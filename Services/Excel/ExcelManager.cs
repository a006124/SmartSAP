using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SmartSAP.Services.Excel
{
    // ---------------------------------------------------------
    // 1. Classe Helper pour contourner Marshal.GetActiveObject
    // ---------------------------------------------------------
    public static class Win32ComHelper
    {
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string lclassName, string windowTitle);

        [DllImport("oleacc.dll")]
        private static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint dwId, ref Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object ppvObject);

        public static dynamic GetExcelFromProcess(Process proc)
        {
            try
            {
                IntPtr hwnd = proc.MainWindowHandle;
                if (hwnd == IntPtr.Zero) return null;

                // 1. On cherche la sous-fenêtre "XLDESK"
                IntPtr hwndDesk = FindWindowEx(hwnd, IntPtr.Zero, "XLDESK", null);
                if (hwndDesk == IntPtr.Zero) return null;

                // 2. On cherche la sous-fenêtre "EXCEL7" (C'est la grille du tableur)
                IntPtr hwndExcel7 = FindWindowEx(hwndDesk, IntPtr.Zero, "EXCEL7", null);
                if (hwndExcel7 == IntPtr.Zero) return null;

                // 3. On extrait l'objet COM depuis cette fenêtre
                Guid IID_IDispatch = new Guid("{00020400-0000-0000-C000-000000000046}");
                object ptr;
                int hr = AccessibleObjectFromWindow(hwndExcel7, 0xFFFFFFF0, ref IID_IDispatch, out ptr); // OBJID_NATIVEOM = -16

                if (hr >= 0 && ptr != null)
                {
                    // L'objet retourné est une "Window" (dynamic), on remonte à l'Application
                    dynamic win = ptr;
                    return win.Application;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("✗ Erreur méthode AccessibleObject : " + ex.Message);
            }
            return null;
        }
    }

    // ---------------------------------------------------------
    // 2. Service de gestion Excel (SAP -> Local)
    // ---------------------------------------------------------
    public class ExcelManager
    {
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("ole32.dll")]
        private static extern int GetActiveObject(ref Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        [DllImport("ole32.dll")]
        private static extern int CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid lpclsid);

        private object GetActiveObjectWrapper(string progId)
        {
            Guid clsid;
            int hr = CLSIDFromProgID(progId, out clsid);
            if (hr < 0) return null;

            object obj;
            hr = GetActiveObject(ref clsid, IntPtr.Zero, out obj);
            if (hr < 0) return null;

            return obj;
        }


        public string SaveSAPExcelWorkbook(string workbookNamePattern, string destinationPath, int timeoutSeconds = 30)
        {
            dynamic excelApp = null;
            dynamic sapWorkbook = null;

            try
            {
                // 1. Récupérer l'instance Excel (Polling pour attendre que SAP la crée/l'utilise)
                DateTime startTime = DateTime.Now;
                while ((DateTime.Now - startTime).TotalSeconds < timeoutSeconds)
                {
                    excelApp = GetExcelInstance();
                    if (excelApp != null)
                    {
                        // 2. Trouver le classeur SAP spécifique (Polling)
                        sapWorkbook = FindSapWorkbook(excelApp, workbookNamePattern);
                        if (sapWorkbook != null)
                            break; // Le document a été trouvé, on sort de la boucle d'attente
                    }
                    
                    System.Threading.Thread.Sleep(1000); // Attendre 1 seconde avant de réessayer
                }

                if (excelApp == null)
                    return $"✗ Erreur : Aucune instance Excel ouverte par SAP trouvée après {timeoutSeconds}s.";

                if (sapWorkbook == null)
                    return $"✗ Erreur : Classeur SAP contenant '{workbookNamePattern}' introuvable après {timeoutSeconds}s d'attente.";

                // 3. Sauvegarder une copie temporaire (avec polling pour attendre la fin de l'écriture par SAP)
                string tempFilePath = Path.Combine(Path.GetTempPath(), $"SAP_TempWorkbook_{Guid.NewGuid()}.xlsx");

                bool saveSuccess = false;
                DateTime saveStartTime = DateTime.Now;
                
                while ((DateTime.Now - saveStartTime).TotalSeconds < timeoutSeconds)
                {
                    try
                    {
                        // On vérifie que le moteur Excel n'est plus occupé par l'export de SAP
                        if (excelApp.Ready)
                        {
                            if (SaveWorkbookDirectly(sapWorkbook, tempFilePath))
                            {
                                saveSuccess = true;
                                break;
                            }
                        }
                    }
                    catch (Exception)
                    {
                        // Une exception COM (ex: RPC_E_CALL_REJECTED) indique qu'Excel est toujours occupé
                    }
                    
                    System.Threading.Thread.Sleep(1000); // Attendre 1 seconde avant de réessayer
                }

                if (!saveSuccess)
                {
                    return "✗ Erreur : Échec de la sauvegarde temporaire via Excel après attente.";
                }

                // 4. Traiter avec NPOI
                bool success = ProcessWorkbookWithNPOI(tempFilePath, destinationPath);

                // 5. Nettoyer le fichier temporaire
                try
                {
                    if (File.Exists(tempFilePath)) File.Delete(tempFilePath);
                }
                catch (Exception exIO)
                {
                    Debug.WriteLine("Warning: Impossible de supprimer le fichier temporaire : " + exIO.Message);
                }

                if (!success)
                {
                    return "✗ Erreur : Échec du traitement/sauvegarde finale avec NPOI.";
                }

                return $"✅ Succès : Classeur SAP sauvegardé sous '{destinationPath}'.";
            }
            catch (Exception ex)
            {
                return $"✗ Erreur inattendue : {ex.Message}";
            }
            finally
            {
                // Libération des objets COM
                if (sapWorkbook != null) Marshal.ReleaseComObject(sapWorkbook);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            }
        }

        // Module02 
        public string EnrichirFromSAPExcelWorkbookM02_E_1_3(string templatePath, string sourceDataPath)
        {
            try
            {
                using (var wExcelToUpdate = new ClosedXML.Excel.XLWorkbook(templatePath))
                using (var wExcelFrom = new ClosedXML.Excel.XLWorkbook(sourceDataPath))
                {
                    var sheetToUpdate = wExcelToUpdate.Worksheet(1); // Le template n'a qu'un seul onglet "Data"
                    var sheetFrom = wExcelFrom.Worksheet(1);

                    int lRowToUpdate = 2;

                    foreach (var row in sheetFrom.RowsUsed())
                    {
                        // On ne traite pas la ligne de titre (ligne 1)
                        if (row.RowNumber() == 1) continue;

                        var targetRow = sheetToUpdate.Row(lRowToUpdate);

                        targetRow.Cell(1).Value = row.Cell(6).GetString(); // Division
                        targetRow.Cell(2).Value = row.Cell(91).GetString(); // Langue
                        targetRow.Cell(3).Value = row.Cell(1).GetString(); // Poste technique
                        targetRow.Cell(4).Value = row.Cell(2).GetString(); // Désignation
                        targetRow.Cell(5).Value = row.Cell(9).GetString(); // Localisation
                        targetRow.Cell(6).Value = row.Cell(3).GetString(); // Centre de coûts
                        targetRow.Cell(7).Value = row.Cell(4).GetString(); // Poste
                        targetRow.Cell(8).Value = row.Cell(11).GetString(); // Code ABC
                        targetRow.Cell(9).Value = string.Empty; ; // Code Projet
                        targetRow.Cell(10).Value = string.Empty; ; // Code Produit
                        targetRow.Cell(11).Value = string.Empty; // A maintenir

                        lRowToUpdate++;
                    }

                    wExcelToUpdate.Save();
                } // La sortie du bloc using va garantir un Dispose et la libération des pointeurs de fichiers ouverts.

                // Nettoyage : Suppression du fichier temporaire source
                try
                {
                    if (System.IO.File.Exists(sourceDataPath))
                    {
                        System.IO.File.Delete(sourceDataPath);
                    }
                }
                catch (Exception exIO)
                {
                    Debug.WriteLine($"Warning: Impossible de supprimer le fichier source temporaire '{sourceDataPath}' : {exIO.Message}");
                }

                return $"Modèle E2 généré et enrichi avec les données extraites !";
            }
            catch (Exception ex)
            {
                return $"✗ Erreur lors de l'enrichissement du modèle : {ex.Message}";
            }
        }

        public int GetColumnNumberByHeader(ClosedXML.Excel.IXLWorksheet sheetFrom, string columnHeader)
        {
            // Parcourir les cellules de la première ligne (ligne d'en-tête)
            foreach (var cell in sheetFrom.Row(1).CellsUsed())
            {
                // Comparer le texte de la cellule avec le libellé recherché (ignorer la casse)
                if (cell.GetString().Equals(columnHeader, StringComparison.OrdinalIgnoreCase))
                {
                    return cell.WorksheetColumn().ColumnNumber();
                }
            }
            return -1; // Retourne -1 si le libellé n'est pas trouvé
        }

        public string SetCellValue(ClosedXML.Excel.IXLWorksheet sheetFrom, string columnHeader, int rowNumber)
        {
            int targetColumn = GetColumnNumberByHeader(sheetFrom, columnHeader);            
            if (targetColumn == -1) // Le libellé n'a pas été trouvé
            {
                return string.Empty; 
            }
            else 
            {
                var theRowFrom = sheetFrom.Row(rowNumber);
                return theRowFrom.Cell(targetColumn).GetString();
            }
        }


        public string EnrichirFromSAPExcelWorkbookM05_E_1_3(string templatePath, string sourceDataPath)
        {
            try
            {
                using (var wExcelToUpdate = new ClosedXML.Excel.XLWorkbook(templatePath))
                using (var wExcelFrom = new ClosedXML.Excel.XLWorkbook(sourceDataPath))
                {
                    var sheetToUpdate = wExcelToUpdate.Worksheet(1); // Le template n'a qu'un seul onglet "Data"
                    var sheetFrom = wExcelFrom.Worksheet(1);

                    int lRowToUpdate = 2;

                    foreach (var row in sheetFrom.RowsUsed())
                    {
                        // On ne traite pas la ligne de titre (ligne 1)
                        if (row.RowNumber() == 1) continue;

                        var targetRow = sheetToUpdate.Row(lRowToUpdate);
                        targetRow.Cell(1).Value = SetCellValue(sheetFrom, "Division local.", row.RowNumber()); // Division
                        targetRow.Cell(2).Value = SetCellValue(sheetFrom, "Langue", row.RowNumber()); // Langue
                        targetRow.Cell(3).Value = SetCellValue(sheetFrom, "Equipement", row.RowNumber()); // Numéro Equipement
                        targetRow.Cell(4).Value = ""; // SetCellValue(sheetFrom, "Numéro licence", row.RowNumber()); // License

                        // Règle : si Equipement supérieur est documenté alors Poste Technique = "" sinon Equipement Supérieur = ""
                        string equipSup = row.Cell(GetColumnNumberByHeader(sheetFrom, "Equip.supérieur")).GetString();
                        if (!string.IsNullOrEmpty(equipSup))
                        {
                            targetRow.Cell(5).Value = ""; // On efface le poste technique
                            targetRow.Cell(6).Value = equipSup;
                        }
                        else
                        {
                            targetRow.Cell(5).Value = row.Cell(GetColumnNumberByHeader(sheetFrom, "Poste technique")).GetString(); // On garde le poste technique
                            targetRow.Cell(6).Value = ""; 
                        }
                        targetRow.Cell(7).Value = ""; // License du père
                        targetRow.Cell(8).Value = ""; // RFOU
                        targetRow.Cell(9).Value = ""; // REF
                        targetRow.Cell(10).Value = SetCellValue(sheetFrom, "Poste", row.RowNumber()); // Position
                        targetRow.Cell(11).Value = SetCellValue(sheetFrom, "Groupe autoris.", row.RowNumber()); // Groupe d'autorisation
                        targetRow.Cell(12).Value = SetCellValue(sheetFrom, "Catég.équipemnt", row.RowNumber());  // Catégorie de l'équipement
                        targetRow.Cell(13).Value = SetCellValue(sheetFrom, "Désignation", row.RowNumber());  // Libellé fonctionnel
                        targetRow.Cell(14).Value = SetCellValue(sheetFrom, "N° série fabr.", row.RowNumber()); // N° de série fabricant
                        targetRow.Cell(15).Value = SetCellValue(sheetFrom, "Type d'objet", row.RowNumber()); // Type d'équipement
                        targetRow.Cell(16).Value = SetCellValue(sheetFrom, "N° inventaire", row.RowNumber()); // N° inventaire
                        targetRow.Cell(17).Value = SetCellValue(sheetFrom, "Code ABC", row.RowNumber()); // Code ABC
                        targetRow.Cell(18).Value = SetCellValue(sheetFrom, "Localisation", row.RowNumber()); // Localisation
                        targetRow.Cell(19).Value = SetCellValue(sheetFrom, "Local", row.RowNumber()); // Local
                        targetRow.Cell(20).Value = SetCellValue(sheetFrom, "Centre de coûts", row.RowNumber()); // Centre de coûts
                        targetRow.Cell(21).Value = SetCellValue(sheetFrom, "Immobilisation", row.RowNumber()); // Immobilisation principale
                        targetRow.Cell(22).Value = SetCellValue(sheetFrom, "Nº subsidiaire", row.RowNumber()); // Immobilisation secondaire
                        targetRow.Cell(23).Value = SetCellValue(sheetFrom, "Val.acquisition", row.RowNumber()); // Valeur d'acquisition
                        targetRow.Cell(24).Value = SetCellValue(sheetFrom, "Devise", row.RowNumber()); // Devise
                        targetRow.Cell(25).Value = SetCellValue(sheetFrom, "Date acquis.", row.RowNumber()); // Date d'acquisition
                        targetRow.Cell(26).Value = SetCellValue(sheetFrom, "Début gar.fourn", row.RowNumber()); // Début de garantie
                        targetRow.Cell(27).Value = SetCellValue(sheetFrom, "Fin gar. fourn.", row.RowNumber()); // Fin de garantie
                        targetRow.Cell(28).Value = SetCellValue(sheetFrom, "Zone de tri", row.RowNumber()); // Repère / Zone de tri
                        targetRow.Cell(29).Value = SetCellValue(sheetFrom, "Numéro licence", row.RowNumber()); // N° License
                        targetRow.Cell(30).Value = SetCellValue(sheetFrom, "Article", row.RowNumber()); // Code MABEC / Article
                        targetRow.Cell(31).Value = SetCellValue(sheetFrom, "Désignation", row.RowNumber()); // Libellé matériel

                        // Niveau
                        string niveau = row.Cell(GetColumnNumberByHeader(sheetFrom, "Niveau de l'équipeme")).GetString();
                        switch (niveau)
                        {
                            case "Groupe d'ensemble": targetRow.Cell(32).Value = "GE"; break;
                            case "Ensemble": targetRow.Cell(32).Value = "E"; break;
                            case "Sous Ensemble": targetRow.Cell(32).Value = "S/E"; break;
                            default: targetRow.Cell(32).Value = ""; break;
                        }

                        targetRow.Cell(33).Value = SetCellValue(sheetFrom, "Référence fournisseu", row.RowNumber()); // Référence fournisseur
                        targetRow.Cell(34).Value = SetCellValue(sheetFrom, "Nom du fournisseur *", row.RowNumber()); // Nom Fournisseur
                        targetRow.Cell(35).Value = SetCellValue(sheetFrom, "Référence intégrateu", row.RowNumber()); // Référence intégrateur
                        targetRow.Cell(36).Value = SetCellValue(sheetFrom, "Nom intégrateur", row.RowNumber()); // Nom intégrateur
                        targetRow.Cell(37).Value = SetCellValue(sheetFrom, "Quantité d'équipemen", row.RowNumber()); // Quantité équipement
                        targetRow.Cell(38).Value = SetCellValue(sheetFrom, "Mnémonique", row.RowNumber()); // Mnémonique
                        targetRow.Cell(39).Value = ""; // Catégorie - Nature de l'équipement
                        targetRow.Cell(40).Value = ""; // Code projet
                        targetRow.Cell(41).Value = ""; // Modèle
                        targetRow.Cell(42).Value = ""; // Famille
                        targetRow.Cell(43).Value = SetCellValue(sheetFrom, "Capacité - GMAO", row.RowNumber()); // Capacité
                        targetRow.Cell(44).Value = ""; // Alimentation

                        // A maintenir
                        string aMaintenir = row.Cell(GetColumnNumberByHeader(sheetFrom, "A maintenir")).GetString();
                        switch (aMaintenir)
                        {
                            case "Avec Maintenance": targetRow.Cell(45).Value = "1"; break;
                            case "Sans Maintenance": targetRow.Cell(45).Value = "0"; break;
                            default: targetRow.Cell(44).Value = ""; break;
                        }

                        targetRow.Cell(46).Value = ""; // UET de fabrication
                        targetRow.Cell(47).Value = ""; // Dessiné par
                        targetRow.Cell(48).Value = ""; // Indice inventaire
                        targetRow.Cell(49).Value = ""; // Date de l'indice
                        targetRow.Cell(50).Value = ""; // Responsable de l'indice
                        targetRow.Cell(51).Value = ""; // N° pièce produit
                        targetRow.Cell(52).Value = ""; // Indice pièce produit
                        targetRow.Cell(53).Value = ""; // N° pièce produit
                        targetRow.Cell(54).Value = ""; // Indice pièce produit
                        targetRow.Cell(55).Value = ""; // N° pièce produit
                        targetRow.Cell(56).Value = ""; // Indice pièce produit
                        targetRow.Cell(57).Value = ""; // N° pièce produit
                        targetRow.Cell(58).Value = ""; // Indice pièce produit

                        lRowToUpdate++;
                    }

                    wExcelToUpdate.Save();
                } // La sortie du bloc using va garantir un Dispose et la libération des pointeurs de fichiers ouverts.
                
                // Nettoyage : Suppression du fichier temporaire source
                try
                {
                    if (System.IO.File.Exists(sourceDataPath))
                    {
                        System.IO.File.Delete(sourceDataPath);
                    }
                }
                catch (Exception exIO)
                {
                    Debug.WriteLine($"Warning: Impossible de supprimer le fichier source temporaire '{sourceDataPath}' : {exIO.Message}");
                }

                return $"Modèle E2 généré et enrichi avec les données extraites !";
            }
            catch (Exception ex)
            {
                return $"✗ Erreur lors de l'enrichissement du modèle : {ex.Message}";
            }
        }

        private bool SaveWorkbookDirectly(dynamic wb, string destinationPath)
        {
            try
            {
                if (File.Exists(destinationPath)) File.Delete(destinationPath);
                wb.SaveCopyAs(destinationPath);
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("✗ Échec sauvegarde directe : " + ex.Message);
                return false;
            }
        }

        private bool ProcessWorkbookWithNPOI(string sourcePath, string destinationPath)
        {
            try
            {
                using (var fs = new FileStream(sourcePath, FileMode.Open, FileAccess.Read))
                {
                    NPOI.SS.UserModel.IWorkbook workbook;

                    if (sourcePath.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                        workbook = new NPOI.HSSF.UserModel.HSSFWorkbook(fs);
                    else
                        workbook = new NPOI.XSSF.UserModel.XSSFWorkbook(fs);

                    using (var outFs = new FileStream(destinationPath, FileMode.Create, FileAccess.Write))
                    {
                        workbook.Write(outFs);
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("✗ Erreur NPOI : " + ex.Message);
                return false;
            }
        }

        private dynamic GetExcelInstance()
        {
            dynamic oExcelApp = null;

            // Tentative via processus avec fenêtre valide
            Process validProc = Process.GetProcessesByName("EXCEL")
                .FirstOrDefault(p => p.MainWindowHandle != IntPtr.Zero);

            if (validProc != null)
            {
                try
                {
                    oExcelApp = Win32ComHelper.GetExcelFromProcess(validProc);
                    if (oExcelApp != null) return oExcelApp;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("✗ Echec méthode Fenêtre : " + ex.Message);
                }
            }

            // Si toujours rien, on tente de récupérer l'instance active via OLE
            try
            {
                oExcelApp = GetActiveObjectWrapper("Excel.Application");
                if (oExcelApp != null) return oExcelApp;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("✗ Echec GetActiveObject : " + ex.Message);
            }

            // En dernier recours, on crée une nouvelle instance
            try
            {
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType != null)
                {
                    oExcelApp = Activator.CreateInstance(excelType);
                    oExcelApp.Visible = true;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("✗ Echec création nouvelle instance : " + ex.Message);
            }

            return oExcelApp;
        }

        private dynamic FindSapWorkbook(dynamic excelApp, string workbookNamePattern)
        {
            try
            {
                foreach (dynamic wb in excelApp.Workbooks)
                {
                    string wbName = wb.Name;
                    if (wbName.IndexOf(workbookNamePattern, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return wb;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("✗ Erreur lors du parcours des classeurs : " + ex.Message);
            }
            return null;
        }
    }
}
