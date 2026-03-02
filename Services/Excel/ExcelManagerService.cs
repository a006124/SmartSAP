using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Linq;

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
    public class ExcelManagerService
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

        public string SaveSAPExcelWorkbook(string workbookNamePattern, string destinationPath)
        {
            dynamic excelApp = null;
            dynamic sapWorkbook = null;

            try
            {
                // 1. Récupérer l'instance Excel
                excelApp = GetExcelInstance();

                if (excelApp == null)
                    return "✗ Erreur : Aucune instance Excel ouverte par SAP trouvée.";

                // 2. Trouver le classeur SAP spécifique
                sapWorkbook = FindSapWorkbook(excelApp, workbookNamePattern);
                if (sapWorkbook == null)
                    return $"✗ Erreur : Classeur SAP contenant '{workbookNamePattern}' introuvable.";

                // 3. Sauvegarder une copie temporaire
                string tempFilePath = Path.Combine(Path.GetTempPath(), $"SAP_TempWorkbook_{Guid.NewGuid()}.xlsx");

                if (!SaveWorkbookDirectly(sapWorkbook, tempFilePath))
                {
                    return "✗ Erreur : Échec de la sauvegarde temporaire via Excel.";
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
