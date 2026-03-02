using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Window = Microsoft.Office.Interop.Excel.Window;

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

        public static Application GetExcelFromProcess(Process proc)
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
                    // L'objet retourné est une "Window", on remonte à l'Application
                    Window win = (Window)ptr;
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

        private const uint WM_SYSCOMMAND = 0x112;
        private const uint SC_RESTORE = 0xF120;

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        public string SaveSAPExcelWorkbook(string workbookNamePattern, string destinationPath)
        {
            Application excelApp = null;
            Workbook sapWorkbook = null;

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

                // 3. Sauvegarder une copie temporaire via Interop
                string tempFilePath = Path.Combine(Path.GetTempPath(), $"SAP_TempWorkbook_{Guid.NewGuid()}.xlsx");

                if (!SaveWorkbookDirectly(sapWorkbook, tempFilePath))
                {
                    return "✗ Erreur : Échec de la sauvegarde temporaire via Excel Interop.";
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
            catch (COMException comEx)
            {
                return $"✗ Erreur COM ({comEx.ErrorCode:X}) : {comEx.Message}";
            }
            catch (Exception ex)
            {
                return $"✗ Erreur inattendue : {ex.Message}";
            }
            finally
            {
                // Libération propre des objets COM
                if (sapWorkbook != null) Marshal.FinalReleaseComObject(sapWorkbook);
                if (excelApp != null) Marshal.FinalReleaseComObject(excelApp);
            }
        }

        private bool SaveWorkbookDirectly(Workbook wb, string destinationPath)
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
                    {
                        workbook = new NPOI.HSSF.UserModel.HSSFWorkbook(fs);
                    }
                    else
                    {
                        workbook = new NPOI.XSSF.UserModel.XSSFWorkbook(fs);
                    }

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

        private Application GetExcelInstance()
        {
            Application oExcelApp = null;

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
                var activeObj = GetActiveObjectWrapper("Excel.Application");
                if (activeObj != null)
                {
                    oExcelApp = (Application)activeObj;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("✗ Echec GetActiveObject : " + ex.Message);
            }

            // En dernier recours, on crée une nouvelle instance
            if (oExcelApp == null)
            {
                oExcelApp = new Application { Visible = true };
            }

            return oExcelApp;
        }

        private Workbook FindSapWorkbook(Application excelApp, string workbookNamePattern)
        {
            try
            {
                foreach (Workbook wb in excelApp.Workbooks)
                {
                    if (wb.Name.IndexOf(workbookNamePattern, StringComparison.OrdinalIgnoreCase) >= 0)
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
