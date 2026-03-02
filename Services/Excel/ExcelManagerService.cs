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

                // 3. Sauvegarder directement (Interop suffit généralement pour le besoin utilisateur)
                if (File.Exists(destinationPath))
                    File.Delete(destinationPath);

                // SaveCopyAs est idéal car il ne change pas le fichier ouvert dans SAP
                sapWorkbook.SaveCopyAs(destinationPath);

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

            // Si toujours rien, on crée une nouvelle instance (ou on tente GetActiveObject si dispo)
            try
            {
                oExcelApp = (Application)Marshal.GetActiveObject("Excel.Application");
            }
            catch
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
