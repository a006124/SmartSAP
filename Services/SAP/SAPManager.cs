using System;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace SmartSAP.Services.SAP
{
    public class SAPManager
    {
        private const string EcranDemarrageSAP = "SAP Easy Access";

        [DllImport("oleaut32.dll", PreserveSig = false)]
        private static extern void GetActiveObject([In] ref Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        [DllImport("ole32.dll")]
        private static extern int CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid pclsid);

        public static object GetActiveObject(string progId)
        {
            Guid clsid;
            int hr = CLSIDFromProgID(progId, out clsid);
            if (hr < 0) Marshal.ThrowExceptionForHR(hr);

            object obj;
            GetActiveObject(ref clsid, IntPtr.Zero, out obj);
            return obj;
        }

        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable prot);

        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(int reserved, out IBindCtx ppbc);

        [ComImport, Guid("00000010-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IRunningObjectTable
        {
            int Register(int grfFlags, [MarshalAs(UnmanagedType.IUnknown)] object punkObject, IMoniker pmkObjectName);
            int Revoke(int dwRegister);
            int IsRunning(IMoniker pmkObjectName);
            int GetObject(IMoniker pmkObjectName, [MarshalAs(UnmanagedType.IUnknown)] out object ppunkObject);
            int NoteChangeTime(int dwRegister, ref System.Runtime.InteropServices.ComTypes.FILETIME pft);
            int GetTimeOfLastChange(IMoniker pmkObjectName, out System.Runtime.InteropServices.ComTypes.FILETIME pft);
            int EnumRunning(out IEnumMoniker ppenumMoniker);
        }

        [ComImport, Guid("00000102-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IEnumMoniker
        {
            [PreserveSig]
            int Next(int celt, [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] IMoniker[] rgelt, out int pceltFetched);
            [PreserveSig]
            int Skip(int celt);
            [PreserveSig]
            int Reset();
            [PreserveSig]
            int Clone(out IEnumMoniker ppenum);
        }

        [ComImport, Guid("0000000f-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IMoniker
        {
            // Dummy methods to keep the correct vtable order
            void GetClassID(out Guid pClassID);
            [PreserveSig] int IsDirty();
            void Load(object pStm);
            void Save(object pStm, bool fClearDirty);
            void GetSizeMax(out long pcbSize);
            void BindToObject(IBindCtx pbc, IMoniker pmkToLeft, [In] ref Guid riidResult, [MarshalAs(UnmanagedType.IUnknown)] out object ppvResult);
            void BindToStorage(IBindCtx pbc, IMoniker pmkToLeft, [In] ref Guid riidResult, [MarshalAs(UnmanagedType.IUnknown)] out object ppvObj);
            void Reduce(IBindCtx pbc, int dwReduceHowFar, ref IMoniker ppmkToLeft, out IMoniker ppmkReduced);
            void ComposeWith(IMoniker pmkRight, bool fOnlyIfNotGeneric, out IMoniker ppmkComposite);
            void Enum(bool fForward, out IEnumMoniker ppenumMoniker);
            [PreserveSig] int IsEqual(IMoniker pmkOtherMoniker);
            void Hash(out int pdwHash);
            [PreserveSig] int IsRunning(IBindCtx pbc, IMoniker pmkToLeft, IMoniker pmkNewlyRunning);
            void GetTimeOfLastChange(IBindCtx pbc, IMoniker pmkToLeft, out System.Runtime.InteropServices.ComTypes.FILETIME pFileTime);
            void Inverse(out IMoniker ppmk);
            void CommonPrefixWith(IMoniker pmkOther, out IMoniker ppmkPrefix);
            void RelativePathTo(IMoniker pmkOther, out IMoniker ppmkRelPath);
            void GetDisplayName(IBindCtx pbc, IMoniker pmkToLeft, [MarshalAs(UnmanagedType.LPWStr)] out string ppszDisplayName);
            void ParseDisplayName(IBindCtx pbc, IMoniker pmkToLeft, [MarshalAs(UnmanagedType.LPWStr)] string lpszDisplayName, out int pchEaten, out IMoniker ppmkOut);
            void IsSystemMoniker(out int pdwMksys);
        }

        [ComImport, Guid("0000000e-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IBindCtx
        {
            void RegisterObjectBound([MarshalAs(UnmanagedType.IUnknown)] object punk);
            void RevokeObjectBound([MarshalAs(UnmanagedType.IUnknown)] object punk);
            void ReleaseBoundObjects();
            void SetBindOptions(ref object pbindopts);
            void GetBindOptions(ref object pbindopts);
            void GetRunningObjectTable(out IRunningObjectTable pprot);
            void RegisterObjectParam([MarshalAs(UnmanagedType.LPWStr)] string pszKey, [MarshalAs(UnmanagedType.IUnknown)] object punk);
            void GetObjectParam([MarshalAs(UnmanagedType.LPWStr)] string pszKey, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);
            void EnumObjectParam(out object ppenum);
            void RevokeObjectParam([MarshalAs(UnmanagedType.LPWStr)] string pszKey);
        }

        public static object GetObject(string name)
        {
            IRunningObjectTable rot;
            GetRunningObjectTable(0, out rot);
            IEnumMoniker enumMoniker;
            rot.EnumRunning(out enumMoniker);
            enumMoniker.Reset();
            IMoniker[] monikers = new IMoniker[1];
            int fetched;
            IBindCtx bindCtx;
            CreateBindCtx(0, out bindCtx);

            while (enumMoniker.Next(1, monikers, out fetched) == 0)
            {
                string displayName;
                monikers[0].GetDisplayName(bindCtx, null, out displayName);
                if (displayName.Contains(name, StringComparison.OrdinalIgnoreCase))
                {
                    object obj;
                    rot.GetObject(monikers[0], out obj);
                    return obj;
                }
            }
            return null;
        }

        /// <summary>
        /// Vérifie la connexion à SAP et retourne un message d'erreur si la connexion échoue.
        /// Un message vide signifie que la connexion est OK.
        /// </summary>
        public string IsConnectedToSAP()
        {
            try
            {
                // 1. Vérifier si SAP est lancé
                object sapGuiAuto = GetObject("SAPGUI");
                if (sapGuiAuto == null)
                    return "✗ Application SAP non exécutée (Non trouvé dans ROT)";

                // 2. Accéder au moteur de scriptage
                dynamic guiApp = sapGuiAuto.GetType().InvokeMember("GetScriptingEngine", 
                    System.Reflection.BindingFlags.InvokeMethod, null, sapGuiAuto, null);
                
                if (guiApp == null)
                    return "✗ Scriptage SAP non disponible";

                // 3. Vérifier les connexions
                if (guiApp.Children.Count == 0)
                    return "✗ Aucune connexion SAP ouverte";

                dynamic connection = guiApp.Children(0);
                if (connection.Children.Count == 0)
                    return "✗ Aucune session SAP ouverte";

                dynamic session = connection.Children(0);

                // 4. Tenter d'aller au menu principal (Équivalent de /n)
                bool movedToMain = GoMainMenu(session);
                if (!movedToMain)
                    return "✗ Impossible de rejoindre l'écran principal";

                return string.Empty; // Succès
            }
            catch (Exception ex)
            {
                return $"✗ Erreur SAP : {ex.Message}";
            }
        }

        private bool GoMainMenu(dynamic session)
        {
            try
            {
                if (session == null) return false;

                // On envoie /n pour revenir à la racine si on n'est pas déjà sur l'écran Easy Access
                session.findById("wnd[0]/tbar[0]/okcd").Text = "/n";
                session.findById("wnd[0]").sendVKey(0);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public string GetStatus(dynamic session)
        {
            try
            {
                if (session == null) return "✗ Session non connectée";
                return session.ActiveWindow.FindByName("sbar", "GuiStatusbar").Text;
            }
            catch
            {
                return "NOK";
            }
        }
    }
}
