using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics;
using SmartSAP.Services.Excel;

namespace SmartSAP.Services.SAP
{
    public class SAPConnectionResult
    {
        public bool IsSuccess => string.IsNullOrEmpty(ErrorMessage);
        public string ErrorMessage { get; set; } = string.Empty;
        public string? SystemID { get; set; }
        public string? Client { get; set; }
        public string InstanceInfo => IsSuccess ? $"{SystemID}[{Client}]" : "Non connecté";
    }

    public class SAPManager
    {
        private const string EcranDemarrageSAP = "SAP Easy Access";

        [DllImport("oleaut32.dll", PreserveSig = false)]
        private static extern void GetActiveObject([In] ref Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        [DllImport("ole32.dll")]
        private static extern int CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid pclsid);

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

        public SAPConnectionResult IsConnectedToSAP()
        {
            var result = new SAPConnectionResult();
            try
            {
                object sapGuiAuto = GetObject("SAPGUI");
                if (sapGuiAuto == null)
                {
                    result.ErrorMessage = "✗ Application SAP non exécutée (Non trouvée dans ROT)";
                    return result;
                }

                dynamic guiApp = sapGuiAuto.GetType().InvokeMember("GetScriptingEngine", 
                    System.Reflection.BindingFlags.InvokeMethod, null, sapGuiAuto, null);
                
                if (guiApp == null)
                {
                    result.ErrorMessage = "✗ Scriptage SAP non disponible";
                    return result;
                }

                if (guiApp.Children.Count == 0)
                {
                    result.ErrorMessage = "✗ Aucune connexion SAP ouverte";
                    return result;
                }

                dynamic connection = guiApp.Children(0);
                if (connection.Children.Count == 0)
                {
                    result.ErrorMessage = "✗ Aucune session SAP ouverte";
                    return result;
                }

                dynamic session = connection.Children(0);

                try
                {
                    dynamic info = session.Info;
                    result.SystemID = info.SystemName;
                    result.Client = info.Client;
                }
                catch { }

                bool movedToMain = GoMainMenu(session);
                if (!movedToMain)
                {
                    result.ErrorMessage = "✗ Impossible de rejoindre l'écran principal";
                    return result;
                }

                return result;
            }
            catch (Exception ex)
            {
                result.ErrorMessage = $"✗ Erreur SAP : {ex.Message}";
                return result;
            }
        }

        public object? GetActiveSession()
        {
            try
            {
                object sapGuiAuto = GetObject("SAPGUI");
                if (sapGuiAuto == null) return null;

                dynamic guiApp = sapGuiAuto.GetType().InvokeMember("GetScriptingEngine", 
                    System.Reflection.BindingFlags.InvokeMethod, null, sapGuiAuto, null);
                
                if (guiApp == null || guiApp.Children.Count == 0) return null;

                dynamic connection = guiApp.Children(0);
                if (connection.Children.Count == 0) return null;

                return connection.Children(0);
            }
            catch { return null; }
        }

        private bool GoMainMenu(dynamic session)
        {
            try
            {
                if (session == null) return false;
                session.findById("wnd[0]/tbar[0]/okcd").Text = "/n";
                session.findById("wnd[0]").sendVKey(0);
                return true;
            }
            catch { return false; }
        }

        /// <summary>
        /// Détecte la valeur d'un objet de manière sécurisée.
        /// </summary>
        public string SafeGetText(dynamic session, string id, string defaultValue = "")
        {
            try
            {
                var element = session.findById(id);
                if (element == null) return defaultValue;

                string text = element.Text?.ToString().Trim();
                return string.IsNullOrWhiteSpace(text) ? defaultValue : text;
            }
            catch
            {
                return defaultValue;
            }
        }

        /// <summary>
        /// Détecte un objet de manière sécurisée.
        /// </summary>
        public dynamic? SafeFindById(dynamic session, string id, bool throwIfNotFound = true)
        {
            try
            {
                var element = session.findById(id);
                if (element == null && throwIfNotFound)
                {
                    throw new Exception($"✗ Élément SAP introuvable: {id}");
                }
                return element;
            }
            catch (Exception ex)
            {
                if (throwIfNotFound)
                {
                    throw new Exception($"✗ Erreur accès élément SAP '{id}': {ex.Message}", ex);
                }
                return null;
            }
        }

        /// <summary>
        /// Récupère de manière sécurisée le titre d'une fenêtre SAP.
        /// </summary>
        public string SafeGetTitle(dynamic session, string windowId, string defaultValue = "")
        {
            try
            {
                var window = session.findById(windowId);
                if (window == null) return defaultValue;

                string title = window.Text?.ToString().Trim();
                return string.IsNullOrWhiteSpace(title) ? defaultValue : title;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SAP] Erreur récupération titre fenêtre '{windowId}' : {ex.Message}");
                return defaultValue;
            }
        }

        /// <summary>
        /// Affiche le type et les dimensions d'un objet SAP.
        /// </summary>
        public void DetecterTypeObjetSAP(dynamic session, string idObjet)
        {
            try
            {
                var sapObject = session.findById(idObjet);
                if (sapObject != null)
                {
                    string typeObjet = sapObject.Type;
                    Console.WriteLine("Le type de l'objet est : " + typeObjet);

                    // Note: Ces méthodes peuvent ne pas exister sur tous les types d'objets, d'où le dynamic
                    try
                    {
                        long rows = sapObject.RowCount() - 1;
                        long cols = sapObject.ColumnCount() - 1;
                        Console.WriteLine($"Dimensions : {rows} lignes, {cols} colonnes");
                    }
                    catch { }
                }
                else
                {
                    Console.WriteLine("✗ L'objet n'a pas été trouvé");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("✗ Une erreur s'est produite : " + ex.Message);
            }
        }

        /// <summary>
        /// Liste les éléments enfants d'un conteneur SAP.
        /// </summary>
        public void ListerElementsConteneur(dynamic session, string idConteneur)
        {
            try
            {
                var conteneur = session.findById(idConteneur);
                if (conteneur != null)
                {
                    foreach (var element in conteneur.Children)
                    {
                        string typeElement = element.Type;
                        string idElement = element.Id;
                        Console.WriteLine($"Type : {typeElement}, ID : {idElement}");
                    }
                }
                else
                {
                    Console.WriteLine("✗ Le conteneur n'a pas été trouvé");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("✗ Une erreur s'est produite : " + ex.Message);
            }
        }


        // EXÉCUTION DE LA TRANSACTION SAP ZSMNBAO12 : CRÉATION EN MASSE DES ÉQUIPEMENTS
        public string ExecuteZSMNBAO12(dynamic session, string filePath, out string resultFilePath)
        {
            const string sSAPTransaction = "ZSMNBAO12";
            resultFilePath = string.Empty;
            int NombreDeLignesLues = 0;
            int NombreDeLignesEnErreur = 0;
            string MessageErreur = Environment.NewLine;

            try
            {
                SafeFindById(session, "wnd[0]").maximize();
                SafeFindById(session, "wnd[0]/tbar[0]/okcd").Text = sSAPTransaction;
                SafeFindById(session, "wnd[0]").sendVKey(0);

                // Écran de sélection
                SafeFindById(session, "wnd[0]/usr/ctxtFIC_FILE").Text = filePath;
                SafeFindById(session, "wnd[0]/tbar[1]/btn[8]").press(); // Exécuter 

                // Résultat de l'exécution
                string MessageLigne1;
                try
                {
                    MessageLigne1 = SafeFindById(session, "wnd[0]/usr/lbl[0,10]").Text;
                    if (MessageLigne1 == "Nombre de lignes lues :")
                    {
                        NombreDeLignesLues = int.Parse(SafeFindById(session, "wnd[0]/usr/lbl[29,10]").Text);
                    }
                    for (int i = 14; ; i += 2)
                    {
                        string elementPath = $"wnd[0]/usr/lbl[0,{i}]";
                        string valeurRecuperee;
                        try
                        {
                            valeurRecuperee = SafeFindById(session, elementPath).Text;
                        }
                        catch (Exception ex)
                        {
                            valeurRecuperee = string.Empty;
                        }
                        if (string.IsNullOrEmpty(valeurRecuperee))
                        {
                            break;
                        }
                        MessageErreur += valeurRecuperee + Environment.NewLine; // Ajoute la valeur et un saut de ligne pour la lisibilité
                        NombreDeLignesEnErreur += 1;
                    }
                }
                catch (Exception ex)
                {
                }

                // Retour et Nettoyage
                SafeFindById(session, "wnd[0]/tbar[0]/btn[3]").press(); // Retour
                SafeFindById(session, "wnd[0]/tbar[0]/btn[3]").press(); // Retour

                // Formatage du résultat compact
                string result = $"{sSAPTransaction}|NOK|{NombreDeLignesLues}|{NombreDeLignesEnErreur}|{MessageErreur}";
                Console.WriteLine($"[SAP] Resultat : {result}");
                return result;

            }
            catch (Exception ex)
            {
                string errorResult = $"{sSAPTransaction}|ERROR|{ex.Message}";
                Console.WriteLine($"[SAP] Erreur : {errorResult}");
                return errorResult;
            }
        }

        // EXÉCUTION DE LA TRANSACTION SAP ZSMNBAO13 : MODIFICATION EN MASSE DES ÉQUIPEMENTS
        public string ExecuteZSMNBAO13(dynamic session, string filePath, out string resultFilePath)
        {
            const string sSAPTransaction = "ZSMNBAO13";
            resultFilePath = string.Empty;
            int NombreDeLignesLues = 0;
            int NombreDeLignesTraitées = 0;
            int NombreDeLignesIgnorées = 0;
            string MessageErreur = Environment.NewLine;

            try
            {
                SafeFindById(session, "wnd[0]").maximize();
                SafeFindById(session, "wnd[0]/tbar[0]/okcd").Text = sSAPTransaction;
                SafeFindById(session, "wnd[0]").sendVKey(0);

                // Écran de sélection
                SafeFindById(session, "wnd[0]/usr/ctxtFIC_FILE").Text = filePath;
                SafeFindById(session, "wnd[0]/tbar[1]/btn[8]").press(); // Exécuter

                // Résultat de l'exécution
                string MessageLigne1;
                string MessageLigne2;
                string MessageLigne3;
                try
                {
                    MessageLigne1 = SafeFindById(session, "wnd[0]/usr/lbl[0,10]").Text;
                    if (MessageLigne1 == "Nombre de lignes lues :")
                    {
                        NombreDeLignesLues = int.Parse(SafeFindById(session, "wnd[0]/usr/lbl[29,10]").Text);
                    }
                    MessageLigne2 = SafeFindById(session, "wnd[0]/usr/lbl[4,12]").Text;
                    if (MessageLigne2 == "Nombre de lignes ignorées (non traitées) :")
                    {
                        NombreDeLignesIgnorées = int.Parse(SafeFindById(session, "wnd[0]/usr/lbl[46,12]").Text);
                        NombreDeLignesTraitées = NombreDeLignesLues - NombreDeLignesIgnorées;
                    }
                    else
                    {
                        NombreDeLignesIgnorées = 0;
                        NombreDeLignesTraitées = int.Parse(SafeFindById(session, "wnd[0]/usr/lbl[46,12]").Text);
                    }

                    if (NombreDeLignesIgnorées != 0)
                    {
                        for (int i = 15; ; i += 2)
                        {
                            string elementPath = $"wnd[0]/usr/lbl[0,{i}]";
                            string valeurRecuperee;
                            try
                            {
                                valeurRecuperee = SafeFindById(session, elementPath).Text;
                            }
                            catch (Exception ex)
                            {
                                valeurRecuperee = string.Empty;
                            }
                            if (string.IsNullOrEmpty(valeurRecuperee))
                            {
                                break;
                            }
                            MessageErreur += valeurRecuperee + Environment.NewLine; // Ajoute la valeur et un saut de ligne pour la lisibilité
                        }
                    }
                }
                catch (Exception ex)
                {
                }

                // Retour et Nettoyage
                SafeFindById(session, "wnd[0]/tbar[0]/btn[3]").press(); // Retour
                SafeFindById(session, "wnd[0]/tbar[0]/btn[3]").press(); // Retour

                // Formatage du résultat compact
                string result;
                if (NombreDeLignesIgnorées == 0)
                {
                    result = $"{sSAPTransaction}|OK|{NombreDeLignesLues}|{NombreDeLignesIgnorées}";
                }
                else
                {
                    result = $"{sSAPTransaction}|NOK|{NombreDeLignesLues}|{NombreDeLignesIgnorées}|{MessageErreur}";
                }
                Console.WriteLine($"[SAP] Resultat : {result}");
                return result;
            }
            catch (Exception ex)
            {
                return $"{sSAPTransaction}|ERROR|{ex.Message}";
            }
        }
        
        // EXÉCUTION DE LA TRANSACTION SAP ZSMNBAO15 : CRÉATION EN MASSE DES POSTES TECHNIQUES
        public string ExecuteZSMNBAO15(dynamic session, string filePath, out string resultFilePath)
        {
            const string sSAPTransaction = "ZSMNBAO15";
            resultFilePath = string.Empty;
            int NombreDeLignesLues=0;
            int NombreDeLignesTraitées=0;
            int NombreDeLignesIgnorées = 0;
            string MessageErreur = Environment.NewLine;

            try
            {
                SafeFindById(session, "wnd[0]").maximize();
                SafeFindById(session, "wnd[0]/tbar[0]/okcd").Text = sSAPTransaction;
                SafeFindById(session, "wnd[0]").sendVKey(0);

                // Écran de sélection
                SafeFindById(session, "wnd[0]/usr/ctxtP_FIC_IN").Text = filePath;
                SafeFindById(session, "wnd[0]/tbar[1]/btn[8]").press(); // Exécuter

                // Résultat de l'exécution
                string MessageLigne1;
                string MessageLigne2;
                string MessageLigne3;
                try
                {
                    MessageLigne1 = SafeFindById(session, "wnd[0]/usr/lbl[0,10]").Text;
                    if (MessageLigne1 == "Nombre de lignes lues :")
                    {
                        NombreDeLignesLues = int.Parse(SafeFindById(session, "wnd[0]/usr/lbl[29,10]").Text);
                    }
                    MessageLigne2 = SafeFindById(session, "wnd[0]/usr/lbl[4,12]").Text;
                    if (MessageLigne2 == "Nombre de lignes ignorées (non traitées) :")
                    {
                        NombreDeLignesIgnorées = int.Parse(SafeFindById(session, "wnd[0]/usr/lbl[46,12]").Text);
                        NombreDeLignesTraitées = int.Parse(SafeFindById(session, "wnd[0]/usr/lbl[46,14]").Text);
                    }
                    else
                    {
                        NombreDeLignesIgnorées = 0;
                        NombreDeLignesTraitées = int.Parse(SafeFindById(session, "wnd[0]/usr/lbl[46,12]").Text);
                    }

                    if (NombreDeLignesIgnorées != 0)
                    {
                        for (int i = 17; ; i += 2) 
                        {
                            string elementPath = $"wnd[0]/usr/lbl[0,{i}]";
                            string valeurRecuperee;
                            try
                            {
                                valeurRecuperee = SafeFindById(session, elementPath).Text;
                            }
                            catch (Exception ex)
                            {
                                valeurRecuperee = string.Empty;
                            }
                            if (string.IsNullOrEmpty(valeurRecuperee))
                            {
                                break;
                            }
                            MessageErreur += valeurRecuperee + Environment.NewLine; // Ajoute la valeur et un saut de ligne pour la lisibilité
                        }
                    }
                }
                catch (Exception ex)
                {
                }



                // Retour et Nettoyage
                SafeFindById(session, "wnd[0]/tbar[0]/btn[3]").press(); // Retour
                SafeFindById(session, "wnd[0]/tbar[0]/btn[3]").press(); // Retour

                // Formatage du résultat compact
                string result;
                if (NombreDeLignesIgnorées == 0)
                {
                    result = $"{sSAPTransaction}|OK|{NombreDeLignesLues}|{NombreDeLignesIgnorées}";
                }
                else
                {
                    result = $"{sSAPTransaction}|NOK|{NombreDeLignesLues}|{NombreDeLignesIgnorées}|{MessageErreur}";
                }
                Console.WriteLine($"[SAP] Resultat : {result}");
                return result;
            }
            catch (Exception ex)
            {
                string errorResult = $"{sSAPTransaction}|ERROR|{ex.Message}";
                Console.WriteLine($"[SAP] Erreur : {errorResult}");
                return errorResult;
            }
        }

        // EXÉCUTION DE LA TRANSACTION SAP ZSMNBAO16 : MODIFICATION EN MASSE DES POSTES TECHNIQUES
        public string ExecuteZSMNBAO16(dynamic session, string filePath, out string resultFilePath)
        {
            const string sSAPTransaction = "ZSMNBAO16";
            resultFilePath = string.Empty;

            try
            {
                SafeFindById(session, "wnd[0]").maximize();
                SafeFindById(session, "wnd[0]/tbar[0]/okcd").Text = sSAPTransaction;
                SafeFindById(session, "wnd[0]").sendVKey(0);

                // Écran de sélection
                SafeFindById(session, "wnd[0]/usr/ctxtP_FIC_IN").Text = filePath;
                SafeFindById(session, "wnd[0]/tbar[1]/btn[8]").press(); // Exécuter 

                // Résultat de l'exécution
                int NombreDeLignesLues = int.Parse(SafeFindById(session, "wnd[0]/usr/lbl[29,10]").Text);
                string LibelléRésultat = SafeFindById(session, "wnd[0]/usr/lbl[4,12]").Text;
                int NombreDeLignesTraitées = int.Parse(SafeFindById(session, "wnd[0]/usr/lbl[46,12]").Text);
                int NombreDeLignesIgnorées = int.Parse(SafeFindById(session, "wnd[0]/usr/lbl[46,12]").Text);
                string MessageErreur = Environment.NewLine;
                if (LibelléRésultat == "Nombre de lignes traitées :")
                {
                    NombreDeLignesIgnorées = 0;
                }
                else
                {
                    for (int i = 17; ; i += 2) // La condition de fin est gérée à l'intérieur de la boucle
                    {
                        string elementPath = $"wnd[0]/usr/lbl[0,{i}]";
                        string valeurRecuperee;
                        try
                        {
                            valeurRecuperee = SafeFindById(session, elementPath).Text;
                        }
                        catch (Exception ex)
                        {
                            valeurRecuperee = string.Empty;
                        }
                        if (string.IsNullOrEmpty(valeurRecuperee))
                        {
                            break;
                        }
                        MessageErreur += valeurRecuperee + Environment.NewLine; // Ajoute la valeur et un saut de ligne pour la lisibilité
                    }
                }

                // Retour et Nettoyage
                SafeFindById(session, "wnd[0]/tbar[0]/btn[3]").press(); // Retour
                SafeFindById(session, "wnd[0]/tbar[0]/btn[3]").press(); // Retour

                // Formatage du résultat compact
                string result;
                if (NombreDeLignesIgnorées==0)
                {
                    result = $"{sSAPTransaction}|OK|{NombreDeLignesLues}|{NombreDeLignesIgnorées}";
                }
                else
                {
                    result = $"{sSAPTransaction}|NOK|{NombreDeLignesLues}|{NombreDeLignesIgnorées}|{MessageErreur}";
                }
                Console.WriteLine($"[SAP] Resultat : {result}");
                return result;
            }
            catch (Exception ex)
            {
                string errorResult = $"{sSAPTransaction}|ERROR|{ex.Message}";
                Console.WriteLine($"[SAP] Erreur : {errorResult}");
                return errorResult;
            }
        }


        // EXÉCUTION DE LA TRANSACTION SAP IH06 : Poste Technique / Afficher / Afficher poste technique
        public string ExecuteIH06(dynamic session, string filePath, out string resultFilePath)
        {
            const string sSAPTransaction = "IH06";
            resultFilePath = string.Empty;

            try
            {
                SafeFindById(session, "wnd[0]").maximize();
                SafeFindById(session, "wnd[0]/tbar[0]/okcd").Text = sSAPTransaction;
                SafeFindById(session, "wnd[0]").sendVKey(0);

                // Sélection multiple postes techniques via import de fichier texte
                SafeFindById(session, "wnd[0]/usr/btn%_STRNO_%_APP_%-VALU_PUSH").press(); // Sélection multiple équipements
                SafeFindById(session, "wnd[1]/tbar[0]/btn[23]").press(); // Import de fichier texte

                // On sépare le chemin du nom de fichier car c'est requis par SAP
                string directory = Path.GetDirectoryName(filePath) ?? "";
                string filename = Path.GetFileName(filePath);

                SafeFindById(session, "wnd[2]/usr/ctxtDY_PATH").Text = directory; // Répertoire
                SafeFindById(session, "wnd[2]/usr/ctxtDY_FILENAME").Text = filename; // Nom du fichier
                SafeFindById(session, "wnd[2]/tbar[0]/btn[0]").press(); // Suite
                SafeFindById(session, "wnd[1]/tbar[0]/btn[8]").press(); // Reprendre (F8)

                // Exécuter (F8)
                SafeFindById(session, "wnd[0]/tbar[1]/btn[8]").press();

                // Modifier la mise en forme
                SafeFindById(session, "wnd[0]/mbar/menu[5]/menu[2]/menu[0]").select(); // Option / Mise en forme / Actuelle
                SafeFindById(session, "wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectAll(); // Sélectionner tout
                SafeFindById(session, "wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press(); // Flèche gauche
                SafeFindById(session, "wnd[1]/tbar[0]/btn[0]").press(); // Validation

                // Sauvegarde au format EXCEL
                SafeFindById(session, "wnd[0]/tbar[1]/btn[16]").press(); // Tableur
                SafeFindById(session, "wnd[1]/tbar[0]/btn[0]").press(); // Suite Nombre de colonnes clés

                // Table
                var tableOption = SafeFindById(session, "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]");
                tableOption.Select();
                tableOption.SetFocus();
                SafeFindById(session, "wnd[1]/tbar[0]/btn[0]").press(); // Suite
                SafeFindById(session, "wnd[1]/tbar[0]/btn[0]").press(); // Suite Export d'un objet

                // Attendre la fenêtre "Information"
                string windowTitle = SafeGetTitle(session, "wnd[1]", "");
                while (windowTitle != "Information")
                {
                    windowTitle = SafeGetTitle(session, "wnd[1]", "");
                }

                // Sauvegarde du classeur Excel via le nouveau service (le service inclut maintenant une attente dynamique de 30s max)
                var excelService = new ExcelManager();
                string tempExcelPath = Path.Combine(directory, $"IH06_{Guid.NewGuid()}.xlsx");

                // On cherche le classeur "Feuille de calcul dans Basis" (nom standard SAP)
                string saveResult = excelService.SaveSAPExcelWorkbook("Feuille de calcul dans Basis", tempExcelPath);

                if (saveResult.StartsWith("✅"))
                {
                    resultFilePath = tempExcelPath;
                }
                else
                {
                    Debug.WriteLine($"[Excel] {saveResult}");
                }

                // Fermer la fenêtre d'information
                SafeFindById(session, "wnd[1]/tbar[0]/btn[0]").press(); // Suite

                // Retour au menu principal
                SafeFindById(session, "wnd[0]/tbar[0]/btn[3]").press(); // Retour écran IH06
                SafeFindById(session, "wnd[0]/tbar[0]/btn[3]").press(); // Retour menu principal

                string result = $"{sSAPTransaction}|OK|Exporté|0";
                return result;
            }
            catch (Exception ex)
            {
                return $"{sSAPTransaction}|ERROR|{ex.Message}";
            }
        }

        // EXÉCUTION DE LA TRANSACTION SAP IH08 : Equipement / Afficher / Liste
        public string ExecuteIH08(dynamic session, string filePath, out string resultFilePath)
        {
            const string sSAPTransaction = "IH08";
            resultFilePath = string.Empty;
            
            try
            {
                SafeFindById(session, "wnd[0]").maximize();
                SafeFindById(session, "wnd[0]/tbar[0]/okcd").Text = sSAPTransaction;
                SafeFindById(session, "wnd[0]").sendVKey(0);

                // Sélection multiple équipements via import de fichier texte
                SafeFindById(session, "wnd[0]/usr/btn%_EQUNR_%_APP_%-VALU_PUSH").press(); // Sélection multiple équipements
                SafeFindById(session, "wnd[1]/tbar[0]/btn[23]").press(); // Import de fichier texte
                
                // On sépare le chemin du nom de fichier car c'est requis par SAP
                string directory = Path.GetDirectoryName(filePath) ?? "";
                string filename = Path.GetFileName(filePath);

                SafeFindById(session, "wnd[2]/usr/ctxtDY_PATH").Text = directory; // Répertoire
                SafeFindById(session, "wnd[2]/usr/ctxtDY_FILENAME").Text = filename; // Nom du fichier
                SafeFindById(session, "wnd[2]/tbar[0]/btn[0]").press(); // Suite
                SafeFindById(session, "wnd[1]/tbar[0]/btn[8]").press(); // Reprendre (F8)
                
                // Exécuter (F8)
                SafeFindById(session, "wnd[0]/tbar[1]/btn[8]").press();

                // Modifier la mise en forme
                SafeFindById(session, "wnd[0]/mbar/menu[5]/menu[2]/menu[0]").select(); // Option / Mise en forme / Actuelle
                SafeFindById(session, "wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").selectAll(); // Sélectionner tout
                SafeFindById(session, "wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press(); // Flèche gauche
                SafeFindById(session, "wnd[1]/tbar[0]/btn[0]").press(); // Validation

                // Afficher les classifications
                SafeFindById(session, "wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellColumn = ""; // Sélection de la 1ère ligne
                SafeFindById(session, "wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"; // Sélection de la 1ère ligne
                SafeFindById(session, "wnd[0]/mbar/menu[5]/menu[13]").select()  ; // Option / Affich/OccultClass.
                SafeFindById(session, "wnd[1]/usr/chk[1,3]").selected = true; // Option / Affich/OccultClass.
                SafeFindById(session, "wnd[1]/tbar[0]/btn[0]").press(); // Validation

                // Sauvegarde au format EXCEL
                SafeFindById(session, "wnd[0]/tbar[1]/btn[16]").press(); // Tableur
                SafeFindById(session, "wnd[1]/tbar[0]/btn[0]").press(); // Suite Nombre de colonnes clés
                                
                // Table
                var tableOption = SafeFindById(session, "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]");
                tableOption.Select();
                tableOption.SetFocus();
                SafeFindById(session, "wnd[1]/tbar[0]/btn[0]").press(); // Suite
                SafeFindById(session, "wnd[1]/tbar[0]/btn[0]").press(); // Suite Export d'un objet

                // Attendre la fenêtre "Information"
                string windowTitle = SafeGetTitle(session, "wnd[1]", "");
                while (windowTitle != "Information")
                {
                    windowTitle = SafeGetTitle(session, "wnd[1]", "");
                }

                // Sauvegarde du classeur Excel via le nouveau service (le service inclut maintenant une attente dynamique de 30s max)
                var excelService = new ExcelManager();
                string tempExcelPath = Path.Combine(directory, $"IH08_{Guid.NewGuid()}.xlsx");
                
                // On cherche le classeur "Feuille de calcul dans Basis" (nom standard SAP)
                string saveResult = excelService.SaveSAPExcelWorkbook("Feuille de calcul dans Basis", tempExcelPath);
                
                if (saveResult.StartsWith("✅"))
                {
                    resultFilePath = tempExcelPath;
                }
                else
                {
                    Debug.WriteLine($"[Excel] {saveResult}");
                }

                // Fermer la fenêtre d'information
                SafeFindById(session, "wnd[1]/tbar[0]/btn[0]").press(); // Suite
                    
                // Retour au menu principal
                SafeFindById(session, "wnd[0]/tbar[0]/btn[3]").press(); // Retour écran IH08
                SafeFindById(session, "wnd[0]/tbar[0]/btn[3]").press(); // Retour menu principal

                string result = $"{sSAPTransaction}|OK|Exporté|0";
                return result;
            }
            catch (Exception ex)
            {
                return $"{sSAPTransaction}|ERROR|{ex.Message}";
            }
        }




    }
}
