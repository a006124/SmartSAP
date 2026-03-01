using System;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace SmartSAP.Services.SAP
{
    public class SAPManager
    {
        private const string EcranDemarrageSAP = "SAP Easy Access";

        /// <summary>
        /// Vérifie la connexion à SAP et retourne un message d'erreur si la connexion échoue.
        /// Un message vide signifie que la connexion est OK.
        /// </summary>
        public string IsConnectedToSAP()
        {
            try
            {
                // 1. Vérifier si SAP est lancé
                object sapGuiAuto;
                try
                {
                    sapGuiAuto = Marshal.GetActiveObject("SAPGUI");
                }
                catch
                {
                    return "✗ Application SAP non exécutée";
                }

                if (sapGuiAuto == null)
                    return "✗ Application SAP non exécutée";

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
