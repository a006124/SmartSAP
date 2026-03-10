using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace SmartSAP.ViewModels.Modules
{
    // Poste Technique : Création en masse
    public class Module01ViewModel : ModuleDetailViewModelBase
    {
        public Module01ViewModel(MainViewModel mainViewModel, string title) 
            : base(mainViewModel, title)
        {
        }

        protected override void InitializeSteps()
        {
            Steps = new ObservableCollection<WorkflowStep>
            {
                new WorkflowStep { 
                    Title = "1. Saisie des données nécessaires à la création des Postes Techniques dans SAP", 
                    Description = "Crée un nouveau fichier Excel modèle.", 
                    Icon = "\xE70F", 
                    ModuleStep = "E1",
                    ActionCommand = GenerateTemplateCommand 
                },
                new WorkflowStep { 
                    Title = "2. Contrôle et export des données", 
                    Description = "Contrôle et exporte les données (Format SAP). ", 
                    Icon = "\xE762",
                    ModuleStep = "E2",
                    ActionCommand = ExportFixedWidthCommand 
                },
                new WorkflowStep { 
                    Title = "3. Intégration SAP", 
                    Description = "Exécute la transaction SAP 'ZSMNBAO15'.", 
                    Icon = "\xE768", 
                    ModuleStep = "E3",
                    ActionCommand = ExecuteSAPTransactionCommand 
                }
            };
        }

        // EXÉCUTION DE LA TRANSACTION SAP
        protected override async Task ExecuteSAPTransactionAsync(WorkflowStep? step = null)
        {
            await base.ExecuteSAPTransactionAsync(step); // Vérifie la présence du fichier exporté
            
            if (step == null)
            {
                step = Steps.FirstOrDefault(s => s.ActionCommand == ExecuteSAPTransactionCommand);
            }

            if (step != null && step.ResultState == "Error") return;

            try
            {
                // 1. Contrôle de la connexion SAP (Fusionné ici)
                Logs.Add(new LogEntry("INFO", "Contrôle de la connexion SAP..."));
                var connResult = await Task.Run(() => SAPManager.IsConnectedToSAP());
                
                // Mise à jour de la barre d'état globale
                MainViewModel.IsSAPConnected = connResult.IsSuccess;
                MainViewModel.SAPInstanceInfo = connResult.IsSuccess ? $"Instance : {connResult.InstanceInfo}" : "Non connecté";

                if (!connResult.IsSuccess)
                {
                    Logs.Add(new LogEntry("ERROR", connResult.ErrorMessage));
                    if (step != null) { step.Status = "Erreur Connexion"; step.ResultState = "Error"; }
                    return;
                }

                Logs.Add(new LogEntry("SUCCESS", $"✓ Connexion SAP OK : {connResult.InstanceInfo}"));

                // 2. Récupération de la session
                dynamic session = SAPManager.GetActiveSession();
                if (session == null)
                {
                    Logs.Add(new LogEntry("ERROR", "Impossible de récupérer une session SAP active."));
                    if (step != null) { step.Status = "Erreur session"; step.ResultState = "Error"; }
                    return;
                }

                Logs.Add(new LogEntry("INFO", "Lancement de la transaction ZSMNBAO15..."));
                
                string resultFile = string.Empty;
                string result = await Task.Run(() => SAPManager.ExecuteZSMNBAO15(session, LastExportedTextPath, out resultFile));

                // Affichage du résultat brut dans les logs
                Logs.Add(new LogEntry("DEBUG", $"Réponse brute SAP : {result}"));

                var parts = result.Split('|');
                if (parts.Length >= 2 && parts[1] == "OK")
                {
                    Logs.Add(new LogEntry("SUCCESS", $"✓ Transaction terminée avec succès. Lignes lues: {parts[2]}."));
                    if (step != null) { step.Status = "Terminé"; step.ResultState = "Success"; }
                }
                else if (parts.Length >= 2 && parts[1] == "NOK")
                {
                    Logs.Add(new LogEntry("WARNING", $"⚠ Transaction terminée avec {parts[3]} erreur(s)."));
                    if (step != null) { step.Status = "Succès partiel"; step.ResultState = "Error"; }
                }
                else
                {
                    Logs.Add(new LogEntry("ERROR", $"✗ Erreur lors de l'exécution : {result}"));
                    if (step != null) { step.Status = "Erreur SAP"; step.ResultState = "Error"; }
                }
            }
            catch (System.Exception ex)
            {
                Logs.Add(new LogEntry("ERROR", $"Erreur fatale lors de l'intégration SAP : {ex.Message}"));
                if (step != null) { step.Status = "Crash"; step.ResultState = "Error"; }
            }
        }

        // DÉFINITION DES COLONNES DE L'EXCEL MODELE
        // Header - Commentaire - Données d'exemple - Largeur fixe - Majuscules forcées - Valeurs autorisées
        protected override void InitializeExcelColumns(WorkflowStep? step = null)
        {
            ExcelColumns.Clear();
            
            // Chargement des données depuis JSON
            string dataPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data");
            if (!Directory.Exists(dataPath)) 
                dataPath = Path.Combine(Directory.GetCurrentDirectory(), "Data");

            var divisions = LoadJsonValues(Path.Combine(dataPath, "division.json"), "01-Division Localisation");
            var langues = LoadJsonValues(Path.Combine(dataPath, "langue.json"), "Langue préférée (division)");
            var abc = LoadJsonValues(Path.Combine(dataPath, "abc.json"), "abc");
            var a_maintenir = LoadJsonValues(Path.Combine(dataPath, "a_maintenir.json"), "a_maintenir");

            var ExcelModel =new[]
            {
                new { entete="Division - 4 car (*)", commentaires="Division SAP", exemple="MC02", longueurMaxi=4, valeursAutorisees=divisions, forcerMajuscule=true, forcerVide=false,forcerDocumentation=true,règleDeGestion=null },
                new { entete="Langue - 2 car (*)", commentaires="Code langue", exemple="FR", longueurMaxi=2, valeursAutorisees=langues, forcerMajuscule=true, forcerVide=false, forcerDocumentation=true, règleDeGestion=null },
                new { entete="Poste technique - 30 car (*)", commentaires="Poste technique lié", exemple="", longueurMaxi=30, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=true, règleDeGestion=null },
                new { entete="Désignation - 40 car (*)", commentaires="Désignation de l'équipement", exemple="PRESSE TRANSFERT", longueurMaxi=40, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=true, règleDeGestion=null },
                new { entete="Localisation - 10 car", commentaires="Code de localisation", exemple="150", longueurMaxi=10, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null },
                new { entete="Centre de coût - 10 car", commentaires="Code du centre de coût", exemple="AC004510", longueurMaxi=10, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null },
                new { entete="Poste - 4 car", commentaires="Numéro de poste", exemple="0010", longueurMaxi=4, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion="M01.2.G" },
                new { entete="Code ABC - 1 car", commentaires="Indicateur de criticité ABC", exemple="1", longueurMaxi=1, valeursAutorisees=abc, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null },
                new { entete="Code projet - 30 car", commentaires="Référence projet", exemple="", longueurMaxi=30, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null },
                new { entete="Code produit - 30 car", commentaires="Référence produit", exemple="", longueurMaxi=30, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null },
                new { entete="A maintenir - 1 car", commentaires="Indicateur de maintenance (1=Oui)", exemple="1", longueurMaxi=1, valeursAutorisees=a_maintenir, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null },
            };

                ExcelColumns.AddRange(ExcelModel.Select(d =>
                    new Models.ExcelColumnDefinition(
                        entete: d.entete,
                        commentaires: d.commentaires,
                        exemple: d.exemple,
                        longueurMaxi: d.longueurMaxi,
                        valeursAutorisees: d.valeursAutorisees,
                        forcerMajuscule: d.forcerMajuscule,
                        forcerVide: d.forcerVide,
                        forcerDocumentation: d.forcerDocumentation,
                        regleDeGestion: d.règleDeGestion
                )));

        }

        private string[] LoadJsonValues(string filePath, string propertyName)
        {
            try
            {
                if (!File.Exists(filePath)) return Array.Empty<string>();

                string jsonContent = File.ReadAllText(filePath);
                using var doc = JsonDocument.Parse(jsonContent);
                return doc.RootElement.EnumerateArray()
                    .Select(e => e.GetProperty(propertyName).GetString() ?? "")
                    .Where(s => !string.IsNullOrEmpty(s))
                    .ToArray();
            }
            catch (Exception ex)
            {
                Logs.Add(new LogEntry("ERROR", $"Erreur lors du chargement de {filePath} : {ex.Message}"));
                return Array.Empty<string>();
            }
        }
    }
}
