using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Office2019.Drawing.Animation.Model3D;

namespace SmartSAP.ViewModels.Modules
{
    // Poste Technique : Modification en masse
    public class Module02ViewModel : ModuleDetailViewModelBase
    {
        public Module02ViewModel(MainViewModel mainViewModel, string title)
            : base(mainViewModel, title)
        {
        }

        public record ExcelColumnModel(
            string entete,
            string commentaires,
            string exemple,
            int longueurMaxi,
            IEnumerable<string>? valeursAutorisées,
            bool forcerMajuscule,
            bool forcerVide,
            bool forcerDocumentation,
            string règleDeGestion
        );

        // INITIALISATION DES ÉTAPES DU WORKFLOW
        protected override void InitializeSteps()
        {
            Steps = new ObservableCollection<WorkflowStep>
            {
                new WorkflowStep {
                    Title = "[Option1] 1. Modèle Excel",
                    Description = "Crée un fichier Excel pour saisir le code des postes techniques à exporter.",
                    Icon = "\xE70F",
                    ModuleStep = "M02-E1.1",
                    OpenFile = true,
                    ActionCommand = GenerateTemplateCommand
                },
                new WorkflowStep {
                    Title = "[Option1] 2. Excel->TXT",
                    Description = "Contrôle et exporte les données (Format TXT). ",
                    Icon = "\xE762",
                    ModuleStep = "M02-E1.2",
                    NombreMini = 2,
                    OpenFile = false,
                    ActionCommand = ExportFixedWidthCommand
                },
                new WorkflowStep {
                    Title = "[Option1] 3. SAP->Excel",
                    Description = "Récupère les données des postes techniques via la transaction SAP 'IH06'.",
                    Icon = "\xE768",
                    ModuleStep = "M02-E1.3",
                    OpenFile = true,
                    ActionCommand = ExecuteSAPTransactionCommand
                },
                new WorkflowStep {
                    Title = "[Option2] Modèle vierge",
                    Description = "Crée un fichier Excel modèle.",
                    Icon = "\xE70F",
                    ModuleStep = "M02-E2",
                    ActionCommand = GenerateTemplateCommand
                },
                new WorkflowStep {
                    Title = "3. Contrôle et export des données",
                    Description = "Contrôle et exporte les données (Format SAP). ",
                    Icon = "\xE762",
                    ModuleStep = "M02-E3",
                    ActionCommand = ExportFixedWidthCommand
                },
                new WorkflowStep {
                    Title = "4. Intégration des modifications dans SAP",
                    Description = "Exécute la transaction SAP 'ZSMNBAO16'.",
                    Icon = "\xE768",
                    ModuleStep = "M02-E4",
                    ActionCommand = ExecuteSAPTransactionCommand
                }
            };
        }

        // EXÉCUTION DE LA TRANSACTION  SAP
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

                // Exécution de la transaction appropriée en fonction de l'étape :
                string sapTx = step?.ModuleStep == "M02-E1.3" ? "IH06" : "ZSMNBAO16";
                Logs.Add(new LogEntry("INFO", $"Lancement de la transaction {sapTx}..."));

                string resultFile = string.Empty;

                string result = await Task.Run(() =>
                    sapTx == "IH06"
                    ? SAPManager.ExecuteIH06(session, LastExportedTextPath, out resultFile)
                    : SAPManager.ExecuteZSMNBAO16(session, LastExportedTextPath, out resultFile));

                // Affichage du résultat brut dans les logs
                Logs.Add(new LogEntry("DEBUG", $"Réponse brute SAP : {result}"));

                var parts = result.Split('|');
                if (parts.Length >= 2 && parts[1] == "OK")
                {
                    Logs.Add(new LogEntry("SUCCESS", $"✓ Transaction terminée avec succès. Lignes lues: {parts[2]}."));

                    if (!string.IsNullOrEmpty(resultFile))
                    {
                        Logs.Add(new LogEntry("SUCCESS", "Fichier Excel créé : ", resultFile));
                        LastGeneratedSAPExcelPath = resultFile;

                        // 3. Traitement du fichier Excel 
                        if (step?.ModuleStep == "M02-E1.3")
                        {
                            Logs.Add(new LogEntry("INFO", "Génération du modèle M02-E2 pour enrichissement..."));

                            // 1. Fichier modèle type M02-E2 créé
                            var e2Step = Steps.FirstOrDefault(s => s.ModuleStep == "M02-E2") ?? new WorkflowStep { ModuleStep = "M02-E2" };e2Step.OpenFile = false;
                            GenerateExcelTemplate(e2Step);

                            if (!string.IsNullOrEmpty(LastGeneratedSAPExcelPath) && System.IO.File.Exists(LastGeneratedSAPExcelPath))
                            {
                                // 2. Exécution de la fonction EnrichirFromSAPExcelWorkbookM02_E_1_3
                                try
                                {
                                    var excelService = new SmartSAP.Services.Excel.ExcelManager();
                                    string enrichResult = excelService.EnrichirFromSAPExcelWorkbookM02_E_1_3(LastGeneratedExcelPath, LastGeneratedSAPExcelPath);
                                    Logs.Add(new LogEntry("SUCCESS", $"Enrichissement terminé : {enrichResult}"));
                                }
                                catch (System.Exception ex)
                                {
                                    Logs.Add(new LogEntry("ERROR", $"Erreur lors de l'enrichissement : {ex.Message}"));
                                }
                            }
                        }
                    }

                    if (step != null) 
                    { 
                        step.Status = "Terminé"; 
                        step.ResultState = "Success";
                        if (step.OpenFile)
                        {
                            // Ouverture automatique du fichier
                            Process.Start(new ProcessStartInfo(LastGeneratedExcelPath) { UseShellExecute = true });
                        }
                    }
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
        protected override void InitializeExcelColumns(WorkflowStep? step = null)
        {
            // Chargement des données depuis JSON
            string dataPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data");
            if (!Directory.Exists(dataPath))
                dataPath = Path.Combine(Directory.GetCurrentDirectory(), "Data");
            var divisions = LoadJsonValues(Path.Combine(dataPath, "division.json"), "01-Division Localisation");
            var langues = LoadJsonValues(Path.Combine(dataPath, "langue.json"), "Langue préférée (division)");
            var abc = LoadJsonValues(Path.Combine(dataPath, "abc.json"), "abc");
            var a_maintenir = LoadJsonValues(Path.Combine(dataPath, "a_maintenir.json"), "a_maintenir");
            

            ExcelColumns.Clear();

            switch (step?.ModuleStep)
            {
                case "M02-E1.1":
                case "M02-E1.2":
                    // LISTE DES CODES DES POSTES TECHNIQUES À EXPORTER
                    var ExcelModel = new[]
                    {
                        // Entete - Commentaires - Données d'exemple - Longueur maxi - Valeurs autorisées - Majuscules forcées - Vide forcé - Documentation forcée - Règle de gestion
                        new ExcelColumnModel("Poste technique - 30 car", "Poste Technique SAP", "", 30, null, true, false, true, "")
                    };
                    var columnsToAdd1 = ExcelModel.Select(d =>
                        new Models.ExcelColumnDefinition(
                            entete: d.entete,
                            commentaires: d.commentaires,
                            exemple: d.exemple,
                            longueurMaxi: d.longueurMaxi,
                            valeursAutorisées: d.valeursAutorisées?.ToArray(),
                            forcerMajuscule: d.forcerMajuscule,
                            forcerVide: d.forcerVide,
                            forcerDocumentation: d.forcerDocumentation,
                            règleDeGestion: d.règleDeGestion
                    ));

                    foreach (var col in columnsToAdd1)
                    {
                        ExcelColumns.Add(col);
                    }
                    break;
                case "M02-E2":
                case "M02-E3":
                    // Header - Commentaire - Données d'exemple - Largeur fixe - Majuscules forcées - Valeurs autorisées
                    var ExcelModelFull = new List<ExcelColumnModel>
                    {
                        // Entete - Commentaires - Données d'exemple - Longueur maxi - Valeurs autorisées - Majuscules forcées - Vide forcé - Documentation forcée - Règle de gestion
                        new ("Division - 4 car (*)", "Documenter le code suivant les divisions gérées dans SAP", "MC02", 4, divisions, true, false, true, ""),
                        new ("Langue - 2 car (*)", "Documenter le code correspondant à la langue utilisée dans la Désignation", "FR", 2, langues, true, false, true, ""),
                        new ("Poste technique - 30 car (*)", "La valeur saisie doit respecter le code structure défini dans SAP SIMON. Si le poste technique existe déjà dans la base Simon, la ligne est traitée en erreur dans le compte rendu BAO", "", 30, null, true, false, true, ""),
                        new ("Désignation - 40 car (*)", "La désignation saisie sera associée au code langue documenté", "PRESSE TRANSFERT", 40, null, true, false, true, ""),
                        new ("Localisation - 10 car", "Code de localisation, contrôlé suivant table Localisation SAP SIMON", "150", 10, null, true, false, false, ""),
                        new ("Centre de coût - 10 car", "Code du centre de coût, contrôlé dans table des Centres de coûts SAP SIMON", "AC004510", 10, null, true, false, false, ""),
                        new ("Poste - 4 car", "Numéro de poste : Permet dans SAP SIMON de définir un ordre d’affichage du poste technique. Lorsque la donnée est vide, le poste technique sera affiché en 1er dans Simon", "0010", 4, null, true, false, false, "M01.2.G"),
                        new ("Code ABC - 1 car", "Indicateur de criticité 1, 2 ou 3. Si non documenté, Valeur 3 mise par défaut", "1", 1, abc, true, false, false, ""),
                        new ("Code projet - 30 car", "Référence projet", "", 30, null, true, false, false, ""),
                        new ("Code produit - 30 car", "Référence produit", "", 30, null, true, false, false, ""),
                        new ("A maintenir - 1 car", "Indicateur de maintenance (0=Non, 1=Oui)", "1", 1, a_maintenir, true, false, false, "")
                    };

                    var columnsToAdd2 = ExcelModelFull.Select(d =>
                            new Models.ExcelColumnDefinition(
                                entete: d.entete,
                                commentaires: d.commentaires,
                                exemple: d.exemple,
                                longueurMaxi: d.longueurMaxi,
                                valeursAutorisées: d.valeursAutorisées?.ToArray(),
                                forcerMajuscule: d.forcerMajuscule,
                                forcerVide: d.forcerVide,
                                forcerDocumentation: d.forcerDocumentation,
                                règleDeGestion: d.règleDeGestion
                            )
                        );

                    foreach (var col in columnsToAdd2)
                    {
                        ExcelColumns.Add(col);
                    }

                    break;
                case "E3":
                    // DONNÉES COMPLÈTES DES ÉQUIPEMENTS
                    // Header - Commentaire - Données d'exemple - Largeur fixe - Majuscules forcées - Valeurs autorisées

                    ExcelModelFull = new List<ExcelColumnModel>
                    {
                        // Entete - Commentaires - Données d'exemple - Longueur maxi - Valeurs autorisées - Majuscules forcées - Vide forcé - Documentation forcée - Règle de gestion
                        new ("Division - 4 car (*)", "Documenter le code suivant les divisions gérées dans SAP", "MC02", 4, divisions, true, false, true, ""),
                        new ("Langue - 2 car (*)", "Documenter le code correspondant à la langue utilisée dans la Désignation", "FR", 2, langues, true, false, true, ""),
                        new ("Poste technique - 30 car (*)", "La valeur saisie doit respecter le code structure défini dans SAP SIMON. Si le poste technique existe déjà dans la base Simon, la ligne est traitée en erreur dans le compte rendu BAO", "", 30, null, true, false, true, ""),
                        new ("Désignation - 40 car (*)", "La désignation saisie sera associée au code langue documenté", "PRESSE TRANSFERT", 40, null, true, false, true, ""),
                        new ("Localisation - 10 car", "Code de localisation, contrôlé suivant table Localisation SAP SIMON", "150", 10, null, true, false, false, ""),
                        new ("Centre de coût - 10 car", "Code du centre de coût, contrôlé dans table des Centres de coûts SAP SIMON", "AC004510", 10, null, true, false, false, ""),
                        new ("Poste - 4 car", "Numéro de poste : Permet dans SAP SIMON de définir un ordre d’affichage du poste technique. Lorsque la donnée est vide, le poste technique sera affiché en 1er dans Simon", "0010", 4, null, true, false, false, "M01.2.G"),
                        new ("Code ABC - 1 car", "Indicateur de criticité 1, 2 ou 3. Si non documenté, Valeur 3 mise par défaut", "1", 1, abc, true, false, false, ""),
                        new ("Code projet - 30 car", "Référence projet", "", 30, null, true, false, false, ""),
                        new ("Code produit - 30 car", "Référence produit", "", 30, null, true, false, false, ""),
                        new ("A maintenir - 1 car", "Indicateur de maintenance (0=Non, 1=Oui)", "1", 1, a_maintenir, true, false, false, "")
                    };

                    columnsToAdd2 = ExcelModelFull.Select(d =>
                            new Models.ExcelColumnDefinition(
                                entete: d.entete,
                                commentaires: d.commentaires,
                                exemple: d.exemple,
                                longueurMaxi: d.longueurMaxi,
                                valeursAutorisées: d.valeursAutorisées?.ToArray(),
                                forcerMajuscule: d.forcerMajuscule,
                                forcerVide: d.forcerVide,
                                forcerDocumentation: d.forcerDocumentation,
                                règleDeGestion: d.règleDeGestion
                            )
                        );

                    foreach (var col in columnsToAdd2)
                    {
                        ExcelColumns.Add(col);
                    }

                    break;
            }
        }

        private string[] LoadJsonValues(string filePath, string propertyName)
        {
            try
            {
                if (!File.Exists(filePath)) return Array.Empty<string>();

                string jsonContent = File.ReadAllText(filePath);
                using var doc = JsonDocument.Parse(jsonContent);
                return doc.RootElement.EnumerateArray()
                    .Select(e => e.GetProperty(propertyName).GetString())
                    .Where(s => s != null)
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