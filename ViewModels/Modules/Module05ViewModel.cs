using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Linq;

namespace SmartSAP.ViewModels.Modules
{
    public class Module05ViewModel : ModuleDetailViewModelBase
    {
        public Module05ViewModel(MainViewModel mainViewModel, string title) 
            : base(mainViewModel, title)
        {
        }

        protected override void InitializeSteps()
        {
            Steps = new ObservableCollection<WorkflowStep>
            {
                new WorkflowStep { 
                    Title = "[SAP->Excel] 1.1 Saisie des numéros d'équipement à exporter", 
                    Description = "Crée un nouveau fichier Excel à renseigner (numéros d'équipement) à partir d'un modèle.", 
                    Icon = "\xE70F", 
                    ModuleStep = "E1",
                    ActionCommand = GenerateTemplateCommand 
                },
                new WorkflowStep { 
                    Title = "[SAP->Excel] 1.2 Contrôle et export des données", 
                    Description = "Contrôle et exporte les données (Format SAP). ", 
                    Icon = "\xE762", 
                    ModuleStep = "E1bis",
                    ActionCommand = ExportFixedWidthCommand
                },
                new WorkflowStep { 
                    Title = "[SAP->Excel] 1.3 Récupération des données des équipements", 
                    Description = "Contrôle la connexion et exécute la transaction SAP IH08.", 
                    Icon = "\xE768", 
                    ModuleStep = "E1ter",
                    ActionCommand = ExecuteSAPTransactionCommand
                },
                new WorkflowStep { 
                    Title = "[Saisie de toutes les données] 2. Saisie des données des équipements que l'on veut modifier dans SAP", 
                    Description = "Crée un nouveau fichier Excel à renseigner à partir d'un modèle.", 
                    Icon = "\xE70F", 
                    ModuleStep = "E2",
                    ActionCommand = GenerateTemplateCommand 
                },
                new WorkflowStep { 
                    Title = "[Saisie de toutes les données] 2 bis. Contrôle et export des données", 
                    Description = "Contrôle et exporte les données (Format SAP). ", 
                    Icon = "\xE762", 
                    ModuleStep = "E2bis",
                    ActionCommand = ExportFixedWidthCommand
                },
                new WorkflowStep { 
                    Title = "3. Intégration des modifications dans SAP", 
                    Description = "Contrôle la connexion et exécute la transaction SAP ZSMNBAO13.", 
                    Icon = "\xE768", 
                    ModuleStep = "E3",
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
                // - IH08 est utilisée pour l'extraction (Lecture seule) E1ter
                // - ZSMNBAO13 est utilisée pour l'intégration (Écriture/Modification) E3
                string sapTx = step?.ModuleStep == "E1ter" ? "IH08" : "ZSMNBAO13";
                Logs.Add(new LogEntry("INFO", $"Lancement de la transaction {sapTx}..."));
                
                string resultFile = string.Empty;
                
                string result = await Task.Run(() => 
                    sapTx == "IH08" 
                    ? SAPManager.ExecuteIH08(session, LastExportedTextPath, out resultFile)
                    : SAPManager.ExecuteZSMNBAO13(session, LastExportedTextPath, out resultFile));

                // Affichage du résultat brut dans les logs
                Logs.Add(new LogEntry("DEBUG", $"Réponse brute SAP : {result}"));

                var parts = result.Split('|');
                if (parts.Length >= 2 && parts[1] == "OK")
                {
                    Logs.Add(new LogEntry("SUCCESS", $"✓ Transaction terminée avec succès. Lignes lues: {parts[2]}."));
                    
                    if (!string.IsNullOrEmpty(resultFile))
                    {
                        Logs.Add(new LogEntry("SUCCESS", "Fichier Excel créé : ", resultFile));
                        
                        // 3. Traitement du fichier Excel si étape E1ter
                        if (step?.ModuleStep == "E1ter")
                        {
                            Logs.Add(new LogEntry("INFO", "Génération du modèle E2 pour enrichissement..."));
                            
                            // 1. Qu'un fichier modèle type E2 soit créé
                            var e2Step = Steps.FirstOrDefault(s => s.ModuleStep == "E2") ?? new WorkflowStep { ModuleStep = "E2" };
                            GenerateExcelTemplate(e2Step);
                            string templateE2Path = LastGeneratedExcelPath;

                            if (!string.IsNullOrEmpty(templateE2Path) && System.IO.File.Exists(templateE2Path))
                            {
                                // 2. Que la fonction EnrichirFromSAPExcelWorkbook soit exécutée
                                try
                                {
                                    var excelService = new SmartSAP.Services.Excel.ExcelManagerService();
                                    string enrichResult = excelService.EnrichirFromSAPExcelWorkbook(templateE2Path, resultFile);
                                    Logs.Add(new LogEntry("SUCCESS", $"Enrichissement terminé : {enrichResult}"));
                                }
                                catch (System.Exception ex)
                                {
                                    Logs.Add(new LogEntry("ERROR", $"Erreur lors de l'enrichissement : {ex.Message}"));
                                }
                            }
                        }
                    }
                    
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

                // 3. Traitement du fichier Excel si étape E1ter
                if (step?.ModuleStep == "E1ter")
                {
                    var excelResult = await ExcelManager.SaveSAPExcelWorkbook(LastExportedTextPath, "C:\\Users\\charles\\Desktop\\SAP_TempWorkbook.xlsx");
                    Logs.Add(new LogEntry("SUCCESS", excelResult));
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
            ExcelColumns.Clear();

            switch (step?.ModuleStep)
            {
                case "E1":
                    // LISTE DE NUMÉROS D'ÉQUIPEMENTS
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("N° Equ SAP - 18 car", "Numéro équipement SAP", "", 18));
                    break;
                case "E2":
                    // DONNÉES COMPLÈTES DES ÉQUIPEMENTS
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Division - 4 car (*)", "Division SAP", "MC02", 4, true));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Langue - 2 car (*)", "Code langue", "FR", 2, true));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("N° Equ SAP - 18 car", "Numéro équipement SAP", "", 18));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("N° EQU LICENCE - 24 car", "Numéro licence équipement", "", 24));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("(1) Poste technique - 30 car", "Poste technique lié", "", 30));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("(2) Equipement - 18 car", "Equipement lié", "", 18));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("(3) N° LICENCE DU PERE - 24 car", "Licence équipement parent", "", 24));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Statut RFOU - 1 car", "Statut RFOU", "", 1));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Statut REF - 1 car", "Statut REF", "", 1));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Poste - 4 car", "Numéro de poste", "", 4));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Groupe autorisation - 4 car", "Groupe d'autorisation", "", 4));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Catégorie équipement - 1 car (*)", "Catégorie équipement", "", 1, true));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Libellé fonctionnel de l'équip - 40 car", "Désignation fonctionnelle", "", 40));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Numéro de série fabricant - 30 car", "S/N Fabricant", "", 30));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Type équipement - 10 car", "Type d'équipement", "", 10));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("N° inventaire - 25 car", "Numéro d'inventaire", "", 25));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Code ABC - 1 car", "Criticité ABC", "", 1));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Localisation - 10 car", "Localisation technique", "", 10));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Local - 8 car", "Local", "", 8));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Centre de coût - 10 car", "Centre de coût SAP", "", 10));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Immobilisation principale - 12 car", "Immobilisation principale", "", 12));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Immobilisation subsidiaire - 4 car", "Immobilisation subsidiaire", "", 4));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Valeur d'acquisition - 17 car", "Valeur d'acquisition", "", 17));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Devise - 5 car", "Devise", "EUR", 5));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Date d'acquisition - 8 car", "Date d'acquisition (AAAAMMJJ)", "", 8));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Date début garanti - 8 car", "Début de garantie (AAAAMMJJ)", "", 8));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Date fin garanti - 8 car", "Fin de garantie (AAAAMMJJ)", "", 8));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Repère - 30 car", "Repère équipement", "", 30));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("N° LICENCE - 24 car", "Numéro de licence", "", 24));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Code MABEC - 18 car", "Code MABEC", "", 18));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Libellé matériel de l'équipement - 30 car (*)", "Libellé matériel", "", 30, true));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Niveau équipement - 3 car (*)", "Niveau de l'équipement", "", 3, true));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Référence fournisseur - 25 car (*)", "Réf fournisseur", "", 25, true));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Nom fournisseur - 30 car (*)", "Nom fournisseur", "", 30, true));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Référence intégrateur - 30 car", "Réf intégrateur", "", 30));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Nom intégrateur - 30 car", "Nom intégrateur", "", 30));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Quantité équipement - 17 car", "Quantité", "1", 17));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Mnémonique - 10 car", "Mnémonique", "", 10));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Catégorie d'équipement - 1 car (*)", "Catégorie équipement (Bis)", "", 1, true));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Code Projet - 30 car", "Référence projet", "", 30));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Modèle - 25 car", "Modèle fabricant", "", 25));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Famille - 6 car (*)", "Famille équipement", "", 6, true));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Capacité - 25 car", "Capacité", "", 25));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Alimentation - 25 car", "Alimentation", "", 25));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("A maintenir - 1 car", "Indicateur maintenance", "1", 1));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Uet de Fabrication - 30 car", "UET de fabrication", "", 30));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Dessiné par - 30 car", "Dessiné par", "", 30));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Indice Inventaire - 30 car", "Indice inventaire", "", 30));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Date de l'indice - 8 car", "Date de l'indice (AAAAMMJJ)", "", 8));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Responsable de l'indice - 30 car", "Responsable indice", "", 30));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("N° pièce produit (1) - 30 car", "Pièce produit 1", "", 30));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Indice pièce produit (1) - 30 car", "Indice produit 1", "", 30));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("N° pièce produit (2) - 30 car", "Pièce produit 2", "", 30));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Indice pièce produit (2) - 30 car", "Indice produit 2", "", 30));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("N° pièce produit (3) - 30 car", "Pièce produit 3", "", 30));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Indice pièce produit (3) - 30 car", "Indice produit 3", "", 30));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("N° pièce produit (4) - 30 car", "Pièce produit 4", "", 30));
                    ExcelColumns.Add(new Models.ExcelColumnDefinition("Indice pièce produit (4) - 30 car", "Indice produit 4", "", 30));
                    break;
            }
        }
    }
}
