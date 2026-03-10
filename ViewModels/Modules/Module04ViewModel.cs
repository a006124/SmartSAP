using DocumentFormat.OpenXml.Packaging;
using System.Collections.ObjectModel;
using System.IO;
using System.Text.Json;

namespace SmartSAP.ViewModels.Modules
{
    public class Module04ViewModel : ModuleDetailViewModelBase
    {
        public Module04ViewModel(MainViewModel mainViewModel, string title) 
            : base(mainViewModel, title)
        {
            InitializeSteps();
            CompleteInitialization();
        }

        protected override void InitializeSteps()
        {
            Steps = new ObservableCollection<WorkflowStep>
            {
                new WorkflowStep { 
                    Title = "1. Saisie des données nécessaires à la création des Equipements dans SAP", 
                    Description = "Crée un nouveau fichier Excel modèle.", 
                    Icon = "\xE70F", 
                    ModuleStep = "E1_Saisie",
                    ActionCommand = GenerateTemplateCommand 
                },
                new WorkflowStep { 
                    Title = "2. Contrôle et export des données", 
                    Description = "Contrôle et exporte les données (Format SAP). ", 
                    Icon = "\xE762", 
                    ActionCommand = ExportFixedWidthCommand 
                },
                new WorkflowStep { 
                    Title = "3. Intégration SAP", 
                    Description = "Exécute la transaction SAP 'ZSMNBAO12'.", 
                    Icon = "\xE768", 
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

                Logs.Add(new LogEntry("INFO", "Lancement de la transaction ZSMNBAO12..."));
                
                string resultFile = string.Empty;
                string result = await Task.Run(() => SAPManager.ExecuteZSMNBAO12(session, LastExportedTextPath, out resultFile));

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
        protected override void InitializeExcelColumns(WorkflowStep? step = null)
        {
            // Chargement des données depuis JSON
            string dataPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data");
            // Note: En mode Debug/Développement, le chemin peut varier, on essaie aussi le chemin relatif au projet
            if (!Directory.Exists(dataPath))
                dataPath = Path.Combine(Directory.GetCurrentDirectory(), "Data");

            ExcelColumns.Clear();

            var divisions = LoadJsonValues(Path.Combine(dataPath, "division.json"), "01-Division Localisation");
            var langues = LoadJsonValues(Path.Combine(dataPath, "langue.json"), "Langue préférée (division)");
            var groupe_autorisation = LoadJsonValues(Path.Combine(dataPath, "groupe_autorisation.json"), "groupe_autorisation");
            var categorie_equipement = LoadJsonValues(Path.Combine(dataPath, "categorie_equipement.json"), "categorie_equipement");
            var type_equipement = LoadJsonValues(Path.Combine(dataPath, "type_equipement.json"), "type_equipement");
            var abc = LoadJsonValues(Path.Combine(dataPath, "abc.json"), "abc");
            var niveau_equipement = LoadJsonValues(Path.Combine(dataPath, "niveau_equipement.json"), "niveau_equipement");
            var a_maintenir = LoadJsonValues(Path.Combine(dataPath, "a_maintenir.json"), "a_maintenir");

            ExcelColumns.Add(new Models.ExcelColumnDefinition("Division - 4 car (*)", "Division SAP", "MC02", 4, divisions, true, false, true, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Langue - 2 car (*)", "Code langue", "FR", 2, langues, true, false, true, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("N° Equ SAP - 18 car", "Numéro équipement SAP", "", 18, null , false, true, false,["E01.C"]));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("N° EQU LICENCE - 20 car", "Numéro licence équipement", "", 20, null, false, true, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("(1) Poste technique - 30 car", "Poste technique lié", "", 30, null, true, false, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("(2) Equipement - 18 car", "Equipement lié", "", 18, null, true, false, false, ["E01.F"]));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("(3) N° LICENCE DU PERE - 20 car", "Licence équipement parent", "", 20, null ,true,false, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Statut RFOU - 1 car", "Statut RFOU", "", 1, null, false, true, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Statut REF - 1 car", "Statut REF", "", 1, null, false, true, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Poste - 4 car", "Numéro de poste", "", 4, null, false, false, false, ["E01.J"]));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Groupe autorisation - 4 car", "Groupe d'autorisation", "", 4, groupe_autorisation, true, false, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Catégorie équipement - 1 car (*)", "Catégorie équipement", "", 1, categorie_equipement, true, false, true, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Libellé fonctionnel de l'équip - 40 car", "Désignation fonctionnelle", "", 40, null, true, false, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Numéro de série fabricant - 30 car", "S/N Fabricant", "", 30, null, true, false, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Type équipement - 10 car", "Type d'équipement", "", 10, type_equipement, true, false, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("N° inventaire - 25 car", "Numéro d'inventaire", "", 25, null, true, false, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Code ABC - 1 car", "Criticité ABC", "", 1, abc, true, false, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Localisation - 10 car", "Localisation technique", "", 10, null, true, false, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Local - 8 car", "Local", "", 8, null, true, false, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Centre de coût - 10 car", "Centre de coût SAP", "", 10, null, true, false, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Immobilisation principale - 12 car", "Immobilisation principale", "", 12, null, true, false, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Immobilisation subsidiaire - 4 car", "Immobilisation subsidiaire", "", 4, null, true, false, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Valeur d'acquisition - 17 car", "Valeur d'acquisition", "", 17, null, true, false, false, ["E01.W"]));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Devise - 5 car", "Devise", "EUR", 5, null, true, false, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Date d'acquisition - 8 car", "Date d'acquisition (AAAAMMJJ)", "", 8, null, true, false, false, ["E01.Y"]));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Date début garanti - 8 car", "Début de garantie (AAAAMMJJ)", "", 8, null, true, false, false, ["E01.Z"]));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Date fin garanti - 8 car", "Fin de garantie (AAAAMMJJ)", "", 8, null, true, false, false, ["E01.AA"]));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Repère - 30 car", "Repère équipement", "", 30, null, true, false, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("N° LICENCE - 24 car", "Numéro de licence", "", 24, null, true, false, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Code MABEC - 18 car", "Code MABEC", "", 18, null, true, false, false, ["E01.AD"]));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Libellé matériel de l'équipement - 30 car (*)", "Libellé matériel", "", 30, null, true, false, true, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Niveau équipement - 3 car (*)", "Niveau de l'équipement", "", 3, niveau_equipement, true, false,true,null ));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Référence fournisseur - 25 car (*)", "Réf fournisseur", "", 25, null, true, true, true, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Nom fournisseur - 30 car (*)", "Nom fournisseur", "", 30, null, true, true, true, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Référence intégrateur - 30 car", "Réf intégrateur", "", 30, null, true, false, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Nom intégrateur - 30 car", "Nom intégrateur", "", 30, null, true, false, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Quantité équipement - 17 car", "Quantité", "1", 17, null, false, false, false, ["E01.AK"]));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Mnémonique - 10 car", "Mnémonique", "", 10, null, true, false, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Catégorie d'équipement - 1 car (*)", "Catégorie équipement (Bis)", "", 1, categorie_equipement,true,false,true,null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Code Projet - 30 car", "Référence projet", "", 30, null, true, false, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Modèle - 25 car", "Modèle fabricant", "", 25, null, true, false, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Famille - 6 car (*)", "Famille équipement", "", 6, null, false, false, true, ["E01.AP"]));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Capacité - 25 car", "Capacité", "", 25, null, true, false, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Alimentation - 25 car", "Alimentation", "", 25, null, true, false, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("A maintenir - 1 car", "Indicateur maintenance", "1", 1, a_maintenir, true, false, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Uet de Fabrication - 30 car", "UET de fabrication", "", 30, null, true, false, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Dessiné par - 30 car", "Dessiné par", "", 30, null, true, false, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Indice Inventaire - 30 car", "Indice inventaire", "", 30, null, true, false, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Date de l'indice - 8 car", "Date de l'indice (AAAAMMJJ)", "", 8, null, false, false, false, ["E01.AW"]));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Responsable de l'indice - 30 car", "Responsable indice", "", 30, null, true, false, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("N° pièce produit (1) - 30 car", "Pièce produit 1", "", 30, null, true, false, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Indice pièce produit (1) - 30 car", "Indice produit 1", "", 30, null, true, false, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("N° pièce produit (2) - 30 car", "Pièce produit 2", "", 30, null, true, false, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Indice pièce produit (2) - 30 car", "Indice produit 2", "", 30, null, true, false, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("N° pièce produit (3) - 30 car", "Pièce produit 3", "", 30, null, true, false, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Indice pièce produit (3) - 30 car", "Indice produit 3", "", 30, null, true, false, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("N° pièce produit (4) - 30 car", "Pièce produit 4", "", 30, null, true, false, false, null));
            ExcelColumns.Add(new Models.ExcelColumnDefinition("Indice pièce produit (4) - 30 car", "Indice produit 4", "", 30, null, true, false, false, null));
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

