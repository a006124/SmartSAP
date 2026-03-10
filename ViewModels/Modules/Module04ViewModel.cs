using DocumentFormat.OpenXml.Packaging;
using System.Collections.ObjectModel;
using System.IO;
using System.Text.Json;

namespace SmartSAP.ViewModels.Modules
{
    public class Module04ViewModel : ModuleDetailViewModelBase
    {
        // Equipement : Création en masse
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
                    Description = "Exécute la transaction SAP 'ZSMNBAO12'.", 
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
                    var nature_equipement = LoadJsonValues(Path.Combine(dataPath, "nature_equipement.json"), "nature_equipement");
                    var a_maintenir = LoadJsonValues(Path.Combine(dataPath, "a_maintenir.json"), "a_maintenir");

                    var ExcelModel =new[]
                    {
                        new { entete="Division - 4 car (*)", commentaires="Division SAP", exemple="MC02", longueurMaxi=4, valeursAutorisees=divisions, forcerMajuscule=true, forcerVide=false,forcerDocumentation=true,règleDeGestion=null },
                        new { entete="Langue - 2 car (*)", commentaires="Code langue", exemple="FR", longueurMaxi=2, valeursAutorisees=langues, forcerMajuscule=true, forcerVide=false, forcerDocumentation=true, règleDeGestion=null },
                        new { entete="N° Equ SAP - 18 car", commentaires="Numéro équipement SAP", exemple="", longueurMaxi=18, valeursAutorisees=null, forcerMajuscule=true, forcerVide=true, forcerDocumentation=false, règleDeGestion=null },
                        new { entete="N° EQU LICENCE - 20 car", commentaires="Numéro licence équipement", exemple="", longueurMaxi=20, valeursAutorisees=null, forcerMajuscule=true, forcerVide=true, forcerDocumentation=false, règleDeGestion=null },
                        new { entete="(1) Poste technique - 30 car", commentaires="Poste technique lié", exemple="", longueurMaxi=30, valeursAutorisees=null, forcerMajuscule=true, forcerVide=true, forcerDocumentation=false, règleDeGestion=null },
                        new { entete="(2) Equipement - 18 car", commentaires="Equipement lié", exemple="", longueurMaxi=18, valeursAutorisees=null, forcerMajuscule=false, forcerVide=true, forcerDocumentation=false, règleDeGestion=null },
                        new { entete="(3) N° LICENCE DU PERE - 20 car", commentaires="Licence équipement parent", exemple="", longueurMaxi=20, valeursAutorisees=null, forcerMajuscule=true, forcerVide=true, forcerDocumentation=false, règleDeGestion=null },
                        new { entete="Statut RFOU - 1 car", commentaires="Statut RFOU", exemple="", longueurMaxi=1, valeursAutorisees=null, forcerMajuscule=true, forcerVide=true, forcerDocumentation=false, règleDeGestion=null },
                        new { entete="Statut REF - 1 car", commentaires="Statut REF", exemple="", longueurMaxi=1, valeursAutorisees=null, forcerMajuscule=true, forcerVide=true, forcerDocumentation=false, règleDeGestion=null },
                        new { entete="N° position - 4 car", commentaires="Numéro de poste", exemple="", longueurMaxi=4, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion="M04.2.J" },
                        new { entete="Groupe autorisation - 4 car", commentaires="Groupe d'autorisation : SEQR (RE00), SEQD (autre division)", exemple="", longueurMaxi=4, valeursAutorisees=groupe_autorisation, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null },
                        
                        new { entete="Catégorie équipement - 1 car (*)", commentaires="Catégorie équipement : N, I, R", exemple="N", longueurMaxi=1, valeursAutorisees=categorie_equipement, forcerMajuscule=true, forcerVide=false, forcerDocumentation=true, règleDeGestion=null },
                        new { entete="Libellé fonctionnel de l'équip - 40 car", commentaires="Désignation fonctionnelle", exemple="", longueurMaxi=40, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null },
                        new { entete="Numéro de série fabricant - 30 car", commentaires="S/N Fabricant : non documenté pour un équipement de type R", exemple="", longueurMaxi=30, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null },
                        new { entete="Type équipement - 10 car", commentaires="Type d'équipement : SMN-REG, SMN-CSR", exemple="SMN-REG", longueurMaxi=10, valeursAutorisees=type_equipement, forcerMajuscule=false, forcerVide=false, forcerDocumentation=false, règleDeGestion=null },
                        new { entete="N° inventaire - 25 car", commentaires="Numéro d'inventaire", exemple="", longueurMaxi=25, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null },
                        new { entete="Code ABC - 1 car", commentaires="Criticité ABC : si non documenté, il sera mis la valeur 3 - 1, 2, 3", exemple="1", longueurMaxi=1, valeursAutorisees=abc, forcerMajuscule=true, forcerVide=false, forcerDocumentation=true, règleDeGestion=null },
                        new { entete="Localisation - 10 car", commentaires="Localisation technique", exemple="", longueurMaxi=10, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null },
                        new { entete="Local - 8 car", commentaires="Local", exemple="", longueurMaxi=8, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null },
                        new { entete="Centre de coût - 10 car", commentaires="Centre de coût SAP", exemple="AC01130", longueurMaxi=10, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null },
                        new { entete="Immobilisation principale - 12 car", commentaires="Immobilisation principale : non documenté pour un équipement de type R", exemple="", longueurMaxi=12, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null },
                        new { entete="Immobilisation subsidiaire - 4 car", commentaires="Immobilisation subsidiaire : non documenté pour un équipement de type R", exemple="", longueurMaxi=4, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null },
                        
                        new { entete="Valeur d'acquisition - 17 car", commentaires="Valeur d'acquisition", exemple="", longueurMaxi=17, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion="M04.2.W" },
                        new { entete="Devise - 5 car", commentaires="Devise : non documenté pour un équipement de type R", exemple="EUR", longueurMaxi=5, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null },
                        new { entete="Date d'acquisition - 8 car", commentaires="Date d'acquisition (JJMMAAAA)", exemple="10091969", longueurMaxi=8, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion="M04.2.Y" },
                        new { entete="Date début garanti - 8 car", commentaires="Début de garantie (JJMMAAAA)", exemple="10091969", longueurMaxi=8, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion="M04.2.Z" },
                        new { entete="Date fin garanti - 8 car", commentaires="Fin de garantie (JJMMAAAA)", exemple="10091969", longueurMaxi=8, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion="M04.2.AA" },
                        new { entete="Repère - 30 car", commentaires="Repère équipement : non documenté pour un équipement de type R", exemple="", longueurMaxi=30, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null },
                        new { entete="N° LICENCE - 24 car", commentaires="Numéro de licence : non documenté pour un équipement de type R", exemple="", longueurMaxi=24, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null },
                        new { entete="Code MABEC - 18 car", commentaires="Code MABEC", exemple="", longueurMaxi=18, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion="M04.2.AD" },
                        new { entete="Libellé matériel de l'équipement - 30 car (*)", commentaires="Libellé matériel", exemple="", longueurMaxi=30, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=true, règleDeGestion=null },
                        new { entete="Niveau équipement - 3 car (*)", commentaires="Niveau de l'équipement : GE, E, S/E", exemple="S/E", longueurMaxi=3, valeursAutorisees=niveau_equipement, forcerMajuscule=true, forcerVide=false, forcerDocumentation=true, règleDeGestion=null },
                        
                        new { entete="Référence fournisseur - 25 car (*)", commentaires="Réf fournisseur", exemple="", longueurMaxi=25, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=true, règleDeGestion=null },
                        new { entete="Nom fournisseur - 30 car (*)", commentaires="Nom fournisseur", exemple="", longueurMaxi=30, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=true, règleDeGestion=null },
                        new { entete="Référence intégrateur - 30 car", commentaires="Réf intégrateur", exemple="", longueurMaxi=30, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null },
                        new { entete="Nom intégrateur - 30 car", commentaires="Nom intégrateur", exemple="", longueurMaxi=30, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null },
                        new { entete="Quantité équipement - 17 car", commentaires="Quantité : si non documenté, il sera mis la valeur 1 par défaut", exemple="1", longueurMaxi=17, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion="M04.2.AK" },
                        new { entete="Mnémonique - 10 car", commentaires="Mnémonique", exemple="", longueurMaxi=10, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null },
                        new { entete="Nature d'équipement - 1 car (*)", commentaires="Nature équipement : C=Commerce, F=Fournisseur, B=Renault, R=Standard", exemple="C", longueurMaxi=1, valeursAutorisees=nature_equipement, forcerMajuscule=true, forcerVide=false, forcerDocumentation=true, règleDeGestion=null },
                        new { entete="Code Projet - 30 car", commentaires="Référence projet", exemple="", longueurMaxi=30, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null },
                        new { entete="Modèle - 25 car", commentaires="Modèle fabricant", exemple="", longueurMaxi=25, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null },
                        
                        new { entete="Famille - 6 car (*)", commentaires="Famille équipement SAP", exemple="", longueurMaxi=6, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=true, règleDeGestion="M04.2.AP" },
                        new { entete="Capacité - 25 car", commentaires="Capacité", exemple="", longueurMaxi=25, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null },
                        new { entete="Alimentation - 25 car", commentaires="Alimentation", exemple="", longueurMaxi=25, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null },
                        new { entete="A maintenir - 1 car", commentaires="Précise si une maintenance est nécessaire : 0, 1", exemple="1", longueurMaxi=1, valeursAutorisees=a_maintenir, forcerMajuscule=true, forcerVide=false, forcerDocumentation=true, règleDeGestion=null },
                        new { entete="Uet de Fabrication - 30 car", commentaires="UET de fabrication", exemple="", longueurMaxi=30, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null },
                        
                        new { entete="Dessiné par - 30 car", commentaires="Dessiné par", exemple="", longueurMaxi=30, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null },
                        new { entete="Indice Inventaire - 30 car", commentaires="Indice inventaire", exemple="", longueurMaxi=30, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null },
                        new { entete="Date de l'indice - 8 car", commentaires="Date de l'indice (JJMMAAAA)", exemple="10091969", longueurMaxi=8, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion="M04.2.AW" },
                        new { entete="Responsable de l'indice - 30 car", commentaires="Responsable indice", exemple="", longueurMaxi=30, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null },
                        new { entete="N° pièce produit (1) - 30 car", commentaires="Pièce produit 1", exemple="", longueurMaxi=30, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null },
                        new { entete="Indice pièce produit (1) - 30 car", commentaires="Indice produit 1", exemple="", longueurMaxi=30, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null }
                        new { entete="N° pièce produit (2) - 30 car", commentaires="Pièce produit 2", exemple="", longueurMaxi=30, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null },
                        new { entete="Indice pièce produit (2) - 30 car", commentaires="Indice produit 2", exemple="", longueurMaxi=30, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null }
                        new { entete="N° pièce produit (3) - 30 car", commentaires="Pièce produit 3", exemple="", longueurMaxi=30, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null },
                        new { entete="Indice pièce produit (3) - 30 car", commentaires="Indice produit 3", exemple="", longueurMaxi=30, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null }
                        new { entete="N° pièce produit (4) - 30 car", commentaires="Pièce produit 4", exemple="", longueurMaxi=30, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null },
                        new { entete="Indice pièce produit (4) - 30 car", commentaires="Indice produit 4", exemple="", longueurMaxi=30, valeursAutorisees=null, forcerMajuscule=true, forcerVide=false, forcerDocumentation=false, règleDeGestion=null }
                    }

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

