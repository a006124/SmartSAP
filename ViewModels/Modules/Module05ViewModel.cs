using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;

namespace SmartSAP.ViewModels.Modules
{
    // Equipement : Modification en masse
    public class Module05ViewModel : ModuleDetailViewModelBase
    {
        public Module05ViewModel(MainViewModel mainViewModel, string title) 
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


        protected override void InitializeSteps()
        {
            Steps = new ObservableCollection<WorkflowStep>
            {
                new WorkflowStep { 
                    Title = "[Option1] SAP->Excel E1.1", 
                    Description = "Crée un fichier Excel pour saisir le numéros d'équipement à exporter.", 
                    Icon = "\xE70F", 
                    ModuleStep = "M05-E1.1",
                    OpenFile = true,
                    ActionCommand = GenerateTemplateCommand 
                },
                new WorkflowStep { 
                    Title = "[Option1] SAP->Excel E1.2", 
                    Description = "Contrôle et exporte les données (Format SAP). ", 
                    Icon = "\xE762", 
                    ModuleStep = "M05-E1.2",
                    OpenFile = false,
                    ActionCommand = ExportFixedWidthCommand
                },
                new WorkflowStep { 
                    Title = "[Option1] SAP->Excel E1.3", 
                    Description = "Récupère les données des équipements via la transaction SAP 'IH08'.", 
                    Icon = "\xE768", 
                    ModuleStep = "M05-E1.3",
                    OpenFile = true,
                    ActionCommand = ExecuteSAPTransactionCommand
                },
                new WorkflowStep { 
                    Title = "[Option2] Modèle vierge", 
                    Description = "Crée un fichier Excel modèle.", 
                    Icon = "\xE70F", 
                    ModuleStep = "M05-E2",
                    OpenFile = true,
                    ActionCommand = GenerateTemplateCommand 
                },
                new WorkflowStep { 
                    Title = "3. Contrôle et export des données", 
                    Description = "Contrôle et exporte les données (Format SAP). ", 
                    Icon = "\xE762", 
                    ModuleStep = "M05-E3",
                    OpenFile = false,
                    ActionCommand = ExportFixedWidthCommand
                },
                new WorkflowStep { 
                    Title = "4. Intégration des modifications dans SAP", 
                    Description = "Exécute la transaction SAP 'ZSMNBAO13'.", 
                    Icon = "\xE768", 
                    ModuleStep = "M05-E4",
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
                string sapTx = step?.ModuleStep == "M05-E1.3" ? "IH08" : "ZSMNBAO13";
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
                        
                        // 3. Traitement du fichier Excel si étape E1.3
                        if (step?.ModuleStep == "M05-E1.3")
                        {
                            Logs.Add(new LogEntry("INFO", "Génération du modèle E2 pour enrichissement..."));
                            
                            // 1. Fichier modèle type M05-E2 créé
                            var e2Step = Steps.FirstOrDefault(s => s.ModuleStep == "M05-E2") ?? new WorkflowStep { ModuleStep = "M05-E2" }; ; e2Step.OpenFile = false;
                            GenerateExcelTemplate(e2Step);
                            string templateE2Path = LastGeneratedExcelPath;

                            if (!string.IsNullOrEmpty(templateE2Path) && System.IO.File.Exists(templateE2Path))
                            {
                                // 2. Que la fonction EnrichirFromSAPExcelWorkbook soit exécutée
                                try
                                {
                                    var excelService = new SmartSAP.Services.Excel.ExcelManager();
                                    string enrichResult = excelService.EnrichirFromSAPExcelWorkbookM05_E_1_3(templateE2Path, resultFile);
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
            ExcelColumns.Clear();

            // Chargement des données depuis JSON
            string dataPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data");
            // Note: En mode Debug/Développement, le chemin peut varier, on essaie aussi le chemin relatif au projet
            if (!Directory.Exists(dataPath))
                dataPath = Path.Combine(Directory.GetCurrentDirectory(), "Data");

            switch (step?.ModuleStep)
            {
                case "M05-E1.1":
                case "M05-E1.2":
                    // LISTE DE NUMÉROS D'ÉQUIPEMENTS
                    // Header - Commentaire - Données d'exemple - Largeur fixe - Majuscules forcées - Valeurs autorisées
                    var ExcelModel = new[]
                    {
                        new ExcelColumnModel("N° Equ SAP - 18 car", "Numéro équipement SAP", "", 18, null, true, false, true, "M05.1.2.A")
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
                case "M05-E2":
                case "M05-E3":
                    // DONNÉES COMPLÈTES DES ÉQUIPEMENTS
                    // Header - Commentaire - Données d'exemple - Largeur fixe - Majuscules forcées - Valeurs autorisées

                    var divisions = LoadJsonValues(Path.Combine(dataPath, "division.json"), "01-Division Localisation");
                    var langues = LoadJsonValues(Path.Combine(dataPath, "langue.json"), "Langue préférée (division)");
                    var groupe_autorisation = LoadJsonValues(Path.Combine(dataPath, "groupe_autorisation.json"), "groupe_autorisation");
                    var categorie_equipement = LoadJsonValues(Path.Combine(dataPath, "categorie_equipement.json"), "categorie_equipement");
                    var type_equipement = LoadJsonValues(Path.Combine(dataPath, "type_equipement.json"), "type_equipement");
                    var abc = LoadJsonValues(Path.Combine(dataPath, "abc.json"), "abc");
                    var niveau_equipement = LoadJsonValues(Path.Combine(dataPath, "niveau_equipement.json"), "niveau_equipement");
                    var nature_equipement = LoadJsonValues(Path.Combine(dataPath, "nature_equipement.json"), "nature_equipement");
                    var a_maintenir = LoadJsonValues(Path.Combine(dataPath, "a_maintenir.json"), "a_maintenir");

                    var ExcelModelFull = new List<ExcelColumnModel>
                     {
                        // Entete - Commentaires - Données d'exemple - Longueur maxi - Valeurs autorisées - Majuscules forcées - Vide forcé - Documentation forcée - Règle de gestion
                        new ("Division - 4 car (*)","Documenter le code suivant les divisions gérées dans SAP","MC02",4,divisions,true,false,true,""),
                        new ("Langue - 2 car (*)","Documenter le code correspondant à la langue utilisée dans la Désignation","FR",2,langues,true,false,true,""),
                        new ("N° Equ SAP - 18 car","Numéro équipement SAP","",18,null,true,false,true,"M05.3.C"),
                        new ("N° EQU LICENCE - 20 car","Numéro licence équipement","",20,null,true,false,false,""),
                        new ("(1) Poste technique - 30 car","Poste technique lié","",30,null,true,false,false,""),
                        new ("(2) Equipement - 18 car","Equipement lié","",18,null,false,false,false,""),
                        new ("(3) N° LICENCE DU PERE - 20 car","Licence équipement parent","",20,null,true,false,false,""),
                        new ("Statut RFOU - 1 car","Statut RFOU","",1,null,true,false,false,""),
                        new ("Statut REF - 1 car","Statut REF","",1,null,true,false,false,""),
                        new ("N° position - 4 car","Numéro de poste","",4,null,true,false,false,"M05.3.J"),
                        new ("Groupe autorisation - 4 car","Groupe d'autorisation : SEQR (RE00), SEQD (autre division)","",4,groupe_autorisation,true,false,false,""),
                        new ("Catégorie équipement - 1 car (*)","Catégorie équipement : N, I, R","N",1,categorie_equipement,true,false,true,""),
                        new ("Libellé fonctionnel de l'équip - 40 car","Désignation fonctionnelle","",40,null,true,false,false,""),
                        new ("Numéro de série fabricant - 30 car","S/N Fabricant : non documenté pour un équipement de type R","",30,null,true,false,false,""),
                        new ("Type équipement - 10 car","Type d'équipement : SMN-REG, SMN-CSR","SMN-REG",10,type_equipement,false,false,false,""),
                        new ("N° inventaire - 25 car","Numéro d'inventaire","",25,null,true,false,false,""),
                        new ("Code ABC - 1 car","Criticité ABC : si non documenté, il sera mis la valeur 3 - 1, 2, 3","1",1,abc,true,false,true,""),
                        new ("Localisation - 10 car","Localisation technique","",10,null,true,false,false,""),
                        new ("Local - 8 car","Local","",8,null,true,false,false,""),
                        new ("Centre de coût - 10 car","Centre de coût SAP","AC01130",10,null,true,false,false,""),
                        new ("Immobilisation principale - 12 car","Immobilisation principale : non documenté pour un équipement de type R","",12,null,true,false,false,""),
                        new ("Immobilisation subsidiaire - 4 car","Immobilisation subsidiaire : non documenté pour un équipement de type R","",4,null,true,false,false,""),
                        new ("Valeur d'acquisition - 17 car","Valeur d'acquisition","",17,null,true,false,false,"M05.3.W"),
                        new ("Devise - 5 car","Devise : non documenté pour un équipement de type R","EUR",5,null,true,false,false,""),
                        new ("Date d'acquisition - 8 car","Date d'acquisition (JJMMAAAA)","10091969",8,null,true,false,false,"M05.3.Y"),
                        new ("Date début garanti - 8 car","Début de garantie (JJMMAAAA)","10091969",8,null,true,false,false,"M05.3.Z"),
                        new ("Date fin garanti - 8 car","Fin de garantie (JJMMAAAA)","10091969",8,null,true,false,false,"M05.3.AA"),
                        new ("Repère - 30 car","Repère équipement : non documenté pour un équipement de type R","",30,null,true,false,false,""),
                        new ("N° LICENCE - 20 car","Numéro de licence : non documenté pour un équipement de type R","",20,null,true,false,false,""),
                        new ("Code MABEC - 18 car","Code MABEC","",18,null,true,false,false,"M05.3.AD"),
                        new ("Libellé matériel de l'équipement - 30 car (*)","Libellé matériel","",30,null,true,false,true,""),
                        new ("Niveau équipement - 3 car (*)","Niveau de l'équipement : GE, E, S/E","S/E",3,niveau_equipement,true,false,true,""),
                        new ("Référence fournisseur - 25 car (*)","Réf fournisseur","",25,null,true,false,true,""),
                        new ("Nom fournisseur - 30 car (*)","Nom fournisseur","",30,null,true,false,true,""),
                        new ("Référence intégrateur - 30 car","Réf intégrateur","",30,null,true,false,false,""),
                        new ("Nom intégrateur - 30 car","Nom intégrateur","",30,null,true,false,false,""),
                        new ("Quantité équipement - 17 car","Quantité : si non documenté, il sera mis la valeur 1 par défaut","1",17,null,true,false,false,"E05.AK"),
                        new ("Mnémonique - 10 car","Mnémonique","",10,null,true,false,false,""),
                        new ("Nature d'équipement - 1 car (*)","Nature équipement : C=Commerce, F=Fournisseur, B=Renault, R=Standard","C",1,nature_equipement,true,false,true,""),
                        new ("Code Projet - 30 car","Référence projet","",30,null,true,false,false,""),
                        new ("Modèle - 25 car","Modèle fabricant","",25,null,true,false,false,""),
                        new ("Famille - 6 car (*)","Famille équipement SAP","",6,null,true,false,true,""),
                        new ("Capacité - 25 car","Capacité","",25,null,true,false,false,""),
                        new ("Alimentation - 25 car","Alimentation","",25,null,true,false,false,""),
                        new ("A maintenir - 1 car","Précise si une maintenance est nécessaire : 0, 1","1",1,a_maintenir,true,false,true,""),
                        new ("Uet de Fabrication - 30 car","UET de fabrication","",30,null,true,false,false,""),
                        new ("Dessiné par - 30 car","Dessiné par","",30,null,true,false,false,""),
                        new ("Indice Inventaire - 30 car","Indice inventaire","",30,null,true,false,false,""),
                        new ("Date de l'indice - 8 car","Date de l'indice (JJMMAAAA)","10091969",8,null,true,false,false,"M05.3.AW"),
                        new ("Responsable de l'indice - 30 car","Responsable indice","",30,null,true,false,false,""),
                        new ("N° pièce produit (1) - 30 car","Pièce produit 1","",30,null,true,false,false,""),
                        new ("Indice pièce produit (1) - 30 car","Indice produit 1","",30,null,true,false,false,""),
                        new ("N° pièce produit (2) - 30 car","Pièce produit 2","",30,null,true,false,false,""),
                        new ("Indice pièce produit (2) - 30 car","Indice produit 2","",30,null,true,false,false,""),
                        new ("N° pièce produit (3) - 30 car","Pièce produit 3","",30,null,true,false,false,""),
                        new ("Indice pièce produit (3) - 30 car","Indice produit 3","",30,null,true,false,false,""),
                        new ("N° pièce produit (4) - 30 car","Pièce produit 4","",30,null,true,false,false,""),
                        new ("Indice pièce produit (4) - 30 car","Indice produit 4","",30,null,true,false,false,""),
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