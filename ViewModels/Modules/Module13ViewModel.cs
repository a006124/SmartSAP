using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Windows.Input;

namespace SmartSAP.ViewModels.Modules
{
    // Test GCP
    public class Module13ViewModel : ModuleDetailViewModelBase
    {
        public ICommand ExecuteGCPTransactionCommand { get; protected set; }

        public Module13ViewModel(MainViewModel mainViewModel, string title)
            : base(mainViewModel, title)
        {
            ExecuteGCPTransactionCommand = new RelayCommand(async p => await ExecuteGCPTransactionAsync(p as WorkflowStep));
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
                    Title = "1. Saisie de la liste des FID à contrôler dans GCP",
                    Description = "Crée un nouveau fichier Excel modèle.",
                    Icon = "\xE70F",
                    ModuleStep = "M13-E1",
                    OpenFile = true,
                    ActionCommand = GenerateTemplateCommand
                },
                new WorkflowStep {
                    Title = "2. Exécution de la transaction GCP",
                    Description = "Exécute la requête SQL GCP et remplit l'Excel.",
                    Icon = "\xE768",
                    ModuleStep = "M13-E2",                  
                    ActionCommand = ExecuteGCPTransactionCommand
                }
            };
        }


        protected async Task ExecuteGCPTransactionAsync(WorkflowStep? step = null)
        {
            if (step == null)
            {
                step = Steps.FirstOrDefault(s => s.ActionCommand == ExecuteGCPTransactionCommand);
            }

            try
            {
                Logs.Add(new LogEntry("INFO", "Vérification du fichier modèle..."));

                if (string.IsNullOrEmpty(LastGeneratedExcelPath) || !File.Exists(LastGeneratedExcelPath))
                {
                    Logs.Add(new LogEntry("ERROR", "Le fichier de données Excel modèle est introuvable. Veuillez d'abord exécuter l'étape 1."));
                    if (step != null) { step.Status = "Erreur Fichier"; step.ResultState = "Error"; }
                    return;
                }

                Logs.Add(new LogEntry("INFO", "Exécution de la requête SQL sur GCP BigQuery..."));

                // Paramètres du projet GCP (À adapter par l'utilisateur)
                string projectId = "votre_projet_id"; 
                string location = "eu"; 
                string query = @"
                    -- Remplacer par votre requête SQL:
                    -- Il est conseillé de renommer vos colonnes (AS `Nom de la colonne`)
                    -- pour qu'elles correspondent exactement aux en-têtes définis dans InitializedExcelColumns
                    SELECT 
                        'S123456' AS `Nom AVEC LE S DEVANT *`,
                        'ZDOC' AS `Type *`,
                        '00' AS `Version`,
                        'Test depuis GCP' AS `Description`,
                        'Validé' AS `Statut`
                ";

                // Exécution
                var results = await SmartSAP.Services.GCP.GCPManager.ExecuteQueryAsync(projectId, location, query);

                if (results.Count == 0)
                {
                    Logs.Add(new LogEntry("WARNING", "La requête s'est exécutée avec succès mais n'a retourné aucun résultat."));
                    if (step != null) { step.Status = "Succès - 0 ligne"; step.ResultState = "Success"; }
                    return;
                }

                Logs.Add(new LogEntry("INFO", $"{results.Count} ligne(s) récupérée(s) depuis GCP. Écriture dans le fichier Excel modèle..."));

                using (var workbook = new XLWorkbook(LastGeneratedExcelPath))
                {
                    var worksheet = workbook.Worksheets.FirstOrDefault();
                    if (worksheet == null)
                    {
                        Logs.Add(new LogEntry("ERROR", "Le fichier Excel ne contient aucune feuille."));
                        if (step != null) { step.Status = "Erreur Fichier"; step.ResultState = "Error"; }
                        return;
                    }

                    // On efface les données existantes (hors en-tête ligne 1) s'il y en a
                    int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
                    if (lastRow > 1)
                    {
                        worksheet.Range(2, 1, lastRow, ExcelColumns.Count).Clear();
                    }

                    // Transformation des résultats pour l'insertion
                    var dataList = new List<object[]>();
                    foreach (var rowDict in results)
                    {
                        var rowArray = new object[ExcelColumns.Count];
                        for (int i = 0; i < ExcelColumns.Count; i++)
                        {
                            var headerTitle = ExcelColumns[i].Entete;
                            if (rowDict.ContainsKey(headerTitle))
                            {
                                rowArray[i] = rowDict[headerTitle]?.ToString() ?? "";
                            }
                            else
                            {
                                rowArray[i] = "";
                            }
                        }
                        dataList.Add(rowArray);
                    }

                    // Insertion dans ClosedXML (Ligne 2, Colonne 1)
                    if (dataList.Count > 0)
                    {
                        worksheet.Cell(2, 1).InsertData(dataList);
                    }

                    workbook.Save();
                }

                Logs.Add(new LogEntry("SUCCESS", $"✓ Terminé avec succès. {results.Count} ligne(s) écrite(s) dans le fichier Excel."));
                if (step != null) { step.Status = "Terminé"; step.ResultState = "Success"; }
            }
            catch (Exception ex)
            {
                Logs.Add(new LogEntry("ERROR", $"Erreur lors de l'intégration GCP : {ex.Message}"));
                if (step != null) { step.Status = "Crash"; step.ResultState = "Error"; }
            }
        }

        // DÉFINITION DES COLONNES DE L'EXCEL MODELE
        protected override void InitializeExcelColumns(WorkflowStep? step = null)
        {
            ExcelColumns.Clear();

            // Chargement des données depuis JSON
            string dataPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data");
            if (!Directory.Exists(dataPath))
                dataPath = Path.Combine(Directory.GetCurrentDirectory(), "Data");

            var type_FID = LoadJsonValues(Path.Combine(dataPath, "type_FID.json"), "type_FID");
            var langue = LoadJsonValues(Path.Combine(dataPath, "langue.json"), "Langue préférée (division)");
            var statut_FID= LoadJsonValues(Path.Combine(dataPath, "statut_FID.json"), "statut_FID");
            var type_FID_R = LoadJsonValues(Path.Combine(dataPath, "type_FID_R.json"), "type_FID_R");
            var version_FID = LoadJsonValues(Path.Combine(dataPath, "version_FID.json"), "version_FID");
            var action_lien_FID = LoadJsonValues(Path.Combine(dataPath, "action_lien_FID.json"), "action_lien_FID");
            var type_objet_lien_FID = LoadJsonValues(Path.Combine(dataPath, "type_objet_lien_FID.json"), "type_objet_lien_FID");
            var action_original_FID = LoadJsonValues(Path.Combine(dataPath, "action_original_FID.json"), "action_original_FID");
            var application_original_FID = LoadJsonValues(Path.Combine(dataPath, "application_original_FID.json"), "application_original_FID");



            var ExcelModel = new List<ExcelColumnModel>
            {
                // Entete - Commentaires - Données d'exemple - Longueur maxi - Valeurs autorisées - Majuscules forcées - Vide forcé - Documentation forcée - Règle de gestion
                new ("Nom AVEC LE S DEVANT *", "", "", 100, null, true, false, false, ""),
                new ("Type *", "", "", 100, null, true, false, false, ""),
                new ("Version", "", "", 100, null, true, false, false, ""),
                new ("Description", "", "", 100, null, true, false, false, ""),
                new ("Statut", "", "", 100, null, true, false, false, ""),
                new ("Liaison Article (Code)", "", "", 100, null, true, false, false, ""),
                new ("Liaison Article (Désignation)", "", "", 100, null, true, false, false, ""),
                new ("Liaison Equipement (Code)", "", "", 100, null, true, false, false, ""),
                new ("Liaison Equipement (Désignation)", "", "", 100, null, true, false, false, ""),
                new ("Liaison Poste Technique (Code)", "", "", 100, null, true, false, false, ""),
                new ("Liaison Poste Technique (Désignation)", "", "", 100, null, true, false, false, ""),
                new ("Originaux (Nombre)", "", "", 100, null, true, false, false, ""),
                new ("Originaux (Fichier)", "", "", 100, null, true, false, false, ""),
            };

            var columnsToAdd = ExcelModel.Select(d =>
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

            foreach (var col in columnsToAdd)
            {
                ExcelColumns.Add(col);
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
