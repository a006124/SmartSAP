using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Cloud.BigQuery.V2;

namespace SmartSAP.Services.GCP
{
    public class GCPManager
    {
        private static readonly string[] Scopes = new[] { "https://www.googleapis.com/auth/bigquery" };

        /// <summary>
        /// Initialise une connexion BigQuery en cherchant le fichier de secret (client_secret*.json).
        /// </summary>
        public static async Task<BigQueryClient> GetClientAsync(string projectId, string location = null)
        {
            string secretFile = FindSecretFile();
            if (string.IsNullOrEmpty(secretFile))
            {
                throw new FileNotFoundException("Impossible de trouver le fichier client_secret*.json pour l'authentification GCP.");
            }

            using (var stream = new FileStream(secretFile, FileMode.Open, FileAccess.Read))
            {
                var credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None);

                return await BigQueryClient.CreateAsync(projectId, credential);
            }
        }

        /// <summary>
        /// Cherche le fichier secret dans le dossier courant ou les dossiers parents.
        /// </summary>
        private static string FindSecretFile()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            
            // 1. Chercher dans le dossier de l'exécutable
            var files = Directory.GetFiles(baseDir, "client_secret*.json");
            if (files.Length > 0)
                return files[0];

            // 2. Chercher dans le dossier parent (pratique pour l'exécution depuis l'IDE)
            string currentDir = baseDir;
            for (int i = 0; i < 4; i++)
            {
                currentDir = Directory.GetParent(currentDir)?.FullName;
                if (currentDir != null)
                {
                    files = Directory.GetFiles(currentDir, "client_secret*.json");
                    if (files.Length > 0)
                        return files[0];
                }
                else
                {
                    break;
                }
            }

            return null;
        }

        /// <summary>
        /// Exécute une requête SQL sur BigQuery et renvoie les résultats sous forme de liste de dictionnaires.
        /// Chaque dictionnaire représente une ligne (Nom de colonne -> Valeur).
        /// </summary>
        public static async Task<List<Dictionary<string, object>>> ExecuteQueryAsync(string projectId, string location, string query)
        {
            var client = await GetClientAsync(projectId, location);
            
            // Exécution de la requête
            BigQueryJob job = await client.CreateQueryJobAsync(query, null);
            job = await job.PollUntilCompletedAsync();
            BigQueryResults results = await job.GetQueryResultsAsync();

            var data = new List<Dictionary<string, object>>();

            foreach (var row in results)
            {
                var rowData = new Dictionary<string, object>();
                foreach (var field in results.Schema.Fields)
                {
                    var value = row[field.Name];
                    rowData[field.Name] = value;
                }
                data.Add(rowData);
            }

            return data;
        }
    }
}
