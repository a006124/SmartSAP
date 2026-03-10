namespace SmartSAP.Models
{
    public class ExcelColumnDefinition
    {
        // Entête de la colonne dans le fichier Excel
        public string Entete { get; set; } = string.Empty;
        // Commentaires de la cellule pour l'utilisateur concernant cette colonne
        public string Commentaires { get; set; } = string.Empty;
        // Exemple de valeur attendue pour cette colonne
        public string Exemple { get; set; } = string.Empty;
        // Longueur maximale autorisée pour les valeurs de cette colonne
        public int LongueurMaxi { get; set; }
        // Liste des valeurs autorisées pour cette colonne (null si aucune restriction)
        public string[]? ValeursAutorisées { get; set; }
        // Indique si les valeurs de cette colonne doivent être converties en majuscules
        public bool ForcerMajuscule { get; set; } = true;
        // Indique si les cellules de cette colonne doivent être forcées à être vides (true) ou non (false)
        public bool ForcerVide { get; set; } = true;
        // Indique si la documentation est obligatoire pour cette colonne (true) ou non (false)
        public bool ForcerDocumentation { get; set; } = true;

        // Précise la ou les règle(s) de gestion spécifique(s) (vide si pas de règle)
        public string[]? RègleDeGestion { get; set; } = null;

        public ExcelColumnDefinition(string entete, 
                                     string commentaires, 
                                     string exemple, 
                                     int longueurMaxi = 0,
                                     string[]? valeursAutorisées = null,
                                     bool forcerMajuscule = true, 
                                     bool forcerVide = false,
                                     bool forcerDocumentation = false,
                                     string[]? règleDeGestion = null)
        {
            Entete = entete;
            Commentaires = commentaires;
            Exemple = exemple;
            LongueurMaxi = longueurMaxi;
            ValeursAutorisées = valeursAutorisées;
            ForcerMajuscule = forcerMajuscule;
            ForcerVide = forcerVide;
            ForcerDocumentation = forcerDocumentation;
            RègleDeGestion = règleDeGestion;
        }
    }
}
