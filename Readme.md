# Vue d'ensemble

**EDB_Frais_DSN.ps1** est un script PowerShell de traitement et consolidation des données de frais EDB pour la DSN.

Le script lit les données de plusieurs sources, consolide les frais par matricule et par catégorie, et génère un fichier CSV formaté pour l'intégration DSN.

* Sources de données : fichiers Excel (Frais Policy) et CSV (Matricules de paie)
* Génère le fichier **final.csv** avec les frais consolidés au format DSN

## Comment installer ce script

* PowerShell 7 ou supérieur

Récupérer le script sur GitLAB et déposer les fichiers dans un répertoire du serveur de scripts.

### Modules externes

Récupérer les modules nécessaires sur GitLAB et les déposer dans le répertoire `Modules/` du script.

* Modules PowerShell requis :
  - **Ini.ps1** : Gestion des fichiers de configuration .ini
  - **Log.ps1** : Gestion des logs et traces
  - **StrConvert.ps1** : Conversion et manipulation de chaînes
  - **SendEmail.ps1** : Envoi d'emails de notification
  - **XLSX.ps1** : Lecture et requêtage de fichiers Excel
  - **Csv.ps1** : Lecture et export de fichiers CSV

Paramétrer le fichier **EDB_Frais_DSN.ini** avant la première exécution.

## Structure du script

Le script suit une architecture modulaire avec :
1. **Chargement des modules** : Import des modules PowerShell nécessaires
2. **Initialisation** : Chargement du fichier de configuration `.ini`
3. **Configuration de l'encodage** : Passage de la console en UTF-8
4. **Récupération des paramètres** : Chargement des fichiers sources et configuration
5. **Traitement des données** : Exécution séquentielle des traitements de frais par catégories
6. **Consolidation** : Agrégation des frais par matricule et catégorie
7. **Génération du CSV** : Export du fichier final
8. **Finalisation** : Affichage du résumé et envoi optionnel d'email

## Sources des données

Le script traite les données provenant de trois sources :

### Fichiers Excel - Frais Policy (2 sources)

| Section [.ini] | Fichier Excel | Paramètre |
| --- | --- | --- |
| `[XLS_Frais_Policy_FORTIL_GROUP]` | DRH - Frais Policy_FORTIL GROUP_YTD_*.xls | fichierXLS |
| `[XLS_Frais_Policy_FUSION]` | DRH - Frais Policy_Fusion_Group_YTD_*.xls | fichierXLS |

**Colonnes lues :**
- Matricule (RH)
- Catégorie (Description)
- Date de validation
- Montant
- Statut (filtré sur PMNT_NOT_PAID)
- statusId2 (filtré sur EXPENSE_ACCEPTED)

### Fichier CSV - Matricules de paie

| Section [.ini] | Fichier CSV | Paramètre |
| --- | --- | --- |
| `[CSV_REQ20]` | req20_Matricules_paie-*.csv | fichierCSV |

**Colonnes lues :**
- Matricule (clé d'indexation)
- matricule paie
- D Fin contrat (gestion des dates)

## Fichier CSV final

| Paramètre [.ini] | Fichier généré |
| --- | --- |
| `[FINAL][FichierCSV]` | final.csv |

**Structure :**
```
Code PAC;Mois de paie;Matricule paie;S21.G00.54.001;S21.G00.54.002;S21.G00.54.003;S21.G00.54.004
```

## Principe du traitement

À partir des fichiers sources, traiter les données et **générer une hashtable consolidée** avec la même structure que le fichier `[FINAL][FichierCSV]` à générer.

### Flux de traitement

1. **LoadIni** : Charge le fichier de configuration `EDB_Frais_DSN.ini`
   - Initialise les variables de log et d'erreur
   - Vérifie l'existence du fichier `.ini`
   - Résout les chemins des fichiers sources
   - Supprime le fichier log One_Shot (nouveau traitement)
   - Crée les fichiers de log nécessaires
   - Valide l'existence des fichiers sources obligatoires
   - Gère les paramètres de période : CURRENT, PREVIOUS, ou yyyyMM

2. **SetConsoleToUFT8** : Configure l'encodage de la console en UTF-8

3. **Compute_CategorieFrais** : Construit la table de correspondance des catégories
   - Mappe les catégories décrites dans `[Categories]` vers les codes DSN dans `[S21.G00.54.001]`
   - Gère les catégories ignorées (valeur "Ignore")
   - Stocke le résultat dans `$script:Categorie`

4. **Query_CSV_REQ20** : Charge les données du fichier CSV de matricules
   - Utilise `Invoke-CSVQuery` avec séparateur `,`
   - Démarre la lecture à la ligne définie par `[CSV_REQ20][HEADERstartline]`
   - Indexe les données par le champ "Matricule"
   - Gère le formatage des dates pour la colonne "D Fin contrat"
   - Stocke le résultat dans `$script:REQ20`

5. **Query_XLS_Frais_Policy** : Charge les données des deux fichiers Excel
   - Traite **FORTIL GROUP** : `[XLS_Frais_Policy_FORTIL_GROUP][fichierXLS]`
   - Traite **FUSION** : `[XLS_Frais_Policy_FUSION][fichierXLS]`
   - Utilise `Invoke-ExcelQuery` pour requêter chaque feuille
   - Filtre sur : Statut = 'PMNT_NOT_PAID' AND statusId2 = 'EXPENSE_ACCEPTED'
   - Trie les résultats par Matricule
   - Appelle `Consolide_Frais` pour chaque source

6. **Consolide_Frais** : Consolide les frais par matricule et catégorie
   - Filtre sur la période définie dans `[FINAL][Mois de paie]`
   - Valide chaque ligne :
     * Exclut les frais si la date de validation ne correspond pas à la période
     * Exclut les catégories avec valeur "Ignore"
     * Exclut les montants nuls
     * Valide la présence du matricule RH dans REQ20
   - Accumule les montants par matricule RH et par catégorie
   - Calcule les dates de début et fin du mois de validation
   - Établit la correspondance avec le matricule de paie
   - Génère des erreurs pour les matricules paie non trouvés
   - Stocke les données consolidées dans `$script:FRAIS`

7. **Extract_Final** : Génère la hashtable finale
   - Parcourt tous les matricules RH consolidés
   - Pour chaque catégorie de frais d'un matricule :
     * Crée une ligne dans `$script:FINAL`
     * Applique l'arrondi à 2 décimales
   - Compile les données dans la structure du fichier de sortie CSV
   - Stocke le résultat dans `$script:FINAL`

8. **FinalToCSV** : Exporte la hashtable finale vers le fichier CSV
   - Utilise la fonction `ExportCsv` du module `Csv.ps1`
   - Applique le délimiteur défini dans `[FINAL][delimiter]`
   - Génère le fichier `[FINAL][FichierCSV]`

9. **QUIT** : Finalise le traitement
   - Affiche le résumé des erreurs et warnings
   - Affiche le temps d'exécution
   - Envoie un email de notification si configuré

**Nota important** : *Ce script ne traite que les frais validés et acceptés. Les matricules RH doivent avoir une correspondance en matricule paie dans le fichier REQ20.*

## Utilisation

### Exécution du script

```powershell
# Depuis le répertoire du script
.\EDB_Frais_DSN.ps1

# Ou avec un chemin absolu
pwsh -File "C:\chemin\vers\EDB_Frais_DSN.ps1"
```

Le script recherche automatiquement le fichier de configuration `EDB_Frais_DSN.ini` dans le même répertoire.

### Configuration

Avant la première exécution, configurer le fichier `EDB_Frais_DSN.ini` :

1. **Section [intf]** : Chemins des fichiers de logs
   - `pathfilelog` : fichier log du traitement courant (supprimé à chaque exécution)
   - `pathfileerr` : fichier cumulatif des erreurs (conservé)

2. **Section [XLS_Frais_Policy_FORTIL_GROUP]** :
   - `fichierXLS` : chemin complet du fichier Excel FORTIL GROUP
   - `Sheet` : nom de la feuille à lire

3. **Section [XLS_Frais_Policy_FUSION]** :
   - `fichierXLS` : chemin complet du fichier Excel FUSION
   - `Sheet` : nom de la feuille à lire

4. **Section [CSV_REQ20]** :
   - `fichierCSV` : chemin complet du fichier CSV des matricules
   - `HEADERstartline` : ligne de démarrage (exemple : 4)

5. **Section [Categories]** :
   - Mapper chaque catégorie de frais de l'Excel vers un code DSN ou "Ignore"
   - Format : `Catégorie Excel = Code DSN` (exemple : `Frais déplacement = 001`)

6. **Section [S21.G00.54.001]** :
   - Définir les codes DSN et leurs descriptions

7. **Section [Code PAC]** :
   - `Code PAC` : code entreprise pour la DSN (exemple : 900380)

8. **Section [FINAL]** :
   - `FichierCSV` : chemin complet du fichier CSV final généré
   - `delimiter` : séparateur (exemple : `;`)
   - `Mois de paie` : Période à traiter
     * `CURRENT` : traite le mois courant
     * `PREVIOUS` : traite le mois précédent
     * `yyyyMM` : traite un mois spécifique (exemple : `202503`)

9. **Section [email]** (optionnel) :
   - `sendemail` : `yes` / `no` pour activer/désactiver l'envoi d'email
   - `destinataire` : adresse email du destinataire
   - `Subject` : sujet de l'email
   - `emailmode` : `SMTP` ou `Microsoft.Graph`
   - Autres paramètres selon le mode d'envoi

### Variables d'environnement

Le script utilise `$rootpath$` comme variable de substitution dans le fichier `.ini` :
- Par défaut : répertoire du script (`$PSScriptRoot`)
- Peut être redéfini dans `[intf][rootpath]`

Exemple :
```ini
fichierXLS = $rootpath$\Fichiers\mon_fichier.xlsx
```

# Détail des fonctions principales

### LoadIni

Charge et initialise le fichier de configuration `EDB_Frais_DSN.ini`.

**Actions réalisées :**
- Initialise les variables de log (`$script:pathfilelog`, `$script:pathfileerr`, etc.)
- Vérifie l'existence du fichier `.ini`
- Charge les paramètres via la fonction `Add-IniFiles`
- Résout les chemins des fichiers avec `GetFilePath`
- Supprime le fichier log One_Shot s'il existe
- Crée les fichiers de log nécessaires
- Valide l'existence des fichiers sources obligatoires
- Traite la période de paie (CURRENT, PREVIOUS ou yyyyMM)

### Compute_CategorieFrais

Construit la table de correspondance entre les catégories de frais et les codes DSN.

**Logique :**
- Parcourt chaque catégorie définie dans `[Categories]`
- Si valeur = "Ignore" : marque comme ignorée
- Sinon : recherche le code DSN correspondant dans `[S21.G00.54.001]`
- Stocke le résultat dans `$script:Categorie`

**Résultat :**
```powershell
$script:Categorie["Frais déplacement"] = "001"
$script:Categorie["Autre frais"] = "Ignore"
```

### Query_CSV_REQ20

Charge en mémoire le fichier CSV `[CSV_REQ20][fichierCSV]` et le convertit en hashtable.

**Traitement spécifique :**
- Utilise `Invoke-CSVQuery` avec séparateur `,`
- Démarre la lecture à la ligne définie par `[CSV_REQ20][HEADERstartline]`
- Indexe les données par le champ "Matricule"
- Gère le formatage des dates pour la colonne "D Fin contrat"
- Stocke le résultat dans `$script:REQ20`

**Résultat :**
```powershell
$script:REQ20["00001234"]["matricule paie"]  # Valeur trouvée
$script:REQ20["00001234"]["D Fin contrat"]   # Date formatée
```

### Query_XLS_Frais_Policy

Charge les données des deux fichiers Excel de frais policy.

**Traitement :**
1. Charge `[XLS_Frais_Policy_FORTIL_GROUP][fichierXLS]`
   - Utilise `Invoke-ExcelQuery` pour exécuter la requête SQL
   - Filtre : Statut = 'PMNT_NOT_PAID' AND statusId2 = 'EXPENSE_ACCEPTED'
   - Appelle `Consolide_Frais` avec les résultats

2. Charge `[XLS_Frais_Policy_FUSION][fichierXLS]`
   - Même traitement que FORTIL GROUP
   - Les données sont consolidées avec celles de FORTIL GROUP

### Consolide_Frais

Consolide les frais par matricule RH et par catégorie de frais.

**Logique de traitement pour chaque ligne :**
1. **Filtrage de la période** : Exclut si la date de validation ≠ période demandée
2. **Filtrage des catégories** : Exclut si catégorie = "Ignore"
3. **Filtrage des montants** : Exclut si montant = 0
4. **Validation du matricule** :
   - Récupère le matricule RH (complété à 8 chiffres à gauche par des zéros)
   - Recherche dans `$script:REQ20`
   - Valide que le matricule paie est défini et non vide
   - Génère une erreur si non trouvé ou non défini

5. **Initialisation de la structure de données** :
   - Crée une entrée pour le matricule RH si inexistant
   - Calcule les dates de début et fin du mois de validation :
     * **S21.G00.54.003** : 1er jour du mois (yyyyMMdd)
     * **S21.G00.54.004** : Dernier jour du mois (yyyyMMdd)
   - Associe le matricule paie

6. **Accumulation des montants** :
   - Ajoute le montant à la catégorie de frais concernée
   - Consolidation par `[matricule_rh][S21.G00.54.001][catnumber]`

**Résumé des compteurs :**
- Nombre de matricules RH consolidés
- Nombre de lignes de frais consolidées
- Nombre de matricules RH avec erreur

### Extract_Final

Génère la hashtable FINAL comportant les données pour la génération du fichier CSV.

**Structure des données générées :**
- **Code PAC** : Constante définie dans `[Code PAC][Code PAC]`
- **Mois de paie** : Extraite des données de frais (yyyyMM)
- **Matricule paie** : Transcodé depuis le fichier CSV REQ20
- **S21.G00.54.001** : Code de catégorie de frais (DSN)
- **S21.G00.54.002** : Montant arrondi à 2 décimales
- **S21.G00.54.003** : 1er jour du mois de validation (yyyyMMdd)
- **S21.G00.54.004** : Dernier jour du mois de validation (yyyyMMdd)

**Résultat :**
- Chaque ligne = 1 matricule + 1 catégorie de frais
- Les montants sont arrondis à 2 décimales
- Les données sont indexées par numéro de ligne séquentiel

### FinalToCSV

Exporte la hashtable `$script:FINAL` dans le fichier CSV `[FINAL][FichierCSV]`.

**Actions réalisées :**
- Utilise la fonction `ExportCsv` du module `Csv.ps1`
- Applique le délimiteur défini dans `[FINAL][delimiter]`
- Génère l'en-tête : `Code PAC;Mois de paie;Matricule paie;S21.G00.54.001;S21.G00.54.002;S21.G00.54.003;S21.G00.54.004`
- Format des dates dans le CSV : yyyyMMdd

# Fichiers de logs

Les fichiers de logs sont définis dans la section `[intf]` du fichier `.ini`.

## EDB Frais DSN_One_Shot.log

Contient les logs du dernier traitement.

**Comportement :**
- Supprimé automatiquement au démarrage de chaque nouveau traitement
- Contient toutes les traces d'exécution (LOG, DBG, WARN, ERR)
- Paramètre `[start][logtoscreen]` : contrôle l'affichage dans la console
- Paramètre `[start][debug]` : contrôle l'affichage des messages de debug

## EDB Frais DSN_Cumul.err

Contient le cumul des erreurs constatées dans tous les traitements.

**Comportement :**
- Conservé entre les exécutions (mode cumulatif)
- Contient uniquement les erreurs (ERR)
- Paramètre `[start][warntoerr]` : permet d'inclure aussi les warnings (WARN)

# Exemple de fichier .ini

```ini
# -----------------------------------------------------------------------------------------------------------------------------
#    EDB Frais DSN.ini - Necessite Powershell 7 ou +
#      
# -----------------------------------------------------------------------------------------------------------------------------

# -------------------------------------------------------------------
#     Parametrage du comportement de l'interface EDB Frais DSN.ps1
# -------------------------------------------------------------------

[start]
# Le parametre "logtoscreen" contrôle l'affichage de toutes les infos de log/error/warning dans la console
logtoscreen = yes

# Le parametre "debug" contrôle l'affichage des infos de debug dans la console
debug       = no

# Le parametre "warntoerr" permet d'inclure ou pas les warnings dans le fichier EDB Frais DSN_Cumul.err
warntoerr   = yes

# -------------------------------------------------------------------
#     Chemin des fichiers de LOGS
# -------------------------------------------------------------------
[intf]
name        = Generation fichier CSV des Frais DSN 

# Chemin du fichier log : 
pathfilelog = $rootpath$\logs\EDB Frais DSN_One_Shot.log

# Chemin du fichier logs d'erreur
pathfileerr = $rootpath$\logs\EDB Frais DSN_Cumul.err

# -------------------------------------------------------------------
#     Fichiers de données
# -------------------------------------------------------------------

[XLS_Frais_Policy_FORTIL_GROUP]
FichierXLS = $rootpath$\Fichiers\DRH - Frais Policy_FORTIL GROUP_YTD_*.xls

[XLS_Frais_Policy_FUSION]
FichierXLS = $rootpath$\Fichiers\DRH - Frais Policy_Fusion_Group_YTD_*.xls

[CSV_REQ20]
FichierCSV      = $rootpath$\Fichiers\req20_Matricules_paie-*.csv
HEADERstartline = 4

[FINAL]
FichierCSV   = $rootpath$\Fichiers\final.csv
delimiter    = ;   
# Mois de paie   CURRENT / PREVIOUS / 202302  
Mois de paie = 202502

# -------------------------------------------------------------------
#     Parametrage des catégories
# -------------------------------------------------------------------
[Code PAC]
Code PAC = 900380

[S21.G00.54.001]
Frais_Forfait    = 07
Frais_Transports = 19
Frais_Reels      = 09

[Categories]
Forfait hebergement                 = Frais_Forfait
Forfait kilométriques               = Frais_Forfait
Forfait repas midi                  = Frais_Forfait
Forfait repas soir                  = Frais_Forfait
36. Abonnement Transports en commun = Frais_Transports
55. Event Paris 06/06/2025          = Ignore
Default                             = Frais_Reels

# -------------------------------------------------------------------
#     Parametrage des Emails
# Parametre pour l'envoi de mails (Protocoles possible : Microsoft.Graph ou SMTP)
# Le parametre "emailmode" permet de choisir le mode d'emission d'un mail (GRAPH ou SMTP)
# Envoi de mail si sendemail = "yes", sinon "no"
# -------------------------------------------------------------------
[email]
sendemail    = no
emailmode    = SMTP
expediteur   = adremaildetest@gmail.com
destinataire = adremaildetest@gmail.com
server       = smtp.gmail.com
port         = 587
password     = nazr giyx tylx dadz
UseSSL       = true
```

# Modules PowerShell

Le script charge automatiquement les modules suivants depuis le répertoire `Modules/` :

| Module | Description | Utilisation |
|--------|-------------|-------------|
| **Ini.ps1** | Gestion des fichiers .ini | Chargement et parsing du fichier de configuration |
| **Log.ps1** | Système de logging | Fonctions LOG, DBG, WARN, ERR, QUIT |
| **StrConvert.ps1** | Conversion de chaînes | Manipulation et formatage de données |
| **SendEmail.ps1** | Envoi d'emails | Notification de fin de traitement (optionnel) |
| **XLSX.ps1** | Lecture de fichiers Excel | Fonction `Invoke-ExcelQuery` pour requêter les fichiers .xlsx |
| **Csv.ps1** | Gestion des fichiers CSV | Fonctions `Invoke-CSVQuery` et `ExportCsv` |

**Chargement des modules :**
```powershell
. "$PSScriptRoot\Modules\Ini.ps1"        > $null 
. "$PSScriptRoot\Modules\Log.ps1"        > $null
. "$PSScriptRoot\Modules\StrConvert.ps1" > $null
. "$PSScriptRoot\Modules\SendEmail.ps1"  > $null
. "$PSScriptRoot\Modules\XLSX.ps1"       > $null
. "$PSScriptRoot\Modules\Csv.ps1"        > $null
```

Les modules sont chargés en mode "dot sourcing" (`.`) pour rendre leurs fonctions disponibles dans le scope du script principal.

# Notes importantes

## Récurrence recommandée

Ce script est généralement exécuté :
- **Mensuellement** : une fois par mois après validation des frais
- **Sur déclenchement manuel** : lors de corrections ou modifications

## Fichiers d'entrée obligatoires

Pour exécuter le script, les fichiers suivants doivent être présents :
1. Fichier Excel FORTIL GROUP (chemin défini dans `[XLS_Frais_Policy_FORTIL_GROUP][fichierXLS]`)
2. Fichier Excel FUSION (chemin défini dans `[XLS_Frais_Policy_FUSION][fichierXLS]`)
3. Fichier CSV REQ20 (chemin défini dans `[CSV_REQ20][fichierCSV]`)

**Le script générera une erreur et s'arrêtera si l'un de ces fichiers manque.**

## Gestion des erreurs

Le script enregistre tous les problèmes rencontrés :
- **ERR** : Erreurs graves (matricule non trouvé, fichier manquant, etc.)
- **WRN** : Avertissements (montant nul, catégorie ignorée, etc.)
- **LOG** : Informations standard du traitement
- **DBG** : Messages de debug (si `[start][debug] = yes`)

Les erreurs sont cumulées dans le fichier `EDB Frais DSN_Cumul.err` pour suivi.