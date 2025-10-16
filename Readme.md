# Vue d'ensemble

**EDB_Frais_DSN.ps1** est un script PowerShell de traitement des données RH pour la gestion des frais EDB.

Le script lit les données de diverses sources, et en extrait les frais d'une période.

* Source de données provenant de plusieurs sources au format **CSV, XLSX**
* Génère le fichier final.csv

## Comment installer ce script

* PowerShell 7 ou supérieur

Recuperer le script sur GitLAB, et déposer les fichiers dans un répertoire du serveur de Script.

### Modules externes

Recuperer les modules nécessaire sur GitLAB, et les déposer dans le répertoire Modules du script.

* Modules PowerShell requis dans le dossier `Modules/`) :
  - **Ini.ps1** : Gestion des fichiers de configuration .ini
  - **Log.ps1** : Gestion des logs et traces
  - **Encode.ps1** : Gestion de l'encodage (UTF-8)
  - **SendEmail.ps1** : Envoi d'emails de notification
  - **StrConvert.ps1** : Conversion et manipulation de chaînes
  - **XLSX.ps1** : Lecture et requêtage de fichiers Excel
  - **Csv.ps1** : Lecture et export de fichiers CSV

Paramétrer le fichier EDB_Frais_DSN.ini

## Structure du script

Le script suit une architecture modulaire avec :
1. **Chargement des modules** : Import des modules PowerShell nécessaires
2. **Initialisation** : Chargement du fichier de configuration `.ini`
3. **Configuration de l'encodage** : Passage de la console en UTF-8
4. **Traitement des données** : Exécution séquentielle des fonctions métier
5. **Finalisation** : Génération du fichier CSV et envoi optionnel d'email

## Source des données  

|                   Paramètres .ini                   |                    Fichiers                     |
| --------------------------------------------------- | ----------------------------------------------- |
| [**XLSX_EDBDSN**][fichierXLSX]                      | EDB Frais DSN.xlsx                              |
| [**CSV_REQ20**][fichierCSV]                         | req20_Matricules_paie-xx-xx-xxxx.csv            |


## Fichier CSV final

|                   Paramètres .ini                   |            Fichier            |
| --------------------------------------------------- | ----------------------------- |
| [FINAL][FichierCSV]                                 | final.csv                     |


## Principe du traitement

À partir des fichiers sources, traiter les données et **générer une hashtable ayant la même structure que le fichier [FINAL][FichierCSV] à générer**.

### Flux de traitement

1. **LoadIni** : Charge le fichier de configuration `EDB_Frais_DSN.ini`
   - Initialise les variables de log et d'erreur
   - Vérifie l'existence des fichiers sources
   - Crée les fichiers de log nécessaires

2. **SetConsoleToUFT8** : Configure l'encodage de la console en UTF-8

3. **Query_XLSX_EDBDSN** : Charge les données depuis le fichier Excel
   - Lit la feuille spécifiée dans `[XLSX_EDBDSN][SheetSage]`
   - Extrait et formate le matricule RH depuis le code Tiers
   - Stocke les données dans la hashtable `$script:EDBDSN`

4. **Query_CSV_REQ20** : Charge les données depuis le fichier CSV
   - Lit le fichier CSV avec gestion des dates
   - Utilise le matricule comme clé d'indexation
   - Stocke les données dans la hashtable `$script:REQ20`

5. **Transcode_Matricule** : Établit la correspondance matricule RH ↔ matricule paie
   - Recherche chaque matricule RH dans le fichier REQ20
   - Ajoute le matricule paie correspondant dans `$script:EDBDSN`
   - Marque comme "UNKNOWN" les matricules non trouvés ou non définis

6. **Extract_Final** : Génère la hashtable finale
   - Filtre selon la période définie dans `[FINAL][Mois de paie]` :
     * **CURRENT** : traite la période du mois de la date courante
     * **PREVIOUS** : traite la période du mois précédant la date courante
     * **ALL** : traite toutes les périodes
     * **yyyyMM** : exemple 202302 (ne traite que la paie de février 2023, correspondant aux données du 01/01/2023 au 31/01/2023)
   - Calcule les dates de début et fin de période (mois précédent le mois de paie)
   - Exclut les matricules "UNKNOWN"
   - Stocke les données dans la hashtable `$script:FINAL`

7. **FinalToCSV** : Exporte la hashtable finale vers le fichier CSV
   - Utilise le délimiteur défini dans `[FINAL][delimiter]`
   - Génère le fichier `[FINAL][FichierCSV]`

8. **QUIT** : Finalise le traitement
   - Affiche le résumé des erreurs et warnings
   - Envoie un email de notification si configuré

**Nota important** : *Ce script ne gère que les matricules RH qui ont une correspondance en matricule paie.*

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
1. Définir les chemins des fichiers sources dans les sections `[XLSX_EDBDSN]` et `[CSV_REQ20]`
2. Définir le chemin du fichier de sortie dans la section `[FINAL]`
3. Choisir la période à traiter avec le paramètre `[FINAL][Mois de paie]`
4. Configurer les options de log dans la section `[start]`
5. Configurer l'envoi d'email dans la section `[email]` (optionnel)

### Variables d'environnement

Le script utilise `$rootpath$` comme variable de substitution dans le fichier `.ini` :
- Par défaut : répertoire du script (`$PSScriptRoot`)
- Peut être redéfini dans `[intf][rootpath]`

# Détail des fonctions principales

### LoadIni

Charge et initialise le fichier de configuration `EDB_Frais_DSN.ini`.

**Actions réalisées :**
- Initialise les variables de log (`$script:pathfilelog`, `$script:pathfileerr`, etc.)
- Vérifie l'existence du fichier `.ini`
- Charge les paramètres via la fonction `Add-IniFiles`
- Résout les chemins des fichiers avec `GetFilePath`
- Supprime le fichier log One_Shot s'il existe (nouveau traitement)
- Crée les fichiers de log nécessaires
- Valide l'existence des fichiers sources obligatoires

### Query_XLSX_EDBDSN

Charge en mémoire le contenu du fichier Excel `[XLSX_EDBDSN][fichierXLSX]` et le convertit en hashtable.

**Traitement spécifique :**
- Utilise `Invoke-ExcelQuery` pour lire la feuille `[XLSX_EDBDSN][SheetSage]`
- Extrait le matricule RH depuis le champ "Tiers - Code" (caractères 5 à 8, complétés à gauche par des zéros sur 8 positions)
- Indexe les données par un compteur séquentiel
- Stocke le résultat dans `$script:EDBDSN`

### Query_CSV_REQ20

Charge en mémoire le contenu du fichier CSV `[CSV_REQ20][fichierCSV]` et le convertit en hashtable.

**Traitement spécifique :**
- Utilise `Invoke-CSVQuery` avec le séparateur `,`
- Démarre la lecture à la ligne définie par `[CSV_REQ20][HEADERstartline]`
- Indexe les données par le champ "Matricule"
- Gère le formatage des dates pour la colonne "D Fin contrat"
- Stocke le résultat dans `$script:REQ20`

### Transcode_Matricule

Établit la correspondance entre matricules RH et matricules paie.

**Logique de traitement :**
- Parcourt chaque entrée de `$script:EDBDSN`
- Recherche le matricule RH dans `$script:REQ20`
- Si trouvé et défini : ajoute le matricule paie dans `$script:EDBDSN`
- Si non trouvé ou non défini : marque comme "UNKNOWN" et génère une erreur
- Affiche un résumé : OK / UNKNOWN / UNDEF

### Extract_Final

Génère la hashtable FINAL comportant les données nécessaires pour la génération du fichier CSV final.

**Structure des données générées :**
- **Code PAC** : Constante définie dans `[XLSX_EDBDSN][Code PAC]`
- **Mois de paie** : Extrait du champ "Période" du fichier Excel
- **Matricule paie** : Transcodé depuis le fichier CSV REQ20
- **S21.G00.54.001** : Constante définie dans `[XLSX_EDBDSN][S21.G00.54.001]`
- **S21.G00.54.002** : Extrait du champ "Solde Tenue de Compte" du fichier Excel
- **S21.G00.54.003** : Premier jour du mois précédant le mois de paie (calculé)
- **S21.G00.54.004** : Dernier jour du mois précédant le mois de paie (calculé)

**Filtrage :**
- Exclut les matricules marqués "UNKNOWN"
- Ne prend en compte que la période correspondant au paramètre `[FINAL][Mois de paie]`

**Nota** : *Il est possible de traiter l'ensemble des périodes en indiquant **ALL** dans le paramètre `[FINAL][Mois de paie]`*

### FinalToCSV

Exporte la hashtable `$script:FINAL` dans le fichier CSV `[FINAL][FichierCSV]`.

**Actions réalisées :**
- Utilise la fonction `ExportCsv` du module `Csv.ps1`
- Applique le délimiteur défini dans `[FINAL][delimiter]`
- Génère l'en-tête : `Code PAC;Mois de paie;Matricule paie;S21.G00.54.001;S21.G00.54.002;S21.G00.54.003;S21.G00.54.004` 

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

# Format de date
DateFormat = yyyy/MM/dd

# -------------------------------------------------------------------
#     Fichiers de données
# -------------------------------------------------------------------
[XLSX_EDBDSN]
fichierXLSX    = $rootpath$\Fichiers\EDB Frais DSN.xlsx
SheetSage      = source-sage
sheetTransco   = transco
Code PAC       = 900380
S21.G00.54.001 = 07

[CSV_REQ20]
fichierCSV      = $rootpath$\Fichiers\req20_Matricules_paie-28-06-2025.csv
HEADERstartline = 4

[FINAL]
FichierCSV   = $rootpath$\Fichiers\final.csv
delimiter    = ;   
# Mois de paie : CURRENT / PREVIOUS / ALL / yyyyMM (ex: 202302)
Mois de paie = 202302

# -------------------------------------------------------------------
#     Parametrage des Emails
# -------------------------------------------------------------------

# Parametre pour l'envoi de mails (Protocoles possible : Microsoft.Graph ou SMTP)
# Le parametre "emailmode" permet de choisir le mode d'emission d'un mail (GRAPH ou SMTP)
# Envoi de mail si sendemail = "yes" / "no"
[email]
sendemail    = no
destinataire = btran56@gmail.com
Subject      = EDB Frais DSN
emailmode    = SMTP
UseSSL       = false

# Login pour SMTP
expediteur   = btran56@gmail.com
server       = smtp.gmail.com
port         = 
password     = 
```

# Modules PowerShell

Le script charge automatiquement les modules suivants depuis le répertoire `Modules/` :

| Module | Description | Utilisation dans le script |
|--------|-------------|----------------------------|
| **Ini.ps1** | Gestion des fichiers .ini | Chargement et parsing du fichier de configuration |
| **Log.ps1** | Système de logging | Fonctions LOG, DBG, WARN, ERR, QUIT |
| **Encode.ps1** | Gestion de l'encodage | Configuration UTF-8 de la console |
| **SendEmail.ps1** | Envoi d'emails | Notification de fin de traitement (optionnel) |
| **StrConvert.ps1** | Conversion de chaînes | Manipulation et formatage de données |
| **XLSX.ps1** | Lecture de fichiers Excel | Fonction `Invoke-ExcelQuery` pour requêter les fichiers .xlsx |
| **Csv.ps1** | Gestion des fichiers CSV | Fonctions `Invoke-CSVQuery` et `ExportCsv` |

**Chargement des modules :**
```powershell
. "$PSScriptRoot\Modules\Ini.ps1" > $null 
. "$PSScriptRoot\Modules\Log.ps1" > $null 
. "$PSScriptRoot\Modules\Encode.ps1" > $null 
. "$PSScriptRoot\Modules\SendEmail.ps1" > $null 
. "$PSScriptRoot\Modules\StrConvert.ps1" > $null 
. "$PSScriptRoot\Modules\XLSX.ps1" > $null 
. "$PSScriptRoot\Modules\Csv.ps1" > $null 
```

Les modules sont chargés en mode "dot sourcing" (`.`) pour rendre leurs fonctions disponibles dans le scope du script principal.

**Note :** La documentation détaillée de chaque module sera fournie séparément.
