# Vue d'ensemble

**EDB_Frais_DSN.ps1** est un script PowerShell de traitement des données RH pour la gestion des frais EDB

Le script lit les données de diverses sources, et en extrait les frais d'une période.

* Source de données provenant de plusieurs sources au format **CSV, XLSX.**
* Génère le fichier final.csv

## Source des données  

|                   Paramètres .ini                   |                    Fichiers                     |
| --------------------------------------------------- | ----------------------------------------------- |
| [**XLSX_EDBDSN**][FichierXLSX]                      | EDB Frais DSN.xlsx                              |
| [**CSV_REQ20**][fichierXLS]                         | req20_Matricules_paie-xx-xx-xxxx.csv            |


## Fichier CSV final

|                   Paramètres .ini                   |            Fichier            |
| --------------------------------------------------- | ----------------------------- |
| [FINAL][FichierCSV]                                 | final.csv                     |


## Principe du traitement

A partir des fichiers sources, traiter les datas et **générer une hashtable ayant la même structure que le fichier [FINAL][FichierCSV] à générer**.

* Charge les données depuis les fichiers sources **CSV** et **XLSX.**
* Transcode les matricules RH en matricules paie
* Restreint le traitement sur la période défini dans le parametre [FINAL][Mois de paie] du fichier .ini
    * CURRENT : traite la période du mois de la date courrante
    * PREVIOUS : traite la période du mois précédant la date courrante
	* ALL : traite toutes les périodes
    * xxxxxx : exemple 202302 ( ne traite que la paie de fevrier 2023, correspondant aux données du 01/01/2023 au 31/01/2023 )
* Génère la hashtable FINAL
* Génère le fichier final.csv

**Nota important** : *Ce script ne gère que les matricules RH qui ont une correspondance en matricule paie.*

# Traitements

* LoadIni
* Query_XLSX_EDBDSN
* Query_CSV_REQ20
* Transcode_Matricule
* Extract_Final
* FinalToCSV

### Query_XLSX_EDBDSN

Charge en memoire le contenu du fichier [**XLSX_EDBDSN**][FichierXLSX] et le convertit en hash table.

### Query_CSV_REQ20

Charge en memoire le contenu du fichier [**CSV_REQ20**][fichierXLS]  et le convertit en hash table.

### Transcode_Matricule

Etabli la correspondance entre matricules RH et matricules paie.

### Extract_Final

Génère la hash table FINAL comportant les données nécessaires pour la génération du fichier [FINAL][FichierCSV] 

* Code PAC : Constante défini dans le parametre [XLSX_EDBDSN][Code PAC] du fichier .ini
* Mois de paie : Extrait du fichier [**XLSX_EDBDSN**][FichierXLSX], champ "Période"
* Matricule paie : Transcodé depuis le fichier [**CSV_REQ20**][fichierXLS]
* S21.G00.54.001 : Constante défini dans le parametre [XLSX_EDBDSN][S21.G00.54.001] du fichier .ini
* S21.G00.54.002 : Extrait du fichier [**XLSX_EDBDSN**][FichierXLSX], champ "Solde Tenue de Compte"
* S21.G00.54.003 Premier jour du mois précédant le Mois de paie (calculé)
* S21.G00.54.004 Dernier jour du mois précédant le Mois de paie (calculé)

Ne prend en compte que la période de paie correspondant au parametre [FINAL][Mois de paie] du fichier .ini

**Nota** : *Il est possible de traiter l'ensemble des périodes en indiquant **ALL** dans le parametre [FINAL][Mois de paie] du fichier .ini*

### FinalToCSV

Exporte la hashtable FINAL dans le fichier [FINAL][FichierCSV] 

# Fichiers de LOGS

## EDB Frais DSN_One_Shot.log

Contient les logs du dernier traitement.

## EDB Frais DSN_Cumul.err

Contient le cumul des erreurs constatées dans tous les traitements

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
FichierXLSX    = $rootpath$\Fichiers\EDB Frais DSN.xlsx
SheetSage      = source-sage
Code PAC       = 900380
S21.G00.54.001 = 07

[CSV_REQ20]
FichierCSV      = $rootpath$\Fichiers\req20_Matricules_paie-28-06-2025.csv
HEADERstartline = 4

[FINAL]
FichierCSV   = $rootpath$\Fichiers\final.csv
delimiter    = ;   
# Mois de paie   CURRENT / PREVIOUS / ALL / 202302  
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
