# CegedimToSQL.ps1

# --------------------------------------------------------
#               Chargement fichier .ini
# --------------------------------------------------------

function LoadIni {
	# initialisation variables liste des logs
	$script:pathfilelog = @()
	$script:pathfileerr = @()
	$script:pathfileina = @()
	$script:pathfiledlt = @()
	$script:pathfilemod = @()
	
	# sections de base du fichier .ini
	$script:cfg = @{
        "start"                   = @{}
        "intf"                    = @{}
        "email"                   = @{}
    }
    # Recuperation des parametres passes au script 
    $script:execok  = $false

    if (-not(Test-Path $($script:cfgFile) -PathType Leaf)) { Write-Host "Fichier de parametrage $script:cfgFile innexistant"; exit 1 }
    Write-Host "Fichier de parametrage $script:cfgFile"

    # Initialisation des sections parametres.
    $script:start    = [System.Diagnostics.Stopwatch]::startNew()
    $script:MailErr  = $false
    $script:WARNING  = 0
    $script:ERREUR   = 0
	
	$script:emailtxt = New-Object 'System.Collections.Generic.List[string]'

	$script:cfg = Add-IniFiles $script:cfg $script:cfgFile

	# Recherche des chemins de tous les fichiers et verification de leur existence
	if (-not ($script:cfg["intf"].ContainsKey("rootpath")) ) {
		$script:cfg["intf"]["rootpath"] = $PSScriptRoot
	}
	$script:cfg["intf"]["pathfilelog"] 	= GetFilePath $script:cfg["intf"]["pathfilelog"]
	$script:cfg["intf"]["pathfileerr"]	= GetFilePath $script:cfg["intf"]["pathfileerr"]

	# Suppression des fichiers One_Shot
	if ((Test-Path $($script:cfg["intf"]["pathfilelog"]) -PathType Leaf)) { Remove-Item -Path $script:cfg["intf"]["pathfilelog"]}    

	# Création des fichiers innexistants
	$null = New-Item -type file $($script:cfg["intf"]["pathfilelog"]) -Force;
	if (-not(Test-Path $($script:cfg["intf"]["pathfileerr"]) -PathType Leaf)) { $null = New-Item -type file $($script:cfg["intf"]["pathfileerr"]) -Force; }

	
	$script:cfg["XLS_Frais_Policy_FORTIL_GROUP"]["fichierXLS"] = GetFilePath $script:cfg["XLS_Frais_Policy_FORTIL_GROUP"]["fichierXLS"] -Needed
	$script:cfg["XLS_Frais_Policy_FUSION"]["fichierXLS"]       = GetFilePath $script:cfg["XLS_Frais_Policy_FUSION"]["fichierXLS"] -Needed
	$script:cfg["CSV_REQ20"]["fichierCSV"]                     = GetFilePath $script:cfg["CSV_REQ20"]["fichierCSV"] -Needed
	$script:cfg["FINAL"]["FichierCSV"]                         = GetFilePath $script:cfg["FINAL"]["FichierCSV"]

	if ( $script:cfg['FINAL']['Mois de paie'] -eq "CURRENT" ) {
		$script:cfg['FINAL']['Mois de paie'] = (Get-Date).ToString("yyyyMM")
	}
	if ( $script:cfg['FINAL']['Mois de paie'] -eq "PREVIOUS" ) {
		$script:cfg['FINAL']['Mois de paie'] = (Get-Date).AddMonths(-1).ToString("yyyyMM")
	}
}
function Compute_CategorieFrais {
	$script:Categorie = @{}
	foreach ($key in $script:cfg['Categories'].Keys) {
		$cat = $script:cfg['Categories'][$key]
		if ( $cat -eq "Ignore") {
			$val = "Ignore"
		} else {
			foreach ($s21 in $script:cfg['S21.G00.54.001'].Keys) {
				if ($cat -eq $s21) {
					$val = $script:cfg['S21.G00.54.001'][$s21]
					break
				}
			}
		}
		$script:Categorie[$key] = $val
		DBG "Compute_CategorieFrais" "Catégorie : $key >> $val"	
	}
}
function Query_CSV_REQ20 {
	$file            = $script:cfg["CSV_REQ20"]["FichierCSV"]
	$headerstartline = $script:cfg["CSV_REQ20"]["HEADERstartline"] - 1
	$datecol 		 = @("D Fin contrat")

	LOG "Query_CSV_REQ20" "Chargement du fichier $file" -CRLF
	$script:REQ20 = Invoke-CSVQuery -csvfile $file -key "Matricule" -separator "," -row $headerstartline -frmtdateOUT $script:cfg["intf"]["DateFormat"] -datecol $datecol
}
function Get_CategorieNumber {
	param ( [string]$categorie )

	if ( $script:Categorie.ContainsKey($categorie) ) {
		return $script:Categorie[$categorie]
	}
	return $script:Categorie["Default"]
}
function Query_XLS_Frais_Policy {
	$script:FRAIS = @{}

	$query = "SELECT [Matricule],[Catégorie (Description)],[Date de validation],[Montant] FROM [Sheet0$] WHERE [Statut] = 'PMNT_NOT_PAID' AND [statusId2] = 'EXPENSE_ACCEPTED' ORDER BY [Matricule]"
	$file  = $script:cfg["XLS_Frais_Policy_FORTIL_GROUP"]["fichierXLS"] 
	$sheet = $script:cfg["XLS_Frais_Policy_FORTIL_GROUP"]["Sheet"]
	LOG "Query_XLS_Frais_Policy" "Chargement du fichier $file - $sheet"
	$FORTILGROUP = Invoke-ExcelQuery -filePath $file -sqlQuery $query -key "_index" -ConvertToHashtable
	Consolide_Frais -Data $FORTILGROUP

	$file  = $script:cfg["XLS_Frais_Policy_FUSION"]["fichierXLS"] 
	$sheet = $script:cfg["XLS_Frais_Policy_FUSION"]["Sheet"]
	LOG "Query_XLS_Frais_Policy" "Chargement du fichier $file - $sheet" -CRLF
	$FUSION = Invoke-ExcelQuery -filePath $file -sqlQuery $query -key "_index" -ConvertToHashtable
	Consolide_Frais -Data $FUSION
}
function Consolide_Frais {
    param ( [hashtable]$data )

	LOG "Consolide_Frais" "Consolidation des frais par type de frais et par matricule"
	$cptuser = 0
	$cptlignes = 0
	$cptundef = 0
	$lastmatricule_rh = ""

	foreach ($key in $data.Keys) {
		# passer au suivant si pas bon mois de paie
		$dateval = [datetime]::ParseExact($data[$key]['Date de validation'], "yyyy-MM-dd HH:mm:ss.0", $null)
		$datemonth = $dateval.ToString("yyyyMM")
		if ( $datemonth -ne $script:cfg['FINAL']['Mois de paie'] ) { continue }
		
		# passer au suivant si Categorie = Ignore
		$catnumber = Get_CategorieNumber $data[$key]['Catégorie (Description)']
		if ( $catnumber -eq "Ignore" ) { continue }

		# passer au suivant si montant = 0
		$montant = [decimal]$data[$key]['Montant']
		if ( $montant -eq 0 ) { continue }

		# passer au suivant si [Matricule paie] non défini dans REQ20
		[string]$matricule_rh = $data[$key]['Matricule']
		$zeromatricule_rh = $matricule_rh.PadLeft(8, '0')
		if ( $script:REQ20.ContainsKey($zeromatricule_rh) ) {
			$matricule_paie = $script:REQ20[$zeromatricule_rh]['matricule paie']
			if ([string]::IsNullOrEmpty($matricule_paie)) {
				if ( $lastmatricule_rh -ne $zeromatricule_rh ) {
					ERR "Transcode_Matricule" "[Matricule paie] non défini dans REQ20 pour [matricule rh] : $zeromatricule_rh"
					$cptundef++
				}
				$matricule_paie = "UNKNOWN"
				$lastmatricule_rh = $zeromatricule_rh
				continue
			}
		} else {
				if ( $lastmatricule_rh -ne $zeromatricule_rh ) {
					ERR "Transcode_Matricule" "[Matricule paie] non défini dans REQ20 pour [matricule rh] : $zeromatricule_rh"
					$cptundef++
				}
			$matricule_paie = "UNKNOWN"
			$lastmatricule_rh = $zeromatricule_rh
			continue
		}

		# Initialisation de la structure de données [Matricule_rh] si matricule inexistant
		if ( -not ($script:FRAIS.ContainsKey($matricule_rh)) ) { 
			$cptuser++
			$moisdebut = (Get-Date -Year $dateval.Year -Month $dateval.Month -Day 1).ToString('yyyyMMdd')
			$moisfin   = (Get-Date -Year $dateval.Year -Month $dateval.Month -Day 1).AddMonths(1).AddDays(-1).ToString('yyyyMMdd')
			$script:FRAIS[$matricule_rh] = @{} 
			$script:FRAIS[$matricule_rh]['Code PAC']       = $script:cfg['Code PAC']['Code PAC']
			$script:FRAIS[$matricule_rh]['Mois de paie']   = $datemonth
			$script:FRAIS[$matricule_rh]['Matricule paie'] = $matricule_paie
			$script:FRAIS[$matricule_rh]['S21.G00.54.001'] = @{}
			$script:FRAIS[$matricule_rh]['S21.G00.54.003'] = $moisdebut
			$script:FRAIS[$matricule_rh]['S21.G00.54.004'] = $moisfin
		}

		# Initialisation de la structure de données [Matricule_rh][catnumber] si catnumber inexistant (S21.G00.54.001)
		if ( -not ($script:FRAIS[$matricule_rh]['S21.G00.54.001'].ContainsKey($catnumber)) ) { 
			$script:FRAIS[$matricule_rh]['S21.G00.54.001'][$catnumber] = 0
			$cptlignes++
		}
		$script:FRAIS[$matricule_rh]['S21.G00.54.001'][$catnumber] += $montant  # S21.G00.54.002
	}
	LOG "Consolide_Frais" "Total [Matricule] : $cptuser - Total lignes de frais consolidées [Matricule][Categories] : $cptlignes"
	if ( $cptundef -gt 0 ) {
		WRN "Consolide_Frais" "Total [Matricule paie] non définis dans REQ20 : $cptundef"
	}
}

function Extract_Final {
	LOG "Extract_Final" "Create CSV final file : Mois de paie : $($script:cfg['FINAL']['Mois de paie'])" -CRLF
	$cpt = 0
	$script:FINAL = @{}

    foreach ($matricule_rh in $script:FRAIS.Keys) {
		foreach ($catnumber in $script:FRAIS[$matricule_rh]['S21.G00.54.001'].Keys) {
			$script:FINAL[$cpt] = @{}
			$script:FINAL[$cpt]['Code PAC']       = $script:FRAIS[$matricule_rh]['Code PAC']
			$script:FINAL[$cpt]['Mois de paie']   = $script:FRAIS[$matricule_rh]['Mois de paie']
			$script:FINAL[$cpt]['Matricule paie'] = $script:FRAIS[$matricule_rh]['Matricule paie']
			$script:FINAL[$cpt]['S21.G00.54.001'] = $catnumber
			$script:FINAL[$cpt]['S21.G00.54.002'] = [math]::Round($script:FRAIS[$matricule_rh]['S21.G00.54.001'][$catnumber], 2)
			$script:FINAL[$cpt]['S21.G00.54.003'] = $script:FRAIS[$matricule_rh]['S21.G00.54.003']
			$script:FINAL[$cpt]['S21.G00.54.004'] = $script:FRAIS[$matricule_rh]['S21.G00.54.004']
			$cpt++
		}
	}
	LOG "Extract_Final" "$cpt lignes traités."
}

function FinalToCSV {
	$filepath = $script:cfg["FINAL"]["FichierCSV"]
	$delimiter = $script:cfg["FINAL"]["delimiter"]
	$header = "Code PAC;Mois de paie;Matricule paie;S21.G00.54.001;S21.G00.54.002;S21.G00.54.003;S21.G00.54.004"
	$header = $header.Replace(";", $delimiter)
	ExportCsv $filepath $script:FINAL $header $delimiter
}

# --------------------------------------------------------
#               Main
# --------------------------------------------------------
	# Chargement des modules
	$pathmodule = "$PSScriptRoot\Modules"

	if (Test-Path "$pathmodule\Ini.ps1" -PathType Leaf) {
		. "$pathmodule\Ini.ps1"                        > $null 
		. (GetPathScript "$pathmodule\Log.ps1")        > $null
		. (GetPathScript "$pathmodule\StrConvert.ps1") > $null
		. (GetPathScript "$pathmodule\SendEmail.ps1")  > $null
		. (GetPathScript "$pathmodule\XLSX.ps1")       > $null
		. (GetPathScript "$pathmodule\Csv.ps1")        > $null
	} else {
		Write-Host "Fichier manquant : $pathmodule\Ini.ps1" -ForegroundColor Red
		exit (1)
	}

	# Recuperation des parametres passes au script dans $script:cfg
	$script:cfgFile = "$PSScriptRoot\EDB_Frais_DSN.ini"
	LoadIni

	# Parametrage console en UFT8 (chcp 65001 ou 850) pour carractères accentués
	SetConsoleToUFT8

	Compute_CategorieFrais
	Query_CSV_REQ20
	Query_XLS_Frais_Policy
	Extract_Final
	FinalToCSV

	QUIT "MAIN" "Process terminé"


