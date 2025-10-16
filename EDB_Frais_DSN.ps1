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

	
	$script:cfg["XLSX_EDBDSN"]["fichierXLSX"] = GetFilePath $script:cfg["XLSX_EDBDSN"]["fichierXLSX"] -Needed
	$script:cfg["CSV_REQ20"]["fichierCSV"]    = GetFilePath $script:cfg["CSV_REQ20"]["fichierCSV"] -Needed
	$script:cfg["FINAL"]["FichierCSV"]        = GetFilePath $script:cfg["FINAL"]["FichierCSV"]
}
function Query_XLSX_EDBDSN {
	$file         = $script:cfg["XLSX_EDBDSN"]["fichierXLSX"] 
	$sheetSage    = $script:cfg["XLSX_EDBDSN"]["SheetSage"]
	$sheetTransco = $script:cfg["XLSX_EDBDSN"]["sheetTransco"]

	LOG "Query_XLSX_EDBDSN" "Chargement du fichier $file - $sheetSage"
	$result = Invoke-ExcelQuery -filePath $file -sqlQuery "SELECT * FROM [$sheetSage$]" -key "_index"
	$table  = $result.Table
	$cpt = 0
	$script:EDBDSN = @{}
    foreach ($row in $table.Rows) {
		$script:EDBDSN[$cpt] = @{}
		foreach ($col in $table.Columns) {
			$cap = $col.Caption
			$script:EDBDSN[$cpt][$cap] = $row[$col]
		}
		[string]$s = $row['Tiers - Code']
		if ($s.Length -ge 9) {
			$s = $s.Substring(4,4)
			$s = $s.PadLeft(8, '0')
			$script:EDBDSN[$cpt]['matricule rh'] = $s
		} else {
			$script:EDBDSN[$cpt]['matricule rh'] = ""
		}
		$cpt++
	}
}

function Query_CSV_REQ20 {
	$file            = $script:cfg["CSV_REQ20"]["FichierCSV"]
	$headerstartline = $script:cfg["CSV_REQ20"]["HEADERstartline"] - 1
	$datecol 		 = @("D Fin contrat")

	LOG "Query_CSV_REQ20" "Chargement du fichier $file"
	$script:REQ20   = @{}
	$script:REQ20   = Invoke-CSVQuery -csvfile $file -key "Matricule" -separator "," -row $headerstartline -frmtdateOUT $script:cfg["intf"]["DateFormat"] -datecol $datecol
}

function Transcode_Matricule {
	LOG "Transcode_Matricule" "Transcode Matricule RH >> Matricule Paie"

	$ok = 0
	$unknow = 0
	$undef = 0
	$last = ""
    foreach ($key in $script:EDBDSN.Keys) {
        $matricule_rh = $script:EDBDSN[$key]['matricule rh']
		$nom = $script:EDBDSN[$key]['Tiers - Libellé']
		if ( $script:REQ20.ContainsKey($matricule_rh) ) {
			$matricule_paie = $script:REQ20[$matricule_rh]['matricule paie']
			if ([string]::IsNullOrEmpty($matricule_paie)) {
				if ( $last -ne $matricule_rh ) {
					ERR "Transcode_Matricule" "[Matricule paie] non défini dans REQ20 pour [matricule rh] : $matricule_rh >> $nom"
					$undef++
				}
				$script:EDBDSN[$key]['matricule paie'] = "UNKNOWN"
			} else {
				DBG "Transcode_Matricule" "$matricule_rh >> $matricule_paie >> $nom"
				$script:EDBDSN[$key]['matricule paie'] = $matricule_paie
				$ok++
			}
		} else {
			if ( $last -ne $matricule_rh ) {
				ERR "Transcode_Matricule" "[matricule rh] n'existe pas dans le fichier REQ20 : $matricule_rh >> $nom"
				$unknow++
			}
			$script:EDBDSN[$key]['matricule paie'] = "UNKNOWN"
		}
		$last = $matricule_rh
	}
	LOG "Transcode_Matricule" "matricule : OK $ok >> UNKNOWN $unknow >> UNDEF $undef"
}

function Extract_Final {
	LOG "Extract_Final" "Create CSV final file : Mois de paie : $($script:cfg['FINAL']['Mois de paie'])"
	$cpt = 0
	$script:FINAL = @{}

	if ( $script:cfg['FINAL']['Mois de paie'] -eq "CURRENT" ) {
		$script:cfg['FINAL']['Mois de paie'] = (Get-Date).ToString("yyyyMM")
	}
	if ( $script:cfg['FINAL']['Mois de paie'] -eq "PREVIOUS" ) {
		$script:cfg['FINAL']['Mois de paie'] = (Get-Date).AddMonths(-1).ToString("yyyyMM")
	}
    foreach ($key in $script:EDBDSN.Keys) {
		if ( $script:EDBDSN[$key]['matricule paie'] -ne "UNKNOWN" ) {
			$periode = $script:EDBDSN[$key]['Période']
			if ( ($script:cfg['FINAL']['Mois de paie'] -eq "ALL") -or ($periode -eq $script:cfg['FINAL']['Mois de paie']) ) {
				$periodeDate = [datetime]::ParseExact($periode + "01", "yyyyMMdd", $null)

				$moisPrecedent = $periodeDate.AddMonths(-1)
				$start = $moisPrecedent.ToString("yyyyMM") + "01"
				$end = $moisPrecedent.AddMonths(1).AddDays(-1).ToString("yyyyMMdd")

				$script:FINAL[$key] = @{}
				$script:FINAL[$key]['Code PAC']       = $script:cfg['XLSX_EDBDSN']['Code PAC']
				$script:FINAL[$key]['Mois de paie']   = $periode
				$script:FINAL[$key]['Matricule paie'] = $script:EDBDSN[$key]['matricule paie']
				$script:FINAL[$key]['S21.G00.54.001'] = $script:cfg['XLSX_EDBDSN']['S21.G00.54.001']
				$script:FINAL[$key]['S21.G00.54.002'] = $script:EDBDSN[$key]['Solde Tenue de Compte']
				$script:FINAL[$key]['S21.G00.54.003'] = $start
				$script:FINAL[$key]['S21.G00.54.004'] = $end
				$cpt++
			}
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

$script:cfgFile = "$PSScriptRoot\EDB_Frais_DSN.ini"
# Chargement des modules

. "$PSScriptRoot\Modules\Ini.ps1" > $null 
. "$PSScriptRoot\Modules\Log.ps1" > $null 
. "$PSScriptRoot\Modules\Encode.ps1" > $null 
. "$PSScriptRoot\Modules\SendEmail.ps1"  > $null 
. "$PSScriptRoot\Modules\StrConvert.ps1" > $null 
. "$PSScriptRoot\Modules\XLSX.ps1"       > $null 
. "$PSScriptRoot\Modules\Csv.ps1"        > $null 

LoadIni

SetConsoleToUFT8

Add-Type -AssemblyName System.Web

# Chargement des modules

Query_XLSX_EDBDSN
Query_CSV_REQ20
Transcode_Matricule
Extract_Final
FinalToCSV

QUIT "MAIN" "Process terminé"


