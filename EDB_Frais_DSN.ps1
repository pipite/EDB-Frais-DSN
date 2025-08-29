# CegedimToSQL.ps1

# --------------------------------------------------------
#               Traitement des LOGS
# --------------------------------------------------------
function Level {
    Param ( [string]$func )
    
    $n = (Get-PSCallStack).Count-3
    $s = ' ' * $n +  $func
    $s = "{0,-36}" -f $s
    return $s
} # code   param -func                                                        Return : [string]    -> Formate la chaine (func)
function Var_Add {
	param ( [string]$file, [string]$value)

	if ( $file -eq $script:cfg["intf"]["pathfilelog"] )     { $script:pathfilelog     += $value }
	if ( $file -eq $script:cfg["intf"]["pathfileerr"] )     { $script:pathfileerr     += $value }
	if ( $file -eq $script:cfg["intf"]["pathfilemod"] )     { $script:pathfilemod     += $value }
}
function Save_logs {
	$script:pathfilelog | Add-Content -Path $script:cfg["intf"]["pathfilelog"]
	$script:pathfileerr | Add-Content -Path $script:cfg["intf"]["pathfileerr"]
	$script:pathfilemod | Add-Content -Path $script:cfg["intf"]["pathfilemod"]
}
function OUT {
	Param ( [string]$trigramme, [string]$func, [string]$msg, [string]$color="White",[bool]$CRLF=$false, [bool]$EMAIL=$false, [bool]$NOSCREEN=$false, [switch]$DBG, [switch]$LOG, [switch]$MOD, [switch]$DLT, [switch]$INA, [switch]$ERR, [switch]$ADDERR, [switch]$ADDWRN)

	# Chaine a afficher
	$f = Level $func
	$Stamp = (Get-Date).toString("yyyy-MM-dd HH:mm:ss")
	$str = "$trigramme : $f : $msg"
	$stampstr = "$Stamp : $str"

	# Affichage a l'ecran

	if ( $script:cfg["start"]["logtoscreen"] -eq "yes" -and -not $NOSCREEN ) { 
		if ( $CRLF ) { Write-Host "" }
		try {
			Write-Host $str -ForegroundColor $color 
		} catch {
			Write-Host $str -ForegroundColor Green
		}
	}

	# Ajout dans les fichiers de logs
	if ( $CRLF ) {
		if ( $DBG -and $script:cfg["start"]["debug"] -eq "yes" ) { Var_Add $($script:cfg["intf"]["pathfilelog"]) -value $stampstr }
		if ( $LOG ) { Var_Add $($script:cfg["intf"]["pathfilelog"])     -value "" }
		if ( $ERR ) { Var_Add $($script:cfg["intf"]["pathfileerr"])     -value "" }

	}
	if ( $DBG -and $script:cfg["start"]["debug"] -eq "yes" ) { Var_Add $($script:cfg["intf"]["pathfilelog"]) -value $stampstr }
	if ( $LOG ) { Var_Add $($script:cfg["intf"]["pathfilelog"])     -value $stampstr }
	if ( $ERR ) { Var_Add $($script:cfg["intf"]["pathfileerr"])     -value $stampstr }
	
	if ( $ADDERR ) { $script:ERREUR  += 1 }
	if ( $ADDWRN ) { $script:WARNING += 1 }

	# Ajoute a Email
	if ( $EMAIL ) { 
		if ( $CRLF ) { $script:emailtxt.Add("") }
		$script:emailtxt.Add($stampstr)	
	}
}
function DBG {
	Param ( [string]$func, [string]$msg, [switch]$CRLF )
	if ( $script:cfg["start"]["debug"] -eq "yes" ) { 
		OUT "DBG" $func $msg "Gray" -DBG -LOG -CRLF $CRLF
	}
} # code   param -func, -msg                                                  Return : N/A         -> Ecrit DBG (func) (msg) dans LOG si [start][debugtolog], et a l'ecran [start][debug] = yes
function LOG {
	Param ( [string]$func, [string]$msg, [string]$color = "Cyan", [switch]$CRLF, [switch]$EMAIL)
	OUT "LOG" $func $msg $color -LOG -CRLF $CRLF -EMAIL $EMAIL
} # code   param -func, -msg, -color, -CRLF, -EMAIL                           Return : N/A         -> Ecrit LOG (func) (msg) dans LOG, et a l'ecran [start][debug] = yes couleur (color)
function ERR {
	Param ( [string]$func, [string]$msg, [switch]$CRLF )

	if ( $script:ERREUR -eq 0 ) {
		OUT "ERR" $func $entete "Red" -ERR -CRLF $CRLF -NOSCREEN $true -EMAIL $true
	}
	OUT "ERR" $func $msg "Red" -ERR -LOG -CRLF $CRLF -ADDERR -EMAIL $true
} # code   param -func, -msg, -CRLF                                           Return : N/A         -> Ecrit ERR (func) (msg) dans ERR, et a l'ecran [start][debug] = yes
function WRN {
	Param ( [string]$func, [string]$msg, [switch]$CRLF )

	if ( $script:cfg["start"]["warntoerr"] -eq "yes" ) {
		OUT "WRN" $func $msg "Magenta" -LOG -CRLF $CRLF -ERR -EMAIL $true -ADDWRN
	} else {
		OUT "WRN" $func $msg "Magenta" -LOG -CRLF $CRLF -EMAIL $true -ADDWRN
	}
} # code   param -func, -msg, -CRLF                                           Return : N/A         -> Ecrit WRN (func) (msg) dans et sort du script
function QUIT {
    Param ( [string]$func, [string]$msg )

    $s = "Duree d'execution : {0:N1} secondes" -f $script:start.Elapsed.TotalSeconds

    if ( $script:ERREUR -eq 0 ) { $c = "Green" } else { $c = "Red" }
	LOG "QUIT" "$($script:ERREUR) erreur, $($script:WARNING) warning, $s" $c -EMAIL
    #$script:emailtxt.Add("$Stamp : QUIT : $f : $msg")
    if ( $script:ERREUR -ne 0 ) {
		OUT "END" $func $msg $c -LOG

        # Contexte
        Get-PSCallStack | Where-Object { $_.Command -and $_.Location } | ForEach-Object {
            if ($_.Command -ne "QUIT") { 
				OUT "END" "$($_.Command)" "$($_.Location)" "Gray" -DBG -LOG -EMAIL $true
            } 
        }
    }

	# Sujet de l'email
	if ( -not $script:MailErr ) {
		$subject = (
			"$($script:cfg['email']['Subject']) : " +
			"[$($script:ERREUR) Erreurs], " +
			"[$($script:WARNING) Warnings]"
		)
		SendEmail $subject $script:emailtxt 
	}

	Save_logs
    exit 0
} # code   param -func, -msg                                                  Return : N/A         -> Ecrit ERR (func) (msg) dans et sort du script 
function QUITEX {
    Param ( [string]$func, [string]$msg, [switch]$ADDERR )

    $lines = $msg -split "`n"
    foreach ($line in $lines) { 
        if ($ADDERR) { 
            ERR "$func" "$line" }
        else { LOG "$func" "$line" }
    }
    QUIT "$func" "Script interrompu."
} # code   param -func -msg, -ADDERR                                          Return : N/A         -> Ecrit ERR (func) (msg) dans et sort du script
function MOD {
	Param ( [string]$func, [string]$msg, [switch]$CRLF )

	if ( $script:cfg["start"]["ApplyUpdate"] -eq "no" ) { $mod = "SIM" } else { $mod = "MOD" }
	OUT $mod $func $msg "Yellow" -LOG -MOD -CRLF $CRLF -EMAIL $true
} # code   param -func, -msg, -CRLF                                           Return : N/A         -> Ecrit MOD (func) (msg) dans LOG, et a l'ecran [start][debug] = yes, saut de ligne si switch -CRLF

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
function GetFilePath {
	param ( [string]$pattern, [switch]$Needed )

	# Remplacement de la chaîne $rootpath$ par le contenu de $script:cfg["intf"]["rootpath"]
	if ($pattern -match '\$rootpath\$') {
		$pattern = $pattern -replace '\$rootpath\$', $script:cfg["intf"]["rootpath"]
	}

	$folder = Split-Path $pattern -Parent
	$filter = Split-Path $pattern -Leaf

	# Créer le répertoire s'il n'existe pas
	if (-not (Test-Path -Path $folder -PathType Container)) {
		try {
			New-Item -Path $folder -ItemType Directory -Force | Out-Null
			DBG "GetFilePath" "Répertoire créé : $folder"
		}
		catch {
			QUITEX "GetFilePath" "Impossible de créer le répertoire '$folder' : $($_.Exception.Message)"
		}
	}

	$files = Get-ChildItem -Path $folder -Filter $filter -File

	if ($files.Count -eq 1) {
		$filepath = $files[0].FullName
	} elseif ($files.Count -eq 0) {
		if ($Needed) {
			QUITEX "GetFilePath" "Aucun fichier ne correspond au filtre '$filter' dans '$folder'" -ADDERR
		} else {
			WRN "GetFilePath" "Aucun fichier ne correspond au filtre '$filter' dans '$folder'"
			$filepath = $pattern
		}
	} else {
		QUITEX "GetFilePath" "Plusieurs fichiers correspondent au filtre '$filter' dans '$folder'" -ADDERR
	}
	return $filepath
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

# Initialisation Culture pour encodage UTF8 et separator numerique "."
chcp 65001 > $null # Encodage avec accent
# Cloner la culture actuelle
$culture = [System.Globalization.CultureInfo]::CurrentCulture.Clone()
# Modifier uniquement le séparateur décimal (de ',' à '.')
$culture.NumberFormat.NumberDecimalSeparator = '.'
# Appliquer cette culture modifiée à la session en cours
[System.Threading.Thread]::CurrentThread.CurrentCulture = $culture

$script:cfgFile = "$PSScriptRoot\EDB_Frais_DSN.ini"
. "$PSScriptRoot\Modules\Ini.ps1" > $null 

LoadIni

Add-Type -AssemblyName System.Web

# Chargement des modules
. "$PSScriptRoot\Modules\SendEmail.ps1"  > $null 
. "$PSScriptRoot\Modules\StrConvert.ps1" > $null 
. "$PSScriptRoot\Modules\XLSX.ps1"       > $null 
. "$PSScriptRoot\Modules\Csv.ps1"        > $null 

Query_XLSX_EDBDSN
Query_CSV_REQ20
Transcode_Matricule
Extract_Final
FinalToCSV

QUIT "MAIN" "Process terminé"


