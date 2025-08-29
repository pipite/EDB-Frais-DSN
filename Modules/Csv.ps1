


function ImportCSV {
    <#
    param (
        [string]$csvfile,
        [string]$key,
        [string]$separator,
        [int]$row = 0,
        [string]$header = $null,
        [string]$frmtdateOUT = $null,
        [string]$dateLocale = "FR"
    )

    $hashtable = @{}
    $index = 1
    $isAutoIndex = ($key -eq "_index")
    
    # Mode 1 ou 4 : header en première ligne + pas de $row + pas de $header
    if (-not $header -and $row -eq 0) {
        $csv = Import-Csv -Path $csvfile -Delimiter $separator
    }
    else {
        # Lire toutes les lignes du fichier
        $lines = Get-Content -Path $csvfile

        # Mode 2 : header dans le fichier à la ligne $row
        if (-not $header) {
            $headerLine = $lines[$row]
            $headers = ($headerLine -split $separator) | ForEach-Object { $_.Trim('"') }
            $dataLines = $lines[($row + 1)..($lines.Count - 1)]
        }
        # Mode 3 : header fourni manuellement
        else {
            $headers = ($header -split $separator) | ForEach-Object { $_.Trim('"') }
            $dataLines = $lines[$row..($lines.Count - 1)]
        }

        # Convertir les lignes en objets CSV
        $csv = $dataLines | ConvertFrom-Csv -Delimiter $separator -Header $headers
    }

    foreach ($line in $csv) {
        $rowKey = $isAutoIndex ? $index++ : $line.$key
        if ($rowKey) {
            $hashtable[$rowKey] = @{}
            foreach ($col in $line.PSObject.Properties.Name) {
                $value = $line.$col
                # Appliquer la conversion de date si le format de sortie est spécifié
                if ($frmtdateOUT) {
                    $value = ConvertDateToString -value $value -formatOut $frmtdateOUT -locale $dateLocale
                }
                $hashtable[$rowKey][$col] = $value
            }
        }
    }

    return $hashtable
    #>
    return $null
}

function Invoke-CSVQuery {
    <#
    .SYNOPSIS
        Version optimisée de ImportCSV avec gestion ciblée des colonnes de dates
    
    .PARAMETER csvfile
        Chemin vers le fichier CSV
    
    .PARAMETER key
        Nom de la colonne à utiliser comme clé (ou "_index" pour auto-index)
    
    .PARAMETER separator
        Séparateur du fichier CSV
    
    .PARAMETER row
        Numéro de ligne contenant les données (0 = première ligne)
    
    .PARAMETER header
        Header personnalisé (optionnel)
    
    .PARAMETER frmtdateOUT
        Format de sortie pour les dates (ex: "yyyy-MM-dd")
    
    .PARAMETER dateLocale
        Locale pour l'interprétation des dates ("FR" ou "US")
    
    .PARAMETER datecol
        Liste des colonnes contenant des dates à convertir
    
    .EXAMPLE
        ImportCSVOptimized -csvfile "data.csv" -key "ID" -separator ";" -frmtdateOUT "yyyy-MM-dd" -datecol @("Date_Entree", "Date_Sortie")
    #>
    param (
        [string]$csvfile,
        [string]$key,
        [string]$separator,
        [int]$row = 0,
        [string]$header = $null,
        [string]$frmtdateOUT = $null,
        [string]$dateLocale = "FR",
        [string[]]$datecol = @()
    )

    $hashtable = @{}
    $index = 1
    $isAutoIndex = ($key -eq "_index")
    
    # Mode 1 ou 4 : header en première ligne + pas de $row + pas de $header
    if (-not $header -and $row -eq 0) {
        $csv = Import-Csv -Path $csvfile -Delimiter $separator
    }
    else {
        # Lire toutes les lignes du fichier
        $lines = Get-Content -Path $csvfile

        # Mode 2 : header dans le fichier à la ligne $row
        if (-not $header) {
            $headerLine = $lines[$row]
            $headers = ($headerLine -split $separator) | ForEach-Object { $_.Trim('"') }
            $dataLines = $lines[($row + 1)..($lines.Count - 1)]
        }
        # Mode 3 : header fourni manuellement
        else {
            $headers = ($header -split $separator) | ForEach-Object { $_.Trim('"') }
            $dataLines = $lines[$row..($lines.Count - 1)]
        }

        # Convertir les lignes en objets CSV
        $csv = $dataLines | ConvertFrom-Csv -Delimiter $separator -Header $headers
    }

    # Vérifier les colonnes de dates spécifiées
    $validDateColumns = @()
    if ($datecol.Count -gt 0) {
        $availableColumns = $csv[0].PSObject.Properties.Name
        foreach ($colName in $datecol) {
            if ($availableColumns -contains $colName) {
                $validDateColumns += $colName
            } else {
                WRN "ImportCSVOptimized" "Colonne de date '$colName' introuvable dans le fichier CSV"
            }
        }
        if ($validDateColumns.Count -gt 0) {
            DBG "ImportCSVOptimized" "Colonnes de dates à convertir : $($validDateColumns -join ', ')"
        }
    }

    foreach ($line in $csv) {
        $rowKey = $isAutoIndex ? $index++ : $line.$key
        if ($rowKey) {
            $hashtable[$rowKey] = @{}
            foreach ($col in $line.PSObject.Properties.Name) {
                $value = $line.$col
                
                # Appliquer la conversion de date seulement aux colonnes spécifiées
                if ($frmtdateOUT -and $validDateColumns -contains $col) {
                    $value = ConvertDateToString -value $value -formatOut $frmtdateOUT -locale $dateLocale
                }
                
                $hashtable[$rowKey][$col] = $value
            }
        }
    }

    return $hashtable
}

function ExportCsv {
    param ( [string]$filepath, [hashtable]$hash, [string]$header, [string]$delimiter)

    $out = @()
    LOG "ExportCsv" "Creation du fichier $filepath"
    $out += $header
    $header = $header -replace '"', ''
    $fields = $header -split $delimiter
    foreach ($key in $hash.Keys) {
        $line = ""
        foreach ($field in $fields) { 
            if ( [string]::IsNullOrEmpty($hash[$key][$field]) ) {
                $val = ""
            } else {
                $val = $hash[$key][$field]
            }
            $line += '"' + $val + '"' + $delimiter 
        }
        $out += $line   
    }
    $out | Out-File -FilePath $filepath -Encoding UTF8
} # code   param -filepath -hash -header -delimiter                           Return : N/A         >> Exporte (hash[][]) vers (filepath) avec (header) et (delimiter)
