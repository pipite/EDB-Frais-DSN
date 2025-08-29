# Fonction utilitaire pour l'accès Excel avec gestion d'erreurs centralisée

function Invoke-ExcelQuery {
    <#
    .SYNOPSIS
        Exécute une requête SQL sur un fichier Excel avec gestion optimisée des dates
    
    .PARAMETER filePath
        Chemin vers le fichier Excel
    
    .PARAMETER sqlQuery
        Requête SQL à exécuter
    
    .PARAMETER functionName
        Nom de la fonction appelante (pour les logs)
    
    .PARAMETER header
        Header personnalisé pour les colonnes
    
    .PARAMETER columnMapping
        Mapping des noms de colonnes vers leurs positions
    
    .PARAMETER frmtdateOUT
        Format de sortie pour les dates (ex: "yyyy-MM-dd", "dd/MM/yyyy")
    
    .PARAMETER dateLocale
        Locale pour l'interprétation des dates ("FR" ou "US")
    
    .PARAMETER datecol
        Liste des colonnes contenant des dates à convertir
    
    .EXAMPLE
        # Conversion des dates dans les colonnes spécifiées
        Invoke-ExcelQuery -filePath "data.xlsx" -sqlQuery "SELECT * FROM [Sheet1$]" -functionName "Test" -frmtdateOUT "yyyy-MM-dd" -datecol @("Date_Entree", "Date_Sortie")
    
    .EXAMPLE
        # Avec locale US pour interpréter 04/02/2000 comme 2 avril 2000
        Invoke-ExcelQuery -filePath "data.xlsx" -sqlQuery "SELECT * FROM [Sheet1$]" -functionName "Test" -frmtdateOUT "yyyy-MM-dd" -dateLocale "US" -datecol @("StartDate", "EndDate")
    #>
    param(
        [string]$filePath,
        [string]$sqlQuery,
        [string]$functionName,
        [string[]]$header = $null,
        [hashtable]$columnMapping = $null,
        [string]$frmtdateOUT = $null,
        [string]$dateLocale = "FR",  # "FR" pour dd/MM/yyyy, "US" pour MM/dd/yyyy
        [string[]]$datecol = @()     # Liste des colonnes contenant des dates
    )
    
    DBG "Invoke-ExcelQuery" "Exécution requête Excel via $functionName"
    
    # Si un columnMapping est fourni, on remplace les noms personnalisés par les positions dans la requête
    if ($columnMapping) {
        DBG "Invoke-ExcelQuery" "Utilisation d'un mapping de colonnes pour éviter les doublons"
        $modifiedQuery = $sqlQuery
        
        foreach ($personalizedName in $columnMapping.Keys) {
            $position = $columnMapping[$personalizedName]
            $columnName = "F$position"
            $modifiedQuery = $modifiedQuery -replace "\[$personalizedName\]", "[$columnName]"
        }
        
        $sqlQuery = $modifiedQuery
        # Utiliser IMEX=1 et MaxScanRows=0 pour forcer une analyse complète en mode mixte
        $connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$filePath;Extended Properties='Excel 12.0 Xml;HDR=NO;IMEX=1;MaxScanRows=0';"
        DBG "Invoke-ExcelQuery" "Requête modifiée : $sqlQuery"
    }
    # Si un header personnalisé est fourni (sans columnMapping), on ignore le header du fichier Excel
    elseif ($header) {
        $connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$filePath;Extended Properties='Excel 12.0 Xml;HDR=NO';"
        DBG "Invoke-ExcelQuery" "Utilisation d'un header personnalisé avec $($header.Count) colonnes"
    } else {
        $connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$filePath;Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1';"
        DBG "Invoke-ExcelQuery" "Utilisation du header du fichier Excel"
    }
    
    $connection = New-Object System.Data.OleDb.OleDbConnection
    $connection.ConnectionString = $connectionString
    
    try {
        $connection.Open()
        $cmd = $connection.CreateCommand()
        $cmd.CommandText = $sqlQuery
        $adapter = New-Object System.Data.OleDb.OleDbDataAdapter($cmd)
        $table = New-Object System.Data.DataTable
        $n = $adapter.Fill($table)
        $connection.Close()
        
        # Si un header personnalisé est fourni, renommer les colonnes
        if ($header) {
            # Vérifier que le nombre de colonnes correspond
            if ($header.Count -ne $table.Columns.Count) {
                QUITEX $functionName "Le nombre de colonnes dans le header personnalisé ($($header.Count)) ne correspond pas au nombre de colonnes dans le fichier ($($table.Columns.Count))"
            }
            
            # Renommer les colonnes avec le header personnalisé
            for ($i = 0; $i -lt $header.Count; $i++) {
                $table.Columns[$i].ColumnName = $header[$i]
            }
            
            # Si HDR=NO (header personnalisé ou columnMapping), la première ligne contient les données du header original, on la supprime
            if ($table.Rows.Count -gt 0 -and ($header -or $columnMapping)) {
                $table.Rows.RemoveAt(0)
            }
        }
        # Si columnMapping est utilisé sans header personnalisé, on utilise les noms du mapping
        elseif ($columnMapping) {
            # Créer un mapping inverse (position -> nom personnalisé)
            $inverseMapping = @{}
            foreach ($name in $columnMapping.Keys) {
                $position = $columnMapping[$name]
                $inverseMapping["F$position"] = $name
            }
            
            # Renommer les colonnes retournées par la requête
            foreach ($col in $table.Columns) {
                $originalName = $col.ColumnName
                if ($inverseMapping.ContainsKey($originalName)) {
                    $col.ColumnName = $inverseMapping[$originalName]
                    DBG "Invoke-ExcelQuery" "Renommage colonne $originalName -> $($col.ColumnName)"
                }
            }
            
            # Supprimer la première ligne (header original)
            if ($table.Rows.Count -gt 0) {
                $table.Rows.RemoveAt(0)
            }
        }
        
        # Convertir les dates en chaîne selon le format spécifié (OPTIMISÉ avec liste explicite)
        if ($frmtdateOUT -and $datecol.Count -gt 0) {
            DBG "Invoke-ExcelQuery" "Conversion des dates au format $frmtdateOUT (locale: $dateLocale)"
            DBG "Invoke-ExcelQuery" "Colonnes de dates spécifiées : $($datecol -join ', ')"
            DBG "Invoke-ExcelQuery" "Paramètre frmtdateOUT reçu : [$frmtdateOUT]"
            
            # Vérifier que les colonnes spécifiées existent dans la table
            $validDateColumns = @()
            foreach ($colName in $datecol) {
                if ($table.Columns.Contains($colName)) {
                    $validDateColumns += $colName
                } else {
                    WRN "Invoke-ExcelQuery" "Colonne de date '$colName' introuvable dans la table"
                }
            }
            
            if ($validDateColumns.Count -gt 0) {
                $dateConversions = 0
                foreach ($row in $table.Rows) {
                    foreach ($colName in $validDateColumns) {
                        $originalValue = $row[$colName]
                        if ($originalValue -ne $null -and $originalValue.ToString().Trim() -ne "" -and $originalValue.ToString().Trim() -ne "-") {
                            $convertedValue = ConvertDateToString -value $originalValue -formatOut $frmtdateOUT -locale $dateLocale
                            
                            # Maintenant que les dates sont des chaînes depuis Excel, on peut les assigner directement
                            $row[$colName] = $convertedValue
                            if ($convertedValue -ne $originalValue.ToString()) {
                                $dateConversions++
                            }
                        }
                    }
                }
                
                if ($dateConversions -gt 0) {
                    DBG "Invoke-ExcelQuery" "$dateConversions dates converties dans $($validDateColumns.Count) colonnes"
                } else {
                    DBG "Invoke-ExcelQuery" "Aucune conversion de date nécessaire"
                }
            } else {
                WRN "Invoke-ExcelQuery" "Aucune colonne de date valide trouvée"
            }
        } elseif ($frmtdateOUT -and $datecol.Count -eq 0) {
            DBG "Invoke-ExcelQuery" "Format de date spécifié mais aucune colonne de date fournie - conversion ignorée"
        }
        
        return @{
            Success = $true
            Table = $table
            RowCount = $n
            Error = $null
        }
    } catch {
        if ($connection.State -eq 'Open') {
            $connection.Close()
        }
        return @{
            Success = $false
            Table = $null
            RowCount = 0
            Error = $_.Exception.Message
        }
    }
}
