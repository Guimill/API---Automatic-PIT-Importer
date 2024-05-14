"""
/***************************************************************************
 API

                              -------------------
        begin                : 10/04/2024
        copyright            : (C) 2023 by Guillaume Milleret
        email                : guillaume.milleret@gmail.com
 ***************************************************************************/

/***************************************************************************
 *                                                                         *
 *   This program is free software; you can redistribute it and/or modify  *
 *   it under the terms of the GNU General Public License as published by  *
 *   the Free Software Foundation; either version 2 of the License, or     *
 *   (at your option) any later version.                                   *
 *                                                                         *
 ***************************************************************************/
"""


######################### API - Automatic Pit Importer ######################################

$programFiles = "C:\Program Files"
$qgisFolders = Get-ChildItem $programFiles | Where-Object { $_.Name -like "QGIS*" } | Sort-Object Name
$qgisVersion = $qgisFolders[-1].Name
$OSGEO4W_ROOT = "$programFiles\$qgisVersion"

################ Find a way to now wich distribution is used #################################


$env:Path = "$OSGEO4W_ROOT\bin;$env:Path"
$env:GDAL_DATA = "$OSGEO4W_ROOT\share\gdal"
$env:GDAL_DRIVER_PATH = "$OSGEO4W_ROOT\bin\gdalplugins"
$env:PROJ_LIB = "$OSGEO4W_ROOT\share\proj"
$env:PYTHONHOME = "$OSGEO4W_ROOT\apps\Python39"


# Set up user input for the directory path
$PATHWAY_NRO = Read-Host "Entrer le dossier mère "

# Set up user input for the directory path
$PATHWAY = "$PATHWAY_NRO\PIT"



############################### Unzip all the files ################################################



# Set the number of iterations
$iterations = 2

# Loop for the specified number of iterations
for ($i = 1; $i -le $iterations; $i++) {
    # Get a list of all zip files in the directory and its subdirectories
    $zipFiles = Get-ChildItem -Path $PATHWAY_NRO -Filter *.zip -Recurse

    # Check if zip files were found and dezziper
    if ($zipFiles.Count -gt 0) {
        # Iterate through each zip file and unzip it
        foreach ($zipFile in $zipFiles) {
            $destinationPath = Join-Path -Path $zipFile.Directory.FullName -ChildPath $zipFile.BaseName
            Expand-Archive -Path $zipFile.FullName -DestinationPath $destinationPath -Force
        }
    } else {
        Write-Host "No zip files found in the directory: $directory"
    }
}


############################### Créer les variables et les csv contenant le appui et ch orange depuis la t_ptech ########################


# Path to the CSV file
$csvFilePath_t_ptech = Get-ChildItem -Path "$PATHWAY_NRO\1_JDD" -Filter t_ptech.csv -Recurse


# Output file path
$outputFilePath_ap_or = "$PATHWAY\appui_orange.csv"
$outputFilePath_ch_or = "$PATHWAY\chambre_orange.csv"

# Load the CSV file
$data_t_ptech = Import-Csv -Path $csvFilePath_t_ptech.FullName -Delimiter ';'

# Filter the data based on your SQL-like conditions and select only the column "pt_codeext"
$filteredData_ch_or = $data_t_ptech | Where-Object { $_.pt_prop -like '*ORMB0000000001*' -and $_.pt_typephy -like '*C*' -and $null -eq $_.pt_ad_code }
$filteredData_ap_or = $data_t_ptech | Where-Object { $_.pt_prop -like '*ORMB0000000001*' -and $_.pt_typephy -like '*A*' -and $null -eq $_.pt_ad_code }

# Debug Output: Check filtered data count
Write-Host "Filtered chambre data count: $($filteredData_ch_or.Count)"
Write-Host "Filtered appui data count: $($filteredData_ap_or.Count)"

# Export the filtered data to the output file
$filteredData_ch_or | Export-Csv -Path $outputFilePath_ch_or -NoTypeInformation -Delimiter ';'
$filteredData_ap_or | Export-Csv -Path $outputFilePath_ap_or -NoTypeInformation -Delimiter ';'



############################### Créer le code chambre dans chambre.csv de GFI ########################



# Set the file path
$ChambreGFIFilePath = "$PATHWAY\GFI\Chambre.csv"

# Get the Chambre.csv file in the pathway
$csvFile_FGI = Get-ChildItem -Path $ChambreGFIFilePath -Filter "Chambre.csv"

# Check if the file exists
if ($csvFile_FGI) {
    # Check if CONCAT_CH column already exists
    $ogrinfoOutput = ogrinfo -q -sql "SELECT concat_ch_gfi FROM Chambre" $csvFile_FGI.FullName
    if ($ogrinfoOutput -notlike "*ERROR*Column*not*found*") {
        Write-Host "concat_ch_gfi column already exists in $($csvFile_FGI.FullName). Skipping creation."
    }
    else {
        # Run the ogrinfo command to add the CONCAT_CH column
        ogrinfo -dialect OGRSQL -sql "ALTER TABLE Chambre ADD COLUMN concat_ch_gfi VARCHAR(15);" $csvFile_FGI.FullName
        Write-Host "concat_ch_gfi column added to $($csvFile_FGI.FullName)."
    }
    
    # Run the ogrinfo command to update CONCAT_CH values
    ogrinfo -dialect SQLite -sql "UPDATE Chambre SET concat_ch_gfi = printf('%05d', code_ch1) || '/' || code_ch2;" $csvFile_FGI.FullName
    Write-Host "concat_ch_gfi values updated in $($csvFile_FGI.FullName)."
} else {
    Write-Host "No Chambre.csv file found in $($ChambreGFIFilePath)."
}



############################### Créer les codes chambres pour les FT_Chambre.shp UNIR et SEAFILE ########################



# Use Get-ChildItem to iterate through all FT_Chambre.shp files recursively in the specified directory
Get-ChildItem -Path $PATHWAY -Recurse -Filter "FT_Chambre.shp" | ForEach-Object {
    # Check if CONCAT_CH column already exists
    $ogrinfoOutput = ogrinfo -q -sql "SELECT CONCAT_CH FROM FT_Chambre" $_.FullName
    if ($ogrinfoOutput -notlike "*ERROR*Column*not*found*") {
        Write-Host "CONCAT_CH column already exists in $($_.FullName). Skipping creation."
    }
    else {
        # Run the ogrinfo command to add the CONCAT_CH column
        ogrinfo -dialect OGRSQL -sql "ALTER TABLE FT_Chambre ADD COLUMN CONCAT_CH VARCHAR(15);" $_.FullName
    }
    
    # Run the ogrinfo command to update CONCAT_CH values
    ogrinfo -dialect SQLite -sql "UPDATE FT_Chambre SET CONCAT_CH = printf('%05d', code_ch1) || '/' || code_ch2;" $_.FullName
    
    # Run the ogrinfo command to update CODE_VOIE values (if needed)
    ogrinfo -dialect SQLite -sql "UPDATE FT_Chambre SET CODE_VOIE=NULL" $_.FullName
}



############################### Créer les codes appui pour les FtAppui.shp UNIR et SEAFILE ########################



# Use Get-ChildItem to iterate through all FT_Chambre.shp files recursively in the specified directory
Get-ChildItem -Path $PATHWAY -Recurse -Filter "FT_Appui.shp" | ForEach-Object {
    # Check if CONCAT_CH column already exists
    $ogrinfoOutput = ogrinfo -q -sql "SELECT CONCAT_AP FROM FT_Appui" $_.FullName
    if ($ogrinfoOutput -notlike "*ERROR*Column*not*found*") {
        Write-Host "CONCAT_AP column already exists in $($_.FullName). Skipping creation."
    }
    else {
        # Run the ogrinfo command to add the CONCAT_CH column
        ogrinfo -dialect OGRSQL -sql "ALTER TABLE FT_Appui ADD COLUMN CONCAT_AP VARCHAR(15);" $_.FullName
    }
    
    # Run the ogrinfo command to update CONCAT_CH values
    ogrinfo -dialect SQLite -sql "UPDATE FT_Appui SET CONCAT_AP = printf('%07d', NUM_APPUI) || '/' || CODE_COMMU;" $_.FullName
    
    # Run the ogrinfo command to update CODE_VOIE values (if needed)
    ogrinfo -dialect SQLite -sql "UPDATE FT_Appui SET CODE_VOIE=NULL" $_.FullName
}



################################ Compte le nombre de chambre présente dans le pt_codeext en fonction de chaque shape ########################
#
## Arrays to store counts
#$chambreCounts = @()
#$appuiCounts = @()
#
## Loop through each shapefile for FT_Chambre
#foreach ($shapefile in $shapefiles_ch) {
#    # Execute for each shapefile
#    $shapefilePath_ch = $shapefile.FullName
#    
#    # Get the shortened directory name
#    $shortenedDirectoryName = $shapefile.DirectoryName.Replace($PATHWAY, "")
#    
#    # Execute OGR SQL query to select matching features from shapefile
#    $query_FT_ch = "SELECT * FROM FT_Chambre WHERE CONCAT_CH IN $csvValuesString_ch_or"
#    $ogrOutput_FT_ch = ogrinfo -q -dialect SQLite -sql $query_FT_ch $shapefilePath_ch
#    
#    $concatCount_FT_ch = ($ogrOutput_FT_ch -match "CONCAT_CH").Count
#    
#    # Store count in chambreCounts array
#    $chambreCounts += [PSCustomObject]@{
#        Directory = $shortenedDirectoryName
#        Count = $concatCount_FT_ch
#    }
#}
#
## Loop through each shapefile for FT_Appui
#foreach ($shapefile in $shapefiles_ap) {
#    # Execute for each shapefile
#    $shapefilePath_ap = $shapefile.FullName
#    
#    # Get the shortened directory name
#    $shortenedDirectoryName = $shapefile.DirectoryName.Replace($PATHWAY, "")
#    
#    # Execute OGR SQL query to select matching features from shapefile
#    $query_FT_ap = "SELECT * FROM FT_Appui WHERE CONCAT_AP IN $csvValuesString_ap_or"
#    $ogrOutput_FT_ap = ogrinfo -q -dialect SQLite -sql $query_FT_ap $shapefilePath_ap
#    
#    $concatCount_FT_ap = ($ogrOutput_FT_ap -match "CONCAT_AP").Count
#    
#    # Store count in appuiCounts array
#    $appuiCounts += [PSCustomObject]@{
#        Directory = $shortenedDirectoryName
#        Count = $concatCount_FT_ap
#    }
#}
#
## Execute OGR SQL query for the CSV file
#$query_GFI_ap = "SELECT * FROM Appui WHERE id_metier_site IN $csvValuesString_ap_or"
#$ogrOutput_GFI_ap = ogrinfo -q -dialect SQLite -sql $query_GFI_ap $outputFilePath__GFI_ap
#
## Count the number of matching entities
#$concatCount_GFI_ap = ($ogrOutput_GFI_ap -match "id_metier_site").Count
#
## Add the count for Appui to the appuiCounts array
#$appuiCounts += [PSCustomObject]@{
#    Directory = "GFI\Appui"
#    Count = $concatCount_GFI_ap
#}
#
## Execute OGR SQL query for the CSV file
#$query_GFI_ch = "SELECT * FROM Chambre WHERE concat_ch_gfi IN $csvValuesString_ch_or"
#$ogrOutput_GFI_ch = ogrinfo -q -dialect SQLite -sql $query_GFI_ch $outputFilePath__GFI_ch
#
## Count the number of matching entities
#$concatCount_GFI_ch = ($ogrOutput_GFI_ch -match "concat_ch_gfi").Count
#
## Add the count for Appui to the appuiCounts array
#$chambreCounts += [PSCustomObject]@{
#    Directory = "GFI\Chambre"
#    Count = $concatCount_GFI_ch
#}
#
## Export FT_Chambre and FT_Appui counts to CSV
#$chambreCounts | Export-Csv -Path "$PATHWAY\FT_Chambre_Counts.csv" -NoTypeInformation -Delimiter ";"
#$appuiCounts | Export-Csv -Path "$PATHWAY\FT_Appui_Counts.csv" -NoTypeInformation -Delimiter ";"
#
#
#
#
################################ Récupération des INSEE appui #################################################


# Initialize an array to hold the full paths of found shapefiles
$directories_ap = @()

# Recursively search for all occurrences of "FT_Appui.shp" and collect their full paths
Get-ChildItem -Path $PATHWAY -Recurse -Filter "FT_Appui.shp" | ForEach-Object {
    # Append the full path of each found shapefile to the array
    $directories_ap += $_.FullName
}

$inseeCodes_ap = @()
foreach ($directory in $directories_ap) {
    # Trouver la première séquence de cinq chiffres entourée de soulignés
    if ($directory -match "_\d{5}_") {
        # Capture les cinq chiffres sans les underscores
        $inseeCode = $matches[0] -replace "_", ""
        $inseeCodes_ap += $inseeCode
    }
}

# Get unique INSEE codes
$uniqueInseeCodes_ap = $inseeCodes_ap | Select-Object -Unique

############################### Récupération des INSEE chambre #################################################


# Initialize an array to hold the full paths of found shapefiles
$directories_ch = @()

# Recursively search for all occurrences of "FT_Appui.shp" and collect their full paths
Get-ChildItem -Path $PATHWAY -Recurse -Filter "FT_Chambre.shp" | ForEach-Object {
    # Append the full path of each found shapefile to the array
    $directories_ch += $_.FullName
}

$inseeCodes_ch = @()
foreach ($directory in $directories_ch) {
    # Trouver la première séquence de cinq chiffres entourée de soulignés
    if ($directory -match "_\d{5}_") {
        # Capture les cinq chiffres sans les underscores
        $inseeCode = $matches[0] -replace "_", ""
        $inseeCodes_ch += $inseeCode
    }
}

# Get unique INSEE codes
$uniqueInseeCodes_ch = $inseeCodes_ch | Select-Object -Unique

############################### Création de différent csv, pour chaque insee, pour comparer les appui #######################################



# Import the CSV with data for the SQL query
$csvAppuiPath = Join-Path -Path $PATHWAY -ChildPath "appui_orange.csv"
$csvAppuiData = Import-Csv -Path $csvAppuiPath -Delimiter ";"

# Extract the unique pt_codeext values
$ptCodeExtValues_ap = $csvAppuiData | ForEach-Object { "'$($_.pt_codeext)'" }
$formattedValues_ap = $ptCodeExtValues_ap -join ", "

# Process each FT_Appui.shp for each unique INSEE code
foreach ($inseeCode in $uniqueInseeCodes_ap) {
    # Create the output path for this INSEE code
    $finalAppuiCsvPath_Seafile = Join-Path -Path $PATHWAY -ChildPath "Appui_SEAFILE_$inseeCode.csv"
    $finalAppuiCsvPath_UNIR = Join-Path -Path $PATHWAY -ChildPath "Appui_UNIR_$inseeCode.csv"

    # Find the shapefile that matches the INSEE code
    $ftAppuiShapefilesSeafile = $directories_ap | Where-Object {($_ -match "(^|_)$inseeCode(_|$)") -and ($_ -like "*SEAFILE*")}

    $ftAppuiShapefilesUNIR = $directories_ap | Where-Object {($_ -match "(^|_)$inseeCode(_|$)") -and ($_ -like "*UNIR*")}

    $query_ap = "SELECT CONCAT_AP FROM FT_Appui WHERE CONCAT_AP IN ($formattedValues_ap)"

    # Execute ogr2ogr to generate a CSV for this INSEE code
    ogr2ogr -f "CSV" -dialect SQLite -sql $query_ap $finalAppuiCsvPath_Seafile $ftAppuiShapefilesSeafile
    ogr2ogr -f "CSV" -dialect SQLite -sql $query_ap $finalAppuiCsvPath_UNIR $ftAppuiShapefilesUNIR

    # Convert commas to semicolons
    $intermediateAppuiCsvData_Seafile = Get-Content $finalAppuiCsvPath_Seafile
    $convertedAppuiCsvData_Seafile = $intermediateAppuiCsvData_Seafile -replace ",", ";"
    
    $intermediateAppuiCsvData_UNIR = Get-Content $finalAppuiCsvPath_UNIR
    $convertedAppuiCsvData_UNIR = $intermediateAppuiCsvData_UNIR -replace ",", ";"

    $convertedAppuiCsvData_Seafile | Out-File -FilePath $finalAppuiCsvPath_Seafile -Encoding utf8
    $convertedAppuiCsvData_UNIR | Out-File -FilePath $finalAppuiCsvPath_UNIR -Encoding utf8

}


############################### Création de différent csv, pour chaque insee, pour comparer les chambres #######################################



# Import the CSV with data for the SQL query
$csvChambrePath = Join-Path -Path $PATHWAY -ChildPath "chambre_orange.csv"
$csvChambreData = Import-Csv -Path $csvChambrePath -Delimiter ";"

# Extract the unique pt_codeext values
$ptCodeExtValues_ch = $csvChambreData | ForEach-Object { "'$($_.pt_codeext)'" }
$formattedValues_ch = $ptCodeExtValues_ch -join ", "

# Process each FT_Appui.shp for each unique INSEE code
foreach ($inseeCode in $uniqueInseeCodes_ch) {
    # Create the output path for this INSEE code
    $finalChambreCsvPath_Seafile = Join-Path -Path $PATHWAY -ChildPath "Chambre_SEAFILE_$inseeCode.csv"
    $finalChambreCsvPath_UNIR = Join-Path -Path $PATHWAY -ChildPath "Chambre_UNIR_$inseeCode.csv"

    # Find the shapefile that matches the INSEE code
    $ftChambreShapefilesSeafile = $directories_ch | Where-Object {($_ -match "(^|_)$inseeCode(_|$)") -and ($_ -like "*SEAFILE*")}

    $ftChambreShapefilesUNIR = $directories_ch | Where-Object {($_ -match "(^|_)$inseeCode(_|$)") -and ($_ -like "*UNIR*")}

    $query_ch = "SELECT CONCAT_CH FROM FT_Chambre WHERE CONCAT_CH IN ($formattedValues_ch)"

    # Execute ogr2ogr to generate a CSV for this INSEE code
    ogr2ogr -f "CSV" -dialect SQLite -sql $query_ch $finalChambreCsvPath_Seafile $ftChambreShapefilesSeafile
    ogr2ogr -f "CSV" -dialect SQLite -sql $query_ch $finalChambreCsvPath_UNIR $ftChambreShapefilesUNIR

    # Convert commas to semicolons
    $intermediateChambreCsvData_Seafile = Get-Content $finalChambreCsvPath_Seafile
    $convertedChambreCsvData_Seafile = $intermediateChambreCsvData_Seafile -replace ",", ";"
    
    $intermediateChambreCsvData_UNIR = Get-Content $finalChambreCsvPath_UNIR
    $convertedChambreCsvData_UNIR = $intermediateChambreCsvData_UNIR -replace ",", ";"

    $convertedChambreCsvData_Seafile | Out-File -FilePath $finalChambreCsvPath_Seafile -Encoding utf8
    $convertedChambreCsvData_UNIR | Out-File -FilePath $finalChambreCsvPath_UNIR -Encoding utf8

}



############################## Left Join sur de tout les appui pour un code insee ############################################

############################## Ne pas oublier faire un Résultat par INSEE ############################################

#foreach ($inseeCode in $uniqueInseeCodes_ch) {
#        # Importez les fichiers CSV
#        $csv1_ap = Import-Csv -Path "$PATHWAY\Appui_SEAFILE_$inseeCode.csv" -Delimiter ';'
#        $csv2_ap = Import-Csv -Path "$PATHWAY\Appui_UNIR_$inseeCode.csv" -Delimiter ';'
#        $csv3_ap = Import-Csv -Path "$PATHWAY\GFI\Appui.csv" -Delimiter ';'
#        $csv4_ap = Import-Csv -Path "$PATHWAY\appui_orange.csv" -Delimiter ';'
#     
#        # Mettez les données dans des dictionnaires pour faciliter les jointures
#        $seafile_dict_ap = @{}
#        $csv1_ap | ForEach-Object { $seafile_dict_ap[$_.CONCAT_AP] = $_ }
#     
#        $unir_dict_ap = @{}
#        $csv2_ap | ForEach-Object { $unir_dict_ap[$_.CONCAT_AP] = $_ }
#     
#        $gfi_dict_ap = @{}
#        $csv3_ap | ForEach-Object { $gfi_dict_ap[$_.id_metier_site] = $_ }
#
#        # Initialise un tableau pour le CSV fusionné
#        $merged_appui = @()
#
#
#            foreach ($record in $csv4_ch) {
#                    # Obtenez les 5 derniers chiffres du 'pt_codeext'
#                    $insee_from_pt_codeext = $record.pt_codeext.Substring($record.pt_codeext.Length - 5, 5)
#        
#                    # Effectuez le join uniquement si les codes INSEE correspondent
#                    if ($insee_from_pt_codeext -eq $inseeCode) {
#                        $key = $record.pt_codeext
#        
#                        # Créez un objet pour contenir les valeurs
#                $row = [PSCustomObject]@{
#                    appui_orange = $record.pt_codeext
#                    gfi = if ($gfi_dict_ap.ContainsKey($key)) { $gfi_dict_ap[$key].id_metier_site } else { "" }
#                    seafile = if ($seafile_dict_ap.ContainsKey($key)) { $seafile_dict_ap[$key].CONCAT_AP } else { "" }
#                    unir = if ($unir_dict_ap.ContainsKey($key)) { $unir_dict_ap[$key].CONCAT_AP } else { "" }
#                }
#
#                $merged_appui += $row
#            }
#        }
#
#        # Triez par une propriété donnée (par exemple, appui_orange)
#        $sortedCsv_ap = $merged_appui | Sort-Object -Property appui_orange
#
#        # Exportez le CSV fusionné trié
#        $sortedCsv_ap | Export-Csv -Path "$PATHWAY\Resultat_Appui_$inseeCode.csv" -NoTypeInformation -Delimiter ';'
#}
#


foreach ($inseeCode in $uniqueInseeCodes_ch) {
    # Importez les fichiers CSV
    $csv1_ap = Import-Csv -Path "$PATHWAY\Appui_SEAFILE_$inseeCode.csv" -Delimiter ';'
    $csv2_ap = Import-Csv -Path "$PATHWAY\Appui_UNIR_$inseeCode.csv" -Delimiter ';'
    $csv3_ap = Import-Csv -Path "$PATHWAY\GFI\Appui.csv" -Delimiter ';'
    $csv4_ap = Import-Csv -Path "$PATHWAY\appui_orange.csv" -Delimiter ';'

    # Mettez les données dans des hashtables
    $seafile_dict_ap = @{}
    foreach ($item in $csv1_ap) {
        $seafile_dict_ap[$item.CONCAT_AP] = $item
    }

    $unir_dict_ap = @{}
    foreach ($item in $csv2_ap) {
        $unir_dict_ap[$item.CONCAT_AP] = $item
    }

    $gfi_dict_ap = @{}
    foreach ($item in $csv3_ap) {
        $gfi_dict_ap[$item.id_metier_site] = $item
    }

    # Initialise un tableau pour le CSV fusionné
    $merged_appui = @()

    # Parcourez le CSV d'appui_orange pour effectuer des jointures
    foreach ($record in $csv4_ap) {
        # Obtenez les 5 derniers chiffres du 'pt_codeext'
        $insee_from_pt_codeext = $record.pt_codeext.Substring($record.pt_codeext.Length - 5, 5)

        # Effectuez le join uniquement si les codes INSEE correspondent
        if ($insee_from_pt_codeext -eq $inseeCode) {
            $key = $record.pt_codeext

            # Créez des hashtables pour stocker les valeurs
            $row = @{}
            $row["appui_orange"] = $record.pt_codeext
            $row["gfi"] = if ($gfi_dict_ap.ContainsKey($key)) { $gfi_dict_ap[$key].id_metier_site } else { "" }
            $row["seafile"] = if ($seafile_dict_ap.ContainsKey($key)) { $seafile_dict_ap[$key].CONCAT_AP } else { "" }
            $row["unir"] = if ($unir_dict_ap.ContainsKey($key)) { $unir_dict_ap[$key].CONCAT_AP } else { "" }

            $merged_appui += $row
        }
    }

    # Triez par une propriété donnée (par exemple, appui_orange)
    $sortedCsv_ap = $merged_appui | Sort-Object -Property "appui_orange"

    # Déterminez les en-têtes des colonnes
    $headers = @("appui_orange", "gfi", "seafile", "unir")

    # Créez le CSV sous forme de chaîne de texte
    $csv_text = [string]::Join(";", $headers) + "`n"

    # Ajoutez chaque ligne au CSV
    foreach ($item in $sortedCsv_ap) {
        $csv_text += [string]::Join(";", ($item["appui_orange"], $item["gfi"], $item["seafile"], $item["unir"])) + "`n"
    }

    # Enregistrez le CSV final dans un fichier
    Set-Content -Path "$PATHWAY\Resultat_Appui_$inseeCode.csv" -Value $csv_text
}


############################## Left Join sur de tout les chambre pour un code insee ############################################

#foreach ($inseeCode in $uniqueInseeCodes_ch) {
#    # Importez les fichiers CSV
#    $csv1_ch = Import-Csv -Path "$PATHWAY\Chambre_SEAFILE_$inseeCode.csv" -Delimiter ';'
#    $csv2_ch = Import-Csv -Path "$PATHWAY\Chambre_UNIR_$inseeCode.csv" -Delimiter ';'
#    $csv3_ch = Import-Csv -Path "$PATHWAY\GFI\Chambre.csv" -Delimiter ';'
#    $csv4_ch = Import-Csv -Path "$PATHWAY\chambre_orange.csv" -Delimiter ';'
#
#    # Mettez les données dans des dictionnaires pour faciliter les jointures
#    $seafile_dict_ch = @{}
#    $csv1_ch | ForEach-Object { $seafile_dict_ch[$_.CONCAT_CH] = $_ }
#
#    $unir_dict_ch = @{}
#    $csv2_ch | ForEach-Object { $unir_dict_ch[$_.CONCAT_CH] = $_ }
#
#    $gfi_dict_ch = @{}
#    $csv3_ch | ForEach-Object { $gfi_dict_ch[$_.concat_ch_gfi] = $_ }
#
#    # Initialise un tableau pour le CSV fusionné
#    $merged_chambre = @()
#
#        foreach ($record in $csv4_ch) {
#                # Obtenez les 5 derniers chiffres du 'pt_codeext'
#                $insee_from_pt_codeext = $record.pt_codeext.Substring($record.pt_codeext.Length - 5, 5)
#        
#                # Effectuez le join uniquement si les codes INSEE correspondent
#                if ($insee_from_pt_codeext -eq $inseeCode) {
#                    $key = $record.pt_codeext
#        
#                    # Créez un objet pour contenir les valeurs
#                    $row = [PSCustomObject]@{
#                        chambre_orange = $record.pt_codeext
#                        gfi = if ($gfi_dict_ch.ContainsKey($key)) { $gfi_dict_ch[$key].concat_ch_gfi } else { "" }
#                        seafile = if ($seafile_dict_ch.ContainsKey($key)) { $seafile_dict_ch[$key].CONCAT_CH } else { "" }
#                        unir = if ($unir_dict_ch.ContainsKey($key)) { $unir_dict_ch[$key].CONCAT_CH } else { "" }
#                    }
#
#                    $merged_chambre += $row
#                }
#            }
#
#            
#    # Triez par une propriété donnée (par exemple, appui_orange)
#    $sortedCsv_ch = $merged_chambre | Sort-Object -Property chambre_orange
#
#    # Exportez le CSV fusionné trié
#    $sortedCsv_ch | Export-Csv -Path "$PATHWAY\Resultat_Chambre_$inseeCode.csv" -NoTypeInformation -Delimiter ';'
#}
#

foreach ($inseeCode in $uniqueInseeCodes_ch) {
    # Importez les fichiers CSV
    $csv1_ch = Import-Csv -Path "$PATHWAY\Chambre_SEAFILE_$inseeCode.csv" -Delimiter ';'
    $csv2_ch = Import-Csv -Path "$PATHWAY\Chambre_UNIR_$inseeCode.csv" -Delimiter ';'
    $csv3_ch = Import-Csv -Path "$PATHWAY\GFI\Chambre.csv" -Delimiter ';'
    $csv4_ch = Import-Csv -Path "$PATHWAY\chambre_orange.csv" -Delimiter ';'

    # Mettez les données dans des hashtables
    $seafile_dict_ch = @{}
    foreach ($item in $csv1_ch) {
        $seafile_dict_ch[$item.CONCAT_CH] = $item
    }

    $unir_dict_ch = @{}
    foreach ($item in $csv2_ch) {
        $unir_dict_ch[$item.CONCAT_CH] = $item
    }

    $gfi_dict_ch = @{}
    foreach ($item in $csv3_ch) {
        $gfi_dict_ch[$item.concat_ch_gfi] = $item
    }

    # Initialisez le tableau fusionné
    $merged_chambre = @()

    # Construisez les enregistrements en utilisant des hashtables
    foreach ($record in $csv4_ch) {
        $insee_from_pt_codeext = $record.pt_codeext.Substring($record.pt_codeext.Length - 5, 5)

        if ($insee_from_pt_codeext -eq $inseeCode) {
            $key = $record.pt_codeext

            $row = @{}
            $row["chambre_orange"] = $record.pt_codeext
            $row["gfi"] = if ($gfi_dict_ch.ContainsKey($key)) { $gfi_dict_ch[$key].concat_ch_gfi } else { "" }
            $row["seafile"] = if ($seafile_dict_ch.ContainsKey($key)) { $seafile_dict_ch[$key].CONCAT_CH } else { "" }
            $row["unir"] = if ($unir_dict_ch.ContainsKey($key)) { $unir_dict_ch[$key].CONCAT_CH } else { "" }

            $merged_chambre += $row
        }
    }

    # Triez les données par une propriété donnée
    $sortedCsv_ch = $merged_chambre | Sort-Object -Property "chambre_orange"

    # Déterminez les en-têtes des colonnes
    $headers = @("chambre_orange", "seafile", "gfi", "unir")

    # Créez le CSV sous forme de chaîne de texte
    $csv_text = [string]::Join(";", $headers) + "`n"

    # Ajoutez chaque ligne au CSV
    foreach ($item in $sortedCsv_ch) {
        $csv_text += [string]::Join(";", ($item["chambre_orange"], $item["seafile"], $item["gfi"], $item["unir"])) + "`n"
    }

    # Enregistrez le CSV final dans un fichier
    Set-Content -Path "$PATHWAY\Resultat_Chambre_$inseeCode.csv" -Value $csv_text
}



############################## Suppression de toutes les collonnes CONCAT_ ################################




# Use Get-ChildItem to iterate through all FT_Chambre.shp files recursively in the specified directory
Get-ChildItem -Path $PATHWAY -Recurse -Filter "FT_Chambre.shp" | ForEach-Object {
    # Run the ogrinfo command to drop the CONCAT_CH column
    ogrinfo -dialect OGRSQL -sql "ALTER TABLE FT_Chambre DROP COLUMN CONCAT_CH" $_.FullName
}




# Use Get-ChildItem to iterate through all FT_Chambre.shp files recursively in the specified directory
Get-ChildItem -Path $PATHWAY -Recurse -Filter "FT_Appui.shp" | ForEach-Object {
    # Run the ogrinfo command to drop the CONCAT_CH column
    ogrinfo -dialect OGRSQL -sql "ALTER TABLE FT_Appui DROP COLUMN CONCAT_AP" $_.FullName
}



################################ Rezippage de tout les shapefiles ###########################################
#
#
#
## Get a list of all directories to zip
#$dir = Get-ChildItem -Path $PATHWAY -Recurse | Where-Object { $_.PSIsContainer }
#
## Check if directories were found
#if ($dir.Count -gt 0) {
#    # Zip each directory
#    foreach ($directory in $dir) {
#        $zipFilePath = Join-Path -Path $directory.Parent.FullName -ChildPath "$($directory.Name).zip"
#        Compress-Archive -Path $directory.FullName -DestinationPath $zipFilePath -Force
#    }
#} else {
#    Write-Host "No directories found to zip in the specified path: $PATHWAY_NRO"
#}
#

# Obtenir tous les dossiers et sous-dossiers
$allSubDirs = Get-ChildItem -Path $PATHWAY -Recurse | Where-Object { $_.PSIsContainer }

# Compresser chaque sous-dossier dans un fichier ZIP séparé
foreach ($subDir in $allSubDirs) {
    # Nom du fichier ZIP basé sur le nom du sous-dossier
    $zipFileName = "$($subDir.Name).zip"
    
    # Chemin de destination pour le fichier ZIP, basé sur le répertoire parent du sous-dossier
    $zipDestinationPath = Join-Path -Path $subDir.Parent.FullName -ChildPath $zipFileName
    
    # Obtenir tous les éléments du sous-dossier (fichiers et sous-dossiers directs)
    $itemsToCompress = Get-ChildItem -Path $subDir.FullName | Where-Object { $_.PSIsContainer -or -not $_.PSIsContainer }
    
    # Compresser les éléments sans inclure le dossier parent
    Compress-Archive -Path $itemsToCompress.FullName -DestinationPath $zipDestinationPath -Force
}

Write-Host "Tous les dossiers et sous-dossiers ont été compressés avec succès."

	Pause
