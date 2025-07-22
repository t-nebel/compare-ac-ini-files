<#
.SYNOPSIS
    Compares two AudioCodes SBC INI configuration files and generates an HTML report of differences.

.DESCRIPTION
    This script parses and analyzes two AudioCodes SBC INI configuration files, comparing their sections,
    parameters, and table entries. It identifies differences between the files and generates a detailed
    HTML report highlighting all discrepancies.

.PARAMETER SBCIniFilePath1
    The path to the first SBC INI file to compare.

.PARAMETER SBCName1
    The name to use for the first SBC in the report. Defaults to "SBC1".

.PARAMETER SBCIniFilePath2
    The path to the second SBC INI file to compare.

.PARAMETER SBCName2
    The name to use for the second SBC in the report. Defaults to "SBC2".

.PARAMETER ReportFilePath
    The path where the HTML report will be saved. Defaults to "SBC-Compare-Report.html" in the current directory.

.EXAMPLE
    .\Compare-ACSBCIniFiles.ps1 -SBCIniFilePath1 "C:\Temp\SBC1.ini" -SBCName1 "PROD-SBC" -SBCIniFilePath2 "C:\Temp\SBC2.ini" -SBCName2 "DR-SBC" -ReportFilePath "C:\Temp\SBC-Compare-Report.html"

    Compares the INI files from PROD-SBC and DR-SBC and generates an HTML report of differences. The report will be saved to "C:\Temp\SBC-Compare-Report.html".
#>

param (
    [Parameter(Mandatory = $true)]
    $SBCIniFilePath1,
    [Parameter(Mandatory = $true)]
    $SBCName1 = "SBC1",
    [Parameter(Mandatory = $true)]
    $SBCIniFilePath2,
    [Parameter(Mandatory = $true)]
    $SBCName2 = "SBC2",
    [Parameter(Mandatory = $false)]
    $reportFilePath = (Join-Path -Path ((Get-Location).Path) -ChildPath "SBC-Compare-Report.html")
)
########################################################################
#region Functions
########################################################################

# Function to parse an INI file into an array of PSCustomObjects (one object per section)
function ConvertFrom-IniData {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string]$FilePath
    )

    if (-not (Test-Path $FilePath)) {
        Write-Error "File not found: $FilePath"
        return $null
    }

    $iniData = [System.Collections.ArrayList]@()
    $currentSection = $null
    $currentSectionName = $null
    $currentSectionIsTable = $false
    $tableHeaders = @()
    $tableRows = @()

    Get-Content $FilePath | ForEach-Object {
        $line = $_.Trim()

        # Section header: [SectionName]
        if ($line -match '^\[\s*([^\\\]]+)\s*\]$') {
            # If previous section was a table, save its rows
            if ($currentSectionIsTable -and $currentSectionName) {
                $currentSection.Content = $tableRows
                $iniData.Add($currentSection) | Out-Null
                $tableRows = @()
                $currentSectionIsTable = $false
            }
            # If there was a previous non-table section, add it to the result
            elseif ($currentSection -ne $null) {
                $iniData.Add($currentSection) | Out-Null
            }

            $currentSectionName = $matches[1].Trim()
            $currentSection = [PSCustomObject]@{
                Name    = $currentSectionName
                Content = [ordered]@{}
                Type    = "KeyValue"
            }
            $tableHeaders = @()
        }
        # Section closing tag: [ \SectionName ]
        elseif ($line -match '^\[\s*\\([^\]]+)\s*\]$') {
            if ($currentSectionIsTable -and $currentSectionName) {
                $currentSection.Content = $tableRows
                $currentSection.Type = "Table"
                $iniData.Add($currentSection) | Out-Null
                $tableRows = @()
                $currentSectionIsTable = $false
            }
            elseif ($currentSection -ne $null) {
                $iniData.Add($currentSection) | Out-Null
            }
            $currentSection = $null
            $currentSectionName = $null
        }
        # FORMAT line (table header)
        elseif ($currentSectionName -and $line -match '^FORMAT\s+Index\s*=\s*(.*)$') {
            $tableHeaders = $matches[1].Split(',') | ForEach-Object { $_.Trim() }
            $tableHeaders = @("Index") + $tableHeaders  # Add Index as the first column
            $tableRows = @()
            $currentSectionIsTable = $true
            $currentSection.Type = "Table"
        }
        # Table row
        elseif ($currentSectionIsTable -and $line -match '^([^;#]\S+)\s+(\d+)\s*=\s*(.*)$') {
            $tableName = $matches[1]
            $rowIndex = $matches[2]
            $rowData = $matches[3]

            $rowValues = $rowData.Split(',') | ForEach-Object { $_.Trim() }
            $rowObject = [ordered]@{}

            # Add the index as the first column
            $rowObject["Index"] = $rowIndex

            # Add the rest of the values
            for ($i = 0; $i -lt $tableHeaders.Count - 1; $i++) {
                # -1 because we already added Index
                if ($i -lt $rowValues.Count) {
                    $rowObject[$tableHeaders[$i + 1]] = $rowValues[$i]  # +1 to skip the Index column we added
                }
            }
            $tableRows += $rowObject
        }
        # Key-value pair
        elseif ($currentSectionName -and -not $currentSectionIsTable -and $line -match '^\s*([^#;].*?)\s*=\s*(.*)') {
            $key = $matches[1].Trim()
            $value = $matches[2].Trim()
            $currentSection.Content[$key] = $value
        }
    }
    # If file ends while a section is still open, save it
    if ($currentSection -ne $null) {
        if ($currentSectionIsTable) {
            $currentSection.Content = $tableRows
            $currentSection.Type = "Table"
        }
        $iniData.Add($currentSection) | Out-Null
    }
    return $iniData
}

#endregion

########################################################################
#region Main script execution
########################################################################

#region 1. Parse both INI files into memory
###########################################
Write-Output "Parsing INI files..."
$ini1 = ConvertFrom-IniData -FilePath $SBCIniFilePath1
$ini2 = ConvertFrom-IniData -FilePath $SBCIniFilePath2

if (-not $ini1 -or -not $ini2) {
    Write-Error "Could not parse one or both INI files. Aborting." -ErrorAction Stop
}
#endregion

#region Initialize differences array
$differencesBySection = [ordered]@{}
#endregion

#region 2. Compare sections
###########################
Write-Output "`n--- Comparing Sections ---"
$sectionsInSBC1 = $ini1 | ForEach-Object { $_.Name }
$sectionsInSBC2 = $ini2 | ForEach-Object { $_.Name }
$sectionComparison = Compare-Object -ReferenceObject $sectionsInSBC1 -DifferenceObject $sectionsInSBC2 -IncludeEqual
$sectionDifferenceExistsSBC1 = $false
$sectionDifferenceExistsSBC2 = $false
$sectionsOnlyInSBC1 = @()
$sectionsOnlyInSBC2 = @()
$sectionDifferences = [System.Collections.ArrayList]@()

foreach ($diff in $sectionComparison) {
    if ($diff.SideIndicator -eq '<=') {
        $sectionDifferenceExistsSBC1 = $true
        $sectionsOnlyInSBC1 += $diff.InputObject
    }
    elseif ($diff.SideIndicator -eq '=>') {
        $sectionDifferenceExistsSBC2 = $true
        $sectionsOnlyInSBC2 += $diff.InputObject
    }
}

if ($sectionDifferenceExistsSBC1 -or $sectionDifferenceExistsSBC2) {
    Write-Output "Sections would be found, which are not in both INI files."
    if ($sectionDifferenceExistsSBC1) {
        Write-Output "Sections which are only in $($SBCName1):"
        foreach ($currentItemName in $sectionsOnlyInSBC1) {
            Write-Output " - $($currentItemName)"
            $sectionDifferences.Add("Section '<b>$($currentItemName)</b>' only exists in <b>$SBCName1</b>") | Out-Null
        }
    }
    if ($sectionDifferenceExistsSBC2) {
        Write-Output "Sections which are only in $($SBCName2):"
        foreach ($currentItemName in $sectionsOnlyInSBC2) {
            Write-Output " - $($currentItemName)"
            $sectionDifferences.Add("Section '<b>$($currentItemName)</b>' only exists in <b>$SBCName2</b>") | Out-Null
        }
    }
    if ($sectionDifferences.Count -gt 0) {
        $differencesBySection["General Section Differences"] = $sectionDifferences
    }
}
else {
    Write-Output "The same sections are in both INI files."
}

if ($sectionsOnlyInSBC1) {
    Clear-Variable sectionsOnlyInSBC1
}
if ($sectionsOnlyInSBC2) {
    Clear-Variable sectionsOnlyInSBC2
}
#endregion

#region 3. Compare content of  sections
#############################################
Write-Output "`n--- Comparing Content of  Sections ---"
$commonSections = $sectionComparison | Where-Object { $_.SideIndicator -eq '==' } | ForEach-Object { $_.InputObject }

foreach ($sectionName in $commonSections) {
    Write-Output "`n----- Section: [$sectionName] -----"
    $section1 = $ini1 | Where-Object { $_.Name -eq $sectionName }
    $section2 = $ini2 | Where-Object { $_.Name -eq $sectionName }
    $content1 = $section1.Content
    $content2 = $section2.Content
    $sectionDifferences = [System.Collections.ArrayList]@()

    # Check if the section is a table or key-value pairs based on its Type property
    if ($section1.Type -eq "Table") {
        # Table comparison with Index matching
        $rowMissmatch = $false
        $hasTableDiff = $false
        if ($content1.Count -eq 0 -and $content2.Count -eq 0) {
            Write-Host "No differences found (both tables are empty)." -ForegroundColor Green
            continue
        }
        if ($content1.Count -gt 0 -and $content2.Count -eq 0) {
            $message = "Section table in <b>$($SBCName1)</b> has <b>$($content1.Count)</b> rows, but table in <b>$($SBCName2)</b> is empty."
            Write-Warning "Section table in $($SBCName1) has $($content1.Count) rows, but table in $($SBCName2) is empty."
            $sectionDifferences.Add($message) | Out-Null
            $differencesBySection[$sectionName] = $sectionDifferences
            continue
        }
        if ($content2.Count -gt 0 -and $content1.Count -eq 0) {
            $message = "Section table in <b>$($SBCName2)</b> has <b>$($content2.Count)</b> rows, but table in <b>$($SBCName1)</b> is empty."
            Write-Warning "Section table in $($SBCName2) has $($content2.Count) rows, but table in $($SBCName1) is empty."
            $sectionDifferences.Add($message) | Out-Null
            $differencesBySection[$sectionName] = $sectionDifferences
            continue
        }
        if ($content1.Count -gt $content2.Count) {
            $message = "Section table in <b>$($SBCName1)</b> has more rows (<b>$($content1.Count)</b>) than table in <b>$($SBCName2)</b> (<b>$($content2.Count)</b>)."
            Write-Warning "Section table in $($SBCName1) has more rows ($($content1.Count)) than table in $($SBCName2) ($($content2.Count))."
            $sectionDifferences.Add($message) | Out-Null
            $hasTableDiff = $true
            $rowMissmatch = $true
        }
        elseif ($content2.Count -gt $content1.Count) {
            $message = "Section table in <b>$($SBCName2)</b> has more rows (<b>$($content2.Count)</b>) than table in <b>$($SBCName1)</b> (<b>$($content1.Count)</b>)."
            Write-Warning "Section table in $($SBCName2) has more rows ($($content2.Count)) than table in $($SBCName1) ($($content1.Count))."
            $sectionDifferences.Add($message) | Out-Null
            $hasTableDiff = $true
            $rowMissmatch = $true
        }

        $currentProperties = $content1.GetEnumerator().Keys
        $currentSectionTableName = $currentProperties | Where-Object { $_ -ne 'Index' } | Select-Object -First 1
        if ($currentSectionTableName -notlike $($content2.GetEnumerator().Keys | Where-Object { $_ -ne 'Index' } | Select-Object -First 1)) {
            $message = "Section table - displayname property for the index '<b>$($currentSectionTableName)</b>' does not match in both INI files."
            Write-Warning "Section table - displayname property for the index '$($currentSectionTableName)' does not match in both INI files."
            $sectionDifferences.Add($message) | Out-Null
            $hasTableDiff = $true
        }

        # Compare Objects of each with each other, even if one has more. $Content1[0] with $Content2[0], $Content1[1] with $Content2[1], etc.
        if ($content1.Count -ne $content2.Count) {
            $hasTableDiff = $true
            # Based on the index number of the rows, compare if the properties are the same.
            # Get all Index values from both tables
            $indexValues1 = $content1 | ForEach-Object { $_.Index }
            $indexValues2 = $content2 | ForEach-Object { $_.Index }

            # If there were missmatching indexes, output them with the current section name
            $missmatchedIndexes1 = $indexValues1 | Where-Object { -not ($indexValues2 -contains $_) }
            if ($missmatchedIndexes1) {
                foreach ($index in $missmatchedIndexes1) {
                    $message = "Section row with Index '<b>$($index)</b>', $($currentSectionTableName) '<b>$((($content1 | Where-Object Index -like $index ).$currentSectionTableName))</b>' only exists in <b>$SBCName1</b>"
                    Write-Warning "Row with Index: '$($index)' and $($currentSectionTableName): '$((($content1 | Where-Object Index -like $index ).$currentSectionTableName))' only exists in $($SBCName1)"
                    $sectionDifferences.Add($message) | Out-Null
                    $hasTableDiff = $true
                }
            }
            $missmatchedIndexes2 = $indexValues2 | Where-Object { -not ($indexValues1 -contains $_) }
            if ($missmatchedIndexes2) {
                foreach ($index in $missmatchedIndexes2) {
                    $message = "Section row with Index '<b>$($index)</b>', $($currentSectionTableName) '<b>$((($content2 | Where-Object Index -like $index ).$currentSectionTableName))</b>' only exists in <b>$SBCName2</b>"
                    Write-Warning "Row with Index: '$($index)' and $($currentSectionTableName): '$((($content2 | Where-Object Index -like $index ).$currentSectionTableName))' only exists in $($SBCName2)"
                    $sectionDifferences.Add($message) | Out-Null
                    $hasTableDiff = $true
                }
            }

            # Based on the matching Index values, compare the properties
            $matchedIndices = $indexValues1 | Where-Object { $indexValues2 -contains $_ }
            foreach ($index in $matchedIndices) {
                $row1 = $content1 | Where-Object { $_.Index -eq $index }
                $row2 = $content2 | Where-Object { $_.Index -eq $index }
                $rowDifferences = [System.Collections.ArrayList]@()
                $processedProperties = @{}

                foreach ($property in $currentProperties) {
                    # Only process this property if it hasn't been processed before
                    if ($row1.$property -ne $row2.$property -and -not $processedProperties.ContainsKey($property)) {
                        $rowDifferences.Add(@{Property = $property; Value1 = $row1.$property; Value2 = $row2.$property }) | Out-Null
                        $hasTableDiff = $true
                        # Mark this property as processed so we don't add it multiple times
                        $processedProperties[$property] = $true
                    }
                }

                if ($rowDifferences.Count -gt 0) {
                    $htmlTable = "Section row with Index '<b>$($index)</b>', $($currentSectionTableName) '<b>$($row1.$($CurrentSectionTableName))</b>' has differences:<br>"
                    Write-Warning "Section row with Index: '$($index)' and $($currentSectionTableName): '$($row1.$($CurrentSectionTableName))' has differences:"
                    $htmlTable += "<table class='diff-table'><tr><th>Property</th><th>$($SBCName1)</th><th>$($SBCName2)</th></tr>"
                    foreach ($diff in $rowDifferences) {
                        $htmlTable += "<tr><td>$($diff.Property)</td><td>$($diff.Value1)</td><td>$($diff.Value2)</td></tr>"
                        Write-Warning "$($diff.Property): $($diff.Value1) (in $($SBCName1)) vs $($diff.Value2) (in $($SBCName2))"
                    }
                    $htmlTable += "</table>"
                    $sectionDifferences.Add($htmlTable) | Out-Null
                }
            }


        }
        else {
            $amountOfRows = $content1.Count
            $counter = 0
            while ($amountOfRows -ne $counter) {
                $rowDifferences = [System.Collections.ArrayList]@()
                $processedProperties = @{}

                foreach ($property in $currentProperties) {
                    # Only process this property if it hasn't been processed before
                    if ($content1[$counter].$property -ne $content2[$counter].$property -and -not $processedProperties.ContainsKey($property)) {
                        $rowDifferences.Add(@{Property = $property; Value1 = $content1[$counter].$property; Value2 = $content2[$counter].$property }) | Out-Null
                        $hasTableDiff = $true
                        # Mark this property as processed
                        $processedProperties[$property] = $true
                    }
                }

                if ($rowDifferences.Count -gt 0) {
                    $htmlTable = "Section table '<b>$($sectionName)</b>', row with Index '<b>$($content1[$counter].Index)</b>', $($currentSectionTableName) '<b>$($content1[$counter].$($currentSectionTableName))</b>' has differences:<br>"
                    Write-Warning "Section table '$($sectionName)', row with Index: '$($content1[$counter].Index)', $($currentSectionTableName): '$($content1[$counter].$($currentSectionTableName))' has differences:"
                    $htmlTable += "<table class='diff-table'><tr><th>Property</th><th>$($SBCName1)</th><th>$($SBCName2)</th></tr>"
                    foreach ($diff in $rowDifferences) {
                        $htmlTable += "<tr><td>$($diff.Property)</td><td>$($diff.Value1)</td><td>$($diff.Value2)</td></tr>"
                        Write-Warning "$($diff.Property): $($diff.Value1) (in $($SBCName1)) vs $($diff.Value2) (in $($SBCName2))"
                    }
                    $htmlTable += "</table>"
                    $sectionDifferences.Add($htmlTable) | Out-Null
                }
                $counter++
            }
        }
    }
    else {
        # Key-value comparison
        $keyComparison = Compare-Object -ReferenceObject $content1.Keys -DifferenceObject $content2.Keys -IncludeEqual
        $hasDifferences = $false

        # Check for keys only present in one or the other
        foreach ($diff in $keyComparison | Where-Object { $_.SideIndicator -ne '==' }) {
            $hasDifferences = $true
            if ($diff.SideIndicator -eq '<=') {
                $message = "Parameter '<b>$($diff.InputObject)</b>' only exists in <b>$($SBCName1)</b>"
                Write-Warning "Parameter '$($diff.InputObject)' only exists in $($SBCName1)"
                $sectionDifferences.Add($message) | Out-Null
            }
            else {
                $message = "Parameter '<b>$($diff.InputObject)</b>' only exists in <b>$($SBCName2)</b>"
                Write-Warning "Parameter '$($diff.InputObject)' only exists in $($SBCName2)"
                $sectionDifferences.Add($message) | Out-Null
            }
        }

        # Check for different values on  keys
        $Keys = $keyComparison | Where-Object { $_.SideIndicator -eq '==' } | ForEach-Object { $_.InputObject }
        foreach ($key in $Keys) {
            if ($content1.$key -ne $content2.$key) {
                $hasDifferences = $true
                $message = "Value for parameter '<b>$($key)</b>' differs:<br>&nbsp;&nbsp;<b>$($SBCName1)</b>: $($content1.$key)<br>&nbsp;&nbsp;<b>$($SBCName2)</b>: $($content2.$key)"
                Write-Warning "Value for parameter '$($key)' differs:"
                Write-Host "  $($SBCName1)`: $($content1.$key)"
                Write-Host "  $($SBCName2)`: $($content2.$key)"
                $sectionDifferences.Add($message) | Out-Null
            }
        }
    }

    if ($sectionDifferences.Count -gt 0) {
        $differencesBySection[$sectionName] = $sectionDifferences
    }
    elseif ((-not $hasTableDiff -and -not $rowMissmatch) -or -not $hasDifferences) {
        Write-Host "No differences found" -ForegroundColor Green
    }
}

#endregion

#region 4. Generate HTML Report or output no differences
#######################################################
if ($differencesBySection.Count -gt 0) {
    Write-Output "`n--- Generating HTML Report ---"

    $htmlHead = @"
<!DOCTYPE html>
<html>
<head>
<title>SBC INI File Comparison Report</title>
<link href="https://fonts.googleapis.com/css2?family=Lato:wght@400;700&display=swap" rel="stylesheet">
<style>
  body {
    font-family: 'Lato', sans-serif;
    background-color: #f8f9fa;
    color: #343a40;
    margin: 20px;
  }
  h1, h2, h3 {
    color: #0056b3;
  }
  h1 {
    border-bottom: 2px solid #0056b3;
    padding-bottom: 10px;
  }
  h2 {
    color: #5a6268;
    border-bottom: 1px solid #dee2e6;
    padding-bottom: 5px;
  }
  h3 {
    background-color: #e9ecef;
    padding: 12px;
    color: #495057;
    margin-top: 30px;
    border-left: 4px solid #007bff;
  }
  .report {
    border-collapse: collapse;
    width: 100%;
    margin-top: 10px;
    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
  }
  .report th, .report td {
    border: 1px solid #dee2e6;
    padding: 12px;
    text-align: left;
  }
  .report th {
    background-color: #e9ecef;
    font-weight: 700;
  }
  .difference td {
    /* No specific color, inherit from body */
  }
  b {
    color: #004085;
    font-weight: 700;
  }
  .diff-table {
    border-collapse: collapse;
    width: 95%;
    margin-top: 10px;
    margin-left: auto;
    margin-right: auto;
    border: 1px solid #c8cbcf;
  }
  .diff-table th, .diff-table td {
    border: 1px solid #dee2e6;
    padding: 8px;
    text-align: left;
    background-color: #ffffff;
  }
  .diff-table th {
    background-color: #f2f2f2;
  }
</style>
</head>
<body>
<h1>SBC INI File Comparison Report</h1>
<h2>Comparison between '$($SBCName1)' and '$($SBCName2)'</h2>
"@

    $htmlBody = ""
    foreach ($section in $differencesBySection.Keys) {
        $htmlBody += "<h3>$($section)</h3>"

        # Group differences by category within each section
        $categorizedDifferences = @{}

        # Process each difference in the section
        foreach ($item in $differencesBySection[$section]) {
            # Determine the category of the difference
            $category = "Difference Description"

            if ($section -eq "General Section Differences") {
                $category = "Section Mismatch"
            }
            elseif ($item -like "*only exists in*" -and $item -like "*Parameter*") {
                $category = "Missing Parameters"
            }
            elseif ($item -like "*Value for parameter*differs*") {
                $category = "Parameter Value Mismatch"
            }
            elseif ($item -like "*row with Index*only exists*") {
                $category = "Missing Table Rows"
            }
            elseif ($item -like "*has differences*") {
                $category = "Different Row Values"
            }
            elseif ($item -like "*has more rows*") {
                $category = "Row Count Mismatch"
            }
            elseif ($item -like "*is empty*") {
                $category = "Empty Table"
            }

            # Add to the appropriate category
            if (-not $categorizedDifferences.Contains($category)) {
                $categorizedDifferences[$category] = [System.Collections.ArrayList]@()
            }
            $categorizedDifferences[$category].Add($item) | Out-Null
        }

        # Track the order of categories for consistent output
        $categoryOrder = @(
            "Section Mismatch",
            "Missing Parameters",
            "Parameter Value Mismatch",
            "Row Count Mismatch",
            "Empty Table",
            "Missing Table Rows",
            "Different Row Values",
            "Difference Description"
        )

        # Create a table for each category in the section, maintaining a logical order
        foreach ($cat in $categoryOrder) {
            if ($categorizedDifferences.Contains($cat)) {
                $htmlBody += "<table class='report'><tr><th>$cat</th></tr>"
                foreach ($item in $categorizedDifferences[$cat]) {
                    $htmlBody += "<tr class='difference'><td>$($item)</td></tr>"
                }
                $htmlBody += "</table>"
                $htmlBody += "<br>" # Add space between tables
            }
        }
    }

    $htmlFoot = @"
</body>
</html>
"@

    # Validate the report file path directory and add warning if it does not exist
    $reportDirectory = Split-Path -Path $ReportFilePath -Parent
    if (-not (Test-Path -Path $reportDirectory)) {
        Write-Warning "The directory '$reportDirectory' does not exist. The report will not be saved!"
        Exit 1
    }
    else {
        $htmlContent = $htmlHead + $htmlBody + $htmlFoot
        $htmlContent | Out-File -FilePath $ReportFilePath -Encoding UTF8

        Write-Host "HTML report generated at: $ReportFilePath" -ForegroundColor Cyan
    }

}
else {
    Write-Host "`n--- No differences found between the two INI files. ---" -ForegroundColor Green
}
#endregion

#endregion