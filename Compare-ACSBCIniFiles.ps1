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

.PARAMETER IncludeNameValidation
    When set to $true, table entries with the same Index will also be compared by their name (first column after Index).
    This enables secondary comparison based on the ProfileName for tables like IpProfile. Defaults to $false.

.EXAMPLE
    .\Compare-ACSBCIniFiles.ps1 -SBCIniFilePath1 "C:\Temp\SBC1.ini" -SBCName1 "PROD-SBC" -SBCIniFilePath2 "C:\Temp\SBC2.ini" -SBCName2 "DR-SBC" -ReportFilePath "C:\Temp\SBC-Compare-Report.html"

    Compares the INI files from PROD-SBC and DR-SBC and generates an HTML report of differences. The report will be saved to "C:\Temp\SBC-Compare-Report.html".

.EXAMPLE
    .\Compare-ACSBCIniFiles.ps1 -SBCIniFilePath1 "C:\Temp\Original.ini" -SBCIniFilePath2 "C:\Temp\Abweichung.ini" -IncludeNameValidation $true

    Compares the INI files with name validation enabled, which will also compare table entries by their name (ProfileName) in addition to Index.
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
    $ReportFilePath = (Join-Path -Path ((Get-Location).Path) -ChildPath "SBC-Compare-Report.html"),
    [Parameter(Mandatory = $false)]
    [bool]$IncludeNameValidation = $false
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
if ($IncludeNameValidation) {
    Write-Host "ℹ️ Name validation is ENABLED - Table entries will also be compared by name if indices differ" -ForegroundColor Cyan
} else {
    Write-Host "ℹ️ Name validation is DISABLED - Table entries will only be compared by index" -ForegroundColor Gray
}
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
                    # If IncludeNameValidation is enabled, check if this entry can be matched by name before reporting as missing
                    $skipMissingReport = $false
                    if ($IncludeNameValidation) {
                        $entryName = ($content1 | Where-Object { $_.Index -eq $index }).$currentSectionTableName
                        $nameValues2 = $content2 | ForEach-Object { $_.$currentSectionTableName }
                        if ($nameValues2 -contains $entryName) {
                            $skipMissingReport = $true
                        }
                    }

                    if (-not $skipMissingReport) {
                        $message = "Section row with Index '<b>$($index)</b>', $($currentSectionTableName) '<b>$((($content1 | Where-Object Index -like $index ).$currentSectionTableName))</b>' only exists in <b>$SBCName1</b>"
                        Write-Warning "Row with Index: '$($index)' and $($currentSectionTableName): '$((($content1 | Where-Object Index -like $index ).$currentSectionTableName))' only exists in $($SBCName1)"
                        $sectionDifferences.Add($message) | Out-Null
                        $hasTableDiff = $true
                    }
                }
            }
            $missmatchedIndexes2 = $indexValues2 | Where-Object { -not ($indexValues1 -contains $_) }
            if ($missmatchedIndexes2) {
                foreach ($index in $missmatchedIndexes2) {
                    # If IncludeNameValidation is enabled, check if this entry can be matched by name before reporting as missing
                    $skipMissingReport = $false
                    if ($IncludeNameValidation) {
                        $entryName = ($content2 | Where-Object { $_.Index -eq $index }).$currentSectionTableName
                        $nameValues1 = $content1 | ForEach-Object { $_.$currentSectionTableName }
                        if ($nameValues1 -contains $entryName) {
                            $skipMissingReport = $true
                        }
                    }

                    if (-not $skipMissingReport) {
                        $message = "Section row with Index '<b>$($index)</b>', $($currentSectionTableName) '<b>$((($content2 | Where-Object Index -like $index ).$currentSectionTableName))</b>' only exists in <b>$SBCName2</b>"
                        Write-Warning "Row with Index: '$($index)' and $($currentSectionTableName): '$((($content2 | Where-Object Index -like $index ).$currentSectionTableName))' only exists in $($SBCName2)"
                        $sectionDifferences.Add($message) | Out-Null
                        $hasTableDiff = $true
                    }
                }
            }

            # Based on the matching Index values, compare the properties
            $matchedIndices = $indexValues1 | Where-Object { $indexValues2 -contains $_ }
            foreach ($index in $matchedIndices) {
                $row1 = $content1 | Where-Object { $_.Index -eq $index }
                $row2 = $content2 | Where-Object { $_.Index -eq $index }

                # If IncludeNameValidation is enabled and the name property differs, look for a matching name in the other table
                if ($IncludeNameValidation -and $row1.$currentSectionTableName -ne $row2.$currentSectionTableName) {
                    # Look for a row with the same name in the other table
                    $matchingRowsByName = $content2 | Where-Object { $_.$currentSectionTableName -eq $row1.$currentSectionTableName }

                    if ($matchingRowsByName) {
                        # Check if there are multiple entries with the same name (duplicates)
                        if ($matchingRowsByName -is [array] -and $matchingRowsByName.Count -gt 1) {
                            # Multiple entries with same name found - use fallback to index-based comparison
                            Write-Host "⚠️ Multiple entries named '$($row1.$currentSectionTableName)' found - using fallback to index-based comparison" -ForegroundColor Yellow

                            # Add HTML message about duplicate fallback
                            $duplicateMessage = "<span style='color: #856404; font-weight: bold;'><strong>Duplicate Name Detected:</strong></span> Multiple entries named '<b>$($row1.$currentSectionTableName)</b>' found in target file. Using index-based comparison fallback for Index <b>$($index)</b>."
                            $sectionDifferences.Add($duplicateMessage) | Out-Null
                            $hasTableDiff = $true
                            # Don't change $row2, keep the original index-based comparison
                        } else {
                            # Single matching entry found - use name-based matching
                            $matchingRowByName = if ($matchingRowsByName -is [array]) { $matchingRowsByName[0] } else { $matchingRowsByName }
                            Write-Host "✓ Name-based matching activated: '$($row1.$currentSectionTableName)' found at different indices (Index $($index) in $($SBCName1) <-> Index $($matchingRowByName.Index) in $($SBCName2))" -ForegroundColor Cyan
                            $row2 = $matchingRowByName

                            # Add a note about the index difference with success indication
                            $message = "<span style='color: #28a745;'><strong>Name-based match found:</strong></span> Entry '<b>$($row1.$currentSectionTableName)</b>' exists in both files but with different Index values: <b>$($index)</b> in $($SBCName1) vs <b>$($matchingRowByName.Index)</b> in $($SBCName2)"
                            Write-Host "✓ Successfully matched entry '$($row1.$currentSectionTableName)' despite index difference: $($index) in $($SBCName1) vs $($matchingRowByName.Index) in $($SBCName2)" -ForegroundColor Green
                            $sectionDifferences.Add($message) | Out-Null
                            $hasTableDiff = $true
                        }
                    }
                }                $rowDifferences = [System.Collections.ArrayList]@()
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

            # If IncludeNameValidation is enabled, also check for entries that have the same name but different indices
            if ($IncludeNameValidation) {
                # Get all name values from both tables
                $nameValues1 = $content1 | ForEach-Object { $_.$currentSectionTableName }
                $nameValues2 = $content2 | ForEach-Object { $_.$currentSectionTableName }

                # Find names that exist in both tables but haven't been compared yet (due to different indices)
                $commonNames = $nameValues1 | Where-Object { $nameValues2 -contains $_ }

                foreach ($name in $commonNames) {
                    $rows1 = $content1 | Where-Object { $_.$currentSectionTableName -eq $name }
                    $rows2 = $content2 | Where-Object { $_.$currentSectionTableName -eq $name }

                    # Check for duplicate names - if found, skip name-based comparison for this name
                    if (($rows1 -is [array] -and $rows1.Count -gt 1) -or ($rows2 -is [array] -and $rows2.Count -gt 1)) {
                        Write-Host "⚠️ Duplicate name '$($name)' found in one or both files - skipping name-based comparison" -ForegroundColor Yellow
                        continue
                    }

                    # Skip if we already compared these based on index matching
                    $alreadyCompared = $false
                    foreach ($tempRow1 in $rows1) {
                        foreach ($tempRow2 in $rows2) {
                            if ($tempRow1.Index -eq $tempRow2.Index) {
                                $alreadyCompared = $true
                                break
                            }
                        }
                        if ($alreadyCompared) { break }
                    }

                    if (-not $alreadyCompared) {
                        # Compare the first occurrence of each name
                        $row1 = $rows1[0]
                        $row2 = $rows2[0]

                        # Check if we got actual objects or just strings (indicates a logic error elsewhere)
                        if ($row1 -is [string] -or $row2 -is [string]) {
                            # Skip this comparison as it's invalid - the filtering logic returned strings instead of objects
                            continue
                        }

                        $rowDifferences = [System.Collections.ArrayList]@()
                        $processedProperties = @{}

                        foreach ($property in $currentProperties) {
                            # Skip the name property itself and Index as we know they're different
                            if ($property -ne $currentSectionTableName -and $property -ne "Index") {
                                if ($row1.$property -ne $row2.$property -and -not $processedProperties.ContainsKey($property)) {
                                    $rowDifferences.Add(@{Property = $property; Value1 = $row1.$property; Value2 = $row2.$property }) | Out-Null
                                    $hasTableDiff = $true
                                    $processedProperties[$property] = $true
                                }
                            }
                        }

                        if ($rowDifferences.Count -gt 0) {
                            $htmlTable = "Entry with name '<b>$($name)</b>' (Index $($row1.Index) in $($SBCName1), Index $($row2.Index) in $($SBCName2)) has differences:<br>"
                            Write-Warning "Entry with name '$($name)' (Index $($row1.Index) in $($SBCName1), Index $($row2.Index) in $($SBCName2)) has differences:"
                            $htmlTable += "<table class='diff-table'><tr><th>Property</th><th>$($SBCName1)</th><th>$($SBCName2)</th></tr>"
                            foreach ($diff in $rowDifferences) {
                                $htmlTable += "<tr><td>$($diff.Property)</td><td>$($diff.Value1)</td><td>$($diff.Value2)</td></tr>"
                                Write-Warning "$($diff.Property): $($diff.Value1) (in $($SBCName1)) vs $($diff.Value2) (in $($SBCName2))"
                            }
                            $htmlTable += "</table>"
                            $sectionDifferences.Add($htmlTable) | Out-Null
                        }
                        else {
                            # No differences found despite different indices - this is a successful name-based match
                            # Capture index values immediately to avoid variable scope issues
                            $index1 = $row1.Index
                            $index2 = $row2.Index
                            $message = "<span style='color: #28a745;'><strong>Perfect name-based match:</strong></span> Entry '<b>$($name)</b>' is identical in both files despite different Index values: <b>$($index1)</b> in $($SBCName1) vs <b>$($index2)</b> in $($SBCName2)"
                            Write-Host "✓ Perfect match found for '$($name)' despite index difference: $($index1) in $($SBCName1) vs $($index2) in $($SBCName2)" -ForegroundColor Green
                            $sectionDifferences.Add($message) | Out-Null
                            $hasTableDiff = $true
                        }
                    }
                }
            }
        }
        else {
            $amountOfRows = $content1.Count
            $counter = 0
            while ($amountOfRows -ne $counter) {
                $row1 = $content1[$counter]
                $row2 = $content2[$counter]

                # If IncludeNameValidation is enabled and the name property differs, look for a matching name in the other table
                if ($IncludeNameValidation -and $row1.$currentSectionTableName -ne $row2.$currentSectionTableName) {
                    # Look for a row with the same name in the other table
                    $matchingRowsByName = $content2 | Where-Object { $_.$currentSectionTableName -eq $row1.$currentSectionTableName }

                    if ($matchingRowsByName) {
                        # Check if there are multiple entries with the same name (duplicates)
                        if ($matchingRowsByName -is [array] -and $matchingRowsByName.Count -gt 1) {
                            # Multiple entries with same name found - use fallback to index-based comparison
                            Write-Host "[WARNING] Multiple entries named '$($row1.$currentSectionTableName)' found - using fallback to index-based comparison" -ForegroundColor Yellow

                            # Add HTML message about duplicate fallback
                            $duplicateMessage = "<span style='color: #856404; font-weight: bold;'><strong>Duplicate Name Detected:</strong></span> Multiple entries named '<b>$($row1.$currentSectionTableName)</b>' found in target file. Using index-based comparison fallback for Index <b>$($row1.Index)</b>."
                            $sectionDifferences.Add($duplicateMessage) | Out-Null
                            $hasTableDiff = $true
                            # Don't change $row2, keep the original index-based comparison
                        } else {
                            # Single matching entry found - use name-based matching
                            $matchingRowByName = if ($matchingRowsByName -is [array]) { $matchingRowsByName[0] } else { $matchingRowsByName }
                            Write-Host "✓ Name-based matching activated: '$($row1.$currentSectionTableName)' found at different indices (Index $($row1.Index) in $($SBCName1) <-> Index $($matchingRowByName.Index) in $($SBCName2))" -ForegroundColor Cyan
                            $row2 = $matchingRowByName

                            # Add a note about the index difference with success indication
                            $message = "<span style='color: #28a745;'><strong>Name-based match found:</strong></span> Entry '<b>$($row1.$currentSectionTableName)</b>' exists in both files but with different Index values: <b>$($row1.Index)</b> in $($SBCName1) vs <b>$($matchingRowByName.Index)</b> in $($SBCName2)"
                            Write-Host "✓ Successfully matched entry '$($row1.$currentSectionTableName)' despite index difference: $($row1.Index) in $($SBCName1) vs $($matchingRowByName.Index) in $($SBCName2)" -ForegroundColor Green
                            $sectionDifferences.Add($message) | Out-Null
                            $hasTableDiff = $true
                        }
                    }
                }                $rowDifferences = [System.Collections.ArrayList]@()
                $processedProperties = @{}

                foreach ($property in $currentProperties) {
                    # Only process this property if it hasn't been processed before
                    if ($row1.$property -ne $row2.$property -and -not $processedProperties.ContainsKey($property)) {
                        $rowDifferences.Add(@{Property = $property; Value1 = $row1.$property; Value2 = $row2.$property }) | Out-Null
                        $hasTableDiff = $true
                        # Mark this property as processed
                        $processedProperties[$property] = $true
                    }
                }

                if ($rowDifferences.Count -gt 0) {
                    $htmlTable = "Section table '<b>$($sectionName)</b>', row with Index '<b>$($row1.Index)</b>', $($currentSectionTableName) '<b>$($row1.$($currentSectionTableName))</b>' has differences:<br>"
                    Write-Warning "Section table '$($sectionName)', row with Index: '$($row1.Index)', $($currentSectionTableName): '$($row1.$($currentSectionTableName))' has differences:"
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

            # If IncludeNameValidation is enabled, check for entries that exist by name but weren't compared due to different positioning
            if ($IncludeNameValidation) {
                # Get all name values from both tables
                $nameValues1 = $content1 | ForEach-Object { $_.$currentSectionTableName }
                $nameValues2 = $content2 | ForEach-Object { $_.$currentSectionTableName }

                # Find names that exist in both tables
                $commonNames = $nameValues1 | Where-Object { $nameValues2 -contains $_ }

                # Track which names we've already compared
                $comparedNames = @()

                for ($i = 0; $i -lt $content1.Count; $i++) {
                    $row1 = $content1[$i]
                    $row2 = $content2[$i]

                    if ($row1.$currentSectionTableName -eq $row2.$currentSectionTableName) {
                        $comparedNames += $row1.$currentSectionTableName
                    }
                }

                # For names that exist in both but weren't compared at the same index position
                foreach ($name in $commonNames) {
                    if ($name -notin $comparedNames) {
                        $rows1ForName = $content1 | Where-Object { $_.$currentSectionTableName -eq $name }
                        $rows2ForName = $content2 | Where-Object { $_.$currentSectionTableName -eq $name }

                        # Check for duplicate names - if found, skip name-based comparison for this name
                        if (($rows1ForName -is [array] -and $rows1ForName.Count -gt 1) -or ($rows2ForName -is [array] -and $rows2ForName.Count -gt 1)) {
                            Write-Host "⚠️ Duplicate name '$($name)' found in one or both files - skipping name-based comparison" -ForegroundColor Yellow
                            continue
                        }

                        $row1 = $rows1ForName | Select-Object -First 1
                        $row2 = $rows2ForName | Select-Object -First 1

                        $rowDifferences = [System.Collections.ArrayList]@()
                        $processedProperties = @{}

                        foreach ($property in $currentProperties) {
                            # Skip the name property itself as we know they match
                            if ($property -ne $currentSectionTableName) {
                                if ($row1.$property -ne $row2.$property -and -not $processedProperties.ContainsKey($property)) {
                                    $rowDifferences.Add(@{Property = $property; Value1 = $row1.$property; Value2 = $row2.$property }) | Out-Null
                                    $hasTableDiff = $true
                                    $processedProperties[$property] = $true
                                }
                            }
                        }

                        if ($rowDifferences.Count -gt 0) {
                            $htmlTable = "Entry with name '<b>$($name)</b>' (Index $($row1.Index) in $($SBCName1), Index $($row2.Index) in $($SBCName2)) has differences:<br>"
                            Write-Warning "Entry with name '$($name)' (Index $($row1.Index) in $($SBCName1), Index $($row2.Index) in $($SBCName2)) has differences:"
                            $htmlTable += "<table class='diff-table'><tr><th>Property</th><th>$($SBCName1)</th><th>$($SBCName2)</th></tr>"
                            foreach ($diff in $rowDifferences) {
                                $htmlTable += "<tr><td>$($diff.Property)</td><td>$($diff.Value1)</td><td>$($diff.Value2)</td></tr>"
                                Write-Warning "$($diff.Property): $($diff.Value1) (in $($SBCName1)) vs $($diff.Value2) (in $($SBCName2))"
                            }
                            $htmlTable += "</table>"
                            $sectionDifferences.Add($htmlTable) | Out-Null
                        }

                        # Note the index difference - capture index values immediately to avoid variable scope issues
                        $index1 = $row1.Index
                        $index2 = $row2.Index
                        if ($index1 -ne $index2) {
                            if ($rowDifferences.Count -eq 0) {
                                # Perfect match despite different indices
                                $message = "<span style='color: #28a745;'><strong>Perfect name-based match:</strong></span> Entry '<b>$($name)</b>' is identical in both files despite different Index values: <b>$($index1)</b> in $($SBCName1) vs <b>$($index2)</b> in $($SBCName2)"
                                Write-Host "✓ Perfect match found for '$($name)' despite index difference: $($index1) in $($SBCName1) vs $($index2) in $($SBCName2)" -ForegroundColor Green
                            } else {
                                # Match found but with some differences
                                $message = "<span style='color: #ffc107;'><strong>Name-based match with differences:</strong></span> Entry '<b>$($name)</b>' found in both files with different Index values: <b>$($index1)</b> in $($SBCName1) vs <b>$($index2)</b> in $($SBCName2)"
                                Write-Host "ℹ️ Name-based match found for '$($name)' with index difference: $($index1) in $($SBCName1) vs $($index2) in $($SBCName2) (but content differs)" -ForegroundColor Yellow
                            }
                            $sectionDifferences.Add($message) | Out-Null
                            $hasTableDiff = $true
                        }
                    }
                }
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
  .success-message {
    color: #28a745;
    font-weight: bold;
    background-color: #d4edda;
    padding: 8px;
    border-left: 4px solid #28a745;
    margin: 10px 0;
  }
  .warning-message {
    color: #856404;
    font-weight: bold;
    background-color: #fff3cd;
    padding: 8px;
    border-left: 4px solid #ffc107;
    margin: 10px 0;
  }
</style>
</head>
<body>
<h1>SBC INI File Comparison Report</h1>
<h2>Comparison between '$($SBCName1)' and '$($SBCName2)'</h2>
"@

if ($IncludeNameValidation) {
    $htmlHead += "<div style='background-color: #d1ecf1; border: 1px solid #bee5eb; padding: 10px; margin: 15px 0; border-radius: 5px;'>
        <strong>[INFO] Name Validation Enabled:</strong> Table entries are compared by both index and name. Entries with matching names but different indices are automatically cross-referenced.
    </div>"
} else {
    $htmlHead += "<div style='background-color: #f8f9fa; border: 1px solid #dee2e6; padding: 10px; margin: 15px 0; border-radius: 5px;'>
        <strong>[INFO] Name Validation Disabled:</strong> Table entries are compared by index only. Use -IncludeNameValidation `$true for enhanced comparison.
    </div>"
}    $htmlBody = ""
    # Initialize script-level variables for tracking swapped entries
    $script:hasSwappedEntries = $false
    $script:swappedSectionName = ""
    $script:swappedEntryNames = @()

    foreach ($section in $differencesBySection.Keys) {
        # Reset swapped entries tracking for each section
        $script:hasSwappedEntries = $false
        $script:swappedSectionName = ""
        $script:swappedEntryNames = @()

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
            elseif ($item -like "*Perfect name-based match*") {
                $category = "Successful Name-Based Matches"
            }
            elseif ($item -like "*Name-based match found*" -or $item -like "*Name-based match with differences*") {
                $category = "Name-Based Index Corrections"
            }
            elseif ($item -like "*Duplicate Name Detected*") {
                $category = "Duplicate Name Warnings"
            }

            # Add to the appropriate category
            if (-not $categorizedDifferences.Contains($category)) {
                $categorizedDifferences[$category] = [System.Collections.ArrayList]@()
            }
            $categorizedDifferences[$category].Add($item) | Out-Null
        }

        # Track the order of categories for consistent output
        $categoryOrder = @(
            "Duplicate Name Warnings",
            "Successful Name-Based Matches",
            "Name-Based Index Corrections",
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
                # Special handling for Name-Based Index Corrections to detect swapped entries
                if ($cat -eq "Name-Based Index Corrections") {
                    $swappedEntries = @{}
                    $regularCorrections = @()

                    # Parse the corrections to detect swapped pairs
                    foreach ($item in $categorizedDifferences[$cat]) {
                        if ($item -match "Entry '<b>([^']+)</b>' exists in both files but with different Index values: <b>(\d+)</b> in ([^v]+) vs <b>(\d+)</b> in (.+)") {
                            $entryName = $matches[1]
                            $index1 = $matches[2]
                            $index2 = $matches[4]

                            if (-not $swappedEntries.ContainsKey($entryName)) {
                                $swappedEntries[$entryName] = @{
                                    Index1 = $index1
                                    Index2 = $index2
                                }
                            }
                        } elseif ($item -match "Entry '<b>([^']+)</b>' found in both files with different Index values: <b>(\d+)</b> in ([^v]+) vs <b>(\d+)</b> in (.+)") {
                            $entryName = $matches[1]
                            $index1 = $matches[2]
                            $index2 = $matches[4]

                            if (-not $swappedEntries.ContainsKey($entryName)) {
                                $swappedEntries[$entryName] = @{
                                    Index1 = $index1
                                    Index2 = $index2
                                }
                            }
                        } else {
                            $regularCorrections += $item
                        }
                    }

                    # Check if we have exactly 2 entries that are swapped
                    if ($swappedEntries.Count -eq 2) {
                        $entries = $swappedEntries.Keys | Sort-Object
                        $entry1 = $entries[0]
                        $entry2 = $entries[1]

                        # Check if they are truly swapped (entry1's index1 = entry2's index2 and vice versa)
                        if ($swappedEntries[$entry1].Index1 -eq $swappedEntries[$entry2].Index2 -and
                            $swappedEntries[$entry1].Index2 -eq $swappedEntries[$entry2].Index1) {

                            # Mark that we found swapped entries for this section
                            $script:hasSwappedEntries = $true
                            $script:swappedSectionName = $section
                            $script:swappedEntryNames = @($entry1, $entry2)

                            # Generate special swapped entries display
                            $htmlBody += "<div class='warning-message'>"
                            $htmlBody += "<strong>Index Order Swapped:</strong> The entries are arranged in different order in the two files."
                            $htmlBody += "</div>"
                            $htmlBody += "<table class='report'>"
                            $htmlBody += "<tr><th>Entry Comparison</th></tr>"
                            $htmlBody += "<tr class='difference'>"
                            $htmlBody += "<td>"
                            $htmlBody += "<strong>Swapped Index Assignment Detected:</strong><br>"
                            $htmlBody += "<table class='diff-table' style='margin-top: 10px;'>"
                            $htmlBody += "<tr><th>Entry</th><th>$($SBCName1) (Index)</th><th>$($SBCName2) (Index)</th></tr>"
                            $htmlBody += "<tr>"
                            $htmlBody += "<td><strong>$($entry1)</strong></td>"
                            $htmlBody += "<td>Index $($swappedEntries[$entry1].Index1)</td>"
                            $htmlBody += "<td>Index $($swappedEntries[$entry1].Index2)</td>"
                            $htmlBody += "</tr>"
                            $htmlBody += "<tr>"
                            $htmlBody += "<td><strong>$($entry2)</strong></td>"
                            $htmlBody += "<td>Index $($swappedEntries[$entry2].Index1)</td>"
                            $htmlBody += "<td>Index $($swappedEntries[$entry2].Index2)</td>"
                            $htmlBody += "</tr>"
                            $htmlBody += "</table>"
                            $htmlBody += "<div style='margin-top: 10px; padding: 8px; background-color: #e7f3ff; border-left: 4px solid #0066cc;'>"
                            $htmlBody += "<strong>Explanation:</strong> The entries $($entry1) and $($entry2) are functionally identical, "
                            $htmlBody += "but their index positions are swapped in the two files."
                            $htmlBody += "</div>"
                            $htmlBody += "</td>"
                            $htmlBody += "</tr>"
                            $htmlBody += "</table>"

                            # Show any additional regular corrections if they exist
                            if ($regularCorrections.Count -gt 0) {
                                $htmlBody += "<table class='report'><tr><th>Additional Index Corrections</th></tr>"
                                foreach ($item in $regularCorrections) {
                                    $htmlBody += "<tr class='difference'><td>$($item)</td></tr>"
                                }
                                $htmlBody += "</table>"
                            }
                        } else {
                            # Not a simple swap, use regular display
                            $htmlBody += "<table class='report'><tr><th>$cat</th></tr>"
                            foreach ($item in $categorizedDifferences[$cat]) {
                                $htmlBody += "<tr class='difference'><td>$($item)</td></tr>"
                            }
                            $htmlBody += "</table>"
                        }
                    } else {
                        # Not exactly 2 entries or not swapped, use regular display
                        $htmlBody += "<table class='report'><tr><th>$cat</th></tr>"
                        foreach ($item in $categorizedDifferences[$cat]) {
                            $htmlBody += "<tr class='difference'><td>$($item)</td></tr>"
                        }
                        $htmlBody += "</table>"
                    }
                } elseif ($cat -eq "Different Row Values" -and $script:hasSwappedEntries -and $script:swappedSectionName -eq $section) {
                    # Filter out differences that are only about index swapping for the swapped entries
                    $filteredItems = @()
                    foreach ($item in $categorizedDifferences[$cat]) {
                        $shouldSkip = $false
                        foreach ($swappedName in $script:swappedEntryNames) {
                            if ($item -like "*$swappedName*" -and $item -like "*Index*" -and ($item -like "*has differences*")) {
                                # Check if this item only shows Index differences for swapped entries
                                if ($item -match "Section table.*Name '<b>([^']+)</b>' has differences.*<tr><td>Index</td>" -and $matches[1] -eq $swappedName) {
                                    # Count number of <tr><td> entries - if only Index, skip it
                                    $trCount = ($item | Select-String "<tr><td>" -AllMatches).Matches.Count
                                    if ($trCount -eq 1) {
                                        $shouldSkip = $true
                                        break
                                    }
                                }
                                elseif ($item -match "Entry with name '<b>([^']+)</b>'.*has differences.*<tr><td>Index</td>" -and $matches[1] -eq $swappedName) {
                                    # Count number of <tr><td> entries - if only Index, skip it
                                    $trCount = ($item | Select-String "<tr><td>" -AllMatches).Matches.Count
                                    if ($trCount -eq 1) {
                                        $shouldSkip = $true
                                        break
                                    }
                                }
                            }
                        }
                        if (-not $shouldSkip) {
                            $filteredItems += $item
                        }
                    }

                    # Only show the table if there are remaining items after filtering
                    if ($filteredItems.Count -gt 0) {
                        $htmlBody += "<table class='report'><tr><th>$cat</th></tr>"
                        foreach ($item in $filteredItems) {
                            $htmlBody += "<tr class='difference'><td>$($item)</td></tr>"
                        }
                        $htmlBody += "</table>"
                    }
                } else {
                    # Regular category handling
                    $htmlBody += "<table class='report'><tr><th>$cat</th></tr>"
                    foreach ($item in $categorizedDifferences[$cat]) {
                        $htmlBody += "<tr class='difference'><td>$($item)</td></tr>"
                    }
                    $htmlBody += "</table>"
                }
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