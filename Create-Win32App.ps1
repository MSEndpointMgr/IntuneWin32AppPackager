<#
.SYNOPSIS
    Create a Win32 app in Microsoft Intune based on input from app manifest file.

.DESCRIPTION
    Create a Win32 app in Microsoft Intune based on input from app manifest file.

.EXAMPLE
    .\Create-Win32App.ps1

.NOTES
    FileName:    Create-Win32App.ps1
    Author:      Nickolaj Andersen
    Contact:     @NickolajA
    Created:     2020-09-26
    Updated:     2020-09-26

    Version history:
    1.0.0 - (2020-09-26) Script created
#>
Process {
    # Read app data from JSON manifest
    $AppDataFile = Join-Path -Path $PSScriptRoot -ChildPath "App.json"
    $AppData = Get-Content -Path $AppDataFile | ConvertFrom-Json

    # Required packaging variables
    $SourceFolder = Join-Path -Path $PSScriptRoot -ChildPath $AppData.PackageInformation.SourceFolder
    $OutputFolder = Join-Path -Path $PSScriptRoot -ChildPath $AppData.PackageInformation.OutputFolder
    $AppIconFile = Join-Path -Path $PSScriptRoot -ChildPath $AppData.PackageInformation.IconFile

    # Connect and retrieve authentication token
    Connect-MSIntuneGraph -TenantName $AppData.TenantInformation.Name -PromptBehavior $AppData.TenantInformation.PromptBehavior -ApplicationID $AppData.TenantInformation.ApplicationID -Verbose

    # Create required .intunewin package from source folder
    $IntuneAppPackage = New-IntuneWin32AppPackage -SourceFolder $SourceFolder -SetupFile $AppData.PackageInformation.SetupFile -OutputFolder $OutputFolder -Verbose

    # Create default requirement rule
    $RequirementRule = New-IntuneWin32AppRequirementRule -Architecture $AppData.RequirementRule.Architecture -MinimumSupportedOperatingSystem $AppData.RequirementRule.MinimumRequiredOperatingSystem

    # Create additional custom requirement rules
    if ($AppData.CustomRequirementRule.Count -ge 1) {
        $RequirementRules = New-Object -TypeName System.Collections.ArrayList
        foreach ($RequirementRuleItem in $AppData.CustomRequirementRule) {
            switch ($AppData.CustomRequirementRule.Type) {
                "File" {
                    switch ($RequirementRuleItem.DetectionMethod) {
                        "Existence" {
                            # Create a custom file based requirement rule
                            $RequirementRuleArgs = @{
                                "Existence" = $true
                                "Path" = $AppData.CustomRequirementRule.Path
                                "FileOrFolder" = $AppData.CustomRequirementRule.FileOrFolder
                                "DetectionType" = $AppData.CustomRequirementRule.DetectionType
                                "Check32BitOn64System" = $AppData.CustomRequirementRule.Check32BitOn64System
                            }
                            $CustomRequirementRule = New-IntuneWin32AppRequirementRuleFile @RequirementRuleArgs
                        }
                        "DateModified" {
                            # Create a custom file based requirement rule
                            $RequirementRuleArgs = @{
                                "DateModified" = $true
                                "Path" = $AppData.CustomRequirementRule.Path
                                "FileOrFolder" = $AppData.CustomRequirementRule.FileOrFolder
                                "Operator" = $AppData.CustomRequirementRule.Operator
                                "DateTimeValue" = $AppData.CustomRequirementRule.DateTimeValue
                                "Check32BitOn64System" = $AppData.CustomRequirementRule.Check32BitOn64System
                            }
                            $CustomRequirementRule = New-IntuneWin32AppRequirementRuleFile @RequirementRuleArgs
                        }
                        "DateCreated" {
                            # Create a custom file based requirement rule
                            $RequirementRuleArgs = @{
                                "DateCreated" = $true
                                "Path" = $AppData.CustomRequirementRule.Path
                                "FileOrFolder" = $AppData.CustomRequirementRule.FileOrFolder
                                "Operator" = $AppData.CustomRequirementRule.Operator
                                "DateTimeValue" = $AppData.CustomRequirementRule.DateTimeValue
                                "Check32BitOn64System" = $AppData.CustomRequirementRule.Check32BitOn64System
                            }
                            $CustomRequirementRule = New-IntuneWin32AppRequirementRuleFile @RequirementRuleArgs
                        }
                        "Version" {
                            # Create a custom file based requirement rule
                            $RequirementRuleArgs = @{
                                "Version" = $true
                                "Path" = $AppData.CustomRequirementRule.Path
                                "FileOrFolder" = $AppData.CustomRequirementRule.FileOrFolder
                                "Operator" = $AppData.CustomRequirementRule.Operator
                                "VersionValue" = $AppData.CustomRequirementRule.VersionValue
                                "Check32BitOn64System" = $AppData.CustomRequirementRule.Check32BitOn64System
                            }
                            $CustomRequirementRule = New-IntuneWin32AppRequirementRuleFile @RequirementRuleArgs
                        }
                        "Size" {
                            # Create a custom file based requirement rule
                            $RequirementRuleArgs = @{
                                "Size" = $true
                                "Path" = $AppData.CustomRequirementRule.Path
                                "FileOrFolder" = $AppData.CustomRequirementRule.FileOrFolder
                                "Operator" = $AppData.CustomRequirementRule.Operator
                                "SizeInMBValue" = $AppData.CustomRequirementRule.SizeInMBValue
                                "Check32BitOn64System" = $AppData.CustomRequirementRule.Check32BitOn64System
                            }
                            $CustomRequirementRule = New-IntuneWin32AppRequirementRuleFile @RequirementRuleArgs
                        }
                    }
                }
                "Registry" {
                    switch ($RequirementRuleItem.DetectionMethod) {
                        "Existence" {
                            # Create a custom registry based requirement rule
                            $RequirementRuleArgs = @{
                                "Existence" = $true
                                "KeyPath" = $AppData.CustomRequirementRule.KeyPath
                                "ValueName" = $AppData.CustomRequirementRule.ValueName
                                "DetectionType" = $AppData.CustomRequirementRule.DetectionType
                                "Check32BitOn64System" = $AppData.CustomRequirementRule.Check32BitOn64System
                            }
                            $CustomRequirementRule = New-IntuneWin32AppRequirementRuleRegistry @RequirementRuleArgs
                        }
                        "StringComparison" {
                            # Create a custom registry based requirement rule
                            $RequirementRuleArgs = @{
                                "StringComparison" = $true
                                "KeyPath" = $AppData.CustomRequirementRule.KeyPath
                                "ValueName" = $AppData.CustomRequirementRule.ValueName
                                "StringComparisonOperator" = $AppData.CustomRequirementRule.Operator
                                "StringComparisonValue" = $AppData.CustomRequirementRule.Value
                                "Check32BitOn64System" = $AppData.CustomRequirementRule.Check32BitOn64System
                            }
                            $CustomRequirementRule = New-IntuneWin32AppRequirementRuleRegistry @RequirementRuleArgs
                        }
                        "VersionComparison" {
                            # Create a custom registry based requirement rule
                            $RequirementRuleArgs = @{
                                "VersionComparison" = $true
                                "KeyPath" = $AppData.CustomRequirementRule.KeyPath
                                "ValueName" = $AppData.CustomRequirementRule.ValueName
                                "VersionComparisonOperator" = $AppData.CustomRequirementRule.Operator
                                "VersionComparisonValue" = $AppData.CustomRequirementRule.Value
                                "Check32BitOn64System" = $AppData.CustomRequirementRule.Check32BitOn64System
                            }
                            $CustomRequirementRule = New-IntuneWin32AppRequirementRuleRegistry @RequirementRuleArgs
                        }
                        "IntegerComparison" {
                            # Create a custom registry based requirement rule
                            $RequirementRuleArgs = @{
                                "IntegerComparison" = $true
                                "KeyPath" = $AppData.CustomRequirementRule.KeyPath
                                "ValueName" = $AppData.CustomRequirementRule.ValueName
                                "IntegerComparisonOperator" = $AppData.CustomRequirementRule.Operator
                                "IntegerComparisonValue" = $AppData.CustomRequirementRule.Value
                                "Check32BitOn64System" = $AppData.CustomRequirementRule.Check32BitOn64System
                            }
                            $CustomRequirementRule = New-IntuneWin32AppRequirementRuleRegistry @RequirementRuleArgs
                        }
                    }
                }
                "Script" {
                    switch ($RequirementRuleItem.DetectionMethod) {
                        "StringOutput" {
                            # Create a custom script based requirement rule
                            $RequirementRuleArgs = @{
                                "StringOutputDataType" = $true
                                "ScriptFile" = $AppData.CustomRequirementRule.ScriptFile
                                "ScriptContext" = $AppData.CustomRequirementRule.ScriptContext
                                "StringComparisonOperator" = $AppData.CustomRequirementRule.Operator
                                "StringValue" = $AppData.CustomRequirementRule.Value
                                "RunAs32BitOn64System" = $AppData.CustomRequirementRule.RunAs32BitOn64System
                                "EnforceSignatureCheck" = $AppData.CustomRequirementRule.EnforceSignatureCheck
                            }
                            $CustomRequirementRule = New-IntuneWin32AppRequirementRuleScript @RequirementRuleArgs
                        }
                        "IntegerOutput" {
                            # Create a custom script based requirement rule
                            $RequirementRuleArgs = @{
                                "IntegerOutputDataType" = $true
                                "ScriptFile" = $AppData.CustomRequirementRule.ScriptFile
                                "ScriptContext" = $AppData.CustomRequirementRule.ScriptContext
                                "IntegerComparisonOperator" = $AppData.CustomRequirementRule.Operator
                                "IntegerValue" = $AppData.CustomRequirementRule.Value
                                "RunAs32BitOn64System" = $AppData.CustomRequirementRule.RunAs32BitOn64System
                                "EnforceSignatureCheck" = $AppData.CustomRequirementRule.EnforceSignatureCheck
                            }
                            $CustomRequirementRule = New-IntuneWin32AppRequirementRuleScript @RequirementRuleArgs
                        }
                        "BooleanOutput" {
                            # Create a custom script based requirement rule
                            $RequirementRuleArgs = @{
                                "BooleanOutputDataType" = $true
                                "ScriptFile" = $AppData.CustomRequirementRule.ScriptFile
                                "ScriptContext" = $AppData.CustomRequirementRule.ScriptContext
                                "BooleanComparisonOperator" = $AppData.CustomRequirementRule.Operator
                                "BooleanValue" = $AppData.CustomRequirementRule.Value
                                "RunAs32BitOn64System" = $AppData.CustomRequirementRule.RunAs32BitOn64System
                                "EnforceSignatureCheck" = $AppData.CustomRequirementRule.EnforceSignatureCheck
                            }
                            $CustomRequirementRule = New-IntuneWin32AppRequirementRuleScript @RequirementRuleArgs
                        }
                        "DateTimeOutput" {
                            # Create a custom script based requirement rule
                            $RequirementRuleArgs = @{
                                "DateTimeOutputDataType" = $true
                                "ScriptFile" = $AppData.CustomRequirementRule.ScriptFile
                                "ScriptContext" = $AppData.CustomRequirementRule.ScriptContext
                                "DateTimeComparisonOperator" = $AppData.CustomRequirementRule.Operator
                                "DateTimeValue" = $AppData.CustomRequirementRule.Value
                                "RunAs32BitOn64System" = $AppData.CustomRequirementRule.RunAs32BitOn64System
                                "EnforceSignatureCheck" = $AppData.CustomRequirementRule.EnforceSignatureCheck
                            }
                            $CustomRequirementRule = New-IntuneWin32AppRequirementRuleScript @RequirementRuleArgs
                        }
                        "FloatOutput" {
                            # Create a custom script based requirement rule
                            $RequirementRuleArgs = @{
                                "FloatOutputDataType" = $true
                                "ScriptFile" = $AppData.CustomRequirementRule.ScriptFile
                                "ScriptContext" = $AppData.CustomRequirementRule.ScriptContext
                                "FloatComparisonOperator" = $AppData.CustomRequirementRule.Operator
                                "FloatValue" = $AppData.CustomRequirementRule.Value
                                "RunAs32BitOn64System" = $AppData.CustomRequirementRule.RunAs32BitOn64System
                                "EnforceSignatureCheck" = $AppData.CustomRequirementRule.EnforceSignatureCheck
                            }
                            $CustomRequirementRule = New-IntuneWin32AppRequirementRuleScript @RequirementRuleArgs
                        }
                        "VersionOutput" {
                            # Create a custom script based requirement rule
                            $RequirementRuleArgs = @{
                                "VersionOutputDataType" = $true
                                "ScriptFile" = $AppData.CustomRequirementRule.ScriptFile
                                "ScriptContext" = $AppData.CustomRequirementRule.ScriptContext
                                "VersionComparisonOperator" = $AppData.CustomRequirementRule.Operator
                                "VersionValue" = $AppData.CustomRequirementRule.Value
                                "RunAs32BitOn64System" = $AppData.CustomRequirementRule.RunAs32BitOn64System
                                "EnforceSignatureCheck" = $AppData.CustomRequirementRule.EnforceSignatureCheck
                            }
                            $CustomRequirementRule = New-IntuneWin32AppRequirementRuleScript @RequirementRuleArgs
                        }
                    }
                }
            }

            # Add requirement rule to list
            $RequirementRules.Add($CustomRequirementRule) | Out-Null
        }
    }

    # Create an array for multiple detection rules if required
    if ($AppData.DetectionRule.Count -gt 1) {
        if ("Script" -in $AppData.DetectionRule.Type) {
            # When a Script detection rule is used, other detection rules cannot be used as well. This should be handled within the module itself by the Add-IntuneWin32App function
        }
    }

    # Create detection rules
    $DetectionRules = New-Object -TypeName System.Collections.ArrayList
    foreach ($DetectionRuleItem in $AppData.DetectionRule) {
        switch ($DetectionRuleItem.Type) {
            "MSI" {
                # Create a MSI installation based detection rule
                $DetectionRuleArgs = @{
                    "ProductCode" = $DetectionRuleItem.ProductCode
                    "Operator" = $DetectionRuleItem.Operator
                    "ProductVersion" = $DetectionRuleItem.ProductVersion
                }
                $DetectionRule = New-IntuneWin32AppDetectionRuleMSI @DetectionRuleArgs
            }
            "Script" {
                # Create a PowerShell script based detection rule
                $DetectionRuleArgs = @{
                    "ScriptFile" = $DetectionRuleItem.ScriptFile
                    "EnforceSignatureCheck" = $DetectionRuleItem.EnforceSignatureCheck
                    "RunAs32Bit" = $DetectionRuleItem.RunAs32Bit
                }
                New-IntuneWin32AppDetectionRuleScript @DetectionRuleArgs
            }
            "Registry" {
                switch ($DetectionRuleItem.DetectionMethod) {
                    "Existence" {
                        # Construct registry existence detection rule parameters
                        $DetectionRuleArgs = @{
                            "Existence" = $true
                            "KeyPath" = $DetectionRuleItem.KeyPath
                            "DetectionType" = $DetectionRuleItem.DetectionType
                            "Check32BitOn64System" = [bool]$DetectionRuleItem.Check32BitOn64System
                        }
                        if (-not([string]::IsNullOrEmpty($DetectionRuleItem.ValueName))) {
                            $DetectionRuleArgs.Add("ValueName", $DetectionRuleItem.ValueName)
                        }
                    }
                    "VersionComparison" {
                        # Construct registry version comparison detection rule parameters
                        $DetectionRuleArgs = @{
                            "VersionComparison" = $true
                            "KeyPath" = $DetectionRuleItem.KeyPath
                            "ValueName" = $DetectionRuleItem.ValueName
                            "VersionComparisonOperator" = $DetectionRuleItem.Operator
                            "VersionComparisonValue" = $DetectionRuleItem.Value
                            "Check32BitOn64System" = [bool]$DetectionRuleItem.Check32BitOn64System
                        }
                    }
                    "StringComparison" {
                        # Construct registry string comparison detection rule parameters
                        $DetectionRuleArgs = @{
                            "StringComparison" = $true
                            "KeyPath" = $DetectionRuleItem.KeyPath
                            "ValueName" = $DetectionRuleItem.ValueName
                            "StringComparisonOperator" = $DetectionRuleItem.Operator
                            "StringComparisonValue" = $DetectionRuleItem.Value
                            "Check32BitOn64System" = [bool]$DetectionRuleItem.Check32BitOn64System
                        }
                    }
                    "IntegerComparison" {
                        # Construct registry integer comparison detection rule parameters
                        $DetectionRuleArgs = @{
                            "IntegerComparison" = $true
                            "KeyPath" = $DetectionRuleItem.KeyPath
                            "ValueName" = $DetectionRuleItem.ValueName
                            "IntegerComparisonOperator" = $DetectionRuleItem.Operator
                            "IntegerComparisonValue" = $DetectionRuleItem.Value
                            "Check32BitOn64System" = [bool]$DetectionRuleItem.Check32BitOn64System
                        }
                    }
                }

                # Create registry based detection rule
                $DetectionRule = New-IntuneWin32AppDetectionRuleRegistry @DetectionRuleArgs
            }
        }

        # Add detection rule to list
        $DetectionRules.Add($DetectionRule) | Out-Null
    }

    # Add icon
    if (Test-Path -Path $AppIconFile) {
        $Icon = New-IntuneWin32AppIcon -FilePath $AppIconFile
    }

    # Construct a table of default parameters for Win32 app
    $Win32AppArgs = @{
        "FilePath" = $IntuneAppPackage.Path
        "DisplayName" = $AppData.Information.DisplayName
        "Description" = $AppData.Information.Description
        "Publisher" = $AppData.Information.Publisher
        "Notes" = $AppData.Information.Notes
        "InstallExperience" = $AppData.Program.InstallExperience
        "RestartBehavior" = $AppData.Program.DeviceRestartBehavior
        "DetectionRule" = $DetectionRules
        "RequirementRule" = $RequirementRule
        "Verbose" = $true
    }

    # Dynamically add additional parameters for Win32 app
    if (Test-Path -Path $AppIconFile) {
        $Win32AppArgs.Add("Icon", $Icon)
    }
    if (-not([string]::IsNullOrEmpty($AppData.Program.InstallCommand))) {
        $Win32AppArgs.Add("InstallCommandLine", $AppData.Program.InstallCommand)
    }
    if (-not([string]::IsNullOrEmpty($AppData.Program.UninstallCommand))) {
        $Win32AppArgs.Add("UninstallCommandLine", $AppData.Program.UninstallCommand)
    }

    # Create Win32 app
    Add-IntuneWin32App @Win32AppArgs
}