$modulePath = Join-Path -Path $PSScriptRoot -ChildPath "..\src\AsiToPix.ExportMasters.psm1"
Import-Module $modulePath -Force

Describe "WBPP master filename parsing" {
    It "removes WBPP metadata tags from a dark master name" {
        $name = "masterDark_BIN-1_6248x4176_EXPOSURE-120.00s_TARGET-H_FILTER-H_TEMP--10C_GAIN-120_EXP-120s.xisf"

        $info = Get-AsiToPixWbppMasterInfo -FileName $name

        $info.MasterType | Should Be "Dark"
        $info.Gain | Should Be "120"
        $info.Temperature | Should Be "-10"
        $info.Exposure | Should Be "120"
        $info.Resolution | Should Be "6248x4176"
        $info.CleanFileName | Should Be "masterDark_BIN-1_6248x4176.xisf"
        $info.Issues.Count | Should Be 0
    }

    It "accepts repeated identical FILTER tags and removes the light EXP tag from a flat name" {
        $name = "masterFlat_BIN-1_6248x4176_FILTER-H_mono_TARGET-H_FILTER-H_TEMP--10C_GAIN-120_EXP-180s.xisf"

        $info = Get-AsiToPixWbppMasterInfo -FileName $name

        $info.MasterType | Should Be "Flat"
        $info.Filter | Should Be "H"
        $info.CleanFileName | Should Be "masterFlat_BIN-1_6248x4176.xisf"
        $info.Issues.Count | Should Be 0
    }

    It "reports conflicting EXP and EXPOSURE values" {
        $name = "masterDark_BIN-1_100x100_EXPOSURE-120.00s_TEMP--10C_GAIN-120_EXP-180s.xisf"

        $info = Get-AsiToPixWbppMasterInfo -FileName $name

        $info.Issues.Count | Should Be 1
        $info.Issues[0] | Should Match "Conflicting EXP/EXPOSURE"
    }
}

Describe "Processing project metadata discovery" {
    It "fuzzy-resolves an object name to its only processing project" {
        $processingRoot = Join-Path -Path $TestDrive -ChildPath "single-processing\AstroPhoto\Processing"
        $projectPath = Join-Path -Path $processingRoot -ChildPath "NGC 7293 - Helix nebula\2026_APO120_ASI2600MM"
        $metadataPath = Join-Path -Path $projectPath -ChildPath "project_meta.json"
        New-Item -ItemType Directory -Path $projectPath -Force | Out-Null
        Set-Content -LiteralPath $metadataPath -Value "{}" -NoNewline

        $candidates = @(Find-AsiToPixProcessingProjectMetadata `
            -ProcessingRoot $processingRoot `
            -ObjectName "Helix")
        $resolvedPath = Resolve-AsiToPixProjectMetadataPath `
            -InputValue "Helix" `
            -ProcessingRoot $processingRoot

        $candidates.Count | Should Be 1
        $candidates[0].ObjectName | Should Be "NGC 7293 - Helix nebula"
        $resolvedPath | Should Be $metadataPath
    }

    It "lists only exact catalog matches and accepts a one-based project index" {
        $processingRoot = Join-Path -Path $TestDrive -ChildPath "multi-processing\AstroPhoto\Processing"
        $firstProject = Join-Path -Path $processingRoot -ChildPath "M16\2025_Setup"
        $secondProject = Join-Path -Path $processingRoot -ChildPath "M16\2026_Setup"
        $otherObjectProject = Join-Path -Path $processingRoot -ChildPath "M17\2026_Setup"
        foreach ($projectPath in @($firstProject, $secondProject, $otherObjectProject)) {
            New-Item -ItemType Directory -Path $projectPath -Force | Out-Null
            Set-Content -LiteralPath (Join-Path $projectPath "project_meta.json") -Value "{}" -NoNewline
        }
        $selectionReader = {
            param($Prompt)
            $null = $Prompt
            return "2"
        }

        $candidates = @(Find-AsiToPixProcessingProjectMetadata `
            -ProcessingRoot $processingRoot `
            -ObjectName "M 16")
        $resolvedPath = Resolve-AsiToPixProjectMetadataPath `
            -InputValue "M 16" `
            -ProcessingRoot $processingRoot `
            -SelectionReader $selectionReader

        $candidates.Count | Should Be 2
        @($candidates | Where-Object { $_.ObjectName -eq "M17" }).Count | Should Be 0
        $resolvedPath | Should Be (Join-Path $secondProject "project_meta.json")
    }

    It "accepts a project folder as direct input" {
        $projectPath = Join-Path -Path $TestDrive -ChildPath "direct-processing\Helix\2026_Setup"
        $metadataPath = Join-Path -Path $projectPath -ChildPath "project_meta.json"
        New-Item -ItemType Directory -Path $projectPath -Force | Out-Null
        Set-Content -LiteralPath $metadataPath -Value "{}" -NoNewline

        Resolve-AsiToPixProjectMetadataPath -InputValue $projectPath | Should Be $metadataPath
    }

    It "reports the searched Processing root when no object matches" {
        $processingRoot = Join-Path -Path $TestDrive -ChildPath "missing-processing\AstroPhoto\Processing"
        New-Item -ItemType Directory -Path $processingRoot -Force | Out-Null

        {
            Resolve-AsiToPixProjectMetadataPath `
                -InputValue "Helix" `
                -ProcessingRoot $processingRoot
        } | Should Throw "No processing project matching object name 'Helix' with project_meta.json was found under '$processingRoot'."
    }
}

Describe "Master export planning" {
    It "uses an existing singular mixed-case calibration folder" {
        $astroPhotoRoot = Join-Path -Path $TestDrive -ChildPath "AstroPhoto"
        $masterRoot = Join-Path -Path $astroPhotoRoot -ChildPath "Calibration\ASI2600MM\Master"
        $existingDarkRoot = Join-Path -Path $masterRoot -ChildPath "dArK"
        New-Item -ItemType Directory -Path $existingDarkRoot -Force | Out-Null
        $env:ASITOPIX_EXPORT_ASTROPHOTO_ROOT = $astroPhotoRoot
        $env:ASITOPIX_EXPORT_DARK_ROOT = $existingDarkRoot

        InModuleScope AsiToPix.ExportMasters {
            $camera = [PSCustomObject]@{ Name = "ASI2600MM" }
            $actual = Get-AsiToPixCalibrationFolder `
                -CameraMetadata $camera `
                -Category Darks `
                -AstroPhotoRoot $env:ASITOPIX_EXPORT_ASTROPHOTO_ROOT

            $actual | Should Be $env:ASITOPIX_EXPORT_DARK_ROOT
        }

        Remove-Item Env:\ASITOPIX_EXPORT_ASTROPHOTO_ROOT
        Remove-Item Env:\ASITOPIX_EXPORT_DARK_ROOT
    }

    It "deduplicates filter-specific dark masters that resolve to one calibration path" {
        $pixPath = Join-Path -Path $TestDrive -ChildPath "project\Pix"
        $masterPath = Join-Path -Path $pixPath -ChildPath "master"
        New-Item -ItemType Directory -Path $masterPath -Force | Out-Null

        $firstName = "masterDark_BIN-1_6248x4176_EXPOSURE-180.00s_TARGET-H_FILTER-H_TEMP--10C_GAIN-120_EXP-180s.xisf"
        $secondName = "masterDark_BIN-1_6248x4176_EXPOSURE-180.00s_TARGET-O_FILTER-O_TEMP--10C_GAIN-120_EXP-180s.xisf"
        Set-Content -LiteralPath (Join-Path $masterPath $firstName) -Value "same master" -NoNewline
        Set-Content -LiteralPath (Join-Path $masterPath $secondName) -Value "same master" -NoNewline

        $destinationFolder = Join-Path -Path $TestDrive -ChildPath "Calibration\ASI2600MM\Master\darks\Gain120\-10C\180sec\26.07"
        $metadata = [PSCustomObject]@{
            PixPath = $pixPath
            Scope = "APO120 @ 0.8x"
            Cameras = @([PSCustomObject]@{ Name = "ASI2600MM" })
        }
        $sourceRecord = [PSCustomObject]@{
            Category = "Darks"
            Camera = "ASI2600MM"
            Gain = "120"
            Temperature = "-10"
            Exposure = "180"
            Filter = $null
            DestinationFolder = $destinationFolder
        }

        $plan = Get-AsiToPixMasterExportPlan -Metadata $metadata -SourceRecord @($sourceRecord)
        $planned = @($plan.Entries | Where-Object { $_.Status -eq "Planned" })
        $duplicates = @($plan.Entries | Where-Object { $_.Status -eq "Duplicate" })

        $planned.Count | Should Be 1
        $duplicates.Count | Should Be 1
        $planned[0].DestinationPath | Should Be (Join-Path $destinationFolder "masterDark_BIN-1_6248x4176.xisf")
        $duplicates[0].DestinationPath | Should Be $planned[0].DestinationPath
    }

    It "deduplicates filter-specific bias masters that resolve to one calibration path" {
        $pixPath = Join-Path -Path $TestDrive -ChildPath "bias-project\Pix"
        $masterPath = Join-Path -Path $pixPath -ChildPath "master"
        New-Item -ItemType Directory -Path $masterPath -Force | Out-Null

        $firstName = "masterBias_BIN-1_6248x4176_TARGET-H_FILTER-H_TEMP--10C_GAIN-120.xisf"
        $secondName = "masterBias_BIN-1_6248x4176_TARGET-O_FILTER-O_TEMP--10C_GAIN-120.xisf"
        Set-Content -LiteralPath (Join-Path $masterPath $firstName) -Value "same bias" -NoNewline
        Set-Content -LiteralPath (Join-Path $masterPath $secondName) -Value "same bias" -NoNewline

        $destinationFolder = Join-Path -Path $TestDrive -ChildPath "Calibration\ASI2600MM\Master\biases\Gain120\-10C\26.07"
        $metadata = [PSCustomObject]@{
            PixPath = $pixPath
            Scope = "APO120 @ 0.8x"
            Cameras = @([PSCustomObject]@{ Name = "ASI2600MM" })
        }
        $sourceRecord = [PSCustomObject]@{
            Category = "Biases"
            Camera = "ASI2600MM"
            Gain = "120"
            Temperature = "-10"
            Exposure = $null
            Filter = $null
            DestinationFolder = $destinationFolder
        }

        $plan = Get-AsiToPixMasterExportPlan -Metadata $metadata -SourceRecord @($sourceRecord)

        @($plan.Entries | Where-Object { $_.Status -eq "Planned" }).Count | Should Be 1
        @($plan.Entries | Where-Object { $_.Status -eq "Duplicate" }).Count | Should Be 1
        @($plan.Entries | Where-Object { $_.Status -eq "Planned" })[0].DestinationPath |
            Should Be (Join-Path $destinationFolder "masterBias_BIN-1_6248x4176.xisf")
    }

    It "deduplicates flat masters independently of the light EXP tag" {
        $pixPath = Join-Path -Path $TestDrive -ChildPath "flat-project\Pix"
        $masterPath = Join-Path -Path $pixPath -ChildPath "master"
        New-Item -ItemType Directory -Path $masterPath -Force | Out-Null

        $firstName = "masterFlat_BIN-1_6248x4176_FILTER-H_mono_TARGET-H_FILTER-H_TEMP--10C_GAIN-120_EXP-120s.xisf"
        $secondName = "masterFlat_BIN-1_6248x4176_FILTER-H_mono_TARGET-H_FILTER-H_TEMP--10C_GAIN-120_EXP-180s.xisf"
        Set-Content -LiteralPath (Join-Path $masterPath $firstName) -Value "same flat" -NoNewline
        Set-Content -LiteralPath (Join-Path $masterPath $secondName) -Value "same flat" -NoNewline

        $destinationFolder = Join-Path -Path $TestDrive -ChildPath "Calibration\ASI2600MM\Master\flats\APO120 @ 0.8x\26.07.08 H 0deg"
        $metadata = [PSCustomObject]@{
            PixPath = $pixPath
            Scope = "APO120 @ 0.8x"
            Cameras = @([PSCustomObject]@{ Name = "ASI2600MM" })
        }
        $sourceRecord = [PSCustomObject]@{
            Category = "Flats"
            Camera = "ASI2600MM"
            Gain = "120"
            Temperature = "-10"
            Exposure = "120"
            Filter = "H"
            DestinationFolder = $destinationFolder
        }

        $plan = Get-AsiToPixMasterExportPlan -Metadata $metadata -SourceRecord @($sourceRecord)

        @($plan.Entries | Where-Object { $_.Status -eq "Planned" }).Count | Should Be 1
        @($plan.Entries | Where-Object { $_.Status -eq "Duplicate" }).Count | Should Be 1
    }

    It "reports differently sized logical flat duplicates as a conflict with choices" {
        $pixPath = Join-Path -Path $TestDrive -ChildPath "flat-selection-project\Pix"
        $masterPath = Join-Path -Path $pixPath -ChildPath "master"
        New-Item -ItemType Directory -Path $masterPath -Force | Out-Null

        $firstName = "masterFlat_BIN-1_6248x4176_FILTER-H_mono_TEMP--10C_GAIN-120_EXP-180s_FILTER-H_TARGET-H.xisf"
        $secondName = "masterFlat_BIN-1_6248x4176_FILTER-H_mono_TEMP--10C_GAIN-120_EXP-300s_FILTER-H_TARGET-H.xisf"
        Set-Content -LiteralPath (Join-Path $masterPath $firstName) -Value "short flat" -NoNewline
        Set-Content -LiteralPath (Join-Path $masterPath $secondName) -Value "longer flat contents" -NoNewline

        $destinationFolder = Join-Path -Path $TestDrive -ChildPath "Calibration\ASI2600MM\Master\flats\SQA55 @ 1.0x\26.07.15 H 180deg"
        $metadata = [PSCustomObject]@{
            PixPath = $pixPath
            Scope = "SQA55 @ 1.0x"
            Cameras = @([PSCustomObject]@{ Name = "ASI2600MM" })
        }
        $sourceRecord = [PSCustomObject]@{
            Category = "Flats"
            Camera = "ASI2600MM"
            Gain = "120"
            Temperature = "-10"
            Exposure = "180"
            Filter = "H"
            DestinationFolder = $destinationFolder
        }

        $plan = Get-AsiToPixMasterExportPlan -Metadata $metadata -SourceRecord @($sourceRecord)
        $selection = @($plan.Entries | Where-Object { $_.Status -eq "Conflict" })

        $selection.Count | Should Be 1
        $selection[0].ChoiceCandidates.Count | Should Be 2
        $selection[0].Reason | Should Match "Select one"
        $selection[0].DestinationPath | Should Be (Join-Path $destinationFolder "masterFlat_BIN-1_6248x4176.xisf")
    }

    It "uses CalibrationSources metadata without reading project Source links" {
        $pixPath = Join-Path -Path $TestDrive -ChildPath "metadata-project\Pix"
        $masterPath = Join-Path -Path $pixPath -ChildPath "master"
        New-Item -ItemType Directory -Path $masterPath -Force | Out-Null

        $name = "masterDark_BIN-1_6248x4176_EXPOSURE-120.00s_TEMP--10C_GAIN-120_EXP-120s.xisf"
        Set-Content -LiteralPath (Join-Path $masterPath $name) -Value "metadata dark" -NoNewline

        $cameraRoot = Join-Path -Path $TestDrive -ChildPath "Calibration\ASI2600MM"
        $darkMasterRoot = Join-Path -Path $cameraRoot -ChildPath "Master\darks"
        $destinationFolder = Join-Path -Path $darkMasterRoot -ChildPath "Gain120\-10C\120sec\26.07"
        $metadata = [PSCustomObject]@{
            SchemaVersion = 2
            PixPath = $pixPath
            WbppMasterPath = $masterPath
            ProjectSourcePath = Join-Path -Path $TestDrive -ChildPath "missing-project-source"
            Scope = "APO120 @ 0.8x"
            Cameras = @([PSCustomObject]@{
                Name = "ASI2600MM"
                CalibrationFolders = [PSCustomObject]@{ Darks = $darkMasterRoot }
            })
            CalibrationSources = @([PSCustomObject]@{
                Type = "Darks"
                Camera = "ASI2600MM"
                SourceMode = "Source"
                SourcePath = Join-Path -Path $cameraRoot -ChildPath "Source\darks\Gain120\-10C\120sec\26.07"
                DestinationFolder = $destinationFolder
                Gain = "120"
                TemperatureC = "-10"
                ExposureSeconds = "120"
                Filter = "H"
                Tag = "dark metadata"
            })
        }

        $plan = Get-AsiToPixMasterExportPlan -Metadata $metadata
        $planned = @($plan.Entries | Where-Object { $_.Status -eq "Planned" })

        $planned.Count | Should Be 1
        $planned[0].DestinationPath | Should Be (Join-Path $destinationFolder "masterDark_BIN-1_6248x4176.xisf")
        $plan.SourceRecordCount | Should Be 1
    }

    It "uses project source mapping instead of an exposure threshold for dark classification" {
        $pixPath = Join-Path -Path $TestDrive -ChildPath "short-dark-project\Pix"
        $masterPath = Join-Path -Path $pixPath -ChildPath "master"
        New-Item -ItemType Directory -Path $masterPath -Force | Out-Null

        $name = "masterDark_BIN-1_100x100_EXPOSURE-5.00s_TEMP--10C_GAIN-120_EXP-5s.xisf"
        Set-Content -LiteralPath (Join-Path $masterPath $name) -Value "short regular dark" -NoNewline

        $destinationFolder = Join-Path -Path $TestDrive -ChildPath "Calibration\ASI2600MM\Master\darks\Gain120\-10C\5sec\26.07"
        $metadata = [PSCustomObject]@{
            PixPath = $pixPath
            Scope = "APO120 @ 0.8x"
            Cameras = @([PSCustomObject]@{ Name = "ASI2600MM" })
        }
        $sourceRecord = [PSCustomObject]@{
            Category = "Darks"
            Camera = "ASI2600MM"
            Gain = "120"
            Temperature = "-10"
            Exposure = "5"
            Filter = $null
            DestinationFolder = $destinationFolder
        }

        $plan = Get-AsiToPixMasterExportPlan -Metadata $metadata -SourceRecord @($sourceRecord)
        $planned = @($plan.Entries | Where-Object { $_.Status -eq "Planned" })

        $planned.Count | Should Be 1
        $planned[0].DestinationPath | Should Match ([regex]::Escape("Master\darks"))
    }

    It "detects one legacy destination master and blocks multiple destination masters" {
        $pixPath = Join-Path -Path $TestDrive -ChildPath "existing-project\Pix"
        $masterPath = Join-Path -Path $pixPath -ChildPath "master"
        New-Item -ItemType Directory -Path $masterPath -Force | Out-Null

        $sourceName = "masterDark_BIN-1_6248x4176_EXPOSURE-180.00s_TEMP--10C_GAIN-120_EXP-180s.xisf"
        Set-Content -LiteralPath (Join-Path $masterPath $sourceName) -Value "new dark" -NoNewline

        $destinationFolder = Join-Path -Path $TestDrive -ChildPath "Calibration\ASI2600MM\Master\darks\Gain120\-10C\180sec\26.07"
        New-Item -ItemType Directory -Path $destinationFolder -Force | Out-Null
        $legacyName = "masterDark_BIN-1_6248x4176_EXPOSURE-180.00s_TEMP--10C_GAIN-120_EXP-180s.xisf"
        $legacyPath = Join-Path -Path $destinationFolder -ChildPath $legacyName
        Set-Content -LiteralPath $legacyPath -Value "old dark" -NoNewline

        $metadata = [PSCustomObject]@{
            PixPath = $pixPath
            Scope = "APO120 @ 0.8x"
            Cameras = @([PSCustomObject]@{ Name = "ASI2600MM" })
        }
        $sourceRecord = [PSCustomObject]@{
            Category = "Darks"
            Camera = "ASI2600MM"
            Gain = "120"
            Temperature = "-10"
            Exposure = "180"
            Filter = $null
            DestinationFolder = $destinationFolder
        }

        $legacyPlan = Get-AsiToPixMasterExportPlan -Metadata $metadata -SourceRecord @($sourceRecord)
        $legacyEntry = @($legacyPlan.Entries | Where-Object { $_.Status -eq "Exists" })

        $legacyEntry.Count | Should Be 1
        $legacyEntry[0].ExistingMasterPath | Should Be $legacyPath

        Set-Content `
            -LiteralPath (Join-Path $destinationFolder "masterDark_BIN-2_6248x4176.xisf") `
            -Value "another dark" `
            -NoNewline
        $conflictPlan = Get-AsiToPixMasterExportPlan -Metadata $metadata -SourceRecord @($sourceRecord)
        $conflictEntry = @($conflictPlan.Entries | Where-Object { $_.Status -eq "Conflict" })

        $conflictEntry.Count | Should Be 1
        $conflictEntry[0].Reason | Should Match "multiple masters"
    }
}

Describe "Master export application" {
    It "copies every planned master after one plan confirmation" {
        $root = Join-Path -Path $TestDrive -ChildPath "apply-copy"
        $firstSourcePath = Join-Path -Path $root -ChildPath "source\masterDark.xisf"
        $secondSourcePath = Join-Path -Path $root -ChildPath "source\masterBias.xisf"
        $firstDestinationPath = Join-Path -Path $root -ChildPath "destination\darks\masterDark_BIN-1_6248x4176.xisf"
        $secondDestinationPath = Join-Path -Path $root -ChildPath "destination\biases\masterBias_BIN-1_6248x4176.xisf"
        New-Item -ItemType Directory -Path (Split-Path -Path $firstSourcePath -Parent) -Force | Out-Null
        Set-Content -LiteralPath $firstSourcePath -Value "new dark" -NoNewline
        Set-Content -LiteralPath $secondSourcePath -Value "new bias" -NoNewline
        $plan = [PSCustomObject]@{
            Entries = @(
                [PSCustomObject]@{
                    Status = "Planned"; MasterType = "Dark"; SourcePath = $firstSourcePath
                    SourceLength = (Get-Item -LiteralPath $firstSourcePath).Length; DestinationPath = $firstDestinationPath
                },
                [PSCustomObject]@{
                    Status = "Planned"; MasterType = "Bias"; SourcePath = $secondSourcePath
                    SourceLength = (Get-Item -LiteralPath $secondSourcePath).Length; DestinationPath = $secondDestinationPath
                }
            )
        }
        $prompts = [System.Collections.Generic.List[string]]::new()
        $confirmationReader = {
            param($Prompt)
            $prompts.Add($Prompt) | Out-Null
            return "y"
        }.GetNewClosure()

        $result = Invoke-AsiToPixMasterExportPlan `
            -Plan $plan `
            -ConfirmationReader $confirmationReader `
            -Confirm:$false

        $prompts.Count | Should Be 1
        $result.CopiedCount | Should Be 2
        (Get-Content -LiteralPath $firstDestinationPath -Raw) | Should Be "new dark"
        (Get-Content -LiteralPath $secondDestinationPath -Raw) | Should Be "new bias"
    }

    It "keeps all planned masters uncreated when plan execution is declined" {
        $sourcePath = Join-Path -Path $TestDrive -ChildPath "apply-decline\source\masterBias.xisf"
        $destinationPath = Join-Path -Path $TestDrive -ChildPath "apply-decline\destination\masterBias_BIN-1_6248x4176.xisf"
        New-Item -ItemType Directory -Path (Split-Path -Path $sourcePath -Parent) -Force | Out-Null
        Set-Content -LiteralPath $sourcePath -Value "new bias" -NoNewline
        $plan = [PSCustomObject]@{
            Entries = @([PSCustomObject]@{
                Status = "Planned"; MasterType = "Bias"; SourcePath = $sourcePath
                SourceLength = (Get-Item -LiteralPath $sourcePath).Length; DestinationPath = $destinationPath
            })
        }
        $confirmationReader = { param($Prompt) $null = $Prompt; return [string][char]0x043D }

        $result = Invoke-AsiToPixMasterExportPlan `
            -Plan $plan `
            -ConfirmationReader $confirmationReader `
            -Confirm:$false

        $result.DeclinedCount | Should Be 1
        Test-Path -LiteralPath $destinationPath | Should Be $false
    }

    It "does not overwrite an existing master or prompt for execution" {
        $root = Join-Path -Path $TestDrive -ChildPath "apply-existing"
        $sourcePath = Join-Path -Path $root -ChildPath "source\masterBias-source.xisf"
        $destinationPath = Join-Path -Path $root -ChildPath "destination\masterBias_BIN-1_6248x4176.xisf"
        New-Item -ItemType Directory -Path (Split-Path -Path $sourcePath -Parent) -Force | Out-Null
        New-Item -ItemType Directory -Path (Split-Path -Path $destinationPath -Parent) -Force | Out-Null
        Set-Content -LiteralPath $sourcePath -Value "new bias contents" -NoNewline
        Set-Content -LiteralPath $destinationPath -Value "old bias contents" -NoNewline
        $plan = [PSCustomObject]@{
            Entries = @([PSCustomObject]@{
                Status = "Exists"; MasterType = "Bias"; SourcePath = $sourcePath
                SourceLength = (Get-Item -LiteralPath $sourcePath).Length; DestinationPath = $destinationPath
            })
        }
        $confirmationReader = { param($Prompt) $null = $Prompt; throw "Existing masters must not prompt." }

        $result = Invoke-AsiToPixMasterExportPlan `
            -Plan $plan `
            -ConfirmationReader $confirmationReader `
            -Confirm:$false

        $result.CopiedCount | Should Be 0
        (Get-Content -LiteralPath $destinationPath -Raw) | Should Be "old bias contents"
    }

    It "does not prompt or copy in WhatIf mode" {
        $sourcePath = Join-Path -Path $TestDrive -ChildPath "apply-whatif\source\masterFlat.xisf"
        $destinationPath = Join-Path -Path $TestDrive -ChildPath "apply-whatif\destination\masterFlat.xisf"
        New-Item -ItemType Directory -Path (Split-Path -Path $sourcePath -Parent) -Force | Out-Null
        Set-Content -LiteralPath $sourcePath -Value "new flat" -NoNewline
        $plan = [PSCustomObject]@{
            Entries = @([PSCustomObject]@{
                Status = "Planned"; MasterType = "Flat"; SourcePath = $sourcePath
                SourceLength = (Get-Item -LiteralPath $sourcePath).Length; DestinationPath = $destinationPath
            })
        }
        $confirmationReader = { param($Prompt) $null = $Prompt; throw "WhatIf must not prompt." }

        $result = Invoke-AsiToPixMasterExportPlan `
            -Plan $plan `
            -ConfirmationReader $confirmationReader `
            -WhatIf `
            -Confirm:$false

        $result.WhatIfCount | Should Be 1
        Test-Path -LiteralPath $destinationPath | Should Be $false
    }

    It "rebuilds the plan from an interactively selected conflict before execution" {
        $root = Join-Path -Path $TestDrive -ChildPath "apply-selection"
        $firstSourcePath = Join-Path -Path $root -ChildPath "source\masterFlat_EXP-180s.xisf"
        $secondSourcePath = Join-Path -Path $root -ChildPath "source\masterFlat_EXP-300s.xisf"
        $destinationPath = Join-Path -Path $root -ChildPath "destination\masterFlat_BIN-1_6248x4176.xisf"
        New-Item -ItemType Directory -Path (Split-Path -Path $firstSourcePath -Parent) -Force | Out-Null
        Set-Content -LiteralPath $firstSourcePath -Value "first flat" -NoNewline
        Set-Content -LiteralPath $secondSourcePath -Value "second selected flat" -NoNewline
        $firstCandidate = [PSCustomObject]@{
            Status = "Candidate"; MasterType = "Flat"; SourcePath = $firstSourcePath
            SourceLength = (Get-Item -LiteralPath $firstSourcePath).Length; DestinationPath = $destinationPath
            ExistingMasterPath = $null; ExistingMasterCount = 0; DuplicateOf = $null; ChoiceCandidates = @(); Reason = $null
        }
        $secondCandidate = [PSCustomObject]@{
            Status = "Candidate"; MasterType = "Flat"; SourcePath = $secondSourcePath
            SourceLength = (Get-Item -LiteralPath $secondSourcePath).Length; DestinationPath = $destinationPath
            ExistingMasterPath = $null; ExistingMasterCount = 0; DuplicateOf = $null; ChoiceCandidates = @(); Reason = $null
        }
        $plan = [PSCustomObject]@{
            MasterPath = (Split-Path -Path $firstSourcePath -Parent); ProjectSourcePath = $root
            XisfFileCount = 2; EligibleMasterCount = 2; IgnoredFileCount = 0; Diagnostics = @()
            Entries = @([PSCustomObject]@{
                Status = "Conflict"; MasterType = "Flat"; SourcePath = $firstSourcePath
                SourceLength = $firstCandidate.SourceLength; DestinationPath = $destinationPath
                ExistingMasterPath = $null; ExistingMasterCount = 0; DuplicateOf = $null
                ChoiceCandidates = @($firstCandidate, $secondCandidate)
                Reason = "Logical duplicates have different file sizes."
            })
        }
        $selectionReader = { param($Prompt) $null = $Prompt; "2" }
        $confirmationReader = { param($Prompt) $null = $Prompt; "y" }

        $result = Invoke-AsiToPixMasterExportPlan `
            -Plan $plan `
            -SelectionReader $selectionReader `
            -ConfirmationReader $confirmationReader `
            -Confirm:$false

        $result.SelectedCount | Should Be 1
        $result.CopiedCount | Should Be 1
        @($plan.Entries | Where-Object { $_.Status -eq "Conflict" }).Count | Should Be 0
        @($plan.Entries | Where-Object { $_.Status -eq "Planned" }).Count | Should Be 1
        (Get-Content -LiteralPath $destinationPath -Raw) | Should Be "second selected flat"
    }

    It "does not prompt for a duplicate conflict in WhatIf mode" {
        $root = Join-Path -Path $TestDrive -ChildPath "apply-selection-whatif"
        $firstSourcePath = Join-Path -Path $root -ChildPath "source\masterFlat_EXP-180s.xisf"
        $secondSourcePath = Join-Path -Path $root -ChildPath "source\masterFlat_EXP-300s.xisf"
        $destinationPath = Join-Path -Path $root -ChildPath "destination\masterFlat_BIN-1_6248x4176.xisf"
        New-Item -ItemType Directory -Path (Split-Path -Path $firstSourcePath -Parent) -Force | Out-Null
        Set-Content -LiteralPath $firstSourcePath -Value "first flat" -NoNewline
        Set-Content -LiteralPath $secondSourcePath -Value "second selected flat" -NoNewline
        $plan = [PSCustomObject]@{
            Entries = @([PSCustomObject]@{
                Status = "Conflict"; MasterType = "Flat"; DestinationPath = $destinationPath
                ChoiceCandidates = @(
                    [PSCustomObject]@{ Status = "Candidate"; MasterType = "Flat"; SourcePath = $firstSourcePath; SourceLength = (Get-Item -LiteralPath $firstSourcePath).Length; DestinationPath = $destinationPath },
                    [PSCustomObject]@{ Status = "Candidate"; MasterType = "Flat"; SourcePath = $secondSourcePath; SourceLength = (Get-Item -LiteralPath $secondSourcePath).Length; DestinationPath = $destinationPath }
                )
            })
        }
        $inputReader = { param($Prompt) $null = $Prompt; throw "WhatIf must not prompt." }

        $result = Invoke-AsiToPixMasterExportPlan `
            -Plan $plan `
            -SelectionReader $inputReader `
            -ConfirmationReader $inputReader `
            -WhatIf `
            -Confirm:$false

        $result.ConflictPendingCount | Should Be 1
        $result.WhatIfCount | Should Be 0
        Test-Path -LiteralPath $destinationPath | Should Be $false
    }
}
