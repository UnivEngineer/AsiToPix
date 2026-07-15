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

Describe "Master export planning" {
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
        $legacyEntry = @($legacyPlan.Entries | Where-Object { $_.Status -eq "Legacy" })

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
    It "copies a new master only after confirmation" {
        $sourcePath = Join-Path -Path $TestDrive -ChildPath "apply-copy\source\masterDark.xisf"
        $destinationPath = Join-Path -Path $TestDrive -ChildPath "apply-copy\destination\masterDark.xisf"
        New-Item -ItemType Directory -Path (Split-Path $sourcePath -Parent) -Force | Out-Null
        Set-Content -LiteralPath $sourcePath -Value "new master" -NoNewline
        $sourceLength = (Get-Item -LiteralPath $sourcePath).Length
        $plan = [PSCustomObject]@{
            Entries = @([PSCustomObject]@{
                Status = "Planned"
                MasterType = "Dark"
                SourcePath = $sourcePath
                SourceLength = $sourceLength
                DestinationPath = $destinationPath
            })
        }
        $confirmationReader = { param($Prompt) $null = $Prompt; return [string][char]0x0434 }

        $result = Invoke-AsiToPixMasterExportPlan `
            -Plan $plan `
            -ConfirmationReader $confirmationReader `
            -Confirm:$false

        $result.CopiedCount | Should Be 1
        (Get-Content -LiteralPath $destinationPath -Raw) | Should Be "new master"
    }

    It "keeps a new master uncreated when copying is declined" {
        $sourcePath = Join-Path -Path $TestDrive -ChildPath "apply-decline\source\masterBias.xisf"
        $destinationPath = Join-Path -Path $TestDrive -ChildPath "apply-decline\destination\masterBias.xisf"
        New-Item -ItemType Directory -Path (Split-Path $sourcePath -Parent) -Force | Out-Null
        Set-Content -LiteralPath $sourcePath -Value "new bias" -NoNewline
        $plan = [PSCustomObject]@{
            Entries = @([PSCustomObject]@{
                Status = "Planned"
                MasterType = "Bias"
                SourcePath = $sourcePath
                SourceLength = (Get-Item -LiteralPath $sourcePath).Length
                DestinationPath = $destinationPath
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

    It "renames a legacy master without replacing its contents when accepted" {
        $root = Join-Path -Path $TestDrive -ChildPath "apply-rename"
        $sourcePath = Join-Path -Path $root -ChildPath "source\masterDark-source.xisf"
        $destinationFolder = Join-Path -Path $root -ChildPath "destination"
        $destinationPath = Join-Path -Path $destinationFolder -ChildPath "masterDark_BIN-1_6248x4176.xisf"
        $legacyPath = Join-Path -Path $destinationFolder -ChildPath "masterDark_BIN-1_6248x4176_TEMP--10C_GAIN-120_EXP-180s.xisf"
        New-Item -ItemType Directory -Path (Split-Path $sourcePath -Parent) -Force | Out-Null
        New-Item -ItemType Directory -Path $destinationFolder -Force | Out-Null
        Set-Content -LiteralPath $sourcePath -Value "new contents" -NoNewline
        Set-Content -LiteralPath $legacyPath -Value "old contents" -NoNewline
        $plan = [PSCustomObject]@{
            Entries = @([PSCustomObject]@{
                Status = "Legacy"
                MasterType = "Dark"
                SourcePath = $sourcePath
                SourceLength = (Get-Item -LiteralPath $sourcePath).Length
                DestinationPath = $destinationPath
            })
        }
        $confirmationReader = { param($Prompt) $null = $Prompt; "y" }

        $result = Invoke-AsiToPixMasterExportPlan `
            -Plan $plan `
            -ConfirmationReader $confirmationReader `
            -Confirm:$false

        $result.RenamedCount | Should Be 1
        Test-Path -LiteralPath $legacyPath | Should Be $false
        (Get-Content -LiteralPath $destinationPath -Raw) | Should Be "old contents"
    }

    It "replaces a legacy master transactionally after rename is declined" {
        $root = Join-Path -Path $TestDrive -ChildPath "apply-replace"
        $sourcePath = Join-Path -Path $root -ChildPath "source\masterDark-source.xisf"
        $destinationFolder = Join-Path -Path $root -ChildPath "destination"
        $destinationPath = Join-Path -Path $destinationFolder -ChildPath "masterDark_BIN-1_6248x4176.xisf"
        $legacyPath = Join-Path -Path $destinationFolder -ChildPath "masterDark_BIN-1_6248x4176_TEMP--10C_GAIN-120_EXP-180s.xisf"
        New-Item -ItemType Directory -Path (Split-Path $sourcePath -Parent) -Force | Out-Null
        New-Item -ItemType Directory -Path $destinationFolder -Force | Out-Null
        Set-Content -LiteralPath $sourcePath -Value "new replacement contents" -NoNewline
        Set-Content -LiteralPath $legacyPath -Value "old contents" -NoNewline
        $answers = [System.Collections.Queue]::new()
        $answers.Enqueue("n")
        $answers.Enqueue("y")
        $confirmationReader = {
            param($Prompt)
            $null = $Prompt
            return $answers.Dequeue()
        }.GetNewClosure()
        $plan = [PSCustomObject]@{
            Entries = @([PSCustomObject]@{
                Status = "Legacy"
                MasterType = "Dark"
                SourcePath = $sourcePath
                SourceLength = (Get-Item -LiteralPath $sourcePath).Length
                DestinationPath = $destinationPath
            })
        }

        $result = Invoke-AsiToPixMasterExportPlan `
            -Plan $plan `
            -ConfirmationReader $confirmationReader `
            -Confirm:$false

        $result.ReplacedCount | Should Be 1
        Test-Path -LiteralPath $legacyPath | Should Be $false
        (Get-Content -LiteralPath $destinationPath -Raw) | Should Be "new replacement contents"
        @(Get-ChildItem -LiteralPath $destinationFolder -File -Filter "*.AsToPixBackup-*").Count | Should Be 0
    }

    It "replaces an existing canonical master after confirmation" {
        $root = Join-Path -Path $TestDrive -ChildPath "apply-overwrite"
        $sourcePath = Join-Path -Path $root -ChildPath "source\masterBias-source.xisf"
        $destinationPath = Join-Path -Path $root -ChildPath "destination\masterBias_BIN-1_6248x4176.xisf"
        New-Item -ItemType Directory -Path (Split-Path $sourcePath -Parent) -Force | Out-Null
        New-Item -ItemType Directory -Path (Split-Path $destinationPath -Parent) -Force | Out-Null
        Set-Content -LiteralPath $sourcePath -Value "new bias contents" -NoNewline
        Set-Content -LiteralPath $destinationPath -Value "old bias" -NoNewline
        $plan = [PSCustomObject]@{
            Entries = @([PSCustomObject]@{
                Status = "Exists"
                MasterType = "Bias"
                SourcePath = $sourcePath
                SourceLength = (Get-Item -LiteralPath $sourcePath).Length
                DestinationPath = $destinationPath
            })
        }
        $confirmationReader = { param($Prompt) $null = $Prompt; "Y" }

        $result = Invoke-AsiToPixMasterExportPlan `
            -Plan $plan `
            -ConfirmationReader $confirmationReader `
            -Confirm:$false

        $result.ReplacedCount | Should Be 1
        (Get-Content -LiteralPath $destinationPath -Raw) | Should Be "new bias contents"
        @(Get-ChildItem -LiteralPath (Split-Path $destinationPath -Parent) -File -Filter "*.AsToPixBackup-*").Count |
            Should Be 0
    }

    It "does not prompt or copy in WhatIf mode" {
        $sourcePath = Join-Path -Path $TestDrive -ChildPath "apply-whatif\source\masterFlat.xisf"
        $destinationPath = Join-Path -Path $TestDrive -ChildPath "apply-whatif\destination\masterFlat.xisf"
        New-Item -ItemType Directory -Path (Split-Path $sourcePath -Parent) -Force | Out-Null
        Set-Content -LiteralPath $sourcePath -Value "new flat" -NoNewline
        $plan = [PSCustomObject]@{
            Entries = @([PSCustomObject]@{
                Status = "Planned"
                MasterType = "Flat"
                SourcePath = $sourcePath
                SourceLength = (Get-Item -LiteralPath $sourcePath).Length
                DestinationPath = $destinationPath
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
}
