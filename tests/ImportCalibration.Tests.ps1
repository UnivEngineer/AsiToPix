Describe "Calibration filename parsing" {
    $modulePath = Join-Path -Path $PSScriptRoot -ChildPath "..\src\AsiToPix.ImportCalibration.psm1"
    Import-Module $modulePath -Force

    It "parses ASIAir bias metadata and converts milliseconds to seconds" {
        $name = "Bias_1.0ms_Bin1_2600MC_IRC_gain120_20260709-051753_185deg_-9.9C_0001.fit"

        $info = ConvertFrom-AsiToPixCalibrationFileName -FileName $name -Category Bias

        $info.CameraName | Should Be "ASI2600MC"
        $info.Gain | Should Be "120"
        $info.TemperatureC | Should Be "-9.9"
        $info.ExposureSeconds | Should Be "0.001"
        $info.FilterName | Should Be "IRC"
        $info.AngleDegrees | Should Be "185"
        $info.CapturedAt | Should Be ([datetime]"2026-07-09T05:17:53")
    }

    It "uses the shared PixInsight image extension list" {
        Test-AsiToPixSupportedCalibrationFileName -FileName "frame.FIT" | Should Be $true
        Test-AsiToPixSupportedCalibrationFileName -FileName "frame.fits" | Should Be $true
        Test-AsiToPixSupportedCalibrationFileName -FileName "frame.ARW" | Should Be $true
        Test-AsiToPixSupportedCalibrationFileName -FileName "frame.CR3" | Should Be $true
        Test-AsiToPixSupportedCalibrationFileName -FileName "frame.xisf" | Should Be $true
        Test-AsiToPixSupportedCalibrationFileName -FileName "frame.tiff" | Should Be $true
        Test-AsiToPixSupportedCalibrationFileName -FileName "result.fit.fz" | Should Be $true
        Test-AsiToPixSupportedCalibrationFileName -FileName "notes.txt" | Should Be $false
    }

    It "accepts singular and plural calibration folder names" {
        InModuleScope AsiToPix.ImportCalibration {
            Get-AsiToPixCalibrationCategoryName -FolderName "bias" | Should Be "Bias"
            Get-AsiToPixCalibrationCategoryName -FolderName "biases" | Should Be "Bias"
            Get-AsiToPixCalibrationCategoryName -FolderName "dark" | Should Be "Dark"
            Get-AsiToPixCalibrationCategoryName -FolderName "darks" | Should Be "Dark"
            Get-AsiToPixCalibrationCategoryName -FolderName "flat" | Should Be "Flat"
            Get-AsiToPixCalibrationCategoryName -FolderName "flats" | Should Be "Flat"
            Get-AsiToPixCalibrationCategoryName -FolderName "BiAsEs" | Should Be "Bias"
            Get-AsiToPixCalibrationCategoryName -FolderName "DaRkS" | Should Be "Dark"
            Get-AsiToPixCalibrationCategoryName -FolderName "FlAtS" | Should Be "Flat"
        }
    }

    It "finds direct bias, dark, and flat folders below every Import setup" {
        $importRoot = Join-Path -Path $TestDrive -ChildPath "discovery\AstroPhoto\Import"
        $firstDarkFolder = Join-Path -Path $importRoot -ChildPath "APO120 @ 0.8x\darks"
        $secondDarkFolder = Join-Path -Path $importRoot -ChildPath "SQA55 @ 1.0x\DaRk"
        $biasFolder = Join-Path -Path $importRoot -ChildPath "APO120 @ 0.8x\biases"
        $flatFolder = Join-Path -Path $importRoot -ChildPath "Canon EF 200\flats"
        New-Item -ItemType Directory -Path $firstDarkFolder -Force | Out-Null
        New-Item -ItemType Directory -Path $secondDarkFolder -Force | Out-Null
        New-Item -ItemType Directory -Path $biasFolder -Force | Out-Null
        New-Item -ItemType Directory -Path $flatFolder -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $importRoot "Nested\other\darks") -Force | Out-Null

        $folders = @(Find-AsiToPixCalibrationImportFolder -ImportRoot $importRoot)

        $folders.Count | Should Be 4
        @($folders.SourcePath) | Should Be @($biasFolder, $firstDarkFolder, $flatFolder, $secondDarkFolder)
        @($folders.Category) | Should Be @("Bias", "Dark", "Flat", "Dark")

        $darkFolders = @(Find-AsiToPixCalibrationImportFolder -ImportRoot $importRoot -Category Dark)
        @($darkFolders.SourcePath) | Should Be @($firstDarkFolder, $secondDarkFolder)
    }

    It "uses the parent setup name when a category folder is imported directly" {
        $darkFolder = Join-Path -Path $TestDrive -ChildPath "direct-category\APO120 @ 0.8x\darks"
        New-Item -ItemType Directory -Path $darkFolder -Force | Out-Null
        $env:ASITOPIX_DIRECT_DARK_FOLDER = $darkFolder

        InModuleScope AsiToPix.ImportCalibration {
            Get-AsiToPixCalibrationSetupName -SourcePath $env:ASITOPIX_DIRECT_DARK_FOLDER |
                Should Be "APO120 @ 0.8x"
        }

        Remove-Item Env:\ASITOPIX_DIRECT_DARK_FOLDER
    }

    It "uses automatic calibration-folder discovery only when SourcePath is omitted" {
        $entryScriptPath = Join-Path -Path $PSScriptRoot -ChildPath "..\ImportCalibration.ps1"
        $entryScriptText = Get-Content -LiteralPath $entryScriptPath -Raw

        $entryScriptText | Should Match 'Read from\s+: <selected root>\\Import'
        $entryScriptText | Should Match 'Write to\s+: <selected root>\\Calibration'
        $entryScriptText | Should Match 'Calibration source discovery \(read\): \$importRoot'
        $entryScriptText | Should Match 'Calibration library destination \(write\): \$calibrationRoot'
        $entryScriptText | Should Match 'Find-AsiToPixCalibrationImportFolder -ImportRoot \$importRoot'
        $entryScriptText | Should Not Match 'Find-AsiToPixCalibrationImportFolder -ImportRoot \$importRoot -Category Dark'
        $entryScriptText | Should Match 'Import all \$\(\$discoveredFolders\.Count\) detected calibration folder\(s\)\?'
        $entryScriptText | Should Match '-SkipConfirmation:\$useDiscoveryConfirmation'
    }

    It "accepts null entries when deriving an interactive metadata default" {
        InModuleScope AsiToPix.ImportCalibration {
            Get-AsiToPixCalibrationDefaultValue -Value @("S", $null, "S") | Should Be "S"
            Get-AsiToPixCalibrationDefaultValue -Value @("S", $null, "O") | Should Be ""
            Get-AsiToPixCalibrationDefaultValue -Value @($null, $null) | Should Be ""
        }
    }

    It "shows relative examples before requesting missing metadata" {
        InModuleScope AsiToPix.ImportCalibration {
            $records = @(
                [PSCustomObject]@{ File = [PSCustomObject]@{ FullName = "C:\Import\Flat\example1.fit" } },
                [PSCustomObject]@{ File = [PSCustomObject]@{ FullName = "C:\Import\Flat\example2.fit" } }
            )
            Mock Write-Host {}

            Write-AsiToPixCalibrationMissingMetadataExample `
                -Description "Flat filter metadata" `
                -SourcePath "C:\Import" `
                -Record $records `
                -MaximumExampleCount 1

            Assert-MockCalled Write-Host -Times 1 -ParameterFilter {
                $Object -eq "`n[INFO] Flat filter metadata is missing for 2 file(s). Examples:" -and
                    $ForegroundColor -eq "Yellow"
            }
            Assert-MockCalled Write-Host -Times 1 -ParameterFilter {
                $Object -eq "  Flat\example1.fit" -and $ForegroundColor -eq "DarkGray"
            }
            Assert-MockCalled Write-Host -Times 1 -ParameterFilter {
                $Object -eq "  ... and 1 more file(s)" -and $ForegroundColor -eq "DarkGray"
            }
        }
    }

    It "uses noon as the boundary for calibration night dates" {
        InModuleScope AsiToPix.ImportCalibration {
            (Get-AsiToPixCalibrationNightStart -CapturedAt ([datetime]"2026-07-09T11:59:59")) |
                Should Be ([datetime]"2026-07-08")
            (Get-AsiToPixCalibrationNightStart -CapturedAt ([datetime]"2026-07-09T12:00:00")) |
                Should Be ([datetime]"2026-07-09")
        }
    }
}

Describe "Calibration import planning" {
    $modulePath = Join-Path -Path $PSScriptRoot -ChildPath "..\src\AsiToPix.ImportCalibration.psm1"
    Import-Module $modulePath -Force

    It "builds canonical bias, dark, and normalized mono flat paths" {
        $sourceRoot = Join-Path -Path $TestDrive -ChildPath "APO120 @ 0.8x"
        $calibrationRoot = Join-Path -Path $TestDrive -ChildPath "Calibration"
        New-Item -ItemType Directory -Path (Join-Path $sourceRoot "biases") -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $sourceRoot "dark") -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path $sourceRoot "flats") -Force | Out-Null
        New-Item -ItemType Directory -Path $calibrationRoot -Force | Out-Null

        $biasName = "Bias_1.0ms_Bin1_2600MM_H_gain120_20260709-051753_0deg_-9.9C_0001.fit"
        $darkName = "Dark_180.0s_Bin1_2600MM_H_gain120_20260709-061753_0deg_-10.1C_0001.fits"
        $flatName = "Flat_100.0ms_Bin1_2600MM_Ha_gain120_20260708-211753_0deg_-10.0C_0001.fit"
        Set-Content -LiteralPath (Join-Path $sourceRoot "biases\$biasName") -Value "bias" -NoNewline
        Set-Content -LiteralPath (Join-Path $sourceRoot "dark\$darkName") -Value "dark" -NoNewline
        Set-Content -LiteralPath (Join-Path $sourceRoot "flats\$flatName") -Value "flat" -NoNewline

        $plan = Get-AsiToPixCalibrationImportPlan `
            -SourcePath $sourceRoot `
            -CalibrationRoot $calibrationRoot
        $destinations = @($plan.Entries | Select-Object -ExpandProperty DestinationPath)

        $plan.PlannedCount | Should Be 3
        ($destinations -contains (Join-Path $calibrationRoot "ASI2600MM\Source\biases\gain120\-10C\26.07\$biasName")) | Should Be $true
        ($destinations -contains (Join-Path $calibrationRoot "ASI2600MM\Source\darks\gain120\-10C\180sec\26.07\$darkName")) | Should Be $true
        ($destinations -contains (Join-Path $calibrationRoot "ASI2600MM\Source\flats\APO120 @ 0.8x\26.07.08 H 0deg\$flatName")) | Should Be $true
    }

    It "separates different flat exposures from the same night with a neutral suffix" {
        $sourceRoot = Join-Path -Path $TestDrive -ChildPath "flat-exposure-collision\APO120 @ 0.8x"
        $sourceFlats = Join-Path -Path $sourceRoot -ChildPath "flats"
        $calibrationRoot = Join-Path -Path $TestDrive -ChildPath "flat-exposure-collision\Calibration"
        New-Item -ItemType Directory -Path $sourceFlats -Force | Out-Null
        New-Item -ItemType Directory -Path $calibrationRoot -Force | Out-Null

        $shortName = "Flat_10.0ms_Bin1_2600MM_L_gain120_20260409-033511_2deg_-19.4C_APO120_0001.fit"
        $longName = "Flat_800.0ms_Bin1_2600MM_L_gain120_20260409-033759_2deg_-20.3C_APO120_0001.fit"
        Set-Content -LiteralPath (Join-Path $sourceFlats $shortName) -Value "short flat" -NoNewline
        Set-Content -LiteralPath (Join-Path $sourceFlats $longName) -Value "long flat" -NoNewline

        $plan = Get-AsiToPixCalibrationImportPlan `
            -SourcePath $sourceRoot `
            -CalibrationRoot $calibrationRoot
        $shortEntry = $plan.Entries | Where-Object { $_.SourcePath -like "*\$shortName" }
        $longEntry = $plan.Entries | Where-Object { $_.SourcePath -like "*\$longName" }

        $shortEntry.DestinationFolder |
            Should Be (Join-Path $calibrationRoot "ASI2600MM\Source\flats\APO120 @ 0.8x\26.04.08 L 2deg")
        $longEntry.DestinationFolder |
            Should Be (Join-Path $calibrationRoot "ASI2600MM\Source\flats\APO120 @ 0.8x\26.04.08 L 2deg 800ms")
        $plan.FlatExposureSplits.Count | Should Be 1
        $plan.FlatExposureSplits[0].Suffix | Should Be "800ms"
    }

    It "reuses an exposure suffix when the canonical flat folder holds another exposure" {
        $sourceRoot = Join-Path -Path $TestDrive -ChildPath "existing-flat-collision\APO120 @ 0.8x"
        $sourceFlats = Join-Path -Path $sourceRoot -ChildPath "flats"
        $calibrationRoot = Join-Path -Path $TestDrive -ChildPath "existing-flat-collision\Calibration"
        $baseFolder = Join-Path $calibrationRoot "ASI2600MM\Source\flats\APO120 @ 0.8x\26.04.08 L 2deg"
        New-Item -ItemType Directory -Path $sourceFlats -Force | Out-Null
        New-Item -ItemType Directory -Path $baseFolder -Force | Out-Null

        $existingName = "Flat_10.0ms_Bin1_2600MM_L_gain120_20260409-033511_2deg_-19.4C_APO120_0001.fit"
        $newName = "Flat_800.0ms_Bin1_2600MM_L_gain120_20260409-033759_2deg_-20.3C_APO120_0001.fit"
        Set-Content -LiteralPath (Join-Path $baseFolder $existingName) -Value "existing flat" -NoNewline
        Set-Content -LiteralPath (Join-Path $sourceFlats $newName) -Value "new flat" -NoNewline

        $plan = Get-AsiToPixCalibrationImportPlan `
            -SourcePath $sourceRoot `
            -CalibrationRoot $calibrationRoot

        $plan.Entries[0].DestinationFolder |
            Should Be (Join-Path $calibrationRoot "ASI2600MM\Source\flats\APO120 @ 0.8x\26.04.08 L 2deg 800ms")
        $plan.FlatExposureSplits[0].ExposureSeconds | Should Be "0.8"

        New-Item -ItemType Directory -Path $plan.Entries[0].DestinationFolder -Force | Out-Null
        Copy-Item -LiteralPath $plan.Entries[0].SourcePath -Destination $plan.Entries[0].DestinationPath
        $repeatPlan = Get-AsiToPixCalibrationImportPlan `
            -SourcePath $sourceRoot `
            -CalibrationRoot $calibrationRoot

        $repeatPlan.ExistingCount | Should Be 1
        $repeatPlan.PlannedCount | Should Be 0
        $repeatPlan.Entries[0].DestinationFolder | Should Be $plan.Entries[0].DestinationFolder
    }

    It "uses supplied metadata and file timestamps for RAW calibration files" {
        $sourceRoot = Join-Path -Path $TestDrive -ChildPath "Canon EF 200 F2.8 MK2"
        $calibrationRoot = Join-Path -Path $TestDrive -ChildPath "raw-Calibration"
        foreach ($folder in @("bias", "darks", "flats")) {
            New-Item -ItemType Directory -Path (Join-Path $sourceRoot $folder) -Force | Out-Null
        }
        New-Item -ItemType Directory -Path $calibrationRoot -Force | Out-Null

        $bias = New-Item -ItemType File -Path (Join-Path $sourceRoot "bias\A7400001.ARW") -Force
        $dark = New-Item -ItemType File -Path (Join-Path $sourceRoot "darks\A7400002.ARW") -Force
        $flat = New-Item -ItemType File -Path (Join-Path $sourceRoot "flats\A7400003.ARW") -Force
        $bias.LastWriteTime = [datetime]"2026-07-12T10:00:00"
        $dark.LastWriteTime = [datetime]"2026-07-13T10:00:00"
        $flat.LastWriteTime = [datetime]"2026-07-12T09:00:00"

        $plan = Get-AsiToPixCalibrationImportPlan `
            -SourcePath $sourceRoot `
            -CalibrationRoot $calibrationRoot `
            -CameraName "SonyA7IV" `
            -Gain "100" `
            -TemperatureC "21" `
            -DarkExposureSeconds "120" `
            -FilterName "None"
        $destinations = @($plan.Entries | Select-Object -ExpandProperty DestinationPath)

        ($destinations -contains (Join-Path $calibrationRoot "SonyA7IV\Source\biases\gain100\20C\26.07\A7400001.ARW")) | Should Be $true
        ($destinations -contains (Join-Path $calibrationRoot "SonyA7IV\Source\darks\gain100\20C\120sec\26.07\A7400002.ARW")) | Should Be $true
        ($destinations -contains (Join-Path $calibrationRoot "SonyA7IV\Source\flats\Canon EF 200 F2.8 MK2\26.07.11 None\A7400003.ARW")) | Should Be $true
        @($plan.Entries | Where-Object { $_.UsedFileTime }).Count | Should Be 3
    }

    It "marks repeated files gray and warns when a populated folder receives a new file" {
        $sourceRoot = Join-Path -Path $TestDrive -ChildPath "repeat\setup"
        $sourceBias = Join-Path -Path $sourceRoot -ChildPath "biases"
        $calibrationRoot = Join-Path -Path $TestDrive -ChildPath "repeat\Calibration"
        New-Item -ItemType Directory -Path $sourceBias -Force | Out-Null
        New-Item -ItemType Directory -Path $calibrationRoot -Force | Out-Null

        $existingName = "Bias_1.0ms_Bin1_2600MC_gain120_20260709-051753_-10.0C_0001.fit"
        $newName = "Bias_1.0ms_Bin1_2600MC_gain120_20260709-051754_-10.0C_0002.fit"
        $sourceExisting = Join-Path -Path $sourceBias -ChildPath $existingName
        Set-Content -LiteralPath $sourceExisting -Value "same" -NoNewline
        Set-Content -LiteralPath (Join-Path $sourceBias $newName) -Value "new" -NoNewline

        $destinationFolder = Join-Path $calibrationRoot "ASI2600MC\Source\biases\gain120\-10C\26.07"
        New-Item -ItemType Directory -Path $destinationFolder -Force | Out-Null
        Copy-Item -LiteralPath $sourceExisting -Destination (Join-Path $destinationFolder $existingName)

        $plan = Get-AsiToPixCalibrationImportPlan `
            -SourcePath $sourceRoot `
            -CalibrationRoot $calibrationRoot

        $plan.ExistingCount | Should Be 1
        $plan.PlannedCount | Should Be 1
        $plan.ConflictCount | Should Be 0
        $plan.AdditionWarnings.Count | Should Be 1
        $plan.AdditionWarnings[0].ExistingCount | Should Be 1
        $plan.AdditionWarnings[0].NewCount | Should Be 1
    }

    It "refuses to overwrite a same-name destination file with a different size" {
        $sourceRoot = Join-Path -Path $TestDrive -ChildPath "conflict\setup"
        $sourceBias = Join-Path -Path $sourceRoot -ChildPath "bias"
        $calibrationRoot = Join-Path -Path $TestDrive -ChildPath "conflict\Calibration"
        New-Item -ItemType Directory -Path $sourceBias -Force | Out-Null
        New-Item -ItemType Directory -Path $calibrationRoot -Force | Out-Null

        $name = "Bias_1.0ms_Bin1_2600MC_gain120_20260709-051753_-10.0C_0001.fit"
        Set-Content -LiteralPath (Join-Path $sourceBias $name) -Value "new source" -NoNewline
        $destinationFolder = Join-Path $calibrationRoot "ASI2600MC\Source\biases\gain120\-10C\26.07"
        New-Item -ItemType Directory -Path $destinationFolder -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $destinationFolder $name) -Value "old" -NoNewline

        $plan = Get-AsiToPixCalibrationImportPlan `
            -SourcePath $sourceRoot `
            -CalibrationRoot $calibrationRoot

        $plan.ConflictCount | Should Be 1
        $plan.PlannedCount | Should Be 0
        $plan.Entries[0].Reason | Should Match "different size"
    }
}

Describe "Calibration import application" {
    $modulePath = Join-Path -Path $PSScriptRoot -ChildPath "..\src\AsiToPix.ImportCalibration.psm1"
    Import-Module $modulePath -Force

    It "copies planned files without changing the source" {
        $sourceRoot = Join-Path -Path $TestDrive -ChildPath "apply\setup"
        $sourceBias = Join-Path -Path $sourceRoot -ChildPath "biases"
        $calibrationRoot = Join-Path -Path $TestDrive -ChildPath "apply\Calibration"
        New-Item -ItemType Directory -Path $sourceBias -Force | Out-Null
        New-Item -ItemType Directory -Path $calibrationRoot -Force | Out-Null
        $name = "Bias_1.0ms_Bin1_2600MC_gain120_20260709-051753_-10.0C_0001.fit"
        $sourceFile = Join-Path -Path $sourceBias -ChildPath $name
        Set-Content -LiteralPath $sourceFile -Value "source contents" -NoNewline
        $sourceWriteTime = (Get-Item -LiteralPath $sourceFile).LastWriteTimeUtc
        $plan = Get-AsiToPixCalibrationImportPlan -SourcePath $sourceRoot -CalibrationRoot $calibrationRoot

        $result = Invoke-AsiToPixCalibrationImportPlan -Plan $plan -Confirm:$false

        $result.CopiedCount | Should Be 1
        (Get-Content -LiteralPath $plan.Entries[0].DestinationPath -Raw) | Should Be "source contents"
        (Get-Content -LiteralPath $sourceFile -Raw) | Should Be "source contents"
        (Get-Item -LiteralPath $sourceFile).LastWriteTimeUtc | Should Be $sourceWriteTime
    }

    It "does not create files in WhatIf mode" {
        $sourceRoot = Join-Path -Path $TestDrive -ChildPath "whatif\setup"
        $sourceBias = Join-Path -Path $sourceRoot -ChildPath "biases"
        $calibrationRoot = Join-Path -Path $TestDrive -ChildPath "whatif\Calibration"
        New-Item -ItemType Directory -Path $sourceBias -Force | Out-Null
        New-Item -ItemType Directory -Path $calibrationRoot -Force | Out-Null
        $name = "Bias_1.0ms_Bin1_2600MC_gain120_20260709-051753_-10.0C_0001.fit"
        Set-Content -LiteralPath (Join-Path $sourceBias $name) -Value "source" -NoNewline
        $plan = Get-AsiToPixCalibrationImportPlan -SourcePath $sourceRoot -CalibrationRoot $calibrationRoot

        $result = Invoke-AsiToPixCalibrationImportPlan -Plan $plan -WhatIf -Confirm:$false

        $result.WhatIfCount | Should Be 1
        Test-Path -LiteralPath $plan.Entries[0].DestinationPath | Should Be $false
    }

    It "copies a Robocopy calibration batch through staging without changing sources" {
        $sourcePath = Join-Path -Path $TestDrive -ChildPath "robocopy-calibration-source"
        $calibrationRoot = Join-Path -Path $TestDrive -ChildPath "robocopy-calibration-target"
        $firstDestinationFolder = Join-Path -Path $calibrationRoot -ChildPath "Camera\Source\darks\gain120\-10C\180sec\26.07"
        $secondDestinationFolder = Join-Path -Path $calibrationRoot -ChildPath "Camera\Source\darks\gain120\-10C\300sec\26.07"
        New-Item -ItemType Directory -Path $sourcePath -Force | Out-Null
        New-Item -ItemType Directory -Path $firstDestinationFolder -Force | Out-Null
        New-Item -ItemType Directory -Path $secondDestinationFolder -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $sourcePath "first.fit") -Value "first calibration" -Encoding ASCII
        Set-Content -LiteralPath (Join-Path $sourcePath "second.fit") -Value "second calibration" -Encoding ASCII
        $env:ASITOPIX_CAL_ROBOCOPY_SOURCE = $sourcePath
        $env:ASITOPIX_CAL_ROBOCOPY_ROOT = $calibrationRoot
        $env:ASITOPIX_CAL_ROBOCOPY_FIRST = Join-Path $firstDestinationFolder "first.fit"
        $env:ASITOPIX_CAL_ROBOCOPY_SECOND = Join-Path $secondDestinationFolder "second.fit"

        InModuleScope AsiToPix.ImportCalibration {
            $firstSource = Get-Item -LiteralPath (Join-Path $env:ASITOPIX_CAL_ROBOCOPY_SOURCE "first.fit")
            $secondSource = Get-Item -LiteralPath (Join-Path $env:ASITOPIX_CAL_ROBOCOPY_SOURCE "second.fit")
            $entries = @(
                [PSCustomObject]@{
                    SourcePath        = $firstSource.FullName
                    SourceLength      = $firstSource.Length
                    DestinationFolder = Split-Path -Path $env:ASITOPIX_CAL_ROBOCOPY_FIRST -Parent
                    DestinationPath   = $env:ASITOPIX_CAL_ROBOCOPY_FIRST
                },
                [PSCustomObject]@{
                    SourcePath        = $secondSource.FullName
                    SourceLength      = $secondSource.Length
                    DestinationFolder = Split-Path -Path $env:ASITOPIX_CAL_ROBOCOPY_SECOND -Parent
                    DestinationPath   = $env:ASITOPIX_CAL_ROBOCOPY_SECOND
                }
            )

            $completed = @(
                Invoke-AsiToPixCalibrationRobocopyImport `
                    -Entry $entries `
                    -CalibrationRoot $env:ASITOPIX_CAL_ROBOCOPY_ROOT `
                    -ThreadCount 2
            )

            $completed.Count | Should Be 2
            Test-Path -LiteralPath $firstSource.FullName -PathType Leaf | Should Be $true
            Test-Path -LiteralPath $secondSource.FullName -PathType Leaf | Should Be $true
            (Get-Item -LiteralPath $env:ASITOPIX_CAL_ROBOCOPY_FIRST).Length | Should Be $firstSource.Length
            (Get-Item -LiteralPath $env:ASITOPIX_CAL_ROBOCOPY_SECOND).Length | Should Be $secondSource.Length
            @(Get-ChildItem -LiteralPath $env:ASITOPIX_CAL_ROBOCOPY_ROOT -Filter ".asitopix-calibration-import-*").Count |
                Should Be 0
        }

        Remove-Item Env:\ASITOPIX_CAL_ROBOCOPY_SOURCE
        Remove-Item Env:\ASITOPIX_CAL_ROBOCOPY_ROOT
        Remove-Item Env:\ASITOPIX_CAL_ROBOCOPY_FIRST
        Remove-Item Env:\ASITOPIX_CAL_ROBOCOPY_SECOND
    }

    It "uses Robocopy only when every calibration source is on the destination UNC share" {
        InModuleScope AsiToPix.ImportCalibration {
            $sameShareEntries = @(
                [PSCustomObject]@{
                    SourcePath = '\\NAS\AstroPhoto\Import\SetupA\darks\first.fit'
                },
                [PSCustomObject]@{
                    SourcePath = '\\nas\astrophoto\Import\SetupB\darks\second.fit'
                }
            )
            $differentShareEntries = @(
                [PSCustomObject]@{
                    SourcePath = '\\NAS\OtherShare\Import\Setup\darks\first.fit'
                }
            )

            Test-AsiToPixCalibrationUseRobocopy `
                -Entry $sameShareEntries `
                -DestinationRoot '\\NAS\AstroPhoto\Calibration' |
                Should Be $true
            Test-AsiToPixCalibrationUseRobocopy `
                -Entry $differentShareEntries `
                -DestinationRoot '\\NAS\AstroPhoto\Calibration' |
                Should Be $false
            Test-AsiToPixCalibrationUseRobocopy `
                -Entry $sameShareEntries `
                -DestinationRoot 'C:\AstroPhoto\Calibration' |
                Should Be $false
        }
    }

    It "selects Robocopy for a same-share NAS calibration plan" {
        InModuleScope AsiToPix.ImportCalibration {
            $plan = [PSCustomObject]@{
                CalibrationRoot = '\\NAS\AstroPhoto\Calibration'
                ExistingCount   = 0
                ConflictCount   = 0
                Entries         = @(
                    [PSCustomObject]@{
                        Status            = "Planned"
                        SourcePath        = '\\NAS\AstroPhoto\Import\Setup\darks\first.fit'
                        SourceLength      = 42
                        DestinationFolder = '\\NAS\AstroPhoto\Calibration\Camera\Source\darks\gain120\-10C\180sec\26.07'
                        DestinationPath   = '\\NAS\AstroPhoto\Calibration\Camera\Source\darks\gain120\-10C\180sec\26.07\first.fit'
                    }
                )
            }

            Mock Test-Path { $false }
            Mock New-Item {}
            Mock Test-AsiToPixCalibrationUseRobocopy { $true }
            Mock Invoke-AsiToPixCalibrationRobocopyImport { @($Entry) }
            Mock Copy-Item {}
            Mock Write-Host {}

            $result = Invoke-AsiToPixCalibrationImportPlan -Plan $plan -Confirm:$false

            $result.CopiedCount | Should Be 1
            $result.CopyEngine | Should Be "Robocopy"
            Assert-MockCalled Invoke-AsiToPixCalibrationRobocopyImport -Times 1 -Exactly -ParameterFilter {
                $CalibrationRoot -eq $plan.CalibrationRoot -and $ThreadCount -eq 4
            }
            Assert-MockCalled Copy-Item -Times 0 -Exactly -ParameterFilter {
                $Destination -like '\\NAS\AstroPhoto\Calibration\*'
            }
        }
    }
}
