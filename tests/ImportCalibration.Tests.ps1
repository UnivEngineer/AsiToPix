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

    It "recognizes only the requested FITS and ARW calibration extensions" {
        Test-AsiToPixSupportedCalibrationFileName -FileName "frame.FIT" | Should Be $true
        Test-AsiToPixSupportedCalibrationFileName -FileName "frame.fits" | Should Be $true
        Test-AsiToPixSupportedCalibrationFileName -FileName "frame.ARW" | Should Be $true
        Test-AsiToPixSupportedCalibrationFileName -FileName "result.fit.fz" | Should Be $false
        Test-AsiToPixSupportedCalibrationFileName -FileName "notes.txt" | Should Be $false
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
}
