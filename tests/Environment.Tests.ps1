Describe "Environment path typo warnings" {
    $modulePath = Join-Path -Path $PSScriptRoot -ChildPath "..\src\AsiToPix.Environment.psm1"
    Import-Module $modulePath -Force

    It "detects Cyrillic C in paths" {
        $path = "C:\AstroPhoto\Calibration\ASI2600MC\Master\biases\gain120\-10$([char]0x0421)\26.07"

        Test-AsiToPixPathHasCyrillicC -Path $path | Should Be $true
    }

    It "does not flag Latin C in paths" {
        $path = "C:\AstroPhoto\Calibration\ASI2600MC\Master\biases\gain120\-10C\26.07"

        Test-AsiToPixPathHasCyrillicC -Path $path | Should Be $false
    }

    It "suggests Latin C replacement" {
        $path = "C:\AstroPhoto\Calibration\ASI2600MC\Master\biases\gain120\-10$([char]0x0421)\26.07"

        ConvertTo-AsiToPixLatinCPath -Path $path | Should Be "C:\AstroPhoto\Calibration\ASI2600MC\Master\biases\gain120\-10C\26.07"
    }

    It "suggests canonical exposure folder units" {
        $path = "C:\AstroPhoto\Calibration\ASI2600MC\Master\darks\Gain120\-10C\60s"

        $issue = Get-AsiToPixPathConventionIssue -Path $path | Where-Object { $_.Kind -eq "FolderToken" }

        $issue.Suggestion | Should Be "C:\AstroPhoto\Calibration\ASI2600MC\Master\darks\Gain120\-10C\60sec"
    }

    It "suggests canonical angle tokens without spaces" {
        $path = "C:\AstroPhoto\Calibration\ASI2600MC\Master\flats\SQA55 @ 1.0x\26.07.08 L 185 deg"

        $issue = Get-AsiToPixPathConventionIssue -Path $path | Where-Object { $_.Kind -eq "FolderToken" }

        $issue.Suggestion | Should Be "C:\AstroPhoto\Calibration\ASI2600MC\Master\flats\SQA55 @ 1.0x\26.07.08 L 185deg"
    }

    It "does not interpret the S flat filter as an exposure unit" {
        $paths = @(
            "C:\AstroPhoto\Calibration\ASI2600MM\Master\flats\APO120 @ 0.8x\26.07.08 S 0deg",
            "C:\AstroPhoto\Calibration\ASI2600MM\Source\flats\APO120 @ 0.8x\25.03.27 S",
            "C:\AstroPhoto\Calibration\ASI2600MM\Source\flats\APO120 @ 0.8x\60s"
        )

        foreach ($path in $paths) {
            $exposureIssues = @(Get-AsiToPixPathConventionIssue -Path $path | Where-Object {
                $_.Message -match "Exposure"
            })

            $exposureIssues.Count | Should Be 0
        }
    }

    It "recognizes singular and mixed-case flat calibration folders" {
        $paths = @(
            "C:\AstroPhoto\Calibration\ASI2600MM\MASTER\fLaT\APO120 @ 0.8x\60s",
            "C:\AstroPhoto\Calibration\ASI2600MM\source\FLATS\APO120 @ 0.8x\60s"
        )

        foreach ($path in $paths) {
            $exposureIssues = @(Get-AsiToPixPathConventionIssue -Path $path | Where-Object {
                $_.Message -match "Exposure"
            })
            $exposureIssues.Count | Should Be 0
        }
    }

    It "suggests temperature folders without spaces" {
        $path = "C:\AstroPhoto\Calibration\ASI2600MC\Master\biases\Gain120\-10 C\26.07"

        $issue = Get-AsiToPixPathConventionIssue -Path $path | Where-Object { $_.Kind -eq "FolderToken" }

        $issue.Suggestion | Should Be "C:\AstroPhoto\Calibration\ASI2600MC\Master\biases\Gain120\-10C\26.07"
    }

    It "warns about legacy dark layout without a gain folder" {
        $path = "C:\AstroPhoto\Calibration\ASI2600MC\Source\darks\-10C\180sec\26.07\Dark_180.0s_Bin1_2600MC_gain120_20260709-063548_185deg_-10.0C_SQA55_0003.fit"

        $issue = Get-AsiToPixPathConventionIssue -Path $path | Where-Object { $_.Kind -eq "LegacyCalibrationLayout" }

        $issue.Suggestion | Should Be "C:\AstroPhoto\Calibration\ASI2600MC\Source\darks\Gain120\-10C\180sec\26.07\Dark_180.0s_Bin1_2600MC_gain120_20260709-063548_185deg_-10.0C_SQA55_0003.fit"
    }

    It "warns about legacy singular mixed-case calibration layouts" {
        $path = "C:\AstroPhoto\Calibration\ASI2600MC\source\DaRk\-10C\180sec\26.07"

        $issue = Get-AsiToPixPathConventionIssue -Path $path | Where-Object { $_.Kind -eq "LegacyCalibrationLayout" }

        $issue | Should Not BeNullOrEmpty
        $issue.Suggestion | Should Match 'Gain<gain>'
    }
}
