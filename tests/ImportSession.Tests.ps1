Describe "ImportSession parsing" {
    $modulePath = Join-Path -Path $PSScriptRoot -ChildPath "..\src\AsiToPix.ImportSession.psm1"
    Import-Module $modulePath -Force

    It "parses ASIAir light file metadata" {
        $fileName = "Light_Helix_300.0s_Bin1_2600MM_H_gain120_20260711-050003_0deg_-10.0C_APO120_0004.fit"

        $info = Get-AsiToPixLightFileInfo -FileName $fileName

        $info.ObjectName | Should Be "Helix"
        $info.ExposureSeconds | Should Be "300.0"
        $info.CameraName | Should Be "ASI2600MM"
        $info.FilterName | Should Be "H"
        $info.TelescopeName | Should Be "APO120"
        $info.CapturedAt | Should Be ([datetime]"2026-07-11T05:00:03")
    }

    It "assigns early morning captures to the previous night" {
        $capturedAt = [datetime]"2026-07-11T05:00:03"

        Get-AsiToPixNightDate -CapturedAt $capturedAt | Should Be "26.07.10"
    }

    It "keeps afternoon captures on the same night start date" {
        $capturedAt = [datetime]"2026-07-11T12:00:00"

        Get-AsiToPixNightDate -CapturedAt $capturedAt | Should Be "26.07.11"
    }

    It "matches short object names to descriptive archive names" {
        $nameMatches = Get-AsiToPixNameMatch -DetectedName "Helix" -Candidates @(
            "M 31",
            "NGC 7293 (Helix nebula)",
            "NGC 7000"
        )

        $nameMatches[0].Name | Should Be "NGC 7293 (Helix nebula)"
    }

    It "returns an empty filename set as a usable object" {
        $emptyFolder = Join-Path -Path $TestDrive -ChildPath "empty"
        $env:ASITOPIX_EMPTY_TEST_FOLDER = $emptyFolder
        New-Item -ItemType Directory -Path $emptyFolder -Force | Out-Null

        InModuleScope AsiToPix.ImportSession {
            $set = Get-AsiToPixFileNameSet -RootPath $env:ASITOPIX_EMPTY_TEST_FOLDER

            $set.GetType().Name | Should Be 'HashSet`1'
            $set.Contains("missing.fit") | Should Be $false
        }

        Remove-Item Env:\ASITOPIX_EMPTY_TEST_FOLDER
    }
}
