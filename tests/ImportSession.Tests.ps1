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

    It "does not match neighboring catalog numbers by a shared catalog prefix" {
        $nameMatches = @(Get-AsiToPixNameMatch -DetectedName "M 16" -Candidates @(
            "M 16 - Eagle nebula",
            "M 17 - Omega nebula"
        ))

        $nameMatches.Count | Should Be 1
        $nameMatches[0].Name | Should Be "M 16 - Eagle nebula"
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

    It "detects the import object from the source folder, not from the FITS name" {
        $objectFolder = Join-Path -Path $TestDrive -ChildPath "APO120 @ 0.8x\Light\M 16"
        $env:ASITOPIX_OBJECT_TEST_FOLDER = $objectFolder
        New-Item -ItemType Directory -Path $objectFolder -Force | Out-Null
        New-Item -ItemType File -Path (Join-Path -Path $objectFolder -ChildPath "Light_FOV_180.0s_Bin1_2600MM_S_gain120_20260713-011758_190deg_-10.0C_APO120_0001.fit") -Force | Out-Null

        InModuleScope AsiToPix.ImportSession {
            Get-AsiToPixDetectedObject -SourcePath $env:ASITOPIX_OBJECT_TEST_FOLDER | Should Be "M 16"
        }

        Remove-Item Env:\ASITOPIX_OBJECT_TEST_FOLDER
    }
}
