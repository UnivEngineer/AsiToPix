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

    It "reads metadata from an OSC light accidentally captured with a Dark prefix" {
        $fileName = "Dark_180.0s_Bin1_2600MC_gain120_20260716-193952_180deg_-10.3C_APO120_0001.fit"

        $info = Get-AsiToPixLightFileInfo -FileName $fileName

        $info.ObjectName | Should Be $null
        $info.ExposureSeconds | Should Be "180.0"
        $info.CameraName | Should Be "ASI2600MC"
        $info.FilterName | Should Be "None"
        $info.CapturedAt | Should Be ([datetime]"2026-07-16T19:39:52")
    }

    It "converts a millisecond Dark exposure used as a light" {
        $fileName = "Dark_20.0ms_Bin1_2600MC_gain120_20260716-193133_230deg_-9.6C_APO120_0001.fit"

        $info = Get-AsiToPixLightFileInfo -FileName $fileName

        $info.ObjectName | Should Be $null
        $info.ExposureSeconds | Should Be "0.02"
    }

    It "applies a supplied filter as metadata without changing the filename" {
        $fileName = "Dark_300.0s_Bin1_2600MM_gain120_20260716-193952_180deg_-10.3C_APO120_0001.fit"

        $info = Resolve-AsiToPixLightFileInfo -FileName $fileName -FilterName "H"

        $info.ObjectName | Should Be $null
        $info.ExposureSeconds | Should Be "300.0"
        $info.FilterName | Should Be "H"
        $fileName | Should Be "Dark_300.0s_Bin1_2600MM_gain120_20260716-193952_180deg_-10.3C_APO120_0001.fit"
    }

    It "normalizes ASI camera suffixes for setup names and maps OSC None to RGB" {
        InModuleScope AsiToPix.ImportSession {
            Get-AsiToPixCameraBaseName -CameraName "ASI2600MM" | Should Be "ASI2600"
            Get-AsiToPixCameraBaseName -CameraName "ASI2600MC" | Should Be "ASI2600"
            ConvertTo-AsiToPixSetupCameraName -SetupName "SQA55 @ 1.0x @ ASI2600MC" | Should Be "SQA55 @ 1.0x @ ASI2600"
            ConvertTo-AsiToPixDestinationFilterName -FilterName "None" -CameraName "ASI2600MC" | Should Be "RGB"
            ConvertTo-AsiToPixDestinationFilterName -FilterName "None" -CameraName "ASI2600MM" | Should Be "None"
        }
    }

    It "sanitizes pasted path segment names before building destination paths" {
        InModuleScope AsiToPix.ImportSession {
            $apostrophe = [char]0x2019
            $rawName = "NGC 6334`t - Cat${apostrophe}s Paw nebula"

            $cleanName = ConvertTo-AsiToPixPathSegment -Value $rawName -ValueName "object name" -Quiet

            $cleanName | Should Be "NGC 6334 - Cat${apostrophe}s Paw nebula"
        }
    }

    It "prompts for a missing mono filter when requested" {
        InModuleScope AsiToPix.ImportSession {
            Mock Read-AsiToPixRequiredValue { "O" }
            $fileName = "Dark_300.0s_Bin1_2600MM_gain120_20260716-193952_180deg_-10.3C_APO120_0001.fit"

            $info = Resolve-AsiToPixLightFileInfo `
                -FileName $fileName `
                -PromptForMissingData

            $info.FilterName | Should Be "O"
            $info.ObjectName | Should Be $null
            Assert-MockCalled Read-AsiToPixRequiredValue -Times 1
        }
    }

    It "keeps the original filename as the incremental import identity" {
        InModuleScope AsiToPix.ImportSession {
            Mock Read-AsiToPixConfirmation { $true }
            $sourcePath = Join-Path -Path $TestDrive -ChildPath "incremental\Light\Ome Cen"
            $astroPhotoRoot = Join-Path -Path $TestDrive -ChildPath "incremental-target"
            New-Item -ItemType Directory -Path $sourcePath -Force | Out-Null
            New-Item -ItemType Directory -Path $astroPhotoRoot -Force | Out-Null
            $fileName = "Dark_180.0s_Bin1_2600MC_gain120_20260716-193952_180deg_-10.3C_APO120_0001.fit"
            $sourceFile = Join-Path -Path $sourcePath -ChildPath $fileName
            New-Item -ItemType File -Path $sourceFile | Out-Null

            $parameters = @{
                SourcePath     = $sourcePath
                AstroPhotoRoot = $astroPhotoRoot
                ObjectName     = "Ome Cen"
                SeasonName     = "2026"
                TelescopeName  = "APO120"
                CameraName     = "ASI2600MC"
                ImportMode     = "Copy"
            }
            Import-AsiToPixSession @parameters
            Import-AsiToPixSession @parameters

            $importedFiles = @(Get-ChildItem -LiteralPath $astroPhotoRoot -File -Recurse)
            $importedFiles.Count | Should Be 1
            $importedFiles[0].Name | Should Be $fileName
            $importedFiles[0].DirectoryName.EndsWith("RGB\26.07.16") | Should Be $true
            (Test-Path -LiteralPath $sourceFile -PathType Leaf) | Should Be $true
        }
    }

    It "splits same-night import folders when one filter contains mixed exposures" {
        $sourcePath = Join-Path -Path $TestDrive -ChildPath "mixed-exp\APO120 @ 0.8x\Light\Helix"
        $astroPhotoRoot = Join-Path -Path $TestDrive -ChildPath "mixed-exp-target"
        New-Item -ItemType Directory -Path $sourcePath -Force | Out-Null
        New-Item -ItemType Directory -Path $astroPhotoRoot -Force | Out-Null
        $file180 = "Light_Helix_180.0s_Bin1_2600MM_O_gain120_20260711-033109_180deg_-10.0C_APO120_0001.fit"
        $file300 = "Light_Helix_300.0s_Bin1_2600MM_O_gain120_20260711-033856_180deg_-10.0C_APO120_0001.fit"
        New-Item -ItemType File -Path (Join-Path -Path $sourcePath -ChildPath $file180) | Out-Null
        New-Item -ItemType File -Path (Join-Path -Path $sourcePath -ChildPath $file300) | Out-Null

        $plan = Get-AsiToPixImportPlan `
            -SourcePath $sourcePath `
            -AstroPhotoRoot $astroPhotoRoot `
            -ObjectName "Helix" `
            -SeasonName "2026" `
            -SetupName "APO120 @ 0.8x @ ASI2600MM" `
            -ImportMode "Copy"

        $folders = @($plan.ParsedFiles | Select-Object -ExpandProperty DestinationNightFolder | Sort-Object)
        $folders | Should Be @("26.07.10-180s", "26.07.10-300s")

        Invoke-AsiToPixImportPlan -Plan $plan

        Test-Path -LiteralPath (Join-Path -Path $plan.GoodRoot -ChildPath "O\26.07.10-180s\$file180") -PathType Leaf | Should Be $true
        Test-Path -LiteralPath (Join-Path -Path $plan.GoodRoot -ChildPath "O\26.07.10-300s\$file300") -PathType Leaf | Should Be $true
    }

    It "keeps plain night folders when a filter has one exposure" {
        $sourcePath = Join-Path -Path $TestDrive -ChildPath "single-exp\APO120 @ 0.8x\Light\Helix"
        $astroPhotoRoot = Join-Path -Path $TestDrive -ChildPath "single-exp-target"
        New-Item -ItemType Directory -Path $sourcePath -Force | Out-Null
        New-Item -ItemType Directory -Path $astroPhotoRoot -Force | Out-Null
        New-Item -ItemType File -Path (Join-Path -Path $sourcePath -ChildPath "Light_Helix_180.0s_Bin1_2600MM_O_gain120_20260711-033109_180deg_-10.0C_APO120_0001.fit") | Out-Null
        New-Item -ItemType File -Path (Join-Path -Path $sourcePath -ChildPath "Light_Helix_180.0s_Bin1_2600MM_O_gain120_20260711-033856_180deg_-10.0C_APO120_0002.fit") | Out-Null

        $plan = Get-AsiToPixImportPlan `
            -SourcePath $sourcePath `
            -AstroPhotoRoot $astroPhotoRoot `
            -ObjectName "Helix" `
            -SeasonName "2026" `
            -SetupName "APO120 @ 0.8x @ ASI2600MM" `
            -ImportMode "Copy"

        @($plan.ParsedFiles | Select-Object -ExpandProperty DestinationNightFolder -Unique) | Should Be @("26.07.10")
    }

    It "finds import sessions grouped by setup and object folders" {
        $importRoot = Join-Path -Path $TestDrive -ChildPath "batch-import"
        $objectFolder = Join-Path -Path $importRoot -ChildPath "APO120 @ 0.8x\Lights\M 16"
        New-Item -ItemType Directory -Path $objectFolder -Force | Out-Null
        New-Item -ItemType File -Path (Join-Path -Path $objectFolder -ChildPath "Light_M 16_180.0s_Bin1_2600MC_gain120_20260716-193952_180deg_-10.3C_APO120_0001.fit") | Out-Null
        New-Item -ItemType File -Path (Join-Path -Path $objectFolder -ChildPath "Light_M 16_180.0s_Bin1_2600MC_gain120_20260716-193952_180deg_-10.3C_APO120_0001_thn.jpg") | Out-Null

        $sessions = @(Find-AsiToPixImportSession -ImportRoot $importRoot)

        $sessions.Count | Should Be 1
        $sessions[0].DetectedSetupName | Should Be "APO120 @ 0.8x"
        $sessions[0].DetectedObject | Should Be "M 16"
        $sessions[0].FileCount | Should Be 1
    }

    It "resolves an object name to a matching default import session" {
        $astroPhotoRoot = Join-Path -Path $TestDrive -ChildPath "object-search-root\AstroPhoto"
        $objectFolder = Join-Path -Path $astroPhotoRoot -ChildPath "Import\APO120 @ 0.8x\Light\Fighting Dragons"
        New-Item -ItemType Directory -Path $objectFolder -Force | Out-Null
        New-Item -ItemType File -Path (Join-Path -Path $objectFolder -ChildPath "Light_FOV_180.0s_Bin1_2600MC_gain120_20260716-193952_180deg_-10.3C_APO120_0001.fit") | Out-Null

        $resolved = Resolve-AsiToPixImportSourcePath -SourcePath "Dragons" -AstroPhotoRoot $astroPhotoRoot

        $resolved.SourcePath | Should Be $objectFolder
        $resolved.AstroPhotoRoot | Should Be $astroPhotoRoot
    }

    It "lets the user choose between multiple matching import sessions" {
        $astroPhotoRoot = Join-Path -Path $TestDrive -ChildPath "multi-object-search-root\AstroPhoto"
        $firstFolder = Join-Path -Path $astroPhotoRoot -ChildPath "Import\APO120 @ 0.8x\Light\Dragons"
        $secondFolder = Join-Path -Path $astroPhotoRoot -ChildPath "Import\SQA55 @ 1.0x\Light\Dragons"
        New-Item -ItemType Directory -Path $firstFolder -Force | Out-Null
        New-Item -ItemType Directory -Path $secondFolder -Force | Out-Null
        New-Item -ItemType File -Path (Join-Path -Path $firstFolder -ChildPath "Light_FOV_180.0s_Bin1_2600MC_gain120_20260716-193952_180deg_-10.3C_APO120_0001.fit") | Out-Null
        New-Item -ItemType File -Path (Join-Path -Path $secondFolder -ChildPath "Light_FOV_180.0s_Bin1_2600MC_gain120_20260717-193952_180deg_-10.3C_SQA55_0001.fit") | Out-Null
        $env:ASITOPIX_SEARCH_ASTRO_ROOT = $astroPhotoRoot
        $env:ASITOPIX_SEARCH_SECOND_FOLDER = $secondFolder

        InModuleScope AsiToPix.ImportSession {
            Mock Read-Host { "2" }

            $resolved = Resolve-AsiToPixImportSourcePath `
                -SourcePath "Dragons" `
                -AstroPhotoRoot $env:ASITOPIX_SEARCH_ASTRO_ROOT

            $resolved.SourcePath | Should Be $env:ASITOPIX_SEARCH_SECOND_FOLDER
            Assert-MockCalled Read-Host -Times 1
        }

        Remove-Item Env:\ASITOPIX_SEARCH_ASTRO_ROOT
        Remove-Item Env:\ASITOPIX_SEARCH_SECOND_FOLDER
    }

    It "builds a reusable import plan from explicit batch-level values" {
        $sourcePath = Join-Path -Path $TestDrive -ChildPath "batch-plan\APO120 @ 0.8x\Light\M 16"
        $astroPhotoRoot = Join-Path -Path $TestDrive -ChildPath "batch-plan-target"
        New-Item -ItemType Directory -Path $sourcePath -Force | Out-Null
        New-Item -ItemType Directory -Path $astroPhotoRoot -Force | Out-Null
        New-Item -ItemType File -Path (Join-Path -Path $sourcePath -ChildPath "Light_M 16_180.0s_Bin1_2600MC_gain120_20260716-193952_180deg_-10.3C_APO120_0001.fit") | Out-Null

        $plan = Get-AsiToPixImportPlan `
            -SourcePath $sourcePath `
            -AstroPhotoRoot $astroPhotoRoot `
            -ObjectName "M 16 - Eagle nebula" `
            -SeasonName "2026" `
            -SetupName "APO120 @ 0.8x @ ASI2600MC" `
            -ImportMode "Copy"

        $plan.ObjectName | Should Be "M 16 - Eagle nebula"
        $plan.SeasonName | Should Be "2026"
        $plan.SetupName | Should Be "APO120 @ 0.8x @ ASI2600"
        $plan.ImportMode | Should Be "Copy"
        @($plan.ParsedFiles).Count | Should Be 1
        $plan.ParsedFiles[0].FilterName | Should Be "RGB"
    }

    It "sanitizes explicit object, season, and setup names in reusable import plans" {
        $sourcePath = Join-Path -Path $TestDrive -ChildPath "sanitize-plan\APO120 @ 0.8x\Light\NGC 6334"
        $astroPhotoRoot = Join-Path -Path $TestDrive -ChildPath "sanitize-plan-target"
        New-Item -ItemType Directory -Path $sourcePath -Force | Out-Null
        New-Item -ItemType Directory -Path $astroPhotoRoot -Force | Out-Null
        New-Item -ItemType File -Path (Join-Path -Path $sourcePath -ChildPath "Light_NGC6334_180.0s_Bin1_2600MC_gain120_20260716-193952_180deg_-10.3C_APO120_0001.fit") | Out-Null
        $apostrophe = [char]0x2019

        $plan = Get-AsiToPixImportPlan `
            -SourcePath $sourcePath `
            -AstroPhotoRoot $astroPhotoRoot `
            -ObjectName "NGC 6334`t - Cat${apostrophe}s Paw nebula" `
            -SeasonName "2026`t" `
            -SetupName "APO120 @ 0.8x @ ASI2600MC`t" `
            -ImportMode "Copy"

        $plan.ObjectName | Should Be "NGC 6334 - Cat${apostrophe}s Paw nebula"
        $plan.SeasonName | Should Be "2026"
        $plan.SetupName | Should Be "APO120 @ 0.8x @ ASI2600"
        $plan.ObjectPath.Contains("`t") | Should Be $false
    }

    It "assigns early morning captures to the previous night" {
        $capturedAt = [datetime]"2026-07-11T05:00:03"

        Get-AsiToPixNightDate -CapturedAt $capturedAt | Should Be "26.07.10"
    }

    It "keeps afternoon captures on the same night start date" {
        $capturedAt = [datetime]"2026-07-11T12:00:00"

        Get-AsiToPixNightDate -CapturedAt $capturedAt | Should Be "26.07.11"
    }

    It "uses copy as the default interactive import mode" {
        InModuleScope AsiToPix.ImportSession {
            Mock Read-Host { "" }

            Read-AsiToPixImportMode | Should Be "Copy"

            Assert-MockCalled Read-Host -Times 1
        }
    }

    It "accepts symlink as an interactive import mode" {
        InModuleScope AsiToPix.ImportSession {
            Mock Read-Host { "2" }

            Read-AsiToPixImportMode | Should Be "Symlink"

            Assert-MockCalled Read-Host -Times 1
        }
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

    It "distinguishes a catalog composition from each individual object" {
        $nameMatches = @(Get-AsiToPixNameMatch -DetectedName "M 8 + M 20" -Candidates @(
            "M 8 - Lagoon nebula",
            "M 8 + M 20 - Lagoon + Trifid nebulae",
            "M 20 - Trifid nebula"
        ))

        $nameMatches.Count | Should Be 1
        $nameMatches[0].Name | Should Be "M 8 + M 20 - Lagoon + Trifid nebulae"
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

    It "discovers camera RAW files and detects lens setup from lowercase lights folders" {
        $rawFolder = Join-Path -Path $TestDrive -ChildPath "Canon EF 200 F2.8 MK2\lights\Rho Oph"
        $env:ASITOPIX_RAW_TEST_FOLDER = $rawFolder
        New-Item -ItemType Directory -Path $rawFolder -Force | Out-Null
        $rawFile = New-Item -ItemType File -Path (Join-Path -Path $rawFolder -ChildPath "A7406786.ARW") -Force
        $rawFile.LastWriteTime = [datetime]"2026-07-13T01:17:58"

        InModuleScope AsiToPix.ImportSession {
            $files = @(Get-AsiToPixSourceLightFile -SourcePath $env:ASITOPIX_RAW_TEST_FOLDER)

            $files.Count | Should Be 1
            $files[0].Name | Should Be "A7406786.ARW"
            Get-AsiToPixDetectedObject -SourcePath $env:ASITOPIX_RAW_TEST_FOLDER | Should Be "Rho Oph"
            Get-AsiToPixDetectedTelescope -SourcePath $env:ASITOPIX_RAW_TEST_FOLDER | Should Be "Canon EF 200 F2.8 MK2"
        }

        Remove-Item Env:\ASITOPIX_RAW_TEST_FOLDER
    }

    It "uses file time metadata for a supported image without a timestamp in its name" {
        $sourcePath = Join-Path -Path $TestDrive -ChildPath "tiff-fallback\APO120 @ 0.8x\lights\M 31"
        $astroPhotoRoot = Join-Path -Path $TestDrive -ChildPath "tiff-fallback-target"
        New-Item -ItemType Directory -Path $sourcePath -Force | Out-Null
        New-Item -ItemType Directory -Path $astroPhotoRoot -Force | Out-Null
        $sourceFile = New-Item -ItemType File -Path (Join-Path -Path $sourcePath -ChildPath "Light_M31_120s_IMG_0001.tiff")
        $sourceFile.LastWriteTime = [datetime]"2026-07-18T01:17:58"

        $plan = Get-AsiToPixImportPlan `
            -SourcePath $sourcePath `
            -AstroPhotoRoot $astroPhotoRoot `
            -ObjectName "M 31" `
            -SeasonName "2026" `
            -TelescopeName "APO120" `
            -CameraName "ASI2600MC" `
            -ImportMode "Copy"

        $plan.ParsedFiles.Count | Should Be 1
        $plan.ParsedFiles[0].File.Name | Should Be "Light_M31_120s_IMG_0001.tiff"
        $plan.ParsedFiles[0].CapturedAt | Should Be ([datetime]"2026-07-18T01:17:58")
        $plan.ParsedFiles[0].NightDate | Should Be "26.07.17"
    }
}
