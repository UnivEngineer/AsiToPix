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

    It "checks each destination filter and night directory only once per import plan" {
        InModuleScope AsiToPix.ImportSession {
            $destinationRoot = "C:\AsiToPixTest\ASIAir\Helix\2026\APO120 @ ASI2600"
            $plan = [PSCustomObject]@{
                AsiairRoot  = "C:\AsiToPixTest\ASIAir"
                ObjectPath  = "C:\AsiToPixTest\ASIAir\Helix"
                SeasonPath  = "C:\AsiToPixTest\ASIAir\Helix\2026"
                SetupPath   = $destinationRoot
                GoodRoot    = Join-Path -Path $destinationRoot -ChildPath "Good"
                TrashRoot   = Join-Path -Path $destinationRoot -ChildPath "Trash"
                ImportMode  = "Copy"
                SourcePath  = "C:\Import\Helix"
                ParsedFiles = @(
                    [PSCustomObject]@{
                        File                   = [PSCustomObject]@{ Name = "first.fit"; FullName = "C:\Import\Helix\first.fit" }
                        FilterName             = "O"
                        DestinationNightFolder = "26.07.10"
                    },
                    [PSCustomObject]@{
                        File                   = [PSCustomObject]@{ Name = "second.fit"; FullName = "C:\Import\Helix\second.fit" }
                        FilterName             = "O"
                        DestinationNightFolder = "26.07.10"
                    }
                )
            }

            Mock New-AsiToPixDirectory {}
            Mock Copy-Item {}
            Mock Write-Host {}

            $result = Invoke-AsiToPixImportPlan -Plan $plan

            $result.Imported | Should Be 2
            Assert-MockCalled Copy-Item -Times 2 -Exactly
            Assert-MockCalled New-AsiToPixDirectory -Times 10 -Exactly
            Assert-MockCalled New-AsiToPixDirectory -Times 1 -Exactly -ParameterFilter {
                $Path -eq (Join-Path -Path $plan.GoodRoot -ChildPath "O\26.07.10")
            }
            Assert-MockCalled New-AsiToPixDirectory -Times 1 -Exactly -ParameterFilter {
                $Path -eq (Join-Path -Path $plan.TrashRoot -ChildPath "O\26.07.10")
            }
        }
    }

    It "recognizes matching UNC shares as a fast NAS copy location" {
        InModuleScope AsiToPix.ImportSession {
            Get-AsiToPixNetworkLocationKey -Path '\\NAS\AstroPhoto\Import\M 31\first.fit' |
                Should Be '\\nas\astrophoto'
            Get-AsiToPixNetworkLocationKey -Path 'C:\AstroPhoto\Import\M 31\first.fit' |
                Should Be ""

            $matchingItems = @(
                [PSCustomObject]@{
                    Entry = [PSCustomObject]@{
                        File = [PSCustomObject]@{
                            Name     = "first.fit"
                            FullName = '\\NAS\AstroPhoto\Import\M 31\first.fit'
                        }
                    }
                }
            )

            Test-AsiToPixUseRobocopy `
                -WorkItem $matchingItems `
                -DestinationRoot '\\nas\AstroPhoto\ASIAir\M 31\2026\Setup' |
                Should Be $true
            Test-AsiToPixUseRobocopy `
                -WorkItem $matchingItems `
                -DestinationRoot '\\nas\OtherShare\ASIAir\M 31\2026\Setup' |
                Should Be $false
        }
    }

    It "copies a Robocopy batch through staging without modifying source files" {
        $sourcePath = Join-Path -Path $TestDrive -ChildPath "robocopy-source"
        $setupPath = Join-Path -Path $TestDrive -ChildPath "robocopy-target"
        $firstDestinationFolder = Join-Path -Path $setupPath -ChildPath "Good\L\26.07.10"
        $secondDestinationFolder = Join-Path -Path $setupPath -ChildPath "Good\O\26.07.10"
        New-Item -ItemType Directory -Path $sourcePath -Force | Out-Null
        New-Item -ItemType Directory -Path $firstDestinationFolder -Force | Out-Null
        New-Item -ItemType Directory -Path $secondDestinationFolder -Force | Out-Null
        Set-Content -LiteralPath (Join-Path -Path $sourcePath -ChildPath "first.fit") -Value "first frame" -Encoding ASCII
        Set-Content -LiteralPath (Join-Path -Path $sourcePath -ChildPath "second.fit") -Value "second frame" -Encoding ASCII
        $env:ASITOPIX_ROBOCOPY_SOURCE = $sourcePath
        $env:ASITOPIX_ROBOCOPY_SETUP = $setupPath
        $env:ASITOPIX_ROBOCOPY_FIRST_DESTINATION = Join-Path -Path $firstDestinationFolder -ChildPath "first.fit"
        $env:ASITOPIX_ROBOCOPY_SECOND_DESTINATION = Join-Path -Path $secondDestinationFolder -ChildPath "second.fit"

        InModuleScope AsiToPix.ImportSession {
            $firstFile = Get-Item -LiteralPath (Join-Path -Path $env:ASITOPIX_ROBOCOPY_SOURCE -ChildPath "first.fit")
            $secondFile = Get-Item -LiteralPath (Join-Path -Path $env:ASITOPIX_ROBOCOPY_SOURCE -ChildPath "second.fit")
            $workItems = @(
                [PSCustomObject]@{
                    Entry = [PSCustomObject]@{
                        File       = $firstFile
                        FilterName = "L"
                    }
                    SourceDirectory        = $env:ASITOPIX_ROBOCOPY_SOURCE
                    DestinationFile        = $env:ASITOPIX_ROBOCOPY_FIRST_DESTINATION
                    DestinationNightFolder = "26.07.10"
                },
                [PSCustomObject]@{
                    Entry = [PSCustomObject]@{
                        File       = $secondFile
                        FilterName = "O"
                    }
                    SourceDirectory        = $env:ASITOPIX_ROBOCOPY_SOURCE
                    DestinationFile        = $env:ASITOPIX_ROBOCOPY_SECOND_DESTINATION
                    DestinationNightFolder = "26.07.10"
                }
            )

            $completed = @(
                Invoke-AsiToPixRobocopyImport `
                    -WorkItem $workItems `
                    -SetupPath $env:ASITOPIX_ROBOCOPY_SETUP `
                    -ThreadCount 2
            )

            $completed.Count | Should Be 2
            Test-Path -LiteralPath $firstFile.FullName -PathType Leaf | Should Be $true
            Test-Path -LiteralPath $secondFile.FullName -PathType Leaf | Should Be $true
            (Get-Item -LiteralPath $env:ASITOPIX_ROBOCOPY_FIRST_DESTINATION).Length | Should Be $firstFile.Length
            (Get-Item -LiteralPath $env:ASITOPIX_ROBOCOPY_SECOND_DESTINATION).Length | Should Be $secondFile.Length
            @(Get-ChildItem -LiteralPath $env:ASITOPIX_ROBOCOPY_SETUP -Filter ".asitopix-import-*").Count |
                Should Be 0
        }

        Remove-Item Env:\ASITOPIX_ROBOCOPY_SOURCE
        Remove-Item Env:\ASITOPIX_ROBOCOPY_SETUP
        Remove-Item Env:\ASITOPIX_ROBOCOPY_FIRST_DESTINATION
        Remove-Item Env:\ASITOPIX_ROBOCOPY_SECOND_DESTINATION
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

    It "prompts again after an interactive object-name typo" {
        $astroPhotoRoot = Join-Path -Path $TestDrive -ChildPath "object-typo-root\AstroPhoto"
        $objectFolder = Join-Path -Path $astroPhotoRoot -ChildPath "Import\Canon EF 200\Light\Butterfly"
        New-Item -ItemType Directory -Path $objectFolder -Force | Out-Null
        New-Item -ItemType File -Path (Join-Path -Path $objectFolder -ChildPath "A7406786.ARW") | Out-Null
        $env:ASITOPIX_TYPO_ASTRO_ROOT = $astroPhotoRoot
        $env:ASITOPIX_TYPO_OBJECT_FOLDER = $objectFolder

        InModuleScope AsiToPix.ImportSession {
            Mock Read-Host { "Butterfly" }
            Mock Write-Host {}

            $resolved = Read-AsiToPixImportSource `
                -InitialValue "betterfly" `
                -AstroPhotoRoot $env:ASITOPIX_TYPO_ASTRO_ROOT

            $resolved.SourcePath | Should Be $env:ASITOPIX_TYPO_OBJECT_FOLDER
            Assert-MockCalled Read-Host -Times 1 -Exactly
            Assert-MockCalled Write-Host -Times 1 -Exactly -ParameterFilter {
                $Object -match "^\[!\] No import object folder matching 'betterfly'" -and
                    $ForegroundColor -eq "Red"
            }
        }

        Remove-Item Env:\ASITOPIX_TYPO_ASTRO_ROOT
        Remove-Item Env:\ASITOPIX_TYPO_OBJECT_FOLDER
    }

    It "uses retry handling only for interactively prompted source input" {
        $entryScriptPath = Join-Path -Path $PSScriptRoot -ChildPath "..\ImportSession.ps1"
        $entryScriptText = Get-Content -LiteralPath $entryScriptPath -Raw

        $entryScriptText | Should Match '\$sourcePathWasProvided = -not \[string\]::IsNullOrWhiteSpace\(\$SourcePath\)'
        $entryScriptText | Should Match 'Read-AsiToPixImportSource -InitialValue \$SourcePath'
        $entryScriptText | Should Match 'Resolve-AsiToPixImportSourcePath -SourcePath \$SourcePath'
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

    It "reads files only from object folders matching an object-name search" {
        $astroPhotoRoot = Join-Path -Path $TestDrive -ChildPath "filtered-object-search\AstroPhoto"
        $targetFolder = Join-Path -Path $astroPhotoRoot -ChildPath "Import\Canon EF 200\Light\Butterfly"
        $unrelatedFolder = Join-Path -Path $astroPhotoRoot -ChildPath "Import\APO120 @ 0.8x\Light\Andromeda"
        New-Item -ItemType Directory -Path $targetFolder -Force | Out-Null
        New-Item -ItemType Directory -Path $unrelatedFolder -Force | Out-Null
        $env:ASITOPIX_FILTERED_SEARCH_ROOT = $astroPhotoRoot
        $env:ASITOPIX_FILTERED_TARGET = $targetFolder
        $env:ASITOPIX_FILTERED_UNRELATED = $unrelatedFolder

        InModuleScope AsiToPix.ImportSession {
            Mock Get-AsiToPixSourceLightFile {
                [PSCustomObject]@{ Name = "A7406786.ARW" }
            }

            $resolved = Resolve-AsiToPixImportSourcePath `
                -SourcePath "Butterfly" `
                -AstroPhotoRoot $env:ASITOPIX_FILTERED_SEARCH_ROOT

            $resolved.SourcePath | Should Be $env:ASITOPIX_FILTERED_TARGET
            Assert-MockCalled Get-AsiToPixSourceLightFile -Times 1 -Exactly -ParameterFilter {
                $SourcePath -eq $env:ASITOPIX_FILTERED_TARGET
            }
            Assert-MockCalled Get-AsiToPixSourceLightFile -Times 0 -Exactly -ParameterFilter {
                $SourcePath -eq $env:ASITOPIX_FILTERED_UNRELATED
            }
        }

        Remove-Item Env:\ASITOPIX_FILTERED_SEARCH_ROOT
        Remove-Item Env:\ASITOPIX_FILTERED_TARGET
        Remove-Item Env:\ASITOPIX_FILTERED_UNRELATED
    }

    It "keeps Robocopy progress in one banner without persistent batch lines" {
        $sourcePath = Join-Path -Path $TestDrive -ChildPath "robocopy-progress-source"
        $setupPath = Join-Path -Path $TestDrive -ChildPath "robocopy-progress-target"
        $destinationFolder = Join-Path -Path $setupPath -ChildPath "Good\L\26.07.10"
        New-Item -ItemType Directory -Path $sourcePath -Force | Out-Null
        New-Item -ItemType Directory -Path $destinationFolder -Force | Out-Null
        1..17 | ForEach-Object {
            Set-Content `
                -LiteralPath (Join-Path -Path $sourcePath -ChildPath ("frame-{0:00}.fit" -f $_)) `
                -Value ("frame {0}" -f $_) `
                -Encoding ASCII
        }
        $env:ASITOPIX_PROGRESS_SOURCE = $sourcePath
        $env:ASITOPIX_PROGRESS_SETUP = $setupPath
        $env:ASITOPIX_PROGRESS_DESTINATION = $destinationFolder

        InModuleScope AsiToPix.ImportSession {
            $workItems = foreach ($sourceFile in Get-ChildItem -LiteralPath $env:ASITOPIX_PROGRESS_SOURCE -File) {
                [PSCustomObject]@{
                    Entry = [PSCustomObject]@{
                        File       = $sourceFile
                        FilterName = "L"
                    }
                    SourceDirectory        = $env:ASITOPIX_PROGRESS_SOURCE
                    DestinationFile        = Join-Path -Path $env:ASITOPIX_PROGRESS_DESTINATION -ChildPath $sourceFile.Name
                    DestinationNightFolder = "26.07.10"
                }
            }

            Mock Invoke-AsiToPixRobocopyBatch {
                foreach ($name in $FileName) {
                    [System.IO.File]::Copy(
                        (Join-Path -Path $SourceDirectory -ChildPath $name),
                        (Join-Path -Path $DestinationDirectory -ChildPath $name)
                    )
                }
            }
            Mock Write-Host {}
            Mock Write-Progress {}

            $completed = @(
                Invoke-AsiToPixRobocopyImport `
                    -WorkItem @($workItems) `
                    -SetupPath $env:ASITOPIX_PROGRESS_SETUP `
                    -ThreadCount 4
            )

            $completed.Count | Should Be 17
            Assert-MockCalled Invoke-AsiToPixRobocopyBatch -Times 2 -Exactly
            Assert-MockCalled Write-Host -Times 0 -Exactly -ParameterFilter {
                $Object -match '^\s+\[copy\]'
            }
            Assert-MockCalled Write-Progress -Times 3 -Exactly -ParameterFilter {
                $Activity -eq "Copying lights from NAS"
            }
            Assert-MockCalled Write-Progress -Times 2 -Exactly -ParameterFilter {
                $Status -match '^\d+/17 files \(\d+%\), \d+([,.]\d+)? MB/s avg, elapsed \d{2}:\d{2}:\d{2} \(Robocopy /MT:4\)$'
            }
        }

        Remove-Item Env:\ASITOPIX_PROGRESS_SOURCE
        Remove-Item Env:\ASITOPIX_PROGRESS_SETUP
        Remove-Item Env:\ASITOPIX_PROGRESS_DESTINATION
    }

    It "selects the fast copy engine for a same-share NAS import plan" {
        InModuleScope AsiToPix.ImportSession {
            $destinationRoot = '\\NAS\AstroPhoto\ASIAir\Helix\2026\APO120 @ ASI2600'
            $plan = [PSCustomObject]@{
                AsiairRoot  = '\\NAS\AstroPhoto\ASIAir'
                ObjectPath  = '\\NAS\AstroPhoto\ASIAir\Helix'
                SeasonPath  = '\\NAS\AstroPhoto\ASIAir\Helix\2026'
                SetupPath   = $destinationRoot
                GoodRoot    = Join-Path -Path $destinationRoot -ChildPath "Good"
                TrashRoot   = Join-Path -Path $destinationRoot -ChildPath "Trash"
                ImportMode  = "Copy"
                SourcePath  = '\\NAS\AstroPhoto\Import\Helix'
                ParsedFiles = @(
                    [PSCustomObject]@{
                        File = [PSCustomObject]@{
                            Name     = "first.fit"
                            FullName = '\\NAS\AstroPhoto\Import\Helix\first.fit'
                            Length   = 42
                        }
                        FilterName             = "O"
                        DestinationNightFolder = "26.07.10"
                    }
                )
            }

            Mock New-AsiToPixDirectory {}
            Mock Test-Path { $false }
            Mock Test-AsiToPixUseRobocopy { $true }
            Mock Invoke-AsiToPixRobocopyImport { @($WorkItem) }
            Mock Copy-Item {}
            Mock Write-Host {}

            $result = Invoke-AsiToPixImportPlan -Plan $plan

            $result.Imported | Should Be 1
            $result.CopyEngine | Should Be "Robocopy"
            Assert-MockCalled Test-AsiToPixUseRobocopy -Times 1 -Exactly
            Assert-MockCalled Invoke-AsiToPixRobocopyImport -Times 1 -Exactly -ParameterFilter {
                $SetupPath -eq $plan.SetupPath -and $ThreadCount -eq 4
            }
            Assert-MockCalled Copy-Item -Times 0 -Exactly -ParameterFilter {
                $Destination -like '\\NAS\AstroPhoto\*'
            }
        }
    }

    It "recognizes a mapped network drive as its backing UNC share" {
        InModuleScope AsiToPix.ImportSession {
            Mock Get-PSDrive {
                [PSCustomObject]@{
                    DisplayRoot = '\\NAS\AstroPhoto'
                }
            }

            Get-AsiToPixNetworkLocationKey -Path 'Z:\AstroPhoto\Import\M 31\first.fit' |
                Should Be '\\nas\astrophoto'
            Assert-MockCalled Get-PSDrive -Times 1 -Exactly -ParameterFilter {
                $Name -eq "Z" -and $PSProvider -eq "FileSystem"
            }
        }
    }

}
