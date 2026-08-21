$importSessionModulePath = Join-Path -Path $PSScriptRoot -ChildPath "..\src\AsiToPix.ImportSession.psm1"
Import-Module $importSessionModulePath -Force

$importReportModulePath = Join-Path -Path $PSScriptRoot -ChildPath "..\src\AsiToPix.ImportReport.psm1"
Import-Module $importReportModulePath -Force

$modulePath = Join-Path -Path $PSScriptRoot -ChildPath "..\src\AsiToPix.Paths.psm1"
Import-Module $modulePath -Force

Describe "Astrophotography root aliases" {
    It "discovers AstroPhoto and Astro roots on filesystem drives" {
        $driveRoot = Join-Path -Path $TestDrive -ChildPath "root-discovery"
        $drive = [PSCustomObject]@{ Root = $driveRoot }
        $astroPhotoRoot = Join-Path -Path $driveRoot -ChildPath "AstroPhoto"
        $astroRoot = Join-Path -Path $driveRoot -ChildPath "Astro"
        New-Item -ItemType Directory -Path $astroPhotoRoot, $astroRoot -Force | Out-Null

        $candidates = @(Get-AsiToPixAstroRootCandidate -FileSystemDrive @($drive))

        $candidates.Count | Should Be 2
        ($candidates -contains $astroPhotoRoot) | Should Be $true
        ($candidates -contains $astroRoot) | Should Be $true
    }

    It "discovers child folders below both root aliases" {
        $driveRoot = Join-Path -Path $TestDrive -ChildPath "child-discovery"
        $drive = [PSCustomObject]@{ Root = $driveRoot }
        $astroPhotoImport = Join-Path -Path $driveRoot -ChildPath "AstroPhoto\Import"
        $astroImport = Join-Path -Path $driveRoot -ChildPath "Astro\Import"
        New-Item -ItemType Directory -Path $astroPhotoImport, $astroImport -Force | Out-Null

        $candidates = @(
            Get-AsiToPixAstroRootChildCandidate -ChildPath "Import" -FileSystemDrive @($drive)
        )

        $candidates.Count | Should Be 2
        ($candidates -contains (Resolve-Path -LiteralPath $astroPhotoImport).ProviderPath) | Should Be $true
        ($candidates -contains (Resolve-Path -LiteralPath $astroImport).ProviderPath) | Should Be $true
    }

    It "ignores missing root aliases" {
        $driveRoot = Join-Path -Path $TestDrive -ChildPath "missing-aliases"
        $drive = [PSCustomObject]@{ Root = $driveRoot }
        $unrelatedRoot = Join-Path -Path $driveRoot -ChildPath "Other"
        New-Item -ItemType Directory -Path $unrelatedRoot -Force | Out-Null

        @(Get-AsiToPixAstroRootCandidate -FileSystemDrive @($drive)).Count | Should Be 0
    }

    It "can require a role-specific choice even when only one root is discovered" {
        $rootPath = Join-Path -Path $TestDrive -ChildPath "single-choice\AstroPhoto"
        New-Item -ItemType Directory -Path $rootPath -Force | Out-Null
        $env:ASITOPIX_SINGLE_CHOICE_ROOT = $rootPath
        $env:ASITOPIX_SELECTION_PROMPT = "Select SOURCE root folder (ASIAir and Calibration)"

        InModuleScope AsiToPix.Paths {
            Mock Get-AsiToPixAstroRootCandidate {
                @($env:ASITOPIX_SINGLE_CHOICE_ROOT)
            }
            Mock Read-Host { "0" }
            Mock Write-Host {}

            $resolved = Resolve-AstroPhotoRoot `
                -Purpose "CreateProject source" `
                -SelectionPrompt $env:ASITOPIX_SELECTION_PROMPT `
                -AlwaysPrompt

            $resolved | Should Be (Resolve-Path -LiteralPath $env:ASITOPIX_SINGLE_CHOICE_ROOT).ProviderPath
            Assert-MockCalled Read-Host -Times 1 -Exactly -ParameterFilter {
                $Prompt -eq "Enter root index or a full root path"
            }
            Assert-MockCalled Write-Host -Times 1 -Exactly -ParameterFilter {
                $Object -eq "`n$($env:ASITOPIX_SELECTION_PROMPT):"
            }
        }

        Remove-Item Env:\ASITOPIX_SINGLE_CHOICE_ROOT
        Remove-Item Env:\ASITOPIX_SELECTION_PROMPT
    }

    It "uses shared aliases for default import session discovery" {
        InModuleScope AsiToPix.ImportSession {
            Mock Get-AsiToPixAstroRootChildCandidate {
                @("C:\Astro\Import", "Z:\AstroPhoto\Import")
            }

            $candidates = @(Get-AsiToPixDefaultImportRoot)

            $candidates.Count | Should Be 2
            ($candidates -contains "C:\Astro\Import") | Should Be $true
            ($candidates -contains "Z:\AstroPhoto\Import") | Should Be $true
            Assert-MockCalled Get-AsiToPixAstroRootChildCandidate -Times 1 -Exactly -ParameterFilter {
                $ChildPath -eq "Import"
            }
        }
    }

    It "uses shared aliases for import report discovery" {
        InModuleScope AsiToPix.ImportReport {
            Mock Get-AsiToPixAstroRootChildCandidate {
                @("C:\Astro\Import", "Z:\AstroPhoto\Import")
            }

            $candidates = @(Get-AsiToPixImportRoot)

            $candidates.Count | Should Be 2
            ($candidates -contains "C:\Astro\Import") | Should Be $true
            ($candidates -contains "Z:\AstroPhoto\Import") | Should Be $true
            Assert-MockCalled Get-AsiToPixAstroRootChildCandidate -Times 1 -Exactly -ParameterFilter {
                $ChildPath -eq "Import"
            }
        }
    }
}
