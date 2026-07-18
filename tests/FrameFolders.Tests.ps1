$modulePath = Join-Path -Path $PSScriptRoot -ChildPath "..\src\AsiToPix.FrameFolders.psm1"
Import-Module $modulePath -Force

Describe "Shared frame folder conventions" {
    It "accepts singular, plural, and mixed-case names" {
        $cases = @(
            @{ Name = "lIgHt"; Kind = "Light" },
            @{ Name = "LIGHTS"; Kind = "Light" },
            @{ Name = "bIaS"; Kind = "Bias" },
            @{ Name = "BIASES"; Kind = "Bias" },
            @{ Name = "dArK"; Kind = "Dark" },
            @{ Name = "DARKS"; Kind = "Dark" },
            @{ Name = "fLaT"; Kind = "Flat" },
            @{ Name = "FLATS"; Kind = "Flat" },
            @{ Name = "FlatDark"; Kind = "FlatDark" },
            @{ Name = "FLAT-DARKS"; Kind = "FlatDark" }
        )

        foreach ($case in $cases) {
            Test-AsiToPixFrameFolderName -Name $case.Name -Kind $case.Kind | Should Be $true
        }
    }

    It "finds every matching child folder without changing its actual name" {
        $root = Join-Path -Path $TestDrive -ChildPath "frames"
        New-Item -ItemType Directory -Path (Join-Path -Path $root -ChildPath "LiGhT") -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path -Path $root -ChildPath "LIGHTS") -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path -Path $root -ChildPath "Other") -Force | Out-Null

        $folders = @(Get-AsiToPixChildFrameFolder -Path $root -Kind Light)

        $folders.Count | Should Be 2
        @($folders | Where-Object { $_.Name -ceq "LiGhT" }).Count | Should Be 1
        @($folders | Where-Object { $_.Name -ceq "LIGHTS" }).Count | Should Be 1
    }

    It "returns canonical project folder names" {
        Get-AsiToPixCanonicalFrameFolderName -Kind Light | Should Be "lights"
        Get-AsiToPixCanonicalFrameFolderName -Kind Bias | Should Be "biases"
        Get-AsiToPixCanonicalFrameFolderName -Kind Dark | Should Be "darks"
        Get-AsiToPixCanonicalFrameFolderName -Kind Flat | Should Be "flats"
        Get-AsiToPixCanonicalFrameFolderName -Kind FlatDark | Should Be "flat-darks"
    }
}
