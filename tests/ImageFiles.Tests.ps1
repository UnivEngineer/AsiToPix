$modulePath = Join-Path -Path $PSScriptRoot -ChildPath "..\src\AsiToPix.ImageFiles.psm1"
Import-Module $modulePath -Force

Describe "Shared PixInsight image formats" {
    It "recognizes astronomical, bitmap, and compressed FITS formats" {
        foreach ($fileName in @(
            "frame.fit",
            "frame.FITS",
            "frame.fit.fz",
            "frame.fts.gz",
            "frame.xisf",
            "frame.tiff",
            "frame.png",
            "frame.jp2"
        )) {
            Test-AsiToPixSupportedImageFileName -FileName $fileName | Should Be $true
        }
    }

    It "recognizes the PixInsight camera RAW extension families" {
        foreach ($fileName in @(
            "canon.CR3",
            "nikon.NEF",
            "sony.ARW",
            "fuji.RAF",
            "panasonic.RW2",
            "olympus.ORF",
            "pentax.PEF",
            "phase-one.IIQ",
            "hasselblad.3FR",
            "sigma.X3F",
            "generic.RAW"
        )) {
            Test-AsiToPixRawImageFileName -FileName $fileName | Should Be $true
            Test-AsiToPixSupportedImageFileName -FileName $fileName | Should Be $true
        }
    }

    It "rejects unrelated files" {
        Test-AsiToPixSupportedImageFileName -FileName "notes.txt" | Should Be $false
    }

    It "rejects ASIAir thumbnail JPEG files but keeps ordinary JPEG images" {
        Test-AsiToPixThumbnailImageFileName -FileName "Light_M16_180s_0001_thn.jpg" | Should Be $true
        Test-AsiToPixSupportedImageFileName -FileName "Light_M16_180s_0001_thn.jpg" | Should Be $false
        Test-AsiToPixSupportedImageFileName -FileName "LIGHT_M16_180S_0001_THN.JPG" | Should Be $false
        Test-AsiToPixSupportedImageFileName -FileName "Light_M16_180s_0001.jpg" | Should Be $true
    }

    It "removes compound and single image extensions" {
        Get-AsiToPixImageFileStem -FileName "Light_M31_180s_20260718-010203.fts.gz" |
            Should Be "Light_M31_180s_20260718-010203"
        Get-AsiToPixImageFileStem -FileName "Light_M31_180s_20260718-010203.CR3" |
            Should Be "Light_M31_180s_20260718-010203"
    }
}
