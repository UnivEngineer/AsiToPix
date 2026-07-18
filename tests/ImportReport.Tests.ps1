Describe "Import report" {
    $modulePath = Join-Path -Path $PSScriptRoot -ChildPath "..\src\AsiToPix.ImportReport.psm1"
    Import-Module $modulePath -Force

    It "maps OSC and mono filter names to report columns" {
        ConvertTo-AsiToPixReportFilter -FilterName "None" | Should Be "RGB"
        ConvertTo-AsiToPixReportFilter -FilterName "Luminance" | Should Be "L"
        ConvertTo-AsiToPixReportFilter -FilterName "Ha" | Should Be "H"
        ConvertTo-AsiToPixReportFilter -FilterName "SII" | Should Be "S"
        ConvertTo-AsiToPixReportFilter -FilterName "OIII" | Should Be "O"
    }

    It "chooses the exposure used by the largest number of subs" {
        $frames = @(
            1..50 | ForEach-Object { [PSCustomObject]@{ ExposureSeconds = [decimal]180 } }
            1..10 | ForEach-Object { [PSCustomObject]@{ ExposureSeconds = [decimal]300 } }
        )

        Get-AsiToPixCharacteristicExposure -Frame $frames | Should Be ([decimal]180)
    }

    It "separates nights at noon and normalizes alternate exposures" {
        $frames = @(
            [PSCustomObject]@{ ExposureSeconds = [decimal]180; NightDate = "26.07.10" }
            [PSCustomObject]@{ ExposureSeconds = [decimal]180; NightDate = "26.07.10" }
            [PSCustomObject]@{ ExposureSeconds = [decimal]300; NightDate = "26.07.10" }
            [PSCustomObject]@{ ExposureSeconds = [decimal]180; NightDate = "26.07.11" }
        )

        Format-AsiToPixExposureExpression -Frame $frames -BaseExposure 180 | Should Be "(2+1)*180+1*300"
    }

    It "groups the same alternate exposure multiplier across nights" {
        $frames = @()
        foreach ($night in @(
            [PSCustomObject]@{ Date = "26.07.10"; BaseCount = 10; AlternateCount = 7 },
            [PSCustomObject]@{ Date = "26.07.11"; BaseCount = 30; AlternateCount = 5 },
            [PSCustomObject]@{ Date = "26.07.12"; BaseCount = 29; AlternateCount = 13 }
        )) {
            $frames += @(1..$night.BaseCount | ForEach-Object {
                [PSCustomObject]@{ ExposureSeconds = [decimal]180; NightDate = $night.Date }
            })
            $frames += @(1..$night.AlternateCount | ForEach-Object {
                [PSCustomObject]@{ ExposureSeconds = [decimal]300; NightDate = $night.Date }
            })
        }

        Format-AsiToPixExposureExpression -Frame $frames -BaseExposure 180 |
            Should Be "(10+30+29)*180+(7+5+13)*300"
    }

    It "formats total integration time beyond 24 hours" {
        Format-AsiToPixIntegrationTime -Seconds 93784.5 | Should Be "26h 03m 4.5s"
    }

    It "formats seconds as rounded H:MM" {
        Format-AsiToPixHourMinute -Seconds 19260 | Should Be "5:21"
        Format-AsiToPixHourMinute -Seconds 90 | Should Be "0:02"
    }

    It "builds a tab-separated report from an Import tree" {
        $objectPath = Join-Path -Path $TestDrive -ChildPath "APO120 @ 0.8x\Light\47 Tuc"
        New-Item -ItemType Directory -Path $objectPath -Force | Out-Null
        @(
            "Light_47 Tuc_180.0s_Bin1_2600MM_L_gain120_20260710-051612_50deg_-10.0C_APO120_0001.fit",
            "Light_NGC 104_300.0s_Bin1_2600MM_L_gain120_20260710-061612_50deg_-10.0C_APO120_0002.fit",
            "Light_47 Tuc_180.0s_Bin1_2600MM_L_gain120_20260711-131612_50deg_-10.0C_APO120_0003.fit"
        ) | ForEach-Object {
            New-Item -ItemType File -Path (Join-Path -Path $objectPath -ChildPath $_) | Out-Null
        }

        $lines = @(Get-AsiToPixImportReportLine -ImportPath $TestDrive)

        $lines[0] | Should Be "APO120`t0.8x"
        $lines[1] | Should Be "Object`tExpo`tRGB`tL`tH`tO`tS"
        $lines[2] | Should Be "47 Tuc`t180`t-`t(1+1)*180+1*300`t-`t-`t-"

        $report = @(Get-AsiToPixImportReport -ImportPath $TestDrive)
        $report[0].FrameCount | Should Be 3
        $report[0].IntegrationSeconds | Should Be ([decimal]660)
        $report[0].LSeconds | Should Be ([decimal]660)

        $summaryLines = @(Get-AsiToPixIntegrationSummaryPrettyLine -Report $report)
        $summaryLines[0] | Should Be "APO120  0.8x"
        $summaryLines[1] | Should Be "Object  L"
        $summaryLines[2] | Should Be "47 Tuc  0:11"
    }

    It "includes misplaced OSC lights with a Dark filename prefix" {
        $importPath = Join-Path -Path $TestDrive -ChildPath "dark-prefix"
        $objectPath = Join-Path -Path $importPath -ChildPath "APO120 @ 0.8x\Light\Ome Cen"
        New-Item -ItemType Directory -Path $objectPath -Force | Out-Null
        $fileName = "Dark_180.0s_Bin1_2600MC_gain120_20260716-193952_180deg_-10.3C_APO120_0001.fit"
        $filePath = Join-Path -Path $objectPath -ChildPath $fileName
        New-Item -ItemType File -Path $filePath | Out-Null

        $report = @(Get-AsiToPixImportReport -ImportPath $importPath)

        $report.Count | Should Be 1
        $report[0].Object | Should Be "Ome Cen"
        $report[0].Exposure | Should Be ([decimal]180)
        $report[0].RGB | Should Be "1*180"
        (Get-Item -LiteralPath $filePath).Name | Should Be $fileName
    }

    It "aligns columns using the longest value" {
        $importPath = Join-Path -Path $TestDrive -ChildPath "pretty"
        $shortPath = Join-Path -Path $importPath -ChildPath "APO120 @ 0.8x\Light\Short"
        $longPath = Join-Path -Path $importPath -ChildPath "APO120 @ 0.8x\Light\Long object"
        New-Item -ItemType Directory -Path $shortPath -Force | Out-Null
        New-Item -ItemType Directory -Path $longPath -Force | Out-Null

        New-Item -ItemType File -Path (Join-Path -Path $shortPath -ChildPath (
            "Light_Short_180.0s_Bin1_2600MM_L_gain120_20260710-051612_0deg_-10.0C_APO120_0001.fit"
        )) | Out-Null
        @(
            "Light_Long_180.0s_Bin1_2600MM_L_gain120_20260710-051612_0deg_-10.0C_APO120_0001.fit",
            "Light_Long_300.0s_Bin1_2600MM_L_gain120_20260710-061612_0deg_-10.0C_APO120_0002.fit"
        ) | ForEach-Object {
            New-Item -ItemType File -Path (Join-Path -Path $longPath -ChildPath $_) | Out-Null
        }

        $lines = @(Get-AsiToPixImportReportPrettyLine -ImportPath $importPath)

        $shortLine = $lines | Where-Object { $_ -match '^Short\s' }
        $longLine = $lines | Where-Object { $_ -match '^Long object\s' }
        $headerLPosition = $lines[1].IndexOf("L")
        $shortLPosition = $shortLine.IndexOf("1*180")
        $longLPosition = $longLine.IndexOf("1*180")
        $headerLPosition | Should Be $shortLPosition
        $headerLPosition | Should Be $longLPosition
    }
}
