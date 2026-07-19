Describe "Import report" {
    $modulePath = Join-Path -Path $PSScriptRoot -ChildPath "..\src\AsiToPix.ImportReport.psm1"
    Import-Module $modulePath -Force

    It "maps OSC and mono filter names to report columns" {
        ConvertTo-AsiToPixReportFilter -FilterName "None" | Should Be "RGB"
        ConvertTo-AsiToPixReportFilter -FilterName "Luminance" | Should Be "L"
        ConvertTo-AsiToPixReportFilter -FilterName "Ha" | Should Be "H"
        ConvertTo-AsiToPixReportFilter -FilterName "SII" | Should Be "S"
        ConvertTo-AsiToPixReportFilter -FilterName "OIII" | Should Be "O"

        ConvertTo-AsiToPixTsvFilter -FilterName "None" | Should Be "RGB"
        ConvertTo-AsiToPixTsvFilter -FilterName "L" | Should Be "L"
        ConvertTo-AsiToPixTsvFilter -FilterName "R" | Should Be "R"
        ConvertTo-AsiToPixTsvFilter -FilterName "G" | Should Be "G"
        ConvertTo-AsiToPixTsvFilter -FilterName "B" | Should Be "B"
        ConvertTo-AsiToPixTsvFilter -FilterName "UHC" | Should Be "HO"
        ConvertTo-AsiToPixTsvFilter -FilterName "SO" | Should Be "SO"
        ConvertTo-AsiToPixTsvFilter -FilterName "H" | Should Be "Ha"
        ConvertTo-AsiToPixTsvFilter -FilterName "O" | Should Be "OIII"
        ConvertTo-AsiToPixTsvFilter -FilterName "S" | Should Be "SII"
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

    It "formats TSV exposure formulas and empty cells for Google Sheets" {
        ConvertTo-AsiToPixTsvExpression -Expression "(84+59)*60+1*120+1*180" |
            Should Be "=(84+59)*60+1*120+1*180"
        ConvertTo-AsiToPixTsvExpression -Expression "-" | Should Be ""
        ConvertTo-AsiToPixTsvExpression -Expression "" | Should Be ""
    }

    It "builds a tab-separated report from an Import tree" {
        $astroPhotoRoot = Join-Path -Path $TestDrive -ChildPath "report-tree\AstroPhoto"
        $importPath = Join-Path -Path $astroPhotoRoot -ChildPath "Import"
        $objectPath = Join-Path -Path $importPath -ChildPath "APO120 @ 0.8x\Light\47 Tuc"
        $archiveObjectPath = Join-Path -Path $astroPhotoRoot -ChildPath "ASIAir\NGC 104 - 47 Tuc"
        New-Item -ItemType Directory -Path $objectPath -Force | Out-Null
        New-Item -ItemType Directory -Path $archiveObjectPath -Force | Out-Null
        @(
            "Light_47 Tuc_180.0s_Bin1_2600MM_L_gain120_20260710-051612_50deg_-10.0C_APO120_0001.fit",
            "Light_NGC 104_300.0s_Bin1_2600MM_L_gain120_20260710-061612_50deg_-10.0C_APO120_0002.fit",
            "Light_47 Tuc_180.0s_Bin1_2600MM_L_gain120_20260711-131612_50deg_-10.0C_APO120_0003.fit",
            "Light_47 Tuc_180.0s_Bin1_2600MM_L_gain120_20260711-131612_50deg_-10.0C_APO120_0003_thn.jpg"
        ) | ForEach-Object {
            New-Item -ItemType File -Path (Join-Path -Path $objectPath -ChildPath $_) | Out-Null
        }

        $lines = @(Get-AsiToPixImportReportLine -ImportPath $importPath)

        $lines[0] | Should Be "APO120`t0.8x"
        $lines[1] | Should Be (@("Catalog number", "Name", "Exposure", "RGB", "L", "R", "G", "B", "HO", "SO", "Ha", "OIII", "SII") -join "`t")
        $lines[2] | Should Be (@("NGC 104", "47 Tuc", "180", "", "=(1+1)*180+1*300", "", "", "", "", "", "", "", "") -join "`t")

        $report = @(Get-AsiToPixImportReport -ImportPath $importPath)
        $report[0].FrameCount | Should Be 3
        $report[0].IntegrationSeconds | Should Be ([decimal]660)
        $report[0].LSeconds | Should Be ([decimal]660)

        $summaryLines = @(Get-AsiToPixIntegrationSummaryPrettyLine -Report $report)
        $summaryLines[0] | Should Be "APO120  0.8x"
        $summaryLines[1] | Should Be "Object  L"
        $summaryLines[2] | Should Be "47 Tuc  0:11"
    }

    It "emits TSV as one clipboard-safe multiline string" {
        $astroPhotoRoot = Join-Path -Path $TestDrive -ChildPath "clipboard-tsv\AstroPhoto"
        $importPath = Join-Path -Path $astroPhotoRoot -ChildPath "Import"
        $objectPath = Join-Path -Path $importPath -ChildPath "APO120 @ 0.8x\Light\M 16"
        $archiveObjectPath = Join-Path -Path $astroPhotoRoot -ChildPath "ASIAir\M 16 - Eagle nebula"
        New-Item -ItemType Directory -Path $objectPath -Force | Out-Null
        New-Item -ItemType Directory -Path $archiveObjectPath -Force | Out-Null
        New-Item -ItemType File -Path (Join-Path -Path $objectPath -ChildPath "Light_M16_180.0s_Bin1_2600MM_H_gain120_20260710-051612_0deg_-10.0C_APO120_0001.fit") | Out-Null
        $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath "..\Get-ImportReport.ps1"

        $output = @(& $scriptPath -ImportPath $importPath -Tsv)

        $output.Count | Should Be 1
        $header = @("Catalog number", "Name", "Exposure", "RGB", "L", "R", "G", "B", "HO", "SO", "Ha", "OIII", "SII") -join "`t"
        $dataRow = @("M 16", "Eagle nebula", "180", "", "", "", "", "", "", "", "=1*180", "", "") -join "`t"
        $output[0] | Should Be "APO120`t0.8x`r`n$header`r`n$dataRow"
    }

    It "uses the import folder name for both TSV name columns when archive matching is ambiguous" {
        $astroPhotoRoot = Join-Path -Path $TestDrive -ChildPath "ambiguous-name\AstroPhoto"
        $importPath = Join-Path -Path $astroPhotoRoot -ChildPath "Import"
        $objectPath = Join-Path -Path $importPath -ChildPath "APO120 @ 0.8x\Light\Eagle"
        New-Item -ItemType Directory -Path $objectPath -Force | Out-Null
        New-Item -ItemType File -Path (Join-Path -Path $objectPath -ChildPath "Light_Eagle_180.0s_Bin1_2600MM_H_gain120_20260710-051612_0deg_-10.0C_APO120_0001.fit") | Out-Null
        New-Item -ItemType Directory -Path (Join-Path -Path $astroPhotoRoot -ChildPath "ASIAir\M 16 - Eagle nebula") -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path -Path $astroPhotoRoot -ChildPath "ASIAir\IC 4703 - Eagle nebula") -Force | Out-Null

        $index = Get-AsiToPixArchiveObjectIndex -ImportRoot $importPath
        $resolution = Resolve-AsiToPixTsvObjectName -ObjectName "Eagle" -ArchiveIndex $index

        $resolution.IsResolved | Should Be $false
        $resolution.CatalogNumber | Should Be "Eagle"
        $resolution.Name | Should Be "Eagle"
        $resolution.Warning | Should Match "Ambiguous"
        $resolution.Warning | Should Match "M 16 - Eagle nebula"
        $resolution.Warning | Should Match "IC 4703 - Eagle nebula"

        $warnings = @()
        $lines = @(Get-AsiToPixImportReportLine -ImportPath $importPath -WarningVariable warnings -WarningAction SilentlyContinue)
        $warnings.Count | Should Be 1
        [string]$warnings[0] | Should Match "Ambiguous"
        $lines[2] | Should Be (@("Eagle", "Eagle", "180", "", "", "", "", "", "", "", "=1*180", "", "") -join "`t")
    }

    It "resolves a catalog composition without matching an individual object" {
        $astroPhotoRoot = Join-Path -Path $TestDrive -ChildPath "composition-name\AstroPhoto"
        $importPath = Join-Path -Path $astroPhotoRoot -ChildPath "Import"
        New-Item -ItemType Directory -Path $importPath -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path -Path $astroPhotoRoot -ChildPath "ASIAir\M 8 - Lagoon nebula") -Force | Out-Null
        New-Item -ItemType Directory -Path (Join-Path -Path $astroPhotoRoot -ChildPath "ASIAir\M 8 + M 20 - Lagoon + Trifid nebulae") -Force | Out-Null

        $index = Get-AsiToPixArchiveObjectIndex -ImportRoot $importPath
        $resolution = Resolve-AsiToPixTsvObjectName -ObjectName "M 8 + M 20" -ArchiveIndex $index

        $resolution.IsResolved | Should Be $true
        $resolution.CatalogNumber | Should Be "M 8 + M 20"
        $resolution.Name | Should Be "Lagoon + Trifid nebulae"
        $resolution.Warning | Should Be ""
    }

    It "keeps every full TSV filter separate while the console report stays compact" {
        $importPath = Join-Path -Path $TestDrive -ChildPath "full-filters"
        $objectPath = Join-Path -Path $importPath -ChildPath "APO120 @ 0.8x\Light\Filter test"
        New-Item -ItemType Directory -Path $objectPath -Force | Out-Null
        $filterFiles = @(
            @{ Camera = "2600MC"; Filter = ""; Index = "01" },
            @{ Camera = "2600MM"; Filter = "L"; Index = "02" },
            @{ Camera = "2600MM"; Filter = "R"; Index = "03" },
            @{ Camera = "2600MM"; Filter = "G"; Index = "04" },
            @{ Camera = "2600MM"; Filter = "B"; Index = "05" },
            @{ Camera = "2600MC"; Filter = "HO"; Index = "06" },
            @{ Camera = "2600MC"; Filter = "SO"; Index = "07" },
            @{ Camera = "2600MM"; Filter = "H"; Index = "08" },
            @{ Camera = "2600MM"; Filter = "O"; Index = "09" },
            @{ Camera = "2600MM"; Filter = "S"; Index = "10" }
        )
        foreach ($item in $filterFiles) {
            $filterToken = if ([string]::IsNullOrWhiteSpace($item.Filter)) { "" } else { "_$($item.Filter)" }
            $fileName = "Light_Filter test_60.0s_Bin1_$($item.Camera)${filterToken}_gain120_20260710-05$($item.Index)12_0deg_-10.0C_APO120_00$($item.Index).fit"
            New-Item -ItemType File -Path (Join-Path -Path $objectPath -ChildPath $fileName) | Out-Null
        }

        $report = @(Get-AsiToPixImportReport -ImportPath $importPath)

        $report.Count | Should Be 1
        $report[0].FrameCount | Should Be 10
        foreach ($propertyName in @("TsvRGB", "TsvL", "TsvR", "TsvG", "TsvB", "TsvHO", "TsvSO", "TsvHa", "TsvOIII", "TsvSII")) {
            $report[0].$propertyName | Should Be "1*60"
        }
        $report[0].RGB | Should Be "1*60"
        $report[0].L | Should Be "1*60"
        $report[0].H | Should Be "1*60"
        $report[0].O | Should Be "1*60"
        $report[0].S | Should Be "1*60"
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

    It "scans a plural Lights folder and shared non-FITS formats" {
        $importPath = Join-Path -Path $TestDrive -ChildPath "plural-lights"
        $objectPath = Join-Path -Path $importPath -ChildPath "SQA55 @ 1.0x\lIgHtS\M 31"
        New-Item -ItemType Directory -Path $objectPath -Force | Out-Null
        @(
            "Light_M31_120.0s_Bin1_2600MC_gain120_20260717-193952_0deg_-10.0C_SQA55_0001.xisf",
            "Light_M31_120.0s_Bin1_2600MC_gain120_20260718-193952_0deg_-10.0C_SQA55_0002.tiff"
        ) | ForEach-Object {
            New-Item -ItemType File -Path (Join-Path -Path $objectPath -ChildPath $_) | Out-Null
        }

        $report = @(Get-AsiToPixImportReport -ImportPath $importPath)

        $report.Count | Should Be 1
        $report[0].Object | Should Be "M 31"
        $report[0].FrameCount | Should Be 2
        $report[0].RGB | Should Be "(1+1)*120"
    }

    It "prompts once for missing RAW exposure and uses the file timestamp" {
        $importPath = Join-Path -Path $TestDrive -ChildPath "raw-report"
        $objectPath = Join-Path -Path $importPath -ChildPath "Canon 200mm\Lights\Rho Oph"
        New-Item -ItemType Directory -Path $objectPath -Force | Out-Null
        $first = New-Item -ItemType File -Path (Join-Path -Path $objectPath -ChildPath "IMG_0001.CR3")
        $second = New-Item -ItemType File -Path (Join-Path -Path $objectPath -ChildPath "IMG_0002.CR3")
        $first.LastWriteTime = [datetime]"2026-07-18T01:00:00"
        $second.LastWriteTime = [datetime]"2026-07-18T02:00:00"

        Mock Read-AsiToPixRequiredValue { "30" } -ModuleName AsiToPix.ImportSession

        $report = @(Get-AsiToPixImportReport -ImportPath $importPath -PromptForMissingData)

        $report.Count | Should Be 1
        $report[0].FrameCount | Should Be 2
        $report[0].RGB | Should Be "2*30"
        Assert-MockCalled Read-AsiToPixRequiredValue -Times 1 -ModuleName AsiToPix.ImportSession
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
