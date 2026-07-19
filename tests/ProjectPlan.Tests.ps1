$planModulePath = Join-Path -Path $PSScriptRoot -ChildPath "..\src\AsiToPix.ProjectPlan.psm1"
$metadataModulePath = Join-Path -Path $PSScriptRoot -ChildPath "..\src\AsiToPix.ProjectMetadata.psm1"
Import-Module $planModulePath -Force
Import-Module $metadataModulePath -Force

function New-TestFlatPendingLink {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,

        [Parameter(Mandatory = $true)]
        [string]$FlatSetId,

        [Parameter(Mandatory = $true)]
        [string]$Tag,

        [Parameter(Mandatory = $true)]
        [string]$LightSession
    )

    return [PSCustomObject]@{
        Type                 = "Flats"
        Tag                  = $Tag
        Src                  = $SourcePath
        Display              = "Source\flats\SQA55 @ 1.0x\26.07.15 H 180deg"
        Cam                  = "ASI2600MM"
        Gain                 = "120"
        Temperature          = "-10C"
        Exposure             = "0.81s"
        Filter               = "H"
        Session              = "26.07.15"
        LightSessions        = @($LightSession)
        Target               = "H"
        FlatSetId            = $FlatSetId
        Binning              = "1"
        Setup                = "SQA55 @ 1.0x"
        Angle                = "180deg"
        FlatCompatibilityKey = "ASI2600MM|1|H|H|SQA55 @ 1.0x|120|-10C"
    }
}

Describe "CreateProject flat-set planning" {
    It "deduplicates the SMC 180s and 300s sessions by their physical flat source" {
        $cameraRoot = Join-Path -Path $TestDrive -ChildPath "Calibration\ASI2600MM"
        $flatSource = Join-Path -Path $cameraRoot -ChildPath "Source\flats\SQA55 @ 1.0x\26.07.15 H 180deg"
        $masterRoot = Join-Path -Path $cameraRoot -ChildPath "Master\flats"
        [void](New-Item -ItemType Directory -Path $flatSource -Force)
        foreach ($number in 1..20) {
            $fileName = "Flat_0.810s_Bin1_2600MM_H_gain120_20260715-120000_180deg_-10.0C_SQA55_{0:D4}.fit" -f $number
            [void](New-Item -ItemType File -Path (Join-Path -Path $flatSource -ChildPath $fileName))
        }

        $flatSetId = Get-AsiToPixFlatSetId -SourcePath $flatSource
        $flatTag = ConvertTo-AsiToPixFlatSetTag `
            -FlatSetId $flatSetId `
            -FlatDate "26.07.15" `
            -Angle "180deg" `
            -Setup "SQA55 @ 1.0x" `
            -Binning "1" `
            -Filter "H" `
            -Target "H" `
            -Gain "120" `
            -Temperature "-10C" `
            -Camera "ASI2600MM"
        $pendingLinks = @(
            New-TestFlatPendingLink -SourcePath $flatSource -FlatSetId $flatSetId -Tag $flatTag -LightSession "26.07.15-180s"
            New-TestFlatPendingLink -SourcePath $flatSource -FlatSetId $flatSetId -Tag $flatTag -LightSession "26.07.16"
            New-TestFlatPendingLink -SourcePath $flatSource -FlatSetId $flatSetId -Tag $flatTag -LightSession "26.07.15-300s"
        )

        $plan = Get-AsiToPixUniqueFlatPlan -PendingLink $pendingLinks
        $plannedFlats = @($plan.PendingLinks | Where-Object { $_.Type -eq "Flats" })

        $plannedFlats.Count | Should Be 1
        $plannedFlats[0].Tag | Should Match "FlatSet_$flatSetId"
        $plannedFlats[0].Tag | Should Match "FlatDate_26.07.15_Angle_180deg"
        $plannedFlats[0].Tag | Should Match "Temp_-10C"
        $plannedFlats[0].Tag | Should Not Match '(?:^|_)Exp_'
        $plannedFlats[0].Tag | Should Not Match 'Session_26\.07\.'
        $plan.DuplicateGroups.Count | Should Be 1
        $plan.DuplicateGroups[0].Count | Should Be 3
        $plan.DuplicateGroups[0].LightSessions.Count | Should Be 3

        $projectFlatFiles = @(
            $plannedFlats | ForEach-Object {
                Get-ChildItem -LiteralPath $_.Src -File -Filter "*.fit" | Select-Object -ExpandProperty FullName
            }
        )
        $projectFlatFiles.Count | Should Be 20
        @($projectFlatFiles | Sort-Object -Unique).Count | Should Be 20

        $camera = [PSCustomObject]@{
            Name = "ASI2600MM"
            CalibrationFolders = [PSCustomObject]@{ Flats = $masterRoot }
        }
        $metadata = @(
            ConvertTo-AsiToPixCalibrationSourceMetadata `
                -PendingLink $plan.PendingLinks `
                -CameraMetadata @($camera)
        )
        $metadata.Count | Should Be 1
        $metadata[0].FlatSetId | Should Be $flatSetId
        $metadata[0].ExposureSeconds | Should Be "0.81"
        $metadata[0].LightSessions.Count | Should Be 3
    }

    It "compares canonical flat paths without regard to Windows path casing" {
        $flatSource = Join-Path -Path $TestDrive -ChildPath "CaseTest\FlatSet"
        [void](New-Item -ItemType Directory -Path $flatSource -Force)
        $flatSetId = Get-AsiToPixFlatSetId -SourcePath $flatSource
        $flatTag = ConvertTo-AsiToPixFlatSetTag `
            -FlatSetId $flatSetId -Setup "SQA55 @ 1.0x" -Filter "H" -Target "H" `
            -Gain "120" -Temperature "-10C" -Camera "ASI2600MM"
        $pendingLinks = @(
            New-TestFlatPendingLink -SourcePath $flatSource -FlatSetId $flatSetId -Tag $flatTag -LightSession "26.07.15"
            New-TestFlatPendingLink -SourcePath $flatSource.ToUpperInvariant() -FlatSetId $flatSetId -Tag $flatTag -LightSession "26.07.16"
        )

        $plan = Get-AsiToPixUniqueFlatPlan -PendingLink $pendingLinks

        @($plan.PendingLinks | Where-Object { $_.Type -eq "Flats" }).Count | Should Be 1
        $plan.DuplicateGroups[0].Count | Should Be 2
    }

    It "keeps distinct physical flat sets separate and reports their compatibility collision" {
        $firstSource = Join-Path -Path $TestDrive -ChildPath "Separate\FlatSetA"
        $secondSource = Join-Path -Path $TestDrive -ChildPath "Separate\FlatSetB"
        [void](New-Item -ItemType Directory -Path $firstSource -Force)
        [void](New-Item -ItemType Directory -Path $secondSource -Force)
        $firstId = Get-AsiToPixFlatSetId -SourcePath $firstSource
        $secondId = Get-AsiToPixFlatSetId -SourcePath $secondSource
        $firstTag = ConvertTo-AsiToPixFlatSetTag `
            -FlatSetId $firstId -Setup "SQA55 @ 1.0x" -Filter "H" -Target "H" `
            -Gain "120" -Temperature "-10C" -Camera "ASI2600MM"
        $secondTag = ConvertTo-AsiToPixFlatSetTag `
            -FlatSetId $secondId -Setup "SQA55 @ 1.0x" -Filter "H" -Target "H" `
            -Gain "120" -Temperature "-10C" -Camera "ASI2600MM"
        $pendingLinks = @(
            New-TestFlatPendingLink -SourcePath $firstSource -FlatSetId $firstId -Tag $firstTag -LightSession "26.07.15"
            New-TestFlatPendingLink -SourcePath $secondSource -FlatSetId $secondId -Tag $secondTag -LightSession "26.07.16"
        )

        $plan = Get-AsiToPixUniqueFlatPlan -PendingLink $pendingLinks

        @($plan.PendingLinks | Where-Object { $_.Type -eq "Flats" }).Count | Should Be 2
        $plan.SeparateSetGroups.Count | Should Be 1
        $plan.SeparateSetGroups[0].FlatSets.Count | Should Be 2
    }
}
