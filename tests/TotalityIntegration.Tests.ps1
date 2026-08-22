$modulePath = Join-Path -Path $PSScriptRoot -ChildPath "..\src\AsiToPix.TotalityIntegration.psm1"
Import-Module $modulePath -Force

function Assert-TestTotalityError {
    param(
        [Parameter(Mandatory)]
        [scriptblock]$Script,

        [Parameter(Mandatory)]
        [string]$MessagePattern
    )

    $caughtError = $null
    try {
        $null = & $Script
    }
    catch {
        $caughtError = $_
    }
    $caughtError | Should Not BeNullOrEmpty
    $caughtError.Exception.Message | Should Match $MessagePattern
}

function Get-TestTotalityFrame {
    param(
        [int]$Index
    )

    return [PSCustomObject]@{
        Index     = $Index
        IndexText = "{0:D5}" -f $Index
        Name      = "Totality_1ms_{0:D5}_c_d_r.xisf" -f $Index
        Path      = Join-Path -Path $TestDrive -ChildPath ("frame_{0:D5}.xisf" -f $Index)
        Extension = "xisf"
    }
}

Describe "Totality frame discovery" {
    It "finds registered frames and sorts their numeric indexes" {
        $registeredPath = Join-Path -Path $TestDrive -ChildPath "registered-éclipse"
        New-Item -ItemType Directory -Path $registeredPath -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $registeredPath "Totality_1ms_00082_c_cc_d_r.xisf") -Value "82"
        Set-Content -LiteralPath (Join-Path $registeredPath "Totality_1ms_00080_arbitrary_changed_suffix.XISF") -Value "80"
        Set-Content -LiteralPath (Join-Path $registeredPath "Totality_1ms_00079.xisf") -Value "79"
        Set-Content -LiteralPath (Join-Path $registeredPath "Totality_1ms_00081.fit") -Value "transition"
        Set-Content -LiteralPath (Join-Path $registeredPath "notes.txt") -Value "ignore"

        $frames = @(Get-AsiToPixTotalityFrame -RegisteredPath $registeredPath)

        $frames.Count | Should Be 4
        $frames[0].Index | Should Be 79
        $frames[1].Index | Should Be 80
        $frames[3].Index | Should Be 82
        $frames[1].Name | Should Be "Totality_1ms_00080_arbitrary_changed_suffix.XISF"
    }

    It "rejects duplicate numeric indexes" {
        $registeredPath = Join-Path -Path $TestDrive -ChildPath "duplicate-registered"
        New-Item -ItemType Directory -Path $registeredPath -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $registeredPath "Totality_1ms_00080_c_d_r.xisf") -Value "xisf"
        Set-Content -LiteralPath (Join-Path $registeredPath "Alternate_1ms_00080_c_d_r.fit") -Value "fit"

        Assert-TestTotalityError `
            -Script { Get-AsiToPixTotalityFrame -RegisteredPath $registeredPath } `
            -MessagePattern "Duplicate input frame indexes"
    }
}

function Get-TestTotalityRawFrame {
    param(
        [int]$Index
    )

    return [PSCustomObject]@{
        Index     = $Index
        IndexText = "{0:D5}" -f $Index
        Name      = "Totality_100ms_{0:D5}.fit" -f $Index
        Path      = Join-Path -Path $TestDrive -ChildPath ("raw_{0:D5}.fit" -f $Index)
        Extension = "fit"
    }
}

Describe "Totality sequence origin detection" {
    It "detects a one-based sequence from a Transition frame" {
        $exposurePath = Join-Path -Path $TestDrive -ChildPath "origin-one\Total-10ms"
        $transitionPath = Join-Path -Path $exposurePath -ChildPath "Transition"
        $registeredPath = Join-Path -Path $exposurePath -ChildPath "Totality\registered"
        $integratedPath = Join-Path -Path $exposurePath -ChildPath "Totality\integrated"
        New-Item -ItemType Directory -Path $transitionPath, $registeredPath, $integratedPath -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $transitionPath "Totality_10ms_00001.fit") -Value "transition"
        Set-Content -LiteralPath (Join-Path $registeredPath "Totality_10ms_00027_c_d_r.xisf") -Value "frame"
        Set-Content -LiteralPath (Join-Path $integratedPath "Block_001_00000_00025.xisf") -Value "ignore"

        Get-AsiToPixTotalitySequenceStartIndex -ExposureDirectory $exposurePath | Should Be 1
    }

    It "prefers zero when both zero and one indexes exist" {
        $exposurePath = Join-Path -Path $TestDrive -ChildPath "origin-zero\Total-1ms"
        $transitionPath = Join-Path -Path $exposurePath -ChildPath "Transition"
        New-Item -ItemType Directory -Path $transitionPath -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $transitionPath "Totality_1ms_00001.fit") -Value "one"
        Set-Content -LiteralPath (Join-Path $transitionPath "Totality_1ms_00000.fit") -Value "zero"

        Get-AsiToPixTotalitySequenceStartIndex -ExposureDirectory $exposurePath | Should Be 0
    }

    It "requires an explicit override when neither origin index is available" {
        $exposurePath = Join-Path -Path $TestDrive -ChildPath "origin-unknown\Total-1ms\Totality\registered"
        New-Item -ItemType Directory -Path $exposurePath -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $exposurePath "Totality_1ms_00023_c_d_r.xisf") -Value "frame"

        Assert-TestTotalityError `
            -Script {
                Get-AsiToPixTotalitySequenceStartIndex `
                    -ExposureDirectory (Split-Path -Path (Split-Path -Path $exposurePath -Parent) -Parent)
            } `
            -MessagePattern "Cannot determine whether the sequence"
    }
}

Describe "Totality frame manifest" {
    It "creates a missing human-editable manifest from interactive answers" {
        $framesPath = Join-Path -Path $TestDrive -ChildPath "create-manifest\Frames"
        New-Item -ItemType Directory -Path $framesPath -Force | Out-Null
        $manifestPath = Join-Path -Path $framesPath -ChildPath "manifest.json"
        $answers = [System.Collections.Generic.Queue[string]]::new()
        @("25", "1", "22", "83", "1, 2, 26-27") | ForEach-Object { $answers.Enqueue($_) }
        $reader = { param($Prompt) $null = $Prompt; $answers.Dequeue() }.GetNewClosure()

        $result = Resolve-AsiToPixTotalityFrameManifest `
            -ManifestPath $manifestPath `
            -MinimumIndex 1 `
            -MaximumIndex 100 `
            -InputReader $reader

        $result.WasWritten | Should Be $true
        Test-Path -LiteralPath $manifestPath -PathType Leaf | Should Be $true
        $saved = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $saved.BlockLength | Should Be 25
        $saved.BlockStartIndex | Should Be 1
        $saved.C2Index | Should Be 22
        $saved.C3Index | Should Be 83
        @($saved.SkipIndices).Count | Should Be 4
        $saved.SkipIndices[3] | Should Be 27
        [System.IO.File]::ReadAllBytes($manifestPath)[0] | Should Not Be 239
    }

    It "fills only missing fields and preserves additional user properties" {
        $framesPath = Join-Path -Path $TestDrive -ChildPath "complete-manifest\Frames"
        New-Item -ItemType Directory -Path $framesPath -Force | Out-Null
        $manifestPath = Join-Path -Path $framesPath -ChildPath "manifest.json"
        [ordered]@{
            BlockLength      = 30
            BlockStartIndex  = 1
            C2Index          = 22
            RegistrationIndex = 160
        } | ConvertTo-Json | Set-Content -LiteralPath $manifestPath -Encoding UTF8
        $answers = [System.Collections.Generic.Queue[string]]::new()
        @("183", "1,2,31") | ForEach-Object { $answers.Enqueue($_) }
        $reader = { param($Prompt) $null = $Prompt; $answers.Dequeue() }.GetNewClosure()

        $result = Resolve-AsiToPixTotalityFrameManifest `
            -ManifestPath $manifestPath `
            -MinimumIndex 1 `
            -MaximumIndex 300 `
            -InputReader $reader

        $result.Manifest.RegistrationIndex | Should Be 160
        $result.Manifest.C3Index | Should Be 183
        @($result.Manifest.SkipIndices).Count | Should Be 3
    }

    It "accepts an explicit null C2 boundary without prompting" {
        $framesPath = Join-Path -Path $TestDrive -ChildPath "null-c2-manifest\Frames"
        New-Item -ItemType Directory -Path $framesPath -Force | Out-Null
        $manifestPath = Join-Path -Path $framesPath -ChildPath "manifest.json"
        [ordered]@{
            BlockLength     = 30
            BlockStartIndex = 1
            C2Index         = $null
            C3Index         = 287
            SkipIndices     = @(1, 2, 31)
        } | ConvertTo-Json | Set-Content -LiteralPath $manifestPath -Encoding UTF8
        $reader = { param($Prompt) throw "Unexpected prompt: $Prompt" }

        $result = Resolve-AsiToPixTotalityFrameManifest `
            -ManifestPath $manifestPath `
            -MinimumIndex 1 `
            -MaximumIndex 300 `
            -InputReader $reader

        $result.WasWritten | Should Be $false
        $result.Manifest.C2Index | Should Be $null
        $result.Manifest.C3Index | Should Be 287
    }

    It "explains how to replace a suspicious C2 index with null" {
        $framesPath = Join-Path -Path $TestDrive -ChildPath "invalid-c2-manifest\Frames"
        New-Item -ItemType Directory -Path $framesPath -Force | Out-Null
        $manifestPath = Join-Path -Path $framesPath -ChildPath "manifest.json"
        [ordered]@{
            BlockLength     = 30
            BlockStartIndex = 1
            C2Index         = 0
            C3Index         = 287
            SkipIndices     = @()
        } | ConvertTo-Json | Set-Content -LiteralPath $manifestPath -Encoding UTF8
        $message = ""

        try {
            Resolve-AsiToPixTotalityFrameManifest `
                -ManifestPath $manifestPath `
                -MinimumIndex 1 `
                -MaximumIndex 300 | Out-Null
        }
        catch {
            $message = $_.Exception.Message
        }

        $message | Should Match "Suspicious index for C2Index: 0"
        $message | Should Match 'Please type "C2Index": null if your dataset has no C2 contact frame'
        $message | Should Match "raw frame range 1-300"
    }
}

Describe "Manifest-based totality planning" {
    It "marks the C2 containing and previous blocks and the C3 containing and next blocks as beads" {
        $rawFrames = @(1..60 | ForEach-Object { Get-TestTotalityRawFrame -Index $_ })
        $registeredFrames = @(1..60 | ForEach-Object { Get-TestTotalityFrame -Index $_ })
        $manifest = [PSCustomObject]@{
            BlockLength     = 10
            BlockStartIndex = 1
            C2Index         = 15
            C3Index         = 45
            SkipIndices     = @()
        }

        $plan = @(Get-AsiToPixTotalityFrameManifestPlan `
            -RawFrames $rawFrames `
            -RegisteredFrames $registeredFrames `
            -Manifest $manifest `
            -OutputDirectory (Join-Path $TestDrive "manifest-plan") `
            -ExposureLabel "100ms")

        $plan.Count | Should Be 6
        @($plan.Phase) -join "," | Should Be "BeadsC2,BeadsC2,Totality,Totality,BeadsC3,BeadsC3"
        $plan[0].OutputPath | Should Be (Join-Path $TestDrive "manifest-plan\Block_001_100ms_BeadsC2_00001_00010.xisf")
        $plan[2].OutputPath | Should Be (Join-Path $TestDrive "manifest-plan\Block_003_100ms_Totality_00021_00030.xisf")
        $plan[5].OutputPath | Should Be (Join-Path $TestDrive "manifest-plan\Block_006_100ms_BeadsC3_00051_00060.xisf")
    }

    It "starts with Totality when the C2 boundary is null" {
        $rawFrames = @(1..60 | ForEach-Object { Get-TestTotalityRawFrame -Index $_ })
        $registeredFrames = @(1..60 | ForEach-Object { Get-TestTotalityFrame -Index $_ })
        $manifest = [PSCustomObject]@{
            BlockLength     = 10
            BlockStartIndex = 1
            C2Index         = $null
            C3Index         = 45
            SkipIndices     = @()
        }

        $plan = @(Get-AsiToPixTotalityFrameManifestPlan `
            -RawFrames $rawFrames `
            -RegisteredFrames $registeredFrames `
            -Manifest $manifest `
            -OutputDirectory (Join-Path $TestDrive "manifest-no-c2") `
            -ExposureLabel "100ms")

        $plan.Count | Should Be 6
        @($plan.Phase) -join "," | Should Be "Totality,Totality,Totality,Totality,BeadsC3,BeadsC3"
        $plan[0].OutputPath | Should Be (Join-Path $TestDrive "manifest-no-c2\Block_001_100ms_Totality_00001_00010.xisf")
    }

    It "ends with Totality when the C3 boundary is null" {
        $rawFrames = @(1..60 | ForEach-Object { Get-TestTotalityRawFrame -Index $_ })
        $registeredFrames = @(1..60 | ForEach-Object { Get-TestTotalityFrame -Index $_ })
        $manifest = [PSCustomObject]@{
            BlockLength     = 10
            BlockStartIndex = 1
            C2Index         = 15
            C3Index         = $null
            SkipIndices     = @()
        }

        $plan = @(Get-AsiToPixTotalityFrameManifestPlan `
            -RawFrames $rawFrames `
            -RegisteredFrames $registeredFrames `
            -Manifest $manifest `
            -OutputDirectory (Join-Path $TestDrive "manifest-no-c3") `
            -ExposureLabel "100ms")

        $plan.Count | Should Be 6
        @($plan.Phase) -join "," | Should Be "BeadsC2,BeadsC2,Totality,Totality,Totality,Totality"
        $plan[5].OutputPath | Should Be (Join-Path $TestDrive "manifest-no-c3\Block_006_100ms_Totality_00051_00060.xisf")
    }

    It "uses declared transition indexes without treating them as missing frames" {
        $rawFrames = @(1..60 | ForEach-Object { Get-TestTotalityRawFrame -Index $_ })
        $skipIndexes = @(1, 2, 21)
        $registeredFrames = @(1..60 | Where-Object { $_ -notin $skipIndexes } | ForEach-Object {
            Get-TestTotalityFrame -Index $_
        })
        $manifest = [PSCustomObject]@{
            BlockLength     = 10
            BlockStartIndex = 1
            C2Index         = 15
            C3Index         = 45
            SkipIndices     = $skipIndexes
        }

        $plan = @(Get-AsiToPixTotalityFrameManifestPlan `
            -RawFrames $rawFrames `
            -RegisteredFrames $registeredFrames `
            -Manifest $manifest `
            -OutputDirectory (Join-Path $TestDrive "manifest-skips") `
            -ExposureLabel "100ms")

        $plan[0].CanIntegrate | Should Be $true
        Format-AsiToPixTotalityIndexList -Indexes $plan[0].IntegrateIndexes | Should Be "3-10"
        Format-AsiToPixTotalityIndexList -Indexes $plan[0].SkippedIndexes | Should Be "1-2"
        $plan[0].OutputPath | Should Be (Join-Path $TestDrive "manifest-skips\Block_001_100ms_BeadsC2_00003_00010.xisf")
        $plan[2].OutputPath | Should Be (Join-Path $TestDrive "manifest-skips\Block_003_100ms_Totality_00022_00030.xisf")
    }

    It "reports indexes whose registration is not finished" {
        $rawFrames = @(1..60 | ForEach-Object { Get-TestTotalityRawFrame -Index $_ })
        $registeredFrames = @(1..24 | ForEach-Object { Get-TestTotalityFrame -Index $_ })
        $manifest = [PSCustomObject]@{
            BlockLength     = 10
            BlockStartIndex = 1
            C2Index         = 15
            C3Index         = 45
            SkipIndices     = @()
        }

        $plan = @(Get-AsiToPixTotalityFrameManifestPlan `
            -RawFrames $rawFrames `
            -RegisteredFrames $registeredFrames `
            -Manifest $manifest `
            -OutputDirectory (Join-Path $TestDrive "manifest-incomplete") `
            -ExposureLabel "100ms")

        $plan[2].CanIntegrate | Should Be $false
        Format-AsiToPixTotalityIndexList -Indexes $plan[2].MissingRegisteredIndexes | Should Be "25-30"
        $plan[2].SkipReason | Should Match "registration incomplete: 25-30"
    }

    It "reports incomplete debayering when Debayered is the input stage" {
        $rawFrames = @(1..30 | ForEach-Object { Get-TestTotalityRawFrame -Index $_ })
        $inputFrames = @(1..24 | ForEach-Object { Get-TestTotalityFrame -Index $_ })
        $manifest = [PSCustomObject]@{
            BlockLength     = 10
            BlockStartIndex = 1
            C2Index         = $null
            C3Index         = $null
            SkipIndices     = @()
        }

        $plan = @(Get-AsiToPixTotalityFrameManifestPlan `
            -RawFrames $rawFrames `
            -InputFrames $inputFrames `
            -Manifest $manifest `
            -OutputDirectory (Join-Path $TestDrive "manifest-debayered-incomplete") `
            -ExposureLabel "100ms" `
            -InputStage "Debayered")

        $plan[2].InputStage | Should Be "Debayered"
        Format-AsiToPixTotalityIndexList -Indexes $plan[2].MissingInputIndexes | Should Be "25-30"
        $plan[2].SkipReason | Should Match "debayering incomplete: 25-30"
    }
}

Describe "Totality input stage selection" {
    It "uses Registered by default" {
        $reader = { param($Prompt) $null = $Prompt; "" }

        Read-AsiToPixTotalityInputStage -InputReader $reader | Should Be "Registered"
    }

    It "accepts Debayered by number or name" {
        $numberReader = { param($Prompt) $null = $Prompt; "2" }
        $nameReader = { param($Prompt) $null = $Prompt; "debayered" }

        Read-AsiToPixTotalityInputStage -InputReader $numberReader | Should Be "Debayered"
        Read-AsiToPixTotalityInputStage -InputReader $nameReader | Should Be "Debayered"
    }
}

Describe "PixInsight execution mode selection" {
    It "uses a running instance by default" {
        $reader = { param($Prompt) $null = $Prompt; "" }

        Read-AsiToPixPixInsightMode -RunningInstanceCount 1 -InputReader $reader | Should Be "Reuse"
    }

    It "accepts a dedicated instance by number or name" {
        $numberReader = { param($Prompt) $null = $Prompt; "2" }
        $nameReader = { param($Prompt) $null = $Prompt; "dedicated" }

        Read-AsiToPixPixInsightMode -RunningInstanceCount 2 -InputReader $numberReader | Should Be "Dedicated"
        Read-AsiToPixPixInsightMode -RunningInstanceCount 1 -InputReader $nameReader | Should Be "Dedicated"
    }

    It "uses a dedicated instance when none is running" {
        $reader = { param($Prompt) throw "Unexpected prompt: $Prompt" }

        Read-AsiToPixPixInsightMode -RunningInstanceCount 0 -InputReader $reader | Should Be "Dedicated"
    }
}

Describe "Totality block length input" {
    It "selects automatic gap splitting when Enter is pressed" {
        $reader = { param($Prompt) $null = $Prompt; "" }

        Read-AsiToPixTotalityBlockLength -InputReader $reader | Should Be 0
    }

    It "accepts a nominal block size" {
        $reader = { param($Prompt) $null = $Prompt; "25" }

        Read-AsiToPixTotalityBlockLength -InputReader $reader | Should Be 25
    }
}

Describe "Totality integration planning" {
    It "splits a one-based sequence into nominal index ranges" {
        $frames = @(61..105 | ForEach-Object { Get-TestTotalityFrame -Index $_ })
        $outputDirectory = Join-Path -Path $TestDrive -ChildPath "integrated"

        $plan = @(Get-AsiToPixTotalityIntegrationPlan `
            -Frames $frames `
            -OutputDirectory $outputDirectory `
            -BlockLength 20 `
            -SequenceStartIndex 1)

        $plan.Count | Should Be 3
        $plan[0].FrameCount | Should Be 20
        $plan[1].FrameCount | Should Be 20
        $plan[2].FrameCount | Should Be 5
        $plan[0].OutputPath | Should Be (Join-Path $outputDirectory "Block_001_00061_00080.xisf")
        $plan[1].OutputPath | Should Be (Join-Path $outputDirectory "Block_002_00081_00100.xisf")
        $plan[2].OutputPath | Should Be (Join-Path $outputDirectory "Block_003_00101_00105.xisf")
        $plan[2].CanIntegrate | Should Be $true
        $plan[0].NominalFirstIndex | Should Be 61
        $plan[0].NominalLastIndex | Should Be 80
        $plan[0].SplitMode | Should Be "NominalIndexRanges"
    }

    It "separates adjacent runs at nominal boundaries even without an index gap" {
        $frames = @(27..74 | ForEach-Object { Get-TestTotalityFrame -Index $_ })
        $outputDirectory = Join-Path -Path $TestDrive -ChildPath "nominal-adjacent"

        $plan = @(Get-AsiToPixTotalityIntegrationPlan `
            -Frames $frames `
            -OutputDirectory $outputDirectory `
            -BlockLength 25 `
            -SequenceStartIndex 1)

        $plan.Count | Should Be 2
        $plan[0].NominalFirstIndex | Should Be 26
        $plan[0].NominalLastIndex | Should Be 50
        $plan[0].FirstIndex | Should Be 27
        $plan[0].LastIndex | Should Be 50
        $plan[0].FrameCount | Should Be 24
        $plan[1].NominalFirstIndex | Should Be 51
        $plan[1].NominalLastIndex | Should Be 75
        $plan[1].FirstIndex | Should Be 51
        $plan[1].LastIndex | Should Be 74
        $plan[1].FrameCount | Should Be 24
    }

    It "rejects a nominal block with a missing index inside its registered run" {
        $frames = @(26..35 | Where-Object { $_ -ne 30 } | ForEach-Object {
            Get-TestTotalityFrame -Index $_
        })
        $outputDirectory = Join-Path -Path $TestDrive -ChildPath "non-contiguous"

        $plan = @(Get-AsiToPixTotalityIntegrationPlan `
            -Frames $frames `
            -OutputDirectory $outputDirectory `
            -BlockLength 25 `
            -SequenceStartIndex 1)

        $plan.Count | Should Be 1
        $plan[0].CanIntegrate | Should Be $false
        $plan[0].IsContiguous | Should Be $false
        @($plan[0].MissingIndexes).Count | Should Be 1
        $plan[0].MissingIndexes[0] | Should Be 30
        $plan[0].SkipReason | Should Match "non-contiguous frame indexes"
    }

    It "splits automatically whenever a numeric index is missing" {
        $frames = @(61, 62, 63, 65, 66, 67, 70, 71 | ForEach-Object {
            Get-TestTotalityFrame -Index $_
        })
        $outputDirectory = Join-Path -Path $TestDrive -ChildPath "gap-integrated"

        $plan = @(Get-AsiToPixTotalityIntegrationPlan `
            -Frames $frames `
            -OutputDirectory $outputDirectory `
            -BlockLength 0)

        $plan.Count | Should Be 3
        $plan[0].FirstIndex | Should Be 61
        $plan[0].LastIndex | Should Be 63
        $plan[1].FirstIndex | Should Be 65
        $plan[1].LastIndex | Should Be 67
        $plan[2].FirstIndex | Should Be 70
        $plan[2].LastIndex | Should Be 71
        $plan[2].CanIntegrate | Should Be $false
        $plan[2].SkipReason | Should Be "fewer than 3 frames"
        $plan[2].SplitMode | Should Be "IndexGaps"
    }

    It "rejects fixed blocks too small for ImageIntegration" {
        $frames = @(1..4 | ForEach-Object { Get-TestTotalityFrame -Index $_ })

        Assert-TestTotalityError `
            -Script {
                Get-AsiToPixTotalityIntegrationPlan `
                    -Frames $frames `
                    -OutputDirectory $TestDrive `
                    -BlockLength 2
            } `
            -MessagePattern "Block length must be zero"
    }
}

Describe "PixInsight integration handoff" {
    It "writes one UTF-8 JSON manifest containing all integration blocks" {
        $manifestDirectory = Join-Path -Path $TestDrive -ChildPath "manifest"
        $outputDirectory = Join-Path -Path $TestDrive -ChildPath "manifest-output"
        New-Item -ItemType Directory -Path $manifestDirectory, $outputDirectory -Force | Out-Null
        $frames = @(21..26 | ForEach-Object {
            $frame = Get-TestTotalityFrame -Index $_
            Set-Content -LiteralPath $frame.Path -Value "frame"
            $frame
        })
        $blocks = @(Get-AsiToPixTotalityIntegrationPlan `
            -Frames $frames `
            -OutputDirectory $outputDirectory `
            -BlockLength 3)

        $manifestInfo = New-AsiToPixTotalityManifest `
            -Blocks $blocks `
            -ManifestDirectory $manifestDirectory
        $manifest = Get-Content -LiteralPath $manifestInfo.ManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json

        $manifest.schemaVersion | Should Be 2
        $manifest.statusPath | Should Be $manifestInfo.StatusPath
        @($manifest.blocks).Count | Should Be 2
        $manifest.blocks[0].blockNumber | Should Be 1
        $manifest.blocks[0].outputPath | Should Be $blocks[0].OutputPath
        @($manifest.blocks[0].inputFiles).Count | Should Be 3
        $manifest.blocks[0].inputFiles[0] | Should Be $frames[0].Path
        $manifest.blocks[1].outputPath | Should Be $blocks[1].OutputPath
        [System.IO.File]::ReadAllBytes($manifestInfo.ManifestPath)[0] | Should Not Be 239
    }

    It "uses yes as the default plan confirmation" {
        InModuleScope AsiToPix.TotalityIntegration {
            Mock Read-Host { "" }

            Read-AsiToPixTotalityPlanConfirmation | Should Be $true

            Assert-MockCalled Read-Host -Times 1
        }
    }

    It "accepts English and Russian yes answers" {
        InModuleScope AsiToPix.TotalityIntegration {
            $answers = @("y", "Y", [string][char]0x0434, [string][char]0x0414)
            foreach ($answer in $answers) {
                $reader = { param($Prompt) $null = $Prompt; $answer }.GetNewClosure()
                Read-AsiToPixTotalityPlanConfirmation -InputReader $reader | Should Be $true
            }
        }
    }

    It "accepts English and Russian no answers" {
        InModuleScope AsiToPix.TotalityIntegration {
            $answers = @("n", "N", [string][char]0x043D, [string][char]0x041D)
            foreach ($answer in $answers) {
                $reader = { param($Prompt) $null = $Prompt; $answer }.GetNewClosure()
                Read-AsiToPixTotalityPlanConfirmation -InputReader $reader | Should Be $false
            }
        }
    }

    It "builds the documented PixInsight automation arguments" {
        $scriptPath = "C:\Tools\AsiToPix\pixinsight\IntegrateTotality.js"
        $manifestPath = "C:\Temp\AsiToPix-Totality\Block_001.json"

        $arguments = @(Get-AsiToPixPixInsightArgumentList `
            -ScriptPath $scriptPath `
            -ManifestPath $manifestPath)

        $arguments.Count | Should Be 4
        $arguments[0] | Should Be "-n"
        $arguments[1] | Should Be "--automation-mode"
        $arguments[2] | Should Be ('-r="{0},manifest={1}"' -f $scriptPath, $manifestPath)
        $arguments[3] | Should Be "--force-exit"
    }

    It "builds IPC arguments that neither create nor close a PixInsight instance" {
        $scriptPath = "C:\Tools\AsiToPix\pixinsight\IntegrateTotality.js"
        $manifestPath = "C:\Temp\AsiToPix-Totality\IntegrationPlan.json"
        $ipcScriptPath = "C:\Temp\AsiToPix-Totality\IntegrateTotality-IPC.js"

        $arguments = @(Get-AsiToPixPixInsightArgumentList `
            -ScriptPath $scriptPath `
            -ManifestPath $manifestPath `
            -Mode Reuse `
            -IpcScriptPath $ipcScriptPath)

        $arguments.Count | Should Be 1
        $arguments[0] | Should Be ('--execute="{0}"' -f $ipcScriptPath)
        @($arguments | Where-Object { $_ -in @("-n", "--automation-mode", "--force-exit") }).Count | Should Be 0
    }

    It "rejects comma-delimited automation paths before launching PixInsight" {
        Assert-TestTotalityError `
            -Script {
                Get-AsiToPixPixInsightArgumentList `
                    -ScriptPath "C:\Tools,Old\IntegrateTotality.js" `
                    -ManifestPath "C:\Temp\Block.json"
            } `
            -MessagePattern "cannot contain commas"
    }

    It "does not create the output directory or launch PixInsight in WhatIf mode" {
        $outputDirectory = Join-Path -Path $TestDrive -ChildPath "whatif-output"
        $frames = @(31..33 | ForEach-Object {
            $frame = Get-TestTotalityFrame -Index $_
            Set-Content -LiteralPath $frame.Path -Value "frame"
            $frame
        })
        $plan = @(Get-AsiToPixTotalityIntegrationPlan `
            -Frames $frames `
            -OutputDirectory $outputDirectory `
            -BlockLength 3 `
            -SequenceStartIndex 1)
        $pixInsightPath = Join-Path -Path $TestDrive -ChildPath "PixInsight.exe"
        $scriptPath = Join-Path -Path $TestDrive -ChildPath "IntegrateTotality.js"
        Set-Content -LiteralPath $pixInsightPath -Value "not launched"
        Set-Content -LiteralPath $scriptPath -Value "not launched"

        $result = Invoke-AsiToPixTotalityIntegrationPlan `
            -Plan $plan `
            -PixInsightPath $pixInsightPath `
            -PixInsightScriptPath $scriptPath `
            -WhatIf

        $result.IntegratedCount | Should Be 0
        $result.NotApprovedCount | Should Be 1
        Test-Path -LiteralPath $outputDirectory | Should Be $false
    }

    It "skips a non-contiguous nominal block before launching PixInsight" {
        $outputDirectory = Join-Path -Path $TestDrive -ChildPath "non-contiguous-output"
        $frames = @(1, 2, 4 | ForEach-Object { Get-TestTotalityFrame -Index $_ })
        $plan = @(Get-AsiToPixTotalityIntegrationPlan `
            -Frames $frames `
            -OutputDirectory $outputDirectory `
            -BlockLength 5 `
            -SequenceStartIndex 0)
        $pixInsightPath = Join-Path -Path $TestDrive -ChildPath "skip-PixInsight.exe"
        $scriptPath = Join-Path -Path $TestDrive -ChildPath "skip-IntegrateTotality.js"
        Set-Content -LiteralPath $pixInsightPath -Value "not launched"
        Set-Content -LiteralPath $scriptPath -Value "not launched"

        $result = Invoke-AsiToPixTotalityIntegrationPlan `
            -Plan $plan `
            -PixInsightPath $pixInsightPath `
            -PixInsightScriptPath $scriptPath

        $result.IntegratedCount | Should Be 0
        $result.RejectedCount | Should Be 1
        Test-Path -LiteralPath $outputDirectory | Should Be $false
    }

    It "skips an existing integration output without overwriting it" {
        $outputDirectory = Join-Path -Path $TestDrive -ChildPath "existing-output"
        New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
        $frames = @(42..44 | ForEach-Object {
            $frame = Get-TestTotalityFrame -Index $_
            Set-Content -LiteralPath $frame.Path -Value "frame"
            $frame
        })
        $plan = @(Get-AsiToPixTotalityIntegrationPlan `
            -Frames $frames `
            -OutputDirectory $outputDirectory `
            -BlockLength 3)
        Set-Content -LiteralPath $plan[0].OutputPath -Value "keep me"
        $pixInsightPath = Join-Path -Path $TestDrive -ChildPath "existing-PixInsight.exe"
        $scriptPath = Join-Path -Path $TestDrive -ChildPath "existing-IntegrateTotality.js"
        Set-Content -LiteralPath $pixInsightPath -Value "not launched"
        Set-Content -LiteralPath $scriptPath -Value "not launched"

        $result = Invoke-AsiToPixTotalityIntegrationPlan `
            -Plan $plan `
            -PixInsightPath $pixInsightPath `
            -PixInsightScriptPath $scriptPath

        $result.IntegratedCount | Should Be 0
        $result.ExistingCount | Should Be 1
        Get-Content -LiteralPath $plan[0].OutputPath | Should Be "keep me"
    }

    It "launches one PixInsight process for all approved blocks" {
        $integrationRoot = Join-Path -Path $TestDrive -ChildPath "single-instance"
        New-Item -ItemType Directory -Path $integrationRoot -Force | Out-Null
        $env:ASITOPIX_TOTALITY_TEST_ROOT = $integrationRoot

        InModuleScope AsiToPix.TotalityIntegration {
            $outputDirectory = Join-Path -Path $env:ASITOPIX_TOTALITY_TEST_ROOT -ChildPath "integrated"
            $frames = @(51..56 | ForEach-Object {
                $path = Join-Path -Path $env:ASITOPIX_TOTALITY_TEST_ROOT -ChildPath ("frame_{0:D5}.xisf" -f $_)
                Set-Content -LiteralPath $path -Value "frame"
                [PSCustomObject]@{
                    Index = $_
                    Path  = $path
                }
            })
            $plan = @(Get-AsiToPixTotalityIntegrationPlan `
                -Frames $frames `
                -OutputDirectory $outputDirectory `
                -BlockLength 3)
            $pixInsightPath = Join-Path -Path $env:ASITOPIX_TOTALITY_TEST_ROOT -ChildPath "PixInsight.exe"
            $scriptPath = Join-Path -Path $env:ASITOPIX_TOTALITY_TEST_ROOT -ChildPath "IntegrateTotality.js"
            Set-Content -LiteralPath $pixInsightPath -Value "mock executable"
            Set-Content -LiteralPath $scriptPath -Value "mock script"

            Mock Start-Process {
                param(
                    $FilePath,
                    $ArgumentList,
                    $WindowStyle,
                    $RedirectStandardOutput,
                    $RedirectStandardError,
                    $PassThru
                )
                $null = @(
                    $FilePath,
                    $WindowStyle,
                    $RedirectStandardOutput,
                    $RedirectStandardError,
                    $PassThru
                )

                $runArgument = @($ArgumentList | Where-Object { $_ -like '-r=*' })[0]
                $manifestPath = ($runArgument -split ',manifest=', 2)[1].TrimEnd('"')
                $manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
                foreach ($block in $manifest.blocks) {
                    Set-Content -LiteralPath $block.outputPath -Value "integrated"
                }
                [ordered]@{
                    schemaVersion         = 1
                    state                 = "completed"
                    currentBlockNumber    = $null
                    completedBlockNumbers = @($manifest.blocks.blockNumber)
                    errorMessage          = ""
                    errorStack            = ""
                } | ConvertTo-Json | Set-Content -LiteralPath $manifest.statusPath -Encoding UTF8

                $process = [PSCustomObject]@{
                    ExitCode = 0
                    HasExited = $true
                }
                $process | Add-Member -MemberType ScriptMethod -Name Refresh -Value { }
                $process | Add-Member -MemberType ScriptMethod -Name WaitForExit -Value { }
                $process
            }

            $result = Invoke-AsiToPixTotalityIntegrationPlan `
                -Plan $plan `
                -PixInsightPath $pixInsightPath `
                -PixInsightScriptPath $scriptPath `
                -Confirm:$false

            $result.IntegratedCount | Should Be 2
            Test-Path -LiteralPath $plan[0].OutputPath -PathType Leaf | Should Be $true
            Test-Path -LiteralPath $plan[1].OutputPath -PathType Leaf | Should Be $true
            Assert-MockCalled Start-Process -Times 1
        }

        Remove-Item Env:\ASITOPIX_TOTALITY_TEST_ROOT
    }

    It "reuses a running PixInsight instance without creating or closing one" {
        $integrationRoot = Join-Path -Path $TestDrive -ChildPath "reuse-instance"
        New-Item -ItemType Directory -Path $integrationRoot -Force | Out-Null
        $env:ASITOPIX_TOTALITY_REUSE_TEST_ROOT = $integrationRoot

        InModuleScope AsiToPix.TotalityIntegration {
            $outputDirectory = Join-Path -Path $env:ASITOPIX_TOTALITY_REUSE_TEST_ROOT -ChildPath "integrated"
            $frames = @(61..63 | ForEach-Object {
                $path = Join-Path -Path $env:ASITOPIX_TOTALITY_REUSE_TEST_ROOT -ChildPath ("frame_{0:D5}.xisf" -f $_)
                Set-Content -LiteralPath $path -Value "frame"
                [PSCustomObject]@{
                    Index = $_
                    Path  = $path
                }
            })
            $plan = @(Get-AsiToPixTotalityIntegrationPlan `
                -Frames $frames `
                -OutputDirectory $outputDirectory `
                -BlockLength 3 `
                -SequenceStartIndex 1)
            $pixInsightPath = Join-Path -Path $env:ASITOPIX_TOTALITY_REUSE_TEST_ROOT -ChildPath "PixInsight.exe"
            $scriptPath = Join-Path -Path $env:ASITOPIX_TOTALITY_REUSE_TEST_ROOT -ChildPath "IntegrateTotality.js"
            Set-Content -LiteralPath $pixInsightPath -Value "mock executable"
            Copy-Item -LiteralPath (Join-Path -Path $PSScriptRoot -ChildPath "..\pixinsight\IntegrateTotality.js") `
                -Destination $scriptPath

            $script:reuseArguments = @()
            $script:ipcManifestWasInjected = $false
            Mock Get-Process { [PSCustomObject]@{ Id = 12345 } }
            Mock Start-Process {
                param(
                    $FilePath,
                    $ArgumentList,
                    $WindowStyle,
                    $RedirectStandardOutput,
                    $RedirectStandardError,
                    $PassThru
                )
                $null = @(
                    $FilePath,
                    $WindowStyle,
                    $RedirectStandardOutput,
                    $RedirectStandardError,
                    $PassThru
                )
                $script:reuseArguments = @($ArgumentList)
                $ipcScriptPath = $ArgumentList[0].Substring('--execute='.Length).Trim('"')
                $ipcScriptText = Get-Content -LiteralPath $ipcScriptPath -Raw -Encoding UTF8
                $manifestMatch = [regex]::Match(
                    $ipcScriptText,
                    'var ASITOPIX_MANIFEST_PATH = (?<literal>".*?");'
                )
                $script:ipcManifestWasInjected = $manifestMatch.Success
                $manifestPath = $manifestMatch.Groups["literal"].Value | ConvertFrom-Json
                $manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
                foreach ($block in $manifest.blocks) {
                    Set-Content -LiteralPath $block.outputPath -Value "integrated"
                }
                [ordered]@{
                    schemaVersion         = 1
                    state                 = "completed"
                    currentBlockNumber    = $null
                    completedBlockNumbers = @($manifest.blocks.blockNumber)
                    errorMessage          = ""
                    errorStack            = ""
                } | ConvertTo-Json | Set-Content -LiteralPath $manifest.statusPath -Encoding UTF8

                $process = [PSCustomObject]@{
                    ExitCode = 0
                    HasExited = $true
                }
                $process | Add-Member -MemberType ScriptMethod -Name Refresh -Value { }
                $process | Add-Member -MemberType ScriptMethod -Name WaitForExit -Value { }
                $process
            }

            $result = Invoke-AsiToPixTotalityIntegrationPlan `
                -Plan $plan `
                -PixInsightPath $pixInsightPath `
                -PixInsightScriptPath $scriptPath `
                -PixInsightMode Reuse `
                -Confirm:$false

            $result.IntegratedCount | Should Be 1
            $script:ipcManifestWasInjected | Should Be $true
            $script:reuseArguments.Count | Should Be 1
            $script:reuseArguments[0] | Should Match '^--execute='
            @($script:reuseArguments | Where-Object { $_ -in @("-n", "--automation-mode", "--force-exit") }).Count |
                Should Be 0
            Assert-MockCalled Start-Process -Times 1
        }

        Remove-Item Env:\ASITOPIX_TOTALITY_REUSE_TEST_ROOT
    }

    It "contains the requested ImageIntegration settings and saves the integrated window" {
        $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath "..\pixinsight\IntegrateTotality.js"
        $scriptText = Get-Content -LiteralPath $scriptPath -Raw -Encoding UTF8

        $scriptText | Should Match 'P\.rejection = ImageIntegration\.prototype\.WinsorizedSigmaClip;'
        $scriptText | Should Match 'P\.sigmaLow = 3\.500;'
        $scriptText | Should Match 'P\.sigmaHigh = 4\.000;'
        $scriptText | Should Match 'P\.clipHigh = false;'
        $scriptText | Should Not Match 'P\.clipHigh = true;'
        $scriptText | Should Match 'P\.largeScaleClipLow = true;'
        $scriptText | Should Match 'P\.generateRejectionMaps = true;'
        $scriptText | Should Match 'P\.showImages = true;'
        $scriptText | Should Match 'P\.images = images;'
        $scriptText | Should Not Match 'P\.images\.push'
        $scriptText | Should Match 'for \( var i = 0; i < manifest\.blocks\.length; \+\+i \)'
        $scriptText | Should Match 'integrationWindow\.saveAs\( block\.outputPath, false, false, false, false \);'
        $scriptText | Should Match 'File\.writeTextFile\( statusPath, JSON\.stringify\( status, null, 2 \) \);'
        $scriptText | Should Match 'typeof ASITOPIX_MANIFEST_PATH != "undefined"'
        $scriptText | Should Match '// ASITOPIX_IPC_MANIFEST'
    }
}
