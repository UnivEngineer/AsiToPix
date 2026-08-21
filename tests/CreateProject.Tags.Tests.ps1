Describe "CreateProject WBPP calibration tags" {
    $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath "..\CreateProject.ps1"
    $scriptText = Get-Content -Path $scriptPath -Raw

    It "keeps real filter and target values for dark calibration tags" {
        $scriptText | Should Match '\$dTag = "FlatSet_\$\{flatSetId\}_Gain_\$\{curGain\}_Temp_\$\{roundT\}_Exp_\$\{curExp\}_Target_\$\{targetGrp\}_Filter_\$\{filt\}_Cam_\$\{sessionCamFull\}"'
    }

    It "keeps real filter and target values for bias calibration tags" {
        $scriptText | Should Match '\$bTag = "FlatSet_\$\{flatSetId\}_Gain_\$\{curGain\}_Temp_\$\{roundT\}_Target_\$\{targetGrp\}_Filter_\$\{filt\}_Cam_\$\{sessionCamFull\}"'
    }

    It "builds flat tags from the physical flat set without light session or exposure" {
        $scriptText | Should Match '\$flatTag = ConvertTo-AsiToPixFlatSetTag'
        $scriptText | Should Match '-FlatDate \$flatDate'
        $scriptText | Should Match '-Angle \$flatAngle'
        $scriptText | Should Match '-Binning \$flatBinning'
        $scriptText | Should Not Match '\$flatTag = "Session_\$\{sessionDate\}'
        $scriptText | Should Not Match '\$flatTag = .*Exp_\$\{curExp\}'
    }

    It "keeps real flat-dark calibration tag values" {
        $scriptText | Should Match '\$fdTag = "FlatSet_\$\{flatSetId\}_Exp_\$\{fdExp\}_Gain_\$\{flatGain\}_Temp_\$\{flatTemperature\}_Target_\$\{targetGrp\}_Filter_\$\{filt\}_Cam_\$\{sessionCamFull\}"'
    }

    It "deduplicates planned flat links by canonical target before applying the project" {
        $scriptText | Should Match 'Get-AsiToPixUniqueFlatPlan -PendingLink \$pendingLinks'
        $scriptText | Should Match '\$pendingLinks = @\(\$flatPlan\.PendingLinks\)'
        $scriptText | Should Match 'Write-CreateProjectFlatPlanWarning -FlatPlan \$flatPlan'
        $scriptText | Should Match 'physical flat-set identifier'
        $scriptText | Should Match 'FLATSET\s+: physical flat-set identifier \(add to pre only after a flat-set collision warning\)'
        $scriptText | Should Match 'EXP\s+: optional custom grouping\s+\(WBPP also reads image headers\)'
    }

    It "reports repeated source calibration paths without changing link tags" {
        $scriptText | Should Match 'function Write-CreateProjectDuplicateCalibrationWarning'
        $scriptText | Should Match 'Repeated Source calibration folders detected'
        $scriptText | Should Match '\$_.Display -notlike "Master\\\*"'
        $scriptText | Should Match 'Full WBPP tags are kept'
        $scriptText | Should Match 'Run WBPP once on the unique calibration set'
        $scriptText | Should Match 'Write-CreateProjectDuplicateCalibrationWarning -PendingLink \$pendingLinks'
        $scriptText | Should Not Match 'Select-CreateProjectUniquePendingLink'
        $scriptText | Should Not Match 'Target_DUMMY|Filter_DUMMY|Session_DUMMY|Exp_DUMMY'
    }

    It "handles repeated source calibrations with one unique tag under StrictMode" {
        $tokens = $null
        $parseErrors = $null
        $scriptAst = [System.Management.Automation.Language.Parser]::ParseFile(
            $scriptPath,
            [ref]$tokens,
            [ref]$parseErrors
        )
        $functionNames = @(
            'Get-CreateProjectDuplicateCalibrationKey',
            'Write-CreateProjectDuplicateCalibrationWarning'
        )
        $functionAsts = @($scriptAst.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
                $node.Name -in $functionNames
        }, $true))
        $pendingLinks = @(1..3 | ForEach-Object {
            [PSCustomObject]@{
                Type = 'Darks'
                Tag = 'FlatSet_same_Gain_120_Temp_-10C_Exp_60s_Target_L_Filter_L_Cam_ASI2600MM'
                Src = 'C:\Calibration\ASI2600MM\Source\darks\Gain120\-10C\60sec\25.09'
                Display = 'Source\darks\Gain120\-10C\60sec\25.09'
            }
        })

        {
            & {
                Set-StrictMode -Version Latest
                foreach ($functionAst in $functionAsts) {
                    . ([scriptblock]::Create($functionAst.Extent.Text))
                }
                Write-CreateProjectDuplicateCalibrationWarning -PendingLink $pendingLinks
            }
        } | Should Not Throw
    }

    It "does not use prefix matching for calibration exposure folders" {
        $scriptText | Should Match 'ConvertTo-CreateProjectExposureNumber -ExposureText \$pathPart'
        $scriptText | Should Match '\$null -ne \$folderExp -and \$folderExp -eq \$cleanExp'
        $scriptText | Should Not Match '\$pathPart -like "\$cleanExp\*"'
    }

    It "normalizes fractional temperatures independently of the current locale" {
        $tokens = $null
        $parseErrors = $null
        $scriptAst = [System.Management.Automation.Language.Parser]::ParseFile(
            $scriptPath,
            [ref]$tokens,
            [ref]$parseErrors
        )
        $functionAst = $scriptAst.Find({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
                $node.Name -eq 'ConvertTo-CreateProjectTemperatureFolderName'
        }, $true)
        . ([scriptblock]::Create($functionAst.Extent.Text))

        ConvertTo-CreateProjectTemperatureFolderName -Temperature '-9.8' | Should Be '-10C'
        ConvertTo-CreateProjectTemperatureFolderName -Temperature '-9,8' | Should Be '-10C'
        ConvertTo-CreateProjectTemperatureFolderName -Temperature ([double]-9.8) | Should Be '-10C'
    }

    It "converts millisecond flat exposures to seconds" {
        $tokens = $null
        $parseErrors = $null
        $scriptAst = [System.Management.Automation.Language.Parser]::ParseFile(
            $scriptPath,
            [ref]$tokens,
            [ref]$parseErrors
        )
        $functionAst = $scriptAst.Find({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
                $node.Name -eq 'ConvertTo-CreateProjectExposureNumber'
        }, $true)
        . ([scriptblock]::Create($functionAst.Extent.Text))

        ConvertTo-CreateProjectExposureNumber -ExposureText '490.0ms' | Should Be '0.49'
        ConvertTo-CreateProjectExposureNumber -ExposureText '180s' | Should Be '180'
    }

    It "resolves object-name input from the ASIAir archive" {
        $scriptText | Should Match 'function Resolve-CreateProjectInputPath'
        $scriptText | Should Match 'function Find-CreateProjectAsiairSourceCandidate'
        $scriptText | Should Match 'Join-Path -Path \$AstroPhotoRoot -ChildPath "ASIAir"'
        $scriptText | Should Match 'Matching ASIAir projects'
        $scriptText | Should Match 'Resolve-CreateProjectInputPath -InputPath \$inputPath -AstroPhotoRoot \$sourceAstroPhotoRoot'
        $scriptText | Should Not Match 'Join-Path -Path \$AstroPhotoRoot -ChildPath "Import"'
    }

    It "uses independent source and destination astrophotography roots" {
        $scriptText | Should Match '\[string\]\$SourceAstroPhotoRoot'
        $scriptText | Should Match '\[string\]\$DestinationAstroPhotoRoot'
        $scriptText | Should Match 'Join-Path -Path \$sourceAstroPhotoRoot -ChildPath "ASIAir"'
        $scriptText | Should Match 'Join-Path -Path \$sourceAstroPhotoRoot -ChildPath "Calibration"'
        $scriptText | Should Match 'Join-Path -Path \$destinationAstroPhotoRoot -ChildPath "Processing"'
        $scriptText | Should Match 'Source root\s+\(read\)'
        $scriptText | Should Match 'Destination root \(write\)'
        $scriptText | Should Match '-SelectionPrompt "Select SOURCE root folder \(ASIAir and Calibration\)"'
        $scriptText | Should Match '-SelectionPrompt "Select DESTINATION folder \(Processing/project\)"'
        $scriptText | Should Match 'Write-Host ""\s+\$inputPath = \(Read-Host "Paste lights path, supported image file, or ASIAir object name"\)'
        $scriptText | Should Not Match '\$baseZ'
    }

    It "warns about mixed light exposures without splitting CreateProject links" {
        $scriptText | Should Match 'Mixed light exposures in ASIAir session folder'
        $scriptText | Should Match 'ImportSession.ps1 or ImportAll.ps1'
        $scriptText | Should Not Match 'function Get-CreateProjectLightFileGroup'
    }

    It "offers preview mode after the OSC merge prompt with a blank line and No as the default" {
        $previewPrompt = '-Prompt "Use preview mode (10-15 frames per filter)?"'
        $previewPromptIndex = $scriptText.IndexOf($previewPrompt)
        $cameraSummaryIndex = $scriptText.IndexOf('--- CAMERA MAPPING SUMMARY ---')
        $promptWithDefaultPattern = [regex]::Escape($previewPrompt) + '\s+`\s*\r?\n\s+-DefaultYes \$false'
        $promptPlacementPattern = 'Allow WBPP to merge OSC sessions with filters IRC and Trib into a single one L\?"\s*\r?\n\s*Write-Host ""\s*\r?\n\s*\$previewMode = Read-CreateProjectConfirmation'

        $scriptText | Should Match 'Select-AsiToPixPreviewLightPlan -PendingLink \$pendingLinks -MaxFramesPerFilter 15'
        $scriptText.Contains($previewPrompt) | Should Be $true
        $scriptText | Should Match $promptWithDefaultPattern
        $scriptText | Should Match $promptPlacementPattern
        ($previewPromptIndex -ge 0) | Should Be $true
        ($previewPromptIndex -lt $cameraSummaryIndex) | Should Be $true
    }

    It "creates preview projects from selected files instead of whole light folders" {
        $scriptText | Should Match 'SourceFiles\s+= @\(\$imageFiles \| Select-Object -ExpandProperty FullName\)'
        $scriptText | Should Match 'function New-CreateProjectPreviewSymbolicLinkSet'
        $scriptText | Should Match 'function Copy-CreateProjectFileSet'
        $scriptText | Should Match 'Join-Path -Path \$Destination -ChildPath \(Split-Path -Path \$sourcePath -Leaf\)'
        $scriptText | Should Match 'PreviewMode\s+= \$previewMode'
        $scriptText | Should Match 'PreviewFrameCount\s+= if'
    }

    It "preserves full preview filenames and uses extended Windows paths for long symlinks" {
        $tokens = $null
        $parseErrors = $null
        $scriptAst = [System.Management.Automation.Language.Parser]::ParseFile(
            $scriptPath,
            [ref]$tokens,
            [ref]$parseErrors
        )
        $functionNames = @(
            'ConvertTo-CreateProjectExtendedPath'
        )
        $functionAsts = @($scriptAst.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
                $node.Name -in $functionNames
        }, $true))
        foreach ($functionAst in $functionAsts) {
            . ([scriptblock]::Create($functionAst.Extent.Text))
        }

        ConvertTo-CreateProjectExtendedPath -Path 'C:\Astro\frame.fit' |
            Should Be '\\?\C:\Astro\frame.fit'
        ConvertTo-CreateProjectExtendedPath -Path '\\server\Astro\frame.fit' |
            Should Be '\\?\UNC\server\Astro\frame.fit'
        $scriptText | Should Match 'Link \(\$\(\$Path\.Length\) chars\)'
        $scriptText | Should Match 'Target \(\$\(\$Target\.Length\) chars\)'
        $scriptText | Should Not Match 'preview_\{0:D3\}'
        $scriptText | Should Match 'Copy-Item -LiteralPath \$copySourcePath -Destination \$copyTargetPath'
    }

    It "keeps the original flat selection prompt behavior" {
        $scriptText | Should Match 'Select Index \(Enter to accept default\$dtext\)'
        $scriptText | Should Not Match '\$flatSelectionCache'
        $scriptText | Should Not Match 'or S to skip'
        $scriptText | Should Not Match 'Flats skipped for'
    }

    It "offers to reuse previous flat selections and asks only for unmatched nights" {
        $previousRunIndex = $scriptText.IndexOf('Previous run detected:')
        $scannerIndex = $scriptText.IndexOf('# SCANNER - iterates filter-group folders')

        $scriptText | Should Match 'Join-Path -Path \$setupRoot -ChildPath "project_meta\.json"'
        $scriptText | Should Match 'Get-AsiToPixPreviousFlatSelectionPlan'
        $scriptText | Should Match 'Previous run detected: \{0\} \{1\} \{2\} \{3\} \{4\} selected'
        $scriptText | Should Match '-Prompt "Reuse flat selection from the previous run\?"\s+`\s*\r?\n\s+-DefaultYes \$true'
        $scriptText | Should Match 'Get-AsiToPixFlatSelectionKey'
        $scriptText | Should Match '\$previousFlatSelectionByKey\.ContainsKey\(\$previousSelectionKey\)'
        $scriptText | Should Match 'Previous flat selection is unavailable and will be requested again'
        $scriptText | Should Match 'if \(\$null -eq \$fFound\) \{\s+# Build list of all raw filter names'
        ($previousRunIndex -ge 0) | Should Be $true
        ($previousRunIndex -lt $scannerIndex) | Should Be $true
    }

    It "maps unfiltered OSC lights to RGB without merging them into L" {
        $scriptText | Should Match "'\^\(None\|RGB\)\$'"
        $scriptText | Should Match 'Filter="RGB";\s+Target="RGB"'
        $scriptText | Should Match '\$rawFilt -eq "IRC" -or \$rawFilt -eq "Trib"'
        $scriptText | Should Not Match '\$rawFilt -eq "None" -or \$rawFilt -eq "IRC"'
    }

    It "shows the supported image count for each light session in the project tree" {
        $scriptText | Should Match 'FrameCount\s+= \$imageFiles\.Count'
        $scriptText | Should Match '\$frameCountPart = if \(\$_\.Type -eq "Lights"\)'
        $scriptText | Should Match 'Write-Host " \|- \$frameCountPart\$\(\$_\.Tag\.PadRight\(65\)\)'
    }

    It "uses the shared PixInsight image format helper" {
        $scriptText | Should Match 'AsiToPix\.ImageFiles\.psm1'
        $scriptText | Should Match 'Test-AsiToPixSupportedImageFileName'
        $scriptText | Should Not Match 'Filter "\*\.fit\*"'
    }
}
