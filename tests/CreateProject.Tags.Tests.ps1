Describe "CreateProject WBPP calibration tags" {
    $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath "..\CreateProject.ps1"
    $scriptText = Get-Content -Path $scriptPath -Raw

    It "keeps real filter and target values for dark calibration tags" {
        $scriptText | Should Match '\$dTag = "Gain_\$\{curGain\}_Temp_\$\{roundT\}_Exp_\$\{curExp\}_Target_\$\{targetGrp\}_Filter_\$\{filt\}_Cam_\$\{sessionCamFull\}"'
    }

    It "keeps real filter and target values for bias calibration tags" {
        $scriptText | Should Match '\$bTag = "Gain_\$\{curGain\}_Temp_\$\{roundT\}_Target_\$\{targetGrp\}_Filter_\$\{filt\}_Cam_\$\{sessionCamFull\}"'
    }

    It "keeps real session, filter, target, and exposure values for flat calibration tags" {
        $scriptText | Should Match '\$flatTag = "Session_\$\{sessionDate\}_Filter_\$\{filt\}_Target_\$\{targetGrp\}_Gain_\$\{curGain\}_Temp_\$\{roundT\}_Exp_\$\{curExp\}_Cam_\$\{sessionCamFull\}"'
    }

    It "keeps real flat-dark calibration tag values" {
        $scriptText | Should Match '\$fdTag = "Exp_\$\{fdExp\}_Gain_\$\{curGain\}_Temp_\$\{roundT\}_Target_\$\{targetGrp\}_Filter_\$\{filt\}_Cam_\$\{sessionCamFull\}"'
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
        $scriptText | Should Match 'Resolve-CreateProjectInputPath -InputPath \$inputPath -AstroPhotoRoot \$baseZ'
        $scriptText | Should Not Match 'Join-Path -Path \$AstroPhotoRoot -ChildPath "Import"'
    }

    It "warns about mixed light exposures without splitting CreateProject links" {
        $scriptText | Should Match 'Mixed light exposures in ASIAir session folder'
        $scriptText | Should Match 'ImportSession.ps1 or ImportAll.ps1'
        $scriptText | Should Not Match 'function Get-CreateProjectLightFileGroup'
        $scriptText | Should Not Match 'SourceFiles ='
        $scriptText | Should Not Match 'Copy-CreateProjectFileSet'
    }

    It "keeps the original flat selection prompt behavior" {
        $scriptText | Should Match 'Select Index \(Enter to accept default\$dtext\)'
        $scriptText | Should Not Match '\$flatSelectionCache'
        $scriptText | Should Not Match 'or S to skip'
        $scriptText | Should Not Match 'Flats skipped for'
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
