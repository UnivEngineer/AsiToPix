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
}
