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
}
