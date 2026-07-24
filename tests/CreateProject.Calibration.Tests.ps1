$scriptPath = Join-Path -Path $PSScriptRoot -ChildPath "..\CreateProject.ps1"
$imageFilesModulePath = Join-Path -Path $PSScriptRoot -ChildPath "..\src\AsiToPix.ImageFiles.psm1"
$frameFoldersModulePath = Join-Path -Path $PSScriptRoot -ChildPath "..\src\AsiToPix.FrameFolders.psm1"
Import-Module $imageFilesModulePath -Force
Import-Module $frameFoldersModulePath -Force

$tokens = $null
$parseErrors = $null
$scriptAst = [System.Management.Automation.Language.Parser]::ParseFile(
    $scriptPath,
    [ref]$tokens,
    [ref]$parseErrors
)
$functionNames = @(
    'ConvertTo-CreateProjectDate',
    'Get-CreateProjectNightDate',
    'Get-CreateProjectDateFromFileName',
    'Get-CreateProjectDateFromPath',
    'Test-CreateProjectPathHasMonthDate',
    'Get-CreateProjectDateDiff',
    'ConvertTo-CreateProjectExposureNumber',
    'ConvertTo-CreateProjectTemperatureFolderName',
    'Get-CreateProjectCalibrationFile',
    'Test-CreateProjectExposureMatch',
    'Get-CreateProjectSourceCounterpartPath',
    'Get-CreateProjectCalibrationDate',
    'ConvertTo-CreateProjectTemperatureFolderKey',
    'Get-CreateProjectTemperatureDirectory',
    'Get-CalibPath'
)
$functionAsts = @($scriptAst.FindAll({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
        $node.Name -in $functionNames
}, $true))
foreach ($functionAst in $functionAsts) {
    . ([scriptblock]::Create($functionAst.Extent.Text))
}

function Write-AsiToPixCyrillicPathWarning {
    param($Path, $Context)
}

function Write-CreateProjectCalibrationDirectoryWarning {
    param($RootPath)
}

Describe "CreateProject calibration selection" {
    It "parses exposure-suffixed session folder dates" {
        (ConvertTo-CreateProjectDate -DateText '26.07.17-60s').ToString('yyyy-MM-dd') |
            Should Be '2026-07-17'
    }

    It "prefers a mirrored master over its more precisely dated Source files" {
        $cameraRoot = Join-Path -Path $TestDrive -ChildPath 'Calibration\ASI2600MM'
        $relativePath = 'darks\Gain120\-10C\60sec\25.09'
        $masterPath = Join-Path -Path (Join-Path -Path $cameraRoot -ChildPath 'Master') -ChildPath $relativePath
        $sourcePath = Join-Path -Path (Join-Path -Path $cameraRoot -ChildPath 'Source') -ChildPath $relativePath
        [void](New-Item -ItemType Directory -Path $masterPath -Force)
        [void](New-Item -ItemType Directory -Path $sourcePath -Force)
        [void](New-Item -ItemType File -Path (Join-Path -Path $masterPath -ChildPath 'masterDark_BIN-1_6248x4176.xisf'))
        [void](New-Item -ItemType File -Path (
            Join-Path -Path $sourcePath -ChildPath 'Dark_60.0s_Bin1_2600MM_L_gain120_20250917-001631_-10.0C_0010.fit'
        ))

        $plainSession = Get-CalibPath 'Darks' '120' '-10' '60s' $cameraRoot '26.07.18'
        $suffixedSession = Get-CalibPath 'Darks' '120' '-10' '60s' $cameraRoot '26.07.17-60s'

        $plainSession.Mode | Should Be 'Master'
        $plainSession.Path | Should Be $masterPath
        $suffixedSession.Mode | Should Be 'Master'
        $suffixedSession.Path | Should Be $masterPath
    }
}
