$modulePath = Join-Path -Path $PSScriptRoot -ChildPath "..\src\AsiToPix.ProjectMetadata.psm1"
Import-Module $modulePath -Force

Describe "Project calibration metadata" {
    It "uses flat-darks as the project Source folder name" {
        Get-AsiToPixProjectSourceFolderName -Type "FlatDarks" | Should Be "flat-darks"
        Get-AsiToPixProjectSourceFolderName -Type "flat-darks" | Should Be "flat-darks"
        Get-AsiToPixProjectSourceFolderName -Type "Darks" | Should Be "darks"
        Get-AsiToPixProjectSourceFolderName -Type "dArK" | Should Be "darks"
        Get-AsiToPixProjectSourceFolderName -Type "BIASES" | Should Be "biases"
        Get-AsiToPixProjectSourceFolderName -Type "fLaT" | Should Be "flats"
        Get-AsiToPixProjectSourceFolderName -Type "LIGHT" | Should Be "lights"
    }

    It "maps a Source calibration folder to the equivalent Master destination" {
        $cameraRoot = Join-Path -Path $TestDrive -ChildPath "Calibration\ASI2600MM"
        $masterRoot = Join-Path -Path $cameraRoot -ChildPath "Master\darks"
        $sourcePath = Join-Path -Path $cameraRoot -ChildPath "Source\darks\Gain120\-10C\120sec\26.07"
        $expectedDestination = Join-Path -Path $masterRoot -ChildPath "Gain120\-10C\120sec\26.07"
        $camera = [PSCustomObject]@{
            Name = "ASI2600MM"
            CalibrationFolders = [PSCustomObject]@{ Darks = $masterRoot }
        }
        $pendingLink = [PSCustomObject]@{
            Type = "Darks"
            Tag = "Gain_120_Temp_-10C_Exp_120s_Target_H_Filter_H_Cam_ASI2600MM"
            Src = $sourcePath
            Cam = "ASI2600MM"
            Gain = "120"
            Temperature = "-10C"
            Exposure = "120s"
            Filter = "H"
            Session = "26.07.08"
            Target = "H"
        }

        $metadata = @(ConvertTo-AsiToPixCalibrationSourceMetadata -PendingLink @($pendingLink) -CameraMetadata @($camera))

        $metadata.Count | Should Be 1
        $metadata[0].SourceMode | Should Be "Source"
        $metadata[0].DestinationFolder | Should Be $expectedDestination
        $metadata[0].Gain | Should Be "120"
        $metadata[0].TemperatureC | Should Be "-10"
        $metadata[0].ExposureSeconds | Should Be "120"
    }

    It "keeps an existing Master calibration folder as its export destination" {
        $cameraRoot = Join-Path -Path $TestDrive -ChildPath "Calibration\ASI2600MM"
        $masterRoot = Join-Path -Path $cameraRoot -ChildPath "Master\biases"
        $sourcePath = Join-Path -Path $masterRoot -ChildPath "Gain120\-10C\26.07"
        $camera = [PSCustomObject]@{
            Name = "ASI2600MM"
            CalibrationFolders = [PSCustomObject]@{ Biases = $masterRoot }
        }
        $pendingLink = [PSCustomObject]@{
            Type = "Biases"
            Tag = "Gain_120_Temp_-10C_Target_H_Filter_H_Cam_ASI2600MM"
            Src = $sourcePath
            Cam = "ASI2600MM"
            Gain = "120"
            Temperature = "-10C"
            Filter = "H"
            Session = "26.07.08"
            Target = "H"
        }

        $metadata = @(ConvertTo-AsiToPixCalibrationSourceMetadata -PendingLink @($pendingLink) -CameraMetadata @($camera))

        $metadata.Count | Should Be 1
        $metadata[0].SourceMode | Should Be "Master"
        $metadata[0].DestinationFolder | Should Be $sourcePath
    }

    It "rejects calibration sources outside the configured Source and Master roots" {
        $cameraRoot = Join-Path -Path $TestDrive -ChildPath "Calibration\ASI2600MM"
        $masterRoot = Join-Path -Path $cameraRoot -ChildPath "Master\darks"
        $outsidePath = Join-Path -Path $TestDrive -ChildPath "Other\darks\Gain120\-10C\120sec\26.07"
        $camera = [PSCustomObject]@{
            Name = "ASI2600MM"
            CalibrationFolders = [PSCustomObject]@{ Darks = $masterRoot }
        }
        $pendingLink = [PSCustomObject]@{
            Type = "Darks"
            Tag = "outside"
            Src = $outsidePath
            Cam = "ASI2600MM"
            Gain = "120"
            Temperature = "-10C"
            Exposure = "120s"
        }

        { ConvertTo-AsiToPixCalibrationSourceMetadata -PendingLink @($pendingLink) -CameraMetadata @($camera) } |
            Should Throw
    }
}
