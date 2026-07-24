$modulePath = Join-Path -Path $PSScriptRoot -ChildPath "..\src\AsiToPix.CleanupWbpp.psm1"
Import-Module $modulePath -Force

Describe "WBPP cleanup planning" {
    It "keeps only masterLight files as cleanup exclusions" {
        $processingRoot = Join-Path $TestDrive "AstroPhoto\Processing"
        $projectRoot = Join-Path $processingRoot "M 31\2026_Setup"
        $pixPath = Join-Path $projectRoot "Pix"
        $masterPath = Join-Path $pixPath "master\masterLight_BIN-1.xisf"
        $temporaryPath = Join-Path $pixPath "calibrated\light_c.xisf"
        $calibrationMasterPath = Join-Path $pixPath "master\masterDark.xisf"
        New-Item -ItemType Directory -Path (Split-Path $masterPath), (Split-Path $temporaryPath) -Force | Out-Null
        Set-Content -LiteralPath $masterPath -Value "light"
        Set-Content -LiteralPath $temporaryPath -Value "temporary"
        Set-Content -LiteralPath $calibrationMasterPath -Value "dark"
        @{ PixPath = $pixPath } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $projectRoot "project_meta.json")

        $plan = @(Get-AsiToPixWbppCleanupPlan -ProcessingRoot $processingRoot)

        $plan.Count | Should Be 1
        @($plan[0].FilesToRemove).Count | Should Be 2
        (@($plan[0].FilesToRemove.Name) -contains "light_c.xisf") | Should Be $true
        (@($plan[0].FilesToRemove.Name) -contains "masterDark.xisf") | Should Be $true
        @($plan[0].PreservedFiles).Count | Should Be 1
        $plan[0].PreservedFiles[0].Name | Should Be "masterLight_BIN-1.xisf"
    }

    It "uses the sibling Pix folder when metadata contains a stale path" {
        $processingRoot = Join-Path $TestDrive "renamed\AstroPhoto\Processing"
        $projectRoot = Join-Path $processingRoot "M 42\Hakos2026_Setup"
        $pixPath = Join-Path $projectRoot "Pix"
        $temporaryPath = Join-Path $pixPath "calibrated\temporary.xisf"
        $stalePixPath = Join-Path $processingRoot "M 42\2026_Setup\Pix"
        New-Item -ItemType Directory -Path (Split-Path $temporaryPath) -Force | Out-Null
        Set-Content -LiteralPath $temporaryPath -Value "temporary"
        @{ PixPath = $stalePixPath } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $projectRoot "project_meta.json")

        $plan = @(Get-AsiToPixWbppCleanupPlan -ProcessingRoot $processingRoot)

        $plan.Count | Should Be 1
        $plan[0].PixPath | Should Be $pixPath
        $plan[0].FilesToRemove[0].FullName | Should Be $temporaryPath
    }

    It "does not parse project metadata contents" {
        $processingRoot = Join-Path $TestDrive "invalid-json\AstroPhoto\Processing"
        $projectRoot = Join-Path $processingRoot "M 42\Hakos2026_Setup"
        $pixPath = Join-Path $projectRoot "Pix"
        New-Item -ItemType Directory -Path $pixPath -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $pixPath "temporary.xisf") -Value "temporary"
        Set-Content -LiteralPath (Join-Path $projectRoot "project_meta.json") -Value "not valid JSON"

        $plan = @(Get-AsiToPixWbppCleanupPlan -ProcessingRoot $processingRoot)

        $plan.Count | Should Be 1
        $plan[0].PixPath | Should Be $pixPath
    }
}

Describe "WBPP cleanup application" {
    It "removes temporary files and empty folders after project confirmation" {
        $processingRoot = Join-Path $TestDrive "apply\AstroPhoto\Processing"
        $projectRoot = Join-Path $processingRoot "M 45\2026_Setup"
        $pixPath = Join-Path $projectRoot "Pix"
        $masterPath = Join-Path $pixPath "master\masterLight_RGB.xisf"
        $temporaryPath = Join-Path $pixPath "debayered\temporary.xisf"
        New-Item -ItemType Directory -Path (Split-Path $masterPath), (Split-Path $temporaryPath) -Force | Out-Null
        Set-Content -LiteralPath $masterPath -Value "light"
        Set-Content -LiteralPath $temporaryPath -Value "temporary"
        @{ PixPath = $pixPath } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $projectRoot "project_meta.json")
        $expectedReclaimedBytes = (Get-Item -LiteralPath $temporaryPath).Length
        $plan = @(Get-AsiToPixWbppCleanupPlan -ProcessingRoot $processingRoot)

        $result = Invoke-AsiToPixWbppCleanupPlan -Plan $plan -ConfirmationReader { param($Prompt) "y" }

        $result.RemovedFileCount | Should Be 1
        $result.ReclaimedBytes | Should Be $expectedReclaimedBytes
        $result.WhatIfBytes | Should Be 0
        Test-Path -LiteralPath $temporaryPath | Should Be $false
        Test-Path -LiteralPath (Split-Path $temporaryPath) | Should Be $false
        Test-Path -LiteralPath $masterPath | Should Be $true
    }

    It "does not remove anything when the project is declined" {
        $processingRoot = Join-Path $TestDrive "decline\AstroPhoto\Processing"
        $projectRoot = Join-Path $processingRoot "M 51\2026_Setup"
        $pixPath = Join-Path $projectRoot "Pix"
        $temporaryPath = Join-Path $pixPath "registered\temporary.xisf"
        New-Item -ItemType Directory -Path (Split-Path $temporaryPath) -Force | Out-Null
        Set-Content -LiteralPath $temporaryPath -Value "temporary"
        @{ PixPath = $pixPath } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $projectRoot "project_meta.json")
        $plan = @(Get-AsiToPixWbppCleanupPlan -ProcessingRoot $processingRoot)

        $result = Invoke-AsiToPixWbppCleanupPlan -Plan $plan -ConfirmationReader { param($Prompt) "N" }

        $result.DeclinedProjectCount | Should Be 1
        $result.ReclaimedBytes | Should Be 0
        Test-Path -LiteralPath $temporaryPath | Should Be $true
    }

    It "reports the total size that WhatIf would reclaim" {
        $processingRoot = Join-Path $TestDrive "whatif-size\AstroPhoto\Processing"
        $projectRoot = Join-Path $processingRoot "M 81\2026_Setup"
        $pixPath = Join-Path $projectRoot "Pix"
        $temporaryPath = Join-Path $pixPath "calibrated\temporary.xisf"
        New-Item -ItemType Directory -Path (Split-Path $temporaryPath) -Force | Out-Null
        Set-Content -LiteralPath $temporaryPath -Value "temporary data"
        Set-Content -LiteralPath (Join-Path $projectRoot "project_meta.json") -Value "{}"
        $expectedBytes = (Get-Item -LiteralPath $temporaryPath).Length
        $plan = @(Get-AsiToPixWbppCleanupPlan -ProcessingRoot $processingRoot)

        $result = Invoke-AsiToPixWbppCleanupPlan -Plan $plan -WhatIf

        $result.WhatIfBytes | Should Be $expectedBytes
        $result.ReclaimedBytes | Should Be 0
        Test-Path -LiteralPath $temporaryPath | Should Be $true
    }
}
