$ErrorActionPreference = 'Stop'
$devRoot = Split-Path $PSScriptRoot -Parent
$root = Split-Path $devRoot -Parent
$assemblyPath = (Resolve-Path (Join-Path $devRoot 'OJT_Exam_Maker_patch.exe')).Path
$asm = [Reflection.Assembly]::LoadFrom($assemblyPath)
$readerType = $asm.GetType('OjtExamPatch.Reader', $true)
$reader = [Activator]::CreateInstance($readerType, $true)
$workbookPath = (Get-ChildItem -LiteralPath $root -Filter '*.xlsm' | Where-Object { -not $_.Name.StartsWith('~$') } | Select-Object -First 1).FullName
$banks = $readerType.GetMethod('Load').Invoke($reader, [object[]]@($workbookPath))

$hanging = $null
$multiImage = $null
foreach ($bank in $banks) {
    foreach ($q in $bank.GetType().GetField('Questions').GetValue($bank)) {
        $qt = $q.GetType()
        $textValue = [string]$qt.GetField('Text').GetValue($q)
        $imageCount = $qt.GetField('Images').GetValue($q).Count
        if ($null -eq $hanging -and $textValue -match 'V-PRESS') { $hanging = $q }
        if ($null -eq $multiImage -and $imageCount -eq 4 -and $textValue -match '03BGASEMBLDN') { $multiImage = $q }
    }
}
if ($null -eq $hanging -or $null -eq $multiImage) { throw 'Verification samples were not found.' }

$previewType = $asm.GetType('OjtExamPatch.Preview', $true)
$flags = [Reflection.BindingFlags]'Static,NonPublic'
$formatMethod = $previewType.GetMethod('FormatQuestionText', $flags)
$scoreMethod = $previewType.GetMethod('ScoreText', $flags)
$wrapMethod = $previewType.GetMethods($flags) | Where-Object { $_.Name -eq 'WrapPrintLines' -and $_.GetParameters().Count -eq 2 } | Select-Object -First 1
$imageLinesMethod = $previewType.GetMethod('ImageChoiceLines', $flags)
$imageListMethod = $previewType.GetMethod('ImageChoiceLineList', $flags)

$body = [string]$formatMethod.Invoke($null, [object[]]@($hanging))
$wrapped = $wrapMethod.Invoke($null, [object[]]@([string]('1.  ' + $body), [int]94))
$hasFiveSpaceContinuation = $false
foreach ($line in $wrapped) {
    Write-Output ('WRAP=[' + ($line -replace ' ', '·') + ']')
    if ($line.StartsWith('     ')) { $hasFiveSpaceContinuation = $true }
}
$score = [string]$scoreMethod.Invoke($null, [object[]]@($hanging))
$pointChar = [char]0xC810
if (-not $hasFiveSpaceContinuation) { throw 'Five-space continuation indent failed.' }
if ($body -match ('\(\s*\d+(?:\.\d+)?\s*' + $pointChar + '\s*\)')) { throw 'Score remained in question body.' }
if ($score -ne ('(5' + $pointChar + ')')) { throw ('Score label failed: ' + $score) }

$multiText = [string]$formatMethod.Invoke($null, [object[]]@($multiImage))
$imageLineText = [string]$imageLinesMethod.Invoke($null, [object[]]@($multiText))
$parsedChoices = $imageListMethod.Invoke($null, [object[]]@($imageLineText, [int]4))
if ($parsedChoices.Count -ne 4) { throw ('Image choice parsing failed: ' + $parsedChoices.Count) }

Write-Output ('ASSERT_INDENT=' + $hasFiveSpaceContinuation)
Write-Output ('ASSERT_SCORE=' + $score)
Write-Output ('ASSERT_IMAGE_CHOICES=' + $parsedChoices.Count)
