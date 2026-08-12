param(
    [Parameter(Mandatory = $true)]
    [string]$SourcePath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$targetSheetName = 'X-Ray ' + (-join (0xC77C, 0xBC18, 0xC6A9 | ForEach-Object { [char]$_ }))
$objectiveType = -join (0xAC1D, 0xAD00, 0xC2DD | ForEach-Object { [char]$_ })
$subjectiveType = -join (0xC8FC, 0xAD00, 0xC2DD | ForEach-Object { [char]$_ })
$expectedAnswer = (-join (0xC791, 0xC5C5 | ForEach-Object { [char]$_ })) +
    ', ' +
    (-join (0xB9AC, 0xB354 | ForEach-Object { [char]$_ })) +
    '/' +
    (-join (0xC870, 0xC7A5 | ForEach-Object { [char]$_ }))

function Read-ZipEntryText {
    param(
        [System.IO.Compression.ZipArchive]$Archive,
        [string]$EntryName
    )

    $entry = $Archive.GetEntry($EntryName)
    if ($null -eq $entry) {
        throw "ZIP entry was not found: $EntryName"
    }

    $stream = $entry.Open()
    try {
        $reader = [System.IO.StreamReader]::new(
            $stream,
            [System.Text.UTF8Encoding]::new($false),
            $true
        )
        try {
            return $reader.ReadToEnd()
        }
        finally {
            $reader.Dispose()
        }
    }
    finally {
        $stream.Dispose()
    }
}

function Resolve-CellText {
    param(
        [System.Xml.XmlElement]$Cell,
        [System.Xml.XmlNodeList]$SharedItems,
        [System.Xml.XmlNamespaceManager]$NamespaceManager
    )

    if ($null -eq $Cell) {
        return ''
    }

    if ($Cell.GetAttribute('t') -eq 'inlineStr') {
        return $Cell.SelectSingleNode('m:is', $NamespaceManager).InnerText
    }

    $valueNode = $Cell.SelectSingleNode('m:v', $NamespaceManager)
    if ($null -eq $valueNode) {
        return ''
    }

    if ($Cell.GetAttribute('t') -eq 's') {
        return $SharedItems[[int]$valueNode.InnerText].InnerText
    }

    return $valueNode.InnerText
}

if (-not (Test-Path -LiteralPath $SourcePath -PathType Leaf)) {
    throw "Source file was not found: $SourcePath"
}

Copy-Item -LiteralPath $SourcePath -Destination $OutputPath -Force

$archive = [System.IO.Compression.ZipFile]::Open(
    $OutputPath,
    [System.IO.Compression.ZipArchiveMode]::Update
)

try {
    [xml]$workbookXml = Read-ZipEntryText $archive 'xl/workbook.xml'
    [xml]$relationsXml = Read-ZipEntryText $archive 'xl/_rels/workbook.xml.rels'
    [xml]$sharedXml = Read-ZipEntryText $archive 'xl/sharedStrings.xml'

    $workbookNs = [System.Xml.XmlNamespaceManager]::new($workbookXml.NameTable)
    $workbookNs.AddNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')

    $sheetNode = $workbookXml.SelectSingleNode(
        "//m:sheet[@name=`"$targetSheetName`"]",
        $workbookNs
    )
    if ($null -eq $sheetNode) {
        throw 'The X-Ray general sheet was not found.'
    }

    $relationshipId = $sheetNode.GetAttribute(
        'id',
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    )

    $relationsNs = [System.Xml.XmlNamespaceManager]::new($relationsXml.NameTable)
    $relationsNs.AddNamespace('p', 'http://schemas.openxmlformats.org/package/2006/relationships')
    $relationship = $relationsXml.SelectSingleNode(
        "//p:Relationship[@Id='$relationshipId']",
        $relationsNs
    )
    if ($null -eq $relationship) {
        throw "The X-Ray sheet relationship was not found: $relationshipId"
    }

    $sheetPath = ('xl/' + $relationship.Target).Replace('xl/../', '')
    $sheetEntry = $archive.GetEntry($sheetPath)
    if ($null -eq $sheetEntry) {
        throw "The X-Ray sheet XML was not found: $sheetPath"
    }

    $sheetTimestamp = $sheetEntry.LastWriteTime
    $sheetText = Read-ZipEntryText $archive $sheetPath
    [xml]$sheetXml = $sheetText

    $sheetNs = [System.Xml.XmlNamespaceManager]::new($sheetXml.NameTable)
    $sheetNs.AddNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')

    $sharedNs = [System.Xml.XmlNamespaceManager]::new($sharedXml.NameTable)
    $sharedNs.AddNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    $sharedItems = $sharedXml.SelectNodes('//m:si', $sharedNs)

    $scoreCell = $sheetXml.SelectSingleNode('//m:c[@r="C26"]', $sheetNs)
    $typeCell = $sheetXml.SelectSingleNode('//m:c[@r="E26"]', $sheetNs)
    $questionCell = $sheetXml.SelectSingleNode('//m:c[@r="F26"]', $sheetNs)
    $answerCell = $sheetXml.SelectSingleNode('//m:c[@r="O26"]', $sheetNs)
    $nearbySubjectiveCell = $sheetXml.SelectSingleNode('//m:c[@r="E28"]', $sheetNs)

    $oldScore = Resolve-CellText $scoreCell $sharedItems $sheetNs
    $oldType = Resolve-CellText $typeCell $sharedItems $sheetNs
    $question = Resolve-CellText $questionCell $sharedItems $sheetNs
    $answer = Resolve-CellText $answerCell $sharedItems $sheetNs
    $nearbySubjective = Resolve-CellText $nearbySubjectiveCell $sharedItems $sheetNs

    if ($oldScore -ne '4' -or $oldType -ne $objectiveType) {
        throw "Unexpected original values: C26=$oldScore, E26=$oldType"
    }
    if ($question -notlike '*(RFC)*') {
        throw "The target question did not match: $question"
    }
    if ($answer -ne $expectedAnswer) {
        throw "The target answer did not match: $answer"
    }
    if ($nearbySubjective -ne $subjectiveType) {
        throw "The nearby subjective reference cell E28 did not match: $nearbySubjective"
    }

    $subjectiveIndex = $nearbySubjectiveCell.SelectSingleNode('m:v', $sheetNs).InnerText
    $typeIndex = $typeCell.SelectSingleNode('m:v', $sheetNs).InnerText

    $scorePattern = '(<c\b(?=[^>]*\br="C26")[^>]*>.*?<v>)' +
        [regex]::Escape($scoreCell.SelectSingleNode('m:v', $sheetNs).InnerText) +
        '(</v>.*?</c>)'
    $typePattern = '(<c\b(?=[^>]*\br="E26")[^>]*>.*?<v>)' +
        [regex]::Escape($typeIndex) +
        '(</v>.*?</c>)'

    $scoreRegex = [regex]::new(
        $scorePattern,
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )
    $typeRegex = [regex]::new(
        $typePattern,
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )

    if ($scoreRegex.Matches($sheetText).Count -ne 1) {
        throw 'The C26 score XML location was not unique.'
    }
    if ($typeRegex.Matches($sheetText).Count -ne 1) {
        throw 'The E26 type XML location was not unique.'
    }

    $patchedText = $scoreRegex.Replace($sheetText, '${1}5${2}', 1)
    $patchedText = $typeRegex.Replace(
        $patchedText,
        ('${1}' + $subjectiveIndex + '${2}'),
        1
    )

    if ($patchedText.Length -ne $sheetText.Length) {
        throw 'The XML length changed unexpectedly.'
    }

    $differentCharacters = 0
    for ($index = 0; $index -lt $sheetText.Length; $index++) {
        if ($sheetText[$index] -ne $patchedText[$index]) {
            $differentCharacters++
        }
    }
    if ($differentCharacters -ne 2) {
        throw "Unexpected XML differences were found: $differentCharacters characters"
    }

    $sheetEntry.Delete()
    $newEntry = $archive.CreateEntry(
        $sheetPath,
        [System.IO.Compression.CompressionLevel]::Optimal
    )
    $newEntry.LastWriteTime = $sheetTimestamp
    $outputStream = $newEntry.Open()
    try {
        $writer = [System.IO.StreamWriter]::new(
            $outputStream,
            [System.Text.UTF8Encoding]::new($false)
        )
        try {
            $writer.Write($patchedText)
        }
        finally {
            $writer.Dispose()
        }
    }
    finally {
        $outputStream.Dispose()
    }

    Write-Output "PATCHED|sheet=$sheetPath|C26=5|E26=$subjectiveType|xml_differences=$differentCharacters"
}
finally {
    $archive.Dispose()
}
