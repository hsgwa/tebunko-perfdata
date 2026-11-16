# サイズ上限の確認用の .docx を作る。性能テスト・上限の確認用。
# Word・Office は使わず、docx の中の XML を直接書く。中身は [Content_Types].xml・_rels/.rels・word/document.xml だけ（docProps は入れない）。
#
#   .\tools\new_docx.ps1 -Dest <置くフォルダ> -DocumentMB <word/document.xml の展開後の大きさ>
#   -Count 2         … 作る数（1〜16。既定 1）
#   -DeclaredMB 1    … word/document.xml の展開後の大きさを、中央ディレクトリとローカルヘッダーの両方でこの値に偽る
#                      （0 より大きく 4096 未満。上限の判定をすり抜けるかを試す）
#   -Name doc        … ファイル名は <Name>001.docx から番号を振る
#
# 1MB = 1,048,576 バイト。document.xml は大きな文字列を作らず、ストリームに少しずつ書く。
# 同じ指定からは、同じ .NET の上では同じバイト列ができる（zip の中の日時と圧縮の水準を固定）。
# すでにあるファイルは作り直さず飛ばす。<名前>.tmp に書き、偽り終えてから本当の名前に移す。
param (
    [Parameter(Mandatory = $true)][string]$Dest,
    [Parameter(Mandatory = $true)][double]$DocumentMB,
    [int]$Count = 1,
    [double]$DeclaredMB = 0,
    [string]$Name = "doc"
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression

# 範囲の外の指定は、作る前に断る
if ($DocumentMB -lt 0.1 -or $DocumentMB -gt 1000) { throw "-DocumentMB は 0.1〜1000 で指定してください: $DocumentMB" }
if ($Count -lt 1 -or $Count -gt 16) { throw "-Count は 1〜16 で指定してください: $Count" }
if ($PSBoundParameters.ContainsKey("DeclaredMB") -and ($DeclaredMB -le 0 -or $DeclaredMB -ge 4096)) {
    throw "-DeclaredMB は 0 より大きく 4096 未満で指定してください: $DeclaredMB"
}
if ($Name -notmatch '^[^\\/:*?"<>|\x00-\x1F]+$') { throw "-Name にファイル名に使えない文字があります: $Name" }

$MB = 1048576
$utf8 = New-Object System.Text.UTF8Encoding($false)
# zip の中のファイルの日時をそろえる（new_books.ps1 と同じ）
$entryTime = [DateTimeOffset]::new(2026, 1, 1, 0, 0, 0, [TimeSpan]::Zero)
$xmlHead = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`r`n"
$wNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$relType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

$contentTypes = $xmlHead + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '</Types>'
$rootRels = $xmlHead + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    "<Relationship Id=`"rId1`" Type=`"$relType/officeDocument`" Target=`"word/document.xml`"/></Relationships>"

# document.xml の頭・段落・末尾（本文は架空の文の繰り返し）
$docHead = $utf8.GetBytes($xmlHead + "<w:document xmlns:w=`"$wNs`"><w:body>")
$docTail = $utf8.GetBytes('</w:body></w:document>')
$paragraph = $utf8.GetBytes('<w:p><w:r><w:t>これは計測用の架空の文書の本文です。</w:t></w:r></w:p>')

function addEntry($zip, [string]$entryName, [byte[]]$bytes) {
    $entry = $zip.CreateEntry($entryName, [System.IO.Compression.CompressionLevel]::Optimal)
    $entry.LastWriteTime = $entryTime
    $stream = $entry.Open()
    try { $stream.Write($bytes, 0, $bytes.Length) } finally { $stream.Dispose() }
}

# word/document.xml を、目標の大きさになるまで少しずつ書く。書いたバイト数を返す
function addDocument($zip, [long]$targetBytes) {
    $entry = $zip.CreateEntry("word/document.xml", [System.IO.Compression.CompressionLevel]::Optimal)
    $entry.LastWriteTime = $entryTime
    $stream = $entry.Open()
    $written = [long]0
    try {
        $stream.Write($docHead, 0, $docHead.Length)
        $written += $docHead.Length
        $paragraphs = [long][Math]::Max(1, [Math]::Floor(($targetBytes - $docHead.Length - $docTail.Length) / $paragraph.Length))
        $perChunk = 4096
        $chunk = New-Object byte[] ($paragraph.Length * $perChunk)
        for ($i = 0; $i -lt $perChunk; $i++) { [Array]::Copy($paragraph, 0, $chunk, $i * $paragraph.Length, $paragraph.Length) }
        while ($paragraphs -gt 0) {
            $n = [int][Math]::Min($perChunk, $paragraphs)
            $stream.Write($chunk, 0, $paragraph.Length * $n)
            $written += [long]$paragraph.Length * $n
            $paragraphs -= $n
        }
        $stream.Write($docTail, 0, $docTail.Length)
        $written += $docTail.Length
    } finally {
        $stream.Dispose()
    }
    return $written
}

# 書き終えた zip の word/document.xml の「展開後の大きさ」を、中央ディレクトリとローカルヘッダーの両方で書き換える。
# 書き換える前に、両方の今の値が実際に書いたバイト数と同じことを確かめ、違えば止める。
function forgeSize([string]$path, [long]$actual, [uint32]$declared) {
    $fs = [System.IO.File]::Open($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite)
    try {
        $len = $fs.Length
        $tailLen = [int][Math]::Min($len, 65557)
        $buf = New-Object byte[] $tailLen
        [void]$fs.Seek($len - $tailLen, [System.IO.SeekOrigin]::Begin)
        [void]$fs.Read($buf, 0, $tailLen)
        $eocd = -1
        for ($i = $tailLen - 22; $i -ge 0; $i--) {
            if ($buf[$i] -eq 0x50 -and $buf[$i + 1] -eq 0x4B -and $buf[$i + 2] -eq 0x05 -and $buf[$i + 3] -eq 0x06) { $eocd = $i; break }
        }
        if ($eocd -lt 0) { throw "中央ディレクトリの終わりが見つかりません: $path" }
        $entries = [BitConverter]::ToUInt16($buf, $eocd + 10)
        $cdOffset = [BitConverter]::ToUInt32($buf, $eocd + 16)
        if ($entries -eq 0xFFFF -or $cdOffset -eq 0xFFFFFFFF) { throw "ZIP64 の zip は偽れません: $path" }
        $cdSize = [BitConverter]::ToUInt32($buf, $eocd + 12)
        $cd = New-Object byte[] $cdSize
        [void]$fs.Seek($cdOffset, [System.IO.SeekOrigin]::Begin)
        [void]$fs.Read($cd, 0, $cdSize)
        $pos = 0
        $cdField = -1
        $localPos = -1
        for ($e = 0; $e -lt $entries; $e++) {
            if ([BitConverter]::ToUInt32($cd, $pos) -ne 0x02014B50) { throw "中央ディレクトリの項目が壊れています: $path" }
            $nameLen = [BitConverter]::ToUInt16($cd, $pos + 28)
            $extraLen = [BitConverter]::ToUInt16($cd, $pos + 30)
            $commentLen = [BitConverter]::ToUInt16($cd, $pos + 32)
            $entryName = [System.Text.Encoding]::UTF8.GetString($cd, $pos + 46, $nameLen)
            if ($entryName -eq "word/document.xml") {
                $cdField = $pos + 24
                $localPos = [BitConverter]::ToUInt32($cd, $pos + 42)
                $flags = [BitConverter]::ToUInt16($cd, $pos + 8)
                if (($flags -band 0x8) -ne 0) { throw "データ記述子のある zip は偽れません: $path" }
                if ($extraLen -ne 0) { throw "追加の項目（ZIP64 など）のある zip は偽れません: $path" }
                break
            }
            $pos += 46 + $nameLen + $extraLen + $commentLen
        }
        if ($cdField -lt 0) { throw "word/document.xml が見つかりません: $path" }
        $cdValue = [BitConverter]::ToUInt32($cd, $cdField)
        if ($cdValue -eq 0xFFFFFFFF -or $cdValue -ne $actual) { throw "中央ディレクトリの大きさ ($cdValue) が書いたバイト数 ($actual) と違います: $path" }

        $lh = New-Object byte[] 30
        [void]$fs.Seek($localPos, [System.IO.SeekOrigin]::Begin)
        [void]$fs.Read($lh, 0, 30)
        if ([BitConverter]::ToUInt32($lh, 0) -ne 0x04034B50) { throw "ローカルヘッダーの位置が正しくありません: $path" }
        if (([BitConverter]::ToUInt16($lh, 6) -band 0x8) -ne 0) { throw "ローカルヘッダーにデータ記述子のビットが立っています: $path" }
        $lhValue = [BitConverter]::ToUInt32($lh, 22)
        if ($lhValue -eq 0xFFFFFFFF -or $lhValue -ne $actual) { throw "ローカルヘッダーの大きさ ($lhValue) が書いたバイト数 ($actual) と違います: $path" }

        $bytes = [BitConverter]::GetBytes($declared)
        [void]$fs.Seek($cdOffset + $cdField, [System.IO.SeekOrigin]::Begin)
        $fs.Write($bytes, 0, 4)
        [void]$fs.Seek($localPos + 22, [System.IO.SeekOrigin]::Begin)
        $fs.Write($bytes, 0, 4)
    } finally {
        $fs.Dispose()
    }
}

[void][System.IO.Directory]::CreateDirectory($Dest)
$Dest = (Resolve-Path -LiteralPath $Dest).ProviderPath.TrimEnd("\")
$targetBytes = [long][Math]::Round($DocumentMB * $MB)
$declared = $null
if ($PSBoundParameters.ContainsKey("DeclaredMB")) { $declared = [uint32][Math]::Round($DeclaredMB * $MB) }

$watch = [System.Diagnostics.Stopwatch]::StartNew()
$made = 0
$skipped = 0
for ($i = 1; $i -le $Count; $i++) {
    $target = Join-Path $Dest ("{0}{1:000}.docx" -f $Name, $i)
    if ([System.IO.File]::Exists($target)) {
        $skipped++
        continue
    }
    $tmp = "$target.tmp"
    try {
        $file = [System.IO.File]::Open($tmp, [System.IO.FileMode]::Create)
        try {
            $zip = New-Object System.IO.Compression.ZipArchive($file, [System.IO.Compression.ZipArchiveMode]::Create)
            try {
                addEntry $zip "[Content_Types].xml" $utf8.GetBytes($contentTypes)
                addEntry $zip "_rels/.rels" $utf8.GetBytes($rootRels)
                $actual = addDocument $zip $targetBytes
            } finally {
                $zip.Dispose()
            }
        } finally {
            $file.Dispose()
        }
        if ($null -ne $declared) { forgeSize $tmp $actual $declared }
        [System.IO.File]::Move($tmp, $target)
        $made++
    } finally {
        # 途中で止まったとき（Ctrl+C・エラー）は、書きかけを残さない
        if ([System.IO.File]::Exists($tmp)) { [System.IO.File]::Delete($tmp) }
    }
}
Write-Host ("作成 {0:N0} 個・作成済みで飛ばした {1:N0} 個（{2:N1} 秒）" -f $made, $skipped, $watch.Elapsed.TotalSeconds)
