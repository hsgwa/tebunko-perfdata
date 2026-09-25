# tebunko のインデックス（場所ごとの TSV）から、同じ中身の Excel ブック（.xlsx）を作る。性能テスト用。
# Excel の COM は使わず、xlsx の中の XML を直接書く（5 万冊を COM で作ると十数時間かかるため）。
# 作ったブックを tebunko で取り込むと、元の TSV と同じ TSV ができる（数値・引用符の扱いは下の convertTsvToRows）。
#
#   .\tools\new_books.ps1 -Source <index のフォルダ> -Dest <ブックを置くフォルダ>
#   -Limit 20    … 先頭から N 冊だけ作る（確かめるとき）
#   -Workers 4   … 並行して作るスレッドの数
#
# <Source>\部署0\年度0\資料1.xlsx\シート1.tsv → <Dest>\部署0\年度0\資料1.xlsx の「シート1」。
# 変換先に同じブックがあれば飛ばすため、止めても続きから作れる。元の TSV は読むだけ。
param (
    [Parameter(Mandatory = $true)][string]$Source,
    [Parameter(Mandatory = $true)][string]$Dest,
    [int]$Limit = 0,
    [int]$Workers = 4
)

$ErrorActionPreference = "Stop"
$Source = (Resolve-Path -LiteralPath $Source).ProviderPath.TrimEnd("\")
[void][System.IO.Directory]::CreateDirectory($Dest)
$Dest = (Resolve-Path -LiteralPath $Dest).ProviderPath.TrimEnd("\")

# 1 冊を作る処理（スレッドごとに読み込む）
$worker = {
    param ($books, $source, $dest, $progress)

    Add-Type -AssemblyName System.IO.Compression
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    # zip の中のファイルの日時をそろえ、同じ TSV からは同じ xlsx ができるようにする
    $entryTime = [DateTimeOffset]::new(2026, 1, 1, 0, 0, 0, [TimeSpan]::Zero)
    $xmlHead = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`r`n"
    $ns = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    $nsRel = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    $pkgRel = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"'
    $relType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    # 表示形式 1 = 桁区切り（#,##0。組み込みの番号 3）
    $styles = $xmlHead + "<styleSheet $ns><fonts count=`"1`"><font><sz val=`"11`"/><name val=`"Yu Gothic`"/></font></fonts>" +
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>' +
        '</styleSheet>'
    $rootRels = $xmlHead + "<Relationships $pkgRel><Relationship Id=`"rId1`" Type=`"$relType/officeDocument`" Target=`"xl/workbook.xml`"/></Relationships>"

    # 数値（先頭の 0 や 16 桁以上は文字のまま。Excel の表示が変わるため）
    $reInt = [regex]::new('<F>(-?(?:0|[1-9][0-9]{0,14}))</F>', 'Compiled')
    # Excel が桁区切りの数値を引用符で囲んだもの（"346,965"）
    $reComma = [regex]::new('<F>"(-?[1-9][0-9]{0,2}(?:,[0-9]{3})+)"</F>', 'Compiled')
    $reCommaStrip = [regex]::new('(<N>[^<,]*),', 'Compiled')
    $reNum = [regex]::new('<N>([^<]*)</N>', 'Compiled')
    # Excel が引用符で囲んだ文字（セル内改行・" を含む）。中の "" は " に戻す
    $reQuoted = [regex]::new('<F>"((?:[^"<]|"")*)"</F>', 'Compiled')
    $reText = [regex]::new('<F>(.*?)</F>', 'Compiled')
    $reBadChar = [regex]::new('[\x00-\x08\x0B\x0C\x0E-\x1F]', 'Compiled')

    function convertTsvToRows([string]$text) {
        # TSV の全文を <row>…</row> の並びにする。1 セルずつ PowerShell で回さず、全文の置き換えだけで行う
        $text = $text.Replace("`r`n", "`n").TrimEnd("`n")
        $text = $reBadChar.Replace($text, "")
        $text = $text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")
        # セル内改行（インデックスでは U+2028）は、セルの中の改行に戻す
        $text = $text.Replace([string][char]0x2028, "&#10;")
        $text = "<row><F>" + $text.Replace("`n", "</F></row>`n<row><F>").Replace("`t", "</F><F>") + "</F></row>"
        $text = $text.Replace("<F></F>", "<c/>")
        $text = $reComma.Replace($text, '<N>$1</N>')
        for ($i = 0; $i -lt 6 -and $text.Contains(","); $i++) {
            $next = $reCommaStrip.Replace($text, '$1')
            if ($next.Length -eq $text.Length) { break }
            $text = $next
        }
        $text = $reNum.Replace($text, '<c s="1"><v>$1</v></c>')
        $text = $reInt.Replace($text, '<c><v>$1</v></c>')
        $text = $reQuoted.Replace($text, '<Q>$1</Q>')
        $text = $text.Replace('""', '"')
        $text = $text.Replace("<Q>", "<F>").Replace("</Q>", "</F>")
        return $reText.Replace($text, '<c t="inlineStr"><is><t xml:space="preserve">$1</t></is></c>')
    }

    function decodePlace([string]$name) {
        # TSV のファイル名は、シート名の使えない文字を %XX にしている
        return [regex]::Replace($name, '%([0-9A-Fa-f]{2})', { param($m) [string][char][Convert]::ToInt32($m.Groups[1].Value, 16) })
    }

    function addEntry($zip, [string]$name, [string]$text) {
        $entry = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
        $entry.LastWriteTime = $entryTime
        $stream = $entry.Open()
        try {
            $bytes = $utf8.GetBytes($text)
            $stream.Write($bytes, 0, $bytes.Length)
        } finally {
            $stream.Dispose()
        }
    }

    foreach ($book in $books) {
        $rel = $book.Substring($source.Length + 1)
        $target = Join-Path $dest $rel
        if ([System.IO.File]::Exists($target)) {
            [void]$progress.Skipped.Add($rel)
            continue
        }
        # シートは TSV の名前の数字の順（シート2 → シート10）
        $sheets = @([System.IO.Directory]::GetFiles($book, "*.tsv") | Sort-Object {
            $n = [System.IO.Path]::GetFileNameWithoutExtension($_)
            [regex]::Replace($n, '\d+', { param($m) $m.Value.PadLeft(10, "0") })
        })
        try {
            [void][System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($target))
            $tmp = "$target.tmp"
            $file = [System.IO.File]::Open($tmp, [System.IO.FileMode]::Create)
            $zip = New-Object System.IO.Compression.ZipArchive($file, [System.IO.Compression.ZipArchiveMode]::Create)
            try {
                $types = New-Object System.Text.StringBuilder
                $sheetList = New-Object System.Text.StringBuilder
                $wbRels = New-Object System.Text.StringBuilder
                for ($i = 1; $i -le $sheets.Count; $i++) {
                    $rows = convertTsvToRows ([System.IO.File]::ReadAllText($sheets[$i - 1], [System.Text.Encoding]::UTF8))
                    addEntry $zip "xl/worksheets/sheet$i.xml" ($xmlHead + "<worksheet $ns><sheetData>" + $rows + "</sheetData></worksheet>")
                    $name = [System.Security.SecurityElement]::Escape((decodePlace ([System.IO.Path]::GetFileNameWithoutExtension($sheets[$i - 1]))))
                    [void]$sheetList.Append("<sheet name=`"$name`" sheetId=`"$i`" r:id=`"rId$i`"/>")
                    [void]$wbRels.Append("<Relationship Id=`"rId$i`" Type=`"$relType/worksheet`" Target=`"worksheets/sheet$i.xml`"/>")
                    [void]$types.Append("<Override PartName=`"/xl/worksheets/sheet$i.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>")
                }
                $n = $sheets.Count + 1
                [void]$wbRels.Append("<Relationship Id=`"rId$n`" Type=`"$relType/styles`" Target=`"styles.xml`"/>")
                addEntry $zip "[Content_Types].xml" ($xmlHead + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
                    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
                    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
                    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
                    $types.ToString() + '</Types>')
                addEntry $zip "_rels/.rels" $rootRels
                addEntry $zip "xl/workbook.xml" ($xmlHead + "<workbook $ns $nsRel><sheets>" + $sheetList.ToString() + "</sheets></workbook>")
                addEntry $zip "xl/_rels/workbook.xml.rels" ($xmlHead + "<Relationships $pkgRel>" + $wbRels.ToString() + "</Relationships>")
                addEntry $zip "xl/styles.xml" $styles
            } finally {
                $zip.Dispose()
                $file.Dispose()
            }
            [System.IO.File]::Move($tmp, $target)
            [void]$progress.Done.Add($rel)
        } catch {
            [void]$progress.Failed.Add("${rel}: $($_.Exception.Message)")
        }
    }
}

# 元のブックのフォルダ（<ファイル名.xlsx>、中に TSV がある）を集める
$watch = [System.Diagnostics.Stopwatch]::StartNew()
$books = @([System.IO.Directory]::GetDirectories($Source, "*.xlsx", "AllDirectories") | Sort-Object)
if ($Limit -gt 0) {
    $books = @($books | Select-Object -First $Limit)
}
Write-Host ("ブック {0:N0} 冊。列挙 {1:N1} 秒" -f $books.Count, $watch.Elapsed.TotalSeconds)

$progress = [hashtable]::Synchronized(@{
    Done    = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
    Skipped = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
    Failed  = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
})
$pool = [runspacefactory]::CreateRunspacePool(1, $Workers)
$pool.Open()
$jobs = foreach ($k in 0..($Workers - 1)) {
    # 冊をスレッドに交互に配る（フォルダごとの偏りを均す）
    $part = @(for ($i = $k; $i -lt $books.Count; $i += $Workers) { $books[$i] })
    $ps = [powershell]::Create()
    $ps.RunspacePool = $pool
    [void]$ps.AddScript($worker).AddArgument($part).AddArgument($Source).AddArgument($Dest).AddArgument($progress)
    @{ PowerShell = $ps; Handle = $ps.BeginInvoke() }
}
while (@($jobs | Where-Object { !$_.Handle.IsCompleted }).Count -gt 0) {
    Start-Sleep -Seconds 5
    $done = $progress.Done.Count + $progress.Skipped.Count + $progress.Failed.Count
    Write-Host ("  {0:N0} / {1:N0} 冊（{2:N0} 秒）" -f $done, $books.Count, $watch.Elapsed.TotalSeconds)
}
foreach ($job in $jobs) {
    $job.PowerShell.EndInvoke($job.Handle)
    foreach ($e in $job.PowerShell.Streams.Error) { Write-Warning $e }
    $job.PowerShell.Dispose()
}
$pool.Close()

Write-Host ("作成 {0:N0} 冊・作成済みで飛ばした {1:N0} 冊・失敗 {2:N0} 冊（{3:N1} 秒）" -f $progress.Done.Count, $progress.Skipped.Count, $progress.Failed.Count, $watch.Elapsed.TotalSeconds)
foreach ($f in $progress.Failed) { Write-Warning $f }
if ($progress.Failed.Count -gt 0) { exit 1 }
