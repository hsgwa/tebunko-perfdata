# 乱数から、tebunko のインデックス（場所ごとの TSV）と同じ形の性能テスト用データを作る。
# 形は tebunko の検索の計測に使った合成インデックス（ブック 5.4 万・TSV 16 万・1GB）に合わせる。中身の文字は乱数なので同じにはならない。
# 作った TSV は、tebunko の pack の作成と検索の計測に使う。new_books.ps1 で Excel ブックにもできる。
#
#   .\tools\new_index.ps1 -Dest <index のフォルダ>
#   -Seed 1       … 乱数の種。同じ種・同じ -Scale からは、スレッドの数によらず同じ TSV ができる
#   -Scale 0.1    … ブックの数を減らす（0.1 で 1 割。大きいブックは常に作る）
#   -Workers 4    … 並行して書き出すスレッドの数（既定は論理コアの数）
#
# <Dest>\部署<0〜51>\年度<0〜9>\資料<番号>.xlsx\シート<1〜5>.tsv と、<Dest>\大きい\乱数.xlsx\シート.tsv を作る。
# 作成済みのブックは飛ばすため、止めても続きから作れる。
#
# 速さのため、行は 1 行ずつ作らない。乱数の行を先に 20 万行作っておき、各シートはその連続した範囲を切り出す。
# そのため同じ行が複数のシートに出る（読む量・照合する量は変わらないが、乱数の語のヒット件数は偏る）。
# 件数を決めて測れるよう、計測の語（words.tsv）の「見積書」は、行の文字（漢字 U+4E00〜U+5057・カタカナ）に無い文字で作り、3 冊に 1 回ずつ入れる。
param (
    [Parameter(Mandatory = $true)][string]$Dest,
    [int]$Seed = 1,
    [double]$Scale = 1.0,
    [int]$Workers = [Environment]::ProcessorCount
)

$ErrorActionPreference = "Stop"
[void][System.IO.Directory]::CreateDirectory($Dest)
$Dest = (Resolve-Path -LiteralPath $Dest).ProviderPath.TrimEnd("\")

# 形（計測に使った合成インデックスの実測に合わせる）
$bookCount = [int][Math]::Round(53603 * $Scale)
$deptCount = 52
$yearCount = 10
$poolRows = 200000
$rareWord = "見積書"
$rareBooks = 3

$watch = [System.Diagnostics.Stopwatch]::StartNew()
$random = New-Object Random $Seed
$culture = [System.Globalization.CultureInfo]::InvariantCulture

function newCharPool([Random]$r, [int]$length, [int]$first, [int]$last, [int]$kanaPercent) {
    # 乱数の文字の列。kanaPercent の割合でカタカナ（U+30A1〜U+30F3）を混ぜる
    $chars = New-Object char[] $length
    for ($i = 0; $i -lt $length; $i++) {
        if ($r.Next(100) -lt $kanaPercent) {
            $chars[$i] = [char](0x30A1 + $r.Next(0x30F3 - 0x30A1 + 1))
        } else {
            $chars[$i] = [char]($first + $r.Next($last - $first + 1))
        }
    }
    return New-Object string (, $chars)
}

# 行の列（5 列: 整数、漢字 0〜8 文字、漢字 0〜4 文字、桁区切りの数値、〜の〜）と、各行の先頭の位置
$chars = newCharPool $random 20000 0x4E00 0x5057 4
$charMax = $chars.Length - 10
$sb = New-Object System.Text.StringBuilder ($poolRows * 40)
$starts = New-Object int[] ($poolRows + 1)
for ($row = 0; $row -lt $poolRows; $row++) {
    $starts[$row] = $sb.Length
    $len2 = if ($random.Next(1000) -lt 46) { 0 } else { 1 + $random.Next(8) }
    [void]$sb.Append($random.Next(100000)).Append("`t").
        Append($chars, $random.Next($charMax), $len2).Append("`t").
        Append($chars, $random.Next($charMax), $random.Next(5)).Append("`t`"").
        Append((1000 + $random.Next(999000)).ToString("#,##0", $culture)).Append("`"`t").
        Append($chars, $random.Next($charMax), $random.Next(5)).Append("の").
        Append($chars, $random.Next($charMax), $random.Next(5)).Append("`r`n")
}
$starts[$poolRows] = $sb.Length
$rowText = $sb.ToString()
$sb = $null
Write-Host ("行の列 {0:N0} 行を作りました（{1:N1} 秒）。" -f $poolRows, $watch.Elapsed.TotalSeconds)

# ブックの番号をフォルダ（部署 × 年度）にばらまく。まれな語を入れるブックも決める
$books = New-Object System.Collections.Generic.List[object]
for ($n = 0; $n -lt $bookCount; $n++) {
    $folder = "部署$($random.Next($deptCount))\年度$($random.Next($yearCount))"
    $books.Add(@{ Number = $n; RelPath = "$folder\資料$n.xlsx"; Rare = $false })
}
$marked = 0
while ($marked -lt [Math]::Min($rareBooks, $books.Count)) {
    $book = $books[$random.Next($books.Count)]
    if (!$book.Rare) { $book.Rare = $true; $marked++ }
}
$rareLine = "0`t$rareWord`t`t`"1,000`"`tの`r`n"

# 大きいブック（1 列 × 70,003 行。1 行 50 文字の漢字）。文字の切れ端をつなぎ、50 文字ごとに改行する
$bigDir = Join-Path $Dest "大きい\乱数.xlsx"
if (![System.IO.Directory]::Exists($bigDir)) {
    $bigChars = newCharPool $random 20000 0x4E00 0x9FA5 0
    $big = New-Object System.Text.StringBuilder (70003 * 52)
    while ($big.Length -lt 70003 * 50) {
        [void]$big.Append($bigChars, $random.Next($bigChars.Length - 1000), 1000)
    }
    $bigText = [regex]::Replace($big.ToString(0, 70003 * 50), '(.{50})', "`$1`r`n")
    [void][System.IO.Directory]::CreateDirectory("$bigDir.tmp")
    [System.IO.File]::WriteAllText("$bigDir.tmp\シート.tsv", $bigText, (New-Object System.Text.UTF8Encoding($true)))
    [System.IO.Directory]::Move("$bigDir.tmp", $bigDir)
}
Write-Host ("ブック {0:N0} 冊を書き出します。" -f $books.Count)

# 1 冊を書き出す処理（スレッドごとに読み込む）。行の列から切り出して書くだけにし、PowerShell の処理を少なくする
$worker = {
    param ($books, $dest, $seed, $rowText, $starts, $poolRows, $rareLine, $progress)

    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    foreach ($book in $books) {
        $bookDir = Join-Path $dest $book.RelPath
        if ([System.IO.Directory]::Exists($bookDir)) {
            [void]$progress.Skipped.Add($book.RelPath)
            continue
        }
        try {
            # ブックの番号ごとの乱数（スレッドの数・作る順によらず同じ中身にする）
            $r = New-Object Random ($seed * 1000003 + $book.Number)
            $tmpDir = "$bookDir.tmp"
            if ([System.IO.Directory]::Exists($tmpDir)) { [System.IO.Directory]::Delete($tmpDir, $true) }
            [void][System.IO.Directory]::CreateDirectory($tmpDir)
            $sheetCount = 1 + $r.Next(5)
            for ($s = 1; $s -le $sheetCount; $s++) {
                $u = $r.NextDouble()
                $rows = 20 + [int][Math]::Floor(300 * $u * $u)
                $first = $r.Next($poolRows - $rows)
                $text = $rowText.Substring($starts[$first], $starts[$first + $rows] - $starts[$first])
                if ($book.Rare -and $s -eq 1) { $text = $rareLine + $text }
                [System.IO.File]::WriteAllText("$tmpDir\シート$s.tsv", $text, $utf8Bom)
            }
            [System.IO.Directory]::Move($tmpDir, $bookDir)
            [void]$progress.Done.Add($book.RelPath)
        } catch {
            [void]$progress.Failed.Add("$($book.RelPath): $($_.Exception.Message)")
        }
    }
}

$progress = [hashtable]::Synchronized(@{
    Done    = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
    Skipped = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
    Failed  = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
})
$pool = [runspacefactory]::CreateRunspacePool(1, $Workers)
$pool.Open()
$jobs = foreach ($k in 0..($Workers - 1)) {
    $part = @(for ($i = $k; $i -lt $books.Count; $i += $Workers) { $books[$i] })
    $ps = [powershell]::Create()
    $ps.RunspacePool = $pool
    [void]$ps.AddScript($worker).AddArgument($part).AddArgument($Dest).AddArgument($Seed).AddArgument($rowText).AddArgument($starts).AddArgument($poolRows).AddArgument($rareLine).AddArgument($progress)
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
