# 乱数から、tebunko のインデックス（場所ごとの TSV）と同じ形の性能テスト用データを作る。
# 形は tebunko の検索の計測に使った合成インデックス（ブック 5.4 万・TSV 16 万・1GB）に合わせる。中身の文字は乱数なので同じにはならない。
# 作った TSV は、そのまま tebunko のインデックスとして検索を測れる。new_books.ps1 で Excel ブックにもできる。
#
#   .\tools\new_index.ps1 -Dest <index のフォルダ>
#   -Seed 1       … 乱数の種。同じ種・同じ -Scale からは、スレッドの数によらず同じ TSV ができる
#   -Scale 0.1    … ブックの数を減らす（0.1 で 1 割。大きいブックは常に作る）
#   -Workers 4    … 並行して作るスレッドの数
#
# <Dest>\部署<0〜51>\年度<0〜9>\資料<番号>.xlsx\シート<1〜5>.tsv と、<Dest>\大きい\乱数.xlsx\シート.tsv を作る。
# 作成済みのブックは飛ばすため、止めても続きから作れる。
param (
    [Parameter(Mandatory = $true)][string]$Dest,
    [int]$Seed = 1,
    [double]$Scale = 1.0,
    [int]$Workers = 4
)

$ErrorActionPreference = "Stop"
[void][System.IO.Directory]::CreateDirectory($Dest)
$Dest = (Resolve-Path -LiteralPath $Dest).ProviderPath.TrimEnd("\")

# 形（計測に使った合成インデックスの実測に合わせる）
$bookCount = [int][Math]::Round(53603 * $Scale)
$deptCount = 52
$yearCount = 10

# 1 冊を作る処理（スレッドごとに読み込む）
$worker = {
    param ($books, $dest, $seed, $progress)

    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    # 文字は漢字（U+4E00〜U+5057）に、カタカナ（U+30A1〜U+30F3）を 4% ほど混ぜる。
    # 1 セルずつ文字を選ぶと遅いため、乱数の文字の長い列を先に作り、そこから切り出す
    $poolRandom = New-Object Random $seed
    $chars = New-Object char[] 1000000
    for ($i = 0; $i -lt $chars.Length; $i++) {
        if ($poolRandom.Next(100) -lt 4) {
            $chars[$i] = [char](0x30A1 + $poolRandom.Next(0x30F3 - 0x30A1 + 1))
        } else {
            $chars[$i] = [char](0x4E00 + $poolRandom.Next(0x5057 - 0x4E00 + 1))
        }
    }
    $pool = New-Object string (, $chars)
    $poolMax = $pool.Length - 10
    $culture = [System.Globalization.CultureInfo]::InvariantCulture

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
                $sb = New-Object System.Text.StringBuilder ($rows * 64)
                for ($row = 0; $row -lt $rows; $row++) {
                    $len2 = if ($r.Next(1000) -lt 46) { 0 } else { 1 + $r.Next(8) }
                    $pre = $r.Next(5)
                    $post = $r.Next(5)
                    [void]$sb.Append($r.Next(100000)).Append("`t").
                        Append($pool, $r.Next($poolMax), $len2).Append("`t").
                        Append($pool, $r.Next($poolMax), $r.Next(5)).Append("`t`"").
                        Append((1000 + $r.Next(999000)).ToString("#,##0", $culture)).Append("`"`t").
                        Append($pool, $r.Next($poolMax), $pre).Append("の").
                        Append($pool, $r.Next($poolMax), $post).Append("`r`n")
                }
                [System.IO.File]::WriteAllText("$tmpDir\シート$s.tsv", $sb.ToString(), $utf8Bom)
            }
            [System.IO.Directory]::Move($tmpDir, $bookDir)
            [void]$progress.Done.Add($book.RelPath)
        } catch {
            [void]$progress.Failed.Add("$($book.RelPath): $($_.Exception.Message)")
        }
    }
}

$watch = [System.Diagnostics.Stopwatch]::StartNew()

# ブックの番号をフォルダ（部署 × 年度）にばらまく
$random = New-Object Random $Seed
$books = New-Object System.Collections.Generic.List[object]
for ($n = 0; $n -lt $bookCount; $n++) {
    $folder = "部署$($random.Next($deptCount))\年度$($random.Next($yearCount))"
    $books.Add(@{ Number = $n; RelPath = "$folder\資料$n.xlsx" })
}
Write-Host ("ブック {0:N0} 冊と大きいブック 1 冊を作ります。" -f $books.Count)

# 大きいブック（1 列 × 70,003 行。1 行 50 文字の漢字）
$bigDir = Join-Path $Dest "大きい\乱数.xlsx"
if (![System.IO.Directory]::Exists($bigDir)) {
    $r = New-Object Random ($Seed * 1000003 - 1)
    $sb = New-Object System.Text.StringBuilder (70003 * 52)
    $line = New-Object char[] 50
    for ($row = 0; $row -lt 70003; $row++) {
        for ($i = 0; $i -lt 50; $i++) { $line[$i] = [char](0x4E00 + $r.Next(0x9FA5 - 0x4E00 + 1)) }
        [void]$sb.Append($line).Append("`r`n")
    }
    [void][System.IO.Directory]::CreateDirectory("$bigDir.tmp")
    [System.IO.File]::WriteAllText("$bigDir.tmp\シート.tsv", $sb.ToString(), (New-Object System.Text.UTF8Encoding($true)))
    [System.IO.Directory]::Move("$bigDir.tmp", $bigDir)
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
    [void]$ps.AddScript($worker).AddArgument($part).AddArgument($Dest).AddArgument($Seed).AddArgument($progress)
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
