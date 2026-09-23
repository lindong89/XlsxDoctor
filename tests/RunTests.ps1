<#
    XlsxDoctor 回归测试
    构造一个塞满陷阱的最小 xlsx，逐项验证清理器不会误删。

    用法：
        .\tests\RunTests.ps1
        .\tests\RunTests.ps1 -Exe ..\dist\XlsxDoctor-cli.exe

    覆盖：
        1. 默认模式：该留的留、该删的删、公式一字不改
        2. 保守模式：只删明确失效的名称
        3. 幂等性：清过的文件再清一次 = 无可清理
        4. 防呆：输出=源文件 / .xlsb / 非 zip 文件 都必须拒绝
        5. --dry-run：不写出任何文件
        6. --in-place：原文件被替换，.bak 与原文件逐字节相同
        7. 输出文件确实存在（防止"报告成功却删掉产物"）
        8. 清理后的包结构校验通过
#>
[CmdletBinding()]
param(
    [string]$Exe
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Definition)
if (-not $Exe) { $Exe = Join-Path $root 'dist\XlsxDoctor-cli.exe' }
if (-not (Test-Path -LiteralPath $Exe)) { throw "找不到可执行文件: $Exe`n请先运行 .\build.ps1" }
$Exe = (Resolve-Path -LiteralPath $Exe).Path

$work = Join-Path $env:TEMP ('xlsxdoctor-tests-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $work | Out-Null

$script:pass = 0
$script:fail = 0
$script:failedNames = @()

function Check([string]$name, [bool]$ok, [string]$detail) {
    if ($ok) {
        $script:pass++
        Write-Host ("  [PASS] {0}" -f $name) -ForegroundColor Green
    } else {
        $script:fail++
        $script:failedNames += $name
        Write-Host ("  [FAIL] {0}" -f $name) -ForegroundColor Red
        if ($detail) { Write-Host ("         {0}" -f $detail) -ForegroundColor DarkRed }
    }
}

function Section([string]$title) {
    Write-Host ''
    Write-Host ("=== {0} ===" -f $title) -ForegroundColor Cyan
}

function Run([string[]]$argv) {
    # 被测程序会往 stderr 写错误信息，这里必须让 EAP 退到 Continue，
    # 否则 PowerShell 会把原生命令的 stderr 当成终止错误直接中断测试。
    $prevEap = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $out = & $Exe @argv 2>&1 | Out-String
        $code = $LASTEXITCODE
    } finally { $ErrorActionPreference = $prevEap }
    return [pscustomobject]@{ Out = $out; Code = $code }
}

# ============================================================ 夹具

$formulaList = @(
    'MyRate*2', '[1]Sheet1!$A$1', 'SUM(A1:B1)', 'AN+1',
    'BrokenName', '"UnusedGood"', 'Data!A1', 'Data!$A$1+MyRate', 'DerivedRate', '税率_日本*2'
)

function New-Fixture([string]$path) {
    $rows = ''
    for ($i = 0; $i -lt $formulaList.Count; $i++) {
        $r = $i + 1
        $rows += ('<row r="{0}"><c r="A{0}"><f>{1}</f><v>1</v></c></row>' -f $r, $formulaList[$i])
    }

    $parts = [ordered]@{
        '[Content_Types].xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/><Override PartName="/xl/externalLinks/externalLink2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/></Types>'

        '_rels/.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'

        'xl/workbook.xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="MyRate">Sheet1!$A$1</definedName><definedName name="AN">Sheet1!$A$2</definedName><definedName name="UnusedGood">Sheet1!$A$3</definedName><definedName name="SUM">Sheet1!$A$4</definedName><definedName name="Data">Sheet1!$A$5</definedName><definedName name="BrokenName">#REF!</definedName><definedName name="ExtUsed">[1]Sheet1!$A$1</definedName><definedName name="ExtUnused">[2]Sheet1!$A$1</definedName><definedName name="BaseRate">Sheet1!$A$6</definedName><definedName name="DerivedRate">BaseRate*2</definedName><definedName name="税率_日本">Sheet1!$A$7</definedName><definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$B$10</definedName><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Sheet1!$A$1:$B$10</definedName></definedNames><externalReferences><externalReference r:id="rId2"/><externalReference r:id="rId3"/></externalReferences></workbook>'

        'xl/_rels/workbook.xml.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink2.xml"/></Relationships>'

        'xl/worksheets/sheet1.xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' + $rows + '</sheetData></worksheet>'

        'xl/externalLinks/externalLink1.xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><externalBook r:id="rId1"><sheetNames><sheetName val="Sheet1"/></sheetNames></externalBook></externalLink>'
        'xl/externalLinks/_rels/externalLink1.xml.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="file:///C:/temp/book1.xlsx" TargetMode="External"/></Relationships>'

        'xl/externalLinks/externalLink2.xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><externalBook r:id="rId1"><sheetNames><sheetName val="Sheet1"/></sheetNames></externalBook></externalLink>'
        'xl/externalLinks/_rels/externalLink2.xml.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="file:///C:/temp/book2.xlsx" TargetMode="External"/></Relationships>'
    }

    if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    $fs = [System.IO.File]::Create($path)
    try {
        $zip = [System.IO.Compression.ZipArchive]::new($fs, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            $utf8 = [System.Text.UTF8Encoding]::new($false)
            foreach ($k in $parts.Keys) {
                $e = $zip.CreateEntry($k)
                $s = $e.Open()
                try {
                    $sw = [System.IO.StreamWriter]::new($s, $utf8)
                    try { $sw.Write($parts[$k]) } finally { $sw.Dispose() }
                } finally { $s.Dispose() }
            }
        } finally { $zip.Dispose() }
    } finally { $fs.Dispose() }
}

function Get-ZipText([string]$zipPath, [string]$entryName) {
    $z = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
    try {
        $e = $z.Entries | Where-Object { $_.FullName -eq $entryName }
        if (-not $e) { return $null }
        $s = $e.Open()
        try {
            $sr = [System.IO.StreamReader]::new($s)
            try { return $sr.ReadToEnd() } finally { $sr.Dispose() }
        } finally { $s.Dispose() }
    } finally { $z.Dispose() }
}

function Get-ZipNames([string]$zipPath) {
    $z = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
    try { return @($z.Entries | ForEach-Object { $_.FullName }) } finally { $z.Dispose() }
}

function Get-Sha([string]$path) {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $fs = [System.IO.File]::OpenRead($path)
        try { return [System.BitConverter]::ToString($sha.ComputeHash($fs)) } finally { $fs.Dispose() }
    } finally { $sha.Dispose() }
}

function HasName([string]$wb, [string]$name) {
    return ($wb -match ('name="' + [regex]::Escape($name) + '"'))
}

function Get-Formulas([string]$sheetXml) {
    $ms = [regex]::Matches($sheetXml, '<f>([^<]*)</f>')
    return @($ms | ForEach-Object { $_.Groups[1].Value })
}

# ============================================================ 开始

Write-Host ''
Write-Host 'XlsxDoctor 回归测试' -ForegroundColor White
Write-Host ("可执行文件: {0}" -f $Exe)
Write-Host ("工作目录  : {0}" -f $work)

$fixture    = Join-Path $work 'fixture.xlsx'
$defaultOut = Join-Path $work 'out-default.xlsx'
$consOut    = Join-Path $work 'out-conservative.xlsx'
$dryOut     = Join-Path $work 'out-dryrun.xlsx'
$idemOut    = Join-Path $work 'out-idempotent.xlsx'
$inplace    = Join-Path $work 'inplace.xlsx'

New-Fixture $fixture
$fixtureSha = Get-Sha $fixture
$sheetBefore = Get-ZipText $fixture 'xl/worksheets/sheet1.xml'

# ---------------------------------------------------------------- 1. 默认模式
Section '1. 默认模式'

$r = Run @($fixture, '-o', $defaultOut)
Check '退出码为 0' ($r.Code -eq 0) ("实际退出码 " + $r.Code)
Check '输出文件存在' (Test-Path -LiteralPath $defaultOut) '输出文件被删掉了'
Check '包校验通过' ($r.Out -match '包校验\s*:\s*通过') $r.Out

$wb = Get-ZipText $defaultOut 'xl/workbook.xml'
Check '保留被公式引用的 MyRate'   (HasName $wb 'MyRate')   'MyRate 被误删'
Check '保留被公式引用的短名 AN'   (HasName $wb 'AN')       'AN 被误删（可能被当成列号）'
Check '保留被间接引用的 BaseRate' (HasName $wb 'BaseRate') '定义名称依赖链被截断'
Check '保留被公式引用的 Unicode 名称' (HasName $wb '税率_日本') 'Unicode 名称被误删'
Check '保留系统名 Print_Area'     (HasName $wb '_xlnm.Print_Area') '打印区域被误删'
Check '保留系统名 _FilterDatabase' (HasName $wb '_xlnm._FilterDatabase') '筛选区域被误删'
Check '删除仅出现在字符串里的 UnusedGood' (-not (HasName $wb 'UnusedGood')) 'UnusedGood 未被删'
Check '删除与内置函数同名的 SUM'  (-not (HasName $wb 'SUM'))   'SUM 未被删'
Check '删除与工作表同名的 Data'   (-not (HasName $wb 'Data'))  'Data 未被删'
Check '删除值为 #REF! 的 BrokenName' (-not (HasName $wb 'BrokenName')) 'BrokenName 未被删'
Check '删除无人引用的 ExtUsed'    (-not (HasName $wb 'ExtUsed'))   'ExtUsed 未被删'
Check '删除无人引用的 ExtUnused'  (-not (HasName $wb 'ExtUnused')) 'ExtUnused 未被删'

$sheetAfter = Get-ZipText $defaultOut 'xl/worksheets/sheet1.xml'
Check '工作表部件逐字节未改动' ($sheetAfter -eq $sheetBefore) 'sheet1.xml 被改动了'

$fAfter = Get-Formulas $sheetAfter
$fSame = ($fAfter.Count -eq $formulaList.Count)
if ($fSame) {
    for ($i = 0; $i -lt $formulaList.Count; $i++) {
        if ($fAfter[$i] -ne $formulaList[$i]) { $fSame = $false; break }
    }
}
Check ('全部 {0} 条公式一字未改' -f $formulaList.Count) $fSame ("实际: " + ($fAfter -join ' | '))

Check '保留被公式引用的外部链接 [1]' ($wb -match '<externalReference\b') 'externalReferences 被整块删掉了'
Check '保留 externalLink1.xml 部件' ((Get-ZipNames $defaultOut) -contains 'xl/externalLinks/externalLink1.xml') 'externalLink1.xml 被误删'
Check '删除无人引用的 externalLink2.xml' (-not ((Get-ZipNames $defaultOut) -contains 'xl/externalLinks/externalLink2.xml')) 'externalLink2.xml 未被删'

$rels = Get-ZipText $defaultOut 'xl/_rels/workbook.xml.rels'
Check '保留 rId2 关系' ($rels -match 'Id="rId2"') 'rId2 关系被误删'
Check '删除 rId3 关系' (-not ($rels -match 'Id="rId3"')) 'rId3 关系未被删'

$ct = Get-ZipText $defaultOut '[Content_Types].xml'
Check '删除 externalLink2 的 ContentType 记录' (-not ($ct -match 'externalLink2')) 'Content_Types 里残留 externalLink2'
Check '保留 externalLink1 的 ContentType 记录' ($ct -match 'externalLink1') 'Content_Types 里丢了 externalLink1'

# ---------------------------------------------------------------- 2. 保守模式
Section '2. 保守模式'

$r = Run @($fixture, '-o', $consOut, '--conservative')
Check '退出码为 0' ($r.Code -eq 0) ("实际退出码 " + $r.Code)
$wb = Get-ZipText $consOut 'xl/workbook.xml'
Check '仍删除 BrokenName'        (-not (HasName $wb 'BrokenName')) 'BrokenName 未被删'
Check '保留 UnusedGood'          (HasName $wb 'UnusedGood')   'UnusedGood 被误删'
Check '保留 SUM'                 (HasName $wb 'SUM')          'SUM 被误删'
Check '保留 Data'                (HasName $wb 'Data')         'Data 被误删'
Check '保留 ExtUnused'           (HasName $wb 'ExtUnused')    'ExtUnused 被误删'
Check '因 ExtUnused 保留而留下 [2]' ((Get-ZipNames $consOut) -contains 'xl/externalLinks/externalLink2.xml') 'externalLink2.xml 被误删'

# ---------------------------------------------------------------- 3. 幂等性
Section '3. 幂等性'

$r = Run @($defaultOut, '-o', $idemOut)
Check '退出码为 0' ($r.Code -eq 0) ("实际退出码 " + $r.Code)
Check '报告无可清理' ($r.Out -match '没有发现可清理的垃圾') $r.Out
Check '未写出新文件' (-not (Test-Path -LiteralPath $idemOut)) '不该写文件却写了'

# ---------------------------------------------------------------- 4. 防呆
Section '4. 防呆'

$r = Run @($fixture, '-o', $fixture)
Check '拒绝 输出=源文件' ($r.Code -ne 0) '没有拒绝，可能覆盖了源文件'
Check '源文件未被破坏' ((Get-Sha $fixture) -eq $fixtureSha) '夹具被改动了'

$fakeXlsb = Join-Path $work 'fake.xlsb'
Set-Content -LiteralPath $fakeXlsb -Value 'not a real xlsb' -Encoding ASCII
$r = Run @($fakeXlsb)
Check '拒绝 .xlsb' ($r.Code -ne 0) '没有拒绝 .xlsb'

$notZip = Join-Path $work 'notxlsx.xlsx'
Set-Content -LiteralPath $notZip -Value 'this is plain text, not a zip' -Encoding ASCII
$r = Run @($notZip)
Check '拒绝 非 zip 的 .xlsx' ($r.Code -ne 0) '没有拒绝非 zip 文件'
Check '给出友好错误信息' ($r.Out -match '无法识别的文件格式|无法作为 zip/xlsx 打开') $r.Out

$r = Run @()
Check '无参数时报错（cli 版）' ($r.Code -eq 2) ("实际退出码 " + $r.Code)

# ---------------------------------------------------------------- 5. dry-run
Section '5. --dry-run'

$r = Run @($fixture, '-o', $dryOut, '--dry-run')
Check '退出码为 0' ($r.Code -eq 0) ("实际退出码 " + $r.Code)
Check '未写出文件' (-not (Test-Path -LiteralPath $dryOut)) 'dry-run 却写了文件'
Check '报告里仍给出分析结果' ($r.Out -match '值已失效，删除\s*:\s*1') $r.Out

# ---------------------------------------------------------------- 6. in-place
Section '6. --in-place'

Copy-Item -LiteralPath $fixture -Destination $inplace -Force
$beforeInplace = Get-Sha $inplace
$r = Run @($inplace, '--in-place')
Check '退出码为 0' ($r.Code -eq 0) ("实际退出码 " + $r.Code)
Check '原文件被就地替换' ((Get-Sha $inplace) -ne $beforeInplace) '原文件没变'
Check '生成了 .bak 备份' (Test-Path -LiteralPath ($inplace + '.bak')) '没有 .bak'
Check '.bak 与原文件逐字节相同' ((Get-Sha ($inplace + '.bak')) -eq $beforeInplace) '.bak 内容不对'
Check '替换后的文件已清理' (-not (HasName (Get-ZipText $inplace 'xl/workbook.xml') 'BrokenName')) '没清理干净'
Check '没有残留临时文件' (-not (Test-Path -LiteralPath ($inplace + '.xlsxdoctor.tmp'))) '临时文件没清掉'

# ---------------------------------------------------------------- 7. 结构校验
Section '7. 清理后包结构'

$names = Get-ZipNames $defaultOut
Check '无残留 externalLinks 部件' (@($names | Where-Object { $_ -like 'xl/externalLinks/*' }).Count -eq 2) ("实际: " + (@($names | Where-Object { $_ -like 'xl/externalLinks/*' }) -join ', '))
Check '无残留临时文件' (-not ($names -contains 'xl/workbook.xml.tmp')) '有残留'

try {
    $x = [xml](Get-ZipText $defaultOut 'xl/workbook.xml')
    Check 'workbook.xml 可被 XML 解析' ($null -ne $x) ''
} catch {
    Check 'workbook.xml 可被 XML 解析' $false $_.Exception.Message
}

# ---------------------------------------------------------------- 8. 真实 Excel 打开（可选）
Section '8. 真实 Excel 打开（可选）'
$excelTested = $false
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try {
        $book = $excel.Workbooks.Open($defaultOut, 0, $true)
        try {
            Check 'Excel 能打开清理后的文件' ($book.Sheets.Count -ge 1) '打不开'
            Check '工作表数量正确' ($book.Sheets.Count -eq 1) ("实际 " + $book.Sheets.Count)
            $excelTested = $true
        } finally { $book.Close($false) }
    } finally {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
} catch {
    Write-Host '  [SKIP] 本机没有可用的 Excel，跳过' -ForegroundColor DarkYellow
}

# ---------------------------------------------------------------- 9. 关系目标边界情况
Section '9. 关系目标边界情况'

function Write-ZipParts([string]$path, $parts) {
    if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    $fs = [System.IO.File]::Create($path)
    try {
        $zip = [System.IO.Compression.ZipArchive]::new($fs, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            $utf8 = [System.Text.UTF8Encoding]::new($false)
            foreach ($k in $parts.Keys) {
                $e = $zip.CreateEntry($k)
                $s = $e.Open()
                try { $sw = [System.IO.StreamWriter]::new($s, $utf8); try { $sw.Write($parts[$k]) } finally { $sw.Dispose() } }
                finally { $s.Dispose() }
            }
        } finally { $zip.Dispose() }
    } finally { $fs.Dispose() }
}

# 造一个带「形状超链接」的最小 xlsx。Target 由参数决定：
#   '#Sheet1!A1'          -> 纯片段 URI（指向本工作簿单元格），完全合法
#   'does-not-exist.xml'  -> 真正悬空的关系，源文件自带的缺陷
function New-RelsFixture([string]$path, [string]$relTarget) {
    $drawingRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="' + $relTarget + '"/></Relationships>'
    Write-ZipParts $path ([ordered]@{
        '[Content_Types].xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>'
        '_rels/.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
        'xl/workbook.xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="Junk">#REF!</definedName></definedNames></workbook>'
        'xl/_rels/workbook.xml.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'
        'xl/worksheets/sheet1.xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><drawing r:id="rId1"/></worksheet>'
        'xl/worksheets/_rels/sheet1.xml.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>'
        'xl/drawings/drawing1.xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp macro=""><xdr:nvSpPr><xdr:cNvPr id="2" name="Shape 1"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/></xdr:sp></xdr:twoCellAnchor></xdr:wsDr>'
        'xl/drawings/_rels/drawing1.xml.rels' = $drawingRels
    })
}

# 9.1 纯片段 URI（形状超链接指向本工作簿单元格）—— 合法构造，不能被当成错误
$fragFix = Join-Path $work 'rels-frag.xlsx'
New-RelsFixture $fragFix '#Sheet1!A1'
$fragOut = Join-Path $work 'rels-frag-out.xlsx'
$r = Run @($fragFix, '-o', $fragOut)
Check '纯片段 URI 关系不再被误判' ($r.Code -eq 0) ("退出码 " + $r.Code + " | " + ($r.Out -replace "\s+", " "))
Check '仍产出输出文件' (Test-Path -LiteralPath $fragOut) '输出被误删'
Check '报告显示校验通过' ($r.Out -match '包校验\s*:\s*通过') $r.Out
if (Test-Path -LiteralPath $fragOut) {
    Check 'drawing 的 .rels 逐字节未改动' ((Get-ZipText $fragOut 'xl/drawings/_rels/drawing1.xml.rels') -eq (Get-ZipText $fragFix 'xl/drawings/_rels/drawing1.xml.rels')) '部件被改动了'
}

# 9.2 真正悬空的关系 —— 源文件自带的缺陷，工具应放行并标注，而不是把自己的产物删掉
$dangleFix = Join-Path $work 'rels-dangle.xlsx'
New-RelsFixture $dangleFix 'does-not-exist.xml'
$dangleOut = Join-Path $work 'rels-dangle-out.xlsx'
$r = Run @($dangleFix, '-o', $dangleOut)
Check '源文件自带的悬空关系不再导致失败' ($r.Code -eq 0) ("退出码 " + $r.Code + " | " + ($r.Out -replace "\s+", " "))
Check '仍产出输出文件' (Test-Path -LiteralPath $dangleOut) '输出被误删'
Check '报告标注为源文件自带' ($r.Out -match '源文件自带') $r.Out
Check '报告不再说"校验失败"' (-not ($r.Out -match '校验失败')) $r.Out
# ============================================================ 汇总

Write-Host ''
Write-Host '============================================' -ForegroundColor White
Write-Host ("通过 {0}   失败 {1}" -f $script:pass, $script:fail) -ForegroundColor $(if ($script:fail -eq 0) { 'Green' } else { 'Red' })
if ($script:fail -gt 0) {
    Write-Host '失败项：' -ForegroundColor Red
    foreach ($n in $script:failedNames) { Write-Host ("  - {0}" -f $n) -ForegroundColor Red }
}
Write-Host ("工作目录: {0}" -f $work)
Write-Host ''

if ($script:fail -eq 0) {
    Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
    exit 0
}
exit 1
