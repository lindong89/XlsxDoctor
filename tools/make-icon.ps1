<#
    PNG -> ICO 转换器（不依赖任何第三方工具）

    为什么要自己写：csc.exe 的 /win32icon 只认 .ico，不认 .png；
    本机也没有 ImageMagick。

    生成的 .ico 是多尺寸的，小尺寸用 BMP(DIB)、大尺寸用 PNG 压缩 ——
    这是资源管理器/任务栏/Alt-Tab 都能正确取到清晰图标的通行做法。

    用法：
        .\tools\make-icon.ps1
        .\tools\make-icon.ps1 -Png D:\DEV\icons\xlsxdoctor2.png -Ico .\assets\XlsxDoctor.ico
#>
[CmdletBinding()]
param(
    [string]$Png = 'D:\DEV\icons\xlsxdoctor2.png',
    # 留空则自动定位到项目下的 assets\XlsxDoctor.ico
    # （注意：$PSScriptRoot 在 param() 默认值里取不到，只能在函数体里算）
    [string]$Ico,
    # 16/20/24/32/40/48/64 覆盖资源管理器各视图与任务栏，96/128/256 供大图标和高 DPI
    [int[]]$Sizes = @(16, 20, 24, 32, 40, 48, 64, 96, 128, 256)
)

$ErrorActionPreference = 'Stop'

if (-not $Ico) {
    $projectRoot = Split-Path -Parent $PSScriptRoot
    $Ico = Join-Path $projectRoot 'assets\XlsxDoctor.ico'
}

if (-not (Test-Path -LiteralPath $Png)) { throw "找不到源 PNG: $Png" }

Add-Type -AssemblyName System.Drawing

if (-not ('IcoBuilder' -as [type])) {
    # 注意：Add-Type 编译 C# 时不会自动带上已加载的程序集，必须显式 -ReferencedAssemblies
    Add-Type -ReferencedAssemblies 'System.Drawing' -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;

public class IcoBuilder
{
    // 生成多尺寸 ico。尺寸 >= 96 用 PNG 压缩，更小用 BMP(DIB) —— 与主流图标工具一致。
    public static void Build(string pngPath, string icoPath, int[] sizes)
    {
        using (Image src = Image.FromFile(pngPath))
        {
            List<int> dims = new List<int>();
            List<byte[]> blobs = new List<byte[]>();
            foreach (int s in sizes)
            {
                using (Bitmap bmp = Resize(src, s))
                {
                    blobs.Add(s >= 96 ? ToPng(bmp) : ToDib(bmp));
                    dims.Add(s);
                }
            }
            Write(icoPath, dims, blobs);
        }
    }

    static Bitmap Resize(Image src, int size)
    {
        Bitmap bmp = new Bitmap(size, size, PixelFormat.Format32bppArgb);
        using (Graphics g = Graphics.FromImage(bmp))
        {
            g.CompositingQuality = CompositingQuality.HighQuality;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.SmoothingMode = SmoothingMode.HighQuality;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
            g.Clear(Color.Transparent);
            // 等比缩放并居中（源图不是正方形时也不会变形）
            double scale = Math.Min((double)size / src.Width, (double)size / src.Height);
            int w = Math.Max(1, (int)Math.Round(src.Width * scale));
            int h = Math.Max(1, (int)Math.Round(src.Height * scale));
            g.DrawImage(src, new Rectangle((size - w) / 2, (size - h) / 2, w, h));
        }
        return bmp;
    }

    static byte[] ToPng(Bitmap bmp)
    {
        using (MemoryStream ms = new MemoryStream())
        {
            bmp.Save(ms, ImageFormat.Png);
            return ms.ToArray();
        }
    }

    // ICO 里的 BMP 不是完整的 .bmp 文件，而是裸的 BITMAPINFOHEADER + 像素 + AND 掩码，
    // 且高度要写成两倍（上半 XOR 位图、下半 AND 掩码），像素自下而上。
    static byte[] ToDib(Bitmap bmp)
    {
        int w = bmp.Width, h = bmp.Height;
        int stride = w * 4;
        int maskStride = ((w + 31) / 32) * 4;
        byte[] data = new byte[40 + stride * h + maskStride * h];

        WriteU32(data, 0, 40);                    // biSize
        WriteU32(data, 4, (uint)w);               // biWidth
        WriteU32(data, 8, (uint)(h * 2));         // biHeight（两倍）
        WriteU16(data, 12, 1);                    // biPlanes
        WriteU16(data, 14, 32);                   // biBitCount
        WriteU32(data, 16, 0);                    // biCompression = BI_RGB
        WriteU32(data, 20, (uint)(stride * h + maskStride * h));

        BitmapData bd = bmp.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
        try
        {
            int offset = 40;
            for (int y = h - 1; y >= 0; y--)
            {
                Marshal.Copy((IntPtr)((long)bd.Scan0 + (long)y * bd.Stride), data, offset, stride);
                offset += stride;
            }
        }
        finally { bmp.UnlockBits(bd); }

        // AND 掩码留全 0：32bpp 图标的透明度由 alpha 通道决定，掩码只是给老渲染器的兜底
        return data;
    }

    static void WriteU16(byte[] b, int o, int v) { b[o] = (byte)(v & 0xFF); b[o + 1] = (byte)((v >> 8) & 0xFF); }
    static void WriteU32(byte[] b, int o, uint v)
    {
        b[o] = (byte)(v & 0xFF); b[o + 1] = (byte)((v >> 8) & 0xFF);
        b[o + 2] = (byte)((v >> 16) & 0xFF); b[o + 3] = (byte)((v >> 24) & 0xFF);
    }

    static void Write(string path, List<int> dims, List<byte[]> blobs)
    {
        int n = dims.Count;
        int headerSize = 6 + 16 * n;
        byte[] head = new byte[headerSize];
        WriteU16(head, 0, 0);          // idReserved
        WriteU16(head, 2, 1);          // idType = 1 (icon)
        WriteU16(head, 4, n);          // idCount

        int offset = headerSize;
        for (int i = 0; i < n; i++)
        {
            int e = 6 + 16 * i;
            int d = dims[i];
            head[e]     = (byte)(d >= 256 ? 0 : d);   // 256 记作 0
            head[e + 1] = (byte)(d >= 256 ? 0 : d);
            head[e + 2] = 0;                          // bColorCount
            head[e + 3] = 0;                          // bReserved
            WriteU16(head, e + 4, 1);                 // wPlanes
            WriteU16(head, e + 6, 32);                // wBitCount
            WriteU32(head, e + 8, (uint)blobs[i].Length);
            WriteU32(head, e + 12, (uint)offset);
            offset += blobs[i].Length;
        }

        string dir = Path.GetDirectoryName(Path.GetFullPath(path));
        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir)) Directory.CreateDirectory(dir);

        using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
        {
            fs.Write(head, 0, head.Length);
            for (int i = 0; i < n; i++) fs.Write(blobs[i], 0, blobs[i].Length);
        }
    }
}
'@
}

[IcoBuilder]::Build((Resolve-Path -LiteralPath $Png).Path, $Ico, $Sizes)

$fi = Get-Item -LiteralPath $Ico
Write-Host ("  源 PNG : {0}" -f (Resolve-Path -LiteralPath $Png).Path)
Write-Host ("  输出   : {0}" -f $fi.FullName)
Write-Host ("  大小   : {0:N0} 字节" -f $fi.Length)
Write-Host ("  尺寸   : {0}" -f ($Sizes -join ', '))

# 读回自检 1：System.Drawing.Icon 能按尺寸取出多少个
# 变量名不能叫 $ico —— PowerShell 变量名不区分大小写，会和参数 $Ico 撞成同一个变量
$iconProbe = New-Object -TypeName System.Drawing.Icon -ArgumentList $fi.FullName
Write-Host ("  自检   : 默认尺寸 {0}x{1}" -f $iconProbe.Width, $iconProbe.Height)
$iconProbe.Dispose()

# 注意：System.Drawing.Icon 有个已知限制 —— 取不出 256x256 的条目（内部用的
# CreateIconFromResourceEx 不支持），资源管理器走的是另一套 API 所以不受影响。
# 因此 256 单独用 PNG 解码来验证，不能把它算作失败。
$ok = 0
$missing = @()
foreach ($s in $Sizes) {
    $one = New-Object -TypeName System.Drawing.Icon -ArgumentList $fi.FullName, $s, $s
    if ($one.Width -eq $s) { $ok++ }
    elseif ($s -eq 256) { }   # 已知限制，下面单独验
    else { $missing += "$s(得到 $($one.Width))" }
    $one.Dispose()
}

# 读回自检 2：直接解析 ICO 目录，逐个条目校验数据完整性
$bytes = [System.IO.File]::ReadAllBytes($fi.FullName)
$count = [BitConverter]::ToUInt16($bytes, 4)
if ($count -ne $Sizes.Count) { throw "ico 条目数 $count 与预期 $($Sizes.Count) 不符" }
$pngChecked = 0
for ($k = 0; $k -lt $count; $k++) {
    $e = 6 + 16 * $k
    $dim = if ($bytes[$e] -eq 0) { 256 } else { [int]$bytes[$e] }
    $len = [BitConverter]::ToUInt32($bytes, $e + 8)
    $off = [BitConverter]::ToUInt32($bytes, $e + 12)
    if (($off + $len) -gt $bytes.Length) { throw "第 $k 个条目越界" }

    $isPng = ($bytes[$off] -eq 0x89 -and $bytes[$off+1] -eq 0x50)
    if ($isPng) {
        # PNG 条目：解出来确认尺寸正确（256 只能靠这条路验）
        $ms = New-Object System.IO.MemoryStream($bytes, [int]$off, [int]$len)
        $img = [System.Drawing.Image]::FromStream($ms)
        if ($img.Width -ne $dim) { throw "PNG 条目 $dim 实际是 $($img.Width)" }
        $img.Dispose(); $ms.Dispose()
        $pngChecked++
    } else {
        # BMP 条目：BITMAPINFOHEADER 的宽/高（高是两倍）
        $w = [BitConverter]::ToUInt32($bytes, $off + 4)
        $h = [BitConverter]::ToUInt32($bytes, $off + 8)
        if ($w -ne $dim -or $h -ne ($dim * 2)) { throw "BMP 条目 $dim 的头不对 (w=$w h=$h)" }
    }
}
Write-Host ("  自检   : Icon 取出 {0}/{1} 个尺寸（256 属已知限制，已单独校验）" -f $ok, $Sizes.Count)
Write-Host ("  自检   : {0} 个 PNG 条目解码尺寸正确，其余 BMP 条目头正确" -f $pngChecked)
if ($missing.Count -gt 0) { throw ("ico 自检失败，异常尺寸: " + ($missing -join ', ')) }
Write-Host '  图标生成成功。' -ForegroundColor Green
