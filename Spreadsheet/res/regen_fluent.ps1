# Appends new Fluent icons (or rendered glyphs) onto fluent_small.bmp /
# fluent_large.bmp by **rebuilding the BMP byte stream**, not by round-
# tripping through System.Drawing.Bitmap.Save — GDI+ collapses the
# original strip's straight-alpha channel into a solid A=255, which
# breaks MFC's alpha-blended ribbon rendering.
#
# Run from this directory:
#     pwsh -File regen_fluent.ps1
#
# Each entry is (folder, snake) for a Fluent SVG, or
# ('__glyph_<char>', '<snake>') to render a single glyph for icons
# Fluent doesn't ship (used for %).

Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase, System.Drawing

$newIcons = @(
    @('Table Simple',   'table_simple'),    # slot 28: fmt_general
    @('Number Symbol',  'number_symbol'),   # slot 29: fmt_number
    @('Wallet',         'wallet'),          # slot 30: fmt_currency
    @('__glyph_%',      'percent'),         # slot 31: fmt_percent
    @('Calendar LTR',   'calendar_ltr')     # slot 32: fmt_date
)

$baseUrl = 'https://raw.githubusercontent.com/microsoft/fluentui-system-icons/main/assets'
$cacheDir = Join-Path $PSScriptRoot 'fluent_cache'
if (-not (Test-Path $cacheDir)) { New-Item -ItemType Directory $cacheDir | Out-Null }

function Get-Svg([string]$folder, [string]$snake, [int]$size) {
    $url = "$baseUrl/$([uri]::EscapeDataString($folder))/SVG/ic_fluent_${snake}_${size}_regular.svg"
    $cachePath = Join-Path $cacheDir "${snake}_${size}.svg"
    if (-not (Test-Path $cachePath)) {
        Invoke-WebRequest -Uri $url -OutFile $cachePath -UseBasicParsing
    }
    return Get-Content $cachePath -Raw
}

# Render an icon to a $size x $size BGRA byte buffer with straight
# (non-premultiplied) alpha. Background pixels get (0,0,0,0) so MFC's
# alpha blend leaves them invisible; glyph strokes get Fluent's #212121
# ink at the appropriate alpha.
function Render-Svg-Bgra([string]$svg, [int]$size) {
    $pathMatches = [regex]::Matches($svg, '<path\s+[^/]*?d="([^"]+)"[^/]*?/>')
    $dv = [System.Windows.Media.DrawingVisual]::new()
    $ctx = $dv.RenderOpen()
    $ink = [System.Windows.Media.Color]::FromRgb(33, 33, 33)
    $brush = [System.Windows.Media.SolidColorBrush]::new($ink)
    foreach ($m in $pathMatches) {
        $geom = [System.Windows.Media.Geometry]::Parse($m.Groups[1].Value)
        $ctx.DrawGeometry($brush, $null, $geom)
    }
    $ctx.Close()
    # RenderTargetBitmap produces premultiplied alpha. Copy into a
    # FormatConvertedBitmap targeting straight Bgra32.
    $rtb = [System.Windows.Media.Imaging.RenderTargetBitmap]::new($size, $size, 96, 96, [System.Windows.Media.PixelFormats]::Pbgra32)
    $rtb.Render($dv)
    $conv = [System.Windows.Media.Imaging.FormatConvertedBitmap]::new($rtb, [System.Windows.Media.PixelFormats]::Bgra32, $null, 0.0)
    $stride = $size * 4
    $buf = New-Object byte[] ($size * $stride)
    $conv.CopyPixels($buf, $stride, 0)
    return ,$buf
}

function Render-Glyph-Bgra([string]$ch, [int]$size) {
    # Draw the glyph onto a transparent Bgra32 wbm using DrawingContext.DrawText,
    # then read the pixels straight back.
    $dv = [System.Windows.Media.DrawingVisual]::new()
    $ctx = $dv.RenderOpen()
    $ink = [System.Windows.Media.Color]::FromRgb(33, 33, 33)
    $brush = [System.Windows.Media.SolidColorBrush]::new($ink)
    $tf = [System.Windows.Media.Typeface]::new(
        [System.Windows.Media.FontFamily]::new('Segoe UI'),
        [System.Windows.FontStyles]::Normal,
        [System.Windows.FontWeights]::Bold,
        [System.Windows.FontStretches]::Normal)
    $emSize = [double]($size * 0.78)
    $ft = [System.Windows.Media.FormattedText]::new(
        $ch,
        [System.Globalization.CultureInfo]::InvariantCulture,
        [System.Windows.FlowDirection]::LeftToRight,
        $tf, $emSize, $brush, 1.0)
    $x = ($size - $ft.Width) / 2.0
    $y = ($size - $ft.Height) / 2.0
    $ctx.DrawText($ft, [System.Windows.Point]::new($x, $y))
    $ctx.Close()
    $rtb = [System.Windows.Media.Imaging.RenderTargetBitmap]::new($size, $size, 96, 96, [System.Windows.Media.PixelFormats]::Pbgra32)
    $rtb.Render($dv)
    $conv = [System.Windows.Media.Imaging.FormatConvertedBitmap]::new($rtb, [System.Windows.Media.PixelFormats]::Bgra32, $null, 0.0)
    $stride = $size * 4
    $buf = New-Object byte[] ($size * $stride)
    $conv.CopyPixels($buf, $stride, 0)
    return ,$buf
}

# Append icons to a 32bpp BMP at byte level. Preserves the existing
# pixel data verbatim and grafts the new icons onto the right side of
# every row. BMPs are bottom-up so we have to be careful with row order.
function Append-To-Bmp([string]$inPath, [string]$outPath, [int]$size) {
    $bytes = [System.IO.File]::ReadAllBytes($inPath)

    $dataOffset = [BitConverter]::ToUInt32($bytes, 10)
    $bpp        = [BitConverter]::ToUInt16($bytes, 28)
    $oldW       = [BitConverter]::ToInt32($bytes, 18)
    $h          = [BitConverter]::ToInt32($bytes, 22)
    if ($bpp -ne 32) { throw "expected 32bpp; got $bpp" }
    if ($h -ne $size) { throw "BMP height ($h) != strip size ($size)" }

    $oldStride = $oldW * 4
    $newStride = ($oldW + $newIcons.Count * $size) * 4
    $newW = $oldW + $newIcons.Count * $size

    # Render every new icon into a [byte[]] of size $size*$size*4 (BGRA, top-down).
    $glyphBufs = @()
    foreach ($pair in $newIcons) {
        $folder, $snake = $pair
        if ($folder -like '__glyph_*') {
            $ch = $folder.Substring(8)
            $glyphBufs += ,(Render-Glyph-Bgra $ch $size)
        } else {
            $svg = Get-Svg $folder $snake $size
            $glyphBufs += ,(Render-Svg-Bgra $svg $size)
        }
    }

    # New file: header (54 bytes) + new pixel data.
    $newPixelBytes = $newStride * $h
    $newFile = New-Object byte[] ($dataOffset + $newPixelBytes)

    # Copy and patch the header.
    [Array]::Copy($bytes, 0, $newFile, 0, $dataOffset)
    # Width (offset 18) and image size (offset 34) are little-endian Int32.
    [BitConverter]::GetBytes([Int32]$newW).CopyTo($newFile, 18)
    [BitConverter]::GetBytes([UInt32]$newPixelBytes).CopyTo($newFile, 34)
    # File size (offset 2) is total file size.
    [BitConverter]::GetBytes([UInt32]($dataOffset + $newPixelBytes)).CopyTo($newFile, 2)

    # Each BMP row corresponds to image y = (h - 1 - rowIdx). Copy old pixels
    # verbatim and append the new icons.
    for ($rowIdx = 0; $rowIdx -lt $h; $rowIdx++) {
        # Old → new (existing portion of the row).
        [Array]::Copy($bytes, $dataOffset + $rowIdx * $oldStride,
                      $newFile, $dataOffset + $rowIdx * $newStride,
                      $oldStride)
        # The image y for this BMP row.
        $y = $h - 1 - $rowIdx
        # Append each new icon's row at column oldW + i*size.
        for ($i = 0; $i -lt $glyphBufs.Count; $i++) {
            $buf = $glyphBufs[$i]
            $srcRowOff = $y * ($size * 4)
            $dstOff = $dataOffset + $rowIdx * $newStride + ($oldW + $i * $size) * 4
            [Array]::Copy($buf, $srcRowOff, $newFile, $dstOff, $size * 4)
        }
    }

    [System.IO.File]::WriteAllBytes($outPath, $newFile)
    Write-Host "wrote $outPath  ($newW x $h, $($newIcons.Count) icons appended starting at x=$oldW)"
}

Append-To-Bmp (Join-Path $PSScriptRoot 'fluent_small.bmp') (Join-Path $PSScriptRoot 'fluent_small.bmp') 16
Append-To-Bmp (Join-Path $PSScriptRoot 'fluent_large.bmp') (Join-Path $PSScriptRoot 'fluent_large.bmp') 32
