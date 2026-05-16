# Appends new Fluent icons (or rendered glyphs) onto fluent_small.bmp and
# fluent_large.bmp, in-place. Run from this directory:
#     pwsh -File regen_fluent.ps1
#
# This script is non-destructive to the existing strip — it loads the
# current images, widens the canvas, and paints new icons at the right
# edge. To add more icons, append to $newIcons; the index in the array
# defines the slot index after the existing icons (28 today, so first
# appended slot is 28).
#
# Each entry is (folder, snake) for a Fluent SVG, or
# ('__glyph_<char>', '<snake>') to render a single glyph for icons
# Fluent doesn't ship (we use this for %).

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

function Render-Svg([string]$svg, [int]$size) {
    # MFC's ribbon image loader color-keys pure black (RGB 0,0,0) as
    # transparent, so the strokes have to be drawn in #212121 (33,33,33)
    # — which also matches the Fluent SVG fill ("fill=#212121").
    $pathMatches = [regex]::Matches($svg, '<path\s+[^/]*?d="([^"]+)"[^/]*?/>')
    $dv = [System.Windows.Media.DrawingVisual]::new()
    $ctx = $dv.RenderOpen()
    $fluentColor = [System.Windows.Media.Color]::FromRgb(33, 33, 33)
    $brush = [System.Windows.Media.SolidColorBrush]::new($fluentColor)
    foreach ($m in $pathMatches) {
        $geom = [System.Windows.Media.Geometry]::Parse($m.Groups[1].Value)
        $ctx.DrawGeometry($brush, $null, $geom)
    }
    $ctx.Close()
    $rtb = [System.Windows.Media.Imaging.RenderTargetBitmap]::new($size, $size, 96, 96, [System.Windows.Media.PixelFormats]::Pbgra32)
    $rtb.Render($dv)
    # PNG round-trip into System.Drawing — simplest interop.
    $enc = [System.Windows.Media.Imaging.PngBitmapEncoder]::new()
    $enc.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($rtb))
    $ms = [System.IO.MemoryStream]::new()
    $enc.Save($ms)
    $ms.Position = 0
    return [System.Drawing.Bitmap]::new($ms)
}

function Render-Glyph([string]$ch, [int]$size) {
    # For glyphs Fluent doesn't ship (e.g. %). Segoe UI Semibold at ~70% of
    # the slot height visually matches the stroke weight of regular Fluent
    # icons; tweak as needed.
    $bmp = [System.Drawing.Bitmap]::new($size, $size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    $g.Clear([System.Drawing.Color]::Transparent)
    $fontSize = [int]($size * 0.72)
    $font = [System.Drawing.Font]::new('Segoe UI', $fontSize, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
    $sf = [System.Drawing.StringFormat]::new()
    $sf.Alignment = [System.Drawing.StringAlignment]::Center
    $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
    # Match the Fluent #212121 ink color so MFC's color-keying treats only
    # the background as transparent.
    $brush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 33, 33, 33))
    $rect = [System.Drawing.RectangleF]::new(0, 0, $size, $size)
    $g.DrawString($ch, $font, $brush, $rect, $sf)
    $brush.Dispose()
    $g.Dispose()
    return $bmp
}

function Extend-Strip([string]$inPath, [string]$outPath, [int]$size) {
    # Load via MemoryStream so GDI+ doesn't hold an exclusive lock on the
    # source file — otherwise saving back to the same path fails.
    $bytes = [System.IO.File]::ReadAllBytes($inPath)
    $existing = [System.Drawing.Bitmap]::new([System.IO.MemoryStream]::new($bytes))
    try {
        $oldWidth = $existing.Width
        $newWidth = $oldWidth + ($newIcons.Count * $size)
        $strip = [System.Drawing.Bitmap]::new($newWidth, $size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
        $g = [System.Drawing.Graphics]::FromImage($strip)
        $g.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceCopy
        $g.Clear([System.Drawing.Color]::Transparent)
        # Copy the existing strip verbatim.
        $g.DrawImage($existing, 0, 0, $oldWidth, $size)
        # Paint the new icons at the right edge.
        for ($i = 0; $i -lt $newIcons.Count; $i++) {
            $folder, $snake = $newIcons[$i]
            if ($folder -like '__glyph_*') {
                $ch = $folder.Substring(8)
                $bmp = Render-Glyph $ch $size
            } else {
                $svg = Get-Svg $folder $snake $size
                $bmp = Render-Svg $svg $size
            }
            $g.DrawImage($bmp, ($oldWidth + $i * $size), 0, $size, $size)
            $bmp.Dispose()
        }
        $g.Dispose()
        $strip.Save($outPath, [System.Drawing.Imaging.ImageFormat]::Bmp)
        $strip.Dispose()
        Write-Host "wrote $outPath  ($newWidth x $size, $($newIcons.Count) icons appended starting at x=$oldWidth)"
    } finally {
        $existing.Dispose()
    }
}

Extend-Strip (Join-Path $PSScriptRoot 'fluent_small.bmp') (Join-Path $PSScriptRoot 'fluent_small.bmp') 16
Extend-Strip (Join-Path $PSScriptRoot 'fluent_large.bmp') (Join-Path $PSScriptRoot 'fluent_large.bmp') 32
