Add-Type -AssemblyName System.Drawing

function Convert-PngToIco {
    param(
        [string]$PngPath,
        [string]$IcoPath
    )

    $srcBmp = [System.Drawing.Bitmap]::FromFile($PngPath)
    $sizes = @(16, 32, 48, 64, 128, 256)

    $pngStreams = @()
    foreach ($sz in $sizes) {
        $resized = New-Object System.Drawing.Bitmap($sz, $sz)
        $g = [System.Drawing.Graphics]::FromImage($resized)
        $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
        $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
        $g.DrawImage($srcBmp, 0, 0, $sz, $sz)
        $g.Dispose()

        $ms = New-Object System.IO.MemoryStream
        $resized.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
        $resized.Dispose()
        $pngStreams += $ms
    }
    $srcBmp.Dispose()

    $fs = [System.IO.File]::Create($IcoPath)
    $bw = New-Object System.IO.BinaryWriter($fs)

    # ICONDIR (6 bytes)
    $bw.Write([uint16]0)
    $bw.Write([uint16]1)
    $bw.Write([uint16]$sizes.Count)

    $offset = 6 + (16 * $sizes.Count)

    for ($i = 0; $i -lt $sizes.Count; $i++) {
        $sz = $sizes[$i]
        $len = $pngStreams[$i].Length
        $w = if ($sz -ge 256) { [byte]0 } else { [byte]$sz }
        $h = if ($sz -ge 256) { [byte]0 } else { [byte]$sz }

        $bw.Write($w)
        $bw.Write($h)
        $bw.Write([byte]0)
        $bw.Write([byte]0)
        $bw.Write([uint16]1)
        $bw.Write([uint16]32)
        $bw.Write([uint32]$len)
        $bw.Write([uint32]$offset)

        $offset += $len
    }

    for ($i = 0; $i -lt $sizes.Count; $i++) {
        $bytes = $pngStreams[$i].ToArray()
        $bw.Write($bytes)
        $pngStreams[$i].Dispose()
    }

    $bw.Flush()
    $bw.Close()
    $fs.Close()
    Write-Host "Icono multi-resolución generado con éxito en: $IcoPath"
}

Convert-PngToIco "OCCESS.png" "icon.ico"
