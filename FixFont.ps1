Set-StrictMode -Version Latest

$version = "2019-05-23"
Write-Output "FixFont - built $version by Mikio Ogawa, Shinichi Akiyama"


function Write-Message ([string]$message) {
    Write-Output $message | Out-File $logfile -Append
    Write-Output $message
}

function backupFile ($path) {
    $directory = Split-Path $path
    $fileNameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($path)
    $extension = $path.Extension
    $backup = Join-Path $directory "$fileNameWithoutExtension - backup$extension"
    $num = 2
    while (Test-Path $backup) {
        $backup = Join-Path $directory "$fileNameWithoutExtension - backup ($num)$extension"
        $num = $num + 1
    }
    Copy-Item $path $backup
    $backup
}
function ReplaceMajorFont ($textRange) {
    $text = $textRange.Text -replace "[\x0b\x0d]", "`n"
    $font = $textRange.Font
    $fontNameFarEast = $font.NameFarEast
    $font.NameFarEast = "+mj-ea"
    if (($null -ne $text) -and ($null -ne $fontNameFarEast) -and ($font.NameFarEast -ne $fontNameFarEast)) {
        Write-Message "[$text] Asian font has been converted from $fontNameFarEast to +mj-ea"
    }
    $fontNameAscii = $font.NameAscii
    $font.NameAscii = "+mj-lt"
    if (($text -ne "") -and ($null -ne $fontNameAscii) -and ($font.NameAscii -ne $fontNameAscii)) {
        Write-Message "[$text] Latin font has been converted from $fontNameAscii to +mj-lt"
    }
}

function ReplaceMinorFont ($textRange) {
    $text = $textRange.Text -replace "[\x0b\x0d]", "`n"
    $font = $textRange.Font
    $fontNameFarEast = $font.NameFarEast
    $font.NameFarEast = "+mn-ea"
    if (($text -ne "") -and ($null -ne $fontNameFarEast) -and ($font.NameFarEast -ne $fontNameFarEast)) {
        Write-Message "[$text] Asian font has been converted from $fontNameFarEast to +mn-ea"
    }
    $fontNameAscii = $font.NameAscii
    $font.NameAscii = "+mn-lt"
    if (($text -ne "") -and ($null -ne $fontNameAscii) -and ($font.NameAscii -ne $fontNameAscii)) {
        Write-Message "[$text] Latin font has been converted from $fontNameAscii to +mn-lt"
    }
}

function treatShape ($shape) {
    if ($shape.HasTextFrame -eq [Microsoft.Office.Core.MsoTriState]::msoTrue) {
        foreach ($textRange in $shape.TextFrame.TextRange) {
            ReplaceMinorFont $textRange
        }
    } elseif ($shape.HasTable -eq [Microsoft.Office.Core.MsoTriState]::msoTrue) {
        foreach ($row in $shape.Table.rows) {
            foreach ($cell in $row.cells) {
                ReplaceMinorFont $cell.shape.TextFrame.TextRange
            }
        }
    } elseif ($shape.GroupItems) {
        foreach ($item in $shape.GroupItems) {
            treatShape $item
        }
    }
}

Add-Type -AssemblyName Office
$app = New-Object -ComObject PowerPoint.Application
Write-Output "PowoerPoint version $($app.version)"

foreach ($file in Get-ChildItem -Path $args) {
    $logfile = [System.IO.Path]::GetDirectoryName($file) + "\" + [System.IO.Path]::GetFileNameWithoutExtension($file) + ".log"

    $backup = backupFile $file
    Write-Message "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") $file was backed up to $backup."

    $presentation = $app.Presentations.Open($file)
    $app.WindowState = [Microsoft.Office.Interop.PowerPoint.PpWindowState]::ppWindowMinimized
    Write-Message "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") $file was opened."

    foreach ($slide in $presentation.Slides) {
        Write-Message "--- Slide $($slide.SlideIndex) ---"
        foreach ($shape in $slide.Shapes) {
            if (($null -ne $slide.Shapes.Title) -and $shape.Id -eq $slide.Shapes.Title.Id) {
                if ($shape.HasTextFrame -eq [Microsoft.Office.Core.MsoTriState]::msoTrue) {
                    foreach ($textRange in $shape.TextFrame.TextRange) {
                        ReplaceMajorFont $textRange
                    }
                }
            } else {
                treatShape $shape
            }
        }
    }

    $presentation.Save()
    Write-Message "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss") $file was saved."
    $presentation.Close()
}
Write-Output "All files were processed."
