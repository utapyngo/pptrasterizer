param(
    [string] $pfilename,
    [string] $slideShowFileName
)

if (-not $pfilename) {
    Write-Host "Usage: powershell -ExecutionPolicy Bypass ""$($script:MyInvocation.MyCommand.Path)"" ""Presentation.pptx"""
    return
}

if (-not (Test-Path $pfilename)) {
    Write-Host "File ""$pfilename"" not found"
    return
}

$ppSaveAsShow = 7
$ppSaveAsPDF = 32
$ppLayoutBlank = 12
$msoPlaceholder = 14
$ppPlaceholderBody = 2


$transitionMembers = ('AdvanceOnClick', 'AdvanceOnTime', 'AdvanceTime', 'Duration', 'EntryEffect', 'Hidden', 'Speed')


function Find-Notes($slide) {
    foreach ($shape in $slide.NotesPage.Shapes) {
        if (($shape.Type -eq $msoPlaceholder) -and ($shape.PlaceholderFormat.Type -eq $ppPlaceholderBody)) {
            return $shape
        }
    }
}

function Convert-Slide($original_slide, $slide, $slidesPath) {
    # image
    $width = $slide.Parent.PageSetup.SlideWidth
    $height = $slide.Parent.PageSetup.SlideHeight
    $slide_image_file_name = Join-Path $slidesPath "Slide$i.png"
    $original_slide.export($slide_image_file_name, "PNG",
        $width * 2, $height * 2) | Out-Null
    $slide.Shapes.AddPicture($slide_image_file_name, $false, $true, 0, 0,
        $width, $height) | Out-Null
    # media
    foreach ($shape in $original_slide.Shapes) {
        if ($shape.MediaType) {
            $shape.Copy() | Out-Null
            $slide.Shapes.Paste() | Out-Null
        }
    }
    # notes
    $original_notes = Find-Notes $original_slide
    $original_notes.TextFrame.TextRange.Copy() | Out-Null
    $notes = Find-Notes $slide
    $notes.TextFrame.TextRange.Paste() | Out-Null

    # transition
    foreach ($memberName in $transitionMembers) {
        if (Get-Member -InputObject $original_slide.SlideShowTransition -Name $memberName) {
            $member = Invoke-Expression ('$' + "original_slide.SlideShowTransition.$memberName")
            Invoke-Expression ('$' + "slide.SlideShowTransition.$memberName = ""$member""")
        }
    }
}

function Convert-Presentation($pfilename, $slideShowFileName) {
    $pfilename = Resolve-Path $pfilename
    $path = Split-Path $pfilename
    $filename = Split-Path $pfilename -Leaf
    $name = $filename.substring(0, $filename.lastindexOf("."))
    if (-not $slideShowFileName) {
        $slideShowFileName = "$path\$name - rasterized.pps"
        $rasterizedPdfFileName = "$path\$name - rasterized.pdf"
    }
    if (($env:temp) -and (Test-Path $env:temp)) {
        $slidesPath = "$env:temp\PhotoAlbumSlides$name"
    } else {
        $slidesPath = "$path\PhotoAlbumSlides"
    }
    if (Test-Path $slidesPath) {
        Write-Host "Cleaning $slidesPath"
        Remove-Item -Recurse $slidesPath -ErrorAction SilentlyContinue | Out-Null
    }    
    mkdir $slidesPath -ErrorAction SilentlyContinue | Out-Null
    $application = New-Object -ComObject "PowerPoint.Application"
    try {
        Write-Host "Loading $pfilename"
        $presentation = $application.Presentations.Open($pfilename)
        try {
            $photoAlbum = $application.Presentations.Add($true)
            try {
                $photoAlbum.PageSetup.SlideSize = $presentation.PageSetup.SlideSize
                $photoAlbum.PageSetup.SlideOrientation = $presentation.PageSetup.SlideOrientation
                $photoAlbum.PageSetup.SlideWidth = $presentation.PageSetup.SlideWidth
                $photoAlbum.PageSetup.SlideHeight = $presentation.PageSetup.SlideHeight
                foreach ($original_slide in $presentation.Slides) {
                    $i = $original_slide.SlideIndex
                    Write-Host "Processing slide $i"
                    $slide = $photoAlbum.Slides.Add($photoAlbum.Slides.Count + 1, $ppLayoutBlank)
                    Convert-Slide $original_slide $slide $slidesPath
                }
                Write-Host "Saving $slideShowFileName"
                $photoAlbum.SaveAs($slideShowFileName, $ppSaveAsShow, 0) | Out-Null
                if ($rasterizedPdfFileName) {
                    Write-Host "Saving $rasterizedPdfFileName"
                    $photoAlbum.SaveAs($rasterizedPdfFileName, $ppSaveAsPDF, 0) | Out-Null
                }
            }
            finally {
                $photoAlbum.Close() | Out-Null
            }
        }
        finally {
            $presentation.Close() | Out-Null
            Remove-Item -Recurse $slidesPath -ErrorAction SilentlyContinue | Out-Null
        }
    }
    finally {
        if ($application.Presentations.Count -eq 0) {
            $application.Quit()
        }
    }
}

try {
    (Get-Host).UI.RawUI.WindowTitle = "PowerPoint Presentation Rasterizer"
    Convert-Presentation $pfilename $slideShowFileName
} catch {
    Write-Host $_.Exception.ToString()
    Write-Host "An error occurred. Please report this bug at"
    Write-Host "http://github.com/utapyngo/pptrasterizer/issues"
    cmd /c pause | Out-Null
}

