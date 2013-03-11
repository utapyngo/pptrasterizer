param(
    [string] $pfilename
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

$transitionMembers = ('AdvanceOnClick', 'AdvanceOnTime', 
    'AdvanceTime', 'Duration', 'EntryEffect', 'Hidden', 'Speed')


function Convert-Slide($original_slide, $slide, $slidesPath) {
    # image
    $slide_image_file_name = Join-Path $slidesPath "Slide$i.png"
    $original_slide.export($slide_image_file_name, "PNG") | Out-Null
    $slide.Shapes.AddPicture($slide_image_file_name, $false, $true, 0, 0, $slide.Parent.PageSetup.SlideHeight, $slide.Parent.PageSetup.SlideWidth) | Out-Null
    # media
    foreach ($shape in $original_slide.Shapes) {
        if ($shape.MediaType) {
            $shape.Copy() | Out-Null
            $slide.Shapes.Paste() | Out-Null
        }
    }
    # notes
    $original_slide.NotesPage.Shapes.Item(2).TextFrame.TextRange.Copy() | Out-Null
    $slide.NotesPage.Shapes.Item(2).TextFrame.TextRange.Paste() | Out-Null
    # transition
    foreach ($memberName in $transitionMembers) {
        if (Get-Member -InputObject $original_slide.SlideShowTransition -Name $memberName) {
            $member = Invoke-Expression ('$' + "original_slide.SlideShowTransition.$memberName")
            Invoke-Expression ('$' + "slide.SlideShowTransition.$memberName = " + $member)
        }
    }
}

function Convert-Presentation($pfilename) {
    $pfilename = Resolve-Path $pfilename
    $path = Split-Path $pfilename
    $filename = Split-Path $pfilename -Leaf
    $name = $filename.substring(0, $filename.lastindexOf("."))
    $slidesPath = "$path\PhotoAlbumSlides"
    $slideShowFileName = "$path\$name - rasterized.pps"
    $rasterizedPdfFileName = "$path\$name - rasterized.pdf"
    mkdir $slidesPath -ErrorAction SilentlyContinue | Out-Null
    $application = New-Object -ComObject "PowerPoint.Application"
    try {
        $presentation = $application.Presentations.Open($pfilename)
        try {
            $photoAlbum = $application.Presentations.Add($true)
            try {
                $photoAlbum.PageSetup.SlideSize = $presentation.PageSetup.SlideSize
                foreach ($original_slide in $presentation.Slides) {
                    $i = $original_slide.SlideIndex
                    Write-Host "Processing slide $i"
                    $slide = $photoAlbum.Slides.Add($photoAlbum.Slides.Count + 1, $ppLayoutBlank)
                    Convert-Slide $original_slide $slide $slidesPath
                }
                $photoAlbum.SaveAs($slideShowFileName, $ppSaveAsShow, 0) | Out-Null
                $photoAlbum.SaveAs($rasterizedPdfFileName, $ppSaveAsPDF, 0) | Out-Null
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

Convert-Presentation $pfilename
