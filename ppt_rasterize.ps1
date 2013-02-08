<#
As a farsighted presenter
I want to rasterize my slides
So that they do not shift when presenting.

As a fanciful presenter
I want to keep transitions effects between slides
So that everyone sees my effects.

As a cautious presenter
I want my presentation to be converted to PDF
So that I can present it even if PowerPoint is not installed.
#>

param(
    [string] $pfilename
)

if (-not $pfilename) {
    echo "Usage: powershell -ExecutionPolicy Bypass ""$($script:MyInvocation.MyCommand.Path)"" ""Presentation.pptx"""
    return
}

if (-not (Test-Path $pfilename)) {
    echo "File ""$pfilename"" not found"
    return
}

$ppSaveAsShow = 7
$ppSaveAsPNG = 18
$ppSaveAsPDF = 32
$ppLayoutBlank = 12

$pfilename = Resolve-Path $pfilename
$path = Split-Path $pfilename
$filename = Split-Path $pfilename -Leaf
$name = $filename.substring(0, $filename.lastindexOf("."))
$slidesPath = "$path\PhotoAlbumSlides"
$slideShowFileName = "$path\$name.pps"
$pdfFileName = "$path\$name.pdf"
$imagePdfFileName = "$path\$name.img.pdf"

$transitionMembers = ('AdvanceOnClick', 'AdvanceOnTime', 
    'AdvanceTime', 'Duration', 'EntryEffect', 'Hidden', 'Speed')

#$slide_size = None
$application = New-Object -ComObject "PowerPoint.Application"
try {
    $presentation = $application.Presentations.Open($pfilename)
    try {
        # save slide size
        $slide_size = $presentation.PageSetup.SlideSize
        # save transitions
        $transitions = @()
        foreach ($slide in $presentation.Slides) {
            echo "Checking slide $($slide.SlideIndex)"
            $d = @{}
            foreach ($member in $transitionMembers) {
                if (Get-Member -InputObject $slide.SlideShowTransition -Name $member) {
                    $d[$member] = Invoke-Expression ('$' + "slide.SlideShowTransition.$member")
                }
            }
            $transitions += $d
        }
        # save slide images
        echo "Saving pictures"
        $presentation.SaveAs($slidesPath, $ppSaveAsPNG, 0)
    }
    finally {
        $presentation.Close()
    }
    
    # create photo album
    try {
        $photoAlbum = $application.Presentations.Add($true)
        # restore slide size
        if ($slide_size) {
            $photoAlbum.PageSetup.SlideSize = $slide_size
        }
        # restore slide images amd transitions
        $slides = Get-ChildItem -Path $slidesPath -Filter *.png | Sort-Object { [regex]::Replace($_, '\d+', { $args[0].Value.PadLeft(20) }) }
        $slides | ForEach-Object -Begin { $i = 0 } -Process {
            $fn = $_
            echo "Restoring $fn"
            $slide = $photoAlbum.Slides.Add($photoAlbum.Slides.Count + 1, $ppLayoutBlank)
            $dummy = $slide.Shapes.AddPicture((Join-Path $slidesPath $fn), $false, $true, 0, 0)
            foreach ($member in $transitionMembers) {
                if ($transitions[$i].contains($member)) {
                    # setattr(slide.SlideShowTransition, member, transitions[i][member])
                    Invoke-Expression ('$' + "slide.SlideShowTransition.$member = " + $transitions[$i][$member])
                }
            }
            $i++
        }
        # save as PPS
        $photoAlbum.SaveAs($slideShowFileName, $ppSaveAsShow, 0)
        # save as PDF
        $photoAlbum.SaveAs($imagePdfFileName, $ppSaveAsPDF, 0)
        $photoAlbum.Close()
    }
    finally {
        # clean
        Remove-Item -Recurse $slidesPath
    }
    
}
finally {
    $application.Quit()
}
