# Unregister the "Rasterize" command

$nothing = $true
Set-Location HKCU:\Software\Classes
'Show.8 SlideShow.8 Show.12 SlideShow.12 ShowMacroEnabled.12 SlideShowMacroEnabled.12'.Split() |% {
    if (Test-Path "PowerPoint.$_\shell\Rasterize") {
        echo "Uninstalling for PowerPoint.$_"
        Remove-Item -Recurse "PowerPoint.$_\shell\Rasterize"
        if (((Get-ChildItem "PowerPoint.$_\shell").Count -eq 0) -and 
            ((Get-ChildItem "PowerPoint.$_").Count -eq 1) -and 
            ((Get-ItemProperty "PowerPoint.$_\shell").Count -eq 0) -and
            ((Get-ItemProperty "PowerPoint.$_").Count -eq 0)) {
            Remove-Item -Recurse "PowerPoint.$_"
        }
        $nothing = $false
    }
}
if ($nothing) {
    echo "Nothing to uninstall."
}
