@echo off

set "files=vars.cmd register.cmd ppt_rasterize.ps1 ppt_rasterize.cmd uninstall.cmd unreg.ps1 update.cmd version.txt"

set "product_name=PowerPoint Presentation Rasterizer"
set "short_name=pptrasterizer"
set "installer_name=setup_%short_name%.bat"
set "version=34"

setlocal DisableDelayedExpansion

echo This will install %product_name% version %version% to you computer.
pause

for /f %%p in ('where powershell') do set "powershell=%%p"

if not exist %powershell% (
    echo Windows Powershell not found
    echo Please install Windows Powershell first.
    pause
    goto :eof
)

set "install_dir=%USERPROFILE%\.%short_name%\"

if exist "%install_dir%uninstall.cmd" (
    echo Uninstalling previous version...
    call %install_dir%uninstall.cmd 2>nul
    pause
)

if not exist "%install_dir%" (
    mkdir "%install_dir%"
)

:: only extract if we are not already there
for %%A in ("%~dp0") do for %%B in ("%install_dir%") do if %%~fA NEQ %%~fB (
    echo Copying files...
    type nul >%install_dir%files.txt
    for %%f in (%files%) do (
        call :get %%f > "%install_dir%\%%f"
        echo   %%f
        echo %%f>>%install_dir%files.txt
    )
    echo files.txt>>%install_dir%files.txt
)

:: add uninstall information to Programs and Features
set "uninstall_key=HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\%short_name%"
reg add %uninstall_key% /f /v DisplayName /d "%product_name%" >nul
reg add %uninstall_key% /f /v InstallLocation /d %install_dir% >nul
reg add %uninstall_key% /f /t REG_DWORD /v NoModify /d 1 >nul
reg add %uninstall_key% /f /t REG_DWORD /v NoRepair /d 1 >nul
reg add %uninstall_key% /f /v UninstallString /d %install_dir%uninstall.cmd >nul
reg add %uninstall_key% /f /v DisplayVersion /d %version% >nul
:: register shell command
call "%install_dir%register.cmd"

:: thank for installing if only not updating
if not defined "%installer_url%" (
    start http://j.mp/pptrasterizer-installed
)

exit /b

:get
set "skip="
for /f "delims=:" %%N in ('findstr /x /n "::begin.%~1" "%~f0"') do if not defined skip set skip=%%N
set "end="
for /f "delims=:" %%N in ('findstr /x /n "::end.%~1" "%~f0"') do if %%N gtr %skip% if not defined end set end=%%N
for /f "skip=%skip% tokens=*" %%A in ('findstr /n "^" "%~f0"') do (
    for /f "delims=:" %%N in ("%%A") do if %%N geq %end% exit /b
    set "line=%%A"
    setlocal EnableDelayedExpansion
    echo(!line:*:=!
    endlocal
)
exit /b

::begin vars.cmd
set "product_name=PowerPoint Presentation Rasterizer"
set "short_name=pptrasterizer"
set "installer_name=setup_%short_name%.bat"

::end vars.cmd

::begin register.cmd
@echo off
for /f %%p in ('where powershell') do set "powershell=%%p"

for %%v in (Show.8 SlideShow.8 Show.12 SlideShow.12 ShowMacroEnabled.12 SlideShowMacroEnabled.12) do (
    echo Installing for PowerPoint.%%v
    reg add HKCU\Software\Classes\PowerPoint.%%v\shell\Rasterize\command /f /ve /d "\"%powershell%\" -executionpolicy bypass -file \"%~dp0ppt_rasterize.ps1\" \"^%%1\""
)
pause
::end register.cmd

::begin ppt_rasterize.ps1
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
        $slidesPath = "$env:temp\PhotoAlbumSlides"
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
    Convert-Presentation $pfilename $slideShowFileName
} catch {
    Write-Host $_.Exception.ToString()
    Write-Host "An error occurred. Please report this bug at"
    Write-Host "http://github.com/utapyngo/pptrasterizer/issues"
    cmd /c pause | Out-Null
}

::end ppt_rasterize.ps1

::begin ppt_rasterize.cmd
@echo off
if [%1]==[] (
    echo Usage: ppt_rasterize "Presentation.pptx"
    goto :eof
)

powershell -ExecutionPolicy Bypass "%~dp0ppt_rasterize.ps1" '%1' '%2'

::end ppt_rasterize.cmd

::begin uninstall.cmd
@echo off
setlocal EnableDelayedExpansion

call %~dp0vars.cmd

echo %product_name% is about to be removed from your computer.
pause

:: unregister the "Rasterize" command
powershell -executionpolicy bypass -file %~dp0unreg.ps1

:: remove from Programs and Features
reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\%short_name%" /f >nul

:: remove the files
set "files="
for /f %%f in (files.txt) do (
    set "files=!files! %%f"
)
for %%f in (!files!) do (
    :: skip this file
    if %%f neq %~nx0 (
        if exist "%~dp0\%%f" (
            echo Removing %%f
            del "%~dp0\%%f"
        )
    )
)

:: count files left
set cnt=0
for %%f in (%~dp0) do set /a cnt+=1

:: remove the directory if it is empty
:: remove uninstall.cmd otherwise
if !cnt!==1 (
    echo Removing %~dp0
    cd ..
    echo %product_name% has been uninstalled from your computer.
    rmdir %~dp0 /s /q
) else (
    echo %product_name% has been uninstalled from your computer.
    del %0
)

::end uninstall.cmd

::begin unreg.ps1
# Unregister the "Rasterize" command

$nothing = $true
Set-Location HKCU:\Software\Classes
'Show.8 SlideShow.8 Show.12 SlideShow.12 ShowMacroEnabled.12 SlideShowMacroEnabled.12'.Split() |% {
    if (Test-Path "PowerPoint.$_\shell\Rasterize") {
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

::end unreg.ps1

::begin update.cmd
@echo off

call %~dp0vars.cmd

set "version_url=http://utapyngo.github.com/pptrasterizer/version.txt"
set "installer_url=http://utapyngo.github.com/pptrasterizer/%installer_name%"
set "install_dir=%~dp0"

echo Checking for the latest version online...
powershell -command "(New-Object Net.WebClient).DownloadFile(\"%version_url%\", \"%TEMP%\%short_name%_latest_version.txt\")"
if errorlevel 1 (
    echo Unable to check version online. Check your Internet connection and firewall rules.
    pause
    goto :eof
)
set /p latest_version=<"%TEMP%\%short_name%_latest_version.txt"
set /p current_version=<"%install_dir%version.txt"
if %current_version% GEQ %latest_version% (
    echo You are already using the latest version of %product_name%.
    echo Press any key to exit . . .
    pause>nul
    goto :eof
)

echo Downloading the installer...
powershell -command "(New-Object Net.WebClient).DownloadFile(\"%installer_url%\", \"%TEMP%\%installer_name%\")"

if errorlevel 1 (
    echo Unable to download the installer. Check your Internet connection and firewall rules.
    pause
    goto :eof
)

call %TEMP%\%installer_name%

del %TEMP%\%installer_name%

::end update.cmd

::begin version.txt
34

::end version.txt
