set "product_name=PowerPoint Presentation Rasterizer"
set "short_name=pptrasterizer"
set "installer_name=setup_%short_name%.bat"
set "install_dir=%~dp0"
if exist "%install_dir%version.txt" (
    set /p version=<"%install_dir%version.txt"
)
