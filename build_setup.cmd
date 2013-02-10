@echo off
set "outfile=setup_ppt_rasterize.bat"
set "files=register.cmd ppt_rasterize.ps1 ppt_rasterize.cmd uninstall.cmd unreg.ps1"

echo @echo off>%outfile%
echo.>>%outfile%
echo set "files=%files%">>%outfile%
echo.>>%outfile%
type install.cmd>>%outfile%

for %%f in (%files%) do (
    echo.>>%outfile%
    echo ::begin %%f>>%outfile%
    type %%f>>%outfile%
    echo.>>%outfile%
    echo ::end %%f>>%outfile%
)
