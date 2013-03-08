@echo off

call %~dp0vars.cmd

set "files=vars.cmd register.cmd ppt_rasterize.ps1 ppt_rasterize.cmd uninstall.cmd unreg.ps1 update.cmd version.txt"

for /f %%v in ('git rev-list HEAD --count') do set "version=%%v"
echo %version%>version.txt

echo @echo off>%installer_name%
echo.>>%installer_name%
echo set "files=%files%">>%installer_name%
echo.>>%installer_name%
type vars.cmd>>%installer_name%
echo set "version=%version%">>%installer_name%
echo.>>%installer_name%
type install.cmd>>%installer_name%

for %%f in (%files%) do (
    echo.>>%installer_name%
    echo ::begin %%f>>%installer_name%
    type %%f>>%installer_name%
    echo.>>%installer_name%
    echo ::end %%f>>%installer_name%
)
