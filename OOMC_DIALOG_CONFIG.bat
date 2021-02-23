@echo off

pushd "%~dp0"

rem ###########################################################################################
rem # OOMC ���̾�α�(XE7) ȯ�漳�� �� ����ȭ�� ��������� �ڵ����� ver 3.0
rem # 
rem # Author : Lee, Se-Ha
rem # Date   : 2017.12.13
rem # Special Thanks to : ����ĭ (Ÿ��������_��ġ�ϱ�.BAT���� ������..)
rem #
rem # (ver 2.0) Application ����� �ý��� ������ �ʼ����� ��� �ڵ����� �� Console ���� ����.
rem # (ver 3.0) �ý��� ������ Lib �ϰ����� ��� App.������ ���� Path ȯ�溯�� ���� ���� [M180718]
rem ###########################################################################################


rem �� LIB �� ȯ��/�ý��� ���� Path ����

set OOMCLIB=C:\OOMC\LIB\
set OOMCEXE=C:\OOMC\EXE\
set OOMCTEMP=C:\OOMC\TEMP\
set OOMCENV=C:\OOMC\ENV\
set OOMCDEVEXE=C:\OOMC_DEV\EXE\
set OOMCDEVTEMP=C:\OOMC_DEV\TEMP\
set OOMCDEVENV=C:\OOMC_DEV\ENV\
set OS32SYSPATH=C:\Windows\System32\
set OS64SYSPATH=C:\WINDOWS\SYSWOW64\



rem �� ������ �޾ƿ���
set CURR_DIR=%cd%\

echo �� ������ : %CURR_DIR%



rem �����ο� ���� ���丮 ���� 

if %CURR_DIR%==%OOMCEXE% (
	
	echo.
	echo [OOMCEXE] ��� ���� ����
	
	set NOWEXEPATH=C:\OOMC\EXE\
	
	echo.
	
	rem echo %OOMCEXE%	
	rem echo %NOWEXEPATH%
	
	echo TMAX_D0.ENV ���� ���� ���� �� �űԻ���
	
	C:
	cd\
	cd C:\OOMC\ENV\

	del TMAX_D0.ENV
	
	echo TMAX_D0.ENV App. �����ο� ���� �����
	
	echo [01]>> TMAX_D0.ENV
	echo TMAXDIR=C:\TMAX>> TMAX_D0.ENV
	echo TMAX_HOST_ADDR=���߼���IP>> TMAX_D0.ENV
	echo TMAX_HOST_PORT=���߼���PORT>> TMAX_D0.ENV
	echo TMAX_BACKUP_ADDR=���߼���IP>> TMAX_D0.ENV
	echo TMAX_BACKUP_PORT=���߼���PORT>> TMAX_D0.ENV
	echo FDLFILE=C:\OOMC\ENV\HIS.FDL>> TMAX_D0.ENV
	echo TMAX_CONNECT_TIMEOUT=^5>> TMAX_D0.ENV
	echo --TMAX_TRACE=*:ulog:dye>> TMAX_D0.ENV
	echo --TMAX_DEBUG=C:\DEV\TMAX\ULOG\TMAX.dbg>> TMAX_D0.ENV
	echo --ULOGPFX=C:\DEV\TMAX\ULOG\TMAX.ulog>> TMAX_D0.ENV
	
	echo.>> TMAX_D0.ENV
	echo [02]>> TMAX_D0.ENV
	echo TMAXDIR=C:\TMAX>> TMAX_D0.ENV
	echo TMAX_HOST_ADDR=���߼���IP>> TMAX_D0.ENV
	echo TMAX_HOST_PORT=���߼���PORT>> TMAX_D0.ENV
	echo TMAX_BACKUP_ADDR=���߼���IP>> TMAX_D0.ENV
	echo TMAX_BACKUP_PORT=���߼���PORT>> TMAX_D0.ENV
	echo FDLFILE=C:\OOMC\ENV\HIS.FDL>> TMAX_D0.ENV
	echo TMAX_CONNECT_TIMEOUT=^5>> TMAX_D0.ENV
	echo --TMAX_TRACE=*:ulog:dye>> TMAX_D0.ENV
	echo --TMAX_DEBUG=C:\DEV\TMAX\ULOG\TMAX.dbg>> TMAX_D0.ENV
	echo --ULOGPFX=C:\DEV\TMAX\ULOG\TMAX.ulog>> TMAX_D0.ENV
	
	echo.>> TMAX_D0.ENV
	echo [DR]>> TMAX_D0.ENV
	echo TMAXDIR=C:\TMAX>> TMAX_D0.ENV
	echo TMAX_HOST_ADDR=���߼���IP>> TMAX_D0.ENV
	echo TMAX_HOST_PORT=���߼���PORT>> TMAX_D0.ENV
	echo FDLFILE=C:\OOMC\ENV\HIS.FDL>> TMAX_D0.ENV
	echo TMAX_CONNECT_TIMEOUT=^5>> TMAX_D0.ENV
	echo --TMAX_TRACE=*:ulog:dye>> TMAX_D0.ENV
	echo --TMAX_DEBUG=C:\DEV\TMAX\ULOG\TMAX.dbg>> TMAX_D0.ENV
	echo --ULOGPFX=C:\DEV\TMAX\ULOG\TMAX.ulog>> TMAX_D0.ENV
	echo.>> TMAX_D0.ENV
	
	echo.
	
	cd\
	cd %OOMCEXE%	
	
	echo [batch ������ OOMC �����Ϸ�]
	echo.
	
	
	
)

if %CURR_DIR%==%OOMCDEVEXE% (

	echo.
	echo [OOMCDEVEXE] ��� ���� ����

	set NOWEXEPATH=C:\OOMC_DEV\EXE\
	
	echo.
	rem echo %OOMCDEVEXE%	
	rem echo %NOWEXEPATH%

	rem TMAX_D0.ENV ���� ���� ���� �� �űԻ���
	
	C:
	cd\
	cd C:\OOMC_DEV\ENV\

	del TMAX_D0.ENV
	
	rem TMAX_D0.ENV App. �����ο� ���� �����
	
	echo [01]>> TMAX_D0.ENV
	echo TMAXDIR=C:\TMAX>> TMAX_D0.ENV
	echo TMAX_HOST_ADDR=���߼���IP>> TMAX_D0.ENV
	echo TMAX_HOST_PORT=���߼���PORT>> TMAX_D0.ENV
	echo TMAX_BACKUP_ADDR=���߼���IP>> MW�ַ�Ǹ�_D0.ENV
	echo MW�ַ�Ǹ�_BACKUP_PORT=���߼���PORT>> MW�ַ�Ǹ�_D0.ENV
	echo FDLFILE=C:\OOMC_DEV\ENV\HIS.FDL>> MW�ַ�Ǹ�_D0.ENV
	echo MW�ַ�Ǹ�_CONNECT_TIMEOUT=^5>> MW�ַ�Ǹ�_D0.ENV
	echo --MW�ַ�Ǹ�_TRACE=*:ulog:dye>> MW�ַ�Ǹ�_D0.ENV
	echo --MW�ַ�Ǹ�_DEBUG=C:\DEV\MW�ַ�Ǹ�\ULOG\MW�ַ�Ǹ�.dbg>> MW�ַ�Ǹ�_D0.ENV
	echo --ULOGPFX=C:\DEV\MW�ַ�Ǹ�\ULOG\MW�ַ�Ǹ�.ulog>> MW�ַ�Ǹ�_D0.ENV
	
	echo.>> MW�ַ�Ǹ�_D0.ENV
	echo [02]>> MW�ַ�Ǹ�_D0.ENV
	echo MW�ַ�Ǹ�DIR=C:\MW�ַ�Ǹ�>> MW�ַ�Ǹ�_D0.ENV
	echo MW�ַ�Ǹ�_HOST_ADDR=���߼���IP>> MW�ַ�Ǹ�_D0.ENV
	echo MW�ַ�Ǹ�_HOST_PORT=���߼���PORT>> MW�ַ�Ǹ�_D0.ENV
	echo MW�ַ�Ǹ�_BACKUP_ADDR=���߼���IP>> MW�ַ�Ǹ�_D0.ENV
	echo MW�ַ�Ǹ�_BACKUP_PORT=���߼���PORT>> MW�ַ�Ǹ�_D0.ENV
	echo FDLFILE=C:\OOMC_DEV\ENV\HIS.FDL>> MW�ַ�Ǹ�_D0.ENV
	echo MW�ַ�Ǹ�_CONNECT_TIMEOUT=^5>> MW�ַ�Ǹ�_D0.ENV
	echo --MW�ַ�Ǹ�_TRACE=*:ulog:dye>> MW�ַ�Ǹ�_D0.ENV
	echo --MW�ַ�Ǹ�_DEBUG=C:\DEV\MW�ַ�Ǹ�\ULOG\MW�ַ�Ǹ�.dbg>> MW�ַ�Ǹ�_D0.ENV
	echo --ULOGPFX=C:\DEV\MW�ַ�Ǹ�\ULOG\MW�ַ�Ǹ�.ulog>> MW�ַ�Ǹ�_D0.ENV
	
	echo.>> MW�ַ�Ǹ�_D0.ENV
	echo [DR]>> MW�ַ�Ǹ�_D0.ENV
	echo MW�ַ�Ǹ�DIR=C:\MW�ַ�Ǹ�>> MW�ַ�Ǹ�_D0.ENV
	echo MW�ַ�Ǹ�_HOST_ADDR=���߼���IP>> MW�ַ�Ǹ�_D0.ENV
	echo MW�ַ�Ǹ�_HOST_PORT=���߼���PORT>> MW�ַ�Ǹ�_D0.ENV
	echo FDLFILE=C:\OOMC_DEV\ENV\HIS.FDL>> MW�ַ�Ǹ�_D0.ENV
	echo MW�ַ�Ǹ�_CONNECT_TIMEOUT=^5>> MW�ַ�Ǹ�_D0.ENV
	echo --MW�ַ�Ǹ�_TRACE=*:ulog:dye>> MW�ַ�Ǹ�_D0.ENV
	echo --MW�ַ�Ǹ�_DEBUG=C:\DEV\MW�ַ�Ǹ�\ULOG\MW�ַ�Ǹ�.dbg>> MW�ַ�Ǹ�_D0.ENV
	echo --ULOGPFX=C:\DEV\MW�ַ�Ǹ�\ULOG\MW�ַ�Ǹ�.ulog>> MW�ַ�Ǹ�_D0.ENV
	echo.>> MW�ַ�Ǹ�_D0.ENV

	echo.

	cd\
	cd %OOMCDEVEXE%		
	
	echo [batch ������ OOMC_DEV �����Ϸ�]
	echo.
	
	
	
)

rem �����ڱ���(�ý�������)���� Bat �����ؾ� �ϴ� ��� (���� ��� OOMC_DIALOG_XE7.exe�� C:\OOMC\EXE�� �����ؾ� ��!!��)

if %CURR_DIR%==%OS32SYSPATH% (

	echo.
	echo [OOMCEXE] ��� ���� ����

	set NOWEXEPATH=C:\OOMC\EXE\
	
	echo.
	
	rem echo %OOMCEXE%	
	rem echo %NOWEXEPATH%
	
	echo MW�ַ�Ǹ�_D0.ENV ���� ���� ���� �� �űԻ���

	del %OOMCENV%MW�ַ�Ǹ�_D0.ENV
	
	echo MW�ַ�Ǹ�_D0.ENV App. �����ο� ���� �����
	
	echo [01]>> %OOMCENV%MW�ַ�Ǹ�_D0.ENV
	echo MW�ַ�Ǹ�DIR=C:\MW�ַ�Ǹ�>> %OOMCENV%MW�ַ�Ǹ�_D0.ENV
	echo MW�ַ�Ǹ�_HOST_ADDR=���߼���IP>> %OOMCENV%MW�ַ�Ǹ�_D0.ENV
	echo MW�ַ�Ǹ�_HOST_PORT=���߼���PORT>> %OOMCENV%MW�ַ�Ǹ�_D0.ENV
	echo MW�ַ�Ǹ�_BACKUP_ADDR=���߼���IP>> %OOMCENV%MW�ַ�Ǹ�_D0.ENV
	echo MW�ַ�Ǹ�_BACKUP_PORT=���߼���PORT>> %OOMCENV%MW�ַ�Ǹ�_D0.ENV
	echo FDLFILE=C:\OOMC\ENV\HIS.FDL>> %OOMCENV%MW�ַ�Ǹ�_D0.ENV
	echo MW�ַ�Ǹ�_CONNECT_TIMEOUT=^5>> %OOMCENV%MW�ַ�Ǹ�_D0.ENV
	echo --MW�ַ�Ǹ�_TRACE=*:ulog:dye>> %OOMCENV%MW�ַ�Ǹ�_D0.ENV
	echo --MW�ַ�Ǹ�_DEBUG=C:\DEV\MW�ַ�Ǹ�\ULOG\MW�ַ�Ǹ�.dbg>> %OOMCENV%MW�ַ�Ǹ�_D0.ENV
	echo --ULOGPFX=C:\DEV\MW�ַ�Ǹ�\ULOG\MW�ַ�Ǹ�.ulog>> %OOMCENV%MW�ַ�Ǹ�_D0.ENV
	
	echo.>> %OOMCENV%MW�ַ�Ǹ�_D0.ENV
	echo [02]>> %OOMCENV%MW�ַ�Ǹ�_D0.ENV
	echo MW�ַ�Ǹ�DIR=C:\MW�ַ�Ǹ�>> %OOMCENV%MW�ַ�Ǹ�_D0.ENV
	echo MW�ַ�Ǹ�_HOST_ADDR=���߼���IP>> %OOMCENV%MW�ַ�Ǹ�_D0.ENV
	echo MW�ַ�Ǹ�_HOST_PORT=���߼���PORT>> %OOMCENV%MW�ַ�Ǹ�_D0.ENV
	echo MW�ַ�Ǹ�_BACKUP_ADDR=���߼���IP>> %OOMCENV%MW�ַ�Ǹ�_D0.ENV
	echo MW�ַ�Ǹ�_BACKUP_PORT=���߼���PORT>> %OOMCENV%MW�ַ�Ǹ�_D0.ENV
	echo FDLFILE=C:\OOMC\ENV\HIS.FDL>> %OOMCENV%MW�ַ�Ǹ�_D0.ENV
	echo MW�ַ�Ǹ�_CONNECT_TIMEOUT=^5>> %OOMCENV%MW�ַ�Ǹ�_D0.ENV
	echo --MW�ַ�Ǹ�_TRACE=*:ulog:dye>> %OOMCENV%MW�ַ�Ǹ�_D0.ENV
	echo --MW�ַ�Ǹ�_DEBUG=C:\DEV\MW�ַ�Ǹ�\ULOG\MW�ַ�Ǹ�.dbg>> %OOMCENV%MW�ַ�Ǹ�_D0.ENV
	echo --ULOGPFX=C:\DEV\MW�ַ�Ǹ�\ULOG\MW�ַ�Ǹ�.ulog>> %OOMCENV%MW�ַ�Ǹ�_D0.ENV
	
	echo.>> %OOMCENV%MW�ַ�Ǹ�_D0.ENV
	echo [DR]>> %OOMCENV%MW�ַ�Ǹ�_D0.ENV
	echo MW�ַ�Ǹ�DIR=C:\MW�ַ�Ǹ�>> %OOMCENV%MW�ַ�Ǹ�_D0.ENV
	echo MW�ַ�Ǹ�_HOST_ADDR=���߼���IP>> %OOMCENV%MW�ַ�Ǹ�_D0.ENV
	echo MW�ַ�Ǹ�_HOST_PORT=���߼���PORT>> %OOMCENV%MW�ַ�Ǹ�_D0.ENV
	echo FDLFILE=C:\OOMC\ENV\HIS.FDL>> %OOMCENV%MW�ַ�Ǹ�_D0.ENV
	echo MW�ַ�Ǹ�_CONNECT_TIMEOUT=^5>> %OOMCENV%MW�ַ�Ǹ�_D0.ENV
	echo --MW�ַ�Ǹ�_TRACE=*:ulog:dye>> %OOMCENV%MW�ַ�Ǹ�_D0.ENV
	echo --MW�ַ�Ǹ�_DEBUG=C:\DEV\MW�ַ�Ǹ�\ULOG\MW�ַ�Ǹ�.dbg>> %OOMCENV%MW�ַ�Ǹ�_D0.ENV
	echo --ULOGPFX=C:\DEV\MW�ַ�Ǹ�\ULOG\MW�ַ�Ǹ�.ulog>> %OOMCENV%MW�ַ�Ǹ�_D0.ENV
	echo.>> %OOMCENV%MW�ַ�Ǹ�_D0.ENV
	
	echo.	
	
	echo [batch ������ OOMC �����Ϸ�]
	echo.
	
)





rem OS�� 32��Ʈ���� 64��Ʈ���� Ȯ���� Root �������� Path ���� Console (KDialConfigConsole.exe) ȣ�� �� �ٷΰ��� ������ ����

reg Query "HKLM\Hardware\Description\System\CentralProcessor\0" | find /i "x86" > NUL && set OS=32BIT || set OS=64BIT



if %OS%==32BIT (
	
	del /q /f %OS32SYSPATH%*.bpl
	
	echo.
	
	rem 32��Ʈ OS ����ȭ�鿡 �ٷΰ��� ICON �ڵ����� �ϱ�

	%OS32SYSPATH%WindowsPowerShell\v1.0\powershell "$s=(New-Object -COM WScript.Shell).CreateShortcut('%userprofile%\DeskTop\OOMC ���̾�α�.lnk');$s.TargetPath='%NOWEXEPATH%KDialConfigConsole.exe';$s.Save()"

	echo [����ȭ�鿡 OOMC ���̾�α� ����Ű �����Ϸ�]
	echo.	
	
)



if %OS%==64BIT (
	
	del /q /f %OS64SYSPATH%*.bpl
	
	echo.
	
	rem 64��Ʈ OS ����ȭ�鿡 �ٷΰ��� ICON �ڵ����� �ϱ�

	%OS64SYSPATH%WindowsPowerShell\v1.0\powershell "$s=(New-Object -COM WScript.Shell).CreateShortcut('%userprofile%\DeskTop\OOMC ���̾�α�.lnk');$s.TargetPath='%NOWEXEPATH%KDialConfigConsole.exe';$s.Save()"

	echo [����ȭ�鿡 OOMC ���̾�α� ����Ű �����Ϸ�]
	echo.		
)



if exist %NOWEXEPATH%OOMC_DIALOG_XE7.exe (
	echo �ڡ� ��� ������ �Ϸ�Ǿ�����, ����ȭ�鿡 ������� [OOMC ���̾�α�] ���� :-p
	echo. 
) else (
	if %CURR_DIR%==%OOMCEXE% (
		echo �ڡ� C:\OOMC\EXE �����ȿ� [OOMC_DIALOG_XE7.exe] ���������� �����ϴ�. ���������� �ش� ����� ������ �����Ͻ� �� ����ϼ��� �ڡ�	
	)
	
	if %CURR_DIR%==%OOMCDEVEXE% (
		echo �ڡ� C:\OOMC_DEV\EXE �����ȿ� [OOMC_DIALOG_XE7.exe] ���������� �����ϴ�. ���������� �ش� ����� ������ �����Ͻ� �� ����ϼ��� �ڡ�
	)
	
	if %CURR_DIR%==%OS32SYSPATH% (
		echo �ڡ� C:\OOMC\EXE �����ȿ� [OOMC_DIALOG_XE7.exe] ���������� �����ϴ�. ���������� �ش� ����� ������ �����Ͻ� �� ����ϼ��� �ڡ�	
	)	
)

echo.
echo.

popd

pause 
