@echo off

pushd "%~dp0"

rem ###########################################################################################
rem # OOMC 다이얼로그(XE7) 환경설정 및 바탕화면 단축아이콘 자동생성 ver 3.0
rem # 
rem # Author : Lee, Se-Ha
rem # Date   : 2017.12.13
rem # Special Thanks to : 모히칸 (타병원접속_설치하기.BAT에서 영감을..)
rem #
rem # (ver 2.0) Application 실행시 시스템 폴더내 필수구동 모듈 자동복사 및 Console 종료 적용.
rem # (ver 3.0) 시스템 폴더내 Lib 일괄복사 대신 App.실행경로 동적 Path 환경변수 설정 적용 [M180718]
rem ###########################################################################################


rem 각 LIB 및 환경/시스템 폴더 Path 정의

set OOMCLIB=C:\OOMC\LIB\
set OOMCEXE=C:\OOMC\EXE\
set OOMCTEMP=C:\OOMC\TEMP\
set OOMCENV=C:\OOMC\ENV\
set OOMCDEVEXE=C:\OOMC_DEV\EXE\
set OOMCDEVTEMP=C:\OOMC_DEV\TEMP\
set OOMCDEVENV=C:\OOMC_DEV\ENV\
set OS32SYSPATH=C:\Windows\System32\
set OS64SYSPATH=C:\WINDOWS\SYSWOW64\



rem 현 실행경로 받아오기
set CURR_DIR=%cd%\

echo 현 실행경로 : %CURR_DIR%



rem 실행경로에 맞춰 디렉토리 변경 

if %CURR_DIR%==%OOMCEXE% (
	
	echo.
	echo [OOMCEXE] 경로 맵핑 시작
	
	set NOWEXEPATH=C:\OOMC\EXE\
	
	echo.
	
	rem echo %OOMCEXE%	
	rem echo %NOWEXEPATH%
	
	echo TMAX_D0.ENV 기존 파일 삭제 및 신규생성
	
	C:
	cd\
	cd C:\OOMC\ENV\

	del TMAX_D0.ENV
	
	echo TMAX_D0.ENV App. 실행경로에 맞춰 재생성
	
	echo [01]>> TMAX_D0.ENV
	echo TMAXDIR=C:\TMAX>> TMAX_D0.ENV
	echo TMAX_HOST_ADDR=개발서버IP>> TMAX_D0.ENV
	echo TMAX_HOST_PORT=개발서버PORT>> TMAX_D0.ENV
	echo TMAX_BACKUP_ADDR=개발서버IP>> TMAX_D0.ENV
	echo TMAX_BACKUP_PORT=개발서버PORT>> TMAX_D0.ENV
	echo FDLFILE=C:\OOMC\ENV\HIS.FDL>> TMAX_D0.ENV
	echo TMAX_CONNECT_TIMEOUT=^5>> TMAX_D0.ENV
	echo --TMAX_TRACE=*:ulog:dye>> TMAX_D0.ENV
	echo --TMAX_DEBUG=C:\DEV\TMAX\ULOG\TMAX.dbg>> TMAX_D0.ENV
	echo --ULOGPFX=C:\DEV\TMAX\ULOG\TMAX.ulog>> TMAX_D0.ENV
	
	echo.>> TMAX_D0.ENV
	echo [02]>> TMAX_D0.ENV
	echo TMAXDIR=C:\TMAX>> TMAX_D0.ENV
	echo TMAX_HOST_ADDR=개발서버IP>> TMAX_D0.ENV
	echo TMAX_HOST_PORT=개발서버PORT>> TMAX_D0.ENV
	echo TMAX_BACKUP_ADDR=개발서버IP>> TMAX_D0.ENV
	echo TMAX_BACKUP_PORT=개발서버PORT>> TMAX_D0.ENV
	echo FDLFILE=C:\OOMC\ENV\HIS.FDL>> TMAX_D0.ENV
	echo TMAX_CONNECT_TIMEOUT=^5>> TMAX_D0.ENV
	echo --TMAX_TRACE=*:ulog:dye>> TMAX_D0.ENV
	echo --TMAX_DEBUG=C:\DEV\TMAX\ULOG\TMAX.dbg>> TMAX_D0.ENV
	echo --ULOGPFX=C:\DEV\TMAX\ULOG\TMAX.ulog>> TMAX_D0.ENV
	
	echo.>> TMAX_D0.ENV
	echo [DR]>> TMAX_D0.ENV
	echo TMAXDIR=C:\TMAX>> TMAX_D0.ENV
	echo TMAX_HOST_ADDR=개발서버IP>> TMAX_D0.ENV
	echo TMAX_HOST_PORT=개발서버PORT>> TMAX_D0.ENV
	echo FDLFILE=C:\OOMC\ENV\HIS.FDL>> TMAX_D0.ENV
	echo TMAX_CONNECT_TIMEOUT=^5>> TMAX_D0.ENV
	echo --TMAX_TRACE=*:ulog:dye>> TMAX_D0.ENV
	echo --TMAX_DEBUG=C:\DEV\TMAX\ULOG\TMAX.dbg>> TMAX_D0.ENV
	echo --ULOGPFX=C:\DEV\TMAX\ULOG\TMAX.ulog>> TMAX_D0.ENV
	echo.>> TMAX_D0.ENV
	
	echo.
	
	cd\
	cd %OOMCEXE%	
	
	echo [batch 실행경로 OOMC 설정완료]
	echo.
	
	
	
)

if %CURR_DIR%==%OOMCDEVEXE% (

	echo.
	echo [OOMCDEVEXE] 경로 맵핑 시작

	set NOWEXEPATH=C:\OOMC_DEV\EXE\
	
	echo.
	rem echo %OOMCDEVEXE%	
	rem echo %NOWEXEPATH%

	rem TMAX_D0.ENV 기존 파일 삭제 및 신규생성
	
	C:
	cd\
	cd C:\OOMC_DEV\ENV\

	del TMAX_D0.ENV
	
	rem TMAX_D0.ENV App. 실행경로에 맞춰 재생성
	
	echo [01]>> TMAX_D0.ENV
	echo TMAXDIR=C:\TMAX>> TMAX_D0.ENV
	echo TMAX_HOST_ADDR=개발서버IP>> TMAX_D0.ENV
	echo TMAX_HOST_PORT=개발서버PORT>> TMAX_D0.ENV
	echo TMAX_BACKUP_ADDR=개발서버IP>> MW솔루션명_D0.ENV
	echo MW솔루션명_BACKUP_PORT=개발서버PORT>> MW솔루션명_D0.ENV
	echo FDLFILE=C:\OOMC_DEV\ENV\HIS.FDL>> MW솔루션명_D0.ENV
	echo MW솔루션명_CONNECT_TIMEOUT=^5>> MW솔루션명_D0.ENV
	echo --MW솔루션명_TRACE=*:ulog:dye>> MW솔루션명_D0.ENV
	echo --MW솔루션명_DEBUG=C:\DEV\MW솔루션명\ULOG\MW솔루션명.dbg>> MW솔루션명_D0.ENV
	echo --ULOGPFX=C:\DEV\MW솔루션명\ULOG\MW솔루션명.ulog>> MW솔루션명_D0.ENV
	
	echo.>> MW솔루션명_D0.ENV
	echo [02]>> MW솔루션명_D0.ENV
	echo MW솔루션명DIR=C:\MW솔루션명>> MW솔루션명_D0.ENV
	echo MW솔루션명_HOST_ADDR=개발서버IP>> MW솔루션명_D0.ENV
	echo MW솔루션명_HOST_PORT=개발서버PORT>> MW솔루션명_D0.ENV
	echo MW솔루션명_BACKUP_ADDR=개발서버IP>> MW솔루션명_D0.ENV
	echo MW솔루션명_BACKUP_PORT=개발서버PORT>> MW솔루션명_D0.ENV
	echo FDLFILE=C:\OOMC_DEV\ENV\HIS.FDL>> MW솔루션명_D0.ENV
	echo MW솔루션명_CONNECT_TIMEOUT=^5>> MW솔루션명_D0.ENV
	echo --MW솔루션명_TRACE=*:ulog:dye>> MW솔루션명_D0.ENV
	echo --MW솔루션명_DEBUG=C:\DEV\MW솔루션명\ULOG\MW솔루션명.dbg>> MW솔루션명_D0.ENV
	echo --ULOGPFX=C:\DEV\MW솔루션명\ULOG\MW솔루션명.ulog>> MW솔루션명_D0.ENV
	
	echo.>> MW솔루션명_D0.ENV
	echo [DR]>> MW솔루션명_D0.ENV
	echo MW솔루션명DIR=C:\MW솔루션명>> MW솔루션명_D0.ENV
	echo MW솔루션명_HOST_ADDR=개발서버IP>> MW솔루션명_D0.ENV
	echo MW솔루션명_HOST_PORT=개발서버PORT>> MW솔루션명_D0.ENV
	echo FDLFILE=C:\OOMC_DEV\ENV\HIS.FDL>> MW솔루션명_D0.ENV
	echo MW솔루션명_CONNECT_TIMEOUT=^5>> MW솔루션명_D0.ENV
	echo --MW솔루션명_TRACE=*:ulog:dye>> MW솔루션명_D0.ENV
	echo --MW솔루션명_DEBUG=C:\DEV\MW솔루션명\ULOG\MW솔루션명.dbg>> MW솔루션명_D0.ENV
	echo --ULOGPFX=C:\DEV\MW솔루션명\ULOG\MW솔루션명.ulog>> MW솔루션명_D0.ENV
	echo.>> MW솔루션명_D0.ENV

	echo.

	cd\
	cd %OOMCDEVEXE%		
	
	echo [batch 실행경로 OOMC_DEV 설정완료]
	echo.
	
	
	
)

rem 관리자권한(시스템폴더)으로 Bat 실행해야 하는 경우 (★이 경우 OOMC_DIALOG_XE7.exe는 C:\OOMC\EXE에 존재해야 함!!★)

if %CURR_DIR%==%OS32SYSPATH% (

	echo.
	echo [OOMCEXE] 경로 맵핑 시작

	set NOWEXEPATH=C:\OOMC\EXE\
	
	echo.
	
	rem echo %OOMCEXE%	
	rem echo %NOWEXEPATH%
	
	echo MW솔루션명_D0.ENV 기존 파일 삭제 및 신규생성

	del %OOMCENV%MW솔루션명_D0.ENV
	
	echo MW솔루션명_D0.ENV App. 실행경로에 맞춰 재생성
	
	echo [01]>> %OOMCENV%MW솔루션명_D0.ENV
	echo MW솔루션명DIR=C:\MW솔루션명>> %OOMCENV%MW솔루션명_D0.ENV
	echo MW솔루션명_HOST_ADDR=개발서버IP>> %OOMCENV%MW솔루션명_D0.ENV
	echo MW솔루션명_HOST_PORT=개발서버PORT>> %OOMCENV%MW솔루션명_D0.ENV
	echo MW솔루션명_BACKUP_ADDR=개발서버IP>> %OOMCENV%MW솔루션명_D0.ENV
	echo MW솔루션명_BACKUP_PORT=개발서버PORT>> %OOMCENV%MW솔루션명_D0.ENV
	echo FDLFILE=C:\OOMC\ENV\HIS.FDL>> %OOMCENV%MW솔루션명_D0.ENV
	echo MW솔루션명_CONNECT_TIMEOUT=^5>> %OOMCENV%MW솔루션명_D0.ENV
	echo --MW솔루션명_TRACE=*:ulog:dye>> %OOMCENV%MW솔루션명_D0.ENV
	echo --MW솔루션명_DEBUG=C:\DEV\MW솔루션명\ULOG\MW솔루션명.dbg>> %OOMCENV%MW솔루션명_D0.ENV
	echo --ULOGPFX=C:\DEV\MW솔루션명\ULOG\MW솔루션명.ulog>> %OOMCENV%MW솔루션명_D0.ENV
	
	echo.>> %OOMCENV%MW솔루션명_D0.ENV
	echo [02]>> %OOMCENV%MW솔루션명_D0.ENV
	echo MW솔루션명DIR=C:\MW솔루션명>> %OOMCENV%MW솔루션명_D0.ENV
	echo MW솔루션명_HOST_ADDR=개발서버IP>> %OOMCENV%MW솔루션명_D0.ENV
	echo MW솔루션명_HOST_PORT=개발서버PORT>> %OOMCENV%MW솔루션명_D0.ENV
	echo MW솔루션명_BACKUP_ADDR=개발서버IP>> %OOMCENV%MW솔루션명_D0.ENV
	echo MW솔루션명_BACKUP_PORT=개발서버PORT>> %OOMCENV%MW솔루션명_D0.ENV
	echo FDLFILE=C:\OOMC\ENV\HIS.FDL>> %OOMCENV%MW솔루션명_D0.ENV
	echo MW솔루션명_CONNECT_TIMEOUT=^5>> %OOMCENV%MW솔루션명_D0.ENV
	echo --MW솔루션명_TRACE=*:ulog:dye>> %OOMCENV%MW솔루션명_D0.ENV
	echo --MW솔루션명_DEBUG=C:\DEV\MW솔루션명\ULOG\MW솔루션명.dbg>> %OOMCENV%MW솔루션명_D0.ENV
	echo --ULOGPFX=C:\DEV\MW솔루션명\ULOG\MW솔루션명.ulog>> %OOMCENV%MW솔루션명_D0.ENV
	
	echo.>> %OOMCENV%MW솔루션명_D0.ENV
	echo [DR]>> %OOMCENV%MW솔루션명_D0.ENV
	echo MW솔루션명DIR=C:\MW솔루션명>> %OOMCENV%MW솔루션명_D0.ENV
	echo MW솔루션명_HOST_ADDR=개발서버IP>> %OOMCENV%MW솔루션명_D0.ENV
	echo MW솔루션명_HOST_PORT=개발서버PORT>> %OOMCENV%MW솔루션명_D0.ENV
	echo FDLFILE=C:\OOMC\ENV\HIS.FDL>> %OOMCENV%MW솔루션명_D0.ENV
	echo MW솔루션명_CONNECT_TIMEOUT=^5>> %OOMCENV%MW솔루션명_D0.ENV
	echo --MW솔루션명_TRACE=*:ulog:dye>> %OOMCENV%MW솔루션명_D0.ENV
	echo --MW솔루션명_DEBUG=C:\DEV\MW솔루션명\ULOG\MW솔루션명.dbg>> %OOMCENV%MW솔루션명_D0.ENV
	echo --ULOGPFX=C:\DEV\MW솔루션명\ULOG\MW솔루션명.ulog>> %OOMCENV%MW솔루션명_D0.ENV
	echo.>> %OOMCENV%MW솔루션명_D0.ENV
	
	echo.	
	
	echo [batch 실행경로 OOMC 설정완료]
	echo.
	
)





rem OS가 32비트인지 64비트인지 확인후 Root 실행경로의 Path 설정 Console (KDialConfigConsole.exe) 호출 및 바로가기 아이콘 생성

reg Query "HKLM\Hardware\Description\System\CentralProcessor\0" | find /i "x86" > NUL && set OS=32BIT || set OS=64BIT



if %OS%==32BIT (
	
	del /q /f %OS32SYSPATH%*.bpl
	
	echo.
	
	rem 32비트 OS 바탕화면에 바로가기 ICON 자동생성 하기

	%OS32SYSPATH%WindowsPowerShell\v1.0\powershell "$s=(New-Object -COM WScript.Shell).CreateShortcut('%userprofile%\DeskTop\OOMC 다이얼로그.lnk');$s.TargetPath='%NOWEXEPATH%KDialConfigConsole.exe';$s.Save()"

	echo [바탕화면에 OOMC 다이얼로그 단축키 생성완료]
	echo.	
	
)



if %OS%==64BIT (
	
	del /q /f %OS64SYSPATH%*.bpl
	
	echo.
	
	rem 64비트 OS 바탕화면에 바로가기 ICON 자동생성 하기

	%OS64SYSPATH%WindowsPowerShell\v1.0\powershell "$s=(New-Object -COM WScript.Shell).CreateShortcut('%userprofile%\DeskTop\OOMC 다이얼로그.lnk');$s.TargetPath='%NOWEXEPATH%KDialConfigConsole.exe';$s.Save()"

	echo [바탕화면에 OOMC 다이얼로그 단축키 생성완료]
	echo.		
)



if exist %NOWEXEPATH%OOMC_DIALOG_XE7.exe (
	echo ★★ 모든 설정이 완료되었으니, 바탕화면에 만들어진 [OOMC 다이얼로그] 고고싱 :-p
	echo. 
) else (
	if %CURR_DIR%==%OOMCEXE% (
		echo ★★ C:\OOMC\EXE 폴더안에 [OOMC_DIALOG_XE7.exe] 실행파일이 없습니다. 실행파일을 해당 경로의 폴더에 복사하신 후 사용하세요 ★★	
	)
	
	if %CURR_DIR%==%OOMCDEVEXE% (
		echo ★★ C:\OOMC_DEV\EXE 폴더안에 [OOMC_DIALOG_XE7.exe] 실행파일이 없습니다. 실행파일을 해당 경로의 폴더에 복사하신 후 사용하세요 ★★
	)
	
	if %CURR_DIR%==%OS32SYSPATH% (
		echo ★★ C:\OOMC\EXE 폴더안에 [OOMC_DIALOG_XE7.exe] 실행파일이 없습니다. 실행파일을 해당 경로의 폴더에 복사하신 후 사용하세요 ★★	
	)	
)

echo.
echo.

popd

pause 
