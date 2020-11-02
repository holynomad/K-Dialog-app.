program KDialConfigConsole;

{$APPTYPE CONSOLE}

{$R *.res}

uses
  System.SysUtils,
  Winapi.Windows,
  StringLib,   // Search Path에서 잡아줌(C:\KUMC_DEV\DCU)
  vcl.dialogs  // ShowMessage 디버깅용 추가
  ;

var
   sNewPath, sRootPath: String;
   sNowExecPath  : String;
   sParamStrings : String;

   startupinfo: TStartupInfo;
   processinfo: TProcessInformation;
begin
   try
      { TODO -oUser -cConsole Main : Insert code here }

      // CreateProcess 사용위한 Params.
      fillchar(startupinfo, sizeof(Tstartupinfo), 0);
      GetStartupInfo(startupinfo);

      // 현 App. 실행 Root 디렉토리
      sRootPath := UpperCase(TokenStr(ExtractFileDir(ParamStr(0)), '\', 1)) +
                   '\'                                                      +
                   UpperCase(TokenStr(ExtractFileDir(ParamStr(0)), '\', 2)) +
                   '\';

      //showmessage(sRootPath);


      // 아래 환경변수(Path) 잡아주는 부분은 Auto_LoadU.pas 참조 (헬퍼: 장은석)
      sNewPath  := GetEnvironmentVariable('Path');

      // 2016-03-04 장은석 , Path에 LIB 포함 안되어 있는 경우만 추가.
      // 2016-03-11 장은석 , Path에 이미 추가 되어있어서 더이상 안추가하는 경우
      //                     KUMC <> KUMC_LIB 상호 교차 실행시키는 경우 순서변동이
      //                     없기 때문에 제거해주고 재추가 하는 방식 추가.
      //                     기존에 추가되어 있는 LIB를 전부 제거해야 할 듯 -.-;;
      sNewPath := StringReplace(sNewPath, 'C:\KUMC_DEV\LIB;', '', [rfIgnoreCase]);
      sNewPath := StringReplace(sNewPath, 'C:\KUMC\LIB;', '', [rfIgnoreCase]);
      sNewPath := Format('%sLIB;', [sRootPath])+ sNewPath;

      // ★ 비개발자 User PC에 C:\TMAX\DLL 경로 강제 설정
      //sNewPath := 'C:\TMAX\DLL;' + sNewPath;

      //showmessage(sNewPath);

      // App. 실행경로 잡아주기 (추가)
      sNowExecPath := Format('%sEXE\', [sRootPath]);

      //showmessage(sNowExecPath);

      // Path 사용자(환경)변수 등록
      SetEnvironmentVariable('Path', PWideChar(sNewPath));

      // test
      //showmessage('KDialConfingConSole : SetEnvironmentVariable Finished');

      // App. 실행(txInit 등)위한 최소한의 Params. 설정
      sParamStrings := ParamStr(0)  + ' ' +     // 0
                       'sUSERID'    + ' ' +     // 1
                       'sUSERNM'    + ' ' +     // 2
                       'sDeptcd'    + ' ' +     // 3
                       'sDeptcd'    + ' ' +     // 4
                       'sGrade'     + ' ' +     // 5
                       'sGLocateCD' + ' ' +     // 6
                       'KUMC'       + ' ' +     // 7
                       'SDS'        + ' ' +     // 8
                       'AB1'        + ' ' +     // 9
                       'B01'        + ' ' +     // 10
                       'sExternalProgramCode' + ' ' + // 11
                       '_'          + ' ' +     // 12
                       'sTmpIP'     + ' ' +     // 13
                       'sExternalProgramCode' + ' ' + // 14
                       'sExternalProgram'     + ' ' + // 15
                       '01'         + ' ' +     // 16 : G_SVRID (1호기)
                       '9999'       + ' ' +     // 17
                       '12345'      + ' ' +     // 18
                       '98765'      + ' ' +     // 19
                       'sVendor'    + ' ' +     // 20
                       'D0'         + ' ' +     // 21
                       '개발'       + ' ' +     // 22
                       'sEMRFlag'   + ' ' +     // 23
                       'sDeveloper' + ' ' +     // 24
                       'sEMRTest'   + ' ' +     // 25
                       ''           + ' ' +     // 26
                       '';                      // 27


      // KUMC 다이얼로그(exe) 실행
      CreateProcess(
                     PChar(sNowExecPath + 'KUMC_DIALOG_XE7.exe'),
                     PChar(sParamStrings),
                     nil,
                     nil,
                     False,
                     NORMAL_PRIORITY_CLASS,
                     nil,
                     nil,
                     startupinfo,
                     processinfo
                   );

   // test
   //showmessage('KDialConfingConSole : CreateProcess Finished')

   except
      on E: Exception do
         showmessage(E.Message);
   end;
end.
