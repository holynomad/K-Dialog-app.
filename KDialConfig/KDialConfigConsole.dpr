program KDialConfigConsole;

{$APPTYPE CONSOLE}

{$R *.res}

uses
  System.SysUtils,
  Winapi.Windows,
  StringLib,   // Search Path���� �����(C:\KUMC_DEV\DCU)
  vcl.dialogs  // ShowMessage ������ �߰�
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

      // CreateProcess ������� Params.
      fillchar(startupinfo, sizeof(Tstartupinfo), 0);
      GetStartupInfo(startupinfo);

      // �� App. ���� Root ���丮
      sRootPath := UpperCase(TokenStr(ExtractFileDir(ParamStr(0)), '\', 1)) +
                   '\'                                                      +
                   UpperCase(TokenStr(ExtractFileDir(ParamStr(0)), '\', 2)) +
                   '\';

      //showmessage(sRootPath);


      // �Ʒ� ȯ�溯��(Path) ����ִ� �κ��� Auto_LoadU.pas ���� (����: ������)
      sNewPath  := GetEnvironmentVariable('Path');

      // 2016-03-04 ������ , Path�� LIB ���� �ȵǾ� �ִ� ��츸 �߰�.
      // 2016-03-11 ������ , Path�� �̹� �߰� �Ǿ��־ ���̻� ���߰��ϴ� ���
      //                     KUMC <> KUMC_LIB ��ȣ ���� �����Ű�� ��� ����������
      //                     ���� ������ �������ְ� ���߰� �ϴ� ��� �߰�.
      //                     ������ �߰��Ǿ� �ִ� LIB�� ���� �����ؾ� �� �� -.-;;
      sNewPath := StringReplace(sNewPath, 'C:\KUMC_DEV\LIB;', '', [rfIgnoreCase]);
      sNewPath := StringReplace(sNewPath, 'C:\KUMC\LIB;', '', [rfIgnoreCase]);
      sNewPath := Format('%sLIB;', [sRootPath])+ sNewPath;

      // �� �񰳹��� User PC�� C:\TMAX\DLL ��� ���� ����
      //sNewPath := 'C:\TMAX\DLL;' + sNewPath;

      //showmessage(sNewPath);

      // App. ������ ����ֱ� (�߰�)
      sNowExecPath := Format('%sEXE\', [sRootPath]);

      //showmessage(sNowExecPath);

      // Path �����(ȯ��)���� ���
      SetEnvironmentVariable('Path', PWideChar(sNewPath));

      // test
      //showmessage('KDialConfingConSole : SetEnvironmentVariable Finished');

      // App. ����(txInit ��)���� �ּ����� Params. ����
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
                       '01'         + ' ' +     // 16 : G_SVRID (1ȣ��)
                       '9999'       + ' ' +     // 17
                       '12345'      + ' ' +     // 18
                       '98765'      + ' ' +     // 19
                       'sVendor'    + ' ' +     // 20
                       'D0'         + ' ' +     // 21
                       '����'       + ' ' +     // 22
                       'sEMRFlag'   + ' ' +     // 23
                       'sDeveloper' + ' ' +     // 24
                       'sEMRTest'   + ' ' +     // 25
                       ''           + ' ' +     // 26
                       '';                      // 27


      // KUMC ���̾�α�(exe) ����
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
