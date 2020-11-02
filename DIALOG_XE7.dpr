program KUMC_DIALOG_XE7;

uses
  Forms,
  VarCom,
  Vcl.Dialogs,
  MainDialog in 'MainDialog.pas' {MainDlg},
  DlgAbout in 'DlgAbout.pas' {About},
  DlgUser in 'DlgUser.pas' {SelUser},
  DlgPrint in 'DlgPrint.pas' {FtpPrint},
  ShellAPI,       // ShellExec 위해 추가 @ 2018.06.12 LSH
  Winapi.Windows, // SW_% 사용위해 추가 @ 2018.06.12 LSH
  Vcl.Themes,
  Vcl.Styles;

{$R *.RES}

begin


//   // 필수 구동 모듈(32/64Bit) 시스템 폴더 복사 @ 2018.06.12 LSH
//    --> EMRCOM210.bpl이 시스템폴더에서 먼저 loading 되는 이슈때문에 중지.
//    --> KDialConfigConsole.exe 통해 동적으로 실행 Library Path 설정 적용/전환 @ 2018.07.18 LSH
//   ShellExecute(Application.Handle,
//                'open',
//                PCHAR('KUMC_DIALOG_CONFIG.bat'),
//                nil,
//                PCHAR(G_HOMEDIR + 'EXE\'),
//                SW_SHOWNORMAL);

   Application.Initialize;

   SetVarCom;

   Application.Title := 'KUMC 다이얼로그(XE7)';
   TStyleManager.TrySetStyle('Metropolis UI Blue');
   Application.CreateForm(TMainDlg, MainDlg);
   Application.Run;
end.
