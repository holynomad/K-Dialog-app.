program KUMC_DIALOG_XE7;

uses
  Forms,
  VarCom,
  Vcl.Dialogs,
  MainDialog in 'MainDialog.pas' {MainDlg},
  DlgAbout in 'DlgAbout.pas' {About},
  DlgUser in 'DlgUser.pas' {SelUser},
  DlgPrint in 'DlgPrint.pas' {FtpPrint},
  ShellAPI,       // ShellExec ���� �߰� @ 2018.06.12 LSH
  Winapi.Windows, // SW_% ������� �߰� @ 2018.06.12 LSH
  Vcl.Themes,
  Vcl.Styles;

{$R *.RES}

begin


//   // �ʼ� ���� ���(32/64Bit) �ý��� ���� ���� @ 2018.06.12 LSH
//    --> EMRCOM210.bpl�� �ý����������� ���� loading �Ǵ� �̽������� ����.
//    --> KDialConfigConsole.exe ���� �������� ���� Library Path ���� ����/��ȯ @ 2018.07.18 LSH
//   ShellExecute(Application.Handle,
//                'open',
//                PCHAR('KUMC_DIALOG_CONFIG.bat'),
//                nil,
//                PCHAR(G_HOMEDIR + 'EXE\'),
//                SW_SHOWNORMAL);

   Application.Initialize;

   SetVarCom;

   Application.Title := 'KUMC ���̾�α�(XE7)';
   TStyleManager.TrySetStyle('Metropolis UI Blue');
   Application.CreateForm(TMainDlg, MainDlg);
   Application.Run;
end.
