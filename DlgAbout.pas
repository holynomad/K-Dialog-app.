{===============================================================================
   Program ID    : DlgAbout.pas
   Program Name  : 고대의료원 정보전산팀 비상연락망 프로그램 정보
   Program Desc. : 프로그램 version 및 개발자 정보.

   Author        : Lee, Se-Ha
   Date          : 2013.08.21
===============================================================================}
unit DlgAbout;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, 
  ExtCtrls, StdCtrls, shellapi, ComCtrls, Variants;

type
  TAbout = class(TForm)
    Label3: TLabel;
    lb_Version: TLabel;
    Image1: TImage;
    Label2: TLabel;
    Button1: TButton;
    Label5: TLabel;
    Label6: TLabel;
    procedure Label3Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  About: TAbout;

implementation

{$R *.DFM}

procedure TAbout.Label3Click(Sender: TObject);
begin
  ShellExecute(Application.Handle,'open','http://www.tmssoftware.com', nil, nil, SW_NORMAL);
end;

end.
