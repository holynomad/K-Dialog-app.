{===============================================================================
   Program ID    : DlgUser.pas
   Program Name  : 다이얼로그 단일 IP 다중 User 선택
   Program Desc. : 동일 IP를 공유하는 다중 User (예: OP 선생님..)인 경우,
                   세부 담당자를 선택할 수 있도록 도와줌.

   Author        : Lee, Se-Ha
   Date          : 2013.08.29
===============================================================================}
unit DlgUser;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, 
  Grids, BaseGrid, AdvGrid, StdCtrls, Variants, AdvObj;

type
  TSelUser = class(TForm)
    asg_UserList: TAdvStringGrid;
    lb_UserCnt: TLabel;
    procedure asg_UserListDblClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure asg_UserListGetAlignment(Sender: TObject; ARow,
      ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  SelUser: TSelUser;

implementation

uses
   MainDialog;

{$R *.DFM}

procedure TSelUser.asg_UserListDblClick(Sender: TObject);
begin
   //-----------------------------------------------------------
   // 선택한 사용자 정보를 MainDlg로 CallBack 처리
   //-----------------------------------------------------------
   MainDlg.GetMultiUser(asg_UserList.Cells[0, asg_UserList.Row],
                        asg_UserList.Cells[2, asg_UserList.Row],
                        asg_UserList.Cells[3, asg_UserList.Row],
                        asg_UserList.Cells[4, asg_UserList.Row],
                        asg_UserList.Cells[5, asg_UserList.Row],
                        asg_UserList.Cells[6, asg_UserList.Row],
                        asg_UserList.Cells[7, asg_UserList.Row]                                                                        
                        );

   // 선택화면 종료.
   Close;
end;

procedure TSelUser.FormDestroy(Sender: TObject);
begin
   SelUser := nil;
end;

procedure TSelUser.asg_UserListGetAlignment(Sender: TObject; ARow,
  ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
   // Horizontal Alignment
   HAlign := taCenter;

   // Vertical Alignment
   VAlign := vtaCenter;
end;

end.

