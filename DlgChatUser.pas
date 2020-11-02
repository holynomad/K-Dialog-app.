unit DlgChatUser;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, 
  Grids, BaseGrid, AdvGrid, StdCtrls, Variants, AdvObj;

type
  TSelChatUser = class(TForm)
    lb_UserCnt: TLabel;
    asg_ChatList: TAdvStringGrid;
    procedure asg_ChatListGetAlignment(Sender: TObject; ARow,
      ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure FormDestroy(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }



    //---------------------------------------------------------------------------
    // [함수] TCP Port 오픈여부 Check
    //    - 2013.08.30 LSH
    //    - 소스출처 : Google
    //---------------------------------------------------------------------------
    function PortTCPIsOpen(dwPort : Word; ipAddressStr:string) : boolean;



    //---------------------------------------------------------------------------
    // String 내의 특정 문자 삭제
    //    - 소스출처 : MComFunc.pas
    //---------------------------------------------------------------------------
    function DelChar( const Str : String ; DelC : Char ) : String;



    procedure GetChatUser;



  public
    { Public declarations }
    UserCount : Integer;


  end;



var
  SelChatUser: TSelChatUser;

implementation

{$R *.DFM}

uses
  CComFunc,
   Winsock, // Windows Sockets API Unit
   TpSvc,
   MainDialog;


procedure TSelChatUser.asg_ChatListGetAlignment(Sender: TObject; ARow,
  ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
   // Horizontal Alignment
   HAlign := taCenter;

   // Vertical Alignment
   VAlign := vtaCenter;
end;

procedure TSelChatUser.FormDestroy(Sender: TObject);
begin
   SelChatUser := nil;
end;

procedure TSelChatUser.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
   Action := caFree;
end;

procedure TSelChatUser.FormShow(Sender: TObject);
var
   i : Integer;
   iTargetServerPort  : Integer;
   sTargetServerIp    : String;
begin


   GetChatUser;


end;



procedure TSelChatUser.GetChatUser;
var
   i, iRowCnt : Integer;
   iTargetServerPort  : Integer;
   sTargetServerIp    : String;
   TpGetSearch : TTpSvc;
begin


//-----------------------------------------------------------------
   // 2. 조회
   //-----------------------------------------------------------------
   Screen.Cursor := crHourGlass;





   //-----------------------------------------------------------------
   // 2-1. Load Data by TpSvc
   //-----------------------------------------------------------------
   TpGetSearch := TTpSvc.Create;
   TpGetSearch.Init(Self);


   Screen.Cursor := crHourglass;


   try
      TpGetSearch.CountField  := 'S_CODE1';
      TpGetSearch.ShowMsgFlag := False;

      if TpGetSearch.GetSvc('MD_KUMCM_S1',
                          ['S_TYPE1'  , '1'
                          ],
                          [
                           'S_CODE1'  , 'sLocate'
                         , 'S_CODE2'  , 'sUserId'
                         , 'S_CODE3'  , 'sUserNm'
                         , 'S_CODE4'  , 'sDutyPart'
                         , 'S_CODE5'  , 'sDutySpec'
                         , 'S_CODE6'  , 'sDutyRmk'
                         , 'S_CODE7'  , 'sMobile'
                         , 'S_CODE8'  , 'sEmail'
                         , 'S_CODE9'  , 'sUserIp'
                         , 'S_CODE10' , 'sDutyUser'
                         , 'S_CODE11' , 'sDutyPtnr'
                         , 'S_CODE12' , 'sCallNo'
                         , 'S_CODE13' , 'sDelDate'
                         , 'S_CODE14' , 'sDeptCd'
                         , 'S_CODE15' , 'sFlag'
                         , 'S_CODE16' , 'sEditIp'
                         , 'S_CODE17' , 'sEditDate'
                         , 'S_CODE18' , 'sDocList'
                         , 'S_CODE19' , 'sDocYear'
                         , 'S_CODE20' , 'sDocSeq'
                         , 'S_CODE21' , 'sDocTitle'
                         , 'S_CODE22' , 'sRegDate'
                         , 'S_CODE23' , 'sRegUser'
                         , 'S_CODE24' , 'sRelDept'
                         , 'S_CODE25' , 'sDocRmk'
                         , 'S_CODE26' , 'sDocLoc'
                         , 'S_STRING1', 'sBoardSeq'
                         , 'S_STRING2', 'sCateUp'
                         , 'S_STRING3', 'sCateDown'
                         , 'S_STRING4', 'sContext'
                         , 'S_STRING5', 'sAttachNm'
                         , 'S_STRING6', 'sHideFile'
                         , 'S_STRING7', 'sServerIp'
                         , 'S_STRING8', 'sHeadTail'
                         , 'S_STRING9', 'sHeadSeq'
                         , 'S_STRING10','sTailSeq'
                         , 'S_STRING11','sReplyCnt'
                         , 'S_STRING12','sLikeCnt'
                         , 'S_STRING14','sNickNm'
                          ]) then

         if TpGetSearch.RowCount < 0 then
         begin
            Exit;
         end
         else if TpGetSearch.RowCount = 0 then
         begin

            Exit;
         end;






      with asg_ChatList do
      begin

         iRowCnt  := TpGetSearch.RowCount;
         RowCount := iRowCnt + FixedRows + 1;


         for i := 0 to iRowCnt - 1 do
         begin
            Cells[0,  i+FixedRows] := TpGetSearch.GetOutputDataS('sUserNm'    , i);
            Cells[1,  i+FixedRows] := TpGetSearch.GetOutputDataS('sUserIp'    , i);


            // 접속 서버의 IP 및 Port 정보
            if (Cells[1, i+FixedRows] <> '') then
            begin
               sTargetServerIp    := Cells[1, i+FixedRows];
               iTargetServerPort  := StrToInt(DelChar(CopyByte(Cells[1, i+FixedRows], LengthByte(Cells[1, i+FixedRows])-3, 4), '.'));


               //Colors[C_M_USERNM, i+FixedRows] := clGreen

               //---------------------------------------------------------
               // 접속하려는 Server Port 사용가능 여부 Check
               //---------------------------------------------------------
               if (PortTCPIsOpen(iTargetServerPort, sTargetServerIp)) then
                  Colors[0, i+FixedRows] := clGreen
               else
                  Colors[0, i+FixedRows] := clSilver;

            end
            else
               Colors[0, i+FixedRows] := clSilver;
         end;



         {

            // 접속 서버의 IP 및 Port 정보
            if (Cells[1, i+FixedRows] <> '') then
            begin
               sTargetServerIp    := Cells[1, i+FixedRows];
               iTargetServerPort  := StrToInt(DelChar(CopyByte(Cells[1, i+FixedRows], LengthByte(Cells[1, i+FixedRows])-3, 4), '.'));



               //---------------------------------------------------------
               // 접속하려는 Server Port 사용가능 여부 Check
               //---------------------------------------------------------
               if (PortTCPIsOpen(iTargetServerPort, sTargetServerIp)) then
                  Colors[0, i+FixedRows] := clGreen
               else
                  Colors[0, i+FixedRows] := clSilver;

            end
            else


               Colors[0, i+FixedRows] := clSilver;
         end;
         }
      end;

   finally
      FreeAndNil(TpGetSearch);
      Screen.Cursor := crDefault;
   end;


end;



function TSelChatUser.PortTCPIsOpen(dwPort : Word; ipAddressStr:string) : boolean;
var
   client : sockaddr_in;                                                // sockaddr_in is used by Windows Sockets to specify a local or remote endpoint address
   sock   : Integer;
begin
   client.sin_family       := AF_INET;
   client.sin_port         := htons(dwPort);                            // htons converts a u_short from host to TCP/IP network byte order.
   client.sin_addr.s_addr  := inet_addr(PAnsiChar(ipAddressStr));           // the inet_addr function converts a string containing an IPv4 dotted-decimal address into a proper address for the IN_ADDR structure.
   sock                    := socket(AF_INET, SOCK_STREAM, 0);          // The socket function creates a socket
   Result                  := connect(sock,client,SizeOf(client)) = 0;  // establishes a connection to a specified socket.
end;




//------------------------------------------------------------------------------
// String 내의 특정 문자 삭제
//       - 소스 출처 : MComFunc.pas
//
// Date   : 2013.08.29
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
function TSelChatUser.DelChar( const Str : String ; DelC : Char ) : String;
var
  I, L : Integer;
begin
  L := LengthByte(Str);
  result := '';

    for I := 1 to L  do begin
      if DelC <> '' then
      begin
        if Str[I] <> DelC then
          result := result + Str[I];
      end
      else
      begin
        if Str[I] in ['0'..'9'] then
          result := result + Str[I];
      end;
    end;
end;



end.
