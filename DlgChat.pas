{===============================================================================
   Program ID    : DlgChat.pas
   Program Name  : 채팅(Chatting) 클라이언트
   Program Desc. : 다이얼로그 담당자정보에 IP가 저장된 상호 User간 1:1 TCP 기반
                   통신을 지원하는 Client.

   Author        : Lee, Se-Ha
   Date          : 2013.08.29
===============================================================================}
unit DlgChat;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, 
  Menus, ScktComp, ComCtrls, StdCtrls, ExtCtrls, Grids, BaseGrid, AdvGrid, 
  ExtDlgs, Variants, AdvObj;


const
   // Chat Send Columns
   C_SD_TEXT      = 0;
   C_SD_BUTTON    = 1;


   // Chat 박스 Columns
   C_CH_MYCOL     = 0;
   C_CH_TEXT      = 1;
   C_CH_YOURCOL   = 2;
   C_CH_HIDDEN    = 3;

type
  TClient = class(TForm)
    Bevel1: TBevel;
    Bevel2: TBevel;
    MemoChat: TMemo;
    ButtonConnect: TButton;
    StatusBar: TStatusBar;
    ButtonChangeNick: TButton;
    ClientSocket: TClientSocket;
    EditNick: TEdit;
    EditIp: TEdit;
    EditPort: TEdit;
    asg_Chat: TAdvStringGrid;
    EditSay: TEdit;
    ButtonSend: TButton;
    asg_ChatSend: TAdvStringGrid;
    pm_Chat: TPopupMenu;
    mi_Emoti: TMenuItem;
    OpenPictureDialog1: TOpenPictureDialog;
    Image1: TImage;
    mi_FileSend: TMenuItem;
    od_File: TOpenDialog;
    procedure ButtonSendClick(Sender: TObject);
    procedure ConnectDisconnect1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure ClientSocketConnect(Sender: TObject;
      Socket: TCustomWinSocket);
    procedure ClientSocketDisconnect(Sender: TObject;
      Socket: TCustomWinSocket);
    procedure ClientSocketRead(Sender: TObject; Socket: TCustomWinSocket);
    procedure ButtonConnectClick(Sender: TObject);
    procedure EditSayKeyPress(Sender: TObject; var Key: Char);
    procedure FormDestroy(Sender: TObject);
    procedure asg_ChatSendButtonClick(Sender: TObject; ACol,
      ARow: Integer);
    procedure asg_ChatSendKeyPress(Sender: TObject; var Key: Char);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure mi_EmotiClick(Sender: TObject);
    procedure ClientSocketError(Sender: TObject; Socket: TCustomWinSocket;
      ErrorEvent: TErrorEvent; var ErrorCode: Integer);
    procedure mi_FileSendClick(Sender: TObject);
    procedure asg_ChatButtonClick(Sender: TObject; ACol, ARow: Integer);
  private
    { Private declarations }
    Reciving     : Boolean;
    DataSize     : integer;
    Data         : TMemoryStream;
    iChatRowCnt  : Integer;
    IsEmotiSend  : Boolean;
    IsFileSend   : Boolean;


   //---------------------------------------------------------------------------
   // FTP 업로드
   //    - 소스출처 : MDN191F2.pas
   //---------------------------------------------------------------------------
   function  FileUpLoad(TargetFile, DelFile: String; var S_IP: String): Boolean;


  public
    { Public declarations }
  end;

var
  Client: TClient;

implementation

uses
  CComFunc,
   VarCom,
   TuxCom,
   MainDialog,
   ShellApi,   // ShellExec 사용관련 추가 @ 2013.08.22 LSH
   WinSock;

{$R *.DFM}


procedure TClient.ButtonSendClick(Sender: TObject);
begin
   ClientSocket.Socket.SendText(EditSay.Text + '&[' + EditNick.Text + ']');

   asg_Chat.InsertRows(iChatRowCnt, 1);
   //asg_Chat.MergeCells(C_CH_TEXT, iChatRowCnt, 2, 1);

   asg_Chat.Cells[C_CH_MYCOL, iChatRowCnt] := '[' + EditNick.Text + '] ';
   asg_Chat.Cells[C_CH_TEXT,  iChatRowCnt] := EditSay.Text;
   asg_Chat.Row := iChatRowCnt;
   asg_Chat.Alignments[C_CH_TEXT, iChatRowCnt] := taLeftJustify;
   asg_Chat.AutoSizeRow(iChatRowCnt);

   Inc(iChatRowCnt);


   EditSay.Text := '';
end;




procedure TClient.ConnectDisconnect1Click(Sender: TObject);
begin
   ClientSocket.Active := not ClientSocket.Active;
end;



procedure TClient.FormShow(Sender: TObject);
begin
   //---------------------------------------------------------
   // [Chat Box] Frame Set.
   //---------------------------------------------------------
   with asg_Chat do
   begin
      ColWidths[C_CH_MYCOL]   := 70;
      ColWidths[C_CH_TEXT]    := 220;
      ColWidths[C_CH_YOURCOL] := 75;
//      ColWidths[C_CH_HIDDEN]  := 0;
   end;

   //---------------------------------------------------------
   // [Chat 전송] Frame Set.
   //---------------------------------------------------------
   with asg_ChatSend do
   begin
      ColWidths[C_SD_TEXT]   := 318;
//      ColWidths[C_SD_BUTTON  := 227;

      AddButton(C_SD_BUTTON, 0, ColWidths[C_SD_BUTTON]-5, 20, '전송', haBeforeText, vaCenter);     // 전송 Button
   end;

   // 파일 및 이모티콘 전송 모드 Off
   IsEmotiSend := False;
   IsFileSend  := False;

end;




procedure TClient.ClientSocketConnect(Sender: TObject;
  Socket: TCustomWinSocket);
begin
   Self.Caption         := '챗(Chat) :: ' + ClientSocket.Socket.RemoteHost + ' 대화중';
   StatusBar.SimpleText := 'Status: Connected with ' + ClientSocket.Socket.RemoteHost;
end;




procedure TClient.ClientSocketDisconnect(Sender: TObject;
  Socket: TCustomWinSocket);
begin
   try
      Self.Caption         := '챗(Chat) :: ' + ClientSocket.Socket.RemoteHost + ' 대화종료';
      StatusBar.SimpleText := 'Status: just Disconnected.';

      ClientSocket.Close;

   except
      on E : Exception do
      ShowMessage(E.ClassName+' error raised by ClientSocket.active off, with message : ' + E.Message);
   end;



   { ------ 험난한 Debugging 끝에 일단, 오픈은 가능한 상태 !!!
   try

      if (ClientSocket.Active) then
      begin



         try
            ClientSocket.Active := False;

         except
            on E : Exception do
            ShowMessage(E.ClassName+' error raised by ClientSocket.active off, with message : ' + E.Message);
         end;
      end;


   except
      on E : Exception do
      ShowMessage(E.ClassName+' error raised by ClientSocket, with message : ' + E.Message);
   end;
   }
end;




procedure TClient.ClientSocketRead(Sender: TObject;
  Socket: TCustomWinSocket);
var
   sl, sReceived, tmpOrgReceived : string;
   sRead : string;
   Buffer: Array[0..10000] of byte;
   iReceived: Integer;
   hThreadID: Cardinal;
   FocusWnd : THandle;
begin

   //-----------------------------------------------------------------
   // Window API : 창을 비활성화 --> 활성화 @ 2014.02.28 LSH
   //     - 델마당 참조
   //-----------------------------------------------------------------
   ShowWindow(Self.Handle, SW_RESTORE);


   {  -- Chat 클라이언트는 윈도우가 비활성화 (화면 좌측하단)되기때문에
         SetForeGroundWindow가 먹질 않아서.. 그냥 매번 ShowWindow를 통해
         Restore 하도록 처리함.
   if IsWindowEnabled(Self.Handle) then
   begin
      hThreadID:= GetWindowThreadProcessId((GetForegroundWindow), nil);
      AttachThreadInput(GetCurrentThreadID, hThreadID, True);
      SetForegroundWindow(Self.Handle);
      AttachThreadInput(GetCurrentThreadID, hThreadID, False);
   end
   else
   begin
      Application.Restore;
   end;
   }


   sRead := '';
   sRead := socket.ReceiveText;


   Inc(iChatRowCnt);
   asg_Chat.RowCount := iChatRowCnt;   // Grid Index out of Range 방지위해 반드시 필요 !


   if (PosByte('just connected',     sRead) > 0)  or
      (PosByte('just disconnected',  sRead) > 0)  or
      (PosByte('[대화중]',           sRead) > 0)  then
   begin
      asg_Chat.InsertRows(iChatRowCnt, 1);
      asg_Chat.Cells[C_CH_TEXT,      iChatRowCnt] := sRead;
      asg_Chat.Row                                := iChatRowCnt;
      asg_Chat.Alignments[C_CH_TEXT, iChatRowCnt] := taCenter;
      asg_Chat.FontColors[C_CH_TEXT, iChatRowCnt] := $009589CC;
      asg_Chat.FontStyles[C_CH_TEXT, iChatRowCnt] := [fsBold];
      asg_Chat.AutoSizeRow(iChatRowCnt);



   end
   //-----------------------------------------------------------
   // 이모티콘(BMP 이미지) 수신모드
   //-----------------------------------------------------------
   else if (PosByte('[E]',  sRead) > 0) then
   begin

      // 원본 수신내역 Temp 파일에 저장 (Client User 이름 추출위함)
      tmpOrgReceived := sRead;

      DataSize := 0;
      sl := '';

      if not Reciving then
      begin
         // Now we need to get the LengthByte of the data stream.
         SetLength(sl, StrLen(PChar(sRead)) +1);           // +1 for the null terminator
         StrLCopy(@sl[1], PChar(sRead), LengthByte(sl)- 13);   // 이름태그(8) + &(1) + [E] (3) + 기존 Null (1)

         DataSize := StrToInt(sl);



         Data     := TMemoryStream.Create;
         Data.Size := 0;

         // DeleteByte the size information from the data.
         DeleteByte(sRead, 1, LengthByte(sl));

         Reciving:= True;
      end;




      // Store the data to the file, until we've received all the data.
      try
         Data.Write(sRead[1], LengthByte(sRead));


         if Data.Size = DataSize then
         begin

            Data.Position:= 0;


            try
               // 1줄 추가
               asg_Chat.InsertRows(iChatRowCnt, 1);

            except
               on e : Exception do
               showmessage('Client Read ERROR : ' +  e.Message);
            end;



            try
               // 수신한 Data-Stream을 Grid에 표기
               asg_Chat.CreatePicture(C_CH_TEXT, iChatRowCnt, True, ShrinkWithAspectRatio, 0, haRight, vaTop).Bitmap.LoadFromStream(Data);

            except
               asg_Chat.Cells[C_CH_TEXT, iChatRowCnt] := '[이미지 수신오류]';
            end;

            // Data 송신자 표기
            asg_Chat.Cells[C_CH_YOURCOL,   iChatRowCnt] := CopyByte(tmpOrgReceived, PosByte('&', tmpOrgReceived) + 1, 8);

            // Scroll 조절
            asg_Chat.Row := iChatRowCnt;

            // 송신자 컬럼 Align
            asg_Chat.Alignments[C_CH_YOURCOL, iChatRowCnt] := taLeftJustify;

            // 해당 Row Auto-Sizing
            asg_Chat.AutoSizeRow(iChatRowCnt);

            // 메모리 해제
            Data.Free;


            Reciving:= False;

         end;

      except
         on e : Exception do
         begin
            showmessage('Client Read Data.Write ERROR : ' + e.Message);

            Data.Free;
         end;
      end;

   end
   //-----------------------------------------------------------
   // 파일 수신모드
   //-----------------------------------------------------------
   else if (PosByte('[A]',  sRead) > 0) then
   begin

      asg_Chat.InsertRows(iChatRowCnt, 1);


      asg_Chat.AddButton(C_CH_MYCOL,iChatRowCnt, asg_Chat.ColWidths[C_CH_MYCOL]-20,  20, '다운', haBeforeText, vaCenter);
      asg_Chat.Cells[C_CH_TEXT,     iChatRowCnt] := CopyByte(sRead, 1, PosByte('$', sRead) - 1);  //CopyByte(sReceived, PosByte('$', sReceived) + 1, PosByte('[', sReceived) - PosByte('$', sReceived) -1); //
      asg_Chat.Cells[C_CH_YOURCOL,  iChatRowCnt] := CopyByte(sRead, PosByte('&', sRead) + 1, 8);
      asg_Chat.Cells[C_CH_HIDDEN,   iChatRowCnt] := CopyByte(sRead, PosByte('$', sRead) + 1, PosByte('[', sRead) - PosByte('$', sRead) -1);

      asg_Chat.Row := iChatRowCnt;

      asg_Chat.Alignments[C_CH_YOURCOL, iChatRowCnt] := taLeftJustify;

      if LengthByte(asg_Chat.Cells[C_CH_TEXT,     iChatRowCnt]) < 36 then
         asg_Chat.Alignments[C_CH_TEXT, iChatRowCnt] := taRightJustify
      else
         asg_Chat.Alignments[C_CH_TEXT, iChatRowCnt] := taLeftJustify;

      //asg_Chat.AutoSizeRow(iChatRowCnt);


   end
   //-----------------------------------------------------------
   // 일반 Text 수신모드
   //-----------------------------------------------------------
   else
   begin
      asg_Chat.InsertRows(iChatRowCnt, 1);
      //asg_Chat.MergeCells(C_CH_MYCOL, iChatRowCnt, 2, 1);

      asg_Chat.Cells[C_CH_YOURCOL,  iChatRowCnt] := CopyByte(sRead, PosByte('&', sRead) + 1, 10);
      asg_Chat.Cells[C_CH_TEXT,     iChatRowCnt] := CopyByte(sRead, 1, PosByte('&', sRead)-1);
      asg_Chat.Row := iChatRowCnt;

      asg_Chat.Alignments[C_CH_YOURCOL, iChatRowCnt] := taLeftJustify;

      if LengthByte(asg_Chat.Cells[C_CH_TEXT,     iChatRowCnt]) < 36 then
         asg_Chat.Alignments[C_CH_TEXT, iChatRowCnt] := taRightJustify
      else
         asg_Chat.Alignments[C_CH_TEXT, iChatRowCnt] := taLeftJustify;

      asg_Chat.AutoSizeRow(iChatRowCnt);
   end;


end;




procedure TClient.ButtonConnectClick(Sender: TObject);
begin
   ClientSocket.Host    := EditIp.Text;
   ClientSocket.Port    := StrToInt(MainDlg.DelChar(CopyByte(MainDlg.asg_Master.Cells[C_M_USERIP, MainDlg.asg_Master.Row], LengthByte(MainDlg.asg_Master.Cells[C_M_USERIP, MainDlg.asg_Master.Row])-3, 4), '.'));
   ClientSocket.Active  := True;
end;




procedure TClient.EditSayKeyPress(Sender: TObject; var Key: Char);
begin
   if (Key = #13) and (EditSay.Text <> '') then
      ButtonSendClick(Sender);
end;



procedure TClient.FormDestroy(Sender: TObject);
begin
   Client := Nil;
end;



//------------------------------------------------------------------------------
// [Chat 전송] AdvStringGrid onButtonClick Event Handler
//
// Author : Lee, Se-Ha
// Date   : 2013.09.09
//------------------------------------------------------------------------------
procedure TClient.asg_ChatSendButtonClick(Sender: TObject; ACol,
  ARow: Integer);
var
   ms : TMemoryStream;
   sFileName, sHideFile, sServerIp : String;
begin

   if (ACol = C_SD_BUTTON) then
   begin

      Inc(iChatRowCnt);
      asg_Chat.RowCount := iChatRowCnt;   // Grid Index out of Range 방지위해 반드시 필요 !


      //----------------------------------------------------------
      // 이모티콘(Bitmap) 전송 모드
      //----------------------------------------------------------
      if (IsEmotiSend) then
      begin

         ms := TMemoryStream.Create;

         try
            // 현재 불러온 이미지를 Memory-Stream으로 저장
            asg_ChatSend.GetPicture(C_SD_TEXT, 0).Bitmap.SaveToStream(ms);


            if ms.Size > 3000 then
            begin
               //
               MessageBox(self.Handle,
                          PChar('이모티콘용 Bitmap 이미지는 3KB 이하만 가능합니다.'),
                          '챗(Chat) :: 이모티콘(Bitmap) 전송 Size 제한',
                          MB_OK + MB_ICONERROR);

               // 메모리 해제
               ms.Free;

               // 입력창 이미지 제거
               asg_ChatSend.RemovePicture(C_SD_TEXT, 0);

               // 상태 Chagne
               IsEmotiSend := False;

               Exit;
            end;


            ms.Position := 0;


            // Data Stream 사이즈 정보 및 구분 Flag + 전송자 이름
            ClientSocket.Socket.SendText(IntToStr(ms.Size) + '[E]' + '&[' + EditNick.Text + ']' + #0);

            // 서버로 Data-Stream 전송
            ClientSocket.Socket.SendStream(ms);



            try
               // 1줄 추가
               asg_Chat.InsertRows(iChatRowCnt, 1);

            except
               on e : Exception do
               showmessage(e.Message);
            end;



            try
               // 전송창에 불러온 이미지를 표기
               asg_Chat.CreatePicture(C_CH_TEXT, iChatRowCnt, True, ShrinkWithAspectRatio, 0, haleft, vaTop).Bitmap.LoadFromFile(OpenPictureDialog1.FileName);

            except
               asg_Chat.Cells[C_CH_TEXT, iChatRowCnt] := '[이미지 전송실패]';
            end;

            // 전송자 표기
            asg_Chat.Cells[C_CH_MYCOL, iChatRowCnt]      := '[' + EditNick.Text + '] ';

            // Scroll 조절
            asg_Chat.Row   := iChatRowCnt;

            // 입력창 이미지 제거
            asg_ChatSend.RemovePicture(C_SD_TEXT, 0);

            // 상태 Change
            IsEmotiSend := False;

         except
            ms.Free;
         end;


      end
      //----------------------------------------------------------
      // File 전송 모드
      //----------------------------------------------------------
      else if (IsFileSend) then
      begin

         sFileName   := '';
         sHideFile   := '';
         sServerIp   := '';



         // 첨부파일 Upload
         if (Trim(asg_ChatSend.Cells[C_SD_TEXT, 0]) <> '') then
         begin

            sFileName := ExtractFileName(Trim(asg_ChatSend.Cells[C_SD_TEXT, 0]));
            sHideFile := 'CHATAPPEND' + MainDlg.DelChar(CopyByte(MainDlg.asg_Master.Cells[C_M_USERIP, MainDlg.asg_Master.Row], LengthByte(MainDlg.asg_Master.Cells[C_M_USERIP, MainDlg.asg_Master.Row])-3, 4), '.') +
                          FormatDateTime('YYYYMMDDHHNNSS', Now) + '.KDIALFILE';


            if Not FileUpLoad(sHideFile, asg_ChatSend.Cells[C_SD_TEXT, 0], sServerIp) then
            begin
               Showmessage('첨부파일 UpLoad 중 에러가 발생했습니다.' + #13#10 + #13#10 +
                           '다시한번 시도해 주시기 바랍니다.');
               Exit;
            end;
         end;


         ClientSocket.Socket.SendText(sFileName + '$' + sHideFile + '[A]' + '&[' + EditNick.Text + ']');


         asg_Chat.InsertRows(iChatRowCnt, 1);
         asg_Chat.Cells[C_CH_MYCOL,    iChatRowCnt]      := '[' + EditNick.Text + ']';
         asg_Chat.Cells[C_CH_TEXT,     iChatRowCnt]      := sFileName;
         asg_Chat.Cells[C_CH_HIDDEN,   iChatRowCnt]      := sHideFile;
         asg_Chat.Cells[C_CH_YOURCOL,  iChatRowCnt]      := '<전송됨>';
         //asg_Chat.AddButton(C_CH_YOURCOL,  iChatRowCnt, asg_Chat.ColWidths[C_CH_YOURCOL]-20,  20, '다운', haBeforeText, vaCenter);

         asg_Chat.Row                                 := iChatRowCnt;
         asg_Chat.Alignments[C_CH_TEXT, iChatRowCnt]  := taLeftJustify;
         //asg_Chat.AutoSizeRow(iChatRowCnt);



         asg_ChatSend.Cells[C_SD_TEXT, 0] := '';


         IsFileSend := False;

      end
      //----------------------------------------------------------
      // 일반 Text 전송 모드
      //----------------------------------------------------------
      else
      begin
         ClientSocket.Socket.SendText(asg_ChatSend.Cells[C_SD_TEXT, 0] + '&[' + EditNick.Text + ']');


         asg_Chat.InsertRows(iChatRowCnt, 1);
         asg_Chat.Cells[C_CH_MYCOL, iChatRowCnt]      := '[' + EditNick.Text + ']';
         asg_Chat.Cells[C_CH_TEXT,  iChatRowCnt]      := asg_ChatSend.Cells[C_SD_TEXT, 0];

         asg_Chat.Row                                 := iChatRowCnt;
         asg_Chat.Alignments[C_CH_TEXT, iChatRowCnt]  := taLeftJustify;
         asg_Chat.AutoSizeRow(iChatRowCnt);

         asg_ChatSend.Cells[C_SD_TEXT, 0] := '';
      end;



   {
      SendToAllClients(asg_ChatSend.Cells[C_SD_TEXT, 0] + '&[' + FsUserNm + ']');


      asg_Chat.InsertRows(iChatRowCnt, 1);
      asg_Chat.Cells[C_CH_MYCOL, iChatRowCnt] := '[' + FsUserNm + '] ';
      asg_Chat.Cells[C_CH_TEXT,  iChatRowCnt] := asg_ChatSend.Cells[C_SD_TEXT, 0];
      asg_Chat.Row := iChatRowCnt;
      asg_Chat.Alignments[C_CH_TEXT, iChatRowCnt] := taLeftJustify;
      asg_Chat.AutoSizeRow(iChatRowCnt);

   }


   end;
end;


//------------------------------------------------------------------------------
// [Chat 전송] AdvStringGrid onKeyPress Event Handler
//
// Author : Lee, Se-Ha
// Date   : 2013.09.09
//------------------------------------------------------------------------------
procedure TClient.asg_ChatSendKeyPress(Sender: TObject; var Key: Char);
begin
   if (Key = #13) and (asg_ChatSend.Cells[C_SD_TEXT, 0] <> '') then
      asg_ChatSendButtonClick(Sender, C_SD_BUTTON, 0);
end;



//------------------------------------------------------------------------------
// [종료] Form onClose Event Handler
//
// Author : Lee, Se-Ha
// Date   : 2013.09.02
//------------------------------------------------------------------------------
procedure TClient.FormClose(Sender: TObject; var Action: TCloseAction);
begin

   {----------------> Debugging
   try
         ClientSocket.Active := False;

      except
         on E : Exception do
         ShowMessage(E.ClassName+' error raised by ClientSocket.active off, with message : ' + E.Message);
      end;

   }
   {
   try
      ClientSocket.Socket.Disconnect(ClientSocket.Socket.SocketHandle);

      //ServerSocket.Socket.Disconnect(ServerSocket.Socket.SocketHandle);
      //ServerSocket.Socket.Close;
      //Action := caFree;
   except
      on E : Exception do
      ShowMessage(E.ClassName+' error raised by FormClose(Client), with message : ' + E.Message);
   end;
   }

   Action := caFree;

end;



//------------------------------------------------------------------------------
// [팝업메뉴] Chat 이모티콘 전송
//
// Author : Lee, Se-Ha
// Date   : 2013.09.09
//------------------------------------------------------------------------------
procedure TClient.mi_EmotiClick(Sender: TObject);
begin
   if OpenPictureDialog1.Execute then
   begin
      with asg_ChatSend do
      begin
         CreatePicture(C_SD_TEXT, 0, True, StretchWithAspectRatio {ShrinkWithAspectRatio}, 0, haLeft, vaTop).LoadFromFile(OpenPictureDialog1.FileName);
      end;
   end;

   if asg_ChatSend.GetPicture(C_SD_TEXT, 0) <> nil then
   begin
      // 이모티콘 전송모드 True
      IsEmotiSend := True;

      // 자동 전송
      asg_ChatSendButtonClick(Sender, C_SD_BUTTON, 0);
   end;
end;




//------------------------------------------------------------------------------
// ClientSocket onError Event Handler
//
// Author : Lee, Se-Ha
// Date   : 2013.09.02
//------------------------------------------------------------------------------
procedure TClient.ClientSocketError(Sender: TObject;
  Socket: TCustomWinSocket; ErrorEvent: TErrorEvent;
  var ErrorCode: Integer);
begin
   if (ErrorCode = 10053) then
   begin
      // ErrorCode = 0으로 두면, '응답없음'으로 강제 종료됨.. 버그 잡아야 함..
      ErrorCode := 1;
      MessageBox(self.Handle,
                 PChar('ClientSocketError(' + EditNick.Text + '): ' + Socket.RemoteHost + '(' + Socket.RemoteAddress + ') 로부터 잘못된 Connection Error 발생(소켓핸들값: ' + IntToStr(Socket.SocketHandle) + ')'),
                 '챗(Chat) :: 잘못된 Socket 통신 오류',
                 MB_OK + MB_ICONERROR);
   end;
end;



//------------------------------------------------------------------------------
// [팝업메뉴] Chat 파일 전송
//
// Author : Lee, Se-Ha
// Date   : 2013.09.09
//------------------------------------------------------------------------------
procedure TClient.mi_FileSendClick(Sender: TObject);
begin
   if od_File.Execute then
      asg_ChatSend.Cells[C_SD_TEXT, 0] := ExtractFileName(od_File.FileName);

   if asg_ChatSend.Cells[C_SD_TEXT, 0] <> '' then
   begin
      // 파일 전송모드 True
      IsFileSend := True;

      // 자동 전송
      asg_ChatSendButtonClick(Sender, C_SD_BUTTON, 0);
   end;


end;


//------------------------------------------------------------------------------
// [Chat 첨부] AdvStringGrid onButtonClick Event Handler
//
// Author : Lee, Se-Ha
// Date   : 2013.09.09
//------------------------------------------------------------------------------
procedure TClient.asg_ChatButtonClick(Sender: TObject; ACol,
  ARow: Integer);
var
   sLocalFile, sRemoteFile : String;
   sServerIp,  sFtpUserID, sFtpPasswd, sFtpHostName, sFtpDir : String;
begin
   //---------------------------------------------
   // 첨부파일 Click시 Event 처리
   //---------------------------------------------
   with asg_Chat do
   begin
      if ((ACol = C_CH_MYCOL) and (Cells[C_CH_HIDDEN, ARow] <> '')) then
      begin
         // 파일 업/다운로드를 위한 정보 조회
         if not GetBinUploadInfo(sServerIp, sFtpUserID, sFtpPasswd, sFtpHostName, sFtpDir) then
         begin
            MessageDlg('파일 저장을 위한 사용자 정보 조회중, 오류가 발생했습니다.', Dialogs.mtError, [Dialogs.mbOK], 0);
            exit;
         end;



         //실제 서버에 저장되어 있는 파일명 지정
         if PosByte('/ftpspool/CHATFILE/',Cells[C_CH_HIDDEN, ARow]) > 0 then
            sRemoteFile := Cells[C_CH_HIDDEN, ARow]
         else
            sRemoteFile := '/ftpspool/CHATFILE/' + Cells[C_CH_HIDDEN, ARow];

         // C:\DEV --> G_HOMEDIR 전환 @ 2016.07.26 LSH
         sLocalFile  := G_HOMEDIR + 'TEMP\SPOOL\' + Cells[C_CH_TEXT, ARow]; //lb_Filename.Caption;



         if (GetBINFTP(sServerIp, sFtpUserID, sFtpPasswd, sRemoteFile, sLocalFile, False)) then
         begin
            //	 ShowMessage('정상적으로 저장되었습니다.');
         end;



         Screen.Cursor := crDefault;



         try

            if (PosByte('.exe', sLocalFile) > 0) or
               (PosByte('.zip', sLocalFile) > 0) or
               (PosByte('.rar', sLocalFile) > 0) then
            begin
               MessageBox(self.Handle,
                          PChar('첨부파일(' + sLocalFile + ') 다운로드가 완료되었습니다.' + #13#10 + #13#10 +
                                '※ 임시 다운로드 폴더 --> C:\KUMC(_DEV)\TEMP\SPOOL\'),
                          '첨부파일 다운로드 실행 완료',
                          MB_OK + MB_ICONINFORMATION);

            end
            else
            begin
               ShellExecute(HANDLE, 'open',
                            PCHAR(Cells[C_CH_TEXT, ARow]),
                            PCHAR(''),
                            // C:\DEV --> G_HOMEDIR 전환 @ 2016.07.26 LSH
                            PCHAR(G_HOMEDIR + 'TEMP\SPOOL\'),
                            SW_SHOWNORMAL);
            end;


         except
            MessageBox(self.Handle,
                       PChar('해당 프로그램 다운로드중 오류가 발생하였습니다.' + #13#10 + #13#10 +
                             '프로그램 종료 후 다시 실행해 주시기 바랍니다.'),
                       '첨부파일 다운로드 중 오류',
                       MB_OK + MB_ICONERROR);

            Exit;
         end;
      end;
   end;
end;



//------------------------------------------------------------------------------
// [FTP] Chat-File Upload function Event Handler
//
// Date   : 2013.09.04
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
function TClient.FileUpLoad(TargetFile, DelFile: String; var S_IP: String):Boolean;
var
   sServerIp, sFtpUserId, sFtpPasswd, sFtpHostName, sFtpDir: String;
begin

   Result := True;


   // 파일 업로드를 위한 정보 조회
   if not GetBinUploadInfo(sServerIp, sFtpUserID, sFtpPasswd, sFtpHostName, sFtpDir) then
   begin
      ShowMessage('파일 저장을 위한 사용자 정보 조회중, 오류가 발생했습니다.');
      Result := False;
      exit;
   end;

   try
      if id_Ftp.Connected then
         nm_Ftp.Disconnect;
   except
      Showmessage('초기설정 에러');
      Result := False;
      Exit;
   end;


   nm_Ftp.Host     := sServerIp;
   nm_Ftp.UserID   := sFtpUserId;
   nm_Ftp.Password := sFtpPasswd;


   try
      nm_Ftp.Connect;

   except
      Showmessage('FTP 연결에러');
      Result := False;
      Exit;
   end;


   try
      nm_Ftp.ChangeDir('/ftpspool/CHATFILE');

   except
      Showmessage('ChangeDir 에러');
      Result := False;
      Exit;
   end;


   try
      nm_Ftp.mode(MODE_BYTE);

   except
      Showmessage('Mode 변경 에러');
      Result := False;
      Exit;
   end;

   //  -- 첨부내용 수정기능 미지원.. 주석처리 @ 2013.08.28 LSH
   {
   if DelFile <> '' then
   begin
      if Trim(fed_Attached.Text) <> DelFile then
      begin
         try
            nm_Ftp.Delete(DelFile);

         except
            Showmessage('파일삭제 에러');
            Result := False;
            Exit;
         end;
      end;
   end;
  }


   if Trim(asg_ChatSend.Cells[C_SD_TEXT, 0]) <> '' then
   begin
      try
         nm_Ftp.Upload(Trim(asg_ChatSend.Cells[C_SD_TEXT, 0]), TargetFile);

      except
         Showmessage('파일전송에러');
         Result := False;
         Exit;
      end;

      try
         nm_Ftp.DoCommand('SITE CHMOD 644 ' + TargetFile)

      except
         Showmessage('SITE CHMOD Error');
         Result := False;
         Exit;
      end;
   end;

   try
      nm_Ftp.Disconnect;

   except
      Showmessage('객체해제 에러');
      Result := False;
      Exit;
   end;


   S_IP := sServerIp;


end;


end.
