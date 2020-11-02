{===============================================================================
   Program ID    : DlgChat.pas
   Program Name  : ä��(Chatting) Ŭ���̾�Ʈ
   Program Desc. : ���̾�α� ����������� IP�� ����� ��ȣ User�� 1:1 TCP ���
                   ����� �����ϴ� Client.

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


   // Chat �ڽ� Columns
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
   // FTP ���ε�
   //    - �ҽ���ó : MDN191F2.pas
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
   ShellApi,   // ShellExec ������ �߰� @ 2013.08.22 LSH
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
   // [Chat ����] Frame Set.
   //---------------------------------------------------------
   with asg_ChatSend do
   begin
      ColWidths[C_SD_TEXT]   := 318;
//      ColWidths[C_SD_BUTTON  := 227;

      AddButton(C_SD_BUTTON, 0, ColWidths[C_SD_BUTTON]-5, 20, '����', haBeforeText, vaCenter);     // ���� Button
   end;

   // ���� �� �̸�Ƽ�� ���� ��� Off
   IsEmotiSend := False;
   IsFileSend  := False;

end;




procedure TClient.ClientSocketConnect(Sender: TObject;
  Socket: TCustomWinSocket);
begin
   Self.Caption         := 'ê(Chat) :: ' + ClientSocket.Socket.RemoteHost + ' ��ȭ��';
   StatusBar.SimpleText := 'Status: Connected with ' + ClientSocket.Socket.RemoteHost;
end;




procedure TClient.ClientSocketDisconnect(Sender: TObject;
  Socket: TCustomWinSocket);
begin
   try
      Self.Caption         := 'ê(Chat) :: ' + ClientSocket.Socket.RemoteHost + ' ��ȭ����';
      StatusBar.SimpleText := 'Status: just Disconnected.';

      ClientSocket.Close;

   except
      on E : Exception do
      ShowMessage(E.ClassName+' error raised by ClientSocket.active off, with message : ' + E.Message);
   end;



   { ------ �賭�� Debugging ���� �ϴ�, ������ ������ ���� !!!
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
   // Window API : â�� ��Ȱ��ȭ --> Ȱ��ȭ @ 2014.02.28 LSH
   //     - ������ ����
   //-----------------------------------------------------------------
   ShowWindow(Self.Handle, SW_RESTORE);


   {  -- Chat Ŭ���̾�Ʈ�� �����찡 ��Ȱ��ȭ (ȭ�� �����ϴ�)�Ǳ⶧����
         SetForeGroundWindow�� ���� �ʾƼ�.. �׳� �Ź� ShowWindow�� ����
         Restore �ϵ��� ó����.
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
   asg_Chat.RowCount := iChatRowCnt;   // Grid Index out of Range �������� �ݵ�� �ʿ� !


   if (PosByte('just connected',     sRead) > 0)  or
      (PosByte('just disconnected',  sRead) > 0)  or
      (PosByte('[��ȭ��]',           sRead) > 0)  then
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
   // �̸�Ƽ��(BMP �̹���) ���Ÿ��
   //-----------------------------------------------------------
   else if (PosByte('[E]',  sRead) > 0) then
   begin

      // ���� ���ų��� Temp ���Ͽ� ���� (Client User �̸� ��������)
      tmpOrgReceived := sRead;

      DataSize := 0;
      sl := '';

      if not Reciving then
      begin
         // Now we need to get the LengthByte of the data stream.
         SetLength(sl, StrLen(PChar(sRead)) +1);           // +1 for the null terminator
         StrLCopy(@sl[1], PChar(sRead), LengthByte(sl)- 13);   // �̸��±�(8) + &(1) + [E] (3) + ���� Null (1)

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
               // 1�� �߰�
               asg_Chat.InsertRows(iChatRowCnt, 1);

            except
               on e : Exception do
               showmessage('Client Read ERROR : ' +  e.Message);
            end;



            try
               // ������ Data-Stream�� Grid�� ǥ��
               asg_Chat.CreatePicture(C_CH_TEXT, iChatRowCnt, True, ShrinkWithAspectRatio, 0, haRight, vaTop).Bitmap.LoadFromStream(Data);

            except
               asg_Chat.Cells[C_CH_TEXT, iChatRowCnt] := '[�̹��� ���ſ���]';
            end;

            // Data �۽��� ǥ��
            asg_Chat.Cells[C_CH_YOURCOL,   iChatRowCnt] := CopyByte(tmpOrgReceived, PosByte('&', tmpOrgReceived) + 1, 8);

            // Scroll ����
            asg_Chat.Row := iChatRowCnt;

            // �۽��� �÷� Align
            asg_Chat.Alignments[C_CH_YOURCOL, iChatRowCnt] := taLeftJustify;

            // �ش� Row Auto-Sizing
            asg_Chat.AutoSizeRow(iChatRowCnt);

            // �޸� ����
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
   // ���� ���Ÿ��
   //-----------------------------------------------------------
   else if (PosByte('[A]',  sRead) > 0) then
   begin

      asg_Chat.InsertRows(iChatRowCnt, 1);


      asg_Chat.AddButton(C_CH_MYCOL,iChatRowCnt, asg_Chat.ColWidths[C_CH_MYCOL]-20,  20, '�ٿ�', haBeforeText, vaCenter);
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
   // �Ϲ� Text ���Ÿ��
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
// [Chat ����] AdvStringGrid onButtonClick Event Handler
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
      asg_Chat.RowCount := iChatRowCnt;   // Grid Index out of Range �������� �ݵ�� �ʿ� !


      //----------------------------------------------------------
      // �̸�Ƽ��(Bitmap) ���� ���
      //----------------------------------------------------------
      if (IsEmotiSend) then
      begin

         ms := TMemoryStream.Create;

         try
            // ���� �ҷ��� �̹����� Memory-Stream���� ����
            asg_ChatSend.GetPicture(C_SD_TEXT, 0).Bitmap.SaveToStream(ms);


            if ms.Size > 3000 then
            begin
               //
               MessageBox(self.Handle,
                          PChar('�̸�Ƽ�ܿ� Bitmap �̹����� 3KB ���ϸ� �����մϴ�.'),
                          'ê(Chat) :: �̸�Ƽ��(Bitmap) ���� Size ����',
                          MB_OK + MB_ICONERROR);

               // �޸� ����
               ms.Free;

               // �Է�â �̹��� ����
               asg_ChatSend.RemovePicture(C_SD_TEXT, 0);

               // ���� Chagne
               IsEmotiSend := False;

               Exit;
            end;


            ms.Position := 0;


            // Data Stream ������ ���� �� ���� Flag + ������ �̸�
            ClientSocket.Socket.SendText(IntToStr(ms.Size) + '[E]' + '&[' + EditNick.Text + ']' + #0);

            // ������ Data-Stream ����
            ClientSocket.Socket.SendStream(ms);



            try
               // 1�� �߰�
               asg_Chat.InsertRows(iChatRowCnt, 1);

            except
               on e : Exception do
               showmessage(e.Message);
            end;



            try
               // ����â�� �ҷ��� �̹����� ǥ��
               asg_Chat.CreatePicture(C_CH_TEXT, iChatRowCnt, True, ShrinkWithAspectRatio, 0, haleft, vaTop).Bitmap.LoadFromFile(OpenPictureDialog1.FileName);

            except
               asg_Chat.Cells[C_CH_TEXT, iChatRowCnt] := '[�̹��� ���۽���]';
            end;

            // ������ ǥ��
            asg_Chat.Cells[C_CH_MYCOL, iChatRowCnt]      := '[' + EditNick.Text + '] ';

            // Scroll ����
            asg_Chat.Row   := iChatRowCnt;

            // �Է�â �̹��� ����
            asg_ChatSend.RemovePicture(C_SD_TEXT, 0);

            // ���� Change
            IsEmotiSend := False;

         except
            ms.Free;
         end;


      end
      //----------------------------------------------------------
      // File ���� ���
      //----------------------------------------------------------
      else if (IsFileSend) then
      begin

         sFileName   := '';
         sHideFile   := '';
         sServerIp   := '';



         // ÷������ Upload
         if (Trim(asg_ChatSend.Cells[C_SD_TEXT, 0]) <> '') then
         begin

            sFileName := ExtractFileName(Trim(asg_ChatSend.Cells[C_SD_TEXT, 0]));
            sHideFile := 'CHATAPPEND' + MainDlg.DelChar(CopyByte(MainDlg.asg_Master.Cells[C_M_USERIP, MainDlg.asg_Master.Row], LengthByte(MainDlg.asg_Master.Cells[C_M_USERIP, MainDlg.asg_Master.Row])-3, 4), '.') +
                          FormatDateTime('YYYYMMDDHHNNSS', Now) + '.KDIALFILE';


            if Not FileUpLoad(sHideFile, asg_ChatSend.Cells[C_SD_TEXT, 0], sServerIp) then
            begin
               Showmessage('÷������ UpLoad �� ������ �߻��߽��ϴ�.' + #13#10 + #13#10 +
                           '�ٽ��ѹ� �õ��� �ֽñ� �ٶ��ϴ�.');
               Exit;
            end;
         end;


         ClientSocket.Socket.SendText(sFileName + '$' + sHideFile + '[A]' + '&[' + EditNick.Text + ']');


         asg_Chat.InsertRows(iChatRowCnt, 1);
         asg_Chat.Cells[C_CH_MYCOL,    iChatRowCnt]      := '[' + EditNick.Text + ']';
         asg_Chat.Cells[C_CH_TEXT,     iChatRowCnt]      := sFileName;
         asg_Chat.Cells[C_CH_HIDDEN,   iChatRowCnt]      := sHideFile;
         asg_Chat.Cells[C_CH_YOURCOL,  iChatRowCnt]      := '<���۵�>';
         //asg_Chat.AddButton(C_CH_YOURCOL,  iChatRowCnt, asg_Chat.ColWidths[C_CH_YOURCOL]-20,  20, '�ٿ�', haBeforeText, vaCenter);

         asg_Chat.Row                                 := iChatRowCnt;
         asg_Chat.Alignments[C_CH_TEXT, iChatRowCnt]  := taLeftJustify;
         //asg_Chat.AutoSizeRow(iChatRowCnt);



         asg_ChatSend.Cells[C_SD_TEXT, 0] := '';


         IsFileSend := False;

      end
      //----------------------------------------------------------
      // �Ϲ� Text ���� ���
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
// [Chat ����] AdvStringGrid onKeyPress Event Handler
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
// [����] Form onClose Event Handler
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
// [�˾��޴�] Chat �̸�Ƽ�� ����
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
      // �̸�Ƽ�� ���۸�� True
      IsEmotiSend := True;

      // �ڵ� ����
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
      // ErrorCode = 0���� �θ�, '�������'���� ���� �����.. ���� ��ƾ� ��..
      ErrorCode := 1;
      MessageBox(self.Handle,
                 PChar('ClientSocketError(' + EditNick.Text + '): ' + Socket.RemoteHost + '(' + Socket.RemoteAddress + ') �κ��� �߸��� Connection Error �߻�(�����ڵ鰪: ' + IntToStr(Socket.SocketHandle) + ')'),
                 'ê(Chat) :: �߸��� Socket ��� ����',
                 MB_OK + MB_ICONERROR);
   end;
end;



//------------------------------------------------------------------------------
// [�˾��޴�] Chat ���� ����
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
      // ���� ���۸�� True
      IsFileSend := True;

      // �ڵ� ����
      asg_ChatSendButtonClick(Sender, C_SD_BUTTON, 0);
   end;


end;


//------------------------------------------------------------------------------
// [Chat ÷��] AdvStringGrid onButtonClick Event Handler
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
   // ÷������ Click�� Event ó��
   //---------------------------------------------
   with asg_Chat do
   begin
      if ((ACol = C_CH_MYCOL) and (Cells[C_CH_HIDDEN, ARow] <> '')) then
      begin
         // ���� ��/�ٿ�ε带 ���� ���� ��ȸ
         if not GetBinUploadInfo(sServerIp, sFtpUserID, sFtpPasswd, sFtpHostName, sFtpDir) then
         begin
            MessageDlg('���� ������ ���� ����� ���� ��ȸ��, ������ �߻��߽��ϴ�.', Dialogs.mtError, [Dialogs.mbOK], 0);
            exit;
         end;



         //���� ������ ����Ǿ� �ִ� ���ϸ� ����
         if PosByte('/ftpspool/CHATFILE/',Cells[C_CH_HIDDEN, ARow]) > 0 then
            sRemoteFile := Cells[C_CH_HIDDEN, ARow]
         else
            sRemoteFile := '/ftpspool/CHATFILE/' + Cells[C_CH_HIDDEN, ARow];

         // C:\DEV --> G_HOMEDIR ��ȯ @ 2016.07.26 LSH
         sLocalFile  := G_HOMEDIR + 'TEMP\SPOOL\' + Cells[C_CH_TEXT, ARow]; //lb_Filename.Caption;



         if (GetBINFTP(sServerIp, sFtpUserID, sFtpPasswd, sRemoteFile, sLocalFile, False)) then
         begin
            //	 ShowMessage('���������� ����Ǿ����ϴ�.');
         end;



         Screen.Cursor := crDefault;



         try

            if (PosByte('.exe', sLocalFile) > 0) or
               (PosByte('.zip', sLocalFile) > 0) or
               (PosByte('.rar', sLocalFile) > 0) then
            begin
               MessageBox(self.Handle,
                          PChar('÷������(' + sLocalFile + ') �ٿ�ε尡 �Ϸ�Ǿ����ϴ�.' + #13#10 + #13#10 +
                                '�� �ӽ� �ٿ�ε� ���� --> C:\KUMC(_DEV)\TEMP\SPOOL\'),
                          '÷������ �ٿ�ε� ���� �Ϸ�',
                          MB_OK + MB_ICONINFORMATION);

            end
            else
            begin
               ShellExecute(HANDLE, 'open',
                            PCHAR(Cells[C_CH_TEXT, ARow]),
                            PCHAR(''),
                            // C:\DEV --> G_HOMEDIR ��ȯ @ 2016.07.26 LSH
                            PCHAR(G_HOMEDIR + 'TEMP\SPOOL\'),
                            SW_SHOWNORMAL);
            end;


         except
            MessageBox(self.Handle,
                       PChar('�ش� ���α׷� �ٿ�ε��� ������ �߻��Ͽ����ϴ�.' + #13#10 + #13#10 +
                             '���α׷� ���� �� �ٽ� ������ �ֽñ� �ٶ��ϴ�.'),
                       '÷������ �ٿ�ε� �� ����',
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


   // ���� ���ε带 ���� ���� ��ȸ
   if not GetBinUploadInfo(sServerIp, sFtpUserID, sFtpPasswd, sFtpHostName, sFtpDir) then
   begin
      ShowMessage('���� ������ ���� ����� ���� ��ȸ��, ������ �߻��߽��ϴ�.');
      Result := False;
      exit;
   end;

   try
      if id_Ftp.Connected then
         nm_Ftp.Disconnect;
   except
      Showmessage('�ʱ⼳�� ����');
      Result := False;
      Exit;
   end;


   nm_Ftp.Host     := sServerIp;
   nm_Ftp.UserID   := sFtpUserId;
   nm_Ftp.Password := sFtpPasswd;


   try
      nm_Ftp.Connect;

   except
      Showmessage('FTP ���ῡ��');
      Result := False;
      Exit;
   end;


   try
      nm_Ftp.ChangeDir('/ftpspool/CHATFILE');

   except
      Showmessage('ChangeDir ����');
      Result := False;
      Exit;
   end;


   try
      nm_Ftp.mode(MODE_BYTE);

   except
      Showmessage('Mode ���� ����');
      Result := False;
      Exit;
   end;

   //  -- ÷�γ��� ������� ������.. �ּ�ó�� @ 2013.08.28 LSH
   {
   if DelFile <> '' then
   begin
      if Trim(fed_Attached.Text) <> DelFile then
      begin
         try
            nm_Ftp.Delete(DelFile);

         except
            Showmessage('���ϻ��� ����');
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
         Showmessage('�������ۿ���');
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
      Showmessage('��ü���� ����');
      Result := False;
      Exit;
   end;


   S_IP := sServerIp;


end;


end.
