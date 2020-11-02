{===============================================================================
   Program ID    : KDialChatClient.pas
   Program Name  : ���̾�α� Chat Client ����(BPL)
   Program Desc. : Chat ��� Client Socket ����

   Author        : Lee, Se-Ha
   Date          : 2015.03.27
===============================================================================}
unit KDialChatClient;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, 
  ExtDlgs, Menus, ScktComp, Grids, BaseGrid, AdvGrid, StdCtrls, ComCtrls, 
  ExtCtrls, Variants, AdvObj;


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
  TChatClient = class(TForm)
    Bevel1: TBevel;
    Bevel2: TBevel;
    Image1: TImage;
    MemoChat: TMemo;
    ButtonConnect: TButton;
    StatusBar: TStatusBar;
    ButtonChangeNick: TButton;
    EditNick: TEdit;
    EditIp: TEdit;
    EditPort: TEdit;
    asg_Chat: TAdvStringGrid;
    EditSay: TEdit;
    ButtonSend: TButton;
    asg_ChatSend: TAdvStringGrid;
    ClientSocket: TClientSocket;
    pm_Chat: TPopupMenu;
    mi_Emoti: TMenuItem;
    mi_FileSend: TMenuItem;
    OpenPictureDialog1: TOpenPictureDialog;
    od_File: TOpenDialog;
    procedure ClientSocketConnect(Sender: TObject;
      Socket: TCustomWinSocket);
    procedure ClientSocketDisconnect(Sender: TObject;
      Socket: TCustomWinSocket);
    procedure ClientSocketError(Sender: TObject; Socket: TCustomWinSocket;
      ErrorEvent: TErrorEvent; var ErrorCode: Integer);
    procedure ClientSocketRead(Sender: TObject; Socket: TCustomWinSocket);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure asg_ChatButtonClick(Sender: TObject; ACol, ARow: Integer);
    procedure asg_ChatSendButtonClick(Sender: TObject; ACol,
      ARow: Integer);
    procedure asg_ChatSendKeyPress(Sender: TObject; var Key: Char);
    procedure mi_EmotiClick(Sender: TObject);
    procedure mi_FileSendClick(Sender: TObject);
  private
    { Private declarations }
    FsEditIp,
    FsChatIp,
    FsChatUserNm,
    FsMyPhoto,
    FsChatPhoto,
    FsEditNick   : String;

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

    //--------------------------------------------------------------------------
    // String ���� Ư�� ���� ����
    //    - �ҽ���ó : MComfunc.pas
    //--------------------------------------------------------------------------
    function DelChar( const Str : String ; DelC : Char ) : String;

    //--------------------------------------------------------------------------
    // [�Լ�] TCP Port ���¿��� Check
    //--------------------------------------------------------------------------
    function PortTCPIsOpen(dwPort : Word; ipAddressStr : AnsiString) : boolean;

    //---------------------------------------------------------------------------
    // [������Ʈ] Log ������Ʈ (�Խñ� ��õ ��)
    //    - 2013.08.20 LSH
    //---------------------------------------------------------------------------
    procedure UpdateLog(in_Gubun,
                        in_BoardSeq,
                        in_UserIp,
                        in_LogFlag,
                        in_Item1,
                        in_Item2  : String;
                        var varResult : String);

   //----------------------------------------------------------------
   // Variant Array ������ �ڸ��� ��ĭ �ϳ� ����
   //    - �ҽ���ó : CComFunc.pas
   //----------------------------------------------------------------
   function AppendVariant(var in_var : Variant; in_str : String ) : Integer;

  public
    { Public declarations }
  published
    property prop_EditIp    : String read FsEditIp       write FsEditIp;
    property prop_EditNick  : String read FsEditNick     write FsEditNick;
    property prop_ChatIp    : String read FsChatIp       write FsChatIp;
    property prop_ChatUserNm: String read FsChatUserNm   write FsChatUserNm;
    property prop_ChatPhoto : String read FsChatPhoto    write FsChatPhoto;
    property prop_MyPhoto   : String read FsMyPhoto      write FsMyPhoto;

  end;

var
  ChatClient: TChatClient;

implementation


uses
   VarCom,
   CComfunc,
   StringLib,
   CMsg,
   MsgCom,
   Uencrypt,
   TuxMsg,
   SysMsg,
   TuxCom,
   HisUtil,
   TpSvc,   // TpCall Agent ������
//   Excel97,
   Imm,     // HIMC �������߰� @ 2013.08.22 LSH
   ShellApi,// ShellExec ������ �߰� @ 2013.08.22 LSH
   Winsock, // Windows Sockets API Unit
   ComObj,
   Qrctrls,
   QuickRpt,
   IdFTPCommon;   // ������ XE ���������� Id_FTP --> IndyFTP ��ȯ @ 2014.04.07 LSH

{$R *.DFM}


procedure TChatClient.ClientSocketConnect(Sender: TObject;
  Socket: TCustomWinSocket);
begin
   Self.Caption         := 'ê(Chat) :: ' + FsChatUserNm + '(' + ClientSocket.Socket.RemoteHost + ') ��ȭ��';
   StatusBar.SimpleText := 'Status: Connected with ' + ClientSocket.Socket.RemoteHost;
end;

procedure TChatClient.ClientSocketDisconnect(Sender: TObject;
  Socket: TCustomWinSocket);
begin
   try
      Self.Caption         := 'ê(Chat) :: ' + FsChatUserNm + '(' + ClientSocket.Socket.RemoteHost + ') ��ȭ����';
      StatusBar.SimpleText := 'Status: just Disconnected.';

      ClientSocket.Close;

   except
      on E : Exception do
      ShowMessage(E.ClassName+' error raised by ClientSocket.active off, with message : ' + E.Message);
   end;

end;

procedure TChatClient.ClientSocketError(Sender: TObject;
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

procedure TChatClient.ClientSocketRead(Sender: TObject;
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


         StrLCopy(@sl[1], PChar(sRead), {PosByte('[', PChar(sRead))-1}LengthByte(sl)- 13);   // �̸��±�(8) + &(1) + [E] (3) + ���� Null (1)


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
            //asg_Chat.Cells[C_CH_YOURCOL,   iChatRowCnt] := CopyByte(tmpOrgReceived, PosByte('&', tmpOrgReceived) + 1, 8);

            // �̸�Ƽ�� ���Ž�, ������ ������ �̹���(Ÿ��Ʋ) ���� �߰� @ 2015.03.31 LSH
            if FileExists(CopyByte(tmpOrgReceived, PosByte('&', tmpOrgReceived) + 2, LengthByte(tmpOrgReceived) - PosByte('&', tmpOrgReceived) - 2)) then
            begin
               // ��ȭ��� Image ������ Grid�� ǥ��
               asg_Chat.CreatePicture(C_CH_YOURCOL,
                                      iChatRowCnt,
                                      False,
                                      ShrinkWithAspectRatio,
                                      0,
                                      haLeft,
                                      vaTop).LoadFromFile(CopyByte(tmpOrgReceived, PosByte('&', tmpOrgReceived) + 2, LengthByte(tmpOrgReceived) - PosByte('&', tmpOrgReceived) - 2));

               asg_Chat.RowHeights[iChatRowCnt] := 40;

            end
            else
               asg_Chat.Cells[C_CH_YOURCOL, iChatRowCnt]      := '[' + FsChatUserNm + ']';



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


      asg_Chat.AddButton(C_CH_MYCOL,iChatRowCnt, asg_Chat.ColWidths[C_CH_MYCOL],  20, '�ٿ�', haBeforeText, vaCenter);

      asg_Chat.Cells[C_CH_TEXT,     iChatRowCnt] := CopyByte(sRead, 1, PosByte('$', sRead) - 1);  //CopyByte(sReceived, PosByte('$', sReceived) + 1, PosByte('[', sReceived) - PosByte('$', sReceived) -1); //

      //asg_Chat.Cells[C_CH_YOURCOL,  iChatRowCnt] := CopyByte(sRead, PosByte('&', sRead) + 1, 8);

      // ���� ���Ž�, ������ ������ �̹���(Ÿ��Ʋ) ���� �߰� @ 2015.03.31 LSH
      if FileExists(CopyByte(sRead, PosByte('&', sRead) + 2, LengthByte(sRead) - PosByte('&', sRead) - 2)) then
      begin
         // ��ȭ��� Image ������ Grid�� ǥ��
         asg_Chat.CreatePicture(C_CH_YOURCOL,
                                iChatRowCnt,
                                False,
                                ShrinkWithAspectRatio,
                                0,
                                haLeft,
                                vaTop).LoadFromFile(CopyByte(sRead, PosByte('&', sRead) + 2, LengthByte(sRead) - PosByte('&', sRead) - 2));

         asg_Chat.RowHeights[iChatRowCnt] := 40;

      end
      else
         asg_Chat.Cells[C_CH_YOURCOL, iChatRowCnt]      := '[' + FsChatUserNm + ']';

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
      //asg_Chat.Cells[C_CH_YOURCOL,  iChatRowCnt] := CopyByte(sRead, PosByte('&', sRead) + 1, 10);

      // test
      //showmessage('Chat Client���� ���Ž� : ' + CopyByte(sRead, PosByte('&', sRead) + 2, LengthByte(sRead) - PosByte('&', sRead) - 2));

      if FileExists(CopyByte(sRead, PosByte('&', sRead) + 2, LengthByte(sRead) - PosByte('&', sRead) - 2)) then
      begin
         // ��ȭ��� Image ������ Grid�� ǥ��
         asg_Chat.CreatePicture(C_CH_YOURCOL,
                                iChatRowCnt,
                                False,
                                ShrinkWithAspectRatio,
                                0,
                                haLeft,
                                vaTop).LoadFromFile(CopyByte(sRead, PosByte('&', sRead) + 2, LengthByte(sRead) - PosByte('&', sRead) - 2));

         asg_Chat.RowHeights[iChatRowCnt] := 40;

      end
      else
         asg_Chat.Cells[C_CH_YOURCOL, iChatRowCnt]      := '[' + FsChatUserNm + ']';



      asg_Chat.AutoSizeRow(iChatRowCnt);

      asg_Chat.Cells[C_CH_TEXT,     iChatRowCnt] := CopyByte(sRead, 1, PosByte('&', sRead)-1);
      asg_Chat.Row := iChatRowCnt;

      asg_Chat.Alignments[C_CH_YOURCOL, iChatRowCnt] := taLeftJustify;

      if LengthByte(asg_Chat.Cells[C_CH_TEXT,     iChatRowCnt]) < 36 then
         asg_Chat.Alignments[C_CH_TEXT, iChatRowCnt] := taRightJustify
      else
         asg_Chat.Alignments[C_CH_TEXT, iChatRowCnt] := taLeftJustify;


   end;

end;



//------------------------------------------------------------------------------
// [FTP] Chat-File Upload function Event Handler
//
// Date   : 2013.09.04
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
function TChatClient.FileUpLoad(TargetFile, DelFile: String; var S_IP: String):Boolean;
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

   //
   ShowMessage('FTP ���ε�/�ٿ�ε� ��� ������ ���� �ð��� �غ��ؼ� �����ҰԿ� :-)');

//   try
//      if nm_Ftp.Connected then
//         nm_Ftp.Disconnect;
//   except
//      Showmessage('�ʱ⼳�� ����');
//      Result := False;
//      Exit;
//   end;
//
//
//   nm_Ftp.Host     := sServerIp;
//   nm_Ftp.UserID   := sFtpUserId;
//   nm_Ftp.Password := sFtpPasswd;
//
//
//   try
//      nm_Ftp.Connect;
//
//   except
//      Showmessage('FTP ���ῡ��');
//      Result := False;
//      Exit;
//   end;


//   try
//      nm_Ftp.ChangeDir('/ftpspool/CHATFILE');
//
//   except
//      Showmessage('ChangeDir ����');
//      Result := False;
//      Exit;
//   end;
//
//
//   try
//      nm_Ftp.mode(MODE_BYTE);
//
//   except
//      Showmessage('Mode ���� ����');
//      Result := False;
//      Exit;
//   end;
//
//   //  -- ÷�γ��� ������� ������.. �ּ�ó�� @ 2013.08.28 LSH
//   {
//   if DelFile <> '' then
//   begin
//      if Trim(fed_Attached.Text) <> DelFile then
//      begin
//         try
//            nm_Ftp.Delete(DelFile);
//
//         except
//            Showmessage('���ϻ��� ����');
//            Result := False;
//            Exit;
//         end;
//      end;
//   end;
//  }
//
//
//   if Trim(asg_ChatSend.Cells[C_SD_TEXT, 0]) <> '' then
//   begin
//      try
//         nm_Ftp.Upload(Trim(asg_ChatSend.Cells[C_SD_TEXT, 0]), TargetFile);
//
//      except
//         Showmessage('�������ۿ���');
//         Result := False;
//         Exit;
//      end;
//
//      try
//         nm_Ftp.DoCommand('SITE CHMOD 644 ' + TargetFile)
//
//      except
//         Showmessage('SITE CHMOD Error');
//         Result := False;
//         Exit;
//      end;
//   end;
//
//   try
//      nm_Ftp.Disconnect;
//
//   except
//      Showmessage('��ü���� ����');
//      Result := False;
//      Exit;
//   end;


   S_IP := sServerIp;


end;



procedure TChatClient.FormClose(Sender: TObject; var Action: TCloseAction);
begin
   ChatClient := nil;
   Action := caFree;
end;



procedure TChatClient.FormShow(Sender: TObject);
type
   TProcedureType  = Procedure(Sender: TObject) of Object;
   TProcedureType2 = Procedure(Sender: TObject; in_Gubun : String) of Object;
var
   varResult   : String;
   Procedure1  : TProcedureType;
   Procedure2  : TProcedureType2;
   PProcedure1 : PMethod;
   PProcedure2 : PMethod;
begin
   //---------------------------------------------------------
   // [Chat Box] Frame Set.
   //---------------------------------------------------------
   with asg_Chat do
   begin
      ColWidths[C_CH_MYCOL]   := 40;
      ColWidths[C_CH_TEXT]    := 285;
      ColWidths[C_CH_YOURCOL] := 40;
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



   // ���� �� �̸�Ƽ�� ���۸�� Off
   IsEmotiSend := False;
   IsFileSend  := False;


   //---------------------------------------------------------
   // Main ȭ�� Chatting - Log ������Ʈ ����
   //---------------------------------------------------------
   if FindForm('MAINDLG') <> nil then
   begin
      PProcedure1 := @TMethod(Procedure1);

      (GetComp('MAINDLG', 'pn_ChatUserIp') as TPanel).Caption := FsChatIp;

      if FindFormMethod('MAINDLG', 'SetUpdChatLog', PProcedure1) = True then
      begin
         Procedure1(Sender);
      end;
   end;



   try
      Screen.Cursor := crHourGlass;


      // ���Ӵ�� IP�� �г��� ����
      EditIp.Text   := FsEditIp;
      EditNick.Text := FsEditNick;

      // ���� ������ IP �� Port ����
      ClientSocket.Host    := FsChatIp;
      ClientSocket.Port    := StrToInt(DelChar(CopyByte(FsChatIp, LengthByte(FsChatIp)-3, 4), '.'));


      //---------------------------------------------------------
      // �����Ϸ��� Server Port ��밡�� ���� Check
      //---------------------------------------------------------
      if PortTCPIsOpen(ClientSocket.Port, ClientSocket.Host) then
      begin
         ClientSocket.Active  := True;
      end
      else
      begin

         MessageBox(self.Handle,
                    PChar('�����Ϸ��� ��� Port�� �����ֽ��ϴ�.' + #13#10 + #13#10 + '���� ������� ���α׷��� ���������� Ȯ���� �ʿ��մϴ�.'),
                    'ê(Chat) �ڽ� :: ���� ���� Port ����',
                    MB_OK + MB_ICONWARNING);


         ClientSocket.Active  := False;

         //---------------------------------------------------------
         // Main ȭ�� ������ User Log-out ������Ʈ ����
         //---------------------------------------------------------
         if FindForm('MAINDLG') <> nil then
         begin
            PProcedure1 := @TMethod(Procedure1);

            (GetComp('MAINDLG', 'pn_ChatUserIp') as TPanel).Caption := FsChatIp;

            if FindFormMethod('MAINDLG', 'SetUpdLogOut', PProcedure1) = True then
            begin
               Procedure1(Sender);
            end;
         end;


         //---------------------------------------------------------
         // ���̾� Chat ����Ʈ Refresh
         //---------------------------------------------------------
         if FindForm('KDIALCHAT') <> nil then
         begin
            PProcedure2 := @TMethod(Procedure2);

            if FindFormMethod('KDIALCHAT', 'SelChatList', PProcedure2) = True then
            begin
               Procedure2(Sender, 'NETWORK');
            end;
         end;

         // Chat-Client �� ����
         Close;

      end;


   finally
      Screen.Cursor := crDefault;
   end;
end;


//------------------------------------------------------------------------------
// [�Լ�] TCP Port ���¿��� Check
//    - 2013.08.30 LSH
//    - �ҽ���ó : Google
//    - XE7 ��ȯ�� ipAddressStr Ÿ�� String -> AnsiString ��ȯ
//------------------------------------------------------------------------------
function TChatClient.PortTCPIsOpen(dwPort : Word; ipAddressStr : AnsiString) : boolean;
var
   client : sockaddr_in;                                                // sockaddr_in is used by Windows Sockets to specify a local or remote endpoint address
   sock   : Integer;
begin
   client.sin_family       := AF_INET;
   client.sin_port         := htons(dwPort);                            // htons converts a u_short from host to TCP/IP network byte order.
   client.sin_addr.s_addr  := inet_addr(PAnsiChar(ipAddressStr));       // the inet_addr function converts a string containing an IPv4 dotted-decimal address into a proper address for the IN_ADDR structure. --> XE7 ���̱׷��̼ǽ� PChar -> PAnsiChar ��ȯ
   sock                    := socket(AF_INET, SOCK_STREAM, 0);          // The socket function creates a socket
   Result                  := connect(sock,client,SizeOf(client)) = 0;  // establishes a connection to a specified socket.
end;



// String ���� Ư�� ���� ����
function TChatClient.DelChar( const Str : String ; DelC : Char ) : String;
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



procedure TChatClient.asg_ChatButtonClick(Sender: TObject; ACol,
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

procedure TChatClient.asg_ChatSendButtonClick(Sender: TObject; ACol,
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
      // �̸�Ƽ��(Bitmap) ���۸��
      //----------------------------------------------------------
      if (IsEmotiSend) then
      begin
         // test
         //showmessage('�̸�Ƽ�� ���۸��');

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


            // Data Stream ������ ���� �� ���� Flag + ������ �̸� @ 2015.03.31 LSH
            if FileExists(FsMyPhoto) then
               ClientSocket.Socket.SendText(IntToStr(ms.Size) + #0 + '[E]' + '&[' + FsMyPhoto + ']' + '%[' + FsChatPhoto + ']' + #0)
            else
               ClientSocket.Socket.SendText(IntToStr(ms.Size) + #0 + '[E]' + '&[' + EditNick.Text + ']' + '%[' + FsChatPhoto + ']' + #0);

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

            // ������ ǥ�� (�̹��� ������ �̹��� ǥ��� ��ȯ) @ 2015.03.30 LSH
            if FileExists(FsMyPhoto) then
            begin
               // ���� Image ������ Grid�� ǥ��
               asg_Chat.CreatePicture(C_CH_MYCOL,
                                      iChatRowCnt,
                                      False,
                                      ShrinkWithAspectRatio,
                                      0,
                                      haLeft,
                                      vaTop).LoadFromFile(FsMyPhoto);

               asg_Chat.RowHeights[iChatRowCnt] := 40;

            end
            else
               asg_Chat.Cells[C_CH_MYCOL, iChatRowCnt]      := '[' + EditNick.Text + '] ';


            asg_Chat.AutoSizeRow(iChatRowCnt);

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
      // File ���۸��
      //----------------------------------------------------------
      else if (IsFileSend) then
      begin
         // test
         //showmessage('���� ���۸��');

         sFileName   := '';
         sHideFile   := '';
         sServerIp   := '';



         // ÷������ Upload
         if (Trim(asg_ChatSend.Cells[C_SD_TEXT, 0]) <> '') then
         begin

            sFileName := ExtractFileName(Trim(asg_ChatSend.Cells[C_SD_TEXT, 0]));
            sHideFile := 'CHATAPPEND' + DelChar(CopyByte(FsChatIp, LengthByte(FsChatIp)-3, 4), '.') +
                          FormatDateTime('YYYYMMDDHHNNSS', Now) + '.KDIALFILE';


            if Not FileUpLoad(sHideFile, asg_ChatSend.Cells[C_SD_TEXT, 0], sServerIp) then
            begin
               Showmessage('÷������ UpLoad �� ������ �߻��߽��ϴ�.' + #13#10 + #13#10 +
                           '�ٽ��ѹ� �õ��� �ֽñ� �ٶ��ϴ�.');
               Exit;
            end;
         end;



         if FileExists(FsMyPhoto) then
            ClientSocket.Socket.SendText(sFileName + '$' + sHideFile + '[A]' + '&[' + FsMyPhoto + ']' + '%[' + FsChatPhoto + ']')
         else
            ClientSocket.Socket.SendText(sFileName + '$' + sHideFile + '[A]' + '&[' + EditNick.Text + ']' + '%[' + FsChatPhoto + ']');


         asg_Chat.InsertRows(iChatRowCnt, 1);

         //asg_Chat.Cells[C_CH_MYCOL,    iChatRowCnt]      := '[' + EditNick.Text + ']';

         // ������ ǥ�� (�̹��� ������ �̹��� ǥ��� ��ȯ) @ 2015.03.30 LSH
         if FileExists(FsMyPhoto) then
         begin
            // ���� Image ������ Grid�� ǥ��
            asg_Chat.CreatePicture(C_CH_MYCOL,
                                   iChatRowCnt,
                                   False,
                                   ShrinkWithAspectRatio,
                                   0,
                                   haLeft,
                                   vaTop).LoadFromFile(FsMyPhoto);

            asg_Chat.RowHeights[iChatRowCnt] := 40;

         end
         else
            asg_Chat.Cells[C_CH_MYCOL,    iChatRowCnt]      := '[' + EditNick.Text + ']';


         asg_Chat.AutoSizeRow(iChatRowCnt);

         asg_Chat.Cells[C_CH_TEXT,     iChatRowCnt]      := sFileName;
         asg_Chat.Cells[C_CH_HIDDEN,   iChatRowCnt]      := sHideFile;
         asg_Chat.Cells[C_CH_YOURCOL,  iChatRowCnt]      := '<���۵�>';
         //asg_Chat.AddButton(C_CH_YOURCOL,  iChatRowCnt, asg_Chat.ColWidths[C_CH_YOURCOL]-20,  20, '�ٿ�', haBeforeText, vaCenter);

         asg_Chat.Row                                 := iChatRowCnt;
         asg_Chat.Alignments[C_CH_TEXT, iChatRowCnt]  := taLeftJustify;




         asg_ChatSend.Cells[C_SD_TEXT, 0] := '';


         IsFileSend := False;

      end
      //----------------------------------------------------------
      // �Ϲ� Text ���۸��
      //----------------------------------------------------------
      else
      begin
         // test
         //showmessage('��� : ' + FsChatPhoto + ' , �� : ' + FsMyPhoto);


         if FileExists(FsMyPhoto) then
            ClientSocket.Socket.SendText(asg_ChatSend.Cells[C_SD_TEXT, 0] + '&[' + FsMyPhoto + ']' + '%[' + FsChatPhoto + ']')
         else
            ClientSocket.Socket.SendText(asg_ChatSend.Cells[C_SD_TEXT, 0] + '&[' + EditNick.Text + ']' + '%[' + FsChatPhoto + ']');



         asg_Chat.InsertRows(iChatRowCnt, 1);

         //asg_Chat.Cells[C_CH_MYCOL, iChatRowCnt]      := '[' + EditNick.Text + ']';

         // ������ ǥ�� (�̹��� ������ �̹��� ǥ��� ��ȯ) @ 2015.03.30 LSH
         if FileExists(FsMyPhoto) then
         begin

            // ���� Image ������ Grid�� ǥ��
            asg_Chat.CreatePicture(C_CH_MYCOL,
                                   iChatRowCnt,
                                   False,
                                   ShrinkWithAspectRatio,
                                   0,
                                   haLeft,
                                   vaTop).LoadFromFile(FsMyPhoto);

            asg_Chat.RowHeights[iChatRowCnt] := 40;

         end
         else
            asg_Chat.Cells[C_CH_MYCOL, iChatRowCnt]      := '[' + EditNick.Text + ']';


         asg_Chat.AutoSizeRow(iChatRowCnt);

         asg_Chat.Cells[C_CH_TEXT,  iChatRowCnt]      := asg_ChatSend.Cells[C_SD_TEXT, 0];

         asg_Chat.Row                                 := iChatRowCnt;
         asg_Chat.Alignments[C_CH_TEXT, iChatRowCnt]  := taLeftJustify;


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

procedure TChatClient.asg_ChatSendKeyPress(Sender: TObject; var Key: Char);
begin
   if (Key = #13) and (asg_ChatSend.Cells[C_SD_TEXT, 0] <> '') then
      asg_ChatSendButtonClick(Sender, C_SD_BUTTON, 0);
end;

procedure TChatClient.mi_EmotiClick(Sender: TObject);
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

procedure TChatClient.mi_FileSendClick(Sender: TObject);
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
// [Update] �Խ��� Log ����
//
// Author : Lee, Se-Ha
// Date   : 2013.08.20
//------------------------------------------------------------------------------
procedure TChatClient.UpdateLog(in_Gubun,
                                in_BoardSeq,
                                in_UserIp,
                                in_LogFlag,
                                in_Item1,
                                in_Item2  : String;
                                var varResult : String);
var
   TpUpdLog : TTpSvc;
   vType1   ,
   vBoardNm ,
   vBoardSeq,
   vUserIp  ,
   vLogFlag ,
   vItem1   ,
   vItem2   ,
   vEditIp  : Variant;
begin
   varResult := 'N';


   //------------------------------------------------------------------
   // 2-1. Append Variables
   //------------------------------------------------------------------
   AppendVariant(vType1   ,  '5');
   AppendVariant(vBoardNm ,  in_Gubun);
   AppendVariant(vBoardSeq,  in_BoardSeq);
   AppendVariant(vUserIp  ,  in_UserIp);
   AppendVariant(vLogFlag ,  in_LogFlag);
   AppendVariant(vItem1   ,  in_Item1);
   AppendVariant(vItem2   ,  in_Item2);
   AppendVariant(vEditIp  ,  FsEditIp);


   //-------------------------------------------------------------------
   // 3. Insert by TpSvc
   //-------------------------------------------------------------------
   TpUpdLog := TTpSvc.Create;
   TpUpdLog.Init(Self);



   Screen.Cursor := crHourGlass;






   try
      if TpUpdLog.PutSvc('MD_KUMCM_M1',
                          [
                            'S_TYPE1'   , vType1
                         ,  'S_STRING35', vBoardNm
                         ,  'S_STRING25', vBoardSeq
                         ,  'S_STRING7' , vUserIp
                         ,  'S_STRING36', vLogFlag
                         ,  'S_STRING37', vItem1
                         ,  'S_STRING38', vItem2
                         ,  'S_STRING8' , vEditIp
                         ] ) then
      begin


         {
         if (in_LogFlag = 'L') then
            MessageBox(self.Handle,
                       PChar('������ �Խñ��� ���������� [��õ] �Ǿ����ϴ�.'),
                       '[KUMC ���̾�α�] Ŀ�´�Ƽ ������Ʈ �˸� ',
                       MB_OK + MB_ICONINFORMATION);
         }
         varResult := 'Y';
      end
      else
      begin

         // SVC �ܿ��� GetTxMsg ó����. ���� ó�� �ʿ� ����.
         varResult := 'N';

      end;


   finally
      FreeAndNil(TpUpdLog);
      Screen.Cursor  := crDefault;
   end;

end;



//------------------------------------------------------------------------------
// [�Լ�] Variant �Ӽ� ���� �Լ�
//       - �ҽ���ó : CComFunc.pas
//
// Author : Lee, Se-Ha
// Date   : 2013.08.08
//------------------------------------------------------------------------------
function TChatClient.AppendVariant(var in_var : Variant; in_str : String ) : Integer;
var
   cnt     : Integer;
   tmp_str : String;
begin
   if ( VarIsArray(in_var) ) then
   begin
      cnt := VarArrayHighBound(in_var,1);

      VarArrayRedim(in_var,cnt + 1 );

      in_var[cnt+1] := in_str;
      result := cnt+1;
   end
   else
   begin
      if ( VarIsEmpty(in_var) ) or
         ( VarIsNull(in_var) ) then
      begin
         in_var    := VarArrayCreate([0,0],varVariant);
         in_var[0] := in_str;
         result    := 0;
      end
      else
      begin
         tmp_str   := VarToStr(in_var);
         in_var    := VarArrayCreate([0,1],varVariant);
         in_var[0] := tmp_str;
         in_var[1] := in_str;
         result    := 1;
      end;
   end;
end;

initialization
   RegisterClass(TChatClient);

finalization
   UnRegisterClass(TChatClient);

end.