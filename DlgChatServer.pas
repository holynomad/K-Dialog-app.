unit DlgChatServer;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, 
  Menus, ScktComp, ComCtrls, StdCtrls, Variants;

type
  TServer = class(TForm)
    Label1: TLabel;
    Label2: TLabel;
    MemoChat: TMemo;
    EditSay: TEdit;
    ButtonSend: TButton;
    StatusBar: TStatusBar;
    Button1: TButton;
    ServerSocket: TServerSocket;
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    Listen: TMenuItem;
    ChangeNickname1: TMenuItem;
    Exit1: TMenuItem;
    EditNick1: TEdit;
    function SendToAllClients( s : string) : boolean;
    procedure ServerSocketClientRead(Sender: TObject;
      Socket: TCustomWinSocket);
    procedure FormCreate(Sender: TObject);
    procedure ServerSocketClientConnect(Sender: TObject;
      Socket: TCustomWinSocket);
    procedure ServerSocketClientDisconnect(Sender: TObject;
      Socket: TCustomWinSocket);
    procedure ButtonSendClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Server: TServer;

implementation

{$R *.DFM}

procedure TServer.ServerSocketClientRead(Sender: TObject;
  Socket: TCustomWinSocket);
var
   received : string;
begin
   received := Socket.ReceiveText;
   SendToAllClients(received);
   MemoChat.Lines.Add(received);
end;

function TServer.SendToAllClients( s : string) : boolean;
var
   i : integer;
begin
   if ServerSocket.Socket.ActiveConnections = 0 then
   begin
          ShowMessage('There are no connected clients.');
          Exit;
   end;

   for i := 0 to (ServerSocket.Socket.ActiveConnections - 1) do
   begin
          ServerSocket.Socket.Connections[i].SendText(s);
   end;
end;

procedure TServer.FormCreate(Sender: TObject);
begin
   MemoChat.Clear;
   //EditNick1.Text := 'Unnamed Soldier';
end;

procedure TServer.ServerSocketClientConnect(Sender: TObject;
  Socket: TCustomWinSocket);
begin
   SendToAllClients('<'+Socket.RemoteHost+' just connected.>');
   MemoChat.Lines.Add('<'+Socket.RemoteHost+' just connected.>');
end;

procedure TServer.ServerSocketClientDisconnect(Sender: TObject;
  Socket: TCustomWinSocket);
begin
   SendToAllClients('<'+Socket.RemoteHost+' just disconnected.>');
   MemoChat.Lines.Add('<'+Socket.RemoteHost+' just disconnected.>');
end;

procedure TServer.ButtonSendClick(Sender: TObject);
begin
   SendToAllClients(EditNick1.Text + ' [Server] says: ' + EditSay.Text);
   MemoChat.Lines.Add(EditNick1.Text +' [Server] says: ' + EditSay.Text);
   EditSay.Text := '';
end;

end.
