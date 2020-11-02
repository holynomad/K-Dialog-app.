object Server: TServer
  Left = 431
  Top = 320
  Width = 600
  Height = 523
  Caption = 'Server'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  FormStyle = fsStayOnTop
  Menu = MainMenu1
  OldCreateOrder = False
  Position = poMainFormCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Transparent = False
    Left = 304
    Top = 424
    Width = 254
    Height = 13
    Caption = 'Created by Wack-a-Mole [wackamonster@gmail.com]'
    Enabled = False
  end
  object Label2: TLabel
    Transparent = False
    Left = 328
    Top = 8
    Width = 196
    Height = 37
    Caption = 'Wack-a-Mole!'
    Enabled = False
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -32
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
  end
  object MemoChat: TMemo
    Left = 8
    Top = 56
    Width = 561
    Height = 345
    BorderStyle = bsNone
    ImeName = 'Microsoft IME 2010'
    Lines.Strings = (
      'Memo1')
    TabOrder = 0
  end
  object EditSay: TEdit
    Left = 8
    Top = 400
    Width = 497
    Height = 21
    BorderStyle = bsNone
    ImeName = 'Microsoft IME 2010'
    TabOrder = 1
  end
  object ButtonSend: TButton
    Left = 504
    Top = 400
    Width = 65
    Height = 21
    Caption = 'Send'
    TabOrder = 2
    OnClick = ButtonSendClick
  end
  object StatusBar: TStatusBar
    Left = 0
    Top = 450
    Width = 592
    Height = 19
    Panels = <>
    SimplePanel = True
    SimpleText = 'Status: Ready [not connected].'
  end
  object Button1: TButton
    Left = 192
    Top = 24
    Width = 73
    Height = 21
    Caption = 'Change'
    TabOrder = 4
  end
  object EditNick1: TEdit
    Left = 9
    Top = 24
    Width = 170
    Height = 21
    BorderStyle = bsNone
    ImeName = 'Microsoft IME 2010'
    TabOrder = 5
  end
  object ServerSocket: TServerSocket
    Active = False
    Port = 0
    ServerType = stNonBlocking
    OnClientConnect = ServerSocketClientConnect
    OnClientDisconnect = ServerSocketClientDisconnect
    OnClientRead = ServerSocketClientRead
    Left = 336
    Top = 72
  end
  object MainMenu1: TMainMenu
    Left = 424
    Top = 72
    object N1: TMenuItem
      Caption = 'File'
      object Listen: TMenuItem
        Caption = 'Listen for connections...'
      end
      object ChangeNickname1: TMenuItem
        Caption = 'Change Nickname...'
      end
      object Exit1: TMenuItem
        Caption = 'Exit'
      end
    end
  end
end
