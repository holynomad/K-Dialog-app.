object Client: TClient
  Left = 942
  Top = 434
  HorzScrollBar.Visible = False
  VertScrollBar.Visible = False
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = #52311'(Chat)'
  ClientHeight = 222
  ClientWidth = 367
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  FormStyle = fsStayOnTop
  OldCreateOrder = False
  Position = poMainFormCenter
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Bevel1: TBevel
    Left = 384
    Top = 80
    Width = 305
    Height = 49
  end
  object Bevel2: TBevel
    Left = 384
    Top = 136
    Width = 273
    Height = 49
  end
  object Image1: TImage
    Left = 368
    Top = 0
    Width = 42
    Height = 41
    Stretch = True
  end
  object MemoChat: TMemo
    Left = 648
    Top = 112
    Width = 588
    Height = 329
    BorderStyle = bsNone
    ImeName = 'Microsoft IME 2010'
    Lines.Strings = (
      '')
    TabOrder = 0
  end
  object ButtonConnect: TButton
    Left = 608
    Top = 104
    Width = 73
    Height = 21
    Caption = 'Connect'
    TabOrder = 1
    OnClick = ButtonConnectClick
  end
  object StatusBar: TStatusBar
    Left = 0
    Top = 203
    Width = 367
    Height = 19
    Panels = <>
    SimplePanel = True
  end
  object ButtonChangeNick: TButton
    Left = 584
    Top = 160
    Width = 73
    Height = 21
    Caption = 'Change'
    TabOrder = 3
  end
  object EditNick: TEdit
    Left = 392
    Top = 160
    Width = 182
    Height = 21
    BorderStyle = bsNone
    ImeName = 'Microsoft IME 2010'
    TabOrder = 4
  end
  object EditIp: TEdit
    Left = 392
    Top = 83
    Width = 182
    Height = 19
    BorderStyle = bsNone
    ImeName = 'Microsoft IME 2010'
    TabOrder = 5
  end
  object EditPort: TEdit
    Left = 392
    Top = 104
    Width = 182
    Height = 21
    BorderStyle = bsNone
    ImeName = 'Microsoft IME 2010'
    TabOrder = 6
  end
  object asg_Chat: TAdvStringGrid
    Left = 0
    Top = 0
    Width = 366
    Height = 180
    Cursor = crDefault
    ColCount = 4
    Ctl3D = False
    DefaultColWidth = 110
    DrawingStyle = gdsClassic
    FixedColor = 2784461
    FixedCols = 0
    RowCount = 1
    FixedRows = 0
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRowSelect]
    ParentCtl3D = False
    ParentFont = False
    ParentShowHint = False
    ScrollBars = ssVertical
    ShowHint = False
    TabOrder = 7
    GridLineColor = clWhite
    HoverRowCells = [hcNormal, hcSelected]
    OnButtonClick = asg_ChatButtonClick
    HintColor = 15718911
    HintShowCells = True
    ActiveCellFont.Charset = DEFAULT_CHARSET
    ActiveCellFont.Color = clWindowText
    ActiveCellFont.Height = -11
    ActiveCellFont.Name = 'MS Sans Serif'
    ActiveCellFont.Style = [fsBold]
    CellNode.ShowTree = False
    ControlLook.FixedGradientHoverFrom = clGray
    ControlLook.FixedGradientHoverTo = clWhite
    ControlLook.FixedGradientDownFrom = clGray
    ControlLook.FixedGradientDownTo = clSilver
    ControlLook.ControlStyle = csWinXP
    ControlLook.DropDownHeader.Font.Charset = DEFAULT_CHARSET
    ControlLook.DropDownHeader.Font.Color = clWindowText
    ControlLook.DropDownHeader.Font.Height = -11
    ControlLook.DropDownHeader.Font.Name = 'Tahoma'
    ControlLook.DropDownHeader.Font.Style = []
    ControlLook.DropDownHeader.Visible = True
    ControlLook.DropDownHeader.Buttons = <>
    ControlLook.DropDownFooter.Font.Charset = DEFAULT_CHARSET
    ControlLook.DropDownFooter.Font.Color = clWindowText
    ControlLook.DropDownFooter.Font.Height = -11
    ControlLook.DropDownFooter.Font.Name = 'Tahoma'
    ControlLook.DropDownFooter.Font.Style = []
    ControlLook.DropDownFooter.Visible = True
    ControlLook.DropDownFooter.Buttons = <>
    ControlLook.NoDisabledButtonLook = True
    EnhRowColMove = False
    Filter = <>
    FilterDropDown.Font.Charset = DEFAULT_CHARSET
    FilterDropDown.Font.Color = clWindowText
    FilterDropDown.Font.Height = -11
    FilterDropDown.Font.Name = 'Tahoma'
    FilterDropDown.Font.Style = []
    FilterDropDownClear = '(All)'
    FilterEdit.TypeNames.Strings = (
      'Starts with'
      'Ends with'
      'Contains'
      'Not contains'
      'Equal'
      'Not equal'
      'Larger than'
      'Smaller than'
      'Clear')
    FixedColWidth = 60
    FixedRowHeight = 22
    FixedFont.Charset = DEFAULT_CHARSET
    FixedFont.Color = clWhite
    FixedFont.Height = -11
    FixedFont.Name = 'MS Sans Serif'
    FixedFont.Style = []
    Flat = True
    FloatFormat = '%.2f'
    HideFocusRect = True
    HoverButtons.Buttons = <>
    HoverButtons.Position = hbLeftFromColumnLeft
    HTMLSettings.ImageFolder = 'images'
    HTMLSettings.ImageBaseName = 'img'
    Look = glSoft
    MouseActions.DirectEdit = True
    Navigation.AdvanceOnEnter = True
    Navigation.AutoComboDropSize = True
    PrintSettings.DateFormat = 'dd/mm/yyyy'
    PrintSettings.Font.Charset = DEFAULT_CHARSET
    PrintSettings.Font.Color = clWindowText
    PrintSettings.Font.Height = -11
    PrintSettings.Font.Name = 'MS Sans Serif'
    PrintSettings.Font.Style = []
    PrintSettings.FixedFont.Charset = DEFAULT_CHARSET
    PrintSettings.FixedFont.Color = clWindowText
    PrintSettings.FixedFont.Height = -11
    PrintSettings.FixedFont.Name = 'Tahoma'
    PrintSettings.FixedFont.Style = []
    PrintSettings.HeaderFont.Charset = DEFAULT_CHARSET
    PrintSettings.HeaderFont.Color = clWindowText
    PrintSettings.HeaderFont.Height = -11
    PrintSettings.HeaderFont.Name = 'MS Sans Serif'
    PrintSettings.HeaderFont.Style = []
    PrintSettings.FooterFont.Charset = DEFAULT_CHARSET
    PrintSettings.FooterFont.Color = clWindowText
    PrintSettings.FooterFont.Height = -11
    PrintSettings.FooterFont.Name = 'MS Sans Serif'
    PrintSettings.FooterFont.Style = []
    PrintSettings.Borders = pbNoborder
    PrintSettings.Centered = False
    PrintSettings.PagePrefix = 'page'
    PrintSettings.PageNumSep = '/'
    ScrollProportional = True
    ScrollType = ssFlat
    ScrollWidth = 16
    SearchFooter.Font.Charset = DEFAULT_CHARSET
    SearchFooter.Font.Color = clWindowText
    SearchFooter.Font.Height = -11
    SearchFooter.Font.Name = 'Tahoma'
    SearchFooter.Font.Style = []
    SelectionColor = 12570091
    SelectionTextColor = clNavy
    ShowSelection = False
    SortSettings.DefaultFormat = ssAutomatic
    SortSettings.Column = 0
    URLColor = clBlack
    Version = '7.8.6.0'
    ColWidths = (
      60
      227
      78
      110)
    RowHeights = (
      22)
  end
  object EditSay: TEdit
    Left = 1
    Top = 297
    Width = 303
    Height = 24
    BorderStyle = bsNone
    Color = 12705777
    ImeName = 'Microsoft IME 2010'
    TabOrder = 8
    OnKeyPress = EditSayKeyPress
  end
  object ButtonSend: TButton
    Left = 304
    Top = 299
    Width = 60
    Height = 21
    Caption = 'Send'
    TabOrder = 9
    OnClick = ButtonSendClick
  end
  object asg_ChatSend: TAdvStringGrid
    Left = 0
    Top = 179
    Width = 366
    Height = 27
    Cursor = crDefault
    Hint = 'Text '#51077#47141#54980' '#50644#53552' '#46608#45716' '#51204#49569#48260#53948#51012' '#45580#47084#51452#49464#50836'.'#13#10'('#47560#50864#49828' '#50864#53364#47533' --> '#54028#51068' '#48143' '#51060#48120#51648' '#51204#49569#47700#45684')'
    ColCount = 2
    Ctl3D = False
    DefaultColWidth = 110
    DrawingStyle = gdsClassic
    FixedColor = 2784461
    FixedCols = 0
    RowCount = 1
    FixedRows = 0
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    Options = [goDrawFocusSelected, goEditing]
    ParentCtl3D = False
    ParentFont = False
    ParentShowHint = False
    PopupMenu = pm_Chat
    ScrollBars = ssVertical
    ShowHint = True
    TabOrder = 10
    OnKeyPress = asg_ChatSendKeyPress
    GridLineColor = clWhite
    HoverRowCells = [hcNormal, hcSelected]
    OnButtonClick = asg_ChatSendButtonClick
    HintColor = 11002871
    HintShowCells = True
    ActiveCellFont.Charset = DEFAULT_CHARSET
    ActiveCellFont.Color = clWindowText
    ActiveCellFont.Height = -11
    ActiveCellFont.Name = 'MS Sans Serif'
    ActiveCellFont.Style = [fsBold]
    CellNode.ShowTree = False
    ColumnHeaders.Strings = (
      ''
      #54616#46168#49483#45367#45796#50668#51068#50668#50500#50676)
    ControlLook.FixedGradientHoverFrom = clGray
    ControlLook.FixedGradientHoverTo = clWhite
    ControlLook.FixedGradientDownFrom = clGray
    ControlLook.FixedGradientDownTo = clSilver
    ControlLook.ControlStyle = csWinXP
    ControlLook.DropDownHeader.Font.Charset = DEFAULT_CHARSET
    ControlLook.DropDownHeader.Font.Color = clWindowText
    ControlLook.DropDownHeader.Font.Height = -11
    ControlLook.DropDownHeader.Font.Name = 'Tahoma'
    ControlLook.DropDownHeader.Font.Style = []
    ControlLook.DropDownHeader.Visible = True
    ControlLook.DropDownHeader.Buttons = <>
    ControlLook.DropDownFooter.Font.Charset = DEFAULT_CHARSET
    ControlLook.DropDownFooter.Font.Color = clWindowText
    ControlLook.DropDownFooter.Font.Height = -11
    ControlLook.DropDownFooter.Font.Name = 'Tahoma'
    ControlLook.DropDownFooter.Font.Style = []
    ControlLook.DropDownFooter.Visible = True
    ControlLook.DropDownFooter.Buttons = <>
    ControlLook.NoDisabledButtonLook = True
    EnhRowColMove = False
    Filter = <>
    FilterDropDown.Font.Charset = DEFAULT_CHARSET
    FilterDropDown.Font.Color = clWindowText
    FilterDropDown.Font.Height = -11
    FilterDropDown.Font.Name = 'Tahoma'
    FilterDropDown.Font.Style = []
    FilterDropDownClear = '(All)'
    FilterEdit.TypeNames.Strings = (
      'Starts with'
      'Ends with'
      'Contains'
      'Not contains'
      'Equal'
      'Not equal'
      'Larger than'
      'Smaller than'
      'Clear')
    FixedColWidth = 318
    FixedRowHeight = 22
    FixedFont.Charset = DEFAULT_CHARSET
    FixedFont.Color = clWhite
    FixedFont.Height = -11
    FixedFont.Name = 'MS Sans Serif'
    FixedFont.Style = []
    Flat = True
    FloatFormat = '%.2f'
    HideFocusRect = True
    HoverButtons.Buttons = <>
    HoverButtons.Position = hbLeftFromColumnLeft
    HTMLSettings.ImageFolder = 'images'
    HTMLSettings.ImageBaseName = 'img'
    Look = glSoft
    MouseActions.DirectEdit = True
    Navigation.AutoComboDropSize = True
    PrintSettings.DateFormat = 'dd/mm/yyyy'
    PrintSettings.Font.Charset = DEFAULT_CHARSET
    PrintSettings.Font.Color = clWindowText
    PrintSettings.Font.Height = -11
    PrintSettings.Font.Name = 'MS Sans Serif'
    PrintSettings.Font.Style = []
    PrintSettings.FixedFont.Charset = DEFAULT_CHARSET
    PrintSettings.FixedFont.Color = clWindowText
    PrintSettings.FixedFont.Height = -11
    PrintSettings.FixedFont.Name = 'Tahoma'
    PrintSettings.FixedFont.Style = []
    PrintSettings.HeaderFont.Charset = DEFAULT_CHARSET
    PrintSettings.HeaderFont.Color = clWindowText
    PrintSettings.HeaderFont.Height = -11
    PrintSettings.HeaderFont.Name = 'MS Sans Serif'
    PrintSettings.HeaderFont.Style = []
    PrintSettings.FooterFont.Charset = DEFAULT_CHARSET
    PrintSettings.FooterFont.Color = clWindowText
    PrintSettings.FooterFont.Height = -11
    PrintSettings.FooterFont.Name = 'MS Sans Serif'
    PrintSettings.FooterFont.Style = []
    PrintSettings.Borders = pbNoborder
    PrintSettings.Centered = False
    PrintSettings.PagePrefix = 'page'
    PrintSettings.PageNumSep = '/'
    ScrollProportional = True
    ScrollType = ssFlat
    ScrollWidth = 16
    SearchFooter.Font.Charset = DEFAULT_CHARSET
    SearchFooter.Font.Color = clWindowText
    SearchFooter.Font.Height = -11
    SearchFooter.Font.Name = 'Tahoma'
    SearchFooter.Font.Style = []
    SelectionColor = 12570091
    SelectionTextColor = clNavy
    ShowSelection = False
    SortSettings.DefaultFormat = ssAutomatic
    SortSettings.Column = 0
    URLColor = clBlack
    Version = '7.8.6.0'
    ColWidths = (
      318
      45)
    RowHeights = (
      22)
  end
  object ClientSocket: TClientSocket
    Active = False
    ClientType = ctNonBlocking
    Port = 21
    OnConnect = ClientSocketConnect
    OnDisconnect = ClientSocketDisconnect
    OnRead = ClientSocketRead
    OnError = ClientSocketError
    Left = 280
    Top = 16
  end
  object pm_Chat: TPopupMenu
    Left = 207
    Top = 16
    object mi_Emoti: TMenuItem
      Caption = #51060#47784#54000#53080
      OnClick = mi_EmotiClick
    end
    object mi_FileSend: TMenuItem
      Caption = #54028#51068' '#48372#45236#44592
      OnClick = mi_FileSendClick
    end
  end
  object OpenPictureDialog1: TOpenPictureDialog
    Filter = 'Bitmaps (*.bmp)|*.bmp'
    InitialDir = 'C:\Documents and Settings\Administrator\My Documents\My Pictures'
    Left = 128
    Top = 16
  end
  object od_File: TOpenDialog
    Filter = '*.*'
    Left = 57
    Top = 17
  end
end
