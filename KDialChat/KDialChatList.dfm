object KDialChat: TKDialChat
  Left = 1164
  Top = 194
  BorderIcons = [biSystemMenu]
  BorderStyle = bsSingle
  Caption = #45796#51060#50620' Chat'
  ClientHeight = 566
  ClientWidth = 153
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  FormStyle = fsStayOnTop
  OldCreateOrder = False
  Position = poDesigned
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object lb_Status: TLabel
    Left = 5
    Top = 32
    Width = 142
    Height = 13
    Caption = #9654' 1'#48516#47560#45796' '#51088#46041' Refresh'#51473'..'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clMaroon
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    Transparent = False
  end
  object fsbt_Network: TFlatSpeedButton
    Left = 123
    Top = 3
    Width = 26
    Height = 23
    Hint = #48708#49345#50672#46973#47581
    Color = clBtnFace
    ColorFocused = 10053171
    ColorDown = 14410211
    ColorBorder = 11254206
    Glyph.Data = {
      36030000424D3603000000000000360000002800000010000000100000000100
      18000000000000030000130B0000130B00000000000000000000FF00FFFF00FF
      FF00FFFF00FFFF00FFFF00FFA86852A86852A77662FF00FFFF00FFFF00FFFF00
      FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFA86852A86852F8ECE9C9
      B6ADF8ECE9733614FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF
      A86852A86852E8D1C9F8ECE9CA9B8CC6A89AF8ECE9F8ECE9733614FF00FFFF00
      FFFF00FFFF00FFFF00FFFF00FFA86852E8D1C9F8ECE9F2E3DFCA9B8CB7857397
      5A40733614F8ECE9F8ECE9733614FF00FFFF00FFFF00FFFF00FFA86852E8C5BD
      F8ECE9CA9B8CCA9B8CAF735CA96A50A06044965739733614F8ECE9F8ECE97336
      14FF00FFFF00FFFF00FFA86852DBADA2C3907FBF8772C38A72BD836AB4795FAA
      6D52A26348965739733614F8ECE9F8ECE9733614FF00FFFF00FFA86852D3A392
      DEB09FDAA490D19B84C78F77BD846BB3785DAB6E53A2634796563A733614F8EC
      E9F8ECE9733614FF00FFFF00FFA86852EECBBEF7E4DDDEAB96D39D87CA917ABE
      856BB57A60AB6E53A2634898583C733614F8ECE9F8ECE9733614FF00FFFF00FF
      A86852EECDC1F7E4DDDDAA96D49D88C89079C0866DB57A60AB6E53A263489757
      3B733614B68F809B624DFF00FFFF00FFFF00FFA86852EECBBFF7E4DDDEAA96D3
      9D86CA927BC0866DB67B61AD7055A4654A9A5A3E733614FF00FFFF00FFFF00FF
      FF00FFFF00FFA86852EDCCC0F7E4DDDDA994D49D88C99179BF866DB57A60AC6F
      54A264499B60479B624DFF00FFFF00FFFF00FFFF00FFFF00FFA86852EECBBFF7
      E4DDE1AD99D59F89CA937CBB8068A96E56A868529B624DFF00FFFF00FFFF00FF
      FF00FFFF00FFFF00FFFF00FFA86852EDC9BCE9BEAED5A08BB3745EB3745E0000
      98A86852FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFA8
      6852B3745EB3745E9A8FBD5E7CFB3056FA000098FF00FFFF00FFFF00FFFF00FF
      FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FF0030F8BDC8FF0030
      F8FF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF00FFFF
      00FFFF00FFFF00FFFF00FF0030F8FF00FFFF00FFFF00FFFF00FF}
    ParentColor = False
    ParentShowHint = False
    ShowHint = True
    OnClick = fsbt_NetworkClick
  end
  object fsbt_DialBook: TFlatSpeedButton
    Left = 98
    Top = 3
    Width = 26
    Height = 23
    Hint = 'My '#45796#51060#50620#48513
    Color = clBtnFace
    ColorFocused = 10053171
    ColorDown = 14410211
    ColorBorder = 11254206
    Glyph.Data = {
      36030000424D3603000000000000360000002800000010000000100000000100
      1800000000000003000012170000121700000000000000000000FF00FF518F56
      497F4E42724836603D36603D36603D36603D36603D36603D36603D36603D86A0
      8B6C6C6B7A7777FF00FFB9B4B2837B7760606096C09A92BD968EBA928AB78E86
      B48A82B1867EAE8279AB7E75A77A8994869A9798C8C4C6908D8C8D8581FFFFFF
      A49B969AC39E96C09A92BD968EBA928AB78EB6D2B9B4D0B6B2CEB4AAC4A97372
      71C8C4C6E0E0E0948F8FFF00FF518F56A2C9A69EC6A29AC39E96C09A92BC95B3
      C9B295998F81807B797A757D8A79807C79E0E0E0948E90FF00FFFF00FF518F56
      A6CCAAA2C9A69EC6A29AC39EB0C0AD908D878F8B85AB9C909D93877D7C787878
      729B9593928E8AFF00FFB9B4B2837B77606060A6CCAAA2C9A6BED2BD96928CA4
      9F94CFC2A6DFC2A8DDC7B0DDC4AB8D88827A7D7686A08BFF00FF8D8581FFFFFF
      A49B96ABCFAECAE0CCA8A99E9E9891D3BA9DD6B698DFC0A6EADCCBE4D3BFDDC3
      AD7F7D79767C72FF00FFFF00FF518F56AED1B1AED1B1CDE2CE9A958EBEAFA0D4
      AD8BD9B598EAD4C0E6EBDFEBDECCE2CDB7A59C917E7E7AFF00FFFF00FF518F56
      AED1B1AED1B1CEE3D099938CC4AE9CD29D77DDB596E6CCB6E4CAB1E8D2BEEAD2
      BDB0A79D8B8683FF00FFB9B4B2837B77606060AED1B191B3959A9F90A89E94CA
      996ED6B190FAF5EFDCC0A5D9BA9ED7B799908C8695968CFF00FF8D8581FFFFFF
      A49B96AED1B147814EC7DAC89C978EB5A79AD6B796FAF6F0EBEEE3DED5C1A39C
      939A948F869B87FF00FFFF00FF518F56AED1B1AED1B147814ECCE2CEBFCEBC9D
      988FA59F95BFB6ACB9B0A79A958F9F9B94B4C7B237603DFF00FFFF00FF518F56
      AED1B1AED1B147814E47814E91B39591B094ACAA9F9E98929D9A91989F90C1D6
      C19AC39E36603DFF00FFB9B4B2837B77606060AED1B1AED1B1AED1B1AED1B1AE
      D1B1CEE3D0CEE3D0CDE2CECAE0CCA2C9A69EC6A2427248FF00FF8D8581FFFFFF
      A49B96AED1B1AED1B1AED1B1AED1B1AED1B1AED1B1AED1B1AED1B1ABCFAEA6CC
      AAA2C9A6497F4EFF00FFFF00FF518F56518F56518F56518F56518F56518F5651
      8F56518F56518F56518F56518F56518F56518F56518F56FF00FF}
    ParentColor = False
    ParentShowHint = False
    ShowHint = True
    OnClick = fsbt_DialBookClick
  end
  object fsbt_Search: TFlatSpeedButton
    Left = 73
    Top = 3
    Width = 26
    Height = 23
    Hint = #53301' '#49436#52824'('#50672#46973#52376' '#44160#49353')'
    ColorFocused = 10053171
    Glyph.Data = {
      22050000424D2205000000000000360000002800000013000000150000000100
      180000000000EC04000000000000000000000000000000000000FFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF000000FFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF000000FFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF932A00CD6600FFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFF000000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF932A00CD6600FFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFF000000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFF932A00FFD8A9CA6200FFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFF000000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFF932A00FFD8A9C96000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF00
      0000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      932A00FFD4A4FFD4A4CA6100FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF000000FFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF932A00FF
      D2A1FFD2A2CC6300FFFFFFFFFFFFFFFFFFFFFFFFFFFFFF000000FFFFFFFFFFFF
      FFFFFFB24B00932A00932A00932A00932A00922900932A00FFCE9BFFCE9BFFD1
      A0C95E00902700B24B00FFFFFFFFFFFFFFFFFF000000FFFFFFFFFFFFCB6200FF
      DDAEFFDAADFFDBAEFFDBAEFFDBAEFFDAAEFFD9ADFFD4A8FFD2A5FFD3A7FFD7AB
      FFDCAECB6200FFFFFFFFFFFFFFFFFF000000FFFFFFFFFFFFC96000FFDFB6FF99
      00FF9C05FF9C05FF9C05FF9C05FF9C05FF9C05FF9C05FF9C05FF9900FFDFB6C9
      6000FFFFFFFFFFFFFFFFFF000000FFFFFFFFFFFFC95F00FFE5C0FFA826FFAB2D
      FFAB2DFFAB2DFFAB2DFFAB2DFFAB2DFFAB2DFFAB2DFFA826FFE5C0C95F00FFFF
      FFFFFFFFFFFFFF000000FFFFFFFFFFFFC95F00FFEBC7FFB545FFB74BFFB74BFF
      B74BFFB74BFFB74BFFB74BFFB74BFFB74BFFB545FFEBC7C95F00FFFFFFFFFFFF
      FFFFFF000000FFFFFFFFFFFFC95F00FFF3D0FFC264FFC469FFC469FFC469FFC4
      69FFC469FFC469FFC469FFC469FFC264FFF3D0C95F00FFFFFFFFFFFFFFFFFF00
      0000FFFFFFFFFFFFC95E00FFF9D8FFCE84FFD088FFD088FFD088FFD088FFD088
      FFD088FFD088FFD088FFCE84FFF9D8C95E00FFFFFFFFFFFFFFFFFF000000FFFF
      FFFFFFFFC95E00FFFFDFFFDAA1FFDBA4FFDBA4FFDBA4FFDBA4FFDBA4FFDBA4FF
      DBA4FFDBA4FFDAA1FFFFDFC95E00FFFFFFFFFFFFFFFFFF000000FFFFFFFFFFFF
      CA5F00FFFFF2FFFFE9FFFFE9FFFFE9FFFFE9FFFFE9FFFFE9FFFFE9FFFFE9FFFF
      E9FFFFE9FFFFF2CA5F00FFFFFFFFFFFFFFFFFF000000FFFFFFFFFFFFFFFFFFCA
      5F00C95E00C95E00C95E00C95E00C95E00C95E00C95E00C95E00C95E00C95E00
      CA5F00FFFFFFFFFFFFFFFFFFFFFFFF000000FFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFF000000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFF000000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF
      FFFFFF000000}
    ParentShowHint = False
    ShowHint = True
    OnClick = fsbt_SearchClick
  end
  object fgau_Progress: TFlatGauge
    Left = 0
    Top = 541
    Width = 153
    Height = 24
    AdvColorBorder = 0
    ColorBorder = 8623776
    Progress = 25
  end
  object asg_ChatList: TAdvStringGrid
    Left = 0
    Top = 55
    Width = 153
    Height = 487
    Cursor = crDefault
    Hint = #50672#46160#49353#51004#47196' '#54364#44592#46108' '#47196#44536#51064#54620' User'#47484' '#45908#48660#53364#47533'!'
    ColCount = 2
    Ctl3D = False
    DefaultColWidth = 40
    DefaultRowHeight = 60
    DrawingStyle = gdsClassic
    FixedColor = 13610846
    FixedCols = 0
    RowCount = 4
    FixedRows = 0
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goDrawFocusSelected, goRowSelect]
    ParentCtl3D = False
    ParentFont = False
    ParentShowHint = False
    ScrollBars = ssVertical
    ShowHint = True
    TabOrder = 0
    OnClick = asg_ChatListClick
    OnDblClick = asg_ChatListDblClick
    HoverRowCells = [hcNormal, hcSelected]
    HintColor = clSilver
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
    ControlLook.ControlStyle = csClassic
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
    FixedColWidth = 40
    FixedRowHeight = 60
    FixedFont.Charset = DEFAULT_CHARSET
    FixedFont.Color = clWhite
    FixedFont.Height = -11
    FixedFont.Name = 'Tahoma'
    FixedFont.Style = []
    Flat = True
    FloatFormat = '%.2f'
    HoverButtons.Buttons = <>
    HoverButtons.Position = hbLeftFromColumnLeft
    HTMLSettings.ImageFolder = 'images'
    HTMLSettings.ImageBaseName = 'img'
    Look = glSoft
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
    PrintSettings.PageNumSep = '/'
    SearchFooter.Font.Charset = DEFAULT_CHARSET
    SearchFooter.Font.Color = clWindowText
    SearchFooter.Font.Height = -11
    SearchFooter.Font.Name = 'Tahoma'
    SearchFooter.Font.Style = []
    SelectionColor = 10053171
    SelectionTextColor = clHighlightText
    SortSettings.DefaultFormat = ssAutomatic
    SortSettings.Column = 0
    Version = '7.8.6.0'
    ColWidths = (
      40
      93)
  end
  object fed_Scan: TFlatEdit
    Left = 4
    Top = 30
    Width = 145
    Height = 20
    Hint = #51060#47492', '#50672#46973#52376', '#49324#48264', '#50836#52397#51088' '#46321#51004#47196' '#44160#49353#54644' '#48372#49464#50836'.'
    ColorBorder = 14013909
    ColorFlat = 14013909
    ParentShowHint = False
    ShowHint = True
    TabOrder = 1
    Visible = False
    OnClick = fed_ScanClick
    OnKeyPress = fed_ScanKeyPress
  end
  object fed_Statusbar: TFlatEdit
    Left = 1
    Top = 526
    Width = 151
    Height = 19
    ColorBorder = clBtnFace
    ColorFlat = clBtnFace
    ParentShowHint = False
    ShowHint = True
    TabOrder = 2
    OnClick = fed_ScanClick
    OnKeyPress = fed_ScanKeyPress
  end
  object tm_Chat: TTimer
    Interval = 60000
    OnTimer = tm_ChatTimer
    Left = 16
    Top = 291
  end
  object ServerSocket: TServerSocket
    Active = False
    Port = 21
    ServerType = stNonBlocking
    Left = 86
    Top = 292
  end
end
