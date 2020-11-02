object FtpPrint: TFtpPrint
  Left = 604
  Top = 79
  BorderStyle = bsSingle
  Caption = '[FtpPrint] KUMC '#45796#51060#50620#47196#44536' FTP '#52636#47141#50577#49885
  ClientHeight = 1042
  ClientWidth = 930
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesigned
  Scaled = False
  OnClose = FormClose
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object qr_KDial: TQuickRep
    Left = 5
    Top = 6
    Width = 794
    Height = 1123
    BeforePrint = qr_KDialBeforePrint
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Arial'
    Font.Style = []
    Functions.Strings = (
      'PAGENUMBER'
      'COLUMNNUMBER'
      'REPORTTITLE')
    Functions.DATA = (
      '0'
      '0'
      #39#39)
    Options = [FirstPageHeader, LastPageFooter]
    Page.Columns = 1
    Page.Orientation = poPortrait
    Page.PaperSize = A4
    Page.Continuous = False
    Page.Values = (
      0.000000000000000000
      2970.000000000000000000
      0.000000000000000000
      2100.000000000000000000
      0.000000000000000000
      0.000000000000000000
      0.000000000000000000)
    PrinterSettings.Copies = 1
    PrinterSettings.OutputBin = Auto
    PrinterSettings.Duplex = False
    PrinterSettings.FirstPage = 0
    PrinterSettings.LastPage = 0
    PrinterSettings.UseStandardprinter = False
    PrinterSettings.UseCustomBinCode = False
    PrinterSettings.CustomBinCode = 0
    PrinterSettings.ExtendedDuplex = 0
    PrinterSettings.UseCustomPaperCode = False
    PrinterSettings.CustomPaperCode = 0
    PrinterSettings.PrintMetaFile = False
    PrinterSettings.PrintQuality = 0
    PrinterSettings.Collate = 0
    PrinterSettings.ColorOption = 0
    PrintIfEmpty = True
    SnapToGrid = True
    Units = MM
    Zoom = 100
    PrevFormStyle = fsNormal
    PreviewWidth = 500
    PreviewHeight = 500
    PrevInitialZoom = qrZoomToFit
    PreviewDefaultSaveType = stQRP
    PreviewLeft = 0
    PreviewTop = 0
    object DetailBand1: TQRBand
      Left = 0
      Top = 0
      Width = 794
      Height = 1123
      AlignToBottom = False
      TransparentBand = False
      ForceNewColumn = False
      ForceNewPage = False
      Size.Values = (
        2971.270833333333000000
        2100.791666666667000000)
      PreCaluculateBandHeight = False
      KeepOnOnePage = False
      BandType = rbDetail
      object qrimg_KDial: TQRImage
        Left = 1
        Top = 1
        Width = 792
        Height = 1122
        Size.Values = (
          2968.625000000000000000
          2.645833333333330000
          2.645833333333330000
          2095.500000000000000000)
        XLColumn = 0
        Center = True
        Stretch = True
      end
      object qrlb_DocTitle: TQRLabel
        Left = 144
        Top = 253
        Width = 523
        Height = 19
        Size.Values = (
          50.270833333333300000
          381.000000000000000000
          669.395833333333000000
          1383.770833333330000000)
        XLColumn = 0
        Alignment = taLeftJustify
        AlignToBand = False
        Caption = 'PACS '#49884#49828#53596' '#51452'/'#48372#51312' '#51200#51109#51109#52824'(EMC) 2014'#45380' 2'#50900' '#50976#51648#48372#49688' '#51648#44553
        Color = clWhite
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -17
        Font.Name = #44404#47548#52404
        Font.Style = []
        ParentFont = False
        Transparent = False
        ExportAs = exptText
        WrapStyle = BreakOnSpaces
        FontSize = 13
      end
      object qrlb_HqAmt: TQRLabel
        Left = 158
        Top = 594
        Width = 105
        Height = 15
        Size.Values = (
          39.687500000000000000
          418.041666666667000000
          1571.625000000000000000
          277.812500000000000000)
        XLColumn = 0
        Alignment = taCenter
        AlignToBand = False
        AutoSize = False
        Caption = '0'#50896
        Color = clWhite
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #44404#47548#52404
        Font.Style = []
        ParentFont = False
        Transparent = False
        ExportAs = exptText
        WrapStyle = BreakOnSpaces
        FontSize = 10
      end
      object qrlb_Period: TQRLabel
        Left = 268
        Top = 462
        Width = 201
        Height = 18
        Size.Values = (
          47.625000000000000000
          709.083333333333000000
          1222.375000000000000000
          531.812500000000000000)
        XLColumn = 0
        Alignment = taLeftJustify
        AlignToBand = False
        Caption = '2013. 7. 1 ~ 2014. 6. 30 '
        Color = clWhite
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -15
        Font.Name = #44404#47548#52404
        Font.Style = []
        ParentFont = False
        Transparent = False
        ExportAs = exptText
        WrapStyle = BreakOnSpaces
        FontSize = 11
      end
      object qrlb_GrAmt: TQRLabel
        Left = 388
        Top = 594
        Width = 90
        Height = 15
        Size.Values = (
          39.687500000000000000
          1026.583333333330000000
          1571.625000000000000000
          238.125000000000000000)
        XLColumn = 0
        Alignment = taCenter
        AlignToBand = False
        AutoSize = False
        Caption = '1,330,000'#50896
        Color = clWhite
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #44404#47548#52404
        Font.Style = []
        ParentFont = False
        Transparent = False
        ExportAs = exptText
        WrapStyle = BreakOnSpaces
        FontSize = 10
      end
      object qrlb_DocName: TQRLabel
        Left = 268
        Top = 438
        Width = 417
        Height = 20
        Size.Values = (
          52.916666666666700000
          709.083333333333000000
          1158.875000000000000000
          1103.312500000000000000)
        XLColumn = 0
        Alignment = taLeftJustify
        AlignToBand = False
        Caption = 'PACS '#49884#49828#53596' '#44288#47144' '#51452'/'#48372#51312' '#51200#51109#51109#52824'(EMC) '#50976#51648#48372#49688' '#51648#44553
        Color = clWhite
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -15
        Font.Name = #44404#47548#52404
        Font.Style = []
        ParentFont = False
        Transparent = False
        ExportAs = exptText
        WrapStyle = BreakOnSpaces
        FontSize = 11
      end
      object qrlb_BarPatno: TQRLabel
        Left = 914
        Top = 21
        Width = 171
        Height = 33
        Size.Values = (
          87.312500000000000000
          2418.291666666670000000
          55.562500000000000000
          452.437500000000000000)
        XLColumn = 0
        Alignment = taLeftJustify
        AlignToBand = False
        Caption = '*01484864*'
        Color = clWhite
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -32
        Font.Name = 'Code39(1:2)'
        Font.Style = []
        ParentFont = False
        Transparent = False
        ExportAs = exptText
        WrapStyle = BreakOnSpaces
        FontSize = 24
      end
      object qrlb_AaAmt: TQRLabel
        Left = 275
        Top = 594
        Width = 94
        Height = 15
        Size.Values = (
          39.687500000000000000
          727.604166666667000000
          1571.625000000000000000
          248.708333333333000000)
        XLColumn = 0
        Alignment = taCenter
        AlignToBand = False
        AutoSize = False
        Caption = '1,365,000'#50896
        Color = clWhite
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #44404#47548#52404
        Font.Style = []
        ParentFont = False
        Transparent = False
        ExportAs = exptText
        WrapStyle = BreakOnSpaces
        FontSize = 10
      end
      object qrlb_TotAmt1: TQRLabel
        Left = 295
        Top = 509
        Width = 89
        Height = 18
        Size.Values = (
          47.625000000000000000
          780.520833333333000000
          1346.729166666670000000
          235.479166666667000000)
        XLColumn = 0
        Alignment = taLeftJustify
        AlignToBand = False
        Caption = '3,500,000'#50896
        Color = clWhite
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -15
        Font.Name = #44404#47548#52404
        Font.Style = []
        ParentFont = False
        Transparent = False
        ExportAs = exptText
        WrapStyle = BreakOnSpaces
        FontSize = 11
      end
      object qrlb_Company: TQRLabel
        Left = 268
        Top = 415
        Width = 113
        Height = 16
        Size.Values = (
          42.333333333333300000
          709.083333333333000000
          1098.020833333330000000
          298.979166666667000000)
        XLColumn = 0
        Alignment = taLeftJustify
        AlignToBand = False
        Caption = 'LG '#50644#49884#49828' ('#51452')'
        Color = clWhite
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -15
        Font.Name = #44404#47548#52404
        Font.Style = []
        ParentFont = False
        Transparent = False
        ExportAs = exptText
        WrapStyle = BreakOnSpaces
        FontSize = 11
      end
      object qrlb_AsAmt: TQRLabel
        Left = 494
        Top = 594
        Width = 84
        Height = 15
        Size.Values = (
          39.687500000000000000
          1307.041666666670000000
          1571.625000000000000000
          222.250000000000000000)
        XLColumn = 0
        Alignment = taCenter
        AlignToBand = False
        AutoSize = False
        Caption = '805,000'#50896
        Color = clWhite
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #44404#47548#52404
        Font.Style = []
        ParentFont = False
        Transparent = False
        ExportAs = exptText
        WrapStyle = BreakOnSpaces
        FontSize = 10
      end
      object qrlb_MedDate_Day: TQRLabel
        Left = 801
        Top = 687
        Width = 25
        Height = 21
        Size.Values = (
          55.562500000000000000
          2119.312500000000000000
          1817.687500000000000000
          66.145833333333300000)
        XLColumn = 0
        Alignment = taLeftJustify
        AlignToBand = False
        AutoSize = False
        Caption = '30'
        Color = clWhite
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -19
        Font.Name = #44404#47548
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = False
        ExportAs = exptText
        WrapStyle = BreakOnSpaces
        FontSize = 14
      end
      object qrlb_MedDate_Hour: TQRLabel
        Left = 969
        Top = 687
        Width = 25
        Height = 21
        Size.Values = (
          55.562500000000000000
          2563.812500000000000000
          1817.687500000000000000
          66.145833333333300000)
        XLColumn = 0
        Alignment = taLeftJustify
        AlignToBand = False
        AutoSize = False
        Caption = '09'
        Color = clWhite
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -19
        Font.Name = #44404#47548
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = False
        ExportAs = exptText
        WrapStyle = BreakOnSpaces
        FontSize = 14
      end
      object qrlb_MedDate_Min: TQRLabel
        Left = 1015
        Top = 687
        Width = 25
        Height = 21
        Size.Values = (
          55.562500000000000000
          2685.520833333330000000
          1817.687500000000000000
          66.145833333333300000)
        XLColumn = 0
        Alignment = taLeftJustify
        AlignToBand = False
        AutoSize = False
        Caption = '30'
        Color = clWhite
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -19
        Font.Name = #44404#47548
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = False
        ExportAs = exptText
        WrapStyle = BreakOnSpaces
        FontSize = 14
      end
      object qrlb_TotAmt2: TQRLabel
        Left = 597
        Top = 594
        Width = 93
        Height = 15
        Size.Values = (
          39.687500000000000000
          1579.562500000000000000
          1571.625000000000000000
          246.062500000000000000)
        XLColumn = 0
        Alignment = taCenter
        AlignToBand = False
        AutoSize = False
        Caption = '3,500,000'#50896
        Color = clWhite
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #44404#47548#52404
        Font.Style = []
        ParentFont = False
        Transparent = False
        ExportAs = exptText
        WrapStyle = BreakOnSpaces
        FontSize = 10
      end
      object qrlb_ImgCode: TQRLabel
        Left = 186
        Top = 21
        Width = 112
        Height = 19
        Size.Values = (
          50.270833333333300000
          492.125000000000000000
          55.562500000000000000
          296.333333333333000000)
        XLColumn = 0
        Alignment = taLeftJustify
        AlignToBand = False
        Caption = 'qrlb_ImgCode'
        Color = clWhite
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -17
        Font.Name = #44404#47548
        Font.Style = [fsBold]
        ParentFont = False
        Transparent = False
        ExportAs = exptText
        WrapStyle = BreakOnSpaces
        FontSize = 13
      end
      object qrlb_DutyUser: TQRLabel
        Left = 124
        Top = 883
        Width = 49
        Height = 17
        Size.Values = (
          44.979166666666700000
          328.083333333333000000
          2336.270833333330000000
          129.645833333333000000)
        XLColumn = 0
        Alignment = taLeftJustify
        AlignToBand = False
        Caption = #51060' '#49464' '#54616
        Color = clWhite
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = #44404#47548#52404
        Font.Style = []
        ParentFont = False
        Transparent = False
        ExportAs = exptText
        WrapStyle = BreakOnSpaces
        FontSize = 9
      end
      object qrlb_Manager: TQRLabel
        Left = 326
        Top = 883
        Width = 49
        Height = 17
        Size.Values = (
          44.979166666666700000
          862.541666666667000000
          2336.270833333330000000
          129.645833333333000000)
        XLColumn = 0
        Alignment = taLeftJustify
        AlignToBand = False
        Caption = #53428' '#44592' '#54364
        Color = clWhite
        Font.Charset = HANGEUL_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = #44404#47548#52404
        Font.Style = []
        ParentFont = False
        Transparent = False
        ExportAs = exptText
        WrapStyle = BreakOnSpaces
        FontSize = 9
      end
    end
  end
end
