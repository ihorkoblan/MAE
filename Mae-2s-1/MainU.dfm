object fmMain: TfmMain
  Left = 230
  Top = 68
  Width = 992
  Height = 608
  Caption = #1052#1040#1045' '#1082#1072#1083#1110#1073#1088#1091#1074#1072#1085#1085#1103
  Color = clBtnFace
  Font.Charset = RUSSIAN_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'Tahoma'
  Font.Style = []
  Menu = mm1
  OldCreateOrder = False
  Position = poDesktopCenter
  OnClose = FormClose
  PixelsPerInch = 96
  TextHeight = 16
  object StatusBar1: TStatusBar
    Left = 0
    Top = 530
    Width = 976
    Height = 19
    Panels = <>
    SimplePanel = True
    SimpleText = #1042#1077#1088#1089#1110#1103' (4 Slavko) '#1074#1110#1076' 28 '#1075#1088#1091#1076#1085#1103' 2009 '#1088'.'
  end
  object Chart1: TChart
    Left = 0
    Top = 209
    Width = 976
    Height = 321
    BackWall.Brush.Color = clWhite
    BackWall.Brush.Style = bsClear
    Title.Font.Charset = DEFAULT_CHARSET
    Title.Font.Color = clBlue
    Title.Font.Height = -13
    Title.Font.Name = 'Arial'
    Title.Font.Style = []
    Title.Text.Strings = (
      #1061#1074#1080#1083#1100#1086#1074#1077' '#1074#1110#1076#1086#1073#1088#1072#1078#1077#1085#1085#1103)
    BottomAxis.Automatic = False
    BottomAxis.AutomaticMinimum = False
    BottomAxis.AxisValuesFormat = '###0.###'
    BottomAxis.LabelsMultiLine = True
    BottomAxis.LabelStyle = talValue
    LeftAxis.AxisValuesFormat = '###0.###'
    LeftAxis.LabelStyle = talValue
    Legend.Visible = False
    View3D = False
    View3DWalls = False
    Align = alClient
    BevelInner = bvSpace
    BevelOuter = bvNone
    TabOrder = 1
    object pnlMark: TPanel
      Left = 8
      Top = 8
      Width = 25
      Height = 25
      BevelOuter = bvNone
      TabOrder = 0
    end
    object Series1: TLineSeries
      Marks.ArrowLength = 8
      Marks.Visible = False
      SeriesColor = clBlue
      ValueFormat = '###0.###'
      Pointer.InflateMargins = True
      Pointer.Style = psRectangle
      Pointer.Visible = False
      XValues.DateTime = False
      XValues.Name = 'X'
      XValues.Multiplier = 1.000000000000000000
      XValues.Order = loAscending
      YValues.DateTime = False
      YValues.Name = 'Y'
      YValues.Multiplier = 1.000000000000000000
      YValues.Order = loNone
    end
    object Series2: TPointSeries
      Marks.Arrow.Color = clBlack
      Marks.ArrowLength = 15
      Marks.Visible = True
      SeriesColor = clGreen
      ValueFormat = '###0.###'
      Pointer.Brush.Color = clRed
      Pointer.InflateMargins = False
      Pointer.Pen.Color = clRed
      Pointer.Pen.Visible = False
      Pointer.Style = psDownTriangle
      Pointer.Visible = True
      XValues.DateTime = False
      XValues.Name = 'X'
      XValues.Multiplier = 1.000000000000000000
      XValues.Order = loAscending
      YValues.DateTime = False
      YValues.Name = 'Y'
      YValues.Multiplier = 1.000000000000000000
      YValues.Order = loNone
    end
    object Series3: TLineSeries
      Marks.ArrowLength = 8
      Marks.Visible = False
      SeriesColor = clFuchsia
      ShowInLegend = False
      LinePen.Width = 2
      Pointer.InflateMargins = True
      Pointer.Style = psRectangle
      Pointer.Visible = False
      XValues.DateTime = False
      XValues.Name = 'X'
      XValues.Multiplier = 1.000000000000000000
      XValues.Order = loAscending
      YValues.DateTime = False
      YValues.Name = 'Y'
      YValues.Multiplier = 1.000000000000000000
      YValues.Order = loNone
    end
    object Series4: TLineSeries
      Marks.ArrowLength = 8
      Marks.Visible = False
      SeriesColor = clRed
      Pointer.InflateMargins = True
      Pointer.Style = psRectangle
      Pointer.Visible = False
      XValues.DateTime = False
      XValues.Name = 'X'
      XValues.Multiplier = 1.000000000000000000
      XValues.Order = loAscending
      YValues.DateTime = False
      YValues.Name = 'Y'
      YValues.Multiplier = 1.000000000000000000
      YValues.Order = loNone
    end
  end
  object PageControl1: TPageControl
    Left = 0
    Top = 0
    Width = 976
    Height = 209
    ActivePage = TabSheet1
    Align = alTop
    Style = tsFlatButtons
    TabOrder = 2
    object TabSheet1: TTabSheet
      Caption = ' '#1053#1072#1083#1072#1096#1090#1091#1074#1072#1085#1085#1103' '
      object gbxSettings: TGroupBox
        Left = 224
        Top = 8
        Width = 281
        Height = 161
        Caption = ' '#1053#1072#1083#1072#1096#1090#1091#1074#1072#1085#1085#1103' '
        TabOrder = 0
        object Label1: TLabel
          Left = 16
          Top = 22
          Width = 101
          Height = 16
          Caption = #1055#1086#1090#1086#1095#1085#1072' '#1074#1080#1073#1086#1088#1082#1072
        end
        object Label2: TLabel
          Left = 16
          Top = 55
          Width = 187
          Height = 16
          Caption = #1064#1091#1082#1072#1090#1080' '#1084#1072#1082#1089#1080#1084#1091#1084' '#1074#1110#1076'            '#1076#1086
        end
        object Label7: TLabel
          Left = 16
          Top = 82
          Width = 153
          Height = 28
          Caption = #1055#1086#1088#1086#1075#1086#1074#1077' '#1079#1085#1072#1095#1077#1085#1085#1103' '#1076#1083#1103' '#1086#1073#1095#1080#1089#1083#1077#1085#1085#1103' '#1089#1091#1084#1080' '#1072#1084#1087#1083#1110#1090#1091#1076
          Font.Charset = RUSSIAN_CHARSET
          Font.Color = clWindowText
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
          WordWrap = True
        end
        object speSampleNo: TSpinEdit
          Left = 204
          Top = 18
          Width = 49
          Height = 26
          EditorEnabled = False
          Enabled = False
          MaxValue = 1
          MinValue = 1
          TabOrder = 0
          Value = 1
          OnChange = speSampleNoChange
        end
        object edEndPoint: TEdit
          Left = 204
          Top = 52
          Width = 49
          Height = 24
          Enabled = False
          TabOrder = 2
          Text = '20000'
          OnExit = edEndPointExit
        end
        object edStartPoint: TEdit
          Left = 142
          Top = 52
          Width = 45
          Height = 24
          Enabled = False
          TabOrder = 1
          Text = '1'
          OnExit = edStartPointExit
        end
        object edThreshold: TSpinEdit
          Left = 204
          Top = 84
          Width = 50
          Height = 26
          EditorEnabled = False
          Enabled = False
          MaxValue = 100
          MinValue = 1
          TabOrder = 3
          Value = 30
          OnChange = edThresholdChange
        end
        object ButtonShow: TButton
          Left = 136
          Top = 120
          Width = 113
          Height = 33
          Caption = #1042#1110#1076#1086#1073#1088#1072#1079#1080#1090#1080
          Enabled = False
          TabOrder = 4
          OnClick = ButtonShowClick
        end
        object ButtonCalcAll: TButton
          Left = 8
          Top = 120
          Width = 113
          Height = 33
          Caption = #1054#1073#1095#1080#1089#1083#1080#1090#1080
          Enabled = False
          TabOrder = 5
          OnClick = ButtonCalcAllClick
        end
      end
      object gbxExperements: TGroupBox
        Left = 8
        Top = 0
        Width = 209
        Height = 169
        Caption = #1045#1082#1089#1087#1077#1088#1077#1084#1077#1085#1090#1080
        TabOrder = 1
        object ListBoxExperements: TListBox
          Left = 8
          Top = 64
          Width = 185
          Height = 97
          ItemHeight = 16
          TabOrder = 0
          OnClick = ListBoxExperementsClick
        end
        object CheckBoxReadNewFormat: TCheckBox
          Left = 8
          Top = 16
          Width = 113
          Height = 17
          Caption = #1053#1086#1074#1080#1081' '#1092#1086#1088#1084#1072#1090
          Checked = True
          State = cbChecked
          TabOrder = 1
        end
      end
      object grp1: TGroupBox
        Left = 512
        Top = 8
        Width = 449
        Height = 161
        Caption = 'grp1'
        TabOrder = 2
        object lbl1: TLabel
          Left = 216
          Top = 16
          Width = 154
          Height = 16
          Caption = #1064#1080#1088#1080#1085#1072' '#1082#1086#1074#1079#1072#1102#1095#1086#1075#1086' '#1074#1110#1082#1085#1072
        end
        object chklst1: TCheckListBox
          Left = 16
          Top = 16
          Width = 193
          Height = 105
          Columns = 1
          ItemHeight = 16
          Items.Strings = (
            #1043#1088#1072#1092#1110#1082' MAE'
            #1050#1086#1074#1079#1085#1077' '#1089#1077#1088#1077#1076#1085#1108' '#1082#1074#1072#1076#1088#1072#1090#1080#1095#1085#1077)
          TabOrder = 0
        end
        object movingAvgWidthSpinEdit: TSpinEdit
          Left = 220
          Top = 36
          Width = 50
          Height = 26
          EditorEnabled = False
          MaxValue = 300
          MinValue = 1
          TabOrder = 1
          Value = 50
          OnChange = movingAvgWidthSpinEditChange
        end
      end
    end
    object TabSheet2: TTabSheet
      Caption = ' '#1056#1077#1079#1091#1083#1100#1090#1072#1090#1080' '
      ImageIndex = 1
      object gbxCurrentDataSample: TGroupBox
        Left = 8
        Top = 2
        Width = 257
        Height = 119
        Caption = ' '#1055#1086#1090#1086#1095#1085#1072' '#1074#1080#1073#1086#1088#1082#1072' '
        TabOrder = 0
        object lb11: TLabel
          Left = 16
          Top = 27
          Width = 4
          Height = 16
        end
        object lb12: TLabel
          Left = 16
          Top = 43
          Width = 4
          Height = 16
        end
        object lb13: TLabel
          Left = 16
          Top = 59
          Width = 4
          Height = 16
        end
        object lb14: TLabel
          Left = 16
          Top = 75
          Width = 4
          Height = 16
        end
      end
      object gbxMeanValues: TGroupBox
        Left = 272
        Top = 2
        Width = 257
        Height = 119
        Caption = ' '#1057#1077#1088#1077#1076#1085#1108' '#1079#1085#1072#1095#1077#1085#1085#1103'... '
        TabOrder = 1
        object lb22: TLabel
          Left = 16
          Top = 43
          Width = 4
          Height = 16
        end
        object lb23: TLabel
          Left = 16
          Top = 59
          Width = 4
          Height = 16
        end
        object lb24: TLabel
          Left = 16
          Top = 75
          Width = 4
          Height = 16
        end
        object lb21: TLabel
          Left = 16
          Top = 27
          Width = 4
          Height = 16
        end
      end
      object GroupBox1: TGroupBox
        Left = 536
        Top = 2
        Width = 329
        Height = 119
        Caption = ' '#1044#1086#1074#1110#1088#1085#1080#1081' '#1110#1085#1090#1077#1088#1074#1072#1083' '
        TabOrder = 2
        object lb32: TLabel
          Left = 16
          Top = 43
          Width = 4
          Height = 16
        end
        object lb33: TLabel
          Left = 16
          Top = 59
          Width = 4
          Height = 16
        end
        object lb34: TLabel
          Left = 16
          Top = 75
          Width = 4
          Height = 16
        end
        object lb31: TLabel
          Left = 16
          Top = 27
          Width = 4
          Height = 16
        end
      end
    end
    object TabSheet3: TTabSheet
      Caption = 'TabSheet3'
      ImageIndex = 2
      object btn1: TButton
        Left = 16
        Top = 16
        Width = 265
        Height = 25
        Caption = #1054#1073#1095#1080#1089#1083#1080#1090#1080' '#1082#1086#1074#1079#1085#1077' '#1089#1077#1088#1077#1076#1085#1108
        TabOrder = 0
        OnClick = btn1Click
      end
      object btn2: TButton
        Left = 16
        Top = 56
        Width = 265
        Height = 25
        Caption = #1054#1073#1095#1080#1089#1083#1080#1090#1080' '#1082#1086#1074#1079#1085#1077' '#1089#1077#1088#1077#1076#1085#1108' '#1082#1074#1072#1076#1088#1072#1090#1080#1095#1085#1077
        TabOrder = 1
        OnClick = btn2Click
      end
    end
  end
  object OpnDlg: TOpenDialog
    Filter = #1060#1072#1081#1083#1080' '#1076#1072#1085#1080#1093'|*.dat|'#1058#1077#1082#1089#1090#1086#1074#1110' '#1092#1072#1081#1083#1080'|*.txt'
    Options = [ofAllowMultiSelect, ofEnableSizing]
    Title = #1042#1082#1072#1078#1110#1090#1100' '#1092#1072#1081#1083' '#1076#1083#1103' '#1086#1073#1088#1086#1073#1082#1080
    Left = 1328
    Top = 48
  end
  object actlMain: TActionList
    Left = 1328
    Top = 16
    object actLoadFromFile: TAction
      Caption = #1042#1110#1076#1082#1088#1080#1090#1080' '#1092#1072#1081#1083
      ShortCut = 16463
    end
    object actExportAll: TAction
      Caption = #1045#1082#1089#1087#1086#1088#1090' '#1088#1077#1079#1091#1083#1100#1090#1072#1090#1110#1074
      Enabled = False
      ShortCut = 16453
    end
    object actExportOne: TAction
      Caption = #1045#1082#1089#1087#1086#1088#1090' '#1087#1086#1090#1086#1095#1085#1086#1111' '#1074#1080#1073#1086#1088#1082#1080
      ShortCut = 24645
    end
    object actCalc: TAction
      Caption = #1055#1088#1086#1074#1077#1089#1090#1080' '#1086#1073#1095#1080#1089#1083#1077#1085#1085#1103
      Enabled = False
      ShortCut = 16459
    end
    object actExportToTXT: TAction
      Caption = #1045#1082#1089#1087#1086#1088#1090' '#1088#1077#1079#1091#1083#1100#1090#1072#1090#1110#1074' (TXT)'
    end
  end
  object ppmExport: TPopupMenu
    Left = 1328
    Top = 80
    object pmiExportAll: TMenuItem
      Action = actExportAll
      Caption = #1045#1082#1089#1087#1086#1088#1090' '#1074#1089#1110#1093' '#1074#1080#1073#1086#1088#1086#1082
    end
    object pmiExportOne: TMenuItem
      Action = actExportOne
    end
  end
  object mm1: TMainMenu
    Left = 60
    Top = 246
    object N1: TMenuItem
      Caption = 'File'
      object Open1: TMenuItem
        Caption = 'Open'
        OnClick = Open1Click
      end
      object Export1: TMenuItem
        Caption = 'Export'
        object ExportsampleExcell1: TMenuItem
          Caption = 'Export sample to Excell'
          OnClick = ExportsampleExcell1Click
        end
        object All1: TMenuItem
          Caption = 'Export all samples to Excel'
          OnClick = All1Click
        end
        object ExportallexperimentstoTXT1: TMenuItem
          Caption = 'Export all experiments to TXT'
          OnClick = ExportallexperimentstoTXT1Click
        end
        object ExportallexperimentstoExscel1: TMenuItem
          Caption = 'Export all experiments to Excel'
          OnClick = ExportallexperimentstoExscel1Click
        end
      end
      object Delete1: TMenuItem
        Caption = 'Delete'
      end
      object Deleteall1: TMenuItem
        Caption = 'Delete all'
        OnClick = Deleteall1Click
      end
      object Exit1: TMenuItem
        Caption = 'Exit'
      end
    end
    object Edit1: TMenuItem
      Caption = 'Edit'
    end
  end
end
