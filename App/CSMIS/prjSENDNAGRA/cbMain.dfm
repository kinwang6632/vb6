object fmMain: TfmMain
  Left = 388
  Top = 192
  Width = 640
  Height = 539
  Caption = 'SEND_NAGRA '#23567#31243#24335
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 14
  object pnTitle: TPanel
    Left = 0
    Top = 0
    Width = 632
    Height = 42
    Align = alTop
    BevelOuter = bvNone
    Color = clBtnShadow
    Font.Charset = ANSI_CHARSET
    Font.Color = clWhite
    Font.Height = -21
    Font.Name = 'Arial'
    Font.Style = []
    ParentBackground = False
    ParentFont = False
    TabOrder = 0
    object Label55: TLabel
      Left = 17
      Top = 14
      Width = 583
      Height = 17
      Caption = #27161#31034#28858#32005#33394#30340#23383', '#30342#28858#24517#22635#27396#20301', '#35531#30906#23450#36664#20837#24460', '#20877#25353#19979#25353#37397#23531#36914' SEND_NAGRA Table'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWhite
      Font.Height = -15
      Font.Name = 'Arial'
      Font.Style = []
      ParentFont = False
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 42
    Width = 190
    Height = 463
    Align = alLeft
    BevelOuter = bvNone
    TabOrder = 1
    object cxGroupBox2: TcxGroupBox
      Left = 0
      Top = 0
      Align = alTop
      Caption = 'Oracle '#36039#26009#24235
      TabOrder = 0
      Height = 203
      Width = 190
      object lblConnectStatus: TLabel
        Left = 17
        Top = 173
        Width = 34
        Height = 16
        Caption = #38626#32218
        Font.Charset = ANSI_CHARSET
        Font.Color = clMaroon
        Font.Height = -13
        Font.Name = 'Verdana'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object Label52: TLabel
        Left = 17
        Top = 21
        Width = 62
        Height = 14
        Caption = 'TNS Name:'
        Font.Charset = ANSI_CHARSET
        Font.Color = clRed
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
      end
      object Label53: TLabel
        Left = 18
        Top = 68
        Width = 63
        Height = 14
        Caption = 'User Name:'
        Font.Charset = ANSI_CHARSET
        Font.Color = clRed
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
      end
      object Label54: TLabel
        Left = 18
        Top = 112
        Width = 55
        Height = 14
        Caption = 'Password:'
        Font.Charset = ANSI_CHARSET
        Font.Color = clRed
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
      end
      object btnConnect: TcxButton
        Left = 100
        Top = 170
        Width = 80
        Height = 25
        Caption = #36899#32218
        TabOrder = 0
        OnClick = btnConnectClick
        Glyph.Data = {
          36040000424D3604000000000000360000002800000010000000100000000100
          2000000000000004000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          00000000000006151D310D4154AC050A0D230000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          000000000000687D84C390BDEFFF2082ADFF070C102800000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          00000000000006151D3192C6F0FF8287E1FF1D789FFF09111633000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          00000000000000000000557480B576D1FAFF8287E1FF1F6F95FF0910142E0000
          0000000000000000000000000000000000000000000000000000000000000000
          000000000000000000000507080C749CAAE376D1FAFF8287E1FF22688CFF0C16
          1C3D000000000000000000000000000000000000000000000000000000000000
          000000000000558CA3F448788CF43C626FF5567C86F976D1FAFF8287E1FF294F
          6BFF0B1419330000000000000000000000000000000000000000000000000000
          00000000000059C8E7FF7FE4FFFF7FE4FFFF94CCF2FF90BDEFFF81A0E9FF8287
          E1FF30475EFF090F132900000000000000000000000000000000000000000000
          00000000000029454C5559C8E7FFD9F2FEFF6EC0F5FF8287E1FF458A9FFE0E14
          1539010101030000000000000000000000000000000000000000000000000000
          0000000000000000000059C8E7FFB0E4FCFF76D1FAFF7FB0EAFF8287E1FF4A7A
          88FF0A0D0E270000000000000000000000000000000000000000000000000000
          0000000000000000000021363B4259C8E7FFD9F2FEFF76D1FAFF7EAAE9FF8287
          E1FF48626AFF0A0D0E2700000000000000000000000000000000000000000000
          000000000000000000000000000059C8E7FFB0E4FCFF7FE4FFFF89C2EEFF7CA5
          E8FF8287E1FF4B5358FF0304040C000000000000000000000000000000000000
          00000000000000000000000000001D2E33398ECCDDFF89D0E7FF89CFE6FF82C5
          DDFF78B2C8FF6A97ACFF14344272000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000}
      end
      object txtDbUserId: TcxTextEdit
        Left = 26
        Top = 87
        Properties.MaxLength = 0
        TabOrder = 1
        Width = 120
      end
      object txtDbPassword: TcxTextEdit
        Left = 26
        Top = 132
        Properties.MaxLength = 0
        TabOrder = 2
        Width = 120
      end
      object cmbTNS: TcxComboBox
        Left = 26
        Top = 41
        Properties.DropDownListStyle = lsFixedList
        TabOrder = 3
        Width = 154
      end
    end
    object cxGroupBox5: TcxGroupBox
      Left = 0
      Top = 203
      Align = alTop
      Caption = #27231#19978#30418#33287#26234#24935#21345
      TabOrder = 1
      Height = 227
      Width = 190
      object Label1: TLabel
        Left = 17
        Top = 25
        Width = 64
        Height = 14
        Caption = #27231#19978#30418#24207#34399':'
        Font.Charset = ANSI_CHARSET
        Font.Color = clRed
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
      end
      object Label9: TLabel
        Left = 17
        Top = 72
        Width = 64
        Height = 14
        Caption = #26234#24935#21345#21345#34399':'
        Font.Charset = ANSI_CHARSET
        Font.Color = clRed
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
      end
      object Label2: TLabel
        Left = 17
        Top = 120
        Width = 76
        Height = 14
        Caption = #33290#27231#19978#30418#24207#34399':'
        Font.Charset = ANSI_CHARSET
        Font.Color = clBlue
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
      end
      object Label3: TLabel
        Left = 17
        Top = 167
        Width = 76
        Height = 14
        Caption = #33290#26234#24935#21345#21345#34399':'
        Font.Charset = ANSI_CHARSET
        Font.Color = clBlue
        Font.Height = -12
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
      end
      object txtNewSmartCard: TcxTextEdit
        Left = 32
        Top = 92
        Properties.MaxLength = 10
        TabOrder = 1
        Width = 120
      end
      object txtNewStepBox: TcxTextEdit
        Left = 32
        Top = 45
        Properties.MaxLength = 14
        TabOrder = 0
        Width = 120
      end
      object txtOldStepBox: TcxTextEdit
        Left = 32
        Top = 140
        Properties.MaxLength = 14
        TabOrder = 2
        Width = 120
      end
      object txtOldSmartCard: TcxTextEdit
        Left = 32
        Top = 187
        Properties.MaxLength = 10
        TabOrder = 3
        Width = 120
      end
    end
  end
  object pnClient: TPanel
    Left = 190
    Top = 42
    Width = 442
    Height = 463
    Align = alClient
    TabOrder = 2
    object Label56: TLabel
      Left = 24
      Top = 268
      Width = 52
      Height = 14
      Caption = #22519#34892#32080#26524':'
    end
    object lbSendResult: TLabel
      Left = 88
      Top = 268
      Width = 21
      Height = 14
      Caption = 'N/A'
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlue
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object Label57: TLabel
      Left = 24
      Top = 300
      Width = 52
      Height = 14
      Caption = #37679#35492#20195#30908':'
    end
    object lbErrCode: TLabel
      Left = 88
      Top = 300
      Width = 21
      Height = 14
      Caption = 'N/A'
      Font.Charset = ANSI_CHARSET
      Font.Color = clRed
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object Label58: TLabel
      Left = 24
      Top = 328
      Width = 52
      Height = 14
      Caption = #37679#35492#35338#24687':'
    end
    object lbErrMsg: TLabel
      Left = 88
      Top = 330
      Width = 21
      Height = 14
      Caption = 'N/A'
      Font.Charset = ANSI_CHARSET
      Font.Color = clRed
      Font.Height = -12
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object CommandPage: TcxPageControl
      Left = 1
      Top = 1
      Width = 440
      Height = 172
      ActivePage = B5
      Align = alTop
      TabOrder = 0
      OnChange = CommandPageChange
      ClientRectBottom = 170
      ClientRectLeft = 2
      ClientRectRight = 438
      ClientRectTop = 23
      object A1: TcxTabSheet
        Caption = 'A1.'#38283#27231
        ImageIndex = 0
        object Label13: TLabel
          Left = 21
          Top = 18
          Width = 52
          Height = 14
          Caption = 'Zip Code:'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label14: TLabel
          Left = 166
          Top = 18
          Width = 55
          Height = 14
          Caption = 'BAT '#21312#30908':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object txtA1ZipCode: TcxTextEdit
          Left = 79
          Top = 15
          Properties.MaxLength = 5
          TabOrder = 0
          Text = '0'
          Width = 65
        end
        object txtA1BatId: TcxTextEdit
          Left = 227
          Top = 15
          Properties.MaxLength = 5
          TabOrder = 1
          Text = '0'
          Width = 65
        end
      end
      object A2: TcxTabSheet
        Caption = 'A2.'#38364#27231
        ImageIndex = 1
        object lbNoParamNeeded: TLabel
          Left = 21
          Top = 30
          Width = 156
          Height = 14
          Caption = #36889#20491#25351#20196#19981#38920#35201#36664#20837#20854#23427#21443#25976
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
      end
      object A3: TcxTabSheet
        Caption = 'A3.'#20572#27231
        ImageIndex = 2
      end
      object A4: TcxTabSheet
        Caption = 'A4.'#24489#27231
        ImageIndex = 3
      end
      object A6: TcxTabSheet
        Caption = 'A6.'#32173#20462#25563#25286'('#25972#32068')'
        ImageIndex = 4
        DesignSize = (
          436
          147)
        object Label4: TLabel
          Left = 21
          Top = 18
          Width = 52
          Height = 14
          Caption = 'Zip Code:'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label5: TLabel
          Left = 162
          Top = 18
          Width = 55
          Height = 14
          Caption = 'BAT '#21312#30908':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label6: TLabel
          Left = 21
          Top = 46
          Width = 91
          Height = 14
          Caption = 'NAGRA'#38971#36947#20195#30908':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label7: TLabel
          Left = 22
          Top = 67
          Width = 228
          Height = 14
          Caption = '( '#22810#32068#38971#36947#35531#29992#36887#40670#38548#38283', '#31684#20363': 10,20,30 )'
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label8: TLabel
          Left = 21
          Top = 91
          Width = 64
          Height = 14
          Caption = #25480#27402#36215#22987#26085':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label10: TLabel
          Left = 209
          Top = 91
          Width = 64
          Height = 14
          Caption = #25480#27402#25130#27490#26085':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object txtA6ZipCode: TcxTextEdit
          Left = 79
          Top = 15
          Properties.MaxLength = 5
          TabOrder = 0
          Text = '0'
          Width = 65
        end
        object txtA6BatId: TcxTextEdit
          Left = 223
          Top = 15
          Properties.MaxLength = 5
          TabOrder = 1
          Text = '0'
          Width = 65
        end
        object txtA6ProductId: TcxTextEdit
          Left = 119
          Top = 43
          Anchors = [akLeft, akTop, akRight]
          Properties.MaxLength = 0
          TabOrder = 2
          Text = '0'
          Width = 287
        end
        object txtA6StartDate: TcxMaskEdit
          Left = 91
          Top = 87
          Properties.MaskKind = emkRegExprEx
          Properties.EditMask = 
            '([123][0-9])? [0-9][0-9] / (0?[1-9] | 1[012]) / ([012]?[1-9] | [' +
            '123]0 |31)'
          Properties.MaxLength = 0
          TabOrder = 3
          Text = '2009/10/15'
          Width = 98
        end
        object txtA6EndDate: TcxMaskEdit
          Left = 279
          Top = 87
          Properties.MaskKind = emkRegExprEx
          Properties.EditMask = 
            '([123][0-9])? [0-9][0-9] / (0?[1-9] | 1[012]) / ([012]?[1-9] | [' +
            '123]0 |31)'
          Properties.MaxLength = 0
          TabOrder = 4
          Text = '2009/10/15'
          Width = 98
        end
      end
      object A8: TcxTabSheet
        Caption = 'A8.'#32173#20462#25563#25286'('#25563'STB)'
        ImageIndex = 6
        object Label11: TLabel
          Left = 26
          Top = 18
          Width = 55
          Height = 14
          Caption = 'BAT '#21312#30908':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object txtA8BatId: TcxTextEdit
          Left = 87
          Top = 15
          Properties.MaxLength = 5
          TabOrder = 0
          Text = '0'
          Width = 65
        end
      end
      object A9: TcxTabSheet
        Caption = 'A9.'#32173#20462#25563#25286'(ICC)'
        ImageIndex = 7
        DesignSize = (
          436
          147)
        object Label12: TLabel
          Left = 21
          Top = 18
          Width = 52
          Height = 14
          Caption = 'Zip Code:'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label15: TLabel
          Left = 21
          Top = 46
          Width = 91
          Height = 14
          Caption = 'NAGRA'#38971#36947#20195#30908':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label16: TLabel
          Left = 22
          Top = 67
          Width = 228
          Height = 14
          Caption = '( '#22810#32068#38971#36947#35531#29992#36887#40670#38548#38283', '#31684#20363': 10,20,30 )'
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label17: TLabel
          Left = 21
          Top = 91
          Width = 64
          Height = 14
          Caption = #25480#27402#36215#22987#26085':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label18: TLabel
          Left = 209
          Top = 91
          Width = 64
          Height = 14
          Caption = #25480#27402#25130#27490#26085':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object txtA9ZipCode: TcxTextEdit
          Left = 79
          Top = 15
          Properties.MaxLength = 5
          TabOrder = 0
          Text = '0'
          Width = 65
        end
        object txtA9ProductId: TcxTextEdit
          Left = 118
          Top = 43
          Anchors = [akLeft, akTop, akRight]
          Properties.MaxLength = 0
          TabOrder = 1
          Text = '0'
          Width = 287
        end
        object txtA9StartDate: TcxMaskEdit
          Left = 91
          Top = 87
          Properties.MaskKind = emkRegExprEx
          Properties.EditMask = 
            '([123][0-9])? [0-9][0-9] / (0?[1-9] | 1[012]) / ([012]?[1-9] | [' +
            '123]0 |31)'
          Properties.MaxLength = 0
          TabOrder = 2
          Text = '2009/10/15'
          Width = 98
        end
        object txtA9EndDate: TcxMaskEdit
          Left = 279
          Top = 87
          Properties.MaskKind = emkRegExprEx
          Properties.EditMask = 
            '([123][0-9])? [0-9][0-9] / (0?[1-9] | 1[012]) / ([012]?[1-9] | [' +
            '123]0 |31)'
          Properties.MaxLength = 0
          TabOrder = 3
          Text = '2009/10/15'
          Width = 98
        end
      end
      object B1: TcxTabSheet
        Caption = 'B1.'#38283#38971#36947
        ImageIndex = 8
        DesignSize = (
          436
          147)
        object Label19: TLabel
          Left = 21
          Top = 46
          Width = 91
          Height = 14
          Caption = 'NAGRA'#38971#36947#20195#30908':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label20: TLabel
          Left = 22
          Top = 67
          Width = 228
          Height = 14
          Caption = '( '#22810#32068#38971#36947#35531#29992#36887#40670#38548#38283', '#31684#20363': 10,20,30 )'
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label21: TLabel
          Left = 21
          Top = 87
          Width = 64
          Height = 14
          Caption = #25480#27402#36215#22987#26085':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label22: TLabel
          Left = 209
          Top = 91
          Width = 64
          Height = 14
          Caption = #25480#27402#25130#27490#26085':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object txtB1ProductId: TcxTextEdit
          Left = 119
          Top = 43
          Anchors = [akLeft, akTop, akRight]
          Properties.MaxLength = 0
          TabOrder = 0
          Text = '0'
          Width = 287
        end
        object txtB1StartDate: TcxMaskEdit
          Left = 91
          Top = 87
          Properties.MaskKind = emkRegExprEx
          Properties.EditMask = 
            '([123][0-9])? [0-9][0-9] / (0?[1-9] | 1[012]) / ([012]?[1-9] | [' +
            '123]0 |31)'
          Properties.MaxLength = 0
          TabOrder = 1
          Text = '2009/10/15'
          Width = 98
        end
        object txtB1EndDate: TcxMaskEdit
          Left = 279
          Top = 87
          Properties.MaskKind = emkRegExprEx
          Properties.EditMask = 
            '([123][0-9])? [0-9][0-9] / (0?[1-9] | 1[012]) / ([012]?[1-9] | [' +
            '123]0 |31)'
          Properties.MaxLength = 0
          TabOrder = 2
          Text = '2009/10/15'
          Width = 98
        end
      end
      object B2: TcxTabSheet
        Caption = 'B2.'#38364#38971#36947
        ImageIndex = 9
        DesignSize = (
          436
          147)
        object Label23: TLabel
          Left = 21
          Top = 46
          Width = 91
          Height = 14
          Caption = 'NAGRA'#38971#36947#20195#30908':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label24: TLabel
          Left = 22
          Top = 67
          Width = 228
          Height = 14
          Caption = '( '#22810#32068#38971#36947#35531#29992#36887#40670#38548#38283', '#31684#20363': 10,20,30 )'
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object txtB2ProductId: TcxTextEdit
          Left = 119
          Top = 43
          Anchors = [akLeft, akTop, akRight]
          Properties.MaxLength = 0
          TabOrder = 0
          Text = '0'
          Width = 287
        end
      end
      object B3: TcxTabSheet
        Caption = 'B3.'#38364#25152#26377#38971#36947
        ImageIndex = 10
      end
      object B4: TcxTabSheet
        Caption = 'B4.'#26283#20572#38971#36947
        ImageIndex = 11
        DesignSize = (
          436
          147)
        object Label25: TLabel
          Left = 21
          Top = 46
          Width = 91
          Height = 14
          Caption = 'NAGRA'#38971#36947#20195#30908':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label26: TLabel
          Left = 22
          Top = 67
          Width = 228
          Height = 14
          Caption = '( '#22810#32068#38971#36947#35531#29992#36887#40670#38548#38283', '#31684#20363': 10,20,30 )'
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object txtB4ProductId: TcxTextEdit
          Left = 119
          Top = 43
          Anchors = [akLeft, akTop, akRight]
          Properties.MaxLength = 0
          TabOrder = 0
          Text = '0'
          Width = 287
        end
      end
      object B5: TcxTabSheet
        Caption = 'B5.'#24674#24489#38971#36947
        ImageIndex = 12
        DesignSize = (
          436
          147)
        object Label27: TLabel
          Left = 21
          Top = 46
          Width = 91
          Height = 14
          Caption = 'NAGRA'#38971#36947#20195#30908':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label28: TLabel
          Left = 22
          Top = 67
          Width = 228
          Height = 14
          Caption = '( '#22810#32068#38971#36947#35531#29992#36887#40670#38548#38283', '#31684#20363': 10,20,30 )'
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label59: TLabel
          Left = 21
          Top = 91
          Width = 64
          Height = 14
          Caption = #25480#27402#36215#22987#26085':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label60: TLabel
          Left = 209
          Top = 91
          Width = 64
          Height = 14
          Caption = #25480#27402#25130#27490#26085':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object txtB5ProductId: TcxTextEdit
          Left = 119
          Top = 43
          Anchors = [akLeft, akTop, akRight]
          Properties.MaxLength = 0
          TabOrder = 0
          Text = '0'
          Width = 287
        end
        object txtB5StartDate: TcxMaskEdit
          Left = 91
          Top = 87
          Properties.MaskKind = emkRegExprEx
          Properties.EditMask = 
            '([123][0-9])? [0-9][0-9] / (0?[1-9] | 1[012]) / ([012]?[1-9] | [' +
            '123]0 |31)'
          Properties.MaxLength = 0
          TabOrder = 1
          Text = '2009/10/15'
          Width = 98
        end
        object txtB5EndDate: TcxMaskEdit
          Left = 279
          Top = 87
          Properties.MaskKind = emkRegExprEx
          Properties.EditMask = 
            '([123][0-9])? [0-9][0-9] / (0?[1-9] | 1[012]) / ([012]?[1-9] | [' +
            '123]0 |31)'
          Properties.MaxLength = 0
          TabOrder = 2
          Text = '2009/10/15'
          Width = 98
        end
      end
      object B6: TcxTabSheet
        Caption = 'B6.'#24310#38263'/'#26356#25913#26399#38480
        ImageIndex = 13
        DesignSize = (
          436
          147)
        object Label29: TLabel
          Left = 21
          Top = 46
          Width = 91
          Height = 14
          Caption = 'NAGRA'#38971#36947#20195#30908':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label30: TLabel
          Left = 22
          Top = 67
          Width = 228
          Height = 14
          Caption = '( '#22810#32068#38971#36947#35531#29992#36887#40670#38548#38283', '#31684#20363': 10,20,30 )'
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label32: TLabel
          Left = 48
          Top = 91
          Width = 64
          Height = 14
          Caption = #25480#27402#25130#27490#26085':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object txtB6ProductId: TcxTextEdit
          Left = 118
          Top = 43
          Anchors = [akLeft, akTop, akRight]
          Properties.MaxLength = 0
          TabOrder = 0
          Text = '0'
          Width = 287
        end
        object txtB6EndDate: TcxMaskEdit
          Left = 118
          Top = 87
          Properties.MaskKind = emkRegExprEx
          Properties.EditMask = 
            '([123][0-9])? [0-9][0-9] / (0?[1-9] | 1[012]) / ([012]?[1-9] | [' +
            '123]0 |31)'
          Properties.MaxLength = 0
          TabOrder = 1
          Text = '2009/10/15'
          Width = 98
        end
      end
      object B7: TcxTabSheet
        Caption = 'B7.'#25910#35222#27402#38480#35079#35069
        ImageIndex = 14
        DesignSize = (
          436
          147)
        object Label31: TLabel
          Left = 21
          Top = 14
          Width = 91
          Height = 14
          Caption = 'NAGRA'#38971#36947#20195#30908':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label33: TLabel
          Left = 22
          Top = 35
          Width = 228
          Height = 14
          Caption = '( '#22810#32068#38971#36947#35531#29992#36887#40670#38548#38283', '#31684#20363': 10,20,30 )'
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label34: TLabel
          Left = 21
          Top = 59
          Width = 64
          Height = 14
          Caption = #25480#27402#36215#22987#26085':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label35: TLabel
          Left = 209
          Top = 59
          Width = 64
          Height = 14
          Caption = #25480#27402#25130#27490#26085':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label36: TLabel
          Left = 21
          Top = 90
          Width = 112
          Height = 14
          Caption = #35079#35069#21040#19979#21015#30340#26234#24935#21345':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label37: TLabel
          Left = 22
          Top = 115
          Width = 312
          Height = 14
          Caption = '( '#22810#32068#21345#34399#35531#29992#36887#40670#38548#38283', '#31684#20363': 100001,200001,300001 )'
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object txtB7ProductId: TcxTextEdit
          Left = 119
          Top = 11
          Anchors = [akLeft, akTop, akRight]
          Properties.MaxLength = 0
          TabOrder = 0
          Text = '0'
          Width = 287
        end
        object txtB7StartDate: TcxMaskEdit
          Left = 91
          Top = 55
          Properties.MaskKind = emkRegExprEx
          Properties.EditMask = 
            '([123][0-9])? [0-9][0-9] / (0?[1-9] | 1[012]) / ([012]?[1-9] | [' +
            '123]0 |31)'
          Properties.MaxLength = 0
          TabOrder = 1
          Text = '2009/10/15'
          Width = 98
        end
        object txtB7EndDate: TcxMaskEdit
          Left = 279
          Top = 55
          Properties.MaskKind = emkRegExprEx
          Properties.EditMask = 
            '([123][0-9])? [0-9][0-9] / (0?[1-9] | 1[012]) / ([012]?[1-9] | [' +
            '123]0 |31)'
          Properties.MaxLength = 0
          TabOrder = 2
          Text = '2009/10/15'
          Width = 98
        end
        object txtB7IccNo: TcxTextEdit
          Left = 139
          Top = 87
          Anchors = [akLeft, akTop, akRight]
          Properties.MaxLength = 0
          TabOrder = 3
          Text = '0'
          Width = 287
        end
      end
      object B9: TcxTabSheet
        Caption = 'B9.'#37325#38283#24050#38283#38971#36947
        ImageIndex = 15
        DesignSize = (
          436
          147)
        object Label38: TLabel
          Left = 21
          Top = 18
          Width = 52
          Height = 14
          Caption = 'Zip Code:'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label39: TLabel
          Left = 166
          Top = 18
          Width = 55
          Height = 14
          Caption = 'BAT '#21312#30908':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label40: TLabel
          Left = 21
          Top = 46
          Width = 91
          Height = 14
          Caption = 'NAGRA'#38971#36947#20195#30908':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label41: TLabel
          Left = 22
          Top = 67
          Width = 228
          Height = 14
          Caption = '( '#22810#32068#38971#36947#35531#29992#36887#40670#38548#38283', '#31684#20363': 10,20,30 )'
          Font.Charset = ANSI_CHARSET
          Font.Color = clBlack
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label42: TLabel
          Left = 21
          Top = 91
          Width = 64
          Height = 14
          Caption = #25480#27402#36215#22987#26085':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label43: TLabel
          Left = 209
          Top = 91
          Width = 64
          Height = 14
          Caption = #25480#27402#25130#27490#26085':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object txtB9ZipCode: TcxTextEdit
          Left = 79
          Top = 15
          Properties.MaxLength = 5
          TabOrder = 0
          Text = '0'
          Width = 65
        end
        object txtB9BatId: TcxTextEdit
          Left = 227
          Top = 15
          Properties.MaxLength = 5
          TabOrder = 1
          Text = '0'
          Width = 65
        end
        object txtB9ProductId: TcxTextEdit
          Left = 119
          Top = 43
          Anchors = [akLeft, akTop, akRight]
          Properties.MaxLength = 0
          TabOrder = 2
          Text = '0'
          Width = 287
        end
        object txtB9StartDate: TcxMaskEdit
          Left = 91
          Top = 87
          Properties.MaskKind = emkRegExprEx
          Properties.EditMask = 
            '([123][0-9])? [0-9][0-9] / (0?[1-9] | 1[012]) / ([012]?[1-9] | [' +
            '123]0 |31)'
          Properties.MaxLength = 0
          TabOrder = 3
          Text = '2009/10/15'
          Width = 98
        end
        object txtB9EndDate: TcxMaskEdit
          Left = 279
          Top = 87
          Properties.MaskKind = emkRegExprEx
          Properties.EditMask = 
            '([123][0-9])? [0-9][0-9] / (0?[1-9] | 1[012]) / ([012]?[1-9] | [' +
            '123]0 |31)'
          Properties.MaxLength = 0
          TabOrder = 4
          Text = '2009/10/15'
          Width = 98
        end
      end
      object C4: TcxTabSheet
        Caption = 'C4.'#35373#23450#37109#36958#21312#34399
        ImageIndex = 16
        object Label44: TLabel
          Left = 21
          Top = 18
          Width = 52
          Height = 14
          Caption = 'Zip Code:'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object txtC4ZipCode: TcxTextEdit
          Left = 79
          Top = 15
          Properties.MaxLength = 5
          TabOrder = 0
          Text = '0'
          Width = 65
        end
      end
      object E1: TcxTabSheet
        Caption = 'E1.'#35373#23450#35242#23376#23494#30908
        ImageIndex = 17
        object Label45: TLabel
          Left = 21
          Top = 19
          Width = 76
          Height = 14
          Caption = #26032#30340#35242#23376#23494#30908':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object txtE1PinCode: TcxTextEdit
          Left = 103
          Top = 15
          Properties.MaxLength = 4
          TabOrder = 0
          Text = '0'
          Width = 70
        end
      end
      object E2: TcxTabSheet
        Caption = 'E2.'#20659#36865#35338#24687
        ImageIndex = 18
        DesignSize = (
          436
          147)
        object Label46: TLabel
          Left = 25
          Top = 19
          Width = 28
          Height = 14
          Caption = #35338#24687':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label47: TLabel
          Left = 25
          Top = 55
          Width = 40
          Height = 14
          Caption = #32202#24613#24230':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object txtE2Message: TcxTextEdit
          Left = 59
          Top = 16
          Anchors = [akLeft, akTop, akRight]
          Properties.MaxLength = 100
          TabOrder = 0
          Width = 352
        end
        object txtE2Priority: TcxComboBox
          Left = 71
          Top = 51
          Properties.DropDownListStyle = lsFixedList
          Properties.Items.Strings = (
            #27491#24120
            #32202#24613)
          TabOrder = 1
          Text = #27491#24120
          Width = 98
        end
      end
      object E3: TcxTabSheet
        Caption = 'E3.'#24375#21046#25563#21488
        ImageIndex = 19
        object Label48: TLabel
          Left = 25
          Top = 19
          Width = 111
          Height = 14
          Caption = 'Orignal_Network_id:'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label49: TLabel
          Left = 25
          Top = 46
          Width = 73
          Height = 14
          Caption = 'Transport_id:'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object Label50: TLabel
          Left = 25
          Top = 74
          Width = 59
          Height = 14
          Caption = 'Service_id:'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object txtE3NetworkId: TcxTextEdit
          Left = 142
          Top = 16
          Properties.MaxLength = 0
          TabOrder = 0
          Text = '0'
          Width = 91
        end
        object txtE3TransportId: TcxTextEdit
          Left = 104
          Top = 44
          Properties.MaxLength = 0
          TabOrder = 1
          Text = '0'
          Width = 91
        end
        object txtE3ServiceId: TcxTextEdit
          Left = 90
          Top = 72
          Properties.MaxLength = 0
          TabOrder = 2
          Text = '0'
          Width = 91
        end
      end
      object E9: TcxTabSheet
        Caption = 'E9.'#35373#23450'BAT'#21312#30908
        ImageIndex = 20
        object Label51: TLabel
          Left = 26
          Top = 18
          Width = 55
          Height = 14
          Caption = 'BAT '#21312#30908':'
          Font.Charset = ANSI_CHARSET
          Font.Color = clRed
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = []
          ParentFont = False
        end
        object txtE9BatId: TcxTextEdit
          Left = 87
          Top = 15
          Properties.MaxLength = 5
          TabOrder = 0
          Text = '0'
          Width = 65
        end
      end
    end
    object pnHor: TPanel
      Left = 1
      Top = 173
      Width = 440
      Height = 38
      Align = alTop
      BevelOuter = bvNone
      TabOrder = 1
      object btnWriteSendNagra: TcxButton
        Left = 294
        Top = 6
        Width = 130
        Height = 25
        Caption = #23531'SEND_NAGRA'
        TabOrder = 0
        OnClick = btnWriteSendNagraClick
        Glyph.Data = {
          36040000424D3604000000000000360000002800000010000000100000000100
          2000000000000004000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          00000000000006151D310D4154AC050A0D230000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          000000000000687D84C390BDEFFF2082ADFF070C102800000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          00000000000006151D3192C6F0FF8287E1FF1D789FFF09111633000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          00000000000000000000557480B576D1FAFF8287E1FF1F6F95FF0910142E0000
          0000000000000000000000000000000000000000000000000000000000000000
          000000000000000000000507080C749CAAE376D1FAFF8287E1FF22688CFF0C16
          1C3D000000000000000000000000000000000000000000000000000000000000
          000000000000558CA3F448788CF43C626FF5567C86F976D1FAFF8287E1FF294F
          6BFF0B1419330000000000000000000000000000000000000000000000000000
          00000000000059C8E7FF7FE4FFFF7FE4FFFF94CCF2FF90BDEFFF81A0E9FF8287
          E1FF30475EFF090F132900000000000000000000000000000000000000000000
          00000000000029454C5559C8E7FFD9F2FEFF6EC0F5FF8287E1FF458A9FFE0E14
          1539010101030000000000000000000000000000000000000000000000000000
          0000000000000000000059C8E7FFB0E4FCFF76D1FAFF7FB0EAFF8287E1FF4A7A
          88FF0A0D0E270000000000000000000000000000000000000000000000000000
          0000000000000000000021363B4259C8E7FFD9F2FEFF76D1FAFF7EAAE9FF8287
          E1FF48626AFF0A0D0E2700000000000000000000000000000000000000000000
          000000000000000000000000000059C8E7FFB0E4FCFF7FE4FFFF89C2EEFF7CA5
          E8FF8287E1FF4B5358FF0304040C000000000000000000000000000000000000
          00000000000000000000000000001D2E33398ECCDDFF89D0E7FF89CFE6FF82C5
          DDFF78B2C8FF6A97ACFF14344272000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000000000000000
          0000000000000000000000000000000000000000000000000000}
      end
    end
  end
  object cxLookAndFeelController1: TcxLookAndFeelController
    Kind = lfStandard
    Left = 352
    Top = 240
  end
  object OraSession: TOraSession
    Left = 224
    Top = 244
  end
  object OraWriter: TOraSQL
    Session = OraSession
    AutoCommit = True
    Left = 268
    Top = 244
  end
  object OraReader: TOraQuery
    Session = OraSession
    FetchAll = True
    Left = 308
    Top = 244
  end
end
