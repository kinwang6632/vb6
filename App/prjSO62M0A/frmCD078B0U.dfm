object frmCD078B0: TfrmCD078B0
  Left = 346
  Top = 172
  ActiveControl = edtDateSt
  BorderStyle = bsSingle
  Caption = #20778#24800#32068#21512
  ClientHeight = 526
  ClientWidth = 759
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Tahoma'
  Font.Style = []
  Icon.Data = {
    0000010002002020100000000000E80200002600000010101000000000002801
    00000E0300002800000020000000400000000100040000000000800200000000
    0000000000000000000000000000000000000000800000800000008080008000
    0000800080008080000080808000C0C0C0000000FF0000FF000000FFFF00FF00
    0000FF00FF00FFFF0000FFFFFF00000000000000000000000000000000000000
    00000000BBBBBBBB000000000000000000000BBBBBBBBBBBBBB0000000000000
    0000BBBBBBBBBBBBBBBB00000000000000BBBBBBBBBBBBBBBBBBBB0000000000
    0BBBBBBBBBBBBBBBBBBBBBB000000000BBBBBBBBB0000000BBBBBBBB00000000
    BBBBBBB00BBBBBBB00BBBBBB0000000BBBBBBB00BBBBBBBBB00BBBBBB00000BB
    BBBBB00BBBBBBBBBBB00BBBBBB0000BBBBBBB0BBBBBBBBBBBBB0BBBBBB0000BB
    BBBB0BBBBBBBBBBBBBBB0BBBBB000BBBBBBB0BBBBBBBBBBBBBBB0BBBBBB00BBB
    BBBB0BBBBBBBBBBBBBBB0BBBBBB00BBBBBBBBBBBBBBBBBBBBBBBBBBBBBB00BBB
    BBBBBBBBBBBBBBBBBBBBBBBBBBB00BBBBBBBBBBBBBBBBBBBBBBBBBBBBBB00BBB
    BBBBBBB00BBBBBB00BBBBBBBBBB00BBBBBBBBB0000BBBB0000BBBBBBBBB00BBB
    BBBBBB0000BBBB0000BBBBBBBBB000BBBBBBBB0000BBBB0000BBBBBBBB0000BB
    BBBBBB0000BBBB0000BBBBBBBB0000BBBBBBBB0000BBBB0000BBBBBBBB00000B
    BBBBBBB00BBBBBB00BBBBBBBB0000000BBBBBBBBBBBBBBBBBBBBBBBB00000000
    BBBBBBBBBBBBBBBBBBBBBBBB000000000BBBBBBBBBBBBBBBBBBBBBB000000000
    00BBBBBBBBBBBBBBBBBBBB00000000000000BBBBBBBBBBBBBBBB000000000000
    00000BBBBBBBBBBBBBB000000000000000000000BBBBBBBB0000000000000000
    0000000000000000000000000000FFF00FFFFF8001FFFE00007FFC00003FF800
    001FF000000FE0000007C0000003C00000038000000180000001800000010000
    0000000000000000000000000000000000000000000000000000000000008000
    00018000000180000001C0000003C0000003E0000007F000000FF800001FFC00
    003FFE00007FFF8001FFFFF00FFF280000001000000020000000010004000000
    0000C00000000000000000000000000000000000000000000000000080000080
    00000080800080000000800080008080000080808000C0C0C0000000FF0000FF
    000000FFFF00FF000000FF00FF00FFFF0000FFFFFF0000000000000000000000
    0BBBBBB00000000BBBBBBBBBB00000BBB000003BBB0000BB00BBBB03BB000BBB
    0BBBBBB0BBB00BBB0BBBBBB0BBB00BBBBBBBBBBBBBB00BBBB00BB00BBBB00BBB
    B00BB00BBBB00BBBB00BB00BBBB000BBB00BB00BBB0000BBBBBBBBBBBB00000B
    BBBBBBBBB00000000BBBBBB000000000000000000000F81F0000E0070000C003
    0000840100008801000000000000000000000000000000000000000000000000
    00008001000080010000C0030000E0070000F81F0000}
  OldCreateOrder = False
  Position = poScreenCenter
  OnCloseQuery = FormCloseQuery
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 14
  object Panel1: TPanel
    Left = 0
    Top = 490
    Width = 759
    Height = 36
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 1
    object btnSearch: TButton
      Left = 7
      Top = 5
      Width = 60
      Height = 26
      Action = actSearch
      TabOrder = 0
    end
    object btnClear: TButton
      Left = 68
      Top = 5
      Width = 60
      Height = 26
      Action = actClear
      TabOrder = 1
    end
    object btnPrint: TButton
      Left = 250
      Top = 5
      Width = 60
      Height = 26
      Action = actPrint
      TabOrder = 2
      Visible = False
    end
    object btnInsert: TButton
      Left = 129
      Top = 5
      Width = 60
      Height = 26
      Action = actInsert
      TabOrder = 3
    end
    object btnModify: TButton
      Left = 190
      Top = 5
      Width = 60
      Height = 26
      Action = actUpdate
      TabOrder = 4
    end
    object btnCopyLocal: TButton
      Left = 531
      Top = 6
      Width = 60
      Height = 26
      Hint = #26412#21312#35079#35069
      Caption = #26412#21312#35079#35069
      TabOrder = 5
      OnClick = btnCopyLocalClick
    end
    object btnCopyOther: TButton
      Left = 596
      Top = 5
      Width = 60
      Height = 26
      Hint = #36328#21312#35079#35069
      Caption = #36328#21312#35079#35069
      TabOrder = 6
      OnClick = btnCopyOtherClick
    end
    object btnQueryOther: TButton
      Left = 659
      Top = 5
      Width = 88
      Height = 26
      Hint = #20419#37559#27963#21205#26597#35426
      Caption = #20419#37559#27963#21205#26597#35426
      TabOrder = 7
      OnClick = btnQueryOtherClick
    end
    object Button7: TButton
      Left = 398
      Top = 7
      Width = 60
      Height = 26
      Cancel = True
      Caption = #32080#26463'(&X)'
      TabOrder = 8
      OnClick = Button7Click
    end
    object btnFTP: TButton
      Left = 467
      Top = 6
      Width = 60
      Height = 26
      Hint = #19978#20659'FTP'
      Caption = #19978#20659'FTP'
      TabOrder = 9
      OnClick = btnFTPClick
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 759
    Height = 490
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 0
    object Panel3: TPanel
      Left = 0
      Top = 205
      Width = 759
      Height = 285
      Align = alClient
      BevelOuter = bvNone
      BorderWidth = 8
      TabOrder = 1
      object Bevel1: TBevel
        Left = 8
        Top = 254
        Width = 743
        Height = 2
        Align = alBottom
        Shape = bsSpacer
      end
      object DBGrid1: TDBGrid
        Left = 8
        Top = 8
        Width = 743
        Height = 246
        Align = alClient
        DataSource = DataSource1
        Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgConfirmDelete, dgCancelOnExit]
        ReadOnly = True
        TabOrder = 0
        TitleFont.Charset = ANSI_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -12
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
        OnDblClick = DBGrid1DblClick
        Columns = <
          item
            Expanded = False
            FieldName = 'CODENO'
            Title.Caption = #20778#24800#32068#21512#20195#30908
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'DESCRIPTION'
            Title.Caption = #20778#24800#32068#21512#21517#31281
            Width = 372
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'ONSALESTARTCDATE'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'ONSALESTOPCDATE'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'ONSALESTARTDATE'
            Visible = False
          end
          item
            Expanded = False
            FieldName = 'ONSALESTOPDATE'
            Visible = False
          end
          item
            Expanded = False
            FieldName = 'MASTERPROD'
            Width = 61
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'BUNDLE'
            Width = 59
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'BUNDLESAMEMON'
            Width = 81
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'BUNDLEMON'
            Width = 59
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'SAMEPERIOD'
            Width = 82
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'PERIOD'
            Width = 60
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'MERGEPRINT'
            Width = 81
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'PENAL'
            Width = 85
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'PENALTYPE'
            Width = 65
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'MERGEWORK'
            Width = 59
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'BUNDLEPRODNOTE'
            Title.Caption = #20778#24800#32068#21512#35498#26126
            Width = 300
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'WORKNOTE'
            Width = 300
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'REFNO'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'UPDTIME'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'UPDEN'
            Width = 101
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'STOPFLAG'
            Width = 63
            Visible = True
          end>
      end
      object Panel4: TPanel
        Left = 8
        Top = 256
        Width = 743
        Height = 21
        Align = alBottom
        BevelOuter = bvLowered
        Color = clWhite
        TabOrder = 1
      end
    end
    object Panel5: TPanel
      Left = 0
      Top = 0
      Width = 759
      Height = 205
      Align = alTop
      BevelOuter = bvNone
      TabOrder = 0
      object SpeedButton1: TSpeedButton
        Left = 133
        Top = 7
        Width = 24
        Height = 24
        Glyph.Data = {
          72020000424D720200000000000036000000280000000E0000000D0000000100
          1800000000003C020000C40E0000C40E00000000000000000000C8D0D4C8D0D4
          C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D40000
          80FFFFFF0000C8D0D4C8D0D4000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0
          D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D40000C8D0D4000080000080000080FFFF
          FFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4000080FFFFFFC8D0D40000C8D0
          D4000080000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D4000080FF
          FFFFC8D0D4C8D0D40000C8D0D4C8D0D4000080000080000080FFFFFFC8D0D4C8
          D0D4C8D0D4000080000080FFFFFFC8D0D4C8D0D40000C8D0D4C8D0D4C8D0D400
          0080000080000080FFFFFFC8D0D4000080000080FFFFFFC8D0D4C8D0D4C8D0D4
          0000C8D0D4C8D0D4C8D0D4C8D0D4000080000080000080000080000080FFFFFF
          C8D0D4C8D0D4C8D0D4C8D0D40000C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4000080
          000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D40000C8D0D4C8D0D4
          C8D0D4C8D0D4000080000080000080000080000080FFFFFFC8D0D4C8D0D4C8D0
          D4C8D0D40000C8D0D4C8D0D4C8D0D4000080000080000080FFFFFFC8D0D40000
          80FFFFFFC8D0D4C8D0D4C8D0D4C8D0D40000C8D0D40000800000800000800000
          80FFFFFFC8D0D4C8D0D4C8D0D4000080000080FFFFFFC8D0D4C8D0D400000000
          80000080000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D400008000
          0080FFFFFFC8D0D40000000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8
          D0D4C8D0D4C8D0D4C8D0D4000080000080FFFFFF0000}
        OnClick = SpeedButton1Click
      end
      object SpeedButton2: TSpeedButton
        Left = 133
        Top = 62
        Width = 24
        Height = 24
        Glyph.Data = {
          72020000424D720200000000000036000000280000000E0000000D0000000100
          1800000000003C020000C40E0000C40E00000000000000000000C8D0D4C8D0D4
          C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D40000
          80FFFFFF0000C8D0D4C8D0D4000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0
          D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D40000C8D0D4000080000080000080FFFF
          FFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4000080FFFFFFC8D0D40000C8D0
          D4000080000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D4000080FF
          FFFFC8D0D4C8D0D40000C8D0D4C8D0D4000080000080000080FFFFFFC8D0D4C8
          D0D4C8D0D4000080000080FFFFFFC8D0D4C8D0D40000C8D0D4C8D0D4C8D0D400
          0080000080000080FFFFFFC8D0D4000080000080FFFFFFC8D0D4C8D0D4C8D0D4
          0000C8D0D4C8D0D4C8D0D4C8D0D4000080000080000080000080000080FFFFFF
          C8D0D4C8D0D4C8D0D4C8D0D40000C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4000080
          000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D40000C8D0D4C8D0D4
          C8D0D4C8D0D4000080000080000080000080000080FFFFFFC8D0D4C8D0D4C8D0
          D4C8D0D40000C8D0D4C8D0D4C8D0D4000080000080000080FFFFFFC8D0D40000
          80FFFFFFC8D0D4C8D0D4C8D0D4C8D0D40000C8D0D40000800000800000800000
          80FFFFFFC8D0D4C8D0D4C8D0D4000080000080FFFFFFC8D0D4C8D0D400000000
          80000080000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D400008000
          0080FFFFFFC8D0D40000000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8
          D0D4C8D0D4C8D0D4C8D0D4000080000080FFFFFF0000}
        OnClick = SpeedButton2Click
      end
      object SpeedButton3: TSpeedButton
        Left = 133
        Top = 91
        Width = 24
        Height = 24
        Glyph.Data = {
          72020000424D720200000000000036000000280000000E0000000D0000000100
          1800000000003C020000C40E0000C40E00000000000000000000C8D0D4C8D0D4
          C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D40000
          80FFFFFF0000C8D0D4C8D0D4000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0
          D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D40000C8D0D4000080000080000080FFFF
          FFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4000080FFFFFFC8D0D40000C8D0
          D4000080000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D4000080FF
          FFFFC8D0D4C8D0D40000C8D0D4C8D0D4000080000080000080FFFFFFC8D0D4C8
          D0D4C8D0D4000080000080FFFFFFC8D0D4C8D0D40000C8D0D4C8D0D4C8D0D400
          0080000080000080FFFFFFC8D0D4000080000080FFFFFFC8D0D4C8D0D4C8D0D4
          0000C8D0D4C8D0D4C8D0D4C8D0D4000080000080000080000080000080FFFFFF
          C8D0D4C8D0D4C8D0D4C8D0D40000C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4000080
          000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D40000C8D0D4C8D0D4
          C8D0D4C8D0D4000080000080000080000080000080FFFFFFC8D0D4C8D0D4C8D0
          D4C8D0D40000C8D0D4C8D0D4C8D0D4000080000080000080FFFFFFC8D0D40000
          80FFFFFFC8D0D4C8D0D4C8D0D4C8D0D40000C8D0D40000800000800000800000
          80FFFFFFC8D0D4C8D0D4C8D0D4000080000080FFFFFFC8D0D4C8D0D400000000
          80000080000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D400008000
          0080FFFFFFC8D0D40000000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8
          D0D4C8D0D4C8D0D4C8D0D4000080000080FFFFFF0000}
        OnClick = SpeedButton3Click
      end
      object SpeedButton4: TSpeedButton
        Left = 133
        Top = 122
        Width = 24
        Height = 24
        Glyph.Data = {
          72020000424D720200000000000036000000280000000E0000000D0000000100
          1800000000003C020000C40E0000C40E00000000000000000000C8D0D4C8D0D4
          C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D40000
          80FFFFFF0000C8D0D4C8D0D4000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0
          D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D40000C8D0D4000080000080000080FFFF
          FFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4000080FFFFFFC8D0D40000C8D0
          D4000080000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D4000080FF
          FFFFC8D0D4C8D0D40000C8D0D4C8D0D4000080000080000080FFFFFFC8D0D4C8
          D0D4C8D0D4000080000080FFFFFFC8D0D4C8D0D40000C8D0D4C8D0D4C8D0D400
          0080000080000080FFFFFFC8D0D4000080000080FFFFFFC8D0D4C8D0D4C8D0D4
          0000C8D0D4C8D0D4C8D0D4C8D0D4000080000080000080000080000080FFFFFF
          C8D0D4C8D0D4C8D0D4C8D0D40000C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4000080
          000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D40000C8D0D4C8D0D4
          C8D0D4C8D0D4000080000080000080000080000080FFFFFFC8D0D4C8D0D4C8D0
          D4C8D0D40000C8D0D4C8D0D4C8D0D4000080000080000080FFFFFFC8D0D40000
          80FFFFFFC8D0D4C8D0D4C8D0D4C8D0D40000C8D0D40000800000800000800000
          80FFFFFFC8D0D4C8D0D4C8D0D4000080000080FFFFFFC8D0D4C8D0D400000000
          80000080000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D400008000
          0080FFFFFFC8D0D40000000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8
          D0D4C8D0D4C8D0D4C8D0D4000080000080FFFFFF0000}
        OnClick = SpeedButton4Click
      end
      object SpeedButton5: TSpeedButton
        Left = 133
        Top = 150
        Width = 24
        Height = 24
        Glyph.Data = {
          72020000424D720200000000000036000000280000000E0000000D0000000100
          1800000000003C020000C40E0000C40E00000000000000000000C8D0D4C8D0D4
          C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D40000
          80FFFFFF0000C8D0D4C8D0D4000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0
          D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D40000C8D0D4000080000080000080FFFF
          FFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4000080FFFFFFC8D0D40000C8D0
          D4000080000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D4000080FF
          FFFFC8D0D4C8D0D40000C8D0D4C8D0D4000080000080000080FFFFFFC8D0D4C8
          D0D4C8D0D4000080000080FFFFFFC8D0D4C8D0D40000C8D0D4C8D0D4C8D0D400
          0080000080000080FFFFFFC8D0D4000080000080FFFFFFC8D0D4C8D0D4C8D0D4
          0000C8D0D4C8D0D4C8D0D4C8D0D4000080000080000080000080000080FFFFFF
          C8D0D4C8D0D4C8D0D4C8D0D40000C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4000080
          000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D40000C8D0D4C8D0D4
          C8D0D4C8D0D4000080000080000080000080000080FFFFFFC8D0D4C8D0D4C8D0
          D4C8D0D40000C8D0D4C8D0D4C8D0D4000080000080000080FFFFFFC8D0D40000
          80FFFFFFC8D0D4C8D0D4C8D0D4C8D0D40000C8D0D40000800000800000800000
          80FFFFFFC8D0D4C8D0D4C8D0D4000080000080FFFFFFC8D0D4C8D0D400000000
          80000080000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D400008000
          0080FFFFFFC8D0D40000000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8
          D0D4C8D0D4C8D0D4C8D0D4000080000080FFFFFF0000}
        OnClick = SpeedButton5Click
      end
      object SpeedButton6: TSpeedButton
        Left = 133
        Top = 180
        Width = 24
        Height = 24
        Glyph.Data = {
          72020000424D720200000000000036000000280000000E0000000D0000000100
          1800000000003C020000C40E0000C40E00000000000000000000C8D0D4C8D0D4
          C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D40000
          80FFFFFF0000C8D0D4C8D0D4000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0
          D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D40000C8D0D4000080000080000080FFFF
          FFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4000080FFFFFFC8D0D40000C8D0
          D4000080000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D4000080FF
          FFFFC8D0D4C8D0D40000C8D0D4C8D0D4000080000080000080FFFFFFC8D0D4C8
          D0D4C8D0D4000080000080FFFFFFC8D0D4C8D0D40000C8D0D4C8D0D4C8D0D400
          0080000080000080FFFFFFC8D0D4000080000080FFFFFFC8D0D4C8D0D4C8D0D4
          0000C8D0D4C8D0D4C8D0D4C8D0D4000080000080000080000080000080FFFFFF
          C8D0D4C8D0D4C8D0D4C8D0D40000C8D0D4C8D0D4C8D0D4C8D0D4C8D0D4000080
          000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D40000C8D0D4C8D0D4
          C8D0D4C8D0D4000080000080000080000080000080FFFFFFC8D0D4C8D0D4C8D0
          D4C8D0D40000C8D0D4C8D0D4C8D0D4000080000080000080FFFFFFC8D0D40000
          80FFFFFFC8D0D4C8D0D4C8D0D4C8D0D40000C8D0D40000800000800000800000
          80FFFFFFC8D0D4C8D0D4C8D0D4000080000080FFFFFFC8D0D4C8D0D400000000
          80000080000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8D0D400008000
          0080FFFFFFC8D0D40000000080000080FFFFFFC8D0D4C8D0D4C8D0D4C8D0D4C8
          D0D4C8D0D4C8D0D4C8D0D4000080000080FFFFFF0000}
        OnClick = SpeedButton6Click
      end
      object Label1: TLabel
        Left = 88
        Top = 38
        Width = 64
        Height = 14
        Caption = #29986#21697#19978#26550#26085':'
      end
      object Label2: TLabel
        Left = 244
        Top = 38
        Width = 12
        Height = 14
        Caption = #33267
      end
      object Button1: TButton
        Left = 8
        Top = 7
        Width = 125
        Height = 24
        Caption = #20778#24800#32068#21512
        TabOrder = 0
        OnClick = Button1Click
      end
      object Button2: TButton
        Left = 8
        Top = 62
        Width = 125
        Height = 24
        Caption = #32371#36027#26399#25976
        TabOrder = 4
        OnClick = Button2Click
      end
      object Button3: TButton
        Left = 8
        Top = 91
        Width = 125
        Height = 24
        Caption = #32129#32004#26381#21209#39006#21029
        TabOrder = 6
        OnClick = Button3Click
      end
      object Button4: TButton
        Left = 8
        Top = 121
        Width = 125
        Height = 24
        Caption = #32129#32004#26376#25976
        TabOrder = 8
        OnClick = Button4Click
      end
      object Button5: TButton
        Left = 8
        Top = 150
        Width = 125
        Height = 24
        Caption = #20778#24800#32068#21512'--'#25910#36027#38917#30446
        TabOrder = 10
        OnClick = Button5Click
      end
      object Button6: TButton
        Left = 8
        Top = 180
        Width = 125
        Height = 24
        Caption = #20778#24800#32068#21512'--'#35037#27231#39006#21029
        TabOrder = 12
        OnClick = Button6Click
      end
      object edtBpCode: TEdit
        Left = 156
        Top = 8
        Width = 355
        Height = 22
        Color = 14737632
        TabOrder = 1
      end
      object edtPeriod: TEdit
        Left = 156
        Top = 63
        Width = 355
        Height = 22
        Color = 14737632
        TabOrder = 5
      end
      object edtServiceType: TEdit
        Left = 156
        Top = 92
        Width = 355
        Height = 22
        Color = 14737632
        TabOrder = 7
      end
      object edtBundleMon: TEdit
        Left = 156
        Top = 122
        Width = 355
        Height = 22
        Color = 14737632
        TabOrder = 9
      end
      object edtCItemCode: TEdit
        Left = 156
        Top = 151
        Width = 355
        Height = 22
        Color = 14737632
        TabOrder = 11
      end
      object edtInstCode: TEdit
        Left = 156
        Top = 181
        Width = 355
        Height = 22
        Color = 14737632
        TabOrder = 13
      end
      object GroupBox1: TGroupBox
        Left = 527
        Top = 3
        Width = 220
        Height = 196
        TabOrder = 14
        object chkMasterProd: TCheckBox
          Left = 18
          Top = 24
          Width = 93
          Height = 17
          Caption = #20027#25512#21830#21697
          TabOrder = 0
        end
        object chkStopFlag: TCheckBox
          Left = 18
          Top = 48
          Width = 134
          Height = 17
          Caption = #20572#29992#26159#21542#39023#31034
          TabOrder = 1
        end
      end
      object edtDateSt: TMaskEdit
        Left = 156
        Top = 35
        Width = 83
        Height = 22
        EditMask = '!9999/99/99;0;_'
        MaxLength = 10
        TabOrder = 2
        OnExit = edtDateStExit
      end
      object edtDateEd: TMaskEdit
        Left = 260
        Top = 35
        Width = 83
        Height = 22
        EditMask = '!9999/99/99;0;_'
        MaxLength = 10
        TabOrder = 3
        OnExit = edtDateStExit
      end
    end
  end
  object CD078DataSet: TClientDataSet
    Aggregates = <>
    Params = <>
    ProviderName = 'CD078DataProvider'
    AfterOpen = CD078DataSetAfterOpen
    AfterScroll = CD078DataSetAfterScroll
    OnCalcFields = CD078DataSetCalcFields
    Left = 144
    Top = 362
    object CD078DataSetCODENO: TStringField
      DisplayLabel = #32068#21512#29986#21697#20195#30908
      DisplayWidth = 15
      FieldName = 'CODENO'
      Size = 10
    end
    object CD078DataSetDESCRIPTION: TStringField
      DisplayLabel = #32068#21512#29986#21697#21517#31281
      DisplayWidth = 63
      FieldName = 'DESCRIPTION'
      Size = 50
    end
    object CD078DataSetONSALESTARTDATE: TDateTimeField
      DisplayLabel = #29986#21697#19978#26550#26085'('#36215')'
      DisplayWidth = 19
      FieldName = 'ONSALESTARTDATE'
    end
    object CD078DataSetONSALESTOPDATE: TDateTimeField
      DisplayLabel = #29986#21697#19978#26550#26085'('#36804')'
      DisplayWidth = 18
      FieldName = 'ONSALESTOPDATE'
    end
    object CD078DataSetMASTERPROD: TIntegerField
      DisplayLabel = #20027#25512#21830#21697
      DisplayWidth = 11
      FieldName = 'MASTERPROD'
    end
    object CD078DataSetBUNDLE: TIntegerField
      DisplayLabel = #30342#38920#32129#32004
      DisplayWidth = 12
      FieldName = 'BUNDLE'
    end
    object CD078DataSetBUNDLESAMEMON: TIntegerField
      DisplayLabel = #32129#32004#26376#25976#19968#33268
      DisplayWidth = 15
      FieldName = 'BUNDLESAMEMON'
    end
    object CD078DataSetBUNDLEMON: TIntegerField
      DisplayLabel = #32129#32004#26376#25976
      DisplayWidth = 12
      FieldName = 'BUNDLEMON'
    end
    object CD078DataSetSAMEPERIOD: TIntegerField
      DisplayLabel = #32371#36027#26399#25976#19968#33268
      DisplayWidth = 15
      FieldName = 'SAMEPERIOD'
    end
    object CD078DataSetPERIOD: TIntegerField
      DisplayLabel = #32371#36027#26399#25976
      DisplayWidth = 14
      FieldName = 'PERIOD'
    end
    object CD078DataSetMERGEPRINT: TIntegerField
      DisplayLabel = #24115#27454#21512#20341#21015#21360
      DisplayWidth = 15
      FieldName = 'MERGEPRINT'
    end
    object CD078DataSetPENAL: TIntegerField
      DisplayLabel = #36949#32004#35336#31639#19968#33268
      DisplayWidth = 15
      FieldName = 'PENAL'
    end
    object CD078DataSetPENALTYPE: TIntegerField
      DisplayLabel = #36949#32004#35336#31639
      DisplayWidth = 14
      FieldName = 'PENALTYPE'
    end
    object CD078DataSetMERGEWORK: TIntegerField
      DisplayLabel = #21512#20341#27966#24037
      DisplayWidth = 14
      FieldName = 'MERGEWORK'
    end
    object CD078DataSetBUNDLEPRODNOTE: TStringField
      DisplayLabel = #32068#21512#29986#21697#35498#26126
      DisplayWidth = 50
      FieldName = 'BUNDLEPRODNOTE'
      Size = 2000
    end
    object CD078DataSetWORKNOTE: TStringField
      DisplayLabel = #24037#21934#21015#21360#35498#26126
      DisplayWidth = 2800
      FieldName = 'WORKNOTE'
      Size = 2000
    end
    object CD078DataSetREFNO: TIntegerField
      DisplayLabel = #21443#32771#34399
      DisplayWidth = 14
      FieldName = 'REFNO'
    end
    object CD078DataSetUPDTIME: TStringField
      DisplayLabel = #30064#21205#26178#38291
      DisplayWidth = 27
      FieldName = 'UPDTIME'
      Size = 19
    end
    object CD078DataSetUPDEN: TStringField
      DisplayLabel = #30064#21205#20154#21729
      DisplayWidth = 28
      FieldName = 'UPDEN'
    end
    object CD078DataSetSTOPFLAG: TIntegerField
      DisplayLabel = #26159#21542#20572#29992
      DisplayWidth = 14
      FieldName = 'STOPFLAG'
    end
    object CD078DataSetONSALESTARTCDATE: TStringField
      DisplayLabel = #29986#21697#19978#26550#26085'('#36215')'
      FieldKind = fkCalculated
      FieldName = 'ONSALESTARTCDATE'
      Size = 10
      Calculated = True
    end
    object CD078DataSetONSALESTOPCDATE: TStringField
      DisplayLabel = #29986#21697#19978#26550#26085'('#36804')'
      FieldKind = fkCalculated
      FieldName = 'ONSALESTOPCDATE'
      Size = 10
      Calculated = True
    end
    object CD078DataSetCanUseType: TIntegerField
      FieldName = 'CanUseType'
    end
  end
  object DataSource1: TDataSource
    DataSet = CD078DataSet
    Left = 108
    Top = 362
  end
  object ActionList1: TActionList
    Left = 72
    Top = 361
    object actSearch: TAction
      Caption = 'F3.'#23563#25214
      ShortCut = 114
      OnExecute = actSearchExecute
    end
    object actClear: TAction
      Caption = 'F4.'#28165#38500
      ShortCut = 115
      OnExecute = actClearExecute
    end
    object actPrint: TAction
      Caption = 'F5.'#21015#21360
      ShortCut = 116
      OnExecute = actPrintExecute
    end
    object actInsert: TAction
      Caption = 'F6.'#26032#22686
      ShortCut = 117
      OnExecute = actInsertExecute
    end
    object actUpdate: TAction
      Caption = 'F11.'#20462#25913
      ShortCut = 122
      OnExecute = actUpdateExecute
    end
  end
  object CD078DataProvider: TDataSetProvider
    DataSet = DBController.DataReader
    Options = [poAllowCommandText]
    Left = 180
    Top = 361
  end
end
