object frmGift: TfrmGift
  Left = 90
  Top = 172
  BorderStyle = bsDialog
  Caption = #36104#21697#26041#26696' [SO64N0A1]'
  ClientHeight = 312
  ClientWidth = 704
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 271
    Width = 704
    Height = 41
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 0
    object Bevel1: TBevel
      Left = 0
      Top = 0
      Width = 704
      Height = 50
      Align = alTop
      Shape = bsTopLine
    end
    object btnSave: TButton
      Left = 24
      Top = 6
      Width = 75
      Height = 28
      Action = actSave
      TabOrder = 0
    end
    object btnCancel: TButton
      Left = 602
      Top = 6
      Width = 75
      Height = 28
      Action = actCancel
      TabOrder = 1
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 704
    Height = 271
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 1
    object Label1: TLabel
      Left = 11
      Top = 48
      Width = 12
      Height = 13
      Caption = #19968
    end
    object Label2: TLabel
      Left = 11
      Top = 84
      Width = 12
      Height = 13
      Caption = #20108
    end
    object Label3: TLabel
      Left = 11
      Top = 122
      Width = 12
      Height = 13
      Caption = #19977
    end
    object Label4: TLabel
      Left = 11
      Top = 158
      Width = 12
      Height = 13
      Caption = #22235
    end
    object Label5: TLabel
      Left = 11
      Top = 195
      Width = 12
      Height = 13
      Caption = #20116
    end
    object Label6: TLabel
      Left = 11
      Top = 233
      Width = 12
      Height = 13
      Caption = #20845
    end
    object Label7: TLabel
      Left = 79
      Top = 22
      Width = 48
      Height = 13
      Caption = #35330#36092#37329#38989
    end
    object Label9: TLabel
      Left = 205
      Top = 22
      Width = 48
      Height = 13
      Caption = #36104#36865#26041#24335
    end
    object Label10: TLabel
      Left = 304
      Top = 22
      Width = 48
      Height = 13
      Caption = #36104#36865#38917#30446
    end
    object Label11: TLabel
      Left = 512
      Top = 22
      Width = 48
      Height = 13
      Caption = #36104#21697#20839#23481
    end
    object Panel3: TPanel
      Left = 29
      Top = 41
      Width = 661
      Height = 27
      BevelOuter = bvNone
      TabOrder = 0
      object Label8: TLabel
        Left = 68
        Top = 6
        Width = 8
        Height = 13
        Caption = '~'
      end
      object EDT_GiftOrderPrice1: TEdit
        Left = 2
        Top = 3
        Width = 62
        Height = 21
        MaxLength = 8
        TabOrder = 0
        OnChange = EDT_GiftOrderPrice1Change
        OnExit = EDT_GiftOrderPrice1Exit
        OnKeyPress = EDT_GiftOrderPrice1KeyPress
      end
      object EDT_GiftOrderPriceMax1: TEdit
        Left = 82
        Top = 3
        Width = 62
        Height = 21
        MaxLength = 8
        TabOrder = 1
        OnChange = EDT_GiftOrderPrice1Change
        OnExit = EDT_GiftOrderPriceMax1Exit
        OnKeyPress = EDT_GiftOrderPrice1KeyPress
      end
      object EDT_GiftType1: TComboBox
        Left = 150
        Top = 3
        Width = 100
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 2
      end
      object EDT_GiftKind1: TComboBox
        Left = 259
        Top = 3
        Width = 100
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 3
        OnChange = EDT_GiftKind1Change
      end
      object Button1: TButton
        Left = 365
        Top = 3
        Width = 22
        Height = 21
        Caption = '..'
        TabOrder = 4
        OnClick = Button1Click
      end
      object X_ArticleNoStr1: TBitBtn
        Left = 387
        Top = 3
        Width = 22
        Height = 21
        TabOrder = 5
        OnClick = X_ArticleNoStr1Click
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
      end
      object EDT_ArticleNoStr1: TEdit
        Left = 408
        Top = 3
        Width = 240
        Height = 20
        Color = 14737632
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ParentFont = False
        ReadOnly = True
        TabOrder = 6
      end
      object EDT_ArticleNo1: TListBox
        Left = 487
        Top = 3
        Width = 100
        Height = 20
        ItemHeight = 13
        TabOrder = 7
        Visible = False
      end
    end
    object Panel4: TPanel
      Left = 29
      Top = 77
      Width = 661
      Height = 27
      BevelOuter = bvNone
      TabOrder = 1
      object Label12: TLabel
        Left = 68
        Top = 6
        Width = 8
        Height = 13
        Caption = '~'
      end
      object EDT_GiftOrderPrice2: TEdit
        Left = 2
        Top = 3
        Width = 62
        Height = 21
        MaxLength = 8
        TabOrder = 0
        OnChange = EDT_GiftOrderPrice1Change
        OnExit = EDT_GiftOrderPrice1Exit
        OnKeyPress = EDT_GiftOrderPrice1KeyPress
      end
      object EDT_GiftOrderPriceMax2: TEdit
        Left = 82
        Top = 3
        Width = 62
        Height = 21
        MaxLength = 8
        TabOrder = 1
        OnChange = EDT_GiftOrderPrice1Change
        OnKeyPress = EDT_GiftOrderPrice1KeyPress
      end
      object EDT_GiftType2: TComboBox
        Left = 150
        Top = 3
        Width = 100
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 2
      end
      object EDT_GiftKind2: TComboBox
        Left = 259
        Top = 3
        Width = 100
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 3
        OnChange = EDT_GiftKind1Change
      end
      object Button2: TButton
        Left = 365
        Top = 3
        Width = 22
        Height = 21
        Caption = '..'
        TabOrder = 4
        OnClick = Button1Click
      end
      object X_ArticleNoStr2: TBitBtn
        Left = 387
        Top = 3
        Width = 22
        Height = 21
        TabOrder = 5
        OnClick = X_ArticleNoStr1Click
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
      end
      object EDT_ArticleNoStr2: TEdit
        Left = 408
        Top = 3
        Width = 240
        Height = 20
        Color = 14737632
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ParentFont = False
        ReadOnly = True
        TabOrder = 6
      end
      object EDT_ArticleNo2: TListBox
        Left = 487
        Top = 3
        Width = 100
        Height = 20
        ItemHeight = 13
        TabOrder = 7
        Visible = False
      end
    end
    object Panel5: TPanel
      Left = 29
      Top = 113
      Width = 661
      Height = 27
      BevelOuter = bvNone
      TabOrder = 2
      object Label13: TLabel
        Left = 68
        Top = 6
        Width = 8
        Height = 13
        Caption = '~'
      end
      object EDT_GiftOrderPrice3: TEdit
        Left = 2
        Top = 3
        Width = 62
        Height = 21
        MaxLength = 8
        TabOrder = 0
        OnChange = EDT_GiftOrderPrice1Change
        OnExit = EDT_GiftOrderPrice1Exit
        OnKeyPress = EDT_GiftOrderPrice1KeyPress
      end
      object EDT_GiftOrderPriceMax3: TEdit
        Left = 82
        Top = 3
        Width = 62
        Height = 21
        MaxLength = 8
        TabOrder = 1
        OnChange = EDT_GiftOrderPrice1Change
        OnKeyPress = EDT_GiftOrderPrice1KeyPress
      end
      object EDT_GiftType3: TComboBox
        Left = 150
        Top = 3
        Width = 100
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 2
      end
      object EDT_GiftKind3: TComboBox
        Left = 259
        Top = 3
        Width = 100
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 3
        OnChange = EDT_GiftKind1Change
      end
      object Button3: TButton
        Left = 365
        Top = 3
        Width = 22
        Height = 21
        Caption = '..'
        TabOrder = 4
        OnClick = Button1Click
      end
      object X_ArticleNoStr3: TBitBtn
        Left = 387
        Top = 3
        Width = 22
        Height = 21
        TabOrder = 5
        OnClick = X_ArticleNoStr1Click
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
      end
      object EDT_ArticleNoStr3: TEdit
        Left = 408
        Top = 3
        Width = 240
        Height = 20
        Color = 14737632
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ParentFont = False
        ReadOnly = True
        TabOrder = 6
      end
      object EDT_ArticleNo3: TListBox
        Left = 487
        Top = 3
        Width = 100
        Height = 20
        ItemHeight = 13
        TabOrder = 7
        Visible = False
      end
    end
    object Panel6: TPanel
      Left = 29
      Top = 149
      Width = 661
      Height = 27
      BevelOuter = bvNone
      TabOrder = 3
      object Label14: TLabel
        Left = 68
        Top = 6
        Width = 8
        Height = 13
        Caption = '~'
      end
      object EDT_GiftOrderPrice4: TEdit
        Left = 2
        Top = 3
        Width = 62
        Height = 21
        MaxLength = 8
        TabOrder = 0
        OnChange = EDT_GiftOrderPrice1Change
        OnExit = EDT_GiftOrderPrice1Exit
        OnKeyPress = EDT_GiftOrderPrice1KeyPress
      end
      object EDT_GiftOrderPriceMax4: TEdit
        Left = 82
        Top = 3
        Width = 62
        Height = 21
        MaxLength = 8
        TabOrder = 1
        OnChange = EDT_GiftOrderPrice1Change
        OnKeyPress = EDT_GiftOrderPrice1KeyPress
      end
      object EDT_GiftType4: TComboBox
        Left = 150
        Top = 3
        Width = 100
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 2
      end
      object EDT_GiftKind4: TComboBox
        Left = 259
        Top = 3
        Width = 100
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 3
        OnChange = EDT_GiftKind1Change
      end
      object Button4: TButton
        Left = 365
        Top = 3
        Width = 22
        Height = 21
        Caption = '..'
        TabOrder = 4
        OnClick = Button1Click
      end
      object X_ArticleNoStr4: TBitBtn
        Left = 387
        Top = 3
        Width = 22
        Height = 21
        TabOrder = 5
        OnClick = X_ArticleNoStr1Click
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
      end
      object EDT_ArticleNoStr4: TEdit
        Left = 408
        Top = 3
        Width = 240
        Height = 20
        Color = 14737632
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ParentFont = False
        ReadOnly = True
        TabOrder = 6
      end
      object EDT_ArticleNo4: TListBox
        Left = 487
        Top = 3
        Width = 100
        Height = 20
        ItemHeight = 13
        TabOrder = 7
        Visible = False
      end
    end
    object Panel7: TPanel
      Left = 29
      Top = 188
      Width = 661
      Height = 27
      BevelOuter = bvNone
      TabOrder = 4
      object Label15: TLabel
        Left = 68
        Top = 6
        Width = 8
        Height = 13
        Caption = '~'
      end
      object EDT_GiftOrderPrice5: TEdit
        Left = 2
        Top = 3
        Width = 62
        Height = 21
        MaxLength = 8
        TabOrder = 0
        OnChange = EDT_GiftOrderPrice1Change
        OnExit = EDT_GiftOrderPrice1Exit
        OnKeyPress = EDT_GiftOrderPrice1KeyPress
      end
      object EDT_GiftOrderPriceMax5: TEdit
        Left = 82
        Top = 3
        Width = 62
        Height = 21
        MaxLength = 8
        TabOrder = 1
        OnChange = EDT_GiftOrderPrice1Change
        OnKeyPress = EDT_GiftOrderPrice1KeyPress
      end
      object EDT_GiftType5: TComboBox
        Left = 150
        Top = 3
        Width = 100
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 2
      end
      object EDT_GiftKind5: TComboBox
        Left = 259
        Top = 3
        Width = 100
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 3
        OnChange = EDT_GiftKind1Change
      end
      object Button5: TButton
        Left = 365
        Top = 3
        Width = 22
        Height = 21
        Caption = '..'
        TabOrder = 4
        OnClick = Button1Click
      end
      object X_ArticleNoStr5: TBitBtn
        Left = 387
        Top = 3
        Width = 22
        Height = 21
        TabOrder = 5
        OnClick = X_ArticleNoStr1Click
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
      end
      object EDT_ArticleNoStr5: TEdit
        Left = 408
        Top = 3
        Width = 240
        Height = 20
        Color = 14737632
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ParentFont = False
        ReadOnly = True
        TabOrder = 6
      end
      object EDT_ArticleNo5: TListBox
        Left = 487
        Top = 3
        Width = 100
        Height = 20
        ItemHeight = 13
        TabOrder = 7
        Visible = False
      end
    end
    object Panel8: TPanel
      Left = 29
      Top = 226
      Width = 661
      Height = 27
      BevelOuter = bvNone
      TabOrder = 5
      object Label16: TLabel
        Left = 69
        Top = 6
        Width = 8
        Height = 13
        Caption = '~'
      end
      object EDT_GiftOrderPrice6: TEdit
        Left = 2
        Top = 3
        Width = 62
        Height = 21
        MaxLength = 8
        TabOrder = 0
        OnChange = EDT_GiftOrderPrice1Change
        OnExit = EDT_GiftOrderPrice1Exit
        OnKeyPress = EDT_GiftOrderPrice1KeyPress
      end
      object EDT_GiftOrderPriceMax6: TEdit
        Left = 83
        Top = 3
        Width = 62
        Height = 21
        MaxLength = 8
        TabOrder = 1
        OnChange = EDT_GiftOrderPrice1Change
        OnKeyPress = EDT_GiftOrderPrice1KeyPress
      end
      object EDT_GiftType6: TComboBox
        Left = 150
        Top = 3
        Width = 100
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 2
      end
      object EDT_GiftKind6: TComboBox
        Left = 260
        Top = 3
        Width = 100
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 3
        OnChange = EDT_GiftKind1Change
      end
      object Button6: TButton
        Left = 365
        Top = 3
        Width = 22
        Height = 21
        Caption = '..'
        TabOrder = 4
        OnClick = Button1Click
      end
      object X_ArticleNoStr6: TBitBtn
        Left = 387
        Top = 3
        Width = 22
        Height = 21
        TabOrder = 5
        OnClick = X_ArticleNoStr1Click
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
      end
      object EDT_ArticleNoStr6: TEdit
        Left = 408
        Top = 3
        Width = 240
        Height = 20
        Color = 14737632
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -12
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ParentFont = False
        ReadOnly = True
        TabOrder = 6
      end
      object EDT_ArticleNo6: TListBox
        Left = 487
        Top = 3
        Width = 100
        Height = 20
        ItemHeight = 13
        TabOrder = 7
        Visible = False
      end
    end
  end
  object ActionList1: TActionList
    Left = 124
    Top = 278
    object actSave: TAction
      Caption = 'F2.'#30906#23450
      ShortCut = 113
      OnExecute = actSaveExecute
    end
    object actCancel: TAction
      Caption = #21462#28040'(&X)'
      ShortCut = 32856
      OnExecute = actCancelExecute
    end
  end
end
