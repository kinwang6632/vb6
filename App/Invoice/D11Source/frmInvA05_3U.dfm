object frmInvA05_3: TfrmInvA05_3
  Left = 541
  Top = 369
  BorderStyle = bsToolWindow
  Caption = #21015#21360#36984#38917
  ClientHeight = 192
  ClientWidth = 210
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 210
    Height = 192
    Align = alClient
    TabOrder = 0
    object btnPrintNew: TBitBtn
      Left = 16
      Top = 147
      Width = 81
      Height = 26
      Caption = #30906#23450
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ModalResult = 1
      ParentFont = False
      TabOrder = 1
      NumGlyphs = 2
    end
    object btnExit: TBitBtn
      Left = 120
      Top = 147
      Width = 81
      Height = 26
      Caption = #21462#28040
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ModalResult = 2
      ParentFont = False
      TabOrder = 2
      NumGlyphs = 2
    end
    object GroupBox1: TGroupBox
      Left = 22
      Top = 47
      Width = 171
      Height = 82
      Caption = '  '#26684#24335#20839#23481'  '
      Font.Charset = ANSI_CHARSET
      Font.Color = clTeal
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 0
      object ckbInvTitle: TCheckBox
        Left = 24
        Top = 16
        Width = 97
        Height = 17
        Caption = '  '#30332#31080#25260#38957
        Checked = True
        State = cbChecked
        TabOrder = 0
      end
      object ckbMailAddr: TCheckBox
        Left = 24
        Top = 37
        Width = 97
        Height = 17
        Caption = '  '#30332#31080#22320#22336
        Checked = True
        State = cbChecked
        TabOrder = 1
      end
      object ckbMemo1: TCheckBox
        Left = 24
        Top = 57
        Width = 97
        Height = 17
        Caption = '  '#30332#31080#20633#35387
        Checked = True
        State = cbChecked
        TabOrder = 2
      end
    end
    object Panel2: TPanel
      Left = 1
      Top = 1
      Width = 208
      Height = 32
      Align = alTop
      Color = clSkyBlue
      TabOrder = 3
      object Label1: TLabel
        Left = 16
        Top = 7
        Width = 68
        Height = 16
        Caption = #21015#21360#36984#38917
        Font.Charset = ANSI_CHARSET
        Font.Color = clBlue
        Font.Height = -16
        Font.Name = #26032#32048#26126#39636
        Font.Style = [fsBold]
        ParentFont = False
      end
    end
  end
end
