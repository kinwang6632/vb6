object frmSO8F10_4: TfrmSO8F10_4
  Left = 276
  Top = 200
  Width = 307
  Height = 186
  Caption = 'frmSO8F10_4'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 299
    Height = 158
    Align = alClient
    TabOrder = 0
    object lblUserID: TLabel
      Left = 9
      Top = 28
      Width = 80
      Height = 19
      Caption = #20351#29992#32773#20195#30908
      Color = clBtnFace
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlack
      Font.Height = -16
      Font.Name = 'Times New Roman'
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
    object lblPassword: TLabel
      Left = 41
      Top = 71
      Width = 48
      Height = 19
      Caption = #23494#12288#30908
      Color = clBtnFace
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlack
      Font.Height = -16
      Font.Name = 'Times New Roman'
      Font.Style = []
      ParentColor = False
      ParentFont = False
    end
    object edtUserID: TEdit
      Left = 106
      Top = 27
      Width = 166
      Height = 21
      TabOrder = 0
    end
    object edtPassword: TEdit
      Left = 106
      Top = 75
      Width = 166
      Height = 21
      PasswordChar = '*'
      TabOrder = 1
    end
    object btnCancel: TBitBtn
      Left = 175
      Top = 112
      Width = 75
      Height = 31
      Caption = #21462#28040
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlack
      Font.Height = -16
      Font.Name = 'Times New Roman'
      Font.Style = []
      ModalResult = 2
      ParentFont = False
      TabOrder = 3
      OnClick = btnCancelClick
    end
    object btnOK: TBitBtn
      Left = 48
      Top = 112
      Width = 75
      Height = 31
      Caption = #30906#23450
      Default = True
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlack
      Font.Height = -16
      Font.Name = 'Times New Roman'
      Font.Style = []
      ParentFont = False
      TabOrder = 2
      OnClick = btnOKClick
    end
  end
end
