object frmSo18C4: TfrmSo18C4
  Left = 189
  Top = 105
  Width = 696
  Height = 325
  Caption = 'frmSo18C4'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 16
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 688
    Height = 298
    Align = alClient
    BevelInner = bvRaised
    BevelOuter = bvLowered
    TabOrder = 0
    object Label1: TLabel
      Left = 80
      Top = 32
      Width = 48
      Height = 16
      Caption = #25910#36027#21934#34399
    end
    object btnQuery: TBitBtn
      Left = 305
      Top = 27
      Width = 75
      Height = 25
      Caption = #26597#35426
      TabOrder = 1
      OnClick = btnQueryClick
    end
    object btnExit: TBitBtn
      Left = 388
      Top = 27
      Width = 75
      Height = 25
      Caption = #32080#26463
      TabOrder = 4
      OnClick = btnExitClick
    end
    object GroupBox1: TGroupBox
      Left = 32
      Top = 152
      Width = 625
      Height = 129
      Caption = ' '#25910#36027#21934#20839#23481' '
      TabOrder = 5
      object dbgSo034: TDBGrid
        Left = 2
        Top = 18
        Width = 621
        Height = 109
        Align = alClient
        Color = clInfoBk
        DataSource = dsrSo034
        ReadOnly = True
        TabOrder = 0
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -13
        TitleFont.Name = 'MS Sans Serif'
        TitleFont.Style = []
      end
    end
    object edtBillNo: TEdit
      Left = 152
      Top = 28
      Width = 145
      Height = 24
      MaxLength = 15
      TabOrder = 0
    end
    object edtSrcHandPaperNo: TEdit
      Left = 152
      Top = 69
      Width = 145
      Height = 24
      MaxLength = 15
      ReadOnly = True
      TabOrder = 8
    end
    object edtNewHandPaperNo: TEdit
      Left = 152
      Top = 109
      Width = 145
      Height = 24
      MaxLength = 15
      TabOrder = 2
    end
    object btnUpdate: TButton
      Left = 305
      Top = 108
      Width = 75
      Height = 25
      Caption = #26356#26032
      TabOrder = 3
      OnClick = btnUpdateClick
    end
    object StaticText1: TStaticText
      Left = 32
      Top = 69
      Width = 116
      Height = 24
      Caption = #23565#25033#20043#25163#38283#21934#34399
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 6
    end
    object StaticText2: TStaticText
      Left = 63
      Top = 108
      Width = 84
      Height = 24
      Caption = #26032#25163#38283#21934#34399
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 7
    end
    object btnReport: TBitBtn
      Left = 388
      Top = 108
      Width = 75
      Height = 25
      Caption = #30064#21205#22577#34920
      TabOrder = 9
      OnClick = btnReportClick
    end
  end
  object dsrSo034: TDataSource
    DataSet = dtmMain.qrySo034
    OnDataChange = dsrSo034DataChange
    Left = 480
    Top = 112
  end
end
