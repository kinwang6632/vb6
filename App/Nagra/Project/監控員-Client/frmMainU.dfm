object frmMain: TfrmMain
  Left = 302
  Top = 217
  Width = 369
  Height = 215
  BorderIcons = []
  Caption = 'CA Monitor Client'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 16
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 361
    Height = 188
    Align = alClient
    BevelInner = bvLowered
    TabOrder = 0
    object StaticText1: TStaticText
      Left = 12
      Top = 31
      Width = 167
      Height = 20
      Caption = 'CA Command Gateway IP : '
      TabOrder = 6
    end
    object edtIP: TEdit
      Left = 180
      Top = 28
      Width = 121
      Height = 24
      TabOrder = 0
    end
    object btnSo8D10: TBitBtn
      Left = 67
      Top = 72
      Width = 100
      Height = 40
      Caption = #20597#28204'CA'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clNavy
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 2
      OnClick = btnSo8D10Click
    end
    object btnSo8D20A: TBitBtn
      Left = 67
      Top = 128
      Width = 100
      Height = 40
      Caption = #20572#27490'CA'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clNavy
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 4
      OnClick = btnSo8D20AClick
    end
    object btnSo8D20B: TBitBtn
      Left = 195
      Top = 128
      Width = 100
      Height = 40
      Caption = #21855#21205'CA'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clNavy
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 5
      OnClick = btnSo8D20BClick
    end
    object btnExit: TBitBtn
      Left = 195
      Top = 72
      Width = 100
      Height = 40
      Caption = #32080#26463
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clNavy
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 3
      OnClick = btnExitClick
    end
    object btnConn: TBitBtn
      Left = 304
      Top = 27
      Width = 51
      Height = 25
      Caption = #36899#25509
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
      OnClick = btnConnClick
    end
  end
  object ClientSocket: TClientSocket
    Active = False
    ClientType = ctNonBlocking
    Port = 1024
    OnConnect = ClientSocketConnect
    OnRead = ClientSocketRead
    OnError = ClientSocketError
    Left = 58
    Top = 5
  end
end
