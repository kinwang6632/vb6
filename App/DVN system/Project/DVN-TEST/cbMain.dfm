object fmMain: TfmMain
  Left = 365
  Top = 76
  ActiveControl = txtDvnIp
  BorderStyle = bsDialog
  Caption = 'fmMain'
  ClientHeight = 639
  ClientWidth = 504
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
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 14
  object SocketImage: TImage
    Left = 28
    Top = 50
    Width = 32
    Height = 32
    AutoSize = True
  end
  object Label1: TLabel
    Left = 134
    Top = 57
    Width = 142
    Height = 14
    Caption = 'DVN SMS Gateway IP :'
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Verdana'
    Font.Style = []
    ParentFont = False
  end
  object Label2: TLabel
    Left = 121
    Top = 87
    Width = 155
    Height = 14
    Caption = 'DVN SMS Gateway Port :'
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Verdana'
    Font.Style = []
    ParentFont = False
  end
  object Bevel1: TBevel
    Left = 8
    Top = 119
    Width = 490
    Height = 3
    Shape = bsBottomLine
  end
  object lblConnectStatus: TLabel
    Left = 8
    Top = 86
    Width = 80
    Height = 16
    Caption = 'Disconnect'
    Font.Charset = ANSI_CHARSET
    Font.Color = clRed
    Font.Height = -13
    Font.Name = 'Verdana'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Label9: TLabel
    Left = 13
    Top = 471
    Width = 124
    Height = 14
    Caption = 'Acknowledgement :'
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Verdana'
    Font.Style = []
    ParentFont = False
  end
  object Bevel2: TBevel
    Left = 142
    Top = 476
    Width = 235
    Height = 3
    Shape = bsBottomLine
  end
  object txtDvnIp: TcxTextEdit
    Left = 280
    Top = 54
    Style.LookAndFeel.NativeStyle = True
    StyleDisabled.LookAndFeel.NativeStyle = True
    StyleFocused.LookAndFeel.NativeStyle = True
    StyleHot.LookAndFeel.NativeStyle = True
    TabOrder = 0
    Width = 121
  end
  object txtDvnPort: TcxTextEdit
    Left = 281
    Top = 83
    Style.LookAndFeel.NativeStyle = True
    StyleDisabled.LookAndFeel.NativeStyle = True
    StyleFocused.LookAndFeel.NativeStyle = True
    StyleHot.LookAndFeel.NativeStyle = True
    TabOrder = 1
    Text = '2100'
    Width = 60
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 504
    Height = 45
    Align = alTop
    BevelOuter = bvNone
    BorderWidth = 3
    TabOrder = 2
    object Panel2: TPanel
      Left = 3
      Top = 3
      Width = 498
      Height = 39
      Align = alClient
      BevelOuter = bvNone
      Color = clBtnShadow
      ParentBackground = False
      TabOrder = 0
      object Label8: TLabel
        Left = 20
        Top = 11
        Width = 363
        Height = 18
        Caption = 'DVN SMS Gateway Simple Command Test'
        Font.Charset = ANSI_CHARSET
        Font.Color = clWhite
        Font.Height = -16
        Font.Name = 'Verdana'
        Font.Style = [fsBold]
        ParentFont = False
      end
    end
  end
  object btnConnect: TcxButton
    Left = 411
    Top = 53
    Width = 85
    Height = 51
    Caption = 'Connect'
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Verdana'
    Font.Style = []
    ParentFont = False
    TabOrder = 3
    OnClick = btnConnectClick
  end
  object gbCmdHeader: TcxGroupBox
    Left = 10
    Top = 132
    Caption = ' Command Header '
    Style.LookAndFeel.NativeStyle = False
    StyleDisabled.LookAndFeel.NativeStyle = False
    StyleFocused.LookAndFeel.NativeStyle = False
    StyleHot.LookAndFeel.NativeStyle = False
    TabOrder = 4
    Height = 152
    Width = 485
    object Label3: TLabel
      Left = 10
      Top = 28
      Width = 76
      Height = 14
      Caption = 'Smart Card:'
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlue
      Font.Height = -12
      Font.Name = 'Verdana'
      Font.Style = []
      ParentFont = False
    end
    object Label4: TLabel
      Left = 255
      Top = 28
      Width = 49
      Height = 14
      Caption = 'STB No:'
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlue
      Font.Height = -12
      Font.Name = 'Verdana'
      Font.Style = []
      ParentFont = False
    end
    object Label5: TLabel
      Left = 15
      Top = 58
      Width = 71
      Height = 14
      Caption = 'Area Code:'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'Verdana'
      Font.Style = []
      ParentFont = False
    end
    object Label6: TLabel
      Left = 15
      Top = 89
      Width = 69
      Height = 14
      Caption = 'Start Time:'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'Verdana'
      Font.Style = []
      ParentFont = False
    end
    object Label7: TLabel
      Left = 228
      Top = 89
      Width = 76
      Height = 14
      Caption = 'Expiry Time:'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'Verdana'
      Font.Style = []
      ParentFont = False
    end
    object Label10: TLabel
      Left = 235
      Top = 59
      Width = 65
      Height = 14
      Caption = 'Frame No:'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'Verdana'
      Font.Style = []
      ParentFont = False
    end
    object Label36: TLabel
      Left = 18
      Top = 119
      Width = 65
      Height = 14
      Caption = 'AccountNo:'
      Transparent = True
    end
    object Label39: TLabel
      Left = 196
      Top = 119
      Width = 129
      Height = 14
      Caption = 'Name of SMS Gateway:'
      Transparent = True
    end
    object txtSC: TcxTextEdit
      Left = 88
      Top = 25
      Style.LookAndFeel.NativeStyle = True
      StyleDisabled.LookAndFeel.NativeStyle = True
      StyleFocused.LookAndFeel.NativeStyle = True
      StyleHot.LookAndFeel.NativeStyle = True
      TabOrder = 0
      Text = '8888888888888888'
      Width = 130
    end
    object txtSTB: TcxTextEdit
      Left = 308
      Top = 25
      Style.LookAndFeel.NativeStyle = True
      StyleDisabled.LookAndFeel.NativeStyle = True
      StyleFocused.LookAndFeel.NativeStyle = True
      StyleHot.LookAndFeel.NativeStyle = True
      TabOrder = 1
      Text = '12345678'
      Width = 130
    end
    object txtAreaCode: TcxTextEdit
      Left = 88
      Top = 55
      Style.LookAndFeel.NativeStyle = True
      StyleDisabled.LookAndFeel.NativeStyle = True
      StyleFocused.LookAndFeel.NativeStyle = True
      StyleHot.LookAndFeel.NativeStyle = True
      TabOrder = 2
      Text = '0-0-0-0'
      Width = 130
    end
    object txtStartTime: TcxTextEdit
      Left = 88
      Top = 85
      Style.LookAndFeel.NativeStyle = True
      StyleDisabled.LookAndFeel.NativeStyle = True
      StyleFocused.LookAndFeel.NativeStyle = True
      StyleHot.LookAndFeel.NativeStyle = True
      TabOrder = 3
      Text = '2005/07/01 00:00:00'
      Width = 129
    end
    object txtExpiryTime: TcxTextEdit
      Left = 308
      Top = 85
      Style.LookAndFeel.NativeStyle = True
      StyleDisabled.LookAndFeel.NativeStyle = True
      StyleFocused.LookAndFeel.NativeStyle = True
      StyleHot.LookAndFeel.NativeStyle = True
      TabOrder = 4
      Width = 130
    end
    object txtFrameNumber: TcxTextEdit
      Left = 308
      Top = 55
      Enabled = False
      Style.LookAndFeel.NativeStyle = True
      StyleDisabled.LookAndFeel.NativeStyle = True
      StyleFocused.LookAndFeel.NativeStyle = True
      StyleHot.LookAndFeel.NativeStyle = True
      TabOrder = 5
      Width = 130
    end
    object txtGlobalAccount: TcxTextEdit
      Left = 88
      Top = 116
      Style.LookAndFeel.NativeStyle = True
      StyleDisabled.LookAndFeel.NativeStyle = True
      StyleFocused.LookAndFeel.NativeStyle = True
      StyleHot.LookAndFeel.NativeStyle = True
      TabOrder = 6
      Width = 98
    end
    object txtSMSName: TcxTextEdit
      Left = 339
      Top = 116
      Style.LookAndFeel.NativeStyle = True
      StyleDisabled.LookAndFeel.NativeStyle = True
      StyleFocused.LookAndFeel.NativeStyle = True
      StyleHot.LookAndFeel.NativeStyle = True
      TabOrder = 7
      Text = 'Central'
      Width = 98
    end
  end
  object cmdPage: TcxPageControl
    Left = 10
    Top = 297
    Width = 486
    Height = 159
    ActivePage = cxTabSheet5
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Verdana'
    Font.Style = []
    LookAndFeel.NativeStyle = True
    ParentFont = False
    TabOrder = 5
    ClientRectBottom = 156
    ClientRectLeft = 3
    ClientRectRight = 483
    ClientRectTop = 23
    object cxTabSheet1: TcxTabSheet
      Caption = 'PAIRING_STB'
      ImageIndex = 0
      object Label11: TLabel
        Left = 24
        Top = 22
        Width = 71
        Height = 13
        Caption = 'ClientName:'
        Transparent = True
      end
      object Label12: TLabel
        Left = 236
        Top = 21
        Width = 65
        Height = 13
        Caption = 'AccountNo:'
        Transparent = True
      end
      object Label13: TLabel
        Left = 16
        Top = 50
        Width = 79
        Height = 13
        Caption = 'TelephoneNo:'
        Transparent = True
      end
      object Label14: TLabel
        Left = 251
        Top = 50
        Width = 50
        Height = 13
        Caption = 'EmailID:'
        Transparent = True
      end
      object Label15: TLabel
        Left = 43
        Top = 81
        Width = 51
        Height = 13
        Caption = 'Address:'
        Transparent = True
      end
      object txtClientName: TcxTextEdit
        Left = 97
        Top = 18
        Style.LookAndFeel.NativeStyle = True
        StyleDisabled.LookAndFeel.NativeStyle = True
        StyleFocused.LookAndFeel.NativeStyle = True
        StyleHot.LookAndFeel.NativeStyle = True
        TabOrder = 0
        Width = 130
      end
      object txtAccountNo: TcxTextEdit
        Left = 306
        Top = 18
        Enabled = False
        Style.LookAndFeel.NativeStyle = True
        StyleDisabled.LookAndFeel.NativeStyle = True
        StyleFocused.LookAndFeel.NativeStyle = True
        StyleHot.LookAndFeel.NativeStyle = True
        TabOrder = 1
        Text = 'using global account'
        Width = 130
      end
      object txtTelNo: TcxTextEdit
        Left = 97
        Top = 47
        Style.LookAndFeel.NativeStyle = True
        StyleDisabled.LookAndFeel.NativeStyle = True
        StyleFocused.LookAndFeel.NativeStyle = True
        StyleHot.LookAndFeel.NativeStyle = True
        TabOrder = 2
        Width = 130
      end
      object txtEmail: TcxTextEdit
        Left = 306
        Top = 46
        Style.LookAndFeel.NativeStyle = True
        StyleDisabled.LookAndFeel.NativeStyle = True
        StyleFocused.LookAndFeel.NativeStyle = True
        StyleHot.LookAndFeel.NativeStyle = True
        TabOrder = 3
        Width = 130
      end
      object txtAddress: TcxTextEdit
        Left = 97
        Top = 78
        Style.LookAndFeel.NativeStyle = True
        StyleDisabled.LookAndFeel.NativeStyle = True
        StyleFocused.LookAndFeel.NativeStyle = True
        StyleHot.LookAndFeel.NativeStyle = True
        TabOrder = 4
        Width = 340
      end
    end
    object cxTabSheet2: TcxTabSheet
      Caption = 'STB_ON'
      ImageIndex = 1
      object Label16: TLabel
        Left = 40
        Top = 56
        Width = 118
        Height = 13
        Caption = 'No Need Parameters'
        Transparent = True
      end
    end
    object cxTabSheet3: TcxTabSheet
      Caption = 'STB_OFF'
      ImageIndex = 2
      object Label17: TLabel
        Left = 40
        Top = 56
        Width = 118
        Height = 13
        Caption = 'No Need Parameters'
        Transparent = True
      end
    end
    object cxTabSheet4: TcxTabSheet
      Caption = 'ADD_PRODUCT'
      ImageIndex = 3
      object Label18: TLabel
        Left = 66
        Top = 24
        Width = 62
        Height = 13
        Caption = 'BankType:'
        Transparent = True
      end
      object Label19: TLabel
        Left = 209
        Top = 24
        Width = 76
        Height = 13
        Caption = 'ProductType:'
        Transparent = True
      end
      object Label20: TLabel
        Left = 14
        Top = 55
        Width = 114
        Height = 13
        Caption = 'Number Of Product:'
        Transparent = True
      end
      object Label21: TLabel
        Left = 197
        Top = 55
        Width = 102
        Height = 13
        Caption = 'Name Of Product:'
        Transparent = True
      end
      object Label22: TLabel
        Left = 15
        Top = 85
        Width = 138
        Height = 13
        Caption = 'Name Of SMS Gateway:'
        Transparent = True
      end
      object txtBankType: TcxTextEdit
        Left = 133
        Top = 21
        Style.LookAndFeel.NativeStyle = True
        StyleDisabled.LookAndFeel.NativeStyle = True
        StyleFocused.LookAndFeel.NativeStyle = True
        StyleHot.LookAndFeel.NativeStyle = True
        TabOrder = 0
        Text = '0'
        Width = 48
      end
      object cmbProductType: TcxComboBox
        Left = 292
        Top = 21
        Properties.DropDownListStyle = lsFixedList
        Properties.Items.Strings = (
          'GP'
          'PP')
        TabOrder = 1
        Text = 'GP'
        Width = 77
      end
      object txtNumberProduct: TcxTextEdit
        Left = 133
        Top = 51
        Style.LookAndFeel.NativeStyle = True
        StyleDisabled.LookAndFeel.NativeStyle = True
        StyleFocused.LookAndFeel.NativeStyle = True
        StyleHot.LookAndFeel.NativeStyle = True
        TabOrder = 2
        Text = '1'
        Width = 48
      end
      object txtProductName: TcxTextEdit
        Left = 308
        Top = 52
        Style.LookAndFeel.NativeStyle = True
        StyleDisabled.LookAndFeel.NativeStyle = True
        StyleFocused.LookAndFeel.NativeStyle = True
        StyleHot.LookAndFeel.NativeStyle = True
        TabOrder = 3
        Text = '310'
        Width = 152
      end
      object txtGatewayName: TcxTextEdit
        Left = 159
        Top = 81
        Enabled = False
        Style.LookAndFeel.NativeStyle = True
        StyleDisabled.LookAndFeel.NativeStyle = True
        StyleFocused.LookAndFeel.NativeStyle = True
        StyleHot.LookAndFeel.NativeStyle = True
        TabOrder = 4
        Text = 'using global sms name'
        Width = 166
      end
    end
    object cxTabSheet5: TcxTabSheet
      Caption = 'REMOVE_PRODUCT'
      ImageIndex = 4
      object Label23: TLabel
        Left = 66
        Top = 24
        Width = 62
        Height = 13
        Caption = 'BankType:'
        Transparent = True
      end
      object Label24: TLabel
        Left = 209
        Top = 24
        Width = 76
        Height = 13
        Caption = 'ProductType:'
        Transparent = True
      end
      object Label25: TLabel
        Left = 14
        Top = 55
        Width = 114
        Height = 13
        Caption = 'Number Of Product:'
        Transparent = True
      end
      object Label26: TLabel
        Left = 197
        Top = 55
        Width = 102
        Height = 13
        Caption = 'Name Of Product:'
        Transparent = True
      end
      object Label27: TLabel
        Left = 15
        Top = 85
        Width = 138
        Height = 13
        Caption = 'Name Of SMS Gateway:'
        Transparent = True
      end
      object txtBankType2: TcxTextEdit
        Left = 133
        Top = 21
        Style.LookAndFeel.NativeStyle = True
        StyleDisabled.LookAndFeel.NativeStyle = True
        StyleFocused.LookAndFeel.NativeStyle = True
        StyleHot.LookAndFeel.NativeStyle = True
        TabOrder = 0
        Text = '0'
        Width = 48
      end
      object cmbProductType2: TcxComboBox
        Left = 292
        Top = 21
        Properties.DropDownListStyle = lsFixedList
        Properties.Items.Strings = (
          'GP'
          'PP')
        TabOrder = 1
        Text = 'GP'
        Width = 77
      end
      object txtNumberProduct2: TcxTextEdit
        Left = 133
        Top = 51
        Style.LookAndFeel.NativeStyle = True
        StyleDisabled.LookAndFeel.NativeStyle = True
        StyleFocused.LookAndFeel.NativeStyle = True
        StyleHot.LookAndFeel.NativeStyle = True
        TabOrder = 2
        Text = '1'
        Width = 48
      end
      object txtProductName2: TcxTextEdit
        Left = 308
        Top = 52
        Style.LookAndFeel.NativeStyle = True
        StyleDisabled.LookAndFeel.NativeStyle = True
        StyleFocused.LookAndFeel.NativeStyle = True
        StyleHot.LookAndFeel.NativeStyle = True
        TabOrder = 3
        Text = '310'
        Width = 152
      end
      object txtGatewayName2: TcxTextEdit
        Left = 159
        Top = 81
        Enabled = False
        Style.LookAndFeel.NativeStyle = True
        StyleDisabled.LookAndFeel.NativeStyle = True
        StyleFocused.LookAndFeel.NativeStyle = True
        StyleHot.LookAndFeel.NativeStyle = True
        TabOrder = 4
        Text = 'using global sms name'
        Width = 166
      end
    end
    object cxTabSheet6: TcxTabSheet
      Caption = 'MAIL_MESSAGE'
      ImageIndex = 5
      object Label28: TLabel
        Left = 12
        Top = 20
        Width = 97
        Height = 26
        Caption = 'Message:'#13'( Max 100 bytes)'
        Transparent = True
      end
      object txtMessage: TcxMemo
        Left = 114
        Top = 16
        Properties.ScrollBars = ssBoth
        Style.LookAndFeel.NativeStyle = True
        StyleDisabled.LookAndFeel.NativeStyle = True
        StyleFocused.LookAndFeel.NativeStyle = True
        StyleHot.LookAndFeel.NativeStyle = True
        TabOrder = 0
        Height = 99
        Width = 347
      end
    end
    object cxTabSheet7: TcxTabSheet
      Caption = 'ADD_TOKEN'
      ImageIndex = 6
      object Label29: TLabel
        Left = 44
        Top = 62
        Width = 49
        Height = 13
        Caption = 'Amount:'
        Transparent = True
      end
      object Label37: TLabel
        Left = 26
        Top = 32
        Width = 67
        Height = 13
        Caption = 'Action Flag:'
        Transparent = True
      end
      object txtAddAmount: TcxTextEdit
        Left = 101
        Top = 59
        Style.LookAndFeel.NativeStyle = True
        StyleDisabled.LookAndFeel.NativeStyle = True
        StyleFocused.LookAndFeel.NativeStyle = True
        StyleHot.LookAndFeel.NativeStyle = True
        TabOrder = 0
        Text = '0'
        Width = 96
      end
      object cmbTokenAction: TcxComboBox
        Left = 101
        Top = 29
        Properties.DropDownListStyle = lsFixedList
        Properties.Items.Strings = (
          '0.Confirm'
          '1.Request'
          '2.Direct')
        TabOrder = 1
        Text = '0.Confirm'
        Width = 96
      end
    end
    object cxTabSheet8: TcxTabSheet
      Caption = 'DEDUCT_TOKEN'
      ImageIndex = 7
      object Label30: TLabel
        Left = 43
        Top = 63
        Width = 49
        Height = 13
        Caption = 'Amount:'
        Transparent = True
      end
      object Label38: TLabel
        Left = 26
        Top = 32
        Width = 67
        Height = 13
        Caption = 'Action Flag:'
        Transparent = True
      end
      object txtDecAmount: TcxTextEdit
        Left = 100
        Top = 60
        Style.LookAndFeel.NativeStyle = True
        StyleDisabled.LookAndFeel.NativeStyle = True
        StyleFocused.LookAndFeel.NativeStyle = True
        StyleHot.LookAndFeel.NativeStyle = True
        TabOrder = 0
        Text = '0'
        Width = 96
      end
      object cmbTokenAction2: TcxComboBox
        Left = 101
        Top = 29
        Properties.DropDownListStyle = lsFixedList
        Properties.Items.Strings = (
          '0.Confirm'
          '1.Request'
          '2.Direct')
        TabOrder = 1
        Text = '0.Confirm'
        Width = 96
      end
    end
    object cxTabSheet9: TcxTabSheet
      Caption = 'USER_DEFINED'
      ImageIndex = 8
      object Label31: TLabel
        Left = 20
        Top = 37
        Width = 112
        Height = 13
        Caption = 'User Define Value :'
        Transparent = True
      end
      object txtUserDefineValue: TcxTextEdit
        Left = 20
        Top = 58
        Style.LookAndFeel.NativeStyle = True
        StyleDisabled.LookAndFeel.NativeStyle = True
        StyleFocused.LookAndFeel.NativeStyle = True
        StyleHot.LookAndFeel.NativeStyle = True
        TabOrder = 0
        Width = 437
      end
    end
    object cxTabSheet10: TcxTabSheet
      Caption = 'SYS_CONFIG'
      ImageIndex = 9
      object Label32: TLabel
        Left = 24
        Top = 17
        Width = 94
        Height = 13
        Caption = 'Config Code ID:'
        Transparent = True
      end
      object Label33: TLabel
        Left = 15
        Top = 44
        Width = 103
        Height = 13
        Caption = 'Number Of Value:'
        Transparent = True
      end
      object Label34: TLabel
        Left = 47
        Top = 73
        Width = 71
        Height = 13
        Caption = 'Code Value:'
        Transparent = True
      end
      object Label35: TLabel
        Left = 15
        Top = 103
        Width = 138
        Height = 13
        Caption = 'Name Of SMS Gateway:'
        Transparent = True
      end
      object cmbConfigCode: TcxComboBox
        Left = 123
        Top = 13
        Properties.DropDownListStyle = lsFixedList
        Properties.Items.Strings = (
          '1.Key Index ( DVN Reserved )'
          '2.Area Code'
          '3.Network Code'
          '4.Service Code'
          '5.Region Code'
          '6.Zip Code'
          '7.Attributes ( DVN Reserved )')
        TabOrder = 0
        Text = '2.Area Code'
        Width = 168
      end
      object txtNumberOfValue: TcxTextEdit
        Left = 123
        Top = 41
        Style.LookAndFeel.NativeStyle = True
        StyleDisabled.LookAndFeel.NativeStyle = True
        StyleFocused.LookAndFeel.NativeStyle = True
        StyleHot.LookAndFeel.NativeStyle = True
        TabOrder = 1
        Text = '1'
        Width = 71
      end
      object txtConfigValue: TcxTextEdit
        Left = 124
        Top = 70
        Style.LookAndFeel.NativeStyle = True
        StyleDisabled.LookAndFeel.NativeStyle = True
        StyleFocused.LookAndFeel.NativeStyle = True
        StyleHot.LookAndFeel.NativeStyle = True
        TabOrder = 2
        Width = 305
      end
      object cxTextEdit13: TcxTextEdit
        Left = 159
        Top = 100
        Enabled = False
        Style.LookAndFeel.NativeStyle = True
        StyleDisabled.LookAndFeel.NativeStyle = True
        StyleFocused.LookAndFeel.NativeStyle = True
        StyleHot.LookAndFeel.NativeStyle = True
        TabOrder = 3
        Text = 'using global sms name'
        Width = 166
      end
    end
  end
  object btnSendCmd: TcxButton
    Left = 395
    Top = 464
    Width = 98
    Height = 28
    Caption = 'Send'
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Verdana'
    Font.Style = []
    ParentFont = False
    TabOrder = 6
    OnClick = btnSendCmdClick
  end
  object AckList: TcxMemo
    Left = 12
    Top = 502
    ParentFont = False
    Properties.ScrollBars = ssBoth
    Properties.WordWrap = False
    Style.Font.Charset = ANSI_CHARSET
    Style.Font.Color = clWindowText
    Style.Font.Height = -11
    Style.Font.Name = 'Verdana'
    Style.Font.Style = []
    Style.LookAndFeel.NativeStyle = True
    Style.IsFontAssigned = True
    StyleDisabled.LookAndFeel.NativeStyle = True
    StyleFocused.LookAndFeel.NativeStyle = True
    StyleHot.LookAndFeel.NativeStyle = True
    TabOrder = 7
    Height = 122
    Width = 481
  end
  object cxLookAndFeelController1: TcxLookAndFeelController
    Kind = lfFlat
    Left = 456
    Top = 12
  end
  object SocketImageList: TImageList
    Height = 32
    Width = 32
    Left = 424
    Top = 12
    Bitmap = {
      494C010102000400040020002000FFFFFFFFFF10FFFFFFFFFFFFFFFF424D3600
      0000000000003600000028000000800000002000000001002000000000000040
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000081CD3000818C8000719
      C3000719C3000719C3000B20C1000B20C1000B20C100091DBC000817B6000817
      B600080EAC00080EAC000808B600000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000000000002E832E00087E
      0900087E0900087E090009810D0009810D0009810D00087E0900087E09000872
      0800087208000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000122ED5000926E000102DDD001432
      DD001738E0001A3CE0001738E0001738E0001432DD001431D3000F29CF000B20
      C1000719C3000717C000030EBF00161DB6000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000000000005C98610020932B000D95
      1C000E991E000F9920000F9920000E991E000D951C000D951C00088712000887
      120006820E001D80220000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000263EDE000B2DEA00072BE5001238E8002149ED002954
      F4002E5AF8002E5AF8002A58F6002954F400264EED002149ED001F43E6001A3C
      E0001432DD00122ED5000719C3000413BF000717C0002D33BA00000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000109B130013961800159D220019A931001BAF
      39001EB43F001EB43F001EB43F001BAF390019A9310016A4300016A43000159D
      22000D951C000887120011821400107D11001182140000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000001238E800123FFA001D47F0002954F4003360F7003866
      FA003C69FD003C69FD003C69FD003866FA003864F7003864F7003360F7003058
      EE002C53EA002246E2001431D3000818C8000416CC00121FC100000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000013A31A000CA81B0012AB2A001EB43F002FB9470031BF
      4F002DC355002DC355002DC3550029C1530028BE500026BC4D0026BC4D0025B5
      450028AE3E0024A431001396180006820E000284070011821400000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00002246E200163AEE001641F6002552F9003866FA004171FD004675FD004876
      FE004876FE004876FE004675FD004876FE004675FD004675FD004675FD004675
      FD004572FA003C67F4002246E2000F2CD300081ECA000719C3000719C3002D33
      BA00000000000000000000000000000000000000000000000000000000000000
      00000000000013A81A0013A81A0018B3340026BC4D0034C2560049C85E0047CC
      64003CD16B0036D26D0036D26D0036D26D0036D26D0036D26D0035D16A0039CE
      660047CC640041BD50001FA42F000790160005880E0006820E0009810D000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00001239F5001843FC002A56F9003866FA004675FD004675FD004A79FE004A79
      FE005080FE005080FE005080FE005080FE005080FE00507EFF00507EFF004A79
      FE004A79FE004675FD00345DEF002448E2001432DD00081ECA000416CC000F1B
      C100000000000000000000000000000000000000000000000000000000000000
      000011AB140012B121001EB43F0021C04E002CCC610039CE660045CE680045D1
      6C003CD570003CD570003CD570003CD570003CD570003CD570003CD570003CD5
      700045CE680047CC64002BBA4C0018A935000F99200005880E0006820E001881
      1900000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000000000000000000000000000163A
      EE00123FFA002A56F9004171FD004876FE004876FE004A79FE004A79FE00507E
      FF005080FE005A8CFF005A8CFF005080FE005A8CFF005A8CFF005788FE005A8C
      FF004A79FE004A79FE004675FD003C67F4002448E2000D2AD8000818C8000717
      C0001219BB0000000000000000000000000000000000000000000000000018A8
      1F0013A81A001DBA3B0029C85C0035D16A0036D26D003CD16B003CD570003CD5
      70003CD5700039D5710039D5710039D5710039D5710039D571003CD570003CD5
      70003CD5700036D26D0035D16A002BC55B0019A931000D951C00028407001881
      1900000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000082DF7001843
      FC002A56F9003C69FD004876FE004A79FE004A79FE004A79FE004A79FE005080
      FE005A8CFF005A8CFF005A8CFF005A8CFF005A8CFF005A8CFF005A8CFF005080
      FE005788FE005080FE005080FE004675FD003058EE000F2FE6000F2CD3000719
      C3000717C0001012B8000000000000000000000000000000000008A50F0012B1
      210023B93C002DC3550032CE680036D26D003CD5700039D571003CD570003CD5
      70003CD5700039D571003CD5700039D571003CD5700039D5710039D5710039D5
      71003CD570003CD570003CD5700035D16A001BAF39000A981F00098D15001384
      19005F9161000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000082DF700244D
      FE004876FE00507EFF004A79FE004A79FE004A79FE004A79FE004A79FE00507E
      FF005788FE005A8CFF005A8CFF005A8CFF005A8CFF005788FE005A8CFF005A8C
      FF005A8CFF005080FE005A8CFF004A79FE002A58F600163EEB001432DD000F2C
      D3000717C0000817B600000000000000000000000000000000000A9E12001FB3
      35003CC8600039CE66003CD16B0039D5710039D5710039D5710039D5710039D5
      71003DD8730039D571003DD8730039D571003DD8730039D571003DD8730039D5
      71003DD8730039D5710039D5710035D16A001BB33F000CA026000D951C000D8D
      19001D8022000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000000000000D34F700315B
      FF005D89FF005788FE005080FE00507EFF00507EFF00507EFF00507EFF005080
      FE005788FE005A8CFF005A8CFF005A8CFF005A8CFF005A8CFF005A8CFF005A8C
      FF005A8CFF005A8CFF005A8CFF005080FE003864F7001C48F200193FE5001432
      DD000B20C100161DB6000000000000000000000000000000000012A220002FB9
      47004BD3740044D3730039D5710036D26D0039D571003DD8730039D5710039D5
      710039D571003DD8730039D5710039D5710039D571003DD8730039D5710039D5
      710039D571003DD8730039D5710035D16A001EBD490010A92E000F9F27000D95
      1C0006820E00087E090000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000000000001641F6003C69
      FD006F97FF006995FF005A8CFF005788FE005788FE005788FE005788FE005A8C
      FF005788FE005A8CFF005A8CFF005788FE005A8CFF005A8CFF005A8CFF005A8C
      FF005A8CFF005A8CFF005A8CFF005080FE003769FD002552F9001C4CEA001A3C
      E000081ECA00101EB5000000000000000000000000000000000019A9310034C2
      560052DB83004CDA7D003CD5700039D5710039D5710039D571003DD8730039D5
      710039D5710039D5710039D571003DD8730039D5710039D5710039D571003DD8
      730039D5710039D5710039D5710035D16A001EBD490014B3370015A62F000F99
      200008871200107D110000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000214EF7004876
      FE0079A1FF00719AFF005D89FF005D89FF005A8CFF005A8CFF005A8CFF005A8C
      FF005A8CFF005788FE005A8CFF005A8CFF005A8CFF005788FE005A8CFF005A8C
      FF005788FE005A8CFF005A8CFF005080FE003C69FD002859FB00264EED001F43
      E6000F29CF001022BA000000000000000000000000000000000026B23F003CC8
      60005BDD8A0052DB83003CD570003CD570003CD5700039D5710039D571003DD8
      730039D571003DD8730039D5710039D571003DD8730039D571003DD8730039D5
      710039D571003DD8730039D5710036D26D0029C1530017B73D0017AC3500149F
      2A000D8D1900107D110000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000000000003360F700507E
      FF007DA6FF0079A1FF006A98FF006191FF005A8CFF005A8CFF005A8CFF005788
      FE005A8CFF005A8CFF005A8CFF005788FE005A8CFF005A8CFF005A8CFF005A8C
      FF005A8CFF005A8CFF005A8CFF005080FE003E70FF002E5AF8002954F4002448
      E2001432DD001227BE000000000000000000000000000000000030B7490045CE
      68005BDD8A0052DB830042DB75003DD8730039D5710039D5710039D5710039D5
      71003DD8730039D5710039D571003DD8730039D5710039D5710039D5710039D5
      71003DD8730039D5710039D5710036D26D0029C153001ABA430014B3370015A6
      2F000D951C001182140000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000000000003360F7005788
      FE007DA6FF0079A1FF006F9CFF006394FF005A8CFF005A8CFF005A8CFF005A8C
      FF005A8CFF005A8CFF005A8CFF005A8CFF005A8CFF005788FE005A8CFF005A8C
      FF005788FE005A8CFF005A8CFF005080FE003E70FF002F60FF002A58F600264E
      ED001432DD001227BE000000000000000000000000000000000035BC52004BD3
      740067E3920059E3870042DD78003DD873003CD570003DD8730039D5710039D5
      710039D5710039D571003DD8730039D5710039D5710039D571003DD8730039D5
      710039D5710039D5710039D5710036D26D0025C453001ABA43001BB33F0018A9
      35000D951C001182140000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000000000003C67F400507E
      FF0079A1FF0079A1FF0076A1FF006995FF005A8CFF005A8CFF005A8CFF005A8C
      FF005788FE005D89FF005A8CFF005A8CFF005A8CFF005A8CFF005788FE005A8C
      FF005A8CFF005A8CFF005A8CFF005080FE004778FF002F60FF002A58F6002B59
      F0001738E0001329C2000000000000000000000000000000000038BD550051D4
      740070E7990059E387004CE681003DD873003DD8730039D571003DD8730039D5
      71003DD8730039D5710039D5710039D5710039D571003DD8730039D5710039D5
      71003DD8730039D5710039D5710039D5710029C85C001EBD49001BB33F001BAF
      39000F9920001182140000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000000000003866FA005788
      FE007DA6FF0083AAFF0079A1FF00719AFF005A8CFF005A8CFF005A8CFF005A8C
      FF005A8CFF005A8CFF005A8CFF005788FE005A8CFF005A8CFF005A8CFF005A8C
      FF005A8CFF005A8CFF005A8CFF005080FE004675FD003769FD003866FA003058
      EE001432DD001329C2000000000000000000000000000000000035BC520054D6
      79007BEBA10070E799004CE6810046E17C0039D5710039D5710039D571003DD8
      730039D5710039D571003DD8730039D571003DD8730039D5710039D5710039D5
      710039D571003DD8730039D5710036D26D0029C85C0021C04E001EBD490020B2
      40000F9920001086100000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000000000003360F7005D89
      FF009DC2FF0094BEFF007DA6FF00719AFF005D89FF005A8CFF005A8CFF005788
      FE005D89FF005A8CFF005A8CFF005A8CFF005A8CFF005A8CFF005788FE005D89
      FF005A8CFF005A8CFF005A8CFF005080FE004778FF003866FA003866FA002954
      F4001738E0001329C2000000000000000000000000000000000035BC520059D5
      7D0089F2AC006EF49E0053EB880046E17C003DD8730039D571003DD8730039D5
      710039D5710039D5710039D571003DD8730039D5710039D5710039D5710039D5
      71003DD8730039D5710039D5710036D26D0029C85C0021C04E001EBD49001BB3
      3F00159D22000887120000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000000000002954F4006995
      FF00A8CBFF009DC2FF0088B2FF0076A1FF006394FF005788FE005A8CFF005A8C
      FF005A8CFF005A8CFF005788FE005A8CFF005A8CFF005A8CFF005D89FF005A8C
      FF005A8CFF005788FE005A8CFF005788FE004778FF003769FD003866FA002A58
      F6001738E000081ECA00000000000000000000000000000000002CBB4E0054D9
      7B0089F2AC007CF3A5005FEF92004CE6810042DD78003DD8730039D5710039D5
      710039D5710039D571003DD8730039D5710039D5710039D5710039D5710039D5
      710039D5710039D5710039D5710039D5710029C85C0021C04E001EBD490025B5
      45000E991E001182140000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000244DFE005788
      FE009DC2FF00A8CBFF009DC8FF0082B1FF006394FF005A8CFF005A8CFF005A8C
      FF005788FE005A8CFF005A8CFF005A8CFF005788FE005D89FF005A8CFF005A8C
      FF005A8CFF005A8CFF005A8CFF005080FE005080FE004675FD003864F7003058
      EE001234E1000F29CF000000000000000000000000000000000026BC4D0051D4
      740082EDA4007CF3A50071F5A0005FEF92004CE6810042DD78003DD8730039D5
      710039D5710039D5710039D5710039D5710039D5710039D5710039D571003DD8
      730039D571003DD8730039D5710036D26D002CCC610025C453001EBD49001BAF
      39000E991E001086100000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000000000001C48F2004876
      FE0083AAFF00A8CBFF00ADDBFF008CB9FF006394FF005A8CFF005A8CFF005A8C
      FF005A8CFF005A8CFF005788FE005D89FF005D89FF005A8CFF005A8CFF005A8C
      FF005A8CFF005A8CFF005A8CFF005A8CFF004A79FE004675FD003360F7001D47
      F0000F2FE6001026D4000000000000000000000000000000000045AA5E0055C3
      730072E495007BEBA10084FDB0006EF49E0053EB880046E17C003DD873003DD8
      730039D571003DD8730039D571003DD8730039D571003DD8730039D5710039D5
      71003DD8730039D5710039D5710039D571002CCC610025C453001EBD49001FA4
      2F0020932B003C8C3C0000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000000000000000000000000000002552F9003C69
      FD006394FF0088B2FF00B9E3FF0099C4FF006394FF005A8CFF005A8CFF005788
      FE005A8CFF005D89FF005A8CFF005D89FF005A8CFF005A8CFF005A8CFF005A8C
      FF005788FE005A8CFF005A8CFF005788FE005080FE004675FD002954F4001036
      EB00082DF7001028E100000000000000000000000000000000000000000078B1
      890059D57D0073EC9B0093FABA007CF3A50053EB88004CE6810046E17C0046E1
      7C003DD8730039D5710039D5710039D571003DD8730039D5710039D571003DD8
      730039D5710039D5710039D5710039D5710036D26D0029C85C0014B337001B9E
      2700659C68000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000003D69
      EA004675FD0076A1FF00B4DBFF00A8CBFF0076A1FF006394FF006394FF005D89
      FF005A8CFF005A8CFF005A8CFF005A8CFF005A8CFF005A8CFF005A8CFF005A8C
      FF005A8CFF005788FE005A8CFF005080FE004A79FE003C69FD00214EF7000B2D
      EA001836E2000000000000000000000000000000000000000000000000000000
      000045CE680067E3920094F7B70089F2AC005FEF92005FEF920053EB88004CE6
      810042DD78003DD873003DD873003DD873003DD8730039D571003DD8730039D5
      710039D5710039D5710039D5710039D571002ECD640025C4530010A92E001A9C
      1F00000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000006B82
      C5003C67F4005A8CFF0085B2FF0097C3FF00A6D1FF0097C3FF0076A1FF006995
      FF005A8CFF005A8CFF005A8CFF005A8CFF005A8CFF005A8CFF005788FE005A8C
      FF005A8CFF005788FE005A8CFF00507EFF003C69FD002A56F900123FFA001036
      EB00000000000000000000000000000000000000000000000000000000000000
      00003CD16B0051E37E0070E799007CF3A50087F5B0007CF3A50073FAA40062F4
      960053EB88004CE6810046E17C0046E17C0046E17C003DD873003DD873003DD8
      73003DD8730039D5710039D5710032CE680029C1530017B73D000DB11B0018A4
      1A00000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000486ED9004B79EE005A8CFF008CB9FF00B9E3FF00ACD5FF0088B2FF0079A1
      FF006F9CFF006394FF005D89FF006394FF005D89FF006394FF005D89FF005D89
      FF005D89FF005A8CFF005080FE004876FE002E5AF8001343FA001D47F000344F
      D500000000000000000000000000000000000000000000000000000000000000
      00004BCB720042DB750051E37E0073EC9B0093FABA0093FABA0084FDB00073FA
      A40062F4960059F28E0053EB880053EB880046E17C0046E17C0046E17C0042DD
      780039D571003CD5700039D5710029C85C001ABA430012B1210019AB200034A5
      3400000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000007188C5004B79EE005A8CFF007CAAFF0088B2FF0092BDFF0092BD
      FF0092BDFF0092BDFF0085B2FF0082B1FF0079A1FF0076A1FF00719AFF00719A
      FF00719AFF005D89FF004675FD002552F900164AFE001D47F0005769C5000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000071BB870042DD780053EB870067E392007BEBA1007CF3A5007FF7
      AB007FF7AB0071F5A00069F097005FEF920058EC8B0058EC8B0052DB830042DD
      78004CDA7D0035D16A002BC55B001ABA430014BA2A001AB1250060A663000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000486ED9003D6BF1004A79FE006394FF008CB9FF00A8CB
      FF00A8CBFF00A8CBFF0097C3FF008CB9FF0085B2FF007CAAFF0076A1FF00719A
      FF00719AFF005D89FF003866FA00123FFA001044F0003453D500000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000004BCB720045D16C003FD76E0059E3870073EC9B0082F0
      A60089F2AC0089F2AC006EF49E0067E3920059E3870059E387005BDD8A004CDA
      7D0044D3730038C85D0025B545001AB52C001AB1250036AC3C00000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000006B82C5003D69EA004675FD004A79FE005A8C
      FF006394FF006394FF006394FF006394FF005A8CFF005080FE004A79FE004675
      FD003866FA002E5AF800164AFE001C4CEA005F74C30000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      00000000000000000000000000006DB6830039CE66003CD16B004BD3740054D9
      7B0054D97B0052DB830054D97B0054D97B004BD3740044D3730039CE660038C8
      5D002CBB4E0023B93C0019BD300025B5340063A9680000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000486ED9002954F4003058EE002A58
      F6003861F0003D6BEE004871F0004871F000456EF000456EF0003360F7002149
      ED001C48F2001641F6001044F000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000000000002BC55B002DBA560033B9
      59003CBC5C0044BE620044BE620044BE620045C065003CBC5C0037B551002DAF
      450028AE3E001DA92F001AB52C00000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      0000000000000000000000000000000000000000000000000000000000000000
      000000000000000000000000000000000000424D3E000000000000003E000000
      2800000080000000200000000100010000000000000200000000000000000000
      000000000000000000000000FFFFFF00FFFFFFFFFFFFFFFF0000000000000000
      FFFFFFFFFFFFFFFF0000000000000000FF8001FFFFC007FF0000000000000000
      FF0000FFFF8003FF0000000000000000FC00003FFE00007F0000000000000000
      FC00003FFC00003F0000000000000000F000000FF800001F0000000000000000
      F000000FF000000F0000000000000000E0000007E000000F0000000000000000
      C0000003C00000070000000000000000C0000003C00000070000000000000000
      C0000003C00000030000000000000000C0000003C00000030000000000000000
      C0000003C00000030000000000000000C0000003C00000030000000000000000
      C0000003C00000030000000000000000C0000003C00000030000000000000000
      C0000003C00000030000000000000000C0000003C00000030000000000000000
      C0000003C00000030000000000000000C0000003C00000030000000000000000
      C0000003C00000030000000000000000C0000003E00000070000000000000000
      E0000007F000000F0000000000000000E000000FF000000F0000000000000000
      F000000FF000000F0000000000000000F800001FF800001F0000000000000000
      FC00003FFC00003F0000000000000000FE00007FFE00007F0000000000000000
      FF0001FFFF8001FF0000000000000000FFFFFFFFFFFFFFFF0000000000000000
      FFFFFFFFFFFFFFFF000000000000000000000000000000000000000000000000
      000000000000}
  end
  object IdSocket: TIdTCPClient
    MaxLineAction = maException
    Port = 0
    Left = 392
    Top = 12
  end
end
