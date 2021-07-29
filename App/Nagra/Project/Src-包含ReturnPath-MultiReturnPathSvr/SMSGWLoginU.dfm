object frmSMSGW0000A: TfrmSMSGW0000A
  Left = 428
  Top = 226
  Width = 356
  Height = 202
  AutoSize = True
  BorderIcons = []
  Caption = '系統登入[SMSGW0000A]'
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Times New Roman'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  Visible = True
  PixelsPerInch = 96
  TextHeight = 15
  object bvlLogin: TBevel
    Left = 0
    Top = 72
    Width = 348
    Height = 103
    Align = alClient
  end
  object lblUserID: TLabel
    Left = 97
    Top = 84
    Width = 60
    Height = 15
    Caption = '使用者代碼'
    Color = clBtnFace
    Font.Charset = ANSI_CHARSET
    Font.Color = clBlack
    Font.Height = -12
    Font.Name = 'Times New Roman'
    Font.Style = []
    ParentColor = False
    ParentFont = False
  end
  object lblPassword: TLabel
    Left = 121
    Top = 116
    Width = 36
    Height = 15
    Caption = '密　碼'
    Color = clBtnFace
    Font.Charset = ANSI_CHARSET
    Font.Color = clBlack
    Font.Height = -12
    Font.Name = 'Times New Roman'
    Font.Style = []
    ParentColor = False
    ParentFont = False
  end
  object lblTitle1: TLabel
    Left = 0
    Top = 0
    Width = 348
    Height = 65
    Align = alTop
    Alignment = taCenter
    AutoSize = False
    Caption = '   SMS Gateway for VIACESS V1.0'
    Color = clWhite
    ParentColor = False
  end
  object lblTitle2: TLabel
    Left = 0
    Top = 65
    Width = 348
    Height = 7
    Align = alTop
    AutoSize = False
    Color = clTeal
    ParentColor = False
  end
  object edtUserID: TEdit
    Left = 171
    Top = 80
    Width = 166
    Height = 23
    TabOrder = 0
  end
  object edtPassword: TEdit
    Left = 171
    Top = 112
    Width = 166
    Height = 23
    PasswordChar = '*'
    TabOrder = 1
  end
  object btnOK: TBitBtn
    Left = 192
    Top = 144
    Width = 65
    Height = 25
    Caption = '確定(&Y)'
    Default = True
    TabOrder = 2
    OnClick = btnOKClick
  end
  object btnCancel: TBitBtn
    Left = 272
    Top = 144
    Width = 65
    Height = 25
    Caption = '取消(&C)'
    TabOrder = 3
    OnClick = btnCancelClick
  end
  object qryUserInfo: TQuery
    DatabaseName = 'SMSGW1_DB'
    SQL.Strings = (
      'select * from USER_INFO'
      'where USER_ID = :USER_ID'
      'AND PASSWD = :PASSWD')
    Left = 8
    Top = 112
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'USER_ID'
        ParamType = ptUnknown
      end
      item
        DataType = ftUnknown
        Name = 'PASSWD'
        ParamType = ptUnknown
      end>
  end
end
