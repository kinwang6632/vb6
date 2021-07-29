object frmChoiceRpt: TfrmChoiceRpt
  Left = 357
  Top = 195
  BorderStyle = bsDialog
  Caption = #22577#34920#36984#25799
  ClientHeight = 254
  ClientWidth = 389
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnClose = FormClose
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 16
  object Panel2: TPanel
    Left = 0
    Top = 213
    Width = 389
    Height = 41
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 0
    DesignSize = (
      389
      41)
    object btnOK: TcxButton
      Left = 221
      Top = 9
      Width = 75
      Height = 25
      Anchors = [akRight, akBottom]
      Caption = #30906#35469
      Default = True
      ModalResult = 1
      TabOrder = 0
    end
    object btnCancel: TcxButton
      Left = 304
      Top = 9
      Width = 75
      Height = 25
      Anchors = [akRight, akBottom]
      Caption = #21462#28040
      ModalResult = 2
      TabOrder = 1
    end
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 389
    Height = 213
    Align = alClient
    BevelOuter = bvNone
    BorderWidth = 2
    TabOrder = 1
    object ClientBox: TScrollBox
      Left = 2
      Top = 2
      Width = 385
      Height = 209
      HorzScrollBar.Style = ssHotTrack
      VertScrollBar.Style = ssHotTrack
      Align = alClient
      BevelEdges = [beBottom]
      BevelInner = bvNone
      BevelKind = bkFlat
      BorderStyle = bsNone
      TabOrder = 0
      object Image1: TImage
        Left = 4
        Top = 4
        Width = 65
        Height = 49
      end
    end
  end
end
