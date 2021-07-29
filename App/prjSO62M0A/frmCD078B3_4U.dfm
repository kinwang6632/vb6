object frmCD078B3_4: TfrmCD078B3_4
  Left = 490
  Top = 351
  BorderStyle = bsDialog
  Caption = #35722#26356#25910#36027#38917#30446'[frmCD078B3]'
  ClientHeight = 101
  ClientWidth = 435
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  DesignSize = (
    435
    101)
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 11
    Top = 17
    Width = 48
    Height = 14
    Caption = #25910#36027#38917#30446
    Font.Charset = ANSI_CHARSET
    Font.Color = clRed
    Font.Height = -12
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object Bevel1: TBevel
    Left = 11
    Top = 52
    Width = 414
    Height = 3
    Anchors = [akLeft, akTop, akRight]
    Shape = bsTopLine
  end
  inline CItem: TDataLookup
    Left = 63
    Top = 12
    Width = 362
    Height = 23
    HorzScrollBar.Visible = False
    VertScrollBar.Visible = False
    TabOrder = 0
    inherited CodeNo: TcxTextEdit
      Left = 2
      Properties.MaxLength = 5
      Properties.OnChange = CItemCodeNoPropertiesChange
      Width = 66
    end
    inherited CodeName: TcxLookupComboBox
      Left = 70
      Properties.OnChange = CItemCodeNamePropertiesChange
      Properties.OnInitPopup = CItemCodeNamePropertiesInitPopup
      Width = 292
    end
  end
  object btnSave: TButton
    Left = 277
    Top = 64
    Width = 69
    Height = 26
    Caption = #30906#23450
    Default = True
    Enabled = False
    TabOrder = 1
    OnClick = btnSaveClick
  end
  object Button1: TButton
    Left = 354
    Top = 64
    Width = 69
    Height = 26
    Cancel = True
    Caption = #21462#28040
    ModalResult = 2
    TabOrder = 2
  end
end
