object frmCD078B3_7: TfrmCD078B3_7
  Left = 288
  Top = 320
  BorderStyle = bsDialog
  Caption = 'frmCD078B3_7'
  ClientHeight = 206
  ClientWidth = 257
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poOwnerFormCenter
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label5: TLabel
    Left = 23
    Top = 13
    Width = 48
    Height = 14
    Caption = #26376#25976#22411#24907
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object Label1: TLabel
    Left = 24
    Top = 64
    Width = 48
    Height = 14
    Caption = #32371#36027#26399#25976
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object Label3: TLabel
    Left = 24
    Top = 92
    Width = 48
    Height = 14
    Caption = #20027#25512#26399#25976
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object Label6: TLabel
    Left = 24
    Top = 117
    Width = 48
    Height = 14
    Caption = #36027#29575#20381#25818
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object Label7: TLabel
    Left = 24
    Top = 144
    Width = 48
    Height = 14
    Caption = #21934#26376#37329#38989
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object Label2: TLabel
    Left = 24
    Top = 37
    Width = 48
    Height = 14
    Caption = #20351#29992#26376#25976
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object cmbUseMonth: TcxComboBox
    Left = 75
    Top = 9
    ParentFont = False
    Properties.DropDownListStyle = lsEditFixedList
    Properties.Items.Strings = (
      '1.'#33287#32129#32004#26376#25976#21516
      '2.'#33287#32371#36027#26399#25976#21516
      '3.'#33258#34892#35373#23450)
    Properties.OnChange = cmbUseMonthPropertiesChange
    Style.StyleController = StyleController.EditorStyle
    TabOrder = 0
    Text = '1.'#33287#32129#32004#26376#25976#21516
    Width = 160
  end
  object edtBillPeriod: TcxCurrencyEdit
    Left = 75
    Top = 60
    ParentFont = False
    Properties.Alignment.Horz = taRightJustify
    Properties.AssignedValues.MinValue = True
    Properties.DecimalPlaces = 0
    Properties.DisplayFormat = ',0;-,0'
    Properties.MaxValue = 999.000000000000000000
    Properties.NullString = '0'
    Style.StyleController = StyleController.EditorStyle
    TabOrder = 2
    Width = 41
  end
  object edtMasterPeriod: TcxCurrencyEdit
    Left = 75
    Top = 86
    ParentFont = False
    Properties.Alignment.Horz = taRightJustify
    Properties.AssignedValues.MinValue = True
    Properties.DecimalPlaces = 0
    Properties.DisplayFormat = ',0;-,0'
    Properties.MaxValue = 999.000000000000000000
    Properties.NullString = '0'
    Style.StyleController = StyleController.EditorStyle
    TabOrder = 3
    Width = 41
  end
  object cmbBillType: TcxComboBox
    Left = 75
    Top = 113
    ParentFont = False
    Properties.DropDownListStyle = lsEditFixedList
    Properties.Items.Strings = (
      '1.'#20381#36027#29575#34920
      '2.'#20381#20778#24800#37329#38989
      '')
    Style.StyleController = StyleController.EditorStyle
    TabOrder = 4
    Text = '2.'#20381#20778#24800#37329#38989
    Width = 160
  end
  object edtMonthAmt1: TcxCurrencyEdit
    Left = 183
    Top = 138
    EditValue = 0
    ParentFont = False
    Properties.Alignment.Horz = taRightJustify
    Properties.AssignedValues.MinValue = True
    Properties.DecimalPlaces = 3
    Properties.DisplayFormat = ',0.000;-,0.000'
    Properties.MaxValue = 9999999.999000000000000000
    Properties.NullString = '0'
    Style.StyleController = StyleController.EditorStyle
    TabOrder = 5
    Visible = False
    Width = 75
  end
  object btnSave: TButton
    Left = 23
    Top = 174
    Width = 69
    Height = 22
    Action = actSave
    Default = True
    TabOrder = 6
  end
  object edtUseMonth: TcxCurrencyEdit
    Left = 75
    Top = 34
    ParentFont = False
    Properties.Alignment.Horz = taRightJustify
    Properties.AssignedValues.MinValue = True
    Properties.DecimalPlaces = 0
    Properties.DisplayFormat = ',0;-,0'
    Properties.MaxValue = 999.000000000000000000
    Properties.NullString = '0'
    Style.StyleController = StyleController.EditorStyle
    TabOrder = 1
    Width = 41
  end
  object btnCancel: TButton
    Left = 163
    Top = 174
    Width = 69
    Height = 23
    Action = actCancel
    Cancel = True
    ModalResult = 1
    TabOrder = 7
  end
  object edtMonthAmt: TcxCurrencyEdit
    Left = 75
    Top = 139
    EditValue = 0.000000000000000000
    ParentFont = False
    Properties.Alignment.Horz = taRightJustify
    Properties.DecimalPlaces = 3
    Properties.DisplayFormat = ',0.000;-,0.000'
    Properties.MaxValue = 9999.999000000000000000
    Properties.NullString = '0'
    Style.StyleController = StyleController.EditorStyle
    TabOrder = 8
    Width = 80
  end
  object ActionList1: TActionList
    Left = 180
    Top = 56
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
