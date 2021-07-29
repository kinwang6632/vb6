object fmGroupSet: TfmGroupSet
  Left = 286
  Top = 118
  Width = 650
  Height = 550
  BorderStyle = bsSizeToolWin
  Caption = #20351#29992#32773#32676#32068#35373#23450
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
  object Bevel1: TBevel
    Left = 0
    Top = 474
    Width = 642
    Height = 3
    Align = alBottom
    Shape = bsBottomLine
  end
  object Panel1: TPanel
    Left = 0
    Top = 477
    Width = 642
    Height = 41
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 0
    DesignSize = (
      642
      41)
    object btnConfirm: TcxButton
      Left = 465
      Top = 6
      Width = 81
      Height = 25
      Anchors = [akRight, akBottom]
      Caption = #30906#23450
      Default = True
      ModalResult = 1
      TabOrder = 0
      OnClick = btnConfirmClick
      LookAndFeel.Kind = lfFlat
      LookAndFeel.NativeStyle = True
    end
    object btnCancel: TcxButton
      Left = 552
      Top = 6
      Width = 81
      Height = 25
      Anchors = [akRight, akBottom]
      Cancel = True
      Caption = #21462#28040
      ModalResult = 2
      TabOrder = 1
      LookAndFeel.Kind = lfFlat
      LookAndFeel.NativeStyle = True
    end
  end
  object HeaderPanel: TPanel
    Left = 0
    Top = 0
    Width = 642
    Height = 42
    Align = alTop
    BevelOuter = bvNone
    Color = 14211288
    ParentBackground = False
    TabOrder = 1
    object HeaderImage: TImage
      Left = 15
      Top = 6
      Width = 32
      Height = 32
      AutoSize = True
      Transparent = True
    end
    object txtCompInfo: TcxImageComboBox
      Left = 67
      Top = 11
      Properties.Images = StyleModule.SoImgList
      Properties.ImmediatePost = True
      Properties.Items = <>
      Properties.OnChange = txtCompInfoPropertiesChange
      TabOrder = 0
      Width = 233
    end
  end
  object CH009Tree: TcxDBTreeList
    Left = 0
    Top = 42
    Width = 300
    Height = 432
    Align = alLeft
    Bands = <
      item
        Width = 283
      end>
    BufferedPaint = False
    DataController.DataSource = dsCH009
    DataController.ImageIndexFieldName = 'ImageIndex'
    DataController.ParentField = 'RGroupId'
    DataController.KeyField = 'GroupId'
    DragMode = dmAutomatic
    Images = StyleModule.ImgList16
    OptionsBehavior.DragFocusing = True
    OptionsBehavior.Sorting = False
    OptionsBehavior.DragDropText = True
    OptionsCustomizing.ColumnMoving = False
    OptionsCustomizing.ColumnVertSizing = False
    OptionsSelection.CellSelect = False
    OptionsSelection.MultiSelect = True
    OptionsView.CellEndEllipsis = True
    RootValue = -1
    TabOrder = 2
    OnCustomDrawCell = CH009TreeCustomDrawCell
    OnDragDrop = CH009TreeDragDrop
    OnDragOver = CH009TreeDragOver
    object CH009TreeCol1: TcxDBTreeListColumn
      Visible = False
      Caption.Text = #20844#21496#21029
      DataBinding.FieldName = 'CompCode'
      Width = 70
      Position.ColIndex = 0
      Position.RowIndex = 0
      Position.BandIndex = 0
    end
    object CH009TreeCol2: TcxDBTreeListColumn
      Visible = False
      Caption.Text = #20844#21496#21517#31281
      DataBinding.FieldName = 'CompName'
      Width = 70
      Position.ColIndex = 1
      Position.RowIndex = 0
      Position.BandIndex = 0
    end
    object CH009TreeCol3: TcxDBTreeListColumn
      Caption.Text = #32676#32068#20195#30908
      DataBinding.FieldName = 'GroupId'
      Width = 100
      Position.ColIndex = 3
      Position.RowIndex = 0
      Position.BandIndex = 0
    end
    object CH009TreeCol4: TcxDBTreeListColumn
      Caption.Text = #32676#32068#21517#31281
      DataBinding.FieldName = 'GroupName'
      Width = 195
      Position.ColIndex = 4
      Position.RowIndex = 0
      Position.BandIndex = 0
    end
    object CH009TreeCol5: TcxDBTreeListColumn
      Visible = False
      Caption.Text = #20572#29992
      DataBinding.FieldName = 'StopFlag'
      Position.ColIndex = 2
      Position.RowIndex = 0
      Position.BandIndex = 0
    end
  end
  object cxSplitter1: TcxSplitter
    Left = 300
    Top = 42
    Width = 8
    Height = 432
    HotZoneClassName = 'TcxXPTaskBarStyle'
    Control = CH009Tree
  end
  object CD071Tree: TcxDBTreeList
    Left = 308
    Top = 42
    Width = 334
    Height = 432
    Align = alClient
    Bands = <
      item
        Width = 304
      end>
    BufferedPaint = False
    DataController.DataSource = dsCD071
    DataController.ImageIndexFieldName = 'ImageIndex'
    DataController.ParentField = 'CodeNo'
    DataController.KeyField = 'CodeNo'
    DragMode = dmAutomatic
    Images = StyleModule.ImgList16
    OptionsBehavior.Sorting = False
    OptionsBehavior.DragDropText = True
    OptionsCustomizing.ColumnMoving = False
    OptionsCustomizing.ColumnVertSizing = False
    OptionsSelection.CellSelect = False
    OptionsSelection.MultiSelect = True
    OptionsView.CellEndEllipsis = True
    OptionsView.TreeLineStyle = tllsNone
    RootValue = -1
    TabOrder = 4
    OnCustomDrawCell = CD071TreeCustomDrawCell
    OnDragDrop = CD071TreeDragDrop
    OnDragOver = CD071TreeDragOver
    object CD071TreeCol1: TcxDBTreeListColumn
      Visible = False
      Caption.Text = #20844#21496#21029
      DataBinding.FieldName = 'CompCode'
      Width = 101
      Position.ColIndex = 0
      Position.RowIndex = 0
      Position.BandIndex = 0
    end
    object CD071TreeCol2: TcxDBTreeListColumn
      Visible = False
      Caption.Text = #20844#21496#21517#31281
      DataBinding.FieldName = 'CompName'
      Width = 194
      Position.ColIndex = 1
      Position.RowIndex = 0
      Position.BandIndex = 0
    end
    object CD071TreeCol3: TcxDBTreeListColumn
      Caption.Text = #32676#32068#20195#30908
      DataBinding.FieldName = 'CodeNo'
      Width = 100
      Position.ColIndex = 3
      Position.RowIndex = 0
      Position.BandIndex = 0
    end
    object CD071TreeCol4: TcxDBTreeListColumn
      Caption.Text = #32676#32068#21517#31281
      DataBinding.FieldName = 'Description'
      Width = 195
      Position.ColIndex = 4
      Position.RowIndex = 0
      Position.BandIndex = 0
    end
    object CD071TreeCol5: TcxDBTreeListColumn
      Visible = False
      Caption.Text = #20572#29992
      DataBinding.FieldName = 'StopFlag'
      Position.ColIndex = 2
      Position.RowIndex = 0
      Position.BandIndex = 0
    end
  end
  object dsCH009: TDataSource
    DataSet = CH009Buffer
    Left = 132
    Top = 220
  end
  object dsCD071: TDataSource
    DataSet = CD071Buffer
    Left = 452
    Top = 220
  end
  object CH009Buffer: TVirtualTable
    Left = 96
    Top = 220
    Data = {03000000000000000000}
  end
  object CD071Buffer: TVirtualTable
    Left = 416
    Top = 220
    Data = {03000000000000000000}
  end
  object LoadTimer: TTimer
    Enabled = False
    Interval = 300
    OnTimer = LoadTimerTimer
    Left = 176
    Top = 220
  end
end
