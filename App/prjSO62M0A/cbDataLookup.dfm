object DataLookup: TDataLookup
  Left = 0
  Top = 0
  Width = 309
  Height = 25
  HorzScrollBar.Visible = False
  VertScrollBar.Visible = False
  TabOrder = 0
  DesignSize = (
    309
    25)
  object CodeNo: TcxTextEdit
    Left = 0
    Top = 1
    ParentFont = False
    Properties.OnChange = CodeNoPropertiesChange
    Style.StyleController = StyleController.EditorStyle
    TabOrder = 0
    OnExit = CodeNoExit
    Width = 60
  end
  object CodeName: TcxLookupComboBox
    Left = 62
    Top = 1
    Anchors = [akLeft, akTop, akRight]
    ParentFont = False
    Properties.DropDownSizeable = True
    Properties.ListColumns = <>
    Properties.ListOptions.ColumnSorting = False
    Properties.OnChange = CodeNamePropertiesChange
    Properties.OnInitPopup = CodeNamePropertiesInitPopup
    Style.StyleController = StyleController.EditorStyle
    TabOrder = 1
    Width = 244
  end
end
