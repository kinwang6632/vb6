unit frmMultiSelectU2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, StdCtrls, ComCtrls, ExtCtrls, DB, DBClient, ADODB, Provider, ActnList,
  cxLookAndFeelPainters, cxButtons, cxCheckBox, cxControls, cxContainer,
  cxEdit, cxTextEdit, cxRadioGroup, dxSkinsCore, dxSkinsDefaultPainters,
  dxSkinBlack, dxSkinBlue, dxSkinCaramel, dxSkinCoffee;

type
  TfrmMultiSelect2 = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    ScrollBox1: TScrollBox;
    ListView: TListView;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    Timer1: TTimer;
    ActionList1: TActionList;
    actConfirm: TAction;
    actSelectAll: TAction;
    actClear: TAction;
    actCancel: TAction;
    GroupBox1: TGroupBox;
    rdbCodeNo: TcxRadioButton;
    rdbDescription: TcxRadioButton;
    edtText: TcxTextEdit;
    chklShowSelected: TcxCheckBox;
    btnNext: TcxButton;
    btnPrior: TcxButton;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure actSelectAllExecute(Sender: TObject);
    procedure actConfirmExecute(Sender: TObject);
    procedure actClearExecute(Sender: TObject);
    procedure actCancelExecute(Sender: TObject);
    procedure edtTextPropertiesChange(Sender: TObject);
    procedure chklShowSelectedPropertiesChange(Sender: TObject);
    procedure ListViewDblClick(Sender: TObject);
    procedure btnNextClick(Sender: TObject);
    procedure btnPriorClick(Sender: TObject);
  private
    { Private declarations }
    FConnection: TADOConnection;
    FDataReader: TADOQuery;
    FProvider: TDataSetProvider;
    FDataSet: TClientDataSet;
    FSQL: TStringList;
    FKeyFields: String;
    FKeyValues: String;
    FDisplayFields: String;
    FResultFields: String;
    FSelectedValue: String;
    FSelectedDisplay: String;
    FQuotedString: String;
    FKeyFieldList: TStringList;
    FKeyValueList: TStringList;
    FDisplayFieldList: TStringList;
    FDisplayCaptionList: TStringList;
    FResultList: TStringList;
    procedure SetKeyFields(const Value: String);
    procedure SetKeyValues(const Value: String);
    procedure SetDisplayFields(const Value: String);
    procedure SetResultFields(const Value: String);
    procedure CreateDisplayColumns;
    procedure FreeupBookmark;
    procedure AddDataItem;
    procedure ClearDataItem;
    procedure ExtractFieldsToList(aSource: String; aList: TStringList);
    procedure SetSelectedResult;
    procedure SetDataSet(const AValue: TClientDataSet);
  public
    { Public declarations }
    property Connection: TADOConnection read FConnection write FConnection;
    property DataSet: TClientDataSet read FDataSet write SetDataSet;
    property SQL: TStringList read FSQL;
    property KeyFields: String read FKeyFields write SetKeyFields;
    property KeyValues: String read FKeyValues write SetKeyValues;
    property DisplayFields: String read FDisplayFields write SetDisplayFields;
    property ResultFields: String read FResultFields write SetResultFields;
    property SelectedValue: String read FSelectedValue;
    property SelectedDisplay: String read FSelectedDisplay;
    property QuotedString: String read FQuotedString write FQuotedString;
  end;

var
  frmMultiSelect2: TfrmMultiSelect2;

implementation

uses cbUtilis;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect2.FormCreate(Sender: TObject);
begin
  FDataReader := TADOQuery.Create( nil );
  FDataSet := TClientDataSet.Create( Self );
  FProvider := TDataSetProvider.Create( Self );
  {}
  FProvider.Name := 'FProvider';
  FProvider.Options := ( FProvider.Options + [poAllowCommandText] );
  FProvider.DataSet := FDataReader;
  {}
  FDataSet.ProviderName := 'FProvider';
  FSQL := TStringList.Create;
  FDisplayFields := EmptyStr;
  FResultFields := EmptyStr;
  FSelectedValue := EmptyStr;
  FSelectedDisplay := EmptyStr;
  FQuotedString := '''';
  FKeyFieldList := TStringList.Create;
  FKeyValueList := TStringList.Create;
  FDisplayFieldList := TStringList.Create;
  FDisplayCaptionList := TStringList.Create;
  FResultList := TStringList.Create;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect2.FormShow(Sender: TObject);
begin
  actConfirm.Enabled := False;
  actSelectAll.Enabled := False;
  actClear.Enabled := False;
  actCancel.Enabled := True;
  CreateDisplayColumns;
  Timer1.Enabled := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect2.FormDestroy(Sender: TObject);
begin
  ClearDataItem;
  if FDataReader.Active then FDataReader.Close;
  FDataReader.Connection := nil;
  FDataReader.Free;
  if ( FDataSet.Active ) then FDataSet.Close;
  FDataSet.Free;
  FProvider.Free;
  FSQL.Free;
  FKeyFieldList.Free;
  FKeyValueList.Free;
  FDisplayFieldList.Free;
  FDisplayCaptionList.Free;
  FResultList.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect2.Timer1Timer(Sender: TObject);
begin
  Timer1.Enabled := False;
  Screen.Cursor := crSQLWait;
  try
    if ( FSQL.Count > 0 ) then
    begin
      FDataReader.Connection := FConnection;
      FDataSet.CommandText := FSQL.Text;
      try
        FDataSet.Open;
      except
        on E: Exception do
        begin
          ErrorMsg( Format( '代碼複選資料顯示錯誤, 原因:%s', [E.Message] ) );
          Exit;
        end;
      end;
    end;
    if ( not FDataSet.Active ) or ( FDataSet.IsEmpty ) then
    begin
      WarningMsg( '無資料可顯示。' );
      Exit;
    end;
    AddDataItem;
    actConfirm.Enabled := True;
    actSelectAll.Enabled := True;
    actClear.Enabled := True;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect2.SetDataSet(const AValue: TClientDataSet);
begin
  FDataSet.Data := Null;
  if Assigned( AValue ) then
    FDataSet.Data := AValue.Data;
  if not VarIsNull( FDataSet.Data ) then
  begin
    FDataSet.Filter := AValue.Filter;
    FDataSet.FilterOptions := AValue.FilterOptions;
    FDataSet.Filtered := AValue.Filtered;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect2.SetKeyFields(const Value: String);
begin
  if ( FKeyFields <> Value ) then
  begin
    FKeyFields := Value;
    ExtractFieldsToList( FKeyFields, FKeyFieldList );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect2.SetKeyValues(const Value: String);
begin
  if ( FKeyValues <> Value ) then
  begin
    FKeyValues := StringReplace( Value, '''', EmptyStr, [rfReplaceAll] );
    ExtractFieldsToList( FKeyValues, FKeyValueList );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect2.SetDisplayFields(const Value: String);
var
  aIndex: Integer;
begin
  if ( FDisplayFields <> Value ) then
  begin
    FDisplayFields := Value;
    ExtractFieldsToList( FDisplayFields, FDisplayFieldList );
    for aIndex := FDisplayFieldList.Count - 1 downto 0 do
      if Odd( aIndex ) then FDisplayFieldList.Delete( aIndex );
    ExtractFieldsToList( FDisplayFields, FDisplayCaptionList );
    for aIndex := FDisplayCaptionList.Count - 1 downto 0 do
      if not Odd( aIndex ) then FDisplayCaptionList.Delete( aIndex );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect2.SetResultFields(const Value: String);
begin
  if ( FResultFields <> Value ) then
  begin
    FResultFields := Value;
    ExtractFieldsToList( FResultFields, FResultList );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect2.ExtractFieldsToList(aSource: String; aList: TStringList);
var
  aField: String;
begin
  aList.Clear;
  repeat
    aField := ExtractValue( aSource );
    if ( aField <> EmptyStr ) then aList.Add( aField ); 
  until ( aSource = EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect2.CreateDisplayColumns;
var
  aColumn: TListColumn;
  aIndex, aWidth: Integer;
begin
  ListView.Columns.Clear;
  aColumn := ListView.Columns.Add;
  aColumn.Caption := '選取';
  aColumn.Width := 50;
  aColumn.Alignment := taCenter;
  for aIndex := 0 to FDisplayFieldList.Count - 1 do
  begin
    aColumn := ListView.Columns.Add;
    aColumn.Caption := FDisplayCaptionList[aIndex];
  end;
  if ( ListView.Columns.Count - 1 ) > 0 then
  begin
    aWidth := ListView.Width div ( ListView.Columns.Count - 1 ) ;
    for aIndex := 1 to ListView.Columns.Count - 1 do
    begin
      ListView.Column[aIndex].Width := aWidth;
      if ( aIndex = 1 ) then
        ListView.Column[aIndex].Width := ( ListView.Column[aIndex].Width -
          ( aWidth div 3 ) );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect2.FreeupBookmark;
var
  aIndex: Integer;
begin
  for aIndex := 0 to ListView.Items.Count - 1 do
    if Assigned( ListView.Items[aIndex].Data ) then
      FDataSet.FreeBookmark( ListView.Items[aIndex].Data );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect2.AddDataItem;
var
  aItem: TListItem;
  aIndex, aFieldIndex, aLvIndex: Integer;
  aName: String;
begin
  FDataSet.First;
  while not FDataSet.Eof do
  begin
    aFieldIndex := 0;
    if ( ListView.Columns.Count - 1 ) <= 0 then Continue;
    aItem := ListView.Items.Add;
    for aIndex := 1 to ListView.Columns.Count - 1 do
    begin
      if aFieldIndex in [0..FDisplayFieldList.Count - 1] then
      begin
        aName := FDisplayFieldList[aFieldIndex];
        aItem.SubItems.Add( FDataSet.FieldByName( aName ).AsString );
        aItem.Data := FDataSet.GetBookmark;
        aItem.Checked := False;
        aItem.Selected := False;
      end;
      Inc( aFieldIndex );
    end;
    FDataSet.Next;
  end;
  for aIndex := 0 to FKeyValueList.Count - 1 do
  begin
    for aLvIndex := 0 to ListView.Items.Count - 1 do
    begin
      if ( ListView.Items[aLvIndex].SubItems[0] <> FKeyValueList[aIndex] ) then
        Continue;
      ListView.Items[aLvIndex].Checked := True;
      ListView.Items[aLvIndex].Selected := True;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect2.ClearDataItem;
begin
  FreeupBookmark;
  ListView.Clear;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect2.actSelectAllExecute(Sender: TObject);
var
  aIndex: Integer;
begin
  for aIndex := 0 to ListView.Items.Count - 1 do
    ListView.Items[aIndex].Checked := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect2.actClearExecute(Sender: TObject);
var
  aIndex: Integer;
begin
  for aIndex := 0 to ListView.Items.Count - 1 do
  begin
    ListView.Items[aIndex].Checked := False;
    ListView.Items[aIndex].Selected := False;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect2.actConfirmExecute(Sender: TObject);
begin
  SetSelectedResult;
  Self.ModalResult := mrOk;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect2.actCancelExecute(Sender: TObject);
begin
  Self.ModalResult := mrCancel;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect2.SetSelectedResult;
var
  aIndex, aIndex2: Integer;
  aValues, aDisplay: TStringList;
begin
  aValues := TStringList.Create;
  aDisplay := TStringList.Create;
  try
    aValues.Delimiter := ',';
    aDisplay.Delimiter := ',';
    for aIndex := 0 to ListView.Items.Count - 1 do
    begin
      if not ( ListView.Items[aIndex].Checked ) then Continue;
      if not Assigned( ListView.Items[aIndex].Data ) then Continue;
      FDataSet.GotoBookmark( ListView.Items[aIndex].Data );
      for aIndex2 := 0 to FKeyFieldList.Count - 1 do
        aValues.Add( Format( '%s%s%s', [FQuotedString,
          FDataSet.FieldByName( FKeyFieldList[aIndex2] ).AsString, FQuotedString] ) );
      for aIndex2 := 0 to FResultList.Count - 1 do
        aDisplay.Add( FDataSet.FieldByName( FResultList[aIndex2] ).AsString );
    end;
    FSelectedValue := aValues.DelimitedText;
    FSelectedDisplay := StringReplace( aDisplay.DelimitedText, '"', EmptyStr, [rfReplaceAll] );
  finally
    aValues.Free;
    aDisplay.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect2.edtTextPropertiesChange(Sender: TObject);
begin
  btnNext.Enabled := ( edtText.Text <> EmptyStr );
  btnPrior.Enabled := ( edtText.Text <> EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect2.chklShowSelectedPropertiesChange(Sender: TObject);
var
  aCounter: Integer;
begin
  if chklShowSelected.Checked then
  begin
    aCounter := ListView.Items.Count;
    while aCounter > 0 do
    begin
      if not ListView.Items[aCounter-1].Checked then
      begin
        if Assigned( ListView.Items[aCounter-1].Data ) then
          FDataReader.FreeBookmark( ListView.Items[aCounter-1].Data );
        ListView.Items[aCounter-1].Delete;
      end;
      Dec (aCounter);
    end;
  end else
  begin
    ClearDataItem;
    Timer1.Enabled := True;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect2.ListViewDblClick(Sender: TObject);
var
  aItem: TListItem;
begin
  aItem := ListView.Selected;
  if Assigned( aItem ) then
    aItem.Checked := not aItem.Checked;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect2.btnNextClick(Sender: TObject);
var
  aIndex, aSearchCol: Integer;
begin
  aSearchCol := 1;
  if ( rdbCodeNo.Checked ) or ( ListView.Columns.Count <= 2 )  then aSearchCol := 0;
  aIndex := ListView.ItemIndex + 1;
  if ( aIndex > ListView.Items.Count - 1 )then aIndex := 0;
  while ( aIndex <= ListView.Items.Count - 1 ) and  ( ListView.Items.Count > 0 ) do
  begin
    if Pos( edtText.Text, ListView.Items[aIndex].SubItems[aSearchCol] ) > 0 then
    begin
      ListView.ItemIndex := aIndex;
      ListView.Items[aIndex].MakeVisible(False);
      aIndex := ListView.Items.Count;
    end;
    Inc(aIndex);
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect2.btnPriorClick(Sender: TObject);
var
  aIndex, aSearchCol: Integer;
begin
  aSearchCol := 1;
  if ( rdbCodeNo.Checked ) or ( ListView.Columns.Count <= 2 ) then aSearchCol := 0;
  aIndex := ListView.ItemIndex - 1;
  if ( aIndex < 0 ) then aIndex := 0;
  while ( aIndex >= 0 ) and ( ListView.Items.Count > 0 ) do
  begin
    if Pos( edtText.Text,ListView.Items[aIndex].SubItems[aSearchCol] ) > 0 then
    begin
      ListView.ItemIndex := aIndex;
      ListView.Items[aIndex].MakeVisible(False);
      aIndex := 0;
    end;
    Dec( aIndex );
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
