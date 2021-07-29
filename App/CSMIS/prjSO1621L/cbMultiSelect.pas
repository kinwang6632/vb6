unit cbMultiSelect;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, Menus, StdCtrls, ComCtrls, ExtCtrls, ActnList,
  cxLookAndFeelPainters, cxButtons, cxCheckBox, cxControls, cxContainer,
  cxEdit, cxTextEdit, cxRadioGroup, cxGroupBox, dxSkinsCore, dxSkinsDefaultPainters;

type
  TfmMultiSelect = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    ScrollBox1: TScrollBox;
    ListView: TListView;
    Button1: TcxButton;
    Button2: TcxButton;
    Button3: TcxButton;
    Button4: TcxButton;
    Timer1: TTimer;
    ActionList1: TActionList;
    actConfirm: TAction;
    actSelectAll: TAction;
    actClear: TAction;
    actCancel: TAction;
    GroupBox1: TcxGroupBox;
    rdbCodeNo: TcxRadioButton;
    rdbDescription: TcxRadioButton;
    edtText: TcxTextEdit;
    chklShowSelected: TcxCheckBox;
    btnNext: TcxButton;
    btnPrior: TcxButton;
    Bevel1: TBevel;
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
    FSQL: TStringList;
    FKeyFields: String;
    FKeyValues: String;
    FDisplayFields: String;
    FResultFields: String;
    FSelectedValue: String;
    FSelectedText: String;
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
  public
    { Public declarations }
    property Connection: TADOConnection read FConnection write FConnection;
    property SQL: TStringList read FSQL;
    property KeyFields: String read FKeyFields write SetKeyFields;
    property KeyValues: String read FKeyValues write SetKeyValues;
    property DisplayFields: String read FDisplayFields write SetDisplayFields;
    property ResultFields: String read FResultFields write SetResultFields;
    property SelectedValue: String read FSelectedValue;
    property SelectedText: String read FSelectedText;
    property QuotedString: String read FQuotedString write FQuotedString;
  end;

var
  fmMultiSelect: TfmMultiSelect;

implementation

uses cbUtilis;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TfmMultiSelect.FormCreate(Sender: TObject);
begin
  FDataReader := TADOQuery.Create( nil );
  FSQL := TStringList.Create;
  FDisplayFields := EmptyStr;
  FResultFields := EmptyStr;
  FSelectedValue := EmptyStr;
  FSelectedText := EmptyStr;
  FQuotedString := '''';
  FKeyFieldList := TStringList.Create;
  FKeyValueList := TStringList.Create;
  FDisplayFieldList := TStringList.Create;
  FDisplayCaptionList := TStringList.Create;
  FResultList := TStringList.Create;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMultiSelect.FormShow(Sender: TObject);
begin
  actConfirm.Enabled := False;
  actSelectAll.Enabled := False;
  actClear.Enabled := False;
  actCancel.Enabled := True;
  CreateDisplayColumns;
  Timer1.Enabled := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMultiSelect.FormDestroy(Sender: TObject);
begin
  ClearDataItem;
  if FDataReader.Active then FDataReader.Close;
  FDataReader.Connection := nil;
  FDataReader.Free;
  FSQL.Free;
  FKeyFieldList.Free;
  FKeyValueList.Free;
  FDisplayFieldList.Free;
  FDisplayCaptionList.Free;
  FResultList.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMultiSelect.Timer1Timer(Sender: TObject);
begin
  Timer1.Enabled := False;
  Screen.Cursor := crSQLWait;
  try
    FDataReader.Connection := FConnection;
    FDataReader.SQL.Text := FSQL.Text;
    try
      FDataReader.Open;
      if FDataReader.IsEmpty then
      begin
        WarningMsg( '無資料可顯示。' );
        Exit;
      end;
      AddDataItem;
      actConfirm.Enabled := True;
      actSelectAll.Enabled := True;
      actClear.Enabled := True;
    except
      on E: Exception do
        ErrorMsg( Format( '複選資料顯示錯誤, 原因:%s', [E.Message] ) );
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }


procedure TfmMultiSelect.SetKeyFields(const Value: String);
begin
  if ( FKeyFields <> Value ) then
  begin
    FKeyFields := Value;
    ExtractFieldsToList( FKeyFields, FKeyFieldList );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMultiSelect.SetKeyValues(const Value: String);
begin
  if ( FKeyValues <> Value ) then
  begin
    FKeyValues := StringReplace( Value, '''', EmptyStr, [rfReplaceAll] );
    ExtractFieldsToList( FKeyValues, FKeyValueList );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMultiSelect.SetDisplayFields(const Value: String);
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

procedure TfmMultiSelect.SetResultFields(const Value: String);
begin
  if ( FResultFields <> Value ) then
  begin
    FResultFields := Value;
    ExtractFieldsToList( FResultFields, FResultList );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMultiSelect.ExtractFieldsToList(aSource: String; aList: TStringList);
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

procedure TfmMultiSelect.CreateDisplayColumns;
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

procedure TfmMultiSelect.FreeupBookmark;
var
  aIndex: Integer;
begin
  for aIndex := 0 to ListView.Items.Count - 1 do
    if Assigned( ListView.Items[aIndex].Data ) then
      FDataReader.FreeBookmark( ListView.Items[aIndex].Data );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMultiSelect.AddDataItem;
var
  aItem: TListItem;
  aIndex, aFieldIndex, aLvIndex: Integer;
  aName: String;
begin
  FDataReader.First;
  while not FDataReader.Eof do
  begin
    aFieldIndex := 0;
    if ( ListView.Columns.Count - 1 ) <= 0 then Continue;
    aItem := ListView.Items.Add;
    for aIndex := 1 to ListView.Columns.Count - 1 do
    begin
      if aFieldIndex in [0..FDisplayFieldList.Count - 1] then
      begin
        aName := FDisplayFieldList[aFieldIndex];
        aItem.SubItems.Add( FDataReader.FieldByName( aName ).AsString );
        aItem.Data := FDataReader.GetBookmark;
        aItem.Checked := False;
        aItem.Selected := False;
      end;
      Inc( aFieldIndex );
    end;
    FDataReader.Next;
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
  if ( ListView.Items.Count > 0 ) then ListView.ItemIndex := 0;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMultiSelect.ClearDataItem;
begin
  FreeupBookmark;
  ListView.Clear;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMultiSelect.actSelectAllExecute(Sender: TObject);
var
  aIndex: Integer;
begin
  for aIndex := 0 to ListView.Items.Count - 1 do
    ListView.Items[aIndex].Checked := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMultiSelect.actClearExecute(Sender: TObject);
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

procedure TfmMultiSelect.actConfirmExecute(Sender: TObject);
begin
  SetSelectedResult;
  Self.ModalResult := mrOk;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMultiSelect.actCancelExecute(Sender: TObject);
begin
  Self.ModalResult := mrCancel;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMultiSelect.SetSelectedResult;
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
      FDataReader.GotoBookmark( ListView.Items[aIndex].Data );
      for aIndex2 := 0 to FKeyFieldList.Count - 1 do
        aValues.Add( Format( '%s%s%s', [FQuotedString,
          FDataReader.FieldByName( FKeyFieldList[aIndex2] ).AsString, FQuotedString] ) );
      for aIndex2 := 0 to FResultList.Count - 1 do
        aDisplay.Add( FDataReader.FieldByName( FResultList[aIndex2] ).AsString );
    end;
    FSelectedValue := aValues.DelimitedText;
    FSelectedText := StringReplace( aDisplay.DelimitedText, '"', EmptyStr, [rfReplaceAll] );
  finally
    aValues.Free;
    aDisplay.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMultiSelect.edtTextPropertiesChange(Sender: TObject);
begin
  btnNext.Enabled := ( edtText.Text <> EmptyStr );
  btnPrior.Enabled := ( edtText.Text <> EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMultiSelect.chklShowSelectedPropertiesChange(
  Sender: TObject);
var
  aCounter: Integer;
begin
  if chklShowSelected.Checked then
  begin
    aCounter := ListView.Items.Count;
    while ( aCounter > 0 ) do
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

procedure TfmMultiSelect.ListViewDblClick(Sender: TObject);
var
  aItem: TListItem;
begin
  aItem := ListView.Selected;
  if Assigned( aItem ) then
    aItem.Checked := not aItem.Checked;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMultiSelect.btnNextClick(Sender: TObject);
var
  aIndex, aSearchCol: Integer;
begin
  aSearchCol := 1;
  if rdbCodeNo.Checked then aSearchCol := 0;
  aIndex := ListView.ItemIndex + 1;
  if ( aIndex > ListView.Items.Count - 1 ) then aIndex := 0;
  while ( aIndex <= ListView.Items.Count - 1 ) do
  begin
    if Pos( edtText.Text,ListView.Items[aIndex].SubItems[aSearchCol] ) > 0 then
    begin
      ListView.ItemIndex := aIndex;
      ListView.Items[aIndex].MakeVisible( False );
      aIndex := ListView.Items.Count;
    end;
    Inc(aIndex);
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMultiSelect.btnPriorClick(Sender: TObject);
var
  aIndex, aSearchCol: Integer;
begin
  aSearchCol := 1;
  if rdbCodeNo.Checked then aSearchCol := 0;
  aIndex := ListView.ItemIndex - 1;
  if ( aIndex < 0 ) then aIndex := 0;
  while ( aIndex >= 0 ) do
  begin
    if Pos(edtText.Text,ListView.Items[aIndex].SubItems[aSearchCol]) > 0 then
    begin
      ListView.ItemIndex := aIndex;
      ListView.Items[aIndex].MakeVisible( False );
      aIndex := 0;
    end;
    Dec(aIndex);
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
