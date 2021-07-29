unit frmMultiSelectU2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, StdCtrls, ComCtrls, ExtCtrls, ActnList;

type
  TfrmMultiSelect1 = class(TForm)
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
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure ListViewMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure actSelectAllExecute(Sender: TObject);
    procedure actConfirmExecute(Sender: TObject);
    procedure ListViewSelectItem(Sender: TObject; Item: TListItem;
      Selected: Boolean);
    procedure actClearExecute(Sender: TObject);
    procedure actCancelExecute(Sender: TObject);
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
  public
    { Public declarations }
    property Connection: TADOConnection read FConnection write FConnection;
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
  frmMultiSelect1: TfrmMultiSelect1;

implementation

uses cbUtilis;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect1.FormCreate(Sender: TObject);
begin
  FDataReader := TADOQuery.Create( nil );
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

procedure TfrmMultiSelect1.FormShow(Sender: TObject);
begin
  actConfirm.Enabled := False;
  actSelectAll.Enabled := False;
  actClear.Enabled := False;
  actCancel.Enabled := True;
  CreateDisplayColumns;
  Timer1.Enabled := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect1.FormDestroy(Sender: TObject);
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

procedure TfrmMultiSelect1.Timer1Timer(Sender: TObject);
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
        ErrorMsg( Format( '代碼複選資料顯示錯誤, 原因:%s', [E.Message] ) );
    end;
  finally
    Screen.Cursor := crDefault;
  end;  
end;

{ ---------------------------------------------------------------------------- }


procedure TfrmMultiSelect1.SetKeyFields(const Value: String);
begin
  if ( FKeyFields <> Value ) then
  begin
    FKeyFields := Value;
    ExtractFieldsToList( FKeyFields, FKeyFieldList );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect1.SetKeyValues(const Value: String);
begin
  if ( FKeyValues <> Value ) then
  begin
    FKeyValues := StringReplace( Value, '''', EmptyStr, [rfReplaceAll] );
    ExtractFieldsToList( FKeyValues, FKeyValueList );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect1.SetDisplayFields(const Value: String);
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

procedure TfrmMultiSelect1.SetResultFields(const Value: String);
begin
  if ( FResultFields <> Value ) then
  begin
    FResultFields := Value;
    ExtractFieldsToList( FResultFields, FResultList );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect1.ExtractFieldsToList(aSource: String; aList: TStringList);
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

procedure TfrmMultiSelect1.CreateDisplayColumns;
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

procedure TfrmMultiSelect1.FreeupBookmark;
var
  aIndex: Integer;
begin
  for aIndex := 0 to ListView.Items.Count - 1 do
    if Assigned( ListView.Items[aIndex].Data ) then
      FDataReader.FreeBookmark( ListView.Items[aIndex].Data );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect1.AddDataItem;
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
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect1.ClearDataItem;
begin
  FreeupBookmark;
  ListView.Clear;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect1.ListViewMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  aItem: TListItem;
begin
 if htOnStateIcon in ListView.GetHitTestInfoAt( X, Y ) then
 begin
   aItem := ListView.GetItemAt( X, Y );
   if Assigned( aItem ) then
     aItem.Selected := aItem.Checked;
 end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect1.ListViewSelectItem(Sender: TObject;
  Item: TListItem; Selected: Boolean);
begin
 Item.Selected := Item.Checked;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect1.actSelectAllExecute(Sender: TObject);
var
  aIndex: Integer;
begin
  for aIndex := 0 to ListView.Items.Count - 1 do
  begin
    ListView.Items[aIndex].Checked := True;
    ListView.Items[aIndex].Selected := True;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect1.actClearExecute(Sender: TObject);
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

procedure TfrmMultiSelect1.actConfirmExecute(Sender: TObject);
begin
  SetSelectedResult;
  Self.ModalResult := mrOk;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect1.actCancelExecute(Sender: TObject);
begin
  Self.ModalResult := mrCancel;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect1.SetSelectedResult;
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
    FSelectedDisplay := StringReplace( aDisplay.DelimitedText, '"', EmptyStr, [rfReplaceAll] );
  finally
    aValues.Free;
    aDisplay.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
