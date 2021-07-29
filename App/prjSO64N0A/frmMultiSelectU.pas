unit frmMultiSelectU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, StdCtrls, ComCtrls, ExtCtrls, DB, DBClient, ADODB, Provider, ActnList,
  cxLookAndFeelPainters, cxButtons, cxCheckBox, cxControls, cxContainer,
  cxEdit, cxTextEdit, cxRadioGroup, dxSkinsCore, dxSkinsDefaultPainters,
  dxSkinBlack, dxSkinBlue, dxSkinCaramel, dxSkinCoffee;

type
  TfrmMultiSelect = class(TForm)
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
    procedure ListViewChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
  private
    { Private declarations }
    FConnection: TADOConnection;
    FDataReader: TADOQuery;
    FProvider: TDataSetProvider;
    FDataSet: TClientDataSet;
    FSQL: TStringList;
    FKeyFields: String;
    FKeyValues: String;
    FKeyValuesOrd: String;
    FDisplayFields: String;
    FResultFields: String;
    FSelectedValue: String;
    FSelectedOrd:String;
    FSelectedDisplay: String;
    FQuotedString: String;
    FKeyFieldList: TStringList;
    FKeyValueList: TStringList;
    FKeyValueOrdList: TStringList;
    FDisplayFieldList: TStringList;
    FDisplayCaptionList: TStringList;
    FResultList: TStringList;
    FLoadComplete:Boolean;
    FUseOrd : Boolean;
    procedure SetKeyFields(const Value: String);
    procedure SetKeyValues(const Value: String);
    procedure SetKeyValuesOrd(const Value:string);
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
    property KeyValueOrd:String read FKeyValuesOrd write SetKeyValuesOrd;
    property DisplayFields: String read FDisplayFields write SetDisplayFields;
    property ResultFields: String read FResultFields write SetResultFields;
    property SelectedValue: String read FSelectedValue;
    property SelectedDisplay: String read FSelectedDisplay;
    property SelectOrd:String read FSelectedOrd;
    property QuotedString: String read FQuotedString write FQuotedString;
    property IsUseOrd : Boolean read FUseOrd write FUseOrd;
  end;

var
  frmMultiSelect: TfrmMultiSelect;

implementation

uses cbUtilis;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect.FormCreate(Sender: TObject);
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
  FSelectedOrd := EmptyStr;
  FQuotedString := '''';
  FKeyFieldList := TStringList.Create;
  FKeyValueList := TStringList.Create;
  FKeyValueOrdList := TStringList.Create;
  FDisplayFieldList := TStringList.Create;
  FDisplayCaptionList := TStringList.Create;
  FResultList := TStringList.Create;
  FLoadComplete := False;

end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect.FormShow(Sender: TObject);
begin
  actConfirm.Enabled := False;
  actSelectAll.Enabled := False;
  actClear.Enabled := False;
  actCancel.Enabled := True;
  CreateDisplayColumns;
  Timer1.Enabled := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect.FormDestroy(Sender: TObject);
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
  FKeyValueOrdList.Free;
  FDisplayFieldList.Free;
  FDisplayCaptionList.Free;
  FResultList.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect.Timer1Timer(Sender: TObject);
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
    FLoadComplete := True;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect.SetDataSet(const AValue: TClientDataSet);
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

procedure TfrmMultiSelect.SetKeyFields(const Value: String);
begin
  if ( FKeyFields <> Value ) then
  begin
    FKeyFields := Value;
    ExtractFieldsToList( FKeyFields, FKeyFieldList );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect.SetKeyValues(const Value: String);
begin
  if ( FKeyValues <> Value ) then
  begin
    FKeyValues := StringReplace( Value, '''', EmptyStr, [rfReplaceAll] );
    ExtractFieldsToList( FKeyValues, FKeyValueList );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect.SetDisplayFields(const Value: String);
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

procedure TfrmMultiSelect.SetResultFields(const Value: String);
begin
  if ( FResultFields <> Value ) then
  begin
    FResultFields := Value;
    ExtractFieldsToList( FResultFields, FResultList );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect.ExtractFieldsToList(aSource: String; aList: TStringList);
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

procedure TfrmMultiSelect.CreateDisplayColumns;
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
  if FUseOrd then
  begin
    aColumn := ListView.Columns.Add;
    aColumn.Width :=50;
    aColumn.Caption := '順序';
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

      if FUseOrd then
      begin
        if ( aIndex = 2 ) then
        begin
           ListView.Column[aIndex].Width := ( ListView.Column[aIndex].Width * 2)-70;
        end;
      end;

      if aIndex = 3 then
        ListView.Column[aIndex].Width := 50;

    end;
  end;

end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect.FreeupBookmark;
var
  aIndex: Integer;
begin
  for aIndex := 0 to ListView.Items.Count - 1 do
    if Assigned( ListView.Items[aIndex].Data ) then
      FDataSet.FreeBookmark( ListView.Items[aIndex].Data );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect.AddDataItem;
var
  aItem: TListItem;
  aIndex, aFieldIndex, aLvIndex: Integer;
  aName: String;
begin
  FDataSet.First;
  FLoadComplete := False;
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
    aitem.SubItems.Add('');
    FDataSet.Next;
  end;
  FLoadComplete := True;
  for aIndex := 0 to FKeyValueList.Count - 1 do
  begin
    for aLvIndex := 0 to ListView.Items.Count - 1 do
    begin
      if ( ListView.Items[aLvIndex].SubItems[0] <> FKeyValueList[aIndex] ) then
        Continue;
      ListView.Items[aLvIndex].Checked := True;
      ListView.Items[aLvIndex].Selected := True;
      if ( FKeyValueOrdList.Count > 0 ) and ( FKeyValueOrdList.Count > aIndex ) then
      begin
        ListView.Items[aLvIndex].SubItems[2]:= FKeyValueOrdList[aIndex];
      end;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect.ClearDataItem;
begin
  FreeupBookmark;
  ListView.Clear;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect.actSelectAllExecute(Sender: TObject);
var
  aIndex: Integer;
begin
  for aIndex := 0 to ListView.Items.Count - 1 do
    ListView.Items[aIndex].Checked := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect.actClearExecute(Sender: TObject);
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

procedure TfrmMultiSelect.actConfirmExecute(Sender: TObject);
begin
  SetSelectedResult;
  Self.ModalResult := mrOk;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect.actCancelExecute(Sender: TObject);
begin
  Self.ModalResult := mrCancel;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect.SetSelectedResult;
var
  aIndex, aIndex2: Integer;
  aValues, aDisplay,aOrd: TStringList;
begin
  aValues := TStringList.Create;
  aDisplay := TStringList.Create;
  aOrd := TStringList.Create;
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
        if FUseOrd then
        begin
          aOrd.Add(ListView.Items[aIndex].SubItems.Strings[2]);
        end;

      for aIndex2 := 0 to FResultList.Count - 1 do
        aDisplay.Add( FDataSet.FieldByName( FResultList[aIndex2] ).AsString );
    end;
    FSelectedValue := aValues.DelimitedText;
    FSelectedDisplay := StringReplace( aDisplay.DelimitedText, '"', EmptyStr, [rfReplaceAll] );
    FSelectedOrd := aOrd.DelimitedText;
  finally
    aValues.Free;
    aDisplay.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect.edtTextPropertiesChange(Sender: TObject);
begin
  btnNext.Enabled := ( edtText.Text <> EmptyStr );
  btnPrior.Enabled := ( edtText.Text <> EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect.chklShowSelectedPropertiesChange(Sender: TObject);
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

procedure TfrmMultiSelect.ListViewDblClick(Sender: TObject);
var
  aItem: TListItem;
begin
  aItem := ListView.Selected;
  if Assigned( aItem ) then
    aItem.Checked := not aItem.Checked;


  {
  begin
    for aLvIndex := 0 to ListView.Items.Count - 1 do
    begin
      if ( ListView.Items[aLvIndex].SubItems[0] <> FKeyValueList[aIndex] ) then
        Continue;
      ListView.Items[aLvIndex].Checked := True;
      ListView.Items[aLvIndex].Selected := True;
    end;
  end;
  }
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMultiSelect.btnNextClick(Sender: TObject);
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

procedure TfrmMultiSelect.btnPriorClick(Sender: TObject);
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

procedure TfrmMultiSelect.ListViewChange(Sender: TObject; Item: TListItem;
  Change: TItemChange);
  var i,min,max:Integer;
begin
    min:=0;
    max:=0;

    if FUseOrd then
    begin
      if FLoadComplete then
      begin

        if item.Checked  then
        begin
           for i := 0 to ListView.items.Count-1 do
           begin
              if ListView.Items[i].SubItems.Strings[2] <> '' then
              begin
                if min < StrToInt(ListView.Items[i].SubItems.Strings[2]) then
                begin
                  min := StrToInt(ListView.Items[i].SubItems.Strings[2]);
                end;
              end;
           end;
           if ( item.SubItems.Strings[2] = '' )  then
             item.SubItems.Strings[2] := IntToStr(min+1);
         end else
         begin
          if item.SubItems.Strings[2] <> '' then
          begin
            max := StrToInt(item.SubItems.Strings[2]);


            for i := 0 to ListView.items.Count-1 do
            begin
              if ListView.Items[i].SubItems.Strings[2] <> '' then
              begin
                if max < StrToInt(ListView.Items[i].SubItems.Strings[2]) then
                begin
                  ListView.Items[i].SubItems.Strings[2] := IntToStr(
                    StrToInt(ListView.Items[i].SubItems.Strings[2])-1 );
                end;
              end;
             end;
            item.SubItems.Strings[2] := '';
          end;
         end;
      end;
    end;

end;

procedure TfrmMultiSelect.SetKeyValuesOrd(const Value: string);
begin
  if ( FKeyValuesOrd  <> Value ) then
  begin
    FKeyValuesOrd := StringReplace( Value, '''', EmptyStr, [rfReplaceAll] );
    ExtractFieldsToList( FKeyValuesOrd, FKeyValueOrdList);

  end;
end;

end.
