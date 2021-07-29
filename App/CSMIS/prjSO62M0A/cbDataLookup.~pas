unit cbDataLookup;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, DBClient,
  cbStyleController, cbDBController,
  cxControls, cxContainer, cxGraphics,
  cxMaskEdit, cxDropDownEdit, cxLookupEdit, cxDBLookupEdit, cxLookupDBGrid,
  cxDBLookupComboBox,cxEdit, cxTextEdit,
  dxSkinsCore, dxSkinsDefaultPainters, dxSkinBlack, dxSkinBlue,
  dxSkinCaramel, dxSkinCoffee;


type
  TDataLookup = class(TFrame)
    CodeNo: TcxTextEdit;
    CodeName: TcxLookupComboBox;
    procedure CodeNamePropertiesChange(Sender: TObject);
    procedure CodeNoPropertiesChange(Sender: TObject);
    procedure CodeNoExit(Sender: TObject);
    procedure CodeNamePropertiesInitPopup(Sender: TObject);
  private
    { Private declarations }
    FDataReader: TADOQuery;
    FDataSource: TDataSource;
    FSQL: TStringList;
    FShowWarning: Boolean;
    FDataSet: TClientDataSet;
    FKeyFieldName: String;
    FLockCount: Integer;
    FServicetype: String;
    FRefNo: String;
    function GetValue: String;
    function GetValueName: String;
    procedure SetValue(const Value: String);
    procedure SetValueName(const Value: String);
    procedure InternalInitializa;
    procedure LoadOtherFieldValue;
  protected
    procedure SetEnabled(Value: Boolean); override;
  public
    { Public declarations }
    property DataSet: TClientDataSet read FDataSet write FDataSet;
    property SQL: TStringList read FSQL;
    property Value: String read GetValue write SetValue;
    property ValueName: String read GetValueName write SetValueName;
    property ShowWarningWhenNotFound: Boolean read FShowWarning write FShowWarning;
    property ServiceType: String read FServiceType;
    property RefNo: String read FRefNo;
    { Must Call This First, Because TFrame haven't OnCreate Event }
    procedure Initializa;
    procedure Finaliza;
    procedure Execute;
    procedure Clear;
    procedure LockChangeEvent;
    procedure UnlockChangeEvent;
  end;


implementation

uses cbUtilis;

{$R *.dfm}


{ TDataLookup }

{ ---------------------------------------------------------------------------- }

procedure TDataLookup.InternalInitializa;
var
  aColumn: TcxLookupDBGridColumn;
begin
  CodeName.Properties.ListSource := FDataSource;
  CodeName.Properties.ListOptions.ShowHeader := False;
  if ( CodeName.Properties.ListFieldNames = EmptyStr ) then
  begin
    aColumn := CodeName.Properties.ListColumns.Add;
    aColumn.FieldName := 'CODENO';
    aColumn := CodeName.Properties.ListColumns.Add;
    aColumn.FieldName := 'DESCRIPTION';
    CodeName.Properties.KeyFieldNames := 'CODENO';
    CodeName.Properties.ListFieldIndex := 1;
    CodeName.Properties.ListFieldNames := 'CODENO;DESCRIPTION';
  end;
  FKeyFieldName := CodeName.Properties.KeyFieldNames;
  CodeName.Properties.DropDownRows := 15;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataLookup.LoadOtherFieldValue;
begin
  if ( Value = EmptyStr) or ( ValueName = EmptyStr ) then
  begin
    FServicetype := EmptyStr;
    FRefNo := EmptyStr;
  end else
  begin
    if FDataSource.DataSet.Locate( FKeyFieldName, Value, [] ) then
    begin
      if Assigned( FDataSource.DataSet.FindField( 'SERVICETYPE' ) ) then
        FServicetype := FDataSource.DataSet.FieldByName( 'SERVICETYPE' ).AsString;
      if Assigned( FDataSource.DataSet.FindField( 'REFNO' ) ) then
        FRefNo := FDataSource.DataSet.FieldByName( 'REFNO' ).AsString;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataLookup.Initializa;
begin
  FDataReader := TADOQuery.Create( nil );
  FDataReader.Connection := DBController.DataConnection;
  FDataSource := TDataSource.Create( nil );
  FDataSource.DataSet := FDataReader;
  FSQL := TStringList.Create;
  CodeNo.Text := EmptyStr;
  CodeName.Text := EmptyStr;
  InternalInitializa;
  FShowWarning := True;
  FLockCount := 0;
  FServicetype := EmptyStr;
  FRefNo := EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataLookup.Finaliza;
begin
  FDataSet := nil;
  FSQL.Free;
  FreeAndNil( FDataSource );
  FDataReader.Close;
  FDataReader.Connection := nil;
  FreeAndNil( FDataReader );
end;

{ ---------------------------------------------------------------------------- }

procedure TDataLookup.Execute;
begin
  CodeName.LockChangeEvents( True, False );
  try
    if ( Assigned( FDataSet ) ) then
      FDataSource.DataSet := FDataSet
    else begin
      FDataSource.DataSet := FDataReader;
      FDataReader.Close;
      FDataReader.SQL.Text := FSQL.Text;
      FDataReader.Open;
    end;
    FDataSource.DataSet.FieldByName( FKeyFieldName ).Alignment := taLeftJustify;
    FDataSource.DataSet.FieldByName( FKeyFieldName ).DisplayWidth := 12;
  finally
    CodeName.LockChangeEvents( False, False );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataLookup.Clear;
begin
  CodeNo.Text := EmptyStr;
  CodeName.Text := EmptyStr;
  FServicetype := EmptyStr;
  FRefNo := EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

function TDataLookup.GetValue: String;
begin
  Result := CodeNo.Text;
end;

{ ---------------------------------------------------------------------------- }

function TDataLookup.GetValueName: String;
begin
  Result := CodeName.Text;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataLookup.SetValue(const Value: String);
begin
  CodeNo.Text := Value;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataLookup.SetValueName(const Value: String);
begin
  CodeName.Text := Value;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataLookup.SetEnabled(Value: Boolean);
begin
  inherited SetEnabled( Value );
  CodeNo.Enabled := Value;
  CodeName.Enabled := Value;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataLookup.CodeNamePropertiesChange(Sender: TObject);
begin
  CodeNo.LockChangeEvents( True, False );
  try
    if ( CodeName.EditingText = EmptyStr ) then
      CodeNo.EditValue := Null
    else
      CodeNo.EditValue := CodeName.EditValue;
  finally
    CodeNo.LockChangeEvents( False, False );
  end;
  LoadOtherFieldValue;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataLookup.CodeNoPropertiesChange(Sender: TObject);
begin
  CodeName.Properties.OnInitPopup( CodeName );
  CodeName.LockChangeEvents( True, False );
  try
    if CodeNo.EditingText = EmptyStr then
      CodeName.EditValue := Null
    else
      CodeName.EditValue := CodeNo.EditingText;
  finally
    CodeName.LockChangeEvents( False, False );
  end;
  LoadOtherFieldValue;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataLookup.CodeNoExit(Sender: TObject);
begin
  if ( FShowWarning ) and
     ( CodeNo.Text <> EmptyStr ) and
     ( CodeName.Text = EmptyStr ) then
  begin
    CodeNo.Text := EmptyStr;
    FServicetype := EmptyStr;
    FRefNo := EmptyStr;
    WarningMsg( '查應無對應資料。' );
    if CodeNo.CanFocusEx then CodeNo.SetFocus;
    Abort;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataLookup.CodeNamePropertiesInitPopup(Sender: TObject);
begin
  Self.Execute;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataLookup.LockChangeEvent;
begin
  Inc( FLockCount );
  if ( FLockCount > 0 ) then
  begin
    CodeNo.LockChangeEvents( True, False );
    CodeName.LockChangeEvents( True, False );
  end;    
end;

{ ---------------------------------------------------------------------------- }

procedure TDataLookup.UnlockChangeEvent;
begin
  Dec( FLockCount );
  if ( FLockCount <= 0 ) then
  begin
    CodeNo.LockChangeEvents( False, False );
    CodeName.LockChangeEvents( False, False );
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
