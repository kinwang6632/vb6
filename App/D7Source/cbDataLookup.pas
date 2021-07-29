unit cbDataLookup;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, DBClient, 
  cxMaskEdit, cxDropDownEdit, cxLookupEdit, cxDBLookupEdit, cxLookupDBGrid,
  cxDBLookupComboBox, cxControls, cxContainer, cxEdit, cxTextEdit,
  cxGraphics, dxSkinsCore, dxSkinsDefaultPainters, dxSkinBlack, dxSkinBlue,
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
    FConnection: TADOConnection;
    FCodeNoFieldName: String;
    FDescriptionFieldName: String;
    FOnSelectChanged : TNotifyEvent;
    function GetValue: String;
    function GetValueName: String;
    procedure SetValue(const Value: String);
    procedure SetValueName(const Value: String);
  protected
    procedure SetEnabled(Value: Boolean); override;
  public
    { Public declarations }
    property DataSet: TClientDataSet read FDataSet write FDataSet;
    property SQL: TStringList read FSQL;
    property Value: String read GetValue write SetValue;
    property ValueName: String read GetValueName write SetValueName;
    property ShowWarningWhenNotFound: Boolean read FShowWarning write FShowWarning;
    property Connection: TADOConnection read FConnection write FConnection;
    property CodeNoFieldName: String read FCodeNoFieldName write FCodeNoFieldName;
    property DescriptionFieldName: String read FDescriptionFieldName write FDescriptionFieldName;
    function FieldByName(const AFieldName: string): TField;
    { Must Call This First, Because TFrame haven't OnCreate Event }
    procedure Initializa;
    procedure Finaliza;
    procedure Execute;
    procedure Clear;
  published
    property OnSelectChange: TNotifyEvent read FOnSelectChanged write FOnSelectChanged;

  end;


implementation

uses cbUtilis;

{$R *.dfm}


{ TDataLookup }

{ ---------------------------------------------------------------------------- }

procedure TDataLookup.Initializa;
var
  aColumn: TcxLookupDBGridColumn;
begin
  FDataReader := TADOQuery.Create( nil );
  FDataReader.Connection := FConnection;
  FDataSource := TDataSource.Create( nil );
  FDataSource.DataSet := FDataReader;
  FSQL := TStringList.Create;
  CodeName.Properties.ListSource := FDataSource;
  CodeName.Properties.ListOptions.ShowHeader := False;
  CodeName.Properties.ListColumns.Clear;
  aColumn := CodeName.Properties.ListColumns.Add;
  aColumn.FieldName := FCodeNoFieldName;
  aColumn := CodeName.Properties.ListColumns.Add;
  aColumn.FieldName := FDescriptionFieldName;
  CodeName.Properties.KeyFieldNames := FCodeNoFieldName;
  CodeName.Properties.ListFieldIndex := 1;
  CodeName.Properties.ListFieldNames := Format( '%s;%s', [FCodeNoFieldName,FDescriptionFieldName]);
  CodeName.Properties.DropDownRows := 15;
  CodeNo.Text := EmptyStr;
  CodeName.Text := EmptyStr;
  FShowWarning := True;
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
    begin
      FDataSource.DataSet := FDataSet;
    end else
    begin
      FDataSource.DataSet := FDataReader;
      FDataReader.Close;
      FDataReader.SQL.Text := FSQL.Text;
      FDataReader.Open;
      FDataReader.FieldByName( FCodeNoFieldName ).Alignment := taLeftJustify;
      FDataReader.FieldByName( FCodeNoFieldName ).DisplayWidth := 15;
    end;
  finally
    CodeName.LockChangeEvents( False, False );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataLookup.Clear;
begin
  CodeNo.Text := EmptyStr;
  CodeName.Text := EmptyStr;
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
    if Assigned(FOnSelectChanged) then
      FOnSelectChanged(Sender);
  end;
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
    if Assigned(FOnSelectChanged) then
      FOnSelectChanged(Sender);
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataLookup.CodeNoExit(Sender: TObject);
begin
  if ( FShowWarning ) and
     ( CodeNo.Text <> EmptyStr ) and
     ( CodeName.Text = EmptyStr ) then
  begin
    CodeNo.Text := '';
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

function TDataLookup.FieldByName(const AFieldName: string): TField;
begin
  if Assigned( FDataSet ) then
    Result := FDataSet.FieldByName( AFieldName )
  else
    Result := FDataReader.FieldByName( AFieldName );  
end;

{ ---------------------------------------------------------------------------- }


end.
