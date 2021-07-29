unit DBDatePicker;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, Db, DbCtrls;

type
  TDBDatePicker = class(TDateTimePicker)
  private
    { Private declarations }
    FDataLink: TFieldDataLink;
    FOrgOnCloseUp: TNotifyEvent;
    procedure DataChange(Sender: TObject);
    procedure UpdateData(Sender: TObject);
    function GetDataField: string;
    function GetDataSource: TDataSource;
    function GetField: TField;
    function GetReadOnly: Boolean;
    procedure SetDataField(const Value: string);
    procedure SetDataSource(Value: TDataSource);
    procedure SetReadOnly(Value: Boolean);
    procedure ChangeOnCloseUp(Sender: TObject);
  protected
    { Protected declarations }
  public
    { Public declarations }
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    property Field: TField read GetField;
  published
    { Published declarations }
    property DataField: string read GetDataField write SetDataField;
    property DataSource: TDataSource read GetDataSource write SetDataSource;
    property ReadOnly: Boolean read GetReadOnly write SetReadOnly default False;
  end;

procedure Register;

implementation

procedure Register;
begin
  RegisterComponents('DbUtil', [TDBDatePicker]);
end;

constructor TDBDatePicker.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FDataLink := TFieldDataLink.Create;
  FDataLink.Control := Self;
  FDataLink.OnDataChange := DataChange;
  FDataLink.OnUpdateData := UpdateData;
  FOrgOnCloseUp := OnCloseUp;
  OnCloseUp := ChangeOnCloseUp;
end;

destructor TDBDatePicker.Destroy;
begin
  FDataLink.Free;
  FDataLink := nil;
  inherited Destroy;
end;

procedure TDBDatePicker.ChangeOnCloseUp(Sender: TObject);
begin
  if FDataLink.Editing then FDataLink.Modified;
  if Assigned(FOrgOnCloseUp) then FOrgonCloseUp(Sender);
end;

procedure TDBDatePicker.DataChange(Sender: TObject);
begin
  if FDataLink.Field <> nil then
    Self.Date := FDataLink.Field.AsDateTime;
end;

procedure TDBDatePicker.UpdateData(Sender: TObject);
begin
  if FDataLink.Field <> nil then FDataLink.Field.AsDateTime := Self.Date;
end;


function TDBDatePicker.GetDataSource: TDataSource;
begin
  Result := FDataLink.DataSource;
end;

procedure TDBDatePicker.SetDataSource(Value: TDataSource);
begin
  FDataLink.DataSource := Value;
  if Value <> nil then Value.FreeNotification(Self);
end;

function TDBDatePicker.GetDataField: string;
begin
  Result := FDataLink.FieldName;
end;

procedure TDBDatePicker.SetDataField(const Value: string);
begin
  FDataLink.FieldName := Value;
end;

function TDBDatePicker.GetReadOnly: Boolean;
begin
  Result := FDataLink.ReadOnly;
end;

procedure TDBDatePicker.SetReadOnly(Value: Boolean);
begin
  FDataLink.ReadOnly := Value;
end;

function TDBDatePicker.GetField: TField;
begin
  Result := FDataLink.Field;
end;

end.
