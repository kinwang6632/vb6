unit EnumDBRadioGroup;

interface

uses Windows, SysUtils, Messages, Classes, Controls, Forms,
     Graphics, Menus, StdCtrls, ExtCtrls, Mask, Buttons, ComCtrls, Db, dbctrls;
type
  TEnumDBRadioGroup = class(TCustomRadioGroup)
  private
    FDataLink: TFieldDataLink;
    FValue: string;
    FValues: TStrings;
    FInSetValue: Boolean;
    FOnChange: TNotifyEvent;
    FOrgOnGetText: TFieldGetTextEvent;
    procedure ActiveChange(Sender: TObject);
    procedure DataChange(Sender: TObject);
    procedure UpdateData(Sender: TObject);
    function GetDataField: string;
    function GetDataSource: TDataSource;
    function GetField: TField;
    function GetReadOnly: Boolean;
    function GetButtonValue(Index: Integer): string;
    procedure SetDataField(const Value: string);
    procedure SetDataSource(Value: TDataSource);
    procedure SetReadOnly(Value: Boolean);
    procedure SetValue(const Value: string);
    procedure SetItems(Value: TStrings);
    procedure SetValues(Value: TStrings);
    procedure CMExit(var Message: TCMExit); message CM_EXIT;
    procedure ChangeOnGetText(Sender: TField; var aText: string;
      DisplayText: Boolean);
  protected
    procedure Change; dynamic;
    procedure Click; override;
    procedure KeyPress(var Key: Char); override;
    function CanModify: Boolean; override;
    procedure Notification(AComponent: TComponent;
      Operation: TOperation); override;
    property DataLink: TFieldDataLink read FDataLink;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    function ExecuteAction(Action: TBasicAction): Boolean; override;
    function UpdateAction(Action: TBasicAction): Boolean; override;
    function UseRightToLeftAlignment: Boolean; override;
    property Field: TField read GetField;
    property ItemIndex;
    property Value: string read FValue write SetValue;
  published
    property Align;
    property Anchors;
    property BiDiMode;
    property Caption;
    property Color;
    property Columns;
    property Constraints;
    property Ctl3D;
    property DataField: string read GetDataField write SetDataField;
    property DataSource: TDataSource read GetDataSource write SetDataSource;
    property DragCursor;
    property DragKind;
    property DragMode;
    property Enabled;
    property Font;
    property Items write SetItems;
    property ParentBiDiMode;
    property ParentColor;
    property ParentCtl3D;
    property ParentFont;
    property ParentShowHint;
    property PopupMenu;
    property ReadOnly: Boolean read GetReadOnly write SetReadOnly default False;
    property ShowHint;
    property TabOrder;
    property TabStop;
    property Values: TStrings read FValues write SetValues;
    property Visible;
    property OnChange: TNotifyEvent read FOnChange write FOnChange;
    property OnClick;
    property OnDragDrop;
    property OnDragOver;
    property OnEndDock;
    property OnEndDrag;
    property OnEnter;
    property OnExit;
    property OnStartDock;
    property OnStartDrag;
  end;
procedure Register;

implementation

uses Clipbrd, DBConsts, Dialogs;

procedure Register;
begin
  RegisterComponents('DbUtil', [TEnumDBRadioGroup]);
end;


procedure TEnumDBRadioGroup.ActiveChange(Sender: TObject);
begin
//870910 modified
  if (Values.Count>0) and FDataLink.Active and Assigned(FDataLink.Field) then
    FDataLink.Field.OnGetText := ChangeOnGetText;
//870910 modified
end;

procedure TEnumDBRadioGroup.ChangeOnGetText(Sender: TField; var aText: string;
  DisplayText: Boolean);
begin
//870910 modified
  if Values.IndexOf(Sender.AsString) >= 0 then
    aText := Items[Values.IndexOf(Sender.AsString)] ;
  if Assigned(FOrgOnGetText) then
    FOrgOnGetText(Sender, aText, DisplayText);
//870910 modified
end;


constructor TEnumDBRadioGroup.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FDataLink := TFieldDataLink.Create;
  FDataLink.Control := Self;
  FDataLink.OnDataChange := DataChange;
  FDataLink.OnUpdateData := UpdateData;
  //870910 modified
  FDataLink.OnActiveChange := ActiveChange;
  //870910 modified
  FValues := TStringList.Create;
end;

destructor TEnumDBRadioGroup.Destroy;
begin
  FDataLink.Free;
  //870910 modified
  if Assigned(FDataLink.Field) then
    FDataLink.Field.OnGetText := nil ;
  //870910 modified
  FDataLink := nil;
  FValues.Free;
  inherited Destroy;
end;

procedure TEnumDBRadioGroup.Notification(AComponent: TComponent;
  Operation: TOperation);
begin
  inherited Notification(AComponent, Operation);
  if (Operation = opRemove) and (FDataLink <> nil) and
    (AComponent = DataSource) then DataSource := nil;
end;

function TEnumDBRadioGroup.UseRightToLeftAlignment: Boolean;
begin
  Result := DBUseRightToLeftAlignment(Self, Field);
end;

procedure TEnumDBRadioGroup.DataChange(Sender: TObject);
begin
  if FDataLink.Field <> nil then
    Value := FDataLink.Field.AsString else
    Value := '';
end;

procedure TEnumDBRadioGroup.UpdateData(Sender: TObject);
begin
  if FDataLink.Field <> nil then FDataLink.Field.AsString := Value;
end;

function TEnumDBRadioGroup.GetDataSource: TDataSource;
begin
  Result := FDataLink.DataSource;
end;

procedure TEnumDBRadioGroup.SetDataSource(Value: TDataSource);
begin
  FDataLink.DataSource := Value;
  if Value <> nil then Value.FreeNotification(Self);
end;

function TEnumDBRadioGroup.GetDataField: string;
begin
  Result := FDataLink.FieldName;
end;

procedure TEnumDBRadioGroup.SetDataField(const Value: string);
begin
  FDataLink.FieldName := Value;
end;

function TEnumDBRadioGroup.GetReadOnly: Boolean;
begin
  Result := FDataLink.ReadOnly;
end;

procedure TEnumDBRadioGroup.SetReadOnly(Value: Boolean);
begin
  FDataLink.ReadOnly := Value;
end;

function TEnumDBRadioGroup.GetField: TField;
begin
  Result := FDataLink.Field;
end;

function TEnumDBRadioGroup.GetButtonValue(Index: Integer): string;
begin
  if (Index < FValues.Count) and (FValues[Index] <> '') then
    Result := FValues[Index]
  else if Index < Items.Count then
    Result := Items[Index]
  else
    Result := '';
end;

procedure TEnumDBRadioGroup.SetValue(const Value: string);
var
  I, Index: Integer;
begin
  if FValue <> Value then
  begin
    FInSetValue := True;
    try
      Index := -1;
      for I := 0 to Items.Count - 1 do
        if Value = GetButtonValue(I) then
        begin
          Index := I;
          Break;
        end;
      ItemIndex := Index;
    finally
      FInSetValue := False;
    end;
    FValue := Value;
    Change;
  end;
end;

procedure TEnumDBRadioGroup.CMExit(var Message: TCMExit);
begin
  try
    FDataLink.UpdateRecord;
  except
    if ItemIndex >= 0 then
      TRadioButton(Controls[ItemIndex]).SetFocus else
      TRadioButton(Controls[0]).SetFocus;
    raise;
  end;
  inherited;
end;

procedure TEnumDBRadioGroup.Click;
begin
  if not FInSetValue then
  begin
    inherited Click;
    if ItemIndex >= 0 then
      Value := GetButtonValue(ItemIndex);
    if FDataLink.Editing then
      FDataLink.Modified;
  end;
end;

procedure TEnumDBRadioGroup.SetItems(Value: TStrings);
begin
  Items.Assign(Value);
  DataChange(Self);
end;

procedure TEnumDBRadioGroup.SetValues(Value: TStrings);
begin
  FValues.Assign(Value);
  DataChange(Self);
end;

procedure TEnumDBRadioGroup.Change;
begin
  if Assigned(FOnChange) then FOnChange(Self);
end;

procedure TEnumDBRadioGroup.KeyPress(var Key: Char);
begin
  inherited KeyPress(Key);
  case Key of
    #8, ' ': FDataLink.Edit;
    #27: FDataLink.Reset;
  end;
end;

function TEnumDBRadioGroup.CanModify: Boolean;
begin
  Result := FDataLink.Edit;
end;

function TEnumDBRadioGroup.ExecuteAction(Action: TBasicAction): Boolean;
begin
  Result := inherited ExecuteAction(Action) or (DataLink <> nil) and
    DataLink.ExecuteAction(Action);
end;

function TEnumDBRadioGroup.UpdateAction(Action: TBasicAction): Boolean;
begin
  Result := inherited UpdateAction(Action) or (DataLink <> nil) and
    DataLink.UpdateAction(Action);
end;

end.
