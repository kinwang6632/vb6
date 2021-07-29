unit EnumDBComboBox;

interface

uses Windows, SysUtils, Messages, Classes, Controls, Forms,
     Graphics, Menus, StdCtrls, ExtCtrls, Mask, Buttons, ComCtrls, Db, dbctrls;

type
  TEnumDBComboBox = class(TCustomComboBox)
  private
    FDataLink: TFieldDataLink;
    FPaintControl: TPaintControl;
    FValues : string;
    FValueList : TStrings;
    FOrgOnGetText: TFieldGetTextEvent;
    procedure DataChange(Sender: TObject);
    procedure EditingChange(Sender: TObject);
    procedure ActiveChange(Sender: TObject);
    function GetComboText: string;
    function GetDataField: string;
    function GetDataSource: TDataSource;
    function GetField: TField;
    function GetReadOnly: Boolean;
    procedure SetComboText(const Value: string);
    procedure SetDataField(const Value: string);
    procedure SetDataSource(Value: TDataSource);
    procedure SetEditReadOnly;
    procedure SetItems(Value: TStrings);
    procedure SetReadOnly(Value: Boolean);
    procedure UpdateData(Sender: TObject);
    procedure CMEnter(var Message: TCMEnter); message CM_ENTER;
    procedure CMExit(var Message: TCMExit); message CM_EXIT;
    procedure CMGetDataLink(var Message: TMessage); message CM_GETDATALINK;
    procedure WMPaint(var Message: TWMPaint); message WM_PAINT;
    procedure SetValues(const Value: string);
    procedure ChangeOnGetText(Sender: TField; var aText: string;
      DisplayText: Boolean);
  protected
    procedure Change; override;
    procedure Click; override;
    procedure ComboWndProc(var Message: TMessage; ComboWnd: HWnd;
      ComboProc: Pointer); override;
    procedure CreateWnd; override;
    procedure DropDown; override;
    procedure KeyDown(var Key: Word; Shift: TShiftState); override;
    procedure KeyPress(var Key: Char); override;
    procedure Loaded; override;
    procedure Notification(AComponent: TComponent;
      Operation: TOperation); override;
    procedure SetStyle(Value: TComboboxStyle); override;
    procedure WndProc(var Message: TMessage); override;
    function ParseStrings(Strs:string; Sep:Char) : TStringList ;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    function ExecuteAction(Action: TBasicAction): Boolean; override;
    function UpdateAction(Action: TBasicAction): Boolean; override;
    function UseRightToLeftAlignment: Boolean; override;
    property Field: TField read GetField;
    property Text;
  published
    property Style; {Must be published before Items}
    property Anchors;
    property BiDiMode;
    property Color;
    property Constraints;
    property Ctl3D;
    property DataField: string read GetDataField write SetDataField;
    property DataSource: TDataSource read GetDataSource write SetDataSource;
    property DragCursor;
    property DragKind;
    property DragMode;
    property DropDownCount;
    property Enabled;
    property Values:string read FValues write SetValues ;
    property Font;
    property ImeMode;
    property ImeName;
    property ItemHeight;
    property Items write SetItems;
    property ParentBiDiMode;
    property ParentColor;
    property ParentCtl3D;
    property ParentFont;
    property ParentShowHint;
    property PopupMenu;
    property ReadOnly: Boolean read GetReadOnly write SetReadOnly default False;
    property ShowHint;
    property Sorted;
    property TabOrder;
    property TabStop;
    property Visible;
    property OnChange;
    property OnClick;
    property OnDblClick;
    property OnDragDrop;
    property OnDragOver;
    property OnDrawItem;
    property OnDropDown;
    property OnEndDock;
    property OnEndDrag;
    property OnEnter;
    property OnExit;
    property OnKeyDown;
    property OnKeyPress;
    property OnKeyUp;
    property OnMeasureItem;
    property OnStartDock;
    property OnStartDrag;
  end;

procedure Register;

implementation

uses Clipbrd, DBConsts, Dialogs;

procedure Register;
begin
  RegisterComponents('DBUtil', [TEnumDBComboBox]);
end;

procedure TEnumDBComboBox.ActiveChange(Sender: TObject);
begin
//870910 modified
  if Assigned(FValueList) and FDataLink.Active then
    FDataLink.Field.OnGetText := ChangeOnGetText;
//870910 modified
end;

procedure TEnumDBComboBox.ChangeOnGetText(Sender: TField; var aText: string;
  DisplayText: Boolean);
begin
//870910 modified
  if Assigned(FValueList) then
    aText := Items[FValueList.IndexOf(Sender.AsString)] ;
  if Assigned(FOrgOnGetText) then
    FOrgOnGetText(Sender, aText, DisplayText);
//870910 modified
end;

constructor TEnumDBComboBox.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  ControlStyle := ControlStyle + [csReplicatable];
  FDataLink := TFieldDataLink.Create;
  FDataLink.Control := Self;
  FDataLink.OnDataChange := DataChange;
  FDataLink.OnUpdateData := UpdateData;
  FDataLink.OnEditingChange := EditingChange;
  //870910 modified
  FDataLink.OnActiveChange := ActiveChange;
  //870910 modified
  FPaintControl := TPaintControl.Create(Self, 'COMBOBOX');
end;

destructor TEnumDBComboBox.Destroy;
begin
  FPaintControl.Free;
  //870910 modified
  if Assigned(FDataLink.Field) then
    FDataLink.Field.OnGetText := nil ;
  //870910 modified
  FDataLink.Free;
  FDataLink := nil;
  if Assigned(FValueList) then FValueList.Free ;
  inherited Destroy;
end;

procedure TEnumDBComboBox.Loaded;
begin
  inherited Loaded;
  if (csDesigning in ComponentState) then DataChange(Self);
end;

procedure TEnumDBComboBox.Notification(AComponent: TComponent;
  Operation: TOperation);
begin
  inherited Notification(AComponent, Operation);
  if (Operation = opRemove) and (FDataLink <> nil) and
    (AComponent = DataSource) then DataSource := nil;
end;

procedure TEnumDBComboBox.CreateWnd;
begin
  inherited CreateWnd;
  SetEditReadOnly;
end;

procedure TEnumDBComboBox.DataChange(Sender: TObject);
var
  tmp_Text : String;
  ii : Integer;
begin
  //870910 modified
  if not (Style = csSimple) and DroppedDown then Exit;
  if FDataLink.Field <> nil then
  begin
    tmp_Text := '';
    if not Assigned(FValueList) then
      tmp_Text := FDataLink.Field.AsString
    else
      for ii:=0 to FValueList.Count-1 do
      begin
        if FDataLink.Field.AsString = FValueList[ii] then
        begin
          tmp_Text := Items.Strings[ii];
          break;
        end;
      end;
    SetComboText(tmp_Text);
  end
  else
    if csDesigning in ComponentState then
      SetComboText(Name)
    else
      SetComboText('');
  //870910 modified
end;

procedure TEnumDBComboBox.UpdateData(Sender: TObject);
begin
  //870910 modified
  if not Assigned(FValueList) then
    FDataLink.Field.Text := GetComboText;//±NComboBox Text¼g¦^DataField
  //870910 modified
end;

procedure TEnumDBComboBox.SetComboText(const Value: string);
var
  I: Integer;
  Redraw: Boolean;
begin
  if Value <> GetComboText then
  begin
    if Style <> csDropDown then
    begin
      Redraw := (Style <> csSimple) and HandleAllocated;
      if Redraw then SendMessage(Handle, WM_SETREDRAW, 0, 0);
      try
        if Value = '' then I := -1 else I := Items.IndexOf(Value);
        ItemIndex := I;
      finally
        if Redraw then
        begin
          SendMessage(Handle, WM_SETREDRAW, 1, 0);
          Invalidate;
        end;
      end;
      if I >= 0 then Exit;
    end;
    if Style in [csDropDown, csSimple] then Text := Value;
  end;
end;

function TEnumDBComboBox.GetComboText: string;
var
  I: Integer;
begin
  if Style in [csDropDown, csSimple] then Result := Text else
  begin
    I := ItemIndex;
    if I < 0 then Result := '' else Result := Items[I];
  end;
end;

procedure TEnumDBComboBox.Change;
begin
  FDataLink.Edit;
  //870910 modified
  if Assigned(FValueList) and (ItemIndex >= 0) then
    FDataLink.Field.AsString := FValueList[ItemIndex] ;
  //870910 modified
  inherited Change;
  FDataLink.Modified;
end;

procedure TEnumDBComboBox.Click;
begin
  FDataLink.Edit;
  inherited Click;
  FDataLink.Modified;
end;

procedure TEnumDBComboBox.DropDown;
begin
  inherited DropDown;
end;

function TEnumDBComboBox.GetDataSource: TDataSource;
begin
  Result := FDataLink.DataSource;
end;

procedure TEnumDBComboBox.SetDataSource(Value: TDataSource);
begin
  if not (FDataLink.DataSourceFixed and (csLoading in ComponentState)) then
    FDataLink.DataSource := Value;
  if Value <> nil then Value.FreeNotification(Self);
end;

function TEnumDBComboBox.GetDataField: string;
begin
  Result := FDataLink.FieldName;
end;

procedure TEnumDBComboBox.SetDataField(const Value: string);
begin
  FDataLink.FieldName := Value;
end;

function TEnumDBComboBox.GetReadOnly: Boolean;
begin
  Result := FDataLink.ReadOnly;
end;

procedure TEnumDBComboBox.SetReadOnly(Value: Boolean);
begin
  FDataLink.ReadOnly := Value;
end;

function TEnumDBComboBox.GetField: TField;
begin
  Result := FDataLink.Field;
end;

procedure TEnumDBComboBox.KeyDown(var Key: Word; Shift: TShiftState);
begin
  inherited KeyDown(Key, Shift);
  if Key in [VK_BACK, VK_DELETE, VK_UP, VK_DOWN, 32..255] then
  begin
    if not FDataLink.Edit and (Key in [VK_UP, VK_DOWN]) then
      Key := 0;
  end;
end;

procedure TEnumDBComboBox.KeyPress(var Key: Char);
begin
  inherited KeyPress(Key);
  if (Key in [#32..#255]) and (FDataLink.Field <> nil) and
    not FDataLink.Field.IsValidChar(Key) then
  begin
    MessageBeep(0);
    Key := #0;
  end;
  case Key of
    ^H, ^V, ^X, #32..#255:
      FDataLink.Edit;
    #27:
      begin
        FDataLink.Reset;
        SelectAll;
      end;
  end;
end;

procedure TEnumDBComboBox.EditingChange(Sender: TObject);
begin
  SetEditReadOnly;
end;

procedure TEnumDBComboBox.SetEditReadOnly;
begin
  if (Style in [csDropDown, csSimple]) and HandleAllocated then
    SendMessage(EditHandle, EM_SETREADONLY, Ord(not FDataLink.Editing), 0);
end;

procedure TEnumDBComboBox.WndProc(var Message: TMessage);
begin
  if not (csDesigning in ComponentState) then
    case Message.Msg of
      WM_COMMAND:
        if TWMCommand(Message).NotifyCode = CBN_SELCHANGE then
          if not FDataLink.Edit then
          begin
            if Style <> csSimple then
              PostMessage(Handle, CB_SHOWDROPDOWN, 0, 0);
            Exit;
          end;
      CB_SHOWDROPDOWN:
        if Message.WParam <> 0 then FDataLink.Edit else
          if not FDataLink.Editing then DataChange(Self); {Restore text}
      WM_CREATE,
      WM_WINDOWPOSCHANGED,
      CM_FONTCHANGED:
        FPaintControl.DestroyHandle;
    end;
  inherited WndProc(Message);
end;

procedure TEnumDBComboBox.ComboWndProc(var Message: TMessage; ComboWnd: HWnd;
  ComboProc: Pointer);
begin
  if not (csDesigning in ComponentState) then
    case Message.Msg of
      WM_LBUTTONDOWN:
        if (Style = csSimple) and (ComboWnd <> EditHandle) then
          if not FDataLink.Edit then Exit;
    end;
  inherited ComboWndProc(Message, ComboWnd, ComboProc);
end;

procedure TEnumDBComboBox.CMEnter(var Message: TCMEnter);
begin
  inherited;
  if SysLocale.FarEast and FDataLink.CanModify then
    SendMessage(EditHandle, EM_SETREADONLY, Ord(False), 0);
end;

procedure TEnumDBComboBox.CMExit(var Message: TCMExit);
begin
  try
    FDataLink.UpdateRecord;
  except
    SelectAll;
    SetFocus;
    raise;
  end;
  inherited;
end;

procedure TEnumDBComboBox.WMPaint(var Message: TWMPaint);
var
  S: string;
  R: TRect;
  P: TPoint;
  Child: HWND;
begin
  if csPaintCopy in ControlState then
  begin
    if FDataLink.Field <> nil then S := Items[ItemIndex] else S := '';
    if Style = csDropDown then
    begin
      SendMessage(FPaintControl.Handle, WM_SETTEXT, 0, Longint(PChar(S)));
      SendMessage(FPaintControl.Handle, WM_PAINT, Message.DC, 0);
      Child := GetWindow(FPaintControl.Handle, GW_CHILD);
      if Child <> 0 then
      begin
        Windows.GetClientRect(Child, R);
        Windows.MapWindowPoints(Child, FPaintControl.Handle, R.TopLeft, 2);
        GetWindowOrgEx(Message.DC, P);
        SetWindowOrgEx(Message.DC, P.X - R.Left, P.Y - R.Top, nil);
        IntersectClipRect(Message.DC, 0, 0, R.Right - R.Left, R.Bottom - R.Top);
        SendMessage(Child, WM_PAINT, Message.DC, 0);
      end;
    end else
    begin
      SendMessage(FPaintControl.Handle, CB_RESETCONTENT, 0, 0);
      if Items.IndexOf(S) <> -1 then
      begin
        SendMessage(FPaintControl.Handle, CB_ADDSTRING, 0, Longint(PChar(S)));
        SendMessage(FPaintControl.Handle, CB_SETCURSEL, 0, 0);
      end;
      SendMessage(FPaintControl.Handle, WM_PAINT, Message.DC, 0);
    end;
  end else
    inherited;
end;

procedure TEnumDBComboBox.SetItems(Value: TStrings);
begin
  Items.Assign(Value);
  DataChange(Self);
end;

procedure TEnumDBComboBox.SetStyle(Value: TComboboxStyle);
begin
{
  if (Value = csSimple) and Assigned(FDatalink) and FDatalink.DatasourceFixed then
    DatabaseError(SNotReplicatable);
}    
  inherited SetStyle(Value);
end;

function TEnumDBComboBox.UseRightToLeftAlignment: Boolean;
begin
  Result := DBUseRightToLeftAlignment(Self, Field);
end;

procedure TEnumDBComboBox.CMGetDatalink(var Message: TMessage);
begin
  Message.Result := Integer(FDataLink);
end;

function TEnumDBComboBox.ExecuteAction(Action: TBasicAction): Boolean;
begin
  Result := inherited ExecuteAction(Action) or (FDataLink <> nil) and
    FDataLink.ExecuteAction(Action);
end;

function TEnumDBComboBox.UpdateAction(Action: TBasicAction): Boolean;
begin
  Result := inherited UpdateAction(Action) or (FDataLink <> nil) and
    FDataLink.UpdateAction(Action);
end;

//870910 modified
procedure TEnumDBComboBox.SetValues(const Value: string);
begin
  FValues := Value;
  FValueList := ParseStrings(Value, ';');
end;

function TEnumDBComboBox.ParseStrings(Strs: string;
  Sep: Char): TStringList;
var
  ii, jj : Integer ;
  slst : TStringList ;
begin
  if Strs = '' then
  begin
    Result := nil ;
    Exit ;
  end ;
  slst := TStringList.Create ;
  jj := 1 ;
  for ii:= 1 to Length(Strs) do
  begin
    if Strs[ii] = Sep then
    begin
      slst.Add(Trim(Copy(Strs, jj, ii-jj))) ;
      jj := ii + 1 ;
    end ;
  end ;
  if jj <> Length(Strs) + 1 then
    slst.Add(Trim(Copy(Strs, jj, Length(Strs)-jj+1))) ;
  Result := slst ;
end ;
//870910 modified
end.
