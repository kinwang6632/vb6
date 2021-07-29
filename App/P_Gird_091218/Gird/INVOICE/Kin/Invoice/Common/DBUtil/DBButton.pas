unit DBButton;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Buttons, Db;

type

  TDBBtnDataLink = class;

  TDBBtnOperation = (boNew, boDelete, boUpdate, boSave, boCancel);

  TDBButton = class(TSpeedButton)
  private
    { Private declarations }
    FDataLink: TDBBtnDataLink;
    FOperation: TDBBtnOperation;
    FAfterOperate: TNotifyEvent;
    FBeforeOperate: TNotifyEvent;
    function GetDataSource: TDataSource;
    procedure SetDataSource(Value: TDataSource);
    procedure SetOperation(const Value: TDBBtnOperation);
  protected
    { Protected declarations }
    procedure DataChanged;
    procedure EditingChanged;
    procedure ActiveChanged;
  public
    { Public declarations }
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure Click; override;
  published
    { Published declarations }
    property DataSource: TDataSource read GetDataSource write SetDataSource;
    property Operation: TDBBtnOperation read FOperation write SetOperation;
    property BeforeOperate: TNotifyEvent read FBeforeOperate write FBeforeOperate;
    property AfterOperate: TNotifyEvent read FAfterOperate write FAfterOperate;
  end;

{ TNavDataLink }

  TDBBtnDataLink = class(TDataLink)
  private
    FDBBtn: TDBButton;
  protected
    procedure EditingChanged; override;
    procedure DataSetChanged; override;
    procedure ActiveChanged; override;
  public
    constructor Create(ABtn: TDBButton);
    destructor Destroy; override;
  end;

procedure Register;

implementation

procedure Register;
begin
  RegisterComponents('DBUtil', [TDBButton]);
end;

constructor TDBButton.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FDataLink := TDBBtnDataLink.Create(Self);
end;

destructor TDBButton.Destroy;
begin
  FDataLink.Free;
  FDataLink := nil;
  inherited Destroy;
end;

procedure TDBButton.Click;
begin
  if Assigned(FBeforeOperate) then FBeforeOperate(Self);
  case FOperation of
    boNew: FDataLink.DataSet.Append;
    boDelete:
      if MessageDlg('確定要刪除嗎?', mtConfirmation,
        [mbYes, mbNo], 0) = mrYes then FDataLink.DataSet.Delete;
    boUpdate: FDataLink.DataSet.Edit;
    boSave:  FDataLink.DataSet.Post;
    boCancel:  FDataLink.DataSet.Cancel;
  end;
  inherited Click;
  if Assigned(FAfterOperate) then FAfterOperate(Self);
end;

function TDBButton.GetDataSource: TDataSource;
begin
  Result := FDataLink.DataSource;
end;

procedure TDBButton.SetDataSource(Value: TDataSource);
begin
  FDataLink.DataSource := Value;
  if Value <> nil then Value.FreeNotification(Self);
end;

procedure TDBButton.DataChanged;
var
  CanModify: Boolean;
begin
  CanModify := FDataLink.Active and FDataLink.DataSet.CanModify;
  case FOperation of
    boNew:
      Enabled := CanModify and not FDataLink.Editing;
    boUpdate, boDelete:
      Enabled := FDataLink.Active and FDataLink.DataSet.CanModify and
        not (FDataLink.DataSet.BOF and FDataLink.DataSet.EOF)
        and not FDataLink.Editing;
    boSave, boCancel:
      Enabled := CanModify and FDataLink.Editing;
  end;
end;

procedure TDBButton.EditingChanged;
var
  CanModify: Boolean;
begin
  CanModify := FDataLink.Active and FDataLink.DataSet.CanModify;
  case FOperation of
    boNew:
      Enabled := CanModify and not FDataLink.Editing;
    boUpdate, boDelete:
      Enabled := FDataLink.Active and FDataLink.DataSet.CanModify and
        not (FDataLink.DataSet.BOF and FDataLink.DataSet.EOF)
        and not FDataLink.Editing;
    boSave, boCancel:
      Enabled := CanModify and FDataLink.Editing;
  end;
end;

procedure TDBButton.ActiveChanged;
begin
  if not FDataLink.Active then Enabled := False
  else
  begin
    DataChanged;
  end;
end;

procedure TDBButton.SetOperation(const Value: TDBBtnOperation);
begin
  FOperation := Value;
end;

{ TDBBtnDataLink }

constructor TDBBtnDataLink.Create(ABtn: TDBButton);
begin
  inherited Create;
  FDBBtn := ABtn;
  VisualControl := True;
end;

destructor TDBBtnDataLink.Destroy;
begin
  FDBBtn := nil;
  inherited Destroy;
end;

procedure TDBBtnDataLink.EditingChanged;
begin
  if FDBBtn <> nil then FDBBtn.EditingChanged;
end;

procedure TDBBtnDataLink.DataSetChanged;
begin
  if FDBBtn <> nil then FDBBtn.DataChanged;
end;

procedure TDBBtnDataLink.ActiveChanged;
begin
  if FDBBtn <> nil then FDBBtn.ActiveChanged;
end;

end.
