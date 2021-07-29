unit DBListView;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, Db;

type
  TDBListViewDataLink = class;

  TDBListView = class(TListView)
  private
    { Private declarations }
    FDataLink: TDBListViewDataLink;
    function GetDataSource: TDataSource;
    procedure SetDataSource(const Value: TDataSource);
  protected
    { Protected declarations }
    procedure DataChanged;
    procedure ActiveChanged;
    procedure EditingChanged;    
  public
    { Public declarations }
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  published
    { Published declarations }
    property DataSource: TDataSource read GetDataSource write SetDataSource;
  end;

  TDBListViewDataLink = class(TDataLink)
  private
    FDBlvu: TDBListView;
  protected
    procedure EditingChanged; override;
    procedure DataSetChanged; override;
    procedure ActiveChanged; override;
  public
    constructor Create(ALvu: TDBListView);
    destructor Destroy; override;
  end;

procedure Register;

implementation

procedure Register;
begin
  RegisterComponents('DBUtil', [TDBListView]);
end;

{ TDBListView }

procedure TDBListView.ActiveChanged;
var
   i : Integer ;
   NewColumn : TListColumn;
   ListItem: TListItem;

begin

     if not FDataLink.Active then
     begin
       Columns.Clear;
       Items.Clear;
       exit;
     end;

     if ViewStyle<>vsReport then Exit;
     ListItem := nil;
     Items.Clear;
     for i:=0 to FDataLink.dataset.FieldCount -1 do
     begin
       NewColumn := Columns.Add;
       NewColumn.Width := FDataLink.dataset.Fields[i].DisplayWidth*4;
       NewColumn.Caption := FDataLink.dataset.Fields[i].DisplayLabel;
     end;

     while not FDataLink.dataset.Eof do
     begin
       for i:=0 to FDataLink.dataset.FieldCount -1 do
       begin
         if i=0 then
         begin
           ListItem := Items.Add;
           ListItem.Caption := FDataLink.dataset.Fields[i].AsString;
         end
         else
           ListItem.SubItems.Add(FDataLink.dataset.Fields[i].AsString);
       end;
       FDataLink.dataset.Next;
     end;


end;

constructor TDBListView.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FDataLink := TDBListViewDataLink.Create(Self);
end;

procedure TDBListView.DataChanged;
begin
{}
end;

destructor TDBListView.Destroy;
begin
  FDataLink.Free;
  FDataLink := nil;
  inherited Destroy;
end;

procedure TDBListView.EditingChanged;
begin
{}
end;

function TDBListView.GetDataSource: TDataSource;
begin
  Result := FDataLink.DataSource;
end;

procedure TDBListView.SetDataSource(const Value: TDataSource);
begin
  FDataLink.DataSource := Value;
  if Value <> nil then Value.FreeNotification(Self);
end;

{ TDBListViewDataLink }

procedure TDBListViewDataLink.ActiveChanged;
begin
  if FDBLvu <> nil then FDBLvu.ActiveChanged;
end;

constructor TDBListViewDataLink.Create(ALvu: TDBListView);
begin
  inherited Create;
  FDBLvu := ALvu;
  VisualControl := True;
end;

procedure TDBListViewDataLink.DataSetChanged;
begin
  if FDBLvu <> nil then FDBLvu.DataChanged;
end;

destructor TDBListViewDataLink.Destroy;
begin
  FDBLvu := nil;
  inherited Destroy;
end;

procedure TDBListViewDataLink.EditingChanged;
begin
{}
end;

end.
