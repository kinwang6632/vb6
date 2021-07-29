unit frmSelUserU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  utlUCU, ComCtrls, StdCtrls, Buttons;

type
  TfrmSelUser = class(TForm)
    lsvUser: TListView;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure lsvUserDblClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    { Private declarations }
    FSubjectNode: TTreeNode; 
  public
    { Public declarations }
    procedure Select(UserGrpNode: TTreeNode; ListItems: TListItems);
  end;

  procedure ShowSelectUser(UserGrpNode: TTreeNode; ListItems: TListItems);

var
  frmSelUser: TfrmSelUser;

implementation

{$R *.DFM}

procedure ShowSelectUser(UserGrpNode: TTreeNode; ListItems: TListItems);
begin
  frmSelUser := TfrmSelUser.Create(Application);
  frmSelUser.Select(UserGrpNode, ListItems);
  frmSelUser.Free;
end;

{ TfrmSelUser }

procedure TfrmSelUser.Select(UserGrpNode: TTreeNode; ListItems: TListItems);
var
  ii: Integer;
  ListItem: TListItem;
begin
  FSubjectNode := UserGrpNode;
  Caption := '新增使用者選擇';
  lsvUser.Items.BeginUpdate;
  for ii:=0 to UCMgr.UserCount-1 do
  begin
    if TUserGrp(UserGrpNode.Data).HasUser(UCMgr.Users[ii]) then Continue;
    ListItem := lsvUser.Items.Add;
    ListItem.Data := UCMgr.Users[ii];
    UCMgr.DisplayUser(ListItem);
  end;
  lsvUser.Items.EndUpdate;
  if ShowModal = mrOk then
  begin
    for ii:=0 to lsvUser.Items.Count-1 do
      if lsvUser.Items[ii].Selected then
        UCMgr.AddUserToUserGrp(FSubjectNode,
          TUser(lsvUser.Items[ii].Data), ListItems);
  end;
end;

procedure TfrmSelUser.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  if (ModalResult = mrOk) and
    not Assigned(lsvUser.Selected) then CanClose := False;
end;

procedure TfrmSelUser.lsvUserDblClick(Sender: TObject);
begin
  ModalResult := mrOk;
end;

procedure TfrmSelUser.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_ESCAPE then ModalResult := mrCancel;
end;

end.
