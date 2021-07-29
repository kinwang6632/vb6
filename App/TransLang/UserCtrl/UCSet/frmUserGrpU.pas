unit frmUserGrpU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, ExtCtrls, ToolWin, Menus, IniFiles, utlUCU;

type
  TfrmUserGrp = class(TForm)
    trvUserGrp: TTreeView;
    Splitter1: TSplitter;
    lsvUser: TListView;
    pmuUserGrp: TPopupMenu;
    mnuNewUserGrp: TMenuItem;
    mnuUpdateUserGrp: TMenuItem;
    mnuDeleteUserGrp: TMenuItem;
    N1: TMenuItem;
    mnuUpUserGrp: TMenuItem;
    mnuDownUserGrp: TMenuItem;
    pmuUser: TPopupMenu;
    mnuNewUser: TMenuItem;
    mnuUpdateUser: TMenuItem;
    mnuDeleteUser: TMenuItem;
    N2: TMenuItem;
    mnuUserGrpFunc: TMenuItem;
    N3: TMenuItem;
    mnuUpUser: TMenuItem;
    mnuDownUser: TMenuItem;
    procedure trvUserGrpChange(Sender: TObject; Node: TTreeNode);
    procedure pmuUserGrpPopup(Sender: TObject);
    procedure pmuUserPopup(Sender: TObject);
    procedure NewUserGrp(Sender: TObject);
    procedure UpdateUserGrp(Sender: TObject);
    procedure UpdateUser(Sender: TObject);
    procedure DeleteUserGrp(Sender: TObject);
    procedure DeleteUser(Sender: TObject);
    procedure mnuUserGrpFuncClick(Sender: TObject);
  private
    { Private declarations }
    procedure NewUser(Sender: TObject);
    procedure SelectUser(Sender: TObject);
  public
    { Public declarations }
  end;

var
  frmUserGrp: TfrmUserGrp;

implementation

uses frmUserGrpAttrU, frmUserAttrU, frmSelFuncU, frmSelUserU;

{$R *.DFM}

procedure TfrmUserGrp.trvUserGrpChange(Sender: TObject; Node: TTreeNode);
var
  ii: Integer;
  UserGrp: TUserGrp;
  ListItem: TListItem;
begin
  lsvUser.Items.BeginUpdate;
  lsvUser.Items.Clear;
  UserGrp := Node.Data;
  for ii:=0 to UserGrp.UserCount-1 do
  begin
    ListItem := lsvUser.Items.Add;
    ListItem.Data := UserGrp.Users[ii];
    UCMgr.DisplayUser(ListItem);
  end;
  lsvUser.Items.EndUpdate;  
end;

procedure TfrmUserGrp.pmuUserGrpPopup(Sender: TObject);
begin
  trvUserGrp.Selected := trvUserGrp.Selected;
  mnuNewUserGrp.Enabled := False;
  mnuUpdateUserGrp.Enabled := False;
  mnuDeleteUserGrp.Enabled := False;
  mnuDownUserGrp.Enabled := False;
  mnuUpUserGrp.Enabled := False;
  mnuUserGrpFunc.Enabled := False;
  if not Assigned(trvUserGrp.Selected) then Exit;
  mnuNewUserGrp.Enabled := True;
  if trvUserGrp.Selected.AbsoluteIndex <> 0 then
  begin
    mnuUpdateUserGrp.Enabled := True;
    mnuDeleteUserGrp.Enabled := True;
    mnuDownUserGrp.Enabled := True;
    mnuUpUserGrp.Enabled := True;
    mnuUserGrpFunc.Enabled := True;    
  end;
end;

procedure TfrmUserGrp.pmuUserPopup(Sender: TObject);
begin
  mnuNewUser.Enabled := False;
  mnuUpdateUser.Enabled := False;
  mnuDeleteUser.Enabled := False;
  mnuUpUser.Enabled := False;
  mnuDownUser.Enabled := False;
  if not Assigned(trvUserGrp.Selected) then Exit;
  if Assigned(lsvUser.Selected) then
  begin
    if (trvUserGrp.Selected.AbsoluteIndex = 0) then mnuUpdateUser.Enabled := True;
    mnuDeleteUser.Enabled := True;
    mnuUpUser.Enabled := True;
    mnuDownUser.Enabled := True;    
  end;
  mnuNewUser.Enabled := True;
  mnuNewUser.OnClick := SelectUser;
  if (trvUserGrp.Selected.AbsoluteIndex = 0) then mnuNewUser.OnClick := NewUser;
end;

procedure TfrmUserGrp.NewUserGrp(Sender: TObject);
var
  UserGrp: TUserGrp;
begin
  UserGrp := TUserGrp.Create;
  if EditUserGrpAttr(UserGrp) = mrOk then UCMgr.NewUserGrp(UserGrp, trvUserGrp.Selected)
  else UserGrp.Free;
end;

procedure TfrmUserGrp.UpdateUserGrp(Sender: TObject);
var
  UserGrp: TUserGrp;
begin
  UserGrp := TUserGrp(trvUserGrp.Selected.Data);
  if EditUserGrpAttr(UserGrp) = mrOk then  trvUserGrp.Selected.Text := UserGrp.Desc;
end;

procedure TfrmUserGrp.DeleteUserGrp(Sender: TObject);
begin
  if MessageDlg('確定要刪除嗎??', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    UCMgr.DeleteUserGrp(trvUserGrp.Selected);
end;

procedure TfrmUserGrp.NewUser(Sender: TObject);
var
  User: TUser;
begin
  User := TUser.Create;
  if EditUserAttr(User) = mrOk then UCMgr.NewUser(User, lsvUser.Items)
  else User.Free;
end;

procedure TfrmUserGrp.UpdateUser(Sender: TObject);
var
  User: TUser;
begin
  User := TUser(lsvUser.Selected.Data);
  if EditUserAttr(User) = mrOk then UCMgr.DisplayUser(lsvUser.Selected);
end;

procedure TfrmUserGrp.DeleteUser(Sender: TObject);
begin
  if MessageDlg('確定要刪除嗎??', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    if trvUserGrp.Selected.AbsoluteIndex = 0 then UCMgr.DeleteUser(lsvUser.Selected)
    else UCMgr.DeleteUserFromUserGrp(trvUserGrp.Selected, lsvUser.Selected);
  end;
end;

procedure TfrmUserGrp.SelectUser(Sender: TObject);
begin
  ShowSelectUser(trvUserGrp.Selected, lsvUser.Items);
end;

procedure TfrmUserGrp.mnuUserGrpFuncClick(Sender: TObject);
begin
  ShowCheckFuncOfUserGrp(trvUserGrp.Selected);
end;

end.
