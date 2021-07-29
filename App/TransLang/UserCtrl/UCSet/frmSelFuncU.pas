unit frmSelFuncU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  utlUCU, ComCtrls, StdCtrls, Buttons;

type

  TShowStyle = (ssSelect, ssCheck);

  TfrmSelFunc = class(TForm)
    lsvFunc: TListView;
    btnOk: TBitBtn;
    btnCancel: TBitBtn;
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure lsvFuncDblClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    { Private declarations }
    FShowStyle: TShowStyle;
    FSubjectNode: TTreeNode;
  public
    { Public declarations }
    procedure Select(AppNode: TTreeNode;  ListItems: TListItems);
    procedure CheckFuncOfUserGrp(UserGrpNode: TTreeNode);
  end;

  procedure ShowSelectFunc(AppNode: TTreeNode; ListItems: TListItems);
  procedure ShowCheckFuncOfUserGrp(UserGrpNode: TTreeNode);

var
  frmSelFunc: TfrmSelFunc;

implementation

{$R *.DFM}

procedure ShowSelectFunc(AppNode: TTreeNode; ListItems: TListItems);
begin
  frmSelFunc := TfrmSelFunc.Create(Application);
  frmSelFunc.Select(AppNode, ListItems);
  frmSelFunc.Free;
end;

procedure ShowCheckFuncOfUserGrp(UserGrpNode: TTreeNode);
begin
  frmSelFunc := TfrmSelFunc.Create(Application);
  frmSelFunc.CheckFuncOfUserGrp(UserGrpNode);
  frmSelFunc.Free;
end;

{ TfrmSelFunc }

procedure TfrmSelFunc.Select(AppNode: TTreeNode; ListItems: TListItems);
var
  ii: Integer;
  ListItem: TListItem;
begin
  FShowStyle := ssSelect;
  FSubjectNode := AppNode;
  Caption := '新增功能選擇';
  lsvFunc.Items.BeginUpdate;
  for ii:=0 to UCMgr.FuncCount-1 do
  begin
    if TApp(AppNode.Data).HasFunc(UCMgr.Funcs[ii]) then Continue;
    ListItem := lsvFunc.Items.Add;
    ListItem.Data := UCMgr.Funcs[ii];
    UCMgr.DisplayFunc(ListItem);
  end;
  lsvFunc.Items.EndUpdate;
  if ShowModal = mrOk then
  begin
    for ii:=0 to lsvFunc.Items.Count-1 do
      if lsvFunc.Items[ii].Selected then
        UCMgr.AddFuncToApp(FSubjectNode,
          TFunc(lsvFunc.Items[ii].Data), ListItems);
  end;
end;

procedure TfrmSelFunc.CheckFuncOfUserGrp(UserGrpNode: TTreeNode);
var
  ii: Integer;
  UserGrp: TUserGrp;
  ListItem: TListItem;
begin
  FShowStyle := ssCheck;
  FSubjectNode := UserGrpNode;
  Caption := '群組權限設定';  
  lsvFunc.CheckBoxes := True;
  UserGrp := TUserGrp(UserGrpNode.Data);
  lsvFunc.Items.BeginUpdate;
  for ii:=0 to UCMgr.FuncCount-1 do
  begin
    ListItem := lsvFunc.Items.Add;
    ListItem.Data := UCMgr.Funcs[ii];
    UCMgr.DisplayFunc(ListItem);
    if UserGrp.HasFunc(UCMgr.Funcs[ii]) then
      ListItem.Checked := True;
  end;
  lsvFunc.Items.EndUpdate;  
  if ShowModal = mrOk then
  begin
    UserGrp.DeleteAllFunc;
    for ii:=0 to lsvFunc.Items.Count-1 do
      if lsvFunc.Items[ii].Checked then
        UCMgr.AddFuncToUserGrp(UserGrpNode, lsvFunc.Items[ii]);
  end;
end;

procedure TfrmSelFunc.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  if (ModalResult = mrOk) and (FShowStyle = ssSelect) and
    not Assigned(lsvFunc.Selected) then CanClose := False;
end;

procedure TfrmSelFunc.lsvFuncDblClick(Sender: TObject);
begin
  if FShowStyle = ssSelect then ModalResult := mrOk;
end;

procedure TfrmSelFunc.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_ESCAPE then ModalResult := mrCancel;
end;


end.
