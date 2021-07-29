unit frmAppU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, ExtCtrls, ToolWin, Menus, IniFiles, utlUCU;

type
  TfrmApp = class(TForm)
    trvApp: TTreeView;
    Splitter1: TSplitter;
    lsvFunc: TListView;
    pmuApp: TPopupMenu;
    mnuNewApp: TMenuItem;
    mnuUpdateApp: TMenuItem;
    mnuDeleteApp: TMenuItem;
    N1: TMenuItem;
    mnuUpApp: TMenuItem;
    mnuDownApp: TMenuItem;
    pmuFunc: TPopupMenu;
    mnuNewFunc: TMenuItem;
    mnuUpdateFunc: TMenuItem;
    mnuDeleteFunc: TMenuItem;
    MenuItem4: TMenuItem;
    mnuUpFunc: TMenuItem;
    mnuDownFunc: TMenuItem;
    N2: TMenuItem;
    mnuFuncDataFile: TMenuItem;
    procedure trvAppChange(Sender: TObject; Node: TTreeNode);
    procedure pmuAppPopup(Sender: TObject);
    procedure pmuFuncPopup(Sender: TObject);
    procedure NewApp(Sender: TObject);
    procedure UpdateApp(Sender: TObject);
    procedure UpdateFunc(Sender: TObject);
    procedure DeleteApp(Sender: TObject);
    procedure DeleteFunc(Sender: TObject);
    procedure mnuFuncDataFileClick(Sender: TObject);
  private
    { Private declarations }
    procedure NewFunc(Sender: TObject);
    procedure SelectFunc(Sender: TObject);
  public
    { Public declarations }
  end;

var
  frmApp: TfrmApp;

implementation

uses frmAppAttrU, frmFuncAttrU, frmSelFuncU, frmDataFileU;

{$R *.DFM}

procedure TfrmApp.trvAppChange(Sender: TObject; Node: TTreeNode);
var
  ii: Integer;
  App: TApp;
  ListItem: TListItem;
begin
  lsvFunc.Items.BeginUpdate;
  lsvFunc.Items.Clear;
  App := TApp(Node.Data);
  for ii:=0 to App.FuncCount-1 do
  begin
    ListItem := lsvFunc.Items.Add;
    ListItem.Data := App.Funcs[ii];
    UCMgr.DisplayFunc(ListItem);
  end;
  lsvFunc.Items.EndUpdate;  
end;

procedure TfrmApp.pmuAppPopup(Sender: TObject);
begin
  trvApp.Selected := trvApp.Selected;
  mnuNewApp.Enabled := False;
  mnuUpdateApp.Enabled := False;
  mnuDeleteApp.Enabled := False;
  mnuDownApp.Enabled := False;
  mnuUpApp.Enabled := False;
  if not Assigned(trvApp.Selected) then Exit;
  mnuNewApp.Enabled := True;
  if trvApp.Selected.AbsoluteIndex <> 0 then
  begin
    mnuUpdateApp.Enabled := True;
    mnuDeleteApp.Enabled := True;
    mnuDownApp.Enabled := True;
    mnuUpApp.Enabled := True;
  end;
end;

procedure TfrmApp.pmuFuncPopup(Sender: TObject);
begin
  mnuNewFunc.Enabled := False;
  mnuUpdateFunc.Enabled := False;
  mnuDeleteFunc.Enabled := False;
  mnuDownFunc.Enabled := False;
  mnuUpFunc.Enabled := False;
  mnuFuncDataFile.Enabled := False;
  if not Assigned(trvApp.Selected) then Exit;
  if Assigned(lsvFunc.Selected) then
  begin
    mnuUpdateFunc.Enabled := True;
    mnuDeleteFunc.Enabled := True;
    mnuDownFunc.Enabled := True;
    mnuUpFunc.Enabled := True;
  end;
  mnuNewFunc.Enabled := True;
  mnuFuncDataFile.Enabled := True;
  mnuNewFunc.OnClick := SelectFunc;
  if (trvApp.Selected.AbsoluteIndex = 0) then mnuNewFunc.OnClick := NewFunc;
end;

procedure TfrmApp.NewApp(Sender: TObject);
var
  App: TApp;
begin
  App := TApp.Create;
  if EditAppAttr(App) = mrOk then UCMgr.NewApp(App, trvApp.Selected)
  else App.Free;
end;

procedure TfrmApp.UpdateApp(Sender: TObject);
var
  App: TApp;
begin
  App := TApp(trvApp.Selected.Data);
  if EditAppAttr(App) = mrOk then  trvApp.Selected.Text := App.Desc;
end;

procedure TfrmApp.DeleteApp(Sender: TObject);
begin
  if MessageDlg('確定要刪除嗎??', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    UCMgr.DeleteApp(trvApp.Selected);
end;

procedure TfrmApp.NewFunc(Sender: TObject);
var
  Func: TFunc;
begin
  Func := TFunc.Create;
  if EditFuncAttr(Func) = mrOk then UCMgr.NewFunc(Func, lsvFunc.Items)
  else Func.Free;
end;

procedure TfrmApp.UpdateFunc(Sender: TObject);
var
  Func: TFunc;
begin
  Func := TFunc(lsvFunc.Selected.Data);
  if EditFuncAttr(Func) = mrOk then UCMgr.DisplayFunc(lsvFunc.Selected);
end;

procedure TfrmApp.DeleteFunc(Sender: TObject);
begin
  if MessageDlg('確定要刪除嗎??', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    if trvApp.Selected.AbsoluteIndex = 0 then UCMgr.DeleteFunc(lsvFunc.Selected)
    else UCMgr.DeleteFuncFromApp(trvApp.Selected, lsvFunc.Selected);
  end;
end;

procedure TfrmApp.SelectFunc(Sender: TObject);
begin
  ShowSelectFunc(trvApp.Selected, lsvFunc.Items);
end;

procedure TfrmApp.mnuFuncDataFileClick(Sender: TObject);
begin
  ShowSelectDataFile(lsvFunc.Selected);
end;

end.
