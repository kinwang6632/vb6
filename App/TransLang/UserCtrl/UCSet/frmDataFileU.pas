unit frmDataFileU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  utlUCU, ComCtrls, StdCtrls, Buttons, Menus;

type
  TfrmDataFile = class(TForm)
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    lsvDataFile: TListView;
    pmuDataFile: TPopupMenu;
    mnuNewDataFile: TMenuItem;
    mnuUpdateDataFile: TMenuItem;
    mnuDeleteDataFile: TMenuItem;
    N1: TMenuItem;
    mnuAllDataFile: TMenuItem;
    N2: TMenuItem;
    mnuSelect: TMenuItem;
    mnuUnSelect: TMenuItem;
    procedure lsvDataFileDblClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure pmuDataFilePopup(Sender: TObject);
    procedure mnuNewDataFileClick(Sender: TObject);
  private
    { Private declarations }
    FSubjectItem: TListItem;
  public
    { Public declarations }
    procedure Select(FuncItem: TListItem);
  end;

  procedure ShowSelectDataFile(FuncItem: TListItem);

var
  frmDataFile: TfrmDataFile;

implementation

uses frmDataFileAttrU;

{$R *.DFM}

procedure ShowSelectDataFile(FuncItem: TListItem);
begin
  frmDataFile := TfrmDataFile.Create(Application);
  frmDataFile.Select(FuncItem);
  frmDataFile.Free;
end;

{ TfrmDataFile }

procedure TfrmDataFile.Select(FuncItem: TListItem);
var
  ii: Integer;
  Func: TFunc;
  ListItem: TListItem;
begin
  FSubjectItem := FuncItem;
  Caption := '檔案套件設定';
  Func := nil;
  if Assigned(FuncItem) then
  begin
    lsvDataFile.CheckBoxes := True;
    Func := TFunc(FuncItem.Data);
  end;
  lsvDataFile.Items.BeginUpdate;
  for ii:=0 to UCMgr.DataFileCount-1 do
  begin
    ListItem := lsvDataFile.Items.Add;
    ListItem.Data := UCMgr.DataFiles[ii];
    UCMgr.DisplayDataFile(ListItem);
    if Assigned(Func) and Func.HasDataFile(UCMgr.DataFiles[ii]) then
      ListItem.Checked := True;
  end;
  lsvDataFile.Items.EndUpdate;  
  if (ShowModal = mrOk) and Assigned(FuncItem) then
  begin
    Func.DeleteAllDataFile;
    for ii:=0 to lsvDataFile.Items.Count-1 do
      if lsvDataFile.Items[ii].Checked then
        Func.AddDataFile(TDataFile(lsvDataFile.Items[ii].Data));
  end;
end;

procedure TfrmDataFile.lsvDataFileDblClick(Sender: TObject);
begin
  ModalResult := mrOk;
end;

procedure TfrmDataFile.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_ESCAPE then ModalResult := mrCancel;
end;

procedure TfrmDataFile.pmuDataFilePopup(Sender: TObject);
begin
  mnuUpdateDataFile.Enabled := False;
  mnuDeleteDataFile.Enabled := False;
  mnuSelect.Enabled := False;
  mnuUnSelect.Enabled := False;
  if Assigned(lsvDataFile.Selected) then
  begin
    mnuUpdateDataFile.Enabled := True;
    mnuDeleteDataFile.Enabled := True;
    mnuSelect.Enabled := False;
    mnuUnSelect.Enabled := False;
  end;
end;

procedure TfrmDataFile.mnuNewDataFileClick(Sender: TObject);
var
  DataFile: TDataFile;
  ListItem: TListItem;
begin
  DataFile := TDataFile.Create;
  if EditDataFileAttr(DataFile) = mrOk then
  begin
    ListItem := UCMgr.NewDataFile(DataFile, lsvDataFile.Items);
    if Assigned(FSubjectItem) then ListItem.Checked := True;
  end
  else DataFile.Free;
end;

end.
