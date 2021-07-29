unit frmUserU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, DBGrids, Buttons, ExtCtrls, Mask, DBCtrls, DB,
  DBClient;

type
  TfrmUser = class(TForm)
    Panel2: TPanel;
    pnlSingleData: TPanel;
    StaticText1: TStaticText;
    StaticText2: TStaticText;
    StaticText3: TStaticText;
    dbtID: TDBEdit;
    DBEdit4: TDBEdit;
    DBEdit5: TDBEdit;
    btnAppend: TBitBtn;
    btnEdit: TBitBtn;
    btnDelete: TBitBtn;
    btnCancel: TBitBtn;
    btnSave: TBitBtn;
    btnExit: TBitBtn;
    pnlMultiData: TPanel;
    dbgUser: TDBGrid;
    dsrUser: TDataSource;
    procedure ChangeBtnStatus;    
    procedure btnExitClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure btnAppendClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmUser: TfrmUser;

implementation

uses dtmMainU, ConstObjU, frmMainU;

{$R *.dfm}

procedure TfrmUser.btnExitClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmUser.btnSaveClick(Sender: TObject);
begin
    dtmMain.cdsUser.Post;
    ChangeBtnStatus;
    dtmMain.SaveCDS(dtmMain.cdsUser, USER_INFO_FILENAME);
end;

procedure TfrmUser.btnCancelClick(Sender: TObject);
begin
      dtmMain.cdsUser.Cancel;
      ChangeBtnStatus;

end;

procedure TfrmUser.btnDeleteClick(Sender: TObject);
begin
    if MessageDlg('是否確認刪除?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      dtmMain.cdsUser.Delete;
      ChangeBtnStatus;

    end;

end;

procedure TfrmUser.btnEditClick(Sender: TObject);
begin
    dtmMain.cdsUser.Edit;
    ChangeBtnStatus;

end;

procedure TfrmUser.btnAppendClick(Sender: TObject);
begin
    dtmMain.cdsUser.Append;
    ChangeBtnStatus;

end;

procedure TfrmUser.ChangeBtnStatus;
begin
     with dtmMain.cdsUser do
     begin
       if state in [dsInactive] then
         Exit     
       else if state in [dsEdit, dsInsert] then
       begin
          btnCancel.Enabled := TRUE;
          btnSave.Enabled := TRUE;
          btnDelete.Enabled := FALSE;
          btnEdit.Enabled := FALSE;
          btnExit.Enabled := FALSE;
          btnAppend.Enabled := FALSE;
          pnlSingleData.Enabled := TRUE;
          pnlMultiData.Enabled := FALSE;
          dbtID.SetFocus;
       end
       else
       begin
          if RecordCount>0 then
          begin
            btnEdit.Enabled := TRUE;
            btnDelete.Enabled := TRUE;
          end
          else
          begin
            btnEdit.Enabled := FALSE;
            btnDelete.Enabled := FALSE;
          end;
          btnCancel.Enabled := FALSE;
          btnSave.Enabled := FALSE;
          btnAppend.Enabled := TRUE;
          btnExit.Enabled := TRUE;          
          pnlSingleData.Enabled := FALSE;
          pnlMultiData.Enabled := TRUE;
       end;
     end;

end;

procedure TfrmUser.FormShow(Sender: TObject);
begin
    Caption := frmMain.getFrmCaption('使用者資料維護');
    ChangeBtnStatus;
end;

end.
