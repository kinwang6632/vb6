unit frmSysParamU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, DBClient, StdCtrls, Mask, DBCtrls, Buttons, ComCtrls, ExtCtrls,
  Grids, DBGrids, ImgList, Variants;


type
  TfrmSysParam = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    btnExit: TBitBtn;
    PageControl1: TPageControl;
    tab_1: TTabSheet;
    pnlMultiData: TPanel;
    Panel3: TPanel;
    pnlSingleData: TPanel;
    dbgUserInfo: TDBGrid;
    dsrUser: TDataSource;
    Lab_5: TLabel;
    dbtID: TDBEdit;
    Lab_6: TLabel;
    DBEdit4: TDBEdit;
    Lab_7: TLabel;
    DBEdit5: TDBEdit;
    btnAppend: TBitBtn;
    btnEdit: TBitBtn;
    btnDelete: TBitBtn;
    btnCancel: TBitBtn;
    btnSave: TBitBtn;
    ImageList1: TImageList;
    dlgOpen: TOpenDialog;
    procedure btnExitClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnAppendClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }

    procedure setLanguageString;
    procedure SaveCDS(I_CDS:TClientDataSet; sI_RefFileName:String);
    procedure ChangeBtnStatus;
    procedure preSaveDB;
  public
    { Public declarations }
    sG_LoginName : String;

  end;

var
  frmSysParam: TfrmSysParam;

implementation

uses ConstObjU, UdateTimeu, frmMainU, dtmMainU;


{$R *.DFM}

procedure TfrmSysParam.btnExitClick(Sender: TObject);
begin
    Close;
    {
    if ISDataOK then
      Close
    else
      MessageDlg('系統參數必須輸入',mtWarning,[mbOK],0);
    }        
end;

procedure TfrmSysParam.FormActivate(Sender: TObject);
begin
//    ActiveCDS(dtmMain.cdsUser, USER_INFO_FILENAME, 'B');  //B:browse mode; E:edit mode => 改在 dtmMain 中 active

    ChangeBtnStatus;

end;

procedure TfrmSysParam.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
    preSaveDB;

end;


procedure TfrmSysParam.btnAppendClick(Sender: TObject);
begin
    dtmMain.cdsUser.Append;
    ChangeBtnStatus;
end;

procedure TfrmSysParam.ChangeBtnStatus;
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
          if dtmMain.cdsUser.RecordCount>0 then
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

procedure TfrmSysParam.btnEditClick(Sender: TObject);
begin
    dtmMain.cdsUser.Edit;
    ChangeBtnStatus;

end;

procedure TfrmSysParam.btnDeleteClick(Sender: TObject);
begin
    if MessageDlg(frmMain.getLanguageSettingInfo('Sim_CA_Msg_26_Content'),mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      dtmMain.cdsUser.Delete;
      ChangeBtnStatus;

    end;
end;

procedure TfrmSysParam.btnCancelClick(Sender: TObject);
begin
      dtmMain.cdsUser.Cancel;
      ChangeBtnStatus;

end;

procedure TfrmSysParam.btnSaveClick(Sender: TObject);
begin
      dtmMain.cdsUser.Post;
      ChangeBtnStatus;

end;

procedure TfrmSysParam.SaveCDS(I_CDS: TClientDataSet; sI_RefFileName:String);
var
    sL_Path : String;
begin
//    SetCurrentDir(ExtractFilePath(Application.ExeName));



    if (I_CDS.State in [dsEdit, dsInsert]) then
      I_CDS.Post;

    if (I_CDS.ChangeCount>0) then
    begin
      sL_Path := ExtractFilePath(Application.ExeName);
      I_CDS.SaveToFile(sL_Path + sI_RefFileName);
    end;
end;





procedure TfrmSysParam.preSaveDB;
begin

    SaveCDS(dtmMain.cdsUser, USER_INFO_FILENAME);

end;



procedure TfrmSysParam.FormShow(Sender: TObject);
begin
    setLanguageString;
end;

procedure TfrmSysParam.setLanguageString;
begin
    Caption := frmMain.getLanguageSettingInfo('frmSimCA_3_Caption');

    tab_1.Caption := frmMain.getLanguageSettingInfo('Sim_CA_Tab_1_Caption');

    btnExit.Caption := frmMain.getLanguageSettingInfo('Sim_CA_btnExit_Caption');
    btnAppend.Caption := frmMain.getLanguageSettingInfo('Sim_CA_btnAppend_Caption');

    btnEdit.Caption := frmMain.getLanguageSettingInfo('Sim_CA_btnEdit_Caption');
    btnDelete.Caption := frmMain.getLanguageSettingInfo('Sim_CA_btnDelete_Caption');
    btnCancel.Caption := frmMain.getLanguageSettingInfo('Sim_CA_btnCancel_Caption');
    btnSave.Caption := frmMain.getLanguageSettingInfo('Sim_CA_btnSave_Caption');


    Lab_5.Caption := frmMain.getLanguageSettingInfo('Sim_CA_Lab_5_Caption');
    Lab_6.Caption := frmMain.getLanguageSettingInfo('Sim_CA_Lab_6_Caption');
    Lab_7.Caption := frmMain.getLanguageSettingInfo('Sim_CA_Lab_7_Caption');

    dbgUserInfo.Columns[0].Title.Caption := Lab_5.Caption;
    dbgUserInfo.Columns[1].Title.Caption := Lab_6.Caption;
    dbgUserInfo.Columns[2].Title.Caption := Lab_7.Caption;
end;

end.
