unit frmSysParamU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ImgList, DB, Grids, DBGrids, Buttons, StdCtrls, DBCtrls, Mask,
  ComCtrls, ExtCtrls ,DBClient;

type
  TfrmSysParam = class(TForm)
    Panel1: TPanel;
    lblTitle: TLabel;
    btnExit: TBitBtn;
    Panel2: TPanel;
    PageControl1: TPageControl;
    TabSheet2: TTabSheet;
    pnlMultiData: TPanel;
    dbgUser: TDBGrid;
    Panel3: TPanel;
    btnAppend: TBitBtn;
    btnEdit: TBitBtn;
    btnDelete: TBitBtn;
    btnCancel: TBitBtn;
    btnSave: TBitBtn;
    pnlSingleData: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    dbtID: TDBEdit;
    DBEdit4: TDBEdit;
    DBEdit5: TDBEdit;
    dsrParam: TDataSource;
    dsrUser: TDataSource;
    ImageList1: TImageList;
    dlgOpen: TOpenDialog;
    OpenDialog1: TOpenDialog;
    procedure btnAppendClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnParamCancelClick(Sender: TObject);
    procedure btnParamEditClick(Sender: TObject);
    procedure btnParamSaveClick(Sender: TObject);
    procedure spbChangeDirClick(Sender: TObject);
  private
    { Private declarations }
    procedure ChangeBtnStatus;
    procedure ChangeSysParamBtnStatus;
    function isUserDataOK : Boolean;
    function isParamDataOK : Boolean;

    procedure SaveCDS(I_CDS:TClientDataSet; sI_RefFileName:String);
    procedure setLanguageString;
  public
    { Public declarations }


    procedure preSaveDB;
  end;

var
  frmSysParam: TfrmSysParam;

implementation

uses dtmMainU, ConstU, frmMainU;

{$R *.dfm}

procedure TfrmSysParam.btnAppendClick(Sender : TObject);
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
    if MessageDlg(dtmMain.getLanguageSettingInfo('frmSysParam_Msg_5_Content'),mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      dtmMain.cdsUser.Delete;
      preSaveDB;
      ChangeBtnStatus;
    end;
end;

procedure TfrmSysParam.btnCancelClick(Sender: TObject);
begin
    dtmMain.cdsUser.Cancel;
    ChangeBtnStatus;
end;

procedure TfrmSysParam.btnSaveClick(Sender: TObject);
var
    bL_IsOnlyDta : Boolean;
    sL_UserId : String;
begin
    if isUserDataOK then
    begin
      if dsrUser.DataSet.State in [dsInsert] then
      begin
        sL_UserId := dsrUser.DataSet.FieldByName('SID').AsString;
        bL_IsOnlyDta := dtmMain.checkUserIdIsOnly(sL_UserId);
      end
      else
      begin
        bL_IsOnlyDta := true;
      end;



      if not bL_IsOnlyDta then
      begin
        dbtID.SetFocus;
        MessageDlg(dtmMain.getLanguageSettingInfo('frmSysParam_Msg_4_Content'),mtError, [mbOK],0);

        dtmMain.cdsUser.Delete;
        dtmMain.cdsUser.Append;
        preSaveDB;
      end
      else
      begin
        dtmMain.cdsUser.Post;
        preSaveDB;
        ChangeBtnStatus;
        dtmMain.getAllUserID;
      end;
    end;
end;

procedure TfrmSysParam.btnExitClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmSysParam.FormCreate(Sender: TObject);
begin
    ChangeSysParamBtnStatus;
    ChangeBtnStatus;

    setLanguageString;
end;

procedure TfrmSysParam.ChangeSysParamBtnStatus;
begin
{930301
     with dtmMain.cdsParam do
     begin
       if state in [dsInactive] then
         Exit
       else if state in [dsEdit, dsInsert] then
       begin
          btnParamCancel.Enabled := TRUE;
          btnParamSave.Enabled := TRUE;
          btnParamEdit.Enabled := FALSE;
          btnExit.Enabled := FALSE;
          pnlSysParam.Enabled := TRUE;

       end
       else
       begin
          if RecordCount>0 then
          begin
            btnParamEdit.Enabled := TRUE;
          end
          else
          begin
            btnParamEdit.Enabled := FALSE;
          end;

          btnParamCancel.Enabled := FALSE;
          btnParamSave.Enabled := FALSE;
          btnExit.Enabled := TRUE;
          pnlSysParam.Enabled := FALSE;
       end;
     end;
}
end;



function TfrmSysParam.isUserDataOK: Boolean;
var
    sL_UserId,sL_User,sL_Password : String;
begin
    sL_UserId := dsrUser.DataSet.FieldByName('SID').AsString;
    if sL_UserId = '' then
    begin
      dbtID.SetFocus;

      MessageDlg(dtmMain.getLanguageSettingInfo('frmSysParam_Msg_1_Content'),mtError, [mbOK],0);
      exit;
    end;

    sL_User := dsrUser.DataSet.FieldByName('SNAME').AsString;
    if sL_User = '' then
    begin
      DBEdit4.SetFocus;
      MessageDlg(dtmMain.getLanguageSettingInfo('frmSysParam_Msg_2_Content'),mtError, [mbOK],0);
      exit;
    end;

    sL_Password := dsrUser.DataSet.FieldByName('SPASSWORD').AsString;
    if sL_Password = '' then
    begin
      DBEdit5.SetFocus;
      MessageDlg(dtmMain.getLanguageSettingInfo('frmSysParam_Msg_3_Content'),mtError, [mbOK],0);
      exit;
    end;


    end;


procedure TfrmSysParam.btnParamCancelClick(Sender: TObject);
begin
{930301
    dtmMain.cdsParam.Cancel;
    ChangeSysParamBtnStatus;
}
end;

procedure TfrmSysParam.btnParamEditClick(Sender: TObject);
begin
{930301
    dtmMain.cdsParam.Edit;
    ChangeSysParamBtnStatus;
    dbtSysName.SetFocus;
}
end;

procedure TfrmSysParam.btnParamSaveClick(Sender: TObject);
begin
    if isParamDataOK then
    begin


      preSaveDB;
      ChangeSysParamBtnStatus;      
      MessageDlg(dtmMain.getLanguageSettingInfo('frmSysParam_Msg_6_Content'), mtConfirmation, [mbOK], 0);

    end;
end;

function TfrmSysParam.isParamDataOK: Boolean;
begin

end;

procedure TfrmSysParam.preSaveDB;
begin
{930301
    dsrParam.DataSet.FieldByName('sUptName').AsString := frmMain.sG_OperatorName;
    dsrParam.DataSet.FieldByName('dUptTime').AsDateTime := now;

    SaveCDS(dtmMain.cdsParam, DISPATCHER_SYS_PARAM_FILENAME);
}
    SaveCDS(dtmMain.cdsUser, USER_INFO_FILENAME);
end;

procedure TfrmSysParam.SaveCDS(I_CDS: TClientDataSet;
  sI_RefFileName: String);
begin
    SetCurrentDir(ExtractFilePath(Application.ExeName));
    if (I_CDS.State in [dsEdit, dsInsert]) then
      I_CDS.Post;

    if (I_CDS.ChangeCount>0) then
      I_CDS.SaveToFile(sI_RefFileName);
end;


procedure TfrmSysParam.spbChangeDirClick(Sender: TObject);
var
    sL_PathFileName,sL_Path : String;
begin
{930301
    if (dlgOpen.Execute) then
    begin
      with dlgOpen.Files do
      begin
        sL_PathFileName := dlgOpen.Files.Text;

        sL_Path := ExtractFileDir(sL_PathFileName);

        dtmMain.cdsParam.FieldByName('sLogPath').AsString := sL_Path;
      end;
    end;
}
end;

procedure TfrmSysParam.setLanguageString;
begin
    Caption := dtmMain.getLanguageSettingInfo('frmSysParam_Caption');
    lblTitle.Caption := dtmMain.getLanguageSettingInfo('frmSysParam_lblTitle_Caption');
    TabSheet2.Caption := dtmMain.getLanguageSettingInfo('frmSysParam_TabSheet2_Caption');
    btnAppend.Caption := dtmMain.getLanguageSettingInfo('frmSysParam_btnAppend_Caption');
    btnEdit.Caption := dtmMain.getLanguageSettingInfo('frmSysParam_btnEdit_Caption');
    btnDelete.Caption := dtmMain.getLanguageSettingInfo('frmSysParam_btnDelete_Caption');
    btnCancel.Caption := dtmMain.getLanguageSettingInfo('frmSysParam_btnCancel_Caption');
    btnSave.Caption := dtmMain.getLanguageSettingInfo('frmSysParam_btnSave_Caption');
    btnExit.Caption := dtmMain.getLanguageSettingInfo('frmSysParam_btnExit_Caption');
    Label1.Caption := dtmMain.getLanguageSettingInfo('frmSysParam_Label1_Caption');
    Label2.Caption := dtmMain.getLanguageSettingInfo('frmSysParam_Label2_Caption');
    Label3.Caption := dtmMain.getLanguageSettingInfo('frmSysParam_Label3_Caption');

    dbgUser.Columns[0].Title.Caption := dtmMain.getLanguageSettingInfo('frmSysParam_Label1_Caption');
    dbgUser.Columns[1].Title.Caption := dtmMain.getLanguageSettingInfo('frmSysParam_Label2_Caption');
end;

end.

