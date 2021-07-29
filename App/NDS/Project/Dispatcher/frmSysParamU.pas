unit frmSysParamU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ImgList, DB, Grids, DBGrids, Buttons, StdCtrls, DBCtrls, Mask,
  ComCtrls, ExtCtrls ,DBClient;

type
  TfrmSysParam = class(TForm)
    Panel1: TPanel;
    Label7: TLabel;
    btnExit: TBitBtn;
    Panel2: TPanel;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    Panel4: TPanel;
    btnParamEdit: TBitBtn;
    btnParamCancel: TBitBtn;
    btnParamSave: TBitBtn;
    pnlSysParam: TPanel;
    TabSheet2: TTabSheet;
    pnlMultiData: TPanel;
    DBGrid1: TDBGrid;
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
    GroupBox5: TGroupBox;
    spbChangeDir: TSpeedButton;
    StaticText5: TStaticText;
    dbtLogPath: TDBEdit;
    GroupBox1: TGroupBox;
    DBText1: TDBText;
    DBText2: TDBText;
    StaticText3: TStaticText;
    StaticText4: TStaticText;
    dbtSysName: TDBEdit;
    DBEdit6: TDBEdit;
    StaticText10: TStaticText;
    StaticText11: TStaticText;
    GroupBox2: TGroupBox;
    StaticText6: TStaticText;
    dbtTimeout: TDBEdit;
    GroupBox3: TGroupBox;
    DBCheckBox1: TDBCheckBox;
    DBCheckBox4: TDBCheckBox;
    GroupBox4: TGroupBox;
    DBCheckBox5: TDBCheckBox;
    StaticText15: TStaticText;
    dbtCmdRefreshRate1: TDBEdit;
    StaticText16: TStaticText;
    StaticText7: TStaticText;
    dbtMaxCommandCount: TDBEdit;
    StaticText8: TStaticText;
    StaticText18: TStaticText;
    dbtCmdRefreshRate2: TDBEdit;
    StaticText19: TStaticText;
    StaticText17: TStaticText;
    dbtResentTimes: TDBEdit;
    StaticText20: TStaticText;
    StaticText21: TStaticText;
    dbtSHotTime: TDBEdit;
    dbtEHotTime: TDBEdit;
    OpenDialog1: TOpenDialog;
    procedure btnAppendClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
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
    if MessageDlg('是否確認刪除?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
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
        MessageDlg('此代碼已有人使用,請重新輸入',mtError, [mbOK],0);

        dtmMain.cdsUser.Delete;
        dtmMain.cdsUser.Append;
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


end;

procedure TfrmSysParam.ChangeSysParamBtnStatus;
begin
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
end;

procedure TfrmSysParam.FormShow(Sender: TObject);
var
    sL_FormTitle : String;
begin
    sL_FormTitle := 'NDS GateWay --[系統參數設定]';

    frmSysParam.Caption := sL_FormTitle;
end;

function TfrmSysParam.isUserDataOK: Boolean;
var
    sL_UserId,sL_User,sL_Password : String;
begin
    sL_UserId := dsrUser.DataSet.FieldByName('SID').AsString;
    if sL_UserId = '' then
    begin
      dbtID.SetFocus;
      MessageDlg('請輸入代碼',mtError, [mbOK],0);
      exit;
    end;

    sL_User := dsrUser.DataSet.FieldByName('SNAME').AsString;
    if sL_User = '' then
    begin
      DBEdit4.SetFocus;
      MessageDlg('請輸入姓名',mtError, [mbOK],0);
      exit;
    end;

    sL_Password := dsrUser.DataSet.FieldByName('SPASSWORD').AsString;
    if sL_Password = '' then
    begin
      DBEdit5.SetFocus;
      MessageDlg('請輸入密碼',mtError, [mbOK],0);
      exit;
    end;

end;


procedure TfrmSysParam.btnParamCancelClick(Sender: TObject);
begin
    dtmMain.cdsParam.Cancel;
    ChangeSysParamBtnStatus;
end;

procedure TfrmSysParam.btnParamEditClick(Sender: TObject);
begin
    dtmMain.cdsParam.Edit;
    ChangeSysParamBtnStatus;
    dbtSysName.SetFocus;
end;

procedure TfrmSysParam.btnParamSaveClick(Sender: TObject);
begin
    if isParamDataOK then
    begin


      preSaveDB;
      ChangeSysParamBtnStatus;      
      MessageDlg('更改設定後,必須重新啟動程式,才能使得新的設定值生效!', mtConfirmation, [mbOK], 0);

    end;
end;

function TfrmSysParam.isParamDataOK: Boolean;
begin

end;

procedure TfrmSysParam.preSaveDB;
begin
    dsrParam.DataSet.FieldByName('sUptName').AsString := frmMain.sG_OperatorName;
    dsrParam.DataSet.FieldByName('dUptTime').AsDateTime := now;

    SaveCDS(dtmMain.cdsParam, DISPATCHER_SYS_PARAM_FILENAME);
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

    if (dlgOpen.Execute) then
    begin
      with dlgOpen.Files do
      begin
        sL_PathFileName := dlgOpen.Files.Text;

        sL_Path := ExtractFileDir(sL_PathFileName);

        dtmMain.cdsParam.FieldByName('sLogPath').AsString := sL_Path;
      end;
    end;

end;

end.

