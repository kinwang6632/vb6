unit frmSysParamU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, DBClient, StdCtrls, Mask, DBCtrls, Buttons, ComCtrls, ExtCtrls,
  Grids, DBGrids, ImgList, Variants;


type
  TfrmSysParam = class(TForm)
    cdsParam: TClientDataSet;
    dsrParam: TDataSource;
    Panel1: TPanel;
    Panel2: TPanel;
    btnExit: TBitBtn;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    cdsUser: TClientDataSet;
    pnlMultiData: TPanel;
    Panel3: TPanel;
    pnlSingleData: TPanel;
    DBGrid1: TDBGrid;
    dsrUser: TDataSource;
    cdsUsersID: TStringField;
    cdsUsersName: TStringField;
    cdsUsersPassword: TStringField;
    Label1: TLabel;
    dbtID: TDBEdit;
    Label2: TLabel;
    DBEdit4: TDBEdit;
    Label3: TLabel;
    DBEdit5: TDBEdit;
    btnAppend: TBitBtn;
    btnEdit: TBitBtn;
    btnDelete: TBitBtn;
    btnCancel: TBitBtn;
    btnSave: TBitBtn;
    ImageList1: TImageList;
    cdsParamsServerIP: TStringField;
    cdsParamsSysName: TStringField;
    cdsParamsSysVersion: TStringField;
    cdsParamsMisCommandPath: TStringField;
    dlgOpen: TOpenDialog;
    cdsParamnTimeOut: TIntegerField;
    cdsParamnMaxCommandCount: TIntegerField;
    cdsParambCommandLog: TBooleanField;
    cdsParambErrorLog: TBooleanField;
    cdsParamdUptTime: TDateTimeField;
    cdsParamsUptName: TStringField;
    Panel4: TPanel;
    btnParamEdit: TBitBtn;
    btnParamCancel: TBitBtn;
    btnParamSave: TBitBtn;
    pnlSysParam: TPanel;
    GroupBox2: TGroupBox;
    DBText1: TDBText;
    DBText2: TDBText;
    StaticText3: TStaticText;
    StaticText4: TStaticText;
    dbtSysName: TDBEdit;
    DBEdit6: TDBEdit;
    StaticText10: TStaticText;
    StaticText11: TStaticText;
    GroupBox1: TGroupBox;
    StaticText1: TStaticText;
    StaticText2: TStaticText;
    dbtIP: TDBEdit;
    dbtSPort: TDBEdit;
    StaticText6: TStaticText;
    dbtTimeout: TDBEdit;
    GroupBox4: TGroupBox;
    DBCheckBox3: TDBCheckBox;
    DBCheckBox1: TDBCheckBox;
    GroupBox3: TGroupBox;
    spbChangeDir: TSpeedButton;
    StaticText5: TStaticText;
    dbtLogPath: TDBEdit;
    Label4: TLabel;
    DBEdit1: TDBEdit;
    Label5: TLabel;
    DBEdit2: TDBEdit;
    Label6: TLabel;
    DBEdit3: TDBEdit;
    StaticText12: TStaticText;
    dbtRPort: TDBEdit;
    cdsParamnRPortNo: TIntegerField;
    cdsParamnSPortNo: TIntegerField;
    cdsParambSetZipCode: TBooleanField;
    cdsParamnCreditLimit: TIntegerField;
    cdsParamnSourceID: TIntegerField;
    cdsParamnDestID: TIntegerField;
    cdsParamnMopPPID: TIntegerField;
    cdsParamsBroadcastSDate: TStringField;
    cdsParamsBroadcastEDate: TStringField;
    cdsParamnResponseLog: TBooleanField;
    DBCheckBox4: TDBCheckBox;
    cdsParamnCmdRefreshRate: TIntegerField;
    cdsParamnCmdResentTimes: TIntegerField;
    cdsParambShowUI: TBooleanField;
    cdsParambAssignBroadcastDate: TBooleanField;
    cdsParamnCmdRefreshRate2: TIntegerField;
    GroupBox5: TGroupBox;
    DBCheckBox2: TDBCheckBox;
    StaticText9: TStaticText;
    DBEdit7: TDBEdit;
    DBCheckBox5: TDBCheckBox;
    DBCheckBox6: TDBCheckBox;
    StaticText13: TStaticText;
    dbtBroadcastSDate: TDBEdit;
    StaticText14: TStaticText;
    dbtBroadcastEDate: TDBEdit;
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
    cdsParamsSHotTime: TStringField;
    cdsParamsEHotTime: TStringField;
    dbtSHotTime: TDBEdit;
    dbtEHotTime: TDBEdit;
    procedure btnExitClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnAppendClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure spbChangeDirClick(Sender: TObject);
    procedure dbtLogPathExit(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure btnParamEditClick(Sender: TObject);
    procedure btnParamCancelClick(Sender: TObject);
    procedure btnParamSaveClick(Sender: TObject);
  private
    { Private declarations }
    function ISDataOK:boolean;
    procedure SaveCDS(I_CDS:TClientDataSet; sI_RefFileName:String);
    procedure ChangeBtnStatus;
    procedure ChangeSysParamBtnStatus;
    procedure preSaveDB;
  public
    { Public declarations }
    sG_LoginName : String;
    

  end;

var
  frmSysParam: TfrmSysParam;

implementation

uses  UdateTimeu, frmMainU, ConstU;


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
    frmMain.ActiveCDS(cdsParam, SYS_PARAM_FILENAME, 'B'); //B:browse mode; E:edit mode
    frmMain.ActiveCDS(cdsUser, USER_INFO_FILENAME, 'B');  //B:browse mode; E:edit mode

    ChangeBtnStatus;
    ChangeSysParamBtnStatus;

end;

procedure TfrmSysParam.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
    preSaveDB;

end;


procedure TfrmSysParam.btnAppendClick(Sender: TObject);
begin
    cdsUser.Append;
    ChangeBtnStatus;
end;

procedure TfrmSysParam.ChangeBtnStatus;
begin
     with cdsUser do   
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
          if cdsUser.RecordCount>0 then
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
    cdsUser.Edit;
    ChangeBtnStatus;

end;

procedure TfrmSysParam.btnDeleteClick(Sender: TObject);
begin
    if MessageDlg('是否確認刪除?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      cdsUser.Delete;
      ChangeBtnStatus;

    end;
end;

procedure TfrmSysParam.btnCancelClick(Sender: TObject);
begin
      cdsUser.Cancel;
      ChangeBtnStatus;

end;

procedure TfrmSysParam.btnSaveClick(Sender: TObject);
begin
      cdsUser.Post;
      ChangeBtnStatus;

end;

procedure TfrmSysParam.SaveCDS(I_CDS: TClientDataSet; sI_RefFileName:String);
begin
    SetCurrentDir(ExtractFilePath(Application.ExeName));
    if (I_CDS.State in [dsEdit, dsInsert]) then
      I_CDS.Post;
    
    if (I_CDS.ChangeCount>0) then
      I_CDS.SaveToFile(sI_RefFileName);

end;

procedure TfrmSysParam.spbChangeDirClick(Sender: TObject);
begin
    dlgOpen.InitialDir := GetCurrentDir;
    if (dlgOpen.Execute) then
    begin
      cdsParam.FieldByName('sLogPath').AsString := ExtractFilePath(dlgOpen.FileName);
    end;
end;

procedure TfrmSysParam.dbtLogPathExit(Sender: TObject);
begin
    if cdsParam.FieldByName('sLogPath').AsString='' then Exit;
    if (not (SetCurrentDir(cdsParam.FieldByName('sLogPath').AsString))) then
    begin                 
      MessageDlg('目錄不存在!!請重新點選.', mtInformation,[mbOk], 0);
      dbtLogPath.SetFocus;
    end;

end;

function TfrmSysParam.ISDataOK: boolean;
var
    bL_Result : boolean;
    nL_DateLength : Integer;
    sL_BroadcastStartDate, sL_BroadcastEndDate : STring;

begin

    bL_Result := true;
    if cdsParam.RecordCount = 0 then
    begin
      result := bL_Result;
      Exit;
    end;

    with cdsParam do
    begin
      nL_DateLength := 10;
      sL_BroadcastStartDate := FieldByName('sBroadcastSDate').AsString;
      sL_BroadcastEndDate := FieldByName('sBroadcastEDate').AsString;

      if (FieldByName('sSysName').Value=NULL) or (FieldByName('sSysName').Value='') then
      begin
        MessageDlg('請輸入系統名稱!', mtError, [mbOK],0);
        dbtSysName.SetFocus;
        Result := false;
        Exit;
      end;
      
      if (FieldByName('sSysVersion').Value=NULL) or (FieldByName('sSysVersion').Value='') then
      begin
        MessageDlg('請輸入系統版本!', mtError, [mbOK],0);
        DBEdit6.SetFocus;
        Result := false;
        Exit;
      end;


      if (FieldByName('sServerIP').Value=NULL) or (FieldByName('sServerIP').Value='') then
      begin
        MessageDlg('請輸入 IP!', mtError, [mbOK],0);
        dbtIP.SetFocus;
        Result := false;
        Exit;
      end;

      if (FieldByName('nSPortNo').Value=NULL) then
      begin
        MessageDlg('請輸入 S-Port!', mtError, [mbOK],0);
        dbtSPort.SetFocus;
        Result := false;
        Exit;
      end;
      if (FieldByName('nRPortNo').Value=NULL) then
      begin
        MessageDlg('請輸入 R-Port!', mtError, [mbOK],0);
        dbtRPort.SetFocus;
        Result := false;
        Exit;
      end;
      if (FieldByName('nTimeOut').Value=NULL) then
      begin
        MessageDlg('請輸入 Timeout!', mtError, [mbOK],0);
        dbtTimeout.SetFocus;
        Result := false;
        Exit;
      end;
      
      if (length(sL_BroadcastStartDate)<>nL_DateLength) then
      begin
        MessageDlg('日期格式錯誤!', mtError, [mbOK],0);
        dbtBroadcastSDate.SetFocus;
        Result := false;
        Exit;
      end;

      if (length(sL_BroadcastEndDate)<>nL_DateLength) then
      begin
        MessageDlg('日期格式錯誤!', mtError, [mbOK],0);
        dbtBroadcastEDate.SetFocus;
        Result := false;
        Exit;
      end;
      if not TUdateTime.IsDateStr(sL_BroadcastStartDate,'/') then
      begin
        MessageDlg('日期格式錯誤!', mtError, [mbOK],0);
        dbtBroadcastSDate.SetFocus;
        Result := false;
        Exit;
      end;

      if not TUdateTime.IsDateStr(sL_BroadcastEndDate,'/') then
      begin
        MessageDlg('日期格式錯誤!', mtError, [mbOK],0);
        dbtBroadcastEDate.SetFocus;
        Result := false;
        Exit;
      end;

      if (FieldByName('sLogPath').Value=NULL) or (FieldByName('sLogPath').Value='') then
      begin
        MessageDlg('請輸入 Log檔路徑!', mtError, [mbOK],0);
        dbtLogPath.SetFocus;
        Result := false;
        Exit;
      end;


      if (FieldByName('nMaxCommandCount').Value=NULL)then
      begin
        MessageDlg('請輸入可同時處理指令之數目!', mtError, [mbOK],0);
        dbtMaxCommandCount.SetFocus;
        Result := false;
        Exit;
      end;

      if (FieldByName('nMaxCommandCount').AsInteger>MAX_COMMAND_COUNT_PER_TIMES)then
      begin
        MessageDlg('每次最多只可同時處理 ' + IntToStr(MAX_COMMAND_COUNT_PER_TIMES) + ' 筆指令!', mtError, [mbOK],0);
        dbtMaxCommandCount.SetFocus;
        Result := false;
        Exit;
      end;

      if (FieldByName('nSourceID').Value=NULL)then
      begin
        MessageDlg('請輸入 Source ID!', mtError, [mbOK],0);
        DBEdit1.SetFocus;
        Result := false;
        Exit;
      end;

      if (FieldByName('nDestID').Value=NULL)then
      begin
        MessageDlg('請輸入 Dest ID!', mtError, [mbOK],0);
        DBEdit2.SetFocus;
        Result := false;
        Exit;
      end;

      if (FieldByName('nMopPPID').Value=NULL) then
      begin
        MessageDlg('請輸入 Mop PPID!', mtError, [mbOK],0);
        DBEdit3.SetFocus;
        Result := false;
        Exit;
      end;            
    end;
    Result := true;
end;

procedure TfrmSysParam.SpeedButton1Click(Sender: TObject);
begin
    dlgOpen.InitialDir := GetCurrentDir;
    if (dlgOpen.Execute) then
    begin
      cdsParam.FieldByName('sMisDbPath').AsString := ExtractFilePath(dlgOpen.FileName);
    end;

end;

procedure TfrmSysParam.btnParamEditClick(Sender: TObject);
begin
    cdsParam.Edit;
    ChangeSysParamBtnStatus;

end;

procedure TfrmSysParam.ChangeSysParamBtnStatus;
begin
     with cdsParam do
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
          dbtSysName.SetFocus;
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

procedure TfrmSysParam.btnParamCancelClick(Sender: TObject);
begin
      cdsParam.Cancel;
      ChangeSysParamBtnStatus;

end;

procedure TfrmSysParam.btnParamSaveClick(Sender: TObject);
begin

    if ISDataOK then
    begin
      cdsParam.FieldByName('sUptName').AsString := sG_LoginName;
      cdsParam.FieldByName('dUptTime').AsDateTime := now;
      if cdsParam.FieldByName('nCreditLimit').AsInteger<=0 then
         cdsParam.FieldByName('nCreditLimit').AsInteger := 0;
      cdsParam.Post;
      ChangeSysParamBtnStatus;
      preSaveDB;
      MessageDlg('更改設定後,必須重新啟動程式,才能使得新的設定值生效!', mtConfirmation, [mbOK], 0);
        
    end;
end;



procedure TfrmSysParam.preSaveDB;
begin
    SaveCDS(cdsParam, SYS_PARAM_FILENAME);
    SaveCDS(cdsUser, USER_INFO_FILENAME);
end;


end.
