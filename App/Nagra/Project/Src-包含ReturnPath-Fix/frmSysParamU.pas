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
    TabSheet3: TTabSheet;
    cdsEmm: TClientDataSet;
    cdsEmmsCommandType: TStringField;
    cdsEmmsBroadcastMode: TStringField;
    cdsEmmsAddressType: TStringField;
    Panel5: TPanel;
    btnEmmEdit: TBitBtn;
    btnEmmCancel: TBitBtn;
    btnEmmSave: TBitBtn;
    pnlEmmParam: TPanel;
    Label7: TLabel;
    dsrEmm: TDataSource;
    Label8: TLabel;
    dbtEmmCommandType: TDBEdit;
    Label9: TLabel;
    TabSheet4: TTabSheet;
    TabSheet5: TTabSheet;
    TabSheet6: TTabSheet;
    TabSheet7: TTabSheet;
    Panel6: TPanel;
    btnControlEdit: TBitBtn;
    btnControlCancel: TBitBtn;
    btnControlSave: TBitBtn;
    pnlControlParam: TPanel;
    cdsControl: TClientDataSet;
    cdsControlsCommandType: TStringField;
    dsrControl: TDataSource;
    Label10: TLabel;
    dbtControlCommandType: TDBEdit;
    cdsProduct: TClientDataSet;
    cdsFeedback: TClientDataSet;
    cdsOperation: TClientDataSet;
    Panel7: TPanel;
    btnProductEdit: TBitBtn;
    btnProductCancel: TBitBtn;
    btnProductSave: TBitBtn;
    pnlProductParam: TPanel;
    Label11: TLabel;
    dbtProductCommandType: TDBEdit;
    Panel9: TPanel;
    btnFeedbackEdit: TBitBtn;
    btnFeedbackCancel: TBitBtn;
    btnFeedbackSave: TBitBtn;
    pnlFeedbackParam: TPanel;
    Label12: TLabel;
    dbtFeedbackCommandType: TDBEdit;
    Panel11: TPanel;
    btnOperationEdit: TBitBtn;
    btnOperationCancel: TBitBtn;
    btnOperationSave: TBitBtn;
    pnlOperationParam: TPanel;
    Label13: TLabel;
    dbtOperationCommandType: TDBEdit;
    cdsProductsCommandType: TStringField;
    cdsFeedbacksCommandType: TStringField;
    cdsOperationsCommandType: TStringField;
    dsrProduct: TDataSource;
    dsrFeedback: TDataSource;
    dsrOperation: TDataSource;
    cdsControlsBroadcastMode: TStringField;
    cdsControlsAddressType: TStringField;
    Label14: TLabel;
    dbtBroadcastMode: TDBEdit;
    Label15: TLabel;
    dbtAddressType: TDBEdit;
    Label16: TLabel;
    Label17: TLabel;
    Label18: TLabel;
    Label19: TLabel;
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
    TabSheet8: TTabSheet;
    cdsNetworkID: TClientDataSet;
    cdsNetworkIDCompCode: TIntegerField;
    cdsNetworkIDCompName: TStringField;
    cdsNetworkIDNetworkID: TStringField;
    cdsNetworkIDOperation: TStringField;
    Panel8: TPanel;
    btnNetworkIDAppend: TBitBtn;
    btnNetworkIDEdit: TBitBtn;
    btnNetworkIDDelete: TBitBtn;
    btnNetworkIDCancel: TBitBtn;
    btnNetworkIDSave: TBitBtn;
    pnlSingleNetworkIdData: TPanel;
    Label22: TLabel;
    Label23: TLabel;
    Label24: TLabel;
    DBEdit9: TDBEdit;
    DBEdit10: TDBEdit;
    DBEdit11: TDBEdit;
    pnlMultiNetworkIdData: TPanel;
    DBGrid2: TDBGrid;
    Label25: TLabel;
    DBEdit12: TDBEdit;
    dsrNetworkId: TDataSource;
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
    DBEdit8: TDBEdit;
    DBEdit13: TDBEdit;
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
    procedure btnEmmSaveClick(Sender: TObject);
    procedure btnEmmEditClick(Sender: TObject);
    procedure btnEmmCancelClick(Sender: TObject);
    procedure btnControlSaveClick(Sender: TObject);
    procedure btnControlCancelClick(Sender: TObject);
    procedure btnControlEditClick(Sender: TObject);
    procedure btnProductSaveClick(Sender: TObject);
    procedure btnProductCancelClick(Sender: TObject);
    procedure btnProductEditClick(Sender: TObject);
    procedure btnFeedbackEditClick(Sender: TObject);
    procedure btnFeedbackCancelClick(Sender: TObject);
    procedure btnFeedbackSaveClick(Sender: TObject);
    procedure btnOperationSaveClick(Sender: TObject);
    procedure btnOperationEditClick(Sender: TObject);
    procedure btnOperationCancelClick(Sender: TObject);
    procedure btnNetworkIDEditClick(Sender: TObject);
    procedure btnNetworkIDCancelClick(Sender: TObject);
    procedure btnNetworkIDSaveClick(Sender: TObject);
    procedure btnNetworkIDAppendClick(Sender: TObject);
    procedure btnNetworkIDDeleteClick(Sender: TObject);
  private
    { Private declarations }
    function ISDataOK:boolean;
    procedure SaveCDS(I_CDS:TClientDataSet; sI_RefFileName:String);
    procedure ChangeBtnStatus;
    procedure ChangeSysParamBtnStatus;
    procedure ChangeEmmParamBtnStatus;
    procedure ChangeControlParamBtnStatus;
    procedure ChangeProductParamBtnStatus;
    procedure ChangeFeedbackParamBtnStatus;
    procedure ChangeOperationParamBtnStatus;
    procedure ChangeNetworkIDBtnStatus;    
    procedure preSaveDB;
  public
    { Public declarations }
    sG_LoginName : String;
    procedure getControlCommandSectionParam(var sI_BroadastMode, sI_AddressType:String);
  end;

var
  frmSysParam: TfrmSysParam;

implementation

uses ConstObjU, UdateTimeu, frmMainU;


{$R *.DFM}

procedure TfrmSysParam.btnExitClick(Sender: TObject);
begin
    Close;
    {
    if ISDataOK then
      Close
    else
      MessageDlg('?t?????????????J',mtWarning,[mbOK],0);
    }        
end;

procedure TfrmSysParam.FormActivate(Sender: TObject);
begin
    frmMain.ActiveCDS(cdsParam, SYS_PARAM_FILENAME, 'B'); //B:browse mode; E:edit mode
    frmMain.ActiveCDS(cdsUser, USER_INFO_FILENAME, 'B');  //B:browse mode; E:edit mode
    frmMain.ActiveCDS(cdsEmm, EMM_INFO_FILENAME, 'B');  //B:browse mode; E:edit mode
    frmMain.ActiveCDS(cdsControl, CONTROL_INFO_FILENAME, 'B');  //B:browse mode; E:edit mode

    frmMain.ActiveCDS(cdsProduct, PRODUCT_INFO_FILENAME, 'B');  //B:browse mode; E:edit mode
    frmMain.ActiveCDS(cdsFeedback, FEEDBACK_INFO_FILENAME, 'B');  //B:browse mode; E:edit mode
    frmMain.ActiveCDS(cdsOperation, OPERATION_INFO_FILENAME, 'B');  //B:browse mode; E:edit mode
    frmMain.ActiveCDS(cdsNetworkID, NETWORK_ID_INFO_FILENAME, 'B');  //B:browse mode; E:edit mode

    ChangeBtnStatus;
    ChangeSysParamBtnStatus;
    ChangeEmmParamBtnStatus;
    ChangeControlParamBtnStatus;
    ChangeProductParamBtnStatus;
    ChangeFeedbackParamBtnStatus;
    ChangeOperationParamBtnStatus;

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
    if MessageDlg('?O?_?T?{?R???',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
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
      MessageDlg('???????s?b!!?????s?I??.', mtInformation,[mbOk], 0);
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
        MessageDlg('?????J?t???W??!', mtError, [mbOK],0);
        dbtSysName.SetFocus;
        Result := false;
        Exit;
      end;
      
      if (FieldByName('sSysVersion').Value=NULL) or (FieldByName('sSysVersion').Value='') then
      begin
        MessageDlg('?????J?t??????!', mtError, [mbOK],0);
        DBEdit6.SetFocus;
        Result := false;
        Exit;
      end;


      if (FieldByName('sServerIP').Value=NULL) or (FieldByName('sServerIP').Value='') then
      begin
        MessageDlg('?????J IP!', mtError, [mbOK],0);
        dbtIP.SetFocus;
        Result := false;
        Exit;
      end;

      if (FieldByName('nSPortNo').Value=NULL) then
      begin
        MessageDlg('?????J S-Port!', mtError, [mbOK],0);
        dbtSPort.SetFocus;
        Result := false;
        Exit;
      end;
      if (FieldByName('nRPortNo').Value=NULL) then
      begin
        MessageDlg('?????J R-Port!', mtError, [mbOK],0);
        dbtRPort.SetFocus;
        Result := false;
        Exit;
      end;
      if (FieldByName('nTimeOut').Value=NULL) then
      begin
        MessageDlg('?????J Timeout!', mtError, [mbOK],0);
        dbtTimeout.SetFocus;
        Result := false;
        Exit;
      end;
      
      if (length(sL_BroadcastStartDate)<>nL_DateLength) then
      begin
        MessageDlg('???????????~!', mtError, [mbOK],0);
        dbtBroadcastSDate.SetFocus;
        Result := false;
        Exit;
      end;

      if (length(sL_BroadcastEndDate)<>nL_DateLength) then
      begin
        MessageDlg('???????????~!', mtError, [mbOK],0);
        dbtBroadcastEDate.SetFocus;
        Result := false;
        Exit;
      end;
      if not TUdateTime.IsDateStr(sL_BroadcastStartDate,'/') then
      begin
        MessageDlg('???????????~!', mtError, [mbOK],0);
        dbtBroadcastSDate.SetFocus;
        Result := false;
        Exit;
      end;

      if not TUdateTime.IsDateStr(sL_BroadcastEndDate,'/') then
      begin
        MessageDlg('???????????~!', mtError, [mbOK],0);
        dbtBroadcastEDate.SetFocus;
        Result := false;
        Exit;
      end;

      if (FieldByName('sLogPath').Value=NULL) or (FieldByName('sLogPath').Value='') then
      begin
        MessageDlg('?????J Log?????|!', mtError, [mbOK],0);
        dbtLogPath.SetFocus;
        Result := false;
        Exit;
      end;


      if (FieldByName('nMaxCommandCount').Value=NULL)then
      begin
        MessageDlg('?????J?i?P???B?z???O??????!', mtError, [mbOK],0);
        dbtMaxCommandCount.SetFocus;
        Result := false;
        Exit;
      end;

      if (FieldByName('nMaxCommandCount').AsInteger>MAX_COMMAND_COUNT_PER_TIMES)then
      begin
        MessageDlg('?C?????h?u?i?P???B?z ' + IntToStr(MAX_COMMAND_COUNT_PER_TIMES) + ' ?????O!', mtError, [mbOK],0);
        dbtMaxCommandCount.SetFocus;
        Result := false;
        Exit;
      end;

      if (FieldByName('nSourceID').Value=NULL)then
      begin
        MessageDlg('?????J Source ID!', mtError, [mbOK],0);
        DBEdit1.SetFocus;
        Result := false;
        Exit;
      end;

      if (FieldByName('nDestID').Value=NULL)then
      begin
        MessageDlg('?????J Dest ID!', mtError, [mbOK],0);
        DBEdit2.SetFocus;
        Result := false;
        Exit;
      end;

      if (FieldByName('nMopPPID').Value=NULL) then
      begin
        MessageDlg('?????J Mop PPID!', mtError, [mbOK],0);
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
      MessageDlg('?????]?w??,???????s?????{??,?~?????o?s???]?w??????!', mtConfirmation, [mbOK], 0);
        
    end;
end;

procedure TfrmSysParam.btnEmmSaveClick(Sender: TObject);
begin

      cdsEmm.Post;
      ChangeEmmParamBtnStatus;

end;

procedure TfrmSysParam.ChangeEmmParamBtnStatus;
begin
     with cdsEmm do
     begin
       if state in [dsInactive] then
         Exit
       else if state in [dsEdit, dsInsert] then
       begin
          btnEmmCancel.Enabled := TRUE;
          btnEmmSave.Enabled := TRUE;
          btnEmmEdit.Enabled := FALSE;
          btnExit.Enabled := FALSE;
          pnlEmmParam.Enabled := TRUE;
          dbtEmmCommandType.SetFocus;
       end
       else
       begin
          if RecordCount>0 then
          begin
            btnEmmEdit.Enabled := TRUE;
          end
          else
          begin
            btnEmmEdit.Enabled := FALSE;
          end;
          btnEmmCancel.Enabled := FALSE;
          btnEmmSave.Enabled := FALSE;
          btnExit.Enabled := TRUE;
          pnlEmmParam.Enabled := FALSE;
       end;
     end;

end;

procedure TfrmSysParam.btnEmmEditClick(Sender: TObject);
begin
    cdsEmm.Edit;
    ChangeEmmParamBtnStatus;

end;

procedure TfrmSysParam.btnEmmCancelClick(Sender: TObject);
begin
      cdsEmm.Cancel;
      ChangeEmmParamBtnStatus;

end;

procedure TfrmSysParam.btnControlSaveClick(Sender: TObject);
begin
    if (Uppercase(cdsControl.FieldByName('sBroadcastMode').AsString) <> 'N') and
       (Uppercase(cdsControl.FieldByName('sBroadcastMode').AsString) <> 'B') then
    begin
      MessageDlg('?????s?]?w broadcast mode.(???????? N or B)', mtError, [mbOK],0);
      dbtBroadcastMode.SetFocus;
      exit;
    end;
    if (Uppercase(cdsControl.FieldByName('sAddressType').AsString) <> 'U') and
       (Uppercase(cdsControl.FieldByName('sAddressType').AsString) <> 'G') then
    begin
      MessageDlg('?????s?]?w address type.(???????? U or G)', mtError, [mbOK],0);
      dbtAddressType.SetFocus;
      exit;
    end;
    cdsControl.FieldByName('sAddressType').AsString := UpperCase(cdsControl.FieldByName('sAddressType').AsString);
    cdsControl.FieldByName('sBroadcastMode').AsString := UpperCase(cdsControl.FieldByName('sBroadcastMode').AsString);    
    cdsControl.Post;
    ChangeControlParamBtnStatus;

end;

procedure TfrmSysParam.btnControlCancelClick(Sender: TObject);
begin
      cdsControl.Cancel;
      ChangeControlParamBtnStatus;

end;

procedure TfrmSysParam.btnControlEditClick(Sender: TObject);
begin
    cdsControl.Edit;
    ChangeControlParamBtnStatus;

end;

procedure TfrmSysParam.ChangeControlParamBtnStatus;
begin
     with cdsControl do
     begin
       if state in [dsInactive] then
         Exit
       else if state in [dsEdit, dsInsert] then
       begin
          btnControlCancel.Enabled := TRUE;
          btnControlSave.Enabled := TRUE;
          btnControlEdit.Enabled := FALSE;
          btnExit.Enabled := FALSE;
          pnlControlParam.Enabled := TRUE;
          dbtControlCommandType.SetFocus;
       end
       else
       begin
          if RecordCount>0 then
          begin
            btnControlEdit.Enabled := TRUE;
          end
          else
          begin
            btnControlEdit.Enabled := FALSE;
          end;
          btnControlCancel.Enabled := FALSE;
          btnControlSave.Enabled := FALSE;
          btnExit.Enabled := TRUE;
          pnlControlParam.Enabled := FALSE;
       end;
     end;

end;

procedure TfrmSysParam.ChangeFeedbackParamBtnStatus;
begin
     with cdsFeedback do
     begin
       if state in [dsInactive] then
         Exit
       else if state in [dsEdit, dsInsert] then
       begin
          btnFeedbackCancel.Enabled := TRUE;
          btnFeedbackSave.Enabled := TRUE;
          btnFeedbackEdit.Enabled := FALSE;
          btnExit.Enabled := FALSE;
          pnlFeedbackParam.Enabled := TRUE;
          dbtFeedbackCommandType.SetFocus;
       end
       else
       begin
          if RecordCount>0 then
          begin
            btnFeedbackEdit.Enabled := TRUE;
          end
          else
          begin
            btnFeedbackEdit.Enabled := FALSE;
          end;
          btnFeedbackCancel.Enabled := FALSE;
          btnFeedbackSave.Enabled := FALSE;
          btnExit.Enabled := TRUE;
          pnlFeedbackParam.Enabled := FALSE;
       end;
     end;

end;

procedure TfrmSysParam.ChangeOperationParamBtnStatus;
begin
     with cdsOperation do
     begin
       if state in [dsInactive] then
         Exit
       else if state in [dsEdit, dsInsert] then
       begin
          btnOperationCancel.Enabled := TRUE;
          btnOperationSave.Enabled := TRUE;
          btnOperationEdit.Enabled := FALSE;
          btnExit.Enabled := FALSE;
          pnlOperationParam.Enabled := TRUE;
          dbtOperationCommandType.SetFocus;
       end
       else
       begin
          if RecordCount>0 then
          begin
            btnOperationEdit.Enabled := TRUE;
          end
          else
          begin
            btnOperationEdit.Enabled := FALSE;
          end;
          btnOperationCancel.Enabled := FALSE;
          btnOperationSave.Enabled := FALSE;
          btnExit.Enabled := TRUE;
          pnlOperationParam.Enabled := FALSE;
       end;
     end;

end;

procedure TfrmSysParam.ChangeProductParamBtnStatus;
begin
     with cdsProduct do
     begin
       if state in [dsInactive] then
         Exit
       else if state in [dsEdit, dsInsert] then
       begin
          btnProductCancel.Enabled := TRUE;
          btnProductSave.Enabled := TRUE;
          btnProductEdit.Enabled := FALSE;
          btnExit.Enabled := FALSE;
          pnlProductParam.Enabled := TRUE;
          dbtProductCommandType.SetFocus;
       end
       else
       begin
          if RecordCount>0 then
          begin
            btnProductEdit.Enabled := TRUE;
          end
          else
          begin
            btnProductEdit.Enabled := FALSE;
          end;
          btnProductCancel.Enabled := FALSE;
          btnProductSave.Enabled := FALSE;
          btnExit.Enabled := TRUE;
          pnlProductParam.Enabled := FALSE;
       end;
     end;

end;

procedure TfrmSysParam.btnProductSaveClick(Sender: TObject);
begin
      cdsProduct.Post;
      ChangeProductParamBtnStatus;

end;

procedure TfrmSysParam.btnProductCancelClick(Sender: TObject);
begin
      cdsProduct.Cancel;
      ChangeProductParamBtnStatus;

end;

procedure TfrmSysParam.btnProductEditClick(Sender: TObject);
begin
      cdsProduct.Edit;
      ChangeProductParamBtnStatus;

end;

procedure TfrmSysParam.btnFeedbackEditClick(Sender: TObject);
begin
      cdsFeedback.Edit;
      ChangeFeedbackParamBtnStatus;

end;

procedure TfrmSysParam.btnFeedbackCancelClick(Sender: TObject);
begin
      cdsFeedback.Cancel;
      ChangeFeedbackParamBtnStatus;

end;

procedure TfrmSysParam.btnFeedbackSaveClick(Sender: TObject);
begin
      cdsFeedback.Post;
      ChangeFeedbackParamBtnStatus;

end;

procedure TfrmSysParam.btnOperationSaveClick(Sender: TObject);
begin
      cdsOperation.Post;
      ChangeOperationParamBtnStatus;

end;

procedure TfrmSysParam.btnOperationEditClick(Sender: TObject);
begin
      cdsOperation.Edit;
      ChangeOperationParamBtnStatus;

end;

procedure TfrmSysParam.btnOperationCancelClick(Sender: TObject);
begin
      cdsOperation.Cancel;
      ChangeOperationParamBtnStatus;

end;

procedure TfrmSysParam.getControlCommandSectionParam(var sI_BroadastMode,
  sI_AddressType: String);
begin
    frmMain.ActiveCDS(cdsControl, CONTROL_INFO_FILENAME, 'B');  //B:browse mode; E:edit mode
    sI_BroadastMode := cdsControl.fieldByName('sBroadcastMode').AsString;
    sI_AddressType := cdsControl.fieldByName('sAddressType').AsString;
    cdsControl.Close;    
end;

procedure TfrmSysParam.preSaveDB;
begin
    SaveCDS(cdsParam, SYS_PARAM_FILENAME);
    SaveCDS(cdsUser, USER_INFO_FILENAME);
    SaveCDS(cdsEmm, EMM_INFO_FILENAME);
    SaveCDS(cdsControl, CONTROL_INFO_FILENAME);
    SaveCDS(cdsProduct, PRODUCT_INFO_FILENAME);
    SaveCDS(cdsFeedback, FEEDBACK_INFO_FILENAME);
    SaveCDS(cdsOperation, OPERATION_INFO_FILENAME);
    SaveCDS(cdsNetworkID, NETWORK_ID_INFO_FILENAME);
end;

procedure TfrmSysParam.ChangeNetworkIDBtnStatus;
begin
     with cdsNetworkID do
     begin
       if state in [dsInactive] then
         Exit     
       else if state in [dsEdit, dsInsert] then
       begin
          btnNetworkIDCancel.Enabled := TRUE;
          btnNetworkIDSave.Enabled := TRUE;
          btnNetworkIDDelete.Enabled := FALSE;
          btnNetworkIDEdit.Enabled := FALSE;
          btnExit.Enabled := FALSE;
          btnNetworkIDAppend.Enabled := FALSE;
          pnlSingleData.Enabled := TRUE;
          pnlMultiData.Enabled := FALSE;
          DBEdit9.SetFocus;
       end
       else
       begin
          if cdsUser.RecordCount>0 then
          begin
            btnNetworkIDEdit.Enabled := TRUE;
            btnNetworkIDDelete.Enabled := TRUE;
          end
          else
          begin
            btnNetworkIDEdit.Enabled := FALSE;
            btnNetworkIDDelete.Enabled := FALSE;
          end;
          btnNetworkIDCancel.Enabled := FALSE;
          btnNetworkIDSave.Enabled := FALSE;
          btnNetworkIDAppend.Enabled := TRUE;
          btnExit.Enabled := TRUE;          
          pnlSingleData.Enabled := FALSE;
          pnlMultiData.Enabled := TRUE;
       end;
     end;


end;

procedure TfrmSysParam.btnNetworkIDEditClick(Sender: TObject);
begin
      cdsNetworkID.Edit;
      ChangeNetworkIDBtnStatus;

end;

procedure TfrmSysParam.btnNetworkIDCancelClick(Sender: TObject);
begin
      cdsNetworkID.Cancel;
      ChangeNetworkIDBtnStatus;

end;

procedure TfrmSysParam.btnNetworkIDSaveClick(Sender: TObject);
begin
      cdsNetworkID.Post;
      ChangeNetworkIDBtnStatus;

end;

procedure TfrmSysParam.btnNetworkIDAppendClick(Sender: TObject);
begin
      cdsNetworkID.Append;
      ChangeNetworkIDBtnStatus;

end;

procedure TfrmSysParam.btnNetworkIDDeleteClick(Sender: TObject);
begin
    if MessageDlg('?O?_?T?{?R???',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      cdsNetworkID.Delete;
      ChangeNetworkIDBtnStatus;

    end;

end;

end.
