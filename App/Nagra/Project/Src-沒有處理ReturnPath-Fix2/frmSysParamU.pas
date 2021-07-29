unit frmSysParamU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, DBClient, StdCtrls, Mask, DBCtrls, Buttons, ComCtrls, ExtCtrls,
  Grids, DBGrids, ImgList, Variants,  dtmMainU;


type
  TfrmSysParam = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    btnExit: TBitBtn;
    PageControl1: TPageControl;
    frmSysParam_Tab_1: TTabSheet;
    frmSysParam_Tab_2: TTabSheet;
    pnlMultiData: TPanel;
    Panel3: TPanel;
    pnlSingleData: TPanel;
    dbgUser: TDBGrid;
    dsrUser: TDataSource;
    frmSysParam_Lbl_14: TLabel;
    dbtID: TDBEdit;
    frmSysParam_Lbl_15: TLabel;
    DBEdit4: TDBEdit;
    frmSysParam_Lbl_16: TLabel;
    DBEdit5: TDBEdit;
    btnAppend: TBitBtn;
    btnEdit: TBitBtn;
    btnDelete: TBitBtn;
    btnCancel: TBitBtn;
    btnSave: TBitBtn;
    ImageList1: TImageList;
    dlgOpen: TOpenDialog;
    Panel4: TPanel;
    btnParamEdit: TBitBtn;
    btnParamCancel: TBitBtn;
    btnParamSave: TBitBtn;
    pnlSysParam: TPanel;
    frmSysParam_GB_1: TGroupBox;
    DBText1: TDBText;
    DBText2: TDBText;
    frmSysParam_Lbl_1: TStaticText;
    frmSysParam_Lbl_2: TStaticText;
    dbtSysName: TDBEdit;
    DBEdit6: TDBEdit;
    frmSysParam_Lbl_4: TStaticText;
    frmSysParam_Lbl_3: TStaticText;
    GroupBox1: TGroupBox;
    StaticText1: TStaticText;
    StaticText2: TStaticText;
    dbtIP: TDBEdit;
    dbtSPort: TDBEdit;
    StaticText6: TStaticText;
    dbtTimeout: TDBEdit;
    GroupBox4: TGroupBox;
    frmSysParam_Chb_1: TDBCheckBox;
    frmSysParam_GB_3: TGroupBox;
    spbChangeDir: TSpeedButton;
    frmSysParam_Lbl_13: TStaticText;
    dbtLogPath: TDBEdit;
    Label4: TLabel;
    DBEdit1: TDBEdit;
    Label5: TLabel;
    DBEdit2: TDBEdit;
    Label6: TLabel;
    DBEdit3: TDBEdit;
    TabSheet3: TTabSheet;
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
    dsrControl: TDataSource;
    Label10: TLabel;
    dbtControlCommandType: TDBEdit;
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
    dsrProduct: TDataSource;
    dsrFeedback: TDataSource;
    dsrOperation: TDataSource;
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
    frmSysParam_Chb_2: TDBCheckBox;
    frmSysParam_Tab_3: TTabSheet;
    Panel8: TPanel;
    btnNetworkIDAppend: TBitBtn;
    btnNetworkIDEdit: TBitBtn;
    btnNetworkIDDelete: TBitBtn;
    btnNetworkIDCancel: TBitBtn;
    btnNetworkIDSave: TBitBtn;
    pnlSingleNetworkIdData: TPanel;
    frmSysParam_Lbl_17: TLabel;
    frmSysParam_Lbl_18: TLabel;
    Label24: TLabel;
    DBEdit9: TDBEdit;
    DBEdit10: TDBEdit;
    DBEdit11: TDBEdit;
    pnlMultiNetworkIdData: TPanel;
    dbgNetwordID: TDBGrid;
    frmSysParam_Lbl_19: TLabel;
    DBEdit12: TDBEdit;
    dsrNetworkId: TDataSource;
    frmSysParam_GB_2: TGroupBox;
    frmSysParam_Chb_3: TDBCheckBox;
    StaticText9: TStaticText;
    DBEdit7: TDBEdit;
    frmSysParam_Chb_4: TDBCheckBox;
    frmSysParam_Chb_5: TDBCheckBox;
    StaticText13: TStaticText;
    dbtBroadcastSDate: TDBEdit;
    StaticText14: TStaticText;
    dbtBroadcastEDate: TDBEdit;
    frmSysParam_Lbl_8: TStaticText;
    dbtCmdRefreshRate1: TDBEdit;
    frmSysParam_Lbl_9: TStaticText;
    frmSysParam_Lbl_6: TStaticText;
    dbtMaxCommandCount: TDBEdit;
    frmSysParam_Lbl_7: TStaticText;
    frmSysParam_Lbl_11: TStaticText;
    dbtCmdRefreshRate2: TDBEdit;
    frmSysParam_Lbl_12: TStaticText;
    frmSysParam_Lbl_10: TStaticText;
    dbtResentTimes: TDBEdit;
    frmSysParam_Lbl_5: TStaticText;
    StaticText21: TStaticText;
    dbtSHotTime: TDBEdit;
    dbtEHotTime: TDBEdit;
    DBEdit8: TDBEdit;
    DBEdit13: TDBEdit;
    dsrParam: TDataSource;
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
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
    procedure setLanguageString;
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
      MessageDlg('系統參數必須輸入',mtWarning,[mbOK],0);
    }        
end;

procedure TfrmSysParam.FormActivate(Sender: TObject);
begin
    frmMain.ActiveCDS(dtmMain.cdsParam, SYS_PARAM_FILENAME, 'B'); //B:browse mode; E:edit mode
    frmMain.ActiveCDS(dtmMain.cdsUser, USER_INFO_FILENAME, 'B');  //B:browse mode; E:edit mode
    frmMain.ActiveCDS(dtmMain.cdsEmm, EMM_INFO_FILENAME, 'B');  //B:browse mode; E:edit mode
    frmMain.ActiveCDS(dtmMain.cdsControl, CONTROL_INFO_FILENAME, 'B');  //B:browse mode; E:edit mode

    frmMain.ActiveCDS(dtmMain.cdsProduct, PRODUCT_INFO_FILENAME, 'B');  //B:browse mode; E:edit mode
    frmMain.ActiveCDS(dtmMain.cdsFeedback, FEEDBACK_INFO_FILENAME, 'B');  //B:browse mode; E:edit mode
    frmMain.ActiveCDS(dtmMain.cdsOperation, OPERATION_INFO_FILENAME, 'B');  //B:browse mode; E:edit mode
    frmMain.ActiveCDS(dtmMain.cdsNetworkID, NETWORK_ID_INFO_FILENAME, 'B');  //B:browse mode; E:edit mode

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

procedure TfrmSysParam.spbChangeDirClick(Sender: TObject);
begin
    dlgOpen.InitialDir := GetCurrentDir;
    if (dlgOpen.Execute) then
    begin
      dtmMain.cdsParam.FieldByName('sLogPath').AsString := ExtractFilePath(dlgOpen.FileName);
    end;
end;

procedure TfrmSysParam.dbtLogPathExit(Sender: TObject);
begin
    if dtmMain.cdsParam.FieldByName('sLogPath').AsString='' then Exit;
    if (not (SetCurrentDir(dtmMain.cdsParam.FieldByName('sLogPath').AsString))) then
    begin                 
      MessageDlg(frmMain.getLanguageSettingInfo('Sim_CA_Msg_29_Content'), mtInformation,[mbOk], 0);

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
    if dtmMain.cdsParam.RecordCount = 0 then
    begin
      result := bL_Result;
      Exit;
    end;

    with dtmMain.cdsParam do
    begin
      nL_DateLength := 10;
      sL_BroadcastStartDate := FieldByName('sBroadcastSDate').AsString;
      sL_BroadcastEndDate := FieldByName('sBroadcastEDate').AsString;

      if (FieldByName('sSysName').Value=NULL) or (FieldByName('sSysName').Value='') then
      begin
        MessageDlg(frmMain.getLanguageSettingInfo('Sim_CA_Msg_30_Content'), mtError, [mbOK],0);

        dbtSysName.SetFocus;
        Result := false;
        Exit;
      end;
      
      if (FieldByName('sSysVersion').Value=NULL) or (FieldByName('sSysVersion').Value='') then
      begin
        MessageDlg(frmMain.getLanguageSettingInfo('Sim_CA_Msg_31_Content'), mtError, [mbOK],0);

        DBEdit6.SetFocus;
        Result := false;
        Exit;
      end;


      if (FieldByName('sServerIP').Value=NULL) or (FieldByName('sServerIP').Value='') then
      begin
        MessageDlg(frmMain.getLanguageSettingInfo('Sim_CA_Msg_32_Content')+ ' IP!', mtError, [mbOK],0);
        dbtIP.SetFocus;
        Result := false;
        Exit;
      end;

      if (FieldByName('nSPortNo').Value=NULL) then
      begin
        MessageDlg(frmMain.getLanguageSettingInfo('Sim_CA_Msg_32_Content') + ' S-Port!', mtError, [mbOK],0);
        dbtSPort.SetFocus;
        Result := false;
        Exit;
      end;
      if (FieldByName('nRPortNo').Value=NULL) then
      begin
        MessageDlg(frmMain.getLanguageSettingInfo('Sim_CA_Msg_32_Content') +' R-Port!', mtError, [mbOK],0);
        dbtRPort.SetFocus;
        Result := false;
        Exit;
      end;
      if (FieldByName('nTimeOut').Value=NULL) then
      begin
        MessageDlg(frmMain.getLanguageSettingInfo('Sim_CA_Msg_32_Content') + ' Timeout!', mtError, [mbOK],0);
        dbtTimeout.SetFocus;
        Result := false;
        Exit;
      end;
      
      if (length(sL_BroadcastStartDate)<>nL_DateLength) then
      begin
        MessageDlg(frmMain.getLanguageSettingInfo('Sim_CA_Msg_33_Content'), mtError, [mbOK],0);

        dbtBroadcastSDate.SetFocus;
        Result := false;
        Exit;
      end;

      if (length(sL_BroadcastEndDate)<>nL_DateLength) then
      begin
        MessageDlg(frmMain.getLanguageSettingInfo('Sim_CA_Msg_33_Content'), mtError, [mbOK],0);
        dbtBroadcastEDate.SetFocus;
        Result := false;
        Exit;
      end;
      if not TUdateTime.IsDateStr(sL_BroadcastStartDate,'/') then
      begin
        MessageDlg(frmMain.getLanguageSettingInfo('Sim_CA_Msg_33_Content'), mtError, [mbOK],0);
        dbtBroadcastSDate.SetFocus;
        Result := false;
        Exit;
      end;

      if not TUdateTime.IsDateStr(sL_BroadcastEndDate,'/') then
      begin
        MessageDlg(frmMain.getLanguageSettingInfo('Sim_CA_Msg_33_Content'), mtError, [mbOK],0);
        dbtBroadcastEDate.SetFocus;
        Result := false;
        Exit;
      end;

      if (FieldByName('sLogPath').Value=NULL) or (FieldByName('sLogPath').Value='') then
      begin
        MessageDlg(frmMain.getLanguageSettingInfo('Sim_CA_Msg_35_Content'), mtError, [mbOK],0);
        dbtLogPath.SetFocus;
        Result := false;
        Exit;
      end;


      if (FieldByName('nMaxCommandCount').Value=NULL)then
      begin
        MessageDlg(frmMain.getLanguageSettingInfo('Sim_CA_Msg_33_Content'), mtError, [mbOK],0);

        dbtMaxCommandCount.SetFocus;
        Result := false;
        Exit;
      end;

      if (FieldByName('nMaxCommandCount').AsInteger>MAX_COMMAND_COUNT_PER_TIMES)then
      begin
        MessageDlg( frmMain.getLanguageSettingInfo('Sim_CA_Msg_36_Content')+ IntToStr(MAX_COMMAND_COUNT_PER_TIMES) +frmMain.getLanguageSettingInfo('Sim_CA_Msg_37_Content'), mtError, [mbOK],0);

        dbtMaxCommandCount.SetFocus;
        Result := false;
        Exit;
      end;

      if (FieldByName('nSourceID').Value=NULL)then
      begin
        MessageDlg(frmMain.getLanguageSettingInfo('Sim_CA_Msg_32_Content') + ' Source ID!', mtError, [mbOK],0);
        DBEdit1.SetFocus;
        Result := false;
        Exit;
      end;

      if (FieldByName('nDestID').Value=NULL)then
      begin
        MessageDlg(frmMain.getLanguageSettingInfo('Sim_CA_Msg_32_Content')+ ' Dest ID!', mtError, [mbOK],0);
        DBEdit2.SetFocus;
        Result := false;
        Exit;
      end;

      if (FieldByName('nMopPPID').Value=NULL) then
      begin
        MessageDlg(frmMain.getLanguageSettingInfo('Sim_CA_Msg_32_Content') + ' Mop PPID!', mtError, [mbOK],0);
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
      dtmMain.cdsParam.FieldByName('sMisDbPath').AsString := ExtractFilePath(dlgOpen.FileName);
    end;

end;

procedure TfrmSysParam.btnParamEditClick(Sender: TObject);
begin

    
    dtmMain.cdsParam.Edit;
    ChangeSysParamBtnStatus;

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
      dtmMain.cdsParam.Cancel;
      ChangeSysParamBtnStatus;

end;

procedure TfrmSysParam.btnParamSaveClick(Sender: TObject);
begin

    if ISDataOK then
    begin
      dtmMain.cdsParam.FieldByName('sUptName').AsString := sG_LoginName;
      dtmMain.cdsParam.FieldByName('dUptTime').AsDateTime := now;
      if dtmMain.cdsParam.FieldByName('nCreditLimit').AsInteger<=0 then
         dtmMain.cdsParam.FieldByName('nCreditLimit').AsInteger := 0;
      dtmMain.cdsParam.Post;
      ChangeSysParamBtnStatus;
      preSaveDB;

      MessageDlg(frmMain.getLanguageSettingInfo('Sim_CA_Msg_38_Content'), mtConfirmation, [mbOK], 0);

        
    end;
end;

procedure TfrmSysParam.btnEmmSaveClick(Sender: TObject);
begin

      dtmMain.cdsEmm.Post;
      ChangeEmmParamBtnStatus;

end;

procedure TfrmSysParam.ChangeEmmParamBtnStatus;
begin
     with dtmMain.cdsEmm do
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
    dtmMain.cdsEmm.Edit;
    ChangeEmmParamBtnStatus;

end;

procedure TfrmSysParam.btnEmmCancelClick(Sender: TObject);
begin
      dtmMain.cdsEmm.Cancel;
      ChangeEmmParamBtnStatus;

end;

procedure TfrmSysParam.btnControlSaveClick(Sender: TObject);
begin
    if (Uppercase(dtmMain.cdsControl.FieldByName('sBroadcastMode').AsString) <> 'N') and
       (Uppercase(dtmMain.cdsControl.FieldByName('sBroadcastMode').AsString) <> 'B') then
    begin
    
      MessageDlg(frmMain.getLanguageSettingInfo('Sim_CA_Msg_27_Content')+' broadcast mode.(' + frmMain.getLanguageSettingInfo('Sim_CA_Msg_27_Content') +' N or B)', mtError, [mbOK],0);
      dbtBroadcastMode.SetFocus;
      exit;
    end;
    if (Uppercase(dtmMain.cdsControl.FieldByName('sAddressType').AsString) <> 'U') and
       (Uppercase(dtmMain.cdsControl.FieldByName('sAddressType').AsString) <> 'G') then
    begin
      MessageDlg(frmMain.getLanguageSettingInfo('Sim_CA_Msg_27_Content') +' address type.(' + frmMain.getLanguageSettingInfo('Sim_CA_Msg_27_Content') +' U or G)', mtError, [mbOK],0);
      dbtAddressType.SetFocus;
      exit;
    end;
    dtmMain.cdsControl.FieldByName('sAddressType').AsString := UpperCase(dtmMain.cdsControl.FieldByName('sAddressType').AsString);
    dtmMain.cdsControl.FieldByName('sBroadcastMode').AsString := UpperCase(dtmMain.cdsControl.FieldByName('sBroadcastMode').AsString);    
    dtmMain.cdsControl.Post;
    ChangeControlParamBtnStatus;

end;

procedure TfrmSysParam.btnControlCancelClick(Sender: TObject);
begin
      dtmMain.cdsControl.Cancel;
      ChangeControlParamBtnStatus;

end;

procedure TfrmSysParam.btnControlEditClick(Sender: TObject);
begin
    dtmMain.cdsControl.Edit;
    ChangeControlParamBtnStatus;

end;

procedure TfrmSysParam.ChangeControlParamBtnStatus;
begin
     with dtmMain.cdsControl do
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
     with dtmMain.cdsFeedback do
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
     with dtmMain.cdsOperation do
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
     with dtmMain.cdsProduct do
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
      dtmMain.cdsProduct.Post;
      ChangeProductParamBtnStatus;

end;

procedure TfrmSysParam.btnProductCancelClick(Sender: TObject);
begin
      dtmMain.cdsProduct.Cancel;
      ChangeProductParamBtnStatus;

end;

procedure TfrmSysParam.btnProductEditClick(Sender: TObject);
begin
      dtmMain.cdsProduct.Edit;
      ChangeProductParamBtnStatus;

end;

procedure TfrmSysParam.btnFeedbackEditClick(Sender: TObject);
begin
      dtmMain.cdsFeedback.Edit;
      ChangeFeedbackParamBtnStatus;

end;

procedure TfrmSysParam.btnFeedbackCancelClick(Sender: TObject);
begin
      dtmMain.cdsFeedback.Cancel;
      ChangeFeedbackParamBtnStatus;

end;

procedure TfrmSysParam.btnFeedbackSaveClick(Sender: TObject);
begin
      dtmMain.cdsFeedback.Post;
      ChangeFeedbackParamBtnStatus;

end;

procedure TfrmSysParam.btnOperationSaveClick(Sender: TObject);
begin
      dtmMain.cdsOperation.Post;
      ChangeOperationParamBtnStatus;

end;

procedure TfrmSysParam.btnOperationEditClick(Sender: TObject);
begin
      dtmMain.cdsOperation.Edit;
      ChangeOperationParamBtnStatus;

end;

procedure TfrmSysParam.btnOperationCancelClick(Sender: TObject);
begin
      dtmMain.cdsOperation.Cancel;
      ChangeOperationParamBtnStatus;

end;

procedure TfrmSysParam.getControlCommandSectionParam(var sI_BroadastMode,
  sI_AddressType: String);
begin
    frmMain.ActiveCDS(dtmMain.cdsControl, CONTROL_INFO_FILENAME, 'B');  //B:browse mode; E:edit mode
    sI_BroadastMode := dtmMain.cdsControl.fieldByName('sBroadcastMode').AsString;
    sI_AddressType := dtmMain.cdsControl.fieldByName('sAddressType').AsString;
    dtmMain.cdsControl.Close;    
end;

procedure TfrmSysParam.preSaveDB;
begin
    SaveCDS(dtmMain.cdsParam, SYS_PARAM_FILENAME);
    SaveCDS(dtmMain.cdsUser, USER_INFO_FILENAME);
    SaveCDS(dtmMain.cdsEmm, EMM_INFO_FILENAME);
    SaveCDS(dtmMain.cdsControl, CONTROL_INFO_FILENAME);
    SaveCDS(dtmMain.cdsProduct, PRODUCT_INFO_FILENAME);
    SaveCDS(dtmMain.cdsFeedback, FEEDBACK_INFO_FILENAME);
    SaveCDS(dtmMain.cdsOperation, OPERATION_INFO_FILENAME);
    SaveCDS(dtmMain.cdsNetworkID, NETWORK_ID_INFO_FILENAME);
end;

procedure TfrmSysParam.ChangeNetworkIDBtnStatus;
begin
     with dtmMain.cdsNetworkID do
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
          if dtmMain.cdsUser.RecordCount>0 then
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
      dtmMain.cdsNetworkID.Edit;
      ChangeNetworkIDBtnStatus;

end;

procedure TfrmSysParam.btnNetworkIDCancelClick(Sender: TObject);
begin
      dtmMain.cdsNetworkID.Cancel;
      ChangeNetworkIDBtnStatus;

end;

procedure TfrmSysParam.btnNetworkIDSaveClick(Sender: TObject);
begin
      dtmMain.cdsNetworkID.Post;
      ChangeNetworkIDBtnStatus;

end;

procedure TfrmSysParam.btnNetworkIDAppendClick(Sender: TObject);
begin
      dtmMain.cdsNetworkID.Append;
      ChangeNetworkIDBtnStatus;

end;

procedure TfrmSysParam.btnNetworkIDDeleteClick(Sender: TObject);
begin

    if MessageDlg(frmMain.getLanguageSettingInfo('Sim_CA_Msg_26_Content'),mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      dtmMain.cdsNetworkID.Delete;
      ChangeNetworkIDBtnStatus;

    end;

end;

procedure TfrmSysParam.FormShow(Sender: TObject);
begin
    setLanguageString;
end;

procedure TfrmSysParam.setLanguageString;
begin
    frmSysParam.Caption := frmMain.getLanguageSettingInfo('frmSysParam_Caption');

    btnExit.Caption := frmMain.getLanguageSettingInfo('frmSysParam_btnExit_Caption');
    btnParamEdit.Caption := frmMain.getLanguageSettingInfo('frmSysParam_btnEdit_Caption');
    btnParamCancel.Caption := frmMain.getLanguageSettingInfo('frmSysParam_btnCancel_Caption');
    btnParamSave.Caption := frmMain.getLanguageSettingInfo('frmSysParam_btnSave_Caption');

    btnAppend.Caption := frmMain.getLanguageSettingInfo('frmSysParam_btnAppend_Caption');
    btnEdit.Caption := frmMain.getLanguageSettingInfo('frmSysParam_btnEdit_Caption');
    btnDelete.Caption := frmMain.getLanguageSettingInfo('frmSysParam_Delete_Caption');
    btnCancel.Caption := frmMain.getLanguageSettingInfo('frmSysParam_btnCancel_Caption');
    btnSave.Caption := frmMain.getLanguageSettingInfo('frmSysParam_btnSave_Caption');

    btnEmmEdit.Caption := frmMain.getLanguageSettingInfo('frmSysParam_btnEdit_Caption');
    btnEmmCancel.Caption := frmMain.getLanguageSettingInfo('frmSysParam_btnCancel_Caption');
    btnEmmSave.Caption := frmMain.getLanguageSettingInfo('frmSysParam_btnSave_Caption');

    btnControlEdit.Caption := frmMain.getLanguageSettingInfo('frmSysParam_btnEdit_Caption');
    btnControlCancel.Caption := frmMain.getLanguageSettingInfo('frmSysParam_btnCancel_Caption');
    btnControlSave.Caption := frmMain.getLanguageSettingInfo('frmSysParam_btnSave_Caption');

    btnProductEdit.Caption := frmMain.getLanguageSettingInfo('frmSysParam_btnEdit_Caption');
    btnProductCancel.Caption := frmMain.getLanguageSettingInfo('frmSysParam_btnCancel_Caption');
    btnProductSave.Caption := frmMain.getLanguageSettingInfo('frmSysParam_btnSave_Caption');

    btnFeedbackEdit.Caption := frmMain.getLanguageSettingInfo('frmSysParam_btnEdit_Caption');
    btnFeedbackCancel.Caption := frmMain.getLanguageSettingInfo('frmSysParam_btnCancel_Caption');
    btnFeedbackSave.Caption := frmMain.getLanguageSettingInfo('frmSysParam_btnSave_Caption');


    btnOperationEdit.Caption := frmMain.getLanguageSettingInfo('frmSysParam_btnEdit_Caption');
    btnOperationCancel.Caption := frmMain.getLanguageSettingInfo('frmSysParam_btnCancel_Caption');
    btnOperationSave.Caption := frmMain.getLanguageSettingInfo('frmSysParam_btnSave_Caption');


    btnNetworkIDEdit.Caption := frmMain.getLanguageSettingInfo('frmSysParam_btnEdit_Caption');
    btnNetworkIDCancel.Caption := frmMain.getLanguageSettingInfo('frmSysParam_btnCancel_Caption');
    btnNetworkIDSave.Caption := frmMain.getLanguageSettingInfo('frmSysParam_btnSave_Caption');

    btnNetworkIDAppend.Caption := frmMain.getLanguageSettingInfo('frmSysParam_btnAppend_Caption');
    btnNetworkIDDelete.Caption := frmMain.getLanguageSettingInfo('frmSysParam_Delete_Caption');

    frmSysParam_Chb_1.Caption := frmMain.getLanguageSettingInfo('frmSysParam_Chb_1_Caption');
    frmSysParam_Chb_2.Caption := frmMain.getLanguageSettingInfo('frmSysParam_Chb_2_Caption');
    frmSysParam_Chb_3.Caption := frmMain.getLanguageSettingInfo('frmSysParam_Chb_3_Caption');
    frmSysParam_Chb_4.Caption := frmMain.getLanguageSettingInfo('frmSysParam_Chb_4_Caption');
    frmSysParam_Chb_5.Caption := frmMain.getLanguageSettingInfo('frmSysParam_Chb_5_Caption');

    frmSysParam_Lbl_1.Caption := frmMain.getLanguageSettingInfo('frmSysParam_Lbl_1_Caption');
    frmSysParam_Lbl_2.Caption := frmMain.getLanguageSettingInfo('frmSysParam_Lbl_2_Caption');
    frmSysParam_Lbl_3.Caption := frmMain.getLanguageSettingInfo('frmSysParam_Lbl_3_Caption');
    frmSysParam_Lbl_4.Caption := frmMain.getLanguageSettingInfo('frmSysParam_Lbl_4_Caption');

    frmSysParam_Lbl_5.Caption := frmMain.getLanguageSettingInfo('frmSysParam_Lbl_5_Caption');
    frmSysParam_Lbl_6.Caption := frmMain.getLanguageSettingInfo('frmSysParam_Lbl_6_Caption');
    frmSysParam_Lbl_7.Caption := frmMain.getLanguageSettingInfo('frmSysParam_Lbl_7_Caption');
    frmSysParam_Lbl_8.Caption := frmMain.getLanguageSettingInfo('frmSysParam_Lbl_8_Caption');
    frmSysParam_Lbl_9.Caption := frmMain.getLanguageSettingInfo('frmSysParam_Lbl_9_Caption');
    frmSysParam_Lbl_10.Caption := frmMain.getLanguageSettingInfo('frmSysParam_Lbl_10_Caption');
    frmSysParam_Lbl_11.Caption := frmMain.getLanguageSettingInfo('frmSysParam_Lbl_11_Caption');
    frmSysParam_Lbl_12.Caption := frmMain.getLanguageSettingInfo('frmSysParam_Lbl_12_Caption');
    frmSysParam_Lbl_13.Caption := frmMain.getLanguageSettingInfo('frmSysParam_Lbl_13_Caption');


    frmSysParam_Lbl_14.Caption := frmMain.getLanguageSettingInfo('frmSysParam_Lbl_14_Caption');
    frmSysParam_Lbl_15.Caption := frmMain.getLanguageSettingInfo('frmSysParam_Lbl_15_Caption');
    frmSysParam_Lbl_16.Caption := frmMain.getLanguageSettingInfo('frmSysParam_Lbl_16_Caption');

    frmSysParam_Lbl_17.Caption := frmMain.getLanguageSettingInfo('frmSysParam_Lbl_17_Caption');
    frmSysParam_Lbl_18.Caption := frmMain.getLanguageSettingInfo('frmSysParam_Lbl_18_Caption');
    frmSysParam_Lbl_19.Caption := frmMain.getLanguageSettingInfo('frmSysParam_Lbl_19_Caption');

    dbgUser.Columns[0].Title.Caption := frmSysParam_Lbl_14.Caption;
    dbgUser.Columns[1].Title.Caption := frmSysParam_Lbl_15.Caption;
    dbgUser.Columns[2].Title.Caption := frmSysParam_Lbl_16.Caption;

    dbgNetwordID.Columns[0].Title.Caption := frmSysParam_Lbl_17.Caption;
    dbgNetwordID.Columns[1].Title.Caption := frmSysParam_Lbl_18.Caption;
    dbgNetwordID.Columns[3].Title.Caption := frmSysParam_Lbl_19.Caption;
    



    frmSysParam_Tab_1.Caption := frmMain.getLanguageSettingInfo('frmSysParam_Tab_1_Caption');
    frmSysParam_Tab_2.Caption := frmMain.getLanguageSettingInfo('frmSysParam_Tab_2_Caption');
    frmSysParam_Tab_3.Caption := frmMain.getLanguageSettingInfo('frmSysParam_Tab_3_Caption');

    frmSysParam_GB_1.Caption := frmMain.getLanguageSettingInfo('frmSysParam_GB_1_Caption');
    frmSysParam_GB_2.Caption := frmMain.getLanguageSettingInfo('frmSysParam_GB_2_Caption');
    frmSysParam_GB_3.Caption := frmMain.getLanguageSettingInfo('frmSysParam_GB_3_Caption');
end;

end.
