unit frmMainU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  DBTables, Db, ImgList, ExtCtrls, Menus, StdCtrls, ComCtrls, ToolWin,
  DBClient, Math, scktcomp, Buttons, ADODB, inifiles, Grids, DBGrids, winsock,
  ReceivingCommandResponseThreadU, HandleResponseU, SendCommandU,
  FetchComDataThreadU,HandleUIU,
  syncobjs,   UdateTimeu, Sockets, ConstObjU, Provider, Psock, NMEcho, Variants;

  {
  Classes, forms, comctrls, syncobjs, dbtables, Sysutils, Dialogs,
  ConstObjU, Ustru, UdateTimeu, , inifiles, ADODB;
  }

type


  TfrmMain = class(TForm)
    bvlMain: TBevel;
    lblTitle: TLabel;
    stbMain: TStatusBar;
    tbrMain: TToolBar;
    tbnOption: TToolButton;
    tbnHistory: TToolButton;
    ToolButton1: TToolButton;
    tbnExit: TToolButton;
    Panel1: TPanel;
    frmMain_Pnl_1: TPanel;
    Panel4: TPanel;
    btnExit: TButton;
    mmuMain: TMainMenu;
    mniOption: TMenuItem;
    mniExit: TMenuItem;
    imlEnabled: TImageList;
    imlStatus: TImageList;
    Panel3: TPanel;
    pnlHistory: TPanel;
    Panel6: TPanel;
    frmMain_Lbl_1: TLabel;
    frmMain_Lbl_2: TLabel;
    frmMain_Lbl_5: TLabel;
    frmMain_Lbl_6: TLabel;
    Splitter1: TSplitter;
    PageControl1: TPageControl;
    frmMain_Tab_1: TTabSheet;
    frmMain_Tab_2: TTabSheet;
    BitBtn1: TBitBtn;
    Label6: TLabel;
    Label7: TLabel;
    rgpCommandLevel: TRadioGroup;
    edtHighCommandID: TEdit;
    edtLowCommandID: TEdit;
    btnShowLog: TBitBtn;
    dbgCmdLog: TDBGrid;
    dbgCmdResult: TDBGrid;
    BitBtn2: TBitBtn;
    chbStartFun: TCheckBox;
    frmMain_Lbl_3: TLabel;
    frmMain_Lbl_4: TLabel;
    StaticText1: TStaticText;
    StaticText2: TStaticText;
    StaticText4: TStaticText;
    dtpEDate: TDateTimePicker;
    StaticText6: TStaticText;
    dtpBDate: TDateTimePicker;
    StaticText5: TStaticText;
    memStbNo: TMemo;
    memProductID: TMemo;
    memIccNo: TMemo;
    edtNewPinCode: TEdit;
    chbNewPinCode: TCheckBox;
    ADOQuery1: TADOQuery;
    memErrorLog: TMemo;
    livStatus: TListView;
    Memo1: TMemo;
    Button1: TButton;
    Timer1: TTimer;
    Timer2: TTimer;
    Timer3: TTimer;
    procedure Panel3Click(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure tbnExitClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure mniOptionClick(Sender: TObject);
    procedure mniExitClick(Sender: TObject);
    procedure tbnOptionClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure frmMain_Tab_2Show(Sender: TObject);
    procedure btnShowLogClick(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure chbStartFunClick(Sender: TObject);
    procedure SendCmdClntSocketError(Sender: TObject; Socket: TCustomWinSocket;
      ErrorEvent: TErrorEvent; var ErrorCode: Integer);
    procedure SendCmdClntSocketDisconnect(Sender: TObject;
      Socket: TCustomWinSocket);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Button1Click(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure Timer2Timer(Sender: TObject);
    procedure Timer3Timer(Sender: TObject);



  private
    { Private declarations }
    dG_Sys1, dG_Sys2, dG_Sys3 : TDate;
    sG_Sys4, sG_Sys5 : String;
    bG_LoginOk : boolean;

    G_FetchComDataThread_ExitCode : Cardinal;
    G_SendCommandThread_ExitCode : Cardinal;
    G_ReceivingCommandResponseThread_ExitCode : Cardinal;
    G_THandleResponse_ExitCode, G_HandleUIThread_ExitCode : Cardinal;


    sG_MonitorIDStatus: String;
    sG_LoginID, sG_LoginName : String;
    bG_AutoInvoke : boolean;

    G_LanguageSettingIni : TInifile;
    
    procedure setLanguageString;
    procedure initLanguageIniFile;

    function GetRegistryInfo:boolean;
    function writeRegData(sI_EncryptedLicenseKey:STring):boolean;    
    procedure createLivItem;
    procedure saveThreadInfo(sI_Info:String);
    function getVirtualNotes(sI_HighLevelCmdID:String):String;
    procedure FetchCmdDataTimer(Sender: TObject);
    procedure createThread;
    function IsDataOK:boolean;
    procedure GetSysParam;
    procedure GetEmmParam;
    procedure GetControlParam;
    procedure GetProductParam;
    procedure GetFeedbackParam;
    procedure GetOperationParam;
    function getSysIniSettiing:WideString;
    procedure  closeProg;
    procedure releaseResource;
    procedure closeSocket;
    procedure transNetworkIdInfo;//將 networkid 之資料由 CDS 轉到 string list
  public
    { Public declarations }
//    bG_CriticalSection : boolean;

//down, 以下是 ReceiveResponse thread 之變數..
    sG_RR_TestInfo : String;
    G_RR_DeviceData : TDeviceData;
//up, 以下是 Receiveresponse thread 之變數..


//down, 以下是 FetchComData thread 之變數..
//    CRS: TCriticalSection;
    nG_FD_CurActiveConn : Integer;

//    G_FD_CmdInfoObj : TCmdInfoObject;
    sG_FD_SeqNos : STring;
    nG_FD_TestingCmdSeqNo, nG_FD_MaxTestingCmdSeqNo : Integer;
//up, 以下是 FetchComData thread 之變數..

//down, 以下是 HandleResponse thread 之變數..
//    CRS: TCriticalSection;
//    sG_SeqNoForUpdateCmdStatus, sG_TransNumForUpdateCmdStatus : String;
    sG_HR_ResponseStr : String;
    nG_HR_MsgStrListPos : Integer;
//    sG_CmdStatus : String;
//    sG_ErrorCode, sG_ErrorMsg : String;
    nG_HR_CompCode, nG_HR_ConnectionID : Integer;

//up, 以下是 HandleResponse thread 之變數..

//down, 以下是 SendCommand thread 之變數..
//    G_WinSocketStreamForSendCmd : TWinSocketStream;
    sG_SC_ActiveMessage, sG_SC_MasterIccNo : String;
    nG_SC_SegmentNumber, nG_SC_TotalSeqment, nG_SC_MailId, nG_SC_RealCmdSeqNo, nG_SC_MaxRealCmdSeqNo : Integer;
    sG_SC_CmdStatus : String;

    sG_SC_TransNumForUpdateCmdStatus : String;





    sG_SC_FullCommand : String;
//    sG_CommandNo : String;
//    sG_ZipCode : String;
//    sG_ImsProductID : String;

//    sG_PhoneNum1, sG_PhoneNum2, sG_PhoneNum3 : String;
//    nG_SC_CreditMode, nG_SC_Credit, nG_SC_ThresholdCredit : Integer;
//    sG_EventName, sG_CCNum1 : String;
//    sG_IpAddr, sG_CallbackDate, sG_CallbackTime : String;
    nG_SC_CreditLimit : Integer;

    sG_SC_ErrorCode,sG_SC_ErrorMsg, sG_RecSeqNo : String;

//    nG_Timer: integer;        //Time Out計數器
    bG_SC_MakeSmsCommand : boolean;
//  sG_Notes,  sG_SeqNo : String;
    sG_SC_Operator : String;
//    sG_IccNo, sG_StbNo : String;
//    sG_CallFreq, sG_FirstCallbackDate : String;
//===============================================


//    sG_BroadcastStartDate, sG_BroadcastEndDate : String;
//    CRS: TCriticalSection;

    {
    sG_EmmBroadcastMode, sG_EmmAddressType : String;

    sG_EmmCommandType : String
    sG_ControlCommandType : String;
    sG_ProductCommandType : String;
    sG_FeedbackCommandType : String;
    sG_OperationCommandType : String;
    }

    bG_SC_buildSendCmdConnection, bG_SC_SendComm : boolean;

    nG_SC_ConnectionID : Integer;
    sG_SrcTransNum : String;// 會被填入 Nagra  傳來的 transaction num => return path 時會用到
//Up, 以下是 SendCommand thread 之變數..

    sG_CurTimeStamp : String;
    sG_CurActiveLowLevelCmd : String;    
    sG_TransactionNum : String;
    G_NetworkIdInfoStrList : TStringList;
    G_DisConnDbStrList : TStringList;
    G_SendCmdClntSocket: TClientSocket;
    nG_LivItemPos : Integer;
    G_TimeStampStrList : TStringList;
    sG_NewIniFileName : String;

    G_HandleUIThread : THandleUI;
    G_FetchComDataThread : FetchComDataThread;
    G_SendCommandThread : TSendCommand;
    G_ReceivingCommandResponseThread : ReceivingCommandResponseThread;
    G_THandleResponse : THandleResponse;


    nG_CurActiveConn, nG_DbCount : Integer;
    G_CompCodeStrList:TStringList;
    G_DbAliasStrList:TStringList;
    bG_StartFetching, bG_Simulator, bG_InputLicense : boolean;
//    G_Timer: TTimer;
    G_NotesHasValueStrList : TStringList;
    sG_Notes_Has_Value, sG_TimeDefine, sG_WriteTimeStamp : String;

    G_CommandInfoStrList : TSTringList;
    bG_HandleUIThread, bG_HandleResponseThread, bG_RunFetchCommandThread, bG_RunReceivingCommandThread,bG_RunSendingCommandThread : boolean;
    bG_HasBuildSendCmdConnection, bG_HasBuildReturnPathConnection : boolean;

    G_SendCmdWStream, G_ReceiveCmdWStream, G_ReturnPathWStream: TWinSocketStream;
    sG_SourceID, sG_DestID, sG_MopPPID : String;
    sG_BroadcastStartDate, sG_BroadcastEndDate : String;
    bG_SetZipCode, bG_CommandLog, bG_ErrorLog, bG_ResponseLog, bG_ShowUI, bG_AssignBroadcastDate : boolean;
    sG_CommandLog, sG_ErrorLog, sG_ResponseLog: String;
    sG_ServerIP, sG_LogPath : String;
    nG_CmdResentTimes, nG_DefaultCreditLimit : Integer;
    nG_SPortNo, nG_RPortNo, nG_Timeout, nG_MaxCommandCount, nG_CmdRefreshRate1, nG_CmdRefreshRate2 : Integer;
    sG_EmmCommandType, sG_EmmBroadcastMode, sG_EmmAddressType : String;
    sG_ControlCommandType : String;
    sG_ProductCommandType : String;
    sG_FeedbackCommandType : String;
    sG_OperationCommandType : String;
    G_ResponseMsgStrList : TStringList;
    function getLanguageSettingInfo(sI_Key:String):String;    
    procedure reBuildNagraConn;
    procedure showLogOnMemo(nI_ActiveConn:Integer; sI_Detail:String);
    procedure setTransCaption;
    procedure buildSendCmdConnection;
    function reDefineFetchDataInterval(dI_Time:TDate):Integer;
    function getConnectionID(nI_CompCode:Integer):Integer;
    function getCompNetworkInfo(nI_CompCode:Integer; sI_IrmCmdID:String):String;
    procedure ReturnPathSvrThreadDone(Sender: TObject);
    procedure ReceiveResponseThreadDone(Sender: TObject);
    procedure FetchCmdDataThreadDone(Sender: TObject);
    procedure HandleResponseThreadDone(Sender: TObject);
    procedure SendCommandThreadDone(Sender: TObject);
    procedure HandleUIThreadDone(Sender: TObject);
    procedure ActiveCDS(I_CDS:TClientDataSet; sI_RefFileName:String; sI_DefaultStatus:String);
//    procedure updateCmdStatusToProcessing(sI_SeqNos : String);
    procedure getAppInfo(var sI_AppPath, sI_AppId: String);
    function getSendCommandServerIP:STring;
    function getSendCommandPort:Integer;
    function getReturnPathServerIP:STring;
    function getReturnPathPort:Integer;

    function getListenCommandPort:Integer;
    function getBinaryVal(nI_Value: Integer; nI_Digits: Integer; bI_Reverse: boolean):TTempRecord;
    function reverseBinaryVal(I_TmpAry:array of Char):Integer;
    function transReceivedDataCharAryToStr(I_TmpAry:array of Char; nI_DataLength:Integer):String;
    procedure getErrorInfo(sI_ResErrorCode,sI_ResErrorCodeExt: String; var sI_ResErrorCodeDetail,sI_ResErrorCodeExtDetail: String);

    procedure createThreadForNagra(nI_Operation:Integer; sI_ErrorMsg:String);    
  end;


var
  frmMain: TfrmMain;

implementation

uses frmSysParamU, frmLoginU, Ustru, dtmMainU, prjMonitor_TLB, Uotheru;


{$R *.DFM}

procedure TfrmMain.Panel3Click(Sender: TObject);
begin
    pnlHistory.Visible := true;
end;

procedure TfrmMain.btnExitClick(Sender: TObject);
begin

    if MessageDlg(getLanguageSettingInfo('frmMain_Msg_1_Content'), mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      closeProg;
    end;  
end;

procedure TfrmMain.tbnExitClick(Sender: TObject);
begin
    btnExitClick(Sender);
end;

procedure TfrmMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    Action := caFree;
end;

procedure TfrmMain.mniOptionClick(Sender: TObject);
begin
    tbnOptionClick(Sender);
end;

procedure TfrmMain.mniExitClick(Sender: TObject);
begin
    btnExitClick(Sender);
end;

procedure TfrmMain.tbnOptionClick(Sender: TObject);
begin
    frmSysParam := TfrmSysParam.Create(Application);
    frmSysParam.sG_LoginName := sG_LoginName;
    frmSysParam.ShowModal;
    frmSysParam.free;

end;



procedure TfrmMain.FormCreate(Sender: TObject);
var
    ii : Integer;
    sL_BdeAlias, sL_Result, sL_EncryptedLicenseKey : String;
    sL_DbAlias, sL_DbUserName : String;
    dL_SysDate : TDate;
    procedure writeFlagToLockSoftware;
    begin
      TUother.RegWrite(REGISTRY_ROOT,'SYS5',2,SYS5_Y); //寫入停用本軟體之FLAG
    end;

begin
    initLanguageIniFile;
    setLanguageString;

    bG_Simulator := false;

    bG_AutoInvoke := false;
    sG_NewIniFileName := '';

    if ParamCount > 0 then
    begin
        for ii:=1 to ParamCount do
        begin
          if UpperCase(ParamStr(ii)) = AUTO_INVOKE then
          begin
            bG_AutoInvoke := true;
          end;
          if AnsiPos('.INI',UpperCase(ParamStr(ii)))>0 then
          begin
            //表示指明要用哪一個 ini file...
            sG_NewIniFileName := ParamStr(ii);
          end;
          if UpperCase(ParamStr(ii)) = 'SIM' then
          begin
            bG_Simulator := true;
          end;
          if UpperCase(ParamStr(ii)) = 'IL' then
          begin
            bG_InputLicense := true;
          end;
        end;
    end;

    if bG_InputLicense then //表示使用者要輸入 license key...
    begin
      try

        sL_EncryptedLicenseKey := InputBox(getLanguageSettingInfo('frmMain_Msg_2_Content'), 'License Key:', '');

        if sL_EncryptedLicenseKey='' then
          Application.Terminate
        else if writeRegData(sL_EncryptedLicenseKey) then
        begin
          MessageDlg(getLanguageSettingInfo('frmMain_Msg_3_Content'), mtInformation, [mbOK],0);
          Application.Terminate;
        end
        else
        begin
          MessageDlg(getLanguageSettingInfo('frmMain_Msg_4_Content'), mtError, [mbOK],0);
          Application.Terminate;
        end;
      except
        MessageDlg(getLanguageSettingInfo('frmMain_Msg_5_Content'), mtError, [mbOK],0);
        Application.Terminate;
      end;
    end;

    if not GetRegistryInfo then
    begin
      MessageDlg(getLanguageSettingInfo('frmMain_Msg_6_Content'), mtError, [mbOK],0 );
      Application.Terminate;
    end
    else
    begin
      dL_SysDate := Date;
      if sG_Sys5=SYS5_Y then //表示停用軟體
      begin
        MessageDlg(getLanguageSettingInfo('frmMain_Msg_7_Content'), mtError, [mbOK],0 );
        Application.Terminate;
      end
      else if (dL_SysDate= dG_Sys1)  then
      begin
        writeFlagToLockSoftware;
        MessageDlg(getLanguageSettingInfo('frmMain_Msg_8_Content'), mtError, [mbOK],0 );
        Application.Terminate;
      end
      else if (dL_SysDate < dG_Sys3)  then
      begin
        writeFlagToLockSoftware;
        MessageDlg(getLanguageSettingInfo('frmMain_Msg_9_Content'), mtError, [mbOK],0 );
        Application.Terminate;
      end
      else
      begin
        G_DisConnDbStrList := TStringList.Create;
        G_NetworkIdInfoStrList := TStringList.Create;



        //系統登入畫面

    //    dtmMain := TDtmMain.Create(Application);

        if not bG_AutoInvoke then
        begin

          frmLogin := TfrmLogin.Create(Application);
          with frmLogin do
          begin

            if ShowModal=mrCancel then
            begin
              bG_LoginOk := false;
            end
            else
            begin
              bG_LoginOk := true;
              sG_LoginID := sG_USerID;
              sG_LoginName := sG_UserName;
              Close;
              free;
            end;
          end;

        end
        else
        begin
          bG_LoginOk := true;
          sG_LoginID := 'SYS';
          sG_LoginName := 'Monitor';
        end;

        if bG_LoginOk then
        begin

          GetSysParam;

          GetEmmParam;

          GetControlParam;
          GetProductParam;
          GetFeedbackParam;
          GetOperationParam;

          ActiveCDS(dtmMain.cdsNetworkID, NETWORK_ID_INFO_FILENAME, 'B');
          transNetworkIdInfo;

          try
            //down, 從 ini 檔讀取連線至 DB 的 information
            sL_Result := getSysIniSettiing;
            if sL_Result<>'' then
            begin
              if not bG_AutoInvoke then
                MessageDlg(sL_Result,mtError,[mbOK],0);
              Application.Terminate;
              exit;
            end;
            //up, 從 ini 檔讀取連線至 DB 的 information
            caption := caption + ' - ('+sG_LoginName+')' + '--' + sG_NewIniFileName;
            //down, 建立DB的 connection
            G_CompCodeStrList := TStringList.Create;
            G_DbAliasStrList := TStringList.Create;

            sL_Result := dtmMain.connectToDB(nG_DbCount, G_CompCodeStrList, G_DbAliasStrList);

            if sL_Result<>'' then
            begin
              if not bG_AutoInvoke then
                MessageDlg(sL_Result,mtError,[mbOK],0);
              Application.Terminate;
              exit;
            end;
            //up, 建立DB的 connection

          except
//            mniHistory.Visible := false;
            tbnHistory.Visible := false;
            MessageDlg(getLanguageSettingInfo('frmMain_Msg_10_Content') + sL_DbAlias + ',UserID:' + sL_DbUserName + '),請修改您的系統參數,並重新啟動本程式!',mtError,[mbOK],0);
            Application.Terminate;
          end;


          G_ResponseMsgStrList := TStringList.Create;
          G_CommandInfoStrList := TSTringList.Create;

          G_TimeStampStrList := TStringList.Create;
      //    createThread;
          Timer1.Enabled := false;
          Timer2.Enabled := false;
          Timer3.Enabled := false;
        end
        else
          Application.Terminate;
      end;

    end;

end;
procedure TfrmMain.GetSysParam;
var
    sL_FormatStr, sL_Path : String;
begin
    with dtmMain.cdsParam do
    begin
      if (State = dsInactive) then
        CreateDataSet;
      sL_Path := ExtractFilePath(Application.ExeName);
      if (FileExists(sL_Path + SYS_PARAM_FILENAME)) then
        LoadFromFile(sL_Path + SYS_PARAM_FILENAME);

      sG_ServerIP := FieldByName('sServerIP').AsString;
      sG_LogPath := FieldByName('sLogPath').AsString;
      nG_SPortNo := FieldByName('nSPortNo').AsInteger;
      nG_RPortNo := FieldByName('nRPortNo').AsInteger;
      nG_Timeout := FieldByName('nTimeOut').AsInteger;
      nG_MaxCommandCount := FieldByName('nMaxCommandCount').AsInteger;
      bG_CommandLog := FieldByName('bCommandLog').Asboolean;
      bG_ErrorLog := FieldByName('bErrorLog').Asboolean;
      bG_ResponseLog := FieldByName('nResponseLog').Asboolean;  //欄位命名錯誤....

      sL_FormatStr := '%.' + IntToStr(SOURCE_ID_LENGTH) + 'd';
      sG_SourceID := Format(sL_FormatStr ,  [FieldByName('nSourceID').AsInteger]);

      sL_FormatStr := '%.' + IntToStr(DEST_ID_LENGTH) + 'd';
      sG_DestID := Format(sL_FormatStr ,  [FieldByName('nDestID').AsInteger]);

      sL_FormatStr := '%.' + IntToStr(MOP_PPID_LENGTH) + 'd';
      sG_MopPPID := Format(sL_FormatStr ,  [FieldByName('nMopPPID').AsInteger]);

      bG_SetZipCode := FieldByName('bSetZipCode').Asboolean;
      nG_DefaultCreditLimit := FieldByName('nCreditLimit').AsInteger;

      nG_CmdRefreshRate1 := FieldByName('nCmdRefreshRate1').AsInteger; //尖峰時段?秒讀取一次指令資料庫
      nG_CmdRefreshRate2 := FieldByName('nCmdRefreshRate2').AsInteger; //離峰時段?秒讀取一次指令資料庫

      nG_CmdResentTimes := FieldByName('nCmdResentTimes').AsInteger;
      bG_ShowUI := FieldByName('bShowUI').Asboolean;
      bG_AssignBroadcastDate := FieldByName('bAssignBroadcastDate').Asboolean;
      if bG_AssignBroadcastDate then
      begin
        sG_BroadcastStartDate := TUstr.replaceStr(FieldByName('sBroadcastSDate').AsString, '/','');
        sG_BroadcastEndDate := TUstr.replaceStr(FieldByName('sBroadcastEDate').AsString,'/','');
      end
      else
      begin
        sG_BroadcastStartDate := TUdateTime.GetPureDateStr(date-1);//Nagra建議傳送昨天
        sG_BroadcastStartDate := TUstr.replaceStr(sG_BroadcastStartDate, '/','');

        sG_BroadcastEndDate := TUdateTime.GetPureDateStr(date+1);//Nagra建議傳送明天
        sG_BroadcastEndDate := TUstr.replaceStr(sG_BroadcastEndDate, '/','');

      end;
      sG_TimeDefine := FieldByName('sSHotTime').AsString + '-' + FieldByName('sEHotTime').AsString; // 尖峰時段定義

      if RecordCount=0 then
        MessageDlg(getLanguageSettingInfo('frmMain_Msg_11_Content'), mtWarning, [mbOK],0);
      Close;
    end;
end;

procedure TfrmMain.FormShow(Sender: TObject);
var
    ii : Integer;
begin
    if bG_LoginOk then
    begin
    
      showLogOnMemo(EXECUTE_SEND_NAGRA_FLAG,getLanguageSettingInfo('frmMain_Msg_12_Content'));
      //建立監控視窗
      sG_MonitorIDStatus := '';
  //    Timer1.Enabled := true;

      frmMain_Tab_2.TabVisible := false;
      if ParamCount > 0 then
      begin

          for ii:=1 to ParamCount do
          begin
            if UpperCase(ParamStr(ii)) = 'TESTING' then
              frmMain_Tab_2.TabVisible := true;
          end;
      end;

      chbStartFun.Checked := false;
      {
      TcpServer1.LocalHost := '10.40.16.1';
      TcpServer1.LocalPort := '60002';
      TcpServer1.Active := True;
      }

      nG_CurActiveConn := nG_DbCount-1;
      frmMain.bG_HasBuildSendCmdConnection := false;
      frmMain.bG_HasBuildReturnPathConnection := false;


      createThread;
      createLivItem;

      if bG_AutoInvoke then
        chbStartFun.Checked := true;

      //down, 讓程式畫面產生 refresh 的效果, 若不做,CA有時會無法開始傳送指令

      self.WindowState := wsMinimized;
      self.WindowState := wsMaximized;
      self.Repaint;

      //up, 讓程式畫面產生 refresh 的效果, 若不做,CA有時會無法開始傳送指令

    end;

end;

procedure TfrmMain.GetEmmParam;
var
    sL_Path : String;
begin
    with dtmMain.cdsEmm do
    begin
      if (State = dsInactive) then
        CreateDataSet;
      sL_Path := ExtractFilePath(Application.ExeName);
      if (FileExists(sL_Path + EMM_INFO_FILENAME)) then
        LoadFromFile(sL_Path + EMM_INFO_FILENAME);

      sG_EmmCommandType := FieldByName('sCommandType').AsString;
      sG_EmmBroadcastMode := FieldByName('sBroadcastMode').AsString;
      sG_EmmAddressType := FieldByName('sAddressType').AsString;
      if RecordCount=0 then
        MessageDlg(getLanguageSettingInfo('frmMain_Msg_13_Content'), mtWarning, [mbOK],0);
      Close;
    end;
end;

procedure TfrmMain.GetControlParam;
var
    sL_Path : String;
begin
    with dtmMain.cdsControl do
    begin
      if (State = dsInactive) then
        CreateDataSet;
      sL_Path := ExtractFilePath(Application.ExeName);
      if (FileExists(sL_Path + CONTROL_INFO_FILENAME)) then
        LoadFromFile(sL_Path + CONTROL_INFO_FILENAME);

      sG_ControlCommandType := FieldByName('sCommandType').AsString;
      if RecordCount=0 then
        MessageDlg(getLanguageSettingInfo('frmMain_Msg_14_Content'), mtWarning, [mbOK],0);
      Close;
    end;

end;

procedure TfrmMain.GetFeedbackParam;
var
    sL_Path : String;
begin
    with dtmMain.cdsFeedback do
    begin
      if (State = dsInactive) then
        CreateDataSet;
      sL_Path := ExtractFilePath(Application.ExeName);
      if (FileExists(sL_Path + FEEDBACK_INFO_FILENAME)) then
        LoadFromFile(sL_Path + FEEDBACK_INFO_FILENAME);

      sG_FeedbackCommandType := FieldByName('sCommandType').AsString;
      if RecordCount=0 then
        MessageDlg(getLanguageSettingInfo('frmMain_Msg_15_Content'), mtWarning, [mbOK],0);
      Close;
    end;

end;

procedure TfrmMain.GetOperationParam;
var
    sL_Path : String;
begin
    with dtmMain.cdsOperation do
    begin
      if (State = dsInactive) then
        CreateDataSet;
      sL_Path := ExtractFilePath(Application.ExeName);
      if (FileExists(sL_Path + OPERATION_INFO_FILENAME)) then
        LoadFromFile(sL_Path + OPERATION_INFO_FILENAME);

      sG_OperationCommandType := FieldByName('sCommandType').AsString;
      if RecordCount=0 then
        MessageDlg(getLanguageSettingInfo('frmMain_Msg_16_Content'), mtWarning, [mbOK],0);
      Close;
    end;

end;

procedure TfrmMain.GetProductParam;
var
    sL_Path : String;
begin
    with dtmMain.cdsProduct do
    begin
      if (State = dsInactive) then
        CreateDataSet;
      sL_Path := ExtractFilePath(Application.ExeName);
      if (FileExists(sL_Path + PRODUCT_INFO_FILENAME)) then
        LoadFromFile(sL_Path + PRODUCT_INFO_FILENAME);

      sG_ProductCommandType := FieldByName('sCommandType').AsString;
      if RecordCount=0 then
        MessageDlg(getLanguageSettingInfo('frmMain_Msg_17_Content'), mtWarning, [mbOK],0);
      Close;
    end;


end;

function TfrmMain.getSendCommandPort: Integer;
begin
    result := nG_SPortNo;
end;

function TfrmMain.getSendCommandServerIP: STring;
begin
    result := sG_ServerIP;
end;

procedure TfrmMain.BitBtn1Click(Sender: TObject);
var
    bL_UpdateIni : boolean;
    sL_SQL,sL_ProductID, sL_Notes : String;
    sL_CommandNo, sL_IccNo, sL_BoxNo : String;
    sL_ExeFileName, sL_ExePath, sL_IniFileName : String;
    L_IniFile : TIniFile;
begin

    if IsDataOK then
    begin
      sL_IccNo := memIccNo.Lines.Strings[0];
      sL_BoxNo := memStbNo.Lines.Strings[0];
      sL_ProductID := memProductID.Lines.Strings[0];

      bL_UpdateIni := true;
      case rgpCommandLevel.ItemIndex of
        0://高階指令
         begin
           bL_UpdateIni := false;
          sL_CommandNo := edtHighCommandID.Text;
         end;
        1://低階指令
         begin
           bL_UpdateIni := true;
          sL_CommandNo := edtLowCommandID.Text;
         end;
      end;
      sL_Notes := getVirtualNotes(Uppercase(edtHighCommandID.Text));
      //down, 更改 send_nagra 的設定
      sL_SQL := 'update send_nagra set cmd_status=' + '''' + 'W' + '''' + ',  HIGH_LEVEL_CMD_ID=' + '''' + sL_CommandNo + '''';
      sL_SQL := sL_SQL + ', ICC_No=' + '''' + sL_IccNo + '''';
      sL_SQL := sL_SQL + ', STB_No=' + '''' + sL_BoxNo + '''';
      sL_SQL := sL_SQL + ',RESENTTIMES=0';
      sL_SQL := sL_SQL + ', ERR_CODE=' + '''' + ' ' + '''';
      sL_SQL := sL_SQL + ', ERR_MSG=' + '''' + ' ' + '''';
      sL_SQL := sL_SQL + ', IMS_PRODUCT_ID=' + '''' + sL_ProductID + '''';
      sL_SQL := sL_SQL + ', NOTES=' + '''' + sL_Notes + '''';
      sL_SQL := sL_SQL + ', SUBSCRIPTION_BEGIN_DATE=' + TUstr.getOracleSQLDateStr(dtpBDate.date);
      sL_SQL := sL_SQL + ', SUBSCRIPTION_END_DATE=' + TUstr.getOracleSQLDateStr(dtpEDate.date);
      sL_SQL := sL_SQL + ', SEQNO=SEQNO+1 ';
      if chbNewPinCode.Checked then
      begin
        sL_SQL := sL_SQL + ', MIS_IRD_CMD_DATA=' + '''' + edtNewPinCode.Text + '''';
        sL_SQL := sL_SQL + ', MIS_IRD_CMD_ID=' + '1';
      end;




      dtmMain.G_AdoTmp[0].Close;

      dtmMain.G_AdoTmp[0].SQL.Clear;
      dtmMain.G_AdoTmp[0].SQL.Add(sL_SQL);
      dtmMain.G_AdoTmp[0].ExecSQL;
      dtmMain.G_AdoTmp[0].Close;
      //up, 更改 send_nagra 的設定

      //down, 更改 ini 的設定
      if bL_UpdateIni then
      begin
        sL_ExeFileName := Application.ExeName;
        sL_ExePath := ExtractFileDir(sL_ExeFileName);
        if sG_NewIniFileName='' then
        begin
          sG_NewIniFileName := CABLE_NAGRA_INI_FILE_NAME;
          sL_IniFileName := sL_ExePath + '\' + CABLE_NAGRA_INI_FILE_NAME;
        end
        else
          sL_IniFileName := sL_ExePath + '\' + sG_NewIniFileName;



        if  FileExists(sL_IniFileName) then
        begin
          L_IniFile := TIniFile.Create(sL_IniFileName);
          L_IniFile.WriteString('COMMANDMATRIX', sL_CommandNo, sL_CommandNo);
          L_IniFile.Free;
        end;

      end;
      //up, 更改 ini 的設定

      chbStartFun.Checked := true;

      edtHighCommandID.SetFocus;
    end;

end;

procedure TfrmMain.frmMain_Tab_2Show(Sender: TObject);
begin
    edtHighCommandID.SetFocus;
end;

function TfrmMain.IsDataOK: boolean;
begin

    case rgpCommandLevel.ItemIndex of
      0://高階指令
       begin
         if edtHighCommandID.Text = '' then
         begin
          MessageDlg(getLanguageSettingInfo('frmMain_Msg_18_Content'), mtWarning, [mbOK],0);
          edtHighCommandID.SetFocus;
          Result := false;
          exit;
         end;
       end;
      1://低階指令
       begin
         if edtLowCommandID.Text = '' then
         begin
          MessageDlg(getLanguageSettingInfo('frmMain_Msg_19_Content'), mtWarning, [mbOK],0);
          edtLowCommandID.SetFocus;
          Result := false;
          exit;
         end;
       end;
    end;

    if memIccNo.Lines.Count = 0 then
    begin
      MessageDlg(getLanguageSettingInfo('frmMain_Msg_20_Content'), mtWarning, [mbOK],0);
      memIccNo.SetFocus;
      Result := false;
      exit;
    end;
    if memProductID.Lines.Count = 0 then
    begin
      MessageDlg(getLanguageSettingInfo('frmMain_Msg_21_Content'), mtWarning, [mbOK],0);
      memProductID.SetFocus;
      Result := false;
      exit;
    end;

    {
    if length(edtSerialNo.Text)<>10 then
    begin
      MessageDlg('Serial No 應輸入 10 碼!', mtWarning, [mbOK],0);
      edtSerialNo.SetFocus;
      Result := false;
      exit;
    end;
    }
    if memStbNo.Lines.Count=0 then
    begin
      MessageDlg(getLanguageSettingInfo('frmMain_Msg_22_Content'), mtWarning, [mbOK],0);
      memStbNo.SetFocus;
      Result := false;
      exit;
    end;
    {
    由程式中補足14碼
    if length(edtBoxNo.Text)<>14 then
    begin
      MessageDlg('Box No 應輸入 14 碼!', mtWarning, [mbOK],0);
      edtBoxNo.SetFocus;
      Result := false;
      exit;
    end;
    }
    Result := true;
end;

procedure TfrmMain.btnShowLogClick(Sender: TObject);
var
    sL_TransactionNum, sL_WhereClause : String;
    sL_SQL, sL_Value : String;
begin
    case rgpCommandLevel.ItemIndex of
      0://高階指令
       begin
         sL_WhereClause := 'where HighLevelCmdNo=' + STR_SEP + edtHighCommandID.Text + STR_SEP;
       end;
      1://低階指令
       begin
         sL_WhereClause := 'where LowLevelCmdNo=' + STR_SEP + edtLowCommandID.Text + STR_SEP;
       end;
    end;
    {
    sL_Value := '(';
    with adoCmdLog do
    begin
      SQL.Clear;
      SQL.Add('select * from CmdLog ' + sL_WhereClause);
      Open;
      while not Eof do
      begin
        sL_TransactionNum := FieldByName('TransactionNum').AsString;
        Next;
        if sL_Value='(' then
          sL_Value := sL_Value + STR_SEP + sL_TransactionNum + STR_SEP
        else
          sL_Value := sL_Value + ',' + STR_SEP + sL_TransactionNum + STR_SEP ;
      end;
      sL_Value := sL_Value + ')';
    end;

    if length(sL_Value) <> 2 then
    begin
      with adoCmdResult do
      begin
        SQL.Clear;
        sL_SQL := 'select * from CmdResult where TransactionNum in ' + sL_Value ;
        SQL.Add(sL_SQL);
        Open;
      end;
    end;
    }
end;

procedure TfrmMain.BitBtn2Click(Sender: TObject);
begin
    {
    with adoCmdLog do
    begin
      Close;
      SQL.Clear;
      SQL.Add('delete from CmdLog');
      ExecSQL;
      Close;
    end;

    with adoCmdResult do
    begin
      Close;
      SQL.Clear;
      SQL.Add('delete from CmdResult');
      ExecSQL;
      Close;
    end;
    }
    {
    with adoRealCmd do
    begin
      Close;
      SQL.Clear;
      SQL.Add('delete from CmdLog');
      ExecSQL;

      Close;
      SQL.Clear;
      SQL.Add('delete from CmdResult');
      ExecSQL;
    end;
    }
end;

function TfrmMain.getListenCommandPort: Integer;
begin
    result := nG_RPortNo;
end;

function TfrmMain.getSysIniSettiing: WideString;
var
    L_IniFile : TIniFile;
    sL_ExeFileName, sL_ExePath, sL_IniFileName : STring;

begin
    sL_ExeFileName := Application.ExeName;
    sL_ExePath := ExtractFileDir(sL_ExeFileName);

    if sG_NewIniFileName='' then
    begin
      sG_NewIniFileName := CABLE_NAGRA_INI_FILE_NAME;
      sL_IniFileName := sL_ExePath + '\' + CABLE_NAGRA_INI_FILE_NAME;
    end
    else
      sL_IniFileName := sL_ExePath + '\' + sG_NewIniFileName;


    if not FileExists(sL_IniFileName) then
    begin
      result := getLanguageSettingInfo('frmMain_Msg_23_Content')+sL_IniFileName +getLanguageSettingInfo('frmMain_Msg_24_Content');
      exit;
    end;
    L_IniFile := TIniFile.Create(sL_IniFileName);

    sG_Notes_Has_Value := L_IniFile.ReadString('SYS_SETTING','NOTES_HAS_VALUE','');
    sG_WriteTimeStamp := L_IniFile.ReadString('SYS_SETTING','WRITE_TIME_STAMP','N');



    G_NotesHasValueStrList := TUStr.ParseStrings(sG_Notes_Has_Value,',');


    L_IniFile.Free;
    result := '';
end;

{
function TfrmMain.connectToDB(sI_DbConnStr, sI_DbAlias, sI_DbUserName,
  sI_DbPassword: String): WideString;
begin
    try

      dbSendNagra.AliasName := sI_DbAlias;
      dbSendNagra.Params.Clear;
      dbSendNagra.Params.Add('UserName='+sI_DbUserName);
      dbSendNagra.Params.Add('Password='+sI_DbPassword);
      dbSendNagra.DatabaseName := SEND_NAGRA_DB_NAME;
      qrySendNagra.DatabaseName := SEND_NAGRA_DB_NAME;
      qryTmp.DatabaseName := SEND_NAGRA_DB_NAME;
      dbSendNagra.Open;

    except
      result := '-1: 與CC&B資料庫連線失敗, 資料庫別名=<'+sI_DbAlias+'>, DB User=<'+sI_DbUserName +'>';
      exit;
    end;
    result := '';

end;

}
function TfrmMain.getBinaryVal(nI_Value, nI_Digits: Integer;
  bI_Reverse: boolean): TTempRecord;
var
    ResultVal: TTempRecord;
    i, nL_Index:integer;
    sL_HexVal: string;
//nI_Value:十進位的數值、nI_Digits:Binary Length (in Bytes)、bI_Reverse:是否顛倒、傳回值為一個Record Type
begin
  sL_HexVal := InttoHex(nI_Value, nI_Digits*2);
  nL_Index:=1;
  if bI_Reverse then
  begin
	  for i:= nI_Digits - 1 downto 0 do
    begin
      ResultVal.TempAry[i] := chr(strtoint('$'+copy(sL_HexVal, nL_Index, 2)));
      Inc(nL_Index,2);
  	end;
  end
  else
  begin
	  for i:= 0 to nI_Digits - 1 do
    begin
      ResultVal.TempAry[i] := chr(strtoint('$'+copy(sL_HexVal, nL_Index, 2)));
      Inc(nL_Index,2);
  	end;
  end;
  result:= ResultVal;
end;

procedure TfrmMain.getErrorInfo(sI_ResErrorCode,
  sI_ResErrorCodeExt: String; var sI_ResErrorCodeDetail,
  sI_ResErrorCodeExtDetail: String);
var
    sL_ExeFileName, sL_ExePath, sL_IniFileName : String;
    L_IniFile : TIniFile;
begin
    sL_ExeFileName := Application.ExeName;
    sL_ExePath := ExtractFileDir(sL_ExeFileName);
    if sG_NewIniFileName='' then
      sL_IniFileName := sL_ExePath + '\' + CABLE_NAGRA_INI_FILE_NAME
    else
      sL_IniFileName := sL_ExePath + '\' + sG_NewIniFileName;



    if  FileExists(sL_IniFileName) then
    begin
      L_IniFile := TIniFile.Create(sL_IniFileName);
      sI_ResErrorCodeDetail := L_IniFile.ReadString('ERROR_CODE', sI_ResErrorCode, '');
      sI_ResErrorCodeExtDetail := L_IniFile.ReadString('ERROR_CODE_EXT', sI_ResErrorCodeExt, '');
      L_IniFile.Free;
    end;

end;

procedure TfrmMain.ReceiveResponseThreadDone(Sender: TObject);
begin

    if (Sender as TThread).FatalException <> nil then
    begin
//      showmessage('777');
      saveThreadInfo('Receive Response Thread Done.');

    end;
end;




procedure TfrmMain.createThread;
var
    sL_Info : String;
begin

    sL_Info := getLanguageSettingInfo('frmMain_Msg_25_Content');
    G_HandleUIThread := THandleUI.Create;
    if GetExitCodeThread(G_HandleUIThread.ThreadID,G_HandleUIThread_ExitCode) then
    begin
      if not bG_AutoInvoke then
        MessageDlg(getLanguageSettingInfo('frmMain_Msg_26_Content') + sL_Info, mtError, [mbOK],0);
      Application.Terminate;
      Exit;
    end;


    G_FetchComDataThread := FetchComDataThread.Create;
    if GetExitCodeThread(G_FetchComDataThread.ThreadID,G_FetchComDataThread_ExitCode) then
    begin
      if not bG_AutoInvoke then
        MessageDlg(getLanguageSettingInfo('frmMain_Msg_27_Content') + sL_Info, mtError, [mbOK],0);
      Application.Terminate;
      Exit;
    end;


    createThreadForNagra(1,'');

    G_THandleResponse := THandleResponse.Create;
    if GetExitCodeThread(G_THandleResponse.ThreadID,G_THandleResponse_ExitCode) then
    begin
      if not bG_AutoInvoke then
        MessageDlg(getLanguageSettingInfo('frmMain_Msg_28_Content') + sL_Info, mtError, [mbOK],0);
      Application.Terminate;
      Exit;
    end;




    //傳入必要參數


    if G_SendCommandThread <> nil then
    begin
//      G_SendCommandThread.sG_WriteTimeStamp := sG_WriteTimeStamp;
{
      G_SendCommandThread.sG_LogPath := sG_LogPath;
      G_SendCommandThread.nG_Timeout := nG_Timeout;
}
{
      G_SendCommandThread.sG_SourceID := sG_SourceID;
      G_SendCommandThread.sG_DestID := sG_DestID;
      G_SendCommandThread.sG_MopPPID := sG_MopPPID;
}
{
      G_SendCommandThread.sG_EmmBroadcastMode := sG_EmmBroadcastMode;
      G_SendCommandThread.sG_EmmAddressType := sG_EmmAddressType;
}
{
      G_SendCommandThread.sG_EmmCommandType := sG_EmmCommandType;
      G_SendCommandThread.sG_ControlCommandType := sG_ControlCommandType;
      G_SendCommandThread.sG_ProductCommandType := sG_ProductCommandType;
      G_SendCommandThread.sG_FeedbackCommandType := sG_FeedbackCommandType;
      G_SendCommandThread.sG_OperationCommandType := sG_OperationCommandType;
}
{
      G_SendCommandThread.bG_CommandLog := bG_CommandLog;
      G_SendCommandThread.bG_ErrorLog := bG_ErrorLog;
      G_SendCommandThread.bG_ResponseLog := bG_ResponseLog;
}
{
      G_SendCommandThread.sG_BroadcastStartDate := sG_BroadcastStartDate;
      G_SendCommandThread.sG_BroadcastEndDate := sG_BroadcastEndDate;
}
    end;


    bG_StartFetching := false;
{
    G_Timer:= TTimer.Create(nil);

    G_Timer.Interval := reDefineFetchDataInterval(now);
    G_Timer.OnTimer := FetchCmdDataTimer;
    G_Timer.Enabled := false;
}
end;

procedure TfrmMain.chbStartFunClick(Sender: TObject);
begin
    bG_RunFetchCommandThread :=  chbStartFun.Checked;
    bG_RunReceivingCommandThread := chbStartFun.Checked;

    bG_RunSendingCommandThread := chbStartFun.Checked;
    bG_HandleResponseThread := chbStartFun.Checked;
    bG_HandleUIThread := chbStartFun.Checked;


//    G_Timer.Enabled :=  chbStartFun.Checked;


    if chbStartFun.Checked then
    begin
      buildSendCmdConnection;


    end;
    Timer1.Enabled := chbStartFun.Checked;
    Timer2.Enabled := chbStartFun.Checked;
    Timer3.Enabled := chbStartFun.Checked;
end;

{
procedure TfrmMain.updateCmdStatusToProcessing(sI_SeqNos : String);
var
    sL_SQL : String;
begin

    if sI_SeqNos='' then Exit;
    with qryTmp do
    begin
      sL_SQL := 'update SEND_NAGRA set CMD_STATUS=' + STR_SEP + 'P' + STR_SEP;
      sL_SQL := sL_SQL + ',RESENTTIMES=NVL(RESENTTIMES,0)+1 ';
      sL_SQL := sL_SQL +' where SEQNO in (' + sI_SeqNos + ') and CMD_STATUS not in (' + STR_SEP + 'C' + STR_SEP +',' + STR_SEP + 'E' + STR_SEP + ')'  ;
      SQL.Clear;
      SQL.Add(sL_SQL);
      ExecSQL;
      Close;
    end;
end;
}
procedure TfrmMain.SendCmdClntSocketError(Sender: TObject;
  Socket: TCustomWinSocket; ErrorEvent: TErrorEvent;
  var ErrorCode: Integer);
begin
    Showmessage(getLanguageSettingInfo('frmMain_Msg_29_Content'));
end;

procedure TfrmMain.FetchCmdDataThreadDone(Sender: TObject);
begin
    if (Sender as TThread).FatalException <> nil then
    begin
//      showmessage('111');
      saveThreadInfo('Fetch Command Data Thread Done.');
    end;
end;

procedure TfrmMain.HandleResponseThreadDone(Sender: TObject);
begin
    if (Sender as TThread).FatalException <> nil then
    begin
//      showmessage('222');
      saveThreadInfo('Handle Response Thread Done.');
    end;


end;

procedure TfrmMain.SendCommandThreadDone(Sender: TObject);
begin
//    self.G_SendCommandThread.G_SendCmdWStream.Free;  //0110

    if (Sender as TThread).FatalException <> nil then
    begin
//      showmessage('333');
      saveThreadInfo('Send Command Thread Done.');
    end;


end;

procedure TfrmMain.ActiveCDS(I_CDS: TClientDataSet; sI_RefFileName,
  sI_DefaultStatus: String);
var
    sL_Path : String;
begin
//    SetCurrentDir(ExtractFilePath(Application.ExeName));

    sL_Path := ExtractFilePath(Application.ExeName);

    if (I_CDS.State = dsInactive) then
      I_CDS.CreateDataSet;
    if (FileExists(sL_Path + sI_RefFileName)) then
      I_CDS.LoadFromFile(sL_Path + sI_RefFileName);

    if (I_CDS.RecordCount=0) then
      I_CDS.append
    else
      I_CDS.edit;

    if (sI_DefaultStatus='E') then //edit mode
    begin
    end
    else if (sI_DefaultStatus='B') then//browse mode
    begin
      I_CDS.Post;
    end;

end;


function TfrmMain.getCompNetworkInfo(nI_CompCode:Integer; sI_IrmCmdID:String): String;
var

    nL_Ndx, nL_NetworkID:Integer;
    sL_Result, sL_NetworkId, sL_Operation : String;
begin
    sL_Result := '';
    nL_Ndx := G_NetworkIdInfoStrList.IndexOf(IntToStr(nI_CompCode));

    if nL_Ndx<>-1 then
    begin

      sL_NetworkId := (G_NetworkIdInfoStrList.Objects[nL_Ndx] as TNetworkIDInfoObj).NetworkID;
      nL_NetworkID := StrToInt(sL_NetworkId);
      sL_NetworkId := IntToHex(nL_NetworkID,0);
      sL_NetworkId := TUstr.AddString(sL_NetworkId,'0',false,4);

      if sI_IrmCmdID=NETWORK_OPERATION_194 then  //194=>operation 要為'', 且 data length=2;
        sL_Operation := ''
      else
      begin  //195=>operation表示多久要 reset STB, data length=4
        sL_Operation := (G_NetworkIdInfoStrList.Objects[nL_Ndx] as TNetworkIDInfoObj).Operation;
        sL_Operation := TUstr.AddString(sL_Operation,'0',false,4);
      end;
      sL_Result := sL_NetworkId + sL_Operation;

    end;


    result := sL_Result;
end;

{
function TfrmMain.getCompNetworkInfo(nI_CompCode:Integer; sI_IrmCmdID:String): String;
var
    nL_NetworkID:Integer;
    sL_Result, sL_NetworkId, sL_Operation : String;
begin
    sL_Result := '';
    with cdsNetworkID do
    begin
      if recordCount>0 then
      begin
        if Locate('CompCode',nI_CompCode, [loPartialKey]) then
        begin
          sL_NetworkId := FieldByName('NetworkID').AsString;
          nL_NetworkID := StrToInt(sL_NetworkId);
          sL_NetworkId := IntToHex(nL_NetworkID,0);
          sL_NetworkId := TUstr.AddString(sL_NetworkId,'0',false,4);

          if sI_IrmCmdID=NETWORK_OPERATION_194 then  //194=>operation 要為'', 且 data length=2;
            sL_Operation := ''
          else
          begin  //195=>operation表示多久要 reset STB, data length=4
            sL_Operation := FieldByName('Operation').AsString;
            sL_Operation := TUstr.AddString(sL_Operation,'0',false,4);
          end;
          sL_Result := sL_NetworkId + sL_Operation;
        end;
      end;

    end;
    result := sL_Result;
end;
}
procedure TfrmMain.FetchCmdDataTimer(Sender: TObject);
begin
//    showmessage('timer...');
    bG_StartFetching := true;
end;

function TfrmMain.getConnectionID(nI_CompCode: Integer): Integer;
var
    nL_Ndx : Integer;
begin
    if Assigned(G_CompCodeStrList) then
      nL_Ndx := G_CompCodeStrList.IndexOf(IntToStr(nI_CompCode))
    else
      nL_Ndx := -1;

    result := nL_Ndx;

end;

procedure TfrmMain.getAppInfo(var sI_AppPath, sI_AppId: String);
begin
    sI_AppPath := Application.ExeName;
    sI_AppId :=  IntToStr(GetCurrentProcessId);
end;

procedure TfrmMain.HandleUIThreadDone(Sender: TObject);
begin
    if (Sender as TThread).FatalException <> nil then
    begin
//      showmessage('555');
      saveThreadInfo('Handle UI Thread Done.');


    end;
end;

function TfrmMain.getVirtualNotes(sI_HighLevelCmdID: String): String;
var
    ii,jj : Integer;
    sL_BDate, sL_EDate, sL_Notes : String;
begin
    if (sI_HighLevelCmdID='B1') or (sI_HighLevelCmdID='B6') or (sI_HighLevelCmdID='B9') then
    begin
      for ii:=0 to memProductID.Lines.Count -1 do
      begin
        if ii=0 then
          sL_Notes := 'A~' + memProductID.Lines.Strings[ii]
        else
          sL_Notes := sL_Notes + ',' + 'A~' + memProductID.Lines.Strings[ii];
      end;
    end
    else if (sI_HighLevelCmdID='B2')or (sI_HighLevelCmdID='B4') or (sI_HighLevelCmdID='B5')  then
    begin
      for ii:=0 to memProductID.Lines.Count -1 do
      begin
        if ii=0 then
          sL_Notes := memProductID.Lines.Strings[ii]
        else
          sL_Notes := sL_Notes + ',' + memProductID.Lines.Strings[ii];
      end;
    end
    else if (sI_HighLevelCmdID='B7') then
    begin
      sL_BDate := TUdateTime.GetPureDateStr(dtpBDate.date,'');
      sL_EDate := TUdateTime.GetPureDateStr(dtpEDate.date,'');
      for ii:=0 to memProductID.Lines.Count -1 do
      begin
        if ii=0 then
          sL_Notes := memProductID.Lines.Strings[ii] + '^' + sL_BDate + '^' + sL_EDate
        else
          sL_Notes := sL_Notes + ',' + memProductID.Lines.Strings[ii] + '^' + sL_BDate + '^' + sL_EDate;
      end;
      for jj:=0 to memIccNo.Lines.Count -1 do
      begin
        if jj=0 then
          sL_Notes := sL_Notes + '~' + memIccNo.Lines.Strings[jj]
        else
          sL_Notes := sL_Notes + ';' + memIccNo.Lines.Strings[jj]
      end;
//      Ex: 111111111111^20020801^20020901,222222222222^20020901^20021231~3333333333;4444444444
    end
    else if (sI_HighLevelCmdID='C1') then
    begin
      for jj:=0 to memIccNo.Lines.Count -1 do
      begin
        if jj=0 then
          sL_Notes := memIccNo.Lines.Strings[jj]
        else
          sL_Notes := sL_Notes + ',' + memIccNo.Lines.Strings[jj]
      end;

    end
    else if (sI_HighLevelCmdID='E2') then
    begin
      sL_Notes := 'Hi, This is Howard.';
    end    
    else
      sL_Notes := '';

    result := sL_Notes;
end;

function TfrmMain.reDefineFetchDataInterval(dI_Time: TDate): Integer;
var
    nL_Result, nL_Time1, nL_Time2, nL_Time3 : Integer;
    sL_Time1, sL_Time2 : String;
    wL_Hour, wL_Min, wL_Sec, wL_MSec : word;
begin
    sL_Time1 := Copy(sG_TimeDefine,1,4);
    sL_Time2 := Copy(sG_TimeDefine,6,4);
    DecodeTime(dI_Time,wL_Hour,wL_Min, wL_Sec, wL_MSec);
    nL_Time1 := StrToInt(sL_Time1);
    nL_Time2 := StrToInt(sL_Time2);
    nL_Time3 := wL_Hour*100+wL_Min;

    if (nL_Time3>nL_Time1) and (nL_Time3<nL_Time2) then
      nL_Result := nG_CmdRefreshRate1*1000 //傳入之時間為尖峰時段
    else
      nL_Result := nG_CmdRefreshRate2*1000; //傳入之時間為離峰時段

    result := nL_Result;
end;

procedure TfrmMain.saveThreadInfo(sI_Info: String);
var
    L_ThreadInfoStrList : TStringList;
    sL_ExeFileName, sL_ExePath, sL_ThreadInfoFileName : String;
begin
    sL_ExeFileName := Application.ExeName;
    sL_ExePath := ExtractFileDir(sL_ExeFileName);
    sL_ThreadInfoFileName := sL_ExePath + '\' + THREAD_INFO_FILE;

    L_ThreadInfoStrList := TStringList.Create;    
    if FileExists(sL_ThreadInfoFileName) then
      L_ThreadInfoStrList.LoadFromFile(sL_ThreadInfoFileName);
    if L_ThreadInfoStrList.Count>500 then
      L_ThreadInfoStrList.Clear;
    L_ThreadInfoStrList.Insert(0,sI_Info + ':' + DateTimeToStr(now));
    L_ThreadInfoStrList.SaveToFile(sL_ThreadInfoFileName);
    L_ThreadInfoStrList.Free;
end;

procedure TfrmMain.SendCmdClntSocketDisconnect(Sender: TObject;
  Socket: TCustomWinSocket);
var
    L_SocketInfoStrList : TStringList;
    sL_ExeFileName, sL_ExePath, sL_ThreadInfoFileName : String;
begin
    sL_ExeFileName := Application.ExeName;
    sL_ExePath := ExtractFileDir(sL_ExeFileName);
    sL_ThreadInfoFileName := sL_ExePath + '\' + SOCKET_INFO_FILE;

    L_SocketInfoStrList := TStringList.Create;    
    if FileExists(sL_ThreadInfoFileName) then
      L_SocketInfoStrList.LoadFromFile(sL_ThreadInfoFileName);
    if L_SocketInfoStrList.Count>500 then
      L_SocketInfoStrList.Clear;
    L_SocketInfoStrList.Insert(0,getLanguageSettingInfo('frmMain_Msg_30_Content') + DateTimeToStr(now));
    L_SocketInfoStrList.SaveToFile(sL_ThreadInfoFileName);
    L_SocketInfoStrList.Free;
end;


function TfrmMain.getReturnPathPort: Integer;
begin
    result := nG_RPortNo;
end;

function TfrmMain.getReturnPathServerIP: STring;
begin
    result := sG_ServerIP;
end;

procedure TfrmMain.ReturnPathSvrThreadDone(Sender: TObject);
begin
    if (Sender as TThread).FatalException <> nil then
    begin
//      showmessage('777');
      saveThreadInfo('Return Path Server Thread Done.');
    end;

end;

procedure TfrmMain.setTransCaption;
begin
    frmMain_Pnl_1.Caption := getLanguageSettingInfo('frmMain_Pnl_1_Caption_2');
//    frmMain_Pnl_1.Caption := ' 即時監控--指令傳送中';
    if frmMain_Pnl_1.Font.Color = clBlue then
      frmMain_Pnl_1.Font.Color := clRed
    else
      frmMain_Pnl_1.Font.Color := clBlue;


    self.Invalidate;
end;


procedure TfrmMain.showLogOnMemo(nI_ActiveConn: Integer; sI_Detail:String);
var
    bL_KillSelf, bL_ShowMemo : boolean;
    sL_DbAlias, sL_Msg : String;
    L_DbLogStrList : TStringList;
    sL_ExeFileName, sL_ExePath, sL_RealDbLogFileName : String;
    L_Intf : IMonitor;
    procedure writeMsg;
    begin
      if bL_ShowMemo then
      begin
        if memErrorLog.Lines.Count > MAX_SHOW_MEMO_ERROR_COUNT then
          memErrorLog.Lines.Clear;
        memErrorLog.Lines.insert(0,sL_Msg );
      end;
      //down, 將訊息寫到檔案中
      sL_ExeFileName := Application.ExeName;
      sL_ExePath := ExtractFileDir(sL_ExeFileName)+'\CaLog';
      if not DirectoryExists(sL_ExePath) then
        CreateDir(sL_ExePath);
      sL_RealDbLogFileName := sL_ExePath + '\' + DB_LOG_FILE_NAME;

      L_DbLogStrList := TStringList.Create;

      if FileExists(sL_RealDbLogFileName) then
        L_DbLogStrList.LoadFromFile(sL_RealDbLogFileName);
      if L_DbLogStrList.Count > MAX_SHOW_MEMO_ERROR_COUNT then
        L_DbLogStrList.Clear;
      L_DbLogStrList.insert(0,sL_Msg);
      L_DbLogStrList.SaveToFile(sL_RealDbLogFileName);
      L_DbLogStrList.Free;
      //up, 將訊息寫到檔案中
    end;
begin
    try
      bL_KillSelf := false;
      bL_ShowMemo := true;
      
      if nI_ActiveConn>=0 then
      begin
        sL_DbAlias := G_DbAliasStrList.Strings[nI_ActiveConn];
        sL_Msg := '['+ datetimetostr(now)+'] ' + sL_DbAlias + getLanguageSettingInfo('frmMain_Msg_31_Content') + getLanguageSettingInfo('frmMain_Msg_32_Content') + sI_Detail;
        writeMsg;
      end
      else if nI_ActiveConn=SEND_CMD_ERROR_FLAG then
      begin
        bL_KillSelf := true;
        sL_DbAlias := '';
        sL_Msg := '['+ datetimetostr(now)+ '] ' + getLanguageSettingInfo('frmMain_Msg_33_Content') + getLanguageSettingInfo('frmMain_Msg_32_Content') + sI_Detail ;

        writeMsg;
        releaseResource;
        //down, 呼叫監控員, 讓監控員把 CA 叫起來
        L_Intf := CoMonitor.Create;
        L_Intf.RunCa(sG_NewIniFileName);
        L_Intf := nil;
        //up, 呼叫監控員, 讓監控員把 CA 叫起來
      end
      else if nI_ActiveConn=REBUILD_NAGRA_CONNECTION_FAILED then
      begin
        bL_KillSelf := true;
        sL_DbAlias := '';
        sL_Msg := '['+ datetimetostr(now)+ '] ' + getLanguageSettingInfo('frmMain_Msg_34_Content') + getLanguageSettingInfo('frmMain_Msg_32_Content') + sI_Detail ;

        writeMsg;
        releaseResource;
        //down, 呼叫監控員, 讓監控員把 CA 叫起來
        L_Intf := CoMonitor.Create;
        L_Intf.RunCa(sG_NewIniFileName);
        L_Intf := nil;
        //up, 呼叫監控員, 讓監控員把 CA 叫起來
      end
      else if nI_ActiveConn=SOCKET_NOT_READY_FOR_READING_FLAG then
      begin
        bL_KillSelf := true;
        sL_DbAlias := '';
        sL_Msg := '['+ datetimetostr(now)+ '] ' +'Socket is not ready for reading.' + getLanguageSettingInfo('frmMain_Msg_32_Content') + sI_Detail ;

        writeMsg;
        releaseResource;
        //down, 呼叫監控員, 讓監控員把 CA 叫起來
        L_Intf := CoMonitor.Create;
        L_Intf.RunCa(sG_NewIniFileName);
        L_Intf := nil;
        //up, 呼叫監控員, 讓監控員把 CA 叫起來


      end

      else if nI_ActiveConn=EXECUTE_ORACLE_SP_ERROR_FLAG then
      begin
        bL_KillSelf := true;
        sL_DbAlias := '';
        sL_Msg := '['+ datetimetostr(now)+ '] ' + getLanguageSettingInfo('frmMain_Msg_35_Content')+ getLanguageSettingInfo('frmMain_Msg_32_Content') + sI_Detail ;

        writeMsg;
        releaseResource;
        //down, 呼叫監控員, 讓監控員把 CA 叫起來
        L_Intf := CoMonitor.Create;
        L_Intf.RunCa(sG_NewIniFileName);
        L_Intf := nil;
        //up, 呼叫監控員, 讓監控員把 CA 叫起來


      end

      else if nI_ActiveConn=REBUILD_SEND_CMD_CONNECTION_TO_NAGRA then
      begin
        bL_ShowMemo := false;

        sL_DbAlias := '';
        sL_Msg := '['+ datetimetostr(now)+ '] ' +  getLanguageSettingInfo('frmMain_Msg_36_Content')+ getLanguageSettingInfo('frmMain_Msg_32_Content') + sI_Detail;
        writeMsg;
      end
      else if nI_ActiveConn=UPDATE_CMD_STATUS_ERROR_FLAG then
      begin
        sL_DbAlias := '';
        sL_Msg := '['+ datetimetostr(now)+ '] ' + getLanguageSettingInfo('frmMain_Msg_37_Content')+ getLanguageSettingInfo('frmMain_Msg_32_Content') + sI_Detail;
        writeMsg;
      end
      else if nI_ActiveConn=RE_CONNECT_NAGRA_ERROR_FLAG then
      begin
        sL_DbAlias := '';
        sL_Msg := '['+ datetimetostr(now)+ '] ' + getLanguageSettingInfo('frmMain_Msg_38_Content')+ getLanguageSettingInfo('frmMain_Msg_32_Content') + sI_Detail;
        writeMsg;
      end
      else if nI_ActiveConn=CLOSE_SEND_NAGRA_FLAG then
      begin
        sL_DbAlias := '';
        sL_Msg := '['+ datetimetostr(now)+']' +getLanguageSettingInfo('frmMain_Msg_32_Content') + sI_Detail;
        writeMsg;
      end
      else if nI_ActiveConn=HANDLE_RESPONSE_ERROR_FLAG then
      begin
        sL_DbAlias := '';
        sL_Msg := '['+ datetimetostr(now)+ '] ' + getLanguageSettingInfo('frmMain_Msg_39_Content')+ getLanguageSettingInfo('frmMain_Msg_32_Content') + sI_Detail;
        writeMsg;
      end
      else if nI_ActiveConn=WRITE_TO_MIS_ERROR_FLAG then
      begin
        sL_DbAlias := '';
        sL_Msg := '['+ datetimetostr(now)+ '] ' + getLanguageSettingInfo('frmMain_Msg_40_Content')+ getLanguageSettingInfo('frmMain_Msg_32_Content') + sI_Detail;
        writeMsg;
      end
      else if nI_ActiveConn=RECEIVE_CMD_RESPONSE_ERROR_FLAG then
      begin
        sL_DbAlias := '';
        sL_Msg := '['+ datetimetostr(now)+ '] ' + getLanguageSettingInfo('frmMain_Msg_41_Content')+ getLanguageSettingInfo('frmMain_Msg_32_Content') + sI_Detail;
        writeMsg;

        releaseResource;
        //down, 呼叫監控員, 讓監控員把 CA 叫起來
        L_Intf := CoMonitor.Create;
        L_Intf.RunCa(sG_NewIniFileName);
        L_Intf := nil;
        //up, 呼叫監控員, 讓監控員把 CA 叫起來
      end
      else if nI_ActiveConn=EXECUTE_SEND_NAGRA_FLAG then
      begin
        sL_DbAlias := '';
        sL_Msg := '['+ datetimetostr(now)+']' + getLanguageSettingInfo('frmMain_Msg_32_Content') + sI_Detail + getLanguageSettingInfo('frmMain_Msg_42_Content')+ sG_LoginName;
        writeMsg;
      end;

      if bL_KillSelf then //未必會被執行到, 因為呼叫  L_Intf.RunCa 時, 就會把 ca command gateway kill 掉
        closeProg;
  //      Application.Terminate;
    except

    end;
end;

procedure TfrmMain.createLivItem; //0219
var
    ii, jj : Integer;
    L_ListItem : TListItem;
    L_Splitter: TSplitter;
begin
    if bG_ShowUI then
    begin
      nG_LivItemPos := 0;
      memErrorLog.Align := alBottom;
      memErrorLog.Height := 122;

      L_Splitter:= TSplitter.Create(Application);
      L_Splitter.Parent := frmMain_Tab_1;
      L_Splitter.Height := 4;
      L_Splitter.Align := alBottom;


      livStatus.Visible := true;
      livStatus.Align := alClient;

      frmMain.livStatus.Items.Clear;
      for ii:=0 to CA_UI_MAX_COUNT -1 do
      begin
        L_ListItem := frmMain.livStatus.Items.Add;

        for jj:=0 to LISTVIEW_SUB_ITEM_COUNT-1 do
         L_ListItem.SubItems.Add('');

      end;
    end
    else
    begin
      livStatus.Visible := false;
      memErrorLog.Align := alClient;
    end;
end;

procedure TfrmMain.closeProg;
begin
    showLogOnMemo(CLOSE_SEND_NAGRA_FLAG,getLanguageSettingInfo('frmMain_Msg_43_Content'));
    releaseResource;

//      Application.Terminate;
    Close;
end;

procedure TfrmMain.releaseResource;
begin
    closeSocket;

    chbStartFun.Checked := false;

    bG_RunFetchCommandThread := false;
    bG_RunReceivingCommandThread := false;
    bG_RunSendingCommandThread := false;
    bG_HandleResponseThread := false;
    bG_HandleUIThread := false;


    if Assigned(G_FetchComDataThread) then
      G_FetchComDataThread.Terminate;

    if Assigned(G_SendCommandThread) then
      G_SendCommandThread.Terminate;

    if Assigned(G_ReceivingCommandResponseThread) then
      G_ReceivingCommandResponseThread.Terminate;

    if Assigned(G_THandleResponse) then
      G_THandleResponse.Terminate;

    if Assigned(G_HandleUIThread) then
      G_HandleUIThread.Terminate;

    { 加此段..程式會當掉
    if Assigned(G_HandleUIThread) then
      G_HandleUIThread.Suspend;
    }

    if Assigned(G_DisConnDbStrList) then
    begin
      G_DisConnDbStrList.Free;
      G_DisConnDbStrList := nil;
    end;

    {
    if Assigned(G_SendCmdClntSocket) then
    begin
      G_SendCmdClntSocket.Close;
      G_SendCmdClntSocket := nil;
    end;
    }
    if Assigned(G_ResponseMsgStrList) then
    begin
      G_ResponseMsgStrList.Free;
      G_ResponseMsgStrList := nil;
    end;

    if Assigned(G_CompCodeStrList) then
    begin
      G_CompCodeStrList.Free;
      G_CompCodeStrList := nil;
    end;


    if Assigned(G_DbAliasStrList) then
    begin
      G_DbAliasStrList.Free;
      G_DbAliasStrList := nil;
    end;


    if Assigned(G_NetworkIdInfoStrList) then
    begin
      G_NetworkIdInfoStrList.Free;
      G_NetworkIdInfoStrList := nil;
    end;


    if Assigned(G_CommandInfoStrList) then
    begin
      G_CommandInfoStrList.Free;
      G_CommandInfoStrList := nil;
    end;


    if Assigned(G_TimeStampStrList) then
    begin
      G_TimeStampStrList.Free;
      G_TimeStampStrList := nil;
    end;

    if Assigned(G_LanguageSettingIni) then
    begin
      G_LanguageSettingIni.Free;
    end;  
end;

procedure TfrmMain.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
    sL_EncryptedLicenseKey : STring;
begin
    if (ssCtrl in Shift) and (Key=83) then //Ctrl + s
    begin
      bG_ShowUI := not bG_ShowUI;
      createLivItem;
    end
    else if (ssShift in Shift) and (Key=192) then //Shift + ~
    begin
      try
        sL_EncryptedLicenseKey := InputBox(getLanguageSettingInfo('frmMain_Msg_2_Content'), 'License Key:', '');
        if sL_EncryptedLicenseKey='' then
        begin
          exit;
        end
        else if writeRegData(sL_EncryptedLicenseKey) then
        begin
          MessageDlg(getLanguageSettingInfo('frmMain_Msg_3_Content'), mtInformation, [mbOK],0);
          exit;
        end
        else
        begin
          MessageDlg(getLanguageSettingInfo('frmMain_Msg_44_Content'), mtError, [mbOK],0);
          Exit;
        end;
      except

      end;
    end;
end;


procedure TfrmMain.buildSendCmdConnection;
begin
      //down, 建立用來 send command 的 connection..
      if not bG_HasBuildSendCmdConnection then
      begin
        if Assigned(G_SendCommandThread) then
//         (G_SendCommandThread as TSendCommand).buildSendCmdConnection(G_SendCmdClntSocket, G_SendCmdWStream);
//         (G_SendCommandThread as TSendCommand).buildSendCmdConnection(G_SendCmdClntSocket, G_SendCmdWStream, G_ReceiveCmdWStream);
         (G_SendCommandThread as TSendCommand).buildSendCmdConnection(G_SendCmdClntSocket);
      end;
      //up, 建立用來 send command 的 connection..

end;

procedure TfrmMain.reBuildNagraConn;
begin
    try

      bG_HasBuildSendCmdConnection := false;

      closeSocket;
      sleep(12000);
//        sleep(3000);
      buildSendCmdConnection;
    except
      frmMain.showLogOnMemo(REBUILD_NAGRA_CONNECTION_FAILED,' frmMain.reBuildNagraConn .' + getLanguageSettingInfo('frmMain_Msg_45_Content')); //執行完此行程式後, 程式會結束掉
    end;
end;

function TfrmMain.reverseBinaryVal(I_TmpAry: array of Char): Integer;
var
    sL_TmpStr : String;
    L_TmpChar : Char;
    ii : Integer;
begin
    sL_TmpStr := '';
    for ii:=0 to High(I_TmpAry) do
    begin
      L_TmpChar := I_TmpAry[ii];
      sL_TmpStr := sL_TmpStr + IntToStr(ord(L_TmpChar));
    end;
    result := StrToInt(sL_TmpStr);

end;

function TfrmMain.transReceivedDataCharAryToStr(I_TmpAry: array of Char; nI_DataLength:Integer): String;
var
    nL_Seed, nL_TmpLength, nL_LoopCount : Integer;
    sL_Result : String;
    L_TmpAry : array[0..1] of Char;
    function copyAryToStr(I_SrcAry:array of Char; nI_BeginPos, nI_Count:Integer):String;
    var
        ii : Integer;
        sL_TmpResult : String;
    begin
      sL_TmpResult := '';
      for ii:=0 to nI_Count-1 do
        sL_TmpResult := sL_TmpResult + I_SrcAry[nI_BeginPos+ii];
      result := sL_TmpResult;  
    end;
begin
    nL_Seed := 0;
    nL_LoopCount := 0;
    sL_Result := '';
    while true do
    begin
      Inc(nL_LoopCount);
      L_TmpAry[0] := I_TmpAry[nL_Seed];
      L_TmpAry[1] := I_TmpAry[nL_Seed+1];
      nL_TmpLength := reverseBinaryVal(L_TmpAry);

      if sL_Result='' then
        sL_Result := copyAryToStr(I_TmpAry, nL_Seed+2, nL_TmpLength)
      else
        sL_Result := sL_Result + RECEIVED_DATA_SEP_FLAG + copyAryToStr(I_TmpAry, nL_Seed+2, nL_TmpLength);
      nL_Seed := length(sL_Result)+2*nL_LoopCount-(nL_LoopCount-1);
//      if (I_TmpAry[nL_Seed]=char(0)) and (I_TmpAry[nL_Seed+1]=char(0)) then
      if nL_Seed>=nI_DataLength then
        break;

    end;
    result :=sL_Result;
end;

procedure TfrmMain.Button1Click(Sender: TObject);
begin
    Memo1.Lines.Add(inttostr(TUstr.binToInt('1010')));
    Memo1.Lines.Add(inttostr(TUstr.binToInt('000000000100000101000001')));
end;

procedure TfrmMain.closeSocket;
begin
//showmessage('1');
{
    if Assigned(G_SendCmdWStream) then
    begin
      G_SendCmdWStream.Free;
      G_SendCmdWStream := nil;
    end;

    if Assigned(G_ReceiveCmdWStream) then
    begin
      G_ReceiveCmdWStream.Free;
      G_ReceiveCmdWStream := nil;
    end;
}

    if Assigned(G_SendCmdClntSocket) and (G_SendCmdClntSocket.Active) then
      G_SendCmdClntSocket.Close;
//showmessage('2');
end;

procedure TfrmMain.createThreadForNagra(nI_Operation:Integer; sI_ErrorMsg:String);
var
    sL_Info : String;
    procedure initThreadForNagra;
    begin
      if Assigned(G_SendCommandThread) then
      begin
        G_SendCommandThread.Terminate;
        G_SendCommandThread.Free;
      end;  

      G_SendCommandThread := TSendCommand.Create;
      if GetExitCodeThread(G_SendCommandThread.ThreadID,G_SendCommandThread_ExitCode) then
      begin
        if not bG_AutoInvoke then
          MessageDlg(getLanguageSettingInfo('frmMain_Msg_46_Content') + sL_Info, mtError, [mbOK],0);
        Application.Terminate;
        Exit;
      end;


      if Assigned(G_ReceivingCommandResponseThread) then
      begin
        G_ReceivingCommandResponseThread.Terminate;
        G_ReceivingCommandResponseThread.Free;
      end;

      G_ReceivingCommandResponseThread := ReceivingCommandResponseThread.Create(false);
      if GetExitCodeThread(G_ReceivingCommandResponseThread.ThreadID,G_ReceivingCommandResponseThread_ExitCode) then
      begin
        if not bG_AutoInvoke then
          MessageDlg(getLanguageSettingInfo('frmMain_Msg_47_Content') + sL_Info, mtError, [mbOK],0);

        Application.Terminate;
        Exit;
      end;

    end;
begin

    sL_Info := getLanguageSettingInfo('frmMain_Msg_48_Content');
    closeSocket;

    case nI_Operation of
      1:
       begin
         initThreadForNagra;
       end;
      { 
      2:
       begin
         initThreadForNagra;

        if AnsiPos('網路名稱無法使用', sI_ErrorMsg) >0 then
        begin
          showLogOnMemo(REBUILD_SEND_CMD_CONNECTION_TO_NAGRA,sI_ErrorMsg + '--' + datetimetostr(now));

          reBuildNagraConn;

        end
        else if AnsiPos('violation', sI_ErrorMsg) >0 then
        begin
        end
        else
          showLogOnMemo(RECEIVE_CMD_RESPONSE_ERROR_FLAG,' createThreadForNagra.程式即將結束,然後重新啟動.' +'--錯誤訊息:' + sI_ErrorMsg);

       end;
      }
    end;


end;

procedure TfrmMain.transNetworkIdInfo;
var
    sL_CompCode : STring;
    L_Obj : TNetworkIDInfoObj;
begin
    if not assigned(G_NetworkIdInfoStrList) then
      Exit;
    G_NetworkIdInfoStrList.Clear;

    with dtmMain.cdsNetworkID do
    begin
      First;
      while not Eof do
      begin
        sL_CompCode := FieldByName('COMPCODE').AsString;
        L_Obj := TNetworkIDInfoObj.Create;
        L_Obj.NetworkID := FieldByName('NETWORKID').AsString;
        L_Obj.Operation := FieldByName('OPERATION').AsString;
        G_NetworkIdInfoStrList.AddObject(sL_CompCode, L_Obj);
        Next;
      end;
    end;
end;

function TfrmMain.GetRegistryInfo: boolean;
var
    L_StrList : TStringList;
    sL_EncryptedLicenseKey, sL_LicenseKey : STring;

    bL_Result : boolean;
    vL_Result : variant;


    procedure readRegData;
    begin
      vL_Result := TUother.RegRead(REGISTRY_ROOT,'SYS1',3); //軟體之使用到期日
        dG_Sys1 := StrToDate(VarToStr(vL_Result));

      vL_Result := TUother.RegRead(REGISTRY_ROOT,'SYS2',3); //軟體之安裝日
        dG_Sys2 := StrToDate(VarToStr(vL_Result));

      vL_Result := TUother.RegRead(REGISTRY_ROOT,'SYS3',3); //軟體之上次使用日期
        dG_Sys3 := StrToDate(VarToStr(vL_Result));

      vL_Result := TUother.RegRead(REGISTRY_ROOT,'SYS4',2); //是否啟動軟體保護機制
        sG_Sys4 := VarToStr(vL_Result);

      vL_Result := TUother.RegRead(REGISTRY_ROOT,'SYS5',2); //是否停用本軟體
        sG_Sys5 := VarToStr(vL_Result);


    end;
begin
    try
      bL_Result := true;
      readRegData;
    except
      sL_EncryptedLicenseKey := InputBox(getLanguageSettingInfo('frmMain_Msg_2_Content'), 'License Key:', '');
      if sL_EncryptedLicenseKey='' then
      begin
        bL_Result := false;
      end  
      else if writeRegData(sL_EncryptedLicenseKey) then
      begin
        try
          readRegData;
        except
          bL_Result := false;
        end;
      end
      else
        bL_Result := false;
    end;

    result := bL_Result;

end;

function TfrmMain.writeRegData(sI_EncryptedLicenseKey: STring): boolean;
var
  L_StrList : TStringList;
  sL_LicenseKey : String;
  bL_WriteRegResult : boolean;
begin
  bL_WriteRegResult := true;
  sL_LicenseKey := TUother.myDecrypt(sI_EncryptedLicenseKey, ENCRYPTION_KEY);

  L_StrList := TUstr.ParseStrings(sL_LicenseKey,ENCRYPTION_SEP);
  if (L_StrList.Count=3) and (TUdateTime.IsDateStr(L_StrList.Strings[0],'/'))
    and ((L_StrList.Strings[1]=SYS4_Y) or (L_StrList.Strings[1]=SYS4_N))
    and ((L_StrList.Strings[2]=SYS5_Y) or (L_StrList.Strings[2]=SYS5_N)) then
  begin
    //表示 License Key 格式正確...
    TUother.RegWrite(REGISTRY_ROOT,'SYS1',3,StrToDate(L_StrList.Strings[0])); //寫入軟體之使用到期日
    TUother.RegWrite(REGISTRY_ROOT,'SYS2',3,Date); //寫入軟體之安裝日
    TUother.RegWrite(REGISTRY_ROOT,'SYS3',3,Date); //寫入軟體之上次使用日期
    TUother.RegWrite(REGISTRY_ROOT,'SYS4',2,L_StrList.Strings[1]); //寫入是否啟動軟體保護機制
    TUother.RegWrite(REGISTRY_ROOT,'SYS5',2,L_StrList.Strings[2]); //寫入是否停用本軟體

   end
   else
     bL_WriteRegResult := false;
   L_StrList.Free;
   result := bL_WriteRegResult;
end;

procedure TfrmMain.Timer1Timer(Sender: TObject);
begin
    (Sender As TTimer).Enabled := false;

    if Assigned(G_FetchComDataThread) then
    begin
      Timer1.Interval := reDefineFetchDataInterval(now);
      frmMain.bG_StartFetching := true;

      G_FetchComDataThread.onTimerStartAction;
    end;
    (Sender As TTimer).Enabled := true;
end;

procedure TfrmMain.Timer2Timer(Sender: TObject);
begin
    (Sender As TTimer).Enabled := false;
    if Assigned(G_THandleResponse) then
      G_THandleResponse.onTimerStartAction;
    (Sender As TTimer).Enabled := true;

end;

procedure TfrmMain.Timer3Timer(Sender: TObject);
begin
    (Sender As TTimer).Enabled := false;
    if Assigned(G_SendCommandThread) then
      G_SendCommandThread.onTimerStartAction;
    (Sender As TTimer).Enabled := true;

end;

procedure TfrmMain.setLanguageString;
begin


    frmMain_Tab_1.Caption := getLanguageSettingInfo('frmMain_Tab_1_Caption');
    frmMain_Tab_2.Caption := getLanguageSettingInfo('frmMain_Tab_2_Caption');
    frmMain_Lbl_1.Caption := getLanguageSettingInfo('frmMain_Lbl_1_Caption');
    frmMain_Lbl_2.Caption := getLanguageSettingInfo('frmMain_Lbl_2_Caption');
    frmMain_Lbl_3.Caption := getLanguageSettingInfo('frmMain_Lbl_3_Caption');
    frmMain_Lbl_4.Caption := getLanguageSettingInfo('frmMain_Lbl_4_Caption');
    frmMain_Lbl_5.Caption := getLanguageSettingInfo('frmMain_Lbl_5_Caption');
    frmMain_Lbl_6.Caption := getLanguageSettingInfo('frmMain_Lbl_6_Caption');

    frmMain_Pnl_1.Caption := getLanguageSettingInfo('frmMain_Pnl_1_Caption');
    chbStartFun.Caption := getLanguageSettingInfo('chbStartFun_Caption');
    btnExit.Caption := getLanguageSettingInfo('btnExit_Caption');

    mniOption.Caption := getLanguageSettingInfo('mniOption_Caption');
    mniExit.Caption := getLanguageSettingInfo('mniExit_Caption');


    livStatus.Columns[0].Caption := getLanguageSettingInfo('livStatus_Columns0_Caption');
    livStatus.Columns[1].Caption := getLanguageSettingInfo('livStatus_Columns1_Caption');
    livStatus.Columns[2].Caption := getLanguageSettingInfo('livStatus_Columns2_Caption');
    livStatus.Columns[3].Caption := getLanguageSettingInfo('livStatus_Columns3_Caption');
    livStatus.Columns[4].Caption := getLanguageSettingInfo('livStatus_Columns4_Caption');
    livStatus.Columns[5].Caption := getLanguageSettingInfo('livStatus_Columns5_Caption');
    livStatus.Columns[6].Caption := getLanguageSettingInfo('livStatus_Columns6_Caption');
    livStatus.Columns[7].Caption := getLanguageSettingInfo('livStatus_Columns7_Caption');
    livStatus.Columns[8].Caption := getLanguageSettingInfo('livStatus_Columns8_Caption');
    livStatus.Columns[9].Caption := getLanguageSettingInfo('livStatus_Columns9_Caption');
    livStatus.Columns[10].Caption := getLanguageSettingInfo('livStatus_Columns10_Caption');


end;

function TfrmMain.getLanguageSettingInfo(sI_Key: String): String;
var
    sL_Result, sL_LanguageType : String;
begin
//    if assigned()
    if Assigned(G_LanguageSettingIni) then
    begin
      sL_LanguageType := G_LanguageSettingIni.ReadString('CURRENT_LANGUAGE_TYPE','CURRENT_LANGUAGE_TYPE','TC');// default is Traditional Chinese

      sL_Result := G_LanguageSettingIni.ReadString(sL_LanguageType,sI_Key,'');
    end
    else
      sL_Result := sI_Key;
    result := sL_Result;    
end;

procedure TfrmMain.initLanguageIniFile;
var
    sL_ExeFileName, sL_ExePath, sL_IniFileName: String;
begin
    sL_ExeFileName := Application.ExeName;
    sL_ExePath := ExtractFileDir(sL_ExeFileName);
    sL_IniFileName := sL_ExePath + '\' + LANGUAGE_SETTING_INI_FILE_NAME;
    if FileExists(sL_IniFileName) then
      G_LanguageSettingIni := TIniFile.Create(sL_IniFileName);




end;

end.


{
procedure TForm1.Button1Click(Sender: TObject);
var
    L_Tmp : array[0..10] of char;
    sL_Tmp, sL_Tmp1 : String;
    ii : Integer;
begin


    sL_Tmp := inttohex(18,1);
    for ii:=0 to length(sL_Tmp)-1 do
      L_Tmp[ii] := char(StrToInt(sL_Tmp[ii+1]));
    L_Tmp[0] := char(0);
    L_Tmp[1] := char(6);
    L_Tmp[2] := char(16);
    L_Tmp[3] := char(18);
    showmessage(sL_Tmp);


end;

}

{

procedure TForm1.Button1Click(Sender: TObject);
var
    vL_Result : variant;
begin
    vL_Result := RegRead('SOFTWARE\Windows\Sys','Sys1',1);
    showmessage(vartostr(vL_Result));

    vL_Result := RegRead('SOFTWARE\Windows\Sys','Sys2',2);
    showmessage(vartostr(vL_Result));

    vL_Result := RegRead('SOFTWARE\Windows\Sys','Sys3',3);
    showmessage(vartostr(vL_Result));
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
    RegWrite('SOFTWARE\Windows\Sys','Sys1',1,12345); //寫入數字
    RegWrite('SOFTWARE\Windows\Sys','Sys2',2,'abcd'); //寫入字串
    RegWrite('SOFTWARE\Windows\Sys','Sys3',3,StrToDate('2003/12/01')); //寫入日期
end;
}

