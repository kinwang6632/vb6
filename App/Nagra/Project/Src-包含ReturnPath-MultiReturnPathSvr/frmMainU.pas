//待測指令 : 101,
//96 => Purge PPV and IPPV records
//97 => Set IPPV records as Reported

unit frmMainU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  DBTables, Db, ImgList, ExtCtrls, Menus, StdCtrls, ComCtrls, ToolWin,
  DBClient, Math, scktcomp, Buttons, ADODB, inifiles, Grids, DBGrids, winsock,
  ReceivingCommandResponseThreadU, HandleResponseU, SendCommandU,
  FetchComDataThreadU,HandleUIU, SlaveReturnPathSvrU,
  syncobjs,   UdateTimeu, Sockets, ConstObjU, Provider, Psock, NMEcho, MasterReturnPathSvrU, ProcessReturnPathDataU;

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
    Panel2: TPanel;
    Panel4: TPanel;
    btnExit: TButton;
    mmuMain: TMainMenu;
    mniOption: TMenuItem;
    mniHistory: TMenuItem;
    mniRequestLog: TMenuItem;
    mniResponseLog: TMenuItem;
    N3ErrorLog1: TMenuItem;
    mniErrorLog: TMenuItem;
    N1: TMenuItem;
    mniSendLog: TMenuItem;
    mniExit: TMenuItem;
    imlEnabled: TImageList;
    imlStatus: TImageList;
    Panel3: TPanel;
    pnlHistory: TPanel;
    Panel6: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    cdsEmm: TClientDataSet;
    cdsEmmsCommandType: TStringField;
    cdsEmmsBroadcastMode: TStringField;
    cdsEmmsAddressType: TStringField;
    cdsControl: TClientDataSet;
    cdsControlsCommandType: TStringField;
    cdsOperation: TClientDataSet;
    cdsOperationsCommandType: TStringField;
    cdsFeedback: TClientDataSet;
    cdsFeedbacksCommandType: TStringField;
    cdsProduct: TClientDataSet;
    cdsProductsCommandType: TStringField;
    Splitter1: TSplitter;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    tasTesting: TTabSheet;
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
    cdsNetworkID: TClientDataSet;
    cdsNetworkIDCompCode: TIntegerField;
    cdsNetworkIDCompName: TStringField;
    cdsNetworkIDNetworkID: TStringField;
    cdsNetworkIDOperation: TStringField;
    Label5: TLabel;
    Label8: TLabel;
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
    cdsParam: TClientDataSet;
    cdsParamsServerIP: TStringField;
    cdsParamnSPortNo: TIntegerField;
    cdsParamnRPortNo: TIntegerField;
    cdsParamsSysName: TStringField;
    cdsParamsSysVersion: TStringField;
    cdsParamsMisCommandPath: TStringField;
    cdsParamnTimeOut: TIntegerField;
    cdsParamnMaxCommandCount: TIntegerField;
    cdsParambCommandLog: TBooleanField;
    cdsParambErrorLog: TBooleanField;
    cdsParamnResponseLog: TBooleanField;
    cdsParamdUptTime: TDateTimeField;
    cdsParamsUptName: TStringField;
    cdsParambSetZipCode: TBooleanField;
    cdsParamnCreditLimit: TIntegerField;
    cdsParamnSourceID: TIntegerField;
    cdsParamnDestID: TIntegerField;
    cdsParamnMopPPID: TIntegerField;
    cdsParamsBroadcastSDate: TStringField;
    cdsParamsBroadcastEDate: TStringField;
    cdsParamnCmdRefreshRate: TIntegerField;
    cdsParamnCmdRefreshRate2: TIntegerField;
    cdsParamnCmdResentTimes: TIntegerField;
    cdsParambShowUI: TBooleanField;
    cdsParambAssignBroadcastDate: TBooleanField;
    cdsParamsSHotTime: TStringField;
    cdsParamsEHotTime: TStringField;
    ADOQuery1: TADOQuery;
    TabSheet2: TTabSheet;
    ReturnPathClntSocket: TClientSocket;
    Memo1: TMemo;
    Memo2: TMemo;
    Button1: TButton;
    memErrorLog: TMemo;
    Splitter2: TSplitter;
    livStatus: TListView;
    Button2: TButton;
    NoUseClientSocket1: TClientSocket;
    Memo3: TMemo;
    Button3: TButton;
    Memo4: TMemo;
    TcpServer1: TTcpServer;
    Button4: TButton;
    Memo5: TMemo;
    TcpClient1: TTcpClient;
    Button5: TButton;
    Button6: TButton;
    Button7: TButton;
    Button8: TButton;
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
    procedure tasTestingShow(Sender: TObject);
    procedure btnShowLogClick(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure chbStartFunClick(Sender: TObject);
    procedure SendCmdClntSocketError(Sender: TObject; Socket: TCustomWinSocket;
      ErrorEvent: TErrorEvent; var ErrorCode: Integer);
    procedure SendCmdClntSocketDisconnect(Sender: TObject;
      Socket: TCustomWinSocket);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure TcpServer1Accept(Sender: TObject;
      ClientSocket: TCustomIpClient);
    procedure TcpClient1Receive(Sender: TObject; Buf: PChar;
      var DataLen: Integer);
    procedure Button5Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure Button8Click(Sender: TObject);

  private
    { Private declarations }
    bG_LoginOk : boolean;
    G_ProcessReturnPathData_ExitCode : Cardinal;
    G_MasterReturnPathSvr_ExitCode : Cardinal;
    G_FetchComDataThread_ExitCode : Cardinal;
    G_SendCommandThread_ExitCode : Cardinal;
    G_ReceivingCommandResponseThread_ExitCode : Cardinal;
    G_THandleResponse_ExitCode, G_HandleUIThread_ExitCode : Cardinal;


    sG_MonitorIDStatus: String;
    sG_LoginID, sG_LoginName : String;
    bG_AutoInvoke : boolean;

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

  public
    { Public declarations }
//    bG_CriticalSection : boolean;
    bG_MasterReturnPathSvrIsAvailable : boolean;
    nG_ActivedSlaveReturnPathSvr : Integer;
    G_AryOfSlaveReturnPathSvr : array of SlaveReturnPathSvr;



    G_DisConnDbStrList : TStringList;
    G_SendCmdClntSocket: TClientSocket;
    nG_LivItemPos : Integer;
    G_ReturnPathDataStrList1 : TStringList;
    G_ReturnPathDataStrList : TStringList;
    G_MasterReturnPathSvr : MasterReturnPathSvr;
    G_ProcessReturnPathData : ProcessReturnPathData;
    G_FetchComDataThread : FetchComDataThread;
    G_SendCommandThread : TSendCommand;
    G_ReceivingCommandResponseThread : ReceivingCommandResponseThread;
    G_THandleResponse : THandleResponse;

    G_TimeStampStrList : TStringList;
    sG_NewIniFileName : String;
    G_HandleUIThread : THandleUI;
    nG_CurActiveConn, nG_DbCount : Integer;
    G_CompCodeStrList:TStringList;
    G_DbAliasStrList:TStringList;
    bG_StartFetching, bG_Simulator : boolean;
    G_Timer: TTimer;
    G_NotesHasValueStrList : TStringList;
    sG_Notes_Has_Value, sG_TimeDefine, sG_WriteTimeStamp : String;

    G_CommandInfoStrList : TSTringList;
    bG_HandleUIThread, bG_HandleResponseThread, bG_RunFetchCommandThread, bG_RunReceivingCommandThread,bG_RunSendingCommandThread, bG_MasterReturnPathSvrThread, bG_ProcessReturnPathDataThread : boolean;
    bG_HasBuildSendCmdConnection, bG_HasBuildReturnPathConnection : boolean;
    bG_CreateThread : boolean;
    G_SendCmdWStream, G_ReturnPathWStream: TWinSocketStream;
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
    procedure reBuildNagraConn;    
    procedure createLivItem;
    procedure showLogOnMemo(nI_ActiveConn:Integer; sI_Detail:String);
    procedure setTransCaption;    
    procedure buildSendCmdConnection;
    procedure buildReturnPathConnection;
    function reDefineFetchDataInterval(dI_Time:TDate):Integer;
    function getConnectionID(nI_CompCode:Integer):Integer;
    function getCompNetworkInfo(nI_CompCode:Integer; sI_IrmCmdID:String):String;
    procedure ReturnPathSvrThreadDone(Sender: TObject);
    procedure ProcessReturnPathDataThreadDone(Sender: TObject);

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
    procedure getErrorInfo(sI_ResErrorCode,sI_ResErrorCodeExt: String; var sI_ResErrorCodeDetail,sI_ResErrorCodeExtDetail: String);
  end;


var
  frmMain: TfrmMain;

implementation

uses frmSysParamU, frmLoginU, Ustru, dtmMainU, prjMonitor_TLB;


{$R *.DFM}

procedure TfrmMain.Panel3Click(Sender: TObject);
begin
    pnlHistory.Visible := true;
end;

procedure TfrmMain.btnExitClick(Sender: TObject);
begin
    if MessageDlg('是否確定要結束程式?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
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
    sL_BdeAlias, sL_Result : String;
    sL_DbAlias, sL_DbUserName : String;

begin
    G_DisConnDbStrList := TStringList.Create;
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
        end;
    end;
    //系統登入畫面
    
    dtmMain := TDtmMain.Create(Application);

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
      
      ActiveCDS(cdsNetworkID, NETWORK_ID_INFO_FILENAME, 'B');
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
        mniHistory.Visible := false;
        tbnHistory.Visible := false;
        MessageDlg('無法連結到高階指令資料庫(Alias:' + sL_DbAlias + ',UserID:' + sL_DbUserName + '),請修改您的系統參數,並重新啟動本程式!',mtError,[mbOK],0);
        Application.Terminate;
      end;



      G_ResponseMsgStrList := TStringList.Create;
      G_CommandInfoStrList := TSTringList.Create;
      G_ReturnPathDataStrList := TSTringList.Create;
      G_ReturnPathDataStrList1 := TSTringList.Create;

      G_TimeStampStrList := TStringList.Create;

      nG_ActivedSlaveReturnPathSvr := 0;
  //    createThread;

    end
    else
      Application.Terminate;

end;

procedure TfrmMain.GetSysParam;
var
    sL_FormatStr, sL_Path : String;
begin
    with cdsParam do
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
      bG_ResponseLog := FieldByName('nResponseLog').Asboolean;

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
        MessageDlg('系統參數尚未設定! 請先設定後,再執行程式.', mtWarning, [mbOK],0); 
      Close;
    end;
end;

procedure TfrmMain.FormShow(Sender: TObject);
var
    ii : Integer;
begin
    if bG_LoginOk then
    begin
      //建立監控視窗
      sG_MonitorIDStatus := '';
  //    Timer1.Enabled := true;

      tasTesting.TabVisible := false;
      if ParamCount > 0 then
      begin

          for ii:=1 to ParamCount do
          begin
            if UpperCase(ParamStr(ii)) = 'TESTING' then
              tasTesting.TabVisible := true;
          end;
      end;

      bG_CreateThread := true;
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
    with cdsEmm do
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
        MessageDlg('Nagra-Emm 參數尚未設定! 請先設定後,再執行程式.', mtWarning, [mbOK],0);      
      Close;
    end;
end;

procedure TfrmMain.GetControlParam;
var
    sL_Path : String;
begin
    with cdsControl do
    begin
      if (State = dsInactive) then
        CreateDataSet;
      sL_Path := ExtractFilePath(Application.ExeName);
      if (FileExists(sL_Path + CONTROL_INFO_FILENAME)) then
        LoadFromFile(sL_Path + CONTROL_INFO_FILENAME);

      sG_ControlCommandType := FieldByName('sCommandType').AsString;
      if RecordCount=0 then
        MessageDlg('Nagra-Control 參數尚未設定! 請先設定後,再執行程式.', mtWarning, [mbOK],0);      
      Close;
    end;

end;

procedure TfrmMain.GetFeedbackParam;
var
    sL_Path : String;
begin
    with cdsFeedback do
    begin
      if (State = dsInactive) then
        CreateDataSet;
      sL_Path := ExtractFilePath(Application.ExeName);
      if (FileExists(sL_Path + FEEDBACK_INFO_FILENAME)) then
        LoadFromFile(sL_Path + FEEDBACK_INFO_FILENAME);

      sG_FeedbackCommandType := FieldByName('sCommandType').AsString;
      if RecordCount=0 then
        MessageDlg('Nagra-Feedback 參數尚未設定! 請先設定後,再執行程式.', mtWarning, [mbOK],0);      
      Close;
    end;

end;

procedure TfrmMain.GetOperationParam;
var
    sL_Path : String;
begin
    with cdsOperation do
    begin
      if (State = dsInactive) then
        CreateDataSet;
      sL_Path := ExtractFilePath(Application.ExeName);
      if (FileExists(sL_Path + OPERATION_INFO_FILENAME)) then
        LoadFromFile(sL_Path + OPERATION_INFO_FILENAME);

      sG_OperationCommandType := FieldByName('sCommandType').AsString;
      if RecordCount=0 then
        MessageDlg('Nagra-Operation 參數尚未設定! 請先設定後,再執行程式.', mtWarning, [mbOK],0);      
      Close;
    end;

end;

procedure TfrmMain.GetProductParam;
var
    sL_Path : String;
begin
    with cdsProduct do
    begin
      if (State = dsInactive) then
        CreateDataSet;
      sL_Path := ExtractFilePath(Application.ExeName);
      if (FileExists(sL_Path + PRODUCT_INFO_FILENAME)) then
        LoadFromFile(sL_Path + PRODUCT_INFO_FILENAME);

      sG_ProductCommandType := FieldByName('sCommandType').AsString;
      if RecordCount=0 then
        MessageDlg('Nagra-Product 參數尚未設定! 請先設定後,再執行程式.', mtWarning, [mbOK],0);      
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
      sL_SQL := sL_SQL + ', SEQNO=SEQNO+1';
      
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

procedure TfrmMain.tasTestingShow(Sender: TObject);
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
          MessageDlg('請輸入高階指令!', mtWarning, [mbOK],0);
          edtHighCommandID.SetFocus;
          Result := false;
          exit;
         end;
       end;
      1://低階指令
       begin
         if edtLowCommandID.Text = '' then
         begin
          MessageDlg('請輸入低階指令!', mtWarning, [mbOK],0);
          edtLowCommandID.SetFocus;
          Result := false;
          exit;
         end;
       end;
    end;

    if memIccNo.Lines.Count = 0 then
    begin
      MessageDlg('請輸入Icc No!', mtWarning, [mbOK],0);
      memIccNo.SetFocus;
      Result := false;
      exit;
    end;
    if memProductID.Lines.Count = 0 then
    begin
      MessageDlg('請輸入 ProductID!', mtWarning, [mbOK],0);
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
      MessageDlg('請輸入Box No!', mtWarning, [mbOK],0);
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
      result := '-1: 讀取參數檔<'+sL_IniFileName +'>失敗';
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
    ii : Integer;
    sL_Info : String;
begin
    sL_Info := '程式即將結束,請重新執行此程式!';
    G_HandleUIThread := THandleUI.Create;

    if GetExitCodeThread(G_HandleUIThread.ThreadID,G_HandleUIThread_ExitCode) then
    begin
      if not bG_AutoInvoke then
        MessageDlg('執行緒 5 產生失敗!' + sL_Info, mtError, [mbOK],0);
      Application.Terminate;
      Exit;
    end;

    G_FetchComDataThread := FetchComDataThread.Create;
    if GetExitCodeThread(G_FetchComDataThread.ThreadID,G_FetchComDataThread_ExitCode) then
    begin
      if not bG_AutoInvoke then
        MessageDlg('執行緒 4 產生失敗!' + sL_Info, mtError, [mbOK],0);
      Application.Terminate;
      Exit;
    end;


    G_SendCommandThread := TSendCommand.Create;
    if GetExitCodeThread(G_SendCommandThread.ThreadID,G_SendCommandThread_ExitCode) then
    begin
      if not bG_AutoInvoke then
        MessageDlg('執行緒 3 產生失敗!' + sL_Info, mtError, [mbOK],0);
      Application.Terminate;
      Exit;
    end;



    G_ReceivingCommandResponseThread := ReceivingCommandResponseThread.Create(false);
    if GetExitCodeThread(G_ReceivingCommandResponseThread.ThreadID,G_ReceivingCommandResponseThread_ExitCode) then
    begin
      if not bG_AutoInvoke then
        MessageDlg('執行緒 1 產生失敗!' + sL_Info, mtError, [mbOK],0);

      Application.Terminate;
      Exit;
    end;

    G_THandleResponse := THandleResponse.Create;
    if GetExitCodeThread(G_THandleResponse.ThreadID,G_THandleResponse_ExitCode) then
    begin
      if not bG_AutoInvoke then
        MessageDlg('執行緒 2 產生失敗!' + sL_Info, mtError, [mbOK],0);
      Application.Terminate;
      Exit;
    end;




    G_MasterReturnPathSvr := MasterReturnPathSvr.Create;
    if GetExitCodeThread(G_MasterReturnPathSvr.ThreadID,G_MasterReturnPathSvr_ExitCode) then
    begin
      if not bG_AutoInvoke then
        MessageDlg('執行緒 6 產生失敗!' + sL_Info, mtError, [mbOK],0);
      Application.Terminate;
      Exit;
    end;
    

    
    G_ProcessReturnPathData := ProcessReturnPathData.Create;
    if GetExitCodeThread(G_ProcessReturnPathData.ThreadID,G_ProcessReturnPathData_ExitCode) then
    begin
      if not bG_AutoInvoke then
        MessageDlg('執行緒 7 產生失敗!' + sL_Info, mtError, [mbOK],0);
      Application.Terminate;
      Exit;
    end;

    setlength(G_AryOfSlaveReturnPathSvr, SLAVE_RETURN_PATH_SVR_COUNT);
    for ii:=0 to high(G_AryOfSlaveReturnPathSvr) do
    begin
      if ii=0 then
        G_AryOfSlaveReturnPathSvr[ii] := SlaveReturnPathSvr.Create(true, false)
      else
        G_AryOfSlaveReturnPathSvr[ii] := SlaveReturnPathSvr.Create(true, false);
//        G_AryOfSlaveReturnPathSvr[ii] := SlaveReturnPathSvr.Create(true, true);
    end;

    //傳入必要參數


    if G_SendCommandThread <> nil then
    begin
      G_SendCommandThread.sG_WriteTimeStamp := sG_WriteTimeStamp;
      G_SendCommandThread.sG_LogPath := sG_LogPath;
      G_SendCommandThread.nG_Timeout := nG_Timeout;
      G_SendCommandThread.sG_SourceID := sG_SourceID;
      G_SendCommandThread.sG_DestID := sG_DestID;
      G_SendCommandThread.sG_MopPPID := sG_MopPPID;
      G_SendCommandThread.sG_EmmCommandType := sG_EmmCommandType;
      G_SendCommandThread.sG_EmmBroadcastMode := sG_EmmBroadcastMode;
      G_SendCommandThread.sG_EmmAddressType := sG_EmmAddressType;
      G_SendCommandThread.sG_ControlCommandType := sG_ControlCommandType;
      G_SendCommandThread.sG_ProductCommandType := sG_ProductCommandType;
      G_SendCommandThread.sG_FeedbackCommandType := sG_FeedbackCommandType;
      G_SendCommandThread.sG_OperationCommandType := sG_OperationCommandType;
      G_SendCommandThread.bG_CommandLog := bG_CommandLog;
      G_SendCommandThread.bG_ErrorLog := bG_ErrorLog;
      G_SendCommandThread.bG_ResponseLog := bG_ResponseLog;
      G_SendCommandThread.sG_BroadcastStartDate := sG_BroadcastStartDate;
      G_SendCommandThread.sG_BroadcastEndDate := sG_BroadcastEndDate;
    end;


    bG_StartFetching := false;
    G_Timer:= TTimer.Create(nil);

    G_Timer.Interval := reDefineFetchDataInterval(now);
    G_Timer.OnTimer := FetchCmdDataTimer;
    G_Timer.Enabled := false;

end;

procedure TfrmMain.chbStartFunClick(Sender: TObject);
begin
    bG_RunFetchCommandThread :=  chbStartFun.Checked;
    bG_RunReceivingCommandThread := chbStartFun.Checked;

    bG_RunSendingCommandThread := chbStartFun.Checked;
    bG_HandleResponseThread := chbStartFun.Checked;
    bG_HandleUIThread := chbStartFun.Checked;

    bG_MasterReturnPathSvrThread := chbStartFun.Checked;
    bG_ProcessReturnPathDataThread := chbStartFun.Checked;

    G_Timer.Enabled :=  chbStartFun.Checked;


    if chbStartFun.Checked then
    begin
      buildSendCmdConnection;
      buildReturnPathConnection;

    end;

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
    Showmessage('網路無法使用....^_^');
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
begin
    if (I_CDS.State = dsInactive) then
      I_CDS.CreateDataSet;
    if (FileExists(sI_RefFileName)) then
      I_CDS.LoadFromFile(sI_RefFileName);

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

procedure TfrmMain.FetchCmdDataTimer(Sender: TObject);
begin
//    showmessage('timer...');
    bG_StartFetching := true;
end;

function TfrmMain.getConnectionID(nI_CompCode: Integer): Integer;
var
    nL_Ndx : Integer;
begin
    nL_Ndx := G_CompCodeStrList.IndexOf(IntToStr(nI_CompCode));
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
    L_SocketInfoStrList.Insert(0,'Socket 斷線:' + DateTimeToStr(now));
    L_SocketInfoStrList.SaveToFile(sL_ThreadInfoFileName);
    L_SocketInfoStrList.Free;
end;




function TfrmMain.getReturnPathPort: Integer;
begin
    result := nG_RPortNo;
end;

function TfrmMain.getReturnPathServerIP: String;
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
    Panel2.Caption := ' 即時監控--指令傳送中';
    if Panel2.Font.Color = clBlue then
      Panel2.Font.Color := clRed
    else
      Panel2.Font.Color := clBlue;

    self.Invalidate;
end;






procedure TfrmMain.ProcessReturnPathDataThreadDone(Sender: TObject);
begin

end;

procedure TfrmMain.showLogOnMemo(nI_ActiveConn: Integer; sI_Detail: String);
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
        sL_Msg := '['+ datetimetostr(now)+'] ' + sL_DbAlias + ' 之資料庫連線有誤.' + '  作業項目:' + sI_Detail;
        writeMsg;
      end
      else if nI_ActiveConn=SEND_CMD_ERROR_FLAG then
      begin
        bL_KillSelf := true;
        sL_DbAlias := '';
        sL_Msg := '['+ datetimetostr(now)+ '] ' +'無法傳送指令給 Nagra. ' + '  作業項目:' + sI_Detail ;

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
        sL_Msg := '['+ datetimetostr(now)+ '] ' +' 無法重新建立 Nagra Connection. ' + '  作業項目:' + sI_Detail ;

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
        sL_Msg := '['+ datetimetostr(now)+ '] ' +'Socket is not ready for reading.' + '  作業項目:' + sI_Detail ;

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
        sL_Msg := '['+ datetimetostr(now)+ '] ' + '無法執行 oracle 後端程式.'+ '  作業項目:' + sI_Detail ;

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
        sL_Msg := '['+ datetimetostr(now)+ '] ' +  ' 重新建立與 Nagra 的 connection.'+ '  作業項目:' + sI_Detail;
        writeMsg;
      end
      else if nI_ActiveConn=UPDATE_CMD_STATUS_ERROR_FLAG then
      begin
        sL_DbAlias := '';
        sL_Msg := '['+ datetimetostr(now)+ '] ' +  '無法執行 SP 來更新指令狀態.'+ '  作業項目:' + sI_Detail;
        writeMsg;
      end
      else if nI_ActiveConn=RE_CONNECT_NAGRA_ERROR_FLAG then
      begin
        sL_DbAlias := '';
        sL_Msg := '['+ datetimetostr(now)+ '] ' + '無法與 Nagra 建立連線.'+ '  作業項目:' + sI_Detail;
        writeMsg;
      end
      else if nI_ActiveConn=CLOSE_SEND_NAGRA_FLAG then
      begin
        sL_DbAlias := '';
        sL_Msg := '['+ datetimetostr(now)+']' + '  作業項目:' + sI_Detail;
        writeMsg;
      end
      else if nI_ActiveConn=HANDLE_RESPONSE_ERROR_FLAG then
      begin
        sL_DbAlias := '';
        sL_Msg := '['+ datetimetostr(now)+ '] ' + '處理 Nagra 之回應有誤.'+ '  作業項目:' + sI_Detail;
        writeMsg;
      end
      else if nI_ActiveConn=WRITE_TO_MIS_ERROR_FLAG then
      begin
        sL_DbAlias := '';
        sL_Msg := '['+ datetimetostr(now)+ '] ' + '回寫指令狀態有誤.'+ '  作業項目:' + sI_Detail;
        writeMsg;
      end
      else if nI_ActiveConn=RECEIVE_CMD_RESPONSE_ERROR_FLAG then
      begin
        sL_DbAlias := '';
        sL_Msg := '['+ datetimetostr(now)+ '] ' + '接收 Nagra 之回應有誤.'+ '  作業項目:' + sI_Detail;
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
        sL_Msg := '['+ datetimetostr(now)+']' + '  作業項目:' + sI_Detail + '. 啟動者:' + sG_LoginName;
        writeMsg;
      end;

      if bL_KillSelf then //未必會被執行到, 因為呼叫  L_Intf.RunCa 時, 就會把 ca command gateway kill 掉
        closeProg;
  //      Application.Terminate;
    except

    end;
end;

procedure TfrmMain.Button1Click(Sender: TObject);
var
    ii : Integer;
    L_ReturnPathDataObj : TReturnPathDataObj;
begin
    Memo2.Lines.Clear;
    Memo2.Lines.Add('count==' + IntToStr(frmMain.G_ReturnPathDataStrList1.Count));

    for ii:=0 to frmMain.G_ReturnPathDataStrList1.Count-1 do
    begin
      Memo2.Lines.Add(G_ReturnPathDataStrList1.Strings[ii]);
    end;
end;

procedure TfrmMain.createLivItem;
var
    ii, jj : Integer;
    L_ListItem : TListItem;
    L_Splitter: TSplitter;
begin
    frmMain.livStatus.Items.Clear;
    if bG_ShowUI then
    begin
      nG_LivItemPos := 0;
      memErrorLog.Align := alBottom;
      memErrorLog.Height := 122;

      L_Splitter:= TSplitter.Create(Application);
      L_Splitter.Parent := TabSheet1;
      L_Splitter.Height := 4;
      L_Splitter.Align := alBottom;


      livStatus.Visible := true;
      livStatus.Align := alClient;


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

procedure TfrmMain.Button2Click(Sender: TObject);
begin
frmMain.Memo3.Lines.SaveToFile('c:\aacc.txt');
end;

procedure TfrmMain.closeProg;
begin
    showLogOnMemo(CLOSE_SEND_NAGRA_FLAG,'正常結束ca command gateway');
    releaseResource;

//      Application.Terminate;
    Close;
end;

procedure TfrmMain.releaseResource;
var
    ii : Integer;
begin
    chbStartFun.Checked := false;

    bG_RunFetchCommandThread := false;
    bG_RunReceivingCommandThread := false;
    bG_RunSendingCommandThread := false;
    bG_HandleResponseThread := false;
    bG_HandleUIThread := false;


    G_FetchComDataThread.Terminate;
    G_SendCommandThread.Terminate;
    G_ReceivingCommandResponseThread.Terminate;
    G_THandleResponse.Terminate;
    G_HandleUIThread.Terminate;


    for ii:=0 to high(G_AryOfSlaveReturnPathSvr) do
    begin
      G_AryOfSlaveReturnPathSvr[ii].Terminate;
    end;
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



end;

procedure TfrmMain.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
    if (ssCtrl in Shift) and (Key=83) then //Ctrl + s  => 切換 UI 的顯示
    begin
      bG_ShowUI := not bG_ShowUI;
      createLivItem;
    end;

end;

procedure TfrmMain.buildSendCmdConnection;
begin
//down, 建立用來 send command 的 connection..
    if not bG_HasBuildSendCmdConnection then
    begin
      if G_SendCommandThread<> nil then
       (G_SendCommandThread as TSendCommand).buildSendCmdConnection(G_SendCmdClntSocket, G_SendCmdWStream);
//       (G_SendCommandThread as TSendCommand).buildSendCmdConnection(G_SendCmdClntSocket);
    end;
    //up, 建立用來 send command 的 connection..

end;

procedure TfrmMain.buildReturnPathConnection;
begin
    //down, 建立用來處理 return path 的 connection..

    if not bG_HasBuildReturnPathConnection then
    begin
      if G_MasterReturnPathSvr<> nil then
       (G_MasterReturnPathSvr as MasterReturnPathSvr).buildConnection(ReturnPathClntSocket, G_ReturnPathWStream);
//         (G_MasterReturnPathSvr as ReturnPathSvr).buildConnection(ReturnPathClntSocket);
    end;

    //up, 建立用來處理 return path 的 connection..

end;

procedure TfrmMain.reBuildNagraConn;
begin
    try
      bG_HasBuildSendCmdConnection := false;

      G_SendCmdWStream.Free;
      G_SendCmdWStream := nil;

      G_SendCmdClntSocket.Close;
//        sleep(3000);
      buildSendCmdConnection;
    except
      frmMain.showLogOnMemo(REBUILD_NAGRA_CONNECTION_FAILED,' frmMain.reBuildNagraConn .程式即將結束,然後重新啟動. '); //執行完此行程式後, 程式會結束掉
    end;
end;
procedure TfrmMain.Button3Click(Sender: TObject);
begin
    G_MasterReturnPathSvr.sendCmd('1000', '123');
end;

procedure TfrmMain.Button4Click(Sender: TObject);
begin
    TcpServer1.LocalHost := '10.40.16.100';
    TcpServer1.LocalPort := '60003';
//    TcpServer1.OnAccept := TcpServerAccept;
    TcpServer1.Open;

end;

procedure TfrmMain.TcpServer1Accept(Sender: TObject;
  ClientSocket: TCustomIpClient);
var
  sL_ClientNdsData: string;
  ii, nL_CmdType : Integer;
begin
  sL_ClientNdsData := ClientSocket.Receiveln;
  while sL_ClientNdsData <> '' do
  begin
//    Application.ProcessMessages;
    Memo5.Lines.Add(sL_ClientNdsData);

    sL_ClientNdsData := ClientSocket.Receiveln;
  end;

end;

procedure TfrmMain.TcpClient1Receive(Sender: TObject; Buf: PChar;
  var DataLen: Integer);
var
    sL_Data : String;  
begin
    sL_Data := TcpClient1.Receiveln;
    Memo5.Lines.Add(sL_Data);
end;

procedure TfrmMain.Button5Click(Sender: TObject);
var
    ii : Integer;
    sL_Cmd : String;
    L_TmpChar : array[0..1024] of char;
begin
    sL_Cmd := '098000001050003025744289200304301002';

    L_TmpChar[0] := char(0);
    L_TmpChar[1] := '$';

    for ii:=1 to length(sL_Cmd) do
      L_TmpChar[ii+1] := sL_Cmd[ii];
//    StrPCopy(L_TmpChar, sL_Cmd);

    TcpClient1.RemoteHost := '10.40.16.1';
    TcpClient1.RemotePort := '60003';
    TcpClient1.Open;
    TcpClient1.SendBuf(L_TmpChar, length(sL_Cmd));
end;

procedure TfrmMain.Button6Click(Sender: TObject);
var
    ii : Integer;
begin
    Memo3.Clear;
    for ii:=0 to high(G_AryOfSlaveReturnPathSvr) do
    begin
      if frmMain.G_AryOfSlaveReturnPathSvr[ii].Suspended then
        Memo3.Lines.Add(inttostr(ii) + '=>Suspended')
      else
        Memo3.Lines.Add(inttostr(ii) + '=>not Suspended');        
    end;
end;

procedure TfrmMain.Button7Click(Sender: TObject);
begin
    frmMain.G_AryOfSlaveReturnPathSvr[8].Execute;
end;

procedure TfrmMain.Button8Click(Sender: TObject);
begin
    G_ReturnPathDataStrList1.Clear;
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

procedure TfrmMain.ReturnPathClntSocketRead(Sender: TObject;
  Socket: TCustomWinSocket);
var
  sL_ClientNdsData: string;
  ii, nL_CmdType : Integer;
begin
  sL_ClientNdsData := Socket.ReceiveText;
  while sL_ClientNdsData <> '' do
  begin
//    Application.ProcessMessages;
    Memo5.Lines.Add(sL_ClientNdsData);

    sL_ClientNdsData := Socket.ReceiveText;
  end;

end;
}
