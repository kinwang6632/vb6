unit frmMainU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  DBTables, Db, ImgList, ExtCtrls, Menus, StdCtrls, ComCtrls, ToolWin,
  DBClient, Buttons, IniFiles,
  ReceivingCommandResponseThreadU, HandleResponseU, SendCommandU,
  FetchComDataThreadU, HandleUIU, ProcessReturnPathDataU, ReturnPathSvrU,
  UdateTimeu, Sockets, ConstObjU, Variants, CsIntf, Grids, DBGrids, ADODB, CbQueue;


var
  _NAGRA_REJECT: Boolean = True;
  _FIRST_START: Boolean = True;

  _AUTO_SCROLL: Boolean = True;

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
    ToolButton2: TToolButton;
    TabSheet1: TTabSheet;
    ListView1: TListView;
    ControlSocket: TTcpClient;
    FeedbackSocket: TTcpClient;
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
    procedure chbStartFunClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Button1Click(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure Timer2Timer(Sender: TObject);
    procedure Timer3Timer(Sender: TObject);
    procedure ToolButton2Click(Sender: TObject);
    procedure tbnHistoryClick(Sender: TObject);
  private
    { Private declarations }
    dG_Sys1, dG_Sys2, dG_Sys3 : TDate;
    sG_Sys4, sG_Sys5 : String;
    bG_LoginOk : boolean;


    sG_MonitorIDStatus: String;
    sG_LoginID, sG_LoginName : String;
    bG_AutoInvoke : boolean;

    G_LanguageSettingIni : TInifile;
    
    procedure SetLanguageString;
    procedure InitLanguageIniFile;
    procedure CreateLivItem;
    procedure SaveThreadInfo(const AInfo:String);
    procedure CreateThread;
    procedure GetSysParam;
    procedure GetEmmParam;
    procedure GetControlParam;
    procedure GetProductParam;
    procedure GetFeedbackParam;
    procedure GetOperationParam;

    procedure CloseProg;
    procedure ReleaseResource;
    procedure TransNetworkIdInfo;  //將 networkid 之資料由 CDS 轉到 string list
    function GetRegistryInfo: Boolean;
    function WriteRegData(const sI_EncryptedLicenseKey: string): Boolean;
    function GetVirtualNotes(const sI_HighLevelCmdID: string): String;
    function IsDataOK: Boolean;
    function GetSysIniSettiing: WideString;
    procedure ReadTransNumber;
    procedure SaveTransNumber;
  public
    { Public declarations }


//down, 以下是 ReceiveResponse thread 之變數..
    sG_RR_TestInfo : String;
    G_RR_DeviceData : TDeviceData;
    



//down, 以下是 FetchComData thread 之變數..
//    CRS: TCriticalSection;
    nG_FD_CurActiveConn : Integer;


    sG_FD_SeqNos : STring;

    { Send Command Transcation Number }
    nG_FD_TestingCmdSeqNo,
    nG_FD_MaxTestingCmdSeqNo,
    nG_SC_RealCmdSeqNo,
    nG_SC_MaxRealCmdSeqNo: Integer;

    { Return Path Command Transcation Number }
    nG_FD_TestingCmdSeqNo2,
    nG_FD_MaxTestingCmdSeqNo2,
    nG_SC_RealCmdSeqNo2,
    nG_SC_MaxRealCmdSeqNo2: Integer;



//down, 以下是 HandleResponse thread 之變數..
//    CRS: TCriticalSection;
//    sG_SeqNoForUpdateCmdStatus, sG_TransNumForUpdateCmdStatus : String;
    sG_HR_ResponseStr : String;
    nG_HR_MsgStrListPos : Integer;
    nG_HR_CompCode, nG_HR_ConnectionID : Integer;



//down, 以下是 SendCommand thread 之變數..

    sG_SC_ActiveMessage, sG_SC_MasterIccNo : String;
    nG_SC_SegmentNumber, nG_SC_TotalSeqment, nG_SC_MailId: Integer;
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

    bG_FirstRun: Boolean;


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

    // 會被填入 Nagra  傳來的 transaction num => return path 時會用到
    sG_SrcTransNum : String;

//Up, 以下是 SendCommand thread 之變數..

    sG_CurTimeStamp : String;
    sG_CurActiveLowLevelCmd : String;
    sG_TransactionNum : String;
    G_NetworkIdInfoStrList : TStringList;
    G_DisConnDbStrList : TStringList;

    nG_LivItemPos : Integer;
    G_TimeStampStrList : TStringList;
    sG_NewIniFileName : String;


    G_ReceivingCommandResponseThread : ReceivingCommandResponseThread;
    G_ProcessReturnPathData: ProcessReturnPathData;
    G_ReturnPathSvr: ReturnPathSvr;

    nG_CurActiveConn, nG_DbCount : Integer;
    G_CompCodeStrList:TStringList;
    G_DbAliasStrList:TStringList;
    bG_StartFetching, bG_Simulator, bG_InputLicense : boolean;
    G_NotesHasValueStrList : TStringList;
    sG_Notes_Has_Value, sG_TimeDefine, sG_WriteTimeStamp : String;

    //G_CommandInfoStrList : TSTringList;
    G_CommandInfoStrList: TCommandStringList;
    //G_ResponseMsgStrList : TStringList;
    G_ResponseMsgStrList : TCommandStringList;
    //G_ReturnPathDataStrList : TStringList;
    G_ReturnPathDataStrList : TCommandStringList;

    bG_HandleResponseThread, bG_RunFetchCommandThread,
    bG_RunReceivingCommandThread, bG_RunSendingCommandThread,
    bG_ProcessReturnPathDataThread, bG_ReturnPathSvrThread: Boolean;

    bG_HasBuildSendCmdConnection, bG_HasBuildReturnPathConnection : boolean;

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



    function GetLanguageSettingInfo(const AKey:String):String;
    procedure showLogOnMemo(nI_ActiveConn:Integer; sI_Detail:String);
    procedure setTransCaption;
    procedure BuildSendCmdConnection;
    procedure BuildReturnPathConnection;
    function reDefineFetchDataInterval(dI_Time:TDate):Integer;
    function getConnectionID(nI_CompCode:Integer):Integer;
    function getCompNetworkInfo(aCompCode: Integer; aIrmCmdID:String):String;
    procedure ReturnPathSvrThreadDone(Sender: TObject);
    procedure ReceiveResponseThreadDone(Sender: TObject);
    procedure SendCommandThreadDone(Sender: TObject);
    procedure ProcessReturnPathDataThreadDone(Sender: TObject);
    procedure ActiveCDS(I_CDS:TClientDataSet; sI_RefFileName:String; sI_DefaultStatus:String);
    function getSendCommandServerIP:STring;
    function getSendCommandPort:Integer;
    function getReturnPathServerIP:STring;
    function getReturnPathPort:Integer;
    function getListenCommandPort:Integer;
    function getBinaryVal(nI_Value: Integer; nI_Digits: Integer; bI_Reverse: boolean):TTempRecord;
    function ReverseBinaryVal(ArrayChars: array of Char): Integer;
    function TransReceivedDataCharAryToStr(ATmpAry:array of Char;
      ALength:Integer ):String;
    function TransReceivedDataCharAryToStr2(ATmpAry:array of Char;
      ALength:Integer ):String;
    procedure getErrorInfo(sI_ResErrorCode,sI_ResErrorCodeExt: String; var sI_ResErrorCodeDetail,sI_ResErrorCodeExtDetail: String);
  end;


var
  frmMain: TfrmMain;

implementation

uses frmSysParamU, frmLoginU, Ustru, dtmMainU, prjMonitor_TLB, Uotheru;

{$R *.DFM}

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.Panel3Click(Sender: TObject);
begin
  pnlHistory.Visible := true;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.btnExitClick(Sender: TObject);
var
  AMsg: String;
begin
  AMsg := getLanguageSettingInfo('frmMain_Msg_1_Content');
  if Application.MessageBox( PChar( AMsg ), 'Confirmation',
    MB_OKCANCEL + MB_ICONQUESTION ) = ID_OK then
  begin
    CloseProg;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.tbnExitClick(Sender: TObject);
begin
  btnExitClick( Sender );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caFree;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.mniOptionClick(Sender: TObject);
begin
  tbnOptionClick( Sender );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.mniExitClick(Sender: TObject);
begin
  btnExitClick( Sender );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.tbnOptionClick(Sender: TObject);
begin
  frmSysParam := TfrmSysParam.Create( Application );
  try
    frmSysParam.sG_LoginName := sG_LoginName;
    frmSysParam.ShowModal;
  finally
    frmSysParam.Free;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.FormCreate(Sender: TObject);
var
    AIndex: Integer;
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

    bG_Simulator := False;
    bG_AutoInvoke := False;
    sG_NewIniFileName := '';

    if ParamCount > 0 then
    begin
      for AIndex := 1 to ParamCount do
      begin
        //起動時,自動開始傳送及接收 
        if ( UpperCase( ParamStr(AIndex ) ) = 'AUTO_INVOKE' ) then
          bG_AutoInvoke := True;
        //表示指明要用哪一個 ini file...
        if AnsiPos('.INI', UpperCase(ParamStr(AIndex ) ) ) > 0 then
          sG_NewIniFileName := ParamStr( AIndex );
        // Nagar給的模擬器
        if UpperCase( ParamStr( AIndex ) ) = 'SIM' then
          bG_Simulator := True;
        if UpperCase( ParamStr(AIndex ) ) = 'IL' then
          bG_InputLicense := True;
      end;
    end;

    if bG_InputLicense then
    begin
      try
        sL_EncryptedLicenseKey := InputBox(
          getLanguageSettingInfo( 'frmMain_Msg_2_Content'), 'License Key:', '' );
        if sL_EncryptedLicenseKey = '' then
          Application.Terminate
        else if writeRegData( sL_EncryptedLicenseKey ) then
        begin
          MessageDlg(
            getLanguageSettingInfo('frmMain_Msg_3_Content'), mtInformation,
            [mbOK],0);
          Application.Terminate;
        end
        else
        begin
          MessageDlg(
            getLanguageSettingInfo('frmMain_Msg_4_Content'), mtError,
            [mbOK],0);
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
        //dtmMain := TDtmMain.Create(Application);
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

            tbnHistory.Visible := false;
            MessageDlg(getLanguageSettingInfo('frmMain_Msg_10_Content') + sL_DbAlias + ',UserID:' + sL_DbUserName + '),請修改您的系統參數,並重新啟動本程式!',mtError,[mbOK],0);
            Application.Terminate;
          end;


          //G_ResponseMsgStrList := TStringList.Create;
          //G_CommandInfoStrList := TSTringList.Create;
          G_ResponseMsgStrList := TCommandStringList.Create;
          G_CommandInfoStrList := TCommandStringList.Create;
          G_TimeStampStrList := TStringList.Create;
          //G_ReturnPathDataStrList := TSTringList.Create;
          G_ReturnPathDataStrList := TCommandStringList.Create;
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

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.FormShow(Sender: TObject);
var
  AIndex : Integer;
begin
  if bG_LoginOk then
  begin
    ShowLogOnMemo( EXECUTE_SEND_NAGRA_FLAG,
      getLanguageSettingInfo( 'frmMain_Msg_12_Content' ) );
    //建立監控視窗
    sG_MonitorIDStatus := '';
    frmMain_Tab_2.TabVisible := False;
    TabSheet1.TabVisible := False;
    if ParamCount > 0 then
    begin
      for AIndex := 1 to ParamCount do
      begin
        frmMain_Tab_2.TabVisible :=
          ( UpperCase( ParamStr( AIndex ) ) = 'TESTING' );
      end;
    end;
    nG_CurActiveConn := nG_DbCount - 1;
    frmMain.bG_HasBuildSendCmdConnection := False;
    frmMain.bG_HasBuildReturnPathConnection := False;

    nG_FD_MaxTestingCmdSeqNo :=
      StrToIntDef(TUstr.AddString( '', '9', True,SEQ_NUM_LENGTH ),999999 );
    nG_SC_MaxRealCmdSeqNo :=
      StrToIntDef( TUstr.AddString( '', '9' , True, SEQ_NUM_LENGTH ), 999999 );

    nG_FD_MaxTestingCmdSeqNo2 :=
      StrToIntDef(TUstr.AddString( '', '9', True,SEQ_NUM_LENGTH ),999999 );
    nG_SC_MaxRealCmdSeqNo2 :=
      StrToIntDef( TUstr.AddString( '', '9' , True, SEQ_NUM_LENGTH ), 999999 );

    ReadTransNumber;  
    CreateThread;
    CreateLivItem;
    chbStartFun.Checked := bG_AutoInvoke;
    { down, 讓程式畫面產生 refresh 的效果, 若不做,CA有時會無法開始傳送指令 }
    Self.WindowState := wsMinimized;
    Self.WindowState := wsMaximized;
    Self.WindowState := wsNormal;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.GetEmmParam;
var
  AFileName, AMsg : String;
begin
  with dtmMain.cdsEmm do
  begin
    if ( State = dsInactive ) then CreateDataSet;
    AFileName := IncludeTrailingPathDelimiter( ExtractFilePath(
      Application.ExeName ) ) + EMM_INFO_FILENAME;
    if FileExists( AFileName ) then LoadFromFile( AFileName );
    sG_EmmCommandType := FieldByName('sCommandType').AsString;
    sG_EmmBroadcastMode := FieldByName('sBroadcastMode').AsString;
    sG_EmmAddressType := FieldByName('sAddressType').AsString;
    if RecordCount = 0 then
    begin
      AMsg := getLanguageSettingInfo( 'frmMain_Msg_13_Content' );
      Application.MessageBox( PChar( AMsg ), 'warning', MB_OK + MB_ICONWARNING );
    end;
    Close;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.GetControlParam;
var
  AFileName, AMsg : String;
begin
  with dtmMain.cdsControl do
  begin
    if ( State = dsInactive ) then CreateDataSet;
    AFileName := IncludeTrailingPathDelimiter( ExtractFilePath(
      Application.ExeName ) ) + CONTROL_INFO_FILENAME;
    if FileExists( AFileName ) then LoadFromFile( AFileName );
    sG_ControlCommandType := FieldByName('sCommandType').AsString;
    if RecordCount = 0 then
    begin
      AMsg := getLanguageSettingInfo( 'frmMain_Msg_14_Content' );
      Application.MessageBox( PChar( AMsg ), 'warning', MB_OK + MB_ICONWARNING );
    end;
    Close;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.GetFeedbackParam;
var
  AFileName, AMsg : String;
begin
  with dtmMain.cdsFeedback do
  begin
    if ( State = dsInactive ) then CreateDataSet;
    AFileName := IncludeTrailingPathDelimiter( ExtractFilePath(
      Application.ExeName ) ) + FEEDBACK_INFO_FILENAME;
    if FileExists( AFileName ) then LoadFromFile( AFileName );
    sG_FeedbackCommandType := FieldByName('sCommandType').AsString;
    if RecordCount = 0 then
    begin
      AMsg := getLanguageSettingInfo( 'frmMain_Msg_15_Content' );
      Application.MessageBox( PChar( AMsg ), 'warning', MB_OK + MB_ICONWARNING );
    end;
    Close;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.GetOperationParam;
var
  AFileName, AMsg : String;
begin
  with dtmMain.cdsOperation do
  begin
    if ( State = dsInactive ) then CreateDataSet;
    AFileName := IncludeTrailingPathDelimiter( ExtractFilePath(
      Application.ExeName ) ) + OPERATION_INFO_FILENAME;
    if FileExists( AFileName ) then LoadFromFile( AFileName );
    sG_OperationCommandType := FieldByName('sCommandType').AsString;
    if RecordCount = 0 then
    begin
      AMsg := getLanguageSettingInfo( 'frmMain_Msg_16_Content' );
      Application.MessageBox( PChar( AMsg ), 'warning', MB_OK + MB_ICONWARNING );
    end;
    Close;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.GetProductParam;
var
  AFileName, AMsg : String;
begin
  with dtmMain.cdsProduct do
  begin
    if ( State = dsInactive ) then CreateDataSet;
    AFileName := IncludeTrailingPathDelimiter( ExtractFilePath(
      Application.ExeName ) ) + PRODUCT_INFO_FILENAME;
    if FileExists( AFileName ) then LoadFromFile( AFileName );
    sG_ProductCommandType := FieldByName('sCommandType').AsString;
    if RecordCount = 0 then
    begin
      AMsg := getLanguageSettingInfo( 'frmMain_Msg_17_Content' );
      Application.MessageBox( PChar( AMsg ), 'warning', MB_OK + MB_ICONWARNING );
    end;
    Close;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmMain.getSendCommandPort: Integer;
begin
  Result := nG_SPortNo;
end;

{ ---------------------------------------------------------------------------- }

function TfrmMain.getSendCommandServerIP: STring;
begin
  Result := sG_ServerIP;
end;

{ ---------------------------------------------------------------------------- }

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
      sL_Notes := GetVirtualNotes( Uppercase( edtHighCommandID.Text ) );
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
  sL_WhereClause : String;
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
end;

{ ---------------------------------------------------------------------------- }

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
    saveThreadInfo('Receive Response Thread Done.');
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.CreateThread;
begin
  //G_ReturnPathSvr := ReturnPathSvr.Create;
  //G_ProcessReturnPathData := ProcessReturnPathData.Create;
  G_ReceivingCommandResponseThread :=
    ReceivingCommandResponseThread.Create;
  bG_StartFetching := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.chbStartFunClick(Sender: TObject);
begin
  bG_RunFetchCommandThread :=  chbStartFun.Checked;
  bG_RunReceivingCommandThread := chbStartFun.Checked;
  bG_RunSendingCommandThread := chbStartFun.Checked;
  bG_HandleResponseThread := chbStartFun.Checked;
  bG_ReturnPathSvrThread := chbStartFun.Checked;
  bG_ProcessReturnPathDataThread := chbStartFun.Checked;
  bG_FirstRun := chbStartFun.Checked;
  if chbStartFun.Checked then
  begin
    BuildSendCmdConnection;
    BuildReturnPathConnection;
  end;
  Timer1.Enabled := chbStartFun.Checked;
  Timer2.Enabled := chbStartFun.Checked;
  Timer3.Enabled := chbStartFun.Checked;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.ProcessReturnPathDataThreadDone(Sender: TObject);
begin
  if (Sender as TThread).FatalException <> nil then
  begin
    SaveThreadInfo( 'Receive Response Thread Done.' );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.SendCommandThreadDone(Sender: TObject);
begin
  if (Sender as TThread).FatalException <> nil then
  begin
    saveThreadInfo('Send Command Thread Done.');
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.ActiveCDS(I_CDS: TClientDataSet; sI_RefFileName,
  sI_DefaultStatus: String);
var
    sL_Path : String;
begin


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

{ ---------------------------------------------------------------------------- }

function TfrmMain.getCompNetworkInfo(aCompCode: Integer; aIrmCmdID:String): String;
var
  aIndex, aIndex2: Integer;
  aNetworkId, aNetworkHex, aOperation : String;
begin
  aNetworkId := EmptyStr;
  aOperation := EmptyStr;
  aIndex := G_NetworkIdInfoStrList.IndexOf( IntToStr( aCompCode ) );
  if aIndex <> -1 then
  begin
    aNetworkId :=
      TNetworkIDInfoObj( G_NetworkIdInfoStrList.Objects[aIndex] ).NetworkID;
    if ( aIrmCmdID = NETWORK_OPERATION_194 ) then
    begin
      aNetworkId := IntToHex( StrToInt( aNetworkId ), 0 );
      aNetworkId := TUstr.AddString( aNetworkId,'0', False, 4 );
    end else
    if ( aIrmCmdID = NETWORK_OPERATION_198 ) then
    begin
      { 10 進制數字 --> 轉 16 進制 --> 每一個 Char 取 ASCII,
        Ex: NetworkId = 10
            轉 16 進制 = 000A
            取 ASCII = 48 48 48 65 }
      aNetworkId := IntToHex( StrToInt( aNetworkId ), 4 );
      for aIndex := 1 to Length( aNetworkId ) do
        aNetworkHex := ( aNetworkHex + IntToStr( Ord( aNetworkId[aIndex] ) ) );
      aNetworkId := aNetworkHex;
    end else
    if ( aIrmCmdID = NETWORK_OPERATION_195 ) then
    begin
      aOperation := TNetworkIDInfoObj( G_NetworkIdInfoStrList.Objects[aIndex] ).Operation;
      aOperation := TUstr.AddString( aOperation, '0', False, 4 );
    end;
  end;
  Result := aNetworkId + aOperation;
end;

{ ---------------------------------------------------------------------------- }

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

{ ---------------------------------------------------------------------------- }

function TfrmMain.GetVirtualNotes(const sI_HighLevelCmdID: String): String;
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

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.SaveThreadInfo(const AInfo: String);
var
  AContent: TStringList;
  AFileName: String;
begin
  AFileName := IncludeTrailingPathDelimiter( ExtractFileDir(
    Application.ExeName ) ) +  THREAD_INFO_FILE;
  AContent := TStringList.Create;
  try
    if FileExists( AFileName ) then AContent.LoadFromFile( AFileName );
    if AContent.Count > 500 then AContent.Clear;
    AContent.Insert( 0, AInfo + ':' + DateTimeToStr( Now ) );
    AContent.SaveToFile( AFileName );
  finally
    AContent.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmMain.getReturnPathPort: Integer;
begin
  Result := nG_RPortNo;
end;

{ ---------------------------------------------------------------------------- }

function TfrmMain.getReturnPathServerIP: STring;
begin
  Result := sG_ServerIP;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.ReturnPathSvrThreadDone(Sender: TObject);
begin
  if (Sender as TThread).FatalException <> nil then
    SaveThreadInfo( 'Return Path Server Thread Done.' );
end;

{ ---------------------------------------------------------------------------- }

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
        SaveTransNumber;
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
        SaveTransNumber;
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
        SaveTransNumber;
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

{ ---------------------------------------------------------------------------- }


procedure TfrmMain.CloseProg;
begin
  ShowLogOnMemo( CLOSE_SEND_NAGRA_FLAG,
    getLanguageSettingInfo('frmMain_Msg_43_Content' ) );
  ReleaseResource;
  SaveTransNumber;
  Self.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.ReleaseResource;
var
  aIndex: Integer;
begin
    //CloseSocket;
    chbStartFun.Checked := False;
    bG_RunFetchCommandThread := False;
    bG_RunReceivingCommandThread := False;
    bG_RunSendingCommandThread := False;
    bG_HandleResponseThread := False;
    bG_ReturnPathSvrThread := False;
    bG_ProcessReturnPathDataThread := False;

    ControlSocket.Active := False;

    if Assigned( G_ReceivingCommandResponseThread ) then
    begin
      G_ReceivingCommandResponseThread.Terminate;
      G_ReceivingCommandResponseThread.WaitFor;
      G_ReceivingCommandResponseThread.Free;
    end;

    if Assigned( G_ProcessReturnPathData ) then
    begin
      G_ProcessReturnPathData.Terminate;
      G_ProcessReturnPathData.WaitFor;
      G_ProcessReturnPathData.Free;
    end;

    FeedBackSocket.Active := False;

    if Assigned( G_ReturnPathSvr ) then
    begin
      G_ReturnPathSvr.Terminate;
      G_ReturnPathSvr.WaitFor;
      G_ReturnPathSvr.Free;
    end;


    if Assigned( G_DisConnDbStrList ) then
    begin
      G_DisConnDbStrList.Free;
      G_DisConnDbStrList := nil;
    end;

    if Assigned( G_ResponseMsgStrList ) then
    begin
      G_ResponseMsgStrList.Free;
      G_ResponseMsgStrList := nil;
    end;

    if Assigned( G_CompCodeStrList ) then
    begin
      G_CompCodeStrList.Free;
      G_CompCodeStrList := nil;
    end;

    if Assigned( G_ReturnPathDataStrList ) then
    begin
      G_ReturnPathDataStrList.Free;
      G_ReturnPathDataStrList := nil;
    end;

    if Assigned( G_DbAliasStrList ) then
    begin
      G_DbAliasStrList.Free;
      G_DbAliasStrList := nil;
    end;

    if Assigned( G_NetworkIdInfoStrList ) then
    begin
      for aIndex := 0 to G_NetworkIdInfoStrList.Count - 1 do
      begin
        if Assigned( G_NetworkIdInfoStrList.Objects[aIndex] ) then
          TNetworkIDInfoObj( G_NetworkIdInfoStrList.Objects[aIndex] ).Free;
      end;  
      G_NetworkIdInfoStrList.Free;
      G_NetworkIdInfoStrList := nil;
    end;

    if Assigned(G_CommandInfoStrList) then
    begin
      for aIndex := 0 to G_CommandInfoStrList.Count - 1 do
      begin
        if Assigned( G_CommandInfoStrList.Objects[aIndex] ) then
          TCmdInfoObject( G_CommandInfoStrList.Objects[aIndex] ).Free;
      end; 
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

{ ---------------------------------------------------------------------------- }

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


{ ---------------------------------------------------------------------------- }

procedure TfrmMain.BuildSendCmdConnection;
begin
  if not bG_HasBuildSendCmdConnection then
  begin
    SendCommandU.BuildSendCmdConnection( ControlSocket );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmMain.ReverseBinaryVal(ArrayChars: array of Char): Integer;
var
  ATemp : String;
  AChar : Char;
  AIndex : Integer;
begin
  ATemp := '';
  for AIndex := 0 to High( ArrayChars ) do
  begin
    AChar := ArrayChars[AIndex];
    ATemp := ATemp + IntToStr( Ord( AChar ) );
  end;
  Result := StrToInt( ATemp );
end;

{ ---------------------------------------------------------------------------- }

function TfrmMain.TransReceivedDataCharAryToStr(ATmpAry: array of Char;
  ALength: Integer): String;
var
  ASeed, ATmpLength, ALoopCount : Integer;
  L_TmpAry : array[0..1] of Char;

  { ------------------------------------------------------- }

  function CopyAryToStr(ASrc: array of Char; const APos,
    ACount: Integer): String;
  var
    AIndex : Integer;
  begin
    Result := '';
    for AIndex := 0 to ACount - 1 do
      Result := Result + ASrc[APos + AIndex];
  end;

  { ------------------------------------------------------- }

begin
  ASeed := 0;
  ALoopCount := 0;
  Result := '';
  while True do
  begin
    Inc( ALoopCount );
    L_TmpAry[0] := ATmpAry[ASeed];
    L_TmpAry[1] := ATmpAry[ASeed + 1];
    ATmpLength := ReverseBinaryVal( L_TmpAry );
    if Result = '' then
      Result := CopyAryToStr( ATmpAry, ASeed + 2, ATmpLength )
    else
      Result := Result + RECEIVED_DATA_SEP_FLAG + CopyAryToStr(
        ATmpAry, ASeed + 2, ATmpLength );
    ASeed := Length( Result ) + 2 * ALoopCount - ( ALoopCount - 1 );
    if ASeed >= ALength then Break;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmMain.TransReceivedDataCharAryToStr2(ATmpAry: array of Char;
  ALength: Integer): String;
var
  ASeed, ATmpLength, ALoopCount : Integer;
  L_TmpAry : array[0..1] of Char;

  { ------------------------------------------------------- }

  function CopyAryToStr(ASrc: array of Char; const APos,
    ACount: Integer): String;
  var
    AIndex : Integer;
  begin
    Result := '';
    for AIndex := 0 to ACount - 1 do
      Result := Result + ASrc[APos + AIndex];
  end;

  { ------------------------------------------------------- }

begin
  ASeed := 0;
  ALoopCount := 0;
  Result := '';
  while True do
  begin
    Inc( ALoopCount );
    L_TmpAry[0] := ATmpAry[ASeed];
    L_TmpAry[1] := ATmpAry[ASeed + 1];
    ATmpLength := ReverseBinaryVal( L_TmpAry );
    if Result = '' then
      Result := CopyAryToStr( ATmpAry, ASeed + 2, ATmpLength )
    else
      Result := Result + RECEIVED_DATA_SEP_FLAG + CopyAryToStr(
        ATmpAry, ASeed + 2, ATmpLength );
    ASeed := Length( Result ) + 2 * ALoopCount - ( ALoopCount - 1 );
    if ASeed >= ALength then Break;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.Button1Click(Sender: TObject);
begin
    Memo1.Lines.Add(Inttostr(TUstr.binToInt('1010')));
    Memo1.Lines.Add(Inttostr(TUstr.binToInt('000000000100000101000001')));
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.TransNetworkIdInfo;
var
  sL_CompCode : STring;
  L_Obj : TNetworkIDInfoObj;
begin
  if not Assigned( G_NetworkIdInfoStrList ) then Exit;
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
      G_NetworkIdInfoStrList.AddObject( sL_CompCode, L_Obj );
      Next;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

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

{ ---------------------------------------------------------------------------- }

function TfrmMain.WriteRegData(const sI_EncryptedLicenseKey: STring): Boolean;
var
  L_StrList : TStringList;
  sL_LicenseKey : String;
begin
  Result := False; 
  sL_LicenseKey := TUother.myDecrypt( sI_EncryptedLicenseKey, ENCRYPTION_KEY );
  L_StrList := TUstr.ParseStrings( sL_LicenseKey,ENCRYPTION_SEP );
  try
    if ( L_StrList.Count = 3 ) and
       ( TUdateTime.IsDateStr( L_StrList.Strings[0], '/' ) ) and
       ( ( L_StrList.Strings[1] = SYS4_Y ) or
         ( L_StrList.Strings[1] = SYS4_N ) ) and
       ( ( L_StrList.Strings[2] = SYS5_Y ) or
         ( L_StrList.Strings[2] = SYS5_N ) ) then
    begin
      //表示 License Key 格式正確...
      //寫入軟體之使用到期日
      TUother.RegWrite(REGISTRY_ROOT,'SYS1',3,StrToDate(L_StrList.Strings[0]));
      //寫入軟體之安裝日
      TUother.RegWrite(REGISTRY_ROOT,'SYS2',3,Date);
      //寫入軟體之上次使用日期
      TUother.RegWrite(REGISTRY_ROOT,'SYS3',3,Date);
      //寫入是否啟動軟體保護機制
      TUother.RegWrite(REGISTRY_ROOT,'SYS4',2,L_StrList.Strings[1]);
      //寫入是否停用本軟體
      TUother.RegWrite(REGISTRY_ROOT,'SYS5',2,L_StrList.Strings[2]);
      Result := True;
     end;
   finally
     L_StrList.Free;
   end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.Timer1Timer(Sender: TObject);
begin
  ( Sender as TTimer ).Enabled := False;
  Timer1.Interval := reDefineFetchDataInterval(now);
  frmMain.bG_StartFetching := True;
  FetchComDataThreadU.onTimerStartAction;
  ( Sender as TTimer ).Enabled := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.Timer2Timer(Sender: TObject);
begin
  ( Sender as TTimer ).Enabled := False;
  HandleResponseU.OnTimerStartAction;
  ( Sender as TTimer ).Enabled := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.Timer3Timer(Sender: TObject);
begin
   ( Sender as TTimer ).Enabled := False;
     SendCommandU.onTimerStartAction;
   (Sender as TTimer ).Enabled := True;
end;

{ ---------------------------------------------------------------------------- }

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

{ ---------------------------------------------------------------------------- }

function TfrmMain.GetLanguageSettingInfo(const AKey: String): String;
var
  ALanguage : String;
begin
  Result := AKey;
  if Assigned( G_LanguageSettingIni ) then
  begin
    ALanguage := G_LanguageSettingIni.ReadString( 'CURRENT_LANGUAGE_TYPE',
      'CURRENT_LANGUAGE_TYPE', 'TC' );
    Result := G_LanguageSettingIni.ReadString( ALanguage, AKey, '' );
  end
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.InitLanguageIniFile;
var
  AFileName: String;
begin
  AFileName := IncludeTrailingPathDelimiter( ExtractFilePath(
    Application.ExeName ) ) + LANGUAGE_SETTING_INI_FILE_NAME;
  if FileExists( AFileName ) then
    G_LanguageSettingIni := TIniFile.Create( AFileName );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.BuildReturnPathConnection;
begin
  if not bG_HasBuildReturnPathConnection then
  begin
    { 啟動 Feedback Command 接收 }
    if G_ReturnPathSvr <> nil then
     G_ReturnPathSvr.BuildConnection( FeedbackSocket );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.ToolButton2Click(Sender: TObject);
begin
  livStatus.Items.Clear;
  createLivItem;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.ReadTransNumber;
var
  aFileName: String;
  aIniFile: TIniFile;
  aSeq1, aSeq2, aSeq3, aSeq4: Integer;
begin
  aFileName := IncludeTrailingPathDelimiter( ExtractFilePath(
  Application.ExeName ) ) + 'TransactionNumber.ini';
  aIniFile := TIniFile.Create( aFileName );
  try
    aSeq1 := aIniFile.ReadInteger( 'TransactionNumber', 'Z1', 0 );
    aSeq2 := aIniFile.ReadInteger( 'TransactionNumber', 'Real', 0 );
    aSeq3 := aIniFile.ReadInteger( 'TransactionNumber', 'FeedBack_Z1', 0 );
    aSeq4 := aIniFile.ReadInteger( 'TransactionNumber', 'FeedBack_Real', 0 );
  finally
    aIniFile.Free;
  end;
  frmMain.nG_FD_TestingCmdSeqNo := aSeq1;
  frmMain.nG_SC_RealCmdSeqNo := aSeq2;
  frmMain.nG_FD_TestingCmdSeqNo2 := aSeq3;
  frmMain.nG_SC_RealCmdSeqNo2 := aSeq4;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.SaveTransNumber;
var
  aFileName: String;
  aIniFile: TIniFile;
begin
  aFileName := IncludeTrailingPathDelimiter( ExtractFilePath(
    Application.ExeName ) ) + 'TransactionNumber.ini';
  aIniFile := TIniFile.Create( aFileName );
  try
    aIniFile.WriteInteger( 'TransactionNumber', 'Z1', frmMain.nG_FD_TestingCmdSeqNo );
    aIniFile.WriteInteger( 'TransactionNumber', 'Real', frmMain.nG_SC_RealCmdSeqNo );
    aIniFile.WriteInteger( 'TransactionNumber', 'FeedBack_Z1', frmMain.nG_FD_TestingCmdSeqNo2 );
    aIniFile.WriteInteger( 'TransactionNumber', 'FeedBack_Real', frmMain.nG_SC_RealCmdSeqNo2 );
    aIniFile.UpdateFile;
  finally
    aIniFile.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmMain.tbnHistoryClick(Sender: TObject);
begin
  _AUTO_SCROLL := tbnHistory.Down;
end;

end.



