unit frmMainU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ScktComp, StdCtrls, Sockets, inifiles, 
  comctrls, syncobjs,  winsock, DB, DBClient, Menus, ExtCtrls, ImgList, TListernerU, ConstU,
  Buttons ;

type
  TfrmMain = class(TForm)
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    StatusBar1: TStatusBar;
    Panel1: TPanel;
    Panel3: TPanel;
    Panel2: TPanel;
    Splitter1: TSplitter;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    imlStatus: TImageList;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    livStatus: TListView;
    memErrorLog: TMemo;
    N5: TMenuItem;
    N6: TMenuItem;
    N7: TMenuItem;
    Dispatcher21: TMenuItem;
    TcpClient1: TTcpClient;
    TabSheet2: TTabSheet;
    BitBtn2: TBitBtn;
    Edit1: TEdit;
    StaticText1: TStaticText;
    BitBtn1: TBitBtn;
    Button1: TButton;
    Memo2: TMemo;
    Memo1: TMemo;
    Memo4: TMemo;
    Memo3: TMemo;
    Edit2: TEdit;
    StaticText2: TStaticText;
    Button2: TButton;
    BitBtn3: TBitBtn;
    BitBtn4: TBitBtn;
    Button3: TButton;
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);

    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure Button1Click(Sender: TObject);
    procedure N7Click(Sender: TObject);
    procedure Dispatcher21Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure BitBtn4Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
  private
    { Private declarations }
    sG_AllSignatureKey, sG_CountryNumber : String;
    nG_TimeZoneOffset1, nG_TimeZoneOffset2 : Integer;
    sG_PinControl, sG_PinCode, sG_HighLevelCmdId : String;
    sG_CardID , sG_RegionKey , sG_ReportbackAvailability , sG_ReportbackDate , sG_BackwardTolerance , sG_ForwardTolerance , sG_Currency , sG_PopulationID : String;
    sG_Version, sG_SecurityType : String;
    sG_ToID : String;
    nG_ConnectionID, nG_SequenceID : Integer;
    G_CmdMatrixStrList : TStringList;
    G_ActionTypeAndPriorityStrList : TStringList;

    G_SeqTransMappingStrList : TStringList;
    G_ClientIpStrList : TStringList;
    G_AryOfTcpClient:  array of TTcpClient;


    G_ListenerThread : TListerner;
    G_ListenerThread_ExitCode : Cardinal;


    //down, 0328
    nG_LivItemPos : Integer;
    sG_SrcTransNum : String;// 會被填入 Nagra  傳來的 transaction num => return path 時會用到
    bG_SendComm : boolean;
    nG_MaxRealCmdSeqNo, nG_RealCmdSeqNo, nG_CCPort, nG_CompCode : Integer;
    sG_CompName, sG_FullCommand, sG_IrdDataLength, sG_IrdOperation, sG_IrdCmdID, sG_IrdData : String;
    nG_CreditMode, nG_CreditLimit, nG_Price, nG_ThresholdCredit, nG_Credit : Integer;

    sG_RealCommand, sG_MisIrdCmdID, sG_MisIrdCmdData : String;
    sG_FirstCallbackDate, sG_PhoneNum1, sG_PhoneNum2, sG_PhoneNum3 : String;
    sG_EventName, sG_CCNum1, sG_IpAddr, sG_CallbackTime, sG_CallbackDate : String;
    sG_MisSubBeginDate, sG_MisSubEndDate,sG_CallFreq, sG_Operator, sG_CommandStatus,  sG_StbNo : String;
    sG_IccNo, sG_SubBeginDate, sG_SubEndDate, sG_ImsProductID, sG_Notes, sG_ZipCode, sG_CurActiveLowLevelCmd, sG_CommandNo, sG_SeqNo : String;

    G_ErrorMsgStrList : TStringList;
    G_TimeStampStrList : TStringList;    
    //up, 0328


    bG_ShowUI, bG_CommandLog, bG_ErrorLog, bG_ResponseLog : boolean;
    nG_CmdRefreshRate1, nG_CmdRefreshRate2, nG_CmdResentTimes : Integer;
//    sG_TimeDefine : String;
    bG_LoginOk, bG_AutoResponse, bG_AutoInvoke : boolean;
    sG_ControlCommandType : String;
    sG_ProductCommandType : String;
    sG_FeedbackCommandType : String;
    sG_OperationCommandType : String;
    sG_EmmCommandType, sG_EmmBroadcastMode, sG_EmmAddressType : String;


    sG_LogPath, sG_ServerIP : String;
    nG_MaxCommandCount, nG_Timeout, nG_RPortNo, nG_SPortNo : Integer;
    bG_Debug : boolean;
    sG_ProcessorIp : String;
    sG_NewIniFileName,sG_Notes_Has_Value, sG_WriteTimeStamp : String;
    sG_ProcessorListenPort, sG_DispatcherListenPort : String;
    G_NotesHasValueStrList : TStringList;
    G_SignatureKey: TStringList;
    function getErrorMsg(sI_ErrorCode:String):String;
    procedure transIntoSignatureKeyStrList(sI_AllSignatureKey:String);
    function getSignatureKey(sI_FromID : String):String;
    procedure createTcpClient;
    procedure preSendCmd(sI_FullCmd : String);
    function getSignature(sI_Data, sI_Key, sI_SecurityType:String) : String;
    function getActionBody(sI_ActionType, sI_LowLevelCmdID, sI_Notes : String):String;
    procedure getActionTypeAndPriority(sI_HighLevelCommand:String; var sI_ActionType, sI_ActionPriority, sI_PriorityReassignment, sI_FromID : String); //0328
    procedure loadErrorMsgContent(I_StrList : TStringList);
    procedure getDispatcherSeqNo(sI_SequenceID:String; var sI_CompCode, sI_HighLevelCmdID, sI_DispatcherSequenceNo:String);

    procedure clearStatusItem(nI_ItemIndex:Integer); overload;
    procedure createLivItem;
    procedure recordSeqAndTransMapping(sI_CompCode, sI_SeqNo, sI_SequenceID, sI_HighLevelCmdID:String);overload;
    procedure showUI1(nI_CompCode:Integer; sI_CompName, sI_SeqNo, sI_CommandID, sI_Operator, sI_TransNum : String);
    procedure showUI2(sI_TransNum, sI_ErrorCode, sI_ErrorMsg:String; bI_HasError:boolean);
    procedure sendToDispatcher(nI_Type:Integer; sI_CompID, sI_Info, sI_HighLevelCmdID, sI_Operator, sI_SrcTransNum, sI_ResponseCmdID:String);
    procedure preSendToDispatcher(sI_CompID, sI_HighLevelCmdID, sI_DispatcherSeqNo, sI_ResponseDataStatus, sI_ErrorCode, sI_ErrorMsg, sI_SubscriberID, sI_StatusCode:String);
    procedure writeTimeStamp;//0328
    function isCmdDataOK(var nI_ResultFlag:Integer) : boolean; //0328
    function getDetailCommandIDs(sI_HighLevelCommand:String):String; //0328

    procedure parseCommandData(I_CmdInfoObject:TCmdInfoObject); //0328


    procedure releaseResource; //0328

    procedure getReturnDataForDispatcher(sI_Action, sI_SequenceID:String);
    function getFormatIpAddr(sI_SrcIpAddr:String):String; //0328
    procedure createThread;
    procedure saveThreadInfo(sI_Info:String);
    function getSysIniSettiing:String;
    function getSendCommandServerIP:STring;
    function getSendCommandPort:Integer;
//    function getBinaryVal(nI_Value: Integer; nI_Digits: Integer; bI_Reverse: boolean):TTempRecord;
    procedure buildNdsConnectionForSendCmd(var I_ClientSocket : TClientSocket; var I_WinSocketStream : TWinSocketStream);overload;
    function getCardIdChecksum(fI_CardID:double):Integer;
    function Write(I_Buf: Pointer; I_Count: Integer; I_WStream: TWinSocketStream):Integer;
    function Increment(Ptr: Pointer; Increment: Integer): Pointer;
    procedure terminateApp;
    procedure GetSysParam;
    procedure sendDisconnectInfoToDispatcher;
    procedure myTcpClient1Disconnect(Sender: TObject);
    procedure myTcpClient1Error(Sender: TObject; SocketError: Integer);
  public
    { Public declarations }
    G_ResponseMsgStrList: TStringList;
    G_RemoteInfoStrList : TStringList;
    G_SendCmdClntSocket: TClientSocket;
    G_SendCmdWStream : TWinSocketStream;
    sG_LoginID, sG_LoginName : String;
    bG_buildSendCmdConnection : boolean;    
    procedure parseClientNdsData(sI_ClientNdsData:String);
    procedure dispatcherConnectSettingCmd(sI_RemoteIP, sI_SettingStr:String);
    procedure dispatcherDisconnectSettingCmd(sI_RemoteIP : String);    

    procedure setTransCaption;
    procedure getErrorInfo(sI_ResErrorCode,sI_ResErrorCodeExt: String; var sI_ResErrorCodeDetail,sI_ResErrorCodeExtDetail: String);
    procedure handleCmdResponse(sI_ResponseStr:String);
    function reDefineFetchDataInterval(dI_Time:TDate):Integer;
    procedure reBuildNdsConn; //0328
    procedure showLogOnMemo(nI_ActiveConn:Integer; sI_Detail:String); //0328
    procedure activeCDS(I_CDS:TClientDataSet; sI_RefFileName:String; sI_DefaultStatus:String);
  end;

var
  frmMain: TfrmMain;

implementation

uses Ustru, UdateTimeu, frmSysParamU, frmLoginU, dllNdsU, frmAboutU,
  frmDispatcherInfoU, dtmMainU;

{$R *.dfm}

procedure TfrmMain.FormShow(Sender: TObject);
begin
    {
    TcpServer2.LocalHost := sG_ProcessorIp;
    TcpServer2.LocalPort := sG_ProcessorListenPort;
    TcpServer2.Open;
    }
    createThread;
    createLivItem;
end;

function TfrmMain.getSysIniSettiing: String;
var
    L_IniFile : TIniFile;
    sL_ExeFileName, sL_ExePath, sL_IniFileName : STring;

begin

    sL_ExeFileName := Application.ExeName;
    sL_ExePath := ExtractFileDir(sL_ExeFileName);

    if sG_NewIniFileName='' then
    begin
      sG_NewIniFileName := CA_INI_FILE_NAME;
      sL_IniFileName := sL_ExePath + '\' + CA_INI_FILE_NAME;
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

    sG_ProcessorListenPort := L_IniFile.ReadString('CA_DISPATCHER_PROCESSOR','PROCESSOR_LISTEN_PORT','1680');
    sG_DispatcherListenPort := L_IniFile.ReadString('CA_DISPATCHER_PROCESSOR','DISPATCHER_LISTEN_PORT','1681');

    sG_ProcessorIp := L_IniFile.ReadString('CA_DISPATCHER_PROCESSOR','PROCESSOR_IP','');

    if (sG_ProcessorIp='')  then
    begin
      result := '-1: 請設定參數檔<'+sL_IniFileName +'>之 PROCESSOR_IP ';
      exit;
    end;
    if not assigned(G_CmdMatrixStrList) then
      G_CmdMatrixStrList := TStringList.Create;

    if not assigned(G_ActionTypeAndPriorityStrList) then
      G_ActionTypeAndPriorityStrList := TStringList.Create;


    L_IniFile.ReadSectionValues('COMMANDMATRIX',G_CmdMatrixStrList);
    L_IniFile.ReadSectionValues('ACTIONTYPE',G_ActionTypeAndPriorityStrList);
        
    G_NotesHasValueStrList := TUStr.ParseStrings(sG_Notes_Has_Value,',');



    L_IniFile.Free;
    result := '';
end;

procedure TfrmMain.FormCreate(Sender: TObject);
var
    ii : Integer;
    sL_Result : String;

begin

    bG_AutoResponse := false;
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
          if UpperCase(ParamStr(ii)) = 'AR' then
          begin
            bG_AutoResponse := true;
          end;
        end;
    end;
    //系統登入畫面



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


      nG_MaxRealCmdSeqNo := StrToIntDef(TUstr.AddString('','9',true,SEQ_NUM_LENGTH),999999);


      try
        //down, 從 ini 檔讀取 information
        sL_Result := getSysIniSettiing;
        if sL_Result<>'' then
        begin
          if not bG_AutoInvoke then
            MessageDlg(sL_Result,mtError,[mbOK],0);
          Application.Terminate;
          exit;
        end;
        //up, 從 ini 檔讀取 information
        caption := caption + ' - ('+sG_LoginName+')' + '--' + sG_NewIniFileName;

      except
      end;
      try
        buildNdsConnectionForSendCmd(G_SendCmdClntSocket,G_SendCmdWStream);
      except
        frmMain.showLogOnMemo(REBUILD_NDS_CONNECTION_FAILED,' frmMain.reBuildNdsConn .程式即將結束,然後重新啟動. '); //執行完此行程式後, 程式會結束掉
      end;


//      G_ResponseMsgStrList := TStringList.Create;
      {
      G_CommandInfoStrList := TSTringList.Create;
      }
      G_TimeStampStrList := TStringList.Create;
      G_RemoteInfoStrList := TStringList.Create;
      G_ClientIpStrList := TStringList.Create;
      G_SeqTransMappingStrList := TStringList.Create;
      G_SignatureKey := TStringList.Create;
      transIntoSignatureKeyStrList(sG_AllSignatureKey);


      G_ErrorMsgStrList := TStringList.Create;
      loadErrorMsgContent(G_ErrorMsgStrList);



      if not assigned(G_CmdMatrixStrList) then
        G_CmdMatrixStrList := TStringList.Create;

      if not assigned(G_ActionTypeAndPriorityStrList) then
        G_ActionTypeAndPriorityStrList := TStringList.Create;
        


      nG_SequenceID := 1;
    end
    else
      Application.Terminate;

end;
procedure TfrmMain.parseClientNdsData(sI_ClientNdsData:String);
var
    nL_SeedValue, nL_NotesLen : Integer;
    sL_NotesLen, sL_Notes : String;
    sL_CompCode, sL_ServiceID, sL_SeqNo : String;


    sL_Operator : String;
    sL_IccNo,sL_HighLevelCmdId, sL_SubBeginDate, sL_SubEndDate : String;

    sL_CompName,sL_RegionKey, sL_Reportbackavailability, sL_SubscriberID,sL_PinCode: String;
    L_CmdInfoObject : TCmdInfoObject;
    sL_PopolcationID,sL_ReportbackDate, sL_PinControl : String;
begin
    try
      Application.ProcessMessages;
      {
      //傳入的字串如下(包含所有空白)
          範例：'00123456  9    陽明山  A1    2DC3  1234567890        2003043088992398760054E28Howard      11thisisnotesL'
      }
      sL_SeqNo := trim(Copy(sI_ClientNdsData,1,SEQ_NO_LENGTH));
      sL_CompCode := trim(Copy(sI_ClientNdsData,9,COMP_CODE_LENGTH));
      sL_CompName := trim(Copy(sI_ClientNdsData,12, COMP_NAME_LENGTH) );
      sL_HighLevelCmdId := trim(Copy(sI_ClientNdsData,22,HIGH_LEVEL_CMD_ID_LENGTH));
      sL_SubscriberID := trim(Copy(sI_ClientNdsData,26,SUBSCRIBER_ID_LENGTH));
      sL_IccNo := trim(Copy(sI_ClientNdsData,34,ICC_NO_LENGTH));
      sL_SubBeginDate := trim(Copy(sI_ClientNdsData,46,SUB_BEGIN_DATE_LENGTH));
      sL_SubEndDate := trim(Copy(sI_ClientNdsData,54,SUB_END_DATE_LENGTH));
      sL_PinCode := trim(Copy(sI_ClientNdsData,62,PIN_CODE_LENGTH));
      sL_PopolcationID :=trim(Copy(sI_ClientNdsData,66,POPULATION_ID_LENGTH));
      sL_RegionKey := trim(Copy(sI_ClientNdsData,68,REGION_KEY_LENGTH));
      sL_Reportbackavailability := trim(Copy(sI_ClientNdsData,76,REPORTBACK_AVAILABILITY_LENGTH));
      sL_ReportbackDate := trim(Copy(sI_ClientNdsData,77,REPORTBACK_DATE_LENGTH));
      sL_Operator := trim(Copy(sI_ClientNdsData,79,OPERATOR_LENGTH));

      sL_NotesLen := trim(Copy(sI_ClientNdsData,89,NOTES_LENGTH));
      nL_NotesLen := StrToIntDef(sL_NotesLen,0);
      sL_Notes := trim(Copy(sI_ClientNdsData,93,nL_NotesLen));

      nL_SeedValue := 93 + nL_NotesLen;
      sL_PinControl := trim(Copy(sI_ClientNdsData,nL_SeedValue,PIN_CONTROL_LENGTH));
      nL_SeedValue := nL_SeedValue+PIN_CONTROL_LENGTH;
      sL_ServiceID :=  Trim(Copy(sI_ClientNdsData,nL_SeedValue,SERVICE_ID_LENGTH));

      L_CmdInfoObject := TCmdInfoObject.Create;
      L_CmdInfoObject.SeqNo := sL_SeqNo;
      L_CmdInfoObject.CompCode := sL_CompCode;
      L_CmdInfoObject.CompName := sL_CompName;
      L_CmdInfoObject.HighLLevelCmdID := sL_HighLevelCmdId;
      L_CmdInfoObject.SubscriberID := sL_SubscriberID;
      L_CmdInfoObject.IccNo := sL_IccNo;
      L_CmdInfoObject.SubscriptionBeginDate := sL_SubBeginDate;
      L_CmdInfoObject.SubscriptionEndDate := sL_SubEndDate;
      L_CmdInfoObject.PinCode := sL_PinCode;
      L_CmdInfoObject.PopulationID := sL_PopolcationID;
      L_CmdInfoObject.RegionKey := sL_RegionKey;
      L_CmdInfoObject.ReportbackAvailability := sL_Reportbackavailability;
      L_CmdInfoObject.ReportbackDate := sL_ReportbackDate;
      L_CmdInfoObject.NOTES := sL_Notes;
      L_CmdInfoObject.Operator := sL_Operator;
      L_CmdInfoObject.PinControl := sL_PinControl;
      L_CmdInfoObject.ServiceID := sL_ServiceID;

      parseCommandData(L_CmdInfoObject);
      L_CmdInfoObject.Free;

      bG_Debug := true;//測試用...
      if bG_Debug then
      begin
        Memo1.Lines.Clear;
        Memo1.Lines.Add('sL_HighLevelCmdId=>' + sL_HighLevelCmdId);
        Memo1.Lines.Add('sL_IccNo=>' + sL_IccNo);
        Memo1.Lines.Add('sL_SubBeginDate=>' + sL_SubBeginDate);
        Memo1.Lines.Add('sL_SubEndDate=>' + sL_SubEndDate);
        Memo1.Lines.Add('sL_Operator=>' + sL_Operator);
        Memo1.Lines.Add('sL_SeqNo=>' + sL_SeqNo);
        Memo1.Lines.Add('sL_CompCode=>' + sL_CompCode);
        Memo1.Lines.Add('sL_NotesLen=>' + sL_NotesLen);
        Memo1.Lines.Add('sL_Notes=>' + sL_Notes);

        Memo1.Lines.Add('sL_SubscriberID=>' + sL_SubscriberID);
        Memo1.Lines.Add('sL_PinCode=>' + sL_PinCode);
        Memo1.Lines.Add('sL_PopolcationID=>' + sL_PopolcationID);
        Memo1.Lines.Add('sL_RegionKey=>' + sL_RegionKey);
        Memo1.Lines.Add('sL_Reportbackavailability=>' + sL_Reportbackavailability);
        Memo1.Lines.Add('sL_ReportbackDate=>' + sL_ReportbackDate);

        Memo1.Lines.Add('sL_PinControl=>' + sL_PinControl);
        Memo1.Lines.Add('sL_ServiceID=>' + sL_ServiceID);


      end;
    except

    end;
end;


procedure TfrmMain.buildNdsConnectionForSendCmd(
  var I_ClientSocket: TClientSocket;
  var I_WinSocketStream: TWinSocketStream);
var
    nL_Words : Integer;
//    L_TmpRecord : TTempRecord;
    ii, nL_CmdLen : Integer;
    L_Socket : TClientSocket;
    L_WStream: TWinSocketStream;
//    L_DeviceCall : TDeviceCall;
//    L_DeviceStatus : TDeviceStatus;
//    L_DeviceAnswer : TDeviceAnswer;
    sL_CmdObjName : String;
begin
    if Assigned(I_ClientSocket) then
    begin
      I_ClientSocket := nil;
      I_ClientSocket.Free;
    end;
    I_ClientSocket:= TClientSocket.Create(nil);
    I_ClientSocket.ClientType := ctBlocking;


    L_Socket := I_ClientSocket;
    try
    begin
      with L_Socket do
      begin

        Address := getSendCommandServerIP;
        Port := getSendCommandPort;

{
        Address := '203.193.62.251';
        Port := 12222;
}


//        Open;

        L_WStream := TWinSocketStream.Create(L_Socket.Socket, 60000);

        I_WinSocketStream := L_WStream;



      end;

    end;
    except

      raise;
      exit;
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

{
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
}

function TfrmMain.Increment(Ptr: Pointer; Increment: Integer): Pointer;
begin
    Result := Pointer(Longint(Ptr) + Increment);
end;

function TfrmMain.Write(I_Buf: Pointer; I_Count: Integer;
  I_WStream: TWinSocketStream): Integer;
var
    nL_Idx: Integer;
begin
    nL_Idx := 0;
    while (nL_Idx < I_Count) do
      nL_Idx := nL_Idx + I_WStream.Write(Increment(I_Buf, nL_Idx)^, I_Count-nL_Idx);

    result := nL_Idx;
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
      if (FileExists(sL_Path + PROCESSOR_SYS_PARAM_FILENAME)) then
        LoadFromFile(sL_Path + PROCESSOR_SYS_PARAM_FILENAME);

      sG_ServerIP := FieldByName('sServerIP').AsString;
      sG_LogPath := FieldByName('sLogPath').AsString;
      nG_SPortNo := FieldByName('nSPortNo').AsInteger;
      nG_RPortNo := FieldByName('nRPortNo').AsInteger;
      nG_Timeout := FieldByName('nTimeOut').AsInteger;
//      nG_MaxCommandCount := FieldByName('nMaxCommandCount').AsInteger;
      bG_CommandLog := FieldByName('bCommandLog').Asboolean;

      bG_ResponseLog := FieldByName('bResponseLog').Asboolean;  //欄位命名錯誤....



//      nG_CmdRefreshRate1 := FieldByName('nCmdRefreshRate1').AsInteger; //尖峰時段?秒讀取一次指令資料庫
//      nG_CmdRefreshRate2 := FieldByName('nCmdRefreshRate2').AsInteger; //離峰時段?秒讀取一次指令資料庫

//      nG_CmdResentTimes := FieldByName('nCmdResentTimes').AsInteger;
      bG_ShowUI := FieldByName('bShowUI').Asboolean;
//      sG_TimeDefine := FieldByName('sSHotTime').AsString + '-' + FieldByName('sEHotTime').AsString; // 尖峰時段定義

//***************************************************************


      sG_ToID := IntToHex(FieldByName('nToID').AsInteger,4);


      sG_Version := IntToHex(FieldByName('nVersion').AsInteger,4);


      sG_SecurityType := FieldByName('sSecurityType').AsString;
      nG_ConnectionID := FieldByName('nConnectionID').AsInteger;

      sG_AllSignatureKey := FieldByName('sKey').AsString; //0,00000000~1,12345678~2,12345678


      sG_BackwardTolerance := IntToHex(FieldByName('nBackwardTolerance').AsInteger,2);

      sG_ForwardTolerance := IntToHex(FieldByName('nForwardTolerance').AsInteger,2);

      sG_Currency := IntToHex(FieldByName('nCurrency').AsInteger,4);

      sG_CountryNumber := FieldByName('sCountryNumber').AsString;
      nG_TimeZoneOffset1 := FieldByName('nTimeZoneOffset').AsInteger;
      nG_TimeZoneOffset2 := FieldByName('nTimeZoneOffset2').AsInteger; 


      if RecordCount=0 then
        MessageDlg('系統參數尚未設定! 請先設定後,再執行程式.', mtWarning, [mbOK],0);
      Close;
    end;
end;

procedure TfrmMain.terminateApp;
begin
    if MessageDlg('是否確定要結束程式?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      sendDisconnectInfoToDispatcher;
      Close;
    end;
end;

procedure TfrmMain.N4Click(Sender: TObject);
begin
    terminateApp;
end;

procedure TfrmMain.N2Click(Sender: TObject);
begin
    frmSysParam := TfrmSysParam.Create(Application);
//    frmSysParam.sG_LoginName := sG_LoginName;
    frmSysParam.ShowModal;
    frmSysParam.free;

end;

procedure TfrmMain.activeCDS(I_CDS: TClientDataSet; sI_RefFileName,
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





procedure TfrmMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    {
    if TcpServer2.Active then
      TcpServer2.Close;
    TcpServer2.Free;
    }
    sendDisconnectInfoToDispatcher;
    releaseResource;    
    Action := caFree;
end;


function TfrmMain.getDetailCommandIDs(sI_HighLevelCommand: String): String;
var
    sL_Result : String;
    ii : Integer;
begin
    sL_Result := '';
    for ii:=0 to G_CmdMatrixStrList.Count -1 do
    begin
      if sI_HighLevelCommand = Copy(G_CmdMatrixStrList.Strings[ii],1,2) then
      begin
        sL_Result := Copy(G_CmdMatrixStrList.Strings[ii],4, length(G_CmdMatrixStrList.Strings[ii])-3);
        break;
      end;
    end;

  result := sL_Result;
end;

function TfrmMain.isCmdDataOK(var nI_ResultFlag: Integer): boolean;
begin
    nI_ResultFlag := 0;
    if (length(sG_IccNo) <> 0) and  (length(sG_IccNo) <> REAL_ICC_CARD_NO_LENGTH) then
    begin
      nI_ResultFlag := -2;
      result := false;
      exit;
    end;

    result := true;

end;

procedure TfrmMain.writeTimeStamp;
var
    sL_ExeFileName, sL_ExePath, sL_TimeStampFileName : String;
    sL_Method ,sL_Info, sL_CurTimeStamp : String;
begin
    if (UpperCase(sG_WriteTimeStamp)='Y')  then
    begin
      sL_CurTimeStamp := DateTimeToStr(now);    
      sL_ExeFileName := Application.ExeName;
      sL_ExePath := ExtractFileDir(sL_ExeFileName);
      sL_TimeStampFileName := sL_ExePath + '\' + TIME_STAMP_FILE_NAME;
      if FileExists(sL_TimeStampFileName) then
        frmMain.G_TimeStampStrList.LoadFromFile(sL_TimeStampFileName);
      frmMain.G_TimeStampStrList.Clear;
      frmMain.G_TimeStampStrList.Add(sL_CurTimeStamp);
      frmMain.G_TimeStampStrList.SaveToFile(sL_TimeStampFileName);
    end;
end;


procedure TfrmMain.releaseResource;
begin
    if Assigned(G_ResponseMsgStrList) then
    begin
      G_ResponseMsgStrList.Free;
      G_ResponseMsgStrList := nil;
    end;

    if Assigned(G_ErrorMsgStrList) then
    begin
      G_ErrorMsgStrList.Free;
      G_ErrorMsgStrList := nil;
    end;


    if Assigned(G_ListenerThread) then
    begin
      G_ListenerThread.Free;
      G_ListenerThread := nil;
    end;



    if Assigned(G_TimeStampStrList) then
    begin
      G_TimeStampStrList.Free;
      G_TimeStampStrList := nil;
    end;

    if Assigned(G_RemoteInfoStrList) then
    begin
      G_RemoteInfoStrList.Free;
      G_RemoteInfoStrList := nil;
    end;

    if Assigned(G_ClientIpStrList) then
    begin
      G_ClientIpStrList.Free;
      G_ClientIpStrList := nil;
    end;
    if Assigned(G_SeqTransMappingStrList) then
    begin
      G_SeqTransMappingStrList.Free;
      G_SeqTransMappingStrList := nil;
    end;

    if Assigned(G_CmdMatrixStrList) then
    begin
      G_CmdMatrixStrList.Free;
      G_CmdMatrixStrList := nil;
    end;

    if Assigned(G_ActionTypeAndPriorityStrList) then
    begin
      G_ActionTypeAndPriorityStrList.Free;
      G_ActionTypeAndPriorityStrList := nil;
    end;

    if Assigned(G_SignatureKey) then
    begin
      G_SignatureKey.Free;
      G_SignatureKey := nil;
    end;


end;



procedure TfrmMain.reBuildNdsConn;
begin
    try
      bG_buildSendCmdConnection := false;

      G_SendCmdWStream.Free;
      G_SendCmdWStream := nil;

      G_SendCmdClntSocket.Close;
//        sleep(3000);
      buildNdsConnectionForSendCmd(G_SendCmdClntSocket,G_SendCmdWStream);
    except
      frmMain.showLogOnMemo(REBUILD_NDS_CONNECTION_FAILED,' frmMain.reBuildNdsConn .程式即將結束,然後重新啟動. '); //執行完此行程式後, 程式會結束掉
    end;

end;

procedure TfrmMain.showLogOnMemo(nI_ActiveConn: Integer;
  sI_Detail: String);
var
    bL_KillSelf, bL_ShowMemo : boolean;
    sL_DbAlias, sL_Msg : String;
    L_DbLogStrList : TStringList;
    sL_ExeFileName, sL_ExePath, sL_RealDbLogFileName : String;
//    L_Intf : IMonitor; //0328
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
      
      if nI_ActiveConn=SEND_CMD_ERROR_FLAG then
      begin
        bL_KillSelf := true;
        sL_DbAlias := '';
        sL_Msg := '['+ datetimetostr(now)+ '] ' +'無法傳送指令給 Nagra. ' + '  作業項目:' + sI_Detail ;

        writeMsg;
        releaseResource;
        //down, 呼叫監控員, 讓監控員把 CA 叫起來
        {
        L_Intf := CoMonitor.Create;
        L_Intf.RunCa(sG_NewIniFileName);
        L_Intf := nil;
        }
        //up, 呼叫監控員, 讓監控員把 CA 叫起來
      end
      else if nI_ActiveConn=REBUILD_NDS_CONNECTION_FAILED then
      begin
        bL_KillSelf := true;
        sL_DbAlias := '';
        sL_Msg := '['+ datetimetostr(now)+ '] ' +' 無法重新建立 Nagra Connection. ' + '  作業項目:' + sI_Detail ;

        writeMsg;
        releaseResource;
        //down, 呼叫監控員, 讓監控員把 CA 叫起來
        {
        L_Intf := CoMonitor.Create;
        L_Intf.RunCa(sG_NewIniFileName);
        L_Intf := nil;
        }
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
        {
        L_Intf := CoMonitor.Create;
        L_Intf.RunCa(sG_NewIniFileName);
        L_Intf := nil;
        }
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
        {
        L_Intf := CoMonitor.Create;
        L_Intf.RunCa(sG_NewIniFileName);
        L_Intf := nil;
        }
        //up, 呼叫監控員, 讓監控員把 CA 叫起來


      end

      else if nI_ActiveConn=REBUILD_SEND_CMD_CONNECTION_TO_NDS then
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
      else if nI_ActiveConn=RE_CONNECT_NDS_ERROR_FLAG then
      begin
        sL_DbAlias := '';
        sL_Msg := '['+ datetimetostr(now)+ '] ' + '無法與 Nagra 建立連線.'+ '  作業項目:' + sI_Detail;
        writeMsg;
      end
      else if nI_ActiveConn=CLOSE_SEND_NDS_FLAG then
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
        {
        L_Intf := CoMonitor.Create;
        L_Intf.RunCa(sG_NewIniFileName);
        L_Intf := nil;
        }
        //up, 呼叫監控員, 讓監控員把 CA 叫起來
      end
      else if nI_ActiveConn=EXECUTE_SEND_NDS_FLAG then
      begin
        sL_DbAlias := '';
        sL_Msg := '['+ datetimetostr(now)+']' + '  作業項目:' + sI_Detail + '. 啟動者:' + sG_LoginName;
        writeMsg;
      end;

      if bL_KillSelf then //未必會被執行到, 因為呼叫  L_Intf.RunCa 時, 就會把 ca command gateway kill 掉
        terminateApp;
  //      Application.Terminate;
    except

    end;
end;


function TfrmMain.getFormatIpAddr(sI_SrcIpAddr: String): String;
var
    L_StrList : TStringList;
    ii : Integer;
    sL_Result : String;
begin
    if sI_SrcIpAddr='' then
    begin
      result := '';
      exit;
    end;

    L_StrList := TUStr.ParseStrings(sI_SrcIpAddr,'.');
    for ii:=0 to L_StrList.Count -1 do
    begin
      L_StrList.Strings[ii] := TUstr.AddString(L_StrList.Strings[ii],'0',false,3);
    end;

    for ii:=0 to L_StrList.Count -1 do
    begin
      if ii=0 then
        sL_Result := L_StrList.Strings[ii]
      else
        sL_Result := sL_Result + '.' + L_StrList.Strings[ii];
    end;

    result := sL_Result;
end;

procedure TfrmMain.showUI1(nI_CompCode:Integer; sI_CompName, sI_SeqNo, sI_CommandID,
  sI_Operator, sI_TransNum: String);
var
    jj, nL_NextItemIndex : Integer;
    L_ListItem : TListItem;
begin
    try
      if nI_CompCode=(StrToInt(TESTING_CMD_COMP_CODE)) then Exit;
      if sI_Operator = BATCH_OPERATOR then Exit; // 批次作業 所產生的指令不顯示 UI
//      if StrToIntDef(Copy(sI_TransNo,4,6),0)=0 then Exit;
      if nG_LivItemPos=CA_UI_MAX_COUNT then
      begin
//        clearStatusItem; //呼叫此 function...似乎會讓 performance 下降...
        nG_LivItemPos := 0;
      end;
      L_ListItem := livStatus.Items[frmMain.nG_LivItemPos];



      for jj:=0 to LISTVIEW_SUB_ITEM_COUNT-1 do
       L_ListItem.SubItems.Add('');

      Inc(nG_LivItemPos);

      //down, 將下一個 ListItem 清除成空白
      if nG_LivItemPos=CA_UI_MAX_COUNT then
        nL_NextItemIndex := 0
      else
        nL_NextItemIndex := nG_LivItemPos;
      clearStatusItem(nL_NextItemIndex);
      //up, 將下一個 ListItem 清除成空白

//      L_ListItem.StateIndex := 0;
      L_ListItem.Caption := sI_TransNum;
      L_ListItem.SubItems[0] := sI_CompName; ;
      L_ListItem.SubItems[1] := sI_SeqNo;
      L_ListItem.SubItems[2] := sI_CommandID + '-' + sG_CurActiveLowLevelCmd;
      L_ListItem.SubItemImages[3] := 2;//打勾之icon  => 傳送
      L_ListItem.SubItemImages[4] := -1; //清除 icon => 回應
      L_ListItem.SubItems[7] := sI_Operator;
      L_ListItem.SubItems[8] := DateTimeToStr(now);





    except

    end;
end;

procedure TfrmMain.createLivItem;
var
    ii, jj : Integer;
    L_ListItem : TListItem;

begin
    if bG_ShowUI then
    begin
      nG_LivItemPos := 0;
      memErrorLog.Align := alBottom;
      memErrorLog.Height := 122;


      Splitter1.Visible := true;      
      Splitter1.Parent := Panel2;
      Splitter1.Height := 4;
      Splitter1.Align := alBottom;

      PageControl1.Visible := true;      
      PageControl1.Align := alClient;


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
      PageControl1.Visible := false;
      Splitter1.Visible := false;
      memErrorLog.Align := alClient;
    end;
end;

procedure TfrmMain.clearStatusItem(nI_ItemIndex: Integer);
var
    jj : Integer;
begin
      frmMain.livStatus.Items[nI_ItemIndex].Caption := '';
      frmMain.livStatus.Items[nI_ItemIndex].SubItemImages[3] := -1;
      frmMain.livStatus.Items[nI_ItemIndex].SubItemImages[4] := -1;
      frmMain.livStatus.Items[nI_ItemIndex].SubItemImages[5] := -1;
      for jj:=0 to frmMain.livStatus.Items[nI_ItemIndex].SubItems.Count -1 do
        frmMain.livStatus.Items[nI_ItemIndex].SubItems.Strings[jj] := '';

end;


procedure TfrmMain.createThread;
var
    sL_Info : String;
begin
    sL_Info := '程式即將結束,請重新執行此程式!';
    
    G_ListenerThread := TListerner.Create(sG_ProcessorIp, sG_ProcessorListenPort );
    if GetExitCodeThread(G_ListenerThread.ThreadID,G_ListenerThread_ExitCode) then
    begin
      if not bG_AutoInvoke then
        MessageDlg('執行緒 2 產生失敗!' + sL_Info, mtError, [mbOK],0);

      Application.Terminate;
      Exit;
    end;


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

function TfrmMain.reDefineFetchDataInterval(dI_Time: TDate): Integer;
var
    nL_Result, nL_Time1, nL_Time2, nL_Time3 : Integer;
    sL_Time1, sL_Time2 : String;
    wL_Hour, wL_Min, wL_Sec, wL_MSec : word;
begin
{
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
}    
end;


procedure TfrmMain.handleCmdResponse(sI_ResponseStr: String);
var
    nL_Seed : Integer;
    sL_Version, sL_SecurityType, sL_Length : String;
    sL_FromID, sL_ToID, sL_ConnectionID, sL_Date : String;
    sL_SequenceID, sL_Action, sL_Signature : String;
    sL_ErrorCode, sL_ErrorMsg : String;
begin
    {
    Memo3.Lines.Add(sI_ResponseStr);
    Exit; //測試...
    }
    Application.ProcessMessages;
    if length(sI_ResponseStr)<20 then Exit;

    try
    begin

      nL_Seed := 1;
      sL_Version := Copy(sI_ResponseStr,nL_Seed, VERSION_LENGTH);
      nL_Seed := nL_Seed + VERSION_LENGTH;
      sL_SecurityType := Copy(sI_ResponseStr,nL_Seed, 1);
      nL_Seed := nL_Seed + 1;
      sL_Length := Copy(sI_ResponseStr,nL_Seed, 4);

      nL_Seed := nL_Seed + 4;
      sL_FromID := Copy(sI_ResponseStr,nL_Seed, FROM_ID_LENGTH);

      nL_Seed := nL_Seed + FROM_ID_LENGTH;
      sL_ConnectionID := Copy(sI_ResponseStr,nL_Seed, 2);

      nL_Seed := nL_Seed + 2;
      sL_ToID := Copy(sI_ResponseStr,nL_Seed, TO_ID_LENGTH);


      nL_Seed := nL_Seed + TO_ID_LENGTH;
      sL_Date := Copy(sI_ResponseStr,nL_Seed, 14);

      nL_Seed := nL_Seed + 14;
      sL_SequenceID := Copy(sI_ResponseStr,nL_Seed, 4);

      nL_Seed := nL_Seed + 4;
      sL_Action := Copy(sI_ResponseStr,nL_Seed, length(sI_ResponseStr)-nL_Seed-16+1);
      getReturnDataForDispatcher(sL_Action, sL_SequenceID);

      nL_Seed := nL_Seed + length(sL_Action);
      sL_Signature := Copy(sI_ResponseStr,nL_Seed, 16);

      //action 的可能值(A0) : S000000051I00000001006600000000000000000000000000000000000000987600541C0000000000000000000000000000E051938000003O
        


//      frmMain.setTransCaption;



    end
    except
      on E: Exception do
      begin
        {
        frmMain.showLogOnMemo(WRITE_TO_MIS_ERROR_FLAG,' HandleResponse.writeToMis' + '=>' + E.Message);

        sL_ErrorInfoFileName := 'c:\Error4.txt';
        L_ErrorInfoStrList := TStringList.Create;
        if FileExists(sL_ErrorInfoFileName) then
          L_ErrorInfoStrList.LoadFromFile(sL_ErrorInfoFileName);
        if L_ErrorInfoStrList.Count>500 then
          L_ErrorInfoStrList.Clear;
        L_ErrorInfoStrList.Insert(0,'   Detail Msg: ConnectionID=' + IntToStr(nL_ConnectionID) +', Cmd Sequence No=' + sL_SeqNo + ',ErrorCode=' + sL_ErrorCode + '  ErrorMsg=' +  sL_ErrorMsg);

        L_ErrorInfoStrList.Insert(0,'   Msg:' +E.Message +', HelpContext:' + IntToStr(E.HelpContext));
        L_ErrorInfoStrList.Insert(0,'THandleResponse.writeToMis error:' + DateTimeToStr(now));
        L_ErrorInfoStrList.SaveToFile(sL_ErrorInfoFileName);
        L_ErrorInfoStrList.Free;



        sL_ErrorCode := WRITE_TO_MIS_ERROR;
        sL_ErrorMsg := WRITE_TO_MIS_ERROR_MSG;
        bL_WriteToMis := false;
        nL_CompCode := StrToIntDef(Copy(sL_ResTransNum2,1,3),0);
        }

        exit;
      end;
    end;

end;

procedure TfrmMain.dispatcherConnectSettingCmd(sI_RemoteIP, sI_SettingStr: String);
var
    sL_CompID : String;
    ii, nL_Ndx : Integer;
    L_TmpStrList : TStringList;
    L_Obj : TCompAndIpInfoObj;
begin
    //sI_SettingStr => 由許多公司別所串起來的字串. Ex:9,10,12

    L_TmpStrList := TUStr.ParseStrings(sI_SettingStr,COMP_ID_SEP);
    for ii:=0 to L_TmpStrList.Count -1 do
    begin
      sL_CompID := L_TmpStrList.Strings[ii];
      nL_Ndx := G_RemoteInfoStrList.IndexOf(sL_CompID);
      if nL_Ndx=-1 then
      begin
        L_Obj := TCompAndIpInfoObj.Create;
        L_Obj.IP := sI_RemoteIP;
        G_RemoteInfoStrList.AddObject(sL_CompID, L_Obj);

      end
      else
      begin
        L_Obj := G_RemoteInfoStrList.Objects[nL_Ndx] As TCompAndIpInfoObj;
        L_Obj.IP := sI_RemoteIP;
      end;
    end;



    nL_Ndx := G_ClientIpStrList.IndexOf(sI_RemoteIP);
    if nL_Ndx=-1 then
    begin
      G_ClientIpStrList.Add(sI_RemoteIP);
      setLength(G_AryOfTcpClient, G_ClientIpStrList.Count);
    end;

    //down, 建立 TcpClient 的 array...    
    createTcpClient;
    //up, 建立 TcpClient 的 array...end;

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
      sL_IniFileName := sL_ExePath + '\' + CA_INI_FILE_NAME
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

procedure TfrmMain.setTransCaption;
begin
    Panel1.Caption := ' 即時監控--指令傳送中';
    if Panel1.Font.Color = clBlue then
      Panel1.Font.Color := clRed
    else
      Panel1.Font.Color := clBlue;


    self.Invalidate;

end;


procedure TfrmMain.showUI2(sI_TransNum, sI_ErrorCode, sI_ErrorMsg: String;
  bI_HasError: boolean);
var
    nL_ImgNdx, nL_Pos : Integer;
    L_ListItem: TListItem;    
begin

     nL_Pos := 0;
     L_ListItem := frmMain.livStatus.FindCaption(nL_Pos,sI_TransNum,true,true,false);
     if L_ListItem<>nil then
     begin
       if bI_HasError then
       begin
         nL_ImgNdx := 3;
         L_ListItem.SubItems[5] := sI_ErrorCode;
         L_ListItem.SubItems[6] := sI_ErrorMsg;
       end
       else
       begin
         nL_ImgNdx := 2;
         L_ListItem.SubItems[5] := '';
         L_ListItem.SubItems[6] := '';
       end;

        L_ListItem.SubItemImages[4] := nL_ImgNdx; //清除 icon => 回應
     end;

end;

procedure TfrmMain.preSendToDispatcher(sI_CompID, sI_HighLevelCmdID, sI_DispatcherSeqNo, sI_ResponseDataStatus,
sI_ErrorCode, sI_ErrorMsg, sI_SubscriberID, sI_StatusCode:String);
var
    sL_Response : String;
    L_Obj : TCompAndIpInfoObj;
begin

    sL_Response := sI_ResponseDataStatus + Format('%.2d',[length(sI_CompID)]) + sI_CompID + Format('%.2d',[length(sI_DispatcherSeqNo)]) + sI_DispatcherSeqNo ;
    sL_Response := sL_Response + Format('%.3d',[length(sI_ErrorCode)]) + sI_ErrorCode ;
    sL_Response := sL_Response + Format('%.3d',[length(sI_ErrorMsg)]) + sI_ErrorMsg ;

    if (sI_HighLevelCmdID='A0') or (sI_HighLevelCmdID='A1') then
      sL_Response := sL_Response + sI_SubscriberID + sI_StatusCode;

    {
    可能的 response string :
    高階指令 A0, A1 => 'C01903123000000000000051' => C 01 9 03 123 000 000 00000005 1
    其他高階指令  => 'C01903123000000' => C 01 9 03 123 000 000     
    }
    sendToDispatcher(SEND_TO_DISPATCHER_INFO_TYPE_1,sI_CompID, sL_Response,'','','','');

end;


procedure TfrmMain.getDispatcherSeqNo(sI_SequenceID: String; var sI_CompCode, sI_HighLevelCmdID, sI_DispatcherSequenceNo:String);
var
    sL_SeqNo : String;
    nL_Ndx : Integer;
    L_Obj : TSequenceIdMappingInfoObj;
begin

    nL_Ndx := G_SeqTransMappingStrList.IndexOf(sI_SequenceID);
    if nL_Ndx<>-1 then
    begin

      L_Obj := G_SeqTransMappingStrList.Objects[nL_Ndx] as TSequenceIdMappingInfoObj;
      if L_Obj<>nil then
      begin
        sL_SeqNo := L_Obj.DispatcherSequenceNo;
        sI_CompCode := L_Obj.CompCode;
        sI_HighLevelCmdID := L_Obj.HighLevelCmdId;
        sI_DispatcherSequenceNo := L_Obj.DispatcherSequenceNo;
        G_SeqTransMappingStrList.Delete(nL_Ndx);
      end;
      
    end;

end;

procedure TfrmMain.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
    if (ssCtrl in Shift) and (Key=83) then //Ctrl + s
    begin
      bG_ShowUI := not bG_ShowUI;
      createLivItem;
    end;

end;

procedure TfrmMain.sendToDispatcher(nI_Type: Integer; sI_CompID, sI_Info, sI_HighLevelCmdID, sI_Operator, sI_SrcTransNum, sI_ResponseCmdID: String);
var
    sL_DispatcherIp, sL_Header : String;
    nL_Ndx,ii : Integer;
    L_Obj : TCompAndIpInfoObj;
begin
    try
      nL_Ndx := G_RemoteInfoStrList.IndexOf(sI_CompID);
      if nL_Ndx<>-1 then
      begin
        case nI_Type of
          SEND_TO_DISPATCHER_INFO_TYPE_1 :
            sL_Header := IntToStr(nI_Type);
          SEND_TO_DISPATCHER_INFO_TYPE_2 :
            sL_Header := IntToStr(nI_Type) + Format('%.2d',[length(sI_CompID)]) + sI_CompID + Format('%.4d',[length(sI_HighLevelCmdID)]) + sI_HighLevelCmdID + Format('%.3d',[length(sI_Operator)]) + sI_Operator;
          SEND_TO_DISPATCHER_INFO_TYPE_3 :
            sL_Header := IntToStr(nI_Type) + Format('%.2d',[length(sI_CompID)]) + sI_CompID + sI_SrcTransNum + sI_ResponseCmdID;
          else
            sL_Header := IntToStr(SEND_TO_DISPATCHER_INFO_DEFAULT_TYPE);
        end;

        L_Obj := G_RemoteInfoStrList.Objects[nL_Ndx] As TCompAndIpInfoObj;
        sL_DispatcherIp := L_Obj.IP;
        for ii:=0 to High(G_AryOfTcpClient) do
        begin
          if G_AryOfTcpClient[ii].RemoteHost = sL_DispatcherIp then
          begin
            Memo3.Lines.Add(sL_Header + sI_Info);
            G_AryOfTcpClient[ii].Sendln(sL_Header + sI_Info);

{
傳回的可能值

SEND_TO_DISPATCHER_INFO_TYPE_1(Ex:1C0190800123456000000) => 1(型別) C(指令處理結果) 01(公司代碼長度) 9(公司代碼) 08(SeqNo長度) 00123456(SeqNo) 000 (錯誤訊息代碼長度) [可能會有錯誤訊息代碼] 000 (錯誤訊息長度) [可能會有錯誤訊息代碼]
SEND_TO_DISPATCHER_INFO_TYPE_1(Ex:1C0190800123456000000000000051) => 1(型別) C(指令處理結果) 01(公司代碼長度) 9(公司代碼) 08(SeqNo長度) 00123456(SeqNo) 000 (錯誤訊息代碼長度) [可能會有錯誤訊息代碼] 000 (錯誤訊息長度) [可能會有錯誤訊息代碼] 00000005 (SubscriberID) 1 (StatusCode) 
SEND_TO_DISPATCHER_INFO_TYPE_2 => 2(型別)   01(公司代碼長度) 9(公司代碼) 0002(高階指令代碼長度) A1(高階指令代碼) 006(操作者長度) howard(操作者) 00900000301000302574428920030408N2003040720030409U1160194446005211911825900000 (指令字串)
SEND_TO_DISPATCHER_INFO_TYPE_3 => 3(型別)   01(公司代碼長度) 9(公司代碼) 009000004(交易號碼) 1000(回應之指令代號) 000008972050257000344289200303081000009000004000000000000000000000000(完整之回應字串)
}
            break;
          end;
        end;
      end;
    except

    end;
end;

procedure TfrmMain.Button1Click(Sender: TObject);
var
    sL_Data : String;
begin

    //down, 開機+頻道授權
    sL_Data := '00123456  9    陽明山  A000000000       10066        2003043088992398760054E28    Howard  254~20040531~N;5~20041231~A     4';
    parseClientNdsData(sL_Data);
    //產生之完整 command : '0003S00400001010100200304271617350001SIH    2DC3C00123456789098760054E1C050300030017A000420040531NA000520041231AAFC1115B3EBD670F';
    //up, 開機+頻道授權
{
    Memo2.Lines.Add('-----------------------');
    //down, 開機
    sL_Data := '00123456  9    陽明山  A1    2DC3  1234567890        2003043088992398760054E28    Howard   0     5';
    parseClientNdsData(sL_Data);
    //產生之完整 command : '0003S00240001010100200304271615170001SIH    2DC3C00123456789098760054E1C05030003001725027EB30300A98F';
    //up, 開機


    Memo2.Lines.Add('-----------------------');
    //down, 拆機
    sL_Data := '00123456  9    陽明山  A2    2DC3  1234567890        2003043088992398760054E28    Howard   0     6';
    parseClientNdsData(sL_Data);
    //產生之完整 command : '0003S00010001010100200304271633430003SHH    2DC3D60D6CC1683ADFF1B';
    //up, 拆機

    Memo2.Lines.Add('-----------------------');
    //down, 開頻道
    sL_Data := '00123456  9    陽明山  B1    2DC3  1234567890        2003043088992398760054E28    Howard  254~20040531~N;5~20041231~A     7';
    parseClientNdsData(sL_Data);
    //產生之完整 command : '0003S001C0001010100200304271638490004SIH    2DC3A000420040531NA000520041231A44569EB9D9BB88DD';
    //up, 開頻道

    Memo2.Lines.Add('-----------------------');
    //down, 關頻道
    sL_Data := '00123456  9    陽明山  B2    2DC3  1234567890        2003043088992398760054E28    Howard   57,8,9     8';
    parseClientNdsData(sL_Data);
    //產生之完整 command : '0003S000F0001010100200304271652010005SHH    2DC3X0007X0008X00094F0EBC5442D9A65E';
    //up, 關頻道


    Memo2.Lines.Add('-----------------------');
    //down, reset PIN Code
    sL_Data := '00123456  9    陽明山  E1    2DC3  1234567890        2003043088992398760054E28    Howard   0L    9';
    parseClientNdsData(sL_Data);
    //產生之完整 command : '0003S00170001010100200304281034460001SIN    2DC3q400L8899000000FFFF0000208AC9EE0A9B8208';
    //up,  reset PIN Code



    Memo2.Lines.Add('-----------------------');
    //down, 停機
    sL_Data := '00123456  9    陽明山  A3    2DC3  1234567890        2003043088992398760054E28    Howard   0L   10';
    parseClientNdsData(sL_Data);
    //產生之完整 command : '0003S00170001010100200304281034460001SIN    2DC3q400L8899000000FFFF0000208AC9EE0A9B8208';
    //up,  停機


    Memo2.Lines.Add('-----------------------');
    //down, 復機
    sL_Data := '00123456  9    陽明山  A4    2DC3  1234567890        2003043088992398760054E28    Howard   0L   11';
    parseClientNdsData(sL_Data);
    //產生之完整 command : '0003S00170001010100200304281034460001SIN    2DC3q400L8899000000FFFF0000208AC9EE0A9B8208';
    //up,  復機



    Memo2.Lines.Add('-----------------------');
    //down, Store Region Key
    sL_Data := '00123456  9    陽明山  E2    2DC3  1234567890        2003043088992398760054E28    Howard   0L   12';
    parseClientNdsData(sL_Data);
    //產生之完整 command : '0003S000B0001010100200305051007030001ZIH98760054D2A074D1001484B474CBB00BF29';
    //up,  Store Region Key


    Memo2.Lines.Add('-----------------------');
    //down, Create Service/PPV/OPPV
    sL_Data := '00123456  9    陽明山  B300000004  1234567890        2003043088992398760054E28    Howard  23Super Ball~20030601~O~PL  234';
    parseClientNdsData(sL_Data);
    //產生之完整 command : '0003M00530001010100200305201727540001RINFFFFC0ASuper Ball20030601OPF3B9BDF12069B3E9';
    //up,  Create Service/PPV/OPPV
}

end;

procedure TfrmMain.parseCommandData(I_CmdInfoObject: TCmdInfoObject);
var
    ii, nL_NotesLen, nL_MsgLength : Integer;
    sL_FromID, sL_NotesLen, sL_Notes : String;
    sL_CompCode, sL_SeqNo, sL_SignatureKey : String;
    sL_DetailCommandIDs, sL_ActionBody, sL_ServiceID : String;

    sL_Operator, sL_CmdEnvelop, sL_CmdHeader : String;
    sL_IccNo,sL_HighLevelCmdId, sL_SubBeginDate, sL_SubEndDate : String;

    sL_CompName,sL_RegionKey, sL_Reportbackavailability, sL_SubscriberID,sL_PinCode: String;
    L_CmdInfoObject : TCmdInfoObject;
    sG_FullCmd, sL_PopolcationID,sL_ReportbackDate : String;
    L_StrList : TStringList;

    sL_Signature, sL_PinControl, sL_ActionType, sL_ActionPriority, sL_PriorityReassignment : String;

    sL_CompId, sL_DispatcherSequenceNo : String;
begin
    sL_SeqNo := I_CmdInfoObject.SeqNo;
    sL_CompCode := I_CmdInfoObject.CompCode;
    sL_CompName := I_CmdInfoObject.CompName;
    sL_HighLevelCmdId := I_CmdInfoObject.HighLLevelCmdID;
    sL_SubscriberID := I_CmdInfoObject.SubscriberID;
    sL_IccNo := I_CmdInfoObject.IccNo;
    sL_SubBeginDate := I_CmdInfoObject.SubscriptionBeginDate;
    sL_SubEndDate := I_CmdInfoObject.SubscriptionEndDate;
    sL_PinCode := I_CmdInfoObject.PinCode;
    sL_PopolcationID := I_CmdInfoObject.PopulationID;
    sL_RegionKey := I_CmdInfoObject.RegionKey;
    sL_Reportbackavailability := I_CmdInfoObject.ReportbackAvailability;
    sL_ReportbackDate := I_CmdInfoObject.ReportbackDate;
    sL_Notes := I_CmdInfoObject.NOTES;
    sL_Operator := I_CmdInfoObject.Operator;
    sL_PinControl := I_CmdInfoObject.PinControl;
    sL_ServiceID := I_CmdInfoObject.ServiceID;

    sG_HighLevelCmdId := sL_HighLevelCmdId;
    getActionTypeAndPriority(sL_HighLevelCmdId, sL_ActionType, sL_ActionPriority, sL_PriorityReassignment, sL_FromID);
    sL_SignatureKey := getSignatureKey(sL_FromID);
    sL_FromID := IntToHex(StrToInt(sL_FromID),4);

    sG_PinCode := TUstr.AddString(sL_PinCode,'0', false, PIN_CODE_LENGTH);
    sG_PinControl := sL_PinControl;

    sG_CardID := TUstr.AddString(sL_IccNo,'0', false, ICC_NO_LENGTH);
    sG_RegionKey := TUstr.AddString(sL_RegionKey,' ', false, REGION_KEY_LENGTH);

    sG_ReportbackAvailability := sL_Reportbackavailability;

    sG_ReportbackDate := IntToHex(StrToInt(sL_ReportbackDate), 2);

    sG_PopulationID := IntToHex(StrToIntDef(sL_PopolcationID,0),4);

    sL_DetailCommandIDs := getDetailCommandIDs(sL_HighLevelCmdId);
    if sL_DetailCommandIDs='' then Exit;

    sL_CmdEnvelop := sG_Version + sG_SecurityType ;


    sL_CmdHeader := sL_FromID + IntToHex(nG_ConnectionID,2) + sG_ToID + TUstr.replaceStr(TUdateTime.GetPureDateTimeStr(now,'',''),' ','') + IntToHex(nG_SequenceID,4) + sL_ActionType + sL_ActionPriority + sL_PriorityReassignment;

    if sL_ActionType='S' then
      sL_CmdHeader := sL_CmdHeader + TUstr.AddString(sL_SubscriberID,'0',false,SUBSCRIBER_ID_LENGTH)
    else if sL_ActionType='Z' then
      sL_CmdHeader := sL_CmdHeader + sG_RegionKey
    else if sL_ActionType='R' then
      sL_CmdHeader := sL_CmdHeader + IntToHex(StrToIntDef(sL_ServiceID,0),4);


    L_StrList := TUStr.ParseStrings(sL_DetailCommandIDs,',');

    for ii:=0 to L_StrList.Count -1 do
    begin
      sL_ActionBody := sL_ActionBody + getActionBody(sL_ActionType, L_StrList.Strings[ii], sL_Notes);
      {
      if (sL_HighLevelCmdId='A0') then
      begin
        sL_ActionBody := sL_ActionBody + getActionBody(sL_ActionType, L_StrList.Strings[ii], sL_Notes);
      end;
      }
    end;

    nL_MsgLength := length(sL_CmdEnvelop + sL_CmdHeader + sL_ActionBody) + 4 + 16; // 4=> Length 本身的長度, 16=> Signature 的長度

    sL_CmdEnvelop := sL_CmdEnvelop + IntToHex(nL_MsgLength,4);


    sL_Signature := getSignature(sL_CmdEnvelop + sL_CmdHeader + sL_ActionBody , sL_SignatureKey, sG_SecurityType);


    sG_FullCmd := sL_CmdEnvelop + sL_CmdHeader + sL_ActionBody + sL_Signature;

    recordSeqAndTransMapping(sL_CompCode, sL_SeqNo, IntToStr(nG_SequenceID), sL_HighLevelCmdId);
    preSendCmd(sG_FullCmd);


    if bG_AutoResponse then
    begin
      getDispatcherSeqNo(IntToStr(nG_SequenceID), sL_CompId, sL_HighLevelCmdID, sL_DispatcherSequenceNo);
      preSendToDispatcher(sL_CompId, sL_HighLevelCmdID, sL_DispatcherSequenceNo, 'C', '', '', '00000005', '1'); // 測試回應給 dispatcher...
    end;
    Inc(nG_SequenceID);
end;


procedure TfrmMain.getActionTypeAndPriority(sI_HighLevelCommand: String;
  var sI_ActionType, sI_ActionPriority, sI_PriorityReassignment, sI_FromID: String);
var
    sL_TmpData : String;
    L_StrList : TStringList;
    ii : Integer;
begin
    for ii:=0 to G_ActionTypeAndPriorityStrList.Count -1 do
    begin
      if sI_HighLevelCommand = Copy(G_ActionTypeAndPriorityStrList.Strings[ii],1,2) then
      begin
        sL_TmpData := Copy(G_ActionTypeAndPriorityStrList.Strings[ii],4, length(G_ActionTypeAndPriorityStrList.Strings[ii])-3);
        L_StrList := TUstr.ParseStrings(sL_TmpData,';');
        if L_StrList.Count=4 then
        begin
          sI_ActionType := L_StrList.Strings[0];
          sI_ActionPriority := L_StrList.Strings[1];
          sI_PriorityReassignment := L_StrList.Strings[2];
          sI_FromID := L_StrList.Strings[3]; 
        end
        else
        begin
          MessageDlg('ini 檔之 Action Type 格式設定錯誤.程式即將結束!', mtError, [mbOk],0);
          Application.Terminate;
        end;  
        break;
      end;
    end;


end;

function TfrmMain.getActionBody(sI_ActionType,
  sI_LowLevelCmdID, sI_Notes: String): String;
var
    nL_ParentalRating,nL_EventSpendingLimit,    nL_CumulativeBalance : Integer;

    sL_UserNumber, sL_Status : String;
    nL_TmpParameterValue, ii, jj : Integer;
    sL_Result : String;
    L_StrList1, L_StrList2 : TStringList;
    nL_ServiceID, nL_LengthOfServiceName : Integer;
    sL_ExpirationDate, sL_TapingAuthorization : String;

    sL_PayFreeFlag, sL_ServiceName, sL_TypeOfService : String;
begin
    if UpperCase(sI_ActionType)='S' then
    begin
      if sI_LowLevelCmdID='C' then //create subscriber
      begin
        sL_Result := 'C' + sG_CardID + sG_RegionKey + sG_ReportbackAvailability + sG_ReportbackDate + sG_BackwardTolerance + sG_ForwardTolerance + sG_Currency + sG_PopulationID;
      end
      else if sI_LowLevelCmdID='D' then //delete subscriber
      begin
        sL_Result := 'D';
      end
      else if sI_LowLevelCmdID='A' then //authorize service
      begin
        L_StrList1 := TUstr.ParseStrings(sI_Notes,';');
        if L_StrList1<> nil then
        begin
          for ii:=0 to L_StrList1.Count -1 do
          begin
            L_StrList2 := TUstr.ParseStrings(L_StrList1.Strings[ii],'~');
            if L_StrList2.Count = 3 then
            begin
              nL_ServiceID := StrToInt(L_StrList2.Strings[0]);
              sL_ExpirationDate := L_StrList2.Strings[1];
              sL_TapingAuthorization := L_StrList2.Strings[2];
              sL_Result := sL_Result + 'A' + IntToHex(nL_ServiceID,4) + sL_ExpirationDate + sL_TapingAuthorization;
            end;

          end;
        end;
      end
      else if sI_LowLevelCmdID='X' then //cancel service
      begin
        L_StrList1 := TUstr.ParseStrings(sI_Notes,',');
        for ii:=0 to L_StrList1.Count -1 do
        begin
          nL_ServiceID := StrToInt(L_StrList1.Strings[ii]);
          sL_Result := sL_Result + 'X' + IntToHex(nL_ServiceID,4);

        end;

      end
      else if sI_LowLevelCmdID='q' then //set user parameter
      begin
        nL_TmpParameterValue := 0;
        sL_UserNumber := '0';
        nL_ParentalRating := 0;
        nL_EventSpendingLimit := 65535; // 65535 => no limit
        nL_CumulativeBalance := 0; // 0 => reset the balance
        if sG_PinCode<>'' then
          nL_TmpParameterValue := 64;
        sL_Result := sL_Result + 'q' + IntToHex(nL_TmpParameterValue,2) + sL_UserNumber + sG_PinControl + sG_PinCode + IntToHex(nL_ParentalRating,6) + IntToHex(nL_EventSpendingLimit,4) + IntToHex(nL_CumulativeBalance,4);

      end
      else if sI_LowLevelCmdID='V' then // resend all packets
      begin
        sL_Result := sL_Result + 'V';
      end
      else if sI_LowLevelCmdID='j' then // Activate/Deactivate Subscriber
      begin
        if sG_HighLevelCmdId='A3' then //Deactivate Subscriber
        begin
          sL_Status := '0';
          sL_Result := sL_Result + 'j' + sL_Status;
        end
        else if sG_HighLevelCmdId='A4' then //Activate Subscriber
        begin
          sL_Status := '1';
          sL_Result := sL_Result + 'j' + sL_Status ;  
        end;

      end;
    end
    else if UpperCase(sI_ActionType)='Z' then
    begin
      if sI_LowLevelCmdID='D' then //srore region key
      begin
        sL_Result := sL_Result + 'D' + sG_CountryNumber + IntToHex(nG_TimeZoneOffset1*2,2) + IntToHex(nG_TimeZoneOffset2,2);
      end;
    end
    else if UpperCase(sI_ActionType)='R' then
    begin
      if sI_LowLevelCmdID='C' then //Create Service/PPV/OPPV
      begin
        // sI_Notes 範例: Super Ball~20030601~O~P
        L_StrList1 := TUstr.ParseStrings(sI_Notes,'~');
        if L_StrList1.Count = 4 then
        begin
          sL_ServiceName := L_StrList1.Strings[0];
          nL_LengthOfServiceName := length(sL_ServiceName);
          sL_ExpirationDate := L_StrList1.Strings[1];
          sL_TypeOfService := L_StrList1.Strings[2];
          sL_PayFreeFlag := L_StrList1.Strings[3];
          sL_Result := sL_Result + 'C' + IntToHex(nL_LengthOfServiceName,2) + sL_ServiceName + sL_ExpirationDate + sL_TypeOfService + sL_PayFreeFlag;
        end;
      end;
    end;
    result := sL_Result;
end;

function TfrmMain.getSignature(sI_Data, sI_Key, sI_SecurityType: String): String;
var
    sL_Result : String;
begin
    sL_Result := '';
    myMD5_RealMD5_8(sI_SecurityType, sI_Data, IntToStr(length(sI_Data)), sI_Key, IntToStr(length(sI_Key)), sL_Result);
    result := sL_Result;

end;

procedure TfrmMain.preSendCmd(sI_FullCmd: String);
var
    sL_RealData : String;
    L_WStream: TWinSocketStream;
    L_Socket : TClientSocket;
    L_CmdArray : array[0..1024] of char;
    L_ResponseAry : array[0..1024] of Char;
    nL_Words, nL_Len,ii, nL_WaitForDataTimeOut : Integer;
begin
    Memo2.Lines.Add(sI_FullCmd);

    bG_SendComm := true;
    //down, 傳送command
    try
    begin
      Memo4.Lines.Add('送出的指令: ' + sI_FullCmd);

      nL_Len := length(sI_FullCmd);//ok..
      FillChar(L_CmdArray, high(L_CmdArray), 0); { initialize the buffer }

      for ii:=1 to length(sI_FullCmd) do
      begin
        L_CmdArray[ii-1] := sI_FullCmd[ii];
      end;


      L_Socket := G_SendCmdClntSocket;

      L_WStream := G_SendCmdWStream;


      try
        if (L_WStream<>nil) and (L_Socket.Socket.Connected)then
        begin
  //        nL_Words := G_WinSocketStreamForSendCmd.Write(L_DeviceData, nL_Len+2); //0317
          nL_WaitForDataTimeOut := SOCKET_WAIT_FOR_DATA_TIMEOUT*1000;
          nL_Words := write(@L_CmdArray, nL_Len, L_WStream); //send_data..
  //        nL_Words := write(@L_DeviceData, nL_Len+2, L_WStream); //send_data..   //0317
  //*******************

          if L_WStream.WaitForData(nL_WaitForDataTimeOut) then
          begin
            if (L_Socket.Socket.Connected) then
            begin
              nL_Words := L_WStream.Read(L_ResponseAry, sizeof(L_ResponseAry)); //這種寫法會使得傳送速度變得很快, 但接收速度變慢
      //      nL_Words := recv(frmMain.G_SendCmdClntSocket.Socket.SocketHandle , G_DeviceData, sizeof(G_DeviceData),0);

              if nL_Words>0 then
              begin
                for ii:=0 to nL_Words-1 do
                begin
                  sL_RealData := sL_RealData + L_ResponseAry[ii];
                end;
                Memo4.Lines.Add('收到的指令回應: ' + sL_RealData);
                handleCmdResponse(sL_RealData);

                {
                if Assigned(G_ResponseMsgStrList) then
                  G_ResponseMsgStrList.Add(sL_RealData);
                }                  
Memo4.Lines.Add('===================================================');
              end;
            end
            else
            begin
              //down, 重新連接到 EMMG..
              reBuildNdsConn;
              //up, 重新連接到 EMMG..
            end;
          end;
        end;
//*******************
      except
        bG_buildSendCmdConnection := false;
        try
          frmMain.reBuildNdsConn;

        except

          frmMain.showLogOnMemo(SEND_CMD_ERROR_FLAG,'傳送指令' + sI_FullCmd + ' 給 Nagra.程式即將結束,然後重新啟動. '); //執行完此行程式後, 程式會結束掉

        end;
        {
        buildSendCmdConnection(frmMain.G_SendCmdClntSocket, frmMain.G_SendCmdWStream);
        if not bG_buildSendCmdConnection then
        begin
          sL_LowLevelCmdID := Copy(sG_FullCommand,61,4);
          frmMain.showLogOnMemo(SEND_CMD_ERROR_FLAG,'傳送指令' + sL_LowLevelCmdID + ' 給 Nagra.程式即將結束,然後重新啟動. '); //執行完此行程式後, 程式會結束掉
        end;
        }
      end;


    end
    except


      bG_SendComm := false;






    end;
    //up, 傳送command
end;


procedure TfrmMain.N7Click(Sender: TObject);
begin
    frmAbout := TfrmAbout.Create(Application);
    frmAbout.ShowModal;
    frmAbout.Free;

end;

procedure TfrmMain.recordSeqAndTransMapping(sI_CompCode, sI_SeqNo,
  sI_SequenceID, sI_HighLevelCmdID: String);
var
    L_Obj : TSequenceIdMappingInfoObj;

begin
    if G_SeqTransMappingStrList.Count >= CA_UI_MAX_COUNT then
      G_SeqTransMappingStrList.Clear;

    L_Obj := TSequenceIdMappingInfoObj.Create;
    L_Obj.CompCode := sI_CompCode;
    L_Obj.HighLevelCmdId := sI_HighLevelCmdID;
    L_Obj.DispatcherSequenceNo := sI_SeqNo;
    G_SeqTransMappingStrList.AddObject(sI_SequenceID, L_Obj);



end;

procedure TfrmMain.Dispatcher21Click(Sender: TObject);
begin
    frmDispatcherInfo := TfrmDispatcherInfo.Create(Application);
    frmDispatcherInfo.ShowModal;
    frmDispatcherInfo.free;

end;

procedure TfrmMain.myTcpClient1Disconnect(Sender: TObject);
begin
//    showmessage('disConnect..');
end;

procedure TfrmMain.myTcpClient1Error(Sender: TObject;
  SocketError: Integer);
begin
    showmessage('socket error..');
end;

procedure TfrmMain.dispatcherDisconnectSettingCmd(sI_RemoteIP: String);
var
    ii  : Integer;
begin
    for ii:=G_RemoteInfoStrList.Count -1 downto 0 do
    begin
      if (G_RemoteInfoStrList.Objects[ii] As TCompAndIpInfoObj).IP = sI_RemoteIP then
        G_RemoteInfoStrList.Delete(ii);
    end;
    ii := G_ClientIpStrList.IndexOf(sI_RemoteIP);
    if ii>=0 then
    begin
      G_ClientIpStrList.Delete(ii);
    end;

    createTcpClient;
end;

procedure TfrmMain.createTcpClient;
var
    ii  : Integer;
begin
    //down, 建立 TcpClient 的 array...
    for ii:=0 to G_ClientIpStrList.Count-1 do
    begin
      if Assigned(G_AryOfTcpClient[ii]) then
      begin
        G_AryOfTcpClient[ii].Free;
        G_AryOfTcpClient[ii] := nil;
      end;

      G_AryOfTcpClient[ii] := TTcpClient.Create(Application);
      G_AryOfTcpClient[ii].OnDisconnect := myTcpClient1Disconnect;
      G_AryOfTcpClient[ii].OnError := myTcpClient1Error;

      G_AryOfTcpClient[ii].RemoteHost := G_ClientIpStrList.Strings[ii];
      G_AryOfTcpClient[ii].RemotePort := sG_DispatcherListenPort;
      G_AryOfTcpClient[ii].Connect;
    end;
    //up, 建立 TcpClient 的 array...

end;

procedure TfrmMain.sendDisconnectInfoToDispatcher;
var
    ii : Integer;
begin
    for ii:=0 to High(G_AryOfTcpClient) do
    begin
      G_AryOfTcpClient[ii].Sendln(DISCONNECT_CMD);
    end;
end;

procedure TfrmMain.BitBtn1Click(Sender: TObject);
begin
    Memo4.Lines.Add(getSignature('0001M0035010001000119940206125158ER01','12345678','M'));

end;

function TfrmMain.getCardIdChecksum(fI_CardID:double): Integer;
var
    i, ckdig, digit : Integer;
    odd_sum, even_sum  : Integer;
    nL_CardID : Integer;
    fL_CardID : double;
begin
    //down, Card ID 的 checksum 產生程式...
    odd_sum := 0;
    even_sum := 0;
    fL_CardID := fI_CardID;
    for i:=1 to 11 do
    begin
      nL_CardID := StrToInt(FloatToStr(fL_CardID));
      digit := StrToInt(FloatToStr(nL_CardID mod 10));
      fL_CardID := Trunc(fL_CardID/10);
      if ((i mod 2) = 1) then
      begin
        digit :=  digit*2;
        if (digit >= 10) then
          digit := digit - 9;
          odd_sum := odd_sum + digit;
        end
        else
          even_sum := even_sum + digit;
    end;
    ckdig := (10 - ((even_sum + odd_sum) mod 10)) mod 10;
    result := ckdig;
    //down, Card ID 的 checksum 產生程式...
end;

procedure TfrmMain.BitBtn2Click(Sender: TObject);
var
    nL_CheckSum : Integer;
begin
    nL_CheckSum := getCardIdChecksum(StrToInt(Edit1.Text));
    showmessage(inttostr(nL_CheckSum));
end;

function TfrmMain.getSignatureKey(sI_FromID: String): String;
var
    sL_Result : String;
    nL_Ndx : Integer;
begin
    nL_Ndx := G_SignatureKey.IndexOf(sI_FromID);
    if nL_Ndx>=0 then
      sL_Result := (G_SignatureKey.Objects[nL_Ndx] as TSignatureKeyInfoObj).sSignatureKey
    else
      sL_Result := '';
    result := sL_Result;
end;

procedure TfrmMain.transIntoSignatureKeyStrList(
  sI_AllSignatureKey: String);
var
    L_Obj : TSignatureKeyInfoObj;
    sL_FromID, sL_SignatureKey : String;
    L_TmpStrList1, L_TmpStrList2 : TStringList;
    ii, jj : Integer;
begin
{
    sI_AllSignatureKey 的可能值 : 0,00000000~1,12345678
    範例表示共有三組 signature key 的設定
    from id 是 0 的 signature key 是 00000000
    from id 是 1 的 signature key 是 12345678
    from id 是 2 的 signature key 是 12345678
}
    if sI_AllSignatureKey='' then Exit;
    L_TmpStrList1 := TUstr.ParseStrings(sI_AllSignatureKey, '~');
    for ii:=0 to L_TmpStrList1.Count -1 do
    begin
      L_TmpStrList2 := TUstr.ParseStrings(L_TmpStrList1.Strings[ii], ',');
      for jj :=0 to L_TmpStrList2.Count -1 do
      begin
        sL_FromID := L_TmpStrList2.Strings[0];
        sL_SignatureKey := L_TmpStrList2.Strings[1];
        if G_SignatureKey.IndexOf(sL_FromID)=-1 then
        begin
          L_Obj := TSignatureKeyInfoObj.Create;
          L_Obj.sSignatureKey := sL_SignatureKey;
          G_SignatureKey.AddObject(sL_FromID, L_Obj);
        end;
      end;
    end;

end;

procedure TfrmMain.Button2Click(Sender: TObject);
begin
    showmessage(getSignatureKey(Edit2.Text));
end;

procedure TfrmMain.BitBtn3Click(Sender: TObject);
begin
    Memo4.Lines.Add(getSignature('0001M0035010001000119940206125158ER01','12345678','S'));
end;


procedure TfrmMain.BitBtn4Click(Sender: TObject);
begin
    handleCmdResponse('0003S00A60100020001200306050651280005S000000051I00000001006600000000000000000000000000000000000000987600541C0000000000000000000000000000E051938000003O762198EB8B01B0DF');
end;

procedure TfrmMain.getReturnDataForDispatcher(sI_Action, sI_SequenceID: String);
var
    sL_ErrorCode, sL_ErrorMsg:String;
    nL_CmdFlagPos, nL_ActionLen, nL_PersonalRegionLen : Integer;
    sL_SubscriberID, sL_PersonalRegionLen, sL_PersonalRegion : String;
    nL_CompNdx, nL_Seed : Integer;
    sL_ResponseDataStatusFlag, sL_CompId, sL_HighLevelCmdID, sL_DispatcherSequenceNo : String;
    sL_ReportbackAvailability, sL_ReportbackDay, sL_ReportbackTime, sL_Currency : String;
    sL_StatusCode, sL_Response, sL_CardID, sL_NewCardID, sL_RegionKey, sL_Date : String;
begin
    sL_CompId := '9';
    sL_HighLevelCmdID := 'A0';
    sL_DispatcherSequenceNo := '123';
    {
    getDispatcherSeqNo(sI_SequenceID, sL_CompId, sL_HighLevelCmdID, sL_DispatcherSequenceNo);
    if sL_CompId='' then
    begin
      memErrorLog.Lines.Add('無法將指令處理結果回應給 Dispatcher.因為無法找出指令所屬的系統台.EMMG指令序號:' + sI_SequenceID + ' --'+ DatetimeToStr(now));
      Exit;
    end;


    nL_CompNdx := G_RemoteInfoStrList.IndexOf(sL_CompId);
    if nL_CompNdx=-1 then
    begin
      memErrorLog.Lines.Add('無法將指令處理結果回應給 Dispatcher.因為無法找出系統台的相關資訊.系統台代號:' + sL_CompId + ' --'+ DatetimeToStr(now));
      Exit;
    end;
    }
    sL_ResponseDataStatusFlag := 'C';
    sL_ErrorCode := '';
    sL_ErrorMsg := '';
    nL_CmdFlagPos := 1;
    nL_ActionLen := length(sI_Action);

    while nL_ActionLen >= nL_CmdFlagPos do
    begin
      if sI_Action[nL_CmdFlagPos]='O' then
      begin
        Inc(nL_CmdFlagPos);
        continue;
      end
      else if sI_Action[nL_CmdFlagPos]='E' then
      begin
        sL_ResponseDataStatusFlag := 'E';        
        sL_ErrorCode := Copy(sI_Action,nL_CmdFlagPos, ERROR_CODE_LENGTH);
        sL_ErrorMsg := getErrorMsg(sL_ErrorCode);
        nL_CmdFlagPos := nL_CmdFlagPos + ERROR_CODE_LENGTH;
//        break;
        continue;
      end
      else if sI_Action[nL_CmdFlagPos]='S' then
      begin
        nL_Seed := nL_CmdFlagPos+1;

        sL_SubscriberID := Copy(sI_Action, nL_Seed, SUBSCRIBER_ID_LENGTH);
        nL_Seed := nL_Seed + SUBSCRIBER_ID_LENGTH;

        sL_StatusCode :=  Copy(sI_Action, nL_Seed, 1);
        {
        if (sL_StatusCode = '1') => request accepted
        if (sL_StatusCode = '2') => reserved
        if (sL_StatusCode = '3') => subscriber already exists
        }
        nL_Seed := nL_Seed + 1;

        nL_CmdFlagPos := nL_CmdFlagPos + RESPONSE_636_LENGTH;
        continue;
      end
      else if sI_Action[nL_CmdFlagPos]='I' then
      begin
        nL_Seed := nL_CmdFlagPos+1;

        sL_CardID :=  Copy(sI_Action, nL_Seed, ICC_NO_LENGTH);
        nL_Seed := nL_Seed + ICC_NO_LENGTH;

        sL_NewCardID := Copy(sI_Action, nL_Seed, ICC_NO_LENGTH);
        nL_Seed := nL_Seed + ICC_NO_LENGTH;
        nL_Seed := nL_Seed + 18; //18是兩個保留欄位的長度

        sL_Date := Copy(sI_Action, nL_Seed, DATE_LENGTH);
        nL_Seed := nL_Seed + DATE_LENGTH;

        sL_RegionKey := Copy(sI_Action, nL_Seed, REGION_KEY_LENGTH);
        nL_Seed := nL_Seed + REGION_KEY_LENGTH;

        sL_PersonalRegionLen := Copy(sI_Action, nL_Seed, PERSONAL_REGION_LENGTH_LENGTH);
        nL_PersonalRegionLen := TUstr.hexToInt(sL_PersonalRegionLen);
        nL_Seed := nL_Seed + PERSONAL_REGION_LENGTH_LENGTH;

        sL_PersonalRegion :=  Copy(sI_Action, nL_Seed, nL_PersonalRegionLen);
        nL_Seed := nL_Seed + nL_PersonalRegionLen;

        sL_ReportbackAvailability := Copy(sI_Action, nL_Seed, REPORTBACK_AVAILABILITY_LENGTH);
        nL_Seed := nL_Seed + REPORTBACK_AVAILABILITY_LENGTH;

        sL_ReportbackDay  := Copy(sI_Action, nL_Seed, REPORTBACK_DATE_LENGTH);
        nL_Seed := nL_Seed + REPORTBACK_DATE_LENGTH;


        sL_ReportbackTime  := Copy(sI_Action, nL_Seed, REPORTBACK_TIME_LENGTH);
        nL_Seed := nL_Seed + REPORTBACK_TIME_LENGTH;

        sL_Currency  := Copy(sI_Action, nL_Seed, CURRENCY_LENGTH);
        nL_Seed := nL_Seed + CURRENCY_LENGTH;



        nL_CmdFlagPos := nL_CmdFlagPos + RESPONSE_633_LENGTH;
        continue;

      end;
    end;

    //down, 組合出適當的訊息字串,然後傳回給 dispatcher..

    preSendToDispatcher(sL_CompID, sL_HighLevelCmdID, sL_DispatcherSequenceNo, sL_ResponseDataStatusFlag,sL_ErrorCode, sL_ErrorMsg, sL_SubscriberID, sL_StatusCode);

    //up, 組合出適當的訊息字串,然後傳回給 dispatcher..
end;

procedure TfrmMain.Button3Click(Sender: TObject);
var
    sL_Action, sL_SequenceID : String;
begin
    sL_Action := 'S000000051I00000001006600000000000000000000000000000000000000987600541C0000000000000000000000000000E051938000003O';
    sL_SequenceID := '123';  //隨意填的值
    getReturnDataForDispatcher(sL_Action, sL_SequenceID);
end;

procedure TfrmMain.loadErrorMsgContent(I_StrList: TStringList);
var
    ii : Integer;
    sL_FileName : String;
    L_TmpStrList, L_TmpStrList1 : TStringList;
    sL_ExeFileName, sL_ExePath : STring;
    sL_ErrorCode, sL_ErrorMsg : String;

    L_Obj : TErrorMsgObj;
begin

    sL_ExeFileName := Application.ExeName;
    sL_ExePath := ExtractFileDir(sL_ExeFileName);

    sL_FileName := sL_ExePath + '\' + ERROR_MSG_FILE_NAME;

    if FileExists(sL_FileName) then
    begin

      L_TmpStrList := TStringList.Create;
      L_TmpStrList.LoadFromFile(sL_FileName);
      for ii:=0 to L_TmpStrList.Count -1 do
      begin
        if L_TmpStrList.Strings[ii] <>'' then
        begin
          L_TmpStrList1 := TUstr.ParseStrings(L_TmpStrList.Strings[ii],'=');
          if L_TmpStrList1.Count=2 then
          begin
            sL_ErrorCode := L_TmpStrList1.Strings[0];
            sL_ErrorMsg := L_TmpStrList1.Strings[1];
            L_Obj := TErrorMsgObj.Create;
            L_Obj.ErrorMsg := sL_ErrorMsg;

            I_StrList.AddObject(sL_ErrorCode, L_Obj);
          end;
          if Assigned(L_TmpStrList1) then
            L_TmpStrList1.Free;

        end;
      end;

      L_TmpStrList.Free;
    end
    else
    begin
      MessageDlg('找不到錯誤訊息檔(' + sL_FileName + ')程式即將結束.', mtError, [mbOK],0);
      releaseResource;
      Application.Terminate;
    end;
end;

function TfrmMain.getErrorMsg(sI_ErrorCode: String): String;
var
    sL_Result : String;
    nL_Ndx : Integer;
begin
    sL_Result := '';
    nL_Ndx := G_ErrorMsgStrList.IndexOf(sI_ErrorCode);
    if nL_Ndx>=0 then
      sL_Result := (G_ErrorMsgStrList.Objects[nL_Ndx] as TErrorMsgObj).ErrorMsg;

    result := sL_Result;  
end;

end.




