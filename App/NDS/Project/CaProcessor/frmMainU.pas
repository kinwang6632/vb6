unit frmMainU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ScktComp, StdCtrls, Sockets, inifiles,
  comctrls, syncobjs,  winsock, DB, DBClient, Menus, ExtCtrls, ImgList, TListernerU, ConstU,
  Buttons, IdGlobal ;

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
    tabTesting: TTabSheet;
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
    Button4: TButton;
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
    procedure Button4Click(Sender: TObject);
  private
    { Private declarations }
    sG_AllSignatureKey, sG_CountryNumber : String;
    nG_TimeZoneOffset1, nG_TimeZoneOffset2,nG_DefaultPopulationID : Integer;
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
    sG_SrcTransNum : String;// �|�Q��J Nagra  �ǨӪ� transaction num => return path �ɷ|�Ψ�
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
    bG_LoginOk, bG_AutoResponse, bG_AutoInvoke, bG_NoEmmg ,bG_LabStatus : boolean;
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
    procedure getDispatcherSeqNo(bI_SequenceIdIsHax: Boolean; sI_SequenceID:String; var sI_CompCode, sI_HighLevelCmdID, sI_DispatcherSequenceNo:String);

    procedure clearStatusItem(nI_ItemIndex:Integer); overload;
    procedure createLivItem;
    procedure recordSeqAndTransMapping(sI_CompCode, sI_SeqNo, sI_SequenceID, sI_HighLevelCmdID:String);overload;
    procedure showUI1(nI_CompCode:Integer; sI_CompName, sI_SeqNo, sI_CommandID, sI_Operator, sI_TransNum : String);
    procedure showUI2(sI_TransNum, sI_ErrorCode, sI_ErrorMsg:String; bI_HasError:boolean);
    procedure sendToDispatcher(nI_Type:Integer; sI_CompID, sI_Info, sI_HighLevelCmdID, sI_Operator, sI_SrcTransNum, sI_ResponseCmdID:String);
    procedure preSendToDispatcher(sI_CompID, sI_HighLevelCmdID, sI_DispatcherSeqNo, sI_ResponseDataStatus, sI_ErrorCode, sI_ErrorMsg, sI_AttachedResponseData:String);
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
    procedure closeSocket;
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
    procedure showLogOnMemo(nI_ActiveOperation:Integer; sI_Detail:String); //0328
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
      result := '-1: Ū���Ѽ���<'+sL_IniFileName +'>����';
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
      result := '-1: �г]�w�Ѽ���<'+sL_IniFileName +'>�� PROCESSOR_IP ';
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
    bG_NoEmmg := false;
    bG_AutoResponse := false;
    bG_AutoInvoke := false;
    sG_NewIniFileName := '';
    tabTesting.TabVisible := false;
    bG_Debug := false;
    bG_LabStatus := false;

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
            //��ܫ����n�έ��@�� ini file...
            sG_NewIniFileName := ParamStr(ii);
          end;
          if UpperCase(ParamStr(ii)) = 'AR' then
          begin
            bG_AutoResponse := true;
          end;
          if UpperCase(ParamStr(ii)) = 'TESTING' then
          begin
            bG_Debug := true;
            tabTesting.TabVisible := true;
          end;
          if UpperCase(ParamStr(ii)) = 'NOEMMG' then
          begin
            bG_NoEmmg := true;
          end;
          if UpperCase(ParamStr(ii)) = 'LAB' then
          begin
            bG_LabStatus := true;
          end;

        end;
    end;
    //�t�εn�J�e��



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
        //down, �q ini ��Ū�� information
        sL_Result := getSysIniSettiing;
        if sL_Result<>'' then
        begin
          if not bG_AutoInvoke then
            MessageDlg(sL_Result,mtError,[mbOK],0);
          Application.Terminate;
          exit;
        end;
        //up, �q ini ��Ū�� information
        caption := caption + ' - ('+sG_LoginName+')' + '--' + sG_NewIniFileName;

      except
      end;
      try
        if not bG_NoEmmg then
          buildNdsConnectionForSendCmd(G_SendCmdClntSocket,G_SendCmdWStream);
      except
        frmMain.showLogOnMemo(REBUILD_NDS_CONNECTION_FAILED,' frmMain.reBuildNdsConn .�{���Y�N����,�M�᭫�s�Ұ�. '); //���槹����{����, �{���|������
        Exit;
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
      //�ǤJ���r��p�U(�]�t�Ҧ��ť�)
          �d�ҡG'00123456  9    �����s  A1000009D4  1234567890        2003043088992398760054E28Howard      11thisisnotesL'
      }
     sL_SeqNo := trim(Copy(sI_ClientNdsData,1,SEQ_NO_LENGTH));
      sL_CompCode := trim(Copy(sI_ClientNdsData,9,COMP_CODE_LENGTH));
      sL_CompName := trim(Copy(sI_ClientNdsData,12, COMP_NAME_LENGTH) );
      sL_HighLevelCmdId := trim(Copy(sI_ClientNdsData,32,HIGH_LEVEL_CMD_ID_LENGTH));
      sL_SubscriberID := trim(Copy(sI_ClientNdsData,36,SUBSCRIBER_ID_LENGTH));
      sL_IccNo := trim(Copy(sI_ClientNdsData,44,ICC_NO_LENGTH));
      sL_SubBeginDate := trim(Copy(sI_ClientNdsData,56,SUB_BEGIN_DATE_LENGTH));
      sL_SubEndDate := trim(Copy(sI_ClientNdsData,64,SUB_END_DATE_LENGTH));
      sL_PinCode := trim(Copy(sI_ClientNdsData,72,PIN_CODE_LENGTH));
      sL_PopolcationID :=trim(Copy(sI_ClientNdsData,76,INT_POPULATION_ID_LENGTH));
      sL_RegionKey := trim(Copy(sI_ClientNdsData,81,REGION_KEY_LENGTH));
      sL_Reportbackavailability := trim(Copy(sI_ClientNdsData,89,REPORTBACK_AVAILABILITY_LENGTH));
      sL_ReportbackDate := trim(Copy(sI_ClientNdsData,90,REPORTBACK_DATE_LENGTH));
      sL_Operator := trim(Copy(sI_ClientNdsData,92,OPERATOR_LENGTH));

      sL_NotesLen := trim(Copy(sI_ClientNdsData,112,NOTES_LENGTH));
      nL_NotesLen := StrToIntDef(sL_NotesLen,0);
      sL_Notes := trim(Copy(sI_ClientNdsData,116,nL_NotesLen));

      nL_SeedValue := 116 + nL_NotesLen;

{
      sL_SeqNo := trim(Copy(sI_ClientNdsData,1,SEQ_NO_LENGTH));
      sL_CompCode := trim(Copy(sI_ClientNdsData,9,COMP_CODE_LENGTH));
      sL_CompName := trim(Copy(sI_ClientNdsData,12, COMP_NAME_LENGTH) );
      sL_HighLevelCmdId := trim(Copy(sI_ClientNdsData,22,HIGH_LEVEL_CMD_ID_LENGTH));
      sL_SubscriberID := trim(Copy(sI_ClientNdsData,26,SUBSCRIBER_ID_LENGTH));
      sL_IccNo := trim(Copy(sI_ClientNdsData,34,ICC_NO_LENGTH));
      sL_SubBeginDate := trim(Copy(sI_ClientNdsData,46,SUB_BEGIN_DATE_LENGTH));
      sL_SubEndDate := trim(Copy(sI_ClientNdsData,54,SUB_END_DATE_LENGTH));
      sL_PinCode := trim(Copy(sI_ClientNdsData,62,PIN_CODE_LENGTH));
      sL_PopolcationID :=trim(Copy(sI_ClientNdsData,66,INT_POPULATION_ID_LENGTH));
      sL_RegionKey := trim(Copy(sI_ClientNdsData,71,REGION_KEY_LENGTH));
      sL_Reportbackavailability := trim(Copy(sI_ClientNdsData,79,REPORTBACK_AVAILABILITY_LENGTH));
      sL_ReportbackDate := trim(Copy(sI_ClientNdsData,80,REPORTBACK_DATE_LENGTH));
      sL_Operator := trim(Copy(sI_ClientNdsData,82,OPERATOR_LENGTH));

      sL_NotesLen := trim(Copy(sI_ClientNdsData,92,NOTES_LENGTH));
      nL_NotesLen := StrToIntDef(sL_NotesLen,0);
      sL_Notes := trim(Copy(sI_ClientNdsData,96,nL_NotesLen));

      nL_SeedValue := 96 + nL_NotesLen;
}
      sL_PinControl := trim(Copy(sI_ClientNdsData,nL_SeedValue,PIN_CONTROL_LENGTH));
      nL_SeedValue := nL_SeedValue+PIN_CONTROL_LENGTH;
      sL_ServiceID :=  Trim(Copy(sI_ClientNdsData,nL_SeedValue,INT_SERVICE_ID_LENGTH));
      //abc

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
      if (I_ClientSocket.Active) then
        I_ClientSocket.Close;
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
        Address := '192.168.42.101';
        Port := 2213;
}

{
        Address := '203.193.62.251';
        Port := 12222;
}


        Open;

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
//nI_Value:�Q�i�쪺�ƭȡBnI_Digits:Binary Length (in Bytes)�BbI_Reverse:�O�_�A�ˡB�Ǧ^�Ȭ��@��Record Type
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
//dtmMain.cdsParam.Edit;
//dtmMain.cdsParam.FieldByName('nPopulationID').AsInteger := 1;
//dtmMain.cdsParam.Post;

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

      bG_ResponseLog := FieldByName('bResponseLog').Asboolean;  //���R�W���~....



//      nG_CmdRefreshRate1 := FieldByName('nCmdRefreshRate1').AsInteger; //�y�p�ɬq?��Ū���@�����O��Ʈw
//      nG_CmdRefreshRate2 := FieldByName('nCmdRefreshRate2').AsInteger; //���p�ɬq?��Ū���@�����O��Ʈw

//      nG_CmdResentTimes := FieldByName('nCmdResentTimes').AsInteger;
      bG_ShowUI := FieldByName('bShowUI').Asboolean;
//      sG_TimeDefine := FieldByName('sSHotTime').AsString + '-' + FieldByName('sEHotTime').AsString; // �y�p�ɬq�w�q

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

      nG_DefaultPopulationID := FieldByName('nPopulationID').AsInteger;

      if RecordCount=0 then
        MessageDlg('�t�ΰѼƩ|���]�w! �Х��]�w��,�A����{��.', mtWarning, [mbOK],0);
      Close;
    end;
end;

procedure TfrmMain.terminateApp;
begin
    if MessageDlg('�O�_�T�w�n�����{��?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      closeSocket;    
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
    sL_Result,sL_TempHighLevelCommand : String;
    ii : Integer;
    L_TempStrList : TStringList;
begin
    sL_Result := '';
    for ii:=0 to G_CmdMatrixStrList.Count -1 do
    begin
      //jackal920813�ק�
      L_TempStrList := TUstr.ParseStrings(G_CmdMatrixStrList.Strings[ii],'=');
      if L_TempStrList.Count = 2 then
      begin
        sL_TempHighLevelCommand := L_TempStrList.Strings[0];

        //if sI_HighLevelCommand = Copy(G_CmdMatrixStrList.Strings[ii],1,2) then
        if sI_HighLevelCommand = sL_TempHighLevelCommand then
        begin
          //sL_Result := Copy(G_CmdMatrixStrList.Strings[ii],4, length(G_CmdMatrixStrList.Strings[ii])-3);
          sL_Result := L_TempStrList.Strings[1];;
          break;
        end;
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

      if G_SendCmdClntSocket.Active then
        G_SendCmdClntSocket.Close;
//        sleep(3000);
      buildNdsConnectionForSendCmd(G_SendCmdClntSocket,G_SendCmdWStream);
    except
      frmMain.showLogOnMemo(REBUILD_NDS_CONNECTION_FAILED,' frmMain.reBuildNdsConn .�{���Y�N����,�M�᭫�s�Ұ�. '); //���槹����{����, �{���|������
    end;

end;

procedure TfrmMain.showLogOnMemo(nI_ActiveOperation: Integer;
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
      //down, �N�T���g���ɮפ�
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
      //up, �N�T���g���ɮפ�
    end;
begin
    try
      bL_KillSelf := false;
      bL_ShowMemo := true;

      if nI_ActiveOperation=SEND_CMD_ERROR_FLAG then
      begin
        bL_KillSelf := true;
        sL_DbAlias := '';
        sL_Msg := '['+ datetimetostr(now)+ '] ' +'�L�k�ǰe���O�� Nagra. ' + '  �@�~����:' + sI_Detail ;

        writeMsg;
        releaseResource;
        //down, �I�s�ʱ���, ���ʱ����� CA �s�_��
        {
        L_Intf := CoMonitor.Create;
        L_Intf.RunCa(sG_NewIniFileName);
        L_Intf := nil;
        }
        //up, �I�s�ʱ���, ���ʱ����� CA �s�_��
      end
      else if nI_ActiveOperation=REBUILD_NDS_CONNECTION_FAILED then
      begin
        bL_KillSelf := true;
        sL_DbAlias := '';
        sL_Msg := '['+ datetimetostr(now)+ '] ' +' �L�k���s�إ� Nagra Connection. ' + '  �@�~����:' + sI_Detail ;

        writeMsg;
        releaseResource;
        //down, �I�s�ʱ���, ���ʱ����� CA �s�_��
        {
        L_Intf := CoMonitor.Create;
        L_Intf.RunCa(sG_NewIniFileName);
        L_Intf := nil;
        }
        //up, �I�s�ʱ���, ���ʱ����� CA �s�_��
      end
      else if nI_ActiveOperation=SOCKET_NOT_READY_FOR_READING_FLAG then
      begin
        bL_KillSelf := true;
        sL_DbAlias := '';
        sL_Msg := '['+ datetimetostr(now)+ '] ' +'Socket is not ready for reading.' + '  �@�~����:' + sI_Detail ;

        writeMsg;
        releaseResource;
        //down, �I�s�ʱ���, ���ʱ����� CA �s�_��
        {
        L_Intf := CoMonitor.Create;
        L_Intf.RunCa(sG_NewIniFileName);
        L_Intf := nil;
        }
        //up, �I�s�ʱ���, ���ʱ����� CA �s�_��


      end

      else if nI_ActiveOperation=EXECUTE_ORACLE_SP_ERROR_FLAG then
      begin
        bL_KillSelf := true;
        sL_DbAlias := '';
        sL_Msg := '['+ datetimetostr(now)+ '] ' + '�L�k���� oracle ��ݵ{��.'+ '  �@�~����:' + sI_Detail ;

        writeMsg;
        releaseResource;
        //down, �I�s�ʱ���, ���ʱ����� CA �s�_��
        {
        L_Intf := CoMonitor.Create;
        L_Intf.RunCa(sG_NewIniFileName);
        L_Intf := nil;
        }
        //up, �I�s�ʱ���, ���ʱ����� CA �s�_��


      end

      else if nI_ActiveOperation=REBUILD_SEND_CMD_CONNECTION_TO_NDS then
      begin
        bL_ShowMemo := false;
        sL_DbAlias := '';
        sL_Msg := '['+ datetimetostr(now)+ '] ' +  ' ���s�إ߻P Nagra �� connection.'+ '  �@�~����:' + sI_Detail;
        writeMsg;
      end
      else if nI_ActiveOperation=UPDATE_CMD_STATUS_ERROR_FLAG then
      begin
        sL_DbAlias := '';
        sL_Msg := '['+ datetimetostr(now)+ '] ' +  '�L�k���� SP �ӧ�s���O���A.'+ '  �@�~����:' + sI_Detail;
        writeMsg;
      end
      else if nI_ActiveOperation=RE_CONNECT_NDS_ERROR_FLAG then
      begin
        sL_DbAlias := '';
        sL_Msg := '['+ datetimetostr(now)+ '] ' + '�L�k�P Nagra �إ߳s�u.'+ '  �@�~����:' + sI_Detail;
        writeMsg;
      end
      else if nI_ActiveOperation=CLOSE_SEND_NDS_FLAG then
      begin
        sL_DbAlias := '';
        sL_Msg := '['+ datetimetostr(now)+']' + '  �@�~����:' + sI_Detail;
        writeMsg;
      end
      else if nI_ActiveOperation=HANDLE_RESPONSE_ERROR_FLAG then
      begin
        sL_DbAlias := '';
        sL_Msg := '['+ datetimetostr(now)+ '] ' + '�B�z Nagra ���^�����~.'+ '  �@�~����:' + sI_Detail;
        writeMsg;
      end
      else if nI_ActiveOperation=WRITE_TO_MIS_ERROR_FLAG then
      begin
        sL_DbAlias := '';
        sL_Msg := '['+ datetimetostr(now)+ '] ' + '�^�g���O���A���~.'+ '  �@�~����:' + sI_Detail;
        writeMsg;
      end
      else if nI_ActiveOperation=RECEIVE_CMD_RESPONSE_ERROR_FLAG then
      begin
        sL_DbAlias := '';
        sL_Msg := '['+ datetimetostr(now)+ '] ' + '���� Nagra ���^�����~.'+ '  �@�~����:' + sI_Detail;
        writeMsg;
        
        releaseResource;
        //down, �I�s�ʱ���, ���ʱ����� CA �s�_��
        {
        L_Intf := CoMonitor.Create;
        L_Intf.RunCa(sG_NewIniFileName);
        L_Intf := nil;
        }
        //up, �I�s�ʱ���, ���ʱ����� CA �s�_��
      end
      else if nI_ActiveOperation=BAD_CLIENT_DATA_FORMAT then
      begin
        sL_DbAlias := '';
        sL_Msg := '['+ datetimetostr(now)+']' + '  �@�~����:' + sI_Detail ;
        writeMsg;
      end
      else if nI_ActiveOperation=EXECUTE_SEND_NDS_FLAG then
      begin
        sL_DbAlias := '';
        sL_Msg := '['+ datetimetostr(now)+']' + '  �@�~����:' + sI_Detail + '. �Ұʪ�:' + sG_LoginName;
        writeMsg;
      end;

      if bL_KillSelf then //�����|�Q�����, �]���I�s  L_Intf.RunCa ��, �N�|�� ca command gateway kill ��
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
      if sI_Operator = BATCH_OPERATOR then Exit; // �妸�@�~ �Ҳ��ͪ����O����� UI
//      if StrToIntDef(Copy(sI_TransNo,4,6),0)=0 then Exit;
      if nG_LivItemPos=CA_UI_MAX_COUNT then
      begin
//        clearStatusItem; //�I�s�� function...���G�|�� performance �U��...
        nG_LivItemPos := 0;
      end;
      L_ListItem := livStatus.Items[frmMain.nG_LivItemPos];



      for jj:=0 to LISTVIEW_SUB_ITEM_COUNT-1 do
       L_ListItem.SubItems.Add('');

      Inc(nG_LivItemPos);

      //down, �N�U�@�� ListItem �M�����ť�
      if nG_LivItemPos=CA_UI_MAX_COUNT then
        nL_NextItemIndex := 0
      else
        nL_NextItemIndex := nG_LivItemPos;
      clearStatusItem(nL_NextItemIndex);
      //up, �N�U�@�� ListItem �M�����ť�

//      L_ListItem.StateIndex := 0;
      L_ListItem.Caption := sI_TransNum;
      L_ListItem.SubItems[0] := sI_CompName; ;
      L_ListItem.SubItems[1] := sI_SeqNo;
      L_ListItem.SubItems[2] := sI_CommandID + '-' + sG_CurActiveLowLevelCmd;
      L_ListItem.SubItemImages[3] := 2;//���Ĥ�icon  => �ǰe
      L_ListItem.SubItemImages[4] := -1; //�M�� icon => �^��
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
    sL_Info := '�{���Y�N����,�Э��s���榹�{��!';
    
    G_ListenerThread := TListerner.Create(sG_ProcessorIp, sG_ProcessorListenPort );
    if GetExitCodeThread(G_ListenerThread.ThreadID,G_ListenerThread_ExitCode) then
    begin
      if not bG_AutoInvoke then
        MessageDlg('����� 2 ���ͥ���!' + sL_Info, mtError, [mbOK],0);

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
      nL_Result := nG_CmdRefreshRate1*1000 //�ǤJ���ɶ����y�p�ɬq
    else
      nL_Result := nG_CmdRefreshRate2*1000; //�ǤJ���ɶ������p�ɬq

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
    Exit; //����...
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

      //action ���i���(A0) : S000000051I00000001006600000000000000000000000000000000000000987600541C0000000000000000000000000000E051938000003O
        


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
    //sI_SettingStr => �ѳ\�h���q�O�Ҧ�_�Ӫ��r��. Ex:9,10,12

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

    //down, �إ� TcpClient �� array...    
    createTcpClient;
    //up, �إ� TcpClient �� array...end;

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
    Panel1.Caption := ' �Y�ɺʱ�--���O�ǰe��';
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

        L_ListItem.SubItemImages[4] := nL_ImgNdx; //�M�� icon => �^��
     end;

end;

procedure TfrmMain.preSendToDispatcher(sI_CompID, sI_HighLevelCmdID, sI_DispatcherSeqNo, sI_ResponseDataStatus,
sI_ErrorCode, sI_ErrorMsg, sI_AttachedResponseData:String);
var
    sL_Response : String;
    L_Obj : TCompAndIpInfoObj;
begin

    sL_Response := sI_ResponseDataStatus + Format('%.2d',[length(sI_CompID)]) + sI_CompID + Format('%.2d',[length(sI_DispatcherSeqNo)]) + sI_DispatcherSeqNo ;
    sL_Response := sL_Response + Format('%.3d',[length(sI_ErrorCode)]) + sI_ErrorCode ;
    sL_Response := sL_Response + Format('%.3d',[length(sI_ErrorMsg)]) + sI_ErrorMsg ;
    sL_Response := sL_Response + sI_AttachedResponseData;





    {
    �i�઺ response string :
    �������O A0, A1 => 'C01903123000000000000051' => C 01 9 03 123 000 000 00000005 1
    ��L�������O  => 'C01903123000000' => C 01 9 03 123 000 000
    }
    sendToDispatcher(SEND_TO_DISPATCHER_INFO_TYPE_1,sI_CompID, sL_Response,'','','','');

end;


procedure TfrmMain.getDispatcherSeqNo(bI_SequenceIdIsHax: Boolean;sI_SequenceID: String; var sI_CompCode, sI_HighLevelCmdID, sI_DispatcherSequenceNo:String);
var
    sL_SeqNo, sL_SequenceID : String;
    nL_Ndx : Integer;
    L_Obj : TSequenceIdMappingInfoObj;
begin

    if bI_SequenceIdIsHax then
      sL_SequenceID := IntToStr(TUstr.hexToInt(sI_SequenceID))
    else
      sL_SequenceID := sI_SequenceID;
            
//    sL_SequenceID := IntToStr(StrToInt(sI_SequenceID));
    nL_Ndx := G_SeqTransMappingStrList.IndexOf(sL_SequenceID);
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
//showmessage('comp id:' + sI_CompID + IntTostr(nI_Type));
        case nI_Type of
          SEND_TO_DISPATCHER_INFO_TYPE_1 :
          begin
//            showmessage('a');
            sL_Header := IntToStr(nI_Type);
          end;
          SEND_TO_DISPATCHER_INFO_TYPE_2 :
          begin
//            showmessage('b');
            sL_Header := IntToStr(nI_Type) + Format('%.2d',[length(sI_CompID)]) + sI_CompID + Format('%.4d',[length(sI_HighLevelCmdID)]) + sI_HighLevelCmdID + Format('%.3d',[length(sI_Operator)]) + sI_Operator;
          end;
          SEND_TO_DISPATCHER_INFO_TYPE_3 :
          begin
//            showmessage('c');
            sL_Header := IntToStr(nI_Type) + Format('%.2d',[length(sI_CompID)]) + sI_CompID + sI_SrcTransNum + sI_ResponseCmdID;
          end
          else
            sL_Header := IntToStr(SEND_TO_DISPATCHER_INFO_DEFAULT_TYPE);
        end;
//showmessage('1');
        L_Obj := G_RemoteInfoStrList.Objects[nL_Ndx] As TCompAndIpInfoObj;
//showmessage('2');
        sL_DispatcherIp := L_Obj.IP;
//showmessage('dispatcher IP:' + sL_DispatcherIp);
        for ii:=0 to High(G_AryOfTcpClient) do
        begin
//showmessage('client ip : ' + G_AryOfTcpClient[ii].RemoteHost);
          if G_AryOfTcpClient[ii].RemoteHost = sL_DispatcherIp then
          begin
//            Memo3.Lines.Add(sL_Header + sI_Info);
{
            if G_AryOfTcpClient[ii].Connected then
              showmessage('connected..')
            else
              showmessage('not connected..');
}
            G_AryOfTcpClient[ii].Sendln(sL_Header + sI_Info);

{
�Ǧ^���i���

SEND_TO_DISPATCHER_INFO_TYPE_1(Ex:1C0190800123456000000) => 1(���O) C(���O�B�z���G) 01(���q�N�X����) 9(���q�N�X) 08(SeqNo����) 00123456(SeqNo) 000 (���~�T���N�X����) [�i��|�����~�T���N�X] 000 (���~�T������) [�i��|�����~�T���N�X]
SEND_TO_DISPATCHER_INFO_TYPE_1(Ex:1C0190800123456000000xxxxx) => 1(���O) C(���O�B�z���G) 01(���q�N�X����) 9(���q�N�X) 08(SeqNo����) 00123456(SeqNo) 000 (���~�T���N�X����) [�i��|�����~�T���N�X] 000 (���~�T������) [�i��|�����~�T���N�X] xxxxx (�@�ꪺ response data)
SEND_TO_DISPATCHER_INFO_TYPE_2 => 2(���O)   01(���q�N�X����) 9(���q�N�X) 0002(�������O�N�X����) A1(�������O�N�X) 006(�ާ@�̪���) howard(�ާ@��) 00900000301000302574428920030408N2003040720030409U1160194446005211911825900000 (���O�r��)
SEND_TO_DISPATCHER_INFO_TYPE_3 => 3(���O)   01(���q�N�X����) 9(���q�N�X) 009000004(������X) 1000(�^�������O�N��) 000008972050257000344289200303081000009000004000000000000000000000000(���㤧�^���r��)
}
            break;
          end;
        end;
      end;
    except

    end;
    Application.ProcessMessages;
end;

procedure TfrmMain.Button1Click(Sender: TObject);
var
    sL_Data : String;
begin

{
    //down, �}��+�W�D���v
    sL_Data := '00123456  9    �����s  A000000000       10066        200304308899    198760054E28    Howard  254~20040531~N;5~20041231~A     4';
    parseClientNdsData(sL_Data);
    //���ͤ����� command : '0003S00400001010100200304271617350001SIH000009D4C00123456789098760054E1C050300030017A000420040531NA000520041231AAFC1115B3EBD670F';
    //up, �}��+�W�D���v

    Memo2.Lines.Add('-----------------------');
    //down, �}��
    sL_Data := '00123456  9    �����s  A100000000       10522                8899    198760054E28    Howard   0      ';
    parseClientNdsData(sL_Data);
    //���ͤ����� command : '0003S00240001010100200304271615170001SIH000009D4C00123456789098760054E1C05030003001725027EB30300A98F';
    //up, �}��

    Memo2.Lines.Add('-----------------------');
    //down, Change Subscriber Data
    sL_Data := '00123456  9    �����s  A6000009D4                            8899    198760054E15    Howard   0      ';
    parseClientNdsData(sL_Data);
    //���ͤ����� command : '0003S00240001010100200304271615170001SIH000009D4C00123456789098760054E1C05030003001725027EB30300A98F';
    //up, Change Subscriber Data



    Memo2.Lines.Add('-----------------------');
    //down, ���
    sL_Data := '00123456  9    �����s  A2000009D4  1234567890        200304308899    198760054E28    Howard   0     6';
    parseClientNdsData(sL_Data);
    //���ͤ����� command : '0003S00010001010100200304271633430003SHH000009D4D60D6CC1683ADFF1B';
    //up, ���

    Memo2.Lines.Add('-----------------------');
    //down, �}�W�D
//    sL_Data := '00123456  9    �����s  B1000009D4  1234567890        200304308899    198760054E28    Howard  254~20040531~N;3~20041231~A     7';
    sL_Data := '00123456  9    �����s  B1000009D4  1234567890        200304308899    198760054E28    Howard  124~20051231~A     7';
    parseClientNdsData(sL_Data);
    //���ͤ����� command : '0003S001C0001010100200304271638490004SIH000009D4A000420040531NA000520041231A44569EB9D9BB88DD';
    //up, �}�W�D

    Memo2.Lines.Add('-----------------------');
    //down, ���W�D
    sL_Data := '00123456  9    �����s  B2000009D4  1234567890        200304308899    198760054E28    Howard   14     8';
    parseClientNdsData(sL_Data);
    //���ͤ����� command : '0003S000F0001010100200304271652010005SHH000009D4X0007X0008X00094F0EBC5442D9A65E';
    //up, ���W�D


    Memo2.Lines.Add('-----------------------');
    //down, reset PIN Code
    sL_Data := '00123456  9    �����s  E1000009D4  1234567890        200304308899    198760054E28    Howard   0U    9';
    parseClientNdsData(sL_Data);
    //���ͤ����� command : '0003S00170001010100200304281034460001SIN000009D4q400L8899000000FFFF0000208AC9EE0A9B8208';
    //up,  reset PIN Code



    Memo2.Lines.Add('-----------------------');
    //down, ����
    sL_Data := '00123456  9    �����s  A3000009D4  1234567890        200304308899    198760054E28    Howard   0L   10';
    parseClientNdsData(sL_Data);
    //���ͤ����� command : '0003S00170001010100200304281034460001SIN000009D4q400L8899000000FFFF0000208AC9EE0A9B8208';
    //up,  ����


    Memo2.Lines.Add('-----------------------');
    //down, �_��
    sL_Data := '00123456  9    �����s  A4000009D4  1234567890        200304308899    198760054E28    Howard   0L   11';
    parseClientNdsData(sL_Data);
    //���ͤ����� command : '0003S00170001010100200304281034460001SIN000009D4q400L8899000000FFFF0000208AC9EE0A9B8208';
    //up,  �_��



    Memo2.Lines.Add('-----------------------');
    //down, Store Region Key
    sL_Data := '00123456  9    �����s  E2000009D4  1234567890        200304308899    198760054E28    Howard   0L   12';
    parseClientNdsData(sL_Data);
    //���ͤ����� command : '0003S000B0001010100200305051007030001ZIH98760054D2A074D1001484B474CBB00BF29';
    //up,  Store Region Key


    Memo2.Lines.Add('-----------------------');
    //down, Create Service/PPV/OPPV
    sL_Data := '00123456  9    �����s  B300000004  1234567890        200304308899    198760054E28    Howard  23Super Ball~20030601~O~PL  234';
    parseClientNdsData(sL_Data);
    //���ͤ����� command : '0003M00530001010100200305201727540001RINFFFFC0ASuper Ball20030601OPF3B9BDF12069B3E9';
    //up,  Create Service/PPV/OPPV

    Memo2.Lines.Add('-----------------------');
    //down, resend all packets
    sL_Data := '00123456  9    �����s  A500000006  1234567890        200304308899    198760054E28    Howard   0L   11';
    parseClientNdsData(sL_Data);
    //���ͤ����� command : '';
    //up,  resend all packets


    Memo2.Lines.Add('-----------------------');
    //down, resend all packets
    sL_Data := '00123456  9    �����s  F2000009D4  1234567890        200304308899    198760054E28    Howard  211~3~1000~100~300~16~CL  234';
    parseClientNdsData(sL_Data);
    //���ͤ����� command : '';
    //up,  Create/Modify IPPV Series slot


    Memo2.Lines.Add('-----------------------');
    //down, resend all packets
    sL_Data := '00123456  9    �����s  F1000009D4  1234567890        200304308899    198760054E28    Howard   5D,110L  234';
    parseClientNdsData(sL_Data);
    //���ͤ����� command : '';
    //up,  set personal region bit                                                           


    Memo2.Lines.Add('-----------------------');
    //down, resend all packets
    sL_Data := '00123456  9    �����s  F6000009D4  1234567890        200304308899    198760054E28    Howard   712~13~9L  234';
    parseClientNdsData(sL_Data);
    //���ͤ����� command : '';
    //up,  set personal region bytes


    Memo2.Lines.Add('-----------------------');
    //down, resend all packets
    sL_Data := '00123456  9    �����s  F7000009D4  1234567890        200304308899    198760054E28    Howard  1297~112~25555L  234';
    parseClientNdsData(sL_Data);
    //���ͤ����� command : '';
    //up,  set personal region bit
}


    Memo2.Lines.Add('-----------------------');
    //down, resend all packets
    sL_Data := '00123456  9    �����s  G2000009D4  1234567890        200304308899    198760054E28    Howard   8hello~20L  234';
    parseClientNdsData(sL_Data);
    //���ͤ����� command : '';
    //up,  send OSD
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

    sL_CompId, sL_DispatcherSequenceNo,sL_TempNotes : String;
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



    if Trim(sL_IccNo)='' then
      sL_IccNo:= '0';
    sL_IccNo := sL_IccNo + IntToStr(getCardIdChecksum(StrToFloat(sL_IccNo)));
    sG_CardID := TUstr.AddString(sL_IccNo,'0', false, ICC_NO_LENGTH);
    sG_RegionKey := TUstr.AddString(sL_RegionKey,'0', false, REGION_KEY_LENGTH);

    sG_ReportbackAvailability := sL_Reportbackavailability;

    sG_ReportbackDate := IntToHex(StrToInt(sL_ReportbackDate), 2);

    if sL_PopolcationID = '' then
    begin
      if (sG_HighLevelCmdId='A1') and (bG_LabStatus) then
      begin
        //���p�b�}����Lab���A�U,PopulationID�Ƕi�Ū�,�h�w�]�@�Өt�ΰѼ�
        sG_PopulationID := IntToHex(nG_DefaultPopulationID,4)
      end
      else
         sG_PopulationID := '';
    end
    else
      sG_PopulationID := IntToHex(StrToIntDef(sL_PopolcationID,0),4);


    sL_DetailCommandIDs := getDetailCommandIDs(sL_HighLevelCmdId);
    if sL_DetailCommandIDs='' then Exit;

    sL_CmdEnvelop := sG_Version + sG_SecurityType ;


    sL_CmdHeader := sL_FromID + IntToHex(nG_ConnectionID,2) + sG_ToID + TUstr.replaceStr(TUdateTime.GetPureDateTimeStr(now,'',''),' ','') + IntToHex(nG_SequenceID,4) + sL_ActionType + sL_ActionPriority + sL_PriorityReassignment;

    if sL_ActionType='S' then
      sL_CmdHeader := sL_CmdHeader + TUstr.AddString(sL_SubscriberID,'0',false,SUBSCRIBER_ID_LENGTH)
    else if sL_ActionType='Z' then
      sL_CmdHeader := sL_CmdHeader + sG_RegionKey
    else if sL_ActionType='A' then
      sL_CmdHeader := sL_CmdHeader + TUstr.AddString(sL_SubscriberID,'0',false,SUBSCRIBER_ID_LENGTH)
    else if sL_ActionType='R' then
      sL_CmdHeader := sL_CmdHeader + IntToHex(StrToIntDef(sL_ServiceID,0),4);


    L_StrList := TUStr.ParseStrings(sL_DetailCommandIDs,',');

    for ii:=0 to L_StrList.Count -1 do
    begin
      if (sG_HighLevelCmdId='A1') and (sG_PopulationID='') then
      begin
        sL_ActionBody := BAD_CMD_FORMAT;
      end
      else
      begin
        if (sG_HighLevelCmdId='A1') AND (L_StrList.Strings[ii]='y') then
        begin
          //�Y���}��(A1)�B�n�]Bouquet_ID
          if TUstr.isHexValue(sL_Notes) then   //�P�O�O�_��16�i��榡
          begin
            sL_TempNotes := '12~13~' + IntToStr(TUstr.hexToInt(sL_Notes));
            sL_ActionBody := sL_ActionBody + getActionBody(sL_ActionType, L_StrList.Strings[ii], sL_TempNotes);
          end
          else
            sL_ActionBody := BAD_CMD_FORMAT;
        end
        else
        begin
          sL_TempNotes := sL_Notes;
          sL_ActionBody := sL_ActionBody + getActionBody(sL_ActionType, L_StrList.Strings[ii], sL_TempNotes);
        end;
      end;
      {
      if (sL_HighLevelCmdId='A0') then
      begin
        sL_ActionBody := sL_ActionBody + getActionBody(sL_ActionType, L_StrList.Strings[ii], sL_Notes);
      end;
      }
    end;

    recordSeqAndTransMapping(sL_CompCode, sL_SeqNo, IntToStr(nG_SequenceID), sL_HighLevelCmdId);

    if AnsiPos(BAD_CMD_FORMAT, sL_ActionBody)>0 then //��ܰ�������O�榡����,�ҥH�����^���� dispatcher...
    begin
      getDispatcherSeqNo(false, IntToStr(nG_SequenceID), sL_CompId, sL_HighLevelCmdID, sL_DispatcherSequenceNo);
      preSendToDispatcher(sL_CompId, sL_HighLevelCmdID, sL_DispatcherSequenceNo, 'E', 'E999', BAD_CMD_FORMAT, ''); // ���զ^���� dispatcher...
    end
    else
    begin
      nL_MsgLength := length(sL_CmdEnvelop + sL_CmdHeader + sL_ActionBody) + 4 + 16; // 4=> Length ����������, 16=> Signature ������

      sL_CmdEnvelop := sL_CmdEnvelop + IntToHex(nL_MsgLength,4);


      sL_Signature := getSignature(sL_CmdEnvelop + sL_CmdHeader + sL_ActionBody , sL_SignatureKey, sG_SecurityType);


      sG_FullCmd := sL_CmdEnvelop + sL_CmdHeader + sL_ActionBody + sL_Signature;

//      recordSeqAndTransMapping(sL_CompCode, sL_SeqNo, IntToStr(nG_SequenceID), sL_HighLevelCmdId);

      if bG_AutoResponse then
      begin
        getDispatcherSeqNo(false, IntToStr(nG_SequenceID), sL_CompId, sL_HighLevelCmdID, sL_DispatcherSequenceNo);
        if (sL_HighLevelCmdID='A0') or (sL_HighLevelCmdID='A1') then
          preSendToDispatcher(sL_CompId, sL_HighLevelCmdID, sL_DispatcherSequenceNo, 'C', '', '', '00000004~1#000000010058~000000000000~00000000~98760054~1C~0000000000000000000000000000~D~05~083746~0003') // ���զ^���� dispatcher...
        else if (sL_HighLevelCmdID='A6') then
          preSendToDispatcher(sL_CompId, sL_HighLevelCmdID, sL_DispatcherSequenceNo, 'C', '', '', '12') // ���զ^���� dispatcher...
        else
          preSendToDispatcher(sL_CompId, sL_HighLevelCmdID, sL_DispatcherSequenceNo, 'C', '', '', ''); // ���զ^���� dispatcher...
      end
      else
        preSendCmd(sG_FullCmd);
    end;  
    Inc(nG_SequenceID);
end;


procedure TfrmMain.getActionTypeAndPriority(sI_HighLevelCommand: String;
  var sI_ActionType, sI_ActionPriority, sI_PriorityReassignment, sI_FromID: String);
var
    sL_TmpData,sL_TempHighLevelCommand : String;
    L_StrList,L_TempStrList : TStringList;
    ii : Integer;
begin
    for ii:=0 to G_ActionTypeAndPriorityStrList.Count -1 do
    begin
      //jackal920813�ק�
      L_TempStrList := TUstr.ParseStrings(G_ActionTypeAndPriorityStrList.Strings[ii],'=');
      if L_TempStrList.Count = 2 then
      begin
        sL_TempHighLevelCommand := L_TempStrList.Strings[0];

        //if sI_HighLevelCommand = Copy(G_ActionTypeAndPriorityStrList.Strings[ii],1,2) then
        if sI_HighLevelCommand = sL_TempHighLevelCommand then
        begin
          //sL_TmpData := Copy(G_ActionTypeAndPriorityStrList.Strings[ii],4, length(G_ActionTypeAndPriorityStrList.Strings[ii])-3);
          sL_TmpData := L_TempStrList.Strings[1];
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
            MessageDlg('ini �ɤ� Action Type �榡�]�w���~.�{���Y�N����!', mtError, [mbOk],0);
            Application.Terminate;
          end;
          break;
        end;
      end;

    end;


end;

function TfrmMain.getActionBody(sI_ActionType,
  sI_LowLevelCmdID, sI_Notes: String): String;
var
    sL_FromRegionBit, sL_ToRegionBit, sL_RegionBitMask : String;
    nL_ParentalRating,nL_EventSpendingLimit,    nL_CumulativeBalance : Integer;

    sL_StartServicesID,sL_EndServicesID,sL_ServiceType,sL_TempAuthorizeService : String;
    nL_StartServicesID,nL_EndServicesID,nL_ServicesIDCount : Integer;

    nL_FromRegionByte, nL_ToRegionByte, nL_BitLength : Integer;
    sL_FromRegionByte, sL_ToRegionByte, sL_UserNumber, sL_Status, sL_RegionNum : String;
    nL_TmpParameterValue, ii, jj : Integer;
    sL_TmpBin, sL_TmpHex,sL_TmpResultHex,  sL_Result : String;
    L_StrList1, L_StrList2, L_StrList3 : TStringList;
    nL_ServiceID, nL_LengthOfServiceName : Integer;
    sL_ExpirationDate, sL_TapingAuthorization : String;

    sL_RegionByteMask, sL_PayFreeFlag, sL_ServiceName, sL_TypeOfService : String;

    bL_HasMultiByte : boolean;
    sL_ContentHexValue, sL_Field1, sL_OsdContent, sL_OsdDuration : String;
    nL_MasterSubID, nL_SlotNum, nL_SeriesID, nL_Balance, nL_Limit, nL_Threshold, nL_MasIppvRecords : Integer;
    sL_BalanceReadjusted :String;
    nL_EvenThreshold : Integer;
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
      else if sI_LowLevelCmdID='S' then //suspend services
      begin
        sL_Result := 'S';
      end
      else if sI_LowLevelCmdID='R' then //reactivate subscriber
      begin
        sL_Result := 'R';
      end
      else if sI_LowLevelCmdID='a' then //re-authorize all services
      begin
        sL_Result := 'a';
      end
      else if sI_LowLevelCmdID='Q' then // change subscriber data
      begin
        sL_Result := 'Q'  + sG_RegionKey + sG_ReportbackAvailability + sG_ReportbackDate + sG_BackwardTolerance + sG_ForwardTolerance + sG_Currency + sG_PopulationID;
      end
      else if sI_LowLevelCmdID='J' then //cancal all services
      begin
        sL_Result := 'J';
      end
      else if sI_LowLevelCmdID='G' then //get information
      begin
        sL_Result := 'G';
        if (sI_Notes='M') or  (sI_Notes='I') or  (sI_Notes='P')
          or  (sI_Notes='U') or  (sI_Notes='T') or  (sI_Notes='V') then
          sL_Result := sL_Result + sI_Notes
        else
          sL_Result :=  BAD_CMD_FORMAT;
      end
      else if sI_LowLevelCmdID='e' then //set master/slave relation
      begin
        sL_Result := 'e';
        L_StrList1 := TUstr.ParseStrings(sI_Notes,'~');
        if (L_StrList1<> nil) and ((L_StrList1.Count=2) or (L_StrList1.Count=1)) then
        begin
          if (L_StrList1.Count=1) and (L_StrList1.Strings[0]='1') then //��� set as master
            sL_Result := sL_Result + IntToHex(1,2) + IntToHex(0,8)
          else if (L_StrList1.Count=2) and (L_StrList1.Strings[0]='0') then //��� set as slave
          begin
            nL_MasterSubID := StrToIntDef(L_StrList1.Strings[1],0);
            if nL_MasterSubID=0 then
              sL_Result :=  BAD_CMD_FORMAT
            else
              sL_Result := sL_Result + IntToHex(0,2) + IntToHex(nL_MasterSubID,8);//���data���X�z�� master subscriber id
          end
          else
            sL_Result :=  BAD_CMD_FORMAT;
        end
        else
          sL_Result :=  BAD_CMD_FORMAT;
      end
      else if sI_LowLevelCmdID='W' then //Initiate Reportback
      begin
        sL_Result := 'W';
        L_StrList1 := TUstr.ParseStrings(sI_Notes,'~');
        if (L_StrList1<> nil) and (L_StrList1.Count=2) then
        begin
          if (L_StrList1.Strings[0]='I') or (L_StrList1.Strings[0]='D') and (L_StrList1.Strings[1]='D') then
            sL_Result := sL_Result + L_StrList1.Strings[0] + L_StrList1.Strings[1]
          else
            sL_Result :=  BAD_CMD_FORMAT;
        end
        else
          sL_Result :=  BAD_CMD_FORMAT;
      end
      else if sI_LowLevelCmdID='c' then // Create/Modify IPPV Series Slot
      begin
        sL_Result := 'c';
        L_StrList1 := TUstr.ParseStrings(sI_Notes,'~');
        if (L_StrList1<> nil) and (L_StrList1.Count=7) then
        begin
          nL_SlotNum := StrToIntDef(L_StrList1.Strings[0],0);
          nL_SeriesID := StrToIntDef(L_StrList1.Strings[1],0);
          nL_Balance := StrToIntDef(L_StrList1.Strings[2],0);
          nL_Limit := StrToIntDef(L_StrList1.Strings[3],0);
          nL_Threshold := StrToIntDef(L_StrList1.Strings[4],0);
          nL_MasIppvRecords := StrToIntDef(L_StrList1.Strings[5],0);
          sL_BalanceReadjusted := L_StrList1.Strings[6];
          sL_Result := sL_Result + IntToHex(nL_SlotNum,2) + IntToHex(nL_SeriesID,4);
          sL_Result := sL_Result + IntToHex(nL_Balance,4) + IntToHex(nL_Limit,4);
          sL_Result := sL_Result + IntToHex(nL_Threshold,4) + IntToHex(nL_MasIppvRecords,2);
          sL_Result := sL_Result + sL_BalanceReadjusted;

        end
        else
          sL_Result :=  BAD_CMD_FORMAT;
      end
      else if sI_LowLevelCmdID='y' then //Set Personal Region Bytes
      begin
        sL_Result := 'y';
        L_StrList1 := TUstr.ParseStrings(sI_Notes,'~');
        if (L_StrList1<> nil) and (L_StrList1.Count=3)then
        begin
          sL_FromRegionByte := L_StrList1.Strings[0];
          sL_ToRegionByte := L_StrList1.Strings[1];
          sL_RegionByteMask := L_StrList1.Strings[2];



          nL_FromRegionByte := StrToIntDef(sL_FromRegionByte,-1);
          nL_ToRegionByte := StrToIntDef(sL_ToRegionByte,-1);


          if (nL_ToRegionByte<>-1) and  (nL_FromRegionByte<>-1) and (nL_ToRegionByte >= nL_FromRegionByte)then
          begin

            sL_TmpResultHex := '';
            sL_TmpHex := IntToHex(StrToInt(sL_RegionByteMask),(nL_FromRegionByte-nL_ToRegionByte)*2);
            sL_TmpResultHex := sL_TmpResultHex + sL_TmpHex;

{
            for ii :=nL_FromRegionByte to nL_ToRegionByte do
            begin

              sL_TmpHex := IntToHex(StrToInt(sL_RegionByteMask),2);
              sL_TmpResultHex := sL_TmpResultHex + sL_TmpHex;

            end;
}
            sL_Result := sL_Result + IntToHex(nL_FromRegionByte,2) + IntToHex(nL_ToRegionByte,2) + sL_TmpResultHex;
//            sL_Result := sL_Result + IntToHex(nL_FromRegionByte,2) + IntToHex(nL_ToRegionByte,2) + '6666'; //���ե�...
          end
          else
            sL_Result :=  BAD_CMD_FORMAT;
        end
        else
          sL_Result :=  BAD_CMD_FORMAT;
      end
      else if sI_LowLevelCmdID='Y' then //Set Personal Region Bit
      begin
        sL_Result := 'Y';
        L_StrList1 := TUstr.ParseStrings(sI_Notes,'~');
        if L_StrList1<> nil then
        begin
          if (L_StrList1.Count=2) and ((L_StrList1.Strings[0]='A') or (L_StrList1.Strings[0]='D')) then//��ܬ�set region bit
          begin
            sL_RegionNum := L_StrList1.Strings[1];
            sL_Result := sL_Result + L_StrList1.Strings[0] + IntToHex(StrToIntDef(sL_RegionNum,0),2);
          end
          else
            sL_Result :=  BAD_CMD_FORMAT;
        end;        

      end
      else if sI_LowLevelCmdID='x' then //Set Personal Region Bits
      begin
        sL_Result := 'x';
        L_StrList1 := TUstr.ParseStrings(sI_Notes,'~');
        if L_StrList1<> nil then
        begin
          if (L_StrList1.Count=3) and (StrToInt(L_StrList1.Strings[1]) - StrToInt(L_StrList1.Strings[0])>0) then//��ܬ�set region bits
          begin
            sL_FromRegionBit := L_StrList1.Strings[0];
            sL_ToRegionBit := L_StrList1.Strings[1];

            nL_BitLength :=  StrToInt(sL_ToRegionBit) - StrToInt(sL_FromRegionBit) + 1;
            sL_RegionBitMask := L_StrList1.Strings[2];

            sL_RegionBitMask := IntToBin(StrToInt(sL_RegionBitMask));
            sL_RegionBitMask := Copy(sL_RegionBitMask, length(sL_RegionBitMask)-nL_BitLength+1, nL_BitLength);

            sL_Result := sL_Result + IntToHex(StrToInt(sL_FromRegionBit),2) + IntToHex(StrToInt(sL_ToRegionBit),2) + sL_RegionBitMask;
          end
          else
            sL_Result :=  BAD_CMD_FORMAT;
        end;

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
            end
            else
              sL_Result :=  BAD_CMD_FORMAT;
          end;
        end
        else
          sL_Result :=  BAD_CMD_FORMAT;
      end
      else if sI_LowLevelCmdID='X' then //cancel service
      begin
        L_StrList1 := TUstr.ParseStrings(sI_Notes,'~');
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
          nL_TmpParameterValue := 128;
        sL_Result := sL_Result + 'q' + IntToHex(nL_TmpParameterValue,2) + sL_UserNumber + sG_PinControl + sG_PinCode + IntToHex(nL_ParentalRating,6) + IntToHex(nL_EventSpendingLimit,4) + IntToHex(nL_CumulativeBalance,4);

      end
      else if sI_LowLevelCmdID='V' then // resend all packets
      begin
        sL_Result := sL_Result + 'V';
      end
      else if sI_LowLevelCmdID='j' then // Activate/Deactivate Subscriber
      begin
        sL_Result := 'j' + sI_Notes;
      end
      else if sI_LowLevelCmdID='B' then //Multi Service Authorization
      begin
        L_StrList1 := TUstr.ParseStrings(sI_Notes,'@');
        if L_StrList1<> nil then
        begin
          sL_StartServicesID := L_StrList1.Strings[0];
          sL_EndServicesID := L_StrList1.Strings[1];
          sL_ServiceType := L_StrList1.Strings[2];
          sL_TempAuthorizeService := L_StrList1.Strings[3];

          nL_StartServicesID := StrToIntDef(sL_StartServicesID,-1);
          nL_EndServicesID := StrToIntDef(sL_EndServicesID,-1);
          nL_ServicesIDCount := nL_EndServicesID - nL_StartServicesID + 1;

          //�ܤ֭n���@�� Services ID
          if (nL_StartServicesID<>-1) and  (nL_EndServicesID<>-1) and (nL_ServicesIDCount >= 1)then
          begin
//******************
            //�H�U�P Authorize service
            L_StrList2 := TUstr.ParseStrings(sL_TempAuthorizeService,';');
            if L_StrList2<> nil then
            begin
              //�ĥ|����줺���ռ�,�n����End Services ID-Start Services ID���ƶq,�B�O�n�����d�򤺪�Services ID
              if nL_ServicesIDCount = L_StrList2.Count then
              begin

                for ii:=0 to L_StrList2.Count -1 do
                begin
                  L_StrList3 := TUstr.ParseStrings(L_StrList2.Strings[ii],'~');
                  if L_StrList3.Count = 3 then
                  begin
                    nL_ServiceID := StrToInt(L_StrList3.Strings[0]);
                    if (nL_ServiceID >= nL_StartServicesID) and (nL_ServiceID <= nL_EndServicesID) then
                    begin
                      sL_ExpirationDate := L_StrList3.Strings[1];
                      sL_TapingAuthorization := L_StrList3.Strings[2];
                      sL_Result := sL_Result + 'A' + IntToHex(nL_ServiceID,4) + sL_ExpirationDate + sL_TapingAuthorization;
                    end
                    else
                      sL_Result :=  BAD_CMD_FORMAT;
                  end
                  else
                    sL_Result :=  BAD_CMD_FORMAT;
                end;

                sL_Result := 'B' + IntToHex(nL_StartServicesID,4) + IntToHex(nL_EndServicesID,4) + sL_ServiceType + IntToHex(nL_ServicesIDCount,2) + sL_Result;
              end
              else
                sL_Result :=  BAD_CMD_FORMAT;
            end
            else
              sL_Result :=  BAD_CMD_FORMAT;
//******************
          end
          else
            sL_Result :=  BAD_CMD_FORMAT;
        end
        else
          sL_Result :=  BAD_CMD_FORMAT;
      end
      else if sI_LowLevelCmdID='N' then //New Card
      begin
        sL_Result := 'N';
        L_StrList1 := TUstr.ParseStrings(sI_Notes,'~');
        if L_StrList1<> nil then
        begin
          if (L_StrList1.Count=2) and ((L_StrList1.Strings[1]='1') or (L_StrList1.Strings[1]='2') or (L_StrList1.Strings[1]='3')) then
          begin
            sL_Result := sL_Result + L_StrList1.Strings[0] + L_StrList1.Strings[1];
          end
          else
            sL_Result :=  BAD_CMD_FORMAT;
        end;
      end
      else if sI_LowLevelCmdID='M' then //Clear Card Transfer
      begin
        sL_Result := 'M';
      end
      else if sI_LowLevelCmdID='U' then //Cancel Pairing/Re-pairing
      begin
        sL_Result := 'U';
      end
      else if sI_LowLevelCmdID='g' then //Get Subscriber ID
      begin
        sL_Result := 'g' + sG_CardID;
      end
      else if sI_LowLevelCmdID='k' then //Resend geographic data
      begin
        sL_Result := 'k';
      end
      else if sI_LowLevelCmdID='s' then //Resend geographic data
      begin
        sL_Result := 's' + sG_Currency;
      end
      else if sI_LowLevelCmdID='b' then //Cancel Slot Use
      begin
        nL_SlotNum := StrToIntDef(sI_Notes,0);
        sL_Result := 'b' + IntToHex(nL_SlotNum,2);
      end
      else if sI_LowLevelCmdID='n' then //Set Non-viewable IPPV Bit
      begin
        L_StrList1 := TUstr.ParseStrings(sI_Notes,'~');
        if L_StrList1<> nil then
        begin
          if L_StrList1.Count = 2 then
          begin
            nL_ServiceID := StrToInt(L_StrList1.Strings[0]);
            sL_ExpirationDate := L_StrList1.Strings[1];

            sL_Result := sL_Result + 'n' + IntToHex(nL_ServiceID,4) + sL_ExpirationDate;
          end
          else
            sL_Result :=  BAD_CMD_FORMAT;
        end
        else
          sL_Result :=  BAD_CMD_FORMAT;
      end
      else if sI_LowLevelCmdID='z' then //Delete Non-viewable IPPV Bit
      begin
        L_StrList1 := TUstr.ParseStrings(sI_Notes,'~');
        if L_StrList1<> nil then
        begin
          if L_StrList1.Count = 2 then
          begin
            nL_ServiceID := StrToInt(L_StrList1.Strings[0]);
            sL_ExpirationDate := L_StrList1.Strings[1];

            sL_Result := sL_Result + 'z' + IntToHex(nL_ServiceID,4) + sL_ExpirationDate;
          end
          else
            sL_Result :=  BAD_CMD_FORMAT;
        end
        else
          sL_Result :=  BAD_CMD_FORMAT;
      end
      else if sI_LowLevelCmdID='K' then //Set Event Threshold
      begin
        nL_EvenThreshold := StrToIntDef(sI_Notes,-1);
        if (nL_EvenThreshold>=0) and (nL_EvenThreshold<=24) then
          sL_Result := 'K' + IntToHex(nL_EvenThreshold,2)
        else
          sL_Result :=  BAD_CMD_FORMAT;
      end


    end
    else if UpperCase(sI_ActionType)='A' then
    begin
      if sI_LowLevelCmdID='O' then //send OSD
      begin
        L_StrList1 := TUstr.ParseStrings(sI_Notes,'~');
        if L_StrList1.Count = 2 then
        begin

          sL_OsdContent := L_StrList1.Strings[0];
          sL_OsdDuration := L_StrList1.Strings[1];

          sL_ContentHexValue := TUstr.getMultiBytesHexValue(sL_OsdContent,bL_HasMultiByte);
          if bL_HasMultiByte then
            sL_Field1 := '100'
          else
            sL_Field1 := '0';
          sL_Result := sL_Result + 'O' + 'A' + IntToHex(StrToIntDef(sL_OsdDuration,20),2) + 'H' + '00' + 'FD' + inttoHex(length(sL_OsdContent),4) + '00' + IntToHex(StrToInt(sL_Field1),2) + '00' + sL_ContentHexValue;
        end
        else
          sL_Result :=  BAD_CMD_FORMAT;
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
        // sI_Notes �d��: Super Ball~20030601~O~P
        L_StrList1 := TUstr.ParseStrings(sI_Notes,'~');
        if L_StrList1.Count = 4 then
        begin
          sL_ServiceName := L_StrList1.Strings[0];
          nL_LengthOfServiceName := length(sL_ServiceName);
          sL_ExpirationDate := L_StrList1.Strings[1];
          sL_TypeOfService := L_StrList1.Strings[2];
          sL_PayFreeFlag := L_StrList1.Strings[3];
          sL_Result := sL_Result + 'C' + IntToHex(nL_LengthOfServiceName,2) + sL_ServiceName + sL_ExpirationDate + sL_TypeOfService + sL_PayFreeFlag;
        end
        else
          sL_Result :=  BAD_CMD_FORMAT;
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
    Memo2.Lines.Add('�e�X�����O:' + sI_FullCmd);

    bG_SendComm := true;
    //down, �ǰecommand
    try
    begin
      Memo4.Lines.Add('�e�X�����O: ' + sI_FullCmd);

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
              nL_Words := L_WStream.Read(L_ResponseAry, sizeof(L_ResponseAry)); //�o�ؼg�k�|�ϱo�ǰe�t���ܱo�ܧ�, �������t���ܺC
      //      nL_Words := recv(frmMain.G_SendCmdClntSocket.Socket.SocketHandle , G_DeviceData, sizeof(G_DeviceData),0);

              if nL_Words>0 then
              begin
                for ii:=0 to nL_Words-1 do
                begin
                  sL_RealData := sL_RealData + L_ResponseAry[ii];
                end;
                Memo4.Lines.Add('���쪺���O�^��: ' + sL_RealData);
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
              //down, ���s�s���� EMMG..
              reBuildNdsConn;
              //up, ���s�s���� EMMG..
            end;
          end;
        end;
//*******************
      except
        bG_buildSendCmdConnection := false;
        try
          frmMain.reBuildNdsConn;

        except

          frmMain.showLogOnMemo(SEND_CMD_ERROR_FLAG,'�ǰe���O' + sI_FullCmd + ' �� Nagra.�{���Y�N����,�M�᭫�s�Ұ�. '); //���槹����{����, �{���|������

        end;
        {
        buildSendCmdConnection(frmMain.G_SendCmdClntSocket, frmMain.G_SendCmdWStream);
        if not bG_buildSendCmdConnection then
        begin
          sL_LowLevelCmdID := Copy(sG_FullCommand,61,4);
          frmMain.showLogOnMemo(SEND_CMD_ERROR_FLAG,'�ǰe���O' + sL_LowLevelCmdID + ' �� Nagra.�{���Y�N����,�M�᭫�s�Ұ�. '); //���槹����{����, �{���|������
        end;
        }
      end;


    end
    except


      bG_SendComm := false;






    end;
    //up, �ǰecommand
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
//    showmessage('socket error..');
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
    //down, �إ� TcpClient �� array...
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
    //up, �إ� TcpClient �� array...

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
    sL_CardID : String;
begin
    //down, Card ID �� checksum ���͵{��...
    odd_sum := 0;
    even_sum := 0;
    fL_CardID := fI_CardID;
    for i:=1 to 11 do
    begin

      sL_CardID := FloatToStr(fL_CardID);
{
      nL_CardID := StrToInt(sL_CardID);
      digit := StrToInt(FloatToStr(nL_CardID mod 10));
}      
      digit := StrToInt(Copy(sL_CardID, length(sL_CardID), 1)) ;
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
    //down, Card ID �� checksum ���͵{��...
end;

procedure TfrmMain.BitBtn2Click(Sender: TObject);
var
    nL_CheckSum : Integer;
begin
    nL_CheckSum := getCardIdChecksum(StrToFloat(Edit1.Text));
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
    sI_AllSignatureKey ���i��� : 0,00000000~1,12345678
    �d�Ҫ�ܦ@���T�� signature key ���]�w
    from id �O 0 �� signature key �O 00000000
    from id �O 1 �� signature key �O 12345678
    from id �O 2 �� signature key �O 12345678
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
    sL_SubscriberInfo, sL_SubscriberIdInfo, sL_TmpReportbackDate : String;
    sL_AttachedResponseData, sL_ErrorCode, sL_ErrorMsg:String;
    ii,nL_Currency, nL_CmdFlagPos, nL_ActionLen, nL_PersonalRegionLen : Integer;
    sL_SubscriberID, sL_PersonalRegionLen, sL_PersonalRegion : String;
    nL_CompNdx, nL_Seed, nL_ReportbackDay : Integer;
    sL_ResponseDataStatusFlag, sL_CompId, sL_HighLevelCmdID, sL_DispatcherSequenceNo : String;
    sL_ReportbackAvailability, sL_ReportbackDay, sL_ReportbackTime, sL_Currency : String;
    sL_StatusCode, sL_Response, sL_CardID, sL_NewCardID, sL_RegionKey, sL_Date : String;

    sL_TmpOppvExpirationDate, sL_TmpOppvNum, sL_NumberOfOppv, sL_NumberOfContinuousService, sL_TmpServiceID : String;
    nL_NumberOfSlot, nL_NumberOfOppv, nL_NumberOfContinuousService : Integer;
    sL_NumberOfSlot, sL_F5ResponseData, sL_PopulationID, sL_EMailMsgID: String;

    sL_TmpMaxNumOfSlot, sL_TmpReportbackType, sL_TmpThreshold, sL_TmpLimit, sL_TmpSeriesID, sL_TmpSlotNumber : String;
begin
    {
    //down, for ���ե�..
    sL_CompId := '9';
    sL_HighLevelCmdID := 'A0';
    sL_DispatcherSequenceNo := '123';
    //up, for ���ե�..
    }

    getDispatcherSeqNo(true, sI_SequenceID, sL_CompId, sL_HighLevelCmdID, sL_DispatcherSequenceNo);
    if sL_CompId='' then
    begin
      memErrorLog.Lines.Add('�L�k�N���O�B�z���G�^���� Dispatcher.�]���L�k��X���O���ݪ��t�Υx.EMMG���O�Ǹ�:' + sI_SequenceID + ' --'+ DatetimeToStr(now));
      Exit;
    end;


    nL_CompNdx := G_RemoteInfoStrList.IndexOf(sL_CompId);
    if nL_CompNdx=-1 then
    begin
      memErrorLog.Lines.Add('�L�k�N���O�B�z���G�^���� Dispatcher.�]���L�k��X�t�Υx��������T.�t�Υx�N��:' + sL_CompId + ' --'+ DatetimeToStr(now));
      Exit;
    end;

    sL_ResponseDataStatusFlag := 'C';
    sL_ErrorCode := '';
    sL_ErrorMsg := '';
    nL_CmdFlagPos := 1;
    nL_ActionLen := length(sI_Action);
    sL_AttachedResponseData := '';
    sL_SubscriberIdInfo := '';
    sL_SubscriberInfo := '';
    sL_F5ResponseData := '';
    
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
        if sI_Action[nL_CmdFlagPos+1]='X' then //��ܱ����� 'Ex' �}�Y�� error code, �ҥH�n���� connection
        begin
          reBuildNdsConn;
          break;
        end
        else
        begin
//        break;
          continue;
        end;
      end
      else if sI_Action[nL_CmdFlagPos]='S' then //6.3.6
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
        sL_SubscriberIdInfo := sL_SubscriberID + SUBSCRIBER_INFO_SEP_FLAG + sL_StatusCode;

        continue;
      end
      else if sI_Action[nL_CmdFlagPos]='I' then //6.3.3
      begin
        nL_Seed := nL_CmdFlagPos+1;

        sL_CardID :=  Copy(sI_Action, nL_Seed, ICC_NO_LENGTH);
        nL_Seed := nL_Seed + ICC_NO_LENGTH;

        sL_NewCardID := Copy(sI_Action, nL_Seed, ICC_NO_LENGTH);
        nL_Seed := nL_Seed + ICC_NO_LENGTH;
        nL_Seed := nL_Seed + 18; //18�O��ӫO�d��쪺����

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
        nL_ReportbackDay := TUstr.hexToInt(sL_ReportbackDay);
        nL_Seed := nL_Seed + REPORTBACK_DATE_LENGTH;


        sL_ReportbackTime  := Copy(sI_Action, nL_Seed, REPORTBACK_TIME_LENGTH);
        nL_Seed := nL_Seed + REPORTBACK_TIME_LENGTH;

        sL_Currency  := Copy(sI_Action, nL_Seed, CURRENCY_LENGTH);
        nL_Currency := TUstr.hexToInt(sL_Currency);
        nL_Seed := nL_Seed + CURRENCY_LENGTH;



        nL_CmdFlagPos := nL_CmdFlagPos + RESPONSE_633_LENGTH;
//        continue;
        sL_SubscriberInfo := sL_CardID + SUBSCRIBER_INFO_SEP_FLAG + sL_NewCardID + SUBSCRIBER_INFO_SEP_FLAG;
        sL_SubscriberInfo := sL_SubscriberInfo + sL_Date + SUBSCRIBER_INFO_SEP_FLAG;
        sL_SubscriberInfo := sL_SubscriberInfo + sL_RegionKey + SUBSCRIBER_INFO_SEP_FLAG;
        sL_SubscriberInfo := sL_SubscriberInfo + sL_PersonalRegion + SUBSCRIBER_INFO_SEP_FLAG;
        sL_SubscriberInfo := sL_SubscriberInfo + sL_ReportbackAvailability + SUBSCRIBER_INFO_SEP_FLAG;
        sL_SubscriberInfo := sL_SubscriberInfo + IntToStr(nL_ReportbackDay) + SUBSCRIBER_INFO_SEP_FLAG;
        sL_SubscriberInfo := sL_SubscriberInfo + sL_ReportbackTime + SUBSCRIBER_INFO_SEP_FLAG;
        sL_SubscriberInfo := sL_SubscriberInfo + IntToStr(nL_Currency);

        continue;
      end
      else if sI_Action[nL_CmdFlagPos]='M' then //6.3.4
      begin
        nL_Seed := nL_CmdFlagPos+1;

        sL_EMailMsgID :=  Copy(sI_Action, nL_Seed, EMAIL_MSG_ID_LENGTH);
        sL_EMailMsgID := IntToStr(TUstr.hexToInt(sL_EMailMsgID));
        sL_F5ResponseData := sL_EMailMsgID;
        
        nL_Seed := nL_Seed + EMAIL_MSG_ID_LENGTH;
        nL_CmdFlagPos := nL_CmdFlagPos + RESPONSE_634_LENGTH;

        continue;
      end
      else if sI_Action[nL_CmdFlagPos]='T' then //6.3.7
      begin
        nL_Seed := nL_CmdFlagPos+1;

        sL_NumberOfOppv :=  Copy(sI_Action, nL_Seed, 4);
        nL_Seed := nL_Seed + 4;

        nL_NumberOfOppv := TUstr.hexToInt(sL_NumberOfOppv);
        for ii:=0 to nL_NumberOfOppv-1 do
        begin
          sL_TmpOppvNum :=  Copy(sI_Action, nL_Seed, HEX_OPPV_NUM_LENGTH);
          nL_Seed := nL_Seed + HEX_OPPV_NUM_LENGTH;
          sL_TmpOppvNum := IntToStr(TUstr.hexToInt(sL_TmpOppvNum));

          sL_TmpOppvExpirationDate :=  Copy(sI_Action, nL_Seed, OPPV_EXPIRATION_DATE_LENGTH);
          nL_Seed := nL_Seed + OPPV_EXPIRATION_DATE_LENGTH;

          if sL_F5ResponseData='' then
            sL_F5ResponseData := sL_TmpOppvNum + ',' + sL_TmpOppvExpirationDate
          else
            sL_F5ResponseData := sL_F5ResponseData + '~' + sL_TmpOppvNum + ',' + sL_TmpOppvExpirationDate;
        end;
        nL_CmdFlagPos := nL_CmdFlagPos + 5 + nL_NumberOfOppv*(HEX_OPPV_NUM_LENGTH+OPPV_EXPIRATION_DATE_LENGTH);

        continue;

      end
      else if sI_Action[nL_CmdFlagPos]='P' then //6.3.5
      begin
        nL_Seed := nL_CmdFlagPos+1;

        sL_NumberOfSlot :=  Copy(sI_Action, nL_Seed, 4);
        nL_Seed := nL_Seed + 4;

        nL_NumberOfSlot := TUstr.hexToInt(sL_NumberOfSlot);
        for ii:=0 to nL_NumberOfSlot-1 do
        begin
          sL_TmpSlotNumber :=  Copy(sI_Action, nL_Seed, HEX_SLOT_NUM_LENGTH);
          sL_TmpSlotNumber := IntToStr(TUstr.hexToInt(sL_TmpSlotNumber));
          nL_Seed := nL_Seed + HEX_SLOT_NUM_LENGTH;


          sL_TmpSeriesID :=  Copy(sI_Action, nL_Seed, HEX_SERIES_ID_LENGTH);
          sL_TmpSeriesID := IntToStr(TUstr.hexToInt(sL_TmpSeriesID));
          nL_Seed := nL_Seed + HEX_SERIES_ID_LENGTH;


          sL_TmpLimit :=  Copy(sI_Action, nL_Seed, HEX_LIMIT_LENGTH);
          sL_TmpLimit := IntToStr(TUstr.hexToInt(sL_TmpLimit));
          nL_Seed := nL_Seed + HEX_LIMIT_LENGTH;


          sL_TmpThreshold :=  Copy(sI_Action, nL_Seed, HEX_THRESHOLD_LENGTH);
          sL_TmpThreshold := IntToStr(TUstr.hexToInt(sL_TmpThreshold));
          nL_Seed := nL_Seed + HEX_THRESHOLD_LENGTH;


          sL_TmpReportbackType :=  Copy(sI_Action, nL_Seed, 1);
          nL_Seed := nL_Seed + 1;

          sL_TmpMaxNumOfSlot :=  Copy(sI_Action, nL_Seed, HEX_MAX_NUM_OF_SLOT_LENGTH);
          sL_TmpMaxNumOfSlot := IntToStr(TUstr.hexToInt(sL_TmpMaxNumOfSlot));
          nL_Seed := nL_Seed + HEX_MAX_NUM_OF_SLOT_LENGTH;
          

          if sL_F5ResponseData='' then
            sL_F5ResponseData := sL_TmpSlotNumber + ',' + sL_TmpSeriesID + ',' + sL_TmpLimit + ',' + sL_TmpThreshold + ',' + sL_TmpReportbackType + ',' + sL_TmpMaxNumOfSlot
          else
            sL_F5ResponseData := sL_F5ResponseData + '~' + sL_TmpSlotNumber + ',' + sL_TmpSeriesID + ',' + sL_TmpLimit + ',' + sL_TmpThreshold + ',' + sL_TmpReportbackType + ',' + sL_TmpMaxNumOfSlot;
        end;
        nL_CmdFlagPos := nL_CmdFlagPos + 5 + nL_NumberOfOppv*(HEX_OPPV_NUM_LENGTH+OPPV_EXPIRATION_DATE_LENGTH);

        continue;

      end
      else if sI_Action[nL_CmdFlagPos]='V' then //6.3.8
      begin
        nL_Seed := nL_CmdFlagPos+1;

        sL_NumberOfContinuousService :=  Copy(sI_Action, nL_Seed, 4);
        nL_Seed := nL_Seed + 4;

        nL_NumberOfContinuousService := TUstr.hexToInt(sL_NumberOfContinuousService);
        for ii:=0 to nL_NumberOfContinuousService-1 do
        begin

          sL_TmpServiceID :=  Copy(sI_Action, nL_Seed, HEX_SERVICE_ID_LENGTH);
          nL_Seed := nL_Seed + HEX_SERVICE_ID_LENGTH;

          sL_TmpServiceID := IntToStr(TUstr.hexToInt(sL_TmpServiceID));
          if sL_F5ResponseData='' then
            sL_F5ResponseData := sL_TmpServiceID
          else
            sL_F5ResponseData := sL_F5ResponseData + '~' + sL_TmpServiceID;
        end;
        nL_CmdFlagPos := nL_CmdFlagPos + 5 + nL_NumberOfContinuousService*HEX_SERVICE_ID_LENGTH;

        continue;
      end
      else if sI_Action[nL_CmdFlagPos]='U' then //6.3.9
      begin
        nL_Seed := nL_CmdFlagPos+1;

        sL_PopulationID :=  Copy(sI_Action, nL_Seed, HEX_POPULATION_ID_LENGTH);
        sL_PopulationID := IntToStr(TUstr.hexToInt(sL_PopulationID));
        sL_F5ResponseData := sL_PopulationID;

        nL_Seed := nL_Seed + HEX_POPULATION_ID_LENGTH;
        nL_CmdFlagPos := nL_CmdFlagPos + RESPONSE_639_LENGTH;

        continue;
      end
      else if sI_Action[nL_CmdFlagPos]='B' then //6.3.2
      begin
        nL_Seed := nL_CmdFlagPos+1;

        sL_TmpReportbackDate := Copy(sI_Action, nL_Seed, REPORTBACK_DATE_LENGTH);
        sL_TmpReportbackDate := IntToStr(TUstr.hexToInt(sL_TmpReportbackDate));

        nL_Seed := nL_Seed + REPORTBACK_DATE_LENGTH;
        nL_CmdFlagPos := nL_CmdFlagPos + RESPONSE_632_LENGTH;


        continue;
      end;

    end;

    //down, �զX�X�A���T���r��,�M��Ǧ^�� dispatcher..

    if (sL_HighLevelCmdID='A0') or (sL_HighLevelCmdID='A1') then
      sL_AttachedResponseData := sL_SubscriberIdInfo + SUBSCRIBER_DATA_SEP_FLAG + sL_SubscriberInfo
    else if (sL_HighLevelCmdID='A6') then
      sL_AttachedResponseData := sL_TmpReportbackDate
    else if (sL_HighLevelCmdID='F5') then
      sL_AttachedResponseData := sL_F5ResponseData
    else
      sL_AttachedResponseData := '';
      
    preSendToDispatcher(sL_CompID, sL_HighLevelCmdID, sL_DispatcherSequenceNo, sL_ResponseDataStatusFlag,sL_ErrorCode, sL_ErrorMsg, sL_AttachedResponseData);

    //up, �զX�X�A���T���r��,�M��Ǧ^�� dispatcher..
end;

procedure TfrmMain.Button3Click(Sender: TObject);
var
    sL_Action, sL_SequenceID : String;
begin
    sL_Action := 'S000000051I00000001006600000000000000000000000000000000000000987600541C0000000000000000000000000000E051938000003O';
    sL_SequenceID := '123';  //�H�N�񪺭�
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
      MessageDlg('�䤣����~�T����(' + sL_FileName + ')�{���Y�N����.', mtError, [mbOK],0);
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

procedure TfrmMain.closeSocket;
begin
    if Assigned(G_SendCmdClntSocket) and (G_SendCmdClntSocket.Active) then
      G_SendCmdClntSocket.Close;

end;

procedure TfrmMain.Button4Click(Sender: TObject);
var
    sL_ToRegionByteBin : String;

begin
    sL_ToRegionByteBin := '000000000000000000100';
    sL_ToRegionByteBin := Copy(sL_ToRegionByteBin, length(sL_ToRegionByteBin)-8+1, 8);
    showmessage(sL_ToRegionByteBin);
end;



end.




