//getHighLevelCommandInfo...



// 若用 Synchronize(updateCmdStatus), 則指令的處理速度會比較快(速度差好幾倍), 但 sometimes...程式會被 hold 住...
// SendCommandU 與 HandleResponseU 均有此 function...
unit SendCommandU;

interface

uses
  Classes, forms, comctrls, syncobjs, dbtables, Sysutils, Dialogs,
  ConstObjU, Ustru, UdateTimeu, scktcomp, winsock, inifiles, ADODB, IdGlobal;


type

  TSendCommand = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  public

    constructor Create;overload;
    function genRealIccNo(sI_SourceIccNo : String):String;
    function genRealStbNo(sI_SourceStbNo : String):String;
    function genRealImsproductID(sI_SourceImsProductID : String):String;        

//    constructor Create(I_MFrame: TfrmMonitor; I_ScrollBox: TScrollBox; I_ListItem: TListItem; I_AdoQryCmdLog:TADOQuery);overload;
    procedure onTimerStartAction;
    procedure buildSendCmdConnection(var I_ClientSocket : TClientSocket; var I_SendCmdWinSocketStream, I_ReceiveCmdWinSocketStream : TWinSocketStream);overload;
    procedure buildSendCmdConnection(var I_ClientSocket : TClientSocket; var I_SendCmdWinSocketStream : TWinSocketStream);overload;
    procedure buildSendCmdConnection(var I_ClientSocket : TClientSocket);overload;

// ===========================
    function getTransNum(nI_CompCode:Integer; sI_SeqNo : String):String;

    function getFormatIpAddr(sI_SrcIpAddr:String):String;
    procedure parseCommandData(I_CmdInfoObj:TCmdInfoObject);
    procedure preSendCommand(I_CmdInfoObj:TCmdInfoObject; sI_IccNo, sI_StbNo, sI_ImsProductID, sI_SubBeginDate, sI_SubEndDate :String);
    procedure makeSmsCommand(I_CmdInfoObj:TCmdInfoObject; sI_IccNo, sI_StbNo, sI_ImsProductID,sI_SubBeginDate, sI_SubEndDate:String; I_IrdInfoObj : TIrdInfoObj); //組合出SMS Command
    function getSmsCommandRootHeader(bI_GenTransactionNum:boolean; nI_CompCode, nI_CommandID : Integer; sI_CommandType, sI_SeqNo: String) : String;
    function getCommandSection(sI_CommandType, sI_BDate, sI_EDate,sI_SerialNo:String) : String;
    function getCommandType(nI_CommandID : Integer) : String;
    function getCommandBody(I_CmdInfoObj:TCmdInfoObject; nI_CommandID : Integer; I_IrdInfoObj : TIrdInfoObj; sI_SubBeginDate, sI_SubEndDate, sI_ImsProductID : String) : String;
    function getRecSeqNo(nI_CmdType: integer): string;
    function transIrdInfo(nI_CompCode:Integer; sI_MisTrdCmdID, sI_MisIrdCmdData: String; var  L_IrdInfoObj : TIrdInfoObj ):boolean;
//    function getBinaryVal(nI_Value: Integer; nI_Digits: Integer; bI_Reverse: boolean):TTempRecord;
    function Write(I_Buf: Pointer; I_Count: Integer; I_WStream: TWinSocketStream):Integer;
    function Increment(Ptr: Pointer; Increment: Integer): Pointer;

    procedure SendCommand(sI_CommandNo : STring);

    procedure updateCmdStatus;
    function IsDataOK(sI_ImsProductID, sI_IccNo, sI_StbNo :String; var nI_ResultFlag:Integer) : boolean;
    procedure getHighLevelCommandInfo(I_CmdInfoObj:TCmdInfoObject);overload;


    function getDetailCommandIDs(sI_HighLevelCommand:String):String;

    procedure writeSendCmdLog(sI_CommandNo : STring);overload;
    procedure writeTimeStamp;
    procedure receiveResponse;    
  end;

implementation

uses frmMainU, frmSysParamU, dtmMainU, HandleUIU;


{ Important: Methods and properties of objects in VCL can only be used in a
  method called using Synchronize, for example,

      Synchronize(UpdateCaption);

  and UpdateCaption could look like,

    procedure TSendCommand.UpdateCaption;
    begin
      Form1.Caption := 'Updated in a thread';
    end; }

{ TSendCommand }



procedure TSendCommand.Execute;
var
    ii : Integer;
    sL_SeqNo, sL_TimeStr : String;
    nL_Count : Integer;
    L_CmdInfoObj : TCmdInfoObject;
    wL_Hour, wL_Min, wL_Sec, sL_MSec : word;
begin

    self.FreeOnTerminate := true;
    self.OnTerminate := frmMain.SendCommandThreadDone;
    frmMain.nG_SC_RealCmdSeqNo := 0;
    frmMain.nG_SC_MaxRealCmdSeqNo := StrToIntDef(TUstr.AddString('','9',true,SEQ_NUM_LENGTH),999999);


{
    onTimerStartAction;
    while true do
    begin
      DecodeTime(now, wL_Hour, wL_Min, wL_Sec, sL_MSec);
      sL_TimeStr := Format('%.2d%.2d', [wL_Hour, wL_Min]);
      if sL_TimeStr=DELETE_MAPPING_DATA_TIME then
        dtmMain.deleteMappingData;
      if frmMain.bG_RunSendingCommandThread then
      begin
        if frmMain.G_CommandInfoStrList.Count >0 then
        begin
          try
            L_CmdInfoObj := (frmMain.G_CommandInfoStrList.Objects[0] as TCmdInfoObject);
            parseCommandData(L_CmdInfoObj);
            frmMain.G_CommandInfoStrList.Delete(0);
          except

          end;

        end;
      end;

    end;
}

end;



function TSendCommand.getCommandBody(I_CmdInfoObj:TCmdInfoObject;nI_CommandID: Integer;
I_IrdInfoObj : TIrdInfoObj; sI_SubBeginDate, sI_SubEndDate, sI_ImsProductID : String): String;
var
    sL_Result, sL_CommandID : String;
    sL_StuNumber, sL_ImsProductID : String;
    sL_BDate, sL_EDate, sL_ZipCode, sL_CreditMode, sL_Credit : String;
    sL_ThresholdCredit, sL_EventName, sL_EventNameLength, sL_Price, sL_CCNum1 : String;
    sL_IpAddr, sL_CCPort, sL_CbDate, sL_CbTime : String;
    sL_CallFreq, sL_DateFirstCall, sL_CreditLimit : String;
    sL_PhoneNum1, sL_PhoneNum2, sL_PhoneNum3, sL_SrcTransNum : String;
    sL_IrdCmdID, sL_IrdOperation, sL_IrdDataLength, sL_IrdData : String;
begin
    {
    此 function 用來組合出各指令的 parameters.
    ex: EMM command 的參數請參考 SMS GATEWAY SPECIFICATION 第五章
    }
    sL_CommandID := TUstr.AddString(IntToStr(nI_CommandID),'0',false,COMMAND_ID_LENGTH);


    sL_StuNumber := genRealStbNo(Trim(I_CmdInfoObj.STB_NO));

    sL_ImsProductID := sI_ImsProductID;

    sL_BDate := TUstr.replaceStr(sI_SubBeginDate,'/','');
    sL_EDate := TUstr.replaceStr(sI_SubEndDate,'/','');



    sL_ZipCode := I_CmdInfoObj.ZIP_CODE;

    sL_CreditMode := Format('%.2d', [StrToIntDef(I_CmdInfoObj.CREDIT_MODE,0)]);
    {
      sL_CreditMode :
        01 => ADD
        02 => SUBTRACT
        03 => SET CREDIT
        04 => SET BALANCE
        05 => SUB OFFSET
    }
    sL_Credit := Format('%.7d', [StrToIntDef(I_CmdInfoObj.CREDIT,0)]);

    sL_ThresholdCredit := Format('%.7d', [StrToIntDef(I_CmdInfoObj.THRESHOLD_CREDIT,0)]);
    sL_EventName := Uppercase(TUstr.AddString(I_CmdInfoObj.EVENT_NAME,' ',true,EVENT_NAME_LENGTH));
    sL_EventNameLength := Format('%.2d',[length(I_CmdInfoObj.EVENT_NAME)]);

    sL_Price := Format('%.5d', [StrToIntDef(I_CmdInfoObj.PRICE,0)]);

    sL_IpAddr := getFormatIpAddr(I_CmdInfoObj.IP_ADDR);
    sL_CCPort := Format('%.5d', [StrToIntDef(I_CmdInfoObj.CC_PORT,0)]);
    sL_CbDate := I_CmdInfoObj.CALLBACK_DATE ;
    sL_CbTime := I_CmdInfoObj.CALLBACK_TIME;
    sL_CallFreq := I_CmdInfoObj.CALLBACK_FREQUENCY;
    sL_DateFirstCall := I_CmdInfoObj.FIRST_CALLBACK_DATE;
    sL_CreditLimit := Format('%.7d', [StrToIntDef(I_CmdInfoObj.CREDIT_LIMIT,0)]);

    if I_IrdInfoObj<>nil then
    begin
      sL_IrdCmdID := I_IrdInfoObj.s_IrdCmdID;
      sL_IrdOperation := I_IrdInfoObj.s_IrdOperation;
      sL_IrdDataLength := I_IrdInfoObj.s_IrdDataLength;
      sL_IrdData := I_IrdInfoObj.s_IrdData;
    end
    else
    begin
      sL_IrdCmdID := '';
      sL_IrdOperation := '';
      sL_IrdDataLength := '';
      sL_IrdData := '';

    end;

    //down, 將 IRD_DATA 用0補滿96個bytes
    sL_IrdData := TUstr.AddString(sL_IrdData,'0',true,IRD_COMMAND_DATA_LENGTH);

    //down, 電話號碼的轉換 => 這樣的程式寫法,對simulator 無法 work; 但連線到瑞士測試時,卻又ok.
    sL_CCNum1 := TUstr.AddString(I_CmdInfoObj.CC_NUMBER_1,chr(32),true,CALL_COLLECTOR_PHONE_NUM_LENGTH);

    sL_PhoneNum1 := TUstr.AddString(I_CmdInfoObj.PHONE_NUM_1,chr(32),true,AUTHORIZED_PHONE_NUM_LENGTH);
    sL_PhoneNum2 := TUstr.AddString(I_CmdInfoObj.PHONE_NUM_2,chr(32),true,AUTHORIZED_PHONE_NUM_LENGTH);
    sL_PhoneNum3 := TUstr.AddString(I_CmdInfoObj.PHONE_NUM_3,chr(32),true,AUTHORIZED_PHONE_NUM_LENGTH);
    //up, 電話號碼的轉換 => 這樣的程式寫法,對simulator 無法 work; 但連線到瑞士測試時,卻又ok.

    case nI_CommandID of
      2 :
        sL_Result := sL_CommandID + sL_ImsProductID + sL_BDate + sL_EDate;
      3 :
        sL_Result := sL_CommandID + sL_ImsProductID + sL_EDate;
      4,5,6 :
        sL_Result := sL_CommandID + sL_ImsProductID;
      7,14,15,20,21,50,51, 53, 62,71,105,110,111,1002 :
        sL_Result := sL_CommandID;
      8 :
        sL_Result := sL_CommandID + sL_CreditMode + sL_Credit;
      9 :
        sL_Result := sL_CommandID + sL_ThresholdCredit;
      10 :
        sL_Result := sL_CommandID + sL_ImsProductID + sL_EventNameLength + sL_EventName + sL_Price ;
      13 :
        sL_Result := sL_CommandID + sL_Credit + sL_ThresholdCredit;
      48 :
        sL_Result := sL_CommandID + sL_ZipCode;
      49 : //此指令有問題
        sL_Result := sL_CommandID + sL_CCNum1;
      52, 104 :
        sL_Result := sL_CommandID + sL_StuNumber;
      54 :
        sL_Result := sL_CommandID + sL_IpAddr + sL_CCPort;
      60 :
        sL_Result := sL_CommandID + sL_CbDate + sL_CbTime;
      61 :
        sL_Result := sL_CommandID + sL_CallFreq + sL_DateFirstCall + sL_CbTime;
      69 :
        sL_Result := sL_CommandID + sL_IrdCmdID + sL_IrdOperation + sL_IrdDataLength + sL_IrdData;
      100 :
        sL_Result := sL_CommandID + sL_CreditLimit;
      101 :  //此指令有問題
        sL_Result := sL_CommandID + sL_PhoneNum1 + sL_PhoneNum2 + sL_PhoneNum3;
      1000 :
        sL_Result := sL_CommandID + sL_SrcTransNum + TUstr.AddString('','0',true,IMS_PRODUCT_ID_LENGTH)  + TUstr.AddString('','0',true,IMS_PRODUCT_ID_LENGTH);
    end;


    result := sL_Result;
end;

function TSendCommand.getSmsCommandRootHeader(bI_GenTransactionNum:boolean; nI_CompCode, nI_CommandID : Integer; sI_CommandType, sI_SeqNo : String): String;
var
    sL_Result : String;
begin


    if bI_GenTransactionNum then
      frmMain.sG_TransactionNum := getTransNum(nI_CompCode, sI_SeqNo)
    else
      frmMain.sG_TransactionNum := '';
    sL_Result := frmMain.sG_TransactionNum + sI_CommandType;
    sL_Result := sL_Result + frmMain.sG_SourceID + frmMain.sG_DestID + frmMain.sG_MopPPID;
    sL_Result := sL_Result + TUdateTime.GetPureDateStr(date,'');
    result := sL_Result;
end;

function TSendCommand.getCommandType(nI_CommandID: Integer): String;
var
    sL_Result : String;
begin

    if (EMM_CommandBeginID<=nI_CommandID)and(nI_CommandID <=EMM_CommandEndID) then
      sL_Result := frmMain.sG_EmmCommandType
    else if (CONTROL_CommandBeginID<=nI_CommandID)and(nI_CommandID <=CONTROL_CommandEndID) then
      sL_Result := frmMain.sG_ControlCommandType
    else if (FEEDBACK_CommandBeginID<=nI_CommandID)and(nI_CommandID <=FEEDBACK_CommandEndID) then
      sL_Result := frmMain.sG_FeedbackCommandType
    else if (PRODUCT_CommandBeginID<=nI_CommandID)and(nI_CommandID <=PRODUCT_CommandEndID) then
      sL_Result := frmMain.sG_ProductCommandType
    else if (OPERATION_CommandBeginID<=nI_CommandID)and(nI_CommandID <=OPERATION_CommandEndID) then
      sL_Result := frmMain.sG_OperationCommandType
    else
      sL_Result := '-1';
    result := sL_Result;
end;

function TSendCommand.getCommandSection(sI_CommandType,sI_BDate, sI_EDate,sI_SerialNo:String): String;
var
    sL_Result : String;
    sL_CurDateStr : String;
    sL_BroadastMode, sL_AddressType: String ;
    L_SysParam : TfrmSysParam ;
begin

    if (sI_CommandType= frmMain.sG_EmmCommandType) then
    begin  //SMS Gateway Interface definition, P18
      sL_Result := frmMain.sG_EmmBroadcastMode + sI_BDate + sI_EDate + frmMain.sG_EmmAddressType;
      if (frmMain.sG_EmmAddressType=EMM_UNIQUE_ADDRESS_TYPE) then
        sL_Result := sL_Result + sI_SerialNo;
    end
    else if (sI_CommandType= frmMain.sG_ControlCommandType) then
    begin //SMS Gateway Interface definition, P20
      sL_CurDateStr := TUdateTime.GetPureDateStr(date,'');
      L_SysParam := TfrmSysParam.Create(Application);
      L_SysParam.getControlCommandSectionParam(sL_BroadastMode, sL_AddressType);
      L_SysParam.Free;
      sL_Result := sL_BroadastMode + sL_CurDateStr + sL_CurDateStr + sL_AddressType + sI_SerialNo; //SMS Gateway Interface definition, P20

    end
    else if (sI_CommandType= frmMain.sG_ProductCommandType) then
    begin
      sL_Result := ''; //SMS Gateway Interface definition, P21
    end
    else if (sI_CommandType= frmMain.sG_FeedbackCommandType) then
    begin
      sL_Result := sI_SerialNo; //SMS Gateway Interface definition, P21
    end
    else if (sI_CommandType= frmMain.sG_OperationCommandType) then
    begin
      sL_Result := ''; //SMS Gateway Interface definition, P21
    end;


    result := sL_Result;
end;

procedure TSendCommand.makeSmsCommand(I_CmdInfoObj:TCmdInfoObject; sI_IccNo, sI_StbNo, sI_ImsProductID , sI_SubBeginDate, sI_SubEndDate :String; I_IrdInfoObj : TIrdInfoObj);
var
  {
  F1: TextFile;
  Request : TRequestLib;
  Header: string;  //Request Header;
  Body: string;    //Request
  Footer: string;  //Request UAs
  }
  nL_CompCode : Integer;
  sL_SmsCommandRootHeader, sL_CommandSection : String;
  sL_CommandBody, sL_CommandType : String;
  sL_IccNo, sL_CurActiveLowLevelCmd : String;
begin
    try
      nL_CompCode := StrToIntDef(I_CmdInfoObj.COMPCODE,0);
      sL_IccNo := sI_IccNo;
      frmMain.bG_SC_MakeSmsCommand := false;

      //  sG_RecSeqNo := GetRecSeqNo(1);

      //Analyze SEND_NAGRA Command


      sL_CurActiveLowLevelCmd := frmMain.sG_CurActiveLowLevelCmd;
      sL_CommandType := getCommandType(StrToInt(sL_CurActiveLowLevelCmd));

      sL_SmsCommandRootHeader := getSmsCommandRootHeader(true, nL_CompCode, StrToInt(sL_CurActiveLowLevelCmd), sL_CommandType, I_CmdInfoObj.SEQNO);

  //    showmessage(sL_SmsCommandRootHeader);


      sL_CommandSection := getCommandSection(sL_CommandType, frmMain.sG_BroadcastStartDate, frmMain.sG_BroadcastEndDate,sL_IccNo);

      sL_CommandBody := getCommandBody(I_CmdInfoObj, StrToInt(sL_CurActiveLowLevelCmd), I_IrdInfoObj, sI_SubBeginDate, sI_SubEndDate, sI_ImsProductID);

      frmMain.sG_SC_FullCommand := sL_SmsCommandRootHeader + sL_CommandSection +sL_CommandBody;

  //    showmessage('EmmCommandRootHeader:'+sL_CommandRootHeader)
  //    showmessage(sL_FullCommand);
  //  end;



    except

      frmMain.bG_SC_MakeSmsCommand := false;
      Exit;
    end;
    frmMain.bG_SC_MakeSmsCommand := true;

end;

function TSendCommand.getRecSeqNo(nI_CmdType: integer): string;
var

    sL_SQL, sL_TmpStr: string;
begin
    sL_TmpStr:='';
    {
    with G_AdoQryRealCmd do
    begin
      SQL.Clear;
      case nI_CmdType of
        1:
        sL_SQL := 'select max(REQ_NO) as MaxNo from ERRORLOG where Log_Date=' + STR_SEP + TUdateTime.GetPureDateStr(date,'')+ STR_SEP;
      end;
      Close;
      SQL.Add(sL_SQL);
      Open;
      First;



      if fieldbyname('MaxNo').asstring = '' then
        sL_TmpStr := TUdateTime.GetPureRocDateStr(date,'') + '00000001'
      else
      begin
        sL_TmpStr := inttostr(strtoint64(fieldbyname('MaxNo').asstring)+ 1);
        sL_TmpStr := copy('000000000000000',1,15-length(sL_TmpStr))+sL_TmpStr;
      end;
      Close;
    end;
    }
    result:=sL_TmpStr;
end;



procedure TSendCommand.buildSendCmdConnection(var I_ClientSocket : TClientSocket; var I_SendCmdWinSocketStream, I_ReceiveCmdWinSocketStream : TWinSocketStream);
var
    nL_Words : Integer;
    L_TmpRecord : TTempRecord;
    ii, nL_CmdLen : Integer;
    L_Socket : TClientSocket;
    L_WStream: TWinSocketStream;
    L_DeviceCall : TDeviceCall;
    L_DeviceStatus : TDeviceStatus;
    L_DeviceAnswer : TDeviceAnswer;
    sL_CmdObjName : String;
begin
    frmMain.bG_SC_buildSendCmdConnection := false;
    if Assigned(I_ClientSocket) then
    begin
      I_ClientSocket := nil;
      I_ClientSocket.Free;
    end;
    I_ClientSocket:= TClientSocket.Create(nil);
    I_ClientSocket.ClientType := ctBlocking;


//    L_Socket := G_MonitorFrame.ClientSocket1;
    L_Socket := I_ClientSocket;
    try
    begin
      with L_Socket do
      begin
        Address := frmMain.getSendCommandServerIP;
        Port := frmMain.getSendCommandPort;

        Open;
        L_DeviceCall.op_mode := char(0);

        sL_CmdObjName := 'SMSGW';

        //L_DeviceCall.obj_name :=
        FillChar(L_DeviceCall.obj_name, length(L_DeviceCall.obj_name), char(0)); { initialize the buffer }

        for ii:=1 to length(sL_CmdObjName) do
        begin
          L_DeviceCall.obj_name[ii-1] := sL_CmdObjName[ii];
        end;

//        StrPCopy(L_DeviceCall.obj_name,sL_CmdObjName);
        L_DeviceCall.obj_name_len := char(length(sL_CmdObjName));


        //1002
        {
        sL_CommandNo := '1002';//直接送出指令 1002, 用以確認 connection 是否順利建立 => 欲接收 feedback command 的 process ,才需要送出 command 1002
//        sL_CommandNo := sG_CommandNo;

        sL_CommandType := getCommandType(StrToInt(sL_CommandNo));
        sL_SmsCommandRootHeader := getSmsCommandRootHeader('', StrToInt(sL_CommandNo), sL_CommandType);
  //    showmessage(sL_SmsCommandRootHeader);
        sL_CurrentDate := TUdateTime.GetPureDateStr(date,'');
        sL_BoxNo := '';
        sL_CommandRootHeader := getCommandSection(sL_CommandType, sL_CurrentDate, sL_CurrentDate,sG_IccNo);
        sL_CommandBody := getCommandBody(StrToInt(sL_CommandNo), sL_BoxNo);
        sL_FullCommand := sL_SmsCommandRootHeader + sL_CommandRootHeader + sL_CommandBody;
        }

//        FillChar(L_DeviceCall.user_data, length(L_DeviceCall.user_data), char(0)); //=> 用此行..會有錯誤...

        for ii:=1 to length(L_DeviceCall.user_data) do
          L_DeviceCall.user_data[ii-1] := char(0);

        {
        for ii:=1 to length(sL_FullCommand) do
        begin
          L_DeviceCall.user_data[ii-1] := sL_FullCommand[ii];
        end;
        }
//        nL_CmdLen := 1 + 1 + length(sL_CmdObjName) + length(sL_FullCommand); //ok.. => 其實不確定要不要送 sL_CmdObjName
        nL_CmdLen := 1 + 1 + length(sL_CmdObjName); //ok.. => 其實不確定要不要送 sL_CmdObjName
//        nL_CmdLen := 1+1+ length(sL_FullCommand); //testing..

        L_TmpRecord := frmMain.getBinaryVal(nL_CmdLen,2,false);
        L_DeviceCall.len[0] := L_TmpRecord.TempAry[0];
        L_DeviceCall.len[1] := L_TmpRecord.TempAry[1];

//        L_WStream := TWinSocketStream.Create(L_Socket.Socket, nG_Timer*100); //howard
        I_ReceiveCmdWinSocketStream := TWinSocketStream.Create(L_Socket.Socket, 60000);

        L_WStream := TWinSocketStream.Create(L_Socket.Socket, 60000);
        I_SendCmdWinSocketStream := L_WStream;


        nL_Words := write(@L_DeviceCall, nL_CmdLen+2, L_WStream); //build connection
//            showmessage('return:'+inttostr(nL_Words));




        nL_Words := recv(L_Socket.Socket.SocketHandle , L_DeviceStatus, 3,0);
  //      showmessage('recv 1 的傳回值:'+inttostr(nL_Words));

        for ii:=1 to length(L_DeviceAnswer.data) do
          L_DeviceAnswer.data[ii-1] := char(0);

//        nL_Words := L_WStream.Read(L_DeviceAnswer, L_DeviceStatus.len);
        nL_Words := recv(L_Socket.Socket.SocketHandle , L_DeviceAnswer, L_DeviceStatus.len,0);

  //      showmessage('recv 2 的傳回值:'+inttostr(nL_Words));

//        showmessage('device_call L_DeviceAnswer.code==' + L_DeviceAnswer.code);
//        showmessage('device_call L_DeviceAnswer.data==' + L_DeviceAnswer.data);


        frmMain.bG_HasBuildSendCmdConnection := true;
        frmMain.bG_SC_buildSendCmdConnection := true;
      end;
    end;
    except

      frmMain.sG_SC_ErrorCode := BUILD_CONNECTION_ERROR;
      frmMain.sG_SC_ErrorMsg := BUILD_CONNECTION_ERROR_MSG;

      frmMain.sG_SC_CmdStatus := 'E';
      frmMain.nG_SC_ConnectionID := frmMain.nG_CurActiveConn;
      updateCmdStatus;
//      Synchronize(updateCmdStatus);
      {
      try
        buildSendCmdConnection(I_ClientSocket, I_WinSocketStream);
      except
        frmMain.showLogOnMemo(RE_CONNECT_NAGRA_ERROR_FLAG,'重新連線到 Nagra');
      end;
      }
      raise;
      exit;
    end;


end;
procedure TSendCommand.buildSendCmdConnection(var I_ClientSocket : TClientSocket; var I_SendCmdWinSocketStream: TWinSocketStream);
var
    nL_Words : Integer;
    L_TmpRecord : TTempRecord;
    ii, nL_CmdLen : Integer;
    L_Socket : TClientSocket;
    L_WStream: TWinSocketStream;
    L_DeviceCall : TDeviceCall;
    L_DeviceStatus : TDeviceStatus;
    L_DeviceAnswer : TDeviceAnswer;
    sL_CmdObjName : String;
begin
    frmMain.bG_SC_buildSendCmdConnection := false;
    if Assigned(I_ClientSocket) then
    begin
      I_ClientSocket := nil;
      I_ClientSocket.Free;
    end;
    I_ClientSocket:= TClientSocket.Create(nil);
    I_ClientSocket.ClientType := ctBlocking;


//    L_Socket := G_MonitorFrame.ClientSocket1;
    L_Socket := I_ClientSocket;
    try
    begin
      with L_Socket do
      begin
        Address := frmMain.getSendCommandServerIP;
        Port := frmMain.getSendCommandPort;

        Open;
        L_DeviceCall.op_mode := char(0);

        sL_CmdObjName := 'SMSGW';

        //L_DeviceCall.obj_name :=
        FillChar(L_DeviceCall.obj_name, length(L_DeviceCall.obj_name), char(0)); { initialize the buffer }
        
        for ii:=1 to length(sL_CmdObjName) do
        begin
          L_DeviceCall.obj_name[ii-1] := sL_CmdObjName[ii];
        end;

//        StrPCopy(L_DeviceCall.obj_name,sL_CmdObjName);
        L_DeviceCall.obj_name_len := char(length(sL_CmdObjName));


        //1002
        {
        sL_CommandNo := '1002';//直接送出指令 1002, 用以確認 connection 是否順利建立 => 欲接收 feedback command 的 process ,才需要送出 command 1002
//        sL_CommandNo := sG_CommandNo;

        sL_CommandType := getCommandType(StrToInt(sL_CommandNo));
        sL_SmsCommandRootHeader := getSmsCommandRootHeader('', StrToInt(sL_CommandNo), sL_CommandType);
  //    showmessage(sL_SmsCommandRootHeader);
        sL_CurrentDate := TUdateTime.GetPureDateStr(date,'');
        sL_BoxNo := '';
        sL_CommandRootHeader := getCommandSection(sL_CommandType, sL_CurrentDate, sL_CurrentDate,sG_IccNo);
        sL_CommandBody := getCommandBody(StrToInt(sL_CommandNo), sL_BoxNo);
        sL_FullCommand := sL_SmsCommandRootHeader + sL_CommandRootHeader + sL_CommandBody;
        }

//        FillChar(L_DeviceCall.user_data, length(L_DeviceCall.user_data), char(0)); //=> 用此行..會有錯誤...

        for ii:=1 to length(L_DeviceCall.user_data) do
          L_DeviceCall.user_data[ii-1] := char(0);

        {
        for ii:=1 to length(sL_FullCommand) do
        begin
          L_DeviceCall.user_data[ii-1] := sL_FullCommand[ii];
        end;
        }
//        nL_CmdLen := 1 + 1 + length(sL_CmdObjName) + length(sL_FullCommand); //ok.. => 其實不確定要不要送 sL_CmdObjName
        nL_CmdLen := 1 + 1 + length(sL_CmdObjName); //ok.. => 其實不確定要不要送 sL_CmdObjName
//        nL_CmdLen := 1+1+ length(sL_FullCommand); //testing..

        L_TmpRecord := frmMain.getBinaryVal(nL_CmdLen,2,false);
        L_DeviceCall.len[0] := L_TmpRecord.TempAry[0];
        L_DeviceCall.len[1] := L_TmpRecord.TempAry[1];

//        L_WStream := TWinSocketStream.Create(L_Socket.Socket, nG_Timer*100); //howard


        L_WStream := TWinSocketStream.Create(L_Socket.Socket, 60000);
        I_SendCmdWinSocketStream := L_WStream;


        nL_Words := write(@L_DeviceCall, nL_CmdLen+2, L_WStream); //build connection
//            showmessage('return:'+inttostr(nL_Words));



        nL_Words := recv(L_Socket.Socket.SocketHandle , L_DeviceStatus, 3,0);
  //      showmessage('recv 1 的傳回值:'+inttostr(nL_Words));

        for ii:=1 to length(L_DeviceAnswer.data) do
          L_DeviceAnswer.data[ii-1] := char(0);

        nL_Words := recv(L_Socket.Socket.SocketHandle , L_DeviceAnswer, L_DeviceStatus.len,0);
  //      nL_Words := recv(L_Socket.Socket.SocketHandle , L_DeviceStatus, 3,0);
  //      showmessage('recv 2 的傳回值:'+inttostr(nL_Words));

//        showmessage('device_call L_DeviceAnswer.code==' + L_DeviceAnswer.code);
//        showmessage('device_call L_DeviceAnswer.data==' + L_DeviceAnswer.data);


        frmMain.bG_HasBuildSendCmdConnection := true;
        frmMain.bG_SC_buildSendCmdConnection := true;
      end;
    end;
    except

      frmMain.sG_SC_ErrorCode := BUILD_CONNECTION_ERROR;
      frmMain.sG_SC_ErrorMsg := BUILD_CONNECTION_ERROR_MSG;

      frmMain.sG_SC_CmdStatus := 'E';
      frmMain.nG_SC_ConnectionID := frmMain.nG_CurActiveConn;
      updateCmdStatus;
//      Synchronize(updateCmdStatus);
      {
      try
        buildSendCmdConnection(I_ClientSocket, I_WinSocketStream);
      except
        frmMain.showLogOnMemo(RE_CONNECT_NAGRA_ERROR_FLAG,'重新連線到 Nagra');
      end;
      }
      raise;
      exit;
    end;


end;


function TSendCommand.Write(I_Buf: Pointer; I_Count: Integer;
  I_WStream: TWinSocketStream): Integer;
var
    nL_Idx: Integer;
begin
    nL_Idx := 0;
    while (nL_Idx < I_Count) do
      nL_Idx := nL_Idx + I_WStream.Write(Increment(I_Buf, nL_Idx)^, I_Count-nL_Idx);
    {
    try
      nL_Idx := nL_Idx + I_WStream.Write(Increment(I_Buf, nL_Idx)^, I_Count-nL_Idx);
    except
      buildSendCmdConnection(frmMain.SendCmdClntSocket, frmMain.G_SendCmdWStream);
      Write(I_Buf, I_Count,frmMain.G_SendCmdWStream);
  //    raise;
    end;
    }
    result := nL_Idx;
end;

function TSendCommand.Increment(Ptr: Pointer;
  Increment: Integer): Pointer;
begin
    Result := Pointer(Longint(Ptr) + Increment);
end;


procedure TSendCommand.SendCommand(sI_CommandNo : STring);
var
    nL_Words : Integer;
    ii, nL_Len : Integer;
    L_TmpRecord : TTempRecord;
    L_DeviceData : TDeviceData;
    L_WStream: TWinSocketStream;
    L_Socket : TClientSocket;
    L_StrList:TStringList;
    sL_TransactionNum, sL_LowLevelCmdID, sL_TmpFullCommand : String;
begin

    if frmMain.sG_SC_FullCommand='' then Exit;
//frmMain.Memo1.Lines.Add(sG_FullCommand);
    frmMain.bG_SC_SendComm := true;
    //down, 傳送command
    try
    begin

      sL_TmpFullCommand := frmMain.sG_SC_FullCommand;
      frmMain.sG_SC_TransNumForUpdateCmdStatus := Copy(sL_TmpFullCommand,1,SEND_TRANSACTION_NUM_LENGTH);
      sL_TransactionNum := Copy(sL_TmpFullCommand,1, SEND_TRANSACTION_NUM_LENGTH);
                  
//      sG_FullCommand := '';
      {
      G_MonitorFrame.Gauge1.Progress := 100;
      G_MonitorFrame.chkStatus2.Checked := true;
      }
//L_StrList:=TStringList.Create;
      nL_Len := length(sL_TmpFullCommand);//ok..
//      nL_Len := 1 + 1 + length('EMGR_EME_SVR') + length(sL_FullCommand); //testing..=> 其實不確定要不要送 sL_CmdObjName

      L_TmpRecord := frmMain.GetBinaryVal(nL_Len,2,false);
      L_DeviceData.len[0] := L_TmpRecord.TempAry[0];
      L_DeviceData.len[1] := L_TmpRecord.TempAry[1];

//L_StrList.Add('a');

      FillChar(L_DeviceData.data, length(L_DeviceData.data), char(0)); { initialize the buffer }
      {
      for ii:=1 to length(L_DeviceData.data) do
        L_DeviceData.data[ii-1] := char(0);
      }
//      frmMain.memErrorLog.Lines.Add('==' + sL_TmpFullCommand + '==');
      
      for ii:=1 to length(sL_TmpFullCommand) do
      begin
        L_DeviceData.data[ii-1] := sL_TmpFullCommand[ii];
      end;


      L_Socket := frmMain.G_SendCmdClntSocket;

//      L_WStream := frmMain.G_SendCmdWStream;
      L_WStream := TWinSocketStream.Create(frmMain.G_SendCmdClntSocket.Socket , 60000);
//L_StrList.Add('b');
      //down, 原先希望無法連上 server 時,可以再試著連一次看看,但好像無法達到這樣的效果
      {
      if (not L_WStream.WaitForData(2000)) or (not L_Socket.Active) then
        buildSendCmdConnection(L_Socket, L_WStream);
      }
      //up, 原先希望無法連上 server 時,可以再試著連一次看看,但好像無法達到這樣的效果
      {
      L_Socket := G_MonitorFrame.ClientSocket1;
      L_WStream := G_MonitorFrame.G_WStream;
      }
//L_StrList.Add('c');

      //down, 將資料寫到 send comand log 中
//      Synchronize(writeSendCmdLog);
      writeSendCmdLog(sI_CommandNo);
      //up, 將資料寫到 send comand log 中


      try
        if frmMain.bG_HasBuildSendCmdConnection  and Assigned(L_WStream) then
        begin
          nL_Words := L_WStream.Write(L_DeviceData, nL_Len+2); //0723
//          nL_Words := write(@L_DeviceData, nL_Len+2, L_WStream); //send_data..

//        nL_Words := G_WinSocketStreamForSendCmd.Write(L_DeviceData, nL_Len+2); //0317

//        nL_Words := write(@L_DeviceData, nL_Len+2, L_WStream); //send_data..   //0317
        end;
      except
        if Assigned(L_WStream) then
          L_WStream.Free;
        frmMain.bG_SC_buildSendCmdConnection := false;
        sL_LowLevelCmdID := Copy(sL_TmpFullCommand,61,4);
        frmMain.showLogOnMemo(SEND_CMD_ERROR_FLAG,'傳送指令' + sL_LowLevelCmdID + ' 給 Nagra.程式即將結束,然後重新啟動. '); //執行完此行程式後, 程式會結束掉
        Exit;
{
        try
          frmMain.reBuildNagraConn;

        except
          sL_LowLevelCmdID := Copy(sL_TmpFullCommand,61,4);
          frmMain.showLogOnMemo(SEND_CMD_ERROR_FLAG,'傳送指令' + sL_LowLevelCmdID + ' 給 Nagra.程式即將結束,然後重新啟動. '); //執行完此行程式後, 程式會結束掉

        end;
}
        {
        buildSendCmdConnection(frmMain.G_SendCmdClntSocket, frmMain.G_SendCmdWStream);
        if not bG_buildSendCmdConnection then
        begin
          sL_LowLevelCmdID := Copy(sL_FullCommand,61,4);
          frmMain.showLogOnMemo(SEND_CMD_ERROR_FLAG,'傳送指令' + sL_LowLevelCmdID + ' 給 Nagra.程式即將結束,然後重新啟動. '); //執行完此行程式後, 程式會結束掉
        end;
        }
      end;
      if Assigned(L_WStream) then
        L_WStream.Free;
      
//***************
//      receiveResponse; //接收回應 ...2003/11/03 將此 mark 掉
//***************

//L_StrList.Add('d');
    end
    except
      if Assigned(L_WStream) then
        L_WStream.Free;


      frmMain.bG_SC_SendComm := false;

      if frmMain.bG_ShowUI then
        frmMain.G_HandleUIThread.updateListItemStatus(nil,4,sL_TransactionNum,frmMain.sG_SC_ErrorCode,frmMain.sG_SC_ErrorMsg);



      frmMain.sG_SC_ErrorCode := SEND_COMMAND_ERROR;
      frmMain.sG_SC_ErrorMsg := SEND_COMMAND_ERROR_MSG;
      frmMain.sG_SC_CmdStatus := 'E';




      frmMain.nG_SC_ConnectionID := frmMain.nG_CurActiveConn;
//L_StrList.Add('error' + IntToStr(nG_ConnectionID) + '==' + sG_SeqNo);

//L_StrList.SaveToFile('c:\aaaaa.txt');
//L_StrList.Free;

      updateCmdStatus;

//        Synchronize(updateCmdStatus);


      Exit;




    end;
    //up, 傳送command
//L_StrList.Add('e');
    if frmMain.bG_ShowUI then
      frmMain.G_HandleUIThread.updateListItemStatus(nil,1,sL_TransactionNum,'','');


end;

function TSendCommand.getFormatIpAddr(sI_SrcIpAddr: String): String;
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

function TSendCommand.getDetailCommandIDs(
  sI_HighLevelCommand: String): String;
var
    L_IniFile : TIniFile;
    sL_Result : String;
    sL_ExeFileName, sL_ExePath, sL_IniFileName : STring;  
begin
  sL_ExeFileName := Application.ExeName;
  sL_ExePath := ExtractFileDir(sL_ExeFileName);
    if frmMain.sG_NewIniFileName='' then
      sL_IniFileName := sL_ExePath + '\' + CABLE_NAGRA_INI_FILE_NAME
    else
      sL_IniFileName := sL_ExePath + '\' + frmMain.sG_NewIniFileName;



  if not FileExists(sL_IniFileName) then
  begin
    MessageDlg(frmMain.getLanguageSettingInfo('SendCmd_Thread_Msg_1_Content'),mtWarning,[mbOK],0);
    sL_Result := '';
  end
  else
  begin
    L_IniFile := TIniFile.Create(sL_IniFileName);
    sL_Result:= L_IniFile.ReadString('COMMANDMATRIX',sI_HighLevelCommand,'');

    L_IniFile.Free;
  end;

  result := sL_Result;
end;

procedure TSendCommand.getHighLevelCommandInfo(I_CmdInfoObj:TCmdInfoObject);
var
    sL_ImsProductID : String;
    L_StrList : TStringList;
begin
//L_StrList := TStringList.Create;
//      sG_CommandNo := I_CmdInfoObj.HIGH_LEVEL_CMD_ID;
//frmMain.G_SendCommandThread.sG_CommandNo := I_CmdInfoObj.HIGH_LEVEL_CMD_ID;

{
2003/10/31
      sG_IccNo := Copy(I_CmdInfoObj.ICC_NO,1,10);
      sG_IccNo := TUstr.AddString(sG_IccNo,'0',false,ICC_CARDNO_LENGTH);//10 碼,右靠, 左補0
}
//L_StrList.Add(sG_IccNo);

//      sG_StbNo := I_CmdInfoObj.STB_NO;
//L_StrList.Add('1=' + I_CmdInfoObj.STB_NO);

{
2003/10/31
      sG_StbNo := Copy(I_CmdInfoObj.STB_NO,1,10);
      sG_StbNo := TUstr.AddString(sG_StbNo,'0',false,BOXNO_ID_LENGTH-4);//14 碼,右靠, 左補0,右補4個空白
//L_StrList.Add('2=' + sG_StbNo);
      if frmMain.bG_Simulator then
        sG_StbNo := TUstr.AddString(sG_StbNo,'0',true,BOXNO_ID_LENGTH)//14 碼,右靠, 左補0,右補4個0 => for 模擬器 => 例如使用者輸入 65536, 則轉成 00000655360000
      else
        sG_StbNo := TUstr.AddString(sG_StbNo,' ',true,BOXNO_ID_LENGTH);//14 碼,右靠, 左補0,右補4個空白 => for production
}
//L_StrList.Add(sG_StbNo);
//L_StrList.SaveToFile('c:\aabc.txt');
//L_StrList.Free;

{
2003/11/03
      if I_CmdInfoObj.SUBSCRIPTION_BEGIN_DATE<>'' then
      begin
        sG_SubBeginDate := TUdateTime.GetPureDateStr(StrToDate(I_CmdInfoObj.SUBSCRIPTION_BEGIN_DATE),'');
        sG_MisSubBeginDate := sG_SubBeginDate;
      end;

      if I_CmdInfoObj.SUBSCRIPTION_END_DATE<>'' then
      begin
        sG_SubEndDate := TUdateTime.GetPureDateStr(StrToDate(I_CmdInfoObj.SUBSCRIPTION_END_DATE),'');
        sG_MisSubEndDate := sG_SubEndDate;
      end;
}
//      sG_Notes := I_CmdInfoObj.NOTES;
      {
      if frmMain.G_NotesHasValueStrList.IndexOf(sG_CommandNo)<>-1 then
      begin
        sG_Notes := dtmMain.getNotes(frmMain.nG_CurActiveConn, sG_SeqNo);
      end;
      }
//      sG_Option := fieldbyname('OPTION').asstring;

{
2003/11/03
      sG_CommandStatus := I_CmdInfoObj.CMD_STATUS;
      sG_Operator := I_CmdInfoObj.OPERATOR;
      sG_ErrorCode := '';
      sG_ErrorMsg := '';


      sL_ImsProductID := genRealImsproductID(I_CmdInfoObj.IMS_PRODUCT_ID);


      sG_ZipCode := I_CmdInfoObj.ZIP_CODE;
      if sG_ZipCode<>'' then
        sG_ZipCode := TUstr.AddString(sG_ZipCode,'0',false,5);


      nG_CreditMode := StrToIntDef(I_CmdInfoObj.CREDIT_MODE,0);
      nG_Credit := StrToIntDef(I_CmdInfoObj.CREDIT,0);
      nG_CreditLimit := StrToIntDef(I_CmdInfoObj.CREDIT_LIMIT,0);


      nG_ThresholdCredit := StrToIntDef(I_CmdInfoObj.THRESHOLD_CREDIT,0);
      sG_EventName := I_CmdInfoObj.EVENT_NAME;
      nG_Price := StrToIntDef(I_CmdInfoObj.PRICE,0);
      sG_CCNum1 := I_CmdInfoObj.CC_NUMBER_1;
      sG_IpAddr := I_CmdInfoObj.IP_ADDR;

      nG_CCPort := StrToIntDef(I_CmdInfoObj.CC_PORT,0);
      sG_CallbackDate := I_CmdInfoObj.CALLBACK_DATE;
      sG_CallbackTime := I_CmdInfoObj.CALLBACK_TIME;
      sG_CallFreq :=  I_CmdInfoObj.CALLBACK_FREQUENCY;

      if I_CmdInfoObj.FIRST_CALLBACK_DATE<>'' then
      begin
        sG_FirstCallbackDate := TUdateTime.GetPureDateStr(StrToDate(I_CmdInfoObj.FIRST_CALLBACK_DATE),'/');
        sG_FirstCallbackDate := TUstr.replaceStr(sG_FirstCallbackDate,'/','');
      end;
      sG_PhoneNum1 :=  I_CmdInfoObj.PHONE_NUM_1;
      sG_PhoneNum2 :=  I_CmdInfoObj.PHONE_NUM_2;
      sG_PhoneNum3 :=  I_CmdInfoObj.PHONE_NUM_3;
}
//      sG_MisIrdCmdID :=  I_CmdInfoObj.MIS_IRD_CMD_ID;
//      sG_MisIrdCmdData :=  I_CmdInfoObj.MIS_IRD_CMD_DATA;

//      nG_CompCode :=  StrToIntDef(I_CmdInfoObj.COMPCODE,0);

{
0726
      if sG_MisIrdCmdID <> '' then
      begin
        sG_MisIrdCmdID := TUstr.AddString(I_CmdInfoObj.MIS_IRD_CMD_ID,'0',false,3);
        transIrdInfo;
      end;
}

end;


procedure TSendCommand.preSendCommand(I_CmdInfoObj:TCmdInfoObject; sI_IccNo, sI_StbNo, sI_ImsProductID, sI_SubBeginDate, sI_SubEndDate :String);
var
    L_IrdInfoObj : TIrdInfoObj;
    sL_CommandNo, sL_SeqNo, sL_MisIrdCmdData, sL_MisIrdCmdID : String;
    bL_TransIrdOk : boolean;
    nL_CompCode : Integer;
begin
    Application.ProcessMessages;
    bL_TransIrdOk := true;
    sL_SeqNo := I_CmdInfoObj.SEQNO;
    nL_CompCode := StrToIntDef(I_CmdInfoObj.COMPCODE,0);

    sL_CommandNo := I_CmdInfoObj.HIGH_LEVEL_CMD_ID ;
    sL_MisIrdCmdID := I_CmdInfoObj.MIS_IRD_CMD_ID;

    if sL_MisIrdCmdID <> '' then
    begin
      sL_MisIrdCmdID := TUstr.AddString(sL_MisIrdCmdID,'0',false,3);
      sL_MisIrdCmdData := I_CmdInfoObj.MIS_IRD_CMD_DATA;
      L_IrdInfoObj := TIrdInfoObj.Create;
      bL_TransIrdOk := transIrdInfo(nL_CompCode, sL_MisIrdCmdID, sL_MisIrdCmdData, L_IrdInfoObj);
    end;


    if bL_TransIrdOk then
    begin
  //      CRS.Enter;

  //      Synchronize(makeSmsCommand);       //產生SMS Command

      makeSmsCommand(I_CmdInfoObj, sI_IccNo, sI_StbNo, sI_ImsProductID,sI_SubBeginDate, sI_SubEndDate,  L_IrdInfoObj);       //產生SMS Command

      if frmMain.bG_ShowUI then
        frmMain.G_HandleUIThread.createListItem(frmMain.sG_TransactionNum,I_CmdInfoObj.HIGH_LEVEL_CMD_ID, I_CmdInfoObj.OPERATOR, frmMain.sG_CurTimeStamp);



      if frmMain.bG_SC_MakeSmsCommand then
      begin

      //    sleep(3000);//停頓3 secs

  //        Synchronize(SendCommand);
  //        for jj:=0 to 100 do
          Application.ProcessMessages;
          SendCommand(sL_CommandNo);
          Application.ProcessMessages;
                    
          if Assigned(L_IrdInfoObj) then
          begin
            L_IrdInfoObj.Free;
            L_IrdInfoObj := nil;
          end;  
  //        showmessage('123');
      end
      else
      begin


  //      CRS.Leave;
        exit;
      end;
    end
    else
    begin
      frmMain.sG_SC_ErrorCode := TRANS_IRD_ERROR;
      frmMain.sG_SC_ErrorMsg := TRANS_IRD_ERROR_MSG;

      frmMain.sG_SC_CmdStatus := 'E';
      frmMain.nG_SC_ConnectionID := frmMain.nG_CurActiveConn;
      dtmMain.UpdateCMD_Status_By_Seq(frmMain.nG_SC_ConnectionID,sL_SeqNo,frmMain.sG_SC_CmdStatus, frmMain.sG_SC_ErrorCode, frmMain.sG_SC_ErrorMsg);

    end;

end;

function TSendCommand.IsDataOK(sI_ImsProductID, sI_IccNo, sI_StbNo:String; var nI_ResultFlag:Integer): boolean;
begin
    nI_ResultFlag := 0;
    if (length(sI_StbNo) <> 0) and (length(sI_StbNo) <> BOXNO_ID_LENGTH) then
    begin
      nI_ResultFlag := -1;
      result := false;
      exit;
    end;
    if (length(sI_IccNo) <> 0) and  (length(sI_IccNo) <> ICC_CARDNO_LENGTH) then
    begin
      nI_ResultFlag := -2;
      result := false;
      exit;
    end;

    if (length(sI_ImsProductID) <> 0) and  (length(sI_ImsProductID) <> IMS_PRODUCT_ID_LENGTH) then
    begin
      nI_ResultFlag := -3;
      result := false;
      exit;
    end;
    result := true;

end;

function TSendCommand.transIrdInfo(nI_CompCode : Integer; sI_MisTrdCmdID, sI_MisIrdCmdData: String; var  L_IrdInfoObj : TIrdInfoObj ):boolean;
var
    bL_Result : boolean;
    sL_RealMessage, sL_TmpData1,sL_TmpData : String;
    L_StrList:TStringList;
    nL_AscIIValue,ii : Integer;
    sL_RealMailId, sL_BinValue, sL_RealPriority, sL_RealSegmentNumber, sL_RealTotalSegment : String;

    sL_IrdCmdID, sL_IrdOperation, sL_IrdDataLength, sL_IrdData : String;    
begin
    //以下設定請參考文件 Hitron Technologies StbCakIrdSpe010302.pdf 以及 Nagravision S.A. SMS Gateway intergace definition V02.06.02
    //但這兩份文件與模擬軟體的回應似乎不太一致...

    bL_Result := true;        
    sL_IrdData := sI_MisIrdCmdData;
    case StrToInt(sI_MisTrdCmdID) of
      8: //(設定 IPPV 訂購密碼)
       begin
        sL_IrdCmdID := '018';
        sL_IrdOperation :=  '002';
//        sG_IrdDataLength :=  '00';
        sL_IrdDataLength :=  '04';
        sL_IrdData := '';
        if sI_MisIrdCmdData='' then
          sI_MisIrdCmdData := '0000';//預設之 IPPV 訂購密碼
        for ii:=1 to length(sI_MisIrdCmdData) do
        begin
          nL_AscIIValue := Ord(sI_MisIrdCmdData[ii]);//取出每個 byte 的 ASCII
          sL_IrdData := sL_IrdData + IntToHex(nL_AscIIValue,0); //取出該 ASCII 的16進位值
        end;
       end;
      1: //(設定親子密碼)
       begin
        sL_IrdCmdID := '018';
        sL_IrdOperation :=  '001';
//        sG_IrdDataLength :=  '00';
        sL_IrdDataLength :=  '04';
        sL_IrdData := '';
        if sI_MisIrdCmdData='' then
          sI_MisIrdCmdData := '0000';//預設之PIN Code
        for ii:=1 to length(sI_MisIrdCmdData) do
        begin
          nL_AscIIValue := Ord(sI_MisIrdCmdData[ii]);//取出每個 byte 的 ASCII
          sL_IrdData := sL_IrdData + IntToHex(nL_AscIIValue,0); //取出該 ASCII 的16進位值
        end;
       end;
      2: //(傳送訊息)
       begin
         sL_IrdCmdID := '192';
         sL_IrdOperation :=  '001';

         sL_BinValue := IntToBin(frmMain.nG_SC_MailId);
         sL_RealMailId := Copy(sL_BinValue,length(sL_BinValue)-10+1,10);

         sL_BinValue := IntToBin(frmMain.nG_SC_TotalSeqment);
         sL_RealTotalSegment := Copy(sL_BinValue,length(sL_BinValue)-6+1,6);

         sL_BinValue := IntToBin(StrToIntDef(sL_IrdData, 1)); //0=>normal priority; 1 => high priority; 2=>emergency
         sL_RealPriority :=Copy(sL_BinValue,length(sL_BinValue)-2+1,2);



         sL_BinValue := IntToBin(frmMain.nG_SC_SegmentNumber);
         sL_RealSegmentNumber := Copy(sL_BinValue,length(sL_BinValue)-6+1,6);

         sL_TmpData := sL_RealMailId + sL_RealTotalSegment + sL_RealPriority + sL_RealSegmentNumber;

         sL_TmpData1 := '';
         sL_TmpData1 := intToHex(TUstr.binToInt(Copy(sL_TmpData,1,8)),2);
         sL_TmpData1 := sL_TmpData1 + intToHex(TUstr.binToInt(Copy(sL_TmpData,9,8)),2);
         sL_TmpData1 := sL_TmpData1 + intToHex(TUstr.binToInt(Copy(sL_TmpData,17,8)),2);


//
//         sG_ActiveMessage
        sL_RealMessage := '';
        for ii:=1 to length(frmMain.sG_SC_ActiveMessage) do
        begin
          nL_AscIIValue := Ord(frmMain.sG_SC_ActiveMessage[ii]);//取出每個 byte 的 ASCII
          sL_RealMessage := sL_RealMessage + IntToHex(nL_AscIIValue,0); //取出該 ASCII 的16進位值
        end;

//         sG_IrdData := sL_RealMailId + sL_RealTotalSegment + sL_RealPriority + sL_RealSegmentNumber + sL_RealMessage;
         sL_IrdData := sL_TmpData1 + sL_RealMessage;
         sL_IrdDataLength :=  TUstr.AddString(FloatToStr(length(sL_IrdData)/2),'0',false,2);

       end;
      3: //(Force Tune)
       begin
        sL_IrdCmdID := '193';
        sL_IrdOperation :=  '001';
        sL_IrdDataLength :=  TUstr.AddString(IntToStr(length(sL_IrdData)),'0',false,2);
       end;
      4: //(set network id)
       begin
        sL_IrdCmdID := NETWORK_OPERATION_194; //194=>operation 要為'', 且 data length=2;
//        sG_IrdCmdID := NETWORK_OPERATION_195; //195=>operation表示多久要 reset STB, data length=4
        sL_IrdOperation :=  '001';
        sL_IrdData := frmMain.getCompNetworkInfo(nI_CompCode,sL_IrdCmdID);
        sL_IrdDataLength :=  TUstr.AddString(FloatToStr(length(sL_IrdData)/2),'0',false,2);
        if length(sL_IrdData)=0 then
          bL_Result := false; //表示抓不到 network id 之資料...
       end;
      5: //Master/Slave continuous mode initialisation => 參考文件檔案:030527 IRD Master_Slave Implementation_Guideline v0.0.3.pdf
       begin
        sL_IrdCmdID := '199';
        sL_IrdOperation :=  '001';

        L_StrList := TUstr.ParseStrings(sL_IrdData,'~');
        if L_StrList.Count =3 then
        begin
          {
          sG_IrdData 的格式 => 7~2~24
          L_StrList.Strings[0] => Validation Period
          L_StrList.Strings[1] => Random Period
          L_StrList.Strings[2] => timeout
          }
          sL_IrdData := frmMain.sG_SC_MasterIccNo + L_StrList.Strings[0] + L_StrList.Strings[1] + L_StrList.Strings[2];
          sL_IrdDataLength :=  TUstr.AddString(FloatToStr(length(sL_IrdData)/2),'0',false,2);
        end
        else
        begin
          //表示 sG_IrdData 之格式錯誤...
          sL_IrdData := '';
          sL_IrdDataLength :=  TUstr.AddString(FloatToStr(0),'0',false,2);
        end;
       end;
      6: //Master/Slave cancellation => 參考文件檔案:030527 IRD Master_Slave Implementation_Guideline v0.0.3.pdf
       begin
        sL_IrdCmdID := '199';
        sL_IrdOperation :=  '002';
        sL_IrdData := '';
        sL_IrdDataLength :=  TUstr.AddString(FloatToStr(0),'0',false,2);

       end;
      7: //Master/Slave single shot => 參考文件檔案:030527 IRD Master_Slave Implementation_Guideline v0.0.3.pdf
       begin
        sL_IrdCmdID := '199';
        sL_IrdOperation :=  '003';
        L_StrList := TUstr.ParseStrings(sL_IrdData,'~');
        if L_StrList.Count =2 then
        begin
          {
          sG_IrdData 的格式 => 1160052795~24
          L_StrList.Strings[0] => Master SmartCard ID
          L_StrList.Strings[1] => timeout
          }
          sL_IrdData := L_StrList.Strings[0] + L_StrList.Strings[1];
          sL_IrdDataLength :=  TUstr.AddString(FloatToStr(length(sL_IrdData)/2),'0',false,2);
        end
        else
        begin
          sL_IrdData := '';
          sL_IrdDataLength :=  TUstr.AddString(FloatToStr(0),'0',false,2);
        end;
       end;

    end;
    L_IrdInfoObj.s_IrdCmdID := sL_IrdCmdID;
    L_IrdInfoObj.s_IrdOperation := sL_IrdOperation;
    L_IrdInfoObj.s_IrdData := sL_IrdData;
    L_IrdInfoObj.s_IrdDataLength := sL_IrdDataLength;

    result := bL_Result;

end;


procedure TSendCommand.parseCommandData(I_CmdInfoObj:TCmdInfoObject);
var

    nL_SegmentCount, nL_Mod, nL_Div : Integer;
    sL_Notes, sL_CommandNo, sL_FullMessage : String;
    sL_DetailCommandIDs, sL_ProductInfo : String;
    L_StrList, L_NotesStrList, L_ProductsStrList, L_ProductInfoStrList,  L_CardStrList  : TStringList;
    ii, jj, kk : Integer;
    bL_ProfuctFieldFormatError : boolean;
    nL_ResultFlag  : Integer;
    sL_IccNo, sL_StbNo, sL_ImsProductID, sL_CurActiveLowLevelCmd : String;
    sL_SubBeginDate, sL_SubEndDate : String;    
begin

//    sG_SeqNo := I_CmdInfoObj.SEQNO;

    sL_CommandNo := I_CmdInfoObj.HIGH_LEVEL_CMD_ID ;



    sL_IccNo := genRealIccNo(I_CmdInfoObj.ICC_NO);
    sL_StbNo := genRealStbNo(I_CmdInfoObj.STB_NO);
   
    sL_ImsProductID := genRealImsproductID(I_CmdInfoObj.IMS_PRODUCT_ID);


    try
//    getHighLevelCommandInfo(I_CmdInfoObj);
//    G_CmdInfoObj := I_CmdInfoObj;
//    getHighLevelCommandInfo;


    sL_DetailCommandIDs := getDetailCommandIDs(sL_CommandNo);
    if sL_DetailCommandIDs='' then Exit;
    if not IsDataOK(sL_ImsProductID, sL_IccNo, sL_StbNo, nL_ResultFlag) then
    begin
      case nL_ResultFlag of
        -1 :
         begin
           frmMain.sG_SC_ErrorCode:= STB_NO_LENGTH_ERROR;
           frmMain.sG_SC_ErrorMsg := STB_NO_LENGTH_ERROR_MSG;
         end;
        -2 :
         begin
           frmMain.sG_SC_ErrorCode:= ICC_CARD_NO_LENGTH_ERROR;
           frmMain.sG_SC_ErrorMsg := ICC_CARD_NO_LENGTH_ERROR_MSG;
         end;
        -3 :
         begin
           frmMain.sG_SC_ErrorCode:= IMS_PRODUCT_ID_LENGTH_ERROR;
           frmMain.sG_SC_ErrorMsg := IMS_PRODUCT_ID_LENGTH_ERROR_MSG;
         end;
        else
         begin
           frmMain.sG_SC_ErrorCode:= OTHER_ERROR;
           frmMain.sG_SC_ErrorMsg := OTHER_ERROR_MSG;
         end;
      end;


      try
        frmMain.sG_SC_CmdStatus := 'E';
        frmMain.nG_SC_ConnectionID := frmMain.nG_CurActiveConn;

        frmMain.sG_SC_TransNumForUpdateCmdStatus := Copy(frmMain.sG_SC_FullCommand,1,9);
        updateCmdStatus;
//        Synchronize(updateCmdStatus);

      except

      end;


      Exit;
    end;
//    sG_TransactionNum := getTransNum; //0227
    frmMain.sG_CurTimeStamp := DateTimeToStr(now);
    try
      writeTimeStamp;
    except

    end;


    L_StrList := TUStr.ParseStrings(sL_DetailCommandIDs,',');

    sL_SubBeginDate :=  TUStr.replaceStr(I_CmdInfoObj.SUBSCRIPTION_BEGIN_DATE,'/','');
    sL_SubEndDate := TUStr.replaceStr(I_CmdInfoObj.SUBSCRIPTION_END_DATE,'/','');

    for ii:=0 to L_StrList.Count -1 do
    begin
      frmMain.sG_CurActiveLowLevelCmd := L_StrList.Strings[ii];
      sL_CurActiveLowLevelCmd := L_StrList.Strings[ii];
      if (sL_CommandNo='A1') and (frmMain.sG_CurActiveLowLevelCmd='48') then//執行開卡之set ZipCode 指令時,要做此判斷
      begin
        if (not frmMain.bG_SetZipCode) or (I_CmdInfoObj.ZIP_CODE='')  then
        begin
          continue;
        end;
      end;

      if (sL_CurActiveLowLevelCmd='100') then//執行Redefine Credit Limit 指令時,要做此判斷
      begin
        if StrToIntDef(I_CmdInfoObj.CREDIT_LIMIT,0) =0 then
          frmMain.nG_SC_CreditLimit := frmMain.nG_DefaultCreditLimit;
      end;

      bL_ProfuctFieldFormatError := false;
      if (sL_CommandNo='B7') then //(收視權限複製)
      begin
        //update send_nagra set notes='111111111111^20020801^20020901,222222222222^20020901^20021231~3333333333;4444444444';
        sL_Notes := I_CmdInfoObj.NOTES;
        {
        傳入多卡號,多頻道(欄位Notes)=>將所有頻道 grant 給所有卡號.
        各 product 間以','區隔. 各卡號間以';'區隔.product data與卡號 data以'~'區隔.(先product data,後卡號 data)
        各 product 的info 包含三個部分=>product id, 收視起日,收視迄日.而這三個部分以'^'作為區隔
        Ex: 111111111111^20020801^20020901,222222222222^20020901^20021231~3333333333;4444444444
        }
        L_NotesStrList := TUStr.ParseStrings(sL_Notes,'~');
        if L_NotesStrList.Count = 2 then
        begin
          L_ProductsStrList := TUStr.ParseStrings(L_NotesStrList.Strings[0],',');
          L_CardStrList := TUStr.ParseStrings(L_NotesStrList.Strings[1],';');
          if (L_ProductsStrList= nil ) or (L_ProductsStrList.Count = 0) or
             (L_CardStrList= nil ) or (L_CardStrList.Count = 0) then
            bL_ProfuctFieldFormatError := true;
        end
        else
          bL_ProfuctFieldFormatError := true;

        if bL_ProfuctFieldFormatError then
        begin
          frmMain.sG_SC_ErrorCode:= NOTES_FIELD_FORMAT_ERROR;
          frmMain.sG_SC_ErrorMsg := NOTES_FIELD_FORMAT_ERROR_MSG;
          try
            frmMain.sG_SC_CmdStatus := 'E';
            frmMain.nG_SC_ConnectionID := frmMain.nG_CurActiveConn;

            frmMain.sG_SC_TransNumForUpdateCmdStatus := Copy(frmMain.sG_SC_FullCommand,1,9);
            updateCmdStatus;
//            Synchronize(updateCmdStatus);

          except

          end;
          break;
        end
        else
        begin
          for jj :=0 to L_ProductsStrList.Count -1 do
          begin
            L_ProductInfoStrList := TUStr.ParseStrings(L_ProductsStrList.Strings[jj],'^');
            if (L_ProductInfoStrList<>nil) and (L_ProductInfoStrList.Count=3) then
            begin
              sL_ImsProductID := genRealImsproductID(L_ProductInfoStrList.Strings[0]);


              sL_SubBeginDate := TUstr.replaceStr(L_ProductInfoStrList.Strings[1],'/','');
              sL_SubEndDate := TUstr.replaceStr(L_ProductInfoStrList.Strings[2],'/','');
              for kk:=0 to L_CardStrList.Count -1 do
              begin
                sL_IccNo := genRealIccNo(L_CardStrList.Strings[kk]);

                preSendCommand(I_CmdInfoObj,  sL_IccNo, sL_StbNo, sL_ImsProductID,sL_SubBeginDate, sL_SubEndDate);
//                Synchronize(preSendCommand); //高階指令 B7 比較特殊,要傳送很多次指令,所以由此呼叫 function
              end
            end
            else
            begin
              frmMain.sG_SC_ErrorCode:= NOTES_FIELD_FORMAT_ERROR;
              frmMain.sG_SC_ErrorMsg := NOTES_FIELD_FORMAT_ERROR_MSG;
              try
                frmMain.sG_SC_CmdStatus := 'E';
                frmMain.nG_SC_ConnectionID := frmMain.nG_CurActiveConn;

                frmMain.sG_SC_TransNumForUpdateCmdStatus := Copy(frmMain.sG_SC_FullCommand,1,9);
                updateCmdStatus;
//                Synchronize(updateCmdStatus);

              except

              end;

//              continue; =>用 continue 會有問題
              break;
            end;
          end;

        end;
      end
      else if (sL_CommandNo='C1') then // (解除配對)
      begin
         {
         各ICC卡號間間以','區隔
         }
        //update send_nagra set notes='1111111111,2222222222,3333333333';
        sL_Notes := I_CmdInfoObj.NOTES;
        L_NotesStrList := TUStr.ParseStrings(sL_Notes,',');
        for jj:=0 to L_NotesStrList.Count -1 do//表示有這麼多組 ICC 卡號資料
        begin
          sL_StbNo := genRealStbNo('');
          sL_IccNo := genRealIccNo(L_NotesStrList.Strings[jj]);
          preSendCommand(I_CmdInfoObj, sL_IccNo, sL_StbNo, sL_ImsProductID, sL_SubBeginDate, sL_SubEndDate);
//          Synchronize(preSendCommand);
        end;
      end
      else if (sL_CommandNo='B2') or (sL_CommandNo='B4') or (sL_CommandNo='B5') then //(B2:關頻道-多組, B4:暫停頻道-多組, B5:恢復頻道-多組)
      begin
         {
         V各 product 間以','區隔
         }
        //update send_nagra set notes='111111111111,222222222222,333333333333,555555555555';
        sL_Notes := I_CmdInfoObj.NOTES;
        L_NotesStrList := TUStr.ParseStrings(sL_Notes,',');
        for jj:=0 to L_NotesStrList.Count -1 do//表示有這麼多組 product 資料
        begin
          sL_ImsProductID := genRealImsproductID(L_NotesStrList.Strings[jj]);
          preSendCommand(I_CmdInfoObj, sL_IccNo, sL_StbNo, sL_ImsProductID, sL_SubBeginDate, sL_SubEndDate);
//          Synchronize(preSendCommand);
        end;
      end
      else if (sL_CommandNo='E2')then //(E2:傳送訊息)
      begin
        sL_FullMessage := I_CmdInfoObj.NOTES;

        Inc(frmMain.nG_SC_MailId);
        nL_Mod := length(sL_FullMessage) mod MAX_MESSAGE_LENGTH;
        nL_Div := length(sL_FullMessage) div MAX_MESSAGE_LENGTH;

        if nL_Mod=0 then //表示訊息長度剛好是  MAX_MESSAGE_LENGTH 的整數倍
          nL_SegmentCount := nL_Div
        else
          nL_SegmentCount := nL_Div + 1;

        frmMain.nG_SC_TotalSeqment := nL_SegmentCount;

        for jj:=0 to nL_SegmentCount -1 do
        begin
          frmMain.nG_SC_SegmentNumber := jj+1;
          frmMain.sG_SC_ActiveMessage := Copy(sL_FullMessage,jj*MAX_MESSAGE_LENGTH+1, MAX_MESSAGE_LENGTH);


          preSendCommand(I_CmdInfoObj, sL_IccNo, sL_StbNo, sL_ImsProductID, sL_SubBeginDate, sL_SubEndDate);
        end;


      end
      else if (sL_CommandNo='E5')then //(E5:Master/Slave continuous mode initialisation)
      begin
        frmMain.sG_SC_MasterIccNo := sL_StbNo;
         {
         V各 slave ICC  間以','區隔
         }
        //update send_nagra set notes='1160092498,1160092569';
        sL_Notes := I_CmdInfoObj.NOTES;
        L_NotesStrList := TUStr.ParseStrings(sL_Notes,',');
        for jj:=0 to L_NotesStrList.Count -1 do//表示有這麼多組 slave ICC 資料
        begin

          sL_StbNo := genRealStbNo(L_NotesStrList.Strings[jj]);
          sL_IccNo := genRealIccNo(L_NotesStrList.Strings[jj]);

          preSendCommand(I_CmdInfoObj, sL_IccNo, sL_StbNo, sL_ImsProductID, sL_SubBeginDate, sL_SubEndDate);
//          Synchronize(preSendCommand);
        end;
      end
      else if (sL_CommandNo='B6') then //(延展期限)
      begin
        //update send_nagra set notes='A~111111111111,A~222222222222,A~333333333333,B~555555555555~20021031';
        sL_Notes := I_CmdInfoObj.NOTES;
        {
        V各 product 間以','區隔, product與期間之間以'~'區隔;
        第一碼是'A'者,表示共用此指令之subscription_end_datee,
        第一碼是'B'者,表示自己有自己的 subscription_end_date.
        Ex:A~1,A~2,A~3,B~5~20020205=>product 1,2,3以指令之 subscription_end_date為日期;
        而product 5 的 subscription_end_date 為20020205
        }
        L_NotesStrList := TUStr.ParseStrings(sL_Notes,',');
        for jj:=0 to L_NotesStrList.Count -1 do//表示有這麼多組 product 資料
        begin
        
          sL_ProductInfo := L_NotesStrList.Strings[jj];
          L_ProductsStrList := TUStr.ParseStrings(sL_ProductInfo,'~');
          if (L_ProductsStrList.Strings[0]='A') and (L_ProductsStrList.Count=2) then //ex : A~1
          begin
            sL_ImsProductID := genRealImsproductID(L_ProductsStrList.Strings[1]);



            preSendCommand(I_CmdInfoObj, sL_IccNo, sL_StbNo, sL_ImsProductID,sL_SubBeginDate, sL_SubEndDate);
//            Synchronize(preSendCommand);
          end
          else if (L_ProductsStrList.Strings[0]='B') and (L_ProductsStrList.Count=3) then //ex: B~5~20021231
          begin
            sL_ImsProductID := genRealImsproductID(L_ProductsStrList.Strings[1]);

            sL_SubEndDate := TUstr.replaceStr(L_ProductsStrList.Strings[2],'/','');
            preSendCommand(I_CmdInfoObj, sL_IccNo, sL_StbNo, sL_ImsProductID, sL_SubBeginDate, sL_SubEndDate);
//            Synchronize(preSendCommand);
          end;
        end;
      end
      else if (sL_CommandNo='B1') then //(開頻道-多組)
      begin
        //update send_nagra set notes='A~111111111111,A~222222222222,A~333333333333,B~555555555555~20020201~20020205';
        sL_Notes := I_CmdInfoObj.NOTES;

        {
        各 product 間以,區隔, product與期間之間以~區隔;
        第一碼是A者,表示共用此指令之subscription_begin_date 與 subscription_end_datee,
        第一碼是B者,
        表示自己有自己的 subscription_begin_date 與 subscription_end_date.
        Ex:A~1,A~2,A~3,B~5~20020201~20020205=>product 1,2,3以指令之 subscription_begin_date, subscription_end_date為日期;
        而product 5 的subscription_begin_date 為20020201, subscription_end_date 為20020205
        }


        L_NotesStrList := TUStr.ParseStrings(sL_Notes,',');
        for jj:=0 to L_NotesStrList.Count -1 do//表示有這麼多組 product 資料
        begin
          sL_ProductInfo := L_NotesStrList.Strings[jj];
          L_ProductsStrList := TUStr.ParseStrings(sL_ProductInfo,'~');
          if (L_ProductsStrList.Strings[0]='A') and (L_ProductsStrList.Count=2) then //ex : A~1
          begin
            sL_ImsProductID := genRealImsproductID(L_ProductsStrList.Strings[1]);
            preSendCommand(I_CmdInfoObj, sL_IccNo, sL_StbNo, sL_ImsProductID, sL_SubBeginDate, sL_SubEndDate);
//            Synchronize(preSendCommand);
          end
          else if (L_ProductsStrList.Strings[0]='B') and (L_ProductsStrList.Count=4) then //ex: B~5~20020201~20020205
          begin
            sL_ImsProductID := genRealImsproductID(L_ProductsStrList.Strings[1]);
            sL_SubBeginDate := TUstr.replaceStr(L_ProductsStrList.Strings[2],'/','');
            sL_SubEndDate := TUstr.replaceStr(L_ProductsStrList.Strings[3],'/','');
            preSendCommand(I_CmdInfoObj, sL_IccNo, sL_StbNo, sL_ImsProductID, sL_SubBeginDate, sL_SubEndDate);
//            Synchronize(preSendCommand);
          end;
        end;
      end
      else if (sL_CommandNo='B9') then //(refresh)
      begin
        //update send_nagra set notes='A~111111111111,A~222222222222,A~333333333333,B~555555555555~20020201~20020205';


        {
        各 product 間以,區隔, product與期間之間以~區隔;
        第一碼是A者,表示共用此指令之subscription_begin_date 與 subscription_end_datee,
        第一碼是B者,
        表示自己有自己的 subscription_begin_date 與 subscription_end_date.
        Ex:A~1,A~2,A~3,B~5~20020201~20020205=>product 1,2,3以指令之 subscription_begin_date, subscription_end_date為日期;
        而product 5 的subscription_begin_date 為20020201, subscription_end_date 為20020205
        }

        if sL_CurActiveLowLevelCmd='2' then
        begin
          sL_Notes := I_CmdInfoObj.NOTES;
          L_NotesStrList := TUStr.ParseStrings(sL_Notes,',');
          for jj:=0 to L_NotesStrList.Count -1 do//表示有這麼多組 product 資料
          begin
            sL_ProductInfo := L_NotesStrList.Strings[jj];
            L_ProductsStrList := TUStr.ParseStrings(sL_ProductInfo,'~');
            if (L_ProductsStrList.Strings[0]='A') and (L_ProductsStrList.Count=2) then //ex : A~1
            begin
              sL_ImsProductID := genRealImsproductID(L_ProductsStrList.Strings[1]);
              preSendCommand(I_CmdInfoObj, sL_IccNo, sL_StbNo, sL_ImsProductID, sL_SubBeginDate, sL_SubEndDate);
//              Synchronize(preSendCommand);
            end
            else if (L_ProductsStrList.Strings[0]='B') and (L_ProductsStrList.Count=4) then //ex: B~5~20020201~20020205
            begin
              sL_ImsProductID := genRealImsproductID(L_ProductsStrList.Strings[1]);
              sL_SubBeginDate := TUstr.replaceStr(L_ProductsStrList.Strings[2],'/','');
              sL_SubEndDate := TUstr.replaceStr(L_ProductsStrList.Strings[3],'/','');
              preSendCommand(I_CmdInfoObj, sL_IccNo, sL_StbNo, sL_ImsProductID, sL_SubBeginDate, sL_SubEndDate);
//              Synchronize(preSendCommand);
            end;
          end;
        end
        else
        begin
          preSendCommand(I_CmdInfoObj, sL_IccNo, sL_StbNo, sL_ImsProductID, sL_SubBeginDate, sL_SubEndDate);
//          Synchronize(preSendCommand);
        end;
      end
      else
      begin
        preSendCommand(I_CmdInfoObj, sL_IccNo, sL_StbNo, sL_ImsProductID, sL_SubBeginDate, sL_SubEndDate);
//        Synchronize(preSendCommand); howard
      end;

    end;
    except
      on ESocketError do Showmessage('Network error....');
    else
//      Showmessage('執行緒錯誤....');
    end;


end;

constructor TSendCommand.Create;
begin
//  CRS := TCriticalSection.Create;
    frmMain.nG_SC_MailId := 0;

    inherited Create(false);
end;



function TSendCommand.getTransNum(nI_CompCode:Integer; sI_SeqNo : String): String;
var
    bL_InsertMappingData : boolean;
    ii : Integer;
    sL_TransactionNum, sL_SeqNo, sL_SourceSeqNo : String;
begin
    sL_SourceSeqNo := sI_SeqNo;
    if nI_CompCode=StrToInt(TESTING_CMD_COMP_CODE) then //如果是測試公司別
    begin
      bL_InsertMappingData := false;

      sL_SeqNo := sI_SeqNo;


      sL_TransactionNum := IntToStr(nI_CompCode) + TUstr.AddString(sL_SeqNo,'0',false,SEQ_NUM_LENGTH);
    end
    else
    begin
      bL_InsertMappingData := true;
      if frmMain.nG_SC_RealCmdSeqNo>=frmMain.nG_SC_MaxRealCmdSeqNo then
        frmMain.nG_SC_RealCmdSeqNo := 1
      else
        INC(frmMain.nG_SC_RealCmdSeqNo);
      sL_SeqNo := IntToStr(frmMain.nG_SC_RealCmdSeqNo);

      sL_TransactionNum := IntToStr(nI_CompCode) + TUstr.AddString(sL_SeqNo,'0',false,SEQ_NUM_LENGTH);

    end;

    sL_TransactionNum := TUstr.AddString(sL_TransactionNum,'0',false,SEND_TRANSACTION_NUM_LENGTH);

    //down, 將 SeqNo 與 transaction number 的 mapping 關係紀錄起來
    if bL_InsertMappingData then
      dtmMain.insertMappingData(sL_SourceSeqNo,sL_TransactionNum);
    //up, 將 SeqNo 與 transaction number 的 mapping 關係紀錄起來

    result := sL_TransactionNum;

end;

procedure TSendCommand.writeTimeStamp;
var
    sL_ExeFileName, sL_ExePath, sL_TimeStampFileName : String;
    sL_Method ,sL_Info : String;
begin
    sL_ExeFileName := Application.ExeName;
    sL_ExePath := ExtractFileDir(sL_ExeFileName);
    sL_TimeStampFileName := sL_ExePath + '\' + TIME_STAMP_FILE_NAME;

    if (UpperCase(frmMain.sG_WriteTimeStamp)='Y')  then
    begin
      if FileExists(sL_TimeStampFileName) then
        frmMain.G_TimeStampStrList.LoadFromFile(sL_TimeStampFileName);
      frmMain.G_TimeStampStrList.Clear;
      frmMain.G_TimeStampStrList.Add(frmMain.sG_CurTimeStamp);
      frmMain.G_TimeStampStrList.SaveToFile(sL_TimeStampFileName);
    end
    else
    begin
      if FileExists(sL_TimeStampFileName) then
        deleteFile(sL_TimeStampFileName);
    end;
end;

procedure TSendCommand.updateCmdStatus;
begin
    try
    
      dtmMain.UpdateCMD_Status_For_SendCommand(frmMain.nG_SC_ConnectionID,frmMain.sG_SC_TransNumForUpdateCmdStatus,frmMain.sG_SC_CmdStatus,frmMain.sG_SC_ErrorCode,frmMain.sG_SC_ErrorMsg);
    except
      frmMain.showLogOnMemo(UPDATE_CMD_STATUS_ERROR_FLAG,' SendCommand.updateCmdStatus');
    end;
end;

procedure TSendCommand.writeSendCmdLog(sI_CommandNo : String);
begin
    if not frmMain.bG_CommandLog then Exit;//如果沒有 enable 指令 log, 則跳過此 procedure
    if StrToIntDef(Copy(frmMain.sG_SC_FullCommand,1,3),0)=(StrToInt(TESTING_CMD_COMP_CODE)) then Exit;   // 如果傳送的是測試指令,則跳過此 procedure

    try
      dtmMain.insertSendCmdLog(sI_CommandNo,frmMain.sG_SC_FullCommand, frmMain.sG_SC_Operator);
    except

    end;
end;

procedure TSendCommand.buildSendCmdConnection(
  var I_ClientSocket: TClientSocket);
var
    nL_Words : Integer;
    L_TmpRecord : TTempRecord;
    ii, nL_CmdLen : Integer;
    L_Socket : TClientSocket;
    L_WStream: TWinSocketStream;
    L_DeviceCall : TDeviceCall;
    L_DeviceStatus : TDeviceStatus;
    L_DeviceAnswer : TDeviceAnswer;
    sL_CmdObjName : String;
begin
    frmMain.bG_SC_buildSendCmdConnection := false;
    if Assigned(I_ClientSocket) then
    begin
      I_ClientSocket := nil;
      I_ClientSocket.Free;
    end;
    I_ClientSocket:= TClientSocket.Create(nil);
    I_ClientSocket.ClientType := ctBlocking;


//    L_Socket := G_MonitorFrame.ClientSocket1;
    L_Socket := I_ClientSocket;
    try
    begin
      with L_Socket do
      begin
        Address := frmMain.getSendCommandServerIP;
        Port := frmMain.getSendCommandPort;

        Open;
        L_DeviceCall.op_mode := char(0);

        sL_CmdObjName := 'SMSGW';

        //L_DeviceCall.obj_name :=
        FillChar(L_DeviceCall.obj_name, length(L_DeviceCall.obj_name), char(0)); { initialize the buffer }
        
        for ii:=1 to length(sL_CmdObjName) do
        begin
          L_DeviceCall.obj_name[ii-1] := sL_CmdObjName[ii];
        end;

//        StrPCopy(L_DeviceCall.obj_name,sL_CmdObjName);
        L_DeviceCall.obj_name_len := char(length(sL_CmdObjName));


        //1002
        {
        sL_CommandNo := '1002';//直接送出指令 1002, 用以確認 connection 是否順利建立 => 欲接收 feedback command 的 process ,才需要送出 command 1002
//        sL_CommandNo := sG_CommandNo;

        sL_CommandType := getCommandType(StrToInt(sL_CommandNo));
        sL_SmsCommandRootHeader := getSmsCommandRootHeader('', StrToInt(sL_CommandNo), sL_CommandType);
  //    showmessage(sL_SmsCommandRootHeader);
        sL_CurrentDate := TUdateTime.GetPureDateStr(date,'');
        sL_BoxNo := '';
        sL_CommandRootHeader := getCommandSection(sL_CommandType, sL_CurrentDate, sL_CurrentDate,sG_IccNo);
        sL_CommandBody := getCommandBody(StrToInt(sL_CommandNo), sL_BoxNo);
        sL_FullCommand := sL_SmsCommandRootHeader + sL_CommandRootHeader + sL_CommandBody;
        }

//        FillChar(L_DeviceCall.user_data, length(L_DeviceCall.user_data), char(0)); //=> 用此行..會有錯誤...

        for ii:=1 to length(L_DeviceCall.user_data) do
          L_DeviceCall.user_data[ii-1] := char(0);

        {
        for ii:=1 to length(sL_FullCommand) do
        begin
          L_DeviceCall.user_data[ii-1] := sL_FullCommand[ii];
        end;
        }
//        nL_CmdLen := 1 + 1 + length(sL_CmdObjName) + length(sL_FullCommand); //ok.. => 其實不確定要不要送 sL_CmdObjName
        nL_CmdLen := 1 + 1 + length(sL_CmdObjName); //ok.. => 其實不確定要不要送 sL_CmdObjName
//        nL_CmdLen := 1+1+ length(sL_FullCommand); //testing..

        L_TmpRecord := frmMain.getBinaryVal(nL_CmdLen,2,false);
        L_DeviceCall.len[0] := L_TmpRecord.TempAry[0];
        L_DeviceCall.len[1] := L_TmpRecord.TempAry[1];

        L_WStream := TWinSocketStream.Create(L_Socket.Socket, 60000); //howard




        nL_Words := write(@L_DeviceCall, nL_CmdLen+2, L_WStream); //build connection
//            showmessage('return:'+inttostr(nL_Words));


        nL_Words := L_WStream.Read(L_DeviceStatus, 3);
//        nL_Words := recv(L_Socket.Socket.SocketHandle , L_DeviceStatus, 3,0);
  //      showmessage('recv 1 的傳回值:'+inttostr(nL_Words));

        for ii:=1 to length(L_DeviceAnswer.data) do
          L_DeviceAnswer.data[ii-1] := char(0);

        nL_Words := L_WStream.Read(L_DeviceAnswer, L_DeviceStatus.len);
//        nL_Words := recv(L_Socket.Socket.SocketHandle , L_DeviceAnswer, L_DeviceStatus.len,0);
  //      nL_Words := recv(L_Socket.Socket.SocketHandle , L_DeviceStatus, 3,0);
  //      showmessage('recv 2 的傳回值:'+inttostr(nL_Words));

//        showmessage('device_call L_DeviceAnswer.code==' + L_DeviceAnswer.code);
//        showmessage('device_call L_DeviceAnswer.data==' + L_DeviceAnswer.data);

        if Assigned(L_WStream) then
          L_WStream.Free;
        frmMain.bG_HasBuildSendCmdConnection := true;
        frmMain.bG_SC_buildSendCmdConnection := true;
      end;
    end;
    except

      frmMain.sG_SC_ErrorCode := BUILD_CONNECTION_ERROR;
      frmMain.sG_SC_ErrorMsg := BUILD_CONNECTION_ERROR_MSG;

      frmMain.sG_SC_CmdStatus := 'E';
      frmMain.nG_SC_ConnectionID := frmMain.nG_CurActiveConn;
      updateCmdStatus;
//      Synchronize(updateCmdStatus);
      {
      try
        buildSendCmdConnection(I_ClientSocket, I_WinSocketStream);
      except
        frmMain.showLogOnMemo(RE_CONNECT_NAGRA_ERROR_FLAG,'重新連線到 Nagra');
      end;
      }
      raise;
      exit;
    end;


end;

procedure TSendCommand.receiveResponse;
var
    nL_WaitForDataTimeOut, ii,nL_Words, nL_Times : Integer;
    sL_ResponseString, sL_RealData : String;
    L_ErrorInfoStrList : TStringList;
    sL_ErrorInfoFileName : String;
    L_StrList : TStringList;
    L_TmpAry : array[0..1024] of Char;
    L_Stream : TWinSocketStream;
//    L_DeviceData : TDeviceData;
    procedure reBuildConn(nI_RaiseType:Integer; sI_Msg:String);
    var
        sL_LogInfo : String;
    begin
      //down, 重新建立 connection...

      case nI_RaiseType of
        1:
          sL_LogInfo := ' ReceivingCommandResponseThread.receiveResponse' + '-- test info =>' + '-- WaitForData timeout...' ;

        2:
          sL_LogInfo := ' ReceivingCommandResponseThread.receiveResponse.' + '-- test info =>' + '--Error message:' + sI_Msg;
      end;
      frmMain.showLogOnMemo(REBUILD_SEND_CMD_CONNECTION_TO_NAGRA,sL_LogInfo);

      frmMain.reBuildNagraConn;
      //up, 重新建立 connection...

    end;

begin
    try
      nL_Times := 1;
      while true do
      begin
        if nL_Times=5 then
          break
        else
          Inc(nL_Times);
        sleep(200);
            
        FillChar(L_TmpAry, sizeof(L_TmpAry), char(0));

        nL_Words := recv(frmMain.G_SendCmdClntSocket.Socket.SocketHandle , L_TmpAry, sizeof(L_TmpAry),0);
        if nL_Words>0 then
        begin
          sL_RealData := '';
          sL_RealData := frmMain.transReceivedDataCharAryToStr(L_TmpAry, nL_Words);

          L_StrList := TUstr.ParseStrings(sL_RealData,RECEIVED_DATA_SEP_FLAG);
          for ii:=0 to L_StrList.Count -1 do
          begin
            frmMain.G_ResponseMsgStrList.Add(L_StrList.Strings[ii]);

          end;
          L_StrList.Free;
          break;
        end;
      end;
    Except
      on E: Exception do
      begin
        if Assigned(L_Stream) then
          L_Stream.Free;
        sL_ErrorInfoFileName := 'c:\Error1.txt';
        L_ErrorInfoStrList := TStringList.Create;
        if FileExists(sL_ErrorInfoFileName) then
          L_ErrorInfoStrList.LoadFromFile(sL_ErrorInfoFileName);
        if L_ErrorInfoStrList.Count>500 then
          L_ErrorInfoStrList.Clear;
            L_ErrorInfoStrList.Insert(0,'   Msg:' +E.Message +', HelpContext:' + IntToStr(E.HelpContext));
        L_ErrorInfoStrList.Insert(0,'ReceivingCommandResponseThread.receiveResponse error:' + DateTimeToStr(now));
        L_ErrorInfoStrList.SaveToFile(sL_ErrorInfoFileName);
        L_ErrorInfoStrList.Free;
//        frmMain.createThreadForNagra(2, E.Message);

//        if AnsiPos('指定的網路名稱無法使用', E.Message) >0 then
        if AnsiPos('violation', E.Message) >0 then
        begin
        end
        else
          frmMain.showLogOnMemo(RECEIVE_CMD_RESPONSE_ERROR_FLAG,' SendCommand.receiveResponse.' + frmMain.getLanguageSettingInfo('frmMain_Msg_45_Content') + '-- test info =>' + '--Error message:' + E.Message);
{
        if AnsiPos('網路名稱無法使用', E.Message) >0 then
          reBuildConn(2, E.Message)
        else if AnsiPos('violation', E.Message) >0 then
        begin
        end
        else
          frmMain.showLogOnMemo(RECEIVE_CMD_RESPONSE_ERROR_FLAG,' ReceivingCommandResponseThread.receiveResponse.程式即將結束,然後重新啟動.' + '-- test info =>' + '--錯誤訊息:' + E.Message);
}
      end;
    end;




end;

procedure TSendCommand.onTimerStartAction;
var
    sL_TimeStr : STring;
    wL_Hour, wL_Min, wL_Sec, sL_MSec : word;
    L_CmdInfoObj : TCmdInfoObject;
begin
      DecodeTime(now, wL_Hour, wL_Min, wL_Sec, sL_MSec);
      sL_TimeStr := Format('%.2d%.2d', [wL_Hour, wL_Min]);
      if sL_TimeStr=DELETE_MAPPING_DATA_TIME then
        dtmMain.deleteMappingData;
      if frmMain.bG_RunSendingCommandThread then
      begin
        if frmMain.G_CommandInfoStrList.Count >0 then
        begin
          try
            L_CmdInfoObj := (frmMain.G_CommandInfoStrList.Objects[0] as TCmdInfoObject);
//            G_CmdInfoObj := (frmMain.G_CommandInfoStrList.Objects[0] as TCmdInfoObject);
            parseCommandData(L_CmdInfoObj);
            frmMain.G_CommandInfoStrList.Delete(0);
          except

          end;

        end;
      end;

end;


function TSendCommand.genRealIccNo(sI_SourceIccNo: String): String;
var
    sL_IccNo : String;
begin
    sL_IccNo := Copy(sI_SourceIccNo,1,10);
    sL_IccNo := TUstr.AddString(sL_IccNo,'0',false,ICC_CARDNO_LENGTH);//10 碼,右靠, 左補0

    result := sL_IccNo;
end;

function TSendCommand.genRealStbNo(sI_SourceStbNo: String): String;
var
    sL_StbNo : String;
begin
    sL_StbNo := Copy(Trim(sI_SourceStbNo),1,10);
    sL_StbNo := TUstr.AddString(sL_StbNo,'0',false,BOXNO_ID_LENGTH-4);//14 碼,右靠, 左補0,右補4個空白
//L_StrList.Add('2=' + sG_StbNo);
    if frmMain.bG_Simulator then
      sL_StbNo := TUstr.AddString(sL_StbNo,'0',true,BOXNO_ID_LENGTH)//14 碼,右靠, 左補0,右補4個0 => for 模擬器 => 例如使用者輸入 65536, 則轉成 00000655360000
    else
      sL_StbNo := TUstr.AddString(sL_StbNo,' ',true,BOXNO_ID_LENGTH);//14 碼,右靠, 左補0,右補4個空白 => for production

    result := sL_StbNo;        
end;

function TSendCommand.genRealImsproductID(
  sI_SourceImsProductID: String): String;
begin
    result := TUstr.AddString(sI_SourceImsProductID,'0',false,IMS_PRODUCT_ID_LENGTH);//12 碼, 右靠,左補0
end;

end.
