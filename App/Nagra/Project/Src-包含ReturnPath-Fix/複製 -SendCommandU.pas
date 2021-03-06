//Syn OK
unit SendCommandU;

interface

uses
  Classes, forms, comctrls, syncobjs, dbtables, Sysutils, Dialogs,
  ConstObjU, Ustru, UdateTimeu, scktcomp, winsock, inifiles, ADODB;


type

  TSendCommand = class(TThread)
  private
    { Private declarations }

  protected



  public
    G_AdoQryRealCmd:TADOQuery;
    sG_SubBeginDate, sG_SubEndDate : String;

    sG_BroadcastStartDate, sG_BroadcastEndDate : String;
    CRS: TCriticalSection;
    sG_SourceID, sG_DestID, sG_MopPPID : String;
    nG_Timeout : Integer;
    sG_RowID : String;
    sG_LogPath : String;
    sG_EmmCommandType, sG_EmmBroadcastMode, sG_EmmAddressType : String;
    sG_ControlCommandType : String;
    sG_ProductCommandType : String;
    sG_FeedbackCommandType : String;
    sG_OperationCommandType : String;
    bG_CommandLog, bG_ErrorLog, bG_ResponseLog : boolean;
    bG_BuildConnection, bG_SendComm, bG_WriteToMis : boolean;

    sG_TransactionNum : String;
    sG_MisSubBeginDate, sG_MisSubEndDate : String;
    nG_CompCode, nG_CmdStrListPos : Integer;
    sG_CurActiveLowLevelCmd : String;
    sG_RealCommand : String;
    sG_FullCommand : String;
    sG_CommandNo : String;
    sG_ImsProductID, sG_ZipCode : String;
    sG_MisIrdCmdID, sG_MisIrdCmdData, sG_IrdCmdID, sG_IrdOperation, sG_IrdDataLength, sG_IrdData : String;
    sG_PhoneNum1, sG_PhoneNum2, sG_PhoneNum3 : String;
    nG_CreditMode, nG_Credit, nG_ThresholdCredit : Integer;
    sG_EventName, sG_CCNum1 : String;
    sG_IpAddr, sG_CallbackDate, sG_CallbackTime : String;
    nG_Price, nG_CCPort, nG_CreditLimit : Integer;
    G_ListItem: TListItem;
    sG_ErrorCode,sG_ErrorMsg, sG_RecSeqNo : String;
    nG_CommandStatus: integer; //指令處理進度的狀態
    nG_Timer: integer;        //Time Out計數器
    bG_MakeSmsCommand : boolean;
    sG_SeqNo : String;
    sG_Notes, sG_CommandStatus : String;
    sG_Operator : String;
    sG_IccNo, sG_StbNo : String;
    sG_CallFreq, sG_FirstCallbackDate : String;
    G_CmdInfoObj : TCmdInfoObject;
    bG_RunSendingCommandThread, bG_SetZipCode : boolean;
    G_CommandInfoStrList : TStringList;
    G_Socket : TClientSocket;
    G_WStream: TWinSocketStream;
    function IsDataOK(var nI_ResultFlag:Integer) : boolean;
    procedure getHighLevelCommandInfo(sI_RowID:String; I_CmdInfoObj:TCmdInfoObject);
    procedure Execute; override;
    function getDetailCommandIDs(sI_HighLevelCommand:String):String;
    procedure writeCmdLog(sI_RowID:String; sI_HighLevelCmdNo, sI_LowLevelCmdNo : String; nI_ActionType:Integer);    
    function getTransNum:String;
    procedure changeFetchingAndSending;
    function getFormatIpAddr(sI_SrcIpAddr:String):String;
    procedure parseCommandData(I_CmdInfoObj:TCmdInfoObject);
    procedure preSendCommand;
    procedure makeSmsCommand; //組合出SMS Command
    function getSmsCommandRootHeader(sI_RowID:String; nI_CommandID : Integer; sI_CommandType : String) : String;
    function getCommandSection(sI_CommandType, sI_BDate, sI_EDate,sI_SerialNo:String) : String;
    function getCommandType(nI_CommandID : Integer) : String;
    function getCommandBody(nI_CommandID : Integer; sI_BoxNo:String) : String;
    function getRecSeqNo(nI_CmdType: integer): string;
    procedure transIrdInfo;
//    function getBinaryVal(nI_Value: Integer; nI_Digits: Integer; bI_Reverse: boolean):TTempRecord;

    function Write(I_Buf: Pointer; I_Count: Integer):Integer;overload;

    function Increment(Ptr: Pointer; Increment: Integer): Pointer;
    procedure writeErrorLog;
    procedure SendCommand;
    procedure insertCmdResult(sI_TransNum, sI_RealCommand:String);overload;
    procedure setTimerStatue2True;
    procedure setTimerStatue2False;
    procedure updateFetchCommandThreadToFalse;
    procedure updateCmdInfo;
    procedure createUiListItem;
    procedure getSeqNo;
    procedure getCmdInfoObj;
    procedure getDefaultCreditLimit;
    procedure setIfSetZipCode;
    procedure setRunSendingCommandThreadFlag;
    procedure setCmdInfoStrList;
    procedure getSocket;
    procedure getSocketStream;
    constructor Create;overload;
    constructor Create(I_AdoQryCmdLog:TADOQuery);overload;
    constructor Create(I_ScrollBox: TScrollBox; I_ListItem: TListItem; I_AdoQryCmdLog:TADOQuery);overload;
//    constructor Create(I_MFrame: TfrmMonitor; I_ScrollBox: TScrollBox; I_ListItem: TListItem; I_AdoQryCmdLog:TADOQuery);overload;
    procedure buildConnection(I_ClientSocket : TClientSocket);


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

{
constructor TSendCommand.Create(I_MFrame: TfrmMonitor;
  I_ScrollBox: TScrollBox; I_ListItem: TListItem; I_AdoQryCmdLog:TADOQuery);
begin
  G_AdoQryRealCmd := I_AdoQryCmdLog;
  G_ListItem := I_ListItem;
  G_ListItem.SubItems[4] := '0%';
  sG_ErrorCode:=NO_ERROR; //預設為 -- 成功訊息代碼
  sG_ErrorMsg :='';
  nG_CommandStatus :=0;
  nG_Timer := 0;
  bG_MakeSmsCommand := false;

//  G_MonitorFrame := I_MFrame;
  G_ScbBox := I_ScrollBox;
//  G_MonitorFrame.Timer1.OnTimer := onTimer;

//  FreeOnTerminate := True;
//  Priority := tpnormal;

//  CRS := TCriticalSection.Create;

  inherited Create(true);
end;
}
procedure TSendCommand.Execute;
var
    ii : Integer;
    sL_SeqNo : String;
    nL_Count : Integer;
    L_CmdInfoObj : TCmdInfoObject;
begin

    {
    self.FreeOnTerminate := true;
    self.OnTerminate := frmMain.SendCommandThreadDone;
    Synchronize(getSocket);
    Synchronize(getSocketStream);

    while true do
    begin
      Synchronize(setRunSendingCommandThreadFlag);
      if bG_RunSendingCommandThread then
      begin
//          CRS.Enter; howard
          Synchronize(updateFetchCommandThreadToFalse);
          Synchronize(setCmdInfoStrList);
          if G_CommandInfoStrList<> nil then
          begin
            nL_Count := G_CommandInfoStrList.Count;
            for ii:= 0 to nL_Count -1 do
            begin
              nG_CmdStrListPos := ii;
              Synchronize(getSeqNo);
//              sL_SeqNo := frmMain.G_CommandInfoStrList.Strings[ii];
              sL_SeqNo := sG_SeqNo;
              if sL_SeqNo<>'' then
              begin
                Synchronize(getCmdInfoObj);
//                L_CmdInfoObj := (frmMain.G_CommandInfoStrList.Objects[ii] as TCmdInfoObject);
                L_CmdInfoObj := G_CmdInfoObj;
                parseCommandData(L_CmdInfoObj);
                Synchronize(updateCmdInfo);

              end;
            end;
          end;
//          CRS.Enter;
          Synchronize(changeFetchingAndSending);
//          CRS.Leave;

//          CRS.Leave; howard
      end;
    end;
    }
end;


function TSendCommand.getCommandBody(nI_CommandID: Integer; sI_BoxNo:String): String;
var
    sL_Result, sL_CommandID : String;
    sL_StuNumber, sL_ImsProductID : String;
    sL_BDate, sL_EDate, sL_ZipCode, sL_CreditMode, sL_Credit : String;
    sL_ThresholdCredit, sL_EventName, sL_EventNameLength, sL_Price, sL_CCNum1 : String;
    sL_IpAddr, sL_CCPort, sL_CbDate, sL_CbTime : String;
    sL_CallFreq, sL_DateFirstCall, sL_CreditLimit : String;
    sL_PhoneNum1, sL_PhoneNum2, sL_PhoneNum3 : String;
    sL_IrdCmdID, sL_IrdOperation, sL_IrdDataLength, sL_IrdData : String;
begin
    {
    此 function 用來組合出各指令的 parameters.
    ex: EMM command 的參數請參考 SMS GATEWAY SPECIFICATION 第五章
    }
    sL_CommandID := TUstr.AddString(IntToStr(nI_CommandID),'0',false,COMMAND_ID_LENGTH);
    sL_StuNumber := TUstr.AddString(sI_BoxNo,'0',false,BOXNO_ID_LENGTH);
    sL_ImsProductID := sG_ImsProductID;
    
    sL_BDate := TUstr.replaceStr(sG_SubBeginDate,'/','');
    sL_EDate := TUstr.replaceStr(sG_SubEndDate,'/','');


    sL_ZipCode := sG_ZipCode;
    sL_CreditMode := Format('%.2d', [nG_CreditMode]);
    {
      sL_CreditMode :
        01 => ADD
        02 => SUBTRACT
        03 => SET CREDIT
        04 => SET BALANCE
        05 => SUB OFFSET
    }
    sL_Credit := Format('%.7d', [nG_Credit]);
    sL_ThresholdCredit := Format('%.7d', [nG_ThresholdCredit]);
    sL_EventName := Uppercase(TUstr.AddString(sG_EventName,' ',true,EVENT_NAME_LENGTH));
    sL_EventNameLength := Format('%.2d',[length(sG_EventName)]);
    sL_Price := Format('%.5d', [nG_Price]);

    sL_IpAddr := getFormatIpAddr(sG_IpAddr);
    sL_CCPort := Format('%.5d', [nG_CCPort]);
    sL_CbDate := sG_CallbackDate;
    sL_CbTime := sG_CallbackTime;
    sL_CallFreq := sG_CallFreq;
    sL_DateFirstCall := sG_FirstCallbackDate;
    sL_CreditLimit := Format('%.7d', [nG_CreditLimit]);

    sL_IrdCmdID := sG_IrdCmdID;
    sL_IrdOperation := sG_IrdOperation;
    sL_IrdDataLength := sG_IrdDataLength;
    sL_IrdData := sG_IrdData;

    //down, 將 IRD_DATA 用0補滿96個bytes
    sL_IrdData := TUstr.AddString(sL_IrdData,'0',true,IRD_COMMAND_DATA_LENGTH);

    //down, 電話號碼的轉換 => 這樣的程式寫法,對simulator 無法 work; 但連線到瑞士測試時,卻又ok.
    sL_CCNum1 := TUstr.AddString(sG_CCNum1,chr(32),true,CALL_COLLECTOR_PHONE_NUM_LENGTH);
    sL_PhoneNum1 := TUstr.AddString(sG_PhoneNum1,chr(32),true,AUTHORIZED_PHONE_NUM_LENGTH);
    sL_PhoneNum2 := TUstr.AddString(sG_PhoneNum2,chr(32),true,AUTHORIZED_PHONE_NUM_LENGTH);
    sL_PhoneNum3 := TUstr.AddString(sG_PhoneNum3,chr(32),true,AUTHORIZED_PHONE_NUM_LENGTH);
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

    end;


    result := sL_Result;
end;

function TSendCommand.getSmsCommandRootHeader(sI_RowID:String; nI_CommandID : Integer; sI_CommandType : String): String;
var
    sL_Result : String;
begin

//    sL_TransactionNum :=
    {
    if sI_RowID<>'' then
    begin
      sL_SQL := 'update send_nagra set TRANSACTION_NUM=' + STR_SEP + sG_TransactionNum + STR_SEP;
      sL_SQL := sL_SQL + ' where ROWID=' + STR_SEP + sI_RowID + STR_SEP;
      with frmMain.qryUpdater do
      begin
        Close;
        DatabaseName := SEND_NAGRA_DB_NAME;
        SQL.Clear;
        SQL.Add(sL_SQL);
        ExecSQL;
        Close;

      end;
    end;
    }
    sL_Result := sG_TransactionNum + sI_CommandType;
    sL_Result := sL_Result + sG_SourceID + sG_DestID + sG_MopPPID;
    sL_Result := sL_Result + TUdateTime.GetPureDateStr(date,'');
    result := sL_Result;
end;

function TSendCommand.getCommandType(nI_CommandID: Integer): String;
var
    sL_Result : String;
begin

    if (EMM_CommandBeginID<=nI_CommandID)and(nI_CommandID <=EMM_CommandEndID) then
      sL_Result := sG_EmmCommandType
    else if (CONTROL_CommandBeginID<=nI_CommandID)and(nI_CommandID <=CONTROL_CommandEndID) then
      sL_Result := sG_ControlCommandType
    else if (FEEDBACK_CommandBeginID<=nI_CommandID)and(nI_CommandID <=FEEDBACK_CommandEndID) then
      sL_Result := sG_FeedbackCommandType
    else if (PRODUCT_CommandBeginID<=nI_CommandID)and(nI_CommandID <=PRODUCT_CommandEndID) then
      sL_Result := sG_ProductCommandType
    else if (OPERATION_CommandBeginID<=nI_CommandID)and(nI_CommandID <=OPERATION_CommandEndID) then
      sL_Result := sG_OperationCommandType
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

    if (sI_CommandType=sG_EmmCommandType) then
    begin  //SMS Gateway Interface definition, P18
      sL_Result := sG_EmmBroadcastMode + sI_BDate + sI_EDate + sG_EmmAddressType;
      if (sG_EmmAddressType=EMM_UNIQUE_ADDRESS_TYPE) then
        sL_Result := sL_Result + sI_SerialNo;
    end
    else if (sI_CommandType=sG_ControlCommandType) then
    begin //SMS Gateway Interface definition, P20
      sL_CurDateStr := TUdateTime.GetPureDateStr(date,'');
      L_SysParam := TfrmSysParam.Create(Application);
      L_SysParam.getControlCommandSectionParam(sL_BroadastMode, sL_AddressType);
      L_SysParam.Free;
      sL_Result := sL_BroadastMode + sL_CurDateStr + sL_CurDateStr + sL_AddressType + sI_SerialNo; //SMS Gateway Interface definition, P20

    end
    else if (sI_CommandType=sG_ProductCommandType) then
    begin
      sL_Result := ''; //SMS Gateway Interface definition, P21
    end
    else if (sI_CommandType=sG_FeedbackCommandType) then
    begin
      sL_Result := sI_SerialNo; //SMS Gateway Interface definition, P21
    end
    else if (sI_CommandType=sG_OperationCommandType) then
    begin
      sL_Result := ''; //SMS Gateway Interface definition, P21    
    end;


    result := sL_Result;
end;

procedure TSendCommand.makeSmsCommand;
var
  {
  F1: TextFile;
  Request : TRequestLib;
  Header: string;  //Request Header;
  Body: string;    //Request
  Footer: string;  //Request UAs
  }
  sL_SmsCommandRootHeader, sL_CommandSection : String;
  sL_CommandBody, sL_CommandType : String;
begin
    try
      bG_MakeSmsCommand := false;

    //  sG_RecSeqNo := GetRecSeqNo(1);

      //將SEND_NAGRA的指令狀態更改為處理中：P
//      UpdateCMD_STATUS(sG_RowID, 'P');

      //Analyze SEND_NAGRA Command

      {
      G_MonitorFrame.Gauge1.Progress := 50;
      G_MonitorFrame.chkStatus1.Checked := true;
      G_MonitorFrame.lblRecNo.Caption := sG_RowID;
      }
      { howard....
      G_ListItem.ImageIndex := -1;
      G_ListItem.StateIndex := 0;
      G_ListItem.SubItems[4] := '50%';
      G_ListItem.SubItems[6] := sG_SeqNo;
      G_ListItem.SubItems[5] := sG_Operator;
      G_ListItem.SubItems[9] := datetimetostr(now);

      }


      sL_CommandType := getCommandType(StrToInt(sG_RealCommand));

      sL_SmsCommandRootHeader := getSmsCommandRootHeader(sG_RowID, StrToInt(sG_RealCommand), sL_CommandType);

  //    showmessage(sL_SmsCommandRootHeader);

  //    sL_CommandSection := getCommandSection(sL_CommandType, sG_SubBeginDate, sG_SubEndDate,sG_IccNo);
      sL_CommandSection := getCommandSection(sL_CommandType, sG_BroadcastStartDate, sG_BroadcastEndDate,sG_IccNo);

      sL_CommandBody := getCommandBody(StrToInt(sG_RealCommand), sG_StbNo);

      sG_FullCommand := sL_SmsCommandRootHeader + sL_CommandSection +sL_CommandBody;

  //    showmessage('EmmCommandRootHeader:'+sL_CommandRootHeader)
  //    showmessage(sL_FullCommand);
  //  end;



      if (bG_CommandLog) then // 寫入指令log
      begin
        writeCmdLog(sG_RowID, sG_CommandNo,sG_RealCommand,2) ;
       
      end;
    except
      writeERrorLog;
      bG_MakeSmsCommand := false;
      Exit;
    end;
    bG_MakeSmsCommand := true;

end;

function TSendCommand.getRecSeqNo(nI_CmdType: integer): string;
var

    sL_SQL, sL_TmpStr: string;
begin
    sL_TmpStr:='';

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
    result:=sL_TmpStr;
end;



procedure TSendCommand.buildConnection(I_ClientSocket : TClientSocket);
var
    nL_Words : Integer;
    L_TmpRecord : TTempRecord;
    ii, nL_CmdLen : Integer;
    L_Socket : TClientSocket;
//    L_WStream: TWinSocketStream;
    L_DeviceCall : TDeviceCall;
    L_DeviceStatus : TDeviceStatus;
    L_DeviceAnswer : TDeviceAnswer;
    sL_CmdObjName : String;
begin

    bG_BuildConnection := false;
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
        for ii:=1 to length(L_DeviceCall.obj_name) do
          L_DeviceCall.obj_name[ii-1] := char(0);

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
//        L_WStream := TWinSocketStream.Create(L_Socket.Socket, 60000); //0110
//        L_WStream := G_MonitorFrame.G_WStream;
//        I_WinSocketStream := L_WStream; //0110

        nL_Words := write(@L_DeviceCall, nL_CmdLen+2); //send_data..

//        nL_Words := write(@L_DeviceCall, nL_CmdLen+2, L_WStream); //build connection
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

//        writeCmdLog(sG_RowID, sG_CommandNo,sG_RealCommand, 3) ;
        frmMain.bG_HasBuildConnection := true;
      end;
    end;
    except
      sG_ErrorCode := BUILD_CONNECTION_ERROR;
      sG_ErrorMsg := BUILD_CONNECTION_ERROR_MSG;
      dtmMain.UpdateCMD_STATUS(frmMain.nG_CurActiveConn,sG_SeqNo,'E',sG_ErrorCode,sG_ErrorMsg);
      writeErrorLog;
//      writeCmdLog(sG_RowID, sG_CommandNo,sG_RealCommand, 4) ;
      exit;
    end;

    bG_BuildConnection := true;
end;



function TSendCommand.Increment(Ptr: Pointer;
  Increment: Integer): Pointer;
begin
  Result := Pointer(Longint(Ptr) + Increment);
end;

procedure TSendCommand.writeErrorLog;
begin
    if bG_ErrorLog then
    begin
      try
       sG_RecSeqNo := GetRecSeqNo(1);
        with G_AdoQryRealCmd do
        begin
          SQL.Clear;
          Close;
          SQL.Add('insert into ERRORLOG(Req_No,Row_ID,Command_No,ICC_No,STB_No,BDATE,EDATE,Notes,Cmd_Status,Operator,Err_Code,Err_Msg,Log_Date, Log_Time)');
          SQL.Add(' values('+STR_SEP + sG_RecSeqNo + STR_SEP +','+ STR_SEP + sG_RowID + STR_SEP +','''+sG_CommandNo+''',');
          SQL.Add(''''+sG_IccNo+''','''+sG_StbNo+''','+''''+sG_SubBeginDate+''','''+sG_SubEndDate+''',');
          SQL.Add(''''+sG_Notes+''','''+sG_CommandStatus+''',');
          SQL.Add(''''+sG_Operator+ STR_SEP + ',');
          SQL.Add(''''+sG_ErrorCode + ''','''+ sG_ErrorMsg + ''','''+ TUdateTime.GetPureDateStr(date,'') + STR_SEP + ',' + STR_SEP + TUdateTime.GetPureTimeStr(now,'') + STR_SEP + ')' );
          ExecSQL;
          Close;
        end;
      except

      end;
    end;

end;

procedure TSendCommand.SendCommand;
var
    nL_Words : Integer;
    ii, nL_Len : Integer;
    L_TmpRecord : TTempRecord;
    L_DeviceData : TDeviceData;
    L_WStream: TWinSocketStream;
    L_Socket : TClientSocket;
    L_ErrorInfoStrList : TStringList;
    sL_ErrorInfoFileName : String;
    sL_Info : String;
begin

    bG_SendComm := true;
    //down, 傳送command
    try
    begin
      Synchronize(setTimerStatue2False);

//      G_ListItem.SubItems[4] := '100%'; howard...
      {
      G_MonitorFrame.Gauge1.Progress := 100;
      G_MonitorFrame.chkStatus2.Checked := true;
      }
      nL_Len := length(sG_FullCommand);//ok..
//      nL_Len := 1 + 1 + length('EMGR_EME_SVR') + length(sG_FullCommand); //testing..=> 其實不確定要不要送 sL_CmdObjName

      L_TmpRecord := frmMain.GetBinaryVal(nL_Len,2,false);
      L_DeviceData.len[0] := L_TmpRecord.TempAry[0];
      L_DeviceData.len[1] := L_TmpRecord.TempAry[1];

      for ii:=1 to length(L_DeviceData.data) do
        L_DeviceData.data[ii-1] := char(0);

      for ii:=1 to length(sG_FullCommand) do
      begin
        L_DeviceData.data[ii-1] := sG_FullCommand[ii];
      end;




      nL_Words := write(@L_DeviceData, nL_Len+2); //send_data..
//      writeCmdLog(sG_RowID, sG_CommandNo,sG_RealCommand, 5) ;
      insertCmdResult(Copy(sG_FullCommand,1,9), sG_RealCommand);
      Synchronize(setTimerStatue2True);
      {
      Synchronize(getSocket);
      Synchronize(getSocketStream);
      L_Socket := G_Socket;
      L_WStream := G_WStream;

      nL_Words := write(@L_DeviceData, nL_Len+2, L_WStream); //send_data..
//      writeCmdLog(sG_RowID, sG_CommandNo,sG_RealCommand, 5) ;
      insertCmdResult(Copy(sG_FullCommand,1,9), sG_RealCommand);
      Synchronize(setTimerStatue2True);
      }
    end
    except
      on E: Exception do
      begin

        sL_ErrorInfoFileName := 'c:\Error7.txt';
        L_ErrorInfoStrList := TStringList.Create;
        if FileExists(sL_ErrorInfoFileName) then
          L_ErrorInfoStrList.LoadFromFile(sL_ErrorInfoFileName);
        if L_ErrorInfoStrList.Count>500 then
          L_ErrorInfoStrList.Clear;

        L_ErrorInfoStrList.Insert(0,'   Detail Msg: Full Cmd=' +sG_FullCommand);
        L_ErrorInfoStrList.Insert(0,'   Msg:' +E.Message +', HelpContext:' + IntToStr(E.HelpContext));
        L_ErrorInfoStrList.Insert(0,'TSendCommand.SendCommand error:' + DateTimeToStr(now));
        L_ErrorInfoStrList.SaveToFile(sL_ErrorInfoFileName);
        L_ErrorInfoStrList.Free;

        nG_CommandStatus := 3; //send command 失敗
        G_ListItem.SubItemImages[3] := 3;
        sG_ErrorCode := SEND_COMMAND_ERROR;
        sG_ErrorMsg := SEND_COMMAND_ERROR_MSG;
        bG_SendComm := false;
        dtmMain.UpdateCMD_STATUS(frmMain.nG_CurActiveConn,sG_SeqNo,'E',sG_ErrorCode,sG_ErrorMsg);
        writeERrorLog;
  //      writeCmdLog(sG_RowID, sG_CommandNo,sG_RealCommand, 6) ;
        exit;
      end;
    end;
    //up, 傳送command
    if frmMain.bG_ShowUI then
      frmMain.G_HandleUIThread.updateListItemStatus(G_ListItem,1,'','','');


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

    sL_Result := '';
    L_StrList := TUStr.ParseStrings(sI_SrcIpAddr,'.');
    if L_StrList<> nil then
    begin
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
    MessageDlg('找不到系統環境設定檔.無法傳送指令',mtWarning,[mbOK],0);
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

procedure TSendCommand.getHighLevelCommandInfo(sI_RowID:String; I_CmdInfoObj:TCmdInfoObject);
var
    L_StrList : TStringList;
begin
//L_StrList := TStringList.Create;
      sG_CommandNo := I_CmdInfoObj.HIGH_LEVEL_CMD_ID;

//      sG_IccNo := I_CmdInfoObj.ICC_NO;
      sG_IccNo := Copy(I_CmdInfoObj.ICC_NO,1,10);
      sG_IccNo := TUstr.AddString(sG_IccNo,'0',false,ICC_CARDNO_LENGTH);//10 碼,右靠, 左補0
//L_StrList.Add(sG_IccNo);

//      sG_StbNo := I_CmdInfoObj.STB_NO;
//L_StrList.Add('1=' + I_CmdInfoObj.STB_NO);
      sG_StbNo := Copy(I_CmdInfoObj.STB_NO,1,10);
      sG_StbNo := TUstr.AddString(sG_StbNo,'0',false,BOXNO_ID_LENGTH-4);//14 碼,右靠, 左補0,右補4個空白
//L_StrList.Add('2=' + sG_StbNo);
      if frmMain.bG_Simulator then
        sG_StbNo := TUstr.AddString(sG_StbNo,'0',true,BOXNO_ID_LENGTH)//14 碼,右靠, 左補0,右補4個0 => for 模擬器 => 例如使用者輸入 65536, 則轉成 00000655360000
      else
        sG_StbNo := TUstr.AddString(sG_StbNo,' ',true,BOXNO_ID_LENGTH);//14 碼,右靠, 左補0,右補4個空白 => for production
//L_StrList.Add(sG_StbNo);
//L_StrList.SaveToFile('c:\aabc.txt');
//L_StrList.Free;

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
      sG_Notes := I_CmdInfoObj.NOTES;
      {
      if frmMain.G_NotesHasValueStrList.IndexOf(sG_CommandNo)<>-1 then
      begin
        sG_Notes := dtmMain.getNotes(frmMain.nG_CurActiveConn, sG_SeqNo);
      end;
      }
//      sG_Option := fieldbyname('OPTION').asstring;

      sG_CommandStatus := I_CmdInfoObj.CMD_STATUS;
      sG_Operator := I_CmdInfoObj.OPERATOR;
      sG_ErrorCode := '';
      sG_ErrorMsg := '';

      sG_ImsProductID := I_CmdInfoObj.IMS_PRODUCT_ID;
      sG_ImsProductID := TUstr.AddString(sG_ImsProductID,'0',false,IMS_PRODUCT_ID_LENGTH);//12 碼, 右靠,左補0

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

      sG_MisIrdCmdID :=  I_CmdInfoObj.MIS_IRD_CMD_ID;
      sG_MisIrdCmdData :=  I_CmdInfoObj.MIS_IRD_CMD_DATA;

      nG_CompCode :=  StrToIntDef(I_CmdInfoObj.COMPCODE,0);

      if sG_MisIrdCmdID <> '' then
      begin
        sG_MisIrdCmdID :=  TUstr.AddString(I_CmdInfoObj.MIS_IRD_CMD_ID,'0',false,3);
        transIrdInfo;
      end;


end;

procedure TSendCommand.writeCmdLog(sI_RowID: String; sI_HighLevelCmdNo,
  sI_LowLevelCmdNo: String; nI_ActionType:Integer);
var
    sL_SQL : STring;
begin
    if not bG_CommandLog then Exit;//如果沒有 enable 指令 log, 則跳過此 procedure

    case nI_ActionType of
      1://init command
       begin
         sL_SQL := 'insert into CmdLog(TransactionNum, HighLevelCmdNo, LowLevelCmdNo,CmdText,LogDateTime)';
         sL_SQL := sL_SQL + 'values(' + STR_SEP + sG_TransactionNum + STR_SEP +','+ STR_SEP + sI_HighLevelCmdNo + STR_SEP + ',';
         sL_SQL := sL_SQL + STR_SEP + sI_LowLevelCmdNo + STR_SEP + ',' + STR_SEP +sG_FullCommand + STR_SEP +',' + STR_SEP + DateTimeToStr(now) + STR_SEP + ')';
       end;
      2://build command
       begin
         sL_SQL := 'update CmdLog set CmdText=' +  STR_SEP + sG_FullCommand + STR_SEP ;
         sL_SQL := sL_SQL +' where TransactionNum=' + STR_SEP + sG_TransactionNum + STR_SEP + ' and HighLevelCmdNo=' + STR_SEP + sI_HighLevelCmdNo + STR_SEP;
         sL_SQL := sL_SQL +' and LowLevelCmdNo=' + STR_SEP + sI_LowLevelCmdNo + STR_SEP;
       end;
    end;


    with G_AdoQryRealCmd do
    begin
      SQL.Clear;
      Close;
      SQL.Add(sL_SQL);
      ExecSQL;
      Close;
    end;
end;


procedure TSendCommand.preSendCommand;
var
    sL_ErrorInfoFileName : String;
    L_ErrorInfoStrList : TStringList;
begin
    try
      sG_RealCommand := sG_CurActiveLowLevelCmd;

//      CRS.Enter;

//      Synchronize(makeSmsCommand);       //產生SMS Command
      writeCmdLog(sG_RowID, sG_CommandNo,sG_RealCommand, 1) ;
      makeSmsCommand;       //產生SMS Command


      { howard...
      if ii=0 then
        G_ListItem.SubItems[1] := sG_RowID + ' : ' + sG_RealCommand
      else
        G_ListItem.SubItems[1] := G_ListItem.SubItems[1] + '-' + sG_RealCommand;
      }

      if bG_MakeSmsCommand then
      begin
//        G_ListItem.SubItemImages[2] := 2; howard...
    //    sleep(3000);//停頓3 secs

//        Synchronize(SendCommand);
//        for jj:=0 to 100 do
          SendCommand; //0109
//        showmessage('123');
      end
      else
      begin
        nG_CommandStatus := 1; //產生SMS Command失敗
//        G_ListItem.SubItemImages[2] := 3; howard...
//        CRS.Leave;
        exit;
      end;
    except
      on E: Exception do
      begin

        sL_ErrorInfoFileName := 'c:\Error6.txt';
        L_ErrorInfoStrList := TStringList.Create;
        if FileExists(sL_ErrorInfoFileName) then
          L_ErrorInfoStrList.LoadFromFile(sL_ErrorInfoFileName);
        if L_ErrorInfoStrList.Count>500 then
          L_ErrorInfoStrList.Clear;
        L_ErrorInfoStrList.Insert(0,'   Msg:' +E.Message +', HelpContext:' + IntToStr(E.HelpContext));
        L_ErrorInfoStrList.Insert(0,'TSendCommand.preSendCommand error:' + DateTimeToStr(now));
        L_ErrorInfoStrList.SaveToFile(sL_ErrorInfoFileName);
        L_ErrorInfoStrList.Free;
      end;
    end;
      {
      if (bG_SendComm) then
      begin
        writeToMis(sG_RowID, sG_CommandNo,sG_RealCommand);
      end
      else
      begin
        CRS.Leave;
        exit;
      end;

      if (bG_WriteToMis) then
      begin

      end
      else
      begin
        CRS.Leave;
        exit;
      end;
      }
//      CRS.Leave;
end;

function TSendCommand.IsDataOK(var nI_ResultFlag:Integer): boolean;
begin
    if (length(sG_StbNo) <> 0) and (length(sG_StbNo) <> BOXNO_ID_LENGTH) then
    begin
      nI_ResultFlag := -1;
      result := false;
      exit;
    end;
    if (length(sG_IccNo) <> 0) and  (length(sG_IccNo) <> ICC_CARDNO_LENGTH) then
    begin
      nI_ResultFlag := -2;
      result := false;
      exit;
    end;

    if (length(sG_ImsProductID) <> 0) and  (length(sG_ImsProductID) <> IMS_PRODUCT_ID_LENGTH) then
    begin
      nI_ResultFlag := -3;
      result := false;
      exit;
    end;

end;

procedure TSendCommand.transIrdInfo;
var
    nL_AscIIValue,ii : Integer;
begin
    //以下設定請參考文件 Hitron Technologies StbCakIrdSpe010302.pdf 以及 Nagravision S.A. SMS Gateway intergace definition V02.06.02
    //但這兩份文件與模擬軟體的回應似乎不太一致...

    sG_IrdData := sG_MisIrdCmdData;
    case StrToInt(sG_MisIrdCmdID) of
      1: //(清除親子密碼)
       begin
        sG_IrdCmdID := '018';
        sG_IrdOperation :=  '001';
//        sG_IrdDataLength :=  '00';
        sG_IrdDataLength :=  '04';
        sG_IrdData := '';
        if sG_MisIrdCmdData='' then
          sG_MisIrdCmdData := '0000';//預設之PIN Code
        for ii:=1 to length(sG_MisIrdCmdData) do
        begin
          nL_AscIIValue := Ord(sG_MisIrdCmdData[ii]);//取出每個 byte 的 ASCII
          sG_IrdData := sG_IrdData + IntToHex(nL_AscIIValue,0); //取出該 ASCII 的16進位值
        end;
       end;
      2: //(傳送訊息)
       begin

       end;
      3: //(Force Tune)
       begin
        sG_IrdCmdID := '193';       
        sG_IrdOperation :=  '001';
        sG_IrdDataLength :=  TUstr.AddString(IntToStr(length(sG_IrdData)),'0',false,2);
       end;
      4: //(set network id)
       begin
        sG_IrdCmdID := NETWORK_OPERATION_194; //194=>operation 要為'', 且 data length=2;
//        sG_IrdCmdID := NETWORK_OPERATION_195; //195=>operation表示多久要 reset STB, data length=4
        sG_IrdOperation :=  '001';
        sG_IrdData := frmMain.getCompNetworkInfo(nG_CompCode,sG_IrdCmdID);
        sG_IrdDataLength :=  TUstr.AddString(FloatToStr(length(sG_IrdData)/2),'0',false,2);

       end;
    end;


end;


procedure TSendCommand.insertCmdResult(sI_TransNum, sI_RealCommand: String);
var
    sL_SQL : String;
begin
    if frmMain.bG_ResponseLog then
    begin
      sL_SQL := 'insert into CmdResult(TransactionNum,RealCmd)';
      sL_SQL := sL_SQL + 'values(' + STR_SEP + sI_TransNum + STR_SEP +',' + STR_SEP + sI_RealCommand +STR_SEP + ')';


      with frmMain.adoRealCmd do
      begin
        SQL.Clear;
        Close;
        SQL.Add(sL_SQL);
        ExecSQL;
        Close;
      end;
    end;
end;


constructor TSendCommand.Create(I_ScrollBox: TScrollBox;
  I_ListItem: TListItem; I_AdoQryCmdLog: TADOQuery);
begin
    G_AdoQryRealCmd := I_AdoQryCmdLog;
    G_ListItem := I_ListItem;
    G_ListItem.SubItems[4] := '0%';
    sG_ErrorCode:=NO_ERROR; //預設為 -- 成功訊息代碼
    sG_ErrorMsg :='';
    nG_CommandStatus :=0;
    nG_Timer := 0;
    bG_MakeSmsCommand := false;
    {
    flag := false;  //FTP命令執行成功與否 1.成功 0.失敗
    runFTPFlag := false;
    makeResponseFlag := false;
    RecSeqNo := '';
    }
    {
    FreeOnTerminate := True;
    Priority := tpnormal;
    }
//    CRS := TCriticalSection.Create;

    inherited Create(false);
end;

procedure TSendCommand.parseCommandData(I_CmdInfoObj:TCmdInfoObject);
var
    sL_Notes : String;
    sL_DetailCommandIDs, sL_ProductInfo : String;
    L_StrList, L_NotesStrList, L_ProductsStrList, L_ProductInfoStrList,  L_CardStrList  : TStringList;
    ii, jj, kk : Integer;
    bL_ProfuctFieldFormatError : boolean;
    nL_ResultFlag  : Integer;
begin
    sG_RowID := I_CmdInfoObj.ROWID;



    try
      getHighLevelCommandInfo(sG_RowID, I_CmdInfoObj);
  //    frmMain.G_HandleUIThread.createListItem(sG_SeqNo,sG_CommandNo,sG_Operator);

      sL_DetailCommandIDs := getDetailCommandIDs(sG_CommandNo);
      if sL_DetailCommandIDs='' then Exit;
      if not IsDataOK(nL_ResultFlag) then
      begin
        case nL_ResultFlag of
          -1 :
           begin
             sG_ErrorCode:= STB_NO_LENGTH_ERROR;
             sG_ErrorMsg := STB_NO_LENGTH_ERROR_MSG;
           end;
          -2 :
           begin
             sG_ErrorCode:= ICC_CARD_NO_LENGTH_ERROR;
             sG_ErrorMsg := ICC_CARD_NO_LENGTH_ERROR_MSG;
           end;
          -3 :
           begin
             sG_ErrorCode:= IMS_PRODUCT_ID_LENGTH_ERROR;
             sG_ErrorMsg := IMS_PRODUCT_ID_LENGTH_ERROR_MSG;
           end;
        end;
        dtmMain.UpdateCMD_STATUS(frmMain.nG_CurActiveConn,sG_SeqNo,'E',sG_ErrorCode,sG_ErrorMsg);

        writeERrorLog;

        Exit;
      end;
      sG_TransactionNum := getTransNum;
      if frmMain.bG_ShowUI then
        Synchronize(createUiListItem);  //0109
  //      G_ListItem := frmMain.G_HandleUIThread.createListItem(sG_TransactionNum,sG_CommandNo, sG_Operator);

      L_StrList := TUStr.ParseStrings(sL_DetailCommandIDs,',');

      if L_StrList <> nil then
      begin
        for ii:=0 to L_StrList.Count -1 do
        begin
          sG_CurActiveLowLevelCmd := L_StrList.Strings[ii];
          if (sG_CommandNo='A1') and (sG_CurActiveLowLevelCmd='48') then//執行開卡之set ZipCode 指令時,要做此判斷
          begin
            Synchronize(setIfSetZipCode); //0109
            if (not bG_SetZipCode) or (sG_ZipCode='')  then
            begin
              continue;
            end;
          end;

          if (sG_CurActiveLowLevelCmd='100') then//執行set Credit Limit 指令時,要做此判斷
          begin
            if nG_CreditLimit=0 then
              Synchronize(getDefaultCreditLimit);
  //            nG_CreditLimit := frmMain.nG_DefaultCreditLimit; //0109
          end;

          bL_ProfuctFieldFormatError := false;
          if (sG_CommandNo='B7') then //(收視權限複製)
          begin
            //update send_nagra set notes='111111111111^20020801^20020901,222222222222^20020901^20021231~3333333333;4444444444';
            sL_Notes := sG_Notes;
            {
            傳入多卡號,多頻道(欄位Notes)=>將所有頻道 grant 給所有卡號.
            各 product 間以','區隔. 各卡號間以';'區隔.product data與卡號 data以'~'區隔.(先product data,後卡號 data)
            各 product 的info 包含三個部分=>product id, 收視起日,收視迄日.而這三個部分以'^'作為區隔
            Ex: 111111111111^20020801^20020901,222222222222^20020901^20021231~3333333333;4444444444
            }
            L_NotesStrList := TUStr.ParseStrings(sL_Notes,'~');
            if L_NotesStrList<>nil then
            begin
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
            end;
          
            if bL_ProfuctFieldFormatError then
            begin
              sG_ErrorCode:= NOTES_FIELD_FORMAT_ERROR;
              sG_ErrorMsg := NOTES_FIELD_FORMAT_ERROR_MSG;
              dtmMain.UpdateCMD_STATUS(frmMain.nG_CurActiveConn,sG_SeqNo,'E',sG_ErrorCode,sG_ErrorMsg);
              writeERrorLog;
              break;
            end
            else
            begin
              for jj :=0 to L_ProductsStrList.Count -1 do
              begin
                L_ProductInfoStrList := TUStr.ParseStrings(L_ProductsStrList.Strings[jj],'^');
                if (L_ProductInfoStrList<>nil) and (L_ProductInfoStrList.Count=3) then
                begin
                  sG_ImsProductID := L_ProductInfoStrList.Strings[0];
                  sG_ImsProductID := TUstr.AddString(sG_ImsProductID,'0',false,IMS_PRODUCT_ID_LENGTH);//12 碼, 右靠,左補0
                  sG_SubBeginDate := L_ProductInfoStrList.Strings[1];
                  sG_SubEndDate := L_ProductInfoStrList.Strings[2];
                  for kk:=0 to L_CardStrList.Count -1 do
                  begin
                    sG_IccNo := L_CardStrList.Strings[kk];
                    sG_IccNo := TUstr.AddString(sG_IccNo,'0',false,ICC_CARDNO_LENGTH);//10 碼,右靠, 左補0                
                    preSendCommand; //高階指令 B7 比較特殊,要傳送很多次指令,所以由此呼叫 function
                  end
                end
                else
                begin
                  sG_ErrorCode:= NOTES_FIELD_FORMAT_ERROR;
                  sG_ErrorMsg := NOTES_FIELD_FORMAT_ERROR_MSG;
                  dtmMain.UpdateCMD_STATUS(frmMain.nG_CurActiveConn,sG_SeqNo,'E',sG_ErrorCode,sG_ErrorMsg);
                  writeERrorLog;
    //              continue; =>用 continue 會有問題
                  break;
                end;
              end;

            end;
          end
          else if (sG_CommandNo='C1') then // (解除配對)
          begin
             {
             各ICC卡號間間以','區隔
             }
            //update send_nagra set notes='1111111111,2222222222,3333333333';
            sL_Notes := sG_Notes;
            L_NotesStrList := TUStr.ParseStrings(sL_Notes,',');
            if L_NotesStrList<> nil then
            begin
              for jj:=0 to L_NotesStrList.Count -1 do//表示有這麼多組 ICC 卡號資料
              begin
                sG_StbNo := '';
                sG_StbNo := TUstr.AddString(sG_StbNo,'0',false,BOXNO_ID_LENGTH-4);//14 碼,右靠, 左補0,右補4個空白
                sG_StbNo := TUstr.AddString(sG_StbNo,' ',true,BOXNO_ID_LENGTH);//14 碼,右靠, 左補0,右補4個空白

                sG_IccNo := L_NotesStrList.Strings[jj];
                sG_IccNo := TUstr.AddString(sG_IccNo,'0',false,ICC_CARDNO_LENGTH);//10 碼,右靠, 左補0
                preSendCommand;
              end;
            end;
          end
          else if (sG_CommandNo='B2') or (sG_CommandNo='B4') or (sG_CommandNo='B5') then //(B2:關頻道-多組, B4:暫停頻道-多組, B5:恢復頻道-多組)
          begin
             {
             V各 product 間以','區隔
             }
            //update send_nagra set notes='111111111111,222222222222,333333333333,555555555555';
            sL_Notes := sG_Notes;
            L_NotesStrList := TUStr.ParseStrings(sL_Notes,',');
            if L_NotesStrList<>nil then
            begin
              for jj:=0 to L_NotesStrList.Count -1 do//表示有這麼多組 product 資料
              begin
                sG_ImsProductID := L_NotesStrList.Strings[jj];
                sG_ImsProductID := TUstr.AddString(sG_ImsProductID,'0',false,IMS_PRODUCT_ID_LENGTH);//12 碼, 右靠,左補0

                preSendCommand;
              end;
            end;
          end
          else if (sG_CommandNo='B6') then //(延展期限)
          begin
            //update send_nagra set notes='A~111111111111,A~222222222222,A~333333333333,B~555555555555~20021031';
            sL_Notes := sG_Notes;
            {
            V各 product 間以','區隔, product與期間之間以'~'區隔;
            第一碼是'A'者,表示共用此指令之subscription_end_datee,
            第一碼是'B'者,表示自己有自己的 subscription_end_date.
            Ex:A~1,A~2,A~3,B~5~20020205=>product 1,2,3以指令之 subscription_end_date為日期;
            而product 5 的 subscription_end_date 為20020205
            }
            L_NotesStrList := TUStr.ParseStrings(sL_Notes,',');
            if L_NotesStrList <> nil then
            begin
              for jj:=0 to L_NotesStrList.Count -1 do//表示有這麼多組 product 資料
              begin
                sL_ProductInfo := L_NotesStrList.Strings[jj];
                L_ProductsStrList := TUStr.ParseStrings(sL_ProductInfo,'~');
                if L_ProductsStrList<> nil then
                begin
                  if (L_ProductsStrList.Strings[0]='A') and (L_ProductsStrList.Count=2) then //ex : A~1
                  begin
                    sG_ImsProductID := L_ProductsStrList.Strings[1];
                    sG_ImsProductID := TUstr.AddString(sG_ImsProductID,'0',false,IMS_PRODUCT_ID_LENGTH);//12 碼, 右靠,左補0

                    sG_SubEndDate := sG_MisSubEndDate;

                    preSendCommand;
                  end
                  else if (L_ProductsStrList.Strings[0]='B') and (L_ProductsStrList.Count=3) then //ex: B~5~20021231
                  begin
                    sG_ImsProductID := L_ProductsStrList.Strings[1];
                    sG_SubEndDate := L_ProductsStrList.Strings[2];
                    preSendCommand;
                  end;
                end;
              end;
            end;
          end
          else if (sG_CommandNo='B1') then //(開頻道-多組)
          begin
            //update send_nagra set notes='A~111111111111,A~222222222222,A~333333333333,B~555555555555~20020201~20020205';
            sL_Notes := sG_Notes;

            {
            各 product 間以,區隔, product與期間之間以~區隔;
            第一碼是A者,表示共用此指令之subscription_begin_date 與 subscription_end_datee,
            第一碼是B者,
            表示自己有自己的 subscription_begin_date 與 subscription_end_date.
            Ex:A~1,A~2,A~3,B~5~20020201~20020205=>product 1,2,3以指令之 subscription_begin_date, subscription_end_date為日期;
            而product 5 的subscription_begin_date 為20020201, subscription_end_date 為20020205
            }


            L_NotesStrList := TUStr.ParseStrings(sL_Notes,',');
            if L_NotesStrList<>nil then
            begin
              for jj:=0 to L_NotesStrList.Count -1 do//表示有這麼多組 product 資料
              begin
                sL_ProductInfo := L_NotesStrList.Strings[jj];
                L_ProductsStrList := TUStr.ParseStrings(sL_ProductInfo,'~');
                if L_ProductsStrList<> nil then
                begin
                  if (L_ProductsStrList.Strings[0]='A') and (L_ProductsStrList.Count=2) then //ex : A~1
                  begin
                    sG_ImsProductID := L_ProductsStrList.Strings[1];
                    sG_ImsProductID := TUstr.AddString(sG_ImsProductID,'0',false,IMS_PRODUCT_ID_LENGTH);//12 碼, 右靠,左補0

                    sG_SubBeginDate := sG_MisSubBeginDate;
                    sG_SubEndDate := sG_MisSubEndDate;

                    preSendCommand;
                  end
                  else if (L_ProductsStrList.Strings[0]='B') and (L_ProductsStrList.Count=4) then //ex: B~5~20020201~20020205
                  begin
                    sG_ImsProductID := L_ProductsStrList.Strings[1];
                    sG_ImsProductID := TUstr.AddString(sG_ImsProductID,'0',false,IMS_PRODUCT_ID_LENGTH);//12 碼, 右靠,左補0
                    sG_SubBeginDate := L_ProductsStrList.Strings[2];
                    sG_SubEndDate := L_ProductsStrList.Strings[3];
                    preSendCommand;
                  end;
                end;
              end;
            end;
          end
          else if (sG_CommandNo='B9') then //(refresh)
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

            if sG_CurActiveLowLevelCmd='2' then
            begin
              sL_Notes := sG_Notes;        
              L_NotesStrList := TUStr.ParseStrings(sL_Notes,',');
              if L_NotesStrList<> nil then
              begin
                for jj:=0 to L_NotesStrList.Count -1 do//表示有這麼多組 product 資料
                begin
                  sL_ProductInfo := L_NotesStrList.Strings[jj];
                  L_ProductsStrList := TUStr.ParseStrings(sL_ProductInfo,'~');
                  if L_ProductsStrList<> nil then
                  begin
                    if (L_ProductsStrList.Strings[0]='A') and (L_ProductsStrList.Count=2) then //ex : A~1
                    begin
                      sG_ImsProductID := L_ProductsStrList.Strings[1];
                      sG_ImsProductID := TUstr.AddString(sG_ImsProductID,'0',false,IMS_PRODUCT_ID_LENGTH);//12 碼, 右靠,左補0

                      sG_SubBeginDate := sG_MisSubBeginDate;
                      sG_SubEndDate := sG_MisSubEndDate;

                      preSendCommand;
                    end
                    else if (L_ProductsStrList.Strings[0]='B') and (L_ProductsStrList.Count=4) then //ex: B~5~20020201~20020205
                    begin
                      sG_ImsProductID := L_ProductsStrList.Strings[1];
                      sG_ImsProductID := TUstr.AddString(sG_ImsProductID,'0',false,IMS_PRODUCT_ID_LENGTH);//12 碼, 右靠,左補0
                      sG_SubBeginDate := L_ProductsStrList.Strings[2];
                      sG_SubEndDate := L_ProductsStrList.Strings[3];
                      preSendCommand;
                    end;
                  end;
                end;
              end;
            end
            else
            begin
              preSendCommand;
            end;
          end
          else
          begin
            preSendCommand;

          end;
        end;
      end;
    except
      on ESocketError do Showmessage('網路無法使用....');
    else
//      Showmessage('執行緒錯誤....');
    end;

end;

constructor TSendCommand.Create;
begin
//    CRS := TCriticalSection.Create;

    inherited Create(false);
end;

constructor TSendCommand.Create(I_AdoQryCmdLog: TADOQuery);
begin

    G_AdoQryRealCmd := I_AdoQryCmdLog;
    sG_ErrorCode:=NO_ERROR; //預設為 -- 成功訊息代碼
    sG_ErrorMsg :='';
    nG_CommandStatus :=0;
    nG_Timer := 0;
    bG_MakeSmsCommand := false;
    {
    flag := false;  //FTP命令執行成功與否 1.成功 0.失敗
    runFTPFlag := false;
    makeResponseFlag := false;
    RecSeqNo := '';
    }
    {
    FreeOnTerminate := True;
    Priority := tpnormal;
    }
//    CRS := TCriticalSection.Create;

    inherited Create(false);
    
end;

procedure TSendCommand.changeFetchingAndSending;
begin
    frmMain.G_CommandInfoStrList.Clear;
    frmMain.bG_RunSendingCommandThread := false;
    frmMain.bG_RunFetchCommandThread := true;

end;

function TSendCommand.getTransNum: String;
var
    sL_TransactionNum : String;
begin
    sL_TransactionNum := IntToStr(nG_CompCode) + TUstr.AddString(sG_SeqNo,'0',false,SEQ_NUM_LENGTH);
    sL_TransactionNum := TUstr.AddString(sL_TransactionNum,'0',false,SEND_TRANSACTION_NUM_LENGTH);
    result := sL_TransactionNum;

end;

procedure TSendCommand.setTimerStatue2False;
begin
    frmMain.timSendTestCommand.Enabled := false;
end;

procedure TSendCommand.setTimerStatue2True;
begin
    frmMain.timSendTestCommand.Enabled := True;
end;



procedure TSendCommand.updateFetchCommandThreadToFalse;
begin
    frmMain.bG_RunFetchCommandThread := false;
end;

procedure TSendCommand.updateCmdInfo;
begin
    frmMain.G_CommandInfoStrList.Strings[nG_CmdStrListPos] := '';
    (frmMain.G_CommandInfoStrList.Objects[nG_CmdStrListPos] as TCmdInfoObject).Free;    
//  frmMain.G_CommandStrList.Delete(nG_CmdStrListPos);

end;

procedure TSendCommand.getSeqNo;
begin
    sG_SeqNo := frmMain.G_CommandInfoStrList.Strings[nG_CmdStrListPos];
end;

procedure TSendCommand.getCmdInfoObj;
begin
    G_CmdInfoObj := (frmMain.G_CommandInfoStrList.Objects[nG_CmdStrListPos] as TCmdInfoObject);
end;

procedure TSendCommand.createUiListItem;
begin
    G_ListItem := frmMain.G_HandleUIThread.createListItem(sG_TransactionNum,sG_CommandNo, sG_Operator);
end;

procedure TSendCommand.getDefaultCreditLimit;
begin
    nG_CreditLimit := frmMain.nG_DefaultCreditLimit;
end;

procedure TSendCommand.setIfSetZipCode;
begin
    bG_SetZipCode := frmMain.bG_SetZipCode;
end;

procedure TSendCommand.setRunSendingCommandThreadFlag;
begin
    bG_RunSendingCommandThread := frmMain.bG_RunSendingCommandThread;
end;

procedure TSendCommand.setCmdInfoStrList;
begin
    G_CommandInfoStrList := frmMain.G_CommandInfoStrList;
end;

procedure TSendCommand.getSocket;
begin
    G_Socket := frmMain.ClientSocket1;
end;


procedure TSendCommand.getSocketStream;
begin
    G_WStream := TWinSocketStream.Create(G_Socket.Socket, 120000);
end;

function TSendCommand.Write(I_Buf: Pointer; I_Count: Integer): Integer;
var
    nL_Idx: Integer;
    L_ErrorInfoStrList : TStringList;
    sL_ErrorInfoFileName : String;
begin


    nL_Idx := 0;
    while (nL_Idx < I_Count) do //0110
//    if (not Terminated) and (G_Socket.Active) then
    begin
      try
        try
          nL_Idx := nL_Idx + G_WStream.Write(Increment(I_Buf, nL_Idx)^, I_Count-nL_Idx); //0110

//          nL_Idx := G_WStream.Write((I_Buf)^, I_Count);
        except
          on E: Exception do
          begin

            sL_ErrorInfoFileName := 'c:\Error8.txt';
            L_ErrorInfoStrList := TStringList.Create;
            if FileExists(sL_ErrorInfoFileName) then
              L_ErrorInfoStrList.LoadFromFile(sL_ErrorInfoFileName);
            if L_ErrorInfoStrList.Count>500 then
              L_ErrorInfoStrList.Clear;
            L_ErrorInfoStrList.Insert(0,'   Msg:' +E.Message +', HelpContext:' + IntToStr(E.HelpContext));
            L_ErrorInfoStrList.Insert(0,'TSendCommand.write error:' + DateTimeToStr(now));
            L_ErrorInfoStrList.SaveToFile(sL_ErrorInfoFileName);
            L_ErrorInfoStrList.Free;

            buildConnection(frmMain.ClientSocket1 );
            Write(I_Buf, I_Count);

          end;

      //    raise;
        end;
      finally

      end;  
    end;
    result := nL_Idx;
end;

end.
