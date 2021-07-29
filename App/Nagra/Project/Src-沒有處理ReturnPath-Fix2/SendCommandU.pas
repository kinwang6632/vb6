unit SendCommandU;

interface

uses
  Windows, Classes, Forms, Comctrls, Sysutils, Dialogs, Sockets,
  IniFiles, ConstObjU, Ustru, UdateTimeu, CsIntf;

  procedure OnTimerStartAction;
  procedure BuildSendCmdConnection(const aSocket: TTcpClient);overload;
  procedure UpdateCmdStatus;
  procedure ParseCommandData(I_CmdInfoObj:TCmdInfoObject);
  procedure WriteTimeStamp;
  procedure SendCommand(sI_CommandNo : String);
  procedure WriteSendCmdLog(sI_CommandNo: String);overload;
  procedure preSendCommand( I_CmdInfoObj: TCmdInfoObject; sI_IccNo,
    sI_ImsProductID, sI_SubBeginDate, sI_SubEndDate: String);
  procedure MakeSmsCommand(I_CmdInfoObj:TCmdInfoObject; sI_IccNo,
    sI_ImsProductID,sI_SubBeginDate, sI_SubEndDate: String; I_IrdInfoObj: TIrdInfoObj);
  function GenRealIccNo(const sI_SourceIccNo : String):String;
  function GenRealStbNo(const sI_SourceStbNo : String):String;
  function GenRealImsproductID(sI_SourceImsProductID : String):String;
  function GetDetailCommandIDs(sI_HighLevelCommand:String):String;
  function GetCommandType(const nI_CommandID : Integer) : String;
  function GetTransNum(nI_CompCode:Integer; sI_SeqNo : String):String;
  function GetFormatIpAddr(sI_SrcIpAddr:String):String;
  function IsDataOK(sI_ImsProductID, sI_IccNo, sI_StbNo :String;
    var nI_ResultFlag:Integer) : boolean;
      
  { �ª� SetopBox ���l�Ϊ����O }
  function transIrdInfo(nI_CompCode:Integer; sI_MisTrdCmdID,
    sI_MisIrdCmdData: String; var  L_IrdInfoObj : TIrdInfoObj ):boolean;
  { �s�� SetopBox ���l�Ϊ����O }
  function TransIrdInfo2(const aCompCode: Integer; aMisIrdCmdId,
    aMisIrdCmdData: String; var aIrdInfoObj: TIrdInfoObj ): Boolean;

  function getSmsCommandRootHeader(bI_GenTransactionNum: Boolean;
    nI_CompCode, nI_CommandID : Integer; sI_CommandType,
    sI_SeqNo: String) : String;
  function getCommandSection(sI_CommandType, sI_BDate,
    sI_EDate,sI_SerialNo:String) : String;
  function getCommandBody(I_CmdInfoObj: TCmdInfoObject; nI_CommandID: Integer;
    I_IrdInfoObj : TIrdInfoObj; sI_SubBeginDate, sI_SubEndDate,
    sI_ImsProductID: String) : String;



implementation

uses frmMainU, frmSysParamU, dtmMainU, HandleUIU, CbUtils;

{ ---------------------------------------------------------------------------- }

function BinStrToInt(const ABinary: String): Integer;
var
  aIndex: Integer;
begin
  Result := 0;
  for aIndex := 1 to Length(ABinary) do
  begin
    Result := Result shl 1 or (Byte(ABinary[aIndex]) and 1);
  end;
end;

{ ---------------------------------------------------------------------------- }

function IntToBin(aValue: cardinal): string;
var
  aIndex: Integer;
begin
  SetLength( Result, 32 );
  for aIndex := 1 to 32 do
  begin
    if (( aValue shl ( aIndex - 1 ) ) shr 31 ) = 0 then
      Result[aIndex] := '0'
    else
      Result[aIndex] := '1';
  end;
end;

{ ---------------------------------------------------------------------------- }

function getCommandBody(I_CmdInfoObj: TCmdInfoObject;
  nI_CommandID: Integer; I_IrdInfoObj: TIrdInfoObj; sI_SubBeginDate,
  sI_SubEndDate, sI_ImsProductID: String): String;
var
    sL_Result, sL_CommandID : String;
    sL_StuNumber, sL_ImsProductID : String;
    sL_BDate, sL_EDate, sL_ZipCode, sL_CreditMode, sL_Credit : String;
    sL_ThresholdCredit, sL_EventName, sL_EventNameLength, sL_Price, sL_CCNum1 : String;
    sL_IpAddr, sL_CCPort, sL_CbDate, sL_CbTime : String;
    sL_CallFreq, sL_DateFirstCall, sL_CreditLimit : String;
    sL_PhoneNum1, sL_PhoneNum2, sL_PhoneNum3, sL_SrcTransNum : String;
    sL_IrdCmdID, sL_IrdOperation, sL_IrdDataLength, sL_IrdData : String;
    sL_CleanDate, sL_ConditionDate, sL_CollectDate, sL_Mode: String;
begin
    {
      �� function �ΨӲզX�X�U���O�� parameters.
      ex: EMM command ���ѼƽаѦ� SMS GATEWAY SPECIFICATION �Ĥ���
    }
    sL_CommandID := TUstr.AddString( IntToStr( nI_CommandID ), '0', False,
      COMMAND_ID_LENGTH );

    sL_StuNumber := genRealStbNo( Trim( I_CmdInfoObj.STB_NO ) );

    sL_ImsProductID := sI_ImsProductID;

    sL_BDate := TUstr.replaceStr( sI_SubBeginDate,'/','' );
    sL_EDate := TUstr.replaceStr( sI_SubEndDate,'/','' );

    sL_ZipCode := I_CmdInfoObj.ZIP_CODE;

    sL_CreditMode := Format('%.2d', [StrToIntDef(
      I_CmdInfoObj.CREDIT_MODE,0 )] );

    {
      sL_CreditMode :
        01 => ADD
        02 => SUBTRACT
        03 => SET CREDIT
        04 => SET BALANCE
        05 => SUB OFFSET
    }

    sL_Credit := Format('%.5d', [StrToIntDef( I_CmdInfoObj.CREDIT,0 )]) + '00';

    sL_ThresholdCredit := Format( '%.7d', [
      StrToIntDef( I_CmdInfoObj.THRESHOLD_CREDIT, 0 )] );

    {
      �b STB ���l��, �ʶR�M��|��ܶýX ?? ��������
    }

    sL_EventName := TUstr.AddString( I_CmdInfoObj.EVENT_NAME , '0', True,
      EVENT_NAME_LENGTH ) + '00';

    sL_EventNameLength := Format( '%.2d',[Length( I_CmdInfoObj.EVENT_NAME )] );


    {
      Price �`���� 5 --> 00000, �᭱2�쬰�p���I,
      Ex: ����15��, �h Price ������ 01500,
          ����22.5��, �h Price ������ 02250,
      �]�����椣�|���p���I, �ҥH��2�X�@�߸�0
    }
    
    sL_Price := Format( '%.3d', [StrToIntDef( I_CmdInfoObj.PRICE, 0 )] ) + '00';


    sL_IpAddr := getFormatIpAddr( I_CmdInfoObj.IP_ADDR );
    sL_CCPort := Format( '%.5d', [StrToIntDef( I_CmdInfoObj.CC_PORT,0 )] );
    sL_CbDate := I_CmdInfoObj.CALLBACK_DATE ;
    sL_CbTime := I_CmdInfoObj.CALLBACK_TIME;
    sL_CallFreq := I_CmdInfoObj.CALLBACK_FREQUENCY;


    sL_DateFirstCall := I_CmdInfoObj.FIRST_CALLBACK_DATE;
    sL_DateFirstCall := TUstr.replaceStr( sL_DateFirstCall,'/','' );

    sL_CreditLimit := Format('%.7d', [ StrToIntDef(
      I_CmdInfoObj.CREDIT_LIMIT, 0 )] );

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

    //down, �N IRD_DATA ��0�ɺ�96��bytes

    sL_IrdData := TUstr.AddString( sL_IrdData, '0', True,
      IRD_COMMAND_DATA_LENGTH );

    //down, �q�ܸ��X���ഫ => �o�˪��{���g�k,��simulator �L�k work; ���s�u���h���ծ�,�o�Sok.
    sL_CCNum1 := TUstr.AddString(I_CmdInfoObj.CC_NUMBER_1,chr(32),true,CALL_COLLECTOR_PHONE_NUM_LENGTH);

    sL_PhoneNum1 := TUstr.AddString(I_CmdInfoObj.PHONE_NUM_1,chr(32),true,AUTHORIZED_PHONE_NUM_LENGTH);
    sL_PhoneNum2 := TUstr.AddString(I_CmdInfoObj.PHONE_NUM_2,chr(32),true,AUTHORIZED_PHONE_NUM_LENGTH);
    sL_PhoneNum3 := TUstr.AddString(I_CmdInfoObj.PHONE_NUM_3,chr(32),true,AUTHORIZED_PHONE_NUM_LENGTH);
    //up, �q�ܸ��X���ഫ => �o�˪��{���g�k,��simulator �L�k work; ���s�u���h���ծ�,�o�Sok.


   sL_CleanDate := TUstr.replaceStr( I_CmdInfoObj.CLEANUP_DATE, '/', '' );
   sL_ConditionDate := TUstr.replaceStr( I_CmdInfoObj.CONDITION_DATE, '/', '' );
   sL_CollectDate := TUstr.replaceStr( I_CmdInfoObj.COLLECT_DATE, '/', '' );
   sL_Mode := '2'; 

    sL_SrcTransNum := frmMain.sG_SrcTransNum;

    case nI_CommandID of
      2 :
        sL_Result := sL_CommandID + sL_ImsProductID + sL_BDate + sL_EDate;
      3 :
        sL_Result := sL_CommandID + sL_ImsProductID + sL_EDate;
      4,5,6 :
        sL_Result := sL_CommandID + sL_ImsProductID;
      7, 14, 15, 20, 21, 50, 51, 53, 62, 71, 105, 110, 111, 1002:
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
      49 : //�����O�����D
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
      92:
        sL_Result := sL_CommandID + sL_Mode + sL_CleanDate;
      96:
        sL_Result := sL_CommandID + sL_CleanDate + sL_ConditionDate;
      97:
        sL_Result := sL_CommandID + sL_CollectDate;
      100 :
        sL_Result := sL_CommandID + sL_CreditLimit;
      101 :  //�����O�����D
        sL_Result := sL_CommandID + sL_PhoneNum1 + sL_PhoneNum2 + sL_PhoneNum3;
      1000 :
        sL_Result := sL_CommandID + sL_SrcTransNum + TUstr.AddString('','0',true,IMS_PRODUCT_ID_LENGTH)  + TUstr.AddString('','0',true,IMS_PRODUCT_ID_LENGTH);
    end;

    Result := sL_Result;
end;

{ ---------------------------------------------------------------------------- }

function getSmsCommandRootHeader(bI_GenTransactionNum:boolean; nI_CompCode,
  nI_CommandID: Integer; sI_CommandType, sI_SeqNo: String): String;
var
  sL_Result : String;
begin
  if bI_GenTransactionNum then
    frmMain.sG_TransactionNum := getTransNum( nI_CompCode, sI_SeqNo )
  else
    frmMain.sG_TransactionNum := EmptyStr;
  sL_Result := frmMain.sG_TransactionNum + sI_CommandType;
  sL_Result := sL_Result + frmMain.sG_SourceID + frmMain.sG_DestID + frmMain.sG_MopPPID;
  sL_Result := sL_Result + TUdateTime.GetPureDateStr(date,'');
  Result := sL_Result;
end;

{ ---------------------------------------------------------------------------- }

function GetCommandType(const nI_CommandID: Integer): String;
begin
  Result := '-1';
  if ( EMM_CommandBeginID <= nI_CommandID ) and
     ( nI_CommandID <= EMM_CommandEndID ) then
    Result := frmMain.sG_EmmCommandType
  else
    if ( CONTROL_CommandBeginID <= nI_CommandID ) and
       ( nI_CommandID <= CONTROL_CommandEndID ) then
    Result := frmMain.sG_ControlCommandType
  else
    if ( FEEDBACK_CommandBeginID <= nI_CommandID ) and
       ( nI_CommandID <= FEEDBACK_CommandEndID ) then
    Result := frmMain.sG_FeedbackCommandType
  else
    if ( PRODUCT_CommandBeginID <= nI_CommandID ) and
       ( nI_CommandID <=PRODUCT_CommandEndID ) then
    Result := frmMain.sG_ProductCommandType
  else
    if ( OPERATION_CommandBeginID <= nI_CommandID ) and
       ( nI_CommandID <= OPERATION_CommandEndID ) then
    Result := frmMain.sG_OperationCommandType;
end;

{ ---------------------------------------------------------------------------- }

function getCommandSection(sI_CommandType,sI_BDate, sI_EDate,
  sI_SerialNo: String): String;
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

procedure MakeSmsCommand(I_CmdInfoObj:TCmdInfoObject; sI_IccNo,
  sI_ImsProductID, sI_SubBeginDate, sI_SubEndDate: String; I_IrdInfoObj: TIrdInfoObj);
var
  nL_CompCode : Integer;
  sL_SmsCommandRootHeader, sL_CommandSection : String;
  sL_CommandBody, sL_CommandType : String;
  sL_IccNo, sL_CurActiveLowLevelCmd : String;
begin
  try
    nL_CompCode := StrToIntDef( I_CmdInfoObj.COMPCODE, 0 );
    sL_IccNo := sI_IccNo;
    frmMain.bG_SC_MakeSmsCommand := False;
    sL_CurActiveLowLevelCmd := frmMain.sG_CurActiveLowLevelCmd;
    sL_CommandType := getCommandType( StrToInt( sL_CurActiveLowLevelCmd ) );
    sL_SmsCommandRootHeader := getSmsCommandRootHeader( True, nL_CompCode,
      StrToInt( sL_CurActiveLowLevelCmd ), sL_CommandType, I_CmdInfoObj.SEQNO );
    sL_CommandSection := getCommandSection( sL_CommandType,
      frmMain.sG_BroadcastStartDate, frmMain.sG_BroadcastEndDate, sL_IccNo );
    sL_CommandBody := getCommandBody( I_CmdInfoObj,
      StrToInt( sL_CurActiveLowLevelCmd ), I_IrdInfoObj, sI_SubBeginDate,
      sI_SubEndDate, sI_ImsProductID );
    frmMain.sG_SC_FullCommand := sL_SmsCommandRootHeader +
      sL_CommandSection +sL_CommandBody;
  except
    frmMain.bG_SC_MakeSmsCommand := False;
    Exit;
  end;
  frmMain.bG_SC_MakeSmsCommand := True;
end;

{ ---------------------------------------------------------------------------- }

procedure SendCommand(sI_CommandNo : STring);
var
  aSocket : TTcpClient;
  nL_Words : Integer;
  ii, nL_Len : Integer;
  L_TmpRecord : TTempRecord;
  L_DeviceData : TDeviceData;
  sL_TransactionNum, sL_LowLevelCmdID, sL_TmpFullCommand : String;
  aCheck: Integer;
begin

    if frmMain.sG_SC_FullCommand='' then Exit;

    frmMain.bG_SC_SendComm := true;

    //down, �ǰecommand
    try
    begin

      sL_TmpFullCommand := frmMain.sG_SC_FullCommand;
      frmMain.sG_SC_TransNumForUpdateCmdStatus := Copy(sL_TmpFullCommand,1,SEND_TRANSACTION_NUM_LENGTH);
      sL_TransactionNum := Copy(sL_TmpFullCommand,1, SEND_TRANSACTION_NUM_LENGTH);

      nL_Len := length(sL_TmpFullCommand);//ok..
      //nL_Len := 1 + 1 + length('EMGR_EME_SVR') + length(sL_FullCommand); //testing..=> ��ꤣ�T�w�n���n�e sL_CmdObjName

      L_TmpRecord := frmMain.GetBinaryVal(nL_Len,2,false);
      L_DeviceData.len[0] := L_TmpRecord.TempAry[0];
      L_DeviceData.len[1] := L_TmpRecord.TempAry[1];

      FillChar(L_DeviceData.data, length(L_DeviceData.data), char(0)); { initialize the buffer }

      for ii:=1 to length(sL_TmpFullCommand) do
      begin
        L_DeviceData.data[ii-1] := sL_TmpFullCommand[ii];
      end;


      aSocket := frmMain.ControlSocket;
      
      WriteSendCmdLog(sI_CommandNo);
      try
        if frmMain.bG_HasBuildSendCmdConnection then
        begin
          nL_Words := aSocket.SendBuf( L_DeviceData, nL_Len+2 );
          CodeSite.SendMsg( csmObject,
            'Command Id=' + sI_CommandNo + ', Body=' + sL_TmpFullCommand );
          CodeSite.AddSeparator;
        end;
      except
        on E: Exception do
        begin
          CodeSite.SendError( 'Data Send Error ' + E.Message );
          frmMain.bG_SC_buildSendCmdConnection := False;
          sL_LowLevelCmdID := Copy( sL_TmpFullCommand, 61, 4 );
          frmMain.showLogOnMemo( SEND_CMD_ERROR_FLAG,
            '�ǰe���O' + sL_LowLevelCmdID + ' �� Nagra.�{���Y�N����,�M�᭫�s�Ұ�. '); //���槹����{����, �{���|������
          Abort;
        end;
      end;

      aCheck := StrToIntDef( Copy( sL_TransactionNum, 1, 3 ), 0 );

//***************
//      receiveResponse; //�����^�� ...2003/11/03 �N�� mark ��
//***************

    end
    except
      on e: Exception do
      begin
        CodeSite.SendError( 'Data Send Error ' + E.Message );
        frmMain.bG_SC_SendComm := false;
        if frmMain.bG_ShowUI then
          HandleUIU.UpdateListItemStatus(nil,4,sL_TransactionNum,frmMain.sG_SC_ErrorCode,frmMain.sG_SC_ErrorMsg);
        frmMain.sG_SC_ErrorCode := SEND_COMMAND_ERROR;
        frmMain.sG_SC_ErrorMsg := SEND_COMMAND_ERROR_MSG;
        frmMain.sG_SC_CmdStatus := 'E';
        frmMain.nG_SC_ConnectionID := frmMain.nG_CurActiveConn;
        updateCmdStatus;
        Exit;
      end;
    end;
    //up, �ǰecommand
    if frmMain.bG_ShowUI then
      HandleUIU.UpdateListItemStatus(nil,1,sL_TransactionNum,'','');

    if  ( aCheck <> StrToIntDef( TESTING_CMD_COMP_CODE, 0 ) ) and
        ( aCheck <> StrToIntDef( TESTING_CMD_COMP_CODE2, 0 ) ) then
        Delay( 800 );

end;

function getFormatIpAddr(sI_SrcIpAddr: String): String;
var
  L_StrList : TStringList;
  AIndex : Integer;
begin
  Result := '';
  if sI_SrcIpAddr = '' then Exit;
  L_StrList := TUStr.ParseStrings( sI_SrcIpAddr, '.' );
  try
    for AIndex := 0 to L_StrList.Count -1 do
      L_StrList.Strings[AIndex] := TUstr.AddString( L_StrList.Strings[AIndex],
        '0', False, 3 );
    for AIndex := 0 to L_StrList.Count - 1 do
    begin
      if AIndex = 0 then
        Result := L_StrList.Strings[AIndex]
      else
        Result := Result + '.' + L_StrList.Strings[AIndex];
    end;
  finally
    L_StrList.Free;
  end;
end;

function getDetailCommandIDs(sI_HighLevelCommand: String): String;
var
  L_IniFile : TIniFile;
  AFileName: String;
begin

  Result := '';

  AFileName := IncludeTrailingPathDelimiter(
     ExtractFileDir( Application.ExeName ) );

  if frmMain.sG_NewIniFileName = '' then
    AFileName := AFileName + CABLE_NAGRA_INI_FILE_NAME
  else
    AFileName:= AFileName +  frmMain.sG_NewIniFileName;

  {
  sL_ExeFileName := Application.ExeName;
  sL_ExePath := ExtractFileDir(sL_ExeFileName);
    if frmMain.sG_NewIniFileName='' then
      sL_IniFileName := sL_ExePath + '\' + CABLE_NAGRA_INI_FILE_NAME
    else
      sL_IniFileName := sL_ExePath + '\' + frmMain.sG_NewIniFileName;
  }
  
  if FileExists( AFileName ) then
  begin
    L_IniFile := TIniFile.Create( AFileName );
    try
      Result := L_IniFile.ReadString( 'COMMANDMATRIX', sI_HighLevelCommand, '' );
    finally
      L_IniFile.Free;
    end;
  end else
  begin
    MessageDlg( frmMain.getLanguageSettingInfo( 'SendCmd_Thread_Msg_1_Content' ),
      mtWarning, [mbOK], 0 );
  end;

end;


procedure PreSendCommand(I_CmdInfoObj: TCmdInfoObject;
  sI_IccNo, sI_ImsProductID, sI_SubBeginDate, sI_SubEndDate: String);
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

    L_IrdInfoObj := nil;

    if sL_MisIrdCmdID <> EmptyWideStr then
    begin
      sL_MisIrdCmdID := TUstr.AddString(sL_MisIrdCmdID,'0',false,3);
      sL_MisIrdCmdData := I_CmdInfoObj.MIS_IRD_CMD_DATA;
      L_IrdInfoObj := TIrdInfoObj.Create;

      if ( I_CmdInfoObj.STB_FLAG = '0' ) then
      begin
        bL_TransIrdOk := transIrdInfo( nL_CompCode, sL_MisIrdCmdID,
          sL_MisIrdCmdData, L_IrdInfoObj );
      end else
      begin
        bL_TransIrdOk := TransIrdInfo2( nL_CompCode, sL_MisIrdCmdId,
          sL_MisIrdCmdData, L_IrdInfoObj );
      end;
    end;

    if bL_TransIrdOk then
    begin
      //����SMS Command
      MakeSmsCommand( I_CmdInfoObj, sI_IccNo, sI_ImsProductID,
        sI_SubBeginDate, sI_SubEndDate, L_IrdInfoObj );

      if frmMain.bG_ShowUI then
        HandleUIU.CreateListItem( frmMain.sG_TransactionNum,
          I_CmdInfoObj.HIGH_LEVEL_CMD_ID,
          I_CmdInfoObj.OPERATOR,
          frmMain.sG_CurTimeStamp );

      if frmMain.bG_SC_MakeSmsCommand then
      begin
        Application.ProcessMessages;
        SendCommand( sL_CommandNo );
        Application.ProcessMessages;
      end;
      //else Exit;
    end
    else
    begin
      frmMain.sG_SC_ErrorCode := TRANS_IRD_ERROR;
      frmMain.sG_SC_ErrorMsg := TRANS_IRD_ERROR_MSG;
      frmMain.sG_SC_CmdStatus := 'E';
      frmMain.nG_SC_ConnectionID := frmMain.nG_CurActiveConn;
      dtmMain.UpdateCMD_Status_By_Seq( frmMain.nG_SC_ConnectionID,
        sL_SeqNo,frmMain. sG_SC_CmdStatus, frmMain.sG_SC_ErrorCode,
        frmMain.sG_SC_ErrorMsg );
    end;

    if Assigned( L_IrdInfoObj ) then FreeAndNil( L_IrdInfoObj );

end;

function IsDataOK(sI_ImsProductID, sI_IccNo, sI_StbNo: String;
  var nI_ResultFlag: Integer): Boolean;
begin
  nI_ResultFlag := 0;
  if ( Length( sI_StbNo ) <> 0 ) and
     ( Length( sI_StbNo ) <> BOXNO_ID_LENGTH ) then
    nI_ResultFlag := -1;
  if ( Length( sI_IccNo ) <> 0 ) and
     ( Length( sI_IccNo ) <> ICC_CARDNO_LENGTH ) then
    nI_ResultFlag := -2;
  if ( Length( sI_ImsProductID ) <> 0 ) and
     ( Length( sI_ImsProductID ) <> IMS_PRODUCT_ID_LENGTH ) then
    nI_ResultFlag := -3;
  Result := ( nI_ResultFlag = 0 );
end;

function transIrdInfo(nI_CompCode: Integer; sI_MisTrdCmdID,
  sI_MisIrdCmdData: String; var L_IrdInfoObj: TIrdInfoObj ): Boolean;
var
    bL_Result: Boolean;
    sL_RealMessage, sL_TmpData1, sL_TmpData : String;
    L_StrList: TStringList;
    nL_AscIIValue, ii : Integer;
    sL_RealMailId, sL_BinValue, sL_RealPriority, sL_RealSegmentNumber,
    sL_RealTotalSegment: String;
    sL_IrdCmdID, sL_IrdOperation, sL_IrdDataLength, sL_IrdData: String;
begin
    //�H�U�]�w�аѦҤ�� Hitron Technologies StbCakIrdSpe010302.pdf �H�� Nagravision S.A. SMS Gateway intergace definition V02.06.02
    //���o������P�����n�骺�^�����G���Ӥ@�P...

    bL_Result := true;
    sL_IrdData := sI_MisIrdCmdData;
    case StrToInt(sI_MisTrdCmdID) of
      8: //(�]�w IPPV �q�ʱK�X)
       begin
        { �ª� Setopbox ���O }
        sL_IrdCmdID := '018';
        //sL_IrdOperation :=  '002';
        { �ª����l�S�� IPPV �q�ʱK�X, �ҥH�@�ߤ]�O�]�ˤl�K�X }
        sL_IrdOperation :=  '001';
        sL_IrdDataLength :=  '04';
        sL_IrdData := '';
        //�w�]�� IPPV �q�ʱK�X
        if sI_MisIrdCmdData = '' then
          sI_MisIrdCmdData := '1234';
        for ii:=1 to length(sI_MisIrdCmdData) do
        begin
          nL_AscIIValue := Ord(sI_MisIrdCmdData[ii]);//���X�C�� byte �� ASCII
          sL_IrdData := sL_IrdData + IntToHex(nL_AscIIValue,0); //���X�� ASCII ��16�i���
        end;
       end;
      1: //(�]�w�ˤl�K�X)
       begin
        { �ª� Setop Box }
        sL_IrdCmdID := '018';
        sL_IrdOperation := '001';
        sL_IrdDataLength := '04';
        sL_IrdData := '';
        if sI_MisIrdCmdData = '' then
          sI_MisIrdCmdData := '1234';//�w�]��PIN Code
        for ii:=1 to length(sI_MisIrdCmdData) do
        begin
          nL_AscIIValue := Ord(sI_MisIrdCmdData[ii]);//���X�C�� byte �� ASCII
          sL_IrdData := sL_IrdData + IntToHex(nL_AscIIValue,0); //���X�� ASCII ��16�i���
        end;
       end;
      2: //(�ǰe�T��)
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
          nL_AscIIValue := Ord(frmMain.sG_SC_ActiveMessage[ii]);//���X�C�� byte �� ASCII
          sL_RealMessage := sL_RealMessage + IntToHex(nL_AscIIValue,0); //���X�� ASCII ��16�i���
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
         { �ª� SetopBox �Ϊ����O }
         sL_IrdCmdID := NETWORK_OPERATION_194; //194=>operation �n��'', �B data length=2;
         //sG_IrdCmdID := NETWORK_OPERATION_195; //195=>operation��ܦh�[�n reset STB, data length=4
         sL_IrdOperation := '001';
         sL_IrdData := frmMain.getCompNetworkInfo(nI_CompCode,sL_IrdCmdID);
         sL_IrdDataLength := TUstr.AddString(FloatToStr(length(sL_IrdData)/2),'0',false,2);
         //��ܧ줣�� network id �����...
         if Length( sL_IrdData ) = 0 then bL_Result := False;
       end;
      5: //Master/Slave continuous mode initialisation => �ѦҤ���ɮ�:030527 IRD Master_Slave Implementation_Guideline v0.0.3.pdf
       begin
         sL_IrdCmdID := '199';
         sL_IrdOperation :=  '001';
         L_StrList := TUstr.ParseStrings(sL_IrdData,'~');
         try
           if L_StrList.Count = 3 then
           begin
             {
             sG_IrdData ���榡 => 7~2~24
             L_StrList.Strings[0] => Validation Period
             L_StrList.Strings[1] => Random Period
             L_StrList.Strings[2] => timeout
             }
             sL_IrdData := frmMain.sG_SC_MasterIccNo + L_StrList.Strings[0] + L_StrList.Strings[1] + L_StrList.Strings[2];
             sL_IrdDataLength :=  TUstr.AddString(FloatToStr(length(sL_IrdData)/2),'0',false,2);
           end
           else
           begin
             //��� sG_IrdData ���榡���~...
             sL_IrdData := '';
             sL_IrdDataLength :=  TUstr.AddString(FloatToStr(0),'0',false,2);
             end;
         finally
           L_StrList.Free;
         end;
       end;
      6: //Master/Slave cancellation => �ѦҤ���ɮ�:030527 IRD Master_Slave Implementation_Guideline v0.0.3.pdf
       begin
        sL_IrdCmdID := '199';
        sL_IrdOperation :=  '002';
        sL_IrdData := '';
        sL_IrdDataLength :=  TUstr.AddString(FloatToStr(0),'0',false,2);

       end;
      7: //Master/Slave single shot => �ѦҤ���ɮ�:030527 IRD Master_Slave Implementation_Guideline v0.0.3.pdf
       begin
         sL_IrdCmdID := '199';
         sL_IrdOperation :=  '003';
         L_StrList := TUstr.ParseStrings(sL_IrdData,'~');
         try
           if L_StrList.Count =2 then
           begin
             {
             sG_IrdData ���榡 => 1160052795~24
             L_StrList.Strings[0] => Master SmartCard ID
             L_StrList.Strings[1] => timeout
             }
             //sL_IrdData := L_StrList.Strings[0] + L_StrList.Strings[1];
             sL_IrdData := L_StrList.Strings[0] +
               TUstr.AddString(L_StrList.Strings[1], '0', False, 4 );
             sL_IrdDataLength :=  TUstr.AddString(FloatToStr(length(sL_IrdData)/2),'0',false,2);
           end
           else
           begin
             sL_IrdData := '';
             sL_IrdDataLength :=  TUstr.AddString(FloatToStr(0),'0',false,2);
           end;
         finally
           L_StrList.Free;
         end;  
       end;
    end;
    L_IrdInfoObj.s_IrdCmdID := sL_IrdCmdID;
    L_IrdInfoObj.s_IrdOperation := sL_IrdOperation;
    L_IrdInfoObj.s_IrdData := sL_IrdData;
    L_IrdInfoObj.s_IrdDataLength := sL_IrdDataLength;
    Result := bL_Result;
end;

{ ---------------------------------------------------------------------------- }

function TransIrdInfo2(const aCompCode: Integer; aMisIrdCmdId,
  aMisIrdCmdData: String; var aIrdInfoObj: TIrdInfoObj): Boolean;
var
  aIndex, aIrdCmdId: Integer;
  aIrdData, aRealIrdCmdId, aRealIrdOperation, aRealIrdData: String;
  aValidationPeriod, aRandomPeriod, aTimeOut: String;
  aMasterIccNo: String;
  aMailId, aSegmentId, aTotalSegnment, aPriority, aMessage: String;
  aList: TStringList;
  aTransportId, aServiceId: String;

    { --------------------------------------------------- }

    function ConvertToRealIrdCmdId(const aMisId: Integer): String;
    begin
      Result := EmptyStr;
      case aMisId of
        1: Result := '200'; { �]�w�ˤl�K�X, 0xC8 }
        2: Result := '192'; { �ǰT��, 0xC0 }
        3: Result := '193'; { �j��x, 0xC1 }
        4: Result := '198'; { SetNetwork Id, 0xC6 }
        5: Result := '199'; { ���l���t�� Master/Slave Continuous Mode Initialisation, 0xC7 }
        6: Result := '199'; { ���l���Ѱt�� Master/Slave Cancellation, 0xC7 }
        7: Result := '199'; { ���l���{�����Ҥl�� Master/Slave Single Shot, 0xC7 }
        8: Result := '200'; { �]�w IPPV �q�ʱK�X, 0xC8 }
      end;
    end;

    { --------------------------------------------------- }

begin
  Result := False;
  aRealIrdCmdId := EmptyStr;
  aRealIrdOperation := EmptyStr;
  aRealIrdData := EmptyStr;
  aIrdCmdId := StrToInt( aMisIrdCmdID );
  aIrdData := aMisIrdCmdData;
  aRealIrdCmdId := ConvertToRealIrdCmdId( aIrdCmdId );
  case aIrdCmdId of
     1: { �ˤl�K�X }
       begin
         aRealIrdOperation := '001';
         aIrdData := Nvl( aMisIrdCmdData, '1234' );
         aIrdData := ( Lpad( IntToStr( Length( aIrdData ) ), 2, '0' ) + aIrdData );
         for aIndex := 1 to Length( aIrdData ) do
           aRealIrdData := aRealIrdData + IntToStr( Ord( aIrdData[aIndex] ) );
       end;
     2:{ Mail �ǰe�T�� }
       begin
         aRealIrdOperation := '001';
         { ���ন 2 �i�� }
         aMailId := IntToBin( frmMain.nG_SC_MailId );
         { �T�w�� 10 �X }
         aMailId := Copy( aMailId, Length( aMailId ) - 10 + 1, 10 );
         { �� 2 �i�� }
         aTotalSegnment := IntToBin( frmMain.nG_SC_TotalSeqment );
         { �� 6 �X }
         aTotalSegnment := Copy( aTotalSegnment, Length( aTotalSegnment ) - 6 + 1, 6  );
         { �� 2 �i�� }
         aIrdData := Nvl( aMisIrdCmdData, '0' );
         aPriority := IntToBin( StrToInt( aIrdData ) );
         { �� 2 �X }
         aPriority := Copy( aPriority, Length( aPriority ) - 2 + 1, 2 );
         { �� 2 �i�� }
         aSegmentId := IntToBin( frmMain.nG_SC_SegmentNumber );
         { �� 6 �X }
         aSegmentId := Copy( aSegmentId, Length( aSegmentId ) - 6 + 1, 6 );
         { �N�Ҧ��Ȭۥ[����, �� 16 �i��,
           Ex: 00000000 10000001 00000000 = 0x008100  }
         aIrdData := IntToHex( BinStrToInt(
           aMailId + aTotalSegnment + aPriority + aSegmentId ), 6 );
         { ���X�T��, �N�C�@�Ӧr���� ASCII ��, �� 16 �i�� }
         aMessage := EmptyStr;
         for aIndex := 1 to Length( frmMain.sG_SC_ActiveMessage ) do
         begin
           aMessage := ( aMessage + IntToHex( Ord(
             frmMain.sG_SC_ActiveMessage[aIndex] ), 0 ) );
         end;
         aIrdData := ( aIrdData + aMessage );
         { �N�Ҧ� 16 �i���, �ন ASCII }
         for aIndex := 1 to Length( aIrdData ) do
           aRealIrdData := ( aRealIrdData + IntToStr( Ord( aIrdData[aIndex] ) ) );
       end;
     3: { Force Tune �j��x }
       begin
         aRealIrdOperation := '001';
         if ( aIrdData <> EmptyStr ) then
         begin
           aTransportId := Copy( aIrdData, 1, Pos( '~', aIrdData ) - 1 );
           Delete( aIrdData, 1, Pos( '~', aIrdData ) );
           aServiceId := aIrdData;
           { network_id + transport_id + service_id }
           aRealIrdData := (
             IntToHex( StrToInt( frmMain.getCompNetworkInfo( aCompCode, aRealIrdCmdId ) ), 4 ) +
             IntToHex( StrToInt( aTransportId ), 4 ) +
             IntToHex( StrToInt( aServiceId ), 4 ) );
           { �C�@�� 16 �i��r�� ---> ASCII, ���� !!!! ���������, �৹ 16 �i��� }
           //for aIndex := 1 to Length( aIrdData ) do
           //  aRealIrdData := ( aRealIrdData + IntToStr( Ord( aIrdData[aIndex] ) ) );
         end;
       end;
     4: { Set Network Id }
       begin
         aRealIrdOperation := '001';
         aIrdData := frmMain.getCompNetworkInfo( aCompCode, aRealIrdCmdId );
         if ( aIrdData <> EmptyStr ) then
           aRealIrdData := aIrdData;
       end;
     5: { ���l���t�� }
       begin
         aRealIrdOperation := '001';
         { ����X Master Icc �d�� 16 �i�� }
         aIrdData := IntToHex( StrToInt( frmMain.sG_SC_MasterIccNo ), 0 );
         { ���X���n�ǤJ���Ѽ�, validationPeriod, randomPeriod, timeout }
         aList := TUstr.ParseStrings( aMisIrdCmdData, '~' );
         try
            aValidationPeriod := EmptyStr;
            aRandomPeriod := EmptyStr;
            aTimeOut := EmptyStr;
            if ( aList.Count = 3 ) then
            begin
              aValidationPeriod := IntToHex( StrToInt( aList.Strings[0] ), 2 );
              aRandomPeriod := IntToHex( StrToInt( aList.Strings[1] ), 2 );
              aTimeOut := IntToHex( StrToInt( aList.Strings[2] ), 2 );
            end;
         finally
           aList.Free;
         end;
         aIrdData := ( aIrdData + aValidationPeriod + aRandomPeriod + aTimeOut );
         { �� ASCII }
         for aIndex := 1 to Length( aIrdData ) do
           aRealIrdData := ( aRealIrdData + IntToStr( Ord( aIrdData[aIndex] ) ) );
       end;
     6: { ���l���Ѱt�� }
       begin
         aRealIrdOperation := '002';
         { No Data Value Need }
       end;
     7: { �{�����Ҥl�� }
       begin
         aRealIrdOperation := '003';
         { ����X Master Icc �d�� 16 �i�� }
         aIrdData := IntToHex( StrToInt( frmMain.sG_SC_MasterIccNo ), 0 );
         { Timeout �Ѽƪ� 16 �i�� }
         aTimeOut := IntToHex( StrToInt( Nvl( aMisIrdCmdData, '1' ) ), 2 );
         aIrdData := ( aIrdData + aTimeOut );
         { �� ASCII }
         for aIndex := 1 to Length( aIrdData ) do
           aRealIrdData := ( aRealIrdData + IntToStr( Ord( aIrdData[aIndex] ) ) );
       end;
     8: { IPPV �q�ʱK�X }
       begin
         aRealIrdOperation := '002';
         aIrdData := Nvl( aMisIrdCmdData, '1234' );
         aIrdData := ( Lpad( IntToStr( Length( aIrdData ) ), 2, '0' ) + aIrdData );
         for aIndex := 1 to Length( aIrdData ) do
           aRealIrdData := aRealIrdData + IntToStr( Ord( aIrdData[aIndex] ) );
       end;
  end;
  aIrdInfoObj.s_IrdCmdID := aRealIrdCmdId;
  aIrdInfoObj.s_IrdOperation := aRealIrdOperation;
  { �Ҧ����ױ��� IRD Data �����װ��H 2  }
  { ����, ���ΰ��H 2 }
  aIrdInfoObj.s_IrdDataLength := Lpad( IntToStr( Length( aRealIrdData ) ), 2, '0' );
  aIrdInfoObj.s_IrdData := aRealIrdData;
  Result := ( aRealIrdCmdId <> EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

procedure ParseCommandData(I_CmdInfoObj: TCmdInfoObject);
var
  nL_SegmentCount, nL_Mod, nL_Div : Integer;
  sL_Notes, sL_CommandNo, sL_FullMessage : String;
  sL_DetailCommandIDs, sL_ProductInfo : String;
  L_StrList, L_NotesStrList, L_ProductsStrList, L_ProductInfoStrList,
  L_CardStrList  : TStringList;
  aIndex, ii, kk : Integer;
  bL_ProfuctFieldFormatError : boolean;
  nL_ResultFlag  : Integer;
  sL_IccNo, sL_StbNo, sL_ImsProductID, sL_CurActiveLowLevelCmd : String;
  sL_SubBeginDate, sL_SubEndDate : String;
  AEventName, APrice: String;
  aOldCredit: String;
begin

     L_NotesStrList := nil;

    sL_CommandNo := I_CmdInfoObj.HIGH_LEVEL_CMD_ID ;

    sL_IccNo := genRealIccNo(I_CmdInfoObj.ICC_NO);
    sL_StbNo := genRealStbNo(I_CmdInfoObj.STB_NO);
   
    sL_ImsProductID := genRealImsproductID( I_CmdInfoObj.IMS_PRODUCT_ID );

    AEventName := '';
    APrice := '';


    try

      sL_DetailCommandIDs := getDetailCommandIDs( sL_CommandNo );
      if sL_DetailCommandIDs = '' then Exit;
      if not IsDataOK( sL_ImsProductID, sL_IccNo, sL_StbNo, nL_ResultFlag ) then
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
          frmMain.sG_SC_TransNumForUpdateCmdStatus :=
            Copy( frmMain.sG_SC_FullCommand, 1, 9 );
          updateCmdStatus;
        except
          { ... }
        end;
        Exit;
      end;

      frmMain.sG_CurTimeStamp := DateTimeToStr(now);
      try
        writeTimeStamp;
      except
        { ... }
      end;


      { �}���ɦh�e�@�ӳ]�w�K�X���O }
      //if sL_CommandNo = 'A1' then
      //begin
      //  sL_DetailCommandIDs := ( sL_DetailCommandIDs + ',69' );
      //end;


      L_StrList := TUStr.ParseStrings( sL_DetailCommandIDs, ',' );

      sL_SubBeginDate :=
        TUStr.replaceStr( I_CmdInfoObj.SUBSCRIPTION_BEGIN_DATE, '/', EmptyStr );

      sL_SubEndDate :=
        TUStr.replaceStr( I_CmdInfoObj.SUBSCRIPTION_END_DATE, '/', EmptyStr );

    for ii:=0 to L_StrList.Count -1 do
    begin
      frmMain.sG_CurActiveLowLevelCmd := L_StrList.Strings[ii];
      sL_CurActiveLowLevelCmd := L_StrList.Strings[ii];

      //����}�d��set ZipCode ���O��,�n�����P�_
      //if ( sL_CommandNo = 'A1' ) and
      //   ( frmMain.sG_CurActiveLowLevelCmd = '48' ) then
      //begin
      //  if ( not frmMain.bG_SetZipCode ) or ( I_CmdInfoObj.ZIP_CODE = '' ) then
      //    continue;
      //end;

      //if ( sL_CommandNo = 'A1' ) and
      //   ( frmMain.sG_CurActiveLowLevelCmd = '69' ) and
      //   ( ii = ( L_StrList.Count - 1 ) ) then
      //begin
        { Notes = �s���ˤl�K�X }
      //  I_CmdInfoObj.MIS_IRD_CMD_DATA := '1234';
      //  I_CmdInfoObj.MIS_IRD_CMD_ID := '1';
      //end;

      //����Redefine Credit Limit ���O��,�n�����P�_
      if ( sL_CurActiveLowLevelCmd = '100' ) then
      begin
        if StrToIntDef( I_CmdInfoObj.CREDIT_LIMIT,0 ) = 0 then
          frmMain.nG_SC_CreditLimit := frmMain.nG_DefaultCreditLimit;
      end;

      bL_ProfuctFieldFormatError := false;
      if (sL_CommandNo='B7') then //(�����v���ƻs)
      begin
        //update send_nagra set notes='111111111111^20020801^20020901,222222222222^20020901^20021231~3333333333;4444444444';
        sL_Notes := I_CmdInfoObj.NOTES;
        {
        �ǤJ�h�d��,�h�W�D(���Notes)=>�N�Ҧ��W�D grant ���Ҧ��d��.
        �U product ���H','�Ϲj. �U�d�����H';'�Ϲj.product data�P�d�� data�H'~'�Ϲj.(��product data,��d�� data)
        �U product ��info �]�t�T�ӳ���=>product id, �����_��,��������.�ӳo�T�ӳ����H'^'�@���Ϲj
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
          for aIndex :=0 to L_ProductsStrList.Count -1 do
          begin
            L_ProductInfoStrList := TUStr.ParseStrings(L_ProductsStrList.Strings[aIndex],'^');
            if (L_ProductInfoStrList<>nil) and (L_ProductInfoStrList.Count=3) then
            begin
              sL_ImsProductID := genRealImsproductID(L_ProductInfoStrList.Strings[0]);
              sL_SubBeginDate := TUstr.replaceStr(L_ProductInfoStrList.Strings[1],'/','');
              sL_SubEndDate := TUstr.replaceStr(L_ProductInfoStrList.Strings[2],'/','');
              for kk:=0 to L_CardStrList.Count -1 do
              begin
                sL_IccNo := genRealIccNo(L_CardStrList.Strings[kk]);

                preSendCommand( I_CmdInfoObj,  sL_IccNo, sL_ImsProductID,sL_SubBeginDate, sL_SubEndDate );

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

              except
                { ... }
              end;
              Break;
            end;
          end;

        end;
      end
      else if (sL_CommandNo='C1') then // (�Ѱ��t��)
      begin
         {
         �UICC�d�������H','�Ϲj
         }
        //update send_nagra set notes='1111111111,2222222222,3333333333';
        sL_Notes := I_CmdInfoObj.NOTES;
        L_NotesStrList := TUStr.ParseStrings(sL_Notes,',');
        for aIndex :=0 to L_NotesStrList.Count -1 do//��ܦ��o��h�� ICC �d�����
        begin
          sL_StbNo := genRealStbNo('');
          sL_IccNo := genRealIccNo(L_NotesStrList.Strings[aIndex]);
          preSendCommand(I_CmdInfoObj, sL_IccNo, sL_ImsProductID,
            sL_SubBeginDate, sL_SubEndDate );
        end;
      end
      else if (sL_CommandNo='P15') then // (���� STB �^�����`)
      begin
         {
         �UICC�d�������H','�Ϲj
         }
        //update send_nagra set notes='1111111111,2222222222,3333333333';
        sL_Notes := I_CmdInfoObj.NOTES;
        L_NotesStrList := TUStr.ParseStrings(sL_Notes,',');
        for aIndex :=0 to L_NotesStrList.Count -1 do//��ܦ��o��h�� ICC �d�����
        begin
          sL_IccNo := genRealIccNo(L_NotesStrList.Strings[aIndex]);
          preSendCommand(I_CmdInfoObj, sL_IccNo, sL_ImsProductID,
            sL_SubBeginDate, sL_SubEndDate );
        end;
      end
      //(B2:���W�D-�h��, B4:�Ȱ��W�D-�h��, B5:��_�W�D-�h��, P31:��PPV�`��-�h��)
      else if (sL_CommandNo='B2') or (sL_CommandNo='B4') or
              (sL_CommandNo='B5') or (sL_CommandNo='P31') then
      begin
         {
         V�U product ���H','�Ϲj
         }
        //update send_nagra set notes='111111111111,222222222222,333333333333,555555555555';
        sL_Notes := I_CmdInfoObj.NOTES;
        L_NotesStrList := TUStr.ParseStrings( sL_Notes,',' );
        for aIndex := 0 to L_NotesStrList.Count -1 do//��ܦ��o��h�� product ���
        begin
          sL_ImsProductID := genRealImsproductID(L_NotesStrList.Strings[aIndex]);
          preSendCommand( I_CmdInfoObj, sL_IccNo, sL_ImsProductID,
            sL_SubBeginDate, sL_SubEndDate );
        end;
      end
      { My Add }
      else if (sL_CommandNo='P30' ) then //(B10:�} PPV �`��-�h��)
      begin
        //update send_nagra set notes='A~111111111111~SpiderMan,A~222222222222~Holo,B~333333333333~WOW~550';
        sL_Notes := I_CmdInfoObj.NOTES;

        {
        �U product ���H,�Ϲj, product�Pevent name���H~�Ϲj, event name �P price ���H~�Ϲj;
        Ex:A~1~spiderman,A~2~holo,B~3~WOW~550 = >product 1,2,3 �H���O���@�Τ� price,
        �� product 3 �� price ���ۤv�� price , �� 550 ��
        }

        L_NotesStrList := TUStr.ParseStrings( sL_Notes,',' );
        //��ܦ��o��h�� product ���
        for  aIndex := 0 to L_NotesStrList.Count -1 do
        begin
          sL_ProductInfo := L_NotesStrList.Strings[aIndex];
          L_ProductsStrList := TUStr.ParseStrings( sL_ProductInfo, '~' );
          sL_ImsProductID := genRealImsproductID( L_ProductsStrList.Strings[0] );
          I_CmdInfoObj.EVENT_NAME := L_ProductsStrList.Strings[1];
          I_CmdInfoObj.PRICE := L_ProductsStrList.Strings[2];
          PreSendCommand( I_CmdInfoObj, sL_IccNo, sL_ImsProductID,
            sL_SubBeginDate, sL_SubEndDate );
        end;
      end
      else if (sL_CommandNo='E2')then //(E2:�ǰe�T��)
      begin
        { �ª����O }
        if I_CmdInfoObj.STB_FLAG = '0' then
        begin
          sL_FullMessage := I_CmdInfoObj.NOTES;
          Inc(frmMain.nG_SC_MailId);
          nL_Mod := length(sL_FullMessage) mod MAX_MESSAGE_LENGTH;
          nL_Div := length(sL_FullMessage) div MAX_MESSAGE_LENGTH;
          if nL_Mod=0 then //��ܰT�����׭�n�O  MAX_MESSAGE_LENGTH ����ƭ�
            nL_SegmentCount := nL_Div
          else
            nL_SegmentCount := nL_Div + 1;

          frmMain.nG_SC_TotalSeqment := nL_SegmentCount;

          for aIndex := 0 to nL_SegmentCount -1 do
          begin
            frmMain.nG_SC_SegmentNumber := aIndex + 1;
            frmMain.sG_SC_ActiveMessage := Copy(sL_FullMessage,aIndex * MAX_MESSAGE_LENGTH+1, MAX_MESSAGE_LENGTH);
            preSendCommand( I_CmdInfoObj, sL_IccNo, sL_ImsProductID,
              sL_SubBeginDate, sL_SubEndDate );
          end;
        end else
        begin
          { �s�����O }
          sL_FullMessage := I_CmdInfoObj.NOTES;
          Inc(frmMain.nG_SC_MailId);
          nL_Mod := Length( sL_FullMessage ) mod MAX_MESSAGE_LENGTH_NEW;
          nL_Div := Length( sL_FullMessage ) div MAX_MESSAGE_LENGTH_NEW;
          nL_SegmentCount := nL_Div;
          if nL_Mod <> 0 then
            nL_SegmentCount := nL_Div + 1;
          frmMain.nG_SC_TotalSeqment := nL_SegmentCount;
          for aIndex :=0 to nL_SegmentCount -1 do
          begin
            frmMain.nG_SC_SegmentNumber := aIndex; //aIndex + 1;
            frmMain.sG_SC_ActiveMessage := Copy( sL_FullMessage,
              aIndex * MAX_MESSAGE_LENGTH_NEW + 1, MAX_MESSAGE_LENGTH_NEW );
            preSendCommand( I_CmdInfoObj, sL_IccNo, sL_ImsProductID,
              sL_SubBeginDate, sL_SubEndDate );
          end;
        end;  
      end
      else if (sL_CommandNo='E5')then //(E5:Master/Slave continuous mode initialisation)
      begin
        //frmMain.sG_SC_MasterIccNo := sL_StbNo; //?? --> �u�_��, ��U���O�p���g�� ??
        frmMain.sG_SC_MasterIccNo := sL_IccNo;
         {
         V�U slave ICC  ���H','�Ϲj
         }
        //update send_nagra set notes='1160092498,1160092569';
        sL_Notes := I_CmdInfoObj.NOTES;
        L_NotesStrList := TUStr.ParseStrings(sL_Notes,',');
        for aIndex := 0 to L_NotesStrList.Count -1 do//��ܦ��o��h�� slave ICC ���
        begin
          { �����D, �� pdf ����Ӯ榡���@�� }
          sL_StbNo := genRealStbNo( L_NotesStrList.Strings[aIndex] );
          sL_IccNo := genRealIccNo( L_NotesStrList.Strings[aIndex] );
          preSendCommand( I_CmdInfoObj, sL_IccNo, sL_ImsProductID,
            sL_SubBeginDate, sL_SubEndDate );
        end;
      end
      else if ( sL_CommandNo = 'E7' ) then //( �{�����Ҥl�� )
      begin
        frmMain.sG_SC_MasterIccNo := sL_IccNo;
        sL_Notes := I_CmdInfoObj.NOTES;
        L_NotesStrList := TUStr.ParseStrings(sL_Notes,',');
        for aIndex := 0 to L_NotesStrList.Count -1 do//��ܦ��o��h�� slave ICC ���
        begin
          sL_StbNo := genRealStbNo( L_NotesStrList.Strings[aIndex] );
          sL_IccNo := genRealIccNo( L_NotesStrList.Strings[aIndex] );
          preSendCommand( I_CmdInfoObj, sL_IccNo, sL_ImsProductID,
            sL_SubBeginDate, sL_SubEndDate );
        end;
      end
      else if (sL_CommandNo='B6') then //(���i����)
      begin
        //update send_nagra set notes='A~111111111111,A~222222222222,A~333333333333,B~555555555555~20021031';
        sL_Notes := I_CmdInfoObj.NOTES;
        {
        V�U product ���H','�Ϲj, product�P���������H'~'�Ϲj;
        �Ĥ@�X�O'A'��,��ܦ@�Φ����O��subscription_end_datee,
        �Ĥ@�X�O'B'��,��ܦۤv���ۤv�� subscription_end_date.
        Ex:A~1,A~2,A~3,B~5~20020205=>product 1,2,3�H���O�� subscription_end_date�����;
        ��product 5 �� subscription_end_date ��20020205
        }
        L_NotesStrList := TUStr.ParseStrings( sL_Notes, ',' );
        for aIndex := 0 to L_NotesStrList.Count -1 do//��ܦ��o��h�� product ���
        begin
        
          sL_ProductInfo := L_NotesStrList.Strings[aIndex];
          L_ProductsStrList := TUStr.ParseStrings(sL_ProductInfo,'~');
          if (L_ProductsStrList.Strings[0]='A') and (L_ProductsStrList.Count=2) then //ex : A~1
          begin
            sL_ImsProductID := genRealImsproductID(L_ProductsStrList.Strings[1]);
            preSendCommand( I_CmdInfoObj, sL_IccNo, sL_ImsProductID,
              sL_SubBeginDate, sL_SubEndDate );
          end
          else if (L_ProductsStrList.Strings[0]='B') and (L_ProductsStrList.Count=3) then //ex: B~5~20021231
          begin
            sL_ImsProductID := genRealImsproductID(L_ProductsStrList.Strings[1]);

            sL_SubEndDate := TUstr.replaceStr(L_ProductsStrList.Strings[2],'/','');
            preSendCommand( I_CmdInfoObj, sL_IccNo, sL_ImsProductID,
              sL_SubBeginDate, sL_SubEndDate );
          end;
        end;
      end
      else if (sL_CommandNo='B1') then //(�}�W�D-�h��)
      begin
        //update send_nagra set notes='A~111111111111,A~222222222222,A~333333333333,B~555555555555~20020201~20020205';
        sL_Notes := I_CmdInfoObj.NOTES;

        {
        �U product ���H,�Ϲj, product�P���������H~�Ϲj;
        �Ĥ@�X�OA��,��ܦ@�Φ����O��subscription_begin_date �P subscription_end_datee,
        �Ĥ@�X�OB��,
        ��ܦۤv���ۤv�� subscription_begin_date �P subscription_end_date.
        Ex:A~1,A~2,A~3,B~5~20020201~20020205=>product 1,2,3�H���O�� subscription_begin_date, subscription_end_date�����;
        ��product 5 ��subscription_begin_date ��20020201, subscription_end_date ��20020205
        }

        L_NotesStrList := TUStr.ParseStrings(sL_Notes,',');
        for aIndex := 0 to L_NotesStrList.Count -1 do//��ܦ��o��h�� product ���
        begin
          sL_ProductInfo := L_NotesStrList.Strings[aIndex];
          L_ProductsStrList := TUStr.ParseStrings(sL_ProductInfo,'~');
          if (L_ProductsStrList.Strings[0]='A') and (L_ProductsStrList.Count=2) then //ex : A~1
          begin
            sL_ImsProductID := genRealImsproductID(L_ProductsStrList.Strings[1]);
            preSendCommand( I_CmdInfoObj, sL_IccNo, sL_ImsProductID,
              sL_SubBeginDate, sL_SubEndDate );
          end
          else if (L_ProductsStrList.Strings[0]='B') and (L_ProductsStrList.Count=4) then //ex: B~5~20020201~20020205
          begin
            sL_ImsProductID := genRealImsproductID(L_ProductsStrList.Strings[1]);
            sL_SubBeginDate := TUstr.replaceStr(L_ProductsStrList.Strings[2],'/','');
            sL_SubEndDate := TUstr.replaceStr(L_ProductsStrList.Strings[3],'/','');
            preSendCommand( I_CmdInfoObj, sL_IccNo, sL_ImsProductID,
              sL_SubBeginDate, sL_SubEndDate );
          end;
        end;
      end
      else if (sL_CommandNo='B9') then //(refresh)
      begin
        //update send_nagra set notes='A~111111111111,A~222222222222,A~333333333333,B~555555555555~20020201~20020205';
        {
        �U product ���H,�Ϲj, product�P���������H~�Ϲj;
        �Ĥ@�X�OA��,��ܦ@�Φ����O��subscription_begin_date �P subscription_end_datee,
        �Ĥ@�X�OB��,
        ��ܦۤv���ۤv�� subscription_begin_date �P subscription_end_date.
        Ex:A~1,A~2,A~3,B~5~20020201~20020205=>product 1,2,3�H���O�� subscription_begin_date, subscription_end_date�����;
        ��product 5 ��subscription_begin_date ��20020201, subscription_end_date ��20020205
        }

        if sL_CurActiveLowLevelCmd='2' then
        begin
          sL_Notes := I_CmdInfoObj.NOTES;
          L_NotesStrList := TUStr.ParseStrings(sL_Notes,',');
          for aIndex := 0 to L_NotesStrList.Count -1 do//��ܦ��o��h�� product ���
          begin
            sL_ProductInfo := L_NotesStrList.Strings[aIndex];
            L_ProductsStrList := TUStr.ParseStrings(sL_ProductInfo,'~');
            if (L_ProductsStrList.Strings[0]='A') and (L_ProductsStrList.Count=2) then //ex : A~1
            begin
              sL_ImsProductID := genRealImsproductID(L_ProductsStrList.Strings[1]);
              preSendCommand(I_CmdInfoObj, sL_IccNo, sL_ImsProductID,
                sL_SubBeginDate, sL_SubEndDate );
            end
            else if (L_ProductsStrList.Strings[0]='B') and (L_ProductsStrList.Count=4) then //ex: B~5~20020201~20020205
            begin
              sL_ImsProductID := genRealImsproductID(L_ProductsStrList.Strings[1]);
              sL_SubBeginDate := TUstr.replaceStr(L_ProductsStrList.Strings[2],'/','');
              sL_SubEndDate := TUstr.replaceStr(L_ProductsStrList.Strings[3],'/','');
              preSendCommand(I_CmdInfoObj, sL_IccNo, sL_ImsProductID,
                sL_SubBeginDate, sL_SubEndDate );
            end;
          end;
        end
        else
        begin
          PreSendCommand( I_CmdInfoObj, sL_IccNo, sL_ImsProductID,
            sL_SubBeginDate, sL_SubEndDate );
        end;
      end
      {
      else if ( sL_CommandNo='P1' ) then //(Enable IPPV with return path) ��h���ե�
      begin
        if ( frmMain.sG_CurActiveLowLevelCmd = '13' ) then
        begin
          aOldCredit := I_CmdInfoObj.CREDIT;
          I_CmdInfoObj.CREDIT := I_CmdInfoObj.NOTES;
        end
        else if ( frmMain.sG_CurActiveLowLevelCmd = '8' ) then
        begin
          I_CmdInfoObj.CREDIT := aOldCredit;
        end;
        PreSendCommand( I_CmdInfoObj, sL_IccNo, sL_ImsProductID,
          sL_SubBeginDate, sL_SubEndDate );
      end
      }
      else if ( sL_CommandNo='P100' ) then //(Enable IPPV without return path) ��h���ե�
      begin
        if ( frmMain.sG_CurActiveLowLevelCmd = '13' ) then
        begin
          aOldCredit := I_CmdInfoObj.CREDIT;
          I_CmdInfoObj.CREDIT := I_CmdInfoObj.NOTES;
        end
        else begin
          I_CmdInfoObj.CREDIT := aOldCredit;
        end;  
        PreSendCommand( I_CmdInfoObj, sL_IccNo, sL_ImsProductID,
          sL_SubBeginDate, sL_SubEndDate );
      end
      else
      begin
        PreSendCommand( I_CmdInfoObj, sL_IccNo, sL_ImsProductID,
          sL_SubBeginDate, sL_SubEndDate );
      end;
    end;
    except
      on ESocketError do Showmessage( 'Network error....' );
    else
    end;

    if Assigned( L_NotesStrList ) then L_NotesStrList.Free;

end;


function getTransNum(nI_CompCode:Integer; sI_SeqNo: String): String;
var
    bL_InsertMappingData : boolean;
    ii : Integer;
    sL_TransactionNum, sL_SeqNo, sL_SourceSeqNo : String;
begin
    sL_SourceSeqNo := sI_SeqNo;
    if nI_CompCode=StrToInt(TESTING_CMD_COMP_CODE) then //�p�G�O���դ��q�O
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

    //down, �N SeqNo �P transaction number �� mapping ���Y�����_��
    if bL_InsertMappingData then
      dtmMain.insertMappingData(sL_SourceSeqNo,sL_TransactionNum);
    //up, �N SeqNo �P transaction number �� mapping ���Y�����_��

    result := sL_TransactionNum;

end;

procedure writeTimeStamp;
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

{ ---------------------------------------------------------------------------- }

procedure updateCmdStatus;
begin
  try
    dtmMain.UpdateCMD_Status_For_SendCommand(
      frmMain.nG_SC_ConnectionID,
      frmMain.sG_SC_TransNumForUpdateCmdStatus,
      frmMain.sG_SC_CmdStatus,
      frmMain.sG_SC_ErrorCode,
      frmMain.sG_SC_ErrorMsg );
  except
    frmMain.showLogOnMemo( UPDATE_CMD_STATUS_ERROR_FLAG,
      ' SendCommand.updateCmdStatus' );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure WriteSendCmdLog(sI_CommandNo : String);
begin
  if not frmMain.bG_CommandLog then Exit;
  // �p�G�ǰe���O���ի��O,�h���L�� procedure
  if StrToIntDef( Copy( frmMain.sG_SC_FullCommand, 1, 3 ), 0 ) =
     ( StrToInt( TESTING_CMD_COMP_CODE ) ) then Exit;
  try
    dtmMain.InsertSendCmdLog( sI_CommandNo, frmMain.sG_SC_FullCommand,
      frmMain.sG_SC_Operator );
  except
    { ... }
  end;
end;

{ ---------------------------------------------------------------------------- }
procedure BuildSendCmdConnection(const aSocket: TTcpClient);
const
  AObjName: String = 'SMSGW';
var
  AIndex, AWriteSize, aReadCount: Integer;
  ADeviceCall : TDeviceCall;
  ADeviceAnswer : TDeviceAnswer;
  AString: String;
  aReadOK: Boolean;
begin
  frmMain.bG_SC_buildSendCmdConnection := False;
  try
    aSocket.RemoteHost := frmMain.getSendCommandServerIP;
    aSocket.RemotePort := IntToStr( frmMain.getSendCommandPort );
    if not aSocket.Active then
      aSocket.Active := True;

    ZeroMemory( @ADeviceCall, SizeOf( ADeviceCall ) );
    ZeroMemory( @ADeviceAnswer, SizeOf( ADeviceAnswer ) );

    ADeviceCall.op_mode := 0;

    for AIndex := 1 to Length( AObjName ) do
      ADeviceCall.obj_name[AIndex] := AObjName[AIndex];

    ADeviceCall.obj_name_len := Length( AObjName );

    ADeviceCall.len[1] := 0;

    ADeviceCall.len[2] := Length( AObjName ) +
      SizeOf( ADeviceCall.op_mode ) + SizeOf( ADeviceCall.obj_name_len );

    AWriteSize := Length( AObjName ) +
      SizeOf( ADeviceCall.op_mode ) + SizeOf( ADeviceCall.obj_name_len ) +
      SizeOf( ADeviceCall.len );

    aSocket.SendBuf( ADeviceCall, AWriteSize );
    Delay( 3000 );
    aSocket.ReceiveBuf( ADeviceAnswer, SizeOf( ADeviceAnswer ) );

    if ( ADeviceAnswer.status <> 6 ) or
       ( ADeviceAnswer.code <> 0 ) then
    begin
      frmMain.memErrorLog.Lines.Add(
         '['+ DateTimeToStr(now)+']' + ' NAGRA �ڵ��s�u�ШD.' );
    end
    else begin
      frmMain.memErrorLog.Lines.Add(
         '['+ DateTimeToStr(now)+']' + ' NAGRA �����s�u�ШD.' );
      frmMain.bG_HasBuildSendCmdConnection := True;
      frmMain.bG_SC_buildSendCmdConnection := True;
    end;


  except
    frmMain.sG_SC_ErrorCode := BUILD_CONNECTION_ERROR;
    frmMain.sG_SC_ErrorMsg := BUILD_CONNECTION_ERROR_MSG;
    frmMain.sG_SC_CmdStatus := 'E';
    frmMain.nG_SC_ConnectionID := frmMain.nG_CurActiveConn;
    updateCmdStatus;
    raise;
  end;

end;

{ ---------------------------------------------------------------------------- }

procedure OnTimerStartAction;
var
  ATimeStr: STring;
  wL_Hour, wL_Min, wL_Sec, sL_MSec : word;
  aCmdInfoObj: TCmdInfoObject;
begin
  DecodeTime( Now, wL_Hour, wL_Min, wL_Sec, sL_MSec );
  ATimeStr := Format( '%.2d%.2d', [wL_Hour, wL_Min] );
  if ATimeStr = DELETE_MAPPING_DATA_TIME then dtmMain.DeleteMappingData;
  if not frmMain.bG_RunSendingCommandThread then Exit;
  if frmMain.G_CommandInfoStrList.Count > 0 then
  begin
    try
      aCmdInfoObj := TCmdInfoObject( frmMain.G_CommandInfoStrList.Objects[0] );
      ParseCommandData( aCmdInfoObj );
      TCmdInfoObject( frmMain.G_CommandInfoStrList.Objects[0] ).Free;
      frmMain.G_CommandInfoStrList.Delete( 0 );
    except
      { ... }
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function genRealIccNo(const sI_SourceIccNo: String): String;
begin
  { 10 �X, �k�a, ����0 }
  Result := Copy( sI_SourceIccNo, 1, 10 );
  Result := TUstr.AddString( Result, '0', False, ICC_CARDNO_LENGTH );
end;

{ ---------------------------------------------------------------------------- }

function GenRealStbNo(const sI_SourceStbNo: String): String;
begin
  Result := Copy( Trim( sI_SourceStbNo ), 1, 10 );
  { 14 �X,�k�a, ����0,�k��4�Ӫť� }
  Result := TUstr.AddString( Result, '0', False, BOXNO_ID_LENGTH - 4 );
  if frmMain.bG_Simulator then
    { 14 �X,�k�a, ����0,�k��4��0 => for ������ =>
      �Ҧp�ϥΪ̿�J 65536, �h�ন   00000655360000 }
    Result := TUstr.AddString( Result, '0', True, BOXNO_ID_LENGTH )
  else
    { 14 �X,�k�a, ����0,�k��4�Ӫť� => for production }
    Result := TUstr.AddString( Result, ' ', True, BOXNO_ID_LENGTH );
end;

{ ---------------------------------------------------------------------------- }

function genRealImsproductID(sI_SourceImsProductID: String): String;
begin
  { 12 �X, �k�a,����0 }
  Result := TUstr.AddString( sI_SourceImsProductID, '0' , False,
    IMS_PRODUCT_ID_LENGTH);
end;

{ ---------------------------------------------------------------------------- }

end.
