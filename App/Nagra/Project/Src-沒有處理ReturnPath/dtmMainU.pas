unit dtmMainU;

interface

uses
  SysUtils, Classes, DB, ADODB, inifiles, Forms, Dialogs, DBClient;

type
  TdtmMain = class(TDataModule)
    cdsNetworkID: TClientDataSet;
    cdsNetworkIDCompCode: TIntegerField;
    cdsNetworkIDCompName: TStringField;
    cdsNetworkIDNetworkID: TStringField;
    cdsNetworkIDOperation: TStringField;
    cdsOperation: TClientDataSet;
    cdsOperationsCommandType: TStringField;
    cdsFeedback: TClientDataSet;
    cdsFeedbacksCommandType: TStringField;
    cdsProduct: TClientDataSet;
    cdsProductsCommandType: TStringField;
    cdsEmm: TClientDataSet;
    cdsEmmsCommandType: TStringField;
    cdsEmmsBroadcastMode: TStringField;
    cdsEmmsAddressType: TStringField;
    cdsControl: TClientDataSet;
    cdsControlsCommandType: TStringField;
    cdsControlsBroadcastMode: TStringField;
    cdsControlsAddressType: TStringField;
    cdsUser: TClientDataSet;
    cdsUsersID: TStringField;
    cdsUsersName: TStringField;
    cdsUsersPassword: TStringField;
    cdsParam: TClientDataSet;
    cdsParamsServerIP: TStringField;
    cdsParamnSPortNo: TIntegerField;
    cdsParamnRPortNo: TIntegerField;
    cdsParamsSysName: TStringField;
    cdsParamsSysVersion: TStringField;
    cdsParamsLogPath: TStringField;
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
    cdsParamnCmdRefreshRate1: TIntegerField;
    cdsParamnCmdRefreshRate2: TIntegerField;
    cdsParamnCmdResentTimes: TIntegerField;
    cdsParambShowUI: TBooleanField;
    cdsParambAssignBroadcastDate: TBooleanField;
    cdsParamsSHotTime: TStringField;
    cdsParamsEHotTime: TStringField;
  private
    { Private declarations }


  public
    { Public declarations }
    G_AdoConnAryForCallback : array of TADOConnection;
    G_AdoCallbackSp : array of TADOStoredProc;
    G_AdoSendCmdLogSp : array of TADOStoredProc;
    G_AdoCmdResponseLogSp : array of TADOStoredProc;
    G_AdoInsertMappingData : array of TADOStoredProc; //insert to so

    G_AdodeleteMappingData : array of TADOQuery;

    G_AdoConnAry : array of TADOConnection;
    G_AdoSendNagra : array of TADOQuery;
    G_AdoTmp : array of TADOQuery;


    G_AdoUpdateCmdStatusSp1 : array of TADOStoredProc;
    G_AdoUpdateCmdStatusSp2 : array of TADOStoredProc;
    G_AdoUpdateCmdStatusSp3 : array of TADOStoredProc;
    G_AdoUpdateCmdStatusBySeq : array of TADOStoredProc;


//    procedure insertTestConnestionCommand;
//    function getNotes(nI_ConnectionID:Integer; sI_SeqNo: String):String;
    procedure reConnectCallbackSvr;
    procedure deleteAllMappingData;
    procedure deleteMappingData;    
    function insertSendCmdLog(sI_HighLevelCmdID, sI_FullCmdText, sI_Operator:String):boolean;
    function insertCmdResponseLog(sI_FullCmdResponseText:String):boolean;
    function insertMappingData(sI_SeqNo, sI_TransNum:String):boolean;


    function insertCallbackData(sI_CallbackData:String):boolean;
    procedure UpdateCMD_Status_By_Seq(nI_ConnectionID:Integer; sI_SeqNo, sI_Status, sI_ErrorCode, sI_ErrorMsg: String);    
    procedure UpdateCMD_Status_For_SendCommand(nI_ConnectionID:Integer; sI_TransNum, sI_Status, sI_ErrorCode, sI_ErrorMsg: String);
    procedure UpdateCMD_Status_For_HandleResponse(nI_ConnectionID:Integer; sI_TransNum, sI_Status, sI_ErrorCode, sI_ErrorMsg: String);
    procedure activeSendNagra(nI_ActiveConn:Integer);    
    procedure updateCmdStatusToProcessing(sI_SeqNos : String; nI_ActiveConn:Integer);
    function connectToDB(var nI_DbCount:Integer; var I_CompCodeStrList:TStringList; var I_DbAliasStrList:TStringList):String;overload;
    function connectToDB(nI_ExceptionConnectionID:Integer; bI_IsExceptionConnID:boolean):boolean;overload;
  end;

var
  dtmMain: TdtmMain;

implementation

uses ConstObjU, frmMainU, Ustru;

{$R *.dfm}


procedure TdtmMain.activeSendNagra(nI_ActiveConn: Integer);
var
    sL_SQL, sL_Where, sL_CompCode : String;
    procedure executeSQL;
    begin
      sL_CompCode := frmMain.G_CompCodeStrList.Strings[nI_ActiveConn];
      sL_SQL := 'select HIGH_LEVEL_CMD_ID,ICC_NO,STB_NO,TO_CHAR(SUBSCRIPTION_BEGIN_DATE,' + STR_SEP  + 'YYYY/MM/DD' + STR_SEP + ') SUBSCRIPTION_BEGIN_DATE,TO_CHAR(SUBSCRIPTION_END_DATE,' + STR_SEP + 'YYYY/MM/DD' + STR_SEP + ') SUBSCRIPTION_END_DATE,CMD_STATUS,OPERATOR,IMS_PRODUCT_ID,';
      sL_SQL := sL_SQL +'ZIP_CODE,CREDIT_MODE,CREDIT,CREDIT_LIMIT,THRESHOLD_CREDIT,EVENT_NAME,PRICE,';
      sL_SQL := sL_SQL +'CC_NUMBER_1,IP_ADDR,CC_PORT,CALLBACK_DATE,CALLBACK_TIME,CALLBACK_FREQUENCY,FIRST_CALLBACK_DATE,';
      sL_SQL := sL_SQL +'PHONE_NUM_1,PHONE_NUM_2,PHONE_NUM_3,SEQNO,MIS_IRD_CMD_ID,MIS_IRD_CMD_DATA,';
      sL_SQL := sL_SQL +'COMPCODE, NOTES from send_nagra ';

      sL_Where := 'where (';
      sL_Where := sL_Where +'(CMD_STATUS=' + STR_SEP + 'W' + STR_SEP + ') or  (CMD_STATUS =' + STR_SEP + 'P' + STR_SEP + ' and  Operator=' + STR_SEP + BATCH_OPERATOR + STR_SEP + ')';

      if frmMain.nG_CmdResentTimes<>0 then
        sL_Where := sL_Where + 'or (CMD_STATUS =' + STR_SEP + 'P' + STR_SEP + ' and nvl(ReSentTimes,0)<=' + IntToStr(frmMain.nG_CmdResentTimes) + ' and Operator<>' + STR_SEP + BATCH_OPERATOR + STR_SEP + ')';
      sL_Where := sL_Where + ') ' ;

      sL_Where := sL_Where + ' and COMPCODE=' + sL_CompCode;
      sL_Where := sL_Where + ' and (PROCESSINGDATE< ' + TUstr.getOracleSQLDateTimeStr(now) + ' or PROCESSINGDATE IS NULL)';
      sL_SQL := sL_SQL + sL_Where + ' order by CMD_STATUS, nvl(resenttimes,0) DESC, SEQNO DESC ';

//select HIGH_LEVEL_CMD_ID,ICC_NO,STB_NO,TO_CHAR(SUBSCRIPTION_BEGIN_DATE,'YYYY/MM/DD') SUBSCRIPTION_BEGIN_DATE,TO_CHAR(SUBSCRIPTION_END_DATE,'YYYY/MM/DD') SUBSCRIPTION_END_DATE,CMD_STATUS,OPERATOR,IMS_PRODUCT_ID,ZIP_CODE,CREDIT_MODE,CREDIT,CREDIT_LIMIT,THRESHOLD_CREDIT,EVENT_NAME,PRICE,CC_NUMBER_1,IP_ADDR,CC_PORT,CALLBACK_DATE,CALLBACK_TIME,CALLBACK_FREQUENCY,FIRST_CALLBACK_DATE,PHONE_NUM_1,PHONE_NUM_2,PHONE_NUM_3,SEQNO,MIS_IRD_CMD_ID,MIS_IRD_CMD_DATA,COMPCODE, NOTES from send_nagra where ((CMD_STATUS='W') or  (CMD_STATUS ='P' and  Operator='批次作業')or (CMD_STATUS ='P' and nvl(ReSentTimes,0)<=2 and Operator<>'批次作業'))  and COMPCODE=9 and (PROCESSINGDATE< to_date('2003/10/01 09:47:18','YYYY/MM/DD HH24:MI:SS') or PROCESSINGDATE IS NULL) order by CMD_STATUS, nvl(resenttimes,0) DESC, SEQNO DESC       
      with G_AdoSendNagra[nI_ActiveConn] do
      begin

          Close;
          SQL.Clear;
          SQL.Add(sL_SQL);
          {
          if not G_AdoSendNagra[nI_ActiveConn].Connection.Connected then
            G_AdoSendNagra[nI_ActiveConn].Connection.Connected := true;

    //      G_AdoSendNagra[nI_ActiveConn].SQL.SaveToFile('c:\abcd.txt');
          }
          Open;
          First;
      end;
    end;
begin
  executeSQL;

end;

function TdtmMain.connectToDB(var nI_DbCount:Integer; var I_CompCodeStrList:TStringList; var I_DbAliasStrList:TStringList): String;
var
    L_IniFile : TIniFile;
    sL_ExeFileName, sL_ExePath, sL_IniFileName : STring;
    nL_DbCount, ii : Integer;
    sL_Result, sL_DbConnStr, sL_CompCode : String;
    sL_DbAlias, sL_DbUserName, sL_DbPassword  : String;
begin
    sL_ExeFileName := Application.ExeName;
    sL_ExePath := ExtractFileDir(sL_ExeFileName);
    if frmMain.sG_NewIniFileName='' then
      sL_IniFileName := sL_ExePath + '\' + CABLE_NAGRA_INI_FILE_NAME
    else
      sL_IniFileName := sL_ExePath + '\' + frmMain.sG_NewIniFileName;


    sL_Result := '';
    if not FileExists(sL_IniFileName) then
    begin
      sL_DbConnStr := '';
      result := frmMain.getLanguageSettingInfo('dtmMain_Msg_1_Content')+sL_IniFileName + frmMain.getLanguageSettingInfo('dtmMain_Msg_2_Content');
      exit;
    end;

    try
      L_IniFile := TIniFile.Create(sL_IniFileName);
      nL_DbCount := L_IniFile.ReadInteger('DBINFO','DB_COUNT',0);
      nI_DbCount := nL_DbCount;

      setLength(G_AdoConnAry,nL_DbCount);
      setLength(G_AdoSendNagra,nL_DbCount);
      setLength(G_AdoTmp,nL_DbCount);
      setLength(G_AdoUpdateCmdStatusSp1,nL_DbCount);
      setLength(G_AdoUpdateCmdStatusSp2,nL_DbCount);
      setLength(G_AdoUpdateCmdStatusSp3,nL_DbCount);
      setLength(G_AdoUpdateCmdStatusBySeq,nL_DbCount);
            
      setLength(G_AdoInsertMappingData,nL_DbCount);

      setLength(G_AdoConnAryForCallback,1);
      setLength(G_AdoCallbackSp,1);
      setLength(G_AdoSendCmdLogSp,1);
      setLength(G_AdoCmdResponseLogSp,1);



      setLength(G_AdodeleteMappingData,1);




      //down,  connect 到 send_nagra...

      for ii:=1 to nL_DbCount do
      begin
        sL_DbAlias := L_IniFile.ReadString('DBINFO','ALIAS_' + IntToStr(ii),'');
        sL_DbUserName := L_IniFile.ReadString('DBINFO','USERID_' +  IntToStr(ii),'');
        sL_DbPassword := L_IniFile.ReadString('DBINFO','PASSWORD_' + IntToStr(ii),'');
        sL_CompCode := L_IniFile.ReadString('DBINFO','COMPCODE_' + IntToStr(ii),'');

        I_CompCodeStrList.Add(sL_CompCode);
        I_DbAliasStrList.Add(sL_DbAlias);

  //    sI_DbConnStr := 'Provider=OraOLEDB.Oracle.1;Password=' + sI_DbPassword +';Persist Security Info=True;User ID='+sI_DbUserName+';Data Source='+sI_DbAlias; //=> 適用於my nb
        sL_DbConnStr := 'Provider=MSDAORA.1;Password=' + sL_DbPassword + ';User ID='+sL_DbUserName+';Data Source='+sL_DbAlias+';Persist Security Info=True'; //適用於其他 PC
        G_AdoConnAry[ii-1] := TAdoConnection.Create(nil);
        G_AdoConnAry[ii-1].LoginPrompt := false;
        G_AdoConnAry[ii-1].ConnectionString := sL_DbConnStr;
        G_AdoConnAry[ii-1].Connected := true;

        G_AdoSendNagra[ii-1] := TAdoQuery.Create(nil);
        G_AdoSendNagra[ii-1].Connection := G_AdoConnAry[ii-1];

        G_AdoTmp[ii-1] := TAdoQuery.Create(nil);
        G_AdoTmp[ii-1].Connection := G_AdoConnAry[ii-1];

        G_AdoUpdateCmdStatusSp1[ii-1] := TADOStoredProc.Create(nil);
        G_AdoUpdateCmdStatusSp1[ii-1].Connection := G_AdoConnAry[ii-1];
        G_AdoUpdateCmdStatusSp1[ii-1].ProcedureName := 'SF_UPDATESENDNAGRACMDSTATUS1';
        G_AdoUpdateCmdStatusSp1[ii-1].Parameters.Refresh;

        G_AdoUpdateCmdStatusSp2[ii-1] := TADOStoredProc.Create(nil);
        G_AdoUpdateCmdStatusSp2[ii-1].Connection := G_AdoConnAry[ii-1];
        G_AdoUpdateCmdStatusSp2[ii-1].ProcedureName := 'SF_UPDATESENDNAGRACMDSTATUS2';
        G_AdoUpdateCmdStatusSp2[ii-1].Parameters.Refresh;

        G_AdoUpdateCmdStatusSp3[ii-1] := TADOStoredProc.Create(nil);
        G_AdoUpdateCmdStatusSp3[ii-1].Connection := G_AdoConnAry[ii-1];
        G_AdoUpdateCmdStatusSp3[ii-1].ProcedureName := 'SF_UPDATESENDNAGRACMDSTATUS3';
        G_AdoUpdateCmdStatusSp3[ii-1].Parameters.Refresh;

        G_AdoUpdateCmdStatusBySeq[ii-1] := TADOStoredProc.Create(nil);
        G_AdoUpdateCmdStatusBySeq[ii-1].Connection := G_AdoConnAry[ii-1];
        G_AdoUpdateCmdStatusBySeq[ii-1].ProcedureName := 'SF_UPDATESENDNAGRACMDSTATUS4';
        G_AdoUpdateCmdStatusBySeq[ii-1].Parameters.Refresh;

        G_AdoInsertMappingData[ii-1] := TADOStoredProc.Create(nil);
        G_AdoInsertMappingData[ii-1].Connection := G_AdoConnAry[ii-1];
        G_AdoInsertMappingData[ii-1].ProcedureName := 'SF_INSERTSEQNOTRANSMAPPING';
        G_AdoInsertMappingData[ii-1].Parameters.Refresh;


      end;
      //up,  connect 到 send_nagra...

      //down,  connect 到寫入 callback data 的 db
      sL_DbAlias := L_IniFile.ReadString('CALLBACKSVR','CALLBACK_DATA_ALIAS','');
      sL_DbUserName := L_IniFile.ReadString('CALLBACKSVR','CALLBACK_DATA_USERID','');
      sL_DbPassword := L_IniFile.ReadString('CALLBACKSVR','CALLBACK_DATA_PASSWORD','');

      sL_DbConnStr := 'Provider=MSDAORA.1;Password=' + sL_DbPassword + ';User ID='+sL_DbUserName+';Data Source='+sL_DbAlias+';Persist Security Info=True'; //適用於其他 PC
      G_AdoConnAryForCallback[0] := TAdoConnection.Create(nil);
      G_AdoConnAryForCallback[0].LoginPrompt := false;
      G_AdoConnAryForCallback[0].ConnectionString := sL_DbConnStr;
      G_AdoConnAryForCallback[0].Connected := true;

      G_AdoCallbackSp[0] := TADOStoredProc.Create(nil);
      G_AdoCallbackSp[0].Connection := G_AdoConnAryForCallback[0];
      G_AdoCallbackSp[0].ProcedureName := 'SF_INSERTCACALLBACKDATA';
      G_AdoCallbackSp[0].Parameters.Refresh;

      G_AdoSendCmdLogSp[0] := TADOStoredProc.Create(nil);
      G_AdoSendCmdLogSp[0].Connection := G_AdoConnAryForCallback[0];
      G_AdoSendCmdLogSp[0].ProcedureName := 'SF_INSERTSENDCMDLOG';
      G_AdoSendCmdLogSp[0].Parameters.Refresh;

      G_AdoCmdResponseLogSp[0] := TADOStoredProc.Create(nil);
      G_AdoCmdResponseLogSp[0].Connection := G_AdoConnAryForCallback[0];
      G_AdoCmdResponseLogSp[0].ProcedureName := 'SF_INSERTCMDRESPONSELOG';
      G_AdoCmdResponseLogSp[0].Parameters.Refresh;



      G_AdodeleteMappingData[0] := TADOQuery.Create(nil);
      G_AdodeleteMappingData[0].Connection := G_AdoConnAryForCallback[0];
      deleteAllMappingData;


      //up,  connect 到寫入 callback data 的 db


      L_IniFile.Free;
    except
      result := frmMain.getLanguageSettingInfo('dtmMain_Msg_3_Content') + ' Alias:'+ sL_DbAlias + '  UserID:' +sL_DbUserName;
    end;
end;

{
function TdtmMain.getNotes(nI_ConnectionID: Integer;
  sI_SeqNo: String): String;
var
    sL_Notes : String;
begin
    sL_Notes := '';
    with dtmMain.G_AdoTmp[nI_ConnectionID] do
    begin
        Close;
        SQL.Clear;
        SQL.Add('select NOTES from SEND_NAGRA where SEQNO ='+ STR_SEP + sI_SeqNo + STR_SEP);
        Open;
        sL_Notes := fieldbyname('NOTES').asstring;
        Close;
    end;
    result := sL_Notes;
end;
}
{
procedure TdtmMain.insertTestConnestionCommand;
var
    sL_SQL, sL_CompCode,sL_Operator : String;
begin

    sL_CompCode := frmMain.G_CompCodeStrList.Strings[0];
    sL_SQL := 'delete send_nagra where SEQNO=0';
    dtmMain.G_AdoTmp[0].SQL.Clear;
    dtmMain.G_AdoTmp[0].SQL.Add(sL_SQL);
    dtmMain.G_AdoTmp[0].ExecSQL;


    sL_Operator := '';
    sL_SQL := 'insert into send_nagra(HIGH_LEVEL_CMD_ID, CMD_STATUS,SEQNO, COMPCODE,RESENTTIMES,OPERATOR)';
    sL_SQL := sL_SQL + ' values (' + STR_SEP + CA_TESTING_CMD_ID + STR_SEP + ',' ;
    sL_SQL := sL_SQL + STR_SEP + 'W' + STR_SEP + ',' + '0' + ',' + STR_SEP + sL_CompCode + STR_SEP +',' + IntToStr(frmMain.nG_CmdResentTimes)+ ',' + STR_SEP + sL_Operator + STR_SEP   +')' ;



    dtmMain.G_AdoTmp[0].SQL.Clear;
    dtmMain.G_AdoTmp[0].SQL.Add(sL_SQL);
    dtmMain.G_AdoTmp[0].ExecSQL;
    

end;
}

function TdtmMain.connectToDB(nI_ExceptionConnectionID: Integer; bI_IsExceptionConnID:boolean):boolean;
var
    ii : Integer;
    bL_Result : boolean;
begin
    bL_Result := true;
    try
      if bI_IsExceptionConnID then
      begin
        G_AdoConnAry[nI_ExceptionConnectionID].Connected := false;
        G_AdoConnAry[nI_ExceptionConnectionID].Connected := true;
      end
      else
      begin
        for ii:=0 to High(G_AdoConnAry) do
        begin
          if ii<>nI_ExceptionConnectionID then
          begin
            G_AdoConnAry[ii].Connected := false;
            G_AdoConnAry[ii].Connected := true;
          end;
        end;
      end;
    except
      bL_Result := false;
      raise;
    end;
    result := bL_Result;
end;

function TdtmMain.insertCallbackData(sI_CallbackData: String):boolean;
var
    fL_FunReturn, fL_RetCode : Double;
    sL_Msg : String;
    bL_Result : boolean;
begin
    try
      with G_AdoCallbackSp[0] do
      begin
        Parameters.ParamByName('P_CALLBACKDATA').Value := sI_CallbackData;

        ExecProc;

        fL_FunReturn := Parameters.Items[0].Value; //用ParamByName會有問題
        fL_RetCode := Parameters.Items[2].Value; //用ParamByName會有問題
        sL_Msg := Parameters.Items[3].Value; //用ParamByName會有問題

        if fL_RetCode=0 then
          bL_Result := true
        else
          bL_Result := false;

  //        Close;

      end;
    except
      bL_Result := false;

      frmMain.showLogOnMemo(EXECUTE_ORACLE_SP_ERROR_FLAG, frmMain.getLanguageSettingInfo('dtmMain_Msg_4_Content') + 'SF_INSERTCACALLBACKDATA' + frmMain.getLanguageSettingInfo('dtmMain_Msg_5_Content')); //執行完此行程式後, 程式會結束掉
//      reConnectCallbackSvr;
    end;

    result := bL_Result;
end;

function TdtmMain.insertSendCmdLog(sI_HighLevelCmdID,
  sI_FullCmdText, sI_Operator: String): boolean;
var
    fL_FunReturn, fL_RetCode : Double;
    sL_Msg : String;
    bL_Result : boolean;
begin
    try
      with G_AdoSendCmdLogSp[0] do
      begin
        Parameters.ParamByName('P_HIGHLEVELCMDID').Value := sI_HighLevelCmdID;
        Parameters.ParamByName('P_FULLCMDTEXT').Value := sI_FullCmdText;
        Parameters.ParamByName('P_OPERATOR').Value := sI_Operator;

        ExecProc;

        fL_FunReturn := Parameters.Items[0].Value; //用ParamByName會有問題
        fL_RetCode := Parameters.Items[4].Value; //用ParamByName會有問題
        sL_Msg := Parameters.Items[5].Value; //用ParamByName會有問題

        if fL_RetCode=0 then
          bL_Result := true
        else
          bL_Result := false;

  //        Close;

      end;
    except
      bL_Result := false;

      frmMain.showLogOnMemo(EXECUTE_ORACLE_SP_ERROR_FLAG,frmMain.getLanguageSettingInfo('dtmMain_Msg_4_Content')+ 'SF_INSERTSENDCMDLOG' + frmMain.getLanguageSettingInfo('dtmMain_Msg_5_Content')); //執行完此行程式後, 程式會結束掉
//      reConnectCallbackSvr;
    end;

    result := bL_Result;
end;

procedure TdtmMain.updateCmdStatusToProcessing(sI_SeqNos: String;
  nI_ActiveConn: Integer);
var
    sL_SQL : String;
    L_AdoQuery : TAdoQuery;
    fL_FunReturn, fL_RetCode : Double;
    sL_Msg : String;

    procedure executeSQL;
    begin
      with G_AdoUpdateCmdStatusSp1[nI_ActiveConn] do
      begin
        Parameters.ParamByName('P_CMDSTATUS').Value := 'P';
        Parameters.ParamByName('P_ERRORCODE').Value := sI_SeqNos;

        Parameters.ParamByName('P_ERRORMSG').Value := '';

        Parameters.ParamByName('P_TRANSNUM').Value := '0'; //一定要傳入 0 
        Parameters.ParamByName('P_NAGRASRCID').Value := frmMain.sG_SourceID;


        ExecProc;

        fL_FunReturn := Parameters.Items[0].Value; //用ParamByName會有問題
        fL_RetCode := Parameters.Items[6].Value; //用ParamByName會有問題
        sL_Msg := Parameters.Items[7].Value; //用ParamByName會有問題


//        Close;

      end;

    end;
begin
    executeSQL;
end;

procedure TdtmMain.UpdateCMD_Status_For_SendCommand(nI_ConnectionID: Integer; sI_TransNum,
  sI_Status, sI_ErrorCode, sI_ErrorMsg: String);
var
    sL_SQL : String;
    procedure executeSQL;
    var
      fL_FunReturn, fL_RetCode : Double;
      sL_Msg : String;
      L_StrList : TStringList;
    begin
      with G_AdoUpdateCmdStatusSp2[nI_ConnectionID] do
      begin
        Parameters.ParamByName('P_CMDSTATUS').Value := sI_Status;
        Parameters.ParamByName('P_ERRORCODE').Value := sI_ErrorCode;

        Parameters.ParamByName('P_ERRORMSG').Value := sI_ErrorMsg;

        Parameters.ParamByName('P_TRANSNUM').Value := sI_TransNum;
        Parameters.ParamByName('P_NAGRASRCID').Value := frmMain.sG_SourceID;


        ExecProc;

        fL_FunReturn := Parameters.Items[0].Value; //用ParamByName會有問題
        fL_RetCode := Parameters.Items[6].Value; //用ParamByName會有問題
        sL_Msg := Parameters.Items[7].Value; //用ParamByName會有問題


      end;

    end;
begin
    if nI_ConnectionID<0 then Exit;
    if StrToIntDef(Copy(sI_TransNum,1,3),0)<>(StrToInt(TESTING_CMD_COMP_CODE)) then //若不是連線測試指令
      executeSQL;
end;

procedure TdtmMain.UpdateCMD_Status_For_HandleResponse(nI_ConnectionID: Integer; sI_TransNum,
  sI_Status, sI_ErrorCode, sI_ErrorMsg: String);
var
    sL_SQL : String;
    procedure executeSQL;
    var
      fL_FunReturn, fL_RetCode : Double;
      sL_Msg : String;
      L_StrList : TStringList;
    begin
      with G_AdoUpdateCmdStatusSp3[nI_ConnectionID] do
      begin
        Parameters.ParamByName('P_CMDSTATUS').Value := sI_Status;
        Parameters.ParamByName('P_ERRORCODE').Value := sI_ErrorCode;

        Parameters.ParamByName('P_ERRORMSG').Value := sI_ErrorMsg;

        Parameters.ParamByName('P_TRANSNUM').Value := sI_TransNum;
        Parameters.ParamByName('P_NAGRASRCID').Value := frmMain.sG_SourceID;


        ExecProc;

        fL_FunReturn := Parameters.Items[0].Value; //用ParamByName會有問題
        fL_RetCode := Parameters.Items[6].Value; //用ParamByName會有問題
        sL_Msg := Parameters.Items[7].Value; //用ParamByName會有問題

//        Close;

      end;

    end;
begin
    if nI_ConnectionID<0 then Exit;
      executeSQL;
end;

function TdtmMain.insertMappingData(sI_SeqNo,
  sI_TransNum: String): boolean;
var
    nL_ConnectionID : Integer;
    fL_FunReturn, fL_RetCode : Double;
    sL_Msg, sL_CompCode : String;
    bL_Result : boolean;
begin
    //若將此 mapping data 放在 string list 中, 指令的處理速度會比較快, 但似乎比較不穩定
    try

      sL_CompCode := Copy(sI_TransNum,1,3);
      nL_ConnectionID := frmMain.getConnectionID(StrToInt(sL_CompCode));
      with G_AdoInsertMappingData[nL_ConnectionID] do
      begin
        Parameters.ParamByName('P_SOURCEID').Value := frmMain.sG_SourceID;
        Parameters.ParamByName('P_SEQNO').Value := sI_SeqNo;
        Parameters.ParamByName('P_TRANSNUMBER').Value := sI_TransNum;
        ExecProc;

        fL_FunReturn := Parameters.Items[0].Value; //用ParamByName會有問題
        fL_RetCode := Parameters.Items[4].Value; //用ParamByName會有問題
        sL_Msg := Parameters.Items[5].Value; //用ParamByName會有問題

        if fL_RetCode=0 then
          bL_Result := true
        else
          bL_Result := false;

  //        Close;

      end;
    except
      bL_Result := false;
      frmMain.showLogOnMemo(EXECUTE_ORACLE_SP_ERROR_FLAG,'後端程式 : SF_INSERTSEQNOTRANSMAPPING' + ' .程式即將結束,然後重新啟動. '); //執行完此行程式後, 程式會結束掉
//      reConnectCallbackSvr;
    end;

    result := bL_Result;
end;

procedure TdtmMain.deleteAllMappingData;
begin
    try
      with G_AdodeleteMappingData[0] do
      begin
        SQL.clear;
        SQL.Add('delete CaSeqNoTransNumMappingData where SOURCEID=' + '''' + frmMain.sG_SourceID + '''');
        ExecSQL;
        Close;
      end;
    except

    end;
end;


procedure TdtmMain.deleteMappingData;
var
    sL_SQL : String;
begin
    //down, 刪除兩小時前的 CaSeqNoTransNumMappingData ..
    with G_AdodeleteMappingData[0] do
    begin
      sL_SQL := 'delete CaSeqNoTransNumMappingData where SOURCEID=' + '''' + frmMain.sG_SourceID + '''';
      sL_SQL := sL_SQL + ' and UPDTIME < ' + TUstr.getOracleSQLDateTimeStr(now-1/12);
      SQL.clear;
      SQL.Add(sL_SQL);
      ExecSQL;
      Close;
    end;
    //up, 刪除兩小時前的 CaSeqNoTransNumMappingData ..

end;

function TdtmMain.insertCmdResponseLog(
  sI_FullCmdResponseText: String): boolean;
var
    fL_FunReturn, fL_RetCode : Double;
    sL_Msg : String;
    bL_Result : boolean;
begin
    try
      with G_AdoCmdResponseLogSp[0] do
      begin
        Parameters.ParamByName('P_FULLCMDRESPONSETEXT').Value := sI_FullCmdResponseText;

        ExecProc;

        fL_FunReturn := Parameters.Items[0].Value; //用ParamByName會有問題
        fL_RetCode := Parameters.Items[2].Value; //用ParamByName會有問題
        sL_Msg := Parameters.Items[3].Value; //用ParamByName會有問題

        if fL_RetCode=0 then
          bL_Result := true
        else
          bL_Result := false;

  //        Close;

      end;
    except
      bL_Result := false;

      frmMain.showLogOnMemo(EXECUTE_ORACLE_SP_ERROR_FLAG,frmMain.getLanguageSettingInfo('dtmMain_Msg_4_Content') +'SF_INSERTCMDRESPONSELOG' + frmMain.getLanguageSettingInfo('dtmMain_Msg_5_Content')); //執行完此行程式後, 程式會結束掉
//      reConnectCallbackSvr;

    end;

    result := bL_Result;
end;

procedure TdtmMain.reConnectCallbackSvr;
begin
    G_AdoConnAryForCallback[0].Connected := false;
    G_AdoConnAryForCallback[0].Connected := true;    
end;

procedure TdtmMain.UpdateCMD_Status_By_Seq(nI_ConnectionID: Integer;
  sI_SeqNo, sI_Status, sI_ErrorCode, sI_ErrorMsg: String);
var
    sL_SQL : String;
    procedure executeSQL;
    var
      fL_FunReturn, fL_RetCode : Double;
      sL_Msg : String;
      L_StrList : TStringList;
    begin
      with G_AdoUpdateCmdStatusBySeq[nI_ConnectionID] do
      begin
        Parameters.ParamByName('P_CMDSTATUS').Value := sI_Status;
        Parameters.ParamByName('P_ERRORCODE').Value := sI_ErrorCode;
        Parameters.ParamByName('P_ERRORMSG').Value := sI_ErrorMsg;

        Parameters.ParamByName('P_SEQNO').Value := sI_SeqNo;

        ExecProc;

        fL_FunReturn := Parameters.Items[0].Value; //用ParamByName會有問題
        fL_RetCode := Parameters.Items[5].Value; //用ParamByName會有問題
        sL_Msg := Parameters.Items[6].Value; //用ParamByName會有問題

//        Close;

      end;

    end;
begin
    if nI_ConnectionID<0 then Exit;
      executeSQL;
end;

end.
