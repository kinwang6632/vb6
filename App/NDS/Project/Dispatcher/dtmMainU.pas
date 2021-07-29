unit dtmMainU;

interface

uses
  SysUtils, Classes, DB, DBClient ,Forms ,Variants ,IniFiles,
  ADODB ,Dialogs;

const
  NO_AVAILABLE_QUERY_COMP = '無法取得適當的元件reference.以致無法讀取資料庫之資料.';
  SYS_INI_FILE_NAME = 'CableNDS.ini';
  DATA_ALIAS = 'ALIAS_';
  DATA_USERID= 'USERID_';
  DATA_PASSWORD='PASSWORD_';
  DATA_COMPCODE='COMPCODE_';
  DATA_COMPNAME='COMPNAME_';
  DATA_PROCESSOR_IP='PROCESSORIP_';
  STR_SEP = '''';
  SELECT_MODE = 1;
  IUD_MODE = 2;
  SPACE_STR = '1';
  ZERO_STR = '2';



type
  TProcessorIPData = class(TObject)
    sCompCodeString : String;
    sProcessorIp : String;
  end;

  TCableNdsIniData = class(TObject)
    nDataAreaNo : Integer;
    sAlias : String;
    sUserID : String;
    sPassword : String;
    sCompCode : String;
    sCompName : String;
    sProcessorIP : String;
  end;

  TdtmMain = class(TDataModule)
    cdsUser: TClientDataSet;
    cdsUsersID: TStringField;
    cdsUsersName: TStringField;
    cdsUsersPassword: TStringField;
    cdsParam: TClientDataSet;
    cdsParamsSysName2: TStringField;
    cdsParamsSysVersion2: TStringField;
    cdsParamsLogPath: TStringField;
    cdsParamnTimeOut2: TIntegerField;
    cdsParamnMaxCommandCount2: TIntegerField;
    cdsParambCommandLog2: TBooleanField;
    cdsParamdUptTime2: TDateTimeField;
    cdsParamsUptName2: TStringField;
    cdsParamnCmdRefreshRate1: TIntegerField;
    cdsParamnCmdRefreshRate22: TIntegerField;
    cdsParamnCmdResentTimes2: TIntegerField;
    cdsParambShowUI2: TBooleanField;
    cdsParamsSHotTime2: TStringField;
    cdsParamsEHotTime2: TStringField;
    cdsParamnVersion: TIntegerField;
    cdsParambResponseLog: TBooleanField;
    adoQrySendNdsData: TADOQuery;
    cdsDbLinkSet: TClientDataSet;
    cdsDbLinkSetGroup: TStringField;
    cdsDbLinkSetALIAS: TStringField;
    cdsDbLinkSetUSERID: TStringField;
    cdsDbLinkSetPASSWORD: TStringField;
    cdsDbLinkSetCOMPCODE: TStringField;
    cdsDbLinkSetCOMPNAME: TStringField;
    cdsDbLinkSetPROCESSORIP: TStringField;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
    G_DisconnectProcessorIP : TStringList;
    G_DbConnAry : array of TADOConnection;
    G_AdoQueryAry : array of TADOQuery;
    G_AdoStoredProcAry1 : array of TADOStoredProc; //SF_UPDATESENDNDSCMDSTATUS
    G_AdoStoredProcAry2 : array of TADOStoredProc; //SF_UPDATESENDNDSCMDSTATUS
//    AdoStoredProcAry2 : array of TADOStoredProc; //SF_UPDATESENDNDSCMDSTATUS
    //AdoStoredProcAry3 : array of TADOStoredProc; //SF_UPDATESENDNDSCMDSTATUS
    G_AdoSFSendCmdLogAry : array of TADOStoredProc; //SF_InsertCmdResponseLog
    G_AdoSFResponseLogAry : array of TADOStoredProc; //SF_InsertCmdResponseLog
    procedure transToPureDate(sI_StartTime,sI_EndTime : String;var dI_PureStartTime,dI_PureEndTime : TDateTime);
    procedure saveProcessorIP(sI_CompCode,sI_ProcessorIP : String);

  public
    { Public declarations }
    G_CompInfoStrList : TStringList;
    G_ProcessorIPStrList : TStringList;
    G_UserIDStrList : TStringList;
    sG_DispatcherIP,sG_ProcessorListenPort,sG_DispatcherListenPort : String;
    function getExePath : String;
    procedure activeCDS;
    procedure saveToFile(vI_Content:OleVariant; sI_FileName:String);
    function getActiveQuery(sI_AreaDataNo : String):TADOQuery;
    function checkUserIdIsOnly(sI_UserID : String) : Boolean;
    function getAllUserID : TStringList;
    function getCmdRefreshRate : Integer;
    function getDbConnInfo: String;
    function connectToDB: WideString;
    function getCompName(sI_CompCode : String) : String;
    function getProcessorIP(sI_CompCode : String) : String;
    function getSendNdsData(sI_CompCode : String;var sI_Msg: String) : TADOQuery;
    function fillStringLength(sI_String,sI_FillStr : String; nI_Length : Integer): String;
    function isDisconnect(sI_ProcessorIP:String):boolean;
    procedure hasConnected(sI_ProcessorIP:String);
    procedure runSF1(sI_CompCode,sI_CmdStatus,sI_ErrorCode,sI_ErrorMsg,sI_SeqNo : String;var fI_FunReturn,fI_RetCode : Double;var sI_Msg : String);
    procedure runSF2(sL_ResponseData, sI_CompCode,sI_CmdStatus,sI_ErrorCode,sI_ErrorMsg,sI_SeqNo : String;var fI_FunReturn,fI_RetCode : Double;var sI_Msg : String);
    procedure runSFSendCmdLog(sI_CompCode,sI_HighLevelCmdID,sI_FullCmdText,sI_User : String;var fI_FunReturn,fI_RetCode : Double;var sI_Msg : String);
    procedure runSFCmdResponseLog(sI_CompCode,sI_FullCmdResponseText : String;var fI_FunReturn,fI_RetCode : Double;var sI_Msg : String);

    function runSQL(nI_Mode:Integer; sI_CompCode, sI_SQL:String):TADOQuery;
    function getAreaDataNo(sI_CompCode : String) : String;
    function separateCaReturnString(sI_RemoteIP, sI_ReturnString : String) : String;

    procedure separateCaReturnType1(sI_ReturnString : String);
    procedure separateCaReturnType2(sI_ReturnString : String);
    procedure separateCaReturnType3(sI_ReturnString : String);

    procedure initialCdsDbLinkSet(I_StrList : TStringList);


  end;

var
  dtmMain: TdtmMain;

implementation

uses ConstU, frmLoginU, frmSysParamU, frmSendCommandU, Ustru;

{$R *.dfm}

{ TdtmMain }

procedure TdtmMain.activeCDS;
var
    sL_FileName, sL_ExePath : String;

begin

    sL_ExePath := dtmMain.getExePath;

    sL_FileName := sL_ExePath + '\' + USER_INFO_FILENAME;
    cdsUser.LoadFromFile(sL_FileName);
    if not cdsUser.Active then
      cdsUser.CreateDataSet;


    sL_FileName := sL_ExePath + '\' + DISPATCHER_SYS_PARAM_FILENAME;
    cdsParam.LoadFromFile(sL_FileName);
    if not cdsParam.Active then
      cdsParam.CreateDataSet;

    if not cdsDbLinkSet.Active then
      cdsDbLinkSet.CreateDataSet; 

end;





function TdtmMain.checkUserIdIsOnly(sI_UserID: String): Boolean;
var
    nL_Ndx : Integer;
begin
    //檢查此UserID是否為唯一值
    nL_Ndx := G_UserIDStrList.IndexOf(sI_UserID);

    if nL_Ndx <> -1 then
      Result := false
    else
      Result := true;
end;



function TdtmMain.getAllUserID: TStringList;
begin
    //將 User 資料存於 StringList
    G_UserIDStrList.Clear;

    cdsUser.First;
    while not cdsUser.Eof do
    begin
      G_UserIDStrList.Add(cdsUser.FieldByName('SID').AsString);

      cdsUser.Next;
    end;
end;

function TdtmMain.getExePath: String;
var
    sL_ExeFileName : String;
    sL_ExePath : String;
begin
    //取得執行檔路徑
    sL_ExeFileName := Application.ExeName;

    Result := ExtractFileDir(sL_ExeFileName);
end;

procedure TdtmMain.DataModuleCreate(Sender: TObject);
begin
    G_DisconnectProcessorIP := TStringList.Create;
    G_CompInfoStrList := TStringList.Create;
    G_ProcessorIPStrList := TStringList.Create;
    G_UserIDStrList := TStringList.Create;
    G_UserIDStrList.Clear;

    //getDbConnInfo;
    //connectToDB;
end;

procedure TdtmMain.DataModuleDestroy(Sender: TObject);
begin
    G_DisconnectProcessorIP.Free;
    G_UserIDStrList.Free;
    G_CompInfoStrList.Free;
    G_ProcessorIPStrList.Free;
end;



function TdtmMain.getCmdRefreshRate: Integer;
var
    sL_StartTime,sL_EndTime : String;
    dL_CurDateTime,dL_PureStartTime,dL_PureEndTime : TDateTime;
begin
    with cdsParam do
    begin
      sL_StartTime := FieldByName('sSHotTime').AsString;
      sL_EndTime := FieldByName('sEHotTime').AsString;
      transToPureDate(sL_StartTime,sL_EndTime,dL_PureStartTime,dL_PureEndTime);
      dL_CurDateTime := now;

      if (dL_CurDateTime >= dL_PureStartTime) AND (dL_CurDateTime <= dL_PureEndTime) then
        Result := cdsParam.FieldByName('nCmdRefreshRate1').AsInteger
      else
        Result := cdsParam.FieldByName('nCmdRefreshRate2').AsInteger
    end;

end;

procedure TdtmMain.transToPureDate(sI_StartTime, sI_EndTime: String;
  var dI_PureStartTime, dI_PureEndTime: TDateTime);
var
    sL_CurDate,sL_StartDateTime : String;
    sL_PureStartTime,sL_PureEndTime : String;
begin
    //將時間加上當日日期
    sL_CurDate := DateToStr(now);
    sL_PureStartTime := sL_CurDate + ' ' + Copy(sI_StartTime,1,2) + ':' + Copy(sI_StartTime,3,2);

    sL_PureEndTime := sL_CurDate + ' ' + Copy(sI_EndTime,1,2) + ':' + Copy(sI_EndTime,3,2);

    dI_PureStartTime := StrToDateTime(sL_PureStartTime);
    dI_PureEndTime := StrToDateTime(sL_PureEndTime);


end;

function TdtmMain.getDbConnInfo: String;
Var
  sL_IniFileName,sL_CompCodeString : String;
  L_IniFile : TIniFile;
  ii, nL_DB : Integer;
  sL_ALIAS,sL_USERID,sL_PASSWORD,sL_COMPCODE:String;
  sL_AliasTag,sL_UserIDTag,sL_PassWordTag,sL_CompCodeTag: String;
  sL_CompNameTag,sL_CompName : String;
  L_Obj : TCableNdsIniData;

  sL_ExeFileName,sL_ExePath,sL_ProcessorIPTag,sL_ProcessorIP : String;
begin
    sL_ExeFileName := Application.ExeName;

    sL_ExePath := ExtractFileDir(sL_ExeFileName);

    sL_IniFileName := sL_ExePath + '\' +  SYS_INI_FILE_NAME;

    if not FileExists(sL_IniFileName) then
    begin
      result := '-1: 讀取資料庫連線參數檔<'+sL_IniFileName +'>失敗';
      exit;
    end;

    L_IniFile := TIniFile.Create(sL_IniFileName);

    //取得參數檔之資料區總數
    nL_DB := L_IniFile.ReadInteger('DBINFO','DB_COUNT',0);
    if nL_DB=0 then
    begin
      result := '-1: 資料庫連線參數檔之資料區總數設定錯誤!<'+sL_IniFileName +'>';
      exit;
    end;

    //取得DISPATCHER_IP
    sG_DispatcherIP := L_IniFile.ReadString('CA_DISPATCHER_PROCESSOR','DISPATCHER_IP','');

    //取得PROCESSOR_LISTEN_PORT
    sG_ProcessorListenPort := L_IniFile.ReadString('CA_DISPATCHER_PROCESSOR','PROCESSOR_LISTEN_PORT','');

    //取得DISPATCHER_LISTEN_PORT
    sG_DispatcherListenPort := L_IniFile.ReadString('CA_DISPATCHER_PROCESSOR','DISPATCHER_LISTEN_PORT','');


    try

      for ii :=1 to nL_DB do
      begin
        sL_AliasTag := DATA_ALIAS + IntToStr(ii);
        sL_UserIDTag := DATA_USERID + IntToStr(ii);
        sL_PassWordTag := DATA_PASSWORD + IntToStr(ii);
        sL_CompCodeTag := DATA_COMPCODE + IntToStr(ii);
        sL_CompNameTag := DATA_COMPNAME + IntToStr(ii);
        sL_ProcessorIPTag := DATA_PROCESSOR_IP + IntToStr(ii);


        sL_ALIAS := L_IniFile.ReadString('DBINFO',sL_AliasTag,'');
        sL_USERID := L_IniFile.ReadString('DBINFO',sL_UserIDTag,'');
        sL_PASSWORD := L_IniFile.ReadString('DBINFO',sL_PassWordTag,'');
        sL_COMPCODE := L_IniFile.ReadString('DBINFO',sL_CompCodeTag,'');
        sL_COMPNAME := L_IniFile.ReadString('DBINFO',sL_CompNameTag,'');
        sL_ProcessorIP := L_IniFile.ReadString('DBINFO',sL_ProcessorIPTag,'');

        //將 ProcessorIP 不重複的存起來
        saveProcessorIP(sL_COMPCODE,sL_ProcessorIP);

        L_Obj := TCableNdsIniData.Create;
        L_Obj.nDataAreaNo := ii;
        L_Obj.sAlias := L_IniFile.ReadString('DBINFO',sL_AliasTag,'');
        L_Obj.sUserID := L_IniFile.ReadString('DBINFO',sL_UserIDTag,'');
        L_Obj.sPassword := L_IniFile.ReadString('DBINFO',sL_PassWordTag,'');
        L_Obj.sCompCode := L_IniFile.ReadString('DBINFO',sL_CompCodeTag,'');
        L_Obj.sCompName := L_IniFile.ReadString('DBINFO',sL_CompNameTag,'');
        L_Obj.sProcessorIP := L_IniFile.ReadString('DBINFO',sL_ProcessorIPTag,'');


        G_CompInfoStrList.AddObject(sL_CompCode, L_Obj);
      end;

      L_IniFile.Free;
    except
      result := '-1: 資料庫連線參數檔之資料區設定錯誤!<'+sL_IniFileName +'>';
      exit;
    end;


end;

procedure TdtmMain.saveProcessorIP(sI_CompCode, sI_ProcessorIP: String);
var
    ii,nL_Ndx : Integer;
    sL_OldCompCodeString,sL_NewCompCodeString : String;
    L_CompCodeObj : TProcessorIPData;
begin
{
    //串起ini中各ProcessorIP有多少公司使用
    nL_Ndx := G_ProcessorIPStrList.IndexOf(sI_CompCode);
    if nL_Ndx = -1 then
    begin
      L_CompCodeObj := TProcessorIPData.Create;
      L_CompCodeObj.sCompCodeString := sI_CompCode;
      L_CompCodeObj.sProcessorIp := sI_ProcessorIP;
      G_ProcessorIPStrList.AddObject(sI_CompCode,L_CompCodeObj);
    end;
}
    //串起ini中各ProcessorIP有多少公司使用
    nL_Ndx := G_ProcessorIPStrList.IndexOf(sI_ProcessorIP);
    if nL_Ndx = -1 then
    begin
      L_CompCodeObj := TProcessorIPData.Create;
      L_CompCodeObj.sCompCodeString := sI_CompCode;
      L_CompCodeObj.sProcessorIp := sI_ProcessorIP;
      G_ProcessorIPStrList.AddObject(sI_ProcessorIP,L_CompCodeObj);
    end
    else
    begin
      sL_OldCompCodeString := (G_ProcessorIPStrList.Objects[nL_Ndx] as TProcessorIPData).sCompCodeString;
      sL_NewCompCodeString := sL_OldCompCodeString + ',' + sI_CompCode;
      G_ProcessorIPStrList.Delete(nL_Ndx);

      L_CompCodeObj := TProcessorIPData.Create;
      L_CompCodeObj.sCompCodeString := sL_NewCompCodeString;
      L_CompCodeObj.sProcessorIp := sI_ProcessorIP;

      G_ProcessorIPStrList.AddObject(sI_ProcessorIP,L_CompCodeObj);
    end;

end;

function TdtmMain.getCompName(sI_CompCode: String): String;
var
    nL_Ndx,nL_AreaDatNo : Integer;
begin
    nL_Ndx := G_CompInfoStrList.IndexOf(sI_CompCode);
    if nL_Ndx<>-1 then
       Result := (G_CompInfoStrList.Objects[nL_Ndx] as TCableNdsIniData).sCompName;

end;

function TdtmMain.getProcessorIP(sI_CompCode: String): String;
var
    nL_Ndx,nL_AreaDatNo : Integer;
begin
    nL_Ndx := G_CompInfoStrList.IndexOf(sI_CompCode);
    if nL_Ndx<>-1 then
      Result := (G_CompInfoStrList.Objects[nL_Ndx] as TCableNdsIniData).sProcessorIP;
end;

function TdtmMain.getSendNdsData(sI_CompCode: String;
  var sI_Msg: String): TADOQuery;
var
    sL_SQL,sL_Where : String;
    L_AdoQry : TADOQuery;
begin
{ //原 Nagra
    sL_SQL := 'SELECT * FROM SEND_NDS';

    sL_Where := ' where (';
    sL_Where := sL_Where +'(CMD_STATUS=' + STR_SEP + 'W' + STR_SEP + ') or (CMD_STATUS =' + STR_SEP + 'P' + STR_SEP + ' and Operator=' + STR_SEP + BATCH_OPERATOR + STR_SEP + ')';

    if frmSendCommand.nG_MaxCommandCount<>0 then
      sL_Where := sL_Where + ' or (CMD_STATUS =' + STR_SEP + 'P' + STR_SEP + ' and nvl(ReSentTimes,0)<=' + IntToStr(frmSendCommand.nG_ReSentTimes) + ' and Operator<>' + STR_SEP + BATCH_OPERATOR + STR_SEP + ')';

    sL_Where := sL_Where + ') ' ;

    sL_Where := sL_Where + ' and COMPCODE=' + sI_CompCode;
    sL_Where := sL_Where + ' and (PROCESSINGDATE< ' + TUstr.getOracleSQLDateTimeStr(now) + ' or PROCESSINGDATE IS NULL)';
    sL_SQL := sL_SQL + sL_Where + ' order by CMD_STATUS, nvl(resenttimes,0) DESC, SEQNO ';
}

    sL_SQL := 'SELECT * FROM SEND_NDS';

    sL_Where := ' where (';
    sL_Where := sL_Where +'(CMDSTATUS=' + STR_SEP + 'W' + STR_SEP + ') or (CMDSTATUS =' + STR_SEP + 'P' + STR_SEP + ' and Operator=' + STR_SEP + BATCH_OPERATOR + STR_SEP + ')';

    if frmSendCommand.nG_MaxCommandCount<>0 then
      sL_Where := sL_Where + ' or (CMDSTATUS =' + STR_SEP + 'P' + STR_SEP + ' and nvl(ReSentTimes,0)<=' + IntToStr(frmSendCommand.nG_ReSentTimes) + ' and Operator<>' + STR_SEP + BATCH_OPERATOR + STR_SEP + ')';

    sL_Where := sL_Where + ') ' ;

    sL_Where := sL_Where + ' and COMPCODE=' + sI_CompCode;
    sL_Where := sL_Where + ' and (PROCESSINGDATE< ' + TUstr.getOracleSQLDateTimeStr(now) + ' or PROCESSINGDATE IS NULL)';
    sL_SQL := sL_SQL + sL_Where + ' order by CMDSTATUS, nvl(resenttimes,0) DESC, SEQNO ';

    L_AdoQry := runSQL(SELECT_MODE,sI_CompCode,sL_SQL);

    if L_AdoQry = nil then
    begin
      sI_Msg := NO_AVAILABLE_QUERY_COMP;
      exit;
    end
    else
    begin
      sI_Msg := '';
      //showmessage(IntToStr(L_AdoQry.RecordCount));
      Result := L_AdoQry;
    end;

end;

function TdtmMain.fillStringLength(sI_String, sI_FillStr: String;
  nI_Length: Integer): String;
var
    ii,nL_OldLen,nL_DiffLen : Integer;
    sL_NewStr : String;
begin
    if sI_FillStr = SPACE_STR then
      Result := format('%' + IntToStr(nI_Length) + 's', [sI_String])
    else if sI_FillStr = ZERO_STR then
      Result := format('%.' + IntToStr(nI_Length) + 'd', [StrToInt(sI_String)]);

end;

procedure TdtmMain.runSF1(sI_CompCode, sI_CmdStatus, sI_ErrorCode,
  sI_ErrorMsg, sI_SeqNo: String; var fI_FunReturn, fI_RetCode: Double;
  var sI_Msg: String);
var
  nL_ActiveConn,ii : Integer;
begin
  try
    nL_ActiveConn := StrToInt(getAreaDataNo(sI_CompCode))-1;

    //for ii:=0 to G_AdoStoredProcAry1[nL_ActiveConn].Parameters.Count do
    //0   showmessage(IntToStr(ii) + ' -- ' +  G_AdoStoredProcAry1[nL_ActiveConn].Parameters.Items[ii].Name);

    with G_AdoStoredProcAry1[nL_ActiveConn] do
    begin

      //showmessage(IntToStr(G_AdoStoredProcAry1[nL_ActiveConn].Parameters.Count));

      Parameters.ParamByName('p_ResponseData').Value := '';
      Parameters.ParamByName('p_SubscriberID').Value := '';
      Parameters.ParamByName('P_CMDSTATUS').Value := sI_CmdStatus;
      Parameters.ParamByName('P_ERRORCODE').Value := sI_ErrorCode;
      Parameters.ParamByName('P_ERRORMSG').Value := sI_ErrorMsg;
      Parameters.ParamByName('P_SEQNO').Value := sI_SeqNo;
      ExecProc;

      fI_FunReturn := Parameters.Items[0].Value; //用ParamByName會有問題
      fI_RetCode := Parameters.Items[6].Value; //用ParamByName會有問題
      sI_Msg := Parameters.Items[7].Value; //用ParamByName會有問題
    end;
  except

  end;

end;

function TdtmMain.runSQL(nI_Mode: Integer; sI_CompCode,
  sI_SQL: String): TADOQuery;
var
  L_AdoQuery : TADOQuery;
begin
  L_AdoQuery := getActiveQuery(sI_CompCode);
  if L_AdoQuery<>nil then
  begin
    L_AdoQuery.SQL.Clear;
    L_AdoQuery.SQL.Add(sI_SQL);
    case nI_Mode of
      SELECT_MODE://select
        L_AdoQuery.Open;
      IUD_MODE://DML
        L_AdoQuery.ExecSQL;
    end;
  end;
  result := L_AdoQuery;



end;

function TdtmMain.getAreaDataNo(sI_CompCode: String): String;
var
    nL_Ndx,nL_AreaDatNo : Integer;
begin
    nL_Ndx := G_CompInfoStrList.IndexOf(sI_CompCode);
    if nL_Ndx<>-1 then
    begin
       nL_AreaDatNo := (G_CompInfoStrList.Objects[nL_Ndx] as TCableNdsIniData).nDataAreaNo;

      getAreaDataNo := IntToStr(nL_AreaDatNo);
    end;


end;

function TdtmMain.connectToDB: WideString;
var
 ii : Integer;
 sL_DbConnStr,sL_ProcedureName : String;
 sL_DbPassword, sL_DbAlias, sL_DbUserName,sL_CompCode : String;
begin
  try
  setlength(G_DbConnAry, G_CompInfoStrList.Count);
  setlength(G_AdoQueryAry, G_CompInfoStrList.Count);

  setlength(G_AdoStoredProcAry1, G_CompInfoStrList.Count);
  setlength(G_AdoStoredProcAry2, G_CompInfoStrList.Count);
  //setlength(G_AdoStoredProcAry3, G_CompInfoStrList.Count);
  setlength(G_AdoSFSendCmdLogAry, G_CompInfoStrList.Count);
  setlength(G_AdoSFResponseLogAry, G_CompInfoStrList.Count);





  for ii:=0 to G_CompInfoStrList.Count-1  do
  begin
    sL_DbPassword := (G_CompInfoStrList.Objects[ii] as TCableNdsIniData).sPassword;
    sL_DbAlias := (G_CompInfoStrList.Objects[ii] as TCableNdsIniData).sAlias;
    sL_DbUserName := (G_CompInfoStrList.Objects[ii] as TCableNdsIniData).sUserID;
    sL_CompCode := (G_CompInfoStrList.Objects[ii] as TCableNdsIniData).sCompCode;

//    sL_DbConnStr := 'Provider=OraOLEDB.Oracle.1;Password=' + sL_DbPassword +';Persist Security Info=True;User ID='+sL_DbUserName+';Data Source='+sL_DbAlias; //=> 適用於my nb
    //sL_DbConnStr := 'Provider=MSDASQL ;Password=' + sL_DbPassword + ';User ID='+sL_DbUserName+';Data Source='+sL_DbAlias+';Persist Security Info=True'; //適用於其他 PC
    sL_DbConnStr := 'Provider=MSDAORA.1;Password=' + sL_DbPassword + ';User ID='+sL_DbUserName+';Data Source='+sL_DbAlias+';Persist Security Info=True'; //適用於其他 PC


    G_DbConnAry[ii] := TADOConnection.Create(nil);

    if not G_DbConnAry[ii].Connected then
    begin
      //ADOConnection 資料庫連線設定
      G_DbConnAry[ii].ConnectionString := sL_DbConnStr;
      G_DbConnAry[ii].LoginPrompt := false;
      G_DbConnAry[ii].Connected := true;


      //ADOQuery 連ADOConnection 設定
      G_AdoQueryAry[ii] :=  TADOQuery.Create(nil);
      G_AdoQueryAry[ii].Connection := G_DbConnAry[ii];
      G_AdoQueryAry[ii].Tag := StrToInt(sL_CompCode);

      //ADOStoredProc 連 ADOConnection 設定
      G_AdoStoredProcAry1[ii] := TADOStoredProc.Create(nil);
      G_AdoStoredProcAry1[ii].Connection := G_DbConnAry[ii];
      G_AdoStoredProcAry1[ii].ProcedureName := 'SF_UPDATESENDNDSCMDSTATUS';
      G_AdoStoredProcAry1[ii].Parameters.Refresh;

      G_AdoStoredProcAry2[ii] := TADOStoredProc.Create(nil);
      G_AdoStoredProcAry2[ii].Connection := G_DbConnAry[ii];
      G_AdoStoredProcAry2[ii].ProcedureName := 'SF_UPDATESENDNDSCMDSTATUS';
      G_AdoStoredProcAry2[ii].Parameters.Refresh;
{
      G_AdoStoredProcAry3[ii] := TADOStoredProc.Create(nil);
      G_AdoStoredProcAry3[ii].Connection := G_DbConnAry[ii];
      G_AdoStoredProcAry3[ii].ProcedureName := 'SF_UPDATESENDNDSCMDSTATUS';
      G_AdoStoredProcAry3[ii].Parameters.Refresh;
}

      //ADOStoredProc 連 ADOConnection 設定
      G_AdoSFSendCmdLogAry[ii] := TADOStoredProc.Create(nil);
      G_AdoSFSendCmdLogAry[ii].Connection := G_DbConnAry[ii];
      G_AdoSFSendCmdLogAry[ii].ProcedureName := 'SF_InsertSendCmdLog';
      G_AdoSFSendCmdLogAry[ii].Parameters.Refresh;


      //ADOStoredProc 連 ADOConnection 設定
      G_AdoSFResponseLogAry[ii] := TADOStoredProc.Create(nil);
      G_AdoSFResponseLogAry[ii].Connection := G_DbConnAry[ii];
      G_AdoSFResponseLogAry[ii].ProcedureName := 'SF_InsertCmdResponseLog';
      G_AdoSFResponseLogAry[ii].Parameters.Refresh;
    end;
  end;


except
  result := '-1: 與CC&B資料庫連線失敗, 資料庫別名=<'+sL_DbAlias+'>, DB User=<'+sL_DbUserName +'>';
  exit;
end;
//    savetofile('ok','c:\aa4.txt');
result := '';

end;

function TdtmMain.getActiveQuery(sI_AreaDataNo: String): TADOQuery;
var
    ii, nL_Pos : Integer;
    sL_CompCode : String;
begin
    nL_Pos := -1;
    for ii:=0 to dtmMain.G_CompInfoStrList.Count-1 do
    begin
      sL_CompCode := dtmMain.G_CompInfoStrList.Strings[ii];

      if sL_CompCode = sI_AreaDataNo then
      begin
        nL_Pos := StrToInt(getAreaDataNo(sL_CompCode)) - 1;
        break;
      end;
    end;
{
    for ii:=0 to G_DbAreaNameStrList.Count -1 do
    begin
      if (DATA_AREA_HEADER + sI_AreaDataNo)=G_DbAreaNameStrList.Strings[ii] then
      begin
        nL_Pos := ii;
        break;
      end;
    end;
}


    if nL_Pos<>-1 then
    begin
      result := G_AdoQueryAry[nL_Pos];
      savetofile('get query '+ sI_AreaDataNo, 'c:\bb4.txt');
    end
    else
    begin
      result := nil;
      savetofile('can not get query '+ sI_AreaDataNo, 'c:\bb5.txt');
    end;



end;

procedure TdtmMain.saveToFile(vI_Content: OleVariant; sI_FileName: String);
var
    L_StrList : TStringList;
begin
    L_StrList := TStringList.Create;;
    L_StrList.Add(vI_Content);
    L_StrList.SaveToFile(sI_FileName);
    L_StrList.Free;
end;

function TdtmMain.separateCaReturnString(sI_RemoteIP, sI_ReturnString: String): String;
var
    sL_TempReturnString,sL_TempNewString : String;
    nL_TotalLen,nL_StringType : Integer;
begin
    if sI_ReturnString= DISCONNECT_CMD then
    begin
      if G_DisconnectProcessorIP.IndexOf(sI_RemoteIP)<0 then
        G_DisconnectProcessorIP.Add(sI_RemoteIP);
    end
    else
    begin
      frmSendCommand.Memo1.Lines.Add(sI_ReturnString);

      //拆解CA傳回來的訊息 => 拆解方式請參考 => NDS_系統設計.doc
      {
      ..1C01903123000000009000001
      SEND_TO_DISPATCHER_INFO_TYPE_1 => 1(型別) C(指令處理結果) 01(公司代碼長度) 9(公司代碼) 03(SeqNo長度) 123(SeqNo) 000 (錯誤訊息代碼長度) [可能會有錯誤訊息代碼] 000 (錯誤訊息長度) [可能會有錯誤訊息代碼] 009000001(交易號碼)


      ..20190002A1006howard00900000301000302574428920030408N2003040720030409U1160194446005211911825900000
      SEND_TO_DISPATCHER_INFO_TYPE_2 => 2(型別)   01(公司代碼長度) 9(公司代碼) 0002(高階指令代碼長度) A1(高階指令代碼) 006(操作者長度) howard(操作者) 00900000301000302574428920030408N2003040720030409U1160194446005211911825900000 (指令字串)


      ..30190090000041000000008972050257000344289200303081000009000004000000000000000000000000
      SEND_TO_DISPATCHER_INFO_TYPE_3 => 3(型別)   01(公司代碼長度) 9(公司代碼) 009000004(交易號碼) 1000(回應之指令代號) 000008972050257000344289200303081000009000004000000000000000000000000(完整之回應字串)

      }
      try
        sL_TempReturnString := sI_ReturnString;

        nL_TotalLen := Length(sL_TempReturnString);

        nL_StringType := StrToInt(Copy(sL_TempReturnString,1,1));

        sL_TempNewString := Copy(sL_TempReturnString,2,nL_TotalLen);

        //showmessage(sL_TempNewString);
        if nL_StringType = SEND_TO_DISPATCHER_INFO_TYPE_1 then
        begin
          separateCaReturnType1(sL_TempNewString);
        end
        else if nL_StringType = SEND_TO_DISPATCHER_INFO_TYPE_2 then
        begin
          //frmSendNagra.bG_SaveCommandLog := true;
          if frmSendCommand.bG_SaveCommandLog then
            separateCaReturnType2(sL_TempNewString);
        end
        else if nL_StringType = SEND_TO_DISPATCHER_INFO_TYPE_3 then
        begin
          if frmSendCommand.bG_SaveResponseLog then
            separateCaReturnType3(sL_TempNewString);
        end;


      except

      end;
    end;

end;

procedure TdtmMain.separateCaReturnType1(sI_ReturnString: String);
var
    sL_CmdStatus,sL_CompCode,sL_SeqNo : String;
    sL_ErrCode,sL_ErrMsg,sL_TransNo : String;
    sL_TempReturnString :String;
    ii,nL_StartPos,nL_TotalLen,nL_CompCodeLen,nL_SeqNoLen : Integer;
    nL_ErrCodeLen,nL_ErrMsgLen : Integer;
    fL_FunReturn,fL_RetCode : Double;
    sL_ResponseData, sL_Msg : String;

    nL_CompCodeLength : Integer;
begin
    //拆解CA傳回來的訊息 => 如何拆解,請參考 => NDS_系統設計.doc

    try
      //C0190800123456000000 + subscriber information..
      sL_TempReturnString := sI_ReturnString;
      //jackaltest
      //sL_TempReturnString := 'C0170000000000000001';
      nL_TotalLen := Length(sL_TempReturnString);

      nL_StartPos := 1;
      sL_CmdStatus := Copy(sL_TempReturnString,nL_StartPos,1);

      nL_CompCodeLength := StrToInt(Copy(sL_TempReturnString,2,2));
      sL_CompCode := Copy(sL_TempReturnString,4,nL_CompCodeLength);

      nL_StartPos := 4 + nL_CompCodeLength;


      nL_SeqNoLen := StrToInt(Copy(sL_TempReturnString,nL_StartPos,2));
      nL_StartPos := nL_StartPos +2;
      sL_SeqNo := IntToStr(StrToInt(Copy(sL_TempReturnString,nL_StartPos,nL_SeqNoLen)));
      nL_StartPos := nL_StartPos + nL_SeqNoLen;


      nL_ErrCodeLen := StrToInt(Copy(sL_TempReturnString,nL_StartPos,3));
      nL_StartPos := nL_StartPos + 3;
      sL_ErrCode := Copy(sL_TempReturnString,nL_StartPos,nL_ErrCodeLen);
      nL_StartPos := nL_StartPos + nL_ErrCodeLen;

      nL_ErrMsgLen := StrToInt(Copy(sL_TempReturnString,nL_StartPos,3));
      nL_StartPos := nL_StartPos + 3;
      sL_ErrMsg := Copy(sL_TempReturnString,nL_StartPos,nL_ErrMsgLen);

      nL_StartPos := nL_StartPos + nL_ErrMsgLen;
      sL_ResponseData := Copy(sL_TempReturnString,nL_StartPos,length(sL_TempReturnString)-nL_StartPos+1);
      //showmessage(sL_TempReturnString + '   (拆解完成)');

      frmSendCommand.updateReturnDataUI(sL_CompCode, sL_SeqNo, sL_ErrCode, sL_ErrMsg);
      //呼叫  將CMD_STATUS  P 改為 C 或 E
      dtmMain.runSF1(sL_CompCode,sL_CmdStatus,sL_ErrCode,sL_ErrMsg,Trim(sL_SeqNo),fL_FunReturn,fL_RetCode,sL_Msg);
      runSF2(sL_ResponseData, sL_CompCode,sL_CmdStatus,sL_ErrCode,sL_ErrMsg,sL_SeqNo,fL_FunReturn,fL_RetCode,sL_Msg);


//*******************************************************************
      if sL_Msg <> '更新指令處理狀態完畢' then
      begin
        {
        dtmMain.cdsTemp.Append;
        dtmMain.cdsTemp.FieldByName('COL1').AsString := sL_SEQNO;
        //dtmMain.cdsTemp.FieldByName('COL4').AsString := sI_ProcessorIP;
        //dtmMain.cdsTemp.FieldByName('COL3').AsString := sI_CompName;

        sL_Msg := 'ReturnNO:  ' + FloatToStr(fL_FunReturn) + '  sL_Msg: ' + sL_Msg;

        dtmMain.cdsTemp.FieldByName('COL2').AsString := sL_Msg;
        dtmMain.cdsTemp.Post;

        frmSendCommand.DBGrid1.Refresh;
        }
      end
      else
      begin
        //nG_Count := nG_Count + 1;
        //frmSendNagra.Label1.Caption := IntToStr(nG_Count);

      end;
//*******************************************************************

//      frmSendCommand.Memo1.Lines.Add(sL_TempReturnString);
      //showmessage('SF執行完成   sL_TransNo ' + sL_TransNo);

    except

    end;


end;

procedure TdtmMain.separateCaReturnType2(sI_ReturnString: String);
var
    sL_TempReturnString,sL_CompCode,sL_TempNewString : String;
    sL_HighLevelCmdID,sL_User,sL_FullCmdText : String;
    nL_TotalLen,nL_FirstPoint,nL_CompCodeLen,nL_HighCodeLen : Integer;
    nL_UserLen : Integer;
    fL_FunReturn,fL_RetCode : Double;
    sL_Msg : String;
begin
{
    ..20190002A1006howard00900000301000302574428920030408N2003040720030409U1160194446005211911825900000
    SEND_TO_DISPATCHER_INFO_TYPE_2 => 2(型別)   01(公司代碼長度) 9(公司代碼) 0002(高階指令代碼長度) A1(高階指令代碼) 006(操作者長度) howard(操作者) 00900000301000302574428920030408N2003040720030409U1160194446005211911825900000 (指令字串)
}
      sL_TempReturnString := sI_ReturnString;
      nL_TotalLen := Length(sL_TempReturnString);


      //取公司代碼
      nL_FirstPoint := 1;
      nL_CompCodeLen := StrToInt(Copy(sL_TempReturnString,1,2));
      sL_CompCode := Copy(sL_TempReturnString,3,nL_CompCodeLen);

      //取高階指令代碼
      nL_FirstPoint := nL_FirstPoint + 2 + nL_CompCodeLen;
      sL_TempNewString := Copy(sL_TempReturnString,nL_FirstPoint,nL_TotalLen);
      nL_HighCodeLen := StrToInt(Copy(sL_TempNewString,1,4));
      sL_HighLevelCmdID := Copy(sL_TempNewString,5,nL_HighCodeLen);


      //取操作者
      nL_FirstPoint := nL_FirstPoint + 4 + nL_HighCodeLen;
      sL_TempNewString := Copy(sL_TempReturnString,nL_FirstPoint,nL_TotalLen);
      nL_UserLen := StrToInt(Copy(sL_TempNewString,1,3));
      sL_User := Copy(sL_TempNewString,4,nL_UserLen);


      //取指令字串
      nL_FirstPoint := nL_FirstPoint + 3 + nL_UserLen;
      sL_FullCmdText := Copy(sL_TempReturnString,nL_FirstPoint,nL_TotalLen);

      runSFSendCmdLog(sL_CompCode,sL_HighLevelCmdID,sL_FullCmdText,sL_User,fL_FunReturn,fL_RetCode,sL_Msg);
      showmessage(FloatToStr(fL_FunReturn) + ' == ' + FloatToStr(fL_RetCode) + ' -- ' + sL_Msg);

end;

procedure TdtmMain.separateCaReturnType3(sI_ReturnString: String);
var
    sL_TempReturnString,sL_CompCode,sL_TempNewString : String;
    sL_TransNo,sL_CommandCode,sL_FullCmdResponseText : String;
    nL_TotalLen,nL_FirstPoint,nL_CompCodeLen : Integer;
    fL_FunReturn,fL_RetCode : Double;
    sL_Msg : String;
begin
{
    ..30190090000041000000008972050257000344289200303081000009000004000000000000000000000000
    SEND_TO_DISPATCHER_INFO_TYPE_3 => 3(型別)   01(公司代碼長度) 9(公司代碼) 009000004(交易號碼) 1000(回應之指令代號) 000008972050257000344289200303081000009000004000000000000000000000000(完整之回應字串)
}
      sL_TempReturnString := sI_ReturnString;
      nL_TotalLen := Length(sL_TempReturnString);


      //取公司代碼
      nL_FirstPoint := 1;
      nL_CompCodeLen := StrToInt(Copy(sL_TempReturnString,1,2));
      sL_CompCode := Copy(sL_TempReturnString,3,nL_CompCodeLen);

      //取交易號碼(固定9碼)
      nL_FirstPoint := nL_FirstPoint + 2 + nL_CompCodeLen;
      sL_TempNewString := Copy(sL_TempReturnString,nL_FirstPoint,nL_TotalLen);
      sL_TransNo := Copy(sL_TempNewString,1,9);


      //取回應之指令代號(固定4碼)
      nL_FirstPoint := nL_FirstPoint + 9;
      sL_TempNewString := Copy(sL_TempReturnString,nL_FirstPoint,nL_TotalLen);
      sL_CommandCode := Copy(sL_TempNewString,1,4);


      //取完整之回應字串
      nL_FirstPoint := nL_FirstPoint + 4;
      sL_FullCmdResponseText := Copy(sL_TempReturnString,nL_FirstPoint,nL_TotalLen);

      runSFCmdResponseLog(sL_CompCode,sL_FullCmdResponseText,fL_FunReturn,fL_RetCode,sL_Msg);

      showmessage('Type3  == ' + FloatToStr(fL_FunReturn) + ' == ' + FloatToStr(fL_RetCode) + ' -- ' + sL_Msg);


end;

procedure TdtmMain.runSFSendCmdLog(sI_CompCode, sI_HighLevelCmdID,
  sI_FullCmdText, sI_User: String; var fI_FunReturn, fI_RetCode: Double;
  var sI_Msg: String);
var
  nL_ActiveConn : Integer;
begin
  try
    nL_ActiveConn := StrToInt(getAreaDataNo(sI_CompCode))-1;

    with G_AdoSFSendCmdLogAry[nL_ActiveConn] do
    begin

      Parameters.ParamByName('P_HighLevelCmdID').Value := sI_HighLevelCmdID;
      Parameters.ParamByName('P_FullCmdText').Value := sI_FullCmdText;
      Parameters.ParamByName('P_Operator').Value := sI_User;
      ExecProc;

      fI_FunReturn := Parameters.Items[0].Value; //用ParamByName會有問題
      fI_RetCode := Parameters.Items[4].Value; //用ParamByName會有問題
      sI_Msg := Parameters.Items[5].Value; //用ParamByName會有問題
    end;
  except

  end;

end;

procedure TdtmMain.runSFCmdResponseLog(sI_CompCode,
  sI_FullCmdResponseText: String; var fI_FunReturn, fI_RetCode: Double;
  var sI_Msg: String);
var
  nL_ActiveConn : Integer;
begin
  try
    nL_ActiveConn := StrToInt(getAreaDataNo(sI_CompCode))-1;

    with G_AdoSFResponseLogAry[nL_ActiveConn] do
    begin

      Parameters.ParamByName('p_FullCmdResponseText').Value := sI_FullCmdResponseText;
      ExecProc;

      fI_FunReturn := Parameters.Items[0].Value; //用ParamByName會有問題
      fI_RetCode := Parameters.Items[2].Value; //用ParamByName會有問題
      sI_Msg := Parameters.Items[3].Value; //用ParamByName會有問題
    end;
  except

  end;

end;

procedure TdtmMain.runSF2(sL_ResponseData, sI_CompCode, sI_CmdStatus, sI_ErrorCode,
  sI_ErrorMsg, sI_SeqNo: String; var fI_FunReturn, fI_RetCode: Double;
  var sI_Msg: String);
var
  nL_ActiveConn : Integer;
begin
  try
    nL_ActiveConn := StrToInt(getAreaDataNo(sI_CompCode))-1;

    with G_AdoStoredProcAry2[nL_ActiveConn] do
    begin

      Parameters.ParamByName('P_RESPONSEDATA').Value := sL_ResponseData;
      Parameters.ParamByName('P_SUBSCRIBERID').Value := '';
      Parameters.ParamByName('P_CMDSTATUS').Value := sI_CmdStatus;
      Parameters.ParamByName('P_ERRORCODE').Value := sI_ErrorCode;
      Parameters.ParamByName('P_ERRORMSG').Value := sI_ErrorMsg;
      Parameters.ParamByName('P_SEQNO').Value := sI_SeqNo;
      ExecProc;

      fI_FunReturn := Parameters.Items[0].Value; //用ParamByName會有問題
      fI_RetCode := Parameters.Items[6].Value; //用ParamByName會有問題
      sI_Msg := Parameters.Items[7].Value; //用ParamByName會有問題
    end;
  except

  end;

end;

function TdtmMain.isDisconnect(sI_ProcessorIP: String): boolean;
var
    bL_Result : boolean;
begin
    bL_Result := false;
    if G_DisconnectProcessorIP.IndexOf(sI_ProcessorIP)>=0 then
      bL_Result := true;
    result := bL_Result;

end;

procedure TdtmMain.hasConnected(sI_ProcessorIP: String);
var
    nL_Ndx : Integer;
begin
    nL_Ndx := G_DisconnectProcessorIP.IndexOf(sI_ProcessorIP);
    if nL_Ndx >=0 then
    begin
      G_DisconnectProcessorIP.Delete(nL_Ndx);
    end;

end;

procedure TdtmMain.initialCdsDbLinkSet(I_StrList: TStringList);
var
    ii : Integer;
    L_IniStrList1,L_IniStrList2 : TStringList;
    sL_Value,sL_Group,sL_ColumnName : String;
begin
    L_IniStrList1 := TStringList.Create;
    L_IniStrList2 := TStringList.Create;

    for ii:=1 to I_StrList.Count-1 do
    begin
      //將 ALIAS_1=mis 拆解成  [ALIAS_1], [mis]=Value, [ALIAS]=ColumnName,[1]=Group
      L_IniStrList1 := TUstr.ParseStrings(I_StrList[ii],'=');
      sL_Value := L_IniStrList1[1];

      L_IniStrList2 := TUstr.ParseStrings(L_IniStrList1[0],'_');
      sL_ColumnName := L_IniStrList2[0];
      sL_Group := L_IniStrList2[1];

      with cdsDbLinkSet do
      begin
        if Locate('Group',VarArrayOf([sL_Group]) , []) then
        begin
          Edit;
          FieldByName(sL_ColumnName).AsString := sL_Value;
          Post;
        end
        else
        begin
          Append;
          FieldByName('Group').AsString := sL_Group;
          FieldByName(sL_ColumnName).AsString := sL_Value;
          Post;
        end;
      end;
      //showmessage(IntToStr(ii) + ' = ' + I_StrList[ii] + '---' + sL_Group + '/' + sL_Value);
    end;

    cdsDbLinkSet.First;

    L_IniStrList1.Free;
    L_IniStrList2.Free;
end;

end.
