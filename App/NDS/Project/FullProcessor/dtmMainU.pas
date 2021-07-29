unit dtmMainU;

interface

uses
  SysUtils, Classes, DB, DBClient, ADODB, Variants ,IniFiles ,Forms ,Dialogs,XMLIntf;

type
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
    cdsParam: TClientDataSet;
    cdsParamsServerIP2: TStringField;
    cdsParamnSPortNo2: TIntegerField;
    cdsParamnRPortNo2: TIntegerField;
    cdsParamsSysName2: TStringField;
    cdsParamsSysVersion2: TStringField;
    cdsParamsLogPath: TStringField;
    cdsParamnTimeOut2: TIntegerField;
    cdsParambCommandLog2: TBooleanField;
    cdsParambResponseLog: TBooleanField;
    cdsParamdUptTime2: TDateTimeField;
    cdsParamsUptName2: TStringField;
    cdsParambShowUI2: TBooleanField;
    cdsParamnVersion: TIntegerField;
    cdsParamsSecuritytype: TStringField;
    cdsParamnToID: TIntegerField;
    cdsParamnConnectionID: TIntegerField;
    cdsParamnForwardTolerance: TIntegerField;
    cdsParamnBackwardTolerance: TIntegerField;
    cdsParamnCurrency: TIntegerField;
    cdsParamnFromID: TIntegerField;
    cdsParamnTimeZoneOffset: TIntegerField;
    cdsParamsCountryNumber: TStringField;
    cdsParamnTimeZoneOffset2: TIntegerField;
    cdsUser: TClientDataSet;
    cdsUsersID: TStringField;
    cdsUsersName: TStringField;
    cdsUsersPassword: TStringField;
    cdsParamsKey: TStringField;
    cdsParamnPopulationID: TIntegerField;
    cdsParamnMaxCommandCount: TIntegerField;
    cdsParamnCmdResentTimes: TIntegerField;
    cdsParamsSHotTime: TStringField;
    cdsParamsEHotTime: TStringField;
    cdsParamnCmdRefreshRate1: TIntegerField;
    cdsParamnCmdRefreshRate2: TIntegerField;
    cdsDbLinkSet: TClientDataSet;
    cdsDbLinkSetGroup: TStringField;
    cdsDbLinkSetCOMPNAME: TStringField;
    cdsDbLinkSetALIAS: TStringField;
    cdsDbLinkSetUSERID: TStringField;
    cdsDbLinkSetPASSWORD: TStringField;
    cdsDbLinkSetPROCESSORIP: TStringField;
    cdsDbLinkSetCOMPCODE: TStringField;
    adoQrySendNdsData: TADOQuery;
    cdsXmlData: TClientDataSet;
    cdsXmlDataSeqNO: TStringField;
    cdsXmlDataCompCode: TIntegerField;
    cdsXmlDataCompName: TStringField;
    cdsXmlDataCommandID: TStringField;
    cdsXmlDataSubscriberID: TStringField;
    cdsXmlDataIccNo: TStringField;
    cdsXmlDataSubBeginDate: TDateTimeField;
    cdsXmlDataSubEndDate: TDateTimeField;
    cdsXmlDataPinCode: TIntegerField;
    cdsXmlDataPopulationID: TIntegerField;
    cdsXmlDataRegionKey: TStringField;
    cdsXmlDataReportbackAvailability: TStringField;
    cdsXmlDataReportbackDate: TIntegerField;
    cdsXmlDataNotes: TStringField;
    cdsXmlDataCmdStatus: TStringField;
    cdsXmlDataErrorCode: TStringField;
    cdsXmlDataErrorDesc: TStringField;
    cdsXmlDataOperator: TStringField;
    cdsXmlDataResentTimes: TIntegerField;
    cdsXmlDataProcessingDate: TDateTimeField;
    cdsXmlDataUpdTime: TDateTimeField;
    cdsXmlDataPinControl: TStringField;
    cdsXmlDataServiceID: TStringField;
    cdsXmlDataResponseData: TStringField;
    cdsXmlDataClientIP: TStringField;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
    G_DbConnAry : array of TADOConnection;
    G_AdoQueryAry : array of TADOQuery;
    G_AdoStoredProcAry1 : array of TADOStoredProc; //SF_UPDATESENDNDSCMDSTATUS
    G_AdoStoredProcAry2 : array of TADOStoredProc; //SF_UPDATESENDNDSCMDSTATUS
    G_AdoSFSendCmdLogAry : array of TADOStoredProc; //SF_InsertCmdResponseLog
    G_AdoSFResponseLogAry : array of TADOStoredProc; //SF_InsertCmdResponseLog
    procedure transToPureDate(sI_StartTime,sI_EndTime : String;var dI_PureStartTime,dI_PureEndTime : TDateTime);    
  public
    { Public declarations }
    G_CompInfoStrList : TStringList;
    procedure saveToFile(vI_Content:OleVariant; sI_FileName:String);
    procedure initialCdsDbLinkSet(I_StrList : TStringList);

    function getDbConnInfo: String;
    function connectToDB: WideString;
    function getCmdRefreshRate : Integer;
    function getCompName(sI_CompCode : String) : String;
    function getSendNdsData(sI_CompCode : String;var sI_Msg: String) : TADOQuery;
    function runSQL(nI_Mode:Integer; sI_CompCode, sI_SQL:String):TADOQuery;
    function getActiveQuery(sI_AreaDataNo : String):TADOQuery;
    function getAreaDataNo(sI_CompCode : String) : String;
    function fillStringLength(sI_String,sI_FillStr : String; nI_Length : Integer): String;
    procedure runSF1(sI_CompCode,sI_CmdStatus,sI_ErrorCode,sI_ErrorMsg,sI_SeqNo : String;var fI_FunReturn,fI_RetCode : Double;var sI_Msg : String);
    procedure runSF2(sL_ResponseData, sI_CompCode,sI_CmdStatus,sI_ErrorCode,sI_ErrorMsg,sI_SeqNo : String;var fI_FunReturn,fI_RetCode : Double;var sI_Msg : String);

    //runSF1,runSF2合併成runSF(由傳入參數分別)
    procedure runSF(sI_ResponseData, sI_CompCode,sI_CmdStatus,sI_ErrorCode,sI_ErrorMsg,sI_SeqNo : String;var fI_FunReturn,fI_RetCode : Double;var sI_Msg : String);


    procedure runSFSendCmdLog(sI_CompCode,sI_HighLevelCmdID,sI_FullCmdText,sI_User : String;var fI_FunReturn,fI_RetCode : Double;var sI_Msg : String);
    procedure runSFCmdResponseLog(sI_CompCode,sI_FullCmdResponseText : String;var fI_FunReturn,fI_RetCode : Double;var sI_Msg : String);            

    function separateCaReturnString(sI_ReturnString : String) : String;
    procedure separateCaReturnType1(sI_ReturnString : String);
    procedure separateCaReturnType2(sI_ReturnString : String);
    procedure separateCaReturnType3(sI_ReturnString : String);


  end;

var
  dtmMain: TdtmMain;

implementation

uses Ustru, ConstU, frmMainU, TListernerU;

{$R *.dfm}

{ TdtmMain }

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

    sL_IniFileName := sL_ExePath + '\' +  CA_INI_FILE_NAME;

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
    //sG_DispatcherIP := L_IniFile.ReadString('CA_DISPATCHER_PROCESSOR','DISPATCHER_IP','');

    //取得PROCESSOR_LISTEN_PORT
    //sG_ProcessorListenPort := L_IniFile.ReadString('CA_DISPATCHER_PROCESSOR','PROCESSOR_LISTEN_PORT','');

    //取得DISPATCHER_LISTEN_PORT
    //sG_DispatcherListenPort := L_IniFile.ReadString('CA_DISPATCHER_PROCESSOR','DISPATCHER_LISTEN_PORT','');


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
        //saveProcessorIP(sL_COMPCODE,sL_ProcessorIP);

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

procedure TdtmMain.saveToFile(vI_Content: OleVariant; sI_FileName: String);
var
    L_StrList : TStringList;
begin
    L_StrList := TStringList.Create;;
    L_StrList.Add(vI_Content);
    L_StrList.SaveToFile(sI_FileName);
    L_StrList.Free;
end;

procedure TdtmMain.DataModuleCreate(Sender: TObject);
begin
    G_CompInfoStrList := TStringList.Create;
end;

procedure TdtmMain.DataModuleDestroy(Sender: TObject);
begin
    G_CompInfoStrList.Free;
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

function TdtmMain.getCompName(sI_CompCode: String): String;
var
    nL_Ndx,nL_AreaDatNo : Integer;
begin
    nL_Ndx := G_CompInfoStrList.IndexOf(sI_CompCode);
    if nL_Ndx<>-1 then
       Result := (G_CompInfoStrList.Objects[nL_Ndx] as TCableNdsIniData).sCompName;
end;

function TdtmMain.getSendNdsData(sI_CompCode: String;
  var sI_Msg: String): TADOQuery;
var
    sL_SQL,sL_Where : String;
    L_AdoQry : TADOQuery;
begin
    sL_SQL := 'SELECT * FROM SEND_NDS';

    sL_Where := ' where (';
    sL_Where := sL_Where +'(CMDSTATUS=' + STR_SEP + 'W' + STR_SEP + ') or (CMDSTATUS =' + STR_SEP + 'P' + STR_SEP + ' and Operator=' + STR_SEP + BATCH_OPERATOR + STR_SEP + ')';

    if frmMain.nG_MaxCommandCount<>0 then
      sL_Where := sL_Where + ' or (CMDSTATUS =' + STR_SEP + 'P' + STR_SEP + ' and nvl(ReSentTimes,0)<=' + IntToStr(frmMain.nG_ReSentTimes) + ' and Operator<>' + STR_SEP + BATCH_OPERATOR + STR_SEP + ')';

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
      Result := L_AdoQry;
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

function TdtmMain.separateCaReturnString(sI_ReturnString: String): String;
var
    sL_TempReturnString,sL_TempNewString : String;
    nL_TotalLen,nL_StringType : Integer;
begin
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

      if nL_StringType = SEND_TO_DISPATCHER_INFO_TYPE_1 then
      begin
        separateCaReturnType1(sL_TempNewString);
      end
      else if nL_StringType = SEND_TO_DISPATCHER_INFO_TYPE_2 then
      begin
        if frmMain.bG_CommandLog then
          separateCaReturnType2(sL_TempNewString);
      end
      else if nL_StringType = SEND_TO_DISPATCHER_INFO_TYPE_3 then
      begin
        if frmMain.bG_ResponseLog then
          separateCaReturnType3(sL_TempNewString);
      end;
    except

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
    sL_ResponseData, sL_Msg ,sL_CommandID,sL_ClientIP : String;
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

      if frmMain.bG_IsDBData then
      begin
        //呼叫  將CMD_STATUS  P 改為 C 或 E
        runSF(sL_ResponseData, sL_CompCode,sL_CmdStatus,sL_ErrCode,sL_ErrMsg,Trim(sL_SeqNo),fL_FunReturn,fL_RetCode,sL_Msg);
      end
      else
      begin

        if dtmMain.cdsXmlData.Locate('SeqNo',VarArrayOf([Trim(sL_SeqNo)]),[]) then
        begin
          sL_CommandID := dtmMain.cdsXmlData.FieldByName('COMMANDID').AsString;
          sL_ClientIP := dtmMain.cdsXmlData.FieldByName('ClientIP').AsString;
          frmMain.returnXMLResponseData(sL_CompCode,Trim(sL_SeqNo),sL_CommandID,sL_ErrCode,sL_ErrMsg,sL_ResponseData,sL_ClientIP);

          dtmMain.cdsXmlData.Delete;
        end;
      end;

      frmMain.updateReturnDataUI(sL_CompCode, sL_SeqNo, sL_ErrCode, sL_ErrMsg);
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

procedure TdtmMain.runSF2(sL_ResponseData, sI_CompCode, sI_CmdStatus,
  sI_ErrorCode, sI_ErrorMsg, sI_SeqNo: String; var fI_FunReturn,
  fI_RetCode: Double; var sI_Msg: String);
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

procedure TdtmMain.runSF(sI_ResponseData, sI_CompCode, sI_CmdStatus,
  sI_ErrorCode, sI_ErrorMsg, sI_SeqNo: String; var fI_FunReturn,
  fI_RetCode: Double; var sI_Msg: String);
var
  nL_ActiveConn : Integer;
begin
  try
    nL_ActiveConn := StrToInt(getAreaDataNo(sI_CompCode))-1;

    with G_AdoStoredProcAry2[nL_ActiveConn] do
    begin

      Parameters.ParamByName('P_RESPONSEDATA').Value := sI_ResponseData;
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

end.
