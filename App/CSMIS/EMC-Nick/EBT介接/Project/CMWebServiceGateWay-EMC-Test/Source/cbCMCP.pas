unit cbCMCP;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, DateUtils, ActiveX, DB, ADODB,
  IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient, IdHTTP,
  cbDataController, cbSoInfo;

type

  TCmdType = (ctOnline, ctBatch);

  TCMCPThread = class(TThread)
  private
    { Private declarations }
    FMsgCodes: String;
    FWebServiceCall: TCMWebServiceManager;
    FDataController: TDataController;
    FSoInfo: TSoInfo;
    FConfigInfo: TConfigInfo;
    FCallList: TList;
    FParseList: TList;
    FMsg: PMsg;
    FActiveSO311: PSO311;
    FLastDataRead: TDateTime;
    FLastXmlParse: TDateTime;
    FSOAPHttp: TIdHTTP;
    FRequestList: TStringList;
    function ConnectToCMCPDb: Boolean;
    function ConnectToLogDb: Boolean;
    function CheckCMMCPDbStatus: Boolean;
    function CheckLogDbStatus: Boolean;
    function CheckCanGetData: Boolean;
    function CheckCanParseXml: Boolean;
    function GetCmdSeqLikeStr: String;
    function GetCmdSeqLikeStrEx(const aPreMonth: Boolean): String;
    function GetDataSql(const aCmdType: TCmdType): String;
    function GetParseSql: String;
    function DataToCallList(const aCmdType: TCmdType): Integer;
    function DataToParseList: Integer;
    procedure DisconnectFromCMCPDb;
    procedure DisconnectFromLogDb;
    procedure ReleaseCallList;
    procedure ReleaseParseList;
    procedure MsgNotify;
    procedure RecordNotify;
    procedure DbNotify;
    procedure CallCMWebService;
    procedure ParseCMWebService;
    procedure SaveXml(aSO311: PSO311);
    procedure SaveToLog(aSO311: PSO311);
    function ParseXml(aSO311: PSO311): Boolean;
    procedure UpdateBack(aSo311: PSO311);
    function UpdateOtherData(aSo311: PSO311): Boolean;
    function IsSpecialCmd(aSo311: PSO311): Boolean;
    function UpdateCMReg1(aSo311: PSO311): Boolean;  // 開機回填
    function UpdateCMReg2(aSo311: PSO311): Boolean;  // 關機回填
    function UpdateCPIPReg(aSo311: PSO311): Boolean; // CP Ip 申請回填
  protected
    procedure Execute; override;
  public
    constructor Create(aSo: TSoInfo; aConfig: TConfigInfo); reintroduce;
    destructor Destroy; override;
  end;

implementation

uses cbMain, cbUtils;

{ ---------------------------------------------------------------------------- }

{ TCMCPThread }

constructor TCMCPThread.Create(aSo: TSoInfo; aConfig: TConfigInfo);
begin
  inherited Create( True );
  Self.FreeOnTerminate := False;
  FSoInfo := aSo;
  FConfigInfo := aConfig;
  FCallList := TList.Create;
  FParseList := TList.Create;
  FRequestList := TStringList.Create;
  FDataController := TDataController.Create( nil );
  FSOAPHttp := TIdHTTP.Create( nil );
  New( FMsg );
  FActiveSO311 := nil;
  FWebServiceCall := TCMWebServiceManager.Create( FDataController.XMLDoc,
    FSoInfo.Owner );
  FMsgCodes :=
    ' ''02'', ''12'', ''20'', ''25'', ''13'', ''21'', ''22'', ' +
    ' ''23'', ''26'', ''27'' ';
end;

{ ---------------------------------------------------------------------------- }

destructor TCMCPThread.Destroy;
begin
  FCallList.Free;
  FParseList.Free;
  FDataController.Free;
  FConfigInfo := nil;
  FSoInfo := nil;
  FSOAPHttp.Free;
  Dispose( FMsg );
  FWebServiceCall.Free;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TCMCPThread.MsgNotify;
begin
  Main.MsgNotify( FMsg );
end;

{ ---------------------------------------------------------------------------- }

procedure TCMCPThread.RecordNotify;
begin
  if Assigned( FActiveSO311 ) then
    Main.RecordNotify( FSoInfo, FActiveSO311 );
end;

{ ---------------------------------------------------------------------------- }

procedure TCMCPThread.DbNotify;
begin
  Main.DbNotify( FSoInfo );
end;

{ ---------------------------------------------------------------------------- }

function TCMCPThread.CheckCMMCPDbStatus: Boolean;
var
  aError: Boolean;
begin
  aError := False;
  FDataController.CMDataReader.Close;
  FDataController.CMDataReader.SQL.Text := ' SELECT 1 FROM DUAL ';
  try
    FDataController.CMDataReader.Open;
  except
    aError := True;
  end;
  FDataController.CMDataReader.Close;
  if aError then
  begin
    FMsg.Kind := MB_ICONINFORMATION;
    FMsg.Msg := Format( '系統台[%s]資料庫重新連結。', [FSoInfo.CompName] );
    Synchronize( MsgNotify );
    ConnectToCMCPDb;
  end;  
  Result := FDataController.CMConnection.Connected;
end;

{ ---------------------------------------------------------------------------- }

function TCMCPThread.CheckLogDbStatus: Boolean;
begin
  if not FDataController.LogConnection.Connected then
  begin
    FMsg.Kind := MB_ICONINFORMATION;
    FMsg.Msg := '重新初使化Log記錄檔。';
    Synchronize( MsgNotify );
    ConnectToLogDb;
  end;
  Result := FDataController.LogConnection.Connected;
end;

{ ---------------------------------------------------------------------------- }

function TCMCPThread.ConnectToCMCPDb: Boolean;
begin
  if FDataController.CMConnection.Connected then
    FDataController.CMConnection.Connected := False;
  FSoInfo.DbStatus := dbNone;
  Synchronize( DbNotify );
  try
    FDataController.CMConnection.ConnectionString := Format(
      'Provider=MSDAORA.1;Password=%s;User ID=%s;Data Source=%s;Persist Security Info=True',
      [FSoInfo.Password, FSoInfo.Owner, FSoInfo.Sid] );
    FDataController.CMConnection.Open;
    FMsg.Kind := MB_OK;
    FMsg.Msg := Format( '系統台: %s, 資料庫連結完成。', [FSoInfo.CompName] );
  except
    on E: Exception do
    begin
      FMsg.Kind := MB_ICONERROR;
      FMsg.Msg := Format( '系統台: %s, 資料庫連結有誤, 原因: %s。',
        [FSoInfo.CompName, E.Message] );
    end;
  end;
  Synchronize( MsgNotify );
  Result := FDataController.CMConnection.Connected;
  if Result then
    FSoInfo.DbStatus := dbOK
  else
    FSoInfo.DbStatus := dbNone;
  Synchronize( DbNotify );
end;

{ ---------------------------------------------------------------------------- }

function TCMCPThread.ConnectToLogDb: Boolean;
begin
  Result := True;
  { 沒用到 log }
  Exit;
  if FDataController.LogConnection.Connected then
    FDataController.LogConnection.Connected := False;
  try
    FDataController.LogConnection.Connected := True;
    FMsg.Kind := MB_OK;
    FMsg.Msg := 'Log記錄檔初使化完成。';
  except
    on E: Exception do
    begin
      FMsg.Kind := MB_ICONERROR;
      FMsg.Msg := Format( 'Log記錄檔初使失敗, 原因: %s。', [E.Message] );
    end;
  end;
  Synchronize( MsgNotify );
  Result := FDataController.LogConnection.Connected;
end;
{ ---------------------------------------------------------------------------- }

procedure TCMCPThread.DisconnectFromCMCPDb;
begin
  if FDataController.CMConnection.Connected then
    FDataController.CMConnection.Connected := False;
  FSoInfo.DbStatus := dbNone;
  Synchronize( DbNotify );     
end;

{ ---------------------------------------------------------------------------- }

procedure TCMCPThread.DisconnectFromLogDb;
begin
  if FDataController.LogConnection.Connected then
   FDataController.LogConnection.Connected := False;
end;

{ ---------------------------------------------------------------------------- }

function TCMCPThread.CheckCanGetData: Boolean;
begin
  Result := ( SecondsBetween( Now, FLastDataRead ) > FConfigInfo.DbScanInterval );
end;

{ ---------------------------------------------------------------------------- }

function TCMCPThread.CheckCanParseXml: Boolean;
begin
  Result := ( SecondsBetween( Now, FLastXmlParse ) > FConfigInfo.XmlParseInterval );
end;

{ ---------------------------------------------------------------------------- }

function TCMCPThread.GetCmdSeqLikeStr: String;
begin
  FDataController.CMDataReader.Close;
  FDataController.CMDataReader.SQL.Text :=
    ' SELECT TO_CHAR( SYSDATE, ''YYYYMMDD'' ) AS YRM FROM DUAL ';
  FDataController.CMDataReader.Open;
  Result := FDataController.CMDataReader.FieldByName( 'YRM' ).AsString;
  FDataController.CMDataReader.Close;
  Result := ( Lpad( FSoInfo.CompCode, 2, '0' ) + Result + '%' );
end;

{ ---------------------------------------------------------------------------- }

function TCMCPThread.GetCmdSeqLikeStrEx(const aPreMonth: Boolean): String;
var
  aDec: Integer;
begin
  aDec := 0;
  if ( aPreMonth ) then aDec := -1;
  FDataController.CMDataReader.Close;
  FDataController.CMDataReader.SQL.Text := Format(
    ' SELECT TO_CHAR( ADD_MONTHS( SYSDATE, %d ), ''YYYYMM'' ) AS YRM FROM DUAL ',
    [aDec] );
  FDataController.CMDataReader.Open;
  Result := FDataController.CMDataReader.FieldByName( 'YRM' ).AsString;
  FDataController.CMDataReader.Close;
  Result := ( Lpad( FSoInfo.CompCode, 2, '0' ) + Result + '%' );
end;

{ ---------------------------------------------------------------------------- }

function TCMCPThread.GetDataSql(const aCmdType: TCmdType): String;
var
  aCmdSeq: String;
begin
  aCmdSeq := GetCmdSeqLikeStr;
  if ( aCmdType = ctOnline ) then
  begin
    Result := Format(
      '  SELECT A.CMDSEQNO,              ' +
      '         A.MSGCODE,               ' +
      '         A.MSGNAME,               ' +
      '         A.WORKID,                ' +
      '         A.CUSSO,                 ' +
      '         A.CUSID,                 ' +
      '         A.CMMAC,                 ' +
      '         A.EMTAMAC,               ' +
      '         A.EMTAPORT,              ' +
      '         A.CPNO,                  ' +
      '         A.CLASSID,               ' +
      '         A.PCIPNO,                ' +
      '         A.PCIP,                  ' +
      '         A.OPERTYPE,              ' +
      '         A.STARTDATETIME,         ' +
      '         A.ENDDATETIME,           ' +
      '         A.ZONE,                  ' +
      '         A.LIN,                   ' +
      '         A.SECTION,               ' +
      '         A.LANE,                  ' +
      '         A.ALLEY,                 ' +
      '         A.SUBALLEY,              ' +
      '         A.NO1,                   ' +
      '         A.NO2,                   ' +
      '         A.PROCESSINGDATE         ' +
      '   FROM %s.SO311 A                ' +
      '  WHERE A.MSGCODE IN ( %s )       ' +
      '    AND A.CMDSEQNO LIKE ''%s''    ' +
      '    AND A.CUSSO = ''%s''          ' +
      '    AND A.XMLDATA IS NULL         ' +
      '    AND A.PROCESSINGDATE IS NULL  ' +
      '  ORDER BY A.CMDSEQNO             ',
      [FSoInfo.Owner, FMsgCodes, aCmdSeq, FSoInfo.CompCode] );
  end else
  begin
    Result := Format(
      '  SELECT * FROM (               ' +
      '  SELECT A.CMDSEQNO,            ' +
      '         A.MSGCODE,             ' +
      '         A.MSGNAME,             ' +
      '         A.WORKID,              ' +
      '         A.CUSSO,               ' +
      '         A.CUSID,               ' +
      '         A.CMMAC,               ' +
      '         A.EMTAMAC,             ' +
      '         A.EMTAPORT,            ' +
      '         A.CPNO,                ' +
      '         A.CLASSID,             ' +
      '         A.PCIPNO,              ' +
      '         A.PCIP,                ' +
      '         A.OPERTYPE,            ' +
      '         A.STARTDATETIME,       ' +
      '         A.ENDDATETIME,         ' +
      '         A.ZONE,                ' +
      '         A.LIN,                 ' +
      '         A.SECTION,             ' +
      '         A.LANE,                ' +
      '         A.ALLEY,               ' +
      '         A.SUBALLEY,            ' +
      '         A.NO1,                 ' +
      '         A.NO2,                 ' +
      '         A.PROCESSINGDATE       ' +
      '   FROM %s.SO311 A              ' +
      '  WHERE A.MSGCODE IN ( %s )     ' +
      '    AND A.CMDSEQNO LIKE ''%s''  ' +
      '    AND A.CUSSO = ''%s''        ' +
      '    AND A.XMLDATA IS NULL       ' +
      '    AND A.PROCESSINGDATE <= SYSDATE ' +
      '  ORDER BY A.CMDSEQNO  )        ' +
      ' WHERE ROWNUM <= 3              ',
      [FSoInfo.Owner, FMsgCodes, aCmdSeq, FSoInfo.CompCode] );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCMCPThread.GetParseSql: String;
var
  aCmdSeq1, aCmdSeq2: String;
begin
  aCmdSeq1 := GetCmdSeqLikeStrEx( True );
  aCmdSeq2 := GetCmdSeqLikeStrEx( False );
  Result := Format(
    '  SELECT * FROM (                     ' +
    '  SELECT A.CMDSEQNO,                  ' +
    '         A.MSGCODE,                   ' +
    '         A.MSGNAME,                   ' +
    '         A.WORKID,                    ' +
    '         A.CUSSO,                     ' +
    '         A.CUSID,                     ' +
    '         A.CMMAC,                     ' +
    '         A.OPERTYPE,                  ' +
    '         A.XMLDATA                    ' +
    '   FROM %s.SO311 A                    ' +
    '  WHERE A.MSGCODE IN ( %s )           ' +
    '    AND A.CUSSO = ''%s''              ' +
    '    AND A.XMLDATA IS NOT NULL         ' +
    '    AND A.PROCESSINGDATE IS NOT NULL  ' +
    '    AND A.OPERRESULT IS NULL          ' +
    '    AND A.QUERYRESULT IS NULL         ' +
    '    AND ( A.CMDSEQNO LIKE ''%s'' OR   ' +
    '          A.CMDSEQNO LIKE ''%s'' )    ' +
    '  ORDER BY A.CMDSEQNO  )              ' +
    ' WHERE ROWNUM <=  %d                  ',
    [FSoInfo.Owner, FMsgCodes, FSoInfo.CompCode, aCmdSeq1, aCmdSeq2,
     FConfigInfo.XmlParseRecords] );
end;

{ ---------------------------------------------------------------------------- }

procedure TCMCPThread.Execute;
var
  aRecords: Integer;
begin
  CoInitialize( nil );
  try
    ConnectToCMCPDb;
    ConnectToLogDb;
    ReleaseCallList;
    ReleaseParseList;
    FLastDataRead := Now;
    FLastXmlParse := Now;
    while not Self.Terminated do
    begin
      try
        { 送指令 }
        if ( CheckCanGetData ) then
        begin
          aRecords := DataToCallList( ctOnline );
          { 若是沒有線上客服指令, 則看是否有預約的指令 }
          if ( aRecords <= 0 ) then
            DataToCallList( ctBatch );
          CallCMWebService;
          FLastDataRead := Now;
        end;
        Sleep( 300 );
        if Self.Terminated then Break;
        { 解析批次指令傳回來的 xml, 並填回對應 so311 欄位  }
        if ( CheckCanParseXml ) then
        begin
          DataToParseList;
          ParseCMWebService;
          FLastXmlParse := Now;
        end;
        Sleep( 300 );
        if Self.Terminated then Break;
      except
        on E: Exception do
        begin
          FMsg.Kind := MB_ICONERROR;
          FMsg.Msg := Format( '資料庫執行緒發生錯誤, 原因:%s。', [E.Message] );
          Synchronize( MsgNotify );
          Sleep( 10 );
        end;
      end;
    end;
    ReleaseCallList;
    ReleaseParseList;
    DisconnectFromCMCPDb;
    DisconnectFromLogDb;
  finally
    CoUninitialize;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCMCPThread.DataToCallList(const aCmdType: TCmdType): Integer;
var
  aSO311: PSO311;
begin
  Result := 0;
  if not CheckCMMCPDbStatus then Exit;
  FSoInfo.DbStatus := dbActive;
  Synchronize( DbNotify );
  try
    Sleep( 300 );
    FDataController.CMDataReader.Close;
    try
      FDataController.CMDataReader.SQL.Text := GetDataSql( aCmdType );
      FDataController.CMDataReader.Open;
      FDataController.CMDataReader.First;
      while not FDataController.CMDataReader.Eof do
      begin
        New( aSO311 );
        { 批次指令 ? }
        aSO311.IsBatch := True;
        if VarIsNull( FDataController.CMDataReader.FieldByName( 'ProcessingDate' ).Value ) then
          aSO311.IsBatch := False;
        {}
        aSO311.CmdSeqNO := FDataController.CMDataReader.FieldByName( 'CMDSEQNO' ).AsString;
        aSO311.MsgCode := FDataController.CMDataReader.FieldByName( 'MSGCODE' ).AsString;
        aSO311.MsgName := FDataController.CMDataReader.FieldByName( 'MSGNAME' ).AsString;
        aSO311.WorkId := FDataController.CMDataReader.FieldByName( 'WORKID' ).AsString;
        aSO311.CusSo := FDataController.CMDataReader.FieldByName( 'CUSSO' ).AsString;
        aSO311.CusId := FDataController.CMDataReader.FieldByName( 'CUSID' ).AsString;
        aSO311.CMMac := FDataController.CMDataReader.FieldByName( 'CMMAC' ).AsString;
        aSO311.EMTAMac := FDataController.CMDataReader.FieldByName( 'EMTAMAC' ).AsString;
        aSO311.EMTAPort := FDataController.CMDataReader.FieldByName( 'EMTAPORT' ).AsString;
        aSO311.CPNo := FDataController.CMDataReader.FieldByName( 'CPNO' ).AsString;
        aSO311.ClassId := FDataController.CMDataReader.FieldByName( 'CLASSID' ).AsString;
        aSO311.PCIPNo := FDataController.CMDataReader.FieldByName( 'PCIPNO' ).AsString;
        aSO311.PCIP := FDataController.CMDataReader.FieldByName( 'PCIP' ).AsString;
        aSO311.OperType := FDataController.CMDataReader.FieldByName( 'OPERTYPE' ).AsString;
        {}
        aSO311.StartDateTime := EmptyStr;
        aSO311.EndDateTime := EmptyStr;
        {}
        if not VarIsNull( FDataController.CMDataReader.FieldByName( 'STARTDATETIME' ).Value ) then
          aSO311.StartDateTime := FormatDateTime( 'yyyy-mm-dd hh:nn:ss',
            FDataController.CMDataReader.FieldByName( 'STARTDATETIME' ).AsDateTime );
        {}
        if not VarIsNull( FDataController.CMDataReader.FieldByName( 'ENDDATETIME' ).Value ) then
          aSO311.EndDateTime := FormatDateTime( 'yyyy-mm-dd hh:nn:ss',
            FDataController.CMDataReader.FieldByName( 'ENDDATETIME' ).AsDateTime );
        {}
        aSO311.Zone := FDataController.CMDataReader.FieldByName( 'ZONE' ).AsString;
        aSO311.Lin := FDataController.CMDataReader.FieldByName( 'LIN' ).AsString;
        aSO311.Section := FDataController.CMDataReader.FieldByName( 'SECTION' ).AsString;
        aSO311.Lane := FDataController.CMDataReader.FieldByName( 'LANE' ).AsString;
        aSO311.Alley := FDataController.CMDataReader.FieldByName( 'ALLEY' ).AsString;
        aSO311.SubAlley := FDataController.CMDataReader.FieldByName( 'SUBALLEY' ).AsString;
        aSO311.NO1 := FDataController.CMDataReader.FieldByName( 'NO1' ).AsString;
        aSO311.NO2 := FDataController.CMDataReader.FieldByName( 'NO2' ).AsString;
        {}
        aSO311.SOAPRequest := EmptyStr;
        aSO311.SOAPResponse := EmptyStr;
        {}
        aSO311.DisplayNode := nil;
        aSO311.ErrMsg := EmptyStr;
        aSO311.ExtendedInfo := EmptyStr;
        {}
        FActiveSO311 := aSO311;
        FCallList.Add( aSO311 );
        Synchronize( RecordNotify );
        FDataController.CMDataReader.Next;
        Inc( Result );
      end;
      FDataController.CMDataReader.Close;
      FActiveSO311 := nil;
      if ( Result > 0 ) then
      begin
        FMsg.Kind := MB_ICONINFORMATION;
        FMsg.Msg := Format( '系統台: %s 待傳送指令共 %d 筆。', [FSoInfo.CompName, Result] );
        Synchronize( MsgNotify );
      end;
    except
      on E: Exception  do
      begin
        FMsg.Kind := MB_ICONERROR;
        FMsg.Msg := Format( '系統台: %s 抓取待傳送指令失敗, 原因: %s。',
          [FSoInfo.CompName, E.Message] );
        Synchronize( MsgNotify );
      end;
    end;
  finally
    FSoInfo.DbStatus := dbOK;
    Synchronize( DbNotify );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TCMCPThread.ReleaseCallList;
var
  aIndex: Integer;
begin
  for aIndex := 0 to FCallList.Count - 1 do
  begin
    if Assigned( FCallList[aIndex] ) then
    begin
      Dispose( PSO311( FCallList[aIndex] ) );
      FCallList[aIndex] := nil;
    end;
  end;
  FCallList.Clear;
end;

{ ---------------------------------------------------------------------------- }

procedure TCMCPThread.ReleaseParseList;
var
  aIndex: Integer;
begin
  for aIndex := 0 to FParseList.Count - 1 do
  begin
    if Assigned( FParseList[aIndex] ) then
    begin
      Dispose( PSO311( FParseList[aIndex] ) );
      FParseList[aIndex] := nil;
    end;
  end;
  FParseList.Clear;
end;

{ ---------------------------------------------------------------------------- }

procedure TCMCPThread.CallCMWebService;
var
  aIndex: Integer;
  aSO311: PSO311;
  aUrl, aActionText, aResult: String;
begin
  for aIndex := 0 to FCallList.Count - 1 do
  begin
    aSO311 := PSO311( FCallList[aIndex] );
    try
      FActiveSO311 := aSO311;
      aSO311.ErrMsg := EmptyStr;
      if aSO311.SOAPRequest <> EmptyStr then Continue;
      {}
      FRequestList.Text := FWebServiceCall.GetSOAPRequestText( aSO311 );
      aActionText := FWebServiceCall.GetSOAPActionEndText( aSO311 );
      {}
      aUrl := FConfigInfo.SOAPUrl;
      if not IsDelimiter( '/', aUrl, Length( aUrl )  ) then
        aUrl := ( aUrl + '/' );
      aUrl := ( aUrl + aActionText );
      if ( aSO311.ExtendedInfo <> EmptyStr ) then
        aSO311.SOAPRequest := aSO311.ExtendedInfo
      else
        aSO311.SOAPRequest := FRequestList.Text;
      {}
      Synchronize( RecordNotify );
      Sleep( 100 );
      {}
      aResult := FSOAPHttp.Post( aUrl, FRequestList );
      aSO311.SOAPResponse := Utf8ToAnsi( aResult );
      Synchronize( RecordNotify );
      Sleep( 100 );
    except
      on E: Exception do
      begin
        FMsg.Kind := MB_ICONERROR;
        FMsg.Msg := Format( '系統台: %s, 呼叫WEB服務( %s )失敗, 序號: %s, 原因: %s。',
          [FSoInfo.CompName, aActionText, aSO311.CmdSeqNO, E.Message] );
        Synchronize( MsgNotify );
        aSO311.ErrMsg := E.Message;
        Synchronize( RecordNotify );
        Sleep( 200 );
      end;
    end;
    SaveXml( aSO311 );
    SaveToLog( aSO311 );
    Dispose( PSO311( FCallList[aIndex] ) );
    FCallList[aIndex] := nil;
  end;
  while FCallList.Count > 0 do
    FCallList.Delete( 0 );
end;

{ ---------------------------------------------------------------------------- }

procedure TCMCPThread.SaveXml(aSO311: PSO311);
begin
  if ( aSO311.ErrMsg <> EmptyStr ) then
    aSO311.SOAPResponse := FWebServiceCall.GetErrorXml( aSO311 );
  {}  
  FDataController.CMDataWriter.Close;
  FDataController.CMDataWriter.SQL.Clear;
  FDataController.CMDataWriter.Parameters.Clear;
  FDataController.CMDataWriter.SQL.Text := Format(
    ' UPDATE %s.SO311                    ' +
    '    SET XMLDATA = :DATA             ' +
    '  WHERE CMDSEQNO = :CMDSEQ          ', [FSoInfo.Owner] );
  {}
  FDataController.CMDataWriter.Parameters.ParamByName( 'DATA' ).Value :=
    aSO311.SOAPResponse;
  FDataController.CMDataWriter.Parameters.ParamByName( 'CMDSEQ' ).Value :=
    aSO311.CmdSeqNO;
  try
    FDataController.CMDataWriter.ExecSQL;
    FMsg.Kind := MB_OK;
    FMsg.Msg := Format( '系統台: %s, 呼叫 %s 完成, 指令序號: %s。',
      [FSoInfo.CompName, aSO311.MsgName, aSO311.CmdSeqNO] );
  except
    on E: Exception do
    begin
      FMsg.Kind := MB_ICONERROR;
      FMsg.Msg := Format( '系統台: %s, 呼叫: %s 失敗, 指令序號: %s, 原因: %s。',
        [FSoInfo.CompName, aSO311.MsgName, aSO311.CmdSeqNO, E.Message] );
    end;
  end;
  Synchronize( MsgNotify );
end;

{ ---------------------------------------------------------------------------- }

procedure TCMCPThread.SaveToLog(aSO311: PSO311);
begin
  {....}
end;

{ ---------------------------------------------------------------------------- }

function TCMCPThread.ParseXml(aSo311: PSO311): Boolean;
begin
  try
    { 因為已先將回傳的 xml 轉成 big-5, 但是 xml 抬頭的 encoding 沒改,
      所以在這邊改掉 }
    aSo311.SOAPResponse := StringReplace( aSo311.SOAPResponse,
      'encoding="utf-8"?', 'encoding="Big5"?', [] );
    FWebServiceCall.ParseXml( aSo311 );
    FMsg.Kind := MB_OK;
    FMsg.Msg := Format( '系統台: %s, 批次指令: %s 拆解XML完成, 指令序號: %s。',
      [FSoInfo.CompName, aSo311.MsgName, aSo311.CmdSeqNO] );
  except
    on E: Exception do
    begin
      FMsg.Kind := MB_ICONERROR;
      FMsg.Msg := Format( '系統台: %s, 批次指令: %s 拆解XML失敗, 指令序號: %s, 原因: %s。',
        [FSoInfo.CompName, aSo311.MsgName, aSo311.CmdSeqNO, E.Message] );
    end;
  end;
  Synchronize( MsgNotify );
  Result := ( FMsg.Kind = MB_OK );
end;

{ ---------------------------------------------------------------------------- }

procedure TCMCPThread.UpdateBack(aSo311: PSO311);
begin
  FDataController.CMDataWriter.Close;
  FDataController.CMDataWriter.SQL.Clear;
  FDataController.CMDataWriter.Parameters.Clear;
  {}
  try
    FDataController.CMDataWriter.SQL.Text :=
      FWebServiceCall.GetUpdateXmlFieldSql( aSo311 );
    FDataController.CMDataWriter.ExecSQL;
    FMsg.Kind := MB_OK;
    FMsg.Msg := Format( '系統台: %s, 批次指令: %s 拆解xml回填完成, 指令序號: %s。',
      [FSoInfo.CompName, aSo311.MsgName, aSo311.CmdSeqNO] );
  except
    on E: Exception do
    begin
      FMsg.Kind := MB_ICONERROR;
      FMsg.Msg := Format( '系統台: %s, 批次指令: %s 拆解xml回填失敗, 指令序號: %s, 原因: %s。',
        [FSoInfo.CompName, aSo311.MsgName, aSo311.CmdSeqNO, E.Message] );
    end;
  end;
  Synchronize( MsgNotify );
end;

{ ---------------------------------------------------------------------------- }

function TCMCPThread.UpdateOtherData(aSo311: PSO311): Boolean;
var
  aMsgCode, aOperType: Integer;
begin
  Result := True;
  { 只有成功的指令才可以 Update 其它資料 Table }
  if ( aSo311.OperResult = '1' ) or ( aSo311.QueryResult = '1' ) then
  begin
    aMsgCode := StrToIntDef( aSo311.MsgCode, -1 );
    aOperType := StrToIntDef( aSo311.OperType, -1 );
    case aMsgCode of
      2:
        begin
          { 開機}
          if ( aOperType = 1 ) then
            Result := UpdateCMReg1( aSo311 )
          { 關機 }
          else
            Result := UpdateCMReg2( aSo311 );
        end;
      27:
        { CP Ip 申請 }
        begin
          Result := UpdateCPIPReg( aSo311 );
        end;
    else
      {}
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TCMCPThread.IsSpecialCmd(aSo311: PSO311): Boolean;
var
  aMsgCode, aOperType: Integer;
begin
  Result := False;
  aMsgCode := StrToIntDef( aSo311.MsgCode, -1 );
  aOperType := StrToIntDef( aSo311.OperType, -1 );
  // 開關機指令, 要特別去 update so004 的資料
  // CP 申裝, 也要 update so004
  if ( ( aMsgCode = 2 ) and ( aOperType in [1..2] ) ) or
     ( aMsgCode = 27 ) then
    Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TCMCPThread.DataToParseList: Integer;
var
  aSo311: PSO311;
begin
  Result := 0;
  if not CheckCMMCPDbStatus then Exit;
  FSoInfo.DbStatus := dbActive;
  Synchronize( DbNotify );
  try
    Sleep( 300 );
    FDataController.CMDataReader.Close;
    try
      FDataController.CMDataReader.SQL.Text := GetParseSql;
      FDataController.CMDataReader.Open;
      FDataController.CMDataReader.First;
      while not FDataController.CMDataReader.Eof do
      begin
        New( aSo311 );
        { 抓批次指令來拆解xml資料, 一定是批次指令  }
        aSo311.IsBatch := True;
        {}
        aSO311.CmdSeqNO := FDataController.CMDataReader.FieldByName( 'CMDSEQNO' ).AsString;
        aSO311.MsgCode := FDataController.CMDataReader.FieldByName( 'MSGCODE' ).AsString;
        aSO311.MsgName := FDataController.CMDataReader.FieldByName( 'MSGNAME' ).AsString;
        aSO311.WorkId := FDataController.CMDataReader.FieldByName( 'WORKID' ).AsString;
        aSO311.CusSo := FDataController.CMDataReader.FieldByName( 'CUSSO' ).AsString;
        aSO311.CusId := FDataController.CMDataReader.FieldByName( 'CUSID' ).AsString;
        aSO311.CMMac := FDataController.CMDataReader.FieldByName( 'CMMAC' ).AsString;
        aSO311.OperType := FDataController.CMDataReader.FieldByName( 'OPERTYPE' ).AsString;
        {}
        aSO311.SOAPResponse := FDataController.CMDataReader.FieldByName( 'XMLDATA' ).AsString; 
        {}
        aSO311.DisplayNode := nil;
        aSO311.ErrMsg := EmptyStr;
        aSO311.ExtendedInfo := EmptyStr;
        {}
        FParseList.Add( aSo311 );
        FDataController.CMDataReader.Next;
        Inc( Result );
      end;
      FDataController.CMDataReader.Close;
      if ( Result > 0 ) then
      begin
        FMsg.Kind := MB_ICONINFORMATION;
        FMsg.Msg := Format( '系統台: %s 待拆解xml批次指令共 %d 筆。',
          [FSoInfo.CompName, Result] );
        Synchronize( MsgNotify );
      end;
    except
      on E: Exception  do
      begin
        FMsg.Kind := MB_ICONERROR;
        FMsg.Msg := Format( '系統台: %s 抓取待拆解xml批次指令失敗, 原因: %s。',
          [FSoInfo.CompName, E.Message] );
        Synchronize( MsgNotify );
      end;
    end;
  finally
    FSoInfo.DbStatus := dbOK;
    Synchronize( DbNotify );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TCMCPThread.ParseCMWebService;
var
  aIndex: Integer;
  aSo311: PSO311;
  aCanUpdateXmlField: Boolean;
begin
  { 拆解批次指令 xml, 回填各欄位 }
  { 若 db 不通則不做 }
  if not CheckCMMCPDbStatus then
  begin
    ReleaseParseList;
    Exit;
  end;
  for aIndex := 0 to FParseList.Count - 1 do
  begin
    aSo311 := PSO311( FParseList[aIndex] );
    if ParseXml( aSo311 ) then
    begin
      { 如果是特殊指令須要 update 其它 table , 則必須先處理 }
      { 成功後才可以 update so0311 各拆解的 xml 欄位 }
      { 因為若 update 其它 table 不成功, 則下次撈取資料還可以抓出來再做一次 }
      aCanUpdateXmlField := True;
      if IsSpecialCmd( aSo311 ) then
        aCanUpdateXmlField := UpdateOtherData( aSo311 );
      if ( aCanUpdateXmlField ) then
        UpdateBack( aSo311 );
    end;
    Dispose( PSO311( FParseList[aIndex] ) );
    FParseList[aIndex] := nil;
  end;
  while FParseList.Count > 0 do
    FParseList.Delete( 0 );
end;

{ ---------------------------------------------------------------------------- }

function TCMCPThread.UpdateCMReg1(aSo311: PSO311): Boolean;
var
  aRecordEffect: Integer;
begin
  // 開機
  aRecordEffect := 0;
  FDataController.CMDataWriter.Close;
  FDataController.CMDataWriter.SQL.Clear;
  try
    FDataController.CMDataWriter.SQL.Text := Format(
      ' update %s.so004 set                        ' +
      '    cmopendate = sysdate,                   ' +
      '    ipaddress = nvl( ipaddress, ''%s'' ),   ' +
      '    upden = ''CABLESOFT'',                  ' +
      '    updtime = to_char( to_number( to_char( sysdate, ''YYYY'' ) ) - 1911 ) || ' +
      '       ''/'' || to_char( sysdate, ''mm/dd hh24:mi:ss'' ) ' +
      '  where custid = ''%s''                    ' +
      '    and facisno = ''%s''                   ',
      [FSoInfo.Owner, aSo311.RCTIp, aSo311.CusId, aSo311.CMMac] );
    aRecordEffect := FDataController.CMDataWriter.ExecSQL;
  except
    on E: Exception do
    begin
     FMsg.Kind := MB_ICONERROR;
     FMsg.Msg := Format( '系統台: %s, 批次指令: %s 回填SO004失敗, 原因: %s。',
       [FSoInfo.CompName, aSo311.MsgName, E.Message] );
     Synchronize( MsgNotify );
    end;
  end;
  Result := ( aRecordEffect > 0 );
end;

{ ---------------------------------------------------------------------------- }

function TCMCPThread.UpdateCMReg2(aSo311: PSO311): Boolean;
var
  aRecordEffect: Integer;
begin
  { 關機 }
  aRecordEffect := 0;
  FDataController.CMDataWriter.Close;
  FDataController.CMDataWriter.SQL.Clear;
  try
    FDataController.CMDataWriter.SQL.Text := Format(
      ' update %s.so004 set                         ' +
      '    cmclosedate = sysdate,                  ' +
      '    ipaddress = nvl( ipaddress, ''%s'' ),   ' +
      '    upden = ''CABLESOFT'',                  ' +
      '    updtime = to_char( to_number( to_char( sysdate, ''YYYY'' ) ) - 1911 ) || ' +
      '       ''/'' || to_char( sysdate, ''mm/dd hh24:mi:ss'' ) ' +
      '  where custid = ''%s''                    ' +
      '    and facisno = ''%s''                   ',
      [FSoInfo.Owner, aSo311.RCTIp, aSo311.CusId, aSo311.CMMac] );
    aRecordEffect := FDataController.CMDataWriter.ExecSQL;
  except
    on E: Exception do
    begin
     FMsg.Kind := MB_ICONERROR;
     FMsg.Msg := Format( '系統台: %s, 批次指令: %s 回填SO004失敗, 原因: %s。',
       [FSoInfo.CompName, aSo311.MsgName, E.Message] );
     Synchronize( MsgNotify );
    end;
  end;
  Result := ( aRecordEffect > 0 );
end;

{ ---------------------------------------------------------------------------- }

function TCMCPThread.UpdateCPIPReg(aSo311: PSO311): Boolean;
var
  aRecordEffect: Integer;
begin
  { CP IP 申請 }
  aRecordEffect := 0;
  FDataController.CMDataWriter.Close;
  FDataController.CMDataWriter.SQL.Clear;
  try
    FDataController.CMDataWriter.SQL.Text := Format(
      ' update %s.so004 set                         ' +
      '    ipaddress = nvl( ipaddress, ''%s'' ),   ' +
      '    upden = ''CABLESOFT'',                  ' +
      '    updtime = to_char( to_number( to_char( sysdate, ''YYYY'' ) ) - 1911 ) || ' +
      '       ''/'' || to_char( sysdate, ''mm/dd hh24:mi:ss'' ) ' +
      '  where custid = ''%s''                    ' +
      '    and facisno = ''%s''                   ',
      [FSoInfo.Owner, aSo311.RCTIp, aSo311.CusId, aSo311.CMMac] );
    aRecordEffect := FDataController.CMDataWriter.ExecSQL;
  except
    on E: Exception do
    begin
     FMsg.Kind := MB_ICONERROR;
     FMsg.Msg := Format( '系統台: %s, 批次指令: %s 回填SO004失敗, 原因: %s。',
       [FSoInfo.CompName, aSo311.MsgName, E.Message] );
     Synchronize( MsgNotify );
    end;
  end;
  Result := ( aRecordEffect > 0 );
end;

{ ---------------------------------------------------------------------------- }

end.



