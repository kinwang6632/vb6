unit cbSoDbThread;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, cbClass, cbAppClass,
  DB, ADODB, CodeSiteLogging;

type
  TSoDbThread = class(TMessageQueueThread)
  private
    { Private declarations }
    FSoList: TAppSoList;
    FProcList: TRecycleCommandList;
    FSendNagraList: TRecycleCommandList;
    FMsgSubject: TMessageSubject;
    FCmdSubject: TObjectSubject;
    FDbSubject: TObjectSubject;
    FCmdHelper: TCommandHelper;
    FLastQueryRecordTime: TDateTime;
    FLastQuerySendNagraTime: TDateTime;
    FLastWriteSendNagraTime: TDateTime;
    FLastQueryCallbackDataTime: TDateTime;
    FLastQueryConfirmTime: TDateTime;
    //function IsDeadConnection(aSo: TAppSo): Boolean;
    function GetAppSo(const aCompCode: String): TAppSo;
    function CanRetryDbConnection(aLastCheckTime: TDateTime): Boolean;
    function CanRetryComDbConnection(aLastCheckTime: TDateTime): Boolean;
    function CanGetProcessRecord: Boolean;
    function CanWriteSendNagra: Boolean;
    function CanGetSendNagra: Boolean;
    function CanGetCallbackData: Boolean;
    function CanGetConfirmData: Boolean;
    function GetWriteSendNagraSql(const aCmd: TRecycleCommand; const aCmdCode: String): String;
    function GetUpdateCARecycleSql_1(const aCmd: TRecycleCommand): String;
    function GetUpdateCARecycleSql_2(const aCmd: TRecycleCommand; const aCmdCode: String): String;
    function GetCallbackDataSql(const aCmd: TRecycleCommand): String;
    function GetConfirmDataSql(const aCmd: TRecycleCommand): String;
    procedure AssignSoList(aSource: TAppSoList);
    procedure ReleaseSoList;
    procedure PrepareSoConnection(aSo: TAppSo); overload;
    procedure PrepareSoConnection; overload;
    procedure PrePareCAConnection(aSo: TAppSo); overload;
    procedure PrePareCAConnection; overload;
    procedure UnPrepareSoConnection(aSo: TAppSo); overload;
    procedure UnPrepareSoConnection; overload;
    procedure UnPrepareCAConnection(aSo: TAppSo); overload;
    procedure UnPrepareCAConection; overload;
    procedure GetProcessRecord(aSo: TAppSo; var aDbStatus: TDbConnectStatus); overload;
    procedure GetProcessRecord; overload;
    procedure GetSendNagra; overload;
    procedure WriteToSendNagra;
    procedure GetCallbackData;
    procedure GetConfirmData;
  protected
    procedure Execute; override;
  public
    constructor Create(CreateSuspended: Boolean; aList: TAppSoList); reintroduce;
    destructor Destroy; override;
    property MsgSubject: TMessageSubject read FMsgSubject;
    property CmdSubject: TObjectSubject read FCmdSubject;
    property DbSubject: TObjectSubject read FDbSubject;
  end;

implementation

{ Important: Methods and properties of objects in visual components can only be
  used in a method called using Synchronize, for example,

      Synchronize(UpdateCaption);

  and UpdateCaption could look like,

    procedure TSoDbThread.UpdateCaption;
    begin
      Form1.Caption := 'Updated in a thread';
    end; }

uses cbUtilis, cbSoDataModule, DateUtils;

{ TSoDbThread }

{ ---------------------------------------------------------------------------- }

function CheckPeriod(aPeriod: TDateTime; aSec: Cardinal): Boolean;
begin
  Result := ( aPeriod <= 0 );
  if not Result then
    Result := ( SecondsBetween( Now, aPeriod ) >= aSec );
end;

{ ---------------------------------------------------------------------------- }

constructor TSoDbThread.Create(CreateSuspended: Boolean; aList: TAppSoList);
begin
  inherited Create( CreateSuspended );
  FSoList := TAppSoList.Create;
  FProcList := TRecycleCommandList.Create;
  FSendNagraList := TRecycleCommandList.Create;
  ReleaseSoList;
  AssignSoList( aList );
  FMsgSubject := TMessageSubject.Create;
  FCmdSubject := TObjectSubject.Create;
  FDbSubject := TObjectSubject.Create;
  FCmdHelper := TCommandHelper.Create;
end;

{ ---------------------------------------------------------------------------- }

destructor TSoDbThread.Destroy;
begin
  FCmdHelper.Free;
  FDbSubject.Free;
  FCmdSubject.Free;
  FMsgSubject.Free;
  FSoList.Free;
  FProcList.Free;
  FSendNagraList.Free;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.AssignSoList(aSource: TAppSoList);
var
  aIndex: Integer;
  aSo: TAppSo;
begin
  for aIndex := 0 to aSource.Count - 1 do
  begin
    if ( aSource[aIndex].Selected ) then
    begin
      aSo := TAppSo.Create;
      aSo.Assign( aSource[aIndex] );
      FSoList.Add( aSo );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.ReleaseSoList;
var
  aIndex: Integer;
begin
  for aIndex := 0 to FSoList.Count - 1 do
  begin
    FSoList[aIndex].DataModule.Free;
    FSoList[aIndex].ComDataModule.Free;
  end;  
  FSoList.Clear;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.Execute;
begin
  FLastQueryRecordTime := 0;
  FLastQuerySendNagraTime := 0;
  FLastWriteSendNagraTime := 0;
  FLastQueryCallbackDataTime := 0;
  FLastQueryConfirmTime := 0;
  {}
  FMsgSubject.Normal( '系統台資料庫準備連線中...。' );
  Sleep( 1000 );
  PrepareSoConnection;
  FMsgSubject.Normal( 'COM區資料庫準備連線中...。' );
  Sleep( 1000 );
  PrePareCAConnection;
  {}
  WaitForPlaySignal;
  Sleep( 300 );
  while not ( Terminated or Stop ) do
  begin
    try
      if ( Terminated or Stop ) then Break;
      Sleep( 100 );
      if CanGetProcessRecord then GetProcessRecord;
      if ( Terminated or Stop ) then Break;
      Sleep( 100 );
      if CanWriteSendNagra then WriteToSendNagra;
      if ( Terminated or Stop ) then Break;
      Sleep( 100 );
      if CanGetSendNagra then GetSendNagra;
      if ( Terminated or Stop ) then Break;
      Sleep( 100 );
      if CanGetSendNagra then GetSendNagra;
      if CanGetCallbackData then GetCallbackData;
      Sleep( 100 );
      if CanGetSendNagra then GetSendNagra;
      if CanGetConfirmData then GetConfirmData;
      Sleep( 100 );
    except
      on E: Exception do
      begin
        FMsgSubject.Error( Format( '系統台資料庫執行緒發生錯誤, 訊息:%s',
          [E.Message] ) );
      end;
    end;
    Sleep( 100 );
  end;
  WaitForStopSignal;
  FMsgSubject.Normal( '系統台資料庫離線中...。' );
  Sleep( 500 );
  CodeSite.Send( 'after sleep 1000' );
  UnPrepareSoConnection;
  FMsgSubject.Normal( 'COM區資料庫離線中...。' );
  Sleep( 500 );
  UnPrepareCAConection;
  ReleaseSoList;
end;

{ ---------------------------------------------------------------------------- }
(*
function TSoDbThread.IsDeadConnection(aSo: TAppSo): Boolean;
begin
  TSoDataModule( aSo.DataModule ).SoDataReader.Close;
  TSoDataModule( aSo.DataModule ).SoDataReader.SQL.Text := ' select 1 from dual ';
  try
    TSoDataModule( aSo.DataModule ).SoDataReader.Open;
  except
    { ... }
  end;
  Result := not TSoDataModule( aSo.DataModule ).SoDataReader.Active;
  TSoDataModule( aSo.DataModule ).SoDataReader.Close;
end;
*)
{ ---------------------------------------------------------------------------- }

function TSoDbThread.CanRetryDbConnection(aLastCheckTime: TDateTime): Boolean;
begin
  Result := CheckPeriod( aLastCheckTime, 60 );
end;

{ ---------------------------------------------------------------------------- }

function TSoDbThread.CanRetryComDbConnection(aLastCheckTime: TDateTime): Boolean;
begin
  Result := CheckPeriod( aLastCheckTime, 60 );
end;

{ ---------------------------------------------------------------------------- }

function TSoDbThread.CanGetProcessRecord: Boolean;
begin
  Result := CheckPeriod( FLastQueryRecordTime, 30 );
  if Result then FLastQueryRecordTime := Now;
end;

{ ---------------------------------------------------------------------------- }

function TSoDbThread.CanGetSendNagra: Boolean;
begin
  Result := CheckPeriod( FLastQuerySendNagraTime, 10 );
  if Result then FLastQuerySendNagraTime := Now;
end;

{ ---------------------------------------------------------------------------- }

function TSoDbThread.CanWriteSendNagra: Boolean;
begin
  Result := CheckPeriod( FLastWriteSendNagraTime, 10 );
  if Result then FLastWriteSendNagraTime := Now;
end;

{ ---------------------------------------------------------------------------- }

function TSoDbThread.CanGetCallbackData: Boolean;
begin
  Result := CheckPeriod( FLastQueryCallbackDataTime, 30 );
  if Result then FLastQueryCallbackDataTime := Now;
end;

{ ---------------------------------------------------------------------------- }

function TSoDbThread.CanGetConfirmData: Boolean;
begin
  Result := CheckPeriod( FLastQueryConfirmTime, 10 );
  if Result then FLastQueryConfirmTime := Now;
end;

{ ---------------------------------------------------------------------------- }

function TSoDbThread.GetWriteSendNagraSql(const aCmd: TRecycleCommand;
  const aCmdCode: String): String;
var
  aProcDateTime: String;
  aDate1: TDateTime;
begin
  Result := EmptyStr;
  { Paire }
  if ( aCmdCode = '52_1' ) then
  begin
    Result := Format(
      ' insert into send_nagra ( high_level_cmd_id, icc_no, stb_no,    ' +
      '   cmd_status, operator, updtime, seqno, compcode, resenttimes, ' +
      '   processingdate, stb_flag )                                   ' +
      ' values ( ''A7'', ''%s'', ''%s'', ''W'', ''%s'', sysdate,       ' +
      '   :seqno, ''%d'', 0, sysdate, ''%d'' )                         ',
      [aCmd.IccNo, aCmd.StbNo, aCmd.Operator, aCmd.CompCode, aCmd.StbFlag] );
  end else
  { Set Credit = 0 }
  if ( aCmdCode = '8' ) then
  begin
    aDate1 := Now;
    if ( aCmd.LastDownloadTime <> EmptyStr ) then
    begin
      aDate1 := StrToDateTime( aCmd.LastDownloadTime );
      if ( HoursBetween( Now, aDate1 ) < 25 ) then
        aDate1 := IncHour( aDate1, 25 );
    end;
    aProcDateTime := FormatDateTime( 'yyyy/mm/dd hh:nn:ss', aDate1 );
    Result := Format(
      ' insert into send_nagra ( high_level_cmd_id, icc_no, cmd_status, ' +
      '   operator, updtime, seqno, compcode, resenttimes,              ' +
      '   processingdate, stb_flag, credit_mode, credit )               ' +
      ' values ( ''P21'', ''%s'', ''W'', ''%s'', sysdate, :seqno,       ' +
      '   ''%d'', 0, to_date( ''%s'', ''YYYY/MM/DD HH24:MI:SS'' ),      ' +
      '   ''%d'', ''04'', ''0'' )                                       ',
      [aCmd.IccNo, aCmd.Operator, aCmd.CompCode, aProcDateTime, aCmd.StbFlag] );
    //aCmd.LastDownloadTime := aProcDateTime;  
  end else
  { Set Ippv Record as Reported and Purge }
  if ( aCmdCode = '97_96' ) then
  begin
    Result := Format(
      ' insert into send_nagra ( high_level_cmd_id, icc_no, cmd_status, ' +
      '   operator, updtime, seqno, compcode, resenttimes,              ' +
      '   processingdate, stb_flag, cleanup_date, condition_date,       ' +
      '   collect_date )                                                ' +
      ' values ( ''P4'', ''%s'', ''W'', ''%s'', sysdate, :seqno,        ' +
      '   ''%d'', 0, sysdate, ''%d'', sysdate, sysdate, sysdate )       ',
      [aCmd.IccNo, aCmd.Operator, aCmd.CompCode, aCmd.StbFlag] );
  end else
  { Cancel All Product }
  if ( aCmdCode = '7' ) then
  begin
    Result := Format(
      ' insert into send_nagra ( high_level_cmd_id, icc_no, cmd_status, ' +
      '   operator, updtime, seqno, compcode, resenttimes,              ' +
      '   processingdate, stb_flag )                                    ' +
      ' values ( ''B3'', ''%s'', ''W'', ''%s'', sysdate, :seqno,        ' +
      '   ''%d'', 0, sysdate, ''%d'' )                                  ',
      [aCmd.IccNo, aCmd.Operator, aCmd.CompCode, aCmd.StbFlag] );
  end else
  { Set Zip Code --> Clear Zip Code, always fill 0 }
  if ( aCmdCode = '48' ) then
  begin
    Result := Format(
      ' insert into send_nagra ( high_level_cmd_id, icc_no, cmd_status, ' +
      '   operator, updtime, seqno, compcode, resenttimes,              ' +
      '   processingdate, stb_flag, zip_code )                          ' +
      ' values ( ''C4'', ''%s'', ''W'', ''%s'', sysdate, :seqno,        ' +
      '   ''%d'', 0, sysdate, ''%d'', ''0'' )                           ',
      [aCmd.IccNo, aCmd.Operator, aCmd.CompCode, aCmd.StbFlag] );
  end else
  { UnPaire }
  if ( aCmdCode = '52_2' ) then
  begin
    Result := Format(
      ' insert into send_nagra ( high_level_cmd_id, notes, cmd_status,  ' +
      '   operator, updtime, seqno, compcode, resenttimes,              ' +
      '   processingdate, stb_flag  )                                   ' +
      ' values ( ''C1'', ''%s'', ''W'', ''%s'', sysdate, :seqno,        ' +
      '   ''%d'', 0, sysdate, ''%d'' )                                  ',
      [aCmd.IccNo, aCmd.Operator, aCmd.CompCode, aCmd.StbFlag] );
  end else
  { Disable Automatic callback }
  if ( aCmdCode = '62' ) then
  begin
    Result := Format(
      ' insert into send_nagra ( high_level_cmd_id, icc_no, cmd_status,  ' +
      '   operator, updtime, seqno, compcode, resenttimes,              ' +
      '   processingdate, stb_flag  )                                   ' +
      ' values ( ''P13'', ''%s'', ''W'', ''%s'', sysdate, :seqno,       ' +
      '   ''%d'', 0, sysdate, ''%d'' )                                  ',
      [aCmd.IccNo, aCmd.Operator, aCmd.CompCode, aCmd.StbFlag] );
  end else
  { immediate callback }
  if ( aCmdCode = '60' ) then
  begin
    Result := Format(
      ' insert into send_nagra ( high_level_cmd_id, icc_no, cmd_status,  ' +
      '   operator, updtime, seqno, compcode, resenttimes,              ' +
      '   processingdate, stb_flag, callback_date, callback_time  )     ' +
      ' values ( ''P14'', ''%s'', ''W'', ''%s'', sysdate, :seqno,       ' +
      '   ''%d'', 0, sysdate, ''%d'', to_char( sysdate, ''YYYYMMDD'' ), ' +
      '   to_char( sysdate, ''HH24MISS'' ) )                            ',
      [aCmd.IccNo, aCmd.Operator, aCmd.CompCode, aCmd.StbFlag] );
  end else
end;

{ ---------------------------------------------------------------------------- }

function TSoDbThread.GetUpdateCARecycleSql_1(const aCmd: TRecycleCommand): String;
begin
  Result := Format(
    ' update carecycle set     ' +
    '   updtime = sysdate,     ' +
    '   cmdflag = 1,           ' +
    '   recycletext = ''%s'',  ' +
    '   transnum = ''%s''      ' +
    '  where seqno = ''%s''    ' +
    '    and icc_no = ''%s''   ' +
    '    and compcode = ''%d'' ',
    [aCmd.RecycleText, aCmd.TransNum, aCmd.SeqNo, aCmd.IccNo, aCmd.CompCode] );
  aCmd.UpdTime := Now;  
end;

{ ---------------------------------------------------------------------------- }

function TSoDbThread.GetUpdateCARecycleSql_2(const aCmd: TRecycleCommand;
  const aCmdCode: String): String;
begin
  if ( aCmdCode = '8' ) then
  begin
    aCmd.LastDownloadTime := FormatDateTime( 'yyyy/mm/dd hh:nn:ss', Now );
    Result := Format(
      ' update carecycle                    ' +
      '    set cmdflag = ''%d'',            ' +
      '        lastdownloadtime = to_date( ''%s'', ''YYYY/MM/DD HH24:MI:SS'' ), ' +
      '        recycletext = ''%s'',        ' +
      '        err_msg = ''%s''             ' +
      '  where icc_no = ''%s''              ' +
      '    and seqno = ''%s''               ' +
      '    and compcode = ''%d''            ',
      [aCmd.CMDFlag, aCmd.LastDownloadTime, aCmd.RecycleText, aCmd.ErrMsg,
       aCmd.IccNo, aCmd.SeqNo, aCmd.CompCode] );
  end else
  begin
    Result := Format(
      ' update carecycle              ' +
      '    set cmdflag = ''%d'',      ' +
      '        recycletext = ''%s'',  ' +
      '        err_msg = ''%s''       ' +
      '  where icc_no = ''%s''        ' +
      '    and seqno = ''%s''         ' +
      '    and compcode = ''%d''      ',
      [aCmd.CMDFlag, aCmd.RecycleText, aCmd.ErrMsg, aCmd.IccNo, aCmd.SeqNo,
       aCmd.CompCode] );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TSoDbThread.GetCallbackDataSql(const aCmd: TRecycleCommand): String;
var
  aDate: String;
begin
  { 台灣時間轉成 UTC 時間 }
 aDate := FormatDateTime( 'yyyymmdd', IncHour(
    StrToDateTime( aCmd.RCStartDate + ' 23:59:00' ), -8 ) );
  Result := Format(
    ' select icc_no from recv_nagra        ' +
    '  where icc_no = ''%s''               ' +
    '    and high_level_cmd_id = ''R10''   ' +
    '    and callback_date >= ''%s''       ',
    [Copy( aCmd.IccNo,1, 10 ), aDate] );
end;

{ ---------------------------------------------------------------------------- }


function TSoDbThread.GetConfirmDataSql(const aCmd: TRecycleCommand): String;
begin
  Result := Format(
    ' select icc_no from carecycle    ' +
    '  where seqno = ''%s''           ' +
    '    and icc_no = ''%s''          ' +
    '    and compcode = ''%d''        ' +
    '    and confirm = ''Y''          ',
    [aCmd.SeqNo, aCmd.IccNo, aCmd.CompCode] );
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.PrepareSoConnection(aSo: TAppSo);
var
  aModule: TSoDataModule;
  aErrMsg: String;
begin
  if not Assigned( aSo.DataModule ) then
  begin
    aModule := TSoDataModule.Create( nil );
    aSo.DataModule := aModule;
  end else
  begin
    aModule := TSoDataModule( aSo.DataModule );
  end;
  aModule.SoConnection.Connected := False;
  aModule.SoConnection.ConnectionString := Format(
    'Provider=MSDAORA.1;Password=%s;User ID=%s;Data Source=%s;'+
    'Persist Security Info=True', [aSo.DbLoginPass, aSo.DbLoginUser, aSo.DbAliase] );
  try
    aModule.SoConnection.Connected := True;
  except
    on E: Exception do aErrMsg := E.Message;
  end;
  if aModule.SoConnection.Connected then
  begin
    FMsgSubject.OK( Format( '系統台【%s】資料庫連線完成。', [aSo.CompName] ) );
    aSo.DbConnectStatus := dbOK;
  end else
  begin
    FMsgSubject.Error( Format( '系統台【%s】資料庫連線有誤, 訊息:%s, 60秒後將重試連線。',
      [aSO.CompName, aErrMsg] ) );
    aSo.DbConnectStatus := dbError;
  end;
  FDbSubject.Notify( aSo );
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.PrepareSoConnection;
var
  aIndex: Integer;
begin
  for aIndex := 0 to FSoList.Count - 1 do
  begin
    if ( FSoList[aIndex].DbConnectStatus in [dbNone, dbError] ) then
      PrepareSoConnection( FSoList[aIndex] );
    if ( Self.Terminated ) or ( Self.Stop ) then Break;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.PrePareCAConnection(aSo: TAppSo);
var
  aModule: TSoDataModule;
  aErrMsg: String;
begin
  if not Assigned( aSo.ComDataModule ) then
  begin
    aModule := TSoDataModule.Create( nil );
    aSo.ComDataModule := aModule;
  end else
  begin
    aModule := TSoDataModule( aSo.ComDataModule );
  end;
  aModule.CAConnection.Connected := False;
  aModule.CAConnection.ConnectionString := Format(
    'Provider=MSDAORA.1;Password=%s;User ID=%s;Data Source=%s;'+
    'Persist Security Info=True', [aSo.ComDbLoginPass, aSo.ComDbLoginUser, aSo.ComDbAliase] );
  try
    aModule.CAConnection.Connected := True;
  except
    on E: Exception do aErrMsg := E.Message;
  end;
  if aModule.CAConnection.Connected then
  begin
    FMsgSubject.OK( Format( '系統台【%s】COM區資料庫連線完成。', [aSo.CompName] ) );
    aSo.ComDbConnectStatus := dbOK;
  end else
  begin
    FMsgSubject.Error( Format( '系統台【%s】COM區資料庫連線有誤, 訊息:%s, 60秒後將重試連線。',
      [aSO.CompName, aErrMsg] ) );
    aSo.ComDbConnectStatus := dbError;
  end;
  //FDbSubject.Notify( aSo );
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.PrePareCAConnection;
var
  aIndex: Integer;
begin
  for aIndex := 0 to FSoList.Count - 1 do
  begin
    if ( FSoList[aIndex].ComDbConnectStatus in [dbNone, dbError] ) then
      PrepareCAConnection( FSoList[aIndex] );
    if ( Self.Terminated ) or ( Self.Stop ) then Break;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.UnPrepareSoConnection(aSo: TAppSo);
begin
  if Assigned( aSo.DataModule ) then
  begin
    if TSoDataModule( aSo.DataModule ).SoConnection.Connected then
      TSoDataModule( aSo.DataModule ).SoConnection.Close;
    TSoDataModule( aSo.DataModule ).Free;
    aSo.DataModule := nil;
    aSo.DbConnectStatus := dbError;
  end;
  if ( Self.Terminated or Self.Stop ) then
  begin
    FMsgSubject.OK( Format( '系統台【%s】資料庫已停止連線。', [aSo.CompName] ) );
    aSo.DbConnectStatus := dbNone;
  end;
  FDbSubject.Notify( aSo );
  //CodeSite.Send( aSo.CompName );;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.UnPrepareSoConnection;
var
  aIndex: Integer;
begin
  for aIndex := 0 to FSoList.Count - 1 do
  begin
    CodeSite.Send( FSoList[aIndex].CompName );
    UnPrepareSoConnection( FSoList[aIndex] );
    Sleep( 100 );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.UnPrepareCAConnection(aSo: TAppSo);
begin
  if Assigned( aSo.ComDataModule ) then
  begin
    if TSoDataModule( aSo.ComDataModule ).CAConnection.Connected then
      TSoDataModule( aSo.ComDataModule ).CAConnection.Close;
    TSoDataModule( aSo.ComDataModule ).Free;
    aSo.ComDataModule := nil;
    aSo.ComDbConnectStatus := dbError;
  end;
  if ( Self.Terminated or Self.Stop ) then
  begin
    FMsgSubject.OK( Format( '系統台【%s】COM區資料庫已停止連線。', [aSo.CompName] ) );
    aSo.ComDbConnectStatus := dbNone;
  end;
  //FDbSubject.Notify( aSo );
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.UnPrepareCAConection;
var
  aIndex: Integer;
begin
  for aIndex := 0 to FSoList.Count - 1 do
  begin
    UnPrepareCAConnection( FSoList[aIndex] );
    Sleep( 100 );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.GetProcessRecord(aSo: TAppSo; var aDbStatus: TDbConnectStatus);
var
  aModule: TSoDataModule;

  { ------------------------------------------------ }

  function GetSelectSql(const aFirstGet: Boolean): String;
  begin
    if aFirstGet then
    begin
      Result :=
        ' select * from carecycle     ' +
        '  where compcode = ''%s''    ' +
        '    and cmdflag in ( 0, 1 )  ' +
        '    and confirm = ''N''      ';
    end else
    begin
      Result :=
        ' select * from carecycle     ' +
        '  where compcode = ''%s''    ' +
        '    and cmdflag = 0          ' +
        '    and confirm = ''N''      ';
    end;
  end;

  { ------------------------------------------------ }

  function GetUpdateSql: String;
  begin
    Result :=
      ' update carecycle set cmdflag = 1 ' +
      '  where icc_no = ''%s''           ' +
      '    and seqno = ''%s''            ' +
      '    and compcode = ''%d''         ';
  end;

  { ------------------------------------------------ }

  function RecordToObject: TRecycleCommand;
  begin
    Result := TRecycleCommand.Create;
    Result.SeqNo := aModule.SoDataReader.FieldByName( 'SEQNO' ).AsString;
    Result.IccNo := aModule.SoDataReader.FieldByName( 'ICC_NO' ).AsString;
    Result.StbNo := aModule.SoDataReader.FieldByName( 'STB_NO' ).AsString;
    Result.StbFlag := aModule.SoDataReader.FieldByName( 'STB_FLAG' ).AsInteger;
    Result.StbAutoFlag := aModule.SoDataReader.FieldByName( 'STBAUTOCB' ).AsInteger;
    Result.CompCode := aModule.SoDataReader.FieldByName( 'COMPCODE' ).AsInteger;
    Result.CompName := aSo.CompName;
    Result.RCStartDate := EmptyStr;
    if not VarIsNull( aModule.SoDataReader.FieldByName( 'RCSTARTDATE' ).Value ) then
      Result.RCStartDate := FormatDateTime( 'yyyy/mm/dd', aModule.SoDataReader.FieldByName( 'RCSTARTDATE' ).AsDateTime );
    Result.RCEndDate := EmptyStr;
    if not VarIsNull( aModule.SoDataReader.FieldByName( 'RCENDDATE' ).Value ) then
      Result.RCEndDate := FormatDateTime( 'yyyy/mm/dd', aModule.SoDataReader.FieldByName( 'RCENDDATE' ).AsDateTime );
    Result.Operator := aModule.SoDataReader.FieldByName( 'OPERATOR' ).AsString;
    Result.UpdTime := aModule.SoDataReader.FieldByName( 'UPDTIME' ).AsDateTime;
    Result.LastDownloadTime := EmptyStr;
    if not VarIsNull( aModule.SoDataReader.FieldByName( 'LASTDOWNLOADTIME' ).Value ) then
      Result.LastDownloadTime := FormatDateTime( 'yyyy/mm/dd hh:nn:ss',
        aModule.SoDataReader.FieldByName( 'LASTDOWNLOADTIME' ).AsDateTime );
    Result.CMDFlag := aModule.SoDataReader.FieldByName( 'CMDFLAG' ).AsInteger;
    Result.RecycleText := aModule.SoDataReader.FieldByName( 'RECYCLETEXT' ).AsString;
    Result.TransNum := aModule.SoDataReader.FieldByName( 'TRANSNUM' ).AsString;
    Result.Confirm := aModule.SoDataReader.FieldByName( 'CONFIRM' ).AsString;
    Result.ErrMsg := aModule.SoDataReader.FieldByName( 'ERR_MSG' ).AsString;
    Result.Data := nil;
  end;

  { ------------------------------------------------ }

var
  aCmd: TRecycleCommand;
  aUpdateSql, aSelectSql: String;
begin
  aModule := TSoDataModule( aSo.DataModule );
  aModule.SoDataReader.Close;
  { 重新啟動, 必須連 處理中 的資料一起抓出, 取出一次後, 後續只須抓取
    未處理的 }
  aSelectSql := GetSelectSql( aSo.FirstQuery );
  { 抓出來加進 List 後, 更新狀態 }
  aUpdateSql := GetUpdateSql;
  try
    aModule.SoDataReader.SQL.Text := Format( aSelectSql,
      [aSo.CompCode] );
    aModule.SoDataReader.Open;
  except
    on E: Exception do
    begin
      FMsgSubject.Error( Format( '系統台【%s】取出待處理資料有誤, 訊息:%s, 60秒後將重試連線。',
        [aSo.CompName, E.Message] ) );
      aDbStatus := dbError;
    end;
  end;
  if aModule.SoDataReader.Active then
  begin
    aModule.SoDataReader.First;
    while not aModule.SoDataReader.Eof do
    begin
      try
        if FProcList.IndexOf( aModule.SoDataReader.FieldByName( 'SEQNO' ).AsString,
          aModule.SoDataReader.FieldByName( 'ICC_NO' ).AsString ) = -1 then
        begin
          aCmd := RecordToObject;
          if ( aCmd.CMDFlag = 0 ) then
          begin
            { 更新成處理中 }
            aModule.SoDataWriter.SQL.Text := Format( aUpdateSql,
              [aCmd.IccNo, aCmd.SeqNo, aCmd.CompCode] );
            aModule.SoDataWriter.ExecSQL;
          end;  
          { 加入佇列 }
          FProcList.Add( aCmd );
        end;  
      except
        on E: Exception do
        begin
          FMsgSubject.Error( Format( '系統台【%s】待處理資料加入佇列有誤, 訊息:%s, 60秒後將重試連線。',
            [aSo.CompName, E.Message] ) );
          aDbStatus := dbError;
          Break;
        end;
      end;
      if ( Self.Stop ) or ( Self.Terminated ) then Break;
      aModule.SoDataReader.Next;
    end;
    aModule.SoDataReader.Close;
    aSo.FirstQuery := False;
  end;
  if aDbStatus <> dbError then aDbStatus := dbOK;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.GetProcessRecord;
var
  aIndex: Integer;
  aStatus: TDbConnectStatus;
begin
  { 先重試連線 }
  for aIndex := 0 to FSoList.Count - 1 do
  begin
    if ( FSoList[aIndex].DbConnectStatus in [dbError] ) then
    begin
      if CanRetryDbConnection( FSoList[aIndex].LastDbErrTime ) then
        PrepareSoConnection( FSoList[aIndex] );
    end;
    {}
    if ( FSoList[aIndex].ComDbConnectStatus in [dbError] ) then
    begin
      if CanRetryComDbConnection( FSoList[aIndex].LastComDbErrTime ) then
        PrePareCAConnection( FSoList[aIndex] );
    end;
    Sleep( 100 );
    if ( Self.Stop ) or ( Self.Terminated ) then Break;
  end;
  { 抓取資料 }
  for aIndex := 0 to FSoList.Count - 1 do
  begin
    if ( FSoList[aIndex].DbConnectStatus in [dbOK, dbActive] ) and
       ( FSoList[aIndex].ComDbConnectStatus in [dbOK, dbActive] ) then
    begin
      FSoList[aIndex].DbConnectStatus := dbActive;
      FDbSubject.Notify( FSoList[aIndex] );
      try
        aStatus := FSoList[aIndex].DbConnectStatus;
        GetProcessRecord( FSoList[aIndex], aStatus );
        Sleep( 100 );
        if ( Self.Stop ) or ( Self.Terminated ) then Break;
      finally
        FSoList[aIndex].DbConnectStatus := aStatus;
        FDbSubject.Notify( FSoList[aIndex] );
      end;
      if ( aStatus in [dbError] ) then
        UnPrepareSoConnection( FSoList[aIndex] );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.GetSendNagra;
var
  aIndex, aErrCount: Integer;
  aSo: TAppSo;
  aCmd: TRecycleCommand;
  aExecuteCmd, aCurrentSeqNo, aCmdStatus, aCmdErrMsg, aIccNo: String;
  aModule: TSoDataModule;
begin
  { 先重試連線 }
  for aIndex := 0 to FSoList.Count - 1 do
  begin
    if ( FSoList[aIndex].DbConnectStatus in [dbError] ) then
    begin
      if CanRetryDbConnection( FSoList[aIndex].LastDbErrTime ) then
        PrepareSoConnection( FSoList[aIndex] );
    end;
    {}
    if ( FSoList[aIndex].ComDbConnectStatus in [dbError] ) then
    begin
      if CanRetryComDbConnection( FSoList[aIndex].LastComDbErrTime ) then
        PrePareCAConnection( FSoList[aIndex] );
    end;
    Sleep( 100 );
    if ( Self.Stop ) or ( Self.Terminated ) then Break;
  end;
  { 從 Send_nagra 抓處理完的資料 }
  aErrCount := 0;
  for aIndex := 0 to FProcList.Count - 1 do
  begin
    if ( ( aIndex mod 10 ) = 0 ) then Sleep( 100 );
    if ( Self.Terminated ) or ( Self.Stop ) then Break;
    try
      aCmd := FProcList[aIndex];
      aIccNo := aCmd.IccNo;
      aSo := GetAppSo( IntToStr( aCmd.CompCode ) );
      if not ( aSo.ComDbConnectStatus in [dbOK, dbActive] ) then Continue;
      if not ( aSo.DbConnectStatus in [dbOK, dbActive] ) then Continue;
      FCmdHelper.RecycleCommand := aCmd;
      aExecuteCmd := FCmdHelper.CurrentExecuteCmd;
      if ( aExecuteCmd = EmptyStr ) then Continue;
      { 假如是等待 Callbakc 資料的指令, 則跳過,不須取 Send_Nagra 的資料 }
      if ( aExecuteCmd = '99' ) then Continue;
      { 最後指令 --> 解配對, 若未確認, 尚為送送指令, 所以不須抓 }
      if ( aExecuteCmd = '52_2' ) and ( aCmd.Confirm = 'N' ) then Continue;
      aCurrentSeqNo := FCmdHelper.FindTransByCmd( aExecuteCmd );
      aModule := TSoDataModule( aSo.ComDataModule );
      aModule.CADataReader.Close;
      aModule.CADataReader.SQL.Text := Format(
        ' select cmd_status, err_msg                ' +
        '  from send_nagra where seqno = ''%s''     ' +
        '   and cmd_status in ( ''C'', ''E'' )      ' +
        '   and compcode = ''%d''                   ',
        [aCurrentSeqNo, aCmd.CompCode] );
      aModule.CADataReader.Open;
      if ( aModule.CADataReader.IsEmpty ) then
      begin
        aModule.CADataReader.Close;
        Continue;
      end;
      {}
      //aCmd.UpdTime := Now;
      aCmdStatus := aModule.CADataReader.FieldByName( 'CMD_STATUS' ).AsString;
      aCmdErrMsg := aModule.CADataReader.FieldByName( 'ERR_MSG' ).AsString;
      if ( aCmdStatus = 'C' ) then
        FCmdHelper.CmdOk( aExecuteCmd )
      else if ( aCmdStatus = 'E' ) then
        FCmdHelper.CmdErr( aExecuteCmd, aCmdErrMsg );
      aModule.CADataReader.Close;
      {}
      aModule := TSoDataModule( aSo.DataModule );
      aModule.SoDataWriter.Close;
      aModule.SoDataWriter.SQL.Text := GetUpdateCARecycleSql_2( aCmd, aExecuteCmd );
      aModule.SoDataWriter.ExecSQL;
      {}
      aModule := TSoDataModule( aSo.ComDataModule );
      aModule.CADataWriter.Close;
      aModule.CADataWriter.SQL.Text := Format(
        ' delete from send_nagra where seqno = ''%s''    ' +
        '    and cmd_status in ( ''C'', ''E'' )          ' +
        '    and compcode = ''%d''                       ',
        [aCurrentSeqNo, aCmd.CompCode] );
      aModule.CADataWriter.ExecSQL;
      {}
      FCmdSubject.Notify( aCmd ); 
    except
      on E: Exception do
      begin
        FMsgSubject.Error( Format( 'ICC 卡號:【%s】抓取SEND_NAGRA有誤, 訊息:%s。',
          [aIccNo, E.Message] ) );
        Inc( aErrCount );
      end;
    end;
    if ( aErrCount >= 5 ) then
    begin
      FMsgSubject.Error( '抓取SEND_NAGRA已失敗多次, 所有資料庫將離線。' );
      aSo.DbConnectStatus := dbError;
      aSo.ComDbConnectStatus := dbError;
      UnPrepareSoConnection;
      UnPrepareCAConection;
      Exit;
    end;
  end;
  {}
  if ( Self.Terminated ) or ( Self.Stop ) then Exit;
  { 將錯誤的直接移掉, 全部 OK 的也移掉 }
  for aIndex := FProcList.Count - 1 downto 0 do
  begin
    if ( FProcList[aIndex].CMDFlag in [3,4] ) then
      FProcList.Delete( aIndex );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.WriteToSendNagra;
var
  aIndex, aErrCount: Integer;
  aCmd: TRecycleCommand;
  aExecuteCmd, aIccNo: String;
  aSo: TAppSo;
  aModule: TSoDataModule;

  { ------------------------------------------------ }

  function GetSendNagraSeqNo: String;
  begin
    aModule.CADataReader.Close;
    aModule.CADataReader.SQL.Text :=
      ' select s_sendnagra_seqno.nextval from dual ';
    aModule.CADataReader.Open;
    Result := aModule.CADataReader.Fields[0].AsString;
    aModule.CADataReader.Close;
  end;

  { ------------------------------------------------ }

var
  aSeqNo: String;
  aSuccessCount, aCount: Integer;
begin
  { 先重試連線 }
  for aIndex := 0 to FSoList.Count - 1 do
  begin
    if ( FSoList[aIndex].DbConnectStatus in [dbError] ) then
    begin
      if CanRetryDbConnection( FSoList[aIndex].LastDbErrTime ) then
        PrepareSoConnection( FSoList[aIndex] );
    end;
    {}
    if ( FSoList[aIndex].ComDbConnectStatus in [dbError] ) then
    begin
      if CanRetryComDbConnection( FSoList[aIndex].LastComDbErrTime ) then
        PrePareCAConnection( FSoList[aIndex] );
    end;
    Sleep( 100 );
    if ( Self.Stop ) or ( Self.Terminated ) then Break;
  end;
  aErrCount := 0;
  aCount := 0;
  aSuccessCount := 0;
  for aIndex := 0 to FProcList.Count - 1 do
  begin
    try
      aCmd := FProcList[aIndex];
      if not Assigned( aCmd.Data ) then FCmdSubject.Notify( aCmd );
      Sleep( 10 );
      aIccNo := aCmd.IccNo;
      FCmdHelper.RecycleCommand := aCmd;
      aExecuteCmd := FCmdHelper.NextExecuteCmd;
      if ( aExecuteCmd = EmptyStr ) then Continue;

      { 52_2 --> 解配對, 必須等待 user 確認 }
      if ( aExecuteCmd = '52_2' ) and ( aCmd.Confirm = 'N' ) then Continue;
      aSo := GetAppSo( IntToStr( aCmd.CompCode ) );
      if not ( aSo.ComDbConnectStatus in [dbOK, dbActive] ) then Continue;
      if not ( aSo.DbConnectStatus in [dbOK, dbActive] ) then Continue;
      { 99 --> wait for callback, 不須下指令 }
      if ( aExecuteCmd <> '99' ) then
      begin
        { 切到 COM }
        aModule := TSoDataModule( aSo.ComDataModule );
        aModule.CADataWriter.SQL.Clear;
        aModule.CADataWriter.SQL.Text := GetWriteSendNagraSql( aCmd, aExecuteCmd );
        aSeqNo := GetSendNagraSeqNo;
        aModule.CADataWriter.Parameters.ParamByName( 'seqno' ).Value := aSeqNo;
        aModule.CADataWriter.ExecSQL;
      end;
      FCmdHelper.CmdProc( aExecuteCmd, aSeqNo );
      { 切到 So }
      aModule := TSoDataModule( aSo.DataModule );
      aModule.SoDataWriter.Close;
      aModule.SoDataWriter.SQL.Text := GetUpdateCARecycleSql_1( aCmd );
      aModule.SoDataWriter.ExecSQL;
      FCmdSubject.Notify( aCmd );
      Inc( aSuccessCount );
      Inc( aCount );
      if ( aCount >= 10 ) then
      begin
        Sleep( 20 );
        aCount := 0;
      end;
    except
      on E: Exception do
      begin
        FMsgSubject.Error( Format( 'ICC 卡號:【%s】寫入SEND_NAGRA有誤, 訊息:%s。',
          [aIccNo, E.Message] ) );
        Inc( aErrCount );
      end;
    end;
    if ( aErrCount >= 5 ) then
    begin
      FMsgSubject.Error( '待處理資料寫入SEND_NAGRA已失敗多次, 所有資料庫將離線。' );
      aSo.DbConnectStatus := dbError;
      aSo.ComDbConnectStatus := dbError;
      UnPrepareSoConnection;
      UnPrepareCAConection;
      Break;
    end;
    if ( Self.Stop ) or ( Self.Terminated ) then Break;
  end;
  if ( aSuccessCount > 0 ) and ( aErrCount <= 0 ) then
  begin
    FMsgSubject.Normal( Format( 'SEND_NAGRA 寫入 %d 筆待處理資料。',
      [aSuccessCount] ) );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TSoDbThread.GetAppSo(const aCompCode: String): TAppSo;
var
  aIndex: Integer;
begin
  Result := nil;
  aIndex := FSoList.IndexOf( aCompCode );
  if aIndex >= 0 then Result := FSoList[aIndex];
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.GetCallbackData;
var
  aIndex, aErrCount: Integer;
  aCmd: TRecycleCommand;
  aSo: TAppSo;
  aModule: TSoDataModule;
  aIccNo, aExecuteCmd: String;
  aFound: Boolean;
begin
  { 先重試連線 }
  for aIndex := 0 to FSoList.Count - 1 do
  begin
    if ( FSoList[aIndex].DbConnectStatus in [dbError] ) then
    begin
      if CanRetryDbConnection( FSoList[aIndex].LastDbErrTime ) then
        PrepareSoConnection( FSoList[aIndex] );
    end;
    {}
    if ( FSoList[aIndex].ComDbConnectStatus in [dbError] ) then
    begin
      if CanRetryComDbConnection( FSoList[aIndex].LastComDbErrTime ) then
        PrePareCAConnection( FSoList[aIndex] );
    end;
    Sleep( 100 );
    if ( Self.Stop ) or ( Self.Terminated ) then Break;
  end;
  { 從 recv_nagra 抓callback回來的資料 }
  aErrCount := 0;
 for aIndex := 0 to FProcList.Count - 1 do
  begin
    if ( ( aIndex mod 10 ) = 0 ) then Sleep( 100 );
    if ( Self.Terminated ) or ( Self.Stop ) then Break;
    try
      aCmd := FProcList[aIndex];
      aIccNo := aCmd.IccNo;
      aSo := GetAppSo( IntToStr( aCmd.CompCode ) );
      if not ( aSo.ComDbConnectStatus in [dbOK, dbActive] ) then Continue;
      if not ( aSo.DbConnectStatus in [dbOK, dbActive] ) then Continue;
      FCmdHelper.RecycleCommand := aCmd;
      aExecuteCmd := FCmdHelper.CurrentExecuteCmd;
      if ( aExecuteCmd <> '99' ) then Continue;
      aModule := TSoDataModule( aSo.ComDataModule );
      aModule.CADataReader.Close;
      aModule.CADataReader.SQL.Text := GetCallbackDataSql( aCmd );
      aModule.CADataReader.Open;
      aFound := not aModule.CADataReader.IsEmpty;
      aModule.CADataReader.Close;
      if aFound then
      begin
        FCmdHelper.CmdOk( aExecuteCmd );
        {}
        aModule := TSoDataModule( aSo.DataModule );
        aModule.SoDataWriter.Close;
        aModule.SoDataWriter.SQL.Text := GetUpdateCARecycleSql_2( aCmd, aExecuteCmd );
        aModule.SoDataWriter.ExecSQL;
        {}
        FCmdSubject.Notify( aCmd );
      end;
    except
        on E: Exception do
      begin
        FMsgSubject.Error( Format( 'ICC 卡號:【%s】抓取RECV_NAGRA有誤, 訊息:%s。',
          [aIccNo, E.Message] ) );
        Inc( aErrCount );
      end;
    end;
    if ( aErrCount >= 5 ) then
    begin
      FMsgSubject.Error( '抓取SEND_NAGRA已失敗多次, 所有資料庫將離線。' );
      aSo.DbConnectStatus := dbError;
      aSo.ComDbConnectStatus := dbError;
      UnPrepareSoConnection;
      UnPrepareCAConection;
      Exit;
    end;
  end;
  {}
  if ( Self.Terminated ) or ( Self.Stop ) then Exit;
  { 將錯誤的直接移掉 }
  for aIndex := FProcList.Count - 1 downto 0 do
  begin
    if ( FProcList[aIndex].CMDFlag = 4 ) then
      FProcList.Delete( aIndex );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSoDbThread.GetConfirmData;
var
  aIndex, aErrCount: Integer;
  aCmd: TRecycleCommand;
  aModule: TSoDataModule;
  aIccNo, aSeqNo, aExecuteCmd: String;
  aSo: TAppSo;

  { ------------------------------------------------ }

  function GetSendNagraSeqNo: String;
  begin
    aModule.CADataReader.Close;
    aModule.CADataReader.SQL.Text :=
      ' select s_sendnagra_seqno.nextval from dual ';
    aModule.CADataReader.Open;
    Result := aModule.CADataReader.Fields[0].AsString;
    aModule.CADataReader.Close;
  end;

  { ------------------------------------------------ }
  
begin
  { 先重試連線 }
  for aIndex := 0 to FSoList.Count - 1 do
  begin
    if ( FSoList[aIndex].DbConnectStatus in [dbError] ) then
    begin
      if CanRetryDbConnection( FSoList[aIndex].LastDbErrTime ) then
        PrepareSoConnection( FSoList[aIndex] );
    end;
    {}
    if ( FSoList[aIndex].ComDbConnectStatus in [dbError] ) then
    begin
      if CanRetryComDbConnection( FSoList[aIndex].LastComDbErrTime ) then
        PrePareCAConnection( FSoList[aIndex] );
    end;
    Sleep( 100 );
    if ( Self.Stop ) or ( Self.Terminated ) then Break;
  end;
  {  }
  aErrCount := 0;
  for aIndex := 0 to FProcList.Count - 1 do
  begin
    if ( ( aIndex mod 10 ) = 0 ) then Sleep( 100 );
    if ( Self.Terminated ) or ( Self.Stop ) then Break;
    try
      aCmd := FProcList[aIndex];
      aIccNo := aCmd.IccNo;
      aSo := GetAppSo( IntToStr( aCmd.CompCode ) );
      if not ( aSo.ComDbConnectStatus in [dbOK, dbActive] ) then Continue;
      if not ( aSo.DbConnectStatus in [dbOK, dbActive] ) then Continue;
      FCmdHelper.RecycleCommand := aCmd;
      aExecuteCmd := FCmdHelper.NextExecuteCmd;
      if ( aExecuteCmd <> '52_2' ) then Continue;
      aModule := TSoDataModule( aSo.DataModule );
      aModule.SoDataReader.Close;
      aModule.SoDataReader.SQL.Text := GetConfirmDataSql( aCmd );
      aModule.SoDataReader.Open;
      aCmd.Confirm := IIF( aModule.SoDataReader.IsEmpty, 'N', 'Y' );
      aModule.SoDataReader.Close;
      {}
   except
        on E: Exception do
      begin
        FMsgSubject.Error( Format( 'ICC 卡號:【%s】抓取使用者確認資料有誤, 訊息:%s。',
          [aIccNo, E.Message] ) );
        Inc( aErrCount );
      end;
    end;
    if ( aErrCount >= 5 ) then
    begin
      FMsgSubject.Error( '抓取使用者確認資料已失敗多次, 所有資料庫將離線。' );
      aSo.DbConnectStatus := dbError;
      aSo.ComDbConnectStatus := dbError;
      UnPrepareSoConnection;
      UnPrepareCAConection;
      Exit;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
