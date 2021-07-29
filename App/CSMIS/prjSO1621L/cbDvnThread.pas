unit cbDvnThread;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, DB, DBClient, ADODB,
  cbMain;

type
  TDvnThread = class(TThread)
  private
    { Private declarations }
    FDataSet: TClientDataSet;
    FConnection: TADOConnection;
    FDataReader: TADOQuery;
    FDataWriter: TADOCommand;
    FCompCode: String;
    FSmsOperator: String;
    FChannelDataSet: TClientDataSet;
    FOracleToday: TDateTime;
    FAccTimeOut: Integer;
    FGuiInfo: TGuiInfo;
    FWantToOpenChannel: Boolean;
    FCmdA1Text: String;
    FReSendErrCmd: Boolean;
    FChNotes: String;
    procedure SetSourceDataSet(const AValue: TClientDataSet);
    procedure SetChannelDataSet(const AValue: TClientDataSet);
    procedure WriteToSendDvn(const aCmdType: Integer);
    procedure WriteA1;
    procedure WriteB1;
    procedure ChangeErrCmdStatus;
    procedure MarkHasBeenLog;
    procedure UpdateGUI;
    procedure UpdateAllData;
    procedure DeleteSuccessCmd(const aCmdType: Integer);
    procedure UpdateSOAC0501C;
    procedure UpdateSOAC0502;
    function GetOracleToday: TDateTime;
    function GetAccTimeOut: Integer;
    function GetWantToOpenChannel: Boolean;
    function GetSeqNo: String;
    function BuildNotes: String;
    function IsProcessing(const aCmdType: Integer): Boolean;
    procedure CheckCmdStatus;
    function CheckThreadCanbeBreak: Boolean;
  protected
    procedure Execute; override;
  public
    constructor Create(CreateSuspended: Boolean);
    destructor Destroy; override;
    property SourceDataSet: TClientDataSet read FDataSet write SetSourceDataSet;
    property SourceConnection: TADOConnection read FConnection write FConnection;
    property ChannelDataSet: TClientDataSet read FChannelDataSet write SetChannelDataSet;
    property CompCode: String read FCompCode write FCompCode;
    property SmsOperator: String read FSmsOperator write FSmsOperator;
    property CmdA1Text: String read FCmdA1Text write FCmdA1Text;
    property ReSendErrCmd: Boolean read FReSendErrCmd write FReSendErrCmd;
  end;

var DvnThreadErrMsg: String;

implementation

uses cbUtilis, DateUtils;

{ Important: Methods and properties of objects in visual components can only be
  used in a method called using Synchronize, for example,

      Synchronize(UpdateCaption);

  and UpdateCaption could look like,

    procedure TWatchDog.UpdateCaption;
    begin
      Form1.Caption := 'Updated in a thread';
    end; }

{ TWatchDog }

{ ---------------------------------------------------------------------------- }

constructor TDvnThread.Create(CreateSuspended: Boolean);
begin
  inherited Create( True );
  DvnThreadErrMsg := EmptyStr;
  FDataSet := TClientDataSet.Create( nil );
  FChannelDataSet := TClientDataSet.Create( nil );
  FDataReader := TADOQuery.Create( nil );
  FDataWriter := TADOCommand.Create( nil );
  FWantToOpenChannel := False;
  FreeOnTerminate := True;
  FCmdA1Text := 'A1';
  FReSendErrCmd := False;
  FChNotes := EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

destructor TDvnThread.Destroy;
begin
  FDataSet.Free;
  FChannelDataSet.Free;
  FDataReader.Free;
  FDataWriter.Free;
  FConnection := nil;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TDvnThread.Execute;
begin
  FDataReader.Connection := FConnection;
  FDataWriter.Connection := FConnection;
  Sleep( 100 );
  FOracleToday := GetOracleToday;
  FAccTimeOut := GetAccTimeOut;
  FWantToOpenChannel := GetWantToOpenChannel;
  Sleep( 100 );
  if ( FReSendErrCmd ) then ChangeErrCmdStatus;
  MarkHasBeenLog;
  {}
  while not Self.Terminated do
  begin
    try
      WriteToSendDvn( CmdA1 );
      Sleep( 500 );
      {}
      if Self.Terminated then Break;
      {}
      CheckCmdStatus;
      Sleep( 500 );
      if Self.Terminated then Break;
      CheckCmdStatus;
      Sleep( 500 );
      if Self.Terminated then Break;
      CheckCmdStatus;
      Sleep( 500 );
      if Self.Terminated then Break;
      {}
      WriteToSendDvn( CmdB1 );
      Sleep( 500 );
      if Self.Terminated then Break;
      {}
      if CheckThreadCanbeBreak then Break;
    except
      on E: Exception do
      begin
        DvnThreadErrMsg := E.Message;
        Break;
      end;
    end;
  end;
  UpdateAllData;
  FDataReader.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TDvnThread.SetSourceDataSet(const AValue: TClientDataSet);
begin
  FDataSet.Data := AValue.Data;
end;

{ ---------------------------------------------------------------------------- }

procedure TDvnThread.SetChannelDataSet(const AValue: TClientDataSet);
begin
  FChannelDataSet.Data := AValue.Data;
end;

{ ---------------------------------------------------------------------------- }

procedure TDvnThread.ChangeErrCmdStatus;
begin
  FDataSet.First;
  while not FDataSet.Eof do
  begin
    FDataSet.Edit;
    if ( FDataSet.FieldByName( 'CMDA1' ).AsString = 'E' ) then
    begin
      FDataSet.FieldByName( 'CMDA1' ).AsString := EmptyStr;
      FGuiInfo.RecNo := FDataSet.FieldByName( 'RECNO' ).AsInteger;
      FGuiInfo.SeqNo := FDataSet.FieldByName( 'SEQNO1' ).AsString;
      FGuiInfo.CmdType := CmdA1;
      FGuiInfo.CmdStatus := EmptyStr;
      FGuiInfo.ErrMsg := EmptyStr;
      Synchronize( UpdateGUI );
    end;
    if ( FDataSet.FieldByName( 'CMDB1' ).AsString = 'E' ) then
    begin
      FDataSet.FieldByName( 'CMDB1' ).AsString := EmptyStr;
      FGuiInfo.RecNo := FDataSet.FieldByName( 'RECNO' ).AsInteger;
      FGuiInfo.SeqNo := FDataSet.FieldByName( 'SEQNO2' ).AsString;
      FGuiInfo.CmdType := CmdB1;
      FGuiInfo.CmdStatus := EmptyStr;
      FGuiInfo.ErrMsg := EmptyStr;
      Synchronize( UpdateGUI );
    end;
    {}
    FDataSet.Post;
    FDataSet.Next;
    Sleep( 50 );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDvnThread.MarkHasBeenLog;
begin
  FDataSet.First;
  while not FDataSet.Eof do
  begin
    FDataSet.Edit;
    if ( FDataSet.FieldByName( 'CMDA1' ).AsString = 'C' ) then
      FDataSet.FieldByName( 'HASLOG1' ).AsString := 'Y';
    if ( FDataSet.FieldByName( 'CMDB1' ).AsString = 'C' ) then
      FDataSet.FieldByName( 'HASLOG2' ).AsString := 'Y';
    FDataSet.Post;
    FDataSet.Next;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDvnThread.UpdateGUI;
begin
  fmMain.ProgressReport( FGuiInfo.CmdType, FGuiInfo.RecNo,
    FGuiInfo.CmdStatus, FGuiInfo.ErrMsg );
end;

{ ---------------------------------------------------------------------------- }

procedure TDvnThread.UpdateAllData;
begin
  FDataSet.First;
  while not FDataSet.Eof do
  begin
    DeleteSuccessCmd( CmdA1 );
    DeleteSuccessCmd( CmdB1 );
    UpdateSOAC0501C;
    UpdateSOAC0502;
    FDataSet.Next;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDvnThread.DeleteSuccessCmd(const aCmdType: Integer);
var
  aCmdFieldName, aSeqFieldName: String;
begin
  aCmdFieldName := EmptyWideStr;
  aSeqFieldName := EmptyStr;
  if ( aCmdType = CmdA1 ) then
  begin
    aCmdFieldName := 'CMDA1';
    aSeqFieldName := 'SEQNO1';
  end else
  if ( aCmdType = CmdB1 ) then
  begin
    aCmdFieldName := 'CMDB1';
    aSeqFieldName := 'SEQNO2';
  end;
  {}
  if ( aCmdFieldName = EmptyStr ) then Exit;
  {}
  if ( FDataSet.FieldByName( aCmdFieldName ).AsString = 'C' ) then
  begin
    FDataWriter.CommandText := Format(
      ' DELETE FROM SEND_DVN WHERE SEQNO = ''%s'' ',
      [FDataSet.FieldByName( aSeqFieldName ).AsString] );
    FDataWriter.Execute;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDvnThread.UpdateSOAC0501C;
var
  aRecordEffect: Integer;
begin
  FDataWriter.CommandText := Format(
    ' UPDATE SOAC0501C SET SMARTCARDNO = ''%s'',   ' +
    '   LASTPAIRDATE = SYSDATE                     ' +
    '  WHERE COMPCODE = ''%s''                     ' +
    '    AND STBSNO  = ''%s''                      ',
    [FDataSet.FieldByName( 'ICCNO' ).AsString,
     FCompCode,
     FDataSet.FieldByName( 'STBNO' ).AsString] );
  FDataWriter.Execute( aRecordEffect, EmptyParam );
  if ( aRecordEffect <= 0 ) then
  begin
    FDataWriter.CommandText := Format(
      ' INSERT INTO SOAC0501C ( COMPCODE, STBSNO,   ' +
      '   SMARTCARDNO, LASTPAIRDATE )               ' +
      '  VALUES ( ''%s'', ''%s'', ''%s'', SYSDATE ) ',
      [FCompCode,
       FDataSet.FieldByName( 'STBNO' ).AsString,
       FDataSet.FieldByName( 'ICCNO' ).AsString] );
    FDataWriter.Execute;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDvnThread.UpdateSOAC0502;

  { -------------------------------------- }

  function GetChInfo(const Value: String; var aName, aCode: String): String;
  begin
    aName := EmptyStr;
    aCode := EmptyStr;
    FChannelDataSet.First;
    if ( FChannelDataSet.Locate( 'CHANNELID', Value, [] ) ) then
    begin
      aName := FChannelDataSet.FieldByName( 'DESCRIPTION' ).AsString;
      aCode := FChannelDataSet.FieldByName( 'CODENO' ).AsString;
    end;  
  end;
  
  { -------------------------------------- }

var
  aOldNotes, aChText, aChCode, aChName, aStarDate, aExpiryDate: String;
  aUpdTime, aChId: String;
begin
  aUpdTime := DateConvert( FOracleToday, True ) + ' ' +
    FormatDateTime( 'hh:nn:ss', FOracleToday );
  if ( Copy( aUpdTime, 1, 1 ) = #32 ) then Delete( aUpdTime, 1, 1 );
  if ( FDataSet.FieldByName( 'CMDA1' ).AsString = 'C' ) and
     ( FDataSet.FieldByName( 'HASLOG1' ).AsString <> 'Y'  ) then
  begin
    FDataWriter.CommandText := Format(
      ' INSERT INTO SOAC0502 ( COMPCODE, STBSNO, SMARTCARDNO, ' +
      '    MODETYPE, UPDTIME, UPDEN )                         ' +
      ' VALUES ( ''%S'', ''%S'', ''%S'', ''%S'', ''%S'',      ' +
      '    ''%S'' ) ',
      [FCompCode,
       FDataSet.FieldByName( 'STBNO' ).AsString,
       FDataSet.FieldByName( 'ICCNO' ).AsString,
       FCmdA1Text, aUpdTime, FSmsOperator] );
    FDataWriter.Execute;
  end;
  if ( FWantToOpenChannel ) and
     ( FDataSet.FieldByName( 'CMDB1' ).AsString = 'C' ) and
     ( FDataSet.FieldByName( 'HASLOG2' ).AsString <> 'Y' ) then
  begin
    aOldNotes := FChNotes;
    repeat
      aChText := ExtractValue( aOldNotes );
      //B
      ExtractValue( aChText, '~' );
      //頻道代碼
      aChCode := ExtractValue( aChText, '~' );
      //起始日
      aStarDate := ExtractValue( aChText, '~' );
      //截止日
      aExpiryDate := ExtractValue( aChText, '~' );
      //頻道名稱,跟代碼
      GetChInfo( aChCode, aChName, aChId );
      //
      FDataWriter.CommandText := Format(
        ' INSERT INTO SOAC0502 ( COMPCODE, STBSNO, SMARTCARDNO, ' +
        '    MODETYPE, CHCODE, CHNAME, UPDTIME, UPDEN,          ' +
        '    AUTHORSTARTDATE, AUTHORSTOPDATE )                  ' + 
        ' VALUES ( ''%S'', ''%S'', ''%S'', ''%S'', ''%S'',      ' +
        '          ''%S'', ''%S'', ''%S'',                      ' +
        '          TO_DATE( ''%S'', ''YYYYMMDD'' ),             ' +
        '          TO_DATE( ''%S'', ''YYYYMMDD'' ) )            ',
        [FCompCode,
         FDataSet.FieldByName( 'STBNO' ).AsString,
         FDataSet.FieldByName( 'ICCNO' ).AsString,
         'B1', aChId, aChName, aUpdTime, FSmsOperator, aStarDate, aExpiryDate] );
      FDataWriter.Execute;
    until ( aOldNotes = EmptyStr );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TDvnThread.GetSeqNo: String;
begin
  FDataReader.Close;
  FDataReader.SQL.Text := ' SELECT S_SENDDVN_SEQNO.NEXTVAL FROM DUAL ';
  FDataReader.Open;
  Result := FDataReader.Fields[0].AsString;
  FDataReader.Open;
end;

{ ---------------------------------------------------------------------------- }

function TDvnThread.GetOracleToday: TDateTime;
begin
  FDataReader.Close;
  FDataReader.SQL.Text := ' SELECT SYSDATE FROM DUAL ';
  FDataReader.Open;
  Result := FDataReader.Fields[0].AsDateTime;
  FDataReader.Close;
end;

{ ---------------------------------------------------------------------------- }

function TDvnThread.GetAccTimeOut: Integer;
begin
  FDataReader.Close;
  FDataReader.SQL.Text := 'select acctimeout from so041 ';
  FDataReader.Open;
  Result := FDataReader.Fields[0].AsInteger;
  if ( Result <= 0 ) then Result := 90;
  FDataReader.Close;
end;

{ ---------------------------------------------------------------------------- }

function TDvnThread.GetWantToOpenChannel: Boolean;
begin
  Result := False;
  FChannelDataSet.First;
  while not FChannelDataSet.Eof do
  begin
    Result := ( FChannelDataSet.FieldByName( 'SELECTED' ).AsString = 'Y' );
    if ( Result ) then Break;
    FChannelDataSet.Next;
  end;
  FChannelDataSet.First;
end;

{ ---------------------------------------------------------------------------- }

function TDvnThread.BuildNotes: String;
var
  aChText, aStart, aExpiry: String;
  aEndPos: Integer;
begin
  Result := EmptyStr;
  FChannelDataSet.First;
  while not FChannelDataSet.Eof do
  begin
    if ( FChannelDataSet.FieldByName( 'SELECTED' ).AsString = 'Y' ) then
    begin
      aStart := FormatDateTime( 'yyyymmdd', FOracleToday );
      aExpiry := FormatDateTime(  'yyyymmdd',
        FOracleToday +
        FChannelDataSet.FieldByName( 'CHANCEDAYS' ).AsInteger +
        FChannelDataSet.FieldByName( 'BUFFERDAYS' ).AsInteger );
      aChText := Format( 'B~%s~%s~%s',
        [FChannelDataSet.FieldByName( 'CHANNELID' ).AsString,
         aStart, aExpiry] );
      Result := ( Result + aChText + ',' );
    end;
    FChannelDataSet.Next;
  end;
  aEndPos := Length( Result );
  if IsDelimiter( ',', Result, aEndPos ) then Delete( Result, aEndPos, 1 );
  { 最後回填 SOAC0502 時, 須要知道開了那些頻道 }
  if ( FChNotes = EmptyStr ) then FChNotes := Result;
end;

{ ---------------------------------------------------------------------------- }

function TDvnThread.IsProcessing(const aCmdType: Integer): Boolean;
var
  aCmdA1, aCmdB1: String;
begin
  Result := False;
  aCmdA1 := FDataSet.FieldByName( 'CMDA1' ).AsString;
  aCmdB1 := FDataSet.FieldByName( 'CMDB1' ).AsString;
  if ( aCmdType = CmdA1 ) then
    Result := ( ( aCmdA1 = 'W' ) or ( aCmdA1 = 'P' ) )
  else if ( aCmdType = CmdB1 ) and ( FWantToOpenChannel ) then
    Result := ( ( aCmdB1 = 'W' ) or ( aCmdB1 = 'P' ) )
end;

{ ---------------------------------------------------------------------------- }

procedure TDvnThread.CheckCmdStatus;

    { -------------------------------------- }

    procedure GetStatus(const aCmdType: Integer);
    var
      aCmdFieldName, aSeqFieldName: String;
    begin
      {}
      aCmdFieldName := EmptyStr;
      aSeqFieldName := EmptyStr;
      if ( aCmdType = CmdA1 ) then
      begin
        aCmdFieldName := 'CMDA1';
        aSeqFieldName := 'SEQNO1';
      end else
      if ( aCmdType = CmdB1 ) then
      begin
        aCmdFieldName := 'CMDB1';
        aSeqFieldName := 'SEQNO2';
      end;
      {}
      if ( aCmdFieldName = EmptyStr ) then Exit;
      {}
      FDataReader.Close;
      FDataReader.SQL.Text := Format(
        ' SELECT CMD_STATUS, ERR_MSG ' +
        '  FROM SEND_DVN             ' +
        ' WHERE SEQNO = ''%S''       ',
        [FDataSet.FieldByName( aSeqFieldName ).AsString ] );
      FDataReader.Open;
      {}
      if ( not FDataReader.IsEmpty ) then
      begin
        FDataSet.Edit;
        FDataSet.FieldByName( aCmdFieldName ).AsString :=
          FDataReader.FieldByName( 'CMD_STATUS' ).AsString;
        FDataSet.FieldByName( 'ERRMSG' ).AsString := EmptyStr;
        if ( FDataSet.FieldByName( aCmdFieldName ).AsString = 'E' ) then
        begin
          FDataSet.FieldByName( 'ERRMSG' ).AsString :=
            FDataReader.FieldByName( 'ERR_MSG' ).AsString;
        end;
        FDataSet.Post;
      end;
    end;

    { -------------------------------------- }

    function IsCmdTimeOut(const aCmdType: Integer): Boolean;
    var
      aSec: Integer;
      aCmdFieldName: String;
    begin
      Result := False;
      aCmdFieldName := EmptyStr;
      if ( aCmdType = CmdA1 ) then
        aCmdFieldName := 'SENDTIME1'
      else if ( aCmdType = CmdB1 ) then
        aCmdFieldName := 'SENDTIME2';
      {}
      if ( aCmdFieldName = EmptyStr ) then Exit;
      {}
      if ( IsProcessing( aCmdType ) ) then
      begin
       aSec := SecondsBetween( Now,
         FDataSet.FieldByName( aCmdFieldName ).AsDateTime );
       Result := ( aSec >= FAccTimeOut );
      end;
    end;

    { -------------------------------------- }

    procedure SetCmdStatusTimeOut(const aCmdType: Integer);
    var
      aCmdFieldName, aSeqFieldName: String;
    begin
      aCmdFieldName := EmptyStr;
      aSeqFieldName := EmptyStr;
      if ( aCmdType = CmdA1 ) then
      begin
        aCmdFieldName := 'CMDA1';
        aSeqFieldName := 'SEQNO1';
      end else
      if ( aCmdType = CmdB1 ) then
      begin
        aCmdFieldName := 'CMDB1';
        aSeqFieldName := 'SEQNO2';
      end;  
      {}
      if ( aCmdFieldName = EmptyStr ) then Exit;
      {}
      FDataSet.Edit;
      FDataSet.FieldByName( aCmdFieldName ).AsString := 'E';
      FDataSet.FieldByName( 'ERRMSG' ).AsString := Format(
        '指令處理逾時(%d)秒', [FAccTimeOut] );
      FDataSet.Post;
      { 更新 send_dvn table }
      FDataWriter.CommandText := Format(
        ' UPDATE SEND_DVN SET CMD_STATUS = ''E'',  ' +
        '    ERR_CODE = ''J01'',                   ' +
        '    ERR_MSG = ''作業逾時''                ' +
        '  WHERE SEQNO = ''%S''                    ',
        [FDataSet.FieldByName( aSeqFieldName ).AsString] );
      FDataWriter.Execute;  
    end;

    { -------------------------------------- }

    procedure SetGuiInfo(const aCmdType: Integer);
    var
      aCmdFieldName, aSeqFieldName: String;
    begin
      aCmdFieldName := EmptyStr;
      if ( aCmdType = CmdA1 ) then
      begin
        aCmdFieldName := 'CMDA1';
        aSeqFieldName := 'SEQNO1';
      end else
      if ( aCmdType = CmdB1 ) then
      begin
        aCmdFieldName := 'CMDB1';
        aSeqFieldName := 'SEQNO2';
      end;  
      {}
      if ( aCmdFieldName = EmptyStr ) then Exit;
      {}
      FGuiInfo.RecNo := FDataSet.FieldByName( 'RECNO' ).AsInteger;
      FGuiInfo.SeqNo := FDataSet.FieldByName( aSeqFieldName ).AsString;
      FGuiInfo.CmdType := aCmdType;
      FGuiInfo.CmdStatus := FDataSet.FieldByName( aCmdFieldName ).AsString;
      FGuiInfo.ErrMsg := FDataSet.FieldByName( 'ERRMSG' ).AsString;
    end;

    { -------------------------------------- }

var
  aIndex: Integer;
begin
  FDataSet.First;
  while not FDataSet.Eof do
  begin
    {}
    for aIndex := CmdA1 to CmdB1 do
    begin
      if ( IsProcessing( aIndex ) ) then
      begin
        GetStatus( aIndex );
        SetGuiInfo( aIndex );
        Synchronize( UpdateGUI );
      end;
    end;
    {}
    for aIndex := CmdA1 to CmdB1 do
    begin
      if ( IsCmdTimeOut( aIndex ) ) then
      begin
        SetCmdStatusTimeOut( aIndex );
        SetGuiInfo( aIndex );
        Synchronize( UpdateGUI );
      end;
    end;
    {}
    FDataSet.Next;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TDvnThread.CheckThreadCanbeBreak: Boolean;
begin
  Result := True;
  FDataSet.First;
  while not FDataSet.Eof do
  begin
    if IsProcessing( CmdA1 ) then
    begin
      Result := False;
      Break;
    end;
    {}
    if IsProcessing( CmdB1 ) then
    begin
      Result := False;
      Break;
    end;
    {}
    FDataSet.Next;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDvnThread.WriteToSendDvn(const aCmdType: Integer);
begin
  FDataSet.First;
  while not FDataSet.Eof do
  begin
    if ( aCmdType = CmdA1 ) then
    begin
      if ( FDataSet.FieldByName( 'CMDA1' ).AsString = EmptyStr ) then
        WriteA1;
    end;
    if ( aCmdType = CmdB1 ) and ( FWantToOpenChannel ) then
    begin
      if ( FDataSet.FieldByName( 'CMDA1' ).AsString = 'C' ) and
         ( FDataSet.FieldByName( 'CMDB1' ).AsString = EmptyStr ) then
        WriteB1;
    end;
    FDataSet.Next;
    Sleep( 50 );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDvnThread.WriteA1;
var
  aSeqNo: String;
begin
  aSeqNo := GetSeqNo;
  FDataWriter.CommandText := Format(
    ' insert into send_dvn ( seqno, high_level_cmd_id, icc_no, ' +
    '  stb_no, cmd_status, operator, updtime, compcode  )      ' +
    '  values ( ''%s'', ''%s'', ''%s'',             ' +
    '  ''%s'', ''W'', ''%s'', sysdate, ''%s'' )     ',
    [ aSeqNo, FCmdA1Text,
      FDataSet.FieldByName( 'ICCNO' ).AsString,
      FDataSet.FieldByName( 'STBNO' ).AsString,
      FSmsOperator, FCompCode] );
  FDataWriter.Execute;
  { 記錄傳送指令時間, 檢測指令逾時用 }
  FDataSet.Edit;
  FDataSet.FieldByName( 'SEQNO1' ).AsString := aSeqNo;
  FDataSet.FieldByName( 'CMDA1' ).AsString := 'W';
  FDataSet.FieldByName( 'SENDTIME1' ).AsDateTime := Now;
  FDataSet.Post;
  {}
  FGuiInfo.RecNo := FDataSet.FieldByName( 'RECNO' ).AsInteger;
  FGuiInfo.SeqNo := aSeqNo;
  FGuiInfo.CmdType := CmdA1;
  FGuiInfo.CmdStatus := 'W';
  FGuiInfo.ErrMsg := EmptyStr;
  Synchronize( UpdateGUI );    
end;

{ ---------------------------------------------------------------------------- }

procedure TDvnThread.WriteB1;
var
  aSeqNo, aNotes: String;
begin
  aNotes := BuildNotes;
  if ( aNotes = EmptyStr ) then Exit;  
  aSeqNo := GetSeqNo;
  FDataWriter.CommandText := Format(
    ' insert into send_dvn ( seqno, high_level_cmd_id, icc_no,   ' +
    '  stb_no, cmd_status, operator, updtime, compcode, notes  ) ' +
    '  values ( ''%s'', ''B1'', ''%s'',                  ' +
    '  ''%s'', ''W'', ''%s'', sysdate, ''%s'', ''%s'' )  ',
    [ aSeqNo,
      FDataSet.FieldByName( 'ICCNO' ).AsString,
      FDataSet.FieldByName( 'STBNO' ).AsString,
      FSmsOperator, FCompCode, aNotes] );
  FDataWriter.Execute;
  FDataSet.Edit;
  { 記錄傳送指令時間, 檢測指令逾時用 }
  FDataSet.FieldByName( 'SEQNO2' ).AsString := aSeqNo;
  FDataSet.FieldByName( 'CMDB1' ).AsString := 'W';
  FDataSet.FieldByName( 'SENDTIME2' ).AsDateTime := Now;
  FDataSet.Post;
  {}
  FGuiInfo.RecNo := FDataSet.FieldByName( 'RECNO' ).AsInteger;
  FGuiInfo.SeqNo := aSeqNo;
  FGuiInfo.CmdType := CmdB1;
  FGuiInfo.CmdStatus := 'W';
  FGuiInfo.ErrMsg := EmptyStr;
  Synchronize( UpdateGUI );
end;

{ ---------------------------------------------------------------------------- }

end.
 