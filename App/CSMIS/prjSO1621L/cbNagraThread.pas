unit cbNagraThread;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, DB, DBClient, ADODB,
  cbMain, cbDataController;


type
  TNagraThread = class(TThread)
  private
    { Private declarations }
    FDataSet: TClientDataSet;
    FConnection: TADOConnection;
    FChannelDataSet: TClientDataSet;
    FDataReader: TADOQuery;
    FDataWriter: TADOCommand;
    FCompCode: String;
    FSmsOperator: String;
    FBatId: String;
    FZipCode: String;
    FResendErrCmd: Boolean;
    FOracleToday: TDateTime;
    FAccTimeOut: Integer;
    FSendNagraOwner: String;
    FWantToOpenChannel: Boolean;
    FGuiInfo: TGuiInfo;
    FCmds: array [1..4] of String;
    FCmdTexts: array [1..4] of String;
    FChNotes: String;
    FNagraOperation: Integer;
    FNagraStep: Integer;
    FAutoFill: Boolean;
    function GetOracleToday: TDateTime;
    function GetAccTimeOut: Integer;
    function GetSendNagraTableOwner: String;
    function GetWantToOpenChannel: Boolean;
    function GetSeqNo: String;
    function BuildNotes: String;
    function IsProcessing(const ACmdIndex: Integer): Boolean;
    function CheckThreadCanbeBreak: Boolean;
    procedure SetSourceDataSet(const AValue: TClientDataSet);
    procedure SetChannelDataSet(const AValue: TClientDataSet);
    procedure ChangeErrCmdStatus;
    procedure CheckCmdStatus;
    procedure MarkHasBeenLog;
    procedure MarkSkipFlag;
    procedure UpdateGUI;
    procedure UpdateAllData;
    procedure WriteToSendNagra;
    procedure WriteA1(const ACmdIndex: Integer);
    procedure WriteB1(const ACmdIndex: Integer);
    procedure WriteB3(const ACmdIndex: Integer);
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
    property BatId: String read FBatId write FBatId;
    property ZipCode: String read FZipCode write FZipCode;
    property ResendErrCmd: Boolean read FResendErrCmd write FResendErrCmd;
    property NagraOperation: Integer read FNagraOperation write FNagraOperation;
    property NagraStep: Integer read FNagraStep write FNagraStep;
    property AutoFill:Boolean write FAutoFill;

  end;

var NagraThreadErrMsg: String;

implementation

uses cbUtilis, DateUtils;

{ Important: Methods and properties of objects in visual components can only be
  used in a method called using Synchronize, for example,

      Synchronize(UpdateCaption);

  and UpdateCaption could look like,

    procedure TNagraThread.UpdateCaption;
    begin
      Form1.Caption := 'Updated in a thread';
    end; }

{ TNagraThread }

{ ---------------------------------------------------------------------------- }

constructor TNagraThread.Create(CreateSuspended: Boolean);
begin
  inherited Create( CreateSuspended );
  FDataSet := TClientDataSet.Create( nil );
  FChannelDataSet := TClientDataSet.Create( nil );
  FDataReader := TADOQuery.Create( nil );
  FDataWriter := TADOCommand.Create( nil );
  FResendErrCmd := False;
  FNagraStep := 1;
  FWantToOpenChannel := False;
  {}
  FCmds[1] := 'CMDA1';
  FCmds[2] := 'CMDB1';
  FCmds[3] := 'CMDB3';
  FCmds[4] := 'CMDB1_1';
  {}
  FCmdTexts[1] := 'A1';
  FCmdTexts[2] := 'B1';
  FCmdTexts[3] := 'B3';
  FCmdTexts[4] := 'B1';
end;

{ ---------------------------------------------------------------------------- }

destructor TNagraThread.Destroy;
begin
  FDataSet.Free;
  FChannelDataSet.Free;
  FDataReader.Free;
  FDataWriter.Free;
  FConnection := nil;
  inherited Destroy;
end;

{ ---------------------------------------------------------------------------- }

function TNagraThread.GetOracleToday: TDateTime;
begin
  FDataReader.Close;
  FDataReader.SQL.Text := ' SELECT SYSDATE FROM DUAL ';
  FDataReader.Open;
  Result := FDataReader.Fields[0].AsDateTime;
  FDataReader.Close;
end;

{ ---------------------------------------------------------------------------- }

function TNagraThread.GetAccTimeOut: Integer;
begin
  FDataReader.Close;
  FDataReader.SQL.Text := 'select acctimeout from so041 ';
  FDataReader.Open;
  Result := FDataReader.Fields[0].AsInteger;
  if ( Result <= 0 ) then Result := 90;
  FDataReader.Close;
end;

{ ---------------------------------------------------------------------------- }

function TNagraThread.GetSendNagraTableOwner: String;
begin
  FDataReader.Close;
  FDataReader.SQL.Text := 'select loginid from so041 ';
  FDataReader.Open;
  Result := FDataReader.Fields[0].AsString;
  FDataReader.Close;
end;

{ ---------------------------------------------------------------------------- }

function TNagraThread.GetWantToOpenChannel: Boolean;
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

function TNagraThread.GetSeqNo: String;
begin
  FDataReader.Close;
  FDataReader.SQL.Text := Format(
    ' SELECT %s.S_SENDNAGRA_SEQNO.NEXTVAL FROM DUAL ',
    [Nvl( FSendNagraOwner, DBController.LoginInfo.DbAccount )] );
  FDataReader.Open;
  Result := FDataReader.Fields[0].AsString;
  FDataReader.Open;
end;

{ ---------------------------------------------------------------------------- }

function TNagraThread.BuildNotes: String;
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
        StrToIntDef(FChannelDataSet.FieldByName( 'CHANCEDAYS' ).AsString,0) +
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
  { 最後回填 SOAC0202 時, 需要知道開了那些頻道 }
  if ( FChNotes = EmptyStr ) then FChNotes := Result;
end;

{ ---------------------------------------------------------------------------- }

function TNagraThread.IsProcessing(const ACmdIndex: Integer): Boolean;
var
  aCmdStatus: String;
begin
  aCmdStatus := FDataSet.FieldByName( FCmds[ACmdIndex] ).AsString;
  Result := ( ( aCmdStatus = 'W' ) or ( aCmdStatus = 'P' ) );
end;

{ ---------------------------------------------------------------------------- }

function TNagraThread.CheckThreadCanbeBreak: Boolean;
var
  aIndex: Integer;
begin
  Result := True;
  FDataSet.First;
  while not FDataSet.Eof do
  begin
    for aIndex := Low( FCmds ) to High( FCmds ) do
    begin
      if IsProcessing( aIndex ) then
      begin
        Result := False;
        Break;
      end;
    end;
    FDataSet.Next;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TNagraThread.SetSourceDataSet(const AValue: TClientDataSet);
begin
  FDataSet.Data := AValue.Data;
end;

{ ---------------------------------------------------------------------------- }

procedure TNagraThread.SetChannelDataSet(const AValue: TClientDataSet);
begin
  FChannelDataSet.Data := AValue.Data;
end;

{ ---------------------------------------------------------------------------- }

procedure TNagraThread.ChangeErrCmdStatus;
var
  aIndex: Integer;
begin
  FDataSet.First;
  while not FDataSet.Eof do
  begin
    FDataSet.Edit;
    for aIndex := Low( FCmds ) to High( FCmds ) do
    begin
      if ( FDataSet.FieldByName( FCmds[aIndex] ).AsString = 'E' ) then
      begin
        FDataSet.FieldByName( FCmds[aIndex] ).AsString := EmptyStr;
        FGuiInfo.RecNo := FDataSet.FieldByName( 'RECNO' ).AsInteger;
        FGuiInfo.SeqNo := FDataSet.FieldByName( Format( 'SEQNO%d', [aIndex] ) ).AsString;
        FGuiInfo.CmdType := aIndex;
        FGuiInfo.CmdStatus := EmptyStr;
        FGuiInfo.ErrMsg := EmptyStr;
        Synchronize( UpdateGUI );
      end;
    end;
    FDataSet.Post;
    FDataSet.Next;
    Sleep( 10 );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TNagraThread.CheckCmdStatus;

    { -------------------------------------- }

    procedure GetStatus(const ACmdIndex: Integer);
    begin
      FDataReader.Close;
      FDataReader.SQL.Text := Format(
        ' SELECT CMD_STATUS, ERR_MSG   ' +
        '  FROM %s.SEND_NAGRA          ' +
        ' WHERE COMPCODE = ''%S''      ' +
        '   AND SEQNO = ''%s''         ',
        [Nvl( FSendNagraOwner, DBController.LoginInfo.DbAccount ),
         FCompCode,
         FDataSet.FieldByName( Format( 'SEQNO%d', [ACmdIndex] ) ).AsString] );
      FDataReader.Open;
      if ( not FDataReader.IsEmpty ) then
      begin
        FDataSet.Edit;
        FDataSet.FieldByName( FCmds[ACmdIndex] ).AsString :=
          FDataReader.FieldByName( 'CMD_STATUS' ).AsString;
        FDataSet.FieldByName( 'ERRMSG' ).AsString := EmptyStr;
        if ( FDataSet.FieldByName( FCmds[ACmdIndex] ).AsString = 'E' ) then
        begin
          FDataSet.FieldByName( 'ERRMSG' ).AsString :=
            FDataReader.FieldByName( 'ERR_MSG' ).AsString;
        end;
        FDataSet.Post;
      end;
    end;

    { -------------------------------------- }

    function IsCmdTimeOut(const ACmdIndex: Integer): Boolean;
    var
      aSec: Integer;
    begin
      Result := False;
      if ( IsProcessing( ACmdIndex ) ) then
      begin
       aSec := SecondsBetween( Now,
         FDataSet.FieldByName( Format( 'SENDTIME%d', [ACmdIndex] ) ).AsDateTime );
       Result := ( aSec >= FAccTimeOut );
      end;
    end;

    { -------------------------------------- }

    procedure SetCmdStatusTimeOut(const ACmdIndex: Integer);
    begin
      FDataSet.Edit;
      FDataSet.FieldByName( FCmds[ACmdIndex] ).AsString := 'E';
      FDataSet.FieldByName( 'ERRMSG' ).AsString := Format(
        '指令處理逾時(%d)秒', [FAccTimeOut] );
      FDataSet.Post;
      { 更新 send_nagra table }
      FDataWriter.CommandText := Format(
        ' UPDATE %s.SEND_NAGRA SET CMD_STATUS = ''E'',  ' +
        '    ERR_CODE = ''J01'',                        ' +
        '    ERR_MSG = ''作業逾時''                     ' +
        '  WHERE COMPCODE = ''%s''                      ' +
        '    AND SEQNO = ''%S''                         ',
        [Nvl( FSendNagraOwner, DBController.LoginInfo.DbAccount ),
         FCompCode,
         FDataSet.FieldByName( Format( 'SEQNO%d', [ACmdIndex]) ).AsString] );
      FDataWriter.Execute;
    end;

    { -------------------------------------- }
    
    procedure SetGuiInfo(const ACmdIndex: Integer);
    begin
      FGuiInfo.RecNo := FDataSet.FieldByName( 'RECNO' ).AsInteger;
      FGuiInfo.SeqNo := FDataSet.FieldByName( Format( 'SEQNO%d', [ACmdIndex]) ).AsString;
      FGuiInfo.CmdType := ACmdIndex;
      FGuiInfo.CmdStatus := FDataSet.FieldByName( FCmds[ACmdIndex] ).AsString;
      FGuiInfo.ErrMsg := FDataSet.FieldByName( 'ERRMSG' ).AsString;
    end;

    { -------------------------------------- }

var
  aIndex: Integer;
begin
  FDataSet.First;
  while not FDataSet.Eof do
  begin
    for aIndex := Low( FCmds ) to High( FCmds ) do
    begin
      if ( IsProcessing( aIndex ) ) then
      begin
        GetStatus( aIndex );
        SetGuiInfo( aIndex );
        Synchronize( UpdateGUI );
      end;
      if ( aIndex >= FNagraStep ) then Break;
    end;
    for aIndex := Low( FCmds ) to High( FCmds ) do
    begin
      if ( IsCmdTimeOut( aIndex ) ) then
      begin
        SetCmdStatusTimeOut( aIndex );
        SetGuiInfo( aIndex );
        Synchronize( UpdateGUI );
      end;
      if ( aIndex >= FNagraStep ) then Break;
    end;
    FDataSet.Next;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TNagraThread.MarkHasBeenLog;
var
  aIndex: Integer;
begin
  FDataSet.First;
  while not FDataSet.Eof do
  begin
    FDataSet.Edit;
    for aIndex := Low( FCmds ) to High( FCmds ) do
    begin
      if ( FDataSet.FieldByName( FCmds[aIndex] ).AsString = 'C' ) then
      begin
        FDataSet.FieldByName( Format( 'HASLOG%d', [aIndex] ) ).AsString := 'Y';
      end;
    end;
    FDataSet.Post;
    FDataSet.Next;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TNagraThread.MarkSkipFlag;
begin
  FDataSet.First;
  while not FDataSet.Eof do
  begin
    if ( FNagraOperation in [2,4] ) and ( not FWantToOpenChannel ) then
    begin
      FDataSet.Edit;
      FDataSet.FieldByName( FCmds[2] ).AsString := 'S';
      FDataSet.FieldByName( FCmds[4] ).AsString := 'S';
      FDataSet.Post;
    end;
    FDataSet.Next;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TNagraThread.UpdateGUI;
begin
  fmMain.ProgressReport( FGuiInfo.CmdType, FGuiInfo.RecNo,
    FGuiInfo.CmdStatus, FGuiInfo.ErrMsg );
end;

{ ---------------------------------------------------------------------------- }

procedure TNagraThread.UpdateAllData;

  { ------------------------------------- }

  procedure GetChInfo(const ACodeNo: String; var AChId, AChName: String);
  begin
    aChName := EmptyStr;
    aChId := EmptyStr;
    FChannelDataSet.First;
    if ( FChannelDataSet.Locate( 'CHANNELID', ACodeNo, [] ) ) then
    begin
      aChId := FChannelDataSet.FieldByName( 'CODENO' ).AsString;
      aChName := FChannelDataSet.FieldByName( 'DESCRIPTION' ).AsString;
    end;
  end;

  { ------------------------------------- }

  procedure DeleteSuccessCmds;
  var
    aIndex: Integer;
    aSeqNo: String;
  begin
    for aIndex := Low( FCmds ) to High( FCmds ) do
    begin
      if ( FDataSet.FieldByName( FCmds[aIndex] ).AsString = 'C' ) then
      begin
        aSeqNo := FDataSet.FieldByName( Format( 'SEQNO%d', [aIndex] ) ).AsString;
        if ( aSeqNo <> EmptyStr ) then
        begin
          FDataWriter.CommandText := Format(
            ' DELETE FROM %s.SEND_NAGRA   ' +
            '  WHERE COMPCODE = ''%s''    ' +
            '    AND SEQNO = ''%s''       ',
            [Nvl( FSendNagraOwner, DBController.LoginInfo.DbAccount ),
             FCompCode,
             aSeqNo] );
          FDataWriter.Execute;
        end;
      end;
    end;
  end;

  { ------------------------------------- }

  procedure UpdateSOAC0201C;
  var
    aRecordEffect: Integer;
  begin
    //#5216 如果是CNS則不要進行Upd或Insert By Kin 2009/07/30
    if not FAutoFill then
    begin
      FDataWriter.CommandText := Format(
        ' UPDATE SOAC0201C SET SMARTCARDNO = ''%s'',   ' +
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
          ' INSERT INTO SOAC0201C ( COMPCODE, STBSNO,   ' +
          '   SMARTCARDNO, LASTPAIRDATE )               ' +
          '  VALUES ( ''%s'', ''%s'', ''%s'', SYSDATE ) ',
          [FCompCode,
           FDataSet.FieldByName( 'STBNO' ).AsString,
           FDataSet.FieldByName( 'ICCNO' ).AsString] );
        FDataWriter.Execute;
      end;
    end;
  end;

  { ------------------------------------- }

  procedure UpdateSOAC0202;
  var
    aUpdTime, aCmdText: String;
    aChId, aOldNotes, aChText, aChCode, aChName, aStarDate, aExpiryDate: String;
    aIndex: Integer;
  begin
    aUpdTime := DateConvert( FOracleToday, True ) + ' ' +
      FormatDateTime( 'hh:nn:ss', FOracleToday );
    if ( Copy( aUpdTime, 1, 1 ) = #32 ) then Delete( aUpdTime, 1, 1 );
    {}
    for aIndex := Low( FCmds ) to High( FCmds ) do
    begin
      if ( aIndex in [1,3] ) and
         ( FDataSet.FieldByName( FCmds[aIndex] ).AsString = 'C' ) and
         ( FDataSet.FieldByName( Format( 'HASLOG%d', [aIndex] ) ).AsString <> 'Y'  ) then
      begin
        aCmdText := FCmdTexts[aIndex];
        FDataWriter.CommandText := Format(
          ' INSERT INTO SOAC0202 ( COMPCODE, STBSNO, SMARTCARDNO, ' +
          '    MODETYPE, UPDTIME, UPDEN )                         ' +
          ' VALUES ( ''%S'', ''%S'', ''%S'', ''%S'', ''%S'',      ' +
          '    ''%S'' ) ',
          [FCompCode,
           FDataSet.FieldByName( 'STBNO' ).AsString,
           FDataSet.FieldByName( 'ICCNO' ).AsString,
           aCmdText, aUpdTime, FSmsOperator] );
        FDataWriter.Execute;
      end;
      {}
      if ( aIndex in [2,4] ) and
         ( FWantToOpenChannel ) and
         ( FDataSet.FieldByName( FCmds[aIndex] ).AsString = 'C' ) and
         ( FDataSet.FieldByName( Format( 'HASLOG%d', [aIndex] ) ).AsString <> 'Y' ) then
      begin
        aCmdText := FCmdTexts[aIndex];
        //試看頻道的指令代碼, 改成 BA
        if ( SameText( aCmdText, 'B1' ) ) then aCmdText := 'BA';
        {}
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
          GetChInfo( aChCode, aChId, aChName );
          //
          FDataWriter.CommandText := Format(
            ' INSERT INTO SOAC0202 ( COMPCODE, STBSNO, SMARTCARDNO, ' +
            '    MODETYPE, CHCODE, CHNAME, UPDTIME, UPDEN,          ' +
            '    AUTHORSTARTDATE, AUTHORSTOPDATE )                  ' +
            ' VALUES ( ''%S'', ''%S'', ''%S'', ''%S'', ''%S'',      ' +
            '          ''%S'', ''%S'', ''%S'',                      ' +
            '          TO_DATE( ''%S'', ''YYYYMMDD'' ),             ' +
            '          TO_DATE( ''%S'', ''YYYYMMDD'' ) )            ',
            [FCompCode,
             FDataSet.FieldByName( 'STBNO' ).AsString,
             FDataSet.FieldByName( 'ICCNO' ).AsString,
             aCmdText, aChId, aChName, aUpdTime, FSmsOperator, aStarDate, aExpiryDate] );
          FDataWriter.Execute;
        until ( aOldNotes = EmptyStr );
      end;
    end;
  end;

  { ------------------------------------- }

  procedure MoveToSO005B;
  begin
    FDataWriter.CommandText := Format(
      ' INSERT INTO SO005B ( CUSTID, CVTID, CHCODE, STATUS, ' +
      '   SETTIME, CITEMCODE, CVTSNO, DUEDATE, SETEN,       ' +
      '   SETNAME, SERVICETYPE, COMPCODE, SMARTCARDNO,      ' +
      '   ORDERNO, CLOSETIME  )                             ' +
      ' SELECT CUSTID, CVTID, CHCODE, STATUS,               ' +
      '   SETTIME, CITEMCODE, CVTSNO, DUEDATE, SETEN,       ' +
      '   SETNAME, SERVICETYPE, COMPCODE, SMARTCARDNO,      ' +
      '   ORDERNO, SYSDATE                                  ' +
      '  FROM SO005                                         ' +
      ' WHERE CVTSNO = ''%s''                               ',
      [FDataSet.FieldByName( 'STBNO' ).AsString] );
    FDataWriter.Execute;
    {}
    FDataWriter.CommandText := Format(
      ' DELETE FROM SO005 WHERE CVTSNO = ''%s'' ',
      [FDataSet.FieldByName( 'STBNO' ).AsString] );
    FDataWriter.Execute;  
  end;

  { ------------------------------------- }

begin
  FDataSet.First;
  while not FDataSet.Eof do
  begin
    DeleteSuccessCmds;
    if ( FDataSet.FieldByName( FCmds[1] ).AsString = 'C' ) then
      UpdateSOAC0201C;
    {}
    UpdateSOAC0202;
    {}
    if ( FDataSet.FieldByName( FCmds[3] ).AsString = 'C' ) then
      MoveToSO005B;
    {}
    FDataSet.Next;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TNagraThread.WriteToSendNagra;

  { ------------------------------------- }

  function VdWriteConditions(ACmdIndex: Integer; const ACmdStatus: String): Boolean;
  begin
    Result := ( ACmdStatus = EmptyStr );
    if ( Result ) then
    begin
      while ( ACmdIndex > 1 ) do
      begin
        Dec( ACmdIndex );
        Result :=
          ( ( Result ) and
            ( FDataSet.FieldByName( FCmds[ACmdIndex] ).AsString = 'C' ) or
            ( FDataSet.FieldByName( FCmds[ACmdIndex] ).AsString = 'S' ) );
        if not Result then Break;
      end;
    end;
  end;

  { ------------------------------------- }

  procedure WriteCommand(const ACmdIndex: Integer);
  begin
    case ACmdIndex of
      1: WriteA1( ACmdIndex );
      2: WriteB1( ACmdIndex );
      3: WriteB3( ACmdIndex );
      4: WriteB1( ACmdIndex );
    end;
  end;

  { ------------------------------------- }

var
  aIndex: Integer;
  aCmdStatus: String;
begin
  FDataSet.First;
  while not FDataSet.Eof do
  begin
    for aIndex := Low( FCmds ) to High( FCmds ) do
    begin
      aCmdStatus := FDataSet.FieldByName( FCmds[aIndex] ).AsString;
      { S -> Skip 跳過不處理 }
      if VdWriteConditions( aIndex, aCmdStatus ) then
        WriteCommand( aIndex );
      if ( aIndex = FNagraStep ) then Break;
    end;
    Sleep( 10 );
    FDataSet.Next;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TNagraThread.WriteA1(const ACmdIndex: Integer);
var
  aSeqNo: String;
begin
  aSeqNo := GetSeqNo;
  FDataWriter.CommandText := Format(
    ' insert into %s.send_nagra (                        ' +
    '   high_level_cmd_id, icc_no, stb_no, zip_code,     ' +
    '   mis_ird_cmd_data, pincode, cmd_status,           ' +
    '   operator, updtime, seqno, compcode, resenttimes, ' +
    '   stb_flag )                                       ' +
    ' values ( :1, :2, :3, :4, :5, :6, :7, :8,           ' +
    '   sysdate, :9, :10, :11, :12 )                     ',
    [Nvl( FSendNagraOwner, DBController.LoginInfo.DbAccount )] );
  {}  
  FDataWriter.Parameters.ParamByName( '1' ).Value := FCmdTexts[ACmdIndex];
  FDataWriter.Parameters.ParamByName( '2' ).Value := FDataSet.FieldByName( 'ICCNO' ).AsString;
  FDataWriter.Parameters.ParamByName( '3' ).Value := FDataSet.FieldByName( 'STBNO' ).AsString;
  FDataWriter.Parameters.ParamByName( '4' ).Value := FZipCode;
  FDataWriter.Parameters.ParamByName( '5' ).Value := FBatId;
  FDataWriter.Parameters.ParamByName( '6' ).Value := '0000';
  FDataWriter.Parameters.ParamByName( '7' ).Value := 'W';
  FDataWriter.Parameters.ParamByName( '8' ).Value := FSmsOperator;
  FDataWriter.Parameters.ParamByName( '9' ).Value := aSeqNo;
  FDataWriter.Parameters.ParamByName( '10' ).Value := FCompCode;
  FDataWriter.Parameters.ParamByName( '11' ).Value := '0';
  FDataWriter.Parameters.ParamByName( '12' ).Value := '1';
  {}
  FDataWriter.Execute;
  { 記錄傳送指令時間, 檢測指令逾時用 }
  FDataSet.Edit;
  FDataSet.FieldByName( Format( 'SEQNO%d', [ACmdIndex] ) ).AsString := aSeqNo;
  FDataSet.FieldByName( FCmds[ACmdIndex] ).AsString := 'W';
  FDataSet.FieldByName( Format( 'SENDTIME%d', [ACmdIndex] ) ).AsDateTime := Now;
  FDataSet.Post;
  {}
  FGuiInfo.RecNo := FDataSet.FieldByName( 'RECNO' ).AsInteger;
  FGuiInfo.SeqNo := aSeqNo;
  FGuiInfo.CmdType := ACmdIndex;
  FGuiInfo.CmdStatus := 'W';
  FGuiInfo.ErrMsg := EmptyStr;
  Synchronize( UpdateGUI );
end;

{ ---------------------------------------------------------------------------- }

procedure TNagraThread.WriteB1(const ACmdIndex: Integer);
var
  aSeqNo, aNotes: String;
begin
  aNotes := BuildNotes;
  if ( not FWantToOpenChannel ) and ( aNotes = EmptyStr ) then Exit;
  aSeqNo := GetSeqNo;
  FDataWriter.CommandText := Format(
    ' insert into %s.send_nagra (                  ' +
    '   high_level_cmd_id, icc_no, stb_no,         ' +
    '   notes, cmd_status, operator, updtime,      ' +
    '   seqno, compcode, resenttimes, stb_flag )   ' +
    ' values ( :1, :2, :3, :4, :5, :6, sysdate,    ' +
    '   :7, :8, :9, :10 )                          ',
    [Nvl( FSendNagraOwner, DBController.LoginInfo.DbAccount )] );
  {}
  FDataWriter.Parameters.ParamByName( '1' ).Value := FCmdTexts[ACmdIndex];
  FDataWriter.Parameters.ParamByName( '2' ).Value := FDataSet.FieldByName( 'ICCNO' ).AsString;
  FDataWriter.Parameters.ParamByName( '3' ).Value := FDataSet.FieldByName( 'STBNO' ).AsString;
  FDataWriter.Parameters.ParamByName( '4' ).Value := aNotes;
  FDataWriter.Parameters.ParamByName( '5' ).Value := 'W';
  FDataWriter.Parameters.ParamByName( '6' ).Value := FSmsOperator;
  FDataWriter.Parameters.ParamByName( '7' ).Value := aSeqNo;
  FDataWriter.Parameters.ParamByName( '8' ).Value := FCompCode;
  FDataWriter.Parameters.ParamByName( '9' ).Value := '0';
  FDataWriter.Parameters.ParamByName( '10' ).Value := '1';
  {}
  FDataWriter.Execute;
  {}
  FDataSet.Edit;
  { 記錄傳送指令時間, 檢測指令逾時用 }
  FDataSet.FieldByName( FCmds[ACmdIndex] ).AsString := 'W';
  FDataSet.FieldByName( Format( 'SEQNO%d', [ACmdIndex] ) ).AsString := aSeqNo;
  FDataSet.FieldByName( Format( 'SENDTIME%d', [ACmdIndex] ) ).AsDateTime := Now;
  FDataSet.Post;
  {}
  FGuiInfo.RecNo := FDataSet.FieldByName( 'RECNO' ).AsInteger;
  FGuiInfo.SeqNo := aSeqNo;
  FGuiInfo.CmdType := ACmdIndex;
  FGuiInfo.CmdStatus := FDataSet.FieldByName( FCmds[ACmdIndex] ).AsString;
  FGuiInfo.ErrMsg := EmptyStr;
  Synchronize( UpdateGUI );
end;

{ ---------------------------------------------------------------------------- }

procedure TNagraThread.WriteB3(const ACmdIndex: Integer);
var
  aSeqNo: String;
begin
  aSeqNo := GetSeqNo;
  FDataWriter.CommandText := Format(
    ' insert into %s.send_nagra (                  ' +
    '   high_level_cmd_id, icc_no, stb_no,         ' +
    '   cmd_status, operator, updtime,             ' +
    '   seqno, compcode, resenttimes, stb_flag )   ' +
    ' values ( :1, :2, :3, :4, :5, sysdate,        ' +
    '   :6, :7, :8, :9 )                           ',
    [Nvl( FSendNagraOwner, DBController.LoginInfo.DbAccount )] );
  {}  
  FDataWriter.Parameters.ParamByName( '1' ).Value := FCmdTexts[ACmdIndex];
  FDataWriter.Parameters.ParamByName( '2' ).Value := FDataSet.FieldByName( 'ICCNO' ).AsString;
  FDataWriter.Parameters.ParamByName( '3' ).Value := FDataSet.FieldByName( 'STBNO' ).AsString;
  FDataWriter.Parameters.ParamByName( '4' ).Value := 'W';
  FDataWriter.Parameters.ParamByName( '5' ).Value := FSmsOperator;
  FDataWriter.Parameters.ParamByName( '6' ).Value := aSeqNo;
  FDataWriter.Parameters.ParamByName( '7' ).Value := FCompCode;
  FDataWriter.Parameters.ParamByName( '8' ).Value := '0';
  FDataWriter.Parameters.ParamByName( '9' ).Value := '1';
  {}
  FDataWriter.Execute;
  { 記錄傳送指令時間, 檢測指令逾時用 }
  FDataSet.Edit;
  FDataSet.FieldByName( Format( 'SEQNO%d', [ACmdIndex] ) ).AsString := aSeqNo;
  FDataSet.FieldByName( FCmds[ACmdIndex] ).AsString := 'W';
  FDataSet.FieldByName( Format( 'SENDTIME%d', [ACmdIndex] ) ).AsDateTime := Now;
  FDataSet.Post;
  {}
  FGuiInfo.RecNo := FDataSet.FieldByName( 'RECNO' ).AsInteger;
  FGuiInfo.SeqNo := aSeqNo;
  FGuiInfo.CmdType := ACmdIndex;
  FGuiInfo.CmdStatus := 'W';
  FGuiInfo.ErrMsg := EmptyStr;
  Synchronize( UpdateGUI );
end;

{ ---------------------------------------------------------------------------- }

procedure TNagraThread.Execute;
begin
  FDataReader.Connection := FConnection;
  FDataWriter.Connection := FConnection;
  Sleep( 100 );
  FOracleToday := GetOracleToday;
  FAccTimeOut := GetAccTimeOut;
  FSendNagraOwner := GetSendNagraTableOwner;
  FWantToOpenChannel := GetWantToOpenChannel;
  Sleep( 100 );
  MarkHasBeenLog;
  MarkSkipFlag;
  if ( FResendErrCmd ) then ChangeErrCmdStatus;
  {}
  while ( not Self.Terminated ) do
  begin
    try
      WriteToSendNagra;
      Sleep( 300 );
      if Self.Terminated then Break;
      CheckCmdStatus;
      Sleep( 300 );
      if Self.Terminated then Break;
      if CheckThreadCanbeBreak then Break;
      Sleep( 300 );
    except
      on E: Exception do
      begin
        NagraThreadErrMsg := E.Message;
        Break;
      end;
    end;
  end;
  UpdateAllData;
  FDataReader.Close;
end;

{ ---------------------------------------------------------------------------- }

end.
