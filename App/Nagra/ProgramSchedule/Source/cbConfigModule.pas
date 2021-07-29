unit cbConfigModule;

interface

uses
  SysUtils, Classes, Variants, DB, ADODB, DBClient, cbClass, cbAppClass;

type
  TConfigModule = class(TDataModule)
    AccessConnection: TADOConnection;
    DataReader: TADOQuery;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
    FMsgSubject: TMessageSubject;
    FParam: TAppParam;
    FPeriod: TAppPeriod;
    FDatabase: TAppDatabase;
    FNotify: TAppNotify;
    function ConnectToAccess(aSilent: Boolean = False): Boolean;
    procedure DisconnectFromAccess;
    function GetParams: Boolean;
    function GetPeriod: Boolean;
    function GetNotify: Boolean;
    procedure SetParams;
    procedure SetPeriod;
    procedure SetNotify;
  public
    { Public declarations }
    function LoadFromFile(aSilent: Boolean = False): Boolean;
    function SaveToFile(aSilent: Boolean = False): Boolean;
    property MessageSubject: TMessageSubject read FMsgSubject;
    property Param: TAppParam read FParam;
    property Period: TAppPeriod read FPeriod;
    property Database: TAppDatabase read FDatabase;
    property Notify: TAppNotify read FNotify;
  end;

var
  ConfigModule: TConfigModule;

implementation

uses cbUtilis;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.DataModuleCreate(Sender: TObject);
begin
  FMsgSubject := TMessageSubject.Create;
  FParam := TAppParam.Create;
  FPeriod := TAppPeriod.Create;
  FDatabase := TAppDatabase.Create;
  FNotify := TAppNotify.Create;
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.DataModuleDestroy(Sender: TObject);
begin
  FParam.Free;
  FPeriod.Free;
  FDatabase.Free;
  FNotify.Free;
  FMsgSubject.Free;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.ConnectToAccess(aSilent: Boolean = False): Boolean;
const
  aDbPassword = 'cyc84177282';
  aDbFileName = 'Config.cfg';
var
  aFileName: String;
begin
  aFileName := IncludeTrailingPathDelimiter( ExtractFilePath( ParamStr( 0 ) ) ) +
    aDbFileName;
  AccessConnection.Connected := False;
  AccessConnection.ConnectionString := Format(
    'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=%s;Jet OLEDB:Database Password=%s;'+
    'Mode=Share Deny Read|Share Deny Write;', [aFileName, aDbPassword] );
  try
    if not FileExists( aFileName ) then
      raise Exception.CreateFmt( '砞﹚郎%sぃ', [aDbFileName] );
    AccessConnection.Connected := True;
    if not aSilent then
      FMsgSubject.OK( Format( '砞﹚郎%s更ЧΘ', [aDbFileName] ) );
  except
    on E: Exception do
    begin
      if not aSilent then
        FMsgSubject.Error( Format( '砞﹚郎%s更ア毖, :%s',
          [aDbFileName, E.Message] ) );
    end;
  end;
  Result := AccessConnection.Connected;
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.DisconnectFromAccess;
begin
  if AccessConnection.Connected then
    AccessConnection.Connected := False;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.LoadFromFile(aSilent: Boolean = False): Boolean;
begin
  Result := ConnectToAccess( aSilent );
  if Result then
  begin
    try
      Delay( COMM_DELAY * 5 );
      if not aSilent then
      begin
        FMsgSubject.Normal( '弄把计Paramsい' );
        Delay( COMM_DELAY * 5 );
      end;
      Result := GetParams;
      if not Result then Exit;
      if not aSilent then
      begin
        FMsgSubject.Normal( '弄把计Periodい' );
        Delay( COMM_DELAY * 5 );
      end;
      Result := GetPeriod;
      if not Result then Exit;
      if not aSilent then
      begin
        FMsgSubject.Normal( '弄把计Notifyい' );
        Delay( COMM_DELAY * 5 );
      end;
      Result := GetNotify;
      if not Result then Exit;
      if not aSilent then
      begin
        FMsgSubject.OK( '砞﹚郎把计弄ЧΘ' );
        Delay( COMM_DELAY * 5 );
      end;
    finally
      DisconnectFromAccess;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.SaveToFile(aSilent: Boolean = False): Boolean;
begin
  Result := ConnectToAccess( aSilent );
  if Result then
  begin
    try
      AccessConnection.BeginTrans;
      try
        if not aSilent then
        begin
          FMsgSubject.Normal( '纗把计Paramsい' );
          Delay( COMM_DELAY * 5 );
        end;  
        SetParams;
        AccessConnection.CommitTrans;
      except
        on E: Exception do
        begin
          AccessConnection.RollbackTrans;
          if not aSilent then
            FMsgSubject.Error( Format( '把计Params纗ア毖, :%s',
              [E.Message] ) );
          Result := False;
          Exit;
        end;
      end;
      AccessConnection.BeginTrans;
      try
        if not aSilent then
        begin
          FMsgSubject.Normal( '纗把计Periodい' );
          Delay( COMM_DELAY * 5 );
        end;  
        SetPeriod;
        AccessConnection.CommitTrans;
      except
        on E: Exception do
        begin
          AccessConnection.RollbackTrans;
          if not aSilent then
            FMsgSubject.Error( Format( '把计Period纗ア毖, :%s',
              [E.Message] ) );
          Result := False;
          Exit;
        end;
      end;
      AccessConnection.BeginTrans;
      try
        if not aSilent then
        begin
          FMsgSubject.Normal( '纗把计Notifyい' );
          Delay( COMM_DELAY * 5 );
        end;  
        SetNotify;
        AccessConnection.CommitTrans;
      except
        on E: Exception do
        begin
          AccessConnection.RollbackTrans;
          if not aSilent then
            FMsgSubject.Error( Format( '把计Notify纗ア毖, :%s',
              [E.Message] ) );
          Result := False;
          Exit;
        end;
      end;
      if not aSilent then
      begin
        FMsgSubject.OK( '砞﹚郎把计纗ЧΘ' );
        Delay( COMM_DELAY * 5 );
      end;
    finally
      DisconnectFromAccess;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.GetParams: Boolean;
begin
  Result := True;
  DataReader.Close;
  DataReader.SQL.Text :=
    ' select paramkind, sourcepath, destpath, errpath, backpath,     ' +
    '   LastProceFileName, DbAccount, DbPassword, DbSid, DbRetrySec, ' +
    '   FtpServer, FtpAccount, FtpPassword                           ' +
    ' from params ';
  try
    DataReader.Open;
    DataReader.First;
    while not DataReader.Eof do
    begin
      if ( DataReader.FieldByName( 'paramkind' ).AsString = 'DEX' ) then
      begin
        FParam.DexSrc := DataReader.FieldByName( 'sourcepath' ).AsString;
        FParam.DexDest := DataReader.FieldByName( 'destpath' ).AsString;
        FParam.DexErr := DataReader.FieldByName( 'errpath' ).AsString;
        FParam.DexBackup := DataReader.FieldByName( 'backpath' ).AsString;
        FParam.DexLastProcFile := DataReader.FieldByName( 'LastProceFileName' ).AsString;
      end else
      if ( DataReader.FieldByName( 'paramkind' ).AsString = 'WRAPPER' ) then
      begin
        FParam.WrapperSrc := DataReader.FieldByName( 'sourcepath' ).AsString;
        FParam.WrapperDest := DataReader.FieldByName( 'destpath' ).AsString;
        FParam.WrapperErr := DataReader.FieldByName( 'errpath' ).AsString;
        FParam.WrapperBackup := DataReader.FieldByName( 'backpath' ).AsString;
        FParam.WrapperLastProcFile := DataReader.FieldByName( 'LastProceFileName' ).AsString;
      end else
      if ( DataReader.FieldByName( 'paramkind' ).AsString = 'ASRUN' ) then
      begin
        FParam.AsRunSrc := DataReader.FieldByName( 'sourcepath' ).AsString;
        FParam.AsRunDest := DataReader.FieldByName( 'destpath' ).AsString;
        FParam.AsRunErr := DataReader.FieldByName( 'errpath' ).AsString;
        FParam.AsRunBackup := DataReader.FieldByName( 'backpath' ).AsString;
        FParam.AsRunLastProcFile := DataReader.FieldByName( 'LastProceFileName' ).AsString;
      end else
      if ( DataReader.FieldByName( 'paramkind' ).AsString = 'DB' ) then
      begin
        FDatabase.DbAccount := DataReader.FieldByName( 'dbaccount' ).AsString;
        FDatabase.DbPassoword := DataReader.FieldByName( 'dbpassword' ).AsString;
        FDatabase.DbSid := DataReader.FieldByName( 'dbsid' ).AsString;
        FDatabase.DbRetrySec := DataReader.FieldByName( 'dbretrysec' ).AsInteger;
        if ( FDatabase.DbRetrySec <= 0 ) then FDatabase.DbRetrySec := 60; 
      end else
      if ( DataReader.FieldByName( 'paramkind' ).AsString = 'FTP' ) then
      begin
        FParam.FtpServer := DataReader.FieldByName( 'ftpserver' ).AsString;
        FParam.FtpAccount := DataReader.FieldByName( 'ftpaccount' ).AsString;
        FParam.FtpPassword := DataReader.FieldByName( 'ftppassword' ).AsString;
      end else
      if ( DataReader.FieldByName( 'paramkind' ).AsString = 'FTPDEX' ) then
      begin
        FParam.FtpDexSrc := DataReader.FieldByName( 'sourcepath' ).AsString;
        FParam.FtpDexBackup := DataReader.FieldByName( 'backpath' ).AsString;
      end else
      if ( DataReader.FieldByName( 'paramkind' ).AsString = 'FTPWRAPPER' ) then
      begin
        FParam.FtpWrapperSrc := DataReader.FieldByName( 'sourcepath' ).AsString;
        FParam.FtpWrapperBackup := DataReader.FieldByName( 'backpath' ).AsString;
      end else
      if ( DataReader.FieldByName( 'paramkind' ).AsString = 'FTPASRUN' ) then
      begin
        FParam.FtpAsRunSrc := DataReader.FieldByName( 'sourcepath' ).AsString;
        FParam.FtpAsRunBackup := DataReader.FieldByName( 'backpath' ).AsString;
      end;
      DataReader.Next;
      Delay( COMM_DELAY );
    end;
    DataReader.Close;
  except
    on E: Exception do
    begin
      Result := False;
      FMsgSubject.Error( Format( '把计Params弄ア毖, :%s',
        [E.Message] ) );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.GetPeriod: Boolean;
var
  aIndex: Integer;
  aDateTime: TDateTime;
  aFieldName: String;
begin
  Result := True;
  DataReader.Close;
  DataReader.SQL.Text :=
    ' select checktype, clock1, clock2, clock3, clock4,    ' +
    '   clock5, clock6, clock7, clock8, clock9, clock10,   ' +
    '   clock11, clock12, clock13, clock14, clock15,       ' +
    '   clock16, clock17, clock18, clock19, clock20,       ' +
    '   clock20, clock21, clock22, clock23, clock24,       ' +
    '   clockminute, lastexecute                           ' +
    ' from period                                          ';
  try
    DataReader.Open;
    DataReader.First;
    while not DataReader.Eof do
    begin
      FPeriod.PeriodType := Nvl( DataReader.FieldByName( 'checktype' ).AsString, '1' );
      FPeriod.PeriodMinute := Nvl( DataReader.FieldByName( 'clockminute' ).AsString, 3 );
      FPeriod.LastExecute := 0;
      if TryStrToDateTime(
        DataReader.FieldByName( 'lastexecute' ).AsString, aDateTime ) then
        FPeriod.LastExecute := aDateTime;
      for aIndex := 0 to 23 do
      begin
        if ( aIndex = 0 ) then
          aFieldName := Format( 'clock%d', [24] )
        else
          aFieldName := Format( 'clock%d', [aIndex] );
        FPeriod.Clocks[aIndex].IsSet := DataReader.FieldByName( aFieldName ).AsBoolean;
      end;
      DataReader.Next;
      Delay( COMM_DELAY );
    end;
    DataReader.Close;
  except
    on E: Exception do
    begin
      Result := False;
      FMsgSubject.Error( Format( '把计Period弄ア毖, :%s',
        [E.Message] ) );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.GetNotify: Boolean;
begin
  Result := True;
  DataReader.Close;
  DataReader.SQL.Text :=
    ' select notifytype, notifyenable, email,          ' +
    '  smtpserver, smtpaccount, smtppassword,          ' +
    '  smtpemail                                       ' +
    ' from notify order by notifytype                  ';
  try
    DataReader.Open;
    DataReader.First;
    FNotify.ErrorEmail.Clear;
    FNotify.ChangeEmail.Clear;
    while not DataReader.Eof do
    begin
      if ( DataReader.FieldByName( 'notifytype' ).AsString = '1' ) then
      begin
        FNotify.ErrorNotify := DataReader.FieldByName( 'notifyenable' ).AsBoolean;
        FNotify.ErrorEmail.Add( DataReader.FieldByName( 'email' ).AsString );
      end else
      if ( DataReader.FieldByName( 'notifytype' ).AsString = '2' ) then
      begin
        FNotify.ChangeNotiy := DataReader.FieldByName( 'notifyenable' ).AsBoolean;
        FNotify.ChangeEmail.Add( DataReader.FieldByName( 'email' ).AsString );
      end else
      if ( DataReader.FieldByName( 'notifytype' ).AsString = '3' ) then
      begin
        FNotify.SMTPSetting.Server := DataReader.FieldByName( 'smtpserver' ).AsString;
        FNotify.SMTPSetting.Account := DataReader.FieldByName( 'smtpaccount' ).AsString;
        FNotify.SMTPSetting.Password := DataReader.FieldByName( 'smtppassword' ).AsString;
        FNotify.SMTPSetting.Email := DataReader.FieldByName( 'smtpemail' ).AsString;
      end;
      DataReader.Next;
      Delay( COMM_DELAY );
    end;
    DataReader.Close;
  except
    on E: Exception do
    begin
      Result := False;
      FMsgSubject.Error( Format( '把计Notify弄ア毖, :%s',
        [E.Message] ) );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.SetParams;
var
  aSql: String;
begin
  DataReader.Close;
  DataReader.SQL.Text := ' delete from params ';
  DataReader.ExecSQL;
  {}
  aSql :=
    ' insert into params ( paramkind, sourcepath, destpath,       ' +
    '   errpath, backpath, LastProceFileName, DbAccount,          ' +
    '   DbPassword, DbSid, DbRetrySec, ftpserver, ftpaccount,     ' +
    '   ftppassword )                                             ' +
    ' values ( ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',    ' +
    '   ''%s'', ''%s'', ''%s'', ''%d'', ''%s'', ''%s'', ''%s'' )  ';
  DataReader.SQL.Text := Format( aSql, [
    'DEX', FParam.DexSrc, FParam.DexDest, FParam.DexErr, FParam.DexBackup,
    FParam.DexLastProcFile, EmptyStr, EmptyStr, EmptyStr, 0, EmptyStr,
    EmptyStr, EmptyStr] );
  DataReader.ExecSQL;
  {}
  DataReader.SQL.Text := Format( aSql, [
    'WRAPPER', FParam.WrapperSrc, FParam.WrapperDest, FParam.WrapperErr,
    FParam.WrapperBackup, FParam.WrapperLastProcFile, EmptyStr, EmptyStr,
    EmptyStr, 0, EmptyStr, EmptyStr, EmptyStr] );
  DataReader.ExecSQL;
  {}
  DataReader.SQL.Text := Format( aSql, [
    'ASRUN', FParam.AsRunSrc, FParam.AsRunDest, FParam.AsRunErr,
    FParam.AsRunBackup, FParam.AsRunLastProcFile, EmptyStr, EmptyStr,
    EmptyStr, 0, EmptyStr, EmptyStr,  EmptyStr] );
  DataReader.ExecSQL;
  {}
  DataReader.SQL.Text := Format( aSql, [
    'DB', EmptyStr, EmptyStr, EmptyStr, EmptyStr, EmptyStr, FDatabase.DbAccount,
    FDatabase.DbPassoword, FDatabase.DbSid, FDatabase.DbRetrySec, EmptyStr,
    EmptyStr,  EmptyStr] );
  DataReader.ExecSQL;
  {}
  DataReader.SQL.Text := Format( aSql, [
    'FTP', EmptyStr, EmptyStr, EmptyStr, EmptyStr, EmptyStr, EmptyStr,
     EmptyStr, EmptyStr, 0, FParam.FtpServer, FParam.FtpAccount,
     FParam.FtpPassword] );
  DataReader.ExecSQL;
  {}
  DataReader.SQL.Text := Format( aSql, [
    'FTPDEX', FParam.FtpDexSrc, EmptyStr, EmptyStr, FParam.FtpDexBackup,
    EmptyStr, EmptyStr, EmptyStr, EmptyStr, 0, EmptyStr,
    EmptyStr, EmptyStr] );
  DataReader.ExecSQL;
  {}
  DataReader.SQL.Text := Format( aSql, [
    'FTPWRAPPER', FParam.FtpWrapperSrc, EmptyStr, EmptyStr, FParam.FtpWrapperBackup,
    EmptyStr, EmptyStr, EmptyStr, EmptyStr, 0, EmptyStr,
    EmptyStr, EmptyStr] );
  DataReader.ExecSQL;
  {}
  DataReader.SQL.Text := Format( aSql, [
    'FTPASRUN', FParam.FtpAsRunSrc, EmptyStr, EmptyStr, FParam.FtpAsRunBackup,
    EmptyStr, EmptyStr, EmptyStr, EmptyStr, 0, EmptyStr,
    EmptyStr, EmptyStr] );
  DataReader.ExecSQL;
  {}
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.SetPeriod;
var
  aSql, aLastExecute: String;
begin
  DataReader.Close;
  DataReader.SQL.Text := ' delete from period ';
  DataReader.ExecSQL;
  aSql :=
    ' insert into period ( checktype, clock1, clock2, clock3,     ' +
    '   clock4, clock5, clock6, clock7, clock8, clock9,           ' +
    '   clock10, clock11, clock12, clock13, clock14, clock15,     ' +
    '   clock16, clock17, clock18, clock19, clock20, clock21,     ' +
    '   clock22, clock23, clock24, clockminute, lastexecute  )    ' +
    ' values ( ''%s'', %s, %s, %s, %s, %s,                        ' +
    '   %s, %s, %s, %s, %s, %s, %s, %s,                           ' +
    '   %s, %s, %s, %s, %s, %s, %s,                               ' +
    '   %s, %s, %s, %s, %d, %s  )                                 ';
  aLastExecute := '0';  
  if ( FPeriod.LastExecute <> 0 ) then
    aLastExecute := FormatDateTime( 'yyyy/mm/dd hh:nn:ss', FPeriod.LastExecute );
  DataReader.SQL.Text := Format( aSql, [
    FPeriod.PeriodType,
    IIF( FPeriod.Clocks[1].IsSet, 'True', 'False' ),
    IIF( FPeriod.Clocks[2].IsSet, 'True', 'False' ),
    IIF( FPeriod.Clocks[3].IsSet, 'True', 'False' ),
    IIF( FPeriod.Clocks[4].IsSet, 'True', 'False' ),
    IIF( FPeriod.Clocks[5].IsSet, 'True', 'False' ),
    IIF( FPeriod.Clocks[6].IsSet, 'True', 'False' ),
    IIF( FPeriod.Clocks[7].IsSet, 'True', 'False' ),
    IIF( FPeriod.Clocks[8].IsSet, 'True', 'False' ),
    IIF( FPeriod.Clocks[9].IsSet, 'True', 'False' ),
    IIF( FPeriod.Clocks[10].IsSet, 'True', 'False' ),
    IIF( FPeriod.Clocks[11].IsSet, 'True', 'False' ),
    IIF( FPeriod.Clocks[12].IsSet, 'True', 'False' ),
    IIF( FPeriod.Clocks[13].IsSet, 'True', 'False' ),
    IIF( FPeriod.Clocks[14].IsSet, 'True', 'False' ),
    IIF( FPeriod.Clocks[15].IsSet, 'True', 'False' ),
    IIF( FPeriod.Clocks[16].IsSet, 'True', 'False' ),
    IIF( FPeriod.Clocks[17].IsSet, 'True', 'False' ),
    IIF( FPeriod.Clocks[18].IsSet, 'True', 'False' ),
    IIF( FPeriod.Clocks[19].IsSet, 'True', 'False' ),
    IIF( FPeriod.Clocks[20].IsSet, 'True', 'False' ),
    IIF( FPeriod.Clocks[21].IsSet, 'True', 'False' ),
    IIF( FPeriod.Clocks[22].IsSet, 'True', 'False' ),
    IIF( FPeriod.Clocks[23].IsSet, 'True', 'False' ),
    IIF( FPeriod.Clocks[0].IsSet, 'True', 'False' ),
    FPeriod.PeriodMinute,
    aLastExecute] );
  DataReader.ExecSQL;  
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.SetNotify;
var
  aSql: String;
  aIndex: Integer;
begin
  DataReader.Close;
  DataReader.SQL.Text := ' delete from notify ';
  DataReader.ExecSQL;
  aSql :=
    ' insert into notify ( notifytype, notifyenable, email,   ' +
    '  smtpserver, smtpaccount, smtppassword, smtpemail )     ' +
    ' values ( ''%s'', %s, ''%s'', ''%s'', ''%s'', ''%s'',    ' +
    '  ''%s'' )                                               ';
  for aIndex := 0 to FNotify.ErrorEmail.Count - 1 do
  begin
    DataReader.SQL.Text := Format( aSql, [
      '1',
      IIF( FNotify.ErrorNotify, 'True', 'False' ),
      FNotify.ErrorEmail[aIndex],
      EmptyStr,
      EmptyStr,
      EmptyStr,
      EmptyStr] );
    DataReader.ExecSQL;
  end;
  for aIndex := 0 to FNotify.ChangeEmail.Count - 1 do
  begin
    DataReader.SQL.Text := Format( aSql, [
      '2',
      IIF( FNotify.ChangeNotiy, 'True', 'False' ),
      FNotify.ChangeEmail[aIndex],
      EmptyStr,
      EmptyStr,
      EmptyStr,
      EmptyStr] );
    DataReader.ExecSQL;
  end;
  DataReader.SQL.Text := Format( aSql, [
    '3',
    'null',
    EmptyStr,
    FNotify.SMTPSetting.Server,
    FNotify.SMTPSetting.Account,
    FNotify.SMTPSetting.Password,
    FNotify.SMTPSetting.Email] );
  DataReader.ExecSQL;
end;

{ ---------------------------------------------------------------------------- }

end.
