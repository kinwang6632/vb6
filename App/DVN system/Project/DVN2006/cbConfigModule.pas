unit cbConfigModule;

interface

uses
  SysUtils, Classes, Controls, Forms, DB, ADODB, DBClient, IniFiles,
  cxTextEdit, StdCtrls, cxTL, 
  cbAppClass;

type
  TConfigModule = class(TDataModule)
    ConfigConnection: TADOConnection;
    DataReader: TADOQuery;
    DataWriter: TADOQuery;
    CmdErrorEnv: TClientDataSet;
    HighCmdEnv: TClientDataSet;
    LowCmdEnv: TClientDataSet;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
    FSoList: TSoList;
    FTextSubj: TTextSubject;
    FConfigPath: String;
    FConfigList: TStringList;
    FActiveConfigFile: String;
    FCasEnv: TCasEnv;
    FCommEnv: TCommEnv;
    function GetConfigFile(aIndex: Integer): String;
    function GetConfigFileCount: Integer;
    function ConnectToAccess(const aConfigName: String): Boolean;
    procedure DisconnectFromAccess;
    function GetCasEnv: Boolean;
    function GetCommEnv: Boolean;
    function GetSoEnv: Boolean;
    function GetHighCmdEnv: Boolean;
    function GetLowCmdEnv: Boolean;
    function GetCmdErrEnv: Boolean;
    {}
    procedure SetCASEnv;
    procedure SetCommEnv;
    procedure SetSoEnv;
    procedure SetHighCmdEnv;
    procedure SetLowCmdEnv;
    procedure SetCmdErrEnv;
  public
    { Public declarations }
    procedure RefreshConfigFile;
    procedure SaveActiveConfigFile(aIni: TIniFile);
    procedure RestoreActiveConfigFile(aIni: TIniFile);
    function ChangeConfigFile: Boolean;
    property ConfigFiles[Index: Integer]: String read GetConfigFile;
    property ConfigFileCount: Integer read GetConfigFileCount;
    property ActiveConfigFile: String read FActiveConfigFile write FActiveConfigFile;
    property SoList: TSoList read FSoList;
    property CasEnv: TCasEnv read FCasEnv;
    property CommEnv: TCommEnv read FCommEnv;
    property TextSubj: TTextSubject read FTextSubj;
    function ShowConfigForm: TModalResult;
    procedure GetConfigFromValue;
    procedure ConfigSaveToFile;
  end;

var
  ConfigModule: TConfigModule;

implementation

{$R *.dfm}

uses cbUtilis, cbMain, cbConfig, Math, cxSpinEdit;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.DataModuleCreate(Sender: TObject);
begin
  FConfigPath := IncludeTrailingPathDelimiter( ExtractFilePath(
    ParamStr( 0 ) ) );
  FConfigList := TStringList.Create;
  FActiveConfigFile := EmptyStr;
  FSoList := TSoList.Create;
  FTextSubj := TTextSubject.Create;
  FCasEnv := TCasEnv.Create;
  FCommEnv := TCommEnv.Create;
  HighCmdEnv.CreateDataSet;
  LowCmdEnv.CreateDataSet;
  CmdErrorEnv.CreateDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.DataModuleDestroy(Sender: TObject);
begin
  FCasEnv.Free;
  FCommEnv.Free;
  FSoList.Free;
  FTextSubj.Free;
  FConfigList.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.RefreshConfigFile;
const
  Ext = '*.cfg';
var
  aFound: Integer;
  aPath: String;
  aSearchRect: TSearchRec;
begin
  FConfigList.Clear;
  aPath := IncludeTrailingPathDelimiter( ExtractFilePath(
    ParamStr( 0 ) ) ) + Ext;
  aFound := FindFirst( aPath, faAnyFile, aSearchRect );
  while ( aFound = 0 ) do
  begin
    FConfigList.Add( aSearchRect.Name );
    FTextSubj.InfoFmt( LanguageManager.Get( 'SConfigSearching' ), [aSearchRect.Name] );
    Delay( 50 );
    aFound := FindNext( aSearchRect );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.ChangeConfigFile: Boolean;
const
  Ext = '.cfg';
begin
  Result := FileExists( IncludeTrailingPathDelimiter( ExtractFilePath(
    ParamStr( 0 ) ) ) + FActiveConfigFile + '.cfg' );
  {}
  if ( not Result ) then
  begin
    FTextSubj.ErrorFmt( LanguageManager.Get( 'SConfigFileNotExists' ),
      [Nvl( FActiveConfigFile, LanguageManager.Get( 'SNull' ) )] );
  end;
  {}
  if ( Result ) then Result := ConnectToAccess( FActiveConfigFile );
  if ( Result ) then
  begin
    FTextSubj.InfoFmt( LanguageManager.Get( 'SConfigBeginRead' ),
       [FActiveConfigFile] );
    Delay( 500 );
    Result := ( Result and GetCasEnv );
    Delay( 1200 );
    Result := ( Result and GetCommEnv );
    Delay( 1200 );
    Result := ( Result and GetHighCmdEnv );
    Delay( 1200 );
    Result := ( Result and GetLowCmdEnv );
    Delay( 1200 );
    Result := ( Result and GetCmdErrEnv );
    Delay( 1200 );
    Result := ( Result and GetSoEnv );
  end;
  if ( Result ) then
    FTextSubj.OK( LanguageManager.Get( 'SConfigEndRead' ) )
  else
    FTextSubj.Error( LanguageManager.Get( 'SConfigLoadWarning' ) );
  DisconnectFromAccess;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.GetConfigFile(aIndex: Integer): String;
begin
  Result := FConfigList[aIndex];
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.GetConfigFileCount: Integer;
begin
  Result := FConfigList.Count;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.ConnectToAccess(const aConfigName: String): Boolean;
const
  aPass = 'cyc84177282';
  Ext = '.cfg';
var
  aMsg, aFileName: String;
begin
  aFileName := IncludeTrailingPathDelimiter( ExtractFilePath(
    ParamStr( 0 ) )  )+ aConfigName + Ext;
  ConfigConnection.Connected := False;
  ConfigConnection.ConnectionString := Format(
    'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=%s;Jet OLEDB:Database Password=%s;'+
    'Mode=Share Deny Read|Share Deny Write;', [aFileName, aPass] );
  try
    ConfigConnection.Connected := True;
  except
    on E: Exception do aMsg := E.Message;
  end;
  Result := ConfigConnection.Connected;
  if Result then
    FTextSubj.InfoFmt( LanguageManager.Get( 'SConfigOpenSuccess' ), [aFileName] )
  else
    FTextSubj.ErrorFmt( LanguageManager.Get( 'SConfigOpenError' ), [aFileName, aMsg] );
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.DisconnectFromAccess;
begin
  if ConfigConnection.Connected then ConfigConnection.Close;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.GetSoEnv: Boolean;
var
  aSo: TSo;
begin
  Result := False;
  FSoList.Clear;
  try
    DataReader.Close;
    DataReader.SQL.Text :=
      ' select * from soenv order by cint( so_comp_code ) ';
    DataReader.Open;
    DataReader.First;
    while not DataReader.Eof do
    begin
      aSo := TSo.Create;
      aSo.Selected := ( DataReader.FieldByName( 'so_selected' ).AsString = 'Y' );
      aSo.PosName := DataReader.FieldByName( 'so_pos_name' ).AsString;
      aSo.SoType := DataReader.FieldByName( 'so_type' ).AsString;
      aSo.CompCode := DataReader.FieldByName( 'so_comp_code' ).AsString;
      aSo.CompName := DataReader.FieldByName( 'so_comp_name' ).AsString;
      aSo.NetworkId := DataReader.FieldByName( 'so_network_id' ).AsString;
      aSo.AreaCode := DataReader.FieldByName( 'so_area_code' ).AsString;
      aSo.DbUserId := DataReader.FieldByName( 'so_loginuser' ).AsString;
      aSo.DbUserPass := DataReader.FieldByName( 'so_loginpass' ).AsString;
      aSO.DbAliase := DataReader.FieldByName( 'so_dbaliase' ).AsString;
      aSo.DbState := dbClose;
      FSoList.Add( aSo );
      DataReader.Next;
      FTextSubj.InfoFmt( LanguageManager.Get( 'SConfigSoRead' ),
        [aSo.CompName] );
      Delay( 100 );  
    end;
    if ( FSoList.Count > 0 ) then
    begin
      Result := True;
      FTextSubj.OKFmt( LanguageManager.Get( 'SConfigSoReadSuccess' ), [FSoList.Count] );
    end  else
      FTextSubj.ErrorFmt( LanguageManager.Get( 'SConfigSoReadError' ),
        [LanguageManager.Get( 'SNull' )] )
  except
    on E: Exception do
    begin
      FTextSubj.ErrorFmt( LanguageManager.Get( 'SConfigSoReadError' ), [E.Message] );
      FSoList.Clear;
    end;
  end;
  DataReader.Close;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.GetCasEnv: Boolean;
begin
  Result := False;
  FCasEnv.Ip := '0.0.0.0';
  FCasEnv.SendPort := '0';
  FCasEnv.RecvPort := '0';
  FCasEnv.Protocol := '2';
  try
    DataReader.Close;
    DataReader.SQL.Text := ' select * from casenv ';
    DataReader.Open;
    DataReader.First;
    if ( DataReader.IsEmpty ) then
    begin
      FTextSubj.ErrorFmt( LanguageManager.Get( 'SConfigCAReadError' ),
       [LanguageManager.Get( 'SNull' )] );
    end else
    begin
      FCasEnv.Ip := DataReader.FieldByName( 'ca_ip' ).AsString;
      FTextSubj.InfoFmt( LanguageManager.Get( 'SConfigCAReadingIP' ), [FCasEnv.Ip] );
      Delay( 100 );
      {}
      FCasEnv.SendPort := DataReader.FieldByName( 'ca_send_port' ).AsString;
      FTextSubj.InfoFmt( LanguageManager.Get( 'SConfigCAReadingSPort' ), [FCasEnv.SendPort] );
      Delay( 100 );
      {}
      FCasEnv.RecvPort := DataReader.FieldByName( 'ca_recv_port' ).AsString;
      FTextSubj.InfoFmt( LanguageManager.Get( 'SConfigCAReadingRPort' ), [FCasEnv.RecvPort] );
      Delay( 100 );
      {}
      FCasEnv.Protocol := DataReader.FieldByName( 'ca_protocol' ).AsString;
      FTextSubj.InfoFmt( LanguageManager.Get( 'SConfigCAReadingProtocol' ), [FCasEnv.Protocol] );
      Delay( 100 );
      {}
      Result := True;
      FTextSubj.OK( LanguageManager.Get( 'SConfigCAReadSuccess' ) );
    end;
  except
    on E: Exception do
    begin
      FTextSubj.ErrorFmt( LanguageManager.Get( 'SConfigCAReadError' ), [E.Message] );
    end;
  end;
  DataReader.Close;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.GetCommEnv: Boolean;
var
  aText: String;
begin
  Result := False;
  FCommEnv.DbRetryFreq := 20;
  FCommEnv.BusyTimeStart := EmptyStr;
  FCommEnv.BusyTimeEnd := EmptyStr;
  FCommEnv.BusyTimeReadFreq := 3;
  FCommEnv.NorTimeReadFreq := 3;
  FCommEnv.DbProcRecords := 20;
  FCommEnv.DbWriteError := True;
  FCommEnv.DbProcIPPV := 'N';
  FCommEnv.DbProcBatch := 'Y';
  FCommEnv.DbWaterMark := 200;
  FCommEnv.CARetryFreq := 30;
  FCommEnv.CACheckFreq := 10;
  FCommEnv.CAEnableSend := True;
  FCommEnv.CAEnableRecv := False;
  FCommEnv.CASendDelay := 0.3;
  FCommEnv.CAMaxError := 3;
  FCommEnv.CAProductDefine := EmptyStr;
  FCommEnv.CAIdleTime := 60;
  try
    DataReader.Close;
    DataReader.SQL.Text := ' select * from commenv ';
    DataReader.Open;
    DataReader.First;
    if ( DataReader.IsEmpty ) then
    begin
      FTextSubj.ErrorFmt( LanguageManager.Get( 'SConfigCommReadError' ),
       [LanguageManager.Get( 'SNull' )] );
    end else
    begin
      {}
      FCommEnv.DbRetryFreq := DataReader.FieldByName( 'dbretryfreq' ).AsInteger;
      FTextSubj.InfoFmt( LanguageManager.Get( 'SConfigCommDbRetryFreq' ), [FCommEnv.DbRetryFreq] );
      Delay( 100 );
      {}
      FCommEnv.BusyTimeStart := DataReader.FieldByName( 'busytimestart' ).AsString;
      FTextSubj.InfoFmt( LanguageManager.Get( 'SConfigCommBusyTimeStart' ), [FCommEnv.BusyTimeStart] );
      Delay( 100 );
      {}
      FCommEnv.BusyTimeEnd := DataReader.FieldByName( 'busytimeend' ).AsString;
      FTextSubj.InfoFmt( LanguageManager.Get( 'SConfigCommBusyTimeEnd' ), [FCommEnv.BusyTimeEnd] );
      Delay( 100 );
      {}
      FCommEnv.BusyTimeReadFreq := DataReader.FieldByName( 'busytimereadfreq' ).AsInteger;
      FTextSubj.InfoFmt( LanguageManager.Get( 'SConfigCommBusyTimeReadFreq' ), [FCommEnv.BusyTimeReadFreq] );
      Delay( 100 );
      {}
      FCommEnv.NorTimeReadFreq := DataReader.FieldByName( 'nortimereadfreq' ).AsInteger;
      FTextSubj.InfoFmt( LanguageManager.Get( 'SConfigCommNorTimeReadFreq' ), [FCommEnv.NorTimeReadFreq] );
      Delay( 100 );
      {}
      FCommEnv.DbProcRecords := DataReader.FieldByName( 'dbprocrecords' ).AsInteger;
      FTextSubj.InfoFmt( LanguageManager.Get( 'SConfigCommDbProcRecords' ), [FCommEnv.DbProcRecords] );
      Delay( 100 );
      {}
      FCommEnv.DbWriteError := ( DataReader.FieldByName( 'dbwriteerror' ).AsString = 'Y' );
      aText := LanguageManager.Get( 'SNo' );
      if ( FCommEnv.DbWriteError ) then aText := LanguageManager.Get( 'SYes' );
      FTextSubj.InfoFmt( LanguageManager.Get( 'SConfigCommDbWriteError' ), [aText] );
      Delay( 100 );
      {}
      FCommEnv.DbProcIPPV := DataReader.FieldByName( 'dbprocippv' ).AsString;
      aText := LanguageManager.Get( 'SConfigCommAIPPV' );
      if ( FCommEnv.DbProcIPPV = 'N' ) then
        aText := LanguageManager.Get( 'SConfigCommNIPPV' )
      else if ( FCommEnv.DbProcIPPV = 'O' ) then
        aText := LanguageManager.Get( 'SConfigCommOIPPV' );
      FTextSubj.InfoFmt( LanguageManager.Get( 'SConfigCommDbProcIPPV' ), [aText] );
      Delay( 100 );
      {}
      FCommEnv.DbProcBatch := DataReader.FieldByName( 'dbprocbatch' ).AsString;
      aText := LanguageManager.Get( 'SConfigCommABatch' );
      if ( FCommEnv.DbProcBatch = 'N' ) then
        aText := LanguageManager.Get( 'SConfigCommNBatch' )
      else if ( FCommEnv.DbProcBatch = 'O' ) then
        aText := LanguageManager.Get( 'SConfigCommOBatch' );
      FTextSubj.InfoFmt( LanguageManager.Get( 'SConfigCommDbProcBatch' ), [aText] );
      Delay( 100 );
      {}
      FCommEnv.DbWaterMark := DataReader.FieldByName( 'dbwatermark' ).AsInteger;
      FTextSubj.InfoFmt( LanguageManager.Get( 'SConfigCommDbWaterMark' ), [FCommEnv.DbWaterMark] );
      Delay( 100 );
      {}
      FCommEnv.CARetryFreq := DataReader.FieldByName( 'caretryfreq' ).AsInteger;
      FTextSubj.InfoFmt( LanguageManager.Get( 'SConfigCommCARetryFreq' ), [FCommEnv.CARetryFreq] );
      Delay( 100 );
      {}
      FCommEnv.CACheckFreq := DataReader.FieldByName( 'cacheckfreq' ).AsInteger;
      FTextSubj.InfoFmt( LanguageManager.Get( 'SConfigCommCACheckFreq' ), [FCommEnv.CACheckFreq] );
      Delay( 100 );
      {}
      FCommEnv.CAEnableSend := ( DataReader.FieldByName( 'caenablesend' ).AsString = 'Y' );
      aText := LanguageManager.Get( 'SNo' );
      if ( FCommEnv.CAEnableSend ) then aText := LanguageManager.Get( 'SYes' );
      FTextSubj.InfoFmt( LanguageManager.Get( 'SConfigCommCAEnableSend' ), [aText] );
      Delay( 100 );
      FCommEnv.CAEnableRecv := ( DataReader.FieldByName( 'caenablerecv' ).AsString = 'Y' );
      aText := LanguageManager.Get( 'SNo' );
      if ( FCommEnv.CAEnableRecv ) then aText := LanguageManager.Get( 'SYes' );
      FTextSubj.InfoFmt( LanguageManager.Get( 'SConfigCommCAEnableRecv' ), [aText] );
      Delay( 100 );
      {}
      FCommEnv.CASendDelay := RoundTo(
        DataReader.FieldByName( 'casenddelay' ).AsFloat, -3 );
      FTextSubj.InfoFmt( LanguageManager.Get( 'SConfigCommCASendDelay' ),
        [FloatToStr( FCommEnv.CASendDelay )] );
      Delay( 100 );
      {}
      FCommEnv.CAMaxError := DataReader.FieldByName( 'camaxerror' ).AsInteger;
      FTextSubj.InfoFmt( LanguageManager.Get( 'SConfigCommCAMaxError' ), [FCommEnv.CAMaxError] );
      Delay( 100 );
      {}
      FCommEnv.CAProductDefine := DataReader.FieldByName( 'caproductdefine' ).AsString;
      FTextSubj.InfoFmt( LanguageManager.Get( 'SConfigCommCAProductDefine' ),
        [Nvl( FCommEnv.CAProductDefine, LanguageManager.Get( 'SNull' ) )] );
      Delay( 100 );
      {}
      FCommEnv.CAIdleTime := StrToIntDef( DataReader.FieldByName( 'caidletime' ).AsString, 60 );
      FTextSubj.InfoFmt( LanguageManager.Get( 'SConfigCommCAIdleTime' ), [FCommEnv.CAIdleTime] );
      Delay( 100 );
      {}
      Result := True;
      FTextSubj.OK( LanguageManager.Get( 'SConfigCommReadSuccess' ) );
    end;
  except
    on E: Exception do
    begin
      FTextSubj.ErrorFmt( LanguageManager.Get( 'SConfigCommReadError' ), [E.Message] );
    end;
  end;
  DataReader.Close;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.GetCmdErrEnv: Boolean;
begin
  Result := False;
  CmdErrorEnv.EmptyDataSet;
  try
    DataReader.Close;
    DataReader.SQL.Text := ' select * from cmderrenv ';
    DataReader.Open;
    DataReader.First;
    {}
    while not DataReader.Eof do
    begin
      CmdErrorEnv.AppendRecord( [
        DataReader.FieldByName( 'errorflag' ).AsString,
        DataReader.FieldByName( 'errorcode' ).AsString,
        DataReader.FieldByName( 'errordesc' ).AsString] );
      FTextSubj.InfoFmt( LanguageManager.Get( 'SConfigErrCodeRead' ),
        [CmdErrorEnv.FieldByName( 'errorcode' ).AsString] );
      Delay( 100 );
      DataReader.Next;
    end;
    if ( HighCmdEnv.RecordCount > 0 ) then
    begin
      Result := True;
      FTextSubj.OKFmt( LanguageManager.Get( 'SConfigErrCodeReadSuccess' ),
       [CmdErrorEnv.RecordCount] )
    end else
      FTextSubj.ErrorFmt( LanguageManager.Get( 'SConfigErrCodeReadError' ),
       [LanguageManager.Get( 'SNull' )] );
  except
    on E: Exception do
    begin
      FTextSubj.ErrorFmt( LanguageManager.Get( 'SConfigErrCodeReadError' ), [E.Message] );
    end;
  end;
  DataReader.Close;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.GetHighCmdEnv: Boolean;
begin
  Result := False;
  HighCmdEnv.EmptyDataSet;
  try
    DataReader.Close;
    DataReader.SQL.Text := ' select * from highcmdenv ';
    DataReader.Open;
    DataReader.First;
    {}
    while not DataReader.Eof do
    begin
      HighCmdEnv.AppendRecord( [
        DataReader.FieldByName( 'highlevelcmd' ).AsString,
        DataReader.FieldByName( 'description' ).AsString,
        DataReader.FieldByName( 'lowlevelcmd' ).AsString,
        DataReader.FieldByName( 'cmdtype' ).AsString,
        DataReader.FieldByName( 'cmdtypepriority' ).AsString,
        DataReader.FieldByName( 'cmdpriority' ).AsString,
        DataReader.FieldByName( 'cmdbynote' ).AsString] );
      FTextSubj.InfoFmt( LanguageManager.Get( 'SConfigHighCmdRead' ),
        [HighCmdEnv.FieldByName( 'description' ).AsString] );
      Delay( 100 );
      DataReader.Next;
    end;
    if ( HighCmdEnv.RecordCount > 0 ) then
    begin
      Result := True;
      FTextSubj.OKFmt( LanguageManager.Get( 'SConfigHighCmdReadSuccess' ),
       [HighCmdEnv.RecordCount] )
    end else
      FTextSubj.ErrorFmt( LanguageManager.Get( 'SConfigHighCmdReadError' ),
       [LanguageManager.Get( 'SNull' )] );
  except
    on E: Exception do
    begin
      FTextSubj.ErrorFmt( LanguageManager.Get( 'SConfigHighCmdReadError' ), [E.Message] );
    end;
  end;
  DataReader.Close;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.GetLowCmdEnv: Boolean;
begin
  Result := False;
  LowCmdEnv.EmptyDataSet;
  try
    DataReader.Close;
    DataReader.SQL.Text := ' select * from lowcmdenv ';
    DataReader.Open;
    DataReader.First;
    {}
    while not DataReader.Eof do
    begin
      LowCmdEnv.AppendRecord( [
        DataReader.FieldByName( 'lowlevelcmd' ).AsString,
        DataReader.FieldByName( 'description' ).AsString] );
      FTextSubj.InfoFmt( LanguageManager.Get( 'SConfigLowCmdRead' ),
        [LowCmdEnv.FieldByName( 'lowlevelcmd' ).AsString] );
      Delay( 100 );
      DataReader.Next;
    end;
    if ( LowCmdEnv.RecordCount > 0 ) then
    begin
      Result := True;
      FTextSubj.OKFmt( LanguageManager.Get( 'SConfigLowCmdReadSuccess' ),
       [LowCmdEnv.RecordCount] );
    end else
      FTextSubj.ErrorFmt( LanguageManager.Get( 'SConfigLowCmdReadError' ),
       [LanguageManager.Get( 'SNull' )] );
  except
    on E: Exception do
    begin
      FTextSubj.ErrorFmt( LanguageManager.Get( 'SConfigLowCmdReadError' ), [E.Message] );
    end;
  end;
  DataReader.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.SaveActiveConfigFile(aIni: TIniFile);
begin
  aIni.WriteString( 'ConfigFile', 'Active', ExtractFileName( FActiveConfigFile ) );
  aIni.UpdateFile;
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.RestoreActiveConfigFile(aIni: TIniFile);
begin
  FActiveConfigFile := aIni.ReadString( 'ConfigFile', 'Active', EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.ShowConfigForm: TModalResult;
begin
  fmConfig := TfmConfig.Create( nil );
  try
    fmConfig.btnConfirm.Enabled := ( not IsThreadRun );
    Result := fmConfig.ShowModal;
    if ( Result = mrOk ) then
    begin
      Screen.Cursor := crHourGlass;
      try
        GetConfigFromValue;
        ConfigSaveToFile;
      finally
        Screen.Cursor := crDefault;
      end;  
    end; 
  finally
    fmConfig.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.GetConfigFromValue;
var
  aIndex: Integer;
  aNode: TcxTreeListNode;
  aSo: TSo;
begin
  FCasEnv.Ip := fmConfig.txtCASIP.Text;
  FCasEnv.SendPort := fmConfig.txtSendPort.Text;
  FCasEnv.RecvPort := fmConfig.txtRecvPort.Text;
  FCasEnv.Protocol := fmConfig.txtCAProtocol.Text;
  {}
  FCommEnv.DbRetryFreq := fmConfig.txtDbRetryFreq.Value;
  FCommEnv.BusyTimeStart := fmConfig.txtBusyTimeStart.Text;
  FCommEnv.BusyTimeEnd := fmConfig.txtBusyTimeEnd.Text;
  FCommEnv.BusyTimeReadFreq := fmConfig.txtBusyTimeReadFreq.Value;
  FCommEnv.NorTimeReadFreq := fmConfig.txtNorTimeFreq.Value;
  FCommEnv.DbProcRecords := fmConfig.txtDbProcRecords.Value;
  FCommEnv.DbWriteError := fmConfig.chkWriteYes.Checked;
  {}
  if fmConfig.chkProcessIPPV_A.Checked then
    FCommEnv.DbProcIPPV := 'A'
  else if fmConfig.chkProcessIPPV_N.Checked then
    FCommEnv.DbProcIPPV := 'N'
  else
    FCommEnv.DbProcIPPV := 'O';
  {}
  if fmConfig.ChkProcessBatch_A.Checked then
    FCommEnv.DbProcBatch := 'A'
  else if fmConfig.ChkProcessBatch_N.Checked then
    FCommEnv.DbProcBatch := 'N'
  else
    FCommEnv.DbProcBatch := 'O';
  {}
  FCommEnv.DbWaterMark := FCommEnv.DbWaterMark;
  FCommEnv.CARetryFreq := fmConfig.txtCARertyFeq.Value;
  FCommEnv.CACheckFreq := fmConfig.txtCACheckFreq.Value;
  FCommEnv.CAEnableSend := fmConfig.chkEnableSend.Checked;
  FCommEnv.CAEnableRecv := fmConfig.chkEnableRecv.Checked;
  FCommEnv.CASendDelay := fmConfig.txtCASendDelay.Value;
  FCommEnv.CAMaxError := fmConfig.txtCAMaxError.Value;
  FCommEnv.CAProductDefine := fmConfig.txtCAProductDefine.Text;
  FCommEnv.CAIdleTime := fmConfig.txtCAIdleTime.Value;
  {}
  FSoList.Clear;
  for aIndex := 0 to fmConfig.ConfigSoTree.Nodes.Count - 1 do
  begin
    if fmConfig.ConfigSoTree.Nodes.Items[aIndex].Level = 0 then Continue;
    aNode := fmConfig.ConfigSoTree.Nodes.Items[aIndex];
    aSo := TSo.Create;
    aSo.Selected := ( aNode.Values[fmConfig.ConfigSoTreeCol1.ItemIndex] = True );
    aSo.PosName := aNode.Texts[fmConfig.ConfigSoTreeCol9.ItemIndex];
    aSo.SoType := aNode.Texts[fmConfig.ConfigSoTreeCol8.ItemIndex];
    aSo.CompCode := aNode.Texts[fmConfig.ConfigSoTreeCol2.ItemIndex];
    aSo.CompName := aNode.Texts[fmConfig.ConfigSoTreeCol3.ItemIndex];
    aSo.NetworkId := aNode.Texts[fmConfig.ConfigSoTreeCol4.ItemIndex];
    aSo.AreaCode := aNode.Texts[fmConfig.ConfigSoTreeCol10.ItemIndex];
    aSo.DbUserId := aNode.Texts[fmConfig.ConfigSoTreeCol5.ItemIndex];
    aSo.DbUserPass := aNode.Texts[fmConfig.ConfigSoTreeCol6.ItemIndex];
    aSo.DbAliase := aNode.Texts[fmConfig.ConfigSoTreeCol7.ItemIndex];
    aSo.DbState := dbClose;
    FSoList.Add( aSo );
  end;
  {}
  HighCmdEnv.EmptyDataSet;
  for aIndex := 0 to fmConfig.HighCmdTree.Nodes.Count - 1 do
  begin
    aNode := fmConfig.HighCmdTree.Nodes.Items[aIndex];
    HighCmdEnv.AppendRecord( [
      aNode.Texts[fmConfig.HighCmdTreeCol1.ItemIndex],
      aNode.Texts[fmConfig.HighCmdTreeCol2.ItemIndex],
      aNode.Texts[fmConfig.HighCmdTreeCol4.ItemIndex],
      aNode.Texts[fmConfig.HighCmdTreeCol3.ItemIndex],
      aNode.Texts[fmConfig.HighCmdTreeCol6.ItemIndex],
      aNode.Texts[fmConfig.HighCmdTreeCol7.ItemIndex],
      aNode.Texts[fmConfig.HighCmdTreeCol5.ItemIndex] ] );
  end;
  {}
  LowCmdEnv.EmptyDataSet;
  for aIndex := 0 to fmConfig.LowCmdTree.Nodes.Count - 1 do
  begin
    aNode := fmConfig.LowCmdTree.Nodes.Items[aIndex];
    LowCmdEnv.AppendRecord( [
      aNode.Texts[fmConfig.LowCmdTreeCol1.ItemIndex],
      aNode.Texts[fmConfig.LowCmdTreeCol2.ItemIndex]] );
  end;
  {}
  CmdErrorEnv.EmptyDataSet;
  for aIndex := 0 to fmConfig.ErrorTree.Nodes.Count - 1 do
  begin
    aNode := fmConfig.ErrorTree.Nodes.Items[aIndex];
    CmdErrorEnv.AppendRecord( [
      aNode.Texts[fmConfig.ErrorTreeCol1.ItemIndex],
      aNode.Texts[fmConfig.ErrorTreeCol2.ItemIndex],
      aNode.Texts[fmConfig.ErrorTreeCol3.ItemIndex] ] );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.ConfigSaveToFile;
begin
  if ConnectToAccess( FActiveConfigFile ) then
  begin
    SetCASEnv;
    SetCommEnv;
    SetSoEnv;
    SetHighCmdEnv;
    SetLowCmdEnv;
    SetCmdErrEnv;
  end;
  DisconnectFromAccess;
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.SetCASEnv;
begin
  DataWriter.Close;
  DataWriter.SQL.Text := ' delete from casenv ';
  DataWriter.ExecSQL;
  try
    DataWriter.SQL.Text := Format(
      ' insert into casenv ( ca_ip, ca_send_port, ca_recv_port, ' +
      '   ca_protocol ) values ( ''%s'', ''%s'', ''%s'',        ' +
      '   ''%s'' ) ',
      [FCasEnv.Ip, FCasEnv.SendPort, FCasEnv.RecvPort, FCasEnv.Protocol] );
    DataWriter.ExecSQL;
    FTextSubj.OK( LanguageManager.Get( 'SConfigSaveCasSuccess' ) );
  except
    on E: Exception do
      FTextSubj.ErrorFmt( LanguageManager.Get( 'SConfigSaveCasError' ), [E.Message] );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.SetCommEnv;
begin
  DataWriter.Close;
  DataWriter.SQL.Text := ' delete from commenv ';
  DataWriter.ExecSQL;
  try
    DataWriter.SQL.Text := Format(
      ' insert into commenv ( DbRetryFreq, BusyTimeStart,  ' +
      '   BusyTimeEnd, BusyTimeReadFreq, NorTimeReadFreq,  ' +
      '   DbProcRecords, DbWriteError, DbProcIPPV,         ' +
      '   DbProcBatch, DbWaterMark,                        ' +
      '   CARetryFreq, CACheckFreq, CAEnableSend,          ' +
      '   CAEnableRecv, CASendDelay, CAMaxError,           ' +
      '   CAProductDefine, CAIdleTime )                    ' +
      ' values ( ''%d'', ''%s'',                           ' +
      '   ''%s'', ''%d'', ''%d'',                          ' +
      '   ''%d'', ''%s'', ''%s'',                          ' +
      '   ''%s'', ''%d'',                                  ' +
      '   ''%d'', ''%d'', ''%s'',                          ' +
      '   ''%s'', ''%f'', ''%d'',                          ' +
      '   ''%s'', ''%d''  )                                ',
      [FCommEnv.DbRetryFreq, FCommEnv.BusyTimeStart,
       FCommEnv.BusyTimeEnd, FCommEnv.BusyTimeReadFreq, FCommEnv.NorTimeReadFreq,
       FCommEnv.DbProcRecords, IIF( FCommEnv.DbWriteError, 'Y', 'N' ), FCommEnv.DbProcIPPV,
       FCommEnv.DbProcBatch, FCommEnv.DbWaterMark,
       FCommEnv.CARetryFreq, FCommEnv.CACheckFreq, IIF( FCommEnv.CAEnableSend, 'Y', 'N' ),
       IIF( FCommEnv.CAEnableRecv, 'Y', 'N' ), FCommEnv.CASendDelay, FCommEnv.CAMaxError,
       FCommEnv.CAProductDefine, FCommEnv.CAIdleTime] );
    DataWriter.ExecSQL;
    FTextSubj.OK( LanguageManager.Get( 'SConfigSaveCommSuccess' ) );
  except
    on E: Exception do
      FTextSubj.ErrorFmt( LanguageManager.Get( 'SConfigSaveCommError' ), [E.Message] );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.SetSoEnv;
var
  aSql: String;
  aIndex: Integer;
begin
  DataWriter.Close;
  DataWriter.SQL.Text := ' delete from soenv ';
  DataWriter.ExecSQL;
  aSql :=
    ' insert into soenv ( SO_TYPE, SO_SELECTED, SO_POS_NAME, ' +
    '   SO_COMP_CODE, SO_COMP_NAME, SO_NETWORK_ID,           ' +
    '   SO_AREA_CODE, SO_LOGINUSER, SO_LOGINPASS,            ' +
    '   SO_DBALIASE )                                        ' +
    ' values ( ''%s'', ''%s'', ''%s'',                       ' +
    '   ''%s'', ''%s'', ''%s'',                              ' +
    '   ''%s'', ''%s'', ''%s'',                              ' +
    '   ''%s''  )                                            ';
  try
    for aIndex := 0 to FSoList.Count - 1 do
    begin
      DataWriter.SQL.Text := Format( aSql,
       [FSoList[aIndex].SoType, IIF( FSoList[aIndex].Selected, 'Y', 'N' ), FSoList[aIndex].PosName,
        FSoList[aIndex].CompCode, FSoList[aIndex].CompName, FSoList[aIndex].NetworkId,
        FSoList[aIndex].AreaCode, FSoList[aIndex].DbUserId, FSoList[aIndex].DbUserPass,
        FSoList[aIndex].DbAliase, FSoList[aIndex]] );
      DataWriter.ExecSQL;
    end;
    FTextSubj.OK( LanguageManager.Get( 'SConfigSaveSoSuccess' ) );
  except
    on E: Exception do
      FTextSubj.ErrorFmt( LanguageManager.Get( 'SConfigSaveSoError' ), [E.Message] );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.SetHighCmdEnv;
var
  aSql: String;
begin
  DataWriter.Close;
  DataWriter.SQL.Text := ' delete from highcmdenv ';
  DataWriter.ExecSQL;
  aSql :=
    ' insert into highcmdenv ( HighLevelCmd, Description,    ' +
    '   LowLevelCmd, CmdType, CmdTypePriority,               ' +
    '   CmdPriority, CmdByNote  )                            ' +
    ' values ( ''%s'', ''%s'',                               ' +
    '   ''%s'', ''%s'', ''%s'',                              ' +
    '   ''%s'', ''%s''  )                                    ';
  HighCmdEnv.First;
  try
    while not HighCmdEnv.Eof do
    begin
      DataWriter.SQL.Text := Format( aSql,
        [HighCmdEnv.FieldByName( 'HighLevelCmd' ).AsString,
         HighCmdEnv.FieldByName( 'Description' ).AsString,
         HighCmdEnv.FieldByName( 'LowLevelCmd' ).AsString,
         HighCmdEnv.FieldByName( 'CmdType' ).AsString,
         HighCmdEnv.FieldByName( 'CmdTypePriority' ).AsString,
         HighCmdEnv.FieldByName( 'CmdPriority' ).AsString,
         HighCmdEnv.FieldByName( 'CmdByNote' ).AsString] );
      DataWriter.ExecSQL;
      HighCmdEnv.Next;
    end;
    FTextSubj.OK( LanguageManager.Get( 'SConfigSaveHighCmdSuccess' ) );
  except
    on E: Exception do
      FTextSubj.ErrorFmt( LanguageManager.Get( 'SConfigSaveHighCmdError' ), [E.Message] );
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.SetLowCmdEnv;
var
  aSql: String;
begin
  DataWriter.Close;
  DataWriter.SQL.Text := ' delete from lowcmdenv ';
  DataWriter.ExecSQL;
  aSql :=
    ' insert into lowcmdenv ( LowLevelCmd, Description ) ' +
    '  values ( ''%s'', ''%s'' )                         ';
  LowCmdEnv.First;
  try
    while not LowCmdEnv.Eof do
    begin
      DataWriter.SQL.Text := Format( aSql,
        [LowCmdEnv.FieldByName( 'LowLevelCmd' ).AsString,
         LowCmdEnv.FieldByName( 'Description' ).AsString] );
      DataWriter.ExecSQL;   
      LowCmdEnv.Next;
    end;
    FTextSubj.OK( LanguageManager.Get( 'SConfigSaveLowCmdSuccess' ) );
  except
    on E: Exception do
      FTextSubj.ErrorFmt( LanguageManager.Get( 'SConfigSaveLowCmdError' ), [E.Message] );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.SetCmdErrEnv;
var
  aSql: String;
begin
  DataWriter.Close;
  DataWriter.SQL.Text := ' delete from cmderrenv ';
  DataWriter.ExecSQL;
  aSql :=
    ' insert into cmderrenv ( ErrorFlag, ErrorCode, ErrorDesc ) ' +
    '  values ( ''%s'', ''%s'', ''%s'' )                        ';
  CmdErrorEnv.First;
  try
    while not CmdErrorEnv.Eof do
    begin
      DataWriter.SQL.Text := Format( aSql,
        [CmdErrorEnv.FieldByName( 'ErrorFlag' ).AsString,
         CmdErrorEnv.FieldByName( 'ErrorCode' ).AsString,
         CmdErrorEnv.FieldByName( 'ErrorDesc' ).AsString] );
      DataWriter.ExecSQL;
      CmdErrorEnv.Next;
    end;
    FTextSubj.OK( LanguageManager.Get( 'SConfigSaveErrCodeSuccess' ) );
  except
    on E: Exception do
      FTextSubj.ErrorFmt( LanguageManager.Get( 'SConfigSaveErrCodeError' ), [E.Message] );
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
