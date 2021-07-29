unit cbConfigModule;

interface

uses
  SysUtils, Classes, Controls, Forms, DB, ADODB, DBClient, IniFiles,
  StdCtrls, Math,
{ App: }
  cbSo, cbSrvClass, cbDesignPattern, cbLanguage,
{ Developer Express: }
  cxTL;  


type
  TConfigModule = class(TDataModule)
    ConfigConnection: TADOConnection;
    DataReader: TADOQuery;
    DataWriter: TADOQuery;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
    FCommEnv: TCommEnv;
    FClientEnv: TClientEnv;
    FSoList: TAppSoList;
    FMsgSubj: TMsgSubject;
    FConfigPath: String;
    FConfigList: TStringList;
    FActiveConfigFile: String;
    function GetConfigFile(aIndex: Integer): String;
    function GetConfigFileCount: Integer;
    function ConnectToAccess(const aConfigName: String): Boolean;
    function GetCommEnv: Boolean;
    function GetClientEnv: Boolean;
    function GetSoEnv: Boolean;
    procedure DisconnectFromAccess;
    procedure SetCommEnv;
    procedure SetClientEnv;
    procedure SetSoEnv;
  public
    { Public declarations }
    procedure RefreshConfigFile;
    procedure SaveActiveConfigFile(aIni: TIniFile);
    procedure RestoreActiveConfigFile(aIni: TIniFile);
    property ConfigFiles[Index: Integer]: String read GetConfigFile;
    property ConfigFileCount: Integer read GetConfigFileCount;
    property ActiveConfigFile: String read FActiveConfigFile write FActiveConfigFile;
    property CommEnv: TCommEnv read FCommEnv;
    property ClientEnv: TClientEnv read FClientEnv;
    property SoList: TAppSoList read FSoList;
    property MsgSub: TMsgSubject read FMsgSubj;
    procedure GetConfigFromValue;
    procedure ConfigSaveToFile;
    function ChangeConfigFile: Boolean;
    function ShowConfigForm: TModalResult;
    function ShowGroupSetForm: TModalResult;
  end;

var
  ConfigModule: TConfigModule;

implementation

{$R *.dfm}

uses cbUtilis, cbMain, cbConfig, cbGroupSet;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.DataModuleCreate(Sender: TObject);
begin
  FConfigPath := IncludeTrailingPathDelimiter( ExtractFilePath(
    ParamStr( 0 ) ) );
  FConfigList := TStringList.Create;
  FActiveConfigFile := EmptyStr;
  FCommEnv := TCommEnv.Create;
  FClientEnv := TClientEnv.Create;
  FSoList := TAppSoList.Create;
  FMsgSubj := TMsgSubject.Create;
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.DataModuleDestroy(Sender: TObject);
begin
  FMsgSubj.Free;
  FSoList.Free;
  FCommEnv.Free;
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
    FMsgSubj.InfoFmt( LanguageManager.Get( 'SConfigSearching' ), [aSearchRect.Name] );
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
    FMsgSubj.ErrorFmt( LanguageManager.Get( 'SConfigFileNotExists' ),
      [Nvl( FActiveConfigFile, LanguageManager.Get( 'SNull' ) )] );
  end;
  {}
  if ( Result ) then Result := ConnectToAccess( FActiveConfigFile );
  if ( Result ) then
  begin
    FMsgSubj.InfoFmt( LanguageManager.Get( 'SConfigBeginRead' ),
       [FActiveConfigFile] );
    Delay( 500 );
    Result := ( Result and GetCommEnv );
    Delay( 500 );
    Result := ( Result and GetClientEnv );
    Delay( 500 );
    Result := ( Result and GetSoEnv );
    Delay( 500 );
  end;
  if ( Result ) then
    FMsgSubj.OK( LanguageManager.Get( 'SConfigEndRead' ) )
  else
    FMsgSubj.Error( LanguageManager.Get( 'SConfigLoadWarning' ) );
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
    FMsgSubj.InfoFmt( LanguageManager.Get( 'SConfigOpenSuccess' ), [aFileName] )
  else
    FMsgSubj.ErrorFmt( LanguageManager.Get( 'SConfigOpenError' ), [aFileName, aMsg] );
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.DisconnectFromAccess;
begin
  if ConfigConnection.Connected then ConfigConnection.Close;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.GetCommEnv: Boolean;
begin
  Result := False;
  FCommEnv.DbRetryFreq := 60;
  try
    DataReader.Close;
    DataReader.SQL.Text := ' select * from commenv ';
    DataReader.Open;
    DataReader.First;
    if ( DataReader.IsEmpty ) then
    begin
      FMsgSubj.ErrorFmt( LanguageManager.Get( 'SConfigCommReadError' ),
       [LanguageManager.Get( 'SNull' )] );                       
    end else
    begin
      {}
      FCommEnv.DbRetryFreq := DataReader.FieldByName( 'dbretryfreq' ).AsInteger;
      FMsgSubj.InfoFmt( LanguageManager.Get( 'SConfigCommDbRetryFreq' ), [FCommEnv.DbRetryFreq] );
      Delay( 100 );
      {}
      FCommEnv.DbSyncUser := DataReader.FieldByName( 'dbsyncuser' ).AsInteger;
      FMsgSubj.InfoFmt( LanguageManager.Get( 'SConfigCommDbSyncUser' ), [FCommEnv.DbSyncUser] );
      Delay( 100 );
      {}
      FCommEnv.DbVaildateSessionFreq := Nvl( DataReader.FieldByName( 'DbValidateSessionFreq' ).AsString, 60 );
      FMsgSubj.InfoFmt( LanguageManager.Get( 'SConfigCommDbValidateSessionFreq' ), [FCommEnv.DbVaildateSessionFreq] );
      Delay( 100 );
      {}
      FCommEnv.DbGetPoolObjectRertyCount := Nvl( DataReader.FieldByName( 'DbGetPoolObjectRetryCount' ).AsString, 10 );
      FMsgSubj.InfoFmt( LanguageManager.Get( 'SConfigCommDbGetPoolObjectRetryCount' ), [FCommEnv.DbGetPoolObjectRertyCount] );
      Delay( 100 );
      {}
      FCommEnv.UIGetUserListFreq := Nvl( DataReader.FieldByName( 'UIGetUserListFreq' ).AsString, 60 );
      FMsgSubj.InfoFmt( LanguageManager.Get( 'SConfigCommUIGetUser' ), [FCommEnv.UIGetUserListFreq] );
      Delay( 100 );
      {}
      Result := True;
      FMsgSubj.OK( LanguageManager.Get( 'SConfigCommReadSuccess' ) );
    end;
  except
    on E: Exception do
    begin
      FMsgSubj.ErrorFmt( LanguageManager.Get( 'SConfigCommReadError' ), [E.Message] );
    end;
  end;
  DataReader.Close;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.GetClientEnv: Boolean;
begin
  Result := False;
  FClientEnv.AutoRefresh := True;
  FClientEnv.AuthorizeRefreshRate := 60;
  FClientEnv.AnnRefreshRate := 120;
  FClientEnv.UserRefreshRate := 10;
  try
    DataReader.Close;
    DataReader.SQL.Text := ' select * from clientenv ';
    DataReader.Open;
    DataReader.First;
    if ( DataReader.IsEmpty ) then
    begin
      FMsgSubj.ErrorFmt( LanguageManager.Get( 'SConfigClientReadError' ),
       [LanguageManager.Get( 'SNull' )] );
    end else
    begin
      {}
      FClientEnv.AutoRefresh := ( DataReader.FieldByName( 'autorefresh' ).AsString = '1' );
      FMsgSubj.InfoFmt( LanguageManager.Get( 'SConfigClientAutoRefresh' ), [
       IIF( FClientEnv.AutoRefresh, LanguageManager.Get( 'SYes' ), LanguageManager.Get( 'SNo' ) )] );
      Delay( 100 );
      {}
      FClientEnv.AuthorizeRefreshRate := DataReader.FieldByName( 'AuthorizeRefreshRate' ).AsInteger;
      FMsgSubj.InfoFmt( LanguageManager.Get( 'SConfigClientAuthorizeRefreshRate' ), [FClientEnv.AuthorizeRefreshRate] );
      Delay( 100 );
      {}
      FClientEnv.AnnRefreshRate := DataReader.FieldByName( 'AnnRefreshRate' ).AsInteger;
      FMsgSubj.InfoFmt( LanguageManager.Get( 'SConfigClientAnnRefreshRate' ), [FClientEnv.AnnRefreshRate] );
      Delay( 100 );
      {}
      FClientEnv.UserRefreshRate := DataReader.FieldByName( 'UserRefreshRate' ).AsInteger;
      FMsgSubj.InfoFmt( LanguageManager.Get( 'SConfigClientUserRefreshRate' ), [FClientEnv.UserRefreshRate] );
      Delay( 100 );
      {}
      FClientEnv.TryReconnectRate := DataReader.FieldByName( 'TryReconnectRate' ).AsInteger;
      FMsgSubj.InfoFmt( LanguageManager.Get( 'SConfigClientTryReconnectRate' ), [FClientEnv.UserRefreshRate] );
      Delay( 100 );
      {}
      Result := True;
      FMsgSubj.OK( LanguageManager.Get( 'SConfigClientReadSuccess' ) );
    end;
  except
    on E: Exception do
    begin
      FMsgSubj.ErrorFmt( LanguageManager.Get( 'SConfigClientReadError' ), [E.Message] );
    end;
  end;
  DataReader.Close;
end;

{ ---------------------------------------------------------------------------- }

function TConfigModule.GetSoEnv: Boolean;
var
  aSo: TAppSo;
begin
  Result := False;
  FSoList.Clear;
  try
    DataReader.Close;
    DataReader.SQL.Text := ' select * from soenv ';
    DataReader.Open;
    DataReader.First;
    while not DataReader.Eof do
    begin
      aSo := TAppSo.Create;
      aSo.Selected := ( DataReader.FieldByName( 'selected' ).AsString = '1' );
      aSo.CompCode := DataReader.FieldByName( 'compcode' ).AsString;
      aSo.CompName := DataReader.FieldByName( 'compname' ).AsString;
      aSo.DbUserId := DataReader.FieldByName( 'dbusername' ).AsString;
      aSo.DbUserPass := DataReader.FieldByName( 'dbpassword' ).AsString;
      aSO.DbAliase := DataReader.FieldByName( 'dbaliase' ).AsString;
      aSo.SynData := ( DataReader.FieldByName( 'syncdata' ).AsString = '1' );
      aSo.MaxSession := Nvl( DataReader.FieldByName( 'MaxSession' ).AsString, 0 );
      aSo.DbState := dbClose;
      if ( aSo.Selected or aSo.SynData ) and ( aSo.MaxSession <= 0 ) then
        aSo.DbState := dbWarning;
      FSoList.Add( aSo );
      DataReader.Next;
      FMsgSubj.InfoFmt( LanguageManager.Get( 'SConfigSoRead' ),
        [aSo.CompName] );
      Delay( 100 );
    end;
    if ( FSoList.Count > 0 ) then
    begin
      Result := True;
      FMsgSubj.OKFmt( LanguageManager.Get( 'SConfigSoReadSuccess' ), [FSoList.Count] );
    end  else
      FMsgSubj.ErrorFmt( LanguageManager.Get( 'SConfigSoReadError' ),
        [LanguageManager.Get( 'SNull' )] )
  except
    on E: Exception do
    begin
      FMsgSubj.ErrorFmt( LanguageManager.Get( 'SConfigSoReadError' ), [E.Message] );
      FSoList.Clear;
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
    fmConfig.Caption := LanguageManager.Get( fmConfig.Name );
    fmConfig.btnConfirm.Enabled := ( AppRunState = arIdle );
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

function TConfigModule.ShowGroupSetForm: TModalResult;
begin
  fmGroupSet := TfmGroupSet.Create( nil );
  try
    fmGroupSet.btnConfirm.Enabled := ( AppRunState = arIdle );
    fmGroupSet.ShowModal;
  finally
    fmGroupSet.Free;
  end;
  Result := mrOk;
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.GetConfigFromValue;
var
  aIndex: Integer;
  aSo: TAppSo;
  aNode: TcxTreeListNode;
begin
  FCommEnv.DbVaildateSessionFreq := fmConfig.txtDbValidateSessionFreq.Value;
  FCommEnv.DbGetPoolObjectRertyCount := fmConfig.txtDbGetPoolObjectRetryCount.Value;
  FCommEnv.DbSyncUser := fmConfig.txtDbSyncUser.Value;
  FCommEnv.UIGetUserListFreq := fmConfig.txtUIGetUserListFreq.Value;
  {}
  FClientEnv.AutoRefresh := ( fmConfig.txtAutoRefresh.ItemIndex = 0 );
  FClientEnv.AuthorizeRefreshRate := fmConfig.txtAuthorizeRefreshRate.Value;
  FClientEnv.AnnRefreshRate := fmConfig.txtAnnRefreshRate.Value;
  FClientEnv.UserRefreshRate := fmConfig.txtUserRefreshRate.Value;
  FClientEnv.TryReconnectRate := fmConfig.txtTryReconnectRate.Value;
  {}
  FSoList.Clear;
  for aIndex := 0 to fmConfig.SoTree.Count - 1 do
  begin
    aNode := fmConfig.SoTree.Nodes[aIndex];
    aSo := TAppSo.Create;
    aSo.Selected := ( aNode.Values[fmConfig.SoTreeCol1.ItemIndex] = True );
    aSo.CompCode := aNode.Texts[fmConfig.SoTreeCol2.ItemIndex];
    aSo.CompName := aNode.Texts[fmConfig.SoTreeCol3.ItemIndex];
    aSo.DbUserId := aNode.Texts[fmConfig.SoTreeCol4.ItemIndex];
    aSo.DbUserPass := aNode.Texts[fmConfig.SoTreeCol5.ItemIndex];
    aSo.DbAliase :=  aNode.Texts[fmConfig.SoTreeCol6.ItemIndex];
    aSo.SynData := ( aNode.Values[fmConfig.SoTreeCol7.ItemIndex] = True );
    aSo.MaxSession := aNode.Values[fmConfig.SoTreeCol8.ItemIndex];
    aSo.DbState := dbClose;
    FSoList.Add( aSo );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.ConfigSaveToFile;
begin
  if ConnectToAccess( FActiveConfigFile ) then
  begin
    SetCommEnv;
    SetClientEnv;
    SetSoEnv;
  end;
  DisconnectFromAccess;
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.SetCommEnv;
begin
  DataWriter.Close;
  DataWriter.SQL.Text := ' delete from commenv ';
  DataWriter.ExecSQL;
  try
    DataWriter.SQL.Text := Format(
      ' insert into commenv ( DbRetryFreq, DbSyncUser, DbValidateSessionFreq, ' +
      '    DbGetPoolObjectRetryCount, UIGetUserListFreq )    ' +
      ' values ( ''%d'', ''%d'', ''%d'', ''%d'', ''%d'' )    ',
      [FCommEnv.DbRetryFreq, FCommEnv.DbSyncUser, FCommEnv.DbVaildateSessionFreq,
       FCommEnv.DbGetPoolObjectRertyCount, FCommEnv.UIGetUserListFreq] );
    DataWriter.ExecSQL;
    FMsgSubj.OK( LanguageManager.Get( 'SConfigSaveCommSuccess' ) );
  except
    on E: Exception do
      FMsgSubj.ErrorFmt( LanguageManager.Get( 'SConfigSaveCommError' ), [E.Message] );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TConfigModule.SetClientEnv;
begin
  DataWriter.Close;
  DataWriter.SQL.Text := ' delete from clientenv ';
  DataWriter.ExecSQL;
  try
    DataWriter.SQL.Text := Format(
      ' insert into clientenv ( AutoRefresh, AuthorizeRefreshRate,  ' +
      '    AnnRefreshRate, UserRefreshRate, TryReconnectRate )      ' +
      ' values ( ''%d'', ''%d'', ''%d'', ''%d'', ''%d'' )           ',
      [Ord( FClientEnv.AutoRefresh ), FClientEnv.AuthorizeRefreshRate,
       FClientEnv.AnnRefreshRate, FClientEnv.UserRefreshRate,
       FClientEnv.TryReconnectRate] );
    DataWriter.ExecSQL;
    FMsgSubj.OK( LanguageManager.Get( 'SConfigSaveClientSuccess' ) );
  except
    on E: Exception do
      FMsgSubj.ErrorFmt( LanguageManager.Get( 'SConfigSaveClientError' ), [E.Message] );
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
    ' insert into soenv (      ' +
    '   selected,              ' +
    '   compcode,              ' +
    '   compname,              ' +
    '   dbusername,            ' +
    '   dbpassword,            ' +
    '   dbaliase,              ' +
    '   syncdata,              ' +
    '   maxsession )           ' +
    ' values ( ''%d'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%d'', ' +
    '          ''%d''  ) ';
  try
    for aIndex := 0 to FSoList.Count - 1 do
    begin
      DataWriter.SQL.Text := Format( aSql,
       [Ord( FSoList[aIndex].Selected ),
        FSoList[aIndex].CompCode,
        FSoList[aIndex].CompName,
        FSoList[aIndex].DbUserId,
        FSoList[aIndex].DbUserPass,
        FSoList[aIndex].DbAliase,
        Ord( FSoList[aIndex].SynData ),
        FSoList[aIndex].MaxSession] );
      DataWriter.ExecSQL;
    end;
    FMsgSubj.OK( LanguageManager.Get( 'SConfigSaveSoSuccess' ) );
  except
    on E: Exception do
      FMsgSubj.ErrorFmt( LanguageManager.Get( 'SConfigSaveSoError' ), [E.Message] );
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
