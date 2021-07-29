unit cbLogModule;

interface

uses
  SysUtils, Classes, Variants, DB, ADODB, DBClient, Provider,
  cbClass, cbAppClass;

type
  TLogModule = class(TDataModule)
    AccessConnection: TADOConnection;
    DataReader: TADOQuery;
    ActionDataSet: TClientDataSet;
    DataWriter: TADOQuery;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
    procedure ActionDataSetNewRecord(DataSet: TDataSet);
  private
    { Private declarations }
    FMsgSubject: TMessageSubject;
    FIdGenerator: TIdGenerator;
    FAppLogParam: TAppLogParam;
    function ConnectToAccess(aSilent: Boolean = False): Boolean;
    procedure DisconnectFromAccess;
    procedure PrepareActionDataSet;
    procedure UnPrepareActionDataSet;
    procedure CreateGenerator;
    procedure DestroyGenerator;
    function GetHistoryRecords(var aRecords: Integer): Boolean;
    function SetHistoryRecords(var aRecords: Integer): Boolean;
    function GetLogParam: Boolean;
    procedure SetLogParam;
  public
    { Public declarations }
    function LoadFromFile(aSilent: Boolean = False): Boolean;
    function SaveToFile(aSilent: Boolean = False): Boolean;
    property MessageSubject: TMessageSubject read FMsgSubject;
    property Params: TAppLogParam read FAppLogParam;
  end;

var
  LogModule: TLogModule;

implementation

uses cbUtilis;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TLogModule.DataModuleCreate(Sender: TObject);
begin
  FMsgSubject := TMessageSubject.Create;
  FAppLogParam := TAppLogParam.Create;
end;

{ ---------------------------------------------------------------------------- }

procedure TLogModule.DataModuleDestroy(Sender: TObject);
begin
  DestroyGenerator;
  FMsgSubject.Free;
  FIdGenerator.Free;
end;

{ ---------------------------------------------------------------------------- }

function TLogModule.ConnectToAccess(aSilent: Boolean): Boolean;
const
  aDbPassword = 'cyc84177282';
  aDbFileName = 'Logs.cfg';
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
      raise Exception.CreateFmt( '異動記錄檔%s不存在', [aDbFileName] );
    AccessConnection.Connected := True;
    if not aSilent then
      FMsgSubject.OK( Format( '連結異動記錄檔%s完成。', [aDbFileName] ) );
  except
    on E: Exception do
    begin
      if not aSilent then
        FMsgSubject.Error( Format( '連結異動記錄檔%s失敗, 原因:%s。',
          [aDbFileName, E.Message] ) );
    end;
  end;
  Result := AccessConnection.Connected;
end;

{ ---------------------------------------------------------------------------- }

procedure TLogModule.DisconnectFromAccess;
begin
  if AccessConnection.Connected then
    AccessConnection.Connected := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TLogModule.PrepareActionDataSet;
begin
  ActionDataSet.CreateDataSet;
  ActionDataSet.FieldByName( 'actfilesize' ).Alignment := taRightJustify;
  ActionDataSet.FieldByName( 'actdate' ).Alignment := taCenter;
  ActionDataSet.FieldByName( 'actTime' ).Alignment := taCenter;
end;

{ ---------------------------------------------------------------------------- }

procedure TLogModule.UnPrepareActionDataSet;
begin
  if not VarIsNull( ActionDataSet.Data ) then
    ActionDataSet.EmptyDataSet;
  ActionDataSet.Data := Null;
end;

{ ---------------------------------------------------------------------------- }

procedure TLogModule.CreateGenerator;
var
  aValue: String;
  aCurValue: Integer;
begin
  DataReader.Close;
  DataReader.SQL.Text := ' select max( keyid ) from actionlog ';
  DataReader.Open;
  aValue := DataReader.Fields[0].AsString;
  DataReader.Close;
  aCurValue := 0;
  if ( aValue <> EmptyStr ) then
    aCurValue := StrToInt( Copy( aValue, Length( aValue ) - 6, 7 ) );
  FIdGenerator := TIdGenerator.Create( False,
    TStringGenerateRule.Create( 0, 9999999, aCurValue, gpYearMonthDay ) );
end;

{ ---------------------------------------------------------------------------- }

procedure TLogModule.DestroyGenerator;
begin
  if Assigned( FIdGenerator ) then FIdGenerator.Free;
  FIdGenerator := nil;
end;

{ ---------------------------------------------------------------------------- }

function TLogModule.LoadFromFile(aSilent: Boolean): Boolean;
var
  aRecords: Integer;
begin
  Result := ConnectToAccess( aSilent );
  if Result then
  begin
    try
      Delay( COMM_DELAY * 5 );
      if not aSilent then
      begin
        FMsgSubject.Normal( '讀取參數【LogParams】中' );
        Delay( COMM_DELAY * 5 );
      end;
      Result := GetLogParam;
      if not Result then Exit;
      if not aSilent then
      begin
        FMsgSubject.Normal( '讀取異動記錄中。' );
        Delay( COMM_DELAY * 5 );
      end;
      UnPrepareActionDataSet;
      PrepareActionDataSet;
      DestroyGenerator;
      CreateGenerator;
      Result := GetHistoryRecords( aRecords );
      if not Result then Exit;
      if not aSilent then
      begin
        FMsgSubject.OK( Format( '異動記錄讀取完成, 共 %d 筆記錄。', [aRecords] ) );
        Delay( COMM_DELAY * 5 );
      end;
    finally
      DisconnectFromAccess;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TLogModule.SaveToFile(aSilent: Boolean): Boolean;
var
  aRecords: Integer;
begin
  Result := ConnectToAccess( aSilent );
  if Result then
  begin
    try
      Delay( COMM_DELAY * 5 );
      AccessConnection.BeginTrans;
      try
        if not aSilent then
        begin
          FMsgSubject.Normal( '儲存參數【LogParams】中。' );
          Delay( COMM_DELAY * 5 );
        end;
        SetLogParam;
        AccessConnection.CommitTrans;
      except
        on E: Exception do
        begin
          AccessConnection.RollbackTrans;
          if not aSilent then
            FMsgSubject.Error( Format( '參數【LogParams】儲存失敗, 原因:%s',
              [E.Message] ) );
          Result := False;
          Exit;
        end;
      end;
      if not aSilent then
      begin
        FMsgSubject.Normal( '儲存異動記錄中。' );
        Delay( COMM_DELAY * 5 );
      end;
      Result := SetHistoryRecords( aRecords );
      if not Result then Exit;
      if not aSilent then
      begin
        FMsgSubject.OK( Format( '異動記錄儲存完成, 共 %d 筆記錄。', [aRecords] ) );
        Delay( COMM_DELAY * 5 );
      end;
    finally
      DisconnectFromAccess;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TLogModule.GetHistoryRecords(var aRecords: Integer): Boolean;
var
  aDayText: String;
begin
  Result := True;
  aRecords := 0;
  if ( FAppLogParam.ActLoadDays <= 0 ) then Exit;
  aDayText := FormatDateTime( 'yyyy/mm/dd', ( Date - FAppLogParam.ActLoadDays ) );
  DataReader.Close;
  DataReader.SQL.Text := Format(
    ' select * from actionlog where actdate >= ''%s'' order by keyid',
    [aDayText] );
  try
    DataReader.Open;
    DataReader.First;
    ActionDataSet.DisableControls;
    try
      ActionDataSet.OnNewRecord := nil;
      try
        aRecords := 0;
        while not DataReader.Eof do
        begin
          ActionDataSet.AppendRecord( [
            DataReader.FieldByName( 'keyid' ).AsString,
            DataReader.FieldByName( 'actdate' ).AsString,
            DataReader.FieldByName( 'acttime' ).AsString,
            DataReader.FieldByName( 'actfilename' ).AsString,
            DataReader.FieldByName( 'actfilesize' ).AsString,
            DataReader.FieldByName( 'actfilepath' ).AsString,
            DataReader.FieldByName( 'actsource' ).AsString,
            DataReader.FieldByName( 'actstatus' ).AsString,
            DataReader.FieldByName( 'actcost' ).AsInteger,
            DataReader.FieldByName( 'actprogress' ).AsInteger,
            DataReader.FieldByName( 'acterrmsg' ).AsString,
            EmptyStr] );
          Inc( aRecords );
          DataReader.Next;
        end;
        ActionDataSet.Last;
        DataReader.Close;
      finally
        ActionDataSet.OnNewRecord := ActionDataSetNewRecord;
      end;
    finally
      ActionDataSet.EnableControls;
    end;
  except
    on E: Exception do
    begin
      Result := False;
      FMsgSubject.Error( Format( '讀取異動記錄失敗, 原因:%s。', [E.Message] ) );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TLogModule.SetHistoryRecords(var aRecords: Integer): Boolean;
var
  aStatus: String;
begin
  Result := True;
  ActionDataSet.DisableControls;
  try
    ActionDataSet.First;
    try
      DataWriter.SQL.Text :=
        ' insert into actionlog ( keyid, actdate, acttime,       ' +
        '    actfilename, actfilepath, actsource, actstatus,     ' +
        '    actcost, actprogress, acterrmsg, actfilesize )      ' +
        '  values (  :keyid, :actdate, :acttime,                 ' +
        '    :actfilename, :actfilepath, :actsource, :actstatus, ' +
        '    :actcost, :actprogress, :acterrmsg, :actfilesize )  ';
      aRecords := 0;
      while not ActionDataSet.Eof do
      begin
        if ( ActionDataSet.FieldByName( 'actflag' ).AsString <> EmptyStr ) then
        begin
          DataWriter.Parameters.ParamByName( 'keyid' ).Value :=
            ActionDataSet.FieldByName( 'keyid' ).AsString;
          DataWriter.Parameters.ParamByName( 'actdate' ).Value :=
            ActionDataSet.FieldByName( 'actdate' ).AsString;
          DataWriter.Parameters.ParamByName( 'acttime' ).Value :=
            ActionDataSet.FieldByName( 'acttime' ).AsString;
          DataWriter.Parameters.ParamByName( 'actfilename' ).Value :=
            ActionDataSet.FieldByName( 'actfilename' ).AsString;
          DataWriter.Parameters.ParamByName( 'actfilepath' ).Value :=
            ActionDataSet.FieldByName( 'actfilepath' ).AsString;
          DataWriter.Parameters.ParamByName( 'actsource' ).Value :=
            ActionDataSet.FieldByName( 'actsource' ).AsString;
          aStatus := ActionDataSet.FieldByName( 'actstatus' ).AsString;
          if ( aStatus = 'P' ) then aStatus := EmptyStr;
          DataWriter.Parameters.ParamByName( 'actstatus' ).Value := aStatus;
          DataWriter.Parameters.ParamByName( 'actcost' ).Value :=
            ActionDataSet.FieldByName( 'actcost' ).AsInteger;
          DataWriter.Parameters.ParamByName( 'actprogress' ).Value :=
            ActionDataSet.FieldByName( 'actprogress' ).AsInteger;
          DataWriter.Parameters.ParamByName( 'acterrmsg' ).Value :=
             ActionDataSet.FieldByName( 'acterrmsg' ).AsString;
          DataWriter.Parameters.ParamByName( 'actfilesize' ).Value :=
            ActionDataSet.FieldByName( 'actfilesize' ).AsString;
          DataWriter.ExecSQL;
          ActionDataSet.Edit;
          ActionDataSet.FieldByName( 'actflag' ).AsString := EmptyStr;
          ActionDataSet.Post;
          Inc( aRecords );
        end;
        ActionDataSet.Next;
      end;
    except
      on E: Exception do
      begin
        Result := False;
        FMsgSubject.Error( Format( '儲存異動記錄失敗, 原因:%s。', [E.Message] ) );
      end;
    end;
  finally
    ActionDataSet.EnableControls;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TLogModule.GetLogParam: Boolean;
begin
  Result := True;
  DataReader.Close;
  DataReader.SQL.Text :=  ' select * from params ';
  try
    DataReader.Open;
    DataReader.First;
    FAppLogParam.ActLoadDays := DataReader.FieldByName( 'ActLoadDays' ).AsInteger;
    FAppLogParam.MsgLoadDays := DataReader.FieldByName( 'MsgLoadDays' ).AsInteger;
    DataReader.Close;
  except
    on E: Exception do
    begin
      Result := False;
      FMsgSubject.Error( Format( '參數值【LogParams】讀取失敗, 原因:%s。',
        [E.Message] ) );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TLogModule.SetLogParam;
begin
  DataWriter.Close;
  DataWriter.SQL.Text := ' delete from params ';
  DataWriter.ExecSQL;
  DataWriter.SQL.Text := Format(
    ' insert into params ( actloaddays, msgloaddays ) ' +
    '  values ( ''%d'', ''%d'' )                      ',
    [FAppLogParam.ActLoadDays, FAppLogParam.MsgLoadDays] );
  DataWriter.ExecSQL;  
end;

{ ---------------------------------------------------------------------------- }

procedure TLogModule.ActionDataSetNewRecord(DataSet: TDataSet);
begin
  ActionDataSet.FieldByName( 'keyid' ).Value := FIdGenerator.NextValue;
end;

{ ---------------------------------------------------------------------------- }

end.
