unit cbMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Buttons, ComCtrls, IniFiles, cbClass,
  cxEdit, ImgList, cxControls, cxContainer, cxListView, IdMessage,
  IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient,
  IdMessageClient, IdSMTP;

type
  TMain = class(TForm)
    Panel1: TPanel;
    btnStart: TSpeedButton;
    btnStop: TSpeedButton;
    SpeedButton2: TSpeedButton;
    ImageList1: TImageList;
    Panel2: TPanel;
    Panel3: TPanel;
    Label1: TLabel;
    DisplayTimer: TTimer;
    Label2: TLabel;
    Panel4: TPanel;
    MsgListView: TcxListView;
    btnOption: TSpeedButton;
    cxDefaultEditStyleController1: TcxDefaultEditStyleController;
    IdSMTP1: TIdSMTP;
    IdMessage1: TIdMessage;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
    procedure btnStartClick(Sender: TObject);
    procedure btnStopClick(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure DisplayTimerTimer(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
  private
    { Private declarations }
    FNotify: TNotify;
    FFtpThread: TThread;
    FDataThread: TThread;
    FWatchDogThread: TThread;
    FCA: TFreq;
    FSRA: TFreq;
    FFTP: TFTP;
    FEmailAccount: TEmailAccount;
    FEmailList: TStringList;
    procedure LoadConfig;
    procedure SaveConfig(const aKind: Integer);
  public
    { Public declarations }
    procedure AddMsg(aMsg: TNotify);
    property DataThread: TThread read FDataThread;
    property FtpThread: TThread read  FFtpThread;
  end;

var
  Main: TMain;

  LocalTempDir: String;
  LocalUploadDir: String;
  LocalBackupDir: String;
  LocalDownloadDir: String;
  LocalErrorDir: String;

  FtpEmcTransDir: String;
  FtpEmcUploadDir: String;
  FtpApolUploadDir: String;


implementation

uses cbDataThread, cbFtpThread, cbDataControler, cbWatchDogThread, Encryption_TLB;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

function ExtractValue(var aValue: String; aSeparator: String = ','): String;
var
  aPos: Integer;
begin
  aPos := AnsiPos( aSeparator, aValue );
  if aPos = 0 then
  begin
    Result := aValue;
    aValue := EmptyStr;
  end else
  begin
    Result := Copy( aValue, 1, aPos - 1 );
    Delete( aValue, 1, aPos );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TMain.FormCreate(Sender: TObject);
begin
  LocalTempDir := IncludeTrailingPathDelimiter( ExtractFilePath(
    ParamStr( 0 ) ) ) + 'Temp';
  LocalUploadDir := IncludeTrailingPathDelimiter( ExtractFilePath(
    ParamStr( 0 ) ) ) + 'Upload';
  LocalBackupDir := IncludeTrailingPathDelimiter( ExtractFilePath(
    ParamStr( 0 ) ) ) + 'Backup';
  LocalDownloadDir := IncludeTrailingPathDelimiter( ExtractFilePath(
    ParamStr( 0 ) ) ) + 'Download';
  LocalErrorDir := IncludeTrailingPathDelimiter( ExtractFilePath(
    ParamStr( 0 ) ) ) + 'Error';

  FtpEmcTransDir := 'EMCTrans';
  FtpEmcUploadDir := 'EMCUpload';
  FtpApolUploadDir := 'APOLUpload';

  FEmailAccount.Account := EmptyStr;
  FEmailAccount.Password := EmptyStr;
  FEmailAccount.Server := EmptyStr;
  
  FEmailList := TStringList.Create;

  LoadConfig;

  if not FileExists( IncludeTrailingPathDelimiter( ExtractFilePath(
    ParamStr( 0 ) ) ) + 'CSIS.INI' ) then
  begin
    Application.MessageBox( PChar( '找不到指定的檔案, CSIS.ini。' ), '錯誤', MB_OK + MB_ICONERROR );
    Application.ShowMainForm := False;
    Application.Terminate;
    Exit;
  end;

  TDataControler.CheckFolder;

  FNotify.MsgType := MB_ICONINFORMATION;
  FNotify.MsgText := '啟動';
  AddMsg( FNotify );
  
    
end;

{ ---------------------------------------------------------------------------- }

procedure TMain.FormShow(Sender: TObject);
begin
  DisplayTimerTimer( DisplayTimer );
end;

{ ---------------------------------------------------------------------------- }

procedure TMain.FormResize(Sender: TObject);
begin
  MsgListView.Columns[0].Width := MsgListView.Width + 50;
end;

{ ---------------------------------------------------------------------------- }

procedure TMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if btnStart.Enabled then
  begin
    Action := caFree;
    Exit;
  end;    
  if Application.MessageBox( PChar( '確認結束程式?' ), PChar( '確認' ),
    MB_OK + MB_ICONQUESTION ) = ID_OK then
  begin
    try
      btnStop.Click;
    except
      {}
    end;
    Action := caFree;
  end else
  begin
    Action := caNone;
  end;    
end;

{ ---------------------------------------------------------------------------- }

procedure TMain.FormDestroy(Sender: TObject);
begin
  DataControler.Free;
  FEmailList.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TMain.AddMsg(aMsg: TNotify);
var
  aItem: TListItem;
begin
  aItem := MsgListView.Items.Add;
  case aMsg.MsgType of
    MB_ICONERROR: aItem.ImageIndex := 1;
    MB_ICONWARNING: aItem.ImageIndex := 3;
    MB_OK: aItem.ImageIndex := 0;
  else
    aItem.ImageIndex := 2;
  end;
  aItem.Caption := Format( '%s>   %s', [
    FormatDateTime( 'yyyy-mm-dd hh:nn:ss', Now ), aMsg.MsgText] );
  aItem.MakeVisible( True );
  aItem.Selected := True;  
  Application.ProcessMessages;  
end;

{ ---------------------------------------------------------------------------- }

procedure TMain.btnStartClick(Sender: TObject);
begin
  btnStart.Enabled := False;
  FNotify.MsgType := MB_ICONINFORMATION;
  FNotify.MsgText := '執行。';
  AddMsg( FNotify );
  FFtpThread := TFtpThread.Create( FFTP.Host, FFTP.UserId, FFTP.Password );
  FFtpThread.Resume;
  {}
  FDataThread := TDataThread.Create( FFTP.Comps );
  TDataThread( FDataThread ).CA := FCA;
  TDataThread( FDataThread ).SRA := FSRA;
  FDataThread.Resume;
  {}
  FWatchDogThread := TWatchDogThread.Create;
  TWatchDogThread( FWatchDogThread ).EMailList := FEmailList;
  TWatchDogThread( FWatchDogThread ).EMailAccount := FEmailAccount;
  TWatchDogThread( FWatchDogThread ).Resume;
end;

{ ---------------------------------------------------------------------------- }

procedure TMain.btnStopClick(Sender: TObject);
begin
  if btnStart.Enabled then Exit;
  try
    FNotify.MsgType := MB_ICONINFORMATION;
    FNotify.MsgText := '停止FTP上/下傳執行緒中.....';
    AddMsg( FNotify );
    Application.ProcessMessages;
    {}
    FFtpThread.Terminate;
    FFtpThread.WaitFor;
    FFtpThread.Free;
    FFtpThread := nil;
    {}
    FNotify.MsgType := MB_ICONINFORMATION;
    FNotify.MsgText := '停止資料庫執行緒中.....';
    AddMsg( FNotify );
    Application.ProcessMessages;
    {}
    FDataThread.Terminate;
    FDataThread.WaitFor;
    FDataThread.Free;
    FDataThread := nil;
    {}
    FNotify.MsgType := MB_ICONINFORMATION;
    FNotify.MsgText := '停止監控執行緒中.....';
    AddMsg( FNotify );
    Application.ProcessMessages;
    {}
    FWatchDogThread.Terminate;
    FWatchDogThread.WaitFor;
    FWatchDogThread.Free;
    FWatchDogThread := nil;
    {}
    FNotify.MsgType := MB_ICONINFORMATION;
    FNotify.MsgText := '停止。';
    AddMsg( FNotify );
  except
    on E: Exception do
    begin
      Application.MessageBox( PChar( E.Message ), '錯誤', MB_OK + MB_ICONERROR );
    end;
  end;
  btnStart.Enabled := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TMain.LoadConfig;
var
  aIni: TIniFile;
  aTemp: String;
  aIndex, aPos: Integer;
begin
  TDataControler.CreateTemporyIniFile(
    IncludeTrailingPathDelimiter( ExtractFilePath(
      ParamStr( 0 ) ) ) + 'CONFIG.INI',
    IncludeTrailingPathDelimiter( ExtractFilePath(
      ParamStr( 0 ) ) )+ 'TMPCONFIG.INI' );
  aIni := TIniFile.Create( IncludeTrailingPathDelimiter( ExtractFilePath(
      ParamStr( 0 ) ) )+ 'TMPCONFIG.INI' );
  try
    FFTP.UserId := aIni.ReadString( 'FTP', 'USERID', EmptyStr );
    FFTP.Password := aIni.ReadString( 'FTP', 'PASSWORD', EmptyStr );
    FFTP.Host := aIni.ReadString( 'FTP', 'HOST', EmptyStr );
    FFTP.Comps := aIni.ReadString( 'FTP', 'COMPS', EmptyStr );
    {}
    FCA.CFR := aIni.ReadInteger( 'CAFRQUENCE', 'FR', 60 );
    FCA.LastExc := StrToDateTime( aIni.ReadString( 'CAFRQUENCE', 'LASTEXC',
      FormatDateTime( 'yyyy/mm/dd hh:nn', Now )  )  );
    FCA.NextExc := ( FCA.LastExc + ( ( 1 / 24 / 60 ) * ( FCA.CFR ) ) );
    {}
    FSRA.CFR := aIni.ReadInteger( 'SRAFRQUENCE', 'FR', 60 );
    FSRA.LastExc := StrToDateTime( aIni.ReadString( 'SRAFRQUENCE', 'LASTEXC',
      FormatDateTime( 'yyyy/mm/dd hh:nn', Now )  )  );
    FSRA.NextExc := ( FCA.LastExc + ( ( 1 / 24 / 60 ) * ( FSRA.CFR ) ) );
    {}
    aTemp := aIni.ReadString( 'SENDER', 'ACCOUNT', EmptyStr );
    if ( aTemp <> EmptyStr ) then
    begin
      FEmailAccount.Account := ExtractValue( aTemp, '/' );
      FEmailAccount.Password := ExtractValue( aTemp, ':' );
      FEmailAccount.Server := ExtractValue( aTemp, ':' );
      FEmailAccount.From := aTemp;
    end;
    FEmailList.Clear;
    aIni.ReadSectionValues( 'EMAIL-LIST', FEmailList );
    for aIndex := 0 to FEmailList.Count - 1 do
    begin
      aPos := Pos( '=', FEmailList[aIndex] );
      FEmailList[aIndex] := Copy( FEmailList[aIndex], aPos + 1,
        Length(  FEmailList[aIndex] ) - aPos );
    end;
  finally
    aIni.Free;
  end;
  if FileExists( IncludeTrailingPathDelimiter( ExtractFilePath(
      ParamStr( 0 ) ) )+ 'TMPCONFIG.INI' ) then
    DeleteFile( IncludeTrailingPathDelimiter( ExtractFilePath(
      ParamStr( 0 ) ) )+ 'TMPCONFIG.INI' );
end;

{ ---------------------------------------------------------------------------- }

procedure TMain.DisplayTimerTimer(Sender: TObject);
begin
  Label1.Caption := FormatDateTime( 'yyyy-mm-dd hh:nn:ss', Now );
end;

{ ---------------------------------------------------------------------------- }

procedure TMain.SpeedButton2Click(Sender: TObject);
begin
  btnStop.Click;
  Self.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TMain.SaveConfig(const aKind: Integer);
var
  aIndex: Integer;
  aIni: TIniFile;
  aTmp: TStringList;
  aObj: _Password;
  aKey: WideString;
begin
  TDataControler.CreateTemporyIniFile(
    IncludeTrailingPathDelimiter( ExtractFilePath(
      ParamStr( 0 ) ) ) + 'CONFIG.INI',
    IncludeTrailingPathDelimiter( ExtractFilePath(
      ParamStr( 0 ) ) )+ 'TMPCONFIG.INI' );
  aIni := TIniFile.Create( IncludeTrailingPathDelimiter( ExtractFilePath(
      ParamStr( 0 ) ) )+ 'TMPCONFIG.INI' );
  try
    if ( aKind in [0,1] ) then
    begin
      aIni.WriteString( 'FTP', 'HOST', FFTP.Host );
      aIni.WriteString( 'FTP', 'USERID', FFTP.UserId );
      aIni.WriteString( 'FTP', 'PASSWORD', FFTP.UserId );
      aIni.WriteString( 'FTP', 'PASSWORD', FFTP.Comps );
    end;
    {}
    if ( aKind in [0,2]) then
    begin
      aIni.WriteInteger( 'CAFRQUENCE', 'FR', FCA.CFR  );
      aIni.WriteString( 'CAFRQUENCE', 'LASTEXC', FormatDateTime( 'yyyy/mm/dd hh:nn:ss', FCA.LastExc ) );
    end;
    {}
    if ( aKind in [0,3] ) then
    begin
      aIni.WriteInteger( 'SRAFRQUENCE', 'FR', FSRA.CFR  );
      aIni.WriteString( 'SRAFRQUENCE', 'LASTEXC', FormatDateTime( 'yyyy/mm/dd hh:nn:ss', FSRA.LastExc ) );
    end;
    {}
    aIni.UpdateFile;
  finally
    aIni.Free;
  end;
  aTmp := TStringList.Create;
  try
    aTmp.LoadFromFile( IncludeTrailingPathDelimiter( ExtractFilePath(
      ParamStr( 0 ) ) )+ 'TMPCONFIG.INI' );
    aObj := CoPassword.Create;
    try
      aKey := 'CS';
      for aIndex := 0 to aTmp.Count - 1 do
        aTmp[aIndex] := aObj.Encrypt( aTmp[aIndex], aKey );
      aTmp.SaveToFile( IncludeTrailingPathDelimiter( ExtractFilePath(
        ParamStr( 0 ) ) )+ 'CONFIG.INI' );
    finally
      aObj := nil;
    end;
  finally
    aTmp.Free;
  end;
  if FileExists( IncludeTrailingPathDelimiter( ExtractFilePath(
      ParamStr( 0 ) ) )+ 'TMPCONFIG.INI' ) then
    DeleteFile( IncludeTrailingPathDelimiter( ExtractFilePath(
      ParamStr( 0 ) ) )+ 'TMPCONFIG.INI' );
end;

{ ---------------------------------------------------------------------------- }

end.
