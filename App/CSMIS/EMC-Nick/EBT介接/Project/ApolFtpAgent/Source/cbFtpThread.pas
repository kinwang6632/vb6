unit cbFtpThread;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, cbClass,
  IdTCPConnection, IdTCPClient, IdFTP, IdFTPList;

type
  TFtpThread = class(TThread)
  private
    { Private declarations }
    FNotify: TNotify;
    FLastExc: TDateTime;
    function LocalFileDetect: Integer;
    function FtpFileDetect: Integer;
    function DoFTPConnect: Boolean;
    procedure DoFTPDisconnect;
    procedure DoFTPUpload;
    procedure DoFTPDownlaod;
    procedure Notify;
    function CheckCanExc: Boolean;
  protected
    procedure Execute; override;
  public
    constructor Create(FtpIp, UserId, Password: String);
    destructor Destroy; override;
    procedure OnDownloadFile(const aErrorText, aFileName: String);
  end;

implementation

uses cbMain, cbDataControler, cbFTPControler, DateUtils;

{ Important: Methods and properties of objects in visual components can only be
  used in a method called using Synchronize, for example,

      Synchronize(UpdateCaption);

  and UpdateCaption could look like,

    procedure TFtpThread.UpdateCaption;
    begin
      Form1.Caption := 'Updated in a thread';
    end; }

{ TFtpThread }

{ ---------------------------------------------------------------------------- }

constructor TFtpThread.Create(FtpIp, UserId, Password: String);
begin
  inherited Create( True );
  Self.FreeOnTerminate := False;
  FTPControler := TFTPControler.Create( nil );
  FTPControler.IdFTP.Host := FtpIp;
  FTPControler.IdFTP.Username := UserId;
  FTPControler.IdFTP.Password := Password;
  FTPControler.IdFTP.Passive := False;
  FTPControler.IdFTP.ReadTimeout := 60000;
  FTPControler.OnDownloadStatus := Self.OnDownloadFile;
  FLastExc := -1;
end;

{ ---------------------------------------------------------------------------- }

destructor TFtpThread.Destroy;
begin
  FTPControler.Free;
  FTPControler := nil;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

function TFtpThread.DoFTPConnect: Boolean;
begin
  if not FTPControler.Active then
  begin
    FNotify.MsgType := MB_ICONINFORMATION;
    FNotify.MsgText := Format( 'FTP Server: %s 連結中。', [FTPControler.IdFTP.Host] );
    Synchronize( Notify );
    Sleep( 300 );
    FTPControler.Active := True;
    if FTPControler.Active then
    begin
      FNotify.MsgType := MB_ICONINFORMATION;
      FNotify.MsgText := Format( 'FTP Server: %s 成功登入。',[FTPControler.IdFTP.Host] );
      Synchronize( Notify );
      Sleep( 300 );
    end else
    begin
      FNotify.MsgType := MB_ICONERROR;
      FNotify.MsgText := Format( 'FTP Server: %s 連結失敗, 原因:%s。',
        [FTPControler.IdFTP.Host, FTPControler.ErrorText] );
      Synchronize( Notify );
      Sleep( 300 );
    end;
  end;
  Result := FTPControler.IdFTP.Connected;
end;

{ ---------------------------------------------------------------------------- }

procedure TFtpThread.DoFTPDisconnect;
begin
  if FTPControler.Active then
  begin
    FNotify.MsgType := MB_ICONINFORMATION;
    FNotify.MsgText := Format( '與 FTP Server: %s 中斷連線。', [FTPControler.IdFTP.Host] );
    Synchronize( Notify );
    Sleep( 300 );
    FTPControler.Active := False;
    FNotify.MsgType := MB_ICONINFORMATION;
    FNotify.MsgText := Format( '與 FTP Server: %s 連結已關閉。', [FTPControler.IdFTP.Host] );
    Synchronize( Notify );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TFtpThread.DoFTPUpload;
var
  aSearch: TSearchRec;
  aSearchMask, aSource, aDest: String;
  aUploadResult: Boolean;
begin
  aSearchMask := IncludeTrailingPathDelimiter( LocalUploadDir ) + '*.txt';
  if FindFirst( aSearchMask, faAnyFile, aSearch ) = 0 then
  begin
    try
      FTPControler.SwitchToFTPTransFolder;
      repeat
         if ( aSearch.Attr and faDirectory ) <> 0 then
         begin
           if ( aSearch.Name = '.' ) or ( aSearch.Name = '..' ) then Continue;
         end else
         if ( aSearch.Attr and faArchive ) <> 0 then
         begin
           FNotify.MsgType := MB_ICONINFORMATION;
           FNotify.MsgText := Format( '檔案:%s上傳中。', [aSearch.Name] );
           Synchronize( Notify );
           Sleep( 300 );
           aUploadResult := FTPControler.UploadFile(
             IncludeTrailingPathDelimiter( LocalUploadDir ) + aSearch.Name );
           if aUploadResult then
           begin
             aSource := IncludeTrailingPathDelimiter( LocalUploadDir ) + aSearch.Name;
             aDest := IncludeTrailingPathDelimiter( LocalBackupDir ) + aSearch.Name;
             MoveFile( PChar( aSource ), PChar( aDest ) );
             FNotify.MsgType := MB_ICONINFORMATION;
             FNotify.MsgText := Format( '檔案:%s已上傳完成。', [aSearch.Name] );
           end else
           begin
             FNotify.MsgType := MB_ICONERROR;
             FNotify.MsgText := Format( '檔案:%s上傳失敗, 原因:%s。',
               [aSearch.Name, FTPControler.ErrorText] );
           end;
           Synchronize( Notify );
           Sleep( 300 );
         end;
      until FindNext( aSearch ) <> 0;
    finally
      FindClose( aSearch );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TFtpThread.Execute;
var
  aFiles: Integer;
begin
  FNotify.MsgType := MB_ICONINFORMATION;
  FNotify.MsgText := 'FTP檔案上/下傳執行緒開始執行。';
  Synchronize( Notify );
  Sleep( 10000 );
  while not Self.Terminated do
  begin
    if Self.Terminated then Break;
    Sleep( 10000 );
    if Self.Terminated then Break;
    try
      {}
      aFiles := LocalFileDetect;
      if ( aFiles > 0 ) then
      begin
        FNotify.MsgType := MB_ICONINFORMATION;
        FNotify.MsgText := Format( '上傳資料夾共有%d檔案, 準備上傳。', [aFiles] );
        Synchronize( Notify );
        if DoFTPConnect then
        begin
          DoFTPUpload;
        end;
        DoFTPDisconnect;
      end;
    except
      on E: Exception do
      begin
        FNotify.MsgType := MB_ICONERROR;
        FNotify.MsgText := Format( 'FTP檔案上傳檔案執行發生錯誤, 原因:%s', [E.Message] );
        Synchronize( Notify );
      end;
    end;
    if ( Self.Terminated ) then Break;       
    Sleep( 5000 );
    try
      if CheckCanExc then
      begin
        aFiles := 0;
        if DoFTPConnect then
        begin
          aFiles := FtpFileDetect;
        end;
        if ( aFiles > 0 ) then
        begin
          FNotify.MsgType := MB_ICONINFORMATION;
          FNotify.MsgText := Format( 'FTP上共有%d個回覆檔案, 準備下載。', [aFiles] );
          Synchronize( Notify );
          if DoFTPConnect then
          begin
            DoFTPDownlaod;
          end;
        end else
        begin
          FNotify.MsgType := MB_ICONINFORMATION;
          FNotify.MsgText := Format( '無回覆檔案可下載。', [aFiles] );
          Synchronize( Notify );
        end;
        FLastExc := Now;
      end;
      DoFTPDisconnect;  
    except
      on E: Exception do
      begin
        FNotify.MsgType := MB_ICONERROR;
        FNotify.MsgText := Format( 'FTP檔案下載回覆檔執行發生錯誤, 原因:%s', [E.Message] );
        Synchronize( Notify );
      end;
    end;
  end;
  FNotify.MsgType := MB_ICONINFORMATION;
  FNotify.MsgText := 'FTP檔案上/下傳執行緒執行結束。';
  Synchronize( Notify );
end;

{ ---------------------------------------------------------------------------- }

function TFtpThread.LocalFileDetect: Integer;
var
  aSearch: TSearchRec;
  aSearchMask, aPrefix: String;
begin
  Result := 0;
  aSearchMask := IncludeTrailingPathDelimiter( LocalUploadDir ) + '*.txt';
  if FindFirst( aSearchMask, faAnyFile, aSearch ) = 0 then
  begin
    try
      repeat
         if ( aSearch.Attr and faDirectory ) <> 0 then
         begin
           if ( aSearch.Name = '.' ) or ( aSearch.Name = '..' ) then Continue;
         end else
         if ( aSearch.Attr and faArchive ) <> 0 then
         begin
           aPrefix := UpperCase( Copy( aSearch.Name, 1, 3 ) );
           if ( aPrefix = 'CA-' ) or
              ( aPrefix = 'SA-' ) or
              ( aPrefix = 'RA-' ) then
             Inc( Result );
         end;
      until FindNext( aSearch ) <> 0;
    finally
      FindClose( aSearch );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TFtpThread.FtpFileDetect: Integer;
begin
  if not FTPControler.SwitchToFTPDownloadFolder then
  begin
    Result := -1;
  end else
  begin
    Result := FTPControler.DownloadFileCount;
  end;  
  if ( Result = -1 ) then
  begin
    FNotify.MsgType := MB_ICONERROR;
    FNotify.MsgText := Format( '偵測遠端FTP下載檔案數量發生錯誤, 原因:%s。',
      [FTPControler.ErrorText] );
    Synchronize( Notify );
    Sleep( 100 );  
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TFtpThread.Notify;
begin
  Main.AddMsg( FNotify );
end;

{ ---------------------------------------------------------------------------- }
procedure TFtpThread.DoFTPDownlaod;
begin
  FTPControler.DownloadFile;
end;

{ ---------------------------------------------------------------------------- }

procedure TFtpThread.OnDownloadFile(const aErrorText, aFileName: String);
begin
  if ( aErrorText <> EmptyStr ) then
  begin
    FNotify.MsgType := MB_ICONERROR;
    FNotify.MsgText := Format( '下載回覆檔%s發生錯誤, 原因:%s。',
      [aFileName, aErrorText] );
  end else
  begin
    FNotify.MsgType := MB_ICONINFORMATION;
    FNotify.MsgText := Format( '回覆檔%s下載完成。', [aFileName] );
  end;
  Synchronize( Notify );
end;

{ ---------------------------------------------------------------------------- }

function TFtpThread.CheckCanExc: Boolean;
begin
  Result := ( FLastExc <= 0 );
  if Result then Exit;   
  Result := ( MinutesBetween( Now, FLastExc ) >= 1 );      
end;

{ ---------------------------------------------------------------------------- }

end.
 