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
    FNotify.MsgText := Format( 'FTP Server: %s �s�����C', [FTPControler.IdFTP.Host] );
    Synchronize( Notify );
    Sleep( 300 );
    FTPControler.Active := True;
    if FTPControler.Active then
    begin
      FNotify.MsgType := MB_ICONINFORMATION;
      FNotify.MsgText := Format( 'FTP Server: %s ���\�n�J�C',[FTPControler.IdFTP.Host] );
      Synchronize( Notify );
      Sleep( 300 );
    end else
    begin
      FNotify.MsgType := MB_ICONERROR;
      FNotify.MsgText := Format( 'FTP Server: %s �s������, ��]:%s�C',
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
    FNotify.MsgText := Format( '�P FTP Server: %s ���_�s�u�C', [FTPControler.IdFTP.Host] );
    Synchronize( Notify );
    Sleep( 300 );
    FTPControler.Active := False;
    FNotify.MsgType := MB_ICONINFORMATION;
    FNotify.MsgText := Format( '�P FTP Server: %s �s���w�����C', [FTPControler.IdFTP.Host] );
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
           FNotify.MsgText := Format( '�ɮ�:%s�W�Ǥ��C', [aSearch.Name] );
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
             FNotify.MsgText := Format( '�ɮ�:%s�w�W�ǧ����C', [aSearch.Name] );
           end else
           begin
             FNotify.MsgType := MB_ICONERROR;
             FNotify.MsgText := Format( '�ɮ�:%s�W�ǥ���, ��]:%s�C',
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
  FNotify.MsgText := 'FTP�ɮפW/�U�ǰ�����}�l����C';
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
        FNotify.MsgText := Format( '�W�Ǹ�Ƨ��@��%d�ɮ�, �ǳƤW�ǡC', [aFiles] );
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
        FNotify.MsgText := Format( 'FTP�ɮפW���ɮװ���o�Ϳ��~, ��]:%s', [E.Message] );
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
          FNotify.MsgText := Format( 'FTP�W�@��%d�Ӧ^���ɮ�, �ǳƤU���C', [aFiles] );
          Synchronize( Notify );
          if DoFTPConnect then
          begin
            DoFTPDownlaod;
          end;
        end else
        begin
          FNotify.MsgType := MB_ICONINFORMATION;
          FNotify.MsgText := Format( '�L�^���ɮץi�U���C', [aFiles] );
          Synchronize( Notify );
        end;
        FLastExc := Now;
      end;
      DoFTPDisconnect;  
    except
      on E: Exception do
      begin
        FNotify.MsgType := MB_ICONERROR;
        FNotify.MsgText := Format( 'FTP�ɮפU���^���ɰ���o�Ϳ��~, ��]:%s', [E.Message] );
        Synchronize( Notify );
      end;
    end;
  end;
  FNotify.MsgType := MB_ICONINFORMATION;
  FNotify.MsgText := 'FTP�ɮפW/�U�ǰ�������浲���C';
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
    FNotify.MsgText := Format( '��������FTP�U���ɮ׼ƶq�o�Ϳ��~, ��]:%s�C',
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
    FNotify.MsgText := Format( '�U���^����%s�o�Ϳ��~, ��]:%s�C',
      [aFileName, aErrorText] );
  end else
  begin
    FNotify.MsgType := MB_ICONINFORMATION;
    FNotify.MsgText := Format( '�^����%s�U�������C', [aFileName] );
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
 