unit cbFTPControler;

interface

uses
  Windows, SysUtils, Classes, IdIntercept, IdLogBase, IdLogFile, IdBaseComponent,
  IdComponent, IdTCPConnection, IdTCPClient, IdFTP, IdFTPList, IdThreadMgr,
  IdThreadMgrDefault, IdThreadComponent, cbClass;

type
  TFTPControler = class(TDataModule)
    IdFTP: TIdFTP;
    procedure DataModuleCreate(Sender: TObject);
  private
    { Private declarations }
    FErrorText: String;
    FDownloadStatus: TDownloadStatus;
    function GetActive: Boolean;
    procedure SetActive(const Value: Boolean);
  public
    { Public declarations }
    property Active: Boolean read GetActive write SetActive;
    property ErrorText: String read FErrorText;
    function UploadFile(aFileName: String): Boolean;
    function DownloadFile: Boolean;
    function SwitchToFTPTransFolder: Boolean;
    function SwitchToFTPUploadFolder: Boolean;
    function SwitchToFTPDownloadFolder: Boolean;
    function DownloadFileCount: Integer;
    property OnDownloadStatus: TDownloadStatus read FDownloadStatus write FDownloadStatus;
  end;



var FTPControler: TFTPControler;
  
implementation

uses cbMain, cbDataControler;


{$R *.dfm}

{ TFTPControler }

{ ---------------------------------------------------------------------------- }

procedure TFTPControler.DataModuleCreate(Sender: TObject);
begin
  FErrorText := EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

function TFTPControler.GetActive: Boolean;
begin
  Result := IdFTP.Connected;
end;

{ ---------------------------------------------------------------------------- }

procedure TFTPControler.SetActive(const Value: Boolean);
begin
  FErrorText := EmptyStr;
  if Value then
  begin
    if not IdFTP.Connected then
    begin
      try
        IdFTP.Connect( True, 60000 );
      except
        on E: Exception do FErrorText := E.Message;
      end;
    end;
  end else
  begin
    if IdFTP.Connected then IdFTP.Disconnect;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TFTPControler.SwitchToFTPTransFolder: Boolean;
begin
  FErrorText := EmptyStr;
  Result := False;
  try
    IdFTP.ChangeDirUp;
    IdFTP.ChangeDir( FtpEmcTransDir );
    Result := True;
  except
    on E: Exception do FErrorText := E.Message;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TFTPControler.SwitchToFTPUploadFolder: Boolean;
begin
  FErrorText := EmptyStr;
  Result := False;
  try
    IdFTP.ChangeDirUp;
    IdFTP.ChangeDir(  FtpEmcUploadDir );
    Result := True;
  except
    on E: Exception do FErrorText := E.Message;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TFTPControler.SwitchToFTPDownloadFolder: Boolean;
begin
  FErrorText := EmptyStr;
  Result := False;
  try
    IdFTP.ChangeDirUp;
    IdFTP.ChangeDir( FtpApolUploadDir );
    Result := True;
  except
    on E: Exception do
    begin
      FErrorText := E.Message;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TFTPControler.UploadFile(aFileName: String): Boolean;
var
  aSourceFileName, aDestFileName: String;
begin
  FErrorText := EmptyStr;
  Result := False;
  try
    aSourceFileName := Format( '/%s/%s' , [FtpEmcTransDir, ExtractFileName( aFileName )] );
    aDestFileName := Format( '/%s/%s' , [FtpEmcUploadDir, ExtractFileName( aFileName )] );
    try
      IdFTP.Delete( aFileName );
    except
      {...}
    end;
    IdFTP.Put( aFileName, ExtractFileName( aFileName ) );
    IdFTP.Rename( aSourceFileName, aDestFileName );
    Result := True;
  except
    on E: Exception do FErrorText := E.Message;
  end;
end;

{ ---------------------------------------------------------------------------- }


function TFTPControler.DownloadFile: Boolean;
var
  aIndex: Integer;
  aSource, aDest, aPrefix, aExten: String;
begin
  FErrorText := EmptyStr;
  IdFTP.List( nil );
  for aIndex := 0 to IdFTP.DirectoryListing.Count - 1 do
  begin
    try
      if IdFTP.DirectoryListing.Items[aIndex].ItemType = ditFile then
      begin
        aPrefix := Copy( IdFTP.DirectoryListing.Items[aIndex].FileName, 1, 3 );
        aExten := ExtractFileExt( IdFTP.DirectoryListing.Items[aIndex].FileName );
        if ( aPrefix = 'RP-' ) and ( UpperCase( aExten ) = '.TXT' ) then
        begin
          aSource := IncludeTrailingPathDelimiter( LocalTempDir ) +
            IdFTP.DirectoryListing.Items[aIndex].FileName;
          aDest := IncludeTrailingPathDelimiter( LocalDownloadDir ) +
            IdFTP.DirectoryListing.Items[aIndex].FileName;
          IdFTP.Get( IdFTP.DirectoryListing.Items[aIndex].FileName, aSource );
          MoveFile( PChar( aSource ), PChar( aDest ) );
          IdFTP.Delete( IdFTP.DirectoryListing.Items[aIndex].FileName );
          if Assigned( FDownloadStatus ) then FDownloadStatus( FErrorText,
            IdFTP.DirectoryListing.Items[aIndex].FileName );            
        end;
      end;
    except
      on E: Exception do
      begin
        FErrorText := E.Message;
        if Assigned( FDownloadStatus ) then FDownloadStatus( FErrorText,
          IdFTP.DirectoryListing.Items[aIndex].FileName );
      end;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TFTPControler.DownloadFileCount: Integer;
var
  aIndex: Integer;
  aPrefix: String;
begin
  FErrorText := EmptyStr;
  Result := 0;
  try
    IdFTP.List( nil );
    for aIndex := 0 to IdFTP.DirectoryListing.Count - 1 do
    begin
      if IdFTP.DirectoryListing.Items[aIndex].ItemType = ditFile then
      begin
        aPrefix := UpperCase( Copy( IdFTP.DirectoryListing.Items[aIndex].FileName, 1, 3 ) );
        if ( aPrefix = 'RP-' ) then Inc( Result );
      end;
    end;
  except
    on E: Exception do
    begin
      Result := -1;
      FErrorText := E.Message;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
