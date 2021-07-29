unit cbWatchDogThread;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, cbClass,
  IdMessage, IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient,
  IdMessageClient, IdSMTP;

type
  TWatchDogThread = class(TThread)
  private
    { Private declarations }
    FNotify: TNotify;
    FEmailAccount: TEmailAccount;
    FEMailList: TStringList;
    FLastExc: TDateTime;
    FWatchFile: TWatchFile;
    FIdSMTP: TIdSMTP;
    procedure SetEMailList(const Value: TStringList);
    procedure Notify;
    function CheckCanExc: Boolean;
    function LocalFileDetect: Integer;
    procedure SendEMail;
    procedure SetEmailAccount(const Value: TEmailAccount);
  protected
    procedure Execute; override;
  public
    property EMailList: TStringList read FEMailList write SetEMailList;
    property EMailAccount: TEmailAccount read FEmailAccount write SetEmailAccount;
    constructor Create;
    destructor Destroy; override;
  end;

implementation

uses cbMain, cbDataControler, cbFTPControler, DateUtils;

{ Important: Methods and properties of objects in visual components can only be
  used in a method called using Synchronize, for example,

      Synchronize(UpdateCaption);

  and UpdateCaption could look like,

    procedure TWatchDogThread.UpdateCaption;
    begin
      Form1.Caption := 'Updated in a thread';
    end; }

{ TWatchDogThread }

{ ---------------------------------------------------------------------------- }

constructor TWatchDogThread.Create;
begin
  inherited Create( True );
  FEmailAccount.Account := EmptyStr;
  FEmailAccount.Password := EmptyStr;
  FEmailAccount.Server := EmptyStr;
  FEMailList := TStringList.Create;
  FWatchFile.FileName := EmptyStr;
  FWatchFile.LastSee := 0;
  FWatchFile.SendEmailCount := 0;
  FIdSMTP := TIdSMTP.Create( nil );
  FIdSMTP.AuthenticationType := atLogin;
end;

{ ---------------------------------------------------------------------------- }

destructor TWatchDogThread.Destroy;
begin
  FEMailList.Free;
  FIdSMTP.Free;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TWatchDogThread.SetEMailList(const Value: TStringList);
begin
  FEMailList.Clear;
  FEMailList.Assign( Value );
end;

{ ---------------------------------------------------------------------------- }

procedure TWatchDogThread.SetEmailAccount(const Value: TEmailAccount);
begin
  FEmailAccount.Account := Value.Account;
  FEmailAccount.Password := Value.Password;
  FEmailAccount.Server := Value.Server;
  FEmailAccount.From := Value.From;
end;

{ ---------------------------------------------------------------------------- }

procedure TWatchDogThread.Notify;
begin
  Main.AddMsg( FNotify );
end;

{ ---------------------------------------------------------------------------- }

procedure TWatchDogThread.Execute;
var
  aFileCount: Integer;
begin
  FNotify.MsgType := MB_ICONINFORMATION;
  FNotify.MsgText := '監控執行緒開始執行。';
  Synchronize( Notify );
  FLastExc := Now;
  while not Self.Terminated do
  begin
    if CheckCanExc then
    begin
      aFileCount := LocalFileDetect;
      FLastExc := Now;
    end;
    Sleep( 300 );
  end;
  FNotify.MsgType := MB_ICONINFORMATION;
  FNotify.MsgText := '監控執行緒執行結束。';
  Synchronize( Notify );
end;

{ ---------------------------------------------------------------------------- }

function TWatchDogThread.CheckCanExc: Boolean;
begin
  Result := ( FLastExc <= 0 );
  if Result then Exit;   
  Result := ( MinutesBetween( Now, FLastExc ) >= 3 );
end;

{ ---------------------------------------------------------------------------- }

function TWatchDogThread.LocalFileDetect: Integer;
var
  aSearch: TSearchRec;
  aFileFound: Boolean;
  aSearchMask, aPrefix: String;
  aHighWaterMark: Integer;
begin
  Result := 0;
  aFileFound := False;
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
           begin
             aFileFound := True;
             if FWatchFile.FileName = EmptyStr then
             begin
               FWatchFile.FileName := aSearch.Name;
               FWatchFile.LastSee := Now;
               FWatchFile.SendEmailCount := 0;
             end;
             if ( FWatchFile.FileName = aSearch.Name ) then
             begin
               aHighWaterMark := 10;
               if ( MinutesBetween( Now, FWatchFile.LastSee ) >= aHighWaterMark ) then
               begin
                 { 寄送?次 }
                 if ( FWatchFile.SendEmailCount <= 0 ) and
                    ( FEMailList.Count > 0 ) then
                 begin
                   FNotify.MsgType := MB_ICONWARNING;
                   FNotify.MsgText :=
                     Format( '監控執行緒偵測到交換檔案已超過%d分鐘未上傳, 開始發送通知。', [aHighWaterMark] );
                   Synchronize( Notify );
                   Sleep( 100 );
                   SendEMail;
                 end;
                 FWatchFile.SendEmailCount := ( FWatchFile.SendEmailCount + 1 );
               end;
             end;
             Inc( Result );
           end;
         end;
      until FindNext( aSearch ) <> 0;
    finally
      FindClose( aSearch );
    end;
  end;
  if not aFileFound then
  begin
    FWatchFile.FileName := EmptyStr;
    FWatchFile.LastSee := 0;
    FWatchFile.SendEmailCount := 0;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TWatchDogThread.SendEMail;
var
  aDatTime, aEmailList: String;
  aIdMessage: TIdMessage;
begin
  aDatTime := FormatDateTime( 'yyyy-mm-dd hh:nn:ss',  Now );
  aEmailList := StringReplace( FEMailList.Text, #13#10, ';', [rfReplaceAll] );
  FIdSMTP.Host := FEmailAccount.Server;
  FIdSMTP.Username := FEmailAccount.Account;
  FIdSMTP.Password := FEmailAccount.Password;
  aIdMessage := TIdMessage.Create( nil );
  try
    aIdMessage.From.Text := FEmailAccount.From;
    aIdMessage.Sender.Text := aIdMessage.From.Text;
    aIdMessage.Recipients.EMailAddresses := aEmailList;
    aIdMessage.Subject := Format( 'APOL FTP Agent 運作異常通知 (%s)', [aDatTime] );
    aIdMessage.Body.Text :=
     ' 此通知信由FTP Agent發送。'#13#10#13#10 +
     ' FTP檔案交換運作可能出現異常, 請盡速查核。';
    try
      FIdSMTP.Connect( 10000 );
      FIdSMTP.Send( aIdMessage );
      FNotify.MsgType := MB_ICONWARNING;
      FNotify.MsgText := '警告信函已發送。';
      Synchronize( Notify );
    except
      on E: Exception do
      begin
        FNotify.MsgType := MB_ICONERROR;
        FNotify.MsgText := Format( '發送警告信函有誤, 原因:%s。', [E.Message] );
        Synchronize( Notify );
      end;
    end;
    try
      if FIdSMTP.Connected then FIdSMTP.Disconnect;
    except
      { ... }
    end;
  finally
    aIdMessage.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
 