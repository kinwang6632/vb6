unit cbWatchThread;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, cbClass, cbAppClass,
  IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient, IdFTP, IdFTPList,
  IdFTPListParseUnix;

type
  TWatchThread = class(TMessageQueueThread)
  private
    { Private declarations }
    FMsgSubject: TMessageSubject;
    FLastExecute: TDateTime;
    FParam: TAppParam;
    FPeriod: TAppPeriod;
    FIdFtp: TIdFTP;
    procedure SetParam(const Value: TAppParam);
    procedure SetPeriod(const Value: TAppPeriod);
    function CheckPeriodTimeToReach: Boolean;
    function XmlFileDetected: Boolean;
    function FtpFileDetected: Boolean;
    function FtpFileCount(const aDir: String; const aType: TXmlDirectory; var aErrMsg: String): Integer;
    function DownloadFile(const aDir: String; const aType: TXmlDirectory; var aErrMsg: String): Integer;
    function DownloadAllFile: Boolean;
    function DoFtpConnect: Boolean;
    function DoFtpDisconnect: Boolean;
    procedure XmlFileMove;
  protected
    procedure Execute; override;
    procedure WndProc(var Msg: TMessage); override;
  public
    constructor Create;
    destructor Destroy; override;
    property MessageSubject: TMessageSubject read FMsgSubject;
    property Param: TAppParam read FParam write SetParam;
    property Perido: TAppPeriod read FPeriod write SetPeriod;
  end;

implementation

uses DateUtils;

{ ---------------------------------------------------------------------------- }

{ TWatchThread }

constructor TWatchThread.Create;
begin
  inherited Create( True );
  FMsgSubject := TMessageSubject.Create;
  FParam := TAppParam.Create;
  FPeriod := TAppPeriod.Create;
  FIdFtp := TIdFTP.Create( nil );
end;

{ ---------------------------------------------------------------------------- }

destructor TWatchThread.Destroy;
begin
  FParam.Free;
  FPeriod.Free;
  FMsgSubject.Free;
  FIdFtp.Free;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TWatchThread.SetParam(const Value: TAppParam);
begin
  FParam.Assign( Value );
  FIdFtp.Host := FParam.FtpServer;
  FIdFtp.Port := 21;
  FIdFtp.Username := FParam.FtpAccount;
  FIdFtp.Password := FParam.FtpPassword;
  FIdFtp.ReadTimeout := 10000;
  FIdFtp.Passive := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TWatchThread.SetPeriod(const Value: TAppPeriod);
begin
  FPeriod.Assign( Value );
end;

{ ---------------------------------------------------------------------------- }

procedure TWatchThread.Execute;
begin
  WaitForPlaySignal;
  FLastExecute := 0;
  while not Self.Terminated do
  begin
    Sleep( 100 );
    try
      if CheckPeriodTimeToReach then
      begin
        FLastExecute := Now;
        if DoFtpConnect then
        begin
          if FtpFileDetected then
          begin
            if Self.Terminated then Break;
            DownloadAllFile;
            if Self.Terminated then Break;
          end;
        end;
        DoFtpDisconnect;
        if XmlFileDetected then
        begin
          if Self.Terminated then Break;
          XmlFileMove;
        end;
      end;
      Sleep( 100 );
      if Self.Terminated then Break;
    except
      on E: Exception do
      begin
        FMsgSubject.Error( Format( '【檔案監控】執行緒發生錯誤, 原因:%s。',
          [E.Message] ) );
      end;
    end;
    Sleep( 300 );
  end;
  WaitForStopSignal;
end;

{ ---------------------------------------------------------------------------- }

procedure TWatchThread.WndProc(var Msg: TMessage);
begin
  inherited WndProc( Msg );
end;

{ ---------------------------------------------------------------------------- }

function TWatchThread.CheckPeriodTimeToReach: Boolean;
var
  aIndex: Integer;
  aHour: Word;
begin
  Result := False;
  if ( FPeriod.PeriodType = '1' ) then
  begin
    if ( FLastExecute = 0 ) then
    begin
      Result := True;
      Sleep( 1000 * 10 );
    end else
      Result := ( MinutesBetween( Now, FLastExecute ) >= FPeriod.PeriodMinute );
  end else
  begin
    FPeriod.CheckExpire;
    if ( DayOf( FLastExecute ) <> DayOf( Now ) ) then
      FPeriod.ResetExecute;
    aHour := HourOf( Now );
    for aIndex := 0 to FPeriod.ClockCount - 1 do
    begin
      if ( FPeriod.Clocks[aIndex].Expire ) then Continue;
      if ( not FPeriod.Clocks[aIndex].IsSet ) then Continue;
      if ( FPeriod.Clocks[aIndex].Execute ) then Continue;
      Result := ( aHour >= aIndex );
      if ( Result ) then
      begin
        FPeriod.Clocks[aIndex].Execute := True;
        Break;
      end;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TWatchThread.XmlFileDetected: Boolean;
var
  aMask:String;
  aSearch: TSearchRec;
  aFileCount: Integer;
  aIndex: TXmlDirectory;
begin
  Result := False;
  for aIndex := Low( TXmlDirectory ) to High( TXmlDirectory ) do
  begin
    if ( Self.Terminated ) then Break;
    case aIndex of
      xdsWrapper:
        aMask := IncludeTrailingPathDelimiter( FParam.WrapperSrc ) + 'Nagra_*.xml';
      xdsDex:
        aMask := IncludeTrailingPathDelimiter( FParam.DexSrc ) + 'CA*.xml';
      xdsAsRun:
        aMask := IncludeTrailingPathDelimiter( FParam.AsRunSrc ) + 'Export*.xml';
    else
      Continue;
    end;
    aFileCount := 0;
    if FindFirst( aMask, faAnyFile, aSearch ) = 0 then
    begin
      try
        repeat
          if ( aSearch.Attr and faDirectory ) <> 0 then
          begin
            if ( aSearch.Name = '.' ) or ( aSearch.Name = '..' ) then Continue;
          end else
          if ( aSearch.Attr and faArchive ) <> 0 then
          begin
            Inc( aFileCount );
          end;
        until FindNext( aSearch ) <> 0;
      finally
        FindClose( aSearch );
      end;
    end;
    if ( Self.Terminated ) then Break;
    Result := ( Result or ( aFileCount > 0 ) );
    if ( aFileCount > 0 ) then
    begin
      case aIndex of
        xdsWrapper:
          FMsgSubject.Normal( Format( '【Warpper】來源目錄: %s 下, 偵測到%d個檔案。',
            [FParam.WrapperSrc, aFileCount] ) );
        xdsDex:
          FMsgSubject.Normal( Format( '【Dex目錄】來源目錄: %s 下, 偵測到%d個檔案。',
            [FParam.DexSrc, aFileCount] ) );
        xdsAsRun:
          FMsgSubject.Normal( Format( '【AsRun】來源目錄: %s 下, 偵測到%d個檔案。',
            [FParam.AsRunSrc, aFileCount] ) );
      end;
      Sleep( 300 );
    end;
    if ( Self.Terminated ) then Break;  
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TWatchThread.XmlFileMove;
var
  aMask, aSource, aDest: String;
  aSearch: TSearchRec;
  aIndex: TXmlDirectory;
begin
  for aIndex := Low( TXmlDirectory ) to High( TXmlDirectory ) do
  begin
    if ( Self.Terminated ) then Break;
    case aIndex of
      xdsWrapper:
        aMask := IncludeTrailingPathDelimiter( FParam.WrapperSrc ) + 'Nagra_*.xml';
      xdsDex:
        aMask := IncludeTrailingPathDelimiter( FParam.DexSrc ) + 'CA*.xml';
      xdsAsRun:
        aMask := IncludeTrailingPathDelimiter( FParam.AsRunSrc ) + 'Export*.xml';
    else
      Continue;
    end;
    if FindFirst( aMask, faAnyFile, aSearch ) = 0 then
    begin
      try
        repeat
          if ( aSearch.Attr and faDirectory ) <> 0 then
          begin
            if ( aSearch.Name = '.' ) or ( aSearch.Name = '..' ) then Continue;
          end else
          if ( aSearch.Attr and faArchive ) <> 0 then
          begin
            aSource := EmptyStr;
            aDest := EmptyStr;
            case aIndex of
              xdsWrapper:
                begin
                  aSource := IncludeTrailingPathDelimiter(
                    FParam.WrapperSrc ) + aSearch.Name;
                  aDest := IncludeTrailingPathDelimiter(
                    FParam.WrapperDest ) + aSearch.Name;
                end;
              xdsDex:
                begin
                  aSource := IncludeTrailingPathDelimiter(
                    FParam.DexSrc ) + aSearch.Name;
                  aDest := IncludeTrailingPathDelimiter(
                    FParam.DexDest ) + aSearch.Name;
                end;
              xdsAsRun:
                begin
                  aSource := IncludeTrailingPathDelimiter(
                    FParam.AsRunSrc ) + aSearch.Name;
                  aDest := IncludeTrailingPathDelimiter(
                    FParam.AsRunDest ) + aSearch.Name;
                end;
            end;
            { 來源及目的的目錄皆有設定, 且不相同才搬移 }  
            if ( aSource <> EmptyStr ) and ( aDest <> EmptyStr ) and
               ( aSource <> aDest ) then
            begin
              MoveFileEx( PChar( aSource ), PChar( aDest ), MOVEFILE_REPLACE_EXISTING );
              FMsgSubject.Normal( Format( '【檔案監控】移動檔案: %s 至 %s。',
                [aSearch.Name, ExtractFilePath( aDest ) ] ) );
              Sleep( 100 );
            end;
          end;
        until FindNext( aSearch ) <> 0;
      finally
        FindClose( aSearch );
      end;
    end;
    if ( Self.Terminated ) then Break;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TWatchThread.DoFtpConnect: Boolean;
begin
  if ( FIdFtp.Connected ) then FIdFtp.Disconnect;
  try
    FIdFtp.Connect;
  except
    on E: Exception do
    begin
      FMsgSubject.Error( Format( '【檔案監控】連接FTP時發生錯誤, 訊息:%s。',
        [E.Message] ) );
    end;
  end;
  Result := FIdFtp.Connected;
end;

{ ---------------------------------------------------------------------------- }

function TWatchThread.DoFtpDisconnect: Boolean;
begin
  try
    if ( FIdFtp.Connected ) then FIdFtp.Disconnect;
  except
    { ... }
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TWatchThread.FtpFileDetected: Boolean;
var
  aIndex: TXmlDirectory;
  aCount, aSum: Integer;
  aDir, aErrMsg, aText: String;
begin
  aSum := 0;
  for aIndex := Low( TXmlDirectory ) to High( TXmlDirectory ) do
  begin
    aCount := 0;
    if ( Self.Terminated ) then Break;
    case aIndex of
      xdsWrapper:
        aDir := FParam.FtpWrapperSrc;
      xdsDex:
        aDir := FParam.FtpDexSrc;
      xdsAsRun:
        aDir := FParam.FtpAsRunSrc;
    else
      Continue;
    end;
    aErrMsg := EmptyStr;
    try
      aCount := FtpFileCount( aDir, aIndex, aErrMsg );
      aSum := ( aSum + aCount );
    except
      on E: Exception do aErrMsg := E.Message;
    end;
    if ( aErrMsg <> EmptyStr ) then
    begin
      case aIndex of
        xdsWrapper:
          aText := Format( '【檔案監控】【Warpper】FTP資料夾 <%s> 無法存取, 訊息:%s。',
            [aDir, aErrMsg ] );
        xdsDex:
          aText := Format( '【檔案監控】【Dex目錄】FTP資料夾 <%s> 無法存取, 訊息:%s。',
            [aDir, aErrMsg] );
        xdsAsRun:
          aText := Format( '【檔案監控】【AsRun】FTP資料夾 <%s> 無法存取, 訊息:%s。',
            [aDir, aErrMsg] );
      end;
      FMsgSubject.Warning( aText );
    end else
    begin
      case aIndex of
        xdsWrapper:
          aText := Format( '【檔案監控】【Warpper】FTP資料夾 <%s> 偵測到%d個檔案。',
            [aDir, aCount ] );
        xdsDex:
          aText := Format( '【檔案監控】【Dex目錄】FTP資料夾 <%s> 偵測到%d個檔案。',
            [aDir, aCount] );
        xdsAsRun:
          aText := Format( '【檔案監控】【AsRun】FTP資料夾 <%s> 偵測到%d個檔案。',
            [aDir, aCount] );
      end;
      FMsgSubject.Normal( aText );
    end;
  end;
  Result := ( aSum > 0 );
end;

{ ---------------------------------------------------------------------------- }

function TWatchThread.FtpFileCount(const aDir: String; const aType: TXmlDirectory;
  var aErrMsg: String): Integer;
var
  aIndex: Integer;
  aItem: TIdFTPListItem;
  aExt: String;
  aBool: Boolean;
begin
  Result := 0;
  aErrMsg := EmptyStr;
  try
    FIdFtp.ChangeDir( '/' );
    FIdFtp.ChangeDir( aDir );
  except
    on E: Exception do
    begin
      aErrMsg := E.Message;
    end;
  end;
  {}
  if ( aErrMsg <> EmptyStr ) then Exit;
  {}
  FIdFtp.List;
  {}
  for aIndex := 0 to FIdFtp.DirectoryListing.Count - 1 do
  begin
    aItem := FIdFtp.DirectoryListing[aIndex];
    if ( aItem.ItemType = ditFile ) then
    begin
      aExt := ExtractFileExt( aItem.FileName );
      aBool := False;
      case aType of
        xdsWrapper:
          begin
            aBool :=
              ( UpperCase( Copy( aItem.FileName, 1, 6 ) ) = 'NAGRA_' ) and
              ( UpperCase( aExt ) = '.XML' );
          end;
        xdsDex:
          begin
            aBool :=
              ( UpperCase( Copy( aItem.FileName, 1, 2 ) ) = 'CA' ) and
              ( UpperCase( aExt ) = '.XML' );
          end;
        xdsAsRun:
          begin
            aBool :=
              ( UpperCase( Copy( aItem.FileName, 1, 6 ) ) = 'EXPORT' ) and
              ( UpperCase( aExt ) = '.XML' );
          end;
      end;
      if ( aBool ) then Inc( Result );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TWatchThread.DownloadFile(const aDir: String; const aType: TXmlDirectory;
  var aErrMsg: String): Integer;
var
  aIndex: Integer;
  aItem: TIdFTPListItem;
  aExt, aDest, aText, aBackDir: String;
  aBool: Boolean;
begin
  Result := 0;
  aErrMsg := EmptyStr;
  try
    FIdFtp.ChangeDir( '/' );
    FIdFtp.ChangeDir( aDir );
  except
    on E: Exception do
    begin
      aErrMsg := E.Message;
    end;
  end;
  {}
  if ( aErrMsg <> EmptyStr ) then Exit;
  {}
  FIdFtp.List( nil );
  {}
  for aIndex := 0 to FIdFtp.DirectoryListing.Count - 1 do
  begin
    aItem := FIdFtp.DirectoryListing[aIndex];
    if ( aItem.ItemType = ditFile ) then
    begin
      aExt := ExtractFileExt( aItem.FileName );
      aBool := False;
      case aType of
        xdsWrapper:
          begin
            aBool :=
              ( UpperCase( Copy( aItem.FileName, 1, 6 ) ) = 'NAGRA_' ) and
              ( UpperCase( aExt ) = '.XML' );
            aDest := IncludeTrailingPathDelimiter( FParam.WrapperSrc ) + aItem.FileName;
          end;
        xdsDex:
          begin
            aBool :=
              ( UpperCase( Copy( aItem.FileName, 1, 2 ) ) = 'CA' ) and
              ( UpperCase( aExt ) = '.XML' );
            aDest := IncludeTrailingPathDelimiter( FParam.DexSrc ) + aItem.FileName;
          end;
        xdsAsRun:
          begin
            aBool :=
              ( UpperCase( Copy( aItem.FileName, 1, 6 ) ) = 'EXPORT' ) and
              ( UpperCase( aExt ) = '.XML' );
            aDest := IncludeTrailingPathDelimiter( FParam.AsRunSrc ) + aItem.FileName;
          end;
      end;
      {}
      if ( not aBool ) then Continue;
      case aType of
        xdsWrapper:
          begin
            aText := Format( '【檔案監控】【Warpper】下載檔案 <%s>。',
              [aItem.FileName] );
            aBackDir := FParam.FtpWrapperBackup;
          end;
        xdsDex:
          begin
            aText := Format( '【檔案監控】【Dex目錄】下載檔案 <%s>。',
              [aItem.FileName] );
            aBackDir := FParam.FtpDexBackup;
          end;
        xdsAsRun:
          begin
            aText := Format( '【檔案監控】【AsRun】下載檔案 <%s>。',
              [aItem.FileName] );
            aBackDir := FParam.FtpAsRunBackup;
          end;
      end;
      if not IsDelimiter( '/', aBackDir, Length( aBackDir ) ) then
        aBackDir := ( aBackDir + '/' );
      try
        FMsgSubject.Normal( aText );
        FIdFtp.Get( aItem.FileName, aDest, True );
        { 移到備份目錄 }
        FIdFtp.Rename( aItem.FileName, aBackDir + aItem.FileName );
        Inc( Result );
      except
        on E: Exception do aErrMsg := E.Message;
      end;
      { 該檔案失敗, 刪掉 }
      if ( aErrMsg <> EmptyStr ) and ( FileExists( aDest ) ) then
      begin
        DeleteFile( aDest );
        Break;
      end;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }


function TWatchThread.DownloadAllFile: Boolean;
var
  aIndex: TXmlDirectory;
  aCount, aSum: Integer;
  aDir, aErrMsg, aText: String;
begin
  aSum := 0;
  for aIndex := Low( TXmlDirectory ) to High( TXmlDirectory ) do
  begin
    aCount := 0;
    if ( Self.Terminated ) then Break;
    case aIndex of
      xdsWrapper:
        aDir := FParam.FtpWrapperSrc;
      xdsDex:
        aDir := FParam.FtpDexSrc;
      xdsAsRun:
        aDir := FParam.FtpAsRunSrc;
    else
      Continue;
    end;
    aErrMsg := EmptyStr;
    try
      aCount := DownloadFile( aDir, aIndex, aErrMsg );
      aSum := ( aSum + aCount );
    except
      on E: Exception do aErrMsg := E.Message;
    end;
    if ( aErrMsg <> EmptyStr ) then
    begin
      case aIndex of
        xdsWrapper:
          aText := Format( '【檔案監控】【Warpper】FTP資料夾 <%s> 無法存取, 訊息:%s。',
            [aDir, aErrMsg ] );
        xdsDex:
          aText := Format( '【檔案監控】【Dex】FTP資料夾 <%s> 無法存取, 訊息:%s。',
            [aDir, aErrMsg] );
        xdsAsRun:
          aText := Format( '【檔案監控】【AsRun】FTP資料夾 <%s> 無法存取, 訊息:%s。',
            [aDir, aErrMsg] );
      end;
      FMsgSubject.Warning( aText );
    end else
    begin
      case aIndex of
        xdsWrapper:
          aText := Format( '【檔案監控】【Warpper】FTP資料夾 <%s> 下載%d個檔案。',
            [aDir, aCount ] );
        xdsDex:
          aText := Format( '【檔案監控】【Dex】FTP資料夾 <%s> 下載%d個檔案。',
            [aDir, aCount] );
        xdsAsRun:
          aText := Format( '【檔案監控】【AsRun】FTP資料夾 <%s> 下載%d個檔案。',
            [aDir, aCount] );
      end;
      FMsgSubject.Normal( aText );
    end;
  end;
  Result := ( aSum > 0 );
end;

{ ---------------------------------------------------------------------------- }
end.
