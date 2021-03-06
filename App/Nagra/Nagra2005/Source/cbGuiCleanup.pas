unit cbGuiCleanup;

interface

uses
  SysUtils, Classes, Windows, Messages, Variants,
  cbClass, cxTL, cxListView;

type

  { 清除主畫面顯示的各種指令及訊息 }

  TGuiCleanup = class(TSMSCommandThread)
  private
    { Private declarations }
    FLogList: TStringList;
    FLastCheckControlCmd: Cardinal;
    FLastCheckFeedbackCmd: Cardinal;
    FLastCheckConsoleCmd: Cardinal;
    FLastCheckMessage: Cardinal;
    FTotalCount: Integer;
    FCleanupCount: Integer;
    FCleanupTreeList: TcxTreeList;
    FCleanupListView: TcxListView;
    function CheckCanCleanupControlCmd: Boolean;
    function CheckCanCleanupFeedbackCmd: Boolean;
    function CheckCanCleanupConsoleCmd: Boolean;
    function CheckCanCleanupMessage: Boolean;
    procedure CleanupControlCmd;
    procedure CleanupFeedbackCmd;
    procedure CleanupConsoleCmd;
    procedure CleanupMessage;
    procedure CleanupCmd;
    procedure CleanupMsg;
  protected
    procedure Execute; override;
  public
    constructor Create;
    destructor Destroy; override;
  end;


  procedure SaveMessageToLog(aLog: TStringList; const aDirPrefix: String);
  

implementation

uses cbMain, cbResStr, cbUtilis;

{ Important: Methods and properties of objects in visual components can only be
  used in a method called using Synchronize, for example,

      Synchronize(UpdateCaption);

  and UpdateCaption could look like,

    procedure TGuiCleanup.UpdateCaption;
    begin
      Form1.Caption := 'Updated in a thread';
    end; }

{ TGuiCleanup }

constructor TGuiCleanup.Create;
begin
  inherited Create;
  FLogList := TStringList.Create;
  FLastCheckControlCmd := 0;
  FLastCheckFeedbackCmd := 0;
  FLastCheckConsoleCmd := 0;
  FLastCheckMessage := 0;
  FTotalCount := 0;
  FCleanupCount := 0;
  FCleanupTreeList := nil;
  FCleanupListView := nil;
end;

{ ---------------------------------------------------------------------------- }

destructor TGuiCleanup.Destroy;
begin
  FLogList.Free;
  inherited Destroy;
end;

{ ---------------------------------------------------------------------------- }

procedure TGuiCleanup.Execute;
begin
  { 等待 Main Thread 打出執行的訊號 }
  WaitForPlaySignal;
  while not Stop do
  begin
    MessageSubject.RunState := rsRunning;
    if Stop then Break;
    Sleep( 200 );
    if CheckCanCleanupControlCmd then CleanupControlCmd;
    if Stop then Break;
    Sleep( 200 );
    if CheckCanCleanupFeedbackCmd then CleanupFeedbackCmd;
    if Stop then Break;
    Sleep( 200 );
    if CheckCanCleanupConsoleCmd then CleanupConsoleCmd;
    if Stop then Break;
    Sleep( 200 );
    if CheckCanCleanupMessage then CleanupMessage;
  end;
  Sleep( GetWaitWhileFrquence );
  MessageSubject.RunState := rsStop;
  WaitForTerminalSignal;
end;

{ ---------------------------------------------------------------------------- }


function TGuiCleanup.CheckCanCleanupControlCmd: Boolean;
var
  aRetry: Cardinal;
begin
  { 傳送指令的 TreeList 每 120 秒檢查一次 }
  if FLastCheckControlCmd <= 0 then
    FLastCheckControlCmd := GetTickCount;
  aRetry := GetBetweenFrquence( FLastCheckControlCmd );
  if aRetry <= 0 then aRetry := 120 * 1000;
  Result := aRetry >= ( 120 * 1000 );
end;

{ ---------------------------------------------------------------------------- }

function TGuiCleanup.CheckCanCleanupFeedbackCmd: Boolean;
var
  aRetry: Cardinal;
begin
  { 回傳機制指令每 360 秒檢查一次 }
  if FLastCheckFeedbackCmd <= 0 then
    FLastCheckFeedbackCmd := GetTickCount;
  aRetry := GetBetweenFrquence( FLastCheckFeedbackCmd );
  if aRetry <= 0 then aRetry := 360 * 1000;
  Result := aRetry >= ( 360 * 1000 );
end;

{ ---------------------------------------------------------------------------- }

function TGuiCleanup.CheckCanCleanupConsoleCmd: Boolean;
var
  aRetry: Cardinal;
begin
  { 命令模式顯示的指令每 180 秒檢查一次 }
  if FLastCheckConsoleCmd <= 0 then
    FLastCheckConsoleCmd := GetTickCount;
  aRetry := GetBetweenFrquence( FLastCheckConsoleCmd );
  if aRetry <= 0 then aRetry := 180 * 1000;
  Result := aRetry >= ( 180 * 1000 );
end;

{ ---------------------------------------------------------------------------- }

function TGuiCleanup.CheckCanCleanupMessage: Boolean;
var
  aRetry: Cardinal;
begin
  { 訊息每 360 秒檢查一次 }
  if FLastCheckMessage <= 0 then
    FLastCheckMessage := GetTickCount;
  aRetry := GetBetweenFrquence( FLastCheckMessage );
  if aRetry <= 0 then aRetry := 360 * 1000;
  Result := aRetry >= ( 360 * 1000 );
end;

{ ---------------------------------------------------------------------------- }

procedure TGuiCleanup.CleanupControlCmd;
begin
  FCleanupTreeList := fmMain.ControlSendTree;
  FTotalCount := FCleanupTreeList.Count;
  if FTotalCount > 1000 then
  begin
    { 清掉 1/5 }
    FCleanupCount := ( FTotalCount div 5 ) + 1;
    while FCleanupCount > 0 do
    begin
      Synchronize( CleanupCmd );
      if Stop then Break;
      Sleep( 200 );
      if Stop then Break;
    end;
  end;  
  FLastCheckControlCmd := GetTickCount;
end;

{ ---------------------------------------------------------------------------- }

procedure TGuiCleanup.CleanupFeedbackCmd;
begin
  FCleanupTreeList := fmMain.FeedbackTree;
  FTotalCount := FCleanupTreeList.Count;
  if FTotalCount > 1000 then
  begin
    { 清掉 1/5 }
    FCleanupCount := ( FTotalCount div 5 ) + 1;
    while FCleanupCount > 0 do
    begin
      Synchronize( CleanupCmd );
      if Stop then Break;
      Sleep( 200 );
      if Stop then Break;
    end;
  end;  
  FLastCheckFeedbackCmd := GetTickCount;
end;

{ ---------------------------------------------------------------------------- }

procedure TGuiCleanup.CleanupConsoleCmd;
begin
  FCleanupTreeList := fmMain.ConsoleTree;
  FTotalCount := FCleanupTreeList.Count;
  if FTotalCount > 1000 then
  begin
    { 清掉 1/5 }
    FCleanupCount := ( FTotalCount div 5 ) + 1;
    while FCleanupCount > 0 do
    begin
      Synchronize( CleanupCmd );
      if Stop then Break;
      Sleep( 200 );
      if Stop then Break;
    end;
  end;  
  FLastCheckConsoleCmd := GetTickCount;
end;

{ ---------------------------------------------------------------------------- }

procedure TGuiCleanup.CleanupMessage;
var
  aIndex: Integer;
begin
  for aIndex := 1 to 4 do
  begin
    case aIndex of
      1: FCleanupListView := fmMain.StartupMsgList;
      2: FCleanupListView := fmMain.SoDatabaseMsgList;
      3: FCleanupListView := fmMain.ControlMsgList;
      4: FCleanupListView := fmMain.FeedbackMsgList;
    else
      FCleanupListView := nil;
    end;
    if Stop then Break;
    if not Assigned( FCleanupListView ) then Continue;
    FTotalCount := FCleanupListView.Items.Count;
    if FTotalCount > 500 then
    begin
      { 清掉 1/2 }
      FCleanupCount := ( FTotalCount div 2 ) + 1;
      while FCleanupCount > 0 do
      begin
        Synchronize( CleanupMsg );
        if Stop then Break;
        Sleep( 100 );
        if Stop then Break;
      end;
      case aIndex of
        1: SaveMessageToLog( FLogList, 'Log-Startup' );
        2: SaveMessageToLog( FLogList, 'Log-Database' );
        3: SaveMessageToLog( FLogList, 'Log-Control' );
        4: SaveMessageToLog( FLogList, 'Log-Feedback' );
      end;
      if Stop then Break;
    end;
  end;  
  FLastCheckMessage := GetTickCount;
end;

{ ---------------------------------------------------------------------------- }

procedure TGuiCleanup.CleanupCmd;
var
  aCount: Integer;
begin
  { 一次刪除 10 筆 }
  aCount := 10;
  if aCount > FCleanupCount then aCount := FCleanupCount;
  FCleanupTreeList.BeginUpdate;
  try
   while aCount > 0 do
   begin
     if Assigned( FCleanupTreeList.TopNode ) then
     begin
       if FCleanupTreeList.TopNode.HasChildren then
       begin
         Dec( FCleanupCount, FCleanupTreeList.TopNode.Count );
         Dec( aCount, FCleanupTreeList.TopNode.Count );
       end else
       begin
         Dec( FCleanupCount );
         Dec( aCount );
       end;
       FCleanupTreeList.TopNode.Delete;
     end else
     begin
       Dec( aCount );
     end;  
   end;
  finally
    FCleanupTreeList.EndUpdate;
  end; 
end;

{ ---------------------------------------------------------------------------- }

procedure TGuiCleanup.CleanupMsg;
var
  aCount: Integer;
begin
  { 一次清 30 筆訊息 }
  aCount := 30;
  if aCount >= FCleanupCount then aCount := FCleanupCount;
  FCleanupListView.Items.BeginUpdate;
  try
    while aCount > 0 do
    begin
      FLogList.Add( FCleanupListView.Items[0].Caption );
      FCleanupListView.Items[0].Delete;
      Dec( FCleanupCount );
      Dec( aCount );
    end;
  finally
    FCleanupListView.Items.EndUpdate;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure SaveMessageToLog(aLog: TStringList; const aDirPrefix: String);
var
  aFileWnd: Integer;
  aDir, aDateTime, aFileName: String;
  aStream: TMemoryStream;
begin
  aDir := IncludeTrailingPathDelimiter( ExtractFilePath( ParamStr( 0 ) ) ) + aDirPrefix;
  if not DirectoryExists( aDir ) then
    ForceDirectories( aDir );
  aDateTime := Format( '%s.log' , [FormatDateTime( 'yyyy-mm-dd', Now )] );
  aFileName := IncludeTrailingPathDelimiter( aDir ) + aDateTime;
  if not FileExists( aFileName ) then
  begin
    aFileWnd := FileCreate( aFileName );
    FileClose( aFileWnd );
  end;
  aFileWnd := FileOpen( aFileName, fmOpenReadWrite );
  try
    aStream := TMemoryStream.Create;
    try
      aLog.SaveToStream( aStream );
      aLog.Clear;
      aStream.Position := 0;
      FileSeek( aFileWnd, 0, 2 );
      FileWrite( aFileWnd, aStream.Memory^ , aStream.Size );
      aStream.Clear;
    finally
      aStream.Free;
    end;  
  finally
    FileClose( aFileWnd );
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
