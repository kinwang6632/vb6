unit cbSendThread;

interface

uses
  SysUtils, Classes, Windows, Messages, Variants, DB, DBClient,
  {$IFDEF APPDEBUG} CodeSiteLogging, {$ENDIF} cbAppClass,
  IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient;

type
  TSendThread = class(TAppThread)
  private
    { Private declarations }
    FDvnCmdBuilder: TDvnCommandBuilder;
    FTextSubj: TTextSubject;
    FCmdSubj: TSendDvnSubject;
    FWMSocket: TWMSocket;
    FSocketErrCount: Integer;
    FLastSocketErr: TDateTime;
    FLastSendCmdTime: TDateTime;
    function OpenConnection: Boolean;
    procedure CloseConnection;
    procedure SetWMSocket(aReason: TSocketMsgReason; aErrText: String = '');
    procedure ThreadNotify(aHandle: THandle);
    function CheckSocketReConnect: Boolean;
    function CheckIdleTime: Boolean;
    procedure SendCmd;
  protected
    procedure Execute; override;
  public
    constructor Create(CreateSuspended: Boolean); override;
    destructor Destroy; override;
    property TextSubj: TTextSubject read FTextSubj;
    property CmdSubj: TSendDvnSubject read FCmdSubj;
  end;

  TRecvThread = class(TAppThread)
  private
    FRecv: Boolean;
    FCmdSubj: TSendDvnSubject;
    FTextSubj: TTextSubject;
    FWMSocket: TWMSocket;
    FAckParser: TDvnAckParser;
    procedure SetWMSocket(aReason: TSocketMsgReason; aErrText: String = '');
    procedure RecvAck;
    procedure BufferToList(aRecvText: String);
    procedure ProcAck;
  protected
    procedure WndProc(var Msg: TMessage); override;
    procedure Execute; override;
  public
    constructor Create(CreateSuspended: Boolean); override;
    destructor Destroy; override;
    property TextSubj: TTextSubject read FTextSubj;
    property CmdSubj: TSendDvnSubject read FCmdSubj;
  end;


implementation

uses cbMain, cbUtilis, DateUtils;

{ Important: Methods and properties of objects in visual components can only be
  used in a method called using Synchronize, for example,

      Synchronize(UpdateCaption);

  and UpdateCaption could look like,

    procedure TSendThread.UpdateCaption;
    begin
      Form1.Caption := 'Updated in a thread';
    end; }

var
  SendThread, RecvThread: TMessageThread;
  FSocket: TIdTCPClient;    

{ TSendThread }

constructor TSendThread.Create(CreateSuspended: Boolean);
begin
  inherited Create( CreateSuspended );
  SendThread := Self;
  FDvnCmdBuilder := TDvnCommandBuilder.Create;
  FSocket := TIdTCPClient.Create( nil );
  FSocket.ReadTimeout := 3000;
  FTextSubj := TTextSubject.Create;
  FCmdSubj := TSendDvnSubject.Create;
  FSocketErrCount := 0;
  FLastSocketErr := 0;
  FLastSendCmdTime := 0;
  SetWMSocket( smrStop );
end;

{ ---------------------------------------------------------------------------- }

destructor TSendThread.Destroy;
begin
  StrDispose( FWMSocket.ErrorDescription );
  FSocket.Free;
  FTextSubj.Free;
  FCmdSubj.Free;
  FDvnCmdBuilder.Free;
  inherited;
  SendThread := nil;
  FSocket := nil;
end;

{ ---------------------------------------------------------------------------- }

function TSendThread.OpenConnection: Boolean;
var
  aErrText: String;
begin
  FTextSubj.InfoFmt( LanguageManager.Get( 'SSocketOpenConnection' ),
    [CasEnv.Ip, CasEnv.SendPort] );
  Sleep( 500 );
  try
    CloseConnection;
    FSocket.Host := CasEnv.Ip;
    FSocket.Port := StrToIntDef( CasEnv.SendPort, 2100 );
    { Connect TimeOut 設成 10 秒 }
    FSocket.Connect( 10000 );
    FSocket.ReadTimeout := 3000;
  except
    on E: Exception do
    begin
      aErrText := E.Message;
      FTextSubj.ErrorFmt( LanguageManager.Get( 'SSocketOpenConnectionError' ),
        [aErrText] );
    end;
  end;
  Result := FSocket.Connected;
  if ( Result ) then
  begin
    FSocketErrCount := 0;
    FLastSocketErr := 0;
    FLastSendCmdTime := 0;
    SetWMSocket( smrAck );
  end else
  begin
    Inc( FSocketErrCount );
    FLastSocketErr := Now;
    SetWMSocket( smrSocketError, aErrText );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSendThread.CloseConnection;
begin
  if FSocket.Connected then FSocket.Disconnect;
end;

{ ---------------------------------------------------------------------------- }

procedure TSendThread.SetWMSocket(aReason: TSocketMsgReason;
  aErrText: String);
begin
  FWMSocket.Msg := WM_SOCKET;
  FWMSocket.ThreadType := Ord( tdSend );
  FWMSocket.SocketMsgReason := Ord( aReason );
  FWMSocket.ErrorCode := 0;
  StrDispose( FWMSocket.ErrorDescription );
  FWMSocket.ErrorDescription := nil;
  if ( aReason = smrSocketError ) then
  begin
    FWMSocket.ErrorCode := -1000;
    if ( aErrText <> EmptyStr ) then
      FWMSocket.ErrorDescription := StrNew( PChar( aErrText ) );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSendThread.ThreadNotify(aHandle: THandle);
begin
  PostMessage( aHandle, FWMSocket.Msg, TMessage( FWMSocket ).WParam,
    TMessage( FWMSocket ).LParam );
end;

{ ---------------------------------------------------------------------------- }

function TSendThread.CheckSocketReConnect: Boolean;
begin
  Result := ( FLastSocketErr = 0 );
  if ( not Result ) then
  begin
    Result := ( not Terminated ) and
      ( SecondsBetween( Now, FLastSocketErr ) > CommEnv.CARetryFreq );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TSendThread.CheckIdleTime: Boolean;
begin
  if ( FLastSendCmdTime = 0 ) then FLastSendCmdTime := Now;
  Result := ( not Terminated ) and
    ( MinutesBetween( Now, FLastSendCmdTime ) >= CommEnv.CAIdleTime );
end;

{ ---------------------------------------------------------------------------- }

procedure TSendThread.Execute;
begin
  FDvnCmdBuilder.CasEnv.Assign( CasEnv );
  FDvnCmdBuilder.CommEnv.Assign( CommEnv );
  FDvnCmdBuilder.HighCmdEnv.Data := HighCmdEnv.Data;
  FDvnCmdBuilder.LowCmdEnv.Data := LowCmdEnv.Data;
  Sleep( 500 ); 
  while ( not Self.Terminated ) do
  begin
    if ( not FSocket.Connected ) and ( CheckSocketReConnect ) then
    begin
      { 通知接收指令回應執行緒, 先停止接收資料, 不打出此信號,
        一旦 Socket Open 就會跟 RecvThread 搶接收資料, 導致 Dead Lock }
      SetWMSocket( smrStop );
      ThreadNotify( RecvThread.MsgHandle );
      {}
      Sleep( 300 );
      if OpenConnection then
      begin
        { 連上了, 通知接收指令回應執行緒接收資料 }
        SetWMSocket( smrStart );
        ThreadNotify( RecvThread.MsgHandle );
        FTextSubj.OK( LanguageManager.Get( 'SSocketOpenConnectionSuccess' ) );
        Sleep( 300 );
      end else
      begin
        { 仍然連不上, 顯示重試訊息 }
        FTextSubj.WarningFmt( LanguageManager.Get( 'SSocketWaitRetry' ),
          [CommEnv.CARetryFreq] );
        Sleep( 300 );
      end;
      { 不管有沒有連上, 通知 MainForm, MainForm 會轉通知 SoDbThread }
      ThreadNotify( MainFormHandle );
    end;
    if Self.Terminated then Break;
    if ( FSocket.Connected ) then
    begin
      SendCmd;
      if ( FSocketErrCount >= CommEnv.CAMaxError ) or
         ( not FSocket.Connected ) then
      begin
        FTextSubj.WarningFmt( LanguageManager.Get(  'SSocketHasProblem' ),
          [FSocketErrCount] );
        FTextSubj.WarningFmt( LanguageManager.Get( 'SSocketWaitRetry' ),
          [CommEnv.CARetryFreq] );
        Sleep( 100 );
        FTextSubj.Warning( LanguageManager.Get( 'SSocketCloseConnection' ) );
        { 通知接收指令回應執行緒, 停止接收資料 }
        ThreadNotify( RecvThread.MsgHandle );
        { 通知主執行緒 SOCKET 連線有誤 }
        ThreadNotify( MainFormHandle );
        { 停止連線 }
        CloseConnection;
        { 最後傳送指令時間歸零 }
        FLastSendCmdTime := 0;
        { 記錄 Socket Error Time }
        FLastSocketErr := Now;
      end;
      {}
    end;
    { 假如已到 IdleTime 最大時間 ( 沒有傳送指令 ) }
    if ( CheckIdleTime ) then
    begin
      FTextSubj.WarningFmt( LanguageManager.Get( 'SSocketIdleTimeReach' ),
        [CommEnv.CAIdleTime] );
      Sleep( 1000 );
      FTextSubj.Warning( LanguageManager.Get( 'SSocketCloseConnection' ) );
      SetWMSocket( smrIdleTimeReach );
      { 通知接收指令回應執行緒, 停止接收資料 }
      ThreadNotify( RecvThread.MsgHandle );
      { 通知主執行緒準備重新連線 }
      ThreadNotify( MainFormHandle );
      { 停止連線 }
      CloseConnection;
      Sleep( 5000 );
      FLastSocketErr := 0;
    end;
    {}
    if Self.Terminated then Break;
    {}
    Sleep( 100 );
  end;
  FTextSubj.Info( LanguageManager.Get( 'SSocketCloseConnection' ) );
  CloseConnection;
  { 通知接收執行緒 --> 停止接收資料 }
  SetWMSocket( smrStop );
  ThreadNotify( RecvThread.MsgHandle );
  Sleep( 500 );
end;

{ ---------------------------------------------------------------------------- }

procedure TSendThread.SendCmd;
var
  aHighProcCount, aLowProcCount, aIndex: Integer;
  aHighCmd, aLowCmd: TSendDvn;
  aSocketErr: Boolean;
begin
  aHighProcCount := 0;
  aLowProcCount := 0;
  while ( HCmdBeginList.Count > 0 ) and
        ( aHighProcCount <= CommEnv.DbProcRecords ) do
  begin
    if ( Self.Terminated ) then Break;
    aHighCmd := HCmdBeginList.Objects[0];
    FDvnCmdBuilder.BuildLowCommand( aHighCmd );
    { 記錄高階指令所產生出來的低階指令數, Ack 回來時好判斷該高階指令是否已經完成 }
    aHighCmd.LowCmdCount := FDvnCmdBuilder.LowCmdCount;
    { 1.將本筆高階指令移到 待傳送 佇列 }
    HCmdWorkingList.AddObject( aHighCmd.HighSeqNo, aHighCmd );
    { 2.然後移到 已傳送 佇列 }
    HCmdBeginList.Delete( 0 );
    { TCP/IP 傳送錯誤 Flag }
    aSocketErr := False;
    for aIndex := 0 to FDvnCmdBuilder.LowCmdCount - 1 do
    begin
      aLowCmd := FDvnCmdBuilder.LowCmd[aIndex];
      aLowCmd.SendTime := Now;
      { 當為該高階指令的第一筆低階指令時,
        將該筆高階指令的 SendTime 定為該低階的 send time }
      if ( aIndex = 0 ) then aHighCmd.SendTime := aLowCmd.SendTime;
      try
        if ( not aSocketErr ) then
        begin
          FSocket.Write( aLowCmd.SendCmdText );
{$IFDEF APPDEBUG}
          CodeSite.Send( 'SendText=' + aLowCmd.SendCmdText );
{$ENDIF}
          aLowCmd.LowCmdStatus := 'P';
          Inc( aLowProcCount );
        end;
      except
        on E: Exception do
        begin
          aSocketErr := True;
          Inc( FSocketErrCount );
          SetWMSocket( smrSocketError, E.Message );
          FTextSubj.ErrorFmt( LanguageManager.Get( 'SSocketWriteError' ), [E.Message] );
        end;
      end;
      { 有傳送過指令, 不管是否有錯誤, 記錄最後一次傳送指令時間 }
      FLastSendCmdTime := Now;
      if ( aSocketErr ) then
      begin
        aLowCmd.LowCmdStatus := 'E';
        aLowCmd.ErrCode := IntToStr( FWMSocket.ErrorCode );
        aLowCmd.ErrMsg := FWMSocket.ErrorDescription;
      end;
      { 將低階指令移至 Log List 裏 }
      LCmdLogList.AddObject( aLowCmd.UniqueId, aLowCmd );
      { 更新狀態 }
      FCmdSubj.Notify( aLowCmd );
      { 控制傳送的速度 }
      Sleep( Trunc( CommEnv.CASendDelay * 1000 )  );
    end;
    FDvnCmdBuilder.Clear;
    { 處理完低階指令 }
    if ( aSocketErr ) then
    begin
      { 若是有誤無法傳送, 直接將高階指令移至完成的 List 裏 }
      aHighCmd.HighCmdStatus := 'E';
      aHighCmd.ErrCode := IntToStr( FWMSocket.ErrorCode );
      aHighCmd.ErrMsg := FWMSocket.ErrorDescription;
      FCmdSubj.Notify( aHighCmd );
      { 從已傳送佇列移走 }
      HCmdWorkingList.Delete( aHighCmd );
      { 移進已完成佇列 }
      HCmdFinishList.AddObject( aHighCmd.UniqueId, aHighCmd );
    end else
    begin
      Inc( aHighProcCount );
    end;
    if ( FSocketErrCount > CommEnv.CAMaxError ) then Break;
    if ( Self.Terminated ) then Break;
  end;
  if ( aHighProcCount > 0 ) then
    FTextSubj.InfoFmt( LanguageManager.Get( 'SSocketSendRecords' ),
      [aHighProcCount, aLowProcCount]  );
end;

{ ---------------------------------------------------------------------------- }

{ TRecvThread }

constructor TRecvThread.Create(CreateSuspended: Boolean);
begin
  inherited Create( CreateSuspended );
  RecvThread := Self;
  FRecv := False;
  FTextSubj := TTextSubject.Create;
  FCmdSubj := TSendDvnSubject.Create;
  FAckParser := TDvnAckParser.Create;
  SetWMSocket( smrStop );
end;

{ ---------------------------------------------------------------------------- }

destructor TRecvThread.Destroy;
begin
  FTextSubj.Free;
  FCmdSubj.Free;
  inherited;
  RecvThread := nil;
end;

{ ---------------------------------------------------------------------------- }

procedure TRecvThread.WndProc(var Msg: TMessage);
begin
  case Msg.Msg of
    WM_SOCKET:
      begin
        FRecv :=  ( TWMSocket( Msg ).SocketMsgReason in
          [Ord( smrStart ), Ord( smrAck )] );
      end;
  end;
  inherited WndProc( Msg );
end;

{ ---------------------------------------------------------------------------- }

procedure TRecvThread.SetWMSocket(aReason: TSocketMsgReason;
  aErrText: String);
begin
  FWMSocket.Msg := WM_SOCKET;
  FWMSocket.ThreadType := Ord( tdRecv );
  FWMSocket.SocketMsgReason := Ord( aReason );
  FWMSocket.ErrorCode := 0;
  FWMSocket.ErrorDescription := nil;
  if ( aReason = smrSocketError ) then
  begin
    FWMSocket.ErrorCode := -1000;
    if ( aErrText <> EmptyStr ) then
      FWMSocket.ErrorDescription := PChar( aErrText );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TRecvThread.Execute;
begin
  FAckParser.CmdErrCnv.Data := CmdErrCnv.Data;
  while ( not Self.Terminated ) do
  begin
    if ( FSocket.Connected ) and ( FRecv ) then
      RecvAck;
    if ( Self.Terminated ) then Break;
    ProcAck;
    if ( Self.Terminated ) then Break;
    Sleep( 100 );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TRecvThread.RecvAck;
var
  aByte: Integer;
  aData: String;
begin
  aByte := -1;
  try
    aByte := FSocket.ReadFromStack( False, -1, False );
  except
    { ... }
  end;
  if aByte > 0 then
  begin
    SetString( aData, PChar( FSocket.InputBuffer.Memory ), aByte );
    FSocket.InputBuffer.Remove( aByte );
{$IFDEF APPDEBUG}
    CodeSite.Send( 'RecvText=' + aData );
{$ENDIF}
    BufferToList( aData );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TRecvThread.BufferToList(aRecvText: String);
var
  aAck: TSendDvn;

  { --------------------------------------------- }

  function IsIgnor(aText: String): Boolean;
  begin
    Result := ( UpperCase( aText ) = 'DISCONTINUITY_FRAME_NO' );
  end;

  { --------------------------------------------- }

begin
  repeat
    aAck := TSendDvn.Create( ctLow );
    aAck.RecvTime := Now;
    aAck.RecvCmdText := ExtractValue( aRecvText, ')' ) + ')';
    FAckParser.Parse( aAck );
    { ErrCode = 'DISCONTINUITY_FRAME_NO' 表示此指令回應不須處理 }
    if ( IsIgnor( aAck.Ack ) ) then
    begin
      FTextSubj.WarningFmt( LanguageManager.Get( 'SSocketAckWarning1' ),
       [aAck.Ack, aAck.LowCmd, aAck.FrameNo] );
      aAck.Free;
    end else
      LCmdAckList.AddObject( aAck.UniqueId, aAck );
  until ( aRecvText = EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

procedure TRecvThread.ProcAck;
var
  aAck, aLowCmd, aHighCmd: TSendDvn;
  aIndex, aCmdIdx: Integer;
  aTmp: String;
begin
  aIndex := 0;
  while ( aIndex < LCmdAckList.Count ) do
  begin
    aAck := LCmdAckList.Objects[aIndex];
    { 找該Ack回來的對應的低階指令 }
    aCmdIdx := LCmdLogList.IndexOf( aAck.UniqueId );
    if ( aCmdIdx < 0 ) then
    begin
      Inc( aIndex );
      Continue;
    end;
    aLowCmd := LCmdLogList.Objects[aCmdIdx];
    aLowCmd.RecvCmdText := aAck.RecvCmdText;
    aLowCmd.RecvTime := aAck.RecvTime;
    aLowCmd.LowCmdStatus := aAck.LowCmdStatus;
    aLowCmd.ErrCode := aAck.ErrCode;
    aLowCmd.ErrMsg := aAck.ErrMsg;
    aLowCmd.Ack := aAck.Ack;
    {}
    FCmdSubj.Notify( aLowCmd );
    Sleep( 100 );
    { 釋放 Ack }
    LCmdAckList.Delete( aIndex );
    aAck.Free;
    {}
    { 找該低階指令的原高階指令 }
    aCmdIdx := HCmdWorkingList.IndexOf( aLowCmd.HighSeqNo );
    if ( aCmdIdx >= 0 ) then
    begin
      aHighCmd := HCmdWorkingList.Objects[aCmdIdx];
      { 必須該高階指令狀態尚未錯誤, 才更新其狀態 }
      if ( aHighCmd.HighCmdStatus <> 'E' ) then
      begin
        aHighCmd.HighCmdStatus := aLowCmd.LowCmdStatus;
        aHighCmd.ErrCode := aLowCmd.ErrCode;
        aHighCmd.ErrMsg := aLowCmd.ErrMsg;
      end;
      { ADD_PRODUCT 及 REMOVE_PRODUCT 會有 Ack 回來 2 筆的現象, 一模一樣的 Ack }
      if ( AnsiPos( aLowCmd.FrameNo, aHighCmd.LowCmdNote ) > 0 ) then
      begin
        aTmp := aHighCmd.LowCmdNote;
        ExtractValue( aTmp );
        aHighCmd.LowCmdNote := aTmp;
        { 該高階對應送出的低階指令數 }
        aHighCmd.LowCmdCount := ( aHighCmd.LowCmdCount - 1 );
      end;   
      { 已全部 Ack 回來, 或是該低階指令已經是錯誤, 直接更新畫面上該高階指令的狀態 }
      if ( aHighCmd.HighCmdStatus = 'E' ) or ( aHighCmd.LowCmdCount <= 0 ) then
      begin
        { 該高階指令收到最後一筆低階指令的時間,代表該高階指令完成時間 }
        aHighCmd.RecvTime := aLowCmd.RecvTime;
        FCmdSubj.Notify( aHighCmd );
        Sleep( 100 );
      end;
      { 所有低階指令都 Ack 回來 }
      if ( aHighCmd.LowCmdCount <= 0 ) then
      begin
        { 刪除處理中佇列 }
        HCmdWorkingList.Delete( aHighCmd );
        { 移到已完成佇列 }
        HCmdFinishList.AddObject( aHighCmd.HighSeqNo, aHighCmd );
      end;
      { 低階指令不可因為 CA Ack 回來就釋放,
        SoDbThread 裏面會去寫 Log, 寫完才會釋放 }
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
