unit cbMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  ImgList, ExtCtrls, StdCtrls, ComCtrls, CommCtrl, SyncObjs,
  { Indy }
  IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient,
  {$IFDEF DEBUG} CsIntf, {$ENDIF}
  { App }
  cbStyleModule, cbConfigModule, cbLiceMgr, cbClass,
  { Developer Express Suite }
  dxBar, dxBarExtItems, dxDockControl, dxDockPanel, cxControls,
  cxContainer, cxListView, cxPC, cxGraphics, cxCustomData, cxStyles, cxTL,
  cxInplaceContainer, cxTextEdit, cxEdit, cxLabel;


type
  TfmMain = class(TForm)
    dxBarManager: TdxBarManager;
    dxFiles: TdxBarSubItem;
    dxAbout: TdxBarSubItem;
    dxPlay: TdxBarButton;
    dxPause: TdxBarButton;
    dxStop: TdxBarButton;
    dxComputer: TdxBarStatic;
    dxReload: TdxBarButton;
    dxServer: TdxBarStatic;
    dxLED: TdxBarControlContainerItem;
    LedPanel: TPanel;
    Image1: TImage;
    Image2: TImage;
    dxDockingManager1: TdxDockingManager;
    dxDockHost: TdxDockSite;
    dxDockPanel1: TdxDockPanel;
    dxDockPanel2: TdxDockPanel;
    dxLayoutDockSite2: TdxLayoutDockSite;
    dxTabContainerDockSite1: TdxTabContainerDockSite;
    dxDockPanel3: TdxDockPanel;
    dxDockPanel5: TdxDockPanel;
    dxLayoutDockSite1: TdxLayoutDockSite;
    dxLayoutDockSite3: TdxLayoutDockSite;
    AutoExecTimer: TTimer;
    cxPageControl1: TcxPageControl;
    PlayCmdGroup: TdxBarGroup;
    StopCmdGroup: TdxBarGroup;
    StartupTimer: TTimer;
    dxDockPanel6: TdxDockPanel;
    dxDockPanel4: TdxDockPanel;
    dxTabContainerDockSite3: TdxTabContainerDockSite;
    StartupMsgList: TcxListView;
    Panel1: TPanel;
    Panel2: TPanel;
    SoDbMsgList: TcxListView;
    SoTree: TcxTreeList;
    SoTreecxTreeListColumn2: TcxTreeListColumn;
    SoTreecxTreeListColumn1: TcxTreeListColumn;
    SoTreecxTreeListColumn3: TcxTreeListColumn;
    ControlSendTree: TcxTreeList;
    ControlSendSocket: TIdTCPClient;
    dxDockPanel7: TdxDockPanel;
    dxTabContainerDockSite2: TdxTabContainerDockSite;
    ControlMsgList: TcxListView;
    lblConfigFileName: TcxLabel;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormDestroy(Sender: TObject);
    procedure StartupTimerTimer(Sender: TObject);
    procedure AutoExecTimerTimer(Sender: TObject);
    procedure dxPlayClick(Sender: TObject);
    procedure dxStopClick(Sender: TObject);
    procedure OnStartupMsgNotify(aSubject: TSubject);
    procedure OnSoDbMsgNotify(aSubject: TSubject);
    procedure OnControlMsgNotify(aSubject: TSubject);
  private
    { Private declarations }
    FLicenceManager: TLicenceManager;
    FStyleModule: TStyleModule;
    FConfigModule: TConfigModule;
    FSoThreadManager: TThreadManager;
    FControlManager: TThreadManager;
    FStartupObServer: TObServer;
    FSoDbObServer: TObServer;
    FControlObServer: TObServer;
    FMsgSubject: TMessageSubject;
    procedure InitiateUI;
    procedure InitiateUILanguage;
    procedure CreateControlThread;
    procedure ReleaseControlThread;
    procedure CreateSoDbThread;
    procedure ReleaseSoDbThread;
    procedure AllocateTaskResource;
    procedure CleanupTaskResource;
  public
    { Public declarations }
    property StartupObServer: TObServer read FStartupObServer;
    property SoDbObServer: TObserver read FSoDbObServer;
    procedure BuildSoTree;
    procedure ClearSoTree;
    procedure BeginExecute;
    procedure EndExecute;
    procedure UpdateSoTreeStatus(aNodeIndex: Integer);
  end;

var
  fmMain: TfmMain;

  { 以下是全域變數, 因為要在 3 個 Thread 中共享, 所以用全域變數方式宣告 }

  { 準備傳送出去的指令 }
  { ControlSend Thread 會負責把它送出去 }
  ControlSendManager: TCommandListMagager;

  { 已經送出去的指令, 由 SoDbThread 讀出來後更新回 Send_Nagar }
  ControlSendDoneManager: TCommandListMagager;

  { 送出去的指令, 當 Nagra Ack 回來後, 會放到這裏,
    SoDbThread 會負責將它讀出來後, 更新回 Send_Nagar  }
  ControlRecvManager: TCommandListMagager;

  { 傳送指令出去的 Log 通通會在這裏, 包含 Send 出去 及 Ack 回來的 Log,
    SoDbThread 會負責將它讀出來後寫進 Log Table }
  ControlSendLogManager: TCommandListMagager;

  { 以上 4 個全域變數以 Main Form 負責建立及釋放 }

implementation

{$R *.dfm}

uses cbResStr, cbUtils, cbSoDbThread, cbControlThread;


{ ---------------------------------------------------------------------------- }

function GetMsgImageIndex(const aType: Longint): Integer;
begin
  case aType of
    MB_OK: Result := 1;
    MB_ICONERROR: Result := 2;
  else
    Result := 0;  
  end;
end;

{ ---------------------------------------------------------------------------- }

function GetSoImageIndex(const aStatus: TDbConnectStatus): Integer;
begin
  case aStatus of
    dbOK, dbActive: Result := 3;
    dbError: Result := 4;
    dbNone: Result := 2;
  else
    Result := 1;
  end;
end;

{ ---------------------------------------------------------------------------- }

function GetSoStateIndex(const aStatus: TDbConnectStatus): Integer;
begin
  case aStatus of
    dbOK, dbError: Result := 6;
    dbActive: Result := 7;
  else
    Result := 5;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.OnStartupMsgNotify(aSubject: TSubject);
var
  aText: String;
  aItem: TListItem;
begin
  if aSubject.RunState in [rsStartup, rsStop] then
  begin
    aText := Format( ' %s >   %s', [FormatDateTime( 'yyyy-mm-dd hh:mm', Now ),
      TMessageSubject( aSubject ).NotifyMessage] );
    aItem := StartupMsgList.Items.Add;
    aItem.Caption := aText;
    aItem.ImageIndex := GetMsgImageIndex( TMessageSubject( aSubject ).NotifyType );
    aItem.MakeVisible( True );
    Application.ProcessMessages;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.OnSoDbMsgNotify(aSubject: TSubject);
var
  aText: String;
  aItem: TListItem;
begin
  if aSubject.RunState in [rsRunning] then
  begin
    aText := Format( ' %s >   %s', [FormatDateTime( 'yyyy-mm-dd hh:mm', Now ),
      TMessageSubject( aSubject ).NotifyMessage] );
    aItem := SoDbMsgList.Items.Add;
    aItem.Caption := aText;
    aItem.ImageIndex := GetMsgImageIndex( TMessageSubject( aSubject ).NotifyType );
    aItem.MakeVisible( True );
    Application.ProcessMessages;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.OnControlMsgNotify(aSubject: TSubject);
var
  aText: String;
  aItem: TListItem;
begin
  if aSubject.RunState in [rsRunning] then
  begin
    aText := Format( ' %s >   %s', [FormatDateTime( 'yyyy-mm-dd hh:mm', Now ),
      TMessageSubject( aSubject ).NotifyMessage] );
    aItem := ControlMsgList.Items.Add;
    aItem.Caption := aText;
    aItem.ImageIndex := GetMsgImageIndex( TMessageSubject( aSubject ).NotifyType );
    aItem.MakeVisible( True );
    Application.ProcessMessages;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormCreate(Sender: TObject);
begin
  FStyleModule := TStyleModule.Create( nil );
  FSoThreadManager := TThreadManager.Create;
  FControlManager := TThreadManager.Create;
{$IFDEF DEBUG}
   if ( Screen.MonitorCount > 1 ) then
   begin
     Self.Position := poDesigned;
     Self.Left := Screen.Monitors[1].Left + ( ( Screen.Monitors[1].Width - Width ) div 2 );
     Self.Top := Screen.Monitors[1].Top + ( ( Screen.Monitors[1].Height - Height) div 2);
     Self.WindowState := wsMaximized;
   end;
{$ENDIF}
  ControlSendManager := TCommandListMagager.Create;
  ControlSendDoneManager := TCommandListMagager.Create;
  ControlRecvManager := TCommandListMagager.Create;
  ControlSendLogManager := TCommandListMagager.Create;
  { 回饋訊息 }
  FStartupObServer := TObServer.Create;
  FStartupObServer.OnUpdate := OnStartupMsgNotify;
  FSoDbObServer := TObServer.Create;
  FSoDbObServer.OnUpdate := OnSoDbMsgNotify;
  FMsgSubject := TMessageSubject.Create;
  FMsgSubject.AddObServer( FStartupObServer );
  FMsgSubject.AutoNotify := True;
  { Control Thread }
  FControlObServer := TObServer.Create;
  InitiateUI;
  InitiateUILanguage;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormShow(Sender: TObject);
begin
  StartupTimer.Enabled := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  CanClose := ConfirmMsg( SDlgConfimQuit );
  if CanClose then
  begin
    FMsgSubject.NotifyMessage := SAppCloseing;
    ReleaseSoDbThread;
    CleanupTaskResource;    
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormDestroy(Sender: TObject);
begin
  FSoThreadManager.Free;
  FControlManager.Free;
  FStyleModule.Free;
  FConfigModule.Free;
  ControlSendManager.Free;
  ControlSendDoneManager.Free;
  ControlRecvManager.Free;
  FSoDbObServer.Free;
  FStartupObServer.Free;
  FControlObServer.Free;
  FMsgSubject.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.InitiateUI;
begin
  ClearSoTree;
  StartupMsgList.Clear;
  PlayCmdGroup.Enabled := False;
  StopCmdGroup.Enabled := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.InitiateUILanguage;
begin
  Caption := Format( '%s-%s', [SVendor, SAppTitle] );
  lblConfigFileName.Caption := Format( SlblConfigFileName,
    [FConfigModule.AppStartupInfo.ConfigFileName] );    
  dxDockPanel1.Caption := SDockPanel1;
  dxDockPanel2.Caption := SDockPanel2;
  dxDockPanel3.Caption := SDockPanel3;
  dxDockPanel4.Caption := SDockPanel4;
  dxDockPanel5.Caption := SDockPanel5;
  dxDockPanel6.Caption := SDockPanel6;
  dxDockPanel7.Caption := SDockPanel7;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.CreateControlThread;
var
  aThread: TGateWaySocketThread;
begin
  FMsgSubject.NotifyMessage := SControlSendOpen;
  Delay( COMMON_DELAY );
  aThread := TControlSendThread.Create( ControlSendSocket );
  aThread.CommEnv := FConfigModule.CommEnv;
  aThread.CmdEnv := FConfigModule.CmdEnv;
  TControlSendThread( aThread ).CASEnv := FConfigModule.CASEnv;
  aThread.MessageSubject.AddObServer( FStartupObServer );
  aThread.MessageSubject.AddObServer( FControlObServer );
  if not TControlSendThread( aThread ).BeginConnection then
  begin
    aThread.Resume;
    aThread.Terminate;
    aThread.WaitFor;
    Exit;
  end;
  FControlManager.Add( aThread );
  FMsgSubject.NotifyMessage := SControlSendOpenReady;
  Delay( COMMON_DELAY );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.ReleaseControlThread;
var
  aIndex: Integer;
begin
  if FControlManager.Count = 0 then Exit;
  FMsgSubject.NotifyMessage := SControlSendClose;
  Delay( COMMON_DELAY );
  { 先打出停止的訊息 }
  for aIndex := 0 to FControlManager.Count - 1 do
    PostMessage( TGateWaySocketThread(
      FControlManager.Items[aIndex] ).WndHandle, WM_STOP, 0, 0 );
  Delay( 1000 );
  { 正式停止 }
  for aIndex := 0 to FControlManager.Count - 1 do
  begin
    try
      FControlManager.StopThread( aIndex );
      Delay( 100 );
      FMsgSubject.NotifyMessage := SControlSendCloseDone;
    except
      on E: Exception do
      begin
        FMsgSubject.NotifyMessage := Format( SControlSendCloseError, [E.Message] );
        Delay( 100 );
      end;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.CreateSoDbThread;
var
  aSoDbThread: TSoDbThread;
  aIndex: Integer;
  aList: TList;
begin
  { 各系統台自獨立的執行緒 }
  if FConfigModule.CommEnv.DbMultiThread then
  begin
    aList := FConfigModule.SoList.LockList;
    try
      for aIndex := 0 to aList.Count - 1 do
      begin
        if PSoInfo( aList.Items[aIndex] ).Selected then
        begin
          aSoDbThread := TSoDbThread.Create( PSoInfo( aList.Items[aIndex] ) );
          aSoDbThread.CommEnv := FConfigModule.CommEnv;
          aSoDbThread.CmdEnv := FConfigModule.CmdEnv;
          aSoDbThread.MessageSubject.AddObServer( FStartupObServer );
          aSoDbThread.MessageSubject.AddObServer( FSoDbObServer );
          FSoThreadManager.Add( aSoDbThread );
        end;
      end;
    finally
      FConfigModule.SoList.UnlockList;
    end;
  end else
  begin
    aSoDbThread := TSoDbThread.Create( FConfigModule.SoList );
    aSoDbThread.CommEnv := FConfigModule.CommEnv;
    aSoDbThread.CmdEnv := FConfigModule.CmdEnv;
    aSoDbThread.MessageSubject.AddObServer( FStartupObServer );
    aSoDbThread.MessageSubject.AddObServer( FSoDbObServer );
    FSoThreadManager.Add( aSoDbThread );
  end;
  FMsgSubject.NotifyMessage := SDbConnectStart;
  Delay( COMMON_DELAY );
  { 恢復執行緒的暫停狀態 }
  for aIndex := 0 to FSoThreadManager.Count - 1 do
     TSoDbThread( FSoThreadManager.Items[aIndex] ).Resume;
  Delay( 5000 );
  { 打出開始執行的訊號 }
  for aIndex := 0 to FSoThreadManager.Count - 1 do
    PostMessage( TSoDbThread( FSoThreadManager.Items[aIndex] ).WndHandle, WM_PLAY, 0, 0 );
  Delay( COMMON_DELAY );
  FMsgSubject.NotifyMessage := SDbConnectEnd;
  Delay( COMMON_DELAY );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.ReleaseSoDbThread;
var
  aIndex: Integer;
begin
  if FSoThreadManager.Count = 0 then Exit;
  FMsgSubject.NotifyMessage := SDbDisConnectStart;
  Delay( COMMON_DELAY );
  { 先打出停止的訊息 }
  for aIndex := 0 to FSoThreadManager.Count - 1 do
    PostMessage( TSoDbThread( FSoThreadManager.Items[aIndex] ).WndHandle, WM_STOP, 0, 0 );
  Delay( 3000 );
  { 正式停止 }
  for aIndex := 0 to FSoThreadManager.Count - 1 do
  begin
    try
      FSoThreadManager.StopThread( aIndex );
      Delay( 100 );
      FMsgSubject.NotifyMessage := SDbDisConnectEnd;
    except
      on E: Exception do
      begin
        FMsgSubject.NotifyMessage := Format( SDbDisConnectThreadError, [E.Message] );
        Delay( 100 );
      end;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.StartupTimerTimer(Sender: TObject);
begin
  StartupTimer.Enabled := False;
  Screen.Cursor := crHourGlass;
  try
    FMsgSubject.NotifyMessage := SAppStartup;
    Delay( 500 );
    FConfigModule := TConfigModule.Create;
    ClearSoTree;
    if ( FConfigModule.SoCount > 0 ) then
    begin
      FMsgSubject.NotifyMessage := SConfigBuildSoTreeStart;
      Delay( 500 );
      BuildSoTree;
      FMsgSubject.NotifyMessage := SConfigBuildSoTreeEnd;
      Delay( 1000 );
      AutoExecTimer.Enabled := FConfigModule.AppStartupInfo.AutoExecute;
      if AutoExecTimer.Enabled then
        FMsgSubject.NotifyMessage := SAutoExec
      else
        FMsgSubject.NotifyMessage := SNoneAutoExec;
      PlayCmdGroup.Enabled := True;  
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.AutoExecTimerTimer(Sender: TObject);
begin
  { 開始執行, 將 Start 這個 Group 的 Button Enable 起來 }
  dxPlay.Click;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.BuildSoTree;

  { -------------------------------------------------- }

  function GetPosNode(const aPosName: String; const aCanAdd: Boolean = True):
    TcxTreeListNode;
  begin
    Result := nil;
    if ( aPosName = EmptyStr ) then Exit;
    Result := SoTree.FindNodeByText( aPosName, SoTreecxTreeListColumn2 );
    if not Assigned( Result ) and aCanAdd then
    begin
      Result := SoTree.Add( nil );
      Result.AssignValues( [EmptyStr, aPosName, EmptyStr] );
      Result.ImageIndex := 0;
      Result.SelectedIndex := 0;
    end;
  end;

  { -------------------------------------------------- }

var
  aIndex: Integer;
  aList: TList;
  aSo: PSoInfo;
  aNodePos, aNodeComp: TcxTreeListNode;
begin
  aList := FConfigModule.SoList.LockList;
  try
    for aIndex := 0 to aList.Count - 1 do
    begin
      aSo := PSoInfo( aList.Items[aIndex] );
      { 取平台節點, 此系統台要掛在那一個平台下 }
      aNodePos := GetPosNode( aSo.PosName );
      aNodeComp := SoTree.AddChild( aNodePos, aSO );
      aNodeComp.AssignValues( [aSo.CompCode, aSo.CompName, EmptyStr] );
      aNodeComp.ImageIndex := GetSoImageIndex( aSo.DbConnectStatus );
      aNodeComp.StateIndex := GetSoStateIndex( aSo.DbConnectStatus );
      aNodeComp.SelectedIndex := aNodeComp.ImageIndex;
      aNodePos.Expanded := True;
      FMsgSubject.NotifyMessage := Format( SConfigBuildSoTree, [aSo.CompName] );
      Delay( COMMON_DELAY );
    end;
    { 將所有系統的樹狀節點重新對應一次 }
    for aIndex := 0 to SoTree.Nodes.Count - 1 do
    begin
      if SoTree.Nodes.Items[aIndex].Level > 0 then
        PSoInfo( SoTree.Nodes.Items[aIndex].Data ).ItemIndex := aIndex;
    end;
  finally
    FConfigModule.SoList.UnlockList;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.ClearSoTree;
var
  aIndex: Integer;
begin
  SoTree.BeginUpdate;
  try
    for aIndex := 0 to SoTree.Nodes.Count - 1 do
    begin
      if ( SoTree.Nodes.Items[aIndex].Level > 0 ) then
        SoTree.Nodes.Items[aIndex].Data := nil;
    end;
  finally
    SoTree.EndUpdate;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxPlayClick(Sender: TObject);
begin
  BeginExecute;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxStopClick(Sender: TObject);
begin
  EndExecute;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.BeginExecute;
begin
  AllocateTaskResource;
  CreateControlThread;
  //CreateSoDbThread;
  StopCmdGroup.Enabled := True;
  PlayCmdGroup.Enabled := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.EndExecute;
begin
  ReleaseControlThread;
  //ReleaseSoDbThread;
  CleanupTaskResource;
  StopCmdGroup.Enabled := False;
  PlayCmdGroup.Enabled := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.UpdateSoTreeStatus(aNodeIndex: Integer);

  { --------------------------------------- }

  procedure UpdateNode(aNode: TcxTreeListNode);
  var
    aSoInfo: PSoInfo;
  begin
    if Assigned( aNode ) then
    begin
      aSoInfo := PSoInfo( aNode.Data );
      if Assigned( aSoInfo ) then
      begin
        aNode.ImageIndex := GetSoImageIndex( aSoInfo.DbConnectStatus );
        aNode.SelectedIndex := aNode.ImageIndex;
        aNode.StateIndex := GetSoStateIndex( aSoInfo.DbConnectStatus );
        if aSoInfo.DbConnectStatus = dbNone then
          aNode.Values[SoTreecxTreeListColumn3.ItemIndex] := EmptyStr
        else
          aNode.Values[SoTreecxTreeListColumn3.ItemIndex] := aSoInfo.RecordCount;
      end;
    end;
  end;

  { --------------------------------------- }

var
  aIndex: Integer;
  aUseLock: Boolean;
begin
  aUseLock := False;
  if aUseLock then FSoThreadManager.Lock;
  try
     if ( aNodeIndex > 0 ) then
       UpdateNode( SoTree.Nodes.Items[aNodeIndex] )
    else begin
      for aIndex := 0 to SoTree.Nodes.Count - 1 do
      begin
        if SoTree.Nodes.Items[aIndex].Level = 0 then Continue;
        if PSoInfo( SoTree.Nodes.Items[aIndex].Data ).Selected then
          UpdateNode( SoTree.Nodes.Items[aIndex] );
      end;
    end;
  finally
    if aUseLock then FSoThreadManager.Unlock;
  end;
end;


(*
procedure TfmMain.UpdateSoTreeStatus(aNodeIndex: Integer);

  { --------------------------------------- }

  procedure UpdateNode(aNode: TTreeNode);
  var
    aDisplay: String;
    aSoInfo: PSoInfo;
  begin
    {$IFDEF DEBUG}
      if not Assigned( aNode ) then
      begin
        CodeSite.SendError( 'UpdateSoTreeStatus Error, The TreeNode Is Null.' );
      end;
    {$ENDIF}
    if Assigned( aNode ) then
    begin
      aSoInfo := PSoInfo( aNode.Data );
      {$IFDEF DEBUG}
        if not Assigned( aSoInfo ) then
        begin
          CodeSite.SendError( 'UpdateSoTreeStatus Error, The TreeNode.Data Is Null.' );
        end;
      {$ENDIF}
      if Assigned( aSoInfo ) then
      begin
        aNode.ImageIndex := GetSoImageIndex( aSoInfo.DbConnectStatus );
        aNode.SelectedIndex := aNode.ImageIndex;
        aNode.StateIndex := GetSoStateIndex( aSoInfo.DbConnectStatus );
        aDisplay := ( Rpad( Lpad( aSoInfo.CompCode, 3, #32 ), 5, #32 ) + aSoInfo.CompName );
        if aSoInfo.DbConnectStatus in [dbOK, dbActive] then
          aDisplay := Format( '%s(%d)', [Rpad( aDisplay, 12, #32 ), aSoInfo.RecordCount] );
        aNode.Text := aDisplay;
      end;
    end;
  end;

  { --------------------------------------- }

var
  aIndex: Integer;
  aUseLock: Boolean;
begin
  aUseLock := False;
  if aUseLock then FSoThreadManager.Lock;
  try
     if ( aNodeIndex > 0 ) then
       UpdateNode( SoTree.Items[aNodeIndex] )
    else begin
      for aIndex := 0 to SoTree.Items.Count - 1 do
      begin
        if SoTree.Items[aIndex].Level = 0 then Continue;
        if PSoInfo( SoTree.Items[aIndex].Data ).Selected then
          UpdateNode( SoTree.Items[aIndex] );
      end;
    end;
  finally
    if aUseLock then FSoThreadManager.Unlock;
  end;
end;
*)

{ ---------------------------------------------------------------------------- }

procedure TfmMain.AllocateTaskResource;
var
  aList: TList;
  aIndex: Integer;
  aCmdList: TThreadStringList;
begin
  aList := FConfigModule.SoList.LockList;
  try
    for aIndex := 0 to aList.Count - 1 do
    begin
      if PSoInfo( aList.Items[aIndex] ).Selected then
      begin
        { 1. 建立 ControlSend 指令傳送的 List, 各系統台獨立 }
        aCmdList := TThreadStringList.Create;
        { ControlSendManager 的識別碼是 SEQNO, 而且有指令傳送先後順序的問題,
          所以 Sort 的屬性不可打開 }
        aCmdList.Sorted := False;
        ControlSendManager.Add( aCmdList );
        { 2. 建立 AlreadSend 已經傳送指令出去傳送的 List, 各系統台獨立 }
        aCmdList := TThreadStringList.Create;
        { AlreadySendManager, ControlAckManager是處理已經送出去的指令, 所以
          識別碼是 TransactionNumber, 可以打開 Sort 屬性加快尋搜尋速度 }
        aCmdList.Sorted := True;
        ControlSendDoneManager.Add( aCmdList );
        { 3. 建立 Ack 回來的 List, 各系統台獨立 }
        aCmdList := TThreadStringList.Create;
        aCmdList.Sorted := True;
        ControlRecvManager.Add( aCmdList );
        { 4. 建立寫 Log List, 各系統台獨立 }
        aCmdList := TThreadStringList.Create;
        aCmdList.Sorted := True;
        ControlSendLogManager.Add( aCmdList );
      end;
    end;
  finally
    FConfigModule.SoList.UnlockList;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.CleanupTaskResource;
var
  aManagerIndex, aIndex, aIndex2: Integer;
  aCommandManager: TCommandListMagager;
  aManagerName: String;
begin
  { 在這裏最後清除尚未處理的 List 資料 }
  for aManagerIndex := 1 to 4 do
  begin
    try
      case aManagerIndex of
        1: begin
             aCommandManager := ControlSendManager; { 1.清除傳送指令的 List }
             aManagerName := SControlSendManager;
           end;
        2: begin
             aCommandManager := ControlSendDoneManager; { 2.指令已送出去的 List }
             aManagerName := SControlSendDoneManager;
           end;
        3: begin
             aCommandManager := ControlRecvManager;  { 3.指令Ack回來的 List }
             aManagerName := SControlRecvManager;
           end;
        4: begin
             aCommandManager := ControlSendLogManager;
             aManagerName := SControlSendLogManager;
           end;
      else
        aCommandManager := nil;      
      end;
      for aIndex := 0 to aCommandManager.ItemCount - 1 do
      begin
        for aIndex2 := 0 to aCommandManager.Items[aIndex].Count - 1 do
        begin
          Dispose( PSendNagra( aCommandManager.Items[aIndex].Objects[aIndex2] ) );
          aCommandManager.Items[aIndex].Objects[aIndex2] := nil;
          {$IFDEF DEBUG}
            CodeSite.SendError( Format( '%s still have item', [aManagerName] ) );
          {$ENDIF}
        end;
      end;
      aCommandManager.Clear;
    except
      on E: Exception do
      begin
        FMsgSubject.NotifyMessage := Format( SCleanupCommandManagerError,
          [aManagerName, E.Message] );
        Delay( 100 );
      end;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }
end.
