unit cbMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  ImgList, ExtCtrls, StdCtrls, ComCtrls, CommCtrl, IniFiles,
  {$IFDEF APPDEBUG} CodeSiteLogging, {$ENDIF}
  { App }
  cbAppClass, cbStyleModule, cbConfigModule, cbSoDbThread, cbSendThread,
  { Developer Express Suite }
  dxBar, dxBarExtItems, dxDockControl, dxDockPanel, cxControls,
  cxContainer, cxListView, cxPC, cxGraphics, cxCustomData, cxStyles, cxTL,
  cxInplaceContainer, cxImageComboBox, cxEdit, cxMaskEdit, cxDropDownEdit,
  cxTextEdit, cxCheckBox, dxSkinsCore, dxSkinsDefaultPainters,
  dxSkinsdxDockControlPainter, cxClasses, dxSkinsdxBarPainter;


type
  TfmMain = class(TForm)
    dxBarManager: TdxBarManager;
    dxMenuItem1: TdxBarSubItem;
    dxPlay: TdxBarButton;
    dxStop: TdxBarButton;
    dxAutoScroll: TdxBarButton;
    dxConsoleTree: TdxBarButton;
    dxCmdExpand: TdxBarButton;
    dxCmdCollapse: TdxBarButton;
    dxStatusItem1: TdxBarStatic;
    dxTabContainerDockSite2: TdxTabContainerDockSite;
    dxTabContainerDockSite3: TdxTabContainerDockSite;
    dxDockingManager: TdxDockingManager;
    dxDockHost: TdxDockSite;
    dxLayoutDockSite1: TdxLayoutDockSite;
    dxLayoutDockSite2: TdxLayoutDockSite;
    dxLayoutDockSite3: TdxLayoutDockSite;
    dxDockPanel1: TdxDockPanel;
    dxDockPanel2: TdxDockPanel;
    dxDockPanel3: TdxDockPanel;
    dxDockPanel5: TdxDockPanel;
    dxDockPanel6: TdxDockPanel;
    dxDockPanel7: TdxDockPanel;
    PlayGroup: TdxBarGroup;
    StopGroup: TdxBarGroup;
    LoadConfigTimer: TTimer;
    MainMsgList: TcxListView;
    Panel1: TPanel;
    Panel2: TPanel;
    SoMsgList: TcxListView;
    SendMsgList: TcxListView;
    lblActiveConfig: TLabel;
    SoTree: TcxTreeList;
    SoTreeCol1: TcxTreeListColumn;
    SoTreeCol2: TcxTreeListColumn;
    SoTreeCol3: TcxTreeListColumn;
    ConsoleTree: TcxTreeList;
    ConsoleTreeCol1: TcxTreeListColumn;
    ConsoleTreeCol2: TcxTreeListColumn;
    ConsoleTreeCol3: TcxTreeListColumn;
    ConsoleTreeCol4: TcxTreeListColumn;
    ConsoleTreeCol5: TcxTreeListColumn;
    dxDbTree: TdxBarButton;
    dxSendTree: TdxBarButton;
    SendTree: TcxTreeList;
    SendTreeCol1: TcxTreeListColumn;
    SendTreeCol2: TcxTreeListColumn;
    SendTreeCol3: TcxTreeListColumn;
    SendTreeCol4: TcxTreeListColumn;
    SendTreeCol5: TcxTreeListColumn;
    SendTreeCol6: TcxTreeListColumn;
    SendTreeCol7: TcxTreeListColumn;
    SendTreeCol8: TcxTreeListColumn;
    SendTreeCol9: TcxTreeListColumn;
    SendTreeCol10: TcxTreeListColumn;
    SendTreeCol11: TcxTreeListColumn;
    dxMenuItem11: TdxBarButton;
    dxStatusItem2: TdxBarStatic;
    cmbConfigList: TcxImageComboBox;
    dxMenuItem22: TdxBarButton;
    dxMenuItem2: TdxBarSubItem;
    dxMenuItem23: TdxBarButton;
    dxMenuItem24: TdxBarButton;
    dxMenuItem25: TdxBarButton;
    dxMenuItem26: TdxBarButton;
    dxMenuItem27: TdxBarButton;
    dxMenuItem21: TdxBarButton;
    dxMenuItem3: TdxBarButton;
    SocketWarningTimer: TTimer;
    LicenceTimer: TTimer;
    SendTreeCol12: TcxTreeListColumn;
    ConsoleTreeCol6: TcxTreeListColumn;
    dxStatusItem3: TdxBarStatic;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormDestroy(Sender: TObject);
    procedure LoadConfigTimerTimer(Sender: TObject);
    procedure cmbConfigListPropertiesChange(Sender: TObject);
    procedure OnTextNotify(aSubject: TSubject);
    procedure OnSoTextNotify(aSubject: TSubject);
    procedure OnSendTextNotify(aSubject: TSubject);
    procedure OnSoNotify(aSubject: TSubject);
    procedure OnSendCmdNotify(aSubject: TSubject);
    procedure OnConsoleNotify(aSubject: TSubject);
    procedure dxPlayClick(Sender: TObject);
    procedure dxStopClick(Sender: TObject);
    procedure SendTreeCustomDrawCell(Sender: TObject; ACanvas: TcxCanvas;
      AViewInfo: TcxTreeListEditCellViewInfo; var ADone: Boolean);
    procedure dxCmdExpandClick(Sender: TObject);
    procedure dxCmdCollapseClick(Sender: TObject);
    procedure dxDbTreeClick(Sender: TObject);
    procedure dxSendTreeClick(Sender: TObject);
    procedure dxConsoleTreeClick(Sender: TObject);
    procedure dxAutoScrollClick(Sender: TObject);
    procedure ConsoleTreeKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure SocketWarningTimerTimer(Sender: TObject);
    procedure dxMenuItem21Click(Sender: TObject);
    procedure dxMenuItem22Click(Sender: TObject);
    procedure dxMenuItem11Click(Sender: TObject);
  private
    { Private declarations }
    FAppIni: TIniFile;
    FAppFirstRun: Boolean;
    FAutoScroll: Boolean;
    FDocklayoutStream: TMemoryStream;
    FTextSubj: TTextSubject;
    FTextObsrv: TObServer;
    FSoDbThread: TSoDbThread;
    FSoTextObsrv: TObServer;
    FSoObsrv: TObServer;
    FSendThread: TSendThread;
    FRecvThread: TRecvThread;
    FSendTextObsrv: TObServer;
    FSendCmdObsrv: TObServer;
    FConsoleObsrv: TObServer;
    function LoadLanguage: Boolean;
    function CreateAppIniFile: TIniFile;
    procedure AppointConfigFile;
    procedure InitControlState;
    procedure InitControlLanguage;
    procedure BuildSoDbTree;
    function CreateSoDbThread: TSoDbThread;
    function CreateSendThread: TSendThread;
    function CreateRecvThread: TRecvThread;
    procedure ReleaseSoDbThread;
    procedure ReleaseSendThread;
    procedure ReleaseRecvThread;
    procedure DisposeSyncList;
  protected
    procedure WMSocket(var Msg: TWMSocket); message WM_SOCKET;
  public
    { Public declarations }
    procedure BuildConfigFileList;
    procedure LoadWindowPositionFromFile;
    procedure SaveWindowPositionToFile;
    procedure LoadToolBarStateFromFile;
    procedure SaveToolBarStateToFile;
    procedure LoadDocklayoutFromFile;
    procedure SaveDocklayoutToFile;
    procedure BeginExecute;
    procedure EndExecute;
  end;


var
  fmMain: TfmMain;

  { 高階指令 }
  { 1. 先 Read from DB --> 放進 CmdBeginList -->
    2. 送指令 Send To DVN --> 移到 CmdWorkingList -->
    3. 傳送 OK 或有錯 Send Ok/Error --> 移到 CmdFinishList -->
    4. 然後更新回資料庫 Wirte to DB  }
  HCmdBeginList, HCmdWorkingList, HCmdFinishList: TSyncList;

  { 低階指令 }
  { 存放每一筆送出去的低階指令 }
  LCmdLogList: TSyncList;

  { CA->DVN Ack 回來的指令 }
  { 送指令到 DVN 後, Ack 回來的指令先全部放進這個 List,
    然後必須跟傳送出去的指令串列比對 LCmdLogList }
  LCmdAckList: TSyncList;

  IsThreadRun: Boolean;


implementation

{$R *.dfm}

uses cbUtilis;


{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormCreate(Sender: TObject);
begin
  Application.ShowMainForm := False;
  if not LoadLanguage then
  begin
    Application.Terminate;
    Exit;
  end;
  Application.ShowMainForm := True;
  FAppIni := CreateAppIniFile;
  FAppFirstRun := True;
  FDocklayoutStream := TMemoryStream.Create;
  FAutoScroll := False;
  IsThreadRun := False;
  {}
  FTextObsrv := TObServer.Create;
  FTextObsrv.OnUpdate := OnTextNotify;
  FTextSubj := TTextSubject.Create;
  FTextSubj.AddObServer( FTextObsrv );
  {}
  FSoTextObsrv := TObServer.Create;
  FSoTextObsrv.OnUpdate := OnSoTextNotify;
  FSoObsrv := TObServer.Create;
  FSoObsrv.OnUpdate := OnSoNotify;
  {}
  FSendTextObsrv := TObServer.Create;
  FSendTextObsrv.OnUpdate := OnSendTextNotify;
  FSendCmdObsrv := TObServer.Create;
  FSendCmdObsrv.OnUpdate := OnSendCmdNotify;
  FConsoleObsrv := TObServer.Create;
  FConsoleObsrv.OnUpdate := OnConsoleNotify;
  {}
  ConfigModule := TConfigModule.Create( nil );
  ConfigModule.TextSubj.AddObServer( FTextObsrv );
  {}
  HCmdBeginList := TSyncList.Create;
  HCmdWorkingList := TSyncList.Create;
  HCmdFinishList := TSyncList.Create;
  {}
  LCmdLogList := TSyncList.Create;
  LCmdAckList := TSyncList.Create;
  {}
  InitControlLanguage;
  InitControlState;
  {}
  Panel1.Align := alBottom;
  Panel1.Align := alTop;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormShow(Sender: TObject);
begin
  LoadConfigTimer.Enabled := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  if IsThreadRun then CanClose := ConfirmMsg(
    LanguageManager.Get( 'SDlgConfimQuit' ) );
  if CanClose then
  begin
    dxStop.Click;
    SaveToolBarStateToFile;
    SaveWindowPositionToFile;
    SaveDocklayoutToFile;
    ConfigModule.SaveActiveConfigFile( FAppIni );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormDestroy(Sender: TObject);
begin
  FTextSubj.Free;
  FTextObsrv.Free;
  FSoTextObsrv.Free;
  FSoObsrv.Free;
  FSendTextObsrv.Free;
  FSendCmdObsrv.Free;
  FConsoleObsrv.Free;
  {}
  ConfigModule.Free;
  FAppIni.Free;
  FDocklayoutStream.Free;
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.LoadLanguage: Boolean;
begin
  Result := False;
  try
    LanguageManager.LoadFromFile;
  except
    on E: Exception do
    begin
      ErrorMsg( E.Message );
      Exit;
    end;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.CreateAppIniFile: TIniFile;
begin
  Result := FAppIni;
  if not Assigned( Result ) then
    Result := TIniFile.Create( IncludeTrailingPathDelimiter( ExtractFilePath(
     ParamStr( 0 ) ) ) + 'AppInfo.ini' );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.AppointConfigFile;
var
  aIndex: Integer;
  aCompareText: String;
begin
  cmbConfigList.Properties.OnChange := nil;
  try
    cmbConfigList.ItemIndex := -1;
    for aIndex := 0 to cmbConfigList.Properties.Items.Count - 1 do
    begin
      aCompareText := cmbConfigList.Properties.Items[aIndex].Description;
      if AnsiCompareText( aCompareText,
         ConfigModule.ActiveConfigFile ) = 0 then
      begin
        cmbConfigList.ItemIndex := aIndex;
        Break;
      end;
    end;
  finally
    cmbConfigList.Properties.OnChange := cmbConfigListPropertiesChange;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.InitControlState;
begin
  LoadWindowPositionFromFile;
  LoadToolBarStateFromFile;
  dxDockingManager.SaveLayoutToStream( FDocklayoutStream );
  LoadDocklayoutFromFile;
  cmbConfigList.Left := ( lblActiveConfig.Left + lblActiveConfig.Width + 5 );
  ConsoleTree.Styles.StyleSheet := StyleModule.TreeListStyleSheetConsoleBlack;
  ConsoleTree.OptionsView.Headers := False;
  {}
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.InitControlLanguage;
var
  aIndex: Integer;
begin
  { Forms }
  Application.Title := Format( '%s - %s', [LanguageManager.Get( 'SVendor' ),
    LanguageManager.Get( 'SAppTitle' )] );
  Self.Caption := Format( '%s - %s', [LanguageManager.Get( 'SVendor' ),
    LanguageManager.Get( 'SAppTitle' )] );
  { Menu }
  dxMenuItem1.Caption := LanguageManager.Get( dxMenuItem1.Name );
  dxMenuItem11.Caption := LanguageManager.Get( dxMenuItem11.Name );
  dxMenuItem2.Caption := LanguageManager.Get( dxMenuItem2.Name );
  dxMenuItem21.Caption := LanguageManager.Get( dxMenuItem21.Name );
  dxMenuItem22.Caption := LanguageManager.Get( dxMenuItem22.Name );
  dxMenuItem23.Caption := LanguageManager.Get( dxMenuItem23.Name );
  dxMenuItem24.Caption := LanguageManager.Get( dxMenuItem24.Name );
  dxMenuItem25.Caption := LanguageManager.Get( dxMenuItem25.Name );
  dxMenuItem26.Caption := LanguageManager.Get( dxMenuItem26.Name );
  dxMenuItem27.Caption := LanguageManager.Get( dxMenuItem27.Name );
  dxMenuItem3.Caption := LanguageManager.Get( dxMenuItem3.Name );
  { Status }
  dxStatusItem1.Caption := Format( LanguageManager.Get( dxStatusItem1.Name ),
    [LanguageManager.Get( 'SNull' ), LanguageManager.Get( 'SNull' )] );
  dxStatusItem2.Caption := LanguageManager.Get( dxStatusItem2.Name );
  dxStatusItem3.Caption := Format( LanguageManager.Get( dxStatusItem3.Name ),
    [LanguageManager.Get( 'SDisconnct' )] );
  dxStatusItem3.ImageIndex := StyleModule.SocketDisconnectImageIndex;  
  { ToolBar }
  for aIndex := 0 to dxBarManager.Bars.Count - 1 do
    dxBarManager.Bars[aIndex].Caption := LanguageManager.Get(
      dxBarManager.Bars[aIndex].Name );
  { DockPanel }
  dxDockPanel1.Caption := LanguageManager.Get( dxDockPanel1.Name );
  dxDockPanel2.Caption := LanguageManager.Get( dxDockPanel2.Name );
  dxDockPanel3.Caption := LanguageManager.Get( dxDockPanel3.Name );
  dxDockPanel5.Caption := LanguageManager.Get( dxDockPanel5.Name );
  dxDockPanel6.Caption := LanguageManager.Get( dxDockPanel6.Name );
  dxDockPanel7.Caption := LanguageManager.Get( dxDockPanel7.Name );
  { SoTree }
  for aIndex := 0 to SoTree.ColumnCount - 1 do
    SoTree.Columns[aIndex].Caption.Text := LanguageManager.Get(
      SoTree.Columns[aIndex].Name );
  { SendTree }
  for aIndex := 0 to SendTree.ColumnCount - 1 do
    SendTree.Columns[aIndex].Caption.Text := LanguageManager.Get(
      SendTree.Columns[aIndex].Name );
  { ConsoleTree }
  for aIndex := 0 to ConsoleTree.ColumnCount - 1 do
    ConsoleTree.Columns[aIndex].Caption.Text := LanguageManager.Get(
      ConsoleTree.Columns[aIndex].Name );
  { Others }
  lblActiveConfig.Caption := LanguageManager.Get( lblActiveConfig.Name );
  { ToolBar Button }
  dxPlay.Caption := LanguageManager.Get( dxPlay.Name );
  dxPlay.Hint := LanguageManager.Get( dxPlay.Name + 'Hint' );
  dxStop.Caption := LanguageManager.Get( dxStop.Name );
  dxStop.Hint := LanguageManager.Get( dxStop.Name + 'Hint' );
  dxDbTree.Caption := LanguageManager.Get( dxDbTree.Name );
  dxDbTree.Hint := LanguageManager.Get( dxDbTree.Name + 'Hint' );
  dxSendTree.Caption := LanguageManager.Get( dxSendTree.Name );
  dxSendTree.Hint := LanguageManager.Get( dxSendTree.Name + 'Hint' );
  dxConsoleTree.Caption := LanguageManager.Get( dxConsoleTree.Name );
  dxConsoleTree.Hint := LanguageManager.Get( dxConsoleTree.Name + 'Hint' );
  dxAutoScroll.Caption := LanguageManager.Get( dxAutoScroll.Name );
  dxAutoScroll.Hint := LanguageManager.Get( dxAutoScroll.Name + 'Hint' );
  dxCmdExpand.Caption := LanguageManager.Get( dxCmdExpand.Name );
  dxCmdExpand.Hint := LanguageManager.Get( dxCmdExpand.Name + 'Hint' );
  dxCmdCollapse.Caption := LanguageManager.Get( dxCmdCollapse.Name );
  dxCmdCollapse.Hint := LanguageManager.Get( dxCmdCollapse.Name + 'Hint' );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.WMSocket(var Msg: TWMSocket);
var
  aErr: String;
begin
  { 通知 SoDbThread, Socket 連線狀態 }
  if Assigned( FSoDbThread ) then
  begin
    PostMessage( FSoDbThread.MsgHandle, Msg.Msg, TMessage( Msg ).WParam,
      TMessage( Msg ).LParam );
  end;
  { 在狀態列上顯示 DVN 的 Socket 連線狀態 }
  case TSocketMsgReason( TWMSocket( Msg ).SocketMsgReason ) of
    smrStart, smrAck, smrNack:
      begin
        SocketWarningTimer.Enabled := False;
        dxStatusItem3.Caption := Format( LanguageManager.Get( dxStatusItem3.Name ),
          [LanguageManager.Get( 'SConnect' )] );
        dxStatusItem3.ImageIndex := StyleModule.SocketConnectImageIndex;
      end;
    smrSocketError:
     begin
       { 若是 Socket 錯誤, 啟動 Timer 來做警告效果 }
       SocketWarningTimer.Enabled := True;
       aErr := TWMSocket( Msg ).ErrorDescription;
       dxStatusItem3.Caption := Format( LanguageManager.Get( dxStatusItem3.Name ),
         [Format( LanguageManager.Get( 'SConnectError' ), [aErr] ) ] );
     end;
    smrIdleTimeReach:
     begin
       SocketWarningTimer.Enabled := True;
       dxStatusItem3.Caption := Format( LanguageManager.Get( dxStatusItem3.Name ),
        [LanguageManager.Get( 'SPrepareReconnect' )] );
     end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.LoadConfigTimerTimer(Sender: TObject);
var
  aExcResult: Boolean;
begin
  Delay( 180 );
  LoadConfigTimer.Enabled := False;
  PlayGroup.Enabled := False;
  cmbConfigList.Enabled := False;
  dxStatusItem2.ImageIndex := 18;
  {}
  Screen.Cursor := crHourGlass;
  try
    if ( FAppFirstRun ) then
    begin
      FAppFirstRun := False;
      FTextSubj.Info( LanguageManager.Get( 'SConfigSearchBegin' ) );
      Delay( 500 );
      ConfigModule.RefreshConfigFile;
      Delay( 500 );
      BuildConfigFileList;
      FTextSubj.OKFmt( LanguageManager.Get( 'SConfigSearchEnd' ),
        [ConfigModule.ConfigFileCount] );
      Delay( 500 );
      ConfigModule.RestoreActiveConfigFile( FAppIni );
      FTextSubj.InfoFmt( LanguageManager.Get( 'SConfigActiveFile' ),
        [Nvl( ConfigModule.ActiveConfigFile, LanguageManager.Get( 'SNull' ) )] );
      Delay( 500 );
      AppointConfigFile;
    end;
    Delay( 500 );
    aExcResult := ConfigModule.ChangeConfigFile;
    Delay( 500 );
    if ( aExcResult ) then
    begin
      FTextSubj.Info( LanguageManager.Get( 'SConfigBuildSoTreeBegin' ) );
      Delay( 500 );
      BuildSoDbTree;
      Delay( 500 );
      FTextSubj.OK( LanguageManager.Get( 'SConfigBuildSoTreeEnd' ) );
    end;
  finally
    Screen.Cursor := crDefault;
  end;
  cmbConfigList.Enabled := True;
  if ( not ConfigModule.CommEnv.CAEnableSend ) then
  begin
    FTextSubj.Warning( LanguageManager.Get( 'SConfigCommCAEnableSendWarning' ) );
    aExcResult := False;
  end;
  {}
  if ( ConfigModule.SoList.Count <= 0 ) then
  begin
    FTextSubj.Warning( LanguageManager.Get( 'SConfigBuildSoTreeWarning1' ) );
    aExcResult := False;
  end else
  if ( ConfigModule.SoList.SelectCount <= 0 ) then
  begin
    FTextSubj.Warning( LanguageManager.Get( 'SConfigBuildSoTreeWarning2' ) );
    aExcResult := False;
  end;
  {}
  if ( aExcResult ) then
  begin
    FTextSubj.OK( LanguageManager.Get( 'SReadyToRun' ) );
    PlayGroup.Enabled := True;
    dxStatusItem1.Caption := Format( LanguageManager.Get( dxStatusItem1.Name ),
      [ConfigModule.CasEnv.Ip, ConfigModule.CasEnv.SendPort] );
    if ( ConfigModule.CommEnv.CAEnableSend ) then dxStatusItem2.ImageIndex := 17;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.BuildConfigFileList;
var
  aIndex, aPos: Integer;
  aItem: TcxImageComboBoxItem;
begin
  cmbConfigList.Properties.OnChange := nil;
  try
    cmbConfigList.Properties.Items.Clear;
    for aIndex := 0 to ConfigModule.ConfigFileCount - 1 do
    begin
      aPos := LastDelimiter( '.', ConfigModule.ConfigFiles[aIndex] );
      aItem := cmbConfigList.Properties.Items.Add;
      aItem.ImageIndex := 7;
      aItem.Value := ConfigModule.ConfigFiles[aIndex];
      aItem.Description := Copy( ConfigModule.ConfigFiles[aIndex], 1, aPos - 1 );
    end;
  finally
    cmbConfigList.Properties.OnChange := cmbConfigListPropertiesChange;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.LoadWindowPositionFromFile;
begin
  Self.WindowState := TWindowState( FAppIni.ReadInteger(
    'WindowPositon', 'State', Ord( wsNormal ) ) );
  if ( Self.WindowState = wsNormal ) then
  begin
    Self.Width := FAppIni.ReadInteger( 'WindowPositon', 'Width', Self.Width );
    Self.Height := FAppIni.ReadInteger( 'WindowPositon', 'Height', Self.Height );
    Self.Left := FAppIni.ReadInteger( 'WindowPositon', 'Left', Self.Left );
    Self.Top := FAppIni.ReadInteger( 'WindowPositon', 'Top', Self.Top );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.SaveWindowPositionToFile;
begin
  FAppIni.WriteInteger( 'WindowPositon', 'State', Ord( Self.WindowState ) );
  if ( Self.WindowState = wsNormal ) then
  begin
    FAppIni.WriteInteger( 'WindowPositon', 'Width', Self.Width );
    FAppIni.WriteInteger( 'WindowPositon', 'Height', Self.Height );
    FAppIni.WriteInteger( 'WindowPositon', 'Left', Self.Left );
    FAppIni.WriteInteger( 'WindowPositon', 'Top', Self.Top );
    FAppIni.UpdateFile;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.LoadToolBarStateFromFile;
begin
  dxDbTree.Down := FAppIni.ReadBool( 'ToolBarState', dxDbTree.Name, False );
  dxSendTree.Down := FAppIni.ReadBool( 'ToolBarState', dxSendTree.Name, False );
  dxConsoleTree.Down := FAppIni.ReadBool( 'ToolBarState', dxConsoleTree.Name, False );
  dxAutoScroll.Down := FAppIni.ReadBool( 'ToolBarState', dxAutoScroll.Name, False );
  FAutoScroll := dxAutoScroll.Down;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.SaveToolBarStateToFile;
begin
  FAppIni.WriteBool( 'ToolBarState', dxDbTree.Name, dxDbTree.Down );
  FAppIni.WriteBool( 'ToolBarState', dxSendTree.Name, dxSendTree.Down );
  FAppIni.WriteBool( 'ToolBarState', dxConsoleTree.Name, dxConsoleTree.Down );
  FAppIni.WriteBool( 'ToolBarState', dxAutoScroll.Name, dxAutoScroll.Down );
  FAppIni.UpdateFile;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.LoadDocklayoutFromFile;
begin
  dxDockingManager.LoadLayoutFromIniFile( IncludeTrailingPathDelimiter(
    ExtractFilePath( ParamStr( 0 ) ) ) + 'DockLayout.ini' );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.SaveDocklayoutToFile;
begin
  dxDockingManager.SaveLayoutToIniFile( IncludeTrailingPathDelimiter(
    ExtractFilePath( ParamStr( 0 ) ) ) + 'DockLayout.ini' );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.cmbConfigListPropertiesChange(Sender: TObject);
begin
  ConfigModule.ActiveConfigFile := cmbConfigList.Text;
  LoadConfigTimer.Interval := 500;
  LoadConfigTimer.Enabled := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.OnTextNotify(aSubject: TSubject);
var
  aItem: TListItem;
begin
  aItem := MainMsgList.Items.Add;
  aItem.ImageIndex := StyleModule.GetMsgImageIndex( TTextSubject( aSubject ).Notify.Flag );
  aItem.Caption := Format( ' %s >   %s', [FormatDateTime( 'yyyy-mm-dd hh:mm', Now ),
      TTextSubject( aSubject ).Notify.Text] );
  aItem.MakeVisible( True );
  Application.ProcessMessages;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.OnSoTextNotify(aSubject: TSubject);
var
  aItem: TListItem;
begin
  aItem := SoMsgList.Items.Add;
  aItem.ImageIndex := StyleModule.GetMsgImageIndex( TTextSubject( aSubject ).Notify.Flag );
  aItem.Caption := Format( ' %s >   %s', [FormatDateTime( 'yyyy-mm-dd hh:mm', Now ),
      TTextSubject( aSubject ).Notify.Text] );
  aItem.MakeVisible( True );
  Application.ProcessMessages;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.OnSendTextNotify(aSubject: TSubject);
var
  aItem: TListItem;
begin
  aItem := SendMsgList.Items.Add;
  aItem.ImageIndex := StyleModule.GetMsgImageIndex( TTextSubject( aSubject ).Notify.Flag );
  aItem.Caption := Format( ' %s >   %s', [FormatDateTime( 'yyyy-mm-dd hh:mm', Now ),
      TTextSubject( aSubject ).Notify.Text] );
  aItem.MakeVisible( True );
  Application.ProcessMessages;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.OnSoNotify(aSubject: TSubject);
var
  aSubj: TSoSubject;
  aRecords: Integer;
begin
  aSubj := TSoSubject( aSubject );
  if not Assigned( aSubj ) then Exit;
  if not Assigned( aSubj.So ) then Exit;
  if not Assigned( aSubj.So.DisplayNode ) then Exit;
  SoTree.BeginUpdate;
  try
    TcxTreeListNode( aSubj.So.DisplayNode ).ImageIndex :=
      StyleModule.GetSoImageIndex( aSubj.So );
    TcxTreeListNode( aSubj.So.DisplayNode ).SelectedIndex :=
      TcxTreeListNode( aSubj.So.DisplayNode ).ImageIndex;
    TcxTreeListNode( aSubj.So.DisplayNode ).StateIndex :=
      StyleModule.GetSoStateIndex( aSubj.So );
    aRecords := 0;
    if ( aSubj.So.DbState <> dbClose ) then
      aRecords := aSubj.So.RecordCount;
    TcxTreeListNode( aSubj.So.DisplayNode ).Values[SoTreeCol3.ItemIndex] :=
      aRecords;
  finally
    SoTree.EndUpdate;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.OnSendCmdNotify(aSubject: TSubject);
var
  aSubj: TSendDvnSubject;
  aCmd: TSendDvn;
  aJustAdded: Boolean;
  aNode: TcxTreeListNode;

  { ------------------------------------------------------ }

  function GetHighCmdDes(aCmdId: String): String;
  begin
    Result := EmptyStr;
    if ConfigModule.HighCmdEnv.Locate( 'HIGHLEVELCMD', aCmdId, [] ) then
      Result := ConfigModule.HighCmdEnv.FieldByName( 'DESCRIPTION' ).AsString;
  end;

  { ------------------------------------------------------ }

  function DoAddHigh: TcxTreeListNode;
  begin
    Result := SendTree.Add;
    Result.ImageIndex := StyleModule.GetCmdImageIndex( aCmd );
    Result.SelectedIndex := Result.ImageIndex;
    Result.AssignValues( [
      aCmd.CompName,                               { COMPNAME }
      aCmd.HighSeqNo ,                             { SEQNO }
      aCmd.StbNo,                                  { STBNO }
      aCmd.IccNo,                                  { ICCNO }
      Format( '%s - ( %s )', [aCmd.HighCmd,
        GetHighCmdDes( aCmd.HighCmd )] ),          { 指令 }
      aCmd.HighCmdStatus,                          { SENDSTATUS }
      aCmd.HighCmdStatus,                          { RECVSTATUS }
      aCmd.Operator,                               { OPERATOR }
      aCmd.ErrCode,                                { ERRORCODE }
      aCmd.ErrMsg,                                 { ERRORMSG }
      FormatDateTime( 'yyyy-mm-dd hh:mm:ss', aCmd.UpdTime ), { UPDTIME }
      EmptyStr] );
    aCmd.DisplayNode := Result;
    aCmd.ParentNode := nil;
  end;

  { ------------------------------------------------ }

  function DoAddLow: TcxTreeListNode;
  begin
    Result := SendTree.AddChild( TcxTreeListNode( aCmd.ParentNode ) );
    Result.ImageIndex := StyleModule.GetCmdImageIndex( aCmd );
    Result.SelectedIndex := Result.ImageIndex;
    Result.AssignValues( [
      EmptyStr,                                     { COMPNAME }
      Format( '        %s', [aCmd.FrameNo] ),       { SEQNO }
      EmptyStr,                                     { STBNO }
      EmptyStr,                                     { ICCNO }
      Format( '%s', [aCmd.LowCmd] ),                { 指令 }
      aCmd.LowCmdStatus,                            { SENDSTATUS }
      aCmd.LowCmdStatus,                            { RECVSTATUS }
      EmptyStr,                                     { OPERATOR }
      aCmd.ErrCode,                                 { ERRORCODE }
      aCmd.ErrMsg,                                   { ERRORMSG }
      FormatDateTime( 'yyyy-mm-dd hh:mm:ss', aCmd.SendTime ),   { SendTime }
      EmptyStr] ); { RecvTime }
    if ( aCmd.LowCmdStatus = 'P' ) or
       ( aCmd.LowCmdStatus = 'C' ) then
    begin
      if Result.Parent.Values[SendTreeCol6.ItemIndex] = 'W' then
        Result.Parent.Values[SendTreeCol6.ItemIndex] := aCmd.LowCmdStatus;
    end;
    aCmd.DisplayNode := Result;
  end;

  { ------------------------------------------------ }

  procedure MakeSureNodeVisible(aNode: TcxTreeListNode);
  begin
    if Assigned( aNode.Parent ) then
      if not aNode.Parent.Expanded then
        aNode.Parent.Expand( True );
    aNode.MakeVisible;
  end;

  { ------------------------------------------------------ }

begin
  aSubj := TSendDvnSubject( aSubject );
  aCmd := aSubj.SendDvn;
  aJustAdded := ( not Assigned( aCmd.DisplayNode ) );
  SendTree.BeginUpdate;
  try
    aNode := nil;
    if ( aJustAdded ) then
    begin
      if ( aCmd.CmdType = ctHigh ) then
        aNode := DoAddHigh
      else begin
        if Assigned( aCmd.ParentNode ) then 
          aNode := DoAddLow;
      end;  
    end else
    begin
      aNode := TcxTreeListNode( aCmd.DisplayNode );
      if ( aCmd.CmdType = ctHigh ) then
        aNode.Values[SendTreeCol7.ItemIndex] := aCmd.HighCmdStatus
      else
        aNode.Values[SendTreeCol7.ItemIndex] := aCmd.LowCmdStatus;
      if aCmd.ErrCode <> EmptyStr then
        aNode.Values[SendTreeCol9.ItemIndex] := aCmd.ErrCode;
      if aCmd.ErrMsg <> EmptyStr then
        aNode.Values[SendTreeCol10.ItemIndex] := aCmd.ErrMsg;
      {}
      if ( aCmd.CmdType = ctLow ) then
        aNode.Values[SendTreeCol12.ItemIndex] :=
          FormatDateTime( 'yyyy/mm/dd hh:nn:ss', aCmd.RecvTime )
      else begin
        if ( aCmd.LowCmdCount <= 0 ) then
        aNode.Values[SendTreeCol12.ItemIndex] :=
          FormatDateTime( 'yyyy/mm/dd hh:nn:ss', aCmd.RecvTime )
      end;
    end;
  finally
    SendTree.EndUpdate;
  end;
  if FAutoScroll and aJustAdded and ( aCmd.CmdType = ctLow ) and Assigned( aNode ) then
    MakeSureNodeVisible( aNode );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.OnConsoleNotify(aSubject: TSubject);
var
  aSubj: TSendDvnSubject;
  aCmd: TSendDvn;
  aNode: TcxTreeListNode;

  { ------------------------------------------------ }

  function AddNode: TcxTreeListNode;
  begin
    Result := ConsoleTree.Add;
    Result.ImageIndex := StyleModule.GetCmdImageIndex( aCmd );
    Result.SelectedIndex := Result.ImageIndex;
    Result.AssignValues( [
      aCmd.FrameNo,
      FormatDateTime( 'yyyy/mm/dd hh:nn:ss', aCmd.SendTime ),
      EmptyStr,
      LanguageManager.Get( 'SUnknow' ),
      aCmd.SendCmdText,
      EmptyStr] );
  end;

  { ------------------------------------------------ }

var
  aText: String;
begin
  aSubj := TSendDvnSubject( aSubject );
  aCmd := aSubj.SendDvn;
  if ( aCmd.CmdType = ctHigh ) then Exit;
  aNode := aCmd.ConsoleNode;
  if not Assigned( aNode ) then
  begin
    aNode := AddNode;
    aCmd.ConsoleNode := aNode;
    if ( FAutoScroll ) then aNode.MakeVisible;
  end else
  begin
    aText := LanguageManager.Get( 'SUnknow' );
    if ( aCmd.LowCmdStatus = 'C' ) then
      aText := LanguageManager.Get( 'SAck' )
    else if ( aCmd.LowCmdStatus = 'E' ) then
      aText := LanguageManager.Get( 'SNack' );
    { 回應 }
    aNode.Texts[ConsoleTreeCol4.ItemIndex] := aText;
    { 回應時間 }
    aNode.Texts[ConsoleTreeCol3.ItemIndex] :=
      FormatDateTime( 'yyyy/mm/dd hh:nn:ss', aCmd.RecvTime );
    { 回應格式 }
    aNode.Texts[ConsoleTreeCol6.ItemIndex] := aCmd.RecvCmdText;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.BuildSoDbTree;

  { -------------------------------------------------- }

  function GetParnetNode(const aPosName: String):
    TcxTreeListNode;
  begin
    Result := nil;
    if ( aPosName = EmptyStr ) then Exit;
    Result := SoTree.FindNodeByText( aPosName, SoTreeCol2 );
    if not Assigned( Result ) then
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
  aSo: TSo;
  aNode, aParnetNode: TcxTreeListNode;
begin
  SoTree.Clear;
  for aIndex := 0 to ConfigModule.SoList.Count - 1 do
  begin
    aSo := ConfigModule.SoList[aIndex];
    aParnetNode := GetParnetNode( aSo.PosName );
    aNode := SoTree.AddChild( aParnetNode );
    aNode.AssignValues( [aSO.CompCode, aSo.CompName, aSo.RecordCount] );
    aNode.ImageIndex := StyleModule.GetSoImageIndex( aSo );
    aNode.StateIndex := StyleModule.GetSoStateIndex( aSo );
    aNode.SelectedIndex := aNode.ImageIndex;
    aSo.DisplayNode := aNode;
    aParnetNode.Expanded := True;
    FTextSubj.InfoFmt( LanguageManager.Get( 'SConfigBuildSoTree' ),[aSo.CompName] );
    Delay( 180 );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.CreateSoDbThread: TSoDbThread;
begin
  Result := TSoDbThread.Create( True );
  Result.FreeOnTerminate := False;
  Result.MainFormHandle := Self.Handle;
  Result.CasEnv.Assign( ConfigModule.CasEnv );
  Result.CommEnv.Assign( ConfigModule.CommEnv );
  Result.HighCmdEnv.Data := ConfigModule.HighCmdEnv.Data;
  Result.LowCmdEnv.Data := ConfigModule.LowCmdEnv.Data;
  Result.CmdErrCnv.Data := ConfigModule.CmdErrorEnv.Data;
  Result.SoList := ConfigModule.SoList;
  Result.TextSubj.AddObServer( FSoTextObsrv );
  Result.SoSubj.AddObServer( FSoObsrv );
  Result.CmdSubj.AddObServer( FSendCmdObsrv );
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.CreateSendThread: TSendThread;
begin
  Result := TSendThread.Create( True );
  Result.FreeOnTerminate := False;
  Result.MainFormHandle := Self.Handle;
  Result.CasEnv.Assign( ConfigModule.CasEnv );
  Result.CommEnv.Assign( ConfigModule.CommEnv );
  Result.HighCmdEnv.Data := ConfigModule.HighCmdEnv.Data;
  Result.LowCmdEnv.Data := ConfigModule.LowCmdEnv.Data;
  Result.CmdErrCnv.Data := ConfigModule.CmdErrorEnv.Data;
  Result.TextSubj.AddObServer( FSendTextObsrv );
  Result.CmdSubj.AddObServer( FSendCmdObsrv );
  Result.CmdSubj.AddObServer( FConsoleObsrv );
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.CreateRecvThread: TRecvThread;
begin
  Result := TRecvThread.Create( True );
  Result.FreeOnTerminate := False;
  Result.MainFormHandle := Self.Handle;
  Result.CasEnv.Assign( ConfigModule.CasEnv );
  Result.CommEnv.Assign( ConfigModule.CommEnv );
  Result.HighCmdEnv.Data := ConfigModule.HighCmdEnv.Data;
  Result.LowCmdEnv.Data := ConfigModule.LowCmdEnv.Data;
  Result.CmdErrCnv.Data := ConfigModule.CmdErrorEnv.Data;
  Result.TextSubj.AddObServer( FSendTextObsrv );
  Result.CmdSubj.AddObServer( FSendCmdObsrv );
  Result.CmdSubj.AddObServer( FConsoleObsrv );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.ReleaseSoDbThread;
begin
  if Assigned( FSoDbThread ) then
  begin
    FSoDbThread.Terminate;
    FSoDbThread.WaitFor;
    FSoDbThread.Free;
    FSoDbThread := nil;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.ReleaseSendThread;
begin
  if Assigned( FSendThread ) then
  begin
    FSendThread.Terminate;
    FSendThread.WaitFor;
    FSendThread.Free;
    FSendThread := nil;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.ReleaseRecvThread;
begin
  if Assigned( FRecvThread ) then
  begin
    FRecvThread.Terminate;
    FRecvThread.WaitFor;
    FRecvThread.Free;
    FRecvThread := nil;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.DisposeSyncList;
begin
  while ( HCmdBeginList.Count > 0 ) do
    HCmdBeginList.FreeObject( 0 );
  while ( HCmdWorkingList.Count > 0 ) do
    HCmdWorkingList.FreeObject( 0 );
  while ( HCmdFinishList.Count > 0 ) do
    HCmdFinishList.FreeObject( 0 );
  {}
  while ( LCmdLogList.Count > 0 ) do
    LCmdLogList.FreeObject( 0 );
  while ( LCmdAckList.Count > 0 ) do
    LCmdAckList.FreeObject( 0 );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.BeginExecute;
begin
  if ( ConfigModule.CommEnv.CAEnableSend ) then
  begin
    FSoDbThread := CreateSoDbThread;
    FSendThread := CreateSendThread;
    FRecvThread := CreateRecvThread;
    FTextSubj.Info( LanguageManager.Get( 'SRun' ) );
    Delay( 500 );
    FSoDbThread.Resume;
    Delay( 500 );
    FRecvThread.Resume;
    Delay( 500 );
    FSendThread.Resume;
    IsThreadRun := True;
  end;
  if ( ConfigModule.CommEnv.CAEnableSend ) then
  begin
    PlayGroup.Enabled := False;
    StopGroup.Enabled := True;
    cmbConfigList.Properties.ReadOnly := True;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.EndExecute;
begin
  IsThreadRun := False;
  ReleaseSendThread;
  Application.ProcessMessages;
  ReleaseRecvThread;
  Application.ProcessMessages;
  ReleaseSoDbThread;
  Delay( 500 );
  FTextSubj.Info( LanguageManager.Get( 'Stop' ) );
  DisposeSyncList;
  if ( ConfigModule.CommEnv.CAEnableSend ) then
  begin
    PlayGroup.Enabled := True;
    StopGroup.Enabled := False;
    cmbConfigList.Properties.ReadOnly := False;
  end;
  {}
  SocketWarningTimer.Enabled := False;
  dxStatusItem3.Caption := Format( LanguageManager.Get( dxStatusItem3.Name ),
    [LanguageManager.Get( 'SDisconnct' )] );
  dxStatusItem3.ImageIndex := StyleModule.SocketDisconnectImageIndex;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxPlayClick(Sender: TObject);
begin
  Screen.Cursor := crHourGlass;
  try
    BeginExecute;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxStopClick(Sender: TObject);
begin
  Screen.Cursor := crHourGlass;
  try
    EndExecute;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.SendTreeCustomDrawCell(Sender: TObject;
  ACanvas: TcxCanvas; AViewInfo: TcxTreeListEditCellViewInfo;
  var ADone: Boolean);
var
  aIconLeft, aIconTop, aFlag, aOffsetX, aImgIdx: Integer;
  aTextRect: TRect;
begin
  if AViewInfo.Node.Level = 0 then
  begin
    ACanvas.FillRect( AViewInfo.VisibleRect );
    if AViewInfo.Column = SendTreeCol8 then
    begin
      aIconLeft := AViewInfo.VisibleRect.Left + 3;
      aIconTop := AViewInfo.VisibleRect.Top;
      ACanvas.DrawImage( StyleModule.CmdTreeImageList, aIconLeft,
        aIconTop, StyleModule.OperatorImageIndex );
      aFlag := ( cxSingleLine	or cxShowEndEllipsis or cxAlignLeft	or cxAlignVCenter );
      aTextRect := AViewInfo.VisibleRect;
      aOffsetX := StyleModule.CmdTreeImageList.Width + 8;
      InflateRect( aTextRect, ( 0 - aOffsetX ), 0 );
      ACanvas.DrawText( AViewInfo.DisplayValue, aTextRect, aFlag );
      ADone := True;
    end;
  end else
  begin
    if ( AViewInfo.Column = SendTreeCol5 ) or
       ( AViewInfo.Column = SendTreeCol11 ) or
       ( AViewInfo.Column = SendTreeCol12 ) then
    begin
      if ( AViewInfo.DisplayValue = EmptyStr ) then Exit;
      ACanvas.FillRect( AViewInfo.ContentRect );
      aIconLeft := AViewInfo.ContentRect.Left + 10;
      aIconTop := AViewInfo.ContentRect.Top + 2;
      if ( AViewInfo.Column = SendTreeCol5 ) then
        aImgIdx := StyleModule.LowCmdImageIndex
      else if ( AViewInfo.Column = SendTreeCol11 ) then
        aImgIdx := StyleModule.ClockSendImageIndex
      else
        aImgIdx := StyleModule.ClockRecvImageIndex;
      ACanvas.DrawImage( StyleModule.CmdTreeImageList, aIconLeft,
        aIconTop, aImgIdx );
      aFlag := ( cxSingleLine	or cxShowEndEllipsis or cxAlignLeft	or cxAlignVCenter );
      aTextRect := AViewInfo.VisibleRect;
      aOffsetX := StyleModule.CmdTreeImageList.Width + 15;
      InflateRect( aTextRect, ( 0 - aOffsetX ), 0 );
      ACanvas.DrawText( AViewInfo.DisplayValue, aTextRect, aFlag );
      ADone := True;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxCmdExpandClick(Sender: TObject);
begin
  SendTree.FullExpand;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxCmdCollapseClick(Sender: TObject);
begin
  SendTree.FullCollapse;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxDbTreeClick(Sender: TObject);
begin
  FSoObsrv.Pause := not dxDbTree.Down;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxSendTreeClick(Sender: TObject);
begin
  FSendCmdObsrv.Pause := not dxSendTree.Down;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxConsoleTreeClick(Sender: TObject);
begin
  FConsoleObsrv.Pause := not dxConsoleTree.Down;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxAutoScrollClick(Sender: TObject);
begin
  FAutoScroll := dxAutoScroll.Down;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.ConsoleTreeKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if ( ssCtrl in Shift  ) then
  begin
    if ( Chr( Key ) = 'C' ) then
      ConsoleTree.CopySelectedToClipboard
    else if ( Chr( Key ) = 'A' ) then
      ConsoleTree.SelectAll;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.SocketWarningTimerTimer(Sender: TObject);
begin
  if ( dxStatusItem3.ImageIndex = StyleModule.SocketNullImageIndex ) then
    dxStatusItem3.ImageIndex := StyleModule.SocketWarningImageIndex
  else
    dxStatusItem3.ImageIndex := StyleModule.SocketNullImageIndex;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxMenuItem21Click(Sender: TObject);
begin
  FDocklayoutStream.Position := 0;
  dxDockingManager.LoadLayoutFromStream( FDocklayoutStream );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxMenuItem22Click(Sender: TObject);
var
  aDockPanel: TdxDockPanel;
begin
  aDockPanel := nil;
  if Sender = dxMenuItem22 then
  begin
    aDockPanel := dxDockPanel1;
  end else
  if Sender = dxMenuItem23 then
  begin
    aDockPanel := dxDockPanel2;
  end else
  if Sender = dxMenuItem24 then
  begin
    aDockPanel := dxDockPanel3;
  end else
  if Sender = dxMenuItem25 then
  begin
    aDockPanel := dxDockPanel5;
  end else
  if Sender = dxMenuItem26 then
  begin
    aDockPanel := dxDockPanel6;
  end else
  if Sender = dxMenuItem27 then
  begin
    aDockPanel := dxDockPanel7;
  end;
  if Assigned( aDockPanel ) then
  begin
    if not aDockPanel.Visible then aDockPanel.Visible := True;
    aDockPanel.Controller.ActiveDockControl := aDockPanel;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxMenuItem11Click(Sender: TObject);
begin
  if ( ConfigModule.ShowConfigForm = mrOk ) then
  begin
    FTextSubj.Info( LanguageManager.Get( 'SConfigBuildSoTreeBegin' ) );
    Delay( 200 );
    BuildSoDbTree;
    Delay( 200 );
    dxStatusItem1.Caption := Format( LanguageManager.Get( dxStatusItem1.Name ),
      [ConfigModule.CasEnv.Ip, ConfigModule.CasEnv.SendPort] );
    FTextSubj.OK( LanguageManager.Get( 'SConfigBuildSoTreeEnd' ) );
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
