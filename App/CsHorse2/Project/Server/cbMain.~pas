unit cbMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  ImgList, ExtCtrls, StdCtrls, ComCtrls, CommCtrl, IniFiles, jpeg,
  {$IFDEF APPDEBUG} CodeSiteLogging, {$ENDIF} DB, 
{ App: }
  cbStyleModule, cbMainUIModule, cbSrvClass, cbDesignPattern, cbLanguage,
  cbROServerModule, cbSo,
{ Developer Express: }
  cxClasses, dxBar, dxBarExtItems, dxDockControl, dxDockPanel, cxControls,
  cxContainer, cxListView, cxPC, cxGraphics, cxCustomData, cxStyles, cxTL,
  cxInplaceContainer, cxImageComboBox, cxEdit, cxTextEdit, cxDropDownEdit,
  dxSkinsCore, dxSkinsDefaultPainters, dxSkinsdxDockControlPainter,
  dxSkinsdxBarPainter, dxGDIPlusClasses, dxSkinscxPCPainter, cxFilter, cxData,
  cxDataStorage, cxDBData, cxGridLevel, cxGridCustomView,  cxGridCustomTableView,
  cxGridTableView, cxGridDBTableView, cxGrid, cxGridStrs, dxSkinBlack,
  dxSkinBlue, dxSkinCaramel, dxSkinCoffee, dxSkinGlassOceans, dxSkiniMaginary,
  dxSkinLilian, dxSkinLiquidSky, dxSkinLondonLiquidSky, dxSkinMcSkin,
  dxSkinMoneyTwins, dxSkinOffice2007Black, dxSkinOffice2007Blue,
  dxSkinOffice2007Green, dxSkinOffice2007Pink, dxSkinOffice2007Silver,
  dxSkinSilver, dxSkinStardust, dxSkinValentine, dxSkinXmas2008Blue,
{ Indy }
  IdStack, dxSkinSummer2008;


type
  TfmMain = class(TForm)
    dxBarManager: TdxBarManager;
    dxMenuItem1: TdxBarSubItem;
    dxPlay: TdxBarButton;
    dxStop: TdxBarButton;
    dxAutoScroll: TdxBarButton;
    dxExpand: TdxBarButton;
    dxCollapse: TdxBarButton;
    dxStatusItem1: TdxBarStatic;
    dxTabContainerDockSite2: TdxTabContainerDockSite;
    dxDockingManager: TdxDockingManager;
    dxDockHost: TdxDockSite;
    dxLayoutDockSite1: TdxLayoutDockSite;
    dxLayoutDockSite2: TdxLayoutDockSite;
    dxLayoutDockSite3: TdxLayoutDockSite;
    dxDockPanel1: TdxDockPanel;
    dxDockPanel2: TdxDockPanel;
    dxDockPanel5: TdxDockPanel;
    dxDockPanel7: TdxDockPanel;
    PlayGroup: TdxBarGroup;
    StopGroup: TdxBarGroup;
    LoadConfigTimer: TTimer;
    MainMsgList: TcxListView;
    UserMsgList: TcxListView;
    SoTcpTree: TcxTreeList;
    SoTcpTreeCol1: TcxTreeListColumn;
    SoTcpTreeCol2: TcxTreeListColumn;
    SoTcpTreeCol3: TcxTreeListColumn;
    dxMenuItem11: TdxBarButton;
    dxMenuItem2: TdxBarSubItem;
    dxMenuItem21: TdxBarButton;
    dxMenuItem3: TdxBarButton;
    SocketWarningTimer: TTimer;
    dxStatusItem2: TdxBarStatic;
    dxConfigList: TdxBarImageCombo;
    gvUser: TcxGridDBTableView;
    glUser: TcxGridLevel;
    UserGrid: TcxGrid;
    gvUserCol2: TcxGridDBColumn;
    gvUserCol1: TcxGridDBColumn;
    gvUserCol3: TcxGridDBColumn;
    gvUserCol4: TcxGridDBColumn;
    gvUserCol6: TcxGridDBColumn;
    gvUserCol5: TcxGridDBColumn;
    gvUserCol8: TcxGridDBColumn;
    gvUserCol9: TcxGridDBColumn;
    gvUserCol10: TcxGridDBColumn;
    gvUserCol7: TcxGridDBColumn;
    gvUserCol20: TcxGridDBColumn;
    gvUserCol21: TcxGridDBColumn;
    gvUserCol11: TcxGridDBColumn;
    gvUserCol12: TcxGridDBColumn;
    gvUserCol13: TcxGridDBColumn;
    gvUserCol14: TcxGridDBColumn;
    SoTcpTreeCol4: TcxTreeListColumn;
    dxDockPanel8: TdxDockPanel;
    dxDockPanel4: TdxDockPanel;
    dxTabContainerDockSite1: TdxTabContainerDockSite;
    SoSessionTree: TcxTreeList;
    SoSessionTreeCol1: TcxTreeListColumn;
    SoSessionTreeCol2: TcxTreeListColumn;
    SoSessionTreeCol3: TcxTreeListColumn;
    SoSessionTreeCol4: TcxTreeListColumn;
    SoSessionTreeCol5: TcxTreeListColumn;
    SoSyncMsgList: TcxListView;
    dxGroupByBox: TdxBarButton;
    gvUserCol15: TcxGridDBColumn;
    dxMenuItem12: TdxBarButton;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormDestroy(Sender: TObject);
    procedure LoadConfigOnTimer(Sender: TObject);
    {}
    procedure OnMsgUpdate(ASubject: TSubject);
    procedure OnDbSyncMsgUpdate(ASubject: TSubject);
    procedure OnUserMsgUpdate(ASubject: TSubject);
    procedure OnUserListUpdate(ASubject: TSubject);
    procedure OnSoSessionUpdate(ASubject: TSubject);
    procedure OnSoTcpUpdate(ASubject: TSubject);
    procedure OnDbWarningUpdate(ASubject: TSubject);
    {}
    procedure dxPlayClick(Sender: TObject);
    procedure dxStopClick(Sender: TObject);
    procedure dxExpandClick(Sender: TObject);
    procedure dxCollapseClick(Sender: TObject);
    procedure dxAutoScrollClick(Sender: TObject);
    procedure SocketWarningTimerTimer(Sender: TObject);
    procedure dxMenuItem21Click(Sender: TObject);
    procedure dxMenuItem11Click(Sender: TObject);
    procedure dxConfigListChange(Sender: TObject);
    procedure dxGroupByBoxClick(Sender: TObject);
    procedure gvUserCustomDrawCell(Sender: TcxCustomGridTableView;
      ACanvas: TcxCanvas; AViewInfo: TcxGridTableDataCellViewInfo;
      var ADone: Boolean);
    procedure dxMenuItem12Click(Sender: TObject);
    procedure dxMenuItem3Click(Sender: TObject);
  private
    { Private declarations }
    FAppFirstRun: Boolean;
    FAutoScroll: Boolean;
    FDocklayoutStream: TMemoryStream;
    FMsgSub: TMsgSubject;
    FMsgObsrv: TObServer;
    FDbSessionObsrv: TObServer;
    FDbSyncMsgObsrv: TObServer;
    FUserMsgObsrv: TObServer;
    FUserListObsrv: TObServer;
    FDbWarningObsrv: TObServer;
    FSessionMonitor: TThread;
    FSyncDbThread: TThread;
    FUIThread: TThread;
    function LoadLanguage: Boolean;
    function CreateAppIniFile: TIniFile;
    procedure AppointConfigFile;
    procedure InitLanguage;
    procedure InitControl;
    procedure BuildSoTree;
    procedure BuildSoTcpTree;
    procedure BuildSoSessionTree;
    procedure MakeMsgItem(AItem: TListItem; ASubject: TSubject);
    procedure CreateSoDbSession;
    procedure DestroySoDbSession;
    procedure CreateThreads;
    procedure DestroyThreads;
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

implementation

{$R *.dfm}

uses
  cbUtilis, cbConfigModule, cbAbout,
  cbSyncDbDataThread, cbSessionMonitorThread, cbUpdateUIThread;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormCreate(Sender: TObject);
begin
  Application.ShowMainForm := False;
  if not LoadLanguage then  //載入*.Lang檔案 裡面包含所有要顯示的資訊資料 用Notepad開啟即可
  begin
    Application.Terminate;
    Exit;
  end;
  Application.ShowMainForm := True;
  AppIni := CreateAppIniFile;
  FAppFirstRun := True;
  FDocklayoutStream := TMemoryStream.Create;
  FAutoScroll := False;
  {}
  FMsgObsrv := TObServer.Create;
  FMsgObsrv.OnUpdate := OnMsgUpdate;
  FMsgSub := TMsgSubject.Create;
  FMsgSub.AddObServer( FMsgObsrv );
  {}
  FDbSessionObsrv := TObServer.Create;
  FDbSessionObsrv.OnUpdate := OnSoSessionUpdate;
  {}
  FDbSyncMsgObsrv := TObServer.Create;
  FDbSyncMsgObsrv.OnUpdate := OnDbSyncMsgUpdate;
  {}
  FUserMsgObsrv := TObServer.Create;
  FUserMsgObsrv.OnUpdate := OnUserMsgUpdate;
  FUserListObsrv := TObServer.Create;
  FUserListObsrv.OnUpdate := OnUserListUpdate;
  {}
  FDbWarningObsrv := TObServer.Create;
  FDbWarningObsrv.OnUpdate := OnDbWarningUpdate;
  {}
  ConfigModule := TConfigModule.Create( nil );
  ConfigModule.MsgSub.AddObServer( FMsgObsrv );
  {}
  MainUIModule := TMainUIModule.Create( nil );
  ROServerModule := TROServerModule.Create( nil );
  ROServerModule.LoginMsgSubject.AddObServer( FUserMsgObsrv );
  {}
  InitLanguage;
  InitControl;
  {}
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormShow(Sender: TObject);
begin
  LoadConfigTimer.Enabled := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  if ( AppRunState = arRun ) then CanClose := ConfirmMsg(
    LanguageManager.Get( 'SDlgConfimQuit' ) );
  if CanClose then
  begin
    dxStop.Click;
    SaveToolBarStateToFile;
    SaveWindowPositionToFile;
    SaveDocklayoutToFile;
    ConfigModule.SaveActiveConfigFile( AppIni );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormDestroy(Sender: TObject);
begin
  FMsgSub.Free;
  FMsgObsrv.Free;
  FDbSessionObsrv.Free;
  FDbSyncMsgObsrv.Free;
  FDbWarningObsrv.Free;
  {}
  ConfigModule.Free;
  MainUIModule.Free;
  ROServerModule.Free;
  AppIni.Free;
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
var
  aFileName: String;
begin
  Result := AppIni;
  if not Assigned( Result ) then
  begin
    aFileName := IncludeTrailingPathDelimiter( ExtractFilePath( ParamStr( 0 ) ) ) + 'Server.ini';
    FileNotExistsAndCreate( aFileName, True );
    Result := TIniFile.Create( aFileName );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.AppointConfigFile;
var
  aIndex: Integer;
  aCompareText: String;
  aEvent: TNotifyEvent;
begin
  aEvent := dxConfigList.OnChange;
  dxConfigList.OnChange := nil;
  try
    dxConfigList.ItemIndex := -1;
    for aIndex := 0 to dxConfigList.Items.Count - 1 do
    begin
      aCompareText := dxConfigList.Items[aIndex];
      if AnsiCompareText( aCompareText, ConfigModule.ActiveConfigFile ) = 0 then
      begin
        dxConfigList.ItemIndex := aIndex;
        Break;
      end;
    end;
  finally
    dxConfigList.OnChange := aEvent;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.InitLanguage;
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
  dxMenuItem12.Caption := LanguageManager.Get( dxMenuItem12.Name );
  dxMenuItem2.Caption := LanguageManager.Get( dxMenuItem2.Name );
  dxMenuItem21.Caption := LanguageManager.Get( dxMenuItem21.Name );
  dxMenuItem3.Caption := LanguageManager.Get( dxMenuItem3.Name );
  { Status }
  dxStatusItem1.Caption := LanguageManager.GetFmt( dxStatusItem1.Name,
    [LanguageManager.Get( 'SNull' ), LanguageManager.Get( 'SNull' )] );
  dxStatusItem2.Caption := LanguageManager.Get( 'Stop' );
  { ToolBar }
  for aIndex := 0 to dxBarManager.Bars.Count - 1 do
    dxBarManager.Bars[aIndex].Caption := LanguageManager.Get(
      dxBarManager.Bars[aIndex].Name );
  { DockPanel }
  dxDockPanel1.Caption := LanguageManager.Get( dxDockPanel1.Name );
  dxDockPanel2.Caption := LanguageManager.Get( dxDockPanel2.Name );
  dxDockPanel5.Caption := LanguageManager.Get( dxDockPanel5.Name );
  dxDockPanel7.Caption := LanguageManager.Get( dxDockPanel7.Name );
  { SoTcpTree }
  for aIndex := 0 to SoTcpTree.ColumnCount - 1 do
    SoTcpTree.Columns[aIndex].Caption.Text := LanguageManager.Get(
      SoTcpTree.Columns[aIndex].Name );
  { SoSessionTree }
  for aIndex := 0 to SoSessionTree.ColumnCount - 1 do
    SoSessionTree.Columns[aIndex].Caption.Text := LanguageManager.Get(
      SoSessionTree.Columns[aIndex].Name );
  { ToolBar Button }
  dxPlay.Caption := LanguageManager.Get( dxPlay.Name );
  dxPlay.Hint := dxPlay.Caption;
  dxStop.Caption := LanguageManager.Get( dxStop.Name );
  dxStop.Hint := dxStop.Caption;
  dxAutoScroll.Caption := LanguageManager.Get( dxAutoScroll.Name );
  dxAutoScroll.Hint := dxAutoScroll.Caption;
  dxExpand.Caption := LanguageManager.Get( dxExpand.Name );
  dxExpand.Hint := dxExpand.Caption;
  dxCollapse.Caption := LanguageManager.Get( dxCollapse.Name );
  dxCollapse.Hint := dxCollapse.Caption;
  dxGroupByBox.Caption := LanguageManager.Get( dxGroupByBox.Name );
  dxGroupByBox.Hint := dxGroupByBox.Caption;
  {}
  TcxImageComboBoxProperties( gvUserCol15.Properties ).Items[0].Description :=
    LanguageManager.Get( 'SOffline' );
  TcxImageComboBoxProperties( gvUserCol15.Properties ).Items[1].Description :=
    LanguageManager.Get( 'SOnline' );
  TcxImageComboBoxProperties( gvUserCol15.Properties ).Items[2].Description :=
    LanguageManager.Get( 'SOnlineBusy' );
  { Change Developer Express Resource String }
  cxSetResourceString( @scxGridNoDataInfoText, '無資料可顯示' );
  cxSetResourceString( @scxGridGroupByBoxCaption, '請拖曳欄位標題做為分組依據' );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.InitControl;
begin
  dxDockingManager.Font.Name := 'Tahoma';
  LoadWindowPositionFromFile;
  LoadToolBarStateFromFile;
  dxDockingManager.SaveLayoutToStream( FDocklayoutStream );
  LoadDocklayoutFromFile;
  dxDockPanel1.Visible := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.LoadConfigOnTimer(Sender: TObject);
var
  aExcResult: Boolean;
begin
  Delay( 180 );
  LoadConfigTimer.Enabled := False;
  PlayGroup.Enabled := False;
  dxConfigList.ReadOnly := True;
  {}
  Screen.Cursor := crHourGlass;
  try
    if ( FAppFirstRun ) then
    begin
      FAppFirstRun := False;
      FMsgSub.Info( LanguageManager.Get( 'SConfigSearchBegin' ) );
      Delay( 500 );
      ConfigModule.RefreshConfigFile;
      Delay( 500 );
      BuildConfigFileList;
      FMsgSub.OKFmt( LanguageManager.Get( 'SConfigSearchEnd' ),
        [ConfigModule.ConfigFileCount] );
      Delay( 500 );
      ConfigModule.RestoreActiveConfigFile( AppIni );
      FMsgSub.InfoFmt( LanguageManager.Get( 'SConfigActiveFile' ),
        [Nvl( ConfigModule.ActiveConfigFile, LanguageManager.Get( 'SNull' ) )] );
      Delay( 500 );
      AppointConfigFile;
    end;
    Delay( 500 );
    aExcResult := ConfigModule.ChangeConfigFile;
    Delay( 500 );
    if ( aExcResult ) then BuildSoTree;
  finally
    Screen.Cursor := crDefault;
  end;
  {}
  if ( ConfigModule.SoList.Count <= 0 ) then
  begin
    FMsgSub.Warning( LanguageManager.Get( 'SConfigBuildSoTreeWarning1' ) );
    aExcResult := False;
  end else
  if ( ConfigModule.SoList.SelectCount <= 0 ) then
  begin
    FMsgSub.Warning( LanguageManager.Get( 'SConfigBuildSoTreeWarning2' ) );
    aExcResult := False;
  end;
  {}
  if ( aExcResult ) then
  begin
    FMsgSub.OK( LanguageManager.Get( 'SReadyToRun' ) );
    PlayGroup.Enabled := True;
  end;
  {}
  dxConfigList.ReadOnly := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.BuildConfigFileList;
var
  aIndex, aPos, aItemIdx: Integer;
  aEvent: TNotifyEvent;
begin
  aEvent := dxConfigList.OnChange;
  dxConfigList.OnChange := nil;
  try
    dxConfigList.Items.Clear;
    for aIndex := 0 to ConfigModule.ConfigFileCount - 1 do
    begin
      aPos := LastDelimiter( '.', ConfigModule.ConfigFiles[aIndex] );
      aItemIdx := dxConfigList.Items.Add( Copy(
        ConfigModule.ConfigFiles[aIndex], 1, aPos - 1 ) );
      dxConfigList.ImageIndexes[aItemIdx] := StyleModule.ConfigFileImageIndex;
    end;
  finally
    dxConfigList.OnChange := aEvent;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.LoadWindowPositionFromFile;
begin
  Self.WindowState := TWindowState( AppIni.ReadInteger(
    'WindowPositon', 'State', Ord( wsNormal ) ) );
  if ( Self.WindowState = wsNormal ) then
  begin
    Self.Width := AppIni.ReadInteger( 'WindowPositon', 'Width', Self.Width );
    Self.Height := AppIni.ReadInteger( 'WindowPositon', 'Height', Self.Height );
    Self.Left := AppIni.ReadInteger( 'WindowPositon', 'Left', Self.Left );
    Self.Top := AppIni.ReadInteger( 'WindowPositon', 'Top', Self.Top );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.SaveWindowPositionToFile;
begin
  AppIni.WriteInteger( 'WindowPositon', 'State', Ord( Self.WindowState ) );
  if ( Self.WindowState = wsNormal ) then
  begin
    AppIni.WriteInteger( 'WindowPositon', 'Width', Self.Width );
    AppIni.WriteInteger( 'WindowPositon', 'Height', Self.Height );
    AppIni.WriteInteger( 'WindowPositon', 'Left', Self.Left );
    AppIni.WriteInteger( 'WindowPositon', 'Top', Self.Top );
    AppIni.UpdateFile;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.LoadToolBarStateFromFile;
begin
  dxAutoScroll.Down := AppIni.ReadBool( 'ToolBarState', dxAutoScroll.Name, False );
  dxGroupByBox.Down := AppIni.ReadBool( 'ToolBarState', dxGroupByBox.Name, False );
  {}
  FAutoScroll := dxAutoScroll.Down;
  dxGroupByBox.DoClick;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.SaveToolBarStateToFile;
begin
  AppIni.WriteBool( 'ToolBarState', dxAutoScroll.Name, dxAutoScroll.Down );
  AppIni.WriteBool( 'ToolBarState', dxGroupByBox.Name, dxGroupByBox.Down );
  AppIni.UpdateFile;
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

procedure TfmMain.dxConfigListChange(Sender: TObject);
begin
  LoadConfigTimer.Enabled := True;
  ConfigModule.ActiveConfigFile := dxConfigList.Text;
  LoadConfigTimer.Interval := 500;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.MakeMsgItem(AItem: TListItem; ASubject: TSubject);
begin
  AItem.ImageIndex := StyleModule.MsgImageIndex( TMsgSubject( ASubject ).MsgText.Flag );
  AItem.Caption := Format( ' %s >   %s',
    [FormatDateTime( 'yyyy-mm-dd hh:mm:ss', Now ), TMsgSubject( ASubject ).MsgText.Text] );
  if ( FAutoScroll ) then
  begin
    AItem.Selected := True;
    AItem.Focused := True;
    AItem.MakeVisible( True );
  end;
  Application.ProcessMessages;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.OnMsgUpdate(ASubject: TSubject);
begin
  MakeMsgItem( MainMsgList.Items.Add, ASubject );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.OnDbSyncMsgUpdate(ASubject: TSubject);
begin
  MakeMsgItem( SoSyncMsgList.Items.Add, ASubject );
 end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.OnUserMsgUpdate(ASubject: TSubject);
begin
  MakeMsgItem( UserMsgList.Items.Add, ASubject );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.OnSoSessionUpdate(ASubject: TSubject);
var
  aSub: TSoSubject;
  aNode: TcxTreeListNode;
begin
  aSub := TSoSubject( ASubject );
  if not Assigned( aSub ) then Exit;
  if not Assigned( aSub.AppSo ) then Exit;
  if not Assigned( aSub.AppSo.Data ) then Exit;
  SoSessionTree.BeginUpdate;
  try
    aNode := TcxTreeListNode( aSub.AppSo.Data );
    aNode.ImageIndex := StyleModule.SoImageIndex( aSub.AppSo );
    aNode.SelectedIndex := aNode.ImageIndex;
    aNode.StateIndex := StyleModule.SoStateIndex( aSub.AppSo );
    aNode.Values[SoSessionTreeCol3.ItemIndex] := aSub.AppSo.FreeSession;
    aNode.Values[SoSessionTreeCol4.ItemIndex] := aSub.AppSo.BusySession;
    aNode.Values[SoSessionTreeCol5.ItemIndex] := aSub.AppSo.ErrSession;
  finally
    SoSessionTree.EndUpdate;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.OnSoTcpUpdate(ASubject: TSubject);
var
  aSubj: TSoSubject;
  aNode: TcxTreeListNode;
begin
  aSubj := TSoSubject( ASubject );
  if not Assigned( aSubj ) then Exit;
  if not Assigned( aSubj.AppSo ) then Exit;
  if not Assigned( aSubj.AppSo.Data ) then Exit;
  SoTcpTree.BeginUpdate;
  try
    aNode := TcxTreeListNode( aSubj.AppSo.Data );
    aNode.ImageIndex := StyleModule.SoImageIndex( aSubj.AppSo );
    aNode.SelectedIndex := aNode.ImageIndex;
    aNode.StateIndex := StyleModule.SoStateIndex( aSubj.AppSo );
    aNode.Values[SoTcpTreeCol3.ItemIndex] := aSubj.AppSo.Onlines;
    aNode.Values[SoTcpTreeCol4.ItemIndex] := aSubj.AppSo.Offlines;
  finally
    SoTcpTree.EndUpdate;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.OnUserListUpdate(ASubject: TSubject);
begin
  gvUser.DataController.BeginUpdate;
  try
    MainUIModule.UpdateUserBuffer( ASubject );
  finally
    gvUser.DataController.EndUpdate;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.OnDbWarningUpdate(ASubject: TSubject);
begin
  SocketWarningTimer.Enabled := ( TMsgSubject( ASubject ).MsgText.Flag <> MB_OK );
  if ( SocketWarningTimer.Enabled ) then
    dxStatusItem2.Caption := LanguageManager.Get( 'SRunWarning' )
  else begin
    dxStatusItem2.Caption := LanguageManager.Get( 'SRun' );
    dxStatusItem2.ImageIndex := StyleModule.SocketConnectImageIndex;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.BuildSoTree;
begin
  FMsgSub.Info( LanguageManager.Get( 'SConfigBuildSoTcpTreeBegin' ) );
  Delay( 1000 );
  BuildSoTcpTree;
  Delay( 500 );
  FMsgSub.OK( LanguageManager.Get( 'SConfigBuildSoTcpTreeEnd' ) );
  Delay( 500 );
  FMsgSub.Info( LanguageManager.Get( 'SConfigBuildSoSessionTreeBegin' ) );
  Delay( 1000 );
  BuildSoSessionTree;
  Delay( 500 );
  FMsgSub.OK( LanguageManager.Get( 'SConfigBuildSoSessionTreeEnd' ) );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.BuildSoTcpTree;
var
  aIndex: Integer;
  aSo: TAppSo;
  aNode: TcxTreeListNode;
begin
  SoTcpTree.Clear;
  for aIndex := 0 to ConfigModule.SoList.Count - 1 do
  begin
    aSo := ConfigModule.SoList[aIndex];
    aNode := SoTcpTree.Add;
    aNode.AssignValues( [aSO.CompCode, aSo.CompName, aSo.Onlines, aSo.Onlines] );
    aNode.ImageIndex := StyleModule.SoImageIndex( aSo );
    aNode.StateIndex := StyleModule.SoStateIndex( aSo );
    aNode.SelectedIndex := aNode.ImageIndex;
    aSo.Data := aNode;
    FMsgSub.InfoFmt( LanguageManager.Get( 'SConfigBuildSoTcpTree' ),[aSo.CompName] );
    Delay( 120 );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.BuildSoSessionTree;
var
  aIndex: Integer;
  aSo: TAppSo;
  aNode: TcxTreeListNode;
begin
  SoSessionTree.Clear;
  for aIndex := 0 to ConfigModule.SoList.Count - 1 do
  begin
    aSo := ConfigModule.SoList[aIndex];
    aNode := SoSessionTree.Add;
    aNode.AssignValues( [aSO.CompCode, aSo.CompName, aSo.FreeSession,
      aSo.BusySession, aSo.ErrSession] );
    aNode.ImageIndex := StyleModule.SoImageIndex( aSo );
    aNode.StateIndex := StyleModule.SoStateIndex( aSo );
    aNode.SelectedIndex := aNode.ImageIndex;
    aSo.Data := aNode;
    FMsgSub.InfoFmt( LanguageManager.Get( 'SConfigBuildSoSessionTree' ),[aSo.CompName] );
    Delay( 120 );
  end;
end;

{ --------------------------------------------------------------------------- }

procedure TfmMain.BeginExecute;
begin
  CreateSoDbSession;
  CreateThreads;
  {}
  ROServerModule.DbSessionManagerList := DbSessionManagerList;
  ROServerModule.ClientEnv.Assign( ConfigModule.ClientEnv );
  ROServerModule.SoList := ConfigModule.SoList;
  ROServerModule.ServiceStart;
  {}
  PlayGroup.Enabled := False;
  StopGroup.Enabled := True;
  dxConfigList.ReadOnly := True;
  {}
  dxStatusItem1.Caption := LanguageManager.GetFmt( dxStatusItem1.Name,
    [Format( '%s:%d',[GStack.LocalAddress, ROServerModule.ROSuperTcpServer.Port]) ] );
  dxStatusItem2.Caption := LanguageManager.Get( 'SRun' );
  dxStatusItem2.ImageIndex := StyleModule.SocketConnectImageIndex;
  {}
  AppRunState := arRun;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.EndExecute;
begin
  ROServerModule.ServiceStop;
  ROServerModule.DbSessionManagerList := nil;
  {}
  DestroyThreads;
  DestroySoDbSession;
  {}
  PlayGroup.Enabled := True;
  StopGroup.Enabled := False;
  dxConfigList.ReadOnly := False;
  {}
  SocketWarningTimer.Enabled := False;
  {}
  dxStatusItem1.Caption := LanguageManager.GetFmt( dxStatusItem1.Name,
    [LanguageManager.Get( 'SNull' )] );
  dxStatusItem2.Caption := LanguageManager.Get( 'Stop' );
  dxStatusItem2.ImageIndex := StyleModule.SocketDisconnectImageIndex;
  {}
  AppRunState := arIdle;
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

procedure TfmMain.dxAutoScrollClick(Sender: TObject);
begin
  FAutoScroll := dxAutoScroll.Down;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxGroupByBoxClick(Sender: TObject);
begin
  gvUser.OptionsView.GroupByBox := dxGroupByBox.Down;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxExpandClick(Sender: TObject);
begin
  gvUser.ViewData.Expand( True );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxCollapseClick(Sender: TObject);
begin
  gvUser.ViewData.Collapse( True );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.SocketWarningTimerTimer(Sender: TObject);
begin
  if ( dxStatusItem2.ImageIndex = StyleModule.SocketNullImageIndex ) then
    dxStatusItem2.ImageIndex := StyleModule.SocketWarningImageIndex
  else
    dxStatusItem2.ImageIndex := StyleModule.SocketNullImageIndex;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxMenuItem21Click(Sender: TObject);
begin
  FDocklayoutStream.Position := 0;
  dxDockingManager.LoadLayoutFromStream( FDocklayoutStream );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxMenuItem3Click(Sender: TObject);
begin
  fmAbout := TfmAbout.Create( nil );
  try
    fmAbout.ShowModal;
  finally
    fmAbout.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxMenuItem11Click(Sender: TObject);
begin
  if ( ConfigModule.ShowConfigForm = mrOk ) then
  begin
    BuildSoTree;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxMenuItem12Click(Sender: TObject);
begin
  ConfigModule.ShowGroupSetForm;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.gvUserCustomDrawCell(Sender: TcxCustomGridTableView;
  ACanvas: TcxCanvas; AViewInfo: TcxGridTableDataCellViewInfo;
  var ADone: Boolean);
var
  aDwRect: TRect;
  aFlag, aStatus: Integer;
begin
  aStatus := StrToInt( VarToStrDef( AViewInfo.GridRecord.Values[gvUserCol15.Index], '0' ) );
  ACanvas.Font.Color := clDefault;
  if aStatus = 0 then
    ACanvas.Font.Color := clGrayText;
  if ( AViewInfo.Item = gvUserCol3 ) then
  begin
    aDwRect := AViewInfo.ContentBounds;
    ACanvas.FillRect( aDwRect );
    InflateRect( aDwRect, -3, 0 );
    ACanvas.DrawImage( StyleModule.ImgList16, aDwRect.Left, aDwRect.Top,
      StyleModule.UserStateImageIndex( aStatus ) );
    aDwRect := AViewInfo.TextAreaBounds;
    InflateRect( aDwRect, ( 0 - ( StyleModule.ImgList16.Width + 8 ) ), 0 );
    aFlag := ( cxSingleLine or cxShowEndEllipsis or cxAlignLeft or cxAlignVCenter );
    ACanvas.DrawTexT( VarToStrDef( AViewInfo.DisplayValue, EmptyStr ), aDwRect, aFlag );
    ADone := True;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.CreateSoDbSession;
var
  aIndex: Integer;
  aAppSo: TAppSo;
begin
  FMsgSub.Info( LanguageManager.Get( 'SDbCreateSessionManager' ) );
  Delay( 1000 );
  DbSessionManagerList := TDbSessionManagerList.Create;
  for aIndex := 0 to ConfigModule.SoList.Count - 1 do
  begin
    aAppSo := ConfigModule.SoList[aIndex];
    if ( aAppSo.Selected ) or ( aAppSo.SynData ) then
    begin
      if ( aAppSo.MaxSession > 0 ) then
        DbSessionManagerList.Add( aAppSo )
      else begin
        FMsgSub.Warning( LanguageManager.GetFmt( 'SDbCreateSessionManagerWarning',
          [aAppSo.CompName, aAppSo.MaxSession] ) );
        Delay( 500 );
      end;
    end;
    Delay( 100 );
  end;
  DbSessionManagerList.CommEnv := ConfigModule.CommEnv;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.DestroySoDbSession;
begin
  try
    DbSessionManagerList.Clear;
    FreeAndNil( DbSessionManagerList );
    FMsgSub.Info( LanguageManager.Get( 'SDbDestorySessionManagerSuccess' ) );
  except
    on E: Exception do
    begin
      FMsgSub.Info( LanguageManager.GetFmt( 'SDbDestorySessionManagerError',
        [E.Message] ) );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.CreateThreads;
begin
  {}
  FMsgSub.Info( LanguageManager.Get( 'SDbCreateSessionMonitorThreadStart' ) );
  Delay( 1000 );
  FSessionMonitor := TSessionMonitorThread.Create( True );
  TSessionMonitorThread( FSessionMonitor ).CommEnv.Assign( ConfigModule.CommEnv );
  TSessionMonitorThread( FSessionMonitor ).MsgSubject.AddObServer( FMsgObsrv );
  TSessionMonitorThread( FSessionMonitor ).SoSubject.AddObServer( FDbSessionObsrv );
  TSessionMonitorThread( FSessionMonitor ).WarningSubject.AddObServer( FDbWarningObsrv );
  TSessionMonitorThread( FSessionMonitor ).DbSessionManagerList := DbSessionManagerList;
  Delay( 1000 );
  {}
  FMsgSub.OK( LanguageManager.Get( 'SDbCreateSyncThreadStart' ) );
  Delay( 1000 );
  FSyncDbThread := TSyncDbDataThread.Create( True );
  TSyncDbDataThread( FSyncDbThread ).CommEnv.Assign( ConfigModule.CommEnv );
  TSyncDbDataThread( FSyncDbThread ).Event := TSessionMonitorThread( FSessionMonitor ).Event;
  TSyncDbDataThread( FSyncDbThread ).MsgSubject.AddObServer( FDbSyncMsgObsrv );
  TSyncDbDataThread( FSyncDbThread ).DbSessionManagerList := DbSessionManagerList;
  Delay( 1000 );
  {}
  FMsgSub.Info( LanguageManager.Get( 'SUICreateThreadStart' ) );
  Delay( 1000 );
  FUIThread := TUpdateUIThread.Create( True );
  TUpdateUIThread( FUIThread ).CommEnv.Assign( ConfigModule.CommEnv );
  TUpdateUIThread( FUIThread ).Event := TSessionMonitorThread( FSessionMonitor ).Event;
  TUpdateUIThread( FUIThread ).MsgSubject.AddObServer( FUserMsgObsrv );
  TUpdateUIThread( FUIThread ).UserSubject.AddObServer( FUserListObsrv );
  TUpdateUIThread( FUIThread ).DbSessionManagerList := DbSessionManagerList;
  Delay( 1000 );
  {}
  Delay( 100 );
  FUIThread.Resume;
  {}
  Delay( 100 );
  FSyncDbThread.Resume;
  {}
  Delay( 100 );
  FSessionMonitor.Resume;
  {}
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.DestroyThreads;
begin
  FMsgSub.Info( LanguageManager.Get( 'SUIDestroyThreadStart' ) );
  Delay( 1000 );
  FUIThread.Terminate;
  FUIThread.WaitFor;
  FUIThread.Free;
  Delay( 1000 );
  {}
  FMsgSub.Info( LanguageManager.Get( 'SDbDestroySyncThreadStart' ) );
  Delay( 1000 );
  FSyncDbThread.Terminate;
  FSyncDbThread.WaitFor;
  FSyncDbThread.Free;
  Delay( 1000 );
  {}
  FMsgSub.Info( LanguageManager.Get( 'SDbDestroySessionMonitorThreadStart' ) );
  Delay( 1000 );
  {}
  FSessionMonitor.Terminate;
  FSessionMonitor.WaitFor;
  FSessionMonitor.Free;
  Delay( 1000 );
end;

{ ---------------------------------------------------------------------------- }

end.
