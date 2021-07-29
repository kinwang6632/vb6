unit cbMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, IniFiles, DB, StdCtrls, ExtCtrls, cbSyncObjs,
  ImgList, TrayIcon,
{$IFDEF APPDEBUG} CodeSiteLogging, {$ENDIF}  
{ App: }
  cbStyleModule, cbClientClass, cbSo, cbDesignPattern, cbDataClientModule,
{ Develpoer Express: }
  dxSkinsCore, dxSkinsDefaultPainters, dxSkinsdxRibbonPainter, dxSkinscxPCPainter,
  dxSkinsdxBarPainter, dxGDIPlusClasses, dxRibbonSkins, dxBarSkinConsts,
  cxStyles, cxLookAndFeels, cxClasses, cxGraphics, cxControls, cxContainer,
  cxCustomData, cxData, cxDataStorage, cxDBData, cxFilter,
  cxPC, cxEdit, cxInplaceContainer, dxBar, dxBarExtItems, dxRibbon, dxRibbonForm,
  cxGridCustomTableView, cxGridTableView, cxGridBandedTableView, cxGridDBBandedTableView,
  cxGridCustomView, cxGrid, cxGridLevel,
  cxGridStrs, dxBarStrs,
  cxTL, cxTLData, cxDBTL,
  cxTextEdit, cxMaskEdit,
{ ODAC: }
  VirtualTable, MemDS, cxLookAndFeelPainters, dxSkinBlack, dxSkinBlue,
  dxSkinCaramel, dxSkiniMaginary, dxSkinMoneyTwins, dxSkinOffice2007Black,
  dxSkinOffice2007Blue, dxSkinOffice2007Silver;

type
  TfmHrMain = class(TdxRibbonForm)
    BarManager: TdxBarManager;
    RibbonBar: TdxRibbon;
    rbTabAppearance: TdxRibbonTab;
    rbTabAction: TdxRibbonTab;
    barQuickAccess: TdxBar;
    dxBarPopupMenu1: TdxBarPopupMenu;
    MainPage: TcxPageControl;
    cxTabSheet1: TcxTabSheet;
    cxTabSheet2: TcxTabSheet;
    AnnGrid: TcxGrid;
    gvSo021: TcxGridDBBandedTableView;
    gvSo021Col1: TcxGridDBBandedColumn;
    gvSo021Col2: TcxGridDBBandedColumn;
    gvSo021Col3: TcxGridDBBandedColumn;
    gvSo021Col4: TcxGridDBBandedColumn;
    gvSo021Col5: TcxGridDBBandedColumn;
    glSo021: TcxGridLevel;
    gvSo021ColSp1: TcxGridDBBandedColumn;
    gvCd042: TcxGridDBBandedTableView;
    gvCd042Col1: TcxGridDBBandedColumn;
    gvCd042Col3: TcxGridDBBandedColumn;
    gvCd042Col4: TcxGridDBBandedColumn;
    gvCd042ColSp1: TcxGridDBBandedColumn;
    gvCd042Col2: TcxGridDBBandedColumn;
    gvSo022: TcxGridDBBandedTableView;
    gvSo022Col1: TcxGridDBBandedColumn;
    gvSo022Col2: TcxGridDBBandedColumn;
    gvSo022Col3: TcxGridDBBandedColumn;
    gvSo022Col4: TcxGridDBBandedColumn;
    gvSo022Col5: TcxGridDBBandedColumn;
    glCd042: TcxGridLevel;
    glSo022: TcxGridLevel;
    gvSo022Col6: TcxGridDBBandedColumn;
    glSo023: TcxGridLevel;
    gvSo023: TcxGridDBBandedTableView;
    gvSo023Col1: TcxGridDBBandedColumn;
    gvSo023Col2: TcxGridDBBandedColumn;
    gvSo023ColSp1: TcxGridDBBandedColumn;
    barDataFunc: TdxBar;
    barFullRefresh: TdxBarLargeButton;
    barMsgFunc: TdxBar;
    barSendMsg: TdxBarLargeButton;
    barViewUser: TdxBarLargeButton;
    barViewAnn: TdxBarLargeButton;
    barColor: TdxBar;
    barBlue: TdxBarLargeButton;
    barBlack: TdxBarLargeButton;
    barSilver: TdxBarLargeButton;
    barMarguee: TdxBarLargeButton;
    barShowApp: TdxBarButton;
    barAppExit: TdxBarButton;
    barFilterMsgIsReadItem: TdxBarLargeButton;
    UserTree: TcxDBTreeList;
    UserTreeCol1: TcxDBTreeListColumn;
    UserTreeCol2: TcxDBTreeListColumn;
    UserTreeCol3: TcxDBTreeListColumn;
    UserTreeCol4: TcxDBTreeListColumn;
    UserTreeCol5: TcxDBTreeListColumn;
    UserTreeCol6: TcxDBTreeListColumn;
    UserTreeCol7: TcxDBTreeListColumn;
    UserTreeCol8: TcxDBTreeListColumn;
    UserTreeCol9: TcxDBTreeListColumn;
    UserTreeCol10: TcxDBTreeListColumn;
    UserTreeCol11: TcxDBTreeListColumn;
    UserTreeCol12: TcxDBTreeListColumn;
    UserTreeCol13: TcxDBTreeListColumn;
    UserTreeCol14: TcxDBTreeListColumn;
    barUserGroupType1: TdxBarButton;
    barUserGroupType2: TdxBarButton;
    SoPage: TcxPageControl;
    barViewAnnSo021: TdxBarButton;
    barViewAnnCd042: TdxBarButton;
    barViewAnnSO022: TdxBarButton;
    UserTreeCol15: TcxDBTreeListColumn;
    cxTabSheet3: TcxTabSheet;
    barViewMsgList: TdxBarLargeButton;
    MsgTree: TcxTreeList;
    MsgTreeCol1: TcxTreeListColumn;
    MsgTreeCol2: TcxTreeListColumn;
    MsgTreeCol3: TcxTreeListColumn;
    MsgTreeCol4: TcxTreeListColumn;
    MsgTreeCol5: TcxTreeListColumn;
    GridPanel: TPanel;
    Image1: TImage;
    NotifyLabel: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
    procedure ColorChangeClick(Sender: TObject);
    procedure barMargueeClick(Sender: TObject);
    procedure gvSO021CustomDrawCell(Sender: TcxCustomGridTableView;
      ACanvas: TcxCanvas; AViewInfo: TcxGridTableDataCellViewInfo;
      var ADone: Boolean);
    procedure TrayIconMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure TrayIconDblClick(Sender: TObject);
    procedure barShowAppClick(Sender: TObject);
    procedure barAppExitClick(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure OnLoginUpdate(ASubject: TSubject);
    procedure OnSoCompanyUpdate(ASubject: TSubject);
    procedure GridPanelResize(Sender: TObject);
    procedure barFullRefreshClick(Sender: TObject);
    procedure barViewUserClick(Sender: TObject);
    procedure barViewAnnClick(Sender: TObject);
    procedure barUserGroupTypeClick(Sender: TObject);
    procedure SoPageChange(Sender: TObject);
    procedure barViewAnnTypeClick(Sender: TObject);
    procedure UserTreeSelectionChanged(Sender: TObject);
    procedure barSendMsgClick(Sender: TObject);
    procedure UserTreeDblClick(Sender: TObject);
    procedure barViewMsgListClick(Sender: TObject);
    procedure MsgTreeDblClick(Sender: TObject);
    procedure MsgTreeCustomDrawCell(Sender: TObject; ACanvas: TcxCanvas;
      AViewInfo: TcxTreeListEditCellViewInfo; var ADone: Boolean);
    procedure GridViewDblClick(Sender: TObject);
  private
    { Private declarations }
    FRealClose: Boolean;
    FSrvConnectSate: TServerConnectState;
    FMarqueeTemplate: String;
    FAppIni: TIniFile;
    FSo: TSo;
    FSoList: TSoList;
    FLoginObsrv: TObServer;
    FCompObsrv: TObServer;
    FMultiSelectUsers: Boolean;
    FOrignalLevel: Integer;
    FTrayIcon: TcbTrayIcon;
    function CreateAppIniFile: TIniFile;
    function AssignMsgRecver(ANode: TcxTreeListNode; ABuffer: TVirtualTable): Integer;
    procedure CreateThreads;
    procedure DestroyThreads;
    procedure InitResourceString;
    procedure LoadParamFromIni;
    procedure Cleanup;
    procedure NotifyProcess(const AProcId: Cardinal; AParamStr: String);
    procedure Logout;
  protected
    procedure WndProc(var Message: TMessage); override;
    procedure WMCopyData(var Message: TWMCopyData); message WM_COPYDATA;
  public
    { Public declarations }
    property TrayIcon: TcbTrayIcon read FTrayIcon;
    property MultiSelectUsers: Boolean read FMultiSelectUsers;
  end;

var
  fmHrMain: TfmHrMain;


implementation

uses
  cbUtilis, cbROClientModule, cbSendMsg, cbReadAnn,
  cbClientThread, cbAnnThread, cbCallbackEventThread,
  CsHorse2Library_Intf;

{$R *.dfm}

type
  PNotifyProcessData = ^TNotifyProcessData;
  TNotifyProcessData = record
    ProcessId: Cardinal;
    ClassName: String;
    WndHandle: THandle;
  end;

var
  FClientThread: TClientThread;
  FAnnThread: TAnnThread;
  FCallbackThread: TCallbackEventThread;
  FHost: String;
  FOldRibbonOwnerMinimalWidth: Integer;
  FOldRibbonOwnerMinimalHeight: Integer;

{ ---------------------------------------------------------------------------- }

function EnumWindowsCallback(Handle: THandle; LParam: PNotifyProcessData): Boolean; stdcall;
var
  aFindProcId: Cardinal;
  aClassName: PChar;
begin
  GetWindowThreadProcessId( Handle, aFindProcId );
  if ( aFindProcId = LParam.ProcessId ) then
  begin
    GetMem( aClassName, MAX_PATH );
    try
      GetClassName( Handle, aClassName, MAX_PATH );
      if ( aClassName = LParam.ClassName ) then
        LParam.WndHandle := Handle;
    finally
      FreeMem( aClassName );
    end;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.FormCreate(Sender: TObject);
var
  aProcId: Cardinal;
  aParam: String;
begin
  if IsProcessExists( ExtractFileName( Application.ExeName ) ) then
  begin
    aProcId := GetProcessId( ExtractFileName( Application.ExeName ) );
    aParam := CmdLine;
    ExtractValue( aParam, #32 );
    aParam := TrimLeft( aParam );
    NotifyProcess( aProcId, aParam );
    Application.ShowMainForm := False;
    Application.Terminate;
    Exit;
  end;
  FOldRibbonOwnerMinimalWidth := 180; //dxRibbonOwnerMinimalWidth;
  FOldRibbonOwnerMinimalHeight := dxRibbonOwnerMinimalHeight;
  {}
  ROClientModule := TROClientModule.Create( nil );
  DataClientModule := TDataClientModule.Create( nil );
  FRealClose := False;
  FSrvConnectSate := sctOnLine;
  FAppIni := CreateAppIniFile;
  LoadParamFromIni;
  InitResourceString;
  FMarqueeTemplate := IncludeTrailingPathDelimiter( ExtractFilePath(
    ParamStr( 0 ) ) ) + 'MarqueeTemplate.htm';
  //Application.OnMessage := OnApplicationMessage;
  {}
  FSo := TSo.Create;
  FSo.CompCode := ParamStr( 5 );
  FSo.DbUserId := ParamStr( 6 );
  {}
  FSoList := TSoList.Create;
  {}
  FLoginObsrv := TObServer.Create;
  FLoginObsrv.OnUpdate := OnLoginUpdate;
  FCompObsrv := TObServer.Create;
  FCompObsrv.OnUpdate := OnSoCompanyUpdate;
  {}
  FMultiSelectUsers := False;
  FTrayIcon := TcbTrayIcon.Create( nil );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.FormShow(Sender: TObject);
begin
  OnShow := nil;
  {}
  Self.Position := poDesigned;
  Self.Width := 300;
  Self.Top := 0;
  Self.Height := Screen.WorkAreaHeight;
  Self.Left := ( Screen.WorkAreaWidth - Self.Width );
  {}
  BarManager.ImageOptions.Images := StyleModule.SmallImages;
  BarManager.ImageOptions.LargeImages := StyleModule.LargeImages;
  {}
  FTrayIcon.Icons := StyleModule.TrayImageList;
  FTrayIcon.Visible := True;
  FTrayIcon.IconIndex := -1;
  FTrayIcon.IconIndex := 0;
  FTrayIcon.AnimateInterval := 300;
  FTrayIcon.BalloonFlags := bfInfo;
  FTrayIcon.BalloonTimeout := 10000;
  FTrayIcon.OnDblClick := TrayIconDblClick;
  FTrayIcon.OnMouseUp := TrayIconMouseUp;
  FTrayIcon.Refresh;
  {}
  MainPage.ActivePageIndex := 1;
  MainPage.HideTabs := True;
  MainPage.Align := alClient;
  MainPage.Visible := False;
  SoPage.Visible := False;
  SoPage.Width := 50;
  {}
  GridPanel.Caption := '尚未登入小幫手控制台';
  GridPanel.Align := alClient;
  GridPanel.Visible := True;
  NotifyLabel.Caption := EmptyStr;
  {}
  RibbonBar.ShowTabGroups := True;
  dxRibbonOwnerMinimalWidth := MaxInt;
  dxRibbonOwnerMinimalHeight := MaxInt;
  RibbonBar.CheckHide;
  {}
  gvSo021.OptionsView.BandHeaders := False;
  gvSo021.OptionsView.Header := False;
  gvCd042.OptionsView.BandHeaders := False;
  gvCd042.OptionsView.Header := False;
  gvSo022.OptionsView.BandHeaders := False;
  gvSo022.OptionsView.Header := False;
  gvSo023.OptionsView.BandHeaders := False;
  gvSo023.OptionsView.Header := False;
  {}
  UserTree.Images := StyleModule.SmallImages;
  UserTree.OptionsView.Headers := False;
  UserTree.OptionsView.Bands := False;
  UserTree.Bands[UserTree.Bands.Count-1].Visible := False;
  {}
  MsgTree.Images := StyleModule.SmallImages;
  MsgTree.StateImages := StyleModule.SmallImages;
  MsgTreeCol4.Visible := False;
  MsgTreeCol5.Visible := False;
  MsgTree.OptionsView.Headers := False;
  MsgTree.OptionsView.Bands := False;
  {}
  barViewUser.Down := ( MainPage.ActivePageIndex = 1 );
  if barViewUser.Down then
  begin
    barUserGroupType1.Visible := ivAlways;
    barUserGroupType2.Visible := ivAlways;
    barViewAnnSo021.Visible := ivNever;
    barViewAnnCd042.Visible := ivNever;
    barViewAnnSO022.Visible := ivNever;
    barUserGroupType1.Click;
  end else
  begin
    barUserGroupType1.Visible := ivNever;
    barUserGroupType2.Visible := ivNever;
    barViewAnnSo021.Visible := ivAlways;
    barViewAnnCd042.Visible := ivAlways;
    barViewAnnSO022.Visible := ivAlways;
    barViewAnnSo021.Click;
  end;
  {}
  barBlue.Click;
  {}
  RibbonBar.QuickAccessToolbar.Position := qtpBelowRibbon;
  StyleModule.ChangeColorSchema( EmptyStr );
  {}
  CreateThreads;
  Delay( 200 );
  DecreaseProcessMemorySize;
  if not Self.Visible then Self.Show;
  SetForegroundWindow( Self.Handle );
  SetActiveWindow( Self.Handle );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if FRealClose then
  begin
    Logout;
    Action := caFree
  end else
  begin
    Action := caNone;
    Self.Hide;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.FormDestroy(Sender: TObject);
begin
  FLoginObsrv.Free;
  FCompObsrv.Free;
  FAppIni.Free;
  FSoList.Free;
  FSo.Free;
  DataClientModule.Free;
  ROClientModule.Free;
  FTrayIcon.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.WndProc(var Message: TMessage);
begin
  inherited WndProc( Message );
  if ( Message.Msg = WM_REALCLOSE ) then
  begin
    {$IFDEF APPDEBUG} CodeSite.SendReminder( 'Recv CSMIS Close Event' ); {$ENDIF}
    FRealClose := True;
    Self.Close;
  end else
  if ( Message.Msg = WM_CSMISCLOSE ) then
  begin
    {$IFDEF APPDEBUG} CodeSite.SendReminder( 'Thread Doesn''t Find CSMIS.exe' ); {$ENDIF}
    FRealClose := True;
    Self.Close;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.WMCopyData(var Message: TWMCopyData);
var
  aCmdParam: String;
begin
  aCmdParam := PChar( Message.CopyDataStruct.lpData );
  {$IFDEF APPDEBUG} CodeSite.Send( 'Recv Change Param:' + aCmdParam ); {$ENDIF}
  Logout;
  FSo.CompCode := ExtractValue( 5, aCmdParam, #32 );
  FSo.DbUserId := ExtractValue( 6, aCmdParam, #32 );
  FormShow( Self );
  Message.Result := 0;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.FormResize(Sender: TObject);
begin
  TcxGridDBBandedTableView( AnnGrid.ActiveView ).Bands[0].Width := Self.Width - 30;
  UserTree.Bands[0].Width := Self.Width - 30;
  MsgTree.Bands[0].Width := Self.Width - 10;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.InitResourceString;
begin
  cxSetResourceString(@dxSBAR_SHOWBELOWRIBBON, '快取列顯示在下方' );
  cxSetResourceString(@dxSBAR_SHOWABOVERIBBON, '快取列顯示在上方' );
  cxSetResourceString(@dxSBAR_MINIMIZERIBBON, '自動隱藏功能列' );
  cxSetResourceString(@scxGridNoDataInfoText, '無符合資料' );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.ColorChangeClick(Sender: TObject);
begin
  case TdxBarItem( Sender ).Tag of
    0: StyleModule.ChangeColorSchema( 'Blue' );
    1: StyleModule.ChangeColorSchema( 'Black' );
    2: StyleModule.ChangeColorSchema( 'Silver' );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.barShowAppClick(Sender: TObject);
begin
  if ( not Self.Visible ) then
    Self.Show;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.barAppExitClick(Sender: TObject);
begin
  FRealClose := True;
  Self.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.barMargueeClick(Sender: TObject);
//var
//  aEnabled: Boolean;
begin
  if barMarguee.Down then
  begin
    //MainPage.ActivePageIndex := 1;
    dxRibbonOwnerMinimalWidth := MaxInt;
    dxRibbonOwnerMinimalHeight := MaxInt;
    RibbonBar.CheckHide;
  end else
  begin
    //MainPage.ActivePageIndex := 0;
    //WebBrowser.Stop;
    dxRibbonOwnerMinimalWidth := FOldRibbonOwnerMinimalWidth;
    dxRibbonOwnerMinimalHeight := FOldRibbonOwnerMinimalHeight;
    RibbonBar.CheckHide;
  end;
  //aEnabled := ( MainPage.ActivePageIndex = 1 );
  //barFilterMsgIsReadItem.Enabled := aEnabled;
  //barSoItem.Enabled := aEnabled;
  //barMsgGroupItem.Enabled := aEnabled;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.SoPageChange(Sender: TObject);
var
  aCompCode: String;
begin
  Screen.Cursor := crAppStart;
  try
    if ( SoPage.ActivePageIndex >= 0 ) then
    begin
      aCompCode := TSo( SoPage.Tabs.Objects[SoPage.ActivePageIndex] ).CompCode;
      gvSo021.BeginUpdate;
      gvCd042.BeginUpdate;
      gvSo022.BeginUpdate;
      gvSo023.BeginUpdate;
      UserTree.BeginUpdate;
      try
        DataClientModule.SwitchCompCode( aCompCode );
      finally
        gvSo021.EndUpdate;
        gvCd042.EndUpdate;
        gvSo022.EndUpdate;
        gvSo023.EndUpdate;
        UserTree.EndUpdate;
      end;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.barViewUserClick(Sender: TObject);
begin
  MainPage.ActivePageIndex := 1;
  barSendMsg.Enabled := True;
  {}
  barUserGroupType1.Visible := ivAlways;
  barUserGroupType2.Visible := ivAlways;
  {}
  barViewAnnSo021.Visible := ivNever;
  barViewAnnCd042.Visible := ivNever;
  barViewAnnSO022.Visible := ivNever;
  RibbonBar.QuickAccessToolbar.Visible := True;
  SoPage.Visible := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.barViewAnnClick(Sender: TObject);
begin
  MainPage.ActivePageIndex := 0;
  barSendMsg.Enabled := False;
  {}
  barUserGroupType1.Visible := ivNever;
  barUserGroupType2.Visible := ivNever;
  {}
  barViewAnnSo021.Visible := ivAlways;
  barViewAnnCd042.Visible := ivAlways;
  barViewAnnSO022.Visible := ivAlways;
  {}
  RibbonBar.QuickAccessToolbar.Visible := True;
  SoPage.Visible := True;
  UserTree.ClearSelection( True );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.barViewMsgListClick(Sender: TObject);
begin
  MainPage.ActivePageIndex := 2;
  barSendMsg.Enabled := False;
  {}
  barUserGroupType1.Visible := ivNever;
  barUserGroupType2.Visible := ivNever;
  {}
  barViewAnnSo021.Visible := ivNever;
  barViewAnnCd042.Visible := ivNever;
  barViewAnnSO022.Visible := ivNever;
  {}
  RibbonBar.QuickAccessToolbar.Visible := False;
  SoPage.Visible := False;
  UserTree.ClearSelection( True );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.gvSO021CustomDrawCell(Sender: TcxCustomGridTableView;
  ACanvas: TcxCanvas; AViewInfo: TcxGridTableDataCellViewInfo;
  var ADone: Boolean);
//var
//  aItem: TcxGridDBBandedColumn;
begin
//  aItem := TcxGridDBBandedColumn( AViewInfo.Item );
//  if ( UpperCase( aItem.DataBinding.FieldName ) = 'MSGTEXT' ) then
//  begin
//    if ( VarToStrDef( AViewInfo.GridRecord.Values[
//      gvMsgTextMsgIsRead.Index], EmptyStr ) <> '1' ) then
//    begin
//      ACanvas.Font.Color := clRed;
//      if ( AViewInfo.Focused or AViewInfo.Selected ) then
//        ACanvas.Font.Color := clHighlightText;
//      ACanvas.Font.Style := ( ACanvas.Font.Style + [fsBold] );
//    end;
//  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.TrayIconDblClick(Sender: TObject);
begin
  if ( Self.Visible ) then
    Self.Hide
  else begin
    if FTrayIcon.Animate then
    begin
      barViewMsgList.DoClick;
      barViewMsgList.Down := True;
    end;
    Self.Show;
    if Self.WindowState = wsMinimized then
      Self.WindowState := wsNormal;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.TrayIconMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  if ( Button = mbRight ) then
  begin
    SetForegroundWindow( Application.Handle );
    Application.ProcessMessages;
    dxBarPopupMenu1.PopupFromCursorPos;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfmHrMain.CreateAppIniFile: TIniFile;
begin
  Result := FAppIni;
  if not Assigned( Result ) then
    Result := TIniFile.Create( IncludeTrailingPathDelimiter( ExtractFilePath(
     ParamStr( 0 ) ) ) + FileNameWithoutExt( ExtractFileName( ParamStr( 0 ) ) ) + '.ini' );
end;

{ ---------------------------------------------------------------------------- }

function TfmHrMain.AssignMsgRecver(ANode: TcxTreeListNode; ABuffer: TVirtualTable): Integer;
begin
  Result := 0;
  while True do
  begin
    if ( ANode.HasChildren ) then
      AssignMsgRecver( ANode.getFirstChild, ABuffer );
    {}
    if VarToStrDef( ANode.Values[UserTreeCol15.ItemIndex], '0' ) = '2' then
    begin
      if not ABuffer.Locate( 'CompCode;UserId', VarArrayOf( [
        VarToStrDef( ANode.Values[UserTreeCol1.ItemIndex], EmptyStr ),
        VarToStrDef( ANode.Values[UserTreeCol3.ItemIndex], EmptyStr )] ), [] ) then
      begin
        ABuffer.Append;
        ABuffer.FieldByName( 'CompCode' ).AsString :=
          VarToStrDef( ANode.Values[UserTreeCol1.ItemIndex], EmptyStr );
        ABuffer.FieldByName( 'UserId' ).AsString :=
          VarToStrDef( ANode.Values[UserTreeCol3.ItemIndex], EmptyStr );
        ABuffer.FieldByName( 'Status' ).AsString :=
          VarToStrDef( ANode.Values[UserTreeCol9.ItemIndex], '0' );
        ABuffer.FieldByName( 'DisplayText' ).AsString :=
          VarToStrDef( ANode.Values[UserTreeCol14.ItemIndex ], EmptyStr );
        ABuffer.Post;
        Inc( Result );
      end;
    end;
    {}
    ANode := ANode.getNextSibling;
    if not Assigned( ANode ) then Break;
    if ANode.Level = FOrignalLevel then Break;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.LoadParamFromIni;
begin
  FHost := FAppIni.ReadString( 'Common', 'Host', EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.OnLoginUpdate(ASubject: TSubject);
var
  aText, aErrCode, aStatus: String;
begin
  NotifyLabel.Caption := EmptyStr;
  aText := TMsgSubject( ASubject ).MsgText.Text;
  aStatus := ExtractValue( aText );
  { 已離線 }
  if ( aStatus = 'A0' ) then
  begin
    GridPanel.Visible := True;
    MainPage.Visible := False;
    SoPage.Visible := False;
    GridPanel.Caption := '已登出小幫手控制台';
    Self.Caption := '已登出';
    Screen.Cursor := crDefault;
  end else
  { 登入中 }
  if ( aStatus = 'A1' ) then
  begin
    GridPanel.Visible := True;
    MainPage.Visible := False;
    SoPage.Visible := False;
    GridPanel.Caption := '登入小幫手控制台中...';
    Self.Caption := '登入中';
    Screen.Cursor := crHourGlass;
  end else
  { 已登入 }
  if ( aStatus = 'A2' ) then
  begin
    GridPanel.Visible := False;
    MainPage.Visible := True;
    SoPage.Visible := True;
    dxRibbonOwnerMinimalWidth := FOldRibbonOwnerMinimalWidth;
    dxRibbonOwnerMinimalHeight := FOldRibbonOwnerMinimalHeight;
    RibbonBar.CheckHide;
    Self.Caption := Format( '客服小幫手-%s(%s)', [ROClientModule.LoginInfo.UserId,
      ROClientModule.LoginInfo.UserName] );
    FTrayIcon.Icons := StyleModule.TrayAnimateImageList;
    FTrayIcon.IconIndex := 0;
    Screen.Cursor := crDefault;
    DataClientModule.NotifyMsgForm( True );
  end else
  { 登出中 }
  if ( aStatus = 'A3' ) then
  begin
    dxRibbonOwnerMinimalWidth := MaxInt;
    dxRibbonOwnerMinimalHeight := MaxInt;
    RibbonBar.CheckHide;
    GridPanel.Visible := True;
    MainPage.Visible := False;
    SoPage.Visible := False;
    GridPanel.Caption := '登出小幫手控制台中...';
    Self.Caption := '登出中';
    FTrayIcon.Icons := StyleModule.TrayImageList;
    FTrayIcon.IconIndex := 0;
    Screen.Cursor := crHourGlass;
  end else
  { 連線中 }
  if ( aStatus = 'C1' ) then
  begin
    dxRibbonOwnerMinimalWidth := MaxInt;
    dxRibbonOwnerMinimalHeight := MaxInt;
    RibbonBar.CheckHide;
    GridPanel.Visible := True;
    MainPage.Visible := False;
    SoPage.Visible := False;
    GridPanel.Caption := '與小幫手控制台連線中...';
    Self.Caption := '連線中';
    Application.Title := '連線中';
    Screen.Cursor := crHourGlass;
  end else
  { 連線發生錯誤 }
  if ( aStatus = 'E0' ) then
  begin
    MainPage.Visible := False;
    SoPage.Visible := False;
    GridPanel.Visible := True;
    GridPanel.Caption := '無法連線小幫手控制台';
    aErrCode := aText;
    if ( aErrCode <> EmptyStr ) then
      NotifyLabel.Caption := '錯誤代碼:' + aErrCode;
    Screen.Cursor := crDefault;
    Self.Caption := '連線錯誤';
    FTrayIcon.Icons := StyleModule.TrayImageList;
    FTrayIcon.IconIndex := 0;
    FTrayIcon.BalloonTitle := EmptyStr;
    FTrayIcon.BalloonHint := EmptyStr;
    FTrayIcon.Animate := False;
    FTrayIcon.Hint := '客服小幫手';
    DataClientModule.NotifyMsgForm( False );
  end;
  { 登入發生錯誤 }
  if ( aStatus = 'E1' ) then
  begin
    MainPage.Visible := False;
    SoPage.Visible := False;
    GridPanel.Visible := True;
    GridPanel.Caption := '無法登入小幫手控制台';
    aErrCode := aText;
    if ( aErrCode <> EmptyStr ) then
      NotifyLabel.Caption := '錯誤代碼:' + aErrCode;
    Screen.Cursor := crDefault;
    Self.Caption := '登入錯誤';
    FTrayIcon.Icons := StyleModule.TrayImageList;
    FTrayIcon.BalloonTitle := EmptyStr;
    FTrayIcon.BalloonHint := EmptyStr;
    FTrayIcon.Animate := False;
    FTrayIcon.Hint := '客服小幫手';
    DataClientModule.NotifyMsgForm( False );
  end;
  {}
  if ( aStatus = 'A0' ) or ( aStatus = 'E0' )  or ( aStatus = 'E1' ) then
    Cleanup;
  {}
  Delay( 250 );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.OnSoCompanyUpdate(ASubject: TSubject);

  { ------------------------------------------ }

  procedure CleanupTabSheets;
  var
    aIndex: Integer;
    aTab: TcxTabSheet;
  begin
    for aIndex := SoPage.PageCount - 1 downto 0 do
    begin
      aTab := SoPage.Pages[aIndex];
      aTab.PageControl := nil;
      FreeAndNil( aTab );
    end;
  end;

  { ------------------------------------------ }

  procedure NewTabSheet(aObj: TObject);
  var
    aTab: TcxTabSheet;
  begin
    aTab := TcxTabSheet.Create( SoPage );
    aTab.Caption := Rpad( Copy( TSo( aObj ).CompName, 1, 6 ), 7, #32 );
    aTab.PageControl := SoPage;
    SoPage.Tabs.Objects[aTab.PageIndex] := aObj;
  end;

  { ------------------------------------------ }

var
  aEvent: TNotifyEvent;
  aSource: TSoList;
  aIndex, aItemIndex: Integer;
  aOldCompCode: String;
  aIsFirtsRun: Boolean;
begin
  Screen.Cursor := crAppStart;
  try
    aSource := TSoList( TDataSubject( ASubject ).Data );
    if Assigned( aSource ) then
    begin
      if ( aSource.Count > 0 ) then
      begin
        aIsFirtsRun := ( FSoList.Count <= 0 );
        aOldCompCode := FSo.CompCode;
        if ( SoPage.ActivePageIndex >= 0 ) then
          aOldCompCode := TSo( SoPage.Tabs.Objects[SoPage.ActivePageIndex] ).CompCode;
        aEvent := SoPage.OnChange;
        SoPage.OnChange := nil;
        try
          CleanupTabSheets;
          FSoList.Clear;
          FSoList.Assign( aSource );
          aItemIndex := -1;
          for aIndex := 0 to FSoList.Count - 1 do
          begin
            NewTabSheet( FSoList[aIndex] );
            if ( FSoList[aIndex].CompCode = aOldCompCode ) then
              aItemIndex := aIndex;
          end;
          if ( aItemIndex = -1 ) and ( SoPage.PageCount > 0 ) then aItemIndex := 0;
          SoPage.ActivePageIndex := aItemIndex;
        finally
          SoPage.OnChange := aEvent;
        end;
        if ( aIsFirtsRun ) then
          SoPageChange( SoPage );
      end;
      aSource.Free;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.CreateThreads;
begin
  FClientThread := TClientThread.Create( True );
  FClientThread.Host := FHost;
  FClientThread.CompCode := FSo.CompCode;
  FClientThread.UserId := FSo.DbUserId;
  FClientThread.UserPass := FSo.DbUserPass;
  FClientThread.MsgSubject.AddObServer( FLoginObsrv );
  FClientThread.MainFormHandle := Self.Handle;
  {}
  FCallbackThread := TCallbackEventThread.Create( True );
  FCallbackThread.ReadyEvent := FClientThread.ReadyEvent;
  FCallbackThread.UserSubject.AddObServer( DataClientModule.UserListObsrv );
  FCallbackThread.MsgSubject.AddObServer( DataClientModule.MsgListObsrv );
  {}
  FAnnThread := TAnnThread.Create( True );
  FAnnThread.ReadyEvent := FClientThread.ReadyEvent;
  FAnnThread.CompSubject.AddObServer( FCompObsrv );
  FAnnThread.So021Subject.AddObServer( DataClientModule.So021Obsrv );
  FAnnThread.SO022Subject.AddObServer( DataClientModule.So022Obsrv );
  FAnnThread.SO023Subject.AddObServer( DataClientModule.So023Obsrv );
  FAnnThread.Cd042Subject.AddObServer( DataClientModule.Cd042Obsrv );  
  {}
  FCallbackThread.Resume;
  FAnnThread.Resume;
  FClientThread.Resume;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.DestroyThreads;
begin
  try
    FCallbackThread.Terminate;
    FCallbackThread.WaitFor;
    FCallbackThread.Free;
  except
    {}
  end;
  Delay( 50 );
  try
    FAnnThread.Terminate;
    FAnnThread.WaitFor;
    FAnnThread.Free;
  except
    {}
  end;
  Delay( 50 );
  try
    FClientThread.Terminate;
    FClientThread.WaitFor;
    FClientThread.Free;
  except
    {}
  end;
  Delay( 50 );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.GridPanelResize(Sender: TObject);
begin
  NotifyLabel.Width := GridPanel.Width - 3;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.barFullRefreshClick(Sender: TObject);
begin
  Screen.Cursor := crAppStart;
  try
    Delay( 100 );
    PostMessage( FAnnThread.MsgWnd, WM_MANUALREFRESH, 0, 0 );
    PostMessage( FCallbackThread.MsgWnd, WM_MANUALREFRESH, 0, 0 );
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.barUserGroupTypeClick(Sender: TObject);
begin
  UserTree.BeginUpdate;
  try
    if Sender = barUserGroupType1 then
      DataClientModule.SwitchUserDisplayType( gkWorkClass )
    else
      DataClientModule.SwitchUserDisplayType( gkStatus );
  finally
    UserTree.EndUpdate;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.barViewAnnTypeClick(Sender: TObject);
begin
  Screen.Cursor := crAppStart;
  try
    if ( Sender = barViewAnnSo021 ) then
    begin
      glSo021.Visible := True;
      glCd042.Visible := False;
      glSo022.Visible := False;
    end else
    if ( Sender = barViewAnnCd042 ) then
    begin
      glSo021.Visible := False;
      glCd042.Visible := True;
      glSo022.Visible := False;
    end else
    if ( Sender = barViewAnnSO022 ) then
    begin
      glSo021.Visible := False;
      glCd042.Visible := False;
      glSo022.Visible := True;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.barSendMsgClick(Sender: TObject);
var
  aList: TList;
  aIndex, aCount: Integer;
  aNode: TcxTreeListNode;
begin
  aCount := 0;
  fmSendMsg := TfmSendMsg.Create( Application );
  FMultiSelectUsers := True;
  try
    aList := TList.Create;
    try
      UserTree.GetSelections( aList );
      for aIndex := 0 to aList.Count - 1 do
      begin
        aNode := TcxTreeListNode( aList[aIndex] );
        FOrignalLevel := aNode.Level;
        if aNode.HasChildren then aNode := aNode.getFirstChild;
        aCount := ( aCount + AssignMsgRecver( aNode, fmSendMsg.BfRecver ) );
      end;
    finally
      aList.Free;
    end;
    if ( aCount <= 0 ) then
      fmSendMsg.Free
    else begin
      fmSendMsg.MsgFormType := 0;
      fmSendMsg.Show;
    end;
  finally
    FMultiSelectUsers := False;
    UserTree.ClearSelection( True );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.UserTreeDblClick(Sender: TObject);
begin
  UserTree.HitTest.ReCalculate;
  if ( UserTree.HitTest.HitAtNode ) and ( not UserTree.HitTest.HitAtButton) then
    barSendMsg.DoClick;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.MsgTreeDblClick(Sender: TObject);
begin
  MsgTree.HitTest.ReCalculate;
  if ( MsgTree.HitTest.HitAtNode ) and Assigned( MsgTree.HitTest.HitNode ) then
  begin
    if Assigned( MsgTree.HitTest.HitNode.Data ) then
    begin
      fmSendMsg := TfmSendMsg( DataClientModule.GetMsgForm( TMsgInfo(
        MsgTree.HitTest.HitNode.Data ) ) );
      if Assigned( fmSendMsg ) then
      begin
        if IsIconic( fmSendMsg.Handle ) then fmSendMsg.WindowState := wsNormal;
        SetActiveWindow( fmSendMsg.Handle  );
        SetForegroundWindow( fmSendMsg.Handle );
      end else
      begin
        fmSendMsg := TfmSendMsg.Create( Application );
        fmSendMsg.MsgFormType := 1;
        fmSendMsg.MsgInfo := TMsgInfo( MsgTree.HitTest.HitNode.Data );
        fmSendMsg.MsgNode := MsgTree.HitTest.HitNode;
        fmSendMsg.Show;
      end;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.UserTreeSelectionChanged(Sender: TObject);
begin
  FMultiSelectUsers := ( UserTree.SelectionCount > 1 );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.Cleanup;
var
  aIndex: Integer;
begin
  for aIndex := 0 to MsgTree.Count - 1 do
  begin
    if Assigned( MsgTree.Nodes.AbsoluteItems[aIndex].Data ) then
    begin
      TMsgInfo( MsgTree.Nodes.AbsoluteItems[aIndex].Data ).Free;
      MsgTree.Nodes.AbsoluteItems[aIndex].Data := nil;
    end;
  end;
  MsgTree.Clear;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.MsgTreeCustomDrawCell(Sender: TObject; ACanvas: TcxCanvas;
  AViewInfo: TcxTreeListEditCellViewInfo; var ADone: Boolean);
var
  aMsgInfo: TMsgInfo;
begin
  if ( AViewInfo.Column = MsgTreeCol1 ) then
  begin
    aMsgInfo := ( AViewInfo.Node.Data );
    if Assigned( aMsgInfo ) then
    begin
      if ( not aMsgInfo.IsRead ) then
      begin
        ACanvas.Font.Color := clRed;
        if ( AViewInfo.Focused or AViewInfo.Selected ) then
          ACanvas.Font.Color := clHighlightText;
        ACanvas.Font.Style := ( ACanvas.Font.Style + [fsBold] );
      end else
      begin
        ACanvas.Font.Color := clWindowText;
        if ( AViewInfo.Focused or AViewInfo.Selected ) then
          ACanvas.Font.Color := clHighlightText;
        ACanvas.Font.Style := ( ACanvas.Font.Style - [fsBold] );
      end;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.NotifyProcess(const AProcId: Cardinal; AParamStr: String);
var
  aEnumInfo: PNotifyProcessData;
  aCopyData: TCopyDataStruct;
begin
  New( aEnumInfo );
  try
    aEnumInfo.ProcessId := AProcId;
    aEnumInfo.ClassName := Self.ClassName;
    EnumWindows( @EnumWindowsCallback, Integer( aEnumInfo ) );
    if ( aEnumInfo.WndHandle > 0 ) then
    begin
      aCopyData.cbData := Length( AParamStr ) + 1;
      aCopyData.lpData := PChar( AParamStr );
      SendMessage( aEnumInfo.WndHandle, WM_COPYDATA, Self.Handle, Integer( @aCopyData ) );
    end;
  finally
    Dispose( aEnumInfo );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.Logout;
begin
  DestroyThreads;
  DataClientModule.ForceFormClose;
  FTrayIcon.Animate := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmHrMain.GridViewDblClick(Sender: TObject);
begin
  fmReadAnn := TfmReadAnn.Create( nil );
  fmReadAnn.AnnSyncMutex := FAnnThread.AnnSyncMutex;
  if ( TcxGridSite( Sender ).GridView = gvSo022 ) then
    fmReadAnn.SetBuffer( gvSo022.DataController.DataSet, gvSo023.DataController.DataSet )
  else
    fmReadAnn.SetBuffer( TcxGridDBBandedTableView( TcxGridSite(
      Sender ).GridView ).DataController.DataSet, nil );
  fmReadAnn.Show;
end;

{ ---------------------------------------------------------------------------- }

end.



