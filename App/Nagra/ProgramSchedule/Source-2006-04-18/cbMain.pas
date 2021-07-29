unit cbMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, ComCtrls, GraphicEx, ActnList,
  cbStyleModule, cbConfigModule, cbLogModule, cbAppClass, cbClass,
  RzPanel,
  dxDockControl, dxBar, dxNavBarCollns, dxNavBarBase, dxNavBar, cxContainer,
  cxListView, cxPC, cxControls, dxDockPanel, cxStyles, cxCustomData, cxGraphics,
  cxFilter, cxData, cxDataStorage, cxEdit, DB, cxDBData, cxGridLevel, cxClasses,
  cxGridCustomView, cxGridCustomTableView, cxGridTableView, cxGridDBTableView,
  cxGrid, cxGridStrs, cxImageComboBox;

type
  TfmMain = class(TForm)
    dxBarManager: TdxBarManager;
    dxPlay: TdxBarButton;
    dxStop: TdxBarButton;
    dxAutoScroll: TdxBarButton;
    dxAbout: TdxBarButton;
    PlayCmdGroup: TdxBarGroup;
    StopCmdGroup: TdxBarGroup;
    ControlSendGroup: TdxBarGroup;
    dxDockingManager: TdxDockingManager;
    StartupTimer: TTimer;
    dxDockMain: TdxDockSite;
    dxDockMsg: TdxDockPanel;
    dxDockClient: TdxDockPanel;
    dxLayoutDockSite2: TdxLayoutDockSite;
    RzTitlePanel: TRzPanel;
    dxLayoutDockSite4: TdxLayoutDockSite;
    Bevel1: TBevel;
    dxDockNavBar: TdxDockPanel;
    dxLayoutDockSite1: TdxLayoutDockSite;
    RzPanelNavBar: TRzPanel;
    dxNavBar: TdxNavBar;
    dxNavBarGroup1: TdxNavBarGroup;
    dxAction: TdxNavBarItem;
    dxList: TdxNavBarItem;
    dxOption: TdxNavBarItem;
    MsgList: TcxListView;
    MainPage: TcxPageControl;
    cxTabAction: TcxTabSheet;
    cxTabList: TcxTabSheet;
    PageMarkImage: TImage;
    lblTitle: TLabel;
    BarActionList: TActionList;
    actPlay: TAction;
    actStop: TAction;
    actAutoScroll: TAction;
    ActionGridTableView: TcxGridDBTableView;
    ActionGridLevel1: TcxGridLevel;
    ActionGrid: TcxGrid;
    ActionDataSource: TDataSource;
    ColumnKeyId: TcxGridDBColumn;
    ColumnActDate: TcxGridDBColumn;
    ColumnActTime: TcxGridDBColumn;
    ColumnActFileName: TcxGridDBColumn;
    ColumnActFilePath: TcxGridDBColumn;
    ColumnActSource: TcxGridDBColumn;
    ColumnActStatus: TcxGridDBColumn;
    ColumnActCost: TcxGridDBColumn;
    ColumnActProgress: TcxGridDBColumn;
    ColumnActErrMsg: TcxGridDBColumn;
    ColumnActFileSize: TcxGridDBColumn;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormDestroy(Sender: TObject);
    procedure OnMsgUpdate(aSubject: TSubject);
    procedure OnXmlUpdate(aSubject: TSubject);
    procedure MainPageChange(Sender: TObject);
    procedure dxNavBarLinkClick(Sender: TObject; ALink: TdxNavBarItemLink);
    procedure StartupTimerTimer(Sender: TObject);
    procedure dxActionClick(Sender: TObject);
    procedure dxListClick(Sender: TObject);
    procedure dxOptionClick(Sender: TObject);
    procedure actPlayExecute(Sender: TObject);
    procedure actStopExecute(Sender: TObject);
    procedure ActionGridTableViewCustomDrawCell(
      Sender: TcxCustomGridTableView; ACanvas: TcxCanvas;
      AViewInfo: TcxGridTableDataCellViewInfo; var ADone: Boolean);
    procedure ColumnActCostGetDisplayText(Sender: TcxCustomGridTableItem;
      ARecord: TcxCustomGridRecord; var AText: String);
    procedure actAutoScrollExecute(Sender: TObject);
  private
    { Private declarations }
    FRunSate: TAppRunSate;
    FAutoScroll: Boolean;
    FMsgSubject: TMessageSubject;
    FMsgObServer: TObServer;
    FParseObServer: TObServer;
    FWatchThread: TThread;
    FParseThread: TThread;
    procedure ChangeDevexComponentLang;
    procedure InitialPageControl;
    procedure InitialUIState(const aEnable: Boolean);
    function CreateAllThread: Boolean;
    procedure DestroyAllThread;
    function CreateWatchThread: Boolean;
    procedure DestroyWatchThread;
    function CreateParseThread: Boolean;
    procedure DestroyParseThread;
  public
    { Public declarations }
  end;

var
  fmMain: TfmMain;

implementation

uses cbUtilis, cbOption, cbWatchThread, cbParseThread, Types;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormCreate(Sender: TObject);
begin
  FRunSate := rsStop;
  FAutoScroll := False;
  FMsgObServer := TObServer.Create;
  FMsgObServer.OnUpdate := OnMsgUpdate;
  FMsgSubject := TMessageSubject.Create;
  FMsgSubject.AddObServer( FMsgObServer );
  FParseObServer := TObServer.Create;
  FParseObServer.OnUpdate := OnXmlUpdate;
  ConfigModule := TConfigModule.Create( nil );
  ConfigModule.MessageSubject.AddObServer( FMsgObServer );
  LogModule := TLogModule.Create( nil );
  LogModule.MessageSubject.AddObServer( FMsgObServer );
  cxTabAction.Caption := dxAction.Caption;
  cxTabList.Caption := dxList.Caption;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormShow(Sender: TObject);
begin
  InitialPageControl;
  ChangeDevexComponentLang;
  dxAutoScroll.Click;
  StartupTimer.Enabled := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  if ( FRunSate = rsPlay ) then
  begin
    CanClose := ConfirmMsg( '確認結束程式 ?' );
    if ( CanClose ) then
    begin
      FMsgSubject.Normal( '停止所有運作執行緒中....' );
      Delay( COMM_DELAY * 5 );
      DestroyAllThread;
      FMsgSubject.Normal( '準備儲存異動資料....' );
      Delay( COMM_DELAY * 5 );
      LogModule.SaveToFile;
      FMsgSubject.OK( '程式運作已停止。' );
      Delay( COMM_DELAY * 5 );
    end;  
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormDestroy(Sender: TObject);
begin
  ConfigModule.Free;
  LogModule.Free;
  FParseObServer.Free;
  FMsgObServer.Free;
  FMsgSubject.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.OnMsgUpdate(aSubject: TSubject);

  { ---------------------------------------------- }

  function GetMsgImageIndex(const aMsgType: Integer): Integer;
  begin
    case aMsgType of
      MB_OK: Result := 1;
      MB_ICONERROR: Result := 2;
      MB_ICONWARNING: Result := 3;
    else
      Result := 0;
    end;
  end;

  { ---------------------------------------------- }

var
  aMsgSubject: TMessageSubject;
  aItem: TListItem;
begin
  aMsgSubject := TMessageSubject( aSubject );
  aItem := MsgList.Items.Add;
  aItem.ImageIndex := GetMsgImageIndex( aMsgSubject.MessageType );
  aItem.Caption := Format( '%s > %s', [
    FormatDateTime( 'yyyy-mm-dd hh:nn:ss', Now ), aMsgSubject.MessageText] );
  if ( FAutoScroll ) then
  begin
    aItem.Selected := True;
    aItem.Focused := True;
    aItem.MakeVisible( True );
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.OnXmlUpdate(aSubject: TSubject);
var
  aObj: PActObject;
begin
  if not Assigned( TParseSubject( aSubject ).ActionObject ) then Exit;
  aObj := TParseSubject( aSubject ).ActionObject;
  if ( aObj.KeyId = EmptyStr ) then
  begin
    LogModule.ActionDataSet.Append;
    LogModule.ActionDataSet.FieldByName( 'actdate' ).AsString := aObj.ActDate;
    LogModule.ActionDataSet.FieldByName( 'acttime' ).AsString := aObj.ActTime;
    LogModule.ActionDataSet.FieldByName( 'actfilename' ).AsString := aObj.ActFileName;
    LogModule.ActionDataSet.FieldByName( 'actfilesize' ).AsString := aObj.ActFileSize;
    LogModule.ActionDataSet.FieldByName( 'actfilepath' ).AsString:= aObj.ActFilePath;
    LogModule.ActionDataSet.FieldByName( 'actsource' ).AsString := aObj.ActSource;
    LogModule.ActionDataSet.FieldByName( 'actstatus' ).AsString := aObj.ActStatus;
    LogModule.ActionDataSet.FieldByName( 'actcost' ).AsInteger := aObj.ActCost;
    LogModule.ActionDataSet.FieldByName( 'actprogress' ).AsInteger := aObj.ActProgress;
    LogModule.ActionDataSet.FieldByName( 'acterrmsg' ).AsString := aObj.ActErrMsg;
    LogModule.ActionDataSet.FieldByName( 'actflag' ).AsString :=  '1';
    LogModule.ActionDataSet.Post;
    aObj.KeyId := LogModule.ActionDataSet.FieldByName( 'keyid' ).AsString;
  end else
  begin
    if LogModule.ActionDataSet.Locate( 'keyid', aObj.KeyId, [] ) then
    begin
      LogModule.ActionDataSet.Edit;
      LogModule.ActionDataSet.FieldByName( 'actstatus' ).AsString := aObj.ActStatus;
      LogModule.ActionDataSet.FieldByName( 'actprogress' ).AsInteger := aObj.ActProgress;
      LogModule.ActionDataSet.FieldByName( 'actcost' ).AsInteger := aObj.ActCost;
      LogModule.ActionDataSet.FieldByName( 'acterrmsg' ).AsString := aObj.ActErrMsg;
      LogModule.ActionDataSet.Post;
      if FAutoScroll then
        ActionGridTableView.DataController.LocateByKey( aObj.KeyId );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.MainPageChange(Sender: TObject);
begin
  lblTitle.Caption := MainPage.ActivePage.Caption;
  StyleModule.PageImageList.GetIcon( MainPage.ActivePageIndex,
    PageMarkImage.Picture.Icon );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxNavBarLinkClick(Sender: TObject;
  ALink: TdxNavBarItemLink);
begin
  if ( ALink.Index in [0..MainPage.PageCount-1] ) then
    MainPage.ActivePageIndex := ALink.Index;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.InitialUIState(const aEnable: Boolean);
var
  aIndex: Integer;
begin
  for aIndex := 0 to BarActionList.ActionCount - 1 do
    TAction( BarActionList.Actions[aIndex] ).Enabled := aEnable;
  dxOption.Enabled := aEnable ;
  if ( aEnable ) then
  begin
    if ( FRunSate = rsStop ) then
    begin
      actPlay.Enabled := ( MainPage.ActivePageIndex = 0 );
      actStop.Enabled := False;
    end;
    if ( FRunSate = rsPlay ) then
    begin
      actStop.Enabled := ( MainPage.ActivePageIndex = 0 );
      actPlay.Enabled := False;
    end;
    actAutoScroll.Enabled := ( MainPage.ActivePageIndex = 0 );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.ChangeDevexComponentLang;
begin
  cxSetResourceString( @scxGridNoDataInfoText, '無異動記錄可顯示' );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.InitialPageControl;
begin
  MainPage.HideTabs := True;
  MainPage.ActivePageIndex := 0;
  MainPageChange( MainPage );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.StartupTimerTimer(Sender: TObject);
var
  aLoaded: Boolean;
begin
  StartupTimer.Enabled := False;
  Screen.Cursor := crHourGlass;
  try
    InitialUIState( False );
    Delay( COMM_DELAY * 3 );
    FMsgSubject.Normal( '程式起始中........' );
    Delay( COMM_DELAY * 3 );
    aLoaded := ConfigModule.LoadFromFile;
    if ( aLoaded ) then aLoaded := LogModule.LoadFromFile;
    if ( aLoaded ) then
    begin
      ActionGridTableView.Controller.GoToLast( False );
      Delay( COMM_DELAY * 3 );
      FMsgSubject.Normal( '程式可開始執行。' );
      InitialUIState( True );
    end else
    begin
      Delay( COMM_DELAY * 3 );
      FMsgSubject.Warning( '程式無法正常運作。' );
      InitialUIState( False );
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.actPlayExecute(Sender: TObject);
begin
  if CreateAllThread then
  begin
    actPlay.Enabled := False;
    actStop.Enabled := True;
    dxOption.Enabled := False;
    FRunSate := rsPlay;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.actStopExecute(Sender: TObject);
begin
  DestroyAllThread;
  Sleep( COMM_DELAY * 3 );
  LogModule.SaveToFile;
  Sleep( COMM_DELAY * 3 );
  FMsgSubject.OK( '程式運作已停止。' );
  Sleep( COMM_DELAY * 3 );
  actPlay.Enabled := True;
  actStop.Enabled := False;
  dxOption.Enabled := True;
  FRunSate := rsStop;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.actAutoScrollExecute(Sender: TObject);
begin
  FAutoScroll := dxAutoScroll.Down;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxActionClick(Sender: TObject);
begin
  MainPage.ActivePageIndex := dxAction.Index;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxListClick(Sender: TObject);
begin
  MainPage.ActivePageIndex := dxList.Index;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxOptionClick(Sender: TObject);
begin
  fmOption := TfmOption.Create( nil );
  try
    if ( fmOption.ShowModal = mrOk ) then
    begin
      Screen.Cursor := crHourGlass;
      try
        Delay( COMM_DELAY * 5 );
        if not ConfigModule.SaveToFile then
          ConfigModule.LoadFromFile( True );
        if not LogModule.SaveToFile then
          LogModule.LoadFromFile( True );  
      finally
        Screen.Cursor := crDefault;
      end;
    end;
  finally
    fmOption.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.CreateAllThread: Boolean;
begin
  Result := CreateParseThread;
  if not Result then Exit;
  Result := CreateWatchThread;
  if not Result then Exit;
  Delay( COMM_DELAY * 3 );
  PostMessage( TWatchThread( FWatchThread ).MsgHandle, WM_PLAY, 0, 0 );
  Delay( COMM_DELAY * 3 );
  PostMessage( TParseThread( FParseThread ).MsgHandle, WM_PLAY, 0, 0 );
  FMsgSubject.OK( '程式已開始運作。' );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.DestroyAllThread;
begin
  if Assigned( FWatchThread ) then
  begin
    FMsgSubject.Normal( '停止【檔案監控】執行緒中....' );
    Delay( COMM_DELAY * 3 );
    DestroyWatchThread;
    FMsgSubject.Normal( '【檔案監控】已停止。' );
    Delay( COMM_DELAY * 3 );
  end;
  if Assigned( FParseThread ) then
  begin
    FMsgSubject.Normal( '停止【解析核心】執行緒中....' );
    Delay( COMM_DELAY * 3 );
    DestroyParseThread;
    FMsgSubject.Normal( '【解析核心】已停止。' );
    Delay( COMM_DELAY * 3 );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.CreateWatchThread: Boolean;
begin
  Result := False;
  FMsgSubject.Normal( '準備【檔案監控】執行緒中....' );
  Delay( COMM_DELAY * 5 );
  FWatchThread := TWatchThread.Create;
  try
    TWatchThread( FWatchThread ).Param := ConfigModule.Param;
    TWatchThread( FWatchThread ).Perido := ConfigModule.Period;
    TWatchThread( FWatchThread ).MessageSubject.AddObServer( FMsgObServer );
    TWatchThread( FWatchThread ).Resume;
    Delay( COMM_DELAY * 3 );
    Result := True;
    FMsgSubject.OK( '【檔案監控】執行緒就序。' );
    Delay( COMM_DELAY * 3 );
  except
    on E: Exception do
    begin
      FMsgSubject.Error( Format( '建立【檔案監控】執行緒錯誤, 原因:%s。',
        [E.Message] ) );
      Delay( COMM_DELAY * 3 );
    end;
  end;
  if not Result then DestroyWatchThread;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.DestroyWatchThread;
begin
  PostMessage( TWatchThread( FWatchThread ).MsgHandle, WM_STOP, 0, 0 );
  Delay( COMM_DELAY * 5 );
  try
    TWatchThread( FWatchThread ).Terminate;
    TWatchThread( FWatchThread ).WaitFor;
    TWatchThread( FWatchThread ).Free;
  except
    { ... }
  end;
  FWatchThread := nil;
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.CreateParseThread: Boolean;
begin
  Result := False;
  FMsgSubject.Normal( '準備【解析核心】執行緒中....' );
  Delay( COMM_DELAY * 3 );
  FParseThread := TParseThread.Create;
  try
    TParseThread( FParseThread ).Param := ConfigModule.Param;
    TParseThread( FParseThread ).DataBase := ConfigModule.Database;
    TParseThread( FParseThread ).MessageSubject.AddObServer( FMsgObServer );
    TParseThread( FParseThread ).ParseSubject.AddObServer( FParseObServer );
    TParseThread( FParseThread ).Resume;
    Delay( COMM_DELAY * 3 );
    Result := True;
    FMsgSubject.OK( '【解析核心】執行緒就序。' );
    Delay( COMM_DELAY * 3 );
  except
    on E: Exception do
    begin
      FMsgSubject.Error( Format( '建立【解析核心】執行緒錯誤, 原因:%s。',
        [E.Message] ) );
      Delay( COMM_DELAY * 3 );
    end;
  end;
  if not Result then DestroyParseThread;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.DestroyParseThread;
begin
  PostMessage( TParseThread( FParseThread ).MsgHandle, WM_STOP, 0, 0 );
  Delay( COMM_DELAY * 5 );
  try
    TParseThread( FParseThread ).Terminate;
    TParseThread( FParseThread ).WaitFor;
    TParseThread( FParseThread ).Free;
  except
    { ... }
  end;
  FParseThread := nil;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.ActionGridTableViewCustomDrawCell(
  Sender: TcxCustomGridTableView; ACanvas: TcxCanvas;
  AViewInfo: TcxGridTableDataCellViewInfo; var ADone: Boolean);
var
  aIconLeft, aIconTop, aFlag, aOffsetX: Integer;
  aTextRect: TRect;

  { ----------------------------------------------- }

  function GetIcon: Integer;
  var
    aSource: String;
  begin
    Result := 5;
    aSource := VarToStrDef( AViewInfo.GridRecord.Values[ColumnActSource.Index],
      EmptyStr );
    if ( aSource = 'Merge' ) then Result := 10;
  end;

  { ----------------------------------------------- }

begin
  ACanvas.FillRect( AViewInfo.Bounds );
  if AViewInfo.Item = ColumnActFileName then
  begin
    aIconLeft := AViewInfo.Bounds.Left + 3;
    aIconTop := AViewInfo.Bounds.Top;
    ACanvas.DrawImage( StyleModule.ActionImageList, aIconLeft,
      aIconTop, GetIcon );
    aFlag := ( cxSingleLine	or cxShowEndEllipsis or cxAlignLeft	or cxAlignVCenter );
    aTextRect := AViewInfo.TextAreaBounds;
    aOffsetX := StyleModule.ActionImageList.Width + 8;
    InflateRect( aTextRect, ( 0 - aOffsetX ), 0 );
    ACanvas.DrawText( AViewInfo.DisplayValue, aTextRect, aFlag );
    ADone := True;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.ColumnActCostGetDisplayText(
  Sender: TcxCustomGridTableItem; ARecord: TcxCustomGridRecord;
  var AText: String);
var
  aSec: Integer;
begin
  if TryStrToInt( AText, aSec ) then
  begin
    if ( aSec > 60 ) then
      AText := Format( '%d 分 %d 秒', [(aSec div 60), ( aSec mod 60)] )
    else
      AText := Format( '%d 秒', [aSec] );
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
