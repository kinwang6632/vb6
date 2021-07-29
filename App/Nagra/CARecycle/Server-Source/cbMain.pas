unit cbMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, DBClient, cbStyleModule, IniFiles, cbAppclass, cbClass,
  cxGraphics, cxCustomData, cxStyles, cxTL, cxTextEdit,
  cxImageComboBox, ComCtrls, cxContainer, cxListView, cxInplaceContainer,
  cxControls, dxDockPanel, dxDockControl, dxBar, dxBarExtItems, cxFilter,
  cxData, cxDataStorage, cxEdit, DB, cxDBData, cxGridCustomView,
  cxGridCustomTableView, cxGridTableView, cxGridBandedTableView,
  cxGridDBBandedTableView, cxClasses, cxGridLevel, cxGrid, cxGridStrs,
  cxCheckBox;

type
  TfmMain = class(TForm)
    dxBarManager: TdxBarManager;
    dxMenuItem1: TdxBarSubItem;
    dxMenuItem2: TdxBarSubItem;
    dxPlay: TdxBarButton;
    dxStop: TdxBarButton;
    dxAutoScroll: TdxBarButton;
    dxMenuItem11: TdxBarButton;
    dxMenuItem3: TdxBarButton;
    dxMenuItem22: TdxBarButton;
    dxMenuItem21: TdxBarButton;
    PlayCmdGroup: TdxBarGroup;
    StopCmdGroup: TdxBarGroup;
    ControlSendGroup: TdxBarGroup;
    dxDockingManager: TdxDockingManager;
    dxDockHost: TdxDockSite;
    dxLayoutDockSite2: TdxLayoutDockSite;
    dxLayoutDockSite3: TdxLayoutDockSite;
    dxLayoutDockSite1: TdxLayoutDockSite;
    dxDockPanel2: TdxDockPanel;
    dxDockPanel5: TdxDockPanel;
    MsgList: TcxListView;
    dxDockPanel1: TdxDockPanel;
    SoTree: TcxTreeList;
    SoTreeColumn1: TcxTreeListColumn;
    SoTreeColumn2: TcxTreeListColumn;
    LoadConfigTimer: TTimer;
    glRecycle: TcxGridLevel;
    RecycleGrid: TcxGrid;
    cdsRecycle: TClientDataSet;
    dsRecycle: TDataSource;
    gvRecycle: TcxGridDBBandedTableView;
    gvRecycleColumn1: TcxGridDBBandedColumn;
    gvRecycleColumn2: TcxGridDBBandedColumn;
    gvRecycleColumn3: TcxGridDBBandedColumn;
    gvRecycleColumn4: TcxGridDBBandedColumn;
    gvRecycleColumn5: TcxGridDBBandedColumn;
    gvRecycleColumn6: TcxGridDBBandedColumn;
    gvRecycleColumn7: TcxGridDBBandedColumn;
    gvRecycleColumn8: TcxGridDBBandedColumn;
    gvRecycleColumn9: TcxGridDBBandedColumn;
    gvRecycleColumn10: TcxGridDBBandedColumn;
    gvRecycleColumn11: TcxGridDBBandedColumn;
    gvRecycleColumn12: TcxGridDBBandedColumn;
    gvRecycleColumn13: TcxGridDBBandedColumn;
    gvRecycleColumn14: TcxGridDBBandedColumn;
    gvRecycleColumn15: TcxGridDBBandedColumn;
    gvRecycleColumn16: TcxGridDBBandedColumn;
    gvRecycleColumn17: TcxGridDBBandedColumn;
    gvRecycleColumn18: TcxGridDBBandedColumn;
    gvRecycleColumn19: TcxGridDBBandedColumn;
    gvRecycleColumn20: TcxGridDBBandedColumn;
    gvRecycleColumn21: TcxGridDBBandedColumn;
    gvRecycleColumn22: TcxGridDBBandedColumn;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure OnMsgUpdate(aSubject: TSubject);
    procedure OnCmdUpdate(aSubject: TSubject);
    procedure OnDbUpdate(aSubject: TSubject);
    procedure LoadConfigTimerTimer(Sender: TObject);
    procedure dxMenuItem11Click(Sender: TObject);
    procedure dxPlayClick(Sender: TObject);
    procedure dxStopClick(Sender: TObject);
    procedure gvRecycleCustomDrawCell(Sender: TcxCustomGridTableView;
      ACanvas: TcxCanvas; AViewInfo: TcxGridTableDataCellViewInfo;
      var ADone: Boolean);
  private
    { Private declarations }
    FAppIni: TIniFile;
    FConfigFileName: String;
    FMsgSubject: TMessageSubject;
    FMsgObserver: TObServer;
    FCmdObserver: TObServer;
    FDbObServer: TObServer;
    FSoDbThread: TThread;
    FThreadActivate: Boolean;
    FCmdHelper: TCommandHelper;
    procedure ChangeDevexComponentLang;
    procedure CleanupResource;
    procedure ClearSoDbTree;
    procedure BuildSoDbTree;
    function ThreadStart: Boolean;
    procedure ThreadStop;
    procedure PrepareDisplayDataSet;
    procedure UnPrepareDisplayDataSet;
    procedure LoadAppIni;
    procedure SaveAppIni;
  public
    { Public declarations }
  end;

var
  fmMain: TfmMain;

implementation

uses Math, cbUtilis, cbConfigModule, cbSoDbThread;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

function GetMsgImageIndex(const aType: Longint): Integer;
begin
  case aType of
    MB_ICONINFORMATION: Result := 0;
    MB_OK: Result := 1;
    MB_ICONWARNING: Result := 2;
    MB_ICONERROR: Result := 3;
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

procedure TfmMain.FormCreate(Sender: TObject);
var
  aFileName: String;
begin
  aFileName := IncludeTrailingPathDelimiter( ExtractFilePath( ParamStr( 0 ) ) ) +
    'AppInfo.ini';
  FAppIni := TIniFile.Create( aFileName );
  LoadAppIni;
  FMsgObserver := TObServer.Create;
  FMsgObserver.OnUpdate := Self.OnMsgUpdate;
  FMsgSubject := TMessageSubject.Create;
  FMsgSubject.AddObServer( FMsgObserver );
  FCmdObserver := TObServer.Create;
  FCmdObserver.OnUpdate := Self.OnCmdUpdate;
  FDbObServer := TObServer.Create;
  FDbObServer.OnUpdate := Self.OnDbUpdate;
  FCmdHelper := TCommandHelper.Create;
  ConfigModule := TConfigModule.Create( nil );
  ConfigModule.MsgSubject.AddObServer( FMsgObserver );
  FThreadActivate := False;
  ChangeDevexComponentLang;
  PrepareDisplayDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormShow(Sender: TObject);
begin
  FMsgSubject.Normal( '�{���Ұʤ�....' );
  LoadConfigTimer.Enabled := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  if FThreadActivate then CanClose := ConfirmMsg( '�T�{�����{�� ?' );
  if CanClose then
  begin
    dxStop.Click;
    CleanupResource;
    UnPrepareDisplayDataSet;
    SaveAppIni;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.OnMsgUpdate(aSubject: TSubject);
var
  aText: String;
  aItem: TListItem;
begin
  aText := Format( ' %s >   %s', [FormatDateTime( 'yyyy-mm-dd hh:mm', Now ),
    TMessageSubject( aSubject ).MessageText] );
  aItem := MsgList.Items.Add;
  aItem.Caption := aText;
  aItem.ImageIndex := GetMsgImageIndex( TMessageSubject( aSubject ).MessageType );
  aItem.MakeVisible( True );
  Application.ProcessMessages;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.OnCmdUpdate(aSubject: TSubject);
var
  aCmd: TRecycleCommand;
begin
  aCmd := TRecycleCommand( TObjectSubject( aSubject ).NotifyObject );
  if not Assigned( aCmd ) then Exit;
  {}
  FCmdHelper.RecycleCommand := aCmd;
  {}
  if cdsRecycle.Locate( 'SEQNO;ICCNO', VarArrayOf( [aCmd.SeqNo, aCmd.IccNo] ), [] ) then
    cdsRecycle.Edit
  else
    cdsRecycle.Append;
  {}
  cdsRecycle.FieldByName( 'SEQNO' ).AsString := aCmd.SeqNo;
  cdsRecycle.FieldByName( 'ICCNO' ).AsString := aCmd.IccNo;
  cdsRecycle.FieldByName( 'STBNO' ).AsString := aCmd.StbNo;
  cdsRecycle.FieldByName( 'STB_FLAG' ).AsInteger := aCmd.StbFlag;
  cdsRecycle.FieldByName( 'STBAUTOCB' ).AsInteger := aCmd.StbAutoFlag;
  cdsRecycle.FieldByName( 'COMPCODE' ).AsInteger := aCmd.CompCode;
  cdsRecycle.FieldByName( 'COMPNAME' ).AsString := aCmd.CompName;
  cdsRecycle.FieldByName( 'RCSTARTDATE' ).AsString := aCmd.RCStartDate;
  cdsRecycle.FieldByName( 'RCENDDATE' ).AsString := aCmd.RCEndDate;
  cdsRecycle.FieldByName( 'OPERATOR' ).AsString := aCmd.Operator;
  cdsRecycle.FieldByName( 'UPDTIME' ).AsDateTime := aCmd.UpdTime;
  cdsRecycle.FieldByName( 'LASTDOWNLOADTIME' ).AsString := aCmd.LastDownloadTime;
  cdsRecycle.FieldByName( 'CMDFLAG' ).AsInteger := aCmd.CMDFlag;
  cdsRecycle.FieldByName( 'RECYCLETEXT' ).AsString := aCmd.RecycleText;
  cdsRecycle.FieldByName( 'TRANSNUM' ).AsString := aCmd.TransNum;
  cdsRecycle.FieldByName( 'CONFIRM' ).AsString := aCmd.Confirm;
  cdsRecycle.FieldByName( 'ERRMSG' ).AsString := aCmd.ErrMsg;
  { �����^�� }
  if ( FCmdHelper.RecycleCommand.StbAutoFlag = 0 ) then
  begin
    cdsRecycle.FieldByName( 'CMD52_1' ).AsString := FCmdHelper.FindTextByCmd( '52_1' );
    cdsRecycle.FieldByName( 'CMD8' ).AsString := FCmdHelper.FindTextByCmd( '8' );
    cdsRecycle.FieldByName( 'CMD97_96' ).AsString := FCmdHelper.FindTextByCmd( '97_96' );
    cdsRecycle.FieldByName( 'CMD7' ).AsString := FCmdHelper.FindTextByCmd( '7' );
    cdsRecycle.FieldByName( 'CMD48' ).AsString := FCmdHelper.FindTextByCmd( '48' );
    cdsRecycle.FieldByName( 'CMD52_2' ).AsString := FCmdHelper.FindTextByCmd( '52_2' );
  end else
  { ���n�^�� }
  begin
    cdsRecycle.FieldByName( 'CALLCMD52_1' ).AsString := FCmdHelper.FindTextByCmd( '52_1' );
    cdsRecycle.FieldByName( 'CALLCMD62' ).AsString := FCmdHelper.FindTextByCmd( '62' );
    cdsRecycle.FieldByName( 'CALLCMD8' ).AsString := FCmdHelper.FindTextByCmd( '8' );
    cdsRecycle.FieldByName( 'CALLCMD60' ).AsString := FCmdHelper.FindTextByCmd( '60' );
    cdsRecycle.FieldByName( 'CALLCMD7' ).AsString := FCmdHelper.FindTextByCmd( '7' );
    cdsRecycle.FieldByName( 'CALLCMD48' ).AsString := FCmdHelper.FindTextByCmd( '48' );
    cdsRecycle.FieldByName( 'CALLCMD99' ).AsString := FCmdHelper.FindTextByCmd( '99' );
    cdsRecycle.FieldByName( 'CALLCMD52_2' ).AsString := FCmdHelper.FindTextByCmd( '52_2' );
  end;
  cdsRecycle.Post;
  aCmd.Data := cdsRecycle.ActiveBuffer;
  Application.ProcessMessages;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.OnDbUpdate(aSubject: TSubject);
var
  aSo: TAppSo;
  aNode: TcxTreeListNode;
begin
  aSo := TAppSo( TObjectSubject( aSubject ).NotifyObject );
  if not Assigned( aSo ) then Exit;
  if not Assigned( aSo.DisplayNode ) then Exit;
  aNode := TcxTreeListNode( aSo.DisplayNode );
  aNode.ImageIndex := GetSoImageIndex( aSo.DbConnectStatus );
  aNode.SelectedIndex := aNode.ImageIndex;
  aNode.StateIndex := GetSoStateIndex( aSo.DbConnectStatus );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.ChangeDevexComponentLang;
begin
  cxSetResourceString( @scxGridNoDataInfoText, '�L��ƥi���' );
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.CleanupResource;
begin
  FMsgSubject.RemoveObServer;
  FMsgSubject.Free;
  FMsgObserver.OnUpdate := nil;
  FMsgObserver.Free;
  FCmdObserver.OnUpdate := nil;
  FCmdObserver.Free;
  FDbObServer.OnUpdate := nil;
  FDbObServer.Free;
  FCmdHelper.Free;
  ConfigModule.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.LoadConfigTimerTimer(Sender: TObject);
begin
  LoadConfigTimer.Enabled := False;
  Screen.Cursor := crHourGlass;
  try
    FMsgSubject.Normal( 'Ū���]�w�ɤ�...�C' );
    Delay( 500 );
    if not ConfigModule.LoadFromFile( FConfigFileName ) then
    begin
      dxPlay.Enabled := False;
      FMsgSubject.Error( '�t�εL�k����C' );
      Exit;
    end;
    ClearSoDbTree;
    FMsgSubject.Normal( '�إߨt�Υx��...�C' );
    Delay( 500 );
    BuildSoDbTree;
    FMsgSubject.OK( '�]�w��Ū�������C' );
    Delay( 500 );
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ --------------------------------------------------------------------------- }

procedure TfmMain.ClearSoDbTree;
begin
  SoTree.BeginUpdate;
  try
    SoTree.Clear;
  finally
    SoTree.EndUpdate;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.BuildSoDbTree;

  { -------------------------------------------------- }

  function GetPosNode(const aPosName: String; const aCanAdd: Boolean = True):
    TcxTreeListNode;
  begin
    Result := nil;
    if ( aPosName = EmptyStr ) then Exit;
    Result := SoTree.FindNodeByText( aPosName, SoTreeColumn2 );
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
  aIndex, aMaxCompCode: Integer;
  aSo: TAppSo;
  aNodePos, aNodeComp: TcxTreeListNode;
begin
  aMaxCompCode := 0;
  for aIndex := 0 to ConfigModule.SoList.Count - 1 do
  begin
    aSo := ConfigModule.SoList[aIndex];
    aMaxCompCode := Max( StrToInt( aSo.CompCode ), aMaxCompCode );
    { �����x�`�I, ���t�Υx�n���b���@�ӥ��x�U }
    aNodePos := GetPosNode( aSo.PosName );
    aNodeComp := SoTree.AddChild( aNodePos, aSO );
    aSo.DisplayNode := aNodeComp;
    aNodeComp.AssignValues( [aSo.CompCode, aSo.CompName, EmptyStr] );
    aNodeComp.ImageIndex := GetSoImageIndex( aSo.DbConnectStatus );
    aNodeComp.StateIndex := GetSoStateIndex( aSo.DbConnectStatus );
    aNodeComp.SelectedIndex := aNodeComp.ImageIndex;
    aNodePos.Expanded := True;
    FMsgSubject.Normal( Format( '�إߨt�Υx[%s]�C', [aSo.CompName] ) );
    Delay( 100 );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxMenuItem11Click(Sender: TObject);
begin
  if ConfigModule.ShowConfigForm = mrOk then
  begin
    LoadConfigTimer.Interval := 500;
    LoadConfigTimer.Enabled := True;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfmMain.ThreadStart: Boolean;
begin
  Result := True;
  try
    FSoDbThread := TSoDbThread.Create( True, ConfigModule.SoList );
    TSoDbThread( FSoDbThread ).MsgSubject.AddObServer( FMsgObserver );
    TSoDbThread( FSoDbThread ).DbSubject.AddObServer( FDbObServer );
    TSoDbThread( FSoDbThread ).CmdSubject.AddObServer( FCmdObserver );
    FSoDbThread.Resume;
    Delay( 3000 );
    PostMessage( TSoDbThread( FSoDbThread ).MsgHandle, WM_PLAY, 0, 0 );
  except
    on E: Exception do
    begin
      FMsgSubject.Error( Format(
        '�إߨt�Υx��Ʈw��������~, �T��:%s�C', [E.Message] ) );
      Result := False;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.ThreadStop;
begin
  if Assigned( FSoDbThread ) then
  begin
    try
      PostMessage( TSoDbThread( FSoDbThread ).MsgHandle, WM_STOP, 0, 0 );
      FSoDbThread.Terminate;
      FSoDbThread.WaitFor;
      FSoDbThread.Free;
    except
      on E: Exception do FMsgSubject.Error( Format(
        '����t�Υx��Ʈw��������~, �T��:%s�C', [E.Message] ) );
    end;
  end;
  FSoDbThread := nil;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxPlayClick(Sender: TObject);
begin
  FMsgSubject.Normal( '�t�Υx��Ʈw������ǳƤ�...�C' );
  Screen.Cursor := crHourGlass;
  try
    Delay( 500 );
    if ThreadStart then
    begin
      FMsgSubject.Normal( '�t�Υx��Ʈw�}�l����C' );
      FThreadActivate := True;
      dxStop.Enabled := FThreadActivate;
      dxPlay.Enabled := not FThreadActivate;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.dxStopClick(Sender: TObject);
begin
  FMsgSubject.Normal( '�t�Υx��Ʈw���...�C' );
  Screen.Cursor := crHourGlass;
  try
    Delay( 500 );
    ThreadStop;
    FMsgSubject.Normal( '�t�Υx��Ʈw������w����C' );
    Delay( 500 );
    FThreadActivate := False;
    dxStop.Enabled := FThreadActivate;
    dxPlay.Enabled := not FThreadActivate;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.PrepareDisplayDataSet;
begin
  if VarIsNull( cdsRecycle.Data ) then
    cdsRecycle.CreateDataSet;
  cdsRecycle.AddIndex( 'IDX1', 'SEQNO;ICCNO', [ixPrimary] );  
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.UnPrepareDisplayDataSet;
begin
  if not VarIsNull( cdsRecycle.Data ) then
  begin
    cdsRecycle.EmptyDataSet;
    cdsRecycle.DeleteIndex( 'IDX1' );
  end;
  cdsRecycle.Data := Null;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.gvRecycleCustomDrawCell(Sender: TcxCustomGridTableView;
  ACanvas: TcxCanvas; AViewInfo: TcxGridTableDataCellViewInfo;
  var ADone: Boolean);

  { ------------------------------------------ }

  function GetDrawImageIndex(const aText: String): Integer;
  begin
    if ( Pos( '2', aText ) > 0 ) then //����
      Result := 3
    else if ( Pos( '0', aText ) > 0 ) then //�|���B�z
      Result :=  5
    else if ( Pos( '3', aText ) > 0 ) then //�B�z��
      Result := 6
    else
      Result := 1;
  end;

  { ------------------------------------------ }

var
  aValue: String;
  aX, aY: Integer;
begin
  if ( AViewInfo.Item.Tag = 1 ) then
  begin
    aValue := VarToStrDef( AViewInfo.GridRecord.Values[AViewInfo.Item.Index], EmptyStr );
    if ( aValue <> EmptyStr ) then
    begin
      ACanvas.FillRect( AViewInfo.Bounds );
      aX := ( ( AViewInfo.Bounds.Right - AViewInfo.Bounds.Left ) -
        StyleModule.MsgImageList.Width ) div 2;
      aY := ( ( AViewInfo.Bounds.Bottom - AViewInfo.Bounds.Top ) -
        StyleModule.MsgImageList.Height ) div 2;
      ACanvas.DrawImage( StyleModule.MsgImageList, AViewInfo.Bounds.Left + aX ,
        AViewInfo.Bounds.Top + aY, GetDrawImageIndex( aValue ) );
      ADone := True;
    end;  
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.LoadAppIni;
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
  FConfigFileName := FAppIni.ReadString( 'ConfigFile', 'ActiveFile', EmptyStr );
  if ( FConfigFileName = EmptyStr ) then
    FConfigFileName := IncludeTrailingPathDelimiter( ExtractFilePath( ParamStr( 0 ) ) ) + 'Config.cfg';
end;

{ ---------------------------------------------------------------------------- }

procedure TfmMain.SaveAppIni;
begin
  FAppIni.WriteInteger( 'WindowPositon', 'State', Ord( Self.WindowState ) );
  if ( Self.WindowState = wsNormal ) then
  begin
    FAppIni.WriteInteger( 'WindowPositon', 'Width', Self.Width );
    FAppIni.WriteInteger( 'WindowPositon', 'Height', Self.Height );
    FAppIni.WriteInteger( 'WindowPositon', 'Left', Self.Left );
    FAppIni.WriteInteger( 'WindowPositon', 'Top', Self.Top );
  end;
  FAppIni.WriteString( 'ConfigFile', 'ActiveFile', FConfigFileName );
  FAppIni.UpdateFile;
end;

{ ---------------------------------------------------------------------------- }

end.
