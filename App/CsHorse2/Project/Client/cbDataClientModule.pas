unit cbDataClientModule;

interface

uses
  SysUtils, Classes, Controls, Variants, Forms, DB, 
{$IFDEF APPDEBUG}
  CodeSiteLogging,
{$ENDIF}
{ App: }
  cbClientClass, cbHrHelper, cbDesignPattern, cbStyleModule,
{ Developer Express: }
  cxCustomData, cxTL,
{ ODAC: }
  MemDS, VirtualTable,
{ RemObject SDK: }
  CsHorse2Library_Intf;

type
  TDataClientModule = class(TDataModule)
    dsSo021: TDataSource;
    BfSo021: TVirtualTable;
    dsCd042: TDataSource;
    BfCd042: TVirtualTable;
    dsSo022: TDataSource;
    BfSo022: TVirtualTable;
    dsSo023: TDataSource;
    BfSo023: TVirtualTable;
    BfUserList: TVirtualTable;
    dsUserList: TDataSource;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
    procedure OnSo021Update(ASubject: TSubject);
    procedure OnSo022Update(ASubject: TSubject);
    procedure OnSo023Update(ASubject: TSubject);
    procedure OnCd042Update(ASubject: TSubject);
    procedure OnUserListUpdate(ASubject: TSubject);
    procedure OnMsgListUpdate(ASubject: TSubject);
    procedure RefreshTrayInfoText(const AIsMsgArrival: Boolean = False);
  private
    { Private declarations }
    FCompCode: String;
    FUserId: String;
    FSo021Obsrv: TObServer;
    FCd042Obsrv: TObServer;
    FSo022Obsrv: TObServer;
    FSo023Obsrv: TObServer;
    FUserListObsrv: TObServer;
    FMsgListObsrv: TObServer;
    FUserDisplayType: TUserListDisplayType;
    procedure InitBuffers;
    procedure ApplyBufferFilter(ABuffer: TVirtualTable; const AFilter: String);
  public
    { Public declarations }
    procedure SwitchCompCode(const ACompCode: String);
    procedure SwitchUserDisplayType(const ADisplayType: TUserListDisplayType);
    procedure UpdateMsgItem(const ANode: TcxTreeListNode);
    procedure DeleteMsgItem(const ANode: TcxTreeListNode);
    procedure AddForm(AForm: TForm);
    procedure RemoveForm(AForm: TForm);
    procedure NotifyMsgForm(const AReadyToSend: Boolean);
    procedure ForceFormClose;
    function GetMsgForm(const AMsgInfo: TMsgInfo): TForm;
    property So021Obsrv: TObServer read FSo021Obsrv;
    property So022Obsrv: TObServer read FSo022Obsrv;
    property So023Obsrv: TObServer read FSo023Obsrv;
    property Cd042Obsrv: TObServer read FCd042Obsrv;
    property UserListObsrv: TObServer read FUserListObsrv;
    property MsgListObsrv: TObServer read FMsgListObsrv;
  end;

var
  DataClientModule: TDataClientModule;
  MsgFormList: TList;
  AnnFormList: TList;

implementation

uses cbUtilis, cbMain, cbSendMsg;


{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TDataClientModule.DataModuleCreate(Sender: TObject);
begin
  InitBuffers;
  FSo021Obsrv := TObServer.Create;
  FSo021Obsrv.OnUpdate := OnSo021Update;
  FSo022Obsrv := TObServer.Create;
  FSo022Obsrv.OnUpdate := OnSo022Update;
  FSo023Obsrv := TObServer.Create;
  FSo023Obsrv.OnUpdate := OnSo023Update;
  FCd042Obsrv := TObServer.Create;
  FCd042Obsrv.OnUpdate := OnCd042Update;
  FUserListObsrv := TObServer.Create;
  FUserListObsrv.OnUpdate := OnUserListUpdate;
  FMsgListObsrv := TObServer.Create;
  FMsgListObsrv.OnUpdate := OnMsgListUpdate;
  FCompCode := EmptyStr;
  FUserId := EmptyStr;
  FUserDisplayType := gkWorkClass;
  MsgFormList := TList.Create;
  AnnFormList := TList.Create;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataClientModule.DataModuleDestroy(Sender: TObject);
begin
  FSo021Obsrv.Free;
  FSo022Obsrv.Free;
  FSo023Obsrv.Free;
  FCd042Obsrv.Free;
  FUserListObsrv.Free;
  FMsgListObsrv.Free;
  MsgFormList.Free;
  AnnFormList.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataClientModule.InitBuffers;
begin
  TBufferHelper.CreateFieldDefs( biSO021, BfSo021 );
  TBufferHelper.CreateFieldDefs( biCD042, BfCd042 );
  TBufferHelper.CreateFieldDefs( biSO022, BfSo022 );
  TBufferHelper.CreateFieldDefs( biSO023, BfSo023 );
  TBufferHelper.CreateFieldDefs( biUserList, BfUserList );
end;

{ ---------------------------------------------------------------------------- }

procedure TDataClientModule.ApplyBufferFilter(ABuffer: TVirtualTable;
  const AFilter: String);
begin
  ABuffer.Filtered := False;
  ABuffer.Filter := EmptyStr;
  if ( AFilter <> EmptyStr ) then
  begin
    ABuffer.Filter := AFilter;
    ABuffer.Filtered := True;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataClientModule.OnSo021Update(ASubject: TSubject);
begin
  Screen.Cursor := crAppStart;
  try
    BfSo021.DisableControls;
    try
      BfSo021.Assign( TVirtualTable( TDataSubject( ASubject ).Data ) );
    finally
      BfSo021.EnableControls;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataClientModule.OnSo022Update(ASubject: TSubject);
begin
  Screen.Cursor := crAppStart;
  try
    BfSo022.DisableControls;
    try
      BfSo022.Assign( TVirtualTable( TDataSubject( ASubject ).Data ) );
    finally
      BfSo022.EnableControls;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataClientModule.OnSo023Update(ASubject: TSubject);
begin
  Screen.Cursor := crAppStart;
  try
    BfSo023.DisableControls;
    try
      BfSo023.Assign( TVirtualTable( TDataSubject( ASubject ).Data ) );
    finally
      BfSo023.EnableControls;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataClientModule.OnCd042Update(ASubject: TSubject);
begin
  Screen.Cursor := crAppStart;
  try
    BfCd042.DisableControls;
    try
      BfCd042.Assign( TVirtualTable( TDataSubject( ASubject ).Data ) );
    finally
      BfCd042.EnableControls;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataClientModule.OnUserListUpdate(ASubject: TSubject);
begin
  Screen.Cursor := crAppStart;
  try
    fmHrMain.UserTree.BeginUpdate;
    try
      BfUserList.DisableControls;
      try
        BfUserList.Assign( TVirtualTable( TDataSubject( ASubject ).Data ) );
      finally
        BfUserList.EnableControls;
      end;
    finally
      fmHrMain.UserTree.EndUpdate;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataClientModule.OnMsgListUpdate(ASubject: TSubject);
var
  aMsgInfo: TMsgInfo;
  aNode: TcxTreeListNode;
  aDisplayText1, aUniqueId: String;
begin
  aMsgInfo := TMsgInfo( TDataSubject( ASubject ).Data );
  aUniqueId := ( aMsgInfo.CompCode + aMsgInfo.MsgId );
  aNode := fmHrMain.MsgTree.FindNodeByText( aUniqueId, fmHrMain.MsgTreeCol3 );
  if not Assigned( aNode ) then
  begin
    fmHrMain.MsgTree.BeginUpdate;
    try
      aDisplayText1 := Format( '%s(%s) - %s', [aMsgInfo.MsgSenderName,
        aMsgInfo.MsgSenderWorkName, aMsgInfo.CompName] );
      aNode := fmHrMain.MsgTree.Add;
      aNode.Texts[fmHrMain.MsgTreeCol1.ItemIndex] := aDisplayText1;
      aNode.Texts[fmHrMain.MsgTreeCol2.ItemIndex] := aMsgInfo.MsgSubject;
      aNode.Texts[fmHrMain.MsgTreeCol3.ItemIndex] := aUniqueId;
      aNode.Texts[fmHrMain.MsgTreeCol4.ItemIndex] := aMsgInfo.MsgPriority;
      aNode.Texts[fmHrMain.MsgTreeCol5.ItemIndex] := aMsgInfo.MsgTime;
      aNode.ImageIndex := StyleModule.GetMsgItemImageIndex( aMsgInfo.IsRead );
      aNode.StateIndex := StyleModule.GetMsgItemsStateIndex( aMsgInfo.MsgPriority );
      aNode.SelectedIndex := aNode.ImageIndex;
      aNode.Data := aMsgInfo;
    finally
      fmHrMain.MsgTree.EndUpdate;
    end;
  end;
  RefreshTrayInfoText( True );
end;

{ ---------------------------------------------------------------------------- }

procedure TDataClientModule.SwitchCompCode(const ACompCode: String);
var
  aFilterText: String;
begin
  if ( FCompCode <> ACompCode ) then
  begin
    FCompCode := ACompCode;
    aFilterText :=  Format( 'COMPCODE=%s', [FCompCode] );
    ApplyBufferFilter( BfSo021, aFilterText );
    ApplyBufferFilter( BfCd042, aFilterText );
    ApplyBufferFilter( BfSo023, aFilterText );
    ApplyBufferFilter( BfSo022, aFilterText );
    aFilterText := Format( '(COMPCODE=%s) AND ((DISPLAYTYPE=%d)OR(DISPLAYTYPE=2))',
      [FCompCode, Ord( FUserDisplayType )] );
    ApplyBufferFilter( BfUserList, aFilterText );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataClientModule.SwitchUserDisplayType(const ADisplayType: TUserListDisplayType);
var
  aFilterText: String;
begin
  if ( FUserDisplayType <> ADisplayType ) then
  begin
    FUserDisplayType := ADisplayType;
    aFilterText := Format( '(COMPCODE=%s) AND ((DISPLAYTYPE=%d)OR(DISPLAYTYPE=2))',
      [FCompCode, Ord( FUserDisplayType )] );
    ApplyBufferFilter( BfUserList, aFilterText );
    fmHrMain.UserTree.ClearSorting;
    if ( FUserDisplayType = gkWorkClass ) then
    begin
      fmHrMain.UserTree.DataController.KeyField := 'UserId';
      fmHrMain.UserTree.DataController.ParentField := 'RGroupId';
      fmHrMain.UserTreeCol13.Position.BandIndex := 0;
      fmHrMain.UserTreeCol14.Position.BandIndex := 1;
      fmHrMain.UserTree.OptionsView.CategorizedColumn := fmHrMain.UserTreeCol13;
      fmHrMain.UserTreeCol9.SortOrder := soDescending;
    end else
    begin
      fmHrMain.UserTree.DataController.KeyField := 'UserId';
      fmHrMain.UserTree.DataController.ParentField := 'RStatus';
      fmHrMain.UserTreeCol13.Position.BandIndex := 1;
      fmHrMain.UserTreeCol14.Position.BandIndex := 0;
      fmHrMain.UserTree.OptionsView.CategorizedColumn := fmHrMain.UserTreeCol14;
      fmHrMain.UserTreeCol6.SortOrder := soAscending;
      fmHrMain.UserTreeCol9.SortOrder := soDescending;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataClientModule.RefreshTrayInfoText(const AIsMsgArrival: Boolean);
var
  aMsgInfo: TMsgInfo;
  aTrayInfoText: String;
  aIndex, aNormalProirity, aHighProirity, aUnRead: Integer;
begin
  aNormalProirity := 0;
  aHighProirity := 0;
  aUnRead := 0;
  aTrayInfoText := EmptyStr;
  for aIndex := 0 to fmHrMain.MsgTree.Count - 1 do
  begin
    aMsgInfo := TMsgInfo( fmHrMain.MsgTree.Nodes.AbsoluteItems[aIndex].Data );
    if ( not aMsgInfo.IsRead ) then
    begin
      Inc( aUnRead );
      if ( aMsgInfo.MsgPriority = '1' ) then
        Inc( aHighProirity )
      else
        Inc( aNormalProirity );
    end;
  end;
  if ( aUnRead > 0 ) then
  begin
    aTrayInfoText := Format( '您共有%d則尚未閱讀的訊息。', [aUnRead] );
    if ( aHighProirity > 0 ) then
      aTrayInfoText := ( aTrayInfoText + Format( #13'%d則重要的訊息。', [aHighProirity] ) );
    if ( aNormalProirity > 0 ) then
      aTrayInfoText := ( aTrayInfoText + Format( #13'%d則一般訊息。', [aNormalProirity] ) );
    fmHrMain.TrayIcon.BalloonTitle := '訊息通知';
    fmHrMain.TrayIcon.BalloonHint := aTrayInfoText;
    fmHrMain.TrayIcon.Hint := aTrayInfoText;
    fmHrMain.TrayIcon.Animate := True;
    if ( AIsMsgArrival or fmHrMain.TrayIcon.BalloonShowing ) then
      fmHrMain.TrayIcon.ShowBalloonHint;
  end else
  begin
    fmHrMain.TrayIcon.BalloonTitle := EmptyStr;
    fmHrMain.TrayIcon.BalloonHint := EmptyStr;
    fmHrMain.TrayIcon.Animate := False;
    fmHrMain.TrayIcon.IconIndex := 0;
    fmHrMain.TrayIcon.Hint := '客服小幫手';
    fmHrMain.TrayIcon.ShowBalloonHint;
  end;
  Application.ProcessMessages;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataClientModule.UpdateMsgItem(const ANode: TcxTreeListNode);
var
  aMsgInfo: TMsgInfo;
begin
  if Assigned( ANode ) then
  begin
    aMsgInfo := TMsgInfo( ANode.Data );
    if Assigned( aMsgInfo ) then aMsgInfo.IsRead := True;
    ANode.ImageIndex := StyleModule.GetMsgItemImageIndex( aMsgInfo.IsRead );
    ANode.SelectedIndex := ANode.ImageIndex;
    RefreshTrayInfoText;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataClientModule.DeleteMsgItem(const ANode: TcxTreeListNode);
begin
  if Assigned( ANode ) then
  begin
    try
      TMsgInfo( ANode.Data ).Free;
      ANode.Data := nil;
      ANode.Delete;
    except
      {}
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataClientModule.AddForm(AForm: TForm);
begin
  if AForm is THrClientForm then
  begin
    case THrClientForm( AForm ).HrClientFormKind of
      ftAnn:
        AnnFormList.Add( AForm );
      ftMsg:
        MsgFormList.Add( AForm );
    end;
  end;
end;             

{ ---------------------------------------------------------------------------- }

procedure TDataClientModule.RemoveForm(AForm: TForm);
begin
  if AForm is THrClientForm then
  begin
    case THrClientForm( AForm ).HrClientFormKind of
      ftAnn:
        AnnFormList.Remove( AForm );
      ftMsg:
        MsgFormList.Remove( AForm );
    end;
  end;
end;
{ ---------------------------------------------------------------------------- }

procedure TDataClientModule.NotifyMsgForm(const AReadyToSend: Boolean);
var
  aIndex: Integer;
  aSMForm: TfmSendMsg;
begin
  for aIndex := 0 to MsgFormList.Count - 1 do
  begin
    aSMForm := TfmSendMsg( MsgFormList[aIndex] );
    aSMForm.ChangeReadyState( AReadyToSend );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TDataClientModule.ForceFormClose;
var
  aForm: THrClientForm;
begin
  while ( MsgFormList.Count > 0 ) do
  begin
    aForm := THrClientForm( MsgFormList[0] );
    RemoveForm( aForm );
    aForm.OnClose := nil;
    aForm.Close;
  end;
  {}
  while ( AnnFormList.Count > 0 ) do
  begin
    aForm := THrClientForm( AnnFormList[0] );
    RemoveForm( aForm );
    aForm.OnClose := nil;
    aForm.Close;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TDataClientModule.GetMsgForm(const AMsgInfo: TMsgInfo): TForm;
var
  aIndex: Integer;
  aSMForm: TfmSendMsg;
begin
  Result := nil;
  for aIndex := 0 to MsgFormList.Count - 1 do
  begin
    if ( TObject( MsgFormList[aIndex] ) is TfmSendMsg ) then
    begin
      aSMForm := TfmSendMsg( MsgFormList[aIndex] );
      if ( aSMForm.MsgInfo.CompCode = AMsgInfo.CompCode ) and
         ( aSMForm.MsgInfo.MsgId = AMsgInfo.MsgId ) then
      begin
        Result := TForm( MsgFormList[aIndex] );
        Break;
      end;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
