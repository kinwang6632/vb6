unit cbCallbackEventThread;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, cbSyncObjs,
{$IFDEF APPDEBUG}
  CodeSiteLogging,
{$ENDIF}
{ App: }
  cbClientClass, cbThread, cbDesignPattern, cbSo, cbROClientModule, cbHrHelper,
{ ODAC: }
  MemDS, VirtualTable,
{ RemObject: }
  uROTypes, uROEventRepository, CsHorse2Library_Intf;


type

  TCallbackAdapter = class(TComponent, ISrvCallbackEvent)
  private
    FOwnerThread: TThread;
    FCallbackEventReceiver: TROEventReceiver;
    FUserBufferLock: TMutex;
  public
    constructor Create(AOwner: TThread; ALock: THandleObject); reintroduce;
    destructor Destroy; override;
    procedure BeginCallback;
    procedure EndCallback;
    procedure UsersChange(const AInfo: TLoginInfo);
    procedure ShutdownServer(const AMessage: String);
    procedure MsgChange(const AMsgInfo: TMsgInfo);
  end;


  TCallbackEventThread = class(TBaseThread)
  private
    { Private declarations }
    FReadyEvent: TSimpleEvent;
    FLastGetUserTime: TDateTime;
    FManualRefresh: Boolean;
    FGroupBuffer: TVirtualTable;
    FUserListBuffer: TVirtualTable;
    FMsgBuffer: TVirtualTable;
    FUserSubject: TDataSubject;
    FGroupSubject: TDataSubject;
    FMsgSubject: TDataSubject;
    FUserBufferLock: TMutex;
    FCallbackAdapter: TCallbackAdapter;
    FMsgList: TSyncList;
    FMsgFullUpdate: Boolean;
    function WantToUserListFullUpdate: Boolean;
    function WantToUpdateUserListGUI: Boolean;
    function WantToMsgListFullUpdate: Boolean;
    function WantToMsgListUpdate: Boolean;
    procedure ApplyBufferFilter(ABuffer: TVirtualTable; const AFilter: String);
    procedure DoUserGroupText;
    procedure DoUserStatusText;
    procedure DoUserStatusImageIndex;
    procedure CheckUserBuffer;
    procedure UpdateUserList; overload;
    procedure UpdateUserList(const AInfo: TLoginInfo); overload;
    procedure RefreshUserBufferDisplayText;
    procedure UpdateMsgList;
    procedure UpdateUserListGUI;
    procedure UpdateMsgListGUI;
    procedure Cleanup;
  protected
    procedure WndProc(var Msg: TMessage); override;
    procedure Execute; override;
  public
    constructor Create(CreateSuspended: Boolean); override;
    destructor Destroy; override;
    property ReadyEvent: TSimpleEvent read FReadyEvent write FReadyEvent;
    property UserSubject: TDataSubject read FUserSubject;
    property MsgSubject: TDataSubject read FMsgSubject;
  end;

implementation

uses cbUtilis, cbMain, cbStyleModule, DateUtils;

var FUpdateUserListCount: Integer = 0;

{ Important: Methods and properties of objects in visual components can only be
  used in a method called using Synchronize, for example,

      Synchronize(UpdateCaption);

  and UpdateCaption could look like,

    procedure TSgThread.UpdateCaption;
    begin
      Form1.Caption := 'Updated in a thread';
    end; }

{ ---------------------------------------------------------------------------- }

{ TCallbackAdapter }

constructor TCallbackAdapter.Create(AOwner: TThread; ALock: THandleObject);
begin
  inherited Create( nil );
  FOwnerThread := AOwner;
  FCallbackEventReceiver := TROEventReceiver.Create( nil );
  FCallbackEventReceiver.SynchronizeInvoke := False;
  FCallbackEventReceiver.Channel := ROClientModule.ROSuperTcpChannel;
  FCallbackEventReceiver.Message := ROClientModule.ROCallbackSrvBinMessage;
  FCallbackEventReceiver.ServiceName := 'CallbackService';
  FUserBufferLock := TMutex( ALock );
end;

{ ---------------------------------------------------------------------------- }

destructor TCallbackAdapter.Destroy;
begin
  FOwnerThread := nil;
  FUserBufferLock := nil;
  FCallbackEventReceiver.Free;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TCallbackAdapter.BeginCallback;
begin
  if ( not FCallbackEventReceiver.IsEventHandlerRegistered( EID_SrvCallbackEvent, Self ) ) then
  begin
    FCallbackEventReceiver.RegisterEventHandler( EID_SrvCallbackEvent, Self );
    FCallbackEventReceiver.Activate;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TCallbackAdapter.EndCallback;
begin
  if ( FCallbackEventReceiver.IsEventHandlerRegistered( EID_SrvCallbackEvent, Self ) ) then
  begin
    FCallbackEventReceiver.Deactivate;
    FCallbackEventReceiver.UnregisterEventHandler( Self );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TCallbackAdapter.UsersChange(const AInfo: TLoginInfo);
var
  aCompStr: String;
begin
  aCompStr := ROClientModule.LoginInfo.CompStr;
  if IsCompCodeFound( AInfo.CompCode, aCompStr ) then
  begin
    FUserBufferLock.Acquire;
    try
      TCallbackEventThread( FOwnerThread ).UpdateUserList( AInfo );
      TCallbackEventThread( FOwnerThread ).RefreshUserBufferDisplayText;
      Inc( FUpdateUserListCount );
    finally
      FUserBufferLock.Release;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TCallbackAdapter.ShutdownServer(const AMessage: String);
begin

end;

{ ---------------------------------------------------------------------------- }

procedure TCallbackAdapter.MsgChange(const AMsgInfo: TMsgInfo);
var
  aMsgInfo2: TMsgInfo;
begin
  aMsgInfo2 := TMsgInfo( AMsgInfo.Clone );
  TCallbackEventThread( FOwnerThread ).FMsgList.AddObject( aMsgInfo2.MsgId, aMsgInfo2 );
end;

{ ---------------------------------------------------------------------------- }

{ TSgThread }

constructor TCallbackEventThread.Create(CreateSuspended: Boolean);
begin
  inherited Create( CreateSuspended );
  FLastGetUserTime := 0;
  FManualRefresh := False;
  FUserBufferLock := TMutex.Create;
  FGroupBuffer := TVirtualTable.Create( nil );
  FGroupBuffer.Open;
  FUserListBuffer := TVirtualTable.Create( nil );
  FUserListBuffer.Open;
  FMsgBuffer := TVirtualTable.Create( nil );
  FMsgBuffer.Open;
  FUserSubject := TDataSubject.Create;
  FCallbackAdapter := TCallbackAdapter.Create( Self, FUserBufferLock );
  FMsgList := TSyncList.Create;
  FMsgFullUpdate := True;
  FMsgSubject := TDataSubject.Create;
end;

{ ---------------------------------------------------------------------------- }

destructor TCallbackEventThread.Destroy;
begin
  FGroupBuffer.Free;
  FUserListBuffer.Free;
  FMsgBuffer.Free;
  FGroupSubject.Free;
  FUserSubject.Free;
  FCallbackAdapter.Free;
  FUserBufferLock.Free;
  FMsgList.Free;
  FMsgSubject.Free;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TCallbackEventThread.WndProc(var Msg: TMessage);
begin
  case Msg.Msg of
    WM_MANUALREFRESH:
      FManualRefresh := True;
  end;
  inherited WndProc( Msg );
end;

{ ---------------------------------------------------------------------------- }

function TCallbackEventThread.WantToUserListFullUpdate: Boolean;
begin
  ROClientModule.BeginLockClientParam;
  try
    Result := not ROClientModule.ClientParam.IsEmpty;
    if ( Result ) then
    begin
      Result := ( FLastGetUserTime <= 0 );
      if ( not Result ) then
      begin
        if ( ROClientModule.ClientParam.FieldByName( 'AutoRefresh' ).AsBoolean ) then
        begin
          Result :=
            ( not Self.Terminated ) and
            ( ( SecondsBetween( Now, FLastGetUserTime ) >=
                  ROClientModule.ClientParam.FieldByName( 'UserRefreshRate' ).AsInteger )
                or
              ( FManualRefresh = True ) );
        end else
        begin
          Result := FManualRefresh;
        end;
      end;
    end;
  finally
    ROClientModule.EndLockClientParam;
  end;
  //Result := ( FUserListBuffer.IsEmpty or FManualRefresh );
end;

{ ---------------------------------------------------------------------------- }

function TCallbackEventThread.WantToUpdateUserListGUI: Boolean;
begin
  Result := ( not fmHrMain.MultiSelectUsers ) and ( FUpdateUserListCount > 0 );
end;

{ ---------------------------------------------------------------------------- }

function TCallbackEventThread.WantToMsgListFullUpdate: Boolean;
begin
  Result := ( FMsgFullUpdate or FManualRefresh );
end;

{ ---------------------------------------------------------------------------- }

function TCallbackEventThread.WantToMsgListUpdate: Boolean;
begin
  Result := ( FMsgList.Count > 0 );
end;

{ ---------------------------------------------------------------------------- }

procedure TCallbackEventThread.ApplyBufferFilter(ABuffer: TVirtualTable;
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

procedure TCallbackEventThread.UpdateUserList;
var
  aBinary: Binary;
begin
  aBinary := ROClientModule.GetGroupList;
  if Assigned( aBinary ) then
  begin
    FGroupBuffer.LoadFromStream( aBinary );
    FGroupBuffer.IndexFieldNames := 'CompCode;GroupId';
    aBinary.Free;
    if ( Self.Terminated ) then Exit;
  end;
  aBinary := ROClientModule.GetUserList;
  if Assigned( aBinary ) then
  begin
    FUserListBuffer.LoadFromStream( aBinary );
    FUserListBuffer.IndexFieldNames := 'CompCode;UserId';
    aBinary.Free;
    if ( Self.Terminated ) then Exit;
    CheckUserBuffer;
    Sleep( 100 );
  end;
  FLastGetUserTime := Now;
  //if FManualRefresh then FManualRefresh := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TCallbackEventThread.UpdateUserList(const AInfo: TLoginInfo);
begin
  if FUserListBuffer.Locate( 'CompCode;UserId', VarArrayOf( [
    AInfo.CompCode, AInfo.UserId] ), [] ) then
  begin
    FUserListBuffer.Edit;
    FUserListBuffer.FieldByName( 'WorkClass' ).AsString := AInfo.WorkClass;
    FUserListBuffer.FieldByName( 'Status' ).AsString := AInfo.Status;
    FUserListBuffer.FieldByName( 'RStatus' ).AsString := AInfo.Status;
    FUserListBuffer.Post;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TCallbackEventThread.CheckUserBuffer;
var
  aIndex: Integer;
  aPreCompCode: String;
begin
  { 1.Append Status Info Record into UserBuffer }
  FUserListBuffer.First;
  FGroupBuffer.First;
  while not FGroupBuffer.Eof do
  begin
    if ( aPreCompCode <> FGroupBuffer.FieldByName( 'CompCode' ).AsString ) then
    begin
      for aIndex := 0 to 1 do
      begin
        FUserListBuffer.Append;
        FUserListBuffer.FieldByName( 'CompCode' ).AsString := FGroupBuffer.FieldByName( 'CompCode' ).AsString;
        FUserListBuffer.FieldByName( 'CompName' ).AsString := FGroupBuffer.FieldByName( 'CompName' ).AsString;
        FUserListBuffer.FieldByName( 'UserId' ).AsString := IntToStr( aIndex );
        FUserListBuffer.FieldByName( 'Status' ).AsString := IntToStr( aIndex );
        FUserListBuffer.FieldByName( 'RStatus' ).AsString := EmptyStr;
        FUserListBuffer.FieldByName( 'DisplayType' ).AsInteger := 1;
        FUserListBuffer.FieldByName( 'ImageIndex' ).AsInteger := -1;
        FUserListBuffer.Post;
        aPreCompCode := FGroupBuffer.FieldByName( 'CompCode' ).AsString;
        if ( Self.Terminated ) then Break;
      end;
    end;
    FGroupBuffer.Next;
  end;
  if ( Self.Terminated ) then Exit;
  { 2.Append Group Info Record into UserBuffer }
  FUserListBuffer.First;
  FGroupBuffer.First;
  while not FGroupBuffer.Eof do
  begin
    FUserListBuffer.Append;
    FUserListBuffer.FieldByName( 'CompCode' ).AsString := FGroupBuffer.FieldByName( 'CompCode' ).AsString;
    FUserListBuffer.FieldByName( 'CompName' ).AsString := FGroupBuffer.FieldByName( 'CompName' ).AsString;
    FUserListBuffer.FieldByName( 'UserId' ).AsString := FGroupBuffer.FieldByName( 'GroupId' ).AsString;
    FUserListBuffer.FieldByName( 'WorkClass' ).AsString := FGroupBuffer.FieldByName( 'GroupId' ).AsString;
    FUserListBuffer.FieldByName( 'GroupName' ).AsString := FGroupBuffer.FieldByName( 'GroupName' ).AsString;
    FUserListBuffer.FieldByName( 'RGroupId' ).AsString := FGroupBuffer.FieldByName( 'RGroupId' ).AsString;
    FUserListBuffer.FieldByName( 'DisplayType' ).AsInteger := 0;
    FUserListBuffer.FieldByName( 'ImageIndex' ).AsInteger := -1;
    FUserListBuffer.Post;
    FGroupBuffer.Next;
    if ( Self.Terminated ) then Break;
  end;
  {3.Delete Slef at UserList }
  if FUserListBuffer.Locate( 'CompCode;UserId', VarArrayOf( [
    ROClientModule.LoginInfo.CompCode, ROClientModule.LoginInfo.UserId] ), [] ) then
   FUserListBuffer.Delete;
end;

{ ---------------------------------------------------------------------------- }

procedure TCallbackEventThread.RefreshUserBufferDisplayText;
begin
  DoUserGroupText;
  if Self.Terminated then Exit;
  DoUserStatusText;
  if Self.Terminated then Exit;
  DoUserStatusImageIndex;
end;

{ ---------------------------------------------------------------------------- }

procedure TCallbackEventThread.UpdateUserListGUI;
begin
  FUserSubject.Notify( FUserListBuffer );
  FUpdateUserListCount := 0;
end;

{ ---------------------------------------------------------------------------- }

procedure TCallbackEventThread.UpdateMsgListGUI;
begin
  while ( FMsgList.Count > 0 ) do
  begin
    FMsgSubject.Notify( FMsgList.Objects[0] );
    { Don't Free, Because It Store fmHrMain.MsgTree.Items
      TMsgInfo( FMsgList.Objects[0] ).Free;
    }
    FMsgList.Delete( 0 );
    if Self.Terminated then Break;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TCallbackEventThread.Cleanup;
begin
  while ( FMsgList.Count > 0 ) do
  begin
    if Assigned( FMsgList.Objects[0] ) then
    begin
      TMsgInfo( FMsgList[0] ).Free;
      FMsgList.Delete( 0 );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TCallbackEventThread.Execute;
begin
  while not Self.Terminated do
  begin
    try
      if ( FReadyEvent.WaitFor( 1000 ) = wrSignaled ) then
      begin
        if ( WantToUserListFullUpdate ) then
        begin
          Sleep( 1000 );
          FUserBufferLock.Acquire;
          try
            UpdateUserList;
            if Self.Terminated then Break;
            RefreshUserBufferDisplayText;
            if Self.Terminated then Break;
            UpdateUserListGUI;
          finally
            FUserBufferLock.Release;
          end;
          if Self.Terminated then Break;
          FCallbackAdapter.BeginCallback;
        end;
        Sleep( 100 );
        if Self.Terminated then Break;
        {}
        if WantToMsgListFullUpdate then
          UpdateMsgList;
      end;
      {}
      if ( WantToUpdateUserListGUI ) then
        UpdateUserListGUI;
      {}
      if Self.Terminated then Break;
      {}
      if ( WantToMsgListUpdate ) then
        UpdateMsgListGUI;
      {}  
      if Self.Terminated then Break;  
    except
      on E: Exception do
      begin
        {$IFDEF APPDEBUG}
          CodeSite.SendError( '%s=%s', [Self.ClassName,E.Message] );
        {$ENDIF}
      end;
    end;
    Sleep( 100 );
  end;
  FCallbackAdapter.EndCallback;
  Cleanup;
end;

{ ---------------------------------------------------------------------------- }

procedure TCallbackEventThread.DoUserGroupText;

  { ------------------------------------------------- }

  function CreateTemporyBuffer: TVirtualTable;
  begin
    Result := TVirtualTable.Create( nil );
    Result.Open;
    Result.Assign( FGroupBuffer );
  end;

  { ------------------------------------------------- }

  procedure DestroyTemporyBuffer(ATempBuffer: TVirtualTable);
  begin
    ATempBuffer.Clear;
    ATempBuffer.Free;
  end;

  { ------------------------------------------------- }

  function GenRecursiveFilter(ACompCode, AWorkClass: String): String;
  begin
    Result := Format( '(CompCode=''%s'') AND (RGroupId=''%s'')',
      [ACompCode, AWorkClass] );
  end;

  { ------------------------------------------------- }

  function GenUserCountFilter(ACompCode, AWorkClass: String): String;
  begin
    Result := Format( '(CompCode=''%s'') AND (WorkClass=''%s'') AND (DisplayType=''2'')',
      [ACompCode, AWorkClass] );
  end;

  { ------------------------------------------------- }

  procedure RecursiveGetCount(ACompCode, AGroupId: String; var ATotal, AOnline: Integer);
  var
    aFilterText: String;
    aTmpBuffer: TVirtualTable;
  begin
    aFilterText := GenRecursiveFilter( ACompCode, AGroupId );
    aTmpBuffer := CreateTemporyBuffer;
    try
      ApplyBufferFilter( aTmpBuffer, aFilterText );
      if ( aTmpBuffer.RecordCount > 0 ) then
      begin
        aTmpBuffer.First;
        while not aTmpBuffer.Eof do
        begin
          RecursiveGetCount( aTmpBuffer.FieldByName( 'CompCode' ).AsString,
            aTmpBuffer.FieldByName( 'GroupId' ).AsString, ATotal, AOnline );
          aTmpBuffer.Next;
        end;
      end;
      aFilterText := GenUserCountFilter( ACompCode, AGroupId );
      ApplyBufferFilter( FUserListBuffer, aFilterText );
      FUserListBuffer.First;
      while not FUserListBuffer.Eof do
      begin
        Inc( ATotal );
        if ( FUserListBuffer.FieldByName( 'Status' ).AsInteger in [1..2] )then
          Inc( AOnline );
        FUserListBuffer.Next;
      end;
      ApplyBufferFilter( FUserListBuffer, EmptyStr );
    finally
      DestroyTemporyBuffer( aTmpBuffer );
    end;
  end;

  { ------------------------------------------------- }

var
  aTotal, aOnline: Integer;
  aDisplayText: String;
begin
  FGroupBuffer.First;
  while not FGroupBuffer.Eof do
  begin
    aTotal := 0;
    aOnline := 0;
    RecursiveGetCount( FGroupBuffer.FieldByName( 'CompCode' ).AsString,
      FGroupBuffer.FieldByName( 'GroupId' ).AsString, aTotal, aOnline );
    aDisplayText := Format( '%s (%d/%d)', [FGroupBuffer.FieldByName( 'GroupName' ).AsString,
      aOnline, aTotal] );
    if FUserListBuffer.Locate( 'CompCode;WorkClass;DisplayType', VarArrayOf( [
      FGroupBuffer.FieldByName( 'CompCode' ).AsString,
      FGroupBuffer.FieldByName( 'GroupId' ).AsString, '0'] ), [] ) then
    begin
      FUserListBuffer.Edit;
      FUserListBuffer.FieldByName( 'DisplayText1' ).AsString := aDisplayText;
      FUserListBuffer.Post;
    end;
    FGroupBuffer.Next;
    if Self.Terminated then Break;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TCallbackEventThread.DoUserStatusText;

  { ------------------------------------------------- }

  function GenStatusFilter(ACompCode, AStatus: String): String;
  begin
    Result := Format( '(CompCode=''%s'') AND (Status=''%s'') AND (DisplayType=''2'')',
      [ACompCode, AStatus] );
  end;

  { ------------------------------------------------- }

  function GenLocateFilter(ACompCode, AStatus: String): String;
  begin
    Result := Format( '(CompCode=''%s'') AND (Status=''%s'') AND (DisplayType=''1'')',
      [ACompCode, AStatus] );
  end;

  { ------------------------------------------------- }

  procedure CountingUserStatus(const ACompCode: String; var AOnline, AOffline: Integer);
  var
    aIndex: Integer;
    aFilterText: String;
  begin
    for aIndex := 0 to 1 do
    begin
      aFilterText := GenStatusFilter( ACompCode, IntToStr( aIndex ) );
      ApplyBufferFilter( FUserListBuffer, aFilterText );
      if TClinetUserStatus( aIndex ) = usOffline then
        Inc( AOffline, FUserListBuffer.RecordCount )
      else
        Inc( AOnline, FUserListBuffer.RecordCount );
    end;
    ApplyBufferFilter( FUserListBuffer, EmptyStr );
  end;

  { ------------------------------------------------- }

var
  aIndex, aOnline, aOffline: Integer;
  aPreCompCode, aDisplayText: String;
begin
  FGroupBuffer.First;
  aPreCompCode := EmptyStr;
  while not FGroupBuffer.Eof do
  begin
    if ( aPreCompCode <> FGroupBuffer.FieldByName( 'CompCode' ).AsString ) then
    begin
      aPreCompCode := FGroupBuffer.FieldByName( 'CompCode' ).AsString;
      aOnline := 0;
      aOffline := 0;
      CountingUserStatus( aPreCompCode, aOnline, aOffline );
      for aIndex := 0 to 1 do
      begin
        if FUserListBuffer.Locate( 'CompCode;UserId;DisplayType', VarArrayOf( [
          aPreCompCode, IntToStr( aIndex ), '1'] ), [] ) then
        begin
          aDisplayText := Format( '%s (%d)', [GetClientUserStatusText( aIndex ),
            IIF( aIndex = 0, aOffline, aOnline )] );
          FUserListBuffer.Edit;
          FUserListBuffer.FieldByName( 'DisplayText2' ).AsString := aDisplayText;
          FUserListBuffer.Post;
        end;
      end;
    end;
    FGroupBuffer.Next;
    if Self.Terminated then Break;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TCallbackEventThread.DoUserStatusImageIndex;

  { ------------------------------------------------- }

  function GenUserFilter: String;
  begin
    Result := 'DisplayType=''2''';
  end;

  { ------------------------------------------------- }

var
  aFilterText: String;
begin
  aFilterText := GenUserFilter;
  ApplyBufferFilter( FUserListBuffer, aFilterText );
  FUserListBuffer.First;
  while ( not FUserListBuffer.Eof ) do
  begin
    FUserListBuffer.Edit;
    FUserListBuffer.FieldByName( 'ImageIndex' ).AsInteger := StyleModule.GetUserStatusImageIndex(
      FUserListBuffer.FieldByName( 'Status' ).AsInteger );
    FUserListBuffer.FieldByName( 'DisplayText1' ).AsString := FUserListBuffer.FieldByName( 'UserName' ).AsString;
    FUserListBuffer.FieldByName( 'DisplayText2' ).AsString := Format( '%s (%s)', [
      FUserListBuffer.FieldByName( 'UserName' ).AsString,
      FUserListBuffer.FieldByName( 'GroupName' ).AsString] );
    FUserListBuffer.Post;
    FUserListBuffer.Next;
    if Self.Terminated then Break;
  end;
  ApplyBufferFilter( FUserListBuffer, EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

procedure TCallbackEventThread.UpdateMsgList;
var
  aBinary: Binary;
  aMsgInfo: TMsgInfo;
begin
  FMsgFullUpdate := False;
  if FManualRefresh then FManualRefresh := False;
  aBinary := ROClientModule.GetMsgList;
  if Assigned( aBinary ) then
  begin
    FMsgBuffer.LoadFromStream( aBinary );
    aBinary.Free;
  end;
  while not FMsgBuffer.Eof do
  begin
    aMsgInfo := TMsgInfo.Create;
    aMsgInfo.CompCode := FMsgBuffer.FieldByName( 'CompCode' ).AsString;
    aMsgInfo.CompName := FMsgBuffer.FieldByName( 'CompName' ).AsString;
    aMsgInfo.MsgPriority := FMsgBuffer.FieldByName( 'MsgPriority' ).AsString;
    aMsgInfo.MsgId := FMsgBuffer.FieldByName( 'MsgId' ).AsString;
    aMsgInfo.MsgSubject := FMsgBuffer.FieldByName( 'MsgSubject' ).AsString;
    aMsgInfo.MsgTime := FMsgBuffer.FieldByName( 'MsgTime' ).AsString;
    aMsgInfo.MsgReply := FMsgBuffer.FieldByName( 'MsgReply' ).AsString;
    aMsgInfo.MsgSenderId := FMsgBuffer.FieldByName( 'MsgSenderId' ).AsString;
    aMsgInfo.MsgSenderName := FMsgBuffer.FieldByName( 'MsgSenderName' ).AsString;
    aMsgInfo.MsgSenderWorkClass := FMsgBuffer.FieldByName( 'MsgSenderWorkClass' ).AsString;
    aMsgInfo.MsgSenderWorkName := FMsgBuffer.FieldByName( 'MsgSenderWorkName' ).AsString;
    aMsgInfo.IsRead := False;
    FMsgList.AddObject( aMsgInfo.MsgId, aMsgInfo );
    FMsgBuffer.Next;
  end;
  FMsgBuffer.Clear;
end;

{ ---------------------------------------------------------------------------- }

end.
