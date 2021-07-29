unit cbAnnThread;

interface

uses
  Windows, Messages, SysUtils, Classes, cbSyncObjs,
{$IFDEF APPDEBUG}
  CodeSiteLogging,
{$ENDIF}
{ App: }
  cbClientClass, cbThread, cbDesignPattern, cbSo, cbROClientModule,
{ ODAC: }
  MemDS, VirtualTable,
{ RemObject: }
  uROTypes;

type
  TAnnThread = class(TBaseThread)
  private
    { Private declarations }
    FReadyEvent: TSimpleEvent;
    FAnnSyncMutex: TMutex;
    FCompSubject: TDataSubject;
    FSo021Subject: TDataSubject;
    FCd042Subject: TDataSubject;
    FSo022Subject: TDataSubject;
    FSo023Subject: TDataSubject;
    FAnnBuffer: TVirtualTable;
    FLastGetCompTime: TDateTime;
    FLastGetAnnDataTime: TDateTime;
    FManualRefresh: Boolean;
    function WantToUpdateClientParameter: Boolean;
    function WantToGetCompInfo: Boolean;
    function WantToGetAnnData: Boolean;
  protected
    procedure WndProc(var Msg: TMessage); override;
    procedure Execute; override;
  public
    constructor Create(CreateSuspended: Boolean); override;
    destructor Destroy; override;
    procedure UpdateClientParameter;
    procedure GetCompInfo;
    procedure GetAnnData;
    procedure GetSO021;
    procedure GetCD042;
    procedure GetSO022;
    procedure GetSO023;
    property ReadyEvent: TSimpleEvent read FReadyEvent write FReadyEvent;
    property AnnSyncMutex: TMutex read FAnnSyncMutex;
    property CompSubject: TDataSubject read FCompSubject;
    property So021Subject: TDataSubject read FSo021Subject;
    property Cd042Subject: TDataSubject read FCd042Subject;
    property SO022Subject: TDataSubject read FSO022Subject;
    property SO023Subject: TDataSubject read FSO023Subject;
  end;

implementation

uses cbUtilis, DateUtils;

{ Important: Methods and properties of objects in visual components can only be
  used in a method called using Synchronize, for example,

      Synchronize(UpdateCaption);

  and UpdateCaption could look like,

    procedure TAnnThread.UpdateCaption;
    begin
      Form1.Caption := 'Updated in a thread';
    end; }

{ ---------------------------------------------------------------------------- }

{ TAnnThread }

constructor TAnnThread.Create(CreateSuspended: Boolean);
begin
  inherited Create( CreateSuspended );
  FAnnSyncMutex := TMutex.Create;
  FAnnBuffer := TVirtualTable.Create( nil );
  FAnnBuffer.Open;
  FCompSubject := TDataSubject.Create;
  FSo021Subject := TDataSubject.Create;
  FCd042Subject := TDataSubject.Create;
  FSo022Subject := TDataSubject.Create;
  FSo023Subject := TDataSubject.Create;
  FLastGetCompTime := 0;
  FLastGetAnnDataTime := 0;
  FManualRefresh := False;
end;

{ ---------------------------------------------------------------------------- }

destructor TAnnThread.Destroy;
begin
  FAnnBuffer.Free;
  FCompSubject.Free;  
  FSo021Subject.Free;
  FCd042Subject.Free;
  FSo022Subject.Free;
  FSo023Subject.Free;
  FAnnSyncMutex.Free;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TAnnThread.WndProc(var Msg: TMessage);
begin
  case Msg.Msg of
    WM_MANUALREFRESH:
      FManualRefresh := True;
  end;
  inherited WndProc( Msg );
end;

{ ---------------------------------------------------------------------------- }

function TAnnThread.WantToUpdateClientParameter: Boolean;
begin
  ROClientModule.BeginLockClientParam;
  try
    Result := ( ROClientModule.ClientParam.IsEmpty ) or ( FManualRefresh );
  finally
    ROClientModule.EndLockClientParam;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TAnnThread.WantToGetCompInfo: Boolean;
begin
  ROClientModule.BeginLockClientParam;
  try
    Result := not ROClientModule.ClientParam.IsEmpty;
    if ( Result ) then
    begin
      Result := ( FLastGetCompTime <= 0 );
      if ( not Result ) then
      begin
        if ( ROClientModule.ClientParam.FieldByName( 'AutoRefresh' ).AsBoolean ) then
        begin
          Result :=
            ( not Self.Terminated ) and
            ( ( SecondsBetween( Now, FLastGetCompTime ) >=
                  ROClientModule.ClientParam.FieldByName( 'AuthorizeRefreshRate' ).AsInteger )
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
end;

{ ---------------------------------------------------------------------------- }

function TAnnThread.WantToGetAnnData: Boolean;
begin
  ROClientModule.BeginLockClientParam;
  try
    Result := not ROClientModule.ClientParam.IsEmpty;
    if ( Result ) then
    begin
      Result := ( FLastGetAnnDataTime <= 0 );
      if ( not Result ) then
      begin
        if ( ROClientModule.ClientParam.FieldByName( 'AutoRefresh' ).AsBoolean ) then
        begin
          Result :=
            ( not Self.Terminated ) and
            ( ( SecondsBetween( Now, FLastGetAnnDataTime ) >=
                  ROClientModule.ClientParam.FieldByName( 'AnnRefreshRate' ).AsInteger )
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
end;

{ ---------------------------------------------------------------------------- }

procedure TAnnThread.UpdateClientParameter;
begin
  ROClientModule.UpdateClientParam;
end;

{ ---------------------------------------------------------------------------- }

procedure TAnnThread.GetCompInfo;
var
  aList: TSoList;
begin
  aList := ROClientModule.GetSoList;
  Sleep( 100 );
  FCompSubject.Notify( aList );
  Sleep( 100 );
  FLastGetCompTime := Now;
end;

{ ---------------------------------------------------------------------------- }

procedure TAnnThread.GetAnnData;
begin
  GetSO021;
  Sleep( 100 );
  if Self.Terminated then Exit;
  GetCD042;
  Sleep( 100 );
  if Self.Terminated then Exit;
  GetSO022;
  Sleep( 100 );
  if Self.Terminated then Exit;
  GetSO023;
  FLastGetAnnDataTime := Now;
  if FManualRefresh then FManualRefresh := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TAnnThread.GetSO021;
var
  aBinary: Binary;
begin
  aBinary := ROClientModule.GetSO021;
  if Assigned( aBinary ) then
  begin
    FAnnBuffer.LoadFromStream( aBinary );
    aBinary.Free;
    if ( Self.Terminated ) then Exit;
    FAnnSyncMutex.Acquire;
    try
      FSo021Subject.Notify( FAnnBuffer );
    finally
      FAnnSyncMutex.Release;
    end;
    FAnnBuffer.Clear;
    Sleep( 10 );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TAnnThread.GetCD042;
var
  aBinary: Binary;
begin
  aBinary := ROClientModule.GetCD042;
  if Assigned( aBinary ) then
  begin
    FAnnBuffer.LoadFromStream( aBinary );
    aBinary.Free;
    if ( Self.Terminated ) then Exit;
    FAnnSyncMutex.Acquire;
    try
      FCd042Subject.Notify( FAnnBuffer );
    finally
      FAnnSyncMutex.Release;
    end;
    FAnnBuffer.Clear;
    Sleep( 10 );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TAnnThread.GetSO022;
var
  aBinary: Binary;
begin
  aBinary := ROClientModule.GetSO022;
  if Assigned( aBinary ) then
  begin
    FAnnBuffer.LoadFromStream( aBinary );
    aBinary.Free;
    if ( Self.Terminated ) then Exit;
    FAnnSyncMutex.Acquire;
    try
      FSo022Subject.Notify( FAnnBuffer );
    finally
      FAnnSyncMutex.Release;
    end;
    FAnnBuffer.Clear;
    Sleep( 10 );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TAnnThread.GetSO023;
var
  aBinary: Binary;
begin
  aBinary := ROClientModule.GetSO023;
  if Assigned( aBinary ) then
  begin
    FAnnBuffer.LoadFromStream( aBinary );
    aBinary.Free;
    if ( Self.Terminated ) then Exit;
    FAnnSyncMutex.Acquire;
    try
      FSo023Subject.Notify( FAnnBuffer );
    finally
      FAnnSyncMutex.Release;
    end;
    FAnnBuffer.Clear;
    Sleep( 10 );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TAnnThread.Execute;
begin
  while not Self.Terminated do
  begin
    try
      if ( FReadyEvent.WaitFor( 1000 ) = wrSignaled ) then
      begin
        if WantToUpdateClientParameter then
           UpdateClientParameter;
        Sleep( 300 );
        if Self.Terminated then Break;
        {}
        if WantToGetCompInfo then
          GetCompInfo;
        Sleep( 100 );
        if Self.Terminated then Break;
        {}
        if WantToGetAnnData then
          GetAnnData;
        Sleep( 100 );
        if Self.Terminated then Break;
      end;
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
end;

{ ---------------------------------------------------------------------------- }

end.
