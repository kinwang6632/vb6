unit cbSgThread;

interface

uses
  Windows, Messages, SysUtils, Classes, SyncObjs,
{ App: }
  cbClientClass, cbThread, cbDesignPattern, cbSo,
{ ODAC: }
  MemDS, VirtualTable,
{ RemObject: }
  cbROClientModule, uROTypes;


type
  TSgThread = class(TBaseThread)
  private
    { Private declarations }
    FEvent: TSimpleEvent;
    FLastGetUserTime: TDateTime;
    FManualRefresh: Boolean;
    FGroupBuffer: TVirtualTable;
    FUserListBuffer: TVirtualTable;
    FUserSubject: TDataSubject;
    FGroupSubject: TDataSubject;
    function WantToGetUserList: Boolean;
    procedure GetUserLisData;
  protected
    procedure WndProc(var Msg: TMessage); override;
    procedure Execute; override;
  public
    constructor Create(CreateSuspended: Boolean); override;
    destructor Destroy; override;
    property Event: TSimpleEvent read FEvent write FEvent;
    property UserSubject: TDataSubject read FUserSubject;
  end;

implementation

uses cbUtilis, DateUtils;

{ Important: Methods and properties of objects in visual components can only be
  used in a method called using Synchronize, for example,

      Synchronize(UpdateCaption);

  and UpdateCaption could look like,

    procedure TSgThread.UpdateCaption;
    begin
      Form1.Caption := 'Updated in a thread';
    end; }

{ TSgThread }

constructor TSgThread.Create(CreateSuspended: Boolean);
begin
  inherited Create( CreateSuspended );
  FGroupBuffer := TVirtualTable.Create( nil );
  FGroupBuffer.Open;
  FUserListBuffer := TVirtualTable.Create( nil );
  FUserListBuffer.Open;
  FUserSubject := TDataSubject.Create;
  FLastGetUserTime := 0;
  FManualRefresh := False;
end;

{ ---------------------------------------------------------------------------- }

destructor TSgThread.Destroy;
begin
  FGroupBuffer.Free;
  FUserListBuffer.Free;
  FGroupSubject.Free;
  FUserSubject.Free;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TSgThread.WndProc(var Msg: TMessage);
begin
  case Msg.Msg of
    WM_MANUALREFRESH:
      FManualRefresh := True;
  end;
  inherited WndProc( Msg );
end;

{ ---------------------------------------------------------------------------- }

function TSgThread.WantToGetUserList: Boolean;
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
end;

{ ---------------------------------------------------------------------------- }

procedure TSgThread.Execute;
begin
  while not Self.Terminated do
  begin
    if ( FEvent.WaitFor( 1000 ) = wrSignaled ) then
    begin
      if WantToGetUserList then
        GetUserLisData;
      Sleep( 100 );
      if Self.Terminated then Break;
    end;
    Sleep( 100 );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TSgThread.GetUserLisData;
var
  aBinary: Binary;
begin
//  aBinary := ROClientModule.GetGroupList;
//  if Assigned( aBinary ) then
//  begin
//    FGroupBuffer.LoadFromStream( aBinary );
//    aBinary.Free;
//  end;
  {}
  aBinary := ROClientModule.GetUserList;
  if Assigned( aBinary ) then
  begin
    FUserListBuffer.LoadFromStream( aBinary );
    aBinary.Free;
    if ( FUserListBuffer.IndexFieldNames = EmptyStr ) then
      FUserListBuffer.IndexFieldNames := 'CompCode;WorkClass;UserId';
    if ( Self.Terminated ) then Exit;  
    FUserSubject.Notify( FUserListBuffer );
    Sleep( 100 );
  end;
  FLastGetUserTime := Now;
  if FManualRefresh then FManualRefresh := False;
end;

{ ---------------------------------------------------------------------------- }


end.
