
unit ReceivingCommandResponseThreadU;

interface

uses
  Windows, Classes, Forms, Comctrls, Sysutils, Dialogs, ScktComp, WinSock,
  CsIntf;


type
  ReceivingCommandResponseThread = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  public
    constructor Create;overload;
    procedure ReceiveResponse;
  end;

implementation

uses frmMainU, HandleResponseU, ConstObjU, Ustru;

{ ReceivingCommandResponseThread }

{ ---------------------------------------------------------------------------- }

constructor ReceivingCommandResponseThread.Create;
begin
  inherited Create( False );
  Self.FreeOnTerminate := False;
  Self.OnTerminate := frmMain.ReceiveResponseThreadDone;
  Self.Resume;
end;

{ ---------------------------------------------------------------------------- }

procedure ReceivingCommandResponseThread.Execute;
begin
  while not Self.Terminated do
  begin
    Sleep( 100 );
    try
      if ( not Self.Terminated ) and
         ( frmMain.ControlSocket.Connected ) then
      begin
        if frmMain.bG_HasBuildSendCmdConnection and
           frmMain.bG_RunReceivingCommandThread then
          ReceiveResponse;
      end;
    except
      on E: Exception do
      begin
        CodeSite.SendError( 'RecvResponse thread inner loop Error: %s',
          [E.Message]  );
      end;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure ReceivingCommandResponseThread.ReceiveResponse;
var
  AIndex, ABytes, aSocketError : Integer;
  AData : String;
  AParseResult: TStringList;
  ABuffer: array[0..1024] of Char;
  //AWStream: TWinSocketStream;
begin
  ZeroMemory( @ABuffer, SizeOf( ABuffer ) );
  try
    ABytes := frmMain.ControlSocket.ReceiveBuf( ABuffer, SizeOf( ABuffer ) );
    aSocketError := WSAGetLastError;

    if aSocketError <> 0 then
    begin
      CodeSite.SendError( ' RecvData Error, Socket Error Code: %d', [aSocketError] );
      _NAGRA_REJECT := True;
    end; 
    
    if ABytes > 0 then
    begin
      AData := '';
      AData := frmMain.TransReceivedDataCharAryToStr( ABuffer, ABytes );
      AParseResult := TUstr.ParseStrings( AData, RECEIVED_DATA_SEP_FLAG );
      try
        for AIndex := 0 to AParseResult.Count -1 do
        begin
          frmMain.G_ResponseMsgStrList.Add( AParseResult.Strings[AIndex] );
          CodeSite.SendMsg( 'Recev Command Data = ' +
            AParseResult.Strings[AIndex] );
        end;
      finally
        AParseResult.Free;
      end;
    end;
    //CodeSite.SendInteger( csmMemoryStatus, 'Data recv byte', ABytes );
    
  Except
    on E: Exception do
    begin
      CodeSite.SendError( 'Data recv Error ' + E.Message );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
