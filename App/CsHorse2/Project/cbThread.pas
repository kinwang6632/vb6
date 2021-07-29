unit cbThread;

interface

uses Windows, Messages, SysUtils, Classes;

type

  { TBaseThread }

  TBaseThread = class(TThread)
  private
    FMsgWnd: THandle;
    FHexThreadId: String;
  protected
    procedure WndProc(var Msg: TMessage); virtual;
  public
    constructor Create(CreateSuspended: Boolean); virtual;
    destructor Destroy; override;
    property MsgWnd: THandle read FMsgWnd;
    property HexThreadId: String read FHexThreadId;
  end;


implementation

{ TMessageQueueThread }

constructor TBaseThread.Create(CreateSuspended: Boolean);
begin
  inherited Create( CreateSuspended );
  FMsgWnd := Classes.AllocateHWnd( WndProc );
  FHexThreadId := Format( '0x%.4x', [ThreadID] );
  FreeOnTerminate := False;
end;

{ ---------------------------------------------------------------------------- }

destructor TBaseThread.Destroy;
begin
  Classes.DeallocateHWnd( FMsgWnd );
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TBaseThread.WndProc(var Msg: TMessage);
begin
  Msg.Result := DefWindowProc( FMsgWnd, Msg.Msg, Msg.wParam, Msg.lParam );
end;

{ ---------------------------------------------------------------------------- }


end.
