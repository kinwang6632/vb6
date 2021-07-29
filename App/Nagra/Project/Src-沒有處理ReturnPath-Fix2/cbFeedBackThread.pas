unit cbFeedBackThread;

interface

uses
  Classes, Windows, Messages, cbQueue, CbUtils;

type
  TFeedBackThread = class(TThread)
  private
    { Private declarations }
    FTransactionNumberGenerator: TTransactionNumberGenerator;
    FAlreadyConnectedToServer: Boolean;
  protected
    procedure Execute; override;
  public
    constructor Create;
    destructor Destroy; override;
  end;

implementation


{ TFeedBackThread }

{ ---------------------------------------------------------------------------- }

constructor TFeedBackThread.Create;
begin
  inherited Create( True );
  Self.Priority := tpLower;
  Self.FreeOnTerminate := False;
  Self.OnTerminate := frmMain.ProcessReturnPathDataThreadDone;
  FTransactionNumberGenerator := TTransactionNumberGenerator.Create( True, 0, 999999 );
  FAlreadyConnectedToServer := False;
  Self.Resume;
end;

{ ---------------------------------------------------------------------------- }

destructor TFeedBackThread.Destroy;
begin
  FTransactionNumberGenerator.Free;
  inherited;
end;

{ ---------------------------------------------------------------------------- }

procedure TFeedBackThread.Execute;
begin
  while not Self.Terminated do
  begin
    Sleep( 100 );
    if FAlreadyConnectedToServer then
    begin


    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
