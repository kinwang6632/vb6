unit TListernerU;

interface

uses
  Classes, Sockets, SysUtils,Dialogs;

type
  TListerner = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  public
    G_TcpListener: TTcpServer;
    procedure TcpServerAccept(Sender: TObject; ClientSocket: TCustomIpClient);

    constructor Create(sI_ListenIP, sI_ListenPort:String);overload;    
  end;

implementation

uses dtmMainU;



{ Important: Methods and properties of objects in VCL or CLX can only be used
  in a method called using Synchronize, for example,

      Synchronize(UpdateCaption);

  and UpdateCaption could look like,

    procedure TListerner.UpdateCaption;
    begin
      Form1.Caption := 'Updated in a thread';
    end; }

{ TListerner }

constructor TListerner.Create(sI_ListenIP, sI_ListenPort: String);
begin
    G_TcpListener:= TTcpServer.Create(nil);
    G_TcpListener.LocalHost := sI_ListenIP;
    G_TcpListener.LocalPort := sI_ListenPort;
    G_TcpListener.OnAccept := TcpServerAccept;
    G_TcpListener.Open;
    inherited Create(false);
end;

procedure TListerner.Execute;
begin
  { Place thread code here }
end;

procedure TListerner.TcpServerAccept(Sender: TObject;
  ClientSocket: TCustomIpClient);
var
  sL_SendNagraData: string;
begin
//  sleep(10000);
  try
    sL_SendNagraData := ClientSocket.Receiveln;
    while sL_SendNagraData <> '' do
    begin
      dtmMain.separateCaReturnString(ClientSocket.RemoteHost, sL_SendNagraData);

      sL_SendNagraData := ClientSocket.Receiveln;
    end;
  except

  end;

end;

end.
