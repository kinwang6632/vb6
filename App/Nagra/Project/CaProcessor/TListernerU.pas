unit TListernerU;

interface

uses
  Classes, Sockets, SysUtils, Forms;

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

uses frmMainU;

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
  sL_ClientNdsData: string;
  ii, nL_CmdType : Integer;
begin
  sL_ClientNdsData := ClientSocket.Receiveln;
  while sL_ClientNdsData <> '' do
  begin
    Application.ProcessMessages;  
    nL_CmdType := StrToIntDef(Copy(sL_ClientNdsData,1,1),0);
    case nL_CmdType of
      1: // ca 相關指令
        frmMain.parseClientNdsData(Copy(sL_ClientNdsData,3,length(sL_ClientNdsData)-2));
      2: // dispatcher 連接設定指令
        frmMain.dispatcherConnectSettingCmd(ClientSocket.RemoteHost, Copy(sL_ClientNdsData,3,length(sL_ClientNdsData)-2));
      3: // dispatcher 斷線設定指令
        frmMain.dispatcherDisConnectSettingCmd(ClientSocket.RemoteHost);

    end;


    sL_ClientNdsData := ClientSocket.Receiveln;
  end;

end;

end.
