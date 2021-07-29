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

uses frmMainU, ConstU;

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
    G_TcpListener.BlockMode := bmThreadBlocking;
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
      SEND_TO_DISPATCHER_INFO_TYPE_1: // ca 相關指令
        frmMain.parseClientNdsData(Copy(sL_ClientNdsData,3,length(sL_ClientNdsData)-2));
      SEND_TO_DISPATCHER_INFO_TYPE_2: // dispatcher 連接設定指令
        frmMain.dispatcherConnectSettingCmd(ClientSocket.RemoteHost, Copy(sL_ClientNdsData,3,length(sL_ClientNdsData)-2));
      SEND_TO_DISPATCHER_INFO_TYPE_3: // dispatcher 斷線設定指令
        frmMain.dispatcherDisConnectSettingCmd(ClientSocket.RemoteHost);
      else
        frmMain.showLogOnMemo(BAD_CLIENT_DATA_FORMAT, ClientSocket.RemoteHost + ' 傳來錯誤的指令字串.如下: ' + sL_ClientNdsData);
    end;


    sL_ClientNdsData := ClientSocket.Receiveln;
  end;

end;

end.
