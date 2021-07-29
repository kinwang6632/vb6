unit TListernerU;

interface

uses
  Classes, Sockets, SysUtils, Forms ,IdBaseComponent,IdComponent,
  IdTCPServer, IdSocketHandle ,IdThreadMgr, IdThreadMgrDefault,Dialogs;

type
  TSimpleClient = class(TObject)
    IP: string;
    ListLink: Integer;
    Thread: Pointer;
    CompName: string;
  end;

  TListerner = class(TThread)
  private
    { Private declarations }
    sG_ListenIP, sG_ListenPort : String;
    procedure UpdateBindings;
  protected
    procedure Execute; override;
  public
    G_TcpListener: TTcpServer;
    G_IdTcpListener: TIdTcpServer;
    G_IdThreadMgrDefault : TIdThreadMgrDefault;
    G_ClientArray : array of TSimpleClient;
    nG_AryLength : Integer;
    procedure TcpServerAccept(Sender: TObject; ClientSocket: TCustomIpClient);
    constructor Create(sI_ListenIP, sI_ListenPort:String);overload;
    procedure IdTcpServerConnect(AThread: TIdPeerThread);
    procedure IdTcpServerExecute(AThread: TIdPeerThread);
    procedure IdTcpServerDisconnect(AThread: TIdPeerThread);
    procedure clientDisconnect(AThread: TIdPeerThread);
    function sendReturnStringToSimpleCA(sI_ClientIP,sI_ReturnString : String) : String;overload;
    procedure disconnectAllClient;
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

procedure TListerner.clientDisconnect(AThread: TIdPeerThread);
var
    Client: TSimpleClient;
    ii : Integer;
    sL_ClientIP : String;
begin
    Client := Pointer(AThread.Data);
    sL_ClientIP := Client.IP;
    //Remove Client from the Clients TList
    for ii:=0 to high(G_ClientArray) do
    begin
      if G_ClientArray[ii].IP = sL_ClientIP then
      begin
        if Assigned(G_ClientArray[ii]) then
        begin
          G_ClientArray[ii].Free;
          G_ClientArray[ii] := nil;
        end;
        break;
      end;
    end;
    //Free the Client object
    //Client.Free;
    //AThread.Data := nil;
    Dec(nG_AryLength);
end;

constructor TListerner.Create(sI_ListenIP, sI_ListenPort: String);
begin
{
    G_TcpListener:= TTcpServer.Create(nil);
    G_TcpListener.LocalHost := sI_ListenIP;
    G_TcpListener.LocalPort := sI_ListenPort;
    G_TcpListener.BlockMode := bmThreadBlocking;
    G_TcpListener.OnAccept := TcpServerAccept;
    G_TcpListener.Open;
    inherited Create(false);
}
//frmMain.Memo5.Lines.Add('Create== ' + IntToStr(nG_AryLength));
    sG_ListenIP := sI_ListenIP;
    sG_ListenPort := sI_ListenPort;

    G_IdTcpListener := TIdTcpServer.Create(nil);
    G_IdThreadMgrDefault := TIdThreadMgrDefault.Create(nil);
    try
      G_IdTcpListener.ThreadMgr := G_IdThreadMgrDefault;
      UpdateBindings;
      G_IdTcpListener.Active := not G_IdTcpListener.Active;
      G_IdTcpListener.OnConnect := IdTcpServerConnect;
      G_IdTcpListener.OnExecute := IdTcpServerExecute;
      G_IdTcpListener.OnDisconnect := IdTcpServerDisconnect;
    except

    end;

    inherited Create(false);
end;

procedure TListerner.disconnectAllClient;
var
    sL_Info,sL_ClientIP : String;
    ii : Integer;
begin
    sL_Info := frmMain.setInfoXml('',SIMPLE_CA_DISCONNECT);

    for ii:=0 to nG_AryLength-1 do
    begin
      sL_ClientIP := G_ClientArray[ii].IP;
      sendReturnStringToSimpleCA(sL_ClientIP,sL_Info);
      G_ClientArray[ii].Free;
      G_ClientArray[ii] := nil;
    end;
end;

procedure TListerner.Execute;
begin
  { Place thread code here }
end;

procedure TListerner.IdTcpServerConnect(AThread: TIdPeerThread);
begin
  //AThread.Connection.WriteLn('您已經連接至server');
//frmMain.Memo5.Lines.Add('nG_AryLength== ' + IntToStr(nG_AryLength));
  Application.ProcessMessages;
  if G_ClientArray = nil then
    nG_AryLength := 1
  else
    Inc(nG_AryLength);


  setLength(G_ClientArray,nG_AryLength);
  G_ClientArray[nG_AryLength-1] := TSimpleClient.Create;
  { Create a client object }
  Application.ProcessMessages;


  G_ClientArray[nG_AryLength-1].IP := AThread.Connection.Socket.Binding.PeerIP;
//JACKAL  G_ClientArray[nG_AryLength-1].ListLink := lbClients.Items.Count;
  { Assign the thread to it for ease in finding }
  G_ClientArray[nG_AryLength-1].Thread := AThread;
  { Add to our clients list box }
  { Assign it to the thread so we can identify it later }
  AThread.Data := G_ClientArray[nG_AryLength-1];
  { Add it to the clients list }
end;

procedure TListerner.IdTcpServerDisconnect(AThread: TIdPeerThread);
begin

end;

procedure TListerner.IdTcpServerExecute(AThread: TIdPeerThread);
var
  Client: TSimpleClient;
  Com, // System command
  sL_Msg,sL_CompName,sL_Content,sL_ClientIP,sL_TempClientIP : string;
  ii,nL_ClientIndex : Integer;
begin
  { Get the text sent from the client }
  try
    if nG_AryLength>0 then
    begin
//frmMain.Memo5.Lines.Add('Execute== ' + IntToStr(nG_AryLength));
      if not frmMain.chbAction.Checked then
        frmMain.chbAction.Checked := true;

//frmMain.Memo5.Lines.Add('=========');
      Client := Pointer(AThread.Data);
      sL_ClientIP := Client.IP;

      for ii:=0 to nG_AryLength-1 do
      begin
        sL_TempClientIP := G_ClientArray[ii].IP;
        if sL_TempClientIP=sL_ClientIP then
          nL_ClientIndex := ii;
//frmMain.Memo5.Lines.Add('ClientIP= ' + sL_TempClientIP);
      end;

      sL_Msg := AThread.Connection.ReadLn;
      if sL_Msg=''then Exit;
      //AThread.Connection.WriteLn('收到您傳送之訊息:' + sL_Msg);
      if AnsiPos('<' + ROOT_NODE_COMMAND_NAME + '>', sL_Msg)>0 then //表示有Command
      begin
        frmMain.handleXMLData(sL_Msg,sL_ClientIP);
      end
      else if AnsiPos('<' + ROOT_NODE_INFO_NAME + '>', sL_Msg)>0 then //表示有訊息
      begin
        frmMain.parseInfoXml(sL_Msg,sL_CompName,sL_Content);

        if sL_Content=SIMPLE_CA_CONNECTION_OK then  //簡易CA程式連接成功
        begin
          frmMain.memErrorLog.Lines.Add(sL_CompName + '--簡易CA程式連接--[IP:' + sL_ClientIP + ']-- '  + DatetimeToStr(now));

          G_ClientArray[nL_ClientIndex].CompName := sL_CompName;
        end
        else if sL_Content=SIMPLE_CA_DISCONNECT then  //簡易CA程式斷線
        begin
          //Disconnect此IP的連線
          clientDisconnect(AThread);

          //Retrieve Client Record from Data pointer
          frmMain.memErrorLog.Lines.Add(sL_CompName + '--簡易CA程式斷線--[IP:' + sL_ClientIP + ']-- '  + DatetimeToStr(now));

        end;
      end;
    end;
  except

  end;

end;

function TListerner.sendReturnStringToSimpleCA(sI_ClientIP,sI_ReturnString: String) : String;
var
    ii : Integer;
    L_ClientArray : TSimpleClient;
    AThread : TIdPeerThread;
begin
{
    L_ClientArray := TSimpleClient.Create;

    L_ClientArray.IP := sI_ClientIP;
    L_ClientArray.Thread := AThread;
    AThread.Data := L_ClientArray;
    TIdPeerThread(L_ClientArray.Thread).Connection.WriteLn(sI_ReturnString);
}

    for ii:=0 to high(G_ClientArray) do
    begin
      if (G_ClientArray[ii]<>nil) AND (G_ClientArray[ii].IP = sI_ClientIP) then
      begin
        TIdPeerThread(G_ClientArray[ii].Thread).Connection.WriteLn(sI_ReturnString);
        break;
      end;
    end;

    Result := '';
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

procedure TListerner.UpdateBindings;
var
  Binding: TIdSocketHandle;
  nL_Port : Integer;
begin
   nL_Port := StrToIntDef(sG_ListenPort,0);
  { Set the TIdTCPServer's port to the chosen value }
  G_IdTcpListener.DefaultPort := nL_Port;
  { Remove all bindings that currently exist }
  G_IdTcpListener.Bindings.Clear;
  { Create a new binding }
  Binding := G_IdTcpListener.Bindings.Add;
  { Assign that bindings port to our new port }
  Binding.Port := nL_Port;
end;

end.
