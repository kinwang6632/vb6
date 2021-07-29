//Syn OK
unit ReceivingCommandResponseThreadU;

interface

uses
  Classes, forms, comctrls, syncobjs, dbtables, Sysutils, Dialogs,
  Ustru, UdateTimeu, scktcomp, winsock, inifiles, ADODB;


type

  ReceivingCommandResponseThread = class(TThread)
  private
    { Private declarations }

  protected
    procedure Execute; override;
  public
//    nG_Times : Integer;


    procedure receiveResponse;
    

//    procedure AddResponseMsg(I_DeviceData : TDeviceData); overload;
    procedure AddResponseMsg; overload;
  end;

implementation

uses frmMainU, ConstU;

{ Important: Methods and properties of objects in VCL or CLX can only be used
  in a method called using Synchronize, for example,

      Synchronize(UpdateCaption);

  and UpdateCaption could look like,

    procedure ReceivingCommandResponseThread.UpdateCaption;
    begin
      Form1.Caption := 'Updated in a thread';
    end; }

{ ReceivingCommandResponseThread }

procedure ReceivingCommandResponseThread.AddResponseMsg;
var
    sL_ResponseString : String;
begin
//    sL_ResponseString := G_DeviceData.data;
{
    if Assigned(frmMain.G_ResponseMsgStrList) then
      frmMain.G_ResponseMsgStrList.Add(G_DeviceData.data);
}      
end;

{
procedure ReceivingCommandResponseThread.AddResponseMsg(
  I_DeviceData: TDeviceData);
begin
    frmMain.handleCmdResponse(I_DeviceData.data);

//    if Assigned(frmMain.G_ResponseMsgStrList) then
//      frmMain.G_ResponseMsgStrList.Add(I_DeviceData.data);

end;
}
procedure ReceivingCommandResponseThread.Execute;
var
    ii : Integer;
    L_ErrorInfoStrList : TStringList;
    sL_ErrorInfoFileName : String;
begin
    try
//      nG_Times := 0;
      self.FreeOnTerminate := true;
//      frmMain.G_ReceivingCommandResponseThread.OnTerminate := frmMain.ReceiveResponseThreadDone;
      self.OnTerminate := frmMain.ReceiveResponseThreadDone;

//      FillChar(G_DeviceData.data, length(G_DeviceData.data)+1, 0); { initialize the buffer }

      {
      for ii:=1 to length(G_DeviceData.data) do
        G_DeviceData.data[ii-1] := char(0);
      }
      while true do
      begin
        try
          if (not Terminated) and Assigned(frmMain.G_SendCmdClntSocket) and (frmMain.G_SendCmdClntSocket.Active) and (Assigned(frmMain.G_SendCmdWStream))then
          begin
            Application.ProcessMessages;
            receiveResponse;
          end;
        except
          on E: Exception do
          begin

            L_ErrorInfoStrList := TStringList.Create;
            sL_ErrorInfoFileName := 'c:\Error3.txt';

            if FileExists(sL_ErrorInfoFileName) then
              L_ErrorInfoStrList.LoadFromFile(sL_ErrorInfoFileName);
            if L_ErrorInfoStrList.Count>500 then
              L_ErrorInfoStrList.Clear;
            L_ErrorInfoStrList.Insert(0,'   Msg:' +E.Message +', HelpContext:' + IntToStr(E.HelpContext));
            L_ErrorInfoStrList.Insert(0,'ReceivingCommandResponseThread.Execute error:' + DateTimeToStr(now));
            L_ErrorInfoStrList.SaveToFile(sL_ErrorInfoFileName);
            L_ErrorInfoStrList.Free;
          end;
        end;
      end;
    finally

    end;
end;

procedure ReceivingCommandResponseThread.receiveResponse;
var
    nL_WaitForDataTimeOut, ii,nL_Words : Integer;
    sL_ResponseString : String;
    L_ErrorInfoStrList : TStringList;
    sL_ErrorInfoFileName : String;

//    L_DeviceData : TDeviceData;
    procedure reBuildConn(nI_RaiseType:Integer; sI_Msg:String);
    var
        sL_LogInfo : String;
    begin
      //down, 重新建立 connection...

      case nI_RaiseType of
        1:
          sL_LogInfo := ' ReceivingCommandResponseThread.receiveResponse' + '-- WaitForData timeout...' ;

        2:
          sL_LogInfo := ' ReceivingCommandResponseThread.receiveResponse.' + '--錯誤訊息:' + sI_Msg;
      end;
      frmMain.showLogOnMemo(REBUILD_SEND_CMD_CONNECTION_TO_NAGRA,sL_LogInfo);

      frmMain.reBuildNagraConn;      
      //up, 重新建立 connection...

    end;

begin
    {
    if (frmMain.ClientSocket1=nil) or (frmMain.ClientSocket1.Socket=nil)or
       (frmMain.ClientSocket1.Socket.SocketHandle<=0) then Exit;
    }
//    if not frmMain.bG_HasBuildSendCmdConnection then Exit; //0304
    try

//    while true do
//    begin
//      sL_ResponseString := '';
//      nL_Words := recv(frmMain.G_SendCmdClntSocket.Socket.SocketHandle , L_DeviceData, sizeof(L_DeviceData),0);

//      AddResponseMsg(L_DeviceData);

//      Inc(nG_Times);
//      if frmMain.G_SendCmdWStream.WaitForData(SOCKET_WAIT_FOR_DATA_TIMEOUT*1000) and (nG_Times<2)then
      nL_WaitForDataTimeOut := frmMain.reDefineFetchDataInterval(now) + SOCKET_WAIT_FOR_DATA_TIMEOUT*1000;
      if frmMain.G_SendCmdWStream.WaitForData(nL_WaitForDataTimeOut) then
      begin
//        nL_Words := frmMain.G_SendCmdWStream.Read(L_DeviceData, sizeof(L_DeviceData)); //這種寫法會使得傳送速度變得很快, 但接收速度變慢
//      nL_Words := recv(frmMain.G_SendCmdClntSocket.Socket.SocketHandle , G_DeviceData, sizeof(G_DeviceData),0);

//      AddResponseMsg;

//        frmMain.handleCmdResponse(L_DeviceData.data);
//        AddResponseMsg(L_DeviceData);
//      Synchronize(AddResponseMsg);
      {
      frmMain.Memo1.Lines.Add(L_DeviceData.data);
      }


      {
      if L_DeviceData.data<>'' then   //0304
      begin
        frmMain.Memo1.Lines.Add(L_DeviceData.data);
        AddResponseMsg(L_DeviceData);
      }//0304
//        Synchronize(AddResponseMsg);
{0304
        for ii:=1 to length(G_DeviceData.data) do
          G_DeviceData.data[ii-1] := char(0);
}
//0304      end;

      end
      else
      begin

//        nG_Times := 0;
        reBuildConn(1,'');

        //down, 重新啟動 AP...
//        frmMain.showLogOnMemo(SOCKET_NOT_READY_FOR_READING_FLAG,' socket 尚未準備好接收回應.程式即將結束,然後重新啟動.' + '-- test info =>' + sG_TestInfo + '-- WaitForData timeout...' );
        //up, 重新啟動 AP...
        Exit;
      end;
//      raise Exception.Create('abc');
    except
      on E: Exception do
      begin

        sL_ErrorInfoFileName := 'c:\Error1.txt';
        L_ErrorInfoStrList := TStringList.Create;
        if FileExists(sL_ErrorInfoFileName) then
          L_ErrorInfoStrList.LoadFromFile(sL_ErrorInfoFileName);
        if L_ErrorInfoStrList.Count>500 then
          L_ErrorInfoStrList.Clear;
            L_ErrorInfoStrList.Insert(0,'   Msg:' +E.Message +', HelpContext:' + IntToStr(E.HelpContext));
        L_ErrorInfoStrList.Insert(0,'ReceivingCommandResponseThread.receiveResponse error:' + DateTimeToStr(now));
        L_ErrorInfoStrList.SaveToFile(sL_ErrorInfoFileName);
        L_ErrorInfoStrList.Free;
//        if AnsiPos('指定的網路名稱無法使用', E.Message) >0 then
        if AnsiPos('網路名稱無法使用', E.Message) >0 then
          reBuildConn(2, E.Message)
        else
          frmMain.showLogOnMemo(RECEIVE_CMD_RESPONSE_ERROR_FLAG,' ReceivingCommandResponseThread.receiveResponse.程式即將結束,然後重新啟動.' + '--錯誤訊息:' + E.Message);

      end;
    end;
end;


end.
