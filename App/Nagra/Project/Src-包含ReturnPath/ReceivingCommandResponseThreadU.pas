//Syn OK
unit ReceivingCommandResponseThreadU;

interface

uses
  Classes, forms, comctrls, syncobjs, dbtables, Sysutils, Dialogs,
  ConstObjU, Ustru, UdateTimeu, scktcomp, winsock, inifiles, ADODB,
  CsIntf;


type

  ReceivingCommandResponseThread = class(TThread)
  private
    { Private declarations }

  protected
    procedure Execute; override;
  public
//    nG_Times : Integer;
    sG_TestInfo : String;
    G_DeviceData : TDeviceData;
    procedure receiveResponse;
    

    procedure AddResponseMsg(I_DeviceData : TDeviceData); overload;
    procedure AddResponseMsg; overload;
  end;

implementation

uses frmMainU, HandleResponseU;

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
    if Assigned(frmMain.G_ResponseMsgStrList) then
      frmMain.G_ResponseMsgStrList.Add(G_DeviceData.data);
end;

procedure ReceivingCommandResponseThread.AddResponseMsg(
  I_DeviceData: TDeviceData);
begin
    if Assigned(frmMain.G_ResponseMsgStrList) then
      frmMain.G_ResponseMsgStrList.Add(I_DeviceData.data);
end;

procedure ReceivingCommandResponseThread.Execute;
var
    ii : Integer;
    L_ErrorInfoStrList : TStringList;
    sL_ErrorInfoFileName : String;
begin
    try

      self.FreeOnTerminate := True;
      self.OnTerminate := frmMain.ReceiveResponseThreadDone;

      
       FillChar( G_DeviceData.data, SizeOf( G_DeviceData.data ), 0 );
      
      while true do
      begin
        Sleep( 50 );
        try
          if (not Terminated) and Assigned(frmMain.G_SendCmdClntSocket) and (frmMain.G_SendCmdClntSocket.Active) then
          begin

            if frmMain.bG_HasBuildSendCmdConnection and frmMain.bG_RunReceivingCommandThread and Assigned(frmMain.G_SendCmdWStream) then
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
        Application.ProcessMessages;
      end;
    finally

    end;
end;

procedure ReceivingCommandResponseThread.receiveResponse;
var
    nL_WaitForDataTimeOut, ii,nL_Words : Integer;
    sL_ResponseString, sL_RealData : String;
    L_ErrorInfoStrList : TStringList;
    sL_ErrorInfoFileName : String;
    L_StrList : TStringList;
    L_TmpAry : array[0..1024] of Char;    
//    L_DeviceData : TDeviceData;
    procedure reBuildConn(nI_RaiseType:Integer; sI_Msg:String);
    var
        sL_LogInfo : String;
    begin
      //down, 重新建立 connection...

      case nI_RaiseType of
        1:
          sL_LogInfo := ' ReceivingCommandResponseThread.receiveResponse' + '-- test info =>' + sG_TestInfo + '-- WaitForData timeout...' ;

        2:
          sL_LogInfo := ' ReceivingCommandResponseThread.receiveResponse.' + '-- test info =>' + sG_TestInfo+ '--錯誤訊息:' + sI_Msg;
      end;
      frmMain.showLogOnMemo(REBUILD_SEND_CMD_CONNECTION_TO_NAGRA,sL_LogInfo);

      frmMain.reBuildNagraConn;      
      //up, 重新建立 connection...

    end;

begin

    try

      FillChar( L_TmpAry, Sizeof( L_TmpAry ), 0 );

      nL_WaitForDataTimeOut := frmMain.reDefineFetchDataInterval(now) + SOCKET_WAIT_FOR_DATA_TIMEOUT*1000;
      if frmMain.G_SendCmdWStream.WaitForData( nL_WaitForDataTimeOut ) then
      begin

        if not Assigned(frmMain.G_SendCmdWStream) then Exit;
        nL_Words := frmMain.G_SendCmdWStream.Read(L_TmpAry, sizeof(L_TmpAry)); //這種寫法會使得傳送速度變得很快, 但接收速度變慢
//      nL_Words := recv(frmMain.G_SendCmdClntSocket.Socket.SocketHandle , G_DeviceData, sizeof(G_DeviceData),0);

        if nL_Words>0 then
        begin
          sL_RealData := '';
          sL_RealData := frmMain.transReceivedDataCharAryToStr(L_TmpAry, nL_Words);

          L_StrList := TUstr.ParseStrings(sL_RealData,RECEIVED_DATA_SEP_FLAG);
          for ii:=0 to L_StrList.Count -1 do
          begin
            frmMain.G_ResponseMsgStrList.Add(L_StrList.Strings[ii]);

          end;

        end;
        {
        if Assigned(frmMain.G_ResponseMsgStrList) then
          frmMain.G_ResponseMsgStrList.Add(L_DeviceData.data);
        }
//        AddResponseMsg(L_DeviceData);
//      Synchronize(AddResponseMsg);
      {
      frmMain.Memo1.Lines.Add(L_DeviceData.data);
      }



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
        else if AnsiPos('violation', E.Message) >0 then
        begin
        end  
        else
          frmMain.showLogOnMemo(RECEIVE_CMD_RESPONSE_ERROR_FLAG,' ReceivingCommandResponseThread.receiveResponse.程式即將結束,然後重新啟動.' + '-- test info =>' + sG_TestInfo+ '--錯誤訊息:' + E.Message);

      end;
    end;
end;


end.
