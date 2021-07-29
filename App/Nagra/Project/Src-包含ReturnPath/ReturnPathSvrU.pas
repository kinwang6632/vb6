unit ReturnPathSvrU;

interface

uses
  Classes, ScktComp, winsock, ConstObjU, SysUtils, Forms;

type
  ReturnPathSvr = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  public
    nG_60003TransID, nG_Max60003TransID : Integer;
    bG_BuildConnection : boolean;
    sG_ErrorCode,sG_ErrorMsg : String;
    G_DeviceData : TDeviceData;
    function isValidCmd(sI_Cmd:String):boolean;
    procedure sendCmd(sI_CmdID, sI_SrcTransNum:String);
    procedure AddReturnPathData;overload;
    procedure AddReturnPathData(sI_CallbackData:String);overload;
    procedure AddReturnPathData(I_DeviceData: TDeviceData);overload;
    procedure receiveReturnPathData;             
    constructor Create;overload;
    function Increment(Ptr: Pointer; Increment: Integer): Pointer;    
    function Write(I_Buf: Pointer; I_Count: Integer; I_WStream: TWinSocketStream):Integer;    
    procedure buildConnection(I_ClientSocket : TClientSocket; var I_WinSocketStream : TWinSocketStream);
  end;

implementation

uses frmMainU, UdateTimeu, Ustru;

{ Important: Methods and properties of objects in VCL or CLX can only be used
  in a method called using Synchronize, for example,

      Synchronize(UpdateCaption);

  and UpdateCaption could look like,

    procedure ReturnPathSvr.UpdateCaption;
    begin
      Form1.Caption := 'Updated in a thread';
    end; }

{ ReturnPathSvr }

procedure ReturnPathSvr.AddReturnPathData;
var
    L_ReturnPathDataObj : TReturnPathDataObj;
begin
    if G_DeviceData.data='' then
      Exit;
//    L_ReturnPathDataObj := TReturnPathDataObj.Create;
//    frmMain.Memo1.Lines.add(G_DeviceData.data);
//    L_ReturnPathDataObj.Data := G_DeviceData.data;
    frmMain.G_ReturnPathDataStrList.Add(G_DeviceData.data);
//    frmMain.G_ReturnPathDataStrList.AddObject('N', L_ReturnPathDataObj);

end;

procedure ReturnPathSvr.AddReturnPathData(I_DeviceData: TDeviceData);
var
    L_ReturnPathDataObj : TReturnPathDataObj;
begin
    if I_DeviceData.data='' then
      Exit;

    L_ReturnPathDataObj := TReturnPathDataObj.Create;
    L_ReturnPathDataObj.Data := I_DeviceData.data;
    frmMain.G_ReturnPathDataStrList.AddObject('N', L_ReturnPathDataObj);

end;

procedure ReturnPathSvr.AddReturnPathData(sI_CallbackData: String);
begin
    frmMain.G_ReturnPathDataStrList.Add(sI_CallbackData);

end;

procedure ReturnPathSvr.buildConnection(I_ClientSocket: TClientSocket;
  var I_WinSocketStream: TWinSocketStream);
var
    nL_Len, nL_Words : Integer;
    L_TmpRecord : TTempRecord;
    ii, nL_CmdLen : Integer;
    L_Socket : TClientSocket;

    L_DeviceCall : TDeviceCall;
    L_DeviceStatus : TDeviceStatus;
    L_DeviceAnswer : TDeviceAnswer;
    sL_CmdObjName : String;
    sL_CurrentDate,sL_BoxNo, sL_CommandNo, sL_CommandType, sL_SmsCommandRootHeader : String;
    sL_FullCommand, sL_IccNo, sL_CommandBody, sL_CommandRootHeader : String;
    L_DeviceData : TDeviceData;
begin

    bG_BuildConnection := false;

    {
    if Assigned(I_ClientSocket) then
    begin
      I_ClientSocket := nil;
      I_ClientSocket.Free;
    end;
    

    I_ClientSocket:= TClientSocket.Create(nil);
    I_ClientSocket.ClientType := ctBlocking;
    }
    L_Socket := I_ClientSocket;
    try
    begin
      with L_Socket do
      begin
        Address := frmMain.getReturnPathServerIP;
        Port := frmMain.getReturnPathPort;

        Open;
        L_DeviceCall.op_mode := char(0);

        sL_CmdObjName := 'SMSGW';

        //L_DeviceCall.obj_name :=
        for ii:=1 to length(L_DeviceCall.obj_name) do
          L_DeviceCall.obj_name[ii-1] := char(0);

        for ii:=1 to length(sL_CmdObjName) do
        begin
          L_DeviceCall.obj_name[ii-1] := sL_CmdObjName[ii];
        end;

//        StrPCopy(L_DeviceCall.obj_name,sL_CmdObjName);
        L_DeviceCall.obj_name_len := char(length(sL_CmdObjName));


        for ii:=1 to length(L_DeviceCall.user_data) do
          L_DeviceCall.user_data[ii-1] := char(0);

        for ii:=1 to length(sL_FullCommand) do
        begin
          L_DeviceCall.user_data[ii-1] := sL_FullCommand[ii];
        end;

//        nL_CmdLen := 1 + 1 + length(sL_CmdObjName) + length(sL_FullCommand); //ok.. => ¨ä¹ê¤£½T©w­n¤£­n°e sL_CmdObjName
        nL_CmdLen := 1 + 1 + length(sL_CmdObjName); //ok.. => ¨ä¹ê¤£½T©w­n¤£­n°e sL_CmdObjName
//        nL_CmdLen := 1+1+ length(sL_FullCommand); //testing..

        L_TmpRecord := frmMain.getBinaryVal(nL_CmdLen,2,false);
        L_DeviceCall.len[0] := L_TmpRecord.TempAry[0];
        L_DeviceCall.len[1] := L_TmpRecord.TempAry[1];

//        L_WStream := TWinSocketStream.Create(L_Socket.Socket, nG_Timer*100); //howard
        I_WinSocketStream := TWinSocketStream.Create(L_Socket.Socket, 10);


        nL_Words := write(@L_DeviceCall, nL_CmdLen+2, I_WinSocketStream); //build connection
//            showmessage('return:'+inttostr(nL_Words));



        nL_Words := recv(L_Socket.Socket.SocketHandle , L_DeviceStatus, 3,0);
  //      showmessage('recv 1 ªº¶Ç¦^­È:'+inttostr(nL_Words));

        for ii:=1 to length(L_DeviceAnswer.data) do
          L_DeviceAnswer.data[ii-1] := char(0);

        nL_Words := recv(L_Socket.Socket.SocketHandle , L_DeviceAnswer, L_DeviceStatus.len,0);
  //      nL_Words := recv(L_Socket.Socket.SocketHandle , L_DeviceStatus, 3,0);
  //      showmessage('recv 2 ªº¶Ç¦^­È:'+inttostr(nL_Words));

//        showmessage('device_call L_DeviceAnswer.code==' + L_DeviceAnswer.code);
//        showmessage('device_call L_DeviceAnswer.data==' + L_DeviceAnswer.data);



        //down,¶Ç°e«ü¥O 1002

        frmMain.G_ProcessReturnPathData.sendCmd('1002','');
        {
        sL_CommandNo := '1002';//ª½±µ°e¥X«ü¥O 1002, ¥Î¥H½T»{ connection ¬O§_¶¶§Q«Ø¥ß => ±ý±µ¦¬ feedback command ªº process ,¤~»Ý­n°e¥X command 1002
        //        sL_CommandNo := sG_CommandNo;

        sL_CommandType := frmMain.G_SendCommandThread.getCommandType(StrToInt(sL_CommandNo));
        sL_SmsCommandRootHeader := frmMain.G_SendCommandThread.getSmsCommandRootHeader('', StrToInt(sL_CommandNo), sL_CommandType);
        //    showmessage(sL_SmsCommandRootHeader);
        sL_CurrentDate := TUdateTime.GetPureDateStr(date,'');
        sL_BoxNo := '';
        sL_IccNo := '';
        sL_CommandRootHeader := frmMain.G_SendCommandThread.getCommandSection(sL_CommandType, sL_CurrentDate, sL_CurrentDate,sL_IccNo);
        sL_CommandBody := frmMain.G_SendCommandThread.getCommandBody(StrToInt(sL_CommandNo), sL_BoxNo);
        sL_FullCommand := sL_SmsCommandRootHeader + sL_CommandRootHeader + sL_CommandBody;

        sL_FullCommand := '099000001' + sL_FullCommand;

        nL_Len := length(sL_FullCommand);//ok..
        L_TmpRecord := frmMain.GetBinaryVal(nL_Len,2,false);
        L_DeviceData.len[0] := L_TmpRecord.TempAry[0];
        L_DeviceData.len[1] := L_TmpRecord.TempAry[1];

        for ii:=1 to length(L_DeviceData.data) do
          L_DeviceData.data[ii-1] := char(0);
        for ii:=1 to length(sL_FullCommand) do
        begin
          L_DeviceData.data[ii-1] := sL_FullCommand[ii];
        end;
        nL_Words := write(@L_DeviceData, nL_Len+2, L_WStream); //send_data..
        }
        //up,¶Ç°e«ü¥O 1002

        frmMain.bG_HasBuildReturnPathConnection := true;
      end;
    end;
    except
      sG_ErrorCode := BUILD_CONNECTION_ERROR;
      sG_ErrorMsg := BUILD_CONNECTION_ERROR_MSG;
//      dtmMain.UpdateCMD_STATUS(frmMain.nG_CurActiveConn,sG_SeqNo,'E',sG_ErrorCode,sG_ErrorMsg);
//      writeErrorLog;

      exit;
    end;

    bG_BuildConnection := true;
end;

constructor ReturnPathSvr.Create;
begin
    inherited Create( True );
    //Priority := tpHighest;
    Self.Priority := tpNormal;
    Self.Resume;
end;

procedure ReturnPathSvr.Execute;
var
    L_DeviceData : TDeviceData;
    ii, nL_Words : Integer;
    L_ErrorInfoStrList : TStringList;
    sL_ErrorInfoFileName : String;
begin
    try

      self.FreeOnTerminate := true;
      self.OnTerminate := frmMain.ReturnPathSvrThreadDone;

      nG_60003TransID := 0;
      nG_Max60003TransID := StrToIntDef(TUstr.AddString('','9',true,SEQ_NUM_LENGTH),999999);

      FillChar(L_DeviceData.data, length(L_DeviceData.data)+1, 0); { initialize the buffer }
      while true do
      begin
        Sleep(50); { My Add, 93-07-30 }
        receiveReturnPathData;
      end;
    except
      on E: Exception do
      begin

        L_ErrorInfoStrList := TStringList.Create;
        sL_ErrorInfoFileName := 'c:\Error11.txt';

        if FileExists(sL_ErrorInfoFileName) then
          L_ErrorInfoStrList.LoadFromFile(sL_ErrorInfoFileName);
        if L_ErrorInfoStrList.Count>500 then
          L_ErrorInfoStrList.Clear;
        L_ErrorInfoStrList.Insert(0,'   Msg:' +E.Message +', HelpContext:' + IntToStr(E.HelpContext));
        L_ErrorInfoStrList.Insert(0,'ReturnPathSvr.Execute error:' + DateTimeToStr(now));
        L_ErrorInfoStrList.SaveToFile(sL_ErrorInfoFileName);
        L_ErrorInfoStrList.Free;
      end;
    end;
end;

function ReturnPathSvr.Increment(Ptr: Pointer;
  Increment: Integer): Pointer;
begin
    Result := Pointer(Longint(Ptr) + Increment);
end;

function ReturnPathSvr.isValidCmd(sI_Cmd: String): boolean;
begin
    //§PÂ_¬O§_¬O¥¿½Tªº«ü¥O®æ¦¡
    result := true;

end;




procedure ReturnPathSvr.receiveReturnPathData;
var
    L_StrList : TStringList;
    sL_RealData, sL_RealData1 : String;
    nL_Words, ii : Integer;
    Stream : TWinSocketStream;
    L_TmpAry : array[0..1024] of Char;
begin
  { make sure connection is active }


  while (not Terminated) and frmMain.ReturnPathClntSocket.Socket.Connected do
  begin
    try
      if frmMain.G_ReturnPathWStream=nil then continue;
      Stream := frmMain.G_ReturnPathWStream; 
//      Stream := TWinSocketStream.Create(frmMain.ReturnPathClntSocket.Socket, 60000);
      try
        FillChar(L_TmpAry, sizeof(L_TmpAry), 0); { initialize the buffer }
        { give the client 60 seconds to start writing }
        if Stream.WaitForData(60000) then
        begin

          nL_Words := Stream.Read(L_TmpAry, sizeof(L_TmpAry));
//          if Stream.Read(L_DeviceData, sizeof(L_DeviceData)) = 0 then { if can’t read in 60 seconds }
//            frmMain.ReturnPathClntSocket.Socket.Close;               { close the connection }
          { now process the request }
          if nL_Words>0 then
          begin
            sL_RealData := '';
            sL_RealData := frmMain.transReceivedDataCharAryToStr(L_TmpAry, nL_Words);

            L_StrList := TUstr.ParseStrings(sL_RealData,RECEIVED_DATA_SEP_FLAG);
            for ii:=0 to L_StrList.Count -1 do
            begin
              frmMain.G_ReturnPathDataStrList.Add(L_StrList.Strings[ii]);            
              frmMain.G_ReturnPathDataStrList1.Add(L_StrList.Strings[ii]);
            end;  
            {
            frmMain.G_ReturnPathDataStrList1.Add(inttostr(nL_Words) + '==' + sL_RealData);
            frmMain.G_ReturnPathDataStrList1.Add(inttostr(nL_Words) + '==' + sL_RealData1);
            }
  //          frmMain.Memo1.Lines.Add(L_DeviceData.data);
          end;  
        end;
{
        else
          frmMain.ReturnPathClntSocket.Socket.Close; { if client doesn’t start, close }

      finally
//        Stream.Free;
      end;
    except
      on E: Exception do
      begin
//        Stream.Free;
        frmMain.Memo1.Lines.Add('Error:' + E.Message);
      end;
    end;
  end;
end;
procedure ReturnPathSvr.sendCmd(sI_CmdID, sI_SrcTransNum: String);
var
    nL_Len, ii, nL_Words : Integer;
    sL_FullCommand: String;
    sL_CommandRootHeader, sL_CommandBody : String;
    sL_BoxNo, sL_IccNo : String;
    sL_CommandNo, sL_CommandType, sL_SmsCommandRootHeader, sL_CurrentDate : String;
    L_TmpRecord : TTempRecord;
    L_DeviceData : TDeviceData;

    sL_RealTransID : String;
begin
    try
//        Application.ProcessMessages;
        frmMain.G_SendCommandThread.sG_SrcTransNum := sI_SrcTransNum;
        sL_CommandNo := sI_CmdID;//ª½±µ°e¥X«ü¥O 1002, ¥Î¥H½T»{ connection ¬O§_¶¶§Q«Ø¥ß => ±ý±µ¦¬ feedback command ªº process ,¤~»Ý­n°e¥X command 1002
        //        sL_CommandNo := sG_CommandNo;
//        Application.ProcessMessages;
        sL_CommandType := frmMain.G_SendCommandThread.getCommandType(StrToInt(sL_CommandNo));
        sL_SmsCommandRootHeader := frmMain.G_SendCommandThread.getSmsCommandRootHeader(false, StrToInt(sL_CommandNo), sL_CommandType);
        //    showmessage(sL_SmsCommandRootHeader);
        sL_CurrentDate := TUdateTime.GetPureDateStr(date,'');
        sL_BoxNo := '';
        sL_IccNo := '';
        sL_CommandRootHeader := frmMain.G_SendCommandThread.getCommandSection(sL_CommandType, sL_CurrentDate, sL_CurrentDate,sL_IccNo);
        sL_CommandBody := frmMain.G_SendCommandThread.getCommandBody(StrToInt(sL_CommandNo), sL_BoxNo);
        sL_FullCommand := sL_SmsCommandRootHeader + sL_CommandRootHeader + sL_CommandBody;
//        frmMain.Memo4.Lines.Add('F1: ' +sL_FullCommand );

                
//        Application.ProcessMessages;

        if nG_60003TransID>=nG_Max60003TransID then
          nG_60003TransID := 1
        else
          INC(nG_60003TransID);

        sL_RealTransID := TUstr.AddString(PORT_60003_COMP_CODE,'0',false,3) +
                          TUstr.AddString(IntToStr(nG_60003TransID),'0',false,SEQ_NUM_LENGTH);
        sL_FullCommand := sL_RealTransID + sL_FullCommand;

        frmMain.Memo4.Lines.Add('F2: ' +sL_FullCommand );
        nL_Len := length(sL_FullCommand);//ok..
//frmMain.Memo3.Lines.Add(sL_FullCommand);
        L_TmpRecord := frmMain.GetBinaryVal(nL_Len,2,false);
        L_DeviceData.len[0] := L_TmpRecord.TempAry[0];
        L_DeviceData.len[1] := L_TmpRecord.TempAry[1];
//        Application.ProcessMessages;
        FillChar(L_DeviceData.data, length(L_DeviceData.data)+1, 0); { initialize the buffer }
        {
        for ii:=1 to length(L_DeviceData.data) do
          L_DeviceData.data[ii-1] := char(0);
        }
        for ii:=1 to length(sL_FullCommand) do
        begin
          L_DeviceData.data[ii-1] := sL_FullCommand[ii];
        end;
//        Application.ProcessMessages;
        nL_Words := write(@L_DeviceData, nL_Len+2, frmMain.G_ReturnPathWStream); //send_data..
    except

    end;
end;

function ReturnPathSvr.Write(I_Buf: Pointer; I_Count: Integer;
  I_WStream: TWinSocketStream): Integer;
var
    nL_Idx: Integer;
begin
  nL_Idx := 0;
  while (nL_Idx < I_Count) do
  try
    nL_Idx := nL_Idx + I_WStream.Write(Increment(I_Buf, nL_Idx)^, I_Count-nL_Idx);
  except
    buildConnection(frmMain.ReturnPathClntSocket, frmMain.G_ReturnPathWStream);
    Write(I_Buf, I_Count,frmMain.G_ReturnPathWStream);
//    raise;
  end;

  result := nL_Idx;
end;

end.
