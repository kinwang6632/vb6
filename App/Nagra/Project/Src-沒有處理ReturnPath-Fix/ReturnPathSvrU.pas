unit ReturnPathSvrU;

interface

uses
  Windows, SysUtils, Classes, Forms, ScktComp, Winsock, CsIntf,
  ConstObjU, SendCommandU;

type
  ReturnPathSvr = class(TThread)
  private
  protected
    procedure Execute; override;
  public
    bG_BuildConnection : boolean;
    sG_ErrorCode,sG_ErrorMsg : String;
    G_DeviceData : TDeviceData;
    constructor Create;overload;
    function isValidCmd(sI_Cmd:String):boolean;
    procedure sendCmd(sI_CmdID: String);
    procedure ReceiveReturnPathData;
    procedure buildConnection(aClientSocket : TClientSocket; var I_WinSocketStream : TWinSocketStream);
  end;

implementation

uses frmMainU, UdateTimeu, Ustru;

{ ReturnPathSvr }

{ ---------------------------------------------------------------------------- }

procedure ReturnPathSvr.BuildConnection(aClientSocket: TClientSocket;
  var I_WinSocketStream: TWinSocketStream);
const
  AObjName: String = 'SMSGW';
var
  AIndex, AWriteSize: Integer;
  AWStream: TWinSocketStream;
  ADeviceCall : TDeviceCall;
  ADeviceAnswer : TDeviceAnswer;
  AString: String;
begin

  try
    AClientSocket.Address := frmMain.getSendCommandServerIP;
    AClientSocket.Port := frmMain.getReturnPathPort;
    AClientSocket.Open;

    ZeroMemory( @ADeviceCall, SizeOf( ADeviceCall ) );
    ZeroMemory( @ADeviceAnswer, SizeOf( ADeviceAnswer ) );

    ADeviceCall.op_mode := 0;

    for AIndex := 1 to Length( AObjName ) do
      ADeviceCall.obj_name[AIndex] := AObjName[AIndex];

    ADeviceCall.obj_name_len := Length( AObjName );

    ADeviceCall.len[1] := 0;

    ADeviceCall.len[2] := Length( AObjName ) +
      SizeOf( ADeviceCall.op_mode ) + SizeOf( ADeviceCall.obj_name_len );

    AWriteSize := Length( AObjName ) +
      SizeOf( ADeviceCall.op_mode ) + SizeOf( ADeviceCall.obj_name_len ) +
      SizeOf( ADeviceCall.len );

    AWStream := TWinSocketStream.Create( AClientSocket.Socket, 10000 );
    AWStream.WriteBuffer( ADeviceCall, AWriteSize );
    Sleep( 3000 );
    AWStream.ReadBuffer( ADeviceAnswer, SizeOf( ADeviceAnswer ) );

    if ( ADeviceAnswer.status <> 6 ) or
       ( ADeviceAnswer.code <> 0 ) then
      raise Exception.Create( 'Nagra reject the connection.' );


    frmMain.G_ReturnPathWStream := AWStream;

    SendCmd( '1002' );
    frmMain.bG_HasBuildReturnPathConnection := True;
    
    except
      sG_ErrorCode := BUILD_CONNECTION_ERROR;
      sG_ErrorMsg := BUILD_CONNECTION_ERROR_MSG;
      Exit;
    end;
end;

{ ---------------------------------------------------------------------------- }

constructor ReturnPathSvr.Create;
begin
  inherited Create( True );
  Self.Priority := tpNormal;
  Self.FreeOnTerminate := False;
  Self.OnTerminate := frmMain.ReturnPathSvrThreadDone;
  Self.Resume;
end;

{ ---------------------------------------------------------------------------- }

procedure ReturnPathSvr.Execute;
begin
  while not Self.Terminated do
  begin
    try
      Sleep( 100 );
      if ( not Self.Terminated ) and
         ( Assigned( frmMain.G_ReturnPathWStream ) ) and
         ( frmMain.ReturnPathClntSocket.Active ) then
      begin
         if ( frmMain.bG_HasBuildReturnPathConnection ) and
            ( frmMain.bG_ReturnPathSvrThread ) then
         begin
           ReceiveReturnPathData;
         end;
      end;
    except
      on E: Exception do
      begin
        CodeSite.SendError( 'Recv Return Data thread inner loop Error: %s',
          [E.Message]  );
      end;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function ReturnPathSvr.isValidCmd(sI_Cmd: String): boolean;
begin
  //判斷是否是正確的指令格式
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

procedure ReturnPathSvr.ReceiveReturnPathData;
var
  AStrList: TStringList;
  AData: String;
  AReadBytes, AIndex: Integer;
  ASocketStream : TWinSocketStream;
  ARecvAry: array [0..1024] of Char;
begin

    if frmMain.G_ReturnPathWStream = nil then Exit;
    ASocketStream := frmMain.G_ReturnPathWStream;
    FillChar( ARecvAry, Sizeof( ARecvAry ), 0 );
     { give the client 10 seconds to start writing }
     if ASocketStream.WaitForData( 10000 ) then
     begin
       if Self.Terminated then Exit;
       AReadBytes := ASocketStream.Read( ARecvAry, SizeOf( ARecvAry ) );
       if AReadBytes > 0 then
       begin
         CodeSite.SendMsg( Format( 'RecvCallBack Data Bytes:%d',
           [AReadBytes] ) );
         AData := '';
         AData := frmMain.transReceivedDataCharAryToStr2(
           ARecvAry, AReadBytes );
         AStrList := TUstr.ParseStrings( AData, RECEIVED_DATA_SEP_FLAG );
         try
           for AIndex := 0 to AStrList.Count - 1 do
           begin
             CodeSite.SendMsg(
               Format( 'RecvCallBack Data: %s', [AStrList.Strings[AIndex]] ) );
             frmMain.G_ReturnPathDataStrList.Add(AStrList.Strings[AIndex] );
           end;
         finally
           AStrList.Free;
         end;
       end;
     end
     { 沒收到資料 }
     else begin
      if Self.Terminated then Exit;
      Self.SendCmd( '1002' );
     end;
end;

{ ---------------------------------------------------------------------------- }

procedure ReturnPathSvr.SendCmd(sI_CmdID: String);
var
  aIndex: Integer;
  sL_FullCommand, sL_CommandRootHeader, sL_CommandBody: String;
  sL_CommandNo, sL_CommandType, sL_SmsCommandRootHeader, sL_CurrentDate: String;
  L_TmpRecord : TTempRecord;
  L_DeviceData : TDeviceData;
  AObj: TCmdInfoObject;
  sL_RealTransID : String;
begin
    try
        sL_CommandNo := sI_CmdID;
        sL_CommandType := SendCommandU.getCommandType( StrToInt( sL_CommandNo ) );

        sL_SmsCommandRootHeader := SendCommandU.getSmsCommandRootHeader( False, 0,
          StrToInt( sL_CommandNo ), sL_CommandType, EmptyStr );

        sL_CurrentDate := TUdateTime.GetPureDateStr( Date, EmptyStr );
        sL_CommandRootHeader := SendCommandU.getCommandSection(
          sL_CommandType, sL_CurrentDate, sL_CurrentDate, EmptyStr );


        { 這個 TCmdInfoObject 是為了傳參數用的, 沒有用途 }
        AObj := TCmdInfoObject.Create;
        try
          sL_CommandBody := SendCommandU.getCommandBody(
            AObj, StrToInt( sL_CommandNo ), nil, EmptyStr, EmptyStr, '0' );
        finally
          AObj.Free;
        end;

        sL_FullCommand :=
          sL_SmsCommandRootHeader + sL_CommandRootHeader + sL_CommandBody;

        if ( frmMain.nG_FD_TestingCmdSeqNo2 >= frmMain.nG_FD_MaxTestingCmdSeqNo2 ) then
          frmMain.nG_FD_TestingCmdSeqNo2 := 1
        else
          Inc( frmMain.nG_FD_TestingCmdSeqNo2 );  

        sL_RealTransID :=
          TUstr.AddString( TESTING_CMD_COMP_CODE2, '0', False, 3 ) +
          TUstr.AddString( IntToStr( frmMain.nG_FD_TestingCmdSeqNo2 ),'0', False,
          SEQ_NUM_LENGTH );

        sL_FullCommand := sL_RealTransID + sL_FullCommand;

        L_TmpRecord := frmMain.GetBinaryVal( Length( sL_FullCommand ), 2, False );
        L_DeviceData.len[0] := L_TmpRecord.TempAry[0];
        L_DeviceData.len[1] := L_TmpRecord.TempAry[1];

        FillChar( L_DeviceData.data, Length( L_DeviceData.data ) + 1, 0 );

        for aIndex := 1 to Length( sL_FullCommand ) do
          L_DeviceData.data[aIndex - 1] := sL_FullCommand[aIndex];

        frmMain.G_ReturnPathWStream.Write( L_DeviceData,  Length( sL_FullCommand ) + 2 );

    except

    end;
end;

{ ---------------------------------------------------------------------------- }

end.
