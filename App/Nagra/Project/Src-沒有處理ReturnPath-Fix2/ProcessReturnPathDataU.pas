unit ProcessReturnPathDataU;

interface

uses
  Classes, SysUtils, ConstObjU, Forms, ScktComp, CsIntf, SendCommandU;

type
  ProcessReturnPathData = class(TThread)
  private
    nG_60003TransID: Integer;
    nG_Max60003TransID: Integer;
    sG_ReturnPathData: String;
    sInfo: string;
  protected
    procedure Execute; override;
  public
    constructor Create; overload;
    procedure PreSendResponse(const AResponseData:String);
    procedure PreParseReturnPathData;
    procedure ParseReturnPathData;
    procedure SendCmd(const sI_CmdID, sI_SrcTransNum:String);
    function IsValidCmd(const sI_Cmd:String): Boolean;
    procedure UpdateUI;  
  end;
                                       
implementation

uses frmMainU, dtmMainU, UdateTimeu, Ustru;

{ ProcessReturnPathData }

{ ---------------------------------------------------------------------------- }

constructor ProcessReturnPathData.Create;
begin
  inherited Create( True );
  Self.Priority := tpLower;
  Self.FreeOnTerminate := False;
  Self.OnTerminate := frmMain.ProcessReturnPathDataThreadDone;
  nG_60003TransID := 0;
  nG_Max60003TransID :=
    StrToIntDef( TUstr.AddString( '', '9', True, SEQ_NUM_LENGTH ), 999999 );
  Self.Resume;
end;

{ ---------------------------------------------------------------------------- }

procedure ProcessReturnPathData.Execute;
begin
  try
    while not Self.Terminated do
    begin
      try
        Sleep( 100 );
        if frmMain.bG_ProcessReturnPathDataThread then
        begin
          PreParseReturnPathData;
        end;
      except
        on E: Exception do
        begin
          CodeSite.SendError( 'Process Return Data thread inner loop Error: %s',
            [E.Message]  );
        end;
      end;
    end;
  except
    on E: Exception do
    begin
     CodeSite.SendError( 'Process Return Data thread Outer loop Error: %s',
       [E.Message]  );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function ProcessReturnPathData.IsValidCmd(const sI_Cmd: String): boolean;
begin
  //判斷是否是正確的指令格式
   Result := True;
end;

{ ---------------------------------------------------------------------------- }

procedure ProcessReturnPathData.ParseReturnPathData;
var
    sL_NumOfIppv, sL_PurchaseDate, sL_WatchedStatus : String;
    sL_FbIccNo, sL_FbCmdID, sL_StbNO, sL_StuNO : String;
    sL_Credit, sL_Debit, sL_Responding, sL_ImsProductID : String;
    nL_FbCmdID : Integer;
    sL_SmsEventID, sL_SmsEventID1, sL_SmsEventID2 : String;
    sL_PhoneNum1, sL_PhoneNum2, sL_PhoneNum3, sL_AbnormalPhone : String;
    sL_ErrorCode, sL_ErrorCodeExt, sL_StuCallbackDate, sL_StuCallbackTime : String;

    sL_PreviousStartDate, sL_PreviousStartTime, sL_PreviousStartDuration : String;
    sL_NewStartDate, sL_NewStartTime, sL_NewStartDuration : String;
    nL_NbOfProducts, ii : Integer;
    sL_OriginalTransactionNum, sL_IccSuspended, sL_NbOfProducts : String;
    nL_TmpBeginNum1, nL_TmpBeginNum2 : Integer;
    sL_ImsProductIdList, sL_ProductSuspendList : String;
    sL_CmdBody, sL_CmdHeader : String;

    sL_TransactionNo, sL_CmdType, sL_SrcID, sL_DestID, sL_MoppID, sL_CreationDate : String;

    nL_CmdID : Integer;
begin

    sInfo := '';

    CodeSite.SendMsg( '拆解FeedBackData開始' );

    nL_CmdID := StrToIntDef(Copy(sG_ReturnPathData,33,4),0);
    if (nL_CmdID <> ACKNOWLEDGE_CMD_ID) and (nL_CmdID <> NON_ACKNOWLEDGE_CMD_ID) then
    begin

      if dtmMain.insertCallbackData(sG_ReturnPathData) then
        sInfo := sG_ReturnPathData +  ' => 處理完成' + '[ ' + datetimetostr(now) + ' ]'
      else
        sInfo := sG_ReturnPathData +  ' => 處理失敗' + '[ ' + datetimetostr(now) + ' ]';

    end
    else if (nL_CmdID = ACKNOWLEDGE_CMD_ID) then
    begin
      sInfo := '接收到 ACK 指令' + '[ ' + datetimetostr(now) + ' ]';
    end
    else if (nL_CmdID = NON_ACKNOWLEDGE_CMD_ID) then
    begin
      sInfo := '接收到 NACK 指令 : ' + sG_ReturnPathData + ' [ ' + datetimetostr(now) + ' ]';
    end
    else
      sInfo := '';

    if sInfo <>'' then
    begin
      CodeSite.SendMsg( '拆解回來的資料=' + sInfo );
      Synchronize( UpdateUI );
    end;

    CodeSite.SendMsg( '拆解FeedBackData結束' );

    {
    Feed Back Command 之格式:

    1~10 碼: UA => command 208, 209, 210 之 UA 是 0000000000
    11~14 碼 : feed back command id
    其餘各碼的拆解方式依各指令而定(每個指令都不一樣)
    }

    {
immediate callback 的 callback 訊息
600011753040257000344289200302181160052756021120030218103540
60001175404025700034428920030218116005275602010000000006553600150000006700
600011755040257000344289200302181160052756020600000000065536N
600011756040257000344289200302181160052756020500000000065536123456          999999999999999999999999999999990000000055528509
600011757040257000344289200302181160052756021200

}

    //frmMain.Memo1.Lines.Add(sG_ReturnPathData);

    sL_CmdHeader := trim(Copy(sG_ReturnPathData,1,32));
    sL_CmdBody := trim(Copy(sG_ReturnPathData,33,length(sG_ReturnPathData)-32));

    sL_TransactionNo := trim(Copy(sL_CmdHeader,1,9));
    sL_CmdType := trim(Copy(sL_CmdHeader,10,2));
    sL_SrcID := trim(Copy(sL_CmdHeader,12,4));
    sL_DestID := trim(Copy(sL_CmdHeader,16,4));
    sL_MoppID := trim(Copy(sL_CmdHeader,20,5));
    sL_CreationDate := trim(Copy(sL_CmdHeader,25,8));




    sL_FbIccNo := trim(Copy(sL_CmdBody,1,10)); //smart card num

    sL_FbCmdID := trim(Copy(sL_CmdBody,11,4)); // feedback command id
    nL_FbCmdID := StrToIntDef(sL_FbCmdID, 0);

    case nL_FbCmdID of
      LOW_CREDIT_ALARM_CMD_ID : //low credit alarm : cmd 200
       begin
         sL_StbNO := trim(Copy(sL_CmdBody,15,14));
         sL_Credit := trim(Copy(sL_CmdBody,29,7));
         sL_Debit := trim(Copy(sL_CmdBody,36,7));
       end;
      CURRENT_DEBIT_AND_CREDIT_CMD_ID :// current debit and credit : cmd 201
       begin
//60001175404025700034428920030218116005275602010000000006553600150000006700       
         sL_StbNO := trim(Copy(sL_CmdBody,15,14));
         sL_Credit := trim(Copy(sL_CmdBody,29,7));
         sL_Debit := trim(Copy(sL_CmdBody,36,7));
       end;
      PPV_PURCHASE_LIST_CMD_ID :// ppv purchase list : cmd 202 => cmd 111 會傳回一串 cmd 202
       begin
//60001677604025700034428920030219116005275602020000000006553600000000006320030219Y

         sL_StbNO := trim(Copy(sL_CmdBody,15,14));
         sL_ImsProductID := trim(Copy(sL_CmdBody,29,12));
         sL_PurchaseDate := trim(Copy(sL_CmdBody,41,8));
         sL_WatchedStatus := trim(Copy(sL_CmdBody,49,1)); // 'Y' or 'N' => 表示此 IPPV 是否已經被看過了
       end;
      PHONE_DISCREPANCIES_CMD_ID :// phone discrepancies : cmd 205
       begin
//600011756040257000344289200302181160052756020500000000065536123456          999999999999999999999999999999990000000055528509
         sL_StbNO := trim(Copy(sL_CmdBody,15,14));
         sL_PhoneNum1 := trim(Copy(sL_CmdBody,29,16));
         sL_PhoneNum2 := trim(Copy(sL_CmdBody,45,16));
         sL_PhoneNum3 := trim(Copy(sL_CmdBody,61,16));
         sL_AbnormalPhone := trim(Copy(sL_CmdBody,77,16));
       end;
      STB_RESPONDING_STATUS_CMD_ID :// stb responding status : cmd 206
       begin
//600011755040257000344289200302181160052756020600000000065536N
         sL_StbNO := trim(Copy(sL_CmdBody,15,14));
         sL_Responding := trim(Copy(sL_CmdBody,29,1));
       end;
      ICC_MEMORY_FULL_ALARM_CMD_ID :// icc memory full alarm : cmd 207
       begin
         sL_StbNO := trim(Copy(sL_CmdBody,15,14));
       end;
      EVENT_DEFINITION_ERROR_CMD_ID :// event definition error : cmd 208
       begin
         sL_SmsEventID1 := trim(Copy(sL_CmdBody,15,12));
         sL_SmsEventID2 := trim(Copy(sL_CmdBody,27,12));
       end;
      NULL_EVENT_ERROR_CMD_ID :// null event error : cmd 209
       begin
         sL_SmsEventID := trim(Copy(sL_CmdBody,15,12));
       end;
      EPG_DATA_FEED_FORMAT_ERROR_CMD_ID :// epg data feed format error : cmd 210
       begin
         sL_ErrorCode := trim(Copy(sL_CmdBody,15,4));
         sL_ErrorCodeExt := trim(Copy(sL_CmdBody,19,4));
       end;
      START_OF_REPORT_CMD_ID :// start of report : cmd 211
       begin
//600011753040257000344289200302181160052756021120030218103540       
         sL_StuCallbackDate := trim(Copy(sL_CmdBody,15,8));
         sL_StuCallbackTime := trim(Copy(sL_CmdBody,23,6));
       end;
      END_OF_REPORT_CMD_ID :// end of report : cmd 212
       begin
//600011757040257000344289200302181160052756021200       
         sL_NumOfIppv := trim(Copy(sL_CmdBody,15,2));
       end;
      EVENT_PRODUCT_SCHEDULE_CHANGE_CMD_ID :// event product schedule change : cmd 213
       begin
         sL_SmsEventID := trim(Copy(sL_CmdBody,15,12));
         sL_PreviousStartDate := trim(Copy(sL_CmdBody,27,8));
         sL_PreviousStartTime := trim(Copy(sL_CmdBody,35,6));
         sL_PreviousStartDuration := trim(Copy(sL_CmdBody,41,4));

         sL_NewStartDate := trim(Copy(sL_CmdBody,45,8));
         sL_NewStartTime := trim(Copy(sL_CmdBody,53,6));
         sL_NewStartDuration := trim(Copy(sL_CmdBody,59,4));

       end;
      EVENT_REMOVE_ERROR_CMD_ID :// event remove error : cmd 214
       begin
         sL_SmsEventID := trim(Copy(sL_CmdBody,15,12));
       end;
      PRODUCTS_LIST_CMD_ID :// products list : cmd 215 => cmd 71 會傳回 cmd 215
       begin
//000003033040002000344289200302191160052756021500900012300000000000000N02000000000004N000000000005N

         sL_OriginalTransactionNum := trim(Copy(sL_CmdBody,15,9));
         sL_StuNO := trim(Copy(sL_CmdBody,24,14));
         sL_IccSuspended := trim(Copy(sL_CmdBody,38,1)); // => 'Y' or 'N'
         sL_NbOfProducts := trim(Copy(sL_CmdBody,39,2));
         nL_NbOfProducts := StrToIntDef(sL_NbOfProducts,0);
         sL_ImsProductIdList := '';
         sL_ProductSuspendList := '';

         for ii:=1 to nL_NbOfProducts do
         begin
           if ii=1 then
           begin
             sL_ImsProductIdList := trim(Copy(sL_CmdBody,41,12));
             sL_ProductSuspendList  := trim(Copy(sL_CmdBody,53,1));
             nL_TmpBeginNum1 := 54;
             nL_TmpBeginNum2 := 66;
           end
           else
           begin
             sL_ImsProductIdList := sL_ImsProductIdList + ',' + trim(Copy(sL_CmdBody,nL_TmpBeginNum1,12));
             sL_ProductSuspendList  := sL_ProductSuspendList + ',' + trim(Copy(sL_CmdBody,nL_TmpBeginNum2,1));
             nL_TmpBeginNum1 := nL_TmpBeginNum1 + 13;
             nL_TmpBeginNum2 := nL_TmpBeginNum2 + 13;
             
           end;
         end;
       end;

    end;
    
end;

{ ---------------------------------------------------------------------------- }

procedure ProcessReturnPathData.PreParseReturnPathData;
begin
  while ( frmMain.G_ReturnPathDataStrList.Count > 0 ) do
  begin
    CodeSite.SendWarning( 'Enter PreParseReturnPathData' );
    sG_ReturnPathData := frmMain.G_ReturnPathDataStrList.Strings[0];
    PreSendResponse( sG_ReturnPathData );
    frmMain.G_ReturnPathDataStrList.Delete(0);
    ParseReturnPathData;
    CodeSite.SendWarning( 'Leave PreParseReturnPathData' );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure ProcessReturnPathData.PreSendResponse(const AResponseData: String);
var
    sL_TransactionNo : String;
    nL_CmdID : Integer;
begin
  nL_CmdID := StrToIntDef( Copy( AResponseData, 33, 4 ), 0 );
  //表示此筆不是 Nagra 丟給我的 ACK or NACK command..
  if ( nL_CmdID <> ACKNOWLEDGE_CMD_ID ) and
     ( nL_CmdID <> NON_ACKNOWLEDGE_CMD_ID ) then
  begin
    sL_TransactionNo := Copy( AResponseData, 1, 9 );
    if IsValidCmd( AResponseData ) then
      SendCmd( '1000', sL_TransactionNo )
    else
      SendCmd( '1001', sL_TransactionNo );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure ProcessReturnPathData.SendCmd(const sI_CmdID, sI_SrcTransNum: String);
var
    L_TmpStr : String;
    L_TmpChar : array[0..1024] of char;
    nL_Len, ii, nL_Words : Integer;
    sL_FullCommand: String;
    sL_CommandRootHeader, sL_CommandBody : String;
    sL_BoxNo, sL_IccNo : String;
    sL_CommandNo, sL_CommandType, sL_SmsCommandRootHeader, sL_CurrentDate : String;
    L_TmpRecord : TTempRecord;
    L_DeviceData : TDeviceData;
    AObj: TCmdInfoObject;
    sL_RealTransID : String;
begin

    try
        frmMain.sG_SrcTransNum := sI_SrcTransNum;

        sL_CommandNo := sI_CmdID;

        sL_CommandType := SendCommandU.getCommandType( StrToInt( sL_CommandNo ) );

        sL_SmsCommandRootHeader := SendCommandU.getSmsCommandRootHeader( False,
          0, StrToInt( sL_CommandNo ), sL_CommandType, '' );

        sL_CurrentDate := TUdateTime.GetPureDateStr( Date, '' );
        
        sL_BoxNo := '';
        sL_IccNo := '';

        sL_CommandRootHeader :=
          SendCommandU.getCommandSection( sL_CommandType,
          sL_CurrentDate, sL_CurrentDate, sL_IccNo );

        { 這個 TCmdInfoObject 是為了傳參數用的, 沒有用途 }
        AObj := TCmdInfoObject.Create;
        try
          sL_CommandBody :=
            SendCommandU.getCommandBody( AObj, StrToInt( sL_CommandNo ),
            nil, '', '', '0' );
        finally
          AObj.Free;
        end;    

        sL_FullCommand := sL_SmsCommandRootHeader + sL_CommandRootHeader +
          sL_CommandBody;
        
        if ( frmMain.nG_SC_RealCmdSeqNo2 >= frmMain.nG_SC_MaxRealCmdSeqNo2 ) then
          frmMain.nG_SC_RealCmdSeqNo2 := 1
        else
          Inc( frmMain.nG_SC_RealCmdSeqNo2 );

        sL_RealTransID :=
          TUstr.AddString( FEEDBACK_COMP_CODE,'0', False, 3 ) +
          TUstr.AddString( IntToStr( frmMain.nG_SC_RealCmdSeqNo2 ), '0', False,
          SEQ_NUM_LENGTH );
          
        sL_FullCommand := sL_RealTransID + sL_FullCommand;

        //CodeSite.SendWarning( 'FeedBack ACK = ' + sI_CmdID );
        //CodeSite.SendWarning( 'ACKCmd = ' + sL_FullCommand );

        nL_Len := Length( sL_FullCommand );

        L_TmpRecord := frmMain.GetBinaryVal( nL_Len,2, False );
        
        L_DeviceData.len[0] := L_TmpRecord.TempAry[0];
        L_DeviceData.len[1] := L_TmpRecord.TempAry[1];

        FillChar( L_DeviceData.Data, Length( L_DeviceData.data ) + 1, 0 );
        
        for ii:=1 to length(sL_FullCommand) do
        begin
          L_DeviceData.data[ii-1] := sL_FullCommand[ii];
        end;

        L_TmpChar[0] := L_DeviceData.len[0];
        L_TmpChar[1] := L_DeviceData.len[1];
        
        for ii:=1 to length(sL_FullCommand) do
        begin
          L_TmpChar[ii+1] := sL_FullCommand[ii];
        end;

        for ii:=0 to High(L_TmpChar) do
          L_TmpStr := L_TmpStr + L_TmpChar[ii];

        //send_data..
        //nL_Words := write(@L_DeviceData, nL_Len+2, frmMain.G_ReturnPathWStream);
        
        frmMain.FeedbackSocket.SendBuf( L_DeviceData, nL_Len + 2 );


    except
      on E: Exception do
      begin
        CodeSite.SendError( 'ProcessReturnPathData inner loop Error: %s ',
          [E.Message] );
      end;
    end;
end;

{ ---------------------------------------------------------------------------- }

procedure ProcessReturnPathData.UpdateUI;
begin
  if frmMain.memErrorLog.Lines.Count >= CA_UI_MAX_COUNT*5 then
    frmMain.memErrorLog.Lines.Clear;
  frmMain.memErrorLog.Lines.Add( sInfo );
end;

{ ---------------------------------------------------------------------------- }

end.
