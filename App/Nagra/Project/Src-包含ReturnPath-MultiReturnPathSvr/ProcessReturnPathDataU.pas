unit ProcessReturnPathDataU;

interface

uses
  Classes, SysUtils, ConstObjU, Forms, ScktComp;

type
  ProcessReturnPathData = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  public
    nG_60003TransID, nG_Max60003TransID : Integer;  
    sG_ReturnPathData : String;
    function Increment(Ptr: Pointer; Increment: Integer): Pointer;    
    function Write(I_Buf: Pointer; I_Count: Integer; I_WStream: TWinSocketStream):Integer;    

    procedure sendCmd(sI_CmdID, sI_SrcTransNum:String);    
    function isValidCmd(sI_Cmd:String):boolean;    
    procedure preSendResponse(sI_ResponseData:String);
    procedure preParseReturnPathData;
    procedure parseReturnPathData;    
    constructor Create;overload;
  end;

implementation

uses frmMainU, dtmMainU, UdateTimeu, Ustru;

{ Important: Methods and properties of objects in VCL or CLX can only be used
  in a method called using Synchronize, for example,

      Synchronize(UpdateCaption);

  and UpdateCaption could look like,

    procedure ProcessReturnPathData.UpdateCaption;
    begin
      Form1.Caption := 'Updated in a thread';
    end; }

{ ProcessReturnPathData }

constructor ProcessReturnPathData.Create;
begin
    inherited Create(false);
end;

procedure ProcessReturnPathData.Execute;
var
    L_ErrorInfoStrList : TStringList;
    sL_ErrorInfoFileName : String;
begin
  { Place thread code here }
  Exit;
  {
  try
    self.FreeOnTerminate := true;
    self.OnTerminate := frmMain.ProcessReturnPathDataThreadDone;

    nG_60003TransID := 0;
    nG_Max60003TransID := StrToIntDef(TUstr.AddString('','9',true,SEQ_NUM_LENGTH),999999);
    while true do
    begin

      if frmMain.bG_ProcessReturnPathDataThread then
      begin
        preParseReturnPathData;
      end;
    end;

  except
      on E: Exception do
      begin

        L_ErrorInfoStrList := TStringList.Create;
        sL_ErrorInfoFileName := 'c:\Error21.txt';

        if FileExists(sL_ErrorInfoFileName) then
          L_ErrorInfoStrList.LoadFromFile(sL_ErrorInfoFileName);
        if L_ErrorInfoStrList.Count>500 then
          L_ErrorInfoStrList.Clear;
        L_ErrorInfoStrList.Insert(0,'   Msg:' +E.Message +', HelpContext:' + IntToStr(E.HelpContext));
        L_ErrorInfoStrList.Insert(0,'ProcessReturnPathData.Execute error:' + DateTimeToStr(now));
        L_ErrorInfoStrList.SaveToFile(sL_ErrorInfoFileName);
        L_ErrorInfoStrList.Free;
      end;
  end;
  }
end;

function ProcessReturnPathData.Increment(Ptr: Pointer;
  Increment: Integer): Pointer;
begin
    Result := Pointer(Longint(Ptr) + Increment);
end;

function ProcessReturnPathData.isValidCmd(sI_Cmd: String): boolean;
begin
    //判斷是否是正確的指令格式
    result := true;

end;

procedure ProcessReturnPathData.parseReturnPathData;
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

    sL_Info : String;
    nL_CmdID : Integer;
begin
    nL_CmdID := StrToIntDef(Copy(sG_ReturnPathData,33,4),0);
    if (nL_CmdID <> ACKNOWLEDGE_CMD_ID) and (nL_CmdID <> NON_ACKNOWLEDGE_CMD_ID) then
    begin

      if dtmMain.insertCallbackData(sG_ReturnPathData) then
        sL_Info := sG_ReturnPathData +  ' => 處理完成' + '[ ' + datetimetostr(now) + ' ]'
      else
        sL_Info := sG_ReturnPathData +  ' => 處理失敗' + '[ ' + datetimetostr(now) + ' ]';

    end
    else if (nL_CmdID = ACKNOWLEDGE_CMD_ID) then
    begin
      sL_Info := '接收到 ACK 指令' + '[ ' + datetimetostr(now) + ' ]';
    end
    else if (nL_CmdID = NON_ACKNOWLEDGE_CMD_ID) then
    begin
      sL_Info := '接收到 NACK 指令 : ' + sG_ReturnPathData + ' [ ' + datetimetostr(now) + ' ]';
    end
    else
      sL_Info := '';

    if sL_Info<>'' then
    begin
      if frmMain.Memo1.Lines.Count >= CA_UI_MAX_COUNT*5 then
        frmMain.Memo1.Lines.Clear;
      frmMain.Memo1.Lines.Add(sL_Info);
    end;
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
    {
    frmMain.Memo1.Lines.Add(sG_ReturnPathData);

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
    }
end;

procedure ProcessReturnPathData.preParseReturnPathData;
var
    ii : Integer;

begin
    if (frmMain.G_ReturnPathDataStrList.Count>0)then
    begin
//      Application.ProcessMessages;
      sG_ReturnPathData := frmMain.G_ReturnPathDataStrList.Strings[0];
//      frmMain.G_ReturnPathDataStrList1.Add(sG_ReturnPathData);
//      preSendResponse(sG_ReturnPathData);
      frmMain.G_ReturnPathDataStrList.Delete(0);
//      Application.ProcessMessages;
//      parseReturnPathData;
//      Application.ProcessMessages;
    end;

{
//    if (frmMain.G_ReturnPathDataStrList.Count>0) and (frmMain.G_ReturnPathDataStrList.Strings[0]='N') then
    if (frmMain.G_ReturnPathDataStrList.Count>0)then
    begin
      Application.ProcessMessages;
      G_ReturnPathDataObj := (frmMain.G_ReturnPathDataStrList.Objects[0] as TReturnPathDataObj);
      sG_ReturnPathData := G_ReturnPathDataObj.Data ;
      frmMain.G_ReturnPathDataStrList.Delete(0);
      Application.ProcessMessages;
      parseReturnPathData;
      Application.ProcessMessages;
    end;
}
end;

procedure ProcessReturnPathData.preSendResponse(sI_ResponseData: String);
var
    sL_TransactionNo : String;
    nL_CmdID : Integer;
begin
      nL_CmdID := StrToIntDef(Copy(sI_ResponseData,33,4),0);
      if (nL_CmdID = ACKNOWLEDGE_CMD_ID) or (nL_CmdID = NON_ACKNOWLEDGE_CMD_ID) then
      begin
        sL_TransactionNo := Copy(sI_ResponseData,1,9);
        if isValidCmd(sI_ResponseData) then
          sendCmd('1000', sL_TransactionNo)
        else
          sendCmd('1001', sL_TransactionNo);
      end;

end;

procedure ProcessReturnPathData.sendCmd(sI_CmdID, sI_SrcTransNum: String);
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

    sL_RealTransID : String;
begin
    try
//        Application.ProcessMessages;
        frmMain.G_SendCommandThread.sG_SrcTransNum := sI_SrcTransNum;
        sL_CommandNo := sI_CmdID;//直接送出指令 1002, 用以確認 connection 是否順利建立 => 欲接收 feedback command 的 process ,才需要送出 command 1002
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
        frmMain.Memo4.Lines.Add('F1: ' +sL_FullCommand );

                
//        Application.ProcessMessages;

        if nG_60003TransID>=nG_Max60003TransID then
          nG_60003TransID := 1
        else
          INC(nG_60003TransID);

        sL_RealTransID := TUstr.AddString(PORT_60003_COMP_CODE,'0',false,3) +
                          TUstr.AddString(IntToStr(nG_60003TransID),'0',false,SEQ_NUM_LENGTH);
        sL_FullCommand := sL_RealTransID + sL_FullCommand;

        frmMain.Memo4.Lines.Add('60003傳回給Nagra: ' +sL_FullCommand );
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

        L_TmpChar[0] := L_DeviceData.len[0];
        L_TmpChar[1] := L_DeviceData.len[1];
        for ii:=1 to length(sL_FullCommand) do
        begin
          L_TmpChar[ii+1] := sL_FullCommand[ii];
        end;

        for ii:=0 to High(L_TmpChar) do
          L_TmpStr := L_TmpStr + L_TmpChar[ii];
//        L_TmpStr := StrPas(L_TmpChar);
        nL_Words := write(@L_DeviceData, nL_Len+2, frmMain.G_ReturnPathWStream); //send_data..
    except

    end;
end;

function ProcessReturnPathData.Write(I_Buf: Pointer; I_Count: Integer;
  I_WStream: TWinSocketStream): Integer;
var
    nL_Idx: Integer;
begin
  nL_Idx := 0;
  while (nL_Idx < I_Count) do
  try
    nL_Idx := nL_Idx + I_WStream.Write(Increment(I_Buf, nL_Idx)^, I_Count-nL_Idx);
  except

    frmMain.G_MasterReturnPathSvr.buildConnection(frmMain.ReturnPathClntSocket, frmMain.G_ReturnPathWStream);
    Write(I_Buf, I_Count,frmMain.G_ReturnPathWStream);
//    raise;
  end;

  result := nL_Idx;
end;

end.
