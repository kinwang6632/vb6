// 若用 Synchronize(updateCmdStatus), 則指令的處理速度會比較快(速度差好幾倍), 但 sometimes...程式會被 hold 住...
// SendCommandU 與 HandleResponseU 均有此 function...
unit HandleResponseU;

interface

uses
  Classes, forms, comctrls, syncobjs, dbtables, Sysutils, Dialogs,
  ConstObjU, Ustru, UdateTimeu, scktcomp, winsock, inifiles, ADODB;


type
  THandleResponse = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  public

    procedure onTimerStartAction;
    procedure updateResponseMsg;
    procedure updateCmdStatus(sI_ErrorCode, sI_ErrorMsg, sI_CmdStatus, sI_TransNumForUpdateCmdStatus : String);
    procedure rebuildResponseMsgStrList;
    procedure writeToMis(sI_ResponseStr:String);overload;
    procedure writeToMis;overload;
      
    constructor Create;overload;
  end;

implementation

uses frmMainU, dtmMainU;

{ Important: Methods and properties of objects in VCL or CLX can only be used
  in a method called using Synchronize, for example,

      Synchronize(UpdateCaption);

  and UpdateCaption could look like,

    procedure THandleResponse.UpdateCaption;
    begin
      Form1.Caption := 'Updated in a thread';
    end; }

{ THandleResponse }


constructor THandleResponse.Create;
begin
//    CRS := TCriticalSection.Create;
    inherited Create(false);

end;

{
procedure THandleResponse.Execute;
var
    sL_Response : String;
    ii, nL_Count: Integer;
    L_ErrorInfoStrList : TStringList;
    sL_ErrorInfoFileName : String;
begin
 
    self.FreeOnTerminate := true;
    self.OnTerminate := frmMain.HandleResponseThreadDone;
    while true do
    begin
      try
        if frmMain.bG_HandleResponseThread then
        begin

          nL_Count := frmMain.G_ResponseMsgStrList.Count;
  //        for ii:= 0 to frmMain.G_ResponseMsgStrList.Count -1 do
          for ii:= 0 to nL_Count -1 do
          begin
            sL_Response := frmMain.G_ResponseMsgStrList.Strings[ii];
            sG_ResponseStr := sL_Response;
            if sL_Response<>'' then
            begin
  //            writeToMis(sL_Response);
  //            Synchronize(writeToMis); howard
              writeToMis;
              nG_MsgStrListPos := ii;
//              Synchronize(updateResponseMsg);
              updateResponseMsg;


            end;
          end;
//          CRS.Enter;
          Synchronize(rebuildResponseMsgStrList);
//          CRS.Leave;
        end;
      except
        on E: Exception do
        begin



          sL_ErrorInfoFileName := 'c:\Error2.txt';
          L_ErrorInfoStrList := TStringList.Create;
          if FileExists(sL_ErrorInfoFileName) then
            L_ErrorInfoStrList.LoadFromFile(sL_ErrorInfoFileName);
          if L_ErrorInfoStrList.Count>500 then
            L_ErrorInfoStrList.Clear;
          L_ErrorInfoStrList.Insert(0,'   Msg:' +E.Message +', HelpContext:' + IntToStr(E.HelpContext));
          L_ErrorInfoStrList.Insert(0,'THandleResponse.Execute error:' + DateTimeToStr(now));

          L_ErrorInfoStrList.SaveToFile(sL_ErrorInfoFileName);
          L_ErrorInfoStrList.Free;
        end;
      end;
    end;
end;


}

procedure THandleResponse.Execute;
var
    sL_Response : String;
    ii, nL_Count: Integer;
    L_ErrorInfoStrList : TStringList;
    sL_ErrorInfoFileName : String;
begin

    self.FreeOnTerminate := true;
    self.OnTerminate := frmMain.HandleResponseThreadDone;

{
    onTimerStartAction;
    while true do
    begin
      try
        if frmMain.bG_HandleResponseThread then
        begin

          if Assigned(frmMain.G_ResponseMsgStrList) and (frmMain.G_ResponseMsgStrList.Count >0) then
          begin

//0304            if frmMain.G_ResponseMsgStrList.Strings[0]<>'' then
//0304            begin
              sG_ResponseStr := frmMain.G_ResponseMsgStrList.Strings[0];
              frmMain.G_ResponseMsgStrList.Delete(0);
  //            writeToMis(sL_Response);
              Synchronize(writeToMis);   //howard 0315
//              writeToMis;

//0304            end;
          end;
        end;
      except
        on E: Exception do
        begin

          frmMain.showLogOnMemo(HANDLE_RESPONSE_ERROR_FLAG,' THandleResponse.Execute');

          sL_ErrorInfoFileName := 'c:\Error2.txt';
          L_ErrorInfoStrList := TStringList.Create;
          if FileExists(sL_ErrorInfoFileName) then
            L_ErrorInfoStrList.LoadFromFile(sL_ErrorInfoFileName);
          if L_ErrorInfoStrList.Count>500 then
            L_ErrorInfoStrList.Clear;
          L_ErrorInfoStrList.Insert(0,'   Msg:' +E.Message +', HelpContext:' + IntToStr(E.HelpContext));
          L_ErrorInfoStrList.Insert(0,'THandleResponse.Execute error:' + DateTimeToStr(now));

          L_ErrorInfoStrList.SaveToFile(sL_ErrorInfoFileName);
          L_ErrorInfoStrList.Free;
        end;
      end;
    end;
}
end;

procedure THandleResponse.writeToMis(sI_ResponseStr: String);
var
    L_ErrorInfoStrList : TStringList;
    sL_ErrorInfoFileName : String;

    sL_ResTransNum1,sL_ResTransNum2, sL_ResCmdType, sL_ResSourceID : String;
    sL_ResDestID, sL_ResMopPPID, sL_ResCreatDate : String;
    sL_ResImsProductID, sL_ResSmsProductID : String;
    nL_InvalidSeedValue, nL_ResCmdID, nL_ResTransNumLength : Integer;

    sL_ResNackStatus, sL_ResErrorCode, sL_ResErrorCodeExt : String;
    sL_ResLengthOfCmdBody, sL_ResCmdSestion : String;
    sL_SeqNo, sL_ResCmdTypeFlag, sL_SQL : String;
    nL_SeedPos, nL_CompCode, nL_ConnectionID : Integer;
    sL_ResErrorCodeDetail,sL_ResErrorCodeExtDetail: String;

    nL_UiStatus : Integer;
    sL_ErrorCode, sL_ErrorMsg : String;
    bL_Error, bL_WriteToMis : boolean;
    nL_TmpResponseErrorCode : Integer;
begin
//    Exit;//temp...

    if length(sI_ResponseStr)<20 then Exit;
    nL_InvalidSeedValue := 999;
    nL_SeedPos := nL_InvalidSeedValue;
    sL_ResCmdTypeFlag := '05';

    nL_ResTransNumLength := 9; //傳回之 transaction num 有 9 碼
    nL_SeedPos := 4;

    if nL_SeedPos = nL_InvalidSeedValue then Exit;

    try
    begin
      sL_ResTransNum2 := Copy(sI_ResponseStr,33+nL_SeedPos,9);

      bL_WriteToMis := true;
      sL_ResTransNum1 := Copy(sI_ResponseStr,1,nL_ResTransNumLength);
      sL_ResCmdType := Copy(sI_ResponseStr,6+nL_SeedPos,2);

      sL_ResSourceID := Copy(sI_ResponseStr,8+nL_SeedPos,4); //此 source id 等於傳入之 dest id
      sL_ResDestID := Copy(sI_ResponseStr,12+nL_SeedPos,4); //此 dest id 等於傳入之 source id
      sL_ResMopPPID := Copy(sI_ResponseStr,16+nL_SeedPos,5);
      sL_ResCreatDate := Copy(sI_ResponseStr,21+nL_SeedPos,8);

      nL_ResCmdID := StrToIntDef(Copy(sI_ResponseStr,29+nL_SeedPos,4),0);



      nL_TmpResponseErrorCode := StrToIntDef(Copy(sI_ResponseStr,43+nL_SeedPos,4),-1);
      if StrToIntDef(Copy(sL_ResTransNum2,1,3),0)<>(StrToInt(TESTING_CMD_COMP_CODE)) then //若不是連線測試指令
      begin
        if (nL_ResCmdID=ACKNOWLEDGE_CMD_ID) or (nL_TmpResponseErrorCode =0 )then  //回傳command 1000 或回傳之錯誤代碼為0者,均視為ok
        begin
          if frmMain.bG_ResponseLog then
          begin
            dtmMain.insertCmdResponseLog(sI_ResponseStr);
          end;

          bL_Error := false;
          sL_ResImsProductID := Copy(sI_ResponseStr,42+nL_SeedPos,12);
          sL_ResSmsProductID := Copy(sI_ResponseStr,54+nL_SeedPos,12);
          sL_ErrorCode := NO_ERROR;
          sL_ErrorMsg := NO_ERROR_MSG;
          nL_CompCode := StrToIntDef(Copy(sL_ResTransNum2,1,3),0);
//          if sL_SeqNo <>'' then
//          begin
            nL_ConnectionID := frmMain.getConnectionID(nL_CompCode);
            try
              nL_UiStatus := 2;
{
              sG_ErrorCode := sL_ErrorCode;
              sG_ErrorMsg := sL_ErrorMsg;
}              
//              sG_SeqNo := sL_SeqNo;
              frmMain.nG_HR_CompCode := nL_CompCode;
              frmMain.nG_HR_ConnectionID := nL_ConnectionID;

//              sG_SeqNoForUpdateCmdStatus := sL_SeqNo;
//              sG_TransNumForUpdateCmdStatus := sL_ResTransNum2;
              updateCmdStatus(sL_ErrorCode, sL_ErrorMsg, 'C', sL_ResTransNum2);
//              Synchronize(updateCmdStatus);



            except

            end;
//          end;
        end
        else if (nL_ResCmdID=NON_ACKNOWLEDGE_CMD_ID) then
        begin
          if frmMain.bG_ResponseLog then
          begin
            dtmMain.insertCmdResponseLog(sI_ResponseStr);
          end;
          bL_Error := true;
          sL_ResNackStatus := Copy(sI_ResponseStr,42+nL_SeedPos,1);
          sL_ResErrorCode := Copy(sI_ResponseStr,43+nL_SeedPos,4);
          sL_ResErrorCodeExt := Copy(sI_ResponseStr,47+nL_SeedPos,4);
          frmMain.getErrorInfo(sL_ResErrorCode,sL_ResErrorCodeExt, sL_ResErrorCodeDetail,sL_ResErrorCodeExtDetail);
          sL_ResLengthOfCmdBody := Copy(sI_ResponseStr,51+nL_SeedPos,3);
          sL_ResCmdSestion := Copy(sI_ResponseStr,54+nL_SeedPos,StrToInt(sL_ResLengthOfCmdBody));
          sL_ErrorCode := sL_ResErrorCode + '--' + sL_ResErrorCodeDetail;
          sL_ErrorMsg := sL_ResErrorCodeExt + '--' + sL_ResErrorCodeExtDetail;
          nL_CompCode := StrToIntDef(Copy(sL_ResTransNum2,1,3),0);

//          if sL_SeqNo <>'' then
//          begin
            nL_ConnectionID := frmMain.getConnectionID(nL_CompCode);
            if nL_ConnectionID>=0 then
            begin
              try
                nL_UiStatus := 3;
{
                sG_ErrorCode := sL_ErrorCode;
                sG_ErrorMsg := sL_ErrorMsg;
}
//                sG_SeqNo := sL_SeqNo;
                frmMain.nG_HR_CompCode := nL_CompCode;
                frmMain.nG_HR_ConnectionID := nL_ConnectionID;
//                sG_SeqNoForUpdateCmdStatus := sL_SeqNo;
//                sG_TransNumForUpdateCmdStatus := sL_ResTransNum2;
                updateCmdStatus(sL_ErrorCode, sL_ErrorMsg, 'E', sL_ResTransNum2);
//                Synchronize(updateCmdStatus);



              except

              end;
            end;
//          end;
        end;



      end
      else
      begin
        if (nL_ResCmdID=ACKNOWLEDGE_CMD_ID) or (nL_TmpResponseErrorCode =0 )then  //回傳command 1000 或回傳之錯誤代碼為0者,均視為ok
          nL_UiStatus := 2
        else
          nL_UiStatus := 3;
      end;
      frmMain.setTransCaption;
      if (frmMain.bG_ShowUI) then
      begin
        frmMain.G_HandleUIThread.updateListItemStatus(nil,nL_UiStatus, sL_ResTransNum2, sL_ErrorCode, sL_ErrorMsg);
      end;



    end
    except
      on E: Exception do
      begin
        frmMain.showLogOnMemo(WRITE_TO_MIS_ERROR_FLAG,' HandleResponse.writeToMis' + '=>' + E.Message);

        sL_ErrorInfoFileName := 'c:\Error4.txt';
        L_ErrorInfoStrList := TStringList.Create;
        if FileExists(sL_ErrorInfoFileName) then
          L_ErrorInfoStrList.LoadFromFile(sL_ErrorInfoFileName);
        if L_ErrorInfoStrList.Count>500 then
          L_ErrorInfoStrList.Clear;
        L_ErrorInfoStrList.Insert(0,'   Detail Msg: ConnectionID=' + IntToStr(nL_ConnectionID) +', Cmd Sequence No=' + sL_SeqNo + ',ErrorCode=' + sL_ErrorCode + '  ErrorMsg=' +  sL_ErrorMsg);

        L_ErrorInfoStrList.Insert(0,'   Msg:' +E.Message +', HelpContext:' + IntToStr(E.HelpContext));
        L_ErrorInfoStrList.Insert(0,'THandleResponse.writeToMis error:' + DateTimeToStr(now));
        L_ErrorInfoStrList.SaveToFile(sL_ErrorInfoFileName);
        L_ErrorInfoStrList.Free;



        sL_ErrorCode := WRITE_TO_MIS_ERROR;
        sL_ErrorMsg := WRITE_TO_MIS_ERROR_MSG;
        bL_WriteToMis := false;
        nL_CompCode := StrToIntDef(Copy(sL_ResTransNum2,1,3),0);

//        if sL_SeqNo <>'' then
//        begin
          nL_ConnectionID := frmMain.getConnectionID(nL_CompCode);
          if nL_ConnectionID>=0 then
          begin

            try
{
              sG_ErrorCode := sL_ErrorCode;
              sG_ErrorMsg := sL_ErrorMsg;
}              
//              sG_SeqNo := sL_SeqNo;
              frmMain.nG_HR_CompCode := nL_CompCode;
              frmMain.nG_HR_ConnectionID := nL_ConnectionID;
//              sG_SeqNoForUpdateCmdStatus := sL_SeqNo;
//              sG_TransNumForUpdateCmdStatus := sL_ResTransNum2;
              updateCmdStatus(sL_ErrorCode, sL_ErrorMsg, 'E', sL_ResTransNum2);
//              Synchronize(updateCmdStatus);



            except

            end;
          end;
//        end;
        exit;
      end;
    end;
    {
    showmessage(sL_ResTransNum1);
    showmessage(sL_ResCmdType);
    showmessage(sL_ResSourceID);
    showmessage(sL_ResDestID);

    showmessage(sL_ResMopPPID);
    showmessage(sL_ResCreatDate);

    showmessage(sL_ResCmdID);
    showmessage(sL_ResTransNum2);

    showmessage(sL_ResImsProductID);
    showmessage(sL_ResSmsProductID);

    }
    if frmMain.bG_ResponseLog then
    begin
      sL_SQL := 'update CmdResult set ResCmdType=' + STR_SEP + sL_ResCmdType + STR_SEP ;
      sL_SQL := sL_SQL + ',ResDestID=' + STR_SEP + sL_ResDestID + STR_SEP ;
      sL_SQL := sL_SQL + ',ResSourceID=' + STR_SEP + sL_ResSourceID + STR_SEP ;
      sL_SQL := sL_SQL + ',ResMopPPID=' + STR_SEP + sL_ResMopPPID + STR_SEP ;
      sL_SQL := sL_SQL + ',ResCreatDate=' + STR_SEP + sL_ResCreatDate + STR_SEP ;
      sL_SQL := sL_SQL + ',ResCmdID=' + STR_SEP + IntToSTr(nL_ResCmdID) + STR_SEP ;
      sL_SQL := sL_SQL + ',ResTransNum2=' + STR_SEP + sL_ResTransNum2 + STR_SEP ;
      sL_SQL := sL_SQL + ',ResImsProductID=' + STR_SEP + sL_ResImsProductID + STR_SEP ;
      sL_SQL := sL_SQL + ',ResSmsProductID=' + STR_SEP + sL_ResSmsProductID + STR_SEP ;
      if bL_Error then
      begin
        sL_SQL := sL_SQL + ',ResErrorCode=' + STR_SEP + sL_ResErrorCode + STR_SEP ;
        sL_SQL := sL_SQL + ',ResErrorCodeExt=' + STR_SEP + sL_ResErrorCodeExt + STR_SEP ;
        sL_SQL := sL_SQL + ',ResErrorCodeDetail=' + STR_SEP + sL_ResErrorCodeDetail + STR_SEP ;
        sL_SQL := sL_SQL + ',ResErrorCodeExtDetail=' + STR_SEP + sL_ResErrorCodeExtDetail + STR_SEP ;
      end;        
      sL_SQL := sL_SQL + ',ResLengthOfCmdBody=' + STR_SEP + sL_ResLengthOfCmdBody + STR_SEP ;
      sL_SQL := sL_SQL + ',ResCmdSestion=' + STR_SEP + sL_ResCmdSestion + STR_SEP ;
      sL_SQL := sL_SQL + ' where TransactionNum=' + STR_SEP + sL_ResTransNum2 +STR_SEP;


      {
      with frmMain.adoRealCmd do
      begin
        SQL.Clear;
        Close;
        SQL.Add(sL_SQL);
        ExecSQL;
        Close;
      end;
      }
    end;
end;


procedure THandleResponse.rebuildResponseMsgStrList;
var
    ii : Integer;
begin
    if frmMain.G_ResponseMsgStrList.Count > 10 then
    begin
      for ii:= frmMain.G_ResponseMsgStrList.Count-1 downto 0 do
      begin
        if frmMain.G_ResponseMsgStrList.Strings[ii]='' then
          frmMain.G_ResponseMsgStrList.Delete(ii);
      end;
    end;
end;

procedure THandleResponse.writeToMis;
begin
    writeToMis(frmMain.sG_HR_ResponseStr);
end;

procedure THandleResponse.updateResponseMsg;
begin
    frmMain.G_ResponseMsgStrList.Strings[frmMain.nG_HR_MsgStrListPos] := '';

end;

procedure THandleResponse.updateCmdStatus(sI_ErrorCode, sI_ErrorMsg, sI_CmdStatus, sI_TransNumForUpdateCmdStatus : String);
begin
    try
      dtmMain.UpdateCMD_Status_For_HandleResponse(frmMain.nG_HR_ConnectionID,sI_TransNumForUpdateCmdStatus, sI_CmdStatus, sI_ErrorCode, sI_ErrorMsg);
    except
      frmMain.showLogOnMemo(UPDATE_CMD_STATUS_ERROR_FLAG,' HandleResponse.updateResponseMsg');
    end;  
end;

procedure THandleResponse.onTimerStartAction;
var
    ii, nL_Count: Integer;
    L_ErrorInfoStrList : TStringList;
    sL_ErrorInfoFileName, sL_ResponseStr : String;

begin
      try
        if frmMain.bG_HandleResponseThread then
        begin

          if Assigned(frmMain.G_ResponseMsgStrList) and (frmMain.G_ResponseMsgStrList.Count >0) then
          begin

//0304            if frmMain.G_ResponseMsgStrList.Strings[0]<>'' then
//0304            begin
//              sG_ResponseStr := frmMain.G_ResponseMsgStrList.Strings[0];
              sL_ResponseStr := frmMain.G_ResponseMsgStrList.Strings[0];
              frmMain.G_ResponseMsgStrList.Delete(0);
              writeToMis(sL_ResponseStr);
//              Synchronize(writeToMis);   //howard 0315
//              writeToMis(frmMain.G_ResponseMsgStrList.Strings[0]);

//0304            end;
          end;
        end;
      except
        on E: Exception do
        begin

          frmMain.showLogOnMemo(HANDLE_RESPONSE_ERROR_FLAG,' THandleResponse.Execute');

          sL_ErrorInfoFileName := 'c:\Error2.txt';
          L_ErrorInfoStrList := TStringList.Create;
          if FileExists(sL_ErrorInfoFileName) then
            L_ErrorInfoStrList.LoadFromFile(sL_ErrorInfoFileName);
          if L_ErrorInfoStrList.Count>500 then
            L_ErrorInfoStrList.Clear;
          L_ErrorInfoStrList.Insert(0,'   Msg:' +E.Message +', HelpContext:' + IntToStr(E.HelpContext));
          L_ErrorInfoStrList.Insert(0,'THandleResponse.Execute error:' + DateTimeToStr(now));

          L_ErrorInfoStrList.SaveToFile(sL_ErrorInfoFileName);
          L_ErrorInfoStrList.Free;
        end;
      end;
end;

end.
