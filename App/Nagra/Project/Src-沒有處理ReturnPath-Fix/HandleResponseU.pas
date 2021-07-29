unit HandleResponseU;

interface

uses
  Classes, Forms, Comctrls, SysUtils, ConstObjU, Ustru, UdateTimeu,
  HandleUIU,
  CsIntf;


  procedure OnTimerStartAction;
  procedure UpdateResponseMsg;
  procedure WriteToMis; overload;
  procedure WriteToMis(const AResponse: String); overload;
  procedure UpdateCmdStatus(const AErrorCode, AErrorMsg, ACmdStatus,
    ATransNumForUpdateCmdStatus : String);


implementation

uses frmMainU, dtmMainU;

{ ---------------------------------------------------------------------------- }

procedure WriteToMis(const AResponse: String);
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

    if Length( AResponse ) < 20 then Exit;
    nL_InvalidSeedValue := 999;
    nL_SeedPos := nL_InvalidSeedValue;
    sL_ResCmdTypeFlag := '05';

    nL_ResTransNumLength := 9; //傳回之 transaction num 有 9 碼
    nL_SeedPos := 4;

    if nL_SeedPos = nL_InvalidSeedValue then Exit;

    try
    begin
      sL_ResTransNum2 := Copy( AResponse, 33 + nL_SeedPos, 9 );
      bL_WriteToMis := True;
      sL_ResTransNum1 := Copy( AResponse, 1, nL_ResTransNumLength );
      sL_ResCmdType := Copy( AResponse, 6 + nL_SeedPos, 2 );

      sL_ResSourceID := Copy( AResponse, 8 + nL_SeedPos, 4 ); //此 source id 等於傳入之 dest id
      sL_ResDestID := Copy( AResponse, 12 + nL_SeedPos, 4 ); //此 dest id 等於傳入之 source id
      sL_ResMopPPID := Copy( AResponse, 16 + nL_SeedPos, 5 );
      sL_ResCreatDate := Copy( AResponse, 21 + nL_SeedPos, 8 );
      nL_ResCmdID := StrToIntDef( Copy( AResponse, 29 + nL_SeedPos, 4 ), 0 );
      nL_TmpResponseErrorCode := StrToIntDef( Copy( AResponse, 43 + nL_SeedPos, 4 ), -1 );
      //若不是連線測試指令
      if StrToIntDef( Copy( sL_ResTransNum2, 1, 3 ),0) <>
         StrToInt( TESTING_CMD_COMP_CODE ) then
      begin
        //回傳command 1000 或回傳之錯誤代碼為0者,均視為OK
        if ( nL_ResCmdID = ACKNOWLEDGE_CMD_ID) or 
           ( nL_TmpResponseErrorCode = 0 ) then
        begin
          if frmMain.bG_ResponseLog then
          begin
            dtmMain.insertCmdResponseLog( AResponse );
          end;

          bL_Error := False;
          sL_ResImsProductID := Copy( AResponse, 42 + nL_SeedPos, 12 );
          sL_ResSmsProductID := Copy( AResponse, 54 + nL_SeedPos, 12 );
          sL_ErrorCode := NO_ERROR;
          sL_ErrorMsg := NO_ERROR_MSG;
          nL_CompCode := StrToIntDef( Copy( sL_ResTransNum2, 1, 3 ), 0 );
          nL_ConnectionID := frmMain.getConnectionID( nL_CompCode );
          try
            nL_UiStatus := 2;
            frmMain.nG_HR_CompCode := nL_CompCode;
            frmMain.nG_HR_ConnectionID := nL_ConnectionID;
            UpdateCmdStatus( sL_ErrorCode, sL_ErrorMsg, 'C', sL_ResTransNum2 );
          except
            { ... }
          end;
        end
        else if (nL_ResCmdID=NON_ACKNOWLEDGE_CMD_ID) then
        begin
          if frmMain.bG_ResponseLog then
          begin
            dtmMain.insertCmdResponseLog( AResponse );
          end;
          bL_Error := true;
          sL_ResNackStatus := Copy( AResponse, 42 + nL_SeedPos, 1 );
          sL_ResErrorCode := Copy( AResponse, 43 + nL_SeedPos, 4 );
          sL_ResErrorCodeExt := Copy( AResponse, 47 + nL_SeedPos,4 );
          frmMain.getErrorInfo( sL_ResErrorCode, sL_ResErrorCodeExt,
            sL_ResErrorCodeDetail, sL_ResErrorCodeExtDetail );
          sL_ResLengthOfCmdBody := Copy( AResponse, 51 + nL_SeedPos, 3 );
          sL_ResCmdSestion := Copy( AResponse, 54 + nL_SeedPos,
            StrToInt( sL_ResLengthOfCmdBody ) );
          sL_ErrorCode := sL_ResErrorCode + '--' + sL_ResErrorCodeDetail;
          sL_ErrorMsg := sL_ResErrorCodeExt + '--' + sL_ResErrorCodeExtDetail;
          nL_CompCode := StrToIntDef( Copy( sL_ResTransNum2, 1, 3 ), 0 );
          nL_ConnectionID := frmMain.getConnectionID(nL_CompCode);
          if nL_ConnectionID>=0 then
          begin
            try
              nL_UiStatus := 3;
              frmMain.nG_HR_CompCode := nL_CompCode;
              frmMain.nG_HR_ConnectionID := nL_ConnectionID;
              UpdateCmdStatus(sL_ErrorCode, sL_ErrorMsg, 'E', sL_ResTransNum2);
            except
              { ... }
            end;
          end;
        end;
      end
      else
      begin
        //回傳command 1000 或回傳之錯誤代碼為0者,均視為ok
        if ( nL_ResCmdID = ACKNOWLEDGE_CMD_ID ) or
           ( nL_TmpResponseErrorCode = 0 ) then
          nL_UiStatus := 2
        else
          nL_UiStatus := 3;
      end;
      frmMain.setTransCaption;
      if (frmMain.bG_ShowUI) then
      begin
        HandleUIU.updateListItemStatus( nil, nL_UiStatus,
          sL_ResTransNum2, sL_ErrorCode, sL_ErrorMsg );
      end;
    end
    except
      on E: Exception do
      begin
        frmMain.showLogOnMemo( WRITE_TO_MIS_ERROR_FLAG,
          ' HandleResponse.writeToMis' + '=>' + E.Message );
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
        nL_ConnectionID := frmMain.getConnectionID(nL_CompCode);
        if nL_ConnectionID>=0 then
        begin
          try
            frmMain.nG_HR_CompCode := nL_CompCode;
            frmMain.nG_HR_ConnectionID := nL_ConnectionID;
            updateCmdStatus(sL_ErrorCode, sL_ErrorMsg, 'E', sL_ResTransNum2);
          except
            { ... }
          end;
        end;
        Exit;
      end;
    end;
    
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
    end;
end;

{ ---------------------------------------------------------------------------- }

procedure WriteToMis;
begin
  WriteToMis( frmMain.sG_HR_ResponseStr );
end;

{ ---------------------------------------------------------------------------- }

procedure UpdateResponseMsg;
begin
  frmMain.G_ResponseMsgStrList.Strings[frmMain.nG_HR_MsgStrListPos] := '';
end;

{ ---------------------------------------------------------------------------- }

procedure UpdateCmdStatus(const AErrorCode, AErrorMsg,
  ACmdStatus, ATransNumForUpdateCmdStatus : String);
begin
  try
    dtmMain.UpdateCMD_Status_For_HandleResponse(
      frmMain.nG_HR_ConnectionID, ATransNumForUpdateCmdStatus,
      ACmdStatus, AErrorCode, AErrorMsg );
  except
    frmMain.ShowLogOnMemo( UPDATE_CMD_STATUS_ERROR_FLAG,
      ' HandleResponse.updateResponseMsg' );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure OnTimerStartAction;
var
  AResponse: String;
begin
  try
    if frmMain.bG_HandleResponseThread then
    begin
      if Assigned( frmMain.G_ResponseMsgStrList ) and
        ( frmMain.G_ResponseMsgStrList.Count > 0 ) then
      begin
        AResponse := frmMain.G_ResponseMsgStrList.Strings[0];
        frmMain.G_ResponseMsgStrList.Delete( 0 );
        WriteToMis( AResponse );
      end;
    end;
  except
    on E: Exception do
    begin
      CodeSite.SendError( 'Handle Response', [E.Message] );
      frmMain.ShowLogOnMemo( HANDLE_RESPONSE_ERROR_FLAG,
        ' THandleResponse.Execute' );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
