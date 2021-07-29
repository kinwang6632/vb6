//Syn OK
unit FetchComDataThreadU;

interface

uses
  Classes, SysUtils, Dialogs, ExtCtrls, ConstObjU, SyncObjs;

  procedure onTimerStartAction;
  procedure UpdateCmdStatusToProcessing; overload;
  procedure UpdateCmdStatusToProcessing(const ASeqNos: String); overload;
  procedure addCmdInfoObj(sI_SeqNo:String; I_CmdInfoObj : TCmdInfoObject);overload;
  procedure FetchData;
  procedure CopyRequestData;

implementation

uses frmMainU, dtmMainU, Ustru;


{ FetchComDataThread }

procedure addCmdInfoObj(sI_SeqNo: String;
  I_CmdInfoObj: TCmdInfoObject);
begin
  if Assigned(frmMain.G_CommandInfoStrList) then
     frmMain.G_CommandInfoStrList.AddObject( sI_SeqNo, I_CmdInfoObj) ;
end;

procedure CopyRequestData;
const
  aDelayNoCommand: Integer = 0;
var
    bL_SendTestCmd : boolean;
    nL_RecCount : Integer;
    sL_SeqNo, sL_SeqNos : String;
    L_CmdInfoObj : TCmdInfoObject;

    nL_Index : Integer;
    L_StrList:TStringList;
begin
    try

    if Assigned( frmMain.G_CommandInfoStrList ) and
      ( frmMain.G_CommandInfoStrList.Count > 0 ) then
      Exit
    else
    begin
      if frmMain.nG_CurActiveConn < 0 then
        Exit
      else
      begin
        with dtmMain.G_AdoSendNagra[frmMain.nG_CurActiveConn] do
        begin

          Last;
          sL_SeqNos := '';
          nL_RecCount := 0;

          if dtmMain.G_AdoSendNagra[frmMain.nG_CurActiveConn].RecordCount >0 then
          begin
            bL_SendTestCmd := false;
            while not Bof do
            begin
              if nL_RecCount=frmMain.nG_MaxCommandCount then
                break;
              INC(nL_RecCount);

              sL_SeqNo := FieldByName('SEQNO').AsString;

              nL_Index := frmMain.G_CommandInfoStrList.IndexOf(sL_SeqNo);

              if nL_Index < 0 then
              begin
                L_CmdInfoObj := TCmdInfoObject.Create;

                L_CmdInfoObj.CMD_STATUS := FieldByName('CMD_STATUS').AsString;
                L_CmdInfoObj.SEQNO  := sL_SeqNo;

                if sL_SeqNos = '' then
                  sL_SeqNos := FieldByName('SEQNO').AsString
                else
                  sL_SeqNos := sL_SeqNos + ',' + FieldByName('SEQNO').AsString;

                L_CmdInfoObj.HIGH_LEVEL_CMD_ID := Trim(FieldByName('HIGH_LEVEL_CMD_ID').AsString);
                L_CmdInfoObj.ICC_NO := Trim(FieldByName('ICC_NO').AsString);
                L_CmdInfoObj.STB_NO := Trim(FieldByName('STB_NO').AsString);
                L_CmdInfoObj.SUBSCRIPTION_BEGIN_DATE := Trim(FieldByName('SUBSCRIPTION_BEGIN_DATE').AsString);
                L_CmdInfoObj.SUBSCRIPTION_END_DATE := Trim(FieldByName('SUBSCRIPTION_END_DATE').AsString);
                L_CmdInfoObj.OPERATOR := Trim(FieldByName('OPERATOR').AsString);
                L_CmdInfoObj.NOTES := Trim(FieldByName('NOTES').AsString);

                L_CmdInfoObj.IMS_PRODUCT_ID := Trim(FieldByName('IMS_PRODUCT_ID').AsString);
                L_CmdInfoObj.ZIP_CODE := Trim(FieldByName('ZIP_CODE').AsString);
                L_CmdInfoObj.CREDIT_MODE  := Trim(FieldByName('CREDIT_MODE').AsString);
                L_CmdInfoObj.CREDIT := Trim(FieldByName('CREDIT').AsString);
                L_CmdInfoObj.CREDIT_LIMIT := Trim(FieldByName('CREDIT_LIMIT').AsString);
                L_CmdInfoObj.THRESHOLD_CREDIT := Trim(FieldByName('THRESHOLD_CREDIT').AsString);
                L_CmdInfoObj.EVENT_NAME := Trim(FieldByName('EVENT_NAME').AsString);
                L_CmdInfoObj.PRICE :=  Trim(FieldByName('PRICE').AsString);
                L_CmdInfoObj.CC_NUMBER_1 := Trim(FieldByName('CC_NUMBER_1').AsString);
                L_CmdInfoObj.IP_ADDR := Trim(FieldByName('IP_ADDR').AsString);
                L_CmdInfoObj.CC_PORT := Trim(FieldByName('CC_PORT').AsString);
                L_CmdInfoObj.CALLBACK_DATE := Trim(FieldByName('CALLBACK_DATE').AsString);
                L_CmdInfoObj.CALLBACK_TIME := Trim(FieldByName('CALLBACK_TIME').AsString);
                L_CmdInfoObj.CALLBACK_FREQUENCY := Trim(FieldByName('CALLBACK_FREQUENCY').AsString);
                L_CmdInfoObj.FIRST_CALLBACK_DATE := Trim(FieldByName('FIRST_CALLBACK_DATE').AsString);
                L_CmdInfoObj.PHONE_NUM_1 := Trim(FieldByName('PHONE_NUM_1').AsString);
                L_CmdInfoObj.PHONE_NUM_2 := Trim(FieldByName('PHONE_NUM_2').AsString);
                L_CmdInfoObj.PHONE_NUM_3 := Trim(FieldByName('PHONE_NUM_3').AsString);

                L_CmdInfoObj.MIS_IRD_CMD_ID := Trim(FieldByName('MIS_IRD_CMD_ID').AsString);
                L_CmdInfoObj.MIS_IRD_CMD_DATA:= Trim(FieldByName('MIS_IRD_CMD_DATA').AsString);
                L_CmdInfoObj.COMPCODE := Trim(FieldByName('COMPCODE').AsString);

                //L_CmdInfoObj.CLEANUP_DATE := Trim(FieldByName('CLEANUP_DATE').AsString);
                //L_CmdInfoObj.CONDITION_DATE := Trim(FieldByName('CONDITION_DATE').AsString);
                //L_CmdInfoObj.COLLECT_DATE := Trim(FieldByName('COLLECT_DATE').AsString);

                addCmdInfoObj(sL_SeqNo, L_CmdInfoObj);
              end;
              Prior;
            end;
            try
              frmMain.sG_FD_SeqNos := sL_SeqNos;
              // 若不用 Synchronize => 在 Delphi 下 run..會有問題
              UpdateCmdStatusToProcessing(sL_SeqNos);
            except
              { ... }
            end;
            dtmMain.G_AdoSendNagra[frmMain.nG_CurActiveConn].Close;
          end
          else
          begin
            //down, 若資料庫沒有未處理之資料,則新增一筆連線測試資料
            bL_SendTestCmd := False;
            { 每6秒送一次, 因為 timer 是每 3 秒 觸發一次, 所以在這
              邊計算 }
            Inc( aDelayNoCommand );
            if ( aDelayNoCommand >= 5 ) then
            begin
              bL_SendTestCmd := True;
              aDelayNoCommand := 0;
            end;
            if bL_SendTestCmd then
            begin
              if frmMain.nG_FD_TestingCmdSeqNo>=frmMain.nG_FD_MaxTestingCmdSeqNo then
                frmMain.nG_FD_TestingCmdSeqNo := 1
              else
                INC(frmMain.nG_FD_TestingCmdSeqNo);
              sL_SeqNo := IntToStr(frmMain.nG_FD_TestingCmdSeqNo);
              L_CmdInfoObj := TCmdInfoObject.Create;

              L_CmdInfoObj.CMD_STATUS := 'W';
              L_CmdInfoObj.SEQNO  := sL_SeqNo;

              L_CmdInfoObj.HIGH_LEVEL_CMD_ID := CA_TESTING_CMD_ID;
              L_CmdInfoObj.ICC_NO := EmptyStr;
              L_CmdInfoObj.STB_NO := EmptyStr;
              L_CmdInfoObj.SUBSCRIPTION_BEGIN_DATE := EmptyStr;
              L_CmdInfoObj.SUBSCRIPTION_END_DATE := EmptyStr;
              L_CmdInfoObj.OPERATOR := EmptyStr;
              L_CmdInfoObj.NOTES := EmptyStr;

              L_CmdInfoObj.IMS_PRODUCT_ID := EmptyStr;
              L_CmdInfoObj.ZIP_CODE := EmptyStr;
              L_CmdInfoObj.CREDIT_MODE  := EmptyStr;
              L_CmdInfoObj.CREDIT := EmptyStr;
              L_CmdInfoObj.CREDIT_LIMIT := EmptyStr;
              L_CmdInfoObj.THRESHOLD_CREDIT := EmptyStr;
              L_CmdInfoObj.EVENT_NAME := EmptyStr;
              L_CmdInfoObj.PRICE :=  EmptyStr;
              L_CmdInfoObj.CC_NUMBER_1 := EmptyStr;
              L_CmdInfoObj.IP_ADDR := EmptyStr;
              L_CmdInfoObj.CC_PORT := EmptyStr;
              L_CmdInfoObj.CALLBACK_DATE := EmptyStr;
              L_CmdInfoObj.CALLBACK_TIME := EmptyStr;
              L_CmdInfoObj.CALLBACK_FREQUENCY := EmptyStr;
              L_CmdInfoObj.FIRST_CALLBACK_DATE := EmptyStr;
              L_CmdInfoObj.PHONE_NUM_1 := EmptyStr;
              L_CmdInfoObj.PHONE_NUM_2 := EmptyStr;
              L_CmdInfoObj.PHONE_NUM_3 := EmptyStr;

              L_CmdInfoObj.MIS_IRD_CMD_ID := EmptyStr;
              L_CmdInfoObj.MIS_IRD_CMD_DATA:= EmptyStr;
              //因為丟測試指令(cmd 1002), 所以公司別用固定值
              L_CmdInfoObj.COMPCODE := TESTING_CMD_COMP_CODE;

              L_CmdInfoObj.CLEANUP_DATE := EmptyStr;
              L_CmdInfoObj.CONDITION_DATE := EmptyStr;
              L_CmdInfoObj.COLLECT_DATE := EmptyStr;

              addCmdInfoObj(sL_SeqNo, L_CmdInfoObj);
          end;
        end;
      end;
    end;
    end;
    except
      { ...}
    end;
end;

{ ---------------------------------------------------------------------------- }

procedure FetchData;
var
  nL_Ndx : Integer;
  sL_CurActiveConn : String;
  bL_HasException : boolean;
begin

  if frmMain.G_CommandInfoStrList.Count <> 0 then
    Exit;

  if frmMain.nG_CurActiveConn = frmMain.nG_DbCount - 1 then
    frmMain.nG_FD_CurActiveConn := 0
  else
    Inc( frmMain.nG_FD_CurActiveConn );

  frmMain.nG_CurActiveConn := frmMain.nG_FD_CurActiveConn;

    try
      bL_HasException := False;
      sL_CurActiveConn := IntToStr( frmMain.nG_FD_CurActiveConn );
      nL_Ndx := frmMain.G_DisConnDbStrList.IndexOf( sL_CurActiveConn );
      if nL_Ndx <> -1 then
      begin
        if dtmMain.ConnectToDB( frmMain.nG_FD_CurActiveConn, True ) then
        begin
          frmMain.G_DisConnDbStrList.Delete( nL_Ndx );
          dtmMain.ActiveSendNagra( frmMain.nG_FD_CurActiveConn );
        end else
        begin
          bL_HasException := True;
        end;  
      end
      else
        dtmMain.ActiveSendNagra( frmMain.nG_FD_CurActiveConn );
    except
      on E: Exception do
      begin
        bL_HasException := True;
      end;
    end;

    if bL_HasException then
    begin
      frmMain.ShowLogOnMemo( frmMain.nG_FD_CurActiveConn,
        frmMain.getLanguageSettingInfo( 'FetchCmd_Thread_Msg_1_Content' ) );
      frmMain.G_CommandInfoStrList.Clear;
      nL_Ndx := frmMain.G_DisConnDbStrList.IndexOf( sL_CurActiveConn );
      if nL_Ndx = -1 then
        frmMain.G_DisConnDbStrList.Add(sL_CurActiveConn);
    end else
    begin
      try
        CopyRequestData;
      except
        { ... }
      end;
    end;
end;

{ ---------------------------------------------------------------------------- }

procedure OnTimerStartAction;
begin
 if frmMain.bG_StartFetching then //表示 timer 到了
  begin
    frmMain.bG_StartFetching := False;
    if frmMain.bG_RunFetchCommandThread then FetchData;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure UpdateCmdStatusToProcessing;
begin
  try
    dtmMain.updateCmdStatusToProcessing( frmMain.sG_FD_SeqNos,
      frmMain.nG_CurActiveConn );
  except
    frmMain.showLogOnMemo( UPDATE_CMD_STATUS_ERROR_FLAG,
      ' FetchComData.updateCmdStatusToProcessing' );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure UpdateCmdStatusToProcessing(const ASeqNos: String);
begin
  try
    dtmMain.updateCmdStatusToProcessing( ASeqNos, frmMain.nG_CurActiveConn );
  except
    frmMain.showLogOnMemo( UPDATE_CMD_STATUS_ERROR_FLAG,
      ' FetchComData.updateCmdStatusToProcessing' );
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
