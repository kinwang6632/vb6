//Syn OK
unit FetchComDataThreadU;

interface

uses
  Classes, SysUtils, Dialogs, ExtCtrls, ConstObjU, SyncObjs;

type
  FetchComDataThread = class(TThread)
  private
    { Private declarations }


  protected
    procedure Execute; override;
  public

    procedure redefineTimerInterval;
    procedure onTimerStartAction;
    procedure updateCmdStatusToProcessing;overload;
    procedure updateCmdStatusToProcessing(sI_SeqNos:String);overload;
    procedure addCmdInfoObj(sI_SeqNo:String; I_CmdInfoObj : TCmdInfoObject);overload;
    procedure fetchData;
    procedure copyRequestData;
    constructor Create;overload;
  end;

implementation

uses frmMainU, dtmMainU, Ustru;


{ Important: Methods and properties of objects in VCL or CLX can only be used
  in a method called using Synchronize, for example,

      Synchronize(UpdateCaption);

  and UpdateCaption could look like,

    procedure FetchComDataThread.UpdateCaption;
    begin
      Form1.Caption := 'Updated in a thread';
    end; }

{ FetchComDataThread }


procedure FetchComDataThread.addCmdInfoObj(sI_SeqNo: String;
  I_CmdInfoObj: TCmdInfoObject);
begin
    if Assigned(frmMain.G_CommandInfoStrList) then
    frmMain.G_CommandInfoStrList.AddObject(sI_SeqNo, I_CmdInfoObj);
end;

procedure FetchComDataThread.copyRequestData;
var
    bL_SendTestCmd : boolean;
    nL_RecCount : Integer;
    sL_SeqNo, sL_SeqNos : String;
    L_CmdInfoObj : TCmdInfoObject;

    nL_Index : Integer;
    L_StrList:TStringList;
begin
    try

    if Assigned(frmMain.G_CommandInfoStrList) and (frmMain.G_CommandInfoStrList.Count>0) then
      Exit
    else
    begin
      if frmMain.nG_CurActiveConn <0 then
        Exit
      else
      begin
        with dtmMain.G_AdoSendNagra[frmMain.nG_CurActiveConn] do
        begin
    //      First;
          Last;
          sL_SeqNos := '';
          nL_RecCount := 0;
    //      while not Eof do
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

              if nL_Index<0 then
              begin
                L_CmdInfoObj := TCmdInfoObject.Create;

                L_CmdInfoObj.CMD_STATUS := FieldByName('CMD_STATUS').AsString;
                L_CmdInfoObj.SEQNO  := sL_SeqNo;

                if sL_SeqNos='' then
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

//                frmMain.G_FD_CmdInfoObj := L_CmdInfoObj;
      //          CRS.Enter;
                addCmdInfoObj(sL_SeqNo, L_CmdInfoObj);

      //          CRS.Leave;

              end;
      //        Next;
              Prior;
            end;
            try
              frmMain.sG_FD_SeqNos := sL_SeqNos;
              updateCmdStatusToProcessing(sL_SeqNos); // 若不用 Synchronize => 在 Delphi 下 run..會有問題
//              Synchronize(updateCmdStatusToProcessing);
            except

            end;
            dtmMain.G_AdoSendNagra[frmMain.nG_CurActiveConn].Close;          
          end
          else
          begin
            //down, 若資料庫沒有未處理之資料,則新增一筆連線測試資料
            {
            if frmMain.nG_CurActiveConn =0 then
              bL_SendTestCmd := true
            else
              bL_SendTestCmd := false;
            }
    //        L_StrList:=TStringList.Create;
    //        L_StrList.Add('1');


            bL_SendTestCmd := true;
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
              L_CmdInfoObj.ICC_NO := '';
              L_CmdInfoObj.STB_NO := '';
              L_CmdInfoObj.SUBSCRIPTION_BEGIN_DATE := '';
              L_CmdInfoObj.SUBSCRIPTION_END_DATE := '';
              L_CmdInfoObj.OPERATOR := '';
              L_CmdInfoObj.NOTES := '';

              L_CmdInfoObj.IMS_PRODUCT_ID := '';
              L_CmdInfoObj.ZIP_CODE := '';
              L_CmdInfoObj.CREDIT_MODE  := '';
              L_CmdInfoObj.CREDIT := '';
              L_CmdInfoObj.CREDIT_LIMIT := '';
              L_CmdInfoObj.THRESHOLD_CREDIT := '';
              L_CmdInfoObj.EVENT_NAME := '';
              L_CmdInfoObj.PRICE :=  '';
              L_CmdInfoObj.CC_NUMBER_1 := '';
              L_CmdInfoObj.IP_ADDR := '';
              L_CmdInfoObj.CC_PORT := '';
              L_CmdInfoObj.CALLBACK_DATE := '';
              L_CmdInfoObj.CALLBACK_TIME := '';
              L_CmdInfoObj.CALLBACK_FREQUENCY := '';
              L_CmdInfoObj.FIRST_CALLBACK_DATE := '';
              L_CmdInfoObj.PHONE_NUM_1 := '';
              L_CmdInfoObj.PHONE_NUM_2 := '';
              L_CmdInfoObj.PHONE_NUM_3 := '';

              L_CmdInfoObj.MIS_IRD_CMD_ID := '';
              L_CmdInfoObj.MIS_IRD_CMD_DATA:= '';

    //          L_CmdInfoObj.COMPCODE := frmMain.G_CompCodeStrList.Strings[nG_CurActiveConn];
              L_CmdInfoObj.COMPCODE := TESTING_CMD_COMP_CODE; //因為丟測試指令(cmd 1002), 所以公司別用固定值


//              frmMain.G_FD_CmdInfoObj := L_CmdInfoObj;
      //          CRS.Enter;
    //        L_StrList.Add('2');
              addCmdInfoObj(sL_SeqNo, L_CmdInfoObj);

      //          CRS.Leave;
    //        L_StrList.Add('3');
            end;
    //        L_StrList.Add('4');
            //up, 若資料庫沒有未處理之資料,則新增一筆連線測試資料

          end;
    //L_StrList.SaveToFile('c:\temp\aa.txt');
        end;
      end;
    end;
    except
    end;
end;

constructor FetchComDataThread.Create;
begin
//    CRS := TCriticalSection.Create;
    inherited Create(false);
end;

procedure FetchComDataThread.Execute;
begin
  { Place thread code here }
    self.FreeOnTerminate := true;
    self.OnTerminate := frmMain.FetchCmdDataThreadDone;

    frmMain.nG_FD_TestingCmdSeqNo := 0;

    frmMain.nG_FD_MaxTestingCmdSeqNo := StrToIntDef(TUstr.AddString('','9',true,SEQ_NUM_LENGTH),999999);
{
2003/10/31
    onTimerStartAction;
    while true do
    begin

      if frmMain.bG_StartFetching then //表示 timer 到了
      begin
        frmMain.G_Timer.Enabled := false;
        frmMain.bG_StartFetching := false;
        if frmMain.bG_RunFetchCommandThread then
          fetchData;
        redefineTimerInterval;
//        Synchronize(redefineTimerInterval);

        frmMain.G_Timer.Enabled := true;

      end;

    end;
}
end;




procedure FetchComDataThread.fetchData;
var
    nL_Ndx : Integer;
    sL_CurActiveConn : String;
    bL_HasException : boolean;
begin
    if frmMain.G_CommandInfoStrList.Count <>0 then
      Exit;
    if frmMain.nG_CurActiveConn=frmMain.nG_DbCount-1 then
      frmMain.nG_FD_CurActiveConn := 0
    else
      Inc(frmMain.nG_FD_CurActiveConn);
    frmMain.nG_CurActiveConn := frmMain.nG_FD_CurActiveConn;

    try
      bL_HasException := false;
      sL_CurActiveConn := IntToStr(frmMain.nG_FD_CurActiveConn);
      nL_Ndx := frmMain.G_DisConnDbStrList.IndexOf(sL_CurActiveConn);
      if nL_Ndx <> -1 then
      begin
        if dtmMain.connectToDB(frmMain.nG_FD_CurActiveConn,true) then
        begin
          frmMain.G_DisConnDbStrList.Delete(nL_Ndx);
          dtmMain.activeSendNagra(frmMain.nG_FD_CurActiveConn);
        end
        else
        begin
          bL_HasException := true;
        end;  
      end
      else
        dtmMain.activeSendNagra(frmMain.nG_FD_CurActiveConn);
    except
      on E: Exception do
      begin
        {
        frmMain.G_Timer.Enabled := true;
        frmMain.showLogOnMemo(nG_CurActiveConn, '取出 CA 指令資料');
        frmMain.G_DisConnDbStrList.Add(sL_CurActiveConn);
        }
        bL_HasException := true;
        {
        try
          dtmMain.connectToDB(nG_CurActiveConn,false);
          dtmMain.connectToDB(nG_CurActiveConn,true);
        except;

        end;
        }
      end;
    end;


    if bL_HasException then
    begin
      frmMain.showLogOnMemo(frmMain.nG_FD_CurActiveConn, frmMain.getLanguageSettingInfo('FetchCmd_Thread_Msg_1_Content'));    
      frmMain.G_CommandInfoStrList.Clear;
//      frmMain.bG_RunSendingCommandThread := false;
//      frmMain.bG_RunFetchCommandThread := true;
//      frmMain.G_Timer.Enabled := true;

      nL_Ndx := frmMain.G_DisConnDbStrList.IndexOf(sL_CurActiveConn);
      if nL_Ndx = -1 then
        frmMain.G_DisConnDbStrList.Add(sL_CurActiveConn);

    end
    else
    begin
      try
        copyRequestData;
//        frmMain.bG_RunSendingCommandThread := true;
//        frmMain.bG_RunFetchCommandThread := false;

      except

      end;
    end;

end;

procedure FetchComDataThread.onTimerStartAction;
begin
    if frmMain.bG_StartFetching then //表示 timer 到了
    begin
//      frmMain.G_Timer.Enabled := false;
      frmMain.bG_StartFetching := false;
      if frmMain.bG_RunFetchCommandThread then
        fetchData;
      redefineTimerInterval;
  //        Synchronize(redefineTimerInterval);

//      frmMain.G_Timer.Enabled := true;
    end;
end;

procedure FetchComDataThread.redefineTimerInterval;
begin
//    frmMain.G_Timer.Interval := frmMain.reDefineFetchDataInterval(now);
end;




procedure FetchComDataThread.updateCmdStatusToProcessing;
begin
    try
      dtmMain.updateCmdStatusToProcessing(frmMain.sG_FD_SeqNos, frmMain.nG_CurActiveConn);
    except
      frmMain.showLogOnMemo(UPDATE_CMD_STATUS_ERROR_FLAG,' FetchComDataThread.updateCmdStatusToProcessing');
    end;
end;

procedure FetchComDataThread.updateCmdStatusToProcessing(
  sI_SeqNos: String);
begin
    try
      dtmMain.updateCmdStatusToProcessing(sI_SeqNos, frmMain.nG_CurActiveConn);
    except
      frmMain.showLogOnMemo(UPDATE_CMD_STATUS_ERROR_FLAG,' FetchComDataThread.updateCmdStatusToProcessing');
    end;
end;

end.
