unit frmMainU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Buttons, DB, ADODB, Menus;
const
    DEFAULT_CHANNEL_DAYS = 7;
    DEFAULT_PIN_CODE = '0000';
    WAREHOURSE_SUPER_USER = 'gs';
    STR_SEP = '''';
    MAX_ACTION_TIMES = 2;
    OPEN_CHANNEL = '開頻道';
type
  TServiceInfoObj = class(TObject)
    sServiceID : String;
  end;
    
  TfrmMain = class(TForm)
    Panel1: TPanel;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    StaticText1: TStaticText;
    StaticText2: TStaticText;
    edtSubscriberID: TEdit;
    edtIccNo: TEdit;
    StaticText3: TStaticText;
    BitBtn3: TBitBtn;
    BitBtn4: TBitBtn;
    BitBtn5: TBitBtn;
    Panel2: TPanel;
    StaticText4: TStaticText;
    StaticText5: TStaticText;
    StaticText6: TStaticText;
    edtAlias: TEdit;
    edtUserID: TEdit;
    edtPasswd: TEdit;
    BitBtn6: TBitBtn;
    ADOConnection1: TADOConnection;
    ADOCommand1: TADOCommand;
    qryChInfo: TADOQuery;
    memCh: TMemo;
    btnChooseCh: TBitBtn;
    qryActionTimes: TADOQuery;
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    qryReport1: TADOQuery;
    BitBtn7: TBitBtn;
    qryReport1ICC_NO: TStringField;
    qryReport1SUBSCRIBER_ID: TStringField;
    qryReport1PRODUCTID: TStringField;
    qryReport1ACTION: TStringField;
    qryReport1AUTHORSTARTDATE: TDateTimeField;
    qryReport1AUTHORSTOPDATE: TDateTimeField;
    qryReport1PINCODE: TStringField;
    qryReport1OPERATOR: TStringField;
    qryReport1UPDTIME: TDateTimeField;
    N3: TMenuItem;
    qryResult: TADOQuery;
    qryGetSeqNo: TADOQuery;
    chbShowResult: TCheckBox;
    Panel3: TPanel;
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure BitBtn4Click(Sender: TObject);
    procedure BitBtn5Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BitBtn6Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure btnChooseChClick(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure BitBtn7Click(Sender: TObject);
    procedure N3Click(Sender: TObject);
  private
    { Private declarations }
    G_ServiceInfo : TStringList;
    bG_IsSuperUser : boolean;
    G_UserNameStrList : TStringList;
    G_UserPasswdStrList : TStringList;
    sG_NewPinCode : String;
    bG_Login : boolean;
    sG_DbUserID, sG_DbUserPasswd : String;

    sG_ProductDescForCloseCh, sG_ProductDescForOpenCh : String;
    function getSeqNo:String;
    procedure RealAction(nI_ActionType:Integer);
    procedure setCursor(nI_State:Integer);

  public
    { Public declarations }
    sG_UserID : String;
    function activeReport1Data(sI_SQL:String):Integer;
    function getOracleSQLDateStr(dI_Date:TDate):String;
    function getOracleSQLDateTimeStr(dI_Date:TDateTime):String;        
  end;

var
  frmMain: TfrmMain;

implementation

uses frmDbMultiSelectu, Ustru, frmReport1U, frmParamU, UdateTimeu,
  frmCmdResultU, ConstU;

{$R *.dfm}

procedure TfrmMain.BitBtn1Click(Sender: TObject);
begin
    if MessageDlg('是否確定要結束程式?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      Application.Terminate;
end;

procedure TfrmMain.BitBtn2Click(Sender: TObject);
var
    sL_IccNo : String;
begin

    sL_IccNo := Trim(edtIccNo.Text);
    if (sL_IccNo='') then
    begin
      MessageDlg('請輸入ICC No.', mtWarning, [mbOK],0);
      edtIccNo.SetFocus;
      Exit;
    end
    else
      RealAction(1);
end;

procedure TfrmMain.BitBtn3Click(Sender: TObject);
var
    sL_SubscriberID : String;
begin
    sL_SubscriberID := Trim(edtSubscriberID.Text);

    if (sL_SubscriberID='') then
    begin
      MessageDlg('請輸入 subscriber id.', mtWarning, [mbOK],0);
      edtSubscriberID.SetFocus;
      Exit;
    end
    else
      RealAction(2);    
end;

procedure TfrmMain.BitBtn4Click(Sender: TObject);
var
    sL_SubscriberID : String;
begin
    sL_SubscriberID := Trim(edtSubscriberID.Text);

    if (sL_SubscriberID='')  then
    begin
      MessageDlg('請輸入 subscriber id.', mtWarning, [mbOK],0);
      edtSubscriberID.SetFocus;
      Exit;
    end
    else if (memCh.Lines.Count=0) then
    begin
      MessageDlg('請點選頻道.', mtWarning, [mbOK],0);
      btnChooseCh.SetFocus;
      Exit;
    end
    else
      RealAction(3);
end;

procedure TfrmMain.BitBtn5Click(Sender: TObject);
var
    sL_SubscriberID : String;
begin
    sL_SubscriberID := Trim(edtSubscriberID.Text);
    if (sL_SubscriberID='')  then
    begin
      MessageDlg('請輸入 subscriber id.', mtWarning, [mbOK],0);
      edtSubscriberID.SetFocus;
      Exit;
    end
    else if (memCh.Lines.Count=0) then
    begin
      MessageDlg('請點選頻道.', mtWarning, [mbOK],0);
      btnChooseCh.SetFocus;
      Exit;
    end    
    else
      RealAction(4);    
end;

procedure TfrmMain.FormCreate(Sender: TObject);
begin
    bG_Login := false;
    G_UserNameStrList := TStringList.Create;
    G_UserPasswdStrList := TStringList.Create;
    G_ServiceInfo := TStringList.Create;

    G_UserNameStrList.Add('a1');
    G_UserPasswdStrList.Add('1688');

    G_UserNameStrList.Add('b1');
    G_UserPasswdStrList.Add('9988');

    G_UserNameStrList.Add('c1');
    G_UserPasswdStrList.Add('3702');

    G_UserNameStrList.Add('d1');
    G_UserPasswdStrList.Add('2581');

    G_UserNameStrList.Add(WAREHOURSE_SUPER_USER); //倉管的 super user
    G_UserPasswdStrList.Add('6980'); //倉管的 super user

end;

procedure TfrmMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    G_UserNameStrList.Free;
    G_UserPasswdStrList.Free;
    G_ServiceInfo.Free;
end;

procedure TfrmMain.BitBtn6Click(Sender: TObject);
var
    sL_DbUserID, sL_DbUserPasswd : String;
    bL_Login : boolean;
    sL_DbAlias, sL_UserID, sL_Passwd, sL_DbConnStr : String;
    nL_Ndx : Integer;
begin
    sL_UserID := Trim(edtUserID.Text);
    sL_Passwd := Trim(edtPasswd.Text);
    bL_Login := false;

    nL_Ndx := G_UserNameStrList.IndexOf(sL_UserID);
    if nL_Ndx<>-1 then
    begin
      if sL_Passwd=G_UserPasswdStrList.Strings[nL_Ndx] then
        bL_Login := true
      else
        bL_Login := false;
    end
    else
      bL_Login := false;

    if not bL_Login then
    begin
      MessageDlg('您輸入的使用者名稱/使用者密碼錯誤!請重新輸入.', mtWarning, [mbOK],0);
      edtUserID.SetFocus;
      Exit;
    end
    else
    begin
      sG_UserID := sL_UserID;
      sL_DbAlias := Trim(edtAlias.Text);
      if sL_DbAlias='' then
      begin
        MessageDlg('請輸入資料庫別名.', mtWarning, [mbOK],0);
        edtAlias.SetFocus;
        Exit;
      end
      else
      begin
        sL_DbUserID := sG_DbUserID;
        sL_DbUserPasswd := sG_DbUserPasswd;
//        sL_DbConnStr := 'Provider=OraOLEDB.Oracle.1;Password=' + sL_DbUserPasswd +';Persist Security Info=True;User ID='+sL_DbUserID+';Data Source='+sL_DbAlias; //=> 適用於my nb
        sL_DbConnStr := 'Provider=MSDAORA.1;Password=' + sL_DbUserPasswd + ';User ID='+sL_DbUserID+';Data Source='+sL_DbAlias+';Persist Security Info=True'; //適用於其他 PC
        try
          if ADOConnection1.Connected then
            ADOConnection1.Connected := false;
          ADOConnection1.ConnectionString := sL_DbConnStr;
          ADOConnection1.Connected := true;
        except
          MessageDlg('資料庫連接失敗!', mtError, [mbOK],0);
          edtAlias.SetFocus;
          Exit;
        end;
        bG_Login := true;
        edtSubscriberID.setFocus;
        if sL_UserID = WAREHOURSE_SUPER_USER then
          bG_IsSuperUser := true
        else
          bG_IsSuperUser := false;
        MessageDlg('資料庫連接成功!', mtInformation, [mbOK],0);
      end;
    end;
end;

procedure TfrmMain.RealAction(nI_ActionType: Integer);
var
    L_TmpStrList1, L_TmpStrList2 : TStringList;
    sL_ChannelDays, sL_MisIrdCmdData,sL_MisIrdCmdID, sL_PinCode : String;
    ii,nL_ChannelDays, nL_Times : Integer;
    bL_CheckTimes, bL_CouldBeExecuted,bL_Exit : boolean;
    sL_CompName, sL_CompCode, sL_SeqNo, sL_Notes, sL_CmdStatus : String;
    sL_AuthorStartDate, sL_AuthorStopDate1, sL_AuthorStopDate2 : String;
    sL_Action, sL_RealSQL, sL_SQL, sL_LogSQL, sL_AlermInfo : String;
    sL_SubscriberID, sL_IccNo,sL_ProductID, sL_PersonalRegionLength : String;
    sL_Operator, sL_Now, sL_HighLevelCmdID : String;
    sL_GetResultSQL, sL_ReportBackAvailability, sL_PinControl : String;
    sL_ReportBackDate, sL_RegionKey, sL_PopulationId : String;
    sL_ErrorCode, sL_ErrorDesc, sL_ResponseData : String;
    sL_ActionResponse, sL_ServiceID : String;
    sL_CreatedSubscriberID, sL_StatusCode : String;
    L_Obj : TServiceInfoObj;
    sL_CardID, sL_NewCardID, sL_Date, sL_ResultRegionKey, sL_PersonalRegion, sL_ResultReportbackAvailability : String;
    sL_ReportbackDay, sL_ReportbackTime, sL_Currency : String;
begin
    if not bG_Login then
    begin
      edtAlias.SetFocus;
      MessageDlg('請先登入系統!', mtWarning, [mbOK],0);
      Exit;
    end;
    bL_Exit := false;
    bL_CheckTimes := false;
    bL_CouldBeExecuted := true;
    sL_SubscriberID := Trim(edtSubscriberID.Text);
    sL_IccNo := Trim(edtIccNo.Text);
    sL_Operator := sG_UserID;
    sL_CmdStatus := 'W';
    sL_CompCode := '7';
    sL_CompName := '開博科技';
    sL_SeqNo := getSeqNo;
    setCursor(1);
    case nI_ActionType of
      1://開機
       begin
         sL_Action :='開機';
         sL_AuthorStartDate := 'NULL';
         sL_AuthorStopDate1 :=  'NULL';
         sL_ProductID := '';
         sL_HighLevelCmdID := 'A1';
         sL_Notes := '';
         sL_MisIrdCmdData := '';
         sL_MisIrdCmdID := '';
         sL_PinCode := '';
         sL_ReportBackAvailability := frmParam.getReportbackAvailability;
         sL_ReportBackDate := frmParam.getReportbackDate;
         sL_RegionKey := frmParam.getRegionKey;
         sL_PopulationId := frmParam.getPopulationID;


         sL_RealSQL := 'insert into SEND_NDS(SeqNo,CompCode,CompName,CommandId,IccNo, ' +
                   'CmdStatus,Operator,ProcessingDate, UpdTime, RegionKey, ReportBackAvailability,' +
                   'ReportBackDate, PopulationId) values ' +
                    ' (' + sL_SeqNo + ' ,' + STR_SEP + sL_CompCode + STR_SEP + ',' + STR_SEP + sL_CompName + STR_SEP + ',' +
                    STR_SEP + sL_HighLevelCmdID + STR_SEP + ',' + STR_SEP + sL_IccNo + STR_SEP + ',' +
                    STR_SEP + sL_CmdStatus + STR_SEP + ',' + STR_SEP + sL_CompName + STR_SEP +
                    ',NULL,sysdate,' + STR_SEP + sL_RegionKey + STR_SEP + ',' +
                    STR_SEP + sL_ReportBackAvailability + STR_SEP + ',' +
                    STR_SEP + sL_ReportBackDate + STR_SEP + ',' +
                    STR_SEP + sL_PopulationId + STR_SEP + ')';
       end;
      2://關機
       begin
         sL_Action :='關機';
         sL_AuthorStartDate := 'NULL';
         sL_AuthorStopDate1 :=  'NULL';
         sL_ProductID := '';
         sL_HighLevelCmdID := 'A2';
         sL_Notes := '';
         sL_MisIrdCmdData := '';
         sL_MisIrdCmdID := '';
         sL_PinCode := '';
         sL_RealSQL := 'insert into SEND_NDS(SeqNo,CompCode,CompName,CommandId, ' +
                   'CmdStatus,Operator,ProcessingDate, UpdTime, SubscriberId) values ' +
                    ' (' + sL_SeqNo + ',' + STR_SEP + sL_CompCode + STR_SEP + ',' + STR_SEP + sL_CompName + STR_SEP + ',' +
                    STR_SEP + sL_HighLevelCmdID + STR_SEP + ',' +
                    STR_SEP + sL_CmdStatus + STR_SEP + ',' + STR_SEP + sL_CompName + STR_SEP +
                    ',NULL,sysdate,' + STR_SEP + sL_SubscriberID + STR_SEP + ')';

       end;
      3://開頻道
       begin
         if not bG_IsSuperUser then //若不是 super user => 只能開頻道兩次
         begin
           nL_ChannelDays := DEFAULT_CHANNEL_DAYS;
           bL_CheckTimes := true;
         end
         else
         begin
           sL_ChannelDays := IntToStr(DEFAULT_CHANNEL_DAYS);
           if InputQuery('請輸入開頻道的天數', '開頻道的天數:',sL_ChannelDays) then
           begin
             nL_ChannelDays := StrToIntDef(sL_ChannelDays,-1);
             if nL_ChannelDays = -1 then
             begin
               bL_Exit := true;
               MessageDlg('開頻道的天數應為數字!', mtWarning, [mbOK],0);
             end;
           end
           else
             bL_Exit := true;
         end;
         if bL_Exit then
         begin
           setCursor(2);         
           Exit;
         end;

         sL_AuthorStartDate := getOracleSQLDateStr(now);
         sL_AuthorStopDate1 := getOracleSQLDateStr(now+nL_ChannelDays);
         sL_AuthorStopDate2 := TUdateTime.GetPureDateStr(now+nL_ChannelDays,'');
         sL_Action := OPEN_CHANNEL;
//         sL_ProductID := '全部';
         sL_ProductID := sG_ProductDescForOpenCh;

         sL_HighLevelCmdID := 'B1';

         sL_MisIrdCmdData := '';
         sL_MisIrdCmdID := '';
         sL_PinCode := '';

         //Notes 範例: 4~20040531~N;5~20041231~A
         sL_Notes := '';
         for ii:=0 to G_ServiceInfo.Count -1 do
         begin
           L_Obj := (G_ServiceInfo.Objects[ii] as TServiceInfoObj);
           sL_ServiceID := L_Obj.sServiceID;
           
           if sL_Notes='' then
             sL_Notes := sL_ServiceID + '~' + sL_AuthorStopDate2 + '~' + 'N'
           else
             sL_Notes := sL_Notes + ';' + sL_ServiceID + '~' + sL_AuthorStopDate2 + '~' + 'N';
         end;

         sL_RealSQL := 'insert into SEND_NDS(SeqNo,CompCode,CompName,CommandId, ' +
                   'CmdStatus,Operator,ProcessingDate, UpdTime, Notes, SubscriberId) values ' +
                    ' (' + sL_SeqNo + ',' + STR_SEP + sL_CompCode + STR_SEP + ',' + STR_SEP + sL_CompName + STR_SEP + ',' +
                    STR_SEP + sL_HighLevelCmdID + STR_SEP + ',' +
                    STR_SEP + sL_CmdStatus + STR_SEP + ',' + STR_SEP + sL_CompName + STR_SEP +
                    ',NULL,sysdate,'  + STR_SEP + sL_Notes + STR_SEP + ',' + STR_SEP + sL_SubscriberID + STR_SEP + ')';



       end;
      4://關頻道
       begin
         sL_AuthorStartDate := 'NULL';
         sL_AuthorStopDate1 :=  'NULL';
         sL_Action :='關頻道';
//         sL_ProductID := '全部';
         sL_ProductID := sG_ProductDescForCloseCh;
         sL_HighLevelCmdID := 'B2';
//         sL_Notes := sG_ProductIdForCloseCh;
         sL_MisIrdCmdData := '';
         sL_MisIrdCmdID := '';
         sL_PinCode := '';

         for ii:=0 to G_ServiceInfo.Count -1 do
         begin
           L_Obj := (G_ServiceInfo.Objects[ii] as TServiceInfoObj);
           sL_ServiceID := L_Obj.sServiceID;
           if sL_Notes='' then
             sL_Notes := sL_ServiceID
           else
             sL_Notes := sL_Notes + ',' + sL_ServiceID;
         end;

         sL_RealSQL := 'insert into SEND_NDS(SeqNo,CompCode,CompName,CommandId, ' +
                   'CmdStatus,Operator,ProcessingDate, UpdTime, Notes, SubscriberId) values ' +
                    ' (' + sL_SeqNo + ',' + STR_SEP + sL_CompCode + STR_SEP + ',' + STR_SEP + sL_CompName + STR_SEP + ',' +
                    STR_SEP + sL_HighLevelCmdID + STR_SEP + ',' +
                    STR_SEP + sL_CmdStatus + STR_SEP + ',' + STR_SEP + sL_CompName + STR_SEP +
                    ',NULL,sysdate,'  + STR_SEP + sL_Notes + STR_SEP + ',' + STR_SEP + sL_SubscriberID + STR_SEP + ')';



       end;
      5://設定親子密碼
       begin
         sL_AuthorStartDate := 'NULL';
         sL_AuthorStopDate1 :=  'NULL';
         sL_Action :='設定親子密碼';
         sL_ProductID := '';
         sL_HighLevelCmdID := 'E1';
         sL_Notes := '';
         sL_MisIrdCmdID := '1';
         sL_PinControl := 'L';
         
         if sG_NewPinCode='' then
           sL_PinCode := DEFAULT_PIN_CODE
         else
           sL_PinCode := sG_NewPinCode;
         sL_MisIrdCmdData := sL_PinCode;

         sL_RealSQL := 'insert into SEND_NDS(SeqNo,CompCode,CompName,CommandId, ' +
                   'CmdStatus,Operator,ProcessingDate, UpdTime, SubscriberId, PinCode, PinControl) values ' +
                    ' (' + STR_SEP + sL_SeqNo + STR_SEP + ',' + sL_CompCode + ',' + STR_SEP + sL_CompName + STR_SEP + ',' +
                    STR_SEP + sL_HighLevelCmdID + STR_SEP + ',' +
                    STR_SEP + sL_CmdStatus + STR_SEP + ',' + STR_SEP + sL_CompName + STR_SEP +
                    ',NULL,sysdate,'  +  STR_SEP + sL_SubscriberID + STR_SEP +
                    ',' + sL_PinCode + ',' + STR_SEP + sL_PinControl + STR_SEP  +')';

       end;

    end;

    try
      if bL_CheckTimes then
      begin
        sL_SQL := ' select count(*) ACTIONTIMES from NDS002 where SUBSCRIBER_ID=' + STR_SEP + sL_SubscriberID + STR_SEP ;
        sL_SQL := sL_SQL + ' and ACTION=' + STR_SEP + OPEN_CHANNEL + STR_SEP;
        qryActionTimes.SQL.Clear;
        qryActionTimes.SQL.Add(sL_SQL);
        qryActionTimes.Open;
        nL_Times := qryActionTimes.FieldByName('ACTIONTIMES').AsInteger;

        if nL_Times >= MAX_ACTION_TIMES then
        begin
          sL_AlermInfo := '此 Subscriber 已經執行過 ' + IntToStr(nL_Times)+' 次開頻道的指令了!所以無法再次開頻道.';
          bL_CouldBeExecuted := false;
        end
        else
          bL_CouldBeExecuted := true;
        qryActionTimes.Close;
      end;

      if bL_CouldBeExecuted then
      begin
        sL_LogSQL := 'insert into NDS002(ICC_NO,SUBSCRIBER_ID,ProductID,ACTION,AUTHORSTARTDATE,AUTHORSTOPDATE,OPERATOR,UPDTIME,PINCODE)';
        sL_LogSQL := sL_LogSQL +' values(' + STR_SEP + sL_IccNo + STR_SEP + ',' + STR_SEP + sL_SubscriberID + STR_SEP + ',';
        sL_LogSQL := sL_LogSQL + STR_SEP + sL_ProductID + STR_SEP + ',' +  STR_SEP + sL_Action + STR_SEP + ',';
        sL_LogSQL := sL_LogSQL + sL_AuthorStartDate + ',' + sL_AuthorStopDate1 + ',' ;
        sL_LogSQL := sL_LogSQL + STR_SEP + sL_Operator +  STR_SEP + ',' + getOracleSQLDateTimeStr(now) + ',' + STR_SEP + sL_PinCode + STR_SEP +')';



        ADOCommand1.CommandText := sL_LogSQL;
        ADOCommand1.Execute;

        ADOCommand1.CommandText := sL_RealSQL;
        ADOCommand1.Execute;

        sL_GetResultSQL := 'select CMDSTATUS, ERRORCODE,ERRORDESC, RESPONSEDATA from SEND_NDS where SEQNO=' + STR_SEP + sL_SeqNo +STR_SEP ;
        sL_GetResultSQL := sL_GetResultSQL + ' and COMPCODE=' + sL_CompCode;

        with qryResult do
        begin

          while true do
          begin

            sleep(2000);
            Close;
            SQL.Clear;
            SQL.Add(sL_GetResultSQL);
            Open;
            if (RecordCount =1) and ((FieldByName('CMDSTATUS').asString='C') or (FieldByName('CMDSTATUS').asString='E')) then
              break;
          end;
          sL_ErrorCode := FieldByName('ERRORCODE').asString;
          sL_ErrorDesc := FieldByName('ERRORDESC').asString;
          sL_ResponseData := FieldByName('RESPONSEDATA').asString;
          Close;
        end;

        if nI_ActionType=1 then
        begin
          L_TmpStrList1 := TUstr.ParseStrings(sL_ResponseData,SUBSCRIBER_DATA_SEP_FLAG);
          if L_TmpStrList1.Count =2 then
          begin
            L_TmpStrList2 := TUstr.ParseStrings(L_TmpStrList1.Strings[0],SUBSCRIBER_INFO_SEP_FLAG);
            if L_TmpStrList2.Count = 2 then
            begin
              sL_CreatedSubscriberID := L_TmpStrList2.Strings[0];
              sL_StatusCode := L_TmpStrList2.Strings[1];
            end
            else
            begin
              L_TmpStrList1.Free;
              L_TmpStrList2.Free;
              MessageDlg(sL_Action + '指令傳回結果格式錯誤,無法解讀!結果字串:' +sL_ResponseData, mtError, [mbOK],0);
              Exit;
            end;
            L_TmpStrList2.Free;


            L_TmpStrList2 := TUstr.ParseStrings(L_TmpStrList1.Strings[1],SUBSCRIBER_INFO_SEP_FLAG);
            if L_TmpStrList2.Count = 10 then
            begin
              sL_CardID := L_TmpStrList2.Strings[0];
              sL_NewCardID := L_TmpStrList2.Strings[1];
              sL_Date := L_TmpStrList2.Strings[2];
              sL_ResultRegionKey := L_TmpStrList2.Strings[3];
              sL_PersonalRegionLength := L_TmpStrList2.Strings[4];
              sL_PersonalRegion := L_TmpStrList2.Strings[5];
              sL_ResultReportbackAvailability := L_TmpStrList2.Strings[6];
              sL_ReportbackDay := L_TmpStrList2.Strings[7];
              sL_ReportbackTime := L_TmpStrList2.Strings[8];
              sL_Currency := L_TmpStrList2.Strings[9];

            end
            else
            begin
              setCursor(2);
              L_TmpStrList1.Free;
              L_TmpStrList2.Free;
              MessageDlg(sL_Action + ' 指令傳回結果格式錯誤,無法解讀!結果字串:' +sL_ResponseData, mtError, [mbOK],0);
              Exit;
            end;
            if chbShowResult.Checked then
            begin
              edtSubscriberID.Text := sL_CreatedSubscriberID;
              frmCmdResult := TfrmCmdResult.Create(Application);
              frmCmdResult.Edit1.Text := sL_CreatedSubscriberID;
              frmCmdResult.Edit2.Text := sL_CardID;
              frmCmdResult.Edit3.Text := sL_NewCardID;
              frmCmdResult.Edit4.Text := sL_Date;
              frmCmdResult.Edit5.Text := sL_ResultRegionKey;
              frmCmdResult.Edit6.Text := sL_PersonalRegion;
              frmCmdResult.Edit7.Text := sL_ResultReportbackAvailability;
              frmCmdResult.Edit8.Text := sL_ReportbackDay;
              frmCmdResult.Edit9.Text := sL_ReportbackTime;
              frmCmdResult.Edit10.Text := sL_Currency;
              if sL_StatusCode='1' then
                frmCmdResult.StaticText11.Caption := 'create subscriber 完成'
              else if sL_StatusCode='3' then
                frmCmdResult.StaticText11.Caption := 'subscriber 已經存在'
              else
                frmCmdResult.StaticText11.Caption := '無法解讀 subscriber status';
              frmCmdResult.ShowModal;
              frmCmdResult.Free;
            end
            else
            begin
              if sL_StatusCode='1' then
              begin
                edtSubscriberID.Text := sL_CreatedSubscriberID;
                sL_ActionResponse := sL_Action + ' 成功!';
//                MessageDlg(sL_Action + '　成功!', mtInformation, [mbOK],0);
                Panel3.Caption := sL_ActionResponse;
              end
              else if sL_StatusCode='3' then
                MessageDlg(sL_Action + '　失敗! 因為subscriber 已經存在.', mtError, [mbOK],0)
              else
                MessageDlg(sL_Action + '　失敗! 因為無法解讀 subscriber status.', mtError, [mbOK],0);

            end;
          end
          else
          begin
            setCursor(2);
            L_TmpStrList1.Free;
            {
            sL_ActionResponse := sL_Action + ' 指令傳回結果格式錯誤,無法解讀!結果字串:' +sL_ResponseData;
            Panel3.Caption := sL_ActionResponse;
            }
            MessageDlg(sL_Action + ' 指令傳回結果格式錯誤,無法解讀!結果字串:' +sL_ResponseData, mtError, [mbOK],0);

            Exit;
          end;

        end
        else
        begin
          if sL_ErrorCode='' then
          begin
            sL_ActionResponse := sL_Action + ' 成功!';
            Panel3.Caption := sL_ActionResponse;
//            MessageDlg(sL_Action + ' 成功!', mtInformation, [mbOK],0)
          end
          else
            MessageDlg(sL_Action + ' 錯誤!錯誤代碼: ' +sL_ErrorCode + ' ,錯誤訊息: ' + sL_ErrorDesc, mtError, [mbOK],0);

        end;

      end
      else
        MessageDlg(sL_AlermInfo, mtWarning, [mbOK],0);
    except
      MessageDlg(sL_Action + ' 錯誤!', mtError, [mbOK],0);
    end;
    setCursor(2);
    edtSubscriberID.SetFocus;
end;

function TfrmMain.getOracleSQLDateTimeStr(dI_Date: TDateTime): String;
var
    wL_Year, wL_Month, wL_Day : Word;
    wL_Hour, wL_Min, wL_Sec, wL_MSec : Word;
    sL_Result, sL_DateTimeStr : String;
begin
    if dI_Date=0 then
      sL_Result := 'null'
    else
    begin
      DecodeDate(dI_Date,wL_Year, wL_Month, wL_Day);
      DecodeTime(dI_Date, wL_Hour, wL_Min, wL_Sec, wL_MSec);
      sL_DateTimeStr := format('%.4d/%.2d/%.2d %.2d:%.2d:%.2d', [wL_Year, wL_Month, wL_Day,wL_Hour, wL_Min, wL_Sec ]);
      sL_Result := 'to_date(' + '''' + sL_DateTimeStr + '''' + ',' + '''' + 'YYYY/MM/DD HH24:MI:SS' + '''' + ')';
    end;
    result := sL_Result;
end;

procedure TfrmMain.FormShow(Sender: TObject);
begin
    edtAlias.SetFocus;
    sG_DbUserID := 'system';
    sG_DbUserPasswd := 'manager';
end;

function TfrmMain.getOracleSQLDateStr(dI_Date: TDate): String;
var
    wL_Year, wL_Month, wL_Day : Word;
    sL_Result, sL_DateStr : String; 
begin
    if (dI_Date=0) or (VarIsNull(dI_Date)) then
      sL_Result := 'null'
    else
    begin
      DecodeDate(dI_Date,wL_Year, wL_Month, wL_Day);
      sL_DateStr := format('%.4d/%.2d/%.2d', [wL_Year, wL_Month, wL_Day]);
      sL_Result := 'to_date(' + '''' + sL_DateStr + '''' + ',' + '''' + 'YYYY/MM/DD' + '''' + ')';
    end;
    result := sL_Result;
end;

procedure TfrmMain.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
    bL_Change1, bL_Change2 : boolean;
    sL_DbUserID, sL_DbUSerPasswd : String;
begin

    if (ssCtrl in Shift) and (Key=VK_F1) then //Ctrl + F1
    begin
      sL_DbUserID := '';
      sL_DbUSerPasswd := '';
      bL_Change1 := false;
      bL_Change2 := false;
      if InputQuery('請輸入資料庫使用者ID', 'ID:',sL_DbUserID) then
      begin
        sG_DbUserID := sL_DbUserID;
        bL_Change1 := true;
      end;

      if InputQuery('請輸入資料庫使用者密碼', 'Password:',sL_DbUSerPasswd) then
      begin
        sG_DbUserPasswd := sL_DbUSerPasswd;
        bL_Change2 := true;
      end;
      
      if bL_Change1 or bL_Change2 then
      begin
        BitBtn6.SetFocus;
        MessageDlg('請重新登入!,才能使得設定生效', mtInformation, [mbOK],0);
      end;
    end;
end;


procedure TfrmMain.btnChooseChClick(Sender: TObject);
var
    L_Obj : TServiceInfoObj;
    sL_ChannelID, sL_Desc : String;
    ii : Integer;
    sL_SQL, sL_TargetSQL, sL_Tmp : String;
    L_TempStrList, L_TargetFieldNamesStrList, L_TargetFieldValueStrList : TStringList;
begin
    if not ADOConnection1.Connected then
    begin
      edtAlias.SetFocus;
      MessageDlg('請先登入系統!', mtWarning, [mbOK],0);
      Exit;
    end;

    with qryChInfo do
    begin
      sL_SQL := 'select DESCRIPTION, CHANNELID from cd024 ';
      Close;                        
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;
      FieldByName('DESCRIPTION').DisplayLabel := '頻道名稱';
    end;
    if qryChInfo.RecordCount = 0 then
    begin
      qryChInfo.Close;
      MessageDlg('沒有頻道可供選取!!', mtWarning, [mbOK],0);
      Exit;
    end
    else
    begin
      L_TargetFieldNamesStrList := TStringList.Create;
      L_TargetFieldValueStrList := TStringList.Create;

      L_TargetFieldNamesStrList.Add('DESCRIPTION');
      L_TargetFieldNamesStrList.Add('CHANNELID');
      if SelectMultiRecords('請點選公司別', qryChInfo, 'DESCRIPTION', ',', L_TargetFieldNamesStrList, L_TargetFieldValueStrList,true,sL_TargetSQL) = mrOk then
      begin
        memCh.Clear;
        sG_ProductDescForOpenCh := '';
        sG_ProductDescForCloseCh := '';

        for ii:=0 to L_TargetFieldValueStrList.Count -1 do
        begin
          sL_Tmp := Trim(L_TargetFieldValueStrList.Strings[ii]);
          L_TempStrList := TUstr.ParseStrings(sL_Tmp,',');

          sL_Desc := L_TempStrList.Strings[0];
          sL_ChannelID := L_TempStrList.Strings[1];
          L_Obj := TServiceInfoObj.Create;
          L_Obj.sServiceID := sL_ChannelID;
          memCh.Lines.Add(sL_Desc);
          G_ServiceInfo.AddObject(sL_Desc, L_Obj);

          
          if sG_ProductDescForOpenCh='' then
            sG_ProductDescForOpenCh := sL_Desc
          else
            sG_ProductDescForOpenCh := sG_ProductDescForOpenCh + ',' + sL_Desc;

          if sG_ProductDescForCloseCh='' then
            sG_ProductDescForCloseCh := sL_Desc
          else
            sG_ProductDescForCloseCh := sG_ProductDescForCloseCh + ',' + sL_Desc;




        end;

      end;
      L_TargetFieldNamesStrList.Free;
      L_TargetFieldValueStrList.Free;

    end;


end;

procedure TfrmMain.N2Click(Sender: TObject);
var
    L_Frm : TfrmReport1;
begin
    if not ADOConnection1.Connected then
    begin
      edtAlias.SetFocus;
      MessageDlg('請先登入系統!', mtWarning, [mbOK],0);
      Exit;
    end
    else
    begin
      L_Frm := TfrmReport1.Create(Application);
      L_Frm.ShowModal;
      L_Frm.Free;
    end;
end;

function TfrmMain.activeReport1Data(sI_SQL: String): Integer;
var
    nL_RecordCount : Integer;
begin
    with qryReport1 do
    begin
      Close;
      SQL.Clear;
      SQL.Add(sI_SQL);
      Open;
      nL_RecordCount := RecordCount;
      if nL_RecordCount=0 then
        Close;
    end;


    result := nL_RecordCount;
end;

procedure TfrmMain.BitBtn7Click(Sender: TObject);
var
    sL_SunscriberID : String;
begin
    sL_SunscriberID := Trim(edtSubscriberID.Text);
    if (sL_SunscriberID='') then
    begin
      MessageDlg('請輸入 subscriber id.', mtWarning, [mbOK],0);
      edtSubscriberID.SetFocus;
      Exit;
    end
    else
    begin
      sG_NewPinCode := DEFAULT_PIN_CODE;
      if InputQuery('請輸入新的親子密碼', '親子密碼:',sG_NewPinCode) then
      begin
        if (length(sG_NewPinCode)<>length(DEFAULT_PIN_CODE)) and (length(sG_NewPinCode)<>0) then
          MessageDlg('請輸入' + IntToStr(length(DEFAULT_PIN_CODE)) + '碼之 PinCode', mtWarning,[mbOK],0)
        else
          RealAction(5);
      end;
    end;
end;

procedure TfrmMain.N3Click(Sender: TObject);
begin
    frmParam.Show;
end;

function TfrmMain.getSeqNo: String;
begin
    with qryGetSeqNo do
    begin
      Close;
      SQL.Clear;
      SQL.Add('select S_SEND_NDS_SEQNO.NEXTVAL from dual');
      Open;
      result := Fields[0].AsString;
      Close;
    end;
end;

procedure TfrmMain.setCursor(nI_State: Integer);
begin
    case nI_State of
      1:
        Screen.Cursor := crHourGlass ;
      2:
        Screen.Cursor := crDefault ;      
    end;
end;

end.
