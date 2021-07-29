unit frmMainU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Buttons, DB, ADODB, Menus, DBClient, inifiles,
  Registry;
const
    DEFAULT_CHANNEL_DAYS = 7;
    DEFAULT_PIN_CODE = '0000';
    { 正式區給倉管用 }
    WAREHOURSE_SUPER_USER_ID = 'gs';
    WAREHOURSE_SUPER_USER_PASSWD = '6980';
    STR_SEP = '''';
    MAX_ACTION_TIMES = 2;

type
  TfrmMain = class(TForm)
    Panel1: TPanel;
    btnExit: TBitBtn;
    btnPair: TBitBtn;
    StaticText1: TStaticText;
    StaticText2: TStaticText;
    edtStbNo: TEdit;
    edtIccNo: TEdit;
    Lab_4: TStaticText;
    btnDepair: TBitBtn;
    btnOpenCh: TBitBtn;
    btnCloseCh: TBitBtn;
    Panel2: TPanel;
    Lab_1: TStaticText;
    Lab_2: TStaticText;
    Lab_3: TStaticText;
    edtAlias: TEdit;
    edtUserID: TEdit;
    edtPasswd: TEdit;
    btnLogin: TBitBtn;
    ADOConnection1: TADOConnection;
    ADOCommand1: TADOCommand;
    qryChInfo: TADOQuery;
    memCh: TMemo;
    btnChooseCh: TBitBtn;
    qryActionTimes: TADOQuery;
    MainMenu1: TMainMenu;
    mitFun1Header: TMenuItem;
    mitFun1: TMenuItem;
    qryReport1: TADOQuery;
    qryReport1ICC_NO: TStringField;
    qryReport1STB_NO: TStringField;
    qryReport1PRODUCTID: TStringField;
    qryReport1ACTION: TStringField;
    qryReport1AUTHORSTARTDATE: TDateTimeField;
    qryReport1AUTHORSTOPDATE: TDateTimeField;
    qryReport1OPERATOR: TStringField;
    qryReport1UPDTIME: TDateTimeField;
    btnPinCode: TBitBtn;
    mitFun2: TMenuItem;
    procedure btnExitClick(Sender: TObject);
    procedure btnPairClick(Sender: TObject);
    procedure btnDepairClick(Sender: TObject);
    procedure btnOpenChClick(Sender: TObject);
    procedure btnCloseChClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnLoginClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure btnChooseChClick(Sender: TObject);
    procedure mitFun1Click(Sender: TObject);
    procedure btnPinCodeClick(Sender: TObject);
    procedure mitFun2Click(Sender: TObject);
  private
    { Private declarations }
    bG_IsSuperUser : boolean;
    sG_NewPinCode : String;
    bG_Login : boolean;
    sG_DbUserID, sG_DbUserPasswd : String;
    sG_ProductIdForOpenCh, sG_ProductIdForCloseCh : String;
    sG_ProductDescForOpenCh, sG_ProductdescForCloseCh : String;
    G_LanguageSettingIni : TInifile;
    procedure RealAction(nI_ActionType:Integer);
    procedure setLanguageString;
    procedure initLanguageIniFile;

  public
    { Public declarations }
    sG_UserID : String;
    function getLanguageSettingInfo(sI_Key:String):String;    
    function activeReport1Data(sI_SQL:String):Integer;
    function getOracleSQLDateStr(dI_Date:TDate):String;
    function getOracleSQLDateTimeStr(dI_Date:TDateTime):String;        
  end;

var
  frmMain: TfrmMain;

implementation

uses frmDbMultiSelectu, Ustru, frmReport1U, frmSysParamU, dtmMainU,
  ConstObjU;

{$R *.dfm}

procedure TfrmMain.btnExitClick(Sender: TObject);
begin
    Application.Terminate;
end;

procedure TfrmMain.btnPairClick(Sender: TObject);
var
    sL_StbNo, sL_IccNo : String;
begin
    sL_StbNo := Trim(edtStbNo.Text);
    sL_IccNo := Trim(edtIccNo.Text);
    if (sL_StbNo='') or (sL_IccNo='') then
    begin
      MessageDlg(getLanguageSettingInfo('Sim_CA_Msg_1_Content'), mtWarning, [mbOK],0);
      edtStbNo.SetFocus;
      Exit;
    end
    else
      RealAction(1);
end;

procedure TfrmMain.btnDepairClick(Sender: TObject);
var
    sL_StbNo, sL_IccNo : String;
begin
    sL_StbNo := Trim(edtStbNo.Text);
    sL_IccNo := Trim(edtIccNo.Text);
    if (sL_StbNo='') or (sL_IccNo='') then
    begin
      MessageDlg(getLanguageSettingInfo('Sim_CA_Msg_1_Content'), mtWarning, [mbOK],0);
      edtStbNo.SetFocus;
      Exit;
    end
    else
      RealAction(2);    
end;

procedure TfrmMain.btnOpenChClick(Sender: TObject);
var
    sL_StbNo, sL_IccNo : String;
begin
    sL_StbNo := Trim(edtStbNo.Text);
    sL_IccNo := Trim(edtIccNo.Text);

    if (sL_StbNo='') or (sL_IccNo='') or (memCh.Lines.Count=0) then
    begin
      MessageDlg(getLanguageSettingInfo('Sim_CA_Msg_1_Content'), mtWarning, [mbOK],0);
      edtStbNo.SetFocus;
      Exit;
    end
    else
      RealAction(3);    
end;

procedure TfrmMain.btnCloseChClick(Sender: TObject);
var
    sL_StbNo, sL_IccNo : String;
begin
    sL_StbNo := Trim(edtStbNo.Text);
    sL_IccNo := Trim(edtIccNo.Text);
    if (sL_StbNo='') or (sL_IccNo='') or (memCh.Lines.Count=0) then
    begin
      MessageDlg(getLanguageSettingInfo('Sim_CA_Msg_1_Content'), mtWarning, [mbOK],0);
      edtStbNo.SetFocus;
      Exit;
    end
    else
      RealAction(4);    
end;

procedure TfrmMain.FormCreate(Sender: TObject);
begin

    initLanguageIniFile;
    setLanguageString;

    bG_Login := false;

end;

procedure TfrmMain.btnLoginClick(Sender: TObject);
var
    sL_DbUserID, sL_DbUserPasswd : String;
    bL_Login : boolean;
    sL_DbAlias, sL_UserID, sL_Passwd, sL_DbConnStr : String;
    nL_Ndx : Integer;
begin
    sL_UserID := Trim(edtUserID.Text);
    sL_Passwd := Trim(edtPasswd.Text);
    bL_Login := false;

    if (sL_UserID=WAREHOURSE_SUPER_USER_ID) and  (sL_Passwd=WAREHOURSE_SUPER_USER_PASSWD) then
    begin
      bL_Login := true;
      bG_IsSuperUser := true;      
    end
    else
      bL_Login := dtmMain.isCorrectID(sL_UserID,sL_Passwd);

    if not bL_Login then
    begin
      MessageDlg(getLanguageSettingInfo('Sim_CA_Msg_2_Content'), mtWarning, [mbOK],0);
      edtUserID.SetFocus;
      Exit;
    end
    else
    begin
      sG_UserID := sL_UserID;
      sL_DbAlias := Trim(edtAlias.Text);
      if sL_DbAlias='' then
      begin
        MessageDlg(getLanguageSettingInfo('Sim_CA_Msg_3_Content'), mtWarning, [mbOK],0);
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
          MessageDlg(getLanguageSettingInfo('Sim_CA_Msg_4_Content'), mtError, [mbOK],0);
          edtAlias.SetFocus;
          Exit;
        end;
        bG_Login := true;
        edtStbNo.setFocus;

        MessageDlg(getLanguageSettingInfo('Sim_CA_Msg_5_Content'), mtInformation, [mbOK],0);
      end;
    end;
end;

procedure TfrmMain.RealAction(nI_ActionType: Integer);
var
    sL_ChannelDays, sL_MisIrdCmdData,sL_MisIrdCmdID, sL_PinCode : String;
    nL_ChannelDays, nL_Times : Integer;
    bL_CheckTimes, bL_CouldBeExecuted,bL_Exit : boolean;
    sL_CompCode, sL_SeqNo, sL_Notes, sL_CmdStatus : String;
    sL_AuthorStartDate, sL_AuthorStopDate : String;
    sL_Action, sL_SQL, sL_LogSQL, sL_AlermInfo : String;
    sL_StbNo, sL_IccNo,sL_ProductID : String;
    sL_Operator, sL_Now, sL_HighLevelCmdID : String;
    aUTCTime: TDateTime;
begin
    if not bG_Login then
    begin
      edtAlias.SetFocus;
      MessageDlg(getLanguageSettingInfo('Sim_CA_Msg_6_Content'), mtWarning, [mbOK],0);
      Exit;
    end;

    bL_CheckTimes := false;
    bL_CouldBeExecuted := true;
    sL_StbNo := Trim(edtStbNo.Text);
    sL_IccNo := Trim(edtIccNo.Text);
    sL_Operator := sG_UserID;
    sL_CmdStatus := 'W';
    sL_CompCode := '20';
    sL_SeqNo := 'S_SENDNAGRA_SEQNO.NEXTVAL';

    case nI_ActionType of
      1://開機
       begin
         sL_Action :=getLanguageSettingInfo('Sim_CA_Action_2_Content');
         sL_AuthorStartDate := 'NULL';
         sL_AuthorStopDate :=  'NULL';
         sL_ProductID := '';
         sL_HighLevelCmdID := 'A1';
         sL_Notes := '';
         sL_MisIrdCmdData := '';
         sL_MisIrdCmdID := '';
         sL_PinCode := '';
       end;
      2://關機
       begin
         sL_Action :=getLanguageSettingInfo('Sim_CA_Action_3_Content');
         sL_AuthorStartDate := 'NULL';
         sL_AuthorStopDate :=  'NULL';
         sL_ProductID := '';
         sL_HighLevelCmdID := 'A2';
         sL_Notes := '';
         sL_MisIrdCmdData := '';
         sL_MisIrdCmdID := '';
         sL_PinCode := '';
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

           if InputQuery(getLanguageSettingInfo('Sim_CA_Msg_7_Content'), getLanguageSettingInfo('Sim_CA_Msg_8_Content'),sL_ChannelDays) then
           begin
             nL_ChannelDays := StrToIntDef(sL_ChannelDays,-1);
             if nL_ChannelDays = -1 then
             begin
               bL_Exit := true;
               MessageDlg(getLanguageSettingInfo('Sim_CA_Msg_9_Content'), mtWarning, [mbOK],0);
             end;
           end
           else
             bL_Exit := true;

         end;

         bL_Exit := False;
         if bL_Exit then
         begin
           Exit;
         end;


         sL_AuthorStartDate := getOracleSQLDateStr(now);

         nL_ChannelDays := DEFAULT_CHANNEL_DAYS;

         sL_AuthorStopDate := getOracleSQLDateStr(now+nL_ChannelDays);
         sL_Action := getLanguageSettingInfo('Sim_CA_Action_1_Content');
//         sL_ProductID := '全部';
         sL_ProductID := sG_ProductDescForOpenCh;

         sL_HighLevelCmdID := 'B1';
         sL_Notes := sG_ProductIdForOpenCh;
         sL_MisIrdCmdData := '';
         sL_MisIrdCmdID := '';
         sL_PinCode := '';
       end;
      4://關頻道
       begin
         sL_AuthorStartDate := 'NULL';
         sL_AuthorStopDate :=  'NULL';
         sL_Action :=getLanguageSettingInfo('Sim_CA_Action_4_Content');
//         sL_ProductID := '全部';
         sL_ProductID := sG_ProductDescForCloseCh;
         sL_HighLevelCmdID := 'B2';
         sL_Notes := sG_ProductIdForCloseCh;
         sL_MisIrdCmdData := '';
         sL_MisIrdCmdID := '';
         sL_PinCode := '';         
       end;
      5://設定親子密碼
       begin
         sL_AuthorStartDate := 'NULL';
         sL_AuthorStopDate :=  'NULL';
         sL_Action :=getLanguageSettingInfo('Sim_CA_Action_5_Content');
         sL_ProductID := '';
         sL_HighLevelCmdID := 'E1';
         sL_Notes := '';
         sL_MisIrdCmdID := '1';
         if sG_NewPinCode='' then
           sL_PinCode := DEFAULT_PIN_CODE
         else
           sL_PinCode := sG_NewPinCode;
         sL_MisIrdCmdData := sL_PinCode;

       end;

    end;

    try
      if bL_CheckTimes then
      begin
        sL_SQL := ' select count(*) ACTIONTIMES from CaWareHouseLog where STB_NO=' + STR_SEP + sL_StbNo + STR_SEP ;
        sL_SQL := sL_SQL + ' and ACTION=' + STR_SEP + getLanguageSettingInfo('Sim_CA_Action_1_Content') + STR_SEP;
        qryActionTimes.SQL.Clear;
        qryActionTimes.SQL.Add(sL_SQL);
        qryActionTimes.Open;
        nL_Times := qryActionTimes.FieldByName('ACTIONTIMES').AsInteger;

        if nL_Times >= MAX_ACTION_TIMES then
        begin
          sL_AlermInfo := getLanguageSettingInfo('Sim_CA_Msg_10_Content') + IntToStr(nL_Times)+ getLanguageSettingInfo('Sim_CA_Msg_11_Content');
          bL_CouldBeExecuted := false;
        end
        else
          bL_CouldBeExecuted := true;
        qryActionTimes.Close;
      end;

      if bL_CouldBeExecuted then
      begin
        sL_LogSQL := 'insert into CaWareHouseLog(ICC_NO,STB_NO,ProductID,ACTION,AUTHORSTARTDATE,AUTHORSTOPDATE,OPERATOR,UPDTIME,PINCODE)';
        sL_LogSQL := sL_LogSQL +' values(' + STR_SEP + sL_IccNo + STR_SEP + ',' + STR_SEP + sL_StbNo + STR_SEP + ',';
        sL_LogSQL := sL_LogSQL + STR_SEP + sL_ProductID + STR_SEP + ',' +  STR_SEP + sL_Action + STR_SEP + ',';
        sL_LogSQL := sL_LogSQL + sL_AuthorStartDate + ',' + sL_AuthorStopDate + ',' ;
        sL_LogSQL := sL_LogSQL + STR_SEP + sL_Operator +  STR_SEP + ',' + getOracleSQLDateTimeStr(now) + ',' + STR_SEP + sL_PinCode + STR_SEP +')';


        sL_SQL := 'insert into send_nagra(HIGH_LEVEL_CMD_ID,ICC_NO,STB_NO,SUBSCRIPTION_BEGIN_DATE,SUBSCRIPTION_END_DATE,NOTES,CMD_STATUS,OPERATOR,SEQNO,COMPCODE,MIS_IRD_CMD_ID,MIS_IRD_CMD_DATA)';
        sL_SQL := sL_SQL + ' values(' + STR_SEP + sL_HighLevelCmdID + STR_SEP + ',' + STR_SEP + sL_IccNo + STR_SEP +',';
        sL_SQL := sL_SQL + STR_SEP + sL_StbNo + STR_SEP + ',' +  sL_AuthorStartDate + ',' +  sL_AuthorStopDate + ',';
        sL_SQL := sL_SQL + STR_SEP + sL_Notes + STR_SEP + ',' + STR_SEP + sL_CmdStatus + STR_SEP + ',' + STR_SEP + sL_Operator + STR_SEP +',' + sL_SeqNo + ',';
        sL_SQL := sL_SQL + STR_SEP + sL_CompCode + STR_SEP + ',' + STR_SEP + sL_MisIrdCmdID + STR_SEP ;
        sL_SQL := sL_SQL + ',' + STR_SEP + sL_MisIrdCmdData + STR_SEP + ')';

        ADOCommand1.CommandText := sL_LogSQL;
        ADOCommand1.Execute;

        ADOCommand1.CommandText := sL_SQL;
        ADOCommand1.Execute;

        MessageDlg(getLanguageSettingInfo('Sim_CA_Msg_12_Content'), mtInformation, [mbOK],0);
      end
      else
        MessageDlg(sL_AlermInfo, mtWarning, [mbOK],0);
    except
      MessageDlg(getLanguageSettingInfo('Sim_CA_Msg_13_Content'), mtError, [mbOK],0);
    end;

    edtStbNo.SetFocus;
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
    sG_DbUserID := 'com';
    sG_DbUserPasswd := 'com';    
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
      if InputQuery(getLanguageSettingInfo('Sim_CA_Msg_14_Content'), 'ID:',sL_DbUserID) then
      begin
        sG_DbUserID := sL_DbUserID;
        bL_Change1 := true;
      end;

      if InputQuery(getLanguageSettingInfo('Sim_CA_Msg_15_Content'), 'Password:',sL_DbUSerPasswd) then
      begin
        sG_DbUserPasswd := sL_DbUSerPasswd;
        bL_Change2 := true;
      end;
      
      if bL_Change1 or bL_Change2 then
      begin
        btnLogin.SetFocus;
        MessageDlg(getLanguageSettingInfo('Sim_CA_Msg_16_Content'), mtInformation, [mbOK],0);
      end;
    end;
end;


procedure TfrmMain.btnChooseChClick(Sender: TObject);
var
    sL_ChannelID, sL_Desc : String;
    ii : Integer;
    sL_SQL, sL_TargetSQL, sL_Tmp : String;
    L_TempStrList, L_TargetFieldNamesStrList, L_TargetFieldValueStrList : TStringList;
begin
    if not ADOConnection1.Connected then
    begin
      edtAlias.SetFocus;
      MessageDlg(getLanguageSettingInfo('Sim_CA_Msg_6_Content'), mtWarning, [mbOK],0);
      Exit;
    end;

    with qryChInfo do
    begin
      sL_SQL := 'select DESCRIPTION, CHANNELID from cd024 ';
      Close;
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;
      FieldByName('DESCRIPTION').DisplayLabel := getLanguageSettingInfo('Sim_CA_Msg_17_Content');
    end;
    if qryChInfo.RecordCount = 0 then
    begin
      qryChInfo.Close;
      MessageDlg(getLanguageSettingInfo('Sim_CA_Msg_18_Content'), mtWarning, [mbOK],0);
      Exit;
    end
    else
    begin
      L_TargetFieldNamesStrList := TStringList.Create;
      L_TargetFieldValueStrList := TStringList.Create;

      L_TargetFieldNamesStrList.Add('DESCRIPTION');
      L_TargetFieldNamesStrList.Add('CHANNELID');
      if SelectMultiRecords(getLanguageSettingInfo('Sim_CA_Msg_19_Content'), qryChInfo, 'DESCRIPTION', ',', L_TargetFieldNamesStrList, L_TargetFieldValueStrList,true,sL_TargetSQL) = mrOk then
      begin
        memCh.Clear;
        sG_ProductDescForOpenCh := '';
        sG_ProductDescForCloseCh := '';
        sG_ProductIDForOpenCh := '';
        sG_ProductIDForCloseCh := '';
                                
        for ii:=0 to L_TargetFieldValueStrList.Count -1 do
        begin
          sL_Tmp := Trim(L_TargetFieldValueStrList.Strings[ii]);
          L_TempStrList := TUstr.ParseStrings(sL_Tmp,',');

          sL_Desc := L_TempStrList.Strings[0];
          sL_ChannelID := L_TempStrList.Strings[1];
          memCh.Lines.Add(sL_Desc);
          if sG_ProductDescForOpenCh='' then
            sG_ProductDescForOpenCh := sL_Desc
          else
            sG_ProductDescForOpenCh := sG_ProductDescForOpenCh + ',' + sL_Desc;

          if sG_ProductDescForCloseCh='' then
            sG_ProductDescForCloseCh := sL_Desc
          else
            sG_ProductDescForCloseCh := sG_ProductDescForCloseCh + ',' + sL_Desc;

          if sG_ProductIdForOpenCh='' then
            sG_ProductIdForOpenCh := 'A~' + sL_ChannelID
          else
            sG_ProductIdForOpenCh := sG_ProductIdForOpenCh + ',' + 'A~' + sL_ChannelID;
                        
          if sG_ProductIdForCloseCh='' then
            sG_ProductIdForCloseCh := sL_ChannelID
          else
            sG_ProductIdForCloseCh := sG_ProductIdForCloseCh + ',' + sL_ChannelID;          


        end;

      end;
      L_TargetFieldNamesStrList.Free;
      L_TargetFieldValueStrList.Free;

    end;


end;

procedure TfrmMain.mitFun1Click(Sender: TObject);
var
    L_Frm : TfrmReport1;
begin
    if not ADOConnection1.Connected then
    begin
      edtAlias.SetFocus;
      MessageDlg(getLanguageSettingInfo('Sim_CA_Msg_6_Content'), mtWarning, [mbOK],0);
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

procedure TfrmMain.btnPinCodeClick(Sender: TObject);
var
    sL_IccNo : String;
begin
    sL_IccNo := Trim(edtIccNo.Text);
    if (sL_IccNo='') then
    begin
      MessageDlg(getLanguageSettingInfo('Sim_CA_Msg_1_Content'), mtWarning, [mbOK],0);
      edtIccNo.SetFocus;
      Exit;
    end
    else
    begin
      sG_NewPinCode := DEFAULT_PIN_CODE;
      if InputQuery(getLanguageSettingInfo('Sim_CA_Msg_20_Content'), 'Pin Code:',sG_NewPinCode) then
      begin
        if (length(sG_NewPinCode)<>length(DEFAULT_PIN_CODE)) and (length(sG_NewPinCode)<>0) then
          MessageDlg(getLanguageSettingInfo('Sim_CA_Msg_21_Content') + IntToStr(length(DEFAULT_PIN_CODE)) + getLanguageSettingInfo('Sim_CA_Msg_22_Content'), mtWarning,[mbOK],0)
        else
          RealAction(5);
      end;
    end;
end;

procedure TfrmMain.mitFun2Click(Sender: TObject);
var
    L_Frm : TfrmSysParam;
begin
    L_Frm := TfrmSysParam.Create(Application);
    L_Frm.ShowModal;
    L_Frm.Free;
end;


procedure TfrmMain.initLanguageIniFile;
var
    sL_ExeFileName, sL_ExePath, sL_IniFileName: String;
begin
    sL_ExeFileName := Application.ExeName;
    sL_ExePath := ExtractFileDir(sL_ExeFileName);
    sL_IniFileName := sL_ExePath + '\' + LANGUAGE_SETTING_INI_FILE_NAME;
    if FileExists(sL_IniFileName) then
      G_LanguageSettingIni := TIniFile.Create(sL_IniFileName);




end;

procedure TfrmMain.setLanguageString;
begin
    Caption := getLanguageSettingInfo('frmSimCA_1_Caption');
    mitFun1Header.Caption := getLanguageSettingInfo('Sim_CA_mitFun1Header_Caption');
    mitFun1.Caption := getLanguageSettingInfo('Sim_CA_mitFun1_Caption');
    mitFun2.Caption := getLanguageSettingInfo('Sim_CA_mitFun2_Caption');


    Lab_1.Caption := getLanguageSettingInfo('Sim_CA_Lab_1_Caption');
    Lab_2.Caption := getLanguageSettingInfo('Sim_CA_Lab_2_Caption');
    Lab_3.Caption := getLanguageSettingInfo('Sim_CA_Lab_3_Caption');
    Lab_4.Caption := getLanguageSettingInfo('Sim_CA_Lab_4_Caption');


    btnLogin.Caption := getLanguageSettingInfo('Sim_CA_btnLogin_Caption');
    btnChooseCh.Caption := getLanguageSettingInfo('Sim_CA_btnChooseCh_Caption');
    btnPair.Caption := getLanguageSettingInfo('Sim_CA_btnPair_Caption');
    btnDepair.Caption := getLanguageSettingInfo('Sim_CA_btnDepair_Caption');
    btnPinCode.Caption := getLanguageSettingInfo('Sim_CA_btnPinCode_Caption');
    btnOpenCh.Caption := getLanguageSettingInfo('Sim_CA_btnOpenCh_Caption');
    btnCloseCh.Caption := getLanguageSettingInfo('Sim_CA_btnCloseCh_Caption');
    btnExit.Caption := getLanguageSettingInfo('Sim_CA_btnExit_Caption');



end;

function TfrmMain.getLanguageSettingInfo(sI_Key: String): String;
var
    sL_Result, sL_LanguageType : String;
begin
//    if assigned()
    if Assigned(G_LanguageSettingIni) then
    begin
      sL_LanguageType := G_LanguageSettingIni.ReadString('CURRENT_LANGUAGE_TYPE','CURRENT_LANGUAGE_TYPE','TC');// default is Traditional Chinese

      sL_Result := G_LanguageSettingIni.ReadString(sL_LanguageType,sI_Key,'');
    end
    else
      sL_Result := sI_Key;
    result := sL_Result;
end;


end.
