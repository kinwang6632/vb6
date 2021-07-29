unit frmMainU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Buttons, DB, ADODB, Menus, xmldom, XMLIntf,
  msxmldom, XMLDoc, IdBaseComponent, IdComponent, IdTCPConnection,
  IdTCPClient, Psock ;
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
    StaticText5: TStaticText;
    StaticText6: TStaticText;
    edtUserID: TEdit;
    edtPasswd: TEdit;
    btnLogin: TBitBtn;
    memCh: TMemo;
    btnChooseCh: TBitBtn;
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    BitBtn7: TBitBtn;
    N3: TMenuItem;
    chbShowResult: TCheckBox;
    Panel3: TPanel;
    N4: TMenuItem;
    StaticText7: TStaticText;
    StaticText8: TStaticText;
    edtRegionKey: TEdit;
    edtBouquetID: TEdit;
    IdTCPClient1: TIdTCPClient;
    Timer1: TTimer;
    BitBtn8: TBitBtn;
    StaticText4: TStaticText;
    edtMSubscriberID: TEdit;
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure BitBtn4Click(Sender: TObject);
    procedure BitBtn5Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnLoginClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure btnChooseChClick(Sender: TObject);
    procedure N2Click(Sender: TObject);
    procedure BitBtn7Click(Sender: TObject);
    procedure N3Click(Sender: TObject);
    procedure N4Click(Sender: TObject);
    procedure edtBouquetIDKeyPress(Sender: TObject; var Key: Char);
    procedure Timer1Timer(Sender: TObject);
    procedure BitBtn8Click(Sender: TObject);
    procedure edtSubscriberIDKeyPress(Sender: TObject; var Key: Char);
    procedure edtSubscriberIDExit(Sender: TObject);
    procedure edtBouquetIDExit(Sender: TObject);
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
    nG_SeqNo : Integer;

    sG_SubscriberID,sG_IccNo,sG_HighLevelCmdID,sG_Action : String;
    sG_ExpirationDate,sG_UpdTime,sG_Operator : String;

    procedure RealAction(nI_ActionType:Integer);
    procedure setCursor(nI_State:Integer);
    function getRegionKey:String;
    function getBouquetID:String;
    procedure handleResponse(sI_ResponseXML:String);
  public
    { Public declarations }
    sG_UserID : String;
    function activeReport1Data(sI_SQL:String):Integer;
    function getOracleSQLDateStr(dI_Date:TDate):String;
    function getOracleSQLDateTimeStr(dI_Date:TDateTime):String;
    function setInfoXml(sI_CompName,sI_Content : String) : WideString;
    procedure parseInfoXml(sI_Msg : String;var sI_CompName,sI_Content : String);
    procedure canAction(bI_CanAction : Boolean;sI_ReturnString : String);
  end;

var
  frmMain: TfrmMain;

implementation

uses frmDbMultiSelectu, Ustru, frmReport1U, frmParamU, UdateTimeu,
  frmCmdResultU, ConstU, frmSyncU, dtmMainU, xmlU;

{$R *.dfm}

procedure TfrmMain.BitBtn1Click(Sender: TObject);
var
    sL_Info : String;
begin
    if MessageDlg('是否確定要結束程式?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      if IdTCPClient1.Connected then
      begin
        sL_Info := setInfoXml(dtmMain.sG_ControlCompName,SIMPLE_CA_DISCONNECT);
        IdTCPClient1.WriteLn(sL_Info);
      end;

      if IdTCPClient1.Connected then
        IdTCPClient1.Disconnect;
      Application.Terminate;
    end;  
end;

procedure TfrmMain.BitBtn2Click(Sender: TObject);
var
    sL_RegionKey, sL_BouquetID : String;
    sL_IccNo : String;
begin

    sL_IccNo := Trim(edtIccNo.Text);
    sL_RegionKey := Trim(edtRegionKey.Text);
    sL_BouquetID := Trim(edtBouquetID.Text);

    if (sL_IccNo='') then
    begin
      MessageDlg('請輸入ICC No.', mtWarning, [mbOK],0);
      edtIccNo.SetFocus;
      Exit;
    end
    else if (sL_RegionKey='') then
    begin
      MessageDlg('請輸入Region Key.', mtWarning, [mbOK],0);
      edtRegionKey.SetFocus;
      Exit;
    end
    else if (sL_BouquetID='') then
    begin
      MessageDlg('請輸入Bouquet ID.', mtWarning, [mbOK],0);
      edtBouquetID.SetFocus;
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
    nG_SeqNo := 1;
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
    if dtmMain.cdsCD024.Active then
      dtmMain.cdsCD024.Close;
      
    G_UserNameStrList.Free;
    G_UserPasswdStrList.Free;
    G_ServiceInfo.Free;
end;

procedure TfrmMain.btnLoginClick(Sender: TObject);
var
    sL_DbUserID, sL_DbUserPasswd ,sL_Info: String;
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
      bG_Login := bL_Login;
      sG_UserID := sL_UserID;
      try
        IdTCPClient1.Host := frmParam.getCaIP;
        IdTCPClient1.Port := StrToIntDef(frmParam.getCaPort,-1);
        IdTCPClient1.Connect;

        if IdTCPClient1.Connected then
        begin
          canAction(true,'已連接上CA G/W');
          sL_Info := setInfoXml(dtmMain.sG_ControlCompName,SIMPLE_CA_CONNECTION_OK);
          IdTCPClient1.WriteLn(sL_Info);
        end;

      except
        MessageDlg('無法連接上CA!請檢查IP與Port之設定', mtError, [mbOK],0);
        Exit;
      end;

      MessageDlg('登入完成', mtInformation, [mbOK],0);
    end;
end;

procedure TfrmMain.RealAction(nI_ActionType: Integer);
var
    sL_XML : String;
    L_TmpStrList1, L_TmpStrList2 : TStringList;
    sL_ChannelDays, sL_MisIrdCmdData,sL_MisIrdCmdID, sL_PinCode : String;
    ii,nL_ChannelDays, nL_Times ,nL_Len : Integer;
    bL_CouldBeExecuted,bL_Exit : boolean;
    sL_CompName, sL_CompCode, sL_SeqNo, sL_Notes, sL_CmdStatus : String;
    sL_ExpirationDate : String;
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

    L_RootNode : IXmlNode;
    L_XMLDocument : TXMLDocument;
    sL_UpdTime, sL_Result : String;
begin
    if not bG_Login then
    begin
      edtUserID.SetFocus;
      MessageDlg('請先登入系統!', mtWarning, [mbOK],0);
      Exit;
    end;
    bL_Exit := false;

    bL_CouldBeExecuted := true;
    sL_SubscriberID := Trim(edtSubscriberID.Text);
    sL_IccNo := Trim(edtIccNo.Text);
    sL_Operator := sG_UserID;
    sL_UpdTime := TUdateTime.GetPureDateTimeStr(now,'/',':');
    sL_CmdStatus := 'W';
    sL_CompCode := dtmMain.sG_ControlCompCode;
    sL_CompName := dtmMain.sG_ControlCompName;


    sL_SeqNo := IntToStr(nG_SeqNo);
    setCursor(1);
    case nI_ActionType of
      1://開機
       begin
         sL_Action :='開機';
         sL_ExpirationDate :=  'NULL';
         sL_ProductID := '';
         sL_HighLevelCmdID := 'A1';
         sL_Notes := getBouquetID ;
         sL_MisIrdCmdData := '';
         sL_MisIrdCmdID := '';
         sL_PinCode := '';
         sL_ReportBackAvailability := frmParam.getReportbackAvailability;
         sL_ReportBackDate := frmParam.getReportbackDate;
         sL_RegionKey := getRegionKey;
         sL_PopulationId := frmParam.getPopulationID;

         //down, 產生 XML 之data...
         L_XMLDocument := TXMLDocument.Create(Application);
         L_XMLDocument.Active := true;
         L_XMLDocument.Encoding := 'Big5';
         L_RootNode := TMyXML.createRootNode(L_XMLDocument, ROOT_NODE_COMMAND_NAME, ROOT_NODE_COMMAND_VALUE, sL_Result);
         TMyXML.addChild(L_RootNode,'SEQNO', sL_SeqNo);
         TMyXML.addChild(L_RootNode,'COMPCODE', sL_CompCode);
         TMyXML.addChild(L_RootNode,'COMPNAME', sL_CompName);
         TMyXML.addChild(L_RootNode,'COMMANDID', sL_HighLevelCmdID);

         TMyXML.addChild(L_RootNode,'SUBSCRIBERID', '');
         TMyXML.addChild(L_RootNode,'ICCNO', sL_IccNo);
         TMyXML.addChild(L_RootNode,'PINCODE', '');
         TMyXML.addChild(L_RootNode,'POPULATIONID', sL_PopulationId);
         TMyXML.addChild(L_RootNode,'REGIONKEY', sL_RegionKey);

         TMyXML.addChild(L_RootNode,'REPORTBACKAVAILABILITY', sL_ReportBackAvailability);
         TMyXML.addChild(L_RootNode,'REPORTBACKDATE', sL_ReportBackDate);
         TMyXML.addChild(L_RootNode,'NOTES', sL_Notes);
         TMyXML.addChild(L_RootNode,'OPERATOR', sL_Operator);
         TMyXML.addChild(L_RootNode,'UPDTIME', sL_UpdTime);
         TMyXML.addChild(L_RootNode,'PINCONTROL', '');
         TMyXML.addChild(L_RootNode,'SERVICEID', '');

         //up, 產生 XML 之data...

         for ii:=0 to L_XMLDocument.XML.Count -1 do
           sL_XML := sL_XML + L_XMLDocument.XML[ii];
         L_XMLDocument.XML.SaveToFile('c:\A1.xml');
         L_XMLDocument.Free;

       end;
      2://關機
       begin
         sL_Action :='關機';
         sL_ExpirationDate :=  'NULL';
         sL_ProductID := '';
         sL_HighLevelCmdID := 'A2';
         sL_Notes := '';
         sL_MisIrdCmdData := '';
         sL_MisIrdCmdID := '';
         sL_PinCode := '';

         //down, 產生 XML 之data...
         L_XMLDocument := TXMLDocument.Create(Application);
         L_XMLDocument.Active := true;
         L_XMLDocument.Encoding := 'Big5';
         L_RootNode := TMyXML.createRootNode(L_XMLDocument, ROOT_NODE_COMMAND_NAME, ROOT_NODE_COMMAND_VALUE, sL_Result);
         TMyXML.addChild(L_RootNode,'SEQNO', sL_SeqNo);
         TMyXML.addChild(L_RootNode,'COMPCODE', sL_CompCode);
         TMyXML.addChild(L_RootNode,'COMPNAME', sL_CompName);
         TMyXML.addChild(L_RootNode,'COMMANDID', sL_HighLevelCmdID);

         TMyXML.addChild(L_RootNode,'SUBSCRIBERID', sL_SubscriberID);
         TMyXML.addChild(L_RootNode,'OPERATOR', sL_Operator);
         TMyXML.addChild(L_RootNode,'UPDTIME', sL_UpdTime);

         //up, 產生 XML 之data...

         for ii:=0 to L_XMLDocument.XML.Count -1 do
           sL_XML := sL_XML + L_XMLDocument.XML[ii];
//         L_XMLDocument.XML.SaveToFile('c:\aaa.xml');
         L_XMLDocument.Free;


       end;
      3://開頻道
       begin
         sL_ChannelDays := IntToStr(DEFAULT_CHANNEL_DAYS);
         if InputQuery('請輸入頻道收視截止日', '收視截止日[Ex:2003/01/31]:',sL_ExpirationDate) then
         begin
           nL_Len := Length(Trim(sL_ExpirationDate));
           if sL_ExpirationDate<>'' then
           begin
             if (nL_Len <> 10) or ((nL_Len = 10) and (not TUdateTime.IsDateStr(sL_ExpirationDate,'/'))) then
             begin
               bL_Exit := true;
               MessageDlg('收視截止日之格式錯誤!', mtError, [mbOK],0);
             end;
           end
           else if sL_ExpirationDate='' then
           begin
             sL_ExpirationDate := TUstr.AddString(sL_ExpirationDate,'0',true,8);
           end;
         end
         else
           bL_Exit := true;

         if bL_Exit then
         begin
           setCursor(2);
           Exit;
         end;
         sL_ExpirationDate := TUstr.replaceStr(sL_ExpirationDate,'/','');
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
             sL_Notes := sL_ServiceID + '~' + sL_ExpirationDate + '~' + 'N'
           else
             sL_Notes := sL_Notes + ';' + sL_ServiceID + '~' + sL_ExpirationDate + '~' + 'N';
         end;

         //down, 產生 XML 之data...
         L_XMLDocument := TXMLDocument.Create(Application);
         L_XMLDocument.Active := true;
         L_XMLDocument.Encoding := 'Big5';
         L_RootNode := TMyXML.createRootNode(L_XMLDocument, ROOT_NODE_COMMAND_NAME, ROOT_NODE_COMMAND_VALUE, sL_Result);
         TMyXML.addChild(L_RootNode,'SEQNO', sL_SeqNo);
         TMyXML.addChild(L_RootNode,'COMPCODE', sL_CompCode);
         TMyXML.addChild(L_RootNode,'COMPNAME', sL_CompName);
         TMyXML.addChild(L_RootNode,'COMMANDID', sL_HighLevelCmdID);
         TMyXML.addChild(L_RootNode,'OPERATOR', sL_Operator);
         TMyXML.addChild(L_RootNode,'UPDTIME', sL_UpdTime);

         TMyXML.addChild(L_RootNode,'SUBSCRIBERID', sL_SubscriberID);


         TMyXML.addChild(L_RootNode,'NOTES', sL_Notes);

         //up, 產生 XML 之data...
         for ii:=0 to L_XMLDocument.XML.Count -1 do
           sL_XML := sL_XML + L_XMLDocument.XML[ii];
//         L_XMLDocument.XML.SaveToFile('c:\aaa.xml');
         L_XMLDocument.Free;

       end;
      4://關頻道
       begin
         sL_ExpirationDate :=  'NULL';
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
             sL_Notes := sL_Notes + '~' + sL_ServiceID;
         end;

         //down, 產生 XML 之data...
         L_XMLDocument := TXMLDocument.Create(Application);
         L_XMLDocument.Active := true;
         L_XMLDocument.Encoding := 'Big5';
         L_RootNode := TMyXML.createRootNode(L_XMLDocument, ROOT_NODE_COMMAND_NAME, ROOT_NODE_COMMAND_VALUE, sL_Result);
         TMyXML.addChild(L_RootNode,'SEQNO', sL_SeqNo);
         TMyXML.addChild(L_RootNode,'COMPCODE', sL_CompCode);
         TMyXML.addChild(L_RootNode,'COMPNAME', sL_CompName);
         TMyXML.addChild(L_RootNode,'COMMANDID', sL_HighLevelCmdID);
         TMyXML.addChild(L_RootNode,'OPERATOR', sL_Operator);
         TMyXML.addChild(L_RootNode,'UPDTIME', sL_UpdTime);

         TMyXML.addChild(L_RootNode,'SUBSCRIBERID', sL_SubscriberID);


         TMyXML.addChild(L_RootNode,'NOTES', sL_Notes);

         //up, 產生 XML 之data...

         for ii:=0 to L_XMLDocument.XML.Count -1 do
           sL_XML := sL_XML + L_XMLDocument.XML[ii];
//         L_XMLDocument.XML.SaveToFile('c:\aaa.xml');
         L_XMLDocument.Free;



       end;
      6: //設定 Bouquet ID
       begin
         sL_ExpirationDate :=  'NULL';
         sL_Action :='設定BouquetID';
         sL_ProductID := '';
         sL_HighLevelCmdID := 'G3';
         sL_MisIrdCmdData := '';
         sL_MisIrdCmdID := '';
         sL_PinCode := '';
         sL_Notes := edtBouquetID.Text;

         //down, 產生 XML 之data...
         L_XMLDocument := TXMLDocument.Create(Application);
         L_XMLDocument.Active := true;
         L_XMLDocument.Encoding := 'Big5';
         L_RootNode := TMyXML.createRootNode(L_XMLDocument, ROOT_NODE_COMMAND_NAME, ROOT_NODE_COMMAND_VALUE, sL_Result);
         TMyXML.addChild(L_RootNode,'SEQNO', sL_SeqNo);
         TMyXML.addChild(L_RootNode,'COMPCODE', sL_CompCode);
         TMyXML.addChild(L_RootNode,'COMPNAME', sL_CompName);
         TMyXML.addChild(L_RootNode,'COMMANDID', sL_HighLevelCmdID);
         TMyXML.addChild(L_RootNode,'OPERATOR', sL_Operator);
         TMyXML.addChild(L_RootNode,'UPDTIME', sL_UpdTime);

         TMyXML.addChild(L_RootNode,'SUBSCRIBERID', sL_SubscriberID);


         TMyXML.addChild(L_RootNode,'NOTES', sL_Notes);

         //up, 產生 XML 之data...

         for ii:=0 to L_XMLDocument.XML.Count -1 do
           sL_XML := sL_XML + L_XMLDocument.XML[ii];
//         L_XMLDocument.XML.SaveToFile('c:\G3.xml');
         L_XMLDocument.Free;

       end;
      5://設定親子密碼
       begin
         sL_ExpirationDate :=  'NULL';
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

         //down, 產生 XML 之data...
         L_XMLDocument := TXMLDocument.Create(Application);
         L_XMLDocument.Active := true;
         L_XMLDocument.Encoding := 'Big5';
         L_RootNode := TMyXML.createRootNode(L_XMLDocument, ROOT_NODE_COMMAND_NAME, ROOT_NODE_COMMAND_VALUE, sL_Result);
         TMyXML.addChild(L_RootNode,'SEQNO', sL_SeqNo);
         TMyXML.addChild(L_RootNode,'COMPCODE', sL_CompCode);
         TMyXML.addChild(L_RootNode,'COMPNAME', sL_CompName);
         TMyXML.addChild(L_RootNode,'COMMANDID', sL_HighLevelCmdID);

         TMyXML.addChild(L_RootNode,'SUBSCRIBERID', sL_SubscriberID);
         TMyXML.addChild(L_RootNode,'PINCODE', sL_PinCode);
         TMyXML.addChild(L_RootNode,'OPERATOR', sL_Operator);
         TMyXML.addChild(L_RootNode,'UPDTIME', sL_UpdTime);
         TMyXML.addChild(L_RootNode,'PINCONTROL', sL_PinControl);

         //up, 產生 XML 之data...

         for ii:=0 to L_XMLDocument.XML.Count -1 do
           sL_XML := sL_XML + L_XMLDocument.XML[ii];

         L_XMLDocument.Free;
       end;
    end;

    try
      if bL_CouldBeExecuted then
      begin
        if IdTCPClient1.Connected then
        begin
          //設定不可下其他Command
          canAction(false,'傳送指令中...');

          IdTCPClient1.WriteLn(sL_XML);//將CA 指令傳給 Gateway

          //Log一筆資料到SimpleCA.mdb
          //dtmMain.insertDataToLog(sL_SubscriberID,sL_IccNo,sL_HighLevelCmdID,sL_Action,sL_ExpirationDate,sL_UpdTime,sL_Operator);
          sG_SubscriberID := sL_SubscriberID;
          sG_IccNo := sL_IccNo;
          sG_HighLevelCmdID := sL_HighLevelCmdID;
          sG_Action := sL_Action;
          sG_ExpirationDate := sL_ExpirationDate;
          sG_UpdTime := sL_UpdTime;
          sG_Operator := sL_Operator;
        end;
      end
      else
        MessageDlg(sL_AlermInfo, mtWarning, [mbOK],0);
    except
      MessageDlg(sL_Action + ' 錯誤!', mtError, [mbOK],0);
    end;
    setCursor(2);
    edtSubscriberID.SetFocus;
    Inc(nG_SeqNo);
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
    edtUserID.SetFocus;
    sG_DbUserID := 'system';
    sG_DbUserPasswd := 'manager';

    canAction(false,'');
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
        btnLogin.SetFocus;
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
    if not FileExists(CHANNEL_DATA_FILENAME) then
    begin
      MessageDlg('沒有頻道資料!!請先執行資料同步功能.', mtWarning, [mbOK],0);
      Exit;
    end
    else
    begin
      if not dtmMain.cdsCD024.Active then
        dtmMain.cdsCD024.LoadFromFile(CHANNEL_DATA_FILENAME);
      dtmMain.cdsCD024.FieldByName('DESCRIPTION').DisplayLabel := '頻道名稱';
      
      L_TargetFieldNamesStrList := TStringList.Create;
      L_TargetFieldValueStrList := TStringList.Create;

      L_TargetFieldNamesStrList.Add('DESCRIPTION');
      L_TargetFieldNamesStrList.Add('CHANNELID');
      if SelectMultiRecords('請點選公司別', dtmMain.cdsCD024, 'DESCRIPTION', ',', L_TargetFieldNamesStrList, L_TargetFieldValueStrList,true,sL_TargetSQL) = mrOk then
      begin
        memCh.Clear;
        sG_ProductDescForOpenCh := '';
        sG_ProductDescForCloseCh := '';
        G_ServiceInfo.Clear;


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
{
    if not bG_Login then
    begin
      edtUserID.SetFocus;
      MessageDlg('請先登入系統!', mtWarning, [mbOK],0);
      Exit;
    end
    else
    begin
      L_Frm := TfrmReport1.Create(Application);
      L_Frm.ShowModal;
      L_Frm.Free;
    end;
}
    L_Frm := TfrmReport1.Create(Application);
    L_Frm.ShowModal;
    L_Frm.Free;
end;

function TfrmMain.activeReport1Data(sI_SQL: String): Integer;
var
    nL_RecordCount : Integer;
begin
    with dtmMain.qryReport1 do
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


procedure TfrmMain.setCursor(nI_State: Integer);
begin
    case nI_State of
      1:
        Screen.Cursor := crHourGlass ;
      2:
        Screen.Cursor := crDefault ;      
    end;
end;

procedure TfrmMain.N4Click(Sender: TObject);
var
    L_Frm : TfrmSync; 
begin
    try
      L_Frm := TfrmSync.Create(Application);
      L_Frm.ShowModal;
      L_Frm.Free;

    except

    end;      
end;

function TfrmMain.getRegionKey: String;
begin
    result := Trim(edtRegionKey.Text);
end;

function TfrmMain.getBouquetID: String;
begin
    result := Trim(edtBouquetID.Text);
end;

procedure TfrmMain.edtBouquetIDKeyPress(Sender: TObject; var Key: Char);
begin
    if Key in['a'..'z'] then
      Dec(Key,32);  //強迫轉大寫

end;

procedure TfrmMain.Timer1Timer(Sender: TObject);
var
    sL_Msg,sL_CompName,sL_Content : string;
begin
  if not IdTcpClient1.Connected then
    exit;

  sL_Msg := IdTCPClient1.ReadLn('', StrToInt(frmParam.getCaTimeout));
  if (sL_Msg<>'') AND (AnsiPos('<' + ROOT_NODE_RESPONSE_NAME + '>', sL_Msg)>0) then
    handleResponse(sL_Msg)
  else if AnsiPos('<' + ROOT_NODE_INFO_NAME + '>', sL_Msg)>0 then
  begin
    parseInfoXml(sL_Msg,sL_CompName,sL_Content);
    if sL_Content=SIMPLE_CA_DISCONNECT then
    begin
      if IdTCPClient1.Connected then
        IdTCPClient1.Disconnect;

      canAction(false,'已與 CA G/W 斷線');
    end;
  end;
end;

procedure TfrmMain.handleResponse(sI_ResponseXML: String);
var
    sL_CardID, sL_NewCardID, sL_Date,sL_ResultRegionKey,sL_PersonalRegionLength : String;
    sL_PersonalRegion, sL_ResultReportbackAvailability,sL_ReportbackDay, sL_ReportbackTime, sL_Currency : String;
    sL_ActionResponse, sL_InfoContent : String;
    sL_Action, sL_CreatedSubscriberID, sL_StatusCode : String;
    L_TmpStrList1, L_TmpStrList2 : TStringList;
    sL_SeqNo, sL_CommandID, sL_ErrorCode, sL_ErrorDesc, sL_ResponseData : String;
    L_RootNode : IXmlNode;
    L_Stream : TMemoryStream;
    L_XMLDocument : TXMLDocument;
begin

    try
      L_Stream := TMemoryStream.Create;
      StreamLn(L_Stream,sI_ResponseXML);
      L_XMLDocument := TXMLDocument.Create(Application);
      L_XMLDocument.Active := true;
      L_XMLDocument.Encoding := 'Big5';
      L_XMLDocument.LoadFromStream(L_Stream);
      L_Stream.Free;

      L_RootNode := L_XMLDocument.DocumentElement;
      if L_RootNode.NodeName = ROOT_NODE_RESPONSE_NAME then
      begin
{
<CA_RESPONSE>
    <COMPCODE>88</COMPCODE>
    <SEQNO>123</SEQNO>
    <COMMANDID>A1</COMMANDID>
    <ERRORCODE />
    <ERRORDESC />
    <RESPONSEDATA />
</CA_RESPONSE>
}
        sL_SeqNo := TMyXML.getChildNodeValue(L_RootNode,'SEQNO');
        sL_CommandID := TMyXML.getChildNodeValue(L_RootNode,'COMMANDID');
        sL_ErrorCode := TMyXML.getChildNodeValue(L_RootNode,'ERRORCODE');
        sL_ErrorDesc := TMyXML.getChildNodeValue(L_RootNode,'ERRORDESC');
        sL_ResponseData := TMyXML.getChildNodeValue(L_RootNode,'RESPONSEDATA');

        L_XMLDocument.Free;

        sL_Action := sL_SeqNo + ' [' + sG_Action + '] ';

        if sL_CommandID='A1' then//表示是開機指令...
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
              canAction(true,'');
              MessageDlg(sL_Action + '  指令傳回結果格式錯誤,無法解讀!結果字串:' +sL_ResponseData, mtError, [mbOK],0);
              Exit;
            end;
            L_TmpStrList2.Free;


            L_TmpStrList2 := TUstr.ParseStrings(L_TmpStrList1.Strings[1],SUBSCRIBER_INFO_SEP_FLAG);
            if L_TmpStrList2.Count = 9 then
            begin
              sL_CardID := L_TmpStrList2.Strings[0];
              sL_NewCardID := L_TmpStrList2.Strings[1];
              sL_Date := L_TmpStrList2.Strings[2];
              sL_ResultRegionKey := L_TmpStrList2.Strings[3];
//              sL_PersonalRegionLength := L_TmpStrList2.Strings[4];
              sL_PersonalRegion := L_TmpStrList2.Strings[4];
              sL_ResultReportbackAvailability := L_TmpStrList2.Strings[5];
              sL_ReportbackDay := L_TmpStrList2.Strings[6];
              sL_ReportbackTime := L_TmpStrList2.Strings[7];
              sL_Currency := L_TmpStrList2.Strings[8];

            end
            else
            begin
              setCursor(2);
              L_TmpStrList1.Free;
              L_TmpStrList2.Free;
              canAction(true,'');
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
                canAction(true,sL_ActionResponse);
              end
              else if sL_StatusCode='3' then
              begin
                canAction(true,'');
                MessageDlg(sL_Action + '　失敗! 因為subscriber 已經存在.', mtError, [mbOK],0)
              end
              else
              begin
                canAction(true,'');
                MessageDlg(sL_Action + '　失敗! 因為無法解讀 subscriber status.', mtError, [mbOK],0);
              end;

            end;
          end
          else
          begin
            setCursor(2);
            L_TmpStrList1.Free;

    //            sL_ActionResponse := sL_Action + ' 指令傳回結果格式錯誤,無法解讀!結果字串:' +sL_ResponseData;
    //            Panel3.Caption := sL_ActionResponse;

            canAction(true,'');
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
            canAction(true,sL_ActionResponse);
          end
          else
          begin
            canAction(true,'');
            MessageDlg(sL_Action + ' 錯誤!錯誤代碼: ' +sL_ErrorCode + ' ,錯誤訊息: ' + sL_ErrorDesc, mtError, [mbOK],0);
          end;

        end;

        if sL_ErrorCode='' then
        begin
          //Log一筆資料到SimpleCA.mdb
          dtmMain.insertDataToLog(sG_SubscriberID,sG_IccNo,sG_HighLevelCmdID,sG_Action,sG_ExpirationDate,sG_UpdTime,sG_Operator);
        end;
      end
      else if L_RootNode.NodeName=ROOT_NODE_INFO_NAME then
      begin
        //傳回的XML 並不是 CA 指令之 response, 而是一般的 information..
{
<INFO>
    <CONTENT>Connection OK</CONTENT>
</INFO>
}
        sL_InfoContent := TMyXML.getChildNodeValue(L_RootNode,'CONTENT');
        //Panel3.Caption := sL_InfoContent;
        canAction(true,sL_InfoContent);
      end;
    except

    end;
end;



function TfrmMain.setInfoXml(sI_CompName,
  sI_Content: String): WideString;
var
    L_XMLDocument : TXMLDocument;
    L_RootNode,L_ChildNode : IXmlNode;
    sL_Result : String;
    ii : Integer;
    wL_WideString : WideString;
begin
    L_XMLDocument := TXMLDocument.Create(Application);
    L_XMLDocument.Active := true;
    L_XMLDocument.Encoding := 'Big5';
    L_RootNode := TMyXML.createRootNode(L_XMLDocument,ROOT_NODE_INFO_NAME,ROOT_NODE_INFO_VALUE, sL_Result);
    L_ChildNode := TMyXML.addChild(L_RootNode,'COMPNAME',sI_CompName);
    L_ChildNode := TMyXML.addChild(L_RootNode,'CONTENT',sI_Content);

    for ii:=0 to L_XMLDocument.XML.Count -1 do
      wL_WideString := wL_WideString + L_XMLDocument.XML[ii];

    Result := wL_WideString;
end;

procedure TfrmMain.canAction(bI_CanAction: Boolean;sI_ReturnString : String);
begin
    if bI_CanAction then
    begin
      Panel3.Caption := sI_ReturnString;
      Panel3.Font.Color := clBlue;

      BitBtn2.Enabled := true;
      BitBtn3.Enabled := true;
      BitBtn4.Enabled := true;
      BitBtn5.Enabled := true;
      BitBtn7.Enabled := true;
      BitBtn8.Enabled := true;
      btnChooseCh.Enabled := true;

      btnLogin.Enabled := false;
    end
    else
    begin
      Panel3.Caption := sI_ReturnString;
      Panel3.Font.Color := clRed;

      BitBtn2.Enabled := false;
      BitBtn3.Enabled := false;
      BitBtn4.Enabled := false;
      BitBtn5.Enabled := false;
      BitBtn7.Enabled := false;
      BitBtn8.Enabled := false;
      btnChooseCh.Enabled := false;

      btnLogin.Enabled := true;
    end;
end;

procedure TfrmMain.BitBtn8Click(Sender: TObject);
var
    sL_BouquetID, sL_SunscriberID : String;

begin


    sL_BouquetID := Trim(edtBouquetID.Text);
    sL_SunscriberID := Trim(edtSubscriberID.Text);
    if (sL_SunscriberID='') then
    begin
      MessageDlg('請輸入 subscriber id.', mtWarning, [mbOK],0);
      edtSubscriberID.SetFocus;
      Exit;
    end
    else if (sL_BouquetID='') then
    begin
      MessageDlg('請輸入Bouquet ID.', mtWarning, [mbOK],0);
      edtBouquetID.SetFocus;
      Exit;
    end    
    else
      RealAction(6);
end;

procedure TfrmMain.edtSubscriberIDKeyPress(Sender: TObject; var Key: Char);
begin
    if Key in['a'..'z'] then
      Dec(Key,32);  //強迫轉大寫

end;

procedure TfrmMain.edtSubscriberIDExit(Sender: TObject);
var
    sL_SubscriberID : String;        
begin
    sL_SubscriberID := edtSubscriberID.Text;
    if (sL_SubscriberID<>'') and (not TUstr.isHexValue(sL_SubscriberID)) then
    begin
      MessageDlg('Subscriber id 格式錯誤!(必須是16進位之值)', mtWarning, [mbOK],0);
      edtSubscriberID.SetFocus;
      Exit;
    end
end;

procedure TfrmMain.edtBouquetIDExit(Sender: TObject);
var
    sL_BouquetID : String;        
begin
    sL_BouquetID := edtBouquetID.Text;
    if (sL_BouquetID<>'') and (not TUstr.isHexValue(sL_BouquetID)) then
    begin
      MessageDlg('Bouquet id 格式錯誤!(必須是16進位之值)', mtWarning, [mbOK],0);
      edtBouquetID.SetFocus;
      Exit;
    end
end;

procedure TfrmMain.parseInfoXml(sI_Msg: String; var sI_CompName,
  sI_Content: String);
var
    L_XMLDocument : TXMLDocument;
    L_Stream : TMemoryStream;
    L_RootNode : IXmlNode;
begin
    L_XMLDocument := TXMLDocument.Create(Application);
    L_XMLDocument.Active := true;
    L_XMLDocument.Encoding := 'Big5';

    L_Stream := TMemoryStream.Create;
    StreamLn(L_Stream,sI_Msg);
    L_XMLDocument.LoadFromStream(L_Stream);

    L_RootNode := L_XMLDocument.DocumentElement;

    sI_CompName := TMyXML.getChildNodeValue(L_RootNode,'COMPNAME');
    sI_Content := TMyXML.getChildNodeValue(L_RootNode,'CONTENT');


end;

end.
