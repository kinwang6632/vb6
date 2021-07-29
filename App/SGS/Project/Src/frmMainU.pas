unit frmMainU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Sockets, ImgList, Menus, StdCtrls, Buttons, ComCtrls, ExtCtrls,
  IdBaseComponent, IdComponent, IdTCPServer, IdHTTPServer, ActnList ,ADODB,
  xmldom, XMLIntf, msxmldom, XMLDoc,THttpThreadU, IdTCPConnection,
  IdTCPClient, IdHTTP, IdCustomHTTPServer, idSocketHandle, IdMultipartFormData;

type
  TfrmMain = class(TForm)
    StatusBar1: TStatusBar;
    pnlTitle: TPanel;
    Panel3: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Panel2: TPanel;
    Splitter1: TSplitter;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    memErrorLog: TMemo;
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    imlStatus: TImageList;
    Label8: TLabel;
    Label9: TLabel;
    GateWay1: TMenuItem;
    N91: TMenuItem;
    alGeneral: TActionList;
    acActivate: TAction;
    cbActive: TCheckBox;
    Timer1: TTimer;
    PopupMenu1: TPopupMenu;
    getXMLDocument: TXMLDocument;
    BitBtn2: TBitBtn;
    setXMLDocument: TXMLDocument;
    tabTesting: TTabSheet;
    N3: TMenuItem;
    livStatus: TListView;
    chkNightRun: TCheckBox;
    NightRunLog1: TMenuItem;
    Log21: TMenuItem;
    Timer2: TTimer;
    BitBtn1: TBitBtn;
    chkReturnFormat: TCheckBox;
    HTTPServer: TIdHTTPServer;
    HTTP: TIdHTTP;
    xmlDocForUpdateContent: TXMLDocument;
    Button1: TButton;
    procedure N2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure N91Click(Sender: TObject);
    procedure GateWay1Click(Sender: TObject);
    procedure acActivateExecute(Sender: TObject);
    procedure HTTPServerCommandGet(AThread: TIdPeerThread;
      RequestInfo: TIdHTTPRequestInfo; ResponseInfo: TIdHTTPResponseInfo);
    procedure FormShow(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure chkNightRunClick(Sender: TObject);
    procedure NightRunLog1Click(Sender: TObject);
    procedure Log21Click(Sender: TObject);
    procedure Timer2Timer(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure BitBtn1Click(Sender: TObject);
    procedure chkReturnFormatClick(Sender: TObject);
    procedure IdHTTPServer1CommandGet(AThread: TIdPeerThread;
      ARequestInfo: TIdHTTPRequestInfo;
      AResponseInfo: TIdHTTPResponseInfo);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
    sG_BaseXmlFilePath : String;
    bG_ActiveCombineData : boolean;
    G_BaseXmlFilePath : TextFile;
    bG_ColorFlag,bG_ActionNightRun,bG_InputLicense : Boolean;
    nG_LivItemPos,nG_CmdSeq : Integer;
    dG_Sys1, dG_Sys2, dG_Sys3 : TDate;
    sG_Sys4, sG_Sys5 : String;
    //sG_CmdSeq : String;
    procedure updateXmlContent(sI_XmlFilePath, sI_QueryDateTime,sI_SrcMsgID,sI_DstMsgID:String);
    procedure terminateApp;
    function writeRegData(sI_EncryptedLicenseKey:STring) : boolean;
    function GetRegistryInfo : boolean;
    procedure setLanguageString;
    procedure StopNightRunProcess;
    procedure createLivItem;
    procedure clearStatusItem(nI_ItemIndex:Integer); overload;
    Procedure createHttpThread;
    Procedure saveToZipFile(SI_WinZipExePath,sI_XmlFilePath,sI_ZipFilePath,
       sI_OriginalErrCode,sI_OriginalErrMsg : String;var sI_ErrCode,sI_ErrMsg : String);

    procedure setDateTimeFormat;
  public
    { Public declarations }
    sG_OperatorName : STring;
    bG_StopActiveFlag,bG_ShowUI,bG_LoginOk,bG_ReadOnly : Boolean;
    G_HttpThread : THttpThread;
    G_HttpThread_ExitCode : Cardinal;
    bG_LongFormat : Boolean;
    procedure showSendMsgUI(sI_CmdSeq,sI_DstMsgID,sI_DstCode,sI_CompCode,sI_HandleEDateTime,
                      sI_QueryType,sI_InstQuery : String);

    procedure updateReturnDataUI(sI_CmdSeq,sI_DstMsgID,sI_RetCode,sI_Msg,sI_SrcMsgID : String);

    procedure sendResponseXml(sI_CompCode,sI_ResponseXml : String);

    function getRendomFileName : String;
    function parseXmlMsgAttribute(sI_TmpXmlFilePath : String;var sI_Version,sI_DstMsgID,sI_Type,sI_SrcCode,sI_QueryDateTime,sI_DstCode,sI_SrcType,sI_CasID,sI_FirstDate,sI_CompCode,sI_ErrCode,sI_ErrMsg : String) : Boolean;
    function parseXmlQueryData(sI_TmpXmlFilePath : String;var sI_InstQuery,sI_HandleDate,sI_QueryType,sI_ErrCode,sI_ErrMsg : String) : Boolean;    
    procedure showLogMsg(sI_LogMsg : String);
    function actionQuery(sI_CmdSeq,sI_ModeType,sI_QueryType,sI_CompCode,sI_SrcCode,sI_DstCode,sI_DstMsgID,sI_CasID,sI_InstQuery,sI_FirstDate,sI_HandleEDateTime,sI_Version,sI_DstType,sI_QueryDateTime,sI_SrcType : String) : String;
    function getCmdSeq(sI_CurrDate : String) : String;
    procedure combineData(sI_CompCode,sI_Mode,sI_QueryStartDate,sI_QueryStopDate : String);
    Procedure StartNightRunProcess(bI_AutoRun : Boolean;sI_CompCode,sI_ModeType,sI_QueryStartDate,sI_QueryStopDate : String; bI_CombineData: boolean);

    function returnExchangeDateQueryXML(bI_HaveData : Boolean;sI_CmdSeq,sI_CompCode,
                                        sI_FirstDate,sI_Version,sI_SrcType,sI_QueryDateTime,
                                        sI_DstType,sI_SrcCode,sI_DstCode,sI_DstMsgID,
                                        sI_HandleEDateTime,sI_InstQuery,sI_QueryType : String) : WideString;

    function returnQueryXML(bI_InitBaseXmlFile:boolean; sI_QueryType,sI_Table,sI_CmdSeq,sI_CompCode,sI_Version,sI_QueryDateTime,
                                  sI_SrcCode,sI_SrcType,sI_SrcMsgID,
                                  sI_DstCode,sI_DstType,sI_DstMsgID,
                                  sI_CasID,sI_InstQuery,sI_FirstDate,sI_HandleEDateTime,
                                  sI_RetCode,sI_Msg,
                                  sI_TotalICCardNum,sI_TotalSubNum,sI_TotalProductNum,sI_TotalChannelNum : String;
                                  nI_TotalExecCounts : Integer) : WideString;

    function returnZipXML(sI_QueryType,sI_XmlFilePath,sI_ResponXml,sI_CompCode,sI_Version,sI_QueryDateTime,
                                     sI_SrcCode,sI_SrcType,sI_SrcMsgID,
                                     sI_DstCode,sI_DstType,sI_DstMsgID,
                                     sI_CasID,sI_InstQuery,sI_FirstDate,sI_HandleEDateTime,
                                     sI_RetCode,sI_Msg,sI_TotalICCardNum,sI_TotalSubNum,
                                     sI_TotalProductNum,sI_TotalChannelNum: String;nI_TotalExecCounts : Integer) : WideString;


    function returnNightRunQueryXML(sI_QueryType,sI_XmlFilePath,sI_ResponXml,
                                          sI_CompCode,sI_Version,sI_QueryDateTime,
                                          sI_SrcCode,sI_SrcType,sI_SrcMsgID,
                                          sI_DstCode,sI_DstType,sI_DstMsgID,
                                          sI_CasID,sI_InstQuery,sI_FirstDate,sI_HandleEDateTime,
                                          sI_RetCode,sI_Msg: String;nI_TotalExecCounts : Integer) : WideString;


    //由已計算出的xml取出總數
    function getXmlICCardNumAndSubNum(sI_XmlFilePath : String;var sI_TotalICCardNum,sI_TotalSubNum : String) : String;
    function getXmlProductNumAndChannelNum(sI_XmlFilePath : String;var sI_TotalProductNum,sI_TotalChannelNum : String) : String;

    //由已計算出的 SO459 取出總數
    function getSO459ICCardNumAndSubNum(sI_CompCode,sI_ModeType,sI_HandleEDateTime : String;var sI_TotalICCardNum,sI_TotalSubNum : String) : String;overload;
    function getSO459ProductNumAndChannelNum(sI_CompCode,sI_ModeType,sI_HandleEDateTime : String;var sI_TotalProductNum,sI_TotalChannelNum : String) : String;

  end;

var
  frmMain: TfrmMain;

implementation

uses frmSysParamU, dtmMainU, frmDbInfoU, ConstU, UdateTimeu, frmLoginU,
  Ustru, xmlU, Uotheru, frmQueryLogU, frmNightRunLogU, frmLoadDataU;

{$R *.dfm}

procedure TfrmMain.N2Click(Sender: TObject);
begin
    if not bG_ReadOnly then
    begin
      frmSysParam := TfrmSysParam.Create(Application);
      frmSysParam.ShowModal;
      frmSysParam.Free;
    end
    else
      MessageDlg(dtmMain.getLanguageSettingInfo('frmMain_Msg_10_Content'),mtError, [mbOK],0);
end;



procedure TfrmMain.terminateApp;
begin
    if MessageDlg(dtmMain.getLanguageSettingInfo('frmMain_Msg_1_Content'), mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      cbActive.Checked := false;
      StopNightRunProcess;
      Application.Terminate;
    end;
end;

procedure TfrmMain.setLanguageString;
begin
    Caption := dtmMain.getLanguageSettingInfo('frmMain_Caption');

    pnlTitle.Caption := dtmMain.getLanguageSettingInfo('frmMain_pnlTitle_1_Caption');
    Label1.Caption := dtmMain.getLanguageSettingInfo('frmMain_Label1_Caption');
    Label2.Caption := dtmMain.getLanguageSettingInfo('frmMain_Label2_Caption');
    Label3.Caption := dtmMain.getLanguageSettingInfo('frmMain_Label3_Caption');
    Label4.Caption := dtmMain.getLanguageSettingInfo('frmMain_Label4_Caption');
    Label5.Caption := dtmMain.getLanguageSettingInfo('frmMain_Label5_Caption');
    Label6.Caption := dtmMain.getLanguageSettingInfo('frmMain_Label6_Caption');
    Label7.Caption := dtmMain.getLanguageSettingInfo('frmMain_Label7_Caption');
    Label8.Caption := dtmMain.getLanguageSettingInfo('frmMain_Label8_Caption');
    Label9.Caption := dtmMain.getLanguageSettingInfo('frmMain_Label9_Caption');

    MainMenu1.Items[0].Caption := dtmMain.getLanguageSettingInfo('MainMenu1_Item1_Caption');
    MainMenu1.Items[0].Items[0].Caption := dtmMain.getLanguageSettingInfo('MainMenu1_Item1_1_Caption');
    MainMenu1.Items[0].Items[1].Caption := dtmMain.getLanguageSettingInfo('MainMenu1_Item1_2_Caption');



    MainMenu1.Items[1].Caption := dtmMain.getLanguageSettingInfo('MainMenu1_Item2_Caption');
    MainMenu1.Items[1].Items[0].Caption := dtmMain.getLanguageSettingInfo('MainMenu1_Item2_1_Caption');
    MainMenu1.Items[1].Items[1].Caption := dtmMain.getLanguageSettingInfo('MainMenu1_Item2_2_Caption');

    MainMenu1.Items[2].Caption := dtmMain.getLanguageSettingInfo('MainMenu1_Item3_Caption');


    livStatus.Column[0].Caption := dtmMain.getLanguageSettingInfo('frmMain_livStatus_Columns_0_Caption');
    livStatus.Column[1].Caption := dtmMain.getLanguageSettingInfo('frmMain_livStatus_Columns_1_Caption');
    livStatus.Column[2].Caption := dtmMain.getLanguageSettingInfo('frmMain_livStatus_Columns_2_Caption');
    livStatus.Column[3].Caption := dtmMain.getLanguageSettingInfo('frmMain_livStatus_Columns_3_Caption');
    livStatus.Column[4].Caption := dtmMain.getLanguageSettingInfo('frmMain_livStatus_Columns_4_Caption');
    livStatus.Column[5].Caption := dtmMain.getLanguageSettingInfo('frmMain_livStatus_Columns_5_Caption');
    livStatus.Column[6].Caption := dtmMain.getLanguageSettingInfo('frmMain_livStatus_Columns_6_Caption');
    livStatus.Column[7].Caption := dtmMain.getLanguageSettingInfo('frmMain_livStatus_Columns_7_Caption');
    livStatus.Column[8].Caption := dtmMain.getLanguageSettingInfo('frmMain_livStatus_Columns_8_Caption');
    livStatus.Column[9].Caption := dtmMain.getLanguageSettingInfo('frmMain_livStatus_Columns_9_Caption');
    livStatus.Column[10].Caption := dtmMain.getLanguageSettingInfo('frmMain_livStatus_Columns_10_Caption');
    livStatus.Column[11].Caption := dtmMain.getLanguageSettingInfo('frmMain_livStatus_Columns_11_Caption');
    livStatus.Column[12].Caption := dtmMain.getLanguageSettingInfo('frmMain_livStatus_Columns_12_Caption');
    livStatus.Column[13].Caption := dtmMain.getLanguageSettingInfo('frmMain_livStatus_Columns_13_Caption');

end;

procedure TfrmMain.FormCreate(Sender: TObject);
var
    sL_EncryptedLicenseKey : String;
    dL_SysDate : TDate;
    ii : Integer;

    procedure writeFlagToLockSoftware;
    begin
      TUother.RegWrite(REGISTRY_ROOT,'SYS5',2,SYS5_Y); //寫入停用本軟體之FLAG
    end;

begin
        //設定系統日期格式
        setDateTimeFormat;

        frmLogin := TfrmLogin.Create(Application);

        with frmLogin do
        begin
          if ShowModal=mrCancel then
          begin
            bG_LoginOk := false;
          end
          else
          begin
            bG_LoginOk := true;
            Close;
            free;
          end;
        end;

{
    if ParamCount > 0 then
    begin
        for ii:=1 to ParamCount do
        begin
          if UpperCase(ParamStr(ii)) = 'IL' then
          begin
            bG_InputLicense := true;
          end;
        end;
    end;

    if bG_InputLicense then //表示使用者要輸入 license key...
    begin
      try

        sL_EncryptedLicenseKey := InputBox(dtmMain.getLanguageSettingInfo('frmMain_Msg_2_Content'), 'License Key:', '');

        if sL_EncryptedLicenseKey='' then
          Application.Terminate
        else if writeRegData(sL_EncryptedLicenseKey) then
        begin
          MessageDlg(dtmMain.getLanguageSettingInfo('frmMain_Msg_3_Content'), mtInformation, [mbOK],0);
          Application.Terminate;
        end
        else
        begin
          MessageDlg(dtmMain.getLanguageSettingInfo('frmMain_Msg_4_Content'), mtError, [mbOK],0);
          Application.Terminate;
        end;
      except
        MessageDlg(dtmMain.getLanguageSettingInfo('frmMain_Msg_5_Content'), mtError, [mbOK],0);
        Application.Terminate;
      end;
    end;



    if not GetRegistryInfo then
    begin
      MessageDlg(dtmMain.getLanguageSettingInfo('frmMain_Msg_6_Content'), mtError, [mbOK],0 );
      Application.Terminate;
    end
    else
    begin
      dL_SysDate := Date;
      if sG_Sys5=SYS5_Y then //表示停用軟體
      begin
        MessageDlg(dtmMain.getLanguageSettingInfo('frmMain_Msg_7_Content'), mtError, [mbOK],0 );
        Application.Terminate;
      end
      else if (dL_SysDate= dG_Sys1)  then
      begin
        writeFlagToLockSoftware;
        MessageDlg(dtmMain.getLanguageSettingInfo('frmMain_Msg_8_Content'), mtError, [mbOK],0 );
        Application.Terminate;
      end
      else if (dL_SysDate < dG_Sys3)  then
      begin
        writeFlagToLockSoftware;
        MessageDlg(dtmMain.getLanguageSettingInfo('frmMain_Msg_9_Content'), mtError, [mbOK],0 );
        Application.Terminate;
      end
      else
      begin
        frmLogin := TfrmLogin.Create(Application);

        with frmLogin do
        begin
          if ShowModal=mrCancel then
          begin
            bG_LoginOk := false;
          end
          else
          begin
            bG_LoginOk := true;
            Close;
            free;
          end;
        end;
      end;
    end;
}
end;

procedure TfrmMain.N91Click(Sender: TObject);
begin
    if not bG_ReadOnly then
      terminateApp
    else
      MessageDlg(dtmMain.getLanguageSettingInfo('frmMain_Msg_10_Content'),mtError, [mbOK],0);
end;

procedure TfrmMain.GateWay1Click(Sender: TObject);
begin
    if not bG_ReadOnly then
    begin
      frmDbInfo := TfrmDbInfo.Create(Application);
      frmDbInfo.ShowModal;
      frmDbInfo.Free;
    end
    else
      MessageDlg(dtmMain.getLanguageSettingInfo('frmMain_Msg_10_Content'),mtError, [mbOK],0);    
end;

procedure TfrmMain.acActivateExecute(Sender: TObject);
var
  Binding : TIdSocketHandle;
begin
  //關 HTTP Server 監聽
  acActivate.Checked := not acActivate.Checked;
  
  if not HTTPServer.Active then
  begin
    HTTPServer.Bindings.Clear;
    Binding := HTTPServer.Bindings.Add;
    Binding.IP := dtmMain.sG_LocalIP;    
    Binding.Port := dtmMaIN.nG_HttpPort;


  end;


  if acActivate.Checked then
  begin
    bG_ReadOnly := true;

    pnlTitle.Caption := dtmMain.getLanguageSettingInfo('frmMain_pnlTitle_2_Caption');

    //啟 NightRun 監聽
    Timer1.Enabled := true;
    Timer1.Interval := dtmMain.nG_NightRunExecTime * 1000 * 60;//以分鐘來計算


    Timer2.Enabled := true;
    Timer2.Interval := 3 * 1000;//以秒來計算



    showLogMsg('Open HTTP Server');
    try
      HTTPServer.Active := true;
    except
      on e: exception do
      begin
        bG_ReadOnly := false;
        acActivate.Checked := False;
      end;
    end;
  end
  else
  begin
    bG_ReadOnly := false;
    //關閉 NightRun 監聽
    Timer1.Enabled := false;
    Timer2.Enabled := false;

    pnlTitle.Caption := dtmMain.getLanguageSettingInfo('frmMain_pnlTitle_1_Caption');

    HTTPServer.Active := false;
    HTTPServer.Intercept := nil;

    showLogMsg('Close HTTP Server');
  end;

end;

procedure TfrmMain.HTTPServerCommandGet(AThread: TIdPeerThread;
  RequestInfo: TIdHTTPRequestInfo; ResponseInfo: TIdHTTPResponseInfo);
var
    L_HttpThread : THttpThread;
    sL_RequestInfo,sL_ResponXml : String;
    F_FilePath : TextFile;
    nL_FileHandel : Integer;
begin
    //取得監管平台 Request
    //sL_RequestInfo := '<?xml version="1.0" encoding="BIG5" standalone="yes" ?><Msg Version="1" MsgID="2" Type="SMSDown"  DateTime="2002-08-17 15:30:00" SrcCode="110000G01"  DstCode="110000S01"><ExchangeDateQuery /></Msg>';
    //sL_RequestInfo := '<?xml version="1.0" encoding="BIG5" standalone="yes" ?><Msg Version="1" MsgID="2" Type="SMSDown"  DateTime="2002-08-17 15:30:00" SrcCode="110000G01"  DstCode="110000S01"><ICCardQuery InstQuery="1" Date="2003-02-01"></ICCardQuery></Msg>';
    //sL_RequestInfo := '<?xml version="1.0" encoding="BIG5" standalone="yes" ?><Msg Version="1" MsgID="2" Type="SMSDown"  DateTime="2002-08-17 15:30:00" SrcCode="110000G01"  DstCode="110000S01"><CAProductQuery InstQuery="0" Date="2003-02-01"></CAProductQuery></Msg>';
    //sL_RequestInfo := '<?xml version="1.0" encoding="BIG5" standalone="yes" ?><Msg Version="1" MsgID="2" Type="SMSDown"  DateTime="2002-08-17 15:30:00" SrcCode="110000G01"  DstCode="110000S01"><ProdPurchaseQuery InstQuery="0" Date="2003-02-01"></ProdPurchaseQuery></Msg>';
    //sL_RequestInfo := '<?xml version="1.0" encoding="BIG5" standalone="yes" ?><Msg Version="1" MsgID="2" Type="SMSDown"  DateTime="2002-08-17 15:30:00" SrcCode="110000G01"  DstCode="110000S01"><EntitlementQuery InstQuery="0" Date="2003-02-01"></EntitlementQuery></Msg>';
    //ResponseInfo.ContentText :=  sL_ResponXml;
    try
      sL_RequestInfo := TUstr.replaceStr(RequestInfo.Params.GetText,''#$D#$A'','');

      //sL_ResponXml := '';//目前不知道規格,先回應空字串
      //ResponseInfo.ContentText := sL_ResponXml;

      dtmMain.savetofile(RequestInfo.RemoteIP,'c:\RemoteIP.txt');

      if not FileExists('c:\RequestInfo.txt') then
      begin
        nL_FileHandel := FileCreate('c:\RequestInfo.txt');

        if (nL_FileHandel>=0) then
          FileClose(nL_FileHandel);
      end;
      AssignFile(F_FilePath,'c:\RequestInfo.txt');
      Append(F_FilePath);
      Writeln(F_FilePath,'[' +  DateTimeToStr(now)+ ']-' + sL_RequestInfo);
      Flush(F_FilePath);
      CloseFile(F_FilePath);


      L_HttpThread := THttpThread.Create;
      L_HttpThread.sG_RequestXml := sL_RequestInfo;
      L_HttpThread.actionQueryThread;
      L_HttpThread.Free;
    except

    end;
end;

procedure TfrmMain.StartNightRunProcess(bI_AutoRun : Boolean;sI_CompCode,sI_ModeType,sI_QueryStartDate,sI_QueryStopDate : String;bI_CombineData: boolean);
    procedure loadBaseXml(sL_XmlFilePath:String; I_XmlDoc:TXMLDocument);
    begin
      if FileExists(sL_XmlFilePath) then
        I_XmlDoc.LoadFromFile(sL_XmlFilePath);
    end;

    procedure  writeIcCardAttribute(I_BaseXmlDoc:TXMLDocument; sI_TotalICCardNum,sI_TotalSubNum, sI_DataNodeName:String);
    var
      L_DataNode : IXmlNode;
    begin
      L_DataNode := I_BaseXmlDoc.DocumentElement.ChildNodes.FindNode (sI_DataNodeName);
      if L_DataNode<>nil then
      begin
        TMyXML.setAttribute(L_DataNode, 'TotalICCardNum', sI_TotalICCardNum);
        TMyXML.setAttribute(L_DataNode, 'TotalSubNum', sI_TotalSubNum);
      end;

    end;

    procedure  writeCaProductAttribute(I_BaseXmlDoc:TXMLDocument; sI_TotalProductNum,sI_TotalChannelNum, sI_DataNodeName:String);
    var
      L_DataNode : IXmlNode;
    begin
      L_DataNode := I_BaseXmlDoc.DocumentElement.ChildNodes.FindNode (sI_DataNodeName);
      if L_DataNode<>nil then
      begin
        TMyXML.setAttribute(L_DataNode, 'TotalProductNum', sI_TotalProductNum);
        TMyXML.setAttribute(L_DataNode, 'TotalChannelNum', sI_TotalChannelNum);
      end;

    end;

var
    L_BaseXmlDoc : TXMLDocument;
    
    ii,nDataAreaNo,nL_TotalExecCounts,nL_TotalProductNum,nL_TotalChannelNum,nL_TotalICCardNum,nL_TotalSubNum : Integer;
    sL_CompCode,sL_CompName,sL_Result,sL_ActionDate,sL_CurrDate : String;
    sL_DataPath,sL_FirstDate,sL_AdministratorMail1,sL_AdministratorMail2 : String;
    sL_Version,sL_SrcCode,sL_SrcType,sL_DstCode,sL_DstType,sL_CasID : String;

    sL_DstMsgID,sL_InstQuery,sL_HandleEDateTime,sL_DataPath1,sL_DataPath2 : String;
    sL_RetCode,sL_Msg,sL_SrcMsgID,sL_QueryType,sL_CmdSeq,sL_Table,sL_QueryDateTime : String;
    sL_RealRetCode,sL_RealMsg,sL_ModeType,sL_FirstFlag : String;
    sL_TotalICCardNum,sL_TotalSubNum,sL_TotalProductNum,sL_TotalChannelNum : String;
    nL_Ndx : Integer;
    dL_QueryStartDate,dL_QueryStopDate,dL_DiffDate,dL_ActionDate : TDate;
    nL_DiffDate,nL_ActionDate : Integer;
    sL_DataNodeName, sL_XmlFilePath, sL_MasterChildNode, sL_RefDate : String;
    bL_UpdateIcCardAttribute, bL_UpdateCaProductAttribute, bL_InitBaseXmlFile : Boolean;
    sL_BaseTotalICCardNum,sL_BaseTotalSubNum,sL_BaseTotalProductNum,sL_BaseTotalChannelNum : String;
begin
    bL_UpdateIcCardAttribute := false;
    bL_UpdateCaProductAttribute := false;

    Timer1.Enabled := false;
    sL_CurrDate := TUdateTime.GetPureDateStr(NOW,'/');


    if bI_AutoRun then   //自動執行NightRun,忽略傳進來的參數
    begin
      sL_ActionDate := TUdateTime.GetPureDateStr(NOW-1,'/');//執行前一天的資料

      for ii:=0 to  dtmMain.G_SGSParamIniStrList.Count-1 do
      begin
        sL_CompCode := (dtmMain.G_SGSParamIniStrList.Objects[ii] as TSGSParamIniData).sCompCode;
        sL_FirstDate := (dtmMain.G_SGSParamIniStrList.Objects[ii] as TSGSParamIniData).sFirstDate;

        //檢核是否已連上該公司資料庫
        if dtmMain.checkIsConnection(sL_CompCode) then
        begin
//          if sL_ActionDate >= sL_FirstDate then  //可開始執行日期必須大於等於第一天
          if sL_ActionDate > sL_FirstDate then  //可開始執行日期必須大於第一天 (Howard)
          begin
            nDataAreaNo := (dtmMain.G_SGSParamIniStrList.Objects[ii] as TSGSParamIniData).nDataAreaNo;
            sL_CompName := (dtmMain.G_SGSParamIniStrList.Objects[ii] as TSGSParamIniData).sCompName;
            sL_DataPath1 := (dtmMain.G_SGSParamIniStrList.Objects[ii] as TSGSParamIniData).sDataPath1;
            sL_DataPath2 := (dtmMain.G_SGSParamIniStrList.Objects[ii] as TSGSParamIniData).sDataPath2;

            sL_AdministratorMail1 := (dtmMain.G_SGSParamIniStrList.Objects[ii] as TSGSParamIniData).sAdministratorMail1;
            sL_AdministratorMail2 := (dtmMain.G_SGSParamIniStrList.Objects[ii] as TSGSParamIniData).sAdministratorMail2;
            sL_Version := (dtmMain.G_SGSParamIniStrList.Objects[ii] as TSGSParamIniData).sVersion;
            sL_SrcCode := (dtmMain.G_SGSParamIniStrList.Objects[ii] as TSGSParamIniData).sSrcCode;
            sL_SrcType := (dtmMain.G_SGSParamIniStrList.Objects[ii] as TSGSParamIniData).sSrcType;

            sL_DstCode := (dtmMain.G_SGSParamIniStrList.Objects[ii] as TSGSParamIniData).sDstCode;
            sL_DstType := (dtmMain.G_SGSParamIniStrList.Objects[ii] as TSGSParamIniData).sDstType;
            sL_CasID := (dtmMain.G_SGSParamIniStrList.Objects[ii] as TSGSParamIniData).sCasID;



            sL_DstMsgID := '';     //執行NightRun時為空值
            sL_InstQuery := '0';   //執行NightRun時固定為歷史查詢0:歷史查詢  1:即時查詢
            sL_HandleEDateTime := sL_ActionDate + ' 23:59:59';
            sL_QueryDateTime := '';

  //執行IC卡資訊之查詢(ICCardQuery)
            sL_CmdSeq := getCmdSeq(sL_CurrDate);
            sL_QueryType := IC_CARD_Query;
            showSendMsgUI(sL_CmdSeq,sL_DstMsgID,sL_DstCode,sL_CompCode,sL_HandleEDateTime,sL_QueryType,sL_InstQuery);

            dtmMain.runICCardQuerySF(sL_CompCode,sL_SrcCode,sL_DstCode,sL_DstMsgID,sL_CasID,
                                     sL_InstQuery,sL_FirstDate,sL_HandleEDateTime,
                                     sL_RetCode,sL_Msg,sL_SrcMsgID,sL_FirstFlag,nL_TotalExecCounts,nL_TotalICCardNum,nL_TotalSubNum);

            //將NightRun 存原始sL_Msg,畫面及回傳Xml回傳處理過的sL_RealRetCode,sL_RealMsg
            sL_ModeType := '1';
            dtmMain.getSFRealMsg(sL_RetCode,sL_Msg,sL_RealRetCode,sL_RealMsg);

            sL_TotalICCardNum := IntToStr(nL_TotalICCardNum);
            sL_TotalSubNum := IntToStr(nL_TotalSubNum);
            sL_TotalProductNum := '';
            sL_TotalChannelNum := '';

            //SO459存原Oracle的錯誤訊息
            dtmMain.logToSO459(sL_CompCode,sL_ModeType,sL_FirstFlag,sL_HandleEDateTime,
                               sL_RetCode,sL_Msg,sL_TotalICCardNum,sL_TotalSubNum,
                               sL_TotalProductNum,sL_TotalChannelNum);

            //轉成XML檔案
            if  sL_InstQuery = '0' then
              sL_Table := 'SO451'
            else
              sL_Table := 'SO452';

            returnQueryXML(bL_InitBaseXmlFile, sL_QueryType,sL_Table,sL_CmdSeq,sL_CompCode,sL_Version,sL_QueryDateTime,
                           sL_SrcCode,sL_SrcType,sL_SrcMsgID,
                           sL_DstCode,sL_DstType,sL_DstMsgID,
                           sL_CasID,sL_InstQuery,sL_FirstDate,sL_HandleEDateTime,
                           sL_RealRetCode,sL_RealMsg,sL_TotalICCardNum,sL_TotalSubNum,
                           sL_TotalProductNum,sL_TotalChannelNum,nL_TotalExecCounts);


            updateReturnDataUI(sL_CmdSeq,sL_DstMsgID,sL_RealRetCode,sL_RealMsg,sL_SrcMsgID);
  //CA產品資訊之查詢(CAProductQuery)
            sL_CmdSeq := getCmdSeq(sL_CurrDate);
            sL_QueryType := CA_PRODUCT_QUERY;
            showSendMsgUI(sL_CmdSeq,sL_DstMsgID,sL_DstCode,sL_CompCode,sL_HandleEDateTime,sL_QueryType,sL_InstQuery);

            dtmMain.runCAProductQuerySF(sL_CompCode,sL_SrcCode,sL_DstCode,sL_DstMsgID,sL_CasID,
                                        sL_InstQuery,sL_FirstDate,sL_HandleEDateTime,
                                        sL_RetCode,sL_Msg,sL_SrcMsgID,sL_FirstFlag,nL_TotalExecCounts,nL_TotalProductNum,nL_TotalChannelNum);

            //將NightRun 存原始sL_Msg,畫面及回傳Xml回傳處理過的sL_RealRetCode,sL_RealMsg
            sL_ModeType := '2';
            dtmMain.getSFRealMsg(sL_RetCode,sL_Msg,sL_RealRetCode,sL_RealMsg);

            sL_TotalICCardNum := '';
            sL_TotalSubNum := '';
            sL_TotalProductNum := IntToStr(nL_TotalProductNum);
            sL_TotalChannelNum := IntToStr(nL_TotalChannelNum);

            //SO459存原Oracle的錯誤訊息
            dtmMain.logToSO459(sL_CompCode,sL_ModeType,sL_FirstFlag,sL_HandleEDateTime,
                               sL_RetCode,sL_Msg,sL_TotalICCardNum,sL_TotalSubNum,
                               sL_TotalProductNum,sL_TotalChannelNum);

            //轉成XML檔案
            if  sL_InstQuery = '0' then
              sL_Table := 'SO453'
            else
              sL_Table := 'SO454';

            returnQueryXML(bL_InitBaseXmlFile, sL_QueryType,sL_Table,sL_CmdSeq,sL_CompCode,sL_Version,sL_QueryDateTime,
                           sL_SrcCode,sL_SrcType,sL_SrcMsgID,
                           sL_DstCode,sL_DstType,sL_DstMsgID,
                           sL_CasID,sL_InstQuery,sL_FirstDate,sL_HandleEDateTime,
                           sL_RealRetCode,sL_RealMsg,sL_TotalICCardNum,sL_TotalSubNum,
                           sL_TotalProductNum,sL_TotalChannelNum,nL_TotalExecCounts);

            updateReturnDataUI(sL_CmdSeq,sL_DstMsgID,sL_RealRetCode,sL_RealMsg,sL_SrcMsgID);
  //產品訂購資訊之查詢(ProdPurchaseQuery)
            sL_CmdSeq := getCmdSeq(sL_CurrDate);
            sL_QueryType := PROD_PURCHASE_QUERY;
            showSendMsgUI(sL_CmdSeq,sL_DstMsgID,sL_DstCode,sL_CompCode,sL_HandleEDateTime,sL_QueryType,sL_InstQuery);

            dtmMain.runProdPurchaseQuerySF(sL_CompCode,sL_SrcCode,sL_DstCode,sL_DstMsgID,sL_CasID,
                                           sL_InstQuery,sL_FirstDate,sL_HandleEDateTime,
                                           dtmMain.sG_CancelChannelStr,sL_RetCode,sL_Msg,sL_SrcMsgID,sL_FirstFlag,nL_TotalExecCounts);

            //將NightRun 存原始sL_Msg,畫面及回傳Xml回傳處理過的sL_RealRetCode,sL_RealMsg
            sL_ModeType := '3';
            dtmMain.getSFRealMsg(sL_RetCode,sL_Msg,sL_RealRetCode,sL_RealMsg);

            sL_TotalICCardNum := '';
            sL_TotalSubNum := '';
            sL_TotalProductNum := '';
            sL_TotalChannelNum := '';

            //SO459存原Oracle的錯誤訊息
            dtmMain.logToSO459(sL_CompCode,sL_ModeType,sL_FirstFlag,sL_HandleEDateTime,
                               sL_RetCode,sL_Msg,sL_TotalICCardNum,sL_TotalSubNum,
                               sL_TotalProductNum,sL_TotalChannelNum);

            //轉成XML檔案
            if  sL_InstQuery = '0' then
              sL_Table := 'SO455'
            else
              sL_Table := 'SO456';

            returnQueryXML(bL_InitBaseXmlFile, sL_QueryType,sL_Table,sL_CmdSeq,sL_CompCode,sL_Version,sL_QueryDateTime,
                           sL_SrcCode,sL_SrcType,sL_SrcMsgID,
                           sL_DstCode,sL_DstType,sL_DstMsgID,
                           sL_CasID,sL_InstQuery,sL_FirstDate,sL_HandleEDateTime,
                           sL_RealRetCode,sL_RealMsg,sL_TotalICCardNum,sL_TotalSubNum,
                           sL_TotalProductNum,sL_TotalChannelNum,nL_TotalExecCounts);

            updateReturnDataUI(sL_CmdSeq,sL_DstMsgID,sL_RealRetCode,sL_RealMsg,sL_SrcMsgID);
  //授權資訊之查詢(EntitlementQuery)
            sL_CmdSeq := getCmdSeq(sL_CurrDate);
            sL_QueryType := ENTITLEMENT_QUERY;
            showSendMsgUI(sL_CmdSeq,sL_DstMsgID,sL_DstCode,sL_CompCode,sL_HandleEDateTime,sL_QueryType,sL_InstQuery);

            dtmMain.runEntitlementQuerySF(sL_CompCode,sL_SrcCode,sL_DstCode,sL_DstMsgID,sL_CasID,
                                          sL_InstQuery,sL_FirstDate,sL_HandleEDateTime,
                                          dtmMain.sG_CancelChannelStr,sL_RetCode,sL_Msg,sL_SrcMsgID,sL_FirstFlag,nL_TotalExecCounts);


            //將NightRun 存原始sL_Msg,畫面及回傳Xml回傳處理過的sL_RealRetCode,sL_RealMsg
            sL_ModeType := '4';
            dtmMain.getSFRealMsg(sL_RetCode,sL_Msg,sL_RealRetCode,sL_RealMsg);

            sL_TotalICCardNum := '';
            sL_TotalSubNum := '';
            sL_TotalProductNum := '';
            sL_TotalChannelNum := '';

            //SO459存原Oracle的錯誤訊息
            dtmMain.logToSO459(sL_CompCode,sL_ModeType,sL_FirstFlag,sL_HandleEDateTime,
                               sL_RetCode,sL_Msg,sL_TotalICCardNum,sL_TotalSubNum,
                               sL_TotalProductNum,sL_TotalChannelNum);


            //轉成XML檔案
            if  sL_InstQuery = '0' then
              sL_Table := 'SO457'
            else
              sL_Table := 'SO458';

            returnQueryXML(bL_InitBaseXmlFile,sL_QueryType,sL_Table,sL_CmdSeq,sL_CompCode,sL_Version,sL_QueryDateTime,
                           sL_SrcCode,sL_SrcType,sL_SrcMsgID,
                           sL_DstCode,sL_DstType,sL_DstMsgID,
                           sL_CasID,sL_InstQuery,sL_FirstDate,sL_HandleEDateTime,
                           sL_RealRetCode,sL_RealMsg,sL_TotalICCardNum,sL_TotalSubNum,
                           sL_TotalProductNum,sL_TotalChannelNum,nL_TotalExecCounts);

            updateReturnDataUI(sL_CmdSeq,sL_DstMsgID,sL_RealRetCode,sL_RealMsg,sL_SrcMsgID);
          end;
        end
        else
        begin
          sL_Result := dtmMain.connectToDBAgain(sL_CompCode);
          if sL_Result <> '' then
            showLogMsg(sL_Result);
        end;
      end;
      chkNightRun.Checked := false;
    end
    else   //手動執行NightRun,使用傳進來的參數
    begin
      sL_CompCode := sI_CompCode;
      sI_QueryStartDate := sI_QueryStartDate;
      sI_QueryStopDate := sI_QueryStopDate;

      dL_QueryStartDate := StrToDate(sI_QueryStartDate);
      dL_QueryStopDate := StrToDate(sI_QueryStopDate);
      nL_DiffDate := Round(dL_QueryStopDate - dL_QueryStartDate);

      if bI_CombineData and (nL_DiffDate >0) then
        bG_ActiveCombineData := true
      else
        bG_ActiveCombineData := false;

      if sI_ModeType = '1' then
      begin
        sL_MasterChildNode := 'ICCardQueryReport';
        sL_QueryType := IC_CARD_Query;

        if bG_ActiveCombineData then
        begin


          sL_DataNodeName := 'ICCardQueryReport';
          sL_RefDate :=  sI_QueryStopDate + ' 23:59:59';
          getSO459ICCardNumAndSubNum(sI_CompCode, sI_ModeType,sL_RefDate,sL_BaseTotalICCardNum,sL_BaseTotalSubNum);
          bL_UpdateIcCardAttribute := true;
        end;
      end
      else if sI_ModeType = '2' then
      begin
        sL_MasterChildNode := 'CAProductQueryReport';
        sL_QueryType := CA_PRODUCT_QUERY;
        if bG_ActiveCombineData then
        begin

          sL_DataNodeName := 'CAProductQueryReport';

          sL_RefDate :=  sI_QueryStopDate + ' 23:59:59';
          getSO459ProductNumAndChannelNum(sI_CompCode, sI_ModeType,sL_RefDate,sL_BaseTotalProductNum,sL_BaseTotalChannelNum);
          bL_UpdateCaProductAttribute := true;
        end;
      end
      else if sI_ModeType = '3' then
      begin
        sL_MasterChildNode := 'ProdPurchaseQueryReport';
        sL_QueryType := PROD_PURCHASE_QUERY;
      end
      else if sI_ModeType = '4' then
      begin
        sL_MasterChildNode := 'EntitlementQueryReport';
        sL_QueryType := ENTITLEMENT_QUERY;
      end;

      frmNightRunLog.showLogMsg('Start...');
      sL_InstQuery := '0';   //執行NightRun時固定為歷史查詢0:歷史查詢  1:即時查詢
      //down, By Howard...2004/07/12
      if bG_ActiveCombineData   then
      begin
        sL_XmlFilePath := dtmMain.getRealXmlFilePath(sL_CompCode,sL_InstQuery,sL_QueryType,sI_QueryStopDate,false,true);
        sG_BaseXmlFilePath := sL_XmlFilePath;

        AssignFile(G_BaseXmlFilePath,sL_XmlFilePath);
        Append(G_BaseXmlFilePath);
      end;
      //up, By Howard...2004/07/12

      for nL_ActionDate:=0 to nL_DiffDate do
      begin
//        dL_ActionDate := dL_QueryStartDate + nL_ActionDate;
        dL_ActionDate := dL_QueryStopDate - nL_ActionDate;//Howard

        sL_ActionDate := DateToStr(dL_ActionDate);

        if bG_ActiveCombineData and (nL_ActionDate=0) then
          bL_InitBaseXmlFile := true
        else
          bL_InitBaseXmlFile := false;

        nL_Ndx := dtmMain.G_SGSParamIniStrList.IndexOf(sI_CompCode);
        if nL_Ndx<>-1 then
        begin
          sL_FirstDate := (dtmMain.G_SGSParamIniStrList.Objects[nL_Ndx] as TSGSParamIniData).sFirstDate;

//          if sL_ActionDate >= sL_FirstDate then  //開始執行日期必須大於等於第一天
          if sL_ActionDate > sL_FirstDate then  //開始執行日期必須大於第一天 (Howard)
          begin
            nDataAreaNo := (dtmMain.G_SGSParamIniStrList.Objects[nL_Ndx] as TSGSParamIniData).nDataAreaNo;
            sL_CompName := (dtmMain.G_SGSParamIniStrList.Objects[nL_Ndx] as TSGSParamIniData).sCompName;
            sL_DataPath1 := (dtmMain.G_SGSParamIniStrList.Objects[nL_Ndx] as TSGSParamIniData).sDataPath1;
            sL_DataPath2 := (dtmMain.G_SGSParamIniStrList.Objects[nL_Ndx] as TSGSParamIniData).sDataPath2;

            sL_AdministratorMail1 := (dtmMain.G_SGSParamIniStrList.Objects[nL_Ndx] as TSGSParamIniData).sAdministratorMail1;
            sL_AdministratorMail2 := (dtmMain.G_SGSParamIniStrList.Objects[nL_Ndx] as TSGSParamIniData).sAdministratorMail2;
            sL_Version := (dtmMain.G_SGSParamIniStrList.Objects[nL_Ndx] as TSGSParamIniData).sVersion;
            sL_SrcCode := (dtmMain.G_SGSParamIniStrList.Objects[nL_Ndx] as TSGSParamIniData).sSrcCode;
            sL_SrcType := (dtmMain.G_SGSParamIniStrList.Objects[nL_Ndx] as TSGSParamIniData).sSrcType;

            sL_DstCode := (dtmMain.G_SGSParamIniStrList.Objects[nL_Ndx] as TSGSParamIniData).sDstCode;
            sL_DstType := (dtmMain.G_SGSParamIniStrList.Objects[nL_Ndx] as TSGSParamIniData).sDstType;
            sL_CasID := (dtmMain.G_SGSParamIniStrList.Objects[nL_Ndx] as TSGSParamIniData).sCasID;


            sL_DstMsgID := '';     //執行NightRun時為空值

            sL_HandleEDateTime := sL_ActionDate + ' 23:59:59';
            sL_QueryDateTime := '';


            if sI_ModeType = '1' then
            begin
              //執行IC卡資訊之查詢(ICCardQuery)
              sL_CmdSeq := getCmdSeq(sL_CurrDate);

              showSendMsgUI(sL_CmdSeq,sL_DstMsgID,sL_DstCode,sL_CompCode,sL_HandleEDateTime,sL_QueryType,sL_InstQuery);

              if not bG_ActiveCombineData then//Howard , 20040713
              begin
                dtmMain.runICCardQuerySF(sL_CompCode,sL_SrcCode,sL_DstCode,sL_DstMsgID,sL_CasID,
                                       sL_InstQuery,sL_FirstDate,sL_HandleEDateTime,
                                       sL_RetCode,sL_Msg,sL_SrcMsgID,sL_FirstFlag,nL_TotalExecCounts,nL_TotalICCardNum,nL_TotalSubNum);


              //將NightRun 存原始sL_Msg,畫面及回傳Xml回傳處理過的sL_RealRetCode,sL_RealMsg
              sL_ModeType := '1';
              dtmMain.getSFRealMsg(sL_RetCode,sL_Msg,sL_RealRetCode,sL_RealMsg);

              sL_TotalICCardNum := IntToStr(nL_TotalICCardNum);
              sL_TotalSubNum := IntToStr(nL_TotalSubNum);
              sL_TotalProductNum := '';
              sL_TotalChannelNum := '';

              //SO459存原Oracle的錯誤訊息
              dtmMain.logToSO459(sL_CompCode,sL_ModeType,sL_FirstFlag,sL_HandleEDateTime,
                                 sL_RetCode,sL_Msg,sL_TotalICCardNum,sL_TotalSubNum,
                                 sL_TotalProductNum,sL_TotalChannelNum);

              end;

              if bG_ActiveCombineData and bL_InitBaseXmlFile then
              begin
                sL_TotalICCardNum := sL_BaseTotalICCardNum;
                sL_TotalSubNum := sL_BaseTotalSubNum;
              end;
              
              //轉成XML檔案
              if  sL_InstQuery = '0' then
                sL_Table := 'SO451'
              else
                sL_Table := 'SO452';


              returnQueryXML(bL_InitBaseXmlFile, sL_QueryType,sL_Table,sL_CmdSeq,sL_CompCode,sL_Version,sL_QueryDateTime,
                             sL_SrcCode,sL_SrcType,sL_SrcMsgID,
                             sL_DstCode,sL_DstType,sL_DstMsgID,
                             sL_CasID,sL_InstQuery,sL_FirstDate,sL_HandleEDateTime,
                             sL_RealRetCode,sL_RealMsg,sL_TotalICCardNum,sL_TotalSubNum,
                             sL_TotalProductNum,sL_TotalChannelNum,nL_TotalExecCounts);

              frmNightRunLog.showLogMsg(Copy(sL_HandleEDateTime,1,10));

              updateReturnDataUI(sL_CmdSeq,sL_DstMsgID,sL_RealRetCode,sL_RealMsg,sL_SrcMsgID);

              if (sL_RealMsg='') or (sL_RealMsg=' ') then
                //MessageDlg(dtmMain.getLanguageSettingInfo('frmNightRunLog_Msg_7_Content'),mtInformation, [mbOK],0)
              else
                frmNightRunLog.showLogMsg(sL_RealMsg);
                //MessageDlg(sL_RealMsg,mtError, [mbOK],0);

            end
            else if sI_ModeType = '2' then
            begin
               //CA產品資訊之查詢(CAProductQuery)
              sL_CmdSeq := getCmdSeq(sL_CurrDate);
              showSendMsgUI(sL_CmdSeq,sL_DstMsgID,sL_DstCode,sL_CompCode,sL_HandleEDateTime,sL_QueryType,sL_InstQuery);
              if not bG_ActiveCombineData then//Howard , 20040713
              begin

              dtmMain.runCAProductQuerySF(sL_CompCode,sL_SrcCode,sL_DstCode,sL_DstMsgID,sL_CasID,
                                          sL_InstQuery,sL_FirstDate,sL_HandleEDateTime,
                                          sL_RetCode,sL_Msg,sL_SrcMsgID,sL_FirstFlag,nL_TotalExecCounts,nL_TotalProductNum,nL_TotalChannelNum);

              //將NightRun 存原始sL_Msg,畫面及回傳Xml回傳處理過的sL_RealRetCode,sL_RealMsg
              sL_ModeType := '2';
              dtmMain.getSFRealMsg(sL_RetCode,sL_Msg,sL_RealRetCode,sL_RealMsg);


              sL_TotalICCardNum := '';
              sL_TotalSubNum := '';
              sL_TotalProductNum := IntToStr(nL_TotalProductNum);
              sL_TotalChannelNum := IntToStr(nL_TotalChannelNum);

              //SO459存原Oracle的錯誤訊息
              dtmMain.logToSO459(sL_CompCode,sL_ModeType,sL_FirstFlag,sL_HandleEDateTime,
                                 sL_RetCode,sL_Msg,sL_TotalICCardNum,sL_TotalSubNum,
                                 sL_TotalProductNum,sL_TotalChannelNum);
              end;

              if bG_ActiveCombineData and bL_InitBaseXmlFile then
              begin
                sL_TotalProductNum := sL_BaseTotalProductNum;
                sL_TotalChannelNum := sL_BaseTotalChannelNum;
              end;

              //轉成XML檔案
              if  sL_InstQuery = '0' then
                sL_Table := 'SO453'
              else
                sL_Table := 'SO454';

              returnQueryXML(bL_InitBaseXmlFile, sL_QueryType,sL_Table,sL_CmdSeq,sL_CompCode,sL_Version,sL_QueryDateTime,
                             sL_SrcCode,sL_SrcType,sL_SrcMsgID,
                             sL_DstCode,sL_DstType,sL_DstMsgID,
                             sL_CasID,sL_InstQuery,sL_FirstDate,sL_HandleEDateTime,
                             sL_RealRetCode,sL_RealMsg,sL_TotalICCardNum,sL_TotalSubNum,
                             sL_TotalProductNum,sL_TotalChannelNum,nL_TotalExecCounts);


              frmNightRunLog.showLogMsg(Copy(sL_HandleEDateTime,1,10));

              updateReturnDataUI(sL_CmdSeq,sL_DstMsgID,sL_RealRetCode,sL_RealMsg,sL_SrcMsgID);

              if (sL_RealMsg='') or (sL_RealMsg=' ') then
                //MessageDlg(dtmMain.getLanguageSettingInfo('frmNightRunLog_Msg_7_Content'),mtInformation, [mbOK],0)
              else
                frmNightRunLog.showLogMsg(sL_RealMsg);
                //MessageDlg(sL_RealMsg,mtError, [mbOK],0);
            end
            else if sI_ModeType = '3' then
            begin
              //產品訂購資訊之查詢(ProdPurchaseQuery)
              sL_CmdSeq := getCmdSeq(sL_CurrDate);
              showSendMsgUI(sL_CmdSeq,sL_DstMsgID,sL_DstCode,sL_CompCode,sL_HandleEDateTime,sL_QueryType,sL_InstQuery);
              if not bG_ActiveCombineData then//Howard , 20040713
              begin

              dtmMain.runProdPurchaseQuerySF(sL_CompCode,sL_SrcCode,sL_DstCode,sL_DstMsgID,sL_CasID,
                                             sL_InstQuery,sL_FirstDate,sL_HandleEDateTime,
                                             dtmMain.sG_CancelChannelStr,sL_RetCode,sL_Msg,sL_SrcMsgID,sL_FirstFlag,nL_TotalExecCounts);

              //將NightRun 存原始sL_Msg,畫面及回傳Xml回傳處理過的sL_RealRetCode,sL_RealMsg
              sL_ModeType := '3';
              dtmMain.getSFRealMsg(sL_RetCode,sL_Msg,sL_RealRetCode,sL_RealMsg);

              sL_TotalICCardNum := '';
              sL_TotalSubNum := '';
              sL_TotalProductNum := '';
              sL_TotalChannelNum := '';

              //SO459存原Oracle的錯誤訊息
              dtmMain.logToSO459(sL_CompCode,sL_ModeType,sL_FirstFlag,sL_HandleEDateTime,
                                 sL_RetCode,sL_Msg,sL_TotalICCardNum,sL_TotalSubNum,
                                 sL_TotalProductNum,sL_TotalChannelNum);
              end;
              //轉成XML檔案
              if  sL_InstQuery = '0' then
                sL_Table := 'SO455'
              else
                sL_Table := 'SO456';

              returnQueryXML(bL_InitBaseXmlFile, sL_QueryType,sL_Table,sL_CmdSeq,sL_CompCode,sL_Version,sL_QueryDateTime,
                             sL_SrcCode,sL_SrcType,sL_SrcMsgID,
                             sL_DstCode,sL_DstType,sL_DstMsgID,
                             sL_CasID,sL_InstQuery,sL_FirstDate,sL_HandleEDateTime,
                             sL_RealRetCode,sL_RealMsg,sL_TotalICCardNum,sL_TotalSubNum,
                             sL_TotalProductNum,sL_TotalChannelNum,nL_TotalExecCounts);


              frmNightRunLog.showLogMsg(Copy(sL_HandleEDateTime,1,10));

              updateReturnDataUI(sL_CmdSeq,sL_DstMsgID,sL_RealRetCode,sL_RealMsg,sL_SrcMsgID);

              if (sL_RealMsg='') or (sL_RealMsg=' ') then
                //MessageDlg(dtmMain.getLanguageSettingInfo('frmNightRunLog_Msg_7_Content'),mtInformation, [mbOK],0)
              else
                frmNightRunLog.showLogMsg(sL_RealMsg);
                //MessageDlg(sL_RealMsg,mtError, [mbOK],0);
            end
            else if sI_ModeType = '4' then
            begin
              //授權資訊之查詢(EntitlementQuery)
              sL_CmdSeq := getCmdSeq(sL_CurrDate);
              showSendMsgUI(sL_CmdSeq,sL_DstMsgID,sL_DstCode,sL_CompCode,sL_HandleEDateTime,sL_QueryType,sL_InstQuery);
              if not bG_ActiveCombineData then//Howard , 20040713
              begin

              dtmMain.runEntitlementQuerySF(sL_CompCode,sL_SrcCode,sL_DstCode,sL_DstMsgID,sL_CasID,
                                            sL_InstQuery,sL_FirstDate,sL_HandleEDateTime,
                                            dtmMain.sG_CancelChannelStr,sL_RetCode,sL_Msg,sL_SrcMsgID,sL_FirstFlag,nL_TotalExecCounts);


              //將NightRun 存原始sL_Msg,畫面及回傳Xml回傳處理過的sL_RealRetCode,sL_RealMsg
              sL_ModeType := '4';
              dtmMain.getSFRealMsg(sL_RetCode,sL_Msg,sL_RealRetCode,sL_RealMsg);

              sL_TotalICCardNum := '';
              sL_TotalSubNum := '';
              sL_TotalProductNum := '';
              sL_TotalChannelNum := '';

              //SO459存原Oracle的錯誤訊息
              dtmMain.logToSO459(sL_CompCode,sL_ModeType,sL_FirstFlag,sL_HandleEDateTime,
                                 sL_RetCode,sL_Msg,sL_TotalICCardNum,sL_TotalSubNum,
                                 sL_TotalProductNum,sL_TotalChannelNum);

              end;
              //轉成XML檔案
              if  sL_InstQuery = '0' then
                sL_Table := 'SO457'
              else
                sL_Table := 'SO458';

              returnQueryXML(bL_InitBaseXmlFile, sL_QueryType,sL_Table,sL_CmdSeq,sL_CompCode,sL_Version,sL_QueryDateTime,
                             sL_SrcCode,sL_SrcType,sL_SrcMsgID,
                             sL_DstCode,sL_DstType,sL_DstMsgID,
                             sL_CasID,sL_InstQuery,sL_FirstDate,sL_HandleEDateTime,
                             sL_RealRetCode,sL_RealMsg,sL_TotalICCardNum,sL_TotalSubNum,
                             sL_TotalProductNum,sL_TotalChannelNum,nL_TotalExecCounts);


              frmNightRunLog.showLogMsg(Copy(sL_HandleEDateTime,1,10));

              updateReturnDataUI(sL_CmdSeq,sL_DstMsgID,sL_RealRetCode,sL_RealMsg,sL_SrcMsgID);

              if (sL_RealMsg='') or (sL_RealMsg=' ') then
                //MessageDlg(dtmMain.getLanguageSettingInfo('frmNightRunLog_Msg_7_Content'),mtInformation, [mbOK],0)
              else
                frmNightRunLog.showLogMsg(sL_RealMsg);
                //MessageDlg(sL_RealMsg,mtError, [mbOK],0);
            end;
          end
          else
          begin
            MessageDlg(dtmMain.getLanguageSettingInfo('frmNightRunLog_Msg_6_Content') + '[' + sL_FirstDate + ']' ,mtError, [mbOK],0);

            //重新執行時間必須大於第一次時間
          end;
        end;
      end;

      frmNightRunLog.showLogMsg('OK.');
      if bG_ActiveCombineData then
      begin

        Writeln(G_BaseXmlFilePath,'</' + sL_MasterChildNode + '>');
        Writeln(G_BaseXmlFilePath,'</Msg>');
        Flush(G_BaseXmlFilePath);
        CloseFile(G_BaseXmlFilePath);
{
        if bL_UpdateIcCardAttribute then
        begin
          L_BaseXmlDoc := TXMLDocument.Create(Application);
          loadBaseXml(sG_BaseXmlFilePath, L_BaseXmlDoc);

          writeIcCardAttribute(L_BaseXmlDoc, sL_BaseTotalICCardNum,sL_BaseTotalSubNum, sL_DataNodeName);
          L_BaseXmlDoc.SaveToFile(sG_BaseXmlFilePath);
          L_BaseXmlDoc.Active := false;
          L_BaseXmlDoc.Free;
        end;

        if bL_UpdateCaProductAttribute then
        begin


          L_BaseXmlDoc := TXMLDocument.Create(Application);
          loadBaseXml(sG_BaseXmlFilePath, L_BaseXmlDoc);

          writeCaProductAttribute(L_BaseXmlDoc,sL_BaseTotalProductNum,sL_BaseTotalChannelNum, sL_DataNodeName);

          L_BaseXmlDoc.SaveToFile(sG_BaseXmlFilePath);
          L_BaseXmlDoc.Active := false;
          L_BaseXmlDoc.Free;

        end;
}
      end;

    end;

    Timer1.Enabled := true;
end;

procedure TfrmMain.FormShow(Sender: TObject);
begin
    if bG_LoginOk then
    begin
      setLanguageString;

      //啟動 HTTP Server 監聽
      cbActive.Checked := true;

      //連接所有資料庫
      dtmMain.connectAllCompanyDB;

      //是否顯示UI
      bG_ShowUI := true;

      createLivItem;

      //指令序號從0開始
      nG_CmdSeq := 0;

      //開啟HttpThread
      createHttpThread;


      tabTesting.TabVisible := false;
    end
    else
      Application.Terminate;
end;

procedure TfrmMain.showLogMsg(sI_LogMsg: String);
begin
    frmMain.memErrorLog.Lines.Add('[' + DateTimeToStr(now) + ']--' + sI_LogMsg);
end;

procedure TfrmMain.StopNightRunProcess;
begin
	Timer1.Enabled := false;
end;

procedure TfrmMain.Timer1Timer(Sender: TObject);
var
    dL_StartRunDateTime,dL_StopRunDateTime,dL_CurrDateTime : TDateTime;
    sL_CurrDate,sL_StartRunDate,sL_StopRunDate : String;
    sL_StopMinute : String;
    nL_StartMinute : Integer;
begin
    nL_StartMinute := 10;
    sL_StartRunDate := TUdateTime.GetPureDateStr(NOW,'/') + ' 00:' + IntToStr(nL_StartMinute) + ':00';
    sL_StopMinute := Format('%.*d',[2,dtmMain.nG_NightRunExecTime + nL_StartMinute + 1]);
    sL_StopRunDate := TUdateTime.GetPureDateStr(NOW,'/') + ' 00:' + sL_StopMinute + ':00';

    dL_StartRunDateTime := StrToDateTime(sL_StartRunDate);
    dL_StopRunDateTime := StrToDateTime(sL_StopRunDate);
    dL_CurrDateTime := now;


    if (dL_CurrDateTime >= dL_StartRunDateTime) and (dL_CurrDateTime <= dL_StopRunDateTime) then
      StartNightRunProcess(true,'','','','',false); //自動執行NightRun,則後面參數省略
end;



procedure TfrmMain.createLivItem;
var
    ii, jj : Integer;
    L_ListItem : TListItem;

begin
    if bG_ShowUI then
    begin
      nG_LivItemPos := 0;
      memErrorLog.Align := alBottom;
      memErrorLog.Height := 120;


      Splitter1.Visible := true;
      Splitter1.Parent := Panel2;
      Splitter1.Height := 4;
      Splitter1.Align := alBottom;

      PageControl1.Visible := true;
      PageControl1.Align := alClient;

      TabSheet1.Caption := dtmMain.getLanguageSettingInfo('frmMain_TabSheet1_1_Caption');


      livStatus.Items.Clear;
      for ii:=0 to SGS_UI_MAX_COUNT -1 do
      begin
        L_ListItem := livStatus.Items.Add;

        for jj:=0 to LISTVIEW_SUB_ITEM_COUNT-1 do
         L_ListItem.SubItems.Add('');

      end;
    end
    else
    begin
      PageControl1.Visible := false;
      TabSheet1.Caption := dtmMain.getLanguageSettingInfo('frmMain_TabSheet1_2_Caption');
      Splitter1.Visible := false;
      memErrorLog.Align := alClient;
    end;
end;

procedure TfrmMain.clearStatusItem(nI_ItemIndex: Integer);
var
    jj : Integer;
begin
      frmMain.livStatus.Items[nI_ItemIndex].Caption := '';
      frmMain.livStatus.Items[nI_ItemIndex].SubItemImages[3] := -1;
      frmMain.livStatus.Items[nI_ItemIndex].SubItemImages[4] := -1;
      frmMain.livStatus.Items[nI_ItemIndex].SubItemImages[5] := -1;
      for jj:=0 to frmMain.livStatus.Items[nI_ItemIndex].SubItems.Count -1 do
        frmMain.livStatus.Items[nI_ItemIndex].SubItems.Strings[jj] := '';
end;



procedure TfrmMain.showSendMsgUI(sI_CmdSeq,sI_DstMsgID, sI_DstCode, sI_CompCode,
  sI_HandleEDateTime, sI_QueryType, sI_InstQuery: String);
var
    jj, nL_NextItemIndex : Integer;
    L_ListItem : TListItem;
    sL_InstQuery : String;
begin
    try
      if sI_CompCode=TESTING_CMD_COMP_CODE then Exit;
//      if StrToIntDef(Copy(sI_TransNo,4,6),0)=0 then Exit;
      if nG_LivItemPos=SGS_UI_MAX_COUNT then
      begin
//        clearStatusItem; //呼叫此 function...似乎會讓 performance 下降...
        nG_LivItemPos := 0;
      end;
      L_ListItem := livStatus.Items[frmMain.nG_LivItemPos];



      for jj:=0 to LISTVIEW_SUB_ITEM_COUNT-1 do
       L_ListItem.SubItems.Add('');

      Inc(nG_LivItemPos);

      //down, 將下一個 ListItem 清除成空白
      if nG_LivItemPos=SGS_UI_MAX_COUNT then
        nL_NextItemIndex := 0
      else
        nL_NextItemIndex := nG_LivItemPos;
      clearStatusItem(nL_NextItemIndex);
      //up, 將下一個 ListItem 清除成空白

//      L_ListItem.StateIndex := 0;
      L_ListItem.Caption := sI_CmdSeq;
      L_ListItem.SubItems[0] := sI_DstMsgID; ;
      L_ListItem.SubItems[1] := sI_DstCode; ;
      L_ListItem.SubItems[2] := sI_CompCode;
      L_ListItem.SubItems[4] := sI_HandleEDateTime ;
      L_ListItem.SubItems[5] := sI_QueryType ;

      if sI_InstQuery = '0' then
        sL_InstQuery := dtmMain.getLanguageSettingInfo('frmMain_Msg_12_Content')
      else if sI_InstQuery = '1' then
        sL_InstQuery := dtmMain.getLanguageSettingInfo('frmMain_Msg_13_Content');

      L_ListItem.SubItems[6] := sL_InstQuery ;
      L_ListItem.SubItemImages[7] := 2;//打勾之icon  => 傳送
      L_ListItem.SubItemImages[8] := -1; //清除 icon => 回應
      L_ListItem.SubItems[11] := sG_OperatorName;
      L_ListItem.SubItems[12] := DateTimeToStr(now);

    except

    end;
    livStatus.Refresh;

end;

procedure TfrmMain.updateReturnDataUI(sI_CmdSeq,sI_DstMsgID,sI_RetCode,sI_Msg,sI_SrcMsgID : String);
var
    L_ListItem : TListItem;
    sL_CmdSequence : String;
    nL_MsgIcon : Integer;
begin
    L_ListItem := livStatus.FindCaption(0,sI_CmdSeq, true, true, false );
    if L_ListItem<>nil then
    begin
      L_ListItem.SubItems[3] := sI_SrcMsgID;
      L_ListItem.SubItems[9] := sI_RetCode;
      L_ListItem.SubItems[10] := sI_Msg;
      if (sI_RetCode = '0') Or (sI_RetCode = '1') Or (sI_RetCode = '') then
        nL_MsgIcon := 2
      else
        nL_MsgIcon := 3;

      L_ListItem.SubItemImages[8] := nL_MsgIcon;
    end;
    livStatus.Refresh;
end;

function TfrmMain.parseXmlMsgAttribute(sI_TmpXmlFilePath : String;var sI_Version,
  sI_DstMsgID, sI_Type, sI_SrcCode, sI_QueryDateTime, sI_DstCode,sI_SrcType,
  sI_CasID,sI_FirstDate,sI_CompCode,sI_ErrCode,sI_ErrMsg: String) : Boolean;
var
    L_XMLDoc : TXMLDocument;
    L_NodeList : IXmlNodeList;
    nL_Index : Integer;
    sL_VersionParm,sL_SrcCodeParm,sL_DstCodeParm,sL_DstTypeParm : String;

    bL_Result : Boolean;
begin
    bL_Result := true;
    getXMLDocument.LoadFromFile(sI_TmpXmlFilePath);

    //取得MSG屬性
    L_NodeList := getXMLDocument.ChildNodes;

    //請求及回應的SrcCode,DstCode剛好相反,以下全部以Src代表SMS端,Dst代表監管端
    sI_QueryDateTime := TMyXML.getAttributeValue(L_NodeList.Nodes[1],'DateTime');
    sI_Version := TMyXML.getAttributeValue(L_NodeList.Nodes[1],'Version');
    sI_Type := TMyXML.getAttributeValue(L_NodeList.Nodes[1],'Type');
    sI_DstCode := TMyXML.getAttributeValue(L_NodeList.Nodes[1],'SrcCode');
    sI_SrcCode := TMyXML.getAttributeValue(L_NodeList.Nodes[1],'DstCode');
    sI_DstMsgID := TMyXML.getAttributeValue(L_NodeList.Nodes[1],'MsgID');

    sI_CompCode := dtmMain.getCompCode(sI_SrcCode,nL_Index);

//                       sdf
    sL_VersionParm := (dtmMain.G_SGSParamIniStrList.Objects[nL_Index] as TSGSParamIniData).sVersion;
    sL_SrcCodeParm := (dtmMain.G_SGSParamIniStrList.Objects[nL_Index] as TSGSParamIniData).sSrcCode;
    sL_DstCodeParm := (dtmMain.G_SGSParamIniStrList.Objects[nL_Index] as TSGSParamIniData).sDstCode;
    sL_DstTypeParm := (dtmMain.G_SGSParamIniStrList.Objects[nL_Index] as TSGSParamIniData).sDstType;

    sI_SrcType := (dtmMain.G_SGSParamIniStrList.Objects[nL_Index] as TSGSParamIniData).sSrcType;
    sI_CasID := (dtmMain.G_SGSParamIniStrList.Objects[nL_Index] as TSGSParamIniData).sCasID;
    sI_FirstDate := (dtmMain.G_SGSParamIniStrList.Objects[nL_Index] as TSGSParamIniData).sFirstDate;


    //檢核監管中心/分中心編碼及SMS編碼是否正確
    if (sL_VersionParm<>sI_Version) or (sL_SrcCodeParm<>sI_SrcCode) or
       (sL_DstCodeParm<>sI_DstCode) or (sL_DstTypeParm<>sI_Type) then
    begin
      sI_ErrCode := '2';
      sI_ErrMsg := dtmMain.getLanguageSettingInfo('Return_Msg_2_Content');
      bL_Result := false;
    end;

    Result := bL_Result;
end;


function TfrmMain.getRendomFileName: String;
var
    sL_TmpXmlFileName : String;
    nL_Seed : Integer;
begin
    Randomize;
    nL_Seed := 100000;
    sL_TmpXmlFileName := 'C:\' + IntToStr(Random(nL_Seed)) + '.xml';

    while FileExists(sL_TmpXmlFileName) do
      getRendomFileName;

    result := sL_TmpXmlFileName;
end;

function TfrmMain.parseXmlQueryData(sI_TmpXmlFilePath: String;
  var sI_InstQuery, sI_HandleDate,sI_QueryType,sI_ErrCode,sI_ErrMsg : String) : Boolean;
var
    L_MsgNodeList : IDOMNodeList;
    ii,jj : Integer;
    sL_TempDate : String;
    bL_Result : Boolean;
begin
    bL_Result := true;
    getXMLDocument.LoadFromFile(sI_TmpXmlFilePath);

    L_MsgNodeList := getXMLDocument.DOMDocument.getElementsByTagName('Msg');

    for ii:=0 to L_MsgNodeList.length-1 do
    begin
//      ShowMessage(L_NodeList.item[ii].nodeName);
      for jj:=0 to L_MsgNodeList.item[ii].childNodes.length-1 do
      begin
        sI_QueryType := L_MsgNodeList.item[ii].childNodes.item[jj].nodeName;

        if UpperCase(sI_QueryType) = UpperCase(EXCHANGE_DATE_QUERY) then
        begin
          sI_QueryType := EXCHANGE_DATE_QUERY;
          sI_InstQuery := '1';
        end
        else if (UpperCase(sI_QueryType) = UpperCase(IC_CARD_QUERY)) Or
                (UpperCase(sI_QueryType) = UpperCase(CA_PRODUCT_QUERY)) Or
                (UpperCase(sI_QueryType) = UpperCase(PROD_PURCHASE_QUERY)) Or
                (UpperCase(sI_QueryType) = UpperCase(ENTITLEMENT_QUERY)) then
        begin
         {InstQuery:是否即時查詢
           0: 指定查詢某日期的資訊,此時 Date 屬性必須有值
           1: 即時查詢目前時刻相對於本日00:00  的資訊,此時 Date 屬性可以忽略}
          sI_InstQuery := TMyXML.getAttributeValue(L_MsgNodeList.item[ii].childNodes.item[jj],'InstQuery');
          sL_TempDate := TMyXML.getAttributeValue(L_MsgNodeList.item[ii].childNodes.item[jj],'Date');

          if sI_InstQuery = '0' then  //歷史資料查到當日最後時刻
          begin
            if sL_TempDate = '' then
            begin
              sI_ErrCode := '2';
              sI_ErrMsg := dtmMain.getLanguageSettingInfo('Return_Msg_2_Content');
              bL_Result := false;
            end
            else
              sI_HandleDate := sL_TempDate;
          end
          else if sI_InstQuery <> '1' then    //非0 非1 
          begin
            sI_ErrCode := '2';
            sI_ErrMsg := dtmMain.getLanguageSettingInfo('Return_Msg_2_Content');
            bL_Result := false;
          end;


          //showmessage('InstQuery : =>' + TMyXML.getAttributeValue(L_MsgNodeList.item[ii].childNodes.item[jj],'InstQuery'));
          //showmessage('Date : =>' + TMyXML.getAttributeValue(L_MsgNodeList.item[ii].childNodes.item[jj],'Date'));
          //showmessage(L_MsgNodeList.item[ii].childNodes.item[jj].nodeName + '=>' + TMyXML.getAttributeValue(L_MsgNodeList.item[ii].childNodes.item[jj],'InstQuery'));
        end;
      end;
    end;

    Result := bL_Result;
end;

procedure TfrmMain.BitBtn2Click(Sender: TObject);
var
    L_HttpThread : THttpThread;
    sL_RequestInfo,sL_ResponXml : String;
begin

    //取得監管平台 Request
    //sL_RequestInfo := TUstr.replaceStr(RequestInfo.Params.GetText,''#$D#$A'','');
    sL_RequestInfo := '<?xml version="1.0" encoding="BIG5" standalone="yes" ?><Msg Version="1" MsgID="2" Type="SMSDown"  DateTime="2002-08-17 15:30:00" SrcCode="110000G01"  DstCode="110000S01"><ExchangeDateQuery /></Msg>';
    //sL_RequestInfo := '<?xml version="1.0" encoding="BIG5" standalone="yes" ?><Msg Version="1" MsgID="2" Type="SMSDown"  DateTime="2002-08-17 15:30:00" SrcCode="110000G01"  DstCode="110000S01"><ICCardQuery InstQuery="1" Date="2003-02-01"></ICCardQuery></Msg>';
    //sL_RequestInfo := '<?xml version="1.0" encoding="BIG5" standalone="yes" ?><Msg Version="1" MsgID="2" Type="SMSDown"  DateTime="2002-08-17 15:30:00" SrcCode="110000G01"  DstCode="110000S01"><CAProductQuery InstQuery="0" Date="2003-02-01"></CAProductQuery></Msg>';
    //sL_RequestInfo := '<?xml version="1.0" encoding="BIG5" standalone="yes" ?><Msg Version="1" MsgID="2" Type="SMSDown"  DateTime="2002-08-17 15:30:00" SrcCode="110000G01"  DstCode="110000S01"><ProdPurchaseQuery InstQuery="0" Date="2003-02-01"></ProdPurchaseQuery></Msg>';
    //sL_RequestInfo := '<?xml version="1.0" encoding="BIG5" standalone="yes" ?><Msg Version="1" MsgID="2" Type="SMSDown"  DateTime="2002-08-17 15:30:00" SrcCode="110000G01"  DstCode="110000S01"><EntitlementQuery InstQuery="0" Date="2003-02-01"></EntitlementQuery></Msg>';

    sL_ResponXml := '';//目前不知道規格,先回應空字串
    //ResponseInfo.ContentText := sL_ResponXml;

    try
      L_HttpThread := THttpThread.Create;
      L_HttpThread.sG_RequestXml := sL_RequestInfo;
      L_HttpThread.actionQueryThread;
      L_HttpThread.Free;
    except

    end;
end;

function TfrmMain.actionQuery(sI_CmdSeq,sI_ModeType,sI_QueryType,sI_CompCode, sI_SrcCode, sI_DstCode,
  sI_DstMsgID, sI_CasID, sI_InstQuery, sI_FirstDate,
  sI_HandleEDateTime,sI_Version,sI_DstType,sI_QueryDateTime,sI_SrcType: String): String;
var
    sL_SQL,sL_RetCode,sL_Msg,sL_SrcMsgID,sL_Table,sL_ResponXml,sL_XmlFilePath,sL_ZipFilePath :  String;
    sL_ActionMsg,sL_FirstFlag,sL_RealRetCode,sL_RealMsg : String;
    sL_TotalICCardNum,sL_TotalSubNum,sL_TotalProductNum,sL_TotalChannelNum : String;
    L_AdoQry : TADOQuery;
    bL_HaveData : Boolean;
    nL_TotalExecCounts,nL_TotalICCardNum,nL_TotalSubNum,nL_TotalProductNum,nL_TotalChannelNum : Integer;
    dL_FileSize : Double;
begin
    sL_XmlFilePath := dtmMain.getRealFilePath(sI_CompCode,sI_InstQuery,sI_QueryType,sI_HandleEDateTime,false);
    sL_ZipFilePath := dtmMain.getRealFilePath(sI_CompCode,sI_InstQuery,sI_QueryType,sI_HandleEDateTime,true);

    if UpperCase(sI_QueryType) = UpperCase(EXCHANGE_DATE_QUERY) then
    begin
      sL_SQL := 'SELECT COUNT(*) COUNTS FROM SO459 WHERE CompCode=' + sI_CompCode +
                ' and FirstFlag=1 and ErrorCode is null and ' +
                'To_Char(HandleEDateTime,''YYYY/MM/DD'')=''' +  sI_FirstDate + '''';
      L_AdoQry := dtmMain.runSQL(SELECT_MODE,sI_CompCode,sL_SQL);

      if L_AdoQry=nil then
      begin
        bL_HaveData := false;
        Result := returnExchangeDateQueryXML(bL_HaveData,sI_CmdSeq,sI_CompCode,sI_FirstDate,sI_Version,sI_SrcType,sI_QueryDateTime,sI_DstType,sI_SrcCode,sI_DstCode,sI_DstMsgID,sI_HandleEDateTime,sI_InstQuery,sI_QueryType);
      end
      else
      begin
        if L_AdoQry.FieldByName('COUNTS').AsInteger > 0 then
          bL_HaveData := true
        else
          bL_HaveData := false;

        //回傳遞一天日期
        Result := returnExchangeDateQueryXML(bL_HaveData,sI_CmdSeq,sI_CompCode,sI_FirstDate,sI_Version,sI_SrcType,sI_QueryDateTime,sI_DstType,sI_SrcCode,sI_DstCode,sI_DstMsgID,sI_HandleEDateTime,sI_InstQuery,sI_QueryType);
      end;
    end
//************************************************************************************
//************************************************************************************
    else
    begin
      if sI_InstQuery = '1' then  //即時查詢
      begin
        if UpperCase(sI_QueryType) = UpperCase(IC_CARD_QUERY) then
        begin
          //1:IC卡資訊
          dtmMain.runICCardQuerySF(sI_CompCode,sI_SrcCode,sI_DstCode,sI_DstMsgID,
                                   sI_CasID,sI_InstQuery,sI_FirstDate,sI_HandleEDateTime,
                                   sL_RetCode,sL_Msg,sL_SrcMsgID,sL_FirstFlag,nL_TotalExecCounts,nL_TotalICCardNum,nL_TotalSubNum);

          //將NightRun 存原始sL_Msg,畫面及回傳Xml回傳處理過的sL_RealRetCode,sL_RealMsg
          dtmMain.getSFRealMsg(sL_RetCode,sL_Msg,sL_RealRetCode,sL_RealMsg);


          sL_TotalICCardNum := IntToStr(nL_TotalICCardNum);
          sL_TotalSubNum := IntToStr(nL_TotalSubNum);
          sL_TotalProductNum := '';
          sL_TotalChannelNum := '';

          if  sI_InstQuery = '0' then
            sL_Table := 'SO451'
          else
            sL_Table := 'SO452';
        end
        else if UpperCase(sI_QueryType) = UpperCase(CA_PRODUCT_QUERY) then
        begin
          //2:產品定義資訊
          dtmMain.runCAProductQuerySF(sI_CompCode,sI_SrcCode,sI_DstCode,sI_DstMsgID,sI_CasID,
                                      sI_InstQuery,sI_FirstDate,sI_HandleEDateTime,
                                      sL_RetCode,sL_Msg,sL_SrcMsgID,sL_FirstFlag,nL_TotalExecCounts,nL_TotalProductNum,nL_TotalChannelNum);

          //將NightRun 存原始sL_Msg,畫面及回傳Xml回傳處理過的sL_RealRetCode,sL_RealMsg
          dtmMain.getSFRealMsg(sL_RetCode,sL_Msg,sL_RealRetCode,sL_RealMsg);

          sL_TotalICCardNum := '';
          sL_TotalSubNum := '';
          sL_TotalProductNum := IntToStr(nL_TotalProductNum);
          sL_TotalChannelNum := IntToStr(nL_TotalChannelNum);

          if  sI_InstQuery = '0' then
            sL_Table := 'SO453'
          else
            sL_Table := 'SO454';
        end
        else if UpperCase(sI_QueryType) = UpperCase(PROD_PURCHASE_QUERY) then
        begin
          //3:產品定購資訊
          dtmMain.runProdPurchaseQuerySF(sI_CompCode,sI_SrcCode,sI_DstCode,sI_DstMsgID,sI_CasID,
                                         sI_InstQuery,sI_FirstDate,sI_HandleEDateTime,
                                         dtmMain.sG_CancelChannelStr,sL_RetCode,sL_Msg,sL_SrcMsgID,sL_FirstFlag,nL_TotalExecCounts);

          //將NightRun 存原始sL_Msg,畫面及回傳Xml回傳處理過的sL_RealRetCode,sL_RealMsg
          dtmMain.getSFRealMsg(sL_RetCode,sL_Msg,sL_RealRetCode,sL_RealMsg);

          sL_TotalICCardNum := '';
          sL_TotalSubNum := '';
          sL_TotalProductNum := '';
          sL_TotalChannelNum := '';

          if  sI_InstQuery = '0' then
            sL_Table := 'SO455'
          else
            sL_Table := 'SO456';
        end
        else if UpperCase(sI_QueryType) = UpperCase(ENTITLEMENT_QUERY) then
        begin
          //4:授權資訊
          dtmMain.runEntitlementQuerySF(sI_CompCode,sI_SrcCode,sI_DstCode,sI_DstMsgID,sI_CasID,
                                        sI_InstQuery,sI_FirstDate,sI_HandleEDateTime,
                                        dtmMain.sG_CancelChannelStr,sL_RetCode,sL_Msg,sL_SrcMsgID,sL_FirstFlag,nL_TotalExecCounts);

          //將NightRun 存原始sL_Msg,畫面及回傳Xml回傳處理過的sL_RealRetCode,sL_RealMsg
          dtmMain.getSFRealMsg(sL_RetCode,sL_Msg,sL_RealRetCode,sL_RealMsg);

          sL_TotalICCardNum := '';
          sL_TotalSubNum := '';
          sL_TotalProductNum := '';
          sL_TotalChannelNum := '';

          if  sI_InstQuery = '0' then
            sL_Table := 'SO457'
          else
            sL_Table := 'SO458';
        end;

        sL_ResponXml := returnQueryXML(false, sI_QueryType,sL_Table,sI_CmdSeq,sI_CompCode,sI_Version,sI_QueryDateTime,
                             sI_SrcCode,sI_SrcType,sL_SrcMsgID,
                             sI_DstCode,sI_DstType,sI_DstMsgID,
                             sI_CasID,sI_InstQuery,sI_FirstDate,sI_HandleEDateTime,
                             sL_RealRetCode,sL_RealMsg,sL_TotalICCardNum,sL_TotalSubNum,
                             sL_TotalProductNum,sL_TotalChannelNum,nL_TotalExecCounts);


        dL_FileSize := dtmMain.getFileSize(sL_XmlFilePath);
        if dL_FileSize > MAX_FILE_SIZE then   //大於 1 MB
        begin
          //將原檔案壓縮
          saveToZipFile(dtmMain.SG_WinZipExePath,sL_XmlFilePath,sL_ZipFilePath,sL_RealRetCode,sL_RealMsg,sL_RealRetCode,sL_RealMsg);

          //產生僅基本資料的XML回傳
          sL_ResponXml := returnZipXML(sI_QueryType,sL_XmlFilePath,sL_ResponXml,sI_CompCode,sI_Version,sI_QueryDateTime,
                                     sI_SrcCode,sI_SrcType,sL_SrcMsgID,
                                     sI_DstCode,sI_DstType,sI_DstMsgID,
                                     sI_CasID,sI_InstQuery,sI_FirstDate,sI_HandleEDateTime,
                                     sL_RealRetCode,sL_RealMsg,sL_TotalICCardNum,sL_TotalSubNum,
                             sL_TotalProductNum,sL_TotalChannelNum,nL_TotalExecCounts);
        end;
      end
      else     //歷史查詢
      begin
        sL_SrcMsgID := dtmMain.getSrcMsgID(sI_CompCode,sL_ActionMsg);

        if FileExists(sL_XmlFilePath) then
        begin
          updateXmlContent(sL_XmlFilePath, sI_QueryDateTime,sL_SrcMsgID,sI_DstMsgID);
          dL_FileSize := dtmMain.getFileSize(sL_XmlFilePath);
          if dL_FileSize > MAX_FILE_SIZE then   //大於 1 MB
          begin

            if UpperCase(sI_QueryType) = UpperCase(IC_CARD_QUERY) then
              getXmlICCardNumAndSubNum(sL_XmlFilePath,sL_TotalICCardNum,sL_TotalSubNum)
            else if UpperCase(sI_QueryType) = UpperCase(CA_PRODUCT_QUERY) then
              getXmlProductNumAndChannelNum(sL_XmlFilePath,sL_TotalProductNum,sL_TotalChannelNum);


            //將原檔案壓縮
            saveToZipFile(dtmMain.SG_WinZipExePath,sL_XmlFilePath,sL_ZipFilePath,sL_RealRetCode,sL_RealMsg,sL_RealRetCode,sL_RealMsg);

            //產生僅基本資料的XML回傳
            sL_ResponXml := returnZipXML(sI_QueryType,sL_XmlFilePath,sL_ResponXml,sI_CompCode,sI_Version,sI_QueryDateTime,
                                       sI_SrcCode,sI_SrcType,sL_SrcMsgID,
                                       sI_DstCode,sI_DstType,sI_DstMsgID,
                                       sI_CasID,sI_InstQuery,sI_FirstDate,sI_HandleEDateTime,
                                       sL_RealRetCode,sL_RealMsg,sL_TotalICCardNum,sL_TotalSubNum,
                             sL_TotalProductNum,sL_TotalChannelNum,nL_TotalExecCounts);
          end
          else
          begin
             //將NightRun產生的xml,補填XML資訊,再回傳
             sL_ResponXml := returnNightRunQueryXML(sI_QueryType,sL_XmlFilePath,
                                       sL_ResponXml,sI_CompCode,sI_Version,sI_QueryDateTime,
                                       sI_SrcCode,sI_SrcType,sL_SrcMsgID,
                                       sI_DstCode,sI_DstType,sI_DstMsgID,
                                       sI_CasID,sI_InstQuery,sI_FirstDate,sI_HandleEDateTime,
                                       sL_RealRetCode,sL_RealMsg,nL_TotalExecCounts);

          end;
        end
        else
        begin
          //JACKAL0000 檔案不存在,則檢查XML是被刪掉或是根本沒產生
          //************************
          //檢核SO459 LOG 是否有做過
          if not dtmMain.checkIsAlreadyAction(sI_CompCode,sI_ModeType,sI_HandleEDateTime) then
          begin
            sL_RetCode := '101';
            sL_Msg := dtmMain.getLanguageSettingInfo('Return_Msg_101_Content') ; //SMS尚未產生此查詢日期資料

            sL_TotalICCardNum := '';
            sL_TotalSubNum := '';
            sL_TotalProductNum := '';
            sL_TotalChannelNum := '';


            sL_ResponXml := returnQueryXML(false, sI_QueryType,sL_Table,sI_CmdSeq,sI_CompCode,sI_Version,sI_QueryDateTime,
                               sI_SrcCode,sI_SrcType,sL_SrcMsgID,
                               sI_DstCode,sI_DstType,sI_DstMsgID,
                               sI_CasID,sI_InstQuery,sI_FirstDate,sI_HandleEDateTime,
                               sL_RetCode,sL_Msg,sL_TotalICCardNum,sL_TotalSubNum,
                               sL_TotalProductNum,sL_TotalChannelNum,nL_TotalExecCounts);


            if FileExists(sL_XmlFilePath) then
              DeleteFile(sL_XmlFilePath);
          end
          else
          begin
            //若有做過但Xml檔案被刪除,則重新取資料產生xml回傳
            if UpperCase(sI_QueryType) = UpperCase(IC_CARD_QUERY) then
            begin
              if  sI_InstQuery = '0' then
                sL_Table := 'SO451'
              else
                sL_Table := 'SO452';

              getSO459ICCardNumAndSubNum(sI_CompCode,sI_ModeType,sI_HandleEDateTime,sL_TotalICCardNum,sL_TotalSubNum);
            end
            else if UpperCase(sI_QueryType) = UpperCase(CA_PRODUCT_QUERY) then
            begin
              if  sI_InstQuery = '0' then
                sL_Table := 'SO453'
              else
                sL_Table := 'SO454';
              getSO459ProductNumAndChannelNum(sI_CompCode,sI_ModeType,sI_HandleEDateTime,sL_TotalProductNum,sL_TotalChannelNum);
            end
            else if UpperCase(sI_QueryType) = UpperCase(PROD_PURCHASE_QUERY) then
            begin
              if  sI_InstQuery = '0' then
                sL_Table := 'SO455'
              else
                sL_Table := 'SO456';
            end
            else if UpperCase(sI_QueryType) = UpperCase(ENTITLEMENT_QUERY) then
            begin
              if  sI_InstQuery = '0' then
                sL_Table := 'SO457'
              else
                sL_Table := 'SO458';
            end;

            //取出原計算過的資料
            L_AdoQry := dtmMain.getQueryData(sI_QueryType,sL_Table,sI_CompCode,sI_HandleEDateTime);
            nL_TotalExecCounts := L_AdoQry.RecordCount;


            sL_ResponXml := returnQueryXML(false, sI_QueryType,sL_Table,sI_CmdSeq,sI_CompCode,sI_Version,sI_QueryDateTime,
                                 sI_SrcCode,sI_SrcType,sL_SrcMsgID,
                                 sI_DstCode,sI_DstType,sI_DstMsgID,
                                 sI_CasID,sI_InstQuery,sI_FirstDate,sI_HandleEDateTime,
                                 sL_RealRetCode,sL_RealMsg,sL_TotalICCardNum,sL_TotalSubNum,
                                 sL_TotalProductNum,sL_TotalChannelNum,nL_TotalExecCounts);


            dL_FileSize := dtmMain.getFileSize(sL_XmlFilePath);
            if dL_FileSize > MAX_FILE_SIZE then   //大於 1 MB
            begin

              if UpperCase(sI_QueryType) = UpperCase(IC_CARD_QUERY) then
                getXmlICCardNumAndSubNum(sL_XmlFilePath,sL_TotalICCardNum,sL_TotalSubNum)
              else if UpperCase(sI_QueryType) = UpperCase(CA_PRODUCT_QUERY) then
                getXmlProductNumAndChannelNum(sL_XmlFilePath,sL_TotalProductNum,sL_TotalChannelNum);


              //將原檔案壓縮
              saveToZipFile(dtmMain.SG_WinZipExePath,sL_XmlFilePath,sL_ZipFilePath,sL_RealRetCode,sL_RealMsg,sL_RealRetCode,sL_RealMsg);

              //產生僅基本資料的XML回傳
              sL_ResponXml := returnZipXML(sI_QueryType,sL_XmlFilePath,sL_ResponXml,sI_CompCode,sI_Version,sI_QueryDateTime,
                                         sI_SrcCode,sI_SrcType,sL_SrcMsgID,
                                         sI_DstCode,sI_DstType,sI_DstMsgID,
                                         sI_CasID,sI_InstQuery,sI_FirstDate,sI_HandleEDateTime,
                                         sL_RealRetCode,sL_RealMsg,sL_TotalICCardNum,sL_TotalSubNum,
                               sL_TotalProductNum,sL_TotalChannelNum,nL_TotalExecCounts);
            end
            else
            begin
               //將NightRun產生的xml,補填XML資訊,再回傳
               sL_ResponXml := returnNightRunQueryXML(sI_QueryType,sL_XmlFilePath,
                                         sL_ResponXml,sI_CompCode,sI_Version,sI_QueryDateTime,
                                         sI_SrcCode,sI_SrcType,sL_SrcMsgID,
                                         sI_DstCode,sI_DstType,sI_DstMsgID,
                                         sI_CasID,sI_InstQuery,sI_FirstDate,sI_HandleEDateTime,
                                         sL_RealRetCode,sL_RealMsg,nL_TotalExecCounts);

            end;
          end;
        end;
      end;



      //if sL_RetCode = '0' then
      //  sL_RealRetCode := '';

      dtmMain.logToSO460(sI_CompCode,sI_ModeType,sI_HandleEDateTime,sI_InstQuery,
                sI_SrcCode,sL_SrcMsgID,sI_DstCode,sI_DstMsgID,sL_RetCode,sL_Msg);

      updateReturnDataUI(sI_CmdSeq,sI_DstMsgID,sL_RealRetCode,sL_RealMsg,sL_SrcMsgID);
      Result := sL_ResponXml;
    end;
end;

function TfrmMain.getCmdSeq(sI_CurrDate: String): String;
begin
    nG_CmdSeq := nG_CmdSeq + 1;
    Result := TUstr.replaceStr(sI_CurrDate,'/','') + Format('%.*d',[5,nG_CmdSeq]);
end;

function TfrmMain.returnExchangeDateQueryXML(bI_HaveData : Boolean;
  sI_CmdSeq,sI_CompCode,sI_FirstDate,sI_Version,sI_SrcType,sI_QueryDateTime,
  sI_DstType,sI_SrcCode,sI_DstCode,sI_DstMsgID,sI_HandleEDateTime,sI_InstQuery,sI_QueryType : String): WideString;
var
  sL_Result,sL_SrcMsgID,sL_ActionMsg,sL_ResultXml,sL_SQL : String;
  L_MsgRootNode,L_ReturnChildNode,L_ExchangeChildNode  : IXmlNode;
  sL_RetCode,sL_Msg,sL_ModeType,sL_InstQuery,sL_XmlFilePath : String;
begin
    sL_SrcMsgID := dtmMain.getSrcMsgID(sI_CompCode,sL_ActionMsg);

    setXMLDocument.Active := true;

    if UpperCase(dtmMAIN.sG_TcOrSc) = 'TC' then
      setXMLDocument.Encoding := XML_ENCODING_TC
    else
      setXMLDocument.Encoding := XML_ENCODING_SC;

    setXMLDocument.StandAlone := XML_STANDALONE;

    L_MsgRootNode := TMyXML.createRootNode(setXMLDocument,'Msg','', sL_Result);
    TMyXML.setAttribute(L_MsgRootNode, 'Version',sI_Version);
    TMyXML.setAttribute(L_MsgRootNode, 'MsgID',sL_SrcMsgID);
    TMyXML.setAttribute(L_MsgRootNode, 'Type',sI_SrcType);
    TMyXML.setAttribute(L_MsgRootNode, 'DateTime',sI_QueryDateTime);
    TMyXML.setAttribute(L_MsgRootNode, 'SrcCode',sI_SrcCode);
    TMyXML.setAttribute(L_MsgRootNode, 'DstCode',sI_DstCode);
    TMyXML.setAttribute(L_MsgRootNode, 'ReplyID',sI_DstMsgID);


    L_ReturnChildNode := TMyXML.addChild(L_MsgRootNode, 'Return','');
    TMyXML.setAttribute(L_ReturnChildNode, 'Type',EXCHANGE_DATE_QUERY);

    if sL_ActionMsg = '' then
    begin
      if bI_HaveData then
      begin
        sL_RetCode := '0';
        TMyXML.setAttribute(L_ReturnChildNode, 'Value','0');    //成功
        TMyXML.setAttribute(L_ReturnChildNode, 'Desc',dtmMain.getLanguageSettingInfo('Return_Msg_0_Content'));
      end
      else
      begin
        TMyXML.setAttribute(L_ReturnChildNode, 'Value','1');   //沒查到相應資料
        TMyXML.setAttribute(L_ReturnChildNode, 'Desc',dtmMain.getLanguageSettingInfo('Return_Msg_1_Content'));
        sL_RetCode := '1';
        sL_Msg := dtmMain.getLanguageSettingInfo('Return_Msg_1_Content');
      end;
    end
    else
    begin
      TMyXML.setAttribute(L_ReturnChildNode, 'Value','100');   //資料庫不連線, 公司名稱
      TMyXML.setAttribute(L_ReturnChildNode, 'Desc',dtmMain.getLanguageSettingInfo('Return_Msg_100_Content'));
      sL_RetCode := '100';
      sL_Msg := dtmMain.getLanguageSettingInfo('Return_Msg_100_Content');
    end;
    TMyXML.setAttribute(L_ReturnChildNode, 'Redirect','');


    L_ExchangeChildNode := TMyXML.addChild(L_MsgRootNode, 'ExchangeDateQueryReport','');
    TMyXML.setAttribute(L_ExchangeChildNode, 'FirstDate',TUstr.replaceStr(sI_FirstDate,'/','-'));

    sL_ModeType := '5';
    sL_InstQuery := '1';

    dtmMain.logToSO460(sI_CompCode,sL_ModeType,sI_HandleEDateTime,sL_InstQuery,
              sI_SrcCode,sL_SrcMsgID,sI_DstCode,sI_DstMsgID,sL_RetCode,sL_Msg);

//******************************************************************************
//******************************************************************************

    //存成檔案
    sL_XmlFilePath := dtmMain.getRealFilePath(sI_CompCode,sI_InstQuery,sI_QueryType,sI_HandleEDateTime,false);

    if FileExists(sL_XmlFilePath) then
      DeleteFile(sL_XmlFilePath);

    setXMLDocument.SaveToFile(sL_XmlFilePath);

    Result := setXMLDocument.XML.GetText;

    setXMLDocument.Active := false;
    updateReturnDataUI(sI_CmdSeq,sI_DstMsgID,sL_RetCode,sL_Msg,sL_SrcMsgID);
end;

{
function TfrmMain.returnQueryXML(sI_QueryType,sI_Table,sI_CmdSeq, sI_CompCode, sI_Version,
  sI_QueryDateTime, sI_SrcCode, sI_SrcType, sI_SrcMsgID, sI_DstCode,
  sI_DstType, sI_DstMsgID, sI_CasID, sI_InstQuery, sI_FirstDate,
  sI_HandleEDateTime, sI_RetCode, sI_Msg,sI_TotalICCardNum,sI_TotalSubNum,sI_TotalProductNum,sI_TotalChannelNum: String;
  nI_TotalExecCounts: Integer): WideString;
var
    L_AdoQry : TADOQuery;
    nL_TotalICCardNum,nL_TotalSubNum,nL_TotalProductNum,nL_TotalChannelNum : Integer;
    L_MsgRootNode,L_ReturnChildNode,L_MasterChildNode,L_DetailChildNode,L_SubDetailChildNode  : IXmlNode;
    sL_Result,sL_ResultXml,sL_SQL,sL_HandleDate, sL_XmlString, sL_TmpXmlFilePath : String;
    sL_RetCode,sL_Msg,sL_ModeType,sL_InstQuery,sL_CompName : String;
    sL_RealDateTime,sL_Status,sL_ICCardNo,sL_ICCardID,sL_CasID,sL_Feature : String;
    sL_Version,sL_SubName,sL_SubVocation,sL_SubAge,sL_SubIDCardNo : String;
    sL_XmlFilePath,sL_DataPath1,sL_DataPath2,sL_FileName  : String;
    dL_FileSize : Double;
    sL_MasterChildNode,sL_DetailChildNode,sL_SubDetailChildNode : String;
    sL_ProductCode,sL_ProductName,sL_Provider,sL_ChannelName,sL_TempProductCode : String;
    sL_PurchaseID,sL_StartTime,sL_EndTime : String;
begin
    setXMLDocument.Active := true;

    if UpperCase(dtmMAIN.sG_TcOrSc) = 'TC' then
      setXMLDocument.Encoding := XML_ENCODING_TC
    else
      setXMLDocument.Encoding := XML_ENCODING_SC;

    setXMLDocument.StandAlone := XML_STANDALONE;



    L_MsgRootNode := TMyXML.createRootNode(setXMLDocument,'Msg','', sL_Result);
    TMyXML.setAttribute(L_MsgRootNode, 'Version',sI_Version);
    TMyXML.setAttribute(L_MsgRootNode, 'MsgID',sI_SrcMsgID);
    TMyXML.setAttribute(L_MsgRootNode, 'Type',sI_SrcType);
    TMyXML.setAttribute(L_MsgRootNode, 'DateTime',sI_QueryDateTime);
    TMyXML.setAttribute(L_MsgRootNode, 'SrcCode',sI_SrcCode);
    TMyXML.setAttribute(L_MsgRootNode, 'DstCode',sI_DstCode);
    TMyXML.setAttribute(L_MsgRootNode, 'ReplyID',sI_DstMsgID);

    L_ReturnChildNode := TMyXML.addChild(L_MsgRootNode, 'Return','');

    //TMyXML.setAttribute(L_ReturnChildNode, 'Type',IC_CARD_QUERY);
    TMyXML.setAttribute(L_ReturnChildNode, 'Type',sI_QueryType);

    if (sI_Msg ='') OR (sI_Msg =' ') then
    begin
      if nI_TotalExecCounts > 0 then
      begin
        TMyXML.setAttribute(L_ReturnChildNode, 'Value','0');    //成功
        TMyXML.setAttribute(L_ReturnChildNode, 'Desc',dtmMain.getLanguageSettingInfo('Return_Msg_0_Content'));
      end
      else
      begin
        sL_RetCode := '1';
        sL_Msg := dtmMain.getLanguageSettingInfo('Return_Msg_1_Content');
        TMyXML.setAttribute(L_ReturnChildNode, 'Value',sL_RetCode);   //沒查到相應資料
        TMyXML.setAttribute(L_ReturnChildNode, 'Desc',sL_Msg);
      end;
    end
    else
    begin
      sL_RetCode := sI_RetCode;
      sL_Msg := sI_Msg;
      TMyXML.setAttribute(L_ReturnChildNode, 'Value',sL_RetCode);
      TMyXML.setAttribute(L_ReturnChildNode, 'Desc',sL_Msg);
    end;

    TMyXML.setAttribute(L_ReturnChildNode, 'Redirect','');

//******************************************************************************
//******************************************************************************
    if UpperCase(sI_QueryType) = UpperCase(IC_CARD_QUERY) then
    begin
      sL_MasterChildNode := 'ICCardQueryReport';
      sL_DetailChildNode := 'ICCard';
      sL_SubDetailChildNode := '';
    end
    else if UpperCase(sI_QueryType) = UpperCase(CA_PRODUCT_QUERY) then
    begin
      sL_MasterChildNode := 'CAProductQueryReport';
      sL_DetailChildNode := 'CAProduct';
      sL_SubDetailChildNode := 'Channel';
    end
    else if UpperCase(sI_QueryType) = UpperCase(PROD_PURCHASE_QUERY) then
    begin
      sL_MasterChildNode := 'ProdPurchaseQueryReport';
      sL_DetailChildNode := 'ProdPurchase';
      sL_SubDetailChildNode := '';
    end
    else if UpperCase(sI_QueryType) = UpperCase(ENTITLEMENT_QUERY) then
    begin
      sL_MasterChildNode := 'EntitlementQueryReport';
      sL_DetailChildNode := 'Entitlement';
      sL_SubDetailChildNode := '';
    end;


    L_MasterChildNode := TMyXML.addChild(L_MsgRootNode, sL_MasterChildNode,'');
    TMyXML.setAttribute(L_MasterChildNode, 'InstQuery',sI_InstQuery);

    sL_HandleDate := TUstr.replaceStr(Copy(sI_HandleEDateTime,0,10),'/','-');
    TMyXML.setAttribute(L_MasterChildNode, 'Date',sL_HandleDate);

    if (sI_Msg ='') OR (sI_Msg =' ')  then
    begin
      if nI_TotalExecCounts > 0 then
      begin
        L_AdoQry := dtmMain.getQueryData(sI_QueryType,sI_Table,sI_CompCode,sI_HandleEDateTime);
        dtmMain.adoQryQueryData.Clone(L_AdoQry,ltUnspecified);

        if dtmMain.adoQryQueryData=nil then
        begin
          L_DetailChildNode := TMyXML.addChild(L_MasterChildNode, sL_DetailChildNode,'');
          sL_CompName := dtmMain.getCompName(sI_CompCode);
          sL_RetCode := '100';
          sL_Msg := dtmMain.getLanguageSettingInfo('Return_Msg_100_Content') + '=<' + sL_CompName + '>'; //資料庫不連線, 公司名稱

          TMyXML.setAttribute(L_ReturnChildNode, 'Value',sL_RetCode);
          TMyXML.setAttribute(L_ReturnChildNode, 'Desc',sL_Msg);

          if UpperCase(sI_QueryType) = UpperCase(IC_CARD_QUERY) then
          begin
            TMyXML.setAttribute(L_MasterChildNode, 'TotalICCardNum','');
            TMyXML.setAttribute(L_MasterChildNode, 'TotalSubNum','');
          end
          else if UpperCase(sI_QueryType) = UpperCase(CA_PRODUCT_QUERY) then
          begin
            TMyXML.setAttribute(L_MasterChildNode, 'TotalProductNum','');
            TMyXML.setAttribute(L_MasterChildNode, 'TotalChannelNum','');
          end
        end
        else
        begin
          with dtmMain.adoQryQueryData do
          begin
//**********************   IC卡資訊
            if UpperCase(sI_QueryType) = UpperCase(IC_CARD_QUERY) then
            begin
              //dtmMain.getICCardNumAndSubNum(sI_CompCode,nL_TotalICCardNum,nL_TotalSubNum);//改填傳入值
              TMyXML.setAttribute(L_MasterChildNode, 'TotalICCardNum',sI_TotalICCardNum);
              TMyXML.setAttribute(L_MasterChildNode, 'TotalSubNum',sI_TotalSubNum);

              dtmMain.adoQryQueryData.First;
              while not dtmMain.adoQryQueryData.Eof do
              begin
                L_DetailChildNode := TMyXML.addChild(L_MasterChildNode, sL_DetailChildNode,'');
                sL_RealDateTime := TUstr.replaceStr(DateTimeToStr(FieldByName('RealDateTime').AsDateTime),'/','-');
                if Length(sL_RealDateTime) <= 10 then
                  sL_RealDateTime := sL_RealDateTime + ' 00:00:00';

                sL_Status := FieldByName('Status').AsString;
                if sL_Status = '0' then
                  sL_Status := 'New'
                else if sL_Status = '1' then
                  sL_Status := 'Delete';

                sL_ICCardNo := FieldByName('ICCardNo').AsString;
                sL_ICCardID := FieldByName('ICCardID').AsString;
                sL_CasID := FieldByName('CasID').AsString;
                sL_Feature := FieldByName('Feature').AsString;
                sL_Version := FieldByName('Version').AsString;
                sL_SubName := FieldByName('SubName').AsString;
                sL_SubVocation := FieldByName('SubVocation').AsString;
                sL_SubAge := FieldByName('SubAge').AsString;
                sL_SubIDCardNo := FieldByName('SubIDCardNo').AsString;

                TMyXML.setAttribute(L_DetailChildNode, 'DateTime',sL_RealDateTime);
                TMyXML.setAttribute(L_DetailChildNode, 'Status',sL_Status);
                TMyXML.setAttribute(L_DetailChildNode, 'ICCardNo',sL_ICCardNo);
                TMyXML.setAttribute(L_DetailChildNode, 'ICCardID',sL_ICCardID);
                TMyXML.setAttribute(L_DetailChildNode, 'CasID',sL_CasID);
                TMyXML.setAttribute(L_DetailChildNode, 'Feature',sL_Feature);
                TMyXML.setAttribute(L_DetailChildNode, 'Version',sL_Version);
                TMyXML.setAttribute(L_DetailChildNode, 'SubName',sL_SubName);
                TMyXML.setAttribute(L_DetailChildNode, 'SubVocation',sL_SubVocation);
                TMyXML.setAttribute(L_DetailChildNode, 'SubAge',sL_SubAge);
                TMyXML.setAttribute(L_DetailChildNode, 'SubIDCardNo',sL_SubIDCardNo);

                Next;
              end;
            end
//**********************   產品定義資訊
            else if UpperCase(sI_QueryType) = UpperCase(CA_PRODUCT_QUERY) then
            begin
              //dtmMain.getProductNumAndChannelNum(sI_CompCode,nL_TotalProductNum,nL_TotalChannelNum);//改填傳入值
              TMyXML.setAttribute(L_MasterChildNode, 'TotalProductNum',sI_TotalProductNum);
              TMyXML.setAttribute(L_MasterChildNode, 'TotalChannelNum',sI_TotalChannelNum);

              sL_TempProductCode := '';
              dtmMain.adoQryQueryData.First;
              while not dtmMain.adoQryQueryData.Eof do
              begin
                sL_RealDateTime := TUstr.replaceStr(DateTimeToStr(FieldByName('RealDateTime').AsDateTime),'/','-');
                if Length(sL_RealDateTime) <= 10 then
                  sL_RealDateTime := sL_RealDateTime + ' 00:00:00';

                sL_ProductCode := FieldByName('ProductCode').AsString;
                sL_ProductName := FieldByName('ProductName').AsString;
                sL_CasID := FieldByName('CasID').AsString;
                sL_Provider := FieldByName('Provider').AsString;
                sL_ChannelName := FieldByName('ChannelName').AsString;

                if sL_TempProductCode <> sL_ProductCode then
                begin
                  sL_TempProductCode := sL_ProductCode;

                  L_DetailChildNode := TMyXML.addChild(L_MasterChildNode, sL_DetailChildNode,'');
                  TMyXML.setAttribute(L_DetailChildNode, 'DateTime',sL_RealDateTime);
                  TMyXML.setAttribute(L_DetailChildNode, 'ProductCode',sL_ProductCode);
                  TMyXML.setAttribute(L_DetailChildNode, 'ProductName',sL_ProductName);
                  TMyXML.setAttribute(L_DetailChildNode, 'CasID',sL_CasID);
                end;

                L_SubDetailChildNode := TMyXML.addChild(L_DetailChildNode, sL_SubDetailChildNode,'');
                TMyXML.setAttribute(L_SubDetailChildNode, 'Provider',sL_Provider);
                TMyXML.setAttribute(L_SubDetailChildNode, 'Name',sL_ChannelName);

                Next;
              end;
            end
//**********************   產品定購資訊
            else if UpperCase(sI_QueryType) = UpperCase(PROD_PURCHASE_QUERY) then
            begin
              dtmMain.adoQryQueryData.First;
              while not dtmMain.adoQryQueryData.Eof do
              begin
                L_DetailChildNode := TMyXML.addChild(L_MasterChildNode, sL_DetailChildNode,'');
                sL_RealDateTime := TUstr.replaceStr(DateTimeToStr(FieldByName('RealDateTime').AsDateTime),'/','-');
                if Length(sL_RealDateTime) <= 10 then
                  sL_RealDateTime := sL_RealDateTime + ' 00:00:00';

                sL_Status := FieldByName('Status').AsString;
                if sL_Status = '0' then
                  sL_Status := 'Modify'
                else if sL_Status = '1' then
                  sL_Status := 'Delete';

                sL_PurchaseID := FieldByName('PurchaseID').AsString;
                sL_ICCardNo := FieldByName('ICCardNo').AsString;
                sL_ICCardID := FieldByName('ICCardID').AsString;
                sL_CasID := FieldByName('CasID').AsString;
                sL_ProductCode := FieldByName('ProductCode').AsString;


                sL_StartTime := TUstr.replaceStr(DateTimeToStr(FieldByName('StartTime').AsDateTime),'/','-');
                if Length(sL_StartTime) <= 10 then
                  sL_StartTime := sL_StartTime + ' 00:00:00';

                sL_EndTime := TUstr.replaceStr(DateTimeToStr(FieldByName('EndTime').AsDateTime),'/','-');
                if Length(sL_EndTime) <= 10 then
                  sL_EndTime := sL_EndTime + ' 00:00:00';


                TMyXML.setAttribute(L_DetailChildNode, 'DateTime',sL_RealDateTime);
                TMyXML.setAttribute(L_DetailChildNode, 'Status',sL_Status);
                TMyXML.setAttribute(L_DetailChildNode, 'PurchaseID',sL_PurchaseID);
                TMyXML.setAttribute(L_DetailChildNode, 'ICCardNo',sL_ICCardNo);
                TMyXML.setAttribute(L_DetailChildNode, 'ICCardID',sL_ICCardID);
                TMyXML.setAttribute(L_DetailChildNode, 'CasID',sL_CasID);
                TMyXML.setAttribute(L_DetailChildNode, 'ProductCode',sL_ProductCode);
                TMyXML.setAttribute(L_DetailChildNode, 'StartTime',sL_StartTime);
                TMyXML.setAttribute(L_DetailChildNode, 'EndTime',sL_EndTime);

                Next;
              end;
            end
//**********************   授權資訊
            else if UpperCase(sI_QueryType) = UpperCase(ENTITLEMENT_QUERY) then
            begin
              dtmMain.adoQryQueryData.First;
              while not dtmMain.adoQryQueryData.Eof do
              begin
                L_DetailChildNode := TMyXML.addChild(L_MasterChildNode, sL_DetailChildNode,'');
                sL_RealDateTime := TUstr.replaceStr(DateTimeToStr(FieldByName('RealDateTime').AsDateTime),'/','-');
                if Length(sL_RealDateTime) <= 10 then
                  sL_RealDateTime := sL_RealDateTime + ' 00:00:00';

                sL_Status := FieldByName('Status').AsString;
                if sL_Status = '0' then
                  sL_Status := 'Create'
                else if sL_Status = '1' then
                  sL_Status := 'Delete';

                sL_ICCardNo := FieldByName('ICCardNo').AsString;
                sL_ICCardID := FieldByName('ICCardID').AsString;
                sL_CasID := FieldByName('CasID').AsString;
                sL_ProductCode := FieldByName('ProductCode').AsString;

                TMyXML.setAttribute(L_DetailChildNode, 'DateTime',sL_RealDateTime);
                TMyXML.setAttribute(L_DetailChildNode, 'ICCardNo',sL_ICCardNo);
                TMyXML.setAttribute(L_DetailChildNode, 'ICCardID',sL_ICCardID);
                TMyXML.setAttribute(L_DetailChildNode, 'Status',sL_Status);
                TMyXML.setAttribute(L_DetailChildNode, 'ProductCode',sL_ProductCode);
                TMyXML.setAttribute(L_DetailChildNode, 'CasID',sL_CasID);

                Next;
              end;
            end;
          end;
        end;
      end
      else
      begin
        if UpperCase(sI_QueryType) = UpperCase(IC_CARD_QUERY) then
        begin
          TMyXML.setAttribute(L_MasterChildNode, 'TotalICCardNum','');
          TMyXML.setAttribute(L_MasterChildNode, 'TotalSubNum','');
        end
        else if UpperCase(sI_QueryType) = UpperCase(CA_PRODUCT_QUERY) then
        begin
          TMyXML.setAttribute(L_MasterChildNode, 'TotalProductNum','');
          TMyXML.setAttribute(L_MasterChildNode, 'TotalChannelNum','');
        end;

        L_DetailChildNode := TMyXML.addChild(L_MasterChildNode, sL_DetailChildNode,'');
      end;
    end
    else
    begin
      if UpperCase(sI_QueryType) = UpperCase(IC_CARD_QUERY) then
      begin
        TMyXML.setAttribute(L_MasterChildNode, 'TotalICCardNum','');
        TMyXML.setAttribute(L_MasterChildNode, 'TotalSubNum','');
      end
      else if UpperCase(sI_QueryType) = UpperCase(CA_PRODUCT_QUERY) then
      begin
        TMyXML.setAttribute(L_MasterChildNode, 'TotalProductNum','');
        TMyXML.setAttribute(L_MasterChildNode, 'TotalChannelNum','');
      end;

      L_DetailChildNode := TMyXML.addChild(L_MasterChildNode, sL_DetailChildNode,'');
      sL_RetCode := sI_RetCode;
      sL_Msg := sI_Msg;
    end;
//******************************************************************************
//******************************************************************************
    //存成檔案
    sL_XmlFilePath := dtmMain.getRealFilePath(sI_CompCode,sI_InstQuery,sI_QueryType,sI_HandleEDateTime,false);
    sL_TmpXmlFilePath := ExtractFilePath(sL_XmlFilePath)  +  TUdateTime.GetPureDateStr(now,'') +'.xml';

    if (sI_InstQuery = '1') or
       ((sI_InstQuery = '0') and ((sI_Msg='') and (nI_TotalExecCounts >= 0)) OR ((sI_Msg=' ') and (nI_TotalExecCounts >= 0))) then
    begin
      if FileExists(sL_XmlFilePath) then
        DeleteFile(sL_XmlFilePath);

      setXMLDocument.SaveToFile(sL_XmlFilePath);
    end;

    sL_XmlString := setXMLDocument.XML.GetText;
    setXMLDocument.Active := false;
    Result := sL_XmlString;
end;

}

function TfrmMain.returnQueryXML(bI_InitBaseXmlFile: boolean; sI_QueryType,sI_Table,sI_CmdSeq, sI_CompCode, sI_Version,
  sI_QueryDateTime, sI_SrcCode, sI_SrcType, sI_SrcMsgID, sI_DstCode,
  sI_DstType, sI_DstMsgID, sI_CasID, sI_InstQuery, sI_FirstDate,
  sI_HandleEDateTime, sI_RetCode, sI_Msg,sI_TotalICCardNum,sI_TotalSubNum,sI_TotalProductNum,sI_TotalChannelNum: String;
  nI_TotalExecCounts: Integer): WideString;
var

    L_AdoQry : TADOQuery;
    nL_TotalICCardNum,nL_TotalSubNum,nL_TotalProductNum,nL_TotalChannelNum : Integer;
    L_MsgRootNode,L_ReturnChildNode,L_MasterChildNode,L_DetailChildNode,L_SubDetailChildNode  : IXmlNode;
    sL_Result,sL_ResultXml,sL_SQL,sL_HandleDate, sL_XmlString, sL_TmpXmlFilePath : String;
    sL_RetCode,sL_Msg,sL_ModeType,sL_InstQuery,sL_CompName : String;
    sL_RealDateTime,sL_Status,sL_ICCardNo,sL_ICCardID,sL_CasID,sL_Feature : String;
    sL_Version,sL_SubName,sL_SubVocation,sL_SubAge,sL_SubIDCardNo : String;
    sL_XmlFilePath,sL_DataPath1,sL_DataPath2,sL_FileName  : String;
    dL_FileSize : Double;
    sL_MasterChildNode,sL_DetailChildNode,sL_SubDetailChildNode : String;
    sL_ProductCode,sL_ProductName,sL_Provider,sL_ChannelName,sL_TempProductCode : String;
    sL_PurchaseID,sL_StartTime,sL_EndTime : String;
//*********************************************************
    F_FilePath : TextFile;
    sL_Value,sL_Desc : String;
    nL_TempCount,nL_HandleDataCounts : Integer;
    L_XmlTxt : TStringList ;
    sL_XmlTitle, sL_MsgXmlStr, sL_ReturnXmlStr,sL_MasterXmlStr,sL_DetailXmlStr,sL_SubDetailXmlStr : String;
    sL_RealXmlStr : String;
begin

//jackal
    //開啟要存放的XML檔案


    if not bI_InitBaseXmlFile then
    begin
      sL_XmlFilePath := dtmMain.getRealXmlFilePath(sI_CompCode,sI_InstQuery,sI_QueryType,sI_HandleEDateTime,false,true);
      AssignFile(F_FilePath,sL_XmlFilePath);
      Append(F_FilePath);
    end;

    setXMLDocument.Active := true;
    if UpperCase(dtmMAIN.sG_TcOrSc) = 'TC' then
    begin
      //setXMLDocument.Encoding := XML_ENCODING_TC;
      sL_XmlTitle := '<?xml version="1.0" encoding="' + XML_ENCODING_TC + '" standalone="' + XML_STANDALONE + '" ?> ';
    end
    else
    begin
      //setXMLDocument.Encoding := XML_ENCODING_SC;
      sL_XmlTitle := '<?xml version="1.0" encoding="' + XML_ENCODING_SC + '" standalone="' + XML_STANDALONE + '" ?> ';
    end;
    //setXMLDocument.StandAlone := XML_STANDALONE;
    if  bI_InitBaseXmlFile and bG_ActiveCombineData then
      Writeln(G_BaseXmlFilePath,sL_XmlTitle)
    else
      Writeln(F_FilePath,sL_XmlTitle);

{
    L_MsgRootNode := TMyXML.createRootNode(setXMLDocument,'Msg','', sL_Result);
    TMyXML.setAttribute(L_MsgRootNode, 'Version',sI_Version);
    TMyXML.setAttribute(L_MsgRootNode, 'MsgID',sI_SrcMsgID);
    TMyXML.setAttribute(L_MsgRootNode, 'Type',sI_SrcType);
    TMyXML.setAttribute(L_MsgRootNode, 'DateTime',sI_QueryDateTime);
    TMyXML.setAttribute(L_MsgRootNode, 'SrcCode',sI_SrcCode);
    TMyXML.setAttribute(L_MsgRootNode, 'DstCode',sI_DstCode);
    TMyXML.setAttribute(L_MsgRootNode, 'ReplyID',sI_DstMsgID);
}

    sL_MsgXmlStr := '<Msg Version="' + sI_Version + '" MsgID="' + sI_SrcMsgID +
                 '" Type="' + sI_SrcType + '" DateTime="' + sI_QueryDateTime +
                 '" SrcCode="' + sI_SrcCode + '" DstCode="' + sI_DstCode +
                 '" ReplyID="' + sI_DstMsgID + '">';
    if  bI_InitBaseXmlFile and bG_ActiveCombineData then
      Writeln(G_BaseXmlFilePath,sL_MsgXmlStr)
    else
      Writeln(F_FilePath,sL_MsgXmlStr);


    //L_ReturnChildNode := TMyXML.addChild(L_MsgRootNode, 'Return','');
    //TMyXML.setAttribute(L_ReturnChildNode, 'Type',sI_QueryType);
    if (sI_Msg ='') OR (sI_Msg =' ') then
    begin
      if (nI_TotalExecCounts > 0) or bG_ActiveCombineData then
      begin
        sL_Value := '0';
        sL_Desc := dtmMain.getLanguageSettingInfo('Return_Msg_0_Content');
        //TMyXML.setAttribute(L_ReturnChildNode, 'Value','0');    //成功
        //TMyXML.setAttribute(L_ReturnChildNode, 'Desc',dtmMain.getLanguageSettingInfo('Return_Msg_0_Content'));
      end
      else
      begin
        sL_RetCode := '1';
        sL_Msg := dtmMain.getLanguageSettingInfo('Return_Msg_1_Content');
        sL_Value := '1';
        sL_Desc := dtmMain.getLanguageSettingInfo('Return_Msg_1_Content');
        //TMyXML.setAttribute(L_ReturnChildNode, 'Value',sL_RetCode);   //沒查到相應資料
        //TMyXML.setAttribute(L_ReturnChildNode, 'Desc',sL_Msg);
      end;
    end
    else
    begin
      sL_RetCode := sI_RetCode;
      sL_Msg := sI_Msg;
      sL_Value := sI_RetCode;
      sL_Desc := sI_Msg;
      //TMyXML.setAttribute(L_ReturnChildNode, 'Value',sL_RetCode);
      //TMyXML.setAttribute(L_ReturnChildNode, 'Desc',sL_Msg);
    end;
    //TMyXML.setAttribute(L_ReturnChildNode, 'Redirect','');

    sL_ReturnXmlStr := '<Return Type="' + sI_QueryType + '" Value="' +  sL_Value +
                 '" Desc="' + sL_Desc + '" Redirect=""></Return>';
    if  bI_InitBaseXmlFile and bG_ActiveCombineData then
      Writeln(G_BaseXmlFilePath,sL_ReturnXmlStr)
    else
      Writeln(F_FilePath,sL_ReturnXmlStr);





//******************************************************************************
//******************************************************************************
    if UpperCase(sI_QueryType) = UpperCase(IC_CARD_QUERY) then
    begin
      sL_MasterChildNode := 'ICCardQueryReport';
      sL_DetailChildNode := 'ICCard';
      sL_SubDetailChildNode := '';
    end
    else if UpperCase(sI_QueryType) = UpperCase(CA_PRODUCT_QUERY) then
    begin
      sL_MasterChildNode := 'CAProductQueryReport';
      sL_DetailChildNode := 'CAProduct';
      sL_SubDetailChildNode := 'Channel';
    end
    else if UpperCase(sI_QueryType) = UpperCase(PROD_PURCHASE_QUERY) then
    begin
      sL_MasterChildNode := 'ProdPurchaseQueryReport';
      sL_DetailChildNode := 'ProdPurchase';
      sL_SubDetailChildNode := '';
    end
    else if UpperCase(sI_QueryType) = UpperCase(ENTITLEMENT_QUERY) then
    begin
      sL_MasterChildNode := 'EntitlementQueryReport';
      sL_DetailChildNode := 'Entitlement';
      sL_SubDetailChildNode := '';
    end;


    //L_MasterChildNode := TMyXML.addChild(L_MsgRootNode, sL_MasterChildNode,'');
    //TMyXML.setAttribute(L_MasterChildNode, 'InstQuery',sI_InstQuery);
    sL_HandleDate := TUstr.replaceStr(Copy(sI_HandleEDateTime,0,10),'/','-');
    //TMyXML.setAttribute(L_MasterChildNode, 'Date',sL_HandleDate);

    sL_MasterXmlStr := '<' + sL_MasterChildNode + '  InstQuery="' + sI_InstQuery +
                       '" Date="' + sL_HandleDate;

    nL_HandleDataCounts := 0;
    if (sI_Msg ='') OR (sI_Msg =' ')  then
    begin
      if (nI_TotalExecCounts > 0) or bG_ActiveCombineData then
      begin

        L_AdoQry := dtmMain.getQueryData(sI_QueryType,sI_Table,sI_CompCode,sI_HandleEDateTime);
        dtmMain.adoQryQueryData.Clone(L_AdoQry,ltUnspecified);

        if dtmMain.adoQryQueryData=nil then
        begin
          {jackal以後再補
          L_DetailChildNode := TMyXML.addChild(L_MasterChildNode, sL_DetailChildNode,'');
          sL_CompName := dtmMain.getCompName(sI_CompCode);
          sL_RetCode := '100';
          sL_Msg := dtmMain.getLanguageSettingInfo('Return_Msg_100_Content') + '=<' + sL_CompName + '>'; //資料庫不連線, 公司名稱

          TMyXML.setAttribute(L_ReturnChildNode, 'Value',sL_RetCode);
          TMyXML.setAttribute(L_ReturnChildNode, 'Desc',sL_Msg);

          if UpperCase(sI_QueryType) = UpperCase(IC_CARD_QUERY) then
          begin
            TMyXML.setAttribute(L_MasterChildNode, 'TotalICCardNum','');
            TMyXML.setAttribute(L_MasterChildNode, 'TotalSubNum','');
          end
          else if UpperCase(sI_QueryType) = UpperCase(CA_PRODUCT_QUERY) then
          begin
            TMyXML.setAttribute(L_MasterChildNode, 'TotalProductNum','');
            TMyXML.setAttribute(L_MasterChildNode, 'TotalChannelNum','');
          end
          jackal以後再補}
        end
        else
        begin
          with dtmMain.adoQryQueryData do
          begin
//**********************   IC卡資訊
            if UpperCase(sI_QueryType) = UpperCase(IC_CARD_QUERY) then
            begin
              //TMyXML.setAttribute(L_MasterChildNode, 'TotalICCardNum',sI_TotalICCardNum);
              //TMyXML.setAttribute(L_MasterChildNode, 'TotalSubNum',sI_TotalSubNum);
              sL_MasterXmlStr := sL_MasterXmlStr + '" TotalICCardNum="' + sI_TotalICCardNum +
                                 '" TotalSubNum="' + sI_TotalSubNum + '">';

              if  bI_InitBaseXmlFile and bG_ActiveCombineData then
                Writeln(G_BaseXmlFilePath,sL_MasterXmlStr)
              else
                Writeln(F_FilePath,sL_MasterXmlStr);

              //showmessage(IntToStr(dtmMain.adoQryQueryData.RecordCount));

              dtmMain.adoQryQueryData.DisableControls;
              dtmMain.adoQryQueryData.First;
              while not dtmMain.adoQryQueryData.Eof do
              begin
                //L_DetailChildNode := TMyXML.addChild(L_MasterChildNode, sL_DetailChildNode,'');
                sL_RealDateTime := TUstr.replaceStr(DateTimeToStr(FieldByName('RealDateTime').AsDateTime),'/','-');
                if Length(sL_RealDateTime) <= 10 then
                  sL_RealDateTime := sL_RealDateTime + ' 00:00:00';

                sL_Status := FieldByName('Status').AsString;
                if sL_Status = '0' then
                  sL_Status := 'New'
                else if sL_Status = '1' then
                  sL_Status := 'Delete';

                sL_ICCardNo := FieldByName('ICCardNo').AsString;
                sL_ICCardID := FieldByName('ICCardID').AsString;
                sL_CasID := FieldByName('CasID').AsString;
                sL_Feature := FieldByName('Feature').AsString;
                sL_Version := FieldByName('Version').AsString;
                sL_SubName := FieldByName('SubName').AsString;
                sL_SubVocation := FieldByName('SubVocation').AsString;
                sL_SubAge := FieldByName('SubAge').AsString;
                sL_SubIDCardNo := FieldByName('SubIDCardNo').AsString;

                {
                TMyXML.setAttribute(L_DetailChildNode, 'DateTime',sL_RealDateTime);
                TMyXML.setAttribute(L_DetailChildNode, 'Status',sL_Status);
                TMyXML.setAttribute(L_DetailChildNode, 'ICCardNo',sL_ICCardNo);
                TMyXML.setAttribute(L_DetailChildNode, 'ICCardID',sL_ICCardID);
                TMyXML.setAttribute(L_DetailChildNode, 'CasID',sL_CasID);
                TMyXML.setAttribute(L_DetailChildNode, 'Feature',sL_Feature);
                TMyXML.setAttribute(L_DetailChildNode, 'Version',sL_Version);
                TMyXML.setAttribute(L_DetailChildNode, 'SubName',sL_SubName);
                TMyXML.setAttribute(L_DetailChildNode, 'SubVocation',sL_SubVocation);
                TMyXML.setAttribute(L_DetailChildNode, 'SubAge',sL_SubAge);
                TMyXML.setAttribute(L_DetailChildNode, 'SubIDCardNo',sL_SubIDCardNo);
                }
                sL_DetailXmlStr := '<' + sL_DetailChildNode + ' DateTime="' + sL_RealDateTime +
                                   '" Status="' + sL_Status + '" ICCardNo="' + sL_ICCardNo +
                                   '" ICCardID="' + sL_ICCardID + '" CASID="' + sL_CasID +
                                   '" Feature="' + sL_Feature + '" Version="' + sL_Version +
                                   '" SubName="' + sL_SubName + '" SubVocation="' + sL_SubVocation +
                                   '" SubAge="' + sL_SubAge + '" SubIDCardNo="' + sL_SubIDCardNo +
                                   '" /> ';
                if  bG_ActiveCombineData then
                  Writeln(G_BaseXmlFilePath,sL_DetailXmlStr);
                if not bI_InitBaseXmlFile then
                  Writeln(F_FilePath,sL_DetailXmlStr);
                Next;
              end;
              dtmMain.adoQryQueryData.EnableControls;
            end
//**********************   產品定義資訊
            else if UpperCase(sI_QueryType) = UpperCase(CA_PRODUCT_QUERY) then
            begin
              //TMyXML.setAttribute(L_MasterChildNode, 'TotalProductNum',sI_TotalProductNum);
              //TMyXML.setAttribute(L_MasterChildNode, 'TotalChannelNum',sI_TotalChannelNum);

              sL_MasterXmlStr := sL_MasterXmlStr + '" TotalProductNum="' + sI_TotalProductNum +
                                 '" TotalChannelNum="' + sI_TotalChannelNum + '">';

              if  bI_InitBaseXmlFile and bG_ActiveCombineData then
                Writeln(G_BaseXmlFilePath,sL_MasterXmlStr)
              else
                Writeln(F_FilePath,sL_MasterXmlStr);


              sL_TempProductCode := '';
              nL_TempCount := 0;

              if dtmMain.adoQryQueryData.RecordCount > 0 then
              begin
                dtmMain.adoQryQueryData.DisableControls;
                dtmMain.adoQryQueryData.First;
                while not dtmMain.adoQryQueryData.Eof do
                begin
                  sL_RealDateTime := TUstr.replaceStr(DateTimeToStr(FieldByName('RealDateTime').AsDateTime),'/','-');
                  if Length(sL_RealDateTime) <= 10 then
                    sL_RealDateTime := sL_RealDateTime + ' 00:00:00';

                  sL_ProductCode := FieldByName('ProductCode').AsString;
                  sL_ProductName := FieldByName('ProductName').AsString;
                  sL_CasID := FieldByName('CasID').AsString;
                  sL_Provider := FieldByName('Provider').AsString;
                  sL_ChannelName := FieldByName('ChannelName').AsString;

                  if sL_TempProductCode <> sL_ProductCode then
                  begin
                    if nL_TempCount <> 0 then
                    begin
                      sL_DetailXmlStr := '</' + sL_DetailChildNode + '>';
                      if  bG_ActiveCombineData then
                        Writeln(G_BaseXmlFilePath,sL_DetailXmlStr);
                      if not bI_InitBaseXmlFile then
                        Writeln(F_FilePath,sL_DetailXmlStr);
                    end;

                    nL_TempCount := nL_TempCount + 1;
                    sL_TempProductCode := sL_ProductCode;
                    {
                    L_DetailChildNode := TMyXML.addChild(L_MasterChildNode, sL_DetailChildNode,'');
                    TMyXML.setAttribute(L_DetailChildNode, 'DateTime',sL_RealDateTime);
                    TMyXML.setAttribute(L_DetailChildNode, 'ProductCode',sL_ProductCode);
                    TMyXML.setAttribute(L_DetailChildNode, 'ProductName',sL_ProductName);
                    TMyXML.setAttribute(L_DetailChildNode, 'CasID',sL_CasID);
                    }
                    sL_DetailXmlStr := '<' + sL_DetailChildNode + ' DateTime="' + sL_RealDateTime +
                                       '" ProductCode="' + sL_ProductCode + '" ProductName="' + sL_ProductName +
                                       '" CASID="' + sL_CasID + '" > ';

                    if  bG_ActiveCombineData then
                      Writeln(G_BaseXmlFilePath,sL_DetailXmlStr);
                    if not bI_InitBaseXmlFile then
                      Writeln(F_FilePath,sL_DetailXmlStr);

                  end;
                  {
                  L_SubDetailChildNode := TMyXML.addChild(L_DetailChildNode, sL_SubDetailChildNode,'');
                  TMyXML.setAttribute(L_SubDetailChildNode, 'Provider',sL_Provider);
                  TMyXML.setAttribute(L_SubDetailChildNode, 'Name',sL_ChannelName);
                  }
                  sL_SubDetailXmlStr := '<' + sL_SubDetailChildNode + ' Provider="' + sL_Provider +
                                        '" Name="' + sL_ChannelName + '" />';

                  if  bG_ActiveCombineData then
                    Writeln(G_BaseXmlFilePath,sL_SubDetailXmlStr);
                  if not bI_InitBaseXmlFile then
                    Writeln(F_FilePath,sL_SubDetailXmlStr);

                  if sL_TempProductCode <> sL_ProductCode then
                  begin
                    sL_DetailXmlStr := '</' + sL_DetailChildNode + '>';

                  if  bG_ActiveCombineData then
                    Writeln(G_BaseXmlFilePath,sL_DetailXmlStr);
                  if not bI_InitBaseXmlFile then
                    Writeln(F_FilePath,sL_DetailXmlStr);
                  end;
                  Next;
                end;
                dtmMain.adoQryQueryData.EnableControls;


                sL_DetailXmlStr := '</' + sL_DetailChildNode + '>';

                if  bG_ActiveCombineData then
                  Writeln(G_BaseXmlFilePath,sL_DetailXmlStr);
                if not bI_InitBaseXmlFile then
                  Writeln(F_FilePath,sL_DetailXmlStr);

              end
              else
              begin
                sL_DetailXmlStr := '<' + sL_DetailChildNode + '>';
                sL_DetailXmlStr := sL_DetailXmlStr + '</' + sL_DetailChildNode + '>';

                if  bG_ActiveCombineData then
                  Writeln(G_BaseXmlFilePath,sL_DetailXmlStr);
                if not bI_InitBaseXmlFile then
                  Writeln(F_FilePath,sL_DetailXmlStr);
              end;
            end
//**********************   產品定購資訊
            else if UpperCase(sI_QueryType) = UpperCase(PROD_PURCHASE_QUERY) then
            begin

              sL_MasterXmlStr := sL_MasterXmlStr + '">';


              if  bI_InitBaseXmlFile and bG_ActiveCombineData then
                Writeln(G_BaseXmlFilePath,sL_MasterXmlStr)
              else
                Writeln(F_FilePath,sL_MasterXmlStr);



              dtmMain.adoQryQueryData.DisableControls;
              dtmMain.adoQryQueryData.First;
              while not dtmMain.adoQryQueryData.Eof do
              begin
                //L_DetailChildNode := TMyXML.addChild(L_MasterChildNode, sL_DetailChildNode,'');
                sL_RealDateTime := TUstr.replaceStr(DateTimeToStr(FieldByName('RealDateTime').AsDateTime),'/','-');
                if Length(sL_RealDateTime) <= 10 then
                  sL_RealDateTime := sL_RealDateTime + ' 00:00:00';

                sL_Status := FieldByName('Status').AsString;
                if sL_Status = '0' then
                  sL_Status := 'Modify'
                else if sL_Status = '1' then
                  sL_Status := 'Delete';

                sL_PurchaseID := FieldByName('PurchaseID').AsString;
                sL_ICCardNo := FieldByName('ICCardNo').AsString;
                sL_ICCardID := FieldByName('ICCardID').AsString;
                sL_CasID := FieldByName('CasID').AsString;
                sL_ProductCode := FieldByName('ProductCode').AsString;


                sL_StartTime := TUstr.replaceStr(DateTimeToStr(FieldByName('StartTime').AsDateTime),'/','-');
                if Length(sL_StartTime) <= 10 then
                  sL_StartTime := sL_StartTime + ' 00:00:00';

                sL_EndTime := TUstr.replaceStr(DateTimeToStr(FieldByName('EndTime').AsDateTime),'/','-');
                if Length(sL_EndTime) <= 10 then
                  sL_EndTime := sL_EndTime + ' 00:00:00';

                {
                TMyXML.setAttribute(L_DetailChildNode, 'DateTime',sL_RealDateTime);
                TMyXML.setAttribute(L_DetailChildNode, 'Status',sL_Status);
                TMyXML.setAttribute(L_DetailChildNode, 'PurchaseID',sL_PurchaseID);
                TMyXML.setAttribute(L_DetailChildNode, 'ICCardNo',sL_ICCardNo);
                TMyXML.setAttribute(L_DetailChildNode, 'ICCardID',sL_ICCardID);
                TMyXML.setAttribute(L_DetailChildNode, 'CasID',sL_CasID);
                TMyXML.setAttribute(L_DetailChildNode, 'ProductCode',sL_ProductCode);
                TMyXML.setAttribute(L_DetailChildNode, 'StartTime',sL_StartTime);
                TMyXML.setAttribute(L_DetailChildNode, 'EndTime',sL_EndTime);
                }
                sL_DetailXmlStr := '<' + sL_DetailChildNode + ' DateTime="' + sL_RealDateTime +
                                   '" Status="' + sL_Status + '" PurchaseID="' + sL_PurchaseID +
                                   '" ICCardNo="' + sL_ICCardNo + '" ICCardID="' + sL_ICCardID +
                                   '" CASID="' + sL_CasID + '" ProductCode="' + sL_ProductCode +
                                   '" StartTime="' + sL_StartTime + '" EndTime="' + sL_EndTime +
                                   '" /> ';


                if  bG_ActiveCombineData then
                  Writeln(G_BaseXmlFilePath,sL_DetailXmlStr);
                if not bI_InitBaseXmlFile then
                    Writeln(F_FilePath,sL_DetailXmlStr);

                Next;
              end;
              dtmMain.adoQryQueryData.EnableControls;
            end
//**********************   授權資訊
            else if UpperCase(sI_QueryType) = UpperCase(ENTITLEMENT_QUERY) then
            begin
              sL_MasterXmlStr := sL_MasterXmlStr + '">';
              if  bI_InitBaseXmlFile and bG_ActiveCombineData then
                Writeln(G_BaseXmlFilePath,sL_MasterXmlStr)
              else
                Writeln(F_FilePath,sL_MasterXmlStr);


              sL_RealXmlStr := '';

              dtmMain.adoQryQueryData.DisableControls;
              dtmMain.adoQryQueryData.First;
              while not dtmMain.adoQryQueryData.Eof do
              begin
                //L_DetailChildNode := TMyXML.addChild(L_MasterChildNode, sL_DetailChildNode,'');
                sL_RealDateTime := TUstr.replaceStr(DateTimeToStr(FieldByName('RealDateTime').AsDateTime),'/','-');
                if Length(sL_RealDateTime) <= 10 then
                  sL_RealDateTime := sL_RealDateTime + ' 00:00:00';

                sL_Status := FieldByName('Status').AsString;
                if sL_Status = '0' then
                  sL_Status := 'Create'
                else if sL_Status = '1' then
                  sL_Status := 'Delete';

                sL_ICCardNo := FieldByName('ICCardNo').AsString;
                sL_ICCardID := FieldByName('ICCardID').AsString;
                sL_CasID := FieldByName('CasID').AsString;
                sL_ProductCode := FieldByName('ProductCode').AsString;
                {
                TMyXML.setAttribute(L_DetailChildNode, 'DateTime',sL_RealDateTime);
                TMyXML.setAttribute(L_DetailChildNode, 'ICCardNo',sL_ICCardNo);
                TMyXML.setAttribute(L_DetailChildNode, 'ICCardID',sL_ICCardID);
                TMyXML.setAttribute(L_DetailChildNode, 'Status',sL_Status);
                TMyXML.setAttribute(L_DetailChildNode, 'ProductCode',sL_ProductCode);
                TMyXML.setAttribute(L_DetailChildNode, 'CasID',sL_CasID);
                }

                sL_DetailXmlStr := '<' + sL_DetailChildNode + ' DateTime="' + sL_RealDateTime +
                                   '" ICCardNo="' + sL_ICCardNo + '" ICCardID="' + sL_ICCardID +
                                   '" Status="' + sL_Status + '" ProductCode="' + sL_ProductCode +
                                   '" CASID="' + sL_CasID + '" /> ';

                if  bG_ActiveCombineData then
                  Writeln(G_BaseXmlFilePath,sL_DetailXmlStr);
                if not bI_InitBaseXmlFile then
                  Writeln(F_FilePath,sL_DetailXmlStr);

                Next;
              end;
              dtmMain.adoQryQueryData.EnableControls;
            end;
          end;
        end;
      end
      else
      begin
        if UpperCase(sI_QueryType) = UpperCase(IC_CARD_QUERY) then
        begin
          //TMyXML.setAttribute(L_MasterChildNode, 'TotalICCardNum','');
          //TMyXML.setAttribute(L_MasterChildNode, 'TotalSubNum','');
          sL_MasterXmlStr := sL_MasterXmlStr + '" TotalICCardNum="" TotalSubNum="">';
        end
        else if UpperCase(sI_QueryType) = UpperCase(CA_PRODUCT_QUERY) then
        begin
          //TMyXML.setAttribute(L_MasterChildNode, 'TotalProductNum','');
          //TMyXML.setAttribute(L_MasterChildNode, 'TotalChannelNum','');
          sL_MasterXmlStr := sL_MasterXmlStr + '" TotalProductNum="" TotalChannelNum="">';
        end
        else
        begin
          sL_MasterXmlStr := sL_MasterXmlStr + '">';
        end;
        if  bI_InitBaseXmlFile and bG_ActiveCombineData then
          Writeln(G_BaseXmlFilePath,sL_MasterXmlStr);        
        if not bI_InitBaseXmlFile then
          Writeln(F_FilePath,sL_MasterXmlStr);
        //L_DetailChildNode := TMyXML.addChild(L_MasterChildNode, sL_DetailChildNode,'');
      end;
    end
    else
    begin
      if UpperCase(sI_QueryType) = UpperCase(IC_CARD_QUERY) then
      begin
        //TMyXML.setAttribute(L_MasterChildNode, 'TotalICCardNum','');
        //TMyXML.setAttribute(L_MasterChildNode, 'TotalSubNum','');
        sL_MasterXmlStr := sL_MasterXmlStr + '" TotalICCardNum="" TotalSubNum="">';
      end
      else if UpperCase(sI_QueryType) = UpperCase(CA_PRODUCT_QUERY) then
      begin
        //TMyXML.setAttribute(L_MasterChildNode, 'TotalProductNum','');
        //TMyXML.setAttribute(L_MasterChildNode, 'TotalChannelNum','');
        sL_MasterXmlStr := sL_MasterXmlStr + '" TotalProductNum="" TotalChannelNum="">';
      end;
      begin
        sL_MasterXmlStr := sL_MasterXmlStr + '">';
      end;
      //L_DetailChildNode := TMyXML.addChild(L_MasterChildNode, sL_DetailChildNode,'');
      if  bI_InitBaseXmlFile and bG_ActiveCombineData then
        Writeln(G_BaseXmlFilePath,sL_MasterXmlStr);
      if not bI_InitBaseXmlFile then
        Writeln(F_FilePath,sL_MasterXmlStr);

      sL_RetCode := sI_RetCode;
      sL_Msg := sI_Msg;
    end;

//******************************************************************************
//******************************************************************************
{
    //存成檔案
    sL_XmlFilePath := dtmMain.getRealFilePath(sI_CompCode,sI_InstQuery,sI_QueryType,sI_HandleEDateTime,false);
    sL_TmpXmlFilePath := ExtractFilePath(sL_XmlFilePath)  +  TUdateTime.GetPureDateStr(now,'') +'.xml';

    if (sI_InstQuery = '1') or
       ((sI_InstQuery = '0') and ((sI_Msg='') and (nI_TotalExecCounts >= 0)) OR ((sI_Msg=' ') and (nI_TotalExecCounts >= 0))) then
    begin
      if FileExists(sL_XmlFilePath) then
        DeleteFile(sL_XmlFilePath);

      setXMLDocument.SaveToFile(sL_XmlFilePath);
    end;

    sL_XmlString := setXMLDocument.XML.GetText;
    setXMLDocument.Active := false;
}


    sL_MasterXmlStr := '</' + sL_MasterChildNode + '>';


    if not bI_InitBaseXmlFile then
    begin
      Writeln(F_FilePath,sL_MasterXmlStr);

      Writeln(F_FilePath,'</Msg>');
      Flush(F_FilePath);
      CloseFile(F_FilePath);
    end;
    sL_XmlString := '';
    //取出該Xml字串,若大於 1 MB 則不用
    if not bI_InitBaseXmlFile then
    begin
      dL_FileSize := dtmMain.getFileSize(sL_XmlFilePath);
      if dL_FileSize <= MAX_FILE_SIZE then   //小於 1 MB 則取出字串
      begin
         L_XmlTxt := TStringList.Create;
         L_XmlTxt.LoadFromFile(sL_XmlFilePath);
         sL_XmlString := L_XmlTxt.GetText;
         L_XmlTxt.Free;
         //sL_XmlString := TUstr.replaceStr(sL_XmlString,''#$D#$A'','');
         //dtmMain.savetofile(sL_XmlString,'c:\KKK.txt');
      end;
    end;
    Result := sL_XmlString;//開啟xml再回傳xml
end;



function TfrmMain.returnZipXML(sI_QueryType,sI_XmlFilePath,sI_ResponXml, sI_CompCode,
  sI_Version, sI_QueryDateTime, sI_SrcCode, sI_SrcType, sI_SrcMsgID,
  sI_DstCode, sI_DstType, sI_DstMsgID, sI_CasID, sI_InstQuery,
  sI_FirstDate, sI_HandleEDateTime, sI_RetCode, sI_Msg,sI_TotalICCardNum,sI_TotalSubNum,
  sI_TotalProductNum,sI_TotalChannelNum: String;
  nI_TotalExecCounts: Integer): WideString;
var
    L_MsgRootNode,L_ReturnChildNode,L_MasterChildNode,L_DetailChildNode,L_SubDetailChildNode  : IXmlNode;
    sL_Result,sL_RetCode,sL_Msg,sL_Redirect,sL_SrcIp,sL_LocalIisPort, sL_HandleDate : String;
    nL_TotalICCardNum,nL_TotalSubNum,nL_TotalProductNum,nL_TotalChannelNum : Integer;
    sL_MasterChildNode,sL_DetailChildNode,sL_SubDetailChildNode,sL_IISName, sL_ReturnType : String;
begin
    setXMLDocument.Active := true;

    if UpperCase(dtmMAIN.sG_TcOrSc) = 'TC' then
      setXMLDocument.Encoding := XML_ENCODING_TC
    else
      setXMLDocument.Encoding := XML_ENCODING_SC;

    setXMLDocument.StandAlone := XML_STANDALONE;

    L_MsgRootNode := TMyXML.createRootNode(setXMLDocument,'Msg','', sL_Result);
    TMyXML.setAttribute(L_MsgRootNode, 'Version',sI_Version);
    TMyXML.setAttribute(L_MsgRootNode, 'Type',sI_SrcType);
    TMyXML.setAttribute(L_MsgRootNode, 'DateTime',sI_QueryDateTime);
    TMyXML.setAttribute(L_MsgRootNode, 'SrcCode',sI_SrcCode);

    if sI_InstQuery = '0' then
    begin
{
      TMyXML.setAttribute(L_MsgRootNode, 'MsgID','');
      TMyXML.setAttribute(L_MsgRootNode, 'DstCode','');
      TMyXML.setAttribute(L_MsgRootNode, 'ReplyID','');
}
      TMyXML.setAttribute(L_MsgRootNode, 'MsgID',sI_SrcMsgID);  //Howard 93/07/06
      TMyXML.setAttribute(L_MsgRootNode, 'DstCode',sI_DstCode); //Howard 93/07/06
      TMyXML.setAttribute(L_MsgRootNode, 'ReplyID',sI_DstMsgID);//Howard 93/07/06
      sL_IISName := IIS_HISTORY;
    end
    else
    begin
      TMyXML.setAttribute(L_MsgRootNode, 'MsgID',sI_SrcMsgID);
      TMyXML.setAttribute(L_MsgRootNode, 'DstCode',sI_DstCode);
      TMyXML.setAttribute(L_MsgRootNode, 'ReplyID',sI_DstMsgID);
      sL_IISName := IIS_IMMEDIATELY;
    end;


    L_ReturnChildNode := TMyXML.addChild(L_MsgRootNode, 'Return','');



    //壓縮檔案一定小於 1MB
    TMyXML.setAttribute(L_ReturnChildNode, 'Value','0');    //成功
    TMyXML.setAttribute(L_ReturnChildNode, 'Desc',dtmMain.getLanguageSettingInfo('Return_Msg_0_Content'));

    sL_SrcIp := dtmMain.getSrcIP(sI_CompCode);
    sL_LocalIisPort := dtmMain.sG_LocalIisPort;
    if sL_LocalIisPort<>'' then
      sL_SrcIp := sL_SrcIp + ':' + sL_LocalIisPort;

    sL_Redirect := 'http://' + sL_SrcIp + '/' + sL_IISName + '/' + dtmMain.getFileName(sI_HandleEDateTime,sI_QueryType,true);
    TMyXML.setAttribute(L_ReturnChildNode, 'Redirect',sL_Redirect);


//******************************************************************************
//******************************************************************************
    if UpperCase(sI_QueryType) = UpperCase(IC_CARD_QUERY) then  //1:IC卡資訊
    begin
      sL_MasterChildNode := 'ICCardQueryReport';
      sL_DetailChildNode := 'ICCard';
      sL_SubDetailChildNode := '';
      sL_ReturnType := 'ICCardQuery';
    end
    else if UpperCase(sI_QueryType) = UpperCase(CA_PRODUCT_QUERY) then   //2:產品定義資訊
    begin
      sL_MasterChildNode := 'CAProductQueryReport';
      sL_DetailChildNode := 'CAProduct';
      sL_SubDetailChildNode := 'Channel';
      sL_ReturnType := 'CAProductQuery';
    end
    else if UpperCase(sI_QueryType) = UpperCase(PROD_PURCHASE_QUERY) then  //3:產品定購資訊
    begin
      sL_MasterChildNode := 'ProdPurchaseQueryReport';
      sL_DetailChildNode := 'ProdPurchase';
      sL_SubDetailChildNode := '';
      sL_ReturnType := 'ProdPurchaseQuery';
    end
    else if UpperCase(sI_QueryType) = UpperCase(ENTITLEMENT_QUERY) then  //4:授權資訊
    begin
      sL_MasterChildNode := 'EntitlementQueryReport';
      sL_DetailChildNode := 'Entitlement';
      sL_SubDetailChildNode := '';
      sL_ReturnType := 'EntitlementQuery';
    end;

    TMyXML.setAttribute(L_ReturnChildNode, 'Type',sL_ReturnType);


    L_MasterChildNode := TMyXML.addChild(L_MsgRootNode, sL_MasterChildNode,'');
    TMyXML.setAttribute(L_MasterChildNode, 'InstQuery',sI_InstQuery);

    sL_HandleDate := TUstr.replaceStr(Copy(sI_HandleEDateTime,0,10),'/','-');
    TMyXML.setAttribute(L_MasterChildNode, 'Date',sL_HandleDate);

    
    if UpperCase(sI_QueryType) = UpperCase(IC_CARD_QUERY) then  //1:IC卡資訊
    begin
      {
      dtmMain.getICCardNumAndSubNum(false,sI_CompCode,sI_HandleEDateTime,nL_TotalICCardNum,nL_TotalSubNum);
      TMyXML.setAttribute(L_MasterChildNode, 'TotalICCardNum',IntToStr(nL_TotalICCardNum));
      TMyXML.setAttribute(L_MasterChildNode, 'TotalSubNum',IntToStr(nL_TotalSubNum));
      }
      TMyXML.setAttribute(L_MasterChildNode, 'TotalICCardNum',sI_TotalICCardNum);
      TMyXML.setAttribute(L_MasterChildNode, 'TotalSubNum',sI_TotalSubNum);
    end
    else if UpperCase(sI_QueryType) = UpperCase(CA_PRODUCT_QUERY) then   //2:產品定義資訊
    begin
      {
      dtmMain.getProductNumAndChannelNum(sI_CompCode,sI_HandleEDateTime,nL_TotalProductNum,nL_TotalChannelNum);
      TMyXML.setAttribute(L_MasterChildNode, 'TotalProductNum',IntToStr(nL_TotalProductNum));
      TMyXML.setAttribute(L_MasterChildNode, 'TotalChannelNum',IntToStr(nL_TotalChannelNum));
      }
      TMyXML.setAttribute(L_MasterChildNode, 'TotalProductNum',sI_TotalProductNum);
      TMyXML.setAttribute(L_MasterChildNode, 'TotalChannelNum',sI_TotalChannelNum);
    end;

    L_DetailChildNode := TMyXML.addChild(L_MasterChildNode, sL_DetailChildNode,'');
    if UpperCase(sI_QueryType) = UpperCase(CA_PRODUCT_QUERY) then   //2:產品定義資訊
      L_SubDetailChildNode := TMyXML.addChild(L_DetailChildNode, sL_SubDetailChildNode,'');

    Result := setXMLDocument.XML.GetText;
    setXMLDocument.Active := false;
end;



function TfrmMain.returnNightRunQueryXML(sI_QueryType,
  sI_XmlFilePath, sI_ResponXml, sI_CompCode, sI_Version, sI_QueryDateTime,
  sI_SrcCode, sI_SrcType, sI_SrcMsgID, sI_DstCode, sI_DstType, sI_DstMsgID,
  sI_CasID, sI_InstQuery, sI_FirstDate, sI_HandleEDateTime, sI_RetCode,
  sI_Msg: String; nI_TotalExecCounts: Integer): WideString;
var
    L_NodeList : IXmlNodeList;
begin
    getXMLDocument.LoadFromFile(sI_XmlFilePath);

    //取得MSG屬性
    L_NodeList := getXMLDocument.ChildNodes;
    TMyXML.setAttribute(L_NodeList.Nodes[1],'MsgID',sI_SrcMsgID);
    TMyXML.setAttribute(L_NodeList.Nodes[1],'DateTime',sI_QueryDateTime);
    TMyXML.setAttribute(L_NodeList.Nodes[1],'ReplyID',sI_DstMsgID);

    Result := getXMLDocument.XML.GetText;
end;

procedure TfrmMain.createHttpThread;
begin
    G_HttpThread := THttpThread.Create;
    if GetExitCodeThread(G_HttpThread.ThreadID,G_HttpThread_ExitCode) then
    begin
      MessageDlg(dtmMain.getLanguageSettingInfo('frmMain_Msg_11_Content') , mtError, [mbOK],0);//執行緒 1 產生失敗!請重新執行程式

      Application.Terminate;
      Exit;
    end;
end;

procedure TfrmMain.sendResponseXml(sI_CompCode,sI_ResponseXml: String);
var
    L_Data : TIdMultiPartFormDataStream;
    L_Response: TStringStream;
    L_XmlString: TStrings;
    sL_ResponseXml,sL_DstIP,sL_HttpTitle : String;
    F_FilePath : TextFile;
    nL_FileHandel : Integer;
begin
    if not FileExists('c:\ResponseXml.txt') then
    begin
      nL_FileHandel := FileCreate('c:\ResponseXml.txt');

      if (nL_FileHandel>=0) then
        FileClose(nL_FileHandel);
    end;
    AssignFile(F_FilePath,'c:\ResponseXml.txt');
    Append(F_FilePath);


    L_Response := TStringStream.Create('');
    L_XmlString := TStringList.Create;

    sL_ResponseXml := TUstr.replaceStr(sI_ResponseXml,''#$D#$A'','');

    L_XmlString.Add(sL_ResponseXml);

    L_Data := TIdMultiPartFormDataStream.Create;
    L_Data.AddFormField('', sL_ResponseXml);

    sL_DstIP := dtmMain.getDstIP(sI_CompCode);
    if bG_LongFormat then
      sL_HttpTitle := 'http://' + sL_DstIP + '/?'
    else
      sL_HttpTitle := 'http://' + sL_DstIP;

    //將要回傳指令儲存
    Writeln(F_FilePath,'[' +  DateTimeToStr(now)+ ']-' + sL_HttpTitle + sL_ResponseXml);


    HTTP.ProtocolVersion := pv1_1;
    HTTP.Request.Username := '';
    HTTP.Request.Password := '';
//    HTTP.Request.ProxyServer := '';
//    HTTP.Request.ProxyPort := dtmMain.nG_DstHttpPort;

//    HTTP.Request.ContentType := 'application/octet-stream';
//    HTTP.Request.ContentType := 'text/html';
    HTTP.Request.ContentType := 'application/x-www-form-urlencoded';
//    HTTP.Request.ContentType := '';
    
    HTTP.HandleRedirects := true;

    try

//      HTTP.Post(sL_HttpTitle,L_XmlString, L_Response);
      HTTP.Post(sL_HttpTitle,L_XmlString);
//        HTTP.Post(sL_HttpTitle,L_Data);
//      showLogMsg('[OK]-' + HTTP.ResponseText + '-' + sL_ResponseXml);
      showLogMsg('[OK]-' + HTTP.ResponseText);
      Writeln(F_FilePath,'[' +  DateTimeToStr(now)+ ']OK -' + HTTP.ResponseText);
    except
      //showLogMsg('[ERROR]-' + HTTP.ResponseText);
      Writeln(F_FilePath,'[' +  DateTimeToStr(now)+ ']ERROR -' + HTTP.ResponseText);
    end;
    L_Data.Free;


{
//***********************************************************
//JACKALTEST
  NMHTTP1.InputFileMode := FALSE;
  NMHTTP1.OutputFileMode := FALSE;
  NMHTTP1.ReportLevel := Status_Basic;
  NMHTTP1.Port := dtmMain.nG_DstHttpPort;

  With NMHTTP1.HeaderInfo do
  Begin
    Cookie := '';
    LocalMailAddress := '';
    LocalProgram := '';
    Referer := '';
    UserID := '';
    Password := '';
  End;

  //sL_HttpTitle := 'http://www.pchome.com.tw';
  //sL_ResponseXml := 'PPP';
  try
    //showLogMsg('[POST]');

    NMHTTP1.Post(sL_HttpTitle,sL_ResponseXml);
    Writeln(F_FilePath,'[' +  DateTimeToStr(now)+ ']OK -' + NMHTTP1.Header);
//    showLogMsg('[OK]--' + NMHTTP1.Header);
    showLogMsg('xml=>' + sL_ResponseXml);
  except
    //showLogMsg('[ERRORQ]');
    //showLogMsg('[ERRORQ]-' + NMHTTP1.ReadLn);
    Writeln(F_FilePath,'[' +  DateTimeToStr(now)+ ']ERROR');
    //Writeln(F_FilePath,'[' +  DateTimeToStr(now)+ ']ERROR -' + NMHTTP1.Header);
  end;

  //Memo1.Text := NMHTTP1.Header;
}  
//***********************************************************


    Flush(F_FilePath);
    CloseFile(F_FilePath);

    L_Response.Free;
    L_XmlString.Free;
end;

{
procedure TfrmMain.sendResponseXml(sI_CompCode,sI_ResponseXml: String);
var
    L_Response: TStringStream;
    L_XmlString: TStrings;
    sL_ResponseXml,sL_DstIP,sL_HttpTitle : String;
    F_FilePath : TextFile;
    nL_FileHandel : Integer;
begin
    if not FileExists('c:\ResponseXml.txt') then
    begin
      nL_FileHandel := FileCreate('c:\ResponseXml.txt');

      if (nL_FileHandel>=0) then
        FileClose(nL_FileHandel);
    end;
    AssignFile(F_FilePath,'c:\ResponseXml.txt');
    Append(F_FilePath);


    L_Response := TStringStream.Create('');
    L_XmlString := TStringList.Create;

    sL_ResponseXml := TUstr.replaceStr(sI_ResponseXml,''#$D#$A'','');

    L_XmlString.Add(sL_ResponseXml);

    sL_DstIP := dtmMain.getDstIP(sI_CompCode);
    if bG_LongFormat then
      sL_HttpTitle := 'http://' + sL_DstIP + '/?'
    else
      sL_HttpTitle := 'http://' + sL_DstIP;

    //將要回傳指令儲存
    Writeln(F_FilePath,'[' +  DateTimeToStr(now)+ ']-' + sL_HttpTitle + sL_ResponseXml);


//***********************************************************
//JACKALTEST
  NMHTTP1.InputFileMode := FALSE;
  NMHTTP1.OutputFileMode := FALSE;
  NMHTTP1.ReportLevel := Status_Basic;
  NMHTTP1.Port := dtmMain.nG_DstHttpPort;

  With NMHTTP1.HeaderInfo do
  Begin
    Cookie := '';
    LocalMailAddress := '';
    LocalProgram := '';
    Referer := '';
    UserID := '';
    Password := '';
  End;

  //sL_HttpTitle := 'http://www.pchome.com.tw';
  //sL_ResponseXml := 'PPP';
  try
    //showLogMsg('[POST]');

    NMHTTP1.Post(sL_HttpTitle,sL_ResponseXml);
    Writeln(F_FilePath,'[' +  DateTimeToStr(now)+ ']OK -' + NMHTTP1.Header);
//    showLogMsg('[OK]--' + NMHTTP1.Header);
    showLogMsg('xml=>' + sL_ResponseXml);
  except
    //showLogMsg('[ERRORQ]');
    //showLogMsg('[ERRORQ]-' + NMHTTP1.ReadLn);
    Writeln(F_FilePath,'[' +  DateTimeToStr(now)+ ']ERROR');
    //Writeln(F_FilePath,'[' +  DateTimeToStr(now)+ ']ERROR -' + NMHTTP1.Header);
  end;

  //Memo1.Text := NMHTTP1.Header;
//***********************************************************


    Flush(F_FilePath);
    CloseFile(F_FilePath);

    L_Response.Free;
    L_XmlString.Free;
end;

}


procedure TfrmMain.chkNightRunClick(Sender: TObject);
begin
    if chkNightRun.Checked then
    begin
      bG_ActionNightRun := true;
      Timer1.Enabled := false;
      StartNightRunProcess(true,'','','','', false); //自動執行NightRun,則後面參數省略
      Timer1.Enabled := true;
    end
    else
    begin
      bG_ActionNightRun := false;
      Timer1.Enabled := true;
    end;
end;

procedure TfrmMain.saveToZipFile(SI_WinZipExePath, sI_XmlFilePath,
  sI_ZipFilePath,sI_OriginalErrCode,sI_OriginalErrMsg : String;var sI_ErrCode,sI_ErrMsg : String);
var
    bL_Status : Boolean;
    sL_ApFilePath,sL_Param,sL_XmlPath,sL_ZipPath,sL_ErrMsg : String;
begin
    //sL_ApFilePath := 'D:\CKD\Tools\WinZip81\WINZIP32.EXE';
    //sL_ZipPath := 'D:\App\SGS\Project\Bin\即時資料\20040319_SMS_ICCardQueryR.zip';
    //sL_XmlPath := 'D:\App\SGS\Project\Bin\即時資料\20040319_SMS_ICCardQueryR.xml';
    sL_Param := '-min -a -r ' + sI_ZipFilePath + ' ' + sI_XmlFilePath;

    if FileExists(sI_ZipFilePath) then
      DeleteFile(sI_ZipFilePath);

    bL_Status := TUother.RunShell(SI_WinZipExePath,sL_Param,false,true, sL_ErrMsg);

    if bL_Status then
    begin
      sI_ErrCode := sI_OriginalErrCode;
      sI_ErrMsg := sI_OriginalErrMsg;
    end
    else
    begin
      sI_ErrCode := '106';
      sI_ErrMsg := dtmMain.getLanguageSettingInfo('Return_Msg_106_Content');
    end;
end;

procedure TfrmMain.NightRunLog1Click(Sender: TObject);
begin
    //if not bG_ReadOnly then
    //begin
      frmNightRunLog := TfrmNightRunLog.Create(Application);
      frmNightRunLog.ShowModal;
      frmNightRunLog.Free;
    //end
    //else
    //  MessageDlg(dtmMain.getLanguageSettingInfo('frmMain_Msg_10_Content'),mtError, [mbOK],0);
end;

procedure TfrmMain.Log21Click(Sender: TObject);
begin
    //if not bG_ReadOnly then
    //begin
      frmQueryLog := TfrmQueryLog.Create(Application);
      frmQueryLog.ShowModal;
      frmQueryLog.Free;
    //end
    //else
    //  MessageDlg(dtmMain.getLanguageSettingInfo('frmMain_Msg_10_Content'),mtError, [mbOK],0);
end;

procedure TfrmMain.Timer2Timer(Sender: TObject);
begin

    bG_ColorFlag := not bG_ColorFlag ;
    if bG_ColorFlag then
    begin
       pnlTitle.Font.Color := clBackground;
       pnlTitle.Refresh;
    end
    else
    begin
       pnlTitle.Font.Color :=clRed;
       pnlTitle.Refresh;
    end;
    
end;

function TfrmMain.writeRegData(sI_EncryptedLicenseKey: STring): boolean;
var
  L_StrList : TStringList;
  sL_LicenseKey : String;
  bL_WriteRegResult : boolean;
begin
  bL_WriteRegResult := true;
  sL_LicenseKey := TUother.myDecrypt(sI_EncryptedLicenseKey, ENCRYPTION_KEY);

  L_StrList := TUstr.ParseStrings(sL_LicenseKey,ENCRYPTION_SEP);
  if (L_StrList.Count=3) and (TUdateTime.IsDateStr(L_StrList.Strings[0],'/'))
    and ((L_StrList.Strings[1]=SYS4_Y) or (L_StrList.Strings[1]=SYS4_N))
    and ((L_StrList.Strings[2]=SYS5_Y) or (L_StrList.Strings[2]=SYS5_N)) then
  begin
    //表示 License Key 格式正確...
    TUother.RegWrite(REGISTRY_ROOT,'SYS1',3,StrToDate(L_StrList.Strings[0])); //寫入軟體之使用到期日
    TUother.RegWrite(REGISTRY_ROOT,'SYS2',3,Date); //寫入軟體之安裝日
    TUother.RegWrite(REGISTRY_ROOT,'SYS3',3,Date); //寫入軟體之上次使用日期
    TUother.RegWrite(REGISTRY_ROOT,'SYS4',2,L_StrList.Strings[1]); //寫入是否啟動軟體保護機制
    TUother.RegWrite(REGISTRY_ROOT,'SYS5',2,L_StrList.Strings[2]); //寫入是否停用本軟體

   end
   else
     bL_WriteRegResult := false;
   L_StrList.Free;
   result := bL_WriteRegResult;

end;

function TfrmMain.GetRegistryInfo: boolean;
var
    L_StrList : TStringList;
    sL_EncryptedLicenseKey, sL_LicenseKey : STring;

    bL_Result : boolean;
    vL_Result : variant;


    procedure readRegData;
    begin
      vL_Result := TUother.RegRead(REGISTRY_ROOT,'SYS1',3); //軟體之使用到期日
        dG_Sys1 := StrToDate(VarToStr(vL_Result));

      vL_Result := TUother.RegRead(REGISTRY_ROOT,'SYS2',3); //軟體之安裝日
        dG_Sys2 := StrToDate(VarToStr(vL_Result));

      vL_Result := TUother.RegRead(REGISTRY_ROOT,'SYS3',3); //軟體之上次使用日期
        dG_Sys3 := StrToDate(VarToStr(vL_Result));

      vL_Result := TUother.RegRead(REGISTRY_ROOT,'SYS4',2); //是否啟動軟體保護機制
        sG_Sys4 := VarToStr(vL_Result);

      vL_Result := TUother.RegRead(REGISTRY_ROOT,'SYS5',2); //是否停用本軟體
        sG_Sys5 := VarToStr(vL_Result);


    end;
begin
    try
      bL_Result := true;
      readRegData;
    except
      sL_EncryptedLicenseKey := InputBox(dtmMain.getLanguageSettingInfo('frmMain_Msg_2_Content'), 'License Key:', '');
      if sL_EncryptedLicenseKey='' then
      begin
        bL_Result := false;
      end  
      else if writeRegData(sL_EncryptedLicenseKey) then
      begin
        try
          readRegData;
        except
          bL_Result := false;
        end;
      end
      else
        bL_Result := false;
    end;

    result := bL_Result;


end;

procedure TfrmMain.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
    if (ssCtrl in Shift) and (Key=80) then //Ctrl + P
    begin
      if not bG_ReadOnly then
      begin
        frmLoadData := TfrmLoadData.Create(Application);
        frmLoadData.ShowModal;
        frmLoadData.Free;
      end
      else
        MessageDlg(dtmMain.getLanguageSettingInfo('frmMain_Msg_10_Content'),mtError, [mbOK],0);
    end;

end;

procedure TfrmMain.setDateTimeFormat;
begin
    DateSeparator := '/';
    TimeSeparator := ':';

    LongDateFormat := 'yyyy/mm/dd';
    ShortDateFormat := 'yyyy/mm/dd';
    ShortTimeFormat :='hh:nn:ss';
    LongTimeFormat :='hh:nn:ss';
    TimeAMString := '';
    TimePMString := '';
end;

procedure TfrmMain.BitBtn1Click(Sender: TObject);
var
    L_MsgNodeList : IDOMNodeList;
    sI_TmpXmlFilePath,sL_NodeName,sL_TotalICCardNum,sL_TotalSubNum : String;
    ii,jj : Integer;
begin
    sI_TmpXmlFilePath := 'D:\app\SGS\Project\Bin\歷史資料\20030818_SMS_ICCardQueryR.xml';
    getXMLDocument.LoadFromFile(sI_TmpXmlFilePath);

    //L_MsgNodeList := getXMLDocument.DOMDocument.getElementsByTagName('ICCardQueryReport');
    L_MsgNodeList := getXMLDocument.DOMDocument.getElementsByTagName('Msg');

    for ii:=0 to L_MsgNodeList.length-1 do
    begin
      for jj:=0 to L_MsgNodeList.item[ii].childNodes.length-1 do
      begin
        sL_NodeName := L_MsgNodeList.item[ii].childNodes.item[jj].nodeName;
        if sL_NodeName = 'ICCardQueryReport' then
        begin
          sL_TotalICCardNum := TMyXML.getAttributeValue(L_MsgNodeList.item[ii].childNodes.item[jj],'TotalICCardNum');
          sL_TotalSubNum := TMyXML.getAttributeValue(L_MsgNodeList.item[ii].childNodes.item[jj],'TotalSubNum');
          showmessage('TotalICCardNum=  ' + sL_TotalICCardNum + 'TotalSubNum=  ' + sL_TotalSubNum);
        end;
      end;
    end;

end;

function TfrmMain.getXmlICCardNumAndSubNum(sI_XmlFilePath: String;
  var sI_TotalICCardNum, sI_TotalSubNum: String): String;
var
    L_MsgNodeList : IDOMNodeList;
    sL_NodeName,sL_TotalICCardNum,sL_TotalSubNum : String;
    ii,jj : Integer;
begin
    //sI_XmlFilePath := 'D:\app\SGS\Project\Bin\歷史資料\20030818_SMS_ICCardQueryR.xml';
    getXMLDocument.LoadFromFile(sI_XmlFilePath);

    L_MsgNodeList := getXMLDocument.DOMDocument.getElementsByTagName('Msg');

    for ii:=0 to L_MsgNodeList.length-1 do
    begin
      for jj:=0 to L_MsgNodeList.item[ii].childNodes.length-1 do
      begin
        sL_NodeName := L_MsgNodeList.item[ii].childNodes.item[jj].nodeName;
        if sL_NodeName = 'ICCardQueryReport' then
        begin
          sI_TotalICCardNum := TMyXML.getAttributeValue(L_MsgNodeList.item[ii].childNodes.item[jj],'TotalICCardNum');
          sI_TotalSubNum := TMyXML.getAttributeValue(L_MsgNodeList.item[ii].childNodes.item[jj],'TotalSubNum');
          //showmessage('TotalICCardNum=  ' + sI_TotalICCardNum + '  TotalSubNum=  ' + sI_TotalSubNum);
        end;
      end;
    end;

end;

function TfrmMain.getXmlProductNumAndChannelNum(sI_XmlFilePath: String;
  var sI_TotalProductNum, sI_TotalChannelNum: String): String;
var
    L_MsgNodeList : IDOMNodeList;
    sL_NodeName,sL_TotalICCardNum,sL_TotalSubNum : String;
    ii,jj : Integer;
begin
    //sI_XmlFilePath := 'D:\app\SGS\Project\Bin\歷史資料\20030818_SMS_CAProductQueryR.xml';
    getXMLDocument.LoadFromFile(sI_XmlFilePath);

    L_MsgNodeList := getXMLDocument.DOMDocument.getElementsByTagName('Msg');

    for ii:=0 to L_MsgNodeList.length-1 do
    begin
      for jj:=0 to L_MsgNodeList.item[ii].childNodes.length-1 do
      begin
        sL_NodeName := L_MsgNodeList.item[ii].childNodes.item[jj].nodeName;
        if sL_NodeName = 'CAProductQueryReport' then
        begin
          sI_TotalProductNum := TMyXML.getAttributeValue(L_MsgNodeList.item[ii].childNodes.item[jj],'TotalProductNum');
          sI_TotalChannelNum := TMyXML.getAttributeValue(L_MsgNodeList.item[ii].childNodes.item[jj],'TotalChannelNum');
          //showmessage('TotalProductNum=  ' + sI_TotalProductNum + '  TotalChannelNum=  ' + sI_TotalChannelNum);
        end;
      end;
    end;


end;

function TfrmMain.getSO459ICCardNumAndSubNum(sI_CompCode, sI_ModeType,
  sI_HandleEDateTime: String; var sI_TotalICCardNum,
  sI_TotalSubNum: String): String;
var
    sL_SQL : String;
    L_AdoQry : TADOQuery;
begin
    sL_SQL := 'select TotalICCardNum,TotalSubNum from SO459 where CompCode=' + sI_CompCode +
              ' and ModeType=' + sI_ModeType + ' and HandleEDateTime=' +
              TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_HandleEDateTime)) +
              'and ((ErrorMsg is null) OR (ErrorMsg='' ''))' +
              ' Order by execdatetime desc';

{

}
    L_AdoQry := dtmMain.runSQL(SELECT_MODE,sI_CompCode,sL_SQL);

    if L_AdoQry <> nil then
    begin
      sI_TotalICCardNum := L_AdoQry.FieldByName('TotalICCardNum').AsString;
      sI_TotalSubNum := L_AdoQry.FieldByName('TotalSubNum').AsString;
    end;
end;

function TfrmMain.getSO459ProductNumAndChannelNum(sI_CompCode, sI_ModeType,
  sI_HandleEDateTime: String; var sI_TotalProductNum,
  sI_TotalChannelNum: String): String;
var
    sL_SQL : String;
    L_AdoQry : TADOQuery;
begin
    sL_SQL := 'select TotalProductNum,TotalChannelNum from SO459 where CompCode=' + sI_CompCode +
              ' and ModeType=' + sI_ModeType + ' and HandleEDateTime=' +
              TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_HandleEDateTime)) +
              'and ((ErrorMsg is null) OR (ErrorMsg='' ''))' +
              ' Order by execdatetime desc';

    L_AdoQry := dtmMain.runSQL(SELECT_MODE,sI_CompCode,sL_SQL);

    if L_AdoQry <> nil then
    begin
      sI_TotalProductNum := L_AdoQry.FieldByName('TotalProductNum').AsString;
      sI_TotalChannelNum := L_AdoQry.FieldByName('TotalChannelNum').AsString;
    end;

end;

procedure TfrmMain.chkReturnFormatClick(Sender: TObject);
begin
    if chkReturnFormat.Checked then
      bG_LongFormat := true
    else
      bG_LongFormat := false;
end;

procedure TfrmMain.IdHTTPServer1CommandGet(AThread: TIdPeerThread;
  ARequestInfo: TIdHTTPRequestInfo; AResponseInfo: TIdHTTPResponseInfo);
var
    L_HttpThread : THttpThread;
    sL_RequestInfo,sL_ResponXml, sL_ReceiveData : String;
    F_FilePath : TextFile;
    nL_FileHandel : Integer;
begin


    //取得監管平台 Request
    //sL_RequestInfo := '<?xml version="1.0" encoding="BIG5" standalone="yes" ?><Msg Version="1" MsgID="2" Type="SMSDown"  DateTime="2002-08-17 15:30:00" SrcCode="110000G01"  DstCode="110000S01"><ExchangeDateQuery /></Msg>';
    //sL_RequestInfo := '<?xml version="1.0" encoding="BIG5" standalone="yes" ?><Msg Version="1" MsgID="2" Type="SMSDown"  DateTime="2002-08-17 15:30:00" SrcCode="110000G01"  DstCode="110000S01"><ICCardQuery InstQuery="1" Date="2003-02-01"></ICCardQuery></Msg>';
    //sL_RequestInfo := '<?xml version="1.0" encoding="BIG5" standalone="yes" ?><Msg Version="1" MsgID="2" Type="SMSDown"  DateTime="2002-08-17 15:30:00" SrcCode="110000G01"  DstCode="110000S01"><CAProductQuery InstQuery="0" Date="2003-02-01"></CAProductQuery></Msg>';
    //sL_RequestInfo := '<?xml version="1.0" encoding="BIG5" standalone="yes" ?><Msg Version="1" MsgID="2" Type="SMSDown"  DateTime="2002-08-17 15:30:00" SrcCode="110000G01"  DstCode="110000S01"><ProdPurchaseQuery InstQuery="0" Date="2003-02-01"></ProdPurchaseQuery></Msg>';
    //sL_RequestInfo := '<?xml version="1.0" encoding="BIG5" standalone="yes" ?><Msg Version="1" MsgID="2" Type="SMSDown"  DateTime="2002-08-17 15:30:00" SrcCode="110000G01"  DstCode="110000S01"><EntitlementQuery InstQuery="0" Date="2003-02-01"></EntitlementQuery></Msg>';
    //ResponseInfo.ContentText :=  sL_ResponXml;
    try
      if dtmMain.isValidRemoteIP(ARequestInfo.RemoteIP) then
      begin
        sL_ReceiveData := ARequestInfo.Params.GetText;
        sL_RequestInfo := TUstr.replaceStr(sL_ReceiveData,''#$D#$A'','');

        //sL_ResponXml := '';//目前不知道規格,先回應空字串
        //ResponseInfo.ContentText := sL_ResponXml;

        dtmMain.savetofile(ARequestInfo.RemoteIP,'c:\RemoteIP.txt');

        if not FileExists('c:\RequestInfo.txt') then
        begin
          nL_FileHandel := FileCreate('c:\RequestInfo.txt');

          if (nL_FileHandel>=0) then
            FileClose(nL_FileHandel);
        end;
        AssignFile(F_FilePath,'c:\RequestInfo.txt');
        Append(F_FilePath);
        Writeln(F_FilePath,'[' +  DateTimeToStr(now)+ ']-' + sL_RequestInfo);
        Flush(F_FilePath);
        CloseFile(F_FilePath);


        L_HttpThread := THttpThread.Create;
        L_HttpThread.sG_RequestXml := sL_RequestInfo;
        L_HttpThread.actionQueryThread;
        L_HttpThread.Free;
      end;
    except

    end;
end;

procedure TfrmMain.updateXmlContent(sI_XmlFilePath, sI_QueryDateTime,
  sI_SrcMsgID, sI_DstMsgID: String);
var
    L_MsgRootNode : IXmlNode;    
begin
    xmlDocForUpdateContent.LoadFromFile(sI_XmlFilePath);

    L_MsgRootNode := xmlDocForUpdateContent.DocumentElement;

    TMyXML.setAttribute(L_MsgRootNode, 'MsgID',sI_SrcMsgID);
    TMyXML.setAttribute(L_MsgRootNode, 'ReplyID',sI_DstMsgID);
    TMyXML.setAttribute(L_MsgRootNode, 'DateTime',sI_QueryDateTime);

    xmlDocForUpdateContent.SaveToFile(sI_XmlFilePath);
    xmlDocForUpdateContent.Active := false;

end;

procedure TfrmMain.combineData(sI_CompCode, sI_Mode, sI_QueryStartDate,
  sI_QueryStopDate: String);
    procedure loadBaseXml(sL_XmlFilePath:String; I_XmlDoc:TXMLDocument);
    begin
      if FileExists(sL_XmlFilePath) then
        I_XmlDoc.LoadFromFile(sL_XmlFilePath);
    end;
    procedure appendXmlData(I_BaseXmlDoc:TXMLDocument;sI_DataNodeName:String; I_AdoQuery:TADOQuery);overload;
    var
        L_TmpNode, L_BaseDataNode : IXmlNode;
        sL_Version, sL_SubName, sL_SubVocation, sL_SubAge, sL_SubIDCardNo : String;
        sL_Feature, sL_CasID, sL_ICCardNo, sL_ICCardID : String;
        sL_Status, sL_RealDateTime : String;
    begin
      with I_AdoQuery do
      begin
          I_AdoQuery.First;
          while not I_AdoQuery.Eof do
          begin
            //L_DetailChildNode := TMyXML.addChild(L_MasterChildNode, sL_DetailChildNode,'');
            sL_RealDateTime := TUstr.replaceStr(DateTimeToStr(FieldByName('RealDateTime').AsDateTime),'/','-');
            if Length(sL_RealDateTime) <= 10 then
              sL_RealDateTime := sL_RealDateTime + ' 00:00:00';

            sL_Status := FieldByName('Status').AsString;
            if sL_Status = '0' then
              sL_Status := 'New'
            else if sL_Status = '1' then
              sL_Status := 'Delete';

            sL_ICCardNo := FieldByName('ICCardNo').AsString;
            sL_ICCardID := FieldByName('ICCardID').AsString;
            sL_CasID := FieldByName('CasID').AsString;
            sL_Feature := FieldByName('Feature').AsString;
            sL_Version := FieldByName('Version').AsString;
            sL_SubName := FieldByName('SubName').AsString;
            sL_SubVocation := FieldByName('SubVocation').AsString;
            sL_SubAge := FieldByName('SubAge').AsString;
            sL_SubIDCardNo := FieldByName('SubIDCardNo').AsString;

            L_BaseDataNode := I_BaseXmlDoc.DocumentElement.ChildNodes.FindNode (sI_DataNodeName);
            if L_BaseDataNode<> nil then
            begin
              L_TmpNode := TMyXML.addChild(L_BaseDataNode, 'ICCard','');
              TMyXML.setAttribute(L_TmpNode, 'DateTime',sL_RealDateTime);
              TMyXML.setAttribute(L_TmpNode, 'Status',sL_Status);
              TMyXML.setAttribute(L_TmpNode, 'ICCardNo',sL_ICCardNo);
              TMyXML.setAttribute(L_TmpNode, 'ICCardID',sL_ICCardID);
              TMyXML.setAttribute(L_TmpNode, 'CasID',sL_CasID);
              TMyXML.setAttribute(L_TmpNode, 'Feature',sL_Feature);
              TMyXML.setAttribute(L_TmpNode, 'Version',sL_Version);
              TMyXML.setAttribute(L_TmpNode, 'SubName',sL_SubName);
              TMyXML.setAttribute(L_TmpNode, 'SubVocation',sL_SubVocation);
              TMyXML.setAttribute(L_TmpNode, 'SubAge',sL_SubAge);
              TMyXML.setAttribute(L_TmpNode, 'SubIDCardNo',sL_SubIDCardNo);

            end;





            Next;
          end;
      end;    
    end;

    procedure appendXmlData(I_BaseXmlDoc, I_TmpXmlDoc:TXMLDocument;sL_TmpXmlFilePath, sI_DataNodeName:String);overload;
    var
        ii : Integer;
        L_DataNode, L_TmpNode, L_BaseDataNode : IXmlNode;
    begin
      if FileExists(sL_TmpXmlFilePath) then
        I_TmpXmlDoc.LoadFromFile(sL_TmpXmlFilePath);


      L_DataNode := I_TmpXmlDoc.DocumentElement.ChildNodes.FindNode (sI_DataNodeName);
      if L_DataNode<>nil then
      begin
        L_TmpNode := L_DataNode.CloneNode(true);
      end;

      if L_TmpNode<>nil then
      begin
        I_BaseXmlDoc.DocumentElement.ChildNodes.Add(L_TmpNode);
        {
        L_BaseDataNode := I_BaseXmlDoc.DocumentElement.ChildNodes.FindNode (sI_DataNodeName);
        if L_BaseDataNode<> nil then
        begin
          for ii:=0 to L_TmpNode.ChildNodes.Count-1 do
          begin

            L_BaseDataNode.ChildNodes.Add(L_TmpNode.ChildNodes[ii]);
          end;

        end;
        }
      end;
    end;

    procedure  writeIcCardAttribute(I_BaseXmlDoc:TXMLDocument; sI_TotalICCardNum,sI_TotalSubNum, sI_DataNodeName:String);
    var
      L_DataNode : IXmlNode;
    begin
      L_DataNode := I_BaseXmlDoc.DocumentElement.ChildNodes.FindNode (sI_DataNodeName);
      if L_DataNode<>nil then
      begin
        TMyXML.setAttribute(L_DataNode, 'TotalICCardNum', sI_TotalICCardNum);
        TMyXML.setAttribute(L_DataNode, 'TotalSubNum', sI_TotalSubNum);
      end;

    end;

    procedure  writeCaProductAttribute(I_BaseXmlDoc:TXMLDocument; sI_TotalProductNum,sI_TotalChannelNum, sI_DataNodeName:String);
    var
      L_DataNode : IXmlNode;
    begin
      L_DataNode := I_BaseXmlDoc.DocumentElement.ChildNodes.FindNode (sI_DataNodeName);
      if L_DataNode<>nil then
      begin
        TMyXML.setAttribute(L_DataNode, 'TotalProductNum', sI_TotalProductNum);
        TMyXML.setAttribute(L_DataNode, 'TotalChannelNum', sI_TotalChannelNum);
      end;

    end;

var
    L_BaseXmlDoc, L_TmpXmlDoc : TXMLDocument;
    ii : Integer;
    sL_BaseXmlFilePath, sL_XmlFilePath, sL_HandleEDateTime, sL_QueryType : String;
    nL_DayDiff : integer;
    sL_RefDate, sL_DataNodeName : String;
    dL_StartDate, dL_EndDate : TDateTime;
    sL_TotalICCardNum,sL_TotalSubNum, sL_Table : String;
    sL_TotalProductNum,sL_TotalChannelNum : String;
    bL_UpdateIcCardAttribute, bL_UpdateCaProductAttribute : boolean;
    L_AdoQry : TADOQuery;
begin
    bL_UpdateIcCardAttribute := false;
    bL_UpdateCaProductAttribute := false;

    sL_RefDate := sI_QueryStopDate + ' 23:59:59';
    if sI_Mode = '1' then
    begin
      //執行IC卡資訊之查詢(ICCardQuery)
      sL_QueryType := IC_CARD_Query;
      sL_DataNodeName := 'ICCardQueryReport';
      getSO459ICCardNumAndSubNum(sI_CompCode, sI_Mode,sL_RefDate,sL_TotalICCardNum,sL_TotalSubNum);
      bL_UpdateIcCardAttribute := true;
      sL_Table := 'SO451';
    end
    else if sI_Mode = '2' then
    begin
       //CA產品資訊之查詢(CAProductQuery)
      sL_QueryType := CA_PRODUCT_QUERY;
      sL_DataNodeName := 'CAProductQueryReport';
      getSO459ProductNumAndChannelNum(sI_CompCode, sI_Mode,sL_RefDate,sL_TotalProductNum,sL_TotalChannelNum);
      bL_UpdateCaProductAttribute := true;
      sL_Table := 'SO453';
    end
    else if sI_Mode = '3' then
    begin
      //產品訂購資訊之查詢(ProdPurchaseQuery)
      sL_QueryType := PROD_PURCHASE_QUERY;
      sL_DataNodeName := 'ProdPurchaseQueryReport';
      sL_Table := 'SO455';
    end
    else if sI_Mode = '4' then
    begin
      //授權資訊之查詢(EntitlementQuery)
      sL_QueryType := ENTITLEMENT_QUERY;
      sL_DataNodeName := 'EntitlementQueryReport';
      sL_Table := 'SO457';      
    end;
//returnQueryXML
    dL_StartDate := Round(StrToDate(sI_QueryStartDate));
    dL_EndDate := Round(StrToDate(sI_QueryStopDate));

    nL_DayDiff := Round(dL_EndDate - dL_StartDate);

    L_BaseXmlDoc := TXMLDocument.Create(Application);
    L_TmpXmlDoc := TXMLDocument.Create(Application);

    for ii:=0 to nL_DayDiff do
    begin
      sL_HandleEDateTime := DateToStr(dL_EndDate-ii) + ' 23:59:59';
      if ii=0 then
      begin
        sL_BaseXmlFilePath := dtmMain.getRealXmlFilePath(sI_CompCode,'0',sL_QueryType,sL_HandleEDateTime,false, false);
        loadBaseXml(sL_BaseXmlFilePath, L_BaseXmlDoc);
      end
      else
      begin
        L_AdoQry := dtmMain.getQueryData(sL_QueryType,sL_Table,sI_CompCode,sL_HandleEDateTime);
        appendXmlData(L_BaseXmlDoc, sL_DataNodeName,L_AdoQry);
//        dtmMain.adoQryQueryData.Clone(L_AdoQry,ltUnspecified);

      {
        sL_XmlFilePath := dtmMain.getRealXmlFilePath(sI_CompCode,'0',sL_QueryType,sL_HandleEDateTime,false,false);
        appendXmlData(L_BaseXmlDoc,L_TmpXmlDoc, sL_XmlFilePath, sL_DataNodeName);
      }
      end;

    end;

    if bL_UpdateIcCardAttribute then
    begin
      writeIcCardAttribute(L_BaseXmlDoc, sL_TotalICCardNum,sL_TotalSubNum, sL_DataNodeName);
    end;

    if bL_UpdateCaProductAttribute then
    begin
      writeCaProductAttribute(L_BaseXmlDoc,sL_TotalProductNum,sL_TotalChannelNum, sL_DataNodeName);
    end;

    L_BaseXmlDoc.SaveToFile(sL_BaseXmlFilePath);

    L_BaseXmlDoc.Active := false;
    L_TmpXmlDoc.Active := false;

    L_BaseXmlDoc.Free;
    L_TmpXmlDoc.Free;
end;

procedure TfrmMain.Button1Click(Sender: TObject);
begin
combineData('7','1','2003/08/19','2003/08/21');
//combineData('7','2','2003/08/19','2003/08/21');
end;


end.
