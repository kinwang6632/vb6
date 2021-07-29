unit THttpThreadU;

interface

uses
  Classes,SysUtils,Dialogs;

type
  THttpThread = class(TThread)
  private
    { Private declarations }
    procedure handleQuery;overload;
  protected
    procedure Execute; override;
  public
    sG_RequestXml: String;
    constructor Create;overload;
    procedure actionQueryThread;overload;

  end;

implementation

uses frmMainU, dtmMainU, UdateTimeu, Ustru, ConstU;

{ Important: Methods and properties of objects in VCL or CLX can only be used
  in a method called using Synchronize, for example,

      Synchronize(UpdateCaption);

  and UpdateCaption could look like,

    procedure THttpThread.UpdateCaption;
    begin
      Form1.Caption := 'Updated in a thread';
    end; }

{ THttpThread }
procedure THttpThread.actionQueryThread;
begin
  try
  
    if  sG_RequestXml <> '' then
      Synchronize(handleQuery);

  except

  end;

end;

constructor THttpThread.Create;
begin
    inherited Create(false);
end;

procedure THttpThread.Execute;
begin
  { Place thread code here }
end;


procedure THttpThread.handleQuery;
var
    sL_TmpXmlFilePath,sL_ResponXml : String;
    sL_Version,sL_DstMsgID,sL_DstType,sL_SrcCode,sL_QueryDateTime,sL_DstCode : String;
    sL_InstQuery,sL_HandleDate,sL_HandleEDateTime,sL_QueryType : String;
    sL_CurrDate,sL_ErrCode,sL_ErrMsg,sL_CompCode : String;
    bL_IsMsgDataOK,bL_IsQueryDataOK : Boolean;
    sL_SrcType,sL_CasID,sL_FirstDate,sL_CmdSeq,sL_Table,sL_SrcMsgID : String;
    sL_ErrCode1,sL_ErrMsg1,sL_ErrCode2,sL_ErrMsg2,sL_ModeType : String;
    sL_TotalICCardNum,sL_TotalSubNum,sL_TotalProductNum,sL_TotalChannelNum : String;
begin
    //down, 直接 loadXML 會有 error, 所以先將此 XML 存到暫存檔中
    sL_TmpXmlFilePath := frmMain.getRendomFileName;
    dtmMain.saveToFile(sG_RequestXml, sL_TmpXmlFilePath);
    //up, 直接 loadXML 會有 error, 所以先將此 XML 存到暫存檔中

    sL_CurrDate := TUdateTime.GetPureDateStr(NOW,'/');
    sL_CmdSeq := frmMain.getCmdSeq(sL_CurrDate);

    bL_IsMsgDataOK := frmMain.parseXmlMsgAttribute(sL_TmpXmlFilePath,sL_Version,sL_DstMsgID,sL_DstType,sL_SrcCode,sL_QueryDateTime,sL_DstCode,sL_SrcType,sL_CasID,sL_FirstDate,sL_CompCode,sL_ErrCode1,sL_ErrMsg1);
    bL_IsQueryDataOK := frmMain.parseXmlQueryData(sL_TmpXmlFilePath,sL_InstQuery,sL_HandleDate,sL_QueryType,sL_ErrCode2,sL_ErrMsg2);
    DeleteFile(sL_TmpXmlFilePath);

    if sL_InstQuery = '0' then
      sL_HandleEDateTime := TUstr.replaceStr(sL_HandleDate,'-','/') + ' 23:59:59'
    else
    begin
      sL_HandleEDateTime := TUdateTime.GetPureDateTimeStr(NOW,'/',':');
    end;

    if UpperCase(sL_QueryType) = UpperCase(IC_CARD_QUERY) then
      sL_ModeType := '1'   //1:IC卡資訊
    else if UpperCase(sL_QueryType) = UpperCase(CA_PRODUCT_QUERY) then
      sL_ModeType := '2'   //2:產品定義資訊
    else if UpperCase(sL_QueryType) = UpperCase(PROD_PURCHASE_QUERY) then
      sL_ModeType := '3'   //3:產品定購資訊
    else if UpperCase(sL_QueryType) = UpperCase(ENTITLEMENT_QUERY) then
      sL_ModeType := '4'  //4:授權資訊
    else if UpperCase(sL_QueryType) = UpperCase(EXCHANGE_DATE_QUERY) then
      sL_ModeType := '5';  //5:數據交換日期

    //將查詢資訊先Log至SO461
    //dtmMain.logToSO461(sL_CompCode,sL_ModeType,sL_HandleEDateTime,sL_QueryDateTime,sL_InstQuery,sL_SrcCode,sL_DstCode,sL_DstMsgID,sG_RequestXml);
    dtmMain.logToSO461(sL_CompCode,sL_ModeType,sL_HandleEDateTime,sL_QueryDateTime,
                       sL_InstQuery,sL_SrcCode,sL_DstCode,sL_DstMsgID,'');


    //顯示要傳送指令
    frmMain.showSendMsgUI(sL_CmdSeq,sL_DstMsgID,sL_SrcCode,sL_CompCode,sL_HandleEDateTime,
                    sL_QueryType,sL_InstQuery);

    if bL_IsQueryDataOK and bL_IsMsgDataOK then
    begin
      sL_ResponXml := frmMain.actionQuery(sL_CmdSeq,sL_ModeType,sL_QueryType,sL_CompCode,sL_SrcCode,sL_DstCode,sL_DstMsgID,sL_CasID,sL_InstQuery,sL_FirstDate,sL_HandleEDateTime,sL_Version,sL_DstType,sL_QueryDateTime,sL_SrcType);
    end
    else
    begin

      sL_Table := '';
      sL_SrcMsgID := dtmMain.getSrcMsgID(sL_CompCode, '');
      if sL_ErrCode1 <> '0' then
      begin
        sL_ErrCode := sL_ErrCode1;
        sL_ErrMsg := sL_ErrMsg1;
      end
      else if sL_ErrCode2 <> '0' then
      begin
        sL_ErrCode := sL_ErrCode2;
        sL_ErrMsg := sL_ErrMsg2;
      end;

      if UpperCase(sL_QueryType) = UpperCase(IC_CARD_QUERY) then
        sL_ModeType := '1'   //1:IC卡資訊
      else if UpperCase(sL_QueryType) = UpperCase(CA_PRODUCT_QUERY) then
        sL_ModeType := '2'   //2:產品定義資訊
      else if UpperCase(sL_QueryType) = UpperCase(PROD_PURCHASE_QUERY) then
        sL_ModeType := '3'   //3:產品定購資訊
      else if UpperCase(sL_QueryType) = UpperCase(ENTITLEMENT_QUERY) then
        sL_ModeType := '4';  //4:授權資訊

        
      sL_TotalICCardNum := '';
      sL_TotalSubNum := '';
      sL_TotalProductNum := '';
      sL_TotalChannelNum := '';

      sL_ResponXml := frmMain.returnQueryXML(sL_QueryType,sL_Table,sL_CmdSeq,sL_CompCode,sL_Version,sL_QueryDateTime,
                           sL_SrcCode,sL_SrcType,sL_SrcMsgID,
                           sL_DstCode,sL_DstType,sL_DstMsgID,
                           sL_CasID,sL_InstQuery,sL_FirstDate,sL_HandleEDateTime,
                           sL_ErrCode,sL_ErrMsg,sL_TotalICCardNum,sL_TotalSubNum,
                           sL_TotalProductNum,sL_TotalChannelNum,0);

      dtmMain.logToSO460(sL_CompCode,sL_ModeType,sL_HandleEDateTime,sL_InstQuery,
                sL_SrcCode,sL_SrcMsgID,sL_DstCode,sL_DstMsgID,sL_ErrCode,sL_ErrMsg);

      frmMain.updateReturnDataUI(sL_CmdSeq,sL_DstMsgID,sL_ErrCode,sL_ErrMsg,sL_SrcMsgID);

    end;

    frmMain.sendResponseXml(sL_CompCode,sL_ResponXml);
end;

end.
