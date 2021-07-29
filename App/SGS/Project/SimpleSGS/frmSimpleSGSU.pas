unit frmSimpleSGSU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ActnList, StdCtrls, ExtCtrls, IdBaseComponent, IdComponent,
  IdTCPServer, IdHTTPServer, IdTCPConnection, IdTCPClient, IdHTTP, fraYMDU,
  xmldom, XMLIntf, msxmldom, XMLDoc ,SyncObjs, Buttons,
  IdCustomHTTPServer, IdMultipartFormData, idSocketHandle;

type
  TfrmSimpleSGS = class(TForm)
    Panel1: TPanel;
    cbActive: TCheckBox;
    alGeneral: TActionList;
    acActivate: TAction;
    Panel2: TPanel;
    Memo1: TMemo;                   

    Panel3: TPanel;
    cobCompCode: TComboBox;
    cobQueryType: TComboBox;
    chkNow: TCheckBox;
    fraYMD1: TfraYMD;
    btnSend: TButton;
    btnExit: TButton;
    setXMLDocument: TXMLDocument;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    lbLog: TListBox;
    BitBtn1: TBitBtn;
    HTTPServer: TIdHTTPServer;
    HTTP: TIdHTTP;
    procedure acActivateExecute(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure chkNowClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnSendClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure HTTPServerCommandGet(AThread: TIdPeerThread;
      ARequestInfo: TIdHTTPRequestInfo;
      AResponseInfo: TIdHTTPResponseInfo);
  private
    { Private declarations }
    UILock: TCriticalSection;
    procedure setCaption;
    function getCobCompCode : String;
    function IsDataOk : Boolean;
    procedure setLanguageString;
    procedure setDateTimeFormat;
  public
    { Public declarations }
    procedure sendRequestXml(sI_CompCode,sI_RequestXml : String);
    function setQueryXML(sI_CompCode,sI_QueryDate,sI_QueryType,sI_InstQuery : String) : String;
    procedure DisplayMessage(const Msg: string);
  end;

var
  frmSimpleSGS: TfrmSimpleSGS;

implementation

uses dtmMainU, Ustru, ConstU, xmlU;

{$R *.dfm}

{ TfrmSimpleSGS }

procedure TfrmSimpleSGS.acActivateExecute(Sender: TObject);
var
  Binding : TIdSocketHandle;
begin
  //關 HTTP Server 監聽
  acActivate.Checked := not acActivate.Checked;
  
  if not HTTPServer.Active then
  begin
    HTTPServer.Bindings.Clear;
    Binding := HTTPServer.Bindings.Add;
    Binding.Port := dtmMain.nG_DstHttpPort;

    Binding.IP := dtmMain.sG_LocalIP;
//    Binding.IP := '127.0.0.1';



  end;


  if acActivate.Checked then
  begin
    try
      HTTPServer.Active := true;
    except
      on e: exception do
      begin
        acActivate.Checked := False;
      end;
    end;
  end
  else
  begin
    HTTPServer.Active := false;
    HTTPServer.Intercept := nil;

  end;

end;

procedure TfrmSimpleSGS.FormShow(Sender: TObject);
var
  ii : Integer;
begin
    setLanguageString;

    cbActive.Checked := true;

    cobCompCode.Clear;
    for ii:=0 to dtmMain.G_CompCodeAndNameStrList.Count -1 do
      cobCompCode.Items.Add(dtmMain.G_CompCodeAndNameStrList.Strings[ii]);


    cobCompCode.ItemIndex := 0;
    cobQueryType.ItemIndex := 1;

    cobCompCode.SetFocus;

    UILock := TCriticalSection.Create;
end;

procedure TfrmSimpleSGS.sendRequestXml(sI_CompCode, sI_RequestXml: String);
var
    L_Data : TIdMultiPartFormDataStream;
    L_Response: TStringStream;
    L_XmlString: TStringList;
    sL_RequestXml,sL_SrcIP,sL_HttpTitle,sL_XmlString : String;
    F_FilePath : TextFile;
    nL_FileHandel : Integer;
begin
    if not FileExists('c:\RequestInfo.txt') then
    begin
      nL_FileHandel := FileCreate('c:\RequestInfo.txt');

      if (nL_FileHandel>=0) then
        FileClose(nL_FileHandel);
    end;
    AssignFile(F_FilePath,'c:\RequestInfo.txt');
    Append(F_FilePath);

    L_Response := TStringStream.Create('');
    L_XmlString := TStringList.Create;

    sL_RequestXml := TUstr.replaceStr(sI_RequestXml,''#$D#$A'','');
    L_XmlString.Add(sL_RequestXml);

    L_Data := TIdMultiPartFormDataStream.Create;
    L_Data.AddFormField('xml', sL_RequestXml);

    sL_SrcIP := dtmMain.getSrcIP(sI_CompCode);
//    sL_HttpTitle := 'http://' + sL_SrcIP + '/?';
    sL_HttpTitle := 'http://' + sL_SrcIP;
    //'http://10.0.0.21/?'

    sL_XmlString :=  sL_HttpTitle + sL_RequestXml;
    Memo1.Clear;
    Memo1.Lines.Add(sL_XmlString);

    Writeln(F_FilePath,'[' +  DateTimeToStr(now)+ ']-' + sL_HttpTitle + sL_RequestXml);



    HTTP.Request.Username := '';
    HTTP.Request.Password := '';
    HTTP.HandleRedirects := true;
//    HTTP.Request.ProxyServer := '';
//    HTTP.Request.ProxyPort := dtmMain.nG_HttpPort;
//    HTTP.Request.ContentType := 'text/html';
//    HTTP.Request.ContentType := '';
    HTTP.Request.ContentType := 'application/x-www-form-urlencoded';
    HTTP.ProtocolVersion := pv1_1;


    try
//      HTTP.Post(sL_HttpTitle,L_XmlString, L_Response);

      HTTP.Post(sL_HttpTitle,L_XmlString);
//    HTTP.Post(sL_HttpTitle,L_Data);

      Writeln(F_FilePath,'[' +  DateTimeToStr(now)+ ']-' + HTTP.ResponseText);
    except
      Writeln(F_FilePath,'[' +  DateTimeToStr(now)+ ']-' + HTTP.ResponseText);
    end;
{

    NMHTTP1.InputFileMode := FALSE;
    NMHTTP1.OutputFileMode := FALSE;
    NMHTTP1.ReportLevel := Status_Basic;
    NMHTTP1.Port := dtmMain.nG_HttpPort;;

    With NMHTTP1.HeaderInfo do
    Begin
      Cookie := '';
      LocalMailAddress := '';
      LocalProgram := '';
      Referer := '';
      UserID := '';
      Password := '';
    End;

    try
      NMHTTP1.Post(sL_HttpTitle,sL_RequestXml);
      Writeln(F_FilePath,'[' +  DateTimeToStr(now)+ ']OK-' + NMHTTP1.Header);
    except
      Writeln(F_FilePath,'[' +  DateTimeToStr(now)+ ']ERROR-' + NMHTTP1.Header);
    end;

    Memo1.Lines.Add('');
    Memo1.Lines.Add('');
    Memo1.Lines.Add(NMHTTP1.Header);
    //Memo1.Text := NMHTTP1.Header;
    //Memo1.Text := NMHTTP1.Body;
}

    Flush(F_FilePath);
    CloseFile(F_FilePath);

    L_Response.Free;
    L_XmlString.Free;
end;


{
procedure TfrmSimpleSGS.sendRequestXml(sI_CompCode, sI_RequestXml: String);
var
    L_Response: TStringStream;
    L_XmlString: TStringList;
    sL_RequestXml,sL_SrcIP,sL_HttpTitle,sL_XmlString : String;
    F_FilePath : TextFile;
    nL_FileHandel : Integer;
begin
    if not FileExists('c:\RequestInfo.txt') then
    begin
      nL_FileHandel := FileCreate('c:\RequestInfo.txt');

      if (nL_FileHandel>=0) then
        FileClose(nL_FileHandel);
    end;
    AssignFile(F_FilePath,'c:\RequestInfo.txt');
    Append(F_FilePath);

    L_Response := TStringStream.Create('');
    L_XmlString := TStringList.Create;

    sL_RequestXml := TUstr.replaceStr(sI_RequestXml,''#$D#$A'','');
    L_XmlString.Add('xml=' + sL_RequestXml);

    sL_SrcIP := dtmMain.getSrcIP(sI_CompCode);
//    sL_HttpTitle := 'http://' + sL_SrcIP + '/?';
    sL_HttpTitle := 'http://' + sL_SrcIP;
    //'http://10.0.0.21/?'

    sL_XmlString :=  sL_HttpTitle + sL_RequestXml;
    Memo1.Clear;
    Memo1.Lines.Add(sL_XmlString);

    Writeln(F_FilePath,'[' +  DateTimeToStr(now)+ ']-' + sL_HttpTitle + sL_RequestXml);



    HTTP.Request.Username := '';
    HTTP.Request.Password := '';
//    HTTP.Request.ProxyServer := '';
//    HTTP.Request.ProxyPort := dtmMain.nG_HttpPort;
    HTTP.Request.ContentType := '';
//    HTTP.Request.ContentType := 'application/x-www-form-urlencoded';
    HTTP.ProtocolVersion := pv1_1;


    try
//      HTTP.Post(sL_HttpTitle,L_XmlString, L_Response);
      HTTP.Post(sL_HttpTitle,L_XmlString);
      Writeln(F_FilePath,'[' +  DateTimeToStr(now)+ ']-' + HTTP.ResponseText);
    except
      Writeln(F_FilePath,'[' +  DateTimeToStr(now)+ ']-' + HTTP.ResponseText);
    end;

    Flush(F_FilePath);
    CloseFile(F_FilePath);

    L_Response.Free;
    L_XmlString.Free;
end;

}
function TfrmSimpleSGS.getCobCompCode: String;
var
    L_CompCodeStrList : TStringList;
    sL_CompCodeAndName : String;
begin
    //取得所屬公司
    sL_CompCodeAndName := cobCompCode.Text;

    if sL_CompCodeAndName <> '' then
    begin
      L_CompCodeStrList := TStringList.Create;
      L_CompCodeStrList := TUstr.ParseStrings(cobCompCode.Text,'-');
      Result := L_CompCodeStrList.Strings[0];
    end
    else
      Result := '';

    L_CompCodeStrList.Free;
end;

procedure TfrmSimpleSGS.chkNowClick(Sender: TObject);
begin
    if chkNow.Checked then
      fraYMD1.Enabled := false
    else
       fraYMD1.Enabled := true;

end;

procedure TfrmSimpleSGS.btnExitClick(Sender: TObject);
begin
    Application.Terminate;
end;

function TfrmSimpleSGS.IsDataOk: Boolean;
var
    sL_CompCode,sL_QueryType,sL_QueryDate : String;
begin
  sL_CompCode := cobCompCode.Text;
  if sL_CompCode = '' then
  begin
      MessageDlg(dtmMain.getLanguageSettingInfo('frmSimpleSGS_Msg_1_Content'),mtError, [mbOK],0);
      cobCompCode.SetFocus;
      IsDataOk := false;
      exit;
  end;

  sL_QueryType := cobQueryType.Text;
  if sL_QueryType = '' then
  begin
      MessageDlg(dtmMain.getLanguageSettingInfo('frmSimpleSGS_Msg_2_Content'),mtError, [mbOK],0);
      cobQueryType.SetFocus;
      IsDataOk := false;
      exit;
  end;

  sL_QueryDate := Trim(TUstr.replaceStr(fraYMD1.getYMD,'/',''));
  if (chkNow.Checked = false) and (sL_QueryDate= '') then
  begin
      MessageDlg(dtmMain.getLanguageSettingInfo('frmSimpleSGS_Msg_3_Content'),mtError, [mbOK],0);
      fraYMD1.SetFocus;
      IsDataOk := false;
      exit;
  end;

end;

procedure TfrmSimpleSGS.btnSendClick(Sender: TObject);
var
    sL_CompCode,sL_QueryDate,sL_QueryType,sL_InstQuery,sL_XmlString : String;
begin
    if IsDataOk then
    begin
      //取得所屬公司
      sL_CompCode := getCobCompCode;

      sL_QueryDate := TUstr.replaceStr(Trim(fraYMD1.getYMD),'/','-');

      if cobQueryType.ItemIndex = 0 then
        sL_QueryType := EXCHANGE_DATE_QUERY
      else if cobQueryType.ItemIndex = 1 then
        sL_QueryType := IC_CARD_QUERY
      else if cobQueryType.ItemIndex = 2 then
        sL_QueryType := CA_PRODUCT_QUERY
      else if cobQueryType.ItemIndex = 3 then
        sL_QueryType := PROD_PURCHASE_QUERY
      else if cobQueryType.ItemIndex = 4 then
        sL_QueryType := ENTITLEMENT_QUERY;


      if chkNow.Checked then
        sL_InstQuery := '1'
      else
        sL_InstQuery := '0';

      sL_XmlString := setQueryXML(sL_CompCode,sL_QueryDate,sL_QueryType,sL_InstQuery);

      sendRequestXml(sL_CompCode,sL_XmlString);
    end;
end;

function TfrmSimpleSGS.setQueryXML(sI_CompCode, sI_QueryDate, sI_QueryType,
  sI_InstQuery: String): String;
var
    nL_Ndx,nL_Seed : Integer;
    sL_Version,sL_MsgID,sL_DstType,sL_QueryDateTime : String;
    sL_DstCode,sL_SrcCode,sL_Result,sL_XmlString : String;
    L_MsgRootNode,L_MasterChildNode : IXmlNode;
begin
    nL_Ndx := dtmMain.G_SGSParamIniStrList.IndexOf(sI_CompCode);
    if nL_Ndx<>-1 then
    begin
       sL_Version := (dtmMain.G_SGSParamIniStrList.Objects[nL_Ndx] as TSGSParamIniData).sVersion;
       sL_DstType := (dtmMain.G_SGSParamIniStrList.Objects[nL_Ndx] as TSGSParamIniData).sDstType;
       sL_DstCode := (dtmMain.G_SGSParamIniStrList.Objects[nL_Ndx] as TSGSParamIniData).sDstCode;
       sL_SrcCode := (dtmMain.G_SGSParamIniStrList.Objects[nL_Ndx] as TSGSParamIniData).sSrcCode;
    end;

    Randomize;
    nL_Seed := 100000;

    sL_MsgID := IntToStr(Random(nL_Seed));
    sL_QueryDateTime := TUstr.replaceStr(DateTimeToStr((NOW)),'/','-');


    //串起XML
    setXMLDocument.Active := true;

    if UpperCase(dtmMAIN.sG_TcOrSc) = 'TC' then
      setXMLDocument.Encoding := XML_ENCODING_TC
    else
      setXMLDocument.Encoding := XML_ENCODING_SC;

    setXMLDocument.StandAlone := 'yes';

    L_MsgRootNode := TMyXML.createRootNode(setXMLDocument,'Msg','', sL_Result);
    TMyXML.setAttribute(L_MsgRootNode, 'DateTime',sL_QueryDateTime);
    TMyXML.setAttribute(L_MsgRootNode, 'DstCode',sL_SrcCode);
    TMyXML.setAttribute(L_MsgRootNode, 'MsgID',sL_MsgID);
    TMyXML.setAttribute(L_MsgRootNode, 'SrcCode',sL_DstCode);
    TMyXML.setAttribute(L_MsgRootNode, 'Type',sL_DstType);
    TMyXML.setAttribute(L_MsgRootNode, 'Version',sL_Version);

    {
    TMyXML.setAttribute(L_MsgRootNode, 'Version',sL_Version);
    TMyXML.setAttribute(L_MsgRootNode, 'MsgID',sL_MsgID);
    TMyXML.setAttribute(L_MsgRootNode, 'Type',sL_DstType);
    TMyXML.setAttribute(L_MsgRootNode, 'DateTime',sL_QueryDateTime);
    TMyXML.setAttribute(L_MsgRootNode, 'SrcCode',sL_DstCode);
    TMyXML.setAttribute(L_MsgRootNode, 'DstCode',sL_SrcCode);
    }

    L_MasterChildNode := TMyXML.addChild(L_MsgRootNode, sI_QueryType,'');

    if UpperCase(sI_QueryType) <> UpperCase(EXCHANGE_DATE_QUERY) then
    begin
      if sI_InstQuery = '0' then
        TMyXML.setAttribute(L_MasterChildNode, 'Date',sI_QueryDate)
      else
        TMyXML.setAttribute(L_MasterChildNode, 'Date','');

      TMyXML.setAttribute(L_MasterChildNode, 'InstQuery',sI_InstQuery);
    end;

    sL_XmlString := setXMLDocument.XML.GetText;

    setXMLDocument.Active := false;
    Result := sL_XmlString;
end;

procedure TfrmSimpleSGS.setLanguageString;
begin
    Caption := dtmMain.getLanguageSettingInfo('frmSimple_Caption');
    Label1.Caption := dtmMain.getLanguageSettingInfo('frmSimple_Label1_Caption');
    Label2.Caption := dtmMain.getLanguageSettingInfo('frmSimple_Label2_Caption');
    chkNow.Caption := dtmMain.getLanguageSettingInfo('frmSimple_chkNow_Caption');
    Label3.Caption := dtmMain.getLanguageSettingInfo('frmSimple_Label3_Caption');

    cobQueryType.Items.Add(dtmMain.getLanguageSettingInfo('frmSimpleSGS_rgpQueryType_Item0_Caption'));
    cobQueryType.Items.Add(dtmMain.getLanguageSettingInfo('frmSimpleSGS_rgpQueryType_Item1_Caption'));
    cobQueryType.Items.Add(dtmMain.getLanguageSettingInfo('frmSimpleSGS_rgpQueryType_Item2_Caption'));
    cobQueryType.Items.Add(dtmMain.getLanguageSettingInfo('frmSimpleSGS_rgpQueryType_Item3_Caption'));
    cobQueryType.Items.Add(dtmMain.getLanguageSettingInfo('frmSimpleSGS_rgpQueryType_Item4_Caption'));
end;

procedure TfrmSimpleSGS.DisplayMessage(const Msg: string);
begin

    UILock.Acquire;
    try
      lbLog.ItemIndex := lbLog.Items.Add(Msg);
      //lbLog.Items.Add(Msg);
    finally
      UILock.Release;
    end;

end;

procedure TfrmSimpleSGS.FormDestroy(Sender: TObject);
begin
    UILock.Free;
end;

procedure TfrmSimpleSGS.setCaption;
begin
    //Label3.Caption := 'xxxx';
end;

procedure TfrmSimpleSGS.setDateTimeFormat;
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

procedure TfrmSimpleSGS.FormCreate(Sender: TObject);
begin
    setDateTimeFormat;
end;


procedure TfrmSimpleSGS.BitBtn1Click(Sender: TObject);
var
    sL_CompCode,sL_QueryDate,sL_QueryType,sL_InstQuery,sL_XmlString : String;
    ii : Integer;
begin
    if IsDataOk then
    begin
      //取得所屬公司
      sL_CompCode := getCobCompCode;

      sL_QueryDate := TUstr.replaceStr(Trim(fraYMD1.getYMD),'/','-');

      if chkNow.Checked then
        sL_InstQuery := '1'
      else
        sL_InstQuery := '0';


      for ii:=1 to 4 do
      begin
        if ii = 1 then
          sL_QueryType := IC_CARD_QUERY
        else if ii = 2 then
          sL_QueryType := CA_PRODUCT_QUERY
        else if ii = 3 then
          sL_QueryType := PROD_PURCHASE_QUERY
        else if ii = 4 then
          sL_QueryType := ENTITLEMENT_QUERY;
{
        if cobQueryType.ItemIndex = 0 then
          sL_QueryType := EXCHANGE_DATE_QUERY
        else if cobQueryType.ItemIndex = 1 then
          sL_QueryType := IC_CARD_QUERY
        else if cobQueryType.ItemIndex = 2 then
          sL_QueryType := CA_PRODUCT_QUERY
        else if cobQueryType.ItemIndex = 3 then
          sL_QueryType := PROD_PURCHASE_QUERY
        else if cobQueryType.ItemIndex = 4 then
          sL_QueryType := ENTITLEMENT_QUERY;
}
        sL_XmlString := setQueryXML(sL_CompCode,sL_QueryDate,sL_QueryType,sL_InstQuery);

        sendRequestXml(sL_CompCode,sL_XmlString);
      end;
    end;
end;

procedure TfrmSimpleSGS.HTTPServerCommandGet(AThread: TIdPeerThread;
  ARequestInfo: TIdHTTPRequestInfo; AResponseInfo: TIdHTTPResponseInfo);
var
    sL_ResponseXml : String;
    F_FilePath : TextFile;
    nL_FileHandel : Integer;
begin
    sL_ResponseXml := TUstr.replaceStr(ARequestInfo.Params.GetText,''#$D#$A'','');

    if not FileExists('c:\ResponseXml.txt') then
    begin
      nL_FileHandel := FileCreate('c:\ResponseXml.txt');

      if (nL_FileHandel>=0) then
        FileClose(nL_FileHandel);
    end;
    AssignFile(F_FilePath,'c:\ResponseXml.txt');
    Append(F_FilePath);

    Writeln(F_FilePath,'[' +  DateTimeToStr(now)+ ']-' + sL_ResponseXml);
    Flush(F_FilePath);
    CloseFile(F_FilePath);
//    Memo1.Lines.Add(sL_ResponseXml);


    //SHOWMESSAGE(sL_ResponseXml);


    {
    ResponseInfo.ResponseText :='';
    setCaption;
    }
//    lbLog.Items.Add(sL_ResponseXml);
    //lbLog.Refresh;
//    DisplayMessage(sL_ResponseXml);
      //  DisplayMessage('xxx');

    //Memo2.Clear;
    //Memo2.Lines.Add('aaa');
end;

end.


