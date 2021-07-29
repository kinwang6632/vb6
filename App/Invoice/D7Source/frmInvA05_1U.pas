unit frmInvA05_1U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Buttons, DB, dtmReportModule,
   frxClass, OleServer;

type
  TfrmInvA05_1 = class(TForm)
    Panel1: TPanel;
    btnExit: TBitBtn;
    btnPrintNew: TBitBtn;
    rgpPrintType: TRadioGroup;
    Panel2: TPanel;
    Label1: TLabel;
    procedure btnExitClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure GenQRCodeEncode(Sender: TObject; var Data: String;
      Barcode: String);

  private
    { Private declarations }
    FCallFromProgram: TClass;
    FInvId: String;
    procedure UnPrepareDataSet;
    procedure PrepareDataSet;
    function InvDataFillToDataSet: Boolean;
    procedure AfterPrintReport;
//    procedure PrepareQRCodeData(a_InvoiceNumber,a_AESKey:String;var aQRCode1,aQRCode2:String);
//    function GetTwInvDate(aInvDate : TDateTime):String;
    function GetFullInvDate(aInvDate,aUpdTime : TDateTime):String;
//    function GetCode39Data(aInvDate:TDateTime;aInvId,aRanNum:String):String;
    procedure Deletefiles(APath, AFileSpec: string);
    function ProcQRCodePic(aText,aPathName:String):Boolean;
  public
    { Public declarations }
    sG_UserID, sG_CompID, sG_PrintTypeFile, sG_OrderByType,sG_originalType : String;
    property CallFromProgram: TClass read FCallFromProgram write FCallFromProgram;
    property InvId: String read FInvId write FInvId;
  end;

var
  frmInvA05_1: TfrmInvA05_1;
  {
  procedure QRCodeINV(a_InvoiceNumber: AnsiString;
                     a_InvoiceDate: AnsiString;
                     a_InvoiceTime: AnsiString;
                     a_RandomNumber: AnsiString;
                     af_SalesAmount: Double;
                     af_TaxAmount: Double;
                     af_TotalAmount: Double;
                     a_BuyerIdentifier: AnsiString;
                     a_RepresentIdentifier: AnsiString;
                     a_SellerIdentifier: AnsiString;
                     a_BusinessIdenti: AnsiString;
                     a_AESKey: AnsiString;
                      a_output: PAnsiChar;
                     var ai_errorCode : Integer)  STDCALL; external 'QRDLL.DLL';
  }
implementation

uses dtmMainJU, frmInvA05_3U, dtmMainU, frmMainU, Ustru, cbUtilis, frmInvA05U,
  frmInvA02_4U,
  frxDesgn, frxRes, frxChBox, frxExportXLS, frxBarcode,ZintInterface;


{$R *.dfm}

procedure TfrmInvA05_1.btnExitClick(Sender: TObject);

begin
  Close;
end;

procedure TfrmInvA05_1.FormCreate(Sender: TObject);
var
  sL_PrintType : String;
  L_Txt : TStringList;
  nL_PrintType: Integer;

begin
  sG_PrintTypeFile := IncludeTrailingPathDelimiter( ExtractFilePath(
    Application.ExeName ) ) + PRINT_TYPE_FILE;

  if FileExists( sG_PrintTypeFile ) then //拆解Text File的部份
  begin
    L_Txt := TStringList.Create;
    try
      L_Txt.LoadFromFile( sG_PrintTypeFile );
      if L_Txt.GetText <> '' then
        sL_PrintType := L_Txt.Strings[0]
      else
        sL_PrintType := '0';
    finally
      L_Txt.Free;
    end;

    nL_PrintType := StrToIntDef(sL_PrintType,0);

    if nL_PrintType = 1 then  //1=群健規格
    begin
      sL_PrintType := '1';
      nL_PrintType :=1;
    end
    else if nL_PrintType = 2 then  //2=TBC新規格
    begin
      sL_PrintType := '2';
      nL_PrintType := 2;
    end
    else  //其他都預設為預設值0
    begin
      sL_PrintType := '0';
      nL_PrintType := 0;
    end;

    rgpPrintType.ItemIndex := nL_PrintType;


  end
  else
  begin
    rgpPrintType.ItemIndex := 0;
    sL_PrintType := '0';
  end;
end;

procedure TfrmInvA05_1.FormShow(Sender: TObject);
begin
  self.Caption := frmMain.GetFormTitleString( 'A05_1', '列印發票型式' );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA05_1.Button1Click(Sender: TObject);
var
  aModalResult: Integer;
  aTxt : TStringList;
  aPath, aPrintTitleSet: String ;
  aPrintTitle, aPrintMailAddr, aPrintMemo: String;

begin

  if not ConfirmMsg( '請確認您的印表機輸出紙張大小設定是否正確。' ) then Exit;
  {}
  UnPrepareDataSet;
  PrepareDataSet;
    {}
  try
    if not InvDataFillToDataSet then
    begin
      WarningMsg( '沒有發票資料可供列印!' );
      Exit;
    end;
  except
    on E: Exception do
    begin
      ErrorMsg( Format( '%s', [E.Message] ) );
      Exit;
    end;
  end;
  try
    aPrintTitleSet := dtmMainJ.getIsPrintTitle( sG_CompID );
    if aPrintTitleSet = EmptyStr then aPrintTitle := 'Y';
    dtmReport.AddA05DataField( dtmMain.GetAutoCreateNum );

    if rgpPrintType.ItemIndex = 0 then
    begin
      aPath := IncludeTrailingPathDelimiter( ExtractFilePath(
        Application.ExeName ) ) + IncludeTrailingPathDelimiter(
        REPORT_FOLDER ) + 'FrptInvA05_1.fr3';
      dtmReport.frxMainReport.LoadFromFile( aPath );
      dtmReport.frxMainReport.Variables.Variables['列印發票抬頭'] :=
        QuotedStr( aPrintTitleSet );
      dtmReport.frxMainReport.ShowReport;
      AfterPrintReport;
      dtmReport.cdsTempory.EmptyDataSet;
    end
    else if rgpPrintType.ItemIndex = 1 then
    begin
      aPrintTitle := 'N';
      aPrintMailAddr := 'N';
      aPrintMemo := 'N';
      frmInvA05_3 := TfrmInvA05_3.Create( Application );
      try
        aModalResult := frmInvA05_3.ShowModal;
        if frmInvA05_3.ckbInvTitle.Checked then aPrintTitle := 'Y';
        if frmInvA05_3.ckbMailAddr.Checked then aPrintMailAddr := 'Y';
        if frmInvA05_3.ckbMemo1.Checked then aPrintMemo := 'Y';
      finally
        frmInvA05_3.Free;
      end;
      if aModalResult = mrOK then
      begin
        //{$IFDEF DEBUG}
        //  aPath := '..\Source\ReportTemplate\FrptInvA05_2.fr3';
        //{$ELSE}
          aPath := IncludeTrailingPathDelimiter( ExtractFilePath(
            Application.ExeName ) ) + IncludeTrailingPathDelimiter(
            REPORT_FOLDER ) + 'FrptInvA05_2.fr3';
        //{$ENDIF}
        dtmReport.frxMainReport.LoadFromFile( aPath );
        dtmReport.frxMainReport.Variables.Variables['列印發票抬頭'] :=
          QuotedStr( aPrintTitle );
        dtmReport.frxMainReport.Variables.Variables['列印發票地址'] :=
          QuotedStr( aPrintMailAddr );
        dtmReport.frxMainReport.Variables.Variables['列印備註'] :=
          QuotedStr( aPrintMemo );
        dtmReport.frxMainReport.ShowReport;
        AfterPrintReport;
        dtmReport.cdsTempory.EmptyDataSet;
      end;
    end
    else if rgpPrintType.ItemIndex = 2 then
    begin


      //{$IFDEF DEBUG}
      //  aPath := '..\Source\ReportTemplate\FrptInvA05_3.fr3';
      //{$ELSE}
        aPath := IncludeTrailingPathDelimiter( ExtractFilePath(
          Application.ExeName ) ) + IncludeTrailingPathDelimiter(
          REPORT_FOLDER ) + 'FrptInvA05_3.fr3';
      //{$ENDIF}

      dtmReport.frxMainReport.LoadFromFile( aPath );
      dtmReport.frxMainReport.Variables.Variables['列印發票抬頭'] :=
        QuotedStr( aPrintTitleSet );
      dtmReport.frxMainReport.ShowReport;
      AfterPrintReport;
      dtmReport.cdsTempory.EmptyDataSet;
    end;
    aTxt := TStringList.Create;
    try
      case rgpPrintType.ItemIndex of
        0: aTxt.Add( '0' );    //TBC格式
        1: aTxt.Add( '1' );     //群健格式
        2: aTxt.Add( '2' );    //TBC新格式
      end;
      if FileExists( sG_PrintTypeFile ) then DeleteFile( sG_PrintTypeFile );
      aTxt.SaveToFile( sG_PrintTypeFile );
    finally
      aTxt.Free;
    end;
  finally
    aPath := IncludeTrailingPathDelimiter( ExtractFilePath(
          Application.ExeName ) ) + IncludeTrailingPathDelimiter(
          REPORT_FOLDER ) + IncludeTrailingPathDelimiter(
          REPORT_QRCODEPIC );
    Deletefiles(aPath,'*.png');

  end;
end;

{ ---------------------------------------------------------------------------- }


procedure TfrmInvA05_1.UnPrepareDataSet;
begin
  if not VarIsNull( dtmReport.cdsTempory.Data ) then
    dtmReport.cdsTempory.Data := Null;
  dtmReport.cdsTempory.FieldDefs.Clear;

end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA05_1.PrepareDataSet;
var
  aAutoCreateNum, aIndex: Integer;
begin
  dtmReport.cdsTempory.FieldDefs.Clear;
  dtmReport.cdsTempory.FieldDefs.Add( 'CHECKNO', ftString, 2 );
  dtmReport.cdsTempory.FieldDefs.Add( 'INVID', ftString, 10 );
  dtmReport.cdsTempory.FieldDefs.Add( 'BUSINESSID', ftString, 8 );
  dtmReport.cdsTempory.FieldDefs.Add( 'INVDATE', ftDateTime );
  dtmReport.cdsTempory.FieldDefs.Add( 'CUSTID', ftString, 100 );
  dtmReport.cdsTempory.FieldDefs.Add( 'CUSTSNAME', ftString, 100 );
  dtmReport.cdsTempory.FieldDefs.Add( 'INVTITLE', ftString, 100 );
  dtmReport.cdsTempory.FieldDefs.Add( 'ZIPCODE', ftString, 100 );
  dtmReport.cdsTempory.FieldDefs.Add( 'MAILADDR', ftString, 100 );
  dtmReport.cdsTempory.FieldDefs.Add( 'INVADDR', ftString, 100 );
  dtmReport.cdsTempory.FieldDefs.Add( 'INSTADDR', ftString, 100 );
  dtmReport.cdsTempory.FieldDefs.Add( 'INVFORMAT', ftString, 100 );
  dtmReport.cdsTempory.FieldDefs.Add( 'TAXTYPE', ftString, 100 );
  dtmReport.cdsTempory.FieldDefs.Add( 'SALEAMOUNT', ftInteger );
  dtmReport.cdsTempory.FieldDefs.Add( 'TAXAMOUNT', ftInteger );
  dtmReport.cdsTempory.FieldDefs.Add( 'INVAMOUNT', ftInteger );
  {}
  dtmReport.cdsTempory.FieldDefs.Add( 'MAININVID', ftString, 10 );
  dtmReport.cdsTempory.FieldDefs.Add( 'MAINSALEAMOUNT', ftInteger );
  dtmReport.cdsTempory.FieldDefs.Add( 'MAINTAXAMOUNT', ftInteger );
  dtmReport.cdsTempory.FieldDefs.Add( 'MAININVAMOUNT', ftInteger );
  {}
  aAutoCreateNum := dtmMain.GetAutoCreateNum;
  if aAutoCreateNum <= 0 then aAutoCreateNum := 5;
  {}
  for aIndex := 1 to aAutoCreateNum do
  begin
    dtmReport.cdsTempory.FieldDefs.Add( Format( 'BILLID%d', [aIndex] ), ftString, 15 );
    dtmReport.cdsTempory.FieldDefs.Add( Format( 'DESCRIPTION%d', [aIndex] ), ftString, 200 );
    dtmReport.cdsTempory.FieldDefs.Add( Format( 'STARTDATE%d', [aIndex] ), ftDateTime );
    dtmReport.cdsTempory.FieldDefs.Add( Format( 'ENDDATE%d', [aIndex] ), ftDateTime );
    dtmReport.cdsTempory.FieldDefs.Add( Format( 'QUANTITY%d', [aIndex] ), ftInteger );
    dtmReport.cdsTempory.FieldDefs.Add( Format( 'UNITPRICE%d', [aIndex] ), ftFloat );  //#4370 改成浮點數 By Kin 2009/02/19
    dtmReport.cdsTempory.FieldDefs.Add( Format( 'SALEAMOUNT%d', [aIndex] ), ftInteger );
    dtmReport.cdsTempory.FieldDefs.Add( Format( 'TOTALAMOUNT%d', [aIndex] ), ftInteger );
    dtmReport.cdsTempory.FieldDefs.Add( Format( 'FACISNO%d', [aIndex] ), ftString, 20 );
  end;
  {}
  dtmReport.cdsTempory.FieldDefs.Add( 'MEMO1', ftString, 60 );
  dtmReport.cdsTempory.FieldDefs.Add( 'ACCOUNTNO', ftString, 1000 );
  dtmReport.cdsTempory.FieldDefs.Add( 'HOWTOCREATE', ftString, 1 );
  dtmReport.cdsTempory.FieldDefs.Add( 'INVUSEDESC', ftString, 20 );
  {}
  {#6600 增加發票隨機碼 By Kin 2013/10/28}
  dtmReport.cdsTempory.FieldDefs.Add( 'RANDOMNUM', ftString, 50 );
  {#6629 增加申報格式 By Kin 2013/10/28}
  dtmReport.cdsTempory.FieldDefs.Add( 'APPLYFORMAT' ,ftString, 50);
  {#7060 增加臨櫃信用卡末4碼 By Kin 2015/08/18}
  dtmReport.cdsTempory.FieldDefs.Add( 'Creditcard4d', ftString,10 );
  dtmReport.cdsTempory.FieldDefs.Add( 'QRCodeChineseInvDate', ftString,30 );
  dtmReport.cdsTempory.FieldDefs.Add( 'QRCodeInvDate', ftString,30 );
  dtmReport.cdsTempory.FieldDefs.Add( 'QRCodeSellerBUSINESSID', ftString,30 );
  dtmReport.cdsTempory.FieldDefs.Add( 'QRCodeBuyerBUSINESSID', ftString,30 );
  dtmReport.cdsTempory.FieldDefs.Add( 'CODE39Data', ftString,30 );
  dtmReport.cdsTempory.FieldDefs.Add( 'QRCodePngPath1', ftString,200 );
  dtmReport.cdsTempory.FieldDefs.Add( 'QRCodePngPath2', ftString,200 );
  dtmReport.cdsTempory.FieldDefs.Add( 'QRCodeInvID', ftString,30 );
  dtmReport.cdsTempory.FieldDefs.Add( 'QRCodeRanNum', ftString,30 );
  dtmReport.cdsTempory.FieldDefs.Add( 'QRCodeTotalAmount', ftString,30 );
  dtmReport.cdsTempory.FieldDefs.Add( 'CanModify',ftString,30);
  dtmReport.cdsTempory.FieldDefs.Add( 'QRCodeEFormat',ftString,30);
  dtmReport.cdsTempory.FieldDefs.Add( 'CompSName',ftString,50);
  {#7128}
  dtmReport.cdsTempory.FieldDefs.Add( 'CompId',ftString,6);
  dtmReport.cdsTempory.FieldDefs.Add( 'CompName',ftString,100);
  dtmReport.cdsTempory.FieldDefs.Add( 'CompBusinessId',ftString,20);
  dtmReport.cdsTempory.FieldDefs.Add( 'CompAddr',ftString,200);
  dtmReport.cdsTempory.FieldDefs.Add( 'CompTel',ftString,20);
  dtmReport.cdsTempory.FieldDefs.Add( 'PrizeType',ftString,5);
  dtmReport.cdsTempory.CreateDataSet;

     {
  COMPID=公司別
COMPNAME=公司全名
COMPBUSINESSID=公司統編
COMPADDR=公司營業地址
COMPTEL=公司電話
PRIZETYPE=中獎記錄
 }

  {
 QRCodeChineseInvDate=QRCODE民國年發票日期
QRCodeInvDate=QRCode發票日期時間
QRCodeSellerBUSINESSID=賣方統一編號
CODE39Data=CDDE39資料內容
QRCodePngPath1=QRCode1圖檔路徑
QRCodePngPath2=QRCode2圖檔路徑
QRCodeInvID=QRCode發票號碼
QRCodeRanNum=QRCode隨機碼
QRCodeTotalAmount=QRCode總計金額
QRCodeBuyerBUSINESSID=買方統一編號
}
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA05_1.InvDataFillToDataSet: Boolean;
var
  aSQL, aLastInvId, aDescription, aDaySt, aDayEd, aFacisNo: String;
  aCurrentFieldNumber: Integer;

  aPath : String;
  aInvDateRage : String;
  aQRCode1,aQRCode2 :String;

begin
  {#6600 增加發票防偽隨機碼 By Kin 2013/10/02}
  {#6966 ApplyFormat 改讀 Inv001.EFormat By Kin 2014/12/18}
  {#7060 用卡後四碼，若"臨櫃信用卡末四碼"有值，以此優先，沒有才取INV008.ACCOUNTNO的後四碼 By Kin 2015/08/18}
  if CallFromProgram = TfrmInvA05 then
  begin
    aSQL := Format(
      ' SELECT A.CHECKNO, A.INVID, A.BUSINESSID, A.INVDATE, A.CUSTID,                   ' +
      '        A.CUSTSNAME, A.INVFORMAT, A.INVTITLE, A.ZIPCODE,                         ' +
      '        A.MAILADDR, A.TAXTYPE, A.SALEAMOUNT, A.TAXAMOUNT,                        ' +
      '        A.INVAMOUNT, A.MEMO1, B.DESCRIPTION, B.STARTDATE,                        ' +
      '        B.ENDDATE, B.QUANTITY, B.UNITPRICE, B.TOTALAMOUNT,                       ' +
      '        B.SEQ, A.MAININVID, A.INVADDR,                                           ' +
      '        A.MAINSALEAMOUNT, A.MAINTAXAMOUNT, A.MAININVAMOUNT,                      ' +
      '        B.SALEAMOUNT AS SALEAMOUNT2, B.BILLID, A.INSTADDR,                       ' +
      '        A.HOWTOCREATE, A.INVUSEDESC, C.IFPRINTTITLE,                             ' +
      '        A.RANDOMNUM, A.FIXFLAG,A.UPTTIME,CANMODIFY,A.PRIZETYPE,                  ' +
      ' (SELECT EFORMAT FROM INV001 WHERE COMPID = A.COMPID ) EFORMAT,                  ' +
      ' (SELECT COMPSNAME FROM INV001 WHERE COMPID = A.COMPID ) COMPSNAME,              ' +
      ' (SELECT COMPID FROM INV001 WHERE COMPID = A.COMPID ) AS COMPID,                 ' +
      ' (SELECT COMPNAME FROM INV001 WHERE COMPID = A.COMPID ) AS COMPNAME,             ' +
      ' (SELECT BUSINESSID FROM INV001 WHERE COMPID = A.COMPID ) AS COMPBUSINESSID,     ' +
      ' (SELECT COMPADDR FROM INV001 WHERE COMPID = A.COMPID ) AS COMPADDR,             ' +
      ' (SELECT TEL FROM INV001 WHERE COMPID = A.COMPID ) AS COMPTEL,                   ' +
      ' (SELECT Decode(BusinessId,Null,MainBusinessId,BusinessId) FROM INV001 WHERE COMPID = A.COMPID ) MainBusinessId,    ' +
      ' Decode(substr(B.Creditcard4d,-4,999),null,substr(B.AccountNo,-4,999),substr( B.Creditcard4d,-4,999))  Creditcard4d ' +
      '   FROM INV023 A, INV008 B, INV002 C                           ' +
      '  WHERE A.IDENTIFYID1 = B.IDENTIFYID1                          ' +
      '    AND A.IDENTIFYID2 = B.IDENTIFYID2                          ' +
      '    AND A.INVID = B.INVID                                      ' +
      '    AND A.IDENTIFYID1 = C.IDENTIFYID1(+)                       ' +
      '    AND A.IDENTIFYID2 = C.IDENTIFYID2(+)                       ' +
      '    AND A.COMPID = C.COMPID(+)                                 ' +
      '    AND A.CUSTID = C.CUSTID(+)                                 ' +
      '    AND A.IDENTIFYID1 = ''%s''                                 ' +
      '    AND A.IDENTIFYID2 = ''%s''                                 ' +
      '    AND A.COMPID = ''%s''                                      ' +
      '    AND A.OWNER = ''%s''                                       ',
      [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, sG_UserID ] );
      if sG_OrderByType = '1' then
        aSQL := ( aSQL + ' ORDER BY A.INVID, B.SEQ ' )
      else if sG_OrderByType = '2' then
        aSQL := ( aSQL + ' ORDER BY A.INVDATE, A.INVID, B.SEQ ' )
      else if sG_OrderByType = '3' then
        aSQL := ( aSQL + ' ORDER BY A.CUSTID, A.INVID, B.SEQ ' );
  end else
  if CallFromProgram = TfrmInvA02_4 then
  begin
    aSQL := Format(
     '  SELECT A.CHECKNO, A.INVID, A.BUSINESSID, A.INVDATE, A.CUSTID,               ' +
     '         A.CUSTSNAME, A.INVFORMAT, A.INVTITLE, A.ZIPCODE,                     ' +
     '         A.MAILADDR, A.TAXTYPE, A.SALEAMOUNT, A.TAXAMOUNT,                    ' +
     '         A.INVAMOUNT, A.MEMO1, B.DESCRIPTION, B.STARTDATE,                    ' +
     '         B.ENDDATE, B.QUANTITY, B.UNITPRICE, B.TOTALAMOUNT,                   ' +
     '         B.SEQ, A.MAININVID, A.INVADDR,                                       ' +
     '         A.MAINSALEAMOUNT, A.MAINTAXAMOUNT, A.MAININVAMOUNT,                  ' +
     '         B.SALEAMOUNT AS SALEAMOUNT2, B.BILLID, A.INSTADDR,                   ' +
     '         A.HOWTOCREATE, A.INVUSEDESC, D.IFPRINTTITLE,                         ' +
     '         A.RANDOMNUM, A.FIXFLAG,A.UPTTIME,CANMODIFY,A.PRIZETYPE,              ' +
     ' (SELECT EFORMAT FROM INV001 WHERE COMPID = A.COMPID ) EFORMAT,               ' +
     ' (SELECT COMPSNAME FROM INV001 WHERE COMPID = A.COMPID ) COMPSNAME,           ' +
     ' (SELECT COMPID FROM INV001 WHERE COMPID = A.COMPID ) AS COMPID,              ' +
     ' (SELECT COMPNAME FROM INV001 WHERE COMPID = A.COMPID ) AS COMPNAME,          ' +
     ' (SELECT BUSINESSID FROM INV001 WHERE COMPID = A.COMPID ) AS COMPBUSINESSID,  ' +
     ' (SELECT COMPADDR FROM INV001 WHERE COMPID = A.COMPID ) AS COMPADDR,          ' +
     ' (SELECT TEL FROM INV001 WHERE COMPID = A.COMPID ) AS COMPTEL,                ' +
     ' (SELECT Decode(BusinessId,Null,MainBusinessId,BusinessId) FROM INV001 WHERE COMPID = A.COMPID ) MainBusinessId,    ' +
     ' Decode(substr(B.Creditcard4d,-4,999),null,substr(B.AccountNo,-4,999),substr( B.Creditcard4d,-4,999))  Creditcard4d ' +
     '    FROM INV007 A, INV008 B,                                    ' +
     '    ( SELECT MAININVID FROM INV007                              ' +
     '       WHERE IDENTIFYID1 = ''%s''                               ' +
     '         AND IDENTIFYID2 = ''%s''                               ' +
     '         AND COMPID = ''%s''                                    ' +
     '         AND INVID = ''%s''                                     ' +
     '     ) C, INV002 D                                              ' +
     '   WHERE A.IDENTIFYID1 = B.IDENTIFYID1                          ' +
     '     AND A.IDENTIFYID2 = B.IDENTIFYID2                          ' +
     '     AND A.INVID = B.INVID                                      ' +
     '     AND A.MAININVID = C.MAININVID                              ' +
     '     AND A.IDENTIFYID1 = D.IDENTIFYID1(+)                       ' +
     '     AND A.IDENTIFYID2 = D.IDENTIFYID2(+)                       ' +
     '     AND A.COMPID = D.COMPID(+)                                 ' +
     '     AND A.CUSTID = D.CUSTID(+)                                 ' +
     '     AND A.IDENTIFYID1 = ''%s''                                 ' +
     '     AND A.IDENTIFYID2 = ''%s''                                 ' +
     '     AND A.COMPID = ''%s''                                      ' +
     '  ORDER BY B.INVID, B.SEQ                                       ',
     [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, FInvId,
      IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID] );
  end;
  {}
  dtmMain.adoComm.Close;
  dtmMain.adoComm.SQL.Text := aSQL;
  dtmMain.adoComm.Open;
  Result := not dtmMain.adoComm.IsEmpty;
  if not Result then
  begin
    dtmMain.adoComm.Close;
    Exit;
  end;
  aLastInvId := EmptyStr;
  aCurrentFieldNumber := 0;
  dtmMain.adoComm.First;
  while not dtmMain.adoComm.Eof do
  begin
    if ( aLastInvId <> dtmMain.adoComm.FieldByName( 'INVID' ).AsString ) then
    begin
      aPath := IncludeTrailingPathDelimiter( ExtractFilePath(
          Application.ExeName ) ) + IncludeTrailingPathDelimiter(
          REPORT_FOLDER ) + IncludeTrailingPathDelimiter(
          REPORT_QRCODEPIC );
     if not DirectoryExists( apath ) then
     begin
       CreateDir( aPath );
     end;
      aQRCode1 := '';
      aQRCode2 := '';
      {#7076 Produce QrCode-data By Kin 2015/09/10}
      dtmMainJ.PrepareQRCodeData(dtmMain.adoComm.FieldByName('INVID').AsString,
                                        dtmMain.sG_AESKey,aQRCode1,aQRCode2);
      try

        if not ProcQRCodePic( aQRCode1 , aPath + dtmMain.adoComm.FieldByName( 'INVID' ).AsString + '_1.Png') then
          Exit;
        if not ProcQRCodePic( aQRCode2 , aPath + dtmMain.adoComm.FieldByName( 'INVID' ).AsString + '_2.Png') then
          Exit;
        {
        GenQRCode.Barcode := aQRCode1;
        SaveQRCode.Save(aPath + dtmMain.adoComm.FieldByName( 'INVID' ).AsString + '_1.Png' );
        GenQRCode.Barcode := EmptyStr;
        GenQRCode.Barcode := aQRCode2;
        SaveQRCode.Save(aPath + dtmMain.adoComm.FieldByName( 'INVID' ).AsString + '_2.Png' );
        }
      except
         raise Exception.Create( '產生QRCode圖檔有誤！' );
      end;

      aLastInvId := dtmMain.adoComm.FieldByName( 'INVID' ).AsString;
      aCurrentFieldNumber := 1;
      dtmReport.cdsTempory.Append;
    end else
    begin
      Inc( aCurrentFieldNumber );
      dtmReport.cdsTempory.Edit;
    end;
    if ( aCurrentFieldNumber <= 1 ) then
    begin
      dtmReport.cdsTempory.FieldByName( 'CHECKNO' ).AsString :=
        dtmMain.adoComm.FieldByName( 'CHECKNO' ).AsString;
      {}
      dtmReport.cdsTempory.FieldByName( 'INVID' ).AsString :=
        dtmMain.adoComm.FieldByName( 'INVID' ).AsString;
      {}
      dtmReport.cdsTempory.FieldByName( 'BUSINESSID' ).AsString :=
        dtmMain.adoComm.FieldByName( 'BUSINESSID' ).AsString;
      {}
      dtmReport.cdsTempory.FieldByName( 'INVDATE' ).Value :=
        dtmMain.adoComm.FieldByName( 'INVDATE' ).Value;
      {}
      dtmReport.cdsTempory.FieldByName( 'CUSTID' ).AsString :=
        dtmMain.adoComm.FieldByName( 'CUSTID' ).AsString;
      {}
      dtmReport.cdsTempory.FieldByName( 'CUSTSNAME' ).AsString :=
        dtmMain.adoComm.FieldByName( 'CUSTSNAME' ).AsString;
      {}
      { 不論統一編號是否有值, 都填入 INVTITLE }
      if dtmMain.GetIfPrintTitle then
        dtmReport.cdsTempory.FieldByName( 'INVTITLE' ).AsString :=
          dtmMain.adoComm.FieldByName( 'INVTITLE' ).AsString
      else begin
        { 統一編號有值, 才填入 INVTITLE }
        dtmReport.cdsTempory.FieldByName( 'INVTITLE' ).AsString := EmptyStr;
        if ( dtmMain.adoComm.FieldByName( 'BUSINESSID' ).AsString <> EmptyStr ) then
          dtmReport.cdsTempory.FieldByName( 'INVTITLE' ).AsString :=
            dtmMain.adoComm.FieldByName( 'INVTITLE' ).AsString
        { 若統編沒值, 但是 INV002 的 IfPrintTitle = Y 仍然要印 }
        else if ( dtmMain.adoComm.FieldByName( 'IFPRINTTITLE' ).AsString = 'Y' ) then
          dtmReport.cdsTempory.FieldByName( 'INVTITLE' ).AsString :=
            dtmMain.adoComm.FieldByName( 'INVTITLE' ).AsString;
      end;
      {#6629 增加申請格式 By Kin 2013/10/28}
      {#發票格式多增加判斷FIXFLAG = 1 By Kin 2013/12/19 For Jacy}
      dtmReport.cdsTempory.FieldByName( 'APPLYFORMAT' ).AsString := EmptyStr;
      if dtmMain.sG_StarComBine then
      begin
        if ( dtmMain.adoComm.FieldByName( 'BUSINESSID' ).AsString <> EmptyStr ) and
          ( dtmMain.adoComm.FieldByName( 'FIXFLAG' ).AsInteger = 1 ) then
        begin
          { #6966 原本寫死21,改為讀取EFORMAT By Kin 2014/12/18}
          dtmReport.cdsTempory.FieldByName( 'APPLYFORMAT' ).AsString :=
              '電子發票格式代號:' + dtmMain.adoComm.FieldByName( 'EFORMAT' ).AsString;
        end;
      end;
      {}
      dtmReport.cdsTempory.FieldByName( 'ZIPCODE' ).AsString :=
        dtmMain.adoComm.FieldByName( 'ZIPCODE' ).AsString;
      // 郵寄地址
      dtmReport.cdsTempory.FieldByName( 'MAILADDR' ).AsString :=
        dtmMain.adoComm.FieldByName( 'MAILADDR' ).AsString;
      // 發票地址
      dtmReport.cdsTempory.FieldByName( 'INVADDR' ).AsString :=
        dtmMain.adoComm.FieldByName( 'INVADDR' ).AsString;
      // 裝機地址
      dtmReport.cdsTempory.FieldByName( 'INSTADDR' ).AsString :=
        dtmMain.adoComm.FieldByName( 'INSTADDR' ).AsString;
      {}
      dtmReport.cdsTempory.FieldByName( 'INVFORMAT' ).AsString :=
        dtmMain.adoComm.FieldByName( 'INVFORMAT' ).AsString;
      {}
      dtmReport.cdsTempory.FieldByName( 'TAXTYPE' ).AsString :=
        dtmMain.adoComm.FieldByName( 'TAXTYPE' ).AsString;
      {}
      dtmReport.cdsTempory.FieldByName( 'SALEAMOUNT' ).AsInteger :=
        dtmMain.adoComm.FieldByName( 'SALEAMOUNT' ).AsInteger;
      {}
      dtmReport.cdsTempory.FieldByName( 'TAXAMOUNT' ).AsInteger :=
        dtmMain.adoComm.FieldByName( 'TAXAMOUNT' ).AsInteger;
      {}
      dtmReport.cdsTempory.FieldByName( 'INVAMOUNT' ).AsInteger :=
        dtmMain.adoComm.FieldByName( 'INVAMOUNT' ).AsInteger;

      dtmReport.cdsTempory.FieldByName( 'MAININVID' ).AsString :=
        dtmMain.adoComm.FieldByName( 'MAININVID' ).AsString;
      {}
      dtmReport.cdsTempory.FieldByName( 'MAINSALEAMOUNT' ).AsInteger :=
        dtmMain.adoComm.FieldByName( 'MAINSALEAMOUNT' ).AsInteger;
      {}
      dtmReport.cdsTempory.FieldByName( 'MAINTAXAMOUNT' ).AsInteger :=
        dtmMain.adoComm.FieldByName( 'MAINTAXAMOUNT' ).AsInteger;
      {}
      dtmReport.cdsTempory.FieldByName( 'MAININVAMOUNT' ).AsInteger :=
        dtmMain.adoComm.FieldByName( 'MAININVAMOUNT' ).AsInteger;
      {}
      dtmReport.cdsTempory.FieldByName( 'MEMO1' ).AsString :=
        dtmMain.adoComm.FieldByName( 'MEMO1' ).AsString;
      {}
      dtmReport.cdsTempory.FieldByName( 'ACCOUNTNO' ).AsString := dtmMainJ.GetAccountNo(
        dtmReport.cdsTempory.FieldByName( 'INVID' ).AsString,
        EmptyStr, EmptyStr, EmptyStr );
      {}
      dtmReport.cdsTempory.FieldByName( 'HOWTOCREATE' ).AsString :=
        dtmMain.adoComm.FieldByName( 'HOWTOCREATE' ).AsString;
      {}
      dtmReport.cdsTempory.FieldByName( 'INVUSEDESC' ).AsString :=
        dtmMain.adoComm.FieldByName( 'INVUSEDESC' ).AsString;
      {#6600 增加發票隨機碼 By Kin 2013/10/02}
      {#6629 增加電子發票隨機碼: By Kin 2013/10/30}
      {隨機碼要判斷沒有統編而且FixFlag =1 才可列印 By Kin 2013/12/19 For Jacy}
       dtmReport.cdsTempory.FieldByName( 'RANDOMNUM' ).AsString := EmptyStr;
      if dtmMain.adoComm.FieldByName( 'BUSINESSID' ).AsString = '' then
      begin
        if dtmMain.adoComm.FieldByName( 'FIXFLAG' ).AsInteger = 1  then
        begin
          if dtmMain.adoComm.FieldByName( 'RANDOMNUM' ).AsString <> EmptyStr then
          begin
            dtmReport.cdsTempory.FieldByName( 'RANDOMNUM' ).AsString :=
              '電子發票隨機碼:' + dtmMain.adoComm.FieldByName( 'RANDOMNUM' ).AsString;
          end;
        end;
      end;
      {#7076 增加臨櫃信用卡末4碼 By Kin 2015/08/18}
      {#7076 Add QRCode-Data of field into Report By Kin 2015/08/18}
      dtmReport.cdsTempory.FieldByName( 'Creditcard4d' ).AsString :=
        dtmMain.adoComm.FieldByName( 'Creditcard4d' ).AsString;

     dtmReport.cdsTempory.FieldByName( 'QRCodeChineseInvDate' ).AsString :=
        GetTwInvDate(  dtmMain.adoComm.FieldByName( 'InvDate' ).AsDateTime );


     dtmReport.cdsTempory.FieldByName( 'QRCodeInvID' ).AsString :=
        format('%s-%s',
          [copy(dtmMain.adoComm.FieldByName('INVID').AsString,1,2),
          copy(dtmMain.adoComm.FieldByName('INVID').AsString,3,8)]);

     dtmReport.cdsTempory.FieldByName( 'QRCodeInvDate' ).AsString :=
        GetFullInvDate(dtmMain.adoComm.FieldByName( 'InvDate' ).AsDateTime,
                      dtmMain.adoComm.FieldByName( 'UptTime' ).AsDateTime);

      dtmReport.cdsTempory.FieldByName( 'QRCodeRanNum' ).AsString :=
        Format('隨機碼 %s',[dtmMain.adoComm.FieldByName('RANDOMNUM').AsString]);

      dtmReport.cdsTempory.FieldByName( 'QRCodeTotalAmount' ).AsString :=
        Format('總計 %s',[dtmMain.adoComm.FieldByName('InvAmount').AsString]);
      dtmReport.cdsTempory.FieldByName( 'QRCodeSellerBUSINESSID' ).AsString :=
        format('賣方%s',[dtmMain.adoComm.FieldByName('MAINBUSINESSID').AsString]);
      {#7076 If Invoice has BUSINESSID data then writing QRCodeEFormat data By Kin   }
      if ( dtmMain.adoComm.FieldByName( 'BUSINESSID' ).AsString <> EmptyStr ) then
      begin
        dtmReport.cdsTempory.FieldByName( 'QRCodeEFormat' ).AsString :=
                  Format('格式 %S',[ dtmMain.adoComm.FieldByName( 'EFORMAT' ).AsString ] );
      end else
      begin
        dtmReport.cdsTempory.FieldByName( 'QRCodeEFormat' ).AsString := '';
      end;
      {#7076 To Add CompSName Field into report By Kin 2015/11/05 }
      if ( dtmMain.adoComm.FieldByName( 'CompSName' ).AsString <> EmptyStr ) then
      begin
        dtmReport.cdsTempory.FieldByName( 'CompSName' ).AsString :=
                  dtmMain.adoComm.FieldByName( 'CompSName' ).AsString;
      end;
      {#7076 Add CanModify Parameter makes qrcode's text visable or enabled By Kin}
      if dtmMain.adoComm.FieldByName( 'CANMODIFY' ).AsString <> '' then
      begin
        dtmReport.cdsTempory.FieldByName( 'CanModify' ).AsString :=
            UpperCase(dtmMain.adoComm.FieldByName( 'CANMODIFY' ).AsString);
      end else
      begin
        dtmReport.cdsTempory.FieldByName( 'CanModify' ).AsString := 'N';
      end;
      if dtmMain.adoComm.FieldByName('BUSINESSID').AsString = EmptyStr then
      begin
        dtmReport.cdsTempory.FieldByName( 'QRCodeBuyerBUSINESSID' ).AsString :='';
      end else
      begin
        dtmReport.cdsTempory.FieldByName( 'QRCodeBuyerBUSINESSID' ).AsString :=
            format('買方%s',[dtmMain.adoComm.FieldByName('BUSINESSID').AsString]);
      end;

      dtmReport.cdsTempory.FieldByName( 'CODE39Data' ).AsString :=
      dtmMainJ.GetCode39Data(dtmMain.adoComm.FieldByName('INVDATE').AsDateTime,
        dtmMain.adoComm.FieldByName('INVID').AsString,
        dtmMain.adoComm.FieldByName('RANDOMNUM').AsString);
      dtmReport.cdsTempory.FieldByName( 'QRCodePngPath1' ).AsString :=
        aPath + dtmMain.adoComm.FieldByName( 'INVID' ).AsString + '_1.Png';
      dtmReport.cdsTempory.FieldByName( 'QRCodePngPath2' ).AsString :=
        aPath + dtmMain.adoComm.FieldByName( 'INVID' ).AsString + '_2.Png';
    end;
    {}
    dtmReport.cdsTempory.FieldByName( Format( 'BILLID%d', [aCurrentFieldNumber] ) ).Value :=
      dtmMain.adoComm.FieldByName( 'BILLID' ).Value;
    {}
    aDescription := dtmMain.adoComm.FieldByName( 'DESCRIPTION' ).AsString;
    aDaySt := EmptyStr;
    aDayEd := EmptyStr;
    if not VarIsNull( dtmMain.adoComm.FieldByName( 'STARTDATE' ).Value ) then
      aDaySt := FormatDateTime( 'eee/mm/dd', dtmMain.adoComm.FieldByName( 'STARTDATE' ).AsDateTime );
    if not VarIsNull( dtmMain.adoComm.FieldByName( 'ENDDATE' ).Value ) then
      aDayEd := FormatDateTime( 'eee/mm/dd', dtmMain.adoComm.FieldByName( 'ENDDATE' ).AsDateTime );
    aDescription := Format( aDescription + '( %s - %s )', [aDaySt, aDayEd] );
    {}
    dtmReport.cdsTempory.FieldByName( Format( 'DESCRIPTION%d', [aCurrentFieldNumber] ) ).AsString :=
      dtmMain.adoComm.FieldByName( 'DESCRIPTION' ).AsString;
    {}
    dtmReport.cdsTempory.FieldByName( Format( 'STARTDATE%d', [aCurrentFieldNumber] ) ).Value :=
      dtmMain.adoComm.FieldByName( 'STARTDATE' ).Value;
    {}
    dtmReport.cdsTempory.FieldByName( Format( 'ENDDATE%d', [aCurrentFieldNumber] ) ).Value :=
      dtmMain.adoComm.FieldByName( 'ENDDATE' ).Value;
    {}
    dtmReport.cdsTempory.FieldByName( Format( 'QUANTITY%d', [aCurrentFieldNumber] ) ).Value :=
      dtmMain.adoComm.FieldByName( 'QUANTITY' ).AsInteger;
    {}
    dtmReport.cdsTempory.FieldByName( Format( 'UNITPRICE%d', [aCurrentFieldNumber] ) ).Value :=
      dtmMain.adoComm.FieldByName( 'UNITPRICE' ).Value;
    {}
    dtmReport.cdsTempory.FieldByName( Format( 'SALEAMOUNT%d', [aCurrentFieldNumber] ) ).Value :=
      dtmMain.adoComm.FieldByName( 'SALEAMOUNT2' ).Value;
    {}
    dtmReport.cdsTempory.FieldByName( Format( 'TOTALAMOUNT%d', [aCurrentFieldNumber] ) ).Value :=
      dtmMain.adoComm.FieldByName( 'TOTALAMOUNT' ).Value;
    {}
    {#7128 Adding fields into report by kin 2015/12/14}
    dtmReport.cdsTempory.FieldByName( 'COMPID' ).Value := dtmMain.adoComm.FieldByName( 'COMPID' ).Value;
    dtmReport.cdsTempory.FieldByName( 'COMPNAME' ).Value := dtmMain.adoComm.FieldByName( 'COMPNAME' ).Value;
    dtmReport.cdsTempory.FieldByName( 'COMPBUSINESSID' ).Value := dtmMain.adoComm.FieldByName( 'COMPBUSINESSID' ).Value;
    dtmReport.cdsTempory.FieldByName( 'COMPADDR' ).Value := dtmMain.adoComm.FieldByName( 'COMPADDR' ).Value;
    dtmReport.cdsTempory.FieldByName( 'COMPTEL' ).Value := dtmMain.adoComm.FieldByName( 'COMPTEL' ).Value;
    dtmReport.cdsTempory.FieldByName( 'PRIZETYPE' ).Value := dtmMain.adoComm.FieldByName( 'PRIZETYPE' ).Value;

    {}
    aFacisNo := dtmMainJ.GetFacisNo( dtmMain.adoComm.FieldByName( 'INVID' ).AsString,
        dtmMain.adoComm.FieldByName( 'SEQ' ).AsString );
    if ( Pos( ',', aFacisNo ) > 0 ) then aFacisNo := Copy( aFacisNo, 1, Pos( ',', aFacisNo ) - 1 );
    aFacisNo := dtmMainU.FacisNoAddMask( aFacisNo );
    dtmReport.cdsTempory.FieldByName( Format( 'FACISNO%d', [aCurrentFieldNumber] ) ).Value := aFacisNo;    
    {}
    dtmReport.cdsTempory.Post;
    dtmMain.adoComm.Next;
  end;
  dtmMain.adoComm.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA05_1.AfterPrintReport;
begin
  if FCallFromProgram = TfrmInvA05 then
  begin
    dtmMainJ.AfterPrintOrTrans( 1, nil,False );
  end else
  if FCallFromProgram = TfrmInvA02_4 then
  begin
    dtmMainJ.AfterPrintOrTrans( 2, dtmReport.cdsTempory,False );
  end;
end;

{ ---------------------------------------------------------------------------- }

{
procedure TfrmInvA05_1.PrepareQRCodeData(a_InvoiceNumber,a_AESKey:String;var aQRCode1,aQRCode2:String);
  var
  Response: Array[0..1023] of Char;
  errorCode : Integer;
  aSQL:string;
  a_InvoiceDate, a_InvoiceTime, a_RandomNumber:String;
  af_SalesAmount, af_TaxAmount, af_TotalAmount:Double;
  a_BuyerIdentifier, a_SellerIdentifier : String;
  aResult : String;
  RecCntTotal,DetailRecTotal :Integer;
  DetailListString: String;
  DetailListString2: String;
begin
  try
    try
      FillChar(Response, Sizeof(Response), 0);
      aSQL := Format('SELECT A.*,B.Description,B.Quantity,B.UnitPrice,      ' +
              ' (SELECT Decode(BusinessId,Null,MainBusinessId,BusinessId) FROM INV001 WHERE COMPID = A.COMPID ) MainBusinessId    ' +
              ' FROM INV007 A,INV008 B                                      ' +
              ' WHERE A.INVID =''%s''                                       ' +
              ' AND A.INVID = B.INVID                                       ',
              [a_InvoiceNumber]);
      dtmMain.adoComm2.Close;
      dtmMain.adoComm2.SQL.Text := aSQL;
      dtmMain.adoComm2.Open;
      a_InvoiceDate := getYearMonthDay7(dtmMain.adoComm2.FieldByName( 'INVDATE' ).AsDateTime,0 );
      a_InvoiceTime := StringReplace( FormatDateTime('hh:nn:ss', dtmMain.adoComm2.FieldByName( 'UPTTIME' ).AsDateTime ),
                                 ':','',[rfReplaceAll, rfIgnoreCase]);
      a_RandomNumber := dtmMain.adoComm2.FieldByName( 'RandomNum' ).AsString;
      if a_RandomNumber = EmptyStr then a_RandomNumber := '0000';
      af_SalesAmount := dtmMain.adoComm2.FieldByName( 'SaleAmount' ).AsFloat;
      af_TaxAmount := dtmMain.adoComm2.FieldByName( 'TaxAmount' ).AsFloat;
      af_TotalAmount := dtmMain.adoComm2.FieldByName( 'InvAmount' ).AsFloat;
      a_BuyerIdentifier := '00000000';
      a_SellerIdentifier := '00000000';
      if dtmMain.adoComm2.FieldByName( 'BusinessId' ).AsString <> EmptyStr then
      begin
        a_BuyerIdentifier := dtmMain.adoComm2.FieldByName( 'BusinessId' ).AsString;
      end;
      if dtmMain.adoComm2.FieldByName( 'MainBusinessId' ).AsString <> EmptyStr then
      begin
        a_SellerIdentifier := dtmMain.adoComm2.FieldByName( 'MainBusinessId' ).AsString;
      end;
      errorCode := -1;
      QRCodeINV(a_InvoiceNumber, a_InvoiceDate,
            a_RandomNumber, a_RandomNumber,af_SalesAmount,
             af_TaxAmount,af_TotalAmount,
             a_BuyerIdentifier, a_BuyerIdentifier,
              a_SellerIdentifier, a_SellerIdentifier, a_AESKey,Response,errorCode);
      if errorCode <> 0 then
      begin
        raise Exception.Create( Format('取出QRCodeData資料有誤！,ErrorCode = %s',[ IntToStr( errorCode)] ));
        Exit;
      end;
      aResult := StrPas(@Response[0]);
      aResult :=  aResult +':**********:';
      RecCntTotal := 0;
      detailRecTotal := 0;
      dtmMain.adoComm2.First;
      DetailListString2 := '**';
      while not dtmMain.adoComm2.Eof do
      begin
        if RecCntTotal = 0 then
        begin
          DetailListString := Format(':%s:%s:%s',
                                [dtmMain.adoComm2.FieldByName( 'Description' ).AsString,
                                dtmMain.adoComm2.FieldByName( 'Quantity' ).AsString,
                                dtmMain.adoComm2.FieldByName( 'UnitPrice' ).AsString]);
        end else
        begin
          if DetailListString2 = '**' then
          begin
            DetailListString2 := DetailListString2 + Format('%s:%s:%s',
                                  [dtmMain.adoComm2.FieldByName( 'Description' ).AsString,
                                  dtmMain.adoComm2.FieldByName( 'Quantity' ).AsString,
                                 dtmMain.adoComm2.FieldByName( 'UnitPrice' ).AsString]);
          end else
          begin
            DetailListString2 := DetailListString + Format(':%s:%s:%s',
                                  [dtmMain.adoComm2.FieldByName( 'Description' ).AsString,
                                  dtmMain.adoComm2.FieldByName( 'Quantity' ).AsString,
                                 dtmMain.adoComm2.FieldByName( 'UnitPrice' ).AsString]);
          end;
        end;

        detailRecTotal := detailRecTotal + dtmMain.adoComm2.FieldByName( 'Quantity' ).AsInteger;
        RecCntTotal := RecCntTotal + 1;
        dtmMain.adoComm2.Next;
      end;
      if DetailListString2 <> '**' then DetailListString := DetailListString + ':';

      aResult := aResult + IntToStr(RecCntTotal) + ':' + IntToStr(DetailRecTotal) + ':1' + DetailListString;
      aQRCode1 := aResult;
      aQRCode2 := DetailListString2;

//      Result := aResult;
    except
       raise Exception.Create( '取出QRCodeData資料有誤！' );
    end;
  finally
     dtmMain.adoComm2.SQL.Clear;
     dtmMain.adoComm2.Close;
  end;
end;
}
procedure TfrmInvA05_1.GenQRCodeEncode(Sender: TObject; var Data: String;
  Barcode: String);
begin
  Data := #$EF#$BB#$BF + AnsiToUTF8(Barcode);
end;
{
function TfrmInvA05_1.GetTwInvDate(aInvDate: TDateTime): String;
 var
  aInvDateRage : String;
  aFmtInvDate : String;
begin
  aFmtInvDate := getYearMonthDay(DateTimeToStr(aInvDate));
  case StrToInt( copy(aFmtInvDate,5,2)) of
       1 : aInvDateRage := '01-02' ;
       2 : aInvDateRage := '01-02' ;
       3 : aInvDateRage := '03-04' ;
       4 : aInvDateRage := '03-04' ;
       5 : aInvDateRage := '05-06' ;
       6 : aInvDateRage := '05-06' ;
       7 : aInvDateRage := '07-08' ;
       8 : aInvDateRage := '07-08' ;
       9 : aInvDateRage := '09-10' ;
       10 : aInvDateRage := '09-10' ;
       11 : aInvDateRage := '11-12' ;
       12 : aInvDateRage := '11-12' ;
  end;
  Result :=Format('%d年%s月',
          [StrToInt(copy(aFmtInvDate,1,4))-1911,
          aInvDateRage]);
end;
}
function TfrmInvA05_1.GetFullInvDate(aInvDate,
  aUpdTime: TDateTime): String;
  var aFmtInvDate : String;
      aYear,aMonth,aDay:String;
begin
  aFmtInvDate := getYearMonthDay(DateTimeToStr(aInvDate));
  aYear := copy(aFmtInvDate,1,4);
  aMonth := copy(aFmtInvDate,5,2);
  aDay := copy(aFmtInvDate,7,2);
  Result := Format('%s-%s-%s %s',
          [aYear,aMonth,aDay,
          FormatDateTime('hh:nn:ss',  aUpdTime)]);
end;
procedure TfrmInvA05_1.Deletefiles(APath, AFileSpec: string);
var
  aSearchRec:TSearchRec;
  aFind:integer;
  aPathName:string;
begin
  aPathName := IncludeTrailingPathDelimiter(APath);

  aFind := FindFirst(aPath+AFileSpec,faAnyFile,aSearchRec);
  while aFind = 0 do
  begin
    DeleteFile(aPathName+aSearchRec.Name);

    aFind := SysUtils.FindNext(aSearchRec);
  end;
  FindClose(aSearchRec);
end;

function TfrmInvA05_1.ProcQRCodePic(aText, aPathName: String): Boolean;
var FSymbol: PZSymbol;
    aBmp : TBitmap;
    aQRCodeData : String;
begin

  aQRCodeData := UTF8Encode(aText);
  aBmp := TBitmap.Create;
  FSymbol := ZBarcode_Create;

  try
    try
      FSymbol.option_1 := 2;
      FSymbol.option_2 := 6;
      FSymbol.symbology := 58;
      FSymbol.scale := 0;
      ZBarcode_Encode_and_Buffer(FSymbol, PChar(aQRCodeData), Length(aQRCodeData), 0);
      ZBarcodeToBitmap(FSymbol, aBmp);
      aBmp.SaveToFile( aPathName );
    except
    on E: Exception do
      begin
        ErrorMsg( Format( '產生QRCODE圖檔失敗, 訊息:%s', [E.Message] ) );
        Result := False;
        Exit;
      end;
    end;
  finally
   ZBarcode_Delete(FSymbol);
   FSymbol := nil;
   aBmp.Free;
   aBmp := nil;
  end;
  Result := True;
end;

end.
