unit dtmReportModule;
  {
    重要:
      1. frxA05_1.CloseDataSource 不可設定為 True, 會影響發票套表後,
         更新 INV007
      2. frxA07_1.CloseDataSource 不可設定為 True, 會影響折讓資料輸入
  }

interface

uses
  SysUtils, Classes, frxClass, frxDBSet, DB, ADODB, DBClient,
  frxDesgn, frxRes, frxChBox, frxExportXLS, frxBarcode;


type
  TdtmReport = class(TDataModule)
    ADOMaster: TADOQuery;
    ADODetail: TADOQuery;
    cdsTempory: TClientDataSet;
    frxA01_1: TfrxDBDataset;
    frxA01_2: TfrxDBDataset;
    frxA05_1: TfrxDBDataset;
    frxA07_1: TfrxDBDataset;
    frxB04_1: TfrxDBDataset;
    frxC01_1: TfrxDBDataset;
    frxC03_1: TfrxDBDataset;
    frxC04_1: TfrxDBDataset;
    frxC05_1: TfrxDBDataset;
    frxC01_2: TfrxDBDataset;
    frxB06_1: TfrxDBDataset;
    frxDBDataset1: TfrxDBDataset;
    frxA07_2: TfrxDBDataset;
    frxMainReport: TfrxReport;
    procedure DataModuleCreate(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
    procedure AddA05DataField(aAutoCreateNum: Integer);
  end;

var
  dtmReport: TdtmReport;

implementation

uses cbUtilis, dtmMainU;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TdtmReport.DataModuleCreate(Sender: TObject);
begin
  if FileExists( IncludeTrailingPathDelimiter(
    ExtractFilePath( ParamStr( 0 ) ) ) + 'frxrcTaiwan.frc' ) then
  begin
    frxResources.LoadFromFile( IncludeTrailingPathDelimiter(
      ExtractFilePath( ParamStr( 0 ) ) ) + 'frxrcTaiwan.frc' );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmReport.AddA05DataField(aAutoCreateNum: Integer);
var
  aIndex: Integer;
begin
  frxA05_1.FieldAliases.Clear;
  frxA05_1.FieldAliases.Add( 'CHECKNO=檢查碼' );
  frxA05_1.FieldAliases.Add( 'INVID=發票號碼' );
  frxA05_1.FieldAliases.Add( 'BUSINESSID=統一編號' );
  frxA05_1.FieldAliases.Add( 'INVDATE=發票日期' );
  frxA05_1.FieldAliases.Add( 'CUSTID=客戶代碼' );
  frxA05_1.FieldAliases.Add( 'CUSTSNAME=客戶簡稱' );
  frxA05_1.FieldAliases.Add( 'INVTITLE=發票抬頭' );
  frxA05_1.FieldAliases.Add( 'ZIPCODE=郵遞區號' );
  frxA05_1.FieldAliases.Add( 'MAILADDR=郵寄地址' );
  frxA05_1.FieldAliases.Add( 'INVADDR=發票地址' );
  frxA05_1.FieldAliases.Add( 'INSTADDR=裝機地址' );
  frxA05_1.FieldAliases.Add( 'INVFORMAT=發票格式' );
  frxA05_1.FieldAliases.Add( 'TAXTYPE=稅別' );
  frxA05_1.FieldAliases.Add( 'SALEAMOUNT=銷售額' );
  frxA05_1.FieldAliases.Add( 'TAXAMOUNT=稅額' );
  frxA05_1.FieldAliases.Add( 'INVAMOUNT=發票總金額' );

  if aAutoCreateNum <= 0 then aAutoCreateNum := 5;

  for aIndex := 1 to aAutoCreateNum do
  begin
    if aIndex >= 100 then Break;
    frxA05_1.FieldAliases.Add( Format( 'BILLID%d=收費單號(%s)', [aIndex, MappingChineseNumber( aIndex )] ) );
    frxA05_1.FieldAliases.Add( Format( 'DESCRIPTION%d=品名(%s)', [aIndex, MappingChineseNumber( aIndex )] ) );
    frxA05_1.FieldAliases.Add( Format( 'STARTDATE%d=有效日起(%s)', [aIndex, MappingChineseNumber( aIndex )] ) );
    frxA05_1.FieldAliases.Add( Format( 'ENDDATE%d=有效日迄(%s)', [aIndex, MappingChineseNumber( aIndex )] ) );
    frxA05_1.FieldAliases.Add( Format( 'QUANTITY%d=數量(%s)', [aIndex, MappingChineseNumber( aIndex )] ) );
    frxA05_1.FieldAliases.Add( Format( 'UNITPRICE%d=單價(%s)', [aIndex, MappingChineseNumber( aIndex )] ) );
    frxA05_1.FieldAliases.Add( Format( 'SALEAMOUNT%d=銷售額(%s)', [aIndex, MappingChineseNumber( aIndex )] ) );
    frxA05_1.FieldAliases.Add( Format( 'TOTALAMOUNT%d=總金額(%s)', [aIndex, MappingChineseNumber( aIndex )] ) );
    frxA05_1.FieldAliases.Add( Format( 'FACISNO%d=設備號碼(%s)', [aIndex, MappingChineseNumber( aIndex )] ) );
  end;
  frxA05_1.FieldAliases.Add( 'MEMO1=備註' );
  frxA05_1.FieldAliases.Add( 'ACCOUNTNO=信用卡號' );
  frxA05_1.FieldAliases.Add( 'MAININVID=主發票號碼' );
  frxA05_1.FieldAliases.Add( 'MAINSALEAMOUNT=主發票總銷售額' );
  frxA05_1.FieldAliases.Add( 'MAINTAXAMOUNT=主發票總稅額' );
  frxA05_1.FieldAliases.Add( 'MAININVAMOUNT=主發票總金額' );
  frxA05_1.FieldAliases.Add( 'HOWTOCREATE=此筆資料如何產生' );
  frxA05_1.FieldAliases.Add( 'INVUSEDESC=發票用途' );
  {#6600 增加發票防偽隨機碼 By Kin 2013/10/02}
  frxA05_1.FieldAliases.Add( 'RANDOMNUM=發票防偽隨機碼' );
  {#6629 增加申報格式 By Kin 2013/11/01}
  frxA05_1.FieldAliases.Add( 'APPLYFORMAT=申報格式' );
  {#7060 增加臨櫃信用卡末4碼 By Kin 2015/08/18}
  frxA05_1.FieldAliases.Add( 'CREDITCARD4D=臨櫃信用卡末四碼');
  frxA05_1.FieldAliases.Add( 'QRCodeChineseInvDate=QRCODE民國年發票日期');
  frxA05_1.FieldAliases.Add( 'QRCodeInvDate=QRCode發票日期時間');
  frxA05_1.FieldAliases.Add( 'QRCodeSellerBUSINESSID=QRCODE賣方統一編號');
  frxA05_1.FieldAliases.Add( 'QRCodeBuyerBUSINESSID=QRcode買方統一編號');
  frxA05_1.FieldAliases.Add( 'CODE39Data=CDDE39資料內容' );
  frxA05_1.FieldAliases.Add( 'QRCodePngPath1=QRCode1圖檔路徑' );
  frxA05_1.FieldAliases.Add( 'QRCodePngPath2=QRCode2圖檔路徑' );
  frxA05_1.FieldAliases.Add( 'QRCodeInvID=QRCode發票號碼' );
  frxA05_1.FieldAliases.Add( 'QRCodeRanNum=QRCode隨機碼' );
  frxA05_1.FieldAliases.Add( 'QRCodeTotalAmount=QRCode總計金額' );
  frxA05_1.FieldAliases.Add( 'QRCodeInvID=QRCode發票號碼' );
  frxA05_1.FieldAliases.Add( 'CanModify=是否可修改' );
  frxA05_1.FieldAliases.Add( 'QRCodeEFormat=QRCode格式' );
  frxA05_1.FieldAliases.Add( 'CompSName=公司簡稱');
  {#7128 Adding some fields into report By Kin 2015/12/14}
  frxA05_1.FieldAliases.Add( 'COMPID=公司別' );
  frxA05_1.FieldAliases.Add( 'COMPNAME=公司全名' );
  frxA05_1.FieldAliases.Add( 'COMPBUSINESSID=公司統編' );
  frxA05_1.FieldAliases.Add( 'COMPADDR=公司營業地址' );
  frxA05_1.FieldAliases.Add( 'COMPTEL=公司電話' );
  frxA05_1.FieldAliases.Add( 'PRIZETYPE=中獎記錄' );

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



end.
