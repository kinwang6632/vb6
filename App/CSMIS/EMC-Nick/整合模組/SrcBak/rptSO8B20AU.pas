unit rptSO8B20AU;

interface

uses Windows, SysUtils, Messages, Classes, Graphics, Controls,
  StdCtrls, ExtCtrls, Forms, QuickRpt, QRCtrls ,DBClient ,DB ,Math;

type
  TrptSO8B20A = class(TQuickRep)
    DetailBand1: TQRBand;
    PageFooterBand1: TQRBand;
    PageHeaderBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel4: TQRLabel;
    QRGroup1: TQRGroup;
    QRBand1: TQRBand;
    QRLabel7: TQRLabel;
    QRLabel9: TQRLabel;
    SummaryBand1: TQRBand;
    QRLabel8: TQRLabel;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRShape3: TQRShape;
    QRLabel10: TQRLabel;
    QRGroup2: TQRGroup;
    QRBand2: TQRBand;
    QRBand3: TQRBand;
    QRLabel11: TQRLabel;
    QRLabel12: TQRLabel;
    QRLabel5: TQRLabel;
    QRLabel6: TQRLabel;
    procedure QuickRepPreview(Sender: TObject);
    procedure QRLabel1Print(sender: TObject; var Value: String);
    procedure QRLabel2Print(sender: TObject; var Value: String);
    procedure QRLabel8Print(sender: TObject; var Value: String);
    procedure DetailBand1BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure QRGroup1BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure QRLabel7Print(sender: TObject; var Value: String);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
    procedure QuickRepAfterPreview(Sender: TObject);
    procedure QRLabel9Print(sender: TObject; var Value: String);
    procedure QRLabel10Print(sender: TObject; var Value: String);
    procedure QRLabel11Print(sender: TObject; var Value: String);
    procedure QRGroup2BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure QRLabel12Print(sender: TObject; var Value: String);
    procedure QRBand2BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure DetailBand1AfterPrint(Sender: TQRCustomBand;
      BandPrinted: Boolean);
    procedure QRLabel6Print(sender: TObject; var Value: String);
  private
     fG_TotalCommAmt,fG_GroupCommAmt,fG_Group2CommAmt : Double;
     sG_UpdTime : String;
     procedure orderByCDSSO121;
  public
    bG_LockData,bG_IsPage : boolean;
    sG_Compute,sG_BelongYM,sG_ManList : String;
  end;

var
  rptSO8B20A: TrptSO8B20A;

implementation

uses dtmMain1U, frmRptPreviewU, UdateTimeu, Ustru, frmMainMenuU;

{$R *.DFM}

procedure TrptSO8B20A.QuickRepPreview(Sender: TObject);
var
    frmPreView : TfrmRptPreview;
begin

  //orderByCDSSO121;
  //查出符合條件的資料
  dtmMain1.getSO121Data(sG_BelongYM,sG_Compute,sG_ManList);
  sG_UpdTime := dtmMain1.cdsSo121.FieldByName('UPDTIME').AsString;

  frmPreView := TfrmRptPreview.Create(nil);

  frmPreView.bG_LockData := bG_LockData;
  frmPreView.G_ActiveDataSet := self.DataSet;

  frmPreView.Report := Self;
  frmPreView.QRPreview.QRPrinter := Self.QRPrinter;
  frmPreView.Show;

end;

procedure TrptSO8B20A.QRLabel1Print(sender: TObject; var Value: String);
begin
  Value := frmMainMenu.sG_CompName;
end;

procedure TrptSO8B20A.QRLabel2Print(sender: TObject; var Value: String);
begin
    Value :='歸屬年月: '+IntToStr(StrToInt(Copy(sG_BelongYM,1,3)))+'年'+Copy(sG_BelongYM,4,2)+'月';
end;



procedure TrptSO8B20A.QRLabel8Print(sender: TObject; var Value: String);
begin
  Value :='總計 : '+ TUstr.CommaNumber(FloatToStr(RoundTo(fG_TotalCommAmt,-2)));

end;

procedure TrptSO8B20A.DetailBand1BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
//    if  self.DataSet.FieldByName('ComputeYM').AsString<>self.DataSet.FieldByName('BelongYM').AsString then
//      PrintBand := false;
    if  sG_BelongYM <> self.DataSet.FieldByName('BelongYM').AsString then
      PrintBand := false;
end;

procedure TrptSO8B20A.orderByCDSSO121;
begin

{JACKAL1221  改為由實際Table來查詢
    //若依部門分群組,則只秀自己公司的資料(CompType=1),不秀關企資料(CompType=2)
    if sG_Compute = '1' then
    begin
      dtmMain1.cdsSo121.Filtered := false;
      dtmMain1.cdsSo121.Filter := 'CompType=1';
      dtmMain1.cdsSo121.Filtered := true;
    end;

    if dtmMain1.cdsSo121 IS TClientDataSet then
    begin
     if TClientDataSet(dtmMain1.cdsSo121).IndexName = 'TmpIndex' then
       TClientDataSet(dtmMain1.cdsSo121).DeleteIndex('TmpIndex');

    if sG_Compute = '0' then        //依公司分群組計算
      TClientDataSet(dtmMain1.cdsSo121).AddIndex('TmpIndex', 'CLANCOMPCODE', [ixCaseInsensitive],'','',0)
    else if sG_Compute = '1' then   //依部門分群組計算
      TClientDataSet(dtmMain1.cdsSo121).AddIndex('TmpIndex', 'MEDIACODE', [ixCaseInsensitive],'','',0);


     TClientDataSet(dtmMain1.cdsSo121).IndexName := 'TmpIndex';
    end;
}
end;

procedure TrptSO8B20A.QRGroup1BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
{
    if sG_Compute = '0' then        //依公司分群組計算
      QRGroup1.Expression := 'CLANCOMPCODE'
    else if sG_Compute = '1' then   //依部門分群組計算
      QRGroup1.Expression := 'MEDIACODE';
}

    fG_GroupCommAmt := 0 ;

    if bG_IsPage = true then    //報表依分群組來分頁
      QRGroup1.ForceNewPage := true
    else
      QRGroup1.ForceNewPage := false;



end;

procedure TrptSO8B20A.QRLabel7Print(sender: TObject; var Value: String);
var
    sL_ClanCompName,sL_MediaCode : String;
begin
    if sG_Compute = '0' then        //依公司分群組計算
    begin
      sL_ClanCompName := dtmMain1.cdsSO121.FieldByName('CLANCOMPNAME').AsString;
      if sL_ClanCompName <> '' then
        value := sL_ClanCompName
      else
        value := frmMainMenu.sG_CompName + '_其他';
    end
    else if sG_Compute = '1' then   //依部門分群組計算
    begin
      sL_MediaCode := dtmMain1.cdsSO121.FieldByName('MEDIANAME').AsString;
      if sL_MediaCode <> '' then
        value := sL_MediaCode
      else
        value := '其他';
    end;
end;

procedure TrptSO8B20A.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
    if sG_Compute = '0' then        //依公司分群組計算
      QRGroup1.Expression := 'CLANCOMPCODE'
    else if sG_Compute = '1' then   //依部門分群組計算
      QRGroup1.Expression := 'MEDIACODE';

    //列印前先歸零
    fG_TotalCommAmt := 0 ;      
end;

procedure TrptSO8B20A.QuickRepAfterPreview(Sender: TObject);
begin
     if TClientDataSet(dtmMain1.cdsSo121).IndexName = 'TmpIndex' then
       TClientDataSet(dtmMain1.cdsSo121).DeleteIndex('TmpIndex');
end;

procedure TrptSO8B20A.QRLabel9Print(sender: TObject; var Value: String);
begin
  Value :='小計 : '+ TUstr.CommaNumber(FloatToStr(RoundTo(fG_GroupCommAmt,-2)));
end;

procedure TrptSO8B20A.QRLabel10Print(sender: TObject; var Value: String);
begin
    Value := '佣金計算時間: ' + sG_UpdTime;
end;

procedure TrptSO8B20A.QRLabel11Print(sender: TObject; var Value: String);
begin
   value:=dtmMain1.cdsSO121.FieldByName('EMPNAME').AsString;
end;

procedure TrptSO8B20A.QRGroup2BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
    fG_Group2CommAmt := 0;
end;

procedure TrptSO8B20A.QRLabel12Print(sender: TObject; var Value: String);
begin
    Value := TUstr.CommaNumber(FloatToStr(RoundTo(fG_Group2CommAmt,-2)));
end;

procedure TrptSO8B20A.QRBand2BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
    if  sG_BelongYM <> self.DataSet.FieldByName('BelongYM').AsString then
      PrintBand := false;
end;

procedure TrptSO8B20A.DetailBand1AfterPrint(Sender: TQRCustomBand;
  BandPrinted: Boolean);
var dL_Comm : double;
begin
    dL_Comm:=dtmMain1.cdsSO121.FieldByName('COMMISSION').AsFloat;

    fG_Group2CommAmt := fG_Group2CommAmt + dL_Comm;
    fG_GroupCommAmt := fG_GroupCommAmt + dL_Comm;
    fG_TotalCommAmt := fG_TotalCommAmt + dL_Comm;
end;

procedure TrptSO8B20A.QRLabel6Print(sender: TObject; var Value: String);
begin
   value:=dtmMain1.cdsSO121.FieldByName('EMPID').AsString;
end;

end.
