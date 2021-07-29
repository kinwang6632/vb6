unit print_frmcd042;

interface

uses Windows, SysUtils, Messages, Classes, Graphics, Controls,
  StdCtrls, ExtCtrls, Forms, QuickRpt, QRCtrls;

type
  TQuickReport1 = class(TQuickRep)
    QRBand2: TQRBand;
    QRBand3: TQRBand;
    QRBand4: TQRBand;
    QRBand5: TQRBand;
    QRLabel4: TQRLabel;
    QRLabel1: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QR_CompName: TQRLabel;
    QR_FormName: TQRLabel;
    QRLabel5: TQRLabel;
    QRDBText1: TQRDBText;
    QRDBText2: TQRDBText;
    QR_User: TQRLabel;
    QRLabel7: TQRLabel;
    QRSysData2: TQRSysData;
    QRLabel6: TQRLabel;
    QRSysData3: TQRSysData;
    QRLabel9: TQRLabel;
    QRDBText3: TQRDBText;
    QRLabel10: TQRLabel;
    QRDBText4: TQRDBText;
    QRLabel11: TQRLabel;
    QRDBText5: TQRDBText;
    QRLabel12: TQRLabel;
    QRDBText6: TQRDBText;
    QRLabel13: TQRLabel;
    QRDBText7: TQRDBText;
    QRLabel15: TQRLabel;
    QRLabel16: TQRLabel;
    QRLabel17: TQRLabel;
    QRLabel18: TQRLabel;
    QRLabel19: TQRLabel;
    QRLabel20: TQRLabel;
    QRLabel21: TQRLabel;
    QRDBText9: TQRDBText;
    QRDBText10: TQRDBText;
    QRDBText11: TQRDBText;
    QRDBText12: TQRDBText;
    QRDBText13: TQRDBText;
    QRDBText14: TQRDBText;
    QRLabel8: TQRLabel;
    QRLabel14: TQRLabel;
    QRLabel22: TQRLabel;
    QRLabel23: TQRLabel;
    QRLabel24: TQRLabel;
    QRLabel25: TQRLabel;
    QRLabel26: TQRLabel;
    QRLabel27: TQRLabel;
    QRLabel28: TQRLabel;
    QRLabel29: TQRLabel;
    QRLabel30: TQRLabel;
    QRLabel31: TQRLabel;
    QRLabel32: TQRLabel;
    QRLabel33: TQRLabel;
    QRLabel34: TQRLabel;
    QRLabel35: TQRLabel;
    QRLabel36: TQRLabel;
    GF_GiftOrderPrice1: TQRDBText;
    GF_GiftOrderPrice2: TQRDBText;
    GF_GiftOrderPrice3: TQRDBText;
    GF_GiftOrderPrice4: TQRDBText;
    GF_GiftOrderPrice5: TQRDBText;
    GF_GiftOrderPrice6: TQRDBText;
    GF_GiftType1: TQRDBText;
    GF_GiftType2: TQRDBText;
    GF_GiftType3: TQRDBText;
    GF_GiftType4: TQRDBText;
    GF_GiftType5: TQRDBText;
    GF_GiftType6: TQRDBText;
    GF_ArticleNoStr1: TQRDBText;
    GF_ArticleNoStr2: TQRDBText;
    GF_ArticleNoStr3: TQRDBText;
    GF_ArticleNoStr4: TQRDBText;
    GF_ArticleNoStr5: TQRDBText;
    GF_ArticleNoStr6: TQRDBText;
    QRLabel37: TQRLabel;
    QRLabel38: TQRLabel;
    QRLabel39: TQRLabel;
    QRLabel40: TQRLabel;
    QRShape1: TQRShape;
    procedure QRBand3BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure QRDBText4Print(sender: TObject; var Value: String);
  private

  public

  end;

var
  QuickReport1: TQuickReport1;

implementation

uses
  frmcd042, cbUtilis;

{$R *.DFM}

procedure TQuickReport1.QRBand3BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
var
  I:Integer;
begin
  if Form1.ClientDataSet2.FieldByName('ProductType').AsInteger=1 then
    begin
      QRLabel23.Caption:='特定產品';
      QRLabel25.Caption:=Form1.ClientDataSet2.FieldByName('CitemCodeStr').AsString;
    end
  else if Form1.ClientDataSet2.FieldByName('ProductType').AsInteger=2 then
    begin
      QRLabel23.Caption:='組合產品';
      QRLabel25.Caption:=Form1.ClientDataSet2.FieldByName('BPCodeStr').AsString;
    end
  else
    begin
      QRLabel23.Caption:='';
      QRLabel25.Caption:='';
    end;

  QRLabel26.Enabled:=Form1.ClientDataSet2.FieldByName('Punish').AsInteger=1;

  QRLabel27.Enabled:=Form1.ClientDataSet2.FieldByName('Price').AsInteger=1;
  QRLabel34.Enabled:=Form1.ClientDataSet2.FieldByName('Price').AsInteger=1;
  QRLabel35.Enabled:=Form1.ClientDataSet2.FieldByName('Price').AsInteger=1;
  if QRLabel27.Enabled then
    if Form1.ClientDataSet2.FieldByName('PriceDiscountType').AsInteger=1 then
      begin
        QRLabel34.Caption:='折扣';
        QRLabel35.Caption:=Form1.ClientDataSet2.FieldByName('PriceDiscountRate').AsString;
      end
    else
      begin
        QRLabel34.Caption:='固定金額';
        QRLabel35.Caption:=Form1.ClientDataSet2.FieldByName('PriceFixAmount').AsString;
      end;

  QRLabel28.Enabled:=Form1.ClientDataSet2.FieldByName('Cash').AsInteger=1;
  QRLabel28.Caption:='現金回饋，項目：'+Form1.ClientDataSet2.FieldByName('CashCitemCodeStr').AsString;
  QRLabel32.Enabled:=Form1.ClientDataSet2.FieldByName('Cash').AsInteger=1;
  QRLabel33.Enabled:=Form1.ClientDataSet2.FieldByName('Cash').AsInteger=1;
  if QRLabel28.Enabled then
    if Form1.ClientDataSet2.FieldByName('CashDiscountType').AsInteger=1 then
      begin
        QRLabel32.Caption:='訂購金額％';
        QRLabel33.Caption:=Form1.ClientDataSet2.FieldByName('CashOrderRate').AsString;
      end
    else
      begin
        QRLabel32.Caption:='固定金額';
        QRLabel33.Caption:=Form1.ClientDataSet2.FieldByName('CashFixAmount').AsString;
      end;

  QRLabel29.Enabled:=Form1.ClientDataSet2.FieldByName('Coupon').AsInteger=1;
  QRLabel30.Enabled:=Form1.ClientDataSet2.FieldByName('Coupon').AsInteger=1;
  QRLabel31.Enabled:=Form1.ClientDataSet2.FieldByName('Coupon').AsInteger=1;
  if QRLabel29.Enabled then
    if Form1.ClientDataSet2.FieldByName('CouponDiscountType').AsInteger=1 then
      begin
        QRLabel30.Caption:='訂購金額％';
        QRLabel31.Caption:=Form1.ClientDataSet2.FieldByName('CouponOrderRate').AsString;
      end
    else
      begin
        QRLabel30.Caption:='固定金額';
        QRLabel31.Caption:=Form1.ClientDataSet2.FieldByName('CouponFixAmount').AsString;
      end;

  for I:=0 to QrBand3.ControlCount-1 do
    if (QrBand3.Controls[I] is TQRDBText) and (Copy(TQRDBText(QrBand3.Controls[I]).Name,1,2)='GF') then
      TQRDBText(QrBand3.Controls[I]).Enabled:=Form1.ClientDataSet2.FieldByName('Gift').AsInteger=1;

  for I:=1 to 6 do
    if Form1.ClientDataSet2.FieldByName('GiftType'+IntToStr(I)).AsInteger<0 then
    begin
      QrBand3.FindChildControl('GF_GiftOrderPrice'+IntToStr(I)).Enabled:=False;
      QrBand3.FindChildControl('GF_GiftType'+IntToStr(I)).Enabled:=False;
      QrBand3.FindChildControl('GF_ArticleNoStr'+IntToStr(I)).Enabled:=False;
    end;
end;

procedure TQuickReport1.QRDBText4Print(sender: TObject; var Value: String);
begin
  Value := DateConvert( Value, True );
end;

end.
