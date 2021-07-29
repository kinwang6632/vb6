unit rptReport1U;

interface

uses Windows, SysUtils, Messages, Classes, Graphics, Controls,
  StdCtrls, ExtCtrls, Forms, QuickRpt, QRCtrls;

type
  TrptReport1 = class(TQuickRep)
    QRBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel4: TQRLabel;
    QRSysData1: TQRSysData;
    QRSysData2: TQRSysData;
    QRBand2: TQRBand;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel8: TQRLabel;
    QRLabel9: TQRLabel;
    QRLabel11: TQRLabel;
    QRLabel12: TQRLabel;
    QRShape1: TQRShape;
    QRDBText1: TQRDBText;
    QRDBText2: TQRDBText;
    QRDBText3: TQRDBText;
    QRDBText5: TQRDBText;
    QRDBText6: TQRDBText;
    QRDBText7: TQRDBText;
    QRShape2: TQRShape;
    QRShape3: TQRShape;
    QRShape4: TQRShape;
    QRShape5: TQRShape;
    QRShape6: TQRShape;
    QRShape7: TQRShape;
    QRShape8: TQRShape;
    QRShape9: TQRShape;
    QRShape10: TQRShape;
    QRShape11: TQRShape;
    QRShape12: TQRShape;
    QRShape13: TQRShape;
    QRShape14: TQRShape;
    QRBand3: TQRBand;
    QRLabel13: TQRLabel;
    QRShape15: TQRShape;
    QRLabel14: TQRLabel;
    procedure QRLabel5Print(sender: TObject; var Value: String);
    procedure QRLabel14Print(sender: TObject; var Value: String);
  private

  public
    nG_RecCount : Integer; 
    sG_Operator : String;
  end;

var
  rptReport1: TrptReport1;

implementation

uses frmMainU, dtmMainU;

{$R *.DFM}

procedure TrptReport1.QRLabel5Print(sender: TObject; var Value: String);
begin
    Value := sG_Operator;
end;

procedure TrptReport1.QRLabel14Print(sender: TObject; var Value: String);
begin
    Value := IntToStr(nG_RecCount);
end;

end.
