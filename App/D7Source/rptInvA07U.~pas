unit rptInvA07U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, QRCtrls, QuickRpt, ExtCtrls;

type
  TrptInvA07 = class(TQuickRep)
    QRBand1: TQRBand;
    QRDBText1: TQRDBText;
    QRDBText2: TQRDBText;
    QRDBText3: TQRDBText;
    QRDBText4: TQRDBText;
    QRDBText5: TQRDBText;
    QRDBText6: TQRDBText;
    QRDBText7: TQRDBText;
    QRDBText8: TQRDBText;
    QRDBText9: TQRDBText;
    QRDBText10: TQRDBText;
    qrbPageHeader: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel5: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel8: TQRLabel;
    QRLabel9: TQRLabel;
    QRLabel10: TQRLabel;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRShape3: TQRShape;
    QRShape4: TQRShape;
    QRShape5: TQRShape;
    QRShape6: TQRShape;
    QRShape7: TQRShape;
    QRShape8: TQRShape;
    QRShape9: TQRShape;
    QRShape10: TQRShape;
    qrlCompName: TQRLabel;
    qrlRptName: TQRLabel;
    lbCondition: TQRLabel;
    QRBand2: TQRBand;
    QRLabel12: TQRLabel;
    QRLabel13: TQRLabel;
    QRBand3: TQRBand;
    QRShape11: TQRShape;
    QRLabel14: TQRLabel;
    QRLabel15: TQRLabel;
    QRLabel16: TQRLabel;
    QRLabel17: TQRLabel;
    procedure QRDBText4Print(sender: TObject; var Value: String);
    procedure QRDBText8Print(sender: TObject; var Value: String);
    procedure QRBand2BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
    procedure QRBand1BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure QRBand3BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure lbConditionPrint(sender: TObject; var Value: String);
  private
    FCompName: String;
    FRptName: String;
    FOperator: String;
    FCondition: String;
    FSaleAmount: Integer;
    FTaxAmount: Integer;
    FInvAmount: Integer;
    procedure SetCompName(const Value: String);
    procedure SetOperator(const Value: String);
    procedure SetRptName(const Value: String);
    { Private declarations }
  public
    { Public declarations }
    property CompName: String read FCompName write SetCompName;
    property RptName: String read FRptName write SetRptName;
    property Operator: String read FOperator write SetOperator;
    property Condition: String read FCondition write FCondition;
  end;

var
  rptInvA07: TrptInvA07;

implementation

uses dtmMainHU;

{$R *.dfm}

{ TrptInvA07 }

{ ---------------------------------------------------------------------------- }

procedure TrptInvA07.SetCompName(const Value: String);
begin
  FCompName := Value;
  qrlCompName.Caption := FCompName;
end;

{ ---------------------------------------------------------------------------- }

procedure TrptInvA07.SetOperator(const Value: String);
begin
  FOperator := Value;
end;

{ ---------------------------------------------------------------------------- }

procedure TrptInvA07.SetRptName(const Value: String);
begin
  FRptName := Value;
  qrlRptName.Caption := FRptName;
end;

{ ---------------------------------------------------------------------------- }

procedure TrptInvA07.QRDBText4Print(sender: TObject; var Value: String);
begin
 if DataSet.FieldByName( 'INVFORMAT' ).AsString = '1' then
   Value := '電子'
 else if DataSet.FieldByName( 'INVFORMAT' ).AsString = '2' then
   Value := '手二'
 else if DataSet.FieldByName( 'INVFORMAT' ).AsString = '3' then
   Value := '手三';
end;

{ ---------------------------------------------------------------------------- }

procedure TrptInvA07.QRDBText8Print(sender: TObject; var Value: String);
begin
 if DataSet.FieldByName( 'TAXTYPE' ).AsString = '1' then
   Value := '應稅'
 else if DataSet.FieldByName( 'TAXTYPE' ).AsString = '2' then
   Value := '零稅率'
 else if DataSet.FieldByName( 'TAXTYPE' ).AsString = '3' then
   Value := '免稅';
end;

{ ---------------------------------------------------------------------------- }

procedure TrptInvA07.QRBand2BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
 QRLabel12.Caption := Format( '列印日期: %s', [FormatDateTime( 'yyyy/mm/dd', Date ) ] );
 QRLabel13.Caption := Format( '印表人: %s', [FOperator] );
end;

{ ---------------------------------------------------------------------------- }

procedure TrptInvA07.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
  QRLabel14.Caption := EmptyStr;
  QRLabel15.Caption := EmptyStr;
  QRLabel16.Caption := EmptyStr;
  FSaleAmount := 0;
  FTaxAmount := 0;
  FInvAmount := 0;
end;

{ ---------------------------------------------------------------------------- }

procedure TrptInvA07.QRBand1BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
  FSaleAmount := ( FSaleAmount + DataSet.FieldByName( 'SALEAMOUNT' ).AsInteger );
  FTaxAmount := ( FTaxAmount + DataSet.FieldByName( 'TAXAMOUNT' ).AsInteger );
  FInvAmount := ( FInvAmount + DataSet.FieldByName( 'INVAMOUNT' ).AsInteger );
end;

{ ---------------------------------------------------------------------------- }

procedure TrptInvA07.QRBand3BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
  QRLabel14.Caption := FormatFloat( '#,##0', FSaleAmount );
  QRLabel15.Caption := FormatFloat( '#,##0',  FTaxAmount );
  QRLabel16.Caption := FormatFloat( '#,##0',  FInvAmount );
end;

{ ---------------------------------------------------------------------------- }

procedure TrptInvA07.lbConditionPrint(sender: TObject; var Value: String);
begin
  Value := FCondition;
end;

{ ---------------------------------------------------------------------------- }

end.
