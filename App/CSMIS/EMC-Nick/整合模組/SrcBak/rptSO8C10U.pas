unit rptSO8C10U;

interface

uses Windows, SysUtils, Messages, Classes, Graphics, Controls,
  StdCtrls, ExtCtrls, Forms, QuickRpt, QRCtrls;

type
  TrptSO8C10 = class(TQuickRep)
    QRBand1: TQRBand;
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
    QRLabel11: TQRLabel;
    QRLabel12: TQRLabel;
    QRLabel13: TQRLabel;
    QRLabel14: TQRLabel;
    QRLabel15: TQRLabel;
    QRLabel16: TQRLabel;
    QRLabel17: TQRLabel;
    QRLabel18: TQRLabel;
    QRLabel19: TQRLabel;
    QRLabel20: TQRLabel;
    QRLabel21: TQRLabel;
    QRLabel22: TQRLabel;
    QRLabel23: TQRLabel;
    QRLabel24: TQRLabel;
    QRLabel25: TQRLabel;
    QRLabel26: TQRLabel;
    QRShape1: TQRShape;
    QRBand2: TQRBand;
    QRDBText1: TQRDBText;
    QRDBText3: TQRDBText;
    QRDBText5: TQRDBText;
    QRDBText6: TQRDBText;
    QRDBText7: TQRDBText;
    QRDBText10: TQRDBText;
    QRDBText11: TQRDBText;
    QRDBText12: TQRDBText;
    QRLabel27: TQRLabel;
    QRLabel28: TQRLabel;
    QRLabel29: TQRLabel;
    procedure QRBand1BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure QuickRepPreview(Sender: TObject);
    procedure QRBand2BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
  private

  public
    aCompName: String;
    aPaperDate: String;
    aOperator: String;
    aPaperNo: String;
  end;

var
  rptSO8C10: TrptSO8C10;

implementation

uses dtmMain3U, UdateTimeu, frmRptPreviewU, frmMainMenuU, Variants;

{$R *.DFM}

{ ---------------------------------------------------------------------------- }

procedure TrptSO8C10.QRBand1BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
  QRLabel10.Caption := aCompName;
  QRLabel11.Caption := aPaperDate;
  QRLabel13.Caption := aOperator;
  QRLabel12.Caption := aPaperNo;
  QRLabel15.Caption := TUdateTime.CDateStr(date,9);
  QRLabel14.Caption := frmMainMenu.sG_User;
end;

{ ---------------------------------------------------------------------------- }

procedure TrptSO8C10.QuickRepPreview(Sender: TObject);
var
  frmPreView : TfrmRptPreview;
begin
  frmPreView := TfrmRptPreview.Create(nil);
  frmPreView.Report := Self;
  frmPreView.QRPreview.QRPrinter := Self.QRPrinter;
  frmPreView.Show;
end;

{ ---------------------------------------------------------------------------- }

procedure TrptSO8C10.QRBand2BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
var
  aYear, aMonth, aDay: Word;
begin
  QRLabel27.Caption := EmptyStr;
  if not VarIsNull( DataSet.FieldByName('GETPAPERDATE').Value ) then
  begin
    DecodeDate( DataSet.FieldByName('GETPAPERDATE').AsDateTime,
      aYear, aMonth, aDay );
    QRLabel27.Caption := Format('%3d/%.2d/%.2d', [aYear-1911, aMonth, aDay] );
  end;
  {}
  if not VarIsNull( DataSet.FieldByName( 'RETURNDATE' ).Value ) then
  begin
    DecodeDate( DataSet.FieldByName( 'RETURNDATE' ).AsDateTime,
      aYear, aMonth, aDay );
    QRLabel28.Caption := Format('%3d/%.2d/%.2d', [aYear-1911, aMonth, aDay] );
  end;
  QRLabel28.Caption := Format('%3d/%.2d/%.2d', [aYear-1911, aMonth, aDay] );
  {}
  if not VarIsNull( DataSet.FieldByName( 'RETURNDATE' ).Value ) then
  begin
    DecodeDate( DataSet.FieldByName( 'CLEARDATE' ).AsDateTime,
      aYear, aMonth, aDay );
    QRLabel29.Caption := Format('%3d/%.2d/%.2d', [aYear-1911, aMonth, aDay] );
  end;
  QRLabel29.Caption := Format('%3d/%.2d/%.2d', [aYear-1911, aMonth, aDay] );
end;

{ ---------------------------------------------------------------------------- }

end.
