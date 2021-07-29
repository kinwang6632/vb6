unit rptSO8B20B_4U;

interface

uses Windows, SysUtils, Messages, Classes, Graphics, Controls,
  StdCtrls, ExtCtrls, Forms, QuickRpt, QRCtrls,Math;

type
  TrptSO8B20B_4 = class(TQuickRep)
    DetailBand1: TQRBand;
    PageFooterBand1: TQRBand;
    PageHeaderBand1: TQRBand;
    QRLabel1: TQRLabel;
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
    QRLabel27: TQRLabel;
    QRLabel28: TQRLabel;
    QRLabel29: TQRLabel;
    QRLabel30: TQRLabel;
    QRLabel31: TQRLabel;
    QRLabel32: TQRLabel;
    QRShape1: TQRShape;
    QRLabel2: TQRLabel;
    QRLabel33: TQRLabel;
    QRLabel34: TQRLabel;
    procedure QuickRepPreview(Sender: TObject);
    procedure QRLabel1Print(sender: TObject; var Value: String);
    procedure QRLabel19Print(sender: TObject; var Value: String);
    procedure QRLabel20Print(sender: TObject; var Value: String);
    procedure QRLabel21Print(sender: TObject; var Value: String);
    procedure QRLabel22Print(sender: TObject; var Value: String);
    procedure QRLabel23Print(sender: TObject; var Value: String);
    procedure QRLabel24Print(sender: TObject; var Value: String);
    procedure QRLabel25Print(sender: TObject; var Value: String);
    procedure QRLabel26Print(sender: TObject; var Value: String);
    procedure QRLabel27Print(sender: TObject; var Value: String);
    procedure QRLabel28Print(sender: TObject; var Value: String);
    procedure QRLabel29Print(sender: TObject; var Value: String);
    procedure QRLabel30Print(sender: TObject; var Value: String);
    procedure QRLabel31Print(sender: TObject; var Value: String);
    procedure QRLabel3Print(sender: TObject; var Value: String);
    procedure QRLabel4Print(sender: TObject; var Value: String);
    procedure QRLabel5Print(sender: TObject; var Value: String);
    procedure QRLabel6Print(sender: TObject; var Value: String);
    procedure QRLabel7Print(sender: TObject; var Value: String);
    procedure QRLabel8Print(sender: TObject; var Value: String);
    procedure QRLabel9Print(sender: TObject; var Value: String);
    procedure QRLabel10Print(sender: TObject; var Value: String);
    procedure QRLabel11Print(sender: TObject; var Value: String);
    procedure QRLabel12Print(sender: TObject; var Value: String);
    procedure QRLabel13Print(sender: TObject; var Value: String);
    procedure QRLabel14Print(sender: TObject; var Value: String);
    procedure QRLabel15Print(sender: TObject; var Value: String);
    procedure QRLabel16Print(sender: TObject; var Value: String);
    procedure QRLabel17Print(sender: TObject; var Value: String);
    procedure QRLabel33Print(sender: TObject; var Value: String);
    procedure QRLabel34Print(sender: TObject; var Value: String);
  private

  public
    sG_Compute,sG_BasicYear,sG_CompName,sG_User,sG_CurDataTime : String;

  end;

var
  rptSO8B20B_4: TrptSO8B20B_4;

implementation

uses dtmMain1U, frmRptPreviewU, Ustru;

{$R *.DFM}


procedure TrptSO8B20B_4.QuickRepPreview(Sender: TObject);
var
    frmPreView : TfrmRptPreview;
begin
    frmPreView := TfrmRptPreview.Create(nil);

    frmPreView.G_ActiveDataSet := self.DataSet;
    frmPreView.Report := Self;
    frmPreView.QRPreview.QRPrinter := Self.QRPrinter;
    frmPreView.Show;
end;

procedure TrptSO8B20B_4.QRLabel1Print(sender: TObject; var Value: String);
begin
//    Value := sG_CompName + ' ' + sG_BasicYear + ' 年度付費頻道介紹佣金攤提表';
Value := sG_CompName + ' ' + IntToStr( StrToInt( sG_BasicYear ) + 1911 - 1911 ) + ' 年度付費頻道介紹佣金攤提表';
end;

procedure TrptSO8B20B_4.QRLabel19Print(sender: TObject; var Value: String);
begin
    Value := IntToStr( StrToInt( sG_BasicYear ) + 1911 - 1911 ) + '/01';
end;

procedure TrptSO8B20B_4.QRLabel20Print(sender: TObject; var Value: String);
begin
    Value := IntToStr( StrToInt( sG_BasicYear ) + 1911 - 1911 ) + '/02';
end;

procedure TrptSO8B20B_4.QRLabel21Print(sender: TObject; var Value: String);
begin
    Value := IntToStr( StrToInt( sG_BasicYear ) + 1911 - 1911 ) + '/03';
end;

procedure TrptSO8B20B_4.QRLabel22Print(sender: TObject; var Value: String);
begin
    Value := IntToStr( StrToInt( sG_BasicYear ) + 1911 - 1911 ) + '/04';
end;

procedure TrptSO8B20B_4.QRLabel23Print(sender: TObject; var Value: String);
begin
    Value := IntToStr( StrToInt( sG_BasicYear ) + 1911 - 1911 ) + '/05';
end;

procedure TrptSO8B20B_4.QRLabel24Print(sender: TObject; var Value: String);
begin
    Value := IntToStr( StrToInt( sG_BasicYear ) + 1911 - 1911 ) + '/06';
end;

procedure TrptSO8B20B_4.QRLabel25Print(sender: TObject; var Value: String);
begin
    Value := IntToStr( StrToInt( sG_BasicYear ) + 1911 - 1911 ) + '/07';
end;

procedure TrptSO8B20B_4.QRLabel26Print(sender: TObject; var Value: String);
begin
    Value := IntToStr( StrToInt( sG_BasicYear ) + 1911 - 1911 ) + '/08';
end;

procedure TrptSO8B20B_4.QRLabel27Print(sender: TObject; var Value: String);
begin
    Value := IntToStr( StrToInt( sG_BasicYear ) + 1911 - 1911  ) + '/09';
end;

procedure TrptSO8B20B_4.QRLabel28Print(sender: TObject; var Value: String);
begin
    Value := IntToStr( StrToInt( sG_BasicYear ) + 1911 - 1911  ) + '/10';
end;

procedure TrptSO8B20B_4.QRLabel29Print(sender: TObject; var Value: String);
begin
    Value := IntToStr( StrToInt( sG_BasicYear ) + 1911 - 1911  ) + '/11';
end;

procedure TrptSO8B20B_4.QRLabel30Print(sender: TObject; var Value: String);
begin
    Value := IntToStr( StrToInt( sG_BasicYear ) + 1911 - 1911  ) + '/12';
end;

procedure TrptSO8B20B_4.QRLabel31Print(sender: TObject; var Value: String);
begin
    Value := IntToStr(StrToInt(IntToStr( StrToInt( sG_BasicYear ) + 1911
       - 1911 ) )+1) + '以後';
end;

procedure TrptSO8B20B_4.QRLabel3Print(sender: TObject; var Value: String);
var
    sL_ComputeYM : String;
    aPos: Integer;
begin
   sL_ComputeYM := self.DataSet.FieldByName('ComputeYM').AsString;
   if ( AnsiPos( '年度', sL_ComputeYM ) > 0 ) then
   begin
     value := Copy( sL_ComputeYM, 1, AnsiPos( '年度', sL_ComputeYM ) - 1 );
     value := IntToStr( StrToInt( value ) + 1911 - 1911 ) + '年度';
   end else
   if ( AnsiPos( '總計', sL_ComputeYM ) > 0 ) then
   begin
     value := sL_ComputeYM;
   end else
   begin
    aPos := Pos( '/',  sL_ComputeYM );
    value := IntToStr( StrToInt( Copy( sL_ComputeYM, 1, aPos - 1 ) ) + 1911 - 1911 ) + '/' +
      Copy( sL_ComputeYM, aPos + 1, 2 );
   end;

    if sL_ComputeYM = '總計' then
    begin
      QRLabel3.Font.Color := clred;
      QRLabel4.Font.Color := clred;
      QRLabel5.Font.Color := clred;
      QRLabel6.Font.Color := clred;
      QRLabel7.Font.Color := clred;
      QRLabel8.Font.Color := clred;
      QRLabel9.Font.Color := clred;
      QRLabel10.Font.Color := clred;
      QRLabel11.Font.Color := clred;
      QRLabel12.Font.Color := clred;
      QRLabel13.Font.Color := clred;
      QRLabel14.Font.Color := clred;
      QRLabel15.Font.Color := clred;
      QRLabel16.Font.Color := clred;
      QRLabel17.Font.Color := clred;
    end;
end;

procedure TrptSO8B20B_4.QRLabel4Print(sender: TObject; var Value: String);
var
    fL_Amount : Double;
begin
    fL_Amount := self.DataSet.FieldByName('Month1').AsFloat;
    value := TUstr.CommaNumber(FloatToStr(RoundTo(fL_Amount,-2)));
end;

procedure TrptSO8B20B_4.QRLabel5Print(sender: TObject; var Value: String);
var
    fL_Amount : Double;
begin
    fL_Amount := self.DataSet.FieldByName('Month2').AsFloat;
    value := TUstr.CommaNumber(FloatToStr(RoundTo(fL_Amount,-2)));
end;

procedure TrptSO8B20B_4.QRLabel6Print(sender: TObject; var Value: String);
var
    fL_Amount : Double;
begin
    fL_Amount := self.DataSet.FieldByName('Month3').AsFloat;
    value := TUstr.CommaNumber(FloatToStr(RoundTo(fL_Amount,-2)));
end;

procedure TrptSO8B20B_4.QRLabel7Print(sender: TObject; var Value: String);
var
    fL_Amount : Double;
begin
    fL_Amount := self.DataSet.FieldByName('Month4').AsFloat;
    value := TUstr.CommaNumber(FloatToStr(RoundTo(fL_Amount,-2)));
end;

procedure TrptSO8B20B_4.QRLabel8Print(sender: TObject; var Value: String);
var
    fL_Amount : Double;
begin
    fL_Amount := self.DataSet.FieldByName('Month5').AsFloat;
    value := TUstr.CommaNumber(FloatToStr(RoundTo(fL_Amount,-2)));
end;

procedure TrptSO8B20B_4.QRLabel9Print(sender: TObject; var Value: String);
var
    fL_Amount : Double;
begin
    fL_Amount := self.DataSet.FieldByName('Month6').AsFloat;
    value := TUstr.CommaNumber(FloatToStr(RoundTo(fL_Amount,-2)));
end;

procedure TrptSO8B20B_4.QRLabel10Print(sender: TObject; var Value: String);
var
    fL_Amount : Double;
begin
    fL_Amount := self.DataSet.FieldByName('Month7').AsFloat;
    value := TUstr.CommaNumber(FloatToStr(RoundTo(fL_Amount,-2)));
end;

procedure TrptSO8B20B_4.QRLabel11Print(sender: TObject; var Value: String);
var
    fL_Amount : Double;
begin
    fL_Amount := self.DataSet.FieldByName('Month8').AsFloat;
    value := TUstr.CommaNumber(FloatToStr(RoundTo(fL_Amount,-2)));
end;

procedure TrptSO8B20B_4.QRLabel12Print(sender: TObject; var Value: String);
var
    fL_Amount : Double;
begin
    fL_Amount := self.DataSet.FieldByName('Month9').AsFloat;
    value := TUstr.CommaNumber(FloatToStr(RoundTo(fL_Amount,-2)));

end;

procedure TrptSO8B20B_4.QRLabel13Print(sender: TObject; var Value: String);
var
    fL_Amount : Double;
begin
    fL_Amount := self.DataSet.FieldByName('Month10').AsFloat;
    value := TUstr.CommaNumber(FloatToStr(RoundTo(fL_Amount,-2)));
end;

procedure TrptSO8B20B_4.QRLabel14Print(sender: TObject; var Value: String);
var
    fL_Amount : Double;
begin
    fL_Amount := self.DataSet.FieldByName('Month11').AsFloat;
    value := TUstr.CommaNumber(FloatToStr(RoundTo(fL_Amount,-2)));
end;

procedure TrptSO8B20B_4.QRLabel15Print(sender: TObject; var Value: String);
var
    fL_Amount : Double;
begin
    fL_Amount := self.DataSet.FieldByName('Month12').AsFloat;
    value := TUstr.CommaNumber(FloatToStr(RoundTo(fL_Amount,-2)));
end;

procedure TrptSO8B20B_4.QRLabel16Print(sender: TObject; var Value: String);
var
    fL_Amount : Double;
begin
    fL_Amount := self.DataSet.FieldByName('OtherMonth').AsFloat;
    value := TUstr.CommaNumber(FloatToStr(RoundTo(fL_Amount,-2)));
end;

procedure TrptSO8B20B_4.QRLabel17Print(sender: TObject; var Value: String);
var
    fL_Amount : Double;
begin
    fL_Amount := self.DataSet.FieldByName('TotalComm').AsFloat;
    value := TUstr.CommaNumber(FloatToStr(RoundTo(fL_Amount,-2)));
end;

procedure TrptSO8B20B_4.QRLabel33Print(sender: TObject; var Value: String);
begin
    Value := '統計時間: ' + DateTimeToStr(now);
end;

procedure TrptSO8B20B_4.QRLabel34Print(sender: TObject; var Value: String);
begin
    Value := '統計人員: ' + sG_User;
end;

end.
