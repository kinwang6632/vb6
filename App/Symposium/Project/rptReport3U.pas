unit rptReport3U;

interface

uses Windows, SysUtils, Messages, Classes, Graphics, Controls,
  StdCtrls, ExtCtrls, Forms, Quickrpt, QRCtrls;

type
  TrptReport3 = class(TQuickRep)
    QRBand2: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel5: TQRLabel;
    QRSysData1: TQRSysData;
    QRLabel7: TQRLabel;
    QRSysData2: TQRSysData;
    QRLabel8: TQRLabel;
    QRLabel9: TQRLabel;
    QRLabel10: TQRLabel;
    QRLabel11: TQRLabel;
    QRLabel12: TQRLabel;
    QRShape2: TQRShape;
    QRShape3: TQRShape;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QRBand1: TQRBand;
    QRDBText1: TQRDBText;
    QRDBText2: TQRDBText;
    QRDBText3: TQRDBText;
    QRDBText4: TQRDBText;
    QRLabel4: TQRLabel;
    QRBand3: TQRBand;
    QRShape1: TQRShape;
    QRShape4: TQRShape;
    qreCallsAnswered: TQRExpr;
    qreTalkTime: TQRExpr;
    QRLabel6: TQRLabel;
    QRExpr1: TQRExpr;
    QRLabel13: TQRLabel;
    QRDBText5: TQRDBText;
    QRExpr2: TQRExpr;
    procedure QRLabel5Print(sender: TObject; var Value: String);
    procedure QRLabel4Print(sender: TObject; var Value: String);
    procedure QRDBText3Print(sender: TObject; var Value: String);
    procedure QRDBText4Print(sender: TObject; var Value: String);
    procedure qreTalkTimePrint(sender: TObject; var Value: String);
    procedure QRLabel6Print(sender: TObject; var Value: String);
    procedure qreCallsAnsweredPrint(sender: TObject; var Value: String);
    procedure QRExpr1Print(sender: TObject; var Value: String);
    procedure QRDBText5Print(sender: TObject; var Value: String);
    procedure QRExpr2Print(sender: TObject; var Value: String);
  private
    nG_TotalTalkTime, nG_TotalCallsAnswered : Integer;
    function transSecsToStr(nI_Secs : Integer):String;
  public
    sG_RptSDate, sG_RptEDate : String;
  end;

var
  rptReport3: TrptReport3;

implementation

uses dtmMainU;

{$R *.DFM}

procedure TrptReport3.QRLabel5Print(sender: TObject; var Value: String);
begin
    Value := sG_RptSDate + ' ~ ' + sG_RptEDate;
end;

procedure TrptReport3.QRLabel4Print(sender: TObject; var Value: String);
var
    nL_CallsAnswered, nL_TalkTime : Integer;
begin
    nL_CallsAnswered := self.DataSet.FieldByName('CallsAnswered').AsInteger;
    nL_TalkTime := self.DataSet.FieldByName('TalkTime').AsInteger;
    if ((nL_CallsAnswered=0) or (nL_TalkTime=0)) then
      Value := '0'
    else
    begin
      Value := FloatToStr(Round(nL_TalkTime/nL_CallsAnswered));
      Value := transSecsToStr(StrToInt(Value));
    end;
end;

procedure TrptReport3.QRDBText3Print(sender: TObject; var Value: String);
begin
    Value := transSecsToStr(StrToInt(Value));
end;

function TrptReport3.transSecsToStr(nI_Secs: Integer): String;
var
    nL_Secs, nL_Min, nL_Hour : Integer;
begin
    nL_Secs := (nI_Secs mod 60); //眔计
    nL_Min := (nI_Secs div 60); // 眔だ牧计

    nL_Hour := (nL_Min div 60); //眔计
    if (nL_Hour>0) then
    begin
      nL_Min := (nL_Min mod 60); // 眔だ牧计
    end;

    result := Format('%.2d:%.2d:%.2d',[nL_Hour, nL_Min, nL_Secs ]);
end;

procedure TrptReport3.QRDBText4Print(sender: TObject; var Value: String);
begin
    Value := transSecsToStr(StrToInt(Value));
end;

procedure TrptReport3.qreTalkTimePrint(sender: TObject; var Value: String);
begin
    nG_TotalTalkTime := StrToInt(Value);
    Value := transSecsToStr(StrToInt(Value));
end;

procedure TrptReport3.QRLabel6Print(sender: TObject; var Value: String);
begin
    Value := FloatToStr(Round(nG_TotalTalkTime/nG_TotalCallsAnswered));
    Value := transSecsToStr(StrToInt(Value));
end;

procedure TrptReport3.qreCallsAnsweredPrint(sender: TObject;
  var Value: String);
begin
    nG_TotalCallsAnswered := StrToInt(Value);
end;

procedure TrptReport3.QRExpr1Print(sender: TObject; var Value: String);
begin
    Value := transSecsToStr(StrToInt(Value));
end;

procedure TrptReport3.QRDBText5Print(sender: TObject; var Value: String);
begin
    Value := transSecsToStr(StrToInt(Value));
end;

procedure TrptReport3.QRExpr2Print(sender: TObject; var Value: String);
begin
    Value := transSecsToStr(StrToInt(Value));
end;

end.
