unit rptReport2U;

interface

uses Windows, SysUtils, Messages, Classes, Graphics, Controls,
  StdCtrls, ExtCtrls, Forms, Quickrpt, QRCtrls, Dialogs;

type
  TrptReport2 = class(TQuickRep)
    QRBand2: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel5: TQRLabel;
    QRSysData1: TQRSysData;
    QRLabel7: TQRLabel;
    QRSysData2: TQRSysData;
    QRLabel8: TQRLabel;
    QRBand1: TQRBand;
    QRDBText1: TQRDBText;
    QRDBText2: TQRDBText;
    QRDBText3: TQRDBText;
    QRDBText4: TQRDBText;
    QRLabel4: TQRLabel;
    QRBand3: TQRBand;
    qreCallsAnswered: TQRExpr;
    qreTalkTime: TQRExpr;
    QRLabel6: TQRLabel;
    QRLabel13: TQRLabel;
    QRDBText5: TQRDBText;
    QRLabel15: TQRLabel;
    QRGroup1: TQRGroup;
    QRLabel9: TQRLabel;
    QRLabel10: TQRLabel;
    QRLabel11: TQRLabel;
    QRLabel12: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel14: TQRLabel;
    QRLabel17: TQRLabel;
    QRDBText6: TQRDBText;
    QRLabel21: TQRLabel;
    QRLabel22: TQRLabel;
    QRExpr8: TQRExpr;
    QRExpr9: TQRExpr;
    QRExpr6: TQRExpr;
    QRLabel23: TQRLabel;
    QRLabel25: TQRLabel;
    QRLabel20: TQRLabel;
    QRDBText7: TQRDBText;
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
    QRLabel16: TQRLabel;
    QRExpr1: TQRExpr;
    QRLabel18: TQRLabel;
    QRLabel19: TQRLabel;
    QRLabel24: TQRLabel;
    QRLabel37: TQRLabel;
    QRShape1: TQRShape;
    QRShape2: TQRShape;
    QRShape3: TQRShape;
    procedure QRLabel5Print(sender: TObject; var Value: String);
    procedure QRLabel4Print(sender: TObject; var Value: String);
    procedure QRDBText3Print(sender: TObject; var Value: String);
    procedure QRDBText4Print(sender: TObject; var Value: String);
    procedure qreTalkTimePrint(sender: TObject; var Value: String);
    procedure QRLabel6Print(sender: TObject; var Value: String);
    procedure qreCallsAnsweredPrint(sender: TObject; var Value: String);
    procedure QRExpr1Print(sender: TObject; var Value: String);
    procedure QRDBText5Print(sender: TObject; var Value: String);
    procedure QRBand1BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure QRLabel22Print(sender: TObject; var Value: String);
    procedure QRLabel15Print(sender: TObject; var Value: String);
    procedure QRExpr9Print(sender: TObject; var Value: String);
    procedure QRExpr2Print(sender: TObject; var Value: String);
    procedure QRExpr6Print(sender: TObject; var Value: String);
    procedure QRLabel20Print(sender: TObject; var Value: String);
    procedure QRLabel16Print(sender: TObject; var Value: String);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
    procedure QRExpr7Print(sender: TObject; var Value: String);
    procedure QRLabel23Print(sender: TObject; var Value: String);
    procedure QRExpr10Print(sender: TObject; var Value: String);
    procedure QRDBText6Print(sender: TObject; var Value: String);
    procedure QRDBText2Print(sender: TObject; var Value: String);
    procedure QRLabel25Print(sender: TObject; var Value: String);
    procedure QRLabel24Print(sender: TObject; var Value: String);
    procedure QRLabel37Print(sender: TObject; var Value: String);
    procedure QRLabel36Print(sender: TObject; var Value: String);
    procedure QRBand2BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure QRGroup1BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure QRBand3BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
  private
    nG_CurCallAnswered : Integer;
    nG_TotalTalkTime, nG_TotalCallsAnswered, nG_TotalCallsPresented: Integer;
    nG_LoggedIn1,nG_LoggedIn2, nG_NotReady1,nG_NotReady2 : Integer;
    nG_TotalShortCalls: Integer;
    FExported: Boolean;
    function transSecsToStr(nI_Secs : Integer):String;
  public
    sG_RptSDate, sG_RptEDate : String;
    property Exported: Boolean read FExported write FExported;
  end;

var
  rptReport2: TrptReport2;

implementation

uses dtmMainU;

{$R *.DFM}

procedure TrptReport2.QRLabel5Print(sender: TObject; var Value: String);
begin
    Value := sG_RptSDate + ' ~ ' + sG_RptEDate;
end;

procedure TrptReport2.QRLabel4Print(sender: TObject; var Value: String);
var
    nL_CallsAnswered, nL_TalkTime : Integer;
begin
    nL_CallsAnswered := self.DataSet.FieldByName('CallsAnswered').AsInteger;
    nG_TotalCallsAnswered := nG_TotalCallsAnswered + nL_CallsAnswered;//howard
    nL_TalkTime := self.DataSet.FieldByName('TalkTime').AsInteger;
    if (nL_CallsAnswered<>0) then
    begin
      Value := FloatToStr(Round(nL_TalkTime/nL_CallsAnswered));
      Value := transSecsToStr(StrToInt(Value));
    end
    else
      Value := '00:00:00';
end;

procedure TrptReport2.QRLabel24Print(sender: TObject; var Value: String);
var
    nL_CallsAnswered, nL_TalkTime, nL_ShortCall: Integer;
begin
    nL_CallsAnswered := self.DataSet.FieldByName('CallsAnswered').AsInteger;
    nL_TalkTime := self.DataSet.FieldByName('TalkTime').AsInteger;
    nL_ShortCall := self.DataSet.FieldByName( 'ShortCallsAns' ).AsInteger;
    nG_TotalShortCalls := ( nG_TotalShortCalls + nL_ShortCall );
    Value := '00:00:00';
    if ( nL_CallsAnswered <> 0 ) and
       ( ( nL_CallsAnswered - nL_ShortCall ) <> 0 )then
    begin
      Value := FloatToStr(Round(nL_TalkTime/(nL_CallsAnswered-nL_ShortCall)));
      Value := transSecsToStr(StrToInt(Value));
    end;
end;


procedure TrptReport2.QRDBText3Print(sender: TObject; var Value: String);
begin
    Value := transSecsToStr(StrToInt(Value));
end;

function TrptReport2.transSecsToStr(nI_Secs: Integer): String;
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

procedure TrptReport2.QRDBText4Print(sender: TObject; var Value: String);
begin
    Value := transSecsToStr(StrToInt(Value));
end;

procedure TrptReport2.qreTalkTimePrint(sender: TObject; var Value: String);
begin
    nG_TotalTalkTime := StrToInt(Value);
    Value := transSecsToStr(StrToInt(Value));
end;

procedure TrptReport2.QRLabel6Print(sender: TObject; var Value: String);
begin
//    howard
    Value := FloatToStr(Round(nG_TotalTalkTime/nG_TotalCallsAnswered));
    Value := transSecsToStr(StrToInt(Value));
end;

procedure TrptReport2.QRLabel37Print(sender: TObject; var Value: String);
begin
    Value := FloatToStr(Round(nG_TotalTalkTime/( nG_TotalCallsAnswered - nG_TotalShortCalls )));
    Value := transSecsToStr(StrToInt(Value));
end;

procedure TrptReport2.qreCallsAnsweredPrint(sender: TObject;
  var Value: String);
begin
//    nG_TotalCallsPresented := StrToInt(Value);
end;

procedure TrptReport2.QRExpr1Print(sender: TObject; var Value: String);
begin
    nG_NotReady1 := StrToInt(Value);
//    showmessage('NR1=='+inttostr(nG_NotReady1));
    Value := transSecsToStr(StrToInt(Value));
end;

procedure TrptReport2.QRDBText5Print(sender: TObject; var Value: String);
begin
    Value := transSecsToStr(StrToInt(Value));
end;

procedure TrptReport2.QRBand1BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
    nG_CurCallAnswered := self.DataSet.fieldByName('CallsAnswered').AsInteger;

end;

procedure TrptReport2.QRLabel22Print(sender: TObject; var Value: String);
var
    dL_Value : Double;
begin

//    howard
    if self.DataSet.fieldByName('CallsPresented').AsInteger<>0 then
    begin
      dL_Value := self.DataSet.fieldByName('CallsAnswered').AsInteger/self.DataSet.fieldByName('CallsPresented').AsInteger*100;
      Value := Format('%.2f',[dL_Value]) + '%';
    end
    else
      Value := '0';  

end;

procedure TrptReport2.QRLabel15Print(sender: TObject; var Value: String);
var
    dL_Value : Double;
    nL_NotReadyTime, nL_LoggedInTime : Integer;
begin
//    howard

    nL_NotReadyTime := StrToIntDef(self.DataSet.fieldByName('NotReadyTime').AsString,0);
    nL_LoggedInTime := StrToIntDef(self.DataSet.fieldByName('LoggedInTime').AsString,0);
    if (nL_NotReadyTime = 0 ) or  ( nL_LoggedInTime=0 ) then
      dL_Value := 0
    else
      dL_Value := nL_NotReadyTime/nL_LoggedInTime*100;
    Value := Format('%.2f',[dL_Value]) + '%';

end;

procedure TrptReport2.QRExpr9Print(sender: TObject; var Value: String);
begin
    nG_NotReady2 := StrToInt(Value);
//    showmessage('NR2=='+inttostr(nG_NotReady2));    
    Value := transSecsToStr(StrToInt(Value));
end;

procedure TrptReport2.QRExpr2Print(sender: TObject; var Value: String);
begin
    nG_LoggedIn1 := StrToInt(Value);
//    showmessage('LG1=='+inttostr(nG_LoggedIn1));    
    Value := transSecsToStr(StrToInt(Value));

end;

procedure TrptReport2.QRExpr6Print(sender: TObject; var Value: String);
begin
    nG_LoggedIn2 := StrToInt(Value);
//    showmessage('LG1=='+inttostr(nG_LoggedIn2));
    Value := transSecsToStr(StrToInt(Value));

end;

procedure TrptReport2.QRLabel20Print(sender: TObject; var Value: String);
var
    dL_Value : double;
begin
//    Value := '';
//    showmessage('T--LG2=='+inttostr(nG_LoggedIn2));
//    showmessage('T--NR2=='+inttostr(nG_NotReady2));

    if nG_LoggedIn2<>0 then
    begin

    dL_Value := (nG_NotReady2/nG_LoggedIn2)*100;
    Value := Format('%.2f',[dL_Value]) + '%';
    end;

end;

procedure TrptReport2.QRLabel16Print(sender: TObject; var Value: String);
var
    dL_Value : double;
begin
//    Value := '';
//    Value := IntToStr(nG_NotReady1) + '==' + IntToStr(nG_LoggedIn1);
//    showmessage('T--LG1=='+inttostr(nG_LoggedIn1));
//    showmessage('T--NR1=='+inttostr(nG_NotReady1));

    if nG_LoggedIn1<>0 then
    begin
      dL_Value := (nG_NotReady1/nG_LoggedIn1)*100;
      Value := Format('%.2f',[dL_Value]) + '%';
    end;
    
end;

procedure TrptReport2.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin

    nG_LoggedIn1 := 0;
    nG_LoggedIn2 := 0;
    nG_NotReady1 := 0;
    nG_NotReady2 := 0;
    nG_TotalCallsPresented := 0;
    nG_TotalCallsAnswered := 0;
    nG_TotalShortCalls := 0;
end;

procedure TrptReport2.QRExpr7Print(sender: TObject; var Value: String);
begin
    nG_TotalCallsPresented := StrToInt(Value);
end;

procedure TrptReport2.QRLabel23Print(sender: TObject; var Value: String);
var
    dL_Value : double;
begin
//    value := Inttostr(nG_TotalCallsAnswered)+'=='+ IntTostr(nG_TotalCallsPresented);

    dL_Value := (nG_TotalCallsAnswered/nG_TotalCallsPresented)*100;
    Value := Format('%.2f',[dL_Value]) + '%';

end;

procedure TrptReport2.QRExpr10Print(sender: TObject; var Value: String);
begin
    nG_TotalCallsAnswered := StrToInt(Value);
end;

procedure TrptReport2.QRDBText6Print(sender: TObject; var Value: String);
begin
    nG_TotalCallsPresented := nG_TotalCallsPresented + StrToInt(Value);
end;

procedure TrptReport2.QRDBText2Print(sender: TObject; var Value: String);
begin
    nG_TotalCallsAnswered := nG_TotalCallsAnswered + StrToInt(Value);
end;

procedure TrptReport2.QRLabel25Print(sender: TObject; var Value: String);
var
    dL_Value : double;
begin

    if nG_LoggedIn2<>0 then
    begin

    dL_Value := (nG_NotReady2/nG_LoggedIn2)*100;
    Value := Format('%.2f',[dL_Value]) + '%';
    end;

end;

procedure TrptReport2.QRLabel36Print(sender: TObject; var Value: String);
begin
 if FExported then
   Value := '--------------------------------------------------------------------------------------------------------------------------------'
 else
   Value := EmptyStr;
end;

procedure TrptReport2.QRBand2BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
  QRShape1.Enabled := not FExported;
end;

procedure TrptReport2.QRGroup1BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
  QRShape2.Enabled := not FExported;
end;

procedure TrptReport2.QRBand3BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
  QRShape3.Enabled := not FExported;
end;

end.
