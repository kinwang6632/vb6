unit rptReport1U;

interface

uses Windows, SysUtils, Messages, Classes, Graphics, Controls,
  StdCtrls, ExtCtrls, Forms, Quickrpt, QRCtrls;

type

  TrptReport1 = class(TQuickRep)
    QRBand3: TQRBand;
    QRDBText1: TQRDBText;
    QRDBText2: TQRDBText;
    QRDBText3: TQRDBText;
    QRLabel3: TQRLabel;
    QRGroup1: TQRGroup;
    QRBand4: TQRBand;
    QRShape1: TQRShape;
    QRBand2: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel5: TQRLabel;
    QRLabel6: TQRLabel;
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
    QRBand1: TQRBand;
    QRLabel13: TQRLabel;
    QRLabel14: TQRLabel;
    QRLabel15: TQRLabel;
    QRLabel16: TQRLabel;
    QRShape4: TQRShape;
    QRLabel17: TQRLabel;
    QRLabel18: TQRLabel;
    QRLabel19: TQRLabel;
    QRLabel20: TQRLabel;
    QRLabel21: TQRLabel;
    QRLabel22: TQRLabel;
    QRLabel23: TQRLabel;
    QRLabel2: TQRLabel;
    QRDBText4: TQRDBText;
    QRDBText5: TQRDBText;
    QRDBText6: TQRDBText;
    QRDBText7: TQRDBText;
    QRDBText8: TQRDBText;
    QRDBText9: TQRDBText;
    procedure QRLabel3Print(sender: TObject; var Value: String);
    procedure QRDBText3Print(sender: TObject; var Value: String);
    procedure QRLabel1Print(sender: TObject; var Value: String);
    procedure QRLabel5Print(sender: TObject; var Value: String);
    procedure QRLabel6Print(sender: TObject; var Value: String);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
    procedure QRLabel17Print(sender: TObject; var Value: String);
    procedure QRBand3BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure QRLabel18Print(sender: TObject; var Value: String);
    procedure QRLabel19Print(sender: TObject; var Value: String);
    procedure QRLabel20Print(sender: TObject; var Value: String);
    procedure QRLabel21Print(sender: TObject; var Value: String);
    procedure QRGroup1BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure QRLabel22Print(sender: TObject; var Value: String);
    procedure QRLabel23Print(sender: TObject; var Value: String);
    procedure QRLabel2Print(sender: TObject; var Value: String);
    procedure QRGroup1AfterPrint(Sender: TQRCustomBand;
      BandPrinted: Boolean);
    procedure QRBand3AfterPrint(Sender: TQRCustomBand;
      BandPrinted: Boolean);
    procedure QRDBText1Print(sender: TObject; var Value: String);
    procedure QRDBText2Print(sender: TObject; var Value: String);
    procedure QRLabel13Print(sender: TObject; var Value: String);
    procedure QRLabel14Print(sender: TObject; var Value: String);
    procedure QRLabel15Print(sender: TObject; var Value: String);
    procedure QRLabel16Print(sender: TObject; var Value: String);
    procedure QRBand1AfterPrint(Sender: TQRCustomBand;
      BandPrinted: Boolean);
    procedure QRDBText4Print(sender: TObject; var Value: String);
    procedure QRDBText5Print(sender: TObject; var Value: String);
    procedure QRDBText6Print(sender: TObject; var Value: String);
  private
    nG_TotalGroupCount : Integer;
    nG_TotalValidCount : Integer;
    nG_TotalInvalidCount : Integer;
    nG_GroupOccurrence: Integer;
    bG_PrintGroupDetail : boolean;
    sG_GroupActivityCode  :  String;
    sG_Sum1,sG_Sum2,sG_Sum3,sG_Sum4,sG_Sum5,sG_Sum6,sG_Sum7,sG_Sum8:String;
    sG_ReportDataID,  sG_ReportDataDesc, sG_ReportDataPercentage : String;
    sG_ReportDataEntry1,
    sG_ReportDataEntry2,
    sG_ReportDataEntry3,
    sG_ReportDataEntry: String;
    function getGroupName(nI_Group : Integer):String;
    function getFullDataString(sI_ID,  sI_Desc, sI_Entry1, sI_Entry2, sI_Entry3, sI_Entry, sI_Percentage:String):String;
    procedure writeFullDataStringToStrList(sI_ID,  sI_FullDataStr:String);
  public
    bG_ShowInvalidDetail : boolean;
    sG_RptSDate, sG_RptEDate : String;
    nG_GroupOccurrance : Integer;
    fG_GroupPercent: double;
    nG_CompNdx : Integer;
    nG_NtyTOccurrences, nG_LscTOccurrences,nG_BestLscTOccurrences,nG_ShLscTOccurrences,  nG_CyLscTOccurrences : Integer;
    nG_BestTOccurrences,   nG_ShTOccurrences, nG_CyTOccurrences   :   Integer;
  end;

var
  rptReport1: TrptReport1;

implementation

uses dtmMainU, frmReport1U, Ustru;

{$R *.DFM}

procedure TrptReport1.QRLabel3Print(sender: TObject; var Value: String);
var
    nL_Code, nL_Occurrence : Integer;
    sL_Value, sL_Percent : String;
    fL_Percent : Double;
begin
    nL_Code := self.DataSet.fieldByName('sActivityCode').AsInteger;
    nL_Occurrence := self.DataSet.fieldByName('nOccurrence').AsInteger;
    case nG_CompNdx of
      0://NTY
       begin
        sL_Percent := Format('%8.1f',[(nL_Occurrence/nG_NtyTOccurrences)*100]);
        fL_Percent := StrToFloat(sL_Percent);
        sL_Value := sL_Percent + '%';
       end;
      1://BEST
       begin
        sL_Percent := Format('%8.1f',[(nL_Occurrence/nG_BestTOccurrences)*100]);
        fL_Percent := StrToFloat(sL_Percent);
        sL_Value := sL_Percent + '%';
       end;
      2://SH
       begin
        sL_Percent := Format('%8.1f',[(nL_Occurrence/nG_ShTOccurrences)*100]);
        fL_Percent := StrToFloat(sL_Percent);
        sL_Value := sL_Percent + '%';
       end;
      3://CY
       begin
        sL_Percent := Format('%8.1f',[(nL_Occurrence/nG_CyTOccurrences)*100]);
        fL_Percent := StrToFloat(sL_Percent);
        sL_Value := sL_Percent + '%';
       end;
      4://LSC
       begin
        sL_Percent := Format('%8.1f',[(nL_Occurrence/nG_LscTOccurrences)*100]);
        fL_Percent := StrToFloat(sL_Percent);
        sL_Value := sL_Percent + '%';
       end;
      5://BEST-LSC
       begin
        sL_Percent := Format('%8.1f',[(nL_Occurrence/nG_BestLscTOccurrences)*100]);
        fL_Percent := StrToFloat(sL_Percent);
        sL_Value := sL_Percent + '%';
       end;
      6://SH-LSC
       begin
        sL_Percent := Format('%8.1f',[(nL_Occurrence/nG_ShLscTOccurrences)*100]);
        fL_Percent := StrToFloat(sL_Percent);
        sL_Value := sL_Percent + '%';
       end;
      7://CY-LSC
       begin
        sL_Percent := Format('%8.1f',[(nL_Occurrence/nG_CyLscTOccurrences)*100]);
        fL_Percent := StrToFloat(sL_Percent);
        sL_Value := sL_Percent + '%';
       end;
    end;
//howrd    fG_GroupPercent := fG_GroupPercent + fL_Percent;
    Value :=  sL_Value;
    sG_ReportDataPercentage := Value;    
end;

function TrptReport1.getGroupName(nI_Group: Integer): String;
var
    sL_Result : String;
begin
    case nG_CompNdx of
     0,1,2,3: //NTY, Best, SH, CY
      begin
          if (nI_Group=1) then
            sL_Result := 'Service Calls'
          else if (nI_Group=2) then
            sL_Result := 'Outage'
          else if (nI_Group=3) then
            sL_Result := 'Request Information'
          else if (nI_Group=4) then
            sL_Result := 'Payment Related'
          else if (nI_Group=5) then
            sL_Result := 'Schedule Installs'
          else if (nI_Group=6) then
            sL_Result := 'Schedule COS'
          else if (nI_Group=7) then
            sL_Result := 'Dispatch'
          else if (nI_Group=8) then
            sL_Result := 'On-Line Troubleshooting'
          else if (nI_Group=9) then
            sL_Result := 'Customer Complaints'
          else if (nI_Group=10) then
            sL_Result := 'Special Request Orders'
          else if (nI_Group=11) then
            sL_Result := 'Disconnect'
          else if (nI_Group=12) then
            sL_Result := 'Others';
      end;

    end;
    result :=  sL_Result;
end;

procedure TrptReport1.QRDBText3Print(sender: TObject; var Value: String);
var
    nL_Occurrence : Integer;
begin
    nL_Occurrence := StrToInt(Value);
    nG_GroupOccurrance := nG_GroupOccurrance + nL_Occurrence;
    sG_ReportDataEntry  :=  Value;

end;

procedure TrptReport1.QRLabel1Print(sender: TObject; var Value: String);
begin
    case nG_CompNdx of
      0:
        Value := 'Nan Taoyuan CATV Co., Ltd';
      1:
        Value := 'BEST CATV Co., Ltd';
      2:
        Value := 'SH CATV Co., Ltd';
      3:
        Value := 'CY CATV Co., Ltd';
      4:
        Value := 'Lightning Speed Communications';
      5:
        Value := 'BEST-LSC Co., Ltd';
      6:
        Value := 'SH-LSC Co., Ltd';
      7:
        Value := 'CY-LSC Co., Ltd';
    end;


end;

procedure TrptReport1.QRLabel5Print(sender: TObject; var Value: String);
begin
    Value := sG_RptSDate + ' ~ ' + sG_RptEDate;
end;

procedure TrptReport1.QRLabel6Print(sender: TObject; var Value: String);
begin
    Value := 'Wrap Up Code Distribution Report';
end;

procedure TrptReport1.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
    nG_TotalGroupCount := 0;
    nG_TotalValidCount := 0;
    nG_TotalInvalidCount := 0;
end;

procedure TrptReport1.QRLabel17Print(sender: TObject; var Value: String);
begin
    Value := IntToStr(nG_TotalGroupCount);
    sG_Sum2  :=  Value;
end;

procedure TrptReport1.QRBand3BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
var
    sL_ActivityCode : String;
    nL_Occurrence : Integer;
begin

    sL_ActivityCode := Trim(self.DataSet.fieldByName('sActivityCode').AsString) ;
    nL_Occurrence := self.DataSet.fieldByName('nOccurrence').AsInteger;

    if (INVALID_GROUP_ID<>self.DataSet.FieldByName('nGroup').AsInteger) then
    begin
      nG_TotalValidCount := nG_TotalValidCount + nL_Occurrence;
      if (nL_Occurrence>0)then
        nG_TotalGroupCount := nG_TotalGroupCount + 1;

    end
    else
    begin
      if (not bG_ShowInvalidDetail) then
        PrintBand := false;
      if (sL_ActivityCode<>'0')then
        nG_TotalInvalidCount := nG_TotalInvalidCount + nL_Occurrence
    end;

    if  not bG_PrintGroupDetail then
        PrintBand := false;

end;

procedure TrptReport1.QRLabel18Print(sender: TObject; var Value: String);
begin
    Value := IntToStr(nG_TotalValidCount);
    sG_Sum4  :=  Value;    
end;

procedure TrptReport1.QRLabel19Print(sender: TObject; var Value: String);
begin
    Value := IntToStr(nG_TotalInvalidCount);
    sG_Sum6  :=  Value;    
end;

procedure TrptReport1.QRLabel20Print(sender: TObject; var Value: String);
begin
    Value := IntToStr(nG_TotalValidCount + nG_TotalInvalidCount);
    sG_Sum8  :=  Value;    
end;

procedure TrptReport1.QRLabel21Print(sender: TObject; var Value: String);
var
    nL_Group : Integer;
begin
    nL_Group   :=   self.DataSet.FieldByName('nGroup').AsInteger;
    if (INVALID_GROUP_ID=nL_Group) then
      Value := 'Total Invalid  Wrap Up Codes Entered'
    else
      Value :=  getGroupName(nL_Group);
    sG_ReportDataDesc  :=  Value;      
end;

procedure TrptReport1.QRGroup1BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
var
    sL_Group   :   String;
    nL_MemberCount : Integer;
begin

    sL_Group   :=   self.DataSet.FieldByName('nGroup').AsString;
    dtmMain.getGroupInfo(sL_Group,   nG_GroupOccurrence, nL_MemberCount,   fG_GroupPercent,  sG_GroupActivityCode);

    if nL_MemberCount=1  then
    begin
      bG_PrintGroupDetail  :=  False;
      QRDBText7.Enabled := True;
      QRDBText8.Enabled := True;
      QRDBText9.Enabled := True;
    end else
    begin
      bG_PrintGroupDetail := True;
      QRDBText7.Enabled := False;
      QRDBText8.Enabled := False;
      QRDBText9.Enabled := False;
    end;
    if (INVALID_GROUP_ID=self.DataSet.FieldByName('nGroup').AsInteger) then
    begin
      if (not bG_ShowInvalidDetail) then
        PrintBand := false;
    end;

end;

procedure TrptReport1.QRLabel22Print(sender: TObject; var Value: String);
begin
    Value := IntToStr(nG_GroupOccurrence);
    sG_ReportDataEntry  :=  Value;    
end;

procedure TrptReport1.QRLabel23Print(sender: TObject; var Value: String);
begin
    Value := FloatToStr(fG_GroupPercent) + '%';
    sG_ReportDataPercentage  :=  Value;    
end;

procedure TrptReport1.QRLabel2Print(sender: TObject; var Value: String);
begin
    Value  :=  sG_GroupActivityCode ;
    if  Value='' then
      sG_ReportDataID  :=  SPECIAL_DATA_ID
    else
      sG_ReportDataID  :=  Value;
end;

procedure TrptReport1.QRGroup1AfterPrint(Sender: TQRCustomBand;
  BandPrinted: Boolean);
var
  sL_FullDataStr : String;
  a1, a2, a3: Integer;
begin
  dtmMain.getGroupInfo2( DataSet.FieldByName('nGroup').AsString, a1, a2, a3 );
  sG_ReportDataEntry1 := IntToStr( a1 );
  sG_ReportDataEntry2 := IntToStr( a2 );
  sG_ReportDataEntry3 := IntToStr( a3 );
  sL_FullDataStr :=  GetFullDataString( sG_ReportDataID, sG_ReportDataDesc,
    sG_ReportDataEntry1, sG_ReportDataEntry2, sG_ReportDataEntry3,
    sG_ReportDataEntry, sG_ReportDataPercentage );
  writeFullDataStringToStrList(sG_ReportDataID,sL_FullDataStr);
end;

function TrptReport1.getFullDataString(
sI_ID,  sI_Desc, sI_Entry1, sI_Entry2, sI_Entry3, sI_Entry, sI_Percentage:String): String;
var
    sL_FullDataStr,  sI_RealID  :  String;
    sL_ID, sL_Desc, sL_Entry1, sL_Entry2, sL_Entry3, sL_Entry,  sL_Percentage: String;
begin
    if sI_ID=SPECIAL_DATA_ID  then
      sI_RealID  := ''
    else
      sI_RealID  :=  sI_ID;

    sL_ID := TUstr.AddString(sI_RealID,' ',false,5);
    sL_Desc := TUstr.AddString(sI_Desc,' ',false,50);
    sL_Entry1 := TUstr.AddString(sI_Entry1,' ',false,10);
    sL_Entry2 := TUstr.AddString(sI_Entry2,' ',false,10);
    sL_Entry3 := TUstr.AddString(sI_Entry3,' ',false,10);
    sL_Entry := TUstr.AddString(sI_Entry,' ',false,10);
    sL_Percentage := TUstr.AddString(sI_Percentage,' ',false,10);

    sL_FullDataStr  :=
      sL_ID  +  sL_Desc  + sL_Entry1 + sL_Entry2 + sL_Entry3 + sL_Entry + sL_Percentage;
    result := sL_FullDataStr;
end;

procedure TrptReport1.QRBand3AfterPrint(Sender: TQRCustomBand;
  BandPrinted: Boolean);
var
    sL_FullDataStr : STring;

begin

    sL_FullDataStr :=
    getFullDataString(sG_ReportDataID,sG_ReportDataDesc,
      sG_ReportDataEntry1,
      sG_ReportDataEntry2,
      sG_ReportDataEntry3,
      sG_ReportDataEntry,
      sG_ReportDataPercentage);
    writeFullDataStringToStrList(sG_ReportDataID,sL_FullDataStr);


end;

procedure TrptReport1.QRDBText1Print(sender: TObject; var Value: String);
begin
    sG_ReportDataID := Value;                
end;

procedure TrptReport1.QRDBText2Print(sender: TObject; var Value: String);
begin
  sG_ReportDataDesc := Value;
end;

procedure TrptReport1.writeFullDataStringToStrList(sI_ID,
  sI_FullDataStr: String);
var
    sI_RealID : String;
    L_Obj : TReport1Obj;
begin

    if   (sI_ID=SPECIAL_DATA_ID)  or  (frmReport1.G_Report1StrList.IndexOf(sI_ID)=-1) then
    begin
      L_Obj := TReport1Obj.Create;
      L_Obj.FullDataString := sI_FullDataStr;


      frmReport1.G_Report1StrList.AddObject(sI_ID,L_Obj);
    end;
end;

procedure TrptReport1.QRLabel13Print(sender: TObject; var Value: String);
begin
    sG_Sum1  :=  Value;
end;

procedure TrptReport1.QRLabel14Print(sender: TObject; var Value: String);
begin
    sG_Sum3  :=  Value;
end;

procedure TrptReport1.QRLabel15Print(sender: TObject; var Value: String);
begin
    sG_Sum5  :=  Value;
end;

procedure TrptReport1.QRLabel16Print(sender: TObject; var Value: String);
begin
    sG_Sum7  :=  Value;
end;

procedure TrptReport1.QRBand1AfterPrint(Sender: TQRCustomBand;
  BandPrinted: Boolean);
var
  sL_FullDataStr :String;  
begin
    sL_FullDataStr :=
      getFullDataString(SPECIAL_DATA_ID,sG_Sum1, EmptyStr, EmptyStr, EmptyStr, sG_Sum2,'');
    writeFullDataStringToStrList(SPECIAL_DATA_ID,sL_FullDataStr);

    sL_FullDataStr :=
      getFullDataString(SPECIAL_DATA_ID,sG_Sum3,EmptyStr, EmptyStr, EmptyStr,  sG_Sum4,'');
    writeFullDataStringToStrList(SPECIAL_DATA_ID,sL_FullDataStr);

    sL_FullDataStr :=
      getFullDataString(SPECIAL_DATA_ID,sG_Sum5, EmptyStr, EmptyStr, EmptyStr, sG_Sum6,'');
    writeFullDataStringToStrList(SPECIAL_DATA_ID,sL_FullDataStr);

    sL_FullDataStr :=
      getFullDataString(SPECIAL_DATA_ID,sG_Sum7,EmptyStr, EmptyStr, EmptyStr, sG_Sum8,'');
    writeFullDataStringToStrList(SPECIAL_DATA_ID,sL_FullDataStr);

end;

procedure TrptReport1.QRDBText4Print(sender: TObject; var Value: String);
begin
 sG_ReportDataEntry1  :=  Value;
end;

procedure TrptReport1.QRDBText5Print(sender: TObject; var Value: String);
begin
  sG_ReportDataEntry2  :=  Value;
end;

procedure TrptReport1.QRDBText6Print(sender: TObject; var Value: String);
begin
  sG_ReportDataEntry3  :=  Value;
end;

end.
