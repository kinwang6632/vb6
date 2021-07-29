unit rptSO8A40AU;

interface

uses Windows, SysUtils, Messages, Classes, Graphics, Controls,
  StdCtrls, ExtCtrls, Forms, QuickRpt, QRCtrls, Dialogs ,Variants ,DB ,DBClient;

type
  //儲存各頻道總收視戶數資料
  TRptChannelViewData = class(TObject)
    sChannel_ID : String;
    sChannel_Name : String;
    sChannelTotalViewCounts : String;
  end;

  TrptSO8A40A = class(TQuickRep)
    DetailBand1: TQRBand;
    PageFooterBand1: TQRBand;
    PageHeaderBand1: TQRBand;
    QRLabel1: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel5: TQRLabel;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    lblChannelName: TQRDBText;
    lblFirstCount: TQRDBText;
    lblLastCount: TQRDBText;
    lblEmcAmt: TQRDBText;
    lblSoAmt: TQRDBText;
    lblProvideAmt: TQRDBText;
    lblSubTotalAmt: TQRDBText;
    lblCompName: TQRLabel;
    lblBelongYM: TQRLabel;
    lblRealDate: TQRLabel;
    QRShape1: TQRShape;
    QRLabel14: TQRLabel;
    lblAddCount: TQRLabel;
    QRLabel17: TQRLabel;
    QRLabel19: TQRLabel;
    QRLabel20: TQRLabel;
    QRLabel21: TQRLabel;
    QRGroup1: TQRGroup;
    QRBand1: TQRBand;
    QRLabel22: TQRLabel;
    QRLabel18: TQRLabel;
    lblInCome: TQRLabel;
    lblOutCome: TQRLabel;
    SummaryBand1: TQRBand;
    QRLabel8: TQRLabel;
    QRLabel9: TQRLabel;
    QRLabel16: TQRLabel;
    QRLabel15: TQRLabel;
    QRLabel23: TQRLabel;
    QRLabel10: TQRLabel;
    QRLabel11: TQRLabel;
    QRLabel12: TQRLabel;
    QRLabel13: TQRLabel;
    QRShape2: TQRShape;
    QRLabel24: TQRLabel;
    QRLabel25: TQRLabel;
    QRLabel26: TQRLabel;
    QRLabel27: TQRLabel;
    QRLabel28: TQRLabel;
    QRLabel29: TQRLabel;
    QRLabel30: TQRLabel;
    QRLabel31: TQRLabel;
    QRLabel32: TQRLabel;
    QRShape3: TQRShape;
    QRLabel33: TQRLabel;
    lblReduceCount: TQRLabel;
    QRLabel34: TQRLabel;
    QRLabel35: TQRLabel;
    procedure lblChannelNamePrint(sender: TObject; var Value: String);
    procedure lblCompNamePrint(sender: TObject; var Value: String);
    procedure lblBelongYMPrint(sender: TObject; var Value: String);
    procedure lblRealDatePrint(sender: TObject; var Value: String);
    procedure lblFirstCountPrint(sender: TObject; var Value: String);
    procedure QRLabel8Print(sender: TObject; var Value: String);
    procedure lblLastCountPrint(sender: TObject; var Value: String);
    procedure QRLabel9Print(sender: TObject; var Value: String);
    procedure lblEmcAmtPrint(sender: TObject; var Value: String);
    procedure QRLabel10Print(sender: TObject; var Value: String);
    procedure lblSoAmtPrint(sender: TObject; var Value: String);
    procedure QRLabel11Print(sender: TObject; var Value: String);
    procedure lblProvideAmtPrint(sender: TObject; var Value: String);
    procedure QRLabel12Print(sender: TObject; var Value: String);
    procedure lblSubTotalAmtPrint(sender: TObject; var Value: String);
    procedure QRLabel13Print(sender: TObject; var Value: String);
    procedure QuickRepPreview(Sender: TObject);
    procedure lblAddCountPrint(sender: TObject; var Value: String);
    procedure QRLabel16Print(sender: TObject; var Value: String);
    procedure QRLabel19Print(sender: TObject; var Value: String);
    procedure QRLabel20Print(sender: TObject; var Value: String);
    procedure QuickRepBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
    procedure QRLabel22Print(sender: TObject; var Value: String);
    procedure lblInComePrint(sender: TObject; var Value: String);
    procedure QRLabel15Print(sender: TObject; var Value: String);
    procedure lblOutComePrint(sender: TObject; var Value: String);
    procedure QRLabel23Print(sender: TObject; var Value: String);
    procedure QRGroup1BeforePrint(Sender: TQRCustomBand;
      var PrintBand: Boolean);
    procedure QRLabel24Print(sender: TObject; var Value: String);
    procedure QRLabel25Print(sender: TObject; var Value: String);
    procedure QRLabel26Print(sender: TObject; var Value: String);
    procedure QRLabel27Print(sender: TObject; var Value: String);
    procedure QRLabel28Print(sender: TObject; var Value: String);
    procedure QRLabel29Print(sender: TObject; var Value: String);
    procedure QRLabel30Print(sender: TObject; var Value: String);
    procedure QRLabel31Print(sender: TObject; var Value: String);
    procedure QRLabel32Print(sender: TObject; var Value: String);
    procedure lblReduceCountPrint(sender: TObject; var Value: String);
    procedure QRLabel34Print(sender: TObject; var Value: String);
    procedure QRLabel35Print(sender: TObject; var Value: String);
  private
    nG_Total_FirstCount,nG_Total_LastCount,nG_Total_AddCount : Integer;
    nG_FirstCount,nG_LastCount,nG_AddCount,nG_ReduceCount : Integer;
    fG_InCome,fG_OutCome,fG_Total_InCome,fG_Total_OutCome : Double;

    fG_Total_EMC_Benefit,fG_Total_So_Benefit,fG_Total_Provider_Benefit,fG_Total_EMC_Income : Double;

    nG_Group_FirstCount,nG_Group_LastCount,nG_Group_AddCount : Integer;
    nG_Group_ReduceCount,nG_Total_ReduceCount : Integer;
    fG_Group_InCome,fG_Group_OutCome,fG_Group_EMC_Income : Double;
    fG_Group_EMC_Benefit,fG_Group_So_Benefit,fG_Group_Provider_Benefit : Double;

    procedure querySQL(sI_SQL : String);
    procedure parseChannelCountsString;
    function getChannelViewCounts(sI_ProductID : String) : String;
    procedure orderByCDSSO114;

  public
    sG_CompName,sG_CompCode,sG_StartDate,sG_EndDate,sG_BelongYM,sG_ShowDetail : String;
    sG_TransData,sG_ProviderGroup : String;
    sG_FormatName,sG_ShowType : String;
    G_ChannelViewDataStrList : TStringList;
    sG_RptRealDate,sG_RptBelongYM : String;
  end;

var
  rptSO8A40A: TrptSO8A40A;

implementation

uses Ustru, frmRptPreviewU, UdateTimeu, dtmMain2U;

{$R *.DFM}
procedure TrptSO8A40A.querySQL(sI_SQL: String);
begin
    with dtmMain2.cdsCom do
    begin
      Close;
      CommandText := sI_SQL;
      Open;
    end;
end;


procedure TrptSO8A40A.lblChannelNamePrint(sender: TObject;
  var Value: String);
var
    sL_SQL,sL_ProductID : String;
begin
    sL_ProductID := dtmMain2.cdsSO114.FieldByName('PRODUCTID').AsString;

    sL_SQL := 'SELECT * FROM So112 WHERE PRODUCT_ID=''' + sL_ProductID + '''';
    querySQL(sL_SQL);
    Value := dtmMain2.cdsCom.FieldByName('PRODUCT_NAME').AsString;
end;

procedure TrptSO8A40A.lblCompNamePrint(sender: TObject; var Value: String);
begin
    Value := sG_CompName;
end;

procedure TrptSO8A40A.lblBelongYMPrint(sender: TObject; var Value: String);
var
    sL_YM : String;
begin

    Value := sG_RptBelongYM;
end;

procedure TrptSO8A40A.lblRealDatePrint(sender: TObject; var Value: String);
begin

    Value := sG_RptRealDate;
end;

procedure TrptSO8A40A.lblFirstCountPrint(sender: TObject; var Value: String);
begin
    nG_FirstCount := dtmMain2.cdsSO114.FieldByName('FIRSTCOUNT').AsInteger;
    Value := TUstr.CommaNumber(IntToStr(nG_FirstCount));

    nG_Group_FirstCount := nG_Group_FirstCount + nG_FirstCount;
    nG_Total_FirstCount := nG_Total_FirstCount + nG_FirstCount;
end;

procedure TrptSO8A40A.QRLabel8Print(sender: TObject; var Value: String);
begin
    Value := TUstr.CommaNumber(IntToStr(nG_Total_FirstCount));
end;

procedure TrptSO8A40A.lblLastCountPrint(sender: TObject; var Value: String);
begin
    nG_LastCount := dtmMain2.cdsSO114.FieldByName('LASTCOUNT').AsInteger;
    Value := TUstr.CommaNumber(IntToStr(nG_LastCount));

    nG_Group_LastCount := nG_Group_LastCount + nG_LastCount;
    nG_Total_LastCount := nG_Total_LastCount + nG_LastCount;
end;

procedure TrptSO8A40A.QRLabel9Print(sender: TObject; var Value: String);
begin
    Value := TUstr.CommaNumber(IntToStr(nG_Total_LastCount));
end;

procedure TrptSO8A40A.lblEmcAmtPrint(sender: TObject; var Value: String);
var
    fL_EMC_Benefit : Double;
begin
    fL_EMC_Benefit := dtmMain2.cdsSO114.FieldByName('EMC_BENEFIT').AsFloat;
    Value := TUstr.CommaNumber(FloatToStr(fL_EMC_Benefit));

    fG_Group_EMC_Benefit := fG_Group_EMC_Benefit + fL_EMC_Benefit;
    fG_Total_EMC_Benefit := fG_Total_EMC_Benefit + fL_EMC_Benefit;
end;

procedure TrptSO8A40A.QRLabel10Print(sender: TObject; var Value: String);
begin
    Value := TUstr.CommaNumber(FloatToStr(fG_Total_EMC_Benefit));
end;

procedure TrptSO8A40A.lblSoAmtPrint(sender: TObject; var Value: String);
var
    fL_So_Benefit : Double;
begin
    fL_So_Benefit := dtmMain2.cdsSO114.FieldByName('So_BENEFIT').AsFloat;
    Value := TUstr.CommaNumber(FloatToStr(fL_So_Benefit));

    fG_Group_So_Benefit := fG_Group_So_Benefit + fL_So_Benefit;
    fG_Total_So_Benefit := fG_Total_So_Benefit + fL_So_Benefit;
end;

procedure TrptSO8A40A.QRLabel11Print(sender: TObject; var Value: String);
begin
    Value := TUstr.CommaNumber(FloatToStr(fG_Total_So_Benefit));
end;

procedure TrptSO8A40A.lblProvideAmtPrint(sender: TObject;
  var Value: String);
var
    fL_Provider_Benefit : Double;
begin
    fL_Provider_Benefit := dtmMain2.cdsSO114.FieldByName('Provider_BENEFIT').AsFloat;
    Value := TUstr.CommaNumber(FloatToStr(fL_Provider_Benefit));

    fG_Group_Provider_Benefit := fG_Group_Provider_Benefit + fL_Provider_Benefit;
    fG_Total_Provider_Benefit := fG_Total_Provider_Benefit + fL_Provider_Benefit;
end;

procedure TrptSO8A40A.QRLabel12Print(sender: TObject; var Value: String);
begin
    Value := TUstr.CommaNumber(FloatToStr(fG_Total_Provider_Benefit));
end;

procedure TrptSO8A40A.lblSubTotalAmtPrint(sender: TObject;
  var Value: String);
var
    fL_EMC_Income,fL_InCome,fL_OutCome : Double;
begin
    //fL_EMC_Income := dtmMain2.cdsSO114.FieldByName('EMC_Income').AsFloat;
    //淨額 = 收入 - 退費
    fL_InCome := dtmMain2.cdsSO114.FieldByName('INCOME').AsFloat;
    fL_OutCome := dtmMain2.cdsSO114.FieldByName('OUTCOME').AsFloat;
    fL_EMC_Income := fL_InCome - fL_OutCome;
    Value := TUstr.CommaNumber(FloatToStr(fL_EMC_Income));

    fG_Group_EMC_Income := fG_Group_EMC_Income + fL_EMC_Income;
    fG_Total_EMC_Income := fG_Total_EMC_Income + fL_EMC_Income;
end;

procedure TrptSO8A40A.QRLabel13Print(sender: TObject; var Value: String);
begin
    Value := TUstr.CommaNumber(FloatToStr(fG_Total_EMC_Income));
end;

procedure TrptSO8A40A.QuickRepPreview(Sender: TObject);
var
    frmPreView : TfrmRptPreview;
    sL_FullPathRptControlIniFileName,sL_YM : String;
begin
  //orderByCDSSO114;
  sG_RptRealDate := '實收日期: ' + TUdateTime.GetFormatDateStr(sG_StartDate,'#',false,true)
                    + ' ~ ' + TUdateTime.GetFormatDateStr(sG_EndDate,'#',false,true);


  sL_YM := Copy(sG_BelongYM,1,4) + '/' + Copy(sG_BelongYM,5,6);
  sG_RptBelongYM := '歸屬年月: ' + TUdateTime.GetFormatDateStr(sL_YM,'#',false,true);


  frmPreView := TfrmRptPreview.Create(nil);

  frmPreView.G_ActiveDataSet := self.DataSet;

  frmPreView.sG_ShowDetail := self.sG_ShowDetail;
  frmPreView.sG_TransData := self.sG_TransData;
  frmPreView.sG_CompCode := self.sG_CompCode;
  frmPreView.sG_BelongYM := self.sG_BelongYM;

  //ToExcel 會用到
  frmPreView.sG_RptName := '頻道拆帳明細表';
  frmPreView.sG_RptCompName := sG_CompName;
  frmPreView.sG_RptRealDate := sG_RptRealDate;
  frmPreView.sG_RptBelongYM := sG_RptBelongYM;
  frmPreView.sG_ShowType := sG_ShowType;

  frmPreView.Report := Self;
  frmPreView.QRPreview.QRPrinter := Self.QRPrinter;
  frmPreView.Show;
end;

procedure TrptSO8A40A.lblAddCountPrint(sender: TObject; var Value: String);
begin
    nG_AddCount := dtmMain2.cdsSO114.FieldByName('AddCount').AsInteger;
    Value := TUstr.CommaNumber(IntToStr(nG_AddCount));

    nG_Group_AddCount := nG_Group_AddCount + nG_AddCount;
    nG_Total_AddCount := nG_Total_AddCount + nG_AddCount;

end;

procedure TrptSO8A40A.QRLabel16Print(sender: TObject; var Value: String);
begin
    Value := TUstr.CommaNumber(IntToStr(nG_Total_AddCount));
end;



procedure TrptSO8A40A.QRLabel19Print(sender: TObject; var Value: String);
begin
    //暫時先不Show
    //由拆帳作業之總和資料找出總有效戶數
    //if dtmMain2.cdsSO117.Locate('COMP_ID;BELONGYM', VarArrayOf([sG_CompCode, sG_BelongYM]), []) then
    //  Value := '總有效戶數: ' + TUstr.CommaNumber(dtmMain2.cdsSO117.FieldByName('TOTALVALIDCOUNTS').AsString) + ' 戶';
    Value := '';

end;

procedure TrptSO8A40A.QRLabel20Print(sender: TObject; var Value: String);
begin
    //暫時先不Show
    //由拆帳作業之總和資料找出總有效戶數
    //if dtmMain2.cdsSO117.Locate('COMP_ID;BELONGYM', VarArrayOf([sG_CompCode, sG_BelongYM]), []) then
    //  Value := '總客戶數: ' + TUstr.CommaNumber(dtmMain2.cdsSO117.FieldByName('TOTALCUSTCOUNTS').AsString) + ' 戶';
    Value := '';
end;

procedure TrptSO8A40A.parseChannelCountsString;
var
    L_Obj : TRptChannelViewData;
    sL_ChannelCountsString : String;
    L_ParseChannelCounts,L_ParseChannelDetailData : TStringList;
    ii,jj : Integer;
begin
    if dtmMain2.cdsSO117.Locate('COMP_ID;BELONGYM', VarArrayOf([sG_CompCode, sG_BelongYM]), []) then
    begin
      //該次歸屬日期所有頻道收視戶的數量
      sL_ChannelCountsString := dtmMain2.cdsSO117.FieldByName('CHANNELCOUNTS').AsString;
    end;
//showmessage(sL_ChannelCountsString);

    //拆解成各頻道收視戶數量
    L_ParseChannelCounts := TStringList.Create;
    if sL_ChannelCountsString <> '' then
      L_ParseChannelCounts := TUstr.ParseStrings(sL_ChannelCountsString,';');

    for ii:=0 to L_ParseChannelCounts.Count-1 do
    begin
      L_ParseChannelDetailData := TUstr.ParseStrings(L_ParseChannelCounts[ii],'~');

      for jj:=0 to L_ParseChannelDetailData.Count-1 do
      begin
        L_Obj := TRptChannelViewData.Create;
        L_Obj.sChannel_ID := L_ParseChannelDetailData[0];
        L_Obj.sChannel_Name := L_ParseChannelDetailData[1];
        L_Obj.sChannelTotalViewCounts := L_ParseChannelDetailData[2];

        G_ChannelViewDataStrList.AddObject(L_ParseChannelDetailData[0], L_Obj);
      end;
    end;

end;


procedure TrptSO8A40A.QuickRepBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
begin
    G_ChannelViewDataStrList := TStringList.Create;

    parseChannelCountsString;

    //列印前歸零
    nG_Total_FirstCount := 0;
    nG_Total_LastCount := 0;
    nG_Total_AddCount := 0;
    nG_Total_ReduceCount := 0;
    fG_Total_EMC_Benefit := 0;
    fG_Total_So_Benefit := 0;
    fG_Total_Provider_Benefit := 0;
    fG_Total_EMC_Income := 0;
    fG_Total_InCome := 0;
    fG_Total_OutCome := 0;
end;

function TrptSO8A40A.getChannelViewCounts(sI_ProductID : String) : String;
var
    nL_Ndx : Integer;
begin
    //頻道之總收視戶數
    nL_Ndx := G_ChannelViewDataStrList.IndexOf(sI_ProductID);
//showmessage(sI_ProductID + '***' + IntToStr(nL_Ndx));
    if nL_Ndx<>-1 then
       Result := (G_ChannelViewDataStrList.Objects[nL_Ndx] as TRptChannelViewData).sChannelTotalViewCounts;


end;


procedure TrptSO8A40A.orderByCDSSO114;
begin
    if dtmMain2.cdsSo114 IS TClientDataSet then
    begin
     if TClientDataSet(dtmMain2.cdsSo114).IndexName = 'TmpIndex' then
       TClientDataSet(dtmMain2.cdsSo114).DeleteIndex('TmpIndex');
     //TClientDataSet(dtmMain2.cdsSo114).AddIndex('TmpIndex', Column.FieldName, [ixCaseInsensitive],'','',0);
     TClientDataSet(dtmMain2.cdsSo114).AddIndex('TmpIndex', 'PROVIDERID', [ixCaseInsensitive],'','',0);
     TClientDataSet(dtmMain2.cdsSo114).IndexName := 'TmpIndex';
    end;
end;

procedure TrptSO8A40A.QRLabel22Print(sender: TObject; var Value: String);
var
    sL_ProviderDesc : String;
begin
    sL_ProviderDesc := dtmMain2.cdsSO114.FieldByName('PROVIDERDESC').AsString;
    Value := sL_ProviderDesc;
end;

procedure TrptSO8A40A.lblInComePrint(sender: TObject; var Value: String);
begin
    fG_InCome := dtmMain2.cdsSO114.FieldByName('INCOME').AsFloat;
    Value := TUstr.CommaNumber(FloatToStr(fG_InCome));

    fG_Group_InCome := fG_Group_InCome + fG_InCome;
    fG_Total_InCome := fG_Total_InCome + fG_InCome;
end;

procedure TrptSO8A40A.QRLabel15Print(sender: TObject; var Value: String);
begin
    Value := TUstr.CommaNumber(FloatToStr(fG_Total_InCome));
end;

procedure TrptSO8A40A.lblOutComePrint(sender: TObject; var Value: String);
begin
    fG_OutCome := dtmMain2.cdsSO114.FieldByName('OUTCOME').AsFloat;
    Value := TUstr.CommaNumber(FloatToStr(fG_OutCome));

    fG_Group_OutCome := fG_Group_OutCome + fG_OutCome;
    fG_Total_OutCome := fG_Total_OutCome + fG_OutCome;
end;

procedure TrptSO8A40A.QRLabel23Print(sender: TObject; var Value: String);
begin
    Value := TUstr.CommaNumber(FloatToStr(fG_Total_OutCome));
end;

procedure TrptSO8A40A.QRGroup1BeforePrint(Sender: TQRCustomBand;
  var PrintBand: Boolean);
begin
    if sG_ProviderGroup = 'Y' then    //報表依更應商分群組來分頁
      QRGroup1.ForceNewPage := true
    else
      QRGroup1.ForceNewPage := false;


    //每個 Group 列印前歸零
    nG_Group_FirstCount := 0;
    nG_Group_LastCount := 0;
    nG_Group_AddCount := 0 ;
    nG_Group_ReduceCount :=0;

    fG_Group_InCome := 0;
    fG_Group_OutCome := 0;

    fG_Group_EMC_Income := 0;
    fG_Group_EMC_Benefit := 0;
    fG_Group_So_Benefit := 0;
    fG_Group_Provider_Benefit := 0;

end;

procedure TrptSO8A40A.QRLabel24Print(sender: TObject; var Value: String);
begin
    Value := TUstr.CommaNumber(IntToStr(nG_Group_FirstCount));
end;

procedure TrptSO8A40A.QRLabel25Print(sender: TObject; var Value: String);
begin
    Value := TUstr.CommaNumber(IntToStr(nG_Group_LastCount));
end;

procedure TrptSO8A40A.QRLabel26Print(sender: TObject; var Value: String);
begin
    Value := TUstr.CommaNumber(IntToStr(nG_Group_AddCount));
end;

procedure TrptSO8A40A.QRLabel27Print(sender: TObject; var Value: String);
begin
    Value := TUstr.CommaNumber(FloatToStr(fG_Group_InCome));
end;

procedure TrptSO8A40A.QRLabel28Print(sender: TObject; var Value: String);
begin
    Value := TUstr.CommaNumber(FloatToStr(fG_Group_OutCome));
end;

procedure TrptSO8A40A.QRLabel29Print(sender: TObject; var Value: String);
begin
    //淨額 = 收入 - 退費

    Value := TUstr.CommaNumber(FloatToStr(fG_Group_EMC_Income));
end;

procedure TrptSO8A40A.QRLabel30Print(sender: TObject; var Value: String);
begin
    Value := TUstr.CommaNumber(FloatToStr(fG_Group_EMC_Benefit));
end;

procedure TrptSO8A40A.QRLabel31Print(sender: TObject; var Value: String);
begin
    Value := TUstr.CommaNumber(FloatToStr(fG_Group_So_Benefit));
end;

procedure TrptSO8A40A.QRLabel32Print(sender: TObject; var Value: String);
begin
    Value := TUstr.CommaNumber(FloatToStr(fG_Group_Provider_Benefit));
end;

procedure TrptSO8A40A.lblReduceCountPrint(sender: TObject;
  var Value: String);
begin
    nG_ReduceCount := dtmMain2.cdsSO114.FieldByName('ReduceCount').AsInteger;
    Value := TUstr.CommaNumber(IntToStr(nG_ReduceCount));

    nG_Group_ReduceCount := nG_Group_ReduceCount + nG_ReduceCount;
    nG_Total_ReduceCount := nG_Total_ReduceCount + nG_ReduceCount;
end;

procedure TrptSO8A40A.QRLabel34Print(sender: TObject; var Value: String);
begin
    Value := TUstr.CommaNumber(IntToStr(nG_Group_ReduceCount));
end;

procedure TrptSO8A40A.QRLabel35Print(sender: TObject; var Value: String);
begin
    Value := TUstr.CommaNumber(IntToStr(nG_Total_ReduceCount));
end;

end.
