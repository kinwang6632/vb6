unit frmInvA06_1U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, DB, Grids, DBGrids, StdCtrls, Mask, Buttons, ExtCtrls,
  fraYMDU, cxStyles, dxSkinsCore, dxSkinBlack, dxSkinBlue, dxSkinCaramel,
  dxSkinCoffee, dxSkinsDefaultPainters, dxSkinscxPCPainter, cxCustomData,
  cxGraphics, cxFilter, cxData, cxDataStorage, cxEdit, cxDBData,
  cxGridLevel, cxClasses, cxControls, cxGridCustomView,
  cxGridCustomTableView, cxGridTableView, cxGridDBTableView, cxGrid,
  cxSpinEdit, cxTextEdit, cxContainer, cxCurrencyEdit;


type
  TfrmInvA06_1 = class(TForm)
    Panel1: TPanel;
    Label5: TLabel;
    Splitter1: TSplitter;
    Panel2: TPanel;
    Label3: TLabel;
    Label4: TLabel;
    lblChargeDate: TLabel;
    Label8: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    lblSaleAmt: TLabel;
    lblTaxAmt: TLabel;
    lblTotalAmt: TLabel;
    lblCounts: TLabel;
    btnExit: TBitBtn;
    btnShowAllData: TBitBtn;
    Panel3: TPanel;
    Label11: TLabel;
    Label12: TLabel;
    btnQuery: TBitBtn;
    dgrChargeDetail: TDBGrid;
    btnOpen: TBitBtn;
    dgrChargeSelected: TDBGrid;
    btnClose: TBitBtn;
    rdbCustID: TRadioButton;
    rdbBillID: TRadioButton;
    medCustID: TMaskEdit;
    medBillID: TMaskEdit;
    dsrInv016: TDataSource;
    dsrInv017: TDataSource;
    BtnDelete: TBitBtn;
    rdbChargeDate: TRadioButton;
    medCustID2: TMaskEdit;
    medBillID2: TMaskEdit;
    fraStartChargeDate2: TfraYMD;
    dgrChargeDetail2: TcxGrid;
    gvCharge: TcxGridDBTableView;
    glCharge: TcxGridLevel;
    gvShouldBeAssigned: TcxGridDBColumn;
    gvSeq: TcxGridDBColumn;
    gvCustid: TcxGridDBColumn;
    gvChargeDate: TcxGridDBColumn;
    gvTaxType: TcxGridDBColumn;
    gvTaxRate: TcxGridDBColumn;
    gvSaleAmount: TcxGridDBColumn;
    gvTaxAmount: TcxGridDBColumn;
    gvInvAmount: TcxGridDBColumn;
    gvIsValid: TcxGridDBColumn;
    gvHowToCreate: TcxGridDBColumn;
    edtAmount1: TcxCurrencyEdit;
    fraStartChargeDate: TfraYMD;
    rdbInvAmount: TRadioButton;
    edtAmount2: TcxCurrencyEdit;
    dgrChargeSelected2: TcxGrid;
    gvSelect: TcxGridDBTableView;
    cvShouldBeAssigned: TcxGridDBColumn;
    cvSeq: TcxGridDBColumn;
    cvBILLID: TcxGridDBColumn;
    cvBILLIDITEMNO: TcxGridDBColumn;
    glSelect: TcxGridLevel;
    cvTAXTYPE: TcxGridDBColumn;
    cvCHARGEDATE: TcxGridDBColumn;
    cvITEMID: TcxGridDBColumn;
    cvDESCRIPTION: TcxGridDBColumn;
    cvQUANTITY: TcxGridDBColumn;
    cvUNITPRICE: TcxGridDBColumn;
    cvTAXRATE: TcxGridDBColumn;
    cvTAXAMOUNT: TcxGridDBColumn;
    cvTOTALAMOUNT: TcxGridDBColumn;
    cvSTARTDATE: TcxGridDBColumn;
    cvENDDATE: TcxGridDBColumn;
    cvCHARGEEN: TcxGridDBColumn;
    procedure FormShow(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure rdbCustIDClick(Sender: TObject);
    procedure rdbCustIDEnter(Sender: TObject);
    procedure btnQueryClick(Sender: TObject);
    procedure dsrInv016DataChange(Sender: TObject; Field: TField);
    procedure dgrChargeDetailDblClick(Sender: TObject);
    procedure btnOpenClick(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure btnShowAllDataClick(Sender: TObject);
    procedure dgrChargeSelectedDblClick(Sender: TObject);
    procedure BtnDeleteClick(Sender: TObject);
    procedure medCustIDExit(Sender: TObject);
    procedure medBillIDExit(Sender: TObject);
    procedure fraStartChargeDatemseYMDExit(Sender: TObject);
    procedure gvChargeDblClick(Sender: TObject);
    procedure gvSelectDblClick(Sender: TObject);
  private
    { Private declarations }
    sG_CustID,sG_ChargeDate,sG_BillID,sG_CustID2,sG_ChargeDate2,sG_BillID2,sG_QueryType,sG_InvAmount,sG_InvAmount2 : String;
    procedure setInitialData;
    function isDataOK : Boolean;
  public
    { Public declarations }
    sG_CompID,sG_UserID,sG_DataArea,sG_StartChargeDate,sG_EndChargeDate : String;
    fG_SaleAmount,fG_TaxAmount,fG_InvAmount,fG_Counts : Double;
    sGMinSeq, sGMaxSeq: String;
  end;

var
  frmInvA06_1: TfrmInvA06_1;

implementation

uses cbUtilis, dtmMainU, dtmMainJU, frmMainU, Ustru; 

{$R *.dfm}



procedure TfrmInvA06_1.setInitialData;
begin


    lblChargeDate.Caption := sG_StartChargeDate + ' ~ ' + sG_EndChargeDate;
    lblSaleAmt.Caption := TUstr.replaceStr(TUstr.ChangeMinusTag(fG_SaleAmount),'.00','');
    lblTaxAmt.Caption := TUstr.replaceStr(TUstr.ChangeMinusTag(fG_TaxAmount),'.00','');
    lblTotalAmt.Caption := TUstr.replaceStr(TUstr.ChangeMinusTag(fG_InvAmount),'.00','');
    lblCounts.Caption := TUstr.replaceStr(TUstr.ChangeMinusTag(fG_Counts),'.00','');

end;

procedure TfrmInvA06_1.FormShow(Sender: TObject);
begin
  setInitialData;
  self.Caption := frmMain.GetFormTitleString( 'A06_1', '選擇開立發票' );
end;

procedure TfrmInvA06_1.btnExitClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmInvA06_1.rdbCustIDClick(Sender: TObject);
begin
{
    if rdbCustID.Checked then
    begin
      rdbCustID.Checked := false;
      rdbBillID.Checked := true;
    end
    else
    begin
      rdbCustID.Checked := true;
      rdbBillID.Checked := false;
    end;
}
end;

procedure TfrmInvA06_1.rdbCustIDEnter(Sender: TObject);
begin
    if rdbCustID.Checked then
    begin
      rdbCustID.Checked := false;
      rdbBillID.Checked := true;
      sG_QueryType := QUERY_TYPE_BILLID;
    end
    else
    begin
      rdbCustID.Checked := true;
      rdbBillID.Checked := false;
      sG_QueryType := QUERY_TYPE_CUSTID;
    end;
end;

function TfrmInvA06_1.isDataOK: Boolean;
begin
//sG_CustID,sG_BillID
    {#6262 增加金額查詢 By Kin 2012/06/27}
    if rdbInvAmount.Checked then
    begin
      sG_InvAmount := Trim(VarToStrDef(edtAmount1.EditValue,EmptyStr));
      sG_InvAmount2 := Trim(VarToStrDef(edtAmount2.EditValue,EmptyStr));
      if ( sG_InvAmount = '' ) or ( sG_InvAmount2 = '' ) then
      begin
        MessageDlg('請填入發票金額範圍',mtError,[mbOK],0);
        edtAmount1.SetFocus;
        Result := False;
        Exit;
      end;
    end;

    if rdbCustID.Checked then
    begin
      sG_CustID := Trim(medCustID.Text);
      sG_CustID2 := Trim(medCustID2.Text);
      if (sG_CustID = '') or (sG_CustID2 = '') then
      begin
        MessageDlg('請填入客戶編號範圍',mtError,[mbOk],0);
        medCustID.SetFocus;
        Result := false;
        Exit;
      end;
    end;

    if rdbCustID.Checked then
    begin
      sG_CustID := Trim(medCustID.Text);
      sG_CustID2 := Trim(medCustID2.Text);
      if (sG_CustID = '') or (sG_CustID2 = '') then
      begin
        MessageDlg('請填入客戶編號範圍',mtError,[mbOk],0);
        medCustID.SetFocus;
        Result := false;
        Exit;
      end;
    end;

    if rdbChargeDate.Checked then
    begin
      sG_ChargeDate := Trim(fraStartChargeDate.mseYMD.Text);
      sG_ChargeDate2 := Trim(fraStartChargeDate2.mseYMD.Text);
      if (sG_ChargeDate = '/  /') or (sG_ChargeDate2 = '/  /') then
      begin
        MessageDlg('請填入收費日期範圍',mtError,[mbOk],0);
        fraStartChargeDate.SetFocus;
        Result := false;
        Exit;
      end;
    end;

    if rdbBillID.Checked then
    begin
      sG_BillID := Trim(medBillID.Text);
      sG_BillID2 := Trim(medBillID2.Text);
      if (sG_BillID = '') or (sG_BillID2 = '') then
      begin
        MessageDlg('請填入單據編號範圍',mtError,[mbOk],0);
        medBillID.SetFocus;
        Result := false;
        Exit;
      end;

      if (Length(sG_BillID) <> 15) or (Length(sG_BillID2) <> 15) then
      begin
          MessageDlg('單據編號請輸入15位數',mtError,[mbOk],0);
          medBillID.SetFocus;
          Result := false;
          Exit;
      end;

    end;

    Result := true;
end;

procedure TfrmInvA06_1.btnQueryClick(Sender: TObject);

begin
    if isDataOK then
    begin
//      sGMinSeq := EmptyStr;
//      sGMaxSeq := EmptyStr;
      

      if rdbCustID.Checked then
      begin
        fG_Counts := dtmMainJ.getInv016Data(QUERY_TYPE_CUSTID,sG_CustID,sG_CustID2,sG_CompID,
          sG_StartChargeDate,sG_EndChargeDate,sGMinSeq,sGMaxSeq,
          fG_SaleAmount,fG_TaxAmount,fG_InvAmount);

        if fG_Counts = 0 then
        begin
          WarningMsg('無此客戶資料，請重新輸入查詢條件！');
          medCustID.SetFocus;
          fG_SaleAmount := 0;
          fG_TaxAmount := 0;
          fG_InvAmount := 0;
          fG_Counts := 0;
        end;
      end
      else if rdbInvAmount.Checked then
      begin
        if StrToInt(sG_InvAmount) > StrToInt( sG_InvAmount2 ) then      {#6262 增加金額查詢 By Kin 2012/06/27}
        begin
          WarningMsg('起始金額必須小於等於截止金額');
          edtAmount1.SetFocus;
          fG_SaleAmount := 0;
          fG_TaxAmount := 0;
          fG_InvAmount := 0;
          fG_Counts := 0;
        end else
        begin
          fG_Counts := dtmMainJ.getInv016Data(QUERY_TYPE_INVAMOUNT,sG_InvAmount,sG_InvAmount2,sG_CompID,
            sG_StartChargeDate,sG_EndChargeDate,sGMinSeq,sGMaxSeq,
            fG_SaleAmount,fG_TaxAmount,fG_InvAmount);
          if fG_Counts = 0 then
          begin
            WarningMsg('無此發票金額資料，請重新輸入查詢條件！');
            edtAmount1.SetFocus;
            fG_SaleAmount := 0;
            fG_TaxAmount := 0;
            fG_InvAmount := 0;
            fG_Counts := 0;
          end;
        end;
      end
      else if rdbChargeDate.Checked then
      begin
        if sG_ChargeDate2 < sG_ChargeDate then
        begin
          WarningMsg('截止日期必須大於等於開始日期');
          fraStartChargeDate.SetFocus;
          fG_SaleAmount := 0;
          fG_TaxAmount := 0;
          fG_InvAmount := 0;
          fG_Counts := 0;
          fraStartChargeDate.SetFocus;
        end
        else
          begin
           fG_Counts := dtmMainJ.getInv016Data(QUERY_TYPE_CHARGEDATE,sG_ChargeDate,sG_ChargeDate2,sG_CompID,sG_StartChargeDate,sG_EndChargeDate,sGMinSeq,sGMaxSeq, fG_SaleAmount,fG_TaxAmount,fG_InvAmount);

           if fG_Counts = 0 then
              begin
                WarningMsg('無此收費日期資料，請重新輸入查詢條件！');
                fraStartChargeDate.SetFocus;
                fG_SaleAmount := 0;
                fG_TaxAmount := 0;
                fG_InvAmount := 0;
                fG_Counts := 0;
              end;
          end;
      end
      else if rdbBillID.Checked then
      begin
        fG_Counts := dtmMainJ.getInv016Data(QUERY_TYPE_BILLID,sG_BillID,sG_BillID2,sG_CompID,sG_StartChargeDate,sG_EndChargeDate,sGMinSeq,sGMaxSeq, fG_SaleAmount,fG_TaxAmount,fG_InvAmount);

        if fG_Counts = 0 then
        begin
          WarningMsg('無此單據編號資料，請重新輸入查詢條件！');
          medBillID.SetFocus;
          fG_SaleAmount := 0;
          fG_TaxAmount := 0;
          fG_InvAmount := 0;
          fG_Counts := 0;
        end;
      end;

      setInitialData;
    end;
end;

procedure TfrmInvA06_1.dsrInv016DataChange(Sender: TObject;
  Field: TField);
var
    sL_SEQ : String;
begin
    sL_SEQ := dsrInv016.DataSet.FieldByName('SEQ').AsString;
    dtmMainJ.getInv017Data(sL_SEQ);
end;

procedure TfrmInvA06_1.dgrChargeDetailDblClick(Sender: TObject);
var
    aRecNo16 : Integer;
    sL_SEQ,sL_ShouldBeAssigned : String;
begin
    aRecNo16 := dsrInv016.DataSet.RecNo;
    sL_SEQ := dsrInv016.DataSet.FieldByName('SEQ').AsString;
    sL_ShouldBeAssigned := dsrInv016.DataSet.FieldByName('ShouldBeAssigned').AsString;

    if sL_ShouldBeAssigned = 'Y' then
      sL_ShouldBeAssigned := 'N'
    else
      sL_ShouldBeAssigned := 'Y';

    //改變是否要開立發票
    dtmMainJ.updateInv016ShouldBeAssigned(sL_SEQ,sL_ShouldBeAssigned,sG_CompID,sG_StartChargeDate,sG_EndChargeDate,
      True, sGMinSeq, sGMaxSeq );


    //若點選的查詢條件是空值,則Show出符合前一畫面的資料
    sG_CustID := Trim(medCustID.Text);
    sG_BillID := Trim(medBillID.Text);
    {
    if (rdbCustID.Checked) and (sG_CustID <> '') and (sG_CustID2 <> '') then
      dtmMainJ.getInv016Data(QUERY_TYPE_CUSTID,sG_CustID,sG_CustID2,sG_CompID,sG_StartChargeDate,sG_EndChargeDate,sGMinSeq,sGMaxSeq,fG_SaleAmount,fG_TaxAmount,fG_InvAmount)
    else if (rdbChargeDate.Checked) and (sG_ChargeDate <> '') and (sG_ChargeDate <> '') then
      dtmMainJ.getInv016Data(QUERY_TYPE_CHARGEDATE,sG_ChargeDate,sG_ChargeDate2,sG_CompID,sG_StartChargeDate,sG_EndChargeDate,sGMinSeq,sGMaxSeq,fG_SaleAmount,fG_TaxAmount,fG_InvAmount)
    else if (rdbBillID.Checked) and (sG_BillID <> '') and (sG_BillID2 <> '') then
      dtmMainJ.getInv016Data(QUERY_TYPE_BILLID,sG_BillID,sG_BillID2,sG_CompID,sG_StartChargeDate,sG_EndChargeDate,sGMinSeq,sGMaxSeq,fG_SaleAmount,fG_TaxAmount,fG_InvAmount)
    else
      dtmMainJ.getInv016Data(QUERY_TYPE_ALL,'','',sG_CompID,sG_StartChargeDate,sG_EndChargeDate,sGMinSeq,sGMaxSeq,fG_SaleAmount,fG_TaxAmount,fG_InvAmount);
}
    dsrInv016.DataSet.RecNo := aRecNo16;

    setInitialData;

end;

procedure TfrmInvA06_1.btnOpenClick(Sender: TObject);
var
    sL_SEQ,sL_ShouldBeAssigned : String;
begin
    //改變是否要開立發票
    Screen.Cursor := crHourGlass;
    sL_SEQ := '';
    sL_ShouldBeAssigned := 'Y';
    dtmMainJ.updateInv016ShouldBeAssigned( sL_SEQ, sL_ShouldBeAssigned, sG_CompID,
      sG_StartChargeDate, sG_EndChargeDate, True, sGMinSeq, sGMaxSeq );

    dtmMainJ.getInv016Data(QUERY_TYPE_ALL,'','', sG_CompID,
      sG_StartChargeDate, sG_EndChargeDate, sGMinSeq, sGMaxSeq,
      fG_SaleAmount, fG_TaxAmount, fG_InvAmount );

    setInitialData;

    Screen.Cursor := crDefault;
end;

procedure TfrmInvA06_1.btnCloseClick(Sender: TObject);
var
    sL_SEQ,sL_ShouldBeAssigned : String;
begin
    //改變是否要開立發票
    Screen.Cursor := crHourGlass;
    sL_SEQ := '';
    sL_ShouldBeAssigned := 'N';
    dtmMainJ.updateInv016ShouldBeAssigned( sL_SEQ, sL_ShouldBeAssigned, sG_CompID,
      sG_StartChargeDate, sG_EndChargeDate, True, sGMinSeq, sGMaxSeq );

    dtmMainJ.getInv016Data(QUERY_TYPE_ALL, '', '', sG_CompID,
      sG_StartChargeDate, sG_EndChargeDate, sGMinSeq, sGMaxSeq,
      fG_SaleAmount,fG_TaxAmount,fG_InvAmount );

    setInitialData;

    Screen.Cursor := crDefault;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA06_1.btnShowAllDataClick(Sender: TObject);
begin
    medCustID.Text := EmptyStr;
    medBillID.Text := EmptyStr;

    sGMinSeq := EmptyStr;
    sGMaxSeq := EmptyStr;

    fG_Counts := dtmMainJ.getInv016Data( QUERY_TYPE_ALL,  EmptyStr, EmptyStr,
      sG_CompID, sG_StartChargeDate,sG_EndChargeDate, sGMinSeq, sGMaxSeq,
      fG_SaleAmount, fG_TaxAmount, fG_InvAmount);

    setInitialData;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA06_1.dgrChargeSelectedDblClick(Sender: TObject);
var
    aRecNo16, aRecNo17 : Integer;
    fL_TotalAmount,fL_TaxAmount,fL_UnitPrice: Double;
    sL_SEQ,sL_ShouldBeAssigned, sL_BillId, sL_BillIdItemNo : String;
begin
    sL_SEQ := dsrInv017.DataSet.FieldByName('SEQ').AsString;
    sL_ShouldBeAssigned := dsrInv017.DataSet.FieldByName('SHOULDBEASSIGNED').AsString;
    sL_BillId := dsrInv017.DataSet.FieldByName('BILLID').AsString;
    sL_BillIdItemNo := dsrInv017.DataSet.FieldByName('BILLIDITEMNO').AsString;

    if sL_ShouldBeAssigned = 'Y' then
      sL_ShouldBeAssigned := 'N'
    else
      sL_ShouldBeAssigned := 'Y';

    //改變是否要開立發票
    dtmMainJ.updateInv017ShouldBeAssigned(sL_SEQ,sL_ShouldBeAssigned,sL_BillId,sL_BillIdItemNo);

    //需重算主檔的金額
    aRecNo16:= dsrInv016.DataSet.RecNo;
    aRecNo17:= dsrInv017.DataSet.RecNo;

    fL_TotalAmount := dsrInv017.DataSet.FieldByName('TotalAmount').AsFloat;
    fL_TaxAmount := dsrInv017.DataSet.FieldByName('TaxAmount').AsFloat;
    fL_UnitPrice := dsrInv017.DataSet.FieldByName('UnitPrice').AsFloat;
    if sL_ShouldBeAssigned = 'N' then
      begin
        fL_TotalAmount := -fL_TotalAmount;
        fL_TaxAmount := -fL_TaxAmount;
        fL_UnitPrice := -fL_UnitPrice;
      end;
    fL_TotalAmount := dsrInv016.DataSet.FieldByName('InvAmount').AsFloat + fL_TotalAmount;
    fL_TaxAmount := dsrInv016.DataSet.FieldByName('TaxAmount').AsFloat + fL_TaxAmount;
    fL_UnitPrice := dsrInv016.DataSet.FieldByName('SaleAmount').AsFloat + fL_UnitPrice;
    //dsrInv016.DataSet.FieldByName('InvAmount').AsFloat := fL_TotalAmount;
    dtmMainJ.updateInv016InvAmount(fL_UnitPrice,fL_TaxAmount,fL_TotalAmount,sL_SEQ,sG_CompID,sG_StartChargeDate,sG_EndChargeDate);

    //如果主檔的金額是 0，則設為不開立
    if fL_TotalAmount > 0 then
      sL_ShouldBeAssigned := 'Y'
    else
      sL_ShouldBeAssigned := 'N';
      
    dtmMainJ.updateInv016ShouldBeAssigned( sL_SEQ, sL_ShouldBeAssigned, sG_CompID,
      sG_StartChargeDate, sG_EndChargeDate, False, sGMinSeq, sGMaxSeq );

    //重讀 Inv016
    sG_CustID := Trim(medCustID.Text);
    sG_BillID := Trim(medBillID.Text);
    if (rdbCustID.Checked) and (sG_CustID <> '') and (sG_CustID2 <> '') then
      dtmMainJ.getInv016Data(QUERY_TYPE_CUSTID,sG_CustID,sG_CustID2,sG_CompID,sG_StartChargeDate,sG_EndChargeDate, sGMinSeq,sGMaxSeq , fG_SaleAmount,fG_TaxAmount,fG_InvAmount)
    else if (rdbChargeDate.Checked) and (sG_ChargeDate <> '') and (sG_ChargeDate <> '') then
      dtmMainJ.getInv016Data(QUERY_TYPE_CHARGEDATE,sG_ChargeDate,sG_ChargeDate2,sG_CompID,sG_StartChargeDate,sG_EndChargeDate,sGMinSeq,sGMaxSeq,  fG_SaleAmount,fG_TaxAmount,fG_InvAmount)
    else if (rdbBillID.Checked) and (sG_BillID <> '') and (sG_BillID2 <> '') then
      dtmMainJ.getInv016Data(QUERY_TYPE_BILLID,sG_BillID,sG_BillID2,sG_CompID,sG_StartChargeDate,sG_EndChargeDate,sGMinSeq,sGMaxSeq,  fG_SaleAmount,fG_TaxAmount,fG_InvAmount)
    else if (rdbInvAmount.Checked ) and (sG_InvAmount <> '') and (sG_InvAmount2 <> '') then
      dtmMainJ.getInv016Data(QUERY_TYPE_INVAMOUNT,sG_InvAmount,sG_InvAmount2,sG_CompID,sG_StartChargeDate,sG_EndChargeDate,sGMinSeq,sGMaxSeq,  fG_SaleAmount,fG_TaxAmount,fG_InvAmount)
    else
      dtmMainJ.getInv016Data(QUERY_TYPE_ALL,'','',sG_CompID,sG_StartChargeDate,sG_EndChargeDate,sGMinSeq,sGMaxSeq, fG_SaleAmount,fG_TaxAmount,fG_InvAmount);
    dsrInv016.DataSet.RecNo := aRecNo16;

    ////重讀 Inv017
    sL_SEQ := dsrInv016.DataSet.FieldByName('SEQ').AsString;
    dtmMainJ.getInv017Data(sL_SEQ);
    dsrInv017.DataSet.RecNo := aRecNo17;

    setInitialData;
end;

procedure TfrmInvA06_1.BtnDeleteClick(Sender: TObject);
var
    aRecNo16:Integer;
    sL_SEQ : String;
begin
    //改變是否要開立發票
    aRecNo16:= dsrInv016.DataSet.RecNo;
    sL_SEQ := dsrInv016.DataSet.FieldByName('SEQ').AsString;
    dtmMainJ.updateInv016StopFlag(sL_SEQ,sG_CompID,sG_StartChargeDate,sG_EndChargeDate);

    //dtmMainJ.getInv016Data(QUERY_TYPE_ALL,'','',sG_CompID,sG_StartChargeDate,sG_EndChargeDate,sGMinSeq,sGMaxSeq,fG_SaleAmount,fG_TaxAmount,fG_InvAmount);

    fG_Counts := dtmMainJ.getInv016Data(QUERY_TYPE_ALL,'','',sG_CompID,sG_StartChargeDate,sG_EndChargeDate,sGMinSeq,sGMaxSeq, fG_SaleAmount,fG_TaxAmount,fG_InvAmount);
    setInitialData;

    if (aRecNo16>dsrInv016.DataSet.RecordCount) and (aRecNo16>0) then
      aRecNo16 := aRecNo16-1;
    if dsrInv016.DataSet.RecordCount>0 then
      dsrInv016.DataSet.RecNo := aRecNo16;

    Screen.Cursor := crDefault;
end;

procedure TfrmInvA06_1.medCustIDExit(Sender: TObject);
begin
  medCustID2.Text := medCustID.Text;
end;

procedure TfrmInvA06_1.medBillIDExit(Sender: TObject);
begin
  medBillID2.Text := medBillID.Text;
end;

procedure TfrmInvA06_1.fraStartChargeDatemseYMDExit(Sender: TObject);
begin
  fraStartChargeDate.mseYMDExit(Sender);
  fraStartChargeDate2.mseYMD.Text := fraStartChargeDate.mseYMD.Text;
end;

procedure TfrmInvA06_1.gvChargeDblClick(Sender: TObject);
var
    aRecNo16 : Integer;
    sL_SEQ,sL_ShouldBeAssigned : String;
begin
    aRecNo16 := dsrInv016.DataSet.RecNo;
    sL_SEQ := dsrInv016.DataSet.FieldByName('SEQ').AsString;
    sL_ShouldBeAssigned := dsrInv016.DataSet.FieldByName('ShouldBeAssigned').AsString;

    if sL_ShouldBeAssigned = 'Y' then
      sL_ShouldBeAssigned := 'N'
    else
      sL_ShouldBeAssigned := 'Y';

    //改變是否要開立發票
    dtmMainJ.updateInv016ShouldBeAssigned(sL_SEQ,sL_ShouldBeAssigned,sG_CompID,sG_StartChargeDate,sG_EndChargeDate,
      True, sGMinSeq, sGMaxSeq );


    //若點選的查詢條件是空值,則Show出符合前一畫面的資料
    sG_CustID := Trim(medCustID.Text);
    sG_BillID := Trim(medBillID.Text);

    


    if (rdbCustID.Checked) and (sG_CustID <> '') and (sG_CustID2 <> '') then
      dtmMainJ.getInv016Data(QUERY_TYPE_CUSTID,sG_CustID,sG_CustID2,sG_CompID,sG_StartChargeDate,sG_EndChargeDate,sGMinSeq,sGMaxSeq, fG_SaleAmount,fG_TaxAmount,fG_InvAmount)
    else if (rdbChargeDate.Checked) and (sG_ChargeDate <> '') and (sG_ChargeDate <> '') then
      dtmMainJ.getInv016Data(QUERY_TYPE_CHARGEDATE,sG_ChargeDate,sG_ChargeDate2,sG_CompID,sG_StartChargeDate,sG_EndChargeDate,sGMinSeq,sGMaxSeq, fG_SaleAmount,fG_TaxAmount,fG_InvAmount)
    else if (rdbBillID.Checked) and (sG_BillID <> '') and (sG_BillID2 <> '') then
      dtmMainJ.getInv016Data(QUERY_TYPE_BILLID,sG_BillID,sG_BillID2,sG_CompID,sG_StartChargeDate,sG_EndChargeDate,sGMinSeq,sGMaxSeq, fG_SaleAmount,fG_TaxAmount,fG_InvAmount)
    else if (rdbInvAmount.Checked) and (sG_InvAmount <> '' ) and (sG_InvAmount2 <> '') then
      dtmMainJ.getInv016Data(QUERY_TYPE_INVAMOUNT, sG_InvAmount,sG_InvAmount2,sG_CompID,sG_StartChargeDate,sG_EndChargeDate,sGMinSeq,sGMaxSeq, fG_SaleAmount,fG_TaxAmount,fG_InvAmount)
    else
      dtmMainJ.getInv016Data(QUERY_TYPE_ALL,'','',sG_CompID,sG_StartChargeDate,sG_EndChargeDate,sGMinSeq,sGMaxSeq, fG_SaleAmount,fG_TaxAmount,fG_InvAmount);
    if dsrInv016.DataSet.RecordCount < aRecNo16 then
      aRecNo16 := dsrInv016.DataSet.RecordCount;
    dsrInv016.DataSet.RecNo := aRecNo16;

    setInitialData;
end;

procedure TfrmInvA06_1.gvSelectDblClick(Sender: TObject);
  var
    aRecNo16, aRecNo17 : Integer;
    fL_TotalAmount,fL_TaxAmount,fL_UnitPrice: Double;
    sL_SEQ,sL_ShouldBeAssigned, sL_BillId, sL_BillIdItemNo : String;
begin
  sL_SEQ := dsrInv017.DataSet.FieldByName('SEQ').AsString;
    sL_ShouldBeAssigned := dsrInv017.DataSet.FieldByName('SHOULDBEASSIGNED').AsString;
    sL_BillId := dsrInv017.DataSet.FieldByName('BILLID').AsString;
    sL_BillIdItemNo := dsrInv017.DataSet.FieldByName('BILLIDITEMNO').AsString;

    if sL_ShouldBeAssigned = 'Y' then
      sL_ShouldBeAssigned := 'N'
    else
      sL_ShouldBeAssigned := 'Y';

    //改變是否要開立發票
    dtmMainJ.updateInv017ShouldBeAssigned(sL_SEQ,sL_ShouldBeAssigned,sL_BillId,sL_BillIdItemNo);

    //需重算主檔的金額
    aRecNo16:= dsrInv016.DataSet.RecNo;
    aRecNo17:= dsrInv017.DataSet.RecNo;

    fL_TotalAmount := dsrInv017.DataSet.FieldByName('TotalAmount').AsFloat;
    fL_TaxAmount := dsrInv017.DataSet.FieldByName('TaxAmount').AsFloat;
    fL_UnitPrice := dsrInv017.DataSet.FieldByName('UnitPrice').AsFloat;
    if sL_ShouldBeAssigned = 'N' then
      begin
        fL_TotalAmount := -fL_TotalAmount;
        fL_TaxAmount := -fL_TaxAmount;
        fL_UnitPrice := -fL_UnitPrice;
      end;
    fL_TotalAmount := dsrInv016.DataSet.FieldByName('InvAmount').AsFloat + fL_TotalAmount;
    fL_TaxAmount := dsrInv016.DataSet.FieldByName('TaxAmount').AsFloat + fL_TaxAmount;
    fL_UnitPrice := dsrInv016.DataSet.FieldByName('SaleAmount').AsFloat + fL_UnitPrice;
    //dsrInv016.DataSet.FieldByName('InvAmount').AsFloat := fL_TotalAmount;
    dtmMainJ.updateInv016InvAmount(fL_UnitPrice,fL_TaxAmount,fL_TotalAmount,sL_SEQ,sG_CompID,sG_StartChargeDate,sG_EndChargeDate);

    //如果主檔的金額是 0，則設為不開立
    if fL_TotalAmount > 0 then
      sL_ShouldBeAssigned := 'Y'
    else
      sL_ShouldBeAssigned := 'N';
      
    dtmMainJ.updateInv016ShouldBeAssigned( sL_SEQ, sL_ShouldBeAssigned, sG_CompID,
      sG_StartChargeDate, sG_EndChargeDate, False, sGMinSeq, sGMaxSeq );

    //重讀 Inv016
    sG_CustID := Trim(medCustID.Text);
    sG_BillID := Trim(medBillID.Text);
    if (rdbCustID.Checked) and (sG_CustID <> '') and (sG_CustID2 <> '') then
      dtmMainJ.getInv016Data(QUERY_TYPE_CUSTID,sG_CustID,sG_CustID2,sG_CompID,sG_StartChargeDate,sG_EndChargeDate, sGMinSeq,sGMaxSeq , fG_SaleAmount,fG_TaxAmount,fG_InvAmount)
    else if (rdbChargeDate.Checked) and (sG_ChargeDate <> '') and (sG_ChargeDate <> '') then
      dtmMainJ.getInv016Data(QUERY_TYPE_CHARGEDATE,sG_ChargeDate,sG_ChargeDate2,sG_CompID,sG_StartChargeDate,sG_EndChargeDate,sGMinSeq,sGMaxSeq,  fG_SaleAmount,fG_TaxAmount,fG_InvAmount)
    else if (rdbBillID.Checked) and (sG_BillID <> '') and (sG_BillID2 <> '') then
      dtmMainJ.getInv016Data(QUERY_TYPE_BILLID,sG_BillID,sG_BillID2,sG_CompID,sG_StartChargeDate,sG_EndChargeDate,sGMinSeq,sGMaxSeq,  fG_SaleAmount,fG_TaxAmount,fG_InvAmount)
    else if (rdbInvAmount.Checked ) and (sG_InvAmount <> '') and (sG_InvAmount2 <> '') then
      dtmMainJ.getInv016Data(QUERY_TYPE_INVAMOUNT,sG_InvAmount,sG_InvAmount2,sG_CompID,sG_StartChargeDate,sG_EndChargeDate,sGMinSeq,sGMaxSeq,  fG_SaleAmount,fG_TaxAmount,fG_InvAmount)
    else
      dtmMainJ.getInv016Data(QUERY_TYPE_ALL,'','',sG_CompID,sG_StartChargeDate,sG_EndChargeDate,sGMinSeq,sGMaxSeq, fG_SaleAmount,fG_TaxAmount,fG_InvAmount);
    dsrInv016.DataSet.RecNo := aRecNo16;

    ////重讀 Inv017
    sL_SEQ := dsrInv016.DataSet.FieldByName('SEQ').AsString;
    dtmMainJ.getInv017Data(sL_SEQ);
    dsrInv017.DataSet.RecNo := aRecNo17;

    setInitialData;
end;

end.



