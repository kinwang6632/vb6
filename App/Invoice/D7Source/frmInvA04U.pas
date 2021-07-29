unit frmInvA04U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Buttons, Grids, DBGrids, DB, Mask,
  Menus, Provider, DBClient,
  cxStyles, cxCustomData, cxGraphics, cxFilter, cxData,
  cxDataStorage, cxEdit, cxDBData, cxGridLevel, cxClasses, cxControls,
  cxGridCustomView, cxGridCustomTableView, cxGridTableView,
  cxGridDBTableView, cxGrid, cxContainer, cxTextEdit, cxCheckBox,
  cxCurrencyEdit, cxLabel, dxSkinsCore,
  dxSkinsDefaultPainters, dxSkinscxPCPainter, dxSkinBlack, dxSkinBlue,
  dxSkinCaramel, dxSkinCoffee, cbDataLookup, cxMaskEdit, cxDropDownEdit,
  cxCalendar, ADODB;

type
  TfrmInvA04 = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    Label1: TLabel;
    btnQueryInvInfo: TBitBtn;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    lblInvDate: TcxLabel;
    lblInvTitle: TcxLabel;
    lblInvFormat: TcxLabel;
    lblTaxType: TcxLabel;
    lblTaxRate: TcxLabel;
    dsrV_ChargeInfo: TDataSource;
    Panel3: TPanel;
    btnQueryCust: TBitBtn;
    Label7: TLabel;
    txtCustId: TcxTextEdit;
    Panel5: TPanel;
    cxGrid1: TcxGrid;
    cxGrid1DBTableView1: TcxGridDBTableView;
    cxGrid1DBTableView1SOURCE: TcxGridDBColumn;
    cxGrid1DBTableView1COMPCODE: TcxGridDBColumn;
    cxGrid1DBTableView1CUSTID: TcxGridDBColumn;
    cxGrid1DBTableView1BILLNO: TcxGridDBColumn;
    cxGrid1DBTableView1ITEM: TcxGridDBColumn;
    cxGrid1DBTableView1CITEMCODE: TcxGridDBColumn;
    cxGrid1DBTableView1CITEMNAME: TcxGridDBColumn;
    cxGrid1DBTableView1SHOULDDATE: TcxGridDBColumn;
    cxGrid1DBTableView1SHOULDAMT: TcxGridDBColumn;
    cxGrid1DBTableView1REALDATE: TcxGridDBColumn;
    cxGrid1DBTableView1REALAMT: TcxGridDBColumn;
    cxGrid1DBTableView1REALPERIOD: TcxGridDBColumn;
    cxGrid1DBTableView1REALSTARTDATE: TcxGridDBColumn;
    cxGrid1DBTableView1REALSTOPDATE: TcxGridDBColumn;
    cxGrid1DBTableView1CLCTNAME: TcxGridDBColumn;
    cxGrid1DBTableView1STNAME: TcxGridDBColumn;
    cxGrid1DBTableView1ACCOUNTNO: TcxGridDBColumn;
    cxGrid1DBTableView1SERVICETYPE: TcxGridDBColumn;
    cxGrid1Level1: TcxGridLevel;
    cxGrid1DBTableView1SELECTED: TcxGridDBColumn;
    txtInvId: TcxTextEdit;
    lblCustName: TcxLabel;
    cxGrid1DBTableView1RATE1: TcxGridDBColumn;
    cxGrid1DBTableView1TAXCODE: TcxGridDBColumn;
    Bevel1: TBevel;
    cxGrid1DBTableView1SOURCENAME: TcxGridDBColumn;
    Panel4: TPanel;
    Panel6: TPanel;
    cxGrid2: TcxGrid;
    cxGridDBTableView1: TcxGridDBTableView;
    cxGridDBColumn1: TcxGridDBColumn;
    cxGridDBColumn2: TcxGridDBColumn;
    cxGridDBColumn3: TcxGridDBColumn;
    cxGridDBColumn4: TcxGridDBColumn;
    cxGridDBColumn5: TcxGridDBColumn;
    cxGridDBColumn6: TcxGridDBColumn;
    cxGridDBColumn7: TcxGridDBColumn;
    cxGridDBColumn8: TcxGridDBColumn;
    cxGridDBColumn9: TcxGridDBColumn;
    cxGridDBColumn10: TcxGridDBColumn;
    cxGridDBColumn11: TcxGridDBColumn;
    cxGridDBColumn12: TcxGridDBColumn;
    cxGridDBColumn13: TcxGridDBColumn;
    cxGridDBColumn14: TcxGridDBColumn;
    cxGridDBColumn15: TcxGridDBColumn;
    cxGridDBColumn16: TcxGridDBColumn;
    cxGridDBColumn17: TcxGridDBColumn;
    cxGridDBColumn18: TcxGridDBColumn;
    cxGridDBColumn19: TcxGridDBColumn;
    cxGridDBColumn20: TcxGridDBColumn;
    cxGridDBColumn21: TcxGridDBColumn;
    cxGridDBColumn22: TcxGridDBColumn;
    cxGridLevel1: TcxGridLevel;
    btnUnSelect: TBitBtn;
    btnSelect: TBitBtn;
    Label8: TLabel;
    Label9: TLabel;
    Panel7: TPanel;
    btnUpdate: TBitBtn;
    Label10: TLabel;
    btnExit: TBitBtn;
    Label11: TLabel;
    cdsCharge: TClientDataSet;
    dsrCharge: TDataSource;
    btnCancel: TBitBtn;
    cxGrid1DBTableView1CUSTNAME: TcxGridDBColumn;
    cxGridDBColumn23: TcxGridDBColumn;
    cdsV_ChargeInfo: TClientDataSet;
    Label12: TLabel;
    txtMduId: TDataLookup;
    txtShouldDate1: TcxDateEdit;
    txtShouldDate2: TcxDateEdit;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    txtRealDate1: TcxDateEdit;
    Label16: TLabel;
    txtRealDate2: TcxDateEdit;
    btnSelAll: TBitBtn;
    btnImportExcel: TBitBtn;
    dlgOpen1: TOpenDialog;
    procedure btnQueryInvInfoClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnQueryCustClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure txtInvIdKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure txtCustIDKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure txtCustIdExit(Sender: TObject);
    procedure btnUpdateClick(Sender: TObject);
    procedure txtCustIdPropertiesChange(Sender: TObject);
    procedure btnSelectClick(Sender: TObject);
    procedure btnUnSelectClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure txtMduIdCodeNamePropertiesInitPopup(Sender: TObject);
    procedure btnSelAllClick(Sender: TObject);
    procedure cxGrid1DBTableView1SELECTEDCustomDrawFooterCell(
      Sender: TcxGridTableView; ACanvas: TcxCanvas;
      AViewInfo: TcxGridColumnHeaderViewInfo; var ADone: Boolean);
    procedure btnImportExcelClick(Sender: TObject);
  private
    { Private declarations }
    FMasterDataSet: TClientDataSet;
    FSourceDataSet: TClientDataSet;
    FInvDate : TDate;
    FInvId, sG_TaxType, sG_HowToCreate: String;
    dG_TaxRate : Double;
    adoImportExcel : TADOQuery;
    procedure HaveDataStatus;
    procedure NoDataStatus;
    procedure SetDBGridColumn;
    procedure PrepareDataSet;
    procedure CheckIfAlreadySelect;
    function ChecckBeforeUpdate: Boolean;
    function GetChargeInfo(const aCustId, aTaxRate, aTaxType,
      aServiceType, aMduId, aShouldDate1, aShouldDate2,
      aRealDate1, aRealDate2: String): Integer;
    function GetExcelChargeInfo(const aTaxRate, aTaxType,
      aServiceType : String;var aErrCount : Integer; var aErrMsg : string ) :Integer;
    procedure PrepareChargeDataSet;
    function GetDate(const aDate : String):String;
  public
    { Public declarations }
  end;

var
  frmInvA04: TfrmInvA04;


implementation

uses cbUtilis, dtmSOU, dtmMainU, frmMainU, frmInvA04_2U;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA04.FormShow(Sender: TObject);
begin
  self.Caption := frmMain.GetFormTitleString( 'A04', '多客戶發票開立' );
  FMasterDataSet := cdsCharge;
  FSourceDataSet := cdsV_ChargeInfo;
  PrepareDataSet;
  NoDataStatus;
  SetDBGridColumn;

end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA04.btnExitClick(Sender: TObject);
begin
  if not VarIsNull( FMasterDataSet.Data ) then
  begin
    FMasterDataSet.EmptyDataSet;
    FMasterDataSet.Data := Null;
  end;
  PrepareChargeDataSet;
  Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA04.btnQueryInvInfoClick(Sender: TObject);
var
  sL_InvDate,sL_InvTitle, sL_InvFormat,sL_TaxType : String;
  aCount: Integer;
begin
  FInvId := Trim( TxtInvId.Text );
  if ( FInvId = '' ) then
  begin
    WarningMsg('請輸入發票號碼');
    TxtInvId.SetFocus;
    Exit;
  end;

  dtmMain.adoComm.Close;
  dtmMain.adoComm.SQL.Text :=
    '  SELECT INVDATE,         ' +
    '         INVTITLE,        ' +
    '         INVFORMAT,       ' +
    '         TAXTYPE,         ' +
    '         TAXRATE,         ' +
    '         HOWTOCREATE      ' +
    '    FROM INV007           ' +
    '   WHERE INVID =          ' + QuotedStr( FInvId );

    dtmMain.adoComm.Open;
    aCount := dtmMain.adoComm.RecordCount;
    if ( aCount <= 0 ) then
    begin
      dtmMain.adoComm.Close;
      WarningMsg( '無此發票號碼,請重新查詢。' );
      NoDataStatus;
      Exit;
    end;

    FInvDate := dtmMain.adoComm.FieldByName('INVDATE').AsDateTime;
    sL_InvDate := dtmMain.adoComm.FieldByName('INVDATE').AsString;
    sL_InvTitle := dtmMain.adoComm.FieldByName('INVTITLE').AsString;
    sL_InvFormat := dtmMain.adoComm.FieldByName('INVFORMAT').AsString;
    sG_TaxType := dtmMain.adoComm.FieldByName('TAXTYPE').AsString;
    dG_TaxRate := dtmMain.adoComm.FieldByName('TAXRATE').AsFloat;
    sG_HowToCreate := dtmMain.adoComm.FieldByName('HOWTOCREATE').AsString;

    lblInvDate.Caption := sL_InvDate;
    lblInvTitle.Caption := sL_InvTitle;

    if sL_InvFormat = '1' then
      sL_InvFormat := '電子'
    else if sL_InvFormat = '2' then
      sL_InvFormat := '手二'
    else if sL_InvFormat = '3' then
      sL_InvFormat := '手三';
    lblInvFormat.Caption := sL_InvFormat;

    if sG_TaxType = '1' then
      sL_TaxType := '應稅'
    else if sG_TaxType = '2' then
      sL_TaxType := '零稅率'
    else if sG_TaxType = '3' then
      sL_TaxType := '免稅';
    lblTaxType.Caption := sL_TaxType;

    lblTaxRate.Caption := FloatToStr(dG_TaxRate);
    
    dtmMain.adoComm.Close;
    HaveDataStatus;
    txtCustId.SetFocus;

end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA04.btnQueryCustClick(Sender: TObject);
var
  aCustId, aServiceType: String;
  aCount: Integer;
begin
  {#6190 把客編條件移除 By Kin 2011/12/20 For Jacy 口述 }
  {
  aCustID := Trim( txtCustId.Text);
  if ( aCustid = '' ) then
  begin
    WarningMsg( '請輸入客戶編號' );
    txtCustId.SetFocus;
    Exit;
  end;
  }
  {#6190 只要有輸入值就好 By Kin 2011/12/20}
  if (Trim(txtCustId.Text)=EmptyStr) and
      (Trim(txtShouldDate1.Text)=EmptyStr) and
      (Trim(txtShouldDate2.Text)=EmptyStr ) and
      (txtRealDate1.Text=EmptyStr) and
      (txtRealDate2.Text=EmptyStr) and
      (txtMduId.Value=EmptyStr) then
  begin
    WarningMsg( '請輸入任一條' );
    txtCustId.SetFocus;
    Exit;
  end;

  aServiceType := dtmMain.GetServiceTypeStrEx;

  Screen.Cursor := crSQLWait;
  try
    aCustId := txtCustId.Text;
    FSourceDataSet.DisableControls;
    try
      aCount := GetChargeInfo( aCustId, FloatToStr( dG_TaxRate ),
       sG_TaxType, aServiceType,txtMduId.Value,
       txtShouldDate1.Text,txtShouldDate2.Text,
       txtRealDate1.Text,txtRealDate2.Text );
      CheckIfAlreadySelect;
    finally
      FSourceDataSet.EnableControls;
    end;
  finally
    Screen.Cursor := crDefault;
  end;

  btnSelect.Enabled := ( FSourceDataSet.RecordCount > 0 );

  if ( aCount <= 0 ) then
  begin
    WarningMsg( '查無資料,請重新查詢。' );
    txtCustId.Text := '';
    txtShouldDate1.Clear;
    txtShouldDate2.Clear;
    txtRealDate1.Clear;
    txtRealDate2.Clear;
    txtMduId.Clear;
    txtCustId.SetFocus;
  end;

end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA04.btnCancelClick(Sender: TObject);
begin
  NoDataStatus;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA04.HaveDataStatus;
begin
  TxtInvId.Enabled := false;
  txtCustId.Enabled := true;
  //#6001 增加大樓、實收、應收條件 By Kin 2011/07/13
  txtMduId.Enabled := True;
  txtShouldDate1.Enabled := True;
  txtShouldDate2.Enabled := True;
  txtRealDate1.Enabled := True;
  txtRealDate2.Enabled := True;

  btnQueryInvInfo.Enabled := false;
  btnExit.Enabled := false;
  btnQueryCust.Enabled := true;
  btnImportExcel.Enabled := True;
  btnUpdate.Enabled := false;
  btnCancel.Enabled := true;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA04.NoDataStatus;
begin
  PrepareChargeDataSet;
  FMasterDataSet.EmptyDataSet;

  TxtInvId.Enabled := true;
  TxtInvId.Text := '';
  TxtInvId.SetFocus;
  txtCustId.Text := '';
  txtCustId.Enabled := false;
  //#6001 增加大樓、實收、應收條件 By Kin 2011/07/13
  txtMduId.Clear;
  txtShouldDate1.Clear;
  txtShouldDate2.Clear;
  txtRealDate1.Clear;
  txtRealDate2.Clear;
  txtMduId.Enabled := False;
  txtShouldDate1.Enabled := False;
  txtShouldDate2.Enabled := False;
  txtRealDate1.Enabled := False;
  txtRealDate2.Enabled := False;


  lblInvDate.Caption := '';
  lblInvTitle.Caption := '';
  lblInvFormat.Caption := '';
  lblTaxType.Caption := '';
  lblTaxRate.Caption := '';
  lblCustName.Caption := '';

  btnQueryInvInfo.Enabled := true;
  btnExit.Enabled := true;
  btnQueryCust.Enabled := false;
  btnImportExcel.Enabled := False;
  btnCancel.Enabled := false;

  btnSelect.Enabled := False;
  btnUnSelect.Enabled := False;
  btnUpdate.Enabled := false;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA04.txtInvIdKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_Return then btnQueryInvInfoClick(Sender);
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA04.txtCustIdKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_Return then btnQueryCustClick( btnQueryCust );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA04.SetDBGridColumn;
var
  aIndex: Integer;
begin

  {
  cxGrid1DBTableView1SELECTED.Width := 63;
  cxGrid1DBTableView1SOURCE.Width := 102;
  cxGrid1DBTableView1SOURCENAME.Width := 98;
  cxGrid1DBTableView1COMPCODE.Width := 88;
  cxGrid1DBTableView1CUSTID.Width := 97;
  cxGrid1DBTableView1BILLNO.Width := 125;
  cxGrid1DBTableView1ITEM.Width := 57;
  cxGrid1DBTableView1CITEMCODE.Width := 131;
  cxGrid1DBTableView1CITEMNAME.Width := 140;
  cxGrid1DBTableView1SHOULDDATE.Width := 88;
  cxGrid1DBTableView1SHOULDAMT.Width := 110;
  cxGrid1DBTableView1REALDATE.Width := 83;
  cxGrid1DBTableView1REALAMT.Width := 108;
  cxGrid1DBTableView1REALPERIOD.Width := 72;
  cxGrid1DBTableView1REALSTARTDATE.Width := 97;
  cxGrid1DBTableView1REALSTOPDATE.Width := 97;
  cxGrid1DBTableView1CLCTNAME.Width := 126;
  cxGrid1DBTableView1STNAME.Width := 162;
  cxGrid1DBTableView1ACCOUNTNO.Width := 155;
  cxGrid1DBTableView1SERVICETYPE.Width := 84;
  for aIndex := 0 to cxGrid1DBTableView1.ColumnCount - 1 do
    cxGridDBTableView1.Columns[aIndex].Width := cxGrid1DBTableView1.Columns[aIndex].Width;
    }
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA04.txtCustIdExit(Sender: TObject);
var
  aOut: Integer;
begin
  lblCustName.Caption := '';
  if TryStrToInt( txtCustId.Text, aOut ) then
  begin
    dtmSO.adoComm.Close;
    dtmSO.adoComm.SQL.Text := Format(
      '  SELECT CUSTNAME FROM %s.SO001 WHERE CUSTID = %d ',
      [dtmMain.getMisDbOwner, aOut] );
    dtmSO.adoComm.Open;
    lblCustName.Caption := dtmSO.adoComm.FieldByName( 'CUSTNAME' ).AsString;
    dtmSO.adoComm.Close;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA04.btnUpdateClick(Sender: TObject);
var
  aSQL, aTableName, aPreInvoice, aErrMsg: String;
  aSuccessCount: Integer;
begin

  if FMasterDataSet.IsEmpty then Exit;

  if not ChecckBeforeUpdate then
  begin
    WarningMsg( '請選取須要回填發票號碼的收費項目。' );
    Exit;
  end; 

  // 0:後開   1:預開   2:現場開 ---- 由此程式產生的一定是現場開
  aPreInvoice := '2';


  if not ConfirmMsg( '確認更新這些選取的資料 ?' ) then
   Exit;

  dtmMain.adoComm.Close;
  aSQL :=
    '  UPDATE %s                                         ' +
    '     SET GUINO = ''%s'',                            ' +
    '         PREINVOICE = %s,                           ' +
    '         INVDATE = TO_DATE( ''%s'', ''YYYYMMDD'' ), ' +
    '         INVOICETIME = SYSDATE                      ' +
    '   WHERE BILLNO = ''%s''                            ' +
    '     AND ITEM = ''%s''                              ';

  Screen.Cursor := crSQLWait;
  try
    FMasterDataSet.DisableControls;
    try
      dtmSO.adoComm.Connection.BeginTrans;
      try
        aSuccessCount := 0;
        aErrMsg := EmptyStr;
        FMasterDataSet.First;
        while not FMasterDataSet.Eof do
        begin
          if ( FMasterDataSet.FieldByName( 'SELECTED' ).AsString = 'Y' ) then
          begin
            aTableName := dtmMain.getMisDbOwner + '.SO033';
            if ( FMasterDataSet.FieldByName( 'SOURCE' ).AsString = '34' ) then
              aTableName := dtmMain.getMisDbOwner + '.SO034';
            dtmSO.adoComm.SQL.Text := Format( aSQL,
              [ aTableName, FInvId, aPreInvoice, FormatDateTime( 'yyyymmdd', FInvDate ),
                FMasterDataSet.FieldByName( 'BILLNO' ).AsString,
                FMasterDataSet.FieldByName( 'ITEM' ).AsString] );
            dtmSO.adoComm.ExecSQL;
            Inc( aSuccessCount );
            Application.ProcessMessages;
          end;
          FMasterDataSet.Next;
        end;
        dtmSO.adoComm.Connection.CommitTrans;
      except
        on E: Exception do
        begin
          dtmSO.adoComm.Connection.RollbackTrans;
          aSuccessCount := 0;
          aErrMsg := E.Message;
        end;
      end;
    finally
      FMasterDataSet.EnableControls;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
  FMasterDataSet.Active := False;
  FMasterDataSet.Active := True;
  btnUpdate.Enabled := ( FMasterDataSet.RecordCount > 0 );
  if ( aErrMsg <> EmptyStr ) then
  begin
    ErrorMsg( Format( '發票號碼回填失敗, 原因: %s。', [aErrMsg] ) );
  end else
  if ( aSuccessCount > 0 ) then
  begin
    InfoMsg( Format( '發票號碼回填完成, 共回填 %d 筆資料。', [aSuccessCount] ) );
    NoDataStatus;
  end else
  begin
    FMasterDataSet.First;
    InfoMsg( '請選取須要回填的收費項目。' );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA04.txtCustIdPropertiesChange(Sender: TObject);
begin
  PrepareChargeDataSet;
  lblCustName.Caption := '';
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA04.PrepareDataSet;
begin
  if VarIsNull( cdsCharge.Data ) then
    cdsCharge.CreateDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA04.btnSelectClick(Sender: TObject);
var
  aIndex: Integer;
begin
  if FSourceDataSet.IsEmpty then Exit;
  FSourceDataSet.DisableControls;
  try
    FSourceDataSet.First;
    while not FSourceDataSet.Eof do
    begin
      if ( FSourceDataSet.FieldByName( 'SELECTED' ).AsString = 'Y' ) then
      begin
        FMasterDataSet.Append;
        for aIndex := 1 to FMasterDataSet.FieldCount - 1 do
          FMasterDataSet.Fields[aIndex].Value := FSourceDataSet.Fields[aIndex].Value;
        FMasterDataSet.FieldByName( 'SELECTED' ).AsString := 'Y';
        FMasterDataSet.Post;
        FSourceDataSet.Delete;
      end
      else FSourceDataSet.Next;
    end;
  finally
    FSourceDataSet.EnableControls;
  end;
  btnUpdate.Enabled := ( FMasterDataSet.RecordCount > 0 );
  btnUnSelect.Enabled := ( FMasterDataSet.RecordCount > 0 );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA04.CheckIfAlreadySelect;
begin
  FSourceDataSet.First;
  while not FSourceDataSet.Eof do
  begin
    if FMasterDataSet.Locate( 'CUSTID;BILLNO;ITEM', VarArrayOf(
      [FSourceDataSet.FieldByName( 'CUSTID' ).Value,
       FSourceDataSet.FieldByName( 'BILLNO' ).Value,
       FSourceDataSet.FieldByName( 'ITEM' ).Value] ), [] ) then
      FSourceDataSet.Delete
    else begin
      FSourceDataSet.Next;
    end;
  end;

end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA04.GetChargeInfo(const aCustId, aTaxRate, aTaxType,
  aServiceType,aMduId, aShouldDate1, aShouldDate2,
  aRealDate1, aRealDate2: String): Integer;
var
  aSql, aWhere: String;
begin

  {#6001 增加大樓編號、應收日期、實收日期條件 By Kin 2011/07/14}
  aWhere := ' 1 = 1 ';
  if aMduId <> EmptyStr then
  begin
    aWhere := aWhere + Format( ' AND D.MDUID = ''%s'' ',[aMduId]);
  end;
  if aShouldDate1 <> EmptyStr then
  begin
    aWhere := aWhere + Format( ' AND SHOULDDATE >= TO_DATE(''%s'',''yyyymmdd'') ',
                  [GetDate(aShouldDate1)]);
  end;
  if aShouldDate2 <> EmptyStr then
  begin
    aWhere := aWhere + Format( ' AND SHOULDDATE <= TO_DATE(''%s'',''yyyymmdd'') ',
                  [GetDate(aShouldDate2)]);
  end;
  if aRealDate1 <> EmptyStr then
  begin
    aWhere := aWhere + Format( ' AND REALDATE >= TO_DATE(''%s'',''yyyymmdd'') ',
                  [GetDate(aRealDate1)]);
  end;
  if aRealDate2 <> EmptyStr then
  begin
    aWhere := aWhere + Format( ' AND REALDATE <= TO_DATE(''%s'',''yyyymmdd'') ',
                  [GetDate(aRealDate2)]);
  end;
  if aCustId <> EmptyStr then
  begin
    aWhere := aWhere + Format( 'AND A.CUSTID = %s',[aCustId]);
  end;
  PrepareChargeDataSet;
  aSql := Format(
    '     SELECT 33 SOURCE,                                  ' +
    '           ''未日結'' SOURCENAME,                       ' +
    '           A.COMPCODE,                                  ' +
    '           A.CUSTID,                                    ' +
    '           D.CUSTNAME,                                  ' +
    '           A.BILLNO,                                    ' +
    '           A.ITEM,                                      ' +
    '           A.CITEMCODE,                                 ' +
    '           A.CITEMNAME,                                 ' +
    '           A.SHOULDDATE,                                ' +
    '           A.SHOULDAMT,                                 ' +
    '           A.REALDATE,                                  ' +
    '           A.REALAMT,                                   ' +
    '           A.REALPERIOD,                                ' +
    '           A.REALSTARTDATE,                             ' +
    '           A.REALSTOPDATE,                              ' +
    '           A.CLCTNAME,                                  ' +
    '           A.STNAME,                                    ' +
    '           A.ACCOUNTNO,                                 ' +
    '           A.SERVICETYPE,                               ' +
    '           C.RATE1,                                     ' +
    '           B.TAXCODE                                    ' +
    '      FROM %s.SO033 A, %s.CD019 B, %s.CD033 C,          ' +
    '           %s.SO001 D                                   ' +
    '     WHERE A.CITEMCODE = B.CODENO                       ' +
    '       AND B.TAXCODE = C.CODENO                         ' +
    '       AND ( B.TAXCODE IS NOT NULL OR C.TAXFLAG = 1)    ' +
    '       AND A.CUSTID = D.CUSTID                          ' +
    '       AND A.CANCELFLAG = 0                             ' +
    '       AND A.GUINO IS NULL                              ' +
    '       AND A.INVOICETIME IS NULL                        ' +
    '       AND C.RATE1 = %s                                 ' +
    '       AND B.TAXCODE = %s                               ' +
    '       AND A.SERVICETYPE IN ( ' + aServiceType + ' )    ' +
    '       AND %s                                           ' +
    '  UNION ALL                                             ' +
    '     SELECT 34 SOURCE,                                  ' +
    '           ''已日結'' SOURCENAME,                       ' +
    '            A.COMPCODE,                                 ' +
    '            A.CUSTID,                                   ' +
    '            D.CUSTNAME,                                 ' +
    '            A.BILLNO,                                   ' +
    '            A.ITEM,                                     ' +
    '            A.CITEMCODE,                                ' +
    '            A.CITEMNAME,                                ' +
    '            A.SHOULDDATE,                               ' +
    '            A.SHOULDAMT,                                ' +
    '            A.REALDATE,                                 ' +
    '            A.REALAMT,                                  ' +
    '            A.REALPERIOD,                               ' +
    '            A.REALSTARTDATE,                            ' +
    '            A.REALSTOPDATE,                             ' +
    '            A.CLCTNAME,                                 ' +
    '            A.STNAME,                                   ' +
    '            A.ACCOUNTNO,                                ' +
    '            A.SERVICETYPE,                              ' +
    '            C.RATE1,                                    ' +
    '            B.TAXCODE                                   ' +
    '       FROM %s.SO034 A, %s.CD019 B, %s.CD033 C,         ' +
    '            %s.SO001 D                                  ' +
    '      WHERE A.CITEMCODE = B.CODENO                      ' +
    '        AND B.TAXCODE = C.CODENO                        ' +
    '        AND ( B.TAXCODE IS NOT NULL OR C.TAXFLAG = 1 )  ' +
    '        AND A.CUSTID = D.CUSTID                         ' +
    '        AND A.CANCELFLAG = 0                            ' +
    '        AND A.GUINO IS NULL                             ' +
    '        AND A.INVOICETIME IS NULL                       ' +
    '        AND C.RATE1 = %s                                ' +
    '        AND B.TAXCODE = %s                              ' +
    '        AND A.SERVICETYPE IN ( ' + aServiceType + ' )   ' +
    '        AND %s                                          ' +
    '     ORDER BY BILLNO, ITEM                              ',
    [ dtmMain.getMisDbOwner, dtmMain.getMisDbOwner, dtmMain.getMisDbOwner, dtmMain.getMisDbOwner,
       aTaxRate, aTaxType, aWhere,
      dtmMain.getMisDbOwner, dtmMain.getMisDbOwner, dtmMain.getMisDbOwner, dtmMain.getMisDbOwner,
      aTaxRate, aTaxType, aWhere ] );
  dtmSO.adoComm.Close;
  dtmSO.adoComm.SQL.Text := aSql;
  dtmSO.adoComm.Open;
  dtmSO.adoComm.First;
  while not dtmSO.adoComm.Eof do
  begin
    cdsV_ChargeInfo.Append;
    cdsV_ChargeInfo.FieldByName( 'SELECTED' ).AsString := 'N';
    cdsV_ChargeInfo.FieldByName( 'SOURCE' ).AsString := dtmSO.adoComm.FieldByName( 'SOURCE' ).AsString;
    cdsV_ChargeInfo.FieldByName( 'SOURCENAME' ).AsString := dtmSO.adoComm.FieldByName( 'SOURCENAME' ).AsString;
    cdsV_ChargeInfo.FieldByName( 'COMPCODE' ).AsString := dtmSO.adoComm.FieldByName( 'COMPCODE' ).AsString;
    cdsV_ChargeInfo.FieldByName( 'CUSTID' ).AsString := dtmSO.adoComm.FieldByName( 'CUSTID' ).AsString;
    cdsV_ChargeInfo.FieldByName( 'CUSTNAME' ).AsString := dtmSO.adoComm.FieldByName( 'CUSTNAME' ).AsString;
    cdsV_ChargeInfo.FieldByName( 'BILLNO' ).AsString := dtmSO.adoComm.FieldByName( 'BILLNO' ).AsString;
    cdsV_ChargeInfo.FieldByName( 'ITEM' ).AsString := dtmSO.adoComm.FieldByName( 'ITEM' ).AsString;
    cdsV_ChargeInfo.FieldByName( 'CITEMCODE' ).AsString := dtmSO.adoComm.FieldByName( 'CITEMCODE' ).AsString;
    cdsV_ChargeInfo.FieldByName( 'CITEMNAME' ).AsString := dtmSO.adoComm.FieldByName( 'CITEMNAME' ).AsString;
    {}
    if not VarIsNull( dtmSO.adoComm.FieldByName( 'SHOULDDATE' ).Value ) then
      cdsV_ChargeInfo.FieldByName( 'SHOULDDATE' ).AsDateTime := dtmSO.adoComm.FieldByName( 'SHOULDDATE' ).AsDateTime;
    cdsV_ChargeInfo.FieldByName( 'SHOULDAMT' ).AsString := dtmSO.adoComm.FieldByName( 'SHOULDAMT' ).AsString;
    {}
    if not VarIsNull( dtmSO.adoComm.FieldByName( 'REALDATE' ).Value ) then
      cdsV_ChargeInfo.FieldByName( 'REALDATE' ).AsDateTime := dtmSO.adoComm.FieldByName( 'REALDATE' ).AsDateTime;
    {}
    cdsV_ChargeInfo.FieldByName( 'REALAMT' ).AsString := dtmSO.adoComm.FieldByName( 'REALAMT' ).AsString;
    cdsV_ChargeInfo.FieldByName( 'REALPERIOD' ).AsString := dtmSO.adoComm.FieldByName( 'REALPERIOD' ).AsString;
    {}
    if not VarIsNull( dtmSO.adoComm.FieldByName( 'REALSTARTDATE' ).Value ) then
      cdsV_ChargeInfo.FieldByName( 'REALSTARTDATE' ).AsDateTime := dtmSO.adoComm.FieldByName( 'REALSTARTDATE' ).AsDateTime;
    {}
    if not VarIsNull( dtmSO.adoComm.FieldByName( 'REALSTOPDATE' ).Value ) then
      cdsV_ChargeInfo.FieldByName( 'REALSTOPDATE' ).AsDateTime := dtmSO.adoComm.FieldByName( 'REALSTOPDATE' ).AsDateTime;
    {}
    cdsV_ChargeInfo.FieldByName( 'CLCTNAME' ).AsString := dtmSO.adoComm.FieldByName( 'CLCTNAME' ).AsString;
    cdsV_ChargeInfo.FieldByName( 'STNAME' ).AsString := dtmSO.adoComm.FieldByName( 'STNAME' ).AsString;
    cdsV_ChargeInfo.FieldByName( 'ACCOUNTNO' ).AsString := dtmSO.adoComm.FieldByName( 'ACCOUNTNO' ).AsString;
    cdsV_ChargeInfo.FieldByName( 'SERVICETYPE' ).AsString := dtmSO.adoComm.FieldByName( 'SERVICETYPE' ).AsString;
    cdsV_ChargeInfo.FieldByName( 'RATE1' ).AsString := dtmSO.adoComm.FieldByName( 'RATE1' ).AsString;
    cdsV_ChargeInfo.FieldByName( 'TAXCODE' ).AsString := dtmSO.adoComm.FieldByName( 'TAXCODE' ).AsString;
    cdsV_ChargeInfo.Post;
    dtmSO.adoComm.Next;
  end;
  Result := cdsV_ChargeInfo.RecordCount;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA04.btnUnSelectClick(Sender: TObject);
var
  aCurrentCustId: String;
  aIndex: Integer;
begin
  aCurrentCustId := Trim( txtCustId.Text );
  FMasterDataSet.DisableControls;
  try
    FMasterDataSet.First;
    while not FMasterDataSet.Eof do
    begin
      if FMasterDataSet.FieldByName( 'SELECTED' ).AsString = 'Y' then
      begin
        if FMasterDataSet.FieldByName( 'CUSTID' ).AsString = aCurrentCustId then
        begin
          FSourceDataSet.Append;
          for aIndex := 1 to FMasterDataSet.FieldCount - 1 do
             FSourceDataSet.Fields[aIndex].Value := FMasterDataSet.Fields[aIndex].Value;
          FSourceDataSet.Post;   
        end;
        FMasterDataSet.Delete;
      end
      else FMasterDataSet.Next;
    end;
  finally
    FMasterDataSet.EnableControls;
  end;
  btnUpdate.Enabled := ( FMasterDataSet.RecordCount > 0 );
  btnUnSelect.Enabled := ( FMasterDataSet.RecordCount > 0 );
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA04.ChecckBeforeUpdate: Boolean;
begin
  Result := False;
  FMasterDataSet.DisableControls;
  try
    FMasterDataSet.First;
    while not FMasterDataSet.Eof do
    begin
      Result := ( FMasterDataSet.FieldByName( 'SELECTED' ).AsString = 'Y' );
      if Result then Break;
      FMasterDataSet.Next;
    end;
  finally
    FMasterDataSet.EnableControls;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA04.PrepareChargeDataSet;
begin
  if not VarIsNull( cdsV_ChargeInfo.Data ) then
    FSourceDataSet.EmptyDataSet;
  FSourceDataSet.Data := Null;
  FSourceDataSet.CreateDataSet;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA04.FormCreate(Sender: TObject);
begin
  txtMduId.Connection := dtmMain.InvConnection;
  txtMduId.CodeNoFieldName := 'MDUID';
  txtMduId.DescriptionFieldName := 'NAME';
  txtMduId.Initializa;
end;

procedure TfrmInvA04.txtMduIdCodeNamePropertiesInitPopup(Sender: TObject);
var aSQL : String;
begin
   aSQL := format( ' SELECT MDUID, NAME                   ' +
                    '   FROM %s.SO017 ORDER BY MDUID      ',
                    [dtmMain.getMisDbOwner]);

  txtMduId.SQL.Text := aSQL;
  txtMduId.CodeNamePropertiesInitPopup( Sender );
end;

function TfrmInvA04.GetDate(const aDate: String): String;
var aStrTmp : string;
begin
  if aDate = EmptyStr then
  begin
    Result := EmptyStr;
  end else
  begin
    aStrTmp := StringReplace( aDate,'/','',[rfReplaceAll]);
    Result := aStrTmp;
  end;
end;

procedure TfrmInvA04.btnSelAllClick(Sender: TObject);
  var
    aBK : TBookmark;
begin
  Screen.Cursor := crSQLWait;
  try
    cdsV_ChargeInfo.DisableControls;
    try
      if Assigned( cdsV_ChargeInfo ) then
      begin
        if not cdsV_ChargeInfo.IsEmpty then
        begin
          try
            aBK := cdsV_ChargeInfo.GetBookmark;
            cdsV_ChargeInfo.First;

            while not cdsV_ChargeInfo.Eof do
            begin
              cdsV_ChargeInfo.Edit;
              cdsV_ChargeInfo.FieldByName( 'SELECTED' ).AsString := 'Y';
              cdsV_ChargeInfo.Post;
              cdsV_ChargeInfo.Next;
            end;
          finally
            cdsV_ChargeInfo.GotoBookmark( aBK );
            cdsV_ChargeInfo.FreeBookmark( aBK );
          end;
        end;
      end;

    finally
      cdsV_ChargeInfo.EnableControls;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TfrmInvA04.cxGrid1DBTableView1SELECTEDCustomDrawFooterCell(
  Sender: TcxGridTableView; ACanvas: TcxCanvas;
  AViewInfo: TcxGridColumnHeaderViewInfo; var ADone: Boolean);
begin



//  AViewInfo.Text := '1';
end;

procedure TfrmInvA04.btnImportExcelClick(Sender: TObject);
var connectString : string;
    aCount: Integer;
    aServiceType : String;
    aErrCount: Integer;
    aErrMsg : String;
    aForm : TfrmInvA04_2;
begin
  dlgOpen1.Filter := 'Excel files (*.xls)|*.xls';
  adoImportExcel := TADOQuery.Create(Self);
  adoImportExcel.Close;
  aErrCount := 0;
  aCount := 0;
  aErrMsg := '';
  if dlgOpen1.Execute then
  begin

    try
      try
         connectString := 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + dlgOpen1.FileName +
                            ';Extended Properties=Excel 8.0;Persist Security Info=False';

        adoImportExcel.ConnectionString := connectString;
        adoImportExcel.SQL.Add('select * from [工作表1$]');
        adoImportExcel.Open;
        if ( adoImportExcel.FindField('客編') = nil )  then
        begin
          MessageDlg('[ 客編 ] 欄位不存在!',mtError,[mbOk],0);
          Exit;
        end;
        if ( adoImportExcel.FindField('單據編號') = nil )  then
        begin
          MessageDlg('[ 單據編號 ]欄位不存在!',mtError,[mbOk],0);
          Exit;
        end;
        if ( adoImportExcel.FindField('項次') = nil )  then
        begin
          MessageDlg('[ 項次 ] 欄位不存在!',mtError,[mbOk],0);
          Exit;
        end;
        if (adoImportExcel.RecordCount = 0 ) then
        begin
          MessageDlg('無任何資料可匯入!',mtError,[mbOk],0);
          Exit;
        end;

        aServiceType := dtmMain.GetServiceTypeStrEx;

        Screen.Cursor := crSQLWait;
        try
          FSourceDataSet.DisableControls;
          try
            aCount := GetExcelChargeInfo( FloatToStr( dG_TaxRate ),
             sG_TaxType, aServiceType,aErrCount,aErrMsg );
             MessageDlg( Format('成功 %d  筆 ' + #10#13 + '失敗 %d 筆 ',[aCount,aErrCount]),mtError,[mbOk],0);
            if ( aErrCount > 0 ) and (aErrMsg <> '') then
            begin
               try
               aForm := TfrmInvA04_2.Create(nil);
                try
                  aForm.ImportErr := aErrMsg;
                  aForm.ShowModal;
                finally
                  aForm.Free;
                end;
              finally
//                Screen.Cursor := crDefault;
              end;
            end;
            CheckIfAlreadySelect;
          finally
            FSourceDataSet.EnableControls;
          end;
        finally
          Screen.Cursor := crDefault;
        end;

        btnSelect.Enabled := ( FSourceDataSet.RecordCount > 0 );
        if ( aCount <= 0 ) then
        begin

        end;
      except
         MessageDlg('Excel格式錯誤!',mtError,[mbOk],0);
      end;


    finally
      adoImportExcel.Close;
      adoImportExcel.Free;
    end;



  end;
end;
//#7609
function TfrmInvA04.GetExcelChargeInfo(const aTaxRate, aTaxType,
      aServiceType : String;var aErrCount : Integer; var aErrMsg : string): Integer;
 var
  aSql, aWhere: String;
  aErrString : TStringList;
begin
  PrepareChargeDataSet;
  adoImportExcel.First;
  aErrString := TStringList.Create;
  try
    while not adoImportExcel.Eof do
    begin
      aWhere := ' 1 = 1 ';
      aWhere :=  aWhere + Format( ' AND A.CUSTID = %s ',[adoImportExcel.FieldByName('客編').AsString]);
      aWhere :=  aWhere + Format( ' AND A.BILLNO = ''%s'' ',[adoImportExcel.FieldByName('單據編號').AsString]);
      aWhere :=  aWhere + Format( ' AND A.ITEM = %s ',[adoImportExcel.FieldByName('項次').AsString]);
       aSql := Format(
      '     SELECT 33 SOURCE,                                  ' +
      '           ''未日結'' SOURCENAME,                       ' +
      '           A.COMPCODE,                                  ' +
      '           A.CUSTID,                                    ' +
      '           D.CUSTNAME,                                  ' +
      '           A.BILLNO,                                    ' +
      '           A.ITEM,                                      ' +
      '           A.CITEMCODE,                                 ' +
      '           A.CITEMNAME,                                 ' +
      '           A.SHOULDDATE,                                ' +
      '           A.SHOULDAMT,                                 ' +
      '           A.REALDATE,                                  ' +
      '           A.REALAMT,                                   ' +
      '           A.REALPERIOD,                                ' +
      '           A.REALSTARTDATE,                             ' +
      '           A.REALSTOPDATE,                              ' +
      '           A.CLCTNAME,                                  ' +
      '           A.STNAME,                                    ' +
      '           A.ACCOUNTNO,                                 ' +
      '           A.SERVICETYPE,                               ' +
      '           A.GUINO,                                     ' +
      '           A.INVOICETIME,                               ' +
      '           C.RATE1,                                     ' +
      '           B.TAXCODE                                    ' +
      '      FROM %s.SO033 A, %s.CD019 B, %s.CD033 C,          ' +
      '           %s.SO001 D                                   ' +
      '     WHERE A.CITEMCODE = B.CODENO                       ' +
      '       AND B.TAXCODE = C.CODENO                         ' +
      '       AND ( B.TAXCODE IS NOT NULL OR C.TAXFLAG = 1)    ' +
      '       AND A.CUSTID = D.CUSTID                          ' +
      '       AND A.CANCELFLAG = 0                             ' +
      '       AND C.RATE1 = %s                                 ' +
      '       AND B.TAXCODE = %s                               ' +
      '       AND A.SERVICETYPE IN ( ' + aServiceType + ' )    ' +
      '       AND %s                                           ' +
      '  UNION ALL                                             ' +
      '     SELECT 34 SOURCE,                                  ' +
      '           ''已日結'' SOURCENAME,                       ' +
      '            A.COMPCODE,                                 ' +
      '            A.CUSTID,                                   ' +
      '            D.CUSTNAME,                                 ' +
      '            A.BILLNO,                                   ' +
      '            A.ITEM,                                     ' +
      '            A.CITEMCODE,                                ' +
      '            A.CITEMNAME,                                ' +
      '            A.SHOULDDATE,                               ' +
      '            A.SHOULDAMT,                                ' +
      '            A.REALDATE,                                 ' +
      '            A.REALAMT,                                  ' +
      '            A.REALPERIOD,                               ' +
      '            A.REALSTARTDATE,                            ' +
      '            A.REALSTOPDATE,                             ' +
      '            A.CLCTNAME,                                 ' +
      '            A.STNAME,                                   ' +
      '            A.ACCOUNTNO,                                ' +
      '            A.SERVICETYPE,                              ' +
      '            A.GUINO,                                    ' +
      '            A.INVOICETIME,                              ' +
      '            C.RATE1,                                    ' +
      '            B.TAXCODE                                   ' +
      '       FROM %s.SO034 A, %s.CD019 B, %s.CD033 C,         ' +
      '            %s.SO001 D                                  ' +
      '      WHERE A.CITEMCODE = B.CODENO                      ' +
      '        AND B.TAXCODE = C.CODENO                        ' +
      '        AND ( B.TAXCODE IS NOT NULL OR C.TAXFLAG = 1 )  ' +
      '        AND A.CUSTID = D.CUSTID                         ' +
      '        AND A.CANCELFLAG = 0                            ' +
      '        AND C.RATE1 = %s                                ' +
      '        AND B.TAXCODE = %s                              ' +
      '        AND A.SERVICETYPE IN ( ' + aServiceType + ' )   ' +
      '        AND %s                                          ' +
      '     ORDER BY BILLNO, ITEM                              ',
      [ dtmMain.getMisDbOwner, dtmMain.getMisDbOwner, dtmMain.getMisDbOwner, dtmMain.getMisDbOwner,
         aTaxRate, aTaxType, aWhere,
        dtmMain.getMisDbOwner, dtmMain.getMisDbOwner, dtmMain.getMisDbOwner, dtmMain.getMisDbOwner,
        aTaxRate, aTaxType, aWhere ] );
      dtmSO.adoComm.Close;
      dtmSO.adoComm.SQL.Text := aSql;
      dtmSO.adoComm.Open;
      if (dtmSO.adoComm.RecordCount = 0 ) then
      begin
        aErrString.Add('[' + adoImportExcel.FieldByName('單據編號').AsString + ' 項次:' +
                            adoImportExcel.FieldByName('項次').AsString + '] 找不到單據');

        adoImportExcel.Next;
        Continue;
      end;
      if (dtmSO.adoComm.FieldByName('SOURCE').AsString = '33' ) then
      begin
        if (dtmSO.adoComm.FieldByName('GUINO').AsString <> '') and
            (dtmSO.adoComm.FieldByName('INVOICETIME').AsString <> '') then
        begin
          aErrString.Add('[' + adoImportExcel.FieldByName('單據編號').AsString + ' 項次:' +
                          adoImportExcel.FieldByName('項次').AsString + '] 已開過發票');
          adoImportExcel.Next;
          Continue;
        end;
      end else
      begin
        if (dtmSO.adoComm.FieldByName('GUINO').AsString <> '') and
            (dtmSO.adoComm.FieldByName('INVOICETIME').AsString <> '') then
        begin
          aErrString.Add('[' + adoImportExcel.FieldByName('單據編號').AsString + ' 項次:' +
                           adoImportExcel.FieldByName('項次').AsString + '] 已開過發票');
          adoImportExcel.Next;
          Continue;
        end;
      end;

      //已拋檔過
      if (dtmSO.adoComm.FieldByName('GUINO').AsString = '') and
            (dtmSO.adoComm.FieldByName('INVOICETIME').AsString <> '') then
      begin
        aErrString.Add('[' + adoImportExcel.FieldByName('單據編號').AsString + ' 項次:' +
                         adoImportExcel.FieldByName('項次').AsString + '] 已拋過檔');
        adoImportExcel.Next;
        Continue;
      end;
      while not dtmSO.adoComm.Eof do
      begin
        cdsV_ChargeInfo.Append;
        cdsV_ChargeInfo.FieldByName( 'SELECTED' ).AsString := 'Y';
        cdsV_ChargeInfo.FieldByName( 'SOURCE' ).AsString := dtmSO.adoComm.FieldByName( 'SOURCE' ).AsString;
        cdsV_ChargeInfo.FieldByName( 'SOURCENAME' ).AsString := dtmSO.adoComm.FieldByName( 'SOURCENAME' ).AsString;
        cdsV_ChargeInfo.FieldByName( 'COMPCODE' ).AsString := dtmSO.adoComm.FieldByName( 'COMPCODE' ).AsString;
        cdsV_ChargeInfo.FieldByName( 'CUSTID' ).AsString := dtmSO.adoComm.FieldByName( 'CUSTID' ).AsString;
        cdsV_ChargeInfo.FieldByName( 'CUSTNAME' ).AsString := dtmSO.adoComm.FieldByName( 'CUSTNAME' ).AsString;
        cdsV_ChargeInfo.FieldByName( 'BILLNO' ).AsString := dtmSO.adoComm.FieldByName( 'BILLNO' ).AsString;
        cdsV_ChargeInfo.FieldByName( 'ITEM' ).AsString := dtmSO.adoComm.FieldByName( 'ITEM' ).AsString;
        cdsV_ChargeInfo.FieldByName( 'CITEMCODE' ).AsString := dtmSO.adoComm.FieldByName( 'CITEMCODE' ).AsString;
        cdsV_ChargeInfo.FieldByName( 'CITEMNAME' ).AsString := dtmSO.adoComm.FieldByName( 'CITEMNAME' ).AsString;

        if not VarIsNull( dtmSO.adoComm.FieldByName( 'SHOULDDATE' ).Value ) then
          cdsV_ChargeInfo.FieldByName( 'SHOULDDATE' ).AsDateTime := dtmSO.adoComm.FieldByName( 'SHOULDDATE' ).AsDateTime;
        cdsV_ChargeInfo.FieldByName( 'SHOULDAMT' ).AsString := dtmSO.adoComm.FieldByName( 'SHOULDAMT' ).AsString;

        if not VarIsNull( dtmSO.adoComm.FieldByName( 'REALDATE' ).Value ) then
          cdsV_ChargeInfo.FieldByName( 'REALDATE' ).AsDateTime := dtmSO.adoComm.FieldByName( 'REALDATE' ).AsDateTime;

        cdsV_ChargeInfo.FieldByName( 'REALAMT' ).AsString := dtmSO.adoComm.FieldByName( 'REALAMT' ).AsString;
        cdsV_ChargeInfo.FieldByName( 'REALPERIOD' ).AsString := dtmSO.adoComm.FieldByName( 'REALPERIOD' ).AsString;

        if not VarIsNull( dtmSO.adoComm.FieldByName( 'REALSTARTDATE' ).Value ) then
          cdsV_ChargeInfo.FieldByName( 'REALSTARTDATE' ).AsDateTime := dtmSO.adoComm.FieldByName( 'REALSTARTDATE' ).AsDateTime;

        if not VarIsNull( dtmSO.adoComm.FieldByName( 'REALSTOPDATE' ).Value ) then
          cdsV_ChargeInfo.FieldByName( 'REALSTOPDATE' ).AsDateTime := dtmSO.adoComm.FieldByName( 'REALSTOPDATE' ).AsDateTime;

        cdsV_ChargeInfo.FieldByName( 'CLCTNAME' ).AsString := dtmSO.adoComm.FieldByName( 'CLCTNAME' ).AsString;
        cdsV_ChargeInfo.FieldByName( 'STNAME' ).AsString := dtmSO.adoComm.FieldByName( 'STNAME' ).AsString;
        cdsV_ChargeInfo.FieldByName( 'ACCOUNTNO' ).AsString := dtmSO.adoComm.FieldByName( 'ACCOUNTNO' ).AsString;
        cdsV_ChargeInfo.FieldByName( 'SERVICETYPE' ).AsString := dtmSO.adoComm.FieldByName( 'SERVICETYPE' ).AsString;
        cdsV_ChargeInfo.FieldByName( 'RATE1' ).AsString := dtmSO.adoComm.FieldByName( 'RATE1' ).AsString;
        cdsV_ChargeInfo.FieldByName( 'TAXCODE' ).AsString := dtmSO.adoComm.FieldByName( 'TAXCODE' ).AsString;
        cdsV_ChargeInfo.Post;
        dtmSO.adoComm.Next;
      end;
      dtmSO.adoComm.Close;
      Application.ProcessMessages;
      adoImportExcel.Next;
    end;

  finally
    aErrCount := aErrString.Count;
    aErrMsg := aErrString.Text;
    aErrString.Free;
    aErrString := nil;
  end;

  Result := cdsV_ChargeInfo.RecordCount;
end;

end.

