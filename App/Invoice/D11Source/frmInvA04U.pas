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
  dxSkinsDefaultPainters, dxSkinscxPCPainter;

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
  private
    { Private declarations }
    FMasterDataSet: TClientDataSet;
    FSourceDataSet: TClientDataSet;
    FInvDate : TDate;
    FInvId, sG_TaxType, sG_HowToCreate: String;
    dG_TaxRate : Double;
    procedure HaveDataStatus;
    procedure NoDataStatus;
    procedure SetDBGridColumn;
    procedure PrepareDataSet;
    procedure CheckIfAlreadySelect;
    function ChecckBeforeUpdate: Boolean;
    function GetChargeInfo(const aCustId, aTaxRate, aTaxType,
      aServiceType: String): Integer;
    procedure PrepareChargeDataSet;
  public
    { Public declarations }
  end;

var
  frmInvA04: TfrmInvA04;

implementation

uses cbUtilis, dtmSOU, dtmMainU, frmMainU;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA04.FormShow(Sender: TObject);
begin
  self.Caption := frmMain.GetFormTitleString( 'A04', '?h?????o???}??' );
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
    WarningMsg('?????J?o?????X');
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
      WarningMsg( '?L???o?????X,?????s?d???C' );
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
      sL_InvFormat := '?q?l'
    else if sL_InvFormat = '2' then
      sL_InvFormat := '???G'
    else if sL_InvFormat = '3' then
      sL_InvFormat := '???T';
    lblInvFormat.Caption := sL_InvFormat;

    if sG_TaxType = '1' then
      sL_TaxType := '???|'
    else if sG_TaxType = '2' then
      sL_TaxType := '?s?|?v'
    else if sG_TaxType = '3' then
      sL_TaxType := '?K?|';
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

  aCustID := Trim( txtCustId.Text);
  if ( aCustid = '' ) then
  begin
    WarningMsg( '?????J?????s??' );
    txtCustId.SetFocus;
    Exit;
  end;

  aServiceType := dtmMain.GetServiceTypeStrEx;

  Screen.Cursor := crSQLWait;
  try
    FSourceDataSet.DisableControls;
    try
      aCount := GetChargeInfo( aCustId, FloatToStr( dG_TaxRate ),
       sG_TaxType, aServiceType );
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
    WarningMsg( '?d?L??????????,?????s?d???C' );
    txtCustId.Text := '';
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
  btnQueryInvInfo.Enabled := false;
  btnExit.Enabled := false;
  btnQueryCust.Enabled := true;
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


  lblInvDate.Caption := '';
  lblInvTitle.Caption := '';
  lblInvFormat.Caption := '';
  lblTaxType.Caption := '';
  lblTaxRate.Caption := '';
  lblCustName.Caption := '';

  btnQueryInvInfo.Enabled := true;
  btnExit.Enabled := true;
  btnQueryCust.Enabled := false;
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
    WarningMsg( '?????????n?^???o?????X?????O?????C' );
    Exit;
  end; 

  // 0:???}   1:?w?}   2:?{???} ---- ?????{?????????@?w?O?{???}
  aPreInvoice := '2';


  if not ConfirmMsg( '?T?{???s?o???????????? ?' ) then
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
    ErrorMsg( Format( '?o?????X?^??????, ???]: %s?C', [aErrMsg] ) );
  end else
  if ( aSuccessCount > 0 ) then
  begin
    InfoMsg( Format( '?o?????X?^??????, ?@?^?? %d ???????C', [aSuccessCount] ) );
    NoDataStatus;
  end else
  begin
    FMasterDataSet.First;
    InfoMsg( '?????????n?^???????O?????C' );
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
  aServiceType: String): Integer;
var
  aSql: String;
begin
  PrepareChargeDataSet;
  aSql := Format(
    '     SELECT 33 SOURCE,                                  ' +
    '           ''??????'' SOURCENAME,                       ' +
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
    '       AND A.CUSTID = ''%s''                            ' +
    '       AND C.RATE1 = %s                                 ' +
    '       AND B.TAXCODE = %s                               ' +
    '       AND A.SERVICETYPE IN ( ' + aServiceType + ' )    ' +
    '  UNION ALL                                             ' +
    '     SELECT 34 SOURCE,                                  ' +
    '           ''?w????'' SOURCENAME,                       ' +
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
    '        AND A.CUSTID = ''%s''                           ' +
    '        AND C.RATE1 = %s                                ' +
    '        AND B.TAXCODE = %s                              ' +
    '        AND A.SERVICETYPE IN ( ' + aServiceType + ' )   ' +
    '     ORDER BY BILLNO, ITEM                              ',   
    [ dtmMain.getMisDbOwner, dtmMain.getMisDbOwner, dtmMain.getMisDbOwner, dtmMain.getMisDbOwner,
      aCustId, aTaxRate, aTaxType,
      dtmMain.getMisDbOwner, dtmMain.getMisDbOwner, dtmMain.getMisDbOwner, dtmMain.getMisDbOwner,
     aCustId, aTaxRate, aTaxType ] );
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

end.
