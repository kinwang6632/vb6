unit frmInvB04U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Math, DBClient, ADODB, Provider,
  cxStyles, cxCustomData, cxGraphics, DB, Menus,
  cxFilter, cxData, cxDataStorage, cxEdit, cxDBData,
  cxLookAndFeelPainters, cxButtons, cxProgressBar, cxGridLevel, cxClasses,
  cxControls, cxGridCustomView, cxGridCustomTableView, cxGridTableView,
  cxGridDBTableView, cxGrid, cxContainer, cxGroupBox, cxRadioGroup,
  cxTextEdit, cxMaskEdit, cxDropDownEdit, cxCalendar, cxButtonEdit,
  cxLabel, cxImageComboBox, cxCurrencyEdit, cxCheckBox, dxSkinsCore,
  dxSkinsDefaultPainters, dxSkinscxPCPainter, dxSkinsdxBarPainter,
  dxSkinBlack, dxSkinBlue, dxSkinCaramel, dxSkinCoffee;

type
  TfrmInvB04 = class(TForm)
    Panel1: TPanel;
    Label1: TLabel;
    Panel3: TPanel;
    Panel2: TPanel;
    cxGroupBox1: TcxGroupBox;
    Bevel1: TBevel;
    gvSO034: TcxGridDBTableView;
    glSO034: TcxGridLevel;
    SO034Grid: TcxGrid;
    cmbServiceType: TcxComboBox;
    Label2: TLabel;
    Label3: TLabel;
    Panel5: TPanel;
    rdoMemberY: TcxRadioButton;
    rdoMemberN: TcxRadioButton;
    Label4: TcxLabel;
    txtRealDateSt: TcxDateEdit;
    Label5: TLabel;
    txtRealDateEd: TcxDateEdit;
    Label6: TLabel;
    txtChargeItem: TcxButtonEdit;
    btnQuery: TcxButton;
    Panel4: TPanel;
    PBar: TcxProgressBar;
    btnExec: TcxButton;
    btnClose: TcxButton;
    cdsSO034: TClientDataSet;
    dsSO034: TDataSource;
    gvSO034BILLNO: TcxGridDBColumn;
    gvSO034ITEM: TcxGridDBColumn;
    gvSO034NOTE: TcxGridDBColumn;
    gvSO034CUSTID: TcxGridDBColumn;
    gvSO034CUSTNAME: TcxGridDBColumn;
    gvSO034CITEMCODE: TcxGridDBColumn;
    gvSO034CITEMNAME: TcxGridDBColumn;
    gvSO034REALDATE: TcxGridDBColumn;
    gvSO034REALAMT: TcxGridDBColumn;
    gvSO034GUINO: TcxGridDBColumn;
    cxGroupBox2: TcxGroupBox;
    Label7: TcxLabel;
    txtYearMonth: TcxMaskEdit;
    gvSO034STATUS: TcxGridDBColumn;
    gvSO034ERRMSG: TcxGridDBColumn;
    Panel6: TPanel;
    btnSelect: TcxButton;
    btnRemove: TcxButton;
    gvSO034SELECTED: TcxGridDBColumn;
    Bevel2: TBevel;
    Label8: TLabel;
    gvSO034SALEAMT: TcxGridDBColumn;
    gvSO034TAXAMT: TcxGridDBColumn;
    cxLabel1: TcxLabel;
    txtPaperDate: TcxDateEdit;
    gvSO034RECNO: TcxGridDBColumn;
    adoSysDate: TADOQuery;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure txtChargeItemPropertiesButtonClick(Sender: TObject;
      AButtonIndex: Integer);
    procedure txtRealDateStExit(Sender: TObject);
    procedure btnQueryClick(Sender: TObject);
    procedure txtRealDateStPropertiesValidate(Sender: TObject;
      var DisplayValue: Variant; var ErrorText: TCaption;
      var Error: Boolean);
    procedure btnExecClick(Sender: TObject);
    procedure rdoMemberYClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnSelectClick(Sender: TObject);
    procedure btnRemoveClick(Sender: TObject);
    procedure gvSO034CustomDrawCell(Sender: TcxCustomGridTableView;
      ACanvas: TcxCanvas; AViewInfo: TcxGridTableDataCellViewInfo;
      var ADone: Boolean);
  private
    { Private declarations }
    FYearMonth: String;
    FPaperDate: TDateTime;
    FInvReader: TADOQuery;
    FKeepInvId: String;
    FInv005: TClientDataSet;
    FsysDate: TDateTime;
    FDouble : Boolean;
    FQryWhere : String;
    procedure SetButtonCompetence;
    procedure initialComboBox;
    procedure DoSelectChargeItem;
    function DoGetInvInformation(const aInvId: String; const aFieldName: String): Variant;
    function IsDataOK(const aType: Integer): Boolean;
    function GetSqlText: String;
    procedure CalcAmt(const aTaxType: String; var aReal, aSale, aTax: Integer);
    procedure PrepareSO034DataSet;
    procedure UnPrepareSO034DataSet;
    procedure DoFillSO034DataSet;
    procedure CreateAllowance(var aError: Boolean);
    function DoInsertInv014(var aErrMsg: String; var aRcdCnt,aDetailCnt: Integer ): Boolean;
    function DoUpdateSO034(var aErrMsg: String): Boolean;
    procedure PrepareReportDataSet;
    procedure UnPrepareReportDataSet;
    procedure DoErrorReport;
    procedure DoSelectRecord(const aSelected: Boolean);
    procedure FillChargeItemDataSet;
    function GetDefaultInvDate(const aDate : TDateTime) :String;
  public
    { Public declarations }
  end;

var
  frmInvB04: TfrmInvB04;

implementation

uses cbUtilis, frmMainU, dtmMainU, dtmMainJU, dtmSOU, dtmReportModule,
  frmMultiSelectU, VarConv;

{$R *.dfm}

var
  aQueryParam: TQueryParam;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB04.FormCreate(Sender: TObject);
begin
  ZeroMemory( @aQueryParam, SizeOf( aQueryParam ) );
  btnSelect.Enabled := False;
  btnRemove.Enabled := False;
  FInvReader := dtmMain.adoComm3;
  FInv005 := TClientDataSet.Create( nil );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB04.FormShow(Sender: TObject);
begin
  Self.Caption := frmMain.GetFormTitleString( 'B04', '?????????????@?~' );
  SetButtonCompetence;
  initialComboBox;
  txtYearMonth.Text := FormatDateTime( 'yyyy/mm', Date );
  PrepareSO034DataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB04.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  UnPrepareReportDataSet;
  FInvReader.Close;
  FInv005.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB04.SetButtonCompetence;
var
  aQuery, aInsert, aDelete, aUpdate, aExecute,aVerify: String;
begin
  //?]?w?v??
  dtmMain.getCompetence( 'B03', aQuery, aInsert, aDelete, aUpdate, aExecute, aVerify );
  btnQuery.Enabled := ( aQuery = 'Y' );
  btnExec.Enabled := ( aExecute = 'Y');
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB04.initialComboBox;
var
  aParam: TComboBoxCreateParam;
begin
  aParam.Sql := Format(
    ' select itemid, description from inv022 ' +
    '  where identifyid1 = ''%s''            ' +
    '    and identifyid2 = ''%s''            ' +
    '    and itemid in  ( %s )               ',
    [IDENTIFYID1, IDENTIFYID2, dtmMain.GetServiceTypeStrEx] );
  aParam.KeyField := 'ITEMID';
  aParam.DescField := 'DESCRIPTION';
  aParam.AddAllText := True;
  CreateCxComboBoxItem( cmbServiceType, aParam );
  if ( cmbServiceType.Properties.Items.Count > 0 ) then
    cmbServiceType.ItemIndex := 0;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB04.PrepareSO034DataSet;
begin
  if ( cdsSO034.FieldDefs.Count <= 0 ) then
  begin
    cdsSO034.FieldDefs.Add( 'RECNO', ftInteger );
    cdsSO034.FieldDefs.Add( 'SELECTED', ftBoolean );
    cdsSO034.FieldDefs.Add( 'STATUS', ftInteger );
    cdsSO034.FieldDefs.Add( 'ERRMSG', ftString, 1024 );
    cdsSO034.FieldDefs.Add( 'BILLNO', ftString, 15 );
    cdsSO034.FieldDefs.Add( 'ITEM', ftInteger );
    cdsSO034.FieldDefs.Add( 'NOTE', ftString, 750 );
    cdsSO034.FieldDefs.Add( 'CUSTID', ftInteger );
    cdsSO034.FieldDefs.Add( 'CUSTNAME', ftString, 50 );
    cdsSO034.FieldDefs.Add( 'CITEMCODE', ftInteger );
    cdsSO034.FieldDefs.Add( 'CITEMNAME', ftString, 40 );
    cdsSO034.FieldDefs.Add( 'REALDATE', ftDateTime );
    cdsSO034.FieldDefs.Add( 'REALAMT', ftInteger );
    cdsSO034.FieldDefs.Add( 'SALEAMT', ftInteger );
    cdsSO034.FieldDefs.Add( 'TAXAMT', ftInteger );
    cdsSO034.FieldDefs.Add( 'GUINO', ftString, 20 );
    cdsSO034.FieldDefs.Add( 'ALLOWANCENO', ftString, 12 );
    {}
    cdsSO034.FieldDefs.Add( 'ISFOUND', ftBoolean );
    cdsSO034.FieldDefs.Add( 'ISOBSOLETE', ftString,1 );
    cdsSO034.FieldDefs.Add( 'BUSINESSID', ftString, 8 );
    cdsSO034.FieldDefs.Add( 'INVFORMAT', ftString, 1 );
    cdsSO034.FieldDefs.Add( 'INVDATE', ftDateTime );
    cdsSO034.FieldDefs.Add( 'TAXTYPE', ftString, 1 );
    cdsSO034.FieldDefs.Add( 'INVAMT', ftInteger );
    {}
    { #4222 ?W?[?I???????BServiceType By Kin 2009/03/10 }
    cdsSO034.FieldDefs.Add( 'PTCODE', ftInteger );
    cdsSO034.FieldDefs.Add( 'PTNAME', ftString ,20 );
    cdsSO034.FieldDefs.Add( 'SERVICETYPE', ftString,1);
    {}
    cdsSO034.CreateDataSet;
  end;
  cdsSO034.EmptyDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB04.UnPrepareSO034DataSet;
begin
  if not VarIsNull( cdsSO034.Data ) then cdsSO034.EmptyDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB04.txtChargeItemPropertiesButtonClick(Sender: TObject;
  AButtonIndex: Integer);
begin
  if AButtonIndex = 1 then
    DoSelectChargeItem
  else begin
    txtChargeItem.Clear;
    aQueryParam.Param1 := EmptyStr;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB04.DoSelectChargeItem;
var
  aMemberFlag: String;
begin
  aMemberFlag := '0';
  if ( rdoMemberY.Checked ) then aMemberFlag := '1';
  FillChargeItemDataSet;
  frmMultiSelect := TfrmMultiSelect.Create( Application );
  try
    {
    frmMultiSelect.Connection := dtmMain.InvConnection;
    frmMultiSelect.KeyFields := 'ITEMID';
    frmMultiSelect.KeyValues := aQueryParam.Param1;
    frmMultiSelect.DisplayFields := 'ITEMID,?N?X,DESCRIPTION,???O????';
    frmMultiSelect.ResultFields := 'DESCRIPTION';
    frmMultiSelect.SQL.Text := Format(
      ' SELECT A.ITEMID, A.DESCRIPTION         ' +
      '   FROM INV005 A, %s.CD019%s B          ' +
      '  WHERE A.ITEMID = B.CODENO             ' +
      '    AND A.IDENTIFYID1 = ''%s''          ' +
      '    AND A.IDENTIFYID2 = ''%s''          ' +
      '    AND A.COMPID = ''%s''               ' +
      '    AND A.SIGN = ''-''                  ' +
      '    AND B.MEMBERFLAG = ''%s''           ' +
      '  ORDER BY ITEMID                       ',
      [dtmMain.getMisDbOwner, dtmMain.GetMisDbLink,
       IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, aMemberFlag] );
    }
    frmMultiSelect.KeyFields := 'ITEMID';
    frmMultiSelect.KeyValues := aQueryParam.Param1;
    frmMultiSelect.DisplayFields := 'ITEMID,?N?X,DESCRIPTION,???O????';
    frmMultiSelect.ResultFields := 'DESCRIPTION';
    frmMultiSelect.DataSet := FInv005;
    if ( frmMultiSelect.ShowModal = mrOk ) then
    begin
      aQueryParam.Param1 := frmMultiSelect.SelectedValue;
      txtChargeItem.Text := frmMultiSelect.SelectedText;
    end;
  finally
    frmMultiSelect.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB04.txtRealDateStExit(Sender: TObject);
begin
  if ( txtRealDateSt.Text <> EmptyStr ) then
  begin
    txtRealDateEd.Date := txtRealDateSt.Date;
    //#6001 ?????a?J?????~?? By Kin 2011/07/13
    txtYearMonth.Text := GetDefaultInvDate( txtRealDateSt.Date);

  end;

end;

{ ---------------------------------------------------------------------------- }

function TfrmInvB04.IsDataOK(const aType: Integer): Boolean;
begin
  Result := False;
  if ( aType = 1 ) then
  begin
    if ( Trim( txtRealDateSt.Text ) = EmptyStr ) or
       ( Trim( txtRealDateEd.Text ) = EmptyStr ) then
    begin
      WarningMsg( '???????J????????(?_)~(??) ?C' );
      if ( Trim( txtRealDateSt.Text ) = EmptyStr ) then
      begin
        if ( txtRealDateSt.CanFocus ) then txtRealDateSt.SetFocus;
      end else
      if ( Trim( txtRealDateEd.Text ) = EmptyStr ) then
      begin
        if ( txtRealDateEd.CanFocus ) then txtRealDateEd.SetFocus;
      end;
      Exit;
    end;
  end else
  if ( aType = 2 ) then
  begin
    if Trim( txtYearMonth.Text ) = EmptyStr then
    begin
      WarningMsg( '???????????i?????~???j???????J?C',  );
      if ( txtYearMonth.CanFocus ) then txtYearMonth.SetFocus;
      Exit;
    end;
    if Trim( txtPaperDate.Text ) = EmptyStr then
    begin
      WarningMsg( '???????????i?????????j???????J?C',  );
      if ( txtPaperDate.CanFocus ) then txtPaperDate.SetFocus;
      Exit;
    end;
  end;
  Result := True;  
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvB04.GetSqlText: String;
var
  aServiceTypeParam: TComboBoxItemParam;
  aMemberFlag: String;
  aSQL : String;

begin

  aMemberFlag := '0';
  { #4222 ?W?[?I???????P?A???O By Kin 2009/03/10 }
  if ( rdoMemberY.Checked ) then aMemberFlag := '1';
  Result := Format(
    ' select a.billno, a.item, a.note,                       ' +
    '        a.custid, b.custName, a.citemcode, a.citemname, ' +
    '        a.realdate, a.realamt, a.guino,                 ' +
    '        a.ServiceType, a.PTCode, a.PTName               ' +
    '   from so034 a, so001 b, cd019 c                       ' +
    '  where a.custid = b.custid                             ' +
    '    and a.citemcode = c.codeno                          ' +
    '    and a.PreInvoice = ''3''                            ' +
    '    and a.guino is not null                             ' +
    '    and c.memberflag = ''%s''                           ',
    [aMemberFlag] );

  aSQL := Format( 'select a.custid,b.custName,                 ' +
                ' max(a.realdate) realdate,                    ' +
                ' sum(a.realamt) realamt,a.guino               ' +
                ' from so034 a, so001 b, cd019 c               ' +
                ' where a.custid = b.custid                    ' +
                ' and a.citemcode = c.codeno                   ' +
                ' and a.PreInvoice = ''3''                     ' +
                ' and a.guino is not null                      ' +
                ' and c.memberflag = ''%s''                    ' ,
                [aMemberFlag]);
  {}
  {
  if GetCxComboBoxItemValue( cmbServiceType, aServiceTypeParam ) then
  begin
    Result := ( Result + Format( ' and a.servicetype = ''%s'' ',
      [aServiceTypeParam.KeyValue] ) );
  end;
  }
  {}
  if Trim( txtRealDateSt.Text ) <> EmptyStr then
  begin
    Result := ( Result + Format( ' and a.realdate >= to_date( ''%s'', ''YYYYMMDD'' ) ',
      [FormatDateTime( 'yyyymmdd', txtRealDateSt.Date )] ) );
    aSQL := ( aSQL + Format( ' and a.realdate >= to_date( ''%s'', ''YYYYMMDD'' ) ',
      [FormatDateTime( 'yyyymmdd', txtRealDateSt.Date )] ) );
    FQryWhere := Format( '  realdate >= to_date( ''%s'', ''YYYYMMDD'' ) ',
      [FormatDateTime( 'yyyymmdd', txtRealDateSt.Date )] ) ;
  end;
  {}
  if Trim( txtRealDateEd.Text ) <> EmptyStr then
  begin
    Result := ( Result + Format( ' and a.realdate <= to_date( ''%s'', ''YYYYMMDD'' ) ',
      [FormatDateTime( 'yyyymmdd', txtRealDateEd.Date )] ) );
    aSQL := ( aSQL + Format( ' and a.realdate <= to_date( ''%s'', ''YYYYMMDD'' ) ',
      [FormatDateTime( 'yyyymmdd', txtRealDateEd.Date )] ) );

    FQryWhere := FQryWhere + Format( ' and realdate <= to_date( ''%s'', ''YYYYMMDD'' ) ',
      [FormatDateTime( 'yyyymmdd', txtRealDateEd.Date )] );

  end;
  {}
  {
  if ( aQueryParam.Param1 <> EmptyStr ) then
  begin
    Result := ( Result + Format( ' and a.citemcode in ( %s ) ',
      [aQueryParam.Param1] ) );
  end;
  }
  if ( aQueryParam.Param1 )<> EmptyStr then
  begin
    aSQL := ( aSQL + Format( ' and a.servicetype in ( %s ) ',
      [aQueryParam.Param1] ) );
    FQryWhere := FQryWhere + Format( ' and servicetype in ( %s ) ',
      [aQueryParam.Param1] );
  end;
  aSQL := ( aSQL + ' group by a.custid,b.custName,a.guino ');
  aSQL := ( 'select * from (' + aSQL + ') a ' );
  Result := ( aSQL + ' order by a.realdate,a.custid ' );
  //Result := ( Result + ' order by a.realdate, a.custid ' );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB04.DoFillSO034DataSet;
var
  aSource: TADOQuery;
  aReal, aSale, aTax, aRecNo: Integer;
begin
  aSale := 0;
  aTax := 0;
  aSource := dtmSO.adoComm;
  aSource.Close;
  aSource.SQL.Text := GetSqlText;
  aSource.Open;
  aSource.First;
  {}
  PBar.Properties.Min := 0;
  PBar.Properties.Max := aSource.RecordCount;
  {}
  PBar.Position := 0;
  try
    cdsSO034.DisableControls;
    try
      FKeepInvId := EmptyStr;
      aRecNo := 0;
      while not aSource.Eof do
      begin
        Inc( aRecNo );
        PBar.Position := ( PBar.Position + 1 );
        Application.ProcessMessages;
        cdsSO034.Append;
        cdsSO034.FieldByName( 'RECNO' ).AsInteger := aRecNo;
        cdsSO034.FieldByName( 'SELECTED' ).AsBoolean := False;
        cdsSO034.FieldByName( 'STATUS' ).AsInteger := 0;
//        cdsSO034.FieldByName( 'BILLNO' ).AsString := aSource.FieldByName( 'BILLNO' ).AsString;
//        cdsSO034.FieldByName( 'ITEM' ).AsString := aSource.FieldByName( 'ITEM' ).AsString;
//        cdsSO034.FieldByName( 'NOTE' ).AsString := aSource.FieldByName( 'NOTE' ).AsString;
        cdsSO034.FieldByName( 'CUSTID' ).AsString := aSource.FieldByName( 'CUSTID' ).AsString;
        cdsSO034.FieldByName( 'CUSTNAME' ).AsString := aSource.FieldByName( 'CUSTNAME' ).AsString;
//        cdsSO034.FieldByName( 'CITEMCODE' ).AsString := aSource.FieldByName( 'CITEMCODE' ).AsString;
//        cdsSO034.FieldByName( 'CITEMNAME' ).AsString := aSource.FieldByName( 'CITEMNAME' ).AsString;
        if not VarIsNull( aSource.FieldByName( 'REALDATE' ).Value ) then
          cdsSO034.FieldByName( 'REALDATE' ).AsString := FormatDateTime( 'yyyy/mm/dd', aSource.FieldByName( 'REALDATE' ).AsDateTime );
        cdsSO034.FieldByName( 'REALAMT' ).AsInteger := aSource.FieldByName( 'REALAMT' ).AsInteger;
        cdsSO034.FieldByName( 'GUINO' ).AsString := aSource.FieldByName( 'GUINO' ).AsString;
        {}
       cdsSO034.FieldByName( 'ISOBSOLETE' ).AsString := VarToStrDef( DoGetInvInformation(
         cdsSO034.FieldByName( 'GUINO' ).AsString, 'ISOBSOLETE' ), EmptyStr );
       {}
       cdsSO034.FieldByName( 'BUSINESSID' ).AsString := VarToStrDef( DoGetInvInformation(
         cdsSO034.FieldByName( 'GUINO' ).AsString, 'BUSINESSID' ), EmptyStr );
       {}
       cdsSO034.FieldByName( 'INVFORMAT' ).AsString := VarToStrDef( DoGetInvInformation(
         cdsSO034.FieldByName( 'GUINO' ).AsString, 'INVFORMAT' ), EmptyStr );
       {}
       cdsSO034.FieldByName( 'INVDATE' ).Value := DoGetInvInformation(
         cdsSO034.FieldByName( 'GUINO' ).AsString, 'INVDATE' );
       {}
       cdsSO034.FieldByName( 'TAXTYPE' ).AsString := VarToStrDef( DoGetInvInformation(
         cdsSO034.FieldByName( 'GUINO' ).AsString, 'TAXTYPE' ), EmptyStr );
       {}
       cdsSO034.FieldByName( 'INVAMT' ).AsString := VarToStrDef( DoGetInvInformation(
         cdsSO034.FieldByName( 'GUINO' ).AsString,  'INVAMOUNT' ), '0' );
       {}
        cdsSO034.FieldByName( 'ISFOUND' ).AsBoolean :=
          not VarIsNull( cdsSO034.FieldByName( 'INVDATE' ).Value );
        {}
        aReal := cdsSO034.FieldByName( 'REALAMT' ).AsInteger;
        CalcAmt( cdsSO034.FieldByName( 'TAXTYPE' ).AsString,
          aReal, aSale, aTax );
        if ( aReal <> cdsSO034.FieldByName( 'REALAMT' ).AsInteger ) then
          cdsSO034.FieldByName( 'REALAMT' ).AsInteger := aReal;
        cdsSO034.FieldByName( 'SALEAMT' ).AsInteger := aSale;
        cdsSO034.FieldByName( 'TAXAMT' ).AsInteger := aTax;
        {}
        { #4222 ?W?[?I???????PSERVICETYPE By Kin 2009/03/10 }
//        cdsSO034.FieldByName( 'SERVICETYPE' ).AsString := aSource.FieldByName( 'SERVICETYPE' ).AsString;
//        cdsSO034.FieldByName( 'PTCode' ).AsString := aSource.FieldByName( 'PTCode' ).AsString;
//        cdsSO034.FieldByName( 'PTName' ).AsString := aSource.FieldByName( 'PTName' ).AsString;
        {}
        cdsSO034.Post;
        aSource.Next;
        Application.ProcessMessages;
      end;
      cdsSO034.First;
      FKeepInvId := EmptyStr;
    finally
      cdsSO034.EnableControls;
    end;
  finally
    PBar.Position := 0;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB04.btnQueryClick(Sender: TObject);
begin
  if not IsDataOK( 1 ) then Exit;
  Screen.Cursor := crSQLWait;
  try
    btnSelect.Enabled := False;
    btnRemove.Enabled := False;
    {}
    btnQuery.Enabled := False;
    btnExec.Enabled := False;
    {}
    try
      UnPrepareSO034DataSet;
      Application.ProcessMessages;
      DoFillSO034DataSet;
    finally
      btnQuery.Enabled := True;
      btnExec.Enabled := True;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
  if ( cdsSO034.IsEmpty ) then
  begin
    WarningMsg( '?????d???L?????i???????????C' );
  end else
  begin
    btnSelect.Enabled := True;
    btnRemove.Enabled := True;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB04.txtRealDateStPropertiesValidate(Sender: TObject;
  var DisplayValue: Variant; var ErrorText: TCaption; var Error: Boolean);
begin
  if Error then
  begin
    WarningMsg( '?z???J???????????T?C' );
    ErrorText := EmptyStr;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB04.btnExecClick(Sender: TObject);
var
  aError: Boolean;
begin
  if not IsDataOK( 2 ) then Exit;
  Screen.Cursor := crSQLWait;
  try
    PBar.Properties.Min := 0;
    PBar.Properties.Max := cdsSO034.RecordCount;
    PBar.Position := 0;
    btnQuery.Enabled := False;
    btnExec.Enabled := False;
    btnClose.Enabled := False;
    try
      try
        FYearMonth := TrimChar( txtYearMonth.Text, ['/'] );
        FPaperDate := txtPaperDate.Date;
        CreateAllowance( aError );
      finally
        PBar.Position := 0;
      end;
    finally
      btnQuery.Enabled := True;
      btnExec.Enabled := True;
      btnClose.Enabled := True;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
  { ?X???`?? }
  if ( aError ) then
  begin
    Screen.Cursor := crSQLWait;
    try
      DoErrorReport;
    finally
      Screen.Cursor := crDefault;
    end;  
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB04.CreateAllowance(var aError: Boolean);
var
  aErrMsg, aAllowNoSt, aAllowNoEd: String;
  aRecords: Integer;
  aRcdCnt: Integer;
  procedure GetSysDate;
  begin
    adoSysDate.Close;
    adoSysDate.SQL.Text := 'SELECT SYSDATE FROM DUAL';
    adoSysDate.Open;
    FsysDate := adoSysDate.FieldList[0].AsDateTime;
    adoSysDate.Close;
  end;
begin
  aError := False;
  aRecords := 0;
  aRcdCnt := 0;
  aAllowNoSt := EmptyStr;
  aAllowNoEd := EmptyStr;
//  aDataSet := dtmMain.adoComm;
  GetSysDate;
  cdsSO034.First;
  while not cdsSO034.Eof do
  begin
    { OK ???????b?B?z?@??, ?????????]?????B?z }
    if ( cdsSO034.FieldByName( 'STATUS' ).AsInteger = 1 ) or
       ( cdsSO034.FieldByName( 'SELECTED' ).AsBoolean = False ) then
    begin
      PBar.Position := ( PBar.Position + 1 );
      cdsSO034.Next;
      Application.ProcessMessages;
      Continue;
    end;
    dtmMain.InvConnection.BeginTrans;
    dtmSO.SOConnection.BeginTrans;
    aErrMsg := EmptyStr;
    try
      if not DoInsertInv014( aErrMsg,aRcdCnt,aRecords ) then
      begin
        raise Exception.Create( aErrMsg );
      end;
      if not DoUpdateSO034( aErrMsg ) then
      begin
        raise Exception.Create( aErrMsg );
      end;
      dtmMain.InvConnection.CommitTrans;
      dtmSO.SOConnection.CommitTrans;
    except
      on E: Exception do
      begin
        aError := True;
        aErrMsg := E.Message;
        dtmMain.InvConnection.RollbackTrans;
        dtmSO.SOConnection.RollbackTrans;
      end;
    end;
    cdsSO034.Edit;
    if ( aErrMsg = EmptyStr ) then
    begin
      cdsSO034.FieldByName( 'STATUS' ).AsInteger := 1;
      cdsSO034.FieldByName( 'ERRMSG' ).AsString := EmptyStr;
      { ???????????? }
      { ?p??????????????PM,?]?????i???o?????D???????X???O?bINV014????}
      //Inc( aRecords );
      { ???????_?l???X }
      if ( aAllowNoSt = EmptyStr ) then
        aAllowNoSt := cdsSO034.FieldByName( 'allowanceno' ).AsString;
      { ?????????????X }
      aAllowNoEd := cdsSO034.FieldByName( 'allowanceno' ).AsString;
    end else
    begin
      cdsSO034.FieldByName( 'STATUS' ).AsInteger := 2;
      cdsSO034.FieldByName( 'ERRMSG' ).AsString := aErrMsg;
    end;
    cdsSO034.Post;
    cdsSO034.Next;
    Application.ProcessMessages;
    PBar.Position := ( PBar.Position + 1 );
    Application.ProcessMessages;
  end;
  { ?????T?? }
  if ( aRecords > 0 ) then
  begin
    InfoMsg( format('????????????????????: %d??'#13#10'????????: %d??" ',
              [aRecords,aRcdCnt]));
    {
    InfoMsg( Format( '?????????????? %d ??'#13#10'????????:%s~%s?C',
      [aRecords, aAllowNoSt, aAllowNoEd] ) );
    }
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvB04.DoUpdateSO034(var aErrMsg: String): Boolean;
var
  aChkSO034 : TADOQuery;
  aNote: String;
begin
  aErrMsg := EmptyStr;
  aChkSO034 := dtmMain.adoComm3;
  aChkSO034.Close;
  aChkSO034.SQL.Text := Format('SELECT BILLNO,ITEM,NOTE,CUSTID FROM %s.SO034   ' +
                    ' WHERE %s AND GUINO = ''%s'' and preinvoice = 3           ',
                    [dtmMain.getMisDbOwner, FQryWhere,
                     cdsSO034.FieldByName( 'GUINO' ).AsString] );
  {
  if FDouble then
  begin
    aChkSO034.SQL.Text := aChkSO034.SQL.Text +
        Format(' AND BILLNO || ITEM NOT IN ( SELECT PaperNo || Seq FROM INV014A  ' +
            ' WHERE AllowanceNo = ''%s'')           ',
            [cdsSO034.FieldByName( 'allowanceno' ).AsString]);
  end;
  }
  {}
  aChkSO034.Open;
  if not aChkSO034.IsEmpty then aChkSO034.First;
  while not aChkSO034.Eof do
  begin
    aNote := aChkSO034.FieldByName( 'note' ).AsString;
    if ( aNote <> EmptyStr ) then aNote := ( aNote + ',' );
    aNote := aNote + Format( '????????:%s',
        [cdsSO034.FieldByName( 'allowanceno' ).AsString] );
    try
      dtmSO.CreateAllowance( aChkSO034.FieldByName( 'billno' ).AsString,
      aChkSO034.FieldByName( 'item' ).AsString,
      aChkSO034.FieldByName( 'custid' ).AsString,
      aNote );
    except
      on E: Exception do aErrMsg := E.Message;
    end;
    aChkSO034.Next;
  end;
  aChkSO034.Close;
  Result := ( aErrMsg = EmptyStr );
  {
  aNote := cdsSO034.FieldByName( 'note' ).AsString;
  if ( aNote <> EmptyStr ) then aNote := ( aNote + ',' );
  aNote := aNote + Format( '????????:%s',
    [cdsSO034.FieldByName( 'allowanceno' ).AsString] );
  }
  {}
  {
  try
    dtmSO.CreateAllowance( cdsSO034.FieldByName( 'billno' ).AsString,
      cdsSO034.FieldByName( 'item' ).AsString,
      cdsSO034.FieldByName( 'custid' ).AsString,
      aNote );
  except
    on E: Exception do aErrMsg := E.Message;
  end;
  Result := ( aErrMsg = EmptyStr );
  }
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvB04.DoInsertInv014(var aErrMsg: String;var aRcdCnt,aDetailCnt: Integer): Boolean;
var
  aWriter, aChkDouble : TADOQuery;
  aAllowanceNo, aInvDate: String;
  aRealAmt, aSaleAmt, aTaxAmt: Integer;
  aNewData : Boolean;
begin
  aWriter := dtmMain.adoComm2;
  aChkDouble := dtmMain.adoComm3;

  FDouble := False;
  {}
  aErrMsg := EmptyStr;
  {}
  try
    if ( not cdsSO034.FieldByName( 'ISFOUND' ).AsBoolean ) then
    begin
      raise Exception.Create( '?o???t???d?L?????o?????X?C' );
    end;
    if ( cdsSO034.FieldByName( 'ISOBSOLETE' ).AsString = 'Y' ) then
    begin
      raise Exception.Create( '?????o?????X?w?@?o?C' );
    end;
    if ( cdsSO034.FieldByName( 'REALAMT' ).AsInteger >
         cdsSO034.FieldByName( 'INVAMT' ).AsInteger ) then
    begin
      raise Exception.Create( '???????B > ???o???}?????B?C' );
    end;
    {}
    aInvDate := 'null';
    if not VarIsNull( cdsSO034.FieldByName( 'INVDATE' ).Value ) then
      aInvDate := QuotedStr( FormatDateTime( 'yyyymmdd',
        cdsSO034.FieldByName( 'invdate' ).AsDateTime ) );
    {}
    aRealAmt := cdsSO034.FieldByName( 'REALAMT' ).AsInteger;
    aSaleAmt := cdsSO034.FieldByName( 'SALEAMT' ).AsInteger;
    aTaxAmt := cdsSO034.FieldByName( 'TAXAMT' ).AsInteger;
    {}
    {}
    aChkDouble.Close;
    //#6001 ?p?GUPLOADFlag = 0 ???B PAPERDATE = ?e???????? ?N??????,????upd By Kin 2011/07/13
    aChkDouble.SQL.Text := Format(
      'SELECT ALLOWANCENO,                              ' +
      ' PAPERDATE, UPLOADFLAG                           ' +
      ' From INV014                                     ' +
      ' WHERE IDENTIFYID1 = ''%s''                      ' +
      ' AND IDENTIFYID2 = ''%s''                        ' +
      ' AND INVID = ''%s''                              ' +
      ' AND NVL(UPLOADFLAG,0) = 0                       ' +
      ' AND PAPERDATE = TO_DATE(''%s'',''yyyy/mm/dd'')  ',
      [ IDENTIFYID1,IDENTIFYID2,
      cdsSO034.FieldByName( 'GUINO').AsString,
      txtPaperDate.Text ]);
    aChkDouble.Open;
    {
    if ( not aChkDouble.IsEmpty ) then
    begin
      aNewData := False;
      aChkDouble.First;
    end;
    while not aChkDouble.Eof do
    begin
      if ( aChkDouble.FieldByName( 'UPLOADFLAG' ).AsString = '1' ) or
          ( aChkDouble.FieldByName( 'PAPERDATE' ).AsDateTime <> txtPaperDate.Date ) then
      begin
      
      end;
      aChkDouble.Next;
    end;
    }
    if ( not aChkDouble.IsEmpty )then
    begin
      FDouble := True;
      aAllowanceNo := aChkDouble.FieldByName( 'ALLOWANCENO' ).AsString;
      aWriter.Close;
      aWriter.SQL.Text := Format(
        ' SELECT * FROM INV014        ' +
        ' WHERE IDENTIFYID1 = ''%s''  ' +
        ' AND IDENTIFYID2 = ''%s''    ' +
        ' AND INVID = ''%s''          ' +
        ' AND ALLOWANCENO = ''%s''    ',
        [IDENTIFYID1,IDENTIFYID2,
          cdsSO034.FieldByName( 'GUINO' ).AsString,
          aAllowanceNo ]);
      aWriter.Open;
      aWriter.Edit;
      aWriter.FieldByName( 'InvAmount' ).AsInteger := aWriter.FieldByName( 'InvAmount' ).AsInteger +
        ( cdsSO034.FieldByName( 'RealAmt' ).AsInteger * -1);
      if ( cdsSO034.FieldByName( 'TAXTYPE' ).AsString = '1' ) then
        aWriter.FieldByName( 'SaleAmount' ).AsInteger := Round( ( aWriter.FieldByName( 'InvAmount').AsFloat / 1.05 ) )
      else
        aWriter.FieldByName( 'SaleAmount' ).AsInteger := aWriter.FieldByName( 'InvAmount' ).AsInteger;
      aWriter.FieldByName( 'TaxAmount' ).AsInteger := aWriter.FieldByName( 'InvAmount' ).AsInteger -
        aWriter.FieldByName( 'SaleAmount' ).AsInteger;
      aWriter.FieldByName( 'UPTTIME' ).AsDateTime := FsysDate;
      aWriter.FieldByName( 'UPTEN' ).AsString := dtmMain.getLoginUser;
      aWriter.Post;
      aWriter.Close;
      aChkDouble.Close;
    end else
    begin
      aChkDouble.Close;
      aAllowanceNo := dtmMainJ.GetAllowanceNo;
      {}
      aWriter.Close;

      inc(aRcdCnt);
      {
      aWriter.SQL.Text := Format(
        ' insert into inv014 ( identifyid1, identifyid2, compid,      ' +
        '    custid, paperno, paperdate, businessid, yearmonth,       ' +
        '    invid, seq, invformat, invdate, taxtype, saleamount,     ' +
        '    taxamount, invamount, upttime, upten, allowanceno,       ' +
        '    source )                                                 ' +
        ' values ( ''%s'', ''%s'', ''%s'',                                         ' +
        '    ''%s'', ''%s'', to_date( ''%s'', ''YYYYMMDD'' ), ''%s'', ''%s'',      ' +
        '    ''%s'', ''%s'', ''%s'', to_date( %s, ''YYYYMMDD'' ), ''%s'', ''%d'',  ' +
        '    ''%d'', ''%d'', sysdate, ''%s'', ''%s'',                 ' +
        '    1 )                                                      ',
        [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID,
         cdsSO034.FieldByName( 'CUSTID' ).AsString,
         cdsSO034.FieldByName( 'BILLNO' ).AsString,
         FormatDateTime( 'yyyymmdd', FPaperDate ),
         cdsSO034.FieldByName( 'BUSINESSID' ).AsString,
         FYearMonth,
         cdsSO034.FieldByName( 'GUINO' ).AsString,
         cdsSO034.FieldByName( 'ITEM' ).AsString,
         cdsSO034.FieldByName( 'INVFORMAT' ).AsString,
         aInvDate,
         cdsSO034.FieldByName( 'TAXTYPE' ).AsString,
         aSaleAmt, aTaxAmt, aRealAmt, dtmMain.getLoginUser,
         aAllowanceNo] );
         }
         aWriter.SQL.Text := Format(
        ' insert into inv014 ( identifyid1, identifyid2, compid,                    ' +
        '    custid, paperdate, businessid, yearmonth,                              ' +
        '    invid, invformat, invdate, taxtype, saleamount,                        ' +
        '    taxamount, invamount, upttime, upten, allowanceno,                     ' +
        '    source,seq )                                                           ' +
        ' values ( ''%s'', ''%s'', ''%s'',                                          ' +
        '    ''%s'',  to_date( ''%s'', ''YYYYMMDD'' ), ''%s'', ''%s'',              ' +
        '    ''%s'', ''%s'',  to_date( %s, ''YYYYMMDD'' ), ''%s'', ''%d'',          ' +
        '    ''%d'', ''%d'', to_date(''%s'',''YYYY/MM/DD HH24:MI:SS''),             ' +
        '    ''%s'', ''%s'',                                                        ' +
        '    1,0 )                                                                  ',
        [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID,
         cdsSO034.FieldByName( 'CUSTID' ).AsString,
         FormatDateTime( 'yyyymmdd', FPaperDate ),
         cdsSO034.FieldByName( 'BUSINESSID' ).AsString,
         FYearMonth,
         cdsSO034.FieldByName( 'GUINO' ).AsString,
         cdsSO034.FieldByName( 'INVFORMAT' ).AsString,
         aInvDate,
         cdsSO034.FieldByName( 'TAXTYPE' ).AsString,
         (aSaleAmt * -1 ), (aTaxAmt * -1) , (aRealAmt * -1),
         FormatDateTime('YYYYMMDD HHMMSS',FsysDate),
         dtmMain.getLoginUser,
         aAllowanceNo] );
       aWriter.ExecSQL;
      end;
      aChkDouble.Close;
      aChkDouble.SQL.Text := Format( 'SELECT * FROM %s.SO034         ' +
                      ' WHERE  %s AND GUINO = ''%s''                 ',
                      [ dtmMain.getMisDbOwner, FQryWhere,
                      cdsSO034.FieldByName( 'GUINO' ).AsString ] );
      if FDouble then
      begin
        aChkDouble.SQL.Text := ( aChkDouble.SQL.Text +
            format( ' AND BillNo || Item Not in(select  PaperNo || Seq from inv014a  ' +
                  ' where AllowanceNo = ''%s'')           ',
                  [aAllowanceNo]) );
      end;
      aChkDouble.Open;
      while not aChkDouble.eof do
      begin
        aWriter.Close;
        aWriter.SQL.Text := Format(
        'INSERT INTO INV014A ( ServiceTypeID,PaperNo,          ' +
        ' Seq,ItemId,Description,PayTypeId,PayTypeDesc,        ' +
        ' InvAmount,UptTime,UptEn,AllowanceNo )                ' +
        ' Values ( ''%s'',''%s'',''%s'',''%s'',                ' +
        ' ''%s'',''%s'',''%s'',%d,                             ' +
        ' to_date(''%s'',''YYYY/MM/DD HH24:MI:SS''),           ' +
        ' ''%s'',''%s'' )                                      ' ,
        [ aChkDouble.FieldByName('SERVICETYPE').AsString,
          aChkDouble.FieldByName('BillNO').AsString,
          aChkDouble.FieldByName('ITEM').AsString,
          aChkDouble.FieldByName('CitemCode').AsString,
          aChkDouble.FieldByName('CitemName').AsString,
          aChkDouble.FieldByName('PTCode').AsString,
          aChkDouble.FieldByName('PTName').AsString,
          (aChkDouble.FieldByName('RealAMT').AsInteger * -1),
          FormatDateTime('YYYYMMDD HHMMSS',FsysDate),
          dtmMain.getLoginUser,
          aAllowanceNo ] );
        aWriter.ExecSQL;
        inc(aDetailCnt);
        aChkDouble.Next;
      end;
      aChkDouble.Close;



   except
    on E: Exception do aErrMsg := E.Message;
  end;
  {}
  aWriter.Close;
  aChkDouble.Close;
  {}
  Result := ( aErrMsg = EmptyStr );
  if Result then
  begin
    cdsSO034.Edit;
    cdsSO034.FieldByName( 'allowanceno' ).AsString := aAllowanceNo;
    cdsSO034.Post;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB04.rdoMemberYClick(Sender: TObject);
begin
  ZeroMemory( @aQueryParam, SizeOf( aQueryParam ) );
  txtChargeItem.Text := EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB04.PrepareReportDataSet;
begin
  dtmReport.cdsTempory.FieldDefs.Clear;
  dtmReport.cdsTempory.FieldDefs.Add( 'CUSTID', ftInteger );
  dtmReport.cdsTempory.FieldDefs.Add( 'CUSTNAME', ftString, 50 );
  dtmReport.cdsTempory.FieldDefs.Add( 'CITEMCODE', ftInteger );
  dtmReport.cdsTempory.FieldDefs.Add( 'CITEMNAME', ftString, 40 );
  dtmReport.cdsTempory.FieldDefs.Add( 'REALDATE', ftDateTime );
  dtmReport.cdsTempory.FieldDefs.Add( 'REALAMT', ftInteger );
  dtmReport.cdsTempory.FieldDefs.Add( 'SALEAMT', ftInteger );
  dtmReport.cdsTempory.FieldDefs.Add( 'TAXAMT', ftInteger );
  dtmReport.cdsTempory.FieldDefs.Add( 'GUINO', ftString, 20 );
  dtmReport.cdsTempory.FieldDefs.Add( 'ERRMSG', ftString, 1024 );
  dtmReport.cdsTempory.CreateDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB04.UnPrepareReportDataSet;
begin
  if not VarIsNull( dtmReport.cdsTempory.Data ) then
    dtmReport.cdsTempory.Data := Null;
  dtmReport.cdsTempory.FieldDefs.Clear;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB04.DoErrorReport;
var
  aPath: String;
begin
  UnPrepareReportDataSet;
  PrepareReportDataSet;
  cdsSO034.First;
  while not cdsSO034.Eof do
  begin
    if ( cdsSO034.FieldByName( 'STATUS' ).AsString = '2' ) then
    begin
      dtmReport.cdsTempory.Append;
      dtmReport.cdsTempory.FieldByName( 'CUSTID' ).AsString :=
        cdsSO034.FieldByName( 'CUSTID' ).AsString;
      dtmReport.cdsTempory.FieldByName( 'CUSTNAME' ).AsString :=
        cdsSO034.FieldByName( 'CUSTNAME' ).AsString;
      dtmReport.cdsTempory.FieldByName( 'CITEMCODE' ).AsString :=
        cdsSO034.FieldByName( 'CITEMCODE' ).AsString;
      dtmReport.cdsTempory.FieldByName( 'CITEMNAME' ).AsString :=
        cdsSO034.FieldByName( 'CITEMNAME' ).AsString;
      dtmReport.cdsTempory.FieldByName( 'REALDATE' ).AsString :=
        cdsSO034.FieldByName( 'REALDATE' ).AsString;
      dtmReport.cdsTempory.FieldByName( 'REALAMT' ).AsString :=
        cdsSO034.FieldByName( 'REALAMT' ).AsString;
      dtmReport.cdsTempory.FieldByName( 'SALEAMT' ).AsString :=
        cdsSO034.FieldByName( 'SALEAMT' ).AsString;
      dtmReport.cdsTempory.FieldByName( 'TAXAMT' ).AsString :=
        cdsSO034.FieldByName( 'TAXAMT' ).AsString;
      dtmReport.cdsTempory.FieldByName( 'GUINO' ).AsString :=
        cdsSO034.FieldByName( 'GUINO' ).AsString;
      dtmReport.cdsTempory.FieldByName( 'ERRMSG' ).AsString :=
        cdsSO034.FieldByName( 'ERRMSG' ).AsString;
      dtmReport.cdsTempory.Post;
    end;
    cdsSO034.Next;
  end;
  aPath := IncludeTrailingPathDelimiter( ExtractFilePath(
    Application.ExeName ) ) + IncludeTrailingPathDelimiter(
    REPORT_FOLDER ) + 'FrptInvB04_1.fr3';
  {}
  dtmReport.frxMainReport.LoadFromFile( aPath );
  {}
  dtmReport.frxMainReport.Variables.Variables['aOperator'] :=
    QuotedStr( dtmMain.getLoginUser );
  {}
  dtmReport.frxMainReport.ShowReport;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB04.DoSelectRecord(const aSelected: Boolean);
var
  aIndex, aCounts: Integer;
  aKeys: array of Integer;
  aBk: TBookmark;
begin
  aCounts := gvSO034.Controller.SelectedRecordCount;
  SetLength( aKeys, aCounts );
  for aIndex := 0 to aCounts - 1 do
  begin
    aKeys[aIndex] := VarAsType( gvSO034.Controller.SelectedRecords[aIndex].Values[
      gvSO034RECNO.Index], varInteger );
  end;
  cdsSO034.DisableControls;
  try
    aBk := cdsSO034.GetBookmark;
    try
      for aIndex := Low( aKeys ) to High( aKeys ) do
      begin
        if cdsSO034.Locate( 'RECNO', aKeys[aIndex], [] ) then
        begin
          cdsSO034.Edit;
          cdsSO034.FieldByName( 'SELECTED' ).AsBoolean := aSelected;
          cdsSO034.Post;
        end;
      end;
      cdsSO034.GotoBookmark( aBk );
    finally
      cdsSO034.FreeBookmark( aBk );
    end;  
  finally
    cdsSO034.EnableControls;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB04.btnSelectClick(Sender: TObject);
begin
  Screen.Cursor := crSQLWait;
  try
    DoSelectRecord( True );
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB04.btnRemoveClick(Sender: TObject);
begin
  Screen.Cursor := crSQLWait;
  try
    DoSelectRecord( False );
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB04.gvSO034CustomDrawCell(Sender: TcxCustomGridTableView;
  ACanvas: TcxCanvas; AViewInfo: TcxGridTableDataCellViewInfo;
  var ADone: Boolean);
begin
  if ( AViewInfo.GridRecord.Values[gvSO034SELECTED.Index] = 1 ) and
     ( not AViewInfo.GridRecord.Selected ) then
  begin
    ACanvas.Brush.Color := clMoneyGreen;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvB04.DoGetInvInformation(const aInvId: String;
  const aFieldName: String): Variant;
begin
  if ( aInvId <> FKeepInvId ) then
  begin
    FInvReader.Close;
    FInvReader.SQL.Text := Format(
      ' SELECT A.BUSINESSID, A.INVFORMAT,  ' +
      '        A.INVDATE, A.INVAMOUNT,     ' +
      '        A.TAXTYPE, A.ISOBSOLETE     ' +
      '   FROM INV007 A                    ' +
      '  WHERE A.IDENTIFYID1 = ''%s''      ' +
      '    AND A.IDENTIFYID2 = ''%s''      ' +
      '    AND A.INVID = ''%s''            ' +
      '    AND A.COMPID = ''%s''           ',
      [IDENTIFYID1, IDENTIFYID2, aInvId, dtmMain.getCompID] );
    FInvReader.Open;
    FKeepInvId := aInvId;
  end;
  Result := FInvReader.FieldByName( aFieldName ).Value;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB04.CalcAmt(const aTaxType: String; var aReal, aSale, aTax: Integer);
 var blnMinus : Boolean;
begin
//  if ( aReal < 0 ) then aReal := Abs( aReal );
  //#4222 ??????OK ?????n??????,?t???n???t?? By Kin 2009/09/01
//  aReal := aReal ;
  blnMinus := False;
  if aReal < 0 then blnMinus := True;
  aReal := Abs(aReal);
  aSale := aReal;
  aTax := 0;
  if ( aTaxType = '1' ) then
  begin
    aSale := Trunc( RoundTo( ( aReal / 1.05 ), 0 ) );
    aTax := ( aReal - aSale );
  end;
  if blnMinus then
  begin
    aReal := aReal * -1;
    aSale := aSale * -1;
    aTax := aTax * -1;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB04.FillChargeItemDataSet;
var
  aSource: TADOQuery;
  aMemberFlag: String;
  aSQL : String;
begin
   aSQL := Format(' select itemid, description from inv022 ' +
                  '  where identifyid1 = ''%s''            ' +
                  '    and identifyid2 = ''%s''            ' ,
                  [IDENTIFYID1,IDENTIFYID2]);

  if ( FInv005.FieldDefs.Count <= 0 ) then
  begin
    FInv005.FieldDefs.Add( 'ITEMID', ftString,5 );
    FInv005.FieldDefs.Add( 'DESCRIPTION', ftString, 40 );
    FInv005.CreateDataSet;
  end;
  if not VarIsNull( FInv005.Data ) then
    FInv005.EmptyDataSet;
  FInvReader.Close;
  { #4222 ???\?????]???s?????H?N SIGN = '-' ???????? By Kin 2009/03/09 }
  {
  FInvReader.SQL.Text := Format(
    ' SELECT A.ITEMID, A.DESCRIPTION   ' +
    '   FROM INV005 A                  ' +
    '  WHERE A.IDENTIFYID1 = ''%s''    ' +
    '    AND A.IDENTIFYID2 = ''%s''    ' +
    '    AND A.COMPID = ''%s''         ' +
    '  ORDER BY A.ITEMID               ',
    [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID] );
  }
  FInvReader.SQL.Text := aSQL;
  FInvReader.Open;
  FInvReader.First;
  while not FInvReader.Eof do
  begin
    FInv005.Append;
    FInv005.FieldByName( 'ITEMID' ).AsString :=
      FInvReader.FieldByName( 'ITEMID' ).AsString;
    FInv005.FieldByName( 'DESCRIPTION' ).AsString :=
      FInvReader.FieldByName( 'DESCRIPTION' ).AsString;
    FInv005.Post;
    FInvReader.Next;
  end;
  //FInvReader.Close;
  {}
  //aMemberFlag := '0';
  {
  if ( rdoMemberY.Checked ) then aMemberFlag := '1';

  aSource := dtmSO.adoComm;
  aSource.Close;
  aSource.SQL.Text := Format(
    '  SELECT A.CODENO FROM CD019 A    ' +
    '   WHERE A.MEMBERFLAG = ''%s''    ' +
    '   ORDER BY A.CODENO              ', [aMemberFlag] );
  aSource.Open;
  }
  {}
  {
  FInv005.First;
  while not FInv005.Eof do
  begin
    if not aSource.Locate( 'CODENO', FInv005.FieldByName( 'ITEMID' ).AsString, [] ) then
      FInv005.Delete
    else
      FInv005.Next;
  end;
  aSource.Close;
  aSource := nil;
  }
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvB04.GetDefaultInvDate(const aDate: TDateTime): String;
var
  aYear,aMonth :String;
begin

  aYear := FormatDateTime('yyyy',aDate);
  aMonth := FormatDateTime('MM',aDate);
  if StrToInt(aMonth) >= 2 then
  begin

    if ( StrToInt( aMonth ) mod 2 ) = 0 then
    begin
      aMonth := IntToStr( StrToInt( aMonth ) - 1 )
    end else
      aMonth := IntToStr( StrToInt( aMonth ) - 0 );

    if Length( aMonth ) = 1 then
      aMonth := '0' + aMonth;
  end else
  begin
//    aMonth := '11';
//    aYear := IntToStr( StrToInt( aYear ) - 1 );
  end;
  Result :=  aYear + '/' + aMonth ;

end;

end.
