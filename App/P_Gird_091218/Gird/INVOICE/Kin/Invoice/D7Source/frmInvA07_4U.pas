unit frmInvA07_4U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, ExtCtrls, StdCtrls,
  cxLookAndFeelPainters, cxGraphics, cxStyles,
  cxCustomData, cxFilter, cxData, cxDataStorage, cxEdit, DB, cxDBData,
  cxControls, cxGridCustomView, cxGridCustomTableView, cxGridTableView,
  cxGridBandedTableView, cxGridDBBandedTableView, cxClasses, cxGridLevel,
  cxGrid, cxSplitter, dxStatusBar, cxButtons,
  cxRadioGroup, cxContainer, cxLabel, cxDBLabel, cxDropDownEdit, cxDBEdit,
  cxMaskEdit, cxCalendar, DBClient, cxTextEdit, cxCurrencyEdit, ADODB,
  dxSkinsCore, dxSkinsDefaultPainters, dxSkinscxPCPainter,
  dxSkinsdxStatusBarPainter, dxSkinBlack, dxSkinBlue, dxSkinCaramel,
  dxSkinCoffee;

type

  TDMLMode = ( dmBrowser, dmInsert, dmUpdate, dmSave, dmCancel, dmDelete );

  TfrmInvA07_1 = class(TForm)
    ScrollBox1: TScrollBox;
    btnInsert: TcxButton;
    btnEdit: TcxButton;
    btnCancel: TcxButton;
    btnSave: TcxButton;
    btnQuit: TcxButton;
    btnQuery: TcxButton;
    Bevel1: TBevel;
    dxStatusBar1: TdxStatusBar;
    dsMaster: TDataSource;
    cdsMaster: TClientDataSet;
    btnReport: TcxButton;
    btnDelete: TcxButton;
    Panel1: TPanel;
    ScrollBox2: TScrollBox;
    Shape1: TShape;
    Label1: TLabel;
    Label2: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label4: TcxLabel;
    Label5: TcxLabel;
    Label6: TcxLabel;
    txtPaperNo: TcxDBTextEdit;
    txtPaperDate: TcxDBDateEdit;
    txtYearMonth: TcxDBMaskEdit;
    txtInvId: TcxDBTextEdit;
    lbCustSName: TcxDBLabel;
    txtSaleAmount: TcxDBCurrencyEdit;
    txtTaxAmount: TcxDBCurrencyEdit;
    txtInvAmount: TcxDBCurrencyEdit;
    cxDBLabel1: TcxDBLabel;
    lbInvDate: TcxDBLabel;
    lbTaxTypeDesc: TcxDBLabel;
    lbInvFormatDesc: TcxDBLabel;
    lbCustId: TcxDBLabel;
    lbBusinessId: TcxDBLabel;
    gvGrid: TcxGrid;
    gvView: TcxGridDBBandedTableView;
    gvViewCOMPSNAME: TcxGridDBBandedColumn;
    gvViewCUSTID: TcxGridDBBandedColumn;
    gvViewCUSTSNAME: TcxGridDBBandedColumn;
    gvViewPAPERNO: TcxGridDBBandedColumn;
    gvViewPAPERDATE: TcxGridDBBandedColumn;
    gvViewBUSINESSID: TcxGridDBBandedColumn;
    gvViewYEARMONTH2: TcxGridDBBandedColumn;
    gvViewINVID: TcxGridDBBandedColumn;
    gvViewINVFORMAT: TcxGridDBBandedColumn;
    gvViewINVDATE: TcxGridDBBandedColumn;
    gvViewTAXTYPE: TcxGridDBBandedColumn;
    gvViewSALEAMOUNT: TcxGridDBBandedColumn;
    gvViewTAXAMOUNT: TcxGridDBBandedColumn;
    gvViewINVAMOUNT: TcxGridDBBandedColumn;
    gvViewUPTEN: TcxGridDBBandedColumn;
    gvLevel1: TcxGridLevel;
    cxSplitter1: TcxSplitter;
    Label3: TLabel;
    Label14: TLabel;
    lblAllowanceNo: TcxDBLabel;
    gvViewALLOWANCENO: TcxGridDBBandedColumn;
    gvViewSOURCE: TcxGridDBBandedColumn;
    Label15: TLabel;
    lblSource2: TcxDBLabel;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure cdsMasterAfterScroll(DataSet: TDataSet);
    procedure btnInsertClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnQuitClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure btnQueryClick(Sender: TObject);
    procedure txtPaperDatePropertiesInitPopup(Sender: TObject);
    procedure txtYearMonthPropertiesValidate(Sender: TObject;
      var DisplayValue: Variant; var ErrorText: TCaption;
      var Error: Boolean);
    procedure txtPaperDatePropertiesValidate(Sender: TObject;
      var DisplayValue: Variant; var ErrorText: TCaption;
      var Error: Boolean);
    procedure txtInvIdExit(Sender: TObject);
    procedure cdsMasterNewRecord(DataSet: TDataSet);
    procedure txtSaleAmountExit(Sender: TObject);
    procedure txtTaxAmountExit(Sender: TObject);
    procedure txtInvAmountExit(Sender: TObject);
    procedure txtPaperNoExit(Sender: TObject);
    procedure btnReportClick(Sender: TObject);
    procedure gvViewDblClick(Sender: TObject);
  private
    { Private declarations }
    FMode: TDMLMode;
    FSqlBody: String;
    FSqlWhere: String;
    FSqlOrder: String;
    FInvExists: Boolean;
    FConditionCompText: String;
    procedure PrepareMasterDataSet;
    function FillMasterDataSet(const aCondition: String): Integer;
    function BuildMasterSQL(const aCondition: String): String;
    function BuildQueryRecordCountSQL(const aCondition: String): String;
    function DataSetStateChange(const aMode: TDMLMode): Boolean;
    procedure GridStateChange(const aMode: TDMLMode);
    procedure ButtonStateChange(const aMode: TDMLMode);
    procedure EditorStateChange(const aMode: TDMLMode);
    procedure SetMasterFieldState;
    function ValidateMasterDataSet(var aErrMsg: String): Boolean;
    function ApplyDataSet(var aErrMsg: String): Boolean;
    function ValidateCanModify(var aErrMsg: String): Boolean;
    function ValidateCanDelete(var aErrMsg: String): Boolean;
    function DetectQueryRecordCount(const aCondition: String): Integer;
    function CheckInvoiceIsObsolete(const aInvId, aCompId: String): Boolean;
    function CheckInvoiceIdExists(const aInvId, aCompId: String): Boolean;
    function CheckPaperNoExists(const aPaperNo, aInvId, aCompId: String): Boolean;
    procedure CalcMoney(aKind: Integer);
    function IsAmtOK: Boolean;
    function GetYearMonth: String;
  public
    { Public declarations }
    procedure DMLModeChange(const aMode: TDMLMode);
  end;

var
  frmInvA07_1: TfrmInvA07_1;

implementation

uses cbUtilis, Uotheru, frmMainU, dtmMainU, dtmMainHU, dtmMainJU, dtmSOU,
     frmInvA07_2U, dtmReportModule, frmInvA07_3U ;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

function GetTaxTypeDesc(const aValue: String): String;
begin
  Result := EmptyStr;
  if ( aValue = '1' ) then
    Result := '1.應稅'
  else if ( aValue = '2' ) then
    Result := '2.零稅率'
  else  if ( aValue = '3' ) then
    Result := '3.免稅';
end;

{ ---------------------------------------------------------------------------- }

function GetInvFormatDesc(const aValue: String): String;
begin
  Result := EmptyStr;
  if ( aValue = '1' ) then
    Result := '1.電子'
  else if ( aValue = '2' ) then
    Result := '2.手二'
  else  if ( aValue = '3' ) then
    Result := '3.手三';
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_1.FormCreate(Sender: TObject);
begin
  FMode := dmBrowser;
  PrepareMasterDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_1.FormShow(Sender: TObject);
begin
  Caption := frmMain.GetFormTitleString( 'A07', '銷貨退回與折讓資料輸入' );
  Self.Update;
  Screen.Cursor := crSQLWait;
  try
    Application.ProcessMessages;
    FillMasterDataSet( EmptyStr );
    SetMasterFieldState;
    Application.ProcessMessages;
    DMLModeChange( dmBrowser );
    Application.ProcessMessages;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_1.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  if FMode in [dmInsert, dmUpdate] then
  begin
    CanClose := ConfirmMsg( '結束作業?' );
    if CanClose then
      DMLModeChange( dmCancel );
  end;
  if CanClose then
  begin
    cdsMaster.AfterScroll := nil;
    cdsMaster.ReadOnly := False;
    cdsMaster.EmptyDataSet;
  end;
end;

{ ---------------------------------------------------------------------------- }



{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_1.cdsMasterAfterScroll(DataSet: TDataSet);
begin
  FInvExists := CheckInvoiceIdExists(
    DataSet.FieldByName( 'INVID' ).AsString, DataSet.FieldByName( 'COMPID' ).AsString );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_1.cdsMasterNewRecord(DataSet: TDataSet);
begin
  cdsMaster.FieldByName( 'COMPID' ).AsString := dtmMain.getCompID;
  cdsMaster.FieldByName( 'COMPSNAME' ).AsString := dtmMain.getCompName;
  cdsMaster.FieldByName( 'SEQ' ).AsString := '0';
  cdsMaster.FieldByName( 'SOURCE' ).AsInteger := 0;
  cdsMaster.FieldByName( 'SOURCE2' ).AsString := '自行輸入';
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA07_1.BuildMasterSQL(const aCondition: String): String;
begin
  FSqlBody := Format(
    ' SELECT A.IDENTIFYID1,                                     ' +
    '        A.IDENTIFYID2,                                     ' +
    '        A.COMPID,                                          ' +
    '        B.COMPSNAME,                                       ' +
    '        A.CUSTID,                                          ' +
    '        C.CUSTSNAME,                                       ' +
    '        A.PAPERNO,                                         ' +
    '        A.PAPERDATE,                                       ' +
    '        A.BUSINESSID,                                      ' +
    '        A.YEARMONTH,                                       ' +
    '        DECODE( A.YEARMONTH, NULL, A.YEARMONTH,            ' +
    '                SUBSTR( A.YEARMONTH, 1, 4 ) || ''/'' ||    ' +
    '                SUBSTR( A.YEARMONTH, 5, 2 ) ) YEARMONTH2,  ' +
    '        A.INVID,                                           ' +
    '        A.SEQ,                                             ' +
    '        A.INVFORMAT,                                       ' +
    '        A.INVDATE,                                         ' +
    '        A.TAXTYPE,                                         ' +
    '        A.SALEAMOUNT,                                      ' +
    '        A.TAXAMOUNT,                                       ' +
    '        A.INVAMOUNT,                                       ' +
    '        A.UPTTIME,                                         ' +
    '        A.UPTEN,                                           ' +
    '        A.ALLOWANCENO,                                     ' +
    '        A.SOURCE,                                          ' +
    '        DECODE( NVL( A.SOURCE, 0 ), 0, ''自行輸入'',       ' +
    '                ''批次產生'' ) SOURCE2                     ' +
    '    FROM INV007 C, INV014 A, INV001 B                      ' +
    '   WHERE A.IDENTIFYID1 = B.IDENTIFYID1                     ' +
    '     AND A.IDENTIFYID2 = B.IDENTIFYID2                     ' +
    '     AND A.COMPID = B.COMPID                               ' +
    '     AND A.IDENTIFYID1 = C.IDENTIFYID1                     ' +
    '     AND A.IDENTIFYID2 = C.IDENTIFYID2                     ' +
    '     AND A.INVID = C.INVID                                 ' +
    '     AND A.IDENTIFYID1 = ''%s''                            ' +
    '     AND A.IDENTIFYID2 = ''%s''                            ',
    [IDENTIFYID1, IDENTIFYID2] );
  {}
  FSqlWhere := ' AND 1 = 2 ';
  {}
  FSqlOrder := ' ORDER BY A.COMPID, A.YEARMONTH, A.PAPERNO ';
  {}
  if ( Trim( aCondition ) = EmptyStr ) then
  begin
    Result := ( FSqlBody + FSqlWhere + FSqlOrder );
  end else
  begin
    Result := ( FSqlBody + aCondition + FSqlOrder );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA07_1.BuildQueryRecordCountSQL(const aCondition: String): String;
begin
  Result := Format(
    ' SELECT COUNT(1) COUNTS                                    ' +
    '    FROM INV007 C, INV014 A, INV001 B                      ' +
    '   WHERE A.IDENTIFYID1 = B.IDENTIFYID1                     ' +
    '     AND A.IDENTIFYID2 = B.IDENTIFYID2                     ' +
    '     AND A.COMPID = B.COMPID                               ' +
    '     AND A.IDENTIFYID1 = C.IDENTIFYID1                     ' +
    '     AND A.IDENTIFYID2 = C.IDENTIFYID2                     ' +
    '     AND A.INVID = C.INVID                                 ' +
    '     AND A.IDENTIFYID1 = ''%s''                            ' +
    '     AND A.IDENTIFYID2 = ''%s''                            ',
    [IDENTIFYID1, IDENTIFYID2] );
  Result := ( Result + aCondition );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_1.PrepareMasterDataSet;
begin
  if ( cdsMaster.FieldDefs.Count <= 0 ) then
  begin
    cdsMaster.FieldDefs.Add( 'IDENTIFYID1', ftString, 1 );
    cdsMaster.FieldDefs.Add( 'IDENTIFYID2', ftString, 2 );
    cdsMaster.FieldDefs.Add( 'COMPID', ftString, 6 );
    cdsMaster.FieldDefs.Add( 'COMPSNAME', ftString, 50 );
    cdsMaster.FieldDefs.Add( 'CUSTID', ftString, 8 );
    cdsMaster.FieldDefs.Add( 'CUSTNAME', ftString, 50 );
    cdsMaster.FieldDefs.Add( 'PAPERNO', ftString, 15 );
    cdsMaster.FieldDefs.Add( 'PAPERDATE', ftDate );
    cdsMaster.FieldDefs.Add( 'BUSINESSID', ftString, 8 );
    cdsMaster.FieldDefs.Add( 'YEARMONTH', ftString, 6 );
    cdsMaster.FieldDefs.Add( 'YEARMONTH2', ftString, 7 );
    cdsMaster.FieldDefs.Add( 'INVID', ftString, 10 );
    cdsMaster.FieldDefs.Add( 'SEQ', ftInteger );
    cdsMaster.FieldDefs.Add( 'INVFORMAT', ftString, 1 );
    cdsMaster.FieldDefs.Add( 'INVFORMATDESC', ftString, 10 );
    cdsMaster.FieldDefs.Add( 'INVDATE', ftDate );
    cdsMaster.FieldDefs.Add( 'TAXTYPE', ftString, 1 );
    cdsMaster.FieldDefs.Add( 'TAXTYPEDESC', ftString, 10 );
    cdsMaster.FieldDefs.Add( 'SALEAMOUNT', ftInteger );
    cdsMaster.FieldDefs.Add( 'TAXAMOUNT', ftInteger );
    cdsMaster.FieldDefs.Add( 'INVAMOUNT', ftInteger );
    cdsMaster.FieldDefs.Add( 'UPTTIME', ftDateTime );
    cdsMaster.FieldDefs.Add( 'UPTEN', ftString, 20  );
    cdsMaster.FieldDefs.Add( 'ALLOWANCENO', ftString, 12 );
    cdsMaster.FieldDefs.Add( 'SOURCE', ftInteger );
    cdsMaster.FieldDefs.Add( 'SOURCE2', ftString, 20 );
    cdsMaster.CreateDataSet;
  end;
  cdsMaster.EmptyDataSet;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA07_1.FillMasterDataSet(const aCondition: String): Integer;
begin
  cdsMaster.EmptyDataSet;
  dtmMain.adoComm.Close;
  dtmMain.adoComm.SQL.Text := BuildMasterSQL( aCondition );
  dtmMain.adoComm.Open;
  dtmMain.adoComm.First;
  cdsMaster.ReadOnly := False;
  cdsMaster.DisableControls;
  cdsMaster.AfterScroll := nil;
  try
    while not dtmMain.adoComm.Eof do
    begin
      cdsMaster.Append;
      cdsMaster.FieldByName( 'IDENTIFYID1' ).AsString := dtmMain.adoComm.FieldByName( 'IDENTIFYID1' ).AsString;
      cdsMaster.FieldByName( 'IDENTIFYID2' ).AsString := dtmMain.adoComm.FieldByName( 'IDENTIFYID2' ).AsString;
      cdsMaster.FieldByName( 'COMPID' ).AsString := dtmMain.adoComm.FieldByName( 'COMPID' ).AsString;
      cdsMaster.FieldByName( 'COMPSNAME' ).AsString := dtmMain.adoComm.FieldByName( 'COMPSNAME' ).AsString;
      cdsMaster.FieldByName( 'CUSTID' ).AsString := dtmMain.adoComm.FieldByName( 'CUSTID' ).AsString;
      cdsMaster.FieldByName( 'CUSTNAME' ).AsString := dtmMain.adoComm.FieldByName( 'CUSTSNAME' ).AsString;
      cdsMaster.FieldByName( 'PAPERNO' ).AsString := dtmMain.adoComm.FieldByName( 'PAPERNO' ).AsString;
      cdsMaster.FieldByName( 'PAPERDATE' ).AsString := dtmMain.adoComm.FieldByName( 'PAPERDATE' ).AsString;
      cdsMaster.FieldByName( 'BUSINESSID' ).AsString := dtmMain.adoComm.FieldByName( 'BUSINESSID' ).AsString;
      cdsMaster.FieldByName( 'YEARMONTH' ).AsString := dtmMain.adoComm.FieldByName( 'YEARMONTH' ).AsString;
      cdsMaster.FieldByName( 'YEARMONTH2' ).AsString := dtmMain.adoComm.FieldByName( 'YEARMONTH2' ).AsString;
      cdsMaster.FieldByName( 'INVID' ).AsString := dtmMain.adoComm.FieldByName( 'INVID' ).AsString;
      cdsMaster.FieldByName( 'SEQ' ).AsString := dtmMain.adoComm.FieldByName( 'SEQ' ).AsString;
      cdsMaster.FieldByName( 'INVFORMAT' ).AsString := dtmMain.adoComm.FieldByName( 'INVFORMAT' ).AsString;
      cdsMaster.FieldByName( 'INVDATE' ).AsString := dtmMain.adoComm.FieldByName( 'INVDATE' ).AsString;
      cdsMaster.FieldByName( 'TAXTYPE' ).AsString := dtmMain.adoComm.FieldByName( 'TAXTYPE' ).AsString;
      cdsMaster.FieldByName( 'SALEAMOUNT' ).AsString := dtmMain.adoComm.FieldByName( 'SALEAMOUNT' ).AsString;
      cdsMaster.FieldByName( 'TAXAMOUNT' ).AsString := dtmMain.adoComm.FieldByName( 'TAXAMOUNT' ).AsString;
      cdsMaster.FieldByName( 'INVAMOUNT' ).AsString := dtmMain.adoComm.FieldByName( 'INVAMOUNT' ).AsString;
      cdsMaster.FieldByName( 'UPTTIME' ).AsString := dtmMain.adoComm.FieldByName( 'UPTTIME' ).AsString;
      cdsMaster.FieldByName( 'UPTEN' ).AsString := dtmMain.adoComm.FieldByName( 'UPTEN' ).AsString;
      cdsMaster.FieldByName( 'TAXTYPEDESC' ).AsString := GetTaxTypeDesc( dtmMain.adoComm.FieldByName( 'TAXTYPE' ).AsString );
      cdsMaster.FieldByName( 'INVFORMATDESC' ).AsString := GetInvFormatDesc( dtmMain.adoComm.FieldByName( 'INVFORMAT' ).AsString );
      cdsMaster.FieldByName( 'ALLOWANCENO' ).AsString := dtmMain.adoComm.FieldByName( 'ALLOWANCENO' ).AsString;
      cdsMaster.FieldByName( 'SOURCE' ).AsInteger := dtmMain.adoComm.FieldByName( 'SOURCE' ).AsInteger;
      cdsMaster.FieldByName( 'SOURCE2' ).AsString := dtmMain.adoComm.FieldByName( 'SOURCE2' ).AsString;
      cdsMaster.Post;
      dtmMain.adoComm.Next;
    end;
    cdsMaster.First;
  finally
    cdsMaster.ReadOnly := True;
    cdsMaster.EnableControls;
    cdsMaster.AfterScroll := cdsMasterAfterScroll;
  end;
  dtmMain.adoComm.Close;
  Result := cdsMaster.RecordCount;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_1.DMLModeChange(const aMode: TDMLMode);
begin
  if not DataSetStateChange( aMode ) then Exit;
  FMode := aMode;
  if ( aMode in [dmSave, dmCancel, dmDelete] ) then
    FMode := dmBrowser;
  GridStateChange( FMode );
  ButtonStateChange( FMode );
  EditorStateChange( FMode );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_1.ButtonStateChange(const aMode: TDMLMode);
begin
  case aMode of
    dmBrowser, dmCancel:
      begin
        btnInsert.Enabled := True;
        btnCancel.Enabled := False;
        btnSave.Enabled := False;
        btnQuery.Enabled := True;
        btnEdit.Enabled := True;
        btnDelete.Enabled := True;
        btnReport.Enabled := True;
      end;
    dmInsert, dmUpdate:
      begin
        btnInsert.Enabled := False;
        btnCancel.Enabled := True;
        btnSave.Enabled := True;
        btnQuery.Enabled := False;
        btnEdit.Enabled := False;
        btnDelete.Enabled := False;
        btnReport.Enabled := False;
      end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA07_1.DataSetStateChange(const aMode: TDMLMode): Boolean;
var
  aErrMsg: String;
begin
  Result := True;
  aErrMsg := '';
  case aMode of
    dmBrowser:
      begin
        cdsMaster.ReadOnly := True;
        if gvView.IsUpdateLocked then gvView.EndUpdate;
        cdsMasterAfterScroll( cdsMaster );
      end;
    dmCancel:
      begin
        cdsMaster.AfterScroll := nil;
        try
          cdsMaster.Cancel;
        finally
          cdsMaster.AfterScroll := cdsMasterAfterScroll;
        end;  
        cdsMaster.ReadOnly := True;
        if gvView.IsUpdateLocked then gvView.EndUpdate;
        cdsMasterAfterScroll( cdsMaster );
      end;
    dmInsert:
      begin
        cdsMaster.ReadOnly := False;
        if not gvView.IsUpdateLocked then gvView.BeginUpdate;
        cdsMaster.Append;
      end;
    dmUpdate:
      begin
        cdsMaster.ReadOnly := False;
        if not gvView.IsUpdateLocked then gvView.BeginUpdate;
        cdsMaster.Edit;
      end;
    dmDelete:
      begin
        FMode := aMode;
        try
          Result := ApplyDataSet( aErrMsg );
        finally
          FMode := dmBrowser;
        end;
        if not Result then
        begin
          WarningMsg( aErrMsg );
          Exit;
        end;
        cdsMaster.ReadOnly := False;
        try
          cdsMaster.Delete;
        finally
          cdsMaster.ReadOnly := True;
        end;
      end;
    dmSave:
      begin
        { 檢查 Master }
        Result := ValidateMasterDataSet( aErrMsg );
        if not Result then
        begin
          WarningMsg( aErrMsg );
          Exit;
        end;
        { 存檔 }
        Result := ApplyDataSet( aErrMsg );
        if not Result then
        begin
          WarningMsg( aErrMsg );
          Exit;
        end;
        if ( cdsMaster.State in [dsInsert, dsEdit] ) then
        begin
          cdsMaster.AfterScroll := nil;
          try
            cdsMaster.Post;
          finally
            cdsMaster.AfterScroll := cdsMasterAfterScroll;
          end;
        end;  
        if gvView.IsUpdateLocked then gvView.EndUpdate;
      end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_1.EditorStateChange(const aMode: TDMLMode);
begin
  if ( aMode in [dmBrowser] ) then
  begin
    txtInvId.Enabled := True;
  end
  else if ( aMode in [dmInsert] ) then
  begin
    txtInvId.Enabled := True;
  end
  else if ( aMode in [dmUpdate] ) then
  begin
    txtInvId.Enabled := False;
  end;
  { 設定 Focus }
  case aMode of
    dmInsert, dmBrowser:
      begin
        if txtPaperNo.CanFocus then txtPaperNo.SetFocus;
      end;
    dmUpdate:
      begin
        if txtPaperNo.CanFocus then txtPaperNo.SetFocus;
      end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_1.GridStateChange(const aMode: TDMLMode);
begin
  case aMode of
    dmBrowser:
      begin
        gvGrid.Enabled := True;
      end;
    dmCancel:
      begin
        gvGrid.Enabled := True;
      end;
    dmInsert:
      begin
        gvGrid.Enabled := False;
      end;
    dmUpdate:
      begin
        gvGrid.Enabled := False;
      end;
    dmSave:
      begin
        gvGrid.Enabled := False;
      end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_1.SetMasterFieldState;
begin
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA07_1.ValidateMasterDataSet(var aErrMsg: String): Boolean;
var
  aFocusControl: TWinControl;
begin
  Result := False;
  aErrMsg := EmptyStr;
  aFocusControl := nil;
  try
    if ( FMode in [dmInsert, dmUpdate] ) then
    begin
      cdsMaster.FieldByName( 'YEARMONTH' ).AsString := GetYearMonth;
      cdsMaster.FieldByName( 'UPTEN' ).AsString := dtmMain.getLoginUser;
    end;
    if ( Trim( txtPaperNo.Text ) = EmptyStr  ) and
       ( cdsMaster.FieldByName( 'SOURCE' ).AsString = '1' ) then
    begin
      aErrMsg := '此單據來源為【批次產生】, 單據編號請填入原【收費單號】。';
      aFocusControl := txtPaperNo;
      Exit;
    end;
    if ( Trim( txtPaperDate.Text ) = EmptyStr ) then
    begin
      aErrMsg := '單據日期必須輸入。';
      aFocusControl := txtPaperDate;
      Exit;
    end;
    if ( Trim( txtYearMonth.Text ) = EmptyStr ) then
    begin
      aErrMsg := '申報年月必須輸入。';
      aFocusControl := txtYearMonth;
      Exit;
    end;
    if ( Trim( txtInvId.Text ) = EmptyStr ) then
    begin
      aErrMsg := '發票號碼必須輸入。';
      aFocusControl := txtInvId;
      Exit;
    end;
    if not CheckInvoiceIdExists( cdsMaster.FieldByName( 'INVID' ).AsString,
      cdsMaster.FieldByName( 'COMPID' ).AsString ) then
    begin
      aErrMsg := '無此發票號碼, 請確認。';
      aFocusControl := txtInvId;
      Exit;
    end;
    if ( cdsMaster.FieldByName( 'SOURCE' ).AsString = '1' ) and
       ( dtmMain.GetLinkToMIS ) then
    begin
      if not dtmSO.CheckAllowacneExists(
        cdsMaster.FieldByName( 'PAPERNO' ).AsString,
        cdsMaster.FieldByName( 'SEQ' ).AsString,
        cdsMaster.FieldByName( 'CUSTID' ).AsString ) then
      begin
        aErrMsg := '此單據來源為【批次產生】, 輸入的單據編號對應不到客服系統【收費單號】。';
        aFocusControl := txtPaperNo;
        Exit; 
      end;
    end;

//    if CheckPaperNoExists( cdsMaster.FieldByName( 'PAPERNO' ).AsString,
//      cdsMaster.FieldByName( 'INVID' ).AsString,
//      cdsMaster.FieldByName( 'COMPID' ).AsString ) then
//    begin
//      aErrMsg := '此單據號碼已存在, 請重新輸入。';
//      aFocusControl := txtPaperNo;
//      Exit;
//    end;
    if not isAmtOK then
    begin
      aErrMsg := '發票金額不等於銷售額加稅額!';
      aFocusControl := txtInvAmount;
      Exit;
    end;
    if dtmMainH.IsDataLocked( cdsMaster.FieldByName( 'YEARMONTH' ).AsString,
       cdsMaster.FieldByName( 'COMPID' ).AsString ) then
    begin
      aErrMsg := '此筆資料所屬公司別已經鎖帳,無法存檔!';
      aFocusControl := txtYearMonth;
      Exit;
    end;
  finally
    Result := ( aErrMsg = EmptyStr );
    if not Result then
    begin
      if Assigned( aFocusControl ) and
        ( aFocusControl.CanFocus ) then
      begin
        aFocusControl.SetFocus;
        if aFocusControl is TcxCustomEdit then
          TcxCustomEdit( aFocusControl ).SelectAll;
      end;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA07_1.ValidateCanDelete(var aErrMsg: String): Boolean;
begin
  aErrMsg := EmptyStr;
  try
    if cdsMaster.IsEmpty then
    begin
      aErrMsg := '無折讓資料, 不可刪除。';
      Result := False;
      Exit;
    end;
    if dtmMainH.IsDataLocked( cdsMaster.FieldByName( 'YEARMONTH' ).AsString,
       cdsMaster.FieldByName( 'COMPID' ).AsString ) then
    begin
      aErrMsg := '此筆資料所屬公司別已經鎖帳,無法刪除!';
      Result := False;
      Exit;
    end;
  finally
    Result := ( aErrMsg = EmptyStr );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA07_1.ValidateCanModify(var aErrMsg: String): Boolean;
begin
  aErrMsg := EmptyStr;
  try
    if cdsMaster.IsEmpty then
    begin
      aErrMsg := '無折讓資料, 不可修改。';
      Result := False;
      Exit;
    end;
    if dtmMainH.IsDataLocked( cdsMaster.FieldByName( 'YEARMONTH' ).AsString,
       cdsMaster.FieldByName( 'COMPID' ).AsString ) then
    begin
      aErrMsg := '此筆資料所屬公司別已經鎖帳,無法修改!';
      Result := False;
      Exit;
    end;
  finally
    Result := ( aErrMsg = EmptyStr );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA07_1.ApplyDataSet(var aErrMsg: String): Boolean;
var
  aDataSet: TADOQuery;

  { ----------------------------------------------- }

  procedure UpdateSoData;
  begin
    dtmSO.DropAllowance(
      cdsMaster.FieldByName( 'PAPERNO' ).AsString,
      cdsMaster.FieldByName( 'SEQ' ).AsString,
      cdsMaster.FieldByName( 'CUSTID' ).AsString,
      cdsMaster.FieldByName( 'ALLOWANCENO' ).AsString );
  end;

  { ----------------------------------------------- }

  procedure DeleteRecord;
  begin
    aDataSet.Close;
    aDataSet.SQL.Text := Format(
      ' DELETE FROM INV014 WHERE IDENTIFYID1 = ''%s''      ' +
      '  AND IDENTIFYID2 = ''%s'' AND YEARMONTH = ''%s''   ' +
      '  AND INVID = ''%s'' AND SEQ = ''%s''               ',
      [IDENTIFYID1, IDENTIFYID2, cdsMaster.FieldByName( 'YEARMONTH' ).AsString,
       cdsMaster.FieldByName( 'INVID' ).AsString, cdsMaster.FieldByName( 'SEQ' ).AsString] );
    aDataSet.ExecSQL;
  end;

  { ----------------------------------------------- }

  procedure ApplyInsertData;
  var
    aAllowanceNo: String;
  begin
    aAllowanceNo := dtmMainJ.GetAllowanceNo;
    aDataSet.Close;
    aDataSet.SQL.Text := Format(
      ' INSERT INTO INV014 ( IDENTIFYID1, IDENTIFYID2, COMPID,                   ' +
      '  CUSTID, PAPERNO, PAPERDATE, BUSINESSID, YEARMONTH,                      ' +
      '  INVID, SEQ, INVFORMAT, INVDATE, TAXTYPE, SALEAMOUNT,                    ' +
      '  TAXAMOUNT, INVAMOUNT, UPTTIME, UPTEN, ALLOWANCENO, SOURCE )             ' +
      ' values ( ''%s'', ''%s'', ''%s'',                                         ' +
      '  ''%s'', ''%s'', to_date( ''%s'', ''YYYYMMDD'' ), ''%s'', ''%s'',        ' +
      '  ''%s'', ''%s'', ''%s'', to_date( ''%s'', ''YYYYMMDD'' ), ''%s'', ''%s'',' +
      '  ''%s'', ''%s'', sysdate, ''%s'', ''%s'', 0 )                            ',
      [IDENTIFYID1, IDENTIFYID2,
       cdsMaster.FieldByName( 'COMPID' ).AsString,
       cdsMaster.FieldByName( 'CUSTID' ).AsString,
       cdsMaster.FieldByName( 'PAPERNO' ).AsString,
       FormatDateTime( 'YYYYMMDD', cdsMaster.FieldByName( 'PAPERDATE' ).AsDateTime ),
       cdsMaster.FieldByName( 'BUSINESSID' ).AsString,
       cdsMaster.FieldByName( 'YEARMONTH' ).AsString,
       cdsMaster.FieldByName( 'INVID' ).AsString,
       cdsMaster.FieldByName( 'SEQ' ).AsString,
       cdsMaster.FieldByName( 'INVFORMAT' ).AsString,
       FormatDateTime( 'YYYYMMDD', cdsMaster.FieldByName( 'invdate' ).AsDateTime ),
       cdsMaster.FieldByName( 'TAXTYPE' ).AsString,
       cdsMaster.FieldByName( 'SALEAMOUNT' ).AsString,
       cdsMaster.FieldByName( 'TAXAMOUNT' ).AsString,
       cdsMaster.FieldByName( 'INVAMOUNT' ).AsString,
       dtmMain.getLoginUser, aAllowanceNo] );
    aDataSet.ExecSQL;
  end;

  { ----------------------------------------------- }

  procedure ApplyUpdateData;
  begin
    aDataSet.Close;
    aDataSet.SQL.Text := Format(
      ' UPDATE INV014                     ' +
      '    SET CUSTID = ''%s'',           ' +
      '        PAPERNO = ''%s'',          ' +
      '        PAPERDATE = TO_DATE( ''%s'', ''YYYYMMDD'' ), ' +
      '        BUSINESSID = ''%s'',       ' +
      '        YEARMONTH = ''%s'',        ' +
      '        INVID = ''%s'',            ' +
      '        SEQ = ''%s'',              ' +
      '        INVFORMAT = ''%s'',        ' +
      '        INVDATE = TO_DATE( ''%s'', ''YYYYMMDD'' ),   ' +
      '        TAXTYPE = ''%s'',          ' +
      '        SALEAMOUNT = ''%s'',       ' +
      '        TAXAMOUNT = ''%s'',        ' +
      '        INVAMOUNT = ''%s'',        ' +
      '        UPTTIME = SYSDATE,         ' +
      '        UPTEN = ''%s''             ' +
      '  WHERE IDENTIFYID1 = ''%s''       ' +
      '    AND IDENTIFYID2 = ''%s''       ' +
      '    AND COMPID = ''%s''            ' +
      '    AND ALLOWANCENO = ''%s''       ',
      [cdsMaster.FieldByName( 'CUSTID' ).AsString,
       cdsMaster.FieldByName( 'PAPERNO' ).AsString,
       FormatDateTime( 'YYYYMMDD', cdsMaster.FieldByName( 'PAPERDATE' ).AsDateTime ),
       cdsMaster.FieldByName( 'BUSINESSID' ).AsString,
       cdsMaster.FieldByName( 'YEARMONTH' ).AsString,
       cdsMaster.FieldByName( 'INVID' ).AsString,
       cdsMaster.FieldByName( 'SEQ' ).AsString,
       cdsMaster.FieldByName( 'INVFORMAT' ).AsString,
       FormatDateTime( 'YYYYMMDD', cdsMaster.FieldByName( 'invdate' ).AsDateTime ),
       cdsMaster.FieldByName( 'TAXTYPE' ).AsString,
       cdsMaster.FieldByName( 'SALEAMOUNT' ).AsString,
       cdsMaster.FieldByName( 'TAXAMOUNT' ).AsString,
       cdsMaster.FieldByName( 'INVAMOUNT' ).AsString,
       dtmMain.getLoginUser,
       IDENTIFYID1,
       IDENTIFYID2,
       cdsMaster.FieldByName( 'COMPID' ).AsString,
       cdsMaster.FieldByName( 'ALLOWANCENO' ).AsString] );
    aDataSet.ExecSQL;
  end;

  { ----------------------------------------------- }

  procedure ApplyDeleteData;
  begin
    DeleteRecord;
    if ( dtmMain.GetLinkToMIS ) and
       ( cdsMaster.FieldByName( 'SOURCE' ).AsInteger = 1 ) then
      UpdateSoData;
  end;

  { ----------------------------------------------- }

var
  aConnection, aConnection2: TADOConnection;
begin
  Result := False;
  aDataSet := dtmMain.adoComm2;
  aConnection := dtmMain.InvConnection;
  aConnection2 := dtmSO.SOConnection;
  aConnection.BeginTrans;
  if ( dtmMain.GetLinkToMIS ) then aConnection2.BeginTrans;
  try
    if ( FMode in [dmInsert] ) then
    begin
      ApplyInsertData;
    end else
    if ( FMode in [dmUpdate] ) then
    begin
      ApplyUpdateData;
    end else
    if ( FMode in [dmDelete] ) then
    begin
      ApplyDeleteData;
    end;
    aConnection.CommitTrans;
    if ( dtmMain.GetLinkToMIS ) then aConnection2.CommitTrans;
    Result := True;
  except
    on E: Exception do
    begin
      TranslateError( E.Message );
      { PKey 重覆, PM 要求改秀訊息 }
      if OracleErros.ErrorCode = 1 then
      begin
        aErrMsg := '【折讓單號】取號重覆,請重新存檔!';
        if ( txtPaperNo.CanFocus ) then txtPaperNo.SetFocus;
      end
      else begin
        aErrMsg := E.Message;
      end;
      aConnection.RollbackTrans;
      if ( dtmMain.GetLinkToMIS ) then aConnection2.RollbackTrans;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_1.btnInsertClick(Sender: TObject);
  var
     aForm : TfrmInvA07_3;
begin
  {
  Screen.Cursor := crSQLWait;
  try
    DMLModeChange( dmInsert );
  finally
    Screen.Cursor := crDefault;
  end;
  Screen.Cursor := crSQLWait;
  }
  try
    aForm := TfrmInvA07_3.Create( nil );
    try
      aForm.INVID := EmptyStr;
      aForm.AllowanceNo := EmptyStr ;
      aForm.PrintConditionCompText := FConditionCompText;
      aForm.MainDataSet := cdsMaster;
      aForm.ShowModal;
    finally
      aForm.Free;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_1.btnEditClick(Sender: TObject);
var
  aErrMsg: String;
begin
  Screen.Cursor := crSQLWait;
  try
    if not ValidateCanModify( aErrMsg ) then
    begin
      WarningMsg( aErrMsg );
      Exit;
    end;
    DMLModeChange( dmUpdate );
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_1.btnCancelClick(Sender: TObject);
begin
  Screen.Cursor := crSQLWait;
  try
    DMLModeChange( dmCancel );
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_1.btnDeleteClick(Sender: TObject);
var
  aErrMsg: String;
begin
  Screen.Cursor := crSQLWait;
  try
    if not ValidateCanDelete( aErrMsg ) then
    begin
      WarningMsg( aErrMsg );
      Exit;
    end;
    if ConfirmMsg( '確認刪除此筆資料?' ) then
      DMLModeChange( dmDelete );
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_1.btnSaveClick(Sender: TObject);
begin
  Screen.Cursor := crSQLWait;
  try
    DMLModeChange( dmSave );
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_1.btnQuitClick(Sender: TObject);
begin
  Self.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_1.btnQueryClick(Sender: TObject);
label
  ReDo;
var
  aCount: Integer;
begin
  frmInvA07_2 := TfrmInvA07_2.Create( nil );
  try
ReDo:
    if ( frmInvA07_2.ShowModal = mrOK ) then
    begin
      aCount := DetectQueryRecordCount( frmInvA07_2.ConditionText );
      if ( aCount <= 0 ) then
      begin
        WarningMsg( '查無折讓資料。' );
        goto Redo;
      end;
      Screen.Cursor := crSQLWait;
      try
        FillMasterDataSet( frmInvA07_2.ConditionText );
        FConditionCompText := frmInvA07_2.ConditionCompText;
      finally
        Screen.Cursor := crDefault;
      end;
    end;
  finally
    frmInvA07_2.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_1.txtPaperDatePropertiesInitPopup(Sender: TObject);
begin
  if not ( FMode in [dmInsert, dmUpdate] ) then Abort;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA07_1.DetectQueryRecordCount(const aCondition: String): Integer;
begin
  dtmMain.adoComm.Close;
  dtmMain.adoComm.SQL.Clear;
  dtmMain.adoComm.SQL.Text := BuildQueryRecordCountSQL( aCondition );
  dtmMain.adoComm.Open;
  Result := dtmMain.adoComm.FieldByName( 'COUNTS' ).AsInteger;
  dtmMain.adoComm.Close;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA07_1.CheckInvoiceIsObsolete(const aInvId, aCompId: String): Boolean;
var
  aSql: String;
begin
  aSql := Format(
    '  SELECT ISOBSOLETE             ' +
    '    FROM INV007                 ' +
    '   WHERE IDENTIFYID1 = ''%s''   ' +
    '     AND IDENTIFYID2 = ''%s''   ' +
    '     AND INVID = ''%s''         ' +
    '     AND COMPID = ''%s''        ',
    [ IDENTIFYID1, IDENTIFYID2, aInvId, dtmMain.getCompID ] );
  dtmMain.adoComm.Close;
  dtmMain.adoComm.SQL.Text := aSQL;
  dtmMain.adoComm.Open;
  Result := ( dtmMain.adoComm.FieldByName( 'ISOBSOLETE' ).AsString = 'Y' );
  dtmMain.adoComm.Close;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA07_1.CheckInvoiceIdExists(const aInvId, aCompId: String): Boolean;
var
  aSql: String;
begin
  aSql := Format(
    '  SELECT INVID                  ' +
    '    FROM INV007                 ' +
    '   WHERE IDENTIFYID1 = ''%s''   ' +
    '     AND IDENTIFYID2 = ''%s''   ' +
    '     AND INVID = ''%s''         ' +
    '     AND COMPID = ''%s''        ',
    [ IDENTIFYID1, IDENTIFYID2, aInvId, aCompId] );
  dtmMain.adoComm.Close;
  dtmMain.adoComm.SQL.Text := aSQL;
  dtmMain.adoComm.Open;
  Result := ( not dtmMain.adoComm.IsEmpty );
  dtmMain.adoComm.Close;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA07_1.CheckPaperNoExists(const aPaperNo,
  aInvId, aCompId: String): Boolean;
var
  aSql: String;
begin
  aSql := Format(
    ' SELECT INVID FROM INV014                  ' +
    '  WHERE IDENTIFYID1 = ''%s''               ' +
    '    AND IDENTIFYID2 = ''%s''               ' +
    '    AND COMPID = ''%s''                    ' +
    '    AND PAPERNO = ''%s''                   ' +
    '    AND INVID <> ''%s''                    ',
    [ IDENTIFYID1, IDENTIFYID2, aCompId, aPaperNo, aInvId ] );
  dtmMain.adoComm.Close;
  dtmMain.adoComm.SQL.Text := aSQL;
  dtmMain.adoComm.Open;
  Result := ( not dtmMain.adoComm.IsEmpty );
  dtmMain.adoComm.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_1.txtYearMonthPropertiesValidate(Sender: TObject;
  var DisplayValue: Variant; var ErrorText: TCaption; var Error: Boolean);
begin
  if ( Error ) then
  begin
    WarningMsg( '您輸入的申報年月不正確。' );
    ErrorText := EmptyStr;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_1.txtPaperDatePropertiesValidate(Sender: TObject;
  var DisplayValue: Variant; var ErrorText: TCaption; var Error: Boolean);
begin
  if ( Error ) then
  begin
    WarningMsg( '您輸入的單據日期不正確。' );
    ErrorText := EmptyStr;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_1.txtInvIdExit(Sender: TObject);
var
  aInvFormat, aInvId, aInvDate, aTaxType, aCustId, aBusinessid, aCustName: String;
  aInvAmt, aSaleAmt, aTaxAmt: Integer;
begin
  if btnCancel.Focused then Exit;
  if not ( FMode in [dmInsert, dmUpdate] ) then Exit;
  aInvId := cdsMaster.FieldByName( 'INVID' ).AsString;
  if ( aInvId = EmptyStr ) then Exit;
  FInvExists := CheckInvoiceIdExists( aInvId,
    cdsMaster.FieldByName( 'COMPID' ).AsString );
  if not FInvExists then
  begin
    WarningMsg( '無此發票號碼' );
    if ( txtInvId.CanFocus ) then txtInvId.SetFocus;
    Exit;
  end;
  if CheckInvoiceIsObsolete( aInvId, cdsMaster.FieldByName( 'COMPID' ).AsString ) then
  begin
    WarningMsg( '此筆發票已做廢。' );
    if txtInvId.CanFocus then txtInvId.SetFocus;
    Exit;
  end;
  if ( dtmMainH.GetInvAmtData( aInvId, aInvAmt, aSaleAmt, aTaxAmt, aTaxType,
     aInvFormat, aInvDate, aCustId, aBusinessid, aCustName ) ) then
   begin
    cdsMaster.FieldByName('INVAMOUNT').AsInteger := aInvAmt;
    cdsMaster.FieldByName('SALEAMOUNT').AsInteger := aSaleAmt;
    cdsMaster.FieldByName('TAXAMOUNT').AsInteger := aTaxAmt;
    cdsMaster.FieldByName('INVFORMAT').AsString := aInvFormat;
    cdsMaster.FieldByName('TAXTYPE').AsString := aTaxType;
    cdsMaster.FieldByName('CUSTID').AsString := aCustId;
    cdsMaster.FieldByName('BUSINESSID').AsString := aBusinessid;
    if aInvDate <> EmptyStr then
      cdsMaster.FieldByName('INVDATE').AsDateTime := StrToDate( aInvDate )
    else
      cdsMaster.FieldByName('INVDATE').Value := EmptyStr;
    cdsMaster.FieldByName( 'TAXTYPEDESC' ).AsString := GetTaxTypeDesc( aTaxType );
    cdsMaster.FieldByName( 'INVFORMATDESC' ).AsString := GetInvFormatDesc( aInvFormat );
    cdsMaster.FieldByName( 'CUSTNAME' ).AsString := aCustName;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_1.CalcMoney(aKind: Integer);
var
  aTaxType: String;
  aTaxRate: Integer;
begin
  aTaxType := cdsMaster.FieldByName( 'TAXTYPE' ).AsString;
  case aKind of
    1:
      begin
        if ( cdsMaster.FieldByName( 'SALEAMOUNT' ).AsInteger = 0 ) then
        begin
          cdsMaster.FieldByName( 'TAXAMOUNT' ).AsInteger := 0;
          cdsMaster.FieldByName( 'INVAMOUNT' ).AsInteger := 0;
        end else
        begin
         aTaxRate := 0;
         if aTaxType = '1' then aTaxRate := 5;
         cdsMaster.FieldByName( 'TAXAMOUNT' ).AsInteger :=
           Round( TUother.RoundFloat( (
             cdsMaster.FieldByName( 'SALEAMOUNT' ).AsInteger * aTaxRate / 100 ), 0 ) );
         cdsMaster.FieldByName( 'INVAMOUNT' ).AsInteger := (
           cdsMaster.FieldByName( 'SALEAMOUNT' ).AsInteger +
           cdsMaster.FieldByName( 'TAXAMOUNT' ).AsInteger );
        end;
      end;
    2:
      begin
         cdsMaster.FieldByName( 'INVAMOUNT' ).AsInteger := (
           cdsMaster.FieldByName( 'SALEAMOUNT' ).AsInteger +
           cdsMaster.FieldByName( 'TAXAMOUNT' ).AsInteger );
      end;
    3:
      begin
        if ( cdsMaster.FieldByName( 'INVAMOUNT' ).AsInteger = 0 ) then
        begin
          cdsMaster.FieldByName( 'TAXAMOUNT' ).AsInteger := 0;
          cdsMaster.FieldByName( 'SALEAMOUNT' ).AsInteger := 0;
        end else
        begin
         aTaxRate := 0;
         if aTaxType = '1' then aTaxRate := 5;
         cdsMaster.FieldByName( 'SALEAMOUNT' ).AsInteger := Round(
           TUother.RoundFloat( cdsMaster.FieldByName( 'INVAMOUNT' ).AsInteger /
            ( 1 + ( aTaxRate / 100 ) ), 0 ) );
         cdsMaster.FieldByName( 'TAXAMOUNT' ).AsInteger := (
           cdsMaster.FieldByName( 'INVAMOUNT' ).AsInteger -
           cdsMaster.FieldByName( 'SALEAMOUNT' ).AsInteger );
        end;
      end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_1.txtSaleAmountExit(Sender: TObject);
begin
  if not ( FMode in [dmInsert, dmUpdate] ) then Exit;
  CalcMoney( 1 );

end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_1.txtTaxAmountExit(Sender: TObject);
begin
  if not ( FMode in [dmInsert, dmUpdate] ) then Exit;
  CalcMoney( 2 );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_1.txtInvAmountExit(Sender: TObject);
begin
  if not ( FMode in [dmInsert, dmUpdate] ) then Exit;
  CalcMoney( 3 );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_1.txtPaperNoExit(Sender: TObject);
begin
  if btnCancel.Focused then Exit;
  if not ( FMode in [dmInsert, dmUpdate] ) then Exit;
  if cdsMaster.FieldByName( 'INVID' ).AsString = EmptyStr then Exit;
  if CheckPaperNoExists( cdsMaster.FieldByName( 'PAPERNO' ).AsString,
    cdsMaster.FieldByName( 'INVID' ).AsString, cdsMaster.FieldByName( 'COMPID' ).AsString ) then
  begin
    WarningMsg( '此單據號碼已存在, 請重新輸入。' );
    if txtPaperNo.CanFocus then txtPaperNo.SetFocus ;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA07_1.IsAmtOK: Boolean;
var
  aInvAmt, aSaleAmt, aTaxAmt: Integer;
begin
  aInvAmt :=  cdsMaster.FieldByName('InvAmount').AsInteger;
  aSaleAmt := cdsMaster.FieldByName('SaleAmount').AsInteger;
  aTaxAmt := cdsMaster.FieldByName('TaxAmount').AsInteger;
  if ( aInvAmt <> aSaleAmt + aTaxAmt ) then
    Result := False
  else
    Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA07_1.GetYearMonth: String;
begin
  Result := EmptyStr;
  if ( cdsMaster.FieldByName( 'YEARMONTH2' ).AsString <> EmptyStr ) then
  begin
    Result :=
      Copy( cdsMaster.FieldByName( 'YEARMONTH2' ).AsString, 1, 4 ) +
      Lpad( Copy( cdsMaster.FieldByName( 'YEARMONTH2' ).AsString, 6, 2 ), 2, '0' );

  end;    
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_1.btnReportClick(Sender: TObject);
var
  WPath: String;
  aBookMark: TBookmark;
begin
  if cdsMaster.IsEmpty then Exit;
  WPath := IncludeTrailingPathDelimiter( ExtractFilePath(
    Application.ExeName ) ) + REPORT_FOLDER + '\FrptInvA07_1.fr3';
  dtmReport.frxMainReport.LoadFromFile( WPath );
  dtmReport.frxMainReport.Variables['公司別'] := QuotedStr( FConditionCompText );
  cdsMaster.AfterScroll := nil;
  try
    cdsMaster.DisableControls;
    try
      aBookMark := cdsMaster.GetBookmark;
      try
        cdsMaster.First;
        dtmReport.frxA07_1.DataSet := cdsMaster;
        try
          dtmReport.frxMainReport.ShowReport;
        finally
          dtmReport.frxA07_1.DataSet := nil;
        end;
        cdsMaster.GotoBookmark( aBookMark );
      finally
        cdsMaster.FreeBookmark( aBookMark );
      end;
    finally
      cdsMaster.EnableControls;
    end;
  finally
    cdsMaster.AfterScroll := cdsMasterAfterScroll;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_1.gvViewDblClick(Sender: TObject);
  var aInvId,aAllowanceNo : string;
    aForm : TfrmInvA07_3;
begin
  Screen.Cursor := crSQLWait;
  try

    if ( not cdsMaster.Active ) or ( cdsMaster.RecordCount=0 )  then Exit;

    aInvId := cdsMaster.FieldByName( 'INVID' ).AsString;
    aAllowanceNo := cdsMaster.FieldByName( 'AllowanceNo').AsString;
    aForm := TfrmInvA07_3.Create( nil );
    try
      aForm.INVID :=aInvId;
      aForm.AllowanceNo := aAllowanceNo;
      aForm.PrintConditionCompText := FConditionCompText;
      aForm.MainDataSet := cdsMaster;
      aForm.ShowModal;
    finally
      aForm.Free;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

end.
