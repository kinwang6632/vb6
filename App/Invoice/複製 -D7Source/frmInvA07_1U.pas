unit frmInvA07_1U;

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
    dsMaster: TDataSource;
    cdsMaster: TClientDataSet;
    btnReport: TcxButton;
    btnDelete: TcxButton;
    Panel1: TPanel;
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
    gvViewALLOWANCENO: TcxGridDBBandedColumn;
    gvViewSOURCE: TcxGridDBBandedColumn;
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
    procedure cdsMasterNewRecord(DataSet: TDataSet);
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
    procedure SetMasterFieldState;
    function ValidateCanModify(var aErrMsg: String): Boolean;
    function ValidateCanDelete(var aErrMsg: String): Boolean;
    function DetectQueryRecordCount(const aCondition: String): Integer;
    function CheckInvoiceIdExists(const aInvId, aCompId: String): Boolean;
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
    Result := '1.???|'
  else if ( aValue = '2' ) then
    Result := '2.?s?|?v'
  else  if ( aValue = '3' ) then
    Result := '3.?K?|';
end;

{ ---------------------------------------------------------------------------- }

function GetInvFormatDesc(const aValue: String): String;
begin
  Result := EmptyStr;
  if ( aValue = '1' ) then
    Result := '1.?q?l'
  else if ( aValue = '2' ) then
    Result := '2.???G'
  else  if ( aValue = '3' ) then
    Result := '3.???T';
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
  Caption := frmMain.GetFormTitleString( 'A07', '?P?f?h?^?P?????????d??' );
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
    CanClose := ConfirmMsg( '?????@?~?' );
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
  cdsMaster.FieldByName( 'SOURCE2' ).AsString := '???????J';
  
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
    '        DECODE( NVL( A.SOURCE, 0 ), 0, ''???????J'',       ' +
    '                ''????????'' ) SOURCE2                     ' +
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
//          Result := ApplyDataSet( aErrMsg );
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
        { ???d Master }
//        Result := ValidateMasterDataSet( aErrMsg );
        if not Result then
        begin
          WarningMsg( aErrMsg );
          Exit;
        end;
        { ?s?? }
//        Result := ApplyDataSet( aErrMsg );
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

function TfrmInvA07_1.ValidateCanDelete(var aErrMsg: String): Boolean;
begin
  aErrMsg := EmptyStr;
  try
    if cdsMaster.IsEmpty then
    begin
      aErrMsg := '?L????????, ???i?R???C';
      Result := False;
      Exit;
    end;
    if dtmMainH.IsDataLocked( cdsMaster.FieldByName( 'YEARMONTH' ).AsString,
       cdsMaster.FieldByName( 'COMPID' ).AsString ) then
    begin
      aErrMsg := '???????????????q?O?w?g???b,?L?k?R??!';
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
      aErrMsg := '?L????????, ???i?????C';
      Result := False;
      Exit;
    end;
    if dtmMainH.IsDataLocked( cdsMaster.FieldByName( 'YEARMONTH' ).AsString,
       cdsMaster.FieldByName( 'COMPID' ).AsString ) then
    begin
      aErrMsg := '???????????????q?O?w?g???b,?L?k????!';
      Result := False;
      Exit;
    end;
  finally
    Result := ( aErrMsg = EmptyStr );
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
    if ConfirmMsg( '?T?{?R???????????' ) then
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
        WarningMsg( '?d?L?????????C' );
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

procedure TfrmInvA07_1.txtYearMonthPropertiesValidate(Sender: TObject;
  var DisplayValue: Variant; var ErrorText: TCaption; var Error: Boolean);
begin
  if ( Error ) then
  begin
    WarningMsg( '?z???J???????~???????T?C' );
    ErrorText := EmptyStr;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_1.txtPaperDatePropertiesValidate(Sender: TObject;
  var DisplayValue: Variant; var ErrorText: TCaption; var Error: Boolean);
begin
  if ( Error ) then
  begin
    WarningMsg( '?z???J???????????????T?C' );
    ErrorText := EmptyStr;
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
  dtmReport.frxMainReport.Variables['???q?O'] := QuotedStr( FConditionCompText );
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
