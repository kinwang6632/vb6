unit frmInvA07_3U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, cxLookAndFeelPainters, StdCtrls, cxButtons, dxSkinsCore,
  dxSkinBlack, dxSkinBlue, dxSkinCaramel, dxSkinCoffee,
  dxSkinsDefaultPainters, cxStyles, dxSkinscxPCPainter, cxCustomData,
  cxGraphics, cxFilter, cxData, cxDataStorage, cxEdit, DB, cxDBData,
  cxCurrencyEdit, cxTextEdit, cxGridLevel, cxGridCustomTableView,
  cxGridTableView, cxGridBandedTableView, cxGridDBBandedTableView,
  cxClasses, cxControls, cxGridCustomView, cxGrid, cxSplitter, cxDBEdit,
  cxDBLabel, cxMaskEdit, cxDropDownEdit, cxCalendar, cxContainer, cxLabel,
  ExtCtrls, DBClient, ADODB, Provider, cxCheckBox;

type
  TDMLMode = ( dmBrowser, dmInsert, dmUpdate, dmSave, dmCancel, dmDelete );
  TfrmInvA07_3 = class(TForm)
    ScrollBox1: TScrollBox;
    btnInsert: TcxButton;
    btnEdit: TcxButton;
    btnCancel: TcxButton;
    btnSave: TcxButton;
    btnQuit: TcxButton;
    btnQuery: TcxButton;
    btnReport: TcxButton;
    btnDelete: TcxButton;
    Bevel1: TBevel;
    cdsMaster: TClientDataSet;
    cdsDetail: TClientDataSet;
    dsDetail: TDataSource;
    dsMaster: TDataSource;
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
    Label14: TLabel;
    Label15: TLabel;
    Label4: TcxLabel;
    Label5: TcxLabel;
    Label6: TcxLabel;
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
    lblAllowanceNo: TcxDBLabel;
    lblSource2: TcxDBLabel;
    cxSplitter1: TcxSplitter;
    gvGrid: TcxGrid;
    gvView: TcxGridDBBandedTableView;
    gvViewServiceTypeID: TcxGridDBBandedColumn;
    gvViewPARERNO: TcxGridDBBandedColumn;
    gvITEMID: TcxGridDBBandedColumn;
    gvViewSEQ: TcxGridDBBandedColumn;
    gvViewDescription: TcxGridDBBandedColumn;
    gvViewPAYTYPEID: TcxGridDBBandedColumn;
    gvViewPAYTYPEDESC: TcxGridDBBandedColumn;
    gvViewINVAMOUNT: TcxGridDBBandedColumn;
    gvViewUPTTIME: TcxGridDBBandedColumn;
    gvViewUPTEN: TcxGridDBBandedColumn;
    gvViewALLOWANCENO: TcxGridDBBandedColumn;
    gvLevel1: TcxGridLevel;
    Label3: TLabel;
    txtPaperNo: TcxDBTextEdit;
    chkUploadFlag: TcxDBCheckBox;
    lblUploadTime: TcxDBLabel;
    Label16: TLabel;
    btnDrop: TcxButton;
    cxLabel15: TcxLabel;
    lblISOBSOLETEDESC: TcxDBLabel;
    cxLabel16: TcxLabel;
    lblOBSOLETEREASON: TcxDBLabel;
    chkObUploadFlag: TcxDBCheckBox;
    lbl1: TLabel;
    lblObUploadTime: TcxDBLabel;
    chkAutoUploadFlag: TcxDBCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure cdsMasterAfterScroll(DataSet: TDataSet);
    procedure btnInsertClick(Sender: TObject);
    procedure cdsDetailNewRecord(DataSet: TDataSet);
    procedure btnCancelClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure cdsMasterNewRecord(DataSet: TDataSet);
    procedure btnSaveClick(Sender: TObject);
    procedure btnReportClick(Sender: TObject);
    procedure txtInvIdExit(Sender: TObject);
    procedure gvViewCustomDrawCell(Sender: TcxCustomGridTableView;
      ACanvas: TcxCanvas; AViewInfo: TcxGridTableDataCellViewInfo;
      var ADone: Boolean);
    procedure txtSaleAmountExit(Sender: TObject);
    procedure txtTaxAmountExit(Sender: TObject);
    procedure txtInvAmountExit(Sender: TObject);
    procedure gvViewINVAMOUNTPropertiesEditValueChanged(Sender: TObject);
    procedure btnQuitClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure cdsDetailAfterDelete(DataSet: TDataSet);
    procedure txtPaperNoExit(Sender: TObject);
    procedure btnDropClick(Sender: TObject);
  private
    { Private declarations }
    FMode: TDMLMode;
    FINVID,FAllowanceNo,FPaperNo : string;
    FSqlBody,FSqlWhere,FSqlOrder: string ;
    FSqlDetailBody,FSqlDetailWhere:string;
    FInvExists: Boolean;
    FConditionCompText: string;
    FDataSet : TClientDataSet;
    procedure PrepareMasterDataSet;
    procedure PrePareDetailDataSet;
    function FillMasterDataSet(const aCondition: String): Integer;
    function FillDetailDataSet(const aCondition: string): Integer;
    function BuildMasterSQL(const aCondition: String): String;
    function BuildDetailSQL(const aCondition: string): string;
    function DataSetStateChange(const aMode: TDMLMode): Boolean;
    procedure GridStateChange(const aMode: TDMLMode);
    procedure ButtonStateChange(const aMode: TDMLMode);
    procedure EditorStateChange(const aMode: TDMLMode);
    function CheckInvoiceIdExists(const aInvId, aCompId: String): Boolean;
    function ApplyDataSet(var aErrMsg: String): Boolean;
    procedure DisableDetailDataSetEvent;
    procedure EnableDetailDataSetEvent;
    function ValidateCanModify(var aErrMsg: String): Boolean;
    function ValidateMasterDataSet(var aErrMsg: String): Boolean;
    function ValidateDetailDataSet(var aErrMsg: String): Boolean;
    function GetYearMonth: String;
    function IsAmtOK: Boolean;
    function CheckInvoiceIsObsolete(const aInvId, aCompId: String): Boolean;
    function ValidateCanDelete(var aErrMsg: String): Boolean;
    procedure CalcMoney(aKind: Integer);
    procedure UpdateMasterInvAmount;
    procedure UpdateDetailInvAmount;
    procedure NewMainDataSet(const aCondition: String);
    function CheckINVAMTOver(aAllowanceNo: String;aInvId:String;var aMsg:string):Boolean;
    function ValidateCanDrop(var aErrMsg: String): Boolean;
  public
    { Public declarations }
    procedure DMLModeChange(const aMode: TDMLMode);
    property INVID: String  write FINVID ;
    property AllowanceNo: string write FAllowanceNo;
    property PrintConditionCompText: String write FConditionCompText;
    property MainDataSet : TClientDataSet write FDataSet;

  end;

var
  frmInvA07_3: TfrmInvA07_3;

implementation
uses frmMainU, dtmMainU, dtmSOU, dtmMainJU, dtmMainHU, cbUtilis,
  dtmReportModule, Uotheru,frmInvA07_5U;
{$R *.dfm}

{ TfrmInvA07_3 }
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


procedure TfrmInvA07_3.PrePareDetailDataSet;
begin
  if ( cdsDetail.FieldDefs.Count <=0 ) then
  begin
    with cdsDetail.FieldDefs do
    begin
      Add( 'ServiceTypeId',ftString,1 );
      Add( 'PaperNo',ftString,15);
      Add( 'Seq',ftInteger);
      Add( 'ItemId',ftString,7);
      Add( 'Description',ftString,40);
      Add( 'PayTypeId',ftString,7);
      Add( 'PayTypeDesc',ftString,40);
      Add( 'InvAmount',ftInteger);
      Add( 'AllowanceNo',ftString,12);
      Add( 'UPTTIME',ftDateTime);
      Add( 'UPTEN',ftString,20);
    end;
    cdsDetail.CreateDataSet;
  end;
  cdsDetail.EmptyDataSet;
end;

procedure TfrmInvA07_3.PrepareMasterDataSet;
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
    //#5881 ?W?[?????W?????O By Kin 2010/12/14
    cdsMaster.FieldDefs.Add( 'UPLOADFLAG',ftInteger);
    //#5945 ?W?[?????W?????? By Kin 2011/03/16
    cdsMaster.FieldDefs.Add( 'UPLOADTIME',ftDateTime);
    //#6001 ?W?[?@?o By Kin 2011/0712
    cdsMaster.FieldDefs.Add( 'ISOBSOLETE' ,ftString,10);
    cdsMaster.FieldDefs.Add( 'OBSOLETEID', ftString,7);
    cdsMaster.FieldDefs.Add( 'OBSOLETEREASON',ftString,100);
    {#6615 ?W?[?@?o?W?????O?P?@?o?W?????? By Kin 2014/01/06}
    cdsMaster.FieldDefs.Add( 'OBUPLOADFLAG', ftInteger );
    cdsMaster.FieldDefs.Add( 'OBUPLOADTIME', ftDateTime );
    cdsMaster.FieldDefs.Add( 'AUTOUPLOADFLAG', ftInteger );
    cdsMaster.CreateDataSet;
  end;
  cdsMaster.EmptyDataSet;
end;

procedure TfrmInvA07_3.FormCreate(Sender: TObject);
begin
  FMode := dmBrowser;
  PrepareMasterDataSet;
//  PrePareDetailDataSet;
end;

procedure TfrmInvA07_3.FormShow(Sender: TObject);
  var   aEvent: TDataSetNotifyEvent;
begin
  Caption := frmMain.GetFormTitleString( 'A07_3', '?P?f?h?^?P???????????J' );
  Self.Update;
  Screen.Cursor := crSQLWait;
  try
    Application.ProcessMessages;
    aEvent := cdsMaster.AfterScroll;
    cdsMaster.AfterScroll := nil;
    try
      FillMasterDataSet( FINVID );
      Application.ProcessMessages;
      if cdsMaster.IsEmpty then
      begin
        PrePareDetailDataSet;
        DMLModeChange( dmInsert );
      end else
        DMLModeChange( dmBrowser );
    finally
      cdsMaster.AfterScroll := aEvent;
    end;
    Application.ProcessMessages;
  finally
    Screen.Cursor := crDefault;
  end;
  
end;
function TfrmInvA07_3.FillMasterDataSet(const aCondition: String): Integer;
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
      {#6423 ?p?GTaxAmount??0???n????Null By Kin 2013/02/08}
      if ( cdsMaster.FieldByName( 'TAXAMOUNT' ).AsString = EmptyStr ) then
        cdsMaster.FieldByName( 'TAXAMOUNT' ).AsString :='0';
      cdsMaster.FieldByName( 'INVAMOUNT' ).AsString := dtmMain.adoComm.FieldByName( 'INVAMOUNT' ).AsString;
      cdsMaster.FieldByName( 'UPTTIME' ).AsString := dtmMain.adoComm.FieldByName( 'UPTTIME' ).AsString;
      cdsMaster.FieldByName( 'UPTEN' ).AsString := dtmMain.adoComm.FieldByName( 'UPTEN' ).AsString;
      cdsMaster.FieldByName( 'TAXTYPEDESC' ).AsString := GetTaxTypeDesc( dtmMain.adoComm.FieldByName( 'TAXTYPE' ).AsString );
      cdsMaster.FieldByName( 'INVFORMATDESC' ).AsString := GetInvFormatDesc( dtmMain.adoComm.FieldByName( 'INVFORMAT' ).AsString );
      cdsMaster.FieldByName( 'ALLOWANCENO' ).AsString := dtmMain.adoComm.FieldByName( 'ALLOWANCENO' ).AsString;
      cdsMaster.FieldByName( 'SOURCE' ).AsInteger := dtmMain.adoComm.FieldByName( 'SOURCE' ).AsInteger;
      cdsMaster.FieldByName( 'SOURCE2' ).AsString := dtmMain.adoComm.FieldByName( 'SOURCE2' ).AsString;
      //#5881 ?W?[?????W?????O By Kin 2010/12/14
      cdsMaster.FieldByName( 'UPLOADFLAG' ).AsInteger := dtmMain.adoComm.FieldByName( 'UPLOADFLAG' ).AsInteger;
      //#5945 ?W?[?????W?????? By Kin 2011/03/16
      cdsMaster.FieldByName( 'UPLOADTIME' ).AsString := dtmMain.adoComm.FieldByName( 'UPLOADTIME' ).AsString;
      //#6001 ?W?[?@?o By Kin 2011/07/12
      if dtmMain.adoComm.FieldByName( 'ISOBSOLETE' ).AsString = 'Y' then
        cdsMaster.FieldByName( 'ISOBSOLETE' ).AsString := '?O'
      else
        cdsMaster.FieldByName( 'ISOBSOLETE' ).AsString := '?_';
      cdsMaster.FieldByName( 'OBSOLETEID' ).AsString := dtmMain.adoComm.FieldByName( 'OBSOLETEID' ).AsString;
      cdsMaster.FieldByName( 'OBSOLETEREASON' ).AsString := dtmMain.adoComm.FieldByName( 'OBSOLETEREASON' ).AsString;
      {#6615 ?W?[?@?o?W???????P?@?o?W?????O By Kin 2014/01/06}
      cdsMaster.FieldByName( 'OBUPLOADFLAG' ).AsString := dtmMain.adoComm.FieldByName( 'OBUPLOADFLAG' ).AsString;
      cdsMaster.FieldByName( 'OBUPLOADTIME' ).AsString := dtmMain.adoComm.FieldByName( 'OBUPLOADTIME' ).AsString;
      {#6615 ?W?[?P?_?O?_??GW?W?? By Kin 2014/01/16}
      cdsMaster.FieldByName( 'AUTOUPLOADFLAG' ).AsString := dtmMain.adoComm.FieldByName( 'AUTOUPLOADFLAG' ).AsString;
      cdsMaster.Post;
      dtmMain.adoComm.Next;
    end;
    cdsMaster.First;
  finally
    cdsMaster.EnableControls;
    cdsMaster.AfterScroll := cdsMasterAfterScroll;
  end;
  dtmMain.adoComm.Close;
  Result := cdsMaster.RecordCount;
end;

function TfrmInvA07_3.BuildMasterSQL(const aCondition: String): String;
begin

 {#5945 ?W?[?W??????(UploadTime) By Kin 2011/03/16 }
 {#6615 ?W?[?@?o?W?????O?P???? By Kin 2014/01/06}
 FSqlBody := Format(
    ' SELECT A.IDENTIFYID1,                                     ' +
    '        A.IDENTIFYID2,                                     ' +
    '        A.COMPID,                                          ' +
    '        B.COMPSNAME,                                       ' +
    '        A.CUSTID,                                          ' +
    '        C.CUSTSNAME,                                       ' +
    '        A.PAPERNO,                                         ' +
    '        A.ISOBSOLETE,                                      ' +
    '        A.OBSOLETEID,                                      ' +
    '        A.OBSOLETEREASON,                                  ' +
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
    '        A.UPLOADTIME,                                      ' +
    '        Nvl(A.UPLOADFLAG,0) UPLOADFLAG,                    ' +
    '        Nvl(A.OBUPLOADFLAG,0) OBUPLOADFLAG,                ' +
    '        A.OBUPLOADTIME,                                    ' +
    '        Nvl(A.AUTOUPLOADFLAG,0) AUTOUPLOADFLAG,            ' +
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
  if aCondition=EmptyStr then
  begin
    FSqlWhere := ' AND 1 = 2 ';
  end else
  begin
    { #6014 ?p?G???????????N?H???????????D?A???M?P?@???o???????|???????? By Kin 2011/05/13 }
    if FAllowanceNo <> EmptyStr then
    begin
      FSqlWhere := ' AND A.ALLOWANCENO = ''' + FAllowanceNo + '''';
    end else
    begin
      FSqlWhere := ' And A.InvId='''+ aCondition +'''';
    end;
  end;
  {}
  FSqlOrder := ' ORDER BY A.COMPID, A.YEARMONTH, A.PAPERNO ';

  Result := ( FSqlBody + FSqlWhere + FSqlOrder );


end;

function TfrmInvA07_3.BuildDetailSQL(const aCondition: string): string;
begin
  FSqlDetailBody:='Select A.* From INV014A A,INV014 B '+
                  'Where A.AllowanceNo=B.AllowanceNo ';
  if aCondition=EmptyStr then
    FSqlDetailWhere := ' And 1 =2 '
  else
    FSqlDetailWhere := ' And A.AllowanceNo=''' + aCondition + '''';
  Result := FSqlDetailBody + FSqlDetailWhere ;


end;

function TfrmInvA07_3.FillDetailDataSet(const aCondition: string): Integer;
begin
  cdsDetail.EmptyDataSet;
  dtmMain.adoComm.Close;
  dtmMain.adoComm.SQL.Text := BuildDetailSQL(aCondition);
  dtmMain.adoComm.Open;
  dtmMain.adoComm.First;
  cdsDetail.ReadOnly := False;
  cdsDetail.DisableControls;
  cdsDetail.AfterScroll := nil;
  try
    while not dtmMain.adoComm.Eof do
    begin
      cdsDetail.Append;
      cdsDetail.FieldByName('ServiceTypeId').AsString := dtmMain.adoComm.FieldByName( 'ServiceTypeId').AsString;
      cdsDetail.FieldByName( 'PaperNo' ).AsString := dtmMain.adoComm.FieldByName( 'PaperNo' ).AsString;
      cdsDetail.FieldByName( 'Seq' ).AsInteger := dtmMain.adoComm.FieldByName ('Seq').AsInteger;
      cdsDetail.FieldByName( 'ItemId' ).AsString := dtmMain.adoComm.FieldByName( 'ItemId' ).AsString;
      cdsDetail.FieldByName( 'Description' ).AsString := dtmMain.adoComm.FieldByName( 'Description' ).AsString;
      cdsDetail.FieldByName( 'PayTypeId' ).AsString := dtmMain.adoComm.FieldByName( 'PayTypeId' ).AsString;
      cdsDetail.FieldByName( 'PayTypeDesc' ).AsString := dtmMain.adoComm.FieldByName( 'PayTypeDesc' ).AsString;
      cdsDetail.FieldByName( 'InvAmount' ).AsInteger := dtmMain.adoComm.FieldByName( 'InvAmount' ).AsInteger;
      cdsDetail.FieldByName( 'UPTTIME' ).AsString := dtmMain.adoComm.FieldByName( 'UPTTIME' ).AsString;
      cdsDetail.FieldByName( 'UPTEN' ).AsString := dtmMain.adoComm.FieldByName( 'UPTEN' ).AsString;
      cdsDetail.FieldByName( 'AllowanceNo').AsString := dtmMain.adoComm.FieldByName( 'AllowanceNo' ).AsString;
      cdsDetail.Post;
      dtmMain.adoComm.Next;
    end;
    cdsDetail.First;
  finally
    cdsDetail.EnableControls;
  end;
  dtmMain.adoComm.Close;
  Result := cdsDetail.RecordCount;
  
end;

procedure TfrmInvA07_3.DMLModeChange(const aMode: TDMLMode);
begin
  if not DataSetStateChange( aMode ) then Exit;
  FMode := aMode;
  if ( aMode in [dmSave, dmCancel, dmDelete] ) then
    FMode := dmBrowser;
  GridStateChange( FMode );
  ButtonStateChange( FMode );
  EditorStateChange( FMode );
end;

function TfrmInvA07_3.DataSetStateChange(const aMode: TDMLMode): Boolean;
  var aErrMsg : string;
begin
 Result := True;
  aErrMsg := '';
  case aMode of
    dmBrowser:
      begin
        cdsMaster.ReadOnly := True;
        cdsDetail.ReadOnly := True;
        if gvView.IsUpdateLocked then gvView.EndUpdate;
        cdsMasterAfterScroll( cdsMaster );
      end;
    dmCancel:
      begin
        cdsMaster.AfterScroll := nil;
        cdsDetail.AfterScroll := nil;
        try
          cdsMaster.Cancel;
          cdsDetail.Cancel;
        finally
          cdsMaster.AfterScroll := cdsMasterAfterScroll;
        end;
        cdsMaster.ReadOnly := True;
        cdsDetail.ReadOnly := True;
        if gvView.IsUpdateLocked then gvView.EndUpdate;
        cdsMasterAfterScroll( cdsMaster );
        cdsMaster.ReadOnly := True;
        cdsDetail.ReadOnly := True;
      end;
    dmInsert:
      begin
        cdsMaster.ReadOnly := False;
        cdsDetail.ReadOnly := False;
        cdsMaster.Append;
      end;
    dmUpdate:
      begin
        cdsMaster.ReadOnly := False;
        cdsDetail.ReadOnly := False;
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
        { ???d Master }
        Result := ValidateMasterDataSet( aErrMsg );
        if Result then
          Result := ValidateDetailDataSet( aErrMsg );
        if not Result then
        begin
          WarningMsg( aErrMsg );
          Exit;
        end;
        { ?s?? }
        Result := ApplyDataSet( aErrMsg );
        if not Result then
        begin
          WarningMsg( aErrMsg );
          cdsMaster.ReadOnly := True;
          cdsDetail.ReadOnly := True;
          Exit;
        end;
        if ( cdsMaster.State in [dsInsert, dsEdit] ) then
        begin
          cdsMaster.AfterScroll := nil;
          try
            cdsMaster.FieldByName('AllowanceNo').AsString := FAllowanceNo;
            cdsMaster.FieldByName( 'PaperNo' ).AsString := FPaperNo;
            cdsMaster.Post;
          finally
            cdsMaster.AfterScroll := cdsMasterAfterScroll;
            cdsMaster.ReadOnly := True;
            cdsDetail.ReadOnly := True;
          end;
        end;
        if gvView.IsUpdateLocked then gvView.EndUpdate;

      end;
  end;
end;

procedure TfrmInvA07_3.GridStateChange(const aMode: TDMLMode);
begin
  case aMode of
    dmBrowser, dmCancel:
      begin
        gvView.OptionsSelection.CellSelect := False;
        gvView.OptionsView.ShowEditButtons := gsebNever;
        gvView.Styles.Content := dtmMain.cxGridBackGroundStyle;
        gvView.OptionsData.Deleting := False;

        //DetailGridView.OptionsSelection.CellSelect := False;
        //DetailGridView.OptionsView.ShowEditButtons := gsebNever;
        //DetailGridView.Styles.Content := dtmMain.cxGridBackGroundStyle;
        {
        DetailGridView.GetColumnByFieldName( 'SEQ' ).Styles.Content := nil;
        DetailGridView.GetColumnByFieldName( 'BILLID' ).Styles.Content := nil;
        DetailGridView.GetColumnByFieldName( 'BILLIDITEMNO' ).Styles.Content := nil;
        DetailGridView.GetColumnByFieldName( 'STARTDATE' ).Styles.Content := nil;
        DetailGridView.GetColumnByFieldName( 'ENDDATE' ).Styles.Content := nil;
        DetailGridView.GetColumnByFieldName( 'SERVICETYPE' ).Styles.Content := nil;
        DetailGridView.GetColumnByFieldName( 'TAXNAME' ).Styles.Content := nil;
        DetailGridView.GetColumnByFieldName( 'SIGN' ).Styles.Content := nil;
        DetailGridLevel2.Visible := True;
        }
      end;
    dmInsert:
      begin
        gvView.OptionsSelection.CellSelect := True;
        gvView.OptionsView.ShowEditButtons := gsebForFocusedRecord;
        gvView.OptionsData.Deleting := True;
        gvView.OptionsData.DeletingConfirmation := True;
        gvView.Styles.Content := nil;
        {
        DetailGridView.OptionsSelection.CellSelect := True;
        DetailGridView.OptionsView.ShowEditButtons := gsebForFocusedRecord;
        DetailGridView.Styles.Content := nil;
        DetailGridView.GetColumnByFieldName( 'TAXNAME' ).Options.Editing := False;
        DetailGridView.GetColumnByFieldName( 'SIGN' ).Options.Editing := False;
        DetailGridView.GetColumnByFieldName( 'TAXNAME' ).Options.Focusing := False;
        DetailGridView.GetColumnByFieldName( 'SIGN' ).Options.Focusing := False;
        DetailGridView.GetColumnByFieldName( 'SEQ' ).Options.Focusing := False;
        DetailGridView.GetColumnByFieldName( 'SEQ' ).Options.Editing := False;
        DetailGridView.GetColumnByFieldName( 'SEQ' ).Styles.Content :=
          dtmMain.cxGridNotEditStyle;
        DetailGridView.GetColumnByFieldName( 'TAXNAME' ).Styles.Content :=
          dtmMain.cxGridNotEditStyle;
        DetailGridView.GetColumnByFieldName( 'SIGN' ).Styles.Content :=
          dtmMain.cxGridNotEditStyle;
        DetailGridLevel2.Visible := False;
        }
      end;
    dmUpdate:
      begin
        gvView.OptionsSelection.CellSelect := True;
        gvView.OptionsView.ShowEditButtons := gsebForFocusedRecord;
//        gvView.OptionsData.Editing := True;

        if cdsMaster.FieldByName( 'Source' ).AsString = '0' then
        begin
          gvView.OptionsSelection.CellSelect := False;
          gvView.OptionsData.Editing := False;
          gvView.OptionsData.Deleting := False;
        end else
        begin
          gvView.OptionsData.Deleting := True;
          gvView.OptionsData.DeletingConfirmation := True;
        end;
        gvView.Styles.Content := nil;
        {
        DetailGridView.OptionsSelection.CellSelect := True;
        DetailGridView.OptionsView.ShowEditButtons := gsebForFocusedRecord;
        DetailGridView.Styles.Content := nil;
        DetailGridView.GetColumnByFieldName( 'SEQ' ).Styles.Content :=
          dtmMain.cxGridNotEditStyle;
        DetailGridView.GetColumnByFieldName( 'TAXNAME' ).Styles.Content :=
          dtmMain.cxGridNotEditStyle;
        DetailGridView.GetColumnByFieldName( 'SIGN' ).Styles.Content :=
          dtmMain.cxGridNotEditStyle;
        DetailGridView.GetColumnByFieldName( 'SEQ' ).Options.Editing := False;
        DetailGridView.GetColumnByFieldName( 'TAXNAME' ).Options.Editing := False;
        DetailGridView.GetColumnByFieldName( 'SIGN' ).Options.Editing := False;
        DetailGridView.GetColumnByFieldName( 'SEQ' ).Options.Focusing := False;
        DetailGridView.GetColumnByFieldName( 'TAXNAME' ).Options.Focusing := False;
        DetailGridView.GetColumnByFieldName( 'SIGN' ).Options.Focusing := False;
        DetailGridLevel2.Visible := False;
        }
      end;

  end;
//  FillDetailColumnPropertiesValue;
end;

procedure TfrmInvA07_3.ButtonStateChange(const aMode: TDMLMode);
begin
  {
  if ( cdsMaster.IsEmpty ) and ( aMode in [dmInsert] ) then
  begin
    btnInsert.Enabled := True;
    btnCancel.Enabled := True;
    btnEdit.Enabled := False;
    btnSave.Enabled := False;
    btnQuit.Enabled := False;
    btnQuery.Enabled := False;
    btnReport.Enabled := False;
    btnDelete.Enabled := True;
    Exit;
  end;
  }
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
        btnQuit.Enabled := True;
        btnDrop.Enabled := True;
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
        btnQuit.Enabled := False;
        btnDrop.Enabled := False;
      end;
  end;
end;

procedure TfrmInvA07_3.EditorStateChange(const aMode: TDMLMode);
begin
   txtPaperNo.Properties.ReadOnly := False;
   txtPaperDate.Properties.ReadOnly := False;
   txtInvId.Properties.ReadOnly := False;
   txtSaleAmount.Properties.ReadOnly := False;
   txtTaxAmount.Properties.ReadOnly := False;
   txtInvAmount.Properties.ReadOnly := False;
  if ( aMode in [dmBrowser] ) then
  begin
    txtInvId.Enabled := True;
  end
  else if ( aMode in [dmInsert] ) then
  begin
    txtInvId.Enabled := True;
    txtPaperNo.Enabled := True;
    txtPaperNo.SetFocus;
  end
  else if ( aMode in [dmUpdate] ) then
  begin
    txtInvId.Enabled := False;
    if cdsMaster.FieldByName( 'Source' ).AsString = '0' then
    begin
      txtPaperNo.Enabled := True;
      txtPaperNo.SetFocus;
    end else
    begin
      txtPaperNo.Enabled := False;
      txtPaperDate.SetFocus;
    end;
    //#6001 ?????~???n???????A???????? By Kin 2011/07/13
    if cdsMaster.FieldByName( 'UPLOADFLAG' ).AsString = '1' then
    begin
      txtPaperNo.Properties.ReadOnly := True;
      txtPaperDate.Properties.ReadOnly := True;
      txtInvId.Properties.ReadOnly := True;
      txtSaleAmount.Properties.ReadOnly := True;
      txtTaxAmount.Properties.ReadOnly := True;
      txtInvAmount.Properties.ReadOnly := True;

    end;

  end;


end;

function TfrmInvA07_3.CheckInvoiceIdExists(const aInvId, aCompId: String): Boolean;
  var aSql : string;
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

procedure TfrmInvA07_3.cdsMasterAfterScroll(DataSet: TDataSet);
begin
  Screen.Cursor := crSQLWait;
  try
    FAllowanceNo := cdsMaster.FieldByName( 'AllowanceNo' ).AsString;
    FINVID := cdsMaster.FieldByName( 'INVID' ).AsString;
    DisableDetailDataSetEvent;
    try
      { INV014A }
      PrepareDetailDataSet;
      FillDetailDataSet(FAllowanceNo) ;
    finally
      EnableDetailDataSetEvent;
    end;

  finally
    Screen.Cursor := crDefault;
  end;



  FInvExists := CheckInvoiceIdExists(
    DataSet.FieldByName( 'INVID' ).AsString, DataSet.FieldByName( 'COMPID' ).AsString );
end;

function TfrmInvA07_3.ApplyDataSet(var aErrMsg: String): Boolean;
  var
  aDataSet: TADOQuery;
  aSysDate : TDateTime;
  aAllowanceNo : string;
  aPaperNo : string;
  {-------------------------------------------------------------------}
  { ???o???s???? }
  procedure GetSysDate;
  begin
    aDataSet.Close;
    aDataSet.SQL.Text := 'SELECT SYSDATE FROM DUAL';
    aDataSet.Open;
    aSysDate := aDataSet.FieldList[0].AsDateTime;
    aDataSet.Close;
  end;
  {-------------------------------------------------------------------}

  {-------------------------------------------------------------------}
  {???o????????}
  procedure GetAllowance;
  begin
    if FMode in [dmInsert] then
      aAllowanceNo := dtmMainJ.GetAllowanceNo
    else
      aAllowanceNo := cdsMaster.FieldByName( 'ALLOWANCENO' ).AsString;
  end;
  {-------------------------------------------------------------------}
  {-------------------------------------------------------------------}
  {???o?????s??}
  procedure GetPaperNo;
  begin
    aPaperNo := cdsMaster.FieldByName( 'PaperNo' ).AsString;
    if ( aPaperNo = '' ) then
    begin
      aPaperNo := aAllowanceNo;
    end;
  end;
  {-------------------------------------------------------------------}
  {--------------------------------------------------------------------}
  { ???s SO034?????? }
  procedure UpdateSoData;
  begin
    try
      cdsDetail.DisableControls;
      cdsDetail.First;
      while not cdsDetail.Eof do
      begin
        dtmSO.DropAllowance(
        cdsDetail.FieldByName( 'PAPERNO' ).AsString,
        cdsDetail.FieldByName( 'SEQ' ).AsString,
        cdsMaster.FieldByName( 'CUSTID' ).AsString,
        cdsMaster.FieldByName( 'ALLOWANCENO' ).AsString );
        cdsDetail.Next;
      end;
    finally
      cdsDetail.EnableControls;
    end;


  end;
  {--------------------------------------------------------------------}

  {--------------------------------------------------------------------}
  { ?R??INV014A?????? }
  procedure DeleteDetailRecord;
  begin
    aDataSet.Close;
    aDataSet.SQL.Text := Format(
                                ' Delete From INV014A Where ALLOWANCENO=''%s''     ',
                                [ cdsMaster.FieldByName( 'ALLOWANCENO').AsString ] );
    aDataSet.ExecSQL;
  end;
  {--------------------------------------------------------------------}

  {--------------------------------------------------------------------}
  { ?R?? Inv014 ?????? }
  procedure DeleteRecord;
  begin
    {
    aDataSet.SQL.Text := Format(
      ' DELETE FROM INV014 WHERE IDENTIFYID1 = ''%s''      ' +
      '  AND IDENTIFYID2 = ''%s'' AND YEARMONTH = ''%s''   ' +
      '  AND INVID = ''%s'' AND SEQ = ''%s''               '  ,
      [IDENTIFYID1, IDENTIFYID2, cdsMaster.FieldByName( 'YEARMONTH' ).AsString,
       cdsMaster.FieldByName( 'INVID' ).AsString, cdsMaster.FieldByName( 'SEQ' ).AsString] );
    }
    {#5302 ?R????,?u?w?????????R??(?]??INVID???i???O????),?y?k????PK????????,???K?R?????? By Kin 2009/09/21}
    aDataSet.SQL.Text := Format(
      ' DELETE FROM INV014 WHERE IDENTIFYID1 =''%s''        ' +
      '   AND IDENTIFYID2 = %s AND COMPID = ''%s''          ' +
      '   AND AllowanceNo = ''%s''                          ' ,
      [IDENTIFYID1,IDENTIFYID2,
       cdsMaster.FieldByName( 'COMPID' ).AsString ,
       cdsMaster.FieldByName( 'AllowanceNo').AsString ] );
    aDataSet.ExecSQL;
  end;
  {--------------------------------------------------------------------}

  {--------------------------------------------------------------------}
  { ???JINV014A??????}
  procedure ApplyInsertDetailData;
  begin
    try
      cdsDetail.DisableControls;
      cdsDetail.First;
      while not cdsDetail.Eof do
      begin
        aDataSet.Close;
        aDataSet.SQL.Text := Format(
                                  ' INSERT INTO INV014A (ServiceTypeID,PaPerNo,      ' +
                                  ' Seq,ItemId,Description,PayTypeId,PayTypeDesc,    ' +
                                  ' InvAmount,AllowanceNo,UPTTIME,UPTEN )            ' +
                                  ' Values (''%s'',''%s'',                           ' +
                                  ' %d,''%s'',''%s'',''%s'',''%s'',                  ' +
                                  ' %d,''%s'',to_date(''%s'',''YYYY/MM/DD HH24:MI:SS''),      ' +
                                  ' ''%s'')                                          ' ,
                                  [ cdsDetail.FieldByName( 'ServiceTypeID' ).AsString,
                                    cdsDetail.FieldByName( 'PaPerNo' ).AsString,
                                    cdsDetail.FieldByName( 'Seq' ).AsInteger,
                                    cdsDetail.FieldByName( 'ItemId' ).AsString,
                                    cdsDetail.FieldByName( 'Description' ).AsString,
                                    cdsDetail.FieldByName( 'PayTypeId' ).AsString,
                                    cdsDetail.FieldByName( 'PayTypeDesc' ).AsString,
                                    cdsDetail.FieldByName( 'InvAmount' ).AsInteger,
                                    aAllowanceNo,
                                    FormatDateTime('YYYYMMDD HHMMSS',aSysDate),
                                    dtmMain.getLoginUser ] );


        aDataSet.ExecSQL;
        cdsDetail.Next;
      end;
    finally
      cdsDetail.EnableControls;
    end;



  end;
  {--------------------------------------------------------------------}

  {--------------------------------------------------------------------}
  { ???JINV014 ????}
  procedure ApplyInsertData;
    var
      aTaxAmount : String;
  begin
    {#6423 ?p?GTaxAmount = null ?n?NtaxAmount????0 ,Jacy?o??A07???O?????D By Kin 2014/02/10}
    aTaxAmount := cdsMaster.FieldByName( 'TAXAMOUNT' ).AsString;
    if aTaxAmount = '' then
    begin
      aTaxAmount := '0';
    end;
    aDataSet.Close;
    aDataSet.SQL.Text := Format(
      ' INSERT INTO INV014 ( IDENTIFYID1, IDENTIFYID2, COMPID,                   ' +
      '  CUSTID, PAPERDATE, BUSINESSID, YEARMONTH,                               ' +
      '  INVID, SEQ, INVFORMAT, INVDATE, TAXTYPE, SALEAMOUNT,                    ' +
      '  TAXAMOUNT, INVAMOUNT, UPTTIME, UPTEN, ALLOWANCENO, SOURCE,PaperNo,      ' +
      '  UPLOADFLAG )                                                            ' +
      ' values ( ''%s'', ''%s'', ''%s'',                                         ' +
      '  ''%s'',  to_date( ''%s'', ''YYYYMMDD'' ), ''%s'', ''%s'',               ' +
      '  ''%s'', ''%s'', ''%s'', to_date( ''%s'', ''YYYYMMDD'' ), ''%s'', ''%s'',' +
      '  ''%s'', ''%s'', to_date(''%s'',''YYYY/MM/DD HH24:MI:SS''),              ' +
      ' ''%s'', ''%s'', 0 ,''%s'',0)                                                      ',
      [IDENTIFYID1, IDENTIFYID2,
       cdsMaster.FieldByName( 'COMPID' ).AsString,
       cdsMaster.FieldByName( 'CUSTID' ).AsString,
       FormatDateTime( 'YYYYMMDD', cdsMaster.FieldByName( 'PAPERDATE' ).AsDateTime ),
       cdsMaster.FieldByName( 'BUSINESSID' ).AsString,
       cdsMaster.FieldByName( 'YEARMONTH' ).AsString,
       cdsMaster.FieldByName( 'INVID' ).AsString,
       cdsMaster.FieldByName( 'SEQ' ).AsString,
       cdsMaster.FieldByName( 'INVFORMAT' ).AsString,
       FormatDateTime( 'YYYYMMDD', cdsMaster.FieldByName( 'invdate' ).AsDateTime ),
       cdsMaster.FieldByName( 'TAXTYPE' ).AsString,
       cdsMaster.FieldByName( 'SALEAMOUNT' ).AsString,
       aTaxAmount,
       cdsMaster.FieldByName( 'INVAMOUNT' ).AsString,
       FormatDateTime('YYYYMMDD HHMMSS',aSysDate),
       dtmMain.getLoginUser, aAllowanceNo,aPaperNo ] );

    aDataSet.ExecSQL;
  end;
  {--------------------------------------------------------------------}

  {--------------------------------------------------------------------}
  {???s INV014 ??????}
  procedure ApplyUpdateData;
    var aTaxAmount : String;
  begin
    {#6423 ?p?GTaxAmount = null ?n?NtaxAmount????0 ,Jacy?o??A07???O?????D By Kin 2014/02/10}
    aTaxAmount := cdsMaster.FieldByName( 'TAXAMOUNT' ).AsString;
    if aTaxAmount = '' then
    begin
      aTaxAmount := '0';
    end;
    aDataSet.Close;
    aDataSet.SQL.Text := Format(
      ' UPDATE INV014                                                  ' +
      '    SET CUSTID = ''%s'',                                        ' +
      '        PAPERNO = ''%s'',                                       ' +
      '        PAPERDATE = TO_DATE( ''%s'', ''YYYYMMDD'' ),            ' +
      '        BUSINESSID = ''%s'',                                    ' +
      '        YEARMONTH = ''%s'',                                     ' +
      '        INVID = ''%s'',                                         ' +
      '        SEQ = ''%s'',                                           ' +
      '        INVFORMAT = ''%s'',                                     ' +
      '        INVDATE = TO_DATE( ''%s'', ''YYYYMMDD'' ),              ' +
      '        TAXTYPE = ''%s'',                                       ' +
      '        SALEAMOUNT = ''%s'',                                    ' +
      '        TAXAMOUNT = ''%s'',                                     ' +
      '        INVAMOUNT = ''%s'',                                     ' +
      '        UPTTIME = TO_DATE(''%s'',''YYYY/MM/DD HH24:MI:SS''),    ' +
      '        UPTEN = ''%s''                                          ' +
      '  WHERE IDENTIFYID1 = ''%s''                                    ' +
      '    AND IDENTIFYID2 = ''%s''                                    ' +
      '    AND COMPID = ''%s''                                         ' +
      '    AND ALLOWANCENO = ''%s''                                    ',
      [cdsMaster.FieldByName( 'CUSTID' ).AsString,
       aPaperNo,
       FormatDateTime( 'YYYYMMDD', cdsMaster.FieldByName( 'PAPERDATE' ).AsDateTime ),
       cdsMaster.FieldByName( 'BUSINESSID' ).AsString,
       cdsMaster.FieldByName( 'YEARMONTH' ).AsString,
       cdsMaster.FieldByName( 'INVID' ).AsString,
       cdsMaster.FieldByName( 'SEQ' ).AsString,
       cdsMaster.FieldByName( 'INVFORMAT' ).AsString,
       FormatDateTime( 'YYYYMMDD', cdsMaster.FieldByName( 'invdate' ).AsDateTime ),
       cdsMaster.FieldByName( 'TAXTYPE' ).AsString,
       cdsMaster.FieldByName( 'SALEAMOUNT' ).AsString,
       aTaxAmount,
       cdsMaster.FieldByName( 'INVAMOUNT' ).AsString,
       FormatDateTime('YYYYMMDD HHMMSS',aSysDate),
       dtmMain.getLoginUser,
       IDENTIFYID1,
       IDENTIFYID2,
       cdsMaster.FieldByName( 'COMPID' ).AsString,
       cdsMaster.FieldByName( 'ALLOWANCENO' ).AsString] );
    aDataSet.ExecSQL;
    try
      if Assigned( FDataSet ) then
      begin
        if ( not FDataSet.IsEmpty ) then
        begin
          FDataSet.ReadOnly := False;
          FDataSet.Edit;
          FDataSet.FieldByName( 'PAPERDATE' ).AsDateTime := cdsMaster.FieldByName( 'PAPERDATE' ).AsDateTime;
          FDataSet.FieldByName( 'YEARMONTH2' ).AsString := cdsMaster.FieldByName( 'YEARMONTH2' ).AsString;
          FDataSet.FieldByName( 'SALEAMOUNT' ).AsInteger := cdsMaster.FieldByName( 'SALEAMOUNT' ).AsInteger;
          FDataSet.FieldByName( 'TAXAMOUNT' ).AsInteger := cdsMaster.FieldByName( 'TAXAMOUNT' ).AsInteger;
          FDataSet.FieldByName( 'INVAMOUNT' ).AsInteger := cdsMaster.FieldByName( 'INVAMOUNT' ).AsInteger;
          FDataSet.FieldByName( 'PaperNo'   ).AsString  := cdsMaster.FieldByName( 'PaperNo' ).AsString;
          FDataSet.Post;
        end else
        begin
          NewMainDataSet(cdsMaster.FieldByName( 'INVID' ).AsString );

        end;
      end;
    finally
        FDataSet.ReadOnly := True;
    end;

  end;
  {--------------------------------------------------------------------}

  {?R??INV014 ????}
  {--------------------------------------------------------------------}
  procedure ApplyDeleteData;
  begin

    DeleteRecord;
    if ( dtmMain.GetLinkToMIS ) and
       ( cdsMaster.FieldByName( 'SOURCE' ).AsInteger = 1 ) then
      UpdateSoData;
  end;
  {--------------------------------------------------------------------}
 var
  aConnection, aConnection2: TADOConnection;
begin
  Result := False;

  aDataSet := dtmMain.adoComm2;
  aConnection := dtmMain.InvConnection;
  aConnection2 := dtmSO.SOConnection;
  GetSysDate;
  GetAllowance;
  GetPaperNo;
  aConnection.BeginTrans;
  if ( dtmMain.GetLinkToMIS ) then aConnection2.BeginTrans;
  try
    if ( FMode in [dmInsert] ) then
    begin
      ApplyInsertData;
      ApplyInsertDetailData;
      cdsMaster.FieldByName( 'ALLOWANCENO' ).AsString := aAllowanceNo;
      //#5302 ?p?G?s?W?????n?P?B???D?e?? By Kin 2009/09/21
      NewMainDataSet(cdsMaster.FieldByName( 'INVID' ).AsString );
    end else
    if ( FMode in [dmUpdate] ) then
    begin
      ApplyUpdateData;
      DeleteDetailRecord;
      ApplyInsertDetailData;
    end else
    if ( FMode in [dmDelete] ) then
    begin
      aDataSet.Connection := dtmMain.InvConnection;
      ApplyDeleteData;
      DeleteDetailRecord;
      if Assigned( FDataSet ) then
      begin
        if ( not FDataSet.IsEmpty ) then
        begin
          FDataSet.ReadOnly := False;
          FDataSet.Delete;
          FDataSet.ReadOnly := True;
        end;
      end;
    end;
    aConnection.CommitTrans;
    FAllowanceNo := aAllowanceNo;
    FPaperNo := aPaperNo;
    if ( dtmMain.GetLinkToMIS ) then aConnection2.CommitTrans;
    Result := True;
  except
    on E: Exception do
    begin
      TranslateError( E.Message );
      { PKey ????, PM ?n?D???q?T?? }
      if OracleErros.ErrorCode = 1 then
      begin
        aErrMsg := '?i?????????j????????,?????s?s??!';
      end
      else begin
        aErrMsg := E.Message;
      end;
      aConnection.RollbackTrans;
      if ( dtmMain.GetLinkToMIS ) then aConnection2.RollbackTrans;
    end;
  end;
end;

procedure TfrmInvA07_3.btnInsertClick(Sender: TObject);
begin
  Screen.Cursor := crSQLWait;
  try
    DMLModeChange( dmInsert );
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TfrmInvA07_3.cdsDetailNewRecord(DataSet: TDataSet);
begin
//  cdsDetail.FieldByName( 'ALLOWANCENO' ).AsString := cdsMaster.FieldByName( 'ALLOWANCENO' ).AsString;
end;

procedure TfrmInvA07_3.DisableDetailDataSetEvent;
begin

  cdsDetail.AfterDelete := nil;
  cdsDetail.AfterInsert := nil;
  cdsDetail.BeforePost := nil;
  cdsDetail.AfterPost := nil;
  cdsDetail.BeforeInsert := nil;
  cdsDetail.OnNewRecord := nil;
  cdsDetail.BeforeDelete := nil;
end;

procedure TfrmInvA07_3.EnableDetailDataSetEvent;
begin
  cdsDetail.OnNewRecord := cdsDetailNewRecord;
  cdsDetail.AfterDelete := cdsDetailAfterDelete;
  //cdsDetail.BeforeDelete := cdsDetailBeforeDelete;

  {
  cdsDetail.AfterDelete := cdsDetailAfterDelete;
  cdsDetail.BeforePost := nil; //cdsDetailBeforePost; ?w?????n
  cdsDetail.AfterPost := cdsDetailAfterPost;
  cdsDetail.BeforeInsert := cdsDetailBeforeInsert;
  cdsDetail.OnNewRecord := cdsDetailNewRecord;
  cdsDetail.BeforeDelete := cdsDetailBeforeDelete;
  }
end;



procedure TfrmInvA07_3.btnCancelClick(Sender: TObject);
begin
  Screen.Cursor := crSQLWait;
  try
    DMLModeChange( dmCancel );
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TfrmInvA07_3.btnEditClick(Sender: TObject);
  var
    aErrMsg : string;
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

function TfrmInvA07_3.ValidateCanModify(var aErrMsg: String): Boolean;
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
    //#5881 ?W?[?w?????W?????O By Kin 2010/12/14
    //#6001 ?????~???n?????? By Kin 2011/07/13
    {
    if cdsMaster.FieldByName( 'UPLOADFLAG' ).AsString = '1' then
    begin
      aErrMsg := '?w?????W??????,?L?k????!';
      Result := False;
      Exit;
    end;
    }
  finally
    Result := ( aErrMsg = EmptyStr );
  end;
end;

procedure TfrmInvA07_3.cdsMasterNewRecord(DataSet: TDataSet);
begin
  cdsMaster.FieldByName( 'COMPID' ).AsString := dtmMain.getCompID;
  cdsMaster.FieldByName( 'COMPSNAME' ).AsString := dtmMain.getCompName;
  cdsMaster.FieldByName( 'SEQ' ).AsString := '0';
  cdsMaster.FieldByName( 'SOURCE' ).AsInteger := 0;
  cdsMaster.FieldByName( 'SOURCE2' ).AsString := '???????J';

end;

procedure TfrmInvA07_3.btnSaveClick(Sender: TObject);
  var aMsg:String;
begin
  Screen.Cursor := crSQLWait;
  try
    {#5302 ?p?G?????s?????????n?????? By Kin 2009/10/01}
    if dtmMainH.IsDouble( txtPaperNo.Text,
       cdsMaster.FieldByName( 'COMPID' ).AsString,FAllowanceNo ) then
    begin
      if MessageDlg('?????s???????I?O?_?s???H',mtWarning,[mbYes,mbNo],0) = mrNO then
      begin
        txtPaperNo.SetFocus;
        Exit;
      end;
    end;
    {#5302 ???d?????????????B?O?_???j???????o?????B By Kin 2009/10/01}
    aMsg := EmptyStr;
    if CheckINVAMTOver( FAllowanceNo,txtInvId.Text,aMsg) then
    begin
      MessageDlg( '???????B?j???o?????B?I'+#10+#13 + aMsg,mtWarning,[mbYes],0);
      txtInvAmount.SetFocus;
      Exit;
    end;
    DMLModeChange( dmSave );
  finally
    Screen.Cursor := crDefault;
  end;
end;

function TfrmInvA07_3.ValidateMasterDataSet(var aErrMsg: String): Boolean;
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

    if ( Trim( txtPaperDate.Text ) = EmptyStr ) then
    begin
      aErrMsg := '???????????????J?C';
      aFocusControl := txtPaperDate;
      Exit;
    end;
    if ( Trim( txtYearMonth.Text ) = EmptyStr ) then
    begin
      aErrMsg := '?????~?????????J?C';
      aFocusControl := txtYearMonth;
      Exit;
    end;
    if ( Trim( txtInvId.Text ) = EmptyStr ) then
    begin
      aErrMsg := '?o?????X???????J?C';
      aFocusControl := txtInvId;
      Exit;
    end;
    if not CheckInvoiceIdExists( cdsMaster.FieldByName( 'INVID' ).AsString,
      cdsMaster.FieldByName( 'COMPID' ).AsString ) then
    begin
      aErrMsg := '?L???o?????X, ???T?{?C';
      aFocusControl := txtInvId;
      Exit;
    end;
    
    if not isAmtOK then
    begin
      aErrMsg := '?o?????B???????P???B?[?|?B!';
      aFocusControl := txtInvAmount;
      Exit;
    end;
    if dtmMainH.IsDataLocked( cdsMaster.FieldByName( 'YEARMONTH' ).AsString,
       cdsMaster.FieldByName( 'COMPID' ).AsString ) then
    begin
      aErrMsg := '???????????????q?O?w?g???b,?L?k?s??!';
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

function TfrmInvA07_3.GetYearMonth: String;
begin
  Result := EmptyStr;
  if ( cdsMaster.FieldByName( 'YEARMONTH2' ).AsString <> EmptyStr ) then
  begin
    Result :=
      Copy( cdsMaster.FieldByName( 'YEARMONTH2' ).AsString, 1, 4 ) +
      Lpad( Copy( cdsMaster.FieldByName( 'YEARMONTH2' ).AsString, 6, 2 ), 2, '0' );
  end;
end;

function TfrmInvA07_3.IsAmtOK: Boolean;
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

function TfrmInvA07_3.ValidateDetailDataSet(var aErrMsg: String): Boolean;
var
  aBookMark : TBookmark;
  aAmtTotal : Integer;
begin
  Result := False;
  aAmtTotal := 0;
  if cdsDetail.IsEmpty then
  begin
    aErrMsg := '?L????????????';
    Exit;
  end;
  aBookMark := cdsDetail.GetBookmark;
  try
    cdsDetail.DisableControls;
    try
      cdsDetail.First;
      while not cdsDetail.Eof do
      begin
        if (cdsMaster.FieldByName( 'SOURCE' ).AsString = '1') then
        begin
          if ( cdsDetail.FieldByName( 'PaperNo' ).AsString = EmptyStr ) then
          begin
            aErrMsg := '?????????????i?????????j, ?????s???????J???i???O?????j?C';
            aBookMark := cdsDetail.GetBookmark;
            Break;
          end;
          if ( dtmMain.GetLinkToMIS ) then
          begin
            if not dtmSO.CheckAllowacneExists(
                cdsDetail.FieldByName( 'PAPERNO' ).AsString,
                cdsDetail.FieldByName( 'SEQ' ).AsString,
                cdsMaster.FieldByName( 'CUSTID' ).AsString ) then
            begin
              aErrMsg :=Format('???J???????s ?i %s ?j?????????????A?t???i???O?????j?C',
                        [ cdsDetail.FieldByName('PAPERNO').AsString ] ) ;

              aBookMark := cdsDetail.GetBookmark;
              Break;
            end;
          end;
        end;
        aAmtTotal := aAmtTotal + cdsDetail.FieldByName( 'InvAmount' ).AsInteger;
        cdsDetail.Next;
      end;
      cdsDetail.GotoBookmark( aBookMark );
    finally
      cdsDetail.EnableControls;
    end;
  finally
    cdsDetail.FreeBookmark( aBookMark );
  end;
  {??#6476???D,???n???????? For Jacy Mail By Kin 2014/01/16 }
  {
  if ( aAmtTotal <> cdsMaster.FieldByName( 'InvAmount' ).AsInteger ) and ( aErrMsg = EmptyStr ) then
    aErrMsg :=Format('?????????o?????B ?i %d ?j ?`?p?????I', [ aAmtTotal ] );
  }
  Result := ( aErrMsg = EmptyStr );
end;

procedure TfrmInvA07_3.btnReportClick(Sender: TObject);
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

procedure TfrmInvA07_3.txtInvIdExit(Sender: TObject);
var
  aInvFormat, aInvId, aInvDate, aTaxType, aCustId, aBusinessid, aCustName: String;
  aInvAmt, aSaleAmt, aTaxAmt: Integer;

  {------------------------------------------------------------------}
  procedure GetInvDetail;
  var
    aSql:string;
  begin
    //#5285 ?p?G?O???????J?u?n?a?J SO007???????N?n By Kin 2009/09/10
    if cdsMaster.FieldByName( 'Source' ).AsString = '0' then
    begin
      aSql := Format('Select * From INV007         ' +
                      ' Where INVID=''%s''         ' +
                      ' AND IDENTIFYID1 = ''%s''   ' +
                      ' AND IDENTIFYID2 = ''%s''   ' ,
                      [ aInvId ,IDENTIFYID1, IDENTIFYID2 ] );
      gvView.OptionsData.Editing := False;
      gvView.OptionsData.Deleting := False;
    end else
    begin
      aSql := Format('Select * From INV008         ' +
                      ' Where INVID=''%s''         ' +
                      ' AND IDENTIFYID1 = ''%s''   ' +
                      ' AND IDENTIFYID2 = ''%s''   ' ,
                      [ aInvId ,IDENTIFYID1, IDENTIFYID2 ] );

      gvView.OptionsData.Editing := True;
      gvView.OptionsData.Deleting := True;
    end;
    dtmMain.adoComm.Close;
    dtmMain.adoComm.SQL.Text := aSql;
    dtmMain.adoComm.Open;
    dtmMain.adoComm.First;
    cdsDetail.EmptyDataSet;

    try
      cdsDetail.DisableControls;
      while not dtmMain.adoComm.Eof do
      begin
        cdsDetail.Append;
        if cdsMaster.FieldByName( 'Source' ).AsString = '0' then
        begin
          cdsDetail.FieldByName( 'INVAMOUNT' ).AsInteger := dtmMain.adoComm.FieldByName( 'InvAmount' ).AsInteger;
          cdsDetail.FieldByName( 'PaperNo' ).AsString := cdsMaster.FieldByName( 'PaperNo' ).AsString;
          cdsDetail.Post;
          dtmMain.adoComm.Next;
        end else
        begin
          cdsDetail.FieldByName( 'SERVICETYPEID' ).AsString := dtmMain.adoComm.FieldByName( 'SERVICETYPE').AsString;
          cdsDetail.FieldByName( 'PAPERNO' ).AsString := dtmMain.adoComm.FieldByName( 'BILLID' ).AsString;
          cdsDetail.FieldByName( 'SEQ').AsString := dtmMain.adoComm.FieldByName( 'BILLIDITEMNO' ).AsString;
          cdsDetail.FieldByName( 'ITEMID' ).AsString := dtmMain.adoComm.FieldByName( 'ITEMID' ).AsString;
          cdsDetail.FieldByName( 'DESCRIPTION' ).AsString := dtmMain.adoComm.FieldByName( 'DESCRIPTION' ).AsString;
          cdsDetail.FieldByName( 'INVAMOUNT' ).AsInteger := dtmMain.adoComm.FieldByName( 'TOTALAMOUNT' ).AsInteger;
          cdsDetail.Post;
          dtmMain.adoComm.Next;
        end;
      end;
    finally
      cdsDetail.EnableControls;
    end;

  end;
  {---------------------------------------------------------------------------}

begin
  if btnCancel.Focused then Exit;
  if not ( FMode in [dmInsert, dmUpdate] ) then Exit;
  aInvId := cdsMaster.FieldByName( 'INVID' ).AsString;
  if ( aInvId = EmptyStr ) then Exit;
  FInvExists := CheckInvoiceIdExists( aInvId,
    cdsMaster.FieldByName( 'COMPID' ).AsString );
  if not FInvExists then
  begin
    WarningMsg( '?L???o?????X' );
    if ( txtInvId.CanFocus ) then txtInvId.SetFocus;
    Exit;
  end;
  if CheckInvoiceIsObsolete( aInvId, cdsMaster.FieldByName( 'COMPID' ).AsString ) then
  begin
    WarningMsg( '?????o???w???o?C' );
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
  GetInvDetail;
end;

function TfrmInvA07_3.CheckInvoiceIsObsolete(const aInvId,
  aCompId: String): Boolean;
  var aSql : string;
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

procedure TfrmInvA07_3.gvViewCustomDrawCell(Sender: TcxCustomGridTableView;
  ACanvas: TcxCanvas; AViewInfo: TcxGridTableDataCellViewInfo;
  var ADone: Boolean);
var
  aColumn: TcxGridDBBandedColumn;
begin
  if ( cdsDetail.Active ) and
     ( not cdsDetail.IsEmpty ) and
     ( FMode in [dmInsert, dmUpdate] ) then
  begin
    aColumn := TcxGridDBBandedColumn( AViewInfo.Item );
    if ( aColumn.DataBinding.FieldName = 'SERVICETYPEID' ) or
        ( aColumn.DataBinding.FieldName = 'PAPERNO' ) or
        ( aColumn.DataBinding.FieldName = 'SEQ' ) or
        ( aColumn.DataBinding.FieldName = 'ITEMID' ) or
        ( aColumn.DataBinding.FieldName = 'DESCRIPTION' ) or
        ( aColumn.DataBinding.FieldName = 'PAYTYPEID' ) or
        ( aColumn.DataBinding.FieldName = 'PAYTYPEDESC' )
    then
    begin
      ACanvas.Brush.Color := dtmMain.cxGridNotEditStyle.Color;
      //ACanvas.Brush.Color := clWhite;
    end
    else begin
      ACanvas.Brush.Color := clWhite;
    end;
  end;
end;
procedure TfrmInvA07_3.CalcMoney(aKind: Integer);
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

procedure TfrmInvA07_3.txtSaleAmountExit(Sender: TObject);
begin
  if not ( FMode in [dmInsert, dmUpdate] ) then Exit;
  CalcMoney( 1 );
  UpdateDetailInvAmount;
end;

procedure TfrmInvA07_3.txtTaxAmountExit(Sender: TObject);
begin
  if not ( FMode in [dmInsert, dmUpdate] ) then Exit;
  CalcMoney( 2 );
  UpdateDetailInvAmount;
end;

procedure TfrmInvA07_3.txtInvAmountExit(Sender: TObject);
begin
  if not ( FMode in [dmInsert, dmUpdate] ) then Exit;
  CalcMoney( 3 );
  UpdateDetailInvAmount;
end;

procedure TfrmInvA07_3.gvViewINVAMOUNTPropertiesEditValueChanged(
  Sender: TObject);
begin
  gvView.DataController.PostEditingData;
  UpdateMasterInvAmount;
end;

procedure TfrmInvA07_3.btnQuitClick(Sender: TObject);
begin
  Self.Close;
end;

procedure TfrmInvA07_3.btnDeleteClick(Sender: TObject);
var
  aErrMsg : string;
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

function TfrmInvA07_3.ValidateCanDelete(var aErrMsg: String): Boolean;
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
    if cdsMaster.FieldByName( 'UPLOADFLAG' ).AsString = '1' then
    begin
      aErrMsg := '?w?????W??????,?L?k?R??!';
      Result := False;
      Exit;
    end;
  finally
    Result := ( aErrMsg = EmptyStr );
  end;
end;

procedure TfrmInvA07_3.cdsDetailAfterDelete(DataSet: TDataSet);
begin
  UpdateMasterInvAmount;
end;
procedure TfrmInvA07_3.UpdateDetailInvAmount;
begin

  {#5285 ?p?G?????O???????J,?h???????B?n?@??????Detail Grid???I?????B By Kin 2009/09/10}
  if cdsMaster.FieldByName( 'Source' ).AsString = '0' then
  begin
    if cdsDetail.RecordCount =1 then
    begin
      cdsDetail.First;
      cdsDetail.ReadOnly := False;
      cdsDetail.Edit;
      cdsDetail.FieldByName( 'InvAmount' ).AsString := cdsMaster.FieldByName( 'InvAmount' ).AsString;
    end;
  end;
end;

procedure TfrmInvA07_3.UpdateMasterInvAmount;
var aInvAmount : Integer;
  aBookMark : TBookmark;
begin
  aInvAmount := 0;
  try
    Screen.Cursor := crSQLWait;
    try
      cdsDetail.DisableControls;
      aBookMark := cdsDetail.GetBookmark;
      try
        cdsDetail.First;
        while ( not cdsDetail.Eof ) do
        begin
          aInvAmount := aInvAmount + cdsDetail.FieldByName( 'InvAmount' ).AsInteger;
          cdsDetail.Next;
        end;
        cdsDetail.GotoBookmark( aBookMark );
        txtInvAmount.Text := IntToStr( aInvAmount );
        cdsMaster.FieldByName( 'InvAmount' ).AsInteger := aInvAmount;
        txtInvAmountExit( txtInvAmount );
      finally
        cdsDetail.FreeBookmark( aBookMark );
      end;
    finally
     cdsDetail.EnableControls;
    end;
  finally
    Screen.Cursor := crDefault;
  end;

end;
procedure TfrmInvA07_3.txtPaperNoExit(Sender: TObject);
begin
  {#5285 ?p?G?????O???????J,?h?n???????JDetail Grid???????s?? By Kin 2009/09/10}
  if cdsMaster.FieldByName( 'Source' ).AsString = '0' then
  begin
    if cdsDetail.RecordCount = 1 then
    begin
      cdsDetail.ReadOnly := False;
      cdsDetail.Edit;
      cdsDetail.FieldByName('PaperNo').AsString := cdsMaster.FieldByName('PaperNo').AsString;
      cdsDetail.Post;
      cdsDetail.ReadOnly := True;
    end;
  end;
end;
{#5302 ?s?W???D?e????Function By Kin 2009/09/21}
procedure TfrmInvA07_3.NewMainDataSet(const aCondition: String);
  var
    aEvent: TDataSetNotifyEvent;
    aEvent2: TDataSetNotifyEvent;
begin
  aEvent := FDataSet.AfterScroll;
  aEvent2 := FDataSet.OnNewRecord;
  dtmMain.adoComm.Close;
  dtmMain.adoComm.SQL.Text := BuildMasterSQL( aCondition );
  dtmMain.adoComm.Open;
  dtmMain.adoComm.First;
  try
    FDataSet.AfterScroll := nil;
    FDataSet.OnNewRecord := nil;
    FDataSet.DisableControls;
    FDataSet.ReadOnly := False;
    while not dtmMain.adoComm.Eof do
    begin
      FDataSet.Append;
      FDataSet.FieldByName( 'IDENTIFYID1' ).AsString := dtmMain.adoComm.FieldByName( 'IDENTIFYID1' ).AsString;
      FDataSet.FieldByName( 'IDENTIFYID2' ).AsString := dtmMain.adoComm.FieldByName( 'IDENTIFYID2' ).AsString;
      FDataSet.FieldByName( 'COMPID' ).AsString := dtmMain.adoComm.FieldByName( 'COMPID' ).AsString;
      FDataSet.FieldByName( 'COMPSNAME' ).AsString := dtmMain.adoComm.FieldByName( 'COMPSNAME' ).AsString;
      FDataSet.FieldByName( 'CUSTID' ).AsString := dtmMain.adoComm.FieldByName( 'CUSTID' ).AsString;
      FDataSet.FieldByName( 'CUSTNAME' ).AsString := dtmMain.adoComm.FieldByName( 'CUSTSNAME' ).AsString;
      FDataSet.FieldByName( 'PAPERNO' ).AsString := dtmMain.adoComm.FieldByName( 'PAPERNO' ).AsString;
      FDataSet.FieldByName( 'PAPERDATE' ).AsString := dtmMain.adoComm.FieldByName( 'PAPERDATE' ).AsString;
      FDataSet.FieldByName( 'BUSINESSID' ).AsString := dtmMain.adoComm.FieldByName( 'BUSINESSID' ).AsString;
      FDataSet.FieldByName( 'YEARMONTH' ).AsString := dtmMain.adoComm.FieldByName( 'YEARMONTH' ).AsString;
      FDataSet.FieldByName( 'YEARMONTH2' ).AsString := dtmMain.adoComm.FieldByName( 'YEARMONTH2' ).AsString;
      FDataSet.FieldByName( 'INVID' ).AsString := dtmMain.adoComm.FieldByName( 'INVID' ).AsString;
      FDataSet.FieldByName( 'SEQ' ).AsString := dtmMain.adoComm.FieldByName( 'SEQ' ).AsString;
      FDataSet.FieldByName( 'INVFORMAT' ).AsString := dtmMain.adoComm.FieldByName( 'INVFORMAT' ).AsString;
      FDataSet.FieldByName( 'INVDATE' ).AsString := dtmMain.adoComm.FieldByName( 'INVDATE' ).AsString;
      FDataSet.FieldByName( 'TAXTYPE' ).AsString := dtmMain.adoComm.FieldByName( 'TAXTYPE' ).AsString;
      FDataSet.FieldByName( 'SALEAMOUNT' ).AsString := dtmMain.adoComm.FieldByName( 'SALEAMOUNT' ).AsString;
      FDataSet.FieldByName( 'TAXAMOUNT' ).AsString := dtmMain.adoComm.FieldByName( 'TAXAMOUNT' ).AsString;
      FDataSet.FieldByName( 'INVAMOUNT' ).AsString := dtmMain.adoComm.FieldByName( 'INVAMOUNT' ).AsString;
      FDataSet.FieldByName( 'UPTTIME' ).AsString := dtmMain.adoComm.FieldByName( 'UPTTIME' ).AsString;
      FDataSet.FieldByName( 'UPTEN' ).AsString := dtmMain.adoComm.FieldByName( 'UPTEN' ).AsString;
      FDataSet.FieldByName( 'TAXTYPEDESC' ).AsString := GetTaxTypeDesc( dtmMain.adoComm.FieldByName( 'TAXTYPE' ).AsString );
      FDataSet.FieldByName( 'INVFORMATDESC' ).AsString := GetInvFormatDesc( dtmMain.adoComm.FieldByName( 'INVFORMAT' ).AsString );
      FDataSet.FieldByName( 'ALLOWANCENO' ).AsString := dtmMain.adoComm.FieldByName( 'ALLOWANCENO' ).AsString;
      FDataSet.FieldByName( 'SOURCE' ).AsInteger := dtmMain.adoComm.FieldByName( 'SOURCE' ).AsInteger;
      FDataSet.FieldByName( 'SOURCE2' ).AsString := dtmMain.adoComm.FieldByName( 'SOURCE2' ).AsString;
      //#6001 ?W?[?O?_?@?o?P?@?o???] By Kin 2011/07/15
      FDataSet.FieldByName( 'ISOBSOLETE' ).AsString := dtmMain.adoComm.FieldByName( 'ISOBSOLETE' ).AsString;
      FDataSet.FieldByName( 'OBSOLETEID' ).AsString := dtmMain.adoComm.FieldByName( 'OBSOLETEID' ).AsString;
      FDataSet.FieldByName( 'OBSOLETEREASON' ).AsString := dtmMain.adoComm.FieldByName( 'OBSOLETEREASON' ).AsString;
      FDataSet.Post;
      dtmMain.adoComm.Next;
    end;
  finally
    FDataSet.EnableControls;
    FDataSet.ReadOnly := True;
    FDataSet.AfterScroll := aEvent;
    FDataSet.OnNewRecord := aEvent2;
    dtmMain.adoComm.Close;
  end;


end;
{#5302 ?P?_???????B?O?_?w?j???o???`???B By Kin 2009/10/01 }
function TfrmInvA07_3.CheckINVAMTOver(aAllowanceNo,
  aInvId: String;var aMsg: String): Boolean;
  var aSQL:string;
      aTotal:Double ;
      aInvAmount:Double;
begin
  Result := False;
  if aAllowanceNo = EmptyStr then
    aAllowanceNo := ' ';
  try
    aSQL := Format('SELECT INVAMOUNT FROM INV007      ' +
        ' WHERE IDENTIFYID1 = ''%s''                  ' +
        ' AND IDENTIFYID2 = %s                        ' +
        ' AND INVID=''%s''                            ' ,
        [IDENTIFYID1,IDENTIFYID2,aInvId]);
    dtmMain.adoComm.Close;
    dtmMain.adoComm.SQL.Text := aSQL;
    dtmMain.adoComm.Open;
    aInvAmount :=dtmMain.adoComm.FieldByName( 'INVAMOUNT' ).AsFloat;
    dtmMain.adoComm.Close;
    {#6309 ?W?[IsObsolete<>'Y'?????? By Kin 2012/09/07}
    aSQL := Format('SELECT A.INVAMOUNT TOTAL,A.ALLOWANCENO  FROM INV014 A              ' +
          ' WHERE A.IDENTIFYID1 =''%s'' AND A.IDENTIFYID2 = %s                         ' +
          ' AND A.INVID = ''%s'' AND ALLOWANCENO <> ''%s''                             ' +
          ' AND A.ISOBSOLETE <> ''Y''                                                   ',
          [ IDENTIFYID1,IDENTIFYID2,aInvId,aAllowanceNo ]);
    dtmMain.adoComm.Close;
    dtmMain.adoComm.SQL.Text := aSQL;
    dtmMain.adoComm.Open;
    if dtmMain.adoComm.IsEmpty then
    begin
      aTotal := 0;
    end else
    begin
      dtmMain.adoComm.First;
      while not dtmMain.adoComm.Eof do
      begin
        aTotal := aTotal + dtmMain.adoComm.FieldByName( 'TOTAL' ).AsFloat;
        aMsg := Format( aMsg + '???????? : %s [ ???B %s ]' +#10+#13 ,
                      [dtmMain.adoComm.FieldByName( 'ALLOWANCENO').AsString,
                      dtmMain.adoComm.FieldByName( 'TOTAL' ).AsString] );

        dtmMain.adoComm.Next;
      end;
    end;
    dtmMain.adoComm.Close;
    Result := (aTotal + txtInvAmount.Value)>aInvAmount;
  except
    on E: Exception do
    begin
      aMsg := Format('???d?o?????B????,?T??:%s',[e.Message]);
      Result := True;
      Exit;
    end;
  end;


{
  aSQL := Format('SELECT SUM(A.INVAMOUNT) TOTAL , B.INVAMOUNT FROM INV014 A,INV007 B ' +
          ' WHERE A.IDENTIFYID1 =''%s'' AND A.IDENTIFYID2 = ''%s''                   ' +
          ' AND A.INVID = ''%s'' AND ALLOWANCENO <> ''%s''                           ' +
          ' AND B.IDENTIFYID1 = A.IDENTIFYID1 AND B.IDENTIFYID2 = A.IDENTIFYID2      ' +
          ' AND B.INVID = A.INVID GROUP BY B.INVAMOUNT                                 ' ,
          [ IDENTIFYID1,IDENTIFYID2,aInvId,aAllowanceNo ] );
  dtmMain.adoComm.Close;
  dtmMain.adoComm.SQL.Text := aSQL;
  dtmMain.adoComm.Open;
  if not dtmMain.adoComm.IsEmpty then
  begin
    dtmMain.adoComm.First;
    aTotal := dtmMain.adoComm.FieldByName( 'TOTAL' ).AsFloat;
    aInvAmount := dtmMain.adoComm.FieldByName( 'INVAMOUNT' ).AsFloat ;
    dtmMain.adoComm.Close;
    Result := (aTotal + txtInvAmount.Value )> aInvAmount;
  end else
  begin
    aSQL := Format('SELECT INVAMOUNT FROM INV007      ' +
        ' WHERE IDENTIFYID1 = ''%s''                  ' +
        ' AND IDENTIFYID2 = ''%s''                    ' +
        ' AND INVID=''%s''                            ' ,
        [IDENTIFYID1,IDENTIFYID2,aInvId]);
    dtmMain.adoComm.Close;
    dtmMain.adoComm.SQL.Text := aSQL;
    dtmMain.adoComm.Open;
    aInvAmount :=dtmMain.adoComm.FieldByName( 'INVAMOUNT' ).AsFloat;
    dtmMain.adoComm.Close;
    Result := txtInvAmount.Value > aInvAmount;

  end;
}
end;

procedure TfrmInvA07_3.btnDropClick(Sender: TObject);
var
  aResult: Integer;
  aMsg: String;
begin
  if not ValidateCanDrop( aMsg ) then
  begin
    WarningMsg( aMsg );
    Exit;
  end;
  frmInvA07_5 := TfrmInvA07_5.Create( nil );
  try
    frmInvA07_5.DisNo := FAllowanceNo;
    aResult := frmInvA07_5.ShowModal;
  finally
    frmInvA07_5.Free;
    frmInvA07_5 := nil;
  end;
  if ( aResult = mrOk ) then { Reload ????} FormShow( Self );
end;

function TfrmInvA07_3.ValidateCanDrop(var aErrMsg: String): Boolean;
begin
  aErrMsg := EmptyStr;

  Result := True;
  try
    if cdsMaster.IsEmpty then
    begin
      aErrMsg := '?L????????, ???i???@?o?C';
      Exit;
    end;
    if ( cdsMaster.FieldByName( 'ISOBSOLETE' ).AsString = '?O' ) then
    begin
      aErrMsg := '?????????w?@?o?C';
      Exit;
    end;
    if ( cdsMaster.FieldByName( 'UPLOADFLAG' ).AsString <> '1' ) then
    begin
      aErrMsg := '?????????????L?W???~?i?@?o?C';
      Exit;
    end;


  finally
    Result := ( aErrMsg = '' );
    dtmMain.adoComm.Close;
  end;
end;

end.
