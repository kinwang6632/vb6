unit frmInvA02_4U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Provider, DBClient,  ADODB, DB, Grids, DBGrids,
  Buttons, cbDataLookup,
  cxGraphics, cxControls, dxStatusBar, cxSplitter, cxStyles, cxCustomData,
  cxFilter, cxData, cxDataStorage, cxEdit, cxDBData, cxGridCustomView,
  cxGridCustomTableView, cxGridTableView, cxGridBandedTableView,
  cxGridDBBandedTableView, cxClasses, cxGridLevel, cxGrid,
  cxCalendar, cxSpinEdit, cxCurrencyEdit, cxLookAndFeelPainters,
  cxButtons, cxDBLabel, cxContainer, cxLabel, cxTextEdit, cxMaskEdit,
  cxButtonEdit, cxDBEdit, cxDropDownEdit, cxGroupBox, cxRadioGroup,
  cxPropertiesStore, cxLookupEdit, cxDBLookupEdit, cxDBExtLookupComboBox,
  Menus, cxGridDBTableView, dxSkinsCore, dxSkinsDefaultPainters,
  dxSkinscxPCPainter, dxSkinsdxBarPainter, dxSkinsdxStatusBarPainter,
  dxSkinBlack, dxSkinBlue, dxSkinCaramel, dxSkinCoffee;

type

  TSubclassPopupEdit = class(TcxPopupEdit)
  public
    property PopupWindow;
  end;

  TSubclassPopupEditPopupWindow = class(TcxPopupEditPopupWindow)
  public
    procedure PopEditWindowKeyPress(Sender: TObject; var Key: Char);
  end;

  TDMLMode = ( dmBrowser, dmInsert, dmUpdate, dmSave, dmCancel, dmDelete, dmLock );

  TInvCreateType = ( icPrev = 1,       // �w�}, MIS ����, �w�}
                     icAfter = 2,      // ��}, MIS ����, ��}
                     icLocale = 3,     // �{���}��, Invoice Create �w�}
                     icNormal = 4 );   // �@��}��, Invoice Create ��}

  TChargeColumn = ( chQuantity, chUnitPrice, chSaleAmount, chTotalAmount, chTaxAmount );

  TBillingInfo = record
    BillId: String;                  { ���O�渹 }
    ItemNo: String;                  { ���� }
  end;

  TTaxInfo = record
    TaxCode: String;                 { �|�O }
    TaxRate: String;                 { �|�v }
  end;

  TChooseInvInfo = record            { ��ܪ��o���O�� }
    YearMonth: String;
    InvFormat: String;
    Prefix: String;
    StartNum: String;
    OldInvDate: String;
  end;


  TfrmInvA02_4 = class(TForm)
    Panel1: TPanel;
    dxStatusBar1: TdxStatusBar;
    ScrollBox1: TScrollBox;
    ScrollBox2: TScrollBox;
    ScrollBox3: TScrollBox;
    Bevel1: TBevel;
    DetailGridLevel1: TcxGridLevel;
    DetailGridLevel2: TcxGridLevel;
    DetailGrid: TcxGrid;
    DetailGridView: TcxGridDBBandedTableView;
    cdsMaster: TClientDataSet;
    dspMaster: TDataSetProvider;
    cdsDetail: TClientDataSet;
    dsMaster: TDataSource;
    dsDetail: TDataSource;
    DetailGridViewBILLID: TcxGridDBBandedColumn;
    DetailGridViewBILLIDITEMNO: TcxGridDBBandedColumn;
    DetailGridViewSEQ: TcxGridDBBandedColumn;
    DetailGridViewSTARTDATE: TcxGridDBBandedColumn;
    DetailGridViewENDDATE: TcxGridDBBandedColumn;
    DetailGridViewDESCRIPTION: TcxGridDBBandedColumn;
    DetailGridViewQUANTITY: TcxGridDBBandedColumn;
    DetailGridViewUNITPRICE: TcxGridDBBandedColumn;
    DetailGridViewTAXAMOUNT: TcxGridDBBandedColumn;
    DetailGridViewTOTALAMOUNT: TcxGridDBBandedColumn;
    DetailGridViewCHARGEEN: TcxGridDBBandedColumn;
    DetailGridViewSERVICETYPE: TcxGridDBBandedColumn;
    DetailGridViewSALEAMOUNT: TcxGridDBBandedColumn;
    btnLocaleCreate: TcxButton;
    btnNormalCreate: TcxButton;
    btnEdit: TcxButton;
    btnCancel: TcxButton;
    btnSave: TcxButton;
    btnQuit: TcxButton;
    cxLabel1: TcxLabel;
    cxLabel3: TcxLabel;
    cxLabel2: TcxLabel;
    txtCUSTID: TcxDBButtonEdit;
    cxLabel4: TcxLabel;
    txtCUSTSNAME: TcxDBTextEdit;
    cxLabel5: TcxLabel;
    txtINVTITLE: TcxDBTextEdit;
    cxLabel6: TcxLabel;
    txtBUSINESSID: TcxDBTextEdit;
    txtCHARGEDATE: TcxDBDateEdit;
    cxLabel7: TcxLabel;
    cxLabel8: TcxLabel;
    txtMAILADDR: TcxDBTextEdit;
    cxLabel9: TcxLabel;
    txtZIPCODE: TcxDBTextEdit;
    cxLabel10: TcxLabel;
    txtMEMO1: TcxDBTextEdit;
    cxLabel11: TcxLabel;
    txtMEMO2: TcxDBTextEdit;
    Shape1: TShape;
    cxSplitter1: TcxSplitter;
    cxLabel12: TcxLabel;
    lblSALEAMOUNT: TcxDBLabel;
    cxLabel13: TcxLabel;
    lblTAXAMOUNT: TcxDBLabel;
    cxLabel14: TcxLabel;
    lblINVAMOUNT: TcxDBLabel;
    cxLabel15: TcxLabel;
    lblISOBSOLETEDESC: TcxDBLabel;
    cxLabel16: TcxLabel;
    lblOBSOLETEREASON: TcxDBLabel;
    Shape2: TShape;
    cxLabel17: TcxLabel;
    lblCANMODIFYDESC: TcxDBLabel;
    cxLabel18: TcxLabel;
    lblINVFORMATDESC: TcxDBLabel;
    cxLabel19: TcxLabel;
    lblHOWTOCREATEDESC: TcxDBLabel;
    grpTaxRate: TcxGroupBox;
    cxLabel20: TcxLabel;
    rdoTAXRATE1: TcxRadioButton;
    rdoTAXRATE2: TcxRadioButton;
    rdoTAXRATE3: TcxRadioButton;
    cxLabel21: TcxLabel;
    DetailGridViewSIGN: TcxGridDBBandedColumn;
    dsChooseInv: TDataSource;
    txtINVDATE: TcxDBPopupEdit;
    ChooseGridLevel1: TcxGridLevel;
    ChooseGrid: TcxGrid;
    ChooseInvView: TcxGridDBBandedTableView;
    ChooseInvViewYEARMONTH: TcxGridDBBandedColumn;
    ChooseInvViewINVFORMAT: TcxGridDBBandedColumn;
    ChooseInvViewPREFIX: TcxGridDBBandedColumn;
    ChooseInvViewSTARTNUM: TcxGridDBBandedColumn;
    ChooseInvViewENDNUM: TcxGridDBBandedColumn;
    ChooseInvViewCURNUM: TcxGridDBBandedColumn;
    ChooseInvViewLASTINVDATE: TcxGridDBBandedColumn;
    ChooseInvViewMEMO: TcxGridDBBandedColumn;
    cxLabel22: TcxLabel;
    txtINVADDR: TcxDBTextEdit;
    DetailGridViewTAXNAME: TcxGridDBBandedColumn;
    cxLabel23: TcxLabel;
    lblPRINTCOUNT: TcxDBLabel;
    DetailGridViewEMPTY: TcxGridDBBandedColumn;
    DetailGridViewITEMID: TcxGridDBBandedColumn;
    dsChooseItem: TDataSource;
    ChooseItemView: TcxGridDBBandedTableView;
    ChooseItemViewITEMID: TcxGridDBBandedColumn;
    ChooseItemViewDESCRIPTION: TcxGridDBBandedColumn;
    ChooseItemViewTAXNAME: TcxGridDBBandedColumn;
    cdsSO: TClientDataSet;
    dspSO: TDataSetProvider;
    txtTAXRATE: TcxDBCurrencyEdit;
    btnPrint: TcxButton;
    btnDrop: TcxButton;
    cxButton1: TcxButton;
    cxLabel24: TcxLabel;
    cxLabel25: TcxLabel;
    cxLabel26: TcxLabel;
    cxLabel27: TcxLabel;
    lblMainSALEAMOUNT: TcxDBLabel;
    lblMainTAXAMOUNT: TcxDBLabel;
    lblMainINVAMOUNT: TcxDBLabel;
    txtMainInvId: TcxDBButtonEdit;
    DetailGridViewLinkToMIS: TcxGridDBBandedColumn;
    txtINVID: TcxDBTextEdit;
    cdsInv008A: TClientDataSet;
    INV008View: TcxGridDBTableView;
    dsInv008: TDataSource;
    INV008ViewBILLID: TcxGridDBColumn;
    INV008ViewBILLIDITEMNO: TcxGridDBColumn;
    INV008ViewSEQ: TcxGridDBColumn;
    INV008ViewSTARTDATE: TcxGridDBColumn;
    INV008ViewENDDATE: TcxGridDBColumn;
    INV008ViewITEMID: TcxGridDBColumn;
    INV008ViewDESCRIPTION: TcxGridDBColumn;
    INV008ViewQUANTITY: TcxGridDBColumn;
    INV008ViewUNITPRICE: TcxGridDBColumn;
    INV008ViewTAXAMOUNT: TcxGridDBColumn;
    INV008ViewSALEAMOUNT: TcxGridDBColumn;
    INV008ViewTOTALAMOUNT: TcxGridDBColumn;
    INV008ViewFACISNO: TcxGridDBColumn;
    INV008ViewACCOUNTNO: TcxGridDBColumn;
    SpeedButton2: TSpeedButton;
    DetailGridViewFacisNo: TcxGridDBBandedColumn;
    DetailGridViewAccountNo: TcxGridDBBandedColumn;
    cxLabel28: TcxLabel;
    txtInstAddr: TcxDBTextEdit;
    cxLabel29: TcxLabel;
    txtInvUse: TDataLookup;
    Shape3: TShape;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure btnLocaleCreateClick(Sender: TObject);
    procedure btnNormalCreateClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnQuitClick(Sender: TObject);
    procedure cdsMasterAfterScroll(DataSet: TDataSet);
    procedure txtCUSTIDPropertiesButtonClick(Sender: TObject;
      AButtonIndex: Integer);
    procedure txtINVDATEPropertiesInitPopup(Sender: TObject);
    procedure txtCHARGEDATEPropertiesInitPopup(Sender: TObject);
    procedure rdoTAXRATE1Click(Sender: TObject);
    procedure txtINVDATEPropertiesValidate(Sender: TObject;
      var DisplayValue: Variant; var ErrorText: TCaption;
      var Error: Boolean);
    procedure txtINVDATEPropertiesCloseUp(Sender: TObject);
    procedure ChooseInvViewDblClick(Sender: TObject);
    procedure ChooseInvViewKeyPress(Sender: TObject; var Key: Char);
    procedure cdsMasterNewRecord(DataSet: TDataSet);
    procedure cdsDetailAfterPost(DataSet: TDataSet);
    procedure txtCUSTIDPropertiesChange(Sender: TObject);
    procedure DetailGridViewBILLIDPropertiesValidate(Sender: TObject;
      var DisplayValue: Variant; var ErrorText: TCaption;
      var Error: Boolean);
    procedure txtCHARGEDATEPropertiesValidate(Sender: TObject;
      var DisplayValue: Variant; var ErrorText: TCaption;
      var Error: Boolean);
    procedure DetailGridViewBILLIDITEMNOPropertiesValidate(Sender: TObject;
      var DisplayValue: Variant; var ErrorText: TCaption;
      var Error: Boolean);
    procedure DetailGridViewSTARTDATEPropertiesValidate(Sender: TObject;
      var DisplayValue: Variant; var ErrorText: TCaption;
      var Error: Boolean);
    procedure DetailGridViewENDDATEPropertiesValidate(Sender: TObject;
      var DisplayValue: Variant; var ErrorText: TCaption;
      var Error: Boolean);
    procedure DetailGridViewBILLIDPropertiesButtonClick(Sender: TObject;
      AButtonIndex: Integer);
    procedure cdsDetailNewRecord(DataSet: TDataSet);
    procedure cdsDetailBeforeInsert(DataSet: TDataSet);
    procedure cdsDetailAfterDelete(DataSet: TDataSet);
    procedure txtCUSTIDExit(Sender: TObject);
    procedure txtBUSINESSIDExit(Sender: TObject);
    procedure DetailGridViewCustomDrawCell(Sender: TcxCustomGridTableView;
      ACanvas: TcxCanvas; AViewInfo: TcxGridTableDataCellViewInfo;
      var ADone: Boolean);
    procedure DetailGridViewEditing(Sender: TcxCustomGridTableView;
      AItem: TcxCustomGridTableItem; var AAllow: Boolean);
    procedure DetailGridViewITEMIDPropertiesInitPopup(Sender: TObject);
    procedure ChooseItemViewDblClick(Sender: TObject);
    procedure DetailGridViewITEMIDPropertiesCloseUp(Sender: TObject);
    procedure DetailGridViewITEMIDPropertiesValidate(Sender: TObject;
      var DisplayValue: Variant; var ErrorText: TCaption;
      var Error: Boolean);
    procedure DetailGridViewITEMIDPropertiesEditValueChanged(
      Sender: TObject);
    procedure DetailGridViewEditKeyDown(Sender: TcxCustomGridTableView;
      AItem: TcxCustomGridTableItem; AEdit: TcxCustomEdit; var Key: Word;
      Shift: TShiftState);
    procedure DetailGridViewQUANTITYPropertiesEditValueChanged(
      Sender: TObject);
    procedure DetailGridViewTOTALAMOUNTPropertiesEditValueChanged(
      Sender: TObject);
    procedure DetailGridViewQUANTITYPropertiesValidate(Sender: TObject;
      var DisplayValue: Variant; var ErrorText: TCaption;
      var Error: Boolean);
    procedure DetailGridViewSALEAMOUNTPropertiesEditValueChanged(
      Sender: TObject);
    procedure DetailGridViewTAXAMOUNTPropertiesValidate(Sender: TObject;
      var DisplayValue: Variant; var ErrorText: TCaption;
      var Error: Boolean);
    procedure txtTAXRATEExit(Sender: TObject);
    procedure txtINVDATEExit(Sender: TObject);
    procedure DetailGridExit(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure DetailGridViewTAXAMOUNTPropertiesEditValueChanged(
      Sender: TObject);
    procedure DetailGridViewUNITPRICEPropertiesEditValueChanged(
      Sender: TObject);
    procedure btnDropClick(Sender: TObject);
    procedure txtMainInvIdPropertiesButtonClick(Sender: TObject;
      AButtonIndex: Integer);
    procedure cdsDetailBeforePost(DataSet: TDataSet);
    procedure DetailGridViewLinkToMISPropertiesEditValueChanged(
      Sender: TObject);
    procedure cdsDetailBeforeDelete(DataSet: TDataSet);
    procedure DetailGridViewBILLIDPropertiesEditValueChanged(
      Sender: TObject);
    procedure txtInvUseCodeNamePropertiesInitPopup(Sender: TObject);
  private
    { Private declarations }
    FMode: TDMLMode;
    FInvCreateType: TInvCreateType;
    FInvoiceId: String;
    FDetailMaxSequence: Integer;
    FBillingInfo: TBillingInfo;
    FTaxInfo: TTaxInfo;
    FChargeWindow: TForm;
    FTitleWindow: TForm;
    FSearchCustWindow: TForm;
    FRelationInvWindow: TForm;
    { �ΨӰO���̫�@�i�o�����X, �u�ΨӦs�ɫ�, ��������̫�@���� }
    FMainInvId: String;
    { �ΨӰO���o�����, �Y User ��Τ⥴, �����P�_�������� }
    FChooseInv: TChooseInvInfo;
    // �O�_�O�h�i�۰ʴ��}���o��
    FMultiInv: Boolean;
    {}
    FCallFromProgram: TClass;
    {}
    function DataSetStateChange(const aMode: TDMLMode): Boolean;
    function ValidateCanModify(var aErrMsg: String): Boolean;
    function ValidateCanDrop(var aErrMsg: String): Boolean;
    function ValidateCanPrint(var aErrMsg: String): Boolean;
    function ValidateMasterDataSet(var aErrMsg: String): Boolean;
    function ValidateDetailDataSet(var aErrMsg: String): Boolean;
    function ApplyDataSet(var aErrMsg: String): Boolean;
    function ValidateCustomer(const aCustId: String; aInvCreateType: TInvCreateType): Boolean;
    function ValidateBillingNo(const aBillId, aItemId: String;  var aErrMsg: String;
      const aIsSaved: Boolean = False): Boolean;
    function ValidateChargeItemId(const aItemId: String; var aErrMsg: String): Boolean;
    function BuildMasterSQL(const aInvId: String): String;
    function BuildDetailSQL(const aInvId: String): String;
    function BuildINV008ASQL(const aInvId: String): String;
    function BuildMasterApplySQL(const aMode: TDMLMode): String;
    function BuildDetailApplySQL(const aMode: TDMLMode): String;
    function BuildInv008AApplySQL(const aMode: TDMLMode): String;
    function DetailSequenceNeeded: Integer;
    function PopChargeWindow(const aCustId: String): TModalResult;
    function PopTitleWindow(const aCustId: String): TModalResult;
    function PopTitleWindow2(const aCustId: String): TModalResult;
    function PopRelationInvId(const aMainInvId: String): TModalResult;
    function GetItemChargeSign(const aItemId: String): String;
    function GetMasterTaxCode: String;
    function GetInvoiceCheckNumber: String;
    function GetCurrentAvailableInvCount: Integer;
    function GetNextInvNumber(const aInvId: String): String;
    function ReverseInvStartNum(const aChoose: TChooseInvInfo): String;
    function ValidateCanSetInvUse(const aInvUseId: String): Boolean;
    procedure DisableDetailDataSetEvent;
    procedure EnableDetailDataSetEvent;
    procedure PrepareDetailDataSet;
    procedure PrepareInv008ADataSet;
    procedure FillDetailDataSet;
    procedure FillInv008ADataSet;
    procedure FillChooseInv;
    procedure ComposeBillingString(var aBillInfo: TBillingInfo);
    procedure GetCustomerInfo(const aCustId: String);
    procedure GetChargeItemInfo(const aItemId: String);
    procedure CalculateChargePrice(const aCalculateColum: TChargeColumn);
    procedure CopyChargeDataToDetail;
    procedure CopyChargeTitleToMaster(const aInvCreateType: TInvCreateType);
    procedure GridStateChange(const aMode: TDMLMode);
    procedure ButtonStateChange(const aMode: TDMLMode);
    procedure EditorStateChange(const aMode: TDMLMode);
    procedure FillDetailColumnPropertiesValue;
    procedure SetMasterFieldState;
    procedure SetDetailFieldState;
    procedure MasterDataToEditor;
    procedure UpdatePriceToMaster;
    procedure UpdateTaxToMaster;
    procedure UpdateTaxToDetail;
    procedure UpdateDeatilPrice;
    procedure UpdateDetailPriceByChargeSign;
    procedure UpdateDetailSourceByBillingNo;
    procedure UpdateMasterPrintCount;
    procedure UpdateDetailMisInfo;
    {}
    function IsLinkToMis(const aInvId, aSeq: String): Boolean;
    function IsMergeRecord(const aInvId, aSeq: String): Boolean;
    function SyncInv008ADataSet(const aMode: TDMLMode; const aOldSeq, aNewSeq, aNewInvId: String;
      var aMsg: String): Boolean;
  public
    { Public declarations }
    constructor Create(AInvoiceId: String); reintroduce;
    procedure DMLModeChange(const aMode: TDMLMode);
    property CallFromProgram: TClass read FCallFromProgram write FCallFromProgram;
  end;

var
  frmInvA02_4: TfrmInvA02_4;

implementation

uses cbUtilis, Uotheru, frmMainU, dtmMainU, dtmMainHU, dtmMainJU, dtmSOU,
     frmInvA02U, frmInvA02_1U, frmInvA02_5U, Math, frmInvA05_1U, frmInvB02U,
     frmSearchCustU, frmInvA02_6U;

{$R *.dfm}


{ ---------------------------------------------------------------------------- }

{ TSubclassPopupEditPopupWindow }

procedure TSubclassPopupEditPopupWindow.PopEditWindowKeyPress(Sender: TObject; var Key: Char);
begin
  if Ord( Key ) = VK_RETURN then
    frmInvA02_4.ChooseItemViewDblClick( frmInvA02_4.ChooseItemView );
end;

{ ---------------------------------------------------------------------------- }

{ String Code Convert To Enum Type --> TInvCreateType }

function StrToInvoicType(const aValue: String): TInvCreateType;
begin
  Result := icNormal;
  if aValue = '1' then
    Result := icPrev
  else if aValue = '2' then
    Result := icAfter
  else if aValue = '3' then
    Result := icLocale
  else if aValue = '4' then
    Result := icNormal
end;

{ ---------------------------------------------------------------------------- }

// �N HowToCreate �X, �ন SO ���O Table ���N�X 
// SO ���O Table ���N�X   0:��}   1:�w�}  2:�{���}��

function InvoiceTypeToSO(aType: TInvCreateType): String;
begin
  Result := EmptyStr;
  case aType of
    icPrev: Result := '1';
    icAfter: Result := '0';
    icLocale: Result := '2';
  end;
end;

{ ---------------------------------------------------------------------------- }

{ TfrmInvA02_4 }

constructor TfrmInvA02_4.Create(AInvoiceId: String);
begin
  inherited Create( Application );
  FInvoiceId := AInvoiceId;
  FBillingInfo.BillId := EmptyStr;
  FBillingInfo.ItemNo := EmptyStr;
  FChooseInv.YearMonth := EmptyStr;
  FChooseInv.InvFormat := EmptyStr;
  FChooseInv.Prefix := EmptyStr;
  FChooseInv.StartNum := EmptyStr;
  FChooseInv.OldInvDate := EmptyStr;
  FMainInvId := EmptyStr;
  FMultiInv := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.FormCreate(Sender: TObject);
begin
  frmInvA02_4 := Self;
  FMode := dmBrowser;
  FInvCreateType := icLocale;
  DetailGridLevel2.Visible := False;
  txtInvUse.Connection := dtmMain.InvConnection;
  txtInvUse.CodeNoFieldName := 'ITEMID';
  txtInvUse.DescriptionFieldName := 'DESCRIPTION';
  txtInvUse.Initializa;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.FormShow(Sender: TObject);
var
  aEvent: TDataSetNotifyEvent;
  aMode: TDMLMode;
begin
  if ( FCallFromProgram = TfrmMain ) then
  begin
    Caption := frmMain.GetFormTitleString( 'A0C', '�o���d��' );
    aMode := dmLock;
  end else
  begin
    Caption := frmMain.GetFormTitleString( 'A02', '�}�ߵo��' );
    aMode := dmBrowser;
  end;  
  Self.Update;
  Screen.Cursor := crSQLWait;
  try
    Application.ProcessMessages;
    aEvent := cdsMaster.AfterScroll;
    cdsMaster.AfterScroll := nil;
    try
      cdsMaster.Close;
      cdsMaster.CommandText := BuildMasterSQL( FInvoiceId );
      cdsMaster.Open;
      SetMasterFieldState;
      Application.ProcessMessages;
      DMLModeChange( aMode );
    finally
      cdsMaster.AfterScroll := aEvent;
    end;
    Application.ProcessMessages;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  if FMode in [dmInsert, dmUpdate] then
  begin
    CanClose := ConfirmMsg( '�����}�ߵo���@�~?' );
    if CanClose then
      DMLModeChange( dmCancel );
  end;
  if CanClose then
  begin
    dsChooseInv.DataSet.Close;
    dsChooseItem.DataSet.Close;
    dspMaster.DataSet.Close;
    cdsMaster.AfterScroll := nil;
    cdsMaster.Close;
    cdsDetail.EmptyDataSet;
    cdsSO.Close;
    txtInvUse.Finaliza;
  end;    
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.DMLModeChange(const aMode: TDMLMode);
begin
  if not DataSetStateChange( aMode ) then Exit;
  FMode := aMode;
  if ( aMode in [dmSave, dmCancel] ) then
    FMode := dmBrowser;
  GridStateChange( FMode );  
  ButtonStateChange( FMode );
  EditorStateChange( FMode );
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA02_4.BuildMasterSQL(const aInvId: String): String;
begin
  Result := Format(
    '  SELECT A.IDENTIFYID1,                      ' +
    '         A.IDENTIFYID2,                      ' +
    '         A.CHECKNO,                          ' +
    '         A.INVID,                            ' +
    '         A.INVDATE,                          ' +
    '         A.CHARGEDATE,                       ' +
    '         A.COMPID,                           ' +
    '         A.CUSTID,                           ' +
    '         A.CUSTSNAME,                        ' +
    '         A.INVTITLE,                         ' +
    '         A.ZIPCODE,                          ' +
    '         A.INVADDR,                          ' +
    '         A.MAILADDR,                         ' +
    '         A.BUSINESSID,                       ' +
    '         A.TAXTYPE,                          ' +
    '         A.TAXRATE,                          ' +
    '         A.SALEAMOUNT,                       ' +
    '         A.TAXAMOUNT,                        ' +
    '         A.INVAMOUNT,                        ' +
    '         A.INVFORMAT,                        ' +
    '         DECODE( A.INVFORMAT,                ' +
    '                 ''1'', ''�q�l'',            ' +
    '                 ''2'', ''��G'',            ' +
    '                 ''3'', ''��T'' ) INVFORMATDESC, ' +
    '         DECODE( A.ISOBSOLETE,               ' +
    '                ''Y'', ''�O'',               ' +
    '                ''�_'' ) ISOBSOLETEDESC,     ' +
    '         A.OBSOLETEREASON,                   ' +
    '         A.CANMODIFY,                        ' +
    '         DECODE( A.CANMODIFY,                ' +
    '                 ''Y'', ''�O'',              ' +
    '                 ''�_'' ) CANMODIFYDESC,     ' +
    '         A.HOWTOCREATE,                      ' +
    '         DECODE( A.HOWTOCREATE,              ' +
    '                 ''1'', ''�w�}'',            ' +
    '                 ''2'', ''��}'',            ' +
    '                 ''3'', ''�{���}��'',        ' +
    '                 ''4'', ''�@��}��'' ) HOWTOCREATEDESC, ' +
    '         NVL( A.PRINTCOUNT, 0 ) PRINTCOUNT,  ' +
    '         A.MEMO1,                            ' +
    '         A.MEMO2,                            ' +
    '         A.MAININVID,                        ' +
    '         A.MAINSALEAMOUNT,                   ' +
    '         A.MAINTAXAMOUNT,                    ' +
    '         A.MAININVAMOUNT,                    ' +
    '         A.INSTADDR,                         ' +
    '         A.INVUSEID,                         ' +
    '         A.INVUSEDESC                        ' +
    '    FROM INV007 A, INV028 B                  ' +
    '   WHERE A.IDENTIFYID1 = ''%s''              ' +
    '     AND A.IDENTIFYID2 = %s                  ' +
    '     AND A.INVID = ''%s''                    ' +
    '     AND A.COMPID = ''%s''                   ',
    [IDENTIFYID1, IDENTIFYID2, aInvId, dtmMain.getCompID] );
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA02_4.BuildDetailSQL(const aInvId: String): String;
begin
  Result := Format(
    '  SELECT A.IDENTIFYID1,                       ' +
    '         A.IDENTIFYID2,                       ' +
    '         A.INVID,                             ' +
    '         A.BILLID,                            ' +
    '         A.BILLIDITEMNO,                      ' +
    '         A.SEQ,                               ' +
    '         A.STARTDATE,                         ' +
    '         A.ENDDATE,                           ' +
    '         A.ITEMID,                            ' +
    '         A.DESCRIPTION,                       ' +
    '         A.QUANTITY,                          ' +
    '         A.UNITPRICE,                         ' +
    '         A.SALEAMOUNT,                        ' +
    '         A.TAXAMOUNT,                         ' +
    '         A.TOTALAMOUNT,                       ' +
    '         A.CHARGEEN,                          ' +
    '         A.UPTTIME,                           ' +
    '         A.UPTEN,                             ' +
    '         A.SERVICETYPE,                       ' +
    '         NVL( A.LINKTOMIS, ''Y'' ) AS LINKTOMIS, ' +
    '         DECODE( NVL( A.LINKTOMIS, ''Y'' ),   ' +
    '           ''Y'', ''�O'', ''�_'' ) AS LINKTOMISDESC, ' +
    '         B.TAXCODE,                           ' +
    '         B.TAXNAME,                           ' +
    '         B.SIGN,                              ' +
    '         '' '' AS SOURCE,                     ' +
    '         DECODE( A.INVID, NULL,               ' +
    '                          ''Y'',              ' +
    '                          ''N'' ) EMPTY,      ' +
    '         A.FACISNO,                           ' +
    '         A.ACCOUNTNO                          ' +
    '    FROM INV008 A,                            ' +
    '         INV005 B                             ' +
    '    WHERE A.IDENTIFYID1 = B.IDENTIFYID1(+)    ' +
    '      AND A.IDENTIFYID2 = B.IDENTIFYID2(+)    ' +
    '      AND B.COMPID(+) = ''%s''                ' +
    '      AND A.ITEMID = B.ITEMID(+)              ' +
    '      AND A.IDENTIFYID1 = ''%s''              ' +
    '      AND A.IDENTIFYID2 = %s                  ' +
    '      AND A.INVID = ''%s''                    ' +
    '    ORDER BY A.SEQ                            ',
    [ dtmMain.getCompID, IDENTIFYID1, IDENTIFYID2, aInvId ] );
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA02_4.BuildINV008ASQL(const aInvId: String): String;
begin
  Result := Format(
    ' SELECT A.ITEMIDREF ,            ' +
    '        A.INVID,                 ' +
    '        A.BILLID,                ' +
    '        A.BILLIDITEMNO,          ' +
    '        A.SEQ,                   ' +
    '        A.STARTDATE,             ' +
    '        A.ENDDATE,               ' +
    '        A.ITEMID,                ' +
    '        A.DESCRIPTION,           ' +
    '        A.QUANTITY,              ' +
    '        A.UNITPRICE,             ' +
    '        A.TAXAMOUNT,             ' +
    '        A.SALEAMOUNT,            ' +
    '        A.TOTALAMOUNT,           ' +
    '        A.SERVICETYPE,           ' +
    '        A.FACISNO,               ' +
    '        A.ACCOUNTNO              ' +
    '   FROM INV008A A                ' +
    '  WHERE A.INVID = ''%s''         ' +
    '  ORDER By A.SEQ                 ', [aInvId] );
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA02_4.BuildInv008AApplySQL(const aMode: TDMLMode): String;
begin
  if ( aMode in [dmInsert] ) then
  begin
    Result :=
      '  INSERT INTO INV008A        ' +
      '       ( ITEMIDREF,          ' +
      '         INVID,              ' +
      '         BILLID,             ' +
      '         BILLIDITEMNO,       ' +
      '         SEQ,                ' +
      '         STARTDATE,          ' +
      '         ENDDATE,            ' +
      '         ITEMID,             ' +
      '         DESCRIPTION,        ' +
      '         QUANTITY,           ' +
      '         UNITPRICE,          ' +
      '         TAXAMOUNT,          ' +
      '         SALEAMOUNT,         ' +
      '         TOTALAMOUNT,        ' +
      '         SERVICETYPE,        ' +
      '         FACISNO,            ' +
      '         ACCOUNTNO  )        ' +
      ' VALUES ( :ITEMIDREF,        ' +
      '         :INVID,             ' +
      '         :BILLID,            ' +
      '         :BILLIDITEMNO,      ' +
      '         :SEQ,               ' +
      '         TO_DATE( :STARTDATE, ''YYYY/MM/DD'' ), ' +
      '         TO_DATE( :ENDDATE, ''YYYY/MM/DD'' ), ' +
      '         :ITEMID,            ' +
      '         :DESCRIPTION,       ' +
      '         :QUANTITY,          ' +
      '         :UNITPRICE,         ' +
      '         :TAXAMOUNT,         ' +
      '         :SALEAMOUNT,        ' +
      '         :TOTALAMOUNT,       ' +
      '         :SERVICETYPE,       ' +
      '         :FACISNO,           ' +
      '         :ACCOUNTNO  )       ';
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA02_4.BuildMasterApplySQL(const aMode: TDMLMode): String;
begin
  if ( aMode in [dmInsert] ) then
    Result :=
      ' INSERT INTO INV007             ' +
      '        ( IDENTIFYID1,          ' +
      '          IDENTIFYID2,          ' +
      '          CHECKNO,              ' +
      '          INVID,                ' +
      '          CANMODIFY,            ' +
      '          INVDATE,              ' +
      '          CHARGEDATE,           ' +
      '          COMPID,               ' +
      '          CUSTID,               ' +
      '          CUSTSNAME,            ' +
      '          INVTITLE,             ' +
      '          ZIPCODE,              ' +
      '          INVADDR,              ' +
      '          MAILADDR,             ' +
      '          BUSINESSID,           ' +
      '          INVFORMAT,            ' +
      '          TAXTYPE,              ' +
      '          TAXRATE,              ' +
      '          SALEAMOUNT,           ' +
      '          TAXAMOUNT,            ' +
      '          INVAMOUNT,            ' +
      '          PRINTCOUNT,           ' +
      '          ISPAST,               ' +
      '          ISOBSOLETE,           ' +
      '          HOWTOCREATE,          ' +
      '          MEMO1,                ' +
      '          MEMO2,                ' +
      '          UPTTIME,              ' +
      '          UPTEN,                ' +
      '          PRINTFUN,             ' +
      '          MAININVID,            ' +
      '          MAINSALEAMOUNT,       ' +
      '          MAINTAXAMOUNT,        ' +
      '          MAININVAMOUNT,        ' +
      '          INSTADDR,             ' +
      '          INVUSEID,             ' +
      '          INVUSEDESC   )        ' +
      ' VALUES ( :IDENTIFYID1,         ' +
      '          :IDENTIFYID2,         ' +
      '          :CHECKNO,             ' +
      '          :INVID,               ' +
      '          :CANMODIFY,           ' +
      '          :INVDATE,             ' +
      '          :CHARGEDATE,          ' +
      '          :COMPID,              ' +
      '          :CUSTID,              ' +
      '          :CUSTSNAME,           ' +
      '          :INVTITLE,            ' +
      '          :ZIPCODE,             ' +
      '          :INVADDR,             ' +
      '          :MAILADDR,            ' +
      '          :BUSINESSID,          ' +
      '          :INVFORMAT,           ' +
      '          :TAXTYPE,             ' +
      '          :TAXRATE,             ' +
      '          :SALEAMOUNT,          ' +
      '          :TAXAMOUNT,           ' +
      '          :INVAMOUNT,           ' +
      '          :PRINTCOUNT,          ' +
      '          :ISPAST,              ' +
      '          :ISOBSOLETE,          ' +
      '          :HOWTOCREATE,         ' +
      '          :MEMO1,               ' +
      '          :MEMO2,               ' +
      '          SYSDATE,              ' +
      '          :UPTEN,               ' +
      '          :PRINTFUN,            ' +
      '          :MAININVID,           ' +
      '          :MAINSALEAMOUNT,      ' +
      '          :MAINTAXAMOUNT,       ' +
      '          :MAININVAMOUNT,       ' +
      '          :INSTADDR,            ' +
      '          :INVUSEID,            ' +
      '          :INVUSEDESC   )       '
  else
    Result :=
      ' UPDATE INV007                            ' +
      '    SET INVDATE = :INVDATE,               ' + 
      '        CHARGEDATE = :CHARGEDATE,         ' +
      '        CUSTID = :CUSTID,                 ' +
      '        CUSTSNAME = :CUSTSNAME,           ' +
      '        INVTITLE = :INVTITLE,             ' +
      '        ZIPCODE = :ZIPCODE,               ' +
      '        INVADDR = :INVADDR,               ' +
      '        MAILADDR = :MAILADDR,             ' +
      '        BUSINESSID = :BUSINESSID,         ' +
      '        TAXTYPE = :TAXTYPE,               ' +
      '        TAXRATE = :TAXRATE,               ' +
      '        SALEAMOUNT = :SALEAMOUNT,         ' +
      '        TAXAMOUNT = :TAXAMOUNT,           ' +
      '        INVAMOUNT = :INVAMOUNT,           ' +
      '        MEMO1 = :MEMO1,                   ' +
      '        MEMO2 = :MEMO2,                   ' +
      '        UPTTIME = SYSDATE,                ' +
      '        UPTEN = :UPTEN,                   ' +
      '        MAININVID = :MAININVID,           ' +
      '        MAINSALEAMOUNT = :MAINSALEAMOUNT, ' +
      '        MAINTAXAMOUNT = :MAINTAXAMOUNT,   ' +
      '        MAININVAMOUNT = :MAININVAMOUNT,   ' +
      '        INSTADDR = :INSTADDR,             ' +
      '        INVUSEID = :INVUSEID,             ' +
      '        INVUSEDESC = :INVUSEDESC          ' +
      '  WHERE IDENTIFYID1 = :IDENTIFYID1        ' +
      '    AND IDENTIFYID2 = :IDENTIFYID2        ' +
      '    AND INVID = :INVID                    ';
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA02_4.BuildDetailApplySQL(const aMode: TDMLMode): String;
begin
  if ( aMode in [dmInsert] ) then
    Result :=
      '   INSERT INTO INV008           ' +
      '        (  IDENTIFYID1,         ' +
      '           IDENTIFYID2,         ' +
      '           INVID,               ' +
      '           BILLID,              ' +
      '           BILLIDITEMNO,        ' +
      '           SEQ,                 ' +
      '           STARTDATE,           ' +
      '           ENDDATE,             ' +
      '           ITEMID,              ' +
      '           DESCRIPTION,         ' +
      '           QUANTITY,            ' +
      '           UNITPRICE,           ' +
      '           TAXAMOUNT,           ' +
      '           TOTALAMOUNT,         ' +
      '           CHARGEEN,            ' +
      '           UPTTIME,             ' +
      '           UPTEN,               ' +
      '           SERVICETYPE,         ' +
      '           SALEAMOUNT,          ' +
      '           LINKTOMIS,           ' +
      '           FACISNO,             ' +
      '           ACCOUNTNO  )         ' +
      '  VALUES ( :IDENTIFYID1,        ' +
      '           :IDENTIFYID2,        ' +
      '           :INVID,              ' +
      '           :BILLID,             ' +
      '           :BILLIDITEMNO,       ' +
      '           :SEQ,                ' +
      '           TO_DATE( :STARTDATE, ''YYYY/MM/DD'' ), ' +
      '           TO_DATE( :ENDDATE, ''YYYY/MM/DD'' ), ' +
      '           :ITEMID,             ' +
      '           :DESCRIPTION,        ' +
      '           :QUANTITY,           ' +
      '           :UNITPRICE,          ' +
      '           :TAXAMOUNT,          ' +
      '           :TOTALAMOUNT,        ' +
      '           :CHARGEEN,           ' +
      '           SYSDATE,             ' +
      '           :UPTEN,              ' +
      '           :SERVICETYPE,        ' +
      '           :SALEAMOUNT,         ' +
      '           :LINKTOMIS,          ' +
      '           :FACISNO,            ' +
      '           :ACCOUNTNO   )       ';
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA02_4.DetailSequenceNeeded: Integer;
var
  aBookMark: TBookmark;
  aMaxSequence: Integer;
begin
  Result := 0;
  aMaxSequence := Result;
  aBookMark := cdsDetail.GetBookmark;
  try
    cdsDetail.DisableControls;
    try
      cdsDetail.First;
      while not cdsDetail.Eof do
      begin
        if aMaxSequence < cdsDetail.FieldByName( 'SEQ' ).AsInteger then
          aMaxSequence := cdsDetail.FieldByName( 'SEQ' ).AsInteger;
        cdsDetail.Next;
      end;
      Result := ( aMaxSequence + 1 );
      cdsDetail.GotoBookmark( aBookMark );
    finally
      cdsDetail.EnableControls;
    end;
  finally
    cdsDetail.FreeBookmark( aBookMark );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.ComposeBillingString(var aBillInfo: TBillingInfo);
var
  aBookMark: TBookmark;
begin
  aBillInfo.BillId := '';
  aBillInfo.ItemNo := '';
  aBookMark := cdsDetail.GetBookmark;
  try
    cdsDetail.DisableControls;
    try
      cdsDetail.First;
      while not cdsDetail.Eof do
      begin
        if cdsDetail.FieldByName( 'BILLID' ).AsString <> '' then
          aBillInfo.BillId := aBillInfo.BillId +
            QuotedStr( cdsDetail.FieldByName( 'BILLID' ).AsString ) + ',';
            
        { #5014 �b�g���ɭԶ��K��Item ����޸��@�_���� By Kin 2009/04/03 }
        if cdsDetail.FieldByName( 'BILLID' ).AsString <> '' then
          aBillInfo.ItemNo := aBillInfo.ItemNo +
             cdsDetail.FieldByName( 'BILLIDITEMNO' ).AsString  + ',';
        cdsDetail.Next;
      end;
      if Copy( aBillInfo.BillId, Length( aBillInfo.BillId ), 1 ) = ',' then
        Delete( aBillInfo.BillId, Length( aBillInfo.BillId ), 1 );
      if Copy( aBillInfo.ItemNo, Length( aBillInfo.ItemNo ), 1 ) = ',' then
        Delete( aBillInfo.ItemNo, Length( aBillInfo.ItemNo ), 1 );
      cdsDetail.GotoBookmark( aBookMark );
    finally
      cdsDetail.EnableControls;
    end;
  finally
    cdsDetail.FreeBookmark( aBookMark );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.GetCustomerInfo(const aCustId: String);
var
  aDataSet: TADOQuery;
begin
  // �@��}��, �Ȥ��Ʊq�o���t�Χ�, PS: �ק�ɥi��Ƚs
  // �����ˬd CompId, �u����X�� CustId ���@����Ƥ~�a�i��
  if ( FInvCreateType in [icNormal] ) then
  begin
    aDataSet := dtmMain.adoComm;
    aDataSet.Close;
    aDataSet.SQL.Text :=  Format(
      '  SELECT A.TITLESNAME,                   ' +
      '         A.TITLENAME,                    ' +
      '         A.BUSINESSID,                   ' +
      '         A.INVADDR,                      ' +
      '         A.MZIPCODE,                     ' +
      '         A.MAILADDR                      ' +
      '    FROM INV019 A, INV002 B              ' +
      '   WHERE A.IDENTIFYID1 = B.IDENTIFYID1   ' +
      '     AND A.IDENTIFYID2 = B.IDENTIFYID2   ' +
      '     AND A.COMPID = B.COMPID             ' +
      '     AND A.CUSTID = B.CUSTID             ' +
      '     AND A.CUSTID = ''%s''               ' +
      '     AND A.IDENTIFYID1 = ''%s''          ' +
      '     AND A.IDENTIFYID2 = %s              ',
      [aCustId, IDENTIFYID1, IDENTIFYID2] );
    aDataSet.Open;
    if ( aDataSet.RecordCount = 1 ) then
    begin
      if ( cdsMaster.FieldByName( 'CUSTSNAME' ).AsString = EmptyStr ) then
        cdsMaster.FieldByName( 'CUSTSNAME' ).AsString := aDataSet.FieldByName( 'TITLESNAME' ).AsString;
      if ( cdsMaster.FieldByName( 'INVTITLE' ).AsString = EmptyStr ) then
        cdsMaster.FieldByName( 'INVTITLE' ).AsString := aDataSet.FieldByName( 'TITLENAME' ).AsString;
      if ( cdsMaster.FieldByName( 'BUSINESSID' ).AsString = EmptyStr ) then
        cdsMaster.FieldByName( 'BUSINESSID' ).AsString := aDataSet.FieldByName( 'BUSINESSID' ).AsString;
      if ( cdsMaster.FieldByName( 'INVADDR' ).AsString = EmptyStr ) then
        cdsMaster.FieldByName( 'INVADDR' ).AsString := aDataSet.FieldByName( 'INVADDR' ).AsString;
      if ( cdsMaster.FieldByName( 'ZIPCODE' ).AsString = EmptyStr ) then
        cdsMaster.FieldByName( 'ZIPCODE' ).AsString := aDataSet.FieldByName( 'MZIPCODE' ).AsString;
      if ( cdsMaster.FieldByName( 'MAILADDR' ).AsString = EmptyStr ) then
        cdsMaster.FieldByName( 'MAILADDR' ).AsString := aDataSet.FieldByName( 'MAILADDR' ).AsString;
    end;    
    aDataSet.Close;
  end
  // �{���}��, �w�}, ��}, �Ȥ��Ʊq SMS ��, PS: �ק�ɤ��i���Ƚs
  else if ( FInvCreateType in [icLocale, icPrev, icAfter] ) then
  begin
    if ( FMode in [dmUpdate] ) then Exit;
    {}
    if ( dtmMain.GetLinkToMIS ) and
       ( cdsMaster.FieldByName( 'CUSTSNAME' ).AsString = EmptyStr ) then
    begin
      aDataSet := dtmSO.adoComm;
      aDataSet.Close;
      aDataSet.SQL.Text := Format(
        ' SELECT CUSTNAME FROM %s.SO001 A   ' +
        '  WHERE A.CUSTID = ''%s''          ', [dtmMain.getMisDbOwner, aCustId ] );
      aDataSet.Open;
      cdsMaster.FieldByName( 'CUSTSNAME' ).AsString :=
        aDataSet.FieldByName( 'CUSTNAME' ).AsString;
      aDataSet.Close;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.GetChargeItemInfo(const aItemId: String);
var
  aDataSet: TADOQuery;
begin
  { �����O���� }
  aDataSet := dtmMain.adoComm;
  aDataSet.Close;
  aDataSet.SQL.Text := Format(
    '  SELECT A.DESCRIPTION,              ' +
    '         A.TAXCODE,                  ' +
    '         A.TAXNAME,                  ' +
    '         A.DESCRIPTION,              ' +
    '         A.SIGN                      ' +
    '    FROM INV005 A                    ' +
    '   WHERE  A.IDENTIFYID1 = ''%s''     ' +
    '      AND A.IDENTIFYID2 = %s         ' +
    '      AND A.COMPID = ''%s''          ' +
    '      AND A.ITEMID = ''%s''          ',
    [ IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, aItemId ] );
  aDataSet.Open;
  cdsDetail.FieldByName( 'DESCRIPTION' ).AsString :=
    aDataSet.FieldByName( 'DESCRIPTION' ).AsString;
  cdsDetail.FieldByName( 'TAXCODE' ).AsString :=
    aDataSet.FieldByName( 'TAXCODE' ).AsString;
  cdsDetail.FieldByName( 'TAXNAME' ).AsString :=
    aDataSet.FieldByName( 'TAXNAME' ).AsString;
  cdsDetail.FieldByName( 'SIGN' ).AsString :=
    aDataSet.FieldByName( 'SIGN' ).AsString;
  aDataSet.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.CalculateChargePrice(const aCalculateColum: TChargeColumn);
var
  aQty, aTax, aSale, aTotal: Integer;
  aTaxRate, aPrice: Double;
begin
  if ( aCalculateColum in [chQuantity] ) then
  begin
    { �H�ƶq�p�� }
    aTax := 0;
    aTaxRate := cdsMaster.FieldByName( 'TAXRATE' ).AsFloat;
    aQty := cdsDetail.FieldByName( 'QUANTITY' ).AsInteger;
    aPrice := cdsDetail.FieldByName( 'UNITPRICE' ).AsFloat;  //#4370 �令�B�I�� By Kin 2009/02/19
    aSale := Round( TUother.RoundFloat( aQty * aPrice, 0 ) );
    if aTaxRate > 0 then
      aTax := Round( TUother.RoundFloat( ( aSale * aTaxRate / 100 ), 0 ) );
    aTotal := aSale + aTax;
    cdsDetail.FieldByName( 'TAXAMOUNT' ).AsInteger := aTax;
    cdsDetail.FieldByName( 'SALEAMOUNT' ).AsInteger := aSale;
    cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsInteger := aTotal;
  end else
  if ( aCalculateColum in [chUnitPrice] ) then
  begin
    { �H����p�� }
    if ( cdsDetail.FieldByName( 'UNITPRICE' ).AsFloat = 0 ) then  //#4370 �令�B�I�� By Kin 2009/02/19
    begin
      cdsDetail.FieldByName( 'TAXAMOUNT' ).AsInteger := 0;
      cdsDetail.FieldByName( 'SALEAMOUNT' ).AsInteger := 0;
      cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsInteger := 0;
    end else
    { �Y�ק�����, Tax, Sale, Total ���Ҧ��� }
    if ( cdsDetail.FieldByName( 'TAXAMOUNT' ).AsInteger <> 0 ) and
       ( cdsDetail.FieldByName( 'SALEAMOUNT' ).AsInteger <> 0 ) and
       ( cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsInteger <> 0 ) then
    begin
      aPrice := cdsDetail.FieldByName( 'UNITPRICE' ).AsFloat; //#4370 �令�B�I�� By Kin 2009/02/19
      aQty := cdsDetail.FieldByName( 'QUANTITY' ).AsInteger;
      aTotal := cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsInteger;
//      aSale := ( aQty * aPrice );
      aSale := Round( TUother.RoundFloat( aQty * aPrice, 0 ) ); //4370 �]��Price�w�令�B�I�� By Kin 2009/02/19
      aTax := (  aTotal - aSale );
      cdsDetail.FieldByName( 'SALEAMOUNT' ).AsInteger := aSale;
      cdsDetail.FieldByName( 'TAXAMOUNT' ).AsInteger := aTax;
    end else
    begin
      aTax := 0;
      aTaxRate := cdsMaster.FieldByName( 'TAXRATE' ).AsFloat;
      aQty := cdsDetail.FieldByName( 'QUANTITY' ).AsInteger;
      aPrice := cdsDetail.FieldByName( 'UNITPRICE' ).AsFloat; //#4370 �令�B�I�� By Kin 2009/02/19
      aSale := Round( TUother.RoundFloat( aQty * aPrice, 0 ) );
      if aTaxRate > 0 then
        aTax := Round( TUother.RoundFloat( ( aSale * aTaxRate / 100 ), 0 ) );
      aTotal := aSale + aTax;
      cdsDetail.FieldByName( 'TAXAMOUNT' ).AsInteger := aTax;
      cdsDetail.FieldByName( 'SALEAMOUNT' ).AsInteger := aSale;
      cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsInteger := aTotal;
    end;
  end else
  if ( aCalculateColum in [chTotalAmount] ) then
  begin
    { �H�`���B�p�� }
    if ( cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsInteger = 0 ) then
    begin
      cdsDetail.FieldByName( 'SALEAMOUNT' ).AsInteger := 0;
      cdsDetail.FieldByName( 'TAXAMOUNT' ).AsInteger := 0;
      cdsDetail.FieldByName( 'UNITPRICE' ).AsInteger := 0;
    end else
    begin
      aQty := cdsDetail.FieldByName( 'QUANTITY' ).AsInteger;
      aPrice := cdsDetail.FieldByName( 'UNITPRICE' ).AsFloat; //#4370 �令�B�I�� By Kin 2009/02/19
      aTaxRate := cdsMaster.FieldByName( 'TAXRATE' ).AsFloat;
      aTotal := cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsInteger;
      aSale := aTotal;
      if aTaxRate > 0 then
        aSale := Round( TUother.RoundFloat(
          aTotal / ( 1 + ( aTaxRate / 100 ) ), 0 ) );
      aTax := ( aTotal - aSale );
      cdsDetail.FieldByName( 'SALEAMOUNT' ).AsInteger := aSale;
      cdsDetail.FieldByName( 'TAXAMOUNT' ).AsInteger := aTax;
      aPrice := Round( TUother.RoundFloat( aSale / aQty, 2 ) );
      cdsDetail.FieldByName( 'UNITPRICE' ).AsFloat := aPrice;
    end;
  end else
  if ( aCalculateColum in [chSaleAmount] ) then
  begin
    { �H�P���B�p�� }
    if ( cdsDetail.FieldByName( 'SALEAMOUNT' ).AsInteger = 0 ) then
    begin
      cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsFloat := 0;
      cdsDetail.FieldByName( 'TAXAMOUNT' ).AsFloat := 0;
      cdsDetail.FieldByName( 'UNITPRICE' ).AsFloat := 0;
    end else
    { �Y�ק�P���B��, Tax, Price, Total ���Ҧ��� }
    if ( cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsInteger <> 0 ) and
       ( cdsDetail.FieldByName( 'TAXAMOUNT' ).AsInteger <> 0 ) and
       ( cdsDetail.FieldByName( 'UNITPRICE' ).AsFloat <> 0 ) then
    begin
      aPrice := cdsDetail.FieldByName( 'UNITPRICE' ).AsFloat; //#4370 �令�B�I�� By Kin 2009/02/19
      aTotal := cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsInteger;
      aSale := cdsDetail.FieldByName( 'SALEAMOUNT' ).AsInteger;
      aTax := ( aTotal - aSale );
      cdsDetail.FieldByName( 'TAXAMOUNT' ).AsInteger := aTax;
    end else
    begin
      aQty := cdsDetail.FieldByName( 'QUANTITY' ).AsInteger;
      aPrice := cdsDetail.FieldByName( 'UNITPRICE' ).AsFloat; //#4370 �令�B�I�� By Kin 2009/02/19
      aTax := 0;
      aTaxRate := cdsMaster.FieldByName( 'TAXRATE' ).AsFloat;
      aSale := cdsDetail.FieldByName( 'SALEAMOUNT' ).AsInteger;
      if aTaxRate > 0 then
        aTax := Round( TUother.RoundFloat( ( aSale * aTaxRate / 100 ), 0 ) );
      aTotal := aSale + aTax;
      cdsDetail.FieldByName( 'TAXAMOUNT' ).AsInteger := aTax;
      cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsInteger := aTotal;
      aPrice := Round( TUother.RoundFloat( aSale / aQty, 2 ) );
      cdsDetail.FieldByName( 'UNITPRICE' ).AsFloat := aPrice; //#4370 �令�B�I�� By Kin 2009/02/19
      UpdateDetailPriceByChargeSign;
    end;
  end else
  if ( aCalculateColum in [chTaxAmount] ) then
  begin
    { �Y�ק�|�B��, Sale, Price, Total ���Ҧ��� }
    if ( cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsInteger <> 0 ) and
       ( cdsDetail.FieldByName( 'SALEAMOUNT' ).AsInteger <> 0 ) and
       ( cdsDetail.FieldByName( 'UNITPRICE' ).AsInteger <> 0 ) then
    begin
      aTotal := cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsInteger;
      aTax := cdsDetail.FieldByName( 'TAXAMOUNT' ).AsInteger;
      aSale := ( aTotal - aTax );
      cdsDetail.FieldByName( 'SALEAMOUNT' ).AsInteger := aSale;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA02_4.PopChargeWindow(const aCustId: String): TModalResult;
begin
  if Assigned( FChargeWindow ) then
  begin
    FChargeWindow.Free;
    FChargeWindow := nil;
  end;
  FTaxInfo.TaxCode := EmptyStr;
  FTaxInfo.TaxRate := EmptyStr;
  if ( cdsDetail.RecordCount > 0 ) then
  begin
    FTaxInfo.TaxCode := cdsMaster.FieldByName( 'TAXTYPE' ).AsString;
    FTaxInfo.TaxRate := cdsMaster.FieldByName( 'TAXRATE' ).AsString;
  end;
  FChargeWindow := TfrmInvA02_1.Create( aCustId, FBillingInfo.BillId,
    FBillingInfo.ItemNo, FTaxInfo.TaxCode, FTaxInfo.TaxRate );
  Result := FChargeWindow.ShowModal;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA02_4.PopTitleWindow(const aCustId: String): TModalResult;
begin
  Result := mrNone;
  if Assigned( FTitleWindow ) then
  begin
    FTitleWindow.Free;
    FTitleWindow := nil;
  end;
  FTitleWindow := TfrmInvA02_5.Create( aCustId );
  Result := FTitleWindow.ShowModal;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA02_4.PopTitleWindow2(const aCustId: String): TModalResult;
begin
  Result := mrNone;
  if Assigned( FSearchCustWindow ) then
  begin
    FSearchCustWindow.Free;
    FSearchCustWindow := nil;
  end;
  FSearchCustWindow := TfrmSearchCust.Create( aCustId );
  Result := FSearchCustWindow.ShowModal;
end;

{ ---------------------------------------------------------------------------- }


function TfrmInvA02_4.PopRelationInvId(const aMainInvId: String): TModalResult;
begin
  Result := mrNone;
  if Assigned( FRelationInvWindow ) then
  begin
    FRelationInvWindow.Free;
    FRelationInvWindow := nil;
  end;
  FRelationInvWindow := TfrmInvA02_6.Create( nil );
  TfrmInvA02_6( FRelationInvWindow ).MainInvId := aMainInvId;
  TfrmInvA02_6( FRelationInvWindow ).InvId := FInvoiceId;
  Result := FRelationInvWindow.ShowModal;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA02_4.GetItemChargeSign(const aItemId: String): String;
var
  aDataSet: TADOQuery;
begin
  { �����O���ت����t�� }
  Result := '+';
  aDataSet := dtmMain.adoComm;
  aDataSet.Close;
  aDataSet.SQL.Text := Format(
     '  SELECT SIGN     ' +
     '    FROM INV005   ' +
     '   WHERE IDENTIFYID1 = ''%s''     ' +
     '     AND IDENTIFYID2 = %s         ' +
     '     AND COMPID = ''%s''          ' +    
     '     AND ITEMID = ''%s''          ',
    [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, aItemId] );
  aDataSet.Open;
  if not aDataSet.IsEmpty then
    Result := Nvl( aDataSet.FieldByName( 'SIGN' ).AsString, '+' );
  aDataSet.Close;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA02_4.GetMasterTaxCode: String;
begin
  Result := '1';
  if rdoTAXRATE2.Checked then
    Result := '2'
  else if rdoTAXRATE3.Checked then
    Result := '3';
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA02_4.GetInvoiceCheckNumber: String;
var
  aSystemId: String;
begin
  Result := EmptyStr;
  aSystemId := dtmMainJ.GetXMLAttribute( Format(
    'SELECT INVPARAM FROM INV003 WHERE COMPID = ''%s'' ', [dtmMain.getCompID] ), 'SysID' );
  Result := dtmMainH.getCheckNo( cdsMaster.FieldByName( 'INVID' ).AsString,
    FormatDateTime( 'yyyy/mm/dd', cdsMaster.FieldByName( 'INVDATE' ).AsDateTime ),
    aSystemId );
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA02_4.GetCurrentAvailableInvCount: Integer;
var
  aYearMonth, aPrefix, aNum: String;
begin
  Result := 0;
  dtmMain.adoComm2.Close;
  if ( FMode in [dmInsert] ) then
  begin
    dtmMain.adoComm2.SQL.Text := Format(
        ' SELECT ( A.ENDNUM - A.CURNUM + 1 ) AS COUNTS  ' +
        '   FROM INV099 A                               ' +
        '  WHERE A.IDENTIFYID1 = ''%s''                 ' +
        '    AND A.IDENTIFYID2 = %s                     ' +
        '    AND A.COMPID = ''%s''                      ' +
        '    AND A.YEARMONTH = ''%s''                   ' +
        '    AND A.PREFIX = ''%s''                      ' +
        '    AND A.STARTNUM = ''%s''                    ' +
        '    AND A.USEFUL = ''Y''                       ',
        [ IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, FChooseInv.YearMonth,
          FChooseInv.Prefix, FChooseInv.StartNum] );
  end else
  begin
    aYearMonth := MapingInvoiceYearMonth(
      cdsMaster.FieldByName( 'INVDATE' ).AsDateTime );
    aPrefix := Copy( FInvoiceId, 1, 2 );
    aNum := Copy( FInvoiceId, 3, Length( FInvoiceId ) - 2 );
    dtmMain.adoComm2.SQL.Text := Format(
        ' SELECT ( A.ENDNUM - A.CURNUM + 1 ) AS COUNTS  ' +
        '   FROM INV099 A                               ' +
        '  WHERE A.IDENTIFYID1 = ''%s''                 ' +
        '    AND A.IDENTIFYID2 = %s                     ' +
        '    AND A.COMPID = ''%s''                      ' +
        '    AND A.YEARMONTH = ''%s''                   ' +
        '    AND A.PREFIX = ''%s''                      ' +
        '    AND ''%s'' BETWEEN A.STARTNUM AND A.ENDNUM ' +
        '    AND A.USEFUL = ''Y''                       ',
        [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, aYearMonth,
         aPrefix, aNum] );
  end;        
  dtmMain.adoComm2.Open;
  Result := dtmMain.adoComm2.FieldByName( 'COUNTS' ).AsInteger;
  dtmMain.adoComm2.Close;
end;

{ ---------------------------------------------------------------------------- }


function TfrmInvA02_4.GetNextInvNumber(const aInvId: String): String;
var
  aNum: Integer;
begin
  aNum := StrToInt( Copy( aInvId, 3, Length( aInvId ) - 2 ) );
  Inc( aNum );
  Result := Copy( aInvId, 1, 2 ) + Lpad( IntToStr( aNum ), 8, '0' );
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA02_4.ReverseInvStartNum(const aChoose: TChooseInvInfo): String;
var
  aNum: String;
begin
  Result := EmptyStr;
  aNum := Copy( FInvoiceId, 3, Length( FInvoiceId ) - 2 );
  dtmMain.adoComm2.Close;
  dtmMain.adoComm2.SQL.Text := Format(
    ' SELECT A.STARTNUM FROM INV099 A                ' +
    '  WHERE A.IDENTIFYID1 = ''%s''                  ' +
    '    AND A.IDENTIFYID2 = ''%s''                  ' +
    '    AND A.COMPID = ''%s''                       ' +
    '    AND A.YEARMONTH = ''%s''                    ' +
    '    AND A.PREFIX = ''%s''                       ' +
    '    AND ''%s'' BETWEEN A.STARTNUM AND A.ENDNUM  ',
    [IDENTIFYID1, IDENTIFYID2,  dtmMain.getCompID, aChoose.YearMonth,
     aChoose.Prefix, aNum] );
  dtmMain.adoComm2.Open;
  Result := dtmMain.adoComm2.FieldByName( 'STARTNUM' ).AsString;
  dtmMain.adoComm2.Close;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA02_4.ValidateCanSetInvUse(const aInvUseId: String): Boolean;
begin
  dtmMain.adoComm2.Close;
  dtmMain.adoComm2.SQL.Text := Format(
    ' SELECT A.REFNO FROM INV028 A                   ' +
    '  WHERE A.IDENTIFYID1 = ''%s''                  ' +
    '    AND A.IDENTIFYID2 = ''%s''                  ' +
    '    AND A.COMPID = ''%s''                       ' +
    '    AND A.ITEMID = ''%s''                       ',
    [IDENTIFYID1, IDENTIFYID2,  dtmMain.getCompID, aInvUseId] );
  dtmMain.adoComm2.Open;
  Result := ( dtmMain.adoComm2.FieldByName( 'REFNO' ).AsString <> '1' );
  dtmMain.adoComm2.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.DisableDetailDataSetEvent;
begin
  cdsDetail.AfterDelete := nil;
  cdsDetail.AfterInsert := nil;
  cdsDetail.BeforePost := nil;
  cdsDetail.AfterPost := nil;
  cdsDetail.BeforeInsert := nil;
  cdsDetail.OnNewRecord := nil;
  cdsDetail.BeforeDelete := nil;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.EnableDetailDataSetEvent;
begin
  cdsDetail.AfterDelete := cdsDetailAfterDelete;
  cdsDetail.BeforePost := nil; //cdsDetailBeforePost; �w�����n
  cdsDetail.AfterPost := cdsDetailAfterPost;
  cdsDetail.BeforeInsert := cdsDetailBeforeInsert;
  cdsDetail.OnNewRecord := cdsDetailNewRecord;
  cdsDetail.BeforeDelete := cdsDetailBeforeDelete;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.PrepareDetailDataSet;
begin
  if ( cdsDetail.FieldDefs.Count <= 0 ) then
  begin
    cdsDetail.FieldDefs.Add( 'IDENTIFYID1', ftString, 2 );
    cdsDetail.FieldDefs.Add( 'IDENTIFYID2', ftString, 2 );
    cdsDetail.FieldDefs.Add( 'INVID', ftString, 10 );
    cdsDetail.FieldDefs.Add( 'BILLID', ftString, 15 );
    cdsDetail.FieldDefs.Add( 'BILLIDITEMNO',ftInteger );
    cdsDetail.FieldDefs.Add( 'SEQ', ftInteger );
    cdsDetail.FieldDefs.Add( 'STARTDATE', ftDateTime );
    cdsDetail.FieldDefs.Add( 'ENDDATE', ftDateTime );
    cdsDetail.FieldDefs.Add( 'ITEMID', ftString, 7 );
    cdsDetail.FieldDefs.Add( 'DESCRIPTION', ftString, 40 );
    cdsDetail.FieldDefs.Add( 'QUANTITY', ftInteger);
    //#4370 ����n�令�B�I�� By Kin 2009/02/19
    cdsDetail.FieldDefs.Add( 'UNITPRICE', ftFloat );
    cdsDetail.FieldDefs.Add( 'SALEAMOUNT', ftBCD );
    cdsDetail.FieldDefs.Add( 'TAXAMOUNT', ftBCD );
    cdsDetail.FieldDefs.Add( 'TOTALAMOUNT', ftBCD );
    cdsDetail.FieldDefs.Add( 'CHARGEEN', ftString, 20 );
    cdsDetail.FieldDefs.Add( 'UPTTIME', ftDateTime );
    cdsDetail.FieldDefs.Add( 'UPTEN', ftString, 20 );
    cdsDetail.FieldDefs.Add( 'SERVICETYPE', ftString, 1 );
    cdsDetail.FieldDefs.Add( 'LINKTOMIS', ftString, 1 );
    cdsDetail.FieldDefs.Add( 'LINKTOMISDESC', ftString, 2 );
    cdsDetail.FieldDefs.Add( 'TAXCODE', ftInteger );
    cdsDetail.FieldDefs.Add( 'TAXNAME', ftString, 12 );
    cdsDetail.FieldDefs.Add( 'SIGN', ftString, 1 );
    cdsDetail.FieldDefs.Add( 'SOURCE', ftString, 10 );
    cdsDetail.FieldDefs.Add( 'EMPTY', ftString, 10 );
    cdsDetail.FieldDefs.Add( 'FACISNO', ftString, 20 );
    cdsDetail.FieldDefs.Add( 'ACCOUNTNO', ftString, 16 );
    cdsDetail.CreateDataSet;
  end;
  cdsDetail.EmptyDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.PrepareInv008ADataSet;
begin
  if ( cdsInv008A.FieldDefs.Count <= 0 ) then
  begin
    cdsInv008A.FieldDefs.Add( 'ITEMIDREF', ftInteger );
    cdsInv008A.FieldDefs.Add( 'INVID', ftString, 10 );
    cdsInv008A.FieldDefs.Add( 'BILLID', ftString, 15 );
    cdsInv008A.FieldDefs.Add( 'BILLIDITEMNO', ftString, 2 );
    cdsInv008A.FieldDefs.Add( 'SEQ', ftInteger );
    cdsInv008A.FieldDefs.Add( 'STARTDATE', ftString, 10 );
    cdsInv008A.FieldDefs.Add( 'ENDDATE', ftString, 10 );
    cdsInv008A.FieldDefs.Add( 'ITEMID', ftString, 7 );
    cdsInv008A.FieldDefs.Add( 'DESCRIPTION', ftString, 40 );
    cdsInv008A.FieldDefs.Add( 'QUANTITY', ftInteger );
    cdsInv008A.FieldDefs.Add( 'UNITPRICE', ftFloat );
    cdsInv008A.FieldDefs.Add( 'TAXAMOUNT', ftInteger );
    cdsInv008A.FieldDefs.Add( 'SALEAMOUNT', ftInteger );
    cdsInv008A.FieldDefs.Add( 'TOTALAMOUNT', ftInteger );
    cdsInv008A.FieldDefs.Add( 'SERVICETYPE', ftString, 1 );
    cdsInv008A.FieldDefs.Add( 'FACISNO', ftString, 20 );
    cdsInv008A.FieldDefs.Add( 'ACCOUNTNO', ftString, 16 );
    cdsInv008A.CreateDataSet;
  end;
  cdsInv008A.EmptyDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.FillDetailDataSet;
var
  aOldReadOnly: Boolean;
  aStartDate, aEndDate: String;
begin
  aOldReadOnly := cdsDetail.ReadOnly;
  try
    cdsDetail.ReadOnly := False;
    dtmMainH.adoA02_Detail.Close;
    dtmMainH.adoA02_Detail.SQL.Text := BuildDetailSQL( FInvoiceId );
    dtmMainH.adoA02_Detail.Open;
    dtmMainH.adoA02_Detail.First;
    while not dtmMainH.adoA02_Detail.Eof do
    begin
      { ���ɮ� ����榡�i��O����, �|�ɦܿ��~ }
      try
        aStartDate := dtmMainH.adoA02_Detail.FieldByName( 'STARTDATE' ).AsString;
      except
        aStartDate := EmptyStr;
      end;
      try
        aEndDate := dtmMainH.adoA02_Detail.FieldByName( 'ENDDATE' ).AsString;
      except
        aEndDate := EmptyStr;
      end;
      cdsDetail.AppendRecord( [
        dtmMainH.adoA02_Detail.FieldByName( 'IDENTIFYID1' ).AsString,
        dtmMainH.adoA02_Detail.FieldByName( 'IDENTIFYID2' ).AsString,
        dtmMainH.adoA02_Detail.FieldByName( 'INVID' ).AsString,
        dtmMainH.adoA02_Detail.FieldByName( 'BILLID' ).AsString,
        dtmMainH.adoA02_Detail.FieldByName( 'BILLIDITEMNO' ).AsString,
        dtmMainH.adoA02_Detail.FieldByName( 'SEQ' ).AsString,
        aStartDate,
        aEndDate,
        dtmMainH.adoA02_Detail.FieldByName( 'ITEMID' ).AsString,
        dtmMainH.adoA02_Detail.FieldByName( 'DESCRIPTION' ).AsString,
        dtmMainH.adoA02_Detail.FieldByName( 'QUANTITY' ).AsString,
        dtmMainH.adoA02_Detail.FieldByName( 'UNITPRICE' ).AsString,
        dtmMainH.adoA02_Detail.FieldByName( 'SALEAMOUNT' ).AsString,
        dtmMainH.adoA02_Detail.FieldByName( 'TAXAMOUNT' ).AsString,
        dtmMainH.adoA02_Detail.FieldByName( 'TOTALAMOUNT' ).AsString,
        dtmMainH.adoA02_Detail.FieldByName( 'CHARGEEN' ).AsString,
        dtmMainH.adoA02_Detail.FieldByName( 'UPTTIME' ).AsString,
        dtmMainH.adoA02_Detail.FieldByName( 'UPTEN' ).AsString,
        dtmMainH.adoA02_Detail.FieldByName( 'SERVICETYPE' ).AsString,
        dtmMainH.adoA02_Detail.FieldByName( 'LINKTOMIS' ).AsString,
        dtmMainH.adoA02_Detail.FieldByName( 'LINKTOMISDESC' ).AsString,
        dtmMainH.adoA02_Detail.FieldByName( 'TAXCODE' ).AsString,
        dtmMainH.adoA02_Detail.FieldByName( 'TAXNAME' ).AsString,
        dtmMainH.adoA02_Detail.FieldByName( 'SIGN' ).AsString,
        dtmMainH.adoA02_Detail.FieldByName( 'SOURCE' ).AsString,
        dtmMainH.adoA02_Detail.FieldByName( 'EMPTY' ).AsString,
        dtmMainH.adoA02_Detail.FieldByName( 'FACISNO' ).AsString,
        dtmMainH.adoA02_Detail.FieldByName( 'ACCOUNTNO' ).AsString] );
      dtmMainH.adoA02_Detail.Next;
    end;
    dtmMainH.adoA02_Detail.Close;
  finally
    cdsDetail.ReadOnly := aOldReadOnly;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.FillInv008ADataSet;
var
  aReader: TADOQuery;
  aStartDate, aEndDate: String;
begin
  aReader := dtmMain.adoComm3;
  aReader.Close;
  aReader.SQL.Text := BuildINV008ASQL( FInvoiceId );
  aReader.Open;
  aReader.First;
  while not aReader.Eof do
  begin
    { ���ɮ� ����榡�i��O����, �|�ɦܿ��~ }
    try
      aStartDate := aReader.FieldByName( 'STARTDATE' ).AsString;
    except
      aStartDate := EmptyStr;
    end;
    try
      aEndDate := aReader.FieldByName( 'ENDDATE' ).AsString;
    except
      aEndDate := EmptyStr;
    end;
    cdsInv008A.AppendRecord( [
      aReader.FieldByName( 'ITEMIDREF' ).AsString,
      aReader.FieldByName( 'INVID' ).AsString,
      aReader.FieldByName( 'BILLID' ).AsString,
      aReader.FieldByName( 'BILLIDITEMNO' ).AsString,
      aReader.FieldByName( 'SEQ' ).AsString,
      aStartDate,
      aEndDate,
      aReader.FieldByName( 'ITEMID' ).AsString,
      aReader.FieldByName( 'DESCRIPTION' ).AsString,
      aReader.FieldByName( 'QUANTITY' ).AsInteger,
      aReader.FieldByName( 'UNITPRICE' ).AsFloat,
      aReader.FieldByName( 'TAXAMOUNT' ).AsInteger,
      aReader.FieldByName( 'SALEAMOUNT' ).AsInteger,
      aReader.FieldByName( 'TOTALAMOUNT' ).AsInteger,
      aReader.FieldByName( 'SERVICETYPE' ).AsString,
      aReader.FieldByName( 'FACISNO' ).AsString,
      aReader.FieldByName( 'ACCOUNTNO' ).AsString] );
    aReader.Next;
  end;
  aReader.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.FillChooseInv;
begin
  ZeroMemory( @FChargeWindow, SizeOf( FChooseInv ) );
  if cdsMaster.FieldByName( 'INVID' ).AsString = EmptyStr then Exit;
  if cdsMaster.FieldByName( 'INVDATE' ).AsString = EmptyStr then Exit;
  FChooseInv.YearMonth := MapingInvoiceYearMonth(
    cdsMaster.FieldByName( 'INVDATE' ).AsDateTime );
  FChooseInv.InvFormat := cdsMaster.FieldByName( 'INVFORMAT' ).AsString;
  FChooseInv.Prefix := Copy( cdsMaster.FieldByName( 'INVID' ).AsString, 1, 2 );
  FChooseInv.OldInvDate := FormatDateTime( 'yyyymmdd',
    cdsMaster.FieldByName( 'INVDATE' ).AsDateTime );
  FChooseInv.StartNum := ReverseInvStartNum( FChooseInv );  
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.CopyChargeDataToDetail;
var
  aDataSet: TClientDataSet;
  aChargeSign, aAccountNo, aFacisNo: String;
  aRay: Double;
begin
  aDataSet := TfrmInvA02_1( FChargeWindow ).ResultDataSet;
  if aDataSet.FieldDefs.Count = 0 then Exit;
  Screen.Cursor := crSQLWait;
  try
    cdsDetail.DisableControls;
    try
      cdsDetail.AfterPost := nil;
      try
        if cdsDetail.State in [dsEdit, dsInsert] then
          cdsDetail.Cancel;
        aDataSet.First;
        while not aDataSet.Eof do
        begin

          cdsDetail.Append;
          
          cdsDetail.FieldByName( 'BILLID' ).AsString :=
            aDataSet.FieldByName( 'BILLNO' ).AsString;

          cdsDetail.FieldByName( 'BILLIDITEMNO' ).AsString :=
            aDataSet.FieldByName( 'ITEM' ).AsString;

          cdsDetail.FieldByName( 'STARTDATE' ).Value := Null;
          if aDataSet.FieldByName( 'REALSTARTDATE' ).AsDateTime > 0 then
            cdsDetail.FieldByName( 'STARTDATE' ).AsDateTime :=
              aDataSet.FieldByName( 'REALSTARTDATE' ).AsDateTime;

          cdsDetail.FieldByName( 'ENDDATE' ).Value := Null;
          if aDataSet.FieldByName( 'REALSTOPDATE' ).AsDateTime > 0 then
            cdsDetail.FieldByName('ENDDATE').AsDateTime :=
              aDataSet.fieldByName( 'REALSTOPDATE' ).AsDateTime;

          cdsDetail.FieldByName( 'ITEMID' ).AsString :=
            aDataSet.FieldByName( 'CITEMCODE' ).AsString;

          cdsDetail.FieldByName( 'DESCRIPTION' ).AsString :=
            aDataSet.FieldByName( 'CITEMNAME' ).AsString;

          cdsDetail.FieldByName( 'QUANTITY' ).AsInteger := 1;
          {}
          { #5057 �p�G�D�諸�OSO033,���P�_�O�_���ꦬ���,�p�G���h��REALAMT By Kin 2009/05/05 }
          {#5285 �P�_�ꦬ����令�P�_������] By Kin 2009/09/09 }
          if aDataSet.FieldByName( 'SOURCE' ).AsString = '33' then
          begin
            if aDataSet.FieldByName( 'UCCode' ).AsString <> ''  then
            begin
              cdsDetail.FieldByName( 'TOTALAMOUNT' ).Value :=
                aDataSet.FieldByName( 'SHOULDAMT' ).Value;
            end else
            begin
              cdsDetail.FieldByName( 'TOTALAMOUNT' ).Value :=
                aDataSet.FieldByName( 'REALAMT' ).Value;
            end;
          end else
          begin
            cdsDetail.FieldByName( 'TOTALAMOUNT' ).Value :=
              aDataSet.FieldByName( 'REALAMT' ).Value;
          end;
          cdsDetail.FieldByName( 'TAXCODE' ).AsString :=
            aDataSet.FieldByName( 'TAXCODE' ).AsString;

          cdsDetail.FieldByName( 'TAXNAME' ).AsString :=
            aDataSet.FieldByName( 'DESCRIPTION' ).AsString;

          { ���| }
          if ( cdsDetail.FieldByName( 'TAXCODE' ).AsString = '1' ) then
          begin
            aRay := ( aDataSet.FieldByName( 'RATE1' ).AsInteger / 100 );
            cdsDetail.FieldByName( 'UNITPRICE' ).AsInteger := Trunc( RoundTo(
              cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsInteger / ( 1 + aRay ), 0 ) );
           cdsDetail.FieldByName( 'SALEAMOUNT' ).AsInteger :=
             cdsDetail.FieldByName( 'UNITPRICE' ).AsInteger;
           cdsDetail.FieldByName( 'TAXAMOUNT' ).AsInteger :=
             ( cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsInteger -
               cdsDetail.FieldByName( 'SALEAMOUNT' ).AsInteger );
          end
          else begin
            cdsDetail.FieldByName( 'UNITPRICE' ).AsInteger :=
              cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsInteger;
            cdsDetail.FieldByName( 'SALEAMOUNT' ).AsInteger :=
              cdsDetail.FieldByName( 'UNITPRICE' ).AsInteger;
           cdsDetail.FieldByName( 'TAXAMOUNT' ).AsInteger := 0;
          end;     

          cdsDetail.FieldByName( 'SERVICETYPE' ).AsString :=
            aDataSet.FieldByName('SERVICETYPE').AsString;

          cdsDetail.FieldByName( 'CHARGEEN' ).AsString :=
            aDataSet.FieldByName( 'CLCTNAME' ).AsString;

          cdsDetail.FieldByName( 'UPTTIME' ).AsDateTime := Now;

          cdsDetail.FieldByName( 'UPTEN' ).AsString := dtmMain.getLoginUser;

          cdsDetail.FieldByName( 'SIGN' ).AsString := GetItemChargeSign(
            cdsDetail.FieldByName( 'ITEMID' ).AsString );

          cdsDetail.FieldByName( 'SOURCE' ).AsString :=
            aDataSet.fieldByName( 'SOURCE' ).AsString;

          { �t�����O }
          aChargeSign := GetItemChargeSign( cdsDetail.FieldByName( 'ITEMID' ).AsString );
          if aChargeSign = '-' then
          begin
            if ( cdsDetail.FieldByName( 'UNITPRICE' ).AsFloat > 0 ) then
              cdsDetail.FieldByName( 'UNITPRICE' ).AsFloat := ( 0 -
              cdsDetail.FieldByName( 'UNITPRICE' ).AsFloat );
            {}
            if ( cdsDetail.FieldByName( 'SALEAMOUNT' ).AsFloat > 0 ) then
              cdsDetail.FieldByName( 'SALEAMOUNT' ).AsFloat := ( 0 -
              cdsDetail.FieldByName( 'SALEAMOUNT' ).AsFloat );
            {}
            if ( cdsDetail.FieldByName( 'TAXAMOUNT' ).AsFloat > 0 ) then
              cdsDetail.FieldByName( 'TAXAMOUNT' ).AsFloat := ( 0 -
              cdsDetail.FieldByName( 'TAXAMOUNT' ).AsFloat );
            {}
            if ( cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsFloat > 0 ) then
              cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsFloat := ( 0 -
              cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsFloat );
          end;

          { �D�X�ֶi�Ӫ�����, �������ΫH�Υd�� }
          if not IsMergeRecord( FInvoiceId,
            cdsDetail.FieldByName( 'SEQ' ).AsString ) then
          begin
            dtmSO.GetMisAccountAndFacisNo(
              cdsDetail.FieldByName( 'BILLID' ).AsString,
              cdsDetail.FieldByName( 'BILLIDITEMNO' ).AsString,
              txtCUSTID.Text, aAccountNo,aFacisNo );
            if ( aFacisNo <> EmptyStr ) then
              cdsDetail.FieldByName( 'FACISNO' ).AsString := aFacisNo;
            if ( aAccountNo <> EmptyStr ) then
              cdsDetail.FieldByName( 'ACCOUNTNO' ).AsString := aAccountNo;
          end;

          cdsDetail.Post;
          aDataSet.Next;
        end;
        cdsDetailAfterPost( cdsDetail );
      finally
        cdsDetail.AfterPost := cdsDetailAfterPost;
      end;
    finally
      cdsDetail.EnableControls;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.CopyChargeTitleToMaster(const aInvCreateType: TInvCreateType);
var
  aDataSet: TClientDataSet;
begin
  if ( aInvCreateType in [icPrev, icAfter, icLocale] ) then
  begin
    aDataSet := TfrmInvA02_1( FChargeWindow ).TitleDataSet;
    if cdsMaster.FieldByName( 'INVTITLE' ).AsString = EmptyStr then
    begin
      cdsMaster.FieldByName( 'INVTITLE' ).AsString :=
        aDataSet.FieldByName( 'INVTITLE' ).AsString;
    end;
    if cdsMaster.FieldByName( 'BUSINESSID' ).AsString = EmptyStr then
    begin
      cdsMaster.FieldByName( 'BUSINESSID' ).AsString :=
        aDataSet.FieldByName( 'BUSINESSID' ).AsString;
    end;
    if cdsMaster.FieldByName( 'INVADDR' ).AsString = EmptyStr then
    begin
      cdsMaster.FieldByName( 'INVADDR' ).AsString :=
        aDataSet.FieldByName( 'INVADDRESS' ).AsString;
    end;
    if cdsMaster.FieldByName( 'ZIPCODE' ).AsString = EmptyStr then
    begin
      cdsMaster.FieldByName( 'ZIPCODE' ).AsString :=
        aDataSet.FieldByName( 'ZIPCODE' ).AsString;
    end;
    if cdsMaster.FieldByName( 'MAILADDR' ).AsString = EmptyStr then
    begin
      cdsMaster.FieldByName( 'MAILADDR' ).AsString :=
        aDataSet.FieldByName( 'MAILADDRESS' ).AsString;
    end;
    if cdsMaster.FieldByName( 'INSTADDR' ).AsString = EmptyStr then
    begin
      cdsMaster.FieldByName( 'INSTADDR' ).AsString :=
        aDataSet.FieldByName( 'INSTADDRESS' ).AsString;
    end;
    if cdsMaster.FieldByName( 'CUSTSNAME' ).AsString = EmptyStr then
    begin
      cdsMaster.FieldByName( 'CUSTSNAME' ).AsString :=
        aDataSet.FieldByName( 'CHARGETITLE' ).AsString;
    end;
    if ( cdsMaster.FieldByName( 'INVUSEID' ).AsString = EmptyStr ) then
    begin
      cdsMaster.FieldByName( 'INVUSEID' ).AsString :=
        aDataSet.FieldByName( 'INVPURPOSECODE' ).AsString;
      cdsMaster.FieldByName( 'INVUSEDESC' ).AsString :=
        aDataSet.FieldByName( 'INVPURPOSENAME' ).AsString;
    end;
  end
  else if ( aInvCreateType in [icNormal] ) then
  begin
    aDataSet := TfrmSearchCust( FSearchCustWindow ).TitleDataSet;
    cdsMaster.FieldByName( 'CUSTID' ).AsString := aDataSet.FieldByName( 'CUSTID' ).AsString;
    cdsMaster.FieldByName( 'CUSTSNAME' ).AsString := aDataSet.FieldByName( 'TITLESNAME' ).AsString;
    cdsMaster.FieldByName( 'INVTITLE' ).AsString := aDataSet.FieldByName( 'TITLENAME' ).AsString;
    cdsMaster.FieldByName( 'BUSINESSID' ).AsString := aDataSet.FieldByName( 'BUSINESSID' ).AsString;
    cdsMaster.FieldByName( 'INVADDR' ).AsString := aDataSet.FieldByName( 'INVADDR' ).AsString;
    cdsMaster.FieldByName( 'ZIPCODE' ).AsString := aDataSet.FieldByName( 'MZIPCODE' ).AsString;
    cdsMaster.FieldByName( 'MAILADDR' ).AsString := aDataSet.FieldByName( 'MAILADDR' ).AsString;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.GridStateChange(const aMode: TDMLMode);
begin
  case aMode of
    dmBrowser, dmCancel, dmLock:
      begin
        DetailGridView.OptionsSelection.CellSelect := False;
        DetailGridView.OptionsView.ShowEditButtons := gsebNever;
        DetailGridView.Styles.Content := dtmMain.cxGridBackGroundStyle;
        DetailGridView.GetColumnByFieldName( 'SEQ' ).Styles.Content := nil;
        DetailGridView.GetColumnByFieldName( 'BILLID' ).Styles.Content := nil;
        DetailGridView.GetColumnByFieldName( 'BILLIDITEMNO' ).Styles.Content := nil;
        DetailGridView.GetColumnByFieldName( 'STARTDATE' ).Styles.Content := nil;
        DetailGridView.GetColumnByFieldName( 'ENDDATE' ).Styles.Content := nil;
        DetailGridView.GetColumnByFieldName( 'SERVICETYPE' ).Styles.Content := nil;
        DetailGridView.GetColumnByFieldName( 'TAXNAME' ).Styles.Content := nil;
        DetailGridView.GetColumnByFieldName( 'SIGN' ).Styles.Content := nil;
        DetailGridLevel2.Visible := True;
      end;
    dmInsert:
      begin
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
      end;
    dmUpdate:
      begin
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
      end;
  end;
  FillDetailColumnPropertiesValue;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.ButtonStateChange(const aMode: TDMLMode);
begin
  case aMode of
    dmBrowser, dmCancel:
      begin
        btnLocaleCreate.Enabled := True;
        btnNormalCreate.Enabled := True;
        btnEdit.Enabled := True;
        btnSave.Enabled := False;
        btnCancel.Enabled := False;
        btnPrint.Enabled := True;
        btnDrop.Enabled := True;
      end;
    dmInsert, dmUpdate:
      begin
        btnLocaleCreate.Enabled := False;
        btnNormalCreate.Enabled := False;
        btnEdit.Enabled := False;
        btnSave.Enabled := True;
        btnCancel.Enabled := True;
        btnPrint.Enabled := False;
        btnDrop.Enabled := False;
      end;
    dmLock:
      begin
        btnLocaleCreate.Enabled := False;
        btnNormalCreate.Enabled := False;
        btnEdit.Enabled := False;
        btnSave.Enabled := False;
        btnCancel.Enabled := False;
        btnPrint.Enabled := False;
        btnDrop.Enabled := False;
      end;
  end;
  { �D�d�߼Ҧ� }
  if not ( aMode in [dmLock] ) then
  begin
    { �S���W�h 1 }
    if dtmMain.GetIfPrintCheck then
    begin
      btnPrint.Enabled := ( ( not cdsMaster.IsEmpty ) and
         ( cdsMaster.FieldByName( 'INVFORMAT' ).AsString = '1' ) and
         ( cdsMaster.FieldByName( 'CANMODIFY' ).AsString = 'Y' ) and
         ( aMode in [dmBrowser, dmCancel] ) );
    end else
    begin
      btnPrint.Enabled := ( ( not cdsMaster.IsEmpty ) and
         ( cdsMaster.FieldByName( 'INVFORMAT' ).AsString = '1' ) and
         ( aMode in [dmBrowser, dmCancel] ) );
    end;
    { �S���W�h 2 }
    if not dtmMain.GetLinkToMIS then
      btnLocaleCreate.Enabled := False;
  end;    
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.EditorStateChange(const aMode: TDMLMode);
begin
  if ( aMode in [dmBrowser, dmLock] ) then
  begin
    txtINVDATE.Enabled := True;
    txtINVDATE.Properties.PopupControl := nil;
    grpTaxRate.Enabled := False;
    txtTAXRATE.Enabled := True;
    txtINVID.Properties.ReadOnly := True;
    txtMainInvId.Properties.ReadOnly := True;
    txtInvUse.Enabled := False;
  end
  else if ( aMode in [dmInsert] ) then
  begin
    txtINVDATE.Enabled := True;
    txtINVDATE.Properties.PopupControl := ChooseGrid;
    txtCUSTID.Properties.ReadOnly := False;
    txtCUSTID.Style.TextColor :=
      txtCUSTID.Style.StyleController.Style.TextColor;
    grpTaxRate.Enabled := True;
    txtTAXRATE.Enabled := True;
    txtINVID.Properties.ReadOnly := True;
    txtMainInvId.Properties.ReadOnly := True;
    txtInvUse.Enabled := True;
  end
  else if ( aMode in [dmUpdate] ) then
  begin
    { �o������i�H�ק�, ���i���s��ܵo����, ���󬰥H�U�榡,
      ��u�G�p = InvFormat = 2 , ��u�T�p = InvFormat = 3  }
    txtINVDATE.Enabled := False;
    if ( cdsMaster.FieldByName( 'INVFORMAT' ).AsString = '2' ) or
       ( cdsMaster.FieldByName( 'INVFORMAT' ).AsString = '3' ) then
    begin
      txtINVDATE.Enabled := True;
      txtINVDATE.Properties.PopupControl := nil;
    end;
    // 1. �Ƚs: �@��}�ߪ��o���~�i�ק�, ���O�D�@��}�ߪ��o�����u�X�������ƥ\��,
    //           ���M�n��ϥ�
    // 2. �|�O: �@��}�ߪ��o���~�i�ק�(�p�G�O�h�i�۰ʶ}�ߪ����i�H�ק�)
    if FInvCreateType = icNormal then
    begin
      txtCUSTID.Properties.ReadOnly := False;
      txtCUSTID.Style.TextColor :=
        txtCUSTID.Style.StyleController.Style.TextColor;
      grpTaxRate.Enabled := not FMultiInv;
      txtTAXRATE.Enabled := not FMultiInv;
    end
    else begin
      txtCUSTID.Properties.ReadOnly := True;
      txtCUSTID.Style.TextColor :=
        txtCUSTID.Style.StyleController.StyleDisabled.TextColor;
      grpTaxRate.Enabled := False;
      txtTAXRATE.Enabled := False;
    end;
    txtINVID.Properties.ReadOnly := True;
    txtMainInvId.Properties.ReadOnly := True;
    // �{���}�ߩΤ@��}��, �~�i�H�ק�o���γ~���, ���ɹL�Ӫ��h����
    txtInvUse.Enabled := ( FInvCreateType in [icLocale, icNormal] );
  end;
  { �]�w Focus }
  case aMode of
    dmInsert, dmBrowser:
      begin
        if txtINVDATE.CanFocus then txtINVDATE.SetFocus;
      end;
    dmUpdate:
      begin
        if ( txtINVDATE.CanFocusEx ) then
        begin
          txtINVDATE.SetFocus;
        end else
        if not ( txtCUSTID.Properties.ReadOnly ) then
        begin
          if txtCUSTID.CanFocus then txtCUSTID.SetFocus
        end else
        begin
          if txtCUSTSNAME.CanFocus then txtCUSTSNAME.SetFocus;
        end;  
      end;
  end;    
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA02_4.DataSetStateChange(const aMode: TDMLMode): Boolean;
var
  aErrMsg: String;
begin
  Result := True;
  aErrMsg := '';
  case aMode of
    dmBrowser, dmLock:
      begin
        cdsMaster.ReadOnly := True;
        cdsDetail.ReadOnly := True;
        cdsMasterAfterScroll( cdsMaster );
      end;
    dmCancel:
      begin
        cdsMaster.Cancel;
        cdsDetail.CancelUpdates;
        cdsMaster.ReadOnly := True;
        cdsDetail.ReadOnly := True;
        FBillingInfo.BillId := EmptyStr;
        FBillingInfo.ItemNo := EmptyStr;
        cdsMasterAfterScroll( cdsMaster );
      end;  
    dmInsert:
      begin
        cdsMaster.ReadOnly := False;
        cdsDetail.ReadOnly := False;
        cdsMaster.Append;
        cdsDetail.Append;
      end;
    dmUpdate:
      begin
        cdsMaster.ReadOnly := False;
        cdsDetail.ReadOnly := False;
        cdsMaster.Edit;
      end;
    dmSave:
      begin
        { ���ˬd Detail }
        Result := ValidateDetailDataSet( aErrMsg );
        if not Result then
        begin
          WarningMsg( aErrMsg );
          Exit;
        end;
        { �ˬd Master }
        Result := ValidateMasterDataSet( aErrMsg );
        if not Result then
        begin
          WarningMsg( aErrMsg );
          Exit;
        end;
        { �s�� }
        Result := ApplyDataSet( aErrMsg );
        if not Result then
        begin
          WarningMsg( aErrMsg );
          Exit;
        end;
        { ���s Open DataSet }
        FInvoiceId := cdsMaster.FieldByName( 'INVID' ).AsString;
        Self.FormShow( Self );
      end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA02_4.ValidateCanModify(var aErrMsg: String): Boolean;
begin
  aErrMsg := EmptyStr;
  Result := True;
  try
    if cdsMaster.IsEmpty then
    begin
      aErrMsg := '�L�o�����, ���i�ק�C';
      Exit;
    end;
    if cdsMaster.FieldByName( 'ISOBSOLETEDESC' ).AsString = '�O' then
    begin
      aErrMsg := '�����o���w�@�o!';
      Exit;
    end;
    if cdsMaster.FieldByName( 'CANMODIFY' ).AsString <> 'Y' then
    begin
      aErrMsg := '�����o���ݰ��L���v�~�i�ק�!';
      Exit;
    end;
{ ���q�����D, �o���O�_��b, �ˬd INV099.USEFUL �O����, �����ˬd INV018 }
//    aPrefix := Copy( cdsMaster.FieldByName( 'INVID' ).AsString, 1, 2 );
//    aNum := Copy( cdsMaster.FieldByName( 'INVID' ).AsString, 3, 8 );
//    dtmMain.adoComm.Close;
//    dtmMain.adoComm.SQL.Text := Format(
//      ' SELECT CURNUM FROM INV099  ' +
//      '  WHERE USEFUL = ''Y''      ' +
//      '    AND PREFIX = ''%s''     ' +
//      '    AND ENDNUM >= ''%s''    ' +
//      '    AND STARTNUM <= ''%s''  ',
//      [ aPrefix, aNum, aNum ] );
//    dtmMain.adoComm.Open;
    dtmMain.adoComm.Close;
    dtmMain.adoComm.SQL.Text := Format(
      ' select count(1) from inv099 a, inv018 b   ' +
      '  where a.compid = b.compid                ' +
      '    and a.yearmonth = b.yearmonth          ' +
      '    and b.islocked = ''Y''                 ' +
      '    and a.compid = ''%s''                  ' +
      '    and ''%s'' between startnum and endnum ',
      [dtmMain.getCompID, cdsMaster.FieldByName( 'INVID' ).AsString] );
    dtmMain.adoComm.Open;
    if ( dtmMain.adoComm.Fields[0].AsInteger > 0 ) then
    begin
      aErrMsg := '�����o���w�L��(�Ӥ���w��b�εo�����X���b���ĵo������),���i�ק�C';
      Exit;
    end;
  finally
    Result := ( aErrMsg = '' );
    dtmMain.adoComm.Close;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA02_4.ValidateCanDrop(var aErrMsg: String): Boolean;
var
  aPrefix, aNum: String;
begin
  aErrMsg := EmptyStr;
  Result := True;
  try
    if cdsMaster.IsEmpty then
    begin
      aErrMsg := '�L�o�����, ���i���@�o�C';
      Exit;
    end;
    if ( cdsMaster.FieldByName( 'ISOBSOLETEDESC' ).AsString = '�O' ) then
    begin
      aErrMsg := '�����o���w�@�o�C';
      Exit;
    end;
    if ( cdsMaster.FieldByName( 'CANMODIFY' ).AsString <> 'Y' ) then
    begin
      aErrMsg := '�����o���ݰ��L���v�~�i�@�o�C';
      Exit;
    end;
    aPrefix := Copy( cdsMaster.FieldByName( 'INVID' ).AsString, 1, 2 );
    aNum := Copy( cdsMaster.FieldByName( 'INVID' ).AsString, 3, 8 );
    dtmMain.adoComm.Close;
    dtmMain.adoComm.SQL.Text := Format(
      ' SELECT CURNUM FROM INV099  ' +
      '  WHERE USEFUL = ''Y''      ' +
      '    AND PREFIX = ''%s''     ' +
      '    AND ENDNUM >= ''%s''    ' +
      '    AND STARTNUM <= ''%s''  ',
      [ aPrefix, aNum, aNum ] );
    dtmMain.adoComm.Open;
    if ( dtmMain.adoComm.IsEmpty ) then
    begin
      aErrMsg := '�����o���w�L��(�Ӥ���w��b�εo�����X���b���ĵo������), ���i�@�o�C';
      Exit;
    end;
  finally
    Result := ( aErrMsg = '' );
    dtmMain.adoComm.Close;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA02_4.ValidateCanPrint(var aErrMsg: String): Boolean;
begin
  aErrMsg := '';
  Result := True;
  try
    if cdsMaster.IsEmpty then
    begin
      aErrMsg := '�L�o�����, ���i�M�L�C';
      Exit;
    end;
    if cdsMaster.FieldByName( 'INVFORMAT' ).AsString <> '1' then
    begin
      aErrMsg := '�����o���榡�������i�q�l�j, �~�i�M�L�����o���C';
      Exit;
    end;
    if dtmMain.GetIfPrintCheck then
    begin
      if cdsMaster.FieldByName( 'CANMODIFY' ).AsString <> 'Y' then
      begin
        aErrMsg := '�����o���ݰ��L���v�~�i�M�L!';
        Exit;
      end;
    end;  
    if cdsMaster.FieldByName( 'ISOBSOLETEDESC' ).AsString = '�O' then
    begin
      aErrMsg := '�����o���w���o, ���i�M�L!';
      Exit;
    end;
  finally
    Result := ( aErrMsg = '' );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA02_4.ValidateMasterDataSet(var aErrMsg: String): Boolean;
var
  aFocusControl: TWinControl;
begin
  Result := False;
  aFocusControl := nil;
  try
    if ( cdsMaster.FieldByName( 'INVID' ).AsString = EmptyStr ) then
    begin
      aErrMsg := '�п�ܵo�����X�C';
      aFocusControl := txtINVDATE;
      Exit;
    end;
    if ( cdsMaster.FieldByName( 'INVDATE' ).AsString = EmptyStr ) then
    begin
      aErrMsg := '�п�J�o������C';
      aFocusControl := txtINVDATE;
      Exit;
    end;
    if ( cdsMaster.FieldByName( 'CUSTID' ).AsString = EmptyStr ) then
    begin
      aErrMsg := '�п�J�Ƚs�C';
      aFocusControl := txtCUSTID;
      Exit;
    end;
    if ( cdsMaster.FieldByName( 'CUSTSNAME' ).AsString = EmptyStr ) then
    begin
      aErrMsg := '�п�J�Ȥ�W�١C';
      aFocusControl := txtCUSTSNAME;
      Exit;
    end;
    {}
    if ( cdsMaster.FieldByName( 'INVFORMAT' ).AsString = '3' ) then
    begin
      if ( cdsMaster.FieldByName( 'BUSINESSID' ).AsString = EmptyStr ) then
      begin
        aErrMsg := '���i����u�T�p���o��, ������J�νs�C';
        aFocusControl := txtBUSINESSID;
        Exit;
      end;
    end;
    if ( cdsMaster.FieldByName( 'INVFORMAT' ).AsString = '2' ) then
    begin
      cdsMaster.FieldByName( 'BUSINESSID' ).AsString := EmptyStr;
    end;
    {}
    if ( cdsMaster.FieldByName( 'BUSINESSID' ).AsString <> EmptyStr ) and
       ( txtInvUse.Value <> EmptyStr ) then
    begin
      if not ValidateCanSetInvUse( txtInvUse.Value ) then
      begin
        aErrMsg := '(��~�H)�o�����i���i���ءj�C';
        aFocusControl := txtInvUse.CodeNo;
        Exit;
      end;
    end;
    {}
    if ( cdsMaster.FieldByName( 'INVAMOUNT' ).AsFloat <= 0 ) then
    begin
      aErrMsg := '�o���`���B���i���s�άO�t�ȡC';
      Exit;
    end;
    if ( GetMasterTaxCode = '1' ) and
       ( cdsMaster.FieldByName( 'TAXRATE' ).AsFloat <= 0 ) then
    begin
      aErrMsg := '�������|�o��, �|�v�����j��s�C';
      aFocusControl := txtTAXRATE;
      Exit;
    end;
    {}
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

function TfrmInvA02_4.ValidateDetailDataSet(var aErrMsg: String): Boolean;
var
  aBookMark: TBookMark;
  aFocusColumn: TcxGridDBBandedColumn;
  aTaxRate, aUnitPtrice, aSaleAmount, aTaxAmount, aTotalAmount: Double;
  aTotalSum: Double;
  aInvCount, aWillCreateCount, aIndex: Integer;
begin
  Result := False;
  aFocusColumn := nil;
  try
    try
      cdsDetail.CheckBrowseMode;
    except
      on E: Exception do
      begin
        aErrMsg := E.Message;
        Exit;
      end;
    end;
    if ( cdsDetail.IsEmpty ) then
    begin
      cdsDetail.Append;
      aErrMsg := '�o�����ӥ�����J�C';
      Exit;
    end;
    aTaxRate := cdsMaster.FieldByName( 'TAXRATE' ).AsFloat;
    aTotalSum := 0;
    cdsDetail.DisableControls;
    try
      aBookMark := cdsDetail.GetBookmark;
      try
        cdsDetail.First;
        while not cdsDetail.Eof do
        begin
          if ( FInvCreateType in [icPrev, icAfter, icLocale] ) then
          begin
            { ���O�渹�ˬd }
            if ( not dtmMain.GetLinkToMIS ) or
               ( cdsDetail.FieldByName( 'LINKTOMIS' ).AsString <> 'Y' ) then
            begin
               { ���P�ȪA�t�Τ���, �ˬd�O�_��J���O�渹�ζ����Y�i }
               if ( cdsDetail.FieldByName( 'BILLID' ).AsString = EmptyStr ) or
                  ( cdsDetail.FieldByName( 'BILLIDITEMNO' ).AsString = EmptyStr ) then
               begin
                 aErrMsg := '���O���Ӹ��,�п�J���O�渹�Φ��O�����C';
                 aFocusColumn := DetailGridView.GetColumnByFieldName( 'BILLID' );
                 Exit;
               end;
            end;
            { ���O�渹�ζ��� }
            if ( dtmMain.GetLinkToMIS ) and
               ( cdsDetail.FieldByName( 'LINKTOMIS' ).AsString = 'Y' ) then
            begin
              Result := ValidateBillingNo( cdsDetail.FieldByName( 'BILLID' ).AsString,
                cdsDetail.FieldByName( 'BILLIDITEMNO' ).AsString, aErrMsg,
                ( cdsDetail.FieldByName( 'EMPTY' ).AsString = 'N' ) );
              if not Result then
              begin
                aFocusColumn := DetailGridView.GetColumnByFieldName( 'BILLID' );
                Exit;
              end;
            end;
          end;
          { �A�ȧO�ˬd }
          if ( FInvCreateType in [icPrev, icAfter, icLocale] ) and
             ( cdsDetail.FieldByName( 'SERVICETYPE' ).AsString = EmptyStr ) then
          begin
            aErrMsg := '���O���Ӹ��,�п�J�A�ȧO�C';
            aFocusColumn := DetailGridView.GetColumnByFieldName( 'SERVICETYPE' );
            Exit;
          end;
          { ���O�����ˬd }
          if ( cdsDetail.FieldByName( 'ITEMID' ).AsString = EmptyStr ) or
             ( cdsDetail.FieldByName( 'DESCRIPTION' ).AsString = EmptyStr ) then
          begin
            aFocusColumn := DetailGridView.GetColumnByFieldName( 'ITEMID' );
            aErrMsg := '���O���Ӹ��, �п�J���O���ءC';
            Exit;
          end;
          if ( ( dtmMain.GetLinkToMIS ) and
               ( cdsDetail.FieldByName( 'LINKTOMIS' ).AsString = 'Y' ) ) or
             ( FInvCreateType in [icNormal] ) then
          begin
            Result := ValidateChargeItemId( cdsDetail.FieldByName( 'ITEMID' ).AsString,
              aErrMsg );
            if not Result then
            begin
              aFocusColumn := DetailGridView.GetColumnByFieldName( 'ITEMID' );
              Exit;
            end;
          end;  
          aUnitPtrice := cdsDetail.FieldByName( 'UNITPRICE' ).AsFloat;
          aSaleAmount := cdsDetail.FieldByName( 'SALEAMOUNT' ).AsFloat;
          aTaxAmount := cdsDetail.FieldByName( 'TAXAMOUNT' ).AsFloat;
          aTotalAmount := cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsFloat;
          { �|�O�ε|�v�ˬd - �K�| }
          if ( cdsDetail.FieldByName( 'TAXCODE' ).AsString <> '' ) then
          begin
            if ( GetMasterTaxCode <> cdsDetail.FieldByName( 'TAXCODE' ).AsString ) then
            begin
              aErrMsg := '���Ӹ��,�|�O�P�o�����šC';
              aFocusColumn := DetailGridView.GetColumnByFieldName( 'ITEMID' );
              Exit;
            end;
          end;
          { ���B�ˬd }
          if ( GetMasterTaxCode <> '1' ) then
          begin
            if ( aTaxAmount <> 0 ) then
            begin
              aErrMsg := '�����K�|�o��,���ӵ|�B���i��J�C';
              aFocusColumn := DetailGridView.GetColumnByFieldName( 'TAXAMOUNT' );
              Exit;
            end
          end else
          begin
            { �Y�����|, �`���B�� 10 �������� -10 ����, �h��X�Ӫ��|�B�|�O�s,
              �ҥH�j��10������-10�����~�|���|�B } 
            if ( aTaxAmount = 0 ) and ( Abs( aTotalAmount ) > 10 ) then
            begin
              aErrMsg := '�������|�o��,���ӵ|�B������J�C';
              aFocusColumn := DetailGridView.GetColumnByFieldName( 'TAXAMOUNT' );
              Exit;
            end;
          end;
          { �P���B�ˬd }
          if ( aSaleAmount = 0 ) then
          begin
            aErrMsg := '���Ӹ��, �P���B������J�C';
            aFocusColumn := DetailGridView.GetColumnByFieldName( 'SALEAMOUNT' );
            Exit;
          end;
          { �`���B }
          if ( aTotalAmount = 0 ) then
          begin
            aErrMsg := '���Ӹ��, �`���B������J�C';
            aFocusColumn := DetailGridView.GetColumnByFieldName( 'TOTALAMOUNT' );
            Exit;
          end;
          { �t�����O }
          if ( cdsDetail.FieldByName( 'SIGN' ).AsString = '-' ) then
          begin
            if ( aUnitPtrice > 0 ) then
            begin
              aErrMsg := '���Ӹ�Ƭ��t�����B, ����������t�ȡC';
              aFocusColumn := DetailGridView.GetColumnByFieldName( 'UNITPRICE' );
              Exit;
            end;
            if ( aTaxAmount > 0 ) then
            begin
              aErrMsg := '���Ӹ�Ƭ��t�����B, �|�B�������t�ȡC';
              aFocusColumn := DetailGridView.GetColumnByFieldName( 'TAXAMOUNT' );
              Exit;
            end;
            if ( aSaleAmount > 0 ) then
            begin
              aErrMsg := '���Ӹ�Ƭ��t�����B, �P���B�������t�ȡC';
              aFocusColumn := DetailGridView.GetColumnByFieldName( 'SALEAMOUNT' );
              Exit;
            end;
            if ( aTotalAmount > 0 ) then
            begin
              aErrMsg := '���Ӹ�Ƭ��t�����B, �`���B�B�������t�ȡC';
              aFocusColumn := DetailGridView.GetColumnByFieldName( 'TOTALAMOUNT' );
              Exit;
            end;
          end else
          begin
            if ( aUnitPtrice < 0 ) then
            begin
              aErrMsg := '���Ӹ�Ƭ��������B, ������������ȡC';
              aFocusColumn := DetailGridView.GetColumnByFieldName( 'UNITPRICE' );
              Exit;
            end;
            if ( aTaxAmount < 0 ) then
            begin
              aErrMsg := '���Ӹ�Ƭ��������B, �|�B���������ȡC';
              aFocusColumn := DetailGridView.GetColumnByFieldName( 'TAXAMOUNT' );
              Exit;
            end;
            if ( aSaleAmount < 0 ) then
            begin
              aErrMsg := '���Ӹ�Ƭ��������B, �P���B���������ȡC';
              aFocusColumn := DetailGridView.GetColumnByFieldName( 'SALEAMOUNT' );
              Exit;
            end;
            if ( aTotalAmount < 0 ) then
            begin
              aErrMsg := '���Ӹ�Ƭ��������B, �`���B�B���������ȡC';
              aFocusColumn := DetailGridView.GetColumnByFieldName( 'TOTALAMOUNT' );
              Exit;
            end;
          end;
          {}
          aTotalSum := ( aTotalSum + aTotalAmount );
          cdsDetail.Next;
        end;
        {}
        if ( aTotalSum <= 0 ) and ( aErrMsg = '' ) then
        begin
          aErrMsg := '���Ӹ��, �o���[�`���B�p��ε���s�C';
          aFocusColumn := nil;
          Exit;
        end;
        // ���}�o���i���ˮ�
        if ( dtmMain.GetAutoCreateNum > 0 ) and
           ( cdsDetail.RecordCount > dtmMain.GetAutoCreateNum ) then
        begin
          if ( FMode in [dmInsert] ) then
          begin
            aInvCount := GetCurrentAvailableInvCount;
            aWillCreateCount := Ceil( cdsDetail.RecordCount / dtmMain.GetAutoCreateNum );
            if ( aWillCreateCount > aInvCount ) then
            begin
              aErrMsg := Format( '�����w�p�}�ߵo���i�Ƭ� %d �i, �Ѿl�i�εo���i�Ƭ� %d �i, �o���i�Ƥ���, �Э��s��ܵo�����C',
                [aWillCreateCount, aInvCount] );
              aFocusColumn := nil;
              Exit;
            end;
          end else
          if ( FMode in [dmUpdate] ) then
          begin
            aErrMsg := Format( '�o�����ӵ��Ƥw�W�X���}���Ƴ]�w, �]�w�����ƭ�� %d ���C',
              [dtmMain.GetAutoCreateNum] );
            aFocusColumn := nil;
            Exit;
          end;  
        end;
        cdsDetail.GotoBookmark( aBookMark );
      finally
        cdsDetail.FreeBookmark( aBookMark );
      end;
    finally
      cdsDetail.EnableControls;
    end;
  finally
    Result := ( aErrMsg = EmptyStr );
    if not Result then
    begin
      if Assigned( aFocusColumn ) then
        aFocusColumn.Focused := True;
      DetailGrid.SetFocus;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA02_4.ApplyDataSet(var aErrMsg: String): Boolean;

type
  PAutoCreate = ^TAutoCreate;
  TAutoCreate = record  //�۰ʵo�����}�O��
    InvId: String;
    InvDate: String;
    MainInvId: String;
    MainSaleAmount: Double;
    MainTaxAmount: Double;
    MainInvAmount: Double;
    SaleAmount: Double;
    TaxAmount: Double;
    InvAmount: Double;
    ChargeDate: String;
    CustId: String;
    CustName: String;
    InvTitle: String;
    ZipCode: String;
    InvAddr: String;
    MailAddr: String;
    InstAddr: String;
    BusinessId: String;
    Memo1: String;
    Memo2: String;
  end;

  PSimpleAutoCreate = ^TSimpleAutoCreate;
  TSimpleAutoCreate = record
    InvId: String;
    Count: Integer;
  end;

var
  aDataSet, aSoDataSet: TADOQuery;
  aConnection, aSoConnection: TADOConnection;

  { -------------------------------------------------------- }

  procedure DeleteAllInvoiceDetail(const aInvId: String);
  begin
    aDataSet.SQL.Text := Format(
      ' DELETE FROM INV008               ' +
      '  WHERE IDENTIFYID1 = ''%s''      ' +
      '    AND IDENTIFYID2 = %s          ' +
      '    AND INVID = ''%s''            ',
      [IDENTIFYID1, IDENTIFYID2, aInvId] );
    aDataSet.ExecSQL;
  end;

  { -------------------------------------------------------- }
  
  { ���N SO033 �� SO034 ���o�����O�ӷ�����ƥ���@��, }
  { �]�� Detail ���ެO�s�W�άO�ק�, �@�߳��O���ΧR�����覡 }
  procedure CloneInvoiceFromSO;
  begin
    cdsSO.Close;
    cdsSO.CommandText := Format(
      '  SELECT 33 AS SOURCETABLE,   ' +
      '         A.BILLNO,            ' +
      '         A.ITEM,              ' +
      '         A.INVOICETIME        ' +
      '    FROM %s.SO033 A           ' +
      '   WHERE A.CUSTID = ''%s''    ' +
      '     AND A.GUINO = ''%s''     ' +
      '  UNION ALL                   ' +
      '  SELECT 34 AS SOURCETABLE,   ' +
      '         A.BILLNO,            ' +
      '         A.ITEM,              ' +
      '         A.INVOICETIME        ' +
      '    FROM %s.SO034 A           ' +
      '   WHERE A.CUSTID = ''%s''    ' +
      '     AND A.GUINO = ''%s''     ',
      [dtmMain.getMisDbOwner,
        cdsMaster.FieldByName( 'CUSTID' ).AsString,
        cdsMaster.FieldByName( 'INVID' ).AsString,
        dtmMain.getMisDbOwner,
        cdsMaster.FieldByName( 'CUSTID' ).AsString,
        cdsMaster.FieldByName( 'INVID' ).AsString] );
     cdsSO.Open;
  end;

  { -------------------------------------------------------- }

  procedure EmptyInvoiceToSO;
  const
    aSOTableName: array [1..2] of String = ( 'SO033', 'SO034' );
  var
    aIndex: Integer;
  begin
    aSoDataSet.Close;
    for aIndex := 1 to 2 do
    begin
      aSoDataSet.SQL.Text := Format(
        ' UPDATE %s.%s                               ' +
        '    SET GUINO = NULL,                       ' +
        '        PREINVOICE = NULL,                  ' +
        '        INVDATE = NULL,                     ' +
        '        INVOICETIME = NULL                  ' +
        '  WHERE CUSTID = ''%s''                     ' +
        '    AND GUINO = ''%s''                      ' +
        '    AND NVL( PREINVOICE, 0 ) IN ( 0, 1, 2 ) ',
      [ dtmMain.getMisDbOwner, aSOTableName[aIndex],
        cdsMaster.FieldByName( 'CUSTID' ).AsString,
        cdsMaster.FieldByName( 'INVID' ).AsString ] );
      aSoDataSet.ExecSQL;
    end;
  end;

  { -------------------------------------------------------- }

  procedure UpdateInvoiceToSO(const aInvId, aBillId, aItemId: String);
  const
    aSOTableName: array [1..2] of String = ( 'SO033', 'SO034' );
  var
    aIndex: Integer;
    aInvoiceTime: String;
  begin
    aSoDataSet.Close;
    for aIndex := 1 to 2 do
    begin
      aInvoiceTime := EmptyStr;
      { �u���ק�~���n��INVOICETIME, �]���ק�s�ɤ��e�@�ߥ��M���ȪA�t�ι�����
        ���, ��s�ɥ������ȶ�^�h }
      if ( FMode in [dmUpdate] ) then
      begin
        if cdsSO.Locate( 'BILLNO;ITEM', VarArrayOf( [aBillId, aItemId] ), [] ) then
        begin
          if ( cdsSO.FieldByName( 'INVOICETIME' ).AsDateTime <> 0 ) then
            aInvoiceTime := FormatDateTime( 'yyyymmdd hhnnss',
              cdsSO.FieldByName( 'INVOICETIME' ).AsDateTime );
        end;
      end;  
      if ( aInvoiceTime = EmptyStr ) then
        aInvoiceTime := FormatDateTime( 'yyyymmdd hhnnss', Now );
      aSoDataSet.SQL.Text := Format(
        ' UPDATE %s.%s                                       ' +
        '    SET GUINO = ''%s'',                             ' +
        '        PREINVOICE = ''%s'',                        ' +
        '        INVDATE = TO_DATE( ''%s'', ''YYYYMMDD'' ),  ' +
        '        INVOICETIME = TO_DATE( ''%s'', ''YYYYMMDD HH24MISS'' ), ' +
        '        INVPURPOSECODE = ''%s'',                    ' +
        '        INVPURPOSENAME = ''%s''                     ' +
        '  WHERE BILLNO = ''%s''                             ' +
        '    AND ITEM = ''%s''                               ',
        [ dtmMain.getMisDbOwner, aSOTableName[aIndex],
          aInvId,
          InvoiceTypeToSO( FInvCreateType ),
          FormatDateTime( 'yyyymmdd', cdsMaster.FieldByName( 'INVDATE' ).AsDateTime ),
          aInvoiceTime,
          txtInvUse.Value,
          txtInvUse.ValueName,
          aBillId, aItemId] );
      aSoDataSet.ExecSQL;
    end;
  end;

  { -------------------------------------------------------- }

  function ApplyMasterDataSet(const aMode: TDMLMode; var aObj: PAutoCreate): Boolean;
  begin
    Result := False;
    aDataSet.Close;
    aDataSet.SQL.Text := EmptyStr;
    aDataSet.Parameters.Clear;
    aDataSet.SQL.Text := BuildMasterApplySQL( aMode );
    if ( aMode in [dmInsert] ) then
    begin
      aDataSet.Parameters.ParamByName( 'IDENTIFYID1' ).Value := IDENTIFYID1;
      aDataSet.Parameters.ParamByName( 'IDENTIFYID2' ).Value := IDENTIFYID2;
      aDataSet.Parameters.ParamByName( 'CHECKNO' ).Value  := GetInvoiceCheckNumber;
      aDataSet.Parameters.ParamByName( 'INVID' ).Value := aObj.InvId;
      aDataSet.Parameters.ParamByName( 'CANMODIFY' ).Value  := 'Y';
      aDataSet.Parameters.ParamByName( 'INVDATE' ).Value :=
        cdsMaster.FieldByName( 'INVDATE' ).AsDateTime;
      if ( cdsMaster.FieldByName( 'CHARGEDATE' ).AsString <> EmptyStr ) then
      begin
        aDataSet.Parameters.ParamByName( 'CHARGEDATE' ).Value :=
          cdsMaster.FieldByName( 'CHARGEDATE' ).AsDateTime;
      end
      else begin
        aDataSet.Parameters.ParamByName( 'CHARGEDATE' ).Value := EmptyStr;
      end;
      aDataSet.Parameters.ParamByName( 'COMPID' ).Value := dtmMain.getCompID;
      aDataSet.Parameters.ParamByName( 'CUSTID' ).Value :=
        cdsMaster.FieldByName( 'CUSTID' ).AsString;
      aDataSet.Parameters.ParamByName( 'CUSTSNAME' ).Value :=
        cdsMaster.FieldByName( 'CUSTSNAME' ).AsString;
      aDataSet.Parameters.ParamByName( 'INVTITLE' ).Value :=
        cdsMaster.FieldByName( 'INVTITLE' ).AsString;
      aDataSet.Parameters.ParamByName( 'ZIPCODE' ).Value :=
        cdsMaster.FieldByName( 'ZIPCODE' ).AsString;
      aDataSet.Parameters.ParamByName( 'INVADDR' ).Value :=
        cdsMaster.FieldByName( 'INVADDR' ).AsString;
      aDataSet.Parameters.ParamByName( 'MAILADDR' ).Value :=
        cdsMaster.FieldByName( 'MAILADDR' ).AsString;
      aDataSet.Parameters.ParamByName( 'BUSINESSID' ).Value :=
        cdsMaster.FieldByName( 'BUSINESSID' ).AsString;
      aDataSet.Parameters.ParamByName( 'INVFORMAT' ).Value :=
        cdsMaster.FieldByName( 'INVFORMAT' ).AsString;
      aDataSet.Parameters.ParamByName( 'TAXTYPE' ).Value := GetMasterTaxCode;
      aDataSet.Parameters.ParamByName( 'TAXRATE' ).Value :=
        cdsMaster.FieldByName( 'TAXRATE' ).AsString;
      aDataSet.Parameters.ParamByName( 'SALEAMOUNT' ).Value :=
        cdsMaster.FieldByName( 'SALEAMOUNT' ).AsString;
      aDataSet.Parameters.ParamByName( 'TAXAMOUNT' ).Value :=
        cdsMaster.FieldByName( 'TAXAMOUNT' ).AsString;
      aDataSet.Parameters.ParamByName( 'INVAMOUNT' ).Value :=
        cdsMaster.FieldByName( 'INVAMOUNT' ).AsString;
      aDataSet.Parameters.ParamByName( 'PRINTCOUNT' ).Value := 0;
      aDataSet.Parameters.ParamByName( 'ISPAST' ).Value := 'N';
      aDataSet.Parameters.ParamByName( 'ISOBSOLETE' ).Value := 'N';
      aDataSet.Parameters.ParamByName( 'HOWTOCREATE' ).Value := Ord( FInvCreateType );
      //#5358
      if aObj.Memo1 <> '' then
      begin
          aDataSet.Parameters.ParamByName( 'MEMO1' ).Value :=
            cdsMaster.FieldByName( 'MEMO1' ).AsString + ' �]��:' + aObj.Memo1;

      end else
      begin
        aDataSet.Parameters.ParamByName( 'MEMO1' ).Value :=
          cdsMaster.FieldByName( 'MEMO1' ).AsString;
      end;
      aDataSet.Parameters.ParamByName( 'MEMO2' ).Value :=
        cdsMaster.FieldByName( 'MEMO2' ).AsString;
      aDataSet.Parameters.ParamByName( 'UPTEN' ).Value := dtmMain.getLoginUserName;
      aDataSet.Parameters.ParamByName( 'PRINTFUN' ).Value := '0';
      {}
      aDataSet.Parameters.ParamByName( 'MAININVID' ).Value := aObj.MainInvId;
      aDataSet.Parameters.ParamByName( 'MAINSALEAMOUNT' ).Value := aObj.MainSaleAmount;
      aDataSet.Parameters.ParamByName( 'MAINTAXAMOUNT' ).Value := aObj.MainTaxAmount;
      aDataSet.Parameters.ParamByName( 'MAININVAMOUNT' ).Value := aObj.MainInvAmount;
      aDataSet.Parameters.ParamByName( 'INSTADDR' ).Value := aObj.InstAddr;
      aDataSet.Parameters.ParamByName( 'INVUSEID' ).Value := txtInvUse.Value;
      aDataSet.Parameters.ParamByName( 'INVUSEDESC' ).Value := txtInvUse.ValueName;
    end
    else begin
      aDataSet.Parameters.ParamByName( 'IDENTIFYID1' ).Value := IDENTIFYID1;
      aDataSet.Parameters.ParamByName( 'IDENTIFYID2' ).Value := IDENTIFYID2;
      aDataSet.Parameters.ParamByName( 'INVID' ).Value  := aObj.InvId;
      aDataSet.Parameters.ParamByName( 'INVDATE' ).Value :=
        cdsMaster.FieldByName( 'INVDATE' ).AsDateTime;
      if ( cdsMaster.FieldByName( 'CHARGEDATE' ).AsString <> EmptyStr ) then
      begin
        aDataSet.Parameters.ParamByName( 'CHARGEDATE' ).Value :=
          cdsMaster.FieldByName( 'CHARGEDATE' ).AsDateTime;
      end
      else begin
        aDataSet.Parameters.ParamByName( 'CHARGEDATE' ).Value := EmptyStr;
      end;    
      aDataSet.Parameters.ParamByName( 'CUSTID' ).Value :=
        cdsMaster.FieldByName( 'CUSTID' ).AsString;
      aDataSet.Parameters.ParamByName( 'CUSTSNAME' ).Value :=
        cdsMaster.FieldByName( 'CUSTSNAME' ).AsString;
      aDataSet.Parameters.ParamByName( 'INVTITLE' ).Value :=
        cdsMaster.FieldByName( 'INVTITLE' ).AsString;
      aDataSet.Parameters.ParamByName( 'ZIPCODE' ).Value :=
        cdsMaster.FieldByName( 'ZIPCODE' ).AsString;
      aDataSet.Parameters.ParamByName( 'INVADDR' ).Value :=
        cdsMaster.FieldByName( 'INVADDR' ).AsString;
      aDataSet.Parameters.ParamByName( 'MAILADDR' ).Value :=
        cdsMaster.FieldByName( 'MAILADDR' ).AsString;
      aDataSet.Parameters.ParamByName( 'BUSINESSID' ).Value :=
        cdsMaster.FieldByName( 'BUSINESSID' ).AsString;
      aDataSet.Parameters.ParamByName( 'TAXTYPE' ).Value := GetMasterTaxCode;
      aDataSet.Parameters.ParamByName( 'TAXRATE' ).Value :=
        cdsMaster.FieldByName( 'TAXRATE' ).AsString;
      aDataSet.Parameters.ParamByName( 'SALEAMOUNT' ).Value :=
        cdsMaster.FieldByName( 'SALEAMOUNT' ).AsString;
      aDataSet.Parameters.ParamByName( 'TAXAMOUNT' ).Value :=
        cdsMaster.FieldByName( 'TAXAMOUNT' ).AsString;
      aDataSet.Parameters.ParamByName( 'INVAMOUNT' ).Value :=
        cdsMaster.FieldByName( 'INVAMOUNT' ).AsString;
      aDataSet.Parameters.ParamByName( 'MEMO1' ).Value :=
        cdsMaster.FieldByName( 'MEMO1' ).AsString;
      aDataSet.Parameters.ParamByName( 'MEMO2' ).Value :=
        cdsMaster.FieldByName( 'MEMO2' ).AsString;
      aDataSet.Parameters.ParamByName( 'UPTEN' ).Value := dtmMain.getLoginUserName;
      aDataSet.Parameters.ParamByName( 'MAININVID' ).Value := aObj.MainInvId;
      aDataSet.Parameters.ParamByName( 'MAINSALEAMOUNT' ).Value := aObj.MainSaleAmount;
      aDataSet.Parameters.ParamByName( 'MAINTAXAMOUNT' ).Value := aObj.MainTaxAmount;
      aDataSet.Parameters.ParamByName( 'MAININVAMOUNT' ).Value := aObj.MainInvAmount;
      aDataSet.Parameters.ParamByName( 'INSTADDR' ).Value := aObj.InstAddr;
      aDataSet.Parameters.ParamByName( 'INVUSEID' ).Value := txtInvUse.Value;
      aDataSet.Parameters.ParamByName( 'INVUSEDESC' ).Value := txtInvUse.ValueName;
    end;
    aDataSet.ExecSQL;
    aObj.InvDate := cdsMaster.FieldByName( 'INVDATE' ).AsString;
    aObj.ChargeDate := cdsMaster.FieldByName( 'CHARGEDATE' ).AsString;
    aObj.CustId := cdsMaster.FieldByName( 'CUSTID' ).AsString;
    aObj.CustName := cdsMaster.FieldByName( 'CUSTSNAME' ).AsString;
    aObj.InvTitle := cdsMaster.FieldByName( 'INVTITLE' ).AsString;
    aObj.ZipCode := cdsMaster.FieldByName( 'ZIPCODE' ).AsString;
    aObj.InvAddr := cdsMaster.FieldByName( 'INVADDR' ).AsString;
    aObj.MailAddr := cdsMaster.FieldByName( 'MAILADDR' ).AsString;
    aObj.InstAddr := cdsMaster.FieldByName( 'INSTADDR' ).AsString;
    aObj.BusinessId := cdsMaster.FieldByName( 'BUSINESSID' ).AsString;
    aObj.Memo1 := cdsMaster.FieldByName( 'MEMO1' ).AsString;
    aObj.Memo2 := cdsMaster.FieldByName( 'MEMO2' ).AsString;
  end;

  { -------------------------------------------------------- }

  function ApplyDetailDataSetByRecord(const aInvId: String;
    const aSeq: Integer): Boolean;
  begin
    { INV008 }
    aDataSet.SQL.Text := BuildDetailApplySQL( dmInsert );
    aDataSet.Parameters.ParamByName( 'IDENTIFYID1' ).Value := IDENTIFYID1;
    aDataSet.Parameters.ParamByName( 'IDENTIFYID2' ).Value := IDENTIFYID2;
    aDataSet.Parameters.ParamByName( 'INVID' ).Value := aInvId;
    aDataSet.Parameters.ParamByName( 'BILLID' ).Value :=
      cdsDetail.FieldByName( 'BILLID' ).AsString;
    aDataSet.Parameters.ParamByName( 'BILLIDITEMNO' ).Value :=
      cdsDetail.FieldByName( 'BILLIDITEMNO' ).AsString;
    aDataSet.Parameters.ParamByName( 'SEQ' ).Value := aSeq;
    if ( cdsDetail.FieldByName( 'STARTDATE' ).AsString <> EmptyStr ) then
    begin
      aDataSet.Parameters.ParamByName( 'STARTDATE' ).Value :=
        cdsDetail.FieldByName( 'STARTDATE' ).AsString;
    end
    else begin
      aDataSet.Parameters.ParamByName( 'STARTDATE' ).Value := EmptyStr;
    end;
    if (  cdsDetail.FieldByName( 'ENDDATE' ).AsString <> EmptyStr ) then
    begin
      aDataSet.Parameters.ParamByName( 'ENDDATE' ).Value :=
        cdsDetail.FieldByName( 'ENDDATE' ).AsString;
    end
    else begin
      aDataSet.Parameters.ParamByName( 'ENDDATE' ).Value := EmptyStr;
    end;
    aDataSet.Parameters.ParamByName( 'ITEMID' ).Value :=
      cdsDetail.FieldByName( 'ITEMID' ).AsString;
    aDataSet.Parameters.ParamByName( 'DESCRIPTION' ).Value :=
      cdsDetail.FieldByName( 'DESCRIPTION' ).AsString;
    aDataSet.Parameters.ParamByName( 'QUANTITY' ).Value :=
      cdsDetail.FieldByName( 'QUANTITY' ).AsInteger;
    aDataSet.Parameters.ParamByName( 'UNITPRICE' ).Value :=
      cdsDetail.FieldByName( 'UNITPRICE' ).AsFloat;
    aDataSet.Parameters.ParamByName( 'TAXAMOUNT' ).Value :=
      cdsDetail.FieldByName( 'TAXAMOUNT' ).AsFloat;
    aDataSet.Parameters.ParamByName( 'TOTALAMOUNT' ).Value :=
      cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsFloat;
    aDataSet.Parameters.ParamByName( 'CHARGEEN' ).Value :=
      cdsDetail.FieldByName( 'CHARGEEN' ).AsString;
    aDataSet.Parameters.ParamByName( 'UPTEN' ).Value := dtmMain.getLoginUserName;
    aDataSet.Parameters.ParamByName( 'SERVICETYPE' ).Value :=
      cdsDetail.FieldByName( 'SERVICETYPE' ).AsString;
    aDataSet.Parameters.ParamByName( 'SALEAMOUNT' ).Value :=
      cdsDetail.FieldByName( 'SALEAMOUNT' ).AsFloat;
    aDataSet.Parameters.ParamByName( 'LINKTOMIS' ).Value :=
      cdsDetail.FieldByName( 'LINKTOMIS' ).AsString;
    aDataSet.Parameters.ParamByName( 'FACISNO' ).Value :=
      cdsDetail.FieldByName( 'FACISNO' ).AsString;
    aDataSet.Parameters.ParamByName( 'ACCOUNTNO' ).Value :=
      cdsDetail.FieldByName( 'ACCOUNTNO' ).AsString;
    aDataSet.ExecSQL;
    if ( FInvCreateType in [icPrev, icAfter, icLocale] ) then
    begin
      if ( dtmMain.GetLinkToMIS ) and
         ( cdsDetail.FieldByName( 'LINKTOMIS' ).AsString = 'Y' ) and
         ( not IsMergeRecord( aInvId, IntToStr( aSeq ) ) ) then
      begin
        { ��{�����Ȥ᦬�O����, ��s�^ SO ���� Table }
        UpdateInvoiceToSO( aInvId,
          cdsDetail.FieldByName( 'BILLID' ).AsString,
          cdsDetail.FieldByName( 'BILLIDITEMNO' ).AsString );
      end;
    end;
  end;

  { -------------------------------------------------------- }

  procedure DeleteAllInv008A(const aId: String);
  begin
    aDataSet.Close;
    aDataSet.SQL.Clear;
    aDataSet.SQL.Text := Format(
      ' DELETE FROM INV008A WHERE INVID = ''%s'' ', [aId] );
    aDataSet.ExecSQL;
  end;

  { -------------------------------------------------------- }

  function ApplyInv008AData: Boolean;
  var
    aLinkToMis: Boolean;
    aPreSeq: String;
  begin
    cdsInv008A.First;
    aPreSeq := EmptyStr;
    {}
    aDataSet.SQL.Text := BuildInv008AApplySQL( dmInsert );
    {}
    while not cdsInv008A.Eof do
    begin
      {}
      aDataSet.Parameters.ParamByName( 'ITEMIDREF' ).Value :=
        cdsInv008A.FieldByName( 'ITEMIDREF' ).AsString;
      {}
      aDataSet.Parameters.ParamByName( 'INVID' ).Value :=
        cdsInv008A.FieldByName( 'INVID' ).AsString;
      {}
      aDataSet.Parameters.ParamByName( 'BILLID' ).Value :=
        cdsInv008A.FieldByName( 'BILLID' ).AsString;
      {}
      aDataSet.Parameters.ParamByName( 'BILLIDITEMNO' ).Value :=
        cdsInv008A.FieldByName( 'BILLIDITEMNO' ).AsString;
      {}
      aDataSet.Parameters.ParamByName( 'SEQ' ).Value :=
        cdsInv008A.FieldByName( 'SEQ' ).AsString;
      {};
      if ( cdsInv008A.FieldByName( 'STARTDATE' ).AsString <> EmptyStr ) then
      begin
        aDataSet.Parameters.ParamByName( 'STARTDATE' ).Value :=
          cdsInv008A.FieldByName( 'STARTDATE' ).AsString;
      end
      else begin
        aDataSet.Parameters.ParamByName( 'STARTDATE' ).Value := EmptyStr;
      end;
      {}
      if ( cdsInv008A.FieldByName( 'ENDDATE' ).AsString <> EmptyStr ) then
      begin
        aDataSet.Parameters.ParamByName( 'ENDDATE' ).Value :=
          cdsInv008A.FieldByName( 'ENDDATE' ).AsString;
      end
      else begin
        aDataSet.Parameters.ParamByName( 'ENDDATE' ).Value := EmptyStr;
      end;
      {}
      aDataSet.Parameters.ParamByName( 'ITEMID' ).Value :=
        cdsInv008A.FieldByName( 'ITEMID' ).AsString;
      {}
      aDataSet.Parameters.ParamByName( 'DESCRIPTION' ).Value :=
        cdsInv008A.FieldByName( 'DESCRIPTION' ).AsString;
      {}
      aDataSet.Parameters.ParamByName( 'QUANTITY' ).Value :=
        cdsInv008A.FieldByName( 'QUANTITY' ).AsInteger;
      {}
      aDataSet.Parameters.ParamByName( 'UNITPRICE' ).Value :=
        cdsInv008A.FieldByName( 'UNITPRICE' ).AsFloat;
      {}
      aDataSet.Parameters.ParamByName( 'TAXAMOUNT' ).Value :=
        cdsInv008A.FieldByName( 'TAXAMOUNT' ).AsFloat;
      {}
      aDataSet.Parameters.ParamByName( 'TOTALAMOUNT' ).Value :=
        cdsInv008A.FieldByName( 'TOTALAMOUNT' ).AsFloat;
      {}
      aDataSet.Parameters.ParamByName( 'SERVICETYPE' ).Value :=
        cdsInv008A.FieldByName( 'SERVICETYPE' ).AsString;
      {}
      aDataSet.Parameters.ParamByName( 'SALEAMOUNT' ).Value :=
        cdsInv008A.FieldByName( 'SALEAMOUNT' ).AsFloat;
      {}
      aDataSet.Parameters.ParamByName( 'FACISNO' ).Value :=
        cdsInv008A.FieldByName( 'FACISNO' ).AsString;
      {}
      aDataSet.Parameters.ParamByName( 'ACCOUNTNO' ).Value :=
        cdsInv008A.FieldByName( 'ACCOUNTNO' ).AsString;
      aDataSet.ExecSQL;
      { ��s�ȪA }
      if ( Nvl( aPreSeq, 'X' ) <> cdsInv008A.FieldByName( 'SEQ' ).AsString ) then
      begin
        aPreSeq := cdsInv008A.FieldByName( 'SEQ' ).AsString;
        aLinkToMis := IsLinkToMis(
          cdsInv008A.FieldByName( 'INVID' ).AsString,
          cdsInv008A.FieldByName( 'SEQ' ).AsString );
      end;
      if ( FInvCreateType in [icPrev, icAfter, icLocale] ) and
         ( dtmMain.GetLinkToMIS ) and ( aLinkToMis ) then
      begin
        UpdateInvoiceToSO(
          cdsInv008A.FieldByName( 'INVID' ).AsString,
          cdsInv008A.FieldByName( 'BILLID' ).AsString,
          cdsInv008A.FieldByName( 'BILLIDITEMNO' ).AsString );
      end;
      cdsInv008A.Next;
    end;
  end;

  { -------------------------------------------------------- }

  function ApplyInv099Data(const aAddCount: Integer): Boolean;
  var
    aPrefix: String;
  begin
    Result := False;
    aDataSet.Close;
    aDataSet.SQL.Text := Format(
      '  UPDATE INV099              ' +
      '     SET CURNUM = LPAD( TO_CHAR( TO_NUMBER( CURNUM ) + %d ), 8, ''0'' ), ' +
      '         LASTINVDATE = TO_DATE( ''%s'', ''YYYYMMDD''  )                  ' +
      '   WHERE IDENTIFYID1 = ''%s''  ' +
      '     AND IDENTIFYID2 = %s      ' +
      '     AND COMPID = ''%s''       ' +
      '     AND YEARMONTH = ''%s''    ' +
      '     AND PREFIX = ''%s''       ' +
      '     AND STARTNUM = ''%s''     ',
      [aAddCount,
       FormatDateTime( 'yyyymmdd', cdsMaster.FieldByName( 'INVDATE' ).AsDateTime ),
       IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, FChooseInv.YearMonth,
       FChooseInv.Prefix, FChooseInv.StartNum ] );
    aDataSet.ExecSQL;
    { �O�_���n�� useful �]�� "N" }
    aDataSet.SQL.Text := Format(
      ' UPDATE INV099                ' +
      '    SET USEFUL = ''N''        ' +
      '  WHERE IDENTIFYID1 = ''%s''  ' +
      '    AND IDENTIFYID2 = %s      ' +
      '    AND COMPID = ''%s''       ' +
      '    AND YEARMONTH = ''%s''    ' +
      '    AND PREFIX = ''%s''       ' +
      '    AND STARTNUM = ''%s''     ' +
      '    AND CURNUM > ENDNUM       ',
      [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, FChooseInv.YearMonth,
       FChooseInv.Prefix, FChooseInv.StartNum ] );
    aDataSet.ExecSQL;   
    Result := True;
  end;

  { -------------------------------------------------------- }

  function UpdateAutoCreateInv(const aObj: PAutoCreate): Boolean;
  begin
    Result := False;
    aDataSet.Close;
    aDataSet.SQL.Text := Format(
      ' UPDATE INV007 SET                            ' +
      '  INVDATE = TO_DATE( ''%s'', ''YYYY/MM/DD'' ),' +
      '  MAININVID = ''%s'',                         ' +
      '  MAINSALEAMOUNT = ''%f'',                    ' +
      '  MAINTAXAMOUNT = ''%f'',                     ' +
      '  MAININVAMOUNT = ''%f'',                     ' +
      '  SALEAMOUNT = ''%f'',                        ' +
      '  TAXAMOUNT = ''%f'',                         ' +
      '  INVAMOUNT = ''%f'',                         ' +
      '  CHARGEDATE = TO_DATE( ''%s'', ''YYYY/MM/DD'' ),' +
      '  CUSTID = ''%s'',                            ' +
      '  CUSTSNAME = ''%s'',                         ' +
      '  INVTITLE = ''%s'',                          ' +
      '  ZIPCODE = ''%s'',                           ' +
      '  INVADDR = ''%s'',                           ' +
      '  MAILADDR = ''%s'',                          ' +
      '  INSTADDR = ''%s'',                          ' +
      '  BUSINESSID = ''%s'',                        ' +
      '  MEMO1 = ''%s'',                             ' +
      '  MEMO2 = ''%s''                              ' +
      ' WHERE IDENTIFYID1 = ''%s''                   ' +
      '   AND IDENTIFYID2 = ''%s''                   ' +
      '   AND INVID = ''%s''                         ' +
      '   AND COMPID = ''%s''                        ',
      [aObj.InvDate,
       aObj.MainInvId, aObj.MainSaleAmount, aObj.MainTaxAmount, aObj.MainInvAmount,
       aObj.SaleAmount, aObj.TaxAmount, aObj.InvAmount, aObj.ChargeDate,
       aObj.CustId, aObj.CustName, aObj.InvTitle, aObj.ZipCode, aObj.InvAddr,
       aObj.MailAddr, aObj.InstAddr, aObj.BusinessId, aObj.Memo1, aObj.Memo2,
       IDENTIFYID1, IDENTIFYID2, aObj.InvId, dtmMain.getCompID] );
    aDataSet.ExecSQL;
    Result := True;
  end;

  { -------------------------------------------------------- }

  function ApplyInsertData: Integer;
  var
    aIndex, aDetailSeq: Integer;
    aFirstTimeApplyMaster: Boolean;
    aMainInvId, aPreInvId, aMsg: String;
    aInvList: TList;
    aAutoCreate: PAutoCreate;
    aFaciMemo : string;
  begin
    Result := 0;
    aFaciMemo := EmptyStr;
    aInvList := TList.Create;
    try
      cdsDetail.DisableControls;
      try
        DisableDetailDataSetEvent;
        try
          aDetailSeq := 0;
          aMainInvId := EmptyStr;
          aFirstTimeApplyMaster := True;
          cdsDetail.First;
          while not cdsDetail.Eof do
          begin
            // ���p���]�w�۰ʴ��}�o��, �B�w�F��۰ʴ��}����, ���m�Ǹ�
            if ( dtmMain.GetAutoCreateNum > 0 ) and
               ( aDetailSeq >= dtmMain.GetAutoCreateNum ) then
            begin
              aDetailSeq := 0;
            end;
            //%5358 �p�G��ܳ]��,�n�N�]�ƧǸ��g�JMemo1 By Kin 2009/12/14
            {--------------------------------------------------------------}
            if (dtmMain.GetShowFaci = 1) then
            begin
              if cdsDetail.FieldByName( 'FaciSNo' ).AsString <>'' then
              begin
                  if aFaciMemo <> EmptyStr then
                  begin
                    if pos(cdsDetail.FieldByName('FaciSNo').AsString,aFaciMemo)=0 then
                    begin
                      aFaciMemo := aFaciMemo +','+
                        cdsDetail.FieldByName('FaciSNo').AsString;
                    end;
                  end else
                  begin
                    aFaciMemo := cdsDetail.FieldByName( 'FaciSNo' ).AsString;
                  end;
              end;
            end;
            {--------------------------------------------------------------}

            Inc( aDetailSeq );
            if ( aDetailSeq = 1 ) then
            begin
              New( aAutoCreate );
              ZeroMemory( aAutoCreate, SizeOf( aAutoCreate^ ) );
              aInvList.Add( aAutoCreate );
              if ( aFirstTimeApplyMaster ) then
              begin
                { �Ĥ@�i�o�����X }
                aAutoCreate.InvId := FInvoiceId;
              end else
              begin
                // �۰ʴ��}�o�����X, ���s���o�����X
                aAutoCreate.InvId := GetNextInvNumber( aPreInvId );
              end;
              {}
              aPreInvId := aAutoCreate.InvId;
              { �s�W�w�]���a�o�����X, �̫�A�γ̫�@�i�o�����X��s }
              aAutoCreate.MainInvId := aAutoCreate.InvId;
              { �Ҧ������o�������B, �D�o�����B }
              aAutoCreate.MainSaleAmount :=
                cdsMaster.FieldByName( 'SALEAMOUNT' ).AsFloat;
              aAutoCreate.MainTaxAmount :=
                cdsMaster.FieldByName( 'TAXAMOUNT' ).AsFloat;
              aAutoCreate.MainInvAmount :=
                cdsMaster.FieldByName( 'INVAMOUNT' ).AsFloat;
              {}
              { �ӱi�o���B���m }
              aAutoCreate.SaleAmount := 0;
              aAutoCreate.TaxAmount := 0;
              aAutoCreate.InvAmount := 0;
              {}
              aFirstTimeApplyMaster := False;
              ApplyMasterDataSet( dmInsert, aAutoCreate );
            end;
            { �P�B INV008A �����, INVID �� SEQ }
            { �w�����n, �u���X�֪��~�|�g�J INV008A, �����{�����|���X�ְʧ@
            if not SyncInv008ADataSet( dmUpdate, cdsDetail.FieldByName( 'SEQ' ).AsString,
              IntToStr( aDetailSeq ), aAutoCreate.InvId, aMsg ) then
            begin
              raise Exception.Create( aMsg );
            end; }
            { �⥻�� Detail �� INVID, �� SEQ �]��s, �d�U���i����,
              ���M�᭱ UpdateInvoiceToSO �|�������� }
            cdsDetail.Edit;
            cdsDetail.FieldByName( 'INVID' ).AsString := aAutoCreate.InvId;
            cdsDetail.FieldByName( 'SEQ' ).AsInteger := aDetailSeq;
            cdsDetail.Post;
            { INV008 }
            ApplyDetailDataSetByRecord(
              cdsDetail.FieldByName( 'INVID' ).AsString,
              cdsDetail.FieldByName( 'SEQ' ).AsInteger );
            { �֭p���Ӫ��B, �����ӱi�o�������B }
            aAutoCreate.SaleAmount := ( aAutoCreate.SaleAmount +
              cdsDetail.FieldByName( 'SALEAMOUNT' ).AsFloat );
            aAutoCreate.TaxAmount := ( aAutoCreate.TaxAmount +
              cdsDetail.FieldByName( 'TAXAMOUNT' ).AsFloat );
            aAutoCreate.InvAmount := ( aAutoCreate.InvAmount +
              cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsFloat );
            cdsDetail.Next;
            { �w�O�̫�@�� Detail, �O���̫�@�i���o�����X }
            if ( cdsDetail.Eof ) then aMainInvId := aAutoCreate.InvId;
          end;
          { �^�Y��s�D�� }
          for aIndex := 0 to aInvList.Count - 1 do
          begin
            PAutoCreate( aInvList[aIndex] ).MainInvId := aMainInvId;
            {--------------------------------------------------}
            //#5358 �W�[ShowFaci �P�_�b�}�ߵo���ɬO�_�n�N�]�ƧǸ��g�JMemo1 By Kin 2009/12/14
            if aFaciMemo <> EmptyStr then
            begin
              PAutoCreate( aInvList[aIndex] ).Memo1 :=
                cdsMaster.FieldByName( 'Memo1' ).AsString + '�]��:' + aFaciMemo;
            end;
            {--------------------------------------------------}
            UpdateAutoCreateInv( PAutoCreate( aInvList[aIndex] ) );
          end;
          Result := aInvList.Count;
        finally
          EnableDetailDataSetEvent;
        end;
      finally
        cdsDetail.EnableControls;
      end;
    finally
      for aIndex := 0 to aInvList.Count - 1 do
        Dispose( PAutoCreate( aInvList[aIndex] ) );
      aInvList.Free;
    end;
  end;

  { -------------------------------------------------------- }

  function GetMainInvList(const aInvId, aMainInvId: String): TList;
  var
    aIndex: Integer;
    aObj: PAutoCreate;
  begin
    Result := TList.Create;
    aDataSet.Close;
    aDataSet.SQL.Text := Format(
      '  SELECT /*+ RULE */                                      ' +
      '         A.INVID, A.SALEAMOUNT, A.TAXAMOUNT, A.INVAMOUNT  ' +
      '    FROM INV007 A                                         ' +
      '   WHERE A.MAININVID = ''%s''                             ' +
      '     AND A.IDENTIFYID1 = ''%s''                           ' +
      '     AND A.IDENTIFYID2 = ''%s''                           ' +
      '     AND A.COMPID = ''%s''                               ',
      [aMainInvId, IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID] );
    aDataSet.Open;
    aDataSet.First;
    while not aDataSet.Eof do
    begin
      New( aObj );
      ZeroMemory( aObj, SizeOf( aObj^ ) );
      aObj.InvId := aDataSet.FieldByName( 'INVID' ).AsString;
      aObj.MainInvId := aMainInvId;
      aObj.MainSaleAmount := 0;
      aObj.MainTaxAmount := 0;
      aObj.MainInvAmount := 0;
      if ( aObj.InvId = aInvId ) then
      begin
        aObj.SaleAmount := 0;
        aObj.TaxAmount := 0;
        aObj.InvAmount := 0;
      end else
      begin
        aObj.SaleAmount := aDataSet.FieldByName( 'SALEAMOUNT' ).AsFloat;
        aObj.TaxAmount := aDataSet.FieldByName( 'TAXAMOUNT' ).AsFloat;
        aObj.InvAmount := aDataSet.FieldByName( 'INVAMOUNT' ).AsFloat;
      end;
      Result.Add( aObj );
      aDataSet.Next;
    end;
    aDataSet.Close;
  end;

  { -------------------------------------------------------- }

  function GetObject(const aInvId: String; aList: TList): PAutoCreate;
  var
    aIndex: Integer;
  begin
    Result := nil;
    for aIndex := 0 to aList.Count - 1 do
    begin
      if ( PAutoCreate( aList[aIndex] ).InvId <> aInvId ) then Continue;
      Result := PAutoCreate( aList[aIndex] );
      Break;
    end;
  end;

  { -------------------------------------------------------- }

  procedure CloneObjData(aSource, aDest: PAutoCreate);
  begin
    aDest.InvDate := aSource.InvDate;
    aDest.ChargeDate := aSource.ChargeDate;
    aDest.CustId := aSource.CustId;
    aDest.CustName := aSource.CustName;
    aDest.InvTitle := aSource.InvTitle;
    aDest.ZipCode := aSource.ZipCode;
    aDest.InvAddr := aSource.InvAddr;
    aDest.MailAddr := aSource.MailAddr;
    aDest.InstAddr := aSource.InstAddr;
    aDest.BusinessId := aSource.BusinessId;
    aDest.Memo1 := aSource.Memo1;
    aDest.Memo2 := aSource.Memo2;
  end;

  { -------------------------------------------------------- }

  function ApplyUpdateData: Integer;
  var
    aList: TList;
    aIndex: Integer;
    aObj: PAutoCreate;
    aSale, aTax, aInv: Double;
    aOldSeq: String;
  begin
    Result := 0;
    { Detail ��Ƥ@�ߧR����A�ηs�W�覡�B�z }
    DeleteAllInvoiceDetail( FInvoiceId );
    if ( FInvCreateType in [icPrev, icAfter, icLocale] ) then
    begin
      if dtmMain.GetLinkToMIS then
      begin
        { �`�N, �o�ؤ��� INV008A �ҰO�������Ӭ���, ������ INVID ��
          �^�Y�h��ȪA�t�ά��������O��ƥ��M��, �U���A�� INV008A
          �ҰO������ƦA�^�Y�h��s�ȪA�t�� }
        { ����@�� SO Table ���o����� }
        CloneInvoiceFromSO;
        { �N�ȪA�t�Ϊ��o����ƲM�� }
        EmptyInvoiceToSO;
      end;
    end;
    aList := GetMainInvList( FInvoiceId, cdsMaster.FieldByName( 'MAININVID' ).AsString );
    try
      cdsDetail.DisableControls;
      try
        { ���N�Ҧ� INV008 �g�� Table }
        cdsDetail.First;
        while not cdsDetail.Eof do
        begin
          ApplyDetailDataSetByRecord( FInvoiceId,
            cdsDetail.FieldByName( 'SEQ' ).AsInteger );
          cdsDetail.Next;
        end;
        cdsDetail.First;
        {}
        aObj := GetObject( FInvoiceId, aList );
        aObj.SaleAmount := cdsMaster.FieldByName( 'SALEAMOUNT' ).AsFloat;
        aObj.TaxAmount := cdsMaster.FieldByName( 'TAXAMOUNT' ).AsFloat;
        aObj.InvAmount := cdsMaster.FieldByName( 'INVAMOUNT' ).AsFloat;
        {}
        ApplyMasterDataSet( dmUpdate, aObj );
        { �⦹���D�ɸ�� Copy ��䥦�� }
        { ���K�N���B�[�_�� }
        aSale := 0;
        aTax := 0;
        aInv := 0;
        for aIndex := 0 to aList.Count - 1 do
        begin
          if ( PAutoCreate( aList[aIndex] ).InvId <> aObj.InvId ) then
            CloneObjData( aObj, PAutoCreate( aList[aIndex] ) );
          aSale := ( aSale + PAutoCreate( aList[aIndex] ).SaleAmount );
          aTax := ( aTax + PAutoCreate( aList[aIndex] ).TaxAmount );
          aInv := ( aInv + PAutoCreate( aList[aIndex] ).InvAmount );
        end;
        { ��s�Ҧ������o�� }
        for aIndex := 0 to aList.Count - 1 do
        begin
          aObj := PAutoCreate( aList[aIndex] );
          aObj.MainSaleAmount := aSale;
          aObj.MainTaxAmount := aTax;
          aObj.MainInvAmount := aInv;
          UpdateAutoCreateInv( aObj );
        end;
        {}
      finally
        cdsDetail.EnableControls;
      end;
    finally
      for aIndex := 0 to aList.Count - 1 do
        Dispose( PAutoCreate( aList[aIndex] ) );
      aList.Free;
    end;
    Result := 1;
  end;  

  { -------------------------------------------------------- }


var
  aUseInvCount: Integer;
begin
  Result := False;
  aDataSet := dtmMain.adoComm;
  aSoDataSet := dtmSO.adoComm;
  aConnection := dtmMain.InvConnection;
  aSoConnection := dtmSO.SOConnection;
  // �۰ʴ��}�o����
  aConnection.BeginTrans;
  if dtmMain.GetLinkToMIS then aSoConnection.BeginTrans;
  try
    if ( FMode in [dmInsert] ) then
    begin
      aUseInvCount := ApplyInsertData;
      ApplyInv099Data( aUseInvCount );
      //ApplyInv008AData; �����n
    end else
    begin
      ApplyUpdateData;
      DeleteAllInv008A( FInvoiceId );
      ApplyInv008AData;
    end;
    aConnection.CommitTrans;
    if dtmMain.GetLinkToMIS then aSoConnection.CommitTrans;
    Result := True;
  except
    on E: Exception do
    begin
      TranslateError( E.Message );
      { PKey ����, PM �n�D��q�T�� }
      if OracleErros.ErrorCode = 1 then
      begin
        aErrMsg := '�ӵo�����X�w�Q�}��,�Ц�<�o�����>�B,���s����r�y,�Y�i�}��!';
        if ( txtINVDATE.CanFocus ) then txtINVDATE.SetFocus; 
      end
      else begin
        aErrMsg := E.Message;
      end;
      if dtmMain.GetLinkToMIS then aSoConnection.RollbackTrans;
      aConnection.RollbackTrans;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA02_4.ValidateCustomer(const aCustId: String;
  aInvCreateType: TInvCreateType): Boolean;
var
  aDataSet: TADOQuery;
begin
  Result := False;
  // �@��}��, �Ȥ��Ʊq�o���t�� INV002 �� INV019 �ˬd
  if ( aInvCreateType in [icNormal] ) then
  begin
    aDataSet := dtmMain.adoComm;
    aDataSet.Close;
    aDataSet.SQL.Text :=  Format(
      '  SELECT COUNT(1) AS COUNTS              ' +
      '    FROM INV019 A, INV002 B              ' +
      '   WHERE A.IDENTIFYID1 = B.IDENTIFYID1   ' +
      '     AND A.IDENTIFYID2 = B.IDENTIFYID2   ' +
      '     AND A.COMPID = B.COMPID             ' +
      '     AND A.CUSTID = B.CUSTID             ' +
      '     AND A.CUSTID = ''%s''               ' +
      '     AND A.IDENTIFYID1 = ''%s''          ' +
      '     AND A.IDENTIFYID2 = %s              ',
      [ aCustId, IDENTIFYID1, IDENTIFYID2 ] );
    aDataSet.Open;
    Result := ( aDataSet.FieldByName( 'COUNTS' ).AsInteger > 0 );
    aDataSet.Close;
  end else
  // �{���}��, �w�}, ��}, �Ȥ��Ʊq SO001 �ˬd, �Y���P�ȪA�t�Τ���, �h���ˬd
  if ( FInvCreateType in [icLocale, icPrev, icAfter] ) then
  begin
    if ( not dtmMain.GetLinkToMIS ) then
      Result := True
    else begin
      aDataSet := dtmSO.adoComm;
      aDataSet.Close;
      aDataSet.SQL.Text := Format(
        ' SELECT COUNT( 1 ) AS COUNTS FROM %s.SO001 A   ' +
        '  WHERE A.CUSTID = ''%s''                      ',
        [ dtmMain.getMisDbOwner, aCustId ] );
      try
        aDataSet.Open;
        Result := ( aDataSet.FieldByName( 'COUNTS' ).AsInteger > 0 );
      except
        { ... }
      end;
      aDataSet.Close;
    end;  
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA02_4.ValidateBillingNo(const aBillId, aItemId: String;
  var aErrMsg: String; const aIsSaved: Boolean = False): Boolean;

  { ----------------------------------------------------- }

  function GetValidateSQL(const aSource: String): String;
  begin
    Result := Format(
      '  SELECT B.TAXCODE, C.RATE1                    ' +
      '    FROM %s.%s A, %s.CD019 B, %s.CD033 C    ' +
      '   WHERE A.CITEMCODE = B.CODENO                ' +
      '     AND B.TAXCODE = C.CODENO                  ' +
      '     AND C.TAXFLAG = 1                         ' +
      '     AND A.CANCELFLAG = 0                      ' +
      '     AND ( A.SHOULDAMT + A.REALAMT <> 0 )      ' +
      '     AND A.GUINO IS NULL                       ' +
      '     AND A.INVOICETIME IS NULL                 ' +
      '     AND A.CUSTID = ''%s''                     ' +
      '     AND A.SERVICETYPE IN ( ' + dtmMain.GetServiceTypeStrEx + ' ) ' +
      '     AND A.BILLNO = ''%s''                     ',
      [ dtmMain.getMisDbOwner, aSource, dtmMain.getMisDbOwner, dtmMain.getMisDbOwner,
        cdsMaster.FieldByName( 'CUSTID' ).AsString, aBillId ] );
    if aItemId <> '' then
      Result := Result + '   AND A.ITEM = ' + QuotedStr( aItemId );
  end;

  { ----------------------------------------------------- }

  function GetSavedValidateSQL(const aSource: String): String;
  begin
    Result := Format(
      '  SELECT B.TAXCODE, C.RATE1                    ' +
      '    FROM %s.%s A, %s.CD019 B, %s.CD033 C    ' +
      '   WHERE A.CITEMCODE = B.CODENO                ' +
      '     AND B.TAXCODE = C.CODENO                  ' +
      '     AND C.TAXFLAG = 1                         ' +
      '     AND A.CANCELFLAG = 0                      ' +
      '     AND ( A.SHOULDAMT + A.REALAMT <> 0 )      ' +
      '     AND A.CUSTID = ''%s''                     ' +
      '     AND A.SERVICETYPE IN ( ' + dtmMain.GetServiceTypeStrEx + ' ) ' +
      '     AND A.BILLNO = ''%s''                     ',
      [ dtmMain.getMisDbOwner, aSource, dtmMain.getMisDbOwner, dtmMain.getMisDbOwner,
        cdsMaster.FieldByName( 'CUSTID' ).AsString, aBillId ] );
    if aItemId <> '' then
      Result := Result + '   AND A.ITEM = ' + QuotedStr( aItemId );
  end;

  { ----------------------------------------------------- }

var
  aSQL, aTaxCode, aTaxRate: String;
  aDataSet: TADOQuery;
  aIndex: Integer;
  aValidateSO033, aValidateSO034: Boolean;
begin
  Result := True;
  aErrMsg := '';
  if ( FInvCreateType in [icLocale, icPrev, icAfter] ) then
  begin
    aDataSet := dtmSO.adoComm;
    aValidateSO033 := False;
    aValidateSO034 := False;
    for aIndex := 1 to 2 do
    begin
      case aIndex of
        1: begin
             if aIsSaved then
               aSQL := GetSavedValidateSQL( 'SO033' )
             else
               aSQL := GetValidateSQL( 'SO033' );
           end;
        2: begin
            if aIsSaved then
              aSQL := GetSavedValidateSQL( 'SO034' )
            else
              aSQL := GetValidateSQL( 'SO034' );
           end;
      end;
      aDataSet.Close;
      aDataSet.SQL.Text := aSQL;
      aDataSet.Open;
      case aIndex of
        1: aValidateSO033 := ( not aDataSet.IsEmpty );
        2: aValidateSO034 := ( not aDataSet.IsEmpty );
      end;
      { �Ĥ@��, �ˬd SO033, �ˬd����h���L�j��, ���� check SO034 }
      if ( aIndex = 1 ) and ( not aValidateSO033 ) then
      begin
        Continue;
      end
      { �Ĥ@��, �ˬd SO033, �ˬd��h�ˬd�|�O�ε|�v }
      else if ( aIndex = 1 ) and ( aValidateSO033 )then
      begin
        if ( aItemId <> '' ) then
        begin
          aTaxCode := aDataSet.FieldByName( 'TAXCODE' ).AsString;
          aTaxRate := aDataSet.FieldByName( 'RATE1' ).AsString;
          if ( aTaxCode <> GetMasterTaxCode ) or
             ( aTaxRate <> cdsMaster.FieldByName( 'TAXRATE' ).AsString ) then
          begin
            aErrMsg := '�������O���ت��|�O/�|�v�P�o�����P�C';
            aValidateSO033 := False;
            Break;
          end;
        end else
        begin
          Break;
        end;
      end
      { �ĤG��, �ˬd SO034, ���M�ˬd����, �h�L���渹 }
      else if ( aIndex = 2 ) and ( not aValidateSO034 ) then
      begin
        aErrMsg := '�L�����O�渹�C';
        Break;
      end
      { �ĤG��, �ˬd SO034, �ˬd��h�ˬd�|�O�ε|�v }
      else if ( aIndex = 2 ) and ( aValidateSO034 ) then
      begin
        if ( aItemId <> '' ) then
        begin
          aTaxCode := aDataSet.FieldByName( 'TAXCODE' ).AsString;
          aTaxRate := aDataSet.FieldByName( 'RATE1' ).AsString;
          if ( aTaxCode <> GetMasterTaxCode ) or
             ( aTaxRate <> cdsMaster.FieldByName( 'TAXRATE' ).AsString ) then
          begin
            aErrMsg := '�������O�渹���|�O/�|�v�P�o�����šC';
            aValidateSO034 := False;
            Break;
          end;
        end;
      end;
    end;
    aDataSet.Close;
    Result := ( aValidateSO033 or aValidateSO034 );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA02_4.ValidateChargeItemId(const aItemId: String;
  var aErrMsg: String): Boolean;
var
  aDataSet: TADOQuery;
begin
  Result := True;
  aErrMsg := '';
  aDataSet := dtmMain.adoComm;
  aDataSet.Close;
  aDataSet.SQL.Text := Format(
    '  SELECT A.ITEMID                   ' +
    '    FROM INV005 A                   ' +
    '   WHERE A.IDENTIFYID1 = ''%s''     ' +
    '     AND A.IDENTIFYID2 = %s         ' +
    '     AND A.COMPID = ''%s''          ' +
    '     AND A.ITEMID = ''%s''          ' +
    '     AND A.TAXCODE = ''%s''         ',
    [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, aItemId, GetMasterTaxCode] );
  aDataSet.Open;
  Result := not aDataSet.IsEmpty;
  if not Result then
  begin
    if aItemId = EmptyStr then
      aErrMsg := '�п�J���O���ءC'
    else
      aErrMsg := '�L�����O����, ���ˬd�|�O�O�_���T�C';
  end;
  aDataSet.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.FillDetailColumnPropertiesValue;
var
  aComboBoxProperties: TcxComboBoxProperties;
  aService: String;
begin
  aComboBoxProperties := TcxComboBoxProperties(
    DetailGridView.GetColumnByFieldName( 'SERVICETYPE' ).Properties );
  aComboBoxProperties.Items.Clear;
  aService := dtmMain.GetServiceTypeStr;
  if ( Pos( ',', aService ) = 0 ) then
    aComboBoxProperties.Items.Add( aService )
  else begin
    while Pos( ',', aService ) > 0 do
    begin
      aComboBoxProperties.Items.Add( Copy( aService, 1, Pos( ',', aService ) - 1 ) );
      Delete( aService, 1, Pos( ',', aService ) );
    end;
    if ( aService <> EmptyStr ) then
      aComboBoxProperties.Items.Add( aService );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.SetMasterFieldState;
begin
  cdsMaster.FieldByName( 'ISOBSOLETEDESC' ).ReadOnly := False;
  cdsMaster.FieldByName( 'CANMODIFYDESC' ).ReadOnly := False;
  cdsMaster.FieldByName( 'INVFORMATDESC' ).ReadOnly := False;
  cdsMaster.FieldByName( 'HOWTOCREATEDESC' ).ReadOnly := False;
  cdsMaster.FieldByName( 'PRINTCOUNT' ).ReadOnly := False;
  TFloatField( cdsMaster.FieldByName( 'TAXAMOUNT' ) ).DisplayFormat := '#,##0';
  TFloatField( cdsMaster.FieldByName( 'SALEAMOUNT' ) ).DisplayFormat := '#,##0';
  TFloatField( cdsMaster.FieldByName( 'INVAMOUNT' ) ).DisplayFormat := '#,##0';
  TFloatField( cdsMaster.FieldByName( 'MAINTAXAMOUNT' ) ).DisplayFormat := '#,##0';
  TFloatField( cdsMaster.FieldByName( 'MAINSALEAMOUNT' ) ).DisplayFormat := '#,##0';
  TFloatField( cdsMaster.FieldByName( 'MAININVAMOUNT' ) ).DisplayFormat := '#,##0';
end;


{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.SetDetailFieldState;
begin
  cdsDetail.FieldByName( 'SEQ' ).Alignment := taCenter;
  cdsDetail.FieldByName( 'BILLIDITEMNO' ).Alignment := taCenter;
  cdsDetail.FieldByName( 'SERVICETYPE' ).Alignment := taCenter;
  cdsDetail.FieldByName( 'TAXNAME' ).Alignment := taCenter;
  cdsDetail.FieldByName( 'SIGN' ).Alignment := taCenter;
  cdsDetail.FieldByName( 'CHARGEEN' ).Alignment := taCenter;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.MasterDataToEditor;
begin
  rdoTAXRATE1.OnClick := nil;
  rdoTAXRATE2.OnClick := nil;
  rdoTAXRATE3.OnClick := nil;
  try
    if cdsMaster.IsEmpty then
    begin
      rdoTAXRATE1.Checked := True;
      txtInvUse.Clear;
    end
    else begin
      if cdsMaster.FieldByName( 'TAXTYPE' ).AsString = '2' then
        rdoTAXRATE2.Checked := True
      else if cdsMaster.FieldByName( 'TAXTYPE' ).AsString = '3' then
        rdoTAXRATE3.Checked := True
      else
        rdoTAXRATE1.Checked := True;
      txtInvUse.Value := cdsMaster.FieldByName( 'INVUSEID' ).AsString;
      txtInvUse.ValueName := cdsMaster.FieldByName( 'INVUSEDESC' ).AsString;
    end;
  finally
    rdoTAXRATE1.OnClick := rdoTAXRATE1Click;
    rdoTAXRATE2.OnClick := rdoTAXRATE1Click;
    rdoTAXRATE3.OnClick := rdoTAXRATE1Click;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.UpdatePriceToMaster;
var
  aBookMark: TBookMark;
  aSale, aTax, aTotal: Integer;
begin
  aSale := 0;
  aTax := 0;
  aTotal := 0;
  aBookMark := cdsDetail.GetBookmark;
  try
    cdsDetail.DisableControls;
    try
      cdsDetail.First;
      while not cdsDetail.Eof do
      begin
        aTax := aTax + cdsDetail.FieldByName( 'TAXAMOUNT' ).AsInteger;
        aSale := aSale + cdsDetail.FieldByName( 'SALEAMOUNT' ).AsInteger;
        aTotal := aTotal + cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsInteger;
        cdsDetail.Next;
      end;
    finally
      cdsDetail.EnableControls;
    end;
  finally
    cdsDetail.FreeBookmark( aBookMark );
  end;
  cdsMaster.FieldByName( 'TAXAMOUNT' ).AsInteger := aTax;
  cdsMaster.FieldByName( 'SALEAMOUNT' ).AsInteger := aSale;
  cdsMaster.FieldByName( 'INVAMOUNT' ).AsInteger := aTotal;
end;

{ ---------------------------------------------------------------------------- }

{ Only Call This Method after run CopyChargeDataToDataSet procedure }
procedure TfrmInvA02_4.UpdateTaxToMaster;
begin
  { �u���P�_�����Y�i }
  cdsMaster.FieldByName( 'TAXTYPE' ).AsString :=
    cdsDetail.FieldByName( 'TAXCODE' ).AsString;
  if cdsMaster.FieldByName( 'TAXTYPE' ).AsString <> '1' then
    cdsMaster.FieldByName( 'TAXRATE' ).AsString := '0';
  MasterDataToEditor;
  { ���s�]�w�@�� TaxRate , �]���b MasterDataToEditor ���w�� rdoTAXRATE ��
    radio button onclick ����, �ҥH���|Ĳ�o }
  if ( cdsMaster.FieldByName( 'TAXTYPE' ).AsString = '1' ) then
  begin
    if cdsMaster.FieldByName( 'TAXRATE' ).AsInteger = 0 then
      cdsMaster.FieldByName( 'TAXRATE' ).AsInteger := 5;
    if not txtTAXRATE.Enabled then txtTAXRATE.Enabled := True;
  end
  else begin
    if cdsMaster.FieldByName( 'TAXRATE' ).AsInteger <> 0 then
      cdsMaster.FieldByName( 'TAXRATE' ).AsInteger := 0;
    if txtTAXRATE.Enabled then txtTAXRATE.Enabled := False;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.UpdateTaxToDetail;
var
  aTaxCode: String;
  aTaxRate: Double;
  aFinish: Boolean;
begin
  if ( cdsDetail.IsEmpty ) or ( cdsDetail.RecordCount <= 0 ) then Exit;
    cdsDetail.DisableControls;
  try
    cdsDetail.CheckBrowseMode;
    { �ˮֵo���D�ɪ��|�O�O�_�P���Ӥ��P, ���P�Y�R�� }
    cdsDetail.First;
    aTaxCode := GetMasterTaxCode;
    aTaxRate := cdsMaster.FieldByName( 'TAXRATE' ).AsFloat;
    aFinish := False;
    while not aFinish do
    begin
      if cdsDetail.FieldByName( 'TAXCODE' ).AsString <> EmptyStr then
      begin
        if cdsDetail.FieldByName( 'TAXCODE' ).AsString <> aTaxCode then
          cdsDetail.Delete
        else
          cdsDetail.Next;
      end else
        cdsDetail.Next;
      aFinish := cdsDetail.Eof;
    end;
  finally
    cdsDetail.EnableControls;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.UpdateDeatilPrice;
var
  aEvent: TDataSetNotifyEvent;
begin
  aEvent := cdsDetail.AfterPost;
  cdsDetail.AfterPost := nil;
  try
    cdsDetail.DisableControls;
    try
      cdsDetail.CheckBrowseMode;
      { ����Ҧ����B }
      cdsDetail.First;
      while not cdsDetail.Eof do
      begin
        cdsDetail.Edit;
        CalculateChargePrice( chTotalAmount );
        cdsDetail.Post;
        cdsDetail.Next;
      end;
    finally
      cdsDetail.EnableControls;
    end;
  finally
    cdsDetail.AfterPost := aEvent;
  end;
  if Assigned( @aEvent ) then aEvent( cdsDetail );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.UpdateDetailPriceByChargeSign;
var
  aChargeSign: String;
begin
  aChargeSign := cdsDetail.FieldByName( 'SIGN' ).AsString;
  { ������X(-)��, �۰��ର�t�� }
  if ( aChargeSign = '-' ) then
  begin
    if ( cdsDetail.FieldByName( 'UNITPRICE' ).Value > 0 ) then
      cdsDetail.FieldByName( 'UNITPRICE' ).Value :=
       ( 0 - cdsDetail.FieldByName( 'UNITPRICE' ).Value );
    if ( cdsDetail.FieldByName( 'SALEAMOUNT' ).Value > 0 ) then
      cdsDetail.FieldByName( 'SALEAMOUNT' ).Value :=
       ( 0 - cdsDetail.FieldByName( 'SALEAMOUNT' ).Value );
    if ( cdsDetail.FieldByName( 'TAXAMOUNT' ).Value > 0 ) then
      cdsDetail.FieldByName( 'TAXAMOUNT' ).Value :=
       ( 0 - cdsDetail.FieldByName( 'TAXAMOUNT' ).Value );
    if ( cdsDetail.FieldByName( 'TOTALAMOUNT' ).Value > 0 ) then
      cdsDetail.FieldByName( 'TOTALAMOUNT' ).Value :=
       ( 0 - cdsDetail.FieldByName( 'TOTALAMOUNT' ).Value );
  end
  { �������J(+)��, �۰��ର����, �άO ChargeSign ����ŭȤ]�Ѭ� (+) }
  else begin
    if ( cdsDetail.FieldByName( 'UNITPRICE' ).Value < 0 ) then
      cdsDetail.FieldByName( 'UNITPRICE' ).Value :=
       ( 0 - cdsDetail.FieldByName( 'UNITPRICE' ).Value );
    if ( cdsDetail.FieldByName( 'SALEAMOUNT' ).Value < 0 ) then
      cdsDetail.FieldByName( 'SALEAMOUNT' ).Value :=
       ( 0 - cdsDetail.FieldByName( 'SALEAMOUNT' ).Value );
    if ( cdsDetail.FieldByName( 'TAXAMOUNT' ).Value < 0 ) then
      cdsDetail.FieldByName( 'TAXAMOUNT' ).Value :=
       ( 0 - cdsDetail.FieldByName( 'TAXAMOUNT' ).Value );
    if ( cdsDetail.FieldByName( 'TOTALAMOUNT' ).Value < 0 ) then
      cdsDetail.FieldByName( 'TOTALAMOUNT' ).Value :=
       ( 0 - cdsDetail.FieldByName( 'TOTALAMOUNT' ).Value );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.UpdateDetailSourceByBillingNo;
var
  aDataSet: TADOQuery;
begin
  aDataSet := dtmSO.adoComm;
  aDataSet.Close;
  aDataSet.SQL.Text := Format(
    '  SELECT 33 AS SOURCETABLE             ' +
    '    FROM SO033 A                       ' +
    '   WHERE A.BILLNO = ''%s''             ' +
    '     AND A.ITEM = ''%s''               ' +
    '  UNION ALL                            ' +
    '  SELECT 34 AS SOURCETABLE             ' +
    '    FROM SO034 A                       ' +
    '   WHERE A.BILLNO = ''%s''             ' +
    '     AND A.ITEM = ''%s''               ',
    [ cdsDetail.FieldByName( 'BILLID' ).AsString,
      cdsDetail.FieldByName( 'BILLIDITEMNO' ).AsString,
      cdsDetail.FieldByName( 'BILLID' ).AsString,
      cdsDetail.FieldByName( 'BILLIDITEMNO' ).AsString] );
  aDataSet.Open;
  cdsDetail.FieldByName( 'SOURCE' ).AsString := aDataSet.FieldByName( 'SOURCETABLE' ).AsString;
  aDataSet.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.UpdateMasterPrintCount;
var
  aSQL, aCanModify: String;
begin
  aSQL := Format(
    ' SELECT PRINTCOUNT, CANMODIFY FROM INV007       ' +
    '  WHERE IDENTIFYID1 = ''%s''                    ' +
    '    AND IDENTIFYID2 = ''%s''                    ' +
    '    AND INVID = ''%s''                          ' +
    '    AND COMPID = ''%s''                         ',
    [ IDENTIFYID1, IDENTIFYID2, cdsMaster.FieldByName( 'INVID' ).AsString,
      dtmMain.getCompID ] );
  dtmMain.adoComm.Close;
  dtmMain.adoComm.SQL.Text := aSQL;
  dtmMain.adoComm.Open;
  if not dtmMain.adoComm.IsEmpty then
  begin
    cdsMaster.ReadOnly := False;
    try
      cdsMaster.Edit;
      cdsMaster.FieldByName( 'PRINTCOUNT' ).AsInteger :=
        dtmMain.adoComm.FieldByName( 'PRINTCOUNT' ).AsInteger;
      cdsMaster.FieldByName( 'CANMODIFY' ).AsString  :=
        dtmMain.adoComm.FieldByName( 'CANMODIFY' ).AsString;
      aCanModify := '�_';
      if dtmMain.adoComm.FieldByName( 'CANMODIFY' ).AsString = 'Y' then
        aCanModify := '�O';
      cdsMaster.FieldByName( 'CANMODIFYDESC' ).AsString  := aCanModify;
      cdsMaster.Post;
    finally
      cdsMaster.ReadOnly := True;
    end;
  end;
  dtmMain.adoComm.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.UpdateDetailMisInfo;
var
  aFacisNo, aAccountNo: String;
begin
  if not IsMergeRecord( FInvoiceId,
    cdsDetail.FieldByName( 'SEQ' ).AsString ) then
  begin
    dtmSO.GetMisAccountAndFacisNo(
      cdsDetail.FieldByName( 'BILLID' ).AsString,
      cdsDetail.FieldByName( 'BILLIDITEMNO' ).AsString,
      cdsMaster.FieldByName( 'CUSTID' ).AsString, aAccountNo,aFacisNo );
      cdsDetail.FieldByName( 'FACISNO' ).AsString := aFacisNo;
      cdsDetail.FieldByName( 'ACCOUNTNO' ).AsString := aAccountNo;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.btnLocaleCreateClick(Sender: TObject);
begin
  Screen.Cursor := crSQLWait;
  try
    FInvCreateType := icLocale;
    DMLModeChange( dmInsert );
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.btnNormalCreateClick(Sender: TObject);
begin
  Screen.Cursor := crSQLWait;
  try
    FInvCreateType := icNormal;
    DMLModeChange( dmInsert );
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.btnEditClick(Sender: TObject);
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

procedure TfrmInvA02_4.btnCancelClick(Sender: TObject);
begin
  Screen.Cursor := crSQLWait;
  try
    DMLModeChange( dmCancel );
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.btnSaveClick(Sender: TObject);
begin
  Screen.Cursor := crSQLWait;
  try
    DMLModeChange( dmSave );
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.btnQuitClick(Sender: TObject);
begin
  Self.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.cdsMasterAfterScroll(DataSet: TDataSet);
begin
  Screen.Cursor := crSQLWait;
  try
    if ( not cdsMaster.IsEmpty ) and
       ( cdsMaster.FieldByName( 'HOWTOCREATE' ).AsString <> '' ) then
    begin
      FInvCreateType := StrToInvoicType(
        cdsMaster.FieldByName( 'HOWTOCREATE' ).AsString );
    end;
    FInvoiceId := cdsMaster.FieldByName( 'INVID' ).AsString;
    MasterDataToEditor;
    FMultiInv := dtmMainJ.IsMultiInv( FInvoiceId );
    { �ϦV�^�񦸱i�o���� INV099 ���c }
    FillChooseInv;
    { �a�o������ }
    DisableDetailDataSetEvent;
    try
      { INV008 }
      PrepareDetailDataSet;
      FillDetailDataSet;
      { INV008A }
      PrepareInv008ADataSet;
      FillInv008ADataSet;
    finally
      EnableDetailDataSetEvent;
    end;
    SetDetailFieldState;
    ComposeBillingString( FBillingInfo );
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.txtCUSTIDExit(Sender: TObject);
begin
  if ( FMode in [dmInsert,dmUpdate] ) then
  begin
    if ( txtCUSTID.EditingText <> EmptyStr ) then
    begin
      if ValidateCustomer( txtCUSTID.EditingText, FInvCreateType ) then
        GetCustomerInfo( txtCUSTID.EditingText )
      else begin
        if ( FInvCreateType in [icLocale, icPrev, icAfter] ) then
          WarningMsg( Format( '�ȪA�t�άd�L���Ƚs(%s), �Э��s��J�C',
            [txtCUSTID.EditingText] ) )
        else
          WarningMsg( Format( '�o���t�άd�L���Ƚs(%s), �Э��s��J�C',
            [txtCUSTID.EditingText] ) );
        if txtCUSTID.CanFocus then
        begin
          txtCUSTID.SetFocus;
          txtCUSTID.SelectAll;
        end;
      end;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.txtBUSINESSIDExit(Sender: TObject);
begin
  if ( FMode in [dmInsert,dmUpdate] ) then
  begin
    if ( txtBUSINESSID.EditingText <> '' ) then
    begin
      if not dtmMain.checkBusinessID( txtBUSINESSID.EditingText ) then
      begin
        WarningMsg( '�Τ@�s����J���~�C' );
        if txtBUSINESSID.CanFocus then
        begin
          txtBUSINESSID.SetFocus;
          txtBUSINESSID.SelectAll;
        end;
      end;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.txtCUSTIDPropertiesButtonClick(Sender: TObject;
  AButtonIndex: Integer);
begin
  if AButtonIndex = 0 then
  begin
    if ( FMode in [dmInsert,dmUpdate] ) then
    begin
      // �{���}��, �w�}, ��}���o��, �Ƚs�ӷ��� SMS �����O���
      if ( FInvCreateType in [icLocale, icPrev, icAfter] ) then
      begin
        if txtCUSTID.EditingText = '' then
        begin
          WarningMsg( '�п�J�Ƚs�C' );
          Exit;
        end;
        if not ValidateCustomer( txtCUSTID.EditingText, FInvCreateType ) then
        begin
          WarningMsg( Format( '�ȪA�t�άd�L���Ƚs(%s), �Э��s��J�C',
            [txtCUSTID.EditingText] ) );
          Exit;
        end;
        { Pop Chioce SMS Data Window }
        if ( dtmMain.GetLinkToMIS ) then
        begin
          if PopChargeWindow( txtCUSTID.EditingText ) = mrOk then
          begin
            CopyChargeDataToDetail;
            CopyChargeTitleToMaster( FInvCreateType );
            UpdateTaxToMaster;
            UpdatePriceToMaster;
            FreeAndNil( FChargeWindow );
            if ( txtCUSTSNAME.CanFocus ) then
            begin
              txtCUSTSNAME.SetFocus;
              txtCUSTSNAME.SelectAll;
            end;
          end;
        end;  
      end else
      // �@��}�ߪ��o��, �Ƚs�ӷ����o���t�Ϊ����O���
      if ( FInvCreateType in [icNormal] ) then
      begin
         { Pop Chioce Inv Data Window }
         if PopTitleWindow2( txtCUSTID.EditingText ) = mrOk then
         begin
           CopyChargeTitleToMaster( FInvCreateType );
           FreeAndNil( FSearchCustWindow );
           txtCUSTID.OnExit := nil;
           try
             if ( txtCUSTSNAME.CanFocus ) then
             begin
               txtCUSTSNAME.SetFocus;
               txtCUSTSNAME.SelectAll;
             end;
           finally
             txtCUSTID.OnExit := txtCUSTIDExit;
           end;
         end;
      end;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.txtINVDATEPropertiesInitPopup(Sender: TObject);
var
  aDataSet: TADOQuery;
  aInvDate: TDateTime;
begin
  if FMode in [dmInsert] then
  begin
    if VarToStr( txtINVDATE.EditText ) = '' then
    begin
      WarningMsg( '�п�J�o������C' );
      Abort;
    end;
    if not TryStrToDate( txtINVDATE.EditingText, aInvDate ) then
    begin
      WarningMsg( '�z��J���o����������T�C' );
      Abort;
    end;
    aDataSet := TADOQuery( dsChooseInv.DataSet );
    aDataSet.Close;
    aDataSet.SQL.Text := Format(
      ' SELECT SUBSTR( A.YEARMONTH, 1, 4 ) ||                    ' +
      '        SUBSTR( A.YEARMONTH, 5, 2 ) YEARMONTH,            ' +
      '        A.INVFORMAT,                                      ' +
      '        DECODE( A.INVFORMAT,                              ' +
      '                ''1'', ''�q�l'',                          ' +
      '                ''2'', ''��G'',                          ' +
      '                ''3'', ''��T'',                          ' +
      '                A.INVFORMAT ) INVFORMATDESC,              ' +
      '        A.PREFIX,                                         ' +
      '        A.STARTNUM,                                       ' +
      '        A.ENDNUM,                                         ' +
      '        A.CURNUM,                                         ' +
      '        TO_CHAR( A.LASTINVDATE,                           ' +
      '                 ''YYYY/MM/DD'' ) LASTINVDATE,            ' +
      '        A.MEMO                                            ' +
      '   FROM INV099 A                                          ' +
      '  WHERE A.IDENTIFYID1 = ''%s''                            ' +
      '    AND A.IDENTIFYID2 = %s                                ' +
      '    AND A.COMPID = ''%s''                                 ' +
      '    AND A.YEARMONTH = ''%s''                              ' +
      '    AND A.USEFUL = ''Y''                                  ' +
      '    AND A.LASTINVDATE <= TO_DATE( ''%s'', ''YYYYMMDD'' )  ' +
      '    AND A.CURNUM <= A.ENDNUM                              ',
      [ IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID,
        MapingInvoiceYearMonth( aInvDate ),
        FormatDateTime( 'yyyymmdd', aInvDate ) ] );
    aDataSet.Open;
    ChooseGrid.ActiveLevel.GridView := ChooseInvView;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.txtCHARGEDATEPropertiesInitPopup(Sender: TObject);
begin
  if not ( FMode in [dmInsert, dmUpdate] ) then Abort;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.rdoTAXRATE1Click(Sender: TObject);
begin
  if ( FMode in [dmInsert, dmUpdate] ) then
  begin
    if Sender = rdoTAXRATE1 then
    begin
      txtTAXRATE.Enabled := True;
      cdsMaster.FieldByName( 'TAXRATE' ).AsInteger := 5;
    end
    else begin
      txtTAXRATE.Enabled := False;
      cdsMaster.FieldByName( 'TAXRATE' ).AsInteger := 0;
    end;
    Screen.Cursor := crSQLWait;
    try
      UpdateTaxToDetail;
      UpdateDeatilPrice;
      if ( cdsDetail.IsEmpty ) then cdsDetail.Append;
    finally
      Screen.Cursor := crDefault;
    end;  
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.txtTAXRATEExit(Sender: TObject);
begin
  if ( FMode in [dmInsert, dmUpdate] ) and
     ( cdsMaster.State in [dsEdit, dsInsert] ) then
  begin
    if ( rdoTAXRATE1.Checked ) and ( txtTAXRATE.EditValue <= 0 ) then
      cdsMaster.FieldByName( 'TAXRATE' ).AsFloat := 5;
    UpdateDeatilPrice;
    if ( cdsDetail.IsEmpty ) then cdsDetail.Append;    
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.txtINVDATEPropertiesValidate(Sender: TObject;
  var DisplayValue: Variant; var ErrorText: TCaption; var Error: Boolean);
begin
  if Error then
  begin
    WarningMsg( '�z��J���o����������T�C' );
    ErrorText := EmptyStr;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.txtCHARGEDATEPropertiesValidate(Sender: TObject;
  var DisplayValue: Variant; var ErrorText: TCaption; var Error: Boolean);
begin
  if Error then
  begin
    WarningMsg( '�z��J�����O��������T�C' );
    ErrorText := '';
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.txtINVDATEPropertiesCloseUp(Sender: TObject);
begin
  dsChooseInv.DataSet.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.txtINVDATEExit(Sender: TObject);
var
  aDate: TDateTime;
  aDataSet: TADOQuery;
  aRealInvMonth: String;
  aNoInv: Boolean;
begin
  Screen.Cursor := crSQLWait;
  try
    if FMode in [dmInsert] then
    begin
      if cdsMaster.State in [dsEdit, dsInsert] then
      begin
        if txtINVDATE.Text = EmptyStr then
        begin
          cdsMaster.FieldByName( 'INVID' ).AsString := EmptyStr;
          cdsMaster.FieldByName( 'INVFORMAT' ).AsString := EmptyStr;
          cdsMaster.FieldByName( 'INVFORMATDESC' ).AsString := EmptyStr;
          cdsMaster.FieldByName( 'BUSINESSID' ).AsString := EmptyStr;
          cxLabel6.Style.StyleController := nil;
          txtBUSINESSID.Properties.ReadOnly := False;
          FChooseInv.YearMonth := EmptyStr;
          FChooseInv.InvFormat := EmptyStr;
          FChooseInv.Prefix := EmptyStr;
          FChooseInv.StartNum := EmptyStr;
          FChooseInv.OldInvDate := EmptyStr;
          FInvoiceId := EmptyStr;
          Exit;
        end;
        if FChooseInv.YearMonth = EmptyStr then Exit;
        { �P�_��ʧ�諸���, ��쥻�w�����쪺�o�����ݦ~��O�_�@�� }
        if not TryStrToDate( txtINVDATE.Text, aDate ) then
        begin
          WarningMsg( '�z��J���o����������T�C' );
          if txtINVDATE.CanFocus then txtINVDATE.SetFocus;
          Exit;
        end;
        aNoInv := False;
        aRealInvMonth := MapingInvoiceYearMonth( aDate );
        { �����쪺�o�����ݦ~��O�_�@�� }
        if aRealInvMonth <> FChooseInv.YearMonth then
        begin
          cdsMaster.FieldByName( 'INVID' ).AsString := '';
          cdsMaster.FieldByName( 'INVFORMAT' ).AsString := '';
          cdsMaster.FieldByName( 'INVFORMATDESC' ).AsString := '';
          FChooseInv.YearMonth := EmptyStr;
          FChooseInv.InvFormat := EmptyStr;
          FChooseInv.Prefix := EmptyStr;
          FChooseInv.StartNum := EmptyStr;
          FChooseInv.OldInvDate := EmptyStr;
          Exit;
        end;
        { �Y�O�����J���T, �O���U�� }
        FChooseInv.OldInvDate :=  FormatDateTime( 'yyyymmdd', aDate );
        { �Y�O�@��, �n�ˬd�̫�o����Φ��~��S���o���� }
        aDataSet :=  TADOQuery( dtmMain.adoComm );
        aDataSet.Close;
        aDataSet.SQL.Text := Format(
          ' SELECT LASTINVDATE FROM INV099             ' +
          '  WHERE IDENTIFYID1  = ''%s''               ' +
          '    AND IDENTIFYID2  = ''%s''               ' +
          '    AND COMPID = ''%s''                     ' +
          '    AND YEARMONTH = ''%s''                  ' +
          '    AND PREFIX = ''%s''                     ' +
          '    AND STARTNUM = ''%s''                   ' +
          '    AND INVFORMAT = ''%s''                  ',
          [ IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, FChooseInv.YearMonth,
            FChooseInv.Prefix, FChooseInv.StartNum, FChooseInv.InvFormat ] );
        aDataSet.Open;
        aNoInv := aDataSet.IsEmpty;
        // �����, �P�_����O�_�j��̫�o���}�ߤ�
        if not aNoInv then
          aNoInv := ( aDate < aDataSet.FieldByName( 'LASTINVDATE' ).AsDateTime );
        if ( aNoInv ) then
        begin
          WarningMsg( '�z��J���o������p��o�������i�̫�o���}�ߤ�j�C' );
          if txtINVDATE.CanFocus then txtINVDATE.SetFocus;
        end;
        aDataSet.Close;
      end;
    end else
    if ( FMode in [dmUpdate] ) then
    begin
      { �P�_��ʧ�諸���, ��쥻�w�����쪺�o�����ݦ~��O�_�@�� }
      if not TryStrToDate( txtINVDATE.Text, aDate ) then
      begin
        WarningMsg( '�z��J���o����������T�C' );
        if txtINVDATE.CanFocus then txtINVDATE.SetFocus;
        Exit;
      end;
      aRealInvMonth := MapingInvoiceYearMonth( aDate );
      { �����쪺�o�����ݦ~��O�_�@�� }
      if aRealInvMonth <> FChooseInv.YearMonth then
      begin
        WarningMsg( '�z�ܧ󪺵o������w���P�󦹱i�o�����ݡi�o���~��j, �Э��s��J�o������C' );
        if txtINVDATE.CanFocus then txtINVDATE.SetFocus; 
      end;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.txtCUSTIDPropertiesChange(Sender: TObject);
var
  aCanEmptyCustInfo: Boolean;
begin
  if ( FMode in [dmInsert] ) and ( FInvCreateType in [icLocale] ) then
  begin
    if ( Trim( txtCUSTID.EditingText ) = '' ) then
    begin
      { to avoid cdsMaster.Cancel trigger this event,
        only at Edit or Insert Mode Change the CustId to empty the cdsDetail DataSet  }
      if cdsMaster.State in [dsEdit, dsInsert] then
        cdsDetail.EmptyDataSet;
    end;  
  end;
  aCanEmptyCustInfo :=
    ( FMode in [dmInsert, dmUpdate] ) and ( cdsMaster.State in [dsEdit, dsInsert] );
  if ( Trim( txtCUSTID.EditingText ) = EmptyStr ) and aCanEmptyCustInfo then
  begin
    cdsMaster.FieldByName( 'CUSTSNAME' ).AsString := '';
    cdsMaster.FieldByName( 'INVTITLE' ).AsString := '';
    cdsMaster.FieldByName( 'BUSINESSID' ).AsString := '';
    cdsMaster.FieldByName( 'INVADDR' ).AsString := '';
    cdsMaster.FieldByName( 'MAILADDR' ).AsString := '';
    cdsMaster.FieldByName( 'CHARGEDATE' ).AsString := '';
    cdsMaster.FieldByName( 'ZIPCODE' ).AsString := '';
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.ChooseInvViewDblClick(Sender: TObject);
var
  aDataSet: TADOQuery;
  aInvDate: TDateTime;
begin
  if ( FMode in [dmInsert] ) then
  begin
    aDataSet :=  TADOQuery( ChooseInvView.DataController.DataSource.DataSet );
    if aDataSet.IsEmpty then Exit;
    if not TryStrToDate( txtINVDATE.EditingText, aInvDate ) then
    begin
      WarningMsg( '�z��J���o����������T�C' );
      Exit;
    end;
    if ( aDataSet.FieldByName( 'PREFIX' ).AsString = EmptyStr ) then
    begin
      WarningMsg( '�ӵo�����d�L�o���r�y, �п�ܨ䥦�o�����C' );
      Exit;
    end;
    cdsMaster.FieldByName( 'INVID' ).AsString :=
      aDataSet.FieldByName( 'PREFIX' ).AsString +
      Lpad( aDataSet.FieldByName( 'CURNUM' ).AsString, 8, '0' );
    cdsMaster.FieldByName( 'MAININVID' ).AsString :=
      cdsMaster.FieldByName( 'INVID' ).AsString;
    cdsMaster.FieldByName( 'INVFORMAT' ).AsString :=
      aDataSet.FieldByName( 'INVFORMAT' ).AsString;
    cdsMaster.FieldByName( 'INVFORMATDESC' ).AsString :=
      aDataSet.FieldByName( 'INVFORMATDESC' ).AsString;
    cdsMaster.FieldByName( 'INVDATE' ).AsDateTime :=
      aDataSet.FieldByName( 'LASTINVDATE' ).AsDateTime;
    FChooseInv.YearMonth := aDataSet.FieldByName( 'YEARMONTH' ).AsString;
    FChooseInv.InvFormat := aDataSet.FieldByName( 'INVFORMAT' ).AsString;
    FChooseInv.Prefix := aDataSet.FieldByName( 'PREFIX' ).AsString;
    FChooseInv.StartNum := aDataSet.FieldByName( 'STARTNUM' ).AsString;
    FChooseInv.OldInvDate := FormatDateTime( 'yyyymmdd',
      aDataSet.FieldByName( 'LASTINVDATE' ).AsDateTime );
    FInvoiceId := cdsMaster.FieldByName( 'INVID' ).AsString;
    {}
    cxLabel6.Style.StyleController := nil;
    txtBUSINESSID.Properties.ReadOnly := False;
    { ��G, ���i��J�νs }
    if ( cdsMaster.FieldByName( 'INVFORMAT' ).AsString = '2' ) then
    begin
      cdsMaster.FieldByName( 'BUSINESSID' ).AsString := EmptyStr;
      txtBUSINESSID.Properties.ReadOnly := True;
    end;
    { ��T, ������J�νs }
    if ( cdsMaster.FieldByName( 'INVFORMAT' ).AsString = '3' ) then
      cxLabel6.Style.StyleController := dtmMain.cxLabelNotNullStyle;

    {}
    txtINVDATE.Deactivate;
    txtINVDATE.OnExit := nil;
    try
      if txtCUSTID.CanFocus then txtCUSTID.SetFocus;
    finally
      txtINVDATE.OnExit := txtINVDATEExit;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.ChooseInvViewKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Ord( Key ) = VK_RETURN then
    ChooseInvViewDblClick( ChooseInvView );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.cdsMasterNewRecord(DataSet: TDataSet);
begin
  case FInvCreateType of
    icLocale: cdsMaster.FieldByName( 'HOWTOCREATEDESC' ).AsString := '�{���}��';
    icNormal: cdsMaster.FieldByName( 'HOWTOCREATEDESC' ).AsString := '�@��}��';
    icPrev: cdsMaster.FieldByName( 'HOWTOCREATEDESC' ).AsString := '�w�}';
    icAfter: cdsMaster.FieldByName( 'HOWTOCREATEDESC' ).AsString := '��}';
  end;
  cdsMaster.FieldByName( 'INVDATE' ).AsDateTime := Date;
  cdsMaster.FieldByName( 'TAXTYPE' ).AsString := '1';
  cdsMaster.FieldByName( 'TAXRATE' ).AsString := '5';
  cdsMaster.FieldByName( 'SALEAMOUNT' ).AsInteger := 0;
  cdsMaster.FieldByName( 'TAXAMOUNT' ).AsInteger := 0;
  cdsMaster.FieldByName( 'INVAMOUNT' ).AsInteger := 0;
  cdsMaster.FieldByName( 'MAINSALEAMOUNT' ).AsInteger := 0;
  cdsMaster.FieldByName( 'MAINTAXAMOUNT' ).AsInteger := 0;
  cdsMaster.FieldByName( 'MAININVAMOUNT' ).AsInteger := 0;
  cdsMaster.FieldByName( 'ISOBSOLETEDESC' ).AsString := '�_';
  cdsMaster.FieldByName( 'CANMODIFYDESC' ).AsString := '�O';
  cdsMaster.FieldByName( 'PRINTCOUNT' ).AsInteger := 0;
  FChooseInv.YearMonth := EmptyStr;
  FChooseInv.InvFormat := EmptyWideStr;
  FChooseInv.Prefix := EmptyStr;
  FChooseInv.StartNum := EmptyStr;
  FChooseInv.OldInvDate := EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.cdsDetailBeforeInsert(DataSet: TDataSet);
begin
  // �����ק�Ҧ���, �w�W�X�]�w�۰ʴ��}����, �����\�s�W }
  // ����s�ɮ��ˬd
  //  if ( FMode in [dmUpdate] ) and
  //     ( dtmMain.GetAutoCreateNum > 0 ) and
  //     ( cdsDetail.RecordCount >= dtmMain.GetAutoCreateNum ) then
  //    Abort;
  FDetailMaxSequence := DetailSequenceNeeded;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.cdsDetailNewRecord(DataSet: TDataSet);
begin
  cdsDetail.FieldByName( 'IDENTIFYID1' ).AsString := IDENTIFYID1;
  cdsDetail.FieldByName( 'IDENTIFYID2' ).AsString := IDENTIFYID2;
  cdsDetail.FieldByName( 'INVID' ).AsString := cdsMaster.FieldByName( 'INVID' ).AsString;
  cdsDetail.FieldByName( 'SEQ' ).AsInteger := FDetailMaxSequence;
  cdsDetail.FieldByName( 'QUANTITY' ).AsInteger := 1;
  cdsDetail.FieldByName( 'UNITPRICE' ).AsInteger := 0;
  cdsDetail.FieldByName( 'SALEAMOUNT' ).AsInteger := 0;
  cdsDetail.FieldByName( 'TAXAMOUNT' ).AsInteger := 0;
  cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsInteger := 0;
  cdsDetail.FieldByName( 'SIGN' ).AsString := '+';
  // �w�}, ��}, �{���}
  if FInvCreateType in [icPrev, icAfter, icLocale] then
  begin
    cdsDetail.FieldByName( 'EMPTY' ).AsString := 'Y';
    cdsDetail.FieldByName( 'LINKTOMIS' ).AsString :=
      IIF( dtmMain.GetLinkToMIS, 'Y', 'N' );
    cdsDetail.FieldByName( 'LINKTOMISDESC' ).AsString := '�O'
  end else
  begin
    // �@��}�ߤ]�n���J ���O�渹�ζ���, �A�ȧO, ���O��_��
    cdsDetail.FieldByName( 'EMPTY' ).AsString := 'Y';
    cdsDetail.FieldByName( 'LINKTOMIS' ).AsString := 'N';
    cdsDetail.FieldByName( 'LINKTOMISDESC' ).AsString := '�_';
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.cdsDetailAfterDelete(DataSet: TDataSet);
begin
  UpdatePriceToMaster;
  ComposeBillingString( FBillingInfo );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.cdsDetailBeforeDelete(DataSet: TDataSet);
var
  aMsg: String;
begin
  { �P�B INV008ADataSet, �s�ɮɼg�J INV008A }
  if not SyncInv008ADataSet( dmDelete,
    cdsDetail.FieldByName( 'SEQ' ).AsString, EmptyStr, FInvoiceId, aMsg ) then
  begin
    ErrorMsg( Format( '�P�B�o���X�ָ���ɿ��~, �T��:%s�C', [aMsg] ) );
    Abort;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.cdsDetailBeforePost(DataSet: TDataSet);
var
  aMsg: String;
begin
  // ���ʧ@�w�����n, �u����r�ɮ׶פJ�ɤ~���X�ְʧ@,
  // ���}�ߤw�����n�g�J INV008A
  { �P�B INV008ADataSet, �s�ɮɼg�J INV008A }
  if not SyncInv008ADataSet( dmInsert,
    cdsDetail.FieldByName( 'SEQ' ).AsString, EmptyStr, FInvoiceId, aMsg ) then
  begin
    ErrorMsg( Format( '�P�B�o���X�ָ���ɿ��~, �T��:%s�C', [aMsg] ) );
    Abort;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.cdsDetailAfterPost(DataSet: TDataSet);
begin
  UpdatePriceToMaster;
  ComposeBillingString( FBillingInfo );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.DetailGridViewBILLIDPropertiesButtonClick(
  Sender: TObject; AButtonIndex: Integer);
var
  aBillId, aItemId: String;
begin
  if ( cdsMaster.FieldByName( 'CUSTID' ).AsString = EmptyStr ) then
  begin
    WarningMsg( '�Х���J�Ƚs�C' );
    Exit;
  end;
  // �{���}��, �w�}, ��}���o��, �Ƚs�ӷ��� SMS �����O���
  if ( FInvCreateType in [icLocale, icPrev, icAfter] ) then
  begin
    { Pop Chioce SMS Data Window }
    if ( dtmMain.GetLinkToMIS ) and
       ( cdsDetail.FieldByName( 'LINKTOMIS' ).AsString = 'Y' ) then
    begin
      if PopChargeWindow( cdsMaster.FieldByName( 'CUSTID' ).AsString ) = mrOk then
      begin
        CopyChargeDataToDetail;
        UpdateTaxToMaster;
        UpdatePriceToMaster;
        FreeAndNil( FChargeWindow );
      end;
    end;  
  end
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.DetailGridViewBILLIDPropertiesValidate(
  Sender: TObject; var DisplayValue: Variant; var ErrorText: TCaption;
  var Error: Boolean);
begin
  if DisplayValue = EmptyStr then Exit;
  // �D�w�}, ��}, �{���}��, ��������J�Ƚs
  if ( FInvCreateType in [icPrev, icAfter, icLocale] ) then
  begin
    Error := ( cdsMaster.FieldByName( 'CUSTID' ).AsString = EmptyStr );
    if Error then
    begin
      WarningMsg( '�Х���J�Ƚs�C' );
      ErrorText := '';
      Abort;
    end;
  end;  
  { �P�ȪA�t�Τ����~�ˬd }
  if ( dtmMain.GetLinkToMIS ) and
     ( cdsDetail.FieldByName( 'LINKTOMIS' ).AsString = 'Y' ) then
  begin
    Screen.Cursor := crSQLWait;
    try
      Error := not ValidateBillingNo( DisplayValue, EmptyStr, String( ErrorText ) );
    finally
      Screen.Cursor := crDefault;
    end;
    if Error then
    begin
      WarningMsg( ErrorText );
      ErrorText := EmptyStr;
      Abort;
    end;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.DetailGridViewBILLIDITEMNOPropertiesValidate(
  Sender: TObject; var DisplayValue: Variant; var ErrorText: TCaption;
  var Error: Boolean);
begin
  if ( cdsDetail.FieldByName( 'BILLID' ).AsString = EmptyStr  ) or
     ( DisplayValue = EmptyStr ) then Exit;
  // �D�w�}, ��}, �{���}��, ��������J�Ƚs
  if ( FInvCreateType in [icPrev, icAfter, icLocale] ) then
  begin
    Error := ( cdsMaster.FieldByName( 'CUSTID' ).AsString = EmptyStr );
    if Error then
    begin
      WarningMsg( '�Х���J�Ƚs�C' );
      ErrorText := '';
      Abort;
    end;
  end;  
  { �P�ȪA�t�Τ����~�ˬd }
  if ( dtmMain.GetLinkToMIS ) and
     ( cdsDetail.FieldByName( 'LINKTOMIS' ).AsString = 'Y' )  then
  begin
    Screen.Cursor := crSQLWait;
    try
      Error := not ValidateBillingNo( cdsDetail.FieldByName( 'BILLID' ).AsString,
        DisplayValue, String( ErrorText ) );
    finally
      Screen.Cursor := crDefault;
    end;
    if Error then
    begin
      WarningMsg( Format( '���O�渹: %s,  �L�����O�����C',
        [cdsDetail.FieldByName( 'BILLID' ).AsString] ) );
      ErrorText := '';
      Abort;
    end;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.DetailGridViewSTARTDATEPropertiesValidate(
  Sender: TObject; var DisplayValue: Variant; var ErrorText: TCaption;
  var Error: Boolean);
begin
  if DisplayValue = EmptyStr then Exit;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.DetailGridViewENDDATEPropertiesValidate(
  Sender: TObject; var DisplayValue: Variant; var ErrorText: TCaption;
  var Error: Boolean);
begin
  if DisplayValue = EmptyStr then Exit;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.DetailGridViewQUANTITYPropertiesValidate(
  Sender: TObject; var DisplayValue: Variant; var ErrorText: TCaption;
  var Error: Boolean);
begin
  Error := ( DisplayValue = 0 );
  if Error then
  begin
    WarningMsg( '�ƶq���i���s�C' );
    ErrorText := '';
    Abort;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.DetailGridViewITEMIDPropertiesValidate(
  Sender: TObject; var DisplayValue: Variant; var ErrorText: TCaption;
  var Error: Boolean);
var
  aInt: Integer;
begin
  if DisplayValue = EmptyStr then Exit;
  { �����\��J�D�Ʀr��� }
  if not TryStrToInt( DisplayValue, aInt ) then
  begin
    WarningMsg( '���O���إN�X���i��J�D�ƭȸ�ơC' );
    Error := True;
    ErrorText := EmptyStr;
    Abort;
  end;
  if ( ( dtmMain.GetLinkToMIS ) and
       ( cdsDetail.FieldByName( 'LINKTOMIS' ).AsString = 'Y' ) ) or
     ( FInvCreateType in [icNormal] ) then
  begin
    Screen.Cursor := crSQLWait;
    try
      Error := not ValidateChargeItemId( DisplayValue, String( ErrorText ) );
    finally
      Screen.Cursor := crDefault;
    end;
    if Error then
    begin
      WarningMsg( ErrorText );
      ErrorText := EmptyStr;
      Abort;
    end;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.DetailGridViewTAXAMOUNTPropertiesValidate(
  Sender: TObject; var DisplayValue: Variant; var ErrorText: TCaption;
  var Error: Boolean);
begin
  { �D���|, �h���J���|�B�@�ߧאּ 0 }
  if ( GetMasterTaxCode <> '1' ) then
    DisplayValue := 0;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.DetailGridViewEditing(
  Sender: TcxCustomGridTableView; AItem: TcxCustomGridTableItem;
  var AAllow: Boolean);
var
  aColumn: TcxGridDBBandedColumn;
begin
  // �w�}, ��}, �{���}��, ���P�_�O�_�n�N���ڳ渹����, ���i��J
  // �@��}��, �}��i�H���i�H��J
  aColumn := TcxGridDBBandedColumn( AItem );
  if ( aColumn.DataBinding.FieldName = 'BILLID' ) or
     ( aColumn.DataBinding.FieldName = 'BILLIDITEMNO' ) then
  begin
    AAllow := True;
    if ( FInvCreateType in [icPrev,icAfter,icLocale] ) then
      AAllow := ( cdsDetail.FieldByName( 'EMPTY' ).AsString = 'Y' );
  end else
  if ( aColumn.DataBinding.FieldName = 'LINKTOMISDESC' ) then
  begin
    AAllow := not ( FInvCreateType in [icNormal] ); 
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.DetailGridViewCustomDrawCell(
  Sender: TcxCustomGridTableView; ACanvas: TcxCanvas;
  AViewInfo: TcxGridTableDataCellViewInfo; var ADone: Boolean);
var
  aColumn, aConditionColumn: TcxGridDBBandedColumn;
  aValue, aChargeSign: String;
begin
  // �w�}, ��}, �{���}��, ���P�_�O�_�n�N���ڳ渹����, ���i��J
  // �@��}��, ������J���O�渹, �@�ߤ��i��J
  // �N���i��J���e���Ǧ�
  if ( cdsDetail.Active ) and
     ( not cdsDetail.IsEmpty ) and
     ( FMode in [dmInsert, dmUpdate] )  then
  begin
    aColumn := TcxGridDBBandedColumn( AViewInfo.Item );
    if ( aColumn.DataBinding.FieldName = 'BILLID' ) or
       ( aColumn.DataBinding.FieldName = 'BILLIDITEMNO' ) then
    begin
      aConditionColumn := DetailGridView.GetColumnByFieldName( 'EMPTY' );
      aValue := VarToStrDef( AViewInfo.GridRecord.Values[aConditionColumn.Index], 'Y' );
      if ( aValue = 'N' ) and ( FInvCreateType in [icPrev, icAfter, icLocale] ) then
        ACanvas.Brush.Color := dtmMain.cxGridNotEditStyle.Color;
    end else
    if ( aColumn.DataBinding.FieldName = 'LINKTOMISDESC' ) then
    begin
      if ( FInvCreateType in [icNormal] ) then
      begin
        ACanvas.Brush.Color := dtmMain.cxGridNotEditStyle.Color;
      end;
    end;
    { ChargeSign ���t����, �N�r���ন���� }
    if ( aColumn.DataBinding.FieldName = 'UNITPRICE' ) or
       ( aColumn.DataBinding.FieldName = 'TAXAMOUNT' ) or
       ( aColumn.DataBinding.FieldName = 'SALEAMOUNT' ) or
       ( aColumn.DataBinding.FieldName = 'TOTALAMOUNT' ) then
    aChargeSign := VarToStrDef( AViewInfo.GridRecord.Values[
      DetailGridView.GetColumnByFieldName( 'SIGN' ).Index ], EmptyStr );
    if ( aChargeSign = '-' ) then
      ACanvas.Font.Color := clRed;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.DetailGridViewITEMIDPropertiesInitPopup(Sender: TObject);
var
  aDataSet: TADOQuery;
  aEdit: TSubclassPopupEdit;
begin
  aDataSet := TADOQuery( dsChooseItem.DataSet );
  aDataSet.Close;
  aDataSet.SQL.Text := Format(
    '  SELECT A.ITEMID,                ' +
    '         A.DESCRIPTION,           ' +
    '         A.SIGN,                  ' +
    '         A.TAXCODE,               ' +
    '         A.TAXNAME                ' +
    '    FROM INV005 A                 ' +
    '   WHERE A.IDENTIFYID1 = ''%s''   ' +
    '     AND A.IDENTIFYID2 = %s       ' +
    '     AND A.COMPID = ''%s''        ' +
    '     AND A.TAXCODE = ''%s''       ' +
    '   ORDER BY A.ITEMID              ',
    [ IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, GetMasterTaxCode ] );
  aDataSet.Open;
  ChooseGrid.ActiveLevel.GridView := ChooseItemView;
  if cdsDetail.FieldByName( 'ITEMID' ).AsString <> '' then
    aDataSet.Locate( 'ITEMID', cdsDetail.FieldByName( 'ITEMID' ).AsString, [] );
  if DetailGridView.Controller.EditingController.Edit is TcxPopupEdit then
  begin
    aEdit := TSubclassPopupEdit( DetailGridView.Controller.EditingController.Edit );
    TcxPopupEditPopupWindow( aEdit.PopupWindow ).KeyPreview := True;
    TcxPopupEditPopupWindow( aEdit.PopupWindow ).OnKeyPress :=
      TSubclassPopupEditPopupWindow( aEdit.PopupWindow ).PopEditWindowKeyPress;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.ChooseItemViewDblClick(Sender: TObject);
var
  aDataSet: TADOQuery;
begin
  aDataSet :=  TADOQuery( ChooseInvView.DataController.DataSource.DataSet );
  if aDataSet.IsEmpty then Exit;
  if ( aDataSet.FieldByName( 'ITEMID' ).AsString <>
       cdsDetail.FieldByName( 'ITEMID' ).AsString ) then
  begin
    cdsDetail.Edit;
    cdsDetail.FieldByName( 'ITEMID' ).AsString :=
      aDataSet.FieldByName( 'ITEMID' ).AsString;
    cdsDetail.FieldByName( 'DESCRIPTION' ).AsString :=
      aDataSet.FieldByName( 'DESCRIPTION' ).AsString;
    cdsDetail.FieldByName( 'TAXCODE' ).AsString :=
      aDataSet.FieldByName( 'TAXCODE' ).AsString;
    cdsDetail.FieldByName( 'TAXNAME' ).AsString :=
      aDataSet.FieldByName( 'TAXNAME' ).AsString;
    cdsDetail.FieldByName( 'SIGN' ).AsString :=
      aDataSet.FieldByName( 'SIGN' ).AsString;
    DetailGridView.Controller.EditingController.HideEdit( True );
    UpdateDetailPriceByChargeSign;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.DetailGridViewITEMIDPropertiesCloseUp(
  Sender: TObject);
begin
  dsChooseItem.DataSet.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.DetailGridViewITEMIDPropertiesEditValueChanged(
  Sender: TObject);
begin
  if ( dtmMain.GetLinkToMIS ) and
     ( cdsDetail.FieldByName( 'LINKTOMIS' ).AsString = 'Y' ) then
  begin
    DetailGridView.DataController.PostEditingData;
    GetChargeItemInfo( TcxPopupEdit(
      DetailGridView.Controller.EditingController.Edit ).EditText );
    UpdateDetailPriceByChargeSign;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.DetailGridViewEditKeyDown(
  Sender: TcxCustomGridTableView; AItem: TcxCustomGridTableItem;
  AEdit: TcxCustomEdit; var Key: Word; Shift: TShiftState);
begin
  if TcxGridDBBandedColumn( AItem ).DataBinding.FieldName = 'CHARGEEN' then
  begin
    if ( Key = VK_RETURN ) and
       ( DetailGridView.Controller.FocusedRowIndex + 1 = 
         DetailGridView.DataController.RecordCount ) then
    begin
      DetailGridView.DataController.Post;
      DetailGridView.DataController.Append;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.DetailGridViewBILLIDPropertiesEditValueChanged(
  Sender: TObject);
begin
  DetailGridView.DataController.PostEditingData;
  if ( FInvCreateType in [icPrev, icAfter, icLocale] ) and
     ( dtmMain.GetLinkToMIS ) and
     ( cdsDetail.FieldByName( 'LINKTOMIS' ).AsString = 'Y' ) then
  begin
    if ( cdsDetail.FieldByName( 'BILLID' ).AsString <> EmptyStr ) and
       ( cdsDetail.FieldByName( 'BILLIDITEMNO' ).AsString <> EmptyStr ) then
    begin
      UpdateDetailSourceByBillingNo;
      UpdateDetailMisInfo;
    end else
    begin
      cdsDetail.FieldByName( 'FACISNO' ).AsString := EmptyStr;
      cdsDetail.FieldByName( 'ACCOUNTNO' ).AsString := EmptyStr;
    end;
  end
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.DetailGridViewQUANTITYPropertiesEditValueChanged(
  Sender: TObject);
begin
  DetailGridView.DataController.PostEditingData;
  CalculateChargePrice( chQuantity );
  UpdateDetailPriceByChargeSign;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.DetailGridViewUNITPRICEPropertiesEditValueChanged(
  Sender: TObject);
begin
  DetailGridView.DataController.PostEditingData;
  CalculateChargePrice( chUnitPrice );
  UpdateDetailPriceByChargeSign;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.DetailGridViewTAXAMOUNTPropertiesEditValueChanged(
  Sender: TObject);
begin
  DetailGridView.DataController.PostEditingData;
  CalculateChargePrice( chTaxAmount );
  UpdateDetailPriceByChargeSign;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.DetailGridViewSALEAMOUNTPropertiesEditValueChanged(
  Sender: TObject);
begin
  DetailGridView.DataController.PostEditingData;
  CalculateChargePrice( chSaleAmount );
  UpdateDetailPriceByChargeSign;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.DetailGridViewTOTALAMOUNTPropertiesEditValueChanged(
  Sender: TObject);
begin
  DetailGridView.DataController.PostEditingData;
  CalculateChargePrice( chTotalAmount );
  UpdateDetailPriceByChargeSign;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.DetailGridExit(Sender: TObject);
begin
  if ( FMode in [dmInsert, dmUpdate] ) and ( cdsDetail.IsEmpty ) then
    cdsDetail.Append;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.btnPrintClick(Sender: TObject);
var
  aErrMsg: String;
begin
  if not ValidateCanPrint( aErrMsg ) then
  begin
    WarningMsg( aErrMsg );
    Exit;
  end;
  frmInvA05_1 := TfrmInvA05_1.Create(Application);
  try
    frmInvA05_1.sG_UserID := dtmMain.getLoginUser;
    frmInvA05_1.sG_CompID := dtmMain.getCompID;
    frmInvA05_1.sG_OrderByType := '1';
    frmInvA05_1.CallFromProgram := Self.ClassType;
    frmInvA05_1.InvId := cdsMaster.FieldByName( 'INVID' ).AsString;
    frmInvA05_1.ShowModal;
  finally
    frmInvA05_1.Free;
    frmInvA05_1 := nil;
  end;
  { Reload ���} 
  FormShow( Self );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.btnDropClick(Sender: TObject);
var
  aResult: Integer;
  aMsg: String;
begin
  if not ValidateCanDrop( aMsg ) then
  begin
    WarningMsg( aMsg );
    Exit;
  end;
  frmInvB02 := TfrmInvB02.Create( nil );
  try
    frmInvB02.InvId := FInvoiceId;
    aResult := frmInvB02.ShowModal;
  finally
    frmInvB02.Free;
    frmInvB02 := nil;
  end;
  if ( aResult = mrOk ) then { Reload ���} FormShow( Self );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.txtMainInvIdPropertiesButtonClick(Sender: TObject;
  AButtonIndex: Integer);
var
  aResult: String;
begin
  if ( cdsMaster.IsEmpty ) then Exit;
  if not ( FMode in [dmBrowser] ) then Exit;
  if PopRelationInvId( txtMainInvId.Text ) = mrOk then
  begin
    aResult := TfrmInvA02_6( FRelationInvWindow ).InvId;
    FreeAndNil( FRelationInvWindow );
    if ( aResult <> EmptyStr ) and ( FInvoiceId <> aResult ) then
    begin
      FInvoiceId := aResult;
      FormShow( Self );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.DetailGridViewLinkToMISPropertiesEditValueChanged(
  Sender: TObject);
begin
  cdsDetail.FieldByName( 'LINKTOMIS' ).AsString := IIF(
    cdsDetail.FieldByName( 'LINKTOMISDESC' ).AsString = '�O', 'Y', 'N' );
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA02_4.IsLinkToMis(const aInvId, aSeq: String): Boolean;
begin
  Result := False;
  cdsDetail.DisableControls;
  try
    if cdsDetail.Locate( 'INVID;SEQ', VarArrayOf( [aInvId,aSeq] ), [] ) then
      Result := ( cdsDetail.FieldByName( 'LINKTOMIS' ).AsString = 'Y' );
  finally
    cdsDetail.EnableControls;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA02_4.IsMergeRecord(const aInvId, aSeq: String): Boolean;
begin
  Result := cdsInv008A.Locate( 'INVID;SEQ', VarArrayOf( [aInvId, aSeq] ), []  );
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA02_4.SyncInv008ADataSet(const aMode: TDMLMode;
  const aOldSeq, aNewSeq, aNewInvId: String;  var aMsg: String): Boolean;

  { -------------------------------------------------- }

  procedure DoAppend;
  var
    aStartDate, aEndDate: String;
    aAccountNo, aFacisNo, aCustId: String;
  begin
    { ���ɮ� ����榡�i��O����, �|�ɦܿ��~ }
    try
      aStartDate := cdsDetail.FieldByName( 'STARTDATE' ).AsString;
    except
      aStartDate := EmptyStr;
    end;
    try
      aEndDate := cdsDetail.FieldByName( 'ENDDATE' ).AsString;
    except
      aEndDate := EmptyStr;
    end;
    aAccountNo := EmptyStr;
    aFacisNo := EmptyStr;
    { ���o�����X��CP���� }
    if ( FInvCreateType in [icPrev, icAfter, icLocale] ) and
       ( dtmMain.GetLinkToMIS ) and
       ( cdsDetail.FieldByName( 'LINKTOMIS' ).AsString = 'Y' ) then
    begin
      // �b Focus �٦b txtCustId �|�����}��, ���� MasterDataSet �� CustId ���
      //  �٬O�­ȩάO�Ū�, �o�ɪ����� txtCustId.Text
      aCustId := cdsMaster.FieldByName( 'CUSTID' ).AsString;
      if ( aCustId = EmptyStr ) or ( aCustId <> txtCUSTID.Text ) then
         aCustId := txtCUSTID.Text;
      dtmSO.GetMisAccountAndFacisNo( cdsDetail.FieldByName( 'BILLID' ).AsString,
        cdsDetail.FieldByName( 'BILLIDITEMNO' ).AsString, aCustId, aAccountNo, aFacisNo );
    end;    
    {}
    cdsInv008A.AppendRecord( [
      EmptyStr,
      FInvoiceId,
      cdsDetail.FieldByName( 'BILLID' ).AsString,
      cdsDetail.FieldByName( 'BILLIDITEMNO' ).AsString,
      aOldSeq,
      aStartDate,
      aEndDate,
      cdsDetail.FieldByName( 'ITEMID' ).AsString,
      cdsDetail.FieldByName( 'DESCRIPTION' ).AsString,
      cdsDetail.FieldByName( 'QUANTITY' ).AsInteger,
      cdsDetail.FieldByName( 'UNITPRICE' ).AsFloat,
      cdsDetail.FieldByName( 'TAXAMOUNT' ).AsInteger,
      cdsDetail.FieldByName( 'SALEAMOUNT' ).AsInteger,
      cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsInteger,
      cdsDetail.FieldByName( 'SERVICETYPE' ).AsString,
      aFacisNo,
      aAccountNo] );
  end;

  { -------------------------------------------------- }

  procedure DoEdit;
  begin
    cdsInv008A.Edit;
    cdsInv008A.FieldByName( 'INVID' ).AsString := aNewInvId;;
    cdsInv008A.FieldByName( 'SEQ' ).AsString := aNewSeq;
    cdsInv008A.Post;
  end;

  { -------------------------------------------------- }

  procedure DoDelete;
  begin
    cdsInv008A.First;
    while not cdsInv008A.Eof do
    begin
      if ( cdsInv008A.FieldByName( 'SEQ' ).AsString = aOldSeq ) then
        cdsInv008A.Delete
      else
        cdsInv008A.Next;  
    end;  
  end;

  { -------------------------------------------------- }

begin
  aMsg := EmptyStr;
  { �����Ӹ�Ʋ��ʮ�, �P�B Inv008A ���, �� OldSeq ����     }
  { �����J�u�ѤU���o�����ӳQ�R����, �Y INV008A ���X�ֶ��خ�, �h�R��,
    �s�ɮɷ|�h��s�ȪA�t��, �s�W�έק�Τ��� }
  if ( aMode in [dmInsert] ) then
  begin
    if not cdsInv008A.Locate( 'SEQ', aOldSeq, [] ) then
    begin
      try
        DoAppend;
      except
        on E: Exception do aMsg := E.Message;
      end;
    end;
  end else
  if ( aMode = dmUpdate ) then
  begin
    try
      if cdsInv008A.Locate( 'INVID;SEQ', VarArrayOf( [FInvoiceId, aOldSeq] ), [] ) then
        DoEdit;
    except
      on E: Exception do aMsg := E.Message;
    end;
  end else
  if ( aMode = dmDelete ) then
  begin
    try
      DoDelete;
    except
      on E: Exception do aMsg := E.Message;
    end;
  end;
  Result := ( aMsg = EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.txtInvUseCodeNamePropertiesInitPopup(
  Sender: TObject);
var
  aSql: String;
begin
  if ( cdsMaster.FieldByName( 'BUSINESSID' ).AsString = EmptyStr ) then
  begin
    aSql := Format(
      ' SELECT ITEMID, DESCRIPTION   ' +
      '   FROM INV028                ' +
      '  WHERE IDENTIFYID1 = ''%S''  ' +
      '    AND IDENTIFYID2 = ''%S''  ' +
      '    AND COMPID = ''%S''       ' +
      '    AND ( REFNO <> 999 OR     ' +
      '          REFNO IS NULL )     ' +
      '    ORDER BY ITEMID           ',
      [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID] );
  end else
  begin
    aSql := Format(
      ' SELECT ITEMID, DESCRIPTION     ' +
      '   FROM INV028                  ' +
      '  WHERE IDENTIFYID1 = ''%S''    ' +
      '    AND IDENTIFYID2 = ''%S''    ' +
      '    AND COMPID = ''%S''         ' +
      '    AND ( ( REFNO <> 999 AND    ' +
      '            REFNO <> 1 )        ' +
      '          OR REFNO IS NULL )    ' +
      '    ORDER BY ITEMID             ',
      [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID] );
  end;
  txtInvUse.SQL.Text := aSql;
  txtInvUse.CodeNamePropertiesInitPopup( Sender );
end;

{ ---------------------------------------------------------------------------- }

end.