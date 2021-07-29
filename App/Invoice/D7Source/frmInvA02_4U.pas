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
  dxSkinBlack, dxSkinBlue, dxSkinCaramel, dxSkinCoffee, cxCheckBox;

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

  TInvCreateType = ( icPrev = 1,       // 預開, MIS 拋檔, 預開
                     icAfter = 2,      // 後開, MIS 拋檔, 後開
                     icLocale = 3,     // 現場開立, Invoice Create 預開
                     icNormal = 4 );   // 一般開立, Invoice Create 後開

  TChargeColumn = ( chQuantity, chUnitPrice, chSaleAmount, chTotalAmount, chTaxAmount );

  TBillingInfo = record
    BillId: String;                  { 收費單號 }
    ItemNo: String;                  { 項次 }
  end;

  TTaxInfo = record
    TaxCode: String;                 { 稅別 }
    TaxRate: String;                 { 稅率 }
  end;

  TChooseInvInfo = record            { 選擇的發票記錄 }
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
    Shape3: TShape;
    grpInvoiceKind: TcxGroupBox;
    cxLabel30: TcxLabel;
    chkEInv: TcxDBCheckBox;
    cxLabel31: TcxLabel;
    DetailGridViewSmartCardNo: TcxGridDBBandedColumn;
    INV008ViewSmartCardNo: TcxGridDBColumn;
    cxLabel32: TcxLabel;
    txtMail: TcxDBTextEdit;
    cxLabel33: TcxLabel;
    txtContmobile: TcxDBTextEdit;
    chkUploadFlag: TcxDBCheckBox;
    txtGiveUnit: TDataLookup;
    cxLabel34: TcxLabel;
    lblPrintEn: TcxDBLabel;
    cxLabel35: TcxLabel;
    lblPrintTime: TcxDBLabel;
    btnLog: TcxButton;
    cxLabel36: TcxLabel;
    lblUploadTime: TcxDBLabel;
    chkObUploadFlag: TcxDBCheckBox;
    cxLabel37: TcxLabel;
    lblOBUPLOADTIME: TcxDBLabel;
    DetailGridViewCMMac: TcxGridDBBandedColumn;
    INV008ViewCMMac: TcxGridDBColumn;
    cxLabel38: TcxLabel;
    lblCreateInvDate: TcxDBLabel;
    ChooseInvViewUPLOADFLAG: TcxGridDBBandedColumn;
    cxLabel39: TcxLabel;
    lblCarrierType: TcxDBLabel;
    cxLabel40: TcxLabel;
    lblCarrierId1: TcxDBLabel;
    cxLabel41: TcxLabel;
    lblCarrierId2: TcxDBLabel;
    btnPrintRecord: TcxButton;
    cxLabel42: TcxLabel;
    txtLoveNum: TcxDBTextEdit;
    cxLabel43: TcxLabel;
    lblRandomNum: TcxDBLabel;
    chkPRIZETYPE: TcxDBCheckBox;
    chkFIXFlag: TcxDBCheckBox;
    txtInvUse: TDataLookup;
    shp1: TShape;
    cxLabel44: TcxLabel;
    cxLabel45: TcxLabel;
    lblA_CarrierId1: TcxDBLabel;
    lblA_CarrierId2: TcxDBLabel;
    cxLabel46: TcxLabel;
    lblDestroyUploadTime: TcxDBLabel;
    cxLabel47: TcxLabel;
    lblDestroyReason: TcxDBLabel;
    chkDestroyFlag: TcxDBCheckBox;
    chkAutoUploadFlag: TcxDBCheckBox;
    btnUploadRecord: TcxButton;
    cdsUploadRecord: TClientDataSet;
    cxDBCheckBox1: TcxDBCheckBox;
    cxLabel48: TcxLabel;
    lblDataType: TcxDBLabel;
    DetailGridViewCreditcard4d: TcxGridDBBandedColumn;
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
    procedure txtGiveUnitCodeNamePropertiesInitPopup(Sender: TObject);
    procedure btnLogClick(Sender: TObject);
    procedure INV008ViewCellDblClick(Sender: TcxCustomGridTableView;
      ACellViewInfo: TcxGridTableDataCellViewInfo; AButton: TMouseButton;
      AShift: TShiftState; var AHandled: Boolean);
    procedure btnPrintRecordClick(Sender: TObject);

    procedure txtInvUse2CodeNamePropertiesInitPopup(Sender: TObject);
    procedure txtInvUseExit(Sender: TObject);
    function IsGiveCanEdit : Boolean;
    procedure btnUploadRecordClick(Sender: TObject);

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
    { 用來記錄最後一張發票號碼, 只用來存檔後, 直接跳到最後一筆用 }
    FMainInvId: String;
    { 用來記錄發票日期, 若 User 改用手打, 必須判斷日期的月份 }
    FChooseInv: TChooseInvInfo;
    // 是否是多張自動換開的發票
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
    function RanDomNum : String;
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
    procedure PrepareUploadRecordDataSet;
    procedure FillDetailDataSet;
    procedure FillInv008ADataSet;
    procedure FillUploadRecordDataSet;
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
    function SyncInv008ADataSet(const aMode: TDMLMode; const aOldSeq,
         aNewSeq, aNewInvId, aItemIdRef: String; var aMsg: String): Boolean;
    function GetItemIdRef(const aItemId: String): String;
    function GetItemDesc(const aItemId: String):String;
    function ChkCreateNum(aInvCnt :Integer):Boolean;
    function ChkChgInvNum(aCount :Integer):Boolean;
    function CalMergeNum():Integer;
    function UpdMergeInv008A(const aItemIdRef, aInvId: String;const aSEQ : Integer) :Boolean;
    function ChkHasMerge(const aRcdNo : Integer; const aItemIdRef : string) : Boolean;
    function ChkCycleMerge(const aItemIdRef, aNewSeq : String;
     aRecNo :Integer ) : Boolean;
    procedure OnInvUseSelectChange(sender : TObject );
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
     frmSearchCustU, frmInvA02_6U,frmInvA02_7U,
     StdConvs,frmInvA02_9U,frmInvA02_AU, frmInvA02_BU;

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

// 將 HowToCreate 碼, 轉成 SO 收費 Table 的代碼 
// SO 收費 Table 的代碼   0:後開   1:預開  2:現場開立

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
  txtInvUse.OnSelectChange := OnInvUseSelectChange;
  {
  txtInvUse2.Connection := dtmMain.InvConnection;
  txtInvUse2.CodeNoFieldName := 'ITEMID';
  txtInvUse2.DescriptionFieldName := 'DESCRIPTION';
  txtInvUse2.Initializa;
  txtInvUse2.OnSelectChange := OnInvUse;
  }
  txtGiveUnit.Connection := dtmMain.InvConnection;
  txtGiveUnit.CodeNoFieldName := 'ITEMID';
  txtGiveUnit.DescriptionFieldName := 'DESCRIPTION';
  txtGiveUnit.Initializa;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.FormShow(Sender: TObject);
var
  aEvent: TDataSetNotifyEvent;
  aMode: TDMLMode;
begin
  if ( FCallFromProgram = TfrmMain ) then
  begin
    Caption := frmMain.GetFormTitleString( 'A0C', '發票查詢' );
    aMode := dmLock;
  end else
  begin
    Caption := frmMain.GetFormTitleString( 'A02', '開立發票' );
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
      PrepareUploadRecordDataSet;
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
    CanClose := ConfirmMsg( '結束開立發票作業?' );
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
    cdsUploadRecord.EmptyDataSet;
    cdsSO.Close;
    txtInvUse.Finaliza;
    //txtInvUse2.Finaliza;
    txtGiveUnit.Finaliza;
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

  {#5945 增加上傳檔案時間 By Kin 2011/03/16}
  {#6262 增加作廢上傳時間與作廢註記 By Kin 2012/07/09}
  {#6337 增加發票開立時間 By Kin 2012/11/13}
  {#6371 增加客戶載具類別號碼、客戶載具顯碼、客戶載具隱碼 By Kin 2013/02/18}
  {#6600 增加愛心碼與防偽隨機碼 By Kin 2013/10/01}
  {#6629 增加中獎註記 By Kin 2013/10/30}
  {#6629 增加字軌屬性 By Kin 2013/11/01}
  {#6764 增加DepositMK與DataType 欄位 By Kin 2014/05/06}
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
    '                 ''1'', ''電子'',            ' +
    '                 ''2'', ''手二'',            ' +
    '                 ''3'', ''手三'' ) INVFORMATDESC, ' +
    '         DECODE( A.ISOBSOLETE,               ' +
    '                ''Y'', ''是'',               ' +
    '                ''否'' ) ISOBSOLETEDESC,     ' +
    '         A.OBSOLETEREASON,                   ' +
    '         A.CANMODIFY,                        ' +
    '         DECODE( A.CANMODIFY,                ' +
    '                 ''Y'', ''是'',              ' +
    '                 ''否'' ) CANMODIFYDESC,     ' +
    '         A.HOWTOCREATE,                      ' +
    '         DECODE( A.HOWTOCREATE,              ' +
    '                 ''1'', ''預開'',            ' +
    '                 ''2'', ''後開'',            ' +
    '                 ''3'', ''現場開立'',        ' +
    '                 ''4'', ''一般開立'' ) HOWTOCREATEDESC, ' +
    '         NVL( A.PRINTCOUNT, 0 ) PRINTCOUNT,  ' +
    '         A.MEMO1,                            ' +
    '         A.MEMO2,                            ' +
    '         A.MAININVID,                        ' +
    '         A.MAINSALEAMOUNT,                   ' +
    '         A.MAINTAXAMOUNT,                    ' +
    '         A.MAININVAMOUNT,                    ' +
    '         A.INSTADDR,                         ' +
    '         A.INVUSEID,                         ' +
    '         A.INVUSEDESC,                       ' +
    '         A.INVOICEKIND,                      ' +
    '         A.GIVEUNITID,                       ' +
    '         A.GIVEUNITDESC,                     ' +
    '         A.CONTMOBILE,                       ' +
    '         A.EMAIL,                            ' +
    '         NVL(A.UPLOADFLAG,0) UPLOADFLAG,     ' +
    '         A.PRINTEN,                          ' +
    '         A.PRINTTIME,                        ' +
    '         A.UPLOADTIME,                       ' +
    '         A.OBUPLOADFLAG,                     ' +
    '         A.OBUPLOADTIME,                     ' +
    '         A.CREATEINVDATE,                    ' +
    '         A.CARRIERTYPE,                      ' +
    '         A.CARRIERID1,                       ' +
    '         A.CARRIERID2,                       ' +
    '         A.RANDOMNUM,                        ' +
    '         A.LOVENUM,                          ' +
    '         NVL(A.FIXFLAG,0) FIXFLAG,           ' +
    '         DECODE(A.PRIZETYPE,NULL,0,1) PRIZETYPE, ' +
    '         NVL(A.DESTROYFLAG,0 ) DESTROYFLAG,  ' +
    '         A.DESTROYUPLOADTIME,                ' +
    '         A.DESTROYREASON,                    ' +
    '         A.A_CarrierId1,                     ' +
    '         A.A_CarrierId2,                     ' +
    '         Nvl(A.AUTOUPLOADFLAG,0) AUTOUPLOADFLAG, ' +
    '         NVL(A.DEPOSITMK,''N'') DEPOSITMK ,  ' +
    '         ( CASE UPPER(A.DATATYPE)            ' +
    '         WHEN ''X'' then ''X-社福團體''      ' +
    '         WHEN ''A'' then ''A-會員廠商''      ' +
    '         WHEN ''B'' then ''B-非會員廠商''    ' +
    '         WHEN ''Y'' then ''Y-無載具紙本''    ' +
    '         ELSE NULL                           ' +
    '         END ) DATATYPE                      ' +
    '    FROM INV007 A                            ' +
    '   WHERE A.IDENTIFYID1 = ''%s''              ' +
    '     AND A.IDENTIFYID2 = %s                  ' +
    '     AND A.INVID = ''%s''                    ' +
    '     AND A.COMPID = ''%s''                   ',
    [IDENTIFYID1, IDENTIFYID2, aInvId, dtmMain.getCompID] );
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA02_4.BuildDetailSQL(const aInvId: String): String;
begin
  {#5439 增加收費資料的合併收費項目 By Kin }
  {#5800 增加CMMAC By Kin 2012/07/02}
  {#7076 增加信用卡末4碼 By Kin 2015/08/20}
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
    '           ''Y'', ''是'', ''否'' ) AS LINKTOMISDESC, ' +
    '         B.TAXCODE,                           ' +
    '         B.TAXNAME,                           ' +
    '         B.SIGN,                              ' +
    '         '' '' AS SOURCE,                     ' +
    '         DECODE( A.INVID, NULL,               ' +
    '                          ''Y'',              ' +
    '                          ''N'' ) EMPTY,      ' +
    '         A.FACISNO,                           ' +
    '         A.ACCOUNTNO,                         ' +
    '         A.SMARTCARDNO,                       ' +
    '         A.COMBCITEMCODE,                     ' +
    '         A.COMBCITEMNAME,                     ' +
    '         A.CMMAC,                             ' +
    '         A.Creditcard4d                       ' +
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
  {#5800 增加CMMac By Kin 2012/07/02}
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
    '        A.ACCOUNTNO,             ' +
    '        A.SMARTCARDNO,           ' +
    '        A.CMMAC                  ' +
    '   FROM INV008A A                ' +
    '  WHERE A.INVID = ''%s''         ' +
    '  ORDER By A.SEQ                 ', [aInvId] );
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA02_4.BuildInv008AApplySQL(const aMode: TDMLMode): String;
begin
  {#5668 增加SmartCardNo By Kin 2010/07/09}
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
      '         ACCOUNTNO,          ' +
      '         SMARTCARDNO )       ' +
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
      '         :ACCOUNTNO,         ' +
      '         :SMARTCARDNO )      ';
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA02_4.BuildMasterApplySQL(const aMode: TDMLMode): String;
begin
  {#5668 增加發票種類 By Kin 2010/06/29 }
  {#6337 增加發票開立時間 By Kin 2012/11/13}
  {#6371 增加客戶載具CarrierType、CarrierId1、CarrierId2 By Kin 2013/02/25}
  {#6600 增加隨機碼與愛心碼 By Kin 2013/10/01}
  {#6629 增加字軌屬性 By Kin 2013/11/01}
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
      '          INVUSEDESC,           ' +
      '          INVOICEKIND,          ' +
      '          GIVEUNITID,           ' +
      '          GIVEUNITDESC,         ' +
      '          EMAIL,                ' +
      '          CONTMOBILE,           ' +
      '          CREATEINVDATE,        ' +
      '          CARRIERTYPE,          ' +
      '          CARRIERID1,           ' +
      '          CARRIERID2,           ' +
      '          RANDOMNUM,            ' +
      '          LOVENUM,              ' +
      '          FIXFLAG,              ' +
      '          A_CARRIERID1,         ' +
      '          A_CARRIERID2 )        ' +
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
      '          :INVUSEDESC,          ' +
      '          :INVOICEKIND,         ' +
      '          :GIVEUNITID,          ' +
      '          :GIVEUNITDESC,        ' +
      '          :EMAIL,               ' +
      '          :CONTMOBILE,          ' +
      '          SYSDATE,              ' +
      '          :CARRIERTYPE,         ' +
      '          :CARRIERID1,          ' +
      '          :CARRIERID2,          ' +
      '          :RANDOMNUM,           ' +
      '          :LOVENUM,             ' +
      '          :FIXFLAG,             ' +
      '          :A_CARRIERID1,        ' +
      '          :A_CARRIERID2   )     '

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
      '        INVUSEDESC = :INVUSEDESC,         ' +
      '        INVOICEKIND = :INVOICEKIND,       ' +
      '        GIVEUNITID = :GIVEUNITID,         ' +
      '        GIVEUNITDESC = :GIVEUNITDESC,     ' +
      '        EMAIL = :EMAIL,                   ' +
      '        CONTMOBILE = :CONTMOBILE,         ' +
      '        CARRIERTYPE = :CARRIERTYPE,       ' +
      '        CARRIERID1 = :CARRIERID1,         ' +
      '        CARRIERID2 = :CARRIERID2,         ' +
      '        RANDOMNUM = :RANDOMNUM,           ' +
      '        LOVENUM = :LOVENUM,               ' +
      '        FIXFLAG = :FIXFLAG                ' +
      '  WHERE IDENTIFYID1 = :IDENTIFYID1        ' +
      '    AND IDENTIFYID2 = :IDENTIFYID2        ' +
      '    AND INVID = :INVID                    ';
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA02_4.BuildDetailApplySQL(const aMode: TDMLMode): String;
begin
  {#7076 增加信用卡末4碼 By Kin 2015/08/20}
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
      '           ACCOUNTNO,           ' +
      '           COMBCITEMCODE,       ' +
      '           COMBCITEMNAME,       ' +
      '           SMARTCARDNO,         ' +
      '           CREDITCARD4D)        ' +
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
      '           :ACCOUNTNO,          ' +
      '           :COMBCITEMCODE,      ' +
      '           :COMBCITEMNAME,      ' +
      '           :SMARTCARDNO,        ' +
      '           :CREDITCARD4D  )     ' ;
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
            
        { #5014 在寫的時候順便把Item 的單引號一起拿掉 By Kin 2009/04/03 }
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
  // 一般開立, 客戶資料從發票系統抓, PS: 修改時可改客編
  // 不須檢查 CompId, 只有抓出來 CustId 為一筆資料才帶進來
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
  // 現場開立, 預開, 後開, 客戶資料從 SMS 抓, PS: 修改時不可更改客編
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
  { 取收費項目 }
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
    { 以數量計算 }
    aTax := 0;
    aTaxRate := cdsMaster.FieldByName( 'TAXRATE' ).AsFloat;
    aQty := cdsDetail.FieldByName( 'QUANTITY' ).AsInteger;
    aPrice := cdsDetail.FieldByName( 'UNITPRICE' ).AsFloat;  //#4370 改成浮點數 By Kin 2009/02/19
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
    { 以單價計算 }
    if ( cdsDetail.FieldByName( 'UNITPRICE' ).AsFloat = 0 ) then  //#4370 改成浮點數 By Kin 2009/02/19
    begin
      cdsDetail.FieldByName( 'TAXAMOUNT' ).AsInteger := 0;
      cdsDetail.FieldByName( 'SALEAMOUNT' ).AsInteger := 0;
      cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsInteger := 0;
    end else
    { 若修改單價時, Tax, Sale, Total 直皆有值 }
    if ( cdsDetail.FieldByName( 'TAXAMOUNT' ).AsInteger <> 0 ) and
       ( cdsDetail.FieldByName( 'SALEAMOUNT' ).AsInteger <> 0 ) and
       ( cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsInteger <> 0 ) then
    begin
      aPrice := cdsDetail.FieldByName( 'UNITPRICE' ).AsFloat; //#4370 改成浮點數 By Kin 2009/02/19
      aQty := cdsDetail.FieldByName( 'QUANTITY' ).AsInteger;
      aTotal := cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsInteger;
//      aSale := ( aQty * aPrice );
      aSale := Round( TUother.RoundFloat( aQty * aPrice, 0 ) ); //4370 因為Price已改成浮點數 By Kin 2009/02/19
      aTax := (  aTotal - aSale );
      cdsDetail.FieldByName( 'SALEAMOUNT' ).AsInteger := aSale;
      cdsDetail.FieldByName( 'TAXAMOUNT' ).AsInteger := aTax;
    end else
    begin
      aTax := 0;
      aTaxRate := cdsMaster.FieldByName( 'TAXRATE' ).AsFloat;
      aQty := cdsDetail.FieldByName( 'QUANTITY' ).AsInteger;
      aPrice := cdsDetail.FieldByName( 'UNITPRICE' ).AsFloat; //#4370 改成浮點數 By Kin 2009/02/19
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
    { 以總金額計算 }
    if ( cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsInteger = 0 ) then
    begin
      cdsDetail.FieldByName( 'SALEAMOUNT' ).AsInteger := 0;
      cdsDetail.FieldByName( 'TAXAMOUNT' ).AsInteger := 0;
      cdsDetail.FieldByName( 'UNITPRICE' ).AsInteger := 0;
    end else
    begin
      aQty := cdsDetail.FieldByName( 'QUANTITY' ).AsInteger;
      aPrice := cdsDetail.FieldByName( 'UNITPRICE' ).AsFloat; //#4370 改成浮點數 By Kin 2009/02/19
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
    { 以銷售額計算 }
    if ( cdsDetail.FieldByName( 'SALEAMOUNT' ).AsInteger = 0 ) then
    begin
      cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsFloat := 0;
      cdsDetail.FieldByName( 'TAXAMOUNT' ).AsFloat := 0;
      cdsDetail.FieldByName( 'UNITPRICE' ).AsFloat := 0;
    end else
    { 若修改銷售額時, Tax, Price, Total 直皆有值 }
    if ( cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsInteger <> 0 ) and
       ( cdsDetail.FieldByName( 'TAXAMOUNT' ).AsInteger <> 0 ) and
       ( cdsDetail.FieldByName( 'UNITPRICE' ).AsFloat <> 0 ) then
    begin
      aPrice := cdsDetail.FieldByName( 'UNITPRICE' ).AsFloat; //#4370 改成浮點數 By Kin 2009/02/19
      aTotal := cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsInteger;
      aSale := cdsDetail.FieldByName( 'SALEAMOUNT' ).AsInteger;
      aTax := ( aTotal - aSale );
      cdsDetail.FieldByName( 'TAXAMOUNT' ).AsInteger := aTax;
    end else
    begin
      aQty := cdsDetail.FieldByName( 'QUANTITY' ).AsInteger;
      aPrice := cdsDetail.FieldByName( 'UNITPRICE' ).AsFloat; //#4370 改成浮點數 By Kin 2009/02/19
      aTax := 0;
      aTaxRate := cdsMaster.FieldByName( 'TAXRATE' ).AsFloat;
      aSale := cdsDetail.FieldByName( 'SALEAMOUNT' ).AsInteger;
      if aTaxRate > 0 then
        aTax := Round( TUother.RoundFloat( ( aSale * aTaxRate / 100 ), 0 ) );
      aTotal := aSale + aTax;
      cdsDetail.FieldByName( 'TAXAMOUNT' ).AsInteger := aTax;
      cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsInteger := aTotal;
      aPrice := Round( TUother.RoundFloat( aSale / aQty, 2 ) );
      cdsDetail.FieldByName( 'UNITPRICE' ).AsFloat := aPrice; //#4370 改成浮點數 By Kin 2009/02/19
      UpdateDetailPriceByChargeSign;
    end;
  end else
  if ( aCalculateColum in [chTaxAmount] ) then
  begin
    { 若修改稅額時, Sale, Price, Total 直皆有值 }
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
  { 取收費項目的正負項 }
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
//  cdsDetail.BeforePost := nil; //cdsDetailBeforePost; 已不須要
  cdsDetail.BeforePost :=cdsDetailBeforePost;         
  cdsDetail.AfterPost := cdsDetailAfterPost;
  cdsDetail.BeforeInsert := cdsDetailBeforeInsert;
  cdsDetail.OnNewRecord := cdsDetailNewRecord;
  cdsDetail.BeforeDelete := cdsDetailBeforeDelete;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.PrepareDetailDataSet;
begin
  {#5668 增加SMARTCARDNO By Kin 2010/07/09}
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
    //#4370 單價要改成浮點數 By Kin 2009/02/19
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
    {#5668 增加SMARTCARDNO By Kin 2010/07/09}
    cdsDetail.FieldDefs.Add( 'SMARTCARDNO', ftString, 20);
    {#5800 增加CMMAC By Kin 2012/07/02}
    cdsDetail.FieldDefs.Add( 'CMMAC' ,ftString,20 );
    {#7076 增加信用卡末4碼 By Kin 2015/08/20}
    cdsDetail.FieldDefs.Add( 'CREDITCARD4D' ,ftString,10 );
    {#5349 增加收費資料的合併項目名稱}
    cdsDetail.FieldDefs.Add( 'COMBCITEMCODE',ftInteger );
    cdsDetail.FieldDefs.Add( 'COMBCITEMNAME',ftString,40 );
    cdsDetail.FieldDefs.Add( 'ITEMIDREF',ftInteger );
    cdsDetail.FieldDefs.Add( 'MERGESEQ' ,ftInteger );
    cdsDetail.CreateDataSet;
    //cdsDetail.AddIndex('SIGN','SIGN',[],'SIGN');
    //cdsDetail.IndexFieldNames := 'SIGN';
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
    cdsInv008A.FieldDefs.Add( 'SMARTCARDNO' , ftString, 20);
    cdsInv008A.FieldDefs.Add( 'CMMAC', ftString, 20);
    cdsInv008A.CreateDataSet;
  end;
  cdsInv008A.EmptyDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.FillDetailDataSet;
var
  aOldReadOnly: Boolean;
  aStartDate, aEndDate: String;
  aItemIdRef, aMergeSeq: string;
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
      { 匯檔時 日期格式可能是錯的, 會導至錯誤 }
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
      //#5800 增加CMMAc By Kin 2012/07/02
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
        dtmMainH.adoA02_Detail.FieldByName( 'ACCOUNTNO' ).AsString,
        dtmMainH.adoA02_Detail.FieldByName( 'SMARTCARDNO' ).AsString,
        dtmMainH.adoA02_Detail.FieldByName( 'CMMAC' ).AsString,
        dtmMainH.adoA02_Detail.FieldByName( 'CREDITCARD4D' ).AsString ] );
      dtmMainH.adoA02_Detail.Next;
    end;
    dtmMainH.adoA02_Detail.Close;
    aItemIdRef := EmptyStr;
    aMergeSeq := EmptyStr;

    if ( cdsDetail.State in [ dsBrowse ] ) and ( cdsInv008A.State in [ dsBrowse ] ) then
    begin
      if ( not cdsDetail.IsEmpty ) then
      begin
        cdsDetail.First;
        while not cdsDetail.Eof do
        begin
          if cdsInv008A.Locate('INVID;SEQ',VarArrayOf([
            cdsDetail.FieldByName( 'INVID' ).AsString,
            cdsDetail.FieldByName( 'SEQ').AsString ]),[]) then
          begin

            aItemIdRef := cdsInv008A.FieldByName( 'ITEMIDREF' ).AsString;
            aMergeSeq := cdsInv008A.FieldByName( 'SEQ' ).AsString;
            cdsDetail.Edit;
            cdsDetail.FieldByName( 'ITEMIDREF' ).AsString := aItemIdRef;
            cdsDetail.FieldByName( 'MERGESEQ' ).AsString := aMergeSeq;
            cdsDetail.Post;
          end;
          cdsDetail.Next;
        end;
      end;
    end;
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
  aReader := dtmMain.adoComm2;
  aReader.Close;
  aReader.SQL.Text := BuildINV008ASQL( FInvoiceId );
  aReader.Open;
  aReader.First;
  {#5668 增加SMARTCARDNO By Kin 2010/07/09}
  while not aReader.Eof do
  begin
    { 匯檔時 日期格式可能是錯的, 會導至錯誤 }
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
      aReader.FieldByName( 'ACCOUNTNO' ).AsString,
      aReader.FieldByName( 'SMARTCARDNO' ).AsString,
      aReader.FieldByName( 'CMMAC' ).AsString ] );
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
  aDataSet,aDataSet2: TClientDataSet;
  aChargeSign, aAccountNo, aFacisNo,aSmartCard: String;
  aRay: Double;
begin
  if ( dtmMain.GetStarEInvoice ) and ( chkEInv.Visible ) then
  begin
    aDataSet2 := TfrmInvA02_1( FChargeWindow ).TitleDataSet;
    if aDataSet2.FieldDefs.Count > 0 then
    begin
      if aDataSet2.FieldByName( 'INVOICEKIND' ).AsString = '電子發票' then
        chkEInv.Checked := True
      else
        chkEInv.Checked := False;
      if aDataSet2.FieldByName( 'EMAIL' ).AsString <> '' then
      begin
        cdsMaster.FieldByName( 'EMAIL' ).AsString :=
          aDataSet2.FieldByName( 'EMAIL').AsString;
      end;
      if aDataSet2.FieldByName( 'TEL3' ).AsString <> EmptyStr then
      begin
        cdsMaster.FieldByName( 'CONTMOBILE' ).AsString :=
          aDataSet2.FieldByName( 'TEL3' ).AsString;
      end;
      cdsMaster.FieldByName( 'GIVEUNITID' ).AsString :=
        aDataSet2.FieldByName( 'DENRECCODE' ).AsString;
      cdsMaster.FieldByName( 'GIVEUNITDESC' ).AsString :=
        aDataSet2.FieldByName( 'DENRECNAME' ).AsString;
    end;

  end;

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
          { #5057 如果挑選的是SO033,先判斷是否有實收日期,如果有則抓REALAMT By Kin 2009/05/05 }
          {#5285 判斷實收日期改成判斷未收原因 By Kin 2009/09/09 }
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

          { 應稅 }
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
            
          { #5349 增加收費資料的合併收費項目 By Kin }
          cdsDetail.FieldByName( 'COMBCITEMCODE' ).AsString :=
            aDataSet.FieldByName( 'COMBCITEMCODE' ).AsString;

          cdsDetail.FieldByName( 'COMBCITEMNAME' ).AsString :=
            aDataSet.FieldByName( 'COMBCITEMNAME' ).AsString;
          {#7076}
          cdsDetail.FieldByName( 'Creditcard4d' ).AsString :=
            aDataSet.FieldByName( 'CardLastNo' ).AsString;
          { 負項收費 }
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

          { 非合併進來的明細, 取門號及信用卡號 }
          if not IsMergeRecord( FInvoiceId,
            cdsDetail.FieldByName( 'SEQ' ).AsString ) then
          begin
            {#5668 增加存入 SmartCard By Kin 2010/07/09}
            dtmSO.GetMisAccountAndFacisNo(
              cdsDetail.FieldByName( 'BILLID' ).AsString,
              cdsDetail.FieldByName( 'BILLIDITEMNO' ).AsString,
              txtCUSTID.Text, aAccountNo,aFacisNo,aSmartCard );
            if ( aFacisNo <> EmptyStr ) then
              cdsDetail.FieldByName( 'FACISNO' ).AsString := aFacisNo;
            if ( aAccountNo <> EmptyStr ) then
              cdsDetail.FieldByName( 'ACCOUNTNO' ).AsString := aAccountNo;
            if ( aSmartCard <> EmptyStr ) then
              cdsDetail.FieldByName( 'SMARTCARDNO' ).AsString := aSmartCard;
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
    //#5668 增加判斷新增欄位 By Kin 2010/07/16
    if dtmMain.GetStarEInvoice then
    begin
      if aDataSet.FieldByName( 'INVOICEKIND' ).AsString = '電子發票' then
        chkEInv.Checked := True
      else
        chkEInv.Checked := False;
    end;
    if aDataSet.FieldByName( 'TEL3' ).AsString <> EmptyStr then
    begin
      cdsMaster.FieldByName( 'CONTMOBILE' ).AsString :=
        aDataSet.FieldByName( 'TEL3' ).AsString;
    end;
    if aDataSet.FieldByName( 'EMAIL' ).AsString <> EmptyStr then
    begin
      cdsMaster.FieldByName( 'EMAIL' ).AsString :=
        aDataSet.FieldByName( 'EMAIL' ).AsString;
    end;

    cdsMaster.FieldByName( 'GIVEUNITID' ).AsString :=
      aDataSet.FieldByName( 'DENRECCODE' ).AsString;
    cdsMaster.FieldByName( 'GIVEUNITDESC' ).AsString :=
      aDataSet.FieldByName( 'DENRECNAME' ).AsString;
    {#6371 增加客戶載具資料 By Kin 2013/02/18}
    
    cdsMaster.FieldByName( 'CARRIERTYPE' ).AsString :=
      TfrmInvA02_1( FChargeWindow ).CarrierType;
    cdsMaster.FieldByName( 'CARRIERID1' ).AsString :=
      TfrmInvA02_1( FChargeWindow ).CarrierId1;
    cdsMaster.FieldByName( 'CARRIERID2' ).AsString :=
      TfrmInvA02_1( FChargeWindow ).CarrierId2;
    {#6600 增加愛心碼 By Kin 2013/10/02}

    cdsMaster.FieldByName( 'LoveNum' ).AsString :=
      TfrmInvA02_1( FChargeWindow ).LoveNum;

    cdsMaster.FieldByName( 'A_CARRIERID1' ).AsString :=
      TfrmInvA02_1( FChargeWindow ).A_CarrierId1;

    cdsMaster.FieldByName( 'A_CARRIERID2' ).AsString :=
      TfrmInvA02_1( FChargeWindow ).A_CarrierId2;

    {
    if aDataSet.FieldByName( 'DENRECCODE' ).AsString <> EmptyStr then
    begin
      txtGiveUnit.Value := aDataSet.FieldByName( 'DENRECCODE' ).AsString;
      txtGiveUnit.ValueName := aDataSet.FieldByName( 'DENRECNAME' ).AsString;

    end;
    }
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
        DetailGridLevel2.Visible := True;
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
        btnLog.Enabled := True;
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
        btnLog.Enabled := False;
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
        btnLog.Enabled := False;
      end;
  end;
  { 非查詢模式 }
  if not ( aMode in [dmLock] ) then
  begin
    { 特殊規則 1 }
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
    { 特殊規則 2 }
    if not dtmMain.GetLinkToMIS then
      btnLocaleCreate.Enabled := False;

    { 特殊規則 3}
    if ( not dtmMain.GetUseCOM ) or ( cdsMaster.RecordCount = 0 ) then
      btnLog.Enabled := False;
  end;    
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.EditorStateChange(const aMode: TDMLMode);
var
  blnLoveCanEdit : Boolean;
begin
  {#5668 增加電子發票種類 By Kin 2010/07/08}
  if not dtmMain.GetStarEInvoice then
  begin
    chkEInv.Checked := False;
    grpInvoiceKind.Visible := False;
  end else
  begin
    grpInvoiceKind.Visible := True;
  end;
  if ( FInvCreateType in [icNormal] ) then
  begin
//    chkEInv.Checked := False;
    //grpInvoiceKind.Visible := False;
  end;

  if ( aMode in [dmBrowser, dmLock] ) then
  begin
    txtINVDATE.Enabled := True;
    txtINVDATE.Properties.PopupControl := nil;
    grpTaxRate.Enabled := False;
    txtTAXRATE.Enabled := True;
    txtINVID.Properties.ReadOnly := True;
    txtMainInvId.Properties.ReadOnly := True;
    txtInvUse.Enabled := False;

    txtGiveUnit.Enabled := False;
    if dtmMain.GetStarEInvoice then
    begin
      grpInvoiceKind.Visible := True;
    end;
    grpInvoiceKind.Enabled := False;
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

//    txtGiveUnit.Enabled := True;
    if dtmMain.GetStarEInvoice then
    begin
      grpInvoiceKind.Enabled := True;
      chkEInv.Enabled := True;
    end;
//    txtGiveUnit.Enabled := ( FInvCreateType in [icLocale,icNormal] );
  end
  else if ( aMode in [dmUpdate] ) then
  begin
    { 發票日期可以修改, 不可重新選擇發票本, 條件為以下格式,
      手工二聯 = InvFormat = 2 , 手工三聯 = InvFormat = 3  }
    txtINVDATE.Enabled := False;
    if ( cdsMaster.FieldByName( 'INVFORMAT' ).AsString = '2' ) or
       ( cdsMaster.FieldByName( 'INVFORMAT' ).AsString = '3' ) then
    begin
      txtINVDATE.Enabled := True;
      txtINVDATE.Properties.PopupControl := nil;
      {#5668 如果是手開,不要顯示電子發票選項 By Kin 2010/07/08}
      chkEInv.Checked := False;
      //#5922 InvFormat不是1,不可以選擇 By Kin 2011/02/11
      grpInvoiceKind.Enabled := False;
      //grpInvoiceKind.Visible := False;
    end;
    // 1. 客編: 一般開立的發票才可修改, 但是非一般開立的發票的彈出視窗選資料功能,
    //           仍然要能使用
    // 2. 稅別: 一般開立的發票才可修改(如果是多張自動開立的不可以修改)
    if dtmMain.GetStarEInvoice then
    begin
      grpInvoiceKind.Visible := True;
      if cdsMaster.FieldByName( 'INVFORMAT' ).AsString = '1' then
      begin
        grpInvoiceKind.Enabled := True;
      end;
    end;
    if FInvCreateType = icNormal then
    begin
      txtCUSTID.Properties.ReadOnly := False;
      txtCUSTID.Style.TextColor :=
        txtCUSTID.Style.StyleController.Style.TextColor;
      grpTaxRate.Enabled := not FMultiInv;
      txtTAXRATE.Enabled := not FMultiInv;
      {#5668 如果是手開,不要顯示電子發票選項 By Kin 2010/07/08}
//      chkEInv.Checked := False;
      //grpInvoiceKind.Visible := False;
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
    // 現場開立或一般開立, 才可以修改發票用途欄位, 拋檔過來的則不行
    txtInvUse.Enabled := ( FInvCreateType in [icLocale, icNormal] );
    if cdsMaster.FieldByName( 'InvUseId' ).AsString = '' then
    begin
      txtGiveUnit.Enabled := False;
    end;
//    txtGiveUnit.Enabled := ( FInvCreateType in [icLocale,icNormal] );
    if ( cdsMaster.FieldByName( 'UPLOADTIME' ).AsString <> '' ) Or
          ( cdsMaster.FieldByName( 'OBUPLOADTIME' ).AsString <> '' ) then
    begin
      txtInvUse.Enabled := False;
      txtGiveUnit.Enabled := False;
      txtLoveNum.Properties.ReadOnly := True;
      //txtLoveNum.Enabled := False;
    end;

  end;

  if ( cdsMaster.FieldByName('ObUploadFlag').AsInteger = 1 ) or
    (cdsMaster.FieldByName( 'UploadFlag' ).AsInteger = 1) then
  begin
    txtInvUse.Enabled := False;
    txtGiveUnit.Enabled := False;
    txtLoveNum.Enabled := False;
  end;
  if ( cdsMaster.FieldByName('DestroyUploadTime').AsString <> '' ) and
      ( not ( aMode in [dmBrowser, dmLock]))      then
  begin
    if (cdsMaster.FieldByName( 'DestroyUploadTime' ).AsDateTime > cdsMaster.FieldByName( 'UploadTime' ).AsDateTime) then
    begin
      if ( cdsMaster.FieldByName( 'DestroyUploadTime' ).AsDateTime > cdsMaster.FieldByName( 'PrintTime').AsDateTime ) or
          (cdsMaster.FieldByName( 'PrintTime' ).AsString = '') then
      begin
        blnLoveCanEdit := IsGiveCanEdit;
        txtInvUse.Enabled := True;
        txtLoveNum.Properties.ReadOnly := not blnLoveCanEdit;
        txtLoveNum.Enabled := blnLoveCanEdit;
        txtGiveUnit.Enabled := blnLoveCanEdit;
      end;

    end;
  end;
  { 設定 Focus }
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
//        txtLoveNum.Enabled := False;
        txtLoveNum.Properties.ReadOnly := True;
        txtGiveUnit.Enabled := False;
      end;
    dmUpdate:
      begin
        cdsMaster.ReadOnly := False;
        cdsDetail.ReadOnly := False;
//        cdsInv008A.ReadOnly := False;
        cdsMaster.Edit;
//        cdsInv008A.Edit;
//        cdsInv008A.FieldByName('QUANTITY').AsString :='10';
//        cdsInv008A.Post;
      end;
    dmSave:
      begin
        { 先檢查 Detail }
        Result := ValidateDetailDataSet( aErrMsg );
        if not Result then
        begin
          WarningMsg( aErrMsg );
          Exit;
        end;
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
        { 重新 Open DataSet }
        FInvoiceId := cdsMaster.FieldByName( 'INVID' ).AsString;
        Self.FormShow( Self );
      end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA02_4.ValidateCanModify(var aErrMsg: String): Boolean;
  var blnLoveCanEdit : Boolean;
begin
  aErrMsg := EmptyStr;
  Result := True;
  blnLoveCanEdit := False;
  try
    if cdsMaster.IsEmpty then
    begin
      aErrMsg := '無發票資料, 不可修改。';
      Exit;
    end;
    if cdsMaster.FieldByName( 'ISOBSOLETEDESC' ).AsString = '是' then
    begin
      aErrMsg := '此筆發票已作廢!';
      Exit;
    end;
    if cdsMaster.FieldByName( 'CANMODIFY' ).AsString <> 'Y' then
    begin
      aErrMsg := '此筆發票需做過授權才可修改!';
      Exit;
    end;
    {
    if ( cdsMaster.FieldByName( 'UPLOADFLAG' ).AsString = '1' ) and
        ( cdsMaster.FieldByName( 'INVOICEKIND' ).AsString = '1' ) then
    begin
      aErrMsg := '不可修改 [已產生上傳檔] 的發票';
    end;
    }


{ 此段有問題, 發票是否鎖帳, 檢查 INV099.USEFUL 是錯的, 應該檢查 INV018 }
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
      aErrMsg := '此筆發票已過期(該月份已鎖帳或發票號碼不在有效發票本內),不可修改。';
      Exit;
    end;

    //#5874 已上傳而且是電子發票,不允許修改下列欄位 By Kin 2010/12/15
    {6716 與Jacy討論後,將InvoiceKind=1的條件拿掉}
    if ( cdsMaster.FieldByName( 'UPLOADFLAG' ).AsString = '1' ) then
    begin
      txtBUSINESSID.Properties.ReadOnly := True;
      txtCHARGEDATE.Properties.ReadOnly := True;
      chkEInv.Properties.ReadOnly := True;
      cdsDetail.ReadOnly :=True;
      {#6600 如果已上傳不可再修改愛心碼 By Kin 2013/10/01}
      txtLoveNum.Properties.ReadOnly := True;
      DetailGridView.OptionsData.Editing := False;
      DetailGridView.OptionsData.Appending := False;
      DetailGridView.OptionsData.Inserting := False;
      DetailGridView.OptionsData.Deleting := False;
    end;
    {#6600 測試不OK 已上傳後(有上傳時間), 發票用途, 受贈單位, 愛心碼不可再被修改 By Kin 2013/11/12}
    if ( cdsMaster.FieldByName( 'UPLOADTIME' ).AsString <> EmptyStr ) then
    begin
      txtInvUse.Enabled := False;

      txtGiveUnit.Enabled := False;
//      txtLoveNum.Enabled := False;
      txtLoveNum.Properties.ReadOnly := True;
    end else
    begin
      {#6600 測試不OK一般開立"及"現場開立", 發票用途為"捐贈"類的參考號, 才可輸入愛心碼或受贈單位 By Kin 2013/11/12}
      blnLoveCanEdit := IsGiveCanEdit;
//      txtLoveNum.Enabled := blnLoveCanEdit;
      txtLoveNum.Properties.ReadOnly := not blnLoveCanEdit;
      txtGiveUnit.Enabled := blnLoveCanEdit;
    end;
    {#6716 若註銷上傳時間>開立上傳時間，且(註銷上傳時間>列印時間 或列印時間為NULL)，該張發票就屬於可修改的狀態 By Kin 2014/02/07}
    if ( cdsMaster.FieldByName('DestroyUploadTime').AsString <> '' ) then
    begin
      if (cdsMaster.FieldByName( 'DestroyUploadTime' ).AsDateTime > cdsMaster.FieldByName( 'UploadTime' ).AsDateTime) then
      begin
        if ( cdsMaster.FieldByName( 'DestroyUploadTime' ).AsDateTime > cdsMaster.FieldByName( 'PrintTime').AsDateTime ) or
            (cdsMaster.FieldByName( 'PrintTime' ).AsString = '') then
        begin
          txtBUSINESSID.Properties.ReadOnly := False ;
          txtCHARGEDATE.Properties.ReadOnly := False;
          chkEInv.Properties.ReadOnly := False;
          cdsDetail.ReadOnly := False;
          txtLoveNum.Properties.ReadOnly := False;
          DetailGridView.OptionsData.Editing := True;
          DetailGridView.OptionsData.Appending := True;
          DetailGridView.OptionsData.Inserting := True;
          DetailGridView.OptionsData.Deleting := True;
          {
          blnLoveCanEdit := IsGiveCanEdit;
          txtLoveNum.Properties.ReadOnly := not blnLoveCanEdit;
          txtGiveUnit.Enabled := blnLoveCanEdit;
          }
        end;

      end;
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
      aErrMsg := '無發票資料, 不可做作廢。';
      Exit;
    end;
    if ( cdsMaster.FieldByName( 'ISOBSOLETEDESC' ).AsString = '是' ) then
    begin
      aErrMsg := '此筆發票已作廢。';
      Exit;
    end;
    if ( cdsMaster.FieldByName( 'CANMODIFY' ).AsString <> 'Y' ) then
    begin
      aErrMsg := '此筆發票需做過授權才可作廢。';
      Exit;
    end;
    {#6660 將USEFUL 條件拿掉 By Kin 2013/11/15}
    aPrefix := Copy( cdsMaster.FieldByName( 'INVID' ).AsString, 1, 2 );
    aNum := Copy( cdsMaster.FieldByName( 'INVID' ).AsString, 3, 8 );
    dtmMain.adoComm.Close;
    dtmMain.adoComm.SQL.Text := Format(
      ' SELECT CURNUM FROM INV099  ' +
      '  WHERE PREFIX = ''%s''     ' +
      '    AND ENDNUM >= ''%s''    ' +
      '    AND STARTNUM <= ''%s''  ',
      [ aPrefix, aNum, aNum ] );
    dtmMain.adoComm.Open;
    if ( dtmMain.adoComm.IsEmpty ) then
    begin
      aErrMsg := '此筆發票已過期(該月份已鎖帳或發票號碼不在有效發票本內), 不可作廢。';
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
      aErrMsg := '無發票資料, 不可套印。';
      Exit;
    end;
    if cdsMaster.FieldByName( 'INVFORMAT' ).AsString <> '1' then
    begin
      aErrMsg := '此筆發票格式必須為【電子】, 才可套印此筆發票。';
      Exit;
    end;
    if dtmMain.GetIfPrintCheck then
    begin
      if cdsMaster.FieldByName( 'CANMODIFY' ).AsString <> 'Y' then
      begin
        aErrMsg := '此筆發票需做過授權才可套印!';
        Exit;
      end;
    end;  
    if cdsMaster.FieldByName( 'ISOBSOLETEDESC' ).AsString = '是' then
    begin
      aErrMsg := '此筆發票已做廢, 不可套印!';
      Exit;
    end;
    {#6764 DEPOSITMK=Y 不可再列印 By Kin 2014/05/06 }
    if cdsMaster.FieldByName( 'DEPOSITMK' ).AsString = 'Y' then
    begin
      aErrMsg := '大平台已匯款註記，不可再次列印發票！';
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
  aDataSet : TADOQuery;
begin
  Result := False;
  aFocusControl := nil;
  try
    if ( cdsMaster.FieldByName( 'INVID' ).AsString = EmptyStr ) then
    begin
      aErrMsg := '請選擇發票號碼。';
      aFocusControl := txtINVDATE;
      Exit;
    end;
    if ( cdsMaster.FieldByName( 'INVDATE' ).AsString = EmptyStr ) then
    begin
      aErrMsg := '請輸入發票日期。';
      aFocusControl := txtINVDATE;
      Exit;
    end;
    if ( cdsMaster.FieldByName( 'CUSTID' ).AsString = EmptyStr ) then
    begin
      aErrMsg := '請輸入客編。';
      aFocusControl := txtCUSTID;
      Exit;
    end;
    if ( cdsMaster.FieldByName( 'CUSTSNAME' ).AsString = EmptyStr ) then
    begin
      aErrMsg := '請輸入客戶名稱。';
      aFocusControl := txtCUSTSNAME;
      Exit;
    end;
    {}
    if ( cdsMaster.FieldByName( 'INVFORMAT' ).AsString = '3' ) then
    begin
      if ( cdsMaster.FieldByName( 'BUSINESSID' ).AsString = EmptyStr ) then
      begin
        aErrMsg := '此張為手工三聯式發票, 必須輸入統編。';
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
        aErrMsg := '(營業人)發票不可為【捐贈】。';
        aFocusControl := txtInvUse.CodeNo;
        Exit;
      end;
    end;
    {}
    if ( cdsMaster.FieldByName( 'INVAMOUNT' ).AsFloat <= 0 ) then
    begin
      aErrMsg := '發票總金額不可為零或是負值。';
      Exit;
    end;
    if ( GetMasterTaxCode = '1' ) and
       ( cdsMaster.FieldByName( 'TAXRATE' ).AsFloat <= 0 ) then
    begin
      aErrMsg := '此為應稅發票, 稅率必須大於零。';
      aFocusControl := txtTAXRATE;
      Exit;
    end;
    {}
    {#5966 統一編號有值不可設定為電子發票 By Kin 2011/03/29}
    if ( cdsMaster.FieldByName( 'BUSINESSID' ).AsString <> EmptyStr ) and
      ( chkEInv.Checked ) and (chkEInv.Visible ) then
    begin
      aErrMsg := '統一編號有值，不可設定為 [ 電子發票 ]！';
      Exit;
    end;
    {#6370 增加判斷是否可選擇上傳註記=1的發票字軌(inv099) By Kin 2012/11/23}
    {#6370 取消判斷是否可選擇上傳註記的發票字軌 By Kin 2012/12/06}
    if ( dtmMain.sG_StarAllUpload ) and ( cdsMaster.FieldByName( 'BusinessId' ).AsString <> EmptyStr ) and
      ( 1 = 0 ) then
    begin

      aDataSet :=  TADOQuery( dtmMain.adoComm );
      aDataSet.Close;
      aDataSet.SQL.Text :=format( 'select count(*) from inv099    ' +
                  ' WHERE IDENTIFYID1 = ''%s''        ' +
                  '     AND IDENTIFYID2 = %s          ' +
                  '     AND COMPID = ''%s''           ' +
                  '     AND PREFIX = ''%s''           ' +
                  '     AND UPLOADFLAG = 1            ' +
                  '     AND INVFORMAT = 1             ' +
                  '     AND YEARMONTH = ''%s''        ' +
                  '     AND STARTNUM = ''%s''         ' ,
                  [IDENTIFYID1,IDENTIFYID2,dtmMain.getCompID,
                  Copy(cdsMaster.FieldByName( 'INVID' ).AsString,1,2),
                  FChooseInv.YearMonth,FChooseInv.StartNum]);

      aDataSet.Open;
      if aDataSet.Fields[0].AsInteger > 0 then
      begin
        aErrMsg := '統一編號有值，不可選取上傳發票字軌！';
        Exit;
      end;

      aDataSet.Close;
    end;
    {#6764 如果uploadflag = 1 則會員載具一定要有值 By Kin 2014/05/02}
    {只有現場開立才需要檢核是否有會員載具 By Kin 2014/07/28}
    if ( cdsMaster.FieldByName('FixFlag').AsInteger = 1 ) and
      ( FInvCreateType in [icLocale] ) then
    begin
      if ( cdsMaster.FieldByName( 'A_CARRIERID1' ).AsString = EmptyStr ) then
      begin
        aErrMsg := '無法取得會員載具資訊，無法存檔！';
        Exit;
      end;
    end;
    if IsGiveCanEdit then
    begin
      if  ( txtLoveNum.Text = EmptyStr ) and ( txtGiveUnit.Value = EmptyStr ) then
      begin
        aErrMsg := '發票為捐贈，愛心碼或受贈單位需有值！';
        aFocusControl := txtLoveNum;
        Exit;
      end;
    end;

    {}
    {#5966 現場開立不可設定為電子發票 By Kin 2011/04/15}
    {此判斷Jacy說PM說明會後決定要拿掉 By Kin 2011/05/10}
    {
    if  FInvCreateType = icLocale then
    begin
       if ( chkEInv.Checked ) and ( chkEInv.Visible ) then
       begin
        aErrMsg := '現場開立，不可設定為 [ 電子發票 ]！';
        Exit;
       end;
    end;
    }
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
  aInvCount, aWillCreateCount, aIndex,aMergeNum,aMainInvCnt : Integer;
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
      aErrMsg := '發票明細必須輸入。';
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
            { 收費單號檢查 }
            if ( not dtmMain.GetLinkToMIS ) or
               ( cdsDetail.FieldByName( 'LINKTOMIS' ).AsString <> 'Y' ) then
            begin
               { 不與客服系統介接, 檢查是否輸入收費單號及項次即可 }
               if ( cdsDetail.FieldByName( 'BILLID' ).AsString = EmptyStr ) or
                  ( cdsDetail.FieldByName( 'BILLIDITEMNO' ).AsString = EmptyStr ) then
               begin
                 aErrMsg := '收費明細資料,請輸入收費單號及收費項次。';
                 aFocusColumn := DetailGridView.GetColumnByFieldName( 'BILLID' );
                 Exit;
               end;
            end;
            { 收費單號及項次 }
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
          { 服務別檢查 }
          if ( FInvCreateType in [icPrev, icAfter, icLocale] ) and
             ( cdsDetail.FieldByName( 'SERVICETYPE' ).AsString = EmptyStr ) then
          begin
            aErrMsg := '收費明細資料,請輸入服務別。';
            aFocusColumn := DetailGridView.GetColumnByFieldName( 'SERVICETYPE' );
            Exit;
          end;
          { 收費項目檢查 }
          if ( cdsDetail.FieldByName( 'ITEMID' ).AsString = EmptyStr ) or
             ( cdsDetail.FieldByName( 'DESCRIPTION' ).AsString = EmptyStr ) then
          begin
            aFocusColumn := DetailGridView.GetColumnByFieldName( 'ITEMID' );
            aErrMsg := '收費明細資料, 請輸入收費項目。';
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
          { 稅別及稅率檢查 - 免稅 }
          if ( cdsDetail.FieldByName( 'TAXCODE' ).AsString <> '' ) then
          begin
            if ( GetMasterTaxCode <> cdsDetail.FieldByName( 'TAXCODE' ).AsString ) then
            begin
              aErrMsg := '明細資料,稅別與發票不符。';
              aFocusColumn := DetailGridView.GetColumnByFieldName( 'ITEMID' );
              Exit;
            end;
          end;
          { 金額檢查 }
          if ( GetMasterTaxCode <> '1' ) then
          begin
            if ( aTaxAmount <> 0 ) then
            begin
              aErrMsg := '此為免稅發票,明細稅額不可輸入。';
              aFocusColumn := DetailGridView.GetColumnByFieldName( 'TAXAMOUNT' );
              Exit;
            end
          end else
          begin
            { 若為應稅, 總金額為 10 塊錢內或 -10 塊錢, 則算出來的稅額會是零,
              所以大於10塊錢或-10塊錢才會有稅額 } 
            if ( aTaxAmount = 0 ) and ( Abs( aTotalAmount ) > 10 ) then
            begin
              aErrMsg := '此為應稅發票,明細稅額必須輸入。';
              aFocusColumn := DetailGridView.GetColumnByFieldName( 'TAXAMOUNT' );
              Exit;
            end;
          end;
          { 銷售額檢查 }
          { #6968 不檢查明細資料 By Kin 2014/12/22}
          if ( aSaleAmount = 0 ) then
          begin
            //aErrMsg := '明細資料, 銷售額必須輸入。';
            //aFocusColumn := DetailGridView.GetColumnByFieldName( 'SALEAMOUNT' );
            //Exit;
          end;
          { 總金額 }
          { #6968 不檢查明細資料 By Kin 2014/12/22}
          if ( aTotalAmount = 0 ) then
          begin
            //aErrMsg := '明細資料, 總金額必須輸入。';
            //aFocusColumn := DetailGridView.GetColumnByFieldName( 'TOTALAMOUNT' );
            //Exit;
          end;
          { 負項收費 }
          if ( cdsDetail.FieldByName( 'SIGN' ).AsString = '-' ) then
          begin
            if ( aUnitPtrice > 0 ) then
            begin
              aErrMsg := '明細資料為負項金額, 單價必須為負值。';
              aFocusColumn := DetailGridView.GetColumnByFieldName( 'UNITPRICE' );
              Exit;
            end;
            if ( aTaxAmount > 0 ) then
            begin
              aErrMsg := '明細資料為負項金額, 稅額必須為負值。';
              aFocusColumn := DetailGridView.GetColumnByFieldName( 'TAXAMOUNT' );
              Exit;
            end;
            if ( aSaleAmount > 0 ) then
            begin
              aErrMsg := '明細資料為負項金額, 銷售額必須為負值。';
              aFocusColumn := DetailGridView.GetColumnByFieldName( 'SALEAMOUNT' );
              Exit;
            end;
            if ( aTotalAmount > 0 ) then
            begin
              aErrMsg := '明細資料為負項金額, 總金額額必須為負值。';
              aFocusColumn := DetailGridView.GetColumnByFieldName( 'TOTALAMOUNT' );
              Exit;
            end;
          end else
          begin
            if ( aUnitPtrice < 0 ) then
            begin
              aErrMsg := '明細資料為正項金額, 單價必須為正值。';
              aFocusColumn := DetailGridView.GetColumnByFieldName( 'UNITPRICE' );
              Exit;
            end;
            if ( aTaxAmount < 0 ) then
            begin
              aErrMsg := '明細資料為正項金額, 稅額必須為正值。';
              aFocusColumn := DetailGridView.GetColumnByFieldName( 'TAXAMOUNT' );
              Exit;
            end;
            if ( aSaleAmount < 0 ) then
            begin
              aErrMsg := '明細資料為正項金額, 銷售額必須為正值。';
              aFocusColumn := DetailGridView.GetColumnByFieldName( 'SALEAMOUNT' );
              Exit;
            end;
            if ( aTotalAmount < 0 ) then
            begin
              aErrMsg := '明細資料為正項金額, 總金額額必須為正值。';
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
          aErrMsg := '明細資料, 發票加總金額小於或等於零。';
          aFocusColumn := nil;
          Exit;
        end;
        // 換開發票張數檢核
        aMergeNum := 0;
        //#5439 計算合併後明細總共幾筆 By Kin 2009/12/21
//        aMergeNum := ChkCreateNum;
        {
        if ( dtmMain.GetAutoCreateNum > 0 ) and
           ( cdsDetail.RecordCount > dtmMain.GetAutoCreateNum ) and
           ( aMergeNum > dtmMain.GetAutoCreateNum ) then
        }
        aMergeNum := CalMergeNum;
        if ( dtmMain.GetAutoCreateNum > 0 ) and
          ( aMergeNum > dtmMain.GetAutoCreateNum ) then
        begin
          if ( FMode in [dmInsert] ) then
          begin
            aInvCount := GetCurrentAvailableInvCount;
            //aWillCreateCount := Ceil( cdsDetail.RecordCount / dtmMain.GetAutoCreateNum );
            aWillCreateCount := Ceil( aMergeNum / dtmMain.GetAutoCreateNum );
            if ( aWillCreateCount > aInvCount ) then
            begin
              aErrMsg := Format( '本次預計開立發票張數為 %d 張, 剩餘可用發票張數為 %d 張, 發票張數不足, 請重新選擇發票本。',
                [aWillCreateCount, aInvCount] );
              aFocusColumn := nil;
              Exit;
            end;
          end else
          if ( FMode in [dmUpdate] ) then
          begin
            aErrMsg := Format( '發票明細筆數已超出換開筆數設定, 設定的筆數限制為 %d 筆。',
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
  TAutoCreate = record  //自動發票換開記錄
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
  
  { 先將 SO033 跟 SO034 的發票收費來源的資料先抓一份, }
  { 因為 Detail 不管是新增或是修改, 一律都是先用刪除的方式 }
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

  procedure UpdateSOFaciToCATV(const aInvId,aBillId,aItemId:String);
  const
    aSOTableName: array [1..2] of String = ( 'SO033', 'SO034' );
  var
    aIndex: Integer;
    aSQL: String;
    aAccountNo,aCustId,aFaciSNO,aSmartCardNo : String;

  begin
    try
      aAccountNo := EmptyStr;
      aCustId := EmptyStr;
      for aIndex := 1 to 2 do
      begin
        if aAccountNo = EmptyStr then
        begin
          aSQL := Format( 'SELECT AccountNo,CUSTID FROM %s.%s WHERE              ' +
              ' BILLNO=''%s'' AND ITEM = %s AND SERVICETYPE = ''C''              ',
              [dtmMain.getMisDbOwner,aSOTableName[aIndex],aBillId,aItemId]);

          dtmMainJ.adoComm.Close;
          dtmMainJ.adoComm.SQL.Text := aSQL;
          dtmMainJ.adoComm.Open;
          if dtmMainJ.adoComm.Recordset.RecordCount > 0 then
          begin
            aAccountNo := dtmMainJ.adoComm.FieldByName( 'ACCOUNTNO' ).AsString;
            aCustId := dtmMainJ.adoComm.FieldByName( 'CUSTID' ).AsString;
          end;
        end;
      end;
      dtmMainJ.adoComm.Close;
      if aAccountNo = EmptyStr then Exit;
      aSQL := Format('SELECT FaciSNO,FaciSeqNo FROM %s.SO003 WHERE         ' +
        ' NVL(STOPFLAG,0) = 0 AND SERVICETYPE IN (''I'',''D'')             ' +
        ' AND ACCOUNTNO = ''%s''                                           ' +
        ' AND CUSTID = %s AND FACISNO IS NOT NULL ORDER BY SERVICETYPE     ',
        [dtmMain.getMisDbOwner,aAccountNo, aCustId]);
      dtmMainJ.adoComm.SQL.Text := aSQL;
      dtmMainJ.adoComm.Open;
      if dtmMainJ.adoComm.Recordset.RecordCount > 0 then
      begin
        aSmartCardNo := EmptyStr;
        aFaciSNO := dtmMainJ.adoComm.FieldByName( 'FACISNO' ).AsString;
        if dtmMainJ.adoComm.FieldByName( 'FACISEQNO' ).AsString <> '' then
        begin
          aSQL := Format( 'SELECT SmartCardNo FROM %s.SO004 WHERE          ' +
            'SEQNO = ''%s''',
            [dtmMain.getMisDbOwner,
            dtmMainJ.adoComm.FieldByName( 'FACISEQNO').AsString] );
          dtmMainJ.adoComm.Close;
          dtmMainJ.adoComm.SQL.Text := aSQL;
          dtmMainJ.adoComm.Open;
          if dtmMainJ.adoComm.Recordset.RecordCount > 0 then
          begin
            if dtmMainJ.adoComm.FieldByName( 'SMARTCARDNO' ).AsString <> '' then
            begin
              aSmartCardNo := dtmMainJ.adoComm.FieldByName( 'SMARTCARDNO' ).AsString;
            end;
          end;
        end;

        aSQL := format('UPDATE INV008 SET FacisNo = ''%s'',SmartCardNo=''%s''  ' +
        ' WHERE INVID=''%s'' AND BillID = ''%s'' AND BillIDItemNo = ''%s''     ' ,
        [aFaciSNO,aSmartCardNo, aInvId,aBillId,aItemId]);
        aDataSet.Close;
        aDataSet.SQL.Text := aSQL;
        aDataSet.ExecSQL;
        aDataSet.Close;
      end else
      begin
        Exit;
      end;

    finally
      dtmMain.adoComm.Close;
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
      { 只有修改才須要抓INVOICETIME, 因為修改存檔之前一律先清掉客服系統對應的
        資料, 更新時必須把原值填回去 }
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
    {#5764 如果是CATV要找I或D服務,將找到的設備填入INV008.FACISNO}
    UpdateSOFaciToCATV(aInvId,aBillId,aItemId);
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
            cdsMaster.FieldByName( 'MEMO1' ).AsString + ' 設備:' + aObj.Memo1;

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
      { #5668 增加INVOICEKIND By Kin 2010/06/29 }
      if ( grpInvoiceKind.Visible ) and ( chkEInv.Checked ) then
        aDataSet.Parameters.ParamByName( 'INVOICEKIND' ).Value := 1
      else
        aDataSet.Parameters.ParamByName( 'INVOICEKIND' ).Value := 0;

      if txtGiveUnit.Value <> EmptyStr then
      begin
        aDataSet.Parameters.ParamByName( 'GIVEUNITID' ).Value :=
          txtGiveUnit.Value;
      end else
      begin
        aDataSet.Parameters.ParamByName( 'GIVEUNITID' ).Value := EmptyStr;
      end;
      if txtGiveUnit.ValueName <> EmptyStr then
      begin
         aDataSet.Parameters.ParamByName( 'GIVEUNITDESC' ).Value :=
          txtGiveUnit.ValueName;
      end else
      begin
        aDataSet.Parameters.ParamByName( 'GIVEUNITDESC' ).Value := EmptyStr;
      end;
      {#5668 增加常用Email與行動電話 By Kin 2010/07/27}
      if cdsMaster.FieldByName( 'EMAIL' ).AsString = EmptyStr then
        aDataSet.Parameters.ParamByName( 'EMAIL' ).Value := EmptyStr
      else
        aDataSet.Parameters.ParamByName( 'EMAIL' ).Value :=
          cdsMaster.FieldByName( 'EMAIL' ).Value;
      if cdsMaster.FieldByName( 'CONTMOBILE' ).AsString = EmptyStr then
        aDataSet.Parameters.ParamByName( 'CONTMOBILE' ).Value := EmptyStr
      else
        aDataSet.Parameters.ParamByName( 'CONTMOBILE' ).Value :=
         cdsMaster.FieldByName( 'CONTMOBILE' ).Value;
      {#6337 增加發票開立時間 By Kin 2012/11/13}
//      aDataSet.Parameters.ParamByName( 'CREATEINVDATE' ).Value := Now;
      {#6371 增加客戶載具CarrierType、CarrierId1、CarrierId2 By Kin 2013/02/25}
      if cdsMaster.FieldByName( 'CARRIERTYPE' ).AsString = EmptyStr then
        aDataSet.Parameters.ParamByName( 'CARRIERTYPE' ).Value := EmptyStr
      else
        aDataSet.Parameters.ParamByName( 'CARRIERTYPE' ).Value :=
          cdsMaster.FieldByName( 'CARRIERTYPE' ).Value;
      if cdsMaster.FieldByName( 'CARRIERID1' ).AsString = EmptyStr then
        aDataSet.Parameters.ParamByName( 'CARRIERID1' ).Value := EmptyStr
      else
        aDataSet.Parameters.ParamByName( 'CARRIERID1' ).Value :=
          cdsMaster.FieldByName( 'CARRIERID1' ).Value;
      if cdsMaster.FieldByName( 'CARRIERID2' ).AsString = EmptyStr then
        aDataSet.Parameters.ParamByName( 'CARRIERID2' ).Value := EmptyStr
      else
        aDataSet.Parameters.ParamByName( 'CARRIERID2' ).Value :=
          cdsMaster.FieldByName( 'CARRIERID2' ).Value;

      {#6600 增加愛心碼與仿偽隨機碼 By Kin 2013/10/01}
      if ( dtmMain.sG_StarComBine ) and ( cdsMaster.FieldByName( 'FIXFLAG' ).AsInteger = 1 ) then
      begin
        aDataSet.Parameters.ParamByName( 'RANDOMNUM' ).Value := RanDomNum
        {
        if FInvCreateType in [icLocale] then
          aDataSet.Parameters.ParamByName( 'RANDOMNUM' ).Value := RanDomNum
        else
          aDataSet.Parameters.ParamByName( 'RANDOMNUM' ).Value := EmptyStr;
        }
      end else
      begin
        aDataSet.Parameters.ParamByName( 'RANDOMNUM' ).Value := EmptyStr;
      end;
      if cdsMaster.FieldByName( 'LOVENUM' ).AsString = EmptyStr then
        aDataSet.Parameters.ParamByName( 'LOVENUM' ).Value := EmptyStr
      else
        aDataSet.Parameters.ParamByName( 'LOVENUM' ).Value :=
          cdsMaster.FieldByName( 'LOVENUM' ).Value;
      aDataSet.Parameters.ParamByName( 'FIXFLAG' ).Value := 0;
      if cdsMaster.FieldByName( 'FIXFLAG' ).AsInteger = 1 then
        aDataSet.Parameters.ParamByName( 'FIXFLAG' ).Value := 1;
      if cdsMaster.FieldByName( 'A_CARRIERID1' ).AsString = EmptyStr then
        aDataSet.Parameters.ParamByName( 'A_CARRIERID1' ).Value := EmptyStr
      else
        aDataSet.Parameters.ParamByName( 'A_CARRIERID1' ).Value :=
          cdsMaster.FieldByName( 'A_CARRIERID1' ).Value;
          
       if cdsMaster.FieldByName( 'A_CARRIERID1' ).AsString = EmptyStr then
        aDataSet.Parameters.ParamByName( 'A_CARRIERID2' ).Value := EmptyStr
      else
        aDataSet.Parameters.ParamByName( 'A_CARRIERID2' ).Value :=
          cdsMaster.FieldByName( 'A_CARRIERID2' ).Value;
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
      { #5668 增加發票種類 By Kin 2010/07/09 }
      if ( grpInvoiceKind.Visible ) and ( chkEInv.Checked ) then
        aDataSet.Parameters.ParamByName( 'INVOICEKIND' ).Value := 1
      else
        aDataSet.Parameters.ParamByName( 'INVOICEKIND' ).Value := 0;
      if txtGiveUnit.Value <> EmptyStr then
      begin
        aDataSet.Parameters.ParamByName( 'GIVEUNITID' ).Value :=
          txtGiveUnit.Value;
      end else
      begin
        aDataSet.Parameters.ParamByName( 'GIVEUNITID' ).Value := EmptyStr;
      end;
      if txtGiveUnit.ValueName <> EmptyStr then
      begin
         aDataSet.Parameters.ParamByName( 'GIVEUNITDESC' ).Value :=
          txtGiveUnit.ValueName;
      end else
      begin
        aDataSet.Parameters.ParamByName( 'GIVEUNITDESC' ).Value := EmptyStr;
      end;
      if txtMail.Text <> EmptyStr then
        aDataSet.Parameters.ParamByName( 'EMAIL' ).Value := txtMail.EditValue
      else
        aDataSet.Parameters.ParamByName( 'EMAIL' ).Value := EmptyStr;
      if txtContmobile.Text <> EmptyStr then
        aDataSet.Parameters.ParamByName( 'CONTMOBILE' ).Value := txtContmobile.EditValue
      else
        aDataSet.Parameters.ParamByName( 'CONTMOBILE' ).Value := EmptyStr;
      {#6371 增加客戶載具CarrierType、CarrierId1、CarrierId2 By Kin 2013/02/25}
      if cdsMaster.FieldByName( 'CARRIERTYPE' ).AsString = EmptyStr then
        aDataSet.Parameters.ParamByName( 'CARRIERTYPE' ).Value := EmptyStr
      else
        aDataSet.Parameters.ParamByName( 'CARRIERTYPE' ).Value :=
          cdsMaster.FieldByName( 'CARRIERTYPE' ).Value;
      if cdsMaster.FieldByName( 'CARRIERID1' ).AsString = EmptyStr then
        aDataSet.Parameters.ParamByName( 'CARRIERID1' ).Value := EmptyStr
      else
        aDataSet.Parameters.ParamByName( 'CARRIERID1' ).Value :=
          cdsMaster.FieldByName( 'CARRIERID1' ).Value;
      if cdsMaster.FieldByName( 'CARRIERID2' ).AsString = EmptyStr then
        aDataSet.Parameters.ParamByName( 'CARRIERID2' ).Value := EmptyStr
      else
        aDataSet.Parameters.ParamByName( 'CARRIERID2' ).Value :=
          cdsMaster.FieldByName( 'CARRIERID2' ).Value;

      
      if cdsMaster.FieldByName( 'RANDOMNUM' ).AsString = EmptyStr then
      begin
        aDataSet.Parameters.ParamByName( 'RANDOMNUM' ).Value := EmptyStr;
        if dtmMain.sG_StarComBine then
        begin
          if FInvCreateType in [icLocale] then
          begin
            aDataSet.Parameters.ParamByName( 'RANDOMNUM' ).Value := RanDomNum;
          end;
        end;
      end else
      begin
        aDataSet.Parameters.ParamByName( 'RANDOMNUM' ).Value :=
          cdsMaster.FieldByName( 'RANDOMNUM' ).Value;
      end;

      if cdsMaster.FieldByName( 'LOVENUM' ).AsString = EmptyStr then
        aDataSet.Parameters.ParamByName( 'LOVENUM' ).Value := EmptyStr
      else
        aDataSet.Parameters.ParamByName( 'LOVENUM' ).Value :=
          cdsMaster.FieldByName( 'LOVENUM' ).Value;
      aDataSet.Parameters.ParamByName( 'FIXFLAG' ).Value := 0;
      if cdsMaster.FieldByName( 'FIXFLAG' ).AsInteger = 1 then
        aDataSet.Parameters.ParamByName( 'FIXFLAG' ).Value := 1;
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
    aDataSet.Parameters.ParamByName( 'COMBCITEMCODE' ).Value :=
      cdsDetail.FieldByName( 'COMBCITEMCODE' ).AsString;
    aDataSet.Parameters.ParamByName( 'COMBCITEMNAME' ).Value :=
      cdsDetail.FieldByName( 'COMBCITEMNAME' ).AsString;
    if cdsDetail.FieldByName( 'SMARTCARDNO' ).AsString = EmptyStr then
    begin
      aDataSet.Parameters.ParamByName( 'SMARTCARDNO' ).Value := EmptyStr;
    end else
    begin
      aDataSet.Parameters.ParamByName( 'SMARTCARDNO' ).Value :=
        cdsDetail.FieldByName( 'SMARTCARDNO' ).Value;
    end;
    {#7076 增加信用卡末4碼 By Kin 2015/08/20}
    if cdsDetail.FieldByName( 'Creditcard4d' ).AsString = EmptyStr then
    begin
       aDataSet.Parameters.ParamByName( 'CREDITCARD4D' ).Value := EmptyStr;
    end else
    begin
      aDataSet.Parameters.ParamByName( 'CREDITCARD4D' ).Value :=
        cdsDetail.FieldByName( 'Creditcard4d' ).Value;
    end;
    aDataSet.ExecSQL;
    if ( FInvCreateType in [icPrev, icAfter, icLocale] ) then
    begin
      if ( dtmMain.GetLinkToMIS ) and
         ( cdsDetail.FieldByName( 'LINKTOMIS' ).AsString = 'Y' ) and
         ( not IsMergeRecord( aInvId, IntToStr( aSeq ) ) ) then
      begin
        { 把現有的客戶收費明細, 更新回 SO 對應 Table }
        UpdateInvoiceToSO( aInvId,
          cdsDetail.FieldByName( 'BILLID' ).AsString,
          cdsDetail.FieldByName( 'BILLIDITEMNO' ).AsString );
      end;
    end;
  end;

  { -------------------------------------------------------- }
  {---------------------------------------------------------- }
  function UpdateDetailDataSetByRecord(const aId, aBillId,
      aBillItemId,aSEQ : string ;aTotal,aPrice,aTax,aSale: Real):Boolean;
    var aSQL : string;
  begin
    if ( aBillId = EmptyStr ) or ( aBillItemId = EmptyStr ) then
    begin
      aSQL := format(' UPDATE INV008 SET UNITPRICE = %f, QUANTITY = 1,          ' +
              ' TAXAMOUNT = %f, TOTALAMOUNT = %f, SALEAMOUNT = %f               ' +
              ' WHERE  INVID = ''%s'' AND SEQ = %s                              ' ,
              [aPrice,aTax,aTotal,aSale,aId,aSEQ ] );
    end else
    begin

      aSQL := format(' UPDATE INV008 SET UNITPRICE = %f, QUANTITY = 1,          ' +
              ' TAXAMOUNT = %f, TOTALAMOUNT = %f, SALEAMOUNT = %f               ' +
              ' WHERE BILLID = ''%s'' AND BILLIDITEMNO = ''%s''                 ' +
              ' AND INVID = ''%s'' AND SEQ = %s                                 ' ,
              [aPrice,aTax,aTotal,aSale,aBillId,aBillItemId,aId,aSEQ ] );
    end;

    aDataSet.SQL.Text := aSQL;
    aDataSet.ExecSQL;

    if ( FInvCreateType in [icPrev, icAfter, icLocale] ) then
    begin
      if ( dtmMain.GetLinkToMIS ) and
         ( cdsDetail.FieldByName( 'LINKTOMIS' ).AsString = 'Y' )  then
      begin
        { 把現有的客戶收費明細, 更新回 SO 對應 Table }
        UpdateInvoiceToSO( aId,
          aBillId,
          aBillItemId );
      end;
    end;
  end;


  {---------------------------------------------------------- }
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
    aPreSeq, aInvId : String;
    aMinDateSQL, aMaxDateSQL,aDateSQL : String;
    aSmartCardSQL : String;
    aSeqSQL : String;
    aSEQ:Integer;
  begin
    cdsInv008A.First;
    aPreSeq := EmptyStr;
    {}

    {}
    while not cdsInv008A.Eof do
    begin
      {}
      //aDataSet.Parameters.Clear;
      aDataSet.SQL.Text := BuildInv008AApplySQL( dmInsert );
      aDataSet.Parameters.ParamByName( 'ITEMIDREF' ).Value :=
        cdsInv008A.FieldByName( 'ITEMIDREF' ).AsString;
      {}
      aInvId := cdsInv008A.FieldByName( 'INVID' ).AsString;
      if ( aInvId = '' ) then
      begin
        cdsDetail.Locate('BILLID;BILLIDITEMNO',
          VarArrayOf([ cdsInv008A.FieldByName('BILLID').AsString,
                     cdsInv008A.FieldByName('BILLIDITEMNO').AsString ]),[]);
        aInvId := cdsDetail.FieldByName( 'INVID' ).AsString;
      end;
      aDataSet.Parameters.ParamByName( 'INVID' ).Value := aInvId;
      {
      aDataSet.Parameters.ParamByName( 'INVID' ).Value :=
        cdsInv008A.FieldByName( 'INVID' ).AsString;
      }
      {}
      if cdsInv008A.FieldByName( 'BILLID' ).AsString = 'X' then
      begin
        aDataSet.Parameters.ParamByName( 'BILLID' ).Value := '';
      end else
      begin
        aDataSet.Parameters.ParamByName( 'BILLID' ).Value :=
          cdsInv008A.FieldByName( 'BILLID' ).AsString;
      end;
      {}
      //#6205 手開如果沒有billiditemno會出現錯誤 By Kin 2011/01/11
      if cdsInv008A.FieldByName( 'BILLIDITEMNO' ).Value = EmptyStr then
      begin
        aDataSet.Parameters.ParamByName( 'BILLIDITEMNO' ).Value := '';
      end else
      begin
        aDataSet.Parameters.ParamByName( 'BILLIDITEMNO' ).Value :=
          cdsInv008A.FieldByName( 'BILLIDITEMNO' ).AsString;
      end;
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
      {}
      aDataSet.Parameters.ParamByName( 'SMARTCARDNO' ).Value :=
        cdsInv008A.FieldByName( 'SMARTCARDNO' ).AsString;
      aDataSet.ExecSQL;

      //因為正負項排序後SEQ會亂掉，在此段程式重新調整SEQ By Kin 2011/03/23
      aSeqSQL := Format( 'SELECT SEQ FROM INV008 WHERE INVID = ''%s'''               +
                ' AND ITEMID = ''%s'''                                               ,
                [ aInvId,cdsInv008A.FieldByName( 'ITEMIDREF' ).AsString] );

      aDataSet.Close;
      aDataSet.SQL.Text := aSeqSQL;
      aDataSet.Open;
      aSEQ := aDataSet.Fields[0].AsInteger;
      aDataSet.Close;
      aSeqSQL := Format( 'UPDATE INV008A SET SEQ = %d                   ' +
            ' WHERE INVID = ''%s'' AND ITEMIDREF = ''%s'''                ,
            [ aSEQ,
            aInvId,
            cdsInv008A.FieldByName( 'ITEMIDREF' ).AsString ] );
      aDataSet.SQL.Text := aSeqSQL;
      aDataSet.ExecSQL;

      {-----------------------------------------------------------------------}
      {將合併後的最大與最小日期回填至INV008}



      aMinDateSQL := Format('(SELECT MIN(A.STARTDATE) FROM INV008A A,INV005 B WHERE    ' +
                     ' A.INVID = ''%s'' AND A.SEQ = %s AND                             ' +
                     ' A.ITEMID = B.ITEMID AND B.SIGN = ''+'' AND B.COMPID = ''%s'' )  ' ,
                     [ aInvId,cdsInv008A.FieldByName('SEQ').AsString,
                     dtmMain.getCompID ]);

      aMaxDateSQL := Format('(SELECT MAX(A.ENDDATE) FROM INV008A A,INV005 B  WHERE    ' +
                     ' A.INVID = ''%s'' AND A.SEQ = %s AND                            ' +
                     ' A.ITEMID = B.ITEMID AND B.SIGN = ''+'' AND B.COMPID = ''%s'' ) ' ,
                     [ aInvId,cdsInv008A.FieldByName('SEQ').AsString,
                     dtmMain.getCompID ]);

      aDateSQL := Format('UPDATE INV008 SET STARTDATE = %s,         ' +
              ' ENDDATE = %s WHERE INVID =''%s''                    ' +
              ' AND SEQ = %s AND IDENTIFYID1 = ''%s''               ' +
              ' AND IDENTIFYID2 = %s                                ' ,
              [aMinDateSQL,aMaxDateSQL,aInvId,
              cdsInv008A.FieldByName('SEQ').AsString,IDENTIFYID1,IDENTIFYID2]);

      {#5668 填入SmartCardNo By Kin 2010/07/09 }
      aSmartCardSQL := Format('(SELECT MAX(SMARTCARDNO) SMARTCARDNO FROM INV008A   ' +
                ' WHERE INVID = ''%s'' AND SEQ = %s AND SMARTCARDNO IS NOT NULL )  ' ,
                [ aInvId,cdsInv008A.FieldByName( 'SEQ' ).AsString ]);

      aSmartCardSQL := Format('UPDATE INV008 SET SMARTCARDNO = %s              ' +
              ' WHERE INVID = ''%s'' AND SEQ = %s AND IDENTIFYID1 = ''%s''     ' +
              ' AND IDENTIFYID2 = %s   ',
              [ aSmartCardSQL,aInvId,
                cdsInv008A.FieldByName( 'SEQ' ).AsString ,
                IDENTIFYID1,IDENTIFYID2 ]);
      aDataSet.SQL.Text := aDateSQL;
      aDataSet.ExecSQL;
      aDataSet.SQL.Text := aSmartCardSQL;
      aDataSet.ExecSQL;
      aDataSet.SQL.Text := BuildInv008AApplySQL( dmInsert );
      {-----------------------------------------------------------------------}

      { 更新客服 }
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
    { 是否須要把 useful 設成 "N" }
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
  { 將正項的SEQ排在前面,負項的排在後面 By Kin 2010/04/07 }
  procedure ChkSign(aInvId : String);
    var aSQL : string;
      aSeq,aIndex : Integer;
      aFindCnt : Integer;
  begin
    aSQL := 'UPDATE INV008 SET SEQ = SEQ - 100 WHERE INVID = ''%s''';
    aSQL := Format( aSQL,[aInvId]);
    aDataSet.Close;
    aDataSet.SQL.Text := aSQL;
    aDataSet.ExecSQL;
    aSQL := 'UPDATE INV008A SET SEQ = SEQ - 100 WHERE INVID =''%s''';
    aSQL := Format( aSQL,[aInvId]);
    aDataSet.Close;
    aDataSet.SQL.Text := aSQL;
    aDataSet.ExecSQL;
    aSeq := 0;
    try
      for aIndex := 0 to 1 do
      begin
        cdsDetail.Filter := EmptyStr;
        cdsDetail.Filtered := False;
        if aIndex = 0 then
          cdsDetail.Filter := 'INVID=''' + aInvId + ''' AND SIGN = ''+'''
        else
          cdsDetail.Filter := 'INVID=''' + aInvId + ''' AND SIGN = ''-''';
        cdsDetail.Filtered := True;
        cdsDetail.First;
        while not cdsDetail.Eof do
        begin
          //無問題集，增加判斷是否有找到SEQ,有找到才需要Seq+1 By Kin 2011/03/23 
          aFindCnt :=0;
          aSQL := 'SELECT COUNT(*) CNT FROM INV008 WHERE INVID = ''%s''' +
                  ' AND SEQ = %d';
          aSQL := Format( aSQL,
              [aInvId,cdsDetail.FieldByName( 'SEQ' ).AsInteger - 100] );
          aDataSet.Close;
          aDataSet.SQL.Text := aSQL;
          aDataSet.Open;
          aFindCnt := aDataSet.Fields[0].AsInteger;
          if aFindCnt > 0 then
          begin
            aSeq := aSeq + 1;
            aSQL := 'UPDATE INV008 SET SEQ = %d Where INVID = ''%s''' +
                    ' AND SEQ = %d';
            aSQL := Format(aSQL, [ aSeq,aInvId,
                cdsDetail.FieldByName( 'SEQ' ).AsInteger - 100 ] );
            aDataSet.Close;
            aDataSet.SQL.Text := aSQL;
            aDataSet.ExecSQL;
            aSQL := 'UPDATE INV008A SET SEQ= %d WHERE INVID = ''%s''' +
                    ' AND SEQ = %d';
            aSQL := Format(aSQL, [ aSeq,aInvId,
                cdsDetail.FieldByName( 'SEQ' ).AsInteger - 100 ] );
            aDataSet.Close;
            aDataSet.SQL.Text := aSQL;
            aDataSet.ExecSQL;
          end;
          cdsDetail.Next;
        end;
      end;
    finally
      cdsDetail.Filter := EmptyStr;
      cdsDetail.Filtered := False;
    end;
  end;
  { -------------------------------------------------------- }
  function ApplyInsertData: Integer;
  var
    aIndex, aDetailSeq,aMergeRcd : Integer;
    aFirstTimeApplyMaster: Boolean;
    aMainInvId, aPreInvId, aMsg,aMergeSEQ: String;
    aInvList: TList;
    aAutoCreate: PAutoCreate;
    aFaciMemo,aItemIdRefId,aItemDesc : string;
    aMerge : Boolean;
    aBookMark, aInv008AMark : Pointer;
    aMergeSale,aMergeUnit,aMergeTax,aMergeTotal : Double;
    aMergeQty : Integer;
    //aHaveDo : Boolean;
    aInvNum : Integer;
    //aHaveMerge : Boolean;
    //aRcdNo: Integer;
    aBillId, aBillItemNo, aInvId : String;
    aSeq : Integer;
  begin
    {#5439要合併收項目,所以整個大改寫 By Kin 2009/12/30}
    Result := 0;
    aInvNum := 0;
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
            // 假如有設定自動換開發票, 且已達到自動換開筆數, 重置序號
            {
            if (( dtmMain.GetAutoCreateNum > 0 ) and
               ( aDetailSeq >= dtmMain.GetAutoCreateNum ) and
               ( aInvNum > dtmMain.GetAutoCreateNum )) or
               (ChkCreateNum(aInvNum)>dtmMain.GetAutoCreateNum)   then
            begin
              aInvNum :=0;
              aDetailSeq := 0 ;
              aFirstTimeApplyMaster := False;
              aHaveDo := False;
            end;
            }
            aMerge := False;
            if ChkCreateNum(aInvNum) then   //判斷是否要換開發票
            begin
              aInvNum := 0;
              aDetailSeq := 0 ;
              aFirstTimeApplyMaster := False;
            end;
            //#5358 如果顯示設備,要將設備序號寫入Memo1 By Kin 2009/12/14
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
              Inc(aInvNum);

              New( aAutoCreate );
              ZeroMemory( aAutoCreate ,SizeOf( aAutoCreate^ ));
              aInvList.Add( aAutoCreate );

              if ( aFirstTimeApplyMaster ) then
              begin
                { 第一張發票號碼 }
                aAutoCreate.InvId := FInvoiceId;
              end else
              begin
                // 自動換開發票號碼, 取新的發票號碼
                aAutoCreate.InvId := GetNextInvNumber( aPreInvId );

              end;
              {}
              aPreInvId := aAutoCreate.InvId;
              { 新增預設先帶發票號碼, 最後再用最後一張發票號碼更新 }
              aAutoCreate.MainInvId := aAutoCreate.InvId;
              { 所有相關發票的金額, 主發票金額 }
              aAutoCreate.MainSaleAmount :=
                cdsMaster.FieldByName( 'SALEAMOUNT' ).AsFloat;
              aAutoCreate.MainTaxAmount :=
                cdsMaster.FieldByName( 'TAXAMOUNT' ).AsFloat;
              aAutoCreate.MainInvAmount :=
                cdsMaster.FieldByName( 'INVAMOUNT' ).AsFloat;
              {}
              { 該張發金額重置 }

                aAutoCreate.SaleAmount := 0;
                aAutoCreate.TaxAmount := 0;
                aAutoCreate.InvAmount := 0;

              {}
              aFirstTimeApplyMaster := False;
              ApplyMasterDataSet( dmInsert, aAutoCreate );

            end;

            { 同步 INV008A 的資料, INVID 跟 SEQ }
            { 已不須要, 只有合併的才會寫入 INV008A, 本隻程式不會做合併動作
            if not SyncInv008ADataSet( dmUpdate, cdsDetail.FieldByName( 'SEQ' ).AsString,
              IntToStr( aDetailSeq ), aAutoCreate.InvId, aMsg ) then
            begin
              raise Exception.Create( aMsg );
            end; }
            { 把本次 Detail 的 INVID, 及 SEQ 也更新, 千萬不可拿掉,
              不然後面 UpdateInvoiceToSO 會對應不到 }

            cdsDetail.Edit;
            cdsDetail.FieldByName( 'INVID' ).AsString := aAutoCreate.InvId;
            cdsDetail.FieldByName( 'SEQ' ).AsInteger := aInvNum;
            cdsDetail.Post;
            aItemIdRefId := cdsDetail.FieldByName( 'ITEMIDREF' ).AsString;
            if aItemIdRefId <> '' then
            begin
              try
                aBookMark := cdsDetail.GetBookmark;
                aMergeSale := 0;
                aMergeUnit := 0;
                aMergeTax := 0;
                aMergeSale := 0;
                 //找之前是否已經有合併過
                if ChkHasMerge( cdsDetail.RecNo ,aItemIdRefId ) then
                begin
                  aMerge := True;
                  aMergeTotal := cdsDetail.FieldByName('TOTALAMOUNT').AsFloat;
                  aMergeUnit := cdsDetail.FieldByName('UNITPRICE').AsFloat ;
                  aMergeTax := cdsDetail.FieldByName('TAXAMOUNT').AsFloat;
                  aBillId := cdsDetail.FieldByName( 'BILLID' ).AsString;
                  aBillItemNo := cdsDetail.FieldByName( 'BILLIDITEMNO' ).AsString;
                  aSeq := cdsDetail.FieldByName( 'SEQ' ).AsInteger;
                  aInvId := cdsDetail.FieldByName( 'INVID' ).AsString;
                  aMergeQty := cdsDetail.FieldByName( 'QUANTITY' ).AsInteger;
                  aMergeUnit := aMergeQty * aMergeUnit;
                  aMergeRcd := cdsDetail.RecNo;
                end;
              finally
                cdsDetail.GotoBookmark( aBookMark );
                cdsDetail.FreeBookmark( aBookMark );
              end;

              if aMerge then
              begin
                //如果有合併資料,要將當筆的SEQ 變更成來源的SEQ
                cdsDetail.Edit;
                cdsDetail.FieldByName( 'INVID' ).AsString := aInvId;
                cdsDetail.FieldByName( 'SEQ' ).AsInteger := aSeq;
                cdsDetail.Post;
              end else
              begin
                //如果不是需要合併的資料,則SEQ要 + 1
                if aDetailSeq > 1 then Inc(aInvNum);
                aSeq := aInvNum;
              end;

              aItemIdRefId := cdsDetail.FieldByName( 'ITEMIDREF' ).AsString;
              aItemDesc := GetItemDesc( aItemIdRefId );

            end else
            begin
              if aDetailSeq > 1 then inc(aInvNum);
              aSeq := aInvNum;
              aItemIdRefId := cdsDetail.FieldByName( 'ITEMID' ).AsString;
              aItemDesc := cdsDetail.FieldByName( 'DESCRIPTION' ).AsString;
            end;

            { 如果有合併資料,將INV008A.SEQ變成此次發票的SEQ }
            if cdsDetail.FieldByName( 'ITEMIDREF' ).AsString <> '' then
            begin
              UpdMergeInv008A( cdsDetail.FieldByName( 'ITEMIDREF' ).AsString,
               cdsDetail.FieldByName( 'INVID' ).AsString,aSeq);
            end;
            cdsDetail.Edit;
            cdsDetail.FieldByName( 'SEQ' ).AsInteger := aSeq;
            cdsDetail.FieldByName( 'ITEMID' ).AsString := aItemIdRefId;
            cdsDetail.FieldByName( 'DESCRIPTION' ).AsString := aItemDesc;
            cdsDetail.Post;

            { INV008 }
            if aMerge then
            begin
              aBookMark := cdsDetail.GetBookmark;
              try
                aMergeTotal := aMergeTotal + cdsDetail.FieldByName('TOTALAMOUNT').AsFloat;
                if cdsMaster.FieldByName( 'TAXTYPE' ).AsInteger = 1 then
                begin
                  aMergeSale := Round( TUother.RoundFloat(
                    ( cdsDetail.FieldByName( 'QUANTITY' ).AsInteger *
                    cdsDetail.FieldByName('UNITPRICE').AsFloat ),0 ));
                  aMergeSale := aMergeSale + aMergeUnit;

                  aMergeTax := (aMergeTotal - aMergeSale);
                  aMergeUnit := aMergeSale;
                end else
                begin
                  aMergeSale := aMergeTotal;
                  aMergeUnit := aMergeSale;
                end;
                //跳到要合併源頭,將源頭的金額一併更新
                cdsDetail.RecNo := aMergeRcd;
                cdsDetail.Edit;
                cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsFloat := aMergeTotal;
                cdsDetail.FieldByName('UNITPRICE').AsFloat := aMergeUnit;
                cdsDetail.FieldByName('TAXAMOUNT').AsFloat := aMergeTax;
                cdsDetail.FieldByName( 'SALEAMOUNT' ).AsFloat := aMergeSale;
                cdsDetail.FieldByName( 'QUANTITY' ).AsInteger := 1;
                cdsDetail.Post;
                UpdateDetailDataSetByRecord( aInvId, aBillId, aBillItemNo, IntToStr(aSeq),
                     aMergeTotal, aMergeUnit,aMergeTax,aMergeSale);

              finally
                cdsDetail.GotoBookmark( aBookMark );
                cdsDetail.FreeBookmark( aBookMark );
              end;
            end else
            begin
              ApplyDetailDataSetByRecord(
                cdsDetail.FieldByName( 'INVID' ).AsString,
                aSeq );
            end;
            { 累計明細金額, 做為該張發票的金額 }
            for aIndex := 0 to aInvList.Count -1 do
            begin
              if PAutoCreate( aInvList[aIndex] ).InvId =
                cdsDetail.FieldByName( 'INVID' ).AsString then
              begin
                PAutoCreate(aInvList[aIndex]).SaleAmount :=
                  ( PAutoCreate(aInvList[aIndex]).SaleAmount +
                  cdsDetail.FieldByName( 'SALEAMOUNT' ).AsFloat );
                PAutoCreate(aInvList[aIndex]).TaxAmount :=
                  ( PAutoCreate(aInvList[aIndex]).TaxAmount +
                  cdsDetail.FieldByName( 'TAXAMOUNT' ).AsFloat );
                PAutoCreate(aInvList[aIndex]).InvAmount :=
                  ( PAutoCreate(aInvList[aIndex]).InvAmount +
                cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsFloat );
              end;
            end;

            cdsDetail.Next;

            { 已是最後一筆 Detail, 記錄最後一張的發票號碼 }
            if ( cdsDetail.Eof )  then aMainInvId := aAutoCreate.InvId;
          end;
          { 回頭更新主檔 }
          for aIndex := 0 to aInvList.Count - 1 do
          begin
            PAutoCreate( aInvList[aIndex] ).MainInvId := aMainInvId;
            {--------------------------------------------------}
            //#5358 增加ShowFaci 判斷在開立發票時是否要將設備序號寫入Memo1 By Kin 2009/12/14
            if aFaciMemo <> EmptyStr then
            begin
              PAutoCreate( aInvList[aIndex] ).Memo1 :=
                cdsMaster.FieldByName( 'Memo1' ).AsString + '設備:' + aFaciMemo;
            end;
            {--------------------------------------------------}
            UpdateAutoCreateInv( PAutoCreate( aInvList[aIndex] ) );
            ChkSign( PAutoCreate( aInvList[aIndex] ).InvId);

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


  { -------------------------------------------------------- }

  function ApplyUpdateData: Integer;
  var
    aList: TList;
    aIndex: Integer;
    aObj: PAutoCreate;
    aSale, aTax, aInv: Double;
    aOldSeq, aItemIdRefId, aItemDesc: String;
    aMerge : Boolean;
    aBookMark : Pointer;
    aMergeSale,aMergeUnit,aMergeTax,aMergeTotal : Double;
//    aRcdNo : Integer;
//    aHaveMerge : Boolean;
    aBillId, aBillItemNo, aInvId : string;
    aMergeQty, aMergeRcd : Integer;
    aInvNum, aSeq : Integer;
  begin
    Result := 0;
    { Detail 資料一律刪除後再用新增方式處理 }
    DeleteAllInvoiceDetail( FInvoiceId );
    if ( FInvCreateType in [icPrev, icAfter, icLocale] ) then
    begin
      if dtmMain.GetLinkToMIS then
      begin
        { 注意, 這裏不管 INV008A 所記錄的明細為何, 直接用 INVID 先
          回頭去把客服系統相關的收費資料先清掉, 下面再用 INV008A
          所記錄的資料再回頭去更新客服系統 }
        { 先抓一份 SO Table 的發票資料 }
        CloneInvoiceFromSO;
        { 將客服系統的發票資料清掉 }
        EmptyInvoiceToSO;
      end;
    end;
    aList := GetMainInvList( FInvoiceId, cdsMaster.FieldByName( 'MAININVID' ).AsString );
    try
      cdsDetail.DisableControls;
      DisableDetailDataSetEvent;
      aInvNum := 0;
      try
        { 先將所有 INV008 寫到 Table }
        cdsDetail.First;

        while not cdsDetail.Eof do
        begin
          aMerge := False;
//          if cdsDetail.RecNo = 1 then inc(aInvNum);
          aItemIdRefId := cdsDetail.FieldByName( 'ITEMIDREF' ).AsString;
          {要將資料合併 By Kin }
          {------------------------------------}
          if aItemIdRefId <> '' then
          begin
            try
              aBookMark := cdsDetail.GetBookmark;
              aMergeSale := 0;
              aMergeUnit := 0;
              aMergeTax := 0;
              aMergeSale := 0;
              if ChkHasMerge( cdsDetail.RecNo,aItemIdRefId ) then
              begin
                aMerge := True;
                aMergeTotal := cdsDetail.FieldByName('TOTALAMOUNT').AsFloat;
                aMergeUnit := cdsDetail.FieldByName('UNITPRICE').AsFloat ;
                aMergeTax := cdsDetail.FieldByName('TAXAMOUNT').AsFloat;
                aBillId := cdsDetail.FieldByName( 'BILLID' ).AsString;
                aBillItemNo := cdsDetail.FieldByName( 'BILLIDITEMNO' ).AsString;
                aSeq := cdsDetail.FieldByName( 'SEQ' ).AsInteger;
                aInvId := cdsDetail.FieldByName( 'INVID' ).AsString;
                aMergeQty := cdsDetail.FieldByName( 'QUANTITY' ).AsInteger;
                aMergeUnit := aMergeQty * aMergeUnit;
                aMergeRcd := cdsDetail.RecNo;
              end;
            finally
              cdsDetail.GotoBookmark( aBookMark );
              cdsDetail.FreeBookmark( aBookMark );
            end;

            if aMerge then
            begin
              cdsDetail.Edit;
              cdsDetail.FieldByName( 'INVID' ).AsString := aInvId;
              cdsDetail.Post;
            end else
            begin
              inc(aInvNum);
              aSeq := aInvNum;
//              aSeq := cdsDetail.FieldByName( 'SEQ' ).AsString;
            end;
            aItemIdRefId := cdsDetail.FieldByName( 'ITEMIDREF' ).AsString;
            aItemDesc := GetItemDesc( aItemIdRefId );

          end else
          begin
            Inc(aInvNum);
            aSeq := aInvNum;
            aItemIdRefId := cdsDetail.FieldByName( 'ITEMID' ).AsString;
            aItemDesc := cdsDetail.FieldByName( 'DESCRIPTION' ).AsString;
          end;

          if cdsDetail.FieldByName( 'ITEMIDREF' ).AsString <> '' then
          begin
            UpdMergeInv008A( cdsDetail.FieldByName( 'ITEMIDREF' ).AsString,
             cdsDetail.FieldByName( 'INVID' ).AsString,aSeq);
          end;

          cdsDetail.Edit;
          cdsDetail.FieldByName( 'SEQ' ).AsInteger := aSeq;
          cdsDetail.FieldByName( 'ITEMID' ).AsString := aItemIdRefId;
          cdsDetail.FieldByName( 'DESCRIPTION' ).AsString := aItemDesc;
          cdsDetail.Post;


          {-----------------------------------------}
          if aMerge then
          begin

            aBookMark := cdsDetail.GetBookmark;
            try
              aMergeTotal := aMergeTotal + cdsDetail.FieldByName('TOTALAMOUNT').AsFloat;

              if cdsMaster.FieldByName( 'TAXTYPE' ).AsInteger = 1 then
              begin

                aMergeSale := Round( TUother.RoundFloat(
                  ( cdsDetail.FieldByName( 'QUANTITY' ).AsInteger *
                  cdsDetail.FieldByName('UNITPRICE').AsFloat ),0 ));
                aMergeSale := aMergeSale + aMergeUnit;

                aMergeTax := (aMergeTotal - aMergeSale);
                aMergeUnit := aMergeSale;
              end;
              cdsDetail.RecNo := aMergeRcd;
              cdsDetail.Edit;
              cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsFloat := aMergeTotal;
              cdsDetail.FieldByName('UNITPRICE').AsFloat := aMergeUnit;
              cdsDetail.FieldByName('TAXAMOUNT').AsFloat := aMergeTax;
              cdsDetail.FieldByName( 'SALEAMOUNT' ).AsFloat := aMergeSale;
              cdsDetail.FieldByName( 'QUANTITY' ).AsInteger := 1;
              cdsDetail.Post;
              UpdateDetailDataSetByRecord( aInvId, aBillId, aBillItemNo,
                 IntToStr(aSeq), aMergeTotal, aMergeUnit,
                 aMergeTax, aMergeSale );

            finally
              cdsDetail.GotoBookmark( aBookMark );
              cdsDetail.FreeBookmark( aBookMark );
            end;
          end else
          begin
            ApplyDetailDataSetByRecord( FInvoiceId,
              aSeq );
          end;
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
        { 把此筆主檔資料 Copy 到其它筆 }
        { 順便將金額加起來 }
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
        { 更新所有相關發票 }
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
        EnableDetailDataSetEvent;
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
  // 自動換開發票用
  aConnection.BeginTrans;
  if dtmMain.GetLinkToMIS then aSoConnection.BeginTrans;
  try
    if ( FMode in [dmInsert] ) then
    begin
      aUseInvCount := ApplyInsertData;

      ApplyInv099Data( aUseInvCount );
      ApplyInv008AData; //不須要


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
      { PKey 重覆, PM 要求改秀訊息 }
      if OracleErros.ErrorCode = 1 then
      begin
        aErrMsg := '該發票號碼已被開立,請至<發票日期>處,重新選取字軌,即可開立!';
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
  // 一般開立, 客戶資料從發票系統 INV002 及 INV019 檢查
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
  // 現場開立, 預開, 後開, 客戶資料從 SO001 檢查, 若不與客服系統介接, 則不檢查
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
      { 第一次, 檢查 SO033, 檢查不到則跳過迴圈, 直接 check SO034 }
      if ( aIndex = 1 ) and ( not aValidateSO033 ) then
      begin
        Continue;
      end
      { 第一次, 檢查 SO033, 檢查到則檢查稅別及稅率 }
      else if ( aIndex = 1 ) and ( aValidateSO033 )then
      begin
        if ( aItemId <> '' ) then
        begin
          aTaxCode := aDataSet.FieldByName( 'TAXCODE' ).AsString;
          aTaxRate := aDataSet.FieldByName( 'RATE1' ).AsString;
          if ( aTaxCode <> GetMasterTaxCode ) or
             ( aTaxRate <> cdsMaster.FieldByName( 'TAXRATE' ).AsString ) then
          begin
            aErrMsg := '此筆收費項目的稅別/稅率與發票不同。';
            aValidateSO033 := False;
            Break;
          end;
        end else
        begin
          Break;
        end;
      end
      { 第二次, 檢查 SO034, 仍然檢查不到, 則無此單號 }
      else if ( aIndex = 2 ) and ( not aValidateSO034 ) then
      begin
        aErrMsg := '無此收費單號。';
        Break;
      end
      { 第二次, 檢查 SO034, 檢查到則檢查稅別及稅率 }
      else if ( aIndex = 2 ) and ( aValidateSO034 ) then
      begin
        if ( aItemId <> '' ) then
        begin
          aTaxCode := aDataSet.FieldByName( 'TAXCODE' ).AsString;
          aTaxRate := aDataSet.FieldByName( 'RATE1' ).AsString;
          if ( aTaxCode <> GetMasterTaxCode ) or
             ( aTaxRate <> cdsMaster.FieldByName( 'TAXRATE' ).AsString ) then
          begin
            aErrMsg := '此筆收費單號的稅別/稅率與發票不符。';
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
      aErrMsg := '請輸入收費項目。'
    else
      aErrMsg := '無此收費項目, 請檢查稅別是否正確。';
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
  cdsMaster.FieldByName( 'FIXFLAG' ).ReadOnly := False;
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
      txtGiveUnit.Clear;
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

      txtGiveUnit.Value := cdsMaster.FieldByName( 'GIVEUNITID' ).AsString;
      txtGiveUnit.ValueName := cdsMaster.FieldByName( 'GIVEUNITDESC' ).AsString;

      {#5668 顯示發票種類 By Kin 2010/07/08}
      if dtmMain.GetStarEInvoice then
      begin
        grpInvoiceKind.Visible := True;
      end;
      { #5668 如果格式不是電子式,不要顯示發票種類 By Kin 2010/07/08}
      if cdsMaster.FieldByName( 'INVFORMAT' ).AsString <> '' then
      begin
        if cdsMaster.FieldByName( 'INVFORMAT' ).AsString <> '1' then
        begin
          chkEInv.Checked := False;
          //#5922 InvFormat不是1不能選擇 By Kin 2011/02/11
          grpInvoiceKind.Enabled := False;
          //grpInvoiceKind.Visible := False;
        end;
      end;
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
  { 只須判斷當筆即可 }
  cdsMaster.FieldByName( 'TAXTYPE' ).AsString :=
    cdsDetail.FieldByName( 'TAXCODE' ).AsString;
  if cdsMaster.FieldByName( 'TAXTYPE' ).AsString <> '1' then
    cdsMaster.FieldByName( 'TAXRATE' ).AsString := '0';
  MasterDataToEditor;
  { 重新設定一次 TaxRate , 因為在 MasterDataToEditor 中已把 rdoTAXRATE 的
    radio button onclick 關掉, 所以不會觸發 }
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
    { 檢核發票主檔的稅別是否與明細不同, 不同即刪除 }
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
      { 重算所有金額 }
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
  { 當為支出(-)時, 自動轉為負值 }
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
  { 當為收入(+)時, 自動轉為正值, 或是 ChargeSign 等於空值也識為 (+) }
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
      aCanModify := '否';
      if dtmMain.adoComm.FieldByName( 'CANMODIFY' ).AsString = 'Y' then
        aCanModify := '是';
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
  aFacisNo, aAccountNo, aSmartCard: String;
begin
  if not IsMergeRecord( FInvoiceId,
    cdsDetail.FieldByName( 'SEQ' ).AsString ) then
  begin
    {#5668 增加存入 SmartCard By Kin 2010/07/09}
    dtmSO.GetMisAccountAndFacisNo(
      cdsDetail.FieldByName( 'BILLID' ).AsString,
      cdsDetail.FieldByName( 'BILLIDITEMNO' ).AsString,
      cdsMaster.FieldByName( 'CUSTID' ).AsString, aAccountNo,aFacisNo,aSmartCard );
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
      btnPrintRecord.Enabled := True;
      btnUploadRecord.Enabled := True;
      FInvCreateType := StrToInvoicType(
        cdsMaster.FieldByName( 'HOWTOCREATE' ).AsString );
    end else
    begin
      btnPrintRecord.Enabled := False;
      btnUploadRecord.Enabled := False;
    end;
    FInvoiceId := cdsMaster.FieldByName( 'INVID' ).AsString;
    MasterDataToEditor;
    FMultiInv := dtmMainJ.IsMultiInv( FInvoiceId );
    { 反向回填次張發票的 INV099 結構 }
    FillChooseInv;
    { 帶發票明細 }
    DisableDetailDataSetEvent;
    try
      { INV008A }
      PrepareInv008ADataSet;
      FillInv008ADataSet;

      { INV008 }
      PrepareDetailDataSet;
      FillDetailDataSet;

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
      begin
        if ( FMode in [ dmInsert ] ) then
        begin
          GetCustomerInfo( txtCUSTID.EditingText );
        end;
      end  else
      begin
        if ( FInvCreateType in [icLocale, icPrev, icAfter] ) then
          WarningMsg( Format( '客服系統查無此客編(%s), 請重新輸入。',
            [txtCUSTID.EditingText] ) )
        else
          WarningMsg( Format( '發票系統查無此客編(%s), 請重新輸入。',
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
  if txtBUSINESSID.Properties.ReadOnly then
  begin
    Exit;
  end;
  if ( FMode in [dmInsert,dmUpdate] ) then
  begin
    if ( txtBUSINESSID.EditingText <> '' ) then
    begin
      if not dtmMain.checkBusinessID( txtBUSINESSID.EditingText ) then
      begin
        WarningMsg( '統一編號輸入有誤。' );
        if txtBUSINESSID.CanFocus then
        begin
          txtBUSINESSID.SetFocus;
          txtBUSINESSID.SelectAll;
        end;
      end;
      if ( cdsMaster.FieldByName('DestroyUploadTime').AsString <> '' ) then
      begin
        if (cdsMaster.FieldByName( 'DestroyUploadTime' ).AsDateTime > cdsMaster.FieldByName( 'UploadTime' ).AsDateTime) then
        begin
          Exit;
        end;
      end;
      if cdsMaster.FieldByName( 'UPLOADFLAG' ).AsString = '1' then
      begin
        WarningMsg( '已產生上傳檔案，不允許輸入統一編號!');
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
      // 現場開立, 預開, 後開的發票, 客編來源為 SMS 的收費資料
      if ( FInvCreateType in [icLocale, icPrev, icAfter] ) then
      begin
        if txtCUSTID.EditingText = '' then
        begin
          WarningMsg( '請輸入客編。' );
          Exit;
        end;
        if not ValidateCustomer( txtCUSTID.EditingText, FInvCreateType ) then
        begin
          WarningMsg( Format( '客服系統查無此客編(%s), 請重新輸入。',
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
      // 一般開立的發票, 客編來源為發票系統的收費資料
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
      WarningMsg( '請輸入發票日期。' );
      Abort;
    end;
    if not TryStrToDate( txtINVDATE.EditingText, aInvDate ) then
    begin
      WarningMsg( '您輸入的發票日期不正確。' );
      Abort;
    end;
    {#6370 增加上傳註記欄位 By Kin 2012/11/26}
    aDataSet := TADOQuery( dsChooseInv.DataSet );
    aDataSet.Close;
    aDataSet.SQL.Text := Format(
      ' SELECT SUBSTR( A.YEARMONTH, 1, 4 ) ||                    ' +
      '        SUBSTR( A.YEARMONTH, 5, 2 ) YEARMONTH,            ' +
      '        A.INVFORMAT,                                      ' +
      '        DECODE( A.INVFORMAT,                              ' +
      '                ''1'', ''電子'',                          ' +
      '                ''2'', ''手二'',                          ' +
      '                ''3'', ''手三'',                          ' +
      '                A.INVFORMAT ) INVFORMATDESC,              ' +
      '        A.PREFIX,                                         ' +
      '        A.STARTNUM,                                       ' +
      '        A.ENDNUM,                                         ' +
      '        A.CURNUM,                                         ' +
      '        TO_CHAR( A.LASTINVDATE,                           ' +
      '                 ''YYYY/MM/DD'' ) LASTINVDATE,            ' +
      '        A.MEMO,                                           ' +
      '        NVL(A.UploadFlag,0) UPLOADFLAG                    ' +
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
    WarningMsg( '您輸入的發票日期不正確。' );
    ErrorText := EmptyStr;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.txtCHARGEDATEPropertiesValidate(Sender: TObject;
  var DisplayValue: Variant; var ErrorText: TCaption; var Error: Boolean);
begin
  if Error then
  begin
    WarningMsg( '您輸入的收費日期不正確。' );
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
        { 判斷手動更改的日期, 跟原本已對應到的發票所屬年月是否一樣 }
        if not TryStrToDate( txtINVDATE.Text, aDate ) then
        begin
          WarningMsg( '您輸入的發票日期不正確。' );
          if txtINVDATE.CanFocus then txtINVDATE.SetFocus;
          Exit;
        end;
        aNoInv := False;
        aRealInvMonth := MapingInvoiceYearMonth( aDate );
        { 對應到的發票所屬年月是否一樣 }
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
        { 若是日期輸入正確, 記錄下來 }
        FChooseInv.OldInvDate :=  FormatDateTime( 'yyyymmdd', aDate );
        { 若是一樣, 要檢查最後發票日或此年月沒有發票本 }
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
        // 有資料, 判斷日期是否大於最後發票開立日
        if not aNoInv then
          aNoInv := ( aDate < aDataSet.FieldByName( 'LASTINVDATE' ).AsDateTime );
        if ( aNoInv ) then
        begin
          WarningMsg( '您輸入的發票日期小於發票本的【最後發票開立日】。' );
          if txtINVDATE.CanFocus then txtINVDATE.SetFocus;
        end;
        aDataSet.Close;
      end;
    end else
    if ( FMode in [dmUpdate] ) then
    begin
      { 判斷手動更改的日期, 跟原本已對應到的發票所屬年月是否一樣 }
      if not TryStrToDate( txtINVDATE.Text, aDate ) then
      begin
        WarningMsg( '您輸入的發票日期不正確。' );
        if txtINVDATE.CanFocus then txtINVDATE.SetFocus;
        Exit;
      end;
      aRealInvMonth := MapingInvoiceYearMonth( aDate );
      { 對應到的發票所屬年月是否一樣 }
      if aRealInvMonth <> FChooseInv.YearMonth then
      begin
        WarningMsg( '您變更的發票日期已不同於此張發票所屬【發票年月】, 請重新輸入發票日期。' );
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
      WarningMsg( '您輸入的發票日期不正確。' );
      Exit;
    end;
    if ( aDataSet.FieldByName( 'PREFIX' ).AsString = EmptyStr ) then
    begin
      WarningMsg( '該發票本查無發票字軌, 請選擇其它發票本。' );
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
    {#6629 增加字軌屬性 By Kin 2013/11/01}
    cdsMaster.FieldByName( 'FIXFLAG' ).AsInteger :=
      aDataSet.FieldByName( 'UPLOADFLAG' ).AsInteger;
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
    { 手二, 不可輸入統編 }
    if ( cdsMaster.FieldByName( 'INVFORMAT' ).AsString = '2' ) then
    begin
      cdsMaster.FieldByName( 'BUSINESSID' ).AsString := EmptyStr;
      txtBUSINESSID.Properties.ReadOnly := True;
    end;
    { 手三, 必須輸入統編 }
    if ( cdsMaster.FieldByName( 'INVFORMAT' ).AsString = '3' ) then
      cxLabel6.Style.StyleController := dtmMain.cxLabelNotNullStyle;
    { '#5668 如果不是電子發票,不要顯示發票種類 By Kin 2010/06/29 }
    {#5922 測試不OK，如果不是電子式發票"電子發票選票不能勾選 By Kin 2011/03/04}
    grpInvoiceKind.Enabled := True;
    if FChooseInv.InvFormat <> '1' then
    begin
      chkEInv.Checked := False;
      cdsMaster.FieldByName( 'INVOICEKIND' ).AsString := '0';
      grpInvoiceKind.Enabled := False;

      //grpInvoiceKind.Visible := False;
    end;
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
    icLocale: cdsMaster.FieldByName( 'HOWTOCREATEDESC' ).AsString := '現場開立';
    icNormal: cdsMaster.FieldByName( 'HOWTOCREATEDESC' ).AsString := '一般開立';
    icPrev: cdsMaster.FieldByName( 'HOWTOCREATEDESC' ).AsString := '預開';
    icAfter: cdsMaster.FieldByName( 'HOWTOCREATEDESC' ).AsString := '後開';
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
  cdsMaster.FieldByName( 'ISOBSOLETEDESC' ).AsString := '否';
  cdsMaster.FieldByName( 'CANMODIFYDESC' ).AsString := '是';
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
  // 當為修改模式時, 已超出設定自動換開筆數, 不允許新增 }
  // 移到存檔時檢查
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
  // 預開, 後開, 現場開
  if FInvCreateType in [icPrev, icAfter, icLocale] then
  begin
    cdsDetail.FieldByName( 'EMPTY' ).AsString := 'Y';
    cdsDetail.FieldByName( 'LINKTOMIS' ).AsString :=
      IIF( dtmMain.GetLinkToMIS, 'Y', 'N' );
    cdsDetail.FieldByName( 'LINKTOMISDESC' ).AsString := '是'
  end else
  begin
    // 一般開立也要能輸入 收費單號及項次, 服務別, 收費日起迄
    cdsDetail.FieldByName( 'EMPTY' ).AsString := 'Y';
    cdsDetail.FieldByName( 'LINKTOMIS' ).AsString := 'N';
    cdsDetail.FieldByName( 'LINKTOMISDESC' ).AsString := '否';
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
  { 同步 INV008ADataSet, 存檔時寫入 INV008A }
  if not SyncInv008ADataSet( dmDelete,
    cdsDetail.FieldByName( 'SEQ' ).AsString, EmptyStr,  FInvoiceId, EmptyStr, aMsg ) then
  begin
    ErrorMsg( Format( '同步發票合併資料檔錯誤, 訊息:%s。', [aMsg] ) );
    Abort;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.cdsDetailBeforePost(DataSet: TDataSet);
var
  aMsg: String;
  aNewSeq: String;
  aItemIdRefId : String;
begin
  // 此動作已不須要, 只有文字檔案匯入時才有合併動作,
  // 此開立已不須要寫入 INV008A
  { 同步 INV008ADataSet, 存檔時寫入 INV008A}
  {#5439 發票要合併,先找收費資料合併,找完後再與INV005合併 By Kin 2009/12/22}
  {
  if FNoUpdate then
  begin
    FNoUpdate := False;
    Exit;
  end;
  }
  {#5439 將收費項目合併}
  aNewSeq := EmptyStr;
  aItemIdRefId := EmptyStr;
  if cdsDetail.FieldByName( 'COMBCITEMCODE' ).AsString <> '' then
  begin
    aItemIdRefId := GetItemIdRef( cdsDetail.FieldByName( 'COMBCITEMCODE' ).AsString );
    if aItemIdRefId = EmptyStr then
    begin
      aItemIdRefId := cdsDetail.FieldByName( 'COMBCITEMCODE' ).AsString;
    end;
  end else
  begin
    aItemIdRefId := GetItemIdRef( cdsDetail.FieldByName( 'ItemID' ).AsString );
  end;
  //有找到需要合併的項目
  if aItemIdRefId <> EmptyStr then
  begin
    if cdsInv008A.Locate( 'ITEMIDREF;INVID', VarArrayOf([ aItemIdRefId,FInvoiceId ]),[] ) then
    begin
      aNewSeq := cdsInv008A.FieldByName( 'SEQ' ).AsString;
    end else
    begin
      if not cdsInv008A.IsEmpty then
      begin
        cdsInv008A.DisableControls;
        cdsInv008A.Last;
        try
          aNewSeq := IntToStr(cdsInv008A.FieldByName('SEQ').AsInteger + 1);
        finally
          cdsInv008A.EnableControls;
        end;
      end else
      begin
        aNewSeq := cdsDetail.FieldByName( 'SEQ' ).AsString;
      end;
    end;
    cdsDetail.FieldByName( 'ITEMIDREF' ).AsString := aItemIdRefId;
    cdsDetail.FieldByName( 'MERGESEQ' ).AsString := aNewSeq;
    if cdsDetail.State = dsEdit then
    begin
      if not SyncInv008ADataSet( dmUpdate , cdsInv008A.FieldByName('SEQ').AsString,
        aNewSeq, FInvoiceId, aItemIdRefId, aMsg ) then
      begin
        ErrorMsg( Format( '同步發票合併資料檔錯誤, 訊息:%s。', [aMsg] ) );
        Abort;
      end;
    end else
    begin

      if not SyncInv008ADataSet( dmInsert,
        aNewSeq,  EmptyStr, FInvoiceId, aItemIdRefId, aMsg ) then
      begin
        ErrorMsg( Format( '同步發票合併資料檔錯誤, 訊息:%s。', [aMsg] ) );
        Abort;
      end;
    end;
    //ChkCycleMerge( aItemIdRefId,aNewSeq,cdsDetail.RecNo );
  end else
  begin
    if cdsInv008A.Locate('ITEMIDREF',cdsDetail.FieldByName('ITEMID').AsVariant,[]) then
    begin
      aItemIdRefId := cdsInv008A.FieldByName( 'ITEMIDREF').AsString;
      aNewSeq := cdsInv008A.FieldByName( 'SEQ' ).AsString;
      cdsDetail.FieldByName( 'ITEMIDREF' ).AsString := aItemIdRefId;
      cdsDetail.FieldByName( 'MERGESEQ' ).AsString := aNewSeq;
      if not SyncInv008ADataSet( dmInsert,
        aNewSeq,  EmptyStr, FInvoiceId, aItemIdRefId, aMsg ) then
      begin
        ErrorMsg( Format( '同步發票合併資料檔錯誤, 訊息:%s。', [aMsg] ) );
        Abort;
      end;
    end;
  end;

end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.cdsDetailAfterPost(DataSet: TDataSet);
begin
  UpdatePriceToMaster;
  ComposeBillingString( FBillingInfo );
  if cdsDetail.FieldByName( 'ITEMIDREF' ).AsString <> '' then
  begin
    ChkCycleMerge( cdsDetail.FieldByName('ITEMIDREF').AsString,
      cdsDetail.FieldByName('MERGESEQ').AsString,
      cdsDetail.RecNo );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA02_4.DetailGridViewBILLIDPropertiesButtonClick(
  Sender: TObject; AButtonIndex: Integer);
var
  aBillId, aItemId: String;
begin
  if ( cdsMaster.FieldByName( 'CUSTID' ).AsString = EmptyStr ) then
  begin
    WarningMsg( '請先輸入客編。' );
    Exit;
  end;
  // 現場開立, 預開, 後開的發票, 客編來源為 SMS 的收費資料
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
  // 非預開, 後開, 現場開立, 不須先輸入客編
  if ( FInvCreateType in [icPrev, icAfter, icLocale] ) then
  begin
    Error := ( cdsMaster.FieldByName( 'CUSTID' ).AsString = EmptyStr );
    if Error then
    begin
      WarningMsg( '請先輸入客編。' );
      ErrorText := '';
      Abort;
    end;
  end;  
  { 與客服系統介接才檢查 }
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
  // 非預開, 後開, 現場開立, 不須先輸入客編
  if ( FInvCreateType in [icPrev, icAfter, icLocale] ) then
  begin
    Error := ( cdsMaster.FieldByName( 'CUSTID' ).AsString = EmptyStr );
    if Error then
    begin
      WarningMsg( '請先輸入客編。' );
      ErrorText := '';
      Abort;
    end;
  end;  
  { 與客服系統介接才檢查 }
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
      WarningMsg( Format( '收費單號: %s,  無此收費項次。',
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
    WarningMsg( '數量不可為零。' );
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
  { 不允許輸入非數字資料 }
  if not TryStrToInt( DisplayValue, aInt ) then
  begin
    WarningMsg( '收費項目代碼不可輸入非數值資料。' );
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
  { 非應稅, 則把輸入的稅額一律改為 0 }
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
  // 預開, 後開, 現場開立, 須判斷是否要將收據單號關掉, 不可輸入
  // 一般開立, 開放可以都可以輸入
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
  // 預開, 後開, 現場開立, 須判斷是否要將收據單號關掉, 不可輸入
  // 一般開立, 不須輸入收費單號, 一律不可輸入
  // 將不可輸入欄位畫成灰色
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
    { ChargeSign 為負項的, 將字體轉成紅色 }
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
  aDataSet.MaxRecords := 0;
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
//  ChooseItemView.ItemCount
//  ChooseItemView.ViewData.RecordCount
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
  { Reload 資料} 
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
  if ( aResult = mrOk ) then { Reload 資料} FormShow( Self );
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
    cdsDetail.FieldByName( 'LINKTOMISDESC' ).AsString = '是', 'Y', 'N' );
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
  const aOldSeq, aNewSeq, aNewInvId, aItemIdRef : String;  var aMsg: String): Boolean;

  { -------------------------------------------------- }

  procedure DoAppend;
  var
    aStartDate, aEndDate: String;
    aAccountNo, aFacisNo, aCustId,aSmartCard: String;
    aBillNo,aBillNoItem : String;
  begin

    { 匯檔時 日期格式可能是錯的, 會導至錯誤 }
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
    { 取發票號碼跟CP門號 }
    if ( FInvCreateType in [icPrev, icAfter, icLocale] ) and
       ( dtmMain.GetLinkToMIS ) and
       ( cdsDetail.FieldByName( 'LINKTOMIS' ).AsString = 'Y' ) then
    begin
      // 在 Focus 還在 txtCustId 尚未離開時, 此時 MasterDataSet 的 CustId 欄位
      //  還是舊值或是空的, 這時直接取 txtCustId.Text
      aCustId := cdsMaster.FieldByName( 'CUSTID' ).AsString;
      if ( aCustId = EmptyStr ) or ( aCustId <> txtCUSTID.Text ) then
         aCustId := txtCUSTID.Text;
      {#5668 增加存入 SmartCard By Kin 2010/07/09}
      dtmSO.GetMisAccountAndFacisNo( cdsDetail.FieldByName( 'BILLID' ).AsString,
        cdsDetail.FieldByName( 'BILLIDITEMNO' ).AsString, aCustId,
         aAccountNo, aFacisNo, aSmartCard );
    end;
    {}

    if not cdsInv008A.Locate('ITEMIDREF;BILLID;BILLIDITEMNO;ITEMID',VarArrayOf(
       [aItemIdRef,cdsDetail.FieldByName( 'BILLID' ).AsString,
       cdsDetail.FieldByName( 'BILLIDITEMNO' ).AsString,
       cdsDetail.FieldByName( 'ITEMID' ).AsString]),[]) then
    begin
      {#5668 增加存入 SmartCard By Kin 2010/07/09}
      aBillNo := cdsDetail.FieldByName( 'BILLID' ).AsString;
      aBillNoItem := cdsDetail.FieldByName( 'BILLIDITEMNO' ).AsString;
      if aBillNo = EmptyStr then
        aBillNo := 'X';
      if aBillNoItem = EmptyStr then
        aBillNoItem := '-99';
      cdsInv008A.AppendRecord( [
      aItemIdRef,
      FInvoiceId,
      aBillNo,
      aBillNoItem,
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
      aAccountNo,
      cdsDetail.FieldByName( 'SMARTCARDNO' ).AsString ] );

      {
      cdsInv008A.AppendRecord( [
        aItemIdRef,
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
        aAccountNo,
        cdsDetail.FieldByName( 'SMARTCARDNO' ).AsString ] );
      }
    end else
    begin
      cdsInv008A.Edit;
      cdsInv008A.FieldByName( 'QUANTITY' ).AsString :=
        cdsDetail.FieldByName( 'QUANTITY' ).AsString;
      cdsInv008A.FieldByName( 'UNITPRICE' ).AsString :=
        cdsDetail.FieldByName( 'UNITPRICE' ).AsString;
      cdsInv008A.FieldByName( 'TAXAMOUNT' ).AsString :=
        cdsDetail.FieldByName( 'TAXAMOUNT' ).AsString;
      cdsInv008A.FieldByName( 'SALEAMOUNT' ).AsString :=
        cdsDetail.FieldByName( 'SALEAMOUNT' ).AsString;
      cdsInv008A.FieldByName( 'TOTALAMOUNT' ).AsString :=
        cdsDetail.FieldByName( 'TOTALAMOUNT' ).AsString;
      cdsInv008A.Post;
    end;
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
      if ( FMode in [dmInsert] ) then
      begin
          if ( cdsInv008A.FieldByName( 'SEQ' ).AsString = aOldSeq ) and
            ( cdsInv008A.FieldByName( 'BILLID' ).AsString = cdsDetail.FieldByName( 'BILLID' ).AsString ) and
            ( cdsInv008A.FieldByName( 'BILLIDITEMNO' ).AsString  = cdsDetail.FieldByName( 'BILLIDITEMNO' ).AsString ) then
          begin
              cdsInv008A.Delete;
          end else
          begin
            cdsInv008A.Next;
          end;
      end else
      begin

        if ( cdsInv008A.FieldByName( 'SEQ' ).AsString = aOldSeq ) then
          cdsInv008A.Delete
        else
          cdsInv008A.Next;
      end;
    end;  
  end;

  { -------------------------------------------------- }

begin
  aMsg := EmptyStr;
  { 當明細資料異動時, 同步 Inv008A 資料, 用 OldSeq 對應     }
  { 此部驟只剩下當發票明細被刪除時, 若 INV008A 有合併項目時, 則刪除,
    存檔時會去更新客服系統, 新增及修改用不到 }
  if ( aMode in [dmInsert] ) then
  begin
    {
    if not cdsInv008A.Locate( 'SEQ', aOldSeq, [] ) then
    begin
      try
        DoAppend;
      except
        on E: Exception do aMsg := E.Message;
      end;
    end;
    }
    try
      DoAppend;
    except
      on E: Exception do aMsg := E.Message;
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

function TfrmInvA02_4.GetItemIdRef(const aItemId: String): String;
 var
  aDataSet: TADOQuery;
begin
  aDataSet := dtmMain.adoComm;
  aDataSet.Close;
  aDataSet.SQL.Text :=  Format(
    '  SELECT ItemIdRef From INV005            ' +
    ' WHERE ItemId=''%s'' AND COMPID = ''%s''  ' ,
    [ aItemId,dtmMain.getCompID ] );
  try
    aDataSet.Open;
    if aDataSet.Eof then
    begin
      Result := EmptyStr;
    end else
    begin
      Result := aDataSet.Fields[0].AsString;
    end;
  finally
    aDataSet.Close;
  end;


end;

function TfrmInvA02_4.GetItemDesc(const aItemId: String): String;
 var
  aDataSet: TADOQuery;
begin
  aDataSet := dtmMain.adoComm;
  aDataSet.Close;
  aDataSet.SQL.Text :=  Format(
    '  SELECT Description From INV005          ' +
    ' WHERE ItemId=''%s'' AND COMPID = ''%s''  ' ,
    [ aItemId,dtmMain.getCompID ] );

  try
    aDataSet.Open;
    if aDataSet.IsEmpty then
    begin
      Result := EmptyStr;
      if cdsDetail.FieldByName( 'COMBCITEMNAME' ).AsString <> '' then
      begin
        Result := cdsDetail.FieldByName( 'COMBCITEMNAME' ).AsString;
      end else
      begin
        Result := cdsDetail.FieldByName( 'DESCRIPTION' ).AsString
      end;
    end else
    begin
      Result := aDataSet.Fields[0].AsString;
    end;
  finally
    aDataSet.Close;
  end;
end;



function TfrmInvA02_4.ChkCreateNum(aInvCnt :Integer): Boolean;
  var aRcdCount : Integer;
   aBookMark : Pointer;
   aRcdNo : Integer;
   aItemIRef : String;
begin
  aBookMark := cdsDetail.GetBookmark;
  aRcdNo := cdsDetail.RecNo;
  aItemIRef := cdsDetail.FieldByName( 'ITEMIDREF' ).AsString;
  cdsDetail.First;
  aRcdCount := 1;
  Result := False;


  if dtmMain.GetAutoCreateNum > 0 then
  begin
    try
      if aItemIRef <> '' then
      begin
        while not cdsDetail.Eof do
        begin
          if cdsDetail.RecNo >= aRcdNo then Break;
          if cdsDetail.FieldByName( 'ITEMIDREF' ).AsString = aItemIRef then
          begin
            aRcdCount :=0;
            Break;
          end;
          cdsDetail.Next;
        end;
      end;
    finally
      cdsDetail.GotoBookmark( aBookMark );
      cdsDetail.FreeBookmark( aBookMark );
    end;
  end else
  begin
    cdsDetail.GotoBookmark( aBookMark );
    cdsDetail.FreeBookmark( aBookMark );
    Exit;
  end;
  Result := ( aInvCnt + aRcdCount ) > dtmMain.GetAutoCreateNum ;
end;

function TfrmInvA02_4.ChkChgInvNum(aCount: Integer): Boolean;
  var
    aItemIdRefId : string;
    aBookMark : Pointer;
begin
  Result := False;
  try
    aBookMark := cdsDetail.GetBookmark;
    if cdsDetail.RecNo > 1 then
    begin
      if cdsDetail.FieldByName( 'COMBCITEMCODE' ).AsString <> '' then
      begin
        aItemIdRefId :=  cdsDetail.FieldByName( 'COMBCITEMCODE' ).AsString ;
        if aItemIdRefId = EmptyStr then
        begin
          aItemIdRefId := GetItemIdRef(cdsDetail.FieldByName( 'COMBCITEMCODE' ).AsString);
        end;
      end else
      begin
        aItemIdRefId := GetItemIdRef( cdsDetail.FieldByName( 'ItemID' ).AsString );
      end;
       //代表有合併資料
      if  cdsInv008A.Locate('BILLID;BILLIDITEMNO',
                  VarArrayOf( [cdsDetail.FieldByName( 'BILLID' ).AsString,
                              cdsDetail.FieldByName( 'BILLIDITEMNO').AsString] ),[]) then
      begin
        if not cdsDetail.Locate('ITEMID',VarArrayOf([cdsInv008A.FieldByName('ITEMIDREF').AsString]),[]) then
        begin
          Inc(aCount);
        end;
      end else
      begin
        inc(aCount);
      end;
    end;

    Result := (aCount >= dtmMain.GetAutoCreateNum);
  finally
    cdsDetail.GotoBookmark( aBookMark );
    cdsDetail.FreeBookmark( aBookMark );
  end;

end;

function TfrmInvA02_4.UpdMergeInv008A(const aItemIdRef, aInvId: String;
  const aSEQ: Integer): Boolean;
begin
  Result := False;
  try
    cdsInv008A.DisableControls;
    cdsInv008A.First;
    while not cdsInv008A.Eof do
    begin
      if cdsInv008A.FieldByName( 'ITEMIDREF' ).AsString = aItemIdRef then
      begin
        cdsInv008A.Edit;
        cdsInv008A.FieldByName( 'INVID' ).AsString := aInvId;
        cdsInv008A.FieldByName( 'SEQ' ).AsInteger := aSEQ;
      end;
      cdsInv008A.Next;
    end;

  finally
    cdsInv008A.EnableControls;
  end;


end;
//計算此次要開幾張發票
function TfrmInvA02_4.CalMergeNum: Integer;
  var aRcdCount : Integer;
  aBookMark : Pointer;
begin
  aRcdCount := 0;
  aBookMark := cdsDetail.GetBookmark;
  cdsDetail.First;

  try
    while not cdsDetail.Eof do
    begin
      if cdsDetail.FieldByName( 'ITEMIDREF' ).AsString = '' then
      begin
        Inc(aRcdCount);
      end else
      begin

        if cdsDetail.FieldByName( 'SEQ' ).AsString =
          cdsDetail.FieldByName( 'MERGESEQ' ).AsString then
        begin
          Inc(aRcdCount);
        end;
      end;
      cdsDetail.Next;
    end;
  finally
    cdsDetail.GotoBookmark( aBookMark );
    cdsDetail.FreeBookmark( aBookMark );
  end;
  Result := aRcdCount;
end;

function TfrmInvA02_4.ChkHasMerge(const aRcdNo : Integer;
  const aItemIdRef : string) : Boolean;
  var aMerge: Boolean;
begin
  aMerge := False;
  cdsDetail.First;
  while not cdsDetail.Eof do
  begin
    if cdsDetail.RecNo >= aRcdNo then Break;
    if cdsDetail.FieldByName( 'ITEMIDREF' ).AsString = aItemIdRef then
    begin
      aMerge := True;
      Break;
    end;
    cdsDetail.Next;
  end;
  Result := aMerge;
end;
{ 再往前找INV008沒有被合併的資料 }
function TfrmInvA02_4.ChkCycleMerge(const aItemIdRef, aNewSeq : String;
     aRecNo :Integer): Boolean;
  var
    aBookMark : Pointer;
    aMsg : String;

begin
  Result := False;
  DisableDetailDataSetEvent;
  aBookMark := cdsDetail.GetBookmark;
  cdsDetail.First;
  try
    cdsDetail.DisableControls;
    while not cdsDetail.Eof do
    begin
      if cdsDetail.RecNo >= aRecNo then Break;
      if cdsDetail.FieldByName( 'ITEMIDREF' ).AsString = '' then
      begin
        if cdsDetail.FieldByName( 'ITEMID' ).AsString = aItemIdRef then
        begin
          Result := True;
          cdsDetail.Edit;
          cdsDetail.FieldByName( 'MERGESEQ' ).AsString := aNewSeq;
          cdsDetail.FieldByName( 'ITEMIDREF' ).AsString := aItemIdRef;
          cdsDetail.Post;
          if not SyncInv008ADataSet( dmInsert,
            aNewSeq,  EmptyStr, FInvoiceId, aItemIdRef, aMsg ) then
          begin
            ErrorMsg( Format( '同步發票合併資料檔錯誤, 訊息:%s。', [aMsg] ) );
            Abort;
          end;
        end;
      end;
      cdsDetail.Next;
    end;
  finally
    cdsDetail.GotoBookmark( aBookMark );
    cdsDetail.FreeBookmark( aBookMark );
    cdsDetail.EnableControls;
    EnableDetailDataSetEvent;
  end;

end;

procedure TfrmInvA02_4.txtGiveUnitCodeNamePropertiesInitPopup(
  Sender: TObject);
var aSQL : String;
begin
  aSQL := Format(
      ' SELECT ITEMID, DESCRIPTION   ' +
      '   FROM INV041                ' +
      '  WHERE IDENTIFYID1 = ''%S''  ' +
      '    AND IDENTIFYID2 = ''%S''  ' +
      '    AND COMPID = ''%S''       ' +
      '    AND ( REFNO <> 999 OR     ' +
      '          REFNO IS NULL )     ' +
      '    ORDER BY ITEMID           ',
      [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID] );
  txtGiveUnit.SQL.Text := aSQL;
  txtGiveUnit.CodeNamePropertiesInitPopup( Sender );
end;

procedure TfrmInvA02_4.btnLogClick(Sender: TObject);
  var aForm : TfrmInvA02_7;
   aInvid,aCompId : String;
begin
  Screen.Cursor := crSQLWait;
  ainvid :=cdsMaster.FieldByName( 'InvId' ).AsString;
  aCompId := cdsMaster.FieldByName( 'CompId' ).AsString;
  aForm := TfrmInvA02_7.Create( aInvid,aCompId );
  try
    aForm.ShowModal;
  finally
    aForm.Free;
    Screen.Cursor := crDefault;
  end;

end;

procedure TfrmInvA02_4.INV008ViewCellDblClick(
  Sender: TcxCustomGridTableView;
  ACellViewInfo: TcxGridTableDataCellViewInfo; AButton: TMouseButton;
  AShift: TShiftState; var AHandled: Boolean);
  var aItemIdColumn,aSeqColumn,aBillIdColumn,aBillIDItemNoColumn: TcxGridDBColumn;
      aItemIdValue,aSeqValue,aBillIdValue,aBillIDItemNoValue : String;
      aForm : TfrmInvA02_9;
begin
  aItemIdColumn := INV008View.GetColumnByFieldName( 'ItemId' );
  aSeqColumn := INV008View.GetColumnByFieldName( 'Seq' );
  aBillIdColumn := INV008View.GetColumnByFieldName( 'BILLID' );
  aBillIDItemNoColumn := INV008View.GetColumnByFieldName( 'BILLIDITEMNO' );
  aBillIdValue := VarToStrDef(ACellViewInfo.GridRecord.Values[aBillIdColumn.Index],'');
  aBillIDItemNoValue := VarToStrDef(ACellViewInfo.GridRecord.Values[aBillIDItemNoColumn.Index],'');
  aItemIdValue := VarToStr( ACellViewInfo.GridRecord.Values[aItemIdColumn.Index] );
  aSeqValue := VarToStr( ACellViewInfo.GridRecord.Values[aSeqColumn.Index] );
  if ( FMode in [dmInsert, dmUpdate] ) then
  begin
    aForm := TfrmInvA02_9.Create(cdsInv008A);
    try
      if cdsInv008A.Locate('ITEMID;SEQ;INVID;BILLID;BILLIDITEMNO',VarArrayOf([
              aItemIdValue,aSeqValue,
              cdsDetail.FieldByName('INVID').AsString,
              aBillIdValue,aBillIDItemNoValue]),[]) then
      begin
        aForm.ShowModal;
      end else
      begin
        WarningMsg( '尋找INV008A有誤！' );
      end;
    finally
      aForm.Free;
    end;
  end;
end;

procedure TfrmInvA02_4.btnPrintRecordClick(Sender: TObject);
 var aForm : TfrmInvA02_A;
   aInvid,aCompId : String;
begin
  ainvid :=cdsMaster.FieldByName( 'InvId' ).AsString;
  aCompId := cdsMaster.FieldByName( 'CompId' ).AsString;
  aForm := TfrmInvA02_A.Create( aInvid,aCompId );
  try
    aForm.ShowModal;
  finally
    aForm.Free;
  end;
end;

function TfrmInvA02_4.RanDomNum: String;
var
  aDataSet: TADOQuery;
begin
  aDataSet := dtmMain.adoComm2;
  aDataSet.Close;
  aDataSet.SQL.Text := 'select LPAD(round(dbms_random.value(1, 9999)),4,''0'') num from dual';
  aDataSet.Open;
  Result :=  aDataSet.FieldByName( 'num' ).AsString;
  aDataSet.Close;
end;



procedure TfrmInvA02_4.txtInvUse2CodeNamePropertiesInitPopup(
  Sender: TObject);
  var aSql : String;
begin
  {
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
  txtInvUse2.SQL.Text := aSql;
  txtInvUse2.CodeNamePropertiesInitPopup( Sender );
  }
end;



procedure TfrmInvA02_4.txtInvUseExit(Sender: TObject);

begin
  {
  blnCanEdit := False;

  try
    blnCanEdit := IsGiveCanEdit;

  finally

    txtLoveNum.Enabled := blnCanEdit;
    txtGiveUnit.Enabled := blnCanEdit;
  end;
  }
end;

function TfrmInvA02_4.IsGiveCanEdit: Boolean;
  var aSql :String;
    blnCanEdit : Boolean;
begin
  aSql := Format('select count(1) cnt from inv028   ' +
          ' where ItemId = %s and compid = ''%s''   ' +
          ' and refno = 1    ',
          [txtInvUse.Value,dtmMain.sG_CompID]);
  dtmMain.adoComm.Close;
  try
    if txtInvUse.Value <> EmptyStr then
    begin
      dtmMain.adoComm.SQL.Text := aSql;
      dtmMain.adoComm.Open;
      if dtmMain.adoComm.FieldByName('cnt').AsInteger > 0 then
        blnCanEdit := True
      else
        blnCanEdit := False;
    end;
  finally
    dtmMain.adoComm.Close;
  end;
  Result := blnCanEdit;

end;



procedure TfrmInvA02_4.OnInvUseSelectChange(sender: TObject);
  var
    blnCanEdit : Boolean;
begin
  {#6600 測試不OK,如果為捐贈愛心碼與受贈單位才可輸入 By Kin 2013/11/12}
  if ( FMode in [dmUpdate, dmInsert] ) then
  begin
    blnCanEdit := IsGiveCanEdit;
    txtLoveNum.Enabled := blnCanEdit;
    txtLoveNum.Properties.ReadOnly := not blnCanEdit;
    txtGiveUnit.Enabled := blnCanEdit;
    if not blnCanEdit then
    begin
      txtLoveNum.Clear;
      txtGiveUnit.Clear;
    end;
  end;
end;

procedure TfrmInvA02_4.btnUploadRecordClick(Sender: TObject);
var aForm : TfrmInvA02_B;
begin
  Screen.Cursor := crSQLWait;
  try
    FillUploadRecordDataSet;
    if ( cdsUploadRecord.IsEmpty ) then
    begin
      WarningMsg( '查無任何上傳記錄！');
    end else
    begin
      aForm := TfrmInvA02_B.Create( cdsUploadRecord );
      try
        aForm.ShowModal;
      finally
        aForm.Free;
      end;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TfrmInvA02_4.FillUploadRecordDataSet;
var
  aReader: TADOQuery;
  aSQL:string;
begin
  PrepareUploadRecordDataSet;
  aSQL := Format('select case uploadtype                 ' +
     ' when 1 then ''一般發票(C0401)''                   ' +
      ' when 2 then ''作廢發票(C0501)''                  ' +
      ' when 3 then ''折讓發票(D0401)''                  ' +
      ' when 4 then ''折讓作廢(D0501)''                  ' +
      ' when 5 then ''註銷發票(C0701)''                  ' +
      ' when 6 then ''E0402''                            ' +
      ' else ''未知'' end uploadtype,                    ' +
      ' uploadtime,                                      ' +
      ' Decode(uploadSource,1,''自動'',''手動'') uploadsource ' +
      ' from inv048                                      ' +
      ' where invid=''%s''                               ' +
      ' and compid = ''%s''                              ' +
      ' order by uploadtime                    ',
      [cdsMaster.FieldByName('Invid').AsString,
        cdsMaster.FieldByName('CompId').AsString]);
  aReader := dtmMain.adoComm2;
  aReader.Close;
  aReader.SQL.Text := aSQL;
  aReader.Open;
  aReader.First;

  while not aReader.Eof do
  begin
    cdsUploadRecord.AppendRecord( [
      aReader.FieldByName( 'uploadtype' ).AsString,
      aReader.FieldByName( 'uploadtime' ).AsString,
      aReader.FieldByName( 'uploadsource' ).AsString]
      );
    aReader.Next;
  end;
  aReader.Close;
end;

procedure TfrmInvA02_4.PrepareUploadRecordDataSet;
begin
  if ( cdsUploadRecord.FieldDefs.Count <= 0 ) then
  begin
    cdsUploadRecord.FieldDefs.Add( 'UPLOADTYPE', ftString, 30 );
    cdsUploadRecord.FieldDefs.Add( 'UPLOADTIME', ftString, 30 );
    cdsUploadRecord.FieldDefs.Add( 'UPLOADSOURCE', ftString, 15 );
    cdsUploadRecord.CreateDataSet;
  end;
  cdsUploadRecord.EmptyDataSet;
end;

end.
