unit frmCD078B3U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Menus, Buttons, DB, DBClient, ADODB, ActnList,
  frmCD078B0U, frmCD078B1U,
  cbDBController, cbStyleController, cbDataLookup,
  cxControls, cxContainer, cxEdit, cxTextEdit, cxGroupBox, cxCheckBox,
  cxMaskEdit, cxDropDownEdit, cxCurrencyEdit, cxGraphics, dxSkinsCore,
  dxSkinsDefaultPainters, dxSkinscxPCPainter, cxPC, cxStyles, 
  cxDataStorage, cxDBData, cxGridLevel, cxClasses,
  cxGridCustomView, cxGridCustomTableView, cxGridTableView, cxGridDBTableView,
  cxGrid, cxGridBandedTableView, cxGridDBBandedTableView,
  cxButtonEdit, cxLookAndFeelPainters, cxButtons, cxGridCommon,
  cxCustomData, cxFilter, cxData, dxSkinBlack, dxSkinBlue, dxSkinCaramel,
  dxSkinCoffee, Mask;

type
  TOldValue = record
    CItemCode: String;
    ServiceType: String;
    PenalCode: String;
  end;

  TCD019Data = record
    Sign: String;
    Amount: Double;
    Period: Integer;
    MonthAmt: Extended;
    DayAmt: Extended;
    IsMinus: Boolean;
    IsByHouse: Boolean;
    CMCode: String;
    CMName: String;
    RefNo: string;
  end;

  TfrmCD078B3 = class(TForm)
    Panel1: TPanel;
    Panel3: TPanel;
    edtDML: TcxTextEdit;
    Panel4: TPanel;
    Bevel1: TBevel;
    ActionList1: TActionList;
    actSave: TAction;
    actCancel: TAction;
    btnSave: TButton;
    Button1: TButton;
    cxGroupBox3: TcxGroupBox;
    chkPenal: TcxCheckBox;
    Label17: TLabel;
    Penal: TDataLookup;
    cxGroupBox4: TcxGroupBox;
    btnUpdate: TButton;
    actUpdate: TAction;
    pnlRank: TPanel;
    Label23: TLabel;
    Label24: TLabel;
    Label26: TLabel;
    lblMon1: TLabel;
    lblMon2: TLabel;
    Label32: TLabel;
    lblMon3: TLabel;
    Label35: TLabel;
    Label37: TLabel;
    lblMon4: TLabel;
    lblMon5: TLabel;
    Label43: TLabel;
    lblMon6: TLabel;
    Label45: TLabel;
    Label46: TLabel;
    lblMon7: TLabel;
    lblMon8: TLabel;
    Label49: TLabel;
    lblMon9: TLabel;
    Label51: TLabel;
    Label52: TLabel;
    lblMon10: TLabel;
    lblMon11: TLabel;
    Label55: TLabel;
    lblMon12: TLabel;
    Label57: TLabel;
    Label58: TLabel;
    Label61: TLabel;
    Label62: TLabel;
    Label63: TLabel;
    Label30: TLabel;
    edtMon1: TcxMaskEdit;
    edtMon2: TcxMaskEdit;
    edtMon3: TcxMaskEdit;
    edtMon4: TcxMaskEdit;
    edtMon5: TcxMaskEdit;
    edtMon6: TcxMaskEdit;
    edtMon7: TcxMaskEdit;
    edtMon8: TcxMaskEdit;
    edtMon9: TcxMaskEdit;
    edtMon10: TcxMaskEdit;
    edtMon11: TcxMaskEdit;
    edtMon12: TcxMaskEdit;
    edtPeriod1: TcxMaskEdit;
    edtPeriod2: TcxMaskEdit;
    edtPeriod3: TcxMaskEdit;
    edtPeriod4: TcxMaskEdit;
    edtPeriod5: TcxMaskEdit;
    edtPeriod6: TcxMaskEdit;
    edtPeriod7: TcxMaskEdit;
    edtPeriod8: TcxMaskEdit;
    edtPeriod9: TcxMaskEdit;
    edtPeriod10: TcxMaskEdit;
    edtPeriod11: TcxMaskEdit;
    edtPeriod12: TcxMaskEdit;
    cmbRateType1: TcxComboBox;
    cmbRateType2: TcxComboBox;
    cmbRateType3: TcxComboBox;
    cmbRateType4: TcxComboBox;
    cmbRateType5: TcxComboBox;
    cmbRateType6: TcxComboBox;
    cmbRateType7: TcxComboBox;
    cmbRateType8: TcxComboBox;
    cmbRateType9: TcxComboBox;
    cmbRateType10: TcxComboBox;
    cmbRateType11: TcxComboBox;
    cmbRateType12: TcxComboBox;
    edtDiscountAmt1: TcxCurrencyEdit;
    edtMonthAmt1: TcxCurrencyEdit;
    edtDayAmt1: TcxCurrencyEdit;
    edtDiscountAmt2: TcxCurrencyEdit;
    edtMonthAmt2: TcxCurrencyEdit;
    edtDayAmt2: TcxCurrencyEdit;
    edtDiscountAmt3: TcxCurrencyEdit;
    edtMonthAmt3: TcxCurrencyEdit;
    edtDayAmt3: TcxCurrencyEdit;
    edtDiscountAmt4: TcxCurrencyEdit;
    edtMonthAmt4: TcxCurrencyEdit;
    edtDayAmt4: TcxCurrencyEdit;
    edtDiscountAmt5: TcxCurrencyEdit;
    edtMonthAmt5: TcxCurrencyEdit;
    edtDayAmt5: TcxCurrencyEdit;
    edtDiscountAmt6: TcxCurrencyEdit;
    edtMonthAmt6: TcxCurrencyEdit;
    edtDayAmt6: TcxCurrencyEdit;
    edtDiscountAmt7: TcxCurrencyEdit;
    edtMonthAmt7: TcxCurrencyEdit;
    edtDayAmt7: TcxCurrencyEdit;
    edtDiscountAmt8: TcxCurrencyEdit;
    edtMonthAmt8: TcxCurrencyEdit;
    edtDayAmt8: TcxCurrencyEdit;
    edtDiscountAmt9: TcxCurrencyEdit;
    edtMonthAmt9: TcxCurrencyEdit;
    edtDayAmt9: TcxCurrencyEdit;
    edtDiscountAmt10: TcxCurrencyEdit;
    edtMonthAmt10: TcxCurrencyEdit;
    edtDayAmt10: TcxCurrencyEdit;
    edtDiscountAmt11: TcxCurrencyEdit;
    edtMonthAmt11: TcxCurrencyEdit;
    edtDayAmt11: TcxCurrencyEdit;
    edtDiscountAmt12: TcxCurrencyEdit;
    edtMonthAmt12: TcxCurrencyEdit;
    edtDayAmt12: TcxCurrencyEdit;
    edtStepNo: TcxTextEdit;
    chkMasterSale: TcxCheckBox;
    edtDescription: TcxTextEdit;
    btnRankList: TBitBtn;
    gvPenal: TcxGridDBTableView;
    glPenal: TcxGridLevel;
    PenalGrid: TcxGrid;
    gvPenalCol1: TcxGridDBColumn;
    gvPenalCol2: TcxGridDBColumn;
    gvPenalCol3: TcxGridDBColumn;
    dsCD078A3: TDataSource;
    btnPenalCondition: TBitBtn;
    Page1: TcxPageControl;
    cxTabSheet1: TcxTabSheet;
    cxTabSheet2: TcxTabSheet;
    cxGroupBox1: TcxGroupBox;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    chkBundle: TcxCheckBox;
    edtBundleMon: TcxMaskEdit;
    cmbPenalType: TcxComboBox;
    cmbExpireType: TcxComboBox;
    cmbDepositAttr: TcxComboBox;
    btnDepositAmt: TBitBtn;
    cxGroupBox2: TcxGroupBox;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    edtMonthAmt: TcxCurrencyEdit;
    edtDayAmt: TcxCurrencyEdit;
    cmbPFlag1: TcxComboBox;
    edtPFlagDate: TcxMaskEdit;
    PFlag: TDataLookup;
    edtTruncAmt: TcxCurrencyEdit;
    btnInstCodeStr: TButton;
    cxButton1: TcxButton;
    edtInstCodeStr: TcxTextEdit;
    Label3: TLabel;
    Label18: TLabel;
    FaciItem: TDataLookup;
    MCItem: TDataLookup;
    Label1: TLabel;
    Label2: TLabel;
    CItem: TDataLookup;
    ServiceType: TDataLookup;
    btnChangeCItemCode: TBitBtn;
    Label9: TLabel;
    edtCMBaudNo: TcxTextEdit;
    edtCMBaudRate: TcxTextEdit;
    Bevel2: TBevel;
    edtMinCount: TcxMaskEdit;
    lblMinCount: TLabel;
    edtMaxCount: TcxMaskEdit;
    lblMaxCount: TLabel;
    MSalePoint: TDataLookup;
    Label8: TLabel;
    btnRankWhere: TBitBtn;
    edtBpStopDate: TMaskEdit;
    Label19: TLabel;
    Label20: TLabel;
    cmbPFlag2: TcxComboBox;
    chkStopFlag: TcxCheckBox;
    chkLONGPAYflag: TcxCheckBox;
    lblContractType: TLabel;
    cmbContractType: TcxComboBox;
    lbl1: TLabel;
    edtAuthTunerCount: TcxMaskEdit;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
    procedure actSaveExecute(Sender: TObject);
    procedure actCancelExecute(Sender: TObject);
    procedure CItemCodeNamePropertiesInitPopup(Sender: TObject);
    procedure ServiceTypeCodeNamePropertiesInitPopup(Sender: TObject);
    procedure CItemCodeNoExit(Sender: TObject);
    procedure CItemCodeNamePropertiesChange(Sender: TObject);
    procedure FaciItemCodeNamePropertiesInitPopup(Sender: TObject);
    procedure FaciItemCodeNoPropertiesChange(Sender: TObject);
    procedure ServiceTypeCodeNoPropertiesChange(Sender: TObject);
    procedure ServiceTypeCodeNamePropertiesChange(Sender: TObject);
    procedure chkPenalPropertiesChange(Sender: TObject);
    procedure MonPropertiesChange(Sender: TObject);
    procedure PenalCodeNamePropertiesInitPopup(Sender: TObject);
    procedure cmbPFlag1PropertiesChange(Sender: TObject);
    procedure PFlagCodeNamePropertiesInitPopup(Sender: TObject);
    procedure CItemCodeNoPropertiesChange(Sender: TObject);
    procedure actUpdateExecute(Sender: TObject);
    procedure cmbDepositAttrPropertiesChange(Sender: TObject);
    procedure btnRankListClick(Sender: TObject);
    procedure CD078A3DataSetNewRecord(DataSet: TDataSet);
    procedure gvPenalCustomDrawCell(Sender: TcxCustomGridTableView;
      ACanvas: TcxCanvas; AViewInfo: TcxGridTableDataCellViewInfo;
      var ADone: Boolean);
    procedure gvPenalCol3PropertiesValidate(Sender: TObject;
      var DisplayValue: Variant; var ErrorText: TCaption; var Error: Boolean);
    procedure btnDepositAmtClick(Sender: TObject);
    procedure gvPenalCol2PropertiesValidate(Sender: TObject;
      var DisplayValue: Variant; var ErrorText: TCaption; var Error: Boolean);
    procedure btnInstCodeStrClick(Sender: TObject);
    procedure cxButton1Click(Sender: TObject);
    procedure btnChangeCItemCodeClick(Sender: TObject);
    procedure btnPenalConditionClick(Sender: TObject);
    procedure MCItemCodeNamePropertiesInitPopup(Sender: TObject);
    procedure MCItemCodeNoPropertiesChange(Sender: TObject);
    procedure MCItemCodeNamePropertiesChange(Sender: TObject);
    procedure FaciItemCodeNamePropertiesChange(Sender: TObject);
    procedure Page1DrawTabEx(AControl: TcxCustomTabControl; ATab: TcxTab;
      Font: TFont);
    procedure edtMonthAmtPropertiesEditValueChanged(Sender: TObject);
    procedure MSalePointCodeNamePropertiesChange(Sender: TObject);
    procedure MSalePointCodeNamePropertiesInitPopup(Sender: TObject);
    procedure MSalePointCodeNoPropertiesChange(Sender: TObject);
    procedure btnRankWhereClick(Sender: TObject);
    procedure edtBpStopDateExit(Sender: TObject);
  private
    { Private declarations }
    FMode: TDML;
    FKeyValue: String;
    FCanEditLongPayFlag : Boolean;
    FDataSet: TClientDataSet;
    FReader: TADOQuery;
    FaciDataSet: TClientDataSet;
    FInstCodeDataSet: TClientDataSet;
    FMCItemDataSet: TClientDataSet;
    FLinkKeyDataSet: TClientDataSet;
    FLockCheck: Boolean;
    FLockCItemChange: Boolean;
    FOldValue: TOldValue;
    FCD019Data: TCD019Data;
    FQueryParam: TQueryParam;
    FAlreadySetCItemValue: String;
    FCD078A1DataSet: TClientDataSet;
    FCD078A2DataSet: TClientDataSet;
    FCD078A3DataSet: TClientDataSet;
    FCD078A4DataSet: TClientDataSet;
    procedure SetCD078A2DataSet(ADataSet: TClientDataSet);
    procedure SetCD078A3DataSet(ADataSet: TClientDataSet);
    procedure SetCD078A4DataSet(ADataSet: TClientDataSet);
    procedure InitControlArray;
    procedure ClearEditor;
    procedure DataToEditor;
    procedure DMLModeChange(aMode: TDML);
    procedure ButtonStateChange;
    procedure EditorStateChange;
    function DataSetStateChange(const aMode: TDML): Boolean;
    function DataValidate: Boolean;
    function EditorToData: Boolean;
    procedure LoadCitemRelationItem;
    procedure InitCD019Data;
    procedure LoadCD019Data;
    procedure CD019DataToFieldValue;
    procedure ChangeRelationItemState;
    procedure ChangeFaciItemState;
    procedure ChangeExpireTypeItem;
    procedure ResetEditorState;
    procedure InternalCalcMonth(aIndex: Integer);
    procedure AcquireDefaulDeposit(var ACode, AName: String);
    {}
    function VdMonthBreakAndInput: Boolean;
    function VdMustInput: Boolean;
    function VdBundle: Boolean;
    function VdDeposit: Boolean;
    function VdPFlagDate: Boolean;
    function VdPenal: Boolean;
    function VdCanShowMultiRank: Boolean;
    function VdCanShowMultiDeposit: Boolean;
    function VdCanShowPenalRule: Boolean;
    function VdCD078A1: Boolean;
    function VdCMBaudNo: Boolean;
    function VdConfirm: Boolean;
    {}
    procedure UpdateMasterSale;
    procedure UpdateMCItemDataSet(const ADeleteCItem: String);
    procedure SyncLongPayFlag;
    procedure UpdateCD078A1KeyValue;
    procedure UpdateCD078A2KeyValue;
    procedure UpdateCD078A3KeyValue;
    procedure UpdateCD078A4KeyValue;
    {}
    procedure FilterCD078A1DataWhenCurrent;
    procedure FilterCD078A2DataWhenCurrent;
    procedure FilterCD078A3DataWhenCurrent;
    procedure FilterCD078A4DataWhenCurrent;
    procedure FilterCD078A1DataWhenNew;
    procedure FilterCD078A2DataWhenNew;
    procedure FilterCD078A3DataWhenNew;
    procedure FilterCD078A4DataWhenNew;
    procedure FilterLinkKeyDataSet;
    {}
    procedure AddCD078A1DefaultRecord;
    procedure BuildCD078A3Record;
    procedure UpdateCD078A1LinkKey;
    procedure EmptyCD078A2;
    procedure EmptyCD078A3;
    procedure EmptyCD078A4;
  public
    { Public declarations }
    constructor Create(aMode: TDML; aKeyValue: String; aDataSet: TClientDataSet); reintroduce;
    property AlreadySetCItemValue: String read FAlreadySetCItemValue write FAlreadySetCItemValue;
    property FaciItemDataSet: TClientDataSet read FaciDataSet write FaciDataSet;
    property InstCodeDataSet: TClientDataSet read FInstCodeDataSet write FInstCodeDataSet;
    property MCItemDataSet: TClientDataSet read FMCItemDataSet write FMCItemDataSet;
    property LinkKeyDataSet: TClientDataSet read FLinkKeyDataSet write FLinkKeyDataSet;
    property CD078A1DataSet: TClientDataSet read FCD078A1DataSet write FCD078A1DataSet;
    property CD078A2DataSet:  TClientDataSet read FCD078A2DataSet write SetCD078A2DataSet;
    property CD078A3DataSet:  TClientDataSet read FCD078A3DataSet write SetCD078A3DataSet;
    property CD078A4DataSet:  TClientDataSet read FCD078A4DataSet write SetCD078A4DataSet;
    property CanEditLongPayFlag : Boolean read FCanEditLongPayFlag write FCanEditLongPayFlag;
  end;

var
  frmCD078B3: TfrmCD078B3;

implementation

uses
  cbUtilis, frmMultiSelectU, frmCD078B3_1U, frmCD078B3_3U,
  frmCD078B3_4U, frmCD078B3_5U, frmCD078B3_7U;

var
  FMon: array [1..12] of TcxMaskEdit;
  FMonLabel: array [1..12] of TLabel;
  FPeriod: array [1..12] of TcxMaskEdit;
  FRateType: array [1..12] of TcxComboBox;
  FDiscountAmt: array [1..12] of TcxCurrencyEdit;
  FMonthAmt: array [1..12] of TcxCurrencyEdit;
  FDayAmt: array [1..12] of TcxCurrencyEdit;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

function ItemIndexOfValue(const AValue: String; AList: TStrings): Integer;
var
  aIndex, aIndex2: Integer;
  aItemText: String;
begin
  Result := -1;
  if ( AList.Count <= 0 ) or ( AValue = EmptyStr ) then Exit;
  for aIndex := 0 to AList.Count - 1 do
  begin
    aItemText := EmptyStr;
    for aIndex2 := 1 to Length( AList[aIndex] ) do
    begin
      if ( IsDelimiter( '.:', AList[aIndex], aIndex2 ) ) then
      begin
        aItemText := Copy( AList[aIndex], 1, aIndex2 - 1 );
        Break;
      end;
    end;
    if AnsiSameText( AValue, aItemText ) then
    begin
      Result := aIndex;
      Break;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function ValueOfItemIndex(const AItemIndex: Integer; AList: TStrings): String;
var
  aIndex: Integer;
  aItemText: String;
begin
  Result := EmptyStr;
  if ( not AItemIndex in [0..AList.Count - 1] ) then Exit;
  aItemText := AList[AItemIndex];
  for aIndex := 1 to Length( aItemText ) do
  begin
    if ( IsDelimiter( '.:', aItemText, aIndex ) ) then
    begin
      Result := Copy( aItemText, 1, aIndex - 1 );
      Break;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

constructor TfrmCD078B3.Create(aMode: TDML; aKeyValue: String;
  aDataSet: TClientDataSet);
begin
  inherited Create( Application );
  FMode := aMode;
  FKeyValue := aKeyValue;
  FDataSet := aDataSet;
  FReader := DBController.CodeReader;
  FLockCheck := False;
  FLockCItemChange := False;
  FOldValue.CItemCode := EmptyStr;
  FOldValue.ServiceType := EmptyStr;
  FOldValue.PenalCode := EmptyStr;
  InitCD019Data;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.FormCreate(Sender: TObject);
begin
  CItem.Initializa;
  MCItem.Initializa;
  ServiceType.Initializa;
  FaciItem.Initializa;
  PFlag.Initializa;
  Penal.Initializa;
  MSalePoint.Initializa;
  InitControlArray;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.FormShow(Sender: TObject);
begin
  Screen.Cursor := crSQLWait;
  try
    FaciItem.DataSet := FaciDataSet;
    MCItem.DataSet := FMCItemDataSet;
    {#6841 ?W?[?P?_?????O???O?O?_?i?H???? By Kin 2014/08/05}
    chkLONGPAYflag.Enabled := FCanEditLongPayFlag;
    DMLModeChange( FMode );
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if ( FMode <> dmCancel ) then
    DMLModeChange( dmCancel );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.FormDestroy(Sender: TObject);
begin
  FReader := nil;
  FDataSet := nil;
  FCD078A1DataSet := nil;
  FCD078A2DataSet := nil;
  FCD078A3DataSet.OnNewRecord := nil;
  FCD078A3DataSet := nil;
  FaciDataSet.Filtered := False;
  FaciDataSet.Filter := EmptyStr;
  FaciDataSet := nil;
  CItem.Finaliza;
  MCItem.Finaliza;
  ServiceType.Finaliza;
  FaciItem.Finaliza;
  PFlag.Finaliza;
  Penal.Finaliza;
  MSalePoint.Finaliza;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.actSaveExecute(Sender: TObject);
begin
  DMLModeChange( dmSave );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.actCancelExecute(Sender: TObject);
begin
  DMLModeChange( dmCancel );
  Self.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.SetCD078A2DataSet(ADataSet: TClientDataSet);
begin
  FCD078A2DataSet := ADataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.SetCD078A3DataSet(ADataSet: TClientDataSet);
begin
  FCD078A3DataSet := ADataSet;
  dsCD078A3.DataSet := FCD078A3DataSet;
  FCD078A3DataSet.OnNewRecord := CD078A3DataSetNewRecord;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.SetCD078A4DataSet(ADataSet: TClientDataSet);
begin
  FCD078A4DataSet := ADataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.InitControlArray;
begin
  { ???q???u?f???? }
  FMon[1] := edtMon1;
  FMon[2] := edtMon2;
  FMon[3] := edtMon3;
  FMon[4] := edtMon4;
  FMon[5] := edtMon5;
  FMon[6] := edtMon6;
  FMon[7] := edtMon7;
  FMon[8] := edtMon8;
  FMon[9] := edtMon9;
  FMon[10] := edtMon10;
  FMon[11] := edtMon11;
  FMon[12] := edtMon12;
  { ???????? - ????}
  FMonLabel[1] := lblMon1;
  FMonLabel[2] := lblMon2;
  FMonLabel[3] := lblMon3;
  FMonLabel[4] := lblMon4;
  FMonLabel[5] := lblMon5;
  FMonLabel[6] := lblMon6;
  FMonLabel[7] := lblMon7;
  FMonLabel[8] := lblMon8;
  FMonLabel[9] := lblMon9;
  FMonLabel[10] := lblMon10;
  FMonLabel[11] := lblMon11;
  FMonLabel[12] := lblMon12;
  { ???O???? }
  FPeriod[1] := edtPeriod1;
  FPeriod[2] := edtPeriod2;
  FPeriod[3] := edtPeriod3;
  FPeriod[4] := edtPeriod4;
  FPeriod[5] := edtPeriod5;
  FPeriod[6] := edtPeriod6;
  FPeriod[7] := edtPeriod7;
  FPeriod[8] := edtPeriod8;
  FPeriod[9] := edtPeriod9;
  FPeriod[10] := edtPeriod10;
  FPeriod[11] := edtPeriod11;
  FPeriod[12] := edtPeriod12;
  { ?O?v???? }
  FRateType[1] := cmbRateType1;
  FRateType[2] := cmbRateType2;
  FRateType[3] := cmbRateType3;
  FRateType[4] := cmbRateType4;
  FRateType[5] := cmbRateType5;
  FRateType[6] := cmbRateType6;
  FRateType[7] := cmbRateType7;
  FRateType[8] := cmbRateType8;
  FRateType[9] := cmbRateType9;
  FRateType[10] := cmbRateType10;
  FRateType[11] := cmbRateType11;
  FRateType[12] := cmbRateType12;
  { ?u?f???B }
  FDiscountAmt[1] := edtDiscountAmt1;
  FDiscountAmt[2] := edtDiscountAmt2;
  FDiscountAmt[3] := edtDiscountAmt3;
  FDiscountAmt[4] := edtDiscountAmt4;
  FDiscountAmt[5] := edtDiscountAmt5;
  FDiscountAmt[6] := edtDiscountAmt6;
  FDiscountAmt[7] := edtDiscountAmt7;
  FDiscountAmt[8] := edtDiscountAmt8;
  FDiscountAmt[9] := edtDiscountAmt9;
  FDiscountAmt[10] := edtDiscountAmt10;
  FDiscountAmt[11] := edtDiscountAmt11;
  FDiscountAmt[12] := edtDiscountAmt12;
  { ???????B }
  FMonthAmt[1] := edtMonthAmt1;
  FMonthAmt[2] := edtMonthAmt2;
  FMonthAmt[3] := edtMonthAmt3;
  FMonthAmt[4] := edtMonthAmt4;
  FMonthAmt[5] := edtMonthAmt5;
  FMonthAmt[6] := edtMonthAmt6;
  FMonthAmt[7] := edtMonthAmt7;
  FMonthAmt[8] := edtMonthAmt8;
  FMonthAmt[9] := edtMonthAmt9;
  FMonthAmt[10] := edtMonthAmt10;
  FMonthAmt[11] := edtMonthAmt11;
  FMonthAmt[12] := edtMonthAmt12;
  { ???????B }
  FDayAmt[1] := edtDayAmt1;
  FDayAmt[2] := edtDayAmt2;
  FDayAmt[3] := edtDayAmt3;
  FDayAmt[4] := edtDayAmt4;
  FDayAmt[5] := edtDayAmt5;
  FDayAmt[6] := edtDayAmt6;
  FDayAmt[7] := edtDayAmt7;
  FDayAmt[8] := edtDayAmt8;
  FDayAmt[9] := edtDayAmt9;
  FDayAmt[10] := edtDayAmt10;
  FDayAmt[11] := edtDayAmt11;
  FDayAmt[12] := edtDayAmt12;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.ClearEditor;
var
  aIndex: Integer;
begin
  FLockCheck := True;
  try
    CItem.Clear;
    MCItem.Clear;
    ServiceType.Clear;
    Label3.Font.Color := clDefault;
    FaciItem.Clear;
    edtCMBaudNo.Clear;
    edtCMBaudRate.Clear;
    { ?j?????? }
    chkBundle.Checked := False;
    {#6700 ?W?[???? By Kin 2014/1/3}
    chkStopFlag.Checked := False;
    edtBundleMon.Text := '0';
    cmbPenalType.ItemIndex := 0;
    cmbExpireType.ItemIndex := 0;
    { ?O???? }
    cmbDepositAttr.ItemIndex := 0;
    { ?P???P?? }
    edtMonthAmt.Value := 0;
    edtDayAmt.Value := 0;
    cmbPFlag1.Properties.OnChange := nil;

    { #4435 ???j?x???P???p?x?? By Kin 2009/03/24}
    edtMinCount.Text :='0';
    edtMaxCount.Text :='0';
    MSalePoint.Clear;
    //#5913 ?W?[PFlag2 By Kin 2011/04/14
    try
      cmbPFlag1.ItemIndex := 0;
      cmbPFlag2.ItemIndex := 0;
    finally
      cmbPFlag1.Properties.OnChange := cmbPFlag1PropertiesChange;
    end;
    edtPFlagDate.Text := '0';
    edtTruncAmt.Value := 0;
    PFlag.Clear;
    { ???q???H?? }
    chkPenal.Properties.OnChange := nil;
    try
      chkPenal.Checked := False;
    finally
      chkPenal.Properties.OnChange := chkPenalPropertiesChange;
    end;
    Penal.Clear;
    { ?D?????q???u?f }
    edtStepNo.Clear;
    edtDescription.Clear;
    chkMasterSale.Checked := False;
    for aIndex := 1 to 12 do
    begin
      FMon[aIndex].Properties.OnChange := nil;
      try
        FMon[aIndex].Text := EmptyStr;
      finally
        FMon[aIndex].Properties.OnChange := MonPropertiesChange;
      end;
      FMonLabel[aIndex].Caption := EmptyStr;
      FPeriod[aIndex].Text := EmptyStr;
      FRateType[aIndex].Text := EmptyStr;
      FDiscountAmt[aIndex].Value := 0;
      FMonthAmt[aIndex].Value := 0;
      FDayAmt[aIndex].Value := 0;
    end;
    {}
    cxButton1.Click;
    {}
    FOldValue.CItemCode := EmptyStr;
    FOldValue.ServiceType := EmptyStr;
    FOldValue.PenalCode := EmptyStr;
    {}
    InitCD019Data;
    {}
    Page1.ActivePageIndex := 0;
  finally
    FLockCheck := False;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.DataToEditor;
begin
  FLockCheck := True;
  try
    ServiceType.Value := FDataSet.FieldByName( 'SERVICETYPE' ).AsString;
    CItem.Value := FDataSet.FieldByName( 'CITEMCODE' ).AsString;
    CItem.ValueName := FDataSet.FieldByName( 'CITEMNAME' ).AsString;
    MCItem.Value := FDataSet.FieldByName( 'MCITEMCODE' ).AsString;
    FaciItem.Value := FDataSet.FieldByName( 'FACIITEM' ).AsString;
    {#5318 ?W?[?I?????P???k By Kin 2009/10/06}
    MSalePoint.Value := FDataSet.FieldByName( 'SALEPOINTCODE' ).AsString;
    MSalePoint.ValueName := FDataSet.FieldByName( 'SALEPOINTCODE' ).AsString;
    chkStopFlag.Checked := ( FDataSet.FieldByName( 'STOPFLAG' ).AsString = '1' );
    {}
    InitCD019Data;
    LoadCD019Data;
    CD019DataToFieldValue;
    ChangeRelationItemState;
    ChangeFaciItemState;
    ChangeExpireTypeItem;
    {}
    edtCMBaudNo.Text := FDataSet.FieldByName( 'CMBAUDNO' ).AsString;
    edtCMBaudRate.Text := FDataSet.FieldByName( 'CMBAUDRATE' ).AsString;
    {}
    edtMonthAmt.Value := FDataSet.FieldByName( 'MONTHAMT' ).AsInteger;
    edtDayAmt.Value := FDataSet.FieldByName( 'DAYAMT' ).AsFloat;
    {}
    chkBundle.Checked := ( FDataSet.FieldByName( 'BUNDLE' ).AsString = '1' );
    {#6700 ?W?[StopFlag By Kin 2014/1/3}
    chkStopFlag.Checked := ( FDataSet.FieldByName( 'StopFlag' ).AsString = '1' );
    edtBundleMon.Text := FDataSet.FieldByName( 'BUNDLEMON' ).AsString;
    {}
    //#5913 ?W?[ PFLAG2 By Kin 2011/04/14
    //#6216 ?????p?G?O2?n?NItemIndex?]??1???M?|????0 By Kin 2012/02/03
    if FDataSet.FieldByName( 'PFLAG2' ).AsInteger = 2 then
      cmbPFlag2.ItemIndex := 1
    else
      cmbPFlag2.ItemIndex := 0;
//    cmbPFlag2.ItemIndex := FDataSet.FieldByName( 'PFLAG2' ).AsInteger-1;

    {}
    if ( FDataSet.FieldByName( 'PENALTYPE' ).AsString = '1' ) or
       ( FDataSet.FieldByName( 'PENALTYPE' ).AsString = '2' ) then
      cmbPenalType.ItemIndex := FDataSet.FieldByName( 'PENALTYPE' ).AsInteger;
    {#6865 ?W?[?@??4.???????B?A???HIndex?n?T?w???b3 By Kin 2014/09/09}
    if ( FDataSet.FieldByName( 'PENALTYPE' ).AsString = '4' )  then
      cmbPenalType.ItemIndex := 3;
    {}
    if ( FDataSet.FieldByName( 'EXPIRETYPE' ).AsString = '1' ) or
       ( FDataSet.FieldByName( 'EXPIRETYPE' ).AsString = '2' ) or
       ( FDataSet.FieldByName( 'EXPIRETYPE' ).AsString = '3' ) or
       ( FDataSet.FieldByName( 'EXPIRETYPE' ).AsString = '5' ) then
    cmbExpireType.ItemIndex := ItemIndexOfValue(
      FDataSet.FieldByName( 'EXPIRETYPE' ).AsString, cmbExpireType.Properties.Items );
    {}
    if ( FDataSet.FieldByName( 'DEPOSITATTR' ).AsString = '1' ) or
       ( FDataSet.FieldByName( 'DEPOSITATTR' ).AsString = '2' ) then
    cmbDepositAttr.ItemIndex := FDataSet.FieldByName( 'DEPOSITATTR' ).AsInteger;
    {}
    if ( FDataSet.FieldByName( 'PFLAG1' ).AsString = '0' ) or
       ( FDataSet.FieldByName( 'PFLAG1' ).AsString = '1' ) or
       ( FDataSet.FieldByName( 'PFLAG1' ).AsString = '2' ) or
       ( FDataSet.FieldByName( 'PFLAG1' ).AsString = '3' ) then
    cmbPFlag1.ItemIndex := FDataSet.FieldByName( 'PFLAG1' ).AsInteger + 1;
    {}


    edtPFlagDate.Text := FDataSet.FieldByName( 'PFLAGDATE' ).AsString;
    edtTruncAmt.Value := FDataSet.FieldByName( 'TRUNCAMT' ).AsInteger;
    {}
    PFlag.Value := FDataSet.FieldByName( 'PFLAGCODE' ).AsString;
    PFlag.ValueName := FDataSet.FieldByName( 'PFLAGNAME' ).AsString;
    {}
    FQueryParam.Param1 := FDataSet.FieldByName( 'INSTCODESTR' ).AsString;
    edtInstCodeStr.Text := DBController.GetInstName( FQueryParam.Param1 );
    {}
    chkPenal.Checked := ( FDataSet.FieldByName( 'PENAL' ).AsString = '1' );
    Penal.Value := FDataSet.FieldByName( 'PENALCODE' ).AsString;
    Penal.ValueName := FDataSet.FieldByName( 'PENALNAME' ).AsString;
    {}
    { #4435 ?N?s?W?????????? By Kin 2009/03/24 }
    edtMinCount.Text := FDataSet.FieldByName( 'MinCount' ).AsString;
    edtMaxCount.Text := FDataSet.FieldByName( 'MaxCount' ).AsString;
    {}
    {#5257 ?W?[?S?w?u?f?I???? By Kin 2010/05/05}

    if FDataSet.FieldByName( 'BPStopDate' ).AsString <> EmptyStr then
    begin
      edtBpStopDate.EditText := FDataSet.FieldByName( 'BPStopDate' ).AsString;
    end;
    {# ?W?[ ContractType By Kin 2015/04/21}
    cmbContractType.ItemIndex := -1;
    if ( FDataSet.FieldByName( 'ContractType' ).AsString <> EmptyStr ) then
    begin
      if  FDataSet.FieldByName( 'ContractType' ).AsInteger <=2 then
      begin
        cmbContractType.ItemIndex := FDataSet.FieldByName( 'ContractType' ).AsInteger;
      end;
    end;

    {# ?W?[ AuthTunerCount By Kin 2015/04/21}
    edtAuthTunerCount.Text := '0';
    if FDataSet.FieldByName( 'AuthTunerCount' ).AsString <> EmptyStr then
      edtAuthTunerCount.Text := FDataSet.FieldByName( 'AuthTunerCount' ).AsString;
    {}
    {#6841 ?W?[?????O???O By Kin 2014/08/05}
    chkLONGPAYflag.Checked := False;
    if FDataSet.FieldByName( 'LONGPAYFLAG' ).AsInteger = 1 then
      chkLONGPAYflag.Checked := True;
    { ???q???u?f???? }
    UpdateMasterSale;
    FOldValue.CItemCode := CItem.Value;
    FOldValue.ServiceType := ServiceType.Value;
    FOldValue.PenalCode := Penal.Value;
    {}
    FilterCD078A1DataWhenCurrent;

    if FCD078A1DataSet.RecordCount <=0 then
      btnRankWhere.Enabled := True
    else
      btnRankWhere.Enabled := False;

  finally
    FLockCheck := False;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.DMLModeChange(aMode: TDML);
begin
  if not DataSetStateChange( aMode ) then Exit;
  if ( aMode = dmSave  ) and ( FMode = dmInsert ) then
  begin
    DMLModeChange( dmInsert );
    Exit;
  end else
  if ( aMode in [dmSave, dmCancel] ) then
    FMode := dmBrowser;
  ButtonStateChange;
  ResetEditorState;
  EditorStateChange;
  edtDML.Text := TDMLString[FMode];
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.ButtonStateChange;
begin
  if ( FMode = dmInsert ) then
  begin
    actSave.Enabled := True;
  end else
  if ( FMode = dmUpdate ) then
  begin
    actSave.Enabled := True;
  end else
  begin
    actSave.Enabled := False;
    actCancel.Caption := '????(&C)';
    actCancel.ShortCut := TextToShortCut( 'Alt+C' );
  end;
  {}
  btnChangeCItemCode.Visible := ( FMode = dmUpdate );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.EditorStateChange;
var
  aIndex: Integer;
begin
  for aIndex := 0 to ControlCount - 1 do
  begin
    if ( Controls[aIndex].Name <> 'btnSave' ) and
       ( Controls[aIndex].Name <> 'Button1' ) and
       ( Controls[aIndex].Name <> 'Panel1' ) and
       ( Controls[aIndex].Name <> 'btnDepositAmt' ) then
      Controls[aIndex].Enabled := not ( FMode in [dmBrowser] );
  end;
  btnUpdate.Enabled := ( FMode in [dmBrowser] );
  if ( FMode in [dmInsert, dmUpdate] ) then
  begin
    { ???q???u?f???? }
    for aIndex := 1 to 12 do
    begin
      FMon[aIndex].Enabled := False;
      { ???O???? }
      FPeriod[aIndex].Enabled := False;
      { ?O?v???? }
      FRateType[aIndex].Enabled := False;
      { ?u?f???B }
      FDiscountAmt[aIndex].Enabled := False;
      { ???????B }
      FMonthAmt[aIndex].Enabled := False;
      { ???????B }
      FDayAmt[aIndex].Enabled := False;
    end;
    {}
    Penal.Enabled := chkPenal.Checked;
    btnPenalCondition.Enabled := chkPenal.Checked;
    { ?O???????? }
    btnDepositAmt.Enabled := ( cmbDepositAttr.ItemIndex > 0 );
  end;
  case FMode of
    dmInsert:
      begin
        PFlag.Enabled := False;
        if ( ServiceType.CodeNo.CanFocusEx ) then ServiceType.CodeNo.SetFocus;
        CItem.Enabled := True;
        ServiceType.Enabled := True;
        gvPenal.Styles.Content := nil;
        cmbPFlag1.ItemIndex := 3;
      end;
    dmUpdate:
      begin
        PFlag.Enabled := ( cmbPFlag1.ItemIndex > 0 );
        edtPFlagDate.Enabled := ( cmbPFlag1.ItemIndex = 4 );
        if ( ServiceType.CodeNo.CanFocusEx ) then ServiceType.CodeNo.SetFocus;
        CItem.Enabled := False;
        ServiceType.Enabled := False;
        gvPenal.Styles.Content := nil;
      end;
    dmBrowser:
      begin
        if ( ServiceType.CodeNo.CanFocusEx ) then ServiceType.CodeNo.SetFocus;
        gvPenal.Styles.Content := StyleController.cxStyle1;;
      end;
  end;
  {}
  edtStepNo.Enabled := False;
  edtDescription.Enabled := False;
  chkMasterSale.Enabled := False;
  edtCMBaudNo.Enabled := False;
  edtCMBaudRate.Enabled := False;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B3.DataSetStateChange(const aMode: TDML): Boolean;

  { ------------------------------------------ }

  procedure DoSubDetailSave;
  begin
    if ( not FCD078A1DataSet.IsEmpty ) then
      UpdateCD078A1KeyValue;
    frmCD078B1.CloneCD078A1Data;
    {}
    if ( not FCD078A2DataSet.IsEmpty ) then
      UpdateCD078A2KeyValue;
    frmCD078B1.CloneCD078A2Data;
    {}
    if ( not FCD078A3DataSet.IsEmpty ) then
      UpdateCD078A3KeyValue;
    frmCD078B1.CloneCD078A3Data;
    {}
    if ( not FCD078A4DataSet.IsEmpty ) then
      UpdateCD078A4KeyValue;
    frmCD078B1.CloneCD078A4Data;
  end;

  { ------------------------------------------ }

begin
  Result := True;
  case aMode of
    dmInsert:
      begin
        FilterCD078A4DataWhenNew;
        FilterCD078A3DataWhenNew;
        FilterCD078A2DataWhenNew;
        FilterCD078A1DataWhenNew;
        ClearEditor;
        FDataSet.Append;
      end;
    dmUpdate:
      begin
        ClearEditor;
        DataToEditor;
        FDataSet.Edit;
      end;
    dmCancel:
      begin
        FDataSet.Cancel;
      end;
    dmBrowser:
      begin
        ClearEditor;
        DataToEditor;
      end;  
    dmSave:
      begin
        Result := DataValidate;
        if Result then Result := EditorToData;
        if Result then
        begin
          try
            FDataSet.Post;
            {#6865 ?????????O?n?P?B???t?? By Kin 2014/10/07}
            SyncLongPayFlag;
            DoSubDetailSave;
            UpdateMCItemDataSet( EmptyStr );
            frmCD078B1.RestoreLinkKeyData;
            FAlreadySetCItemValue := frmCD078B1.GetCD078ACItemValue;
          except
            on E: Exception do
            begin
              ErrorMsg( Format( '?s?????~, ???]:%s?C', [E.Message] ) );
              Result := False;
            end;  
          end;
        end;
      end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B3.DataValidate: Boolean;
begin

  Result := VdMustInput;
  if not Result then Exit;
  {}
  Result := VdBundle;
  if not Result then Exit;
  {}
  Result := VdDeposit;
  if not Result then Exit;
  {}
  Result := VdPFlagDate;
  if not Result then Exit;
  {}
  Result := VdPenal;
  if not Result then Exit;
  {}
  Result := VdCD078A1;
  if not Result then
  begin
    btnRankList.Click;
    Exit;
  end;
  {}
  Result := VdConfirm;
  if not Result then Exit;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B3.EditorToData: Boolean;
var
  aIndex: Integer;
  aDepositCode, aDepositName: String;
begin
  Result := False;
  try
    if ( FMode in [dmInsert] ) then
    begin
      FDataSet.FieldByName( 'BPCODE' ).AsString := FKeyValue;
    end;  
    FDataSet.FieldByName( 'CITEMCODE' ).AsString := CItem.Value;
    FDataSet.FieldByName( 'CITEMNAME' ).AsString := CItem.ValueName;
    FDataSet.FieldByName( 'MCITEMCODE' ).AsString := MCItem.Value;
    FDataSet.FieldByName( 'SERVICETYPE' ).AsString := ServiceType.Value;
    FDataSet.FieldByName( 'FACIITEM' ).AsString := FaciItem.Value;
    FDataSet.FieldByName( 'CMBAUDNO' ).AsString := edtCMBaudNo.Text;
    FDataSet.FieldByName( 'CMBAUDRATE' ).AsString := edtCMBaudRate.Text;
    {}
    FDataSet.FieldByName( 'BUNDLE' ).AsString := '0';
    if chkBundle.Checked then
      FDataSet.FieldByName( 'BUNDLE' ).AsString := '1';
    {}
    {#6700 ?W?[StopFlag By Kin 2014/1/3}
    FDataSet.FieldByName( 'STOPFLAG' ).AsString := '0';
    if chkStopFlag.Checked then
      FDataSet.FieldByName( 'STOPFLAG' ).AsString := '1';
    FDataSet.FieldByName( 'BUNDLEMON' ).AsString := edtBundleMon.Text;
    {}
    FDataSet.FieldByName( 'PENALTYPE' ).AsString := EmptyStr;
    if ( cmbPenalType.ItemIndex > 0 ) then
      FDataSet.FieldByName( 'PENALTYPE' ).AsInteger := cmbPenalType.ItemIndex;
    {#6865 ?W?[4.???????B By Kin 2014/09/09}
    if FDataSet.FieldByName( 'PENALTYPE' ).AsInteger = 3 then
      FDataSet.FieldByName( 'PENALTYPE' ).AsInteger := 4;
    {}
    FDataSet.FieldByName( 'EXPIRETYPE' ).AsString := EmptyStr;
    if ( cmbExpireType.ItemIndex > 0 ) then
      FDataSet.FieldByName( 'EXPIRETYPE' ).AsString := ValueOfItemIndex(
        cmbExpireType.ItemIndex, cmbExpireType.Properties.Items );
    {}
    FDataSet.FieldByName( 'EXPIRECUT' ).AsString := '0';
    {}
    {#5275 ?W?[?S?w?u?f?I???? By Kin 2010/05/05}

    if edtBpStopDate.Text  <> EmptyStr then
    begin
      FDataSet.FieldByName( 'BPStopDate' ).AsString := edtBpStopDate.EditText;
    end else
    begin
      FDataSet.FieldByName( 'BPStopDate' ).AsString :=EmptyStr;
    end;
    {}
    if ( cmbDepositAttr.ItemIndex > 0 ) then
    begin
      FDataSet.FieldByName( 'DEPOSITATTR' ).AsInteger := cmbDepositAttr.ItemIndex;
      { ?N CD078A2, ?w?]???I?????????X, ???? CD078A }
      AcquireDefaulDeposit( aDepositCode, aDepositName );
      FDataSet.FieldByName( 'DEPOSITCODE' ).AsString := aDepositCode;
      FDataSet.FieldByName( 'DEPOSITNAME' ).AsString := aDepositName;
    end else
    begin
      FDataSet.FieldByName( 'DEPOSITATTR' ).AsString := EmptyStr;
      FDataSet.FieldByName( 'DEPOSITCODE' ).AsString := EmptyStr;
      FDataSet.FieldByName( 'DEPOSITNAME' ).AsString := EmptyStr;

      EmptyCD078A2;
    end;
    {}
    FDataSet.FieldByName( 'MONTHAMT' ).AsInteger := Trunc( edtMonthAmt.Value );
    FDataSet.FieldByName( 'DAYAMT' ).AsFloat := edtDayAmt.Value;
    {}
    FDataSet.FieldByName( 'PFLAG1' ).AsString := EmptyStr;
    if ( cmbPFlag1.ItemIndex > 0 ) then
      FDataSet.FieldByName( 'PFLAG1' ).AsInteger := ( cmbPFlag1.ItemIndex - 1 );
    {}
    //#5913 ?W?[PFlag2 By Kin 2011/04/14
    //#6220 ?]?????FIndex1,???HIndex1?n?s??2 By Kin 2012/03/02
    FDataSet.FieldByName( 'PFLAG2' ).AsString := EmptyStr;
    if cmbPFlag2.ItemIndex > -1 then
    begin
      if cmbPFlag2.ItemIndex = 1 then
      begin
        FDataSet.FieldByName( 'PFlag2' ).AsInteger := cmbPFlag2.ItemIndex+1;
      end else
      begin
        FDataSet.FieldByName( 'PFlag2' ).AsInteger := cmbPFlag2.ItemIndex;
      end;

    end;

    {}
    FDataSet.FieldByName( 'DISCOUNTCALC' ).AsString := EmptyStr;
    {}
    FDataSet.FieldByName( 'PFLAGDATE' ).AsInteger := StrToIntDef( edtPFlagDate.Text, 0 );
    FDataSet.FieldByName( 'TRUNCAMT' ).AsInteger := Trunc( edtTruncAmt.Value );
    {}
    FDataSet.FieldByName( 'PFLAGCODE' ).AsString := PFlag.Value;
    FDataSet.FieldByName( 'PFLAGNAME' ).AsString := PFlag.ValueName;
    {}
    FDataSet.FieldByName( 'INSTCODESTR' ).AsString := TrimChar( FQueryParam.Param1, [''''] );
    {}
    { #4435 ?N?s?W??2?????????????J }
    if ( FCD019Data.Sign = '-' ) then
    begin
      FDataSet.FieldByName( 'MinCount' ).AsInteger := StrToIntDef(edtMinCount.Text,0);
      FDataSet.FieldByName( 'MaxCount' ).AsInteger := StrToIntDef(edtMaxCount.Text,0);
    end;
    {#5318 ?W?[?I?????P???? By Kin 2009/10/06}
    FDataSet.FieldByName( 'SalePointcode' ).AsString := MSalePoint.Value;
    FDataSet.FieldByName( 'SalePointName' ).AsString := MSalePoint.ValueName;
    {#6841 ?W?[?????O???O By Kin 2014/08/05}
    if chkLONGPAYflag.Checked then
      FDataSet.FieldByName( 'LONGPAYFLAG' ).AsInteger := 1
    else
      FDataSet.FieldByName( 'LONGPAYFLAG' ).AsInteger := 0;
    {?W?[ContractType By Kin 2015/04/21}
    FDataSet.FieldByName( 'ContractType' ).AsString := EmptyStr;
    if cmbContractType.Text <> EmptyStr then
      FDataSet.FieldByName( 'ContractType' ).AsInteger := cmbContractType.ItemIndex ;
    {# ?W?[AuthTunerCount By Kin 2015/04/21}
    if edtAuthTunerCount.Text <> EmptyStr then
      FDataSet.FieldByName( 'AuthTunerCount' ).AsString := edtAuthTunerCount.Text;
    {}
    { ???q???u?f???? }
    for aIndex := 1 to 12 do
    begin
      FDataSet.FieldByName( Format( 'MON%d', [aIndex] ) ).AsString := EmptyStr;
      FDataSet.FieldByName( Format( 'PERIOD%d', [aIndex] ) ).AsString := EmptyStr;
      FDataSet.FieldByName( Format( 'RATETYPE%d', [aIndex] ) ).AsString := EmptyStr;
      FDataSet.FieldByName( Format( 'DISCOUNTAMT%d', [aIndex] ) ).AsString := EmptyStr;
      FDataSet.FieldByName( Format( 'MONTHAMT%d', [aIndex] ) ).AsString := EmptyStr;
      FDataSet.FieldByName( Format( 'DAYAMT%d', [aIndex] ) ).AsString := EmptyStr;
      if ( FMon[aIndex].Text <> EmptyStr ) then
      begin
        FDataSet.FieldByName( Format( 'MON%d', [aIndex] ) ).AsString := FMon[aIndex].Text;
        FDataSet.FieldByName( Format( 'PERIOD%d', [aIndex] ) ).AsString := FPeriod[aIndex].Text;
        if ( FRateType[aIndex].ItemIndex > 0 ) then
          FDataSet.FieldByName( Format( 'RATETYPE%d', [aIndex] ) ).AsInteger := FRateType[aIndex].ItemIndex;
        if ( FDiscountAmt[aIndex].Text <> EmptyStr ) then
          FDataSet.FieldByName( Format( 'DISCOUNTAMT%d', [aIndex] ) ).AsInteger := Trunc( FDiscountAmt[aIndex].Value );
        if ( FMonthAmt[aIndex].Text <> EmptyStr ) then
          FDataSet.FieldByName( Format( 'MONTHAMT%d', [aIndex] ) ).AsFloat := FMonthAmt[aIndex].Value;
        if ( FDayAmt[aIndex].Text <> EmptyStr ) then
          FDataSet.FieldByName( Format( 'DAYAMT%d', [aIndex] ) ).AsFloat := FDayAmt[aIndex].Value;
      end;
    end;
    {}
    { ???q???H??????, ???M?? }
    FDataSet.FieldByName( 'PENAL' ).AsInteger := 0;
    FDataSet.FieldByName( 'PENALCODE' ).AsString := EmptyStr;
    FDataSet.FieldByName( 'PENALNAME' ).AsString := EmptyStr;
    { ?]?? }
    if ( chkPenal.Checked ) then
    begin
      FDataSet.FieldByName( 'PENAL' ).AsInteger := 1;
      FDataSet.FieldByName( 'PENALCODE' ).AsString := Penal.Value;
      FDataSet.FieldByName( 'PENALNAME' ).AsString := Penal.ValueName;
    end else
    begin
      EmptyCD078A3;
      EmptyCD078A4;
    end;
    { ?u?f???? }
    FCD078A1DataSet.First;
    while not CD078A1DataSet.Eof do
    begin
      if ( CD078A1DataSet.FieldByName( 'MasterSale' ).AsInteger = 1 ) then
      begin
        for aIndex := 1 to 12 do
        begin
          FDataSet.FieldByName( Format( 'Comment%d', [aIndex] ) ).AsString :=
            CD078A1DataSet.FieldByName( Format( 'Comment%d', [aIndex] ) ).AsString;
        end;
        Break;
      end;
      CD078A1DataSet.Next;
    end;
    { ??/?t?? }
    FDataSet.FieldByName( 'SIGN' ).AsString := FCD019Data.Sign;
    { ?????? }
    FDataSet.FieldByName( 'REFNO' ).AsString := FCD019Data.RefNo;


  except
    on E: Exception do
    begin
      ErrorMsg( Format( '?s?????~, ???]:%s', [E.Message] ) );
      Exit;
    end;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.CItemCodeNamePropertiesInitPopup(Sender: TObject);
var
  aCItems: String;
begin
  Screen.Cursor := crSQLWait;
  try
    CItem.SQL.Text := Format(
      ' SELECT A.CODENO, A.DESCRIPTION      ' +
      '  FROM %s.CD019 A                    ' +
      ' WHERE A.PERIODFLAG = ''1''          ' +
      '   AND A.STOPFLAG = ''0''            ',
      [DBController.LoginInfo.DbAccount] );
    {}
    if ( ServiceType.value <> EmptyStr )then
      CItem.SQL.Text := CItem.SQL.Text +
        ' AND A.SERVICETYPE=''' + ServiceType.Value+''' ';
    {}
    if ( FMode in [dmInsert] ) and ( FAlreadySetCItemValue <> EmptyStr ) then
    begin
      aCItems := QuotedValue( FAlreadySetCItemValue );
      CItem.SQL.Text := CItem.SQL.Text +
        ' AND A.CODENO NOT IN ( ' + aCItems + ' ) ';
    end;
    {}
    { #4407 ?p?G?O???~?H???h?W?[?P?_ MASTERPRODUCT <> 1  By Kin 2009/03/06 }
    { ???M???????D By Kin 2009/03/06 }
    { #5692 ??????23?]?n?????X?? By Kin 2010/06/17 }
    if ( frmCD078B1.KindFunction = '1' ) then
    begin
      CItem.SQL.Text := CItem.SQL.Text + ' AND Nvl(A.REFNO,0) IN( 2,23 ) ';
    end;
    {}
    CItem.SQL.Text := CItem.SQL.Text +
      '  ORDER BY A.CODENO  ';
    CItem.CodeNamePropertiesInitPopup( Sender );
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.ServiceTypeCodeNamePropertiesInitPopup(
  Sender: TObject);
begin
  { ?A?????O }
  Screen.Cursor := crSQLWait;
  try
    ServiceType.SQL.Text := Format(
    ' SELECT CODENO, DESCRIPTION    ' +
    '  FROM %s.CD046                ' ,
    [DBController.LoginInfo.DbAccount] );
    ServiceType.CodeNamePropertiesInitPopup( Sender );
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.LoadCitemRelationItem;
begin
  Screen.Cursor := crSQLWait;
  try
    FLockCItemChange := True;
    try
      FReader.Close;
      FReader.SQL.Text := Format(
        ' SELECT A.SERVICETYPE FROM %s.CD019 A ' +
        '  WHERE A.CODENO = ''%s''             ',
        [DBController.LoginInfo.DbAccount, CItem.Value] );
      FReader.Open;
      if ( FReader.FieldByName( 'SERVICETYPE' ).AsString <> EmptyStr ) and
         ( FReader.FieldByName( 'SERVICETYPE' ).AsString <> ServiceType.Value ) then
          ServiceType.Value := FReader.FieldByName( 'SERVICETYPE' ).AsString;
      FReader.Close;
      if not FaciDataSet.IsEmpty then
        FaciItemCodeNoPropertiesChange( FaciItem.CodeNo );
      FaciItem.Value := FaciDataSet.FieldByName( 'CODENO' ).AsString;
    finally
      FLockCItemChange := False;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.InitCD019Data;
begin
  FCD019Data.Sign := '+';
  FCD019Data.Amount := 0;
  FCD019Data.Period := 0;
  FCD019Data.MonthAmt := 0;
  FCD019Data.DayAmt := 0;
  FCD019Data.IsMinus := ( FCD019Data.Sign = '-' );
  FCD019Data.IsByHouse := False;
  FCD019Data.CMCode := EmptyStr;
  FCD019Data.CMName := EmptyStr;
  FCD019Data.RefNo := EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.LoadCD019Data;
begin
  InitCD019Data;
  FReader.Close;
  FReader.SQL.Text := Format(
    ' SELECT A.AMOUNT, A.PERIOD, A.SIGN,          ' +
    '        A.CMCODE, A.CMNAME, A.REFNO,         ' +
    '        A.BYHOUSE                            ' +
    '   FROM %s.CD019 A                           ' +
    '  WHERE A.CODENO = ''%s''                    ' +
    '    AND A.PERIODFLAG = ''1''                 ' +
    '    AND A.STOPFLAG = ''0''                   ',
    [DBController.LoginInfo.DbAccount, CItem.Value] );
  FReader.Open;
  FCD019Data.Sign := FReader.FieldByName( 'SIGN' ).AsString;
  FCD019Data.Amount := FReader.FieldByName( 'AMOUNT' ).AsFloat;
  FCD019Data.Period := FReader.FieldByName( 'PERIOD' ).AsInteger;
  FCD019Data.IsMinus := ( FCD019Data.Sign = '-' );
  FCD019Data.IsByHouse := ( FReader.FieldByName( 'BYHOUSE' ).AsInteger = 1 );
  FCD019Data.CMCode := FReader.FieldByName( 'CMCODE' ).AsString;
  FCD019Data.CMName := FReader.FieldByName( 'CMNAME' ).AsString;
  FCD019Data.RefNo := FReader.FieldByName( 'REFNO' ).AsString;
  FReader.Close;
  if ( FCD019Data.Period > 0 ) then
  begin
    FCD019Data.MonthAmt := CbRoundTo( ( FCD019Data.Amount / FCD019Data.Period ), 3 );
    FCD019Data.DayAmt := CbRoundTo( ( ( FCD019Data.Amount / FCD019Data.Period ) / 30 ), 3 );
    if ( FCD019Data.IsMinus ) then
    begin
      FCD019Data.Amount := ( 0 - FCD019Data.Amount );
      FCD019Data.MonthAmt := ( 0 - FCD019Data.MonthAmt );
      FCD019Data.DayAmt := ( 0 - FCD019Data.DayAmt );
    end;
  end;
  { #4435 ?t???~?i???? By Kin 2009/03/24 }
  lblMinCount.Visible := ( FCD019Data.Sign = '-' );
  lblMaxCount.Visible := ( FCD019Data.Sign = '-' );
  edtMinCount.Visible := ( FCD019Data.Sign = '-' );
  edtMaxCount.Visible := ( FCD019Data.Sign = '-' );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.CD019DataToFieldValue;
begin
  edtDayAmt.Value := FCD019Data.DayAmt;
  edtMonthAmt.Value := FCD019Data.MonthAmt;
  edtCMBaudNo.Text := FCD019Data.CMCode;
  edtCMBaudRate.Text := FCD019Data.CMName;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.ChangeRelationItemState;
begin
  Label18.Enabled := SameText( FCD019Data.Sign, '-' );
  MCItem.Enabled := SameText( FCD019Data.Sign, '-' );
  if ( not MCItem.Enabled ) then
    MCItem.Clear;
  Label9.Font.Color := clWindowText;
  if VdCMBaudNo then Label9.Font.Color := clBlue;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.ChangeFaciItemState;
begin
  Label3.Font.Color := clWindowText;
  if ( not FCD019Data.IsByHouse ) then
    Label3.Font.Color := clRed;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.ChangeExpireTypeItem;
var
  aIndex: Integer;
begin
  if SameText( FCD019Data.Sign, '-' ) then
  begin
    aIndex := ItemIndexOfValue( '5', cmbExpireType.Properties.Items );
    if ( aIndex >= 0 ) then
      cmbExpireType.Properties.Items.Delete( aIndex );
  end else
  begin
    if ( ItemIndexOfValue( '5', cmbExpireType.Properties.Items ) < 0 ) then
      cmbExpireType.Properties.Items.Add( '5.??????????' );
  end;
  {#5257 ??5?????????????t?????n?X?? By Kin 2010/05/05}
  if ( ItemIndexOfValue( '5', cmbExpireType.Properties.Items ) < 0 ) then
      cmbExpireType.Properties.Items.Add( '5.??????????' );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.CItemCodeNoExit(Sender: TObject);
begin
  CItem.CodeNoExit( Sender );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.CItemCodeNoPropertiesChange(Sender: TObject);
begin
  if ( not FLockCheck ) then
  begin
    LoadCitemRelationItem;
    InitCD019Data;
    if ( CItem.Value <> EmptyStr ) then
    begin
      LoadCD019Data;
      ChangeRelationItemState;
      ChangeFaciItemState;
      ChangeExpireTypeItem;
    end;
    CD019DataToFieldValue;
    if ( CItem.Value <> EmptyStr ) then
    begin
      UpdateCD078A1KeyValue;
      UpdateCD078A2KeyValue;
      UpdateCD078A3KeyValue;
      UpdateCD078A4KeyValue;
      FilterCD078A1DataWhenCurrent;
      FilterCD078A2DataWhenCurrent;
      FilterCD078A3DataWhenCurrent;
      FilterCD078A4DataWhenCurrent;
      //AddCD078A1DefaultRecord;
    end;
  end;

  btnRankWhere.Enabled := not ( FCD078A1DataSet.RecordCount > 0 );

  CItem.CodeNoPropertiesChange( Sender );

end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.CItemCodeNamePropertiesChange(Sender: TObject);
begin
  CItem.CodeNamePropertiesChange( Sender );
  if ( not FLockCheck ) then
  begin
    LoadCitemRelationItem;
    InitCD019Data;
    if ( CItem.Value <> EmptyStr ) then
    begin
      LoadCD019Data;
      ChangeRelationItemState;
      ChangeFaciItemState;
      ChangeExpireTypeItem;
    end;
    CD019DataToFieldValue;
    if ( CItem.Value <> EmptyStr ) then
    begin
      UpdateCD078A1KeyValue;
      UpdateCD078A2KeyValue;
      UpdateCD078A3KeyValue;
      UpdateCD078A4KeyValue;
      FilterCD078A1DataWhenCurrent;
      FilterCD078A2DataWhenCurrent;
      FilterCD078A3DataWhenCurrent;
      FilterCD078A4DataWhenCurrent;
      //AddCD078A1DefaultRecord;
    end;
  end;
  btnRankWhere.Enabled := not ( FCD078A1DataSet.RecordCount > 0 );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.FaciItemCodeNamePropertiesChange(Sender: TObject);
begin
  FaciItem.CodeNamePropertiesChange( Sender );
  if not FLockCheck then
    ChangeRelationItemState;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.FaciItemCodeNamePropertiesInitPopup(Sender: TObject);
begin
  FaciDataSet.Filtered := False;
  FaciDataSet.Filter := Format( 'SERVICETYPE=''%s''', [ServiceType.Value] );
  FaciDataSet.Filtered := True;
  FaciItem.CodeNamePropertiesInitPopup( Sender );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.FaciItemCodeNoPropertiesChange(Sender: TObject);
begin
  FaciDataSet.Filtered := False;
  FaciDataSet.Filter := Format( 'SERVICETYPE=''%s''', [ServiceType.Value] );
  FaciDataSet.Filtered := True;
  FaciItem.CodeNoPropertiesChange( Sender );
  if not FLockCheck then
    ChangeRelationItemState;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.ServiceTypeCodeNoPropertiesChange(Sender: TObject);
begin
  ServiceType.CodeNoPropertiesChange( Sender );
  if ( not FLockCItemChange ) then
  begin
    CItem.Clear;
    FaciItem.Clear;
    cmbDepositAttr.ItemIndex := 0;
    Penal.Clear;
    PFlag.Clear;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.ServiceTypeCodeNamePropertiesChange(Sender: TObject);
begin
  ServiceType.CodeNamePropertiesChange( Sender );
  CItem.Clear;
  FaciItem.Clear;
  PFlag.Clear;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.ResetEditorState;
var
  aIndex: Integer;
begin
  { ???q???u?f???? }
  for aIndex := 1 to 12 do
  begin
    { ???????? }
    FMon[aIndex].Enabled := False;
    { ???O???? }
    FPeriod[aIndex].Enabled := False;
    { ?O?v???? }
    FRateType[aIndex].Enabled := False;
    { ?u?f???B }
    FDiscountAmt[aIndex].Enabled := False;
    { ???????B }
    FMonthAmt[aIndex].Enabled := False;
    { ???????B }
    FDayAmt[aIndex].Enabled := False;
  end;
  {}
  Penal.Enabled := False;
end;

{ ---------------------------------------------------------------------------- }

var
  FSValue: Integer;

procedure TfrmCD078B3.InternalCalcMonth(aIndex: Integer);
var
  aSValue, aEvalue: Integer;
  aLabel: TLabel;
  aControl: TcxMaskEdit;
begin
  if ( aIndex = 1 ) then FSValue := 0;
  aLabel := FMonLabel[aIndex];
  aControl := FMon[aIndex];
  aLabel.Caption := EmptyStr;
  if ( aControl.Text <> EmptyStr ) then
  begin
    if ( StrToInt( aControl.Text ) <= 0 ) then
    begin
     aSValue := FSValue;
     aEValue := aSValue;
    end else
    begin
      aSValue := FSValue + 1;
      aEValue := aSValue + StrToInt( aControl.Text ) - 1;
    end;
    FSValue := aEValue;
    aLabel.Caption := Format( '%d~%d', [aSValue, aEValue] );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.AcquireDefaulDeposit(var ACode, AName: String);
begin
  FCD078A2DataSet.First;
  while not FCD078A2DataSet.Eof do
  begin
    if ( FCD078A2DataSet.FieldByName( 'DEPOSITDEFAULT' ).AsInteger = 1 ) then
    begin
      ACode := FCD078A2DataSet.FieldByName( 'DEPOSITCODE' ).AsString;
      AName := FCD078A2DataSet.FieldByName( 'DEPOSITNAME' ).AsString;
      Break;
    end;
    FCD078A2DataSet.Next;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B3.VdMonthBreakAndInput: Boolean;
var
  aIndex, aLastPos: Integer;
begin
  Result := False;
  if ( FMon[1].Text = EmptyStr ) then
  begin
    WarningMsg( '???q???u?f?????i???@???????????j???????J?C' );
    if ( FMon[1].CanFocusEx ) then FMon[1].SetFocus;
    Exit;
  end;
  aLastPos := 0;
  for aIndex := 1 to 12 do
  begin
    if ( FMon[aIndex].Text <> EmptyStr ) then
    begin
      if ( aIndex - 1 ) = aLastPos then
      begin
        aLastPos := aIndex;
        {}
        if ( FPeriod[aIndex].Text = EmptyStr ) then
        begin
          WarningMsg( '?????J?i???O?????j?C' );
          if ( FPeriod[aIndex].CanFocusEx ) then FPeriod[aIndex].SetFocus;
          Exit;
        end;
        {}
        if ( FRateType[aIndex].ItemIndex <= 0 ) then
        begin
          WarningMsg( '?????J?i?O?v?????j?C' );
          if ( FRateType[aIndex].CanFocusEx ) then FRateType[aIndex].SetFocus;
          Exit;
        end;
      end else
      begin
        WarningMsg( '?h???u?f?????i?????????j?????s?????J, ???i???_?C' );
        if ( FMon[aLastPos+1].CanFocusEx ) then
          FMon[aLastPos+1].SetFocus;
        Exit;
      end;
    end
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.chkPenalPropertiesChange(Sender: TObject);
begin
  if chkPenal.Checked then
  begin
    Penal.Enabled := True;
    btnPenalCondition.Enabled := True;
  end else
  begin
    Penal.Clear;
    Penal.Enabled := False;
    btnPenalCondition.Enabled := False;
    EmptyCD078A3;
    EmptyCD078A4;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.MCItemCodeNamePropertiesChange(Sender: TObject);
begin
  MCItem.CodeNamePropertiesChange( Sender );
  if ( not FLockCheck ) then
    UpdateCD078A1LinkKey;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.MCItemCodeNamePropertiesInitPopup(Sender: TObject);
begin
  Screen.Cursor := crSQLWait;
  try
    FMCItemDataSet.Filtered := False;
    FMCItemDataSet.Filter := Format( 'SERVICETYPE=''%s''', [ServiceType.Value] );
    FMCItemDataSet.Filtered := True;
    MCItem.CodeNamePropertiesInitPopup( Sender );
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.MCItemCodeNoPropertiesChange(Sender: TObject);
begin
  Screen.Cursor := crSQLWait;
  try
    FMCItemDataSet.Filtered := False;
    FMCItemDataSet.Filter := Format( 'SERVICETYPE=''%s''', [ServiceType.Value] );
    FMCItemDataSet.Filtered := True;
    MCItem.CodeNoPropertiesChange( Sender );
    if ( not FLockCheck ) then
      UpdateCD078A1LinkKey;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.MonPropertiesChange(Sender: TObject);
var
  aIndex: Integer;
begin
  for aIndex := 1 to 12  do
    InternalCalcMonth( aIndex );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.Page1DrawTabEx(AControl: TcxCustomTabControl;
  ATab: TcxTab; Font: TFont);
begin

end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.PenalCodeNamePropertiesInitPopup(Sender: TObject);
begin
  Penal.SQL.Text := Format(
    '  SELECT A.CODENO, A.DESCRIPTION         ' +
    '    FROM %s.CD019 A                      ' +
    '   WHERE A.PERIODFLAG = ''0''            ' +
    '     AND A.SIGN = ''+''                  ' +
    '     AND A.REFNO = ''14''                ' +
    '     AND A.SERVICETYPE = ''%s''          ' +
    '     AND A.STOPFLAG = ''0''              ' +
    '   ORDER BY A.CODENO                     ',
    [DBController.LoginInfo.DbAccount, ServiceType.Value] );
  Penal.CodeNamePropertiesInitPopup( Sender );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.cmbDepositAttrPropertiesChange(Sender: TObject);
begin
  btnDepositAmt.Enabled := ( cmbDepositAttr.ItemIndex > 0 );
  if not FLockCheck then
    EmptyCD078A2;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.cmbPFlag1PropertiesChange(Sender: TObject);
begin
  if cmbPFlag1.ItemIndex <= 0 then PFlag.Clear;
  PFlag.Enabled := ( cmbPFlag1.ItemIndex > 0 );
  edtPFlagDate.Enabled := ( cmbPFlag1.ItemIndex = 4 );
  if not edtPFlagDate.Enabled then
    edtPFlagDate.Text := '';
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.PFlagCodeNamePropertiesInitPopup(Sender: TObject);
begin
  SCreen.Cursor := crSQLWait;
  try
    PFlag.SQL.Text := Format(
      '  SELECT A.CODENO, A.DESCRIPTION         ' +
      '    FROM %s.CD019 A                      ' +
      '   WHERE A.SIGN = ''+''                  ' +
      '     AND A.SERVICETYPE = ''%s''          ' +
      '     AND A.STOPFLAG = ''0''              ' +
      '   ORDER BY A.CODENO                     ',
      [DBController.LoginInfo.DbAccount, ServiceType.Value] );
    PFlag.CodeNamePropertiesInitPopup( Sender );
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B3.VdMustInput: Boolean;
begin
  Result := False;
  if ( CItem.Value = EmptyStr ) then
  begin
    WarningMsg( '?????J?i???O?????j?C' );
    if CItem.CodeNo.CanFocusEx then CItem.CodeNo.SetFocus;
    Exit;
  end;
  {}
  if ( ServiceType.Value = EmptyStr ) then
  begin
    WarningMsg( '?????J?i?A?????O?j?C' );
    if ServiceType.CodeNo.CanFocusEx then ServiceType.CodeNo.SetFocus;
    Exit;
  end;
  {}
  if ( SameText( FCD019Data.Sign, '-' ) and ( MCItem.Value = EmptyStr ) ) then
  begin
    if ( not ConfirmMsg( '???i???O?????j???t??, ?O?_?????w?i???????????O?????j?' ) ) then
    begin
      if MCItem.CodeNo.CanFocusEx then MCItem.CodeNo.SetFocus;
      Exit;
    end;
  end;
  {}
  if ( not FCD019Data.IsByHouse ) and ( FaciItem.Value = EmptyStr ) then
  begin
    if ( not ConfirmMsg( '???i???O?????j?H????????, ?O?_?????w?i???w?]???j?' ) ) then
    begin
      if FaciItem.CodeNo.CanFocusEx then FaciItem.CodeNo.SetFocus;
      Exit;
    end;
  end;
  {}
  if ( chkPenal.Checked ) then
  begin
    if ( Penal.Value = EmptyStr ) then
    begin
      WarningMsg( '?w???w?p???H????, ?????J?i?H?????????j?C' );
      if Penal.CodeNo.CanFocusEx then Penal.CodeNo.SetFocus;
      Exit;
    end;
    if ( FCD078A3DataSet.IsEmpty ) then
    begin
      WarningMsg( '?w???w?p???H????, ?????]?w?i???q???H???????j?C' );
      btnPenalCondition.Click;
      Exit;
    end;
  end;
  {}
  if ( FQueryParam.Param1 = EmptyStr ) then
  begin
    WarningMsg( '?????J?i???w???u???O?j?C' );
    Page1.ActivePageIndex := 1;
    btnInstCodeStr.Click;
    Exit;
  end;
  {}
  if not DBController.VdInstCodeExitisInCD078B( FQueryParam.Param1,
    FInstCodeDataSet ) then
  begin
    WarningMsg( '?i???w???u???O?j???????b?i???z?????u?????j???????s?b, ?????s?]?w?C' );
    Page1.ActivePageIndex := 1;
    btnInstCodeStr.Click;
    Exit;
  end;
  {}
  if ( FCD019Data.Sign = '-' ) then
  begin
    if  (StrToInt( edtMaxCount.Text) <> 0 ) or (StrToInt(edtMinCount.Text) <> 0)    then
    begin
      if ( StrToInt(edtMinCount.Text) > StrToInt( edtMaxCount.Text ) ) then
      begin
        WarningMsg( '???h?x???????j???????x???I' );
        edtMinCount.SetFocus;
        Exit;
      end;
    end;
  end;
  {}
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B3.VdBundle: Boolean;
begin
  Result := False;
  { ?????????s?????J, ???i???_ }
  if not VdMonthBreakAndInput then Exit;
  if chkBundle.Checked then
  begin
    if Trim( edtBundleMon.Text ) = EmptyStr then
    begin
      WarningMsg( '?w???w?j??, ?????J?i?j???????j?C' );
      Page1.ActivePageIndex := 0;
      if edtBundleMon.CanFocusEx then edtBundleMon.SetFocus;
      Exit;
    end;
    if ( edtBundleMon.Text = '0' ) then
    begin
      WarningMsg( '?w???w?j??, ?i?j???????j?????j???s?C' );
      Page1.ActivePageIndex := 0;
      if edtBundleMon.CanFocusEx then edtBundleMon.SetFocus;
      Exit;
    end;
    if ( cmbPenalType.ItemIndex <= 0 ) then
    begin
      WarningMsg( '?w???w?j??, ?????J?i?H?????p???????j?C' );
      Page1.ActivePageIndex := 0;
      if cmbPenalType.CanFocusEx then cmbPenalType.SetFocus;
      Exit;
    end;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B3.VdDeposit: Boolean;
var
  aDefaulDepositCount: Integer;
begin
  Result := False;
  { ?O???? }
  if ( cmbDepositAttr.ItemIndex in [1..2]) then
  begin
    if ( FCD078A2DataSet.IsEmpty ) then
    begin
      WarningMsg( '?w???w?O????????, ???]?w?i?I???????O?????]?w?j?C' );
      Page1.ActivePageIndex := 0;
      btnDepositAmtClick( btnDepositAmt );
      Exit;
    end;
    {}
    aDefaulDepositCount := 0;
    FCD078A2DataSet.First;
    while not FCD078A2DataSet.Eof do
    begin
      if ( FCD078A2DataSet.FieldByName( 'DepositDefault' ).AsString = '1' ) then
        Inc( aDefaulDepositCount );
      FCD078A2DataSet.Next;
    end;
    if ( aDefaulDepositCount <= 0 ) then
    begin
      WarningMsg( '???????w?@???i?I???????O?????]?w?j?????w?]?I???O?????C' );
      Page1.ActivePageIndex := 0;
      btnDepositAmtClick( btnDepositAmt );
      Exit;
    end else
    if ( aDefaulDepositCount > 1 ) then
    begin
      WarningMsg( '?i?I???????O?????]?w?j?w?]???u?????@??, ?????s?]?w?C' );
      Page1.ActivePageIndex := 0;
      btnDepositAmtClick( btnDepositAmt );
      Exit;
    end;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B3.VdPFlagDate: Boolean;
begin
  Result := False;
  if ( cmbPFlag1.ItemIndex = 4 ) then
  begin
    if edtPFlagDate.Enabled and (Trim( edtPFlagDate.Text ) = EmptyStr) then
    begin
      WarningMsg( '?????J?i?}?????????j?C' );
      Page1.ActivePageIndex := 1;
      if edtPFlagDate.CanFocusEx then edtPFlagDate.SetFocus;
      Exit;
    end;
    if ( cmbPFlag1.ItemIndex = 4 ) and
       ( not ( StrToInt( edtPFlagDate.Text ) in [1..30] ) ) then
    begin
      WarningMsg( '???}??????(???u??)= 3.?}??????????, ?i?}?????????j??????????1~30?C' );
      Page1.ActivePageIndex := 1;
      if edtPFlagDate.CanFocusEx then edtPFlagDate.SetFocus;
      Exit;
    end;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B3.VdPenal: Boolean;
var
  aBMark: TBookmark;
begin
  Result := False;
  if ( chkPenal.Checked ) then
  begin
    FCD078A3DataSet.DisableControls;
    aBMark := FCD078A3DataSet.GetBookmark;
    try
      FCD078A3DataSet.First;
      while not FCD078A3DataSet.Eof do
      begin
        if ( FCD078A3DataSet.FieldByName( 'PENALMON' ).AsInteger <= 0 ) then
        begin
          WarningMsg( '?i???q???H???j?p?????i?????j???????j???~, ???i???s???t???C' );
          if PenalGrid.CanFocusEx then PenalGrid.SetFocus;
          Exit;
        end;
        if ( FCD078A3DataSet.FieldByName( 'PENALAMT' ).AsInteger < 0 ) then
        begin
          WarningMsg( '?i???q???H???j?p?????i?H?????B?j???~, ???i???t???C' );
          if PenalGrid.CanFocusEx then PenalGrid.SetFocus;
          Exit;
        end;
        if ( chkBundle.Checked ) and ( edtBundleMon.Text <> EmptyStr ) then
        begin
          if ( FCD078A3DataSet.FieldByName( 'PENALMON' ).AsInteger > StrToIntDef( edtBundleMon.Text, 0 ) ) then
          begin
            WarningMsg( '?i???q???H???j?p?????i?????j???????j???~, ?`?????w?W?X?i?j???????j?C' );
            if PenalGrid.CanFocusEx then PenalGrid.SetFocus;
            Exit;
          end;
        end;
        FCD078A3DataSet.Next;
      end;
      FCD078A3DataSet.GotoBookmark( aBMark );
    finally
      FCD078A3DataSet.FreeBookmark( aBMark );
      FCD078A3DataSet.EnableControls;
    end;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B3.VdCanShowMultiRank: Boolean;
begin
  Result := False;
  if ( ServiceType.Value = EmptyStr ) then
  begin
    WarningMsg( '?????]?w?i?A?????O?j?C' );
    if ( ServiceType.CodeNo.CanFocusEx ) then ServiceType.CodeNo.SetFocus;
      Exit;
  end;
  {}
  if ( CItem.Value = EmptyStr ) then
  begin
    WarningMsg( '?????]?w?i???O?????j?C' );
    if ( CItem.CodeNo.CanFocusEx ) then CItem.CodeNo.SetFocus;
    Exit;
  end;
  {}
  if ( SameText( FCD019Data.Sign, '-' ) ) then
  begin
    if ( MCItem.Value = EmptyStr ) then
    begin
      WarningMsg( '?n?????????v?p???u?f???B, ?????]?w?i???????????O?????j?C' );
    end;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B3.VdCanShowPenalRule: Boolean;
begin
  Result := False;
  if ( ServiceType.Value = EmptyStr ) then
  begin
    WarningMsg( '?????]?w?i?A?????O?j?C' );
    if ( ServiceType.CodeNo.CanFocusEx ) then ServiceType.CodeNo.SetFocus;
      Exit;
  end;
  {}
  if ( CItem.Value = EmptyStr ) then
  begin
    WarningMsg( '?????]?w?i???O?????j?C' );
    if ( CItem.CodeNo.CanFocusEx ) then CItem.CodeNo.SetFocus;
    Exit;
  end;
  {}
  if ( Penal.Value = EmptyStr ) then
  begin
    WarningMsg( '?????]?w?i?H?????????j?C' );
    if ( Penal.CodeNo.CanFocusEx ) then Penal.CodeNo.SetFocus;
    Exit;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B3.VdCanShowMultiDeposit: Boolean;
begin
  Result := False;
  if ( ServiceType.Value = EmptyStr ) then
  begin
    WarningMsg( '?????]?w?i?A?????O?j?C' );
    if ( ServiceType.CodeNo.CanFocusEx ) then ServiceType.CodeNo.SetFocus;
      Exit;
  end;
  {}
  if ( CItem.Value = EmptyStr ) then
  begin
    WarningMsg( '?????]?w?i???O?????j?C' );
    if ( CItem.CodeNo.CanFocusEx ) then CItem.CodeNo.SetFocus;
    Exit;
  end;
  {}
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B3.VdCD078A1: Boolean;
begin
  Result := True;
  if ( SameText( FCD019Data.Sign, '-' ) and ( MCItem.Value <> EmptyStr ) ) then
  begin
    FCD078A1DataSet.First;
    while not FCD078A1DataSet.Eof do
    begin
      if ( FCD078A1DataSet.FieldByName( 'LinkKey' ).AsString = EmptyStr ) then
      begin
        Result := False;
        WarningMsg( '???]?w?i???q???u?f?????j???i?????????h???????j?C' );
        Exit;
      end;
      FCD078A1DataSet.Next;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B3.VdCMBaudNo: Boolean;
begin
  Result :=
    ( StrToIntDef( FaciItem.RefNo, -1 ) in [2,5] ) and
    ( SameText( ServiceType.Value, 'I' ) );
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B3.VdConfirm: Boolean;
begin
  Result := True;
  {}
  if ( VdCMBaudNo ) and ( edtCMBaudNo.Text = EmptyStr ) then
  begin
    Result := ( ConfirmMsg( '???i???O?????j?????? CM ?t?v?L?]?w??,?O?_?s???' ) );
    if ( not Result ) then
      if CItem.CodeNo.CanFocusEx then CItem.CodeNo.SetFocus;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.actUpdateExecute(Sender: TObject);
begin
  FMode := dmUpdate;
  DMLModeChange(dmUpdate);
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.btnDepositAmtClick(Sender: TObject);
var
  aKeys: array [1..3] of String;
begin
  if not VdCanShowMultiDeposit then Exit;
  {}
  aKeys[1] := FKeyValue;
  aKeys[2] := CItem.Value;
  aKeys[3] := ServiceType.Value;
  {}
  if ( FMode = dmInsert ) then FilterCD078A2DataWhenCurrent;
  frmCD078B3_3 := TfrmCD078B3_3.Create( FMode, FCD078A2DataSet, aKeys );
  try
    frmCD078B3_3.Caption := '?I???????O?????]?w[frmCD078B31]';
    frmCD078B3_3.DepositAttr := cmbDepositAttr.ItemIndex;
    frmCD078B3_3.ShowModal;
  finally
    frmCD078B3_3.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.btnInstCodeStrClick(Sender: TObject);
begin
  if ( FMode in [dmInsert, dmUpdate] ) then
  begin
    frmMultiSelect := TfrmMultiSelect.Create( Application );
    try
      frmMultiSelect.Connection := DBController.DataConnection;
      frmMultiSelect.KeyFields := 'INSTCODE';
      frmMultiSelect.KeyValues := FQueryParam.Param1;
      frmMultiSelect.DisplayFields := 'INSTCODE,?????N?X,INSTNAME,???????O?W??';
      frmMultiSelect.ResultFields := 'INSTNAME';
      FInstCodeDataSet.Filtered := False;
      FInstCodeDataSet.Filter := Format( 'SERVICETYPE=''%s''', [ServiceType.Value] );
      FInstCodeDataSet.Filtered := True;
      try
        frmMultiSelect.DataSet := FInstCodeDataSet;
        if frmMultiSelect.ShowModal = mrOk then
        begin
          if ( FMode in [dmInsert, dmUpdate] ) then
          begin
            FQueryParam.Param1 := frmMultiSelect.SelectedValue;
            edtInstCodeStr.Text := frmMultiSelect.SelectedDisplay;
          end;
        end;
      finally
        FInstCodeDataSet.Filtered := False;
        FInstCodeDataSet.Filter := EmptyStr;
      end;
    finally
      frmMultiSelect.Free;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.cxButton1Click(Sender: TObject);
begin
  if ( FMode in [dmInsert, dmUpdate] ) then
  begin
    FQueryParam.Param1 := EmptyStr;
    edtInstCodeStr.Text := EmptyStr;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.btnRankListClick(Sender: TObject);
var
  aKeys: array [1..4] of String;
begin
  if not VdCanShowMultiRank then Exit;
  {}
  aKeys[1] := FKeyValue;
  aKeys[2] := CItem.Value;
  aKeys[3] := ServiceType.Value;
  aKeys[4] := Penal.Value;
  {}
  if ( FMode = dmInsert ) then FilterCD078A1DataWhenCurrent;
  if ( SameText( FCD019Data.Sign, '-' ) ) and ( MCItem.Value <> EmptyStr ) then
    FilterLinkKeyDataSet;
  {}
  frmCD078B3_1 := TfrmCD078B3_1.Create( FMode, FCD078A1DataSet, aKeys );
  try
    frmCD078B3_1.Caption := '?h?????u?f????[frmCD078B30]';
    frmCD078B3_1.RankAmount.Sign := FCD019Data.Sign;
    {****************************************************************}
    {#4329 ?Y???t??,?hCD078A1 ???u?f???B,?n???????????????P??}
    //frmCD078B3_1.RankAmount.DefaultPeriod := FCD019Data.Period;
    if FCD019Data.Sign = '-' then
    begin
      frmCD078B3_1.RankAmount.DefaultPeriod :=
        DBController.GetPeriod(MCItem.Value) ;
    end else
    begin
      frmCD078B3_1.RankAmount.DefaultPeriod := FCD019Data.Period;
    end;
    {*****************************************************************}
    { ?Y???t??, ?h CD078A1 ???u?f???B, ?n???????????????P?? * ?????v }
    if ( FCD019Data.Sign = '-' ) then
    begin
      { ?????????????P?? }
      frmCD078B3_1.RankAmount.DefaultDiscountAmount :=
        DBController.GetDiscountAmt( MCItem.Value ) ;
    end else
    begin
      frmCD078B3_1.RankAmount.DefaultDiscountAmount := FCD019Data.Amount;

    end;
    frmCD078B3_1.LinkKeyDataSet := FLinkKeyDataSet;
    frmCD078B3_1.EnableLinkKey :=
      ( SameText( FCD019Data.Sign, '-' ) ) and ( MCItem.Value <> EmptyStr );
    frmCD078B3_1.ShowModal;
  finally
    frmCD078B3_1.Free;
  end;
  UpdateMasterSale;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.UpdateMasterSale;
var
  aIndex: Integer;
begin
  edtStepNo.Clear;
  edtDescription.Clear;
  chkMasterSale.Clear;
  {}
  CD078A1DataSet.First;
  while not CD078A1DataSet.Eof do
  begin
    if ( CD078A1DataSet.FieldByName( 'MasterSale' ).AsInteger = 1 ) then
    begin
      edtStepNo.Text := CD078A1DataSet.FieldByName( 'StepNo' ).AsString;
      edtDescription.Text := CD078A1DataSet.FieldByName( 'Description' ).AsString;
      chkMasterSale.Checked := True;
      for aIndex := 1 to 12 do
      begin
        FMon[aIndex].Text := CD078A1DataSet.FieldByName( Format( 'Mon%d', [aIndex] ) ).AsString;
        FPeriod[aIndex].Text :=  CD078A1DataSet.FieldByName( Format( 'Period%d', [aIndex] ) ).AsString;
        FRateType[aIndex].ItemIndex := CD078A1DataSet.FieldByName( Format( 'RateType%d', [aIndex] ) ).AsInteger;
        FDiscountAmt[aIndex].Value := CD078A1DataSet.FieldByName( Format( 'DiscountAmt%d', [aIndex] ) ).AsInteger;
        FMonthAmt[aIndex].Value := CD078A1DataSet.FieldByName( Format( 'MonthAmt%d', [aIndex] ) ).AsFloat;
        FDayAmt[aIndex].Value := CD078A1DataSet.FieldByName( Format( 'Dayamt%d', [aIndex] ) ).AsFloat;
      end;
      Break;
    end;
    CD078A1DataSet.Next;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.UpdateMCItemDataSet(const ADeleteCItem: String);
begin
  if ( ADeleteCItem <> EmptyStr ) then
  begin
    if ( MCItemDataSet.Locate( 'CODENO', ADeleteCItem, [] ) ) then
      MCItemDataSet.Delete;
  end else
  begin
    if ( FOldValue.CItemCode <> EmptyStr ) and
       ( FOldValue.CItemCode <> CItem.Value ) then
    begin
      if ( MCItemDataSet.Locate( 'CODENO', FOldValue.CItemCode, [] ) ) then
        MCItemDataSet.Delete;
    end;
    if SameText( FCD019Data.Sign, '+' ) then
    begin
      if ( not MCItemDataSet.Locate( 'CODENO', CItem.Value, [] ) ) then
        MCItemDataSet.AppendRecord( [CItem.Value, CItem.ValueName, ServiceType.Value] );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.UpdateCD078A1KeyValue;
var
  aRecordCount: Integer;
begin
  aRecordCount := FCD078A1DataSet.RecordCount;
  FCD078A1DataSet.First;
  while not FCD078A1DataSet.Eof do
  begin
    FCD078A1DataSet.Edit;
    FCD078A1DataSet.FieldByName( 'BPCode' ).AsString := FKeyValue;
    FCD078A1DataSet.FieldByName( 'CitemCode' ).AsString := CItem.Value;
    FCD078A1DataSet.FieldByName( 'ServiceType' ).AsString := ServiceType.Value;
    FCD078A1DataSet.FieldByName( 'FaciItem' ).AsString := FaciItem.Value;
    { ???s LinkKey }
    if ( SameText( FCD019Data.Sign, '-' ) ) and ( MCItem.Value = EmptyStr ) then
    begin
      FCD078A1DataSet.FieldByName( 'LinkKey' ).AsString := EmptyStr;
      FCD078A1DataSet.FieldByName( 'LinkKeyName' ).AsString := EmptyStr;
    end;
    FCD078A1DataSet.Post;
    { ?]???]?w Filter ??????, ?Y Key ????????, ?h Post ?????|?Q Filter ?o?? }
    { ?o?????Y?A?I?s Next ???k, ?????|???L???????? }
    if ( aRecordCount = FCD078A1DataSet.RecordCount ) then
      FCD078A1DataSet.Next;
  end;
  FCD078A1DataSet.First;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.UpdateCD078A2KeyValue;
var
  aRecordCount: Integer;
begin
  aRecordCount := FCD078A2DataSet.RecordCount;
  FCD078A2DataSet.First;
  while not FCD078A2DataSet.Eof do
  begin
    FCD078A2DataSet.Edit;
    FCD078A2DataSet.FieldByName( 'BPCODE' ).AsString := FKeyValue;
    FCD078A2DataSet.FieldByName( 'CITEMCODE' ).AsString := CItem.Value;
    FCD078A2DataSet.FieldByName( 'SERVICETYPE' ).AsString := ServiceType.Value;
    FCD078A2DataSet.Post;
    { ?]???]?w Filter ??????, ?Y Key ????????, ?h Post ?????|?Q Filter ?o?? }
    { ?o?????Y?A?I?s Next ???k, ?????|???L???????? }
    if ( aRecordCount = FCD078A2DataSet.RecordCount ) then
      FCD078A2DataSet.Next;
  end;
  FCD078A2DataSet.First;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.UpdateCD078A3KeyValue;
var
  aRecordCount: Integer;
begin
  FCD078A3DataSet.DisableControls;
  try
    aRecordCount := FCD078A3DataSet.RecordCount;
    FCD078A3DataSet.First;
    while not FCD078A3DataSet.Eof do
    begin
      FCD078A3DataSet.Edit;
      FCD078A3DataSet.FieldByName( 'BPCODE' ).AsString := FKeyValue;
      FCD078A3DataSet.FieldByName( 'CITEMCODE' ).AsString := CItem.Value;
      FCD078A3DataSet.FieldByName( 'SERVICETYPE' ).AsString := ServiceType.Value;
      FCD078A3DataSet.FieldByName( 'PENALCODE' ).AsString := Penal.Value;
      FCD078A3DataSet.Post;
      { ?]???]?w Filter ??????, ?Y Key ????????, ?h Post ?????|?Q Filter ?o?? }
      { ?o?????Y?A?I?s Next ???k, ?????|???L???????? }
      if ( aRecordCount = FCD078A3DataSet.RecordCount ) then
        FCD078A3DataSet.Next;
    end;
    FCD078A3DataSet.First;
  finally
    FCD078A3DataSet.EnableControls;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.UpdateCD078A4KeyValue;
var
  aRecordCount: Integer;
begin
  FCD078A4DataSet.DisableControls;
  try
    aRecordCount := FCD078A4DataSet.RecordCount;
    FCD078A4DataSet.First;
    while not FCD078A4DataSet.Eof do
    begin
      FCD078A4DataSet.Edit;
      FCD078A4DataSet.FieldByName( 'BPCODE' ).AsString := FKeyValue;
      FCD078A4DataSet.FieldByName( 'CITEMCODE' ).AsString := CItem.Value;
      FCD078A4DataSet.FieldByName( 'SERVICETYPE' ).AsString := ServiceType.Value;
      FCD078A4DataSet.FieldByName( 'PENALCODE' ).AsString := Penal.Value;
      FCD078A4DataSet.Post;
      { ?]???]?w Filter ??????, ?Y Key ????????, ?h Post ?????|?Q Filter ?o?? }
      { ?o?????Y?A?I?s Next ???k, ?????|???L???????? }
      if ( aRecordCount = FCD078A4DataSet.RecordCount ) then
        FCD078A4DataSet.Next;
    end;
    FCD078A4DataSet.First;
  finally
    FCD078A4DataSet.EnableControls;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.FilterCD078A1DataWhenCurrent;
begin
  FCD078A1DataSet.Filtered := False;
  FCD078A1DataSet.Filter := Format(
    'BPCODE=''%s'' AND CITEMCODE=''%s'' AND SERVICETYPE=''%s''',
    [FKeyValue, CItem.Value, ServiceType.Value] );
  FCD078A1DataSet.Filtered := True;
  
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.FilterCD078A2DataWhenCurrent;
begin
  FCD078A2DataSet.Filtered := False;
  FCD078A2DataSet.Filter := Format(
    'BPCODE=''%s'' AND CITEMCODE=''%s'' AND SERVICETYPE=''%s''',
    [FKeyValue, CItem.Value, ServiceType.Value] );
  FCD078A2DataSet.Filtered := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.FilterCD078A3DataWhenCurrent;
begin
  FCD078A3DataSet.Filtered := False;
  FCD078A3DataSet.Filter := Format(
    'BPCODE=''%s'' AND CITEMCODE=''%s'' AND SERVICETYPE=''%s''',
    [FKeyValue, CItem.Value, ServiceType.Value] );
  FCD078A3DataSet.Filtered := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.FilterCD078A4DataWhenCurrent;
begin
  FCD078A4DataSet.Filtered := False;
  FCD078A4DataSet.Filter := Format(
    'BPCODE=''%s'' AND CITEMCODE=''%s'' AND SERVICETYPE=''%s''',
    [FKeyValue, CItem.Value, ServiceType.Value] );
  FCD078A4DataSet.Filtered := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.FilterCD078A1DataWhenNew;
begin
  FCD078A1DataSet.Filtered := False;
  FCD078A1DataSet.Filter := Format(
    'BPCODE=''%s'' AND CITEMCODE=''-1'' AND SERVICETYPE=''-1''', [FKeyValue] );
  FCD078A1DataSet.Filtered := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.FilterCD078A2DataWhenNew;
begin
  FCD078A2DataSet.Filtered := False;
  FCD078A2DataSet.Filter := Format(
    'BPCODE=''%s'' AND CITEMCODE=''-1'' AND SERVICETYPE=''-1''', [FKeyValue] );
  FCD078A2DataSet.Filtered := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.FilterCD078A3DataWhenNew;
begin
  FCD078A3DataSet.Filtered := False;
  FCD078A3DataSet.Filter := Format(
    'BPCODE=''%s'' AND CITEMCODE=''-1'' AND SERVICETYPE=''-1''', [FKeyValue] );
  FCD078A3DataSet.Filtered := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.FilterCD078A4DataWhenNew;
begin
  FCD078A4DataSet.Filtered := False;
  FCD078A4DataSet.Filter := Format(
    'BPCODE=''%s'' AND CITEMCODE=''-1'' AND SERVICETYPE=''-1''', [FKeyValue] );
  FCD078A4DataSet.Filtered := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.FilterLinkKeyDataSet;
begin
  FLinkKeyDataSet.Filtered := False;
  FLinkKeyDataSet.Filter := Format(
    'BPCODE=''%s'' AND CITEMCODE=''%s'' AND  SERVICETYPE=''%s''',
    [FKeyValue, MCItem.Value, ServiceType.Value] );
  FLinkKeyDataSet.Filtered := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.AddCD078A1DefaultRecord;

   { --------------------------------------------- }

   function OnlyHaveDefaultOne: Boolean;
   begin
     Result := False;
     if ( FCD078A1DataSet.RecordCount = 1 ) then
     begin
       Result :=
         ( FCD078A1DataSet.FieldByName( 'MON1' ).AsString <> EmptyStr ) and
         ( FCD078A1DataSet.FieldByName( 'MON2' ).AsString = EmptyStr ) and
         ( FCD078A1DataSet.FieldByName( 'MON3' ).AsString = EmptyStr ) and
         ( FCD078A1DataSet.FieldByName( 'MON4' ).AsString = EmptyStr ) and
         ( FCD078A1DataSet.FieldByName( 'MON5' ).AsString = EmptyStr ) and
         ( FCD078A1DataSet.FieldByName( 'MON6' ).AsString = EmptyStr ) and
         ( FCD078A1DataSet.FieldByName( 'MON7' ).AsString = EmptyStr ) and
         ( FCD078A1DataSet.FieldByName( 'MON8' ).AsString = EmptyStr ) and
         ( FCD078A1DataSet.FieldByName( 'MON9' ).AsString = EmptyStr ) and
         ( FCD078A1DataSet.FieldByName( 'MON10' ).AsString = EmptyStr ) and
         ( FCD078A1DataSet.FieldByName( 'MON11' ).AsString = EmptyStr ) and
         ( FCD078A1DataSet.FieldByName( 'MON12' ).AsString = EmptyStr );
     end;
   end;

   { --------------------------------------------- }


begin
  if ( FMode in [dmInsert, dmUpdate] ) then
  begin
    if ( FCD078A1DataSet.IsEmpty ) then
    begin
      FCD078A1DataSet.Append;
      {}
      FCD078A1DataSet.FieldByName( 'BPCODE' ).AsString := FKeyValue;
      FCD078A1DataSet.FieldByName( 'CITEMCODE' ).AsString := CItem.Value;
      FCD078A1DataSet.FieldByName( 'SERVICETYPE' ).AsString := ServiceType.Value;
      FCD078A1DataSet.FieldByName( 'FACIITEM' ).AsString := FaciItem.Value;
      FCD078A1DataSet.FieldByName( 'STEPNO' ).AsString := DBController.GetStepNo;
      FCD078A1DataSet.FieldByName( 'DESCRIPTION' ).AsString := '?w?]???O';
      FCD078A1DataSet.FieldByName( 'MASTERSALE' ).AsString := '1';
      {}
      FCD078A1DataSet.FieldByName( 'MON1' ).AsString := '12';
      FCD078A1DataSet.FieldByName( 'PERIOD1' ).AsInteger := FCD019Data.Period;
      FCD078A1DataSet.FieldByName( 'RATETYPE1' ).AsString := '2';
      FCD078A1DataSet.FieldByName( 'DISCOUNTAMT1' ).AsFloat := FCD019Data.Amount;
      FCD078A1DataSet.FieldByName( 'MONTHAMT1' ).AsFloat := FCD019Data.MonthAmt;
      FCD078A1DataSet.FieldByName( 'DAYAMT1' ).AsFloat := FCD019Data.DayAmt;
      {}
      FCD078A1DataSet.Post;
    end else
    if ( OnlyHaveDefaultOne ) then
    begin
      FCD078A1DataSet.Edit;
      {}
      FCD078A1DataSet.FieldByName( 'PERIOD1' ).AsInteger := FCD019Data.Period;
      FCD078A1DataSet.FieldByName( 'DISCOUNTAMT1' ).AsFloat := FCD019Data.Amount;
      FCD078A1DataSet.FieldByName( 'MONTHAMT1' ).AsFloat := FCD019Data.MonthAmt;
      FCD078A1DataSet.FieldByName( 'DAYAMT1' ).AsFloat := FCD019Data.DayAmt;
      {}
      FCD078A1DataSet.Post;
    end;
    UpdateMasterSale;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.BuildCD078A3Record;
var
  aHigh, aIndex, aPriorIndex, aItemSeq, aDecreaseAmt, aRealAmt: Integer;
begin
  Screen.Cursor := crHourGlass;
  try
    EmptyCD078A3;
    aPriorIndex := 0;
    aItemSeq := 0;
    FCD078A4DataSet.First;
    while not FCD078A4DataSet.Eof do
    begin
      aHigh := ( ( FCD078A4DataSet.FieldByName( 'MONTHSTOP' ).AsInteger -
        FCD078A4DataSet.FieldByName( 'MONTHSTART' ).AsInteger ) + 1 );
      aDecreaseAmt := FCD078A4DataSet.FieldByName( 'DECREASEAMT' ).AsInteger;
      aRealAmt := FCD078A4DataSet.FieldByName( 'MONTHAMT' ).AsInteger;
      for aIndex := 1 to aHigh do
      begin
        Inc( aItemSeq );
        FCD078A3DataSet.Append;
        FCD078A3DataSet.FieldByName( 'BPCODE' ).AsString := FKeyValue;
        FCD078A3DataSet.FieldByName( 'CITEMCODE' ).AsString := FCD078A4DataSet.FieldByName( 'CITEMCODE' ).AsString;
        FCD078A3DataSet.FieldByName( 'SERVICETYPE' ).AsString := FCD078A4DataSet.FieldByName( 'SERVICETYPE' ).AsString;
        FCD078A3DataSet.FieldByName( 'PENALCODE' ).AsString := FCD078A4DataSet.FieldByName( 'PENALCODE' ).AsString;
        FCD078A3DataSet.FieldByName( 'ITEM' ).AsInteger := aItemSeq;
        FCD078A3DataSet.FieldByName( 'PENALMON' ).AsInteger := ( aPriorIndex + aIndex );
        if ( aIndex > 1 ) then
          Dec( aRealAmt, aDecreaseAmt );
        FCD078A3DataSet.FieldByName( 'PENALAMT' ).AsInteger := aRealAmt;
        FCD078A3DataSet.Post;
      end;
      Inc( aPriorIndex, aHigh );
      FCD078A4DataSet.Next;
      Application.ProcessMessages;
    end;
    FCD078A3DataSet.First;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.UpdateCD078A1LinkKey;
begin
  if ( MCItem.Value = EmptyStr ) then
  begin
    CD078A1DataSet.First;
    while not CD078A1DataSet.Eof do
    begin
      CD078A1DataSet.Edit;
      CD078A1DataSet.FieldByName( 'LinkKey' ).Clear;
      CD078A1DataSet.Post;
      CD078A1DataSet.Next;
    end;
    CD078A1DataSet.First;
  end else
  begin
    FilterLinkKeyDataSet;
    if ( CD078A1DataSet.RecordCount = 1 ) and
       ( FLinkKeyDataSet.RecordCount = 1 )  then
    begin
      CD078A1DataSet.First;
      CD078A1DataSet.Edit;
      CD078A1DataSet.FieldByName( 'LinkKey' ).AsString :=
        FLinkKeyDataSet.FieldByName( 'StepNo' ).AsString;
      CD078A1DataSet.FieldByName( 'LinkKeyName' ).AsString :=
        FLinkKeyDataSet.FieldByName( 'Description' ).AsString;
      CD078A1DataSet.Post;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.EmptyCD078A2;
begin
  FCD078A2DataSet.First;
  while not FCD078A2DataSet.Eof do
    FCD078A2DataSet.Delete;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.EmptyCD078A3;
begin
  FCD078A3DataSet.DisableControls;
  try
    FCD078A3DataSet.First;
    while not FCD078A3DataSet.Eof do
      FCD078A3DataSet.Delete;
  finally
    FCD078A3DataSet.EnableControls;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.EmptyCD078A4;
begin
  FCD078A4DataSet.First;
  while not FCD078A4DataSet.Eof do
    FCD078A4DataSet.Delete;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.gvPenalCol2PropertiesValidate(Sender: TObject;
  var DisplayValue: Variant; var ErrorText: TCaption; var Error: Boolean);
var
  aValue: String;
begin
  aValue := VarToStrDef( DisplayValue, '0' );
  if ( aValue = '0' ) then
  begin
    ErrorText := '???????J?i?????j???????j, ?B???i???s';
    WarningMsg( ErrorText );
    ErrorText := EmptyStr;
    Error := True;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.gvPenalCol3PropertiesValidate(Sender: TObject;
  var DisplayValue: Variant; var ErrorText: TCaption; var Error: Boolean);
var
  aValue: String;
  aInt: Integer;
begin
  aValue := VarToStrDef( DisplayValue, EmptyStr );
  if ( aValue = EmptyStr ) then Exit;
  if TryStrToInt( aValue, aInt ) then
  begin
    if ( aInt < 0 ) then
    begin
      ErrorText := '?i?H?????B?j???i???J?t??,?????s???J';
      WarningMsg( ErrorText );
      ErrorText := EmptyStr;
      Error := True;
    end;
  end else
  begin
    if ( TcxCurrencyEdit( Sender ).Properties.DecimalPlaces = 0 ) then
    begin
       ErrorText := Format( '???J?????~, ??????????%d~%d',
        [Trunc( TcxCurrencyEdit( Sender ).Properties.MinValue ),
         Trunc( TcxCurrencyEdit( Sender ).Properties.MaxValue )] );
    end else
    begin
       ErrorText := Format( '???J?????~, ??????????%10.3f~%10.3f',
        [TcxCurrencyEdit( Sender ).Properties.MinValue,
         TcxCurrencyEdit( Sender ).Properties.MaxValue] );
    end;
    WarningMsg( ErrorText );
    ErrorText := EmptyStr;
    Error := True;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.gvPenalCustomDrawCell(Sender: TcxCustomGridTableView;
  ACanvas: TcxCanvas; AViewInfo: TcxGridTableDataCellViewInfo;
  var ADone: Boolean);
var
  aGridRecord: TcxGridDataRow;
  aMonth: Integer;
  aAmt: Double;
  aDrawRedFont: Boolean;
begin
  aGridRecord := TcxGridDataRow( AViewInfo.GridRecord );
  if ( not aGridRecord.Focused ) then
  begin
    ACanvas.Brush.Color := clBtnFace;
    ACanvas.Font.Color := clDefault;
    ACanvas.FillRect( AViewInfo.ContentBounds );
  end;
  aMonth := 0;
  if ( VarToStrDef( aGridRecord.Values[gvPenalCol2.Index], EmptyStr ) <> EmptyStr ) then
    aMonth := VarAsType( aGridRecord.Values[gvPenalCol2.Index], varInteger );
  {}
  aAmt := 0;
  if ( VarToStrDef( aGridRecord.Values[gvPenalCol3.Index], EmptyStr ) <> EmptyStr ) then
    aAmt := VarAsType( aGridRecord.Values[gvPenalCol3.Index], varDouble );
  {}
  aDrawRedFont :=
    ( ( chkBundle.Checked and ( aMonth > StrToIntDef( edtBundleMon.Text, 0 ) ) ) or
      ( aAmt <= 0 ) ) and ( not aGridRecord.Focused  );
  {}
  if ( AViewInfo.Item = gvPenalCol2 ) or ( AViewInfo.Item = gvPenalCol3 ) then
  begin
    if aDrawRedFont then
      ACanvas.Font.Color := clRed;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.CD078A3DataSetNewRecord(DataSet: TDataSet);
begin
  FCD078A3DataSet.FieldByName( 'BPCODE' ).AsString := FKeyValue;
  FCD078A3DataSet.FieldByName( 'CITEMCODE' ).AsString := CItem.Value;
  FCD078A3DataSet.FieldByName( 'SERVICETYPE' ).AsString := ServiceType.Value;
  FCD078A3DataSet.FieldByName( 'PENALCODE' ).AsString := Penal.Value;
  FCD078A3DataSet.FieldByName( 'ITEM' ).AsInteger :=( FCD078A3DataSet.RecordCount + 1 );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.btnChangeCItemCodeClick(Sender: TObject);
var
  aPoint: TPoint; 
begin
  frmCD078B3_4 := TfrmCD078B3_4.Create( nil );
  try
    frmCD078B3_4.Caption :=  '???????O????[frmCD078B32]';
    frmCD078B3_4.BPCode := FKeyValue;
    frmCD078B3_4.ServiceType := ServiceType.Value;
    frmCD078B3_4.CItemCode := CItem.Value;
    frmCD078B3_4.AlreadySetCItemValue := FAlreadySetCItemValue;
    aPoint := Label1.ClientToScreen( Point( Label1.Left, Label1.Top ) );
    frmCD078B3_4.Left := aPoint.X;
    frmCD078B3_4.Top := aPoint.Y + 3;
    if ( frmCD078B3_4.ShowModal = mrOk ) then
    begin
      CItem.CodeNo.LockChangeEvents( True, False );
      CItem.CodeName.LockChangeEvents( True, False );
      try
        UpdateMCItemDataSet( CItem.Value );
        FAlreadySetCItemValue := DeleteValue( CItem.Value, FAlreadySetCItemValue );
        CItem.Value := frmCD078B3_4.CItem.Value;
        CItem.ValueName := frmCD078B3_4.CItem.ValueName;
        FAlreadySetCItemValue := AddValue( CItem.Value, FAlreadySetCItemValue );
        InitCD019Data;
        LoadCD019Data;
        CD019DataToFieldValue;
        ChangeRelationItemState;
        ChangeFaciItemState;
        ChangeExpireTypeItem;
        {}
        UpdateCD078A1KeyValue;
        UpdateCD078A2KeyValue;
        UpdateCD078A3KeyValue;
        UpdateCD078A4KeyValue;
        {}
        FilterCD078A1DataWhenCurrent;
        FilterCD078A2DataWhenCurrent;
        FilterCD078A3DataWhenCurrent;
        FilterCD078A4DataWhenCurrent;
        {}
        AddCD078A1DefaultRecord;
        UpdateMasterSale;
        UpdateMCItemDataSet( EmptyStr );
      finally
        CItem.CodeNo.LockChangeEvents( False, False );
        CItem.CodeName.LockChangeEvents( False, False );
      end;
    end;
  finally
    frmCD078B3_4.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.btnPenalConditionClick(Sender: TObject);
var
  aKeys: array [1..4] of String;
begin
  if not VdCanShowPenalRule then Exit;
  {}
  aKeys[1] := FKeyValue;
  aKeys[2] := CItem.Value;
  aKeys[3] := ServiceType.Value;
  aKeys[4] := Penal.Value;
  {}
  if ( FMode = dmInsert ) then FilterCD078A4DataWhenCurrent;
  {}
  frmCD078B3_5 := TfrmCD078B3_5.Create( FMode, FCD078A4DataSet, aKeys );
  try
    frmCD078B3_5.Caption := '?H?????????W?h?]?w[frmCD078B33]';
    if ( frmCD078B3_5.ShowModal = mrOk ) then
      BuildCD078A3Record;
  finally
    frmCD078B3_5.Free;
  end;

end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B3.edtMonthAmtPropertiesEditValueChanged(
  Sender: TObject);
  var aMonthAmt,aDayAmt :Double;
begin
  if FLockCheck then exit;
  aMonthAmt :=0;
  aDayAmt :=0;
  try
    edtMonthAmt.LockChangeEvents(True,False);
    if edtMonthAmt.EditValue<>Null then
    begin
      aMonthAmt:= Abs(edtMonthAmt.Value);
      aDayAmt := CbRoundTo(aMonthAmt/30,3);
    end;
    if FCD019Data.IsMinus then
    begin
      aMonthAmt := aMonthAmt *-1;
      aDayAmt := aDayAmt * -1;
    end;
  finally
    edtDayAmt.Value := aDayAmt;
    edtMonthAmt.Value := aMonthAmt;
    edtMonthAmt.LockChangeEvents(False,False);
  end;

end;

procedure TfrmCD078B3.MSalePointCodeNamePropertiesChange(Sender: TObject);
begin
  MSalePoint.CodeNamePropertiesChange( Sender );
end;

procedure TfrmCD078B3.MSalePointCodeNamePropertiesInitPopup(
  Sender: TObject);
begin
  //#5318 ??????????,?u???L?o?D???????N?n By Kin 2009/10/30
  // Penny?V?T???????O?n?L?o POINTDOU=0?????? By Kin 2009/11/06
  Screen.Cursor := crSQLWait;
  try
    MSalePoint.SQL.Text := Format(
    ' SELECT CODENO, DESCRIPTION    ' +
    '  FROM %s.CD107                ' +
    '  WHERE STOPFLAG <> 1          ' +
    '  AND POINTDOU = 0             ' ,
    [DBController.LoginInfo.DbAccount] );
    MSalePoint.CodeNamePropertiesInitPopup( Sender );
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TfrmCD078B3.MSalePointCodeNoPropertiesChange(Sender: TObject);
begin
  MSalePoint.CodeNoPropertiesChange( Sender );

end;

procedure TfrmCD078B3.btnRankWhereClick(Sender: TObject);
var
  aKeys: array [1..4] of String;
begin
  {#5257 ?s?W???\?? By Kin 2010/05/05}
  if not VdCanShowMultiRank then Exit;
  FilterCD078A1DataWhenCurrent;
  aKeys[1] := FKeyValue;
  aKeys[2] := CItem.Value;
  aKeys[3] := ServiceType.Value;
  aKeys[4] := Penal.Value;
  if ( SameText( FCD019Data.Sign, '-' ) ) and ( MCItem.Value <> EmptyStr ) then
  begin
    if MCItem.Value = EmptyStr then
    begin
      WarningMsg( '???]?w???????????O????' );
    end;
    FilterLinkKeyDataSet;
  end;
  frmCD078B3_7 := TfrmCD078B3_7.Create( FMode,FCD078A1DataSet,aKeys );
  try
    frmCD078B3_7.Caption := '???q???u?f?????]?w[frmCD078B37]';
    frmCD078B3_7.DefaultMonth := edtBundleMon.Text;
    frmCD078B3_7.DefaultAmt := edtMonthAmt.Text;
    frmCD078B3_7.LinkCitemCode := MCItem.Value;
    frmCD078B3_7.EnableLinkKey :=
      ( SameText( FCD019Data.Sign, '-' ) ) and ( MCItem.Value <> EmptyStr );
    frmCD078B3_7.LinkKeyDataSet := FLinkKeyDataSet;
    frmCD078B3_7.ShowModal;
    btnRankWhere.Enabled := False;
    UpdateMasterSale;
    btnRankWhere.Enabled := not  (FCD078A1DataSet.RecordCount > 0);
  finally
    frmCD078B3_7.Free;
  end;

end;

procedure TfrmCD078B3.edtBpStopDateExit(Sender: TObject);
  var aDate : string;
begin
  aDate := TMaskEdit( Sender ).EditText;
  if not DateTextIsValidEx( aDate ) then
  begin
    WarningMsg('?S?w?u?f?I?????????~');
    TMaskEdit( Sender ).SetFocus;
  end;
end;
{#6865 ?????????O?n?P?B???t?? By Kin 2014/10/07}
procedure TfrmCD078B3.SyncLongPayFlag;
   var aBMark: TBookmark;
      MCitemCode : String;
      aCitemCode : string;
begin

  aBMark := FDataSet.GetBookmark;
  MCitemCode := MCItem.Value;
  aCitemCode := CItem.Value;
  try
    if chkLONGPAYflag.Checked then
    begin
      FDataSet.First;
      while not FDataSet.Eof do
      begin
        if MCitemCode <> EmptyStr then
        begin
          if ( MCitemCode = FDataSet.FieldByName( 'CitemCode' ).AsString ) then
          begin
            FDataSet.Edit;
            FDataSet.FieldByName( 'LONGPAYflag' ).AsString := '1';
            FDataSet.Post;
          end;
        end else
        begin
          if aCitemCode = FDataSet.FieldByName( 'MCitemCode' ).AsString then
          begin
            FDataSet.Edit;
            FDataSet.FieldByName( 'LONGPAYflag' ).AsString := '1';
            FDataSet.Post;
          end;
        end;
        FDataSet.Next;
      end;
      FDataSet.GotoBookmark( aBMark );
    end;

  finally
    FDataSet.FreeBookmark( aBMark );
  end;
end;

end.
