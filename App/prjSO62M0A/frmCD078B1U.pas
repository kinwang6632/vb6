unit frmCD078B1U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, ActnList, ADODB,
  DB, DBClient,
  cbDBController, cbStyleController,
  cxPC, cxControls, cxContainer, cxEdit,
  cxTextEdit, cxMaskEdit, cxCheckBox, cxMemo, cxDropDownEdit, cxGridCustomView,
  cxGridCustomTableView, cxGridTableView, cxGridBandedTableView,
  cxGridDBBandedTableView, cxClasses, cxGridLevel, cxGrid, Provider, 
  cxGraphics, cxDataStorage, cxDBData,
  cxImageComboBox, cxCurrencyEdit, cxGridCustomPopupMenu, 
  dxSkinsCore, dxSkinsDefaultPainters, dxSkinscxPCPainter, cxStyles,
  cxCustomData, cxFilter, cxData, dxSkinBlack, dxSkinBlue, dxSkinCaramel,
  dxSkinCoffee, cxLabel, cbDataLookup;


type
  TDML = ( dmBrowser = 0,
           dmInsert = 1,
           dmUpdate = 2,
           dmDelete = 3,
           dmCancel = 4,
           dmSave = 5 );

  TfrmCD078B1 = class(TForm)
    Panel1: TPanel;
    Panel3: TPanel;
    Bevel1: TBevel;
    btnSave: TButton;
    ActionList1: TActionList;
    actSave: TAction;
    MainPageControl: TcxPageControl;
    Sheet2: TcxTabSheet;
    Sheet3: TcxTabSheet;
    Sheet4: TcxTabSheet;
    Sheet1: TcxTabSheet;
    Bevel4: TBevel;
    CD078DataSet: TClientDataSet;
    CD078ADataSet: TClientDataSet;
    CD078ADataSource: TDataSource;
    CD078BDataSet: TClientDataSet;
    CD078BDataSource: TDataSource;
    CD078DDataSet: TClientDataSet;
    CD078DDataSource: TDataSource;
    ScrollBox3: TScrollBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    edtCodeNo: TcxTextEdit;
    edtDescription: TcxTextEdit;
    edtDateSt: TcxMaskEdit;
    edtDateEd: TcxMaskEdit;
    chkMasterProd: TcxCheckBox;
    chkStopFlag: TcxCheckBox;
    edtRefNo: TcxTextEdit;
    Shape1: TShape;
    Panel2: TPanel;
    CD078DGrid: TcxGrid;
    CD078DBandedTableView: TcxGridDBBandedTableView;
    CD078DGridLevel: TcxGridLevel;
    Panel4: TPanel;
    CD078AGrid: TcxGrid;
    CD078ABandedTableView: TcxGridDBBandedTableView;
    CD078AGridLevel: TcxGridLevel;
    Panel5: TPanel;
    CD078BGrid: TcxGrid;
    CD078BBandedTableView: TcxGridDBBandedTableView;
    CD078BGridLevel: TcxGridLevel;
    Panel6: TPanel;
    CD078CGrid: TcxGrid;
    CD078CBandedTableView: TcxGridDBBandedTableView;
    CD078CGridLevel: TcxGridLevel;
    CD078CDataSet: TClientDataSet;
    CD078CDataSource: TDataSource;
    CD078Provider: TDataSetProvider;
    edtDML: TcxTextEdit;
    actInsert: TAction;
    btnInsert: TButton;
    btnUpdate: TButton;
    actUpdate: TAction;
    CD078DBandedTableViewFACIITEM: TcxGridDBBandedColumn;
    CD078DBandedTableViewFACILTYPE: TcxGridDBBandedColumn;
    CD078DBandedTableViewFACILNAME: TcxGridDBBandedColumn;
    CD078DBandedTableViewSERVICETYPE: TcxGridDBBandedColumn;
    CD078DBandedTableViewMODELNAME: TcxGridDBBandedColumn;
    CD078DBandedTableViewBUYNAME: TcxGridDBBandedColumn;
    CD078DBandedTableViewCMBAUDRATE: TcxGridDBBandedColumn;
    CD078DBandedTableViewFIXIPCOUNT: TcxGridDBBandedColumn;
    CD078DBandedTableViewDYNIPCOUNT: TcxGridDBBandedColumn;
    CD078ABandedTableViewCITEMNAME: TcxGridDBBandedColumn;
    CD078ABandedTableViewSERVICETYPE: TcxGridDBBandedColumn;
    CD078ABandedTableViewBUNDLE: TcxGridDBBandedColumn;
    CD078ABandedTableViewBUNDLEMON: TcxGridDBBandedColumn;
    CD078ABandedTableViewPENALTYPE: TcxGridDBBandedColumn;
    CD078ABandedTableViewEXPIRETYPE: TcxGridDBBandedColumn;
    CD078ABandedTableViewDEPOSITATTR: TcxGridDBBandedColumn;
    CD078ABandedTableViewDEPOSITNAME: TcxGridDBBandedColumn;
    CD078ABandedTableViewDEPOSITAMT: TcxGridDBBandedColumn;
    CD078BBandedTableViewINSTNAME: TcxGridDBBandedColumn;
    CD078BBandedTableViewSERVICETYPE: TcxGridDBBandedColumn;
    CD078BBandedTableViewGROUPNAME: TcxGridDBBandedColumn;
    CD078BBandedTableViewWORKUNIT: TcxGridDBBandedColumn;
    CD078BBandedTableViewREFNO: TcxGridDBBandedColumn;
    CD078BBandedTableViewINTERDEPEND: TcxGridDBBandedColumn;
    CD078CBandedTableViewCITEMNAME: TcxGridDBBandedColumn;
    CD078CBandedTableViewSERVICETYPE: TcxGridDBBandedColumn;
    CD078CBandedTableViewFACIITEM: TcxGridDBBandedColumn;
    CD078CBandedTableViewAMOUNT: TcxGridDBBandedColumn;
    CD078CBandedTableViewDISCOUNTAMT: TcxGridDBBandedColumn;
    CD078CBandedTableViewPUNISH: TcxGridDBBandedColumn;
    CD078CBandedTableViewPENALTYPE: TcxGridDBBandedColumn;
    CD078CBandedTableViewINSTCODESTR: TcxGridDBBandedColumn;
    btnDelete: TButton;
    actDelete: TAction;
    actCancel: TAction;
    actClose: TAction;
    btnCancel: TButton;
    OtherDataSet: TClientDataSet;
    InstCodeDataSet: TClientDataSet;
    CD078A1DataSet: TClientDataSet;
    CD078A1CloneDataSet: TClientDataSet;
    MemoPageControl: TcxPageControl;
    cxTabSheet1: TcxTabSheet;
    cxTabSheet2: TcxTabSheet;
    edtBundleProdNote: TcxMemo;
    edtWorkNote: TcxMemo;
    cxTabSheet3: TcxTabSheet;
    edtConditional: TcxMemo;
    CD078A2DataSet: TClientDataSet;
    CD078A3DataSet: TClientDataSet;
    CD078A2CloneDataSet: TClientDataSet;
    CD078A3CloneDataSet: TClientDataSet;
    CD078B1CloneDataSet: TClientDataSet;
    CD078B1DataSet: TClientDataSet;
    CD078ABandedTableViewPENAL: TcxGridDBBandedColumn;
    CD078A4DataSet: TClientDataSet;
    CD078A4CloneDataSet: TClientDataSet;
    MCItemDataSet: TClientDataSet;
    CD078ABandedTableViewSIGN: TcxGridDBBandedColumn;
    LinkKeyDataSet: TClientDataSet;
    Label6: TLabel;
    edtKindFunction: TcxComboBox;
    chkChooseKind: TcxCheckBox;
    CD078GDataSet: TClientDataSet;
    CD078G1DataSet: TClientDataSet;
    Sheet5: TcxTabSheet;
    Panel7: TPanel;
    cxGrid1: TcxGrid;
    CD078GBandedTableView: TcxGridDBBandedTableView;
    CD078GBandedTableViewCITEMNAME: TcxGridDBBandedColumn;
    CD078GBandedTableViewSERVICETYPE: TcxGridDBBandedColumn;
    CD078GBandedTableViewFACIITEM: TcxGridDBBandedColumn;
    CD078GBandedTableViewSALEPOINTNAME: TcxGridDBBandedColumn;
    cxGridLevel1: TcxGridLevel;
    CD078GDataSource: TDataSource;
    CD078G1CloneDataSet: TClientDataSet;
    chkPayKind: TcxCheckBox;
    Sheet6: TcxTabSheet;
    CD078IDataSet: TClientDataSet;
    CD078ICloneDataSet: TClientDataSet;
    CD078IGrid: TcxGrid;
    CD078IBandedTableView: TcxGridDBBandedTableView;
    cxGridDBBandedColumn1: TcxGridDBBandedColumn;
    cxGridDBBandedColumn2: TcxGridDBBandedColumn;
    cxGridLevel2: TcxGridLevel;
    SO562DataSet: TClientDataSet;
    SO562Provider: TDataSetProvider;
    SO562DataSource: TDataSource;
    SO562_1_DataSet: TClientDataSet;
    cxGridDBBandedColumn3: TcxGridDBBandedColumn;
    cxGridDBBandedColumn4: TcxGridDBBandedColumn;
    cxGridDBBandedColumn5: TcxGridDBBandedColumn;
    cxGridDBBandedColumn6: TcxGridDBBandedColumn;
    cxGridDBBandedColumn7: TcxGridDBBandedColumn;
    cxGridDBBandedColumn8: TcxGridDBBandedColumn;
    cxGridDBBandedColumn9: TcxGridDBBandedColumn;
    cxGridDBBandedColumn10: TcxGridDBBandedColumn;
    cxGridDBBandedColumn11: TcxGridDBBandedColumn;
    cxGridDBBandedColumn12: TcxGridDBBandedColumn;
    cxGridDBBandedColumn13: TcxGridDBBandedColumn;
    cxGridDBBandedColumn14: TcxGridDBBandedColumn;
    Sheet7: TcxTabSheet;
    CD078JGrid: TcxGrid;
    CD078JBandedTableView: TcxGridDBBandedTableView;
    cxGridDBBandedColumn16: TcxGridDBBandedColumn;
    cxGridDBBandedColumn18: TcxGridDBBandedColumn;
    cxGridDBBandedColumn20: TcxGridDBBandedColumn;
    cxGridDBBandedColumn21: TcxGridDBBandedColumn;
    cxGridDBBandedColumn23: TcxGridDBBandedColumn;
    cxGridLevel3: TcxGridLevel;
    CD078JDataSet: TClientDataSet;
    CD078JDataSource: TDataSource;
    SyncBPCode: TDataLookup;
    Label7: TLabel;
    lbl1: TLabel;
    edtImgURL: TcxTextEdit;
    lbl2: TLabel;
    edtEPGOrder: TcxTextEdit;
    lbl3: TLabel;
    cobCanUseType: TcxComboBox;
    CD078ABandedTableViewSTOPFLAG: TcxGridDBBandedColumn;
    CD078ABandedTableViewLongPayFlag: TcxGridDBBandedColumn;
    btnIFRS: TButton;
    CD078KDataSet: TClientDataSet;
    btnShare: TButton;
    CD078LDataSet: TClientDataSet;
    chkByHouse: TcxCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure actSaveExecute(Sender: TObject);
    procedure CD078DataSetAfterScroll(DataSet: TDataSet);
    procedure actInsertExecute(Sender: TObject);
    procedure actUpdateExecute(Sender: TObject);
    procedure actDeleteExecute(Sender: TObject);
    procedure actCancelExecute(Sender: TObject);
    procedure actCloseExecute(Sender: TObject);
    procedure DataSetAfterDeleteOrPost(DataSet: TDataSet);
    procedure MainPageControlDrawTab(AControl: TcxCustomTabControl;
      ATab: TcxTab; var DefaultDraw: Boolean);
    procedure CD078ABandedTableViewDEPOSITATTRGetDisplayText(
      Sender: TcxCustomGridTableItem; ARecord: TcxCustomGridRecord;
      var AText: String);
    procedure CD078ABandedTableViewEXPIRETYPEGetDisplayText(
      Sender: TcxCustomGridTableItem; ARecord: TcxCustomGridRecord;
      var AText: String);
    procedure CD078ABandedTableViewPENALTYPEGetDisplayText(
      Sender: TcxCustomGridTableItem; ARecord: TcxCustomGridRecord;
      var AText: String);
    procedure edtDateStPropertiesValidate(Sender: TObject;
      var DisplayValue: Variant; var ErrorText: TCaption;
      var Error: Boolean);
    procedure MainPageControlChange(Sender: TObject);
    procedure BandedTableViewDblClick(Sender: TObject);
    procedure CD078ABandedTableViewCustomDrawCell(
      Sender: TcxCustomGridTableView; ACanvas: TcxCanvas;
      AViewInfo: TcxGridTableDataCellViewInfo; var ADone: Boolean);
    procedure CD078CBandedTableViewCustomDrawCell(
      Sender: TcxCustomGridTableView; ACanvas: TcxCanvas;
      AViewInfo: TcxGridTableDataCellViewInfo; var ADone: Boolean);
    procedure edtKindFunctionPropertiesChange(Sender: TObject);
    procedure CD078GBandedTableViewDblClick(Sender: TObject);
    procedure cxGridDBBandedColumn2GetDisplayText(
      Sender: TcxCustomGridTableItem; ARecord: TcxCustomGridRecord;
      var AText: String);
    procedure cxGridDBBandedColumn3GetDisplayText(
      Sender: TcxCustomGridTableItem; ARecord: TcxCustomGridRecord;
      var AText: String);
    procedure CD078IBandedTableViewCustomDrawCell(
      Sender: TcxCustomGridTableView; ACanvas: TcxCanvas;
      AViewInfo: TcxGridTableDataCellViewInfo; var ADone: Boolean);
    procedure cxGridDBBandedColumn7GetDisplayText(
      Sender: TcxCustomGridTableItem; ARecord: TcxCustomGridRecord;
      var AText: String);
    procedure SyncBPCodeCodeNamePropertiesInitPopup(Sender: TObject);
    procedure SyncBPCodeCodeNamePropertiesChange(Sender: TObject);
    procedure SyncBPCodeCodeNoPropertiesChange(Sender: TObject);
    procedure edtEPGOrderKeyPress(Sender: TObject; var Key: Char);
    procedure cobCanUseTypePropertiesEditValueChanged(Sender: TObject);
    procedure cobCanUseTypePropertiesChange(Sender: TObject);
    procedure btnIFRSClick(Sender: TObject);
    procedure btnShareClick(Sender: TObject);
  private
    { Private declarations }
    FMode: TDML;
    FKeyValue: String;
    FCanModify : Boolean;
    FReader: TADOQuery;
    FForceSave: Boolean;
    FCanUseEPG : Boolean;
    FCanAddEPG : Boolean;
    function DataSetStateChange(const aMode: TDML): Boolean;
    procedure DMLModeChange(aMode: TDML);
    procedure ButtonStateChange;
    procedure EditorStateChange;
    function OneOrZero(Status: Boolean): String;
    { Master Table }
    procedure LoadCD078Data;
    procedure ClearMasterEditor;
    procedure MasterDataToEditor;
    { Detail Table }
    procedure LoadCD078A_Data;
    procedure LoadCD078A1_Data;
    procedure LoadCD078A2_Data;
    procedure LoadCD078A3_Data;
    procedure LoadCD078A4_Data;
    procedure LoadCD078B_Data;
    procedure LoadCD078B1_Data;
    procedure LoadCD078C_Data;
    procedure LoadCD078D_Data;
    procedure LoadCD078G_Data;
    procedure LoadCD078G1_Data;
    procedure LoadCD078I_Data;
    procedure LoadCD078J_Data;
    procedure LoadCD078K_Data;
    procedure LoadCD078L_Data;
    procedure LoadSO562_Data;
    procedure LoadSO562Citem_Data( const aLoad : Boolean ) ;
    procedure PrepareSO562DataSet;
    {}
    procedure ShowDetailForm(aMode: TDML);
    procedure ShowCD078A_From(aMode: TDML);
    procedure ShowCD078B_From(aMode: TDML);
    procedure ShowCD078C_From(aMode: TDML);
    procedure ShowCD078D_Form(aMode: TDML);
    procedure ShowCD078E_Form(aMode: TDML);
    procedure ShowCD078I_Form(aMode: TDML);
    procedure ShowCD078J_Form(aMode: TDML);
    {}
    procedure CloneCodeData;
    procedure CloneInstCodeData;
    procedure LoadMCItetmCodeData;
    procedure RemoveInstCodeStr(const AInstCode: String);
    {}
    function ApplyCD078: Boolean;
    function ApplyCD078A: Boolean;
    function ApplyCD078A1: Boolean;
    function ApplyCD078A2: Boolean;
    function ApplyCD078A3: Boolean;
    function ApplyCD078A4: Boolean;
    function ApplyCD078B: Boolean;
    function ApplyCD078B1: Boolean;
    function ApplyCD078C: Boolean;
    function ApplyCD078D: Boolean;
    function ApplyCD078G: Boolean;
    function ApplyCD078G1: Boolean;
    function ApplyCD078I: Boolean;
    function ApplyCD078J: Boolean;
    function ApplyCD078K: Boolean;
    function ApplyCD078L: Boolean;
    {}
    function ServiceTypeIntoStr: String;
    function DoDeleteCheck: Boolean;
    function IsServiceTypeExists(AServiceType: String): Boolean;
    function IsFaciItemExists(AFaciItem: String): Boolean; overload;
    function IsFaciItemExists(AFaciItem: String; aDataSet: TClientDataSet): Boolean; overload;
    function IsRefCitemExists(ACitem: String): Boolean;
    function GetFaciItemRefNo(AFaciItem: String): String;
    function GetItemByHouse(ACitem: String): Integer;
    function GetKindFunction: string;
    {}
    function DataValidate(var AErrSource: String): Boolean;
    function VdMustInput: Boolean;
    function VdCD078A: Boolean;
    function VdCD078A1: Boolean;
    function VdCD078A2: Boolean;
    function VdCD078B: Boolean;
    function VdCD078C: Boolean;
    function VdCD078D: Boolean;
    {}
    function ApplyDataSet: Boolean;
  public
    { Public declarations }
    constructor Create(aMode: TDML; aKeyValue: String); reintroduce;
    procedure CloneCD078A1Data;
    procedure CloneCD078A2Data;
    procedure CloneCD078A3Data;
    procedure CloneCD078A4Data;
    procedure CloneCD078B1Data;
    procedure CloneCD078G1Data;
    procedure CloneCD078IData;
    procedure RestoreLinkKeyData;
    {}
    procedure FilterCD078A1Data(const AEnableFilter: Boolean; const aMode: TDML);
    procedure FilterCD078G1Data(const AEnableFilter: Boolean; const aMode: TDML);
    procedure FilterSO562Data(const AEnableFilter: Boolean;const aMode: TDML);
    procedure ChangeSO562Data(const aOldCitemCode: String;const aMode: TDML);
    procedure RestoreCD078A1Data;
    procedure RestoreCD078G1Data;
    procedure DeleteCD078A1Data;
    procedure DeleteCD078G1Data;
    {}
    procedure FilterCD078A2Data(const AEnableFilter: Boolean; const aMode: TDML);
    procedure RestoreCD078A2Data;
    procedure DeleteCD078A2Data;
    {}
    procedure FilterCD078A3Data(const AEnableFilter: Boolean; const aMode: TDML);
    procedure RestoreCD078A3Data;
    procedure DeleteCD078A3Data;
    {}
    procedure FilterCD078A4Data(const AEnableFilter: Boolean; const aMode: TDML);
    procedure RestoreCD078A4Data;
    procedure DeleteCD078A4Data;
    procedure DeleteSO562Data;
    {}
    procedure FilterCD078B1Data(const AEnableFilter: Boolean; const aMode: TDML);
    procedure RestoreCD078B1Data;
    procedure DeleteCD078B1Data;
    procedure DeleteCD078CIfInstCodeIsEmpty;
    {}
    function GetCD078ACItemValue: String;
    function GetCD078GCItemValue: String;
    property KeyValue: String read FKeyValue;
    property KindFunction: string read GetKindFunction;
  end;


var
  frmCD078B1: TfrmCD078B1;
  TDMLString: array [TDML] of String =
    ( '顯示', '新增', '修改', '刪除', '取消', '存檔' );
  blnForceSave: Boolean;
implementation

uses cbUtilis, frmCD078B2U, frmCD078B3U, frmCD078B4U,
      frmCD078B5U,frmCD078G0U,frmCD078IU,frmCD078J1U,
      frmCD078K0U,frmCD078L0U;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

function ToACDateText(aRocDateText: String): String;
var
  aYear, aMonth, aDay: String;
begin
  aYear := Copy( aRocDateText, 1, Pos( '/', aRocDateText ) - 1 );
  Delete( aRocDateText, 1, Pos( '/', aRocDateText ) );
  aMonth := Copy( aRocDateText, 1, Pos( '/', aRocDateText ) - 1 );
  Delete( aRocDateText, 1, Pos( '/', aRocDateText ) );
  aDay := aRocDateText;
  {}
  aMonth := Lpad( aMonth, 2, '0' );
  aDay := Lpad( aDay, 2, '0' );
  Result := Format( '%s/%s/%s', [aYear, aMonth, aDay] );

end;



{ ---------------------------------------------------------------------------- }

constructor TfrmCD078B1.Create(aMode: TDML; aKeyValue: String);
begin
  inherited Create( Application );
  FMode := aMode;
  FKeyValue := aKeyValue;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.FormCreate(Sender: TObject);
begin
  //#6560 判斷使用者EPG權限是否可用 By Kin 2013/08/20
  FCanUseEPG := True;
//  FCanUseEPG := DBController.GetUserPriv( 'SO62M021' );
  FCanAddEPG := DBController.GetUserPriv( 'SO62M011' );
  FReader := DBController.DataReader;
  SyncBPCode.Initializa;
  if not DBController.CanModify then
  begin
    btnSave.Enabled := False;
    btnInsert.Enabled := False;
    btnUpdate.Enabled := False;
    btnDelete.Enabled := False;
    btnIFRS.Enabled := False;
    btnShare.Enabled := False;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.FormShow(Sender: TObject);
 var aShow,aPvodShow:Boolean;
begin
  MainPageControl.TabWidth := ( MainPageControl.Width div MainPageControl.PageCount ) - 10;
  {#5318 判斷點數設定是否啟用 BY Kin 2009/10/06}

  FReader.Close;
  FReader.SQL.Text := Format('SELECT Nvl(STARTVOD,0) STARTVOD ,   ' +
              ' NVL(OrderPayNow,0) OrderPayNow ,                  ' +
              ' NVL(STARTPVOD,0) STARTPVOD                       ' +
              ' FROM %s.SO041 Where SYSID=%s',
       [ DBController.LoginInfo.DbAccount,DBController.LoginInfo.CompCode ]);
  FReader.Open;
  aShow :=( FReader.FieldByName( 'STARTVOD' ).AsInteger = 1);
  aPvodShow := ( FReader.FieldByName( 'STARTPVOD' ).AsInteger = 1 );
  chkPayKind.Enabled := FReader.FieldByName( 'OrderPayNow' ).AsString = '1';
  FReader.Close;
  {#6881 判斷是否有啟用IFRS機制 By Kin 2014/10/16}
  btnIFRS.Visible := False;
  FReader.SQL.Text := Format('SELECT NVL(USEIFRS,0) USEIFRS FROM %s.SO043 WHERE USEIFRS = 1',
        [DBController.LoginInfo.DbAccount] );
  FReader.Open;

  if FReader.RecordCount > 0 then
  begin
    FReader.First;
    btnIFRS.Visible := FReader.FieldByName( 'UseIFRS' ).AsString = '1';
  end;
  btnShare.Visible := btnIFRS.Visible;
  FReader.Close;

  {#6035 判斷是否有啟動PVOD By Kin 2011/06/28}
  DBController.LoginInfo.StarPVOD := aPvodShow;
  MainPageControl.Tabs[4].Visible := aShow;
  MainPageControl.Tabs[5].Visible := aPvodShow;
  LoadCD078Data;
  DMLModeChange( FMode );
  MainPageControl.ActivePageIndex := 0;
  MemoPageControl.ActivePageIndex := 0;

  {}

  //#5232 上架日期<=系統日期不允許修改該欄位 By Kin 2009/08/10
  if (edtDateSt.Text) <> EmptyStr then
  begin
    if  StrToDate( edtDateSt.Text )<=Now then
    begin
      edtDateSt.Enabled := False;
      //edtDateEd.Enabled := False;
    end;
  end;

  if not FCanAddEPG then
  begin
    if cobCanUseType.ItemIndex <> 1 then
    begin
      btnSave.Enabled := False;
      btnInsert.Enabled := False;
      btnUpdate.Enabled := False;
      btnDelete.Enabled := False;
    end;
  end;

  {#6560 測試不OK,如果沒有修改權限則不能下拉 By Kin 2013/09/05}
  if FMode in [dmUpdate] then
  begin
    if ( not FCanAddEPG ) then
      cobCanUseType.Enabled := False;
  end;
  {#6560 測試不OK,如果沒有新增權限則預設為1 By Kin 2013/09/05}
  if FMode in [dmInsert] then
  begin
    if ( not FCanAddEPG ) then
      cobCanUseType.ItemIndex := 1;
  end;

end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.FormDestroy(Sender: TObject);
begin
  FReader := nil;
  SyncBPCode.Finaliza;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.DMLModeChange(aMode: TDML);
begin
  if not DataSetStateChange( aMode ) then Exit;
  FMode := aMode;
  if ( aMode in [dmSave, dmCancel] ) then
    FMode := dmBrowser;
  ButtonStateChange;
  EditorStateChange;
  edtDML.Text := TDMLString[FMode];
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.ButtonStateChange;
var
  aDataSet: TClientDataSet;
begin
  case MainPageControl.ActivePageIndex of
    0: aDataSet := CD078BDataSet;
    1: aDataSet := CD078DDataSet;
    2: aDataSet := CD078ADataSet;
    3: aDataSet := CD078CDataSet;
    4: aDataSet := CD078GDataSet;
    5: aDataSet := SO562DataSet;
    6: aDataSet := CD078JDataSet;
  else
    aDataSet := nil;
  end;
  if ( FMode = dmInsert ) then
  begin
    actSave.Enabled := True;
    actInsert.Enabled := True;
    actDelete.Enabled := ( not aDataSet.IsEmpty );
    actUpdate.Enabled := ( not aDataSet.IsEmpty );
    actCancel.Enabled := True;

  end else
  if ( FMode = dmUpdate ) then
  begin
    actSave.Enabled := True;
    actInsert.Enabled := True;
    actDelete.Enabled := ( not aDataSet.IsEmpty );
    actUpdate.Enabled := ( not aDataSet.IsEmpty );
    actCancel.Enabled := True;
  end else
  begin
    actSave.Enabled := False;
    actInsert.Enabled := False;
    actDelete.Enabled := False;
    actUpdate.Enabled := False;
    btnIFRS.Enabled := False;
    btnShare.Enabled := False;
    btnCancel.Action := actClose;
  end;
  if ( aDataSet = SO562DataSet ) then
  begin
    actInsert.Enabled := False;
    actDelete.Enabled := False;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B1.DataSetStateChange(const aMode: TDML): Boolean;
var
  aErrSource, aErrMsg: String;
begin
  Result := True;
  aErrMsg := EmptyStr;
  case aMode of
    dmBrowser:
      begin
        CD078DataSetAfterScroll( CD078DataSet );
        CD078DataSet.ReadOnly := True;
        CD078ADataSet.ReadOnly := True;
        CD078BDataSet.ReadOnly := True;
        CD078CDataSet.ReadOnly := True;
        CD078DDataSet.ReadOnly := True;
        CD078GDataSet.ReadOnly := True;
        CD078JDataSet.ReadOnly := True;
      end;
    dmCancel:
      begin
        CD078DataSet.AfterScroll := nil;
        try
          CD078DataSet.Cancel;
        finally
          CD078DataSet.AfterScroll := CD078DataSetAfterScroll;
        end;
        CD078ADataSet.CancelUpdates;
        CD078BDataSet.CancelUpdates;
        CD078CDataSet.CancelUpdates;
        CD078DDataSet.CancelUpdates;
        CD078DataSetAfterScroll( CD078DataSet );
        CD078DataSet.ReadOnly := True;
        CD078ADataSet.ReadOnly := True;
        CD078BDataSet.ReadOnly := True;
        CD078CDataSet.ReadOnly := True;
        CD078DDataSet.ReadOnly := True;
        CD078GDataSet.ReadOnly := True;
        CD078JDataSet.ReadOnly := True;
      end;
    dmInsert:
      begin
        CD078DataSet.ReadOnly := False;
        CD078DataSetAfterScroll( CD078DataSet );
      end;
    dmUpdate:
      begin
        CD078DataSet.ReadOnly := False;
        CD078DataSetAfterScroll( CD078DataSet );
      end;
    dmSave:
      begin
        Screen.Cursor := crSQLWait;
        try
          Result := DataValidate( aErrSource );
          if not Result then
          begin
            if ( aErrSource <> EmptyStr ) then
              actUpdate.Execute;
            Exit;
          end;
          DBController.DataConnection.BeginTrans;
          try
            Result := ApplyDataSet;
            DBController.DataConnection.CommitTrans;
          except
            on E: Exception do
            begin
              DBController.DataConnection.RollbackTrans;
              ErrorMsg( Format( '存檔有誤, 訊息:%s。', [E.Message] ) );
            end;
          end;
        finally
          Screen.Cursor := crDefault;
        end;
      end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.EditorStateChange;
begin
  if ( FMode = dmBrowser ) then
  begin
    edtCodeNo.Enabled := False;
    edtDescription.Enabled := False;
    edtDateSt.Enabled := False;
    edtDateEd.Enabled := False;
    edtKindFunction.Enabled := False;
    chkMasterProd.Enabled := False;
    chkStopFlag.Enabled := False;
    chkByHouse.Enabled := False;
    edtRefNo.Enabled := False;
    edtBundleProdNote.Enabled := False;
    edtWorkNote.Enabled := False;
    chkChooseKind.Enabled := False;     //4435 增加頻道挑選模式 By Kin 2009/06/06
    chkPayKind.Enabled := False;        //5683 增加繳付類別  By Kin 2010/08/16
    SyncBPCode.Enabled := False;
    { #6388 增加CANUSETYPE、EPGIMAGEURL、EPGORD 欄位 By Kin 2012/12/18}
    cobCanUseType.Enabled := False;
    edtEPGOrder.Enabled := False;
    edtImgURL.Enabled := False;

  end else
  if ( FMode = dmInsert ) then
  begin
    ClearMasterEditor;
    edtCodeNo.Enabled := True;
    edtDescription.Enabled := True;
    edtDateSt.Enabled := True;
    edtDateEd.Enabled := True;
    edtKindFunction.Enabled := True;
    chkMasterProd.Enabled := True;
    chkStopFlag.Enabled := True;
    chkByHouse.Enabled := True;
    edtRefNo.Enabled := True;
    edtBundleProdNote.Enabled := True;
    edtWorkNote.Enabled := True;
    SyncBPCode.Enabled := True;
    { #6388 增加CANUSETYPE、EPGIMAGEURL、EPGORD 欄位 By Kin 2012/12/18}
    cobCanUseType.Enabled := True;
    edtEPGOrder.Enabled := True;
    edtImgURL.Enabled := True;
    if ( edtCodeNo.CanFocusEx ) then edtCodeNo.SetFocus;
  end else
  if ( FMode = dmUpdate ) then
  begin
    edtCodeNo.Enabled := False;
    edtDescription.Enabled := True;
    edtDateSt.Enabled := True;
    edtDateEd.Enabled := True;
    edtKindFunction.Enabled := True;
    chkMasterProd.Enabled := True;
    chkStopFlag.Enabled := True;
    chkByHouse.Enabled := True;
    edtRefNo.Enabled := True;
    edtBundleProdNote.Enabled := True;
    edtWorkNote.Enabled := True;
    SyncBPCode.Enabled := True;
    { #6388 增加CANUSETYPE、EPGIMAGEURL、EPGORD 欄位 By Kin 2012/12/18}
    cobCanUseType.Enabled := True;
    edtEPGOrder.Enabled := True;
    edtImgURL.Enabled := True;
    {#6388 測試不OK,CanUseType =0 一樣不能選 By Kin 2013/03/07}
    if ( cobCanUseType.ItemIndex = 2 ) or (cobCanUseType.ItemIndex = 0 ) then
    begin
//      edtKindFunction.ItemIndex := 0;
      edtKindFunction.Enabled := False;
    end;
    if ( edtDescription.CanFocusEx ) then edtDescription.SetFocus;
  end else
  begin
    {}
  end;      
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.LoadCD078Data;
begin
  CD078DataSet.Close;
  CD078DataSet.CommandText := Format(
    ' SELECT A.CODENO,           ' +
    '        A.DESCRIPTION,      ' +
    '        A.ONSALESTARTDATE,  ' +
    '        A.ONSALESTOPDATE,   ' +
    '        A.MASTERPROD,       ' +
    '        A.BUNDLE,           ' +
    '        A.BUNDLESAMEMON,    ' +
    '        A.BUNDLEMON,        ' +
    '        A.SAMEPERIOD,       ' +
    '        A.PERIOD,           ' +
    '        A.MERGEPRINT,       ' +
    '        A.PENAL,            ' +
    '        A.PENALTYPE,        ' +
    '        A.PENALSINGLE,      ' +
    '        A.MERGEWORK,        ' +
    '        A.BUNDLEPRODNOTE,   ' +
    '        A.WORKNOTE,         ' +
    '        A.REFNO,            ' +
    '        A.STOPFLAG,         ' +
    '        A.CONDITIONAL,      ' +
    '        A.KINDFUNCTION,     ' +
    '        A.CHOOSEKIND,       ' +          //#4435 增加頻道挑選模式 By Kin 2009/06/06
    '        A.CANPAYNOW,        ' +
    '        A.SYNCBPCODE,       ' +
    '        A.CANUSETYPE,       ' +
    '        A.EPGIMAGEURL,      ' +
    '        A.EPGORD,           ' +
    '        A.BYHOUSE           ' +
    '   FROM %s.CD078 A          ' +
    '  WHERE A.CODENO = ''%s''   ', [DBController.LoginInfo.DbAccount, FKeyValue] );
  CD078DataSet.AfterScroll := nil;
  try
    Application.ProcessMessages;
    CD078DataSet.Open;
    Application.ProcessMessages;
  finally
    CD078DataSet.AfterScroll := CD078DataSetAfterScroll;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.LoadCD078A_Data;

  { --------------------------------------------- }

  procedure AcquireDefaultDeposit(ACItemCode, AServiceType: String;
    var ACode, AName: String);
  begin
    ACode := EmptyStr;
    AName := EmptyStr;
    FReader.Close;
    FReader.SQL.Text := Format(
      ' SELECT A.DEPOSITCODE,                ' +
      '        B.DESCRIPTION                 ' +
      '   FROM %s.CD078A2 A, %s.CD019 B      ' +
      '  WHERE A.DEPOSITCODE = B.CODENO      ' +
      '    AND A.BPCODE = ''%s''             ' +
      '    AND A.CITEMCODE = ''%s''          ' +
      '    AND A.SERVICETYPE = ''%s''        ' +
      '    AND A.DEPOSITDEFAULT = 1          ',
      [DBController.LoginInfo.DbAccount,
       DBController.LoginInfo.DbAccount,
       FKeyValue, ACItemCode, AServiceType] );
    FReader.Open;
    ACode := FReader.FieldByName( 'DEPOSITCODE' ).AsString;
    AName := FReader.FieldByName( 'DESCRIPTION' ).AsString;
    FReader.Close;
  end;

  { --------------------------------------------- }

var
  aDepositCode, aDepositName: String;
begin
  { #4435 增加 minCount與MaxCount欄位 By Kin 2009/03/24 }
  { #5318 增加SalePointName與SalePointCode欄位 By Kin 2009/10/06}
  { #5257 增加特定優惠截止日 By Kin 2010/05/05}
  { #5913 增加破月期數(收費單) By Kin 2011/04/14}
  { # 增加ContractType與AuthTunerCount 欄位 By Kin 2015/04/21}
  CD078ADataSet.Close;
  CD078ADataSet.CommandText := Format(
    ' SELECT ' +
    '        A.BPCODE,                       ' +
    '        A.CITEMCODE,                    ' +
    '        A.CITEMNAME,                    ' +
    '        A.SERVICETYPE,                  ' +
    '        A.FACIITEM,                     ' +
    '        A.BUNDLE,                       ' +
    '        A.BUNDLEMON,                    ' +
    '        A.PENALTYPE,                    ' +
    '        A.EXPIRETYPE,                   ' +
    '        A.EXPIRECUT,                    ' +
    '        A.DEPOSITATTR,                  ' +
    '        A.DEPOSITCODE,                  ' +
    '        A.DEPOSITNAME,                  ' +
    '        A.MONTHAMT,                     ' +
    '        A.DAYAMT,                       ' +
    '        A.PFLAG1,                       ' +
    '        A.PFLAG2,                       ' +
    '        A.PFLAGDATE,                    ' +
    '        A.TRUNCAMT,                     ' +
    '        A.PFLAGCODE,                    ' +
    '        A.PFLAGNAME,                    ' +
    '        A.MON1,                         ' +
    '        A.PERIOD1,                      ' +
    '        A.RATETYPE1,                    ' +
    '        A.DISCOUNTAMT1,                 ' +
    '        A.MONTHAMT1,                    ' +
    '        A.DAYAMT1,                      ' +
    '        A.MON2,                         ' +
    '        A.PERIOD2,                      ' +
    '        A.RATETYPE2,                    ' +
    '        A.DISCOUNTAMT2,                 ' +
    '        A.MONTHAMT2,                    ' +
    '        A.DAYAMT2,                      ' +
    '        A.MON3,                         ' +
    '        A.PERIOD3,                      ' +
    '        A.RATETYPE3,                    ' +
    '        A.DISCOUNTAMT3,                 ' +
    '        A.MONTHAMT3,                    ' +
    '        A.DAYAMT3,                      ' +
    '        A.PUNISH3,                      ' +
    '        A.MON4,                         ' +
    '        A.PERIOD4,                      ' +
    '        A.RATETYPE4,                    ' +
    '        A.DISCOUNTAMT4,                 ' +
    '        A.MONTHAMT4,                    ' +
    '        A.DAYAMT4,                      ' +
    '        A.MON5,                         ' +
    '        A.PERIOD5,                      ' +
    '        A.RATETYPE5,                    ' +
    '        A.DISCOUNTAMT5,                 ' +
    '        A.MONTHAMT5,                    ' +
    '        A.DAYAMT5,                      ' +
    '        A.MON6,                         ' +
    '        A.PERIOD6,                      ' +
    '        A.RATETYPE6,                    ' +
    '        A.DISCOUNTAMT6,                 ' +
    '        A.MONTHAMT6,                    ' +
    '        A.DAYAMT6,                      ' +
    '        A.MON7,                         ' +
    '        A.PERIOD7,                      ' +
    '        A.RATETYPE7,                    ' +
    '        A.DISCOUNTAMT7,                 ' +
    '        A.MONTHAMT7,                    ' +
    '        A.DAYAMT7,                      ' +
    '        A.MON8,                         ' +
    '        A.PERIOD8,                      ' +
    '        A.RATETYPE8,                    ' +
    '        A.DISCOUNTAMT8,                 ' +
    '        A.MONTHAMT8,                    ' +
    '        A.DAYAMT8,                      ' +
    '        A.MON9,                         ' +
    '        A.PERIOD9,                      ' +
    '        A.RATETYPE9,                    ' +
    '        A.DISCOUNTAMT9,                 ' +
    '        A.MONTHAMT9,                    ' +
    '        A.DAYAMT9,                      ' +
    '        A.MON10,                        ' +
    '        A.PERIOD10,                     ' +
    '        A.RATETYPE10,                   ' +
    '        A.DISCOUNTAMT10,                ' +
    '        A.MONTHAMT10,                   ' +
    '        A.DAYAMT10,                     ' +
    '        A.MON11,                        ' +
    '        A.PERIOD11,                     ' +
    '        A.RATETYPE11,                   ' +
    '        A.DISCOUNTAMT11,                ' +
    '        A.MONTHAMT11,                   ' +
    '        A.DAYAMT11,                     ' +
    '        A.MON12,                        ' +
    '        A.PERIOD12,                     ' +
    '        A.RATETYPE12,                   ' +
    '        A.DISCOUNTAMT12,                ' +
    '        A.MONTHAMT12,                   ' +
    '        A.DAYAMT12,                     ' +
    '        A.PENAL,                        ' +
    '        A.PENALCODE,                    ' +
    '        A.PENALNAME,                    ' +
    '        A.DEPOSITCODESTR,               ' +
    '        A.DISCOUNTCALC,                 ' +
    '        A.INSTCODESTR,                  ' +
    '        A.COMMENT1,                     ' +
    '        A.COMMENT2,                     ' +
    '        A.COMMENT3,                     ' +
    '        A.COMMENT4,                     ' +
    '        A.COMMENT5,                     ' +
    '        A.COMMENT6,                     ' +
    '        A.COMMENT7,                     ' +
    '        A.COMMENT8,                     ' +
    '        A.COMMENT9,                     ' +
    '        A.COMMENT10,                    ' +
    '        A.COMMENT11,                    ' +
    '        A.COMMENT12,                    ' +
    '        A.MCITEMCODE,                   ' +
    '        A.CMBAUDNO,                     ' +
    '        A.CMBAUDRATE,                   ' +
    '        B.SIGN,                         ' +
    '        B.REFNO,                        ' +
    '        A.MinCount,                     ' +
    '        A.MaxCount,                     ' +
    '        A.SALEPOINTCODE,                ' +
    '        A.SALEPOINTNAME,                ' +
    '        A.BPSTOPDATE,                   ' +
    '        A.STOPFLAG,                     ' +
    '        A.LONGPAYFLAG,                  ' +
    '        A.AUTHTUNERCOUNT,               ' +
    '        A.CONTRACTTYPE                  ' +
    '   FROM %s.CD078A A, %s.CD019 B         ' +
    '  WHERE A.CITEMCODE = B.CODENO(+)       ' +
    '    AND A.BPCODE = ''%s''               ' +
    '  ORDER BY A.CITEMCODE                  ',
    [DBController.LoginInfo.DbAccount,
     DBController.LoginInfo.DbAccount,
     FKeyValue] );
  CD078ADataSet.Open;
  CD078ADataSet.First;
  { 將 CD078A2, 預設的付款種類取出, 填到 CD078A }
  while not CD078ADataSet.Eof do
  begin
    AcquireDefaultDeposit( CD078ADataSet.FieldByName( 'CITEMCODE' ).AsString,
      CD078ADataSet.FieldByName( 'SERVICETYPE' ).AsString, aDepositCode, aDepositName );
    CD078ADataSet.Edit;
    CD078ADataSet.FieldByName( 'DEPOSITCODE' ).AsString := aDepositCode;
    CD078ADataSet.FieldByName( 'DEPOSITNAME' ).AsString := aDepositName;
    CD078ADataSet.Post;
    CD078ADataSet.Next;
  end;
  CD078ADataSet.First;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.LoadCD078A1_Data;
begin
  {#6865 增加PERIOD By Kin 2014/09/10}
  CD078A1DataSet.Close;
  CD078A1DataSet.CommandText := Format(
    ' SELECT ' +
    '        A.BPCODE,                       ' +
    '        A.CITEMCODE,                    ' +
    '        A.SERVICETYPE,                  ' +
    '        A.FACIITEM,                     ' +
    '        A.STEPNO,                       ' +
    '        A.DESCRIPTION,                  ' +
    '        A.MASTERSALE,                   ' +
    '        A.MON1,                         ' +
    '        A.PERIOD1,                      ' +
    '        A.RATETYPE1,                    ' +
    '        A.DISCOUNTAMT1,                 ' +
    '        A.MONTHAMT1,                    ' +
    '        A.DAYAMT1,                      ' +
    '        A.MON2,                         ' +
    '        A.PERIOD2,                      ' +
    '        A.RATETYPE2,                    ' +
    '        A.DISCOUNTAMT2,                 ' +
    '        A.MONTHAMT2,                    ' +
    '        A.DAYAMT2,                      ' +
    '        A.MON3,                         ' +
    '        A.PERIOD3,                      ' +
    '        A.RATETYPE3,                    ' +
    '        A.DISCOUNTAMT3,                 ' +
    '        A.MONTHAMT3,                    ' +
    '        A.DAYAMT3,                      ' +
    '        A.MON4,                         ' +
    '        A.PERIOD4,                      ' +
    '        A.RATETYPE4,                    ' +
    '        A.DISCOUNTAMT4,                 ' +
    '        A.MONTHAMT4,                    ' +
    '        A.DAYAMT4,                      ' +
    '        A.MON5,                         ' +
    '        A.PERIOD5,                      ' +
    '        A.RATETYPE5,                    ' +
    '        A.DISCOUNTAMT5,                 ' +
    '        A.MONTHAMT5,                    ' +
    '        A.DAYAMT5,                      ' +
    '        A.MON6,                         ' +
    '        A.PERIOD6,                      ' +
    '        A.RATETYPE6,                    ' +
    '        A.DISCOUNTAMT6,                 ' +
    '        A.MONTHAMT6,                    ' +
    '        A.DAYAMT6,                      ' +
    '        A.MON7,                         ' +
    '        A.PERIOD7,                      ' +
    '        A.RATETYPE7,                    ' +
    '        A.DISCOUNTAMT7,                 ' +
    '        A.MONTHAMT7,                    ' +
    '        A.DAYAMT7,                      ' +
    '        A.MON8,                         ' +
    '        A.PERIOD8,                      ' +
    '        A.RATETYPE8,                    ' +
    '        A.DISCOUNTAMT8,                 ' +
    '        A.MONTHAMT8,                    ' +
    '        A.DAYAMT8,                      ' +
    '        A.MON9,                         ' +
    '        A.PERIOD9,                      ' +
    '        A.RATETYPE9,                    ' +
    '        A.DISCOUNTAMT9,                 ' +
    '        A.MONTHAMT9,                    ' +
    '        A.DAYAMT9,                      ' +
    '        A.MON10,                        ' +
    '        A.PERIOD10,                     ' +
    '        A.RATETYPE10,                   ' +
    '        A.DISCOUNTAMT10,                ' +
    '        A.MONTHAMT10,                   ' +
    '        A.DAYAMT10,                     ' +
    '        A.MON11,                        ' +
    '        A.PERIOD11,                     ' +
    '        A.RATETYPE11,                   ' +
    '        A.DISCOUNTAMT11,                ' +
    '        A.MONTHAMT11,                   ' +
    '        A.DAYAMT11,                     ' +
    '        A.MON12,                        ' +
    '        A.PERIOD12,                     ' +
    '        A.RATETYPE12,                   ' +
    '        A.DISCOUNTAMT12,                ' +
    '        A.MONTHAMT12,                   ' +
    '        A.DAYAMT12,                     ' +
    '        A.COMMENT1,                     ' +
    '        A.COMMENT2,                     ' +
    '        A.COMMENT3,                     ' +
    '        A.COMMENT4,                     ' +
    '        A.COMMENT5,                     ' +
    '        A.COMMENT6,                     ' +
    '        A.COMMENT7,                     ' +
    '        A.COMMENT8,                     ' +
    '        A.COMMENT9,                     ' +
    '        A.COMMENT10,                    ' +
    '        A.COMMENT11,                    ' +
    '        A.COMMENT12,                    ' +
    '        A.LINKKEY,                      ' +
    '        A.DISCOUNTRATE1,                ' +
    '        A.DISCOUNTRATE2,                ' +
    '        A.DISCOUNTRATE3,                ' +
    '        A.DISCOUNTRATE4,                ' +
    '        A.DISCOUNTRATE5,                ' +
    '        A.DISCOUNTRATE6,                ' +
    '        A.DISCOUNTRATE7,                ' +
    '        A.DISCOUNTRATE8,                ' +
    '        A.DISCOUNTRATE9,                ' +
    '        A.DISCOUNTRATE10,               ' +
    '        A.DISCOUNTRATE11,               ' +
    '        A.DISCOUNTRATE12,               ' +
    '        B.DESCRIPTION AS LINKKEYNAME,   ' +
    '        A.PERIOD                        ' +
    '   FROM %s.CD078A1 A, %s.CD078A1 B      ' +
    '  WHERE A.LINKKEY = B.STEPNO(+)         ' +
    '    AND A.BPCODE = ''%s''               ' +
    '  ORDER BY A.BPCODE, A.CITEMCODE,       ' +
    '           A.SERVICETYPE, A.FACIITEM,   ' +
    '           A.STEPNO                     ',
    [DBController.LoginInfo.DbAccount,
     DBController.LoginInfo.DbAccount, FKeyValue] );
  CD078A1DataSet.Open;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.LoadCD078A2_Data;
begin
  CD078A2DataSet.Close;
  CD078A2DataSet.CommandText := Format(
    ' SELECT ' + 
    '        A.BPCODE,                       ' +
    '        A.CITEMCODE,                    ' +
    '        A.SERVICETYPE,                  ' +
    '        A.DEPOSITCODE,                  ' +
    '        A.PTCODE,                       ' +
    '        A.PTNAME,                       ' +
    '        A.DEPOSITAMT,                   ' +
    '        A.DEPOSITDEFAULT,               ' +
    '        B.DESCRIPTION AS DEPOSITNAME    ' + 
    '   FROM %s.CD078A2 A, %s.CD019 B        ' +
    '  WHERE A.DEPOSITCODE = B.CODENO(+)     ' +
    '    AND A.BPCODE = ''%s''               ' +
    '  ORDER BY A.BPCODE,                    ' +
    '           A.CITEMCODE,                 ' +
    '           A.SERVICETYPE,               ' +
    '           A.DEPOSITCODE,               ' +
    '           A.PTCODE                     ',
    [DBController.LoginInfo.DbAccount,
     DBController.LoginInfo.DbAccount,
     FKeyValue] );
  CD078A2DataSet.Open;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.LoadCD078A3_Data;
begin
  CD078A3DataSet.Close;
  CD078A3DataSet.CommandText := Format(
    ' SELECT ' + 
    '        A.BPCODE,                       ' +
    '        A.CITEMCODE,                    ' +
    '        A.SERVICETYPE,                  ' +
    '        A.PENALCODE,                    ' +
    '        A.ITEM,                         ' +
    '        A.PENALMON,                     ' +
    '        A.PENALAMT                      ' +
    '   FROM %s.CD078A3 A                    ' +
    '  WHERE A.BPCODE = ''%s''               ' +
    '  ORDER BY A.BPCODE, A.CITEMCODE,       ' +
    '           A.SERVICETYPE, A.ITEM        ',
    [DBController.LoginInfo.DbAccount, FKeyValue] );
  CD078A3DataSet.Open;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.LoadCD078A4_Data;
begin
  CD078A4DataSet.Close;
  CD078A4DataSet.CommandText := Format(
    ' SELECT ' +
    '        A.BPCODE,                       ' +
    '        A.CITEMCODE,                    ' +
    '        A.SERVICETYPE,                  ' +
    '        A.PENALCODE,                    ' +
    '        A.ITEM,                         ' +
    '        A.MONTHSTART,                   ' +
    '        A.MONTHSTOP,                    ' +
    '        A.MONTHAMT,                     ' +
    '        A.DECREASEAMT                   ' +
    '   FROM %s.CD078A4 A                    ' +
    '  WHERE A.BPCODE = ''%s''               ' +
    '  ORDER BY A.BPCODE, A.CITEMCODE,       ' +
    '           A.SERVICETYPE, A.ITEM        ',
    [DBController.LoginInfo.DbAccount, FKeyValue] );
  CD078A4DataSet.Open;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.LoadCD078B_Data;
begin
  CD078BDataSet.Close;
  CD078BDataSet.CommandText := Format(
    ' SELECT ' +
    '        A.BPCODE,                         ' +
    '        A.NEWBPCODE,                      ' +
    '        A.INSTCODE,                       ' +
    '        A.INSTNAME,                       ' +
    '        A.SERVICETYPE,                    ' +
    '        A.GROUPNO,                        ' +
    '        A.GROUPNAME,                      ' +
    '        A.WORKUNIT,                       ' +
    '        A.REFNO,                          ' +
    '        A.INTERDEPEND,                    ' +
    '        A.PRCODE,                         ' +
    '        A.PRNAME                          ' +
    '   FROM %s.CD078B A                       ' +
    '  WHERE A.BPCODE = ''%s''                 ' +
    '    AND A.NEWBPCODE IS NULL               ' +
    '  ORDER BY A.SERVICETYPE, A.INSTCODE      ',
    [DBController.LoginInfo.DbAccount, FKeyValue] );
  CD078BDataSet.Open;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.LoadCD078B1_Data;
begin
  CD078B1DataSet.Close;
  CD078B1DataSet.CommandText := Format(
    ' SELECT A.BPCODE,                       ' +
    '        A.INSTCODE,                     ' +
    '        A.SERVICETYPE,                  ' +
    '        A.INSTCODE2,                    ' +
    '        A.INSTNAME2,                    ' +
    '        A.PRCODE2,                      ' +
    '        A.PRNAME2                       ' +
    '   FROM %s.CD078B1 A                    ' +
    '  WHERE A.BPCODE = ''%s''               ' +
    '  ORDER BY                              ' +
    '    A.BPCODE, A.INSTCODE, A.INSTCODE2   ',
    [DBController.LoginInfo.DbAccount, FKeyValue] );
  CD078B1DataSet.Open;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.LoadCD078C_Data;
begin                                 
  CD078CDataSet.Close;
  CD078CDataSet.CommandText := Format(
    ' SELECT ' +
    '        A.BPCODE,                        ' +
    '        A.NEWBPCODE,                     ' +
    '        A.CITEMCODE,                     ' +
    '        A.CITEMNAME,                     ' +
    '        A.SERVICETYPE,                   ' +
    '        A.FACIITEM,                      ' +
    '        A.AMOUNT,                        ' +
    '        A.DISCOUNTAMT,                   ' +
    '        A.PUNISH,                        ' +
    '        A.PENALTYPE,                     ' +
    '        A.REFCITEMCODE,                  ' +
    '        A.REFCITEMNAME,                  ' +
    '        A.INSTCODESTR                    ' +
    '   FROM %s.CD078C A                      ' +
    '  WHERE A.BPCODE = ''%s''                ' +
    '    AND A.NEWBPCODE IS NULL              ' +
    '  ORDER BY A.SERVICETYPE, A.CITEMCODE    ',
    [DBController.LoginInfo.DbAccount, FKeyValue] );
  CD078CDataSet.Open;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.LoadCD078D_Data;
begin
  {#5859 增加CBPCHANGEFIXIP與CBPCHANGEDYNIP By Kin 2011/04/27}
  CD078DDataSet.Close;
  CD078DDataSet.CommandText := Format(
    ' SELECT ' +
    '        A.BPCODE,                       ' +
    '        A.NEWBPCODE,                    ' +
    '        A.FACIITEM,                     ' +
    '        A.FACITYPE,                     ' +
    '        A.FACICODE,                     ' +
    '        A.FACINAME,                     ' +
    '        A.SERVICETYPE,                  ' +
    '        A.MODELCODE,                    ' +
    '        A.MODELNAME,                    ' +
    '        A.BUYCODE,                      ' +
    '        A.BUYNAME,                      ' +
    '        A.CMBAUDNO,                     ' +
    '        A.CMBAUDRATE,                   ' +
    '        A.FIXIPCOUNT,                   ' +
    '        A.DYNIPCOUNT,                   ' +
    '        A.INSTCODESTR,                  ' +
    '        A.CBPCHANGEFIXIP,               ' +
    '        A.CBPCHANGEDYNIP,               ' +
    '        B.REFNO                         ' +
    '   FROM %s.CD078D A,                    ' +
    '        %s.CD022 B                      ' +
    '  WHERE A.FACICODE = B.CODENO(+)        ' +
    '    AND A.BPCODE = ''%s''               ' +
    '    AND A.NEWBPCODE IS NULL             ' +
    '  ORDER BY A.FACIITEM                   ',
    [DBController.LoginInfo.DbAccount,
     DBController.LoginInfo.DbAccount, FKeyValue] );
  CD078DDataSet.Open;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.CD078DataSetAfterScroll(DataSet: TDataSet);
begin
  Screen.Cursor := crSQLWait;
  try
    FKeyValue := CD078DataSet.FieldByName( 'CODENO' ).AsString;
    ClearMasterEditor;
    MasterDataToEditor;
    LoadCD078A_Data;
    LoadCD078A1_Data;
    CloneCD078A1Data;
    LoadCD078A2_Data;
    CloneCD078A2Data;
    {}
    LoadCD078A3_Data;
    CloneCD078A3Data;
    {}
    LoadCD078A4_Data;
    CloneCD078A4Data;
    {}
    LoadCD078B_Data;
    LoadCD078B1_Data;
    CloneCD078B1Data;
    {}
    LoadCD078C_Data;
    LoadCD078D_Data;
    {#5318 增加點數 設定 By Kin 2009/10/06}
    LoadCD078G_Data;
    LoadCD078G1_Data;
    CloneCD078G1Data;
    {#6035 增加PVOD 設定 By Kin 2011/06/08}
    LoadCD078I_Data;
    CloneCD078IData;
    LoadSO562_Data;
    LoadCD078J_Data;
    {}
    {#6881 增加CD078K By Kin 2014/10/14}
    LoadCD078K_Data;
    {#6931 增加CD078L By Kin 2014/11/13}
    LoadCD078L_Data;
    {
    if CD078DataSet.FieldByName( 'CanUseType' ).AsInteger = 2 then
    begin
      if Not FCanUseEPG then
      begin
        btnSave.Enabled := False;
        btnInsert.Enabled := False;
        btnUpdate.Enabled := False;
        btnDelete.Enabled := False;
      end;
    end;
    }
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.actSaveExecute(Sender: TObject);
begin
  if btnSave.Enabled then
    DMLModeChange( dmSave );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.actInsertExecute(Sender: TObject);
begin
  if btnInsert.Enabled then
  begin
    Screen.Cursor := crSQLWait;
    try
      ShowDetailForm( dmInsert );
    finally
      Screen.Cursor := crDefault;
    end;
  end;

  

end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.actDeleteExecute(Sender: TObject);
var
  aDataSet: TClientDataSet;
begin
  if not btnDelete.Enabled then
    Exit;
  case MainPageControl.ActivePageIndex of
    0: aDataSet := CD078BDataSet;
    1: aDataSet := CD078DDataSet;
    2: aDataSet := CD078ADataSet;
    3: aDataSet := CD078CDataSet;
    4: aDataSet := CD078GDataSet;  {#5318 增加點數設定 By Kin 2009/10/06}
    5: aDataSet := SO562DataSet;
    6: aDataSet := CD078JDataSet;
  else
    aDataSet := nil;
  end;
  
  if Assigned( aDataSet ) and ( not aDataSet.IsEmpty ) then
  begin
    if ( aDataSet = CD078DDataSet ) then
    begin
      if DoDeleteCheck then
      begin
        WarningMsg( '本筆【設備參數】資料已被【優惠參數】及【費用參數】使用，無法刪除。' );
        Exit;
      end;
    end;
    if ConfirmMsg( '確認刪除此筆資料?' ) then
    begin
      if ( aDataSet = CD078ADataSet ) then
      begin
        DeleteCD078A1Data;
        DeleteCD078A2Data;
        DeleteCD078A3Data;
        DeleteCD078A4Data;

        {#6035 刪除收費項目要同步SO562.Type=0的資料 By Kin 2011/06/10}
        DeleteSO562Data;
      end else
      if ( aDataSet = CD078BDataSet ) then
      begin
        { 刪掉所有參考到該 裝機類別 的資料 }
        RemoveInstCodeStr( aDataSet.FieldByName( 'InstCode' ).AsString );
        DeleteCD078CIfInstCodeIsEmpty;
        { !!! 在 CD078B1 裏面, 也要執行 RemoveInstCodeStr !!! }
        DeleteCD078B1Data;
      end;
      if (aDataSet = CD078GDataSet ) then
      begin
        DeleteCD078G1Data;
      end;
      aDataSet.Delete;
    end;
    ButtonStateChange;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.actUpdateExecute(Sender: TObject);
var
  aDataSet: TClientDataSet;
begin
  if not btnUpdate.Enabled then
    Exit;
  Screen.Cursor := crSQLWait;
  
  try
    case MainPageControl.ActivePageIndex of
      0: aDataSet := CD078BDataSet;
      1: aDataSet := CD078DDataSet;
      2: aDataSet := CD078ADataSet;
      3: aDataSet := CD078CDataSet;
      4: aDataSet := CD078GDataSet;   {#5318 增加點數設定 By Kin 2009/10/06}
      5: aDataSet := SO562DataSet;    {#6035 增加PVOD設定 By Kin 2011/06/14}
      6: aDataSet := CD078JDataSet;
    else
      aDataSet := nil;
    end;
    if not Assigned( aDataSet ) or ( aDataSet.IsEmpty ) then
      Exit;
    ShowDetailForm( dmUpdate );
  finally
    Screen.Cursor := crDefault;
  end;


end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.actCancelExecute(Sender: TObject);
begin
  if not ( ConfirmMsg( '新增或修改尚未存檔,確認是否取消?' ) ) then Exit;
  Self.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.actCloseExecute(Sender: TObject);
begin
  if ( FMode in [dmInsert,dmUpdate] ) then DMLModeChange( dmCancel );
  Self.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.ClearMasterEditor;
begin
  edtCodeNo.Text := EmptyStr;
  edtDescription.Text := EmptyStr;
  edtDateSt.Text := EmptyStr;
  edtDateEd.Text := EmptyStr;
  chkMasterProd.Checked := False;
  chkStopFlag.Checked := False;
  chkByHouse.Checked := False;
  edtRefNo.Text := EmptyStr;
  edtBundleProdNote.Text := EmptyStr;
  edtWorkNote.Text := EmptyStr;
  edtConditional.Text := EmptyStr;
  edtKindFunction.ItemIndex := 0;
  chkChooseKind.Checked := False;     //#4435 增加頻道挑選模式 By Kin 2009/06/06
  chkPayKind.Checked := False;        //#5683 增加挑選模式 By Kin 2010/08/16
  SyncBPCode.Clear;                   //#6267 增加SYNCBPCODE 欄位 By Kin 2012/07/19
  { #6388 增加CANUSETYPE、EPGIMAGEURL、EPGORD 欄位 By Kin 2012/12/18}
  edtEPGOrder.Clear;
  cobCanUseType.ItemIndex := 0;
  edtImgURL.Clear;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.MasterDataToEditor;
begin
  edtCodeNo.Text := CD078DataSet.FieldByName( 'CODENO' ).AsString;
  edtDescription.Text := CD078DataSet.FieldByName( 'DESCRIPTION' ).AsString;
  if not VarIsNull( CD078DataSet.FieldByName( 'ONSALESTARTDATE' ).Value ) then
    edtDateSt.Text := FormatDateTime( 'yyyy/mm/dd',
      CD078DataSet.FieldByName( 'ONSALESTARTDATE' ).AsDateTime );
  if not VarIsNull( CD078DataSet.FieldByName( 'ONSALESTOPDATE' ).Value ) then
    edtDateEd.Text := FormatDateTime( 'yyyy/mm/dd',
      CD078DataSet.FieldByName( 'ONSALESTOPDATE' ).AsDateTime );
  chkMasterProd.Checked := ( CD078DataSet.FieldByName( 'MASTERPROD' ).AsString = '1' );
  chkStopFlag.Checked := ( CD078DataSet.FieldByName( 'STOPFLAG' ).AsString = '1' );
  chkByHouse.Checked := ( CD078DataSet.FieldByName( 'ByHouse' ).AsString = '1' );
  edtRefNo.Text := CD078DataSet.FieldByName( 'REFNO' ).AsString;
  edtBundleProdNote.Text := CD078DataSet.FieldByName( 'BUNDLEPRODNOTE' ).AsString;
  edtWorkNote.Text := CD078DataSet.FieldByName( 'WORKNOTE' ).AsString;
  edtConditional.Text := CD078DataSet.FieldByName( 'CONDITIONAL' ).AsString;
  edtKindFunction.ItemIndex := 0;
  if ( CD078DataSet.FieldByName( 'KINDFUNCTION' ).AsString = '1' ) then
    edtKindFunction.ItemIndex := 1;
  { #4435 增加頻道挑選模式 By Kin 2009/06/06 }
  if (CD078DataSet.FieldByName( 'CHOOSEKIND' ).AsString = '1' ) then
    chkChooseKind.Checked := True;
  {#5683 增加繳付類別 By Kin 2010/08/16}
  if (CD078DataSet.FieldByName( 'CANPAYNOW' ).AsString = '1') then
    chkPayKind.Checked := True;
  if CD078DataSet.FieldByName( 'SYNCBPCODE' ).AsString <> '' then
    SyncBPCode.Value := CD078DataSet.FieldByName( 'SYNCBPCODE' ).AsString;
  { #6388 增加CANUSETYPE、EPGIMAGEURL、EPGORD 欄位 By Kin 2012/12/18}
  cobCanUseType.ItemIndex := CD078DataSet.FieldByName( 'CANUSETYPE' ).AsInteger;
  edtImgURL.Text := CD078DataSet.FieldByName( 'EPGIMAGEURL' ).AsString;
  edtEPGOrder.Text := CD078DataSet.FieldByName( 'EPGORD' ).AsString;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.MainPageControlChange(Sender: TObject);
begin
  ButtonStateChange;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.MainPageControlDrawTab(AControl: TcxCustomTabControl;
  ATab: TcxTab; var DefaultDraw: Boolean);
var
  aRect: TRect;
begin
  if ATab.IsMainTab then
  begin
    aRect := ATab.VisibleRect;
    AControl.Canvas.FrameRect( aRect, clBlue );
    InflateRect( aRect, -1, -1 );
    AControl.Canvas.FrameRect( aRect, clBlue );
    InflateRect( aRect, -1, -1 );
  end;
  AControl.Canvas.FillRect( aRect );
  AControl.Canvas.Font.Color := clDefault;
  AControl.Canvas.Font.Style := ( AControl.Canvas.Font.Style - [fsBold] );
  if ATab.IsMainTab then
  begin
    AControl.Canvas.Font.Color := clBlue;
    AControl.Canvas.Font.Style := ( AControl.Canvas.Font.Style + [fsBold] );
  end;
  AControl.Canvas.DrawText( ATab.Caption, aRect, cxAlignHCenter );
  DefaultDraw := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.ShowDetailForm(aMode: TDML);
begin
  case MainPageControl.ActivePageIndex of
    0: ShowCD078B_From( aMode );
    1: ShowCD078D_Form( aMode );
    2: ShowCD078A_From( aMode );
    3: ShowCD078C_From( aMode );
    4: ShowCD078E_Form( aMode );
    5: ShowCD078I_Form( aMode );
    6: ShowCD078J_Form( aMode ); 


  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.ShowCD078A_From(aMode: TDML);
var
  aMsg, aCitemCodes: String;
  aControl: TcxCustomTextEdit;
  aOldCitemCode : String;
begin
  aOldCitemCode := EmptyStr;
  if ( aMode in [dmInsert, dmUpdate] ) then
  begin

    aMsg := EmptyStr;
    aControl := nil;
    FKeyValue := edtCodeNo.Text;
    if ( FKeyValue = EmptyStr ) then
    begin
      WarningMsg( '請先輸入優惠組合代碼。' );
      if edtCodeNo.CanFocusEx then edtCodeNo.SetFocus;
      Exit;
    end;
    if ( aMsg <> EmptyStr ) then
    begin
      WarningMsg( aMsg );
      if Assigned( aControl ) then
        if aControl.CanFocusEx then
          aControl.SetFocus;
      Exit;
    end;
  end;
  {}
  CloneCodeData;
  CloneInstCodeData;
  if ( InstCodeDataSet.IsEmpty ) then
  begin
    WarningMsg( '請先設定【派工參數】。' );
    MainPageControl.ActivePageIndex := 0;
    Exit;
  end;
  {}
  RestoreCD078A1Data;
  FilterCD078A1Data( True, aMode );
  {}
  RestoreCD078A2Data;
  FilterCD078A2Data( True, aMode );
  {}
  RestoreCD078A3Data;
  FilterCD078A3Data( True, aMode );
  {}
  RestoreCD078A4Data;
  FilterCD078A4Data( True, aMode );
  {#6035 先過濾SO562包月的收費項目 By Kin 2011/06/10}
  FilterSO562Data( True, aMode);
  {}
  aCitemCodes := GetCD078ACItemValue;
  LoadMCItetmCodeData;
  RestoreLinkKeyData;

  {#6035 記錄原本的CitemCode，存檔後再根據CitemCode回存 By Kin 2011/06/10 }

  if ( aMode = dmUpdate ) then
    aOldCitemCode := CD078ADataSet.FieldByName( 'CITEMCODE' ).AsString;
  {}
  try
    frmCD078B3 := TfrmCD078B3.Create( aMode,  FKeyValue, CD078ADataSet );
    try
      frmCD078B3.Caption := MainPageControl.ActivePage.Caption + ' [' + frmCD078B3.Caption + ']';
      frmCD078B3.AlreadySetCItemValue := aCitemCodes;
      frmCD078B3.FaciItemDataSet := OtherDataSet;
      frmCD078B3.InstCodeDataSet := InstCodeDataSet;
      frmCD078B3.MCItemDataSet := MCItemDataSet;
      frmCD078B3.LinkKeyDataSet := LinkKeyDataSet;
      frmCD078B3.CD078A1DataSet := CD078A1DataSet;
      frmCD078B3.CD078A2DataSet := CD078A2DataSet;
      frmCD078B3.CD078A3DataSet := CD078A3DataSet;
      frmCD078B3.CD078A4DataSet := CD078A4DataSet;
      {#6841 增加判斷長繳別註記是否可以修改 By Kin 2014/08/05}
      frmCD078B3.CanEditLongPayFlag := True;
      if Now >= StrToDateTime( edtDateSt.Text ) then
        frmCD078B3.CanEditLongPayFlag := False;
      frmCD078B3.ShowModal;
    finally
      frmCD078B3.Free;
    end;
  finally
    {#6035 儲存資料，需要判斷SO562是否有符合的收費項目 By Kin 2011/06/14}
    ChangeSO562Data( aOldCitemCode,aMode );
    FilterCD078A1Data( False, aMode );
    FilterCD078A2Data( False, aMode );
    FilterCD078A3Data( False, aMode );
    FilterCD078A4Data( False, aMode );
    FilterSO562Data( False, aMode);
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.ShowCD078B_From(aMode: TDML);
begin
  FKeyValue := edtCodeNo.Text;
  if ( FKeyValue = EmptyStr ) then
  begin
    WarningMsg( '請先輸入優惠組合代碼。' );
    if edtCodeNo.CanFocusEx then edtCodeNo.SetFocus;
    Exit;
  end;
  RestoreCD078B1Data;
  FilterCD078B1Data( True, aMode );
  try
    frmCD078B4 := TfrmCD078B4.Create( aMode, FKeyValue, EmptyStr, CD078BDataSet );
    try
      frmCD078B4.Caption := MainPageControl.ActivePage.Caption + ' [' + frmCD078B4.Caption + ']';
      frmCD078B4.FaciItemDataSet := OtherDataSet;
      frmCD078B4.CD078DDataSet := CD078DDataSet;
      frmCD078B4.CD078ADataSet := CD078ADataSet;
      frmCD078B4.CD078CDataSet := CD078CDataSet;
      frmCD078B4.CD078B1DataSet := CD078B1DataSet;
      frmCD078B4.CD078JDataSet := CD078JDataSet;
      frmCD078B4.ShowModal;
    finally
      frmCD078B4.Free;
    end;
  finally
    FilterCD078B1Data( False, aMode );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.ShowCD078C_From(aMode: TDML);
var
  aService: String;
begin
  CloneCodeData;
  FKeyValue := edtCodeNo.Text;
  if ( FKeyValue = EmptyStr ) then
  begin
    WarningMsg( '請先輸入【優惠組合代碼】。' );
    if edtCodeNo.CanFocusEx then edtCodeNo.SetFocus;
    Exit;
  end;
  aService := ServiceTypeIntoStr;
  if ( aService = EmptyStr ) then
  begin
    WarningMsg( '請先設定【優惠組合參數】。' );
    MainPageControl.ActivePageIndex := 2;
    Exit;
  end;
  {}
  CloneInstCodeData;
  if InstCodeDataSet.IsEmpty then
  begin
    WarningMsg( '請先設定【派工參數】。' );
    MainPageControl.ActivePageIndex := 0;
    Exit;
  end;
  frmCD078B5 := TfrmCD078B5.Create( aMode, FKeyValue, aService, EmptyStr,
    CD078CDataSet );
  try
    frmCD078B5.Caption := MainPageControl.ActivePage.Caption + ' [' + frmCD078B5.Caption + ']';
    frmCD078B5.FaciItemDataSet := OtherDataSet;
    frmCD078B5.InstCodeDataSet := InstCodeDataSet;
    frmCD078B5.ShowModal;
  finally
    frmCD078B5.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.ShowCD078D_Form(aMode: TDML);
begin
  FKeyValue := edtCodeNo.Text;
  if ( FKeyValue = EmptyStr ) then
  begin
    WarningMsg( '請先輸入優惠組合代碼。' );
    if edtCodeNo.CanFocusEx then edtCodeNo.SetFocus;
    Exit;
  end;
  CloneInstCodeData;
  if InstCodeDataSet.IsEmpty then
  begin
    WarningMsg( '請先設定【派工參數】。' );
    MainPageControl.ActivePageIndex := 0;
    Exit;
  end;
  frmCD078B2 := TfrmCD078B2.Create( aMode, FKeyValue, EmptyStr, CD078DDataSet );
  try
    frmCD078B2.Caption := MainPageControl.ActivePage.Caption + ' [' + frmCD078B2.Caption + ']';
    frmCD078B2.InstCodeDataSet := InstCodeDataSet;
    frmCD078B2.ShowModal;
  finally
    frmCD078B2.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.DataSetAfterDeleteOrPost(DataSet: TDataSet);
begin
  ButtonStateChange;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.CloneCodeData;
var
  aBMark: TBookmark;
begin
  if ( OtherDataSet.FieldDefs.Count <= 0 ) then
  begin
    OtherDataSet.FieldDefs.Clear;
    OtherDataSet.FieldDefs.Add( 'CODENO', ftString, 10 );
    OtherDataSet.FieldDefs.Add( 'DESCRIPTION', ftString, 1000 );
    OtherDataSet.FieldDefs.Add( 'SERVICETYPE', ftString, 1 );
    OtherDataSet.FieldDefs.Add( 'FACITYPE', ftString, 1 );
    OtherDataSet.FieldDefs.Add( 'REFNO', ftString, 3 );
    OtherDataSet.CreateDataSet;
  end;
  if not VarIsNull( OtherDataSet.Data ) then
    OtherDataSet.EmptyDataSet;
  aBMark := CD078DDataSet.GetBookmark;
  try
    CD078DDataSet.First;
    while not CD078DDataSet.Eof do
    begin
      OtherDataSet.AppendRecord( [
        CD078DDataSet.FieldByName( 'FACIITEM' ).AsString,
        CD078DDataSet.FieldByName( 'FACINAME' ).AsString,
        CD078DDataSet.FieldByName( 'SERVICETYPE' ).AsString,
        CD078DDataSet.FieldByName( 'FACITYPE' ).AsString,
        CD078DDataSet.FieldByName( 'REFNO' ).AsString] );
      CD078DDataSet.Next;
    end;
    CD078DDataSet.GotoBookmark( aBMark );
  finally
    CD078DDataSet.FreeBookmark( aBMark );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.CloneInstCodeData;
var
  aBMark: TBookmark;
begin
  if not VarIsNull( InstCodeDataSet.Data ) then
    InstCodeDataSet.EmptyDataSet;
  InstCodeDataSet.Data := Null;
  InstCodeDataSet.FieldDefs.Clear;
  InstCodeDataSet.FieldDefs.Add( 'INSTCODE', ftString, 10 );
  InstCodeDataSet.FieldDefs.Add( 'INSTNAME', ftString,  1000 );
  InstCodeDataSet.FieldDefs.Add( 'SERVICETYPE', ftString, 2 );
  InstCodeDataSet.CreateDataSet;
  {}
  aBMark := CD078BDataSet.GetBookmark;
  try
    CD078BDataSet.DisableControls;
    try
      CD078BDataSet.First;
      while not CD078BDataSet.Eof do
      begin
        InstCodeDataSet.AppendRecord( [
          CD078BDataSet.FieldByName( 'INSTCODE' ).AsString,
          CD078BDataSet.FieldByName( 'INSTNAME' ).AsString,
          CD078BDataSet.FieldByName( 'SERVICETYPE' ).AsString] );
        CD078BDataSet.Next;
      end;
      CD078BDataSet.GotoBookmark( aBMark );
    finally
      CD078BDataSet.EnableControls;
    end;
  finally
    CD078BDataSet.FreeBookmark( aBMark );
  end;
  {}
  CD078B1CloneDataSet.First;
  while not CD078B1CloneDataSet.Eof  do
  begin
    if ( not InstCodeDataSet.Locate( 'INSTCODE',
      CD078B1CloneDataSet.FieldByName( 'INSTCODE2' ).AsString, [] ) ) then
    begin
      InstCodeDataSet.AppendRecord( [
        CD078B1CloneDataSet.FieldByName( 'INSTCODE2' ).AsString,
        CD078B1CloneDataSet.FieldByName( 'INSTNAME2' ).AsString,
        CD078B1CloneDataSet.FieldByName( 'SERVICETYPE' ).AsString] );
    end;
    CD078B1CloneDataSet.Next;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.LoadMCItetmCodeData;
var
  aBMark: TBookmark;
begin
  if ( MCItemDataSet.FieldDefs.Count <= 0 ) then
  begin
    MCItemDataSet.FieldDefs.Add( 'CODENO', ftInteger );
    MCItemDataSet.FieldDefs.Add( 'DESCRIPTION', ftString, 100 );
    MCItemDataSet.FieldDefs.Add( 'SERVICETYPE', ftString, 2 );
    MCItemDataSet.CreateDataSet;
  end;
  if not VarIsNull( MCItemDataSet.Data ) then
     MCItemDataSet.EmptyDataSet;
  aBMark := CD078ADataSet.GetBookmark;
  try
    CD078ADataSet.First;
    while not CD078ADataSet.Eof do
    begin
      if not SameText( CD078ADataSet.FieldByName( 'SIGN' ).AsString, '-' ) then
      begin
        MCItemDataSet.AppendRecord( [
          CD078ADataSet.FieldByName( 'CITEMCODE' ).AsString,
          CD078ADataSet.FieldByName( 'CITEMNAME' ).AsString,
          CD078ADataSet.FieldByName( 'SERVICETYPE' ).AsString] );
      end;
      CD078ADataSet.Next;
    end;
    CD078ADataSet.GotoBookmark( aBMark );
  finally
    CD078ADataSet.FreeBookmark( aBMark );
  end;

end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.RemoveInstCodeStr(const AInstCode: String);
var
  aInstCodeStr: String;
begin
  CD078DDataSet.First;
  while not CD078DDataSet.Eof do
  begin
    aInstCodeStr := CD078DDataSet.FieldByName( 'InstCodeStr' ).AsString;
    if ValueExists( AInstCode, aInstCodeStr )then
    begin
      aInstCodeStr := DeleteValue( AInstCode, aInstCodeStr );
      CD078DDataSet.Edit;
      CD078DDataSet.FieldByName( 'InstCodeStr' ).AsString := aInstCodeStr;
      CD078DDataSet.Post;
    end;
    CD078DDataSet.Next;
  end;
  {}
  CD078ADataSet.First;
  while not CD078ADataSet.Eof do
  begin
    aInstCodeStr := CD078ADataSet.FieldByName( 'InstCodeStr' ).AsString;
    if ValueExists( AInstCode, aInstCodeStr )then
    begin
      aInstCodeStr := DeleteValue( AInstCode, aInstCodeStr );
      CD078ADataSet.Edit;
      CD078ADataSet.FieldByName( 'InstCodeStr' ).AsString := aInstCodeStr;
      CD078ADataSet.Post;
    end;
    CD078ADataSet.Next;
  end;
  {}
  CD078CDataSet.First;
  while not CD078CDataSet.Eof do
  begin
    aInstCodeStr := CD078CDataSet.FieldByName( 'InstCodeStr' ).AsString;
    if ValueExists( AInstCode, aInstCodeStr )then
    begin
      aInstCodeStr := DeleteValue( AInstCode, aInstCodeStr );
      CD078CDataSet.Edit;
      CD078CDataSet.FieldByName( 'InstCodeStr' ).AsString := aInstCodeStr;
      CD078CDataSet.Post;
    end;
    CD078CDataSet.Next;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.CloneCD078A1Data;
begin
  CD078A1CloneDataSet.Data := CD078A1DataSet.Data;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.CloneCD078A2Data;
begin
  CD078A2CloneDataSet.Data := CD078A2DataSet.Data;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.CloneCD078A3Data;
begin
  CD078A3CloneDataSet.Data := CD078A3DataSet.Data;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.CloneCD078A4Data;
begin
  CD078A4CloneDataSet.Data := CD078A4DataSet.Data;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.CloneCD078B1Data;
begin
  CD078B1CloneDataSet.Data := CD078B1DataSet.Data;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.FilterCD078A1Data(const AEnableFilter: Boolean; const aMode: TDML);
begin
  CD078A1DataSet.Filtered := False;

  if ( AEnableFilter ) then
  begin
    if ( aMode = dmInsert ) then
    begin
      CD078A1DataSet.Filter := Format(
        'BPCODE=''%s'' AND CITEMCODE=''-1'' AND SERVICETYPE=''-1''',
        [VarToStrDef( CD078ADataSet.FieldByName( 'BPCode' ).Value, '-1' )] );
    end else
    begin
      CD078A1DataSet.Filter := Format(
        'BPCODE=''%s'' AND CITEMCODE=''%s'' AND SERVICETYPE=''%s''',
        [VarToStrDef( CD078ADataSet.FieldByName( 'BPCode' ).Value, '-1' ),
         VarToStrDef( CD078ADataSet.FieldByName( 'CitemCode' ).Value, '-1' ),
         VarToStrDef( CD078ADataSet.FieldByName( 'ServiceType' ).Value, '-1' )] );
    end;
    CD078A1DataSet.Filtered := True;
  end else
  begin
    CD078A1DataSet.Filter := EmptyStr;
  end;
  CD078A1DataSet.First;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.RestoreCD078A1Data;
begin
  FilterCD078A1Data( False, FMode );
  CD078A1DataSet.Data := CD078A1CloneDataSet.Data;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.DeleteCD078A1Data;
begin
  CD078A1CloneDataSet.Filtered := False;
  CD078A1CloneDataSet.Filter := Format(
    'BPCODE=''%s'' AND CITEMCODE=''%s'' AND SERVICETYPE=''%s''',
    [VarToStrDef( CD078ADataSet.FieldByName( 'BPCode' ).Value, '-1' ),
     VarToStrDef( CD078ADataSet.FieldByName( 'CitemCode' ).Value, '-1' ),
     VarToStrDef( CD078ADataSet.FieldByName( 'ServiceType' ).Value, '-1' )] );
  CD078A1CloneDataSet.Filtered := True;
  try
    while not CD078A1CloneDataSet.Eof do
      CD078A1CloneDataSet.Delete;
  finally
    CD078A1CloneDataSet.Filtered := False;
    CD078A1CloneDataSet.Filter := EmptyStr;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.FilterCD078A2Data(const AEnableFilter: Boolean; const aMode: TDML);
begin
  CD078A2DataSet.Filtered := False;
  if ( AEnableFilter ) then
  begin
    if ( aMode = dmInsert ) then
    begin
      CD078A2DataSet.Filter := Format(
        'BPCODE=''%s'' AND CITEMCODE=''-1'' AND SERVICETYPE=''-1''',
        [VarToStrDef( CD078ADataSet.FieldByName( 'BPCode' ).Value, '-1' )] );
    end else
    begin
      CD078A2DataSet.Filter := Format(
        'BPCODE=''%s'' AND CITEMCODE=''%s'' AND SERVICETYPE=''%s''',
        [VarToStrDef( CD078ADataSet.FieldByName( 'BPCode' ).Value, '-1' ),
         VarToStrDef( CD078ADataSet.FieldByName( 'CitemCode' ).Value, '-1' ),
         VarToStrDef( CD078ADataSet.FieldByName( 'ServiceType' ).Value, '-1' )] );
    end;
    CD078A2DataSet.Filtered := True;
  end else
  begin
    CD078A2DataSet.Filter := EmptyStr;
  end;
  CD078A2DataSet.First;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.RestoreCD078A2Data;
begin
  FilterCD078A2Data( False, FMode );
  CD078A2DataSet.Data := CD078A2CloneDataSet.Data;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.DeleteCD078A2Data;
begin
  CD078A2CloneDataSet.Filtered := False;
  CD078A2CloneDataSet.Filter := Format(
    'BPCODE=''%s'' AND CITEMCODE=''%s'' AND SERVICETYPE=''%s''',
    [VarToStrDef( CD078ADataSet.FieldByName( 'BPCode' ).Value, '-1' ),
     VarToStrDef( CD078ADataSet.FieldByName( 'CitemCode' ).Value, '-1' ),
     VarToStrDef( CD078ADataSet.FieldByName( 'ServiceType' ).Value, '-1' )] );
  CD078A2CloneDataSet.Filtered := True;
  try
    while not CD078A2CloneDataSet.Eof do
      CD078A2CloneDataSet.Delete;
  finally
    CD078A2CloneDataSet.Filtered := False;
    CD078A2CloneDataSet.Filter := EmptyStr;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.FilterCD078A3Data(const AEnableFilter: Boolean;
  const aMode: TDML);
begin
  CD078A3DataSet.Filtered := False;
  if ( AEnableFilter ) then
  begin
    if ( aMode = dmInsert ) then
    begin
      CD078A3DataSet.Filter := Format(
        'BPCODE=''%s'' AND CITEMCODE=''-1'' AND SERVICETYPE=''-1''',
        [VarToStrDef( CD078ADataSet.FieldByName( 'BPCode' ).Value, '-1' )] );
    end else
    begin
      CD078A3DataSet.Filter := Format(
        'BPCODE=''%s'' AND CITEMCODE=''%s'' AND SERVICETYPE=''%s''',
        [VarToStrDef( CD078ADataSet.FieldByName( 'BPCode' ).Value, '-1' ),
         VarToStrDef( CD078ADataSet.FieldByName( 'CitemCode' ).Value, '-1' ),
         VarToStrDef( CD078ADataSet.FieldByName( 'ServiceType' ).Value, '-1' )] );
    end;
    CD078A3DataSet.Filtered := True;
  end else
  begin
    CD078A3DataSet.Filter := EmptyStr;
  end;
  CD078A3DataSet.First;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.RestoreCD078A3Data;
begin
  FilterCD078A3Data( False, FMode );
  CD078A3DataSet.Data := CD078A3CloneDataSet.Data;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.DeleteCD078A3Data;
begin
  CD078A3CloneDataSet.Filtered := False;
  CD078A3CloneDataSet.Filter := Format(
    'BPCODE=''%s'' AND CITEMCODE=''%s'' AND SERVICETYPE=''%s''',
    [VarToStrDef( CD078ADataSet.FieldByName( 'BPCode' ).Value, '-1' ),
     VarToStrDef( CD078ADataSet.FieldByName( 'CitemCode' ).Value, '-1' ),
     VarToStrDef( CD078ADataSet.FieldByName( 'ServiceType' ).Value, '-1' )] );
  CD078A3CloneDataSet.Filtered := True;
  try
    while not CD078A3CloneDataSet.Eof do
      CD078A3CloneDataSet.Delete;
  finally
    CD078A3CloneDataSet.Filtered := False;
    CD078A3CloneDataSet.Filter := EmptyStr;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.FilterCD078A4Data(const AEnableFilter: Boolean;
  const aMode: TDML);
begin
  CD078A4DataSet.Filtered := False;
  if ( AEnableFilter ) then
  begin
    if ( aMode = dmInsert ) then
    begin
      CD078A4DataSet.Filter := Format(
        'BPCODE=''%s'' AND CITEMCODE=''-1'' AND SERVICETYPE=''-1''',
        [VarToStrDef( CD078ADataSet.FieldByName( 'BPCode' ).Value, '-1' )] );
    end else
    begin
      CD078A4DataSet.Filter := Format(
        'BPCODE=''%s'' AND CITEMCODE=''%s'' AND SERVICETYPE=''%s''',
        [VarToStrDef( CD078ADataSet.FieldByName( 'BPCode' ).Value, '-1' ),
         VarToStrDef( CD078ADataSet.FieldByName( 'CitemCode' ).Value, '-1' ),
         VarToStrDef( CD078ADataSet.FieldByName( 'ServiceType' ).Value, '-1' )] );
    end;
    CD078A4DataSet.Filtered := True;
  end else
  begin
    CD078A4DataSet.Filter := EmptyStr;
  end;
  CD078A4DataSet.First;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.RestoreCD078A4Data;
begin
  FilterCD078A4Data( False, FMode );
  CD078A4DataSet.Data := CD078A4CloneDataSet.Data;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.DeleteCD078A4Data;
begin
  CD078A4CloneDataSet.Filtered := False;
  CD078A4CloneDataSet.Filter := Format(
    'BPCODE=''%s'' AND CITEMCODE=''%s'' AND SERVICETYPE=''%s''',
    [VarToStrDef( CD078ADataSet.FieldByName( 'BPCode' ).Value, '-1' ),
     VarToStrDef( CD078ADataSet.FieldByName( 'CitemCode' ).Value, '-1' ),
     VarToStrDef( CD078ADataSet.FieldByName( 'ServiceType' ).Value, '-1' )] );
  CD078A3CloneDataSet.Filtered := True;
  try
    while not CD078A4CloneDataSet.Eof do
      CD078A4CloneDataSet.Delete;
  finally
    CD078A4CloneDataSet.Filtered := False;
    CD078A4CloneDataSet.Filter := EmptyStr;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.FilterCD078B1Data(const AEnableFilter: Boolean;
  const aMode: TDML);
begin
  CD078B1DataSet.Filtered := False;
  if ( AEnableFilter ) then
  begin
    if ( aMode = dmInsert ) then
    begin
      CD078B1DataSet.Filter := Format(
        'BPCODE=''%s'' AND INSTCODE=''-1'' AND SERVICETYPE=''-1''',
        [VarToStrDef( CD078BDataSet.FieldByName( 'BPCode' ).Value, '-1' )] );
    end else
    begin
      CD078B1DataSet.Filter := Format(
        'BPCODE=''%s'' AND INSTCODE=''%s'' AND SERVICETYPE=''%s''',
        [VarToStrDef( CD078BDataSet.FieldByName( 'BPCode' ).Value, '-1' ),
         VarToStrDef( CD078BDataSet.FieldByName( 'InstCode' ).Value, '-1' ),
         VarToStrDef( CD078BDataSet.FieldByName( 'ServiceType' ).Value, '-1' )] );
    end;
    CD078B1DataSet.Filtered := True;
  end else
  begin
    CD078B1DataSet.Filter := EmptyStr;
  end;
  CD078B1DataSet.First;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.RestoreCD078B1Data;
begin
  FilterCD078B1Data( False, FMode );
  CD078B1DataSet.Data := CD078B1CloneDataSet.Data;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.DeleteCD078B1Data;
begin
  CD078B1CloneDataSet.Filtered := False;
  CD078B1CloneDataSet.Filter := Format(
    'BPCODE=''%s'' AND INSTCODE=''%s'' AND SERVICETYPE=''%s''',
    [VarToStrDef( CD078BDataSet.FieldByName( 'BPCode' ).Value, '-1' ),
     VarToStrDef( CD078BDataSet.FieldByName( 'InstCode' ).Value, '-1' ),
     VarToStrDef( CD078BDataSet.FieldByName( 'ServiceType' ).Value, '-1' )] );
  CD078B1CloneDataSet.Filtered := True;
  try
    while not CD078B1CloneDataSet.Eof do
    begin
      { 因為 CD078B1 裏面設定的子項裝機類別, 也有自動 Append 到其它 Table 去,
        所以在刪除的時候也須要移掉 }
      RemoveInstCodeStr( CD078B1CloneDataSet.FieldByName( 'InstCode2' ).AsString );
      CD078B1CloneDataSet.Delete;
    end;
  finally
    CD078B1CloneDataSet.Filtered := False;
    CD078B1CloneDataSet.Filter := EmptyStr;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.CD078ABandedTableViewDEPOSITATTRGetDisplayText(
  Sender: TcxCustomGridTableItem; ARecord: TcxCustomGridRecord;
  var AText: String);
begin
  if ( aText = '1' ) then
    AText := '一般'
  else if ( AText = '2' ) then
    AText := '履約';
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.CD078ABandedTableViewEXPIRETYPEGetDisplayText(
  Sender: TcxCustomGridTableItem; ARecord: TcxCustomGridRecord;
  var AText: String);
begin
  if ( aText = '1' ) then
    AText := '恢復最新公告牌價'
  else if ( AText = '2' ) then
    AText := '以最後一階繼續優惠'
  else if ( AText = '3' ) then
    AText := '到期不出'
  else if ( AText = '5' ) then
    AText := '約滿不跨階';

end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.CD078ABandedTableViewPENALTYPEGetDisplayText(
  Sender: TcxCustomGridTableItem; ARecord: TcxCustomGridRecord;
  var AText: String);
begin
  if ( aText = '1' ) then
    AText := '回溯牌價'
  else if ( AText = '2' ) then
    AText := '優惠價'
  else if ( AText = '3' ) then
    AText := '優惠平均價';

end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.RestoreLinkKeyData;
begin
  LinkKeyDataSet.Filtered := False;
  LinkKeyDataSet.Filter := EmptyStr;
  LinkKeyDataSet.Data := CD078A1CloneDataSet.Data;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.DeleteCD078CIfInstCodeIsEmpty;
begin
  CD078CDataSet.First;
  while not CD078CDataSet.Eof do
  begin
    if CD078CDataSet.FieldByName( 'InstCodeStr' ).AsString = EmptyStr then
      CD078CDataSet.Delete
    else
      CD078CDataSet.Next;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B1.OneOrZero(Status: Boolean): String;
begin
  if Status then
    Result := '1'
  else
    Result := '0';
end;

{ ---------------------------------------------------------------------------- }
{#5318 增加此功能 By Kin 2009/10/06}
function TfrmCD078B1.GetCD078ACItemValue: String;
var
  aMark: TBookmark;
  aList: TStringList;
begin
  Result := EmptyStr;
  aMark := CD078ADataSet.GetBookmark;
  try
    CD078ADataSet.DisableControls;
    try
      aList := TStringList.Create;
      try
        aList.Delimiter := ',';
        CD078ADataSet.First;
        while not CD078ADataSet.Eof do
        begin
          if aList.IndexOf( CD078ADataSet.FieldByName( 'CITEMCODE' ).AsString ) < 0 then
            aList.Add( CD078ADataSet.FieldByName( 'CITEMCODE' ).AsString );
          CD078ADataSet.Next;
        end;
        Result := aList.DelimitedText;
        CD078ADataSet.GotoBookmark( aMark );
      finally
        aList.Free;
      end;
    finally
      CD078ADataSet.EnableControls;
    end;
  finally
    CD078ADataSet.FreeBookmark( aMark );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B1.ApplyCD078: Boolean;
var
  aWriter: TADOQuery;
  aDateStr1, aDateStr2: String;
begin
  aWriter := DBController.DataWriter;
  aWriter.Close;
  { CD078A }
  aWriter.SQL.Text := Format( ' DELETE FROM CD078A WHERE BPCODE = ''%s'' ',
    [FKeyValue] );
  aWriter.ExecSQL;
  { CD078A1 }
  aWriter.SQL.Text := Format( ' DELETE FROM CD078A1 WHERE BPCODE = ''%s'' ',
    [FKeyValue] );
  aWriter.ExecSQL;
  { CD078A2 }
  aWriter.SQL.Text := Format( ' DELETE FROM CD078A2 WHERE BPCODE = ''%s'' ',
    [FKeyValue] );
  aWriter.ExecSQL;
  { CD078A3 }
  aWriter.SQL.Text := Format( ' DELETE FROM CD078A3 WHERE BPCODE = ''%s'' ',
    [FKeyValue] );
  aWriter.ExecSQL;
  { CD078A4 }
  aWriter.SQL.Text := Format( ' DELETE FROM CD078A4 WHERE BPCODE = ''%s'' ',
    [FKeyValue] );
  aWriter.ExecSQL;
  { CD078B }
  aWriter.SQL.Text := Format( ' DELETE FROM CD078B WHERE BPCODE = ''%s'' AND NEWBPCODE IS NULL ',
    [FKeyValue] );
  aWriter.ExecSQL;
  { CD078B1 }
  aWriter.SQL.Text := Format( ' DELETE FROM CD078B1 WHERE BPCODE = ''%s'' ',
    [FKeyValue] );
  aWriter.ExecSQL;
  { CD078C }
  aWriter.SQL.Text := Format( ' DELETE FROM CD078C WHERE BPCODE = ''%s'' AND NEWBPCODE IS NULL ',
    [FKeyValue] );
  aWriter.ExecSQL;
  { CD078D }
  aWriter.SQL.Text := Format( ' DELETE FROM CD078D WHERE BPCODE = ''%s'' AND NEWBPCODE IS NULL ',
    [FKeyValue] );
  aWriter.ExecSQL;
  {#5318 增加CD078G 與 CD078G1}
  aWriter.SQL.Text := Format( ' DELETE FROM CD078G WHERE BPCODE = ''%s'' ',
  [FKeyValue]);
  aWriter.ExecSQL;
  aWriter.SQL.Text := Format( ' DELETE FROM CD078G1 WHERE BPCODE = ''%s'' ',
  [FKeyValue]);
  aWriter.ExecSQL;
  {#6035 增加CD078I By Kin 2011/06/14}
  aWriter.SQL.Text := Format( ' DELETE FROM CD078I WHERE BPCODE = ''%s'' ',
  [FKeyValue]);
  aWriter.ExecSQL;
  {}
  aWriter.SQL.Text := Format( ' DELETE FROM CD078J WHERE BPCODE = ''%s'' ',
  [FKeyValue]);
  aWriter.ExecSQL;
  {}
  {#6881 增加CD078K}
  aWriter.SQL.Text := Format( 'DELETE FROM CD078K WHERE BPCODE = ''%s'' ',
  [FKeyValue]);
  aWriter.ExecSQL;
  {#6931 增加CD078L}
  aWriter.SQL.Text := Format( 'DELETE FROM CD078L WHERE BPCODE = ''%s'' ',
  [FKeyValue]);
  aWriter.ExecSQL;
  aDateStr1 := ToACDateText( edtDateSt.Text );
  aDateStr2 := ToACDateText( edtDateEd.Text );
  {}
  {#5683 增加繳付類別 By Kin 2010/08/16}
  {#6267 增加SyncBPCode 欄位 By Kin 2012/07/19}
  if ( FMode = dmInsert ) then
  begin
    aWriter.Close;
    aWriter.SQL.Text := Format (
        ' INSERT INTO %s.CD078 (                                        ' +
        '   CODENO, DESCRIPTION, ONSALESTARTDATE, ONSALESTOPDATE,       ' +
        '   MASTERPROD, BUNDLE, BUNDLESAMEMON, BUNDLEMON, SAMEPERIOD,   ' +
        '   PERIOD, MERGEPRINT, PENAL, PENALTYPE, PENALSINGLE,          ' +
        '   MERGEWORK, REFNO, UPDTIME, UPDEN,                           ' +
        '   STOPFLAG, KINDFUNCTION,CANPAYNOW,SYNCBPCODE,BYHOUSE )       ' +
        ' VALUES (                                                      ' +
        '   ''%s'', ''%s'', TO_DATE( ''%s'', ''YYYY/MM/DD'' ),          ' +
        '   TO_DATE( ''%s'', ''YYYY/MM/DD'' ),                          ' +
        ' ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                       ' +
        ' ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                       ' +
        ' ''%s'', ''%s'', ''%s'', ''%s'',                               ' +
        ' ''%s'', ''%s'',''%s'',''%s'',''%s'' )                         ',
        [DBController.LoginInfo.DbAccount,
        FKeyValue,
        edtDescription.Text,
        aDateStr1,
        aDateStr2,
        OneOrZero(chkMasterProd.checked),
        '0',
        '0',
        '0',
        '0',
        '0',
        '0',
        '0',
        EmptyStr,
        EmptyStr,
        '0',
        edtRefNo.Text,
        FormatDateTime( 'yyyy/mm/dd hh:nn:ss', Now ),
        DBController.LoginInfo.UserName,
        OneOrZero( chkStopFlag.checked ),
        IntToStr( edtKindFunction.ItemIndex ),
        OneOrZero(chkPayKind.Checked),
        SyncBPCode.Value,
        OneOrZero(chkByHouse.checked)] );
  end else
  begin

    aWriter.SQL.Text := Format(
      ' UPDATE %s.CD078 SET                                    ' +
      '   DESCRIPTION = ''%s'',                                ' +
      '   ONSALESTARTDATE = TO_DATE( ''%s'', ''YYYY/MM/DD'' ), ' +
      '   ONSALESTOPDATE = TO_DATE( ''%s'', ''YYYY/MM/DD'' ),  ' +
      '   MASTERPROD = ''%s'',                                 ' +
      '   BUNDLE = ''%s'',                                     ' +
      '   BUNDLESAMEMON = ''%s'',                              ' +
      '   BUNDLEMON = ''%s'',                                  ' +
      '   SAMEPERIOD = ''%s'',                                 ' +
      '   PERIOD = ''%s'',                                     ' +
      '   MERGEPRINT = ''%s'',                                 ' +
      '   PENAL = ''%s'',                                      ' +
      '   PENALTYPE = ''%s'',                                  ' +
      '   PENALSINGLE = ''%s'',                                ' +
      '   MERGEWORK = ''%s'',                                  ' +
      '   REFNO = ''%s'',                                      ' +
      '   UPDTIME = ''%s'',                                    ' +
      '   UPDEN = ''%s'',                                      ' +
      '   STOPFLAG = ''%s'',                                   ' +
      '   KINDFUNCTION = ''%s'',                               ' +
      '   SYNCBPCODE = ''%s'',                                 ' +
      '   BYHOUSE = ''%s''                                     ' +
      ' WHERE CODENO = ''%s''                                  ',
      [DBController.LoginInfo.DbAccount,
       edtDescription.Text,
       aDateStr1,
       aDateStr2,
       OneOrZero(chkMasterProd.checked),
       '0',
       '0',
       '0',
       '0',
       '0',
       '0',
       '0',
       EmptyStr,
       EmptyStr,
       '0',
       edtRefNo.Text,
       FormatDateTime( 'yyyy/mm/dd hh:nn:ss', Now ),
       DBController.LoginInfo.UserName,
       OneOrZero( chkStopFlag.checked ),
       IntToStr( edtKindFunction.ItemIndex ),
       SyncBPCode.Value,
       OneOrZero( chkByHouse.Checked ),
       edtCodeNo.Text] );
  end;
  aWriter.ExecSQL;
  {}
  { #4435 增加頻道挑選模式 By Kin 2009/06/06 }
  { #5683 增加繳付類別 By Kin 2010/08/16}
  { #6388 增加CANUSETYPE、EPGIMAGEURL、EPGORD 欄位 By Kin 2012/12/18}
  aWriter.SQL.Text := Format(
   ' UPDATE %s.CD078 SET        ' +
   '   BUNDLEPRODNOTE = :1,     ' +
   '   WORKNOTE = :2,           ' +
   '   CONDITIONAL = :3,        ' +
   '   CHOOSEKIND = :4,         ' +
   '   CANPAYNOW = :5,          ' +
   '   CANUSETYPE = :6,         ' +
   '   EPGIMAGEURL = :7,        ' +
   '   EPGORD = :8              ' +
   ' WHERE CODENO = ''%s''      ',
  [DBController.LoginInfo.DbAccount, edtCodeNo.Text] );
  aWriter.Parameters.ParamByName( '1' ).Value := edtBundleProdNote.Text;
  aWriter.Parameters.ParamByName( '2' ).Value := edtWorkNote.Text;
  aWriter.Parameters.ParamByName( '3' ).Value := edtConditional.Text;
  aWriter.Parameters.ParamByName( '4' ).Value := StrToInt(OneOrZero(chkChooseKind.Checked)) ;
  aWriter.Parameters.ParamByName( '5' ).Value := StrToInt(OneOrZero(chkPayKind.Checked));
  aWriter.Parameters.ParamByName( '6' ).Value := cobCanUseType.ItemIndex;
  aWriter.Parameters.ParamByName( '7' ).Value := edtImgURL.Text;
  aWriter.Parameters.ParamByName( '8' ).Value := edtEPGOrder.Text;
  {
  if edtEPGOrder.Text <> EmptyStr then
    aWriter.Parameters.ParamByName( '8' ).Value := edtEPGOrder.Text
  else
    aWriter.Parameters.ParamByName( '8' ).Value := null;
  }
  aWriter.ExecSQL;
  {}
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B1.ApplyCD078A: Boolean;
var
  aWriter: TADOQuery;
  aInsertSql, aUpdateSql: String;
begin
  aWriter := DBController.DataWriter;
  aWriter.Close;
  {}
  //#5318 增加SalePoint與NameSalePointCode 欄位 By Kin 2009/10/06
  //#5257 增加特定優惠截止日 By Kin 2010/05/05
  //#5913 增加破月期數(收費單) By Kin 2011/04/14
  {#6700 增加StopFlag By Kin 2014/1/3}
  {#6841 增加長繳別註記 By Kin 2014/08/05}
  aInsertSql :=
    ' INSERT INTO %s.CD078A ( bpcode, citemcode, citemname, servicetype,' +
    '   faciitem, bundle, bundlemon, penaltype, expiretype, expirecut,  ' +
    '   depositattr, depositcode, depositname, monthamt,                ' +
    '   dayamt, pflag1, pflagdate, truncamt, pflagcode, pflagname,      ' +
    '   mon1, period1, ratetype1, discountamt1, monthamt1, dayamt1,     ' +
    '   punish1, mon2, period2, ratetype2, discountamt2, monthamt2,     ' +
    '   dayamt2, punish2, mon3, period3, ratetype3, discountamt3,       ' +
    '   monthamt3, dayamt3, punish3, mon4, period4, ratetype4,          ' +
    '   discountamt4, monthamt4, dayamt4, punish4, mon5, period5,       ' +
    '   ratetype5, discountamt5, monthamt5, dayamt5, punish5, mon6,     ' +
    '   period6, ratetype6, discountamt6, monthamt6, dayamt6, punish6,  ' +
    '   mon7, period7, ratetype7, discountamt7, monthamt7, dayamt7,     ' +
    '   punish7, mon8, period8, ratetype8, discountamt8, monthamt8,     ' +
    '   dayamt8, punish8, mon9, period9, ratetype9, discountamt9,       ' +
    '   monthamt9, dayamt9, punish9, mon10, period10, ratetype10,       ' +
    '   discountamt10, monthamt10, dayamt10, punish10, mon11, period11, ' +
    '   ratetype11, discountamt11, monthamt11, dayamt11, punish11,      ' +
    '   mon12, period12, ratetype12, discountamt12, monthamt12,         ' +
    '   dayamt12, punish12, penal, penalcode, penalname,                ' +
    '   depositcodestr, discountcalc, mcitemcode,MinCount,MaxCount,     ' +
    '   SalePointcode, SalePointName,PFlag2,StopFlag,LONGPAYflag,       ' +
    '   AuthTunerCount, ContractType                             )      ' +
    ' VALUES ( ''%s'', ''%s'', ''%s'', ''%s'',                          ' +
    '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
    '   ''%s'', ''%s'', ''%s'', ''%s'',                                 ' +
    '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
    '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
    '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
    '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
    '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
    '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
    '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
    '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
    '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
    '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
    '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
    '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
    '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
    '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                         ' +
    '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                         ' +
    '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                         ' +
    '   ''%s'', ''%s'', ''%s'',%d, %d,''%s'',''%s'',%d,%s,%s,           ' +
    '   ''%s'',''%s'' )                                                 ' ;

  {}
  aUpdateSql := Format(
    ' UPDATE %s.CD078A                     ' +
    '    SET Comment1 = :Comment1,         ' +
    '        Comment2 = :Comment2,         ' +
    '        Comment3 = :Comment3,         ' +
    '        Comment4 = :Comment4,         ' +
    '        Comment5 = :Comment5,         ' +
    '        Comment6 = :Comment6,         ' +
    '        Comment7 = :Comment7,         ' +
    '        Comment8 = :Comment8,         ' +
    '        Comment9 = :Comment9,         ' +
    '        Comment10 = :Comment10,       ' +
    '        Comment11 = :Comment11,       ' +
    '        Comment12 = :Comment12,       ' +
    '        InstCodeStr = :InstCodeStr,   ' +
    '        CMBaudNo = :CMBaudNo,         ' +
    '        CMBaudRate = :CMBaudRate,     ' +
    '        BPStopDate = :BPStopDate      ' +
    '  WHERE BpCode = ''%s''               ' +
    '    AND CItemCode = :CItemCode        ' +
    '    AND ServiceType = :ServiceType    ',
    [DBController.LoginInfo.DbAccount, FKeyValue] );
  {}
  CD078ADataSet.DisableControls;
  try
    CD078ADataSet.First;
    while not CD078ADataSet.Eof do
    begin

      aWriter.SQL.Text := Format( aInsertSql, [
         DBController.LoginInfo.DbAccount,
         FKeyValue,
         CD078ADataSet.FieldByName( 'citemcode' ).AsString,
         CD078ADataSet.FieldByName( 'citemname' ).AsString,
         CD078ADataSet.FieldByName( 'servicetype' ).AsString,
         CD078ADataSet.FieldByName( 'faciitem' ).AsString,
         Nvl( CD078ADataSet.FieldByName( 'bundle' ).AsString, '0' ),
         CD078ADataSet.FieldByName( 'bundlemon' ).AsString,
         CD078ADataSet.FieldByName( 'penaltype' ).AsString,
         CD078ADataSet.FieldByName( 'expiretype' ).AsString,
         Nvl( CD078ADataSet.FieldByName( 'expirecut' ).AsString, '0' ),
         CD078ADataSet.FieldByName( 'depositattr' ).AsString,
         CD078ADataSet.FieldByName( 'depositcode' ).AsString,
         CD078ADataSet.FieldByName( 'depositname' ).AsString,
         CD078ADataSet.FieldByName( 'monthamt' ).AsString,
         CD078ADataSet.FieldByName( 'dayamt' ).AsString,
         CD078ADataSet.FieldByName( 'pflag1' ).AsString,
         CD078ADataSet.FieldByName( 'pflagdate' ).AsString,
         CD078ADataSet.FieldByName( 'truncamt' ).AsString,
         CD078ADataSet.FieldByName( 'pflagcode' ).AsString,
         CD078ADataSet.FieldByName( 'pflagname' ).AsString,
         {}
         CD078ADataSet.FieldByName( 'mon1').AsString,
         CD078ADataSet.FieldByName( 'period1').AsString,
         CD078ADataSet.FieldByName( 'ratetype1' ).AsString,
         CD078ADataSet.FieldByName( 'discountamt1' ).AsString,
         CD078ADataSet.FieldByName( 'monthamt1' ).AsString,
         CD078ADataSet.FieldByName( 'dayamt1' ).AsString,
         '0',
         {}
         CD078ADataSet.FieldByName( 'mon2').AsString,
         CD078ADataSet.FieldByName( 'period2').AsString,
         CD078ADataSet.FieldByName( 'ratetype2' ).AsString,
         CD078ADataSet.FieldByName( 'discountamt2' ).AsString,
         CD078ADataSet.FieldByName( 'monthamt2' ).AsString,
         CD078ADataSet.FieldByName( 'dayamt2' ).AsString,
         '0',
         {}
         CD078ADataSet.FieldByName( 'mon3').AsString,
         CD078ADataSet.FieldByName( 'period3').AsString,
         CD078ADataSet.FieldByName( 'ratetype3' ).AsString,
         CD078ADataSet.FieldByName( 'discountamt3' ).AsString,
         CD078ADataSet.FieldByName( 'monthamt3' ).AsString,
         CD078ADataSet.FieldByName( 'dayamt3' ).AsString,
         '0',
         {}
         CD078ADataSet.FieldByName( 'mon4').AsString,
         CD078ADataSet.FieldByName( 'period4').AsString,
         CD078ADataSet.FieldByName( 'ratetype4' ).AsString,
         CD078ADataSet.FieldByName( 'discountamt4' ).AsString,
         CD078ADataSet.FieldByName( 'monthamt4' ).AsString,
         CD078ADataSet.FieldByName( 'dayamt4' ).AsString,
         '0',
         {}
         CD078ADataSet.FieldByName( 'mon5').AsString,
         CD078ADataSet.FieldByName( 'period5').AsString,
         CD078ADataSet.FieldByName( 'ratetype5' ).AsString,
         CD078ADataSet.FieldByName( 'discountamt5' ).AsString,
         CD078ADataSet.FieldByName( 'monthamt5' ).AsString,
         CD078ADataSet.FieldByName( 'dayamt5' ).AsString,
         '0',
         {}
         CD078ADataSet.FieldByName( 'mon6').AsString,
         CD078ADataSet.FieldByName( 'period6').AsString,
         CD078ADataSet.FieldByName( 'ratetype6' ).AsString,
         CD078ADataSet.FieldByName( 'discountamt6' ).AsString,
         CD078ADataSet.FieldByName( 'monthamt6' ).AsString,
         CD078ADataSet.FieldByName( 'dayamt6' ).AsString,
         '0',
         {}
         CD078ADataSet.FieldByName( 'mon7').AsString,
         CD078ADataSet.FieldByName( 'period7').AsString,
         CD078ADataSet.FieldByName( 'ratetype7' ).AsString,
         CD078ADataSet.FieldByName( 'discountamt7' ).AsString,
         CD078ADataSet.FieldByName( 'monthamt7' ).AsString,
         CD078ADataSet.FieldByName( 'dayamt7' ).AsString,
         '0',
         {}
         CD078ADataSet.FieldByName( 'mon8').AsString,
         CD078ADataSet.FieldByName( 'period8').AsString,
         CD078ADataSet.FieldByName( 'ratetype8' ).AsString,
         CD078ADataSet.FieldByName( 'discountamt8' ).AsString,
         CD078ADataSet.FieldByName( 'monthamt8' ).AsString,
         CD078ADataSet.FieldByName( 'dayamt8' ).AsString,
         '0',
         {}
         CD078ADataSet.FieldByName( 'mon9').AsString,
         CD078ADataSet.FieldByName( 'period9').AsString,
         CD078ADataSet.FieldByName( 'ratetype9' ).AsString,
         CD078ADataSet.FieldByName( 'discountamt9' ).AsString,
         CD078ADataSet.FieldByName( 'monthamt9' ).AsString,
         CD078ADataSet.FieldByName( 'dayamt9' ).AsString,
         '0',
         {}
         CD078ADataSet.FieldByName( 'mon10').AsString,
         CD078ADataSet.FieldByName( 'period10').AsString,
         CD078ADataSet.FieldByName( 'ratetype10' ).AsString,
         CD078ADataSet.FieldByName( 'discountamt10' ).AsString,
         CD078ADataSet.FieldByName( 'monthamt10' ).AsString,
         CD078ADataSet.FieldByName( 'dayamt10' ).AsString,
         '0',
         {}
         CD078ADataSet.FieldByName( 'mon11').AsString,
         CD078ADataSet.FieldByName( 'period11').AsString,
         CD078ADataSet.FieldByName( 'ratetype11' ).AsString,
         CD078ADataSet.FieldByName( 'discountamt11' ).AsString,
         CD078ADataSet.FieldByName( 'monthamt11' ).AsString,
         CD078ADataSet.FieldByName( 'dayamt11' ).AsString,
         '0',
         {}
         CD078ADataSet.FieldByName( 'mon12').AsString,
         CD078ADataSet.FieldByName( 'period12').AsString,
         CD078ADataSet.FieldByName( 'ratetype12' ).AsString,
         CD078ADataSet.FieldByName( 'discountamt12' ).AsString,
         CD078ADataSet.FieldByName( 'monthamt12' ).AsString,
         CD078ADataSet.FieldByName( 'dayamt12' ).AsString,
         '0',
         {}
         Nvl( CD078ADataSet.FieldByName( 'penal' ).AsString, '0' ),
         CD078ADataSet.FieldByName( 'penalcode' ).AsString,
         CD078ADataSet.FieldByName( 'penalname' ).AsString,
         {}
         CD078ADataSet.FieldByName( 'depositcodestr' ).AsString,
         CD078ADataSet.FieldByName( 'discountcalc' ).AsString,
         CD078ADataSet.FieldByName( 'mcitemcode' ).AsString,
         {}
         //#4435 測試不OK,這裡沒有加上去,導致存檔後資料沒有進去 By Kin 2009/04/30
         //#5913 增加PFlag2 By Kin 2011/04/14
         CD078ADataSet.FieldByName( 'MinCount' ).AsInteger ,
         CD078ADataSet.FieldByName( 'MaxCount' ).AsInteger ,
         CD078ADataSet.FieldByName( 'SalePointCode').AsString ,
         CD078ADataSet.FieldByName( 'SalePointName' ).AsString,
         CD078ADataSet.FieldByName( 'PFlag2').AsInteger,
         CD078ADataSet.FieldByName( 'STOPFLAG' ).AsString,
         CD078ADataSet.FieldByName( 'LONGPAYFLAG').AsString,
         CD078ADataSet.FieldByName ( 'AuthTunerCount' ).AsString,
         CD078ADataSet.FieldByName ( 'ContractType' ).AsString ] );
      aWriter.ExecSQL;
      {}
      aWriter.SQL.Text := aUpdateSql;
      aWriter.Parameters.ParamByName( 'CItemCode' ).Value := CD078ADataSet.FieldByName( 'CItemCode' ).AsString;
      aWriter.Parameters.ParamByName( 'ServiceType' ).Value := CD078ADataSet.FieldByName( 'ServiceType' ).AsString;
      aWriter.Parameters.ParamByName( 'Comment1' ).Value := CD078ADataSet.FieldByName( 'Comment1' ).AsString;
      aWriter.Parameters.ParamByName( 'Comment2' ).Value := CD078ADataSet.FieldByName( 'Comment2' ).AsString;
      aWriter.Parameters.ParamByName( 'Comment3' ).Value := CD078ADataSet.FieldByName( 'Comment3' ).AsString;
      aWriter.Parameters.ParamByName( 'Comment4' ).Value := CD078ADataSet.FieldByName( 'Comment4' ).AsString;
      aWriter.Parameters.ParamByName( 'Comment5' ).Value := CD078ADataSet.FieldByName( 'Comment5' ).AsString;
      aWriter.Parameters.ParamByName( 'Comment6' ).Value := CD078ADataSet.FieldByName( 'Comment6' ).AsString;
      aWriter.Parameters.ParamByName( 'Comment7' ).Value := CD078ADataSet.FieldByName( 'Comment7' ).AsString;
      aWriter.Parameters.ParamByName( 'Comment8' ).Value := CD078ADataSet.FieldByName( 'Comment8' ).AsString;
      aWriter.Parameters.ParamByName( 'Comment9' ).Value := CD078ADataSet.FieldByName( 'Comment9' ).AsString;
      aWriter.Parameters.ParamByName( 'Comment10' ).Value := CD078ADataSet.FieldByName( 'Comment10' ).AsString;
      aWriter.Parameters.ParamByName( 'Comment11' ).Value := CD078ADataSet.FieldByName( 'Comment11' ).AsString;
      aWriter.Parameters.ParamByName( 'Comment12' ).Value := CD078ADataSet.FieldByName( 'Comment12' ).AsString;
      aWriter.Parameters.ParamByName( 'InstCodeStr' ).Value := CD078ADataSet.FieldByName( 'InstCodeStr' ).AsString;
      aWriter.Parameters.ParamByName( 'CMBaudNo' ).Value := CD078ADataSet.FieldByName( 'CMBaudNo' ).AsString;
      aWriter.Parameters.ParamByName( 'CMBaudRate' ).Value := CD078ADataSet.FieldByName( 'CMBaudRate' ).AsString;
      if CD078ADataSet.FieldByName( 'BPStopDate' ).AsString = EmptyStr then
        aWriter.Parameters.ParamByName( 'BPStopDate' ).Value := EmptyStr
      else
        aWriter.Parameters.ParamByName( 'BPStopDate' ).Value := CD078ADataSet.FieldByName( 'BPStopDate' ).AsDateTime;

      aWriter.ExecSQL;
      {}
      aWriter.Parameters.Clear;
      CD078ADataSet.Next;
    end;
    CD078ADataSet.First;
    Result := True;
  finally
    CD078ADataSet.EnableControls;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B1.ApplyCD078A1: Boolean;
var
  aWriter: TADOQuery;
  aInsetSql, aUpdateSql: String;
  aIndex: Integer;
begin
  aWriter := DBController.DataWriter;
  aWriter.Close;
  {}
  {#6865 增加Period By Kin 2014/09/10}
  aInsetSql :=
    ' INSERT INTO %s.CD078A1( bpcode, citemcode, servicetype,           ' +
    '   faciitem, StepNo, Description, MasterSale,                      ' +
    '   mon1, period1, ratetype1, discountamt1, monthamt1, dayamt1,     ' +
    '   punish1, mon2, period2, ratetype2, discountamt2, monthamt2,     ' +
    '   dayamt2, punish2, mon3, period3, ratetype3, discountamt3,       ' +
    '   monthamt3, dayamt3, punish3, mon4, period4, ratetype4,          ' +
    '   discountamt4, monthamt4, dayamt4, punish4, mon5, period5,       ' +
    '   ratetype5, discountamt5, monthamt5, dayamt5, punish5, mon6,     ' +
    '   period6, ratetype6, discountamt6, monthamt6, dayamt6, punish6,  ' +
    '   mon7, period7, ratetype7, discountamt7, monthamt7, dayamt7,     ' +
    '   punish7, mon8, period8, ratetype8, discountamt8, monthamt8,     ' +
    '   dayamt8, punish8, mon9, period9, ratetype9, discountamt9,       ' +
    '   monthamt9, dayamt9, punish9, mon10, period10, ratetype10,       ' +
    '   discountamt10, monthamt10, dayamt10, punish10, mon11, period11, ' +
    '   ratetype11, discountamt11, monthamt11, dayamt11, punish11,      ' +
    '   mon12, period12, ratetype12, discountamt12, monthamt12,         ' +
    '   dayamt12, punish12, LinkKey,Period  )                           ' +
    ' VALUES ( ''%s'', ''%s'', ''%s'',                                  ' +
    '   ''%s'', ''%s'', ''%s'', ''%s'',                                 ' +
    '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
    '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
    '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
    '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
    '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
    '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
    '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
    '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
    '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
    '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
    '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
    '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                 ' +
    '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                         ' +
    '   ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                         ' +
    '   ''%s'', ''%s'', ''%s'',''%s'' )                                 ';
  {}
  aUpdateSql := Format(
    ' UPDATE %s.CD078A1               ' +
    '    SET Comment1 = :Comment1,    ' +
    '        Comment2 = :Comment2,    ' +
    '        Comment3 = :Comment3,    ' +
    '        Comment4 = :Comment4,    ' +
    '        Comment5 = :Comment5,    ' +
    '        Comment6 = :Comment6,    ' +
    '        Comment7 = :Comment7,    ' +
    '        Comment8 = :Comment8,    ' +
    '        Comment9 = :Comment9,    ' +
    '        Comment10 = :Comment10,  ' +
    '        Comment11 = :Comment11,  ' +
    '        Comment12 = :Comment12,  ' +
    '        DiscountRate1 = :DiscountRate1,   ' +
    '        DiscountRate2 = :DiscountRate2,   ' +
    '        DiscountRate3 = :DiscountRate3,   ' +
    '        DiscountRate4 = :DiscountRate4,   ' +
    '        DiscountRate5 = :DiscountRate5,   ' +
    '        DiscountRate6 = :DiscountRate6,   ' +
    '        DiscountRate7 = :DiscountRate7,   ' +
    '        DiscountRate8 = :DiscountRate8,   ' +
    '        DiscountRate9 = :DiscountRate9,   ' +
    '        DiscountRate10 = :DiscountRate10, ' +
    '        DiscountRate11 = :DiscountRate11, ' +
    '        DiscountRate12 = :DiscountRate12  ' +
    '  WHERE BpCode = ''%s''          ' +
    '    AND StepNo = :StepNo         ',
    [DBController.LoginInfo.DbAccount, FKeyValue] );
  {}
  CD078A1CloneDataSet.First;
  while not CD078A1CloneDataSet.Eof do
  begin
    aWriter.SQL.Text := Format( aInsetSql, [
       DBController.LoginInfo.DbAccount,
       FKeyValue,
       CD078A1CloneDataSet.FieldByName( 'citemcode' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'servicetype' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'faciItem' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'stepno' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'description' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'mastersale' ).AsString,
       {}
       CD078A1CloneDataSet.FieldByName( 'mon1').AsString,
       CD078A1CloneDataSet.FieldByName( 'period1').AsString,
       CD078A1CloneDataSet.FieldByName( 'ratetype1' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'discountamt1' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'monthamt1' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'dayamt1' ).AsString,
       '0',
       {}
       CD078A1CloneDataSet.FieldByName( 'mon2').AsString,
       CD078A1CloneDataSet.FieldByName( 'period2').AsString,
       CD078A1CloneDataSet.FieldByName( 'ratetype2' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'discountamt2' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'monthamt2' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'dayamt2' ).AsString,
       '0',
       {}
       CD078A1CloneDataSet.FieldByName( 'mon3').AsString,
       CD078A1CloneDataSet.FieldByName( 'period3').AsString,
       CD078A1CloneDataSet.FieldByName( 'ratetype3' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'discountamt3' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'monthamt3' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'dayamt3' ).AsString,
       '0',
       {}
       CD078A1CloneDataSet.FieldByName( 'mon4').AsString,
       CD078A1CloneDataSet.FieldByName( 'period4').AsString,
       CD078A1CloneDataSet.FieldByName( 'ratetype4' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'discountamt4' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'monthamt4' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'dayamt4' ).AsString,
       '0',
       {}
       CD078A1CloneDataSet.FieldByName( 'mon5').AsString,
       CD078A1CloneDataSet.FieldByName( 'period5').AsString,
       CD078A1CloneDataSet.FieldByName( 'ratetype5' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'discountamt5' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'monthamt5' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'dayamt5' ).AsString,
       '0',
       {}
       CD078A1CloneDataSet.FieldByName( 'mon6').AsString,
       CD078A1CloneDataSet.FieldByName( 'period6').AsString,
       CD078A1CloneDataSet.FieldByName( 'ratetype6' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'discountamt6' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'monthamt6' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'dayamt6' ).AsString,
       '0',
       {}
       CD078A1CloneDataSet.FieldByName( 'mon7').AsString,
       CD078A1CloneDataSet.FieldByName( 'period7').AsString,
       CD078A1CloneDataSet.FieldByName( 'ratetype7' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'discountamt7' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'monthamt7' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'dayamt7' ).AsString,
       '0',
       {}
       CD078A1CloneDataSet.FieldByName( 'mon8').AsString,
       CD078A1CloneDataSet.FieldByName( 'period8').AsString,
       CD078A1CloneDataSet.FieldByName( 'ratetype8' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'discountamt8' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'monthamt8' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'dayamt8' ).AsString,
       '0',
       {}
       CD078A1CloneDataSet.FieldByName( 'mon9').AsString,
       CD078A1CloneDataSet.FieldByName( 'period9').AsString,
       CD078A1CloneDataSet.FieldByName( 'ratetype9' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'discountamt9' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'monthamt9' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'dayamt9' ).AsString,
       '0',
       {}
       CD078A1CloneDataSet.FieldByName( 'mon10').AsString,
       CD078A1CloneDataSet.FieldByName( 'period10').AsString,
       CD078A1CloneDataSet.FieldByName( 'ratetype10' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'discountamt10' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'monthamt10' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'dayamt10' ).AsString,
       '0',
       {}
       CD078A1CloneDataSet.FieldByName( 'mon11').AsString,
       CD078A1CloneDataSet.FieldByName( 'period11').AsString,
       CD078A1CloneDataSet.FieldByName( 'ratetype11' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'discountamt11' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'monthamt11' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'dayamt11' ).AsString,
       '0',
       {}
       CD078A1CloneDataSet.FieldByName( 'mon12').AsString,
       CD078A1CloneDataSet.FieldByName( 'period12').AsString,
       CD078A1CloneDataSet.FieldByName( 'ratetype12' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'discountamt12' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'monthamt12' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'dayamt12' ).AsString,
       '0',
       CD078A1CloneDataSet.FieldByName( 'linkkey' ).AsString,
       CD078A1CloneDataSet.FieldByName( 'Period' ).AsString]);
       {}
    aWriter.ExecSQL;
    {}
    aWriter.SQL.Text := aUpdateSql;
    aWriter.Parameters.ParamByName( 'StepNo' ).Value := CD078A1CloneDataSet.FieldByName( 'StepNo' ).AsString;
    for aIndex := 1 to 12 do
    begin
      aWriter.Parameters.ParamByName( Format( 'Comment%d', [aIndex] ) ).Value :=
        CD078A1CloneDataSet.FieldByName( Format( 'Comment%d', [aIndex] ) ).AsString;
      aWriter.Parameters.ParamByName( Format( 'DiscountRate%d', [aIndex] ) ).Value :=
        CD078A1CloneDataSet.FieldByName( Format( 'DiscountRate%d', [aIndex] ) ).AsString;
    end;
    aWriter.ExecSQL;
    {}
    aWriter.Parameters.Clear;
    CD078A1CloneDataSet.Next;
  end;
  CD078A1CloneDataSet.First;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B1.ApplyCD078A2: Boolean;
var
  aWriter: TADOQuery;
  aSql: String;
begin
  aWriter := DBController.DataWriter;
  aWriter.Close;
  {}
  aSql := 
    ' INSERT INTO %s.CD078A2 (                  ' +
    '    BPCode, CitemCode, ServiceType,        ' +
    '    DepositCode, PTCode, PTName,           ' +
    '    DepositAmt, DepositDefault )           ' +
    ' VALUES ( ''%s'', ''%s'', ''%s'',          ' +
    '    ''%s'', ''%s'', ''%s'', ''%s'',        ' +
    '    ''%s''  )      ';
  CD078A2CloneDataSet.First;
  while ( not CD078A2CloneDataSet.Eof ) do
  begin
    aWriter.SQL.Text := Format( aSql,
      [DBController.LoginInfo.DbAccount,
       FKeyValue,
       CD078A2CloneDataSet.FieldByName( 'citemcode' ).AsString,
       CD078A2CloneDataSet.FieldByName( 'servicetype' ).AsString,
       CD078A2CloneDataSet.FieldByName( 'depositcode' ).AsString,
       CD078A2CloneDataSet.FieldByName( 'ptcode' ).AsString,
       CD078A2CloneDataSet.FieldByName( 'ptname' ).AsString,
       CD078A2CloneDataSet.FieldByName( 'depositamt' ).AsString,
       CD078A2CloneDataSet.FieldByName( 'depositdefault' ).AsString] );
    aWriter.ExecSQL;
    CD078A2CloneDataSet.Next;
  end;
  CD078A2CloneDataSet.First;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B1.ApplyCD078A3: Boolean;
var
  aWriter: TADOQuery;
  aSql: String;
begin
  aWriter := DBController.DataWriter;
  aWriter.Close;
  {}
  aSql := 
    ' INSERT INTO %s.CD078A3 (                  ' +
    '    BPCode, CitemCode, ServiceType,        ' +
    '    PenalCode, Item, PenalMon, PenalAmt )  ' +
    ' VALUES ( ''%s'', ''%s'', ''%s'',          ' +
    '    ''%s'', ''%s'', ''%s'', ''%s''  )      ';
  CD078A3CloneDataSet.First;
  while ( not CD078A3CloneDataSet.Eof ) do
  begin
    aWriter.SQL.Text := Format( aSql,
      [DBController.LoginInfo.DbAccount,
       FKeyValue,
       CD078A3CloneDataSet.FieldByName( 'citemcode' ).AsString,
       CD078A3CloneDataSet.FieldByName( 'servicetype' ).AsString,
       CD078A3CloneDataSet.FieldByName( 'penalCode' ).AsString,
       CD078A3CloneDataSet.FieldByName( 'item' ).AsString,
       CD078A3CloneDataSet.FieldByName( 'PenalMon' ).AsString,
       CD078A3CloneDataSet.FieldByName( 'PenalAmt' ).AsString] );
    aWriter.ExecSQL;
    CD078A3CloneDataSet.Next;
  end;
  CD078A3CloneDataSet.First;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B1.ApplyCD078A4: Boolean;
var
  aWriter: TADOQuery;
  aSql: String;
begin
  aWriter := DBController.DataWriter;
  aWriter.Close;
  {}
  aSql :=
    ' INSERT INTO %s.CD078A4 (                     ' +
    '    BPCode, CitemCode, ServiceType,           ' +
    '    PenalCode, Item, MonthStart, MonthStop,   ' +
    '    MonthAmt, DecreaseAmt )                   ' +
    ' VALUES ( ''%s'', ''%s'', ''%s'',             ' +
    '    ''%s'', ''%s'', ''%s'', ''%s'' ,          ' +
    '    ''%s'', ''%s''  )                         ';
  CD078A4CloneDataSet.First;
  while ( not CD078A4CloneDataSet.Eof ) do
  begin
    aWriter.SQL.Text := Format( aSql,
      [DBController.LoginInfo.DbAccount,
       FKeyValue,
       CD078A4CloneDataSet.FieldByName( 'citemcode' ).AsString,
       CD078A4CloneDataSet.FieldByName( 'servicetype' ).AsString,
       CD078A4CloneDataSet.FieldByName( 'penalCode' ).AsString,
       CD078A4CloneDataSet.FieldByName( 'item' ).AsString,
       CD078A4CloneDataSet.FieldByName( 'MonthStart' ).AsString,
       CD078A4CloneDataSet.FieldByName( 'MonthStop' ).AsString,
       CD078A4CloneDataSet.FieldByName( 'MonthAmt' ).AsString,
       CD078A4CloneDataSet.FieldByName( 'DecreaseAmt' ).AsString] );
    aWriter.ExecSQL;
    CD078A4CloneDataSet.Next;
  end;
  CD078A4CloneDataSet.First;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B1.ApplyCD078B: Boolean;
var
  aWriter: TADOQuery;
  aSql: String;
begin
  aWriter := DBController.DataWriter;
  aWriter.Close;
  {}
  aSql :=
    '  INSERT INTO %s.CD078B ( bpcode, newbpcode, instcode, instname, ' +
    '    servicetype, groupno, groupname, workunit, refno,            ' +
    '    interdepend, prcode, prname  )                               ' +
    '  VALUES ( ''%s'', ''%s'', ''%s'', ''%s'',                       ' +
    '    ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                      ' +
    '    ''%s'', ''%s'', ''%s''  )                                    ';
  CD078BDataSet.DisableControls;
  try
    CD078BDataSet.First;
    while not CD078BDataSet.Eof do
    begin
      aWriter.SQL.Text := Format( aSql, [
        DBController.LoginInfo.DbAccount,
        FKeyValue,
        CD078BDataSet.FieldByName( 'newbpcode' ).AsString,
        CD078BDataSet.FieldByName( 'instcode' ).AsString,
        CD078BDataSet.FieldByName( 'instname' ).AsString,
        CD078BDataSet.FieldByName( 'servicetype' ).AsString,
        CD078BDataSet.FieldByName( 'groupno' ).AsString,
        CD078BDataSet.FieldByName( 'groupname' ).AsString,
        CD078BDataSet.FieldByName( 'workunit' ).AsString,
        CD078BDataSet.FieldByName( 'refno' ).AsString,
        Nvl( CD078BDataSet.FieldByName( 'interdepend' ).AsString, '0' ),
        CD078BDataSet.FieldByName( 'prcode' ).AsString,
        CD078BDataSet.FieldByName( 'prname' ).AsString] );
      aWriter.ExecSQL;
      CD078BDataSet.Next;
    end;
    CD078BDataSet.First;
    Result := True;
  finally
    CD078BDataSet.EnableControls;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B1.ApplyCD078B1: Boolean;
var
  aWriter: TADOQuery;
  aSql: String;
begin
  aWriter := DBController.DataWriter;
  aWriter.Close;
  {}
  aSql := 
    ' INSERT INTO %s.CD078B1 (                  ' +
    '    BPCode, InstCode, ServiceType,         ' +
    '    InstCode2, InstName2,                  ' +
    '    PrCode2, PrName2  )                    ' +
    ' VALUES ( ''%s'', ''%s'', ''%s'',          ' +
    '    ''%s'', ''%s'',                        ' +
    '    ''%s'', ''%s''    )                    ';
  CD078B1CloneDataSet.First;
  while ( not CD078B1CloneDataSet.Eof ) do
  begin
    aWriter.SQL.Text := Format( aSql,
      [DBController.LoginInfo.DbAccount,
       FKeyValue,
       CD078B1CloneDataSet.FieldByName( 'InstCode' ).AsString,
       CD078B1CloneDataSet.FieldByName( 'ServiceType' ).AsString,
       CD078B1CloneDataSet.FieldByName( 'InstCode2' ).AsString,
       CD078B1CloneDataSet.FieldByName( 'InstName2' ).AsString,
       CD078B1CloneDataSet.FieldByName( 'PrCode2' ).AsString,
       CD078B1CloneDataSet.FieldByName( 'PrName2' ).AsString] );
    aWriter.ExecSQL;
    CD078B1CloneDataSet.Next;
  end;
  CD078B1CloneDataSet.First;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B1.ApplyCD078C: Boolean;
var
  aWriter: TADOQuery;
  aSql: String;
begin
  aWriter := DBController.DataWriter;
  aWriter.Close;
  {}
  aSql :=
     ' INSERT INTO %s.CD078C ( bpcode, newbpcode, citemcode, citemname,  ' +
     '    servicetype, faciitem, amount, discountamt, punish,            ' +
     '    penaltype, refcitemcode, refcitemname, instcodestr )           ' +
     ' VALUES ( ''%s'', Null, ''%s'', ''%s'',                            ' +
     '    ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                        ' +
     '    ''%s'', ''%s'', ''%s'', ''%s''  )                              ';
  CD078CDataSet.DisableControls;
  try
    CD078CDataSet.First;
    while not CD078CDataSet.Eof do
    begin
      aWriter.SQL.Text := Format( aSql, [
        DBController.LoginInfo.DbAccount,
        FKeyValue,
        CD078CDataSet.FieldByName( 'citemcode' ).AsString,
        CD078CDataSet.FieldByName( 'citemname' ).AsString,
        CD078CDataSet.FieldByName( 'servicetype' ).AsString,
        CD078CDataSet.FieldByName( 'faciitem' ).AsString,
        CD078CDataSet.FieldByName( 'amount' ).AsString,
        CD078CDataSet.FieldByName( 'discountamt' ).AsString,
        Nvl( CD078CDataSet.FieldByName( 'punish' ).AsString, '0' ),
        CD078CDataSet.FieldByName( 'penaltype' ).AsString,
        CD078CDataSet.FieldByName( 'refcitemcode' ).AsString,
        CD078CDataSet.FieldByName( 'refcitemname' ).AsString,
        CD078CDataSet.FieldByName( 'instcodestr' ).AsString] );
      aWriter.ExecSQL;  
      CD078CDataSet.Next;
    end;
    CD078CDataSet.First;
    Result := True;
  finally
    CD078CDataSet.EnableControls;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B1.ApplyCD078D: Boolean;
var
  aWriter: TADOQuery;
  aSql: String;
begin
  aWriter := DBController.DataWriter;
  aWriter.Close;
  {}
  {#5859 增加方案變更更新固定IP與浮定IP By Kin 2011/04/27}

  aSql :=
    ' INSERT INTO %s.CD078D ( bpcode, newbpcode, faciitem,          ' +
    '    facitype, facicode, faciname, servicetype, modelcode,      ' +
    '    modelname, buycode, buyname, cmbaudno, cmbaudrate,         ' +
    '    fixipcount, dynipcount, instcodestr,                       ' +
    '    cbpchangefixip, cbpchangedynip  )                          ' +
    ' VALUES ( ''%s'', Null, ''%s'',                                ' +
    '    ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                    ' +
    '    ''%s'', ''%s'', ''%s'', ''%s'', ''%s'',                    ' +
    '    ''%s'', ''%s'',  ''%s'', ''%s'', ''%s'' )   ';
  CD078DDataSet.DisableControls;
  try
    CD078DDataSet.First;
    while not CD078DDataSet.Eof do
    begin
      aWriter.SQL.Text := Format( aSql, [
        DBController.LoginInfo.DbAccount,
        FKeyValue,
        CD078DDataSet.FieldByName( 'faciitem' ).AsString,
        CD078DDataSet.FieldByName( 'facitype' ).AsString,
        CD078DDataSet.FieldByName( 'facicode' ).AsString,
        CD078DDataSet.FieldByName( 'faciname' ).AsString,
        CD078DDataSet.FieldByName( 'servicetype' ).AsString,
        CD078DDataSet.FieldByName( 'modelcode' ).AsString,
        CD078DDataSet.FieldByName( 'modelname' ).AsString,
        CD078DDataSet.FieldByName( 'buycode' ).AsString,
        CD078DDataSet.FieldByName( 'buyname' ).AsString,
        CD078DDataSet.FieldByName( 'cmbaudno' ).AsString,
        CD078DDataSet.FieldByName( 'cmbaudrate' ).AsString,
        CD078DDataSet.FieldByName( 'fixipcount' ).AsString,
        CD078DDataSet.FieldByName( 'dynipcount' ).AsString,
        CD078DDataSet.FieldByName( 'instcodestr' ).AsString,
        CD078DDataSet.FieldByName( 'cbpchangefixip' ).AsString,
        CD078DDataSet.FieldByName( 'cbpchangedynip' ).AsString] );
      aWriter.ExecSQL;
      CD078DDataSet.Next;
    end;
    CD078DDataSet.First;
    Result := True;
  finally
    CD078DDataSet.EnableControls;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B1.ServiceTypeIntoStr: String;
var
  aMark: TBookmark;
  aList: TStringList;
begin
  Result := EmptyStr;
  aMark := CD078ADataSet.GetBookmark;
  try
    CD078ADataSet.DisableControls;
    try
      aList := TStringList.Create;
      try
        aList.Delimiter := ',';
        CD078ADataSet.First;
        while not CD078ADataSet.Eof do
        begin
          if aList.IndexOf( CD078ADataSet.FieldByName( 'SERVICETYPE' ).AsString ) < 0 then
            aList.Add( CD078ADataSet.FieldByName( 'SERVICETYPE' ).AsString );
          CD078ADataSet.Next;
        end;
        Result := aList.DelimitedText;
        CD078ADataSet.GotoBookmark( aMark );
      finally
        aList.Free;
      end;
    finally
      CD078ADataSet.EnableControls;
    end;
  finally
    CD078ADataSet.FreeBookmark( aMark );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B1.DoDeleteCheck: Boolean;
var
  aFaciItem: String;
begin
  Result := False;
  //TODO: 檢核條件必須 Review, 有問題
  if ( MainPageControl.ActivePageIndex in [0] ) then
  begin
    aFaciItem := CD078DDataSet.FieldByName( 'FACIITEM' ).AsString;
    { 指定設備是否有被使用 }
    Result := (
      IsFaciItemExists( aFaciItem, CD078ADataSet ) and
      IsFaciItemExists( aFaciItem, CD078CDataSet ) );
  end
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B1.IsServiceTypeExists(AServiceType: String): Boolean;
var
  aBookMark: TBookmark;
  aRefrence: Integer;
begin
  CD078ADataSet.DisableControls;
  try
    aBookMark := CD078ADataSet.GetBookmark;
    try
      CD078ADataSet.First;
      aRefrence := 0;
      while not CD078ADataSet.Eof do
      begin
        if ( UpperCase( CD078ADataSet.FieldByName( 'SERVICETYPE' ).AsString  ) =
          UpperCase( AServiceType ) ) then Inc( aRefrence );
        if aRefrence > 0 then Break;
        CD078ADataSet.Next;
      end;
      CD078ADataSet.GotoBookmark( aBookMark );
    finally
      CD078ADataSet.FreeBookmark( aBookMark );
    end;
  finally
    CD078ADataSet.EnableControls;
  end;
  Result := ( aRefrence > 0 );
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B1.IsFaciItemExists(AFaciItem: String): Boolean;
begin
  Result := IsFaciItemExists( AFaciItem, CD078DDataSet );
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B1.IsFaciItemExists(AFaciItem: String;
  aDataSet: TClientDataSet): Boolean;
var
  aBookMark: TBookmark;
begin
  aDataSet.DisableControls;
  try
    aBookMark := aDataSet.GetBookmark;
    try
      Result := aDataSet.Locate( 'FACIITEM', AFaciItem, [] );
      aDataSet.GotoBookmark( aBookMark );
    finally
      aDataSet.FreeBookmark( aBookMark );
    end;
  finally
    aDataSet.EnableControls;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B1.IsRefCitemExists(ACitem: String): Boolean;
var
  aRefrence: Integer;
  aBMark: TBookmark;
begin
  CD078ADataSet.DisableControls;
  try
    aBMark := CD078ADataSet.GetBookmark;
    try
      CD078ADataSet.First;
      aRefrence := 0;
      while not CD078ADataSet.Eof do
      begin
        if ( CD078ADataSet.FieldByName( 'CITEMCODE' ).AsString = ACitem ) and
           ( CD078ADataSet.FieldByName( 'BUNDLE' ).AsString = '1' ) then
          Inc( aRefrence );
        if aRefrence > 0 then Break;
        CD078ADataSet.Next;
      end;
      CD078ADataSet.GotoBookmark( aBMark );
    finally
      CD078ADataSet.FreeBookmark( aBMark );
    end;
  finally
    CD078ADataSet.EnableControls;
  end;
  Result := ( aRefrence > 0 );
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B1.GetFaciItemRefNo(AFaciItem: String): String;
var
  aBMark: TBookmark;
begin
  CD078DDataSet.DisableControls;
  try
    aBMark := CD078DDataSet.GetBookmark;
    try
      CD078DDataSet.First;
      while not CD078DDataSet.Eof do
      begin
        if ( AFaciItem = CD078DDataSet.FieldByName( 'FaciItem' ).AsString ) then
        begin
          Result := CD078DDataSet.FieldByName( 'REFNO' ).AsString;
          Break;
        end;
        CD078DDataSet.Next;
      end;
    finally
      CD078DDataSet.FreeBookmark( aBMark );
    end;
  finally
    CD078DDataSet.EnableControls;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B1.GetItemByHouse(ACitem: String): Integer;
begin
  FReader.Close;
  FReader.SQL.Text := Format(
    ' SELECT A.BYHOUSE FROM %s.CD019 A  ' +
    '  WHERE A.CODENO = ''%s''          ',
    [DBController.LoginInfo.DbAccount, ACitem] );
  FReader.Open;
  Result := FReader.FieldByName( 'BYHOUSE' ).AsInteger;
  FReader.Close;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B1.GetKindFunction: string;
begin
  Result := IntToStr( edtKindFunction.ItemIndex );
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B1.DataValidate(var AErrSource: String): Boolean;
begin
  AErrSource := EmptyStr;
  Result := VdMustInput;
  if not Result then Exit;
  {}
  Result := VdCD078B;

  {************************************}
  { #5180 讓User可以強迫存檔 By Kin 2009/07/13 }
  if FForceSave then
  begin
    FForceSave:= False;
    Result:= True;
    Exit;
  end;
  {************************************}

  if not Result then
  begin
    AErrSource := 'CD078B';
    Exit;
  end;
  {}
  Result := VdCD078A;
  if not Result then
  begin
    AErrSource := 'CD078A';
    Exit;
  end;
  {}
  Result := VdCD078A1;
  if not Result then
  begin
    AErrSource := 'CD078A';
    Exit;
  end;
  {}
  Result := VdCD078A2;
  if not Result then
  begin
    AErrSource := 'CD078A';
    Exit;
  end;
  {}
  Result := VdCD078C;
  if not Result then
  begin
    AErrSource := 'CD078C';
    Exit;
  end;
  {}
  Result := VdCD078D;
  if not Result then Exit;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B1.VdMustInput: Boolean;
begin
  Result := False;
  if ( edtCodeNo.Text = EmptyStr ) then
  begin
    WarningMsg( '請輸入【優惠組合】。' );
    if edtCodeNo.CanFocusEx then edtCodeNo.SetFocus;
    Exit;
  end;
  if ( edtDateSt.Text = EmptyStr ) or
     ( edtDateEd.Text = EmptyStr ) then
  begin
    WarningMsg( '請輸入【產品上架期間】。' );
    Exit;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B1.VdCD078A: Boolean;
var
  aBookMark: TBookmark;
  aFaciItem, aDeposit, aPenal, aPenalCode, aInstCodeStr: String;
  aIsByHouse: Boolean;
begin
  Result := True;
  if CD078ADataSet.IsEmpty then
  begin
    Result := False;
    MainPageControl.ActivePageIndex := 2;
    WarningMsg( '請設定【優惠組合參數】。' );
    Exit;
  end;
  CD078ADataSet.DisableControls;
  try
    aBookMark := CD078ADataSet.GetBookmark;
    try
      CD078ADataSet.First;
      while not CD078ADataSet.Eof do
      begin
        { 取以機或以戶為單位 }
        {
        aIsByHouse := ( GetItemByHouse( CD078ADataSet.FieldByName( 'CItemCode' ).AsString ) = 1 );
        aFaciItem := CD078ADataSet.FieldByName( 'FACIITEM' ).AsString;
        if ( aFaciItem = EmptyStr ) and ( not aIsByHouse ) then
        begin
          Result := False;
          MainPageControl.ActivePageIndex := 2;
          WarningMsg( '【優惠組合參數】中, 該收費項目以機為單位, 請輸入【指定設備】。' );
          Exit;
        end;
        }
        { 檢查設備參數 }
        if ( aFaciItem <> EmptyStr ) then
        begin
          if ( not IsFaciItemExists( aFaciItem ) ) then
          begin
            Result := False;
            MainPageControl.ActivePageIndex := 2;
            WarningMsg( '【優惠組合參數】中【指定設備】對應不到【設備參數】中的設備順序。' );
            Exit;
          end;
          {
          aFaciItemRefNo := GetFaciItemRefNo( aFaciItem );
          if ( StrToIntDef( aFaciItemRefNo, -1 ) in [2,5] ) and
             ( CD078ADataSet.FieldByName( 'SERVICETYPE' ).AsString = 'I' ) and
             ( CD078ADataSet.FieldByName( 'CMBAUDNO' ).AsString = EmptyStr ) then
          begin
            Result := False;
            MainPageControl.ActivePageIndex := 2;
            WarningMsg( '【優惠組合參數】中【收費項目】對應之 CM 速率無設定值，請設定。。' );
            Exit;
          end;
          }
        end;
        { 檢查階段式優惠參數 }
        FilterCD078A1Data( True, dmSave );
        try
          if ( CD078A1DataSet.RecordCount = 0 ) then
          begin
            Result := False;
            MainPageControl.ActivePageIndex := 2;
            WarningMsg( '【優惠組合參數】中, 尚未設定【階段式優惠參數】。' );
            Exit;
          end;
        finally
          FilterCD078A1Data( False, dmSave );
        end;
        { 檢查保證金 }
        aDeposit := CD078ADataSet.FieldByName( 'DepositCode' ).AsString;
        if ( aDeposit <> EmptyStr ) then
        begin
          FilterCD078A2Data( True, dmSave );
          try
            if ( CD078A2DataSet.RecordCount <= 0 ) then
            begin
              Result := False;
              MainPageControl.ActivePageIndex := 2;
              WarningMsg( '【優惠組合參數】中, 未設定各【付款種類保證金】。' );
              Exit;
            end;
          finally
            FilterCD078A2Data( False, dmSave );
          end;
        end;
        { 檢查階段式違約參數 }
        aPenal := CD078ADataSet.FieldByName( 'Penal' ).AsString;
        if ( aPenal = '1' ) then
        begin
          aPenalCode := CD078ADataSet.FieldByName( 'PenalCode' ).AsString;
          if ( aPenalCode = EmptyStr ) then
          begin
            Result := False;
            MainPageControl.ActivePageIndex := 2;
            WarningMsg( '【優惠組合參數】中, 已設定【違約金按月攤提】, 但是未設定【違約金收費】項目。' );
            Exit;
          end;
          FilterCD078A3Data( True, dmSave );
          try
            if ( CD078A3DataSet.RecordCount <= 0 ) then
            begin
              Result := False;
              MainPageControl.ActivePageIndex := 2;
              WarningMsg( '【優惠組合參數】中, 已設定【違約金按月攤提】, 但是未設定【階段式違約金】項目。' );
              Exit;
            end;
          finally
            FilterCD078A3Data( False, dmSave );
          end;
        end;
        { 檢查派工類別 }
        aInstCodeStr := CD078ADataSet.FieldByName( 'InstCodeStr' ).AsString;
        if ( aInstCodeStr = EmptyStr ) then
        begin
          Result := False;
          MainPageControl.ActivePageIndex := 2;
          WarningMsg( '【優惠組合參數】中, 【指定派工類別】必須輸入。' );
          Exit;
        end;
        CloneInstCodeData;
        if not DBController.VdInstCodeExitisInCD078B( aInstCodeStr, InstCodeDataSet ) then
        begin
          Result := False;
          MainPageControl.ActivePageIndex := 2;
          WarningMsg( '【優惠組合參數】中, 【指定派工類別】中的裝機類別在【智慧型派工參數】中並不存在。' );
          Exit;
        end;
        { 檢查計費機制 }
        {#5663 把此判斷拿掉 By Kin 2010/05/31}
        {
        if ( frmCD078B1.KindFunction = '1' ) then
        begin
          if ( CD078ADataSet.FieldByName( 'REFNO' ).AsString <> '2' ) then
          begin
            Result := False;
            MainPageControl.ActivePageIndex := 2;
            WarningMsg( '計費機制已指定為【挑選搭配】, 【優惠組合參數】中收費項目的參考號必須為2。' );
            Exit;
          end;
        end;
        }
        CD078ADataSet.Next;
      end;
      CD078ADataSet.GotoBookmark( aBookMark );
    finally
      CD078ADataSet.FreeBookmark( aBookMark );
    end;
  finally
    CD078ADataSet.EnableControls;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B1.VdCD078A1: Boolean;
var
  aMasterSaleCount: Integer;
  aOldCItemCode: String;
begin
  Result := False;
  aMasterSaleCount := 0;
  aOldCItemCode := EmptyStr;
  CD078A1CloneDataSet.Filtered := False;
  CD078A1CloneDataSet.Filter := 'MasterSale=1';
  CD078A1CloneDataSet.Filtered := True;
  try
    CD078A1CloneDataSet.First;
    while not CD078A1CloneDataSet.Eof do
    begin
      if ( CD078A1CloneDataSet.FieldByName( 'CItemCode' ).AsString <> aOldCItemCode ) then
      begin
        Inc( aMasterSaleCount );
        aOldCItemCode := CD078A1CloneDataSet.FieldByName( 'CItemCode' ).AsString;
      end;
      CD078A1CloneDataSet.Next;
    end;
    {}
    if ( aMasterSaleCount < CD078A1CloneDataSet.RecordCount ) then
    begin
      MainPageControl.ActivePageIndex := 2;
      WarningMsg( '優惠組合參數, 多階優惠參數必須設定【主推多階】。' );
      Exit;
    end;
    if ( aMasterSaleCount > CD078A1CloneDataSet.RecordCount ) then
    begin
      MainPageControl.ActivePageIndex := 2;
      WarningMsg( Format(
        '優惠組合參數, 多階優惠設定有誤, 【主推多階】只能設定一項, 現設定成%d項。',
        [aMasterSaleCount] ) );
      Exit;
    end;
  finally
    CD078A1CloneDataSet.Filtered := False;
    CD078A1CloneDataSet.Filter := EmptyStr;
  end;
  CD078A1CloneDataSet.First;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B1.VdCD078A2: Boolean;
type
  TA2Record = record
    BpCode: String;
    CitemCode: String;
    ServiceType: String;
  end;
var
  aRecord: TA2Record;
  aDefaultCount: Integer;
begin
  Result := False;
  {}
  aRecord.BpCode := EmptyStr;
  aRecord.CitemCode := EmptyStr;
  aRecord.ServiceType := EmptyStr;
  {}
  CD078A2CloneDataSet.AddIndex( 'IDX1', 'BPCODE;CITEMCODE;SERVICETYPE;PTCODE', [] );
  try
    CD078A2CloneDataSet.IndexName := 'IDX1';
    CD078A2CloneDataSet.First;
    {}
    aRecord.BpCode := CD078A2CloneDataSet.FieldByName( 'BpCode' ).AsString;
    aRecord.CitemCode := CD078A2CloneDataSet.FieldByName( 'CItemCode' ).AsString;
    aRecord.ServiceType := CD078A2CloneDataSet.FieldByName( 'ServiceType' ).AsString;
    {}
    aDefaultCount := 0;
    {}
    while not CD078A2CloneDataSet.Eof do
    begin
      if ( aRecord.BpCode <> CD078A2CloneDataSet.FieldByName( 'BpCode' ).AsString ) or
         ( aRecord.CitemCode <> CD078A2CloneDataSet.FieldByName( 'CItemCode' ).AsString ) or
         ( aRecord.ServiceType <> CD078A2CloneDataSet.FieldByName( 'ServiceType' ).AsString ) then
      begin
        if ( aDefaultCount <> 1 ) then Break;
        aDefaultCount := 0;
        aRecord.BpCode := CD078A2CloneDataSet.FieldByName( 'BpCode' ).AsString;
        aRecord.CitemCode := CD078A2CloneDataSet.FieldByName( 'CItemCode' ).AsString;
        aRecord.ServiceType := CD078A2CloneDataSet.FieldByName( 'ServiceType' ).AsString;
      end;
      if ( CD078A2CloneDataSet.FieldByName( 'DepositDefault' ).AsString = '1' ) then
        Inc( aDefaultCount );
      CD078A2CloneDataSet.Next;
    end;
    if ( not CD078A2CloneDataSet.IsEmpty ) then
    begin
      if ( aDefaultCount <= 0 ) then
      begin
        MainPageControl.ActivePageIndex := 2;
        WarningMsg( '【優惠組合參數】中必須指定一筆【預設付款保證金】。' );
        Exit;
      end else
      if ( aDefaultCount >= 2 ) then
      begin
        MainPageControl.ActivePageIndex := 2;
        WarningMsg( '【優惠組合參數】中【預設付款保證金】設定重覆, 預設付款保證金只可設定一筆。。' );
        Exit;
      end;
    end;
  finally
    CD078A2CloneDataSet.DeleteIndex( 'IDX1' );
  end;
  CD078A2CloneDataSet.First;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B1.VdCD078B: Boolean;
var
  aService: String;
  aBookMark: TBookmark;
begin
  Result := False;
  if ( CD078BDataSet.IsEmpty ) then
  begin
    MainPageControl.ActivePageIndex := 0;
    WarningMsg( '請先設定【智慧型派工參數】' );
    Exit;
  end;
  CD078BDataSet.DisableControls;
  try
    aBookMark := CD078BDataSet.GetBookmark;
    try
      CD078BDataSet.First;
      while not CD078BDataSet.Eof do
      begin
        aService := CD078BDataSet.FieldByName( 'SERVICETYPE' ).AsString;
        if not IsServiceTypeExists( aService ) then
        begin
          MainPageControl.ActivePageIndex := 0;
          { #5180 改成可以讓使用者強迫存檔 By Kin 2009/07/13 }
          if MessageDlg('無設定收費資料，是否存檔？',mtConfirmation,[mbYes,mbNo],0)=mrYes then
          begin
            Result:= True;
            FForceSave:= True;
          end
          else
            FForceSave:=False;
          Exit;

          //WarningMsg( Format( '派工參數中, 無此服務別【%s】。', [aService] ) );
          //Exit;
        end;
        CD078BDataSet.Next;
      end;
      CD078BDataSet.GotoBookmark( aBookMark );
    finally
      CD078BDataSet.FreeBookmark( aBookMark );
    end;
  finally
    CD078BDataSet.EnableControls;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B1.VdCD078C: Boolean;
var
  aService, aFaciItem, aRefCitem, aInstCodeStr: String;
  aBookMark: TBookmark;
begin
  Result := False;
  {
  if ( CD078CDataSet.IsEmpty ) then
  begin
    MainPageControl.ActivePageIndex := 3;
    WarningMsg( '請設定【施工/設定/設備之費用參數】。' );
    Exit;
  end;
  }
  CD078CDataSet.DisableControls;
  try
    aBookMark := CD078CDataSet.GetBookmark;
    try
      CD078CDataSet.First;
      while not CD078CDataSet.Eof do
      begin
        aService := CD078CDataSet.FieldByName( 'SERVICETYPE' ).AsString;
        if not IsServiceTypeExists( aService ) then
        begin
          MainPageControl.ActivePageIndex := 3;
          WarningMsg( Format( '施工/設定/設備之費用參數中, 無此服務別【%s】。',
            [aService]  ) );
          Exit;
        end;
        aFaciItem := CD078CDataSet.FieldByName( 'FACIITEM' ).AsString;
        if ( aFaciItem <> EmptyStr ) and ( not IsFaciItemExists( aFaciItem ) ) then
        begin
          MainPageControl.ActivePageIndex := 3;
          WarningMsg( '施工/設定/設備之費用參數中, 【指定設備】對應不到【設備參數】中的設備順序。' );
          Exit;
        end;
//        if ( DBController.GetItemByHouse( CD078CDataSet.FieldByName( 'CITEMCODE' ).AsString ) = '0' ) then
//        begin
//          if ( aFaciItem = EmptyStr ) or ( not IsFaciItemExists( aFaciItem ) ) then
//          begin
//            MainPageControl.ActivePageIndex := 3;
//            WarningMsg( '施工/設定/設備之費用參數中, 【指定設備】對應不到【設備參數】中的設備順序。' );
//            Exit;
//          end;
//        end;
        aRefCitem := CD078CDataSet.FieldByName( 'REFCITEMCODE' ).AsString;
        if ( aRefCitem <> EmptyStr ) then
        begin
          if not IsRefCitemExists( aRefCitem ) then
          begin
            MainPageControl.ActivePageIndex := 3;
            WarningMsg( '施工/設定/設備之費用參數中, 無此【違約參考收費項目】。' );
            Exit;
          end;
        end;
        aInstCodeStr := CD078CDataSet.FieldByName( 'InstCodeStr' ).AsString;
        if ( aInstCodeStr = EmptyStr ) then
        begin
          MainPageControl.ActivePageIndex := 3;
          WarningMsg( '施工/設定/設備之費用參數中, 【指定派工類別】必須輸入。' );
          Exit;
        end;
        CloneInstCodeData;
        if not DBController.VdInstCodeExitisInCD078B( aInstCodeStr, InstCodeDataSet ) then
        begin
          MainPageControl.ActivePageIndex := 3;
          WarningMsg( '施工/設定/設備之費用參數中, 【指定派工類別】中的裝機類別在【智慧型派工參數】中並不存在。' );
          Exit;
        end;
        CD078CDataSet.Next;
      end;
      CD078CDataSet.GotoBookmark( aBookMark );
    finally
      CD078CDataSet.FreeBookmark( aBookMark );
    end;
  finally
    CD078CDataSet.EnableControls;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B1.VdCD078D: Boolean;
begin
  Result := True;
  if ( CD078DDataSet.IsEmpty ) then
  begin
    MainPageControl.ActivePageIndex := 1;
    Result := ConfirmMsg( '無設定【設備參數】, 是否存檔。' );
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B1.ApplyDataSet: Boolean;
var
  aIndex: Integer;
begin
  {#6881 增加CD078K By Kin 2014/10/15}
  {#6931 增加CD078L By Kin 2014/11/14}
  Result := False;
  if ( FMode in [dmInsert] ) then FKeyValue := edtCodeNo.Text;
  for aIndex := 1 to 12 do
  begin
    case aIndex of
      1: Result := ApplyCD078;
      2: Result := ApplyCD078A;
      3:
        begin
          FilterCD078A1Data( False, FMode );
          Result := ApplyCD078A1;
          if not Result then Exit;
          FilterCD078A2Data( False, FMode );
          Result := ApplyCD078A2;
          if not Result then Exit;
          FilterCD078A3Data( False, FMode );
          Result := ApplyCD078A3;
          if not Result then Exit;
          FilterCD078A4Data( False , FMode );
          Result := ApplyCD078A4;
          if not Result then Exit;
        end;
      4:
        begin
          Result := ApplyCD078B;
          if not Result then Exit;
          FilterCD078B1Data( False, FMode );
          Result := ApplyCD078B1;
        end;
      5: Result := ApplyCD078C;
      6: Result := ApplyCD078D;
      7: Result := ApplyCD078G;
      8: Result := ApplyCD078G1;
      9: Result := ApplyCD078I;
      10: Result := ApplyCD078J;
      11: Result := ApplyCD078K;
      12: Result := ApplyCD078L;
    end;
    if not Result then Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.edtDateStPropertiesValidate(Sender: TObject;
  var DisplayValue: Variant; var ErrorText: TCaption; var Error: Boolean);
begin
  if Error then
  begin
    WarningMsg( '輸入的日期不合法,請重新輸入。' );
    ErrorText := EmptyStr;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.CD078ABandedTableViewCustomDrawCell(
  Sender: TcxCustomGridTableView; ACanvas: TcxCanvas;
  AViewInfo: TcxGridTableDataCellViewInfo; var ADone: Boolean);
var
  aGridRecord: TcxGridDataRow;
begin
  aGridRecord := TcxGridDataRow( AViewInfo.GridRecord );
  if ( not aGridRecord.Focused ) then
    if SameText( aGridRecord.DisplayTexts[CD078ABandedTableViewSIGN.Index], '-' ) then
      ACanvas.Font.Color := clRed;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.CD078CBandedTableViewCustomDrawCell(
  Sender: TcxCustomGridTableView; ACanvas: TcxCanvas;
  AViewInfo: TcxGridTableDataCellViewInfo; var ADone: Boolean);
var
  aGridRecord: TcxGridDataRow;
begin
  aGridRecord := TcxGridDataRow( AViewInfo.GridRecord );
  if ( not aGridRecord.Focused ) then
    if ( StrToInt( VarToStrDef( aGridRecord.Values[CD078CBandedTableViewAMOUNT.Index], '0' ) ) < 0 ) then
      ACanvas.Font.Color := clRed;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.BandedTableViewDblClick(Sender: TObject);
begin
  actUpdate.Execute;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B1.edtKindFunctionPropertiesChange(Sender: TObject);
begin
  if edtKindFunction.ItemIndex = 1 then
    chkChooseKind.Enabled := True
  else
  begin
    chkChooseKind.Enabled := False;
    chkChooseKind.Checked := False;
  end;

end;

procedure TfrmCD078B1.ShowCD078E_Form(aMode: TDML);
var
  aMsg, aCitemCodes: String;
  aControl: TcxCustomTextEdit;

begin
  if ( aMode in [dmInsert, dmUpdate] ) then
  begin
    aMsg := EmptyStr;
    aControl := nil;
    FKeyValue := edtCodeNo.Text;
    if ( FKeyValue = EmptyStr ) then
    begin
      WarningMsg( '請先輸入優惠組合代碼。' );
      if edtCodeNo.CanFocusEx then edtCodeNo.SetFocus;
      Exit;
    end;
    if ( aMsg <> EmptyStr ) then
    begin
      WarningMsg( aMsg );
      if Assigned( aControl ) then
        if aControl.CanFocusEx then
          aControl.SetFocus;
      Exit;
    end;
  end;
  {}
  CloneCodeData;
  CloneInstCodeData;

  if ( InstCodeDataSet.IsEmpty ) then
  begin
    WarningMsg( '請先設定【派工參數】。' );
    MainPageControl.ActivePageIndex := 0;
    Exit;
  end;

  {}

  RestoreCD078G1Data;
  FilterCD078G1Data( True, aMode );

  {}
  {
  RestoreCD078A2Data;
  FilterCD078A2Data( True, aMode );
  }
  {}
  {
  RestoreCD078A3Data;
  FilterCD078A3Data( True, aMode );
  }
  {}
  {
  RestoreCD078A4Data;
  FilterCD078A4Data( True, aMode );
  }
  {}

  aCitemCodes := GetCD078GCItemValue;
  {
  LoadMCItetmCodeData;
  RestoreLinkKeyData;
  }
  {}
  try
    frmCD078GU := TfrmCD078GU.Create( aMode,FKeyValue,CD078GDataSet );
    try
      frmCD078GU.Caption := MainPageControl.ActivePage.Caption + ' [' + frmCD078GU.Caption + ']';
      frmCD078GU.AlreadySetCItemValue := aCitemCodes;
      frmCD078GU.FaciItemDataSet := OtherDataSet;
      frmCD078GU.CD078G1DataSet := CD078G1DataSet;
      frmCD078GU.InstCodeDataSet := InstCodeDataSet;
      frmCD078GU.ShowModal;
    finally
      frmCD078GU.Free;
    end;
  finally
    CD078G1DataSet.AfterScroll := nil;
    FilterCD078G1Data( False, aMode);
  end;


end;

procedure TfrmCD078B1.LoadCD078G_Data;
begin
  CD078GDataSet.Close;
  CD078GDataSet.CommandText := Format(
    ' SELECT ' +
    '        A.BPCODE,                         ' +
    '        A.CITEMCODE,                      ' +
    '        A.CITEMNAME,                      ' +
    '        A.SERVICETYPE,                    ' +
    '        A.FACIITEM,                       ' +
    '        A.SALEPOINTCODE,                  ' +
    '        A.SALEPOINTNAME,                  ' +
    '        A.INSTCODESTR                     ' +
    '   FROM %s.CD078G A                       ' +
    '  WHERE A.BPCODE = ''%s''                 ' +
    '  ORDER BY A.BPCODE                       ' ,
    [DBController.LoginInfo.DbAccount, FKeyValue] );
  CD078GDataSet.Open;
end;

procedure TfrmCD078B1.LoadCD078G1_Data;
begin
  CD078G1DataSet.Close;
  CD078G1DataSet.CommandText := Format(
    ' SELECT A.BPCODE,                       ' +
    '        A.CITEMCODE,                    ' +
    '        A.SERVICETYPE,                  ' +
    '        A.FACIITEM,                     ' +
    '        A.STEPNO,                       ' +
    '        A.DESCRIPTION,                  ' +
    '        A.MASTERSALE,                   ' +
    '        A.DISCOUNTAMT,                  ' +
    '        A.SALEPOINTCODE,                ' +
    '        A.SALEPOINTNAME                 ' +
    '   FROM %s.CD078G1 A                    ' +
    '  WHERE A.BPCODE = ''%s''               ' +
    '  ORDER BY A.BPCODE,STEPNO              ' ,
    [DBController.LoginInfo.DbAccount, FKeyValue] );

  CD078G1DataSet.Open;
end;

procedure TfrmCD078B1.CD078GBandedTableViewDblClick(Sender: TObject);
begin
  actUpdate.Execute;  
end;

function TfrmCD078B1.GetCD078GCItemValue: String;
var
  aMark: TBookmark;
  aList: TStringList;
begin
  Result := EmptyStr;
  aMark := CD078GDataSet.GetBookmark;
  try
    CD078GDataSet.DisableControls;
    try
      aList := TStringList.Create;
      try
        aList.Delimiter := ',';
        CD078GDataSet.First;
        while not CD078GDataSet.Eof do
        begin
          if aList.IndexOf( CD078GDataSet.FieldByName( 'CITEMCODE' ).AsString ) < 0 then
          begin
            if CD078GDataSet.FieldByName( 'CITEMCODE' ).AsString <> EmptyStr then
              aList.Add( CD078GDataSet.FieldByName( 'CITEMCODE' ).AsString );
          end;
          CD078GDataSet.Next;
        end;
        Result := aList.DelimitedText;
        CD078GDataSet.GotoBookmark( aMark );
      finally
        aList.Free;
      end;
    finally
      CD078GDataSet.EnableControls;
    end;
  finally
    CD078GDataSet.FreeBookmark( aMark );
  end;
end;

procedure TfrmCD078B1.CloneCD078G1Data;
begin
  CD078G1CloneDataSet.Data := CD078G1DataSet.Data;
end;
{##5318 增加此功能 By Kin 2009/10/06}
procedure TfrmCD078B1.FilterCD078G1Data(const AEnableFilter: Boolean;
  const aMode: TDML);
begin
  CD078G1DataSet.Filtered := False;
  if ( AEnableFilter ) then
  begin
    if ( aMode = dmInsert ) then
    begin
      CD078G1DataSet.Filter := Format(
        'BPCODE=''%s'' AND CITEMCODE=''-1'' AND SERVICETYPE=''-1''',
        [VarToStrDef( CD078GDataSet.FieldByName( 'BPCode' ).Value, '-1' )] );
    end else
    begin
      CD078G1DataSet.Filter := Format(
        'BPCODE=''%s'' AND CITEMCODE=''%s''                             ',
        [VarToStrDef( CD078GDataSet.FieldByName( 'BPCode' ).Value, '-1' ),
         VarToStrDef( CD078GDataSet.FieldByName( 'CitemCode' ).Value, '-1' ) ] );
    end;
    CD078G1DataSet.Filtered := True;
  end else
  begin
    CD078G1DataSet.Filter := EmptyStr;
  end;
  CD078G1DataSet.First;

end;
{#5318 增加此功能 By Kin 2009/10/06}
procedure TfrmCD078B1.RestoreCD078G1Data;
begin
  FilterCD078G1Data( False, FMode );
  CD078G1DataSet.Data := CD078G1CloneDataSet.Data;
end;

procedure TfrmCD078B1.DeleteCD078G1Data;
begin
  CD078G1CloneDataSet.Filtered := False;
  CD078G1CloneDataSet.Filter := Format(
    'BPCODE=''%s'' AND CITEMCODE=''%s'' AND SERVICETYPE=''%s''',
    [VarToStrDef( CD078GDataSet.FieldByName( 'BPCode' ).Value, '-1' ),
     VarToStrDef( CD078GDataSet.FieldByName( 'CitemCode' ).Value, '-1' ),
     VarToStrDef( CD078GDataSet.FieldByName( 'ServiceType' ).Value, '-1' )] );
  CD078G1CloneDataSet.Filtered := True;
  try
    while not CD078G1CloneDataSet.Eof do
      CD078G1CloneDataSet.Delete;
  finally
    CD078A1CloneDataSet.Filtered := False;
    CD078A1CloneDataSet.Filter := EmptyStr;
  end;
end;

function TfrmCD078B1.ApplyCD078G: Boolean;
var
  aWriter: TADOQuery;
  aSql: String;
begin
  aWriter := DBController.DataWriter;
  aWriter.Close;
  {}
  aSql :=
    ' INSERT INTO %s.CD078G ( BPCODE,CitemCode,CitemName,           ' +
    '   ServiceType,FaciItem,SalePointcode,                         ' +
    'SalePointName,InstCodeStr )                                    ' +
    ' VALUES ( ''%s'', %s , ''%s'', ''%s'',                         ' +
    '       ''%s'', ''%s'', ''%s'',''%s''   )';

  CD078GDataSet.DisableControls;
  try
    CD078GDataSet.First;
    while not CD078GDataSet.Eof do
    begin
      aWriter.SQL.Text := Format( aSql, [
        DBController.LoginInfo.DbAccount,
        FKeyValue,
        CD078GDataSet.FieldByName( 'CitemCode' ).AsString,
        CD078GDataSet.FieldByName( 'CitemName' ).AsString,
        CD078GDataSet.FieldByName( 'ServiceType' ).AsString,
        CD078GDataSet.FieldByName( 'FaciItem' ).AsString,
        CD078GDataSet.FieldByName( 'SalePointCode' ).AsString,
        CD078GDataSet.FieldByName( 'SalePointName').AsString,
        CD078GDataSet.FieldByName( 'InstCodeStr').AsString ] );
      aWriter.ExecSQL;
      CD078GDataSet.Next;

    end;
    CD078DDataSet.First;
    Result := True;
  finally
    CD078GDataSet.EnableControls;
  end;
end;

function TfrmCD078B1.ApplyCD078G1: Boolean;
var
  aWriter: TADOQuery;
  aSql: String;
begin
  aWriter := DBController.DataWriter;
  aWriter.Close;
  {}
  aSql :=
    ' INSERT INTO %s.CD078G1 ( BPCODE,CitemCode,                    ' +
    '   ServiceType,FaciItem,SalePointcode,SalePointName,           ' +
    '   StepNo,Description,MasterSale,DiscountAmt )                 ' +
    ' VALUES ( ''%s'', ''%s'' ,''%s'',                              ' +
    '       ''%s'', ''%s'', ''%s'', %s, ''%s'',                     ' +
    '       %s, %s   )';
  while not CD078G1CloneDataSet.Eof do
  begin
    aWriter.SQL.Text := Format( aSql, [
      DBController.LoginInfo.DbAccount,
      FKeyValue,
      CD078G1CloneDataSet.FieldByName( 'CitemCode' ).AsString,
      CD078G1CloneDataSet.FieldByName( 'ServiceType' ).AsString,
      CD078G1CloneDataSet.FieldByName( 'FaciItem' ).AsString,
      CD078G1CloneDataSet.FieldByName( 'SalePointCode' ).AsString,
      CD078G1CloneDataSet.FieldByName( 'SalePointName' ).AsString,
      CD078G1CloneDataSet.FieldByName( 'StepNo' ).AsString,
      CD078G1CloneDataSet.FieldByName( 'Description' ).AsString,
      CD078G1CloneDataSet.FieldByName( 'MasterSale' ).AsString,
      CD078G1CloneDataSet.FieldByName( 'DiscountAmt' ).AsString ]);
    aWriter.ExecSQL;
    CD078G1CloneDataSet.Next;
  end;
  CD078G1CloneDataSet.First;
  Result := True;

end;

procedure TfrmCD078B1.LoadCD078I_Data;
begin
  CD078IDataSet.Close;
  CD078IDataSet.CommandText := Format(
    ' SELECT * ' +
    '   FROM %s.CD078I A                       ' +
    '  WHERE A.BPCODE = ''%s''                 ' +
    '  ORDER BY A.BPCODE                       ' ,
    [DBController.LoginInfo.DbAccount, FKeyValue] );
  CD078IDataSet.Open;
end;

procedure TfrmCD078B1.CloneCD078IData;
begin
  CD078ICloneDataSet.Data := CD078IDataSet.Data;
end;

procedure TfrmCD078B1.LoadSO562_Data;
 var aCasid, aSQL,aSQL2,aSQL3,aCitems : string;
begin
  if DBController.LoginInfo.StarPVOD then
  begin
    aCitems := GetCD078ACItemValue;

    CD078IDataSet.First;
    while not CD078IDataSet.Eof do
    begin
      if aCasid = EmptyStr then
        aCasid := CD078IDataSet.FieldByName( 'CASID' ).AsString
      else
        aCasid := aCasid + ',' + CD078IDataSet.FieldByName( 'CASID' ).AsString;
      CD078IDataSet.Next;
    end;
    if aCasid = EmptyStr then aCasid := '-1';
    if aCitems = EmptyStr then aCitems := '-1';
    {包含在CD078I的資料}
    aSQL := Format( ' SELECT A.*,1 CHOICE,0 FLAG,B.TITLE SHOWTITLE,   ' +
                    ' -1 CITEMCODE2,LPAD('' '',40) CITEMNAME2,        ' +
                    ' 0 DISCOUNTAMT,-1 FACIITEM,                      ' +
                    ' LPAD('' '',2000) INSTCODESTR,''D'' SERVICETYPE  ' +
                    ' FROM SO562 A,SO563 B                            ' +
                    ' WHERE A.CASID IN ( %s )                         ' +
                    ' AND NVL(LOCALLANGUAGE,0) = 0                    ' +
                    ' AND B.TYPE = 1 AND A.ID = B.ID                  ',
                    [aCasid] );

    {SO562 裡有效的產品}

    aSQL2 := Format( ' SELECT A.*,1 CHOICE,2 FLAG,B.TITLE SHOWTITLE,                ' +
                    ' NVL(A.PREFERENTIALCITEM,-1) CITEMCODE2,                       ' +
                    ' LPAD('' '',40) CITEMNAME2,                                    ' +
                    ' 0 DISCOUNTAMT,-1 FACIITEM,                                    ' +
                    ' LPAD('' '',2000) INSTCODESTR,''D'' SERVICETYPE                ' +
                    ' FROM SO562 A, SO563 B                                         ' +
                    ' WHERE A.ENDPURCHASE >= TO_DATE(''%s'',''YYYYMMDDHH24MISS'' )  ' +
                    ' AND A.CASID NOT IN ( %s )                                     ' +
                    ' AND A.TYPE IN ( 1, 2 )                                        ' +
                    ' AND B.TYPE = 1 AND A.ID = B.ID                                ' +
                    ' AND NVL(LOCALLANGUAGE,0) = 0                                  ' +
                    ' AND ONSALESTOPDATE IS NULL                                    ',
                    [FormatDateTime('yyyymmddhhmmss',Now),aCasid] );

    {包月}
    {#6385 包月的產品名稱要用SO562.Title By Kin 2012/11/29}
    {#6385 選取資料的選項拿掉,修改進入[frmCD078I]時該畫面的資料皆可被修改不需再勾選,Choice 永遠設成1 By Kin 2012/11/29}
    aSQL3 := Format( ' SELECT A.*,1 CHOICE, -1 FLAG,A.TITLE SHOWTITLE,              ' +
                      ' NVL(A.PREFERENTIALCITEM,-1) CITEMCODE2,                     ' +
                      ' LPAD('' '',40) CITEMNAME2,                                  ' +
                      ' 0 DISCOUNTAMT,-1 FACIITEM,                                  ' +
                      ' LPAD('' '',2000) INSTCODESTR,''D'' SERVICETYPE              ' +
                      ' FROM SO562 A , SO563 B                                      ' +
                      ' WHERE A.TYPE = 0 AND A.CITEMCODE IN ( %s )                  ' +
                      ' AND NVL(LOCALLANGUAGE,0) = 0                                ' +
                      ' AND A.ID = B.ID AND B.TYPE= 1                               ',
                      [aCitems] );

    aSQL := 'SELECT DISTINCT *                                                       ' +
            'FROM  ( ' + aSQL + ' UNION ALL ' + aSQL2 + ' UNION ALL ' + aSQL3 + ' )  ' +
            'ORDER BY CHOICE DESC , FLAG ';

    SO562DataSet.Close;
    SO562DataSet.CommandText := aSQL;
    SO562DataSet.Open;
    LoadSO562Citem_Data( True );


    if not SO562DataSet.IsEmpty then
    begin
      if MainPageControl.ActivePageIndex = 5 then
        CD078IGrid.SetFocus;
      CD078IBandedTableView.ViewData.Rows[0].Focused := True;
    end;

    {
    CD078IBandedTableView.Columns[0].SortIndex :=0;
    CD078IBandedTableView.Columns[0].SortOrder := soDescending;

    }


    {
    SO562DataSet.Filtered := False;
    try
      SO562DataSet.Filtered := False;
      SO562DataSet.Filter :='FLAG = 0 ';
      SO562DataSet.Filtered := True;
      SO562DataSet.First;
      while not SO562DataSet.Eof do
      begin

        SO562DataSet.Next;
      end;
    finally
      SO562DataSet.Filter := EmptyStr;
      SO562DataSet.Filtered := False;
      SO562DataSet.First;
    end;
    }







    {
    if SO562DataSet.Locate( 'CASID',VarArrayOf([152]),[]) then
      ShowMessage( 'Yes' )
    else
      ShowMessage( 'No' );
    }
    {
    PrepareSO562DataSet;
    while not SO562DataSet.Eof do
    begin
      SO562_1_DataSet.Append;
      SO562_1_DataSet.FieldByName( 'CHOICE' ).AsString :=
        SO562DataSet.FieldByName( 'CHOICE' ).AsString;
      

      SO562DataSet.Next;
    end;
    }

  end;

end;

procedure TfrmCD078B1.PrepareSO562DataSet;
begin
  if ( SO562_1_DataSet.FieldDefs.Count <= 0 ) then
  begin
    SO562_1_DataSet.FieldDefs.Add( 'CHOICE',ftInteger );
    SO562_1_DataSet.FieldDefs.Add( 'TYPE',ftInteger );
    SO562_1_DataSet.CreateDataSet;
  end;
  SO562_1_DataSet.EmptyDataSet;
end;

procedure TfrmCD078B1.cxGridDBBandedColumn2GetDisplayText(
  Sender: TcxCustomGridTableItem; ARecord: TcxCustomGridRecord;
  var AText: String);
begin
  if ( AText = '1' ) then
    AText := 'PPV';
  if ( AText = '2' ) then
    AText := 'PPP';
  if ( AText = '0' ) then
    AText := '包月';
  {
  if ( aText = '1' ) then
    AText := '一一'
  else if ( AText = '2' ) then
    AText := '履約';
    }
end;

procedure TfrmCD078B1.cxGridDBBandedColumn3GetDisplayText(
  Sender: TcxCustomGridTableItem; ARecord: TcxCustomGridRecord;
  var AText: String);
begin
  if AText = EmptyStr then
  begin
    AText := '無效';
    Exit;
  end;
  try
      if Now <= StrToDateTime( AText ) then
        AText := '有效'
      else
        AText := '無效';
  except
    AText := '無效';
  end;


end;

procedure TfrmCD078B1.DeleteSO562Data;
begin
  if not DBController.LoginInfo.StarPVOD then Exit;
  SO562DataSet.Filtered := False;
  SO562DataSet.Filter := Format(
    'CITEMCODE=''%s'' AND TYPE = 0 ',
    [VarToStrDef( CD078ADataSet.FieldByName( 'CitemCode' ).Value,'-1')] );
  SO562DataSet.Filtered := True;
  try
    while not SO562DataSet.Eof do
      SO562DataSet.Delete;
  finally
    SO562DataSet.Filtered := False;
    SO562DataSet.Filter := EmptyStr;
  end;
end;

procedure TfrmCD078B1.CD078IBandedTableViewCustomDrawCell(
  Sender: TcxCustomGridTableView; ACanvas: TcxCanvas;
  AViewInfo: TcxGridTableDataCellViewInfo; var ADone: Boolean);
  var
    aGridRecord: TcxGridDataRow;
    aDrawDisable : Boolean;
    aStatus: String;
    aDrawRed : Boolean;
begin
  aGridRecord := TcxGridDataRow( AViewInfo.GridRecord );
  aStatus := VarAsType(aGridRecord.Values[cxGridDBBandedColumn3.Index],varString );
  aDrawDisable := VarAsType( aGridRecord.Values[cxGridDBBandedColumn2.Index],varInteger ) = 0;
  aDrawRed := False;
  if aDrawDisable then
  begin
    ACanvas.Brush.Color := clBtnFace;
    ACanvas.Font.Color := clDefault;
  end;

  if aStatus = EmptyStr then
  begin
    aDrawRed := True;
  end else
  begin
    try
      if Now > StrToDateTime( aStatus ) then
        aDrawRed := True;
    except
      aDrawRed := True;
    end;

  end;
  if aDrawRed then
  begin
    ACanvas.Font.Color := clRed;
  end;


end;

procedure TfrmCD078B1.FilterSO562Data(const AEnableFilter: Boolean;
  const aMode: TDML);
begin
  if not DBController.LoginInfo.StarPVOD then
  begin
    Exit;
  end;
  SO562DataSet.Filtered := False;
  if ( AEnableFilter ) then
  begin
      SO562DataSet.Filter := Format(
        'CITEMCODE=''%s'' AND TYPE = 0 ',
        [VarToStrDef( CD078ADataSet.FieldByName( 'CitemCode' ).Value,'-1')] );
      SO562DataSet.Filtered := True;
  end else
  begin
    SO562DataSet.Filter := EmptyStr;
  end;
  SO562DataSet.First;

end;

procedure TfrmCD078B1.ChangeSO562Data(const aOldCitemCode: String;
  const aMode: TDML);
  var aSQL,aCitemName : string;
      i:Integer;

begin
  aSQL := EmptyStr;
  i := 0;
  if not DBController.LoginInfo.StarPVOD then Exit;
  try
    if ( aMode = dmUpdate ) then
    begin
      //修改模式沒有變動收費項目，不做任何事情，直接離開
      if ( CD078ADataSet.FieldByName( 'CITEMCODE' ).AsString = aOldCitemCode ) then
      begin
        Exit;
      end else
      begin
        // 有變更收費項目，將Filter過的資料先刪除掉，
        //下段程式再判斷是否需要補到SO562
        //此段一定要有，因為User可能不是按存檔而是按取消
        while not SO562DataSet.Eof do
        begin
          SO562DataSet.Delete;
        end;
      end;
    end;

    FilterSO562Data( True, aMode );

    aSQL := Format('SELECT A.*,1 CHOICE, -1 FLAG,                ' +
                      ' NVL(A.PREFERENTIALCITEM,-1) CITEMCODE2,  ' +
                      ' 0 DISCOUNTAMT,B.TITLE SHOWTITLE          ' +
                      ' FROM SO562 A , SO563 B                   ' +
                      ' WHERE A.TYPE = 0                         ' +
                      ' AND A.CITEMCODE = %s                     ' +
                      ' AND NVL(LOCALLANGUAGE,0) = 0             ' +
                      ' AND A.ID = B.ID AND B.TYPE = 1           ',
                       [ CD078ADataSet.FieldByName( 'CITEMCODE' ).AsString ] );
    while not SO562DataSet.Eof do
    begin
      SO562DataSet.Delete;
    end;
    try
      DBController.ComReader.Close;
      DBController.ComReader.SQL.Text := aSQL;
      DBController.ComReader.Open;
      if ( not DBController.ComReader.IsEmpty ) then
      begin
        DBController.ComReader.First;
        while not DBController.ComReader.Eof do
        begin
          SO562DataSet.Append;
          for i := 0  to DBController.ComReader.Fields.Count - 1  do
          begin
              SO562DataSet.FieldByName( DBController.ComReader.Fields[i].FieldName  ).AsString :=
                DBController.ComReader.Fields[i].AsString;
              if UpperCase( DBController.ComReader.Fields[i].FieldName ) = UpperCase( 'CITEMCODE2' ) then
              begin
                if DBController.ComReader.Fields[i].AsString <> EmptyStr then
                begin
                  SO562DataSet.FieldByName( 'CITEMNAME2' ).AsString :=
                     DBController.GetCitemName( DBController.ComReader.Fields[i].AsString );
                end;
              end;
          end;
          SO562DataSet.Post;
          DBController.ComReader.Next;
        end;
      end;

    finally
      DBController.ComReader.Close;
    end;

  except
     on E: Exception do
     begin
        ErrorMsg( Format( 'PVOD 收費項目填入有誤, 原因:%s。', [E.Message] ) );
     end;
  end;

end;

procedure TfrmCD078B1.LoadSO562Citem_Data(const aLoad : Boolean);
 var aSQL,aCitemCode : string;

begin
  SO562DataSet.Filtered := False;
  try
    if aLoad then
    begin
      {
      SO562DataSet.Filter := 'FLAG = 0 AND CITEMCODE2 = -1 ';
      SO562DataSet.Filtered := True;
      }
      if  SO562DataSet.IsEmpty then
        Exit;
      if SO562DataSet.RecordCount = 0 then
        Exit;
      SO562DataSet.First;
      while not SO562DataSet.Eof do
      begin
         if ( SO562DataSet.FieldByName( 'FLAG' ).AsInteger = 0 ) and
            ( SO562DataSet.FieldByName( 'CITEMCODE2' ).AsString = '-1' ) then
         begin
          CD078IDataSet.Filtered := False;
          CD078IDataSet.Filter := EmptyStr;
          CD078IDataSet.Filter := Format('CASID = ''%s''',
                          [SO562DataSet.FieldByName( 'CASID' ).AsString]);
          CD078IDataSet.Filtered := True;
           if ( not CD078IDataSet.IsEmpty ) then
           begin
             SO562DataSet.Edit;
             SO562DataSet.FieldByName( 'CITEMCODE2' ).AsString :=
                CD078IDataSet.FieldByName('DISCOUNTCITEMCODE').AsString;
             SO562DataSet.FieldByName( 'DISCOUNTAMT' ).AsString :=
              CD078IDataSet.FieldByName( 'DISCOUNTAMT' ).AsString;

            SO562DataSet.FieldByName( 'FACIITEM' ).AsString :=
              CD078IDataSet.FieldByName( 'FACIITEM' ).AsString;

            SO562DataSet.FieldByName( 'INSTCODESTR' ).AsString :=
              CD078IDataSet.FieldByName( 'INSTCODESTR' ).AsString;
             SO562DataSet.Post;
           end;
         end;

         SO562DataSet.Next;

      end;
    end;
    SO562DataSet.Filter := EmptyStr;
    SO562DataSet.Filtered := False;
    SO562DataSet.First;
    {
    SO562DataSet.Filter := 'CITEMCODE2 > 0';
    SO562DataSet.Filtered := True;
    }
    SO562DataSet.First;
    while not SO562DataSet.Eof do
    begin
      if SO562DataSet.FieldByName( 'CITEMCODE2' ).AsFloat > 0 then
      begin
        if Trim(SO562DataSet.FieldByName( 'CITEMNAME2' ).AsString) = EmptyStr then
        begin
          SO562DataSet.Edit;
          SO562DataSet.FieldByName( 'CITEMNAME2' ).AsString :=
            DBController.GetCitemName( SO562DataSet.FieldByName( 'CITEMCODE2' ).AsString );
          SO562DataSet.Post;
        end;
      end;
      SO562DataSet.Next;
    end;

  finally
    SO562DataSet.Filtered := False;
    CD078IDataSet.Filtered := False;
    CD078IDataSet.Filter := EmptyStr;
    SO562DataSet.Filter := EmptyStr;
    SO562DataSet.First;


  end;


end;

procedure TfrmCD078B1.cxGridDBBandedColumn7GetDisplayText(
  Sender: TcxCustomGridTableItem; ARecord: TcxCustomGridRecord;
  var AText: String);
begin
  if Trim( AText ) = EmptyStr then
  begin
    AText := EmptyStr;
  end;

end;

function TfrmCD078B1.ApplyCD078I: Boolean;
var
  aWriter: TADOQuery;
  aSql: String;
  aFaciItem,aDiscountCitemCode:String;

begin
  if not DBController.LoginInfo.StarPVOD then
  begin
    Result := True;
    Exit;
  end;
  aWriter := DBController.DataWriter;
  aWriter.Close;
  SO562DataSet.First;
  {}
  aSql :=
    ' INSERT INTO %s.CD078I ( BPCODE,CITEMCODE,                     ' +
    '   SERVICETYPE,FACIITEM,DISCOUNTCITEMCODE,CASID,               ' +
    '   DISCOUNTAMT,INSTCODESTR )                                   ' +
    ' VALUES ( ''%s'', ''%s'' ,''%s'',                              ' +
    '       ''%s'', ''%s'', ''%s'', %s, ''%s'' )                    ';

  while not SO562DataSet.Eof do
  begin
    if ( SO562DataSet.FieldByName( 'CHOICE' ).AsString = '1' ) and
        ( SO562DataSet.FieldByName( 'TYPE' ).AsString <> '0' ) then
    begin

      aFaciItem := SO562DataSet.FieldByName( 'FACIITEM' ).AsString;
      aDiscountCitemCode := SO562DataSet.FieldByName( 'CITEMCODE2' ).AsString;
      if aFaciItem = '-1' then aFaciItem :='';
      if aDiscountCitemCode = '-1' then aDiscountCitemCode := '';

      aWriter.SQL.Text := Format( aSql, [
        DBController.LoginInfo.DbAccount,
        FKeyValue,
        Trim( SO562DataSet.FieldByName( 'CITEMCODE' ).AsString ),
        SO562DataSet.FieldByName( 'SERVICETYPE' ).AsString,
        aFaciItem,
        aDiscountCitemCode,
        SO562DataSet.FieldByName( 'CASID' ).AsString,
        SO562DataSet.FieldByName( 'DISCOUNTAMT' ).AsString,
        Trim( SO562DataSet.FieldByName( 'INSTCODESTR' ).AsString ) ]);
      aWriter.ExecSQL;
    end;
    SO562DataSet.Next;
  end;
  SO562DataSet.First;
  Result := True;


end;

procedure TfrmCD078B1.ShowCD078I_Form(aMode: TDML);
var
  aMsg, aCitemCodes: String;
  aControl: TcxCustomTextEdit;
begin
  if ( aMode in [dmInsert, dmUpdate] ) then
  begin
    aMsg := EmptyStr;
    aControl := nil;
    FKeyValue := edtCodeNo.Text;
    if ( FKeyValue = EmptyStr ) then
    begin
      WarningMsg( '請先輸入優惠組合代碼。' );
      if edtCodeNo.CanFocusEx then edtCodeNo.SetFocus;
      Exit;
    end;
    if SO562DataSet.FieldByName( 'TYPE' ).AsString = '0' then
    begin
      aMsg := '產品種類為包月，不允許修改！';
    end;
    if ( aMsg <> EmptyStr ) then
    begin
      WarningMsg( aMsg );
      if Assigned( aControl ) then
        if aControl.CanFocusEx then
          aControl.SetFocus;
      Exit;
    end;

  end;
  {}
  CloneCodeData;
  CloneInstCodeData;

  if ( InstCodeDataSet.IsEmpty ) then
  begin
    WarningMsg( '請先設定【派工參數】。' );
    MainPageControl.ActivePageIndex := 0;
    Exit;
  end;

  aCitemCodes := GetCD078GCItemValue;

  try
    frmCD078I := TfrmCD078I.Create( aMode,FKeyValue,SO562DataSet );
    try
      frmCD078I.Caption := MainPageControl.ActivePage.Caption + ' [' + frmCD078I.Caption + ']';
      frmCD078I.AlreadySetCItemValue := aCitemCodes;
      frmCD078I.FaciItemDataSet := OtherDataSet;
      frmCD078I.InstCodeDataSet := InstCodeDataSet;
      frmCD078I.ShowModal;
    finally
      frmCD078I.Free;
    end;
  finally
    SO562DataSet.AfterScroll := nil;
//    FilterCD078G1Data( False, aMode);
  end;


end;

procedure TfrmCD078B1.LoadCD078J_Data;
begin
  CD078JDataSet.Close;
  CD078JDataSet.CommandText := Format(
    ' SELECT ' +
    '        A.BPCODE,                         ' +
    '        A.PRODUCTCODE,                    ' +
    '        A.PRODUCTNAME,                    ' +
    '        A.SERVICETYPE,                    ' +
    '        A.FACIITEM,                       ' +
    '        A.INSTCODESTR                     ' +
    '   FROM %s.CD078J A                       ' +
    '  WHERE A.BPCODE = ''%s''                 ' +
    '  ORDER BY A.BPCODE                       ' ,
    [DBController.LoginInfo.DbAccount, FKeyValue] );
  CD078JDataSet.Open;
end;

procedure TfrmCD078B1.ShowCD078J_Form(aMode: TDML);
begin
  FKeyValue := edtCodeNo.Text;
  if ( FKeyValue = EmptyStr ) then
  begin
    WarningMsg( '請先輸入優惠組合代碼。' );
    if edtCodeNo.CanFocusEx then edtCodeNo.SetFocus;
    Exit;
  end;
  CloneInstCodeData;
  if InstCodeDataSet.IsEmpty then
  begin
    WarningMsg( '請先設定【派工參數】。' );
    MainPageControl.ActivePageIndex := 0;
    Exit;
  end;
  frmCD078JU := TfrmCD078JU.Create( aMode, FKeyValue, EmptyStr, CD078JDataSet );
  try
    frmCD078JU.Caption := MainPageControl.ActivePage.Caption + ' [' + frmCD078JU.Caption + ']';
    frmCD078JU.InstCodeDataSet := InstCodeDataSet;
    frmCD078JU.ShowModal;
  finally
    frmCD078JU.Free;
  end;
end;

function TfrmCD078B1.ApplyCD078J: Boolean;
var
  aWriter: TADOQuery;
  aSql: String;
begin
  aWriter := DBController.DataWriter;
  aWriter.Close;
  {}
  aSql :=
    ' INSERT INTO %s.CD078J ( BPCODE,PRODUCTCODE,PRODUCTNAME,           ' +
    '   SERVICETYPE,FACIITEM,INSTCODESTR   )                            ' +
    ' VALUES ( ''%s'', %s , ''%s'', ''%s'',                             ' +
    '       ''%s'', ''%s''   )';

  CD078JDataSet.DisableControls;
  try
    CD078JDataSet.First;
    while not CD078JDataSet.Eof do
    begin
      aWriter.SQL.Text := Format( aSql, [
        DBController.LoginInfo.DbAccount,
        FKeyValue,
        CD078JDataSet.FieldByName( 'PRODUCTCODE' ).AsString,
        CD078JDataSet.FieldByName( 'PRODUCTNAME' ).AsString,
        CD078JDataSet.FieldByName( 'SERVICETYPE' ).AsString,
        CD078JDataSet.FieldByName( 'FACIITEM' ).AsString,
        CD078JDataSet.FieldByName( 'INSTCODESTR' ).AsString ] );
      aWriter.ExecSQL;
      CD078JDataSet.Next;

    end;
    CD078JDataSet.First;
    Result := True;
  finally
    CD078JDataSet.EnableControls;
  end;
end;

procedure TfrmCD078B1.SyncBPCodeCodeNamePropertiesInitPopup(
  Sender: TObject);
  var aCodeNo : String;
begin
  aCodeNo := '-X';
  if edtCodeNo.Text <> '' then
    aCodeNo := edtCodeNo.Text;
  SyncBPCode.SQL.Text := Format(
    ' SELECT CODENO, DESCRIPTION    ' +
    '  FROM %s.CD078                ' +
    ' WHERE NVL(STOPFLAG,0) = 0     ' +
    ' AND CODENO <> ''%s''          ',
    [DBController.LoginInfo.DbAccount,aCodeNo] );
  SyncBPCode.CodeNamePropertiesInitPopup( Sender );
  SyncBPCode.CodeName.Properties.ListColumns[0].Width := 250;
end;

procedure TfrmCD078B1.SyncBPCodeCodeNamePropertiesChange(Sender: TObject);
begin
  SyncBPCode.CodeNamePropertiesChange( Sender );
//  SyncBPCodeCodeNamePropertiesChange( Sender );
end;

procedure TfrmCD078B1.SyncBPCodeCodeNoPropertiesChange(Sender: TObject);
begin
  SyncBPCode.CodeNoPropertiesChange( Sender );
end;

procedure TfrmCD078B1.edtEPGOrderKeyPress(Sender: TObject; var Key: Char);
begin
  if not(Key in ['0'..'9', #8]) then
    Key := #0;
end;

procedure TfrmCD078B1.cobCanUseTypePropertiesEditValueChanged(
  Sender: TObject);
begin
  edtKindFunction.Enabled := True;
  {#6388 測試不OK,CanUseType =0 一樣不能選 By Kin 2013/03/07}
  if ( cobCanUseType.ItemIndex = 2 ) or ( cobCanUseType.ItemIndex = 0 ) then
  begin
    edtKindFunction.ItemIndex := 0;
    edtKindFunction.Enabled := False;
  end;
end;

procedure TfrmCD078B1.cobCanUseTypePropertiesChange(Sender: TObject);
begin
  {#6560 測試不OK,新增如果沒有權限則不能選擇0與2 By Kin 2013/09/05}

  if ( cobCanUseType.ItemIndex = 2 ) or ( cobCanUseType.ItemIndex = 0 ) then
  begin
    if ( FMode in [dmInsert] ) then
    begin
      if ( not FCanAddEPG ) then
      begin
        MessageDlg( '無此權限選擇適用對象！',mtWarning,[mbOK],0);
        cobCanUseType.ItemIndex := 1;
      end;
    end;
  end;

end;

procedure TfrmCD078B1.btnIFRSClick(Sender: TObject);
  var aService : String;
begin

  FKeyValue := edtCodeNo.Text;
  if ( FKeyValue = EmptyStr ) then
  begin
    WarningMsg( '請先輸入【優惠組合代碼】。' );
    if edtCodeNo.CanFocusEx then edtCodeNo.SetFocus;
    Exit;
  end;
  aService := ServiceTypeIntoStr;
  if ( aService = EmptyStr ) then
  begin
    WarningMsg( '請先設定【優惠組合參數】。' );
    MainPageControl.ActivePageIndex := 2;
    Exit;
  end;

  try
    frmCD078K := TfrmCD078K.Create( CD078ADataSet,CD078KDataSet );
    try

      frmCD078K.ShowModal;
    finally
      frmCD078K.Free;
    end;
  finally

  end;

end;

procedure TfrmCD078B1.LoadCD078K_Data;
begin
  CD078KDataSet.Close;
  CD078KDataSet.CommandText := Format(
    ' SELECT *                                 ' +
    '   FROM %s.CD078K A                       ' +
    '  WHERE A.BPCODE = ''%s''                 ' +
    '  ORDER BY A.BPCODE                       ' ,
    [DBController.LoginInfo.DbAccount, FKeyValue] );
  CD078KDataSet.Open;
end;

function TfrmCD078B1.ApplyCD078K: Boolean;
var
  aWriter: TADOQuery;
  aSql: String;
begin
  aWriter := DBController.DataWriter;
  aWriter.Close;
  {}
  aSql :=
    ' INSERT INTO %s.CD078K ( BPCODE,CITEMCODE,MASTERITEM,              ' +
    '   AMOUNT,IFRSPERCENT,IFRSAMT   )                                  ' +
    ' VALUES ( ''%s'', %s , ''%s'', ''%s'',                             ' +
    '       ''%s'', ''%s''   )';

  CD078KDataSet.DisableControls;
  try
    CD078KDataSet.First;
    while not CD078KDataSet.Eof do
    begin
      aWriter.SQL.Text := Format( aSql, [
        DBController.LoginInfo.DbAccount,
        FKeyValue,
        CD078KDataSet.FieldByName( 'CITEMCODE' ).AsString,
        CD078KDataSet.FieldByName( 'MASTERITEM' ).AsString,
        CD078KDataSet.FieldByName( 'AMOUNT' ).AsString,
        CD078KDataSet.FieldByName( 'IFRSPERCENT' ).AsString,
        CD078KDataSet.FieldByName( 'IFRSAMT' ).AsString ] );
      aWriter.ExecSQL;
      CD078KDataSet.Next;

    end;
    CD078KDataSet.First;
    Result := True;
  finally
    CD078KDataSet.EnableControls;
  end;
end;

procedure TfrmCD078B1.btnShareClick(Sender: TObject);
begin
  FKeyValue := edtCodeNo.Text;
  if ( FKeyValue = EmptyStr ) then
  begin
    WarningMsg( '請先輸入【優惠組合代碼】。' );
    if edtCodeNo.CanFocusEx then edtCodeNo.SetFocus;
    Exit;
  end;

  try
    frmCD078L := TfrmCD078L.Create( FKeyValue,  CD078LDataSet );
    try

      frmCD078L.ShowModal;
    finally
      frmCD078L.Free;
    end;
  finally

  end;

end;

procedure TfrmCD078B1.LoadCD078L_Data;
begin
  CD078LDataSet.Close;
  CD078LDataSet.CommandText := Format(
    ' SELECT *                                 ' +
    '   FROM %s.CD078L A                       ' +
    '  WHERE A.BPCODE = ''%s''                 ' +
    '  ORDER BY A.BPCODE                       ' ,
    [DBController.LoginInfo.DbAccount, FKeyValue] );
  CD078LDataSet.Open;
end;

function TfrmCD078B1.ApplyCD078L: Boolean;
var
  aWriter: TADOQuery;
  aSql: String;
begin
  aWriter := DBController.DataWriter;
  aWriter.Close;
  {}
  aSql :=
    ' INSERT INTO %s.CD078L ( BPCODE,MONTH,MASTERSERVICETYPE,              ' +
    '   CASHADJ,SUBTOTAL,CMAMT,CMTAX,                                      ' +
    '   CMADJ,DTVAMT,DTVTAX,DTVADJ,                                        ' +
    '   PRVAMT,PRVTAX,PRVADJ,STOPFLAG   )                                  ' +
    ' VALUES ( ''%s'', %s , ''%s'',                                        ' +
    '   ''%s'', ''%s'', ''%s'', ''%s'',                                    ' +
    '   ''%s'', ''%s'', ''%s'', ''%s'',                                    ' +
    '   ''%s'', ''%s'', ''%s'', ''%s''   )';

  CD078LDataSet.DisableControls;
  try
    CD078LDataSet.First;
    while not CD078LDataSet.Eof do
    begin
      aWriter.SQL.Text := Format( aSql, [
        DBController.LoginInfo.DbAccount,
        FKeyValue,
        CD078LDataSet.FieldByName( 'MONTH' ).AsString,
        CD078LDataSet.FieldByName( 'MASTERSERVICETYPE' ).AsString,
        CD078LDataSet.FieldByName( 'CASHADJ' ).AsString,
        CD078LDataSet.FieldByName( 'SUBTOTAL' ).AsString,
        CD078LDataSet.FieldByName( 'CMAMT' ).AsString,
        CD078LDataSet.FieldByName( 'CMTAX' ).AsString,
        CD078LDataSet.FieldByName( 'CMADJ' ).AsString,
        CD078LDataSet.FieldByName( 'DTVAMT' ).AsString,
        CD078LDataSet.FieldByName( 'DTVTAX' ).AsString,
        CD078LDataSet.FieldByName( 'DTVADJ' ).AsString,
        CD078LDataSet.FieldByName( 'PRVAMT' ).AsString,
        CD078LDataSet.FieldByName( 'PRVTAX' ).AsString,
        CD078LDataSet.FieldByName( 'PRVADJ' ).AsString,
        CD078LDataSet.FieldByName( 'STOPFLAG' ).AsString
         ] );
      aWriter.ExecSQL;
      CD078LDataSet.Next;
    end;
    CD078LDataSet.First;
    Result := True;
  finally
    CD078LDataSet.EnableControls;
  end;
end;

end.
