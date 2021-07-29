unit frmcd042;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, DBGrids, ExtCtrls, ADODB, DB, Buttons, ActnList,
  DBCtrls, Mask, Provider, DBClient, dxSkinsCore, dxSkinsDefaultPainters,
  dxSkinscxPCPainter, cxStyles, cxCustomData, cxGraphics, cxFilter, cxData, cxDataStorage,
  cxEdit, cxDBData, cxMaskEdit, cxCurrencyEdit, cxButtonEdit, cxGridLevel,
  cxGridCustomTableView, cxGridTableView, cxGridDBTableView, cxClasses, cxControls,
  cxGridCustomView, cxGrid, cxPC, cxGridStrs, dxSkinBlack, dxSkinBlue,
  dxSkinCaramel, dxSkinCoffee;

type
  TForm1 = class(TForm)
    DBGrid1: TDBGrid;
    btnSave: TButton;
    btnPrint: TButton;
    btnInsert: TButton;
    btnModify: TButton;
    ActionList1: TActionList;
    ADOConnection1: TADOConnection;
    ADOQuery1: TADOQuery;
    DataSource1: TDataSource;
    Button10: TButton;
    Edit35: TEdit;
    Button12: TButton;
    Panel5: TPanel;
    ADOQuery2: TADOQuery;
    Action1: TAction;
    Action2: TAction;
    ClientDataSet1: TClientDataSet;
    ClientDataSet1ActStartDate2: TStringField;
    ClientDataSet1ActStopDate2: TStringField;
    ClientDataSet1LastOrderDate2: TStringField;
    DataSetProvider1: TDataSetProvider;
    Action3: TAction;
    Action4: TAction;
    ClientDataSet1CashOrderRate2: TStringField;
    ClientDataSet1CouponOrderRate2: TStringField;
    ClientDataSet1CODENO: TBCDField;
    ClientDataSet1DESCRIPTION: TStringField;
    ClientDataSet1REFNO: TIntegerField;
    ClientDataSet1UPDTIME: TStringField;
    ClientDataSet1UPDEN: TStringField;
    ClientDataSet1COMPCODE: TIntegerField;
    ClientDataSet1ACTSTARTDATE: TDateTimeField;
    ClientDataSet1ACTSTOPDATE: TDateTimeField;
    ClientDataSet1MANAGEMENT: TStringField;
    ClientDataSet1NOTE: TStringField;
    ClientDataSet1DISCOUNTCODE: TBCDField;
    ClientDataSet1DISCOUNTNAME: TStringField;
    ClientDataSet1STOPFLAG: TIntegerField;
    ClientDataSet1LASTORDERDATE: TDateTimeField;
    ClientDataSet1ORDERCOUNT: TIntegerField;
    ClientDataSet1CHANCEDAYS: TIntegerField;
    ClientDataSet1EBTPROJECT: TStringField;
    ClientDataSet1PROMOTIONCONT: TStringField;
    ClientDataSet1WORKNOTE: TStringField;
    ClientDataSet1SERVICETYPESTR: TStringField;
    ClientDataSet1CUSTSTATUSCODESTR: TStringField;
    ClientDataSet1CLASSCODESTR: TStringField;
    ClientDataSet1MDUIDSTR: TStringField;
    ClientDataSet1NODENOSTR: TStringField;
    ClientDataSet1CIRCUITNOSTR: TStringField;
    ClientDataSet1PRODUCTTYPE: TIntegerField;
    ClientDataSet1CITEMCODESTR: TStringField;
    ClientDataSet1BPCODESTR: TStringField;
    ClientDataSet1PUNISH: TIntegerField;
    ClientDataSet1PRICE: TIntegerField;
    ClientDataSet1CASH: TIntegerField;
    ClientDataSet1COUPON: TIntegerField;
    ClientDataSet1GIFT: TIntegerField;
    ClientDataSet1PRICEDISCOUNTTYPE: TIntegerField;
    ClientDataSet1CASHDISCOUNTTYPE: TIntegerField;
    ClientDataSet1COUPONDISCOUNTTYPE: TIntegerField;
    ClientDataSet1PRICEDISCOUNTRATE: TBCDField;
    ClientDataSet1PRICEFIXAMOUNT: TIntegerField;
    ClientDataSet1CASHORDERRATE: TBCDField;
    ClientDataSet1CASHFIXAMOUNT: TIntegerField;
    ClientDataSet1CASHCITEMCODESTR: TStringField;
    ClientDataSet1COUPONORDERRATE: TBCDField;
    ClientDataSet1COUPONFIXAMOUNT: TIntegerField;
    ClientDataSet1GIFTORDERPRICE1: TIntegerField;
    ClientDataSet1GIFTTYPE1: TIntegerField;
    ClientDataSet1ARTICLENOSTR1: TStringField;
    ClientDataSet1GIFTORDERPRICE2: TIntegerField;
    ClientDataSet1GIFTTYPE2: TIntegerField;
    ClientDataSet1ARTICLENOSTR2: TStringField;
    ClientDataSet1GIFTORDERPRICE3: TIntegerField;
    ClientDataSet1GIFTTYPE3: TIntegerField;
    ClientDataSet1ARTICLENOSTR3: TStringField;
    ClientDataSet1GIFTORDERPRICE4: TIntegerField;
    ClientDataSet1GIFTTYPE4: TIntegerField;
    ClientDataSet1ARTICLENOSTR4: TStringField;
    ClientDataSet1GIFTORDERPRICE5: TIntegerField;
    ClientDataSet1GIFTTYPE5: TIntegerField;
    ClientDataSet1ARTICLENOSTR5: TStringField;
    ClientDataSet1GIFTORDERPRICE6: TIntegerField;
    ClientDataSet1GIFTTYPE6: TIntegerField;
    ClientDataSet1ARTICLENOSTR6: TStringField;
    ClientDataSet2: TClientDataSet;
    StringField1: TStringField;
    StringField2: TStringField;
    StringField3: TStringField;
    StringField4: TStringField;
    StringField5: TStringField;
    BCDField1: TBCDField;
    StringField6: TStringField;
    IntegerField1: TIntegerField;
    StringField7: TStringField;
    StringField8: TStringField;
    IntegerField2: TIntegerField;
    DateTimeField1: TDateTimeField;
    DateTimeField2: TDateTimeField;
    StringField9: TStringField;
    StringField10: TStringField;
    BCDField2: TBCDField;
    StringField11: TStringField;
    IntegerField3: TIntegerField;
    DateTimeField3: TDateTimeField;
    IntegerField4: TIntegerField;
    IntegerField5: TIntegerField;
    StringField12: TStringField;
    StringField13: TStringField;
    StringField14: TStringField;
    StringField15: TStringField;
    StringField16: TStringField;
    StringField17: TStringField;
    StringField18: TStringField;
    StringField19: TStringField;
    StringField20: TStringField;
    IntegerField6: TIntegerField;
    StringField21: TStringField;
    StringField22: TStringField;
    IntegerField7: TIntegerField;
    IntegerField8: TIntegerField;
    IntegerField9: TIntegerField;
    IntegerField10: TIntegerField;
    IntegerField11: TIntegerField;
    IntegerField12: TIntegerField;
    IntegerField13: TIntegerField;
    IntegerField14: TIntegerField;
    BCDField3: TBCDField;
    IntegerField15: TIntegerField;
    BCDField4: TBCDField;
    IntegerField16: TIntegerField;
    StringField23: TStringField;
    BCDField5: TBCDField;
    IntegerField17: TIntegerField;
    IntegerField18: TIntegerField;
    IntegerField19: TIntegerField;
    StringField24: TStringField;
    IntegerField20: TIntegerField;
    IntegerField21: TIntegerField;
    StringField25: TStringField;
    IntegerField22: TIntegerField;
    IntegerField23: TIntegerField;
    StringField26: TStringField;
    IntegerField24: TIntegerField;
    IntegerField25: TIntegerField;
    StringField27: TStringField;
    IntegerField26: TIntegerField;
    IntegerField27: TIntegerField;
    StringField28: TStringField;
    IntegerField28: TIntegerField;
    IntegerField29: TIntegerField;
    StringField29: TStringField;
    cxPageControl1: TcxPageControl;
    cxTabSheet1: TcxTabSheet;
    cxTabSheet2: TcxTabSheet;
    cxPageControl2: TcxPageControl;
    cxTabSheet3: TcxTabSheet;
    cxTabSheet4: TcxTabSheet;
    Panel1: TPanel;
    Edt_Price: TCheckBox;
    RadioButton3: TRadioButton;
    RadioButton4: TRadioButton;
    EDT_PriceDiscountRate: TEdit;
    EDT_PriceFixAmount: TEdit;
    Panel2: TPanel;
    Label10: TLabel;
    EDT_Cash: TCheckBox;
    RadioButton5: TRadioButton;
    RadioButton6: TRadioButton;
    EDT_CashOrderRate: TEdit;
    EDT_CashFixAmount: TEdit;
    Panel3: TPanel;
    Label19: TLabel;
    EDT_Coupon: TCheckBox;
    RadioButton7: TRadioButton;
    RadioButton8: TRadioButton;
    EDT_CouponOrderRate: TEdit;
    EDT_CouponFixAmount: TEdit;
    Panel4: TPanel;
    btnCashCItem: TButton;
    X_CashCItemCodeStr: TBitBtn;
    Edt_CashCItemCodeStr: TEdit;
    Edt_CashCItemCode: TListBox;
    btnGift: TButton;
    EDT_Gift: TCheckBox;
    Panel6: TPanel;
    Button1: TButton;
    EDT_ServiceTypeStr: TEdit;
    Button3: TButton;
    EDT_CustStatusCodeStr: TEdit;
    Button5: TButton;
    EDT_ClassCodeStr: TEdit;
    Button7: TButton;
    EDT_MduIdStr: TEdit;
    Button9: TButton;
    EDT_NodeNoStr: TEdit;
    Button11: TButton;
    EDT_CircuitNoStr: TEdit;
    X_ServiceTypeStr: TBitBtn;
    X_CustStatusCodeStr: TBitBtn;
    X_ClassCodeStr: TBitBtn;
    X_MduIdStr: TBitBtn;
    X_NodeNoStr: TBitBtn;
    X_CircuitNoStr: TBitBtn;
    EDT_ServiceType: TListBox;
    EDT_CustStatusCode: TListBox;
    EDT_ClassCode: TListBox;
    EDT_MduId: TListBox;
    EDT_NodeNo: TListBox;
    EDT_CircuitNo: TListBox;
    Panel7: TPanel;
    btnCitemCodeStr: TButton;
    EDT_CitemCodeStr: TEdit;
    X_CitemCodeStr: TBitBtn;
    EDT_CitemCode: TListBox;
    Panel8: TPanel;
    EDT_Punish: TCheckBox;
    ClientDataSet1DISCOUNTCODESTR: TStringField;
    ClientDataSet1PACKAGENOSTR: TStringField;
    ClientDataSet1BUNDLE: TIntegerField;
    ClientDataSet1BUNDLEMON: TIntegerField;
    ClientDataSet1PENALAMT: TIntegerField;
    ClientDataSet1PENALTYPE: TIntegerField;
    ClientDataSet1EXPIRETYPE: TIntegerField;
    ClientDataSet1EXPIRECUT: TIntegerField;
    ClientDataSet1DEPOSITATTR: TIntegerField;
    ClientDataSet1DEPOSITCODE: TIntegerField;
    ClientDataSet1DEPOSITNAME: TStringField;
    ClientDataSet1DEPOSITAMT: TIntegerField;
    ClientDataSet1CONVERTIBLEPROMOTIONSTR: TStringField;
    ClientDataSet1GIFTORDERPRICEMAX1: TIntegerField;
    ClientDataSet1GIFTKIND1: TIntegerField;
    ClientDataSet1GIFTORDERPRICEMAX2: TIntegerField;
    ClientDataSet1GIFTKIND2: TIntegerField;
    ClientDataSet1GIFTORDERPRICEMAX3: TIntegerField;
    ClientDataSet1GIFTKIND3: TIntegerField;
    ClientDataSet1GIFTORDERPRICEMAX4: TIntegerField;
    ClientDataSet1GIFTKIND4: TIntegerField;
    ClientDataSet1GIFTORDERPRICEMAX5: TIntegerField;
    ClientDataSet1GIFTKIND5: TIntegerField;
    ClientDataSet1GIFTORDERPRICEMAX6: TIntegerField;
    ClientDataSet1GIFTKIND6: TIntegerField;
    Timer1: TTimer;
    btnPackageNoStr: TButton;
    X_PackageNoStr: TBitBtn;
    EDT_PackageNoStr: TEdit;
    EDT_PackageNo: TListBox;
    chkCitemCodeStr: TCheckBox;
    chkPackageNoStr: TCheckBox;
    EDT_DiscountCode: TEdit;
    EDT_DiscountCodeStr: TComboBox;
    chkDiscountCode: TCheckBox;
    cboDiscountCode: TComboBox;
    Panel9: TPanel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    CheckBox5: TCheckBox;
    EDT_BundleMon: TLabeledEdit;
    EDT_PenalType: TComboBox;
    EDT_Expiretype: TComboBox;
    EDT_DepositAttr: TComboBox;
    EDT_DepositCode: TEdit;
    EDT_DepositName: TComboBox;
    ComboBox9: TComboBox;
    EDT_DepositAmt: TEdit;
    Button14: TButton;
    X_CPromotionStr: TBitBtn;
    EDT_CPromotionStr: TEdit;
    EDT_CPromotion: TListBox;
    GroupBox4: TGroupBox;
    cxPageControl3: TcxPageControl;
    cxTabSheet5: TcxTabSheet;
    cxTabSheet6: TcxTabSheet;
    Panel10: TPanel;
    EDT_Note: TMemo;
    Panel11: TPanel;
    EDT_WorkNote: TMemo;
    Panel12: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    EDT_CodeNo: TEdit;
    EDT_OrderCount: TEdit;
    EDT_Description: TEdit;
    EDT_ActStartDate: TMaskEdit;
    EDT_ActStopDate: TMaskEdit;
    EDT_LastOrderDate: TMaskEdit;
    EDT_RefNo: TEdit;
    EDT_StopFlag: TCheckBox;
    EDT_EBTProject: TLabeledEdit;
    EDT_Management: TLabeledEdit;
    EDT_ChanceDays: TLabeledEdit;
    cxTabSheet7: TcxTabSheet;
    Panel13: TPanel;
    CheckBox7: TCheckBox;
    EDT_PenalCode: TEdit;
    EDT_PenalName: TComboBox;
    ComboBox10: TComboBox;
    Label8: TLabel;
    ClientDataSet1PENAL: TIntegerField;
    ClientDataSet1PENALCODE: TIntegerField;
    ClientDataSet1PENALNAME: TStringField;
    ClientDataSet1PENALMON1: TIntegerField;
    ClientDataSet1PENALAMT1: TIntegerField;
    ClientDataSet1PENALMON2: TIntegerField;
    ClientDataSet1PENALAMT2: TIntegerField;
    ClientDataSet1PENALMON3: TIntegerField;
    ClientDataSet1PENALAMT3: TIntegerField;
    ClientDataSet1PENALMON4: TIntegerField;
    ClientDataSet1PENALAMT4: TIntegerField;
    ClientDataSet1PENALMON5: TIntegerField;
    ClientDataSet1PENALAMT5: TIntegerField;
    ClientDataSet1PENALMON6: TIntegerField;
    ClientDataSet1PENALAMT6: TIntegerField;
    ClientDataSet1PENALMON7: TIntegerField;
    ClientDataSet1PENALAMT7: TIntegerField;
    ClientDataSet1PENALMON8: TIntegerField;
    ClientDataSet1PENALAMT8: TIntegerField;
    ClientDataSet1PENALMON9: TIntegerField;
    ClientDataSet1PENALAMT9: TIntegerField;
    ClientDataSet1PENALMON10: TIntegerField;
    ClientDataSet1PENALAMT10: TIntegerField;
    ClientDataSet1PENALMON11: TIntegerField;
    ClientDataSet1PENALAMT11: TIntegerField;
    ClientDataSet1PENALMON12: TIntegerField;
    ClientDataSet1PENALAMT12: TIntegerField;
    btnCopyOther: TButton;
    PenalGrid: TcxGrid;
    gvPenal: TcxGridDBTableView;
    gvPenalCol1: TcxGridDBColumn;
    gvPenalCol2: TcxGridDBColumn;
    gvPenalCol3: TcxGridDBColumn;
    gvPenalCol4: TcxGridDBColumn;
    glPenal: TcxGridLevel;
    dsCD042A: TDataSource;
    CD042A: TClientDataSet;
    CD042AProvider: TDataSetProvider;
    cxStyleRepository1: TcxStyleRepository;
    cxStyle1: TcxStyle;
    ADOQuery3: TADOQuery;
    btnFind: TButton;
    Action5: TAction;
    ClientDataSet1ChooseMulti: TIntegerField;
    EDT_ChooseMulti: TCheckBox;
    btnServCode: TButton;
    X_ServNoStr: TBitBtn;
    EDT_ServCodeStr: TEdit;
    EDT_ServCode: TListBox;
    btnArea: TButton;
    X_AreaNoStr: TBitBtn;
    EDT_AreaNoStr: TEdit;
    EDT_AreaNo: TListBox;
    ClientDataSet1ServCodeStr: TStringField;
    ClientDataSet1AreaCodeStr: TStringField;
    ClientDataSet1BPCodeOrdStr: TStringField;
    btnWorkClass: TButton;
    X_WorkClassStr: TBitBtn;
    EDT_WorkClassStr: TEdit;
    EDT_WorkClass: TListBox;
    btnClctEn: TButton;
    X_ClctEnStr: TBitBtn;
    EDT_ClctEnStr: TEdit;
    EDT_ClctEn: TListBox;
    ClientDataSet1WorkClassStr: TStringField;
    ClientDataSet1EmpNoStr: TStringField;
    lbl1: TLabel;
    lbl2: TLabel;
    EDT_BourgCode: TEdit;
    EDT_BourgName: TComboBox;
    EDTBourgComboBox: TComboBox;
    btnBourgCode: TButton;
    X_BourgStr: TBitBtn;
    EDT_BourgStr: TEdit;
    EDT_Bourg: TListBox;
    ClientDataSet1BourgCodeStr: TStringField;
    EDT_BundleProdCode: TListBox;
    EDT_BundleProdCodeOrd: TListBox;
    chkBundleProdCodeStr: TCheckBox;
    X_BundleProdCodeStr: TBitBtn;
    EDT_BundleProdCodeStr: TEdit;
    btnBundleProdCodeStr: TButton;
    procedure Button10Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Edt_PriceClick(Sender: TObject);
    procedure EDT_CashClick(Sender: TObject);
    procedure EDT_CouponClick(Sender: TObject);
    procedure RadioButton3Click(Sender: TObject);
    procedure RadioButton4Click(Sender: TObject);
    procedure RadioButton5Click(Sender: TObject);
    procedure RadioButton6Click(Sender: TObject);
    procedure RadioButton7Click(Sender: TObject);
    procedure RadioButton8Click(Sender: TObject);
    procedure EDT_OrderCountKeyPress(Sender: TObject; var Key: Char);
    procedure EDT_PriceDiscountRateKeyPress(Sender: TObject;
      var Key: Char);
    procedure EDT_ActStartDateExit(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure EDT_GiftClick(Sender: TObject);
    procedure Button12Click(Sender: TObject);
    procedure btnModifyClick(Sender: TObject);
    procedure X_ServiceTypeStrClick(Sender: TObject);
    procedure Action1Execute(Sender: TObject);
    procedure Action2Execute(Sender: TObject);
    procedure ClientDataSet1AfterScroll(DataSet: TDataSet);
    procedure ClientDataSet1CalcFields(DataSet: TDataSet);
    procedure btnInsertClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure Button9Click(Sender: TObject);
    procedure Button11Click(Sender: TObject);
    procedure btnCitemCodeStrClick(Sender: TObject);
    procedure btnBundleProdCodeStrClick(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure Action3Execute(Sender: TObject);
    procedure Action4Execute(Sender: TObject);
    procedure EDT_ActStopDateExit(Sender: TObject);
    procedure EDT_PriceDiscountRateEnter(Sender: TObject);
    procedure EDT_PriceFixAmountEnter(Sender: TObject);
    procedure EDT_CashOrderRateEnter(Sender: TObject);
    procedure EDT_CashFixAmountEnter(Sender: TObject);
    procedure EDT_CouponOrderRateEnter(Sender: TObject);
    procedure EDT_CouponFixAmountEnter(Sender: TObject);
    procedure EDT_PriceDiscountRateExit(Sender: TObject);
    procedure btnCashCItemClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnGiftClick(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure btnPackageNoStrClick(Sender: TObject);
    procedure chkCitemCodeStrClick(Sender: TObject);
    procedure chkBundleProdCodeStrClick(Sender: TObject);
    procedure chkPackageNoStrClick(Sender: TObject);
    procedure chkDiscountCodeClick(Sender: TObject);
    procedure EDT_DiscountCodeStrSelect(Sender: TObject);
    procedure EDT_DepositAttrSelect(Sender: TObject);
    procedure Button14Click(Sender: TObject);
    procedure CheckBox7Click(Sender: TObject);
    procedure EDT_PenalNameSelect(Sender: TObject);
    procedure btnCopyOtherClick(Sender: TObject);
    procedure EDT_DepositNameSelect(Sender: TObject);
    procedure CD042ANewRecord(DataSet: TDataSet);
    procedure gvPenalCol4PropertiesButtonClick(Sender: TObject;
      AButtonIndex: Integer);
    procedure EDT_BundleMonChange(Sender: TObject);
    procedure btnFindClick(Sender: TObject);
    procedure Action5Execute(Sender: TObject);
    procedure btnAreaClick(Sender: TObject);
    procedure btnServCodeClick(Sender: TObject);
    procedure btnWorkClassClick(Sender: TObject);
    procedure btnClctEnClick(Sender: TObject);
    procedure EDT_WorkClassStrChange(Sender: TObject);
    procedure btnBourgCodeClick(Sender: TObject);
    procedure EDT_BourgNameSelect(Sender: TObject);
    procedure EDT_AreaNoStrChange(Sender: TObject);
  private
    EditMode:string;
    FLockCheck: Boolean;
    FKey: String;
    FGroupId : string;
    FCanInsert : Boolean;
    FCanModify : Boolean;
    FCanPrint : Boolean;
    procedure BrowserMode;
    procedure QueryMode;
    procedure ShowServiceType;
    procedure ShowAreaCode;
    procedure ShowServCode;
    procedure ShowWorkClass;
    procedure ShowClctEn;
    procedure ShowBourgCode;
    procedure ShowCustStatusCode;
    procedure ShowClassCode;
    procedure ShowMduId;
    procedure ShowNodeNo;
    procedure ShowCircuitNo;
    procedure ShowCitemCode;
    procedure ShowBundleProdCode;
    procedure ShowBundleProdCodeOrd;
    procedure ShowCashCitemCode;
    procedure ShowPackageNo;
    procedure ShowCPromotion;
    procedure ClearAllControls;
    function CheckData:Boolean;
    procedure SaveData;
    procedure WMUSER999(var Msg: TMessage); Message WM_USER+999;
    procedure ClearGiftValues;
    function CheckGiftValues: Boolean;
    procedure AddDiscountItem;
    procedure AddDepositItem;
    procedure AddPenalItem;
    procedure LoadCD042A;
    procedure EmptyCD042A;
    procedure ReCalcCD042AItemNo;
    procedure AddBourgItem;
    {}
    function GetServiceType: String;
    function VdCD042A: Boolean;
    function GetUserPriv( const Mid: String ): Boolean;
    procedure GetGropuId;
    {}
  public
    sG_IsSupervisor, sG_OperatorID, sG_Operator, sG_CompID, sG_CompName, sG_DbUserID, sG_DbPasswd, sG_DbAlias, sG_ServiecType : String;
    procedure GetQryData();
  end;

  function StrCount(aStr: String): Integer;
  function GiftCount(aStr: String): Integer; overload;
  function GiftCount(aIndex: Integer): Integer; overload;

var
  Form1: TForm1;

implementation

uses
  cbUtilis, frmMultiSelectU, print_frmcd042, frmGiftU, DateUtils, frmcd0421, frmQryCD042U;

{ ---------------------------------------------------------------------------- }

function StrCount(aStr: String): Integer;
var
  aValue: String;
begin
  Result := 0;
  repeat
    aValue := ExtractValue( aStr );
    Inc( Result );
  until ( aStr = EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

function GiftCount(aStr: String): Integer;
var
  aValue: String;
begin
  aValue := EmptyStr;
  if ByteType( aStr, Length( aStr ) - 1 ) = mbSingleByte then
  begin
    aValue := Copy( aStr, Length( aStr ) - 1, 1 );
  end;
  Result := StrToIntDef( aValue, 0 );
end;

{ ---------------------------------------------------------------------------- }

function GiftCount(aIndex: Integer): Integer;
begin
  Result := aIndex;
end;

{ ---------------------------------------------------------------------------- }

{$R *.dfm}

procedure TForm1.Button10Click(Sender: TObject);
begin
  Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.WMUSER999(var Msg: TMessage);
begin
  Application.Restore;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.BrowserMode;
begin
  ClientDataSet1.Cancel;
  CD042A.Cancel;
  btnFind.Enabled := True;
  btnPrint.Enabled:=not ClientDataSet1.Eof;
  btnInsert.Enabled:=True;
  btnModify.Enabled:=not ClientDataSet1.Eof;
  Button10.Visible:=True;
  DbGrid1.Enabled:=True;
  Button12.Visible:=False;
  btnSave.Enabled:=False;
  {#7222}
  if btnInsert.Enabled then
    btnInsert.Enabled := FCanInsert;
  if btnPrint.Enabled then
    btnPrint.Enabled := FCanPrint;
  if btnModify.Enabled then
    btnModify.Enabled := FCanModify;
  {}
    //7238
//  Panel6.Enabled := False;
//  Panel7.Enabled := False;
  //7238
  chkCitemCodeStr.Enabled := False;
  chkBundleProdCodeStr.Enabled := False;
  chkPackageNoStr.Enabled := False;
  chkDiscountCode.Enabled := False;
  Panel8.Enabled := False;
  Panel9.Enabled := False;
  {}
  Panel1.Enabled := False;
  Panel2.Enabled := False;
  Panel3.Enabled := False;
  Panel4.Enabled := False;
  {}
  Panel10.Enabled := False;
  Panel11.Enabled := False;
  Panel12.Enabled := False;
  Panel13.Enabled := False;
  {}
  PenalGrid.Enabled := False;
  gvPenal.Styles.Content := cxStyle1;
  {}
  if Button10.CanFocus then Button10.SetFocus;
  Edit35.Text:='顯示';
  EditMode:='';
  ClientDataSet1AfterScroll( ClientDataSet1 );
  DbGrid1.Refresh;
  DbGrid1.SetFocus;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.FormShow(Sender: TObject);
begin
  cxPageControl1.ActivePageIndex := 0;
  cxPageControl2.ActivePageIndex := 0;
  cxPageControl3.ActivePageIndex := 0;
  chkCitemCodeStr.Enabled := False;
  chkBundleProdCodeStr.Enabled := False;
  chkPackageNoStr.Enabled := False;
  chkDiscountCode.Enabled := False;
  btnCitemCodeStr.Enabled := False;
  btnBundleProdCodeStr.Enabled := False;
  btnPackageNoStr.Enabled := False;
  QueryMode;

//  Button6.Enabled := True;
//  Timer1.Enabled := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.Edt_PriceClick(Sender: TObject);
var
  I:Integer;
begin
  for I:=0 to Panel1.ControlCount-1 do
    if not (Panel1.Controls[I] is TCheckBox) then
      Panel1.Controls[I].Enabled:=TCheckBox(Sender).Checked;
  if TCheckBox(Sender).Checked then
    RadioButton3.Checked:=True
  else
    begin
      RadioButton3.Checked:=False;
      RadioButton4.Checked:=False;
      EDT_PriceDiscountRate.Text:='';
      EDT_PriceFixAmount.Text:='';
    end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.EDT_CashClick(Sender: TObject);
var
  aIndex:Integer;
begin
  for aIndex := 0 to Panel2.ControlCount-1 do
    if not (Panel2.Controls[aIndex] is TCheckBox) then
      Panel2.Controls[aIndex].Enabled:=TCheckBox(Sender).Checked;
  {}
  if TCheckBox(Sender).Checked then
    RadioButton5.Checked := True
  else
  begin
    RadioButton5.Checked:=False;
    RadioButton6.Checked:=False;
    EDT_CashOrderRate.Text:='';
    EDT_CashFixAmount.Text:='';
    Edt_CashCItemCodeStr.Clear;
    Edt_CashCItemCode.Clear;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.EDT_CouponClick(Sender: TObject);
var
  I:Integer;
begin
  for I:=0 to Panel3.ControlCount-1 do
    if not (Panel3.Controls[I] is TCheckBox) then
      Panel3.Controls[I].Enabled:=TCheckBox(Sender).Checked;

  if TCheckBox(Sender).Checked then
    RadioButton7.Checked:=True
  else
    begin
      RadioButton7.Checked:=False;
      RadioButton8.Checked:=False;
      EDT_CouponOrderRate.Text:='';
      EDT_CouponFixAmount.Text:='';
    end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.RadioButton3Click(Sender: TObject);
begin
  EDT_PriceFixAmount.Text:='';
  if EDT_CashFixAmount.CanFocus then EDT_PriceDiscountRate.SetFocus;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.RadioButton4Click(Sender: TObject);
begin
  EDT_PriceDiscountRate.Text:='';
  if EDT_PriceFixAmount.CanFocus then EDT_PriceFixAmount.SetFocus;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.RadioButton5Click(Sender: TObject);
begin
  EDT_CashFixAmount.Text:='';
  if EDT_CashOrderRate.CanFocus then EDT_CashOrderRate.SetFocus;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.RadioButton6Click(Sender: TObject);
begin
  EDT_CashOrderRate.Text:='';
  if EDT_CashFixAmount.CanFocus then EDT_CashFixAmount.SetFocus;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.RadioButton7Click(Sender: TObject);
begin
  EDT_CouponFixAmount.Text:='';
  if EDT_CouponOrderRate.CanFocus then EDT_CouponOrderRate.SetFocus;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.RadioButton8Click(Sender: TObject);
begin
  EDT_CouponOrderRate.Text:='';
  if EDT_CouponFixAmount.CanFocus then EDT_CouponFixAmount.SetFocus;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.EDT_OrderCountKeyPress(Sender: TObject; var Key: Char);
begin
  if ((Key<#48) or (Key>#57)) and (Key<>#8) then
    Key:=#0;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.EDT_PriceDiscountRateKeyPress(Sender: TObject;
  var Key: Char);
begin
  if ((Key<#48) or (Key>#57)) and (Key<>#8) and (Key<>'.') then
    Key:=#0;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.EDT_ActStartDateExit(Sender: TObject);
var
  aDate, aErrMsg: String;
begin
  aDate := TMaskEdit( Sender ).EditText;
  if not DateTextIsValidEx( aDate ) then
  begin
    aErrMsg := '您輸入的日期格式有誤。';
  end;
  if ( aErrMsg <> EmptyStr ) then
  begin
    WarningMsg( aErrMsg );
    if ( TMaskEdit( Sender ).CanFocus ) then
      TMaskEdit( Sender ).SetFocus;
  end else
  begin
    TMaskEdit( Sender ).EditText := aDate;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.ShowServiceType;
var
  I:Integer;
  WhereString:string;
begin
  EDT_ServiceTypeStr.Text:='';
  WhereString:='';

  for I:=0 to EDT_ServiceType.Items.Count-1 do
    begin
      if I=0 then
        WhereString:=WhereString+'Where '
      else
        WhereString:=WhereString+'Or ';
      WhereString:=WhereString+'(CODENO='''+EDT_ServiceType.Items[I]+''') ';
    end;

  if WhereString<>'' then
    begin
      AdoQuery2.Close;
      AdoQuery2.SQL.Clear;
      AdoQuery2.SQL.Text:='Select DESCRIPTION From '+sG_DbUserID+'.CD046 '+WhereString;
      AdoQuery2.Prepared:=True;
      AdoQuery2.Open;
      while not AdoQuery2.Eof do
        begin
          EDT_ServiceTypeStr.Text:=EDT_ServiceTypeStr.Text+AdoQuery2.FieldByName('DESCRIPTION').AsString;
          AdoQuery2.Next;
          if not AdoQuery2.Eof then
            EDT_ServiceTypeStr.Text:=EDT_ServiceTypeStr.Text+',';
        end;
    end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.ShowCustStatusCode;
var
  I:Integer;
  WhereString:string;
begin
  EDT_CustStatusCodeStr.Text:='';
  WhereString:='';

  for I:=0 to EDT_CustStatusCode.Items.Count-1 do
    begin
      if I=0 then
        WhereString:=WhereString+'Where '
      else
        WhereString:=WhereString+'Or ';
      WhereString:=WhereString+'(CODENO='+EDT_CustStatusCode.Items[I]+') ';
    end;

  if WhereString<>'' then
    begin
      AdoQuery2.Close;
      AdoQuery2.SQL.Clear;
      AdoQuery2.SQL.Text:='Select DESCRIPTION From '+sG_DbUserID+'.CD035 '+WhereString;
      AdoQuery2.Prepared:=True;
      AdoQuery2.Open;
      while not AdoQuery2.Eof do
        begin
          EDT_CustStatusCodeStr.Text:=EDT_CustStatusCodeStr.Text+AdoQuery2.FieldByName('DESCRIPTION').AsString;
          AdoQuery2.Next;
          if not AdoQuery2.Eof then
            EDT_CustStatusCodeStr.Text:=EDT_CustStatusCodeStr.Text+',';
        end;
    end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.ShowClassCode;
var
  I:Integer;
  WhereString:string;
begin
  EDT_ClassCodeStr.Text:='';
  WhereString:='';

  for I:=0 to EDT_ClassCode.Items.Count-1 do
    begin
      if I=0 then
        WhereString:=WhereString+'Where '
      else
        WhereString:=WhereString+'Or ';
      WhereString:=WhereString+'(CODENO='+EDT_ClassCode.Items[I]+') ';
    end;

  if WhereString<>'' then
    begin
      AdoQuery2.Close;
      AdoQuery2.SQL.Clear;
      AdoQuery2.SQL.Text:='Select DESCRIPTION From '+sG_DbUserID+'.CD004 '+WhereString;
      AdoQuery2.Prepared:=True;
      AdoQuery2.Open;
      while not AdoQuery2.Eof do
        begin
          EDT_ClassCodeStr.Text:=EDT_ClassCodeStr.Text+AdoQuery2.FieldByName('DESCRIPTION').AsString;
          AdoQuery2.Next;
          if not AdoQuery2.Eof then
            EDT_ClassCodeStr.Text:=EDT_ClassCodeStr.Text+',';
        end;
    end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.ShowMduId;
var
  I:Integer;
  WhereString:string;
begin
  EDT_MduIdStr.Text:='';
  WhereString:='';

  for I:=0 to EDT_MduId.Items.Count-1 do
    begin
      if I=0 then
        WhereString:=WhereString+'Where ('
      else
        WhereString:=WhereString+'Or ';
      WhereString:=WhereString+'(MduId='''+EDT_MduId.Items[I]+''') ';
    end;

  if WhereString<>'' then
    begin
      AdoQuery2.Close;
      AdoQuery2.SQL.Clear;
      AdoQuery2.SQL.Text:='Select Name From '+sG_DbUserID+'.SO017 '+WhereString+') And (COMPCODE='+sG_CompID+')';
      AdoQuery2.Prepared:=True;
      AdoQuery2.Open;
      while not AdoQuery2.Eof do
        begin
          EDT_MduIdStr.Text:=EDT_MduIdStr.Text+AdoQuery2.FieldByName('Name').AsString;
          AdoQuery2.Next;
          if not AdoQuery2.Eof then
            EDT_MduIdStr.Text:=EDT_MduIdStr.Text+',';
        end;
    end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.ShowNodeNo;
var
  I:Integer;
  WhereString:string;
begin
  EDT_NodeNoStr.Text:='';
  WhereString:='';

  for I:=0 to EDT_NodeNo.Items.Count-1 do
    begin
      if I=0 then
        WhereString:=WhereString+'Where ('
      else
        WhereString:=WhereString+'Or ';
      WhereString:=WhereString+'(CODENO='''+EDT_NodeNo.Items[I]+''') ';
    end;

  if WhereString<>'' then
    begin
      AdoQuery2.Close;
      AdoQuery2.SQL.Clear;
      AdoQuery2.SQL.Text:='Select CODENO From '+sG_DbUserID+'.CD047 '+WhereString+') And (COMPCODE='+sG_CompID+')';
      AdoQuery2.Prepared:=True;
      AdoQuery2.Open;
      while not AdoQuery2.Eof do
        begin
          EDT_NodeNoStr.Text:=EDT_NodeNoStr.Text+AdoQuery2.FieldByName('CODENO').AsString;
          AdoQuery2.Next;
          if not AdoQuery2.Eof then
            EDT_NodeNoStr.Text:=EDT_NodeNoStr.Text+',';
        end;
    end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.FormCreate(Sender: TObject);
begin
  if ParamCount<>9 then
  begin
    ErrorMsg( '傳入之參數個數錯誤!' );
    Application.Terminate;
  end else
  begin
    //N 1 Howard 1 觀昇有限 gicmis may mis
    sG_IsSupervisor := Uppercase(ParamStr(1));
    sG_OperatorID := ParamStr(2);
    sG_Operator := ParamStr(3);
    sG_CompID := ParamStr(4);
    sG_CompName := ParamStr(5);
    sG_DbUserID := ParamStr(6);
    sG_DbPasswd := ParamStr(7);
    sG_DbAlias := ParamStr(8);
    sG_ServiecType := ParamStr(9);
  end;
  ADOConnection1.ConnectionString:='Provider=MSDAORA.1;Password='+sG_DbPasswd+';User ID='+sG_DbUserID+';Data Source='+sG_DbAlias;
  {#7222}
  GetGropuId;
  FCanInsert := GetUserPriv( 'SO64N01' );
  FCanModify := GetUserPriv( 'SO64N02' );
  FCanPrint := GetUserPriv( 'SO64N05' );
  TcxButtonEditProperties( gvPenalCol4.Properties ).Buttons[0].Width :=
    gvPenalCol4.Width - 3;
  cxSetResourceString( @scxGridNoDataInfoText, '無資料可顯示' );
  
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.ShowCircuitNo;
var
  I:Integer;
  WhereString:string;
begin
  EDT_CircuitNoStr.Text:='';
  WhereString:='';

  if EDT_CircuitNo.Items.Count>0 then
    for I:=0 to EDT_CircuitNo.Items.Count-1 do
      begin
        if I=0 then
          WhereString:=WhereString+'Where '
        else
          WhereString:=WhereString+'Or ';
        WhereString:=WhereString+'(CircuitNo='''+EDT_CircuitNo.Items[I]+''') ';
      end;

  if WhereString<>'' then
    begin
      AdoQuery2.Close;
      AdoQuery2.SQL.Clear;
      AdoQuery2.SQL.Text:='Select DISTINCT CircuitNo From '+sG_DbUserID+'.SO016 '+WhereString;
      AdoQuery2.Prepared:=True;
      AdoQuery2.Open;
      while not AdoQuery2.Eof do
        begin
          EDT_CircuitNoStr.Text:=EDT_CircuitNoStr.Text+AdoQuery2.FieldByName('CircuitNo').AsString;
          AdoQuery2.Next;
          if not AdoQuery2.Eof then
            EDT_CircuitNoStr.Text:=EDT_CircuitNoStr.Text+',';
        end;
    end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.ShowCitemCode;
var
  I:Integer;
  WhereString:string;
begin
  EDT_CitemCodeStr.Text:='';
  WhereString:='';

  for I:=0 to EDT_CitemCode.Items.Count-1 do
    begin
      if I=0 then
        WhereString:=WhereString+'Where '
      else
        WhereString:=WhereString+'Or ';
      WhereString:=WhereString+'(CODENO='''+EDT_CitemCode.Items[I]+''') ';
    end;

  if WhereString<>'' then
    begin
      AdoQuery2.Close;
      AdoQuery2.SQL.Clear;
      AdoQuery2.SQL.Text:='Select Description From '+sG_DbUserID+'.CD019 '+WhereString;
      AdoQuery2.Prepared:=True;
      AdoQuery2.Open;
      while not AdoQuery2.Eof do
        begin
          EDT_CitemCodeStr.Text:=EDT_CitemCodeStr.Text+AdoQuery2.FieldByName('Description').AsString;
          AdoQuery2.Next;
          if not AdoQuery2.Eof then
            EDT_CitemCodeStr.Text:=EDT_CitemCodeStr.Text+',';
        end;
    end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.ShowBundleProdCode;
var
  I:Integer;
  WhereString:string;
begin
  EDT_BundleProdCodeStr.Text:='';
  WhereString:='';

  for I:=0 to EDT_BundleProdCode.Items.Count-1 do
    begin
      if I=0 then
        WhereString:=WhereString+'Where '
      else
        WhereString:=WhereString+'Or ';
      WhereString:=WhereString+'(CODENO='''+EDT_BundleProdCode.Items[I]+''') ';
    end;

  if WhereString<>'' then
    begin
      AdoQuery2.Close;
      AdoQuery2.SQL.Clear;
      AdoQuery2.SQL.Text:='Select Description From '+sG_DbUserID+'.CD078 '+WhereString;
      AdoQuery2.Prepared:=True;
      AdoQuery2.Open;
      while not AdoQuery2.Eof do
        begin
          EDT_BundleProdCodeStr.Text:=EDT_BundleProdCodeStr.Text+AdoQuery2.FieldByName('Description').AsString;
          AdoQuery2.Next;
          if not AdoQuery2.Eof then
            EDT_BundleProdCodeStr.Text:=EDT_BundleProdCodeStr.Text+',';
        end;
    end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.ShowCashCitemCode;
var
  I:Integer;
  WhereString:string;
begin
  Edt_CashCItemCodeStr.Text:='';
  WhereString:='';
  for I := 0 to Edt_CashCItemCode.Items.Count-1 do
  begin
    if I=0 then
      WhereString:=WhereString+'Where '
    else
      WhereString:= WhereString+ 'Or ';
    WhereString := WhereString+'(CODENO='+Edt_CashCItemCode.Items[I]+') ';
  end;

  if WhereString<>'' then
  begin
    AdoQuery2.Close;
    AdoQuery2.SQL.Clear;
    AdoQuery2.SQL.Text:='Select DESCRIPTION From '+sG_DbUserID+'.CD019 '+WhereString;
    AdoQuery2.Prepared:=True;
    AdoQuery2.Open;
    while not AdoQuery2.Eof do
    begin
      Edt_CashCItemCodeStr.Text:= Edt_CashCItemCodeStr.Text +
        AdoQuery2.FieldByName('DESCRIPTION').AsString;
      AdoQuery2.Next;
      if not AdoQuery2.Eof then
        Edt_CashCItemCodeStr.Text := Edt_CashCItemCodeStr.Text + ',';
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.ShowPackageNo;
var
  I:Integer;
  WhereString:string;
begin
  EDT_PackageNoStr.Text:='';
  WhereString:='';
  for I := 0 to EDT_PackageNo.Items.Count-1 do
  begin
    if I=0 then
      WhereString:=WhereString+'Where '
    else
      WhereString:= WhereString+ 'Or ';
    WhereString := WhereString+'(CODENO='+ '''' + EDT_PackageNo.Items[I] + '''' +') ';
  end;

  if WhereString<>'' then
  begin
    AdoQuery2.Close;
    AdoQuery2.SQL.Clear;
    AdoQuery2.SQL.Text:='Select DESCRIPTION From '+sG_DbUserID+'.CD027 '+WhereString;
    AdoQuery2.Prepared:=True;
    AdoQuery2.Open;
    while not AdoQuery2.Eof do
    begin
      EDT_PackageNoStr.Text:= EDT_PackageNoStr.Text +
        AdoQuery2.FieldByName('DESCRIPTION').AsString;
      AdoQuery2.Next;
      if not AdoQuery2.Eof then
        EDT_PackageNoStr.Text := EDT_PackageNoStr.Text + ',';
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.ShowCPromotion;
var
  I:Integer;
  WhereString:string;
begin
  EDT_CPromotionStr.Text:='';
  WhereString:='';
  for I := 0 to EDT_CPromotion.Items.Count-1 do
  begin
    if I=0 then
      WhereString:=WhereString+'Where '
    else
      WhereString:= WhereString+ 'Or ';
    WhereString := WhereString+'(CODENO='+ '''' + EDT_CPromotion.Items[I] + '''' +') ';
  end;

  if WhereString<>'' then
  begin
    AdoQuery2.Close;
    AdoQuery2.SQL.Clear;
    AdoQuery2.SQL.Text:='Select DESCRIPTION From '+sG_DbUserID+'.CD042 '+WhereString;
    AdoQuery2.Prepared:=True;
    AdoQuery2.Open;
    while not AdoQuery2.Eof do
    begin
      EDT_CPromotionStr.Text:= EDT_CPromotionStr.Text +
        AdoQuery2.FieldByName('DESCRIPTION').AsString;
      AdoQuery2.Next;
      if not AdoQuery2.Eof then
        EDT_CPromotionStr.Text := EDT_CPromotionStr.Text + ',';
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.EDT_GiftClick(Sender: TObject);
var
  I:Integer;
begin
  for I:= 0 to Panel4.ControlCount-1 do
  begin
    Panel4.Controls[I].Visible := False;
    if ( Panel4.Controls[I] =  btnGift ) then
    begin
      Panel4.Controls[I].Visible := True;
      Panel4.Controls[I].Enabled := TCheckBox(Sender).Checked;
    end else
    if ( Panel4.Controls[I] = Sender ) then
    begin
      Panel4.Controls[I].Visible := True;
    end;
  end;
  if not EDT_Gift.Checked then
    ClearGiftValues;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.Button12Click(Sender: TObject);
begin
  BrowserMode;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.btnModifyClick(Sender: TObject);
begin
  EditMode:='Fix';
  ClientDataSet1.Edit;
  Button12.Visible:=True;
  Button10.Visible:=False;
  btnFind.Enabled :=False;
  btnSave.Enabled:=True;
  btnPrint.Enabled:=False;
  btnInsert.Enabled:=False;
  btnModify.Enabled:=False;
  EDT_CodeNo.Enabled:=False;
  DbGrid1.Enabled:=False;
  DbGrid1.Refresh;
  Edit35.Text:='修改';
  {}
  Panel6.Enabled := True;
  Panel7.Enabled := True;
  //7238
  chkCitemCodeStr.Enabled := true;
  chkBundleProdCodeStr.Enabled := true;
  chkPackageNoStr.Enabled := true;
  chkDiscountCode.Enabled := true;

  Panel8.Enabled := True;
  Panel9.Enabled := True;
  {}
  Panel1.Enabled := True;
  Panel2.Enabled := True;
  Panel3.Enabled := True;
  Panel4.Enabled := True;
  {}
  Panel10.Enabled := True;
  Panel11.Enabled := True;
  Panel12.Enabled := True;
  Panel13.Enabled := True;
  {}
  if EDT_Description.CanFocus then EDT_Description.SetFocus;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.X_ServiceTypeStrClick(Sender: TObject);
var
  EditNames,ListBoxName:string;
begin
  if EditMode='' then Exit;

  EditNames:='EDT_'+Copy(TBitbtn(Sender).Name,3,Length(TBitbtn(Sender).Name)-2);
  ListBoxName:='EDT_'+Copy(TBitbtn(Sender).Name,3,Length(TBitbtn(Sender).Name)-5);
  TEdit(FindComponent(EditNames)).Text:='';
  TListBox(FindComponent(ListBoxName)).Clear;

end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.Action1Execute(Sender: TObject);
begin
  if btnModify.Enabled then
    btnModifyClick(Sender);
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.btnCashCItemClick(Sender: TObject);
var
  I:Integer;
  SelectedValues,String1:string;
begin
  SelectedValues:='';

  for I:=0 to Edt_CashCItemCode.Items.Count-1 do
    begin
      SelectedValues:=SelectedValues+Edt_CashCItemCode.Items[I];
      if I<>Edt_CashCItemCode.Items.Count-1 then
        SelectedValues:=SelectedValues+',';
    end;
  frmMultiSelect:=TfrmMultiSelect.Create(Application);
  try
    frmMultiSelect.Connection:=ADOConnection1;
    frmMultiSelect.KeyFields:='CODENO';
    frmMultiSelect.KeyValues:=SelectedValues;
    frmMultiSelect.DisplayFields:='CODENO,現金回饋項目代碼,DESCRIPTION,現金回饋項目名稱';
    frmMultiSelect.ResultFields:='CODENO';
    frmMultiSelect.SQL.Text:=
      'SELECT CODENO,DESCRIPTION From '+sG_DbUserID+'.CD019 ' +
      ' WHERE STOPFLAG = 0 AND SIGN = ''-''   ' +  
      ' Order By CODENO';
    frmMultiSelect.QuotedString:='';
    if frmMultiSelect.ShowModal=mrOk then
      begin
        SelectedValues:=frmMultiSelect.SelectedValue;
        Edt_CashCItemCode.Clear;
        while SelectedValues<>'' do
          begin
            String1:=ExtractValue(SelectedValues);
            Edt_CashCItemCode.Items.Add(String1);
          end;
        ShowCashCitemCode;
      end;
  finally
    frmMultiSelect.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.btnCopyOtherClick(Sender: TObject);
begin
  Form2 := TForm2.Create( nil );
  try
    Form2.ShowModal;
  finally
    Form2.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.btnSaveClick(Sender: TObject);
begin
  Screen.Cursor := crSQLWait;
  try
    SaveData;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.Action2Execute(Sender: TObject);
begin
  if btnSave.Enabled then
  begin
    btnSave.Click;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.chkCitemCodeStrClick(Sender: TObject);
var
  aIndex: Integer;
begin
  for aIndex:=0 to Panel1.ControlCount-1 do
  begin
    if Panel1.Controls[aIndex] is TCheckBox then
      TCheckBox(Panel1.Controls[aIndex]).Checked := False
    else if Panel1.Controls[aIndex] is TRadioButton then
      TRadioButton(Panel1.Controls[aIndex]).Checked:=False
    else if Panel1.Controls[aIndex] is TEdit then
      TEdit(Panel1.Controls[aIndex]).Text:=''
    else if Panel1.Controls[aIndex] is TComboBox then
      TComboBox(Panel1.Controls[aIndex]).Items.Clear
    else if Panel1.Controls[aIndex] is TListBox then
      TListBox(Panel1.Controls[aIndex]).Items.Clear;
    if Panel1.Controls[aIndex] is TCheckBox then
      Panel1.Controls[aIndex].Enabled:= chkCitemCodeStr.Checked
    else
      Panel1.Controls[aIndex].Enabled:=False;
  end;
  if not chkCitemCodeStr.Checked then X_ServiceTypeStrClick( X_CitemCodeStr );
  btnCitemCodeStr.Enabled := chkCitemCodeStr.Checked;
  X_CitemCodeStr.Enabled := chkCitemCodeStr.Checked;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.chkBundleProdCodeStrClick(Sender: TObject);
begin
  if not chkBundleProdCodeStr.Checked then X_ServiceTypeStrClick( X_BundleProdCodeStr );
  btnBundleProdCodeStr.Enabled := chkBundleProdCodeStr.Checked;
  X_BundleProdCodeStr.Enabled := chkBundleProdCodeStr.Checked;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.chkPackageNoStrClick(Sender: TObject);
begin
  if not chkPackageNoStr.Checked then X_ServiceTypeStrClick( X_PackageNoStr );
  btnPackageNoStr.Enabled := chkPackageNoStr.Checked;
  X_PackageNoStr.Enabled := chkPackageNoStr.Checked;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.chkDiscountCodeClick(Sender: TObject);
begin
  if chkDiscountCode.Checked then
  begin
    AddDiscountItem;
    EDT_DiscountCodeStr.ItemIndex:=0;
    cboDiscountCode.ItemIndex:=0;
    EDT_DiscountCode.Text := cboDiscountCode.Items[cboDiscountCode.ItemIndex];
  end else
  begin
    EDT_DiscountCode.Clear;
    EDT_DiscountCodeStr.ItemIndex := -1;
    cboDiscountCode.Clear;
  end;
  EDT_DiscountCode.Enabled := chkDiscountCode.Checked;
  EDT_DiscountCodeStr.Enabled := chkDiscountCode.Checked;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.CheckBox7Click(Sender: TObject);
begin
  if CheckBox7.Checked then
  begin
    AddPenalItem;
    EDT_PenalName.ItemIndex := 0;
    ComboBox10.ItemIndex := 0;
    EDT_PenalCode.Text := ComboBox10.Items[ComboBox10.ItemIndex];
    PenalGrid.Enabled := True;
    gvPenal.Styles.Content := nil;
    if CD042A.IsEmpty then CD042A.Append;
  end else
  begin
    EDT_PenalCode.Clear;
    EDT_PenalName.ItemIndex := -1;
    ComboBox10.Clear;
    PenalGrid.Enabled := False;
    gvPenal.Styles.Content := cxStyle1;
    EmptyCD042A;
  end;
  EDT_PenalCode.Enabled := CheckBox7.Checked;
  EDT_PenalName.Enabled := CheckBox7.Checked;
end;

{ ---------------------------------------------------------------------------- }

function TForm1.CheckData:Boolean;
begin
  Result:=False;
  if EDT_CodeNo.Text='' then
    begin
      WarningMsg('促銷活動代碼需有值');
      if EDT_CodeNo.CanFocus then EDT_CodeNo.SetFocus;
      Exit;
    end
  else if EDT_Description.Text='' then
    begin
      WarningMsg('促銷活動名稱需有值');
      if EDT_Description.CanFocus then EDT_Description.SetFocus;
      Exit;
    end
  else if TrimChar( EDT_ActStartDate.Text, ['/',#32] ) = EmptyStr then
    begin
      WarningMsg('促銷活動期間需有值');
      if EDT_ActStartDate.CanFocus then EDT_ActStartDate.SetFocus;
      Exit;
    end
  else if TrimChar( EDT_ActStopDate.Text, ['/',#32] ) = EmptyStr then
    begin
      WarningMsg('促銷活動期間需有值');
      if EDT_ActStopDate.CanFocus then EDT_ActStopDate.SetFocus;
      Exit;
    end
  else if TrimChar( EDT_LastOrderDate.Text, ['/',#32] ) = EmptyStr then
    begin
      WarningMsg('最後接單日期需有值');
      if EDT_LastOrderDate.CanFocus then EDT_LastOrderDate.SetFocus;
      Exit;
    end
  else if EDT_OrderCount.Text='' then
    begin
      WarningMsg('每一客戶可使用次數需有值');
      if EDT_OrderCount.CanFocus then EDT_OrderCount.SetFocus;
      Exit;
    end
  else if chkCitemCodeStr.Checked and (EDT_CitemCode.Items.Count<1) then
    begin
      cxPageControl1.ActivePageIndex := 1;
      WarningMsg('促銷產品：特定產品需有值');
      if btnCitemCodeStr.CanFocus then btnCitemCodeStr.SetFocus;
      Exit;
    end
  else if chkBundleProdCodeStr.Checked and (EDT_BundleProdCode.Items.Count<1) then
    begin
      cxPageControl1.ActivePageIndex := 1;
      WarningMsg('促銷產品：優惠組合需有值');
      if btnBundleProdCodeStr.CanFocus then btnBundleProdCodeStr.SetFocus;
      Exit;
    end
  else if chkPackageNoStr.Checked and (EDT_PackageNo.Items.Count<1) then
    begin
      cxPageControl1.ActivePageIndex := 1;
      WarningMsg('促銷產品：套餐需有值');
      if btnPackageNoStr.CanFocus then btnPackageNoStr.SetFocus;
      Exit;
    end
  else if chkDiscountCode.Checked and ( EDT_DiscountCode.Text = EmptyStr ) then
    begin
      cxPageControl1.ActivePageIndex := 1;
      WarningMsg('促銷產品：優惠辦法需有值');
      if EDT_DiscountCodeStr.CanFocus then EDT_DiscountCodeStr.SetFocus;
      Exit;
    end
  else if Edt_Price.Checked and RadioButton3.Checked and (EDT_PriceDiscountRate.Text='') then
    begin
      cxPageControl2.ActivePageIndex := 0;
      WarningMsg('價格優惠：折扣需有值');
      if EDT_PriceDiscountRate.CanFocus then EDT_PriceDiscountRate.SetFocus;
      Exit;
    end
  else if Edt_Price.Checked and RadioButton4.Checked and (EDT_PriceFixAmount.Text='') then
    begin
      cxPageControl2.ActivePageIndex := 0;
      WarningMsg('價格優惠：固定金額需有值');
      if EDT_PriceFixAmount.CanFocus then EDT_PriceFixAmount.SetFocus;
      Exit;
    end
  else if EDT_Cash.Checked and RadioButton5.Checked and (EDT_CashOrderRate.Text='') then
    begin
      cxPageControl2.ActivePageIndex := 0;
      WarningMsg('現金回饋：訂購金額％需有值');
      if EDT_CashOrderRate.CanFocus then EDT_CashOrderRate.SetFocus;
      Exit;
    end
  else if EDT_Cash.Checked and RadioButton6.Checked and (EDT_CashFixAmount.Text='') then
    begin
      cxPageControl2.ActivePageIndex := 0;
      WarningMsg('現金回饋：固定金額需有值');
      if EDT_CashFixAmount.CanFocus then EDT_CashFixAmount.SetFocus;
      Exit;
    end
  else if EDT_Cash.Checked and (Edt_CashCItemCode.Items.Count <= 0) then
    begin
      cxPageControl2.ActivePageIndex := 0;
      WarningMsg('現金回饋：現金回饋項目需有值');
      if btnCashCItem.CanFocus then btnCashCItem.SetFocus;
      Exit;
    end
  else if EDT_Coupon.Checked and RadioButton7.Checked and (EDT_CouponOrderRate.Text='') then
    begin
      cxPageControl2.ActivePageIndex := 0;
      WarningMsg('折價卷回饋：訂購金額％需有值');
      if EDT_CouponOrderRate.CanFocus then EDT_CouponOrderRate.SetFocus;
      Exit;
    end
  else if EDT_Coupon.Checked and RadioButton8.Checked and (EDT_CouponFixAmount.Text='') then
    begin
      cxPageControl2.ActivePageIndex := 0;
      WarningMsg('折價卷回饋：固定金額需有值');
      if EDT_CouponFixAmount.CanFocus then EDT_CouponFixAmount.SetFocus;
      Exit;
    end
  else if ( EDT_DepositAttr.ItemIndex in [1..2] ) and ( EDT_DepositCode.Text = EmptyStr ) then
    begin
      cxPageControl2.ActivePageIndex := 1;
      WarningMsg('保證金屬性已設定, 請指定保證金收費項目' );
      if EDT_DepositName.CanFocus then EDT_DepositName.SetFocus;
      Exit;
    end
  else if ( EDT_DepositAttr.ItemIndex in [1..2] ) and ( EDT_DepositCode.Text = EmptyStr ) then
    begin
      if ( EDT_DepositAmt.Text = EmptyStr ) then
      begin
        cxPageControl2.ActivePageIndex := 1;
        WarningMsg('保證金屬性已設定, 請指定保證金金額' );
        if EDT_DepositAmt.CanFocus then EDT_DepositAmt.SetFocus;
        Exit;
      end;
    end
  else if ( EDT_Gift.Checked ) and ( not CheckGiftValues ) then
    begin
      Exit;
    end
  else if ( CheckBox7.Checked ) and ( ( EDT_PenalCode.Text = EmptyStr ) or ( EDT_PenalName.Text = EmptyStr ) ) then
    begin
      cxPageControl2.ActivePageIndex := 2;
      WarningMsg('已指定計算違約金, 請指定違約金收費項目' );
      if EDT_PenalName.CanFocus then EDT_PenalName.SetFocus;
      Exit;
    end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.ClearAllControls;
var
  I:Integer;
begin
  FLockCheck := True;
  try
    for I:=0 to ComponentCount-1 do
    begin
      if ( Components[I].Name = 'EDT_PenalType' ) or
         ( Components[I].Name = 'EDT_Expiretype' ) or
         ( Components[I].Name = 'EDT_DepositAttr' ) then
       Continue;
      if (Components[I] is TEdit) and (Components[I].Name<>'Edit35') then
        TEdit(Components[I]).Text:=''
      else if Components[I] is TMaskEdit then    //#4311在寫的時候順便發現的 沒有清空日期 By Kin 2009/02/05
        TMaskEdit(Components[I]).Clear
      else if Components[I] is TLabeledEdit then //#4311在寫的時候順便發現的 沒有清空 By Kin 2009/02/05
        (Components[I] as TLabeledEdit).Clear
      else if Components[I] is TCheckBox then
        TCheckBox(Components[I]).Checked:=False
      else if Components[I] is TRadioButton then
        TRadioButton(Components[I]).Checked:=False
      else if Components[I] is TMemo then
        TMemo(Components[I]).Clear
      else if Components[I] is TListBox then
        TListBox(Components[I]).Clear
      else if Components[I] is TComboBox then
        TComboBox(Components[I]).Clear;
    end;
  finally
    FLockCheck := False;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.ClientDataSet1AfterScroll(DataSet: TDataSet);
var
  I:Integer;
  DataString,String1:string;
begin
  I := ClientDataSet1.RecNo;
  if I<1 then
    I:=0;
  Panel5.Caption:=IntToStr(I)+' ／ '+IntToStr(ClientDataSet1.RecordCount);

  ClearAllControls;

  FKey := ClientDataSet1.FieldByName('CodeNo').AsString;

  LoadCD042A;
  
  EDT_CodeNo.Text := FKey;
  EDT_Description.Text:=ClientDataSet1.FieldByName('Description').AsString;
  EDT_OrderCount.Text:=ClientDataSet1.FieldByName('OrderCount').AsString;
  EDT_ActStartDate.Text:=ClientDataSet1.FieldByName('ActStartDate2').AsString;
  EDT_ActStopDate.Text:=ClientDataSet1.FieldByName('ActStopDate2').AsString;
  EDT_LastOrderDate.Text:=ClientDataSet1.FieldByName('LastOrderDate2').AsString;
  EDT_RefNo.Text:=ClientDataSet1.FieldByName('RefNo').AsString;
  EDT_StopFlag.Checked:=ClientDataSet1.FieldByName('StopFlag').AsInteger=1;
  EDT_EBTProject.Text := ClientDataSet1.FieldByName( 'EBTProject' ).AsString;
  EDT_ChanceDays.Text := ClientDataSet1.FieldByName( 'ChanceDays' ).AsString;
  EDT_Management.Text := ClientDataSet1.FieldByName( 'Management' ).AsString;

  EDT_Note.Clear;
  EDT_Note.Text:=ClientDataSet1.FieldByName('Note').AsString;
  EDT_WorkNote.Clear;
  EDT_WorkNote.Text:=ClientDataSet1.FieldByName('WorkNote').AsString;

  DataString:=ClientDataSet1.FieldByName('ServiceTypeStr').AsString;
  if ClientDataSet1.FieldByName('ChooseMulti').AsInteger=1 then
    EDT_ChooseMulti.Checked := True
  else
    EDT_ChooseMulti.Checked := False;
  while DataString<>'' do
    begin
      String1:=ExtractValue(DataString);
      EDT_ServiceType.Items.Add(String1);
    end;
  ShowServiceType;

  DataString:=ClientDataSet1.FieldByName('CustStatusCodeStr').AsString;
  while DataString<>'' do
    begin
      String1:=ExtractValue(DataString);
      EDT_CustStatusCode.Items.Add(String1);
    end;
  ShowCustStatusCode;

  DataString:=ClientDataSet1.FieldByName('ClassCodeStr').AsString;
  while DataString<>'' do
    begin
      String1:=ExtractValue(DataString);
      EDT_ClassCode.Items.Add(String1);
    end;
  ShowClassCode;

  DataString:=ClientDataSet1.FieldByName('MduIdStr').AsString;
  while DataString<>'' do
    begin
      String1:=ExtractValue(DataString);
      EDT_MduId.Items.Add(String1);
    end;
    ShowMduId;
  {#6472 增加行政區 By Kin 2013/05/13}
  DataString:=ClientDataSet1.FieldByName('AreaCodeStr').AsString;
  while DataString<>'' do
  begin
    String1:=ExtractValue(DataString);
    EDT_AreaNo.Items.Add(String1);
  end;
  ShowAreaCode;
  {#7132}
  DataString:=ClientDataSet1.FieldByName('BourgCodeStr').AsString;
  while DataString<>'' do
  begin
    String1:=ExtractValue(DataString);
    EDT_Bourg.Items.Add(String1);
  end;
  ShowBourgCode;
   {
   AddBourgItem;
   EDT_BourgCode.Text := ClientDataSet1.FieldByName( 'BourgCodeStr' ).AsString;
   for I :=0 to EDTBourgComboBox.Items.Count-1 do
   begin
    if EDTBourgComboBox.Items[I] = EDT_BourgCode.Text then
    begin
      EDT_BourgName.ItemIndex := I;
      EDTBourgComboBox.ItemIndex := I;
      Break;
    end;
   end;
   }
   {#6472 增加服務區 By Kin 2013/05/13}
   DataString:=ClientDataSet1.FieldByName('ServCodeStr').AsString;
    while DataString<>'' do
    begin
      String1:=ExtractValue(DataString);
      EDT_ServCode.Items.Add(String1);
    end;
    ShowServCode;


  DataString:=ClientDataSet1.FieldByName('NodeNoStr').AsString;
  while DataString<>'' do
    begin
      String1:=ExtractValue(DataString);
      EDT_NodeNo.Items.Add(String1);
    end;
  ShowNodeNo;

  DataString:=ClientDataSet1.FieldByName('CircuitNoStr').AsString;
  while DataString<>'' do
    begin
      String1:=ExtractValue(DataString);
      EDT_CircuitNo.Items.Add(String1);
    end;
  ShowCircuitNo;

  {#6798 增加工作類別 By Kin 2014/06/11}
  DataString := StringReplace(ClientDataSet1.FieldByName('WorkClassStr').AsString,'''','', [rfReplaceAll]);
  while DataString <> '' do
  begin
    String1 := ExtractValue(DataString);
    EDT_WorkClass.Items.Add( String1 );
  end;
  ShowWorkClass;

  {#6798 增加適用人員 By Kin 2014/06/11}
  DataString := StringReplace(ClientDataSet1.FieldByName('EmpNoStr').AsString,'''','', [rfReplaceAll]);
  while DataString <> '' do
  begin
    String1 := ExtractValue(DataString);
    EDT_ClctEn.Items.Add( String1 );
  end;
  ShowClctEn;

  chkCitemCodeStr.Checked:= ClientDataSet1.FieldByName('CitemCodeStr').AsString <> EmptyStr;
  chkCitemCodeStrClick( chkCitemCodeStr );
//  CheckBox1Click( chkCitemCodeStr );
  if ( chkCitemCodeStr.Checked ) then
  begin
    DataString:=ClientDataSet1.FieldByName('CitemCodeStr').AsString;
    while DataString<>'' do
    begin
      String1:=ExtractValue(DataString);
      EDT_CitemCode.Items.Add(String1);
    end;
    ShowCitemCode;

  end;

  chkBundleProdCodeStr.Checked:= ClientDataSet1.FieldByName('BPCodeStr').AsString <> EmptyStr;
  chkBundleProdCodeStrClick( chkBundleProdCodeStr );
//  CheckBox2Click( CheckBox2 );
  if ( chkBundleProdCodeStr.Checked ) then
  begin
    DataString:=ClientDataSet1.FieldByName('BPCodeStr').AsString;
    while DataString<>'' do
    begin
      String1:=ExtractValue(DataString);
      EDT_BundleProdCode.Items.Add(String1);
    end;
    ShowBundleProdCode;
    ShowBundleProdCodeOrd;
  end;




  chkPackageNoStr.Checked:= ClientDataSet1.FieldByName('PackageNoStr').AsString <> EmptyStr;
  chkPackageNoStrClick( chkPackageNoStr );
//  CheckBox3Click( CheckBox3 );
  if ( chkPackageNoStr.Checked )  then
  begin
    DataString := ClientDataSet1.FieldByName( 'PackageNoStr' ).AsString;
    while ( DataString <> EmptyStr ) do
    begin
      String1 := ExtractValue( DataString );
      EDT_PackageNo.Items.Add( String1 );
    end;
    ShowPackageNo;
  end;

  chkDiscountCode.Checked:= ClientDataSet1.FieldByName('DiscountCodeStr').AsString <> EmptyStr;
  chkDiscountCodeClick( chkDiscountCode );
//  CheckBox4Click( CheckBox4 );
  EDT_DiscountCode.Text := ClientDataSet1.FieldByName('DiscountCodeStr').AsString;
  for I:=0 to cboDiscountCode.Items.Count-1 do
  begin
    if cboDiscountCode.Items[I] = EDT_DiscountCode.Text then
    begin
      EDT_DiscountCodeStr.ItemIndex := I;
      cboDiscountCode.ItemIndex := I;
    end;
  end;



  EDT_Punish.Checked:=ClientDataSet1.FieldByName('Punish').AsInteger=1;
  EDT_Price.Checked:=ClientDataSet1.FieldByName('Price').AsInteger=1;
  EDT_Cash.Checked:=ClientDataSet1.FieldByName('Cash').AsInteger=1;
  EDT_Coupon.Checked:=ClientDataSet1.FieldByName('Coupon').AsInteger=1;
  EDT_Gift.Checked:=ClientDataSet1.FieldByName('Gift').AsInteger=1;
  btnGift.Enabled := EDT_Gift.Checked;

  {}
  btnCashCItem.Enabled := ( EDT_Cash.Checked );
  X_CashCItemCodeStr.Enabled := ( EDT_Cash.Checked );

  {}

  RadioButton3.Checked:=(Panel1.Enabled) and (ClientDataSet1.FieldByName('PriceDiscountType').AsInteger=1);
  RadioButton4.Checked:=ClientDataSet1.FieldByName('PriceDiscountType').AsInteger=2;
  EDT_PriceDiscountRate.Text:=ClientDataSet1.FieldByName('PriceDiscountRate').AsString;
  EDT_PriceFixAmount.Text:=ClientDataSet1.FieldByName('PriceFixAmount').AsString;

  RadioButton5.Checked:=ClientDataSet1.FieldByName('CashDiscountType').AsInteger=1;
  RadioButton6.Checked:=ClientDataSet1.FieldByName('CashDiscountType').AsInteger=2;
  EDT_CashOrderRate.Text:=ClientDataSet1.FieldByName('CashOrderRate2').AsString;
  EDT_CashFixAmount.Text:=ClientDataSet1.FieldByName('CashFixAmount').AsString;

  DataString :=ClientDataSet1.FieldByName('CashCitemCodeStr').AsString;
  while DataString<>'' do
  begin
    String1 := ExtractValue(DataString);
    Edt_CashCItemCode.Items.Add(String1);
  end;
  ShowCashCitemCode;


  RadioButton7.Checked:=ClientDataSet1.FieldByName('CouponDiscountType').AsInteger=1;
  RadioButton8.Checked:=ClientDataSet1.FieldByName('CouponDiscountType').AsInteger=2;
  EDT_CouponOrderRate.Text:=ClientDataSet1.FieldByName('CouponOrderRate2').AsString;
  EDT_CouponFixAmount.Text:=ClientDataSet1.FieldByName('CouponFixAmount').AsString;


  {}

  CheckBox5.Checked := ClientDataSet1.FieldByName('Bundle').AsString = '1';
  EDT_BundleMon.Text := ClientDataSet1.FieldByName('BundleMon').AsString;


  EDT_PenalType.ItemIndex := 0;
  if ClientDataSet1.FieldByName('PenalType').AsInteger in [1..2] then
    EDT_PenalType.ItemIndex := ClientDataSet1.FieldByName('PenalType').AsInteger;


  EDT_Expiretype.ItemIndex := 0;
  if ClientDataSet1.FieldByName('Expiretype').AsInteger in [1..2] then
    EDT_Expiretype.ItemIndex := ClientDataSet1.FieldByName('Expiretype').AsInteger;


  EDT_DepositAttr.ItemIndex := 0;
  EDT_DepositCode.Text := EmptyStr;
  EDT_DepositName.ItemIndex := -1;
  ComboBox9.Items.Clear;
  EDT_DepositAmt.Text := EmptyStr;
  if ClientDataSet1.FieldByName('DepositAttr').AsInteger in [1..2] then
  begin
    EDT_DepositAttr.ItemIndex := ClientDataSet1.FieldByName('DepositAttr').AsInteger;
    EDT_DepositAttrSelect( EDT_DepositAttr );
    EDT_DepositCode.Text := ClientDataSet1.FieldByName('DepositCode').AsString;
    for I := 0 to ComboBox9.Items.Count-1 do
    begin
      if ComboBox9.Items[I] = EDT_DepositCode.Text then
      begin
        EDT_DepositName.ItemIndex := I;
        ComboBox9.ItemIndex := I;
      end;
    end;
    EDT_DepositAmt.Text := ClientDataSet1.FieldByName( 'DepositAmt' ).AsString;
  end;

  DataString := ClientDataSet1.FieldByName('ConvertiblePromotionStr').AsString;
  while DataString <>'' do
  begin
    String1:=ExtractValue(DataString);
    EDT_CPromotion.Items.Add(String1);
  end;
  ShowCPromotion;
  

  CheckBox7.Checked := ( ClientDataSet1.FieldByName( 'Penal' ).AsInteger = 1 );
  EDT_PenalCode.Text := ClientDataSet1.FieldByName( 'PenalCode' ).AsString;
  for I :=0 to ComboBox10.Items.Count-1 do
  begin
    if ComboBox10.Items[I] = EDT_PenalCode.Text then
    begin
      EDT_PenalName.ItemIndex := I;
      ComboBox10.ItemIndex := I;
    end;
  end;
  {}
  if EditMode = '' then
    DbGrid1.SetFocus;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.ClientDataSet1CalcFields(DataSet: TDataSet);
var
  PercentValue:Integer;
  String1:string;
begin

  String1:=TClientDataSet(DataSet).FieldByName('ActStartDate').AsString;
  if String1 <> '' then
    begin
      String1 := FormatDateTime( 'yyyy/mm/dd', DataSet.FieldByName('ActStartDate').AsDateTime );
      TClientDataSet(DataSet).FieldByName('ActStartDate2').AsString:=String1;
    end;

  String1:=TClientDataSet(DataSet).FieldByName('ActStopDate').AsString;
  if String1<>'' then
    begin
      String1 := FormatDateTime( 'yyyy/mm/dd', DataSet.FieldByName('ActStopDate').AsDateTime );
      TClientDataSet(DataSet).FieldByName('ActStopDate2').AsString:=String1;
    end;

  String1:=TClientDataSet(DataSet).FieldByName('LastOrderDate').AsString;
  if String1<>'' then
    begin
      String1 := FormatDateTime( 'yyyy/mm/dd', DataSet.FieldByName('LastOrderDate').AsDateTime );    
      TClientDataSet(DataSet).FieldByName('LastOrderDate2').AsString:=String1;
    end;


  String1:=TClientDataSet(DataSet).FieldByName('CashOrderRate').AsString;
  if String1<>'' then
    begin
      PercentValue:=Trunc(StrtoFloat(String1)*100);
      TClientDataSet(DataSet).FieldByName('CashOrderRate2').AsString:=IntToStr(PercentValue);
    end;

  String1:=TClientDataSet(DataSet).FieldByName('CouponOrderRate').AsString;
  if String1<>'' then
    begin
      PercentValue:=Trunc(StrtoFloat(String1)*100);
      TClientDataSet(DataSet).FieldByName('CouponOrderRate2').AsString:=IntToStr(PercentValue);
    end;
  
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.btnInsertClick(Sender: TObject);
begin
  { #4437 沒有資料也要可以新增 By Kin 2009/03/19 }
  if ( not ClientDataSet1.Active ) then
  begin
    ClientDataSet1.CommandText := Format(
              ' SELECT * FROM %S.CD042 ' +
              '  WHERE 1=0 ', [sG_DbUserID] );
    ClientDataSet1.AfterScroll := nil;
    try
      ClientDataSet1.Open;
    finally
      ClientDataSet1.AfterScroll := ClientDataSet1AfterScroll;
    end;

  end;
  EditMode:='Add';
  ClientDataSet1.Append;
  Button12.Visible:=True;
  Button10.Visible:=False;
  btnFind.Enabled := False;
  btnSave.Enabled:=True;
  btnPrint.Enabled:=False;
  btnInsert.Enabled:=False;
  btnModify.Enabled:=False;
  EDT_CodeNo.Enabled:=True;
  DbGrid1.Enabled:=False;
  DbGrid1.Refresh;
  {}
  Panel6.Enabled := True;
  Panel7.Enabled := True;
  Panel8.Enabled := True;
  Panel9.Enabled := True;
  {}
  Panel1.Enabled := True;
  Panel2.Enabled := True;
  Panel3.Enabled := True;
  Panel4.Enabled := True;
  {}
  Panel10.Enabled := True;
  Panel11.Enabled := True;
  Panel12.Enabled := True;
  Panel13.Enabled := True;
  {}
  //#7322 to fix cannot modify citem and prodcode by Kin 2016/10/12 
  chkCitemCodeStr.Enabled := true;
  chkBundleProdCodeStr.Enabled := true;
  chkPackageNoStr.Enabled := true;
  chkDiscountCode.Enabled := true;
  {}
  Edit35.Text:='新增';
  if EDT_CodeNo.CanFocus then EDT_CodeNo.SetFocus;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.Button1Click(Sender: TObject);
var
  I:Integer;
  SelectedValues,String1:string;
begin
  SelectedValues:='';

  for I:=0 to EDT_ServiceType.Items.Count-1 do
    begin
      SelectedValues:=SelectedValues+EDT_ServiceType.Items[I];
      if I<>EDT_ServiceType.Items.Count-1 then
        SelectedValues:=SelectedValues+',';
    end;
  frmMultiSelect:=TfrmMultiSelect.Create(Application);
  try
    frmMultiSelect.CanSelect := EditMode <> '';
    frmMultiSelect.Connection:=ADOConnection1;
    frmMultiSelect.KeyFields:='CODENO';
    frmMultiSelect.KeyValues:=SelectedValues;
    frmMultiSelect.DisplayFields:='CODENO,服務類別代碼,DESCRIPTION,服務類別名稱';
    frmMultiSelect.ResultFields:='CODENO';
    frmMultiSelect.SQL.Text:='SELECT CODENO,DESCRIPTION From '+sG_DbUserID+'.CD046 Order By CODENO';
    frmMultiSelect.QuotedString:='';
    if frmMultiSelect.ShowModal=mrOk then
      begin
        SelectedValues:=frmMultiSelect.SelectedValue;
        EDT_ServiceType.Clear;
        while SelectedValues<>'' do
          begin
            String1:=ExtractValue(SelectedValues);
            EDT_ServiceType.Items.Add(String1);
          end;
        ShowServiceType;
      end;
  finally
    frmMultiSelect.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.Button3Click(Sender: TObject);
var
  I:Integer;
  SelectedValues,String1:string;
begin
  SelectedValues:='';

  for I:=0 to EDT_CustStatusCode.Items.Count-1 do
    begin
      SelectedValues:=SelectedValues+EDT_CustStatusCode.Items[I];
      if I<>EDT_CustStatusCode.Items.Count-1 then
        SelectedValues:=SelectedValues+',';
    end;
  frmMultiSelect:=TfrmMultiSelect.Create(Application);
  try
    frmMultiSelect.CanSelect := EditMode <> '';
    frmMultiSelect.Connection:=ADOConnection1;
    frmMultiSelect.KeyFields:='CODENO';
    frmMultiSelect.KeyValues:=SelectedValues;
    frmMultiSelect.DisplayFields:='CODENO,客戶狀態代碼,DESCRIPTION,客戶狀態名稱';
    frmMultiSelect.ResultFields:='CODENO';
    frmMultiSelect.SQL.Text:='SELECT CODENO, DESCRIPTION From '+sG_DbUserID+'.CD035 Order By CODENO';
    frmMultiSelect.QuotedString:='';
    if frmMultiSelect.ShowModal=mrOk then
      begin
        SelectedValues:=frmMultiSelect.SelectedValue;
        EDT_CustStatusCode.Clear;
        while SelectedValues<>'' do
          begin
            String1:=ExtractValue(SelectedValues);
            EDT_CustStatusCode.Items.Add(String1);
          end;
        ShowCustStatusCode;
      end;
  finally
    frmMultiSelect.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.Button5Click(Sender: TObject);
var
  I:Integer;
  SelectedValues,String1:string;
begin
  SelectedValues:='';

  for I:=0 to EDT_ClassCode.Items.Count-1 do
    begin
      SelectedValues:=SelectedValues+EDT_ClassCode.Items[I];
      if I<>EDT_ClassCode.Items.Count-1 then
        SelectedValues:=SelectedValues+',';
    end;
  frmMultiSelect:=TfrmMultiSelect.Create(Application);
  try
    frmMultiSelect.CanSelect := EditMode <> '';
    frmMultiSelect.Connection:=ADOConnection1;
    frmMultiSelect.KeyFields:='CODENO';
    frmMultiSelect.KeyValues:=SelectedValues;
    frmMultiSelect.DisplayFields:='CODENO,客戶類別代碼,DESCRIPTION,客戶類別名稱';
    frmMultiSelect.ResultFields:='DESCRIPTION';
    frmMultiSelect.SQL.Text:='SELECT CODENO, DESCRIPTION From '+sG_DbUserID+'.CD004 Where StopFlag=0 '+
                             'Order By CODENO';
    frmMultiSelect.QuotedString:='';
    if frmMultiSelect.ShowModal=mrOk then
      begin
        SelectedValues:=frmMultiSelect.SelectedValue;
        EDT_ClassCode.Clear;
        while SelectedValues<>'' do
          begin
            String1:=ExtractValue(SelectedValues);
            EDT_ClassCode.Items.Add(String1);
          end;
        ShowClassCode;
      end;
  finally
    frmMultiSelect.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.Button7Click(Sender: TObject);
var
  I:Integer;
  SelectedValues,String1:string;
begin
  SelectedValues:='';

  for I:=0 to EDT_MduId.Items.Count-1 do
    begin
      SelectedValues:=SelectedValues+EDT_MduId.Items[I];
      if I<>EDT_MduId.Items.Count-1 then
        SelectedValues:=SelectedValues+',';
    end;
  frmMultiSelect:=TfrmMultiSelect.Create(Application);
  try
    frmMultiSelect.CanSelect := EditMode <> '';
    frmMultiSelect.Connection:=ADOConnection1;
    frmMultiSelect.KeyFields:='mduid';
    frmMultiSelect.KeyValues:=SelectedValues;
    frmMultiSelect.DisplayFields:='mduid,大樓編號,Name,大樓名稱';
    frmMultiSelect.ResultFields:='mduid';
    frmMultiSelect.SQL.Text:='SELECT mduid, Name From '+sG_DbUserID+
                             '.SO017 Where (COMPCODE='+sG_CompID+') Order By MDUID';
    frmMultiSelect.QuotedString:='';
    if frmMultiSelect.ShowModal=mrOk then
      begin
        SelectedValues:=frmMultiSelect.SelectedValue;
        EDT_MduId.Clear;
        while SelectedValues<>'' do
          begin
            String1:=ExtractValue(SelectedValues);
            EDT_MduId.Items.Add(String1);
          end;
        ShowMduId;
      end;
  finally
    frmMultiSelect.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.Button9Click(Sender: TObject);
var
  I:Integer;
  SelectedValues,String1,ServiceTypes:string;
begin
  if EDT_ServiceType.Items.Count<1 then
    begin
      WarningMsg('請先指定服務類別');
      Exit;
    end;

  SelectedValues:='';
  ServiceTypes:='';

  for I:=0 to EDT_ServiceType.Items.Count-1 do
    begin
      ServiceTypes:=ServiceTypes+''''+EDT_ServiceType.Items[I]+'''';
      if I<>EDT_ServiceType.Items.Count-1 then
        ServiceTypes:=ServiceTypes+',';
    end;

  if ServiceTypes='' then
    ServiceTypes:='''''';

  for I:=0 to EDT_NodeNo.Items.Count-1 do
    begin
      SelectedValues:=SelectedValues+EDT_NodeNo.Items[I];
      if I<>EDT_NodeNo.Items.Count-1 then
        SelectedValues:=SelectedValues+',';
    end;
  frmMultiSelect:=TfrmMultiSelect.Create(Application);
  try
    frmMultiSelect.CanSelect := EditMode <> '';
    frmMultiSelect.Connection:=ADOConnection1;
    frmMultiSelect.KeyFields:='CODENO';
    frmMultiSelect.KeyValues:=SelectedValues;
    frmMultiSelect.DisplayFields:='CODENO,NodeNo';
    frmMultiSelect.ResultFields:='CODENO';
    frmMultiSelect.SQL.Text:='SELECT CODENO From '+sG_DbUserID+
                             '.CD047 Where (COMPCODE='+sG_CompID+
                             ') And (SERVICETYPE In ('+ServiceTypes+')) Order By CODENO';
    frmMultiSelect.QuotedString:='';
    if frmMultiSelect.ShowModal=mrOk then
      begin
        SelectedValues:=frmMultiSelect.SelectedValue;
        EDT_NodeNo.Clear;
        while SelectedValues<>'' do
          begin
            String1:=ExtractValue(SelectedValues);
            EDT_NodeNo.Items.Add(String1);
          end;
        ShowNodeNo;
      end;
  finally
    frmMultiSelect.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.Button11Click(Sender: TObject);
var
  I:Integer;
  SelectedValues,String1:string;
begin
  SelectedValues:='';

  for I:=0 to EDT_CircuitNo.Items.Count-1 do
    begin
      SelectedValues:=SelectedValues+EDT_CircuitNo.Items[I];
      if I<>EDT_CircuitNo.Items.Count-1 then
        SelectedValues:=SelectedValues+',';
    end;
  frmMultiSelect:=TfrmMultiSelect.Create(Application);
  try
    frmMultiSelect.CanSelect := EditMode <> '';
    frmMultiSelect.Connection:=ADOConnection1;
    frmMultiSelect.KeyFields:='CIRCUITNO';
    frmMultiSelect.KeyValues:= SelectedValues;
    frmMultiSelect.DisplayFields:='CircuitNo,網路編號';
    frmMultiSelect.ResultFields:='CIRCUITNO';
    frmMultiSelect.SQL.Text:='SELECT DISTINCT COMPCODE, CIRCUITNO FROM '+sG_DbUserID+
                             '.SO016 Where (COMPCODE='+sG_CompID+') And (CircuitNo Is Not Null) '+
                             'Order By CIRCUITNO';
    frmMultiSelect.QuotedString:='';
    if frmMultiSelect.ShowModal=mrOk then
      begin
        SelectedValues:=frmMultiSelect.SelectedValue;
        EDT_CircuitNo.Clear;
        while SelectedValues<>'' do
          begin
            String1:=ExtractValue(SelectedValues);
            EDT_CircuitNo.Items.Add(String1);
          end;
        ShowCircuitNo;
      end;
  finally
    frmMultiSelect.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.btnCitemCodeStrClick(Sender: TObject);
var
  I:Integer;
  SelectedValues,String1:string;
begin
  SelectedValues:='';

  for I:=0 to EDT_CitemCode.Items.Count-1 do
    begin
      SelectedValues:=SelectedValues+EDT_CitemCode.Items[I];
      if I<>EDT_CitemCode.Items.Count-1 then
        SelectedValues:=SelectedValues+',';
    end;
  frmMultiSelect:=TfrmMultiSelect.Create(Application);
  try
    frmMultiSelect.CanSelect := EditMode <> '';
    frmMultiSelect.Connection:=ADOConnection1;
    frmMultiSelect.KeyFields:='CODENO';
    frmMultiSelect.KeyValues:=SelectedValues;
    frmMultiSelect.DisplayFields:='CODENO,產品編號,Description,產品名稱';
    frmMultiSelect.ResultFields:='CODENO';
    frmMultiSelect.SQL.Text:='SELECT CODENO,Description From '+sG_DbUserID+
                             '.CD019 Where StopFlag=0 Order By CODENO';
    frmMultiSelect.QuotedString:='';
    if frmMultiSelect.ShowModal=mrOk then
      begin
        SelectedValues:=frmMultiSelect.SelectedValue;
        EDT_CitemCode.Clear;
        while SelectedValues<>'' do
          begin
            String1:=ExtractValue(SelectedValues);
            EDT_CitemCode.Items.Add(String1);
          end;
        ShowCitemCode;
      end;
  finally
    frmMultiSelect.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.btnBundleProdCodeStrClick(Sender: TObject);
var
  I:Integer;
  SelectedValues,SelectOrds,String1,String2:string;
begin
  SelectedValues:='';
  SelectOrds := '';
  for I:=0 to EDT_BundleProdCode.Items.Count-1 do
    begin
      SelectedValues:=SelectedValues+EDT_BundleProdCode.Items[I];
      if I<>EDT_BundleProdCode.Items.Count-1 then
        SelectedValues:=SelectedValues+',';
      if EDT_BundleProdCodeOrd.Count > I then
      begin
        SelectOrds :=SelectOrds + EDT_BundleProdCodeOrd.Items[I];
        if I<>EDT_BundleProdCode.Items.Count-1 then
          SelectOrds:=SelectOrds+',';
      end;

    end;
  frmMultiSelect:=TfrmMultiSelect.Create(Application);
  try
    frmMultiSelect.CanSelect := EditMode <> '';
    frmMultiSelect.Connection:=ADOConnection1;
    frmMultiSelect.KeyFields:='CODENO';
    frmMultiSelect.KeyValues:=SelectedValues;
    frmMultiSelect.KeyValueOrd := SelectOrds;
    frmMultiSelect.DisplayFields:='CODENO,產品編號,Description,產品名稱';
    frmMultiSelect.ResultFields:='CODENO';
    frmMultiSelect.SQL.Text:='SELECT CODENO,Description From '+sG_DbUserID+
                             '.CD078 Where StopFlag=0 Order BY CODENO';
    frmMultiSelect.IsUseOrd := True;
    frmMultiSelect.QuotedString:='';
    if frmMultiSelect.ShowModal=mrOk then
      begin
        SelectedValues:=frmMultiSelect.SelectedValue;
        SelectOrds := frmMultiSelect.SelectOrd;
        EDT_BundleProdCode.Clear;
        EDT_BundleProdCodeOrd.Clear;
        while SelectedValues<>'' do
          begin
            String1:=ExtractValue(SelectedValues);
            String2:=ExtractValue(SelectOrds);
            EDT_BundleProdCode.Items.Add(String1);
            EDT_BundleProdCodeOrd.Items.Add( String2 );
          end;

        ShowBundleProdCode;
      end;
  finally
    frmMultiSelect.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.Timer1Timer(Sender: TObject);
begin
  Timer1.Enabled := False;
  Screen.Cursor := crSQLWait;
  try
    ClientDataSet1.CommandText := Format(
      ' SELECT * FROM %S.CD042 ' +
      '  WHERE ( DISCOUNTCODE IS NULL ) AND ( PRODUCTTYPE IS NOT NULL ) ' +
      '  ORDER BY CODENO', [sG_DbUserID] );
    ClientDataSet1.AfterScroll := nil;
    try
      ClientDataSet1.Open;
    finally
      ClientDataSet1.AfterScroll := ClientDataSet1AfterScroll;
    end;
//    BrowserMode;
    QueryMode;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.btnPrintClick(Sender: TObject);
begin
  btnPrint.Enabled:=False;
  try
    ClientDataSet2.Data:=ClientDataSet1.Data;
    QuickReport1:=TQuickReport1.Create(Application);
    with QuickReport1 do
      begin
        QR_CompName.Caption:=sG_CompName;
        QR_CompName.Left:=(Width-QR_CompName.Width) div 2;
        QR_FormName.Left:=(Width-QR_FormName.Width) div 2;
        QR_User.Caption:=QR_User.Caption+sG_Operator;
        Preview;
        Refresh;
        Free;
      end;
  finally
    btnPrint.Enabled:=True;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.Action3Execute(Sender: TObject);
begin
  if btnInsert.Enabled then
    btnInsertClick(Sender);
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.Action4Execute(Sender: TObject);
begin
  if btnPrint.Enabled then
    btnPrintClick(Sender);
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.EDT_ActStopDateExit(Sender: TObject);
var
  aDate, aErrMsg: String;
begin
  aDate := TMaskEdit( Sender ).EditText;
  if not DateTextIsValidEx( aDate ) then
  begin
    aErrMsg := '您輸入的日期格式有誤。';
  end;
  if ( aErrMsg <> EmptyStr ) then
  begin
    WarningMsg( aErrMsg );
    if TMaskEdit( Sender ).CanFocus then
      TMaskEdit( Sender ).SetFocus;
  end else
  begin
    TMaskEdit( Sender ).EditText := aDate;
  end;
end;

procedure TForm1.EDT_BundleMonChange(Sender: TObject);
begin

end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.EDT_PriceDiscountRateEnter(Sender: TObject);
begin
  RadioButton3.Checked:=True;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.EDT_PriceFixAmountEnter(Sender: TObject);
begin
  RadioButton4.Checked:=True;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.EDT_CashOrderRateEnter(Sender: TObject);
begin
  RadioButton5.Checked:=True;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.EDT_CashFixAmountEnter(Sender: TObject);
begin
  RadioButton6.Checked:=True;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.EDT_CouponOrderRateEnter(Sender: TObject);
begin
  RadioButton7.Checked:=True;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.EDT_CouponFixAmountEnter(Sender: TObject);
begin
  RadioButton8.Checked:=True;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.EDT_PriceDiscountRateExit(Sender: TObject);
var
  String1:string;
begin
  String1:=EDT_PriceDiscountRate.Text;
  if (String1<>'') and (StrToFloat(String1)>0.99) then
    begin
      WarningMsg('折扣不能大於 0.99');
      if EDT_PriceDiscountRate.CanFocus then
        EDT_PriceDiscountRate.SetFocus;
    end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.SaveData;
var
  I:Integer;
  String1:string;
begin
  if CheckData then
  begin
    if EditMode='Add' then
    begin
      AdoQuery2.Close;
      AdoQuery2.SQL.Clear;
      AdoQuery2.SQL.Text:='Select CODENO From '+sG_DbUserID+'.CD042 Where CODENO='+EDT_CodeNo.Text;
      AdoQuery2.Prepared:=True;
      AdoQuery2.Open;
      if not AdoQuery2.Eof then
      begin
        WarningMsg('促銷活動代碼重覆');
        if EDT_CodeNo.CanFocus then EDT_CodeNo.SetFocus;
        Exit;
      end;
    end;

    AdoQuery2.Close;
    AdoQuery2.SQL.Clear;
    AdoQuery2.SQL.Text:= Format(
      ' SELECT CODENO, DESCRIPTION      ' +
      '   FROM %s.CD042                 ' +
      '  WHERE ( CODENO <> %s )         ' +
      '    AND ( DESCRIPTION = ''%s'' ) ' +
      '    AND ( STOPFLAG = 0 )         ',
      [sG_DbUserID, EDT_CodeNo.Text, EDT_Description.Text] );
    AdoQuery2.Prepared := True;
    AdoQuery2.Open;
    if not AdoQuery2.Eof then
    begin
      AdoQuery2.Close;
      WarningMsg( '促銷活動名稱重覆' );
      if EDT_Description.CanFocus then EDT_Description.SetFocus;
      Exit;
    end;
    AdoQuery2.Close;
    if EditMode <> 'Add' then
      ClientDataSet1.Refresh;
    ClientDataSet1.Edit;
    ClientDataSet1.FieldByName('CodeNo').AsString:=EDT_CodeNo.Text;
    ClientDataSet1.FieldByName('Description').AsString:=EDT_Description.Text;
    ClientDataSet1.FieldByName('OrderCount').AsString:=EDT_OrderCount.Text;

    String1:=EDT_ActStartDate.Text;
    ClientDataSet1.FieldByName('ActStartDate').AsString:=String1;

    String1:=EDT_ActStopDate.Text;
    ClientDataSet1.FieldByName('ActStopDate').AsString:=String1;

    String1:=EDT_LastOrderDate.Text;
    ClientDataSet1.FieldByName('LastOrderDate').AsString:=String1;

    ClientDataSet1.FieldByName('RefNo').AsString:=EDT_RefNo.Text;
    if EDT_StopFlag.Checked then
      ClientDataSet1.FieldByName('StopFlag').AsString:='1'
    else
      ClientDataSet1.FieldByName('StopFlag').AsString:='0';

    ClientDataSet1.FieldByName('UpdTime').AsDateTime:=Now;
    ClientDataSet1.FieldByName('UpdEn').AsString:=sG_Operator;
    ClientDataSet1.FieldByName('CompCode').AsString:=sG_CompID;

    ClientDataSet1.FieldByName('Note').AsString:=EDT_Note.Text;
    ClientDataSet1.FieldByName('WorkNote').AsString:=EDT_WorkNote.Text;

    ClientDataSet1.FieldByName( 'EBTProject' ).AsString := EDT_EBTProject.Text;
    ClientDataSet1.FieldByName( 'ChanceDays' ).AsString := EDT_ChanceDays.Text;
    ClientDataSet1.FieldByName( 'Management' ).AsString := EDT_Management.Text;
    //#5266 多增加 可選多優惠組合 欄位 By Kin 2009/08/06
    if EDT_ChooseMulti.Checked then
      ClientDataSet1.FieldByName( 'ChooseMulti' ).AsInteger :=1
    else
      ClientDataSet1.FieldByName( 'ChooseMulti' ).AsInteger :=0;


    String1:='';
    for I:=0 to EDT_ServiceType.Items.Count-1 do
    begin
      String1:=String1+EDT_ServiceType.Items[I];
      if I<>EDT_ServiceType.Items.Count-1 then
        String1:=String1+',';
    end;
    ClientDataSet1.FieldByName('ServiceTypeStr').AsString:=String1;

    String1:='';
    for I:=0 to EDT_CustStatusCode.Items.Count-1 do
    begin
      String1:=String1+EDT_CustStatusCode.Items[I];
      if I<>EDT_CustStatusCode.Items.Count-1 then
        String1:=String1+',';
    end;
    ClientDataSet1.FieldByName('CustStatusCodeStr').AsString:=String1;

    String1:='';
    for I:=0 to EDT_ClassCode.Items.Count-1 do
    begin
      String1:=String1+EDT_ClassCode.Items[I];
      if I<>EDT_ClassCode.Items.Count-1 then
        String1:=String1+',';
    end;
    ClientDataSet1.FieldByName('ClassCodeStr').AsString:=String1;

    String1:='';
    for I:=0 to EDT_MduId.Items.Count-1 do
    begin
      String1:=String1+EDT_MduId.Items[I];
      if I<>EDT_MduId.Items.Count-1 then
        String1:=String1+',';
    end;
    ClientDataSet1.FieldByName('MduIdStr').AsString:=String1;

    String1:='';
    for I:=0 to EDT_NodeNo.Items.Count-1 do
    begin
      String1:=String1+EDT_NodeNo.Items[I];
      if I<>EDT_NodeNo.Items.Count-1 then
        String1:=String1+',';
    end;
    ClientDataSet1.FieldByName('NodeNoStr').AsString:=String1;

    String1:='';
    for I:=0 to EDT_CircuitNo.Items.Count-1 do
    begin
      String1:=String1+EDT_CircuitNo.Items[I];
      if I<>EDT_CircuitNo.Items.Count-1 then
        String1:=String1+',';
    end;
    ClientDataSet1.FieldByName('CircuitNoStr').AsString:=String1;

    { 因為原本 ProductType 存的內容值是
      0.優待辦法
      1.特定產品
      2.優惠組合
      原本是只能有1種, 2007/10/24 ( KC 的訂單規格 )現在改成 checkBox 多選,
      變成此欄位內容沒有做用, 所以把它存成 3 }
    ClientDataSet1.FieldByName('ProductType').AsString:= '3';

    { 適用之特定產品 }
    String1:='';
    for I:=0 to EDT_CitemCode.Items.Count-1 do
    begin
      String1:=String1+EDT_CitemCode.Items[I];
      if I<>EDT_CitemCode.Items.Count-1 then
        String1:=String1+',';
     end;
    ClientDataSet1.FieldByName('CitemCodeStr').AsString:=String1;


    { 適用之優惠組合 }
    String1:='';
    for I:=0 to EDT_BundleProdCode.Items.Count-1 do
      begin
        String1:=String1+EDT_BundleProdCode.Items[I];
        if I<>EDT_BundleProdCode.Items.Count-1 then
          String1:=String1+',';
      end;
    ClientDataSet1.FieldByName('BPCodeStr').AsString:=String1;

    { #6798 工作類別 By Kin 2014/06/11 }
    String1:='';
    for I:=0 to EDT_WorkClass.Items.Count-1 do
      begin
        String1:=String1+EDT_WorkClass.Items[I];
        if I<>EDT_WorkClass.Items.Count-1 then
          String1:=String1+',';
        {
        if String1 = '' then
          String1 := '''' +  EDT_WorkClass.Items[I] + ''''
        else
          String1:=String1+ '''' + EDT_WorkClass.Items[I] + '''';
        if I<>EDT_WorkClass.Items.Count-1 then
          String1:=String1+',';
        }
      end;
    ClientDataSet1.FieldByName('WorkClassStr').AsString:=String1;

    { #6798 適用人員 By Kin 2014/06/11 }
    String1:='';
    for I:=0 to EDT_ClctEn.Items.Count-1 do
      begin
        String1:=String1+EDT_ClctEn.Items[I];
        if I<>EDT_ClctEn.Items.Count-1 then
          String1:=String1+',';
        {
        if String1 = '' then
          String1 := '''' +  EDT_ClctEn.Items[I] + ''''
        else
          String1:=String1+ '''' + EDT_ClctEn.Items[I] + '''';
        if I<>EDT_ClctEn.Items.Count-1 then
          String1:=String1+',';
        }
      end;
    ClientDataSet1.FieldByName('EmpNoStr').AsString:=String1;
    {#7132 Add cd042.BourgCodeStr field By Kin 2015/11/02}
    String1:= EmptyStr;
    for I:=0 to EDT_Bourg.Items.Count-1 do
    begin
      String1 := ( String1 + EDT_Bourg.Items[I] );
      if ( I <> EDT_Bourg.Items.Count - 1 ) then
        String1 := ( String1 + ',' );
     end;
    ClientDataSet1.FieldByName('BourgCodeStr').AsString := String1;

    { 適用之套餐 }
    String1:= EmptyStr;
    for I:=0 to EDT_PackageNo.Items.Count-1 do
    begin
      String1 := ( String1 + EDT_PackageNo.Items[I] );
      if ( I <> EDT_PackageNo.Items.Count - 1 ) then
        String1 := ( String1 + ',' );
     end;
    ClientDataSet1.FieldByName('PackageNoStr').AsString := String1;

    {#6472增加行政區 By Kin 2013/05/13}
    String1 := EmptyStr;
    for i:=0 to EDT_AreaNo.Items.Count-1 do
    begin
      String1 := ( String1 + EDT_AreaNo.Items[I] );
      if ( I <> EDT_AreaNo.Items.Count - 1 ) then
        String1 := ( String1 + ',' );
    end;
    ClientDataSet1.FieldByName( 'AreaCodeStr' ).AsString := String1;

    {#6472增加服務區 By Kin 2013/05/13}
    String1 := EmptyStr;
    for i:=0 to EDT_ServCode.Items.Count-1 do
    begin
      String1 := ( String1 + EDT_ServCode.Items[I] );
      if ( I <> EDT_ServCode.Items.Count - 1 ) then
        String1 := ( String1 + ',' );
    end;
    ClientDataSet1.FieldByName( 'ServCodeStr' ).AsString := String1;

    {#6465 增加優惠組合順序 By Kin 2013/05/13}
    String1 := EmptyStr;
    for i:=0 to EDT_BundleProdCodeOrd.Items.Count-1 do
    begin
      String1 := ( String1 + EDT_BundleProdCodeOrd.Items[I] );
      if ( I <> EDT_BundleProdCodeOrd.Items.Count - 1 ) then
        String1 := ( String1 + ',' );
    end;
    ClientDataSet1.FieldByName( 'BPCodeOrdStr' ).AsString := String1;


    { 適用之優惠辦法 }
    ClientDataSet1.FieldByName( 'DiscountCodeStr' ).AsString := EDT_DiscountCode.Text;


    if EDT_Punish.Checked then
      String1:='1'
    else
      String1:='0';
    ClientDataSet1.FieldByName('Punish').AsString:=String1;

    if EDT_Price.Checked then
      String1:='1'
    else
      String1:='0';
    ClientDataSet1.FieldByName('Price').AsString:=String1;

    if EDT_Cash.Checked then
      String1:='1'
    else
      String1:='0';
    ClientDataSet1.FieldByName('Cash').AsString:=String1;

    if EDT_Coupon.Checked then
      String1:='1'
    else
      String1:='0';
    ClientDataSet1.FieldByName('Coupon').AsString:=String1;

    if EDT_Gift.Checked then
      String1:='1'
    else
      String1:='0';
    ClientDataSet1.FieldByName('Gift').AsString:=String1;

    if RadioButton3.Checked then
      String1:='1'
    else if RadioButton4.Checked then
      String1:='2'
    else
      String1:='';

    ClientDataSet1.FieldByName('PriceDiscountType').AsString:=String1;
    ClientDataSet1.FieldByName('PriceDiscountRate').AsString:=EDT_PriceDiscountRate.Text;
    ClientDataSet1.FieldByName('PriceFixAmount').AsString:=EDT_PriceFixAmount.Text;

    if RadioButton5.Checked then
      String1:='1'
    else if RadioButton6.Checked then
      String1:='2'
    else
      String1:='';
    ClientDataSet1.FieldByName('CashDiscountType').AsString:=String1;

    String1:=EDT_CashOrderRate.Text;
    if String1<>'' then
      String1:=FloatToStr(StrToInt(String1)/100);
    ClientDataSet1.FieldByName('CashOrderRate').AsString:=String1;

    ClientDataSet1.FieldByName('CashFixAmount').AsString:=EDT_CashFixAmount.Text;

    String1:='';
    for I:=0 to Edt_CashCItemCode.Items.Count-1 do
    begin
      String1:=String1+Edt_CashCItemCode.Items[I];
      if I<>Edt_CashCItemCode.Items.Count-1 then
        String1:=String1+',';
     end;
    ClientDataSet1.FieldByName('CashCitemCodeStr').AsString:=String1;

    if RadioButton7.Checked then
      String1:='1'
    else if RadioButton8.Checked then
      String1:='2'
    else
      String1:='';
    ClientDataSet1.FieldByName('CouponDiscountType').AsString:=String1;

    String1:=EDT_CouponOrderRate.Text;
    if String1<>'' then
      String1:=FloatToStr(StrToInt(String1)/100);
    ClientDataSet1.FieldByName('CouponOrderRate').AsString:=String1;

    ClientDataSet1.FieldByName('CouponFixAmount').AsString:=EDT_CouponFixAmount.Text;

    { 綁約參數 }
    String1 := '0';
    if CheckBox5.Checked then String1 := '1';
    ClientDataSet1.FieldByName( 'Bundle' ).AsString := String1;
    ClientDataSet1.FieldByName( 'BundleMon' ).AsString := EDT_BundleMon.Text;
    ClientDataSet1.FieldByName( 'PenalAmt' ).AsString := '0';
    ClientDataSet1.FieldByName( 'PenalType' ).AsString := EmptyStr;
    if EDT_PenalType.ItemIndex > 0 then
      ClientDataSet1.FieldByName( 'PenalType' ).AsString := IntToStr( EDT_PenalType.ItemIndex );
    ClientDataSet1.FieldByName( 'Expiretype' ).AsString := EmptyStr;
    if EDT_Expiretype.ItemIndex > 0 then
      ClientDataSet1.FieldByName( 'Expiretype' ).AsString := IntToStr( EDT_Expiretype.ItemIndex );

    ClientDataSet1.FieldByName( 'Expirecut' ).AsString := '0';

    ClientDataSet1.FieldByName( 'DepositAttr' ).AsString := EmptyStr;
    ClientDataSet1.FieldByName( 'DepositCode' ).AsString := EmptyStr;
    ClientDataSet1.FieldByName( 'DepositName' ).AsString := EmptyStr;
    ClientDataSet1.FieldByName( 'DepositAmt' ).AsString := EmptyStr;
    if EDT_DepositAttr.ItemIndex > 0 then
    begin
      ClientDataSet1.FieldByName( 'DepositAttr' ).AsString := IntToStr( EDT_DepositAttr.ItemIndex );
      ClientDataSet1.FieldByName( 'DepositCode' ).AsString := EDT_DepositCode.Text;
      ClientDataSet1.FieldByName( 'DepositName' ).AsString := EDT_DepositName.Text;
      ClientDataSet1.FieldByName( 'DepositAmt' ).AsString := EDT_DepositAmt.Text;
    end;

    { 可轉促案 }
    String1 := EmptyStr;
    for I:=0 to EDT_CPromotion.Items.Count-1 do
    begin
      String1 := ( String1 + EDT_CPromotion.Items[I] );
      if ( I <> EDT_CPromotion.Items.Count - 1 ) then
        String1 := ( String1 + ',' );
    end;
    ClientDataSet1.FieldByName('ConvertiblePromotionStr').AsString := String1;

    { 階段式違約 }
    ClientDataSet1.FieldByName( 'Penal' ).AsString := '0';
    ClientDataSet1.FieldByName( 'PenalCode' ).AsString := EmptyStr;
    ClientDataSet1.FieldByName( 'PenalName' ).AsString := EmptyStr;
    for I := 1 to 12 do
    begin
      ClientDataSet1.FieldByName( Format( 'PenalMon%d', [I] ) ).AsString := EmptyStr;
      ClientDataSet1.FieldByName( Format( 'PenalAmt%d', [I] ) ).AsString := EmptyStr;
    end;

    if CheckBox7.Checked then
    begin
      ClientDataSet1.FieldByName( 'Penal' ).AsString := '1';
      ClientDataSet1.FieldByName( 'PenalCode' ).AsString := EDT_PenalCode.Text;
      ClientDataSet1.FieldByName( 'PenalName' ).AsString := EDT_PenalName.Text;
    end;

    CD042A.CheckBrowseMode;
    if not VdCD042A then Exit;

    ClientDataSet1.Post;
    ADOConnection1.BeginTrans;
    try

      ClientDataSet1.ApplyUpdates( -1 );
      CD042A.ApplyUpdates( -1 );
      ADOConnection1.CommitTrans;
    except
      ADOConnection1.RollbackTrans;
      ClientDataSet1.Edit;
      WarningMsg('資料更新錯誤');
    end;
//    ClientDataSet1.Post;
//    ClientDataSet1.Refresh;

   if EditMode='Fix' then
     BrowserMode
   else
     btnInsertClick( btnInsert );

  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.ClearGiftValues;
var
  aIndex: Integer;
begin
  if ( ClientDataSet1.State = dsEdit ) then
  begin
    for aIndex := 1 to 6 do
    begin
      ClientDataSet1.FieldByName( Format( 'GiftOrderPrice%d', [aIndex] ) ).Value := Null;
      ClientDataSet1.FieldByName( Format( 'GiftOrderPriceMax%d', [aIndex] ) ).Value := Null;
      ClientDataSet1.FieldByName( Format( 'GiftType%d', [aIndex] ) ).Value := Null;
      ClientDataSet1.FieldByName( Format( 'GiftKind%d', [aIndex] ) ).Value := Null;
      ClientDataSet1.FieldByName( Format( 'ArticleNoStr%d', [aIndex] ) ).Value := Null;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TForm1.CheckGiftValues: Boolean;
var
  aIndex, aLastIdx: Integer;
  aAllEmpty: Boolean;
  aStVal, aEdVal, aLastVal: String;
  aGiftCount, aArticleNoCount: Integer;
begin
  Result := False;
  aAllEmpty :=
    ( ClientDataSet1.FieldByName( 'GiftOrderPrice1' ).AsString = EmptyStr ) and
    ( ClientDataSet1.FieldByName( 'GiftOrderPriceMax1' ).AsString = EmptyStr ) and
    ( ClientDataSet1.FieldByName( 'GiftOrderPrice2' ).AsString = EmptyStr ) and
    ( ClientDataSet1.FieldByName( 'GiftOrderPriceMax2' ).AsString = EmptyStr ) and
    ( ClientDataSet1.FieldByName( 'GiftOrderPrice3' ).AsString = EmptyStr ) and
    ( ClientDataSet1.FieldByName( 'GiftOrderPriceMax3' ).AsString = EmptyStr ) and
    ( ClientDataSet1.FieldByName( 'GiftOrderPrice4' ).AsString = EmptyStr ) and
    ( ClientDataSet1.FieldByName( 'GiftOrderPriceMax4' ).AsString = EmptyStr ) and
    ( ClientDataSet1.FieldByName( 'GiftOrderPrice5' ).AsString = EmptyStr ) and
    ( ClientDataSet1.FieldByName( 'GiftOrderPriceMax5' ).AsString = EmptyStr ) and
    ( ClientDataSet1.FieldByName( 'GiftOrderPrice6' ).AsString = EmptyStr ) and
    ( ClientDataSet1.FieldByName( 'GiftOrderPriceMax6' ).AsString = EmptyStr );
  if aAllEmpty then
  begin
    WarningMsg( '已指定贈品, 請設定贈品內容(至少指定一項)。' );
    Exit;
  end;
  {}
  for aIndex := 1 to 6 do
  begin
    aStVal := ClientDataSet1.FieldByName( Format( 'GiftOrderPrice%d', [aIndex] ) ).AsString;
    aEdVal := ClientDataSet1.FieldByName( Format( 'GiftOrderPriceMax%d', [aIndex] ) ).AsString;
    if ( ( aStVal <> EmptyStr ) and ( aEdVal = EmptyStr ) ) or
       ( ( aStVal = EmptyStr ) and ( aEdVal <> EmptyStr ) ) then
    begin
      WarningMsg( '贈品設定, 訂購金額【起】【迄】必須輸入。' );
      Exit;
    end;
  end;
  {}
  aLastIdx := MaxInt;
  for aIndex := 1 to 6 do
  begin
    aStVal := ClientDataSet1.FieldByName( Format( 'GiftOrderPrice%d', [aIndex] ) ).AsString;
    aEdVal := ClientDataSet1.FieldByName( Format( 'GiftOrderPriceMax%d', [aIndex] ) ).AsString;
    if ( aStVal = EmptyStr ) then
      aLastIdx := aIndex
    else  begin
      if ( aIndex > aLastIdx ) then
      begin
        WarningMsg( '贈品設定, 必須依照順序一～六輸入,不可中斷。' );
        Exit;
      end;
    end;
  end;
  {}
  for aIndex := 1 to 6 do
  begin
    aStVal := ClientDataSet1.FieldByName( Format( 'GiftOrderPrice%d', [aIndex] ) ).AsString;
    aEdVal := ClientDataSet1.FieldByName( Format( 'GiftOrderPriceMax%d', [aIndex] ) ).AsString;
    if ( aStVal <> EmptyStr ) then
    begin
      if ( StrToFloatDef( aEdVal, 0 ) ) < ( StrToFloatDef( aStVal, 0 ) ) then
      begin
        WarningMsg( Format( '贈品設定, 贈品%d 訂購金額【迄】必須大於或等於【起】。', [aIndex] ) );
        Exit;
      end;
      if ( aIndex > 1 ) then
      begin
        if ( StrToFloatDef( aLastVal, 0 ) + 1 ) <> ( StrToFloatDef( aStVal, 0 ) ) then
        begin
          WarningMsg( Format( '贈品設定, 贈品%d 訂購金額一～六必須按照大小排列, 起迄金額不可重疊。', [aIndex] ) );
          Exit;
        end;
      end;
      aLastVal := aEdVal;
    end;
  end;
  {}
  for aIndex := 1 to 6 do
  begin
    aStVal := ClientDataSet1.FieldByName( Format( 'GiftOrderPrice%d', [aIndex] ) ).AsString;
    aEdVal := ClientDataSet1.FieldByName( Format( 'GiftOrderPriceMax%d', [aIndex] ) ).AsString;
    if ( aStVal <> EmptyStr ) then
    begin
      if ClientDataSet1.FieldByName( Format( 'GiftType%d', [aIndex] ) ).AsString = EmptyStr then
      begin
        WarningMsg( Format( '贈品設定, 贈品%d 已設定訂購金額，但沒有設定【贈送方式】。', [aIndex] ) );
        Exit;
      end;
      if ClientDataSet1.FieldByName( Format( 'GiftKind%d', [aIndex] ) ).AsString = EmptyStr then
      begin
        WarningMsg( Format( '贈品設定, 贈品%d 已設定訂購金額，但沒有設定【贈送項目】。', [aIndex] ) );
        Exit;
      end;
      aArticleNoCount := StrCount( ClientDataSet1.FieldByName( Format( 'ArticleNoStr%d', [aIndex] ) ).AsString );
      aGiftCount := GiftCount( ClientDataSet1.FieldByName( Format( 'GiftType%d', [aIndex] ) ).AsString );
      if ( ( aGiftCount in [1..9] ) and ( aGiftCount > aArticleNoCount ) ) then
      begin
        WarningMsg( Format( '贈品設定, 贈品%d 至少須選定%d組。', [aIndex, aGiftCount] ) );
        Exit;
      end else
      if ( aArticleNoCount = 0 ) then
      begin
        WarningMsg( Format( '贈品設定, 贈品%d 至少須選定1組。', [aIndex] ) );
        Exit;
      end;
    end;
  end;
  {}
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.btnGiftClick(Sender: TObject);
var
  aControl: TWinControl;
begin
  aControl := nil;
  if ( TrimChar( EDT_ActStartDate.Text, [#32,'/'] ) = EmptyStr ) then
    aControl := EDT_ActStartDate
  else if ( TrimChar( EDT_ActStopDate.Text, [#32, '/'] ) = EmptyStr ) then
    aControl := EDT_ActStopDate;
  if Assigned( aControl ) then
  begin
    WarningMsg( '請先輸入【促銷活動期間】起~迄。' );
    if aControl.CanFocus then aControl.SetFocus;
    Exit;
  end;
  frmGift := TfrmGift.Create( nil );
  try
    frmGift.CD042DataSet := ClientDataSet1;
    frmGift.DataReader := ADOQuery2;
    frmGift.ServcieTypes := QuotedValue( GetServiceType );
    frmGift.ActStartDate := EDT_ActStartDate.Text;
    { 改看最後接單日 }
    frmGift.ActStopDate := EDT_LastOrderDate.Text;
    frmGift.ShowModal;
  finally
    frmGift.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.btnPackageNoStrClick(Sender: TObject);
var
  I:Integer;
  SelectedValues,String1:string;
begin
  SelectedValues:='';

  for I:=0 to EDT_PackageNo.Items.Count-1 do
    begin
      SelectedValues:=SelectedValues+EDT_PackageNo.Items[I];
      if I<>EDT_PackageNo.Items.Count-1 then
        SelectedValues:=SelectedValues+',';
    end;
  frmMultiSelect:=TfrmMultiSelect.Create(Application);
  try
    frmMultiSelect.CanSelect := EditMode <> '';
    frmMultiSelect.Connection:=ADOConnection1;
    frmMultiSelect.KeyFields:='CODENO';
    frmMultiSelect.KeyValues:=SelectedValues;
    frmMultiSelect.DisplayFields:='CODENO,套餐編號,Description,套餐名稱';
    frmMultiSelect.ResultFields:='Description';
    frmMultiSelect.SQL.Text:='SELECT CODENO,Description From '+sG_DbUserID+
                             '.CD027 Where StopFlag=0 Order BY CODENO';
    frmMultiSelect.QuotedString := '';
    if frmMultiSelect.ShowModal=mrOk then
      begin
        SelectedValues:=frmMultiSelect.SelectedValue;
        EDT_PackageNo.Clear;
        while SelectedValues<>'' do
        begin
          String1:=ExtractValue(SelectedValues);
          EDT_PackageNo.Items.Add(String1);
        end;
        EDT_PackageNoStr.Text := frmMultiSelect.SelectedDisplay;
      end;
  finally
    frmMultiSelect.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.AddDiscountItem;
var
  aSql: String;
begin
  EDT_DiscountCodeStr.Items.Clear;
  cboDiscountCode.Clear;
  AdoQuery2.Close;
  AdoQuery2.SQL.Clear;
  aSql:='Select CODENO, DESCRIPTION ';
  aSql:=aSql+'From '+sG_DbUserID+'.CD048 ';
  aSql:=aSql+'Where (STOPFLAG=0) ';
  AdoQuery2.SQL.Text:=aSql;
  AdoQuery2.Prepared:=True;
  AdoQuery2.Open;
  while not AdoQuery2.Eof do
  begin
    EDT_DiscountCodeStr.Items.Add(AdoQuery2.FieldByName('DESCRIPTION').AsString);
    cboDiscountCode.Items.Add(AdoQuery2.FieldByName('CODENO').AsString);
    AdoQuery2.Next;
  end;
  ADOQuery2.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.AddDepositItem;
var
  aSql, aRefCode: String;
begin
  EDT_DepositCode.Clear;
  EDT_DepositName.Clear;
  ComboBox9.Clear;
  aSql := EmptyStr;
  if EDT_DepositAttr.ItemIndex = 1 then
  begin
    aRefCode := ' ''15'', ''16''  ';
  end else
  begin
    aRefCode := ' ''14'' ';
  end;
  aSql := Format(
    '  SELECT A.CODENO, A.DESCRIPTION         ' +
    '    FROM %s.CD019 A                      ' +
    '   WHERE A.PERIODFLAG = ''0''            ' +
    '     AND A.SIGN = ''+''                  ' +
    '     AND A.REFNO IN ( %s )               ' +
    '     AND A.STOPFLAG = ''0''              ' +
    '   ORDER BY A.CODENO                     ',
    [sG_DbUserID, aRefCode] );
  {}
  AdoQuery2.SQL.Text:=aSql;
  AdoQuery2.Prepared:=True;
  AdoQuery2.Open;
  while not AdoQuery2.Eof do
  begin
    EDT_DepositName.Items.Add(AdoQuery2.FieldByName('DESCRIPTION').AsString);
    ComboBox9.Items.Add(AdoQuery2.FieldByName('CODENO').AsString);
    AdoQuery2.Next;
  end;
  ADOQuery2.Close;
  EDT_DepositName.ItemIndex:=0;
  ComboBox9.ItemIndex:=0;
  EDT_DepositCode.Text := ComboBox9.Items[ComboBox9.ItemIndex];
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.AddPenalItem;
begin
  EDT_PenalCode.Clear;
  EDT_PenalName.Clear;
  ComboBox10.Clear;
  AdoQuery2.SQL.Text := Format(
    '  SELECT A.CODENO, A.DESCRIPTION  ' +
    '    FROM %s.CD019 A               ' +
    '   WHERE A.PERIODFLAG = ''0''     ' +
    '     AND A.SIGN = ''+''           ' +
    '     AND A.REFNO = ''14''         ' +
    '     AND A.STOPFLAG = ''0''       ' +
    '   ORDER BY A.CODENO              ', [sG_DbUserID] );
  AdoQuery2.Prepared := True;
  AdoQuery2.Open;
  while not AdoQuery2.Eof do
  begin
    EDT_PenalName.Items.Add( AdoQuery2.FieldByName( 'DESCRIPTION' ).AsString );
    ComboBox10.Items.Add( AdoQuery2.FieldByName( 'CODENO' ).AsString );
    AdoQuery2.Next;
  end;
  AdoQuery2.Close;
  EDT_PenalName.ItemIndex := 0;
  ComboBox10.ItemIndex := 0;
  EDT_PenalCode.Text := ComboBox10.Items[ComboBox10.ItemIndex];
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.EDT_DiscountCodeStrSelect(Sender: TObject);
begin
  EDT_DiscountCode.Text := cboDiscountCode.Items[EDT_DiscountCodeStr.ItemIndex];
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.EDT_DepositAttrSelect(Sender: TObject);
begin
  if ( EDT_DepositAttr.ItemIndex > 0 ) then
  begin
    AddDepositItem;
  end else
  begin
    EDT_DepositCode.Text := EmptyStr;
    EDT_DepositName.Items.Clear;
    ComboBox9.Items.Clear;
    EDT_DepositAmt.Clear;
  end;
  EDT_DepositCode.Enabled := ( EDT_DepositAttr.ItemIndex > 0 );
  EDT_DepositName.Enabled := ( EDT_DepositAttr.ItemIndex > 0 );
  EDT_DepositAmt.Enabled := ( EDT_DepositAttr.ItemIndex > 0 );
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.EDT_DepositNameSelect(Sender: TObject);
begin
  EDT_DepositCode.Text := ComboBox9.Items[EDT_DepositName.ItemIndex];
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.Button14Click(Sender: TObject);
var
  SelectedValues, String1: String;
  I: Integer;
begin
  SelectedValues:='';
  for I:=0 to EDT_CPromotion.Items.Count-1 do
  begin
    SelectedValues:=SelectedValues+EDT_CPromotion.Items[I];
    if I<>EDT_CPromotion.Items.Count-1 then
      SelectedValues:=SelectedValues+',';
  end;
  frmMultiSelect:=TfrmMultiSelect.Create(Application);
  try
    frmMultiSelect.Connection:=ADOConnection1;
    frmMultiSelect.KeyFields:='CODENO';
    frmMultiSelect.KeyValues:=SelectedValues;
    frmMultiSelect.DisplayFields:='CODENO,促案代碼,Description,促案名稱';
    frmMultiSelect.ResultFields:='Description';
    frmMultiSelect.SQL.Text:=
      'SELECT CODENO,Description From '+sG_DbUserID+
      '.CD042 Where StopFlag=0 AND CODENO <>' + '''' + EDT_CodeNo.Text + ''''+
      ' Order BY CODENO ';
    frmMultiSelect.QuotedString := '';
    if frmMultiSelect.ShowModal=mrOk then
    begin
      SelectedValues:=frmMultiSelect.SelectedValue;
      EDT_CPromotion.Clear;
      while SelectedValues<>'' do
      begin
        String1:=ExtractValue(SelectedValues);
        EDT_CPromotion.Items.Add(String1);
      end;
      EDT_CPromotionStr.Text := frmMultiSelect.SelectedDisplay;
    end;
  finally
    frmMultiSelect.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TForm1.GetServiceType: String;
var
  aIndex: Integer;
begin
  Result := EmptyStr;
  for aIndex := 0 to EDT_ServiceType.Items.Count - 1 do
    Result := ( Result + EDT_ServiceType.Items[aIndex] + ',' );
  if IsDelimiter( ',', Result, Length( Result ) ) then
    Delete( Result, Length( Result ), 1 );
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.EDT_PenalNameSelect(Sender: TObject);
begin
  EDT_PenalCode.Text := ComboBox10.Items[EDT_PenalName.ItemIndex];
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.LoadCD042A;
begin
  CD042A.Close;
  CD042A.CommandText := Format(
    ' SELECT PROMCODE,         ' +
    '        PENALCODE,        ' +
    '        ITEM,             ' +
    '        PENALMON,         ' +
    '        PENALAMT          ' +
    '   FROM %s.CD042A         ' +
    '  WHERE PROMCODE = ''%s'' ' +
    '  ORDER BY ITEM           ', [sG_DbUserID, FKey] );
  CD042A.Open;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.EmptyCD042A;
begin
  while not CD042A.Eof do
    CD042A.Delete;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.ReCalcCD042AItemNo;
var
  aBMark: TBookmark;
  aIndex: Integer;
begin
  CD042A.DisableControls;
  try
    aBMark := CD042A.GetBookmark;
    try
      CD042A.First;
      aIndex := 0;
      while not CD042A.Eof do
      begin
        Inc( aIndex );
        CD042A.Edit;
        CD042A.FieldByName( 'ITEM' ).AsInteger := aIndex;
        CD042A.Post;
        CD042A.Next;
      end;
      CD042A.GotoBookmark( aBMark );
    finally
      CD042A.FreeBookmark( aBMark );
    end;
  finally
    CD042A.EnableControls;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.CD042ANewRecord(DataSet: TDataSet);
begin
  CD042A.FieldByName( 'PromCode' ).AsString := FKey;
  CD042A.FieldByName( 'PENALCODE' ).AsString := EDT_PenalCode.Text;
  CD042A.FieldByName( 'ITEM' ).AsInteger := ( CD042A.RecordCount + 1 );
  CD042A.FieldByName( 'PENALMON' ).AsInteger := 0;
  CD042A.FieldByName( 'PENALAMT' ).AsInteger := 0;
end;

{ ---------------------------------------------------------------------------- }

procedure TForm1.gvPenalCol4PropertiesButtonClick(Sender: TObject;
  AButtonIndex: Integer);
begin
  if ( AButtonIndex = 0 ) then
  begin
    CD042A.Delete;
    ReCalcCD042AItemNo;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TForm1.VdCD042A: Boolean;
begin
  Result := False;
  {}
  if ( CheckBox7.Checked  ) then
  begin
    if CD042A.IsEmpty then
    begin
      cxPageControl2.ActivePageIndex := 2;
      WarningMsg( '已指定計算違約金, 必須設定【階段式違約參數】。' );
      if PenalGrid.CanFocusEx then PenalGrid.SetFocus;
      CD042A.Append;
      Exit;
    end;
  end;
  {}
  if ( CheckBox7.Checked  ) then
  begin
    CD042A.DisableControls;
    try
      CD042A.First;
      while not CD042A.Eof do
      begin
        if ( CD042A.FieldByName( 'PENALMON' ).AsInteger <= 0 ) then
        begin
          Result := False;
          cxPageControl2.ActivePageIndex := 2;
          WarningMsg( '【階段式違約】設定的【未滿綁約月數】必須輸入且不可為零或負值。' );
          if PenalGrid.CanFocusEx then PenalGrid.SetFocus;
          Exit;
        end;
        if ( CheckBox5.Checked ) and ( EDT_BundleMon.Text <> EmptyStr )  then
        begin
          if ( CD042A.FieldByName( 'PENALMON' ).AsString > EDT_BundleMon.Text ) then
          begin
            Result := False;
            cxPageControl2.ActivePageIndex := 2;
            WarningMsg( '【階段式違約】設定的【未滿綁約月數】已超出綁約參數中的【綁約月數】。' );
            if PenalGrid.CanFocusEx then PenalGrid.SetFocus;
            Exit;
          end;
        end;
        if ( CD042A.FieldByName( 'PENALAMT' ).AsInteger <= 0 ) then
        begin
          Result := False;
          cxPageControl2.ActivePageIndex := 2;
          WarningMsg( '【階段式違約】設定的【違約金額】必須輸入且不可為零或負值。' );
          if PenalGrid.CanFocusEx then PenalGrid.SetFocus;
          Exit;
        end;
        CD042A.Next;
      end;
    finally
      CD042A.EnableControls;
    end;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }



procedure TForm1.GetQryData;
  var aSQL:string;
begin
  {
  aSQL :=Format(' SELECT * FROM %S.CD042 ' +
                '  WHERE ( DISCOUNTCODE IS NULL ) AND ( PRODUCTTYPE IS NOT NULL ) ' + F_strQry +
                '  ORDER BY CODENO', [sG_DbUserID] );
  }
  Screen.Cursor := crSQLWait;

  try
    ClientDataSet1.Active := False;
    ClientDataSet1.CommandText := aSQL;
    ClientDataSet1.AfterScroll := nil;
    try
      ClientDataSet1.Open;
    finally
      ClientDataSet1.AfterScroll := ClientDataSet1AfterScroll;
    end;
    BrowserMode;
  finally
    Screen.Cursor := crDefault;
  end;

end;

procedure TForm1.btnFindClick(Sender: TObject);
  var
    aQryString: string;
begin
  try
    frmQryCD042 := TfrmQryCD042.Create(Application);
    if frmQryCD042.ShowModal= mrOk then
    begin

      try
        Self.Update;
        Screen.Cursor := crSQLWait;
        ClientDataSet1.Active := False;
        aQryString := frmQryCD042.sWhere;
        ClientDataSet1.CommandText := Format(
              ' SELECT * FROM %S.CD042 ' +
              '  WHERE ( DISCOUNTCODE IS NULL ) AND ( PRODUCTTYPE IS NOT NULL ) ' + aQryString +
              '  ORDER BY CODENO', [sG_DbUserID] );

        ClientDataSet1.AfterScroll := nil;
        try
          ClientDataSet1.Open;
        finally
          ClientDataSet1.AfterScroll := ClientDataSet1AfterScroll;
        end;
        if not ClientDataSet1.Eof then

          BrowserMode

        else begin
          WarningMsg('查詢不到任何資料!!');
          Panel5.Caption:='0 ／ 0';
          QueryMode;
        end;
      finally
        Screen.Cursor := crDefault;
      end;
    end;
  finally
    frmQryCD042.Free;
  end;
end;

procedure TForm1.QueryMode;
begin
  ClientDataSet1.Cancel;
  CD042A.Cancel;
  btnPrint.Enabled:= False;
  btnInsert.Enabled:= True;   { #4437 沒有資料也要可以新增 By Kin 2009/03/19 }
  btnModify.Enabled:= False;
  Button10.Visible:=True;
  DbGrid1.Enabled:=False;
  Button12.Visible:=False;
  btnSave.Enabled:=False;
  {#7222}
  if btnInsert.Enabled then
    btnInsert.Enabled := FCanInsert;
  if btnModify.Enabled then
    btnModify.Enabled := FCanModify;
  if btnPrint.Enabled then
    btnPrint.Enabled := FCanPrint;
  {}
  //7238
//  Panel6.Enabled := False;
//  Panel7.Enabled := False;
  Panel8.Enabled := False;
  Panel9.Enabled := False;
  {}
  Panel1.Enabled := False;
  Panel2.Enabled := False;
  Panel3.Enabled := False;
  Panel4.Enabled := False;
  {}
  Panel10.Enabled := False;
  Panel11.Enabled := False;
  Panel12.Enabled := False;
  Panel13.Enabled := False;
  {}
  PenalGrid.Enabled := False;
  gvPenal.Styles.Content := cxStyle1;

  gvPenal.Styles.Content := cxStyle1;
  Edit35.Text:='顯示';
  EditMode:='';
  ClearAllControls;
end;

procedure TForm1.Action5Execute(Sender: TObject);
begin
  if btnFind.Enabled then
    btnFindClick(Sender);
end;

procedure TForm1.btnAreaClick(Sender: TObject);
var
  I:Integer;
  SelectedValues,String1:string;
  oldValue:String;
begin
  SelectedValues:='';
  oldValue := EDT_AreaNoStr.Text;
  for I:=0 to EDT_AreaNo.Items.Count-1 do
  begin
    SelectedValues:=SelectedValues+EDT_AreaNo.Items[I];
    if I<>EDT_AreaNo.Items.Count-1 then
      SelectedValues:=SelectedValues+',';
  end;
  frmMultiSelect:=TfrmMultiSelect.Create(Application);
  try
    frmMultiSelect.CanSelect := EditMode <> '';
    frmMultiSelect.Connection:=ADOConnection1;
    frmMultiSelect.KeyFields:='CODENO';
    frmMultiSelect.KeyValues:= SelectedValues;
    frmMultiSelect.DisplayFields:='CODENO,行政區代碼,DESCRIPTION,行政區名稱';
    frmMultiSelect.ResultFields:='CODENO';
    frmMultiSelect.SQL.Text:='SELECT  CODENO, DESCRIPTION FROM '+sG_DbUserID+
                             '.CD001 Where (COMPCODE='+sG_CompID+')  '+
                             'Order By CODENO';
    frmMultiSelect.QuotedString:='';
    if frmMultiSelect.ShowModal=mrOk then
      begin
        SelectedValues:=frmMultiSelect.SelectedValue;
        EDT_AreaNo.Clear;
        while SelectedValues<>'' do
          begin
            String1:=ExtractValue(SelectedValues);
            EDT_AreaNo.Items.Add(String1);
          end;
          ShowAreaCode;
          if oldValue <> EDT_AreaNoStr.Text then
          begin
            EDT_Bourg.Clear;
            EDT_BourgStr.Clear;
          end;
//        ShowCircuitNo;
      end;
  finally
    frmMultiSelect.Free;
  end;
end;

procedure TForm1.ShowAreaCode;
var
  I:Integer;
  WhereString:string;
begin
  EDT_AreaNoStr.Text:='';
  WhereString:='';

  for I:=0 to EDT_AreaNo.Items.Count-1 do
    begin
      if I=0 then
        WhereString:=WhereString+'Where '
      else
        WhereString:=WhereString+'Or ';
      WhereString:=WhereString+'(CODENO='''+EDT_AreaNo.Items[I]+''') ';
    end;

  if WhereString<>'' then
    begin
      AdoQuery2.Close;
      AdoQuery2.SQL.Clear;
      AdoQuery2.SQL.Text:='Select DESCRIPTION From '+sG_DbUserID+'.CD001 '+WhereString;
      AdoQuery2.Prepared:=True;
      AdoQuery2.Open;
      while not AdoQuery2.Eof do
        begin
          EDT_AreaNoStr.Text:=EDT_AreaNoStr.Text+AdoQuery2.FieldByName('DESCRIPTION').AsString;
          AdoQuery2.Next;
          if not AdoQuery2.Eof then
            EDT_AreaNoStr.Text:=EDT_AreaNoStr.Text+',';

        end;

    end;
end;

procedure TForm1.btnServCodeClick(Sender: TObject);
var
  I:Integer;
  SelectedValues,String1:string;
begin
  SelectedValues:='';

  for I:=0 to EDT_ServCode.Items.Count-1 do
  begin
    SelectedValues:=SelectedValues+EDT_ServCode.Items[I];
    if I<>EDT_ServCode.Items.Count-1 then
      SelectedValues:=SelectedValues+',';
  end;
  frmMultiSelect:=TfrmMultiSelect.Create(Application);
  try
    frmMultiSelect.CanSelect := EditMode <> '';
    frmMultiSelect.Connection:=ADOConnection1;
    frmMultiSelect.KeyFields:='CODENO';
    frmMultiSelect.KeyValues:= SelectedValues;
    frmMultiSelect.DisplayFields:='CODENO,服務區代碼,DESCRIPTION,服務區名稱';
    frmMultiSelect.ResultFields:='CODENO';
    frmMultiSelect.SQL.Text:='SELECT  CODENO, DESCRIPTION FROM '+sG_DbUserID+
                             '.CD002 Where nvl(stopflag,0)=0  '+
                             ' Order By CODENO';
    frmMultiSelect.QuotedString:='';
    if frmMultiSelect.ShowModal=mrOk then
      begin
        SelectedValues:=frmMultiSelect.SelectedValue;
        EDT_ServCode.Clear;
        while SelectedValues<>'' do
          begin
            String1:=ExtractValue(SelectedValues);
            EDT_ServCode.Items.Add(String1);
          end;
          ShowServCode;
//        ShowCircuitNo;
      end;
  finally
    frmMultiSelect.Free;
  end;
end;

procedure TForm1.ShowServCode;
var
  I:Integer;
  WhereString:string;
begin
  EDT_ServCodeStr.Text:='';
  WhereString:='';

  for I:=0 to EDT_ServCode.Items.Count-1 do
    begin
      if I=0 then
        WhereString:=WhereString+'Where '
      else
        WhereString:=WhereString+'Or ';
      WhereString:=WhereString+'(CODENO='''+EDT_ServCode.Items[I]+''') ';
    end;

  if WhereString<>'' then
    begin
      AdoQuery2.Close;
      AdoQuery2.SQL.Clear;
      AdoQuery2.SQL.Text:='Select DESCRIPTION From '+sG_DbUserID+'.CD002 '+WhereString;
      AdoQuery2.Prepared:=True;
      AdoQuery2.Open;
      while not AdoQuery2.Eof do
        begin
          EDT_ServCodeStr.Text:=EDT_ServCodeStr.Text+AdoQuery2.FieldByName('DESCRIPTION').AsString;
          AdoQuery2.Next;
          if not AdoQuery2.Eof then
            EDT_ServCodeStr.Text:=EDT_ServCodeStr.Text+',';
        end;
    end;
end;

procedure TForm1.ShowBundleProdCodeOrd;
var
  I:Integer;
  WhereString:string;
  SelectedValues,String1:string;
begin
  EDT_BundleProdCodeOrd.Clear;


  SelectedValues:=ClientDataSet1.FieldByName( 'BPCodeOrdStr' ).AsString;
  EDT_BundleProdCodeOrd.Clear;
  while SelectedValues<>'' do
    begin
      String1:=ExtractValue(SelectedValues);
      EDT_BundleProdCodeOrd.Items.Add( String1 );
    end;



  {
  WhereString:='';

  for I:=0 to EDT_BundleProdCodeOrd.Items.Count-1 do
    begin
      if I=0 then
        WhereString:=WhereString+'Where '
      else
        WhereString:=WhereString+'Or ';
      WhereString:=WhereString+'(CODENO='''+EDT_BundleProdCode.Items[I]+''') ';
    end;

  if WhereString<>'' then
    begin
      AdoQuery2.Close;
      AdoQuery2.SQL.Clear;
      AdoQuery2.SQL.Text:='Select Description From '+sG_DbUserID+'.CD042 '+WhereString;
      AdoQuery2.Prepared:=True;
      AdoQuery2.Open;
      while not AdoQuery2.Eof do
        begin
          EDT_BundleProdCodeStr.Text:=EDT_BundleProdCodeStr.Text+AdoQuery2.FieldByName('Description').AsString;
          AdoQuery2.Next;
          if not AdoQuery2.Eof then
            EDT_BundleProdCodeStr.Text:=EDT_BundleProdCodeStr.Text+',';
        end;
    end;
    }
end;

procedure TForm1.btnWorkClassClick(Sender: TObject);
var
  I:Integer;
  SelectedValues,String1:string;
begin
  SelectedValues:='';
  for I:=0 to EDT_WorkClass.Items.Count-1 do
  begin
    SelectedValues:=SelectedValues+EDT_WorkClass.Items[I];
    if I<>EDT_WorkClass.Items.Count-1 then
      SelectedValues:=SelectedValues+',';
  end;
  frmMultiSelect:=TfrmMultiSelect.Create(Application);
  try
    frmMultiSelect.CanSelect := EditMode <> '';
    frmMultiSelect.Connection:=ADOConnection1;
    frmMultiSelect.KeyFields:='CODENO';
    frmMultiSelect.KeyValues:= SelectedValues;
    frmMultiSelect.DisplayFields:='CODENO,工作類別代碼,DESCRIPTION,工作類別名稱';
    frmMultiSelect.ResultFields:='CODENO';
    frmMultiSelect.SQL.Text:='SELECT  CODENO, DESCRIPTION FROM '+sG_DbUserID+
                             '.CD071 Where nvl(stopflag,0)=0  '+
                             ' Order By CODENO';
    frmMultiSelect.QuotedString:='';
    if frmMultiSelect.ShowModal=mrOk then
      begin
        SelectedValues:=frmMultiSelect.SelectedValue;
        EDT_WorkClass.Clear;
        while SelectedValues<>'' do
          begin
            String1:=ExtractValue(SelectedValues);
            EDT_WorkClass.Items.Add(String1);
          end;
          ShowWorkClass;
//        ShowCircuitNo;
      end;
  finally
    frmMultiSelect.Free;
  end;
end;

procedure TForm1.ShowWorkClass;
var
  I:Integer;
  WhereString:string;
begin
  edt_WorkClassStr.Text:='';
  WhereString:='';

  for I:=0 to EDT_WorkClass.Items.Count-1 do
    begin
      if I=0 then
        WhereString:=WhereString+'Where '
      else
        WhereString:=WhereString+'Or ';
      WhereString:=WhereString+'(CODENO='''+EDT_WorkClass.Items[I]+''') ';
    end;

  if WhereString<>'' then
    begin
      AdoQuery2.Close;
      AdoQuery2.SQL.Clear;
      AdoQuery2.SQL.Text:='Select DESCRIPTION From '+sG_DbUserID+'.CD071 '+WhereString;
      AdoQuery2.Prepared:=True;
      AdoQuery2.Open;
      while not AdoQuery2.Eof do
        begin
          edt_WorkClassStr.Text:=edt_WorkClassStr.Text+AdoQuery2.FieldByName('DESCRIPTION').AsString;
          AdoQuery2.Next;
          if not AdoQuery2.Eof then
            edt_WorkClassStr.Text:=edt_WorkClassStr.Text+',';
        end;
    end;
end;

procedure TForm1.btnClctEnClick(Sender: TObject);
var
  I:Integer;
  SelectedValues,String1,aWorkClassWhere:string;

begin
  SelectedValues:='';
  aWorkClassWhere := ' 1 = 1 ';
  for I:=0 to EDT_ClctEn.Items.Count-1 do
  begin
    SelectedValues:=SelectedValues+EDT_ClctEn.Items[I];
    if I<>EDT_ClctEn.Items.Count-1 then
      SelectedValues:=SelectedValues+',';
  end;

  for I:=0 to EDT_WorkClass.Items.Count-1 do
  begin
    if ( i = 0 ) then
      aWorkClassWhere := '''' + EDT_WorkClass.Items[I] + ''''
    else
      aWorkClassWhere := aWorkClassWhere + ','''+EDT_WorkClass.Items[I] + '''';
    if ( I = EDT_WorkClass.Count - 1) then
    begin
      aWorkClassWhere :=  'WorkClass IN (' + aWorkClassWhere +') ';
    end;
  end;


  frmMultiSelect:=TfrmMultiSelect.Create(Application);
  try
    frmMultiSelect.CanSelect := EditMode <> '';
    frmMultiSelect.Connection:=ADOConnection1;
    frmMultiSelect.KeyFields:='EMPNO';
    frmMultiSelect.KeyValues:= SelectedValues;
    frmMultiSelect.DisplayFields:='EMPNO,員工代碼,EMPNAME,員工姓名';
    frmMultiSelect.ResultFields:='EMPNO';
    frmMultiSelect.SQL.Text:='SELECT  EMPNO, EMPNAME FROM '+sG_DbUserID+
                             '.CM003 Where nvl(stopflag,0)=0  '+
                             ' AND ' + aWorkClassWhere +
                             ' Order By EMPNO';
    frmMultiSelect.QuotedString:='';
    if frmMultiSelect.ShowModal=mrOk then
      begin
        SelectedValues:=frmMultiSelect.SelectedValue;
        EDT_ClctEn.Clear;
        while SelectedValues<>'' do
          begin
            String1:=ExtractValue(SelectedValues);
            EDT_ClctEn.Items.Add(String1);
          end;
        ShowClctEn;

      end;
  finally
    frmMultiSelect.Free;
  end;
end;

procedure TForm1.ShowClctEn;
var
   I:Integer;
  WhereString:string;
begin
  EDT_ClctEnStr.Text:='';
  WhereString:='';

  for I:=0 to EDT_ClctEn.Items.Count-1 do
    begin
      if I=0 then
        WhereString:=WhereString+'Where '
      else
        WhereString:=WhereString+'Or ';
      WhereString:=WhereString+'(EMPNO='''+EDT_ClctEn.Items[I]+''') ';
    end;

  if WhereString<>'' then
    begin
      AdoQuery2.Close;
      AdoQuery2.SQL.Clear;
      AdoQuery2.SQL.Text:='Select EMPNAME From '+sG_DbUserID+'.CM003 '+WhereString;
      AdoQuery2.Prepared:=True;
      AdoQuery2.Open;
      while not AdoQuery2.Eof do
        begin
          EDT_ClctEnStr.Text:=EDT_ClctEnStr.Text+AdoQuery2.FieldByName('EMPNAME').AsString;
          AdoQuery2.Next;
          if not AdoQuery2.Eof then
            EDT_ClctEnStr.Text:=EDT_ClctEnStr.Text+',';
        end;
    end;
end;

procedure TForm1.EDT_WorkClassStrChange(Sender: TObject);
begin
  EDT_ClctEn.Clear;
  EDT_ClctEnStr.Clear;
end;

procedure TForm1.btnBourgCodeClick(Sender: TObject);
 var
  I:Integer;
  SelectedValues,String1,aBourgWhere:string;
  aSQL : String;
begin

  SelectedValues := '';
  aBourgWhere := ' 1 = 1 ';

  for I:=0 to EDT_Bourg.Items.Count-1 do
  begin
    SelectedValues:=SelectedValues+EDT_Bourg.Items[I];
    if I<>EDT_Bourg.Items.Count-1 then
      SelectedValues:=SelectedValues+',';
  end;


  for I:=0 to EDT_AreaNo.Items.Count-1 do
  begin
    if ( i = 0 ) then
      aBourgWhere := '''' + EDT_AreaNo.Items[I] + ''''
    else
      aBourgWhere := aBourgWhere + ','''+EDT_AreaNo.Items[I] + '''';
    if ( I = EDT_AreaNo.Count - 1) then
    begin
      aBourgWhere :=  ' CODENO IN ( SELECT Nvl(BourgCode,-1) BourgCode FROM  ' + sG_DbUserID +
                      '.CD001A WHERE AreaCode IN (' + aBourgWhere + ' ) )                   ';
       aSQL := 'SELECT COUNT(*) FROM      ' + sG_DbUserID      +
        '.CD058 WHERE ' + aBourgWhere + ' And Nvl(StopFlag,0) = 0';

      ADOQuery3.Close;
      try
        ADOQuery3.SQL.Text := aSQL;
        ADOQuery3.Open;
        if ADOQuery3.Fields[0].AsInteger = 0  then
        begin
          aBourgWhere := ' 1 = 1 ';
        end;
      finally
        ADOQuery3.Close;
      end;



    end;

  end;


  frmMultiSelect:=TfrmMultiSelect.Create(Application);
  try
    frmMultiSelect.CanSelect := EditMode <> '';
    frmMultiSelect.Connection:=ADOConnection1;
    frmMultiSelect.KeyFields:='CodeNo';
    frmMultiSelect.KeyValues:= SelectedValues;
    frmMultiSelect.DisplayFields:='CODENO,代碼,DESCRIPTION,名稱';
    frmMultiSelect.ResultFields:='CODENO';
    frmMultiSelect.SQL.Text:='SELECT  CODENO, DESCRIPTION FROM '+sG_DbUserID+
                             '.CD058 Where nvl(stopflag,0)=0  '+
                             ' AND ' + aBourgWhere +
                             ' Order By CODENO';
    frmMultiSelect.QuotedString:='';
    if frmMultiSelect.ShowModal=mrOk then
      begin
        SelectedValues:=frmMultiSelect.SelectedValue;
        EDT_Bourg.Clear;
        while SelectedValues<>'' do
          begin
            String1:=ExtractValue(SelectedValues);
            EDT_Bourg.Items.Add(String1);
          end;
        ShowBourgCode;

      end;
  finally
    frmMultiSelect.Free;
  end;
end;

procedure TForm1.ShowBourgCode;
  var
   I:Integer;
    WhereString:string;
begin
  EDT_BourgStr.Text:='';
  WhereString:='';
  for I:=0 to EDT_Bourg.Items.Count-1 do
  begin
    if I=0 then
      WhereString := WhereString+'Where '
    else
      WhereString := WhereString+'Or ';
    WhereString := WhereString+'(CODENO = '+EDT_Bourg.Items[I]+ ') ';
  end;

  if WhereString<>'' then
    begin
      AdoQuery2.Close;
      AdoQuery2.SQL.Clear;
      AdoQuery2.SQL.Text:='Select DESCRIPTION From ' + sG_DbUserID+'.CD058 ' + WhereString;
      AdoQuery2.Prepared:=True;
      AdoQuery2.Open;
      while not AdoQuery2.Eof do
        begin
          EDT_BourgStr.Text:=EDT_BourgStr.Text+AdoQuery2.FieldByName('DESCRIPTION').AsString;
          AdoQuery2.Next;
          if not AdoQuery2.Eof then
            EDT_BourgStr.Text:=EDT_BourgStr.Text+',';
        end;
    end;
end;

procedure TForm1.AddBourgItem;
  var
    I : Integer;
    aBourgWhere,aSQL : String;
begin
  aBourgWhere := ' 1 = 1 ';
  for I:=0 to EDT_AreaNo.Items.Count-1 do
  begin
    if ( i = 0 ) then
      aBourgWhere := '''' + EDT_AreaNo.Items[I] + ''''
    else
      aBourgWhere := aBourgWhere + ','''+EDT_AreaNo.Items[I] + '''';
    if ( I = EDT_AreaNo.Count - 1) then
    begin
      aBourgWhere :=  ' CODENO IN ( SELECT Nvl(BourgCode,-1) BourgCode FROM  ' + sG_DbUserID +
                      '.CD001A WHERE AreaCode IN (' + aBourgWhere + ' ) )                   ';
       aSQL := 'SELECT COUNT(*) FROM      ' + sG_DbUserID      +
        '.CD058 WHERE ' + aBourgWhere ;

      ADOQuery3.Close;
      try
        ADOQuery3.SQL.Text := aSQL;
        ADOQuery3.Open;
        if ADOQuery3.Fields[0].AsInteger = 0  then
        begin
          aBourgWhere := ' 1 = 1 ';
        end;
      finally
        ADOQuery3.Close;
      end;
    end;
  end;

  EDT_BourgCode.Clear;
  EDT_BourgName.Clear;
  EDTBourgComboBox.Clear;
  AdoQuery2.SQL.Text := Format(
    'SELECT * FROM (                                 ' +
    '  SELECT NULL CODENO,NULL DESCRIPTION FROM DUAL ' +
    '  UNION ALL                                     ' +
    '  SELECT CODENO,     DESCRIPTION  ' +
    '    FROM %s.CD058                 ' +
    '   WHERE NVL(STOPFLAG,0) = 0      ' +
    '     AND %s                       ' +
    ' ) ORDER BY NVL(CODENO,-1)        ', [sG_DbUserID,aBourgWhere] );
  AdoQuery2.Prepared := True;
  AdoQuery2.Open;
  while not AdoQuery2.Eof do
  begin
    EDT_BourgName.Items.Add( AdoQuery2.FieldByName( 'DESCRIPTION' ).AsString );
    EDTBourgComboBox.Items.Add( AdoQuery2.FieldByName( 'CODENO' ).AsString );
    AdoQuery2.Next;
  end;
  AdoQuery2.Close;
  EDT_BourgName.ItemIndex := 0;
  EDTBourgComboBox.ItemIndex := 0;
  EDT_BourgCode.Text := EDTBourgComboBox.Items[EDTBourgComboBox.ItemIndex];
end;

procedure TForm1.EDT_BourgNameSelect(Sender: TObject);
begin
  EDT_BourgCode.Text := EDTBourgComboBox.Items[EDT_BourgName.ItemIndex];
end;

procedure TForm1.EDT_AreaNoStrChange(Sender: TObject);
begin
  //AddBourgItem;
end;

function TForm1.GetUserPriv(const Mid: String): Boolean;
begin

  if FGroupId = '0' then
  begin
    Result := True;
    Exit;
  end;
  try
    try
      
      ADOQuery2.Close;
      ADOQuery2.SQL.Text := Format(
                        'SELECT COUNT(*) CNT FROM %s.SO029 ' +
                        ' WHERE MID = ''%s''               ' +
                        ' AND GROUP%s = 1                  ',
                        [sG_DbUserID,
                        Mid,
                        FGroupId]);

      ADOQuery2.Open;
      Result := ADOQuery2.FieldByName( 'CNT' ).AsInteger = 1;

    except
      on E: Exception do
      begin
        Result := False;
      end;
    end;
  finally
    ADOQuery2.Close;
  end;

end;

procedure TForm1.GetGropuId;
begin
  FGroupId := '0';
  try
    if sG_IsSupervisor = 'N' then
    begin
      ADOQuery2.Close;
      ADOQuery2.SQL.Text := format('SELECT GroupID FROM %s.SO026 ' +
                                    ' WHERE UserId = ''%s''        '             +
                                    ' AND COMPCODE = %s            ',
                                    [sG_DbUserID,
                                    sG_OperatorID,
                                    sG_CompID]);
      ADOQuery2.Open;
      if ADOQuery2.RecordCount > 0 then
        FGroupId := ADOQuery2.FieldByName( 'GroupID' ).AsString;
    end;
  finally
    ADOQuery2.Close;
  end;
end;

end.
