unit dtmMain2U;

interface

uses
  SysUtils, Classes, DB, DBClient, MConnect,Dialogs ,Variants,Math, XMLDoc, XMLIntf,
  xmldom, msxmldom;

const
  DATA_SEP_1='~';
  DATA_SEP_2='-';

  DATA_AREA_HEADER = 'COMP';
  PARAM_COUNT= 10;
  SERVICE_TYPE = 'C';
  TABLE_SO033 = 'SO033';
  TABLE_SO034 = 'SO034';
  FUNCTION0 = 0;  //璸衡膀セそΑ
  FUNCTION1 = 1;  //璸衡甅ノ "そΑ:ㄢ顶甅繺-ゑㄒ" 繵笵
  FUNCTION2 = 2;  //璸衡甅ノ "そΑ:め计计秖禯-虫基" 籔" そΑ:め计κだゑ禯-虫基"繵笵
  FUNCTION3 = 3;  //璸衡甅ノ "そΑ:め计计秖禯-ゑㄒ" 籔" そΑき:め计κだゑ禯-ゑㄒ"繵笵
  IS_STEP_PERCENT = true;        //め计κだゑ禯---め计 = κだゑ * 羆Τめ计( そΑのそΑき )
  IS_STEP_NOT_PERCENT = false;   //め计计秖禯---ぃノκだゑ璸衡め计秖 ( そΑのそΑ )
  IS_PRICE = true;   //虫基
  IS_NOT_PRICE = false;   //ゑㄒ


  REAL_DATE_TYPE = 0;
  SHOULD_DATE_TYPE = 1;

type

  //纗繵笵羆Μ跌め计戈
  TChannelViewData = class(TObject)
    sProduct_ID : String;
    sCitemCode : String;
    nChannelTotalViewCounts : Integer;
  end;

  TdtmMain2 = class(TDataModule)
    cdsSo110: TClientDataSet;
    DCOMConnection1: TDCOMConnection;
    cdsSo110FORMULA_ID: TStringField;
    cdsSo110FORMULA_NAME: TStringField;
    cdsSo110FORMULA_DESC: TStringField;
    cdsSo110OPERATOR: TStringField;
    cdsSo112M: TClientDataSet;
    cdsSo112MPRODUCT_NAME: TStringField;
    cdsSo112MTYPE: TStringField;
    cdsSo112MOPERATOR: TStringField;
    cdsCd024CT: TClientDataSet;
    cdsSqlStatement: TClientDataSet;
    cdsSo112MPROVIDER_ID: TStringField;
    cdsSo113CT: TClientDataSet;
    cdsSo113CTPROVIDER_ID: TStringField;
    cdsSo113CTPROVIDER_NAME: TStringField;
    cdsSo113CTCONTACTEE: TStringField;
    cdsSo113CTTEL: TStringField;
    cdsSo113CTADDRESS: TStringField;
    cdsSo113CTNOTES: TStringField;
    cdsSo113CTOPERATOR: TStringField;
    cdsSo111D: TClientDataSet;
    cdsSo111DFORMULA_ID: TStringField;
    cdsSo111DOPERATOR: TStringField;
    cdsSo111DFORMULA_NAME_LU: TStringField;
    cdsSqlDelete: TClientDataSet;
    cdsSo113CTATTRIBUTE_NAME: TStringField;
    cdsSo113CTSTOPFLAG_NAME: TStringField;
    cdsSo110UPDTIME: TStringField;
    cdsSo112MUPDTIME: TStringField;
    cdsSo111DUPDTIME: TStringField;
    cdsSo113CTUPDTIME: TStringField;
    cdsSo110STOPFLAG_NAME: TStringField;
    cdsCompute_issue: TClientDataSet;
    cdsSo111DCOMPUTE_ISSUE_NAME: TStringField;
    cdsCD039: TClientDataSet;
    cdsSo113Cd039: TClientDataSet;
    cdsSo111DCOMP_TYPE: TStringField;
    cdsSo111DSO_PROVIDER_ID: TStringField;
    cdsSo113: TClientDataSet;
    cdsSo111DSO_PROVIDER_DESC: TStringField;
    cdsSo113Cd039CLASSIFY_ID: TStringField;
    cdsSo113Cd039CODE: TStringField;
    cdsSo113Cd039DESC: TStringField;
    cdsSo111DPRODUCT_ID: TStringField;
    cdsSo111DPRODUCT_NAME: TStringField;
    cdsRange_Unit: TClientDataSet;
    cdsSo111DRANGE_UNIT_NAME: TStringField;
    cdsSo114N: TClientDataSet;
    cdsCodeCd039: TClientDataSet;
    cdsParam: TClientDataSet;
    cdsProductInfoN: TClientDataSet;
    cdsRefDataAN: TClientDataSet;
    cdsV_ChargeInfoForTally: TClientDataSet;
    cdsChargeDataN: TClientDataSet;
    cdsPackageId: TClientDataSet;
    cdsRefDataN: TClientDataSet;
    cdsPackageInfoN: TClientDataSet;
    cdsSearSo112: TClientDataSet;
    cdsRefData2: TClientDataSet;
    cdsCodeCD039A: TClientDataSet;
    cdsCom: TClientDataSet;
    cdsRefData: TClientDataSet;
    cdsRefDataA: TClientDataSet;
    cdsChargeData: TClientDataSet;
    cdsCustData: TClientDataSet;
    cdsProductInfo: TClientDataSet;
    cdsPackageInfo: TClientDataSet;
    cdsSO115: TClientDataSet;
    cdsSO114: TClientDataSet;
    cdsSO116: TClientDataSet;
    cdsSO117: TClientDataSet;
    cdsCd024Look: TClientDataSet;
    cdsAttribute: TClientDataSet;
    cdsCd024LookCODENO: TStringField;
    cdsCd024LookCODENO_DESCRIPTION: TStringField;
    cdsCd024LookDESCRIPTION: TStringField;
    cdsSo113Look: TClientDataSet;
    cdsSo113LookPROVIDER_ID: TStringField;
    cdsSo113LookPROVIDER_ID_NAME: TStringField;
    cdsSo113LookPROVIDER_NAME: TStringField;
    cdsSo112MPROVIDER_NAME: TStringField;
    cdsSo110STOPFLAG: TIntegerField;
    cdsSo113CTATTRIBUTE: TIntegerField;
    cdsSo113CTSTOPFLAG: TIntegerField;
    cdsSo112MPRICE: TIntegerField;
    cdsSo111DCOMPUTE_ISSUE: TIntegerField;
    cdsSo111DRANGE_UNIT: TStringField;
    cdsSo111DSUBSCRIBER_COUNT_PERCENT1: TBCDField;
    cdsSo111DSUBSCRIBER_COUNT_PERCENT2: TBCDField;
    cdsSo111DSUBSCRIBER_COUNT_PERCENT3: TBCDField;
    cdsSo111DSUBSCRIBER_COUNT_PERCENT4: TBCDField;
    cdsSo111DSUBSCRIBER_COUNT_PERCENT5: TBCDField;
    cdsSo111DAMOUNT_PERCENT1: TBCDField;
    cdsSo111DAMOUNT_PERCENT2: TBCDField;
    cdsSo111DAMOUNT_PERCENT3: TBCDField;
    cdsSo111DAMOUNT_PERCENT4: TBCDField;
    cdsSo111DAMOUNT_PERCENT5: TBCDField;
    cdsSo112MREF_NO: TIntegerField;
    cdsCd024CTCODENO: TStringField;
    cdsCd024CTDESCRIPTION: TStringField;
    cdsCd024CTREFNO: TIntegerField;
    cdsCd024CTPAYFLAG: TIntegerField;
    cdsCd024CTCITEMCODE: TIntegerField;
    cdsCd024CTCITEMNAME: TStringField;
    cdsCd024CTUPDTIME: TStringField;
    cdsCd024CTUPDEN: TStringField;
    cdsCd024CTCOMPCODE: TIntegerField;
    cdsCd024CTSTOPFLAG: TIntegerField;
    cdsCd024CTCHANCEDAYS: TIntegerField;
    cdsCd024CTCHANNELID: TStringField;
    cdsSo112MPRODUCT_ID: TStringField;
    cdsSO115B: TClientDataSet;
    cdsPriv: TClientDataSet;
    cdsSo110Look: TClientDataSet;
    cdsSo110LookFORMULA_ID: TStringField;
    cdsSo110LookFORMULA_ID_NAME: TStringField;
    cdsSo110LookFORMULA_NAME: TStringField;
    cdsChannelCounts: TClientDataSet;
    cdsChannelCountsProductID: TStringField;
    cdsChannelCountsChannelName: TStringField;
    cdsChannelCountsCitemCode: TStringField;
    cdsChannelCountsChannelTotalViewCounts: TIntegerField;
    cdsSO114B: TClientDataSet;
    cdsTallyRefData1: TClientDataSet;
    cdsTallyRefData2: TClientDataSet;
    cdsSO119: TClientDataSet;
    cdsCD019: TClientDataSet;
    cdsTempSO119: TClientDataSet;
    cdsRefNo: TClientDataSet;
    cdsSo110REF_NO: TStringField;
    cdsSo110REFNO_NAME: TStringField;
    cdsSo112MUSEBASEFORMULA: TStringField;
    cdsSo112MPROVIDER_PERCENT: TIntegerField;
    cdsSo112MSO_PERCENT: TIntegerField;
    cdsSO114C: TClientDataSet;
    cdsRefNoCODENO: TStringField;
    cdsRefNoDESCRIPTION: TStringField;
    ClientDataSet1: TClientDataSet;
    StringField1: TStringField;
    IntegerField1: TIntegerField;
    StringField2: TStringField;
    FloatField1: TFloatField;
    FloatField2: TFloatField;
    StringField3: TStringField;
    StringField4: TStringField;
    IntegerField2: TIntegerField;
    IntegerField3: TIntegerField;
    BooleanField1: TBooleanField;
    ClientDataSet2: TClientDataSet;
    IntegerField4: TIntegerField;
    FloatField3: TFloatField;
    FloatField4: TFloatField;
    cdsPackageInfoPackageID: TStringField;
    cdsPackageInfoTotalOrigionalPrice: TFloatField;
    cdsPackageInfoTotalRealCharge: TFloatField;
    cdsProductInfoProductID: TStringField;
    cdsProductInfoPackageID: TStringField;
    cdsProductInfoProviderID: TStringField;
    cdsProductInfoRealCharge: TFloatField;
    cdsProductInfoOrigionalPrice: TFloatField;
    cdsProductInfoFormulaID: TStringField;
    cdsProductInfoUseBaseFormula: TStringField;
    cdsProductInfoProvider_Percent: TIntegerField;
    cdsProductInfoSo_Percent: TIntegerField;
    cdsProductInfoIsInCome: TBooleanField;
    cdsProductInfoFirstFlag: TStringField;
    cdsSO116Rpt: TClientDataSet;
    cdsSO116RptCustID: TStringField;
    cdsSO116RptCustName: TStringField;
    cdsSO116RptChannel_String: TWideStringField;
    cdsSO114Rpt: TClientDataSet;
    cdsSO114RptPROVIDERDESC: TStringField;
    cdsSO114RptProductName: TStringField;
    cdsSO114RptFirstCount: TStringField;
    cdsSO114RptLastCount: TStringField;
    cdsSO114RptAddCount: TStringField;
    cdsSO114RptReduceCount: TStringField;
    cdsSO114RptIncome: TStringField;
    cdsSO114RptOutcome: TStringField;
    cdsSO114RptDiffCome: TStringField;
    cdsSO114RptEMC_Benefit: TStringField;
    cdsSO114RptSo_Benefit: TStringField;
    cdsSO114RptProvider_Benefit: TStringField;
    cdsSO114SubTotal: TClientDataSet;
    cdsSO114SubTotalFirstCount: TIntegerField;
    cdsSO114SubTotalLastCount: TIntegerField;
    cdsSO114SubTotalAddCount: TIntegerField;
    cdsSO114SubTotalReduceCount: TIntegerField;
    cdsSO114SubTotalIncome: TFloatField;
    cdsSO114SubTotalOutcome: TFloatField;
    cdsSO114SubTotalDiffCome: TFloatField;
    cdsSO114SubTotalEMC_Benefit: TFloatField;
    cdsSO114SubTotalSo_Benefit: TFloatField;
    cdsSO114SubTotalProvider_Benefit: TFloatField;
    cdsSO114Total: TClientDataSet;
    IntegerField5: TIntegerField;
    IntegerField6: TIntegerField;
    IntegerField7: TIntegerField;
    IntegerField8: TIntegerField;
    FloatField5: TFloatField;
    FloatField6: TFloatField;
    FloatField7: TFloatField;
    FloatField8: TFloatField;
    FloatField9: TFloatField;
    FloatField10: TFloatField;
    cdsCodeCD019: TClientDataSet;
    cdsSO033: TClientDataSet;
    cdsSO131: TClientDataSet;
    cdsSpSO131: TClientDataSet;
    cdsSO131COMPUTEYM: TStringField;
    cdsSO131REALORSHOULDDATE: TIntegerField;
    cdsSO131SEQNO: TIntegerField;
    cdsSO131BILLNO: TStringField;
    cdsSO131ITEM: TIntegerField;
    cdsSO131CUSTID: TIntegerField;
    cdsSO131CITEMCODE: TIntegerField;
    cdsSO131CITEMNAME: TStringField;
    cdsSO131SHOULDDATE: TDateTimeField;
    cdsSO131SHOULDAMT: TIntegerField;
    cdsSO131REALDATE: TDateTimeField;
    cdsSO131REALAMT: TIntegerField;
    cdsSO131REALPERIOD: TIntegerField;
    cdsSO131REALSTARTDATE: TDateTimeField;
    cdsSO131REALSTOPDATE: TDateTimeField;
    cdsSO131COMPCODE: TIntegerField;
    cdsSO131ORDERNO: TStringField;
    cdsSO131SBILLNO: TStringField;
    cdsSO131SITEM: TIntegerField;
    cdsSO131MEDIACODE: TIntegerField;
    cdsSO131MEDIANAME: TStringField;
    cdsSO131ACCEPTEN: TStringField;
    cdsSO131ACCEPTNAME: TStringField;
    cdsSO131INTROID: TStringField;
    cdsSO131INTRONAME: TStringField;
    cdsCd019A: TClientDataSet;
    cdsCd019CODENO: TIntegerField;
    cdsCd019DESCRIPTION: TStringField;
    cdsCd019REFDATA: TStringField;
    cdsSO131NOTES: TStringField;
    cdsSO131Excel: TClientDataSet;
    cdsSO131ExcelCOMPUTEYM: TStringField;
    cdsSO131ExcelREALORSHOULDDATE: TIntegerField;
    cdsSO131ExcelSEQNO: TIntegerField;
    cdsSO131ExcelBILLNO: TStringField;
    cdsSO131ExcelITEM: TIntegerField;
    cdsSO131ExcelCUSTID: TIntegerField;
    cdsSO131ExcelCITEMNAME: TStringField;
    cdsSO131ExcelSHOULDDATE: TDateTimeField;
    cdsSO131ExcelSHOULDAMT: TStringField;
    cdsSO131ExcelREALDATE: TDateTimeField;
    cdsSO131ExcelREALAMT: TStringField;
    cdsSO131ExcelREALPERIOD: TIntegerField;
    cdsSO131ExcelREALSTARTDATE: TDateTimeField;
    cdsSO131ExcelREALSTOPDATE: TDateTimeField;
    cdsSO131ExcelCOMPCODE: TIntegerField;
    cdsSO131ExcelORDERNO: TStringField;
    cdsSO131ExcelSBILLNO: TStringField;
    cdsSO131ExcelSITEM: TIntegerField;
    cdsSO131ExcelMEDIACODE: TIntegerField;
    cdsSO131ExcelMEDIANAME: TStringField;
    cdsSO131ExcelACCEPTEN: TStringField;
    cdsSO131ExcelACCEPTNAME: TStringField;
    cdsSO131ExcelINTROID: TStringField;
    cdsSO131ExcelINTRONAME: TStringField;
    cdsSO131ExcelNOTES: TStringField;
    cdsSO131ExcelCalcShouldAmt: TIntegerField;
    cdsSO131ExcelCalcRealAmt: TIntegerField;
    cdsSO131ExcelCITEMCODE: TIntegerField;
    cdsOtherSO131Excel: TClientDataSet;
    StringField5: TStringField;
    IntegerField9: TIntegerField;
    IntegerField10: TIntegerField;
    StringField6: TStringField;
    IntegerField11: TIntegerField;
    IntegerField12: TIntegerField;
    IntegerField13: TIntegerField;
    StringField7: TStringField;
    DateTimeField1: TDateTimeField;
    StringField8: TStringField;
    IntegerField14: TIntegerField;
    DateTimeField2: TDateTimeField;
    StringField9: TStringField;
    IntegerField15: TIntegerField;
    IntegerField16: TIntegerField;
    DateTimeField3: TDateTimeField;
    DateTimeField4: TDateTimeField;
    IntegerField17: TIntegerField;
    StringField10: TStringField;
    StringField11: TStringField;
    IntegerField18: TIntegerField;
    IntegerField19: TIntegerField;
    StringField12: TStringField;
    StringField13: TStringField;
    StringField14: TStringField;
    StringField15: TStringField;
    StringField16: TStringField;
    StringField17: TStringField;
    cdsSO131ExcelREFNO: TIntegerField;
    cdsOtherSO131ExcelREFNO: TIntegerField;
    cdsSO131ExcelCalcRefNo: TStringField;
    cdsOtherSO131ExcelCalcRefNo: TStringField;
    cdsSO113ForUSeful: TClientDataSet;
    cdsPercentXmlData: TClientDataSet;
    cdsPercentXmlDatasProviderName: TStringField;
    cdsPercentXmlDatanPercent: TIntegerField;
    cdsSo112A: TClientDataSet;
    cdsPercentXmlDatasProviderID: TStringField;
    cdsSO131ExcelCOMPUTEYM2: TStringField;
    cdsOtherSO131ExcelCOMPUTEYM2: TStringField;
    cdsSO131ExcelSHOULDDATE2: TStringField;
    cdsSO131ExcelREALDATE2: TStringField;
    cdsSO131ExcelREALSTARTDATE2: TStringField;
    cdsSO131ExcelREALSTOPDATE2: TStringField;
    cdsOtherSO131ExcelSHOULDDATE2: TStringField;
    cdsOtherSO131ExcelREALDATE2: TStringField;
    cdsOtherSO131ExcelREALSTARTDATE2: TStringField;
    cdsOtherSO131ExcelREALSTOPDATE2: TStringField;
    procedure cdsSo113CTCalcFields(DataSet: TDataSet);
    procedure cdsSo110CalcFields(DataSet: TDataSet);
    procedure cdsSo111DCalcFields(DataSet: TDataSet);
    procedure cdsSo112MAfterScroll(DataSet: TDataSet);
    procedure cdsSO131ExcelCalcFields(DataSet: TDataSet);
    procedure cdsOtherSO131ExcelCalcFields(DataSet: TDataSet);
    procedure DataModuleCreate(Sender: TObject);
    procedure cdsPercentXmlDataBeforeScroll(DataSet: TDataSet);
    procedure cdsPercentXmlDataAfterPost(DataSet: TDataSet);
    procedure cdsSo112AAfterScroll(DataSet: TDataSet);

  private
    { Private declarations }
    function isNeedCitemCode(nI_CitemCode : Integer) : Boolean;
    procedure initCdsAttribute;


  public
    { Public declarations }
    //sG_IsSupervisor,sG_UserID,sG_User,sG_CompCode,sG_CompName : String;
    //sG_DbUserName,sG_DbPassword,sG_DbAlias,sG_ServiceType,sG_Server_IP : String;


    sG_ChannelViewCountsString : String;

    bG_IsTransData : Boolean;
    nG_TotalValidViewCounts : Integer;
    nG_TotalCustCounts,nG_TotalCustCounts1,nG_TotalCustCounts2,nG_TotalCustCounts3 : Integer;
    G_ChannelViewStrList : TStringList;
    G_ComputeParaStrList,G_CompareNoStrList,G_ProviderStrList : TStringList;
    //Rpt
    procedure transXmlToCds;
    function genXmlDataStr:String;
    function genXmlDataStr1:String;
            
    procedure initialData;
    procedure createCDSActive;
    procedure getRptSQL(I_CDS:TClientDataset; sI_MID :String; var I_SQLStrList:TStringList);overload;
    procedure saveToDB(I_CDS:TClientDataSet);

    procedure getCdsSqlDelete_So111Result;

    procedure saveToFile(vI_Content:OleVariant; sI_FileName:String);
    function ReplaceStr(sI_SrcString, sI_SepFlag: String): String;
    procedure initPercentXmlData;
    //frmSO8A40
    function checkSO114HaveBelongYM(sI_CompCode,sI_BelongYM : String) : Boolean;
    function chineseDateChangeToEnglishDate(sI_Date : String) : String;
    function chineseYMChangeToEnglishYM(sI_Date : String) : String;
    function calculateData(sI_CompCode,sI_StartDate,sI_EndDate,sI_BelongYM,sI_ShowDetail : String) : String; //璸衡戈
    procedure getTotalViewCustCounts(sI_Table,sI_CompCode : String); //璸衡羆ΤΜ跌めの繵笵ぇ羆Μ跌め计
    function getCitemData(nI_Function : Integer) : String;
    function getCustData(sI_Table,sI_CompCode,sI_StartDate,sI_EndDate,sI_CitemString : String) : Integer;
    function getChargeData(sI_Table,sI_CustID,sI_CompCode,sI_StartDate,sI_EndDate,sI_CitemString : String) : Boolean;
    function getProductID(sI_CitemCode : String) : String;
    function getSO112Data(sI_ProductID : String;var sI_ProviderID,sI_UseBaseFormula : String;var fI_OrigionalPrice,fI_ProviderPercent,fI_SoPercent : Double) : Boolean;
    procedure saveTocdsProductInfo(sI_FirstFlag,sI_ProductID,sI_PackageID,sI_FormulaID,sI_ProviderID,sI_UseBaseFormula : String;fI_RealCharge,fI_OrigionalPrice,fI_ProviderPercent,fI_SoPercent : Double;bI_IsInCome : Boolean);
    procedure saveTocdsPackageInfo(sI_PackageID : String;fI_RealCharge,fI_OrigionalPrice : Double);
    function calculateSoOrProviderAmt(sI_ProductID,sI_PackageID,sI_FormulaID,sI_CompType,sI_CompOrProviderID : String;fI_OrigionalPrice : Double) : Double;
    procedure deleteReportData(sI_compcode,sI_BelongYM : String);
    procedure getReportData(sI_compcode,sI_BelongYM : String);
    procedure handleShareAmt1(sI_CompCode,sI_CustID,sI_StartDate,sI_EndDate,sI_BelongYM,sI_CurDate,sI_ShowDetail,sI_PrvBelongYM : String;var sI_ChannelString : String);
    function getChannelName(sI_ProductID : String) : String;
    procedure saveTocdsSo116(sI_CompCode,sI_BelongYM,sI_CustID,sI_CurDate,sI_ChannelString : String);
    procedure saveTocdsSo117(sI_CompCode,sI_BelongYM,sI_CurDate,sI_ChannelViewCountsString : String;nI_CustCounts : Integer);
    function getCustName(sI_CustID : String) : String;
    procedure getSO114Data(sI_compcode,sI_BelongYM : String);
    function getProviderName(sI_ProviderID : String) : String;
    procedure gerAllProvider;
    function getProviderBenefit(sI_compcode,sI_BelongYM,sI_ProductID,sI_ProviderID : String) : Double;
    function getProviderCounts(sI_compcode,sI_BelongYM,sI_ProductID,sI_ProviderID : String) : Integer;
    function functionBasic(sI_CurDate,sI_CompCode,sI_StartDate,sI_EndDate,sI_BelongYM,sI_ShowDetail : String) : String;
    function functionOne(sI_CurDate,sI_CompCode,sI_StartDate,sI_EndDate,sI_BelongYM,sI_ShowDetail : String) : String;
    function functionTwo(sI_CurDate,sI_CompCode,sI_StartDate,sI_EndDate,sI_BelongYM,sI_ShowDetail : String) : String;
    function functionThree(sI_CurDate,sI_CompCode,sI_StartDate,sI_EndDate,sI_BelongYM,sI_ShowDetail : String) : String;
    function getChannelCustData(sI_CompCode,sI_Table,sI_CitemCode,sI_StartDate,sI_EndDate : String) : Integer;
    function handleChannelCustData(sI_CompCode,sI_BelongYM,sI_CurDate,sI_ChannelName : String;var nI_TotalSingleCounts,nI_TotalPackageCounts : Integer) : Double;
    procedure getSharePara(sI_CompCode,sI_ProductID,sI_ProviderID : String);
    function handleShareAmt2(sI_FormulaID : String;nI_ChannelTotalViewCounts : Integer;fI_OrigionalPrice : Double;var fI_Provider_Benefit,fI_So_Benefit : Double) : String;
    function calculateSoOrProviderAmt2(sI_SoOrProvider,sI_FormulaID : String;nI_ChannelTotalViewCounts : Integer;fI_OrigionalPrice : Double;var fI_SoOrProviderBenefit : Double): String;
    function getFormulaRefNo(sI_FormulaID : String) : String;
    procedure getCalculateNo(sI_RangeUnit,sI_SoOrProvider,sI_ComputeIssue : String;nI_ChannelTotalViewCounts : Integer);
    function calculateMulti(nI_ChannelTotalViewCounts : Integer;fI_OrigionalPrice : Double;bI_IsPercent,bI_IsPrice : Boolean) : Double;  //璸衡仓秈禯
    function calculateSingle(nI_ChannelTotalViewCounts : Integer;fI_OrigionalPrice : Double;bI_IsPrice : Boolean) : Double; //璸衡辅翴禯
    procedure saveToSo118(sI_CompCode,sI_BelongYM ,sI_CurDate: String);
    function checkBelongYMIsLock(sI_CompCode,sI_BelongYM : String) : Boolean;
    procedure saveToSO119(sI_CompCode,sI_BelongYM,sI_ProductID,sI_ProviderID,sI_CurDate : String;fI_ProviderPercent,fI_SoPercent : Double);
    function checkIsInCome(sI_TempCitemCode : String;var sI_CitemCode : String) : Boolean;
    function getPrvBelongYM(sI_BelongYM : String) : String;
    procedure getLastMonthComputePara(sI_CompCode,sI_PrvBelongYM,sI_ProductID : String;var fI_OldProviderPercent,fI_OldSoPercent : Double);
    function getOutComeCitemData(sI_CitemString : String) : String;
    function getFirstCounts(sI_CompCode, sI_PrvBelongYM,sI_ProductID : String) : Integer;
    procedure saveOtherDataTocdsSo114(sI_BelongYM,sI_PrvBelongYM,sI_CompCode, sI_StartDate, sI_EndDate,sI_TotalCitemString: String);
    function getTotalInOrOutCome(sI_CompCode, sI_StartDate, sI_EndDate,sI_CitemCode : String;var nI_TotalInOrOutComeCounts :Integer) : Double;


    function getSO110RefNo(sI_Formula_ID : String) : String;
    procedure getSO113AndCD039(sI_ProviderID : String);
    function getProductName(sI_ProductID : String) : String;
    procedure getCodeCD019;
    function calculateDetailData(sI_CompCode,sI_ChargeItemSQL,sI_ComputeYM : String;nI_RealOrShouldDate : Integer) : String;
    function TransToEngDate(sI_Date : String) : String;
    procedure getSO105Data(sI_OrderNo : String;var sI_MediaCode,sI_MediaName,sI_AcceptEn,sI_AcceptName,sI_IntroId,sI_IntroName : String);
    function checkHaveSO131Data(sI_ComputeYM : String;nI_RealOrShouldDate : Integer) : Boolean;
    function runSFDetail(sI_ComputeYM,sI_ChargeItemSQL,sI_StartDate,sI_EndDate : String;nI_RealOrShouldDate : Integer;var fI_RetCode : Double;var sI_Msg : String) : String;
    function getBackCitemCode : String;
    function getBackDetailData(sI_Table,sI_CompCode, sI_ComputeYM : String; nI_RealOrShouldDate: Integer) : String;
    function calculateCompleteDetailData(sI_Table,sI_ComputeYM : String;nI_RealOrShouldDate : Integer) : String;
    procedure addSeqNo;
    procedure getDetailData(sI_CompCode,sI_ComputeYM : String;nI_RealOrShouldDate : Integer);
    procedure connToServer;
  end;

var
  dtmMain2: TdtmMain2;
      bPanelControlG:Boolean;
implementation

uses frmSO8A30U, Emc_Separate_TLB, Ustru, UCommonU, UdateTimeu, frmSO8A50U,
  frmMainMenuU, xmlU;

{$R *.dfm}

{ TfrmDM }
procedure TdtmMain2.getCdsSqlDelete_So111Result;
var sL_SQL:String;
begin
    with dtmMain2.cdsSqlDelete do
    begin
      Close;
      sL_SQL:='DELETE from SO111 where PRODUCT_ID='''+dtmMain2.cdsSo112M.FieldByName('PRODUCT_ID').AsString+'''';
      CommandText:=sL_SQL;
      //CommandText:='DELETE from SO111 where PRODUCT_ID=:PRODUCT_ID';
      //PARAMS.ParamByName('PRODUCT_ID').AsString:=dtmMain2.cdsSo112M.FieldByName('PRODUCT_ID').AsString;
      Execute;
    end;
end;

procedure TdtmMain2.cdsSo113CTCalcFields(DataSet: TDataSet);
begin
    if cdsSo113CT.FieldByName('ATTRIBUTE').AsString='1' then
       cdsSo113CT.FieldByName('ATTRIBUTE_NAME').AsString:='PPV Provider'
    else if cdsSo113CT.FieldByName('ATTRIBUTE').AsString='2' then
       cdsSo113CT.FieldByName('ATTRIBUTE_NAME').AsString:='Pay Channel Provider'
    else if cdsSo113CT.FieldByName('ATTRIBUTE').AsString='3' then
       cdsSo113CT.FieldByName('ATTRIBUTE_NAME').AsString:='PPV and Pay Channel Provider';

    if cdsSo113CT.FieldByName('STOPFLAG').AsString='0' then
       cdsSo113CT.FieldByName('STOPFLAG_NAME').AsString:=''
    else if cdsSo113CT.FieldByName('STOPFLAG').AsString='1' then
       cdsSo113CT.FieldByName('STOPFLAG_NAME').AsString:='琌';
end;

procedure TdtmMain2.cdsSo110CalcFields(DataSet: TDataSet);
var nL_Ref_No:Integer;
begin
    if cdsSo110.FieldByName('Ref_No').AsString<>'' then
    begin
      nL_Ref_No:=StrToInt(cdsSo110.FieldByName('Ref_No').AsString);
      case nL_Ref_No of
        1:cdsSo110.FieldByName('REFNO_NAME').AsString:='ㄢ顶甅繺( ゑㄒ )';
        2:cdsSo110.FieldByName('REFNO_NAME').AsString:='め计计秖禯( 虫基 )';
        3:cdsSo110.FieldByName('REFNO_NAME').AsString:='め计计秖禯( ゑㄒ )';
        4:cdsSo110.FieldByName('REFNO_NAME').AsString:='め计κだゑ禯( 虫基 )';
        5:cdsSo110.FieldByName('REFNO_NAME').AsString:='め计κだゑ禯( ゑㄒ )';
      end;
   end;
   if cdsSo110.FieldByName('STOPFLAG').AsString='0' then
     cdsSo110.FieldByName('STOPFLAG_NAME').AsString:=''
   else if cdsSo110.FieldByName('STOPFLAG').AsString='1' then
     cdsSo110.FieldByName('STOPFLAG_NAME').AsString:='琌';
end;


procedure TdtmMain2.cdsSo111DCalcFields(DataSet: TDataSet);
begin
    if cdsSo111D.FieldByName('RANGE_UNIT').AsString='1' then
       cdsSo111D.FieldByName('RANGE_UNIT_NAME').AsString:='め计'
    else if cdsSo111D.FieldByName('RANGE_UNIT').AsString='2' then
       cdsSo111D.FieldByName('RANGE_UNIT_NAME').AsString:='κだゑ';
end;

procedure TdtmMain2.cdsSo112MAfterScroll(DataSet: TDataSet);
begin
    frmSO8A30.CursorMoveState;
end;





function TdtmMain2.ReplaceStr(sI_SrcString, sI_SepFlag: String): String;
var
    ii : Integer;
    sL_Result : String;
begin
    sL_Result := '';
    for ii:=1 to Length(sI_SrcString) do
    begin
      if sI_SrcString[ii]<>sI_SepFlag then
        sL_Result := sL_Result + sI_SrcString[ii]
      else
        Continue;
    end;
    Result := sL_Result;

end;

function TdtmMain2.checkSO114HaveBelongYM(sI_CompCode,sI_BelongYM : String): Boolean;
var
    sL_SQL : String;
begin
    sL_SQL := 'SELECT * FROM SO114 WHERE BelongYM=''' + sI_BelongYM + '''';

    with cdsCom do
    begin
      Close;
      CommandText := sL_SQL;
      Open;

      if RecordCount > 0 then
        Result := true
      else
        Result := false;
    end;
end;


function TdtmMain2.chineseDateChangeToEnglishDate(sI_Date: String): String;
var
    L_StringList : TStringList;
    sL_Year,sL_Date : String;
begin
    L_StringList := TStringList.Create;
    L_StringList := TUstr.ParseStrings(sI_Date,'/');

    sL_Year  := IntToStr(StrToInt(L_StringList.Strings[0]) + 1911);

    sL_Date := sL_Year + '/' + L_StringList.Strings[1] + '/' + L_StringList.Strings[2];

    Result := sL_Date;
end;

procedure TdtmMain2.saveToFile(vI_Content: OleVariant; sI_FileName: String);
var
    L_StrList : TStringList;
begin
    L_StrList := TStringList.Create;;
    L_StrList.Add(vI_Content);
    L_StrList.SaveToFile(sI_FileName);
    L_StrList.Free;
end;

function TdtmMain2.calculateData(sI_CompCode,sI_StartDate,sI_EndDate,sI_BelongYM,sI_ShowDetail : String) : String;
var
    sL_CurDate,sL_Table,sL_ErrorMsg : String;
    ii : Integer;
begin

    sL_CurDate := DateTimeToStr(Now);
    //睲
    cdsSO114.EmptyDataSet;
    cdsSO115.EmptyDataSet;
    cdsSO116.EmptyDataSet;
    cdsSO117.EmptyDataSet;
    nG_TotalCustCounts := 0;

    //璸衡羆ΤΜ跌めの繵笵ぇ羆Μ跌め计
    getTotalViewCustCounts(TABLE_SO033,sI_CompCode);
    getTotalViewCustCounts(TABLE_SO034,sI_CompCode);

    //璸衡膀セそΑ
    sL_ErrorMsg := functionBasic(sL_CurDate,sI_CompCode,sI_StartDate,sI_EndDate,sI_BelongYM,sI_ShowDetail);
    if sL_ErrorMsg <> '' then
      Result := sL_ErrorMsg;

    //璸衡そΑ
    sL_ErrorMsg := functionOne(sL_CurDate,sI_CompCode,sI_StartDate,sI_EndDate,sI_BelongYM,sI_ShowDetail);
    if sL_ErrorMsg <> '' then
      Result := sL_ErrorMsg;

    //璸衡そΑの
    sL_ErrorMsg := functionTwo(sL_CurDate,sI_CompCode,sI_StartDate,sI_EndDate,sI_BelongYM,sI_ShowDetail);
    if sL_ErrorMsg <> '' then
      Result := sL_ErrorMsg;

    //璸衡そΑのき
    sL_ErrorMsg := functionThree(sL_CurDate,sI_CompCode,sI_StartDate,sI_EndDate,sI_BelongYM,sI_ShowDetail);
    if sL_ErrorMsg <> '' then
      Result := sL_ErrorMsg;


    //盢戈╊眀穨ぇ羆㎝戈
    saveTocdsSo117(sI_CompCode,sI_BelongYM,sL_CurDate,sG_ChannelViewCountsString,nG_TotalCustCounts);
end;

procedure TdtmMain2.getTotalViewCustCounts(sI_Table,sI_CompCode : String);
var
    sL_SQL,sL_CitemCode,sL_ProductID,sL_ChannelName : String;
    L_Obj : TChannelViewData;
    nL_ChannelTotalViewCounts,nL_TempViewCounts : Integer;

begin
    //斗暗Ω珿讽 Query 癸禜琌 SO033 琩高
    if sI_Table = TABLE_SO033 then
    begin
      //璸衡羆ΤΜ跌め
      sL_SQL := 'SELECT COUNT(*) COUNTS FROM SO002 WHERE CUSTSTATUSCODE=1' +
                ' AND COMPCODE=' + sI_CompCode +
                ' AND SERVICETYPE=''' + SERVICE_TYPE + '''';

      with cdsCom do
      begin
        Close;
        CommandText := sL_SQL;
        Open;

        nG_TotalValidViewCounts := cdsCom.FieldByName('COUNTS').AsInteger;
        if nG_TotalValidViewCounts <= 0 then
        begin
          MessageDlg('⊿ΤΜ跌め戈',mtInformation, [mbOK],0);
          exit;
        end;
      end;
    end;


    //璸衡繵笵ぇ羆Μ跌め计
    G_ChannelViewStrList := TStringList.Create;
    //cdsChannelCounts.EmptyDataSet;

    sL_SQL := 'SELECT A.PRODUCT_ID,A.REF_NO,B.CITEMCODE ' +
              'FROM SO112 A,CD019A B WHERE B.CODENO=A.PRODUCT_ID';

    cdsRefData.Close;
    cdsRefData.CommandText := sL_SQL;
    cdsRefData.Open;

    cdsRefData.First;
    while not cdsRefData.Eof do
    begin
      sL_CitemCode := cdsRefData.FieldByName('CITEMCODE').AsString;
      sL_ProductID := cdsRefData.FieldByName('PRODUCT_ID').AsString;
      sL_ChannelName := getChannelName(sL_ProductID);

      //sL_SQL := 'SELECT COUNT(*) COUNTS FROM V_ChargeInfoForTally WHERE ' +
      //V_ChargeInfoForTally эだ癸 SO033 の SO034 穨
      sL_SQL := 'SELECT COUNT(*) COUNTS FROM ' + sI_Table + ' WHERE ' +
                'CITEMCODE=' + sL_CitemCode + ' AND COMPCODE=' + sI_CompCode;

      with cdsCom do
      begin
        Close;
        CommandText := sL_SQL;
        Open;
        nL_ChannelTotalViewCounts := FieldByName('COUNTS').AsInteger;

        //рSO033のSO034戈仓既cdsChannelCountsい
        if cdsChannelCounts.Locate('ProductID',sL_ProductID,[]) then
        begin
          nL_TempViewCounts := cdsChannelCounts.FieldByName('ChannelTotalViewCounts').AsInteger;

          cdsChannelCounts.Delete;
          cdsChannelCounts.Append;

          cdsChannelCounts.FieldByName('ProductID').AsString := sL_ProductID;
          cdsChannelCounts.FieldByName('CitemCode').AsString := sL_CitemCode;
          cdsChannelCounts.FieldByName('ChannelName').AsString := sL_ChannelName;
          cdsChannelCounts.FieldByName('ChannelTotalViewCounts').AsInteger := nL_TempViewCounts + nL_ChannelTotalViewCounts;

          cdsChannelCounts.Post;
        end
        else
        begin
          cdsChannelCounts.Append;

          cdsChannelCounts.FieldByName('ProductID').AsString := sL_ProductID;
          cdsChannelCounts.FieldByName('CitemCode').AsString := sL_CitemCode;
          cdsChannelCounts.FieldByName('ChannelName').AsString := sL_ChannelName;
          cdsChannelCounts.FieldByName('ChannelTotalViewCounts').AsInteger := nL_ChannelTotalViewCounts;

          cdsChannelCounts.Post;
        end;
      end;

      cdsRefData.Next;
    end;

    //矪瞶仓 cdsChannelCounts
    //( 斗暗Ω珿讽 Query 癸禜琌 SO034 琩高 )
    if sI_Table = TABLE_SO034 then
    begin
      cdsChannelCounts.First;
      while not cdsChannelCounts.Eof do
      begin

        sL_ProductID := cdsChannelCounts.FieldByName('ProductID').AsString;
        sL_ChannelName := cdsChannelCounts.FieldByName('ChannelName').AsString;
        sL_CitemCode := cdsChannelCounts.FieldByName('CitemCode').AsString;
        nL_ChannelTotalViewCounts := cdsChannelCounts.FieldByName('ChannelTotalViewCounts').AsInteger;

        L_Obj := TChannelViewData.Create;
        L_Obj.sProduct_ID := sL_ProductID;
        L_Obj.sCitemCode := sL_CitemCode;
        L_Obj.nChannelTotalViewCounts := nL_ChannelTotalViewCounts;

        if sG_ChannelViewCountsString = '' then
          sG_ChannelViewCountsString := sL_ProductID + '~' + sL_ChannelName + '~' + IntToStr(nL_ChannelTotalViewCounts)
        else
          sG_ChannelViewCountsString := sG_ChannelViewCountsString + ';' + sL_ProductID + '~' +sL_ChannelName + '~' + IntToStr(nL_ChannelTotalViewCounts);

        G_ChannelViewStrList.AddObject(sL_ProductID, L_Obj);

        cdsChannelCounts.Next;
      end;
    end;
end;



function TdtmMain2.getCitemData(nI_Function : Integer) : String;
var
    sL_SQL,sL_SQLTitle,sL_CitemCode,sL_CitemString,sL_TempCitemCode,sL_Where : String;
    L_CitemString : TStringList;
    ii : Integer;
begin
    sL_SQLTitle := 'SELECT DISTINCT A.PRODUCT_ID,A.REF_NO,B.CITEMCODE ' +
              'FROM SO112 A,CD019A B';

    if nI_Function = FUNCTION1 then  //そΑ
      sL_Where :=  ',SO110 C WHERE B.CODENO=A.PRODUCT_ID AND A.REF_NO=1 AND C.STOPFLAG=0 AND A.USEBASEFORMULA=''N'''
    else if nI_Function = FUNCTION2 then  //そΑの
      sL_Where :=  ',SO110 C WHERE B.CODENO=A.PRODUCT_ID AND A.REF_NO<>1 AND C.REF_NO IN (''2'',''4'') AND C.STOPFLAG=0 AND A.USEBASEFORMULA=''N'''
    else if nI_Function = FUNCTION3 then  //そΑのき
      sL_Where :=  ',SO110 C WHERE B.CODENO=A.PRODUCT_ID AND A.REF_NO<>1 AND C.REF_NO IN (''3'',''5'') AND C.STOPFLAG=0 AND A.USEBASEFORMULA=''N'''
    else if nI_Function = FUNCTION0 then  //膀セそΑ
      sL_Where :=  ' WHERE B.CODENO=A.PRODUCT_ID AND A.USEBASEFORMULA=''Y''';

    sL_SQL := sL_SQLTitle + sL_Where;

    with cdsRefDataA do
    begin
      Close;
      CommandText := sL_SQL;
      Open;
    end;


    //盢Μ禣兜ヘ絏﹃Θ﹃
    sL_CitemString := '';
    L_CitemString := TStringList.Create;
    cdsRefDataA.First;
    while not cdsRefDataA.Eof do
    begin
      sL_CitemCode := cdsRefDataA.FieldByName('CITEMCODE').AsString;

      if sL_CitemString = '' then
      begin
        L_CitemString.Add(sL_CitemCode);
        sL_CitemString := sL_CitemCode;
      end
      else
      begin
        for ii:=0 to L_CitemString.count-1 do
        begin
          sL_TempCitemCode := L_CitemString.Strings[ii];
          if sL_TempCitemCode = sL_CitemCode then
            break;

          if ii = L_CitemString.count-1 then
          begin
            if sL_TempCitemCode <> sL_CitemCode then
            begin
              L_CitemString.Add(sL_CitemCode);

              sL_CitemString := sL_CitemString + ',' + sL_CitemCode;
            end;
          end;
        end;
      end;

      cdsRefDataA.Next;
    end;

    Result := sL_CitemString;
end;

function TdtmMain2.getCustData(sI_Table,sI_CompCode, sI_StartDate, sI_EndDate,
  sI_CitemString: String): Integer;
var
    sL_SQL : String;
begin
    //眖 view V_ChargeInfoForTally いт龟Μら戳絛瞅ず
    //禣兜ヘ絏(CITEMCODE)妮cdsRefDataA┮Τめ絪腹
    //sL_SQL := 'SELECT DISTINCT CUSTID FROM V_ChargeInfoForTally WHERE CITEMCODE IN (' + sI_CitemString + ')' +
    //V_ChargeInfoForTally эだ癸 SO033 の SO034 穨
    sL_SQL := 'SELECT DISTINCT CUSTID FROM ' + sI_Table + ' WHERE CITEMCODE IN (' + sI_CitemString + ')' +
              ' AND (REALDATE BETWEEN ' + TUstr.getOracleSQLDateStr(StrToDate(sI_StartDate)) +
              ' AND ' + TUstr.getOracleSQLDateStr(StrToDate(sI_EndDate)) + ')' +
              ' AND COMPCODE=' + sI_CompCode;

    with cdsCustData do
    begin
      Close;
      CommandText := sL_SQL;
      Open;
      Result := RecordCount;
    end;

end;

function TdtmMain2.getChargeData(sI_Table,sI_CustID, sI_CompCode, sI_StartDate,
  sI_EndDate, sI_CitemString: String): Boolean;
var
    sL_SQL : String;
begin
    //V_ChargeInfoForTally эだ癸 SO033 の SO034 穨
    //sL_SQL := 'SELECT * FROM V_ChargeInfoForTally WHERE CITEMCODE IN (' + sI_CitemString + ')' +
    sL_SQL := 'SELECT * FROM ' + sI_Table + ' WHERE CITEMCODE IN (' + sI_CitemString + ')' +
              ' AND (REALDATE BETWEEN ' + TUstr.getOracleSQLDateStr(StrToDate(sI_StartDate)) +
              ' AND ' + TUstr.getOracleSQLDateStr(StrToDate(sI_EndDate)) + ')' +
              ' AND CUSTID=' + sI_CustID + ' AND COMPCODE=' + sI_CompCode;

    with cdsChargeData do
    begin
      Close;
      CommandText := sL_SQL;
      Open;

      if RecordCount > 0 then
        Result := true
      else
        Result := false;
    end;
end;

function TdtmMain2.getProductID(sI_CitemCode: String): String;
begin
    //パ CitemCode 癸莱 cdsRefDataA 眔繵笵絏 ( ProductID )
    if cdsRefDataA.Locate('CITEMCODE',sI_CitemCode,[]) then
      Result := cdsRefDataA.FieldByName('PRODUCT_ID').AsString;
end;



function TdtmMain2.getSO112Data(sI_ProductID: String; var
  sI_ProviderID,sI_UseBaseFormula: String; var fI_OrigionalPrice,fI_ProviderPercent,fI_SoPercent: Double): Boolean;
var
    sL_SQL : String;
begin
    sL_SQL := 'SELECT * FROM SO112 WHERE PRODUCT_ID=''' + sI_ProductID + '''';

    with cdsCom do
    begin
      Close;
      CommandText := sL_SQL;
      Open;


      if RecordCount > 0 then
      begin
        sI_ProviderID := FieldByName('PROVIDER_ID').AsString;
        fI_OrigionalPrice := FieldByName('PRICE').AsFloat;

        sI_UseBaseFormula := FieldByName('USEBASEFORMULA').AsString;
        fI_ProviderPercent := FieldByName('PROVIDER_PERCENT').AsFloat;
        fI_SoPercent := FieldByName('SO_PERCENT').AsFloat;

        Result := true;
      end
      else
      begin
        MessageDlg('⊿Τ才兵ン繵笵戈<' +  sI_ProductID + '>',mtInformation, [mbOK],0);
        Result := false;
        exit;
      end;

    end;
end;

procedure TdtmMain2.saveTocdsProductInfo(sI_FirstFlag,sI_ProductID, sI_PackageID,
  sI_FormulaID,sI_ProviderID,sI_UseBaseFormula: String; fI_RealCharge, fI_OrigionalPrice,fI_ProviderPercent,fI_SoPercent: Double;bI_IsInCome : Boolean);
begin
    with cdsProductInfo do
    begin
      Append;

      FieldByName('ProductID').AsString := sI_ProductID;
      FieldByName('PackageID').AsString := sI_PackageID;
      FieldByName('RealCharge').AsFloat := fI_RealCharge;
      FieldByName('OrigionalPrice').AsFloat := fI_OrigionalPrice;
      FieldByName('FormulaID').AsString := sI_FormulaID;
      FieldByName('ProviderID').AsString := sI_ProviderID;

      FieldByName('UseBaseFormula').AsString := sI_UseBaseFormula;
      FieldByName('Provider_Percent').AsFloat := fI_ProviderPercent;
      FieldByName('So_Percent').AsFloat := fI_SoPercent;

      FieldByName('IsInCome').AsBoolean := bI_IsInCome;
      FieldByName('FirstFlag').AsString := sI_FirstFlag;

      Post;
    end;
end;

procedure TdtmMain2.saveTocdsPackageInfo(sI_PackageID: String; fI_RealCharge,
  fI_OrigionalPrice: Double);
var
    fL_TotalRealCharge,fL_TotalOrigionalPrice : Double;
begin
    //璝掸data妮甅繺戈, 玥Update(or Insert) cdsPackageInfo
    if sI_PackageID <> '' then
    begin
      //安璝赣甅繺,玥仓肂
      if cdsPackageInfo.Locate('PACKAGEID',sI_PackageID,[]) then
      begin
        with cdsPackageInfo do
        begin
          fL_TotalRealCharge := FieldByName('TotalRealCharge').AsFloat;
          fL_TotalOrigionalPrice := FieldByName('TotalOrigionalPrice').AsFloat;
          Delete;
          Append;
          fL_TotalRealCharge := fL_TotalRealCharge + fI_RealCharge;
          fL_TotalOrigionalPrice := fL_TotalOrigionalPrice + fI_OrigionalPrice;
          FieldByName('PackageID').AsString := sI_PackageID;
          FieldByName('TotalRealCharge').AsFloat := fL_TotalRealCharge;
          FieldByName('TotalOrigionalPrice').AsFloat := fL_TotalOrigionalPrice;
          Post;
        end;
      end
      else //安璝赣甅繺ぃ,玥穝糤掸戈
      begin
        with cdsPackageInfo do
        begin
          Append;
          FieldByName('PackageID').AsString := sI_PackageID;
          FieldByName('TotalRealCharge').AsFloat := fI_RealCharge;
          FieldByName('TotalOrigionalPrice').AsFloat := fI_OrigionalPrice;
          Post;
        end;
      end;
    end;
end;

function TdtmMain2.calculateSoOrProviderAmt(sI_ProductID,sI_PackageID,sI_FormulaID,sI_CompType,sI_CompOrProviderID: String;fI_OrigionalPrice : Double): Double;
var
    nL_Ndx,nL_ChannelTotalViewCounts : Integer;
    fL_Subscriber : Double;
    sL_SQL,sL_RangeUnit : String;
    fL_Subscriber_Count_Percent1,fL_Subscriber_Count_Percent2 : Double;
    fL_Subscriber_Count_Percent3,fL_Subscriber_Count_Percent4 : Double;
    fL_Subscriber_Count_Percent5 : Double;
    fL_Amount_Percent1,fL_Amount_Percent2,fL_Amount_Percent3 : Double;
    fL_Amount_Percent4,fL_Amount_Percent5 : Double;
    fL_ComputerPara,fL_CountsRate : Double;

    fL_TotalOrigionalPrice,fL_TotalRealCharge : Double;
begin
    //繵笵ぇ羆Μ跌め计
    nL_Ndx := G_ChannelViewStrList.IndexOf(sI_ProductID);
    if nL_Ndx<>-1 then
       nL_ChannelTotalViewCounts := (G_ChannelViewStrList.Objects[nL_Ndx] as TChannelViewData).nChannelTotalViewCounts;

    sL_SQL := 'SELECT * FROM SO111 WHERE PRODUCT_ID=''' + sI_ProductID +
              ''' AND COMP_TYPE=''' +
              sI_CompType + ''' AND So_Provider_ID=''' + sI_CompOrProviderID + '''';


    with cdsCom do
    begin
      Close;
      CommandText := sL_SQL;
      Open;

      //眔赣繵笵め计禯
      sL_RangeUnit := FieldByName('RANGE_UNIT').AsString;

      fL_Subscriber_Count_Percent1 := FieldByName('Subscriber_Count_Percent1').AsFloat;
      fL_Subscriber_Count_Percent2 := FieldByName('Subscriber_Count_Percent2').AsFloat;
      fL_Subscriber_Count_Percent3 := FieldByName('Subscriber_Count_Percent3').AsFloat;
      fL_Subscriber_Count_Percent4 := FieldByName('Subscriber_Count_Percent4').AsFloat;
      fL_Subscriber_Count_Percent5 := FieldByName('Subscriber_Count_Percent5').AsFloat;

      fL_Amount_Percent1 := FieldByName('Amount_Percent1').AsFloat;
      fL_Amount_Percent2 := FieldByName('Amount_Percent2').AsFloat;
      fL_Amount_Percent3 := FieldByName('Amount_Percent3').AsFloat;
      fL_Amount_Percent4 := FieldByName('Amount_Percent4').AsFloat;
      fL_Amount_Percent5 := FieldByName('Amount_Percent5').AsFloat;

      if sL_RangeUnit = '1' then  //め计ゑ耕虫 : め计
      begin
        //め计ゑ耕
        if nL_ChannelTotalViewCounts <= fL_Subscriber_Count_Percent1 then
          fL_ComputerPara := fL_Amount_Percent1
        else if nL_ChannelTotalViewCounts <= fL_Subscriber_Count_Percent2 then
          fL_ComputerPara := fL_Amount_Percent2
        else if nL_ChannelTotalViewCounts <= fL_Subscriber_Count_Percent3 then
          fL_ComputerPara := fL_Amount_Percent3
        else if nL_ChannelTotalViewCounts <= fL_Subscriber_Count_Percent4 then
          fL_ComputerPara := fL_Amount_Percent4
        else if nL_ChannelTotalViewCounts <= fL_Subscriber_Count_Percent5 then
          fL_ComputerPara := fL_Amount_Percent5;

        //虫め计璸衡 = 繵笵ぇ羆Μ跌め计 * 璸衡把计
        Result := nL_ChannelTotalViewCounts * fL_ComputerPara;

      end
      else if sL_RangeUnit = '2' then  //め计ゑ耕虫 : κだゑ
      begin
        //璸衡め计ゑㄒ(繵笵ぇ羆Μ跌め计/羆ΤΜ跌め)*100
        fL_CountsRate := (nL_ChannelTotalViewCounts/nG_TotalValidViewCounts)*100;

        //め计ゑㄒゑ耕
        if fL_CountsRate <= fL_Subscriber_Count_Percent1 then
          fL_ComputerPara := fL_Amount_Percent1
        else if fL_CountsRate <= fL_Subscriber_Count_Percent2 then
          fL_ComputerPara := fL_Amount_Percent2
        else if fL_CountsRate <= fL_Subscriber_Count_Percent3 then
          fL_ComputerPara := fL_Amount_Percent3
        else if fL_CountsRate <= fL_Subscriber_Count_Percent4 then
          fL_ComputerPara := fL_Amount_Percent4
        else if fL_CountsRate <= fL_Subscriber_Count_Percent5 then
          fL_ComputerPara := fL_Amount_Percent5;

        if sI_PackageID <> '' then   //矪瞶甅繺╊眀
        begin
          //赣掸繵笵﹍肂
          cdsPackageInfo.First;
          while not cdsPackageInfo.Eof do
          begin
            if cdsPackageInfo.Locate('PackageID',sI_PackageID,[]) then
            begin
              //赣甅繺羆﹍肂
              fL_TotalOrigionalPrice := cdsPackageInfo.FieldByName('TotalOrigionalPrice').AsFloat;
              //赣甅繺羆龟Μ肂
              fL_TotalRealCharge := cdsPackageInfo.FieldByName('TotalRealCharge').AsFloat;
            end;
            cdsPackageInfo.Next;
          end;


          //虫κだゑ璸衡 =  ( 赣掸繵笵﹍肂 / 赣甅繺羆﹍肂) * 赣甅繺羆龟Μ肂 * 璸衡把计
          Result := RoundTo((fI_OrigionalPrice / fL_TotalOrigionalPrice) * fL_TotalRealCharge * fL_ComputerPara / 100,0);
        end
        else   //矪瞶ぃ琌甅繺╊眀
        begin
          //虫肂璸衡 =  赣掸繵笵﹍肂 * 璸衡把计
          Result := RoundTo((fI_OrigionalPrice * fL_ComputerPara / 100),0);
        end;
      end;
    end;
end;


procedure TdtmMain2.handleShareAmt1(sI_CompCode,sI_CustID,sI_StartDate,sI_EndDate, sI_BelongYM,
  sI_CurDate,sI_ShowDetail,sI_PrvBelongYM : String;var sI_ChannelString : String);
var
    sL_ProductID,sL_FormulaID,sL_ProviderID,sL_PackageID : String;
    sL_CompType,sL_SQL,sL_UseBaseFormula : String;
    fL_OrigionalPrice,fL_RealCharge,fL_ProviderPercent,fL_SoPercent : Double;
    fL_Provider_Benefit,fL_So_Benefit,fL_Emc_Benefit : Double;
    fL_Total_EMC_Income,fL_Total_Emc_Benefit : Double;
    fL_Total_So_Benefit,fL_Total_Provider_Benefit : Double;
    fL_OldProviderPercent,fL_OldSoPercent : Double;
    fL_Income,fL_Outcome,fL_OIncome,fL_OOutcome : Double;
    fL_ORealCharge,fL_OEmc_Benefit,fL_OSo_Benefit,fL_OProvider_Benefit : Double;
    nL_SingleCounts,nL_PackageCounts,nL_OSingleCounts,nL_OPackageCounts : Integer;
    sL_ChannelName,sL_SingleString,sL_PackageString,sL_ProviderDesc : String;
    bL_IsIncome : Boolean;
    nL_FirstCounts,nL_LastCount,nL_OLastCount : Integer;
    nL_OldAddCount,nL_OldReduceCount,nL_AddCount,nL_ReduceCount : Integer;
    sL_FirstFlag : String;
begin
    //璸衡膀セそΑ┪そΑ
    nL_AddCount := 0;
    nL_ReduceCount := 0;
    cdsProductInfo.First;
    while not cdsProductInfo.Eof do
    begin

      sL_ProductID := cdsProductInfo.FieldByName('PRODUCTID').AsString;
      sL_FormulaID := cdsProductInfo.FieldByName('FORMULAID').AsString;
      sL_ProviderID := cdsProductInfo.FieldByName('ProviderID').AsString;
      sL_ProviderDesc := getProviderName(sL_ProviderID);

      sL_PackageID := cdsProductInfo.FieldByName('PackageID').AsString;
      fL_OrigionalPrice := cdsProductInfo.FieldByName('OrigionalPrice').AsFloat;
      fL_RealCharge := cdsProductInfo.FieldByName('RealCharge').AsFloat;

      sL_UseBaseFormula := cdsProductInfo.FieldByName('UseBaseFormula').AsString;
      fL_ProviderPercent := cdsProductInfo.FieldByName('Provider_Percent').AsFloat;
      fL_SoPercent := cdsProductInfo.FieldByName('So_Percent').AsFloat;

      bL_IsIncome := cdsProductInfo.FieldByName('IsInCome').AsBoolean;
      sL_FirstFlag :=  cdsProductInfo.FieldByName('FirstFlag').AsString;

      if UpperCase(sL_UseBaseFormula) = 'Y' then   //璸衡膀セそΑ
      begin
        //╊眀穨ぇ繵笵菌╊眀璸衡把计 ( ヘ玡度続ノ膀セそΑ )
        saveToSO119(sI_CompCode,sI_BelongYM,sL_ProductID,sL_ProviderID,sI_CurDate,fL_ProviderPercent,fL_SoPercent);

        if bL_IsIncome then                    //Μ禣兜ヘ
        begin
          fL_Provider_Benefit :=  fL_RealCharge * fL_ProviderPercent /100;
          fL_So_Benefit := fL_RealCharge * fL_SoPercent /100;


          fL_Income := fL_RealCharge;
          fL_Outcome := 0;

          //狦SO033,SO034  FirstFlag Τ戳,赣る穝糤ぇめ
          if sL_FirstFlag <> '' then
          begin
            nL_LastCount := 1;
            nL_AddCount := 1;
          end
          else
          begin
            nL_LastCount := 0;
          end;
        end
        else   //癶禣兜ヘ
        begin
          //眔るㄑ莱坝のSO╊眀ゑㄒ
          getLastMonthComputePara(sI_CompCode,sI_PrvBelongYM,sL_ProductID,fL_OldProviderPercent,fL_OldSoPercent);

          if (fL_OldProviderPercent=-1) and (fL_OldSoPercent=-1) then
          begin
            //璝琩礚る戈,玥盿セる╊眀ゑㄒ璸衡
            fL_OldProviderPercent := fL_ProviderPercent;
            fL_OldSoPercent := fL_SoPercent;
          end;

          fL_Provider_Benefit := (fL_RealCharge * fL_OldProviderPercent /100) * -1;
          fL_So_Benefit := (fL_RealCharge * fL_OldSoPercent /100) * -1;

          //癶禣
          fL_RealCharge := fL_RealCharge * -1;

          fL_Income := 0;
          fL_Outcome := fL_RealCharge;
          nL_LastCount := -1;
          nL_ReduceCount := 1;
        end;
      end
      else                             //璸衡そΑ
      begin
        //璸衡籔繵笵坝ぇ╊眀肂
        sL_CompType := '2';   //把酚SO111 COMP_TYPE逆
        //sI_CompOrProviderID := 繵笵坝絏
        fL_Provider_Benefit := calculateSoOrProviderAmt(sL_ProductID,sL_PackageID,sL_FormulaID,sL_CompType,sL_ProviderID,fL_OrigionalPrice);

        //璸衡籔 So ぇ╊眀肂
        sL_CompType := '1';   //把酚SO111 COMP_TYPE逆
        //sI_CompOrProviderID := Soそ絏
        fL_So_Benefit := calculateSoOrProviderAmt(sL_ProductID,sL_PackageID,sL_FormulaID,sL_CompType,sI_CompCode,fL_OrigionalPrice);

      end;

      //璸衡籔 Emc ぇ╊眀肂 = 龟Μ肂 - So ぇ╊眀肂 - 繵笵坝ぇ╊眀肂
      fL_Emc_Benefit := fL_RealCharge - fL_So_Benefit - fL_Provider_Benefit;


      //璸衡赣め┮璹繵笵
      sL_ChannelName := getChannelName(sL_ProductID);
      if sL_PackageID = '' then
      begin
        if sL_SingleString = '' then
          sL_SingleString := '虫翴: ' + sL_ChannelName
        else
          sL_SingleString := sL_SingleString + ',' + sL_ChannelName;
      end
      else
      begin
        if sL_PackageString = '' then
          sL_PackageString := '甅繺: ' + sL_ChannelName
        else
          sL_PackageString := sL_PackageString + ',' + sL_ChannelName;
      end;


      if sL_PackageID = '' then
      begin
        sL_PackageID := 'null';
        nL_SingleCounts := 1;
        nL_PackageCounts := 0;
      end
      else
      begin
        nL_SingleCounts := 0;
        nL_PackageCounts := 1;
      end;

      //眔戳计秖
      nL_FirstCounts := getFirstCounts(sI_CompCode, sI_PrvBelongYM,sL_ProductID);

      if cdsSO114.Locate('COMP_ID;BELONGYM;PRODUCTID', VarArrayOf([sI_CompCode, sI_BelongYM,sL_ProductID]), []) then
      begin
        //繵笵盢–掸戈既 cdsSO114 仓
        with cdsSO114 do
        begin
          //ㄓ肂
          fL_ORealCharge := FieldByName('EMC_INCOME').AsFloat;
          fL_OEmc_Benefit := FieldByName('EMC_BENEFIT').AsFloat;
          fL_OSo_Benefit := FieldByName('SO_BENEFIT').AsFloat;
          fL_OProvider_Benefit := FieldByName('PROVIDER_BENEFIT').AsFloat;
          nL_OSingleCounts := FieldByName('SINGLECOUNTS').AsInteger;
          nL_OPackageCounts := FieldByName('PACKAGECOUNTS').AsInteger;

          fL_OIncome := FieldByName('INCOME').AsFloat;
          fL_OOutcome := (FieldByName('OUTCOME').AsFloat * -1);

          nL_OLastCount := FieldByName('LASTCOUNT').AsInteger;
          nL_OldAddCount := FieldByName('AddCount').AsInteger;
          nL_OldReduceCount := FieldByName('ReduceCount').AsInteger;


          Delete;
          Append;
          //仓
          FieldByName('COMP_ID').AsString := sI_CompCode;
          FieldByName('BELONGYM').AsString := sI_BelongYM;
          FieldByName('BEGIN_DATE').AsDateTime := StrToDateTime(sI_StartDate);
          FieldByName('END_DATE').AsDateTime := StrToDateTime(sI_EndDate);
          FieldByName('EMC_INCOME').AsFloat := RoundTo(fL_ORealCharge,0) + RoundTo(fL_RealCharge,0);

          FieldByName('EMC_BENEFIT').AsFloat := RoundTo(fL_OEmc_Benefit,0) + RoundTo(fL_Emc_Benefit,0);
          FieldByName('SO_BENEFIT').AsFloat := RoundTo(fL_OSo_Benefit,0) +  RoundTo(fL_So_Benefit,0);
          FieldByName('PRODUCTID').AsString := sL_ProductID;

          FieldByName('PROVIDERID').AsString := sL_ProviderID;
          FieldByName('PROVIDERDESC').AsString := sL_ProviderDesc;

          FieldByName('PROVIDER_BENEFIT').AsFloat := RoundTo(fL_OProvider_Benefit,0) + RoundTo(fL_Provider_Benefit,0);
          FieldByName('SINGLECOUNTS').AsInteger := nL_OSingleCounts + nL_SingleCounts;
          FieldByName('PACKAGECOUNTS').AsInteger := nL_OPackageCounts + nL_PackageCounts;

          //璸衡そΑノㄤよΑ衡
          if UpperCase(sL_UseBaseFormula) = 'Y' then   //璸衡膀セそΑ
          begin
            FieldByName('INCOME').AsFloat := RoundTo(fL_Income,0) + RoundTo(fL_OIncome,0);
            FieldByName('OUTCOME').AsFloat := ((RoundTo(fL_Outcome,0) + RoundTo(fL_OOutcome,0)) * -1);
            FieldByName('FIRSTCOUNT').AsInteger := nL_FirstCounts;
            FieldByName('LASTCOUNT').AsInteger := nL_OLastCount + nL_LastCount;
            FieldByName('ADDCOUNT').AsInteger := nL_OldAddCount + nL_AddCount;
            FieldByName('REDUCECOUNT').AsInteger := nL_OldReduceCount + nL_ReduceCount;
          end;

          FieldByName('OPERATOR').AsString := frmMainMenu.sG_User;
          FieldByName('UPDTIME').AsString := sI_CurDate;

          Post;
        end;
      end
      else
      begin
        //繵笵盢–掸戈既 cdsSO114
        with cdsSO114 do
        begin
          Append;

          FieldByName('COMP_ID').AsString := sI_CompCode;
          FieldByName('BELONGYM').AsString := sI_BelongYM;
          FieldByName('BEGIN_DATE').AsDateTime := StrToDateTime(sI_StartDate);
          FieldByName('END_DATE').AsDateTime := StrToDateTime(sI_EndDate);
          FieldByName('EMC_INCOME').AsFloat := RoundTo(fL_RealCharge,0);

          FieldByName('EMC_BENEFIT').AsFloat := RoundTo(fL_Emc_Benefit,0);
          FieldByName('SO_BENEFIT').AsFloat := RoundTo(fL_So_Benefit,0);
          FieldByName('PRODUCTID').AsString := sL_ProductID;

          FieldByName('PROVIDERID').AsString := sL_ProviderID;
          FieldByName('PROVIDERDESC').AsString := sL_ProviderDesc;

          FieldByName('PROVIDER_BENEFIT').AsFloat := RoundTo(fL_Provider_Benefit,0);
          FieldByName('SINGLECOUNTS').AsInteger := nL_SingleCounts;
          FieldByName('PACKAGECOUNTS').AsInteger := nL_PackageCounts;

          //璸衡そΑノㄤよΑ衡
          if UpperCase(sL_UseBaseFormula) = 'Y' then   //璸衡膀セそΑ
          begin
            FieldByName('INCOME').AsFloat := RoundTo(fL_Income,0);
            FieldByName('OUTCOME').AsFloat := (RoundTo(fL_Outcome,0) * -1);
            FieldByName('FIRSTCOUNT').AsInteger := nL_FirstCounts;
            FieldByName('LASTCOUNT').AsInteger := nL_LastCount;
            FieldByName('ADDCOUNT').AsInteger := nL_AddCount;
            FieldByName('REDUCECOUNT').AsInteger := nL_ReduceCount;
          end;

          FieldByName('OPERATOR').AsString := frmMainMenu.sG_User;
          FieldByName('UPDTIME').AsString := sI_CurDate;

          Post;
        end;
      end;

      //礶Τつ匡陪ボめΜ跌灿戈(ゲ斗翴匡兜锣タΑ灿戈),璸衡SO115
      if sI_ShowDetail = 'Y' then
      begin
        //盢–掸戈既 cdsSO115
        with cdsSO115 do
        begin
          Append;

          FieldByName('COMP_ID').AsString := sI_CompCode;
          FieldByName('BELONGYM').AsString := sI_BelongYM;
          FieldByName('CUSTID').AsString := sI_CustID;
          FieldByName('PRODUCT_ID').AsString := sL_ProductID;

          if sL_PackageID <> 'null' then
            FieldByName('PACKAGE_ID').AsString := sL_PackageID;

          FieldByName('PROVIDER_ID').AsString := sL_ProviderID;
          FieldByName('ORIGIONALPRICE').AsFloat := RoundTo(fL_OrigionalPrice,0);
          FieldByName('REALCHARGE').AsFloat := RoundTo(fL_RealCharge,0);;
          FieldByName('EMC_BENEFIT').AsFloat := RoundTo(fL_Emc_Benefit,0);
          FieldByName('SO_BENEFIT').AsFloat := RoundTo(fL_So_Benefit,0);
          FieldByName('PROVIDER_BENEFIT').AsFloat := RoundTo(fL_Provider_Benefit,0);
          FieldByName('OPERATOR').AsString := frmMainMenu.sG_User;
          FieldByName('UPDTIME').AsString := sI_CurDate;

          Post;
        end;
      end;
      cdsProductInfo.Next;
    end;

    //矪瞶﹃
    if sL_SingleString = '' then
      sI_ChannelString := sL_PackageString
    else if sL_PackageString = '' then
      sI_ChannelString := sL_SingleString
    else
      sI_ChannelString := sL_SingleString + ' ;        ' +  sL_PackageString;
end;

procedure TdtmMain2.createCDSActive;
begin
    if not cdsSO114.Active = true then
      cdsSO114.Active := true;

    if not cdsSO114B.Active = true then
      cdsSO114B.Active := true;

    if not cdsSO114C.Active = true then
      cdsSO114C.Active := true;

    if not cdsSO115.Active = true then
      cdsSO115.Active := true;

    if not cdsSO116.Active = true then
      cdsSO116.Active := true;

    if not cdsSO117.Active = true then
      cdsSO117.Active := true;

    if not cdsSO115B.Active = true then
      cdsSO115B.Active := true;

    if not cdsSO119.Active = true then
      cdsSO119.Active := true;

    if not cdsCD019.Active = true then
      cdsCD019.Active := true;

    if not cdsTempSO119.Active = true then
      cdsTempSO119.Active := true;


end;

procedure TdtmMain2.getRptSQL(I_CDS: TClientDataset; sI_MID: String;
  var I_SQLStrList: TStringList);
var
    nL_FieldCount, ii : Integer;
    sL_SQLFieldPos, sL_SQL, sL_TmpSQL, sL_PrgType, sL_FieldName : String;
    sL_Header, sL_PrgName : String;
begin
    with I_CDS do
    begin
      sL_SQL := 'select * from AA909 where FORMATNAME=' + STR_SEP + sI_MID + STR_SEP + ' order by PRGTYPE';
      Close;
      CommandText := sL_SQL;
      Open;
      if Recordcount>0 then
      begin
        while not Eof do
        begin
          sL_PrgType := FieldByName('PRGTYPE').AsString;
          sL_PrgName := FieldByName('PRGNAME').AsString;

          nL_FieldCount := Fields.Count;
          for ii:=0 to nL_FieldCount-1 do
          begin
            sL_FieldName := Fields[ii].FieldName;
            if (FieldByName(sL_FieldName).AsString<>'') and (AnsiPos('SQL',sL_FieldName)<>0) then
            begin
              sL_TmpSQL := FieldByName(sL_FieldName).AsString;
              sL_SQLFieldPos := Copy(sL_FieldName,4,length(sL_FieldName)-3);
              if sL_PrgName<>'' then
                sL_Header := '[' + sL_PrgType+'-' + sL_SQLFieldPos + '-' + sL_PrgName + ']'
              else
                sL_Header := '[' + sL_PrgType+'-' + sL_SQLFieldPos  + ']';
              I_SQLStrList.Add(sL_Header) ;
              I_SQLStrList.Add(sL_TmpSQL);

            end;
          end;

          Next;
        end;

      end;
      Close;
    end;

end;

procedure TdtmMain2.saveToDB(I_CDS: TClientDataSet);
begin
    with I_CDS do
    begin
      if state in [dsEdit, dsInsert] then
        Post;
      if ChangeCount >0 then
        ApplyUpdates(-1);

    end;
end;

procedure TdtmMain2.deleteReportData(sI_compcode,sI_BelongYM : String);
var
    sL_SQL : String;
begin
    sL_SQL := 'DELETE SO114 WHERE COMP_ID=''' + sI_compcode +
              ''' AND BelongYM=''' + sI_BelongYM + '''';

//savetofile(sL_SQL, 'C:\EMC\╊眀\project\client\Error.txt');
    with cdsCom do
    begin
      Close;
      CommandText := sL_SQL;
      Execute;
    end;

    sL_SQL := 'DELETE SO115 WHERE COMP_ID=''' + sI_compcode +
              ''' AND BelongYM=''' + sI_BelongYM + '''';

    with cdsCom do
    begin
      Close;
      CommandText := sL_SQL;
      Execute;
    end;

    sL_SQL := 'DELETE SO116 WHERE COMP_ID=''' + sI_compcode +
              ''' AND BelongYM=''' + sI_BelongYM + '''';

    with cdsCom do
    begin
      Close;
      CommandText := sL_SQL;
      Execute;
    end;

    sL_SQL := 'DELETE SO117 WHERE COMP_ID=''' + sI_compcode +
              ''' AND BelongYM=''' + sI_BelongYM + '''';

    with cdsCom do
    begin
      Close;
      CommandText := sL_SQL;
      Execute;
    end;


    sL_SQL := 'DELETE SO118 WHERE COMP_ID=''' + sI_compcode +
              ''' AND LOCKYM=''' + sI_BelongYM + '''';

    with cdsCom do
    begin
      Close;
      CommandText := sL_SQL;
      Execute;
    end;


    sL_SQL := 'DELETE SO119 WHERE COMP_ID=''' + sI_compcode +
              ''' AND BelongYM=''' + sI_BelongYM + '''';

    with cdsCom do
    begin
      Close;
      CommandText := sL_SQL;
      Execute;
    end;
end;

procedure TdtmMain2.getReportData(sI_compcode, sI_BelongYM: String);
var
    sL_SQL : String;
begin
    sL_SQL := 'SELECT * FROM SO114 WHERE COMP_ID=''' + sI_compcode +
              ''' AND BelongYM=''' + sI_BelongYM +
              ''' ORDER BY PROVIDERID,PRODUCTID';


    with cdsSO114 do
    begin
      Close;
      CommandText := sL_SQL;
      Open;
    end;


    sL_SQL := 'SELECT * FROM SO116 WHERE COMP_ID=''' + sI_compcode +
              ''' AND BelongYM=''' + sI_BelongYM + '''';

    with cdsSO116 do
    begin
      Close;
      CommandText := sL_SQL;
      Open;
    end;


    sL_SQL := 'SELECT * FROM SO117 WHERE COMP_ID=''' + sI_compcode +
              ''' AND BelongYM=''' + sI_BelongYM + '''';

    with cdsSO117 do
    begin
      Close;
      CommandText := sL_SQL;
      Open;
    end;


    //if sI_ShowDetail = 'Y' then   //瞷эぃノ翴匡
    getSO114Data(sI_compcode, sI_BelongYM);

end;

function TdtmMain2.getChannelName(sI_ProductID: String): String;
var
    sL_SQL : String;
begin

    sL_SQL := 'SELECT * FROM SO112 WHERE PRODUCT_ID=''' + sI_ProductID + '''';

    with cdsCom do
    begin
      Close;
      CommandText := sL_SQL;
      Open;

      Result := FieldByName('PRODUCT_NAME').AsString;
    end;
end;

procedure TdtmMain2.saveTocdsSo116(sI_CompCode, sI_BelongYM, sI_CustID,
  sI_CurDate, sI_ChannelString: String);
begin
    //盢戈╊眀穨ぇめΜ跌繵笵灿郎
    with cdsSO116 do
    begin
      Append;
      FieldByName('COMP_ID').AsString := sI_CompCode;
      FieldByName('BELONGYM').AsString := sI_BelongYM;
      FieldByName('CUSTID').AsString := sI_CustID;
      FieldByName('CHANNEL_STRING').AsString := sI_ChannelString;
      FieldByName('OPERATOR').AsString := frmMainMenu.sG_User;
      FieldByName('UPDTIME').AsString := sI_CurDate;
      Post;
    end;
end;

procedure TdtmMain2.saveTocdsSo117(sI_CompCode, sI_BelongYM, sI_CurDate,
  sI_ChannelViewCountsString: String; nI_CustCounts: Integer);
begin
    //盢戈╊眀穨ぇ羆㎝戈
    with cdsSO117 do
    begin
      Append;
      FieldByName('COMP_ID').AsString := sI_CompCode;
      FieldByName('BELONGYM').AsString := sI_BelongYM;
      FieldByName('TOTALVALIDCOUNTS').AsInteger := nG_TotalValidViewCounts;
      FieldByName('TOTALCUSTCOUNTS').AsInteger := nI_CustCounts;
      FieldByName('CHANNELCOUNTS').AsString := sI_ChannelViewCountsString;
      FieldByName('OPERATOR').AsString := frmMainMenu.sG_User;
      FieldByName('UPDTIME').AsString := sI_CurDate;
      Post;
    end;
end;

function TdtmMain2.getCustName(sI_CustID: String): String;
var
    sL_SQL : String;
begin
    sL_SQL := 'SELECT CUSTNAME FROM SO001 WHERE COMPCODE=' + frmMainMenu.sG_CompCode +
              ' AND CUSTID=' + sI_CustID;
    with cdsCom do
    begin
      Close;
      CommandText := sL_SQL;
      Open;

      Result := FieldByName('CUSTNAME').AsString;
    end;

end;

procedure TdtmMain2.getSO114Data(sI_compcode, sI_BelongYM: String);
var
    sL_SQL : String;
begin
    //眔才兵ンㄑ莱坝┮Τ戈
    sL_SQL := 'SELECT * FROM SO114 WHERE COMP_ID=''' + sI_compcode +
              ''' AND BelongYM=''' + sI_BelongYM +
              ''' ORDER BY PROVIDERID,PRODUCTID';


    with cdsSO114B do
    begin
      Close;
      CommandText := sL_SQL;
      Open;
    end;

end;

function TdtmMain2.getProviderName(sI_ProviderID: String): String;
var
    sL_SQL : String;
begin
    sL_SQL := 'SELECT PROVIDER_NAME FROM SO113 WHERE PROVIDER_ID=''' + sI_ProviderID + '''';

    with cdsCom do
    begin
      Close;
      CommandText := sL_SQL;
      Open;

      Result := FieldByName('PROVIDER_NAME').AsString
    end;
end;

function TdtmMain2.getProviderBenefit(sI_compcode, sI_BelongYM, sI_ProductID,
  sI_ProviderID: String): Double;
var
    sL_SQL : String;
begin
    sL_SQL := 'SELECT SUM(NVL(PROVIDER_BENEFIT,0)) TOTAL_PROVIDER_BENEFIT FROM SO114 WHERE COMP_ID=''' + sI_compcode +
              ''' AND BelongYM=''' + sI_BelongYM +
              ''' AND PRODUCTID=''' + sI_ProductID + ''' AND PROVIDERID=''' +
              sI_ProviderID + '''';

    with cdsCom do
    begin
      Close;
      CommandText := sL_SQL;
      Open;
      Result := FieldByName('TOTAL_PROVIDER_BENEFIT').AsFloat
    end;
end;

function TdtmMain2.getProviderCounts(sI_compcode, sI_BelongYM, sI_ProductID,
  sI_ProviderID: String): Integer;
var
    sL_SQL : String;
begin
    sL_SQL := 'SELECT SUM(NVL(SINGLECOUNTS,0)) + SUM(NVL(PACKAGECOUNTS,0)) CHANNELCOUNTS' +
              ' FROM SO114 WHERE COMP_ID=''' + sI_compcode + ''' AND BelongYM=''' + sI_BelongYM +
              ''' AND PRODUCTID=''' + sI_ProductID + ''' AND PROVIDERID=''' +
              sI_ProviderID + '''';

    with cdsCom do
    begin
      Close;
      CommandText := sL_SQL;
      Open;

      Result := FieldByName('CHANNELCOUNTS').AsInteger
    end;
end;






procedure TdtmMain2.initialData;
var
    ii,jj : Integer;
begin
    if not FileExists('COMPUTE_ISSUE.cds') then
    begin
        cdsCompute_issue.CreateDataSet;
        cdsCompute_issue.SaveToFile('COMPUTE_ISSUE.cds');
    end;
    cdsCompute_issue.LoadFromFile('COMPUTE_ISSUE.cds');

    with cdsCompute_issue do
    begin
        Append;
        FieldByName('CODE').AsString:='1';
        FieldByName('COMPUTE_ISSUE').AsString:='仓秈禯';
        post;
        Append;
        FieldByName('CODE').AsString:='2';
        FieldByName('COMPUTE_ISSUE').AsString:='辅翴禯';
        post;
    end;


    //そ/ㄑ莱坝 APPEND  cdsSo113Cd039
    cdsSo113.Close;
    cdsSo113.CommandText:=('SELECT PROVIDER_ID,PROVIDER_NAME FROM SO113');
    cdsSo113.open;
    cdsCD039.Close;
    cdsCD039.CommandText:=('SELECT CODENO,DESCRIPTION FROM CD039');
    cdsCD039.open;

    cdsCD039.First;
    cdsSo113Cd039.CreateDataSet;
    While not cdsCD039.Eof do
    begin
      try
        cdsSo113Cd039.Append;
        cdsSo113Cd039.FieldByName('CLASSIFY_ID').AsString:='1';
        cdsSo113Cd039.FieldByName('CODE').AsString:=cdsCD039.FieldByName('CODENO').AsString;
        cdsSo113Cd039.FieldByName('DESC').AsString:=cdsCD039.FieldByName('DESCRIPTION').AsString;
        cdsSo113Cd039.Post;
        cdsCD039.Next;
      except
      end;
    end;
    cdsSo113.First;
    While not cdsSo113.Eof do
    begin
      try
        cdsSo113Cd039.Append;
        cdsSo113Cd039.FieldByName('CLASSIFY_ID').AsString:='2';
        cdsSo113Cd039.FieldByName('CODE').AsString:=cdsSo113.FieldByName('PROVIDER_ID').AsString;
        cdsSo113Cd039.FieldByName('DESC').AsString:=cdsSo113.FieldByName('PROVIDER_NAME').AsString;
        cdsSo113Cd039.Post;
        cdsSo113.Next;
      except
      end;
    end;

    if not FileExists('RANGE_UNIT.cds') then
    begin
        cdsRANGE_UNIT.CreateDataSet;
        cdsRANGE_UNIT.SaveToFile('RANGE_UNIT.cds');
    end;
    cdsRANGE_UNIT.LoadFromFile('RANGE_UNIT.cds');
    with cdsRANGE_UNIT do
    begin
        Append;
        FieldByName('RANGE_UNIT').AsString:='1';
        FieldByName('RANGE_UNIT_NAME').AsString:='め计';
        post;
        Append;
        FieldByName('RANGE_UNIT').AsString:='2';
        FieldByName('RANGE_UNIT_NAME').AsString:='κだゑ';
        post;
    end;
//穝糤   NANSEN-910920
    initCdsAttribute;


    for jj:=1 to 5 do
    begin
      with cdsRefNo do
      begin
        Append;
        FieldByName('CODENO').AsString:=IntToStr(jj);
        case jj of
          1:begin
            FieldByName('DESCRIPTION').AsString:='ㄢ顶甅繺( ゑㄒ )';
            end;
          2:begin
            FieldByName('DESCRIPTION').AsString:='め计计秖禯( 虫基 )';
            end;
          3:begin
            FieldByName('DESCRIPTION').AsString:='め计计秖禯( ゑㄒ )';
            end;
          4:begin
            FieldByName('DESCRIPTION').AsString:='め计κだゑ禯( 虫基 )';
            end;
          5:begin
            FieldByName('DESCRIPTION').AsString:='め计κだゑ禯( ゑㄒ )';
            end;
        end;
        Post;
      end;
    end;

end;




function TdtmMain2.getChannelCustData(sI_CompCode,sI_Table,sI_CitemCode,sI_StartDate,sI_EndDate : String) : Integer;
var
    sL_SQL : String;
begin
    sL_SQL := 'SELECT * FROM ' + sI_Table + ' WHERE CITEMCODE =' + sI_CitemCode  +
              ' AND (REALDATE BETWEEN ' + TUstr.getOracleSQLDateStr(StrToDate(sI_StartDate)) +
              ' AND ' + TUstr.getOracleSQLDateStr(StrToDate(sI_EndDate)) + ')' +
              ' AND COMPCODE=' + sI_CompCode;

    with cdsCustData do
    begin
      Close;
      CommandText := sL_SQL;
      Open;

      Result := RecordCount;
    end;
end;


function TdtmMain2.handleChannelCustData(sI_CompCode,sI_BelongYM,sI_CurDate,sI_ChannelName : String;var nI_TotalSingleCounts,nI_TotalPackageCounts : Integer): Double;
var
    sL_PackageID,sL_CustID,sL_SingleString,sL_PackageString : String;
    fL_RealAmt,fL_TotalChannelRealAmt : Double;
    sL_ChannelString,sL_Temp : String;
    L_StrList : TStringList;
begin
    //盢赣繵笵–め┮繵笵SO116い,璸衡繵笵羆龟Μ肂の繵笵虫翴の甅繺计秖
    L_StrList := TStringList.Create;
    fL_TotalChannelRealAmt := 0;
    nI_TotalSingleCounts := 0;
    nI_TotalPackageCounts := 0;

    //璸衡繵笵┮Τめ
    cdsCustData.First;
    while not cdsCustData.Eof do
    begin
      fL_RealAmt := 0 ;
      sL_CustID := cdsCustData.FieldByName('CUSTID').AsString;
      sL_PackageID := cdsCustData.FieldByName('PACKAGENO').AsString;

      //璸衡繵笵羆龟Μ肂
      fL_RealAmt := cdsCustData.FieldByName('REALAMT').AsFloat;
      fL_TotalChannelRealAmt := fL_TotalChannelRealAmt + fL_RealAmt;


      //璸衡繵笵虫翴の甅繺计秖
      if sL_PackageID = '' then
        nI_TotalSingleCounts := nI_TotalSingleCounts + 1
      else
        nI_TotalPackageCounts := nI_TotalPackageCounts + 1;


      sL_ChannelString := '';
      if cdsSO116.Locate('COMP_ID;BELONGYM;CUSTID', VarArrayOf([sI_CompCode, sI_BelongYM,sL_CustID]), []) then
      begin
        sL_ChannelString := cdsSO116.FieldByName('CHANNEL_STRING').AsString;
        L_StrList := TUstr.ParseStrings(sL_ChannelString,';');

        if L_StrList.Count = 1 then
        begin
          sL_Temp := copy(sL_ChannelString,1,5);
          if sL_PackageID = '' then
          begin
            if sL_Temp = '虫翴:' then
              sL_ChannelString := sL_ChannelString + ',' + sI_ChannelName
            else
              sL_ChannelString := sL_ChannelString + '  ;   甅繺:' + sI_ChannelName;
          end
          else
          begin
            if sL_Temp <> '虫翴:' then
              sL_ChannelString := sL_ChannelString + ',' + sI_ChannelName
            else
              sL_ChannelString := '虫翴:' + sI_ChannelName +  '  ,   ' + sL_ChannelString;
          end;

        end
        else
        begin
          sL_SingleString := L_StrList.Strings[0];
          sL_PackageString := L_StrList.Strings[1];

          if sL_PackageID = '' then
            sL_ChannelString := sL_SingleString + ',' + sI_ChannelName
          else
            sL_ChannelString := sL_PackageString + ',' + sI_ChannelName;

        end;

        with cdsSO116 do
        begin
          Delete;
          Append;

          FieldByName('COMP_ID').AsString := sI_CompCode;
          FieldByName('BELONGYM').AsString := sI_BelongYM;

          FieldByName('CUSTID').AsString := sL_CustID;
          FieldByName('CHANNEL_STRING').AsString := sL_ChannelString;

          FieldByName('OPERATOR').AsString := frmMainMenu.sG_User;
          FieldByName('UPDTIME').AsString := sI_CurDate;

          Post;
        end;
      end
      else
      begin
        //繵笵盢–掸戈既 cdsSO116
        with cdsSO116 do
        begin
          if sL_PackageID = '' then
            sL_ChannelString := '虫翴: ' + sI_ChannelName
          else
            sL_ChannelString := '甅繺: ' + sI_ChannelName;

          Append;

          FieldByName('COMP_ID').AsString := sI_CompCode;
          FieldByName('BELONGYM').AsString := sI_BelongYM;

          FieldByName('CUSTID').AsString := sL_CustID;
          FieldByName('CHANNEL_STRING').AsString := sL_ChannelString;

          FieldByName('OPERATOR').AsString := frmMainMenu.sG_User;
          FieldByName('UPDTIME').AsString := sI_CurDate;

          Post;
        end;
      end;

      cdsCustData.Next;
    end;

    //肚繵笵龟Μ羆肂
    Result := fL_TotalChannelRealAmt;
end;

procedure TdtmMain2.getSharePara(sI_CompCode,sI_ProductID,sI_ProviderID : String);
var
    sL_SQL,sL_SQLTitle,sL_Where1,sL_Where2 : String;
begin
    sL_SQLTitle := 'SELECT * FROM SO111';

    //тㄑ莱坝╊眀璸衡计沮 ( COMP_TYPE=2 , SO_PROVIDER_ID = ㄑ莱坝絏 )
    sL_Where1 := ' WHERE PRODUCT_ID=''' + sI_ProductID +
                 ''' AND COMP_TYPE=''2'' AND SO_PROVIDER_ID=''' +
                 sI_ProviderID + '''';

    sL_SQL := sL_SQLTitle + sL_Where1;
    with cdsTallyRefData1 do
    begin
      Close;
      CommandText := sL_SQL;
      Open;
    end;

    //т SO ╊眀璸衡计沮 ( COMP_TYPE=1 , SO_PROVIDER_ID = そ絏 )
    sL_Where2 := ' WHERE PRODUCT_ID=''' + sI_ProductID +
                 ''' AND COMP_TYPE=''1'' AND SO_PROVIDER_ID=''' +
                 sI_CompCode + '''';

    sL_SQL := sL_SQLTitle + sL_Where2;
    with cdsTallyRefData2 do
    begin
      Close;
      CommandText := sL_SQL;
      Open;
    end;
end;

function TdtmMain2.handleShareAmt2(sI_FormulaID : String;nI_ChannelTotalViewCounts : Integer;fI_OrigionalPrice : Double;var fI_Provider_Benefit,fI_So_Benefit : Double) : String;
var
    sL_SoOrProvider,sL_Msg : String;
begin
    //矪瞶そΑ┪┪┪き繵笵╊眀戈
    sL_SoOrProvider := '1'; //矪瞶繵笵坝╊眀肂
    sL_Msg := calculateSoOrProviderAmt2(sL_SoOrProvider,sI_FormulaID,nI_ChannelTotalViewCounts,fI_OrigionalPrice,fI_Provider_Benefit);
    if sL_Msg <> '' then
    begin
      Result := sL_Msg;
      exit;
    end;

    sL_SoOrProvider := '2'; //矪瞶 SO ╊眀肂
    sL_Msg := calculateSoOrProviderAmt2(sL_SoOrProvider,sI_FormulaID,nI_ChannelTotalViewCounts,fI_OrigionalPrice,fI_So_Benefit);
    if sL_Msg <> '' then
    begin
      Result := sL_Msg;
      exit;
    end;

end;

function TdtmMain2.getFormulaRefNo(sI_FormulaID: String): String;
var
    sL_SQL : String;
begin
    sL_SQL := 'SELECT REF_NO FROM SO110 WHERE FORMULA_ID=''' + sI_FormulaID +
              ''' AND STOPFLAG=0';

    with cdsCom do
    begin
      Close;
      CommandText := sL_SQL;
      Open;

      Result := FieldByName('REF_NO').AsString
    end;
end;


function TdtmMain2.calculateSoOrProviderAmt2(sI_SoOrProvider,
  sI_FormulaID: String; nI_ChannelTotalViewCounts: Integer;fI_OrigionalPrice : Double;var fI_SoOrProviderBenefit : Double) : String;
var
    sL_RangeUnit,sL_ComputeIssue,sL_RefNo : String;
    ii : Integer;
    fL_ComputeAmt : Double;
begin
    G_CompareNoStrList := TStringList.Create;
    G_ComputeParaStrList := TStringList.Create;
    fL_ComputeAmt := 0;

    if sI_SoOrProvider = '1' then        //ㄑ莱坝
    begin
      //L_CDS.Data := cdsTallyRefData1.Data;
      with cdsTallyRefData1 do
      begin
        //眔赣繵笵め计禯
        sL_RangeUnit := FieldByName('RANGE_UNIT').AsString;
        sL_ComputeIssue := FieldByName('COMPUTE_ISSUE').AsString;
      end;
    end
    else if sI_SoOrProvider = '2' then   //SO
    begin
      //L_CDS.Data := cdsTallyRefData2.Data;
      with cdsTallyRefData2 do
      begin
        //眔赣繵笵め计禯
        sL_RangeUnit := FieldByName('RANGE_UNIT').AsString;
        sL_ComputeIssue := FieldByName('COMPUTE_ISSUE').AsString;
      end;
    end;

    sL_RefNo := getFormulaRefNo(sI_FormulaID);

    //TEST Jackal 代刚ノ计沮
    //nG_TotalValidViewCounts := 100;
    //nI_ChannelTotalViewCounts := 6;

    //禯癸莱肂┪κだゑ
    getCalculateNo(sL_RangeUnit,sI_SoOrProvider,sL_ComputeIssue,nI_ChannelTotalViewCounts);

///////////
{
showmessage(IntToStr(G_CompareNoStrList.Count) + '---' + IntToStr(G_ComputeParaStrList.Count));
for ii:=0 to G_ComputeParaStrList.Count-1 do
begin
    showmessage(G_CompareNoStrList[ii] + '---' + G_ComputeParaStrList[ii]);
end;
}
//////////


    if sL_RefNo = '2' then                   //1(そΑ)め计计秖禯-虫基
    begin
      if sL_RangeUnit = '1' then             //2め计(そΑ﹚め计)
      begin

        if sL_ComputeIssue = '1' then        //3仓秈禯
          fL_ComputeAmt := calculateMulti(nI_ChannelTotalViewCounts,fI_OrigionalPrice,IS_STEP_NOT_PERCENT,IS_PRICE)
        else if sL_ComputeIssue = '2' then   //3辅翴禯
          fL_ComputeAmt := calculateSingle(nI_ChannelTotalViewCounts,fI_OrigionalPrice,IS_PRICE);

      end
      else
      begin
        Result := 'め计计秖禯-虫基,め计ゑ耕虫﹚璶砞め计';
        exit;
      end;
    end
    else if sL_RefNo = '4' then              //1(そΑ)め计κだゑ禯-虫基
    begin
      if sL_RangeUnit = '2' then             //2κだゑ(そΑ﹚κだゑ)
      begin

        if sL_ComputeIssue = '1' then        //3仓秈禯
          fL_ComputeAmt := calculateMulti(nI_ChannelTotalViewCounts,fI_OrigionalPrice,IS_STEP_PERCENT,IS_PRICE)
        else if sL_ComputeIssue = '2' then   //3辅翴禯
          fL_ComputeAmt := calculateSingle(nI_ChannelTotalViewCounts,fI_OrigionalPrice,IS_PRICE);

      end
      else
      begin
        Result := 'め计κだゑ禯-虫基,め计ゑ耕虫﹚璶砞κだゑ';
        exit;
      end;
    end
    else if sL_RefNo = '3' then              //1(そΑ)め计计秖禯-ゑㄒ
    begin
      if sL_RangeUnit = '1' then             //2め计(そΑ﹚め计)
      begin

        if sL_ComputeIssue = '1' then        //3仓秈禯
          fL_ComputeAmt := calculateMulti(nI_ChannelTotalViewCounts,fI_OrigionalPrice,IS_STEP_NOT_PERCENT,IS_NOT_PRICE)
        else if sL_ComputeIssue = '2' then   //3辅翴禯
          fL_ComputeAmt := calculateSingle(nI_ChannelTotalViewCounts,fI_OrigionalPrice,IS_NOT_PRICE);

      end
      else
      begin
        Result := 'め计计秖禯-ゑㄒ,め计ゑ耕虫﹚璶砞め计';
        exit;
      end;
    end
    else if sL_RefNo = '5' then              //1(そΑき)め计κだゑ禯-ゑㄒ
    begin
      if sL_RangeUnit = '2' then             //2κだゑ(そΑき﹚κだゑ)
      begin

        if sL_ComputeIssue = '1' then        //3仓秈禯
          fL_ComputeAmt := calculateMulti(nI_ChannelTotalViewCounts,fI_OrigionalPrice,IS_STEP_PERCENT,IS_NOT_PRICE)
        else if sL_ComputeIssue = '2' then   //3辅翴禯
          fL_ComputeAmt := calculateSingle(nI_ChannelTotalViewCounts,fI_OrigionalPrice,IS_NOT_PRICE);

      end
      else
      begin
        Result := 'め计κだゑ禯-ゑㄒ,め计ゑ耕虫﹚璶砞κだゑ';
        exit;
      end;
    end;

    fI_SoOrProviderBenefit := fL_ComputeAmt;

    G_CompareNoStrList.Clear;
    G_ComputeParaStrList.Clear;
end;



procedure TdtmMain2.getCalculateNo(sI_RangeUnit,sI_SoOrProvider,sI_ComputeIssue: String;nI_ChannelTotalViewCounts : Integer);
var
    fL_Subscriber_Count_Percent1,fL_Subscriber_Count_Percent2,fL_Subscriber_Count_Percent3 : Double;
    fL_Subscriber_Count_Percent4,fL_Subscriber_Count_Percent5 : Double;

    fL_Amount_Percent1,fL_Amount_Percent2,fL_Amount_Percent3 : Double;
    fL_Amount_Percent4,fL_Amount_Percent5 : Double;

    fL_ChannelCountsPercent,fL_CompareNo  : Double;
begin
    if sI_SoOrProvider = '1' then        //ㄑ莱坝
    begin
      with cdsTallyRefData1 do
      begin
        fL_Subscriber_Count_Percent1 := FieldByName('Subscriber_Count_Percent1').AsFloat;
        fL_Subscriber_Count_Percent2 := FieldByName('Subscriber_Count_Percent2').AsFloat;
        fL_Subscriber_Count_Percent3 := FieldByName('Subscriber_Count_Percent3').AsFloat;
        fL_Subscriber_Count_Percent4 := FieldByName('Subscriber_Count_Percent4').AsFloat;
        fL_Subscriber_Count_Percent5 := FieldByName('Subscriber_Count_Percent5').AsFloat;

        fL_Amount_Percent1 := FieldByName('Amount_Percent1').AsFloat;
        fL_Amount_Percent2 := FieldByName('Amount_Percent2').AsFloat;
        fL_Amount_Percent3 := FieldByName('Amount_Percent3').AsFloat;
        fL_Amount_Percent4 := FieldByName('Amount_Percent4').AsFloat;
        fL_Amount_Percent5 := FieldByName('Amount_Percent5').AsFloat;
      end;
    end
    else if sI_SoOrProvider = '2' then   //SO
    begin
      with cdsTallyRefData2 do
      begin
        fL_Subscriber_Count_Percent1 := FieldByName('Subscriber_Count_Percent1').AsFloat;
        fL_Subscriber_Count_Percent2 := FieldByName('Subscriber_Count_Percent2').AsFloat;
        fL_Subscriber_Count_Percent3 := FieldByName('Subscriber_Count_Percent3').AsFloat;
        fL_Subscriber_Count_Percent4 := FieldByName('Subscriber_Count_Percent4').AsFloat;
        fL_Subscriber_Count_Percent5 := FieldByName('Subscriber_Count_Percent5').AsFloat;

        fL_Amount_Percent1 := FieldByName('Amount_Percent1').AsFloat;
        fL_Amount_Percent2 := FieldByName('Amount_Percent2').AsFloat;
        fL_Amount_Percent3 := FieldByName('Amount_Percent3').AsFloat;
        fL_Amount_Percent4 := FieldByName('Amount_Percent4').AsFloat;
        fL_Amount_Percent5 := FieldByName('Amount_Percent5').AsFloat;
      end;
    end;

    if sI_RangeUnit = '1' then   //め计璸衡
    begin
      fL_CompareNo := nI_ChannelTotalViewCounts;
    end
    else if sI_RangeUnit = '2' then   //κだゑ璸衡
    begin
      //繵笵羆め计/羆Τめ计
      fL_ChannelCountsPercent := RoundTo(nI_ChannelTotalViewCounts/nG_TotalValidViewCounts*100,-2);
      fL_CompareNo := fL_ChannelCountsPercent;
    end;


    if fL_CompareNo <= fL_Subscriber_Count_Percent1 then
    begin
      //禯
      G_CompareNoStrList.Add(FloatToStr(fL_Subscriber_Count_Percent1));

      //璸衡把计
      G_ComputeParaStrList.Add(FloatToStr(fL_Amount_Percent1));
    end
    else if fL_CompareNo <= fL_Subscriber_Count_Percent2 then
    begin
      //禯
      G_CompareNoStrList.Add(FloatToStr(fL_Subscriber_Count_Percent1));
      G_CompareNoStrList.Add(FloatToStr(fL_Subscriber_Count_Percent2));

      //璸衡把计
      G_ComputeParaStrList.Add(FloatToStr(fL_Amount_Percent1));
      G_ComputeParaStrList.Add(FloatToStr(fL_Amount_Percent2));
    end
    else if fL_CompareNo <= fL_Subscriber_Count_Percent3 then
    begin
      //禯
      G_CompareNoStrList.Add(FloatToStr(fL_Subscriber_Count_Percent1));
      G_CompareNoStrList.Add(FloatToStr(fL_Subscriber_Count_Percent2));
      G_CompareNoStrList.Add(FloatToStr(fL_Subscriber_Count_Percent3));

      //璸衡把计
      G_ComputeParaStrList.Add(FloatToStr(fL_Amount_Percent1));
      G_ComputeParaStrList.Add(FloatToStr(fL_Amount_Percent2));
      G_ComputeParaStrList.Add(FloatToStr(fL_Amount_Percent3));
    end
    else if fL_CompareNo <= fL_Subscriber_Count_Percent4 then
    begin
      //禯
      G_CompareNoStrList.Add(FloatToStr(fL_Subscriber_Count_Percent1));
      G_CompareNoStrList.Add(FloatToStr(fL_Subscriber_Count_Percent2));
      G_CompareNoStrList.Add(FloatToStr(fL_Subscriber_Count_Percent3));
      G_CompareNoStrList.Add(FloatToStr(fL_Subscriber_Count_Percent4));

      //璸衡把计
      G_ComputeParaStrList.Add(FloatToStr(fL_Amount_Percent1));
      G_ComputeParaStrList.Add(FloatToStr(fL_Amount_Percent2));
      G_ComputeParaStrList.Add(FloatToStr(fL_Amount_Percent3));
      G_ComputeParaStrList.Add(FloatToStr(fL_Amount_Percent4));
    end
    else if fL_CompareNo <= fL_Subscriber_Count_Percent5 then
    begin
      //禯
      G_CompareNoStrList.Add(FloatToStr(fL_Subscriber_Count_Percent1));
      G_CompareNoStrList.Add(FloatToStr(fL_Subscriber_Count_Percent2));
      G_CompareNoStrList.Add(FloatToStr(fL_Subscriber_Count_Percent3));
      G_CompareNoStrList.Add(FloatToStr(fL_Subscriber_Count_Percent4));
      G_CompareNoStrList.Add(FloatToStr(fL_Subscriber_Count_Percent5));

      //璸衡把计
      G_ComputeParaStrList.Add(FloatToStr(fL_Amount_Percent1));
      G_ComputeParaStrList.Add(FloatToStr(fL_Amount_Percent2));
      G_ComputeParaStrList.Add(FloatToStr(fL_Amount_Percent3));
      G_ComputeParaStrList.Add(FloatToStr(fL_Amount_Percent4));
      G_ComputeParaStrList.Add(FloatToStr(fL_Amount_Percent5));
    end;
end;



function TdtmMain2.calculateMulti(nI_ChannelTotalViewCounts : Integer;fI_OrigionalPrice : Double;bI_IsPercent,bI_IsPrice : Boolean): Double;
var
    ii,nL_MaxStep,nL_NewCompareNo,nL_OldCompareNo : Integer;
    fL_ComputeAmt,fL_TotalComputeAmt : Double;
begin
    //璸衡そΑ┪┪┪き
    nL_MaxStep := G_CompareNoStrList.Count-1;
    nL_NewCompareNo := 0;
    nL_OldCompareNo := 0;
    fL_TotalComputeAmt := 0;


    //璸衡仓秈禯
    for ii:=0 to nL_MaxStep do
    begin
      nL_NewCompareNo := StrToInt(G_CompareNoStrList[ii]);

      if nL_MaxStep = 0 then
      begin
        if bI_IsPrice then  //虫基
          fL_TotalComputeAmt := nI_ChannelTotalViewCounts * StrToFloat(G_ComputeParaStrList[ii])
        else                //ゑㄒ
          fL_TotalComputeAmt := nI_ChannelTotalViewCounts * (fI_OrigionalPrice * StrToFloat(G_ComputeParaStrList[ii]) / 100);
      end
      else
      begin
        if ii = 0 then
        begin
          if not bI_IsPercent then   //..め计计秖禯
          begin
            if bI_IsPrice then  //そΑ:め计计秖禯-虫基
              fL_TotalComputeAmt := nL_NewCompareNo * StrToFloat(G_ComputeParaStrList[ii])
            else                //そΑ:め计计秖禯-ゑㄒ
              fL_TotalComputeAmt := nL_NewCompareNo * (fI_OrigionalPrice * StrToFloat(G_ComputeParaStrList[ii]) / 100);
          end
          else                       //..め计κだゑ禯
          begin
            if bI_IsPrice then  //そΑ:め计κだゑ禯-虫基
              fL_TotalComputeAmt := (nG_TotalValidViewCounts * nL_NewCompareNo / 100) * StrToFloat(G_ComputeParaStrList[ii])
            else                //そΑき:め计κだゑ禯-ゑㄒ
              fL_TotalComputeAmt := (nG_TotalValidViewCounts * nL_NewCompareNo / 100) * (fI_OrigionalPrice * StrToFloat(G_ComputeParaStrList[ii]) / 100);
          end;
        end
        else
        begin
          nL_OldCompareNo := StrToInt(G_CompareNoStrList[ii-1]);

          if ii <> nL_MaxStep then   //璝ゼ程禯
          begin
            if not bI_IsPercent then   //..め计计秖禯
            begin
              if bI_IsPrice then  //そΑ:め计计秖禯-虫基
                fL_ComputeAmt := (nL_NewCompareNo - nL_OldCompareNo) * StrToFloat(G_ComputeParaStrList[ii])
              else                //そΑ:め计计秖禯-ゑㄒ
                fL_ComputeAmt := (nL_NewCompareNo - nL_OldCompareNo) * (fI_OrigionalPrice * StrToFloat(G_ComputeParaStrList[ii]) /100)
            end
            else
            begin                     //..め计κだゑ禯
              if bI_IsPrice then  //そΑ:め计κだゑ禯-虫基
                fL_ComputeAmt := (nG_TotalValidViewCounts * (nL_NewCompareNo - nL_OldCompareNo) / 100) * StrToFloat(G_ComputeParaStrList[ii])
              else                //そΑき:め计κだゑ禯-ゑㄒ
                fL_ComputeAmt := (nG_TotalValidViewCounts * (nL_NewCompareNo - nL_OldCompareNo) / 100) * (fI_OrigionalPrice * StrToFloat(G_ComputeParaStrList[ii]) /100)
            end
          end
          else                      //璝程禯
          begin
            if not bI_IsPercent then   //..め计计秖禯
            begin
              if bI_IsPrice then  //そΑ:め计计秖禯-虫基
                fL_ComputeAmt := (nI_ChannelTotalViewCounts - nL_OldCompareNo) * StrToFloat(G_ComputeParaStrList[ii])
              else                //そΑ:め计计秖禯-ゑㄒ
                fL_ComputeAmt := (nI_ChannelTotalViewCounts - nL_OldCompareNo) * (fI_OrigionalPrice * StrToFloat(G_ComputeParaStrList[ii]) / 100)
            end
            else
            begin               //..め计κだゑ禯
              if bI_IsPrice then  //そΑ:め计κだゑ禯-虫基
                fL_ComputeAmt := (nG_TotalValidViewCounts * (nI_ChannelTotalViewCounts - nL_OldCompareNo) /100) * StrToFloat(G_ComputeParaStrList[ii])
              else                //そΑき:め计κだゑ禯-ゑㄒ
                fL_ComputeAmt := (nG_TotalValidViewCounts * (nI_ChannelTotalViewCounts - nL_OldCompareNo) /100) * (fI_OrigionalPrice * StrToFloat(G_ComputeParaStrList[ii]) /100)
            end
          end;

          fL_TotalComputeAmt := fL_TotalComputeAmt + fL_ComputeAmt;

        end;
      end;
    end;

    Result := fL_TotalComputeAmt;
end;

function TdtmMain2.calculateSingle(nI_ChannelTotalViewCounts: Integer;fI_OrigionalPrice : Double;bI_IsPrice : Boolean): Double;
var
    fL_TotalComputeAmt : Double;
    nL_MaxStep : Integer;
begin
    //璸衡辅翴禯
    nL_MaxStep := G_CompareNoStrList.Count-1;

    if bI_IsPrice then    //虫基
      fL_TotalComputeAmt := nI_ChannelTotalViewCounts * StrToFloat(G_ComputeParaStrList[nL_MaxStep])
    else                  //ゑㄒ
      fL_TotalComputeAmt := nI_ChannelTotalViewCounts * (fI_OrigionalPrice * StrToFloat(G_ComputeParaStrList[nL_MaxStep]) /100);

    Result := fL_TotalComputeAmt;
end;

function TdtmMain2.functionTwo(sI_CurDate, sI_CompCode, sI_StartDate,
  sI_EndDate, sI_BelongYM, sI_ShowDetail: String): String;
var
    ii,nL_CustCounts,nL_TotalSingleCounts,nL_TotalPackageCounts : Integer;
    sL_Table,sL_CitemString,sL_CitemCode,sL_CustID,sL_ProductID : String;
    sL_ProviderID,sL_FormulaID,sL_ChannelName,sL_Msg : String;
    fL_OrigionalPrice,fL_TotalChannelRealAmt : Double;
    nL_Ndx,nL_ChannelTotalViewCounts,nL_ChannelCustCounts : Integer;
    sL_UseBaseFormula,sL_ProviderDesc,sL_TempCitemCode : String;
    fL_ProviderPercent,fL_SoPercent,fL_Provider_Benefit,fL_So_Benefit,fL_Emc_Benefit : Double;
    fL_ORealCharge,fL_OEmc_Benefit,fL_OSo_Benefit,fL_OProvider_Benefit : Double;
    nL_OSingleCounts,nL_OPackageCounts,nL_OLastCount,nL_LastCount : Integer;
    sL_PrvBelongYM : String;
begin
    //眔才戈Μ禣兜ヘcdsRefDataA
    sL_CitemString := getCitemData(FUNCTION2);

    sL_PrvBelongYM := getPrvBelongYM(sI_BelongYM);

    //璸衡そΑ┪
    for ii:=0 to 1 do
    begin
      //セ V_ChargeInfoForTally 挡 SO033 の SO034 ,瞷э钡だ癸硂ㄢ Table 暗笆
      if ii = 0 then
        sL_Table := TABLE_SO033
      else if ii = 1 then
        sL_Table := TABLE_SO034;

      if sL_CitemString <> '' then
      begin
        nL_CustCounts := getCustData(sL_Table,sI_CompCode,sI_StartDate,sI_EndDate,sL_CitemString);
        nG_TotalCustCounts2 := nG_TotalCustCounts2 + nL_CustCounts;

        //эノ办跑计ㄓ羆そΑΜ跌计
        nG_TotalCustCounts := nG_TotalCustCounts + nL_CustCounts;
{  //JackalパそΑ1012
        if (nG_TotalCustCounts <=0) AND (sL_Table = TABLE_SO034) then
        begin
          Result := '⊿Τ才兵ンめ戈';
          exit;
        end
        else
        begin
}
          //璸衡–繵笵
          cdsRefDataA.First;
          while not cdsRefDataA.Eof do
          begin

            sL_CitemCode := cdsRefDataA.FieldByName('CITEMCODE').AsString;

            //浪琩Μ禣┪癶禣兜ヘ,肚癸莱Μ禣兜ヘ盿絏ㄓ
            //bL_IsInCome := checkIsInCome(sL_TempCitemCode,sL_CitemCode);

            sL_ProductID := getProductID(sL_CitemCode);

            sL_ChannelName := getChannelName(sL_ProductID);

            //繵笵ぇ羆Μ跌め计
            nL_Ndx := G_ChannelViewStrList.IndexOf(sL_ProductID);
            if nL_Ndx<>-1 then
               nL_ChannelTotalViewCounts := (G_ChannelViewStrList.Objects[nL_Ndx] as TChannelViewData).nChannelTotalViewCounts;


            getSO112Data(sL_ProductID,sL_ProviderID,sL_UseBaseFormula,fL_OrigionalPrice,fL_ProviderPercent,fL_SoPercent);

            //眔赣繵笵┮Τめ戈
            nL_ChannelCustCounts := getChannelCustData(sI_CompCode,sL_Table,sL_CitemCode,sI_StartDate,sI_EndDate);

            if nL_ChannelCustCounts > 0 then
            begin
              //盢赣繵笵–め┮Μ跌繵笵SO116い,璸衡繵笵羆龟Μ肂の繵笵虫翴の甅繺计秖
              fL_TotalChannelRealAmt := handleChannelCustData(sI_CompCode,sI_BelongYM,
                  sI_CurDate,sL_ChannelName,nL_TotalSingleCounts,nL_TotalPackageCounts);


              //т SO111 闽╊眀计ㄌ沮
              getSharePara(sI_CompCode,sL_ProductID,sL_ProviderID);

              //矪瞶そΑ┪繵笵╊眀戈
              sL_Msg := handleShareAmt2(sL_FormulaID,nL_ChannelTotalViewCounts,fL_OrigionalPrice,fL_Provider_Benefit,fL_So_Benefit);
              if sL_Msg <> '' then
              begin
                MessageDlg(sL_Msg,mtError, [mbOK],0);
                exit;
              end;


              //Emc龟悔┮眔 = 赣繵笵羆Μ(赣繵笵羆龟Μ肂) - ㄑ莱坝╊眀肂 - SO╊眀肂
              fL_Emc_Benefit := fL_TotalChannelRealAmt - fL_Provider_Benefit - fL_So_Benefit;

              sL_ProviderDesc := getProviderName(sL_ProviderID);


              if cdsSO114.Locate('COMP_ID;BELONGYM;PRODUCTID', VarArrayOf([sI_CompCode, sI_BelongYM,sL_ProductID]), []) then
              begin
                //繵笵盢–掸戈既 cdsSO114 仓
                with cdsSO114 do
                begin
                  //ㄓ肂
                  fL_ORealCharge := FieldByName('EMC_INCOME').AsFloat;
                  fL_OEmc_Benefit := FieldByName('EMC_BENEFIT').AsFloat;
                  fL_OSo_Benefit := FieldByName('SO_BENEFIT').AsFloat;
                  fL_OProvider_Benefit := FieldByName('PROVIDER_BENEFIT').AsFloat;
                  nL_OSingleCounts := FieldByName('SINGLECOUNTS').AsInteger;
                  nL_OPackageCounts := FieldByName('PACKAGECOUNTS').AsInteger;

                  //fL_OIncome := FieldByName('INCOME').AsFloat;
                  //fL_OOutcome := (FieldByName('OUTCOME').AsFloat * -1);

                  //nL_OLastCount := FieldByName('LASTCOUNT').AsInteger;

                  Delete;
                  Append;
                  //仓
                  FieldByName('COMP_ID').AsString := sI_CompCode;
                  FieldByName('BELONGYM').AsString := sI_BelongYM;
                  FieldByName('BEGIN_DATE').AsDateTime := StrToDateTime(sI_StartDate);
                  FieldByName('END_DATE').AsDateTime := StrToDateTime(sI_EndDate);
                  FieldByName('EMC_INCOME').AsFloat := RoundTo(fL_ORealCharge,0) + RoundTo(fL_TotalChannelRealAmt,0);

                  //FieldByName('INCOME').AsFloat := RoundTo(fL_Income,0) + RoundTo(fL_OIncome,0);
                  //FieldByName('OUTCOME').AsFloat := ((RoundTo(fL_Outcome,0) + RoundTo(fL_OOutcome,0)) * -1);

                  FieldByName('EMC_BENEFIT').AsFloat := RoundTo(fL_OEmc_Benefit,0) + RoundTo(fL_Emc_Benefit,0);
                  FieldByName('SO_BENEFIT').AsFloat := RoundTo(fL_OSo_Benefit,0) +  RoundTo(fL_So_Benefit,0);
                  FieldByName('PRODUCTID').AsString := sL_ProductID;

                  FieldByName('PROVIDERID').AsString := sL_ProviderID;
                  FieldByName('PROVIDERDESC').AsString := sL_ProviderDesc;

                  FieldByName('PROVIDER_BENEFIT').AsFloat := RoundTo(fL_OProvider_Benefit,0) + RoundTo(fL_Provider_Benefit,0);
                  FieldByName('SINGLECOUNTS').AsInteger := nL_OSingleCounts + nL_TotalSingleCounts;
                  FieldByName('PACKAGECOUNTS').AsInteger := nL_OPackageCounts + nL_TotalPackageCounts;

                  //FieldByName('LASTCOUNT').AsInteger := nL_OLastCount + nL_LastCount;

                  FieldByName('OPERATOR').AsString := frmMainMenu.sG_User;
                  FieldByName('UPDTIME').AsString := sI_CurDate;

                  Post;
                end;
              end
              else
              begin
                //繵笵盢–掸戈既 cdsSO114
                with cdsSO114 do
                begin
                  Append;

                  FieldByName('COMP_ID').AsString := sI_CompCode;
                  FieldByName('BELONGYM').AsString := sI_BelongYM;
                  FieldByName('BEGIN_DATE').AsDateTime := StrToDateTime(sI_StartDate);
                  FieldByName('END_DATE').AsDateTime := StrToDateTime(sI_EndDate);
                  FieldByName('EMC_INCOME').AsFloat := RoundTo(fL_TotalChannelRealAmt,0);

                  //FieldByName('INCOME').AsFloat := RoundTo(fL_Income,0);
                  //FieldByName('OUTCOME').AsFloat := (RoundTo(fL_Outcome,0) * -1);

                  FieldByName('EMC_BENEFIT').AsFloat := RoundTo(fL_Emc_Benefit,0);
                  FieldByName('SO_BENEFIT').AsFloat := RoundTo(fL_So_Benefit,0);
                  FieldByName('PRODUCTID').AsString := sL_ProductID;

                  FieldByName('PROVIDERID').AsString := sL_ProviderID;
                  FieldByName('PROVIDERDESC').AsString := sL_ProviderDesc;

                  FieldByName('PROVIDER_BENEFIT').AsFloat := RoundTo(fL_Provider_Benefit,0);
                  FieldByName('SINGLECOUNTS').AsInteger := nL_TotalSingleCounts;
                  FieldByName('PACKAGECOUNTS').AsInteger := nL_TotalPackageCounts;

                  //FieldByName('LASTCOUNT').AsInteger := nL_LastCount;

                  FieldByName('OPERATOR').AsString := frmMainMenu.sG_User;
                  FieldByName('UPDTIME').AsString := sI_CurDate;

                  Post;
                end;
              end;
            end;

            cdsRefDataA.Next;
          end;
        //end;  //JackalパそΑ1012
      end;
    end;

    //璸衡戳计秖,戳ソ计秖,セるΜ,セる癶禣 盢戈纗 cdsSO114
    if nG_TotalCustCounts2 > 0 then
      saveOtherDataTocdsSo114(sI_BelongYM,sL_PrvBelongYM,sI_CompCode,sI_StartDate,sI_EndDate,sL_CitemString);
end;

function TdtmMain2.functionThree(sI_CurDate, sI_CompCode, sI_StartDate,
  sI_EndDate, sI_BelongYM, sI_ShowDetail: String): String;
var
    ii,nL_CustCounts,nL_Ndx,nL_ChannelTotalViewCounts,nL_ChannelCustCounts : Integer;
    sL_Table,sL_CitemString,sL_CitemCode,sL_ProductID,sL_ChannelName : String;
    sL_FormulaID,sL_ProviderID,sL_UseBaseFormula,sL_Msg,sL_ProviderDesc : String;
    fL_OrigionalPrice,fL_ProviderPercent,fL_SoPercent : Double;
    fL_TotalChannelRealAmt,fL_Provider_Benefit,fL_So_Benefit,fL_Emc_Benefit : Double;
    nL_TotalSingleCounts,nL_TotalPackageCounts : Integer;
    nL_OSingleCounts,nL_OPackageCounts : Integer;
    fL_ORealCharge,fL_OEmc_Benefit,fL_OSo_Benefit,fL_OProvider_Benefit : Double;
    sL_PrvBelongYM : String;
begin
    //眔才戈Μ禣兜ヘcdsRefDataA
    sL_CitemString := getCitemData(FUNCTION3);

    sL_PrvBelongYM := getPrvBelongYM(sI_BelongYM);

    //璸衡そΑ┪き
    for ii:=0 to 1 do
    begin
      //セ V_ChargeInfoForTally 挡 SO033 の SO034 ,瞷э钡だ癸硂ㄢ Table 暗笆
      if ii = 0 then
        sL_Table := TABLE_SO033
      else if ii = 1 then
        sL_Table := TABLE_SO034;

      if sL_CitemString <> '' then
      begin
        nL_CustCounts := getCustData(sL_Table,sI_CompCode,sI_StartDate,sI_EndDate,sL_CitemString);
        nG_TotalCustCounts3 := nG_TotalCustCounts3 + nL_CustCounts;

        //эノ办跑计ㄓ羆そΑΜ跌计
        nG_TotalCustCounts := nG_TotalCustCounts + nL_CustCounts;

        if (nG_TotalCustCounts <=0) AND (sL_Table = TABLE_SO034) then
        begin
          Result := '⊿Τ才兵ンめ戈';
          exit;
        end
        else
        begin

          //璸衡–繵笵
          cdsRefDataA.First;
          while not cdsRefDataA.Eof do
          begin

            sL_CitemCode := cdsRefDataA.FieldByName('CITEMCODE').AsString;
            sL_ProductID := getProductID(sL_CitemCode);

            sL_ChannelName := getChannelName(sL_ProductID);

            //繵笵ぇ羆Μ跌め计
            nL_Ndx := G_ChannelViewStrList.IndexOf(sL_ProductID);
            if nL_Ndx<>-1 then
               nL_ChannelTotalViewCounts := (G_ChannelViewStrList.Objects[nL_Ndx] as TChannelViewData).nChannelTotalViewCounts;


            getSO112Data(sL_ProductID,sL_ProviderID,sL_UseBaseFormula,fL_OrigionalPrice,fL_ProviderPercent,fL_SoPercent);

            //眔赣繵笵┮Τめ戈
            nL_ChannelCustCounts := getChannelCustData(sI_CompCode,sL_Table,sL_CitemCode,sI_StartDate,sI_EndDate);

            if nL_ChannelCustCounts > 0 then
            begin
              //盢赣繵笵–め┮Μ跌繵笵SO116い,璸衡繵笵羆龟Μ肂の繵笵虫翴の甅繺计秖
              fL_TotalChannelRealAmt := handleChannelCustData(sI_CompCode,sI_BelongYM,
                  sI_CurDate,sL_ChannelName,nL_TotalSingleCounts,nL_TotalPackageCounts);



              //т SO111 闽╊眀计ㄌ沮
              getSharePara(sI_CompCode,sL_ProductID,sL_ProviderID);

              //矪瞶そΑ┪き繵笵╊眀戈
              sL_Msg := handleShareAmt2(sL_FormulaID,nL_ChannelTotalViewCounts,fL_OrigionalPrice,fL_Provider_Benefit,fL_So_Benefit);

              if sL_Msg <> '' then
              begin
                MessageDlg(sL_Msg,mtError, [mbOK],0);
                exit;
              end;


              //Emc龟悔┮眔 = 赣繵笵羆Μ(赣繵笵羆龟Μ肂) - ㄑ莱坝╊眀肂 - SO╊眀肂
              fL_Emc_Benefit := fL_TotalChannelRealAmt - fL_Provider_Benefit - fL_So_Benefit;

              sL_ProviderDesc := getProviderName(sL_ProviderID);


              if cdsSO114.Locate('COMP_ID;BELONGYM;PRODUCTID', VarArrayOf([sI_CompCode, sI_BelongYM,sL_ProductID]), []) then
              begin
                //繵笵盢–掸戈既 cdsSO114 仓
                with cdsSO114 do
                begin
                  //ㄓ肂
                  fL_ORealCharge := FieldByName('EMC_INCOME').AsFloat;
                  fL_OEmc_Benefit := FieldByName('EMC_BENEFIT').AsFloat;
                  fL_OSo_Benefit := FieldByName('SO_BENEFIT').AsFloat;
                  fL_OProvider_Benefit := FieldByName('PROVIDER_BENEFIT').AsFloat;
                  nL_OSingleCounts := FieldByName('SINGLECOUNTS').AsInteger;
                  nL_OPackageCounts := FieldByName('PACKAGECOUNTS').AsInteger;

                  Delete;
                  Append;
                  //仓
                  FieldByName('COMP_ID').AsString := sI_CompCode;
                  FieldByName('BELONGYM').AsString := sI_BelongYM;
                  FieldByName('BEGIN_DATE').AsDateTime := StrToDateTime(sI_StartDate);
                  FieldByName('END_DATE').AsDateTime := StrToDateTime(sI_EndDate);
                  FieldByName('EMC_INCOME').AsFloat := RoundTo(fL_ORealCharge,0) + RoundTo(fL_TotalChannelRealAmt,0);
                  FieldByName('EMC_BENEFIT').AsFloat := RoundTo(fL_OEmc_Benefit,0) + RoundTo(fL_Emc_Benefit,0);
                  FieldByName('SO_BENEFIT').AsFloat := RoundTo(fL_OSo_Benefit,0) +  RoundTo(fL_So_Benefit,0);
                  FieldByName('PRODUCTID').AsString := sL_ProductID;

                  FieldByName('PROVIDERID').AsString := sL_ProviderID;
                  FieldByName('PROVIDERDESC').AsString := sL_ProviderDesc;

                  FieldByName('PROVIDER_BENEFIT').AsFloat := RoundTo(fL_OProvider_Benefit,0) + RoundTo(fL_Provider_Benefit,0);
                  FieldByName('SINGLECOUNTS').AsInteger := nL_OSingleCounts + nL_TotalSingleCounts;
                  FieldByName('PACKAGECOUNTS').AsInteger := nL_OPackageCounts + nL_TotalPackageCounts;
                  FieldByName('OPERATOR').AsString := frmMainMenu.sG_User;
                  FieldByName('UPDTIME').AsString := sI_CurDate;

                  Post;
                end;
              end
              else
              begin
                //繵笵盢–掸戈既 cdsSO114
                with cdsSO114 do
                begin
                  Append;

                  FieldByName('COMP_ID').AsString := sI_CompCode;
                  FieldByName('BELONGYM').AsString := sI_BelongYM;
                  FieldByName('BEGIN_DATE').AsDateTime := StrToDateTime(sI_StartDate);
                  FieldByName('END_DATE').AsDateTime := StrToDateTime(sI_EndDate);
                  FieldByName('EMC_INCOME').AsFloat := RoundTo(fL_TotalChannelRealAmt,0);
                  FieldByName('EMC_BENEFIT').AsFloat := RoundTo(fL_Emc_Benefit,0);
                  FieldByName('SO_BENEFIT').AsFloat := RoundTo(fL_So_Benefit,0);
                  FieldByName('PRODUCTID').AsString := sL_ProductID;

                  FieldByName('PROVIDERID').AsString := sL_ProviderID;
                  FieldByName('PROVIDERDESC').AsString := sL_ProviderDesc;

                  FieldByName('PROVIDER_BENEFIT').AsFloat := RoundTo(fL_Provider_Benefit,0);
                  FieldByName('SINGLECOUNTS').AsInteger := nL_TotalSingleCounts;
                  FieldByName('PACKAGECOUNTS').AsInteger := nL_TotalPackageCounts;
                  FieldByName('OPERATOR').AsString := frmMainMenu.sG_User;
                  FieldByName('UPDTIME').AsString := sI_CurDate;

                  Post;
                end;
              end;
            end;
            cdsRefDataA.Next;
          end;
        end;
      end;

      //浪琩┮ΤそΑ琌Τ才兵ンめ
      if (nG_TotalCustCounts <=0) AND (sL_Table = TABLE_SO034) then
      begin
        Result := '⊿Τ才兵ンめ戈';
        exit;
      end
    end;

    //璸衡戳计秖,戳ソ计秖,セるΜ,セる癶禣 盢戈纗 cdsSO114
    if nG_TotalCustCounts3 > 0 then
      saveOtherDataTocdsSo114(sI_BelongYM,sL_PrvBelongYM,sI_CompCode,sI_StartDate,sI_EndDate,sL_CitemString);
end;

function TdtmMain2.chineseYMChangeToEnglishYM(sI_Date: String): String;
var
    L_StringList : TStringList;
    sL_Year,sL_Date : String;
begin
    L_StringList := TStringList.Create;
    L_StringList := TUstr.ParseStrings(sI_Date,'/');

    sL_Year  := IntToStr(StrToInt(L_StringList.Strings[0]) + 1911);

    sL_Date := sL_Year + L_StringList.Strings[1];

    Result := sL_Date;
end;

procedure TdtmMain2.saveToSo118(sI_CompCode, sI_BelongYM,sI_CurDate: String);
var
    sL_SQL : String;
begin
    sL_SQL := 'INSERT INTO SO118 VALUES(''' + sI_CompCode + ''',''' + sI_BelongYM +
              ''',''' + frmMainMenu.sG_User + ''',''' + sI_CurDate + ''')';

    with cdsCom do
    begin
      Close;
      CommandText := sL_SQL;
      Execute;
    end;

end;

function TdtmMain2.checkBelongYMIsLock(sI_CompCode,
  sI_BelongYM: String): Boolean;
var
    sL_SQL : String;
begin
    sL_SQL := 'SELECT * FROM SO118 WHERE COMP_ID=''' + sI_CompCode +
              ''' AND LOCKYM=''' + sI_BelongYM + '''';


    with cdsCom do
    begin
      Close;
      CommandText := sL_SQL;
      Open;

      if RecordCount > 0 then
        Result := true
      else
        Result := false;
    end;
end;

procedure TdtmMain2.saveToSO119(sI_CompCode, sI_BelongYM, sI_ProductID,
  sI_ProviderID, sI_CurDate: String; fI_ProviderPercent,
  fI_SoPercent: Double);
begin
    //璝ぃ狡玥穝糤掸
    if not cdsSO119.Locate('COMP_ID;BELONGYM;PRODUCT_ID', VarArrayOf([sI_CompCode, sI_BelongYM,sI_ProductID]), []) then
    begin
      //繵笵盢–掸戈既 cdsSO119
      with cdsSO119 do
      begin
        Append;

        FieldByName('COMP_ID').AsString := sI_CompCode;
        FieldByName('BELONGYM').AsString := sI_BelongYM;
        FieldByName('PRODUCT_ID').AsString := sI_ProductID;
        FieldByName('PROVIDER_ID').AsString := sI_ProviderID;
        FieldByName('PROVIDER_PERCENT').AsFloat := fI_ProviderPercent;
        FieldByName('SO_PERCENT').AsFloat := fI_SoPercent;
        FieldByName('OPERATOR').AsString := frmMainMenu.sG_User;
        FieldByName('UPDTIME').AsString := sI_CurDate;

        Post;
      end;
    end;
end;

function TdtmMain2.checkIsInCome(sI_TempCitemCode: String;
  var sI_CitemCode: String): Boolean;
var
    sL_RefNo,sL_ReturnCode : String;
begin

    if cdsCD019.Locate('CodeNo', sI_TempCitemCode, []) then
    begin

      sL_RefNo := cdsCD019.FieldByName('REFNO').AsString;

      if sL_RefNo = '9' then    //CD019 REFNO=9 癶禣兜ヘ
      begin
        sL_ReturnCode := cdsCD019.FieldByName('RETURNCODE').AsString;
        sI_CitemCode := sL_ReturnCode;    //璝癶禣兜ヘ斗肚痷タ癸莱Μ禣兜ヘ
        Result := false;
      end
      else                    //CD019 REFNO<>9 Μ兜ヘ
      begin
        sI_CitemCode := sI_TempCitemCode;
        Result := true;
      end;

    end;
end;

function TdtmMain2.getPrvBelongYM(sI_BelongYM: String): String;
var
    nL_Year,nL_Month : Integer;
    sL_LastYear,sL_LastMonth : String;
begin
    //璸衡る 200210-->>200209
    nL_Year := StrToInt(Copy(sI_BelongYM,1,4));
    nL_Month := StrToInt(Copy(sI_BelongYM,5,6));

    if nL_Month = 1 then
    begin
      sL_LastYear := IntToStr(nL_Year-1);
      sL_LastMonth := '12';
    end
    else
    begin
      sL_LastYear := IntToStr(nL_Year);
      sL_LastMonth := IntToStr(nL_Month-1);

      if Length(sL_LastMonth) = 1 then
        sL_LastMonth := '0' + sL_LastMonth;
    end;

    Result := sL_LastYear + sL_LastMonth;

end;

procedure TdtmMain2.getLastMonthComputePara(sI_CompCode, sI_PrvBelongYM,
  sI_ProductID: String; var fI_OldProviderPercent,
  fI_OldSoPercent: Double);
begin
    //璝琩礚才戈肚-1
    fI_OldProviderPercent := -1;
    fI_OldSoPercent := -1;

    //眔るㄑ莱坝のSO╊眀ゑㄒ
    if cdsTempSO119.Locate('COMP_ID;BELONGYM;PRODUCT_ID', VarArrayOf([sI_CompCode, sI_PrvBelongYM,sI_ProductID]), []) then
    begin
       fI_OldProviderPercent := cdsTempSO119.FieldByName('PROVIDER_PERCENT').AsFloat;
       fI_OldSoPercent := cdsTempSO119.FieldByName('SO_PERCENT').AsFloat;
    end;
end;

function TdtmMain2.getOutComeCitemData(sI_CitemString : String): String;
var
    sL_SQL,sL_CitemString,sL_CitemCode,sL_TempCitemCode : String;
    L_CitemString : TStringList;
    ii : Integer;
begin
    //眔癶禣CitemCode   REFNO=9癶禣
    sL_SQL := 'SELECT CODENO FROM CD019 WHERE REFNO=9 AND STOPFLAG=0' +
              ' AND RETURNCODE IN (' + sI_CitemString + ')';

    with cdsCom do
    begin
      Close;
      CommandText := sL_SQL;
      Open;
    end;

    //盢Μ禣兜ヘ絏﹃Θ﹃
    sL_CitemString := '';
    L_CitemString := TStringList.Create;
    cdsCom.First;
    while not cdsCom.Eof do
    begin
      sL_CitemCode := cdsCom.FieldByName('CODENO').AsString;

      if sL_CitemString = '' then
      begin
        L_CitemString.Add(sL_CitemCode);
        sL_CitemString := sL_CitemCode;
      end
      else
      begin
        for ii:=0 to L_CitemString.count-1 do
        begin
          sL_TempCitemCode := L_CitemString.Strings[ii];
          if sL_TempCitemCode = sL_CitemCode then
            break;

          if ii = L_CitemString.count-1 then
          begin
            if sL_TempCitemCode <> sL_CitemCode then
            begin
              L_CitemString.Add(sL_CitemCode);

              sL_CitemString := sL_CitemString + ',' + sL_CitemCode;
            end;
          end;
        end;
      end;

      cdsCom.Next;
    end;

    Result := sL_CitemString;
end;


function TdtmMain2.functionOne(sI_CurDate, sI_CompCode, sI_StartDate,
  sI_EndDate, sI_BelongYM, sI_ShowDetail: String) : String;
var
    sL_CitemString,sL_CustID,sL_CitemCode,sL_ProductID,sL_PackageID : String;
    ii,nL_CustCounts : Integer;
    sL_FormulaID,sL_ProviderID,sL_CompType,sL_SQL,sL_OutComeCitemString : String;
    fL_RealCharge,fL_OrigionalPrice : Double;
    sL_ChannelString,sL_ChannelViewCountsString,sL_Table : String;
    fL_ProviderPercent,fL_SoPercent : Double;
    sL_UseBaseFormula,sL_TempCitemCode,sL_PrvBelongYM,sL_TotalCitemString : String;
    bL_IsInCome : Boolean;
    sL_FirstFlag : String;
begin
    //眔才戈Μ禣兜ヘcdsRefDataA
    sL_CitemString := getCitemData(FUNCTION1);
    //そΑ璸衡Μ禣兜ヘ
{
    //眔癶禣兜ヘ
    if sL_CitemString <> '' then
      sL_OutComeCitemString := getOutComeCitemData(sL_CitemString);

    if (sL_CitemString <> '') and (sL_OutComeCitemString <> '') then
      sL_TotalCitemString := sL_CitemString + ',' + sL_OutComeCitemString
    else if sL_CitemString = '' then
      sL_TotalCitemString := sL_OutComeCitemString;
}
    sL_PrvBelongYM := getPrvBelongYM(sI_BelongYM);

    //璸衡そΑ
    for ii:=0 to 1 do
    begin
      //セ V_ChargeInfoForTally 挡 SO033 の SO034 ,瞷э钡だ癸硂ㄢ Table 暗笆
      if ii = 0 then
        sL_Table := TABLE_SO033
      else if ii = 1 then
        sL_Table := TABLE_SO034;

      {
      if sL_CitemString = '' then
      begin
        Result := '⊿Τ才兵ンΜ禣絏';
        exit;
      end
      else
       }
      if sL_CitemString <> '' then
      begin
        //眔才龟Μら戳絛瞅ず┮Τめ
        nL_CustCounts := getCustData(sL_Table,sI_CompCode,sI_StartDate,sI_EndDate,sL_CitemString);
        nG_TotalCustCounts1 := nG_TotalCustCounts1 + nL_CustCounts;

        //эノ办跑计ㄓ羆そΑΜ跌计
        nG_TotalCustCounts := nG_TotalCustCounts + nL_CustCounts;
{  //JackalパそΑ1012
        if (nG_TotalCustCounts <=0) AND (sL_Table = TABLE_SO034) then
        begin
          Result := '⊿Τ才兵ンめ戈';
          exit;
        end
        else
        begin
}
          cdsCustData.First;
          while not cdsCustData.Eof do
          begin
            sL_CustID := cdsCustData.FieldByName('CUSTID').AsString;

            //眔赣めΜ禣戈
            getChargeData(sL_Table,sL_CustID,sI_CompCode,sI_StartDate,sI_EndDate,sL_CitemString);
            cdsChargeData.First;

            //睲
            cdsProductInfo.EmptyDataSet;
            cdsPackageInfo.EmptyDataSet;

            while not cdsChargeData.Eof do
            begin
              sL_TempCitemCode := cdsChargeData.FieldByName('CITEMCODE').AsString;

              //浪琩Μ禣┪癶禣兜ヘ,肚癸莱Μ禣兜ヘ盿絏ㄓ
              bL_IsInCome := checkIsInCome(sL_TempCitemCode,sL_CitemCode);

              sL_ProductID := getProductID(sL_CitemCode);

              fL_RealCharge := cdsChargeData.FieldByName('REALAMT').AsFloat;


              //浪琩琌甅繺
              sL_PackageID := cdsChargeData.FieldByName('PACKAGENO').AsString;

              //琌戳
              sL_FirstFlag := cdsChargeData.FieldByName('FirstFlag').AsString;

              getSO112Data(sL_ProductID,sL_ProviderID,sL_UseBaseFormula,fL_OrigionalPrice,fL_ProviderPercent,fL_SoPercent);

              //盢赣掸めΜ禣戈既 cdsProductInfo
              saveTocdsProductInfo(sL_FirstFlag,sL_ProductID,sL_PackageID,sL_FormulaID,sL_ProviderID,sL_UseBaseFormula,fL_RealCharge,fL_OrigionalPrice,fL_ProviderPercent,fL_SoPercent,bL_IsInCome);

              //赣掸めΜ禣戈璝妮甅繺,玥既 dsPackageInfo
              saveTocdsPackageInfo(sL_PackageID,fL_RealCharge,fL_OrigionalPrice);

              cdsChargeData.Next;
            end;

            //矪瞶めΜ禣╊眀戈
            handleShareAmt1(sI_CompCode,sL_CustID,sI_StartDate,sI_EndDate,sI_BelongYM,sI_CurDate,sI_ShowDetail,sL_PrvBelongYM,sL_ChannelString);

            //盢戈╊眀穨ぇめΜ跌繵笵灿郎
            saveTocdsSo116(sI_CompCode,sI_BelongYM,sL_CustID,sI_CurDate,sL_ChannelString);

            cdsCustData.Next;
          end;
        //end; //JackalパそΑ1012

      end;
    end;


    //璸衡戳计秖,戳ソ计秖,セるΜ,セる癶禣 盢戈纗 cdsSO114
    if nG_TotalCustCounts1 > 0 then
      saveOtherDataTocdsSo114(sI_BelongYM,sL_PrvBelongYM,sI_CompCode,sI_StartDate,sI_EndDate,sL_CitemString);
end;


function TdtmMain2.functionBasic(sI_CurDate, sI_CompCode, sI_StartDate,
  sI_EndDate, sI_BelongYM, sI_ShowDetail: String): String;
var
    sL_OutComeCitemString,sL_CitemString,sL_PrvBelongYM,sL_Table : String;
    sL_CustID,sL_TempCitemCode,sL_CitemCode,sL_ProductID,sL_PackageID : String;
    sL_FormulaID,sL_ProviderID,sL_UseBaseFormula,sL_ChannelString : String;
    fL_RealCharge,fL_OrigionalPrice,fL_ProviderPercent,fL_SoPercent : Double;
    ii,nL_CustCounts : Integer;
    bL_IsInCome : Boolean;
    sL_FirstFlag : String;
begin

    //眔才戈Μ禣兜ヘcdsRefDataA
    sL_CitemString := getCitemData(FUNCTION0);

    //眔癶禣兜ヘ
    if sL_CitemString <> '' then
      sL_OutComeCitemString := getOutComeCitemData(sL_CitemString);

    if (sL_CitemString <> '') and (sL_OutComeCitemString <> '') then
      sL_CitemString := sL_CitemString + ',' + sL_OutComeCitemString
    else if sL_CitemString = '' then
      sL_CitemString := sL_OutComeCitemString;

    sL_PrvBelongYM := getPrvBelongYM(sI_BelongYM);

    //璸衡膀セそΑ
    for ii:=0 to 1 do
    begin
      //セ V_ChargeInfoForTally 挡 SO033 の SO034 ,瞷э钡だ癸硂ㄢ Table 暗笆
      if ii = 0 then
        sL_Table := TABLE_SO033
      else if ii = 1 then
        sL_Table := TABLE_SO034;

      if sL_CitemString <> '' then
      begin
        //眔才龟Μら戳絛瞅ず┮Τめ
        nL_CustCounts := getCustData(sL_Table,sI_CompCode,sI_StartDate,sI_EndDate,sL_CitemString);
        //nG_TotalCustCounts1 := nG_TotalCustCounts1 + nL_CustCounts;

        //эノ办跑计ㄓ羆そΑΜ跌计
        nG_TotalCustCounts := nG_TotalCustCounts + nL_CustCounts;

          cdsCustData.First;
          while not cdsCustData.Eof do
          begin
            sL_CustID := cdsCustData.FieldByName('CUSTID').AsString;

            //眔赣めΜ禣戈
            getChargeData(sL_Table,sL_CustID,sI_CompCode,sI_StartDate,sI_EndDate,sL_CitemString);
            cdsChargeData.First;

            //睲
            cdsProductInfo.EmptyDataSet;
            cdsPackageInfo.EmptyDataSet;

            while not cdsChargeData.Eof do
            begin
              sL_TempCitemCode := cdsChargeData.FieldByName('CITEMCODE').AsString;

              //浪琩Μ禣┪癶禣兜ヘ,肚癸莱Μ禣兜ヘ盿絏ㄓ
              bL_IsInCome := checkIsInCome(sL_TempCitemCode,sL_CitemCode);

              sL_ProductID := getProductID(sL_CitemCode);
              fL_RealCharge := cdsChargeData.FieldByName('REALAMT').AsFloat;


              //浪琩琌甅繺
              sL_PackageID := cdsChargeData.FieldByName('PACKAGENO').AsString;

              //琌戳
              sL_FirstFlag := cdsChargeData.FieldByName('FirstFlag').AsString;

              getSO112Data(sL_ProductID,sL_ProviderID,sL_UseBaseFormula,fL_OrigionalPrice,fL_ProviderPercent,fL_SoPercent);

              //盢赣掸めΜ禣戈既 cdsProductInfo
              saveTocdsProductInfo(sL_FirstFlag,sL_ProductID,sL_PackageID,sL_FormulaID,sL_ProviderID,sL_UseBaseFormula,fL_RealCharge,fL_OrigionalPrice,fL_ProviderPercent,fL_SoPercent,bL_IsInCome);

              //赣掸めΜ禣戈璝妮甅繺,玥既 dsPackageInfo
              saveTocdsPackageInfo(sL_PackageID,fL_RealCharge,fL_OrigionalPrice);

              cdsChargeData.Next;
            end;


            //矪瞶めΜ禣╊眀戈
            handleShareAmt1(sI_CompCode,sL_CustID,sI_StartDate,sI_EndDate,sI_BelongYM,sI_CurDate,sI_ShowDetail,sL_PrvBelongYM,sL_ChannelString);

            //盢戈╊眀穨ぇめΜ跌繵笵灿郎
            saveTocdsSo116(sI_CompCode,sI_BelongYM,sL_CustID,sI_CurDate,sL_ChannelString);

            cdsCustData.Next;
          end;
      end;
    end;
end;

function TdtmMain2.getFirstCounts(sI_CompCode, sI_PrvBelongYM,
  sI_ProductID: String): Integer;
var
    nL_FirstCounts : Integer;
begin
    //眔赣繵笵戳计秖
    if cdsSO114C.Locate('COMP_ID;BELONGYM;PRODUCTID', VarArrayOf([sI_CompCode, sI_PrvBelongYM,sI_ProductID]), []) then
      nL_FirstCounts := cdsSO114C.FieldByName('LASTCOUNT').AsInteger
    else
      nL_FirstCounts := 0;

    Result := nL_FirstCounts;
end;



function TdtmMain2.getTotalInOrOutCome(sI_CompCode, sI_StartDate,
  sI_EndDate, sI_CitemCode: String;var nI_TotalInOrOutComeCounts :Integer): Double;
var
    sL_SQL,sL_Table : String;
    fL_InOrOutCome,fL_TotalInOrOutCome : Double;
    ii,nL_InOrOutComeCounts,nL_TotalInOrOutComeCounts : Integer;
begin
    nL_InOrOutComeCounts := 0;
    nL_TotalInOrOutComeCounts := 0;
    fL_InOrOutCome:= 0;
    fL_TotalInOrOutCome := 0;

    for ii:=0 to 1 do
    begin
      if ii = 0 then
        sL_Table := TABLE_SO033
      else if ii = 1 then
        sL_Table := TABLE_SO034;

      sL_SQL := 'SELECT SUM(NVL(REALAMT,0)) TotalInOrOutCome FROM ' + sL_Table + ' WHERE CITEMCODE =' + sI_CitemCode +
                ' AND (REALDATE BETWEEN ' + TUstr.getOracleSQLDateStr(StrToDate(sI_StartDate)) +
                ' AND ' + TUstr.getOracleSQLDateStr(StrToDate(sI_EndDate)) + ')' +
                ' AND COMPCODE=' + sI_CompCode;

      with cdsCom do
      begin
        Close;
        CommandText := sL_SQL;
        Open;

        fL_InOrOutCome := cdsCom.FieldByName('TotalInOrOutCome').AsFloat;
        fL_TotalInOrOutCome := fL_TotalInOrOutCome + fL_InOrOutCome;
      end;
      sL_SQL := 'SELECT COUNT(*) TotalInOrOutComeCounts  FROM ' + sL_Table + ' WHERE CITEMCODE =' + sI_CitemCode +
                ' AND (REALDATE BETWEEN ' + TUstr.getOracleSQLDateStr(StrToDate(sI_StartDate)) +
                ' AND ' + TUstr.getOracleSQLDateStr(StrToDate(sI_EndDate)) + ')' +
                ' AND COMPCODE=' + sI_CompCode;

      with cdsCom do
      begin
        Close;
        CommandText := sL_SQL;
        Open;

        nL_InOrOutComeCounts := cdsCom.FieldByName('TotalInOrOutComeCounts').AsInteger;
        nL_TotalInOrOutComeCounts := nL_TotalInOrOutComeCounts + nL_InOrOutComeCounts;
      end;
    end;

    nI_TotalInOrOutComeCounts  := nL_TotalInOrOutComeCounts;
    Result := fL_TotalInOrOutCome ;

end;

procedure TdtmMain2.saveOtherDataTocdsSo114(sI_BelongYM,sI_PrvBelongYM,sI_CompCode, sI_StartDate, sI_EndDate,sI_TotalCitemString: String);
var
    sL_OutComeCitemCode,sL_CitemCode,sL_ProductID : String;
    L_CitemStrList : TStringList;
    ii,nL_TotalInComeCounts,nL_LastCount,nL_TotalOutComeCounts,nL_FirstCount : Integer;
    fL_TotalInCome,fL_TotalOutCome : Double;
begin
    L_CitemStrList := TStringList.Create;
    L_CitemStrList := TUstr.ParseStrings(sI_TotalCitemString,',');

    for ii:=0 to L_CitemStrList.Count-1 do
    begin
      sL_CitemCode := L_CitemStrList.Strings[ii];
      //眔癶禣兜ヘ絏
      sL_OutComeCitemCode := getOutComeCitemData(sL_CitemCode);

      sL_ProductID := getProductID(sL_CitemCode);

      //眔赣繵笵 (Μ羆の癶禣羆) 籔 (Μの癶禣め计羆)
      fL_TotalInCome := getTotalInOrOutCome(sI_CompCode, sI_StartDate, sI_EndDate,sL_CitemCode,nL_TotalInComeCounts);
      fL_TotalOutCome := getTotalInOrOutCome(sI_CompCode, sI_StartDate, sI_EndDate,sL_OutComeCitemCode,nL_TotalOutComeCounts);

      //セ戳め计 = Μめ计
      nL_LastCount := nL_TotalInComeCounts;

      //眔戳计秖
      nL_FirstCount := getFirstCounts(sI_CompCode, sI_PrvBelongYM,sL_ProductID);

      if cdsSO114.Locate('COMP_ID;BELONGYM;PRODUCTID', VarArrayOf([sI_CompCode, sI_BelongYM,sL_ProductID]), []) then
      begin
        //璸衡そΑノㄤよΑ衡
        cdsSO114.Edit;

        cdsSO114.FieldByName('INCOME').AsFloat := RoundTo(fL_TotalInCome,0);
        cdsSO114.FieldByName('OUTCOME').AsFloat := RoundTo(fL_TotalOutCome,0);
        cdsSO114.FieldByName('FIRSTCOUNT').AsInteger := nL_FirstCount;
        cdsSO114.FieldByName('LASTCOUNT').AsInteger := nL_LastCount;

        cdsSO114.Post;
      end;
//showmessage(sI_CitemCode + '--' + IntToStr(nL_TotalInOrOutComeCounts) + '--' + FloatToStr(fL_TotalInOrOutCome));
    end;
    L_CitemStrList.Free;
end;

function TdtmMain2.getSO110RefNo(sI_Formula_ID: String): String;
var
    sL_SQL : String;
begin
    //ゑ癸そΑ絏そΑ
    sL_SQL := 'SELECT REF_NO FROM SO110 WHERE FORMULA_ID=''' + sI_Formula_ID + '''';

    with cdsCom do
    begin
      Close;
      CommandText := sL_SQL;
      Open;

      Result := FieldByName('REF_NO').AsString
    end;
end;

procedure TdtmMain2.getSO113AndCD039(sI_ProviderID: String);
var
    sL_SQL : String;
begin
    //そ/ㄑ莱坝 APPEND  cdsSo113Cd039
    with cdsCD039 do
    begin
      sL_SQL := 'SELECT CODENO,DESCRIPTION FROM CD039';
      Close;
      CommandText:=(sL_SQL);
      Open;

      First;
      cdsSo113Cd039.EmptyDataSet;
      While not cdsCD039.Eof do
      begin
        try
          cdsSo113Cd039.Append;
          cdsSo113Cd039.FieldByName('CLASSIFY_ID').AsString:='1';
          cdsSo113Cd039.FieldByName('CODE').AsString:=cdsCD039.FieldByName('CODENO').AsString;
          cdsSo113Cd039.FieldByName('DESC').AsString:=cdsCD039.FieldByName('DESCRIPTION').AsString;
          cdsSo113Cd039.Post;
          cdsCD039.Next;
        except
        end;
      end;

    end;



    with cdsSo113 do
    begin
      sL_SQL := 'SELECT PROVIDER_ID,PROVIDER_NAME FROM SO113 WHERE PROVIDER_ID=''' + sI_ProviderID + '''';
      Close;
      CommandText:=(sL_SQL);
      open;

      First;
      While not cdsSo113.Eof do
      begin
        try
          cdsSo113Cd039.Append;
          cdsSo113Cd039.FieldByName('CLASSIFY_ID').AsString:='2';
          cdsSo113Cd039.FieldByName('CODE').AsString:=cdsSo113.FieldByName('PROVIDER_ID').AsString;
          cdsSo113Cd039.FieldByName('DESC').AsString:=cdsSo113.FieldByName('PROVIDER_NAME').AsString;
          cdsSo113Cd039.Post;
          cdsSo113.Next;
        except
        end;
      end;
    end;
end;

function TdtmMain2.getProductName(sI_ProductID: String): String;
var
    sL_SQL,sL_ProductID : String;
begin
    sL_SQL := 'SELECT * FROM So112 WHERE PRODUCT_ID=''' + sI_ProductID + '''';

    with dtmMain2.cdsCom do
    begin
      Close;
      CommandText := sL_SQL;
      Open;
    end;

    Result := cdsCom.FieldByName('PRODUCT_NAME').AsString;
end;

procedure TdtmMain2.gerAllProvider;
var
    sL_SQL,sL_ProviderID : String;
begin
    G_ProviderStrList := TStringList.Create;

    sL_SQL:='SELECT * FROM SO113 WHERE STOPFLAG=0';

    with dtmMain2.cdsCom do
    begin
      Close;
      CommandText := sL_SQL;
      Open;

      First;
      while not Eof do
      begin
        sL_ProviderID := dtmMain2.cdsCom.FieldByName('Provider_ID').AsString;
        G_ProviderStrList.Add(sL_ProviderID);
        Next;
      end;
    end;
end;

procedure TdtmMain2.getCodeCD019;
var
    sL_SQL : String;
begin
    //琩高 CD019 Μ禣兜ヘ絏郎┮Τ戈
    sL_SQL := 'SELECT CODENO,DESCRIPTION FROM CD019 WHERE ' +
              ' SERVICETYPE=''' + frmMainMenu.sG_ServiceType + ''' AND REFNO=2 ';
    with cdsCodeCD019 do
    begin
      Close;
      CommandText := sL_SQL;
      Open;
    end;

end;

function TdtmMain2.calculateDetailData(sI_CompCode, sI_ChargeItemSQL, sI_ComputeYM : String; nI_RealOrShouldDate: Integer): String;
var
    sL_SQL,sL_Where,sL_ComputeYM,sL_Table : String;
    sL_StartDate,sL_EndDate,sL_MonthDays,sL_Msg : String;
    fL_RetCode : Double;
    ii : Integer;
begin
    sL_ComputeYM := Trim(dtmMain2.ReplaceStr(sI_ComputeYM,'/'));

    sL_StartDate := TransToEngDate(sI_ComputeYM + '/01');
    sL_MonthDays := IntToStr(TUdateTime.DaysOfMonth(sL_StartDate));
    sL_EndDate := TransToEngDate(sL_ComputeYM  + sL_MonthDays);
    sL_StartDate := Trim(dtmMain2.ReplaceStr(sL_StartDate,'/'));


    //璸衡タ兜Μ禣兜ヘ灿
    runSFDetail(sL_ComputeYM,sI_ChargeItemSQL,sL_StartDate,sL_EndDate,nI_RealOrShouldDate,fL_RetCode,sL_Msg);

    if sL_Msg = '' then
    begin
{

      for ii:=1 to 2 do
      begin
        if ii = 1 then
          sL_Table := 'SO033'
        else if ii = 2 then
          sL_Table := 'SO034';

        //璸衡璽兜Μ禣兜ヘ灿
        getBackDetailData(sL_Table,sI_CompCode, sI_ComputeYM,nI_RealOrShouldDate);

        //俱タ璽Μ禣兜ヘ
        calculateCompleteDetailData(sL_Table,sL_ComputeYM,nI_RealOrShouldDate);
      end;

      //盢 SeqNo 腹
      addSeqNo;

      if cdsSO131.State in [dsInsert, dsEdit] then
        cdsSO131.Post;
      if cdsSO131.ChangeCount>0 then
        cdsSO131.ApplyUpdates(0);
}        
    end
    else
    begin
      MessageDlg(sL_Msg,mtError, [mbOK],0);
    end;
end;

function TdtmMain2.TransToEngDate(sI_Date: String): String;
var
    sL_Year,sL_EnglishYear,sL_MonthDay,sL_EngDate : String;

    nL_Year,sL_Len : Integer;
begin
    sL_Len := Length(sI_Date);
    nL_Year := StrToInt(Copy(sI_Date,1,3));
    sL_MonthDay := Copy(sI_Date,4,sL_Len);
    sL_EnglishYear := IntToStr(nL_Year + 1911);

    sL_EngDate := sL_EnglishYear + sL_MonthDay;

    Result := sL_EngDate;

end;



procedure TdtmMain2.getSO105Data(sI_OrderNo: String;var sI_MediaCode,
sI_MediaName,sI_AcceptEn,sI_AcceptName,sI_IntroId,sI_IntroName : String);
var
    sL_SQL : String;
begin
    sL_SQL := 'Select MediaCode,MediaName,AcceptEn,AcceptName,IntroId,IntroName '
            + ' from So105 where OrderNo=''' + sI_OrderNo + '''';


    with cdsCom do
    begin
      Close;
      CommandText := sL_SQL;
      Open;

      sI_MediaCode := FieldByName('MediaCode').AsString;
      sI_MediaName := FieldByName('MediaName').AsString;
      sI_AcceptEn := FieldByName('AcceptEn').AsString;
      sI_AcceptName := FieldByName('AcceptName').AsString;
      sI_IntroId := FieldByName('IntroId').AsString;
      sI_IntroName := FieldByName('IntroName').AsString;

    end;
end;

function TdtmMain2.checkHaveSO131Data(sI_ComputeYM: String;
  nI_RealOrShouldDate: Integer): Boolean;
var
    sL_SQL,sL_ComputeYM : String;

begin
    sL_ComputeYM := Trim(dtmMain2.ReplaceStr(sI_ComputeYM,'/'));

    sL_SQL := 'SELECT * FROM SO131 WHERE COMPCODE=' + frmMainMenu.sG_CompCode
              + ' AND COMPUTEYM=''' + sL_ComputeYM
              + ''' AND RealOrShouldDate=' + IntToStr(nI_RealOrShouldDate);


    with cdsCom do
    begin
      Close;
      CommandText := sL_SQL;
      Open;

      if RecordCount = 0 then
        Result := false
      else
        Result := true;
    end;
end;

function TdtmMain2.runSFDetail(sI_ComputeYM, sI_ChargeItemSQL,sI_StartDate,sI_EndDate : String;
  nI_RealOrShouldDate: Integer;var fI_RetCode : Double;var sI_Msg : String): String;
var
    ii : Integer;
    L_Intf : IEmc_Separate1;
    sL_RetCode,sL_Msg : OleVariant;
begin
    dtmMain2.connToServer;
    //if not dtmMain2.DCOMConnection1.Connected then
    //  dtmMain2.DCOMConnection1.Connected := true;

    L_Intf := IUnknown(dtmMain2.DCOMConnection1.AppServer) as IEmc_Separate1;

    L_Intf.runSF_CALCULATESO131(frmMainMenu.sG_CompCode,frmMainMenu.sG_ServiceType,sI_ComputeYM,
    sI_StartDate,sI_EndDate,sI_ChargeItemSQL,nI_RealOrShouldDate,
    sL_RetCode,sL_Msg);

end;

function TdtmMain2.getBackCitemCode: String;
var
    sL_SQL,sL_CitemCode,sL_AllCitemCode : String;
begin
    //琩高 CD019 繵笵癶禣Μ禣兜ヘ絏郎┮Τ戈
    sL_AllCitemCode := '';

    sL_SQL := 'SELECT CODENO,DESCRIPTION FROM CD019 WHERE ' +
              'SERVICETYPE=''' + frmMainMenu.sG_ServiceType + ''' AND REFNO=9 ';

    with cdsCom do
    begin
      Close;
      CommandText := sL_SQL;
      Open;


      First;
      while not cdsCom.Eof do
      begin
        sL_CitemCode := FieldByName('CODENO').AsString;

        if sL_AllCitemCode = '' then
          sL_AllCitemCode := sL_CitemCode
        else
          sL_AllCitemCode := sL_AllCitemCode + ',' + sL_CitemCode;

        Next;
      end;
    end;

    Result := sL_AllCitemCode;
end;

function TdtmMain2.getBackDetailData(sI_Table,sI_CompCode,
  sI_ComputeYM: String; nI_RealOrShouldDate: Integer): String;
var
    sL_BackCitemCodeString,sL_SQL,sL_Where : String;
    sL_ComputeYM,sL_StartDate,sL_MonthDays,sL_EndDate : String;
    ii : Integer;
begin
    sL_ComputeYM := Trim(dtmMain2.ReplaceStr(sI_ComputeYM,'/'));

    sL_StartDate := TransToEngDate(sI_ComputeYM + '/01');
    sL_MonthDays := IntToStr(TUdateTime.DaysOfMonth(sL_StartDate));

    sL_EndDate := TransToEngDate(sI_ComputeYM + '/' + sL_MonthDays + ' 23:59:59');
    sL_StartDate := sL_StartDate + ' 00:00:01';

    sL_BackCitemCodeString := getBackCitemCode;


    sL_SQL := 'SELECT * FROM ' + sI_Table;

    sL_Where := ' WHERE COMPCODE=' + sI_CompCode + ' AND CitemCode IN(' + sL_BackCitemCodeString + ') ';

    if nI_RealOrShouldDate = REAL_DATE_TYPE then
    begin
      sL_Where := sL_Where + ' AND (RealDate BETWEEN ' + TUstr.getOracleSQLDateTimeStr(StrToDateTime(sL_StartDate))
                  + ' AND ' + TUstr.getOracleSQLDateTimeStr(StrToDateTime(sL_EndDate)) + ') ';

    end
    else if nI_RealOrShouldDate = SHOULD_DATE_TYPE then
    begin
      sL_Where := sL_Where + ' AND (ShouldDate BETWEEN ' + TUstr.getOracleSQLDateTimeStr(StrToDateTime(sL_StartDate))
                  + ' AND ' + TUstr.getOracleSQLDateTimeStr(StrToDateTime(sL_EndDate)) + ') ';

    end;

//    sL_Where := sL_Where + ' AND CancelCode IS NULL AND CancelName IS NULL ';
    sL_Where := sL_Where + ' AND CancelFlag=0 ';

    sL_SQL := sL_SQL + sL_Where + ' ORDER BY CitemCode';

    with cdsSO033 do
    begin
      Close;
      CommandText := sL_SQL;
      Open;
    end;
end;

function TdtmMain2.calculateCompleteDetailData(sI_Table,sI_ComputeYM: String;
  nI_RealOrShouldDate: Integer): String;
var
    sL_SQL,sL_SBillNo,sL_SItem,sL_CompCode,sL_OrderNo : String;
    sL_MediaCode,sL_MediaName,sL_AcceptEn : String;
    sL_AcceptName,sL_IntroId,sL_IntroName : String;
    nL_CitemCode : Integer;
    L_Intf : IEmc_Separate1;
begin
    //才兵ン SO131 戈,斗тΩ
    if sI_Table = 'SO033' then
    begin
      sL_SQL := 'SELECT * FROM SO131 WHERE COMPCODE=' + frmMainMenu.sG_CompCode +
                ' AND COMPUTEYM=''' + sI_ComputeYM +
                ''' AND RealOrShouldDate=' + IntToStr(nI_RealOrShouldDate) +
                ' Order By BillNo,Item';


      //パ Server ㄓ Query SO131 Data
      //竒パ ADOQuery  CDS 戈矪瞶 ApplyUpdate  DB
      dtmMain2.connToServer;
      //if not dtmMain2.DCOMConnection1.Connected then
      //  dtmMain2.DCOMConnection1.Connected := true;

      L_Intf := IUnknown(dtmMain2.DCOMConnection1.AppServer) as IEmc_Separate1;

      L_Intf.getSO131Data(sL_SQL);


      with cdsSO131 do
      begin
        Close;
        Open;
      end;
    end;


    //盢璽兜タ兜,璝礚癸莱タ兜戈玥程
    cdsSO033.First;
    while not cdsSO033.Eof do
    begin

      sL_SBillNo := cdsSO033.FieldByName('SBILLNO').AsString;
      sL_SItem := cdsSO033.FieldByName('SITEM').AsString;
      sL_CompCode := cdsSO033.FieldByName('COMPCODE').AsString;
      nL_CitemCode := cdsSO033.FieldByName('CitemCode').AsInteger;

      //眔璽兜癸莱タ兜Μ禣兜ヘ

      //浪琩琌璶璸衡CitemCode
      //isNeedCitemCode(nL_CitemCode);

      //璝ゼタ盽癶禣挡睲繵笵,玥SBILLNOのSITEM琌null,礚猭癸莱タ兜
      if sL_SBillNo <> '' then
      begin
        if cdsSO131.Locate('COMPUTEYM;REALORSHOULDDATE;COMPCODE;BILLNO;ITEM', VarArrayOf([sI_ComputeYM,nI_RealOrShouldDate,StrToInt(sL_CompCode),sL_SBillNo,StrToInt(sL_SItem)]), []) then
        begin
          cdsSO131.Next;
          cdsSO131.Insert;


        end
        else
        begin
          cdsSO131.Append;
        end;
      end
      else
      begin
        cdsSO131.Append;
      end;

      cdsSO131.FieldByName('ComputeYM').AsString := sI_ComputeYM;
      cdsSO131.FieldByName('RealOrShouldDate').AsInteger := nI_RealOrShouldDate;
      cdsSO131.FieldByName('BillNo').AsString := sL_SBillNo;
      cdsSO131.FieldByName('Item').AsString := sL_SItem;
      cdsSO131.FieldByName('CustId').AsInteger := cdsSO033.FieldByName('CustId').AsInteger;
      cdsSO131.FieldByName('CitemCode').AsInteger := nL_CitemCode;
      cdsSO131.FieldByName('CitemName').AsString := cdsSO033.FieldByName('CitemName').AsString;
      cdsSO131.FieldByName('ShouldDate').AsDateTime := cdsSO033.FieldByName('ShouldDate').AsDateTime;
      cdsSO131.FieldByName('ShouldAmt').AsFloat := cdsSO033.FieldByName('ShouldAmt').AsFloat;
      cdsSO131.FieldByName('RealDate').AsDateTime := cdsSO033.FieldByName('RealDate').AsDateTime;
      cdsSO131.FieldByName('RealAmt').AsFloat := cdsSO033.FieldByName('RealAmt').AsFloat;
      cdsSO131.FieldByName('RealPeriod').AsInteger := cdsSO131.FieldByName('RealPeriod').AsInteger;
      cdsSO131.FieldByName('RealStartDate').AsDateTime := cdsSO033.FieldByName('RealStartDate').AsDateTime;
      cdsSO131.FieldByName('RealStopDate').AsDateTime := cdsSO033.FieldByName('RealStopDate').AsDateTime;
      cdsSO131.FieldByName('CompCode').AsInteger := cdsSO033.FieldByName('CompCode').AsInteger;


      sL_OrderNo := cdsSO033.FieldByName('OrderNo').AsString;
      cdsSO131.FieldByName('OrderNo').AsString := sL_OrderNo;

      cdsSO131.FieldByName('SBillNo').AsString := sL_SBillNo;
      cdsSO131.FieldByName('SItem').AsString := sL_SItem;


      //眔SO105 Data(眔碈砰ざ残闽戈癟)
      getSO105Data(sL_OrderNo,sL_MediaCode,sL_MediaName,sL_AcceptEn,
                   sL_AcceptName,sL_IntroId,sL_IntroName);

      cdsSO131.FieldByName('MediaCode').AsString := sL_MediaCode;
      cdsSO131.FieldByName('MediaName').AsString := sL_MediaName;
      cdsSO131.FieldByName('AcceptEn').AsString := sL_AcceptEn;
      cdsSO131.FieldByName('AcceptName').AsString := sL_AcceptName;
      cdsSO131.FieldByName('IntroId').AsString := sL_IntroId;
      cdsSO131.FieldByName('IntroName').AsString := sL_IntroName;


      cdsSO131.FieldByName('Notes').AsString := '癶禣';

      cdsSO131.Post;

      cdsSO033.Next;
    end;
end;

procedure TdtmMain2.addSeqNo;
var
    nL_SeqNo : Integer;
begin
    nL_SeqNo := 0;

    cdsSO131.First;
    while not cdsSO131.Eof do
    begin
      nL_SeqNo := nL_SeqNo + 1;

      cdsSO131.Edit;
      cdsSO131.FieldByName('SEQNO').AsInteger := nL_SeqNo;
      cdsSO131.Post;

      cdsSO131.Next;
    end;
end;

procedure TdtmMain2.getDetailData(sI_CompCode, sI_ComputeYM: String;
  nI_RealOrShouldDate: Integer);
var
    sL_Title,sL_Where,sL_SQL,sL_ComputeYM : String;
    L_Intf : IEmc_Separate1;
    aValue: String;
begin
//showmessage('A');
    sL_ComputeYM := Trim(dtmMain2.ReplaceStr(sI_ComputeYM,'/'));

    //тタ盽戈
    sL_Title := 'SELECT ComputeYM,RealOrShouldDate,SeqNo,BillNo,Item,CustId,' +
                'CitemCode,CitemName,ShouldDate,TO_CHAR(ShouldAmt) ShouldAmt,' +
                'RealDate,TO_CHAR(RealAmt) RealAmt,RealPeriod,RealStartDate,' +
                'RealStopDate,CompCode,OrderNo,SBillNo,SItem,MediaCode,' +
                'MediaName,AcceptEn,AcceptName,IntroId,IntroName,Notes,RefNo FROM SO131 ';

    sL_Where := ' WHERE COMPCODE=' + sI_CompCode +
                ' AND COMPUTEYM=''' + sL_ComputeYM +
                ''' AND RealOrShouldDate=' + IntToStr(nI_RealOrShouldDate) +
                ' AND RefNo=2 ' +
                ' Order by SeqNo';

    sL_SQL := sL_Title + sL_Where;

    with cdsSO131Excel do
    begin
      RemoteServer := DCOMConnection1;
      //ProviderName := 'dspSO131';
      ProviderName := 'dspSO131Excel';
      Close;
      CommandText := sL_SQL;
      Open;
    end;



    cdsSO131Excel.First;
    while not cdsSO131Excel.Eof do
    begin
      aValue := cdsSO131Excel.FieldByName( 'COMPUTEYM' ).AsString;
      if ( aValue <> EmptyStr ) then
      begin
        aValue := IntToStr( StrToInt( Copy( aValue, 1, 3 ) ) + 1911 - 1911 ) + Copy( aValue, 4, 2 );
        cdsSO131Excel.Edit;
        cdsSO131Excel.FieldByName( 'COMPUTEYM2' ).AsString := aValue;
        cdsSO131Excel.Post;
      end;
      {}
      aValue := cdsSO131Excel.FieldByName( 'REALDATE' ).AsString;
      if ( aValue <> EmptyStr ) then
      begin
        aValue := FormatDateTime( 'eee/mm/dd', cdsSO131Excel.FieldByName( 'REALDATE' ).AsDateTime );
        cdsSO131Excel.Edit;
        cdsSO131Excel.FieldByName( 'REALDATE2' ).AsString := aValue;
        cdsSO131Excel.Post;
      end;
      {}
      aValue := cdsSO131Excel.FieldByName( 'SHOULDDATE' ).AsString;
      if ( aValue <> EmptyStr ) then
      begin
        aValue := FormatDateTime( 'eee/mm/dd', cdsSO131Excel.FieldByName( 'SHOULDDATE' ).AsDateTime );
        cdsSO131Excel.Edit;
        cdsSO131Excel.FieldByName( 'SHOULDDATE2' ).AsString := aValue;
        cdsSO131Excel.Post;
      end;
      {}
      aValue := cdsSO131Excel.FieldByName( 'REALSTARTDATE' ).AsString;
      if ( aValue <> EmptyStr ) then
      begin
        aValue := FormatDateTime( 'eee/mm/dd', cdsSO131Excel.FieldByName( 'REALSTARTDATE' ).AsDateTime );
        cdsSO131Excel.Edit;
        cdsSO131Excel.FieldByName( 'REALSTARTDATE2' ).AsString := aValue;
        cdsSO131Excel.Post;
      end;
      {}
      aValue := cdsSO131Excel.FieldByName( 'REALSTOPDATE' ).AsString;
      if ( aValue <> EmptyStr ) then
      begin
        aValue := FormatDateTime( 'eee/mm/dd', cdsSO131Excel.FieldByName( 'REALSTOPDATE' ).AsDateTime );
        cdsSO131Excel.Edit;
        cdsSO131Excel.FieldByName( 'REALSTOPDATE2' ).AsString := aValue;
        cdsSO131Excel.Post;
      end;
      {}
      cdsSO131Excel.Next;
    end;
    cdsSO131Excel.First;


    //т眀戈
    sL_Title := '';
    sL_Where := '';
    sL_SQL := '';

    sL_Title := 'SELECT ComputeYM,RealOrShouldDate,SeqNo,BillNo,Item,CustId,' +
                'CitemCode,CitemName,ShouldDate,TO_CHAR(ShouldAmt) ShouldAmt,' +
                'RealDate,TO_CHAR(RealAmt) RealAmt,RealPeriod,RealStartDate,' +
                'RealStopDate,CompCode,OrderNo,SBillNo,SItem,MediaCode,' +
                'MediaName,AcceptEn,AcceptName,IntroId,IntroName,Notes,RefNo FROM SO131 ';

    sL_Where := ' WHERE COMPCODE=' + sI_CompCode +
                ' AND COMPUTEYM=''' + sL_ComputeYM +
                ''' AND RealOrShouldDate=' + IntToStr(nI_RealOrShouldDate) +
                ' AND RefNo=1 ' +
                ' Order by SeqNo';

    sL_SQL := sL_Title + sL_Where;

    with cdsOtherSO131Excel do
    begin
      RemoteServer := DCOMConnection1;
      //ProviderName := 'dspSO131';
      ProviderName := 'dspOtherSO131Excel';
      Close;
      CommandText := sL_SQL;
      Open;
    end;



    cdsOtherSO131Excel.First;
    while not cdsOtherSO131Excel.Eof do
    begin
      aValue := cdsOtherSO131Excel.FieldByName( 'COMPUTEYM' ).AsString;
      if ( aValue <> EmptyStr ) then
      begin
        aValue := IntToStr( StrToInt( Copy( aValue, 1, 3 ) ) + 1911 + 1911 ) + Copy( aValue, 4, 2 );
        cdsOtherSO131Excel.Edit;
        cdsOtherSO131Excel.FieldByName( 'COMPUTEYM2' ).AsString := aValue;
        cdsOtherSO131Excel.Post;
      end; 
      {}
      aValue := cdsOtherSO131Excel.FieldByName( 'REALDATE' ).AsString;
      if ( aValue <> EmptyStr ) then
      begin
        aValue := FormatDateTime( 'eee/mm/dd', cdsOtherSO131Excel.FieldByName( 'REALDATE' ).AsDateTime );
        cdsOtherSO131Excel.Edit;
        cdsOtherSO131Excel.FieldByName( 'REALDATE2' ).AsString := aValue;
        cdsOtherSO131Excel.Post;
      end;
      {}
      aValue := cdsOtherSO131Excel.FieldByName( 'SHOULDDATE' ).AsString;
      if ( aValue <> EmptyStr ) then
      begin
        aValue := FormatDateTime( 'eee/mm/dd', cdsOtherSO131Excel.FieldByName( 'SHOULDDATE' ).AsDateTime );
        cdsOtherSO131Excel.Edit;
        cdsOtherSO131Excel.FieldByName( 'SHOULDDATE2' ).AsString := aValue;
        cdsOtherSO131Excel.Post;
      end;
      {}
      aValue := cdsOtherSO131Excel.FieldByName( 'REALSTARTDATE' ).AsString;
      if ( aValue <> EmptyStr ) then
      begin
        aValue := FormatDateTime( 'eee/mm/dd', cdsOtherSO131Excel.FieldByName( 'REALSTARTDATE' ).AsDateTime );
        cdsOtherSO131Excel.Edit;
        cdsOtherSO131Excel.FieldByName( 'REALSTARTDATE2' ).AsString := aValue;
        cdsOtherSO131Excel.Post;
      end;
      {}
      aValue := cdsOtherSO131Excel.FieldByName( 'REALSTOPDATE' ).AsString;
      if ( aValue <> EmptyStr ) then
      begin
        aValue := FormatDateTime( 'eee/mm/dd', cdsOtherSO131Excel.FieldByName( 'REALSTOPDATE' ).AsDateTime );
        cdsOtherSO131Excel.Edit;
        cdsOtherSO131Excel.FieldByName( 'REALSTOPDATE2' ).AsString := aValue;
        cdsOtherSO131Excel.Post;
      end;
      {}
      cdsOtherSO131Excel.Next;
    end;
    cdsOtherSO131Excel.First;


{
    sL_SQL := 'SELECT * FROM SO131 WHERE COMPCODE=' + sI_CompCode +
              ' AND COMPUTEYM=''' + sL_ComputeYM +
              ''' AND RealOrShouldDate=' + IntToStr(nI_RealOrShouldDate) +
              ' Order by SeqNo';

    //パ Server ㄓ Query SO131 Data
    //竒パ ADOQuery  CDS 戈矪瞶 ApplyUpdate  DB
    dtmMain2.connToServer;
    //if not dtmMain2.DCOMConnection1.Connected then
    //  dtmMain2.DCOMConnection1.Connected := true;
    L_Intf := IUnknown(dtmMain2.DCOMConnection1.AppServer) as IEmc_Separate1;

    L_Intf.getSO131Data(sL_SQL);

    with cdsSO131 do
    begin
      Close;
      Open;
    end;
}



end;

function TdtmMain2.isNeedCitemCode(nI_CitemCode: Integer): Boolean;
var
    nL_Ndx : Integer;
begin

    nL_Ndx := frmSO8A50.G_ChargeCodeStrList.IndexOf(IntToStr(nI_CitemCode));

    if nL_Ndx <> -1 then
      Result := true
    else
      Result := false;


end;

procedure TdtmMain2.cdsSO131ExcelCalcFields(DataSet: TDataSet);
var
    sL_CalcShouldAmt,sL_CalcRealAmt,sL_CalcRefNo : String;
    nL_CalcShouldAmt : Integer;
begin
    sL_CalcShouldAmt := cdsSO131Excel.FieldByName('ShouldAmt').AsString;
    if sL_CalcShouldAmt <> '' then
      cdsSO131Excel.FieldByName('CalcShouldAmt').asInteger := StrToInt(sL_CalcShouldAmt);

    sL_CalcRealAmt := cdsSO131Excel.FieldByName('RealAmt').AsString;
    if sL_CalcShouldAmt <> '' then
      cdsSO131Excel.FieldByName('CalcRealAmt').asInteger := StrToInt(sL_CalcRealAmt);


    sL_CalcRefNo := cdsSO131Excel.FieldByName('RefNo').AsString;
    if sL_CalcRefNo = '1' then
      cdsSO131Excel.FieldByName('CalcRefNo').AsString := '眀戈'
    else if sL_CalcRefNo = '2' then
      cdsSO131Excel.FieldByName('CalcRefNo').AsString := 'タ盽戈';
end;

procedure TdtmMain2.cdsOtherSO131ExcelCalcFields(DataSet: TDataSet);
var
    sL_CalcShouldAmt,sL_CalcRealAmt,sL_CalcRefNo : String;
    nL_CalcShouldAmt : Integer;
begin
    sL_CalcShouldAmt := cdsOtherSO131Excel.FieldByName('ShouldAmt').AsString;
    if sL_CalcShouldAmt <> '' then
      cdsOtherSO131Excel.FieldByName('CalcShouldAmt').asInteger := StrToInt(sL_CalcShouldAmt);

    sL_CalcRealAmt := cdsOtherSO131Excel.FieldByName('RealAmt').AsString;
    if sL_CalcShouldAmt <> '' then
      cdsOtherSO131Excel.FieldByName('CalcRealAmt').asInteger := StrToInt(sL_CalcRealAmt);


    sL_CalcRefNo := cdsOtherSO131Excel.FieldByName('RefNo').AsString;
    if sL_CalcRefNo = '1' then
      cdsOtherSO131Excel.FieldByName('CalcRefNo').AsString := '眀戈'
    else if sL_CalcRefNo = '2' then
      cdsOtherSO131Excel.FieldByName('CalcRefNo').AsString := 'タ盽戈';

end;

procedure TdtmMain2.connToServer;
begin
    if not dtmMain2.DCOMConnection1.Connected then
    begin
      dtmMain2.DCOMConnection1.ComputerName := frmMainMenu.sG_Server_IP;
      dtmMain2.DCOMConnection1.Connected := true;
    end;
end;

procedure TdtmMain2.DataModuleCreate(Sender: TObject);
begin
    initCdsAttribute;

end;

procedure TdtmMain2.initCdsAttribute;
var
    ii : Integer;
begin
    if cdsAttribute.State = dsInactive then
      cdsAttribute.CreateDataSet;
    cdsAttribute.EmptyDataSet;
    for ii:=1 to 3 do
    begin
      with cdsAttribute do
      begin
        Append;
        if ii=1 then
        begin
          FieldByName('CODENO').AsString:='1';
          FieldByName('DESCRIPTION').AsString:='PPV Provider';
        end;
        if ii=2 then
        begin
          FieldByName('CODENO').AsString:='2';
          FieldByName('DESCRIPTION').AsString:='Pay Channel Provider';
        end;
        if ii=3 then
        begin
          FieldByName('CODENO').AsString:='3';
          FieldByName('DESCRIPTION').AsString:='PPV and Pay Channel Provider';
        end;
        Post;
      end;
    end;

end;

procedure TdtmMain2.initPercentXmlData;
begin
    if cdsPercentXmlData.State = dsInactive then
      cdsPercentXmlData.CreateDataSet;
    cdsPercentXmlData.EmptyDataSet;

end;

procedure TdtmMain2.cdsPercentXmlDataBeforeScroll(DataSet: TDataSet);
begin
{
    if (DataSet.RecordCount>0) and ((DataSet.FieldByName('sProviderName').AsString='') or
       (DataSet.FieldByName('nPercent').AsString='')) then
    begin
      Abort;
    end;
}    
end;

procedure TdtmMain2.cdsPercentXmlDataAfterPost(DataSet: TDataSet);
var
    sL_XmlDataStr : String;
begin
{
      sL_XmlDataStr := genXmlDataStr;
      cdsSo112A.Edit;
      cdsSo112A.FieldByName('PERCENTXMLDATA').AsString := sL_XmlDataStr;
      cdsSo112A.Post;
      saveToDB(cdsSo112A);
}
end;

function TdtmMain2.genXmlDataStr: String;
var
    nL_Percent, ii : Integer;
    sL_ProviderID, sL_ProviderName, sL_Percent : String;
    sL_Result : String;
    L_ChildNode, L_RootNode : IXmlNode;


begin
    sL_Result := '';
    nL_Percent := 0;
    with cdsPercentXmlData do
    begin
      DisableControls;
      First;
      while not Eof do
      begin
        sL_ProviderID := FieldByName('sProviderID').AsString;
        sL_ProviderName := FieldByName('sProviderName').AsString;
        //sL_Percent := FieldByName('nPercent').AsString;

        if sL_Result='' then
          sL_Result := sL_ProviderID + DATA_SEP_1 + sL_ProviderName + DATA_SEP_1 + sL_Percent
        else
          sL_Result := sL_Result + DATA_SEP_2 + sL_ProviderID + DATA_SEP_1 + sL_ProviderName + DATA_SEP_1 + sL_Percent;

        Next;
      end;
      EnableControls;
    end;






    result := sL_Result;
end;

function TdtmMain2.genXmlDataStr1: String;
var
    nL_Amount, ii : Integer;
    sL_ProviderID, sL_ProviderName, sL_Amount : String;
    sL_Result : String;
    L_ChildNode, L_RootNode : IXmlNode;
begin
    sL_Result := '';
    nL_Amount := 0;
    with cdsPercentXmlData do
    begin
      DisableControls;
      First;
      while not Eof do
      begin
        sL_ProviderID := FieldByName('sProviderID').AsString;
        sL_ProviderName := FieldByName('sProviderName').AsString;
        //sL_Percent := FieldByName('nAmount').AsString;

        if sL_Result='' then
          sL_Result := sL_ProviderID + DATA_SEP_1 + sL_ProviderName + DATA_SEP_1 + sL_Amount
        else
          sL_Result := sL_Result + DATA_SEP_2 + sL_ProviderID + DATA_SEP_1 + sL_ProviderName + DATA_SEP_1 + sL_Amount;

        Next;
      end;
      EnableControls;
    end;

    result := sL_Result;
end;

procedure TdtmMain2.transXmlToCds;
var
    sL_ProviderID,sL_ProviderName, sL_Percent : String;
    sL_PercentData, sL_TmpFile : String;
    L_StrList1, L_StrList2 : TStringList;
    ii : Integer;
    L_XmlNode, L_RootNode : IXmlNode;
begin
    cdsPercentXmlData.EmptyDataSet;

    sL_PercentData := cdsSo112A.FieldByName('PERCENTXMLDATA').AsString ;
    if sL_PercentData='' then
      Exit
    else
    begin
      L_StrList1 := TUStr.ParseStrings(sL_PercentData,DATA_SEP_2);


      for ii:=0 to L_StrList1.Count -1 do
      begin
        L_StrList2 := TUStr.ParseStrings(L_StrList1.Strings[ii],DATA_SEP_1);
        if L_StrList2.Count=3 then
        begin
          sL_ProviderID := L_StrList2.Strings[0];
          sL_ProviderName := L_StrList2.Strings[1];
          sL_Percent :=  L_StrList2.Strings[2];
          cdsPercentXmlData.Append;
          cdsPercentXmlData.FieldByName('sProviderID').AsString := sL_ProviderID;
          cdsPercentXmlData.FieldByName('sProviderName').AsString := sL_ProviderName;
          cdsPercentXmlData.FieldByName('nPercent').AsInteger := StrToIntDef(sL_Percent,0);
          cdsPercentXmlData.Post;
        end;
      end;
    end;
end;

procedure TdtmMain2.cdsSo112AAfterScroll(DataSet: TDataSet);
begin
    transXmlToCds;
end;

end.

