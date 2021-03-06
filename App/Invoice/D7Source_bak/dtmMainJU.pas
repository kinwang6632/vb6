unit dtmMainJU;

interface

uses
  SysUtils, Classes, DB, ADODB, Forms, DBClient, Provider ,Dialogs,
  StdCtrls, DBCtrls, Controls, Variants, cbUtilis, dtmMainU,
  msxmldom, XMLDoc, xmldom, XMLIntf;


type
  TNormalObj = class(TObject)
     s_Code : String;
     s_Desc : String;
  end;

  TChargeData = class(TObject)
     s_Code : String;
     s_Desc : String;
  end;

  TdtmMainJ = class(TDataModule)
    adoQryCommon: TADOQuery;
    adoQryCommon2: TADOQuery;
    dspCommon: TDataSetProvider;
    cdsCommon: TClientDataSet;
    cdsPrintInvoice: TClientDataSet;
    cdsPrintInvoiceCHECKNO: TWideStringField;
    cdsPrintInvoiceINVID: TWideStringField;
    cdsPrintInvoiceBUSINESSID: TWideStringField;
    cdsPrintInvoiceINVDATE: TDateTimeField;
    cdsPrintInvoiceCUSTID: TWideStringField;
    cdsPrintInvoiceCUSTSNAME: TWideStringField;
    cdsPrintInvoiceINVTITLE: TWideStringField;
    cdsPrintInvoiceZIPCODE: TWideStringField;
    cdsPrintInvoiceMAILADDR: TWideStringField;
    cdsPrintInvoiceINVFORMAT: TWideStringField;
    cdsPrintInvoiceTAXTYPE: TWideStringField;
    cdsPrintInvoiceSALEAMOUNT: TBCDField;
    cdsPrintInvoiceTAXAMOUNT: TBCDField;
    cdsPrintInvoiceINVAMOUNT: TBCDField;
    cdsPrintInvoiceDESCRIPTION: TWideStringField;
    cdsPrintInvoiceSTARTDATE: TDateTimeField;
    cdsPrintInvoiceENDDATE: TDateTimeField;
    cdsPrintInvoiceQUANTITY: TIntegerField;
    cdsPrintInvoiceUNITPRICE1: TFloatField;
    cdsPrintInvoiceTOTALAMOUNT: TBCDField;
    cdsPrintInvoiceDESCRIPTION2: TWideStringField;
    cdsPrintInvoiceSTARTDATE2: TDateTimeField;
    cdsPrintInvoiceENDDATE2: TDateTimeField;
    cdsPrintInvoiceQUANTITY2: TIntegerField;
    cdsPrintInvoiceUNITPRICE2: TFloatField;
    cdsPrintInvoiceTOTALAMOUNT2: TBCDField;
    cdsPrintInvoiceDESCRIPTION3: TWideStringField;
    cdsPrintInvoiceSTARTDATE3: TDateTimeField;
    cdsPrintInvoiceENDDATE3: TDateTimeField;
    cdsPrintInvoiceQUANTITY3: TIntegerField;
    cdsPrintInvoiceUNITPRICE3: TFloatField;
    cdsPrintInvoiceTOTALAMOUNT3: TBCDField;
    cdsPrintInvoiceDESCRIPTION4: TWideStringField;
    cdsPrintInvoiceSTARTDATE4: TDateTimeField;
    cdsPrintInvoiceENDDATE4: TDateTimeField;
    cdsPrintInvoiceQUANTITY4: TIntegerField;
    cdsPrintInvoiceUNITPRICE4: TFloatField;
    cdsPrintInvoiceTOTALAMOUNT4: TBCDField;
    cdsPrintInvoiceDESCRIPTION5: TWideStringField;
    cdsPrintInvoiceSTARTDATE5: TDateTimeField;
    cdsPrintInvoiceENDDATE5: TDateTimeField;
    cdsPrintInvoiceQUANTITY5: TIntegerField;
    cdsPrintInvoiceUNITPRICE5: TFloatField;
    cdsPrintInvoiceTOTALAMOUNT5: TBCDField;
    cdsPrintInvoiceMEMO1: TStringField;
    adoQryTrans2Txt: TADOQuery;
    ADOQuery1: TADOQuery;
    adoQryChargeMasterData: TADOQuery;
    dspChargeMasterData: TDataSetProvider;
    cdsChargeMasterData: TClientDataSet;
    adoQryChargeDetailData: TADOQuery;
    dspChargeDetailData: TDataSetProvider;
    cdsChargeDetailData: TClientDataSet;
    cdsChargeMasterExcel: TClientDataSet;
    cdsChargeDetailExcel: TClientDataSet;
    cdsChargeMasterExcelDESCRIPTION: TStringField;
    cdsChargeMasterExcelQUANTITY: TStringField;
    cdsChargeMasterExcelSALEAMOUNT: TStringField;
    cdsChargeMasterExcelTAXAMOUNT: TStringField;
    cdsChargeMasterExcelTOTALAMOUNT: TStringField;
    cdsChargeDetailExcelINVDATE: TStringField;
    cdsChargeDetailExcelINVID: TStringField;
    cdsChargeDetailExcelDESCRIPTION: TStringField;
    cdsChargeDetailExcelSALEAMOUNT: TStringField;
    cdsChargeDetailExcelTAXAMOUNT: TStringField;
    cdsChargeDetailExcelTOTALAMOUNT: TStringField;
    cdsChargeDetailExcelCUSTID: TStringField;
    cdsChargeDetailExcelCUSTSNAME: TStringField;
    cdsChargeDetailExcelCHARGEDATE: TStringField;
    cdsChargeDetailExcelMemo1: TStringField;
    cdsChargeDetailExcelMemo2: TStringField;
    adoQrySelectInv016: TADOQuery;
    adoQrySelectInv016SHOULDBEASSIGNED: TStringField;
    adoQrySelectInv016SEQ: TStringField;
    adoQrySelectInv016COMPID: TStringField;
    adoQrySelectInv016CUSTID: TStringField;
    adoQrySelectInv016TEL: TStringField;
    adoQrySelectInv016BUSINESSID: TStringField;
    adoQrySelectInv016TITLE: TStringField;
    adoQrySelectInv016ZIPCODE: TStringField;
    adoQrySelectInv016INVADDR: TStringField;
    adoQrySelectInv016MAILADDR: TStringField;
    adoQrySelectInv016BEASSIGNEDINVID: TStringField;
    adoQrySelectInv016CHARGEDATE: TDateTimeField;
    adoQrySelectInv016TAXTYPE: TStringField;
    adoQrySelectInv016TAXRATE: TBCDField;
    adoQrySelectInv016SALEAMOUNT: TBCDField;
    adoQrySelectInv016TAXAMOUNT: TBCDField;
    adoQrySelectInv016INVAMOUNT: TBCDField;
    adoQrySelectInv016ISVALID: TStringField;
    adoQrySelectInv016HOWTOCREATE: TStringField;
    adoQrySelectInv016UPTTIME: TDateTimeField;
    adoQrySelectInv016UPTEN: TStringField;
    adoQrySelectInv017: TADOQuery;
    adoQrySelectInv017SEQ: TStringField;
    adoQrySelectInv017BILLID: TStringField;
    adoQrySelectInv017BILLIDITEMNO: TIntegerField;
    adoQrySelectInv017TAXTYPE: TStringField;
    adoQrySelectInv017CHARGEDATE: TDateTimeField;
    adoQrySelectInv017ITEMID: TStringField;
    adoQrySelectInv017DESCRIPTION: TStringField;
    adoQrySelectInv017QUANTITY: TIntegerField;
    adoQrySelectInv017UNITPRICE: TBCDField;
    adoQrySelectInv017TAXRATE: TBCDField;
    adoQrySelectInv017TAXAMOUNT: TBCDField;
    adoQrySelectInv017TOTALAMOUNT: TBCDField;
    adoQrySelectInv017STARTDATE: TDateTimeField;
    adoQrySelectInv017ENDDATE: TDateTimeField;
    adoQrySelectInv017CHARGEEN: TStringField;
    InV099DataSetProvider: TDataSetProvider;
    Inv099DataSet: TClientDataSet;
    DataSource1: TDataSource;
    adoSourcePreFix: TADOQuery;
    cdsUsePrefix: TClientDataSet;
    cdsUsePrefixOrder: TStringField;
    cdsUsePrefixPrefix: TStringField;
    cdsUsePrefixStartNum: TStringField;
    cdsUsePrefixEndNum: TStringField;
    adoInv007ToInv024: TADOQuery;
    adoQryInv024: TADOQuery;
    cdsChargeDetailExcelBillID: TStringField;
    adoInv001Code: TADOQuery;
    adoInv002Code: TADOQuery;
    adoInv019Code: TADOQuery;
    adoInv006Code: TADOQuery;
    adoInv022Code: TADOQuery;
    adoInv025Code: TADOQuery;
    cdsPrefix: TClientDataSet;
    cdsPrefixNo: TIntegerField;
    cdsPrefixPrefix: TStringField;
    cdsPrefixLastInvDate: TStringField;
    cdsPrefixIdentifyId1: TStringField;
    cdsPrefixIdentifyId2: TStringField;
    cdsPrefixCompId: TStringField;
    cdsPrefixYearMonth: TStringField;
    cdsPrefixStartNum: TStringField;
    AdoQryInsertInv016: TADOQuery;
    dspQryInsertInv016: TDataSetProvider;
    cdsQryInsertInv016: TClientDataSet;
    AdoQryInsertInv017: TADOQuery;
    dspQryInsertInv017: TDataSetProvider;
    cdsQryInsertInv017: TClientDataSet;
    spAssignInvID: TADOStoredProc;
    adoQryInv031: TADOQuery;
    dspQInv031: TDataSetProvider;
    cdsQInv031: TClientDataSet;
    adoQryInv033: TADOQuery;
    dspInv033: TDataSetProvider;
    cdsInv033: TClientDataSet;
    adoInv099Code: TADOQuery;
    cdsPrefixMemo: TStringField;
    cdsCheckBusinessID: TClientDataSet;
    cdsCheckJumpInvID: TClientDataSet;
    cdsCheckBusinessIDMemo: TStringField;
    cdsCheckJumpInvIDMemo: TStringField;
    adoInv007AmtCheck: TADOQuery;
    adoInv008AmtCheck: TADOQuery;
    cdsCheckAmount: TClientDataSet;
    cdsCheckAmountInvDate: TDateField;
    cdsCheckAmountInvID: TStringField;
    cdsCheckAmountMasterAmt: TIntegerField;
    cdsCheckAmountDetailAmt: TIntegerField;
    cdsCheckAmountIsObsolete: TStringField;
    cdsCheckMisData: TClientDataSet;
    cdsCheckMisDataBillNo: TStringField;
    cdsCheckMisDataInvID: TStringField;
    cdsCheckMisDataMisTitle: TStringField;
    cdsCheckMisDataMisBusinessID: TStringField;
    cdsCheckMisDataInvTitle: TStringField;
    cdsCheckMisDataInvBusinessID: TStringField;
    cdsCheckMisDataInvCustID: TStringField;
    cdsCheckMisDataMisCustID: TStringField;
    cdsCheckMisDataMisCustName: TStringField;
    cdsCheckMisDataReason: TStringField;
    adoInv004Code: TADOQuery;
    cdsCreateTypeInfo: TClientDataSet;
    cdsCreateTypeInfoCreateType1: TStringField;
    cdsCreateTypeInfoCreateType4: TStringField;
    cdsObsoleteInvoiceData: TClientDataSet;
    cdsObsoleteInvoiceDataInvID: TStringField;
    cdsObsoleteInvoiceDataInvDate: TStringField;
    cdsObsoleteInvoiceDataInvAmount: TStringField;
    cdsObsoleteInvoiceDataObsoleteReason: TStringField;
    cdsObsoleteInvoiceDataCustID: TStringField;
    cdsObsoleteInvoiceDataHowToCreate: TStringField;
    cdsCreateTypeInfoIsObsolete: TStringField;
    cdsCreateTypeInfoInvDate: TStringField;
    adoQryInv018: TADOQuery;
    adoInv003Code: TADOQuery;
    cdsSelectInv016: TClientDataSet;
    cdsSelectInv016ShouldBeAssigned: TStringField;
    cdsSelectInv016Seq: TStringField;
    cdsSelectInv016CustId: TStringField;
    cdsSelectInv016ChargeDate: TStringField;
    cdsSelectInv016TaxType: TStringField;
    cdsSelectInv016TaxRate: TStringField;
    cdsSelectInv016SaleAmount: TStringField;
    cdsSelectInv016TaxAmount: TStringField;
    cdsSelectInv016InvAmount: TStringField;
    cdsSelectInv016IsValid: TStringField;
    cdsSelectInv016HowToCreate: TStringField;
    cdsCreateTypeInfoCreateType2: TStringField;
    cdsCreateTypeInfoCreateType3: TStringField;
    cdsCreateTypeInfoRecordTotal: TStringField;
    XMLDoc: TXMLDocument;
    cdsWitchOne: TClientDataSet;
    adoInv012Code: TADOQuery;
    adoQrySelectInv017ShouldBeAssigned: TStringField;
    cdsChargeDetailExcelINVTITLE: TStringField;
    cdsChargeDetailExcelBusinessId: TStringField;
    cdsTemp: TClientDataSet;
    adoInv028Code: TADOQuery;
    adoInv040Code: TADOQuery;
    adoInvD09: TADOConnection;
    adoComm: TADOQuery;
    cdsChargeDetailExcelUploadFlag: TStringField;
    cdsChargeDetailExcelINVOICEKIND: TStringField;
    cdsChargeDetailExcelGIVEUNITDESC: TStringField;
    cdsChargeDetailExcelISGIVE: TStringField;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    G_ChargeItemStrList : TStringList;
    sG_ExcelSavePath : String;
    function OpenSQL(const aSQL: String):String;overload;
    function getIsPrintTitle(sL_CompID:String):String ;
    procedure ExecuteSQL(const ASql: String);
    procedure Trans2Txt(aFilePathName: String; aHowToCreate: Integer;aInvIds: String);
    procedure AfterPrintOrTrans(const aFrom: Integer; aSourceDataSet: TClientDataSet;aUseAllComp : Boolean);
    procedure CheckInvoice3(sI_CompID,sI_UserID:String;var I_LogList:TStringList;sI_InvIds:string);
    function getDataSet(sI_SQL:String):TDataSet;
    function getPrefixIndex(sI_InvID: String;I_Prefix,I_StartNum,I_EndNum:TStringList): Integer;
    procedure getInvoice3(sI_CompID,sI_UserID:String;var I_LogList:TStringList);


    procedure getPrintInvoiceTempData(sI_SInvID,sI_EInvID,
             sI_SInvDate,sI_EInvDate, sI_BusinessType,sI_InvKind,sI_Invs,
             sI_Where,sI_GiftUnit : String;aUseAllComp : Boolean);

    procedure FilterPrintInvoiceTempData(AIsInvUse: Integer; AInvUseId: String);

    procedure createComboBoxItem(COB:TComboBox;sI_SQL,sI_CodeColName,sI_DescColName : String;bI_HaveDefaultItem : Boolean;var I_StrList : TStringList); //?N?????qDB???X?????iComboBox
    procedure getComboBoxCode(COB:TComboBox; I_StrList : TStringList;var sI_Code,sI_Desc : String);
    procedure AddObjectToComboBox(I_StrList:TStrings; I_DS:TDataSet; bI_InsertAllDataItem, bI_CombineCD:boolean);overload;
    function getChargeData( sI_CompID, sI_UserID, sI_SInvDate, sI_EInvDate, sI_SChargeDate,
       sI_EChargeDate, sI_SInvID, sI_EInvID, sI_ChargeItemCode, sI_ChargeItemName, sI_IsPrinted,
       sI_InvFormat, sI_Business, sI_TaxType, sI_ToCreate, sI_RptType, sI_ExcelSavePath: String;
       var bI_PrintMemo1, bI_PrintMemo2 : Boolean;
       aServiceType, aServiceTypeName: String;
       var aText1, aText2, aText3, aText4: String;
       var aInvCount: Integer; aInvKind: Integer;aUseGive: Integer;
       aRptSource : Integer): Boolean;

    procedure calcChargeTotalData(I_CDS,I_SourceCDS:TClientDataSet;sI_RptType : String);

    function getInv016Data(sI_QueryType,sI_CustIDOrBillIDCondition,sI_CustIDOrBillIDCondition2,sI_CompID,sI_StartChargeDate,sI_EndChargeDate,sI_CurMinSeq,sI_CurMaxSeq : String;
             var fI_SaleAmount,fI_TaxAmount,fI_InvAmount : Double) : Integer;

    procedure getInv017Data(ASeq: String);
    procedure updateInv016ShouldBeAssigned(sI_SEQ,sI_ShouldBeAssigned,sI_CompID, sI_StartChargeDate,
               sI_EndChargeDate : String; UpdateInv017:Boolean; aMinSeq, aMaxSeq: String);
    procedure updateInv016InvAmount(sI_UnitPrice,sI_TaxAmount,sI_InvAmount:Double; sI_SEQ,
         sI_CompID, sI_StartChargeDate,sI_EndChargeDate: String);
    procedure updateInv016StopFlag(sI_SEQ,sI_CompID,sI_StartChargeDate,sI_EndChargeDate : String);

    procedure updateInv017ShouldBeAssigned(sI_SEQ, sI_ShouldBeAssigned,
      sI_BILLID, sI_BILLIDITEMNO: String);

    procedure listLastInvDate(sI_CompID, sI_YearMonth:String);//?C?X?????w?????o??????(?e??)

    function getAllInv001Data : Integer;
    function getAllInv040Data : Integer;
    function getInv002Data(bI_Select: Integer;sI_CustID1Str,sI_CustID2Str,
                           sI_CustName,sI_Sno,sI_Tel,sI_CompID : String) : Boolean;
    function checkPK(sI_SQL : String) : Boolean;
    function checkCdsPK(I_CDS : TDataSet;sI_CodeColName,sI_DescColName,sI_Code,sI_Desc : String) : Boolean;
    function getInv019Data(sI_CompID,sI_CustID : String) : Integer;
    function getAllInv006Data  : Integer;
    function getAllInv012Data  : Integer;
    function getAllInv022Data  : Integer;
    function getAllInv003Data  : Integer;
    function getAllInv004Data  : Integer;
    function getInv099Data(aCompID, aYear, aYM: String): Integer;
    function getInv018Data(aCompId, aYear: String): Integer;
    function changePrefixYM(sI_YM : String) : String;

    //??????
    function getInvCreating(sL_CompID:String):String ;
    function assignInvoiceWithSF(nI_TotalUnit:Integer; fI_TotalSaleAmount, fI_TotalTaxAmount, fI_TotalInvAmount:double;
             sI_UptEn,sI_CompID,sI_HowToCreate,sI_InvDate,sI_InvYearMonth,sI_StartDate,sI_EndDate,sI_PrefixString,sI_MisDbOwner: String;
             nI_Order : Integer;bI_InvDateEqualChargeDate : Boolean; sI_ShowFaci,sI_StarCMTVmail : Integer;
             var sI_LogDateTime:String): Boolean;

    function getLastInvDate(sI_CompID:String ; sI_InvYearMonth:String ; sI_IdentifyID1:String ; sI_IdentifyID2:String): String ;
    function getCanUseInvCounts(sI_CompID,sI_InvYearMonth,sI_IdentifyID1,sI_IdentifyID2 : String) : Integer;
    procedure openSQLGetAmount(sI_SQL:String; var fL_SaleAmount, fL_TaxAmount, fL_InvAmount:double);
    function GetXMLAttribute(const aSQL, aAttributeName: String): String;
    procedure reactiveDataSet(I_DataSet:TDataSet);
    procedure ResetInv016;
    procedure ResetInv017;
    procedure ResetWitchOne;
    procedure InsertInv016(dI_CurSeq, sI_CompID, sI_CustID, sI_Tel,
      sI_BusinessID, sI_Title, sI_ZipCode, sI_InvAddr, sI_MailAddr, sI_BeAssignedInvID,
      sI_IsValid,sI_IsPreCreated, sI_UserName,
      sI_ChargeTitle, aInvUseId, aInvUseDesc, sI_InvType, sI_Email,
      sI_Contmobile, sI_GiveId, sI_GiveName: String);
    procedure DeleteInv016(const dI_CurSeq: String);
    procedure UpdateInv016(const dI_CurSeq: String; const dI_SaleAmount,
      dI_MasterTaxAmount: Double; const dI_ChargeDate, dI_TaxType, dI_TaxRate: String);
    procedure InsertInv017(
      dI_CurSeq,
      sI_BillID,sI_BillIDItemNo,sI_MasterTaxType,sI_DetailChargeDate,
      sI_ItemID,sI_Description,sI_Quantity,sI_UnitPrice,
      sI_MasterTaxRate,sI_DetailTaxAmount,sI_TotalAmount,sI_StartDate,
      sI_EndDate,sI_ChargeEn, sI_ServiceType, aLinkToMis, aAccountNo,
      aCombCitemCode, aCombCitemName: String);

    function DetailBelongWitchOne(const aCurSeq, aTaxType: String;
      const aTaxRate: Double): String;

    procedure InsertBelongWitchOne(const aCurSeq, aTaxType: String;
      const aTaxRate: Double);

    procedure getInv031StartAndStopInvID(sI_CompID:String;var sI_StartInvID,sI_StopInvID, sI_AllInvId : String);

    procedure getInv016MinAndMaxChargeDate(sI_CompID,sI_CurMinSeq,sI_CurMaxSeq : String;var sI_MinChargeDate,sI_MaxChargeDate : String);

    function CanDropOrRecoverInv(var aParam: TCommonParam): Boolean;
    function CanDropOrRecoverDis(var aParam: TCommonParam): Boolean;

    function getUnusualInvID(nI_ConditionType : Integer;sI_Where1,sI_Where2 : String; var aSQL: String) : Integer;
    function handleCheckBusinessID(sI_CompID,sI_SDate,sI_EDate : String) : Integer;
    function handleCheckJumpInvID(sI_CompID,sI_InvYM,sI_Prefix : String) : Integer;
    function handleChkAmountData(dI_SDate, dI_EDate:TDate; bI_IncludeObsolete:boolean;sI_CompID : String) : Integer;
    function handleChkMisData(sI_SelectItem,sI_CompID,sI_SInvDate,sI_EInvDate : String) : Integer;


    function getObsoleteInvoiceData(sI_CompID,sI_SInvID,sI_EInvID,sI_SInvDate,sI_EInvDate,sI_ObsoleteId : String) : Integer;

    function DropOrRecoverInv(var aParam: TCommonParam): Boolean;
    function DropOrRecoverDis(var aParam: TCommonParam): Boolean;
      
    function InvHasBeenLock(const aInvYearMonth: String): Boolean;
    
    procedure UpdateObsoleteMisData(const aInvId: String);
    procedure UpdateRecoverMisData(const aInvId, aInvDate, aPreInvoice: String);

    procedure HandleIsLockedData(sI_CompID,sI_IsLockedYear,sI_IsLockedMonth,sI_IsLocked : String);

    function afterMonth(sI_YM : String;nI_Counts : Integer) : String;
    function AuthorizeInv(const aInvIdSt, aInvIdEd, aInvUseId: String; const AIsInvUse: Integer;const aInvKind: Integer;const aInvs:String): String;
    function AuthorizeElectronInv(const aInvDateSt,aInvDateEd,
                        aInvIdSt,aInvIdEd : String;aAllComp : Boolean ) :String;
    function AuthorizeDiscountInv(const aDiscDateSt,aDiscDateEd,aDiscNoSt,aDiscNoEd : string;
                                  aAllComp :Boolean):String;

    procedure synchronizeCompData(sI_MisDbOwner,sI_CurDateTime,sI_UserName : String);
    procedure synchronizeChargeItem(sI_MisDbOwner,sI_CurDateTime,sI_UserName : String);
    function SynchronizeChargeItem2: Boolean;
    procedure synchronizePayType(sI_MisDbOwner,sI_CurDateTime,sI_UserName : String);
    procedure synchronizeServiceType(sI_MisDbOwner,sI_CurDateTime,sI_UserName : String);
    procedure synchronizeCustID(sI_MisDbOwner,sI_CurDateTime,sI_UserName,sI_CompID,sI_CustID,sI_MisDbCompCode : String);
    procedure SynchronizeCD095;
    procedure SynchronizeCD110;
    procedure copyMisCustInfoToInv002(sI_MisDbOwner, sI_CurDateTime,sI_UserName, sI_CompID,sI_CustID,sI_MisDbCompCode : String);
    function copyMisZipcodeFromSo014(sI_MisDbOwner,sI_MisDbCompCode,sI_MailAddrNo : String) : String;
    procedure copyMisCustInfoToInv019(sI_MisDbOwner, sI_CurDateTime,sI_UserName, sI_CompID,sI_CustID,sI_MisDbCompCode,sI_CustName,sI_CustSName,sI_MZipCode,sI_MailAddress : String);
    {}
    function CalcCreateInvCounts(const aChargeDateSt, aChargeDateEd,
      aHowToCreate: String): Integer;
    {}
    function IsMultiInv(const aInvId: String): Boolean;
    function IsMultiInvEx(const aMainInvId: String): Boolean;
    {}
    function InvDeatilLinkToMIS(const aInvId, aSeq: String): Boolean;
    function GetAccountNo(const aInvId, aSeq, aBillId, aBillidItemNo: String): String;
    function GetFacisNo(const aInvId, aSeq: String): String;
    function GetAllowanceNo: String;
    {}
    function GetIfPrintCheckValue: String;
    function GetFulDate(aType: Integer; aDate: String):String;
  end;

var
  dtmMainJ: TdtmMainJ;

implementation

uses
  frmInvA01U, frmInvA05_2U, Ustru, XLSFile,
  UdateTimeu, MaskUtils, dtmSOU;

{$R *.dfm}

{ TdtmMain }

procedure TdtmMainJ.DataModuleCreate(Sender: TObject);
begin
  G_ChargeItemStrList := TStringList.Create;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMainJ.DataModuleDestroy(Sender: TObject);
begin
  G_ChargeItemStrList.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMainJ.AddObjectToComboBox(I_StrList: TStrings; I_DS: TDataSet;
  bI_InsertAllDataItem, bI_CombineCD: boolean);
var
    L_Obj : TNormalObj;
    sL_CodeNo, sL_Desc : String;
begin
    //bI_CombineCD => ?O?_?n?N code and description ???X?b?@?_
    I_StrList.Clear;
    with I_DS do
    begin
      if state = dsInactive then
        active := true;
      first;
      if bI_InsertAllDataItem then
      begin
        L_Obj := TNormalObj.Create;
        sL_CodeNo := DEFAULT_ALL_DATA_ITEM_CODE_VALUE;
        sL_Desc := DEFAULT_ALL_DATA_ITEM_DESC_VALUE;
        L_Obj.s_Code := sL_CodeNo;
        L_Obj.s_Desc := sL_Desc;
        I_StrList.AddObject(sL_CodeNo+CODE_VALUE_SEP+sL_Desc,L_Obj);
      end;

      while not eof do
      begin
        sL_CodeNo := FieldByName('CodeNo').AsString;
        sL_Desc := FieldByName('Description').AsString;
        L_Obj := TNormalObj.Create;
        L_Obj.s_Code := sL_CodeNo;
        L_Obj.s_Desc := sL_Desc;
        if bI_CombineCD then
          I_StrList.AddObject(sL_CodeNo+CODE_VALUE_SEP+sL_Desc,L_Obj)
        else
          I_StrList.AddObject(sL_Desc,L_Obj);          
        Next;
      end;
    end;
end;

procedure TdtmMainJ.CheckInvoice3(sI_CompID, sI_UserID: String;
  var I_LogList: TStringList;sI_InvIds:string );
var
  sL_CompID,sL_SQLPrefix,sL_SQLInvice : String;
  L_Prefix,L_StartNum,L_EndNum : TStringList ;
  L_PrefixDataSet,L_InvoiceDataSet : TDataSet ;
  sL_InvID,sL_InvPrefix : String;
  nL_InvNO,nL_TempInvNO,nL_PrefixIndex,nL_TempPrefixIndex,K,nL_StartNum : Integer ;
  bL_RightFlag : Boolean;
  aInvKindSQL,aInvKind : String;
begin
  sL_CompID := sI_CompID ;
  L_Prefix := TStringList.Create ;
  L_StartNum := TStringList.Create ;
  L_EndNum := TStringList.Create ;
  sL_SQLPrefix :=
    ' SELECT PREFIX,STARTNUM,ENDNUM FROM INV099 '+
    '  WHERE COMPID = ' + STR_SEP + sL_CompID+ STR_SEP + ' AND USEFUL=' + STR_SEP + 'Y' + STR_SEP + ' ORDER BY PREFIX';

  L_PrefixDataSet := getDataSet(sL_SQLPrefix);
  if L_PrefixDataSet.IsEmpty then
    begin
      MessageDlg('?r?y?????s?b', mtWarning,[mbOK],0);
      exit;
    end
  else
  begin
    with L_PrefixDataSet do
    begin
      while not Eof do
      begin
        L_Prefix.Add(FieldByName('PREFIX').AsString);
        L_StartNum.Add(FieldByName('STARTNUM').AsString);
        L_EndNum.Add(FieldByName('ENDNUM').AsString);
        Next;
      end;
      Close;
    end;
  end;

  sL_SQLInvice := 'SELECT INVID FROM INV023 WHERE COMPID='''+sL_CompID +
    ''' AND OWNER='''+sI_UserID+''' ORDER BY INVID' ;

  L_InvoiceDataSet := getDataSet(sL_SQLInvice);
//showmessage('sL_SQLInvice =  ' + IntToStr(L_InvoiceDataSet.RecordCount));


// ???R?}?l
  sL_InvPrefix := '';
  nL_PrefixIndex := -1 ; //?s???e?o?????????@???r?y
  nL_TempInvNO := -1 ;
  nL_TempPrefixIndex := -1 ;
  nL_StartNum := -1;
  bL_RightFlag := true;



  with L_InvoiceDataSet do
  begin
    DisableControls;
    while not Eof do
    begin
      sL_InvID :=  FieldByName('INVID').AsString;  //AK10000001
      nL_PrefixIndex := getPrefixIndex(sL_InvID,L_Prefix,L_StartNum,L_EndNum); //???o???e?o???????????r?y
      sL_InvPrefix := copy(sL_InvID,1,2);  //???oAK
      nL_InvNO := StrToInt(copy(sL_InvID,3,8));//???o 100000001

      //?P?Oinv023 ???o?????X?O?_??????
      if nL_PrefixIndex <> nL_TempPrefixIndex then //?r?y?w?? ???o?????X?P???r?y?????l???X????
      begin

{
        //nL_StartNum := StrToInt(L_StartNum[nL_PrefixIndex]);

        nL_StartNum := nL_InvNO;
        bL_RightFlag := false;
        if nL_InvNO > nL_StartNum then
        begin
          bL_RightFlag := false;
          for J := 0 to (nL_InvNO - nL_StartNum)-1 do
            I_LogList.Add('?o???G'+sL_InvPrefix+adjString(8,IntToStr(nL_StartNum+J),true,true)+' ???L')

        end
        else if nL_InvNO = nL_StartNum then
        begin
          //???`
        end
        else if nL_InvNO < nL_StartNum then
        //?o?????X?p?????l?r?y?????X
        begin
          I_LogList.Add(' ?o?????X?p?????l?r?y?????X');
          //???????|?o???a?K
        end;
}

        //?C???????@?????X,?@?w???|??????????
        nL_TempInvNO := nL_InvNO;
        nL_TempPrefixIndex := nL_PrefixIndex
      end
      else //?r?y???? ???o?????X?P?e?@?????X????
      begin
        if nL_InvNO-nL_TempInvNO=1 then
        begin
          //???`
        end
        else if nL_InvNO-nL_TempInvNO=0 then
        begin
          I_LogList.Add('?o???G'+sL_InvID+' ????');
          bL_RightFlag := false;
        end
        else if (nL_TempInvNO <> -1) AND (nL_InvNO - nL_TempInvNO > 1) then
        begin

          //#5945 ???L?n?W?[?o?????? By Kin 2011/03/16
          //#5945 ?????W???A?p?G?O???J?o?????n?XLog,???D?u???]?t?b???J?o?? By Kin 2011/04/15
          for K := 1 to (nL_InvNO - nL_TempInvNO)-1 do
          begin
            {
            aInvKindSQL := Format('SELECT NVL(INVOICEKIND,0)         ' +
                            ' FROM INV007                            ' +
                            ' WHERE INVID = ''%s''                   ',
                            [sL_InvPrefix+adjString(8,IntToStr(nL_TempInvNO+K),true,true)]);
            if OpenSQL( aInvKindSQL ) = '0' then
              aInvKind := '?q?l?p?????o??'
            else
              aInvKind := '?q?l?o??';
            }
            //?p?G?S?????J?o???A?????XLog
            if sI_InvIds = EmptyStr then
            begin
              bL_RightFlag := false;
              I_LogList.Add('?o???G'  +
                    sL_InvPrefix+adjString(8,IntToStr(nL_TempInvNO+K),true,true)+
                    ' ???L' );
            end else
            begin
              //?M???O?_?????X???J?o???????? By Kin 2011/04/12
              if Pos(sL_InvPrefix+adjString(8,IntToStr(nL_TempInvNO+K),true,true),sI_InvIds) > 0 then
              begin
                bL_RightFlag := false;
                I_LogList.Add('?o???G'  +
                      sL_InvPrefix+adjString(8,IntToStr(nL_TempInvNO+K),true,true)+
                    ' ???L' );
              end;
            end;


          end;
        end;
        nL_TempInvNO := nL_InvNO;
      end;

      Next;
    end;
    EnableControls;
    Close;
  end;
  if bL_RightFlag then
    I_LogList.Add('?L???L???o??');

  L_Prefix.Free;
  L_StartNum.Free;
  L_EndNum.Free;

end;



procedure TdtmMainJ.createComboBoxItem(COB: TComboBox; sI_SQL,sI_CodeColName,sI_DescColName: String;bI_HaveDefaultItem : Boolean;var I_StrList : TStringList);
var
  i,nL_CodeOrder : integer ;
  sL_Code, sL_Desc : String;
  L_StrList : TStringList;
  L_Obj : TNormalObj;
begin
    COB.Items.Clear;
    nL_CodeOrder := 0;
    with adoQryCommon do
    begin
      Close;
      SQL.Clear ;
      SQL.Add(sI_SQL) ;
      Open;

      if bI_HaveDefaultItem then
      begin
        sL_Code := DEFAULT_ALL_DATA_ITEM_CODE_VALUE;
        sL_Desc := DEFAULT_ALL_DATA_ITEM_DESC_VALUE;
        L_Obj := TNormalObj.Create;
        L_Obj.s_Code := sL_Code;
        L_Obj.s_Desc := sL_Desc;
        I_StrList.AddObject(IntToStr(i), L_Obj);

        L_StrList := TStringList.Create;
        L_StrList.Add(sL_Code);
        COB.Items.AddObject(sL_Code + '-' + sL_Desc, L_StrList);
      end;

      for i:=0 to RecordCount-1 do
      begin
        sL_Code := FieldByName(sI_CodeColName).AsString;
        sL_Desc := FieldByName(sI_DescColName).AsString;

        L_Obj := TNormalObj.Create;
        L_Obj.s_Code := sL_Code;
        L_Obj.s_Desc := sL_Desc;

        if bI_HaveDefaultItem then    //???w?]??
          nL_CodeOrder := i+1
        else
          nL_CodeOrder := i;

        I_StrList.AddObject(IntToStr(nL_CodeOrder), L_Obj);

        L_StrList := TStringList.Create;
        L_StrList.Add(sL_Code);
        COB.Items.AddObject(sL_Code + '-' + sL_Desc, L_StrList);
        Next ;
      end;
      Close;
    end;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMainJ.ExecuteSQL(const ASql: String);
begin
  adoQryCommon.Close;
  adoQryCommon.SQL.Clear;
  adoQryCommon.SQL.Text := ASql;
  adoQryCommon.ExecSQL;
  adoQryCommon.Close;
end;

{ ---------------------------------------------------------------------------- }
//#5922 ?W?[aUseGive ?O???????????? By Kin 2011/02/10
//#5922 ?W?[aRptSource ???p????,?p?G?O0 Detail?n????INV008A By Kin 2011/02/10
function TdtmMainJ.getChargeData(sI_CompID, sI_UserID, sI_SInvDate,
  sI_EInvDate, sI_SChargeDate, sI_EChargeDate, sI_SInvID, sI_EInvID,
  sI_ChargeItemCode,sI_ChargeItemName, sI_IsPrinted, sI_InvFormat, sI_Business, sI_TaxType,
  sI_ToCreate,sI_RptType,sI_ExcelSavePath: String;
  var bI_PrintMemo1,bI_PrintMemo2: Boolean;
  aServiceType, aServiceTypeName: String;
  var aText1, aText2, aText3, aText4: String; var aInvCount:Integer; aInvKind: Integer;
  aUseGive:Integer;aRptSource : Integer): Boolean;
var
    sL_MasterSQL,sL_DetailSQL,sL_Where,sL_CompName,sL_ExcelTitle : String;
    sL_FileName,sL_CurrDateTime : String;
    sL_DetailTable : String;
    sL_ConditionInvDate,sL_ConditionChargeDate,sL_ConditionInvID,sL_ConditionChargeItem : String;
    sL_ConditionIsPrinted, sL_ConditionInvFormat, sL_ConditionBusiness, sL_ConditionTaxType: String;
    sL_ConditionIsUseGive: String;
    sL_ConditionToCreate,sL_SQL : String;
    aConditionServiceTypeName: String ;
    sL_NvlInv028 : String;
begin

  Result := False;
  sL_SQL := 'SELECT ITEMID FROM INV028 WHERE REFNO <> 1';
  with adoQryCommon do
  begin

    Close;
    SQL.Clear;
    SQL.Add(sL_SQL);
    Open;
    if adoQryCommon.RecordCount <=0 then
    begin
      InfoMsg( '?????]?w?o?????~??????????1??????!');
      Close;
      Exit;
    end;

    sL_NvlInv028 := FieldByName('ITEMID').AsString;
    Close;
  end;
  sL_DetailTable := 'INV008';
  if aRptSource = 1 then
  begin
    sL_DetailTable := 'INV008A';
  end;
  sL_CurrDateTime := TUstr.replaceStr( TUstr.replaceStr( TUstr.replaceStr(
    DateTimeToStr( Now ), '/', '' ),':','' ),' ','' );

    if sI_RptType = '3' then
      sL_ExcelTitle := '?????W??:???O???????p??-[???p??]'
    else if sI_RptType = '4' then
      sL_ExcelTitle := '?????W??:???O???????p??-[??????]';

    if ( sI_CompID = EmptyStr ) then
    begin
      sL_CompName := dtmMain.getCompNameEx;
      sL_ConditionInvDate := '???q?O:???????q';
      sL_ExcelTitle := sL_ExcelTitle + '@???q?W??:' + sL_CompName;
    end else
    begin
      sL_CompName := dtmMain.getCompName;
      sL_ConditionInvDate := '???q?O:?U?O???q';
      sL_ExcelTitle := sL_ExcelTitle + '@???q?W??:' + sL_CompName;
    end;  

    sL_Where := Format(
      ' WHERE A.INVID = B.INVID               ' +
      '   AND A.IDENTIFYID1 = ''%s''          ' +
      '   AND A.IDENTIFYID2 = ''%s''          ' +
      '   AND A.COMPID LIKE ''%s''            ' +
      '   AND A.ISOBSOLETE = ''N''            ',
      [IDENTIFYID1, IDENTIFYID2, Nvl( sI_CompID, '%' )] );

    //?o??????
    if ( sI_SInvDate <> EmptyStr ) then
    begin
      sL_Where := sL_Where + ' AND A.INVDATE BETWEEN ' +
                  TUstr.getOracleSQLDateStr(StrToDate(sI_SInvDate)) + ' AND ' +
                  TUstr.getOracleSQLDateStr(StrToDate(sI_EInvDate));
      sL_ExcelTitle := sL_ExcelTitle + '@?o??????:' + sI_SInvDate + '~' + sI_EInvDate;
      sL_ConditionInvDate := '?o??????:' + sI_SInvDate + '~' + sI_EInvDate;
    end else
      sL_ConditionInvDate := '?o??????:????';

    //?o?????? #5668 ?W?[?o?????? By Kin 2010/06/28
    case aInvKind of
      1:sL_Where := sL_Where + ' AND A.INVOICEKIND = 0 ';
      2:sL_Where := sL_Where + ' AND A.INVOICEKIND = 1 ';
    else ;
    end;

    //???O????
    if ( sI_SChargeDate <> EmptyStr ) then
    begin
      sL_Where := sL_Where + ' AND A.CHARGEDATE BETWEEN ' +
                  TUstr.getOracleSQLDateStr(StrToDate(sI_SChargeDate)) + ' AND ' +
                  TUstr.getOracleSQLDateStr(StrToDate(sI_EChargeDate));
      sL_ExcelTitle := sL_ExcelTitle + '@???O????:' + sI_SChargeDate + '~' + sI_EChargeDate;
      sL_ConditionChargeDate := '???O????:' + sI_SChargeDate + '~' + sI_EChargeDate;
    end else
      sL_ConditionChargeDate := '???O????:????';

    //?o?????X
    if ( sI_SInvID <> EmptyStr ) then
    begin
      sL_Where := sL_Where + Format( ' AND A.INVID BETWEEN ''%s'' AND ''%s'' ',
        [sI_SInvID, sI_EInvID] );
      sL_ExcelTitle := sL_ExcelTitle + '@?o?????X:' + sI_SInvID + '~' + sI_EInvID;
      sL_ConditionInvID := '?o?????X:' + sI_SInvID + '~' + sI_EInvID;
    end else
      sL_ConditionInvID := '?o?????X:????';


    //???O????
    if ( sI_ChargeItemCode <> EmptyStr ) then
      sL_Where := sL_Where + Format( ' AND B.ITEMID IN ( %s ) ', [sI_ChargeItemCode] )
    else
      sI_ChargeItemName := '????';
    sL_ExcelTitle := sL_ExcelTitle + '@???O????:' + sI_ChargeItemName;
    sL_ConditionChargeItem := '???O????:' + sI_ChargeItemName;

    //?A???O
    if ( aServiceType <> EmptyStr ) then
      sL_Where := sL_Where + Format( ' AND B.SERVICETYPE = ''%s'' ', [aServiceType] )
    else
      aServiceTypeName := '????';
    sL_ExcelTitle := sL_ExcelTitle + '@?A???O:' + aServiceTypeName;
    aConditionServiceTypeName := '?A???O:' + aServiceTypeName;

    //?O?_?i????(?????O?_?w?C?L)
    if sI_IsPrinted = 'Y' then
    begin
      sL_Where := sL_Where + ' AND A.CANMODIFY = ''Y'' ';
      sL_ExcelTitle := sL_ExcelTitle + '@?O?_?i????:?i????';
      sL_ConditionIsPrinted := '?O?_?i????:?i????';
    end else
    if sI_IsPrinted = 'N' then
    begin
      sL_Where := sL_Where + ' AND A.CANMODIFY=''N'' ';
      sL_ExcelTitle := sL_ExcelTitle + '@?O?_?i????:???i????';
      sL_ConditionIsPrinted := '?O?_?i????:???i????';
    end else
    begin
      sL_ExcelTitle := sL_ExcelTitle + '@?O?_?i????:???z?|';
      sL_ConditionIsPrinted := '?O?_?i????:???z?|';
    end;


    //?o??????
    if ( sI_InvFormat <> EmptyStr ) then
    begin
      sL_Where := sL_Where + Format( ' AND A.INVFORMAT = ''%s''', [sI_InvFormat] );
      if sI_InvFormat = '1' then  //?q?l??
      begin
        sL_ExcelTitle := sL_ExcelTitle + '@?o??????:?q?l??';
        sL_ConditionInvFormat := '?o??????:?q?l??';
      end else
      if sI_InvFormat = '2' then  //???u?G?p
      begin
        sL_ExcelTitle := sL_ExcelTitle + '@?o??????:???u?G?p';
        sL_ConditionInvFormat := '?o??????:???u?G?p';
      end else
      if sI_InvFormat = '3' then  //???u?T?p
      begin
        sL_ExcelTitle := sL_ExcelTitle + '@?o??????:???u?T?p';
        sL_ConditionInvFormat := '?o??????:???u?T?p';
      end;
    end else
    begin
      sL_ExcelTitle := sL_ExcelTitle + '@?o??????:????';
      sL_ConditionInvFormat := '?o??????:????';
    end;
      

    //???~?O
    if sI_Business = '1' then  //???~?H
    begin
      sL_Where := sL_Where + ' AND A.BUSINESSID IS NOT NULL ';
      sL_ExcelTitle := sL_ExcelTitle + '@???~?O:???~?H';
      sL_ConditionBusiness := '???~?O:???~?H';
    end else
    if sI_Business = '2' then  //?D???~?H
    begin
      sL_Where := sL_Where + ' AND ( A.BUSINESSID IS NULL OR A.BUSINESSID = '''') ';
      sL_ExcelTitle := sL_ExcelTitle + '@???~?O:?D???~?H';
      sL_ConditionBusiness := '???~?O:?D???~?H';
    end else
    begin
      sL_ExcelTitle := sL_ExcelTitle + '@???~?O:????';
      sL_ConditionBusiness := '???~?O:????';
    end;


    //???|?O
    if ( sI_TaxType <> EmptyStr ) then
    begin
      sL_Where := sL_Where + Format( ' AND A.TAXTYPE = ''%s''', [sI_TaxType] );
      if sI_TaxType = '1' then  //???|
      begin
        sL_ExcelTitle := sL_ExcelTitle + '@???|?O:???|';
        sL_ConditionTaxType := '???|?O:???|';
      end else
      if sI_TaxType = '2' then  //?s?|?v
      begin
        sL_ExcelTitle := sL_ExcelTitle + '@???|?O:?s?|?v';
        sL_ConditionTaxType := '???|?O:?s?|?v';
      end else
      if sI_TaxType = '3' then  //?K?|
      begin
        sL_ExcelTitle := sL_ExcelTitle + '@???|?O:?K?|';
        sL_ConditionTaxType := '???|?O:?K?|';
      end;
    end else
    begin
      sL_ExcelTitle := sL_ExcelTitle + '@???|?O:????';
      sL_ConditionTaxType := '???|?O:????';
    end;

    //?}???O
    if ( sI_ToCreate <> EmptyStr ) then
    begin
      sL_Where := sL_Where + Format( ' AND A.HOWTOCREATE = ''%s'' ', [sI_ToCreate] );
      if sI_ToCreate = '1' then  //????-?w?}
      begin
        sL_ExcelTitle := sL_ExcelTitle + '@?}???O:mis????-?w?}';
        sL_ConditionToCreate := '?}???O:mis????-?w?}';
      end else
      if sI_ToCreate = '2' then  //????-???}
      begin
        sL_ExcelTitle := sL_ExcelTitle + '@?}???O:mis????-???}';
        sL_ConditionToCreate := '?}???O:mis????-???}';
      end
      else if sI_ToCreate = '3' then  //?{???}??
      begin
        sL_ExcelTitle := sL_ExcelTitle + '@?}???O:?{???}??';
        sL_ConditionToCreate := '?}???O:?{???}??';
      end
      else if sI_ToCreate = '4' then  //?@???}??
      begin
        sL_ExcelTitle := sL_ExcelTitle + '@?}???O:?@???}??';
        sL_ConditionToCreate := '?}???O:?@???}??';
      end;
    end else
    begin
      sL_ExcelTitle := sL_ExcelTitle + '@?}???O:????';
      sL_ConditionToCreate := '?}???O:????';
    end;
    //#5922 ?W?[???????O By Kin 2011/02/10
    case aUseGive of
      0 :
      begin
        sL_ExcelTitle := sL_ExcelTitle + '@???????O:????';
        sL_ConditionIsUseGive := '???????O:????';
      end;
      1 :
      begin
        sL_ExcelTitle := sL_ExcelTitle + '@???????O:????';
        sL_ConditionIsUseGive :=  '???????O:????';
      end;
    else
      begin
        sL_ExcelTitle := sL_ExcelTitle + '@???????O:?D????';
        sL_ConditionIsUseGive :=  '???????O:?D????';
      end;
    end;


    { ??????????, ?????b???????Y?????? }

    aText1 := Format( '???q????:%s,%s,%s',
      [sL_CompName, sL_ConditionInvDate, sL_ConditionChargeDate] );

    aText2 := Format( '%s,%s,%s',
      [sL_ConditionInvID, aConditionServiceTypeName, sL_ConditionIsPrinted] );

    aText3 := Format( '%s,%s,%s,%s,%s',
        [sL_ConditionInvFormat, sL_ConditionBusiness,
        sL_ConditionTaxType, sL_ConditionToCreate, sL_ConditionIsUseGive] );

    aText4 := sL_ConditionChargeItem;
    {}

//******************************************************************************
    if (sI_RptType = '1') or (sI_RptType = '3') then   //???p??,???pExcel
    begin
      //#5922 ?W?[???????O,???H?n?h??INV028 By Kin 2011/02/10
      if aUseGive = 0 then
      begin
        sL_MasterSQL :=
          ' SELECT B.DESCRIPTION,                                           ' +
          '        A.TAXTYPE,                                               ' +
          '        SUM( B.QUANTITY ) QUANTITY,                              ' +
          '        SUM( B.SALEAMOUNT ) SALEAMOUNT,                          ' +
          '        SUM( B.TAXAMOUNT ) TAXAMOUNT,                            ' +
          '        SUM( B.TOTALAMOUNT ) TOTALAMOUNT                         ' +
          '   FROM INV007 A, ' + sL_DetailTable + ' B                       ' +
          sL_Where + ' GROUP BY B.DESCRIPTION, A.TAXTYPE ';
      end;
      if aUseGive > 0 then
      begin
          sL_MasterSQL :=
            ' SELECT Distinct B.DESCRIPTION,                                           ' +
            '        A.TAXTYPE,                                               ' +
            '        SUM( B.QUANTITY ) QUANTITY,                              ' +
            '        SUM( B.SALEAMOUNT ) SALEAMOUNT,                          ' +
            '        SUM( B.TAXAMOUNT ) TAXAMOUNT,                            ' +
            '        SUM( B.TOTALAMOUNT ) TOTALAMOUNT                         ' +
            '   FROM INV007 A, ' + sL_DetailTable +' B, INV028 C              ';

        if aUseGive = 1 then
        begin
          sL_Where := sL_Where + ' AND NVL(A.INVUSEID,0) = C.ITEMID AND C.REFNO = 1 And A.CompId = C.CompId ';
        end else
        begin
          sL_Where := sL_Where + ' AND NVL(A.INVUSEID,' + sL_NvlInv028 + ') = C.ITEMID AND C.REFNO <> 1 And A.CompId = C.CompId ';
        end;
        sL_Where := sL_Where + ' GROUP BY B.DESCRIPTION, A.TAXTYPE ';
        sL_MasterSQL := sL_MasterSQL + sL_Where;
      end;
        cdsChargeMasterData.Close;
        cdsChargeMasterData.CommandText := sL_MasterSQL;
        cdsChargeMasterData.Open;
        cdsChargeMasterData.FindField('DESCRIPTION').DisplayLabel := '???O????';
        cdsChargeMasterData.FindField('QUANTITY').DisplayLabel := '???q';
        cdsChargeMasterData.FindField('SALEAMOUNT').DisplayLabel := '?P???B';
        cdsChargeMasterData.FindField('TAXAMOUNT').DisplayLabel := '?|?B';
        cdsChargeMasterData.FindField('TOTALAMOUNT').DisplayLabel := '?o?????B';
        cdsChargeMasterData.FindField('DESCRIPTION').DisplayWidth := 30;
        cdsChargeMasterData.FindField('QUANTITY').DisplayWidth := 12;
        cdsChargeMasterData.FindField('SALEAMOUNT').DisplayWidth := 15;
        cdsChargeMasterData.FindField('TAXAMOUNT').DisplayWidth := 10;
        cdsChargeMasterData.FindField('TOTALAMOUNT').DisplayWidth := 15;
        cdsChargeMasterData.FindField('TaxType').Visible := false;

        if ( cdsChargeMasterData.RecordCount <= 0 ) then
        begin
          WarningMsg( '?S?????X??????!' );
          Exit;
        end;
        Result := True;
        { ???p?? }
        if (sI_RptType = '1') then
        begin

//          rptInvC01_1 := TrptInvC01_1.Create(Application);
//          rptInvC01_1.sG_ConditionInvDate := sL_ConditionInvDate;
//          rptInvC01_1.sG_ConditionChargeDate := sL_ConditionChargeDate;
//          rptInvC01_1.sG_ConditionInvID := sL_ConditionInvID;
//          rptInvC01_1.sG_ConditionChargeItem := sL_ConditionChargeItem;
//          rptInvC01_1.sG_ConditionIsPrinted := sL_ConditionIsPrinted;
//          rptInvC01_1.sG_ConditionInvFormat := sL_ConditionInvFormat;
//          rptInvC01_1.sG_ConditionBusiness := sL_ConditionBusiness;
//          rptInvC01_1.sG_ConditionTaxType := sL_ConditionTaxType;
//          rptInvC01_1.sG_ConditionToCreate := sL_ConditionToCreate;
//          rptInvC01_1.sG_ConditionServiceType := aConditionServiceTypeName;
//          rptInvC01_1.sG_CompName := sL_CompName;
//          rptInvC01_1.sG_UserID := sI_UserID;
//
//          //?^?_?????A
//          dtmMain.setDefaultCursor;
//          rptInvC01_1.Preview;
//          rptInvC01_1.Free;

        end else
        if (sI_RptType = '3') then   //???pExcel
        begin
          if not cdsChargeMasterExcel.Active then
            cdsChargeMasterExcel.CreateDataSet;
          cdsChargeMasterExcel.EmptyDataSet;
          cdsChargeMasterExcel.DisableControls;
          cdsChargeMasterData.First;
          while not cdsChargeMasterData.Eof do
          begin
            cdsChargeMasterExcel.Append;
            cdsChargeMasterExcel.FieldByName('DESCRIPTION').AsString :=
              cdsChargeMasterData.FindField('DESCRIPTION').AsString;
            cdsChargeMasterExcel.FieldByName('QUANTITY').AsString :=
              TUstr.CommaNumber( cdsChargeMasterData.FindField('QUANTITY').AsString );
            cdsChargeMasterExcel.FieldByName('SALEAMOUNT').AsString :=
              TUstr.CommaNumber( cdsChargeMasterData.FindField('SALEAMOUNT').AsString );
            cdsChargeMasterExcel.FieldByName('TAXAMOUNT').AsString :=
              TUstr.CommaNumber( cdsChargeMasterData.FindField('TAXAMOUNT').AsString );
            cdsChargeMasterExcel.FieldByName('TOTALAMOUNT').AsString :=
              TUstr.CommaNumber( cdsChargeMasterData.FindField('TOTALAMOUNT').AsString );
            cdsChargeMasterExcel.Post;
            cdsChargeMasterData.Next;
          end;
          cdsChargeMasterExcel.EnableControls;
          //?[?`
          dtmMainJ.calcChargeTotalData(cdsChargeMasterExcel,cdsChargeMasterData,sI_RptType);

          cdsChargeMasterExcel.FindField('DESCRIPTION').DisplayLabel := '???O????';
          cdsChargeMasterExcel.FindField('QUANTITY').DisplayLabel := '???q';
          cdsChargeMasterExcel.FindField('SALEAMOUNT').DisplayLabel := '?P???B';
          cdsChargeMasterExcel.FindField('TAXAMOUNT').DisplayLabel := '?|?B';
          cdsChargeMasterExcel.FindField('TOTALAMOUNT').DisplayLabel := '?o?????B';

          cdsChargeMasterExcel.FindField('DESCRIPTION').DisplayWidth := 30;
          cdsChargeMasterExcel.FindField('QUANTITY').DisplayWidth := 12;
          cdsChargeMasterExcel.FindField('SALEAMOUNT').DisplayWidth := 15;
          cdsChargeMasterExcel.FindField('TAXAMOUNT').DisplayWidth := 10;
          cdsChargeMasterExcel.FindField('TOTALAMOUNT').DisplayWidth := 15;

          sL_FileName := sI_ExcelSavePath + '\???O???p??_' + dtmMain.getCompName + '_' + sL_CurrDateTime + '.XLS';

          if cdsChargeMasterExcel.RecordCount > 0 then
          begin
            cdsChargeMasterExcel.DisableControls;
            DataSetToXLS(cdsChargeMasterExcel,sL_FileName,sL_ExcelTitle,'');
            cdsChargeMasterExcel.EnableControls;
            InfoMsg( '???XExcel????!' );
          end;
        end;

    end else
    if (sI_RptType = '2') or (sI_RptType = '4') then   //??????,??????Excel
    begin
      case aUseGive of
        0 :
          begin
            sL_Where := sL_Where + ' AND NVL(A.INVUSEID,' + sL_NvlInv028 + ' )= C.ITEMID ' +
                      ' AND A.COMPID = C.COMPID ';
          end;
        1 :
          begin
            sL_Where := sL_Where + ' AND A.INVUSEID = C.ITEMID AND C.REFNO = 1 AND A.COMPID = C.COMPID ';
          end;
      else
        begin
          sL_Where := sL_Where + ' AND NVL(A.INVUSEID,' + sL_NvlInv028 +') = C.ITEMID AND C.REFNO <> 1 AND A.COMPID = C.COMPID ';
        end;
      end;

      //?p???o??????
      sL_SQL :=
        ' SELECT COUNT( DISTINCT A.INVID ) INVIDCOUNT ' +
        '   FROM INV007 A,' + sL_DetailTable +' B,INV028 C ' + sL_Where;

      with adoQryCommon do
      begin
        Close;
        SQL.Clear;
        SQL.Add(sL_SQL);
        Open;
        aInvCount := FieldByName('INVIDCOUNT').AsInteger;
        Close;
      end;


      //???X?????o??????
      //#5922 ?W?[ ?W?????O,?o??????,???????O,???????? By Kin 2011/02/10
      //#6018 ?P?@?????O????+BillNo?|???P?X???u???@???A?{?b?[?JB.SEQ???K??(???????D?g???NDistinct????,?]????INV028?|?????D) By Kin 2011/05/17
      sL_DetailSQL :=
        ' SELECT DISTINCT A.INVDATE, A.INVID, B.DESCRIPTION,  B.SALEAMOUNT,                ' +
        '        B.TAXAMOUNT, B.TOTALAMOUNT, A.CUSTID, A.CUSTSNAME,                        ' +
        '        A.CHARGEDATE, A.MEMO1, A.MEMO2, A.TAXTYPE, B.BILLID,                      ' +
        '        A.INVTITLE, A.BUSINESSID,B.SEQ ,                                          ' +
        '        DECODE(NVL(A.UPLOADFLAG,0),0,''???W??'',''?w?W??'' ) UPLOADFLAG,          ' +
        '        DECODE(NVL(INVOICEKIND,0),0,''?q?l?p?????o??'',''?q?l?o??'') INVOICEKIND, ' +
        '        A.GIVEUNITDESC,DECODE(NVL(C.REFNO,0),1,''????'',''?D????'') ISGIVE        ' +
        '   FROM INV007 A, ' + sL_DetailTable + ' B, INV028 C                              ' +
        sL_Where + ' ORDER BY A.INVDATE,A.INVID  ';


      with cdsChargeDetailData do
      begin
        Close;
        CommandText := sL_DetailSQL;
        cdsChargeDetailData.DisableControls;
        Open;
        cdsChargeDetailData.EnableControls;

        FindField('INVDATE').DisplayLabel := '?o??????';
        FindField('INVID').DisplayLabel := '?o?????X';
        FindField('DESCRIPTION').DisplayLabel := '???O????';
        FindField('SALEAMOUNT').DisplayLabel := '?P???B';
        FindField('TAXAMOUNT').DisplayLabel := '?|?B';
        FindField('TOTALAMOUNT').DisplayLabel := '?o?????B';
        FindField('CUSTID').DisplayLabel := '?????N?X';
        FindField('CUSTSNAME').DisplayLabel := '????????';
        FindField('CHARGEDATE').DisplayLabel := '???O????';
        FindField('MEMO1').DisplayLabel := '?????@';
        FindField('MEMO2').DisplayLabel := '?????G';
        FindField('BillID').DisplayLabel := '???O????';
        FindField('INVTITLE').DisplayLabel := '?o?????Y';
        FindField('BUSINESSID').DisplayLabel := '???s';
        FindField('UPLOADFLAG').DisplayLabel :='?W?????O';
        FindField('INVOICEKIND').DisplayLabel := '?o???}??????';
        FindField('ISGIVE').DisplayLabel := '???????O';
        FindField('GIVEUNITDESC').DisplayLabel := '????????';

        FindField('INVDATE').DisplayWidth := 10;
        FindField('INVID').DisplayWidth := 12;
        FindField('DESCRIPTION').DisplayWidth := 10;
        FindField('CUSTID').DisplayWidth  := 10;
        FindField('CUSTSNAME').DisplayWidth := 20;
        FindField('SALEAMOUNT').DisplayWidth := 10;
        FindField('CHARGEDATE').DisplayWidth := 10;
        FindField('MEMO1').DisplayWidth := 60;
        FindField('MEMO2').DisplayWidth := 60;
        FindField('BillID').DisplayWidth := 15;
        FindField('INVTITLE').DisplayWidth := 20;
        FindField('BUSINESSID').DisplayWidth := 10;
        FindField('UPLOADFLAG').DisplayWidth := 10;
        FindField('INVOICEKIND').DisplayWidth :=10;
        FindField('ISGIVE').DisplayWidth := 10;
        FindField('GIVEUNITDESC').DisplayWidth := 50;

        FindField('TaxType').Visible := false;


        if not bI_PrintMemo1 then
          FindField('Memo1').Visible := false;

        if not bI_PrintMemo2 then
          FindField('Memo2').Visible := false;

        if ( cdsChargeDetailData.RecordCount <= 0 ) then
        begin
          WarningMsg( '?S?????X??????!' );
          Exit;
        end;

        Result := True;
        
          if (sI_RptType = '2') then   //??????
          begin
//            rptInvC01_2 := TrptInvC01_2.Create(Application);
//
//            rptInvC01_2.sG_ConditionInvDate := sL_ConditionInvDate;
//            rptInvC01_2.sG_ConditionChargeDate := sL_ConditionChargeDate;
//            rptInvC01_2.sG_ConditionInvID := sL_ConditionInvID;
//            rptInvC01_2.sG_ConditionChargeItem := sL_ConditionChargeItem;
//            rptInvC01_2.sG_ConditionIsPrinted := sL_ConditionIsPrinted;
//            rptInvC01_2.sG_ConditionInvFormat := sL_ConditionInvFormat;
//            rptInvC01_2.sG_ConditionBusiness := sL_ConditionBusiness;
//            rptInvC01_2.sG_ConditionTaxType := sL_ConditionTaxType;
//            rptInvC01_2.sG_ConditionToCreate := sL_ConditionToCreate;
//            rptInvC01_2.sG_ConditionServiceType := aConditionServiceTypeName;
//            rptInvC01_2.sG_CompName := sL_CompName;
//            rptInvC01_2.sG_UserID := sI_UserID;
//            rptInvC01_2.nG_InvIDCounts := aInvCount;
//
//            //?^?_?????A
//            dtmMain.setDefaultCursor;
//
//            rptInvC01_2.Preview;
//            rptInvC01_2.Free;

          end
          else if (sI_RptType = '4') then   //????Excel
          begin
            if not cdsChargeDetailExcel.Active then
              cdsChargeDetailExcel.CreateDataSet;
            cdsChargeDetailExcel.EmptyDataSet;

            cdsChargeDetailData.DisableControls;
            First;
            while not Eof do
            begin
              cdsChargeDetailExcel.Append;

              cdsChargeDetailExcel.FieldByName('INVDATE').AsString := FieldByName('INVDATE').AsString;
              cdsChargeDetailExcel.FieldByName('INVID').AsString := FieldByName('INVID').AsString;
              cdsChargeDetailExcel.FieldByName('DESCRIPTION').AsString := FieldByName('DESCRIPTION').AsString;
              cdsChargeDetailExcel.FieldByName('SALEAMOUNT').AsString := TUstr.CommaNumber(FieldByName('SALEAMOUNT').AsString);
              cdsChargeDetailExcel.FieldByName('TAXAMOUNT').AsString := TUstr.CommaNumber(FieldByName('TAXAMOUNT').AsString);
              cdsChargeDetailExcel.FieldByName('TOTALAMOUNT').AsString := TUstr.CommaNumber(FieldByName('TOTALAMOUNT').AsString);
              cdsChargeDetailExcel.FieldByName('CUSTID').AsString := FieldByName('CUSTID').AsString;
              cdsChargeDetailExcel.FieldByName('CUSTSNAME').AsString := FieldByName('CUSTSNAME').AsString;
              cdsChargeDetailExcel.FieldByName('CHARGEDATE').AsString := FieldByName('CHARGEDATE').AsString;
              cdsChargeDetailExcel.FieldByName('INVTITLE').AsString := FieldByName('INVTITLE').AsString;
              cdsChargeDetailExcel.FieldByName('BUSINESSID').AsString := FieldByName('BUSINESSID').AsString;
              cdsChargeDetailExcel.FieldByName('UPLOADFLAG').AsString := FieldByName('UPLOADFLAG').AsString;
              cdsChargeDetailExcel.FindField('INVOICEKIND').AsString := FieldByName('INVOICEKIND').AsString;
              cdsChargeDetailExcel.FindField('ISGIVE').AsString := FieldByName('ISGIVE').AsString;
              cdsChargeDetailExcel.FindField('GIVEUNITDESC').AsString := FieldByName('GIVEUNITDESC').AsString;
              if bI_PrintMemo1 then
                cdsChargeDetailExcel.FieldByName('Memo1').AsString := FieldByName('Memo1').AsString;

              if bI_PrintMemo2 then
                cdsChargeDetailExcel.FieldByName('Memo2').AsString := FieldByName('Memo2').AsString;

              cdsChargeDetailExcel.FieldByName('BillID').AsString := FieldByName('BillID').AsString;

              cdsChargeDetailExcel.Post;

              Next;
            end;
            cdsChargeDetailData.EnableControls;


            //?[?`
            dtmMainJ.calcChargeTotalData(cdsChargeDetailExcel,cdsChargeDetailData,sI_RptType);

            cdsChargeDetailExcel.FindField('INVDATE').DisplayLabel := '?o??????';
            cdsChargeDetailExcel.FindField('INVID').DisplayLabel := '?o?????X';
            cdsChargeDetailExcel.FindField('DESCRIPTION').DisplayLabel := '???O????';
            cdsChargeDetailExcel.FindField('SALEAMOUNT').DisplayLabel := '?P???B';
            cdsChargeDetailExcel.FindField('TAXAMOUNT').DisplayLabel := '?|?B';
            cdsChargeDetailExcel.FindField('TOTALAMOUNT').DisplayLabel := '?o?????B';
            cdsChargeDetailExcel.FindField('CUSTID').DisplayLabel := '?????N?X';
            cdsChargeDetailExcel.FindField('CUSTSNAME').DisplayLabel := '????????';
            cdsChargeDetailExcel.FindField('CHARGEDATE').DisplayLabel := '???O????';
            cdsChargeDetailExcel.FindField('INVTITLE').DisplayLabel := '?o?????Y';
            cdsChargeDetailExcel.FindField('BUSINESSID').DisplayLabel := '???s';
            cdsChargeDetailExcel.FindField('INVOICEKIND').DisplayLabel := '?o???}??????';
            cdsChargeDetailExcel.FindField('ISGIVE').DisplayLabel := '???????O';
            cdsChargeDetailExcel.FindField('GIVEUNITDESC').DisplayLabel := '????????';
            if bI_PrintMemo1 then
              cdsChargeDetailExcel.FindField('Memo1').DisplayLabel := '?????@'
            else
              cdsChargeDetailExcel.FindField('Memo1').DisplayLabel := ' ';

            if bI_PrintMemo2 then
              cdsChargeDetailExcel.FindField('Memo2').DisplayLabel := '?????G'
            else
              cdsChargeDetailExcel.FindField('Memo2').DisplayLabel := ' ';

            cdsChargeDetailExcel.FindField('BillID').DisplayLabel := '???O????';
            cdsChargeDetailExcel.FindField('UPLOADFLAG').DisplayLabel := '?W?????O';
            cdsChargeDetailExcel.FindField('INVDATE').DisplayWidth := 10;
            cdsChargeDetailExcel.FindField('INVID').DisplayWidth := 12;
            cdsChargeDetailExcel.FindField('DESCRIPTION').DisplayWidth := 10;
            cdsChargeDetailExcel.FindField('CUSTID').DisplayWidth  := 10;
            cdsChargeDetailExcel.FindField('CUSTSNAME').DisplayWidth := 20;
            cdsChargeDetailExcel.FindField('SALEAMOUNT').DisplayWidth := 10;
            cdsChargeDetailExcel.FindField('CHARGEDATE').DisplayWidth := 10;
            cdsChargeDetailExcel.FindField('Memo1').DisplayWidth := 60;
            cdsChargeDetailExcel.FindField('Memo2').DisplayWidth := 60;
            cdsChargeDetailExcel.FindField('BillID').DisplayWidth := 15;
            cdsChargeDetailExcel.FindField('INVTITLE').DisplayWidth := 50;
            cdsChargeDetailExcel.FindField('BUSINESSID').DisplayWidth := 10;
            cdsChargeDetailExcel.FindField('UPLOADFLAG').DisplayWidth :=10;
            if not bI_PrintMemo1 then
              cdsChargeDetailExcel.FindField('Memo1').Visible := false;

            if not bI_PrintMemo2 then
              cdsChargeDetailExcel.FindField('Memo2').Visible := false;


            sL_FileName := sI_ExcelSavePath + '\???O??????_' + dtmMain.getCompName + '_' + sL_CurrDateTime + '.XLS';;

            if cdsChargeDetailExcel.RecordCount > 0 then
            begin
              cdsChargeDetailExcel.DisableControls;
              DataSetToXLS(cdsChargeDetailExcel,sL_FileName,sL_ExcelTitle,'');
              cdsChargeDetailExcel.EnableControls;
            end;

            InfoMsg( '???XExcel?????C' );

            //?^?_?????A
            dtmMain.setDefaultCursor;
          end;
      end;
    end;
end;

procedure TdtmMainJ.getComboBoxCode(COB: TComboBox;I_StrList : TStringList;var sI_Code,sI_Desc : String);
var
    nL_Ndx : Integer;
begin
    nL_Ndx := I_StrList.IndexOf(IntToStr(COB.ItemIndex));
    if nL_Ndx<>-1 then
    begin
      sI_Code := (I_StrList.Objects[nL_Ndx] as TNormalObj).s_Code;
      sI_Desc := (I_StrList.Objects[nL_Ndx] as TNormalObj).s_Desc;
    end;
end;

function TdtmMainJ.getDataSet(sI_SQL: String): TDataSet;
begin
  with adoQryCommon do
  begin
    Close;
    SQL.Clear;
    SQL.Add(sI_SQL);
    Open;
    Result := adoQryCommon ;
  end;
end;

procedure TdtmMainJ.getInvoice3(sI_CompID, sI_UserID: String;
  var I_LogList: TStringList);
var
  sL_SQL :String ;
begin
  sL_SQL := 'SELECT INVID FROM INV023 WHERE HOWTOCREATE=3 AND OWNER=''' +
            sI_UserID + ''' AND COMPID=''' + sI_CompID + '''';
  I_LogList.Add('');

  with ADOQuery1 do
  begin
    Close;
    SQL.Clear;
    SQL.Add(sL_SQL);
    Open;
    I_LogList.Add('?{???}?????o?????X??'+IntToStr(RecordCount)+'?i?A?o?????X??');
    I_LogList.Add('');
    while not EOF do
    begin
      I_LogList.Add(FieldByName('INVID').AsString);
      Next ;
    end;
    Close;
  end;

end;

function TdtmMainJ.getIsPrintTitle(sL_CompID: String): String;
var
  aSQL : String ;
begin
  aSQL := Format( ' SELECT IFPRINTTITLE FROM INV003 WHERE COMPID = ''%s'' ',
   [sL_CompID] );
  Result := OpenSQL( aSQL );
end;

{ ---------------------------------------------------------------------------- }

function TdtmMainJ.getPrefixIndex(sI_InvID: String; I_Prefix, I_StartNum,
  I_EndNum: TStringList): Integer;
//?P?_?Y?o?????X?s?????@?r?y???????X??
var
  sL_InvPrefix : String;
  nL_InvNO,i,nL_PrefixCount,nL_PrefixIndex : Integer ;
begin
  nL_PrefixIndex := -1 ;
  nL_PrefixCount := I_Prefix.Count ;
  sL_InvPrefix := copy(sI_InvID,1,2);  //???oAK
  nL_InvNO := StrToInt(copy(sI_InvID,3,8));//???o 10000001
  for i:=0 to nL_PrefixCount-1 do
  begin
    if sL_InvPrefix = I_Prefix.Strings[i] then
      if (nL_InvNO >= StrToInt(I_StartNum.Strings[i])) and
      (nL_InvNO <= StrToInt(I_EndNum.Strings[i])) then
        nL_PrefixIndex := i;
  end;



//***********************************************************
{
  if sL_InvPrefix = I_Prefix.Strings[i] then
  begin
    if (nL_InvNO >= StrToInt(I_StartNum.Strings[i])) and
    (nL_InvNO <= StrToInt(I_EndNum.Strings[i])) then
    begin
      nL_PrefixIndex := i;
      nL_PrefixIndex := G_StrList.IndexOf(sI_CompCode);
      I_Prefix
    end;
  end;
}
//***********************************************************




  Result := nL_PrefixIndex ;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMainJ.OpenSQL(const aSQL: String): String;
begin
  adoQryCommon.Close;
  adoQryCommon.SQL.Text := aSQL;
  adoQryCommon.Open;
  Result := adoQryCommon.Fields[0].AsString;
  adoQryCommon.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMainJ.getPrintInvoiceTempData(sI_SInvID,
  sI_EInvID, sI_SInvDate, sI_EInvDate, sI_BusinessType,sI_InvKind,sI_Invs,
  sI_Where,sI_GiftUnit : String;aUseAllComp : Boolean);
var
  aSQL, aSubSql: String;
begin
  if aUseAllComp then
  begin
    aSQL := ' DELETE INV023';
  end else
  begin
    aSQL := Format( ' DELETE FROM INV023 WHERE COMPID = ''%S'' AND OWNER = ''%S'' ',
     [dtmMain.getCompID, dtmMain.getLoginUserName] );
  end;
  ExecuteSQL( aSQL );

  aSubSql := Format(
    ' SELECT UNIQUE MAININVID FROM INV007   ' +
    '  WHERE IDENTIFYID1 = ''%s''           ' +
    '    AND IDENTIFYID2 = ''%s''           ' +
    '    AND CANMODIFY IN ( %s )            ' +
    '    AND INVFORMAT = ''1''              ' +
    '    AND ISOBSOLETE = ''N''             ',
    [IDENTIFYID1, IDENTIFYID2,  GetIfPrintCheckValue] );
  if ( not aUseAllComp) then
  begin
    aSubSql := Format(aSubSql + ' AND COMPID = ''%s''',[dtmMain.getCompID]);
  end;
  if ( sI_SInvID <> EmptyStr ) then
  begin
    if ( sI_SInvID = sI_EInvID ) then
    begin
      aSubSql := aSubSql + Format( ' AND INVID = ''%s'' ', [sI_SInvID] );
    end else
    begin
      aSubSql := aSubSql + Format( ' AND INVID BETWEEN ''%s'' AND ''%s'' ',
        [sI_SInvID, sI_EInvID] );
    end;
  end;

  if ( sI_SInvDate <> EmptyStr ) then
  begin
    if ( sI_SInvDate = sI_EInvDate ) then
    begin
      aSubSql := aSubSql + Format( ' AND INVDATE = TO_DATE( ''%s'', ''YYYY/MM/DD'' ) ',
        [sI_SInvDate] );
    end else
    begin
      aSubSql := aSubSql + Format( ' AND INVDATE BETWEEN TO_DATE( ''%s'', ''YYYY/MM/DD'' ) AND TO_DATE( ''%s'', ''YYYY/MM/DD'' ) ',
        [sI_SInvDate, sI_EInvDate] );
    end;
  end;
  if (sI_GiftUnit <> EmptyStr ) then
  begin
    aSubSql := aSubSql + Format( 'AND GIVEUNITID IN( %s )  ',
        [sI_GiftUnit]);
  end;
  if sI_Where <> EmptyStr then
  begin
    aSubSql := aSubSql +  sI_Where;
  end;
  aSubSql := Format( '( %s ) B ', [aSubSql] );

  aSQL := Format(
    ' INSERT INTO INV023 (                                        ' +
    '    IDENTIFYID1, IDENTIFYID2, CHECKNO, INVID,                ' +
    '    CANMODIFY, INVDATE, CHARGEDATE, COMPID, CUSTID,          ' +
    '    CUSTSNAME, INVTITLE, ZIPCODE, INVADDR, MAILADDR,         ' +
    '    BUSINESSID, INVFORMAT, TAXTYPE, TAXRATE, SALEAMOUNT,     ' +
    '    TAXAMOUNT, INVAMOUNT, PRINTCOUNT, ISPAST, ISOBSOLETE,    ' +
    '    OBSOLETEID, OBSOLETEREASON, HOWTOCREATE, MEMO1, MEMO2,   ' +
    '    OWNER, UPTTIME, UPTEN, MAININVID, MAINSALEAMOUNT,        ' +
    '    MAINTAXAMOUNT, MAININVAMOUNT, INSTADDR, INVUSEID,        ' +
    '    INVUSEDESC,GIVEUNITID,GIVEUNITDESC )                     ' +
    ' SELECT                                                      ' +
    '    A.IDENTIFYID1, A.IDENTIFYID2, A.CHECKNO, A.INVID,        ' +
    '    A.CANMODIFY, A.INVDATE, A.CHARGEDATE, A.COMPID,          ' +
    '    A.CUSTID, A.CUSTSNAME, A.INVTITLE,                       ' +
    '    A.ZIPCODE, A.INVADDR, A.MAILADDR, A.BUSINESSID,          ' +
    '    A.INVFORMAT, A.TAXTYPE, A.TAXRATE, A.SALEAMOUNT,         ' +
    '    A.TAXAMOUNT, A.INVAMOUNT, A.PRINTCOUNT,                  ' +
    '    A.ISPAST, A.ISOBSOLETE, A.OBSOLETEID, A.OBSOLETEREASON,  ' +
    '    A.HOWTOCREATE, A.MEMO1, A.MEMO2, ''%s'' , A.UPTTIME,     ' +
    '    A.UPTEN, A.MAININVID, A.MAINSALEAMOUNT, A.MAINTAXAMOUNT, ' +
    '    A.MAININVAMOUNT, A.INSTADDR, A.INVUSEID, A.INVUSEDESC,   ' +
    '    A.GIVEUNITID,A.GIVEUNITDESC                              ' +
    '  FROM INV007 A, %s                                          ' +
    ' WHERE A.MAININVID = B.MAININVID                             ' +
    '   AND A.IDENTIFYID1 = ''%s''                                ' +
    '   AND A.IDENTIFYID2 = ''%s''                                ' +
    '   AND A.CANMODIFY IN ( %s )                                 ' +
    '   AND A.INVFORMAT = ''1''                                   ' +
    '   AND A.ISOBSOLETE = ''N''                                  ',
     [dtmMain.getLoginUserName, aSubSql, IDENTIFYID1, IDENTIFYID2,
       GetIfPrintCheckValue] );

    if ( not aUseAllComp ) then
    begin
      aSQL := Format(aSQL + ' AND A.COMPID = ''%s''',[dtmMain.getCompID]);
    end;


   //???~?H
   if sI_BusinessType = '1' then
     aSQL := ( aSQL + ' AND A.BUSINESSID IS NOT NULL ' )
   else if sI_BusinessType = '2' then  //?D???~?H
     aSQL := ( aSQL + ' AND A.BUSINESSID IS NULL ' );
   if sI_InvKind <> EmptyStr then
   begin
     aSQL := ( aSQL + ' AND A.INVOICEKIND = ' + sI_InvKind );
   end;
   if sI_Invs <> EmptyStr then
   begin
     aSQL := ( aSQL + ' AND A.INVID IN(' + sI_Invs + ')'  );
   end;
    ExecuteSQL( aSQL );

end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMainJ.FilterPrintInvoiceTempData(AIsInvUse: Integer; AInvUseId: String);

  { ----------------------------------------- }

  function CheckToSeeIfOldNullData(ACodes: String): Boolean;
  var
    aExtCode: String;
  begin
    repeat
      aExtCode := ExtractValue( ACodes );
      adoQryCommon2.Close;
      adoQryCommon2.SQL.Text := Format(
        ' SELECT COUNT(1) FROM INV028   ' +
        '  WHERE IDENTIFYID1 = ''%s''   ' +
        '    AND IDENTIFYID2 = ''%s''   ' +
        '    AND COMPID = ''%s''        ' +
        '    AND ITEMID = %s            ' +
        '    AND REFNO = 999            ',
        [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, aExtCode] );
      adoQryCommon2.Open;
      Result := ( adoQryCommon2.Fields[0].AsInteger > 0 );
      adoQryCommon2.Close;
    until ( ( Result ) or ( ACodes = EmptyStr ) );
  end;

  { ----------------------------------------- }

var
  aSql: String;
begin
  { ?Y???]?w?????o?? }
  { #5922 AIsInvUse?]??Integer 0,?N??????,1?N?????? 2.?N???D???? By Kin 2011/02/10}
  if ( AIsInvUse = 1 ) then   // ?L?o?u????????????
  begin
    adoQryCommon2.Close;
    adoQryCommon2.SQL.Text := Format(
      ' SELECT ITEMID FROM INV028    ' +
      '  WHERE IDENTIFYID1 = ''%S''  ' +
      '    AND IDENTIFYID2 = ''%S''  ' +
      '    AND COMPID = ''%S''       ' +
      '    AND NVL( REFNO, ''999'' ) <> ''1''  ',
      [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID] );
    adoQryCommon2.Open;
    adoQryCommon2.First;
    { ?v?@?R???D?????????? }
    while not adoQryCommon2.Eof do
    begin
      aSql := Format(
        ' DELETE FROM INV023            ' +
        '  WHERE IDENTIFYID1 = ''%S''   ' +
        '    AND IDENTIFYID2 = ''%S''   ' +
        '    AND COMPID = ''%S''        ' +
        '    AND OWNER = ''%S''         ' +
        '    AND INVUSEID = ''%S''      ',
        [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID,
         dtmMain.getLoginUserName,
         adoQryCommon2.FieldByName( 'ITEMID' ).AsString] );
      ExecuteSQL( aSql );
      adoQryCommon2.Next;
    end;
    adoQryCommon2.Close;
    { ?A?R?????o??????, InvUseId ?? Null }
    aSql := Format(
      ' DELETE FROM INV023            ' +
      '  WHERE IDENTIFYID1 = ''%S''   ' +
      '    AND IDENTIFYID2 = ''%S''   ' +
      '    AND COMPID = ''%S''        ' +
      '    AND OWNER = ''%S''         ' +
      '    AND INVUSEID IS NULL       ',
      [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID,
       dtmMain.getLoginUserName] );
    ExecuteSQL( aSql );
  end;
  { ?Y???]?w?D???o???o?? }
  if (  AIsInvUse = 2 ) then  //?N???????????R??
  begin
    { ???h?????????o?? }
    adoQryCommon2.Close;
    adoQryCommon2.SQL.Text := Format(
      ' SELECT ITEMID FROM INV028    ' +
      '  WHERE IDENTIFYID1 = ''%S''  ' +
      '    AND IDENTIFYID2 = ''%S''  ' +
      '    AND COMPID = ''%S''       ' +
      '    AND NVL( REFNO, ''999'' ) = ''1''  ',
      [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID] );
    adoQryCommon2.Open;
    adoQryCommon2.First;
    while not adoQryCommon2.Eof do
    begin
      aSql := Format(
        ' DELETE FROM INV023            ' +
        '  WHERE IDENTIFYID1 = ''%S''   ' +
        '    AND IDENTIFYID2 = ''%S''   ' +
        '    AND COMPID = ''%S''        ' +
        '    AND OWNER = ''%S''         ' +
        '    AND INVUSEID = ''%S''      ',
        [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID,
         dtmMain.getLoginUserName,
         adoQryCommon2.FieldByName( 'ITEMID' ).AsString] );
      ExecuteSQL( aSql );
      adoQryCommon2.Next;
    end;
    adoQryCommon2.Close;
    { ?Y???]?w?????o?????~ }
    if ( AInvUseId <> EmptyStr ) then
    begin
      aSql := Format(
        ' DELETE FROM INV023             ' +
        '  WHERE IDENTIFYID1 = ''%S''    ' +
        '    AND IDENTIFYID2 = ''%S''    ' +
        '    AND COMPID = ''%S''         ' +
        '    AND OWNER = ''%S''          ' +
        '    AND INVUSEID NOT IN ( %S )  ',
        [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID,
         dtmMain.getLoginUserName, AInvUseId ] );
      ExecuteSQL( aSql );
      { ?P?_?O?_?????w?????????o?????~, ?Y?O?S???????h?R?? }
      if not CheckToSeeIfOldNullData( AInvUseId ) then
      begin
        aSql := Format(
          ' DELETE FROM INV023             ' +
          '  WHERE IDENTIFYID1 = ''%S''    ' +
          '    AND IDENTIFYID2 = ''%S''    ' +
          '    AND COMPID = ''%S''         ' +
          '    AND OWNER = ''%S''          ' +
          '    AND INVUSEID IS NULL        ',
          [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID,
           dtmMain.getLoginUserName] );
        ExecuteSQL( aSql );   
      end;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMainJ.Trans2Txt(aFilePathName: String; aHowToCreate: Integer;aInvIds: String);
var
  sL_SQL,sL_InvDate,sL_InvID,sL_CheckNo,sL_BusinessID,sL_Description : String ;
  sL_Quantity,sL_UnitPrice,sL_SaleAmount,sL_TaxAmount,sL_TotalAmount : String ;
  sL_InvAmount,sL_TaxType,sL_InvTitle,sL_ZipCode,sL_MailAddr,sL_CustID : String;
  sL_TxtRecord,sL_PriorInvID,sL_InvCount,sL_ShowDate : String ;
  sL_LogFileName,sL_FirstInvID,sL_Date: String ;
  sL_MinInvDate ,sL_MaxInvDate,sL_CustSName, sL_Seq : String ;
  L_TxtList,L_LogList: TStringList ;
  nL_InvNumber,nL_Inv023Count : Integer;
  Year,Month,Day,Hour,Min,Sec,Msec : Word ;
  dL_SaleAmount,dL_TotalAmount,dL_TaxAmount: Double;
  aTransTotalAmount: Double;
  aServiceType: String;
  aMainInvId, aMainSaleAmount, aMainTaxAmount, aMainInvAmount: String;
  aAccountNo, aFacisNo, aCompTel: String;
  aClientDataSet: TClientDataSet;
  sL_Memo1,sL_Memo2: string;
  nPos :Integer;
  aSQLInv003 : String;
  aPrintZeroTax : Boolean;
  sL_GIVEUNITDESC: string;
begin

  {#4370 ?h?W?[Memo1?PMemo2 By Kin 2009/02/19}
  {#6001 ?W?[?????????W?? By Kin 2011/07/13}
  sL_SQL := Format(
    ' SELECT A.INVDATE,             ' +
    '        A.INVID,               ' +
    '        A.CHECKNO,             ' +
    '        A.BUSINESSID,          ' +
    '        B.SEQ,                 ' +
    '        B.DESCRIPTION,         ' +
    '        B.QUANTITY,            ' +
    '        B.UNITPRICE,           ' +
    '        B.TAXAMOUNT,           ' +
    '        B.SALEAMOUNT,          ' +
    '        B.TOTALAMOUNT,         ' +
    '        A.INVAMOUNT,           ' +
    '        A.TAXTYPE,             ' +
    '        A.INVTITLE,            ' +
    '        A.ZIPCODE,             ' +
    '        A.INVADDR,             ' +
    '        A.MAILADDR,            ' +
    '        A.CUSTID,              ' +
    '        A.CUSTSNAME,           ' +
    '        A.MAININVID,           ' +
    '        A.MAINSALEAMOUNT,      ' +
    '        A.MAINTAXAMOUNT,       ' +
    '        A.MAININVAMOUNT,       ' +
    '        B.STARTDATE,           ' +
    '        B.ENDDATE,             ' +
    '        B.SERVICETYPE,         ' +
    '        C.IFPRINTTITLE,        ' +
    '        A.MEMO1,               ' +
    '        A.MEMO2,               ' +
    '        A.GIVEUNITDESC         ' +
    '   FROM INV023 A, INV008 B,    ' +
    '        INV002 C               ' +
    '  WHERE A.IDENTIFYID1 = ''%s'' ' +
    '    AND A.IDENTIFYID2 = ''%s'' ' +
    '    AND A.COMPID = ''%s''      ' +
    '    AND A.OWNER = ''%s''       ' +
    '    AND A.INVID = B.INVID      ' +
    '    AND A.IDENTIFYID1 = C.IDENTIFYID1(+) ' +
    '    AND A.IDENTIFYID2 = C.IDENTIFYID2(+) ' +
    '    AND A.COMPID = C.COMPID(+) ' +
    '    AND A.CUSTID = C.CUSTID(+) ' +
    '    AND A.INVFORMAT = ''1''    ' +
    '    AND A.CANMODIFY IN ( %s )  ',
    [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID,
    dtmMain.getLoginUserName, GetIfPrintCheckValue] );

  if aHowToCreate = 1 then //?????F???t?o???{???}??
    sL_SQL := ( sL_SQL + ' AND A.HOWTOCREATE <> ''3'' ' );
  sL_SQL := ( sL_SQL + ' ORDER BY A.INVID, B.SEQ ' );
  aSQLInv003 := ' SELECT INVPARAM FROM INV003 WHERE COMPID = ' +
      QuotedStr( dtmMain.getCompID );
  aPrintZeroTax := dtmMainJ.GetXMLAttribute(aSQLInv003,'PrintZeroTax') = 'N';
  { ???s INV007 ?? }
  aClientDataSet := TClientDataSet.Create( nil );
  aClientDataSet.FieldDefs.Add( 'INVID', ftString, 10 );
  aClientDataSet.CreateDataSet;

  L_TxtList := TStringList.Create;

  sL_PriorInvID := '';
  aTransTotalAmount := 0;
  with adoQryTrans2Txt do
  begin
    SQL.Clear;
    SQL.Add(sL_SQL);
    Open ;
    frmInvA05_2.SetProgressMax(RecordCount);
    DisableControls;
    First;
    while not EOF do
    begin
      try

      sL_InvDate := getYearMonthDay7(FieldByName('INVDATE').AsDateTime,0);
      sL_InvID := FieldByName('INVID').AsString;
      if adoQryTrans2Txt.RecNo=1 then  //?u?O?????s???@?????o?????X
        sL_FirstInvID := sL_InvID;

      sL_CheckNo := adjString(2,FieldByName('CHECKNO').AsString,false,false);
      {}

      sL_BusinessID := adjString(8,FieldByName('BUSINESSID').AsString,false,false);


      sL_Description := adjString(40,FieldByName('DESCRIPTION').AsString,false,false);
      sL_Quantity := adjString(6,FieldByName('QUANTITY').AsString,true,true);

      {#5604 ?p?G BUSINESSID=???? ????7(????)?n????INV008.TotalAmount/Inv008.Quantity(?p???I??????)
         ????9(?|?B)?????J0?@By Kin 2010/03/25 }
      {#5764 ?P?_?D???~?H?O?_?C?L0?|?v,?p?G?_,??????unitprice By Kin 2010/09/10}
      if ( FieldByName('BUSINESSID').AsString = '' )  then
      begin
        if aPrintZeroTax then
        begin
          sL_UnitPrice := FloatToStr( StrToFloat(FieldByName( 'TOTALAMOUNT').AsString) /
                      StrtoFloat(FieldByName( 'QUANTITY' ).AsString ));
        end else
        begin
          sL_UnitPrice := adjString(10,FieldByName('UNITPRICE').AsString,true,true);
        end;

        nPos := AnsiPos('.',sL_UnitPrice );
        if nPos > 0 then
        begin
          sL_UnitPrice := FloatToStr( CbRoundTo(StrToFloat(sL_UnitPrice),2) );
        end;
        sL_UnitPrice := adjString( 10,sL_UnitPrice,True,True );
        //#5604 ?????O??10??0,?{?b????10?????? By Kin 2010/04/01
        //#5764 ?P?_?D???~?H?O?_?C?L0?|?v ,?p?G?_?h?] else???q By Kin 2010/09/10
        if aPrintZeroTax then
          sL_TaxAmount := adjString(10,'',True,False)
        else
          sL_TaxAmount := adjString(10,FieldByName('TAXAMOUNT').AsString,true,true);
          //sL_TaxAmount := adjString(10,'0',True, False );
      end else
      begin
        sL_UnitPrice := adjString(10,FieldByName('UNITPRICE').AsString,true,true);
        dL_TaxAmount := StrToFloat(FieldByName('TAXAMOUNT').AsString);
        sL_TaxAmount := adjString(10,FieldByName('TAXAMOUNT').AsString,true,true);
      end;
      dL_TotalAmount := StrToFloat(FieldByName('TOTALAMOUNT').AsString);
      sL_TotalAmount := adjString(10,FieldByName('TOTALAMOUNT').AsString,true,true);

      //dL_TaxAmount := StrToFloat(FieldByName('TAXAMOUNT').AsString);
      //sL_TaxAmount := adjString(10,FieldByName('TAXAMOUNT').AsString,true,true);

      {5604 ????8(?P???B)?n????INV008.TotalAmount By Kin 2010/03/25}
      {5764 ?P?_?D???~?H?????O?_?C?L0?|?v,?p?G?_?h?]else???q By Kin 2010/09/10 }
      if ( FieldByName('BUSINESSID').AsString = '' ) And ( aPrintZeroTax ) then
      begin
        sL_SaleAmount := adjString(10,FieldByName( 'TOTALAMOUNT' ).AsString,True,True);
      end else
      begin
        sL_SaleAmount := adjString(10,FieldByName( 'SALEAMOUNT' ).AsString,True,True);
//        dL_SaleAmount := dL_TotalAmount - dL_TaxAmount;
//        sL_SaleAmount := adjString(10, FloatToStr(dL_SaleAmount),true,true);
      end;
      {#4370 ?W?[Memo1?PMemo2 By Kin 2009/02/19}
      sL_Memo1 :=adjString( 60,FieldByName('Memo1').AsString,False,False );
      sL_Memo2 :=adjString( 60,FieldByName('Memo2').AsString,False,False );

      aTransTotalAmount := ( aTransTotalAmount + dL_TotalAmount );

      if sL_InvID <> sL_PriorInvID then //?N???O???@???o?????X
      begin
        sL_InvAmount := adjString(10,FieldByName('INVAMOUNT').AsString,true,true);

        sL_SQL := 'SELECT COUNT(INVID) FROM INV008 WHERE INVID = '''+sL_InvID +'''';
        sL_InvCount := adjString(2,openSQL(sL_SQL),true,true);
      end
      else
      begin//?Y?????????b???????H?W?A?h??????INVAMOUNT?????|?g?b???@???A???G???????N???s
        sL_InvAmount := adjString(10,'',true,true);
        sL_InvCount := adjString(2,'',true,true);
      end;

      sL_TaxType := FieldByName('TAXTYPE').AsString;
      {}



      {}
      if dtmMain.GetIfPrintTitle then
        { ???????@?s???O?_????, ?????J INVTITLE }
        sL_InvTitle := adjString(50,FieldByName('INVTITLE').AsString,false,false)
      else begin
        { ???@?s??????, ?~???J INVTITLE }
        sL_InvTitle := Lpad( EmptyStr, 50, #32 );
        if FieldByName('BUSINESSID').AsString <> EmptyStr then
          sL_InvTitle := adjString(50,FieldByName('INVTITLE').AsString,false,false)
        { ???@?s???S????, ???O INV002 ?? IfPrintTitle = Y , ???M?n?L }
        else if FieldByName( 'IFPRINTTITLE' ).AsString = 'Y' then
          sL_InvTitle := adjString(50,FieldByName('INVTITLE').AsString,false,false)
      end;
      {}
      sL_ZipCode := adjString(5,FieldByName('ZIPCODE').AsString,false,false);
      {}
      if ( dtmMain.GetExpAddrType = 1 ) then
        sL_MailAddr := adjString(60,FieldByName('MAILADDR').AsString,false,false)
      else
        sL_MailAddr := adjString(60,FieldByName('INVADDR').AsString,false,false);
      {}
      sL_CustID := '???s:'+adjString(8,FieldByName('CUSTID').AsString,false,false);

      sL_Date := FieldByName('STARTDATE').AsString;
      if sL_Date<>  '' then
        sL_ShowDate := '????:'+getYearMonthDay7(FieldByName('STARTDATE').AsDateTime,0)
      else
        sL_ShowDate :=adjString(12,'',False,False); { #4421 ?p?G?O?????n????20???? By Kin 2009/03/11 }
        //sL_ShowDate := adjString(18,'',false,false);

      sL_Date := FieldByName('ENDDATE').AsString;
      if sL_Date<> '' then
        sL_ShowDate := sL_ShowDate+'~'+getYearMonthDay7(FieldByName('ENDDATE').AsDateTime,0);

      {}
      { #4421 ?@?w?n????20 Bytes By Kin 2009/03/11 }
      sL_ShowDate :=sL_ShowDate + adjString(20-length(sL_ShowDate),'',False,False);
      
      {}
      sL_CustSName := adjString(50,FieldByName('CUSTSNAME').AsString,false,false);

      { ?A???O }
      aServiceType := Rpad( FieldByName( 'SERVICETYPE' ).AsString, 1, #32 );

      { ?D?o?????X, ?D?o???P???B,?D?o???|?B, ?D?o???`???B }
      aMainInvId := Rpad( FieldByName( 'MAININVID' ).AsString, 10, #32 );

      {#5604  ????22?n?? Inv007.MainSaleAmount ?B
        ????23?D?o???|?B?n????0 By Kin 2010/03/25 }
      {#5764 ?P?_?D???H?C?L0?|?v,?p?G?_?h?] else ???q By Kin 2010/09/10}
      if (FieldByName('BUSINESSID').AsString = '') And ( aPrintZeroTax ) then
      begin
        aMainSaleAmount := Lpad( FieldByName( 'MAININVAMOUNT' ).AsString, 10, '0' );
        //#5604 ????????10?????? By Kin 2010/04/01
        aMainTaxAmount := adjString( 10,'',True,False );
      end else
      begin
        aMainSaleAmount := Lpad( FieldByName( 'MAINSALEAMOUNT' ).AsString, 10, '0' );
        aMainTaxAmount := Lpad( FieldByName( 'MAINTAXAMOUNT' ).AsString, 10, '0' );
      end;
      aMainInvAmount := Lpad( FieldByName( 'MAININVAMOUNT' ).AsString, 10, '0' );

      { FacisNo ?]??????/CP????, ?n?? ?B?n }
      sL_Seq := FieldByName( 'SEQ' ).AsString;
      aFacisNo := dtmMainJ.GetFacisNo( sL_InvID, sL_Seq );
      if ( Pos( ',', aFacisNo ) > 0 ) then
        aFacisNo := Copy( aFacisNo, 1, Pos( ',', aFacisNo ) - 1 );
      aFacisNo :=  dtmMainU.FacisNoAddMask( aFacisNo );
      aFacisNo := Rpad( aFacisNo, 20, #32 );

      { ?b??/?d?? }
      aAccountNo := dtmMainJ.GetAccountNo( sL_InvID, sL_Seq, EmptyStr, EmptyStr );
      if ( Pos( ',', aAccountNo ) > 0 ) then
        aAccountNo := Copy( aAccountNo, 1, Pos( ',', aAccountNo ) - 1 );
      { ?e???h?[ "?b??" ?G?r }
      if Trim( aAccountNo ) <> EmptyStr then aAccountNo := ( '?b??:' + aAccountNo );
      aAccountNo := Rpad( aAccountNo, 16, #32 );

      { ???q?O?q?????X }
      aCompTel := Rpad( dtmMain.getCompTel, 20, #32 );

      {#6001 ?W?[?????????W?? By Kin 2011/07/13}
      sL_GIVEUNITDESC := adjString(20,FieldByName('GIVEUNITDESC').AsString,false,false);

      {}{
      sL_TxtRecord := sL_InvDate + sL_InvID + sL_CheckNo + sL_BusinessID +
        sL_Description + sL_Quantity + sL_UnitPrice +sL_SaleAmount +
        sL_TaxAmount +sL_TotalAmount + sL_InvAmount +sL_InvCount+
        sL_TaxType + sL_InvTitle + sL_ZipCode + sL_MailAddr + sL_CustID +
        sL_ShowDate + sL_CustSName + aServiceType + aMainInvId +
        aMainSaleAmount + aMainTaxAmount + aMainInvAmount +
        aFacisNo + aAccountNo + aCompTel;
       }
      //#4348 ????1?P18????????,?N????1?P18????28?P29???H????100?~???????? By Kin 2009/02/04
      //#4370 ?W?[????30?P31(Memo1,Memo2) By Kin 2009/02/19
      //#6001 ?W?[???e?????W?? By Kin 2011/07/13
      sL_TxtRecord := adjString(6,#32,True,False) + sL_InvID + sL_CheckNo + sL_BusinessID +
        sL_Description + sL_Quantity + sL_UnitPrice +sL_SaleAmount +
        sL_TaxAmount +sL_TotalAmount + sL_InvAmount +sL_InvCount+
        sL_TaxType + sL_InvTitle + sL_ZipCode + sL_MailAddr + sL_CustID +
        adjString(18,#32,True,False) + sL_CustSName + aServiceType + aMainInvId +
        aMainSaleAmount + aMainTaxAmount + aMainInvAmount +
        aFacisNo + aAccountNo + aCompTel + sL_InvDate +
        sL_ShowDate + sL_Memo1 + sL_Memo2 + sL_GIVEUNITDESC;
      L_TxtList.Add(sL_TxtRecord);

      aClientDataSet.Append;
      aClientDataSet.FieldByName( 'INVID' ).AsString := sL_InvID;
      aClientDataSet.Post;

      sL_PriorInvID := sL_InvID;

      frmInvA05_2.setProgress;

      Application.ProcessMessages;

      except
        on E: Exception do
        begin
          adoQryTrans2Txt.Close;
          L_TxtList.Free;
          aClientDataSet.Free;
          ErrorMsg( Format( '???X???????~, ?o?????X: %s, ???]:%s ', [sL_InvID, E.Message] ) );
          Exit;
        end;
      end;

      Next;

    end;
    EnableControls;

    nL_InvNumber := RecordCount;
    Close;
  end;

  if nL_InvNumber <= 0 then
  begin
    L_TxtList.Free;
    aClientDataSet.Free;
    InfoMsg( '?L?????i???X' );
    Exit;
  end;


  //???r????????
    L_txtList.SaveToFile( aFilePathName );
    L_TxtList.Free;


  //LOG????????
    DecodeDate(now,Year,Month,Day);
    DecodeTime(now,Hour,Min,Sec,Msec);
    sL_LogFileName := Format('%.4d%.2d%.2d-%.2d%.2d%.2d',[Year,Month,Day,Hour,Min,Sec]);


    L_LogList := TStringList.Create ;

    L_LogList.Add('?o???_???G'+sL_FirstInvID+'~'+sL_InvID);

    {}
    sL_SQL := Format(
      ' SELECT MIN(INVDATE) FROM INV023 WHERE IDENTIFYID1 = ''%s''          ' +
      '   AND IDENTIFYID2 = ''%s'' AND COMPID = ''%s'' AND OWNER = ''%s''   ' +
      '   AND INVFORMAT = ''1''  AND CANMODIFY IN ( %s )                    ',
      [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID,
       dtmMain.getLoginUserName, GetIfPrintCheckValue] );

    if ( aHowToCreate = 1 ) then //?????F???t?o???{???}??
    begin
      sL_SQL := ( sL_SQL + ' AND HOWTOCREATE <> ''3'' ' );
    end;

    sL_MinInvDate := OpenSQL(sL_SQL);
    {}

    {}
    sL_SQL := Format(
      ' SELECT MAX(INVDATE) FROM INV023 WHERE IDENTIFYID1 = ''%s''          ' +
      '   AND IDENTIFYID2 = ''%s'' AND COMPID = ''%s'' AND OWNER = ''%s''   ' +
      '   AND INVFORMAT = ''1''  AND CANMODIFY IN ( %s )                    ',
      [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID,
       dtmMain.getLoginUserName, GetIfPrintCheckValue] );

    if ( aHowToCreate = 1 ) then //?????F???t?o???{???}??
    begin
      sL_SQL := ( sL_SQL + ' AND HOWTOCREATE <> ''3'' ' );
    end;

    sL_MaxInvDate := OpenSQL(sL_SQL);
    {}


    L_LogList.Add('?????_?W?G'+sL_MinInvDate+'~'+sL_MaxInvDate);

    sL_SQL := Format(
      ' SELECT COUNT(1) COUNT FROM INV023 WHERE IDENTIFYID1 = ''%s''        ' +
      '   AND IDENTIFYID2 = ''%s'' AND COMPID = ''%s'' AND OWNER = ''%s''   ' +
      '   AND INVFORMAT = ''1''  AND CANMODIFY IN ( %s )                    ',
      [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID,
       dtmMain.getLoginUserName, GetIfPrintCheckValue] );

    if ( aHowToCreate = 1 ) then //?????F???t?o???{???}??
    begin
      sL_SQL := ( sL_SQL + ' AND HOWTOCREATE <> ''3'' ' );
    end;

    nL_Inv023Count := StrToInt( OpenSQL( sL_SQL ) );
    L_LogList.Add( '?o???i??:' + IntToStr( nL_Inv023Count ) );
    L_LogList.Add( '?o???`???B:' + FormatFloat( '#,##0,##', aTransTotalAmount ) );
    L_LogList.Add( '????????:' + DateTimeToStr( now ) );
    L_LogList.Add( '?H??:' + dtmMain.getLoginUserName );


    CheckInvoice3( dtmMain.getCompID, dtmMain.getLoginUserName, L_LogList,aInvIds );

    if aHowToCreate = 1 then //?????F???t?o???{???}??
      //???o???t?{???}?????o?????X
      getInvoice3( dtmMain.getCompID, dtmMain.getLoginUserName ,L_LogList );
      
    L_LogList.SaveToFile(ExtractFileDir( aFilePathName ) + '\' + sL_LogFileName + '.txt' );
    frmInvA05_2.ListBox1.Items.Assign( L_LogList );
    L_LogList.Free;

    AfterPrintOrTrans( 3, aClientDataSet,False );
    aClientDataSet.Free;

    InfoMsg( '???X????' );
    
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMainJ.AfterPrintOrTrans(const aFrom: Integer; aSourceDataSet: TClientDataSet
        ;aUseAllComp : Boolean);
var
  aSql, aPreInvId, aErrMsg: String;
begin
  { #5668 ?n??INVOICEKIND = 1 ???? 0 By Kin 2010/06/29  }
  { #5764 ?W?[?O???C?L?H?? By Kin 2010/09/13}
  aSql:=
    ' UPDATE INV007                                  ' +
    '    SET CANMODIFY = ''N'',                      ' +
    '        PRINTCOUNT = NVL( PRINTCOUNT, 0 ) + 1,  ' +
    '        INVOICEKIND = 0,                        ' +
    '        PRINTTIME = SYSDATE,                    ' +
    '        PRINTEN = ''%s''                        ' +
    '  WHERE IDENTIFYID1 = ''%s''                    ' +
    '    AND IDENTIFYID2 = ''%s''                    ' +
    '    AND INVID = ''%s''                          ';
  {
  if ( not aUseAllComp ) then
  begin
    aSql := aSql + ' AND COMPID = ''%s'' ';
  end;
  }
  if ( aFrom = 3 ) then
  begin
    aErrMsg := '???s?o?????A???w???X???o?????~, ?T??:%s?C';
  end else
  begin
    aErrMsg := '???s?o?????A???w?C?L???o?????~, ?T??:%s?C' ;
  end;  
  try
    { ?M???C?L?e???i??}
    if ( aFrom = 1 ) then
    begin
      adoQryCommon.Close;
      adoQryCommon.SQL.Text := Format(
        ' SELECT INVID FROM INV023 WHERE IDENTIFYID1 = ''%s''   ' +
        '   AND IDENTIFYID2 = ''%s''                            ' +
        '   AND OWNER = ''%s''                                  ',
        [IDENTIFYID1, IDENTIFYID2,  dtmMain.getLoginUserName] );
      if ( not aUseAllComp ) then
      begin
        adoQryCommon.SQL.Text := Format( adoQryCommon.SQL.Text +
              ' AND COMPID = ''%s'' ',[dtmMain.getCompID]);
      end;



      adoQryCommon.Open;
      adoQryCommon.First;
      while not adoQryCommon.Eof do
      begin
        { #5764 ?W?[?O???C?L?H?? By Kin 2010/09/13}
        adoQryCommon2.Close;
        adoQryCommon2.SQL.Text := Format( aSql, [
          dtmMain.getLoginUserName, IDENTIFYID1,
          IDENTIFYID2, adoQryCommon.FieldByName( 'INVID' ).AsString] );
        adoQryCommon2.ExecSQL;
        adoQryCommon.Next;
      end;
      adoQryCommon.Close;
    end else
    // ???X???r?? ???O???{???}???o???i??
    if ( aFrom in [2,3] ) then
    begin
      aSourceDataSet.First;
      aPreInvId := aSourceDataSet.FieldByName( 'INVID' ).AsString;
      while not aSourceDataSet.Eof do
      begin
        aSourceDataSet.Next;
        if ( aPreInvId <> aSourceDataSet.FieldByName( 'INVID' ).AsString ) or
           ( aSourceDataSet.Eof ) then
        begin
          { #5764?C ?W?[?O???C?L?H?? By Kin 2010/09/13}
          adoQryCommon2.Close;
          adoQryCommon2.SQL.Text := Format( aSql,
            [dtmMain.getLoginUserName, IDENTIFYID1,
            IDENTIFYID2, aPreInvId] );
          adoQryCommon2.ExecSQL;
          aPreInvId := aSourceDataSet.FieldByName( 'INVID' ).AsString;
        end;
      end;
    end;
  except
    on E: Exception do
    begin
      adoQryCommon.Close;
      adoQryCommon2.Close;
      ErrorMsg( Format( aErrMsg, [E.Message] ) );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMainJ.calcChargeTotalData(I_CDS,I_SourceCDS: TClientDataSet;
  sI_RptType: String);
var
  nL_ItemCounts : Integer;
  dL_SaleAmt,dL_TaxAmt,dL_SubTotalInvAmt : Double;
  dL_TotalSaleAmt,dL_TotalTaxAmt,dL_TotalInvAmt : Double;
  dL_FreeSaleAmt,dL_ShouldSaleAmt :Double;
  sL_TaxType : String;
begin
    nL_ItemCounts := 0;
    dL_SaleAmt := 0;
    dL_TaxAmt := 0;
    dL_SubTotalInvAmt := 0;
    dL_TotalSaleAmt := 0;
    dL_TotalTaxAmt := 0;
    dL_TotalInvAmt := 0;
    dL_FreeSaleAmt := 0;
    dL_ShouldSaleAmt := 0;


    nL_ItemCounts := I_SourceCDS.RecordCount; //???O?????`??
    I_SourceCDS.First;
    while not I_SourceCDS.Eof do
    begin
      dL_SaleAmt := I_SourceCDS.FieldByName('SALEAMOUNT').AsFloat;
      dL_TaxAmt := I_SourceCDS.FieldByName('TAXAMOUNT').AsFloat;
      dL_SubTotalInvAmt := I_SourceCDS.FieldByName('TOTALAMOUNT').AsFloat;
      sL_TaxType := I_SourceCDS.FieldByName('TaxType').AsString;

      dL_TotalSaleAmt := dL_TotalSaleAmt + dL_SaleAmt;         //?`?P???B
      dL_TotalTaxAmt := dL_TotalTaxAmt + dL_TaxAmt;            //?`?|?B
      dL_TotalInvAmt := dL_TotalInvAmt + dL_SubTotalInvAmt;    //?`?o?????B


      //1:???|(Default)2:?s?|?v 3:?K?|
      if sL_TaxType = '3' then
        dL_FreeSaleAmt := dL_FreeSaleAmt + dL_SaleAmt          //?K?|?P???B
      else if sL_TaxType = '1' then
        dL_ShouldSaleAmt := dL_ShouldSaleAmt + dL_SaleAmt;     //???|?P???B

      {
      if dL_TaxAmt = 0 then
        dL_FreeSaleAmt := dL_FreeSaleAmt + dL_SaleAmt          //?K?|?P???B
      else if dL_TotalSaleAmt <> 0 then
        dL_ShouldSaleAmt := dL_ShouldSaleAmt + dL_SaleAmt;     //???|?P???B
      }
      I_SourceCDS.Next;
    end;


    with I_CDS do
    begin
      if (sI_RptType = '1') OR (sI_RptType = '3') then
      begin
        Append;
        FieldByName('DESCRIPTION').AsString := '';

        Append;
        FieldByName('DESCRIPTION').AsString := '';

        Append;
        FieldByName('DESCRIPTION').AsString := '???O?????`??';
        FieldByName('QUANTITY').AsString := TUstr.CommaNumber(IntToStr(nL_ItemCounts));

        Append;
        FieldByName('DESCRIPTION').AsString := '?`?P???B';
        FieldByName('QUANTITY').AsString := TUstr.CommaNumber(FloatToStr(dL_TotalSaleAmt));

        Append;
        FieldByName('DESCRIPTION').AsString := '?`?|?B';
        FieldByName('QUANTITY').AsString := TUstr.CommaNumber(FloatToStr(dL_TotalTaxAmt));

        Append;
        FieldByName('DESCRIPTION').AsString := '?`?o?????B';
        FieldByName('QUANTITY').AsString := TUstr.CommaNumber(FloatToStr(dL_TotalInvAmt));

        Append;
        FieldByName('DESCRIPTION').AsString := '???|?P???B';
        FieldByName('QUANTITY').AsString := TUstr.CommaNumber(FloatToStr(dL_ShouldSaleAmt));

        Append;
        FieldByName('DESCRIPTION').AsString := '?K?|?P???B';
        FieldByName('QUANTITY').AsString := TUstr.CommaNumber(FloatToStr(dL_FreeSaleAmt));

        Post;
      end
      else if (sI_RptType = '2') OR (sI_RptType = '4') then
      begin
        Append;
        FieldByName('DESCRIPTION').AsString := '';

        Append;
        FieldByName('DESCRIPTION').AsString := '';

        Append;
        FieldByName('DESCRIPTION').AsString := '?`?P???B';
        FieldByName('SALEAMOUNT').AsString := TUstr.CommaNumber(FloatToStr(dL_TotalSaleAmt));

        Append;
        FieldByName('DESCRIPTION').AsString := '?`?|?B';
        FieldByName('SALEAMOUNT').AsString := TUstr.CommaNumber(FloatToStr(dL_TotalTaxAmt));

        Append;
        FieldByName('DESCRIPTION').AsString := '?`?o?????B';
        FieldByName('SALEAMOUNT').AsString := TUstr.CommaNumber(FloatToStr(dL_TotalInvAmt));

        Append;
        FieldByName('DESCRIPTION').AsString := '???|?P???B';
        FieldByName('SALEAMOUNT').AsString := TUstr.CommaNumber(FloatToStr(dL_ShouldSaleAmt));

        Append;
        FieldByName('DESCRIPTION').AsString := '?K?|?P???B';
        FieldByName('SALEAMOUNT').AsString := TUstr.CommaNumber(FloatToStr(dL_FreeSaleAmt));

        Post;
      end;
      {
      showmessage('???O?????`??   ' + IntToStr(nL_ItemCounts));
      showmessage('?`?P???B   ' + FloatToStr(dL_TotalSaleAmt));
      showmessage('?`?|?B   ' + FloatToStr(dL_TotalTaxAmt));
      showmessage('?`?o?????B   ' + FloatToStr(dL_TotalInvAmt));
      showmessage('???|?P???B   ' + FloatToStr(dL_ShouldSaleAmt));
      showmessage('?K?|?P???B   ' + FloatToStr(dL_FreeSaleAmt));
      }
  end;



end;

function TdtmMainJ.getInv016Data(sI_QueryType, sI_CustIDOrBillIDCondition, sI_CustIDOrBillIDCondition2,
  sI_CompID, sI_StartChargeDate, sI_EndChargeDate,sI_CurMinSeq,sI_CurMaxSeq: String;
  var fI_SaleAmount, fI_TaxAmount, fI_InvAmount: Double): Integer;
var
    sL_SQL,sL_Where : String;
    sL_HowToCreate : String;
begin
    sL_SQL := 'SELECT * FROM INV016 ';

    sL_Where := ' WHERE COMPID= '+''''+sI_CompID+''''+
      ' AND CHARGEDATE BETWEEN TO_DATE('+''''+sI_StartChargeDate+''','''+
      Date_Format+''') AND TO_DATE('+''''+sI_EndChargeDate+''','''+
      Date_Format+''') AND BEASSIGNEDINVID=''N'' AND ISVALID=''Y'''+
      ' AND STOPFLAG=0 ';

    if ( sI_CurMinSeq <> '' ) and ( sI_CurMaxSeq <> '' )  then
      sL_Where := sL_Where + ' AND SEQ BETWEEN ''' + sI_CurMinSeq + ''' AND ''' + sI_CurMaxSeq + '''';


    if sI_QueryType = QUERY_TYPE_ALL then

    else if sI_QueryType = QUERY_TYPE_CUSTID then
    begin
      sL_Where := sL_Where + ' AND CUSTID BETWEEN ''' + sI_CustIDOrBillIDCondition + ''' AND ''' + sI_CustIDOrBillIDCondition2 + '''';
    end
    else if sI_QueryType = QUERY_TYPE_CHARGEDATE then
    begin
      sL_Where := sL_Where + ' AND CHARGEDATE BETWEEN TO_DATE(''' + sI_CustIDOrBillIDCondition + ''', ''YYYY/MM/DD'') AND TO_DATE(''' + sI_CustIDOrBillIDCondition2 + ''', ''YYYY/MM/DD'')';
    end
    else if sI_QueryType = QUERY_TYPE_BILLID then
    begin
      sL_Where := sL_Where + ' AND SEQ BETWEEN ''' + sI_CustIDOrBillIDCondition + ''' AND ''' + sI_CustIDOrBillIDCondition2 + '''';
    end;
      //???? BIllID ?? INV017 ???X???? INV016 ?? SEQ
      {sL_SQL2 := 'SELECT SEQ FROM INV017 WHERE BILLID=''' + sI_CustIDOrBillIDCondition + '''';

      with cdsCommon do
      begin
        Close;
        CommandText := sL_SQL2;
        Open;

        L_SeqStrList := TStringList.Create;
        First;
        while not cdsCommon.Eof do
        begin
          sL_Seq := FieldByName('SEQ').AsString;

          nL_Ndx := L_SeqStrList.IndexOf(sL_Seq);
          if nL_Ndx = -1 then
            L_SeqStrList.Add(sL_Seq);

          Next;
        end;

        sL_SeqString := '';
        for ii:=0 to L_SeqStrList.Count-1 do
        begin
          sL_Seq := L_SeqStrList.Strings[ii];

          if sL_SeqString = '' then
            sL_SeqString := '''' + sL_Seq + ''''
          else
            sL_SeqString := sL_SeqString + ',''' + sL_Seq + '''';
        end;

        //?A?? SEQ ?? INV016 ??????????????
        sL_Where := sL_Where + ' AND SEQ IN(' + sL_SeqString + ')';
      end;
    end;

    if (sI_QueryType = QUERY_TYPE_BILLID) AND (sL_SeqString = '') then
      Result := 0
    else}

    //begin
      sL_SQL := sL_SQL + sL_Where + ' ORDER BY SEQ';


      with adoQrySelectInv016 do
      begin
        Close;
        SQL.Clear;
        SQL.Add(sL_SQL);
        Open;


        if VarIsNull( cdsSelectInv016.Data ) then cdsSelectInv016.CreateDataSet;
        cdsSelectInv016.EmptyDataSet;

        Result := RecordCount;
        fI_SaleAmount := 0;
        fI_TaxAmount := 0;
        fI_InvAmount := 0;

        if RecordCount <> 0 then
        begin
          First;
          DisableControls;
          cdsSelectInv016.DisableControls;
          while not Eof do
          begin
            cdsSelectInv016.Append;

            cdsSelectInv016.FieldByName('ShouldBeAssigned').AsString := FieldByName('ShouldBeAssigned').AsString;
            cdsSelectInv016.FieldByName('Seq').AsString := FieldByName('Seq').AsString;
            cdsSelectInv016.FieldByName('CustId').AsString := FieldByName('CustId').AsString;
            cdsSelectInv016.FieldByName('ChargeDate').AsString := FieldByName('ChargeDate').AsString;
            cdsSelectInv016.FieldByName('TaxType').AsString := FieldByName('TaxType').AsString;

            cdsSelectInv016.FieldByName('TaxRate').AsString := FieldByName('TaxRate').AsString;
            cdsSelectInv016.FieldByName('SaleAmount').AsString := FieldByName('SaleAmount').AsString;
            cdsSelectInv016.FieldByName('TaxAmount').AsString := FieldByName('TaxAmount').AsString;
            cdsSelectInv016.FieldByName('InvAmount').AsString := FieldByName('InvAmount').AsString;
            cdsSelectInv016.FieldByName('IsValid').AsString := FieldByName('IsValid').AsString;


            //1-mis????, ?w?}
            //2-mis????, ???}
            //3-invoice create, ?w?}
            //4-invoice create, ???}
            sL_HowToCreate := FieldByName('HowToCreate').AsString;

            if sL_HowToCreate = '1' then
              sL_HowToCreate := '?w?}'
            else if sL_HowToCreate = '2' then
              sL_HowToCreate := '???}'
            else if sL_HowToCreate = '3' then
              sL_HowToCreate := '?{???}??'
            else if sL_HowToCreate = '4' then
              sL_HowToCreate := '?@???}??';


            cdsSelectInv016.FieldByName('HowToCreate').AsString := sL_HowToCreate;

            cdsSelectInv016.Post;

            Next;
          end;
          cdsSelectInv016.EnableControls;
          EnableControls;

          sL_SQL := 'SELECT SUM(NVL(SaleAmount,0)) SaleAmount,SUM(NVL(TaxAmount,0)) TaxAmount,' +
                    'SUM(NVL(InvAmount,0)) InvAmount FROM INV016 ';

          sL_SQL := sL_SQL + sL_Where + ' ORDER BY SEQ';

          with cdsCommon do
          begin
            Close;
            CommandText := sL_SQL;
            Open;

            fI_SaleAmount := FieldByName('SaleAmount').AsFloat;
            fI_TaxAmount := FieldByName('TaxAmount').AsFloat;
            fI_InvAmount := FieldByName('InvAmount').AsFloat;
          end;
        end;
      end;
    //end;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMainJ.getInv017Data(ASeq: String);
begin
  adoQrySelectInv017.Close;
  adoQrySelectInv017.SQL.Text := Format(
    '   SELECT * FROM INV017            ' +
    '    WHERE SEQ = ''%s''             ' +
    '    ORDER BY BILLID, BILLIDITEMNO  ', [ASeq] );
  adoQrySelectInv017.Open;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMainJ.updateInv016ShouldBeAssigned(sI_SEQ,
  sI_ShouldBeAssigned, sI_CompID, sI_StartChargeDate,
  sI_EndChargeDate: String; UpdateInv017:Boolean; aMinSeq, aMaxSeq: String);
var
    sL_SQL,sL_Where : String;
begin
    sL_SQL := 'UPDATE INV016 SET ShouldBeAssigned=''' + sI_ShouldBeAssigned + ''' ';

    if UpdateInv017 then
      if sI_ShouldBeAssigned ='Y' then
        sL_SQL := sL_SQL + ', SALEAMOUNT=IMPORTSALEAMOUNT, INVAMOUNT=IMPORTINVAMOUNT, '+
                           ' TAXAMOUNT=IMPORTTAXAMOUNT '
      else
        sL_SQL := sL_SQL + ', SALEAMOUNT=0, INVAMOUNT=0,TAXAMOUNT=0 ';

    sL_Where := ' WHERE COMPID= '+''''+sI_CompID+''''+
      ' AND CHARGEDATE between TO_DATE('+''''+sI_StartChargeDate+''','''+
      Date_Format+''') AND TO_DATE('+''''+sI_EndChargeDate+''','''+
      Date_Format+''') AND BEASSIGNEDINVID=''N'' AND ISVALID=''Y'' '+
      ' AND STOPFLAG=0 ';

    if ( aMinSeq <> EmptyStr ) and ( aMaxSeq <> EmptyStr ) then
      sL_Where := ( sL_Where + Format( ' AND SEQ BETWEEN ''%s'' AND ''%s'' ',
      [aMinSeq, aMaxSeq] ) );

    if sI_SEQ <> '' then
      sL_Where := sL_Where + ' AND SEQ=''' + sI_SEQ + '''';

    sL_SQL := sL_SQL + sL_Where;

    executeSQL(sL_SQL);

    if UpdateInv017 then
      begin
        //?D???}???A???}???A?????]?n?????]?w
        sL_SQL := 'UPDATE INV017 SET SHOULDBEASSIGNED=''' + sI_ShouldBeAssigned + '''' +
                  ' WHERE SEQ IN ' +
                  '  (SELECT SEQ FROM INV016 ' +
                  sL_Where +
                  ') ';
        executeSQL(sL_SQL);
      end;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMainJ.updateInv017ShouldBeAssigned(sI_SEQ,
  sI_ShouldBeAssigned, sI_BILLID,
  sI_BILLIDITEMNO: String);
var
    sL_SQL,sL_Where : String;
begin
    sL_SQL := 'UPDATE INV017 SET ShouldBeAssigned=''' + sI_ShouldBeAssigned + '''';

    sL_Where := ' WHERE BILLID= '+''''+sI_BILLID+''''+
      ' AND BILLIDITEMNO= '+''''+sI_BILLIDITEMNO+''' ';

    if sI_SEQ <> '' then
      sL_Where := sL_Where + ' AND SEQ=''' + sI_SEQ + '''';

    sL_SQL := sL_SQL + sL_Where;

    executeSQL(sL_SQL);

end;

procedure TdtmMainJ.listLastInvDate(sI_CompID, sI_YearMonth: String);
var
    sL_SQL : String ;
begin
    Inv099DataSet.Close ;
    with adoQryCommon do
    begin
      Close;
      SQL.Clear ;
      if sI_YearMonth='' then
        sL_SQL := 'SELECT A.IDENTIFYID1,A.IDENTIFYID2,A.COMPID,A.STARTNUM,A.PREFIX, A.LASTINVDATE,A.YEARMONTH,A.CURNUM,A.MEMO FROM INV099 A, INV001 B'+
                ' WHERE A.COMPID = :COMPID AND A.COMPID=B.COMPID AND INVFORMAT=1 AND USEFUL=''Y'' ORDER BY A.YEARMONTH DESC '
      else
        sL_SQL := 'SELECT A.IDENTIFYID1,A.IDENTIFYID2,A.COMPID,A.STARTNUM,A.PREFIX, A.LASTINVDATE,A.YEARMONTH,A.CURNUM,A.MEMO FROM INV099 A, INV001 B'+
                ' WHERE A.COMPID = :COMPID AND A.COMPID=B.COMPID AND INVFORMAT=1 AND USEFUL=''Y''' + ' AND YEARMONTH=' + '''' + sI_YearMonth + '''' + ' ORDER BY A.YEARMONTH DESC ';

      SQL.Add(sL_SQL) ; //???X?????q?????q?l?????o??
      adoQryCommon.Parameters.ParamByName('COMPID').Value := sI_CompID ;
      Open;
    end;
    Inv099DataSet.Open ;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMainJ.getAllInv001Data : Integer;
begin
  adoInv001Code.Close;
  adoInv001Code.SQL.Text :=
    ' select * from inv001 order by to_number( compid ) ';
  adoInv001Code.Open;
  Result := adoInv001Code.RecordCount;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMainJ.getInv002Data(bI_Select: Integer;sI_CustID1Str,sI_CustID2Str,
                        sI_CustName,sI_Sno,sI_Tel,sI_CompID : String) : Boolean;
var
    sL_SQL,sL_Where : String;
begin
    sL_SQL := 'select inv002.*,inv019.businessid from inv002 ,inv019 ';

    sL_Where := ' where inv002.custid = inv019.custid(+) '+
                '   and inv002.IdentifyId1=' + QuotedStr(IDENTIFYID1) +
                '   and inv002.IdentifyId2=' + IDENTIFYID2 +
                '   and inv002.CompID=' + QuotedStr(sI_CompID);

    Case bI_Select of
       1:
          sL_Where := sL_Where + ' and inv002.CustID between '+QuotedStr(sI_CustID1Str)+
                                 ' and '+QuotedStr(sI_CustID2Str);
       2:
          sL_Where := sL_Where + ' and inv002.CustsName like '+QuotedStr('%'+sI_CustName+'%');
       3:
          sL_Where := sL_Where + ' and inv019.BUSINESSID= '+QuotedStr(sI_Sno);
       4:
          sL_Where := sL_Where + ' and inv002.Tel1= '+QuotedStr(sI_Tel)+
                                 '  or inv002.Tel2= '+QuotedStr(sI_Tel)+
                                 '  or inv002.Tel3= '+QuotedStr(sI_Tel);
       5: sL_Where := sL_Where + ' and 1 = 2 '; 
    end;
    sL_SQL := sL_SQL + sL_Where;

    with adoInv002Code do
    begin
      Close;
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;
      if RecordCount > 0 then
        Result := True
      else
        Result := False;
    end;
end;

function TdtmMainJ.checkPK(sI_SQL: String): Boolean;
var
    nL_Counts : Integer;
begin
  with adoQryCommon do
  begin
    Close;
    SQL.Clear;
    SQL.Add(sI_SQL);
    Open;
    nL_Counts := adoQryCommon.RecordCount;
    if nL_Counts > 0 then
      Result := true
    else
      Result := false;
    Close;
  end;
end;

function TdtmMainJ.getInv019Data(sI_CompID, sI_CustID: String): Integer;
var
    sL_SQL,sL_Where : String;
begin
    sL_SQL := 'select * from inv019 ';
    sL_Where := ' where IdentifyId1=''' + IDENTIFYID1 +
                ''' and IdentifyId2=' + IDENTIFYID2 +
                ' and CompID=''' + sI_CompID +
                ''' and CustID=''' + sI_CustID + '''';

    sL_SQL := sL_SQL + sL_Where;

    with adoInv019Code do
    begin
      Close;
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;

      Result := RecordCount;
    end;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMainJ.getAllInv006Data: Integer;
begin
  adoInv006Code.Close;
  adoInv006Code.SQL.Text := ' SELECT * FROM INV006 ORDER BY ITEMID ';
  adoInv006Code.Open;
  Result := adoInv006Code.RecordCount;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMainJ.getAllInv012Data: Integer;
begin
  adoInv012Code.Close;
  adoInv012Code.SQL.Text := ' SELECT * FROM INV012 ORDER BY GROUPID ';
  adoInv012Code.Open;
  Result := adoInv012Code.RecordCount;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMainJ.getAllInv022Data: Integer;
begin
  adoInv022Code.Close;
  adoInv022Code.SQL.Text :=
    ' SELECT * FROM INV022  ORDER BY ITEMID ';
  adoInv022Code.Open;
  Result := adoInv022Code.RecordCount;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMainJ.checkCdsPK(I_CDS: TDataSet; sI_CodeColName,sI_DescColName,sI_Code,
  sI_Desc: String): Boolean;
var
    sL_ColStr : String;
begin
    if sI_Desc <> '' then
    begin
      sL_ColStr := sI_CodeColName + ';' + sI_DescColName;
      if I_CDS.Locate(sL_ColStr,VarArrayOf([sI_Code,sI_Desc]) , []) then
        Result := true
      else
        Result := false;
    end
    else
    begin
      sL_ColStr := sI_CodeColName;
      if I_CDS.Locate(sL_ColStr,VarArrayOf([sI_Code]) , []) then
        Result := true
      else
        Result := false;
    end;
end;

function TdtmMainJ.getInvCreating(sL_CompID: String): String;
var
  sL_SQL : String ;
begin
  sL_SQL := 'SELECT INVCREATING FROM INV003 WHERE COMPID ='''+sL_CompID +'''';
  result := OpenSQL(sL_SQL);
end;

function TdtmMainJ.assignInvoiceWithSF(nI_TotalUnit: Integer;
  fI_TotalSaleAmount, fI_TotalTaxAmount,fI_TotalInvAmount: double;
  sI_UptEn,sI_CompID,sI_HowToCreate,sI_InvDate,sI_InvYearMonth,sI_StartDate,sI_EndDate,sI_PrefixString,sI_MisDbOwner: String;
  nI_Order : Integer;bI_InvDateEqualChargeDate : Boolean; sI_ShowFaci,sI_StarCMTVmail : Integer;
  var sI_LogDateTime: String): Boolean;

var
    sL_SystemID,sL_Msg : String;
    fL_InvDateEqualChargeDate : double;
    fL_RetCode : Double;
    fL_SaleAmount,fL_TaxAmount,fL_InvAmount : Double;
    fL_ESaleAmount,fL_ETaxAmount,fL_EInvAmount: Double;
    fL_NoESaleAmount,fL_NoETaxAmount,fL_NoEInvAmount: Double;
    nL_TotalUnit : Integer;
    nL_ETotalUnit , nL_NoETotalUnit: Integer;
    sL_StartInvID,sL_StopInvID, aAllInvId, aCalcSQL, aCalcSQL2: String;
    aCalcNoElectronSQL, aCalcElectronSQL: String;
begin
  Result := False;

  if not dtmMain.InvConnection.Connected then
    dtmMain.InvConnection.Connected := True;
    //dtmMain.InvConnection.ConnectionString
  sL_SystemID := dtmMainJ.GetXMLAttribute( Format(
    'SELECT INVPARAM FROM INV003 WHERE COMPID = ''%s'' ', [dtmMain.getCompID] ), 'SysID' );

  if bI_InvDateEqualChargeDate then
    fL_InvDateEqualChargeDate := 1 //1=>?o?????????????O????;
  else
    fL_InvDateEqualChargeDate := 2;//2=>?o???????????????O????

  try
    spAssignInvID.ProcedureName := 'SF_ASSIGNINVID';
    spAssignInvID.Parameters.Clear;
    //spAssignInvID.Parameters.Count
    //spAssignInvID.Parameters[0].Name
    spAssignInvID.Parameters.Refresh;
    spAssignInvID.Parameters.ParamByName('P_USER').Value := sI_UptEn;
    spAssignInvID.Parameters.ParamByName('P_COMPID').Value := sI_CompID;
    if dtmMain.GetLinkToMIS then
       spAssignInvID.Parameters.ParamByName('P_LINKTOMIS').Value := 'Y'
    else
      spAssignInvID.Parameters.ParamByName('P_LINKTOMIS').Value := 'N';
    spAssignInvID.Parameters.ParamByName('P_DBLINK').Value := dtmMain.GetMisDbLink;
    spAssignInvID.Parameters.ParamByName('P_HOWTOCREATE').Value := sI_HowToCreate;
    spAssignInvID.Parameters.ParamByName('P_INVDATEEQUALTOCHARGEDATE').Value := fL_InvDateEqualChargeDate;
    spAssignInvID.Parameters.ParamByName('P_INVDATE').Value := sI_InvDate;
    spAssignInvID.Parameters.ParamByName('P_INVYEARMONTH').Value := sI_InvYearMonth;
    spAssignInvID.Parameters.ParamByName('P_CHARGESTARTDATE').Value := sI_StartDate;
    spAssignInvID.Parameters.ParamByName('P_CHARGESTOPDATE').Value := sI_EndDate;
    spAssignInvID.Parameters.ParamByName('P_IDENTIFYID1').Value := IDENTIFYID1;
    spAssignInvID.Parameters.ParamByName('P_IDENTIFYID2').Value := IDENTIFYID2;
    spAssignInvID.Parameters.ParamByName('P_SYSTEMID').Value := sL_SystemID;
    spAssignInvID.Parameters.ParamByName('p_PREFIXSTRING').Value := sI_PrefixString;
    spAssignInvID.Parameters.ParamByName('P_ORDERBY').Value := nI_Order;
    spAssignInvID.Parameters.ParamByName('P_MISDBOWNER').Value := sI_MisDbOwner;
    //#5358 ?W?[ShowFaci ?P?_?b?}???o?????O?_?n?N?]???????g?JMemo1 By Kin 2009/12/14
    spAssignInvID.Parameters.ParamByName('P_SHOWFACI').Value := sI_ShowFaci;
    //#6080 ?W?[StarCMTVMail ???? By Kin 2011/0823
    spAssignInvID.Parameters.ParamByName('P_STARCMTVMAIL').Value := sI_StarCMTVmail;
  except
    on E: Exception do
    begin
      WarningMsg( Format( '?????}???o???{?????????????~, ???]: %s?C', [E.Message] ) );
      Exit;
    end;
  end;


  spAssignInvID.Connection.BeginTrans;
  try
    spAssignInvID.ExecProc;
    fL_RetCode := StrToInt( VarToStrDef(
      spAssignInvID.Parameters.ParamByName( 'p_RetCode' ).Value, '-99' ) );
    {}
    sL_Msg := VarToStrDef(
      spAssignInvID.Parameters.ParamByName( 'p_RetMsg' ).Value, '?????{?????????G???????C' );
    {}
    sI_LogDateTime := VarToStrDef(
      spAssignInvID.Parameters.ParamByName( 'p_LogDateTime' ).Value, EmptyStr );
    {}
    if ( fL_RetCode = 0 ) then
      spAssignInvID.Connection.CommitTrans
    else
      spAssignInvID.Connection.RollbackTrans;
    {}
  except
    on E: Exception do
    begin
      fL_RetCode := -99;
      sL_Msg := E.Message;
      spAssignInvID.Connection.RollbackTrans;
    end;
  end;

  if ( fL_RetCode <> 0 ) then
  begin
    WarningMsg( '?o???}??????!???]:' + sL_Msg );
    Exit;
  end;

  with frmInvA01 do
  begin
    {---------------------------????????-----------------------------------------}
    aCalcSQL := Format(
      ' SELECT SUM( A.SALEAMOUNT ) SALEAMOUNT, ' +
      '        SUM( A.TAXAMOUNT ) TAXAMOUNT,   ' +
      '        SUM( A.INVAMOUNT ) INVAMOUNT    ' +
      '  FROM INV031 A                         ' +
      ' WHERE A.IDENTIFYID1 = ''%s''           ' +
      '   AND A.IDENTIFYID2 = %s               ' +
      '   AND A.COMPID = ''%s''                ',
    [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID] );

    dtmMainJ.openSQLGetAmount( aCalcSQL, fL_SaleAmount, fL_TaxAmount, fL_InvAmount );
    {----------------------------------------------------------------------------}
    {#5945 ?W?[?q?l?o???P?D?q?l?o???????p By Kin 2011/03/15}
    if dtmMain.GetStarEInvoice then
    begin
      {------------------------------?q?l?o??--------------------------------------}
      aCalcSQL := Format(
      ' SELECT SUM( A.SALEAMOUNT ) SALEAMOUNT, ' +
      '        SUM( A.TAXAMOUNT ) TAXAMOUNT,   ' +
      '        SUM( A.INVAMOUNT ) INVAMOUNT    ' +
      '  FROM INV031 A                         ' +
      ' WHERE A.IDENTIFYID1 = ''%s''           ' +
      '   AND A.IDENTIFYID2 = %s               ' +
      '   AND A.COMPID = ''%s''                ' +
      '   AND INVOICEKIND = 1                  ' ,
      [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID] );
      dtmMainJ.openSQLGetAmount( aCalcSQL, fL_ESaleAmount, fL_ETaxAmount, fL_EInvAmount );
      {----------------------------------------------------------------------------}

      {------------------------------?D?q?l?o??-------------------------------------}
      aCalcSQL := Format(
      ' SELECT SUM( A.SALEAMOUNT ) SALEAMOUNT, ' +
      '        SUM( A.TAXAMOUNT ) TAXAMOUNT,   ' +
      '        SUM( A.INVAMOUNT ) INVAMOUNT    ' +
      '  FROM INV031 A                         ' +
      ' WHERE A.IDENTIFYID1 = ''%s''           ' +
      '   AND A.IDENTIFYID2 = %s               ' +
      '   AND A.COMPID = ''%s''                ' +
      '   AND INVOICEKIND <> 1                 ' ,
      [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID] );
      dtmMainJ.openSQLGetAmount( aCalcSQL, fL_NoESaleAmount, fL_NoETaxAmount, fL_NoEInvAmount );
      {----------------------------------------------------------------------------}
    end;

    aCalcSQL2 := Format(
      ' SELECT COUNT(1) COUNT                   ' +
      '   FROM INV031 A                         ' +
      '  WHERE A.IDENTIFYID1 = ''%s''           ' +
      '    AND A.IDENTIFYID2 = %s               ' +
      '    AND A.COMPID = ''%s''                ',
    [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID] );

    nL_TotalUnit := StrToInt( dtmMainJ.openSQL( aCalcSQL2 ) ) ;

    {#5945 ?W?[?q?l?o???P?D?q?l?o???????p By Kin 2011/03/15}
    if dtmMain.GetStarEInvoice then
    begin
      {?q?l?o??}
      aCalcSQL2 := Format(
      ' SELECT COUNT(1) COUNT                   ' +
      '   FROM INV031 A                         ' +
      '  WHERE A.IDENTIFYID1 = ''%s''           ' +
      '    AND A.IDENTIFYID2 = %s               ' +
      '    AND A.COMPID = ''%s''                ' +
      '    AND A.INVOICEKIND = 1                ',
      [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID] );
      nL_ETotalUnit := StrToInt( dtmMainJ.openSQL( aCalcSQL2 ) ) ;

      {?D?q?l?o??}
      aCalcSQL2 := Format(
      ' SELECT COUNT(1) COUNT                   ' +
      '   FROM INV031 A                         ' +
      '  WHERE A.IDENTIFYID1 = ''%s''           ' +
      '    AND A.IDENTIFYID2 = %s               ' +
      '    AND A.COMPID = ''%s''                ' +
      '    AND A.INVOICEKIND <> 1               ',
      [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID] );
      nL_NoETotalUnit := StrToInt( dtmMainJ.openSQL( aCalcSQL2 ) ) ;

    end;


    lblSaleAmount3.Caption := '?P?f?`?B: $'+ FormatFloat( '#,##0', fL_SaleAmount);


    lblTaxAmount3.Caption := '?|???`?B: $'+ FormatFloat( '#,##0', fL_TaxAmount);
    lblInvAmount3.Caption := '?`?p: $'+ FormatFloat( '#,##0', fL_InvAmount) ;
    lblAssignInvCount.Caption := Format( '?o???}???i??:%d?i', [nL_TotalUnit] );

    dtmMainJ.getInv031StartAndStopInvID( dtmMain.getCompID, sL_StartInvID,
      sL_StopInvID, aAllInvId );

    lblStartInvID.Caption := '?o???}???_?l???X: ' + sL_StartInvID;
    lblStopInvID.Caption := '?o???}???I?????X: ' + sL_StopInvID;

    {#5945 ?W?[?q?l?o???P?D?q?l?o???????p By Kin 2011/03/15}
    if dtmMain.GetStarEInvoice then
    begin
      pnlTotal4.Visible := True;
      {?P?f?`?B}
      lblNoESaleAmt.Caption := FormatFloat('#,##0',fL_NoESaleAmount);
      lblESaleAmt.Caption := FormatFloat('#,##0',fL_ESaleAmount);
      {?|???`?B}
      lblNoETaxAmt.Caption := FormatFloat('#,##0',fL_NoETaxAmount);
      lblETaxAmt.Caption := FormatFloat('#,##0',fL_ETaxAmount);
      {?`?p}
      lblNoEInvAmt.Caption := FormatFloat( '#,##0', fL_NoEInvAmount);
      lblEInvAmt.Caption := FormatFloat( '#,##0', fL_EInvAmount);
      {?}???i??}
      lblNoEInvCount.Caption := IntToStr( nL_NoETotalUnit );
      lblEInvCount.Caption := IntToStr( nL_ETotalUnit);

    end;

  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMainJ.getLastInvDate(sI_CompID, sI_InvYearMonth,
  sI_IdentifyID1, sI_IdentifyID2: String): String;
var
  sL_SQL,sL_PreFix,sL_StartNum : String ;
  sL_LastInvDate,sL_MaxInvDate : String ;
begin
    sL_MaxInvDate := '';
    with adoQryCommon do
    begin
      //???X?????o???r?y???j????
      cdsPrefix.First;
      while not cdsPrefix.Eof do
      begin

        sL_PreFix := cdsPrefix.FieldByName('PREFIX').AsString;
        sL_StartNum := cdsPrefix.FieldByName('StartNum').AsString;

        Close;
        SQL.Clear ;
        sL_SQL := 'SELECT TO_CHAR(MAX(LASTINVDATE), ''YYYY/MM/DD'') FROM INV099'+
                  ' WHERE COMPID = '+ '''' + sI_CompID + '''' +
                  ' AND YEARMONTH = '+ '''' + sI_InvYearMonth + '''' +
                  ' AND USEFUL =''Y'' '+
                  ' AND INVFORMAT = 1 AND IDENTIFYID1 = ' + '''' + sI_IdentifyID1 + '''' +
                  ' AND IDENTIFYID2 = ' + '''' + sI_IdentifyID2 + '''' +
                  ' AND PREFIX=''' +  sL_PreFix + '''' +
                  ' AND STARTNUM=' + sL_StartNum;
        SQL.Add(sL_SQL) ;

        Open;
        sL_LastInvDate := Fields[0].AsString;


        if sL_MaxInvDate = '' then
          sL_MaxInvDate := sL_LastInvDate
        else
        begin
          if sL_LastInvDate > sL_MaxInvDate then
            sL_MaxInvDate := sL_LastInvDate
        end;
        Close;

        cdsPrefix.Next;
      end;
    end;

    Result := sL_MaxInvDate;

end;

function TdtmMainJ.getCanUseInvCounts(sI_CompID, sI_InvYearMonth,
  sI_IdentifyID1, sI_IdentifyID2: String): Integer;
var
    sL_SQL,sL_PreFix,sL_StartNum : String;
    nL_Counts,nL_TotalConts : Integer;
begin
    nL_TotalConts := 0;
    with adoQryCommon do
    begin
      //???X?????o???r?y???j????
      cdsPrefix.First;
      while not cdsPrefix.Eof do
      begin

        sL_PreFix := cdsPrefix.FieldByName('PREFIX').AsString;
        sL_StartNum := cdsPrefix.FieldByName('StartNum').AsString;

        Close;
        SQL.Clear ;
        sL_SQL := 'SELECT SUM(ENDNUM - CURNUM+1) INVCOUNT FROM INV099'+
                  ' WHERE COMPID = '+ '''' + sI_CompID + '''' +
                  ' AND YEARMONTH = '+ '''' + sI_InvYearMonth + '''' +
                  ' AND USEFUL =''Y'' '+
                  ' AND INVFORMAT = 1 AND IDENTIFYID1 = ' + '''' + sI_IdentifyID1 + '''' +
                  ' AND IDENTIFYID2 = ' + '''' + sI_IdentifyID2 + '''' +
                  ' AND PREFIX=''' +  sL_PreFix + '''' +
                  ' AND STARTNUM=''' + sL_StartNum + '''';
        SQL.Add(sL_SQL) ;

        Open;
        nL_Counts := Fields[0].AsInteger;

        nL_TotalConts := nL_TotalConts + nL_Counts;

        Close;

        cdsPrefix.Next;
      end;
    end;

    Result := nL_TotalConts;


end;

procedure TdtmMainJ.openSQLGetAmount(sI_SQL: String; var fL_SaleAmount,
  fL_TaxAmount, fL_InvAmount: double);
begin
  with adoQryCommon do
  begin
    Close;
    SQL.Clear;
    SQL.Add(sI_SQL);
    Open;
    fL_SaleAmount := FieldByName('SALEAMOUNT').AsInteger;
    fL_TaxAmount := FieldByName('TAXAMOUNT').AsInteger;
    fL_InvAmount := FieldByName('INVAMOUNT').AsInteger;
    Close;
  end;
end;

function TdtmMainJ.GetXMLAttribute(const aSQL, aAttributeName: String): String;
begin
  XMLDoc.XML.Text := dtmMainJ.OpenSQL( aSQL );
  XMLDoc.Active := True;
  try
     Result := XMLDoc.DocumentElement.Attributes[aAttributeName];
  finally
    XMLDoc.Active := False;
  end;
end;

procedure TdtmMainJ.reactiveDataSet(I_DataSet: TDataSet);
begin
  I_DataSet.Close;
  I_DataSet.Open;
end;

procedure TdtmMainJ.InsertInv016(dI_CurSeq, sI_CompID, sI_CustID, sI_Tel,
  sI_BusinessID, sI_Title, sI_ZipCode, sI_InvAddr, sI_MailAddr, sI_BeAssignedInvID,
  sI_IsValid, sI_IsPreCreated, sI_UserName, sI_ChargeTitle,
   aInvUseId, aInvUseDesc,sI_InvType, sI_Email,
   sI_Contmobile, sI_GiveId, sI_GiveName: String);
begin
  cdsQryInsertInv016.Append;
  cdsQryInsertInv016.FieldByName('SEQ').AsString := dI_CurSeq;
  cdsQryInsertInv016.FieldByName('COMPID').AsString := sI_CompID;
  cdsQryInsertInv016.FieldByName('CUSTID').AsString := sI_CustID;
  cdsQryInsertInv016.FieldByName('TEL').AsString := sI_Tel;
  cdsQryInsertInv016.FieldByName('BUSINESSID').AsString := sI_BusinessID;
  cdsQryInsertInv016.FieldByName('TITLE').AsString := sI_Title;
  cdsQryInsertInv016.FieldByName('ZIPCODE').AsString := sI_ZipCode;
  cdsQryInsertInv016.FieldByName('INVADDR').AsString := sI_InvAddr;
  cdsQryInsertInv016.FieldByName('MAILADDR').AsString := sI_MailAddr;
  cdsQryInsertInv016.FieldByName('BEASSIGNEDINVID').AsString := sI_BeAssignedInvID;
  cdsQryInsertInv016.FieldByName('SALEAMOUNT').AsFloat := 0;
  cdsQryInsertInv016.FieldByName('TAXAMOUNT').AsFloat := 0;
  cdsQryInsertInv016.FieldByName('INVAMOUNT').AsFloat := 0;
  cdsQryInsertInv016.FieldByName('ISVALID').AsString := sI_IsValid;
  cdsQryInsertInv016.FieldByName('HOWTOCREATE').AsString := sI_IsPreCreated;
  cdsQryInsertInv016.FieldByName('UPTTIME').AsDateTime := now;
  cdsQryInsertInv016.FieldByName('UPTEN').AsString := sI_UserName;
  cdsQryInsertInv016.FieldByName('CHARGETITLE').AsString := sI_ChargeTitle;
  cdsQryInsertInv016.FieldByName('INVUSEID').AsString := aInvUseId;
  cdsQryInsertInv016.FieldByName('INVUSEDESC').AsString := aInvUseDesc;
  //?w?]?????????n?}???o??
  cdsQryInsertInv016.FieldByName('SHOULDBEASSIGNED').AsString := 'Y';
  cdsQryInsertInv016.FieldByName('IMPORTSALEAMOUNT').AsFloat := 0;
  cdsQryInsertInv016.FieldByName('IMPORTTAXAMOUNT').AsFloat := 0;
  cdsQryInsertInv016.FieldByName('IMPORTINVAMOUNT').AsFloat := 0;
  {#5668 ?W?[5?????? By Kin 2010/07/09}
  cdsQryInsertInv016.FieldByName( 'INVOICEKIND' ).AsString := sI_InvType;
  cdsQryInsertInv016.FieldByName( 'EMAIL' ).AsString := sI_Email;
  cdsQryInsertInv016.FieldByName( 'CONTMOBILE' ).AsString := sI_Contmobile;
  cdsQryInsertInv016.FieldByName( 'GIVEUNITID' ).AsString := sI_GiveId;
  cdsQryInsertInv016.FieldByName( 'GIVEUNITDESC' ).AsString := sI_GiveName;
  cdsQryInsertInv016.Post ;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMainJ.DeleteInv016(const dI_CurSeq: String);
begin
  if cdsQryInsertInv016.Locate( 'SEQ', dI_CurSeq, [] ) then
    cdsQryInsertInv016.Delete;
  cdsQryInsertInv016.Last;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMainJ.UpdateInv016(const dI_CurSeq: String; const dI_SaleAmount,
  dI_MasterTaxAmount: Double; const dI_ChargeDate, dI_TaxType, dI_TaxRate: String);
var
  aSaleAmount, aTaxAmount, aInvAmount: Double;
begin
  if cdsQryInsertInv016.Locate( 'SEQ', dI_CurSeq, [] ) then
  begin
    aInvAmount := ( cdsQryInsertInv016.FieldByName('INVAMOUNT').AsFloat + dI_SaleAmount );
    aTaxAmount := ( cdsQryInsertInv016.FieldByName('TAXAMOUNT').AsFloat + dI_MasterTaxAmount );
    aSaleAmount := ( aInvAmount - aTaxAmount );
    cdsQryInsertInv016.Edit;
    cdsQryInsertInv016.FieldByName('SALEAMOUNT').AsFloat := aSaleAmount;
    cdsQryInsertInv016.FieldByName('TAXAMOUNT').AsFloat := aTaxAmount;
    cdsQryInsertInv016.FieldByName('INVAMOUNT').AsFloat := aInvAmount;

    cdsQryInsertInv016.FieldByName('IMPORTSALEAMOUNT').AsFloat := aSaleAmount;
    cdsQryInsertInv016.FieldByName('IMPORTTAXAMOUNT').AsFloat := aTaxAmount;
    cdsQryInsertInv016.FieldByName('IMPORTINVAMOUNT').AsFloat := aInvAmount;

    cdsQryInsertInv016.FieldByName('CHARGEDATE').AsDateTime :=
      StrToDateTime( FormatMaskText( '####/##/##;0;_', dI_ChargeDate ) );
    cdsQryInsertInv016.FieldByName('TAXTYPE').AsString := dI_TaxType;
    cdsQryInsertInv016.FieldByName('TAXRATE').AsString := dI_TaxRate;
    cdsQryInsertInv016.Post ;
    cdsQryInsertInv016.Last;
  end;
end;

{ ---------------------------------------------------------------------------- }


procedure TdtmMainJ.ResetInv016;
begin
  if cdsQryInsertInv016.Active then
  begin
    cdsQryInsertInv016.EmptyDataSet;
    cdsQryInsertInv016.Close;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMainJ.InsertInv017(dI_CurSeq, sI_BillID,
  sI_BillIDItemNo, sI_MasterTaxType, sI_DetailChargeDate, sI_ItemID,
  sI_Description, sI_Quantity, sI_UnitPrice, sI_MasterTaxRate,
  sI_DetailTaxAmount, sI_TotalAmount, sI_StartDate, sI_EndDate,
  sI_ChargeEn, sI_ServiceType, aLinkToMis, aAccountNo,
  aCombCitemCode, aCombCitemName : String);
var
  sL_DetailChargeDate,sL_StartDate,sL_EndDate : String ;
begin
  sL_DetailChargeDate := getYearMonthDay2(sI_DetailChargeDate);
  sL_StartDate := getYearMonthDay2(sI_StartDate);
  sL_EndDate := getYearMonthDay2(sI_EndDate);
  cdsQryInsertInv017.Append;
  cdsQryInsertInv017.FieldByName('SEQ').AsString := dI_CurSeq;
  cdsQryInsertInv017.FieldByName('BILLID').AsString := sI_BillID;
  cdsQryInsertInv017.FieldByName('BILLIDITEMNO').AsString := sI_BillIDItemNo;
  cdsQryInsertInv017.FieldByName('TAXTYPE').AsString := sI_MasterTaxType;
  cdsQryInsertInv017.FieldByName('CHARGEDATE').AsDateTime := StrToDate(sL_DetailChargeDate);
  cdsQryInsertInv017.FieldByName('ITEMID').AsString := sI_ItemID;
  cdsQryInsertInv017.FieldByName('DESCRIPTION').AsString := sI_Description;
  cdsQryInsertInv017.FieldByName('QUANTITY').AsInteger := StrToInt(sI_Quantity);
//  cdsQryInsertInv017.FieldByName('UNITPRICE').AsInteger := StrToInt(sI_UnitPrice);
//  cdsQryInsertInv017.FieldByName('TAXRATE').AsInteger := StrToInt(sI_MasterTaxRate);
  //?J???p???I?{???|?X??,?{?b????AsString By Kin 2011/03/01
  cdsQryInsertInv017.FieldByName('UNITPRICE').AsString := sI_UnitPrice;
  cdsQryInsertInv017.FieldByName('TAXRATE').AsString := sI_MasterTaxRate;
  cdsQryInsertInv017.FieldByName('TAXAMOUNT').AsInteger := StrToInt(sI_DetailTaxAmount);
  cdsQryInsertInv017.FieldByName('TOTALAMOUNT').AsInteger := StrToInt(sI_TotalAmount);
  if (sI_StartDate<>'') then
    cdsQryInsertInv017.FieldByName('STARTDATE').AsDateTime := StrToDate(sL_StartDate);
  if (sI_EndDate<>'') then
    cdsQryInsertInv017.FieldByName('ENDDATE').Value := StrToDate(sL_EndDate);
  cdsQryInsertInv017.FieldByName('CHARGEEN').AsString := trim(sI_ChargeEn);
  cdsQryInsertInv017.FieldByName('SERVICETYPE').AsString := sI_ServiceType;
  cdsQryInsertInv017.FieldByName('SHOULDBEASSIGNED').AsString := 'Y';
  cdsQryInsertInv017.FieldByName('LINKTOMIS').AsString := Nvl( aLinkToMis, 'Y' );
  cdsQryInsertInv017.FieldByName('ACCOUNTNO').AsString := aAccountNo;
  cdsQryInsertInv017.FieldByName('COMBCITEMCODE').AsString := aCombCitemCode;
  cdsQryInsertInv017.FieldByName('COMBCITEMNAME').AsString := aCombCitemName;

  cdsQryInsertInv017.Post;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMainJ.ResetInv017;
begin
  if cdsQryInsertInv017.Active then
  begin
    cdsQryInsertInv017.EmptyDataSet;
    cdsQryInsertInv017.Close;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMainJ.ResetWitchOne;
begin
  if VarIsNull( cdsWitchOne.Data ) then
    cdsWitchOne.CreateDataSet;
  cdsWitchOne.EmptyDataSet;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMainJ.DetailBelongWitchOne(const aCurSeq, aTaxType: String;
  const aTaxRate: Double): String;
begin
  Result := EmptyStr;
  { ?????P?@?????????D?????? } 
  if cdsQryInsertInv016.Locate( 'SEQ;TAXTYPE;TAXRATE',
    VarArrayOf( [aCurSeq, aTaxType, aTaxRate] ), [] ) then
  begin
    Result := cdsQryInsertInv016.FieldByName( 'SEQ' ).AsString;
  end
  { ?S????, ?h???s?????? }
  else if cdsWitchOne.Locate( 'TAXTYPE;TAXRATE',
    VarArrayOf( [ aTaxType, aTaxRate] ), [] ) then
  begin
    Result := cdsWitchOne.FieldByName( 'SEQ' ).AsString;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMainJ.InsertBelongWitchOne(const aCurSeq, aTaxType: String;
  const aTaxRate: Double);
begin
  cdsWitchOne.Append;
  cdsWitchOne.FieldByName( 'SEQ' ).AsString := aCurSeq;
  cdsWitchOne.FieldByName( 'TAXTYPE' ).AsString := aTaxType;
  cdsWitchOne.FieldByName( 'TAXRATE' ).AsFloat := aTaxRate;
  cdsWitchOne.Post;
end;

{ ---------------------------------------------------------------------------- }


procedure TdtmMainJ.getInv031StartAndStopInvID(sI_CompID: String;
  var sI_StartInvID, sI_StopInvID, sI_AllInvId: String);
var
  aTemp, aInvPrefix: String;
  aPos: Integer;
  aResult: String;
begin
  cdsQInv031.Close;
  cdsQInv031.CommandText := Format(
    ' SELECT MIN( INVID ) MININV,                   ' +
    '        MAX( INVID ) MAXINV                    ' +
    '   FROM ( SELECT INVID FROM INV031             ' +
    '           WHERE COMPID = ''%s''               ' +
    '           ORDER BY ROWID )                    ', [sI_CompID] );
  cdsQInv031.Open;
  sI_StartInvID := cdsQInv031.FieldByName('MININV').AsString;
  sI_StopInvID := cdsQInv031.FieldByName('MAXINV').AsString;
  cdsQInv031.Close;
  { }
  cdsQInv031.CommandText := Format(
   ' SELECT UNIQUE SUBSTR( INVID, 1, 2 ) INVPREFIX ' +
   '   FROM INV031                                 ' +
   '   WHERE COMPID = ''%s''                       ', [sI_CompID] );
  cdsQInv031.Open;
  cdsQInv031.First;
  while not cdsQInv031.Eof do
  begin
    aTemp := ( aTemp + cdsQInv031.FieldByName( 'INVPREFIX' ).AsString );
    cdsQInv031.Next;
    if not cdsQInv031.Eof then
      aTemp := ( aTemp + ';' );
  end;
  cdsQInv031.Close;
  { Ex: CV54338100~CV54338149,CV54338150~CV54338199 }
  aPos := 1;
  aInvPrefix := EmptyStr;
  aResult := EmptyStr;
  while aPos <= Length( aTemp ) do
  begin
    aInvPrefix := ExtractFieldName( aTemp, aPos );
    cdsQInv031.CommandText := Format(
      ' SELECT MIN( INVID ) INVID FROM INV031   ' +
      '  WHERE COMPID = ''%s''                  ' +
      '    AND SUBSTR( INVID, 1, 2 ) = ''%s''   ', [sI_CompID, aInvPrefix] );
    cdsQInv031.Open;
    aResult := ( aResult + cdsQInv031.FieldByName( 'INVID' ).AsString + '~' );
    cdsQInv031.Close;
    cdsQInv031.CommandText := Format(
      ' SELECT MAX( INVID ) INVID FROM INV031   ' +
      '  WHERE COMPID = ''%s''                  ' +
      '    AND SUBSTR( INVID, 1, 2 ) = ''%s''   ', [sI_CompID, aInvPrefix] );
    cdsQInv031.Open;
    aResult := ( aResult + cdsQInv031.FieldByName( 'INVID' ).AsString + ',' );
    cdsQInv031.Close;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMainJ.getInv016MinAndMaxChargeDate(sI_CompID, sI_CurMinSeq,
  sI_CurMaxSeq :String;var sI_MinChargeDate, sI_MaxChargeDate: String);
var
    sL_SQL,sL_Where : String;
begin
    sL_SQL := 'SELECT Min(ChargeDate) MinChargeDate,Max(ChargeDate) MaxChargeDate FROM INV016 ';

    sL_Where := ' WHERE COMPID= ''' + sI_CompID +
                ''' AND Seq between ''' + sI_CurMinSeq + ''' and ''' + sI_CurMaxSeq + '''';

    sL_SQL := sL_SQL + sL_Where;

    with adoQryCommon do
    begin
      Close;
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;

      sI_MinChargeDate := FieldByName('MinChargeDate').AsString;
      sI_MaxChargeDate := FieldByName('MaxChargeDate').AsString;
      Close;
    end;
end;

function TdtmMainJ.getUnusualInvID(nI_ConditionType: Integer; sI_Where1,
  sI_Where2: String; var aSQL: String): Integer;
var
  aWhere: String;
begin
    {
        ?o?????`?????d?? Type
        CONDITION_TYPE_BILLNO = 0;
        CONDITION_TYPE_CUSTID = 1;
        CONDITION_TYPE_CHARGEDATE = 2;
        CONDITION_TYPE_LOGDATE = 3;
        CONDITION_TYPE_ASSIGN_LOGDATE = 4;   //?????????o???????????d?????`????
    }
    if nI_ConditionType = CONDITION_TYPE_BILLNO then
    begin
      aWhere := Format ( ' WHERE BILLID = ''%s'' AND COMPID = ''%s'' ORDER BY BILLID ',
      [sI_Where1, dtmMain.getCompID] );
    end else
    if nI_ConditionType = CONDITION_TYPE_CUSTID then
    begin
      aWhere := Format( ' WHERE CUSTID = ''%s'' AND COMPID = ''%s'' ORDER BY CUSTID ',
      [sI_Where1, dtmMain.getCompID] );
    end else
    if nI_ConditionType = CONDITION_TYPE_CHARGEDATE then
    begin
      aWhere :=
        ' WHERE ( CHARGEDATE BETWEEN ' + TUstr.getOracleSQLDateStr(StrToDate(sI_Where1)) +
        '  AND ' + TUstr.getOracleSQLDateStr(StrToDate(sI_Where2)) + ' ) AND COMPID = ''' + dtmMain.getCompID + '''' +
        '  ORDER BY CHARGEDATE ';
    end
    else if nI_ConditionType = CONDITION_TYPE_LOGDATE then
    begin
      aWhere :=
        ' WHERE (LOGTIME  BETWEEN ' + TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_Where1 + ' 00:00:01')) +
        ' AND ' + TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_Where2 + ' 23:59:59')) + ' ) AND COMPID = ''' + dtmMain.getCompID + '''' +
        ' ORDER BY LOGTIME ';
    end
    else if nI_ConditionType = CONDITION_TYPE_ASSIGN_LOGDATE then
    begin
      aWhere := ' WHERE LOGTIME=' + TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_Where1)) +
       ' AND COMPID = ''' + dtmMain.getCompID + '''';
    end;

    if ( aWhere = EmptyStr ) then
      aWhere := Format( ' WHERE COMPID = ''%s'' ', [dtmMain.getCompID] );

    aSQL := ' SELECT * FROM INV033 ' + aWhere;
    cdsInv033.Close;
    cdsInv033.CommandText := aSQL;
    cdsInv033.Open;
    Result := cdsInv033.RecordCount;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMainJ.getInv099Data(aCompID, aYear, aYM: String): Integer;
var
   aSQL: String;
begin
  aSQL := ' SELECT IDENTIFYID1,           ' +
            '        IDENTIFYID2,         ' +
            '        COMPID,              ' +
            '        YEARMONTH,           ' +
            '        INVFORMAT,           ' +
            '        PREFIX,              ' +
            '        STARTNUM,            ' +
            '        ENDNUM,              ' +
            '        CURNUM,              ' +
            '        LASTINVDATE,         ' +
            '        USEFUL,              ' +
            '        UPTTIME,             ' +
            '        UPTEN,               ' +
            '        MEMO                 ' +
            '   FROM INV099               ' +
            '  WHERE COMPID = ' + QuotedStr( aCompID );

    if aYear <> '' then
      aSQL := aSQL + ' AND SUBSTR( YEARMONTH, 1, 4 ) = ' + QuotedStr( aYear );

    if aYM <> '' then
      aSQL := aSQL + ' AND YEARMONTH = ' + QuotedStr( aYM );

    aSQL := aSQL + ' ORDER BY YEARMONTH, INVFORMAT, PREFIX, STARTNUM  ';  

    adoInv099Code.Close;
    adoInv099Code.SQL.Text := aSQL;
    adoInv099Code.Open;
    adoInv099Code.FieldByName( 'STARTNUM' ).ReadOnly := False;
    adoInv099Code.FieldByName( 'ENDNUM' ).ReadOnly := False;
    adoInv099Code.FieldByName( 'CURNUM' ).ReadOnly := False;
    Result := adoInv099Code.RecordCount;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMainJ.changePrefixYM(sI_YM: String): String;
var
    sL_Year,sL_Month,sL_MonthStr : String;
begin
    sL_Year := IntToStr((StrToInt(Copy(sI_YM,1,4))));

    sL_Month := Copy(sI_YM,5,2);
    if (sL_Month = '01') or (sL_Month = '02')  then
      sL_MonthStr := '01'
    else if (sL_Month = '03') or (sL_Month = '04') then
      sL_MonthStr := '03'
    else if (sL_Month = '05') or (sL_Month = '06') then
      sL_MonthStr := '05'
    else if (sL_Month = '07') or (sL_Month = '08') then
      sL_MonthStr := '07'
    else if (sL_Month = '09') or (sL_Month = '10') then
      sL_MonthStr := '09'
    else if (sL_Month = '11') or (sL_Month = '12') then
      sL_MonthStr := '11';

    Result := sL_Year + sL_MonthStr;
end;

function TdtmMainJ.handleCheckBusinessID(sI_CompID, sI_SDate,
  sI_EDate: String): Integer;
var
    sL_SQL,sL_BusinessID,sL_InvID,sL_CustID : String;
    L_DataSet : TDataSet;
begin
    sL_SQL := 'SELECT CUSTID, BUSINESSID,INVID FROM INV007 WHERE COMPID='''+ sI_CompID +''''+
              ' AND (INVDATE BETWEEN TO_DATE('''+
               sI_SDate + ''',''YYYY/MM/DD'') AND '+
              ' TO_DATE(''' + sI_EDate + ''',''YYYY/MM/DD''))' ;

    L_DataSet := getDataSet(sL_SQL);
    with L_DataSet do
    begin
      while not Eof do
      begin
        sL_BusinessID := trim(FieldByName('BUSINESSID').AsString);
        if sL_BusinessID<>'' then
        begin
          sL_InvID := trim(FieldByName('INVID').AsString);
          sL_CustID := trim(FieldByName('CUSTID').AsString);
          if not dtmMain.checkBusinessID(sL_BusinessID) then
          begin
            with cdsCheckBusinessID do
            begin
              Append;
              FieldByName('Memo').AsString := '????: ' + sL_CustID + '  ???o???G '+ sL_InvID+' ?????s '+sL_BusinessID+' ???~';
              Post;
            end;
          end;
        end;
        Next;
      end;
      Close;
    end;
    Result := cdsCheckBusinessID.RecordCount;
end;

function TdtmMainJ.handleCheckJumpInvID(sI_CompID, sI_InvYM,
  sI_Prefix: String): Integer;
var
  sL_CompID,sL_InvYM,sL_Prefix,sL_SQLPrefix,sL_SQLInvice : String;
  L_Prefix,L_StartNum,L_EndNum : TStringList ;
  L_PrefixDataSet,L_InvoiceDataSet : TDataSet ;
  sL_InvID,sL_InvPrefix,sL_Year : String;
  sL_Month,sL_Date,sL_DaysOfMonth,sL_StrDate,sL_EndDate : String;
  nL_InvNO,nL_TempInvNO,nL_PrefixIndex,nL_TempPrefixIndex,J,K,nL_StartNum : Integer ;
  bL_RightFlag : Boolean;
begin
  sL_CompID := sI_CompID ;
  sL_InvYM := sI_InvYM ;
  sL_Prefix := sI_Prefix ;
  L_Prefix := TStringList.Create ;
  L_StartNum := TStringList.Create ;
  L_EndNum := TStringList.Create ;
  sL_SQLPrefix := 'SELECT PREFIX,STARTNUM,ENDNUM FROM INV099 ' ;
  if sL_Prefix='*'  then
  begin
    sL_SQLPrefix := sL_SQLPrefix + 'WHERE COMPID='''+sL_CompID+
    ''' AND YEARMONTH = '''+getYearMonthDay4(sL_InvYM)+
    ''' ORDER BY PREFIX';
  end
  else
  begin
    sL_SQLPrefix := sL_SQLPrefix + ' WHERE COMPID = '''+sL_CompID+
    ''' AND YEARMONTH ='''+getYearMonthDay4(sL_InvYM)+
    ''' AND PREFIX='''+sL_Prefix+''' ORDER BY PREFIX';
  end;

  L_PrefixDataSet := getDataSet(sL_SQLPrefix);
  if L_PrefixDataSet.IsEmpty then
    begin
      with cdsCheckJumpInvID do
      begin
        Append;
        FieldByName('Memo').AsString := '?r?y?????s?b';
        post;
      end;
      exit;
    end
  else
  begin
    with L_PrefixDataSet do
    begin
      while not Eof do
      begin
        L_Prefix.Add(FieldByName('PREFIX').AsString);
        L_StartNum.Add(FieldByName('STARTNUM').AsString);
        L_EndNum.Add(FieldByName('ENDNUM').AsString);
        Next;
      end;
      Close;
    end;
  end;

  sL_Year := copy(sL_InvYM,1,4);
  sL_Month := copy(sL_InvYM,5,2);
  sL_Date := IntToStr(StrToInt(sL_Year)-1911)+sL_Month+'01';
  sL_DaysOfMonth := IntToStr(TUdateTime.DaysOfMonth(sL_Date));

  sL_StrDate := sL_Year+'/'+sL_Month+'/01';
  sL_EndDate := sL_Year+'/'+sL_Month+'/'+sL_DaysOfMonth;

  sL_SQLInvice := 'SELECT INVID FROM INV007 WHERE COMPID='''+sL_CompID +''''+
        ' AND (INVDATE BETWEEN TO_DATE('''+
        sL_StrDate+''',''YYYY/MM/DD'') AND '+
        ' TO_DATE('''+sL_EndDate+''',''YYYY/MM/DD''))' ;

  if sL_Prefix='*'  then
    sL_SQLInvice :=  sL_SQLInvice + ' ORDER BY INVID '
  else
  begin
    sL_SQLInvice := sL_SQLInvice + ' AND SUBSTRB(INVID,1,2) = '''+ sL_Prefix +''' ORDER BY INVID';
  end;

  L_InvoiceDataSet := getDataSet(sL_SQLInvice);
// ???R?}?l
  sL_InvPrefix := '';
  nL_PrefixIndex := -1 ; //?s???e?o?????????@???r?y
  nL_TempInvNO := -1 ;
  nL_TempPrefixIndex := -1 ;
  nL_StartNum := -1;
  bL_RightFlag := true;
  with L_InvoiceDataSet do
  begin
    DisableControls;
    First;
    while not Eof do
    begin
      sL_InvID :=  FieldByName('INVID').AsString;  //AK10000001
      nL_PrefixIndex := getPrefixIndex(sL_InvID,L_Prefix,L_StartNum,L_EndNum); //???o???e?o???????????r?y
      sL_InvPrefix := copy(sL_InvID,1,2);  //???oAK
      nL_InvNO := StrToInt(copy(sL_InvID,3,8));//???o 10000001
      if nL_PrefixIndex <> nL_TempPrefixIndex then //?r?y?w?? ???o?????X?P???r?y?????l???X????
      begin

        //if nL_TempPrefixIndex <> -1 then  //?????D???????@?????r?y
        //?r?y?????}???o??
        //begin
          //if nL_TempInvNO < StrToInt(G_EndNum[nL_TempPrefixIndex]) then
          //begin
            //for L:=1 to (StrToInt(G_EndNum[nL_TempPrefixIndex])-nL_TempInvNO) do
              //frmCheck.RmoCheck.Lines.Add('?o???G'+G_Prefix.Strings[nL_TempPrefixIndex]+IntToStr(nL_TempInvNO+L)+'?|???}??')
          //end
        //end;

        nL_StartNum := StrToInt(L_StartNum[nL_PrefixIndex]);
        //bL_RightFlag := false;
        if nL_InvNO > nL_StartNum then
        begin
          bL_RightFlag := false;
          for J := 0 to (nL_InvNO - nL_StartNum)-1 do
          begin
            //frmCheck.RmoCheck.Lines.Add('?o???G'+sL_InvPrefix+IntToStr(nL_StartNum+J)+' ????')
            with cdsCheckJumpInvID do
            begin
              Append;
              FieldByName('Memo').AsString := '?o???G'+sL_InvPrefix+IntToStr(nL_StartNum+J)+' ????';
              post;
            end;
          end;

        end
        else if nL_InvNO = nL_StartNum then
        begin
          //???`
        end
        else if nL_InvNO < nL_StartNum then
        //?o?????X?p?????l?r?y?????X
        begin
          //frmCheck.RmoCheck.Lines.Add(' ?o?????X?p?????l?r?y?????X');
          //???????|?o???a?K
          with cdsCheckJumpInvID do
          begin
            Append;
            FieldByName('Memo').AsString := '?o?????X?p???_?l?r?y?????X';
            post;
          end;
        end;

        nL_TempInvNO := nL_InvNO;
        nL_TempPrefixIndex := nL_PrefixIndex
      end

      else //?r?y???? ???o?????X?P?e?@?????X????
      begin
        if nL_InvNO-nL_TempInvNO=1 then
        begin
          //???`
        end
        else if nL_InvNO-nL_TempInvNO=0 then
        begin
          //frmCheck.RmoCheck.Lines.Add('?o???G'+sL_InvID+' ????');
          with cdsCheckJumpInvID do
          begin
            Append;
            FieldByName('Memo').AsString := '?o???G'+sL_InvID+' ????';
            post;
          end;
          bL_RightFlag := false;
        end
        else if nL_InvNO - nL_TempInvNO > 1 then
        begin
          bL_RightFlag := false;
          for K := 1 to (nL_InvNO - nL_TempInvNO)-1 do
          begin
            //frmCheck.RmoCheck.Lines.Add('?o???G'+sL_InvPrefix+IntToStr(nL_TempInvNO+K)+' ????')
            with cdsCheckJumpInvID do
            begin
              Append;
              FieldByName('Memo').AsString := '?o???G'+sL_InvPrefix+IntToStr(nL_TempInvNO+K)+' ????';
              post;
            end;
          end;
        end;

        nL_TempInvNO := nL_InvNO;
      end;
      Next;

    end;
    EnableControls;
    Result := cdsCheckJumpInvID.RecordCount;
    Close;

  end;
  //?B?z?????@???r?y???}?????o??????
  //if nL_TempInvNO <> -1 then  //???B?z???o???????~??
  //begin
  //if nL_TempInvNO < StrToInt(G_EndNum[nL_TempPrefixIndex]) then
    //begin
      //for L:=1 to (StrToInt(G_EndNum[nL_TempPrefixIndex])-nL_TempInvNO) do
       // frmCheck.RmoCheck.Lines.Add('?o???G'+G_Prefix.Strings[nL_TempPrefixIndex]+IntToStr(nL_TempInvNO+L)+'?|???}??')
    //end;
  //end;
{
  if bL_RightFlag then
  begin
    //frmCheck.RmoCheck.Lines.Add('?L???????o??');
    with cdsCheckJumpInvID do
    begin
      Append;
      FieldByName('Memo').AsString := '?L???????o??';
      post;
    end;
  end;
}  

  L_Prefix.Free;
  L_StartNum.Free;
  L_EndNum.Free;
end;

function TdtmMainJ.handleChkAmountData(dI_SDate, dI_EDate: TDate;
  bI_IncludeObsolete: boolean;sI_CompID : String): Integer;
var
  aSql: String;
  nL_ID2 : Integer;
  sL_ID1, sL_InvID: String;
  nL_DetailAmt, nL_MasterAmt : Integer;
  sL_SDate, sL_EDate : String;
begin
    sL_SDate := DateToStr(dI_SDate);
    sL_EDate := DateToStr(dI_EDate);

    adoInv007AmtCheck.SQL.Clear;
    adoInv007AmtCheck.Close;

    aSql := 'select IdentifyID1, IdentifyID2, InvDate, InvID, InvAmount, IsObsolete' +
              ' from inv007 where InvDate between ' +
              'to_date(' + '''' + sL_SDate + ''''+ ',' + '''' + 'YYYY/MM/DD' + ''')' +
              ' and  to_date(' + ''''+ sL_EDate + '''' + ',' + '''' + 'YYYY/MM/DD' + ''') ' +
              ' and CompID=''' + sI_CompID + '''';

    if not bI_IncludeObsolete then
      aSql := ( aSql + ' and IsObsolete = ''N''' );

    adoInv007AmtCheck.SQL.Text := aSql;
    adoInv007AmtCheck.Open;
    if ( adoInv007AmtCheck.RecordCount > 0 ) then
    begin
      adoInv007AmtCheck.First;
      adoInv007AmtCheck.DisableControls;
      while not adoInv007AmtCheck.Eof do
      begin
        sL_ID1 := adoInv007AmtCheck.FieldByName('IdentifyID1').AsString;
        nL_ID2 := adoInv007AmtCheck.FieldByName('IdentifyID2').AsInteger;
        sL_InvID := adoInv007AmtCheck.FieldByName('InvID').AsString;
        nL_MasterAmt := adoInv007AmtCheck.FieldByName('InvAmount').AsInteger;
          
        adoInv008AmtCheck.Close;
        adoInv008AmtCheck.SQL.Clear;
        aSql := Format(
          ' SELECT SUM( TOTALAMOUNT ) AS INVAMOUNT ' +
          '   FROM INV008                          ' +
          '  WHERE IDENTIFYID1 = ''%s'' ' +
          '    AND IDENTIFYID2 = ''%s'' ' +
          '    AND INVID = ''%s''       ' +
          '  GROUP BY INVID             ',
          [sL_ID1, IntToStr( nL_ID2 ), sL_InvID] );
        adoInv008AmtCheck.SQL.Text := aSql;
        adoInv008AmtCheck.Open;

        nL_DetailAmt := 0;
        if ( adoInv008AmtCheck.RecordCount > 0 ) then
          nL_DetailAmt := adoInv008AmtCheck.FieldByName( 'InvAmount' ).AsInteger;

        if ( nL_MasterAmt <> nL_DetailAmt ) then
        begin
          cdsCheckAmount.Append;
          cdsCheckAmount.FieldByName('InvDate').AsDateTime := adoInv007AmtCheck.FieldByName('InvDate').AsDateTime;
          cdsCheckAmount.FieldByName('InvID').AsString := sL_InvID;
          cdsCheckAmount.FieldByName('MasterAmt').AsInteger := nL_MasterAmt;
          cdsCheckAmount.FieldByName('DetailAmt').AsInteger := nL_DetailAmt;
          cdsCheckAmount.FieldByName('IsObsolete').AsString := adoInv007AmtCheck.FieldByName('IsObsolete').AsString;
          cdsCheckAmount.Post;
        end;
        adoInv007AmtCheck.Next;
      end;
      adoInv007AmtCheck.EnableControls;
    end;
    Result := cdsCheckAmount.RecordCount;
end;


function TdtmMainJ.handleChkMisData(sI_SelectItem, sI_CompID, sI_SInvDate,
  sI_EInvDate: String): Integer;
var
    sL_SQL,sL_SQL1,sL_SQL2 : String;
begin
{
//?o???????H?P???A?D???m?W???P???P?b???????H???P  --->>>  SelectItem(1)
//?????T?p?X???G?p(???b??)???????T?p?X???G?p      --->>>  SelectItem(2)
//?h???????X???b?P?@?i?o???W                      --->>>  SelectItem(3)
//?o?????Y?P???A???????o?????Y???P                --->>>  SelectItem(4)
//?o???X???????s?P???A?t???????????s???P          --->>>  SelectItem(5)
}
//?o???????H?P???A?D???m?W???P???P?b???????H???P
    if sI_SelectItem = '1' then
    begin
      sL_SQL1 := 'SELECT ''?o??CUSTSNAME<>?b???????H(???b??)'' Reason,A.CUSTID MisCustID,A.CUSTNAME MisCustName,' +
                 'B.CHARGETITLE MisTitle,c.bankname,c.accountno,B.INVTITLE,B.INVNO,C.BILLNO,D.INVID,' +
                 'D.CustID InvCustID,D.CUSTSNAME,D.INVTITLE,D.BUSINESSID,c.realdate,d.invdate ' +
                 'FROM SO001 A,SO002A B,SO034 C,INV007 D ' +
                 'WHERE A.CUSTID=B.CUSTID AND A.CUSTID=C.CUSTID AND C.GUINO=D.INVID ' +
                 'AND C.SERVICETYPE IN (''I'',''D'')  AND C.ACCOUNTNO IS NOT NULL ' +
                 'AND B.BANKCODE=C.BANKCODE  AND B.ACCOUNTNO=C.ACCOUNTNO AND B.CHARGETITLE<>D.CUSTSNAME '  +
                 'AND D.INVDATE between ' + TUstr.getOracleSQLDateStr(StrToDate(sI_SInvDate)) +
                 'AND ' + TUstr.getOracleSQLDateStr(StrToDate(sI_EInvDate));


      sL_SQL2 := 'SELECT ''?o??CUSTSNAME<>?A???p???H(?L?b??))'' Reason,A.CUSTID MisCustID,A.CUSTNAME MisCustName,' +
                 'B.CONTNAME MisTitle,c.bankname,c.accountno,B.INVTITLE,B.INVNO,C.BILLNO,D.INVID,' +
                 'D.CustID InvCustID,D.CUSTSNAME,D.INVTITLE,D.BUSINESSID,c.realdate,d.invdate ' +
                 'FROM SO001 A,SO002 B,SO034 C,INV007 D ' +
                 'WHERE A.CUSTID=B.CUSTID AND A.CUSTID=C.CUSTID AND C.GUINO=D.INVID ' +
                 'AND C.SERVICETYPE IN (''I'',''D'') AND B.SERVICETYPE=C.SERVICETYPE ' +
                 'AND C.ACCOUNTNO IS  NULL AND A.CUSTNAME<>D.CUSTSNAME ' +
                 'AND D.INVDATE between ' + TUstr.getOracleSQLDateStr(StrToDate(sI_SInvDate)) +
                 'AND ' + TUstr.getOracleSQLDateStr(StrToDate(sI_EInvDate));

      sL_SQL := sL_SQL1 + ' unioN all ' + sL_SQL2;
      with adoQryCommon do
      begin
        Close;
        SQL.Clear;
        SQL.Add(sL_SQL);
        Open;


        First;
        DisableControls;
        while not Eof do
        begin
          cdsCheckMisData.Append;
          cdsCheckMisData.FieldByName('Reason').AsString := FieldByName('Reason').AsString;
          cdsCheckMisData.FieldByName('MisCustID').AsString := FieldByName('MisCustID').AsString;
          cdsCheckMisData.FieldByName('MisCustName').AsString := FieldByName('MisCustName').AsString;
          cdsCheckMisData.FieldByName('BillNo').AsString := FieldByName('BillNo').AsString;
          cdsCheckMisData.FieldByName('InvID').AsString := FieldByName('InvID').AsString;

          cdsCheckMisData.FieldByName('MisTitle').AsString := FieldByName('MisTitle').AsString;
          //cdsCheckMisData.FieldByName('MisBusinessID').AsString := '';
          cdsCheckMisData.FieldByName('InvTitle').AsString := FieldByName('InvTitle').AsString;
          cdsCheckMisData.FieldByName('InvBusinessID').AsString := FieldByName('BusinessID').AsString;
          cdsCheckMisData.FieldByName('InvCustID').AsString := FieldByName('InvCustID').AsString;

          cdsCheckMisData.Post;
          Next;
        end;
        EnableControls;

        Result := adoQryCommon.RecordCount;
        Close;
      end;
    end;


//?????T?p?X???G?p(???b??)???????T?p?X???G?p
    if sI_SelectItem = '2' then
    begin
      sL_SQL1 := 'SELECT ''LSC--?????T?p?}???G?p(???b??)'' Reason,CUSTID,TITLECUST,' +
                 'INVID,BILLNO,Realdate,INVCUSTID,CUSTSNAME,TITLEINV,INVNO,BUSINESSID FROM ' +
                 '(select a.custid,A.CHARGETITLE chargetitle,A.INVTITLE TITLECUST,b.realdate,' +
                 'C.CUSTSNAME,C.INVTITLE TITLEINV,a.accountno,a.invno,C.BUSINESSID,a.invoicetype,' +
                 'B.BILLNO,B.GUINO,B.INVDATE,B.REALAMT,C.INVID,C.CUSTID INVCUSTID,C.INVAMOUNT ' +
                 'from so002a a,so034 b, INV007 C where  b.servicetype in (''I'',''D'')  AND A.CUSTID=B.CUSTID ' +
                 'AND B.accountno IS NOT NULL  AND B.BANKCODE=A.BANKCODE AND A.ACCOUNTNO=B.ACCOUNTNO ' +
                 'AND C.INVID=B.GUINO AND C.BUSINESSID IS NULL AND A.INVNO IS NOT NULL ' +
                 'AND C.INVDATE between ' + TUstr.getOracleSQLDateStr(StrToDate(sI_SInvDate)) +
                 'AND ' + TUstr.getOracleSQLDateStr(StrToDate(sI_EInvDate)) + ')';


      sL_SQL2 := 'SELECT ''LSC--?????T?p?}???G?p(?L?b??)'' Reason,CUSTID,TITLECUST,' +
                 'INVID,BILLNO,Realdate,INVCUSTID,CUSTSNAME,TITLEINV,INVNO,BUSINESSID FROM ' +
                 '(select a.custid,A.CONTNAME chargetitle,A.INVTITLE TITLECUST,' +
                 'C.CUSTSNAME,C.INVTITLE TITLEINV,b.realdate,a.accountno,a.invno,C.BUSINESSID,a.invoicetype,' +
                 'B.BILLNO,B.GUINO,B.INVDATE,B.REALAMT,C.INVID,C.CUSTID INVCUSTID,C.INVAMOUNT' +
                 ' from so002 a,so034 b, INV007 C,so001 d ' +
                 ' where  b.servicetype in (''I'',''D'') and a.servicetype=b.servicetype and a.custid=d.custid ' +
                 'AND A.CUSTID=B.CUSTID AND B.accountno IS NULL AND C.INVID=B.GUINO AND C.BUSINESSID IS NULL AND A.INVNO IS NOT NULL ' +
                 'AND C.INVDATE between ' + TUstr.getOracleSQLDateStr(StrToDate(sI_SInvDate)) +
                 'AND ' + TUstr.getOracleSQLDateStr(StrToDate(sI_EInvDate)) + ')';


      sL_SQL := sL_SQL1 + ' unioN all ' + sL_SQL2;

      with adoQryCommon do
      begin
        Close;
        SQL.Clear;
        SQL.Add(sL_SQL);
        Open;


        First;
        DisableControls;
        while not Eof do
        begin
          cdsCheckMisData.Append;
          cdsCheckMisData.FieldByName('Reason').AsString := FieldByName('Reason').AsString;
          cdsCheckMisData.FieldByName('MisCustID').AsString := FieldByName('CUSTID').AsString;
          //cdsCheckMisData.FieldByName('MisCustName').AsString := '';
          cdsCheckMisData.FieldByName('BillNo').AsString := FieldByName('BillNo').AsString;
          cdsCheckMisData.FieldByName('InvID').AsString := FieldByName('InvID').AsString;

          cdsCheckMisData.FieldByName('MisTitle').AsString := FieldByName('TITLECUST').AsString;
          cdsCheckMisData.FieldByName('MisBusinessID').AsString := '';
          cdsCheckMisData.FieldByName('InvTitle').AsString := FieldByName('TITLEINV').AsString;
          cdsCheckMisData.FieldByName('InvBusinessID').AsString := FieldByName('BusinessID').AsString;
          cdsCheckMisData.FieldByName('InvCustID').AsString := FieldByName('InvCustID').AsString;

          cdsCheckMisData.Post;
          Next;
        end;
        EnableControls;

        Result := adoQryCommon.RecordCount;
        Close;
      end;
    end;



//?h???????X???b?P?@?i?o???W
    if sI_SelectItem = '3' then
    begin
      sL_SQL  := 'SELECT ''?h???????}???b?P?@?o???W'' Reason,A.CUSTID MisCustID,A.GUINO,A.BILLNO,A.ITEM,' +
                 'A.CITEMCODE,A.CITEMNAME,B.CUSTID InvCustID,A.REALDATE,A.MEDIABILLNO,A.ACCOUNTNO,A.invoicetime ,' +
                 'C.INVTITLE SO002TITLE,B.INVTITLE  INV007TITLE,B.InvID,B.BUSINESSID ' +
                 'FROM SO034 A,INV007 B, so002 c ' +
                 'WHERE B.REALDATE between ' + TUstr.getOracleSQLDateStr(StrToDate(sI_SInvDate)) +
                 'AND ' + TUstr.getOracleSQLDateStr(StrToDate(sI_EInvDate)) +
                 'AND A.GUINO IN (' +
                 'SELECT GUINO FROM (SELECT GUINO,CUSTID,BILLNO FROM SO034 ' +
                 'WHERE REALDATE between ' + TUstr.getOracleSQLDateStr(StrToDate(sI_SInvDate)) +
                 'AND ' + TUstr.getOracleSQLDateStr(StrToDate(sI_EInvDate)) +
                 ' GROUP BY GUINO,CUSTID,BILLNO' +
                 ') QQ GROUP BY GUINO HAVING COUNT(*) > 1) AND A.GUINO=B.INVID ' +
                 'AND A.CUSTID<>to_number(B.custid) and a.custid=c.custid and a.servicetype=''C'' AND C.SERVICETYPE=''C''' +
                 'ORDER BY A.GUINO';


      with adoQryCommon do
      begin
        Close;
        SQL.Clear;
        SQL.Add(sL_SQL);
        Open;


        First;
        DisableControls;
        while not Eof do
        begin
          cdsCheckMisData.Append;
          cdsCheckMisData.FieldByName('Reason').AsString := FieldByName('Reason').AsString;
          cdsCheckMisData.FieldByName('MisCustID').AsString := FieldByName('MisCustID').AsString;
          //cdsCheckMisData.FieldByName('MisCustName').AsString := '';
          cdsCheckMisData.FieldByName('BillNo').AsString := FieldByName('BillNo').AsString;
          cdsCheckMisData.FieldByName('InvID').AsString := FieldByName('InvID').AsString;

          cdsCheckMisData.FieldByName('MisTitle').AsString := FieldByName('SO002TITLE').AsString;
          //cdsCheckMisData.FieldByName('MisBusinessID').AsString := '';
          cdsCheckMisData.FieldByName('InvTitle').AsString := FieldByName('INV007TITLE').AsString;
          cdsCheckMisData.FieldByName('InvBusinessID').AsString := FieldByName('BusinessID').AsString;
          cdsCheckMisData.FieldByName('InvCustID').AsString := FieldByName('InvCustID').AsString;

          cdsCheckMisData.Post;
          Next;
        end;
        EnableControls;

        Result := adoQryCommon.RecordCount;
        Close;
      end;
    end;



//?o?????Y?P???A???????o?????Y???P
    if sI_SelectItem = '4' then
    begin
      sL_SQL1 := 'SELECT ''?L?b??--?o?????Y???P'' Reason,A.CUSTID MisCustID,B.CUSTNAME MisCustName,C.BILLNO,C.GUINO,C.REALDATE,' +
                 'A.INVTITLE MisTitle,D.INVTITLE InvTitle,A.INVNO,D.BUSINESSID,D.InvID,D.CUSTID InvCustID ' +
                 'FROM SO002 A,SO001 B,SO034 C,INV007 D ' +
                 'WHERE A.CUSTID=B.CUSTID AND A.CUSTID=C.CUSTID' +
                 ' AND A.SERVICETYPE=C.SERVICETYPE AND A.SERVICETYPE IN (''I'',''D'')' +
                 ' AND C.GUINO=D.INVID AND C.ACCOUNTNO IS NULL ' +
                 ' AND D.INVDATE between ' + TUstr.getOracleSQLDateStr(StrToDate(sI_SInvDate)) +
                 ' AND ' + TUstr.getOracleSQLDateStr(StrToDate(sI_EInvDate)) +
                 ' AND A.INVOICETYPE=3 AND A.INVTITLE<>D.INVTITLE';

      sL_SQL2 := 'SELECT ''???b??--?o?????Y???P'' Reason,A.CUSTID MisCustID,B.CUSTNAME MisCustName,C.BILLNO,C.GUINO,C.REALDATE,' +
                 'A.INVTITLE MisTitle,D.INVTITLE InvTitle,A.INVNO,D.BUSINESSID,D.InvID,D.CUSTID InvCustID ' +
                 'FROM SO002A A,SO001 B,SO034 C,INV007 D' +
                 ' WHERE A.CUSTID=B.CUSTID AND A.CUSTID=C.CUSTID ' +
                 ' AND C.SERVICETYPE IN (''I'',''D'') AND A.BANKCODE IS NOT NULL ' +
                 ' AND A.BANKCODE=C.BANKCODE AND A.ACCOUNTNO=C.ACCOUNTNO ' +
                 ' AND C.GUINO=D.INVID ' +
                 ' AND D.INVDATE between ' + TUstr.getOracleSQLDateStr(StrToDate(sI_SInvDate)) +
                 ' AND ' + TUstr.getOracleSQLDateStr(StrToDate(sI_EInvDate)) +
                 ' AND A.INVOICETYPE=3 AND A.INVTITLE<>D.INVTITLE';

      sL_SQL := sL_SQL1 + ' unioN all ' + sL_SQL2;

      with adoQryCommon do
      begin
        Close;
        SQL.Clear;
        SQL.Add(sL_SQL);
        Open;


        First;
        DisableControls;
        while not Eof do
        begin
          cdsCheckMisData.Append;
          cdsCheckMisData.FieldByName('Reason').AsString := FieldByName('Reason').AsString;
          cdsCheckMisData.FieldByName('MisCustID').AsString := FieldByName('MisCustID').AsString;
          cdsCheckMisData.FieldByName('MisCustName').AsString := FieldByName('MisCustName').AsString;
          cdsCheckMisData.FieldByName('BillNo').AsString := FieldByName('BillNo').AsString;
          cdsCheckMisData.FieldByName('InvID').AsString := FieldByName('InvID').AsString;

          cdsCheckMisData.FieldByName('MisTitle').AsString := FieldByName('MisTitle').AsString;
          //cdsCheckMisData.FieldByName('MisBusinessID').AsString := '';
          cdsCheckMisData.FieldByName('InvTitle').AsString := FieldByName('InvTitle').AsString;
          cdsCheckMisData.FieldByName('InvBusinessID').AsString := FieldByName('BusinessID').AsString;
          cdsCheckMisData.FieldByName('InvCustID').AsString := FieldByName('InvCustID').AsString;

          cdsCheckMisData.Post;
          Next;
        end;
        EnableControls;

        Result := adoQryCommon.RecordCount;
        Close;
      end;
    end;


//?o???X???????s?P???A?t???????????s???P
    if sI_SelectItem = '5' then
    begin
      sL_SQL1 := 'SELECT ''?L?b??--???s???~'' Reason,A.CUSTID MisCustID,B.CUSTNAME MisCustName,C.BILLNO,C.GUINO,C.REALDATE,' +
                 'A.INVTITLE MisTitle,D.INVTITLE InvTitle,A.INVNO MisBusinessID,D.BUSINESSID,D.InvID,D.CUSTID InvCustID ' +
                 'FROM SO002 A,SO001 B,SO034 C,INV007 D ' +
                 'WHERE A.CUSTID=B.CUSTID AND A.CUSTID=C.CUSTID' +
                 ' AND A.SERVICETYPE=C.SERVICETYPE AND A.SERVICETYPE IN (''I'',''D'')' +
                 ' AND C.GUINO=D.INVID AND C.ACCOUNTNO IS NULL' +
                 ' AND A.INVOICETYPE=3 AND A.INVNO<>D.BUSINESSID' +
                 ' AND D.INVDATE between ' + TUstr.getOracleSQLDateStr(StrToDate(sI_SInvDate)) +
                 ' AND ' + TUstr.getOracleSQLDateStr(StrToDate(sI_EInvDate));


      sL_SQL2 := 'SELECT ''???b??--???s???~'' Reason,A.CUSTID MisCustID,B.CUSTNAME MisCustName,C.BILLNO,C.GUINO,C.REALDATE,' +
                 'A.INVTITLE MisTitle,D.INVTITLE InvTitle,A.INVNO MisBusinessID,D.BUSINESSID,D.InvID,D.CUSTID InvCustID ' +
                 'FROM SO002A A,SO001 B,SO034 C,INV007 D ' +
                 'WHERE A.CUSTID=B.CUSTID AND A.CUSTID=C.CUSTID ' +
                 ' AND C.SERVICETYPE IN (''I'',''D'') AND A.BANKCODE IS NOT NULL' +
                 ' AND A.BANKCODE=C.BANKCODE AND A.ACCOUNTNO=C.ACCOUNTNO' +
                 ' AND C.GUINO=D.INVID AND A.INVOICETYPE=3 AND  A.INVNO<>D.BUSINESSID ' +
                 ' AND D.INVDATE between ' + TUstr.getOracleSQLDateStr(StrToDate(sI_SInvDate)) +
                 ' AND ' + TUstr.getOracleSQLDateStr(StrToDate(sI_EInvDate));


      sL_SQL := sL_SQL1 + ' unioN all ' + sL_SQL2;

      with adoQryCommon do
      begin
        Close;
        SQL.Clear;
        SQL.Add(sL_SQL);
        Open;


        First;
        DisableControls;
        while not Eof do
        begin
          cdsCheckMisData.Append;
          cdsCheckMisData.FieldByName('Reason').AsString := FieldByName('Reason').AsString;
          cdsCheckMisData.FieldByName('MisCustID').AsString := FieldByName('MisCustID').AsString;
          cdsCheckMisData.FieldByName('MisCustName').AsString := FieldByName('MisCustName').AsString;
          cdsCheckMisData.FieldByName('BillNo').AsString := FieldByName('BillNo').AsString;
          cdsCheckMisData.FieldByName('InvID').AsString := FieldByName('InvID').AsString;

          cdsCheckMisData.FieldByName('MisTitle').AsString := FieldByName('MisTitle').AsString;
          cdsCheckMisData.FieldByName('MisBusinessID').AsString := FieldByName('MisBusinessID').AsString;
          cdsCheckMisData.FieldByName('InvTitle').AsString := FieldByName('InvTitle').AsString;
          cdsCheckMisData.FieldByName('InvBusinessID').AsString := FieldByName('BusinessID').AsString;
          cdsCheckMisData.FieldByName('InvCustID').AsString := FieldByName('InvCustID').AsString;

          cdsCheckMisData.Post;
          Next;
        end;
        EnableControls;

        Result := adoQryCommon.RecordCount;
        Close;
      end;      
    end;
end;

function TdtmMainJ.getAllInv004Data: Integer;
var
    sL_SQL : String;
begin
    sL_SQL := 'select * from inv004';

    with adoInv004Code do
    begin
      Close;
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;

      Result := RecordCount;
    end;
end;

//function TdtmMainJ.getCreateTypeInfo(sI_CompID, sI_SInvDate, sI_EInvDate,
//  sI_HowToCreate,sI_IsObsolete: String): Integer;
//var
//    sL_SQL,sL_SQLTitle,sL_Where : String;
//    sL_QueryInvDate,sL_QueryHowToCreate,sL_QueryIsObsolete,sL_InvAmount : String;
//begin
//    if not cdsCreateTypeInfo.Active then
//      cdsCreateTypeInfo.CreateDataSet;
//
//    cdsCreateTypeInfo.EmptyDataSet;
//
//    sL_SQLTitle := 'SELECT IsObsolete,InvDate,HowToCreate,SUM(INVAMOUNT) TotalAmount FROM INV007';
//    sL_Where := ' where IdentifyId1=''' + IDENTIFYID1 + ''' and IdentifyId2=' + IDENTIFYID2 +
//                ' and CompID=''' + sI_CompID + '''';
//
//    sL_Where := sL_Where + ' and INVDATE between ' +
//                TUstr.getOracleSQLDateStr(StrToDate(sI_SInvDate)) + ' and ' +
//                TUstr.getOracleSQLDateStr(StrToDate(sI_EInvDate));
//
//    if sI_HowToCreate <> '' then
//      sL_Where := sL_Where + ' and HowToCreate =''' + sI_HowToCreate + '''';
//
//
//    if sI_IsObsolete <> '' then
//      sL_Where := sL_Where + ' and ISOBSOLETE =''' + sI_IsObsolete + '''';
//
//
//
//    sL_SQL := sL_SQLTitle + sL_Where + ' GROUP BY IsObsolete,InvDate,HowToCreate';
//
//
//    with adoQryCommon do
//    begin
//      Close;
//      SQL.Clear;
//      SQL.Add(sL_SQL);
//      Open;
//      with cdsCreateTypeInfo do
//      begin
//        adoQryCommon.First;
//        adoQryCommon.DisableControls;
//        while not adoQryCommon.Eof do
//        begin
//          sL_QueryInvDate := adoQryCommon.FieldByName('InvDate').AsString;
//          sL_QueryHowToCreate := adoQryCommon.FieldByName('HowToCreate').AsString;
//          sL_QueryIsObsolete := adoQryCommon.FieldByName('IsObsolete').AsString;
//          sL_InvAmount := adoQryCommon.FieldByName('TotalAmount').AsString;
//
//
//          //1-mis????, ?w?}
//          //2-mis????, ???}
//          //3-invoice create, ?w?}
//          //4-invoice create, ???}
//          if cdsCreateTypeInfo.Locate('InvDate;IsObsolete',VarArrayOf([sL_QueryInvDate,sL_QueryIsObsolete]) , []) then
//            Edit
//          else
//          begin
//            Append;
//            FieldByName('InvDate').AsString := sL_QueryInvDate;
//            FieldByName('IsObsolete').AsString := sL_QueryIsObsolete;
////=======================================================2004/11 Julien??
//            FieldByName('CreateType1').AsString := '0';
//            FieldByName('CreateType2').AsString := '0';
//            FieldByName('CreateType3').AsString := '0';
//            FieldByName('CreateType4').AsString := '0';
////=======================================================2004/11 Julien??
//          end;
//
//          if sL_QueryHowToCreate = '1' then
//            FieldByName('CreateType1').AsString := sL_InvAmount
//          else if sL_QueryHowToCreate = '2' then
//            FieldByName('CreateType2').AsString := sL_InvAmount
//          else if sL_QueryHowToCreate = '3' then
//            FieldByName('CreateType3').AsString := sL_InvAmount
//          else if sL_QueryHowToCreate = '4' then
//            FieldByName('CreateType4').AsString := sL_InvAmount;
//
//          FieldByName('RecordTotal').AsInteger :=  //2004/11 Julien
//            FieldByName('CreateType1').AsInteger + FieldByName('CreateType2').AsInteger +
//            FieldByName('CreateType3').AsInteger + FieldByName('CreateType4').AsInteger ;
//
//          Post;
//          adoQryCommon.Next;
//        end;
//        adoQryCommon.EnableControls;
//      end;
//    end;
//
//    Result := cdsCreateTypeInfo.RecordCount;
//end;

{ ---------------------------------------------------------------------------- }

function TdtmMainJ.getObsoleteInvoiceData(sI_CompID, sI_SInvID, sI_EInvID,
  sI_SInvDate, sI_EInvDate, sI_ObsoleteId: String): Integer;
var
  aHowToCreate, aSubSql : String;
begin
  if not cdsObsoleteInvoiceData.Active then
    cdsObsoleteInvoiceData.CreateDataSet;
  cdsObsoleteInvoiceData.EmptyDataSet;
  {}
  aSubSql := Format(
    '     SELECT DISTINCT A.MAININVID FROM INV007 A          ' +
    '      WHERE A.IDENTIFYID1 = ''%s''                      ' +
    '        AND A.IDENTIFYID2 = ''%s''                      ' +
    '        AND A.COMPID = ''%s''                           ' +
    '        AND A.ISOBSOLETE=''Y''                          ',
    [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID] );
  {}
  if ( sI_SInvID <> EmptyStr ) then
  begin
    aSubSql := aSubSql +
      ' AND A.INVID BETWEEN ''' +  sI_SInvID + ''' AND ''' + sI_EInvID + '''';
  end;
  {}
  if ( sI_SInvDate <> EmptyStr ) then
  begin
    aSubSql := aSubSql +
      ' AND A.INVDATE BETWEEN ' + TUstr.getOracleSQLDateStr(StrToDate(sI_SInvDate)) +
      ' AND ' + TUstr.getOracleSQLDateStr(StrToDate(sI_EInvDate));
  end;

  if ( sI_ObsoleteId <> EmptyStr ) then
  begin
    aSubSql := aSubSql + ' AND OBSOLETEID=''' + sI_ObsoleteId + '''';
  end;

  adoQryCommon.Close;
  adoQryCommon.SQL.Text := Format(
    ' SELECT A.INVID, A.INVDATE, A.INVAMOUNT,                ' +
    '        A.OBSOLETEREASON, A.CUSTID, A.HOWTOCREATE       ' +
    '   FROM INV007 A, ( %s )  B                             ' +
    '  WHERE A.MAININVID = B.MAININVID                       ' +
    '    AND A.IDENTIFYID1 = ''%s''                          ' +
    '    AND A.IDENTIFYID2 = ''%s'' AND A.COMPID = ''%s''    ' +
    '  ORDER BY A.INVID                                      ',
    [aSubSql, IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID] );
  adoQryCommon.Open;
  adoQryCommon.First;
  while not adoQryCommon.Eof do
  begin
    cdsObsoleteInvoiceData.Append;
    cdsObsoleteInvoiceData.FieldByName('INVID').AsString :=
      adoQryCommon.FieldByName('INVID').AsString;
    {}
    cdsObsoleteInvoiceData.FieldByName('INVDATE').AsString :=
      adoQryCommon.FieldByName('INVDATE').AsString;
    {}
    cdsObsoleteInvoiceData.FieldByName('INVAMOUNT').AsString :=
      adoQryCommon.FieldByName('INVAMOUNT').AsString;
    {}
    cdsObsoleteInvoiceData.FieldByName('OBSOLETEREASON').AsString :=
      adoQryCommon.FieldByName('OBSOLETEREASON').AsString;
    {}
    cdsObsoleteInvoiceData.FieldByName('CUSTID').AsString :=
      adoQryCommon.FieldByName('CUSTID').AsString;

    //0-????
    //1-mis????, ?w?}
    //2-mis????, ???}
    //3-invoice create, ?w?}
    //4-invoice create, ???}
    {}
    aHowToCreate := EmptyStr;
    if adoQryCommon.FieldByName('HowToCreate').AsString = '1' then
      aHowToCreate := '?w?}'
    else if adoQryCommon.FieldByName('HowToCreate').AsString = '2' then
      aHowToCreate := '???}'
    else if adoQryCommon.FieldByName('HowToCreate').AsString = '3' then
      aHowToCreate := '?{???}??'
    else if adoQryCommon.FieldByName('HowToCreate').AsString = '4' then
      aHowToCreate := '?@???}??';
    {}
    cdsObsoleteInvoiceData.FieldByName('HowToCreate').AsString := aHowToCreate;
    cdsObsoleteInvoiceData.Post;
    adoQryCommon.Next;
  end;
  Result := cdsObsoleteInvoiceData.RecordCount;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMainJ.CanDropOrRecoverInv(var aParam: TCommonParam): Boolean;
var
  aYearMonth: String;
  aSoReader: TADOQuery;
  aCount: Integer;
begin
  { aParam.InStr1 = ?o?????X
    aParam.InStr2 = Y/N ?? ???r?? --> ?@?o/?^?_
    aParam.OutStr1 = ?o??????
    aParam.OutStr2 = ???o???O?_?w?Q?@?o --> Y or N }
  Result := False;
  adoQryCommon.Close;
  adoQryCommon.SQL.Text := Format(
     ' SELECT INVDATE, ISOBSOLETE            ' +
     '   FROM INV007                         ' +
     '  WHERE IDENTIFYID1 = ''%s''           ' +
     '    AND IDENTIFYID2 = ''%s''           ' +
     '    AND COMPID = ''%s''                ' +
     '    AND INVID = ''%s''                 ',
     [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, aParam.InStr1] );
  adoQryCommon.Open;
  if ( not adoQryCommon.IsEmpty ) then
  begin
    aParam.OutStr1 := adoQryCommon.FieldByName( 'INVDATE' ).AsString;
    aParam.OutStr2 := adoQryCommon.FieldByName( 'ISOBSOLETE' ).AsString;
  end;
  adoQryCommon.Close;
  aParam.Msg := EmptyStr;
  try
    if ( aParam.OutStr1 = EmptyStr ) then
    begin
      aParam.Msg := Format( '?o???????d?L???i?o??:%s?C', [aParam.OutStr1] );
      Exit;
    end;
    if ( aParam.InStr2 <> EmptyStr ) and ( aParam.OutStr2 = 'Y' ) then
    begin
      aParam.Msg := Format( '???i?o??:%s, ?w?Q?@?o,???i?A?Q?@?o?C',
        [aParam.InStr1] );
      Exit;
    end else
    if ( aParam.InStr2 = EmptyStr ) and ( Nvl( aParam.OutStr2, 'N' ) = 'N' ) then
    begin
      aParam.Msg := Format( '???i?o??:%s, ?|???@?o,???i?^?_?C',
        [aParam.InStr1] );
      Exit;
    end;
//    ?H?U????????, ?S?H?o?????????D, ???????A?g?{?????W
//    if ( aParam.InStr2 = EmptyStr ) and ( aParam.OutStr2 = 'Y' ) then
//    begin
//      //???????n?^?_???o?????????O?_?w?g?S?q???r?????i??, ???}?o??,
//      //?Y???d????????????????, ???i?????\?^?_
//      if CreatrInvAgain() then
//        aParam.Msg := '???i?o???????w???s?}???L?o??, ?????\?^?_?C';
//      Exit
//    end;
    aYearMonth := Copy( aParam.OutStr1, 1, 4 ) + Copy( aParam.OutStr1, 6, 2 );
    if InvHasBeenLock( aYearMonth ) then
    begin
      aParam.Msg := Format( '???i?o???????~??: %s, ?w?g???b,?L?k?????C',
        [FormatMaskText( '####/##;0;_', aYearMonth )] );
      Exit;
    end;
    {}
    { ?@?o??, ?????????o???O?_?w????????, aParam.InStr2 = Y ?N???@?o }
    if ( aParam.InStr2 <> EmptyStr )  then
    begin
      if ( dtmMain.GetLinkToMIS ) then
      begin
        aSoReader := dtmSO.adoComm;
        aSoReader.Close;
        aSoReader.SQL.Text := Format(
          ' select count(1) from so034  ' +
          '  where guino = ''%s''       ' +
          '    and preinvoice = 3       ', [aParam.InStr1] );
        aSoReader.Open;
        aCount := aSoReader.Fields[0].AsInteger;
        aSoReader.Close;
        if ( aCount > 0 ) then
        begin
          aParam.Msg := Format( '???i?o??:%s, ???A?t???????????i???n???????????j, ???i?@?o?C',
            [aParam.InStr1] );
          Exit;
        end;
        {}
        aSoReader.Close;
        aSoReader.SQL.Text := Format(
          ' select count(1) from so034  ' +
          '  where guino = ''%s''       ' +
          '    and preinvoice = 4       ', [aParam.InStr1] );
        aSoReader.Open;
        aCount := aSoReader.Fields[0].AsInteger;
        aSoReader.Close;
        if ( aCount > 0 ) then
        begin
          aParam.Msg := Format( '???i?o??:%s, ???A?t???????????i?w?g???????????j, ???i?@?o?C',
            [aParam.InStr1] );
          Exit;
        end;
      end;
      { ?????d?o???t?????????????? }
      adoQryCommon.Close;
      adoQryCommon.SQL.Text := Format(
        ' select count(1) from inv014    ' +
        '  where identifyid1 = ''%s''    ' +
        '    and identifyid2 = ''%s''    ' +
        '    and compid = ''%s''         ' +
        '    and invid = ''%s''          ',
        [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, aParam.InStr1] );
      adoQryCommon.Open;
      aCount := adoQryCommon.Fields[0].AsInteger;
      adoQryCommon.Close;
      if ( aCount > 0 ) then
      begin
        aParam.Msg := Format( '???i?o??:%s, ?o???t?????i?w?g???????????j, ???i?@?o?C',
          [aParam.InStr1] );
        Exit;
      end;
    end;
  finally
    Result := ( aParam.Msg = EmptyStr );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMainJ.DropOrRecoverInv(var aParam: TCommonParam): Boolean;
var
  aSql: String;
begin
  { aParam.InStr1 = ?o?????X
    aParam.InStr2 = ???o?N?X
    aParam.InStr3 = ???o???]
    aParam.OutStr1 = HowToCreate
    aParam.OutStr2 = ?o??????
    }
  adoQryCommon.Close;
  adoQryCommon.SQL.Text := Format(
    '  SELECT HOWTOCREATE,                ' +
    '         INVDATE                     ' +
    '    FROM INV007                      ' +
    '   WHERE IDENTIFYID1 = ''%s''        ' +
    '     AND IDENTIFYID2 = %s            ' +
    '     AND COMPID = ''%s''             ' +
    '     AND INVID = ''%s''              ',
    [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, aParam.InStr1] );
  adoQryCommon.Open;
  aParam.OutStr1 := adoQryCommon.FieldByName( 'HOWTOCREATE' ).AsString;
  aParam.OutStr2 := adoQryCommon.FieldByName( 'INVDATE' ).AsString;
  adoQryCommon.Close;
  { ?^?_ }
  if ( aParam.InStr2 = EmptyStr ) then
  begin
    aSql := Format(
     '  UPDATE INV007                      ' +
     '     SET ISOBSOLETE = ''N'',         ' +
     '         OBSOLETEID = NULL,          ' +
     '         OBSOLETEREASON = NULL,      ' +
     '         CANMODIFY = ''Y'',          ' +
     '         UPTTIME = SYSDATE,          ' +
     '         UPTEN = ''%s''              ' +
     '   WHERE IDENTIFYID1 = ''%s''        ' +
     '     AND IDENTIFYID2 = %s            ' +
     '     AND COMPID = ''%s''             ' +
     '     AND INVID = ''%s''              ',
     [dtmMain.getLoginUser, IDENTIFYID1, IDENTIFYID2,
      dtmMain.getCompID, aParam.InStr1] );
  end else
  begin
    aSql := Format(
     '  UPDATE INV007                      ' +
     '     SET ISOBSOLETE = ''Y'',         ' +
     '         OBSOLETEID = ''%s'',        ' +
     '         OBSOLETEREASON = ''%s'',    ' +
     '         CANMODIFY = ''N'',          ' +
     '         UPTTIME = SYSDATE,          ' +
     '         UPTEN = ''%s''              ' +
     '   WHERE IDENTIFYID1 = ''%s''        ' +
     '     AND IDENTIFYID2 = ''%s''        ' +
     '     AND COMPID = ''%s''             ' +
     '     AND INVID = ''%s''              ',
     [ aParam.InStr2, aParam.InStr3, dtmMain.getLoginUser,
       IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, aParam.InStr1] );
  end;
  adoQryCommon.Close;
  adoQryCommon.SQL.Text := aSql;
  adoQryCommon.ExecSQL;
  { ???@?o???????????g?J?o???@?o?????O????, ?????r?????J?o???t????,
    ?Y?O???J?????????d?? Table ??????????, ?????i?H?A?????J }
  { ?^?_????, ?R?? INV024 }
  if ( aParam.InStr2 = EmptyStr ) then
  begin
    adoQryCommon.Close;
    adoQryCommon.SQL.Text := Format(
      ' DELETE FROM INV024 WHERE COMPID = ''%s''           ' +
      '    AND INVID = ''%s''                              ',
      [ dtmMain.getCompID, aParam.InStr1] );
    adoQryCommon.ExecSQL;
  end else
  begin
    { ???o, ?g?J INV024, ?????? INV008A + INV008, ?o???~?O?u?????????i????????,
      ?]?? INV008 ???????i???O?h???X?????@??, ?]???i???O?S???X???? --> INV008 }
    {#5363 ???i???S??????CC&B???HBillID?PBillIDitemNo?p?G?ONull,?n???@?????????? By Kin 2009/10/27}
    adoQryCommon.Close;
    adoQryCommon.SQL.Text := Format(
      ' INSERT INTO INV024 (                                        ' +
      '    COMPID, INVID, BILLID, BILLIDITEMNO, CUSTID )            ' +
      '  SELECT DISTINCT A.COMPID, A.INVID,                         ' +
      '  DECODE(B.BILLID,NULL,'' '',B.BILLID) BILLID,               ' +
      '  DECODE(B.BILLIDITEMNO,NULL,0,B.BILLIDITEMNO),              ' +
      '         A.CUSTID                                            ' +
      '    FROM INV007 A, INV008A  B                                ' +
      '   WHERE A.INVID = B.INVID                                   ' +
      '     AND A.ISOBSOLETE = ''Y''                                ' +
      '     AND B.BILLID IS NOT NULL                                ' +
      '     AND B.BILLIDITEMNO IS NOT NULL                          ' +
      '     AND A.IDENTIFYID1 = ''%s''                              ' +
      '     AND A.IDENTIFYID2 = ''%s''                              ' +
      '     AND A.COMPID = ''%s''                                   ' +
      '     AND A.INVID = ''%s''                                    ',
      [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, aParam.InStr1] );
    adoQryCommon.ExecSQL;
    {}
    {#5363 ???i???S??????CC&B???HBillID?PBillIDitemNo?p?G?ONull,?n???@?????????? By Kin 2009/10/27}
    adoQryCommon.SQL.Text := Format(
      ' INSERT INTO INV024 (                                        ' +
      '    COMPID, INVID, BILLID, BILLIDITEMNO, CUSTID )            ' +
      '  SELECT DISTINCT A.COMPID, A.INVID,                         ' +
      '  DECODE(B.BILLID,NULL,'' '',B.BILLID) BILLID,               ' +
      '  DECODE(B.BILLIDITEMNO,NULL,0,B.BILLIDITEMNO),              ' +
      '         A.CUSTID                                            ' +
      '    FROM INV007 A, INV008 B, INV008A C                       ' +
      '   WHERE A.INVID = B.INVID                                   ' +
      '     AND B.INVID = C.INVID(+)                                ' +
      '     AND B.BILLID = C.BILLID(+)                              ' +
      '     AND B.BILLIDITEMNO = C.BILLIDITEMNO(+)                  ' +
      '     AND A.ISOBSOLETE = ''Y''                                ' +
      '     AND A.IDENTIFYID1 = ''%s''                              ' +
      '     AND A.IDENTIFYID2 = ''%s''                              ' +
      '     AND A.COMPID = ''%s''                                   ' +
      '     AND A.INVID = ''%s''                                    ' +
      '     AND C.INVID IS NULL                                     ',
      [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, aParam.InStr1] );
    adoQryCommon.ExecSQL;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMainJ.InvHasBeenLock(const aInvYearMonth: String): Boolean;
begin
  adoQryCommon.Close;
  adoQryCommon.SQL.Text := Format(
    '  SELECT ISLOCKED              ' +
    '    FROM INV018                ' +
    '   WHERE IDENTIFYID1 = ''%s''  ' +
    '     AND IDENTIFYID2 = %s      ' +
    '     AND COMPID = ''%s''       ' +
    '     AND YEARMONTH = ''%s''    ',
   [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, aInvYearMonth] );
   adoQryCommon.Open;
   Result := ( adoQryCommon.FieldByName( 'ISLOCKED' ).AsString = 'Y' );
   adoQryCommon.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMainJ.UpdateObsoleteMisData(const aInvId: String);
begin
  dtmSO.adoComm.Close;
  dtmSO.adoComm.SQL.Text := Format(
       '   UPDATE %s.SO033                      ' +
       '      SET PREINVOICE = NULL,            ' +
       '          GUINO = NULL,                 ' +
       '          INVDATE = NULL,               ' +
       '          INVOICETIME = NULL            ' +
       '    WHERE GUINO = ''%s''                ',
       [dtmMain.getMisDbOwner, aInvId] );
   dtmSO.adoComm.ExecSQL;
   {}
   dtmSO.adoComm.SQL.Text := Format(
       '   UPDATE %s.SO034                      ' +
       '      SET PREINVOICE = NULL,            ' +
       '          GUINO = NULL,                 ' +
       '          INVDATE = NULL,               ' +
       '          INVOICETIME = NULL            ' +
       '    WHERE GUINO = ''%s''                ',
       [dtmMain.getMisDbOwner, aInvId] );
   dtmSO.adoComm.ExecSQL;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMainJ.UpdateRecoverMisData(const aInvId, aInvDate, aPreInvoice: String);
begin
  dtmSO.adoComm.Close;
  dtmSO.adoComm.SQL.Text := Format(
    '   UPDATE %s.SO033                                       ' +
    '      SET PREINVOICE = %s,                               ' +
    '          GUINO = ''%s'',                                ' +
    '          INVDATE = TO_DATE( ''%s'', ''YYYY/MM/DD'' ),   ' +
    '          INVOICETIME = SYSDATE                          ' +
    '    WHERE GUINO = ''%s''                                 ',
    [dtmMain.getMisDbOwner, aPreInvoice, aInvId, aInvDate, aInvId] );
  dtmSO.adoComm.ExecSQL;
   {}
  dtmSO.adoComm.SQL.Text := Format(
    '   UPDATE %s.SO034                                       ' +
    '      SET PREINVOICE = %s,                               ' +
    '          GUINO = ''%s'',                                ' +
    '          INVDATE = TO_DATE( ''%s'', ''YYYY/MM/DD'' ),   ' +
    '          INVOICETIME = SYSDATE                          ' +
    '    WHERE GUINO = ''%s''                                 ',
    [dtmMain.getMisDbOwner, aPreInvoice, aInvId, aInvDate, aInvId] );
   dtmSO.adoComm.ExecSQL;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMainJ.getInv018Data(aCompID, aYear: String): Integer;
begin
  if aYear = EmptyStr then
    aYear := '%'
  else
    aYear := ( aYear + '%' );
  {}
  adoQryInv018.Close;
  adoQryInv018.SQL.Text := Format(
    ' SELECT * FROM INV018          ' +
    '  WHERE IDENTIFYID1= ''%s''    ' +
    '    AND IDENTIFYID2= ''%s''    ' +
    '    AND COMPID = ''%s''        ' +
    '    AND YEARMONTH LIKE ''%s''  ' +
    '  ORDER BY YEARMONTH           ',
    [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, aYear] );
  adoQryInv018.Open;
  Result := adoQryInv018.RecordCount;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMainJ.handleIsLockedData(sI_CompID, sI_IsLockedYear,
  sI_IsLockedMonth, sI_IsLocked: String);
var
    sL_SQL,sL_CanModify,sL_Useful,sL_SInvDate,sL_EInvDate,sL_YM,sL_NextYM,sL_NextYMD : String;
begin
    if sI_IsLocked = 'N' then   //?????b
    begin
      sL_CanModify := 'Y';
      sL_Useful := 'Y';
    end
    else                        //???b
    begin
      sL_CanModify := 'N';
      sL_Useful := 'N';
    end;

    sL_YM := sI_IsLockedYear + Format('%.2d',[StrToInt(sI_IsLockedMonth)]);

    sL_SInvDate := sI_IsLockedYear + '/' + Format('%.2d',[StrToInt(sI_IsLockedMonth)]) + '/01';

    sL_NextYM := afterMonth(sL_YM,1);
    sL_NextYMD := Copy(sL_NextYM,0,4) + '/' + Copy(sL_NextYM,5,2) + '/01';
    sL_EInvDate := Copy(sL_NextYM,0,4) + '/' + Copy(sL_NextYM,5,2) + '/' + IntToStr(TUdateTime.DaysOfMonth(StrToDateTime(sL_NextYMD)));

    //update?r?y??
    sL_SQL := 'update INV099 set USEFUL=''' + sL_Useful +
              ''' where IDENTIFYID1=''' + IDENTIFYID1 + ''' and IDENTIFYID2=' + IDENTIFYID2 +
              ' and COMPID=''' + sI_CompID + ''' and  YEARMONTH =''' + sL_YM + '''';

    executeSQL(sL_SQL);


    if sI_IsLocked = 'Y' then //???b???~????inv007????,?Y?O???}???b?h?o?????????t?????v?~?i????
    begin
      //update?o???D??
      sL_SQL := 'update INV007 set CANMODIFY=''' + sL_CanModify +
                ''' where IDENTIFYID1=''' + IDENTIFYID1 + ''' and IDENTIFYID2=' + IDENTIFYID2 +
                ' and COMPID=''' + sI_CompID + ''' and  INVDATE between ' +
                TUstr.getOracleSQLDateStr(StrToDate(sL_SInvDate)) +
                ' and ' + TUstr.getOracleSQLDateStr(StrToDate(sL_EInvDate));

      executeSQL(sL_SQL);
    end;
end;

function TdtmMainJ.afterMonth(sI_YM: String; nI_Counts: Integer): String;
var
    nL_Year,nL_Month,ii : Integer;
    sL_Year,sL_Month,sL_NextYM : String;
begin
    //afterMonth('09109',5) --->>> 9202
    nL_Year := StrToInt(Copy(sI_YM,1,3));
    nL_Month := StrToInt(Copy(sI_YM,4,5));

    for ii:=1 to nI_Counts do
    begin
      if nL_Month = 12 then
      begin
        nL_Year := nL_Year + 1;
        nL_Month := 1;
      end
      else
      begin
        nL_Month := nL_Month + 1;
      end;
    end;

    sL_Year := IntToStr(nL_Year);
    sL_Month := IntToStr(nL_Month);
    if Length(sL_Month) < 2 then
      sL_Month := '0' + sL_Month;

    sL_NextYM := sL_Year + sL_Month;

    Result := sL_NextYM;
end;


{ ---------------------------------------------------------------------------- }

function TdtmMainJ.AuthorizeInv(const aInvIdSt, aInvIdEd, aInvUseId: String;
  const AIsInvUse: Integer;const aInvKind: Integer;const aInvs:String): String;

  { -------------------------------- }

  procedure ClearTempDataSet;
  begin
    if not VarIsNull( cdsTemp.Data ) then
      cdsTemp.EmptyDataSet;
    cdsTemp.Data := Null;
  end;

  { -------------------------------- }

  procedure InitTempDataSet;
  begin
    cdsTemp.FieldDefs.Clear;
    cdsTemp.FieldDefs.Add( 'INVID', ftString, 20 );
    cdsTemp.FieldDefs.Add( 'INVDATE', ftDateTime );
    cdsTemp.FieldDefs.Add( 'INVUSEID', ftString, 3 );
    cdsTemp.CreateDataSet;
  end;

  { -------------------------------- }

  procedure SetRefNoFilter(AFieldName, ACodes: String);
  var
    aExtCode, aFilerText: String;
  begin
    if ( ACodes = EmptyStr ) then
    begin
      adoInv028Code.Filtered := False;
      adoInv028Code.Filter := EmptyStr;
    end else
    begin
      ACodes := TrimChar( ACodes, [''''] );
      aFilerText := EmptyStr;
      repeat
        aExtCode := ExtractValue( ACodes );
        aFilerText := ( aFilerText + Format( '%s=''%s''', [AFieldName, aExtCode] ) );
        if ( ACodes <> EmptyStr ) then aFilerText := ( aFilerText + ' OR ' );
      until ( ACodes = EmptyStr );
      adoInv028Code.Filtered := False;
      adoInv028Code.Filter := aFilerText; 
      adoInv028Code.Filtered := True;
    end;  
  end;

  { -------------------------------- }

  function FindItemId(const AValue: String): Boolean;
  begin
    Result := False;
    adoInv028Code.First;
    while not adoInv028Code.Eof do
    begin
      Result := ( adoInv028Code.FieldByName( 'ITEMID' ).AsString = AValue );
      if Result then Break;
      adoInv028Code.Next;
    end;
  end;

  { -------------------------------- }

  function CheckToSeeIfOldNullData(ACodes: String): Boolean;
  var
    aExtCode: String;
  begin
    ACodes := TrimChar( ACodes, [''''] );
    SetRefNoFilter( 'REFNO', '999' );
    repeat
      aExtCode := ExtractValue( ACodes );
      Result := FindItemId( aExtCode );
    until ( Result ) or ( ACodes = EmptyStr );
    SetRefNoFilter( EmptyStr, EmptyStr );
  end;

  { -------------------------------- }

var
  aInvDate, aYearMonth, aSql: String;
  aNvlInvSt,aNvlInvEd: String;
begin

  ClearTempDataSet;
  InitTempDataSet;
  {
  adoQryCommon2.Close;
  adoQryCommon2.SQL.Text := Format(
    ' SELECT UNIQUE INVDATE FROM INV007 A,                   ' +
    '   ( SELECT A.MAININVID FROM INV007 A                   ' +
    '      WHERE A.IDENTIFYID1 = ''%s''                      ' +
    '        AND A.IDENTIFYID2 = ''%s''                      ' +
    '        AND A.COMPID = ''%s''                           ' +
    '        AND A.INVID BETWEEN ''%s'' AND ''%s'' ) B       ' +
    '  WHERE A.MAININVID = B.MAININVID                       ' +
    '    AND A.IDENTIFYID1 = ''%s''                          ' +
    '    AND A.IDENTIFYID2 = ''%s'' AND A.COMPID = ''%s''    ',
    [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, aInvIdSt, aInvIdEd,
     IDENTIFYID1, IDENTIFYID2,  dtmMain.getCompID] );
  adoQryCommon2.Open;
  if ( adoQryCommon2.IsEmpty ) then
  begin
    adoQryCommon2.Close;
    Result := '?o???????L???d?????o???C';
    Exit;
  end;
  }
  {}
  { ???????????X???? Temp DataSet }
  {
  adoQryCommon2.SQL.Text := Format(
    ' SELECT UNIQUE INVID, INVDATE, INVUSEID FROM INV007 A,  ' +
    '   ( SELECT A.MAININVID FROM INV007 A                   ' +
    '      WHERE A.IDENTIFYID1 = ''%s''                      ' +
    '        AND A.IDENTIFYID2 = ''%s''                      ' +
    '        AND A.COMPID = ''%s''                           ' +
    '        AND A.INVID  BETWEEN ''%s'' AND ''%s'' ) B      ' +
    '  WHERE A.MAININVID = B.MAININVID                       ' +
    '    AND A.IDENTIFYID1 = ''%s''                          ' +
    '    AND A.IDENTIFYID2 = ''%s'' AND A.COMPID = ''%s''    ',
    [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, aInvIdSt, aInvIdEd,
     IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID] );
  }
  adoQryCommon2.SQL.Text := Format(
    ' SELECT UNIQUE INVID, INVDATE, INVUSEID FROM INV007 A,  ' +
    '   ( SELECT A.MAININVID FROM INV007 A                   ' +
    '      WHERE A.IDENTIFYID1 = ''%s''                      ' +
    '        AND A.IDENTIFYID2 = ''%s''                      ' +
    '        AND A.COMPID = ''%s''                           ',
    [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID]);
  if aInvs <> '' then
  begin
    adoQryCommon2.SQL.Text := adoQryCommon2.SQL.Text +
      ' AND INVID IN(' + aInvs + ')';
  end;
  if ( aInvIdSt <> '') And ( aInvIdEd <> '') then
  begin
    adoQryCommon2.SQL.Text := adoQryCommon2.SQL.Text +
      Format(' AND A.INVID  BETWEEN ''%s'' AND ''%s'' ) B      ',
      [aInvIdSt,aInvIdEd]);
  end else
  begin
    adoQryCommon2.SQL.Text := adoQryCommon2.SQL.Text + '  ) B ';
  end;
  adoQryCommon2.SQL.Text := adoQryCommon2.SQL.Text +
    format('  WHERE A.MAININVID = B.MAININVID                        ' +
            '    AND A.IDENTIFYID1 = ''%s''                          ' +
            '    AND A.IDENTIFYID2 = ''%s'' AND A.COMPID = ''%s''    ',
            [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID]);

  if dtmMain.GetStarEInvoice then
  begin
    adoQryCommon2.SQL.Text := adoQryCommon2.SQL.Text +
          ' AND INVOICEKIND = ' + IntToStr( aInvKind );
  end;
  adoQryCommon2.Open;
  adoQryCommon2.First;
  while not adoQryCommon2.Eof do
  begin
    cdsTemp.Append;
    cdsTemp.FieldByName( 'INVID' ).AsString :=
      adoQryCommon2.FieldByName( 'INVID' ).AsString;
    cdsTemp.FieldByName( 'INVUSEID' ).AsString :=
      adoQryCommon2.FieldByName( 'INVUSEID' ).AsString;
    cdsTemp.FieldByName( 'INVDATE' ).Value :=
      adoQryCommon2.FieldByName( 'INVDATE' ).Value;
    cdsTemp.Post;
    adoQryCommon2.Next;
  end;
  adoQryCommon2.Close;
  { ?}?l?L?o???? }
  adoInv028Code.Close;
  adoInv028Code.Open;
  { ?????? }
  { #5922 ???????O??????Integer???A 0.???? 1.???? 2.?D???? By Kin 2011/02/11}
  if ( AIsInvUse = 1 ) then
  begin
    SetRefNoFilter( 'REFNO', '1' );
    cdsTemp.First;
    while not cdsTemp.Eof do
    begin
      if ( cdsTemp.FieldByName( 'INVUSEID' ).AsString = EmptyStr ) or
         ( not FindItemId( cdsTemp.FieldByName( 'INVUSEID' ).AsString ) ) then
        cdsTemp.Delete
      else
        cdsTemp.Next;
    end;
  end;
  { ?L???? }
  { #5922 ???????O??????Integer???A 0.???? 1.???? 2.?D???? By Kin 2011/02/11}
  if (  AIsInvUse = 2 ) then
  begin
    { ???h???O?????????? }
    SetRefNoFilter( 'REFNO', '1' );
    cdsTemp.First;
    while not cdsTemp.Eof do
    begin
      if ( FindItemId( cdsTemp.FieldByName( 'INVUSEID' ).AsString ) ) then
        cdsTemp.Delete
      else
        cdsTemp.Next;
    end;
    { ?Y???]?w?????o?????~?????? }
    if ( aInvUseId <> EmptyStr ) then
    begin
      SetRefNoFilter( 'ITEMID', aInvUseId );
      cdsTemp.First;
      while not cdsTemp.Eof do
      begin
        if ( cdsTemp.FieldByName( 'INVUSEID' ).AsString <> EmptyStr ) and
           ( not FindItemId( cdsTemp.FieldByName( 'INVUSEID' ).AsString ) ) then
          cdsTemp.Delete
        else
          cdsTemp.Next;
      end;
      { ?L???????????o?????~?????????O?_???t 999 ???????? }
      if not CheckToSeeIfOldNullData( aInvUseId ) then
      begin
        cdsTemp.First;
        while not cdsTemp.Eof do
        begin
          if ( cdsTemp.FieldByName( 'INVUSEID' ).AsString = EmptyStr ) then
            cdsTemp.Delete
          else
            cdsTemp.Next;
        end;
      end;
    end;
  end;

  SetRefNoFilter( EmptyStr, EmptyStr );
  adoInv028Code.Close;

  if ( cdsTemp.IsEmpty ) then
  begin
    Result := '?o???????L???d?????o???C';
    ClearTempDataSet;
    Exit;
  end;
  { ?????O?_???b }
  cdsTemp.First;
  while not cdsTemp.Eof do
  begin
    aInvDate := FormatDateTime( 'yyyy/mm/dd', cdsTemp.FieldByName( 'INVDATE' ).AsDateTime );
    aYearMonth := Copy( aInvDate, 1, 4 ) + Copy( aInvDate, 6, 2 );
    if InvHasBeenLock( aYearMonth ) then
    begin
      Result := '?o???????~??:' + aYearMonth + ', ?w?g???b,?L?k?????C';
      ClearTempDataSet;
      Exit;
    end;
    cdsTemp.Next;
  end;
  {}
  { ?h???????X???o??????, ?q?l?p?????o?????q?l?o?? }
  cdsTemp.First;

  {}
 
  while not cdsTemp.Eof do
  begin
    aSql := Format(
      '  UPDATE /*+ RULE */ INV007                    ' +
      '     SET CANMODIFY = ''Y''                     ' +
      '   WHERE IDENTIFYID1 = ''%s''                  ' +
      '     AND IDENTIFYID2 = %s                      ' +
      '     AND COMPID = %s                           ' +
      '     AND INVID = ''%s''                        ' +
      '     AND ISOBSOLETE = ''N''                    ',
      [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID,
       cdsTemp.FieldByName( 'INVID' ).AsString ] );
    ExecuteSQL( aSql );
    cdsTemp.Next;
  end;
  { ?gLog }
  if ( aInvIdSt = '' ) OR ( aInvIdEd = '' ) then
  begin
    aNvlInvSt := '???J????';
    aNvlInvEd := '???J????';
  end else
  begin
    aNvlInvSt := aInvIdSt;
    aNvlInvEd := aInvIdEd;
  end;
  aSql := Format(
    '  INSERT INTO INV015                         ' +
    '            ( IDENTIFYID1,                   ' +
    '              IDENTIFYID2,                   ' +
    '              USERID,                        ' +
    '              STARTINVID,                    ' +
    '              ENDINVID,                      ' +
    '              UPTTIME      )                 ' +
    '     VALUES ( ''%s'',                        ' +
    '              %s,                            ' +
    '              ''%s'',                        ' +
    '              ''%s'',                        ' +
    '              ''%s'',                        ' +
    '              SYSDATE      )                 ',

    [IDENTIFYID1, IDENTIFYID2, dtmMain.getLoginUser, aNvlInvSt, aNvlInvEd] );
  ExecuteSQL( aSql );
  Result := '???v?????C';
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMainJ.synchronizeCompData(sI_MisDbOwner,sI_CurDateTime,sI_UserName : String);
var
    sL_SQL,sL_SrcSysId,sL_SrcSysName,sL_SrcInvoiceId,sL_SrcAddress : String;
    sL_SrcServiceType,sL_IsSelfCreated,sL_TEL, aServiceTypeEx : String;
    nL_Counts : Integer;

    { ------------------------------------------------------ }

    function CutServiceType(aSrv: String): String;
    begin
      Result := EmptyStr;
      if ( aSrv = EmptyStr ) then Exit;
      repeat
        Result := ( Result + Copy( aSrv, 1, 1 ) );
        Delete( aSrv, 1, 1 );
        if ( aSrv <> EmptyStr ) then Result := ( Result + ',' ); 
      until ( aSrv = EmptyStr );
    end;
    
    { ------------------------------------------------------ }

begin
    //?P?B???q????

    //???XMIS?????q????
    sL_SQL := 'select SYSID,SYSNAME,INVOICEID,ADDRESS,SERVICETYPE from ' + sI_MisDbOwner +
              '.SO041 order by SYSID ';



    with dtmSo.adoQrySynSO041 do
    begin
      Close;
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;

      if RecordCount <> 0 then
      begin
        sL_IsSelfCreated := 'N';
        sL_TEL := '';

        First;
        while not Eof do
        begin
          sL_SrcSysId := FieldByName('SYSID').AsString;
          sL_SrcSysName := FieldByName('SYSNAME').AsString;
          sL_SrcInvoiceId := FieldByName('INVOICEID').AsString;
          sL_SrcAddress := FieldByName('ADDRESS').AsString;
          sL_SrcServiceType := FieldByName('SERVICETYPE').AsString;
          aServiceTypeEx := CutServiceType( sL_SrcServiceType );

          //?Y?????q?N?X???w?s?b?o???t??,?h?s?W????????,?Y?s?b?h??????????????
          sL_SQL := 'SELECT COUNT(*) FROM INV001 WHERE COMPID=''' + sL_SrcSysId + '''';
          nL_Counts := StrToInt(openSQL(sL_SQL));

          if nL_Counts <= 0 then
          begin
            sL_SQL := 'insert into INV001 (IDENTIFYID1,IDENTIFYID2,COMPID,COMPSNAME,COMPNAME,' +
                      'BUSINESSID,TAXID,MANAGER,TEL,COMPADDR,NOTITYPE,DECACOUNT,PAYCOUNT,TAXRATE,TAXDIVNAME,' +
                      'TAXGETDATE,TAXNUM1,TAXNUM2,ISSELFCREATED,RPTNAME1,RPTNAME2,RPTNAME3,RPTNAME4,RPTNAME5,RPTNAME6,UPTTIME,UPTEN,SERVICETYPESTR)' +
                      ' values(''' + IDENTIFYID1 + ''',' + IDENTIFYID2 + ',''' + sL_SrcSysId +
                      ''',''' + sL_SrcSysName + ''',''' + sL_SrcSysName + ''',''' + sL_SrcInvoiceId +
                      ''',null,null,''' + sL_TEL + ''',''' + sL_SrcAddress +
                      ''',null,null,null,null,null,null,null,null,''' + sL_IsSelfCreated +
                      ''','''','''','''','''','''','''',' +
                      TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_CurDateTime)) +
                      ',''' + sI_UserName + ''',''' + aServiceTypeEx + ''')';

             executeSQL(sL_SQL);
          end
          else
          begin
            sL_SQL := 'update inv001 set COMPSNAME=''' + sL_SrcSysName + ''',COMPNAME=''' +
                      sL_SrcSysName + ''',BUSINESSID=''' +  sL_SrcInvoiceId +
                      ''',COMPADDR=''' + sL_SrcAddress + ''',UPTEN=''' + sI_UserName +
                      ''',UPTTIME=' + TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_CurDateTime)) +
                      ',SERVICETYPESTR=''' + aServiceTypeEx + '''' +
                      ' where IdentifyId1=''' + IDENTIFYID1 + ''' and IdentifyId2=' + IDENTIFYID2 +
                      ' and CompID=''' + sL_SrcSysId + '''';

            executeSQL(sL_SQL);
          end;

          Next;
        end;
      end;
    end;
end;

procedure TdtmMainJ.synchronizeChargeItem(sI_MisDbOwner,sI_CurDateTime,
  sI_UserName: String);
var
    sL_SQL,sL_SrcCodeNo,sL_SrcDesc,sL_SrcSign,sL_IsSelfCreated,sL_SrcTaxCode,sL_TaxName : String;
begin
    sL_SQL := Format(
      ' DELETE INV005 WHERE IDENTIFYID1 = ''%s''          ' +
      '   AND IDENTIFYID2 = ''%s'' AND COMPID = ''%s''    ' +
      '   AND ISSELFCREATED = ''N''                       ',
      [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID] );

    executeSQL(sL_SQL);

    sL_IsSelfCreated := 'N';
    //???XMIS?????O????

    dtmSo.adoQrySynCD019.Close;
    dtmSo.adoQrySynCD019.SQL.Clear;
    dtmSo.adoQrySynCD019.SQL.Text := Format(
      ' SELECT A.CODENO, A.DESCRIPTION,        ' +
      '        A.SIGN, A.TAXCODE, A.TAXNAME    ' +
      '   FROM %s.CD019 A, %s.CD033 B          ' +
      '  WHERE A.TAXCODE = B.CODENO            ' +
      '    AND B.TAXFLAG = 1                   ' +
      '  ORDER BY A.CODENO ', [sI_MisDbOwner, sI_MisDbOwner] );
    dtmSo.adoQrySynCD019.Open;
    dtmSo.adoQrySynCD019.First;
    while not dtmSo.adoQrySynCD019.Eof do
    begin
      sL_SrcCodeNo := dtmSo.adoQrySynCD019.FieldByName('CODENO').AsString;
      sL_SrcDesc := dtmSo.adoQrySynCD019.FieldByName('DESCRIPTION').AsString;
      sL_SrcSign := dtmSo.adoQrySynCD019.FieldByName('SIGN').AsString;
      sL_SrcTaxCode := dtmSo.adoQrySynCD019.FieldByName('TaxCode').AsString;

      //??????INVD03
      if sL_SrcTaxCode = '1' then
        sL_TaxName := '???|'
      else if sL_SrcTaxCode = '2' then
        sL_TaxName := '?s?|?v'
      else if sL_SrcTaxCode = '3' then
        sL_TaxName := '?K?|'
      else if sL_SrcTaxCode = '4' then
        sL_TaxName := '?L???|'
      else if sL_SrcTaxCode = '5' then
        sL_TaxName := '???|-????'
      else if sL_SrcTaxCode = '6' then
        sL_TaxName := '?K?|-????'
      else if sL_SrcTaxCode = '' then
      begin
        sL_SrcTaxCode := 'null';
        sL_TaxName := '';
      end;

      sL_SQL := Format(
        ' INSERT INTO INV005 ( IDENTIFYID1, IDENTIFYID2, ITEMID,      ' +
        '    DESCRIPTION, SIGN, ISSELFCREATED, UPTTIME, UPTEN,        ' +
        '    TAXCODE, TAXNAME, ITEMIDREF, COMPID )                    ' +
        ' VALUES ( ''%s'', ''%s'',  ''%s'',                           ' +
        '    ''%s'', ''%s'',  ''%s'', SYSDATE, ''%s'',                ' +
        '    ''%s'', ''%s'', NULL, ''%s''  )                          ',
        [IDENTIFYID1, IDENTIFYID2, sL_SrcCodeNo,
         sL_SrcDesc, sL_SrcSign, sL_IsSelfCreated, dtmMain.getLoginUser,
         sL_SrcTaxCode, sL_TaxName, dtmMain.getCompID] );

      executeSQL(sL_SQL);

      dtmSo.adoQrySynCD019.Next;
    end;
    dtmSo.adoQrySynCD019.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMainJ.SynchronizeCD095;
var
  aSoDataReader: TADOQuery;
  aWriter: TADOQuery;
  aEffects: Integer;
begin
  aSoDataReader := dtmSO.adoQrySynCD019;
  aWriter := dtmMain.adoComm2;
  {}
  aSoDataReader.Close;
  aSoDataReader.SQL.Text := ' SELECT * FROM CD095 ORDER BY CodeNo ';
  aSoDataReader.Open;
  aSoDataReader.First;
  while not aSoDataReader.Eof do
  begin
    try
      aWriter.Close;
      aWriter.SQL.Text := Format(
        ' SELECT COUNT(1) FROM INV028    ' +
        '  WHERE IDENTIFYID1 = ''%S''    ' +
        '    AND IDENTIFYID2 = ''%S''    ' +
        '    AND COMPID = ''%S''         ' +
        '    AND ITEMID = ''%S''         ',
        [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID,
         aSoDataReader.FieldByName( 'CODENO' ).AsString] );
      aWriter.Open;
      aEffects := aWriter.Fields[0].AsInteger;
      aWriter.Close;
      if ( aEffects > 0 ) then
      begin
        aWriter.SQL.Text := Format(
          ' UPDATE INV028                 ' +
          '    SET DESCRIPTION = ''%s'',  ' +
          '        REFNO = ''%s'',        ' +
          '        UPDEN = ''%s'',        ' +
          '        UPDTIME = TO_CHAR( SYSDATE, ''YYYY/MM/DD HH24:MI:SS'' ), ' +
          '        ISSELFCREATED = ''N''  ' +
          '  WHERE IDENTIFYID1 = ''%s''   ' +
          '    AND IDENTIFYID2 = ''%s''   ' +
          '    AND COMPID = ''%s''        ' +
          '    AND ITEMID = ''%s''        ',
          [aSoDataReader.FieldByName( 'Description' ).AsString,
          aSoDataReader.FieldByName( 'RefNo' ).AsString,
          dtmMain.getLoginUser,
          IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID,
          aSoDataReader.FieldByName( 'CODENO' ).AsString] );
        aWriter.ExecSQL;
      end else
      begin
        aWriter.SQL.Text := Format(
          ' INSERT INTO INV028 ( IDENTIFYID1, IDENTIFYID2,  ' +
          '    COMPID, ITEMID, DESCRIPTION, REFNO, UPDEN,   ' +
          '    UPDTIME, ISSELFCREATED )                     ' +
          '  VALUES ( ''%s'', ''%s'', ''%s'', ''%s'',       ' +
          '           ''%s'', ''%s'', ''%s'',               ' +
          '           TO_CHAR( SYSDATE, ''YYYY/MM/DD HH24:MI:SS'' ), ' +
          '           ''N'' )                               ',
          [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID,
           aSoDataReader.FieldByName( 'CODENO' ).AsString,
           aSoDataReader.FieldByName( 'DESCRIPTION' ).AsString,
           aSoDataReader.FieldByName( 'REFNO' ).AsString,
           dtmMain.getLoginUser] );
        aWriter.ExecSQL;
      end;
    except
      on E: Exception do
      begin
        ErrorMsg( Format( '?P?B?o?????~?N?X????, ?T??:%s?C', [E.Message] ) );
        Break;
      end;
    end;
    aSoDataReader.Next;
  end;
  aSoDataReader.Close;
  aWriter.Close;
  aWriter.SQL.Clear;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMainJ.SynchronizeChargeItem2: Boolean;
var
  aSoDataReader: TADOQuery;
  aWriter: TADOQuery;
  aEffects: Integer;
  aTaxCode, aTaxName: String;
begin
  Result := True;
  aSoDataReader := dtmSO.adoQrySynCD019;
  aWriter := dtmMain.adoComm2;
  {}
  aSoDataReader.Close;
  aSoDataReader.SQL.Text := Format(
    ' SELECT A.CODENO, A.DESCRIPTION,        ' +
    '        A.SIGN, A.TAXCODE, A.TAXNAME    ' +
    '   FROM %s.CD019 A, %s.CD033 B          ' +
    '  WHERE A.TAXCODE = B.CODENO            ' +
    '    AND B.TAXFLAG = 1                   ' +
    '  ORDER BY A.CODENO ', [dtmMain.getMisDbOwner, dtmMain.getMisDbOwner] );
  aSoDataReader.Open;
  aSoDataReader.First;
  while not aSoDataReader.Eof do
  begin
    try
      aWriter.Close;
      aWriter.SQL.Text := Format(
        ' SELECT COUNT(1) FROM INV005    ' +
        '  WHERE IDENTIFYID1 = ''%s''    ' +
        '    AND IDENTIFYID2 = ''%s''    ' +
        '    AND COMPID = ''%s''         ' +
        '    AND ITEMID = ''%s''         ' +
        '    AND ISSELFCREATED = ''N''   ',
        [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID,
         aSoDataReader.FieldByName( 'CODENO' ).AsString] );
      aWriter.Open;
      aEffects := aWriter.Fields[0].AsInteger;
      aWriter.Close;

      aTaxCode := aSoDataReader.FieldByName( 'TAXCODE' ).AsString;

      //??????INVD03
      if aTaxCode = '1' then
        aTaxName := '???|'
      else if aTaxCode = '2' then
        aTaxName := '?s?|?v'
      else if aTaxCode = '3' then
        aTaxName := '?K?|'
      else if aTaxCode = '4' then
        aTaxName := '?L???|'
      else if aTaxCode = '5' then
        aTaxName := '???|-????'
      else if aTaxCode = '6' then
        aTaxName := '?K?|-????'
      else if aTaxCode = EmptyStr then
      begin
        aTaxCode := EmptyStr;
        aTaxName := EmptyStr;
      end;

      if ( aEffects > 0 ) then
      begin
        aWriter.SQL.Text := Format(
         ' UPDATE INV005                ' +
         '    SET DESCRIPTION = ''%s'', ' +
         '        SIGN = ''%s'',        ' +
         '        TAXCODE = ''%s'',     ' +
         '        TAXNAME = ''%s'',     ' +
         '        UPTEN = ''%s'',       ' +
         '        UPTTIME = SYSDATE     ' +
         '  WHERE IDENTIFYID1 = ''%s''  ' +
         '    AND IDENTIFYID2 = ''%s''  ' +
         '    AND COMPID = ''%s''       ' +
         '    AND ITEMID = ''%s''       ',
         [aSoDataReader.FieldByName( 'DESCRIPTION' ).AsString,
          aSoDataReader.FieldByName( 'SIGN' ).AsString,
          aTaxCode,
          aTaxName,
          dtmMain.getLoginUser,
          IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID,
          aSoDataReader.FieldByName( 'CODENO' ).AsString] );
      end else
      begin
        aWriter.SQL.Text := Format(
          ' INSERT INTO INV005 ( IDENTIFYID1, IDENTIFYID2, ITEMID,      ' +
          '    DESCRIPTION, SIGN, ISSELFCREATED, UPTTIME, UPTEN,        ' +
          '    TAXCODE, TAXNAME, ITEMIDREF, COMPID )                    ' +
          ' VALUES ( ''%s'', ''%s'',  ''%s'',                           ' +
          '    ''%s'', ''%s'',  ''%s'', SYSDATE, ''%s'',                ' +
          '    ''%s'', ''%s'', NULL, ''%s''  )                          ',
        [IDENTIFYID1, IDENTIFYID2,
         aSoDataReader.FieldByName( 'CODENO' ).AsString,
         aSoDataReader.FieldByName( 'DESCRIPTION' ).AsString,
         aSoDataReader.FieldByName( 'SIGN' ).AsString,
         'N', dtmMain.getLoginUser,
         aTaxCode,
         aTaxName,
         dtmMain.getCompID] );
      end;
      aWriter.ExecSQL;
    except
      on E: Exception do
      begin
        ErrorMsg( Format( '?P?B???O????????, ?T??:%s?C', [E.Message] ) );
        Result := False;
        Break;
      end;
    end;
    aSoDataReader.Next;
  end;
  aSoDataReader.Close;
  aWriter.SQL.Clear;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMainJ.synchronizePayType(sI_MisDbOwner, sI_CurDateTime,
  sI_UserName: String);
var
    sL_SQL,sL_CodeNo,sL_Description,sL_IsSelfCreated : String;
begin
    sL_SQL := 'delete inv020 where ISSELFCREATED=''N''';
    executeSQL(sL_SQL);


    sL_IsSelfCreated := 'N';
    //???XMIS???I??????????
    sL_SQL := 'select CODENO,DESCRIPTION from ' + sI_MisDbOwner +
              '.CD032 order by CODENO ';
    with dtmSo.adoQrySynCD032 do
    begin
      Close;
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;

      First;
      while not Eof do
      begin
        sL_CodeNo := FieldByName('CODENO').AsString;
        sL_Description := FieldByName('DESCRIPTION').AsString;

        sL_SQL := 'insert into INV020 (IDENTIFYID1,IDENTIFYID2,ITEMID,' +
                  'DESCRIPTION,ISSELFCREATED,UPTTIME,UPTEN) ' +
                  ' values(''' + IDENTIFYID1 + ''',' + IDENTIFYID2 + ',''' +
                  sL_CodeNo + ''',''' + sL_Description + ''',''' +
                  sL_IsSelfCreated + ''',' +
                  TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_CurDateTime)) +
                  ',''' + sI_UserName + ''')';


        executeSQL(sL_SQL);
        
        Next;
      end;
    end;
end;

procedure TdtmMainJ.synchronizeServiceType(sI_MisDbOwner, sI_CurDateTime,
  sI_UserName: String);
var
    sL_SQL,sL_IsSelfCreated,sL_CodeNo,sL_Description : String;
begin
   sL_SQL := 'DELETE INV022 WHERE ISSELFCREATED=''N''';
   executeSQL(sL_SQL);


    sL_IsSelfCreated := 'N';
    //???XMIS???I??????????
    sL_SQL := 'SELECT CODENO, DESCRIPTION FROM ' + sI_MisDbOwner +
              '.CD046 ORDER BY CODENO ';

    with dtmSo.adoQrySynCD032 do
    begin
      Close;
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;

      First;
      while not Eof do
      begin
        sL_CodeNo := FieldByName('CODENO').AsString;
        sL_Description := FieldByName('DESCRIPTION').AsString;

        sL_SQL := 'insert into INV022 (IDENTIFYID1,IDENTIFYID2,ITEMID,' +
                  'DESCRIPTION,ISSELFCREATED,UPTTIME,UPTEN)' +
                  ' values(''' + IDENTIFYID1 + ''',' + IDENTIFYID2 + ',''' +
                  sL_CodeNo + ''',''' + sL_Description + ''',''' +
                  sL_IsSelfCreated + ''',' +
                  TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_CurDateTime)) +
                  ',''' + sI_UserName + ''')';

        executeSQL(sL_SQL);

        Next;
      end;
    end;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMainJ.synchronizeCustID(sI_MisDbOwner, sI_CurDateTime,
  sI_UserName, sI_CompID,sI_CustID,sI_MisDbCompCode : String);
var
    sL_SQL : String;
begin
    sL_SQL :='delete inv002 where IdentifyId1=''' + IDENTIFYID1 + ''' and IdentifyId2=' + IDENTIFYID2 +
              ' and CompID=''' + sI_CompID + ''' and CustID=''' + sI_CustID +
              ''' and IsSelfCreated=''N''';
    executeSQL(sL_SQL);



    sL_SQL :='delete inv019 where IdentifyId1=''' + IDENTIFYID1 + ''' and IdentifyId2=' + IDENTIFYID2 +
              ' and CompID=''' + sI_CompID + ''' and CustID=''' + sI_CustID + '''';
    executeSQL(sL_SQL);


    copyMisCustInfoToInv002(sI_MisDbOwner, sI_CurDateTime,sI_UserName, sI_CompID,sI_CustID,sI_MisDbCompCode);

end;

procedure TdtmMainJ.copyMisCustInfoToInv002(sI_MisDbOwner, sI_CurDateTime,
  sI_UserName, sI_CompID, sI_CustID, sI_MisDbCompCode: String);
var
    sL_SQL,sL_CustName,sL_CustSName,sL_Tel1,sL_Tel2,sL_Tel3 : String;
    sL_IsSelfCreated,sL_MailAddrNo,sL_MailAddress : String;
    sL_MZipCode : String;
begin

    sL_SQL := Format(
      ' SELECT CUSTNAME, MAILADDRNO, MAILADDRESS,   ' +
      '        TEL1, TEL2, TEL3                     ' +
      ' FROM %s.SO001                               ' +
      ' WHERE CUSTID = ''%s''                       ' +
      '   AND COMPCODE= ''%s''                      ',
      [sI_MisDbOwner, sI_CustID, sI_MisDbCompCode] );

    with dtmSo.adoQrySynSO001 do
    begin
      Close;
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;

      sL_IsSelfCreated := 'N';
      First;
      while not Eof do
      begin
        sL_CustName := FieldByName('CUSTNAME').AsString;
        sL_CustSName := sL_CustName;

        sL_MailAddrNo := FieldByName('MAILADDRNO').AsString;
        sL_MailAddress := FieldByName('MAILADDRESS').AsString;
        sL_Tel1 := FieldByName('Tel1').AsString;
        sL_Tel2 := FieldByName('Tel2').AsString;
        sL_Tel3 := FieldByName('Tel3').AsString;

        sL_MZipCode := copyMisZipcodeFromSo014(sI_MisDbOwner,sI_MisDbCompCode,sL_MailAddrNo);

        sL_SQL := 'iNSERT INTO INV002 (IDENTIFYID1,IDENTIFYID2,COMPID,CUSTID,CUSTSNAME,CUSTNAME,' +
                  'MZIPCODE,MAILADDR,TEL1,TEL2,TEL3,FAX,APPCONTACTEE1,APPCONTACTEE2,FINACONTACTEE1,' +
                  'FINACONTACTEE2,ISSELFCREATED,UPTTIME,UPTEN)' +
                  ' values(''' + IDENTIFYID1 + ''',' + IDENTIFYID2 + ',''' +
                  sI_CompID + ''',''' + sI_CustID + ''',''' + sL_CustSName + ''',''' +
                  sL_CustName + ''',''' + sL_MZipCode + ''',''' + sL_MailAddress + ''',''' +
                  sL_Tel1 + ''',''' + sL_Tel2 + ''',''' + sL_Tel3 + ''','''','''','''','''','''',''' +
                  sL_IsSelfCreated +  ''',' + TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_CurDateTime)) +
                  ',''' + sI_UserName + ''')';

        executeSQL(sL_SQL);

        copyMisCustInfoToInv019(sI_MisDbOwner, sI_CurDateTime,sI_UserName, sI_CompID,sI_CustID,sI_MisDbCompCode,sL_CustName,sL_CustSName,sL_MZipCode,sL_MailAddress);

        Next;
      end;
    end;
end;

function TdtmMainJ.copyMisZipcodeFromSo014(sI_MisDbOwner, sI_MisDbCompCode,
  sI_MailAddrNo: String): String;
var
    sL_SQL : String;
begin
    sL_SQL := 'SELECT ZIPCODE FROM ' + sI_MisDbOwner +
              '.SO014 WHERE ADDRNO=' + sI_MailAddrNo +
              '  AND COMPCODE=''' + sI_MisDbCompCode + '''';

      with dtmSo.adoComm do
      begin
        Close;
        SQL.Clear;
        SQL.Add(sL_SQL);
        Open;
        if RecordCount <> 0 then
          Result := FieldByName('ZIPCODE').AsString
        else
          Result := '';
        Close;
      end;
end;

procedure TdtmMainJ.copyMisCustInfoToInv019(sI_MisDbOwner, sI_CurDateTime,
  sI_UserName, sI_CompID, sI_CustID, sI_MisDbCompCode, sI_CustName,
  sI_CustSName,sI_MZipCode,sI_MailAddress : String);
var
    sL_SQL,sL_IsSelfCreated,sL_TitleName,sL_BusinessId,sL_InvAddr,sL_TitleSName : String;
    nL_TitleID : Integer;
begin
    sL_SQL := 'SELECT CUSTID,INVTITLE,INVNO,INVADDRESS FROM ' + sI_MisDbOwner +
              '.SO002 WHERE CUSTID=''' + sI_CustID + ''' and CompCode=''' + sI_MisDbCompCode + '''';


    with dtmSo.adoQrySynSO002 do
    begin
      Close;
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;

      sL_IsSelfCreated := 'N';
      nL_TitleID := 0;
      First;
      while not Eof do
      begin
        nL_TitleID := nL_TitleID + 1;
        sL_TitleName := FieldByName('INVTITLE').AsString;  //?G?p??????,so002.INVTITLE?|?O???r??,?T?p???~?|????
        sL_BusinessId := FieldByName('INVNO').AsString;
        sL_InvAddr := FieldByName('INVADDRESS').AsString;
        sL_TitleSName := sL_TitleName;   //?G?p??????,so002.INVTITLE?|?O???r??,?T?p???~?|????

        if sL_TitleName = '' then
        begin
          sL_TitleName := sI_CustName;
          sL_TitleSName := sI_CustSName;
        end;

        sL_SQL := 'insert into Inv019 (IDENTIFYID1,IDENTIFYID2,COMPID,CUSTID,TITLEID,TITLESNAME,' +
                  'TITLENAME,BUSINESSID,MZIPCODE,MAILADDR,INVADDR,MEMO,UPTTIME,UPTEN) ' +
                  ' values(''' + IDENTIFYID1 + ''',' + IDENTIFYID2 + ',''' +
                  sI_CompID + ''',''' + sI_CustID + ''',' + IntToStr(nL_TitleID) + ',''' +
                  sL_TitleSName + ''',''' + sL_TitleName + ''',''' + sL_BusinessId + ''',''' +
                  sI_MZipCode + ''',''' + sI_MailAddress + ''',''' + sL_InvAddr + ''','''',' +
                  TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_CurDateTime)) +
                  ',''' + sI_UserName + ''')';

        executeSQL(sL_SQL);

        Next;
      end;
    end;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMainJ.getAllInv003Data: Integer;
begin
  adoInv003Code.Close;
  adoInv003Code.SQL.Text := ' select * from inv003 ';
  adoInv003Code.Open;
  Result := adoInv003Code.RecordCount;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMainJ.updateInv016InvAmount(sI_UnitPrice,sI_TaxAmount,sI_InvAmount:Double;
      sI_SEQ, sI_CompID,sI_StartChargeDate,sI_EndChargeDate: String);
var
    sL_SQL,sL_Where : String;
begin
    sL_SQL := 'UPDATE INV016 SET SaleAMount=' + FloatToStr(sI_UnitPrice) +
              ' , TaxAmount=' + FloatToStr(sI_TaxAmount) +
              ' , InvAmount=' + FloatToStr(sI_InvAmount);


    sL_Where := ' WHERE COMPID= '+''''+sI_CompID+''''+
      ' AND CHARGEDATE between TO_DATE('+''''+sI_StartChargeDate+''','''+
      Date_Format+''') AND TO_DATE('+''''+sI_EndChargeDate+''','''+
      Date_Format+''') AND BEASSIGNEDINVID=''N'' AND ISVALID=''Y''';


    if sI_SEQ <> '' then
      sL_Where := sL_Where + ' AND SEQ=''' + sI_SEQ + '''';

    sL_SQL := sL_SQL + sL_Where;

    executeSQL(sL_SQL);

end;

procedure TdtmMainJ.updateInv016StopFlag(sI_SEQ,sI_CompID,sI_StartChargeDate,sI_EndChargeDate : String);
var
    sL_SQL,sL_Where : String;
begin
    sL_SQL := 'UPDATE INV016 SET STOPFLAG=1, SHOULDBEASSIGNED=''N'' ';

    sL_Where := ' WHERE COMPID= '+''''+sI_CompID+''''+
      ' AND CHARGEDATE between TO_DATE('+''''+sI_StartChargeDate+''','''+
      Date_Format+''') AND TO_DATE('+''''+sI_EndChargeDate+''','''+
      Date_Format+''') AND BEASSIGNEDINVID=''N'' AND ISVALID=''Y''';


    if sI_SEQ <> '' then
      sL_Where := sL_Where + ' AND SEQ=''' + sI_SEQ + '''';

    sL_SQL := sL_SQL + sL_Where;

    executeSQL(sL_SQL);
end;

function TdtmMainJ.CalcCreateInvCounts(const aChargeDateSt,
  aChargeDateEd, aHowToCreate: String): Integer;
var
  aSql: String;
begin
  aSql := Format(
    ' SELECT SUM( CEIL( COUNT( B.SEQ ) / %d ) ) AS COUNTS  ' +
    '   FROM INV016, A, INV017 B                   ' +
    '  WHERE A.SEQ = B.SEQ                         ' +
    '    AND A.COMPID = ''%s''                     ' +
    '    AND A.CHARGEDATE BETWEEN                  ' +
    '        TO_DATE( ''%s'', ''YYYY/MM/DD'' ) AND ' +
    '        TO_DATE( ''%s'', ''YYYY/MM/DD'' )     ' +
    '    AND A.BEASSIGNEDINVID = ''N''             ' +
    '    AND A.ISVALID = ''Y''                     ' +
    '    AND A.HOWTOCREATE = ''%s''                ' +
    '    AND A.SHOULDBEASSIGNED = ''Y''            ' +
    '    AND A.INVAMOUNT > 0                       ' +
    '    AND A.TAXTYPE <> ''0''                    ' +
    '    AND B.SHOULDBEASSIGNED = ''Y''            ',
    [dtmMain.GetAutoCreateNum, dtmMain.getCompID,
     aChargeDateSt, aChargeDateEd, aHowToCreate] );
  adoQryCommon.Close;
  adoQryCommon.SQL.Text := aSql;
  adoQryCommon.Open;
  Result := adoQryCommon.FieldByName( 'COUNTS' ).AsInteger;
  adoQryCommon.Close;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMainJ.IsMultiInv(const aInvId: String): Boolean;
begin
  adoQryCommon.Close;
  adoQryCommon.SQL.Text := Format(
    ' SELECT COUNT(1) AS COUNTS FROM INV007 A,                 ' +
    '   ( SELECT A.MAININVID FROM INV007 A                     ' +
    '      WHERE A.IDENTIFYID1 = ''%s''                        ' +
    '        AND A.IDENTIFYID2 = ''%s''                        ' +
    '        AND A.COMPID = ''%s'' AND A.INVID = ''%s'' ) B    ' +
    '  WHERE A.MAININVID = B.MAININVID                         ' +
    '    AND A.IDENTIFYID1 = ''%s''                            ' +
    '    AND A.IDENTIFYID2 = ''%s'' AND A.COMPID = ''%s''      ',
    [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, aInvId,
     IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID] );
  adoQryCommon.Open;
  Result := ( adoQryCommon.FieldByName( 'COUNTS' ).AsInteger > 1 );
  adoQryCommon.Close;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMainJ.IsMultiInvEx(const aMainInvId: String): Boolean;
begin
  adoQryCommon.Close;
  adoQryCommon.SQL.Text := Format(
     ' SELECT /*+ RULE */ COUNT(1) FROM INV007 ' +
     '  WHERE MAININVID = ''%s''               ' +
     '   AND IDENTIFYID1 = ''%s''              ' +
     '   AND IDENTIFYID2 = ''%s''              ' +
     '   AND COMPID = ''%s''                   ',
     [aMainInvId, IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID] );
  adoQryCommon.Open;
  Result := ( adoQryCommon.FieldByName( 'COUNTS' ).AsInteger > 1 );
  adoQryCommon.Close;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMainJ.InvDeatilLinkToMIS(const aInvId, aSeq: String): Boolean;
begin
  adoQryCommon2.Close;
  adoQryCommon2.SQL.Text := Format(
    ' SELECT LINKTOMIS FROM INV008    ' +
    '  WHERE INVID = ''%s''           ' +
    '    AND SEQ = ''%s''             ', [aInvId, aSeq] );
  adoQryCommon2.Open;
  Result := ( Nvl( adoQryCommon2.FieldByName( 'LINKTOMIS' ).AsString, 'N' ) = 'Y' );
  adoQryCommon2.Close;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMainJ.GetAccountNo(const aInvId, aSeq, aBillId,
  aBillidItemNo: String): String;
var
  aSql, aSql2: String;
begin
  Result := EmptyStr;
  aSql := Format(
    ' SELECT DISTINCT ACCOUNTNO FROM INV008    ' +
    '  WHERE INVID = ''%s''                    ', [aInvId] );
  aSql2 := Format(
    ' SELECT DISTINCT ACCOUNTNO FROM INV008A   ' +
    '  WHERE INVID = ''%s''                    ', [aInvId] );
  {}
  if ( aSeq <> EmptyStr ) then
  begin
    aSql := ( aSql + Format( ' AND SEQ = ''%s'' ', [aSeq] ) );
    aSql2 := ( aSql2 + Format( ' AND SEQ = ''%s'' ', [aSeq] ) );
  end;
  {}
  if ( aBillId <> EmptyStr ) and ( aBillIdItemNo <> EmptyStr ) then
  begin
    aSql := ( aSql + Format( ' AND BILLID = ''%s'' AND BILLIDITEMNO = ''%s'' ',
      [aBillId, aBillidItemNo] ) );
    aSql2 := ( aSql2 + Format( ' AND BILLID = ''%s'' AND BILLIDITEMNO = ''%s'' ',
      [aBillId, aBillidItemNo] ) );
  end;
  {}
  aSql := ( aSql + ' AND ACCOUNTNO IS NOT NULL ' );
  aSql2 := ( aSql2 + ' AND ACCOUNTNO IS NOT NULL ' );
  {}
  adoQryCommon2.Close;
  adoQryCommon2.SQL.Text := Format( ' %s union %s ', [aSql, aSql2] );
  adoQryCommon2.Open;
  adoQryCommon2.First;
  while not adoQryCommon2.Eof do
  begin
    Result := ( Result + adoQryCommon2.FieldByName( 'ACCOUNTNO' ).AsString );
    adoQryCommon2.Next;
    if not adoQryCommon2.Eof then Result := ( Result + ',' );
  end;
  adoQryCommon2.Close;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMainJ.GetFacisNo(const aInvId, aSeq: String): String;
begin
  Result := EmptyStr;
  adoQryCommon2.Close;
  adoQryCommon2.SQL.Text := Format(
    '  SELECT DISTINCT FACISNO FROM INV008      ' +
    '   WHERE INVID = ''%s'' AND SEQ = ''%s''   ' +
    '     AND FACISNO IS NOT NULL               ' +
    '  UNION                                    ' +
    '  SELECT DISTINCT FACISNO FROM INV008A     ' +
    '   WHERE INVID = ''%s'' AND SEQ = ''%s''   ' +
    '     AND FACISNO IS NOT NULL               ',
     [aInvId, aSeq, aInvId, aSeq] );
  adoQryCommon2.Open;
  adoQryCommon2.First;
  while not adoQryCommon2.Eof do
  begin
    Result := ( Result + adoQryCommon2.FieldByName( 'FACISNO' ).AsString );
    adoQryCommon2.Next;
    if not adoQryCommon2.Eof then Result := ( Result + ',' );
  end;
  adoQryCommon2.Close;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMainJ.GetAllowanceNo: String;
begin
  adoQryCommon2.Close;
  adoQryCommon2.SQL.Text :=
    ' select to_char( sysdate, ''yyyymm'' ) ||            ' +
    '        lpad( s_inv014_allowance.nextval, 6, ''0'' ) ' +
    '   from dual ';
  adoQryCommon2.Open;
  Result := adoQryCommon2.Fields[0].AsString;
  adoQryCommon2.Close;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMainJ.GetIfPrintCheckValue: String;
begin
  Result := ' ''Y'', ''N'' ';
  if dtmMain.GetIfPrintCheck then
    Result := ' ''Y'' ';
end;

{ ---------------------------------------------------------------------------- }
function TdtmMainJ.GetFulDate(aType: Integer; aDate: String): String;
var
  aTmp : string;
begin
  if aType = 0 then
  begin
    Result := aDate + '01';
  end else
  begin
    aTmp := IntToStr(StrToInt(Copy(aDate,5,2))+1);
    if Length(aTmp) = 1 then
    begin
      aTmp := '0'+aTmp;
    end;
    Result := _monthLastDay(Copy(aDate,1,4)+aTmp,0);
  end;
end;

function TdtmMainJ.getAllInv040Data: Integer;
var aWhere : String;
begin
  //#5966 ?W?[?o?????????e By Kin 2011/03/29
  aWhere := EmptyStr;
  adoInv040Code.Close;
  adoInv040Code.SQL.Text :=
    ' select A.COMPID,                   ' +
    ' ( Case SendType                    ' +
    ' WHEN 1 THEN ''E-MAIL''             ' +
    ' WHEN 2 THEN ''???T''               ' +
    ' WHEN 3 THEN ''TVMail''             ' +
    ' WHEN 4 THEN ''CM ???y''            ' +
    ' ELSE ''?S???]?w''                  ' +
    ' END ) AS SENDTYPE2,                ' +
    ' A.EMAILHEAD,A.SendContext,SENDTYPE,' +
    ' A.BatHeadFile,A.BatModuleFile,     ' +
    ' A.SendDonatetext,                  ' +
    ' A.UptEN,A.UptTime                  ' +
    ' from inv040 A                      ';
  aWhere := '0';
  if dtmMain.GetStarEmail then
  begin
     aWhere := aWhere + ',1';
  end;
  if dtmMain.GetStarMessage then aWhere := aWhere + ',2';
  if dtmMain.GetStarTVmail then aWhere := aWhere +',3';
  if dtmMain.GetStarCMSend then aWhere := aWhere +',4';
  adoInv040Code.SQL.Text := adoInv040Code.SQL.Text +
    ' WHERE A.SENDTYPE IN (' + aWhere + ') order by to_number(compid)';
  adoInv040Code.Open;
  Result := adoInv040Code.RecordCount;
end;

procedure TdtmMainJ.SynchronizeCD110;
var
  aSoDataReader: TADOQuery;
  aWriter: TADOQuery;
  aEffects: Integer;
begin

  aSoDataReader := dtmSO.adoQrySynCD110;
  aWriter := dtmMain.adoComm2;
  {}
  aSoDataReader.Close;
  aSoDataReader.SQL.Text := ' SELECT * FROM CD110 ORDER BY CodeNo ';
  aSoDataReader.Open;
  aSoDataReader.First;
  while not aSoDataReader.Eof do
  begin
    try
      aWriter.Close;
      aWriter.SQL.Text := Format(
        ' SELECT COUNT(1) FROM INV041    ' +
        '  WHERE IDENTIFYID1 = ''%S''    ' +
        '    AND IDENTIFYID2 = ''%S''    ' +
        '    AND COMPID = ''%S''         ' +
        '    AND ITEMID = ''%S''         ',
        [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID,
         aSoDataReader.FieldByName( 'CODENO' ).AsString] );
      aWriter.Open;
      aEffects := aWriter.Fields[0].AsInteger;
      aWriter.Close;
      if ( aEffects > 0 ) then
      begin
        aWriter.SQL.Text := Format(
          ' UPDATE INV041                 ' +
          '    SET DESCRIPTION = ''%s'',  ' +
          '        REFNO = ''%s'',        ' +
          '        UPDEN = ''%s'',        ' +
          '        UPDTIME = SYSDATE,     ' +
          '        ISSELFCREATED = ''N''  ' +
          '  WHERE IDENTIFYID1 = ''%s''   ' +
          '    AND IDENTIFYID2 = ''%s''   ' +
          '    AND COMPID = ''%s''        ' +
          '    AND ITEMID = ''%s''        ',
          [aSoDataReader.FieldByName( 'Description' ).AsString,
          aSoDataReader.FieldByName( 'RefNo' ).AsString,
          dtmMain.getLoginUser,
          IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID,
          aSoDataReader.FieldByName( 'CODENO' ).AsString] );
        aWriter.ExecSQL;
      end else
      begin
        aWriter.SQL.Text := Format(
          ' INSERT INTO INV041 ( IDENTIFYID1, IDENTIFYID2,  ' +
          '    COMPID, ITEMID, DESCRIPTION, REFNO, UPDEN,   ' +
          '    UPDTIME, ISSELFCREATED )                     ' +
          '  VALUES ( ''%s'', ''%s'', ''%s'', ''%s'',       ' +
          '           ''%s'', ''%s'', ''%s'',               ' +
          '           SYSDATE,                              ' +
          '           ''N'' )                               ',
          [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID,
           aSoDataReader.FieldByName( 'CODENO' ).AsString,
           aSoDataReader.FieldByName( 'DESCRIPTION' ).AsString,
           aSoDataReader.FieldByName( 'REFNO' ).AsString,
           dtmMain.getLoginUser] );
        aWriter.ExecSQL;
      end;
    except
      on E: Exception do
      begin
        ErrorMsg( Format( '?P?B?????????N?X????, ?T??:%s?C', [E.Message] ) );
        Break;
      end;
    end;
    aSoDataReader.Next;
  end;
  aSoDataReader.Close;
  aWriter.Close;
  aWriter.SQL.Clear;

end;

function TdtmMainJ.AuthorizeElectronInv(const aInvDateSt, aInvDateEd,
  aInvIdSt, aInvIdEd: String;aAllComp : Boolean): String;
  var aSql,aWhere : String;
  aRunCnt : Integer;
begin
  aRunCnt := 0;
  aWhere := EmptyStr;
  {
  aSql := Format(
      '  UPDATE /*+ RULE */ INV007                    ' +
      '     SET UPLOADFLAG = 0,                       ' +
      '     UPLOADTIME = NULL                         ' +
      '   WHERE IDENTIFYID1 = ''%s''                  ' +
      '     AND IDENTIFYID2 = %s                      ' +
      '     AND UPLOADFLAG = 1                        ' +
      '     AND INVOICEKIND = 1                       ' +
      '     AND UPLOADTIME IS NOT NULL                ' ,
      [IDENTIFYID1, IDENTIFYID2] );
   }
  aWhere := Format( ' IDENTIFYID1 = ''%s''               ' +
                    '     AND IDENTIFYID2 = %s           ' +
                    '     AND UPLOADFLAG = 1             ' +
                    '     AND INVOICEKIND = 1            ' +
                    '     AND UPLOADTIME IS NOT NULL     ',
                    [IDENTIFYID1,IDENTIFYID2]);
  if  ( not aAllComp ) then
  begin
    aWhere := aWhere + Format( ' AND COMPID = ''%s''',[dtmMain.getCompID]);
  end;
  if aInvDateSt <> EmptyStr then
  begin
    aWhere := aWhere + Format( ' AND INVDATE >= TO_DATE(''%s'',''yyyymmdd'')    ',
                    [StringReplace(aInvDateSt,'/','',[rfReplaceAll])]);


  end;
  if aInvDateEd <> EmptyStr then
  begin
    aWhere := aWhere + Format( ' AND INVDATE <= TO_DATE(''%s'',''yyyymmdd'')     ',
                    [StringReplace(aInvDateEd,'/','',[rfReplaceAll])]);

  end;
  if aInvIdSt <> EmptyStr then
  begin
    aWhere := aWhere + Format( ' AND INVID >= ''%s''' ,[ aInvIdSt ]);

  end;
  if aInvIdEd <> EmptyStr then
  begin
    aWhere := aWhere + Format( ' AND INVID <= ''%s''' ,[ aInvIdEd ]);

  end;
  aSql := Format( '  UPDATE  INV007                    ' +
                  '     SET UPLOADFLAG = 0,            ' +
                  '     UPLOADTIME = NULL              ' +
                  '   WHERE %s                         ',
                  [aWhere] );

  adoQryCommon.Close;
  adoQryCommon.SQL.Clear;
  adoQryCommon.SQL.Text := aSql;
  aRunCnt := adoQryCommon.ExecSQL;
  if aRunCnt = 0 then
  begin
    Result := '?L???????s???v?????I';
  end else
  begin
    Result := Format('?w???s???v %d ???????I',[aRunCnt]);
  end;


end;

function TdtmMainJ.AuthorizeDiscountInv(const aDiscDateSt,
  aDiscDateEd,aDiscNoSt,aDiscNoEd: string; aAllComp: Boolean): String;
  var aSql,aWhere : String;
  aRunCnt : Integer;
begin
  aWhere := Format( ' A.IDENTIFYID1 = ''%s''               ' +
                    '     AND A.IDENTIFYID2 = %s           ' +
                    '     AND A.UPLOADFLAG = 1             ' +
                    '     AND B.INVOICEKIND = 1            ' +
                    '     AND A.UPLOADTIME IS NOT NULL     ',
                    [IDENTIFYID1,IDENTIFYID2]);
  if  ( not aAllComp ) then
  begin
    aWhere := aWhere + Format( ' AND A.COMPID = ''%s''',[dtmMain.getCompID]);
  end;
  if aDiscDateSt <> EmptyStr then
  begin
    aWhere := aWhere + Format( ' AND A.PaperDate >= TO_DATE(''%s'',''yyyymmdd'')    ',
                    [StringReplace(aDiscDateSt,'/','',[rfReplaceAll])]);

  end;
  if aDiscDateEd <> EmptyStr then
  begin
    aWhere := aWhere + Format( ' AND A.PaperDate <= TO_DATE(''%s'',''yyyymmdd'')    ',
                    [StringReplace(aDiscDateEd,'/','',[rfReplaceAll])]);

  end;
  {#6001 ?W?[?????????_?? By Kin 2011/07/13}
  if aDiscNoSt <> EmptyStr then
  begin
    aWhere := aWhere + Format( ' AND A.ALLOWANCENO >= ''%s''',[aDiscNoSt]);
  end;
  if aDiscNoEd <> EmptyStr then
  begin
    aWhere := aWhere + Format( ' AND A.ALLOWANCENO <= ''%s''',[aDiscNoEd]);
  end;

  aSql := 'SELECT ALLOWANCENO FROM INV014 A,INV007 B WHERE A.INVID = B.INVID     ' +
          ' AND A.BUSINESSID IS NULL AND ' + aWhere;
  aSql := Format( '  UPDATE  INV014                    ' +
                  '     SET UPLOADFLAG = 0,            ' +
                  '     UPLOADTIME = NULL              ' +
                  '   WHERE ALLOWANCENO IN ( %s )      ',
                  [aSql] );
  adoQryCommon.Close;
  adoQryCommon.SQL.Clear;
  adoQryCommon.SQL.Text := aSql;
  aRunCnt := adoQryCommon.ExecSQL;
  if aRunCnt = 0 then
  begin
    Result := '?L???????s???v?????I';
  end else
  begin
    Result := Format('?w???s???v %d ???????I',[aRunCnt]);
  end;

end;

function TdtmMainJ.CanDropOrRecoverDis(var aParam: TCommonParam): Boolean;
  var
  aYearMonth: String;
  aSoReader: TADOQuery;
  aCount: Integer;
begin
  { aParam.InStr1 = ?o?????X
    aParam.InStr2 = Y/N ?? ???r?? --> ?@?o/?^?_
    aParam.OutStr1 = ?o??????
    aParam.OutStr2 = ???o???O?_?w?Q?@?o --> Y or N }
  Result := False;
  adoQryCommon.Close;
  adoQryCommon.SQL.Text := Format(
     ' SELECT INVDATE, ISOBSOLETE, UPLOADFLAG            ' +
     '   FROM INV014                                     ' +
     '  WHERE IDENTIFYID1 = ''%s''                       ' +
     '    AND IDENTIFYID2 = ''%s''                       ' +
     '    AND COMPID = ''%s''                            ' +
     '    AND ALLOWANCENO = ''%s''           ',
     [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, aParam.InStr1] );
  adoQryCommon.Open;
  if ( not adoQryCommon.IsEmpty ) then
  begin
    aParam.OutStr1 := adoQryCommon.FieldByName( 'INVDATE' ).AsString;
    aParam.OutStr2 := adoQryCommon.FieldByName( 'ISOBSOLETE' ).AsString;
    aParam.OutStr3 := adoQryCommon.FieldByName( 'UPLOADFLAG' ).AsString;
  end;
  adoQryCommon.Close;
  aParam.Msg := EmptyStr;
  try
    if ( aParam.OutStr1 = EmptyStr ) then
    begin
      aParam.Msg := Format( '?????????d?L???i????????:%s?C', [aParam.InStr1] );
      Exit;
    end;
    if ( aParam.InStr2 <> EmptyStr ) and ( aParam.OutStr2 = 'Y' ) then
    begin
      aParam.Msg := Format( '???i??????:%s, ?w?Q?@?o,???i?A?Q?@?o?C',
        [aParam.InStr1] );
      Exit;
    end else
    if ( aParam.InStr2 = EmptyStr ) and ( Nvl( aParam.OutStr2, 'N' ) = 'N' ) then
    begin
      aParam.Msg := Format( '???i??????:%s, ?|???@?o,???i?^?_?C',
        [aParam.InStr1] );
      Exit;
    end;
    if (aParam.InStr2 <> EmptyStr ) and ( aParam.OutStr3 <> '1' ) then
    begin
      aParam.Msg := Format( '???i??????:%s, ?|???W??,???i?Q?@?o?C',
        [aParam.InStr1] );
      Exit;
    end;

  finally
    Result := ( aParam.Msg = EmptyStr );
  end;


end;

function TdtmMainJ.DropOrRecoverDis(var aParam: TCommonParam): Boolean;
var
  aSql: String;
begin

  { aParam.InStr1 = ?o?????X
    aParam.InStr2 = ???o?N?X
    aParam.InStr3 = ???o???]
    aParam.OutStr1 = HowToCreate
    aParam.OutStr2 = ?o??????
    }
  adoQryCommon.Close;
  adoQryCommon.SQL.Text := Format(
    '  SELECT HOWTOCREATE,                ' +
    '         INVDATE                     ' +
    '    FROM INV007                      ' +
    '   WHERE IDENTIFYID1 = ''%s''        ' +
    '     AND IDENTIFYID2 = %s            ' +
    '     AND COMPID = ''%s''             ' +
    '     AND INVID = ''%s''              ',
    [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, aParam.InStr1] );
  adoQryCommon.Open;
  aParam.OutStr1 := adoQryCommon.FieldByName( 'HOWTOCREATE' ).AsString;
  aParam.OutStr2 := adoQryCommon.FieldByName( 'INVDATE' ).AsString;
  adoQryCommon.Close;
  { ?^?_ }
  if ( aParam.InStr2 = EmptyStr ) then
  begin
    aSql := Format(
     '  UPDATE INV014                      ' +
     '     SET ISOBSOLETE = ''N'',         ' +
     '         OBSOLETEID = NULL,          ' +
     '         OBSOLETEREASON = NULL,      ' +
     '         UPTTIME = SYSDATE,          ' +
     '         UPTEN = ''%s''              ' +
     '   WHERE IDENTIFYID1 = ''%s''        ' +
     '     AND IDENTIFYID2 = %s            ' +
     '     AND COMPID = ''%s''             ' +
     '     AND ALLOWANCENO = ''%s''        ',
     [dtmMain.getLoginUser, IDENTIFYID1, IDENTIFYID2,
      dtmMain.getCompID, aParam.InStr1] );
  end else
  begin
    aSql := Format(
     '  UPDATE INV014                      ' +
     '     SET ISOBSOLETE = ''Y'',         ' +
     '         OBSOLETEID = ''%s'',        ' +
     '         OBSOLETEREASON = ''%s'',    ' +
     '         UPTTIME = SYSDATE,          ' +
     '         UPTEN = ''%s''              ' +
     '   WHERE IDENTIFYID1 = ''%s''        ' +
     '     AND IDENTIFYID2 = ''%s''        ' +
     '     AND COMPID = ''%s''             ' +
     '     AND ALLOWANCENO = ''%s''        ',
     [ aParam.InStr2, aParam.InStr3, dtmMain.getLoginUser,
       IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, aParam.InStr1] );
  end;
  adoQryCommon.Close;
  adoQryCommon.SQL.Text := aSql;
  adoQryCommon.ExecSQL;
  Result := True;


end;

end.
