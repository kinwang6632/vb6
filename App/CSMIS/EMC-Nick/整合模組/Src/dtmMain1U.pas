
//Provider=MSDAORA.1;Password=V30;User ID=V30;Data Source=EMC;Persist Security Info=True
{
drop table so121;
drop table so122;
drop table so123;
drop table so124;
create table so121 as select * from t121;
create table so122 as select * from t122;
create table so123 as select * from t123;
create table so124 as select * from t124;
}
//Provider=MSDAORA.1;Password=v30;User ID=v30;Data Source=mis
unit dtmMain1U;

interface

uses
  SysUtils, Classes,Math ,ADODB ,IniFiles, DB ,Forms ,Dialogs, DBClient,StdCtrls ,Controls ,Variants,
  Provider, RecError,scExcelExport;

const
      //�R�_
    //BUY_INTR_COM    =500;        //���ФH
    //BUY_ACCNAME_COM =50;    //���z�H
    //BUY_WORNAME_COM =100;  //�u�{�H��
     //���βέp
    //RET_INTR_COM   =300;
    //RET_ACCNAME_COM=50;
    //RET_WORNAME_COM=100;

    BUY_BOX = 1;
    RENT_BOX= 2;

    //�אּ��SO120 �]�w MONTH_DAYS = 30;   //�@�Ӥ몺�Ѽ�
    THIS_MONTH = 1;   //����
    //�אּ��SO125 �]�w RENT_USE_DAYS = 90; //�������L�W�L�T�Ӥ�(90��),�W�L�~�����ФH����
    //CHANNEL_USE_DAYS = 30; //�W�D�������L�W�L�@�Ӥ�(30��),�W�L�~�����ФH����


    //CALL_IN_COM=50;

    CHANNEL_DATA = 'C';  //�I�O�W�D���
    CHANNEL_A = 'CA';    //�I�O�W�D�Ҧ�����
    CHANNEL_I = 'CI';    //�I�O�W�D���ФH���
    CHANNEL_C = 'CC';    //�I�O�W�D���O�����

    BOX_DATA = 'B';      //BOX���
    BOX_A = 'BA';        //BOX�Ҧ�����
    BOX_I = 'BI';        //BOX���ФH���
    BOX_W = 'BW';        //BOX�u�{�H�����

    STATISTIC_DATA = 'S';      //�έp����

    BATCH_PAY = '0';     //�������I����
    ONCE_PAY = '1';      //�@�����I����

    EXCEL_MODE = 'E';       //����Excel���
    REPORT_MODE = 'R';      //����Report���


    EMP_TYPE_WORKER = '0';  //�u�{�H����
    EMP_TYPE_CLCT = '1';    //���O�H����
    EMP_TYPE_INTACC = '2';  //���ФH��


    WORKER_NO = 999;
    WORKER = '�u�{�H��';

    SERVER_NO = 998;
    SERVER = '�ȪA�H��';

    CLCTEN_NO = 997;
    CLCTEN = '���O�H��';
type
  TCommissionIniData = class(TObject)
    nDataAreaNo : Integer;
    sDataArea : String;
    sAlias : String;
    sUserID : String;
    sPassword : String;
    sCompCode : String;
    sCompName : String;
  end;


  TdtmMain1 = class(TDataModule)
    qryCodeCD039: TADOQuery;
    qryCodeCD019: TADOQuery;
    qryCodeCD032: TADOQuery;
    qryCom: TADOQuery;
    qrySO122: TADOQuery;
    qrySo033So034: TADOQuery;
    qryCodeCD019CODENO: TIntegerField;
    qryCodeCD019DESCRIPTION: TStringField;
    qrySo120Formula: TADOQuery;
    qrySO120: TADOQuery;
    qrySO121Z: TADOQuery;
    cdsSo121Z: TClientDataSet;
    cdsSo122: TClientDataSet;
    qryCommon: TADOQuery;
    dspSo121Z: TDataSetProvider;
    dspSo122: TDataSetProvider;
    qryCodeCD009: TADOQuery;
    qrySo123: TADOQuery;
    cdsSo123: TClientDataSet;
    dspSo123: TDataSetProvider;
    cdsSo122ReturnMoney: TClientDataSet;
    qrySo122ReturnMoney: TADOQuery;
    dspSo122ReturnMoney: TDataSetProvider;
    cdsSo121ReturnMoney: TClientDataSet;
    qryCodeCD042: TADOQuery;
    qryCD042: TADOQuery;
    qrySO120COMPCODE: TIntegerField;
    qrySO120CODENO: TIntegerField;
    qrySO120DESCRIPTION: TStringField;
    qrySO120PROMOTECODE: TBCDField;
    qrySO120PROMOTENAME: TStringField;
    qrySO120DISCOUNTCODE: TBCDField;
    qrySO120DISCOUNTNAME: TStringField;
    qrySO120PAYUNIT: TStringField;
    qrySO120FIRSTCREDITCARD1: TBCDField;
    qrySO120FIRSTNOTCREDITCARD1: TBCDField;
    qrySO120FIRSTCREDITCARD2: TBCDField;
    qrySO120FIRSTNOTCREDITCARD2: TBCDField;
    qrySO120OTHERCREDITCARD2: TBCDField;
    qrySO120OTHERNOTCREDITCARD2: TBCDField;
    qrySO120REF_NO: TIntegerField;
    qrySO120OPERATOR: TStringField;
    qrySO120UPDTIME: TStringField;
    qrySO120CompName: TStringField;
    qrySO120PAYUNITNAME: TStringField;
    qryBox: TADOQuery;
    cdsSo125: TClientDataSet;
    qrySO125: TADOQuery;
    dspSO125: TDataSetProvider;
    qrySO124: TADOQuery;
    qryCM003: TADOQuery;
    qrySO013: TADOQuery;
    qryOtherCompCode: TADOQuery;
    cdsSo121ReturnMoneyA: TClientDataSet;
    IntegerField1: TIntegerField;
    StringField1: TStringField;
    StringField2: TStringField;
    StringField3: TStringField;
    StringField4: TStringField;
    BCDField1: TBCDField;
    StringField5: TStringField;
    StringField6: TStringField;
    cdsSo121ReturnMoneyCOMPCODE: TIntegerField;
    cdsSo121ReturnMoneyCOMPUTEYM: TStringField;
    cdsSo121ReturnMoneyBELONGYM: TStringField;
    cdsSo121ReturnMoneyEMPID: TStringField;
    cdsSo121ReturnMoneyEMPNAME: TStringField;
    cdsSo121ReturnMoneyCOMMISSION: TBCDField;
    cdsSo121ReturnMoneyOPERATOR: TStringField;
    cdsSo121ReturnMoneyUPDTIME: TStringField;
    cdsSo121ReturnMoneyMEDIACODE: TIntegerField;
    cdsSo121ReturnMoneyMEDIANAME: TStringField;
    cdsSo121ReturnMoneyCOMPTYPE: TStringField;
    cdsSo121ReturnMoneyCLANCOMPCODE: TIntegerField;
    cdsSo121ReturnMoneyCLANCOMPNAME: TStringField;
    cdsSo121ZCOMPCODE: TIntegerField;
    cdsSo121ZCOMPUTEYM: TStringField;
    cdsSo121ZBELONGYM: TStringField;
    cdsSo121ZEMPID: TStringField;
    cdsSo121ZEMPNAME: TStringField;
    cdsSo121ZCOMMISSION: TBCDField;
    cdsSo121ZMEDIACODE: TIntegerField;
    cdsSo121ZMEDIANAME: TStringField;
    cdsSo121ZCOMPTYPE: TStringField;
    cdsSo121ZCLANCOMPCODE: TIntegerField;
    cdsSo121ZCLANCOMPNAME: TStringField;
    cdsSo121ZOPERATOR: TStringField;
    cdsSo121ZUPDTIME: TStringField;
    qrySO121: TADOQuery;
    dspSo121: TDataSetProvider;
    cdsSo121: TClientDataSet;
    cdsSo121COMPCODE: TIntegerField;
    cdsSo121COMPUTEYM: TStringField;
    cdsSo121BELONGYM: TStringField;
    cdsSo121EMPID: TStringField;
    cdsSo121EMPNAME: TStringField;
    cdsSo121COMMISSION: TBCDField;
    cdsSo121MEDIACODE: TIntegerField;
    cdsSo121MEDIANAME: TStringField;
    cdsSo121COMPTYPE: TStringField;
    cdsSo121CLANCOMPCODE: TIntegerField;
    cdsSo121CLANCOMPNAME: TStringField;
    cdsSo121OPERATOR: TStringField;
    cdsSo121UPDTIME: TStringField;
    ADOQuery1: TADOQuery;
    ADOQuery1COMPCODE: TIntegerField;
    ADOQuery1COMPUTEYM: TStringField;
    ADOQuery1BELONGYM: TStringField;
    ADOQuery1EMPID: TStringField;
    ADOQuery1EMPNAME: TStringField;
    ADOQuery1COMMISSION: TBCDField;
    ADOQuery1MEDIACODE: TIntegerField;
    ADOQuery1MEDIANAME: TStringField;
    ADOQuery1COMPTYPE: TStringField;
    ADOQuery1CLANCOMPCODE: TIntegerField;
    ADOQuery1CLANCOMPNAME: TStringField;
    ADOQuery1OPERATOR: TStringField;
    ADOQuery1UPDTIME: TStringField;
    qrySO120CHANNELVIEWDAYS: TIntegerField;
    qryPriv: TADOQuery;
    qryCD019: TADOQuery;
    qrySO122Data: TADOQuery;
    dspSO122Data: TDataSetProvider;
    cdsSO122BOX: TClientDataSet;
    cdsSO122BOXCOMPCODE: TIntegerField;
    cdsSO122BOXCOMPUTEYM: TStringField;
    cdsSO122BOXBELONGYM: TStringField;
    cdsSO122BOXSTBNO: TStringField;
    cdsSO122BOXSNO: TStringField;
    cdsSO122BOXWORKEREN1: TStringField;
    cdsSO122BOXWORKEREN1NAME: TStringField;
    cdsSO122BOXINTACCID: TStringField;
    cdsSO122BOXINTACCNAME: TStringField;
    cdsSO122BOXSTAFFCODE: TStringField;
    cdsSO122BOXREALSTARTDATE: TDateTimeField;
    cdsSO122BOXREALSTOPDATE: TDateTimeField;
    cdsSO122BOXOPERATOR: TStringField;
    cdsSO122BOXUPDTIME: TStringField;
    cdsSO122BOXBUYORRENT: TIntegerField;
    cdsSO122BOXMEDIACODE: TIntegerField;
    cdsSO122BOXMEDIANAME: TStringField;
    qryCD042Data: TADOQuery;
    qrySO105Prom: TADOQuery;
    qryDateData: TADOQuery;
    cdsSO122Channel: TClientDataSet;
    cdsSO122ChannelCOMPCODE: TIntegerField;
    cdsSO122ChannelCOMPUTEYM: TStringField;
    cdsSO122ChannelBELONGYM: TStringField;
    cdsSO122ChannelBILLNO: TStringField;
    cdsSO122ChannelITEM: TIntegerField;
    cdsSO122ChannelPERIODID: TIntegerField;
    cdsSO122ChannelCLCTEN: TStringField;
    cdsSO122ChannelCLCTENNAME: TStringField;
    cdsSO122ChannelINTACCID: TStringField;
    cdsSO122ChannelINTACCNAME: TStringField;
    cdsSO122ChannelSTAFFCODE: TStringField;
    cdsSO122ChannelREALSTARTDATE: TDateTimeField;
    cdsSO122ChannelREALSTOPDATE: TDateTimeField;
    cdsSO122ChannelOPERATOR: TStringField;
    cdsSO122ChannelUPDTIME: TStringField;
    cdsSO122ChannelCITEMCODE: TIntegerField;
    cdsSO122ChannelCITEMNAME: TStringField;
    cdsSO122ChannelMEDIACODE: TIntegerField;
    cdsSO122ChannelMEDIANAME: TStringField;
    cdsSO122ChannelREALDATE: TDateField;
    cdsSO122ChannelSHOULDDATE: TDateField;
    qrySO132: TADOQuery;
    qrySO132COMPCODE: TIntegerField;
    qrySO132CUSTID: TIntegerField;
    qrySO132CITEMCODE: TIntegerField;
    qrySO132CITEMNAME: TStringField;
    qrySO132EMPNO: TStringField;
    qrySO132EMPNAME: TStringField;
    qrySO132DATACREATOR: TStringField;
    qrySO132UPDEN: TStringField;
    qrySO132UPDTIME: TDateTimeField;
    qrySo133: TADOQuery;
    qrySo133SEQ: TBCDField;
    qrySo133COMPCODE: TIntegerField;
    qrySo133OLDCITEMCODE: TIntegerField;
    qrySo133OLDCITEMNAME: TStringField;
    qrySo133NEWCITEMCODE: TIntegerField;
    qrySo133NEWCITEMNAME: TStringField;
    qrySo133OLDEMPNO: TStringField;
    qrySo133OLDEMPNAME: TStringField;
    qrySo133NEWEMPNO: TStringField;
    qrySo133NEWEMPNAME: TStringField;
    qrySo133OPERATIONMODE: TStringField;
    qrySo133UPDEN: TStringField;
    qrySo133UPDTIME: TDateTimeField;
    ADOQuery2: TADOQuery;
    qryCodeCM003: TADOQuery;
    qryCm003EMPNO: TStringField;
    qryCm003EMPNAME: TStringField;
    qryCodeCD019_2: TADOQuery;
    qryCD019CODENO: TIntegerField;
    qryCD019DESCRIPTION: TStringField;
    ADOCommand1: TADOCommand;
    cdsSo033Log: TClientDataSet;
    cdsSo033LogCUSTID: TStringField;
    cdsSo033LogBILLNO: TStringField;
    cdsSo033LogITEM: TStringField;
    cdsSo033LogCITEMCODE: TStringField;
    cdsSo033LogCITEMNAME: TStringField;
    cdsSo033LogSHOULDDATE: TDateField;
    cdsSo033LogSHOULDAMT: TStringField;
    cdsSo033LogREALDATE: TDateField;
    cdsSo033LogREALAMT: TStringField;
    cdsSo033LogREALPERIOD: TStringField;
    cdsSo033LogREALSTARTDATE: TDateField;
    cdsSO122ChannelCLCTENORIPERCENT: TFloatField;
    cdsSO122ChannelINTACCCOMM: TFloatField;
    cdsSO122ChannelINTACCORIPERCENT: TFloatField;
    cdsSO122ChannelREALAMT: TFloatField;
    cdsSO122BOXWORKEREN1COMM: TBCDField;
    cdsSO122BOXINTACCCOMM: TIntegerField;
    cdsSO122ChannelCLCTENCOMM: TFloatField;
    cdsSO122Data: TClientDataSet;
    cdsSo033LogSBILLNO: TStringField;
    cdsSo033LogSITEM: TIntegerField;
    cdsSO122ChannelPROMCODE: TBCDField;
    cdsSO122BOXPROMCODE: TBCDField;
    cdsSO122BOXPROMNAME: TStringField;
    cdsSO122ChannelPROMNAME: TStringField;
    cdsSO122BOXBuyName: TStringField;
    cdsSO122BOXCustID: TIntegerField;
    cdsSO122ChannelCustID: TIntegerField;
    cdsSo033LogNotes: TStringField;
    cdsSo122COMPCODE: TIntegerField;
    cdsSo122COMPUTEYM: TStringField;
    cdsSo122BELONGYM: TStringField;
    cdsSo122STBNO: TStringField;
    cdsSo122SNO: TStringField;
    cdsSo122BILLNO: TStringField;
    cdsSo122ITEM: TIntegerField;
    cdsSo122PERIODID: TIntegerField;
    cdsSo122WORKEREN1: TStringField;
    cdsSo122WORKEREN1NAME: TStringField;
    cdsSo122WORKEREN1COMM: TBCDField;
    cdsSo122CLCTEN: TStringField;
    cdsSo122CLCTENNAME: TStringField;
    cdsSo122CLCTENCOMM: TBCDField;
    cdsSo122CLCTENORIPERCENT: TBCDField;
    cdsSo122INTACCID: TStringField;
    cdsSo122INTACCNAME: TStringField;
    cdsSo122INTACCCOMM: TBCDField;
    cdsSo122INTACCORIPERCENT: TBCDField;
    cdsSo122STAFFCODE: TStringField;
    cdsSo122REALSTARTDATE: TDateTimeField;
    cdsSo122REALSTOPDATE: TDateTimeField;
    cdsSo122OPERATOR: TStringField;
    cdsSo122UPDTIME: TStringField;
    cdsSo122BUYORRENT: TIntegerField;
    cdsSo122CITEMCODE: TIntegerField;
    cdsSo122CITEMNAME: TStringField;
    cdsSo122MEDIACODE: TIntegerField;
    cdsSo122MEDIANAME: TStringField;
    cdsSo122REALAMT: TIntegerField;
    cdsSo122PROMCODE: TBCDField;
    cdsSo122PROMNAME: TStringField;
    cdsSo122REALDATE: TDateTimeField;
    cdsSo122SHOULDDATE: TDateTimeField;
    cdsSo122CUSTID: TIntegerField;
    cdsSo122CUSTNAME: TStringField;
    dspSo033So034: TDataSetProvider;
    cdsSo033So034: TClientDataSet;
    qryTempBox: TADOQuery;
    dspTempBox: TDataSetProvider;
    cdsTempBox: TClientDataSet;
    qrySO004: TADOQuery;
    cdsSo033LogSNo: TStringField;
    cdsSo033LogSEQNo: TStringField;
    cdsSo033LogPRSNO: TStringField;
    cdsSo033LogInstDate: TDateField;
    cdsSo033LogPRDate: TDateField;
    cdsSo033LogPromCode: TBCDField;
    cdsSo033LogPromName: TStringField;
    cdsSo033LogDiscountCode: TIntegerField;
    cdsSo033LogDiscountName: TStringField;
    qrySO134: TADOQuery;
    dspSO134: TDataSetProvider;
    cdsSO134: TClientDataSet;
    cdsSo122FIRSTFLAG: TStringField;
    cdsSO122ChannelFirstFlagName: TStringField;
    spGetChargeData: TADOStoredProc;
    dspGetChargeData: TDataSetProvider;
    cdsGetChargeData: TClientDataSet;
    qrySO135: TADOQuery;
    dspSO135: TDataSetProvider;
    cdsSO135: TClientDataSet;
    ClientDataSet1: TClientDataSet;
    IntegerField2: TIntegerField;
    StringField7: TStringField;
    StringField8: TStringField;
    StringField9: TStringField;
    StringField10: TStringField;
    StringField11: TStringField;
    IntegerField3: TIntegerField;
    IntegerField4: TIntegerField;
    StringField12: TStringField;
    StringField13: TStringField;
    BCDField2: TBCDField;
    StringField14: TStringField;
    StringField15: TStringField;
    StringField16: TStringField;
    BCDField3: TBCDField;
    StringField17: TStringField;
    BCDField4: TBCDField;
    StringField18: TStringField;
    StringField19: TStringField;
    BCDField5: TBCDField;
    StringField20: TStringField;
    BCDField6: TBCDField;
    StringField21: TStringField;
    DateTimeField1: TDateTimeField;
    DateTimeField2: TDateTimeField;
    StringField22: TStringField;
    StringField23: TStringField;
    IntegerField5: TIntegerField;
    IntegerField6: TIntegerField;
    StringField24: TStringField;
    IntegerField7: TIntegerField;
    StringField25: TStringField;
    IntegerField8: TIntegerField;
    StringField26: TStringField;
    IntegerField9: TBCDField;
    StringField27: TStringField;
    DateTimeField3: TDateTimeField;
    DateTimeField4: TDateTimeField;
    IntegerField10: TIntegerField;
    StringField28: TStringField;
    qrySO134Data: TADOQuery;
    dspSO134Data: TDataSetProvider;
    cdsSO134Data: TClientDataSet;
    cdsStatisticExcel: TClientDataSet;
    cdsStatisticExcelCompCode: TStringField;
    cdsStatisticExcelEmpID: TStringField;
    cdsStatisticExcelEmpName: TStringField;
    cdsStatisticExcelWorkerCounts: TIntegerField;
    cdsStatisticExcelWorkerComm: TFloatField;
    cdsStatisticExcelClctCounts: TIntegerField;
    cdsStatisticExcelClctComm: TFloatField;
    cdsStatisticExcelTotalComm: TFloatField;
    cdsStatisticExcelStrWComm: TStringField;
    cdsStatisticExcelStrCComm: TStringField;
    cdsTotalCommData: TClientDataSet;
    cdsTotalCommDataLastPay: TFloatField;
    cdsTotalCommDataNextPay: TFloatField;
    cdsTotalCommDataMonth1: TFloatField;
    cdsTotalCommDataMonth2: TFloatField;
    cdsTotalCommDataMonth3: TFloatField;
    cdsTotalCommDataMonth4: TFloatField;
    cdsTotalCommDataMonth5: TFloatField;
    cdsTotalCommDataMonth6: TFloatField;
    cdsTotalCommDataMonth7: TFloatField;
    cdsTotalCommDataMonth8: TFloatField;
    cdsTotalCommDataMonth9: TFloatField;
    cdsTotalCommDataMonth10: TFloatField;
    cdsTotalCommDataMonth11: TFloatField;
    cdsTotalCommDataMonth12: TFloatField;
    cdsTotalCommDataOtherMonth: TFloatField;
    qryTotalComm: TADOQuery;
    dspTotalComm: TDataSetProvider;
    cdsTotalComm: TClientDataSet;
    cdsTotalCommDataMode: TStringField;
    cdsTotalCommDataModeName: TStringField;
    cdsTotalMonthData: TClientDataSet;
    cdsTotalMonthDataComputeYM: TStringField;
    cdsTotalMonthDataMonth1: TFloatField;
    cdsTotalMonthDataMonth2: TFloatField;
    cdsTotalMonthDataMonth3: TFloatField;
    cdsTotalMonthDataMonth4: TFloatField;
    cdsTotalMonthDataMonth5: TFloatField;
    cdsTotalMonthDataMonth6: TFloatField;
    cdsTotalMonthDataMonth7: TFloatField;
    cdsTotalMonthDataMonth8: TFloatField;
    cdsTotalMonthDataMonth9: TFloatField;
    cdsTotalMonthDataMonth10: TFloatField;
    cdsTotalMonthDataMonth11: TFloatField;
    cdsTotalMonthDataMonth12: TFloatField;
    cdsTotalMonthDataOtherMonth: TFloatField;
    cdsTotalMonthDataTotalComm: TFloatField;
    cdsStatisticExcelBoxIntAccCounts: TIntegerField;
    cdsStatisticExcelBoxIntAccComm: TFloatField;
    cdsStatisticExcelBoxStrIComm: TStringField;
    cdsStatisticExcelChIntAccCounts: TIntegerField;
    cdsStatisticExcelChIntAccComm: TFloatField;
    cdsStatisticExcelChStrIComm: TStringField;
    qryComm1: TADOQuery;
    dspComm1: TDataSetProvider;
    cdsComm1: TClientDataSet;
    cdsSo122ORDERNO: TStringField;
    cdsSo122CALLSELF: TIntegerField;
    cdsSO122ChannelREALSTARTDATE2: TStringField;
    cdsSO122ChannelREALSTOPDATE2: TStringField;
    cdsSO122ChannelREALDATE2: TStringField;
    cdsSO122ChannelSHOULDDATE2: TStringField;
    cdsSO122BOXREALSTARTDATE2: TStringField;
    cdsSO122BOXREALSTOPDATE2: TStringField;
    cdsSO122BOXREALDATE2: TDateField;
    cdsSO122BOXSHOULDDATE2: TDateField;
    cdsSO122BOXREALDATE22: TStringField;
    cdsSO122BOXSHOULDDATE22: TStringField;
    procedure DataModuleCreate(Sender: TObject);
    procedure qrySO120CalcFields(DataSet: TDataSet);
    procedure cdsSo121ZReconcileError(DataSet: TCustomClientDataSet;
      E: EReconcileError; UpdateKind: TUpdateKind;
      var Action: TReconcileAction);
    procedure qrySO132BeforePost(DataSet: TDataSet);
    procedure qrySO132BeforeDelete(DataSet: TDataSet);


    procedure cdsSO122BOXCalcFields(DataSet: TDataSet);
  private
    { Private declarations }
    G_DbConnAry : array of TADOConnection;
    G_AdoQueryAry : array of TADOQuery;
    sG_IntroId, sG_WorkerId, sG_WorkerName, sG_AcceptId : String;

    G_DbAreaNameStrList,G_DbAliasStrList,G_DbUserNameStrList,G_DbPasswordStrList : TStringList;

    nG_BuyBoxIntrComm,nG_RentBoxIntrComm : Integer;
    nG_BuyBoxAcceptComm,nG_RentBoxAcceptComm : Integer;
    nG_SetBoxWorkerComm,nG_SelfOrderChannelAcceptComm,nG_RentBoxUseDays : Integer;

    fG_TotalLastPay,fG_TotalNextPay,fG_TotalOtherMonth : Double;
    fG_TotalMonth1,fG_TotalMonth2,fG_TotalMonth3 : Double;
    fG_TotalMonth4,fG_TotalMonth5,fG_TotalMonth6 : Double;
    fG_TotalMonth7,fG_TotalMonth8,fG_TotalMonth9 : Double;
    fG_TotalMonth10,fG_TotalMonth11,fG_TotalMonth12 : Double;
    procedure createDbInfoStrList;
    function getEmpData(sI_CompCode,sI_CustID,sI_CitemCode : String;var sI_EmpNo,sI_EmpName : String) : Boolean;
    procedure computeTotalComm(I_CDS : TClientDataSet);

    //procedure cdsPerComsaveToSo121;
  public
    { Public declarations }
    //sG_ServiceType,sG_CurDate,sG_IsSupervisor,sG_UserID,sG_User,sG_CompCode,sG_CompName : String;
    //sG_DbUserName,sG_DbPassword,sG_DbAlias,sG_Server_IP : String;
    sG_CurDate : String;


    G_StrList : TStringList;
    G_OrderStrList : TStringList;
    sG_IsHeadQuarters,sG_SO043Para3 : String;
    sG_ExcelSavePath : String;


    sG_SFIntAccId,sG_SFIntAccName,sG_SFStaffCode : String;
    sG_SFCitemCode,sG_SFCitemName,sG_SFMediaCode,sG_SFMediaName : String;
    sG_SFPromCode,sG_SFPromName,sG_SFCustId,sG_SFPeriod,sG_SFFirstFlag : String;
    sG_SFOrderNo,sG_SFCallSelf : String;
    dG_SFRealStartDate,dG_SFRealStopDate : TDate;
    fG_SFIntAccOriPercent : Double;

    sG_ComputeYM,sG_BelongYM ,sG_SumSo121Commision ,sG_CharMeth,sG_CustId:String;
    sG_AccEn,sG_AccName, sG_CodeNo ,sG_PayUnit : String;
    fG_SpreadComm,fG_ChargeComm,fG_RealAmt : double;
    fG_FirCardSpd,fG_FirNoCardSpd : double;
    fG_FirCardCharge, fG_FirNoCardCharge,fG_ContCardCharge, fG_ContNoCardCharge : double;
    bG_Repeat : boolean;
    nG_RealPeriod ,nG_ChannelViewDays : Integer;
    nG_BuyBoxUseDays,nG_LimitPeriod,nG_BackMonth : Integer;
    sG_OnceComputeYM,sG_OnceBelongYM : String;
    sG_PromCodeAndName : String;
    G_PromCodeStrList: TStringList;
    G_PromNameStrList: TStringList;
    procedure copyCdsSo121ToQry;
    procedure save2Db(I_CDS:TClientDataSet);
    procedure saveSO121ToDB(I_CDS : TClientDataSet);
    procedure transReturnDataCdsIntoDB;//�N�h�O��Ƥ�CDS��X�� database ��
    procedure applyToDB;
    procedure activePriorData(sI_CompCode : String);
    procedure deleteCommData(sI_CompCode  :String);

    procedure dataSetActive;
    function chineseDateChangeToEnglishDate(sI_Date : String) : String;
    procedure saveToFile(vI_Content:OleVariant; sI_FileName:String);
    function checkIsSO120OnlyData(sI_CompCode, sI_CodeNo,sI_PromoteCode,sI_DiscountCode : String) : Boolean;
    function getCurCobRecID(I_Cob:TComboBox): String;
    function getCurCobRecName(I_Cob:TComboBox): String;
    procedure getSO120Data(sI_CompCode : String);
    function chineseYMChangeToEnglishYM(sI_YM : String) : String;
    function ReduceThreeMonths(sI_YM : String) : String;
    procedure dealwithSTBBuy(sI_YM : String);
    procedure dealwithSTBRent(sI_YM : String);
    procedure dealwithSTBMSno(sI_YM : String);
    procedure insertSO121(I_DateSet:TDataset; nI_OperationFlag:Integer);
    procedure insertSO122(I_DateSet:TDataset; nI_OperationFlag:Integer);
    procedure insertMSnoSO122(I_DateSet,I_MSnoDateSet:TDataset);
    function CheckTableSo121(sI_ComputeYM :String) :Boolean;

    function preProcessingData(sI_TableName,sI_YM,sI_PayMode : String) : String;


    function handleInCome(sI_TableName,sI_YM,sI_PayMode : String) : String;

    function handleOutCome(sI_TableName,sI_YM : String) : String;
    function handleOnceOutCome(sI_TableName,sI_YM : String) : String;

    function processIncome(sI_OrderNo,sI_RealDate:String; I_DataSet:TDataSet):Boolean;
    function processOnceIncome(sI_OrderNo,sI_RealDate:String; I_DataSet:TDataSet):Boolean;


    function CustIDCheck(sI_CustID ,sI_CompCode : String) : Boolean;
    function CustOnceIDCheck(sI_CustID ,sI_CompCode : String) : Boolean;

    function ifPayByCreditCard(sI_PTCode : String ) : Boolean;

    procedure SaveTempDataTocdsSo122(sI_RealBelongYM,sI_PromCode,sI_PromName,sI_ClctId, sI_ClctName, sI_IntroId, sI_IntroName ,sI_MediaCode,sI_MediaName,sI_RefNo,sI_OrderNo:STring; I_DataSet : TDataset ;nI_MonthDiff : Integer;dI_RealStartDate,dI_RealStopDate : TDate;fI_ClctEnOriPercent,fI_IntAccOriPercent : Double);
    procedure SaveTempDataTocdsSo134(sI_PromCode,sI_PromName,sI_ClctId, sI_ClctName, sI_IntroId, sI_IntroName ,sI_MediaCode,sI_MediaName,sI_RefNo:STring; I_DataSet : TDataset ;dI_RealStartDate,dI_RealStopDate : TDate;fI_ClctEnOriPercent,fI_IntAccOriPercent : Double);

    procedure SaveTempDataTocdsSo121(sI_ClctId, sI_ClctName, sI_IntroId, sI_IntroName ,sI_MediaCode,sI_MediaName,sI_RefNo:STring; I_DataSet:TDataset;nI_MonthDiff:Integer);
    procedure activeFormula;
    function insertSO123(sI_CompCode,sI_CustID, sI_AcceptEn,sI_AcceptName:String):boolean;
    function insertSO135(sI_CompCode,sI_CustID, sI_AcceptEn,sI_AcceptName:String):boolean;

    procedure insertCallInDataToSo121(sI_CompCode,sI_ComputeYM,sI_AcceptEn,sI_AcceptName : String);
    procedure insertCallInDataToSo122(bI_HasPayComm : Boolean;sI_RealBelongYM,sI_PromCode,sI_PromName,sI_CompCode,sI_ComputeYM,sI_AcceptEn,sI_AcceptName,sI_ClctId,sI_ClctName,sI_OrderNo : String; I_DataSet:TDataSet;fI_IntAccOriPercent,fI_ClctEnOriPercent,fI_RealAmt : Double);
    procedure insertCallInDataToSo134(bI_HasPayComm : Boolean;sI_PromCode,sI_PromName,sI_CompCode,sI_ComputeYM,sI_AcceptEn,sI_AcceptName,sI_ClctId,sI_ClctName : String; I_DataSet:TDataSet;fI_IntAccOriPercent : Double);

    //procedure selectCD009WhereRefNoEqualNight;
    function CalculateCountMonths(sI_BelongYM : String ; nI_RealPeriod :Integer) :String;
    function CheckSo124LockYM(sI_ComputeYM:String) :boolean;
    procedure getCD024Data(sI_PromoteCode : String;var sI_PromoteName,sI_DiscountCode,sI_DiscountName : String);
    function getPromoteCode(sI_OrderNo : String;var sI_PromoteName : String) : String;
    //function getCD009RefNo(sI_MediaCode : String) : String;
    procedure getComputeDays(sI_BelongYM,sI_RealStartDate,sI_RealStopDate : String;var nI_TotalPeriodDays,nI_FirstPeriodDays ,nI_PeriodCounts : Integer;var sI_RealBelongYM : String);
    function afterMonth(sI_YM : String;nI_Counts : Integer) : String;
    function beforeMonth(sI_YM : String;nI_Counts : Integer) : String;
    procedure ComputeReturnData1(sI_ShouldDate,sI_SBillNo,sI_SItem,sI_FunctionNo : String;fI_RealAmt : Double;I_CDSSo033So034 : TClientDataSet);
    function getSo125Param(sI_CompCode : String) : Boolean;
    function getSo125ParamTemp(sI_CompCode : String) : Boolean;
    function getLockYM : String;
    procedure updateLockYM(sI_NewLockYM : String);
    procedure InsertLockYM(sI_NewLockYM : String);
    procedure getSO121Data(sI_BelongYM,sI_Compute,sI_ManList : String);
    procedure getCM003Data;
    procedure getSO013Data;
    procedure getCD019Data;
    procedure getOtherCompCodeData;
    function getRptChannelDetailData(sI_PayType,sI_BelongYM,sI_Compute,sI_EmpNoSQL,sI_SalesCodeSQL,sI_OtherCompCodeSQL,sI_OrderList,sI_CitemCodeSQL : String) : Boolean;
    function getRptBoxDetailData(sI_BelongYM,sI_Compute,sI_EmpNoSQL,sI_SalesCodeSQL,sI_OtherCompCodeSQL,sI_OrderList : String) : Boolean;

    function getLastPeriodDays(dI_MonthFirstDay,dI_RealStopDate : TDate)  :Integer;
    procedure getCodeCD019;
    procedure getSO122Data1(sI_Flag,sI_BelongYM,sI_CompCode,sI_CODENO:String;bI_IncludeWorkerComm : Boolean);
    procedure getSO134Data(sI_BelongYM,sI_CompCode,sI_CODENO:String);

    function TransToExcel(sI_Flag,sI_BelongYM:String):String;
    function TransToChannelExcel(sI_PayMode,sI_Flag,sI_BelongYM,sI_CurrDateTime :String):String;
    function TransToBoxExcel(sI_Flag,sI_BelongYM,sI_CurrDateTime :String;bI_IncludeWorkerComm : Boolean):String;
    function TransToStatisticExcel(sI_BelongYM,sI_CurrDateTime :String;bI_IncludeWorkerComm : Boolean):String;

    procedure getAccountParameter(I_Adoqry:TADOQuery);
    procedure getCD042Data;
    function  CheckProm(sI_OrderNo,sI_PROMCODE,sI_PROMNAME:String):boolean;
    function  getDateData(sI_SEQNo,sI_ComputeYM,sI_Table:String;  Var dI_RealDate,dI_ShouldDate:TDate) : Boolean;
    function ReplaceStr(sI_SrcString, sI_SepFlag: String): String;
    procedure transSo033So034Notation(I_CDS : TClientDataSet);
    procedure transSo122Notation(I_CDS : TClientDataSet);
    procedure transSo134Notation(I_CDS : TClientDataSet);

    function getISno(sI_Sno : String;var sI_ReturnSno : String) : Integer;
    function checkedIsAlreadyPay(sI_YM,sI_Sno,sI_CompCode,sI_StbNo : String) : Boolean;
    function checkIsTheSameMonth(sI_Date1,sI_Date2 : String) : Boolean;
    function callSF_GetChargeData(sI_Table,sI_CompCode,sI_BelongYM,sI_BillNo,sI_Item : String;var sI_Msg : String;var nI_RetCode : Integer) : Boolean;
    procedure getLastData(sI_Table,sI_CompCode : String;var sI_LastUser,sI_LastComputeYM,sI_LastUpdTime : String);
    function getStatisticData(sI_ExcelOrRpt,sI_PayMode,sI_CompCode,sI_BelongYM,sI_CitemCodeStr : String;bI_IncludeWorkerComm : Boolean) : Boolean;
    procedure insertToCdsStatisticExcel(sI_BoxOrChannel,sI_EmpType,sI_CompCode,sI_BelongYM,sI_EmpID,sI_EmpName,sI_StrComm : String;nI_Counts : Integer;fI_Comm : Double);
    procedure getSFReturnData(sI_ReturnData : String);
    procedure computeExcelTotalComm(I_CDS : TClientDataSet;sI_BoxOrChannel,sI_Flag : String);
    procedure computeTotalReport(sI_BelongYM,sI_CompCode : String);
    procedure cdsTotalCommInitial;
    procedure cdsTotalMonthDataInitial(sI_Year : String);
    procedure addToCdsTotalComm1(I_Source,I_Data : TDataSet;sI_BelongYM,sI_FieldName,sI_BoxOrChannel : String;bI_IsIntAccData : Boolean);
    procedure addToCdsTotalComm2(I_Source,I_Data : TDataSet;sI_BelongYM,sI_FieldName,sI_Mode : String);
    function addZero(sI_SourceStr : String;nI_MaxLen : Integer) : String;
    procedure CalculateTotalMonthData(sI_CompCode,sI_Year : String);
    procedure addToCdsTotalMonthData(I_Source,I_Data : TDataSet;sI_CdsComputeYM,sI_Year : String);
    procedure computeCdsTotalMonthDataTotalAmount(I_Source,I_Data :  TDataSet;sI_Year : String);
    procedure showTotalMonthDataExcel(sI_Year : String);
    procedure addColumnZero(I_CDS : TClientDataSet);
  end;

var
  dtmMain1: TdtmMain1;

implementation

uses UCommonU, UObjectu, Ustru, frmSO8B20U, UdateTimeu, frmSO8B20_1U,
  frmSO8B50U, XLSFile, frmSO8B10U, frmMainMenuU, frmSO8B30U;

{$R *.dfm}

{ TdtmMain1 }



procedure TdtmMain1.saveToFile(vI_Content: OleVariant; sI_FileName: String);
var
    L_StrList : TStringList;
begin
    L_StrList := TStringList.Create;;
    L_StrList.Add(vI_Content);
    L_StrList.SaveToFile(sI_FileName);
    L_StrList.Free;
end;



procedure TdtmMain1.DataModuleCreate(Sender: TObject);
var
    sL_Result : String;
begin
{
    //down, �q ini ��Ū���s�u�� DB �� information
    sL_Result := getDbConnInfo(G_DbAliasStrList, G_DbUserNameStrList, G_DbPasswordStrList);
    if sL_Result<>'' then
    begin
      showmessage(sL_Result);
      Application.Terminate;
    end;
    //up, �q ini ��Ū���s�u�� DB �� information
}
    G_OrderStrList := TStringList.Create;
    G_PromCodeStrList:= TStringList.Create;
    G_PromNameStrList:= TStringList.Create;
end;

procedure TdtmMain1.createDbInfoStrList;
begin

    G_DbAreaNameStrList := TStringList.Create;
    G_DbAliasStrList := TStringList.Create;
    G_DbUserNameStrList := TStringList.Create;
    G_DbPasswordStrList := TStringList.Create;


end;


procedure TdtmMain1.dataSetActive;
begin

    {
    if not qrySO121.Active then
        qrySO121.Active := true;

    if not qrySO122.Active then
        qrySO122.Active := true;
    }
    if not qrySO120.Active then
        qrySO120.Active := true;

    if not qryCodeCD019.Active then
        qryCodeCD019.Active := true;

    if not qryCodeCD039.Active then
        qryCodeCD039.Active := true;

    if not qryCodeCD032.Active then
        qryCodeCD032.Active := true;

    if not cdsSO121.Active then
       cdsSO121.Active:=True;

    if not cdsSO122.Active then
       cdsSO122.Active:=True;

    if not cdsSO123.Active then
       cdsSO123.Active:=True;

    if not cdsSO125.Active then
       cdsSO125.Active:=True;

    if not cdsSO134.Active then
       cdsSO134.Active:=True;

    if not cdsSO135.Active then
       cdsSO135.Active:=True;
end;

function TdtmMain1.checkIsSO120OnlyData(sI_CompCode,sI_CodeNo,sI_PromoteCode,sI_DiscountCode: String): Boolean;
var
    sL_SQL : String;
    nL_Count : Integer;
begin
    with qryCom do
    begin
      SQL.Clear;
      if sI_CodeNo <> '' then
      begin
        //�Y�~�Ȭ��ʥN�X���O�� "�L" �ΥN�X�� "0"
        if sI_PromoteCode <> '0' then
        begin
          sL_SQL := 'SELECT COUNT(*) COUNT FROM SO120 WHERE COMPCODE=' + sI_CompCode +
                    ' AND CODENO=' + sI_CodeNo + ' AND PROMOTECODE=' + sI_PromoteCode +
                    ' AND DISCOUNTCODE=' + sI_DiscountCode;
        end
        else
        begin
          sL_SQL := 'SELECT COUNT(*) COUNT FROM SO120 WHERE COMPCODE=' + sI_CompCode +
                    ' AND CODENO=' + sI_CodeNo;
        end;
      end
      else
      begin
        sL_SQL := 'SELECT COUNT(*) COUNT FROM SO120 WHERE COMPCODE=' + sI_CompCode +
                  ' AND CODENO IS NULL';
      end;


      SQL.Add(sL_Sql);
      Open;

      nL_Count := FieldByName('COUNT').AsInteger ;

      if nL_Count >= 1 then
        Result := false
      else
        Result := true;
      Close;
    end;

end;

function TdtmMain1.getCurCobRecID(I_Cob: TComboBox): String;
begin
  Result := (I_Cob.Items.Objects[I_Cob.ItemIndex] as TNormalObj).s_Code;
end;

function TdtmMain1.getCurCobRecName(I_Cob: TComboBox): String;
begin
  Result := (I_Cob.Items.Objects[I_Cob.ItemIndex] as TNormalObj).s_Desc;
end;

procedure TdtmMain1.getSO120Data(sI_CompCode: String);
var
    sL_SQL : String;
begin
    sL_SQL := 'SELECT * FROM SO120 WHERE COMPCODE=' + sI_CompCode;
    qrySO120.SQL.Clear;
    qrySO120.SQL.Add(sL_SQL);
    qrySO120.Open;
end;

function TdtmMain1.chineseYMChangeToEnglishYM(sI_YM: String): String;
var
    L_StringList : TStringList;
    sL_Year,sL_YM : String;
 begin

    L_StringList := TUstr.ParseStrings(sI_YM,'/');

    sL_Year := IntToStr(StrToInt(L_StringList.Strings[0]) + 1911);

    sL_YM := sL_Year + '/' + L_StringList.Strings[1];// + '/' + L_StringList.Strings[2];

    Result := sL_YM;
end;

function TdtmMain1.ReduceThreeMonths(sI_YM: String): String;
var sL_YM:String;
begin

    if (copy(sI_YM,6,2) ='01') then  //�@��
    begin
      sL_YM:=IntToStr(StrToInt(copy(sI_YM,1,4))-1)+'/'+'10';
    end
    else if (copy(sI_YM,6,2) ='02') then //�G��
    begin
      sL_YM:=IntToStr(StrToInt(copy(sI_YM,1,4))-1)+'/'+'11';
    end
    else if (copy(sI_YM,6,2) ='03') then //�T��
    begin
      sL_YM:=IntToStr(StrToInt(copy(sI_YM,1,4))-1)+'/'+'12';
    end
    else if copy(sI_YM,6,1)='0' then //4,5,6,7,8,9
    begin
      sL_YM:=copy(sI_YM,1,4)+'/'+'0'+IntToStr(StrToInt(copy(sI_YM,7,1))-3);
    end
    else if copy(sI_YM,6,1)='1' then //10,11,12
    begin
      sL_YM:=copy(sI_YM,1,4)+'/'+'0'+IntToStr(StrToInt(copy(sI_YM,6,2))-3);
    end;
    Result :=sL_YM;
end;

procedure TdtmMain1.dealwithSTBBuy(sI_YM: String);
var
    sL_SQL,sL_StartInstDate,sL_StopInstDate,sL_MonthDays : String;
begin
    //select �Ӧ~�� (sI_YM) ���Ҧ��˾����
    sL_StartInstDate := sI_YM + '/01';
    sL_MonthDays := IntToStr(TUdateTime.DaysOfMonth(sL_StartInstDate));
    sL_StopInstDate := sI_YM + '/' + sL_MonthDays + ' 23:59:59';
    sL_StartInstDate := sL_StartInstDate + ' 00:00:00';

    sL_SQL := 'select a.SNo, a.IntroId, a.IntroName, a.WorkerEn1,a.WorkerName1,a.WorkerEn2,a.WorkerName2, ';
    sL_SQL := sL_SQL + 'a.WorkerEn3,a.WorkerName3,a.AcceptEn, a.AcceptName, a.MediaCode , a.MediaName ,a.ORDERNO,b.FaciSNo ,b.InstDate ,b.PRDate ,a.CustID,b.SeqNo,a.PromCode,a.PromName  ';
    sL_SQL := sL_SQL + ' from So007 a, So004 b ';
    sL_SQL := sL_SQL + ' where (b.InstDate BETWEEN ' + TUstr.getOracleSQLDateTimeStr(StrToDateTime(sL_StartInstDate));
    sL_SQL := sL_SQL + ' AND ' + TUstr.getOracleSQLDateTimeStr(StrToDateTime(sL_StopInstDate))+')';
    sL_SQL := sL_SQL + ' AND b.BuyCode=1'+' AND a.SNo=b.SNo AND b.SmartCardNo Is Not Null';
    sL_SQL := sL_SQL + ' AND b.FaciCode Is Not Null ';

//    saveToFile(sL_SQL,'C:\temp\buysL_SQL.txt');

    with qryBox do
    begin
      Close;
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;

      First;
      while not Eof do
      begin
        //SO121�������έp����  Jackal 920514
        //insertSO121(qryBox, BUY_BOX);  //�p��]����X box �ӻݭn��������
        insertSO122(qryBox, BUY_BOX);
        Next;
      end;
    end;


end;

procedure TdtmMain1.dealwithSTBRent(sI_YM: String);
var
    sL_SQL,sL_StartInstDate,sL_StopInstDate,sL_MonthDays  : String;
    dL_StartInstDate,dL_StopInstDate : TDateTime;
begin
    sL_StartInstDate := sI_YM + '/01';
    sL_MonthDays := IntToStr(TUdateTime.DaysOfMonth(sL_StartInstDate));
    sL_StopInstDate := sI_YM + '/' + sL_MonthDays + ' 23:59:59';
    sL_StartInstDate := sL_StartInstDate + ' 00:00:00';
    //showmessage(IntToStr(nG_RentBoxUseDays) + ' =�ഫ�e= ' + sL_StartInstDate + ' == ' + sL_StopInstDate);

    dL_StartInstDate := StrToDateTime(sL_StartInstDate)-nG_RentBoxUseDays;
    dL_StopInstDate := StrToDateTime(sL_StopInstDate)-nG_RentBoxUseDays;

    //showmessage('�ഫ��= ' + DateTimeToStr(dL_StartInstDate) + ' == ' + DateTimeToStr(dL_StopInstDate));
    sL_SQL := 'select a.SNo, a.IntroId, a.IntroName, a.WorkerEn1,a.WorkerName1,a.WorkerEn2,a.WorkerName2, ';
    sL_SQL := sL_SQL + 'a.WorkerEn3,a.WorkerName3,a.AcceptEn, a.AcceptName, a.MediaCode, a.MediaName,a.ORDERNO,b.FaciSNo ,b.InstDate ,b.PRDate  ,a.CustID,b.SeqNo,a.PromCode,a.PromName  ';

    sL_SQL := sL_SQL + 'from So007 a, So004 b ';
    sL_SQL := sL_SQL + ' where (b.InstDate BETWEEN ' + TUstr.getOracleSQLDateTimeStr(dL_StartInstDate);
    sL_SQL := sL_SQL + ' AND ' + TUstr.getOracleSQLDateTimeStr(dL_StopInstDate)+')';
    sL_SQL := sL_SQL + ' AND BuyCode=3'+' AND (b.PRDATE IS NULL OR ((b.PRDATE IS NOT NULL) AND (b.PRDate-b.InstDate>=' +IntToStr(nG_RentBoxUseDays) + ')))';

//SHOWMESSAGE('���wSNO ');
//    sL_SQL := sL_SQL + ' AND A.SNO=''200303IC0958595'' AND A.CUSTID=227260';

    sL_SQL := sL_SQL + ' AND a.SNo=b.SNo AND b.SmartCardNo Is Not Null AND b.FaciCode Is Not Null ';
    //saveToFile(sL_SQL,'C:\temp\RetsL_SQL.txt');

    with qryBox do
    begin
      Close;
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;
      while not Eof do
      begin
        //SO121�������έp����  Jackal 920514
        //insertSO121(qryBox, RENT_BOX);  //�p��]�����X box �ӻݭn��������
        insertSO122(qryBox, RENT_BOX);
        Next;
      end;
    end;
end;

procedure TdtmMain1.insertSO121(I_DateSet:TDataset; nI_OperationFlag:Integer);
var
    nL_IntrComm, nL_AccComm, nL_WorkenComm : Integer;

    sL_TargetID ,sL_TargetName ,sL_WorkerId ,sL_WorkerName: String;
    dL_TargetComm ,dL_WorNameComm: double;
    sL_IntroId, sL_AcceptId,sL_MediaCode,sL_RefNo,sL_BelongDays,sL_OrderNo: String;
    sL_MediaName,sL_CompType,sL_ClanCompCode,sL_ClanCompName,sL_PromCode,sL_PromName: String;
    dL_StartDate,dL_StopDate,dL_CurDate : TDateTime;
    nL_DiffDays : Integer;
    sL_NextBelongYM : String;
    bL_Result,bL_IsTheSameMonth : Boolean;

    procedure setTargetPerson;
    begin
      sL_IntroId :=I_DateSet.FieldByName('IntroId').AsString;


      if I_DateSet.FieldByName('WorkerEn1').AsString<>'' then
      begin
        sL_WorkerId:=I_DateSet.FieldByName('WorkerEn1').AsString;
        sL_WorkerName:=I_DateSet.FieldByName('WorkerName1').AsString;
      end
      else if (I_DateSet.FieldByName('WorkerEn1').AsString='')
         and (I_DateSet.FieldByName('WorkerEn2').AsString<>'') then
      begin
        sL_WorkerId:=I_DateSet.FieldByName('WorkerEn2').AsString;
        sL_WorkerName:=I_DateSet.FieldByName('WorkerName2').AsString;
      end
      else if (I_DateSet.FieldByName('WorkerEn1').AsString='')
         and (I_DateSet.FieldByName('WorkerEn2').AsString='')
         and (I_DateSet.FieldByName('WorkerEn3').AsString<>'') then
      begin
        sL_WorkerId:=I_DateSet.FieldByName('WorkerEn3').AsString;
        sL_WorkerName:=I_DateSet.FieldByName('WorkerName3').AsString
      end;


      sL_AcceptId:=I_DateSet.FieldByName('AcceptEn').AsString;


      sL_OrderNo:=I_DateSet.FieldByName('ORDERNO').AsString;
    end;
begin
    sL_TargetID:='';
    sL_TargetName:='';
    sL_WorkerName:='';
    sL_WorkerId:='';
    dL_TargetComm:=0;
    dL_WorNameComm:=0;
    setTargetPerson;
{
    �Y���ФH����,��ܸӵ��ݥI���Ӥ��ФH N ��;
    �Y���ФH�L��,���z�H������, ��ܸӵ��ݥI���Ө��z�H�� M ��. �I���u�{�H�� P ��
}
    bL_Result := CheckProm(sL_OrderNo,sL_PROMCODE,sL_PROMNAME);

    //�Y�O�n�ư����~�Ȭ���,��������
    if not bL_Result then
    begin
      case nI_OperationFlag of
        BUY_BOX :
        begin
          nL_IntrComm := nG_BuyBoxIntrComm;
          nL_AccComm := nG_BuyBoxAcceptComm;
          nL_WorkenComm := nG_SetBoxWorkerComm;
        end;
        RENT_BOX :
        begin
          nL_IntrComm := nG_RentBoxIntrComm;
          nL_AccComm := nG_RentBoxAcceptComm;
          nL_WorkenComm := nG_SetBoxWorkerComm;
        end;
      end;

      with cdsSo121 do
      begin
        if sL_IntroId<>'' then
        begin
          sL_TargetID :=sL_IntroId;      //���ФH�N�X
          sL_TargetName:=I_DateSet.FieldByName('IntroName').AsString;
          dL_TargetComm :=nL_IntrComm; //���ФH �o 500 ��
        end
        else if (sL_IntroId='') and (sL_AcceptId<>'') then
        begin
          sL_TargetID :=sL_AcceptId;    // ���z�H�N�X
          sL_TargetName :=I_DateSet.FieldByName('AcceptName').AsString;
          dL_TargetComm :=nL_AccComm;  //���z�H �o 50 ��;
        end;

        //�ˬd�O�_���ۤv���u,�O���ܤ~�o����
        sL_MediaCode := I_DateSet.FieldByName('MediaCode').AsString;

        sL_RefNo := '';
        sL_MediaName := '';

        if sL_MediaCode <> '' then
        begin
          if qryCodeCD009.Locate('CODENO', StrToInt(sL_MediaCode),[]) then
          begin
            sL_RefNo := qryCodeCD009.FieldByName('RefNo').AsString;
            sL_MediaName := qryCodeCD009.FieldByName('DESCRIPTION').AsString;
          end;
        end;


        if sL_RefNo = '2' then       //SO�ۤv���q
        begin
          sL_CompType := '1';
          sL_ClanCompCode := frmMainMenu.sG_CompCode;
          sL_ClanCompName := frmMainMenu.sG_CompName;
        end
        else if sL_RefNo = '3' then  //�P���I
        begin
          sL_CompType := '2';
          sL_ClanCompCode := sL_MediaCode;
          sL_ClanCompName := sL_MediaName;
        end
        else if sL_RefNo = '5' then  //���Y���~���q
        begin
          sL_CompType := '3';
          sL_ClanCompCode := sL_MediaCode;
          sL_ClanCompName := sL_MediaName;
        end;
  //jackal1223*************************
        if I_DateSet.FieldByName('InstDate').AsString <> '' then
          dL_StartDate := StrToDate(DateToStr(I_DateSet.FieldByName('InstDate').AsDateTime));

        sL_BelongDays := Copy(sG_BelongYM,1,3) + '/' + Copy(sG_BelongYM,4,5);
        sL_BelongDays := chineseYMChangeToEnglishYM(sL_BelongDays) + '/01';


        //�Y�L������,�h�H�k�ݦ~��Ĥ@�Ѩӭp��
        if I_DateSet.FieldByName('PRDate').AsString <> '' then
        begin
          bL_IsTheSameMonth := checkIsTheSameMonth(I_DateSet.FieldByName('InstDate').AsString,I_DateSet.FieldByName('PRDate').AsString);
          if bL_IsTheSameMonth then  //�Y�������P�k�ݤ���P�@�Ӥ�
            dL_StopDate := StrToDate(DateToStr(I_DateSet.FieldByName('PRDate').AsDateTime))
          else //���P��
            dL_StopDate := StrToDate(sL_BelongDays);
        end
        else
          dL_StopDate := StrToDate(sL_BelongDays);


        nL_DiffDays := Round(dL_StopDate - dL_StartDate);


        //�ˬdBox�O�_���,�Y�ܩ���ϥδ���,
        //��|���W�L�W�w����,�]��������
        if I_DateSet.FieldByName('PRDate').AsString <> '' then
        begin
          if ((nI_OperationFlag = 1) and (nL_DiffDays <= nG_BuyBoxUseDays))
             or ((nI_OperationFlag = 2) and (nL_DiffDays <= nG_RentBoxUseDays))then
          begin
            dL_TargetComm := 0;
          end;
        end;


        //���ФH�O�ۤv���u�~�p��
        if sL_RefNo <> '1' then
        begin
          //BOX�R�_�ί��άҭn�L�@�q�ɶ�,�~�o�����,�_�h�U�Ӥ�o��
          if ((nI_OperationFlag = 1) and (nL_DiffDays > nG_BuyBoxUseDays))
             or ((nI_OperationFlag = 2) and (nL_DiffDays > nG_RentBoxUseDays))then
          begin
            if cdsSo121.Locate('COMPCODE;BELONGYM;EMPID;COMPUTEYM',VarArrayOf([frmMainMenu.sG_CompCode,sG_BelongYM,sL_TargetID,sG_ComputeYM]),[]) then
            begin
              Edit;
              FieldByName('Commission').AsFloat:=FieldByName('Commission').AsFloat + dL_TargetComm;
              post;
            end
            else
            begin
              Append;
              FieldByName('CompCode').AsString:=frmMainMenu.sG_CompCode;
              FieldByName('ComputeYM').AsString:=sG_ComputeYM;
              FieldByName('BelongYM').AsString:=sG_BelongYM;
              FieldByName('EmpID').AsString:=sL_TargetID;
              FieldByName('EmpName').AsString:=sL_TargetName;

              FieldByName('Commission').AsFloat := dL_TargetComm;

              if sL_MediaCode <> '' then
                FieldByName('MediaCode').AsInteger := StrToInt(sL_MediaCode);
              FieldByName('MediaName').AsString := sL_MediaName;
              FieldByName('CompType').AsString := sL_CompType;

              if sL_ClanCompCode <> '' then
                FieldByName('ClanCompCode').AsInteger := StrToInt(sL_ClanCompCode);
              FieldByName('ClanCompName').AsString := sL_ClanCompName;

              FieldByName('OPERATOR').AsString:=frmMainMenu.sG_User;
              FieldByName('UpdTime').AsString := sG_CurDate;
              Post;
            end;
          end
          else
          begin
            //�U���k�ݦ~��
            sL_NextBelongYM := afterMonth(sG_BelongYM,1);

            if Length(sL_NextBelongYM) = 4 then
            begin
              sL_NextBelongYM := '0' + sL_NextBelongYM;
            end;

            if cdsSo121.Locate('COMPCODE;BELONGYM;EMPID;COMPUTEYM',VarArrayOf([frmMainMenu.sG_CompCode,sL_NextBelongYM,sL_TargetID,sG_ComputeYM]),[]) then
            begin
              Edit;
              FieldByName('Commission').AsFloat:=FieldByName('Commission').AsFloat + dL_TargetComm;
              post;
            end
            else
            begin
              Append;
              FieldByName('CompCode').AsString:=frmMainMenu.sG_CompCode;
              FieldByName('ComputeYM').AsString:=sG_ComputeYM;

              //�p���U�@�Ӥ�
              FieldByName('BelongYM').AsString:=sL_NextBelongYM;
              FieldByName('EmpID').AsString:=sL_TargetID;
              FieldByName('EmpName').AsString:=sL_TargetName;

              FieldByName('Commission').AsFloat := dL_TargetComm;

              if sL_MediaCode <> '' then
                FieldByName('MediaCode').AsInteger := StrToInt(sL_MediaCode);
              FieldByName('MediaName').AsString := sL_MediaName;
              FieldByName('CompType').AsString := sL_CompType;

              if sL_ClanCompCode <> '' then
                FieldByName('ClanCompCode').AsInteger := StrToInt(sL_ClanCompCode);
              FieldByName('ClanCompName').AsString := sL_ClanCompName;

              FieldByName('OPERATOR').AsString:=frmMainMenu.sG_User;
              FieldByName('UpdTime').AsString := sG_CurDate;
              Post;
            end;
          end;
        end;

        if sL_WorkerId<>'' then  //�u�{�H��
        begin
          dL_WorNameComm:=nL_WorkenComm; //�u�{�H�� �o 100 ��

          if cdsSo121.Locate('COMPCODE;BELONGYM;EMPID;COMPUTEYM',VarArrayOf([frmMainMenu.sG_CompCode,sG_BelongYM,sL_WorkerID,sG_ComputeYM]),[]) then
          begin
            Edit;
            FieldByName('Commission').AsFloat:=FieldByName('Commission').AsFloat+dL_WorNameComm;
            post;
          end
          else if (not cdsSo121.Locate('COMPCODE;BELONGYM;EMPID;COMPUTEYM',VarArrayOf([frmMainMenu.sG_CompCode,sG_BelongYM,sL_WorkerID,sG_ComputeYM]),[])) and (sL_WorkerID<>'') then
          begin
            Append;
            FieldByName('CompCode').AsString:=frmMainMenu.sG_CompCode;
            FieldByName('ComputeYM').AsString:=sG_ComputeYM;
            FieldByName('BelongYM').AsString:=sG_BelongYM;
            FieldByName('EmpID').AsString:=sL_WorkerID;
            FieldByName('EmpName').AsString:=sL_WorkerName;
            FieldByName('Commission').AsFloat:=dL_WorNameComm;

            //�u�{�H�����ݩ�SO���q�ҥH���T�w�`��
            FieldByName('MediaCode').AsInteger := WORKER_NO;
            FieldByName('MediaName').AsString := WORKER;

            FieldByName('CompType').AsString := '1';
            FieldByName('ClanCompCode').AsInteger := StrToInt(frmMainMenu.sG_CompCode);
            FieldByName('ClanCompName').AsString := frmMainMenu.sG_CompName;


            FieldByName('OPERATOR').AsString:=frmMainMenu.sG_User;
            //FieldByName('UpdTime').AsString:=IntToStr(StrToInt(FormatDateTime('YYYY',now))-1911)+ FormatDateTime('/MM/DD hh:mm:ss',now);
            FieldByName('UpdTime').AsString := sG_CurDate;
            Post;
          end;
        end;
      end;
    end;

end;



procedure TdtmMain1.insertSO122(I_DateSet:TDataset; nI_OperationFlag:Integer);
var
    nL_IntrComm, nL_AccComm, nL_WorkenComm ,nL_BuyOrRent : Integer;
    sL_MediaCode,sL_MediaName,sL_RefNo ,sL_BelongDays,sL_NextBelongYM,sL_OrderNo: String;
    dL_StartDate,dL_StopDate,dL_CurDate : TDateTime;
    dL_RealDate,dL_ShouldDate,dL_PayDate : TDate;
    nL_DiffDays : Integer;
    sL_PromCode,sL_PromName,sL_ComputeYM : String;
    bL_Result,bL_HaveDateData,bL_IsTheSameMonth : Boolean;
    sL_CustID,sL_SeqNo : String;

    wL_Year, wL_Month, wL_Day : Word;
    sL_Year,sL_PayYM,sL_Sno : String;


begin
{
    �H�C�@�� box �ΨC�@�ӥI�O�W�D�����, �O���C�@�� box �ΨC�@�ӥI�O�W�D �U�����ǤH�h�ֿ�
}
    sL_Sno := I_DateSet.FieldByName('SNo').AsString;

    sL_OrderNo := I_DateSet.FieldByName('ORDERNO').AsString;

    sL_PromCode := I_DateSet.FieldByName('PromCode').AsString;
    sL_PromName := I_DateSet.FieldByName('PromName').AsString;
    bL_Result := CheckProm(sL_OrderNo,sL_PromCode,sL_PromName);

    //�Y�O�n�ư����~�Ȭ���,��������
    if not bL_Result then
    begin
      case nI_OperationFlag of
        BUY_BOX :
        begin
          nL_IntrComm := nG_BuyBoxIntrComm;
          nL_AccComm := nG_BuyBoxAcceptComm;
          nL_WorkenComm := nG_SetBoxWorkerComm;
          nL_BuyOrRent := BUY_BOX;
        end;
        RENT_BOX :
        begin
          nL_IntrComm := nG_RentBoxIntrComm;
          nL_AccComm := nG_RentBoxAcceptComm;
          nL_WorkenComm := nG_SetBoxWorkerComm;
          nL_BuyOrRent := RENT_BOX;
        end;
      end;

      with cdsSo122 do
      begin
        Append;

        FieldByName('CompCode').AsString:=frmMainMenu.sG_CompCode;
        FieldByName('ComputeYM').AsString:=sG_ComputeYM;
        FieldByName('BELONGYM').AsString:=sG_BelongYM;

        FieldByName('StbNo').AsString:=I_DateSet.FieldByName('FaciSNo').AsString;
        FieldByName('SNo').AsString := sL_Sno;

        //sL_Sno:=I_DateSet.FieldByName('SNO').AsString;

        sL_ComputeYM:=Copy(sG_ComputeYM,1,3) + '/' + Copy(sG_ComputeYM,4,5);
        sL_ComputeYM := chineseYMChangeToEnglishYM(sL_ComputeYM) ;
        //sL_ComputeYM :=Copy(sL_ComputeYM,1,4) + Copy(sL_ComputeYM,6,2)+ '/01';


        //�Y SO033 �S��Ƨ� SO034
        sL_SeqNo := I_DateSet.FieldByName('SeqNo').AsString;
        bL_HaveDateData := getDateData(sL_SeqNo,sL_ComputeYM,'so033',dL_RealDate,dL_ShouldDate);
        if not bL_HaveDateData then
          getDateData(sL_SeqNo,sL_ComputeYM,'so034',dL_RealDate,dL_ShouldDate);


        //�O�_�u�{�H���@�S���h���u�{�H���G....?
        if I_DateSet.FieldByName('WorkerEn1').AsString='' then
            FieldByName('WorkerEn1Comm').AsFloat:=0
        else
        begin
          FieldByName('WorkerEn1').AsString:=I_DateSet.FieldByName('WorkerEn1').AsString;
          FieldByName('WorkerEn1Name').AsString:=I_DateSet.FieldByName('WorkerName1').AsString;
          FieldByName('WorkerEn1Comm').AsFloat:=nL_WorkenComm;
        end;

        //�ˬd�O�_���ۤv���u,�O���ܤ~�o����
        sL_MediaCode := I_DateSet.FieldByName('MediaCode').AsString;
        sL_MediaName := I_DateSet.FieldByName('MediaName').AsString;

        if sL_MediaCode <> '' then
        begin
          if qryCodeCD009.Locate('CODENO', StrToInt(sL_MediaCode),[]) then
          begin
            sL_RefNo := '';
            sL_RefNo := qryCodeCD009.FieldByName('RefNo').AsString;
          end;
        end;

        if I_DateSet.FieldByName('InstDate').AsString <> '' then
          dL_StartDate := StrToDate(DateToStr(I_DateSet.FieldByName('InstDate').AsDateTime));

        sL_BelongDays := Copy(sG_BelongYM,1,3) + '/' + Copy(sG_BelongYM,4,5);
        sL_BelongDays := chineseYMChangeToEnglishYM(sL_BelongDays) + '/01';


        //�Y�L������,�h�H�k�ݦ~��Ĥ@�Ѩӭp��
        if I_DateSet.FieldByName('PRDate').AsString <> '' then
        begin
          bL_IsTheSameMonth := checkIsTheSameMonth(I_DateSet.FieldByName('InstDate').AsString,I_DateSet.FieldByName('PRDate').AsString);
          if bL_IsTheSameMonth then  //�Y�������P�k�ݤ���P�@�Ӥ�
            dL_StopDate := StrToDate(DateToStr(I_DateSet.FieldByName('PRDate').AsDateTime))
          else //���P��
            dL_StopDate := StrToDate(sL_BelongDays);
        end
        else
          dL_StopDate := StrToDate(sL_BelongDays);

        nL_DiffDays := Round(dL_StopDate - dL_StartDate);


        //�Y�����ФH
        if I_DateSet.FieldByName('IntroId').AsString<>'' then
        begin
          FieldByName('IntAccId').AsString:=I_DateSet.FieldByName('IntroId').AsString;
          FieldByName('IntAccName').AsString:=I_DateSet.FieldByName('IntroName').AsString;

          //���ФH�O�ۤv���u�~�p��
          if sL_RefNo <> '1' then
          begin
            if ((nL_BuyOrRent = BUY_BOX) and (nL_DiffDays > nG_BuyBoxUseDays))
               or ((nL_BuyOrRent = RENT_BOX) and (nL_DiffDays > nG_RentBoxUseDays)) then
            begin
              FieldByName('IntAccComm').AsFloat := nL_IntrComm
            end
            else
              FieldByName('IntAccComm').AsFloat := 0;
          end;

          FieldByName('StaffCode').AsString:='1'; //��ܬO���ФH
          FieldByName('CallSelf').AsInteger:=0; //�D�ۨӹq

        end;

        //�Y�L���ФH�������z�H��
        if (I_DateSet.FieldByName('IntroId').AsString='') and (I_DateSet.FieldByName('AcceptEn').AsString<>'') then
        begin
          FieldByName('IntAccId').AsString:=I_DateSet.FieldByName('AcceptEn').AsString;
          FieldByName('IntAccName').AsString:=I_DateSet.FieldByName('AcceptName').AsString;


          if ((nL_BuyOrRent = BUY_BOX) and (nL_DiffDays > nG_BuyBoxUseDays))
             or ((nL_BuyOrRent = RENT_BOX) and (nL_DiffDays > nG_RentBoxUseDays)) then
          begin
            FieldByName('IntAccComm').AsFloat := nL_AccComm
          end
          else
            FieldByName('IntAccComm').AsFloat := 0;

          FieldByName('StaffCode').AsString:='2';  //��ܬO���z�H
          FieldByName('CallSelf').AsInteger:=1; //�ۨӹq
        end;

        FieldByName('OPERATOR').AsString:=frmMainMenu.sG_User;
        //FieldByName('UpdTime').AsString:=IntToStr(StrToInt(FormatDateTime('YYYY',now))-1911)+ FormatDateTime('/MM/DD hh:mm:ss',now);
        FieldByName('UpdTime').AsString := sG_CurDate;

        FieldByName('BUYORRENT').AsInteger := nL_BuyOrRent;

        //BOX �˾����
        if I_DateSet.FieldByName('InstDate').AsString <> '' then
          FieldByName('RealStartDate').AsDateTime := I_DateSet.FieldByName('InstDate').AsDateTime;
        //BOX ������
        if I_DateSet.FieldByName('PRDate').AsString <> '' then
          FieldByName('RealStopDate').AsDateTime := I_DateSet.FieldByName('PRDate').AsDateTime;

        //CD009 �� RefNo=2(���u����),RefNo=3(�P���I)
        //�YRefno=2,�h����쬰SO CompCode,�_�h��MediaCode
        if sL_RefNo='2' then
        begin
          FieldByName('MediaCode').AsInteger := StrToInt(frmMainMenu.sG_CompCode);
          FieldByName('MediaName').AsString := frmMainMenu.sG_CompName;
        end
        else if (sL_RefNo='3') or (sL_RefNo='5') then
        begin
          FieldByName('MediaCode').AsInteger := StrToInt(sL_MediaCode);
          FieldByName('MediaName').AsString := sL_MediaName;
        end;


        if sL_PromCode <> '' then
          FieldByName('PromCode').AsInteger:=StrToInt(sL_PromCode);

        FieldByName('PromName').AsString:=sL_PromName;

        if dL_RealDate <> 0  then
          FieldByName('REALDATE').AsDateTime:=dL_RealDate;

        if dL_ShouldDate <> 0 then
          FieldByName('SHOULDDATE').AsDateTime:=dL_ShouldDate;

        sL_CustID := I_DateSet.FieldByName('CustID').AsString;
        if sL_CustID <> '' then
        FieldByName('CustID').AsInteger := StrToInt(sL_CustID);

        FieldByName('OrderNo').AsString := sL_OrderNo;

        Post;


        //*********************************************************************
        //*********************************************************************
        { �Y�R�_�ί��ΨϥΤ��,���W�L�W�w���� ,  ���F�W�[�@�����k�ݦ~�묰�s����ƥ~
          , �A�N�����W�[�ܸӵ������  }
        if ((nL_BuyOrRent = BUY_BOX) and (nL_DiffDays <= nG_BuyBoxUseDays))
           or ((nL_BuyOrRent = RENT_BOX) and (nL_DiffDays <= nG_RentBoxUseDays)) then
        begin
          //�p��n�k�ݨ���@�Ӥ� --- Down
          if nL_BuyOrRent = BUY_BOX then    //�R
            dL_PayDate := dL_StartDate + nG_BuyBoxUseDays
          else if nL_BuyOrRent = RENT_BOX then    //��
            dL_PayDate := dL_StartDate + nG_RentBoxUseDays;

          DecodeDate(dL_PayDate,wL_Year, wL_Month, wL_Day);
          sL_Year := IntToStr((wL_Year - 1911));
          if Length(sL_Year) < 3 then
            sL_Year := '0' + sL_Year;

          sL_PayYM := sL_Year + IntToStr(wL_Month);

          sL_NextBelongYM := afterMonth(sL_PayYM,1);

          if Length(sL_NextBelongYM) = 4 then
          begin
            sL_NextBelongYM := '0' + sL_NextBelongYM;
          end;

          if sG_BelongYM > sL_NextBelongYM then
            sL_NextBelongYM := sG_BelongYM;
          //�p��n�k�ݨ���@�Ӥ� --- Up


          Append;
          FieldByName('CompCode').AsString:=frmMainMenu.sG_CompCode;
          FieldByName('ComputeYM').AsString:=sG_ComputeYM;

          //�N�����[��U�Ӥ�
          FieldByName('BELONGYM').AsString:=sL_NextBelongYM;

          FieldByName('StbNo').AsString:=I_DateSet.FieldByName('FaciSNo').AsString;
          FieldByName('SNo').AsString:=I_DateSet.FieldByName('SNo').AsString;


          FieldByName('WorkerEn1').AsString:=I_DateSet.FieldByName('WorkerEn1').AsString;
          FieldByName('WorkerEn1Name').AsString:=I_DateSet.FieldByName('WorkerName1').AsString;
          //�u�@�H����e�@�Ӥ�,�w�ˮɧY������
          FieldByName('WorkerEn1Comm').AsFloat:=0;


          //�ˬd�O�_���ۤv���u,�O���ܤ~�o����
          sL_MediaCode := I_DateSet.FieldByName('MediaCode').AsString;
          sL_MediaName := I_DateSet.FieldByName('MediaName').AsString;

          if sL_MediaCode <> '' then
          begin
            if qryCodeCD009.Locate('CODENO', StrToInt(sL_MediaCode),[]) then
            begin
              sL_RefNo := '';
              sL_RefNo := qryCodeCD009.FieldByName('RefNo').AsString;
            end;
          end;

          if I_DateSet.FieldByName('InstDate').AsString <> '' then
            dL_StartDate := StrToDate(DateToStr(I_DateSet.FieldByName('InstDate').AsDateTime));

          sL_BelongDays := Copy(sG_BelongYM,1,3) + '/' + Copy(sG_BelongYM,4,5);
          sL_BelongDays := chineseYMChangeToEnglishYM(sL_BelongDays) + '/01';


          //�Y�L������,�h�H�k�ݦ~��Ĥ@�Ѩӭp��
          if I_DateSet.FieldByName('PRDate').AsString <> '' then
          begin
          {
            bL_IsTheSameMonth := checkIsTheSameMonth(I_DateSet.FieldByName('InstDate').AsString,I_DateSet.FieldByName('PRDate').AsString);
            if bL_IsTheSameMonth then  //�Y�������P�k�ݤ���P�@�Ӥ�
              dL_StopDate := StrToDate(DateToStr(I_DateSet.FieldByName('PRDate').AsDateTime))
            else //���P��
              dL_StopDate := StrToDate(sL_BelongDays);
            }
            //�ĤG���p��H�������p��
            dL_StopDate := StrToDate(DateToStr(I_DateSet.FieldByName('PRDate').AsDateTime));
          end
          else
            dL_StopDate := StrToDate(sL_BelongDays);

          nL_DiffDays := Round(dL_StopDate - dL_StartDate);


          //�Y�����ФH
          if I_DateSet.FieldByName('IntroId').AsString<>'' then
          begin
            FieldByName('IntAccId').AsString:=I_DateSet.FieldByName('IntroId').AsString;
            FieldByName('IntAccName').AsString:=I_DateSet.FieldByName('IntroName').AsString;



            //���ФH�O�ۤv���u�~�p��
            if sL_RefNo <> '1' then
            begin
              //Box�ϥδ����Y���W�L�W�w����,������U�Ӥ뵹,
              //�|���ˬdBox�O�_���,�Y�ܩ���ϥδ���,
              //��|���W�L�W�w����,�]��������
              if I_DateSet.FieldByName('PRDate').AsString <> '' then
              begin
                if ((nI_OperationFlag = 1) and (nL_DiffDays <= nG_BuyBoxUseDays))
                   or ((nI_OperationFlag = 2) and (nL_DiffDays <= nG_RentBoxUseDays))then
                begin
                  FieldByName('IntAccComm').AsFloat := 0
                end
                else
                  FieldByName('IntAccComm').AsFloat := nL_IntrComm;
              end
              else
                FieldByName('IntAccComm').AsFloat := nL_IntrComm;
            end;

            FieldByName('StaffCode').AsString:='1'; //��ܬO���ФH
            FieldByName('CallSelf').AsInteger:=0; //�D�ۨӹq
          end;

          //�Y�L���ФH�������z�H��
          if (I_DateSet.FieldByName('IntroId').AsString='') and (I_DateSet.FieldByName('AcceptEn').AsString<>'') then
          begin
            FieldByName('IntAccId').AsString:=I_DateSet.FieldByName('AcceptEn').AsString;
            FieldByName('IntAccName').AsString:=I_DateSet.FieldByName('AcceptName').AsString;


            FieldByName('StaffCode').AsString:='2';  //��ܬO���z�H
            FieldByName('CallSelf').AsInteger:=1; //�ۨӹq

            //Box�ϥδ����Y���W�L�W�w����,������U�Ӥ뵹,
            //�|���ˬdBox�O�_���,�Y�ܩ���ϥδ���,
            //��|���W�L�W�w����,�]��������
            if I_DateSet.FieldByName('PRDate').AsString <> '' then
            begin
              if ((nI_OperationFlag = 1) and (nL_DiffDays <= nG_BuyBoxUseDays))
                 or ((nI_OperationFlag = 2) and (nL_DiffDays <= nG_RentBoxUseDays))then
              begin
                FieldByName('IntAccComm').AsFloat := 0
              end
              else
                FieldByName('IntAccComm').AsFloat := nL_AccComm;
            end
            else
              FieldByName('IntAccComm').AsFloat := nL_AccComm;
          end;

          FieldByName('OPERATOR').AsString:=frmMainMenu.sG_User;
          FieldByName('UpdTime').AsString := sG_CurDate;

          FieldByName('BUYORRENT').AsInteger := nL_BuyOrRent;

          //BOX �˾����
          if I_DateSet.FieldByName('InstDate').AsString <> '' then
            FieldByName('RealStartDate').AsDateTime := I_DateSet.FieldByName('InstDate').AsDateTime;
          //BOX ������
          if I_DateSet.FieldByName('PRDate').AsString <> '' then
          begin
            FieldByName('RealStopDate').AsDateTime := I_DateSet.FieldByName('PRDate').AsDateTime;

            //showmessage(I_DateSet.FieldByName('PRDate').AsString);
          end;

          //CD009 �� RefNo=2(���u����),RefNo=3(�P���I)
          //�YRefno=2,�h����쬰SO CompCode,�_�h��MediaCode
          if sL_RefNo='2' then
          begin
            FieldByName('MediaCode').AsInteger := StrToInt(frmMainMenu.sG_CompCode);
            FieldByName('MediaName').AsString := frmMainMenu.sG_CompName;
          end
          else if (sL_RefNo='3') or (sL_RefNo='5') then
          begin
            FieldByName('MediaCode').AsInteger := StrToInt(sL_MediaCode);
            FieldByName('MediaName').AsString := sL_MediaName;
          end;



          if sL_PromCode <> '' then
            FieldByName('PromCode').AsInteger:=StrToInt(sL_PromCode);

          FieldByName('PromName').AsString:=sL_PromName;

          if dL_RealDate <> 0  then
            FieldByName('REALDATE').AsDateTime:=dL_RealDate;

          if dL_ShouldDate <> 0 then
            FieldByName('SHOULDDATE').AsDateTime:=dL_ShouldDate;


          sL_CustID := I_DateSet.FieldByName('CustID').AsString;
          if sL_CustID <> '' then
          FieldByName('CustID').AsInteger := StrToInt(sL_CustID);

          FieldByName('OrderNo').AsString := sL_OrderNo;

          Post;
        end;
      end;
    end;

end;

function TdtmMain1.CheckTableSo121(sI_ComputeYM :String): Boolean;
var sL_SQL : String;
begin
    with qryCommon do
    begin
      sL_SQL:='Select Count(*) CNT from So121 where COMPCODE='''+frmMainMenu.sG_CompCode+'''';
      sL_SQL:=sL_SQL+' AND COMPUTEYM='''+sI_ComputeYM+'''';
      Close;
      Sql.Clear;
      Sql.Add(sL_SQL);
      Open;
      if FieldByName('CNT').Value>0 then
        Result:=True
      else
        Result:=False;
      Close;
    end;


end;

function TdtmMain1.preProcessingData(sI_TableName,sI_YM,sI_PayMode: String): String;
begin
    if sI_PayMode = BATCH_PAY then      //�������I
      frmSO8B20.StatusBar1.SimpleText := '�B�z�I�O�W�D-[�������I]-[���O]���...'
    else if sI_PayMode = ONCE_PAY then  //�@�����I
      frmSO8B20.StatusBar1.SimpleText := '�B�z�I�O�W�D-[�@�����I]-[���O]���...';

    //�B�z�������I�W�D����
    handleInCome(sI_TableName,sI_YM,sI_PayMode);


    if sI_PayMode = BATCH_PAY then      //�������I
    begin
      frmSO8B20.StatusBar1.SimpleText := '�B�z�I�O�W�D-[�������I]-[�h�O]���...';
      handleOutCome(sI_TableName,sI_YM);
    end
    else if sI_PayMode = ONCE_PAY then  //�@�����I
    begin
      frmSO8B20.StatusBar1.SimpleText := '�B�z�I�O�W�D-[�@�����I]-[�h�O]���...';
      handleOnceOutCome(sI_TableName,sI_YM);
    end;
end;

function TdtmMain1.processIncome(sI_OrderNo,sI_RealDate: String; I_DataSet:TDataSet): Boolean;
var
    fL_RealAmt : double;
    sL_SQL,sL_RefNo,sL_CustId, sL_CompCode:String;
    sL_AcceptEn, sL_AcceptName,sL_MonthFirstDay : String;
    bL_CallIn,bL_PayByCreditCard:boolean;
    bL_SelfCallIn, bL_HasPayComm,bL_Result : boolean;
    ii,  nL_RealPeriod,nL_ComputePeriodCounts:Integer;                                                                //qrySo033So034
    sL_IntroId, sL_IntroName ,sL_RealStartDate,sL_RealStopDate : String;
    sL_ClctId, sL_ClctName,sL_MediaCode,sL_MediaName,sL_NextYM,sL_beforeMonth : String;
    nL_TotalPeriodDays,nL_FirstPeriodDays,nL_OtherPeriodDays,nL_PeriodCounts : Integer;
    fL_ClctEnOriPercent,fL_IntAccOriPercent : Double;
    dL_RealStartDate,dL_RealStopDate : TDate;
    sL_PromCode,sL_PromName,sL_CitemCode,sL_RealBelongYM : String;
    bL_HaveClctId,bL_HavePayComm : Boolean;
    sL_SelfCallInBelongYM : String;
    bL_HaveSelfCallInBelongYM : Boolean;
begin
    //�B�z�������I�W�D����
    sL_CustId := I_DataSet.fieldByName('CUSTID').AsString;
    sL_CompCode := I_DataSet.fieldByName('COMPCODE').AsString;
    sL_CitemCode := I_DataSet.FieldByName('CitemCode').AsString;
    sL_ClctId := I_DataSet.FieldByName('ClctEn').AsString;
    sL_ClctName := I_DataSet.FieldByName('ClctName').AsString;


    //�P�_�O�_���H�Υd=>�YSo033/So034.PTCode=3,��ܬ��H�Υd�I�O
    bL_PayByCreditCard := ifPayByCreditCard(I_DataSet.FieldByName('PTCode').AsString);

    //���ˬd���L���O�H��,�S���h���p��
    //���O��( �Y�O�H�Υd���O�BSO033�S�����O�H��,�h��SO132�� )
    if (bL_PayByCreditCard = true) and (sL_ClctId = '') then
    begin
      //�� SO132 ��H�Υd�I�O�����k�ݤH��
      bL_HaveClctId := getEmpData(sL_CompCode,sL_CustID,sL_CitemCode,sL_ClctId,sL_ClctName);
    end
    else if sL_ClctId <> '' then
      bL_HaveClctId := true
    else
      bL_HaveClctId := false;


    if bL_HaveClctId then    //�����O�H��
    begin
      sL_SQL:='Select MediaCode,ACCEPTEN, ACCEPTNAME, IntroId, IntroName,PromCode,PromName ';
      sL_SQL:=sL_SQL+' from So105 ';
      sL_SQL:=sL_SQL+' where OrderNo='''+sI_OrderNo+'''';
  //    saveToFile(sL_SQL,'C:\temp\CallInCheck_SQL.txt');

      with qryCom do
      begin
        Close;
        SQL.Clear;
        SQL.Add(sL_SQL);
        Open;

        if RecordCount=0 then
        begin
          Close;
          Exit;
        end
        else
        begin
          sL_MediaCode := FieldByName('MediaCode').AsString;
          sL_AcceptEn := FieldByName('ACCEPTEN').AsString;
          sL_AcceptName := FieldByName('ACCEPTNAME').AsString;
          sL_IntroId := FieldByName('IntroId').AsString;
          sL_IntroName := FieldByName('IntroName').AsString;

          //�~�Ȭ��ʸ��
          sL_PromCode := FieldByName('PromCode').AsString;
          sL_PromName := FieldByName('PromName').AsString;

          bL_SelfCallIn := true;

          sL_RefNo := '';
          sL_MediaName := '';

          if sL_MediaCode <> '' then
          begin
            if qryCodeCD009.Locate('CODENO', StrToInt(sL_MediaCode),[]) then
            begin
              sL_RefNo := qryCodeCD009.FieldByName('RefNo').AsString;
              sL_MediaName := qryCodeCD009.FieldByName('DESCRIPTION').AsString;

              //���p���O�ۨӹq
              if qryCodeCD009.FieldByName('RefNo').AsInteger in [1,2,3,5] then
                bL_SelfCallIn := false;
            end;
          end;
        end;
        Close;
      end;

      sL_RealStartDate := qrySo033So034.FieldByName('RealStartDate').AsString;
      sL_RealStopDate := qrySo033So034.FieldByName('RealStopDate').AsString;

      //�p�� �`���ƤѼ� �� �����p�O�Ѽ� �� ���p�����
      getComputeDays(sG_BelongYM,sL_RealStartDate,sL_RealStopDate,nL_TotalPeriodDays,nL_FirstPeriodDays,nL_PeriodCounts,sL_RealBelongYM);

      fL_ClctEnOriPercent := 0;
      fL_IntAccOriPercent := 0;


//******************************************
//******************************************
      fG_SpreadComm:=0;
      fG_ChargeComm:=0;


      nL_RealPeriod := qrySo033So034.FieldByName('RealPeriod').AsInteger;
      fL_RealAmt := qrySo033So034.FieldByName('RealAmt').AsFloat;
      dL_RealStartDate := qrySo033So034.FieldByName('RealStartDate').AsDateTime;
      dL_RealStopDate := qrySo033So034.FieldByName('RealStopDate').AsDateTime;


      if (qrySo033So034.FieldByName('FirstFlag').AsString='1') then//�O����
      begin
        //���s�H�����������W�L�������
        if nL_PeriodCounts > nG_LimitPeriod then
        begin
          //�p�L�Ĥ@������30��,�h�����n�h���@��
          if nL_FirstPeriodDays < 30 then
            nL_ComputePeriodCounts := nG_LimitPeriod + 1
          else
            nL_ComputePeriodCounts := nG_LimitPeriod;
        end
        else
          nL_ComputePeriodCounts := nL_PeriodCounts;

        //�p����ˤH������
        sL_NextYM := sL_RealBelongYM;

        bL_HaveSelfCallInBelongYM := false;
        for ii:=1 to nL_ComputePeriodCounts do
        begin
          fL_ClctEnOriPercent := 0;
          fL_IntAccOriPercent := 0;

          if ii <> 1 then  //�p��U�Ӧ~��
          begin
            sL_NextYM := afterMonth(sL_NextYM,ii-2);
            nL_OtherPeriodDays := TUdateTime.DaysOfMonth(sL_NextYM + '01');

            //�p��̫�@��
            if ii = nL_ComputePeriodCounts then
            begin
              //�p�G��ڴ��Ƥp�󭭨����,�h�̫�@���Ѽ�,�����HRealStopDate�p��
              if nL_PeriodCounts <= nG_LimitPeriod then
              begin
                //���o�Ӥ몺�Ĥ@��
                sL_MonthFirstDay := TUdateTime.getMonthFirstDay(sL_RealStopDate);
                nL_OtherPeriodDays := getLastPeriodDays(StrToDate(sL_MonthFirstDay),StrToDate(sL_RealStopDate));
              end;
            end;

            //���Y�����p�O�ѼƤp��30��,�h�����é�U���@�ֵ�
            if (ii = 2) AND (nL_FirstPeriodDays <= nG_ChannelViewDays)then
              nL_OtherPeriodDays := nL_OtherPeriodDays + nL_FirstPeriodDays;
          end;
//showmessage(sG_BelongYM + '--' + IntToStr(ii) + '==' + sL_NextYM + '--' + IntToStr(nL_OtherPeriodDays));

          if bL_PayByCreditCard then//�H�H�Υd�I�O
          begin
            if sG_PayUnit='2' then    //��
            begin
              //�ˬd���s�H�O�_���ۤv���u,�O���ܤ~�o����
              if sL_RefNo <> '1' then
              begin
                //�|����Τ��p��,�]�O���_����h�O,�ҥH�����@������(�Ĥ@���N�@������)
                {
                if (ii = 1) AND (nL_FirstPeriodDays > nG_ChannelViewDays) then
                  fG_SpreadComm := fG_FirCardSpd    //���s�H
                else
                  fG_SpreadComm := 0;
                }
                if ii = 1 then
                begin
                  if nL_FirstPeriodDays > nG_ChannelViewDays then
                  begin
                    //�ۨӹq�k�ݦ~��
                    sL_SelfCallInBelongYM := sL_NextYM;

                    fG_SpreadComm := fG_FirCardSpd;    //���s�H
                    bL_HavePayComm := true;
                  end
                  else
                  begin
                    fG_SpreadComm := 0;    //���s�H
                    bL_HavePayComm := false;
                  end;
                end
                else
                begin
                  if bL_HavePayComm then
                    fG_SpreadComm := 0
                  else
                  begin
                    if nL_OtherPeriodDays > nG_ChannelViewDays then
                    begin
                      //�ۨӹq�k�ݦ~��
                      sL_SelfCallInBelongYM := sL_NextYM;

                      fG_SpreadComm := fG_FirCardSpd;    //���s�H
                      bL_HavePayComm := true;
                    end
                    else
                    begin
                      fG_SpreadComm := 0;    //���s�H
                      bL_HavePayComm := false;
                    end;
                  end;
                end;

              end
              else
                fG_SpreadComm := 0;

              //���O���u�p�⭺��
              if ii = 1 then
                fG_ChargeComm:=fG_FirCardCharge
              else
                fG_ChargeComm := 0;
            end
            else if sG_PayUnit='1' then  //�ʤ���
            begin
              //�ˬd���s�H�O�_���ۤv���u,�O���ܤ~�o����
              if sL_RefNo <> '1' then
              begin
                if (ii = 1) AND (nL_FirstPeriodDays > nG_ChannelViewDays) then   //�Ĥ@������=(RealAmt * % ) * �����p�O�Ѽ� / �`���ƤѼ�
                  fG_SpreadComm := (fL_RealAmt*(fG_FirCardSpd/100))*nL_FirstPeriodDays/nL_TotalPeriodDays //0.12;      //���s�H
                else            //�D�Ĥ@������=(RealAmt * % ) * ���Ѽ� / �`���ƤѼ�
                  fG_SpreadComm := (fL_RealAmt*(fG_FirCardSpd/100))*nL_OtherPeriodDays/nL_TotalPeriodDays; //0.12;      //���s�H

                //�ۨӹq�k�ݦ~�묰�Ĥ@�����������Ӧ~��
                if (fG_SpreadComm>0) and (bL_HaveSelfCallInBelongYM=false) then
                begin
                  //�ۨӹq�k�ݦ~��
                  sL_SelfCallInBelongYM := sL_NextYM;
                  bL_HaveSelfCallInBelongYM := true;
                end;
              end
              else
                fG_SpreadComm := 0;


              //���O���u�p�⭺��
              if ii = 1 then
                fG_ChargeComm := (fL_RealAmt*(fG_FirCardCharge/100))  //0.01;      //���O��
                //���O�H���@������
                //fG_ChargeComm := (fL_RealAmt*(fG_FirCardCharge/100))*nL_FirstPeriodDays/nL_TotalPeriodDays  //0.01;      //���O��
              else
                fG_ChargeComm := 0;


              //�����p��ʤ���@���h�O�ɨ̾�
              fL_ClctEnOriPercent := fG_FirCardCharge;
              fL_IntAccOriPercent := fG_FirCardSpd;
            end;
          end
          else//���O�H�H�Υd�I�O
          begin
            if sG_PayUnit='2' then    //��
            begin
              //�ˬd���s�H�O�_���ۤv���u,�O���ܤ~�o����
              if sL_RefNo <> '1' then
              begin
                //�|����Τ��p��,�]�O���_����h�O,�ҥH�����@������(�Ĥ@���N�@������)
                {
                if (ii = 1) AND (nL_FirstPeriodDays > nG_ChannelViewDays) then
                  fG_SpreadComm := fG_FirNoCardSpd  //���s�H
                else
                  fG_SpreadComm := 0;
                }
                if ii = 1 then
                begin
                  if nL_FirstPeriodDays > nG_ChannelViewDays then
                  begin
                    //�ۨӹq�k�ݦ~��
                    sL_SelfCallInBelongYM := sL_NextYM;

                    fG_SpreadComm := fG_FirNoCardSpd;    //���s�H
                    bL_HavePayComm := true;
                  end
                  else
                  begin
                    fG_SpreadComm := 0;    //���s�H
                    bL_HavePayComm := false;
                  end;
                end
                else
                begin
                  if bL_HavePayComm then
                    fG_SpreadComm := 0
                  else
                  begin
                    if nL_OtherPeriodDays > nG_ChannelViewDays then
                    begin
                      //�ۨӹq�k�ݦ~��
                      sL_SelfCallInBelongYM := sL_NextYM;

                      fG_SpreadComm := fG_FirNoCardSpd;    //���s�H
                      bL_HavePayComm := true;
                    end
                    else
                    begin
                      fG_SpreadComm := 0;    //���s�H
                      bL_HavePayComm := false;
                    end;
                  end;
                end;

              end
              else
                fG_SpreadComm := 0;

              //���O���u�p�⭺��
              if ii = 1 then
                fG_ChargeComm := fG_FirNoCardCharge
              else
                fG_ChargeComm := 0;
            end
            else if sG_PayUnit='1' then //�ʤ���
            begin
              //�ˬd���s�H�O�_���ۤv���u,�O���ܤ~�o����
              if sL_RefNo <> '1' then
              begin
                if (ii = 1) AND (nL_FirstPeriodDays > nG_ChannelViewDays) then //�Ĥ@������=(RealAmt * % ) * �����p�O�Ѽ� / �`���ƤѼ�
                  fG_SpreadComm := (fL_RealAmt*(fG_FirNoCardSpd/100))*nL_FirstPeriodDays/nL_TotalPeriodDays//0.12;  //���s�H
                else
                  fG_SpreadComm := (fL_RealAmt*(fG_FirNoCardSpd/100))*nL_OtherPeriodDays/nL_TotalPeriodDays;//0.12;  //���s�H

                //�ۨӹq�k�ݦ~�묰�Ĥ@�����������Ӧ~��
                if (fG_SpreadComm>0) and (bL_HaveSelfCallInBelongYM=false) then
                begin
                  //�ۨӹq�k�ݦ~��
                  sL_SelfCallInBelongYM := sL_NextYM;
                  bL_HaveSelfCallInBelongYM := true;
                end;
              end
              else
                fG_SpreadComm := 0;

              //���O���u�p�⭺��
              if ii = 1 then
                fG_ChargeComm:=(fL_RealAmt*(fG_FirNoCardCharge/100))  //0.03; //���O��
                //���O�H���@������
                //fG_ChargeComm:=(fL_RealAmt*(fG_FirNoCardCharge/100))*nL_FirstPeriodDays/nL_TotalPeriodDays  //0.03; //���O��
              else
                fG_ChargeComm := 0;

              //�����p��ʤ���@���h�O�ɨ̾�
              fL_ClctEnOriPercent := fG_FirNoCardCharge;
              fL_IntAccOriPercent := fG_FirNoCardSpd;
            end;
          end;

          //���ެO�Ĥ@���p��έ��s�p�⪺��ƥ��s�Jcds
          if not bL_SelfCallIn then//���O�ۨӹq
            SaveTempDataTocdsSo122(sL_RealBelongYM,sL_PromCode,sL_PromName,sL_ClctId, sL_ClctName, sL_IntroId, sL_IntroName,sL_MediaCode,sL_MediaName,sL_RefNo,sI_OrderNo, I_DataSet,ii,dL_RealStartDate,dL_RealStopDate,fL_ClctEnOriPercent,fL_IntAccOriPercent);
          //SO121�������έp����  Jackal 920514
          //�ˬd���s�H�O�_���ۤv���u,�O���ܤ~�o����
          //if sL_RefNo <> '1' then
          //  SaveTempDataTocdsSo121(sL_ClctId, sL_ClctName, sL_IntroId, sL_IntroName,sL_MediaCode,sL_MediaName,sL_RefNo, I_DataSet,ii);
        end;
      end
      else//�O��
      begin
        //�򦬥u�p�⦬�O�H��
        fG_SpreadComm := 0;

        sL_beforeMonth := beforeMonth(sG_BelongYM,1);
        nL_OtherPeriodDays := TUdateTime.DaysOfMonth(sL_beforeMonth + '01');

        if bL_PayByCreditCard then    //�H�H�Υd�I�O
        begin
          if sG_PayUnit='2' then       //��
            fG_ChargeComm := fG_ContCardCharge //���O��
          else if sG_PayUnit='1' then //�ʤ���
            fG_ChargeComm := (fL_RealAmt*(fG_ContCardCharge/100));//0.01 //���O��


          //�����p��ʤ���@���h�O�ɨ̾�
          fL_ClctEnOriPercent := fG_ContCardCharge;
          fL_IntAccOriPercent := 0;
        end
        else
        begin                         //���H�H�Υd�I�O
          if sG_PayUnit='2' then //��
            fG_ChargeComm := fG_ContNoCardCharge    //���O��
          else if sG_PayUnit='1' then //�ʤ���
            fG_ChargeComm:=(fL_RealAmt*(fG_ContNoCardCharge/100));//0.03;    //���O��


          //�����p��ʤ���@���h�O�ɨ̾�
          fL_ClctEnOriPercent := fG_ContNoCardCharge;
          fL_IntAccOriPercent := 0;

        end;

        //���ެO�Ĥ@���p��έ��s�p�⪺��ƥ��s�Jcds
        if not bL_SelfCallIn then //���O�ۨӹq
          SaveTempDataTocdsSo122(sL_RealBelongYM,sL_PromCode,sL_PromName,sL_ClctId,sL_ClctName,sL_IntroId,sL_IntroName,sL_MediaCode,sL_MediaName,sL_RefNo,sI_OrderNo,I_DataSet,THIS_MONTH,dL_RealStartDate,dL_RealStopDate,fL_ClctEnOriPercent,fL_IntAccOriPercent);

        //SO121�������έp����  Jackal 920514
        //�ˬd���s�H�O�_���ۤv���u,�O���ܤ~�o����
        //if sL_RefNo <> '1' then
        //  SaveTempDataTocdsSo121(sL_ClctId, sL_ClctName, sL_IntroId, sL_IntroName,sL_MediaCode,sL_MediaName,sL_RefNo,I_DataSet,THIS_MONTH);

      end;


      if (bL_SelfCallIn) and (qrySo033So034.FieldByName('FirstFlag').AsString='1') then //�O�ۨӹq,�B�O����
      begin
        if fL_RealAmt <> 0 then    //�ꦬ���B����0�����έp��
        begin
          bL_HasPayComm := CustIDCheck(sL_CustId,sL_CompCode); //�ˬd�O�_�w�g���L����

          try
            if not bL_HasPayComm then
              bL_HasPayComm := insertSO123(sL_CompCode,sL_CustID, sL_AcceptEn,sL_AcceptName);


            //���Y�O�ۨӹq,���ƥu���@��,�B�n�Q�k�ݪ��~�를�w
            if (nL_ComputePeriodCounts=1) and (bL_HaveSelfCallInBelongYM=false)  then
              sL_SelfCallInBelongYM := afterMonth(sL_NextYM,1);

            sL_SelfCallInBelongYM := Format('%.5d',[StrToInt(sL_SelfCallInBelongYM)]);
            insertCallInDataToSo122(bL_HasPayComm,sL_SelfCallInBelongYM,sL_PromCode,sL_PromName,sL_CompCode,sG_ComputeYM,sL_AcceptEn,sL_AcceptName,sL_ClctId,sL_ClctName,sI_OrderNo, I_DataSet,fL_IntAccOriPercent,fL_ClctEnOriPercent,fL_RealAmt);
          except
            ShowMessage(sL_CustId + '�ۨӹq�����B�z����_1!');
            Result:=False;
            Exit;
          end;
        end;
      end;
    end
    else
    begin  //�S�����O�H��
      cdsSo033Log.Append;
      cdsSo033Log.FieldByName('CUSTID').AsString := qrySo033So034.FieldByName('CUSTID').AsString;
      cdsSo033Log.FieldByName('BILLNO').AsString := qrySo033So034.FieldByName('BILLNO').AsString;
      cdsSo033Log.FieldByName('ITEM').AsString := qrySo033So034.FieldByName('ITEM').AsString;
      cdsSo033Log.FieldByName('CITEMCODE').AsString := qrySo033So034.FieldByName('CITEMCODE').AsString;
      cdsSo033Log.FieldByName('CITEMNAME').AsString := qrySo033So034.FieldByName('CITEMNAME').AsString;
      cdsSo033Log.FieldByName('SHOULDDATE').AsString := qrySo033So034.FieldByName('SHOULDDATE').AsString;
      cdsSo033Log.FieldByName('REALDATE').AsString := qrySo033So034.FieldByName('REALDATE').AsString;
      cdsSo033Log.FieldByName('REALAMT').AsString := qrySo033So034.FieldByName('REALAMT').AsString;
      cdsSo033Log.FieldByName('REALPERIOD').AsString := qrySo033So034.FieldByName('RealPeriod').AsString;
      cdsSo033Log.FieldByName('REALSTARTDATE').AsString := qrySo033So034.FieldByName('REALSTARTDATE').AsString;
      cdsSo033Log.FieldByName('SBILLNO').AsString := qrySo033So034.FieldByName('SBILLNO').AsString;
      cdsSo033Log.FieldByName('SITEM').AsInteger := qrySo033So034.FieldByName('SITEM').AsInteger;
      cdsSo033Log.FieldByName('Notes').AsString := '�������I-�S�����O�H�����';
      cdsSo033Log.Post;
    end;
end;

function TdtmMain1.CustIDCheck(sI_CustID,sI_CompCode: String): Boolean;
var sL_SQL : String;
begin
    //�Y���Ȥ�w�g�Q��L����,�h�Ǧ^true
    with qryCommon do
    begin
      sL_SQL:='Select COUNT(*) CNT from So123 where CUSTID='''+sI_CustID+'''';
      sL_SQL:=sL_SQL+' and CompCode='+sI_CompCode;
      Close;
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;
      if FieldByName('CNT').AsInteger>0 then
        Result:=True
      else
        Result:=False;
      Close;
    end;

end;

function TdtmMain1.ifPayByCreditCard(sI_PTCode: String): Boolean;
begin
    //�P�_�O�_���H�Υd=>�YSo033/So034.PTCode=3(TCB�H�Υd),PTCode=4(�L�d�H�Υd)(TCB�H�Υd),��ܬ��H�Υd�I�O
    if (sI_PTCode='3') OR (sI_PTCode='4') then
      Result := True
    else
      Result:=False;
end;


procedure TdtmMain1.activeFormula;
var sL_SQL:String;
begin
    with qrySo120Formula do
    begin
      sL_SQL:='Select * from So120';
      Close;
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;
    end;
    {
    qrySo120ComPer.First;
    while not qrySo120ComPer.Eof do
    begin
      cdsSo120.Open;
      cdsSo120.Append;
      cdsSo120.FieldByName('CompCode').AsFloat:=qrySo120ComPer.fieldByName('CompCode').AsFloat;
      cdsSo120.FieldByName('CodeNo').AsFloat:=qrySo120ComPer.fieldByName('CodeNo').AsFloat;
      cdsSo120.FieldByName('Description').AsString:=qrySo120ComPer.fieldByName('Description').AsString;
      cdsSo120.FieldByName('PayUnit').AsString:=qrySo120ComPer.fieldByName('PayUnit').AsString;
      cdsSo120.FieldByName('FirstCreditCard1').AsFloat:=qrySo120ComPer.fieldByName('FirstCreditCard1').AsFloat;
      cdsSo120.FieldByName('FirstNotCreditCard1').AsFloat:=qrySo120ComPer.fieldByName('FirstNotCreditCard1').AsFloat;
      cdsSo120.FieldByName('OtherCreditCard1').AsFloat:=qrySo120ComPer.fieldByName('OtherCreditCard1').AsFloat;
      cdsSo120.FieldByName('OtherNotCreditCard1').AsFloat:=qrySo120ComPer.fieldByName('OtherNotCreditCard1').AsFloat;
      cdsSo120.FieldByName('FirstCreditCard2').AsFloat:=qrySo120ComPer.fieldByName('FirstCreditCard2').AsFloat;
      cdsSo120.FieldByName('FirstNotCreditCard2').AsFloat:=qrySo120ComPer.fieldByName('FirstNotCreditCard2').AsFloat;
      cdsSo120.FieldByName('OtherCreditCard2').AsFloat:=qrySo120ComPer.fieldByName('OtherCreditCard2').AsFloat;
      cdsSo120.FieldByName('OtherNotCreditCard2').AsFloat:=qrySo120ComPer.fieldByName('OtherNotCreditCard2').AsFloat;
      cdsSo120.FieldByName('Ref_No').AsFloat:=qrySo120ComPer.fieldByName('Ref_No').AsFloat;
      cdsSo120.FieldByName('OPERATOR').AsString:=qrySo120ComPer.fieldByName('OPERATOR').AsString;
      cdsSo120.FieldByName('UpdTime').AsString:=qrySo120ComPer.fieldByName('UpdTime').AsString;
      cdsSo120.Post;
      qrySo120ComPer.Next;
    end;
    qrySo120ComPer.Close;
    }
end;

function TdtmMain1.insertSO123(sI_CompCode,sI_CustID, sI_AcceptEn,sI_AcceptName:String):boolean;
var
    bL_Result : boolean;
begin
    bL_Result := true;
    with cdsSo123 do
    begin
      if state=dsInactive then
        Open;

      if not Locate('COMPCODE;CUSTID',VarArrayOf([sI_CompCode,sI_CustID]),[]) then
      begin
        Append;
        FieldByName('COMPCODE').AsInteger := StrToInt(sI_CompCode);//qryCallInCheck.FieldByName('COMPCODE').AsString;
        FieldByName('CUSTID').AsInteger := StrToInt(sI_CustID);
        FieldByName('COMPUTEYM').AsString := sG_ComputeYM;

        FieldByName('COMMDATE').AsDateTime:=StrToDate(FormatDateTime('YYYY/MM/DD',Now));
        FieldByName('CSRID').AsString := sI_AcceptEn;
        FieldByName('CSRNAME').AsString := sI_AcceptName;
        FieldByName('COMM').AsFloat := nG_SelfOrderChannelAcceptComm;
        FieldByName('Operator').AsString := frmMainMenu.sG_User;
        FieldByName('UpdTime').AsString := sG_CurDate;
        Post;
        bL_Result := false;
      end;
    end;
    result := bL_Result;
end;


procedure TdtmMain1.insertCallInDataToSo121(sI_CompCode,sI_ComputeYM,sI_AcceptEn,sI_AcceptName :String);
begin
    with cdsSO121 do
    begin
      if Locate('COMPCODE;BELONGYM;EMPID;COMPUTEYM',VarArrayOf([sI_CompCode,sG_BelongYM,sI_AcceptEn,sG_ComputeYM]),[]) then
      begin
        Edit;
        FieldByName('Commission').AsFloat:=FieldByName('Commission').AsFloat+nG_SelfOrderChannelAcceptComm;
        post;
      end
      else
      begin
        Append;
        FieldByName('COMPCODE').AsInteger := StrToInt(sI_CompCode);
        FieldByName('ComputeYM').AsString := sI_ComputeYM;
        FieldByName('BELONGYM').AsString := sG_BelongYM;
        FieldByName('EMPID').AsString := sI_AcceptEn;
        FieldByName('EMPNAME').AsString := sI_AcceptName;
        FieldByName('COMMISSION').AsFloat:=nG_SelfOrderChannelAcceptComm;

        FieldByName('MediaCode').AsInteger := SERVER_NO;
        FieldByName('MediaName').AsString := SERVER;

        FieldByName('CompType').AsString := '1';
        FieldByName('ClanCompCode').AsInteger := StrToInt(frmMainMenu.sG_CompCode);
        FieldByName('ClanCompName').AsString := frmMainMenu.sG_CompName;
//**

        FieldByName('OPERATOR').AsString:=frmMainMenu.sG_User;
        //FieldByName('UPDTIME').AsString:=IntToStr(StrToInt(FormatDateTime('YYYY',now))-1911)+ FormatDateTime('/MM/DD hh:mm:ss',now);
        FieldByName('UpdTime').AsString := sG_CurDate;
        Post;
      end;
    end;
end;

procedure TdtmMain1.insertCallInDataToSo122(bI_HasPayComm : Boolean;sI_RealBelongYM,sI_PromCode,sI_PromName,sI_CompCode,sI_ComputeYM,sI_AcceptEn,sI_AcceptName,sI_ClctId,sI_ClctName,sI_OrderNo : String; I_DataSet:TDataSet;fI_IntAccOriPercent,fI_ClctEnOriPercent,fI_RealAmt : Double);
var
    dL_RealDate,dL_ShouldDate : TDate;
    sL_CustID : String;
    sL_FirstFlag : String;
begin
    with cdsSo122 do
    begin
      Append;
      FieldByName('CompCode').AsString := sI_CompCode;
      FieldByName('ComputeYM').AsString := sI_ComputeYM;
      FieldByName('BELONGYM').AsString := sI_RealBelongYM;
      FieldByName('BillNo').AsString := I_DataSet.FieldByName('BillNo').AsString;
      FieldByName('Item').AsString := I_DataSet.FieldByName('Item').AsString;
      FieldByName('PeriodId').AsInteger := I_DataSet.FieldByName('REALPERIOD').AsInteger;

      FieldByName('RealAmt').AsFloat := I_DataSet.FieldByName('RealAmt').AsFloat;


      //�ۨӹq���ФH�����ȵ��@��
      if not bI_HasPayComm then
      begin
        FieldByName('IntAccId').AsString := sI_AcceptEn;
        FieldByName('IntAccName').AsString := sI_AcceptName;

        //if fI_IntAccOriPercent <> 0 then  //�Y������ҳ]�w����0,�h�������z�H����
        FieldByName('IntAccComm').AsFloat := nG_SelfOrderChannelAcceptComm;
      end;

      FieldByName('ClctEn').AsString := sI_ClctId;
      FieldByName('ClctEnName').AsString := sI_ClctName;
      FieldByName('ClctEnComm').AsFloat := RoundTo(fG_ChargeComm,-2);
      FieldByName('ClctEnOriPercent').AsFloat := RoundTo(fI_ClctEnOriPercent,-2);


      FieldByName('RealStartDate').AsDateTime := I_DataSet.FieldByName('RealStartDate').AsDateTime;
      FieldByName('RealStopDate').AsDateTime := I_DataSet.FieldByName('RealStopDate').AsDateTime;
      FieldByName('CitemCode').AsInteger := I_DataSet.FieldByName('CitemCode').AsInteger;
      FieldByName('CitemName').AsString := I_DataSet.FieldByName('CitemName').AsString;

      FieldByName('MediaCode').AsInteger := StrToInt(frmMainMenu.sG_CompCode);
      FieldByName('MediaName').AsString := frmMainMenu.sG_CompName;

      FieldByName('OPERATOR').AsString := frmMainMenu.sG_User;
      FieldByName('UpdTime').AsString := sG_CurDate;


      if sI_PromCode <> '' then
        FieldByName('PromCode').AsInteger := StrToInt(sI_PromCode);
      FieldByName('PromName').AsString := sI_PromName;


      dL_RealDate := I_DataSet.FieldByName('RealDate').AsDateTime;
      if dL_RealDate <> 0 then
        FieldByName('RealDate').AsDateTime := dL_RealDate;

      dL_ShouldDate := I_DataSet.FieldByName('ShouldDate').AsDateTime;
      if dL_ShouldDate <> 0 then
        FieldByName('ShouldDate').AsDateTime := dL_ShouldDate;

      sL_CustID := I_DataSet.FieldByName('CustID').AsString;
      if sL_CustID <> '' then
        FieldByName('CustID').AsString := sL_CustID;

      sL_FirstFlag := I_DataSet.FieldByName('FirstFlag').AsString;
      if sL_FirstFlag = '1' then
        FieldByName('FirstFlag').AsString := sL_FirstFlag;


      FieldByName('OrderNo').AsString := sI_OrderNo;

      //�O�ۨӹq
      FieldByName('CallSelf').AsInteger := 1;

      Post;
    end;
end;

procedure TdtmMain1.SaveTempDataTocdsSo121(sI_ClctId, sI_ClctName, sI_IntroId, sI_IntroName,sI_MediaCode,sI_MediaName,sI_RefNo:STring;I_DataSet:TDataset;nI_MonthDiff:Integer);
var
    sL_BelongYM,sL_CompType,sL_ClanCompCode,sL_ClanCompName : String;
begin
    try
    sL_BelongYM := CalculateCountMonths(sG_BelongYM,nI_MonthDiff);

    if sI_RefNo = '2' then       //SO�ۤv���q
    begin
      sL_CompType := '1';
      sL_ClanCompCode := frmMainMenu.sG_CompCode;
      sL_ClanCompName := frmMainMenu.sG_CompName;
    end
    else if sI_RefNo = '3' then  //�P���I
    begin
      sL_CompType := '2';
      sL_ClanCompCode := sI_MediaCode;
      sL_ClanCompName := sI_MediaName;
    end
    else if sI_RefNo = '5' then  //���Y���~���q
    begin
      sL_CompType := '3';
      sL_ClanCompCode := sI_MediaCode;
      sL_ClanCompName := sI_MediaName;
    end;

//SHOWMESSAGE(sL_BelongYM + '--' + sG_ComputeYM + '--' + IntToStr(nI_MonthDiff));
    if (fG_SpreadComm>0) and (sI_IntroId<>'') then
    begin
        with cdsSO121 do
        begin
          if Locate('COMPCODE;BELONGYM;EMPID;COMPUTEYM',VarArrayOf([frmMainMenu.sG_CompCode,sL_BelongYM,sI_IntroId,sG_ComputeYM]),[]) then
          begin
            Edit;
            FieldByName('COMMISSION').AsFloat:= FieldByName('COMMISSION').AsFloat + RoundTo(fG_SpreadComm,-2);
          end
          else
          begin
            Append;
            FieldByName('COMPCODE').AsString := frmMainMenu.sG_CompCode;
            FieldByName('ComputeYM').AsString := sG_ComputeYM;
            FieldByName('BELONGYM').AsString := sL_BelongYM;
            FieldByName('EMPID').AsString := sI_IntroId;
            FieldByName('EMPNAME').AsString := sI_IntroName;
            FieldByName('COMMISSION').AsFloat := RoundTo(fG_SpreadComm,-2);

            FieldByName('MediaCode').AsInteger := StrToInt(sI_MediaCode);
            FieldByName('MediaName').AsString := sI_MediaName;
            FieldByName('CompType').AsString := sL_CompType;

            if sL_ClanCompCode <> '' then
              FieldByName('ClanCompCode').AsInteger := StrToInt(sL_ClanCompCode);
            FieldByName('ClanCompName').AsString := sL_ClanCompName;

            FieldByName('OPERATOR').AsString:=frmMainMenu.sG_User;
            //FieldByName('UPDTIME').AsString:=IntToStr(StrToInt(FormatDateTime('YYYY',now))-1911)+ FormatDateTime('/MM/DD hh:mm:ss',now);
            FieldByName('UpdTime').AsString := sG_CurDate;
          end;
          Post;
        end;
    end;

    //���O��
    if (fG_ChargeComm>0) and (sI_ClctId<>'') then
    begin
        with cdsSO121 do
        begin
          if Locate('COMPCODE;BELONGYM;EMPID;COMPUTEYM',VarArrayOf([frmMainMenu.sG_CompCode,sL_BelongYM,sI_ClctId,sG_ComputeYM]),[]) then
          begin
            Edit;
            FieldByName('COMMISSION').AsFloat:= FieldByName('COMMISSION').AsFloat + RoundTo(fG_ChargeComm,-2);
          end
          else
          begin
            Append;
            FieldByName('COMPCODE').AsString:=frmMainMenu.sG_CompCode;
            FieldByName('BELONGYM').AsString:=CalculateCountMonths(sG_BelongYM,nI_MonthDiff);
            FieldByName('ComputeYM').AsString:=sG_ComputeYM;
            FieldByName('EMPID').AsString:=sI_ClctId;
            FieldByName('EMPNAME').AsString:=sI_ClctName;
            FieldByName('COMMISSION').AsFloat:=RoundTo(fG_ChargeComm,-2);

            FieldByName('MediaCode').AsInteger := CLCTEN_NO;
            FieldByName('MediaName').AsString := CLCTEN;
            FieldByName('CompType').AsString := '1';
            FieldByName('ClanCompCode').AsInteger := StrToInt(frmMainMenu.sG_CompCode);
            FieldByName('ClanCompName').AsString := frmMainMenu.sG_CompName;

            FieldByName('OPERATOR').AsString:=frmMainMenu.sG_User;
            //FieldByName('UPDTIME').AsString:=IntToStr(StrToInt(FormatDateTime('YYYY',now))-1911)+ FormatDateTime('/MM/DD hh:mm:ss',now);
            FieldByName('UpdTime').AsString := sG_CurDate;
          end;
          Post;
        end;
    end;
    except
      ShowMessage('cdsSo121,�s�ɥ���');
      Exit;
    end;
end;

procedure TdtmMain1.SaveTempDataTocdsSo122(sI_RealBelongYM,sI_PromCode,sI_PromName,sI_ClctId, sI_ClctName, sI_IntroId,
sI_IntroName,sI_MediaCode,sI_MediaName,sI_RefNo,sI_OrderNo:String;I_DataSet:TDataset;nI_MonthDiff:Integer;dI_RealStartDate,
dI_RealStopDate : TDate;fI_ClctEnOriPercent,fI_IntAccOriPercent : Double);
var
   dL_ShouldDate,dL_RealDate : TDate;
   sL_CompCode,sL_CustID,sL_CitemCode : String;
   fL_RealAmt : Double;
   sL_BelongYM : String;
begin
    try
      //if ((fG_SpreadComm>0) and (sI_IntroId<>'')) or ((fG_ChargeComm>0) and (sI_ClctId<>'')) then
      if (sI_IntroId<>'') or (sI_ClctId<>'') then
      begin
        with cdsSo122 do
        begin
          Append;
          sL_CompCode := I_DataSet.FieldByName('CompCode').AsString;
          FieldByName('CompCode').AsString:= sL_CompCode;
          FieldByName('ComputeYM').AsString:=sG_ComputeYM;

          sL_BelongYM := CalculateCountMonths(sI_RealBelongYM,nI_MonthDiff);
          FieldByName('BelongYM').AsString := sL_BelongYM;
          //showmessage(FieldByName('BelongYM').AsString);
          FieldByName('BillNo').AsString:=I_DataSet.FieldByName('BillNo').AsString;
          FieldByName('Item').AsString:=I_DataSet.FieldByName('Item').AsString;
          FieldByName('PeriodId').AsInteger:=I_DataSet.FieldByName('REALPERIOD').AsInteger;
          FieldByName('FirstFlag').AsString:=I_DataSet.FieldByName('FirstFlag').AsString;

          fL_RealAmt := I_DataSet.FieldByName('RealAmt').AsFloat;
          FieldByName('RealAmt').AsFloat:= fL_RealAmt;
          if (sI_IntroId<>'') and (fG_SpreadComm>0) then//���s�H
          begin
            FieldByName('IntAccId').AsString:=sI_IntroId;
            FieldByName('IntAccName').AsString:=sI_IntroName;
            FieldByName('IntAccComm').AsFloat:=RoundTo(fG_SpreadComm,-2);
            FieldByName('IntAccOriPercent').AsFloat:= RoundTo(fI_IntAccOriPercent,-2);
            FieldByName('StaffCode').AsString:='3';
          end;

          //���O��
          if sI_ClctId<>'' then
          begin
            FieldByName('CLCTEN').AsString:=sI_ClctId;
            FieldByName('CLCTENNAME').AsString:=sI_ClctName;
            FieldByName('CLCTENCOMM').AsFloat:=RoundTo(fG_ChargeComm,-2);
            FieldByName('ClctEnOriPercent').AsFloat:=RoundTo(fI_ClctEnOriPercent,-2);
          end;



          FieldByName('RealStartDate').AsDateTime := dI_RealStartDate;
          FieldByName('RealStopDate').AsDateTime := dI_RealStopDate;

          FieldByName('OPERATOR').AsString:=frmMainMenu.sG_User;
          //FieldByName('UpdTime').AsString:=IntToStr(StrToInt(FormatDateTime('YYYY',now))-1911)+ FormatDateTime('/MM/DD hh:mm:ss',now);
          FieldByName('UpdTime').AsString := sG_CurDate;

          //CD009 �� RefNo=2(���u����),RefNo=3(�P���I)
          //�YRefno=2,�h����쬰SO CompCode,�_�h��MediaCode
          if sI_RefNo='2' then
          begin
            FieldByName('MediaCode').AsInteger := StrToInt(frmMainMenu.sG_CompCode);
            FieldByName('MediaName').AsString := frmMainMenu.sG_CompName;
          end
          else if (sI_RefNo='3') or (sI_RefNo='5') then
          begin
            if sI_MediaCode <> '' then
              FieldByName('MediaCode').AsInteger := StrToInt(sI_MediaCode);

            FieldByName('MediaName').AsString := sI_MediaName;
          end;

          sL_CitemCode := I_DataSet.FieldByName('CitemCode').AsString;
          if sL_CitemCode <> '' then
            FieldByName('CitemCode').AsInteger := StrToInt(sL_CitemCode);
          FieldByName('CitemName').AsString := I_DataSet.FieldByName('CitemName').AsString;


          if sI_PromCode <> '' then
            FieldByName('PromCode').AsInteger := StrToInt(sI_PromCode);

          FieldByName('PromName').AsString := sI_PromName;


          dL_RealDate := I_DataSet.FieldByName('RealDate').AsDateTime;
          if dL_RealDate <> 0 then
            FieldByName('RealDate').AsDateTime := dL_RealDate;

          dL_ShouldDate := I_DataSet.FieldByName('ShouldDate').AsDateTime;
          if dL_ShouldDate <> 0 then
            FieldByName('ShouldDate').AsDateTime := dL_ShouldDate;

          sL_CustID := I_DataSet.FieldByName('CustID').AsString;
          if sL_CustID <> '' then
            FieldByName('CustID').AsString := sL_CustID;

          FieldByName('OrderNo').AsString := sI_OrderNo;

          //���O�ۨӹq
          FieldByName('CallSelf').AsInteger := 0;

          Post;
        end;
      end;
    except
      ShowMessage('cdsSo122�s�ɥ���');
      Exit;
    end;
end;

function TdtmMain1.CalculateCountMonths(sI_BelongYM :String ;
  nI_RealPeriod: Integer): String;
var nL_Divyy:integer;
    nL_Modmm:integer;
    nL_YearAmt,nL_MonthAmt:integer;
    sL_MM,sL_YY,sL_YYAMT:string;
begin

    sL_YY:=copy(sI_BelongYM,1,3);
    sL_MM:=copy(sI_BelongYM,4,2);

    if nI_RealPeriod>=12 then
      begin
        nL_Divyy:=nI_RealPeriod div 12;    //year
        nL_Modmm:=nI_RealPeriod mod 12;    //Month
                                   //-1 �קK�l�Ƭ�5�� ���� 2005/00
        if nL_Modmm+StrToInt(sL_MM)-1>12 then
        begin
          nL_YearAmt:=nL_Divyy+1;
          nL_MonthAmt:=(StrToInt(sL_MM)+nL_Modmm-1)-12;
          if Length(IntToStr(nL_MonthAmt))=1 then //���Ǧ^2���
          begin
            if Length(IntToStr(StrToInt(sL_YY)))=2 then //���Ǧ^3���
              Result:='0'+IntToStr(StrToInt(sL_YY)+nL_YearAmt)+'0'+IntToStr(nL_MonthAmt)
            else
              Result:=IntToStr(StrToInt(sL_YY)+nL_YearAmt)+'0'+IntToStr(nL_MonthAmt);
          end
          else
          begin
            if Length(IntToStr(StrToInt(sL_YY)))=2 then
              Result:='0'+IntToStr(StrToInt(sL_YY)+nL_YearAmt)+IntToStr(nL_MonthAmt)
            else
              Result:=IntToStr(StrToInt(sL_YY)+nL_YearAmt)+IntToStr(nL_MonthAmt);
          end;
        end;
        if nL_Modmm+StrToInt(sL_MM)-1<=12 then
        begin
          nL_YearAmt:=nL_Divyy;
          nL_MonthAmt:=StrToInt(sL_MM)+nL_Modmm-1;
          if Length(IntToStr(nL_MonthAmt))=1 then
          begin
            if Length(IntToStr(StrToInt(sL_YY)))=2 then
              Result:='0'+IntToStr(StrToInt(sL_YY)+nL_YearAmt)+'0'+IntToStr(nL_MonthAmt)
            else
              Result:=IntToStr(StrToInt(sL_YY)+nL_YearAmt)+'0'+IntToStr(nL_MonthAmt);
          end
          else
          begin
            if Length(IntToStr(StrToInt(sL_YY)))=2 then
              Result:='0'+IntToStr(StrToInt(sL_YY)+nL_YearAmt)+IntToStr(nL_MonthAmt)
            else
              Result:=IntToStr(StrToInt(sL_YY)+nL_YearAmt)+IntToStr(nL_MonthAmt);
          end;
      end;
    end;

    if nI_RealPeriod<12 then
    begin
      if nI_RealPeriod+StrToInt(sL_MM)-1<=12 then
      begin
        nL_MonthAmt:=StrToInt(sL_MM)+nI_RealPeriod-1;
        if Length(IntToStr(nL_MonthAmt))=1 then
        begin
          if length(IntToStr(StrToInt(sL_YY)))=2 then
            Result:=sL_YY+'0'+IntToStr(nL_MonthAmt)
          else
            Result:=sL_YY+'0'+IntToStr(nL_MonthAmt);
        end
        else
          Result:=sL_YY+IntToStr(nL_MonthAmt)
      end;

      if nI_RealPeriod+StrToInt(sL_MM)-1>12 then
      begin
        nL_YearAmt:=StrToInt(sL_YY)+1;
        nL_MonthAmt:=(StrToInt(sL_MM)+nI_RealPeriod-1)-12;

        if Length(IntToStr(nL_MonthAmt))=1 then
        begin
          if Length(IntToStr(nL_YearAmt))=2 then
            Result:='0'+IntToStr(nL_YearAmt)+'0'+IntToStr(nL_MonthAmt)
          else
            Result:=IntToStr(nL_YearAmt)+'0'+IntToStr(nL_MonthAmt);
        end
        else
        begin
          if Length(IntToStr(nL_YearAmt))=2 then
            Result:='0'+IntToStr(nL_YearAmt)+IntToStr(nL_MonthAmt)
          else
            Result:=IntToStr(nL_YearAmt)+IntToStr(nL_MonthAmt);
        end;
      end;
    end;
end;



function TdtmMain1.CheckSo124LockYM(
  sI_ComputeYM: String): boolean;
var
  sL_SQL:String;
  nL_LockYM : Integer;
begin
    with qryCommon do
    begin
      sL_SQL:='Select * from So124';
      Close;
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;
      if RecordCount=0 then
        result := false
      else
      begin
        nL_LockYM:=FieldByName('LockYM').AsInteger;
        if StrToInt(sI_ComputeYM)<=nL_LockYM then
          Result:=true
        else
          Result:=false;

      end;
      Close;
    end;

end;


procedure TdtmMain1.qrySO120CalcFields(DataSet: TDataSet);
var
    nL_PayUnit : Integer;
begin
    nL_PayUnit := DataSet.FieldByName('PAYUNIT').AsInteger;

    if nL_PayUnit = 1 then
      DataSet.FieldByName('PAYUNITNAME').AsString := '�ʤ���'
    else if nL_PayUnit = 2 then
      DataSet.FieldByName('PAYUNITNAME').AsString := '��';
end;

procedure TdtmMain1.deleteCommData(sI_CompCode : String);
    procedure executeDeleteData(sI_TableName,sI_ComputeYM:String);
    var
        sL_SQL : String;
    begin
      sL_SQL := 'delete ' + sI_TableName + ' where COMPUTEYM=' + '''' + sI_ComputeYM  + '''';
      sL_SQL := sL_SQL + ' and COMPCODE=' + sI_CompCode;
      with qryCommon do
      begin
        Close;
        SQL.Clear;
        SQL.Add(sL_SQL);
        ExecSQL;
        Close;
      end;
    end;
begin
    //executeDeleteData('SO121');
    executeDeleteData('SO123',sG_ComputeYM);

    //�������I
    executeDeleteData('SO122',sG_ComputeYM);

    //�@�����I
    executeDeleteData('SO134',sG_OnceComputeYM);
    executeDeleteData('SO135',sG_OnceComputeYM);
end;

procedure TdtmMain1.activePriorData(sI_CompCode : String);
var
    sL_SQL : String;
begin
    with cdsSo121 do
    begin
      Close;
      sL_SQL := 'select * from SO121 where COMPCODE=' + sI_CompCode +
                ' AND COMPUTEYM=' + '''' + sG_ComputeYM + '''';
      CommandText := sL_SQL;
      Open;
    end;

    with cdsSo122 do
    begin
      Close;
      sL_SQL := 'select * from SO122 where COMPCODE=' + sI_CompCode +
                ' AND BELONGYM=' + '''' + sG_BelongYM + '''';
      CommandText := sL_SQL;
      Open;
    end;

    //�d�� SO043 ���O�Ѽ��ɤ� Para3 (�����_�l��P�����I���)  0:�_   1:�O
    sL_SQL := 'SELECT PARA3 FROM SO043 WHERE COMPCODE=' + frmMainMenu.sG_CompCode +
              ' AND SERVICETYPE=''' + frmMainMenu.sG_ServiceType + '''';
    with qryCommon do
    begin
      Close;
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;

      sG_SO043Para3 := FieldByName('PARA3').AsString;

      Close;
    end;


    //�d�� CD019 ���O���إN�X�ɩҦ����
    sL_SQL := 'SELECT CODENO,DESCRIPTION FROM CD019 WHERE StopFlag=0' +
              ' AND SERVICETYPE=''' + frmMainMenu.sG_ServiceType + ''' AND REFNO=2';
    with qryCodeCD019 do
    begin
      Close;
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;

      Close;
    end;

end;

procedure TdtmMain1.applyToDB;
var
    sL_SQL : String;
begin
    //saveSO121ToDB(cdsSo121);
//    save2Db(cdsSo121);
    save2Db(cdsSo122);
    save2Db(cdsSo123);
    save2Db(cdsSo134);
    save2Db(cdsSo135);
end;

procedure TdtmMain1.transReturnDataCdsIntoDB;
var
    sL_CompCode,sL_ComputeYM, sL_BelongYM,sL_EmpID : String;
    sL_EmpName, sL_Operator, sL_SQL : String;
    dL_UpdTime : TDateTime;
    fL_Commission : double ;
begin
    //down, �����N So122 ���h�O��Ƽg�J database
    save2Db(cdsSo122ReturnMoney);
    //up, �����N So122 ���h�O��Ƽg�J database

    //down, loop So121 ���h�O���,�Aupdate ��database
    with cdsSo121ReturnMoney do
    begin
      First;
      while not Eof do
      begin
        sL_CompCode := FieldByName('COMPCODE').AsString;
        sL_ComputeYM := FieldByName('COMPUTEYM').AsString;
        sL_BelongYM := FieldByName('BELONGYM').AsString;
        sL_EmpID := FieldByName('EMPID').AsString;
        sL_EmpName := FieldByName('EMPNAME').AsString;
        sL_Operator := FieldByName('OPERATOR').AsString;
        dL_UpdTime := FieldByName('UPDTIME').AsDateTime;
        fL_Commission := FieldByName('COMMISSION').AsFloat;

        sL_SQL := 'select count(*) COUNT from SO121 where COMPCODE=' + sL_CompCode ;
        sL_SQL := sL_SQL + ' and BELONGYM=' + STR_SEP + sL_BelongYM + STR_SEP;
        sL_SQL := sL_SQL + ' and EMPID=' +  STR_SEP + sL_EmpID + STR_SEP;
        qryCommon.Close;
        qryCommon.SQL.Clear;
        qryCommon.SQL.Add(sL_SQL);
        qryCommon.Open;
        if qryCommon.FieldByNAme('COUNT').asInteger>0 then
        begin
          sL_SQL := 'update SO121 set COMMISSION=COMMISSION-' + FloatToStr(fL_Commission);
          sL_SQL := sL_SQL + ' ,OPERATOR=' + STR_SEP + frmMainMenu.sG_User + STR_SEP ;
          //sL_SQL := sL_SQL + ' ,UPDTIME=' +  TUstr.getOracleSQLDateTimeStr(now) ;
          sL_SQL := sL_SQL + ' ,UPDTIME=''' +  sG_CurDate + '''' ;
          sL_SQL := sL_SQL + ' where COMPCODE=' + sL_CompCode ;
          sL_SQL := sL_SQL + ' and BELONGYM=' + STR_SEP + sL_BelongYM + STR_SEP ;
          sL_SQL := sL_SQL + ' and EMPID=' + STR_SEP + sL_EmpID + STR_SEP ;
          qryCommon.Close;
          qryCommon.SQL.Clear;
          qryCommon.SQL.Add(sL_SQL);
          qryCommon.ExecSQL;
        end
        else
        begin
          sL_SQL := 'insert into SO121(COMPCODE,COMPUTEYM,BELONGYM,EMPID,EMPNAME,COMMISSION,OPERATOR,UPDTIME)';
          sL_SQL := sL_SQL + ' values(' + sL_CompCode + ',' + STR_SEP + sL_ComputeYM + STR_SEP ;
          sL_SQL := sL_SQL + ',' + STR_SEP + sL_BelongYM + STR_SEP ;
          sL_SQL := sL_SQL + ',' + STR_SEP + sL_EmpID + STR_SEP ;
          sL_SQL := sL_SQL + ',' + STR_SEP + sL_EmpName + STR_SEP ;
          sL_SQL := sL_SQL + ',' + FloatToStr(0-fL_Commission);
          sL_SQL := sL_SQL + ',' + STR_SEP + sL_Operator + STR_SEP ;
          sL_SQL := sL_SQL + ',' + STR_SEP + sG_CurDate + STR_SEP + ')';
          qryCommon.Close;
          qryCommon.SQL.Clear;
          qryCommon.SQL.Add(sL_SQL);
          qryCommon.ExecSQL;
        end;
        qryCommon.Close;
        Next;
      end;
    end;
    //up, loop So121 ���h�O���,�Aupdate ��database

end;

procedure TdtmMain1.save2Db(I_CDS: TClientDataSet);
begin
    if I_CDS.State in [dsInsert, dsEdit] then
      I_CDS.Post;
    if I_CDS.ChangeCount>0 then
      I_CDS.ApplyUpdates(0);
end;

procedure TdtmMain1.getCD024Data(sI_PromoteCode: String; var sI_PromoteName,
  sI_DiscountCode, sI_DiscountName: String);
var
    sL_SQL : String;
begin
    sL_SQL := 'SELECT * FROM CD042 WHERE CODENO=' + sI_PromoteCode;

    with qryCD042 do
    begin
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;

      sI_PromoteName := FieldByName('DESCRIPTION').AsString;
      sI_DiscountCode := FieldByName('DISCOUNTCODE').AsString;
      sI_DiscountName := FieldByName('DISCOUNTNAME').AsString;
    end;

end;

function TdtmMain1.getPromoteCode(sI_OrderNo: String;var sI_PromoteName : String): String;
var
    sL_SQL : String;
begin
    sL_SQL := 'SELECT PROMCODE,PromName FROM SO105 WHERE ORDERNO=''' + sI_OrderNo + '''';

    with qryCommon do
    begin
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;

      sI_PromoteName := FieldByName('PromName').AsString;
      Result := FieldByName('PROMCODE').AsString;
    end;
end;


{
function TdtmMain1.getCD009RefNo(sI_MediaCode: String): String;
var
    sL_SQL : String;
begin
    sL_SQL := 'SELECT REFNO FROM CD009 WHERE CODENO=' + sI_MediaCode;
    with qryCommon do
    begin
      Close;
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;

      Result := FieldByName('REFNO').AsString;
    end;
end;
}
procedure TdtmMain1.cdsSo121ZReconcileError(DataSet: TCustomClientDataSet;
  E: EReconcileError; UpdateKind: TUpdateKind;
  var Action: TReconcileAction);
begin
    Action := HandleReconcileError(DataSet, UpdateKind, E);

end;

procedure TdtmMain1.saveSO121ToDB(I_CDS: TClientDataSet);
var
    sL_SQL : String;
    sL_CompCode,sL_ComputeYM,sL_BelongYM,sL_EmpID,sL_EmpName : String;
    sL_MediaCode,sL_MediaName,sL_ClanCompCode,sL_ClanCompName : String;
    sL_CompType,sL_Commission,sL_Operator,sL_UpdTime : String;
begin
    with I_CDS  do
    begin
      First;
      while not Eof do
      begin
        sL_CompCode := I_CDS.FieldByName('CompCode').AsString;
        sL_ComputeYM := I_CDS.FieldByName('ComputeYM').AsString;
        sL_BelongYM := I_CDS.FieldByName('BelongYM').AsString;
        sL_EmpID := I_CDS.FieldByName('EmpID').AsString;
        sL_EmpName := I_CDS.FieldByName('EmpName').AsString;
        sL_Commission := I_CDS.FieldByName('Commission').AsString;
        sL_MediaCode := I_CDS.FieldByName('MediaCode').AsString;
        if sL_MediaCode = '' then
          sL_MediaCode := 'null';

        sL_MediaName := I_CDS.FieldByName('MediaName').AsString;
        sL_CompType := I_CDS.FieldByName('CompType').AsString;
        sL_ClanCompCode := I_CDS.FieldByName('ClanCompCode').AsString;
        if sL_ClanCompCode = '' then
          sL_ClanCompCode := 'null';

        sL_ClanCompName := I_CDS.FieldByName('ClanCompName').AsString;
        sL_Operator := I_CDS.FieldByName('Operator').AsString;
        sL_UpdTime := I_CDS.FieldByName('UpdTime').AsString;

        sL_SQL := 'INSERT INTO SO121 VALUES(' + sL_CompCode + ',''' + sL_ComputeYM +
                  ''',''' + sL_BelongYM + ''',''' + sL_EmpID + ''',''' + sL_EmpName +
                  ''',' + sL_Commission + ',' + sL_MediaCode + ',''' + sL_MediaName +
                  ''',''' + sL_CompType + ''',' + sL_ClanCompCode + ',''' + sL_ClanCompName +
                  ''',''' + sL_Operator + ''',''' + sG_CurDate + ''')';

        with qryCommon do
        begin
          Close;
          SQL.Clear;
          SQL.Add(sL_SQL);
          ExecSQL;
        end;

        Next;
      end;
    end;
end;

procedure TdtmMain1.getComputeDays(sI_BelongYM, sI_RealStartDate,
  sI_RealStopDate: String; var nI_TotalPeriodDays,
  nI_FirstPeriodDays ,nI_PeriodCounts: Integer;var sI_RealBelongYM : String);
var
    dL_RealStartDate,dL_RealStopDate : TDate;
    dL_BelongDate : TDate;
    sL_BelongDate,sL_TempBelongYM : String;
    nL_DiffTotalPeriodDays,nL_DiffFirstPeriodDays,nL_DiffMonths : Integer;
    wL_Year, wL_Month, wL_Day : Word;
begin
    //�p���`���ƤѼƤέ����p�O�Ѽ�
    sL_BelongDate := Copy(sI_BelongYM,1,3) + '/' + Copy(sI_BelongYM,4,5) + '/01';
    dL_BelongDate := StrToDate(chineseDateChangeToEnglishDate(sL_BelongDate));

    dL_RealStartDate := StrToDate(sI_RealStartDate);
    dL_RealStopDate := StrToDate(sI_RealStopDate);
//showmessage('dL_RealStartDate   ' + DateToStr(dL_RealStartDate));
//showmessage('dL_RealStopDate   ' + DateToStr(dL_RealStopDate));

    //�`�g���Ѽ� = RealStopDate - RealStartDate
    nL_DiffTotalPeriodDays := Round(dL_RealStopDate - dL_RealStartDate) + 1;


    //���p�k�ݦ~��W�L���ĺI����
    if dL_BelongDate >= dL_RealStopDate then
    begin
      //�����p�O�Ѽ� = dL_RealStopDate - RealStartDate
      nL_DiffFirstPeriodDays := Round(dL_RealStopDate - dL_RealStartDate) + 1;

      //����k�ݦ~��H���k�ݦ~����w
      sI_RealBelongYM := sI_BelongYM;
    end
    else if dL_BelongDate < dL_RealStartDate then   //���p�k�ݦ~��p�󦳮Ķ}�l���
    begin
      //�k�ݦ~��אּRealStartDate���ݦ~�몺�U�Ӥ�
      DecodeDate(dL_RealStartDate,wL_Year, wL_Month, wL_Day);
      sL_TempBelongYM := afterMonth(format('%.3d%.2d', [wL_Year-1911, wL_Month]),1);
      sL_TempBelongYM := format('%.5d', [StrToInt(sL_TempBelongYM)]);
      //����k�ݦ~��HRealStartDate���U�@�Ӥ�����w
      sI_RealBelongYM := sL_TempBelongYM;


      sL_BelongDate := Copy(sL_TempBelongYM,1,3) + '/' + Copy(sL_TempBelongYM,4,2) + '/01';
      dL_BelongDate := StrToDate(chineseDateChangeToEnglishDate(sL_BelongDate));

      //�����p�O�Ѽ� = �k�ݦ~�몺�e�Ӥ�̫�@�� - RealStartDate
      nL_DiffFirstPeriodDays := Round((dL_BelongDate-1) - dL_RealStartDate) + 1;

      //�p�⦩�������٦��X�� = RealStopDate - (�k�ݦ~��o�e�@�Ӥ�)
      nL_DiffMonths := TUdateTime.countMonth(dL_BelongDate-1,dL_RealStopDate);
    end
    else
    begin
//showmessage('dL_RealStartDate   ' + DateToStr(dL_RealStartDate));
//showmessage('dL_BelongDate   ' + DateToStr(dL_BelongDate-1));
      //�����p�O�Ѽ� = �k�ݦ~�몺�e�Ӥ�̫�@�� - RealStartDate
      nL_DiffFirstPeriodDays := Round((dL_BelongDate-1) - dL_RealStartDate) + 1;

      //�p�⦩�������٦��X�� = RealStopDate - (�k�ݦ~��o�e�@�Ӥ�)
      nL_DiffMonths := TUdateTime.countMonth(dL_BelongDate-1,dL_RealStopDate);

      //����k�ݦ~��H���k�ݦ~����w
      sI_RealBelongYM := sI_BelongYM;
    end;
    //SO043...Para3(�����_�l��P�����I��)  0:�_  1:�O
    if sG_SO043Para3 = '0' then
    begin
      //showmessage('�멳' + sI_BelongYM + sL_BelongDate);
      nI_TotalPeriodDays := nL_DiffTotalPeriodDays;
      nI_FirstPeriodDays := nL_DiffFirstPeriodDays;
    end
    else if sG_SO043Para3 = '1' then
    begin
      //showmessage('���');
      nI_TotalPeriodDays := nL_DiffTotalPeriodDays - 1;

      //���p�k�ݦ~��W�L���ĺI����
      if dL_BelongDate >= dL_RealStopDate then
        nI_FirstPeriodDays := nL_DiffFirstPeriodDays - 1
      else
        nI_FirstPeriodDays := nL_DiffFirstPeriodDays;
    end;

    //���p�k�ݦ~��W�L���ĺI����
    if dL_BelongDate >= dL_RealStopDate then
      nI_PeriodCounts := 1  //�u������
    else
      nI_PeriodCounts := nL_DiffMonths + 1;

//SHOWMESSAGE(DateToStr(dL_RealStopDate) + '--' + DateToStr(dL_BelongDate-1) + '--' + IntToStr(nL_DiffMonths));
end;

function TdtmMain1.chineseDateChangeToEnglishDate(sI_Date: String): String;
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

function TdtmMain1.afterMonth(sI_YM: String; nI_Counts: Integer): String;
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

function TdtmMain1.beforeMonth(sI_YM: String; nI_Counts: Integer): String;
var
    nL_Year,nL_Month,ii : Integer;
    sL_Year,sL_Month,sL_BeforeYM : String;
begin
    //beforeMonth('09109',5) --->>> 9104
    if Length( sI_YM ) <= 5 then
    begin
      nL_Year := StrToInt(Copy(sI_YM,1,3));
      nL_Month := StrToInt(Copy(sI_YM,4,5));
    end else
    begin
      nL_Year := StrToInt(Copy(sI_YM,1,4));
      nL_Month := StrToInt(Copy(sI_YM,5,2));
    end;

    for ii:=1 to nI_Counts do
    begin
      if nL_Month = 1 then
      begin
        nL_Year := nL_Year - 1;
        nL_Month := 12;
      end
      else
      begin
        nL_Month := nL_Month - 1;
      end;
    end;

    sL_Year := IntToStr(nL_Year);
    sL_Month := IntToStr(nL_Month);
    if Length(sL_Month) < 2 then
      sL_Month := '0' + sL_Month;

    sL_BeforeYM := sL_Year + sL_Month;

    Result := sL_BeforeYM;

end;

function TdtmMain1.handleInCome(sI_TableName, sI_YM,sI_PayMode: String): String;
var  sL_SQL,sL_OrderNo,sL_TempString1,sL_TempString2 : String;
     bL_Commit,bL_CreditCard,bL_HasReturnData : Boolean;
     sL_SBillNo,sL_SItem,sL_CitemCode,sL_PromoteCode,sL_DiscountCode : String;
     sL_StartInstDate,sL_StopInstDate,sL_MonthDays : String;
     sL_CitemName,sL_DiscountName,sL_PromoteName : String;
     bL_HaveSetData : Boolean;
     sL_BillNo,sL_Item : String;
begin
    sL_StartInstDate := sI_YM + '/01';
    sL_MonthDays := IntToStr(TUdateTime.DaysOfMonth(sL_StartInstDate));
    sL_StopInstDate := sI_YM + '/' + sL_MonthDays + ' 23:59:59';
    sL_StartInstDate := sL_StartInstDate + ' 00:00:00';

    sL_SQL:='SELECT A.RefNo,B.CustId,B.OrderNo,B.FirstFlag,';
    sL_SQL:=sL_SQL+'B.DiscountCode,B.DiscountName,B.OrderNo,TO_DATE(TO_CHAR(B.RealStartDate,'+''''+'YYYY/MM/DD'+''''+'),'+''''+'YYYY/MM/DD'+''''+') RealStartDate ,'+'TO_DATE(TO_CHAR(B.RealStopDate,'+''''+'YYYY/MM/DD'+''''+'),'+''''+'YYYY/MM/DD'+''''+') RealStopDate ,';
    sL_SQL:=sL_SQL+'B.CMName,B.CitemCode,B.CitemName,B.PTCode,B.RealAmt,to_char(B.RealAmt) StrRealAmt ,';
    sL_SQL:=sL_SQL+'B.CompCode,B.ClctEn,B.ClctName,B.BillNo,';
    sL_SQL:=sL_SQL+'B.Item, B.RealPeriod, B.SBILLNO, B.SITEM ,TO_DATE(TO_CHAR(B.RealDate,'+''''+'YYYY/MM/DD'+''''+'),'+''''+'YYYY/MM/DD'+''''+') RealDate,TO_DATE(TO_CHAR(B.ShouldDate,'+''''+'YYYY/MM/DD'+''''+'),'+''''+'YYYY/MM/DD'+''''+') ShouldDate FROM CD019 A, '+sI_TableName+'';
    sL_SQL:=sL_SQL+' B Where B.CancelFlag=0';
    sL_SQL:=sL_SQL+' AND B.UcCode is NULL and B.ServiceType=''' + frmMainMenu.sG_ServiceType + '''';
    sL_SQL := sL_SQL + ' AND (B.RealDate BETWEEN ' + TUstr.getOracleSQLDateTimeStr(StrToDateTime(sL_StartInstDate));
    sL_SQL := sL_SQL + ' AND ' + TUstr.getOracleSQLDateTimeStr(StrToDateTime(sL_StopInstDate))+')';
    sL_SQL:=sL_SQL+' AND B.CitemCode=A.CodeNo and A.RefNo=2 and (b.OrderNo is not null or b.OrderNo<>'''') ';
    //sL_SQL:=sL_SQL+ 'AND (SYSDATE - B.RealDate) > ' + IntToStr(MONTH_DAYS);


    sL_SQL:=sL_SQL+' AND B.COMPCODE='''+frmMainMenu.sG_CompCode+'''';
    sL_SQL:=sL_SQL+' Order by B.ClctEn';

//showmessage('920923_Jackal_BILLNO');
    //sL_SQL:=sL_SQL+' AND B.COMPCODE='''+frmMainMenu.sG_CompCode+''' AND B.CITEMCODE IN (931,932,933) AND CUSTID=171702';
    //sL_SQL:=sL_SQL+' AND B.COMPCODE='''+frmMainMenu.sG_CompCode+''' AND BILLNO=''200308TC0129591''';
    //sL_SQL:=sL_SQL+' AND B.COMPCODE='''+frmMainMenu.sG_CompCode+''' AND FirstFlag<>''1''';
    //sL_SQL:=sL_SQL+' AND B.COMPCODE='''+frmMainMenu.sG_CompCode+''' AND custid=12186';
    //sL_SQL:=sL_SQL+' Order by B.ClctEn';


    with qrySo033So034 do
    begin
      Close;
      SQL.Clear;
      SQL.Add(sL_SQL);
      OPEN;
    end;


    //showmessage(IntTosTR(qrySo033So034.RecordCount));
    while not qrySo033So034.Eof do
    begin
      fG_RealAmt:=0;
      sG_CharMeth:='';
      sG_CustId:='';

      sG_CodeNo:='';

      fG_FirCardSpd:=0;
      fG_FirNoCardSpd:=0;
      fG_FirCardCharge:=0;
      fG_FirNoCardCharge:=0;
      fG_ContCardCharge:=0;
      fG_ContNoCardCharge:=0;
      sG_PayUnit:='';
      nG_RealPeriod:=0;


      with qrySo120Formula do
      begin
        sL_CitemCode := qrySo033So034.FieldByName('CitemCode').AsString;
        sL_DiscountCode := qrySo033So034.FieldByName('DiscountCode').AsString;
        sL_DiscountName := qrySo033So034.FieldByName('DiscountName').AsString;
        sL_OrderNo := qrySo033So034.FieldByName('OrderNo').AsString;
        sL_PromoteCode := getPromoteCode(sL_OrderNo,sL_PromoteName);

{
        //�n�T�{�u�f���ʻP�~�Ȭ��ʪ��ϧO
        if (sL_PromoteCode <> '') and (sL_DiscountCode <> '') then
        begin
          if Locate('Compcode;CodeNo;PROMOTECODE;DISCOUNTCODE',VarArrayof([sG_CompCode,sL_CitemCode,StrToInt(sL_PromoteCode),StrToInt(sL_DiscountCode)]),[]) then
          begin
            getAccountParameter(qrySo120Formula);
            bL_HaveSetData := true;
          end
          else
            bL_HaveSetData := false;
        end
        else if (sL_PromoteCode <> '') and (sL_DiscountCode = '') then
        begin
          if Locate('Compcode;CodeNo;PROMOTECODE;DISCOUNTCODE',VarArrayof([sG_CompCode,sL_CitemCode,StrToInt(sL_PromoteCode),null]),[]) then
          begin
            getAccountParameter(qrySo120Formula);
            bL_HaveSetData := true;
          end
          else
            bL_HaveSetData := false;
        end
        else if (sL_PromoteCode = '') and (sL_DiscountCode <> '') then
        begin
          if Locate('Compcode;CodeNo;PROMOTECODE;DISCOUNTCODE',VarArrayof([sG_CompCode,sL_CitemCode,null,StrToInt(sL_DiscountCode)]),[]) then
          begin
            getAccountParameter(qrySo120Formula);
            bL_HaveSetData := true;
          end
          else
            bL_HaveSetData := false;
        end
        else if(sL_PromoteCode = '') and (sL_DiscountCode = '') then
        begin
          if Locate('Compcode;CodeNo;PROMOTECODE;DISCOUNTCODE',VarArrayof([sG_CompCode,sL_CitemCode,null,null]),[]) then
          begin
            getAccountParameter(qrySo120Formula);
            bL_HaveSetData := true;
          end
          else
            bL_HaveSetData := false;
        end
        else
          bL_HaveSetData := false;

}


        //�ثe���޷~�Ȭ���(Jackal 920530)
        if Locate('Compcode;CodeNo;PROMOTECODE;DISCOUNTCODE',VarArrayof([frmMainMenu.sG_CompCode,sL_CitemCode,null,null]),[]) then
        begin
          getAccountParameter(qrySo120Formula);
          bL_HaveSetData := true;
        end
        else
          bL_HaveSetData := false;

      end;



      //���������/���B�]�w��Ƥ~�p��
      if bL_HaveSetData then
      begin
        //�YCD019.RefNo=2,��ܸӵ����O���I�O�W�D���g���ʦ��O����=>�ݭn������
        if qrySo033So034.FieldByName('RefNo').AsString='2' then
        begin
          if qrySo033So034.FieldByName('OrderNo').AsString<>'' then
          begin
            if sI_PayMode = BATCH_PAY then     //�������I�p��
              bL_Commit := processIncome(qrySo033So034.FieldByName('OrderNo').AsString,sI_YM, qrySo033So034)
            else if sI_PayMode = ONCE_PAY then //�@�����I�p��
              bL_Commit := processOnceIncome(qrySo033So034.FieldByName('OrderNo').AsString,sI_YM, qrySo033So034);
          end
          else
          begin  //�S��OrderNo ��������C�餶�и��
            cdsSo033Log.Append;
            cdsSo033Log.FieldByName('CUSTID').AsString := qrySo033So034.FieldByName('CUSTID').AsString;
            cdsSo033Log.FieldByName('BILLNO').AsString := qrySo033So034.FieldByName('BILLNO').AsString;
            cdsSo033Log.FieldByName('ITEM').AsString := qrySo033So034.FieldByName('ITEM').AsString;
            cdsSo033Log.FieldByName('CITEMCODE').AsString := qrySo033So034.FieldByName('CITEMCODE').AsString;
            cdsSo033Log.FieldByName('CITEMNAME').AsString := qrySo033So034.FieldByName('CITEMNAME').AsString;
            cdsSo033Log.FieldByName('SHOULDDATE').AsString := qrySo033So034.FieldByName('SHOULDDATE').AsString;
            cdsSo033Log.FieldByName('REALDATE').AsString := qrySo033So034.FieldByName('REALDATE').AsString;
            cdsSo033Log.FieldByName('REALAMT').AsString := qrySo033So034.FieldByName('REALAMT').AsString;
            cdsSo033Log.FieldByName('REALPERIOD').AsString := qrySo033So034.FieldByName('RealPeriod').AsString;
            cdsSo033Log.FieldByName('REALSTARTDATE').AsString := qrySo033So034.FieldByName('REALSTARTDATE').AsString;
            cdsSo033Log.FieldByName('SBILLNO').AsString := qrySo033So034.FieldByName('SBILLNO').AsString;
            cdsSo033Log.FieldByName('SITEM').AsInteger := qrySo033So034.FieldByName('SITEM').AsInteger;

            if sI_PayMode = BATCH_PAY then
              cdsSo033Log.FieldByName('Notes').AsString := '�������I-�S���q��渹��������C�餶�и��'
            else if sI_PayMode = ONCE_PAY then
              cdsSo033Log.FieldByName('Notes').AsString := '�@�����I-�S���q��渹��������C�餶�и��';

            cdsSo033Log.Post;
          end;
        end;
      end
      else  //�Y�䤣��������/���B�]�w���
      begin
        cdsSo033Log.Append;
        cdsSo033Log.FieldByName('CUSTID').AsString := qrySo033So034.FieldByName('CUSTID').AsString;
        cdsSo033Log.FieldByName('BILLNO').AsString := qrySo033So034.FieldByName('BILLNO').AsString;
        cdsSo033Log.FieldByName('ITEM').AsString := qrySo033So034.FieldByName('ITEM').AsString;
        cdsSo033Log.FieldByName('CITEMCODE').AsString := qrySo033So034.FieldByName('CITEMCODE').AsString;
        cdsSo033Log.FieldByName('CITEMNAME').AsString := qrySo033So034.FieldByName('CITEMNAME').AsString;
        cdsSo033Log.FieldByName('SHOULDDATE').AsString := qrySo033So034.FieldByName('SHOULDDATE').AsString;
        cdsSo033Log.FieldByName('REALDATE').AsString := qrySo033So034.FieldByName('REALDATE').AsString;
        cdsSo033Log.FieldByName('REALAMT').AsString := qrySo033So034.FieldByName('REALAMT').AsString;
        cdsSo033Log.FieldByName('REALPERIOD').AsString := qrySo033So034.FieldByName('RealPeriod').AsString;
        cdsSo033Log.FieldByName('REALSTARTDATE').AsString := qrySo033So034.FieldByName('REALSTARTDATE').AsString;
        cdsSo033Log.FieldByName('SBILLNO').AsString := qrySo033So034.FieldByName('SBILLNO').AsString;
        cdsSo033Log.FieldByName('SITEM').AsInteger := qrySo033So034.FieldByName('SITEM').AsInteger;

        if sI_PayMode = BATCH_PAY then
          cdsSo033Log.FieldByName('Notes').AsString := '�������I-�䤣��������/���B�]�w���'
        else if sI_PayMode = ONCE_PAY then
          cdsSo033Log.FieldByName('Notes').AsString := '�@�����I-�䤣��������/���B�]�w���';

        cdsSo033Log.FieldByName('PromCode').AsString := sL_PromoteCode;
        cdsSo033Log.FieldByName('PromName').AsString := sL_PromoteName;
        cdsSo033Log.FieldByName('DiscountCode').AsString := sL_DiscountCode;
        cdsSo033Log.FieldByName('DiscountName').AsString := sL_DiscountName;
        cdsSo033Log.Post;
      end;
      qrySo033So034.Next;
    end;
end;

function TdtmMain1.handleOutCome(sI_TableName, sI_YM: String): String;
var
    sL_SQL,sL_OrderNo,sL_TempString1,sL_TempString2 : String;
    //bL_Commit,bL_CreditCard,bL_HasReturnData : Boolean;
    sL_SBillNo,sL_SItem,sL_FunctionNo : String;
    sL_BelongDate,sL_Msg : String;
    dL_BelongDate,dL_RealStartDate,dL_RealStopDate,dL_ShouldDate,dL_ViewStartDate,dL_ViewStopDate : TDate;
    fL_RealAmt : Double;
    nL_DiffDays,nL_RetCode : Integer;
    bL_HaveData : Boolean;
    sL_StartShouldDate,sL_StopShouldDate,sL_MonthDays : String;
begin
    //�ȳB�z���ФH/���z�H�h�O,�]���ФH�������@���N�u��U��(���O�H����C��|�X)
    sL_BelongDate := Copy(sG_BelongYM,1,3) + '/' + Copy(sG_BelongYM,4,5) + '/01';
    dL_BelongDate := StrToDate(chineseDateChangeToEnglishDate(sL_BelongDate));

    sL_StartShouldDate := sI_YM + '/01';
    sL_MonthDays := IntToStr(TUdateTime.DaysOfMonth(sL_StartShouldDate));
    sL_StopShouldDate := sI_YM + '/' + sL_MonthDays + ' 23:59:59';
    sL_StartShouldDate := sL_StartShouldDate + ' 00:00:00';

    sL_SQL:='SELECT A.RefNo,B.CustId,B.OrderNo,B.FirstFlag,TO_DATE(TO_CHAR(B.SHOULDDATE,'+''''+'YYYY/MM/DD'+''''+'),'+''''+'YYYY/MM/DD'+''''+') SHOULDDATE,';
    sL_SQL:=sL_SQL+'B.DiscountCode,B.OrderNo,TO_DATE(TO_CHAR(B.RealStartDate,'+''''+'YYYY/MM/DD'+''''+'),'+''''+'YYYY/MM/DD'+''''+') RealStartDate,TO_DATE(TO_CHAR(B.RealStopDate,'+''''+'YYYY/MM/DD'+''''+'),'+''''+'YYYY/MM/DD'+''''+') RealStopDate,';
    sL_SQL:=sL_SQL+'B.CMName,B.CitemCode,B.CitemName,B.PTCode,TO_DATE(TO_CHAR(B.RealDate,'+''''+'YYYY/MM/DD'+''''+'),'+''''+'YYYY/MM/DD'+''''+') RealDate,B.RealAmt,to_char(B.RealAmt) StrRealAmt,';
    sL_SQL:=sL_SQL+'B.CompCode,B.ClctEn,B.ClctName,B.BillNo,';
    sL_SQL:=sL_SQL+'B.Item, B.REALPERIOD,B.SBILLNO, B.SITEM FROM CD019 A, '+sI_TableName+'';
    sL_SQL:=sL_SQL+' B Where B.CancelFlag=0';
    sL_SQL:=sL_SQL+' AND B.UcCode is NULL and B.ServiceType=''' + frmMainMenu.sG_ServiceType + '''';

    sL_SQL:=sL_SQL+' AND (B.ShouldDate BETWEEN ' + TUstr.getOracleSQLDateTimeStr(StrToDateTime(sL_StartShouldDate));
    sL_SQL:=sL_SQL+' AND ' + TUstr.getOracleSQLDateTimeStr(StrToDateTime(sL_StopShouldDate)) + ')';

    sL_SQL:=sL_SQL+' AND B.CitemCode=A.CodeNo ';
    sL_SQL:=sL_SQL+' AND B.COMPCODE='''+frmMainMenu.sG_CompCode+''' AND A.RefNo=9 ';

//showmessage('920915_Jackal_�B�z�h�O���');
    //sL_SQL:=sL_SQL+' AND B.SBILLNO = ''200304IC0162053''';

    sL_SQL:=sL_SQL+' Order by B.ClctEn';
{
    with qrySo033So034 do
    begin
      Close;
      SQL.Clear;
      SQL.Add(sL_SQL);
      OPEN;
    end;
}

    qrySo033So034.Close;
    with cdsSo033So034 do
    begin
      Close;
      CommandText := sL_SQL;
      OPEN;
    end;

    //���t��
    transSo033So034Notation(cdsSo033So034);

    cdsSo033So034.First;
    while not cdsSo033So034.Eof do
    begin
      sL_SBillNo := cdsSo033So034.FieldByName('SBillNo').AsString;
      sL_SItem := cdsSo033So034.FieldByName('SItem').AsString;
      dL_ShouldDate := cdsSo033So034.FieldByName('ShouldDate').AsDateTime;
      fL_RealAmt := cdsSo033So034.FieldByName('RealAmt').AsFloat;
      dL_RealStartDate:= cdsSo033So034.FieldByName('RealStartDate').AsDateTime;
      dL_RealStopDate:= cdsSo033So034.FieldByName('RealStopDate').AsDateTime;
//so105
      //�Y�h�O��Ʀ� RealStartDate �� RealStopDate �h�Φۨ���RealStartDate �� RealStopDate
      //�Y�h�O��ƨS�� RealStartDate �� RealStopDate �h�Υ�����RealStartDate �� RealStopDate
      if (dL_RealStartDate =0) AND  (dL_RealStopDate =0) then //�h�O��ƨS�� RealStartDate �� RealStopDate
      begin
        if (sL_SBillNo<>'') and (sL_SItem<>'')
        and (cdsSo122.Locate('COMPCODE;BILLNO;ITEM',VarArrayOf([frmMainMenu.sG_CompCode,sL_SBillNo,sL_SItem]),[])) then
        begin
          //�ѭt��(�h�O)��SBillNo��SItem��X�������������(����P������������)
          dL_ViewStartDate := cdsSo122.FieldByName('RealStartDate').AsDateTime;
          dL_ViewStopDate := cdsSo122.FieldByName('RealStopDate').AsDateTime;

          sG_SFFirstFlag := cdsSo122.FieldByName('FirstFlag').AsString;

          bL_HaveData := true;
        end
        else if (sL_SBillNo<>'') and (sL_SItem<>'') then
        begin
          //��SO122��D���p��~�몺�������
          bL_HaveData := callSF_GetChargeData('SO122', frmMainMenu.sG_CompCode, sG_OnceBelongYM,
                               sL_SBillNo, sL_SItem,sL_Msg,nL_RetCode);

          //�N�Ǧ^����ƨ��X
          if bL_HaveData then
            getSFReturnData(sL_Msg);

          dL_ViewStartDate := dG_SFRealStartDate;
          dL_ViewStopDate := dG_SFRealStopDate;
        end
        else
        begin
          bL_HaveData := false;
        end;


        if (bL_HaveData=true) AND (sG_SFFirstFlag='1')then
        begin
          //�ˬd�O�_�������W�L�@�Ӥ�     MONTH_DAYS
          nL_DiffDays := Round((dL_ShouldDate - dL_ViewStartDate));
          //�����@�Ӥ���������B�h,����������h(�]�����O������) Jackal_920905
          //if nL_DiffDays >= nG_ChannelViewDays then
          //begin
            //�ˬd�k�ݦ~��O�_�W�L RealStopDate
            if dL_BelongDate < dL_ViewStopDate then
            begin
              sL_FunctionNo := '1';
              //�t���� ShouldDate �N��h�O���_�l���
              if nL_DiffDays >= nG_ChannelViewDays then
                ComputeReturnData1(DateToStr(dL_ShouldDate),sL_SBillNo,sL_SItem,sL_FunctionNo,fL_RealAmt,cdsSo033So034)
              else
                ComputeReturnData1(DateToStr(dL_ViewStartDate),sL_SBillNo,sL_SItem,sL_FunctionNo,fL_RealAmt,cdsSo033So034);
            end
            else    //�k�ݦ~��W�L RealStopDate
            begin
              sL_FunctionNo := '2';
              //�t���� ShouldDate �N��h�O���_�l���
              if nL_DiffDays >= nG_ChannelViewDays then
                ComputeReturnData1(DateToStr(dL_ShouldDate),sL_SBillNo,sL_SItem,sL_FunctionNo,fL_RealAmt,cdsSo033So034)
              else
                ComputeReturnData1(DateToStr(dL_ViewStartDate),sL_SBillNo,sL_SItem,sL_FunctionNo,fL_RealAmt,cdsSo033So034);
            end;
          //end
          //else      //�����@�Ӥ�ҥH�������B�h
          //begin
            //�����@�Ӥ�]������������,�ҥH���ΰh�O  Jackal920520
          //  sL_FunctionNo := '3';
            //���B�h�h�H�������_�l������h�O���_�l���
          //  ComputeReturnData1(DateToStr(dL_ViewStartDate),sL_SBillNo,sL_SItem,sL_FunctionNo,fL_RealAmt,cdsSo033So034);
          //end;
        end
        else if sG_SFFirstFlag<>'1' then  //�򦬪���Ƥ��ΰh�O
        begin

        end
        else
        begin
          cdsSo033Log.Append;
          cdsSo033Log.FieldByName('CUSTID').AsString:=cdsSo033So034.FieldByName('CUSTID').AsString;
          cdsSo033Log.FieldByName('BILLNO').AsString:=cdsSo033So034.FieldByName('BILLNO').AsString;
          cdsSo033Log.FieldByName('ITEM').AsString:=cdsSo033So034.FieldByName('ITEM').AsString;
          cdsSo033Log.FieldByName('CITEMCODE').AsString:=cdsSo033So034.FieldByName('CITEMCODE').AsString;
          cdsSo033Log.FieldByName('CITEMNAME').AsString:=cdsSo033So034.FieldByName('CITEMNAME').AsString;
          cdsSo033Log.FieldByName('SHOULDDATE').AsString:=cdsSo033So034.FieldByName('SHOULDDATE').AsString;
          cdsSo033Log.FieldByName('REALDATE').AsString:=cdsSo033So034.FieldByName('REALDATE').AsString;
          cdsSo033Log.FieldByName('REALAMT').AsString:=cdsSo033So034.FieldByName('REALAMT').AsString;
          cdsSo033Log.FieldByName('REALPERIOD').AsString:=cdsSo033So034.FieldByName('RealPeriod').AsString;
          cdsSo033Log.FieldByName('REALSTARTDATE').AsString:=cdsSo033So034.FieldByName('REALSTARTDATE').AsString;
          cdsSo033Log.FieldByName('SBILLNO').AsString:=cdsSo033So034.FieldByName('SBILLNO').AsString;
          cdsSo033Log.FieldByName('SITEM').AsInteger:=cdsSo033So034.FieldByName('SITEM').AsInteger;
          cdsSo033Log.FieldByName('Notes').AsString := '�������I-�h�O��ƹ������쥿�����';
          cdsSo033Log.Post;
        end;

      end
      else if (dL_RealStartDate<>0) and (dL_RealStopDate<>0) then //�h�O��Ʀ� RealStartDate �� RealStopDate
      begin
        if (sL_SBillNo<>'') and (sL_SItem<>'')
        and (cdsSo122.Locate('COMPCODE;BILLNO;ITEM',VarArrayOf([frmMainMenu.sG_CompCode,sL_SBillNo,sL_SItem]),[])) then
        begin
          //�ѭt��(�h�O)��SBillNo��SItem��X�������������(����P������������)
          dL_ViewStartDate := cdsSo122.FieldByName('RealStartDate').AsDateTime;

          bL_HaveData := true;
        end
        else if (sL_SBillNo<>'') and (sL_SItem<>'') then
        begin
//if sL_SBillNo = '200302IC0134225' then
//  sL_SBillNo := sL_SBillNo;
          //��SO122��D���p��~�몺�������
          bL_HaveData := callSF_GetChargeData('SO122', frmMainMenu.sG_CompCode, sG_OnceBelongYM,
                               sL_SBillNo, sL_SItem,sL_Msg,nL_RetCode);

          //�N�Ǧ^����ƨ��X
          if bL_HaveData then
            getSFReturnData(sL_Msg);

          dL_ViewStartDate := dG_SFRealStartDate;
          dL_ViewStopDate := dG_SFRealStopDate;
        end
        else
        begin
          bL_HaveData := false;
        end;



        if (bL_HaveData=true) AND (sG_SFFirstFlag='1')then
        begin
          nL_DiffDays := Round((dL_RealStartDate - dL_ViewStartDate));
          //if nL_DiffDays >= nG_ChannelViewDays then
          //begin
            //�ˬd�k�ݦ~��O�_�W�L RealStopDate
            if dL_BelongDate < dL_RealStopDate then
            begin
              sL_FunctionNo := '1';
              //�t���� RealStartDate �N��h�O���_�l���
              if nL_DiffDays >= nG_ChannelViewDays then
                ComputeReturnData1(DateToStr(dL_RealStartDate),sL_SBillNo,sL_SItem,sL_FunctionNo,fL_RealAmt,cdsSo033So034)
              else
                ComputeReturnData1(DateToStr(dL_ViewStartDate),sL_SBillNo,sL_SItem,sL_FunctionNo,fL_RealAmt,cdsSo033So034);
            end
            else    //�k�ݦ~��W�L RealStopDate
            begin
              sL_FunctionNo := '2';
              //�t���� RealStartDate �N��h�O���_�l���
              if nL_DiffDays >= nG_ChannelViewDays then
                ComputeReturnData1(DateToStr(dL_RealStartDate),sL_SBillNo,sL_SItem,sL_FunctionNo,fL_RealAmt,cdsSo033So034)
              else
                ComputeReturnData1(DateToStr(dL_ViewStartDate),sL_SBillNo,sL_SItem,sL_FunctionNo,fL_RealAmt,cdsSo033So034);
            end;
          //end
          //else      //�����@�Ӥ�ҥH�������B�h
          //begin
            //�����@�Ӥ�]������������,�ҥH���ΰh�O  Jackal920520
            //sL_FunctionNo := '3';
            //���B�h�h�H�������_������h�O���_�l���
            //ComputeReturnData1(DateToStr(dL_ViewStartDate),sL_SBillNo,sL_SItem,sL_FunctionNo,fL_RealAmt,cdsSo033So034);
          //end;
        end
        else if sG_SFFirstFlag<>'1' then  //�򦬪���Ƥ��ΰh�O
        begin

        end
        else
        begin
          cdsSo033Log.Append;
          cdsSo033Log.FieldByName('CUSTID').AsString:=cdsSo033So034.FieldByName('CUSTID').AsString;
          cdsSo033Log.FieldByName('BILLNO').AsString:=cdsSo033So034.FieldByName('BILLNO').AsString;
          cdsSo033Log.FieldByName('ITEM').AsString:=cdsSo033So034.FieldByName('ITEM').AsString;
          cdsSo033Log.FieldByName('CITEMCODE').AsString:=cdsSo033So034.FieldByName('CITEMCODE').AsString;
          cdsSo033Log.FieldByName('CITEMNAME').AsString:=cdsSo033So034.FieldByName('CITEMNAME').AsString;
          cdsSo033Log.FieldByName('SHOULDDATE').AsString:=cdsSo033So034.FieldByName('SHOULDDATE').AsString;
          cdsSo033Log.FieldByName('REALDATE').AsString:=cdsSo033So034.FieldByName('REALDATE').AsString;
          cdsSo033Log.FieldByName('REALAMT').AsString:=cdsSo033So034.FieldByName('REALAMT').AsString;
          cdsSo033Log.FieldByName('REALPERIOD').AsString:=cdsSo033So034.FieldByName('RealPeriod').AsString;
          cdsSo033Log.FieldByName('REALSTARTDATE').AsString:=cdsSo033So034.FieldByName('REALSTARTDATE').AsString;
          cdsSo033Log.FieldByName('SBILLNO').AsString:=cdsSo033So034.FieldByName('SBILLNO').AsString;
          cdsSo033Log.FieldByName('SITEM').AsInteger:=cdsSo033So034.FieldByName('SITEM').AsInteger;
          cdsSo033Log.FieldByName('Notes').AsString := '�������I-�h�O��ƹ������쥿�����';
          cdsSo033Log.Post;
        end;
      end;

      cdsSo033So034.Next;
    end;


    end;

procedure TdtmMain1.ComputeReturnData1(sI_ShouldDate,sI_SBillNo,sI_SItem,
sI_FunctionNo : String;fI_RealAmt : Double;I_CDSSo033So034 : TClientDataSet);
var
    ii,nL_TotalPeriodDays,nL_FirstPeriodDays,nL_PeriodCounts,nL_OtherPeriodDays : Integer;
    sL_NextYM,sL_NextBelongYM, sL_MediaCodeA,sL_MediaNameA : String;
    fL_ReturnComm,fL_RealAmt : Double;
    dL_RealStartDate,dL_RealStopDate : TDate;
    fL_IntAccOriPercent,fL_ClctEnOriPercent,fL_IntAccComm : Double;

    sL_CompCode,sL_ComputeYM,sL_BelongYM,sL_StbNo,sL_SNo : String;
    sL_BillNo,sL_Item,sL_Period,sL_WorkerEn1,sL_WorkerEn1Name : String;
    sL_WorkerEn1Comm,sL_ClctEn,sL_ClctEnName,sL_ClctEnComm : String;
    sL_IntAccId,sL_IntAccName,sL_StaffCode : String;
    sL_MediaCode,sL_MediaName : String;
    sL_CompType,sL_ClanCompCode,sL_ClanCompName : String;
    sL_CitemCode,sL_CitemName,sL_MonthFirstDay : String;
    sL_CompTypeA,sL_ClanCompCodeA,sL_ClanCompNameA : String;
    sL_PromCode,sL_PromName,sL_CustID,sL_RealBelongYM : String;
    dL_RealDate,dL_ShouldDate : TDate;
    sL_FirstFlag,sL_OrderNo,sL_CallSelf : String;


    procedure transIntoSo121(I_CDS:TClientDataset);
    begin
      I_CDS.Append;
      I_CDS.FieldByName('COMPCODE').AsString := frmMainMenu.sG_CompCode;
      I_CDS.FieldByName('COMPUTEYM').AsString := sL_ComputeYM;
      I_CDS.FieldByName('BELONGYM').AsString := sL_BelongYM;
      I_CDS.FieldByName('EMPID').AsString := sL_IntAccId;
      I_CDS.FieldByName('EMPNAME').AsString := sL_IntAccName;
      I_CDS.FieldByName('COMMISSION').AsFloat := fL_IntAccComm;

      I_CDS.FieldByName('MediaCode').AsString := sL_MediaCodeA;
      I_CDS.FieldByName('MediaName').AsString := sL_MediaNameA;
      I_CDS.FieldByName('CompType').AsString := sL_CompTypeA;
      I_CDS.FieldByName('ClanCompCode').AsString := sL_ClanCompCodeA;
      I_CDS.FieldByName('ClanCompName').AsString := sL_ClanCompNameA;

      I_CDS.FieldByName('OPERATOR').AsString := frmMainMenu.sG_User;
      I_CDS.FieldByName('UPDTIME').AsString := sG_CurDate;
      I_CDS.Post;
    end;

    procedure transIntoSo122(I_CDS:TClientDataset);
    begin
      I_CDS.Append;
      I_CDS.FieldByName('COMPCODE').AsString := sL_CompCode;
      I_CDS.FieldByName('COMPUTEYM').AsString := sL_ComputeYM;
      I_CDS.FieldByName('BELONGYM').AsString := sL_BelongYM;
      I_CDS.FieldByName('STBNO').AsString := sL_StbNo;
      I_CDS.FieldByName('SNO').AsString := sL_SNo;

      I_CDS.FieldByName('BILLNO').AsString := sL_BillNo;
      I_CDS.FieldByName('ITEM').AsInteger := StrToInt(sL_Item);
      I_CDS.FieldByName('PERIODID').AsInteger := StrToInt(sL_Period);
      I_CDS.FieldByName('WORKEREN1').Value := NULL;
      I_CDS.FieldByName('WORKEREN1NAME').Value := NULL;
      I_CDS.FieldByName('WORKEREN1COMM').AsInteger := StrToInt(sL_WorkerEn1Comm);

      I_CDS.FieldByName('CLCTEN').Value := NULL;
      I_CDS.FieldByName('CLCTENNAME').Value := NULL;
      I_CDS.FieldByName('CLCTENCOMM').AsInteger := StrToInt(sL_ClctEnComm);
      I_CDS.FieldByName('ClctEnOriPercent').AsFloat := fL_ClctEnOriPercent;

      I_CDS.FieldByName('INTACCID').AsString := sL_IntAccId;
      I_CDS.FieldByName('INTACCNAME').AsString := sL_IntAccName;
      I_CDS.FieldByName('INTACCCOMM').AsFloat := fL_IntAccComm;
      I_CDS.FieldByName('IntAccOriPercent').AsFloat := fL_IntAccOriPercent;

      I_CDS.FieldByName('STAFFCODE').AsString := sL_StaffCode;
      I_CDS.FieldByName('RealStartDate').AsDateTime := dL_RealStartDate;
      I_CDS.FieldByName('RealStopDate').AsDateTime := dL_RealStopDate;
      I_CDS.FieldByName('OPERATOR').AsString := frmMainMenu.sG_User;
      //I_CDS.FieldByName('UPDTIME').AsDateTime := now;
      I_CDS.FieldByName('UPDTIME').AsString := sG_CurDate;
      I_CDS.FieldByName('MEDIACODE').AsString := sL_MediaCode;
      I_CDS.FieldByName('MediaName').AsString := sL_MediaName;

      I_CDS.FieldByName('CitemCode').AsString := sL_CitemCode;
      I_CDS.FieldByName('CitemName').AsString := sL_CitemName;
      I_CDS.FieldByName('RealAmt').AsFloat := fL_RealAmt;

      if sL_PromCode <> '' then
        I_CDS.FieldByName('PromCode').AsInteger := StrToInt(sL_PromCode);
      I_CDS.FieldByName('PromName').AsString := sL_PromName;

      dL_RealDate := I_CDSSo033So034.FieldByName('RealDate').AsDateTime;
      if dL_RealDate <> 0 then
        I_CDS.FieldByName('RealDate').AsDateTime := dL_RealDate;

      dL_ShouldDate := I_CDSSo033So034.FieldByName('ShouldDate').AsDateTime;
      if dL_ShouldDate <> 0 then
        I_CDS.FieldByName('ShouldDate').AsDateTime := dL_ShouldDate;

      if sL_CustID <> '' then
        I_CDS.FieldByName('CustID').AsInteger := StrToInt(sL_CustID);

      I_CDS.FieldByName('FirstFlag').AsString := sL_FirstFlag;
      I_CDS.FieldByName('OrderNo').AsString := sL_OrderNo;

      if sL_CallSelf <> '' then
        I_CDS.FieldByName('CallSelf').AsInteger := StrToInt(sL_CallSelf);

      I_CDS.Post;
    end;
begin
    if cdsSo122.Locate('COMPCODE;BILLNO;ITEM',VarArrayOf([frmMainMenu.sG_CompCode,sI_SBillNo,sI_SItem]),[]) then
    begin
      dL_RealStartDate := StrToDate(sI_ShouldDate);
      dL_RealStopDate := cdsSo122.FieldByName('RealStopDate').AsDateTime;
      fL_IntAccOriPercent := cdsSo122.FieldByName('IntAccOriPercent').AsFloat;
      fL_ClctEnOriPercent := cdsSo122.FieldByName('ClctEnOriPercent').AsFloat;

      sL_CompCode := cdsSo122.FieldByName('CompCode').AsString;
      sL_ComputeYM := sG_ComputeYM;
      //sL_BelongYM := cdsSo122.FieldByName('BelongYM').AsString;
      sL_StbNo := cdsSo122.FieldByName('StbNo').AsString;
      sL_SNo := cdsSo122.FieldByName('Sno').AsString;

      //�ﵹ�h�O�� BillNo �� Item   sI_SBillNo,sI_SItem
      sL_BillNo := sI_SBillNo;
      sL_Item := sI_SItem;
      //sL_BillNo := cdsSo122.FieldByName('BillNo').AsString;
      //sL_Item := cdsSo122.FieldByName('Item').AsString;

      sL_Period := cdsSo122.FieldByName('PeriodId').AsString;
      sL_WorkerEn1 := 'NULL';
      sL_WorkerEn1Name := 'NULL';
      sL_WorkerEn1Comm := '0';
      sL_ClctEn := 'NULL';
      sL_ClctEnName := 'NULL';
      sL_ClctEnComm := '0';
      sL_IntAccId := cdsSo122.FieldByName('IntAccId').AsString;
      sL_IntAccName := cdsSo122.FieldByName('IntAccName').AsString;
      //fL_IntAccComm := cdsSo122.FieldByName('IntAccComm').AsFloat;
      sL_StaffCode := '3';
      sL_MediaCode := cdsSo122.FieldByName('MediaCode').AsString;
      sL_MediaName := cdsSo122.FieldByName('MediaName').AsString;

      sL_CitemCode := cdsSo122.FieldByName('CitemCode').AsString;
      sL_CitemName := cdsSo122.FieldByName('CitemName').AsString;

      fL_RealAmt := I_CDSSo033So034.FieldByName('RealAmt').AsFloat;

      sL_PromCode := cdsSo122.FieldByName('PromCode').AsString;
      sL_PromName := cdsSo122.FieldByName('PromName').AsString;

      sL_CustID := cdsSo122.FieldByName('CustID').AsString;

      sL_FirstFlag := cdsSo122.FieldByName('FirstFlag').AsString;
      sL_OrderNo := cdsSo122.FieldByName('OrderNo').AsString;
      sL_CallSelf := cdsSo122.FieldByName('CallSelf').AsString;
    end
    else
    begin
      dL_RealStartDate := StrToDate(sI_ShouldDate);
      dL_RealStopDate := dG_SFRealStopDate;
      fL_IntAccOriPercent := fG_SFIntAccOriPercent;
      //fL_ClctEnOriPercent := cdsSo122.FieldByName('ClctEnOriPercent').AsFloat;

      sL_CompCode := frmMainMenu.sG_CompCode;
      sL_ComputeYM := sG_ComputeYM;

      //�ﵹ�h�O�� BillNo �� Item   sI_SBillNo,sI_SItem
      sL_BillNo := sI_SBillNo;
      sL_Item := sI_SItem;

      sL_Period := sG_SFPeriod;
      sL_WorkerEn1 := 'NULL';
      sL_WorkerEn1Name := 'NULL';
      sL_WorkerEn1Comm := '0';
      sL_ClctEn := 'NULL';
      sL_ClctEnName := 'NULL';
      sL_ClctEnComm := '0';
      sL_IntAccId := sG_SFIntAccId;
      sL_IntAccName := sG_SFIntAccName;

      sL_StaffCode := '3';
      sL_MediaCode := sG_SFMediaCode;
      sL_MediaName := sG_SFMediaName;

      sL_CitemCode := sG_SFCitemCode;
      sL_CitemName := sG_SFCitemName;

      fL_RealAmt := I_CDSSo033So034.FieldByName('RealAmt').AsFloat;

      sL_PromCode := sG_SFPromCode;
      sL_PromName := sG_SFPromName;

      sL_CustID := sG_SFCustId;

      sL_FirstFlag := sG_SFFirstFlag;
      sL_OrderNo := sG_SFOrderNo;
      sL_CallSelf := sG_SFCallSelf;
    end;

    //SO121�������έp����  Jackal 920514
    {
    if cdsSo121.Locate('COMPCODE;EMPID',VarArrayOf([sL_CompCode,sL_IntAccId]),[]) then
    begin
      sL_MediaCodeA := cdsSo121.FieldByName('MediaCode').AsString;
      sL_MediaNameA := cdsSo121.FieldByName('MediaName').AsString;

      sL_CompTypeA := cdsSo121.FieldByName('CompType').AsString;
      sL_ClanCompCodeA := cdsSo121.FieldByName('ClanCompCode').AsString;
      sL_ClanCompNameA := cdsSo121.FieldByName('ClanCompName').AsString;
    end;
    }


    //�p�� �`���ƤѼ� �� �����p�O�Ѽ� �� ���p�����
    getComputeDays(sG_BelongYM,sI_ShouldDate,DateToStr(dL_RealStopDate),
                   nL_TotalPeriodDays,nL_FirstPeriodDays,nL_PeriodCounts,sL_RealBelongYM);

    if sI_FunctionNo = '2' then  //���p�k�ݦ~��W�L RealStopDate ,�h���������֤@���h�O
      nL_PeriodCounts := 1;

    //�p����ˤH���h�O
    for ii:=1 to nL_PeriodCounts do
    begin
      sL_NextYM := sL_RealBelongYM;
      sL_NextBelongYM := sL_RealBelongYM;
      if ii <> 1 then  //�p��U�Ӧ~��
      begin
        sL_NextYM := afterMonth(sL_NextYM,ii-2);
        nL_OtherPeriodDays := TUdateTime.DaysOfMonth(sL_NextYM + '01');

        //���p�O�̫�@��,�h��̫�@�����Ĥ@�Ѧ�RealStopDate����ڤѼ�
        if ii = nL_PeriodCounts then
        begin
          //���o�Ӥ몺�Ĥ@��
          sL_MonthFirstDay := TUdateTime.getMonthFirstDay(DateToStr(dL_RealStopDate));
          nL_OtherPeriodDays := getLastPeriodDays(StrToDate(sL_MonthFirstDay),dL_RealStopDate);
        end;


        //�U�@���k�ݦ~��
        sL_NextBelongYM := afterMonth(sL_NextBelongYM,ii-1);

        if Length(sL_NextBelongYM) = 4 then
        begin
          sL_NextBelongYM := '0' + sL_NextBelongYM;
        end;
      end;

      sL_BelongYM := sL_NextBelongYM;

      //���p���X���Ȥ��O�t��,�N�� * -1
      {
      if fI_RealAmt < 0 then
        fI_RealAmt := fI_RealAmt
      else
        fI_RealAmt := fI_RealAmt * -1;
      }

      if ii = 1 then    //�h�Ĥ@�Ӥ�
      begin
        if sL_CallSelf = '0' then  //�D�ۨӹq
          fL_IntAccComm := (fI_RealAmt * fL_IntAccOriPercent /100)*nL_FirstPeriodDays/nL_TotalPeriodDays*-1
        else
          fL_IntAccComm := nG_SelfOrderChannelAcceptComm*-1;
      end
      else
      begin
        if sL_CallSelf = '0' then  //�D�ۨӹq
          fL_IntAccComm := (fI_RealAmt * fL_IntAccOriPercent /100)*nL_OtherPeriodDays/nL_TotalPeriodDays*-1
        else
          fL_IntAccComm := 0;
      end;


      if sI_FunctionNo = '3' then  //���p��������@�Ӥ�,�h����
      begin
        if cdsSo122.Locate('COMPCODE;BELONGYM;BILLNO;ITEM',VarArrayOf([frmMainMenu.sG_CompCode,sL_BelongYM,sI_SBillNo,sI_SItem]),[]) then
          fL_IntAccComm := cdsSo122.FieldByName('INTACCCOMM').AsFloat * -1;
      end;


      if sG_BelongYM <> sL_BelongYM then
        transIntoSo122(cdsSo122ReturnMoney)
      else
        transIntoSo122(cdsSo122);



      //SO121�������έp����  Jackal 920514
      {
      if sG_BelongYM <> sL_BelongYM then
        //transIntoSo121(cdsSo121ReturnMoney)
        transIntoSo121(cdsSo121)
      else
      begin
        if cdsSo121.Locate('COMPCODE;BELONGYM;EMPID;COMPUTEYM',VarArrayOf([sL_CompCode,sL_BelongYM,sL_IntAccId,sG_ComputeYM]),[]) then
        begin
          cdsSo121.Edit;
          cdsSo121.FieldByName('Commission').AsFloat := cdsSo121.FieldByName('Commission').AsFloat + fL_IntAccComm;
          cdsSo121.post;
        end
        else
        begin
          transIntoSo121(cdsSo121)
        end;
      end;
      }

    end;
end;

function TdtmMain1.getSo125Param(sI_CompCode: String) :Boolean;
var
    sL_SQL,sL_TempExcelSavePath : String;
    L_TempList : TStringList;
    ii : Integer;
begin
    sL_SQL := 'SELECT * FROM SO125 WHERE COMPCODE=' + sI_CompCode;

    with cdsSo125 do
    begin
      Close;
      CommandText := sL_SQL;
      Open;


      //�p�G�S���,�������w�]��0
      if RecordCount = 0 then
      begin
        DisableControls;
        if State = dsInactive then
            Active := True;
        Edit;

        FieldByName('Param1').AsInteger := 0;
        FieldByName('Param2').AsInteger := 0;
        FieldByName('Param3').AsInteger := 0;
        FieldByName('Param4').AsInteger := 0;
        FieldByName('Param5').AsInteger := 0;
        FieldByName('Param6').AsInteger := 0;
        FieldByName('Param7').AsInteger := 0;
        FieldByName('Param8').AsInteger := 0;
        FieldByName('Param9').AsInteger := 0;
        FieldByName('Param10').AsInteger := 0;
        FieldByName('Param11').AsString := 'C:';
        FieldByName('Param12').AsString := '';

        //�L�����q���ѼƥN�X
        Result  := false;
      end
      else
      begin
        //���_STB�����ФH�h�֦���
        nG_BuyBoxIntrComm := FieldByName('Param1').AsInteger;

        //����STB�����ФH�h�֦���
        nG_RentBoxIntrComm := FieldByName('Param2').AsInteger;

        //�w��STB���u�{�H���h�֦���
        nG_SetBoxWorkerComm := FieldByName('Param3').AsInteger;

        //���_STB�����z�H(�ȪA)�h�֦���
        nG_BuyBoxAcceptComm := FieldByName('Param4').AsInteger;

        //����STB�����z�H(�ȪA)�h�֦���
        nG_RentBoxAcceptComm := FieldByName('Param5').AsInteger;

        //�Ȥ�ۨӹq�q���W�D�����s�H(�ȪA)�h�֦���
        nG_SelfOrderChannelAcceptComm := FieldByName('Param6').AsInteger;

        //STB�ݯ����h�֤Ѥ~������
        nG_RentBoxUseDays := FieldByName('Param7').AsInteger;

        //STB�R�_���ϥκ��h�֤Ѥ~�����s�H����
        nG_BuyBoxUseDays := FieldByName('Param8').AsInteger;

        //��������X��
        nG_LimitPeriod := FieldByName('Param9').AsInteger;

        //�@�����I������Ū��??��e�����h�O���
        nG_BackMonth := FieldByName('Param10').AsInteger;


        //BOX�ݱư����~�Ȭ���
        sG_PromCodeAndName := FieldByName('Param12').AsString;

        L_TempList := TUStr.ParseStrings(sG_PromCodeAndName,';');

        G_PromCodeStrList.Clear;
        G_PromNameStrList.Clear;


        if sG_PromCodeAndName <> '' then
        begin
          G_PromCodeStrList := TUStr.ParseStrings(L_TempList.Strings[0],',');
          G_PromNameStrList := TUStr.ParseStrings(L_TempList.Strings[1],',');
        end;

        L_TempList.Free;

        //Excel�����x�s���|(�w�]�bC:\)
        sL_TempExcelSavePath := FieldByName('Param11').AsString;

        if sL_TempExcelSavePath = '' then
          sG_ExcelSavePath := 'C:'
        else
        begin
          if (sL_TempExcelSavePath = 'C:\') OR (sL_TempExcelSavePath = 'D:\') then
            sG_ExcelSavePath := TUstr.replaceStr(sL_TempExcelSavePath,'\','')
          else
            sG_ExcelSavePath := sL_TempExcelSavePath;
        end;

        Edit;
        FieldByName('Param11').AsString := sG_ExcelSavePath;
        Post;



        {
        for ii:=0 to G_PromCodeStrList.Count-1 do
        begin
          showmessage('getSo125Param== ' + G_PromCodeStrList.Strings[ii] + '-' + G_PromNameStrList.Strings[ii]);
        end;
        }

        Result  := true;
      end;
    end;


end;

function TdtmMain1.getLockYM : String;
var
    sL_SQL : String;
begin
    sL_SQL := 'SELECT * FROM SO124 WHERE COMPCODE=' + frmMainMenu.sG_CompCode;

    with qrySO124 do
    begin
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;

      Result := FieldByName('LOCKYM').AsString;
      Result := IntToStr( StrToInt( Copy( Result, 1, 3 ) ) + 1911 ) +
        Copy( Result, 4, 2 );
    end;
end;

procedure TdtmMain1.updateLockYM(sI_NewLockYM : String);
var
  sL_SQL, aROC: String;
begin
  aROC := IntToStr( StrToInt( Copy( sI_NewLockYM, 1, 4 ) ) - 1911 );
  if Length( aROC ) < 3 then aROC := ( '0' + aROC );
  aROC := ( aROC + Copy( sI_NewLockYM, 5, 2 ) );
  sL_SQL := 'UPDATE SO124 SET LOCKYM=''' + aROC +
            ''',OPERATOR=''' + frmMainMenu.sG_User +
            ''',UPDTIME=''' + DateTimeToStr(now) +
            ''' WHERE COMPCODE=' + frmMainMenu.sG_CompCode;

//showmessage(sL_SQL);
    with qrySO124 do
    begin
      SQL.Clear;
      SQL.Add(sL_SQL);
      ExecSQL;
    end;
end;

procedure TdtmMain1.InsertLockYM(sI_NewLockYM: String);
var
    sL_SQL, aROC: String;
begin
  aROC := IntToStr( StrToInt( Copy( sI_NewLockYM, 1, 4 ) ) - 1911 );
  if Length( aROC ) < 3 then aROC := ( '0' + aROC );
  aROC := ( aROC + Copy( sI_NewLockYM, 5, 2 ) );
    sL_SQL := 'INSERT INTO SO124 VALUES(' + frmMainMenu.sG_CompCode +
              ',''' + aROC + ''',''' + frmMainMenu.sG_User +
              ''',''' + DateTimeToStr(now) + ''')';

//showmessage(sL_SQL);
    with qrySO124 do
    begin
      SQL.Clear;
      SQL.Add(sL_SQL);
      ExecSQL;
    end;

end;

procedure TdtmMain1.getSO121Data(sI_BelongYM, sI_Compute,sI_ManList: String);
var
    sL_SQL : String;
begin
    //�Y�̳������s��,�h�u�q�ۤv���q�����(CompType=1),���q�������(CompType=2)
    sL_SQL := 'SELECT * FROM SO121 WHERE COMPCODE=' + frmMainMenu.sG_CompCode +
              ' AND BELONGYM=''' + sI_BelongYM + '''';

    if sI_ManList <> '' then
      sL_SQL := sL_SQL + ' AND EMPID IN(' + sI_ManList + ') ';

    if sI_Compute = '0' then        //�̤��q���s�խp��
    begin
      sL_SQL := sL_SQL + ' ORDER BY CLANCOMPCODE,EMPID';
    end
    else if sI_Compute = '1' then   //�̳������s�խp��
    begin
      sL_SQL := sL_SQL + ' AND COMPTYPE=''1''  ORDER BY MEDIACODE,EMPID';
    end;

    with cdsSo121 do
    begin
      Close;
      CommandText := sL_SQL;
      Open;
    end;
end;

procedure TdtmMain1.getCM003Data;
var
    sL_SQL : String;
begin
    sL_SQL := 'SELECT * FROM CM003 WHERE COMPCODE=' + frmMainMenu.sG_CompCOde +
              ' AND STOPFLAG=0 ORDER BY EMPNO';

    with qryCM003 do
    begin
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;
    end;
end;

procedure TdtmMain1.getSO013Data;
var
    sL_SQL : String;
begin
    sL_SQL := 'SELECT * FROM SO013 WHERE COMPCODE=' + frmMainMenu.sG_CompCOde +
              ' ORDER BY INTROID';

    with qrySO013 do
    begin
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;
    end;
end;

procedure TdtmMain1.getOtherCompCodeData;
var
    sL_SQL : String;
begin
    sL_SQL := 'SELECT DISTINCT a.IntAccId,a.IntAccName FROM SO122 a,CD009 b ' +
              'WHERE a.COMPCODE=' + frmMainMenu.sG_CompCOde + ' AND a.MediaCode=b.CodeNo ' +
              'AND b.RefNo=5 AND b.SERVICETYPE=''' + frmMainMenu.sG_ServiceType + ''' ORDER BY a.IntAccId';

    with qryOtherCompCode do
    begin
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;
    end;

end;

function TdtmMain1.getRptChannelDetailData(sI_PayType,sI_BelongYM,sI_Compute,sI_EmpNoSQL,sI_SalesCodeSQL,sI_OtherCompCodeSQL,sI_OrderList,sI_CitemCodeSQL : String) : Boolean;
var
    sL_SQLTitle,sL_SQL,sL_Where,sL_Table : String;
begin
    if sI_PayType = BATCH_PAY then
      sL_Table := 'SO122'
    else if sI_PayType = ONCE_PAY then
      sL_Table := 'SO134';

    sL_SQLTitle := 'SELECT * FROM ' + sL_Table;
    sL_Where := ' WHERE COMPCODE=' + sI_Compute + ' AND BELONGYM=''' +
                sI_BelongYM + ''' AND BILLNO IS NOT NULL ';

    if sI_CitemCodeSQL <> '' then
    begin
        sL_Where := sL_Where + ' AND CitemCode in (' + sI_CitemCodeSQL + ') ';
    end;


    if ((sI_EmpNoSQL<>'') OR (sI_SalesCodeSQL<>'')
                          OR (sI_OtherCompCodeSQL<>'')) then
    begin
      if sI_EmpNoSQL <> '' then
      begin
        sL_Where := sL_Where + ' AND ((ClctEn IN(' + sI_EmpNoSQL +
                    ') OR IntAccId IN(' + sI_EmpNoSQL + '))';
      end
      else
      begin
        sL_Where := sL_Where + ' AND (' +
                    'ClctEn IN ('''') OR IntAccId IN(''''))';
      end;

      if sI_SalesCodeSQL <> '' then
        sL_Where := sL_Where + ' OR IntAccId IN (' + sI_SalesCodeSQL + ')'
      else
        sL_Where := sL_Where + ' OR IntAccId IN ('''')';


      if sI_OtherCompCodeSQL <> '' then
        sL_Where := sL_Where + ' OR IntAccId IN (' + sI_OtherCompCodeSQL + '))'
      else
        sL_Where := sL_Where + ' OR IntAccId IN (''''))';

    end;
    sL_SQL := sL_SQLTitle + sL_Where + ' ORDER BY ' + sI_OrderList;

    //SO122�PSO134�W�D�������@��,�ҥH�@��CDS
    with cdsSo134 do
    begin
      Close;
      CommandText := sL_SQL;
      Open;

      if RecordCount<>0 then
        Result := true
      else
        Result := false;
    end;


end;

function TdtmMain1.getLastPeriodDays(dI_MonthFirstDay,
  dI_RealStopDate: TDate): Integer;
var
    nL_DiffDays : Integer;
begin
    //�p����������Ѽ�
    nL_DiffDays := Round(dI_RealStopDate - dI_MonthFirstDay + 1);
    Result := nL_DiffDays;
end;

procedure TdtmMain1.copyCdsSo121ToQry;
var
    ii : Integer;
begin
    if ADOQuery1.State=dsInactive then
      ADOQuery1.Open
    else
    begin
      ADOQuery1.Close;
      ADOQuery1.Open;
    end;

    with cdsSo121 do
    begin
      First;
      while not Eof do
      begin
        ADOQuery1.Append;
        for ii:=0 to FieldCount -1 do
        begin
          ADOQuery1.Fields[ii].Value := Fields[ii].Value
        end;
        ADOQuery1.Post;
        Next;
      end;
    end;
end;


procedure TdtmMain1.getCodeCD019;
var
    sL_SQL : String;
begin
    //�d�� CD019 ���O���إN�X�ɩҦ����
    sL_SQL := 'SELECT CODENO,DESCRIPTION FROM CD019 WHERE StopFlag=0' +
              ' AND SERVICETYPE=''' + frmMainMenu.sG_ServiceType + ''' AND REFNO=2';
    with qryCodeCD019 do
    begin
      Close;
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;

    end;
end;

procedure TdtmMain1.getCD019Data;
Var
  sL_SQL:String;
begin
  sL_SQL:='SELECT CODENO,DESCRIPTION FROM CD019 WHERE ' +
          ' SERVICETYPE=''' + frmMainMenu.sG_ServiceType + ''' AND REFNO=2';

  with qryCD019 do
  begin
    Close;
    SQL.Clear;
    SQL.Add(sL_SQL);
    Open;
  end;
end;

procedure TdtmMain1.getSO122Data1(sI_Flag,sI_BelongYM,sI_CompCode,sI_CODENO:String;bI_IncludeWorkerComm : Boolean);
Var
    sL_SQL,sL_Where,sL_SQLTitle,sL_FirstFlag : String;
    fL_WorkerEnComm,fL_ClctComm,fL_IntAccComm,fL_RealAmt : Double;
    fL_TotalWorkerEnComm,fL_TotalClctComm,fL_TotalIntAccComm,fL_TotalRealAmt : Double;
    aValue: String;
begin

   sL_SQLTitle:='select COMPCODE,COMPUTEYM,BELONGYM,STBNO,SNO,BILLNO,ITEM,'+
                'PERIODID,WORKEREN1,WORKEREN1NAME,WORKEREN1COMM,TO_CHAR(WORKEREN1COMM) StrWORKEREN1COMM,CLCTEN,CLCTENNAME,'+
                'CLCTENCOMM,TO_CHAR(CLCTENCOMM) StrCLCTENCOMM,CLCTENORIPERCENT,INTACCID,INTACCNAME,INTACCCOMM,TO_CHAR(INTACCCOMM)StrINTACCCOMM,INTACCORIPERCENT,'+
                'STAFFCODE,REALSTARTDATE,REALSTOPDATE,OPERATOR,UPDTIME,BUYORRENT,CITEMCODE,'+
                'CITEMNAME,MEDIACODE,MEDIANAME,REALAMT,TO_CHAR(REALAMT) StrREALAMT,PROMCODE,PROMNAME,REALDATE,SHOULDDATE,CUSTID,FIRSTFLAG '
                +' from SO122';


    if sI_Flag = CHANNEL_DATA then
    begin

        sL_Where:=' where citemcode in ( '+ sI_CODENO +') AND BelongYM='''
            + sI_BelongYM+''' AND STBNO IS NULL'+' AND COMPCODE='+ sI_CompCode;

    end
    else if sI_Flag = BOX_DATA then
    begin
      sL_Where:=' where BelongYM='''+ sI_BelongYM +''' AND COMPCODE='''+sI_CompCode+
          ''' AND STBNO IS NOT NULL';

    end;

    sL_SQL:=sL_SQLTitle+sL_Where;

    with cdsSO122Data do
    begin
      Close;
      CommandText := sL_SQL;
      Open;
    end;

{
    dtmMain1.cdsSO122Data.Filtered := false;
    dtmMain1.cdsSO122Data.Filter := 'INTACCCOMM<>0';
    dtmMain1.cdsSO122Data.Filtered := true;
}
    //���t��
    transSo122Notation(cdsSO122Data);

    fL_TotalWorkerEnComm := 0;
    fL_TotalClctComm := 0;
    fL_TotalIntAccComm := 0;


    if sI_Flag = CHANNEL_DATA then
    begin

      //�W�D����� Excel
      cdsSO122Data.First;
      while not cdsSO122Data.Eof do
      begin
        cdsSO122Channel.Append;

        cdsSO122Channel.FieldByName('COMPCODE').AsString:= cdsSO122Data.FieldByName('COMPCODE').AsString;

        aValue := cdsSO122Data.FieldByName('COMPUTEYM').AsString;
//        if ( aValue <> EmptyStr ) then
//          aValue := IntToStr( StrToInt( Copy( aValue, 1, 3 ) ) + 1911 ) + Copy( aValue, 4, 2 );

        cdsSO122Channel.FieldByName('COMPUTEYM').AsString:= aValue;

        aValue := cdsSO122Data.FieldByName('BELONGYM').AsString;
//        if ( aValue <> EmptyStr ) then
//          aValue := IntToStr( StrToInt( Copy( aValue, 1, 3 ) ) + 1911 ) + Copy( aValue, 4, 2 );

        cdsSO122Channel.FieldByName('BELONGYM').AsString:= aValue;


        cdsSO122Channel.FieldByName('BILLNO').AsString := cdsSO122Data.FieldByName('BILLNO').AsString;
        cdsSO122Channel.FieldByName('ITEM').AsString := cdsSO122Data.FieldByName('ITEM').AsString;
        cdsSO122Channel.FieldByName('PERIODID').AsString := cdsSO122Data.FieldByName('PERIODID').AsString;
        cdsSO122Channel.FieldByName('CLCTEN').AsString := cdsSO122Data.FieldByName('CLCTEN').AsString;
        cdsSO122Channel.FieldByName('CLCTENNAME').AsString := cdsSO122Data.FieldByName('CLCTENNAME').AsString;

        fL_ClctComm := cdsSO122Data.FieldByName('CLCTENCOMM').AsFloat;
        fL_TotalClctComm := fL_TotalClctComm + fL_ClctComm;
        cdsSO122Channel.FieldByName('CLCTENCOMM').AsFloat := fL_ClctComm;

        cdsSO122Channel.FieldByName('CLCTENORIPERCENT').AsFloat := cdsSO122Data.FieldByName('CLCTENORIPERCENT').AsFloat;
        cdsSO122Channel.FieldByName('STAFFCODE').AsString := cdsSO122Data.FieldByName('STAFFCODE').AsString;
        cdsSO122Channel.FieldByName('INTACCID').AsString := cdsSO122Data.FieldByName('INTACCID').AsString;
        cdsSO122Channel.FieldByName('INTACCNAME').AsString := cdsSO122Data.FieldByName('INTACCNAME').AsString;

        fL_IntAccComm := cdsSO122Data.FieldByName('INTACCCOMM').AsFloat;
        fL_TotalIntAccComm := fL_TotalIntAccComm + fL_IntAccComm;
        cdsSO122Channel.FieldByName('INTACCCOMM').AsFloat := fL_IntAccComm;

        cdsSO122Channel.FieldByName('INTACCORIPERCENT').AsFloat := cdsSO122Data.FieldByName('INTACCORIPERCENT').AsFloat;
        cdsSO122Channel.FieldByName('REALSTARTDATE').AsString := cdsSO122Data.FieldByName('REALSTARTDATE').AsString;
        cdsSO122Channel.FieldByName('REALSTOPDATE').AsString := cdsSO122Data.FieldByName('REALSTOPDATE').AsString;

        try
          aValue := cdsSO122Data.FieldByName('REALSTARTDATE').AsString;
          if ( aValue <> EmptyStr ) then
            aValue := TUstr.ToROCYMD( aValue );
          cdsSO122Channel.FieldByName('REALSTARTDATE2').AsString := aValue;
        except
          cdsSO122Channel.FieldByName('REALSTARTDATE2').AsString := cdsSO122Data.FieldByName('REALSTARTDATE').AsString;
        end;

        try
          aValue := cdsSO122Data.FieldByName('REALSTOPDATE').AsString;
          if ( aValue <> EmptyStr ) then
            aValue := TUstr.ToROCYMD( aValue );
          cdsSO122Channel.FieldByName('REALSTOPDATE2').AsString := aValue;
        except
          cdsSO122Channel.FieldByName('REALSTOPDATE2').AsString := cdsSO122Data.FieldByName('REALSTOPDATE').AsString;
        end;


        cdsSO122Channel.FieldByName('CITEMCODE').AsString := cdsSO122Data.FieldByName('CITEMCODE').AsString;
        cdsSO122Channel.FieldByName('CITEMNAME').AsString := cdsSO122Data.FieldByName('CITEMNAME').AsString;
        cdsSO122Channel.FieldByName('MEDIACODE').AsString := cdsSO122Data.FieldByName('MEDIACODE').AsString;
        cdsSO122Channel.FieldByName('MEDIANAME').AsString := cdsSO122Data.FieldByName('MEDIANAME').AsString;

        fL_RealAmt := cdsSO122Data.FieldByName('REALAMT').AsFloat;
        fL_TotalRealAmt := fL_TotalRealAmt + fL_RealAmt;
        cdsSO122Channel.FieldByName('REALAMT').AsFloat := fL_RealAmt;

        cdsSO122Channel.FieldByName('OPERATOR').AsString := cdsSO122Data.FieldByName('OPERATOR').AsString;
        cdsSO122Channel.FieldByName('UPDTIME').AsString := cdsSO122Data.FieldByName('UPDTIME').AsString;
        cdsSO122Channel.FieldByName('REALDATE').AsString := cdsSO122Data.FieldByName('REALDATE').AsString;
        cdsSO122Channel.FieldByName('SHOULDDATE').AsString := cdsSO122Data.FieldByName('SHOULDDATE').AsString;

        try
          aValue := cdsSO122Data.FieldByName('REALDATE').AsString;
          if ( aValue <> EmptyStr ) then
            aValue := TUstr.ToROCYMD( aValue );
          cdsSO122Channel.FieldByName('REALDATE2').AsString := aValue;
        except
          cdsSO122Channel.FieldByName('REALDATE2').AsString := cdsSO122Data.FieldByName('REALSTARTDATE').AsString;
        end;

        try
          aValue := cdsSO122Data.FieldByName('SHOULDDATE').AsString;
          if ( aValue <> EmptyStr ) then
            aValue := TUstr.ToROCYMD( aValue );
          cdsSO122Channel.FieldByName('SHOULDDATE2').AsString := aValue;
        except
          cdsSO122Channel.FieldByName('SHOULDDATE2').AsString := cdsSO122Data.FieldByName('REALSTOPDATE').AsString;
        end;


        cdsSO122Channel.FieldByName('PROMCODE').AsInteger := cdsSO122Data.FieldByName('PROMCODE').AsInteger;
        cdsSO122Channel.FieldByName('PROMNAME').AsString := cdsSO122Data.FieldByName('PROMNAME').AsString;
        cdsSO122Channel.FieldByName('CustID').AsString := cdsSO122Data.FieldByName('CustID').AsString;

        sL_FirstFlag := cdsSO122Data.FieldByName('FirstFlag').AsString;
        if sL_FirstFlag ='1' then
          cdsSO122Channel.FieldByName('FirstFlagName').AsString := '����'
        else
          cdsSO122Channel.FieldByName('FirstFlagName').AsString := '��';

        cdsSO122Channel.Post;

        cdsSO122Data.Next;
      end;

      cdsSO122Channel.Append;
      cdsSO122Channel.FieldByName('COMPUTEYM').AsString := '�`�p';
      cdsSO122Channel.FieldByName('CLCTENCOMM').AsFloat := fL_TotalClctComm;
      cdsSO122Channel.FieldByName('INTACCCOMM').AsFloat := fL_TotalIntAccComm;
      cdsSO122Channel.FieldByName('REALAMT').AsFloat := fL_TotalRealAmt;

      //�t�X�Ƨ�,�[�`�n�\�̫᭱
      cdsSO122Channel.FieldByName('CustID').AsInteger := 99999999;
      cdsSO122Channel.FieldByName('CLCTEN').AsString := '��';
      cdsSO122Channel.FieldByName('INTACCID').AsString := '��';

      cdsSO122Channel.Post;

    end
    else if sI_Flag = BOX_DATA then
    begin
      cdsSO122Data.First;
      while not cdsSO122Data.Eof do
      begin

        cdsSO122BOX.Append;

        cdsSO122BOX.FieldByName('COMPCODE').AsString:= cdsSO122Data.FieldByName('COMPCODE').AsString;

        aValue := cdsSO122Data.FieldByName('COMPUTEYM').AsString;
//        if ( aValue <> EmptyStr ) then
//          aValue := IntToStr( StrToInt( Copy( aValue, 1, 3 ) ) + 1911 ) + Copy( aValue, 4, 2 );
        cdsSO122BOX.FieldByName('COMPUTEYM').AsString:= aValue;


        aValue := cdsSO122Data.FieldByName('BELONGYM').AsString;
//        if ( aValue <> EmptyStr ) then
//          aValue := IntToStr( StrToInt( Copy( aValue, 1, 3 ) ) + 1911 ) + Copy( aValue, 4, 2 );
        cdsSO122BOX.FieldByName('BELONGYM').AsString:= aValue;

        
        cdsSO122BOX.FieldByName('STBNO').AsString:= cdsSO122Data.FieldByName('STBNO').AsString;
        cdsSO122BOX.FieldByName('SNO').AsString := cdsSO122Data.FieldByName('SNO').AsString;
        cdsSO122BOX.FieldByName('WORKEREN1').AsString := cdsSO122Data.FieldByName('WORKEREN1').AsString;
        cdsSO122BOX.FieldByName('WORKEREN1NAME').AsString := cdsSO122Data.FieldByName('WORKEREN1NAME').AsString;

        //���Ŀ�]�tBOX�u�{�H������
        if bI_IncludeWorkerComm then
          fL_WorkerEnComm := cdsSO122Data.FieldByName('WORKEREN1COMM').AsFloat
        else
          fL_WorkerEnComm := 0;

        fL_TotalWorkerEnComm := fL_TotalWorkerEnComm + fL_WorkerEnComm;
        cdsSO122BOX.FieldByName('WORKEREN1COMM').AsFloat := fL_WorkerEnComm;

        cdsSO122BOX.FieldByName('STAFFCODE').AsString := cdsSO122Data.FieldByName('STAFFCODE').AsString;
        cdsSO122BOX.FieldByName('INTACCID').AsString := cdsSO122Data.FieldByName('INTACCID').AsString;
        cdsSO122BOX.FieldByName('INTACCNAME').AsString := cdsSO122Data.FieldByName('INTACCNAME').AsString;

        fL_IntAccComm := cdsSO122Data.FieldByName('INTACCCOMM').AsFloat;
        fL_TotalIntAccComm := fL_TotalIntAccComm + fL_IntAccComm;
        cdsSO122BOX.FieldByName('INTACCCOMM').AsFloat := fL_IntAccComm;

        cdsSO122BOX.FieldByName('REALSTARTDATE').AsString := cdsSO122Data.FieldByName('REALSTARTDATE').AsString;
        cdsSO122BOX.FieldByName('REALSTOPDATE').AsString := cdsSO122Data.FieldByName('REALSTOPDATE').AsString;


        try
          aValue := cdsSO122Data.FieldByName('REALSTARTDATE').AsString;
          if ( aValue <> EmptyStr ) then
            aValue := TUstr.ToROCYMD( aValue );
          cdsSO122BOX.FieldByName('REALSTARTDATE2').AsString := aValue;
        except
          cdsSO122BOX.FieldByName('REALSTARTDATE2').AsString := cdsSO122Data.FieldByName('REALSTARTDATE').AsString;
        end;

        try
          aValue := cdsSO122Data.FieldByName('REALSTOPDATE').AsString;
          if ( aValue <> EmptyStr ) then
            aValue := TUstr.ToROCYMD( aValue );
          cdsSO122BOX.FieldByName('REALSTOPDATE2').AsString := aValue;
        except
          cdsSO122BOX.FieldByName('REALSTOPDATE2').AsString := cdsSO122Data.FieldByName('REALSTOPDATE').AsString;
        end;


        cdsSO122BOX.FieldByName('BUYORRENT').AsString := cdsSO122Data.FieldByName('BUYORRENT').AsString;
        cdsSO122BOX.FieldByName('MEDIACODE').AsString := cdsSO122Data.FieldByName('MEDIACODE').AsString;
        cdsSO122BOX.FieldByName('MEDIANAME').AsString := cdsSO122Data.FieldByName('MEDIANAME').AsString;
        cdsSO122BOX.FieldByName('OPERATOR').AsString := cdsSO122Data.FieldByName('OPERATOR').AsString;
        cdsSO122BOX.FieldByName('UPDTIME').AsString := cdsSO122Data.FieldByName('UPDTIME').AsString;
        cdsSO122BOX.FieldByName('REALDATE').AsString := cdsSO122Data.FieldByName('RealDate').AsString;
        cdsSO122BOX.FieldByName('SHOULDDATE').AsString := cdsSO122Data.FieldByName('ShouldDate').AsString;


        try
          aValue := cdsSO122Data.FieldByName('REALDATE').AsString;
          if ( aValue <> EmptyStr ) then
            aValue := TUstr.ToROCYMD( aValue );
          cdsSO122BOX.FieldByName('REALDATE2').AsString := aValue;
        except
          cdsSO122BOX.FieldByName('REALDATE2').AsString := cdsSO122Data.FieldByName('REALSTARTDATE').AsString;
        end;

        try
          aValue := cdsSO122Data.FieldByName('SHOULDDATE').AsString;
          if ( aValue <> EmptyStr ) then
            aValue := TUstr.ToROCYMD( aValue );
          cdsSO122BOX.FieldByName('SHOULDDATE2').AsString := aValue;
        except
          cdsSO122BOX.FieldByName('SHOULDDATE2').AsString := cdsSO122Data.FieldByName('REALSTOPDATE').AsString;
        end;






        //cdsSO122BOX.FieldByName('REALAMT').AsString := cdsSO122Data.FieldByName('REALAMT').AsString;
        cdsSO122BOX.FieldByName('PROMCODE').AsString := cdsSO122Data.FieldByName('PROMCODE').AsString;
        cdsSO122BOX.FieldByName('PROMNAME').AsString := cdsSO122Data.FieldByName('PROMNAME').AsString;
        cdsSO122BOX.FieldByName('CustID').AsString := cdsSO122Data.FieldByName('CustID').AsString;

        cdsSO122BOX.Post;

        cdsSO122Data.Next;
      end;

        cdsSO122BOX.Append;
        cdsSO122BOX.FieldByName('COMPUTEYM').AsString := '�`�p';
        cdsSO122BOX.FieldByName('WORKEREN1COMM').AsFloat := fL_TotalWorkerEnComm;
        cdsSO122BOX.FieldByName('INTACCCOMM').AsFloat := fL_TotalIntAccComm;

        //�t�X�Ƨ�,�[�`�n�\�̫᭱
        cdsSO122BOX.FieldByName('CustID').AsInteger := 99999999;
        cdsSO122BOX.FieldByName('WORKEREN1').AsString := '��';
        cdsSO122BOX.FieldByName('INTACCID').AsString := '��';

        cdsSO122BOX.Post;
    end;
end;

function TdtmMain1.TransToExcel(sI_Flag,sI_BelongYM:String): String;
Var
  sL_QueryString,sL_FileName,sL_FileCurrDateTime,sL_CurrDateTime : String;
  L_StrList : TStringList;
begin
    sL_FileCurrDateTime := DateTimeToStr(now);
    sL_CurrDateTime := ReplaceStr(Trim(ReplaceStr(sL_FileCurrDateTime,'/')),':');

    //�W�D����� Excel
    if sI_Flag='1' then
    begin
      cdsSO122Data.First;
      while not cdsSO122Data.Eof do
      begin
        cdsSO122Channel.Append;

        cdsSO122Channel.FieldByName('COMPCODE').AsString:= cdsSO122Data.FieldByName('COMPCODE').AsString;
        cdsSO122Channel.FieldByName('COMPUTEYM').AsString:= cdsSO122Data.FieldByName('COMPUTEYM').AsString;
        cdsSO122Channel.FieldByName('BELONGYM').AsString:= cdsSO122Data.FieldByName('BELONGYM').AsString;
        cdsSO122Channel.FieldByName('BILLNO').AsString := cdsSO122Data.FieldByName('BILLNO').AsString;
        cdsSO122Channel.FieldByName('ITEM').AsString := cdsSO122Data.FieldByName('ITEM').AsString;
        cdsSO122Channel.FieldByName('PERIODID').AsString := cdsSO122Data.FieldByName('PERIODID').AsString;
        cdsSO122Channel.FieldByName('CLCTEN').AsString := cdsSO122Data.FieldByName('CLCTEN').AsString;
        cdsSO122Channel.FieldByName('CLCTENNAME').AsString := cdsSO122Data.FieldByName('CLCTENNAME').AsString;
        cdsSO122Channel.FieldByName('CLCTENCOMM').AsFloat := cdsSO122Data.FieldByName('CLCTENCOMM').AsFloat;
        cdsSO122Channel.FieldByName('CLCTENORIPERCENT').AsFloat := cdsSO122Data.FieldByName('CLCTENORIPERCENT').AsFloat;
        cdsSO122Channel.FieldByName('STAFFCODE').AsString := cdsSO122Data.FieldByName('STAFFCODE').AsString;
        cdsSO122Channel.FieldByName('INTACCID').AsString := cdsSO122Data.FieldByName('INTACCID').AsString;
        cdsSO122Channel.FieldByName('INTACCNAME').AsString := cdsSO122Data.FieldByName('INTACCNAME').AsString;
        cdsSO122Channel.FieldByName('INTACCCOMM').AsFloat := cdsSO122Data.FieldByName('INTACCCOMM').AsFloat;
        cdsSO122Channel.FieldByName('INTACCORIPERCENT').AsFloat := cdsSO122Data.FieldByName('INTACCORIPERCENT').AsFloat;
        cdsSO122Channel.FieldByName('REALSTARTDATE').AsString := cdsSO122Data.FieldByName('REALSTARTDATE').AsString;
        cdsSO122Channel.FieldByName('REALSTOPDATE').AsString := cdsSO122Data.FieldByName('REALSTOPDATE').AsString;
        cdsSO122Channel.FieldByName('CITEMCODE').AsString := cdsSO122Data.FieldByName('CITEMCODE').AsString;
        cdsSO122Channel.FieldByName('CITEMNAME').AsString := cdsSO122Data.FieldByName('CITEMNAME').AsString;
        cdsSO122Channel.FieldByName('MEDIACODE').AsString := cdsSO122Data.FieldByName('MEDIACODE').AsString;
        cdsSO122Channel.FieldByName('MEDIANAME').AsString := cdsSO122Data.FieldByName('MEDIANAME').AsString;
        cdsSO122Channel.FieldByName('REALAMT').AsFloat := cdsSO122Data.FieldByName('REALAMT').AsFloat;
        cdsSO122Channel.FieldByName('OPERATOR').AsString := cdsSO122Data.FieldByName('OPERATOR').AsString;
        cdsSO122Channel.FieldByName('UPDTIME').AsString := cdsSO122Data.FieldByName('UPDTIME').AsString;
        cdsSO122Channel.FieldByName('REALDATE').AsString := cdsSO122Data.FieldByName('REALDATE').AsString;
        cdsSO122Channel.FieldByName('SHOULDDATE').AsString := cdsSO122Data.FieldByName('SHOULDDATE').AsString;
        cdsSO122Channel.FieldByName('PROMCODE').AsInteger := cdsSO122Data.FieldByName('PROMCODE').AsInteger;
        cdsSO122Channel.FieldByName('PROMNAME').AsString := cdsSO122Data.FieldByName('PROMNAME').AsString;
        cdsSO122Channel.FieldByName('CustID').AsString := cdsSO122Data.FieldByName('CustID').AsString;

        cdsSO122Channel.Post;

        cdsSO122Data.Next;
      end;


      // '�p��ɶ�: 2003/05/02 10:54:27@�p��H��: Jackal@���q�W��: �}��@�d�ߦ~����: �ꦬ���@�d�ߦ~��: 092/01@�d�ߦ��O����: �򥻥I�O�W,�ൣ�^�y�о��W�D,���k��T�о��W�D,
       // �����о��W�D,�����Ҳy�о��W�D,���@�о��W�D,�q�J���о��W�D'
      sL_QueryString:='�p��ɶ�:' + sL_FileCurrDateTime + '@' + '����H��:' + frmMainMenu.sG_User+'@' +
                      '���q�W��:' + frmMainMenu.sG_CompName + '@' + '�k�ݦ~��:' + sI_BelongYM + '@' +
                      '�d�ߦ��O����:' + frmSO8B20_1.sG_ChargeNameSQL;
//showmessage(sL_QueryString);

      sL_FileName:='C:\�W�D����' + sL_CurrDateTime + '.XLS';

      cdsSO122Channel.DisableControls;
      DataSetToXLS(cdsSO122Channel,sL_FileName,sL_QueryString,'');
      cdsSO122Channel.EnableControls;

      if cdsSO122Channel.RecordCount > 0 then
        MessageDlg('�p�⧹��,�W�D�����s�� ' + sL_FileName,mtInformation,[mbOK],0)
      else
        MessageDlg('�S���W�D�������',mtInformation,[mbOK],0);
    end;



    //BOX ����� Excel
    if sI_Flag='2' then
    begin
      cdsSO122Data.First;
      while not cdsSO122Data.Eof do
      begin

        cdsSO122BOX.Append;

        cdsSO122BOX.FieldByName('COMPCODE').AsString:= cdsSO122Data.FieldByName('COMPCODE').AsString;
        cdsSO122BOX.FieldByName('COMPUTEYM').AsString:= cdsSO122Data.FieldByName('COMPUTEYM').AsString;
        cdsSO122BOX.FieldByName('BELONGYM').AsString:= cdsSO122Data.FieldByName('BELONGYM').AsString;
        cdsSO122BOX.FieldByName('STBNO').AsString:= cdsSO122Data.FieldByName('STBNO').AsString;
        cdsSO122BOX.FieldByName('SNO').AsString := cdsSO122Data.FieldByName('SNO').AsString;
        cdsSO122BOX.FieldByName('WORKEREN1').AsString := cdsSO122Data.FieldByName('WORKEREN1').AsString;
        cdsSO122BOX.FieldByName('WORKEREN1NAME').AsString := cdsSO122Data.FieldByName('WORKEREN1NAME').AsString;
        cdsSO122BOX.FieldByName('WORKEREN1COMM').AsFloat := cdsSO122Data.FieldByName('WORKEREN1COMM').AsFloat;
        cdsSO122BOX.FieldByName('STAFFCODE').AsString := cdsSO122Data.FieldByName('STAFFCODE').AsString;
        cdsSO122BOX.FieldByName('INTACCID').AsString := cdsSO122Data.FieldByName('INTACCID').AsString;
        cdsSO122BOX.FieldByName('INTACCNAME').AsString := cdsSO122Data.FieldByName('INTACCNAME').AsString;
        cdsSO122BOX.FieldByName('INTACCCOMM').AsFloat := cdsSO122Data.FieldByName('INTACCCOMM').AsFloat;
        cdsSO122BOX.FieldByName('REALSTARTDATE').AsString := cdsSO122Data.FieldByName('REALSTARTDATE').AsString;
        cdsSO122BOX.FieldByName('REALSTOPDATE').AsString := cdsSO122Data.FieldByName('REALSTOPDATE').AsString;
        cdsSO122BOX.FieldByName('BUYORRENT').AsString := cdsSO122Data.FieldByName('BUYORRENT').AsString;
        cdsSO122BOX.FieldByName('MEDIACODE').AsString := cdsSO122Data.FieldByName('MEDIACODE').AsString;
        cdsSO122BOX.FieldByName('MEDIANAME').AsString := cdsSO122Data.FieldByName('MEDIANAME').AsString;
        cdsSO122BOX.FieldByName('OPERATOR').AsString := cdsSO122Data.FieldByName('OPERATOR').AsString;
        cdsSO122BOX.FieldByName('UPDTIME').AsString := cdsSO122Data.FieldByName('UPDTIME').AsString;
        cdsSO122BOX.FieldByName('REALDATE').AsString := cdsSO122Data.FieldByName('RealDate').AsString;
        cdsSO122BOX.FieldByName('SHOULDDATE').AsString := cdsSO122Data.FieldByName('ShouldDate').AsString;
        //cdsSO122BOX.FieldByName('REALAMT').AsString := cdsSO122Data.FieldByName('REALAMT').AsString;
        cdsSO122BOX.FieldByName('PROMCODE').AsString := cdsSO122Data.FieldByName('PROMCODE').AsString;
        cdsSO122BOX.FieldByName('PROMNAME').AsString := cdsSO122Data.FieldByName('PROMNAME').AsString;
        cdsSO122BOX.FieldByName('CustID').AsString := cdsSO122Data.FieldByName('CustID').AsString;

        cdsSO122BOX.Post;

        cdsSO122Data.Next;
      end;



      sL_QueryString:='�p��ɶ�:' + sL_FileCurrDateTime + '@' + '����H��:' + frmMainMenu.sG_User + '@' +
                      '���q�W��:' + frmMainMenu.sG_CompName + '@' + '�k�ݦ~��:' + sI_BelongYM;

      sL_FileName:='C:\BOX����' + sL_CurrDateTime + '.XLS';

      cdsSO122BOX.DisableControls;
      DataSetToXLS(cdsSO122BOX,sL_FileName,sL_QueryString,'BUYORRENT');
      cdsSO122BOX.EnableControls;

      if cdsSO122BOX.RecordCount > 0 then
        MessageDlg('�p�⧹��,BOX������Ʀs�� ' + sL_FileName,mtInformation,[mbOK],0)
      else
        MessageDlg('�S��BOX�������',mtInformation,[mbOK],0)
    end;


end;

procedure TdtmMain1.getAccountParameter(I_Adoqry: TADOQuery);
begin
   //�����H�Υd�ʤ���/���B-���s�H  FirstCreditCard1
  fG_FirCardSpd := I_Adoqry.FieldByName('FirstCreditCard1').AsFloat;
  //�����D�H�Υd�ʤ���/���B-���s�H FirstNotCreditCard1
  fG_FirNoCardSpd := I_Adoqry.FieldByName('FirstNotCreditCard1').AsFloat;

  fG_FirCardCharge := I_Adoqry.FieldByName('FirstCreditCard2').AsFloat;
  //�����D�H�Υd�ʤ���/���B-���O�� FirstNotCreditCard2
  fG_FirNoCardCharge := I_Adoqry.FieldByName('FirstNotCreditCard2').AsFloat;
  //�򦬫H�Υd�ʤ���/���B-���O��   OtherCreditCard2
  fG_ContCardCharge := I_Adoqry.FieldByName('OtherCreditCard2').AsFloat;
  //�򦬫D�H�Υd�ʤ���/���B-���O��  OtherNotCreditCard2
  fG_ContNoCardCharge := I_Adoqry.FieldByName('OtherNotCreditCard2').AsFloat;



  sG_PayUnit := I_Adoqry.FieldByName('PayUnit').AsString;

  //�W�D�ݦ����� �h�֤Ѥ~������
  nG_ChannelViewDays := I_Adoqry.FieldByName('ChannelViewDays').AsInteger;

end;

procedure TdtmMain1.getCD042Data;
var
  sL_SQL:String;
begin
  sL_SQL:='SELECT CODENO,DESCRIPTION FROM CD042';

  with qryCD042Data do
  begin
    Close;
    SQL.Clear;
    SQL.Add(sL_SQL);
    Open;
  end;

end;

function TdtmMain1.CheckProm(sI_OrderNo,sI_PROMCODE,sI_PROMNAME:String):boolean;
Var
  sL_SQL,sL_PROMNAME,sL_PROMCODE : String;
  jj : Integer;
begin
{//�諾����SO007��X���
  sL_SQL:='SELECT PROMCODE,PROMNAME  FROM SO105 WHERE ORDERNO=''' + sI_OrderNo + '''';

  with qrySO105Prom do
  begin
    Close;
    SQL.Clear;
    SQL.Add(sL_SQL);
    OPEN;
  end;

  sI_PROMCODE:=qrySO105Prom.FieldByName('PROMCODE').AsString;
  sI_PROMNAME:=qrySO105Prom.FieldByName('PROMNAME').AsString;
}

  //�ˬd��ƬO�_���n�ư����~�Ȭ���
  for jj:=0 to G_PromCodeStrList.Count-1 do
    begin
      sL_PROMCODE := G_PromCodeStrList.Strings[jj];
      sL_PROMNAME := G_PromNameStrList.Strings[jj];

      if sL_PROMCODE = sI_PROMCODE then
      begin
        Result:=true;
        Break;
      end
      else
        Result:=false;
    end;
end;

function TdtmMain1.getDateData(sI_SEQNo,sI_ComputeYM,sI_Table: String; var dI_RealDate,
  dI_ShouldDate: TDate) : Boolean;
Var
  sL_SQL,sL_StartDate,sL_StopDate,sL_MonthDay:String;
begin
    sL_StartDate:=sI_ComputeYM+'/01';
    sL_MonthDay:=IntToStr(TUdateTime.DaysOfMonth(sL_StartDate));
    sL_StopDate:=sI_ComputeYM+'/'+sL_MonthDay;
    sL_SQL:='select RealDate,ShouldDate from ' + sI_Table +
            ' where FaciSeqNo=''' + sI_SEQNo + ''' AND ROWNUM <=1 ORDER BY ShouldDate ' ;


    with qryDateData do
    begin
      Close;
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;

      dI_RealDate:=qryDateData.FieldByName('REALDATE').AsDateTime;
      dI_ShouldDate:=qryDateData.FieldByName('SHOULDDATE').AsDateTime;

      if qryDateData.RecordCount = 0 then
        Result := false
      else
        Result := true;
    end;

end;

procedure TdtmMain1.qrySO132BeforePost(DataSet: TDataSet);
var
    nL_CompCode, nL_SeqNo : Integer;
    sL_SQL : String;
    bL_BeModified : boolean;
    nL_OldCItemCode, nL_NewCItemCode : Integer;
    sL_OldCItemName, sL_NewCItemName : String;
    sL_OldEmpNo, sL_OldEmpName, sL_NewEmpNo, sL_NewEmpName : String;
begin
  bL_BeModified := false;

    if DataSet.State = dsEdit then
    begin
//      nL_OldCItemCode := 'NULL';
      sL_OldCItemName := 'NULL';
//      sL_NewCItemCode := 'NULL';
      sL_NewCItemName := 'NULL';
      sL_OldEmpNo := 'NULL';
      sL_OldEmpName := 'NULL';
      sL_NewEmpNo := 'NULL';
      sL_NewEmpName := 'NULL';

      //down, �˴� ���O���� �O�_�Q����..
      if VarToStr(DataSet.FieldByName('CITEMCODE').OldValue) <> VarToStr(DataSet.FieldByName('CITEMCODE').NewValue) then
      begin
        bL_BeModified := true;
        nL_OldCItemCode := StrToInt(VarToStr(DataSet.FieldByName('CITEMCODE').OldValue));
        sL_OldCItemName := VarToStr(DataSet.FieldByName('CITEMNAME').OldValue);
        nL_NewCItemCode := StrToInt(VarToStr(DataSet.FieldByName('CITEMCODE').NewValue));
        sL_NewCItemName := VarToStr(DataSet.FieldByName('CITEMNAME').NewValue);
        {
        Memo1.Lines.Add('sL_OldCItemCode=' + sL_OldCItemCode);
        Memo1.Lines.Add('sL_OldCItemName=' + sL_OldCItemName);
        Memo1.Lines.Add('sL_NewCItemCode=' + sL_NewCItemCode);
        Memo1.Lines.Add('sL_NewCItemName=' + sL_NewCItemName);
        }
      end;
      //up, �˴� ���O���� �O�_�Q����..

      //down, �˴� ���O���� �O�_�Q����..
      if VarToStr(DataSet.FieldByName('EMPNO').OldValue) <> VarToStr(DataSet.FieldByName('EMPNO').NewValue) then
      begin
        bL_BeModified := true;
        sL_OldEmpNo := VarToStr(DataSet.FieldByName('EMPNO').OldValue);
        sL_OldEmpName := VarToStr(DataSet.FieldByName('EMPNAME').OldValue);
        sL_NewEmpNo := VarToStr(DataSet.FieldByName('EMPNO').NewValue);
        sL_NewEmpName := VarToStr(DataSet.FieldByName('EMPNAME').NewValue);
        {
        Memo1.Lines.Add('sL_OldEmpNo=' + sL_OldEmpNo);
        Memo1.Lines.Add('sL_OldEmpName=' + sL_OldEmpName);
        Memo1.Lines.Add('sL_NewEmpNo=' + sL_NewEmpNo);
        Memo1.Lines.Add('sL_NewEmpName=' + sL_NewEmpName);
        }
      end;
      //up, �˴� ���O���� �O�_�Q����..

      if bL_BeModified then
      begin
        with qryCommon do
        begin
          SQL.Clear;
          SQL.Add('select nvl(max(seq),1) seq from so133');
          Open;
          nL_SeqNo := FieldByName('seq').AsInteger + 1;
          Close;
        end;

        nL_CompCode := DataSet.FieldByName('COMPCODE').AsInteger;
        with qrySo133 do
        begin
          if state=dsInactive then
            Open;
          append;
          FieldByName('SEQ').AsInteger := nL_SeqNo;
          FieldByName('COMPCODE').AsInteger := nL_CompCode;
          FieldByName('OLDCITEMCODE').AsInteger := nL_OldCItemCode;
          FieldByName('OLDCITEMNAME').Value := sL_OldCItemName;
          FieldByName('NEWCITEMCODE').AsInteger := nL_NewCItemCode;
          FieldByName('NEWCITEMNAME').Value := sL_NewCItemName;

          FieldByName('OLDEMPNO').Value := sL_OldEmpNo;
          FieldByName('OLDEMPNAME').Value := sL_OldEmpName;
          FieldByName('NEWEMPNO').Value := sL_NewEmpNo;
          FieldByName('NEWEMPNAME').Value := sL_NewEmpName;
          FieldByName('OPERATIONMODE').AsString := 'U';
          FieldByName('UPDEN').AsString := frmMainMenu.sG_User;
          FieldByName('UPDTIME').AsDateTime := now;
          post;
        end;

{
        with ADOCommand1 do
        begin
          sL_SQL := 'insert into SO133(SEQ,COMPCODE,OLDCITEMCODE,OLDCITEMNAME,NEWCITEMCODE,NEWCITEMNAME,OLDEMPNO,OLDEMPNAME,NEWEMPNO,NEWEMPNAME,OPERATIONMODE,UPDEN,UPDTIME)';
          sL_SQL := sL_SQL + ' values((select nvl(max(seq),1) from so133)+1,' + IntToStr(nL_CompCode) + ',' +  sL_OldCItemCode + ',' ;
          sL_SQL := sL_SQL + STR_SEP + sL_OldCItemName + STR_SEP + ',' + sL_NewCItemCode + ',' ;
          sL_SQL := sL_SQL + STR_SEP + sL_NewCItemName + STR_SEP + ',' ;
          sL_SQL := sL_SQL + STR_SEP + sL_OldEmpNo + STR_SEP + ',' ;
          sL_SQL := sL_SQL + STR_SEP + sL_OldEmpName + STR_SEP + ',' ;
          sL_SQL := sL_SQL + STR_SEP + sL_NewEmpNo + STR_SEP + ',' ;
          sL_SQL := sL_SQL + STR_SEP + sL_NewEmpName + STR_SEP + ',' ;
          sL_SQL := sL_SQL + STR_SEP + 'U' + STR_SEP + ',' + STR_SEP + sG_User + STR_SEP + ',';
          sL_SQL := sL_SQL + TUstr.getOracleSQLDateTimeStr(now) + ')';
          Memo1.Lines.Add(sL_SQL);
          ADOCommand1.CommandText := sL_SQL;
          Memo1.Lines.Add(ADOCommand1.CommandText);
          ADOCommand1.Execute;

        end;

        Showmessage('modified..');
}
      end;
    end;
end;



procedure TdtmMain1.qrySO132BeforeDelete(DataSet: TDataSet);
var
    nL_CompCode, nL_SeqNo : Integer;
    sL_SQL : String;
    bL_BeModified : boolean;
    nL_OldCItemCode, nL_NewCItemCode : Integer;
    sL_OldCItemName, sL_NewCItemName : String;
    sL_OldEmpNo, sL_OldEmpName, sL_NewEmpNo, sL_NewEmpName : String;
begin
  bL_BeModified := false;

    if DataSet.State = dsBrowse	 then
    begin
//      nL_OldCItemCode := 'NULL';
      sL_OldCItemName := 'NULL';
//      sL_NewCItemCode := 'NULL';
      sL_NewCItemName := 'NULL';
      sL_OldEmpNo := 'NULL';
      sL_OldEmpName := 'NULL';
      sL_NewEmpNo := 'NULL';
      sL_NewEmpName := 'NULL';

      //down, �˴� ���O���� �O�_�Q����..
      if VarToStr(DataSet.FieldByName('CITEMCODE').OldValue) <> VarToStr(DataSet.FieldByName('CITEMCODE').NewValue) then
      begin
        bL_BeModified := true;
        nL_OldCItemCode := StrToInt(VarToStr(DataSet.FieldByName('CITEMCODE').OldValue));
        sL_OldCItemName := VarToStr(DataSet.FieldByName('CITEMNAME').OldValue);
        nL_NewCItemCode := StrToInt(VarToStr(DataSet.FieldByName('CITEMCODE').NewValue));
        sL_NewCItemName := VarToStr(DataSet.FieldByName('CITEMNAME').NewValue);

      end;
      //up, �˴� ���O���� �O�_�Q����..

      //down, �˴� ���O���� �O�_�Q����..
      if VarToStr(DataSet.FieldByName('EMPNO').OldValue) <> VarToStr(DataSet.FieldByName('EMPNO').NewValue) then
      begin
        bL_BeModified := true;
        sL_OldEmpNo := VarToStr(DataSet.FieldByName('EMPNO').OldValue);
        sL_OldEmpName := VarToStr(DataSet.FieldByName('EMPNAME').OldValue);
        sL_NewEmpNo := VarToStr(DataSet.FieldByName('EMPNO').NewValue);
        sL_NewEmpName := VarToStr(DataSet.FieldByName('EMPNAME').NewValue);

      end;
      //up, �˴� ���O���� �O�_�Q����..

      if bL_BeModified then
      begin
        with qryCommon do
        begin
          SQL.Clear;
          SQL.Add('select nvl(max(seq),1) seq from so133');
          Open;
          nL_SeqNo := FieldByName('seq').AsInteger + 1;
          Close;
        end;

        nL_CompCode := DataSet.FieldByName('COMPCODE').AsInteger;
        with qrySo133 do
        begin
          if state=dsInactive then
            Open;
          append;
          FieldByName('SEQ').AsInteger := nL_SeqNo;
          FieldByName('COMPCODE').AsInteger := nL_CompCode;
          FieldByName('OLDCITEMCODE').AsInteger := nL_OldCItemCode;
          FieldByName('OLDCITEMNAME').Value := sL_OldCItemName;
          FieldByName('NEWCITEMCODE').AsInteger := nL_NewCItemCode;
          FieldByName('NEWCITEMNAME').Value := sL_NewCItemName;

          FieldByName('OLDEMPNO').Value := sL_OldEmpNo;
          FieldByName('OLDEMPNAME').Value := sL_OldEmpName;
          FieldByName('NEWEMPNO').Value := sL_NewEmpNo;
          FieldByName('NEWEMPNAME').Value := sL_NewEmpName;
          FieldByName('OPERATIONMODE').AsString := 'D';
          FieldByName('UPDEN').AsString := frmMainMenu.sG_User;
          FieldByName('UPDTIME').AsDateTime := now;
          post;                              
        end;
      end;
    end;
end;






function TdtmMain1.ReplaceStr(sI_SrcString, sI_SepFlag: String): String;
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

procedure TdtmMain1.cdsSO122BOXCalcFields(DataSet: TDataSet);
var
    nL_BuyCode : Integer;
begin

    nL_BuyCode := cdsSO122BOX.FieldByName('BUYORRENT').AsInteger;
    if nL_BuyCode = BUY_BOX then
      cdsSO122BOX.FieldByName('BuyName').AsString := '�R'
    else if nL_BuyCode = RENT_BOX then
      cdsSO122BOX.FieldByName('BuyName').AsString := '��';

end;

function TdtmMain1.getEmpData(sI_CompCode, sI_CustID, sI_CitemCode: String;
  var sI_EmpNo, sI_EmpName: String) : Boolean;
var
    sL_SQL : String;
begin
    sL_SQL := 'SELECT * FROM SO132 WHERE CompCode=' + sI_CompCode +
              ' AND CustID=' + sI_CustID + ' AND CitemCode=' + sI_CitemCode;


    with qryCommon do
    begin
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;

      sI_EmpNo := FieldByName('EmpNo').AsString;
      sI_EmpName := FieldByName('EmpName').AsString;

      if qryCommon.RecordCount = 0 then
        Result := false
      else
        Result := true;
      Close;
    end;
end;

procedure TdtmMain1.transSo033So034Notation(I_CDS : TClientDataSet);
var
    sL_StrRealAmt : String;
begin
    I_CDS.First;
    with I_CDS do
    begin
      while not Eof do
      begin
        sL_StrRealAmt := FieldByName('StrRealAmt').AsString;
        if sL_StrRealAmt <> '' then
        begin
          Edit;
          FieldByName('RealAmt').AsInteger := StrToInt(sL_StrRealAmt);
          Post;
        end;

        Next;
      end;
    end;
end;

procedure TdtmMain1.transSo122Notation(I_CDS: TClientDataSet);
var
    sL_StrWorkeren1Comm,sL_StrClctenComm,sL_StrIntaccComm,sL_StrRealAmt : String;
begin
    I_CDS.First;
    with I_CDS do
    begin
      while not Eof do
      begin
        Edit;

        sL_StrRealAmt := FieldByName('StrRealAmt').AsString;
        if sL_StrRealAmt <> '' then
          FieldByName('RealAmt').AsFloat := StrToFloat(sL_StrRealAmt);

        sL_StrWorkeren1Comm := FieldByName('StrWorkeren1Comm').AsString;
        if sL_StrWorkeren1Comm <> '' then
          FieldByName('Workeren1Comm').AsFloat := StrToFloat(sL_StrWorkeren1Comm);

        sL_StrClctenComm := FieldByName('StrClctenComm').AsString;
        if sL_StrClctenComm <> '' then
          FieldByName('ClctenComm').AsFloat := StrToFloat(sL_StrClctenComm);

        sL_StrIntaccComm := FieldByName('StrIntaccComm').AsString;
        if sL_StrIntaccComm <> '' then
          FieldByName('IntaccComm').AsFloat := StrToFloat(sL_StrIntaccComm);

        Post;

        Next;
      end;
    end;
end;

function TdtmMain1.getISno(sI_Sno : String;var sI_ReturnSno : String): Integer;
var
    sL_SQL,sL_Sno : String;
    nL_TargetPos : Integer;
begin
    sL_SQL := 'select a.SNo, a.IntroId, a.IntroName, a.WorkerEn1,a.WorkerName1,a.WorkerEn2,a.WorkerName2, ';
    sL_SQL := sL_SQL + 'a.WorkerEn3,a.WorkerName3,a.AcceptEn, a.AcceptName, a.MediaCode , a.MediaName ,a.ORDERNO,b.FaciSNo ,b.InstDate ,b.PRDate ,a.CustID,b.SeqNo,b.BuyCode ';
    sL_SQL := sL_SQL + ' from So007 a, So004 b ';
    sL_SQL := sL_SQL + ' where  a.SNo=b.SNo AND b.SmartCardNo Is Not Null';
    sL_SQL := sL_SQL + ' AND b.BuyCode IN (1,3) AND b.FaciCode Is Not Null and b.PRSNO=''' + sI_Sno + '''';

    with cdsTempBox do
    begin
      Close;
      CommandText := sL_SQL;
      Open;

      if RecordCount > 0 then
      begin
        sL_Sno := FieldByName('SNO').AsString;
        nL_TargetPos := Pos('I',sL_Sno);
        if nL_TargetPos > 0 then
          Result := 1     //��� I ��
        else
          Result := -2;   //����Ƥ��O I ��,�~���
      end
      else
        Result := -1;     //�S�� I ��

      sI_ReturnSno := sL_Sno;
    end;


end;

procedure TdtmMain1.dealwithSTBMSno(sI_YM: String);
var
    sL_SQL,sL_StartInstDate,sL_StopInstDate,sL_MonthDays : String;
    sL_Sno,sL_ReturnSno,sL_Year,sL_Month,sL_YM,sL_StbNo : String;
    nL_CheckNo : Integer;
    bL_IsAlreadyPay : Boolean;
    wL_Year, wL_Month, wL_Day : word;
    dL_InstDate : TDateTime;
    dL_MSnoInstDate,dL_MSnoPRDate : TDateTime;
begin
    //select �Ӧ~�� (sI_YM) ���Ҧ��˾����
    sL_StartInstDate := sI_YM + '/01';
    sL_MonthDays := IntToStr(TUdateTime.DaysOfMonth(sL_StartInstDate));
    sL_StopInstDate := sI_YM + '/' + sL_MonthDays + ' 23:59:59';
    sL_StartInstDate := sL_StartInstDate + ' 00:00:00';

    //�B�z M ��,�B�u�B�z���ζR
    sL_SQL := 'select SNo,InstDate ,PRDate,SeqNo,PRSNO,CustId,FaciSNo ';
    sL_SQL := sL_SQL + ' from So004 ';
    sL_SQL := sL_SQL + ' where (InstDate BETWEEN ' + TUstr.getOracleSQLDateTimeStr(StrToDateTime(sL_StartInstDate));
    sL_SQL := sL_SQL + ' AND ' + TUstr.getOracleSQLDateTimeStr(StrToDateTime(sL_StopInstDate))+')';
    sL_SQL := sL_SQL + ' AND SmartCardNo Is Not Null AND FaciCode Is Not Null';
    sL_SQL := sL_SQL + ' AND SNO LIKE ''%M%'' AND BuyCode in (1,3)';

//    saveToFile(sL_SQL,'C:\temp\buysL_SQL.txt');

    with qrySO004 do
    begin
      Close;
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;

      First;
      while not Eof do
      begin
        sL_Sno := qrySO004.FieldByName('SNO').AsString;


        nL_CheckNo := getISno(sL_Sno,sL_ReturnSno);

        while nL_CheckNo = -2 do  //����Ƥ��O I ��,�~���
        begin
          sL_Sno := sL_ReturnSno;
          nL_CheckNo := getISno(sL_Sno,sL_ReturnSno);
        end;

        //�������� I ����
        if nL_CheckNo = 1 then
        begin
          sL_Sno := cdsTempBox.FieldByName('SNO').AsString;
          dL_InstDate := cdsTempBox.FieldByName('InstDate').AsDateTime;
          sL_StbNo := cdsTempBox.FieldByName('FaciSNo').AsString;

          DecodeDate(dL_InstDate,wL_Year, wL_Month, wL_Day);
          sL_Year := IntToStr((wL_Year - 1911));
          if Length(sL_Year) < 3 then
            sL_Year := '0' + sL_Year;

          sL_Month := IntToStr(wL_Month);
          if Length(sL_Month) < 2 then
            sL_Month := '0' + sL_Month;

          sL_YM := sL_Year + sL_Month;

          bL_IsAlreadyPay := checkedIsAlreadyPay(sL_YM,sL_Sno,frmMainMenu.sG_CompCode,sL_StbNo);

          //�Y�w�g���L�����A�h�� M ���ƴN���A�������C
          if not bL_IsAlreadyPay then
          begin
            //showmessage('����');
            //insertMSnoSO122(����Ӹ˾��檺DataSet,���׳檺DataSet);
            insertMSnoSO122(cdsTempBox,qrySO004);
          end
          else
          begin
            //showmessage('���L')
          end
        end
        else if nL_CheckNo = -1 then  //�S�������� I ��
        begin
          //showmessage('�S�������� I ��');
          cdsSo033Log.Append;
          cdsSo033Log.FieldByName('CUSTID').AsString := qrySO004.FieldByName('CUSTID').AsString;
          cdsSo033Log.FieldByName('SNo').AsString := qrySO004.FieldByName('SNo').AsString;

          dL_MSnoInstDate := qrySO004.FieldByName('InstDate').AsDateTime;
          if dL_MSnoInstDate <> 0 then
            cdsSo033Log.FieldByName('InstDate').AsDateTime := dL_MSnoInstDate;

          dL_MSnoPRDate := qrySO004.FieldByName('PRDate').AsDateTime;
          if dL_MSnoPRDate <> 0 then
            cdsSo033Log.FieldByName('PRDate').AsDateTime := dL_MSnoPRDate;

          cdsSo033Log.FieldByName('SeqNo').AsString := qrySO004.FieldByName('SeqNo').AsString;
          cdsSo033Log.FieldByName('PRSNO').AsString := qrySO004.FieldByName('PRSNO').AsString;
          cdsSo033Log.FieldByName('Notes').AsString := 'M ��������� I ��';

          cdsSo033Log.Post;
        end;

        Next;
      end;
    end;
end;

function TdtmMain1.checkedIsAlreadyPay(sI_YM, sI_Sno,sI_CompCode,sI_StbNo : String): Boolean;
var
    sL_SQL : String;
begin
    sL_SQL := 'SELECT COUNT(*) counts FROM SO122 WHERE ComputeYM=''' + sI_YM +
              ''' AND Sno=''' + sI_Sno + ''' AND COMPCODE=' + sI_CompCode +
              ' AND INTACCCOMM > 0 AND StbNo=''' + sI_StbNo + '''';

    with qryCom do
    begin
      Close;
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;

      if FieldByName('COUNTS').AsInteger > 0 then
        Result := true
      else
        Result := false;
    end;
end;

procedure TdtmMain1.insertMSnoSO122(I_DateSet,I_MSnoDateSet: TDataset);
var
    nL_IntrComm, nL_AccComm, nL_WorkenComm ,nL_BuyOrRent : Integer;
    sL_MediaCode,sL_MediaName,sL_RefNo ,sL_BelongDays,sL_NextBelongYM,sL_OrderNo: String;
    dL_StartDate,dL_StopDate,dL_CurDate : TDateTime;
    dL_RealDate,dL_ShouldDate,dL_PayDate : TDate;
    nL_DiffDays : Integer;
    sL_PromCode,sL_PromName,sL_ComputeYM : String;
    bL_Result,bL_HaveDateData : Boolean;
    sL_CustID,sL_SeqNo : String;

    wL_Year, wL_Month, wL_Day : Word;
    sL_Year,sL_PayYM,sL_Sno : String;
    nL_OperationFlag,nL_BuyCode : Integer;
    bL_IsTheSameMonth : Boolean;
begin
    //�N M �����x�s�� SO122
    sL_Sno := I_DateSet.FieldByName('SNo').AsString;
    nL_BuyCode := I_DateSet.FieldByName('BuyCode').AsInteger;
    if nL_BuyCode = 1 then
      nL_OperationFlag := BUY_BOX
    else if nL_BuyCode = 3 then
      nL_OperationFlag := RENT_BOX;

    sL_OrderNo:=I_DateSet.FieldByName('ORDERNO').AsString;

    bL_Result := CheckProm(sL_OrderNo,sL_PromCode,sL_PromName);

    //�Y�O�n�ư����~�Ȭ���,��������
    if not bL_Result then
    begin
      case nL_OperationFlag of
        BUY_BOX :
        begin
          nL_IntrComm := nG_BuyBoxIntrComm;
          nL_AccComm := nG_BuyBoxAcceptComm;
          nL_WorkenComm := nG_SetBoxWorkerComm;
          nL_BuyOrRent := BUY_BOX;
        end;
        RENT_BOX :
        begin
          nL_IntrComm := nG_RentBoxIntrComm;
          nL_AccComm := nG_RentBoxAcceptComm;
          nL_WorkenComm := nG_SetBoxWorkerComm;
          nL_BuyOrRent := RENT_BOX;
        end;
      end;

      with cdsSo122 do
      begin
        Append;

        FieldByName('CompCode').AsString:=frmMainMenu.sG_CompCode;
        FieldByName('ComputeYM').AsString:=sG_ComputeYM;
        FieldByName('BELONGYM').AsString:=sG_BelongYM;

        //�� M ����
        FieldByName('StbNo').AsString := I_MSnoDateSet.FieldByName('FaciSNo').AsString;
        FieldByName('SNo').AsString := I_MSnoDateSet.FieldByName('SNO').AsString;


        sL_ComputeYM:=Copy(sG_ComputeYM,1,3) + '/' + Copy(sG_ComputeYM,4,5);
        sL_ComputeYM := chineseYMChangeToEnglishYM(sL_ComputeYM) ;
        //sL_ComputeYM :=Copy(sL_ComputeYM,1,4) + Copy(sL_ComputeYM,6,2)+ '/01';


        //�Y SO033 �S��Ƨ� SO034
        sL_SeqNo := I_DateSet.FieldByName('SeqNo').AsString;
        bL_HaveDateData := getDateData(sL_SeqNo,sL_ComputeYM,'so033',dL_RealDate,dL_ShouldDate);
        if not bL_HaveDateData then
          getDateData(sL_SeqNo,sL_ComputeYM,'so034',dL_RealDate,dL_ShouldDate);


        if I_DateSet.FieldByName('WorkerEn1').AsString='' then
            FieldByName('WorkerEn1Comm').AsFloat:=0
        else
        begin
          FieldByName('WorkerEn1').AsString:=I_DateSet.FieldByName('WorkerEn1').AsString;
          FieldByName('WorkerEn1Name').AsString:=I_DateSet.FieldByName('WorkerName1').AsString;

          //�� I ��ɵ��L���ΦA��
          //FieldByName('WorkerEn1Comm').AsFloat:=nL_WorkenComm;
          FieldByName('WorkerEn1Comm').AsFloat := 0;
        end;

        //�ˬd�O�_���ۤv���u,�O���ܤ~�o����
        sL_MediaCode := I_DateSet.FieldByName('MediaCode').AsString;
        sL_MediaName := I_DateSet.FieldByName('MediaName').AsString;

        if sL_MediaCode <> '' then
        begin
          if qryCodeCD009.Locate('CODENO', StrToInt(sL_MediaCode),[]) then
          begin
            sL_RefNo := '';
            sL_RefNo := qryCodeCD009.FieldByName('RefNo').AsString;
          end;
        end;

        if I_DateSet.FieldByName('InstDate').AsString <> '' then
          dL_StartDate := StrToDate(DateToStr(I_DateSet.FieldByName('InstDate').AsDateTime));

        sL_BelongDays := Copy(sG_BelongYM,1,3) + '/' + Copy(sG_BelongYM,4,5);
        sL_BelongDays := chineseYMChangeToEnglishYM(sL_BelongDays) + '/01';


        //�Y�L������,�h�H�k�ݦ~��Ĥ@�Ѩӭp��
        if I_MSnoDateSet.FieldByName('PRDate').AsString <> '' then
        begin
          bL_IsTheSameMonth := checkIsTheSameMonth(I_DateSet.FieldByName('InstDate').AsString,I_MSnoDateSet.FieldByName('PRDate').AsString);
          if bL_IsTheSameMonth then  //�Y�������P�k�ݤ���P�@�Ӥ�
            dL_StopDate := StrToDate(DateToStr(I_MSnoDateSet.FieldByName('PRDate').AsDateTime))
          else //���P��
            dL_StopDate := StrToDate(sL_BelongDays);
        end
        else
          dL_StopDate := StrToDate(sL_BelongDays);

        nL_DiffDays := Round(dL_StopDate - dL_StartDate);


        //�Y�����ФH
        if I_DateSet.FieldByName('IntroId').AsString<>'' then
        begin
          FieldByName('IntAccId').AsString:=I_DateSet.FieldByName('IntroId').AsString;
          FieldByName('IntAccName').AsString:=I_DateSet.FieldByName('IntroName').AsString;

          //���ФH�O�ۤv���u�~�p��
          if sL_RefNo <> '1' then
          begin
            if ((nL_BuyOrRent = BUY_BOX) and (nL_DiffDays > nG_BuyBoxUseDays))
               or ((nL_BuyOrRent = RENT_BOX) and (nL_DiffDays > nG_RentBoxUseDays)) then
            begin
              FieldByName('IntAccComm').AsFloat := nL_IntrComm
            end
            else
              FieldByName('IntAccComm').AsFloat := 0;
          end;

          FieldByName('StaffCode').AsString:='1'; //��ܬO���ФH
          FieldByName('CallSelf').AsInteger:=0; //�D�ۨӹq
        end;

        //�Y�L���ФH�������z�H��
        if (I_DateSet.FieldByName('IntroId').AsString='') and (I_DateSet.FieldByName('AcceptEn').AsString<>'') then
        begin
          FieldByName('IntAccId').AsString:=I_DateSet.FieldByName('AcceptEn').AsString;
          FieldByName('IntAccName').AsString:=I_DateSet.FieldByName('AcceptName').AsString;

          if ((nL_BuyOrRent = BUY_BOX) and (nL_DiffDays > nG_BuyBoxUseDays))
             or ((nL_BuyOrRent = RENT_BOX) and (nL_DiffDays > nG_RentBoxUseDays)) then
          begin
            FieldByName('IntAccComm').AsFloat := nL_AccComm
          end
          else
            FieldByName('IntAccComm').AsFloat := 0;

          FieldByName('StaffCode').AsString:='2';  //��ܬO���z�H
          FieldByName('CallSelf').AsInteger:=1; //�ۨӹq
        end;

        FieldByName('OPERATOR').AsString:=frmMainMenu.sG_User;
        //FieldByName('UpdTime').AsString:=IntToStr(StrToInt(FormatDateTime('YYYY',now))-1911)+ FormatDateTime('/MM/DD hh:mm:ss',now);
        FieldByName('UpdTime').AsString := sG_CurDate;

        FieldByName('BUYORRENT').AsInteger := nL_BuyOrRent;

        //BOX �˾����
        if I_DateSet.FieldByName('InstDate').AsString <> '' then
          FieldByName('RealStartDate').AsDateTime := I_DateSet.FieldByName('InstDate').AsDateTime;
        //BOX ������
        if I_MSnoDateSet.FieldByName('PRDate').AsString <> '' then
          FieldByName('RealStopDate').AsDateTime := I_MSnoDateSet.FieldByName('PRDate').AsDateTime;

        //CD009 �� RefNo=2(���u����),RefNo=3(�P���I)
        //�YRefno=2,�h����쬰SO CompCode,�_�h��MediaCode
        if sL_RefNo='2' then
        begin
          FieldByName('MediaCode').AsInteger := StrToInt(frmMainMenu.sG_CompCode);
          FieldByName('MediaName').AsString := frmMainMenu.sG_CompName;
        end
        else if (sL_RefNo='3') or (sL_RefNo='5') then
        begin
          FieldByName('MediaCode').AsInteger := StrToInt(sL_MediaCode);
          FieldByName('MediaName').AsString := sL_MediaName;
        end;


        if sL_PromCode <> '' then
          FieldByName('PromCode').AsInteger:=StrToInt(sL_PromCode);

        FieldByName('PromName').AsString:=sL_PromName;

        if dL_RealDate <> 0  then
          FieldByName('REALDATE').AsDateTime:=dL_RealDate;

        if dL_ShouldDate <> 0 then
          FieldByName('SHOULDDATE').AsDateTime:=dL_ShouldDate;

        sL_CustID := I_DateSet.FieldByName('CustID').AsString;
        if sL_CustID <> '' then
        FieldByName('CustID').AsInteger := StrToInt(sL_CustID);

        FieldByName('OrderNo').AsString := sL_OrderNo;

        Post;


        //*********************************************************************
        //*********************************************************************
        { �Y�R�_�ί��ΨϥΤ��,���W�L�W�w���� ,  ���F�W�[�@�����k�ݦ~�묰�s����ƥ~
          , �A�N�����W�[�ܸӵ������  }
        if ((nL_BuyOrRent = BUY_BOX) and (nL_DiffDays <= nG_BuyBoxUseDays))
           or ((nL_BuyOrRent = RENT_BOX) and (nL_DiffDays <= nG_RentBoxUseDays)) then
        begin
          //�p��n�k�ݨ���@�Ӥ� --- Down
          if nL_BuyOrRent = BUY_BOX then    //�R
            dL_PayDate := dL_StartDate + nG_BuyBoxUseDays
          else if nL_BuyOrRent = RENT_BOX then    //��
            dL_PayDate := dL_StartDate + nG_RentBoxUseDays;

          DecodeDate(dL_PayDate,wL_Year, wL_Month, wL_Day);
          sL_Year := IntToStr((wL_Year - 1911));
          if Length(sL_Year) < 3 then
            sL_Year := '0' + sL_Year;

          sL_PayYM := sL_Year + IntToStr(wL_Month);

          sL_NextBelongYM := afterMonth(sL_PayYM,1);

          if Length(sL_NextBelongYM) = 4 then
          begin
            sL_NextBelongYM := '0' + sL_NextBelongYM;
          end;

          if sG_BelongYM > sL_NextBelongYM then
            sL_NextBelongYM := sG_BelongYM;
          //�p��n�k�ݨ���@�Ӥ� --- Up


          Append;
          FieldByName('CompCode').AsString:=frmMainMenu.sG_CompCode;
          FieldByName('ComputeYM').AsString:=sG_ComputeYM;

          //�N�����[��U�Ӥ�
          FieldByName('BELONGYM').AsString:=sL_NextBelongYM;

          //�� M �檺���
          FieldByName('StbNo').AsString := I_MSnoDateSet.FieldByName('FaciSNo').AsString;
          FieldByName('SNo').AsString := I_MSnoDateSet.FieldByName('SNo').AsString;


          FieldByName('WorkerEn1').AsString:=I_DateSet.FieldByName('WorkerEn1').AsString;
          FieldByName('WorkerEn1Name').AsString:=I_DateSet.FieldByName('WorkerName1').AsString;
          //�u�@�H����e�@�Ӥ�,�w�ˮɧY������
          FieldByName('WorkerEn1Comm').AsFloat:=0;


          //�ˬd�O�_���ۤv���u,�O���ܤ~�o����
          sL_MediaCode := I_DateSet.FieldByName('MediaCode').AsString;
          sL_MediaName := I_DateSet.FieldByName('MediaName').AsString;

          if sL_MediaCode <> '' then
          begin
            if qryCodeCD009.Locate('CODENO', StrToInt(sL_MediaCode),[]) then
            begin
              sL_RefNo := '';
              sL_RefNo := qryCodeCD009.FieldByName('RefNo').AsString;
            end;
          end;

          if I_DateSet.FieldByName('InstDate').AsString <> '' then
            dL_StartDate := StrToDate(DateToStr(I_DateSet.FieldByName('InstDate').AsDateTime));

          sL_BelongDays := Copy(sG_BelongYM,1,3) + '/' + Copy(sG_BelongYM,4,5);
          sL_BelongDays := chineseYMChangeToEnglishYM(sL_BelongDays) + '/01';


          //�Y�L������,�h�H�k�ݦ~��Ĥ@�Ѩӭp��
          if I_MSnoDateSet.FieldByName('PRDate').AsString <> '' then
          begin
          {
            bL_IsTheSameMonth := checkIsTheSameMonth(I_DateSet.FieldByName('InstDate').AsString,I_MSnoDateSet.FieldByName('PRDate').AsString);
            if bL_IsTheSameMonth then  //�Y�������P�k�ݤ���P�@�Ӥ�
              dL_StopDate := StrToDate(DateToStr(I_MSnoDateSet.FieldByName('PRDate').AsDateTime))
            else //���P��
              dL_StopDate := StrToDate(sL_BelongDays);
            }
            //�ĤG���p��H�������p��
            dL_StopDate := StrToDate(DateToStr(I_MSnoDateSet.FieldByName('PRDate').AsDateTime))
          end
          else
            dL_StopDate := StrToDate(sL_BelongDays);

          nL_DiffDays := Round(dL_StopDate - dL_StartDate);


          //�Y�����ФH
          if I_DateSet.FieldByName('IntroId').AsString<>'' then
          begin
            FieldByName('IntAccId').AsString:=I_DateSet.FieldByName('IntroId').AsString;
            FieldByName('IntAccName').AsString:=I_DateSet.FieldByName('IntroName').AsString;



            //���ФH�O�ۤv���u�~�p��
            if sL_RefNo <> '1' then
            begin
              //Box�ϥδ����Y���W�L�W�w����,������U�Ӥ뵹,
              //�|���ˬdBox�O�_���,�Y�ܩ���ϥδ���,
              //��|���W�L�W�w����,�]��������
              if I_MSnoDateSet.FieldByName('PRDate').AsString <> '' then
              begin
                //if nL_DiffDays <= nG_BuyBoxUseDays then
                if ((nL_OperationFlag = 1) and (nL_DiffDays <= nG_BuyBoxUseDays))
                   or ((nL_OperationFlag = 2) and (nL_DiffDays <= nG_RentBoxUseDays))then
                begin
                  FieldByName('IntAccComm').AsFloat := 0
                end
                else
                  FieldByName('IntAccComm').AsFloat := nL_IntrComm;
              end
              else
                FieldByName('IntAccComm').AsFloat := nL_IntrComm;
            end;

            FieldByName('StaffCode').AsString:='1'; //��ܬO���ФH
            FieldByName('CallSelf').AsInteger:=0; //�D�ۨӹq
          end;

          //�Y�L���ФH�������z�H��
          if (I_DateSet.FieldByName('IntroId').AsString='') and (I_DateSet.FieldByName('AcceptEn').AsString<>'') then
          begin
            FieldByName('IntAccId').AsString:=I_DateSet.FieldByName('AcceptEn').AsString;
            FieldByName('IntAccName').AsString:=I_DateSet.FieldByName('AcceptName').AsString;


            FieldByName('StaffCode').AsString:='2';  //��ܬO���z�H
            FieldByName('CallSelf').AsInteger:=1; //�ۨӹq

            //Box�ϥδ����Y���W�L�W�w����,������U�Ӥ뵹,
            //�|���ˬdBox�O�_���,�Y�ܩ���ϥδ���,
            //��|���W�L�W�w����,�]��������
            if I_MSnoDateSet.FieldByName('PRDate').AsString <> '' then
            begin
              if ((nL_OperationFlag = 1) and (nL_DiffDays <= nG_BuyBoxUseDays))
                 or ((nL_OperationFlag = 2) and (nL_DiffDays <= nG_RentBoxUseDays))then
              begin
                FieldByName('IntAccComm').AsFloat := 0
              end
              else
                FieldByName('IntAccComm').AsFloat := nL_AccComm;
            end
            else
              FieldByName('IntAccComm').AsFloat := nL_AccComm;
          end;

          FieldByName('OPERATOR').AsString:=frmMainMenu.sG_User;
          FieldByName('UpdTime').AsString := sG_CurDate;

          FieldByName('BUYORRENT').AsInteger := nL_BuyOrRent;

          //BOX �˾����
          if I_DateSet.FieldByName('InstDate').AsString <> '' then
            FieldByName('RealStartDate').AsDateTime := I_DateSet.FieldByName('InstDate').AsDateTime;
          //BOX ������
          if I_MSnoDateSet.FieldByName('PRDate').AsString <> '' then
          begin
            FieldByName('RealStopDate').AsDateTime := I_MSnoDateSet.FieldByName('PRDate').AsDateTime;

            //showmessage(I_MSnoDateSet.FieldByName('PRDate').AsString);
          end;

          //CD009 �� RefNo=2(���u����),RefNo=3(�P���I)
          //�YRefno=2,�h����쬰SO CompCode,�_�h��MediaCode
          if sL_RefNo='2' then
          begin
            FieldByName('MediaCode').AsInteger := StrToInt(frmMainMenu.sG_CompCode);
            FieldByName('MediaName').AsString := frmMainMenu.sG_CompName;
          end
          else if (sL_RefNo='3') or (sL_RefNo='5') then
          begin
            FieldByName('MediaCode').AsInteger := StrToInt(sL_MediaCode);
            FieldByName('MediaName').AsString := sL_MediaName;
          end;



          if sL_PromCode <> '' then
            FieldByName('PromCode').AsInteger:=StrToInt(sL_PromCode);

          FieldByName('PromName').AsString:=sL_PromName;

          if dL_RealDate <> 0  then
            FieldByName('REALDATE').AsDateTime:=dL_RealDate;

          if dL_ShouldDate <> 0 then
            FieldByName('SHOULDDATE').AsDateTime:=dL_ShouldDate;


          sL_CustID := I_DateSet.FieldByName('CustID').AsString;
          if sL_CustID <> '' then
          FieldByName('CustID').AsInteger := StrToInt(sL_CustID);

          FieldByName('OrderNo').AsString := sL_OrderNo;

          Post;
        end;
      end;
    end;

end;

function TdtmMain1.checkIsTheSameMonth(sI_Date1, sI_Date2: String): Boolean;
var
    wL_Year1, wL_Month1, wL_Day1 : word;
    wL_Year2, wL_Month2, wL_Day2 : word;
begin
    //�ˬd��Ӥ���O���O�P�@�Ӥ�
    DecodeDate(StrToDate(Copy(sI_Date1,1,10)),wL_Year1,wL_Month1,wL_Day2);
    DecodeDate(StrToDate(Copy(sI_Date2,1,10)),wL_Year2,wL_Month2,wL_Day2);
    if wL_Month1 = wL_Month2 then
      Result := true
    else
      Result := false;
end;

function TdtmMain1.TransToChannelExcel(sI_PayMode,sI_Flag,
  sI_BelongYM,sI_CurrDateTime : String): String;
Var
  sL_QueryString,sL_FileName,sL_FileCurrDateTime,sL_CurrDateTime : String;
  L_StrList : TStringList;
  sL_NoDataMsg,sL_HaveDataMsg : String;
  sL_YM,sL_PayModeName : String;
begin
    if sI_PayMode = ONCE_PAY then
      sL_PayModeName := '1�����I'
    else if sI_PayMode = BATCH_PAY then
      sL_PayModeName := '�������I';

    sL_YM := IntToStr(StrToInt(Copy(sI_BelongYM,1,3)))+'�~'+Copy(sI_BelongYM,4,2)+'��';

    sL_FileCurrDateTime := sI_CurrDateTime;
    sL_CurrDateTime := ReplaceStr(Trim(ReplaceStr(sL_FileCurrDateTime,'/')),':');

    //�p������`�p
    //computeExcelTotalComm(cdsSO122Channel,CHANNEL_DATA,sI_Flag);

    if sI_Flag = CHANNEL_A then   //������Ҧ�����
    begin

      if TClientDataSet(cdsSO122Channel).IndexName = 'TmpIndex' then
        TClientDataSet(cdsSO122Channel).DeleteIndex('TmpIndex');
      TClientDataSet(cdsSO122Channel).AddIndex('TmpIndex', 'CustID', [ixCaseInsensitive],'','',0);
      TClientDataSet(cdsSO122Channel).IndexName := 'TmpIndex';

      sL_FileName := sG_ExcelSavePath + '\�W�D�����Ҧ�����_' + frmMainMenu.sG_CompCode + '_' + sI_BelongYM + '_' + sL_CurrDateTime + '.XLS';
      sL_HaveDataMsg := '�p�⧹��,�W�D�����Ҧ����Ӧs��  ';
      sL_NoDataMsg := '�S���W�D�����Ҧ����Ӹ�� ';

      sL_QueryString:='����W��:' + sL_YM + '_�W�D�����Ҧ����Ӫ�' + '@' +
                      '�έp�ɶ�:' + sL_FileCurrDateTime + '@' + '����H��:' + frmMainMenu.sG_User+'@' +
                      '���q�W��:' + TUstr.replaceStr(frmMainMenu.sG_CompName,'�@','��') + '@' + '�k�ݦ~��:' + sL_YM + '@' +
                      '�d�ߦ��O����:' + frmSO8B20_1.sG_ChargeNameSQL + '@' +
                      '�����p���h:' + sL_PayModeName;

      if cdsSO122Channel.RecordCount > 0 then
      begin
        cdsSO122Channel.DisableControls;
        DataSetToXLS(cdsSO122Channel,sL_FileName,sL_QueryString,'REALSTARTDATE,REALSTOPDATE,REALDATE,SHOULDDATE');
        cdsSO122Channel.EnableControls;
      end;
    end
    else if sI_Flag = CHANNEL_I then  //��������ФH����
    begin

      if TClientDataSet(cdsSO122Channel).IndexName = 'TmpIndex' then
        TClientDataSet(cdsSO122Channel).DeleteIndex('TmpIndex');
      TClientDataSet(cdsSO122Channel).AddIndex('TmpIndex', 'INTACCID', [ixCaseInsensitive],'','',0);
      TClientDataSet(cdsSO122Channel).IndexName := 'TmpIndex';


      sL_FileName:= sG_ExcelSavePath + '\�W�D�������ФH����_' + frmMainMenu.sG_CompCode + '_' + sI_BelongYM + '_' + sL_CurrDateTime + '.XLS';
      sL_HaveDataMsg := '�p�⧹��,�W�D�������ФH���Ӧs��  ';
      sL_NoDataMsg := '�S���W�D�������ФH���Ӹ�� ';

      sL_QueryString:='����W��:' + sL_YM + '_�W�D�������ФH���Ӫ�' + '@' +
                      '�έp�ɶ�:' + sL_FileCurrDateTime + '@' + '����H��:' + frmMainMenu.sG_User+'@' +
                      '���q�W��:' + TUstr.replaceStr(frmMainMenu.sG_CompName,'�@','��') + '@' + '�k�ݦ~��:' + sL_YM + '@' +
                      '�d�ߦ��O����:' + frmSO8B20_1.sG_ChargeNameSQL + '@' +
                      '�����p���h:' + sL_PayModeName;

      if cdsSO122Channel.RecordCount > 0 then
      begin
        cdsSO122Channel.DisableControls;
        DataSetToXLS(cdsSO122Channel,sL_FileName,sL_QueryString,'CLCTEN,CLCTENNAME,CLCTENCOMM,CLCTENORIPERCENT,REALSTARTDATE,REALSTOPDATE,REALDATE,SHOULDDATE');
        cdsSO122Channel.EnableControls;
      end;
    end
    else if sI_Flag = CHANNEL_C then  //��������O�H������
    begin

      if TClientDataSet(cdsSO122Channel).IndexName = 'TmpIndex' then
        TClientDataSet(cdsSO122Channel).DeleteIndex('TmpIndex');
      TClientDataSet(cdsSO122Channel).AddIndex('TmpIndex', 'CLCTEN', [ixCaseInsensitive],'','',0);
      TClientDataSet(cdsSO122Channel).IndexName := 'TmpIndex';


      sL_FileName := sG_ExcelSavePath + '\�W�D�������O�H������_' + frmMainMenu.sG_CompCode + '_'  + sI_BelongYM + '_' + sL_CurrDateTime + '.XLS';
      sL_HaveDataMsg := '�p�⧹��,�W�D�������O�H�����Ӧs��  ';
      sL_NoDataMsg := '�S���W�D�������O�H�����Ӹ�� ';


      sL_QueryString:='����W��:' + sL_YM + '_�W�D�������O�H�����Ӫ�' + '@' +
                      '�έp�ɶ�:' + sL_FileCurrDateTime + '@' + '����H��:' + frmMainMenu.sG_User+'@' +
                      '���q�W��:' + TUstr.replaceStr(frmMainMenu.sG_CompName,'�@','��') + '@' + '�k�ݦ~��:' + sL_YM + '@' +
                      '�d�ߦ��O����:' + frmSO8B20_1.sG_ChargeNameSQL + '@' +
                      '�����p���h:' + sL_PayModeName;;

      if cdsSO122Channel.RecordCount > 0 then
      begin
        cdsSO122Channel.DisableControls;
        DataSetToXLS(cdsSO122Channel,sL_FileName,sL_QueryString,'INTACCID,INTACCNAME,INTACCCOMM,INTACCORIPERCENT,REALSTARTDATE,REALSTOPDATE,REALDATE,SHOULDDATE');
        cdsSO122Channel.EnableControls;
      end;
    end;

    //�^�_�쪬�A
    TUCommonFun.setDefaultCursor;

    if cdsSO122Channel.RecordCount > 0 then
      MessageDlg(sL_HaveDataMsg + sL_FileName,mtInformation,[mbOK],0)
    else
      MessageDlg(sL_NoDataMsg,mtInformation,[mbOK],0);
end;

function TdtmMain1.TransToBoxExcel(sI_Flag, sI_BelongYM,
  sI_CurrDateTime: String;bI_IncludeWorkerComm : Boolean): String;
Var
  sL_QueryString,sL_FileName,sL_FileCurrDateTime,sL_CurrDateTime : String;
  L_StrList : TStringList;
  sL_NoDataMsg,sL_HaveDataMsg : String;
  sL_YM,sL_IncludeWorkerComm : String;
begin
    sL_YM := IntToStr(StrToInt(Copy(sI_BelongYM,1,3)))+'�~'+Copy(sI_BelongYM,4,2)+'��';

    sL_FileCurrDateTime := sI_CurrDateTime;
    sL_CurrDateTime := ReplaceStr(Trim(ReplaceStr(sL_FileCurrDateTime,'/')),':');

    if bI_IncludeWorkerComm then
      sL_IncludeWorkerComm := '�]�tBOX�u�{�H������'
    else
      sL_IncludeWorkerComm := '���]�tBOX�u�{�H������';

    if sI_Flag = BOX_A then   //������Ҧ�����
    begin
      if TClientDataSet(cdsSO122BOX).IndexName = 'TmpIndex' then
        TClientDataSet(cdsSO122BOX).DeleteIndex('TmpIndex');
      TClientDataSet(cdsSO122BOX).AddIndex('TmpIndex', 'CustID', [ixCaseInsensitive],'','',0);
      TClientDataSet(cdsSO122BOX).IndexName := 'TmpIndex';

      sL_FileName := sG_ExcelSavePath +'\BOX�����Ҧ�����_' + frmMainMenu.sG_CompCode + '_'  + sI_BelongYM + '_' + sL_CurrDateTime + '.XLS';
      sL_HaveDataMsg := '�p�⧹��,BOX�����Ҧ����Ӧs��  ';
      sL_NoDataMsg := '�S��BOX�����Ҧ����Ӹ�� ';

      sL_QueryString:='����W��:' + sL_YM + '_�Ҧ�BOX�������Ӫ�' + '@' +
                      '�έp�ɶ�:' + sL_FileCurrDateTime + '@' + '����H��:' + frmMainMenu.sG_User+'@' +
                      '���q�W��:' + TUstr.replaceStr(frmMainMenu.sG_CompName,'�@','��') + ' (' + sL_IncludeWorkerComm + ')@' + '�k�ݦ~��:' + sL_YM + '@' +
                      '�d�ߦ��O����:' + frmSO8B20_1.sG_ChargeNameSQL;

      cdsSO122BOX.DisableControls;

      //���Ŀ�]�tBOX�u�{�H������
      if bI_IncludeWorkerComm then
        DataSetToXLS(cdsSO122BOX,sL_FileName,sL_QueryString,'BUYORRENT,REALSTARTDATE,REALSTOPDATE,REALDATE,SHOULDDATE')
      else
        DataSetToXLS(cdsSO122BOX,sL_FileName,sL_QueryString,'BUYORRENT,WORKEREN1,WORKEREN1NAME,WORKEREN1COMM,REALSTARTDATE,REALSTOPDATE,REALDATE,SHOULDDATE');

      cdsSO122BOX.EnableControls;
    end
    else if sI_Flag = BOX_I then  //��������ФH����
    begin
      if TClientDataSet(cdsSO122BOX).IndexName = 'TmpIndex' then
        TClientDataSet(cdsSO122BOX).DeleteIndex('TmpIndex');
      TClientDataSet(cdsSO122BOX).AddIndex('TmpIndex', 'INTACCID', [ixCaseInsensitive],'','',0);
      TClientDataSet(cdsSO122BOX).IndexName := 'TmpIndex';

      sL_FileName := sG_ExcelSavePath + '\BOX�������ФH����_' + frmMainMenu.sG_CompCode + '_' + sI_BelongYM + '_' + sL_CurrDateTime + '.XLS';
      sL_HaveDataMsg := '�p�⧹��,BOX�������ФH���Ӧs��  ';
      sL_NoDataMsg := '�S��BOX�������ФH���Ӹ�� ';

      sL_QueryString:='����W��:' + sL_YM + '_BOX�������ФH���Ӫ�' + '@' +
                      '�έp�ɶ�:' + sL_FileCurrDateTime + '@' + '����H��:' + frmMainMenu.sG_User+'@' +
                      '���q�W��:' + TUstr.replaceStr(frmMainMenu.sG_CompName,'�@','��') + '@' + '�k�ݦ~��:' + sL_YM + '@' +
                      '�d�ߦ��O����:' + frmSO8B20_1.sG_ChargeNameSQL;

      cdsSO122BOX.DisableControls;
      DataSetToXLS(cdsSO122BOX,sL_FileName,sL_QueryString,'WORKEREN1,WORKEREN1NAME,WORKEREN1COMM,REALSTARTDATE,REALSTOPDATE,REALDATE,SHOULDDATE');
      cdsSO122BOX.EnableControls;
    end
    else if sI_Flag = BOX_W then  //������u�{�H������
    begin
      if TClientDataSet(cdsSO122BOX).IndexName = 'TmpIndex' then
        TClientDataSet(cdsSO122BOX).DeleteIndex('TmpIndex');
      TClientDataSet(cdsSO122BOX).AddIndex('TmpIndex', 'WORKEREN1', [ixCaseInsensitive],'','',0);
      TClientDataSet(cdsSO122BOX).IndexName := 'TmpIndex';

      sL_FileName := sG_ExcelSavePath + '\BOX�����u�{�H������_' + frmMainMenu.sG_CompCode + '_' + sI_BelongYM + '_' + sL_CurrDateTime + '.XLS';
      sL_HaveDataMsg := '�p�⧹��,BOX�����u�{�H�����Ӧs��  ';
      sL_NoDataMsg := '�S��BOX�����u�{�H�����Ӹ�� ';


      sL_QueryString:='����W��:' + sL_YM + '_BOX�����u�{�H�����Ӫ�' + '@' +
                      '�έp�ɶ�:' + sL_FileCurrDateTime + '@' + '����H��:' + frmMainMenu.sG_User+'@' +
                      '���q�W��:' + TUstr.replaceStr(frmMainMenu.sG_CompName,'�@','��') + '@' + '�k�ݦ~��:' + sI_BelongYM + '@' +
                      '�d�ߦ��O����:' + frmSO8B20_1.sG_ChargeNameSQL;

      cdsSO122BOX.DisableControls;
      DataSetToXLS(cdsSO122BOX,sL_FileName,sL_QueryString,'INTACCID,INTACCNAME,INTACCCOMM,REALSTARTDATE,REALSTOPDATE,REALDATE,SHOULDDATE');
      cdsSO122BOX.EnableControls;
    end;

    //�^�_�쪬�A
    TUCommonFun.setDefaultCursor;

    if cdsSO122BOX.RecordCount > 0 then
      MessageDlg(sL_HaveDataMsg + sL_FileName,mtInformation,[mbOK],0)
    else
      MessageDlg(sL_NoDataMsg,mtInformation,[mbOK],0);

end;




function TdtmMain1.processOnceIncome(sI_OrderNo, sI_RealDate: String;
  I_DataSet: TDataSet): Boolean;
var
    sL_CustId,sL_CompCode,sL_CitemCode,sL_ClctId,sL_ClctName,sL_SQL,sL_MediaCode : String;
    sL_AcceptEn,sL_AcceptName,sL_IntroId,sL_IntroName,sL_PromCode,sL_PromName : String;
    sL_RefNo,sL_MediaName,sL_RealStartDate,sL_RealStopDate : String;
    bL_PayByCreditCard,bL_HaveClctId,bL_SelfCallIn,bL_HasPayComm : Boolean;
    fL_ClctEnOriPercent,fL_IntAccOriPercent,fL_RealAmt : Double;
    nL_RealPeriod : Integer;
    dL_RealStartDate,dL_RealStopDate,nL_TotalPeriod : TDate;
begin
    //�B�z�@�����I�W�D����
    sL_CustId := I_DataSet.fieldByName('CUSTID').AsString;
    sL_CompCode := I_DataSet.fieldByName('COMPCODE').AsString;
    sL_CitemCode := I_DataSet.FieldByName('CitemCode').AsString;
    sL_ClctId := I_DataSet.FieldByName('ClctEn').AsString;
    sL_ClctName := I_DataSet.FieldByName('ClctName').AsString;


    //�P�_�O�_���H�Υd=>�YSo033/So034.PTCode=3,��ܬ��H�Υd�I�O
    bL_PayByCreditCard := ifPayByCreditCard(I_DataSet.FieldByName('PTCode').AsString);

    //���ˬd���L���O�H��,�S���h���p��
    //���O��( �Y�O�H�Υd���O�BSO033�S�����O�H��,�h��SO132�� )
    if (bL_PayByCreditCard = true) and (sL_ClctId = '') then
    begin
      //�� SO132 ��H�Υd�I�O�����k�ݤH��
      bL_HaveClctId := getEmpData(sL_CompCode,sL_CustID,sL_CitemCode,sL_ClctId,sL_ClctName);
    end
    else if sL_ClctId <> '' then
      bL_HaveClctId := true
    else
      bL_HaveClctId := false;


    if bL_HaveClctId then    //�����O�H��
    begin
      sL_SQL:='Select MediaCode,ACCEPTEN, ACCEPTNAME, IntroId, IntroName,PromCode,PromName ';
      sL_SQL:=sL_SQL+' from So105 ';
      sL_SQL:=sL_SQL+' where OrderNo='''+sI_OrderNo+'''';
  //    saveToFile(sL_SQL,'C:\temp\CallInCheck_SQL.txt');

      with qryCom do
      begin
        Close;
        SQL.Clear;
        SQL.Add(sL_SQL);
        Open;

        if RecordCount=0 then
        begin
          Close;
          Exit;
        end
        else
        begin
          sL_MediaCode := FieldByName('MediaCode').AsString;
          sL_AcceptEn := FieldByName('ACCEPTEN').AsString;
          sL_AcceptName := FieldByName('ACCEPTNAME').AsString;
          sL_IntroId := FieldByName('IntroId').AsString;
          sL_IntroName := FieldByName('IntroName').AsString;

          //�~�Ȭ��ʸ��
          sL_PromCode := FieldByName('PromCode').AsString;
          sL_PromName := FieldByName('PromName').AsString;

          bL_SelfCallIn := true;

          sL_RefNo := '';
          sL_MediaName := '';

          if sL_MediaCode <> '' then
          begin
            if qryCodeCD009.Locate('CODENO', StrToInt(sL_MediaCode),[]) then
            begin
              sL_RefNo := qryCodeCD009.FieldByName('RefNo').AsString;
              sL_MediaName := qryCodeCD009.FieldByName('DESCRIPTION').AsString;

              //���p���O�ۨӹq
              if qryCodeCD009.FieldByName('RefNo').AsInteger in [1,2,3,5] then
                bL_SelfCallIn := false;
            end;
          end;
        end;
        Close;
      end;

      sL_RealStartDate := qrySo033So034.FieldByName('RealStartDate').AsString;
      sL_RealStopDate := qrySo033So034.FieldByName('RealStopDate').AsString;

//*****************************************************************************
//*****************************************************************************
      fL_ClctEnOriPercent := 0;
      fL_IntAccOriPercent := 0;


      fG_SpreadComm:=0;
      fG_ChargeComm:=0;

      nL_RealPeriod := qrySo033So034.FieldByName('RealPeriod').AsInteger;
      fL_RealAmt := qrySo033So034.FieldByName('RealAmt').AsFloat;
      dL_RealStartDate := qrySo033So034.FieldByName('RealStartDate').AsDateTime;
      dL_RealStopDate := qrySo033So034.FieldByName('RealStopDate').AsDateTime;

      if (qrySo033So034.FieldByName('FirstFlag').AsString='1') then//�O����
      begin
        //���s�H�����������W�L�������
        if nL_RealPeriod > nG_LimitPeriod then
          nL_TotalPeriod := nL_RealPeriod
        else
          nL_TotalPeriod := nG_LimitPeriod;


        //�p����ˤH������
        fL_ClctEnOriPercent := 0;
        fL_IntAccOriPercent := 0;

        if bL_PayByCreditCard then//�H�H�Υd�I�O
        begin
          if sG_PayUnit='2' then    //��
          begin
            //�ˬd���s�H�O�_���ۤv���u,�O���ܤ~�o����
            if sL_RefNo <> '1' then
            begin
              //�|����Τ��p��,�]�O���_����h�O,�ҥH�����@������(�Ĥ@���N�@������)
              fG_SpreadComm := fG_FirCardSpd;    //���s�H
              {
              if (ii = 1) AND (nL_FirstPeriodDays > nG_ChannelViewDays) then
                fG_SpreadComm := fG_FirCardSpd    //���s�H
              else
                fG_SpreadComm := 0;
              }
            end
            else
              fG_SpreadComm := 0;

            //���O������
            fG_ChargeComm := fG_FirCardCharge;

          end
          else if sG_PayUnit='1' then  //�ʤ���
          begin
            //�ˬd���s�H�O�_���ۤv���u,�O���ܤ~�o����
            if sL_RefNo <> '1' then
            begin
              //�@�����I����=(RealAmt * % ) * �̦h�i����� / �`����
              fG_SpreadComm := (fL_RealAmt*(fG_FirCardSpd/100))*nG_LimitPeriod/nL_TotalPeriod //0.12;      //���s�H
              {
              if (ii = 1) AND (nL_FirstPeriodDays > nG_ChannelViewDays) then   //�Ĥ@������=(RealAmt * % ) * �����p�O�Ѽ� / �`���ƤѼ�
                fG_SpreadComm := (fL_RealAmt*(fG_FirCardSpd/100))*nL_FirstPeriodDays/nL_TotalPeriodDays //0.12;      //���s�H
              else            //�D�Ĥ@������=(RealAmt * % ) * ���Ѽ� / �`���ƤѼ�
                fG_SpreadComm := (fL_RealAmt*(fG_FirCardSpd/100))*nL_OtherPeriodDays/nL_TotalPeriodDays //0.12;      //���s�H
              }
            end
            else
              fG_SpreadComm := 0;


            //���O���u�p�⭺��
            fG_ChargeComm := (fL_RealAmt*(fG_FirCardCharge/100));  //0.01;      //���O��

            //�����p��ʤ���@���h�O�ɨ̾�
            fL_ClctEnOriPercent := fG_FirCardCharge;
            fL_IntAccOriPercent := fG_FirCardSpd;
          end;
        end
        else//���O�H�H�Υd�I�O
        begin
          if sG_PayUnit='2' then    //��
          begin
            //�ˬd���s�H�O�_���ۤv���u,�O���ܤ~�o����
            if sL_RefNo <> '1' then
            begin
              //�|����Τ��p��,�]�O���_����h�O,�ҥH�����@������(�Ĥ@���N�@������)
              fG_SpreadComm := fG_FirNoCardSpd;  //���s�H
              {
              if (ii = 1) AND (nL_FirstPeriodDays > nG_ChannelViewDays) then
                fG_SpreadComm := fG_FirNoCardSpd  //���s�H
              else
                fG_SpreadComm := 0;
              }
            end
            else
              fG_SpreadComm := 0;

            //���O������
            fG_ChargeComm := fG_FirNoCardCharge;

          end
          else if sG_PayUnit='1' then //�ʤ���
          begin
            //�ˬd���s�H�O�_���ۤv���u,�O���ܤ~�o����
            if sL_RefNo <> '1' then
              //�@�����I����=(RealAmt * % ) * �̦h�i����� / �`����
              fG_SpreadComm := (fL_RealAmt*(fG_FirNoCardSpd/100))*nG_LimitPeriod/nL_TotalPeriod//0.12;  //���s�H
              {
              if (ii = 1) AND (nL_FirstPeriodDays > nG_ChannelViewDays) then //�Ĥ@������=(RealAmt * % ) * �����p�O�Ѽ� / �`���ƤѼ�
                fG_SpreadComm := (fL_RealAmt*(fG_FirNoCardSpd/100))*nL_FirstPeriodDays/nL_TotalPeriodDays//0.12;  //���s�H
              else
                fG_SpreadComm := (fL_RealAmt*(fG_FirNoCardSpd/100))*nL_OtherPeriodDays/nL_TotalPeriodDays//0.12;  //���s�H
                }
            else
              fG_SpreadComm := 0;

            //���O���u�p�⭺��
            fG_ChargeComm:=(fL_RealAmt*(fG_FirNoCardCharge/100));  //0.03; //���O��

            //�����p��ʤ���@���h�O�ɨ̾�
            fL_ClctEnOriPercent := fG_FirNoCardCharge;
            fL_IntAccOriPercent := fG_FirNoCardSpd;
          end;
        end;

        if not bL_SelfCallIn then//���O�ۨӹq
          SaveTempDataTocdsSo134(sL_PromCode,sL_PromName,sL_ClctId, sL_ClctName, sL_IntroId, sL_IntroName,sL_MediaCode,sL_MediaName,sL_RefNo, I_DataSet,dL_RealStartDate,dL_RealStopDate,fL_ClctEnOriPercent,fL_IntAccOriPercent);
      end
      else//�O��
      begin
        //�򦬥u�p�⦬�O�H��
        fG_SpreadComm := 0;

        //sL_beforeMonth := beforeMonth(sG_BelongYM,1);
        //nL_OtherPeriodDays := TUdateTime.DaysOfMonth(sL_beforeMonth + '01');

        if bL_PayByCreditCard then    //�H�H�Υd�I�O
        begin
          if sG_PayUnit='2' then      //��
            fG_ChargeComm := fG_ContCardCharge //���O��
          else if sG_PayUnit='1' then //�ʤ���
            fG_ChargeComm := (fL_RealAmt*(fG_ContCardCharge/100));//0.01 //���O��


          //�����p��ʤ���@���h�O�ɨ̾�
          fL_ClctEnOriPercent := fG_ContCardCharge;
          fL_IntAccOriPercent := 0;
        end
        else
        begin                         //���H�H�Υd�I�O
          if sG_PayUnit='2' then //��
            fG_ChargeComm := fG_ContNoCardCharge    //���O��
          else if sG_PayUnit='1' then //�ʤ���
            fG_ChargeComm := (fL_RealAmt*(fG_ContNoCardCharge/100));//0.03;    //���O��


          //�����p��ʤ���@���h�O�ɨ̾�
          fL_ClctEnOriPercent := fG_ContNoCardCharge;
          fL_IntAccOriPercent := 0;

        end;

        if not bL_SelfCallIn then//���O�ۨӹq
          SaveTempDataTocdsSo134(sL_PromCode,sL_PromName,sL_ClctId,sL_ClctName,sL_IntroId,sL_IntroName,sL_MediaCode,sL_MediaName,sL_RefNo,I_DataSet,dL_RealStartDate,dL_RealStopDate,fL_ClctEnOriPercent,fL_IntAccOriPercent);
      end;


      if (bL_SelfCallIn) and (qrySo033So034.FieldByName('FirstFlag').AsString='1') then //�O�ۨӹq,�B�O����
      begin
        if fL_RealAmt <> 0 then    //�ꦬ���B����0�����έp��
        begin
          bL_HasPayComm := CustOnceIDCheck(sL_CustId,sL_CompCode); //�ˬd���ФH�O�_�w�g���L����

          try
            if not bL_HasPayComm then
              bL_HasPayComm := insertSO135(sL_CompCode,sL_CustID, sL_AcceptEn,sL_AcceptName);

            insertCallInDataToSo134(bL_HasPayComm,sL_PromCode,sL_PromName,sL_CompCode,sG_OnceComputeYM,sL_AcceptEn,sL_AcceptName,sL_ClctId,sL_ClctName, I_DataSet,fL_IntAccOriPercent);
          except
            ShowMessage('�ۨӹq�����B�z����_2!');
            Result:=False;
            Exit;
          end;
        end;
      end;
    end
    else
    begin  //�S�����O�H��
      cdsSo033Log.Append;
      cdsSo033Log.FieldByName('CUSTID').AsString := qrySo033So034.FieldByName('CUSTID').AsString;
      cdsSo033Log.FieldByName('BILLNO').AsString := qrySo033So034.FieldByName('BILLNO').AsString;
      cdsSo033Log.FieldByName('ITEM').AsString := qrySo033So034.FieldByName('ITEM').AsString;
      cdsSo033Log.FieldByName('CITEMCODE').AsString := qrySo033So034.FieldByName('CITEMCODE').AsString;
      cdsSo033Log.FieldByName('CITEMNAME').AsString := qrySo033So034.FieldByName('CITEMNAME').AsString;
      cdsSo033Log.FieldByName('SHOULDDATE').AsString := qrySo033So034.FieldByName('SHOULDDATE').AsString;
      cdsSo033Log.FieldByName('REALDATE').AsString := qrySo033So034.FieldByName('REALDATE').AsString;
      cdsSo033Log.FieldByName('REALAMT').AsString := qrySo033So034.FieldByName('REALAMT').AsString;
      cdsSo033Log.FieldByName('REALPERIOD').AsString := qrySo033So034.FieldByName('RealPeriod').AsString;
      cdsSo033Log.FieldByName('REALSTARTDATE').AsString := qrySo033So034.FieldByName('REALSTARTDATE').AsString;
      cdsSo033Log.FieldByName('SBILLNO').AsString := qrySo033So034.FieldByName('SBILLNO').AsString;
      cdsSo033Log.FieldByName('SITEM').AsInteger := qrySo033So034.FieldByName('SITEM').AsInteger;
      cdsSo033Log.FieldByName('Notes').AsString := '�@�����O-�S�����O�H�����';
      cdsSo033Log.Post;
    end;

end;

procedure TdtmMain1.SaveTempDataTocdsSo134(sI_PromCode, sI_PromName,
  sI_ClctId, sI_ClctName, sI_IntroId, sI_IntroName, sI_MediaCode,
  sI_MediaName, sI_RefNo: STring; I_DataSet: TDataset; dI_RealStartDate, dI_RealStopDate: TDate;
  fI_ClctEnOriPercent, fI_IntAccOriPercent: Double);
var
   dL_ShouldDate,dL_RealDate : TDate;
   sL_CompCode,sL_CustID,sL_CitemCode : String;
begin
    try
      //if ((fG_SpreadComm>0) and (sI_IntroId<>'')) or ((fG_ChargeComm>0) and (sI_ClctId<>'')) then
      if (sI_IntroId<>'') or (sI_ClctId<>'') then
      begin
        with cdsSO134 do
        begin
          Append;
          sL_CompCode := I_DataSet.FieldByName('CompCode').AsString;
          FieldByName('CompCode').AsString:= sL_CompCode;
          FieldByName('ComputeYM').AsString:=sG_OnceComputeYM;
          FieldByName('BelongYM').AsString:= sG_OnceBelongYM;
          FieldByName('BillNo').AsString:=I_DataSet.FieldByName('BillNo').AsString;
          FieldByName('Item').AsString:=I_DataSet.FieldByName('Item').AsString;
          FieldByName('PeriodId').AsInteger:=I_DataSet.FieldByName('REALPERIOD').AsInteger;
          FieldByName('FirstFlag').AsString:=I_DataSet.FieldByName('FirstFlag').AsString;

          FieldByName('RealAmt').AsFloat:=I_DataSet.FieldByName('RealAmt').AsFloat;
          if sI_IntroId<>'' then//���s�H
          begin
            FieldByName('IntAccId').AsString:=sI_IntroId;
            FieldByName('IntAccName').AsString:=sI_IntroName;
            FieldByName('IntAccComm').AsFloat:=RoundTo(fG_SpreadComm,-2);
            FieldByName('IntAccOriPercent').AsFloat:= RoundTo(fI_IntAccOriPercent,-2);
            FieldByName('StaffCode').AsString:='3';
          end;

          //���O��
          if sI_ClctId<>'' then
          begin
            FieldByName('CLCTEN').AsString:=sI_ClctId;
            FieldByName('CLCTENNAME').AsString:=sI_ClctName;
            FieldByName('CLCTENCOMM').AsFloat:=RoundTo(fG_ChargeComm,-2);
            FieldByName('ClctEnOriPercent').AsFloat:=RoundTo(fI_ClctEnOriPercent,-2);
          end;



          FieldByName('RealStartDate').AsDateTime := dI_RealStartDate;
          FieldByName('RealStopDate').AsDateTime := dI_RealStopDate;

          FieldByName('OPERATOR').AsString:=frmMainMenu.sG_User;
          //FieldByName('UpdTime').AsString:=IntToStr(StrToInt(FormatDateTime('YYYY',now))-1911)+ FormatDateTime('/MM/DD hh:mm:ss',now);
          FieldByName('UpdTime').AsString := sG_CurDate;

          //CD009 �� RefNo=2(���u����),RefNo=3(�P���I)
          //�YRefno=2,�h����쬰SO CompCode,�_�h��MediaCode
          if sI_RefNo='2' then
          begin
            FieldByName('MediaCode').AsInteger := StrToInt(frmMainMenu.sG_CompCode);
            FieldByName('MediaName').AsString := frmMainMenu.sG_CompName;
          end
          else if (sI_RefNo='3') or (sI_RefNo='5') then
          begin
            if sI_MediaCode <> '' then
              FieldByName('MediaCode').AsInteger := StrToInt(sI_MediaCode);

            FieldByName('MediaName').AsString := sI_MediaName;
          end;

          sL_CitemCode := I_DataSet.FieldByName('CitemCode').AsString;
          if sL_CitemCode <> '' then
            FieldByName('CitemCode').AsInteger := StrToInt(sL_CitemCode);
          FieldByName('CitemName').AsString := I_DataSet.FieldByName('CitemName').AsString;


          if sI_PromCode <> '' then
            FieldByName('PromCode').AsInteger := StrToInt(sI_PromCode);

          FieldByName('PromName').AsString := sI_PromName;


          dL_RealDate := I_DataSet.FieldByName('RealDate').AsDateTime;
          if dL_RealDate <> 0 then
            FieldByName('RealDate').AsDateTime := dL_RealDate;

          dL_ShouldDate := I_DataSet.FieldByName('ShouldDate').AsDateTime;
          if dL_ShouldDate <> 0 then
            FieldByName('ShouldDate').AsDateTime := dL_ShouldDate;

          sL_CustID := I_DataSet.FieldByName('CustID').AsString;
          if sL_CustID <> '' then
            FieldByName('CustID').AsString := sL_CustID;
          Post;
        end;
      end;
    except
      ShowMessage('cdsSo134�s�ɥ���');
      Exit;
    end;


end;

function TdtmMain1.callSF_GetChargeData(sI_Table, sI_CompCode, sI_BelongYM,
  sI_BillNo, sI_Item: String;var sI_Msg : String;var nI_RetCode : Integer) : Boolean;
var
    sL_Msg : String;
begin
    with spGetChargeData do
    begin
      Parameters.ParamByName('p_Table').Value := sI_Table;
      Parameters.ParamByName('p_CompCode').Value := sI_CompCode;
      Parameters.ParamByName('p_BelongYM').Value := sI_BelongYM;
      Parameters.ParamByName('p_BillNo').Value := sI_BillNo;
      Parameters.ParamByName('p_Item').Value := sI_Item;
      ExecProc;

      nI_RetCode := Parameters.Items[6].Value;
      sL_Msg := Parameters.Items[7].Value;

      sI_Msg := sL_Msg;

      //�p�G�Ǧ^ 1 �N�����
      if nI_RetCode = 1 then
        Result := true
      else
        Result := false;
    end;


{
    with cdsGetChargeData do
    begin
  //  SHOWMESSAGE(INTTOSTR(Params.Count));
      Params.Clear;
      FetchParams;

      Params.ParamByName('p_Table').AsString := sI_Table;
      Params.ParamByName('p_CompCode').AsString := sI_CompCode;
      Params.ParamByName('p_BelongYM').AsString := sI_BelongYM;
      Params.ParamByName('p_BillNo').AsString := sI_BillNo;
      Params.ParamByName('p_Item').AsString := sI_Item;
      Execute;


      sI_Msg := Params.ParamByName('p_RetMsg').AsString;
      nI_RetCode := Params.ParamByName('p_RetCode').AsInteger;
      //nI_RetCode := StrToInt(FloatToStr(fL_RetCode));
    end;
}
end;

function TdtmMain1.handleOnceOutCome(sI_TableName, sI_YM: String): String;
var
    sL_SQL,sL_OrderNo,sL_TempString1,sL_TempString2 : String;
    sL_SBillNo,sL_SItem,sL_FunctionNo : String;
    sL_BelongDate,sL_Msg : String;
    dL_BelongDate,dL_RealStartDate,dL_RealStopDate,dL_ShouldDate,dL_ViewStartDate,dL_ViewStopDate : TDate;
    fL_RealAmt : Double;
    nL_DiffDays,nL_RetCode : Integer;
    bL_HaveData : Boolean;

    sL_BillNo,sL_Item,sL_Period,sL_IntAccId,sL_IntAccName,sL_StaffCode : String;
    fL_IntAccComm,fL_IntAccOriPercent : Double;
    sL_MediaCode,sL_MediaName,sL_CitemCode,sL_CitemName : String;
    sL_PromCode,sL_PromName,sL_CustID,sL_FirstFlag : String;
    dL_RealDate : TDate;
    L_StrList : TStringList;
    nL_RealPeriod,nL_TotalPeriod : Integer;
begin
    //�ȳB�z���ФH/���z�H�h�O,�]���ФH�������@���N�u��U��(���O�H����C��|�X)
    sL_BelongDate := Copy(sG_BelongYM,1,3) + '/' + Copy(sG_BelongYM,4,5) + '/01';
    dL_BelongDate := StrToDate(chineseDateChangeToEnglishDate(sL_BelongDate));

//����jackal
//dL_BelongDate := StrToDate('2003/03/01');
////////

    sL_SQL:='SELECT A.RefNo,B.CustId,B.OrderNo,B.FirstFlag,TO_DATE(TO_CHAR(B.SHOULDDATE,'+''''+'YYYY/MM/DD'+''''+'),'+''''+'YYYY/MM/DD'+''''+') SHOULDDATE,';
    sL_SQL:=sL_SQL+'B.DiscountCode,B.OrderNo,TO_DATE(TO_CHAR(B.RealStartDate,'+''''+'YYYY/MM/DD'+''''+'),'+''''+'YYYY/MM/DD'+''''+') RealStartDate,TO_DATE(TO_CHAR(B.RealStopDate,'+''''+'YYYY/MM/DD'+''''+'),'+''''+'YYYY/MM/DD'+''''+') RealStopDate,';
    sL_SQL:=sL_SQL+'B.CMName,B.CitemCode,B.CitemName,B.PTCode,TO_DATE(TO_CHAR(B.RealDate,'+''''+'YYYY/MM/DD'+''''+'),'+''''+'YYYY/MM/DD'+''''+') RealDate,B.RealAmt,to_char(B.RealAmt) StrRealAmt,';
    sL_SQL:=sL_SQL+'B.CompCode,B.ClctEn,B.ClctName,B.BillNo,';
    sL_SQL:=sL_SQL+'B.Item, B.REALPERIOD,B.SBILLNO, B.SITEM FROM CD019 A, '+sI_TableName+'';
    sL_SQL:=sL_SQL+' B Where B.CancelFlag=0';
    sL_SQL:=sL_SQL+' AND B.UcCode is NULL and B.ServiceType=''' + frmMainMenu.sG_ServiceType + '''';

    sL_SQL:=sL_SQL+' AND B.ShouldDate < ' + TUstr.getOracleSQLDateStr(dL_BelongDate);
    sL_SQL:=sL_SQL+' AND B.CitemCode=A.CodeNo ';
    sL_SQL:=sL_SQL+' AND B.COMPCODE='''+frmMainMenu.sG_CompCode+''' AND A.RefNo=9 ';
    //Jackal_Test_920520
    //sL_SQL:=sL_SQL+' AND B.SBILLNO = ''200302IC0177029''';

    sL_SQL:=sL_SQL+' Order by B.ClctEn';

    qrySo033So034.Close;
    with cdsSo033So034 do
    begin
      Close;
      CommandText := sL_SQL;
      OPEN;
    end;

    //���t��
    transSo033So034Notation(cdsSo033So034);

    L_StrList := TStringList.Create;

    cdsSo033So034.First;
    while not cdsSo033So034.Eof do
    begin
      sL_SBillNo := cdsSo033So034.FieldByName('SBillNo').AsString;
      sL_SItem := cdsSo033So034.FieldByName('SItem').AsString;

      sL_BillNo := cdsSo033So034.FieldByName('BillNo').AsString;
      sL_Item := cdsSo033So034.FieldByName('Item').AsString;
      //sL_Period := cdsSo033So034.FieldByName('RealPeriod').AsString;

      dL_ShouldDate := cdsSo033So034.FieldByName('ShouldDate').AsDateTime;
      fL_RealAmt := cdsSo033So034.FieldByName('RealAmt').AsFloat;

      //�Y�h�O��Ʀ� RealStartDate �� RealStopDate �h�Φۨ���RealStartDate �� RealStopDate
      //�Y�h�O��ƨS�� RealStartDate �� RealStopDate �h�Υ�����RealStartDate �� RealStopDate
      if (dL_RealStartDate =0) AND  (dL_RealStopDate =0) then //�h�O��ƨS�� RealStartDate �� RealStopDate
      begin
        if (sL_SBillNo<>'') and (sL_SItem<>'')
        and (cdsSo134.Locate('COMPCODE;BELONGYM;BILLNO;ITEM',VarArrayOf([frmMainMenu.sG_CompCode,sG_BelongYM,sL_SBillNo,sL_SItem]),[])) then
        begin
          sL_IntAccId := cdsSo134.FieldByName('IntAccId').AsString;
          sL_StaffCode := cdsSo134.FieldByName('StaffCode').AsString;
          fL_IntAccOriPercent := cdsSo134.FieldByName('IntAccOriPercent').AsFloat;

          sL_Period := cdsSo134.FieldByName('PeriodId').AsString;

          sL_MediaCode := cdsSo134.FieldByName('MediaCode').AsString;
          sL_MediaName := cdsSo134.FieldByName('MediaName').AsString;
          sL_CitemCode := cdsSo134.FieldByName('CitemCode').AsString;
          sL_CitemName := cdsSo134.FieldByName('CitemName').AsString;

          sL_PromCode := cdsSo134.FieldByName('PromCode').AsString;
          sL_PromName := cdsSo134.FieldByName('PromName').AsString;
          sL_CustID := cdsSo134.FieldByName('CustID').AsString;

          //�h�O�� RealStartDate �S���ȫh��, �h�O�� ShouldDate �N��h�O���}�l���
          dL_RealStartDate := dL_ShouldDate;
          dL_RealStopDate := cdsSo134.FieldByName('RealStopDate').AsDateTime;

          bL_HaveData := true;
        end
        else
        begin
          bL_HaveData := callSF_GetChargeData('SO134', frmMainMenu.sG_CompCode, sG_OnceBelongYM,
                               sL_SBillNo, sL_SItem,sL_Msg,nL_RetCode);
          if bL_HaveData then
          begin
            L_StrList.Clear;
            L_StrList := TUstr.ParseStrings(sL_Msg,',');
            {�Ǧ^�ƾڶ���
             IntAccId,IntAccName,IntAccOriPercent,
             StaffCode,RealStartDate,RealStopDate,CitemCode,CitemName,
             MediaCode,MediaName,PromCode,PromName,CustId
             }


            if L_StrList.Strings[0] <> 'NULL' then
              sL_IntAccId := L_StrList.Strings[0];

            if L_StrList.Strings[1] <> 'NULL' then
              sL_IntAccName := L_StrList.Strings[1];

            if L_StrList.Strings[2] <> 'NULL' then
              fL_IntAccOriPercent := StrToFloat(L_StrList.Strings[2]);

            if L_StrList.Strings[3] <> 'NULL' then
              sL_StaffCode := L_StrList.Strings[3];

            //�h�O�� RealStartDate �S���ȫh��, �h�O�� ShouldDate �N��h�O���}�l���
            dL_RealStartDate := dL_ShouldDate;

            if L_StrList.Strings[5] <> 'NULL' then
              dL_RealStopDate := StrToDate(L_StrList.Strings[5]);

            if L_StrList.Strings[6] <> 'NULL' then
              sL_CitemCode := L_StrList.Strings[6];

            if L_StrList.Strings[7] <> 'NULL' then
              sL_CitemName := L_StrList.Strings[7];

            if L_StrList.Strings[8] <> 'NULL' then
              sL_MediaCode := L_StrList.Strings[8];

            if L_StrList.Strings[9] <> 'NULL' then
              sL_MediaName := L_StrList.Strings[9];

            if L_StrList.Strings[10] <> 'NULL' then
              sL_PromCode := L_StrList.Strings[10];

            if L_StrList.Strings[11] <> 'NULL' then
              sL_PromName := L_StrList.Strings[11];

            if L_StrList.Strings[12] <> 'NULL' then
              sL_CustId := L_StrList.Strings[12];

            if L_StrList.Strings[13] <> 'NULL' then
              sL_Period := L_StrList.Strings[13];

            if L_StrList.Strings[14] <> 'NULL' then
              sL_FirstFlag := L_StrList.Strings[14];
          end;

        end;

      end
      else if (dL_RealStartDate<>0) and (dL_RealStopDate<>0) then //�h�O��Ʀ� RealStartDate �� RealStopDate
      begin
        if (sL_SBillNo<>'') and (sL_SItem<>'')
        and (cdsSo134.Locate('COMPCODE;BELONGYM;BILLNO;ITEM',VarArrayOf([frmMainMenu.sG_CompCode,sG_BelongYM,sL_SBillNo,sL_SItem]),[])) then
        begin
          sL_IntAccId := cdsSo134.FieldByName('IntAccId').AsString;
          sL_IntAccName := cdsSo134.FieldByName('IntAccName').AsString;
          sL_StaffCode := cdsSo134.FieldByName('StaffCode').AsString;
          fL_IntAccOriPercent := cdsSo134.FieldByName('IntAccOriPercent').AsFloat;

          sL_Period := cdsSo134.FieldByName('PeriodId').AsString;

          sL_MediaCode := cdsSo134.FieldByName('MediaCode').AsString;
          sL_MediaName := cdsSo134.FieldByName('MediaName').AsString;
          sL_CitemCode := cdsSo134.FieldByName('CitemCode').AsString;
          sL_CitemName := cdsSo134.FieldByName('CitemName').AsString;

          sL_PromCode := cdsSo134.FieldByName('PromCode').AsString;
          sL_PromName := cdsSo134.FieldByName('PromName').AsString;
          sL_CustID := cdsSo134.FieldByName('CustID').AsString;


          //�h�O�� RealStartDate ���ȫh��, �h�O�� RealStartDate �N��h�O���}�l���
          dL_RealStartDate := cdsSo033So034.FieldByName('RealStartDate').AsDateTime;
          dL_RealStopDate := cdsSo134.FieldByName('RealStopDate').AsDateTime;

          bL_HaveData := true;
        end
        else
        begin
          bL_HaveData := callSF_GetChargeData('SO134', frmMainMenu.sG_CompCode, sG_OnceBelongYM,
                               sL_SBillNo, sL_SItem,sL_Msg,nL_RetCode);
          if bL_HaveData then
          begin
            L_StrList.Clear;
            L_StrList := TUstr.ParseStrings(sL_Msg,',');

            {�Ǧ^�ƾڶ���
             IntAccId,IntAccName,IntAccComm,IntAccOriPercent,
             StaffCode,RealStartDate,RealStopDate,CitemCode,CitemName,
             MediaCode,MediaName,PromCode,PromName,CustId
            }

            if L_StrList.Strings[0] <> 'NULL' then
              sL_IntAccId := L_StrList.Strings[0];

            if L_StrList.Strings[1] <> 'NULL' then
              sL_IntAccName := L_StrList.Strings[1];

            if L_StrList.Strings[2] <> 'NULL' then
              fL_IntAccOriPercent := StrToFloat(L_StrList.Strings[2]);

            if L_StrList.Strings[3] <> 'NULL' then
              sL_StaffCode := L_StrList.Strings[3];

            //�h�O�� RealStartDate ���ȫh��, �h�O�� RealStartDate �N��h�O���}�l���
            dL_RealStartDate := cdsSo033So034.FieldByName('RealStartDate').AsDateTime;

            if L_StrList.Strings[5] <> 'NULL' then
              dL_RealStopDate := StrToDate(L_StrList.Strings[5]);

            if L_StrList.Strings[6] <> 'NULL' then
              sL_CitemCode := L_StrList.Strings[6];

            if L_StrList.Strings[7] <> 'NULL' then
              sL_CitemName := L_StrList.Strings[7];

            if L_StrList.Strings[8] <> 'NULL' then
              sL_MediaCode := L_StrList.Strings[8];

            if L_StrList.Strings[9] <> 'NULL' then
              sL_MediaName := L_StrList.Strings[9];

            if L_StrList.Strings[10] <> 'NULL' then
              sL_PromCode := L_StrList.Strings[10];

            if L_StrList.Strings[11] <> 'NULL' then
              sL_PromName := L_StrList.Strings[11];

            if L_StrList.Strings[12] <> 'NULL' then
              sL_CustId := L_StrList.Strings[12];

            if L_StrList.Strings[13] <> 'NULL' then
              sL_Period := L_StrList.Strings[13];              

            if L_StrList.Strings[14] <> 'NULL' then
              sL_FirstFlag := L_StrList.Strings[14];
          end;

        end;
      end;

           
      if (bL_HaveData=true) AND (sL_FirstFlag='1')then
      begin
        if sL_Period <> '' then
          nL_RealPeriod := StrToInt(sL_Period);
        //���s�H�����������W�L�������
        if nL_RealPeriod > nG_LimitPeriod then
          nL_TotalPeriod := nL_RealPeriod
        else
          nL_TotalPeriod := nG_LimitPeriod;

        with cdsSo134 do
        begin
          Append;
          FieldByName('COMPCODE').AsString := frmMainMenu.sG_CompCode;
          FieldByName('COMPUTEYM').AsString := sG_OnceComputeYM;
          FieldByName('BELONGYM').AsString := sG_OnceBelongYM;

          FieldByName('BILLNO').AsString := sL_BillNo;
          if sL_Item <> '' then
            FieldByName('ITEM').AsInteger := StrToInt(sL_Item);


          FieldByName('PERIODID').AsInteger := nL_RealPeriod;
          //FieldByName('WORKEREN1').Value := NULL;
          //FieldByName('WORKEREN1NAME').Value := NULL;
          //FieldByName('WORKEREN1COMM').AsInteger := StrToInt(sL_WorkerEn1Comm);

          //I_CDS.FieldByName('CLCTEN').Value := NULL;
          //I_CDS.FieldByName('CLCTENNAME').Value := NULL;
          //I_CDS.FieldByName('CLCTENCOMM').AsInteger := StrToInt(sL_ClctEnComm);
          //FieldByName('ClctEnOriPercent').AsFloat := fL_ClctEnOriPercent;

          FieldByName('INTACCID').AsString := sL_IntAccId;
          FieldByName('INTACCNAME').AsString := sL_IntAccName;
          FieldByName('IntAccOriPercent').AsFloat := fL_IntAccOriPercent;

          //�h�O���B = ��ڰh�O���B * %
          fL_IntAccComm := fL_RealAmt * fL_IntAccOriPercent/100;
          FieldByName('INTACCCOMM').AsFloat := fL_IntAccComm;

          FieldByName('STAFFCODE').AsString := sL_StaffCode;

          if dL_RealStartDate <> 0 then
            FieldByName('RealStartDate').AsDateTime := dL_RealStartDate;

          if dL_RealStopDate <> 0 then
            FieldByName('RealStopDate').AsDateTime := dL_RealStopDate;

          FieldByName('OPERATOR').AsString := frmMainMenu.sG_User;

          FieldByName('UPDTIME').AsString := sG_CurDate;
          FieldByName('MEDIACODE').AsString := sL_MediaCode;
          FieldByName('MediaName').AsString := sL_MediaName;

          FieldByName('CitemCode').AsString := sL_CitemCode;
          FieldByName('CitemName').AsString := sL_CitemName;
          FieldByName('RealAmt').AsFloat := fL_RealAmt;

          if sL_PromCode <> '' then
            FieldByName('PromCode').AsInteger := StrToInt(sL_PromCode);
          FieldByName('PromName').AsString := sL_PromName;

          dL_RealDate := cdsSo033So034.FieldByName('RealDate').AsDateTime;
          if dL_RealDate <> 0 then
            FieldByName('RealDate').AsDateTime := dL_RealDate;

          dL_ShouldDate := cdsSo033So034.FieldByName('ShouldDate').AsDateTime;
          if dL_ShouldDate <> 0 then
            FieldByName('ShouldDate').AsDateTime := dL_ShouldDate;

          if sL_CustID <> '' then
            FieldByName('CustID').AsInteger := StrToInt(sL_CustID);

          Post;
        end;
      end
      else if sL_FirstFlag<>'1' then   //�򦬪���Ƥ��ΰh�O
      begin

      end
      else
      begin
        cdsSo033Log.Append;
        cdsSo033Log.FieldByName('CUSTID').AsString:=cdsSo033So034.FieldByName('CUSTID').AsString;
        cdsSo033Log.FieldByName('BILLNO').AsString:=cdsSo033So034.FieldByName('BILLNO').AsString;
        cdsSo033Log.FieldByName('ITEM').AsString:=cdsSo033So034.FieldByName('ITEM').AsString;
        cdsSo033Log.FieldByName('CITEMCODE').AsString:=cdsSo033So034.FieldByName('CITEMCODE').AsString;
        cdsSo033Log.FieldByName('CITEMNAME').AsString:=cdsSo033So034.FieldByName('CITEMNAME').AsString;
        cdsSo033Log.FieldByName('SHOULDDATE').AsString:=cdsSo033So034.FieldByName('SHOULDDATE').AsString;
        cdsSo033Log.FieldByName('REALDATE').AsString:=cdsSo033So034.FieldByName('REALDATE').AsString;
        cdsSo033Log.FieldByName('REALAMT').AsString:=cdsSo033So034.FieldByName('REALAMT').AsString;
        cdsSo033Log.FieldByName('REALPERIOD').AsString:=cdsSo033So034.FieldByName('RealPeriod').AsString;
        cdsSo033Log.FieldByName('REALSTARTDATE').AsString:=cdsSo033So034.FieldByName('REALSTARTDATE').AsString;
        cdsSo033Log.FieldByName('SBILLNO').AsString:=cdsSo033So034.FieldByName('SBILLNO').AsString;
        cdsSo033Log.FieldByName('SITEM').AsInteger:=cdsSo033So034.FieldByName('SITEM').AsInteger;
        cdsSo033Log.FieldByName('Notes').AsString := '�@�����I-�h�O��ƹ������쥿�����';
        cdsSo033Log.Post;
      end;


      cdsSo033So034.Next;
    end;

    L_StrList.Free;
end;

function TdtmMain1.CustOnceIDCheck(sI_CustID, sI_CompCode: String): Boolean;
var sL_SQL : String;
begin
    //�Y���Ȥ�w�g�Q��L����,�h�Ǧ^true
    with qryCommon do
    begin
      sL_SQL:='Select COUNT(*) CNT from So135 where CUSTID='''+sI_CustID+'''';
      sL_SQL:=sL_SQL+' and CompCode='+sI_CompCode;
      Close;
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;
      if FieldByName('CNT').AsInteger>0 then
        Result:=True
      else
        Result:=False;
      Close;
    end;
end;



function TdtmMain1.insertSO135(sI_CompCode, sI_CustID, sI_AcceptEn,
  sI_AcceptName: String): boolean;
var
    bL_Result : boolean;
begin
    bL_Result := true;
    with cdsSo135 do
    begin
      if state=dsInactive then
        Open;

      if not Locate('COMPCODE;CUSTID',VarArrayOf([sI_CompCode,sI_CustID]),[]) then
      begin
        Append;
        FieldByName('COMPCODE').AsInteger := StrToInt(sI_CompCode);//qryCallInCheck.FieldByName('COMPCODE').AsString;
        FieldByName('CUSTID').AsInteger := StrToInt(sI_CustID);
        FieldByName('COMPUTEYM').AsString := sG_OnceComputeYM;

        FieldByName('COMMDATE').AsDateTime:=StrToDate(FormatDateTime('YYYY/MM/DD',Now));
        FieldByName('CSRID').AsString := sI_AcceptEn;
        FieldByName('CSRNAME').AsString := sI_AcceptName;
        FieldByName('COMM').AsFloat := nG_SelfOrderChannelAcceptComm;
        FieldByName('Operator').AsString := frmMainMenu.sG_User;
        FieldByName('UpdTime').AsString := sG_CurDate;
        Post;
        bL_Result := false;
      end;
    end;
    result := bL_Result;

end;

procedure TdtmMain1.insertCallInDataToSo134(bI_HasPayComm : Boolean;sI_PromCode, sI_PromName,
  sI_CompCode, sI_ComputeYM, sI_AcceptEn, sI_AcceptName,sI_ClctId,sI_ClctName : String;
  I_DataSet: TDataSet;fI_IntAccOriPercent : Double);
var
    dL_RealDate,dL_ShouldDate : TDate;
    sL_CustID,sL_FirstFlag : String;
begin
    with cdsSo134 do
    begin
      Append;
      FieldByName('CompCode').AsString := sI_CompCode;
      FieldByName('ComputeYM').AsString := sG_OnceComputeYM;
      FieldByName('BELONGYM').AsString := sG_OnceBelongYM;
      FieldByName('BillNo').AsString := I_DataSet.FieldByName('BillNo').AsString;
      FieldByName('Item').AsString := I_DataSet.FieldByName('Item').AsString;
      FieldByName('PeriodId').AsInteger := I_DataSet.FieldByName('REALPERIOD').AsInteger;

      //�ۨӹq���ФH�����ȵ��@��
      if not bI_HasPayComm then
      begin
        FieldByName('IntAccId').AsString := sI_AcceptEn;
        FieldByName('IntAccName').AsString := sI_AcceptName;

        //if fI_IntAccOriPercent <> 0 then   //�Y������ҳ]�w����0,�h�������z�H����
        FieldByName('IntAccComm').AsFloat := nG_SelfOrderChannelAcceptComm;
      end;

      FieldByName('ClctEn').AsString := sI_ClctId;
      FieldByName('ClctEnName').AsString := sI_ClctName;
      FieldByName('ClctEnComm').AsFloat := RoundTo(fG_ChargeComm,-2);


      FieldByName('RealStartDate').AsDateTime := I_DataSet.FieldByName('RealStartDate').AsDateTime;
      FieldByName('RealStopDate').AsDateTime := I_DataSet.FieldByName('RealStopDate').AsDateTime;
      FieldByName('CitemCode').AsInteger := I_DataSet.FieldByName('CitemCode').AsInteger;
      FieldByName('CitemName').AsString := I_DataSet.FieldByName('CitemName').AsString;

      FieldByName('MediaCode').AsInteger := StrToInt(frmMainMenu.sG_CompCode);
      FieldByName('MediaName').AsString := frmMainMenu.sG_CompName;

      FieldByName('OPERATOR').AsString := frmMainMenu.sG_User;
      FieldByName('UpdTime').AsString := sG_CurDate;


      if sI_PromCode <> '' then
        FieldByName('PromCode').AsInteger := StrToInt(sI_PromCode);
      FieldByName('PromName').AsString := sI_PromName;


      dL_RealDate := I_DataSet.FieldByName('RealDate').AsDateTime;
      if dL_RealDate <> 0 then
        FieldByName('RealDate').AsDateTime := dL_RealDate;

      dL_ShouldDate := I_DataSet.FieldByName('ShouldDate').AsDateTime;
      if dL_ShouldDate <> 0 then
        FieldByName('ShouldDate').AsDateTime := dL_ShouldDate;

      sL_CustID := I_DataSet.FieldByName('CustID').AsString;
      if sL_CustID <> '' then
        FieldByName('CustID').AsString := sL_CustID;

      sL_FirstFlag := I_DataSet.FieldByName('FirstFlag').AsString;
      if sL_FirstFlag = '1' then
        FieldByName('FirstFlag').AsString := sL_FirstFlag;

      Post;
    end;
end;

procedure TdtmMain1.getSO134Data(sI_BelongYM, sI_CompCode,
  sI_CODENO: String);
Var
    sL_SQL,sL_Where,sL_SQLTitle,sL_FirstFlag : String;
    aValue: String;
begin

   sL_SQLTitle:='select COMPCODE,COMPUTEYM,BELONGYM,BILLNO,ITEM,'+
                'PERIODID,CLCTEN,CLCTENNAME,'+
                'CLCTENCOMM,TO_CHAR(CLCTENCOMM) StrCLCTENCOMM,CLCTENORIPERCENT,INTACCID,INTACCNAME,INTACCCOMM,TO_CHAR(INTACCCOMM)StrINTACCCOMM,INTACCORIPERCENT,'+
                'STAFFCODE,REALSTARTDATE,REALSTOPDATE,OPERATOR,UPDTIME,CITEMCODE,'+
                'CITEMNAME,MEDIACODE,MEDIANAME,REALAMT,TO_CHAR(REALAMT) StrREALAMT,PROMCODE,PROMNAME,REALDATE,SHOULDDATE,CUSTID,FIRSTFLAG '
                +' from SO134';



    sL_Where:=' where citemcode in ( '+ sI_CODENO +') AND BelongYM='''
        + sI_BelongYM+''' AND COMPCODE='+ sI_CompCode;

    sL_SQL:=sL_SQLTitle+sL_Where;

    with cdsSO134Data do
    begin
      Close;
      CommandText := sL_SQL;
      Open;
    end;

    //���t��
    transSo134Notation(cdsSO134Data);


    //�W�D����� Excel
    cdsSO134Data.First;
    while not cdsSO134Data.Eof do
    begin
      cdsSO122Channel.Append;

      cdsSO122Channel.FieldByName('COMPCODE').AsString:= cdsSO134Data.FieldByName('COMPCODE').AsString;

      aValue := cdsSO134Data.FieldByName('COMPUTEYM').AsString;
//      if ( aValue <> EmptyStr ) then
//        aValue := IntToStr( StrToInt( Copy( aValue, 1, 3 ) ) + 1911 ) + Copy( aValue, 4, 2 );

      cdsSO122Channel.FieldByName('COMPUTEYM').AsString:= aValue;

      aValue := cdsSO134Data.FieldByName('BELONGYM').AsString;
//      if ( aValue <> EmptyStr ) then
//        aValue := IntToStr( StrToInt( Copy( aValue, 1, 3 ) ) + 1911 ) + Copy( aValue, 4, 2 );
      cdsSO122Channel.FieldByName('BELONGYM').AsString:= aValue;
      
      cdsSO122Channel.FieldByName('BILLNO').AsString := cdsSO134Data.FieldByName('BILLNO').AsString;
      cdsSO122Channel.FieldByName('ITEM').AsString := cdsSO134Data.FieldByName('ITEM').AsString;
      cdsSO122Channel.FieldByName('PERIODID').AsString := cdsSO134Data.FieldByName('PERIODID').AsString;
      cdsSO122Channel.FieldByName('CLCTEN').AsString := cdsSO134Data.FieldByName('CLCTEN').AsString;
      cdsSO122Channel.FieldByName('CLCTENNAME').AsString := cdsSO134Data.FieldByName('CLCTENNAME').AsString;
      cdsSO122Channel.FieldByName('CLCTENCOMM').AsFloat := cdsSO134Data.FieldByName('CLCTENCOMM').AsFloat;
      cdsSO122Channel.FieldByName('CLCTENORIPERCENT').AsFloat := cdsSO134Data.FieldByName('CLCTENORIPERCENT').AsFloat;
      cdsSO122Channel.FieldByName('STAFFCODE').AsString := cdsSO134Data.FieldByName('STAFFCODE').AsString;
      cdsSO122Channel.FieldByName('INTACCID').AsString := cdsSO134Data.FieldByName('INTACCID').AsString;
      cdsSO122Channel.FieldByName('INTACCNAME').AsString := cdsSO134Data.FieldByName('INTACCNAME').AsString;
      cdsSO122Channel.FieldByName('INTACCCOMM').AsFloat := cdsSO134Data.FieldByName('INTACCCOMM').AsFloat;
      cdsSO122Channel.FieldByName('INTACCORIPERCENT').AsFloat := cdsSO134Data.FieldByName('INTACCORIPERCENT').AsFloat;
      cdsSO122Channel.FieldByName('REALSTARTDATE').AsString := cdsSO134Data.FieldByName('REALSTARTDATE').AsString;
      cdsSO122Channel.FieldByName('REALSTOPDATE').AsString := cdsSO134Data.FieldByName('REALSTOPDATE').AsString;
      cdsSO122Channel.FieldByName('CITEMCODE').AsString := cdsSO134Data.FieldByName('CITEMCODE').AsString;
      cdsSO122Channel.FieldByName('CITEMNAME').AsString := cdsSO134Data.FieldByName('CITEMNAME').AsString;
      cdsSO122Channel.FieldByName('MEDIACODE').AsString := cdsSO134Data.FieldByName('MEDIACODE').AsString;
      cdsSO122Channel.FieldByName('MEDIANAME').AsString := cdsSO134Data.FieldByName('MEDIANAME').AsString;
      cdsSO122Channel.FieldByName('REALAMT').AsFloat := cdsSO134Data.FieldByName('REALAMT').AsFloat;
      cdsSO122Channel.FieldByName('OPERATOR').AsString := cdsSO134Data.FieldByName('OPERATOR').AsString;
      cdsSO122Channel.FieldByName('UPDTIME').AsString := cdsSO134Data.FieldByName('UPDTIME').AsString;
      cdsSO122Channel.FieldByName('REALDATE').AsString := cdsSO134Data.FieldByName('REALDATE').AsString;
      cdsSO122Channel.FieldByName('SHOULDDATE').AsString := cdsSO134Data.FieldByName('SHOULDDATE').AsString;
      cdsSO122Channel.FieldByName('PROMCODE').AsInteger := cdsSO134Data.FieldByName('PROMCODE').AsInteger;
      cdsSO122Channel.FieldByName('PROMNAME').AsString := cdsSO134Data.FieldByName('PROMNAME').AsString;
      cdsSO122Channel.FieldByName('CustID').AsString := cdsSO134Data.FieldByName('CustID').AsString;

      sL_FirstFlag := cdsSO134Data.FieldByName('FirstFlag').AsString;
      if sL_FirstFlag ='1' then
        cdsSO122Channel.FieldByName('FirstFlagName').AsString := '����'
      else
        cdsSO122Channel.FieldByName('FirstFlagName').AsString := '��';

      cdsSO122Channel.Post;

      cdsSO134Data.Next;
    end;

end;

procedure TdtmMain1.transSo134Notation(I_CDS: TClientDataSet);
var
    sL_StrWorkeren1Comm,sL_StrClctenComm,sL_StrIntaccComm,sL_StrRealAmt : String;
begin
    I_CDS.First;
    with I_CDS do
    begin
      while not Eof do
      begin
        Edit;

        sL_StrRealAmt := FieldByName('StrRealAmt').AsString;
        if sL_StrRealAmt <> '' then
          FieldByName('RealAmt').AsFloat := StrToFloat(sL_StrRealAmt);

        //sL_StrWorkeren1Comm := FieldByName('StrWorkeren1Comm').AsString;
        //if sL_StrWorkeren1Comm <> '' then
        //  FieldByName('Workeren1Comm').AsFloat := StrToFloat(sL_StrWorkeren1Comm);

        sL_StrClctenComm := FieldByName('StrClctenComm').AsString;
        if sL_StrClctenComm <> '' then
          FieldByName('ClctenComm').AsFloat := StrToFloat(sL_StrClctenComm);

        sL_StrIntaccComm := FieldByName('StrIntaccComm').AsString;
        if sL_StrIntaccComm <> '' then
          FieldByName('IntaccComm').AsFloat := StrToFloat(sL_StrIntaccComm);

        Post;

        Next;
      end;
    end;


end;

procedure TdtmMain1.getLastData(sI_Table,sI_CompCode: String;var sI_LastUser,sI_LastComputeYM,sI_LastUpdTime : String);
var
    sL_SQL,sL_MaxUpdTime : String;
begin
    sL_SQL := 'SELECT MAX(UPDTIME) UPDTIME FROM ' + sI_Table;

    with qryCom do
    begin
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;

      sL_MaxUpdTime := FieldByName('UPDTIME').AsString;
    end;

    sL_SQL := 'SELECT COMPUTEYM,UPDTIME,OPERATOR FROM ' + sI_Table +
              ' WHERE UPDTIME=''' + sL_MaxUpdTime + '''';

    with qryCom do
    begin
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;

      sI_LastComputeYM := FieldByName('COMPUTEYM').AsString;
      sI_LastUpdTime := FieldByName('UPDTIME').AsString;
      sI_LastUser := FieldByName('OPERATOR').AsString;
    end;
end;

function TdtmMain1.getStatisticData(sI_ExcelOrRpt,sI_PayMode,sI_CompCode,sI_BelongYM,sI_CitemCodeStr : String;bI_IncludeWorkerComm : Boolean): Boolean;
var
    sL_SQL,sL_Table,sL_Where : String;
    sL_EmpID,sL_EmpName,sL_StrComm : String;
    nL_Counts : Integer;
    fL_Comm : Double;
begin
    //���Ŀ�]�tBOX�u�{�H������
    if bI_IncludeWorkerComm then
    begin
      //�έpBOX�u�{�H��***********************************************************
      sL_SQL := 'SELECT WORKEREN1,WORKEREN1NAME,NVL(SUM(WORKEREN1COMM),0) WORKEREN1COMM,' +
                'COUNT(*) COUNTS,TO_CHAR(NVL(SUM(WORKEREN1COMM),0)) STRWCOMM FROM SO122 ' +
                'WHERE COMPCODE=' + sI_CompCode + ' AND BELONGYM=''' + sI_BelongYM +
                ''' AND STBNO IS NOT NULL AND WORKEREN1COMM<>0 ' +
                ' GROUP BY WORKEREN1,WORKEREN1NAME ';
      with qryCom do
      begin
        SQL.Clear;
        SQL.Add(sL_SQL);
        Open;

        qryCom.First;
        while not qryCom.Eof do
        begin
          sL_EmpID := FieldByName('WORKEREN1').AsString;
          sL_EmpName := FieldByName('WORKEREN1NAME').AsString;
          nL_Counts := FieldByName('COUNTS').AsInteger;
          fL_Comm := FieldByName('WORKEREN1COMM').AsFloat;
          sL_StrComm := FieldByName('STRWCOMM').AsString;

          //�s�W��CDS
          insertToCdsStatisticExcel(BOX_DATA,EMP_TYPE_WORKER,sI_CompCode,sI_BelongYM,sL_EmpID,sL_EmpName,sL_StrComm,nL_Counts,fL_Comm);

          qryCom.Next;
        end;
      end;
    end;


    //�έpBOX���ФH************************************************************
    sL_SQL := 'SELECT INTACCID,INTACCNAME,NVL(SUM(INTACCCOMM),0) INTACCCOMM,' +
              'COUNT(*) COUNTS,TO_CHAR(NVL(SUM(INTACCCOMM),0)) STRICOMM FROM SO122 ' +
              'WHERE COMPCODE=' + sI_CompCode + ' AND BELONGYM=''' + sI_BelongYM +
              ''' AND STBNO IS NOT NULL AND INTACCCOMM <>0' +
              ' GROUP BY INTACCID,INTACCNAME ';
    with qryCom do
    begin
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;


      qryCom.First;
      while not qryCom.Eof do
      begin
        sL_EmpID := FieldByName('INTACCID').AsString;
        sL_EmpName := FieldByName('INTACCNAME').AsString;
        nL_Counts := FieldByName('COUNTS').AsInteger;
        fL_Comm := FieldByName('INTACCCOMM').AsFloat;
        sL_StrComm := FieldByName('STRICOMM').AsString;

        //�s�W��CDS
        insertToCdsStatisticExcel(BOX_DATA,EMP_TYPE_INTACC,sI_CompCode,sI_BelongYM,sL_EmpID,sL_EmpName,sL_StrComm,nL_Counts,fL_Comm);

        qryCom.Next;
      end;
    end;

//*****************************************************************************
//*****************************************************************************
    if sI_PayMode = BATCH_PAY then //�������I
      sL_Table := 'SO122'
    else if sI_PayMode = ONCE_PAY then //�@�����I
      sL_Table := 'SO134';


    //�έp�W�D���O�H��
    sL_SQL := 'SELECT CLCTEN,CLCTENNAME,NVL(SUM(CLCTENCOMM),0) CLCTENCOMM,' +
              'COUNT(*) COUNTS,TO_CHAR(NVL(SUM(CLCTENCOMM),0)) STRCCOMM FROM ' + sL_Table ;

    sL_Where := ' WHERE COMPCODE=' + sI_CompCode + ' AND BELONGYM=''' + sI_BelongYM +
                  ''' AND BILLNO IS NOT NULL AND CLCTENCOMM<>0 ';

    if sI_CitemCodeStr <> '' then
      sL_Where := sL_Where + ' AND CITEMCODE IN (' + sI_CitemCodeStr + ') ';

    sL_Where := sL_Where +  ' GROUP BY CLCTEN,CLCTENNAME';
{
    if sI_ExcelOrRpt = EXCEL_MODE then   //Excel�n�[CitemCode����
    begin
      sL_Where := ' WHERE COMPCODE=' + sI_CompCode + ' AND BELONGYM=''' + sI_BelongYM +
                  ''' AND BILLNO IS NOT NULL AND CITEMCODE IN (' + sI_CitemCodeStr + ') ' +
                  ' AND CLCTENCOMM<>0 GROUP BY CLCTEN,CLCTENNAME ';
    end
    else if sI_ExcelOrRpt = REPORT_MODE then
    begin
      sL_Where := ' WHERE COMPCODE=' + sI_CompCode + ' AND BELONGYM=''' + sI_BelongYM +
                  ''' AND BILLNO IS NOT NULL ' +
                  ' AND CLCTENCOMM<>0 GROUP BY CLCTEN,CLCTENNAME ';
    end;
}

    sL_SQL := sL_SQL + sL_Where;

    with qryCom do
    begin
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;


      qryCom.First;
      while not qryCom.Eof do
      begin
        sL_EmpID := FieldByName('CLCTEN').AsString;
        sL_EmpName := FieldByName('CLCTENNAME').AsString;
        nL_Counts := FieldByName('COUNTS').AsInteger;
        fL_Comm := FieldByName('CLCTENCOMM').AsFloat;
        sL_StrComm := FieldByName('STRCCOMM').AsString;

        //�s�W��CDS
        insertToCdsStatisticExcel(CHANNEL_DATA,EMP_TYPE_CLCT,sI_CompCode,sI_BelongYM,sL_EmpID,sL_EmpName,sL_StrComm,nL_Counts,fL_Comm);

        qryCom.Next;
      end;
    end;


    //�έp�W�D���ФH
    sL_SQL := 'SELECT INTACCID,INTACCNAME,NVL(SUM(INTACCCOMM),0) INTACCCOMM,' +
              'COUNT(*) COUNTS,TO_CHAR(NVL(SUM(INTACCCOMM),0)) STRICOMM FROM ' + sL_Table ;

    sL_Where := ' WHERE COMPCODE=' + sI_CompCode + ' AND BELONGYM=''' + sI_BelongYM +
                  ''' AND BILLNO IS NOT NULL AND INTACCCOMM <>0 ';

    if sI_CitemCodeStr <> '' then
      sL_Where := sL_Where + ' AND CITEMCODE IN (' + sI_CitemCodeStr + ') ';

    sL_Where := sL_Where +  ' GROUP BY INTACCID,INTACCNAME ';


{
    if sI_ExcelOrRpt = EXCEL_MODE then   //Excel�n�[CitemCode����
    begin
      sL_Where := ' WHERE COMPCODE=' + sI_CompCode + ' AND BELONGYM=''' + sI_BelongYM +
                  ''' AND BILLNO IS NOT NULL AND CITEMCODE IN (' + sI_CitemCodeStr + ') ' +
                  ' AND INTACCCOMM <>0 GROUP BY INTACCID,INTACCNAME ';
    end
    else if sI_ExcelOrRpt = REPORT_MODE then
    begin
      sL_Where := ' WHERE COMPCODE=' + sI_CompCode + ' AND BELONGYM=''' + sI_BelongYM +
                  ''' AND BILLNO IS NOT NULL ' +
                  ' AND INTACCCOMM <>0 GROUP BY INTACCID,INTACCNAME ';
    end;
}

    sL_SQL := sL_SQL + sL_Where;


    with qryCom do
    begin
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;

      qryCom.First;
      while not qryCom.Eof do
      begin
        sL_EmpID := FieldByName('INTACCID').AsString;
        sL_EmpName := FieldByName('INTACCNAME').AsString;
        nL_Counts := FieldByName('COUNTS').AsInteger;
        fL_Comm := FieldByName('INTACCCOMM').AsFloat;
        sL_StrComm := FieldByName('STRICOMM').AsString;

        //�s�W��CDS
        insertToCdsStatisticExcel(CHANNEL_DATA,EMP_TYPE_INTACC,sI_CompCode,sI_BelongYM,sL_EmpID,sL_EmpName,sL_StrComm,nL_Counts,fL_Comm);

        qryCom.Next;
      end;
    end;

    //�[�`�ζ�t��
    computeTotalComm(cdsStatisticExcel);
    addColumnZero(cdsStatisticExcel);

    if cdsStatisticExcel.RecordCount > 1 then    //�]���N�ⳣ�S���,�٦��@���[�`���
      Result := true
    else
      Result := false;
end;

procedure TdtmMain1.insertToCdsStatisticExcel(sI_BoxOrChannel,sI_EmpType, sI_CompCode,
  sI_BelongYM, sI_EmpID, sI_EmpName,sI_StrComm: String; nI_Counts: Integer;
  fI_Comm: Double);
begin
    if (sI_EmpID<>'') OR (sI_EmpName<>'') then
    begin
      with cdsStatisticExcel do
      begin
        if cdsStatisticExcel.Locate('COMPCODE;EMPID;EMPNAME',VarArrayOf([sI_CompCode,sI_EmpID,sI_EmpName]),[]) then
        begin
          cdsStatisticExcel.Edit;

          if sI_EmpType = EMP_TYPE_WORKER then
          begin
            cdsStatisticExcel.FieldByName('WorkerCounts').AsInteger := nI_Counts;
            cdsStatisticExcel.FieldByName('WorkerComm').AsFloat := fI_Comm;
            cdsStatisticExcel.FieldByName('StrWComm').AsString := sI_StrComm;
          end
          else if sI_EmpType = EMP_TYPE_CLCT then
          begin
            cdsStatisticExcel.FieldByName('ClctCounts').AsInteger := nI_Counts;
            cdsStatisticExcel.FieldByName('ClctComm').AsFloat := fI_Comm;
            cdsStatisticExcel.FieldByName('StrCComm').AsString := sI_StrComm;
          end
          else if sI_EmpType = EMP_TYPE_INTACC then
          begin
            if sI_BoxOrChannel=BOX_DATA then
            begin
              cdsStatisticExcel.FieldByName('BoxIntAccCounts').AsInteger := nI_Counts;
              cdsStatisticExcel.FieldByName('BoxIntAccComm').AsFloat := fI_Comm;
              cdsStatisticExcel.FieldByName('BoxStrIComm').AsString := sI_StrComm;
            end
            else if sI_BoxOrChannel=CHANNEL_DATA then
            begin
              cdsStatisticExcel.FieldByName('ChIntAccCounts').AsInteger := nI_Counts;
              cdsStatisticExcel.FieldByName('ChIntAccComm').AsFloat := fI_Comm;
              cdsStatisticExcel.FieldByName('ChStrIComm').AsString := sI_StrComm;
            end
          end;

          cdsStatisticExcel.Post;
        end
        else
        begin
          cdsStatisticExcel.Append;
          cdsStatisticExcel.FieldByName('COMPCODE').AsString := sI_CompCode;
          cdsStatisticExcel.FieldByName('EMPID').AsString := sI_EmpID;
          cdsStatisticExcel.FieldByName('EMPNAME').AsString := sI_EmpName;

          if sI_EmpType = EMP_TYPE_WORKER then
          begin
            cdsStatisticExcel.FieldByName('WorkerCounts').AsInteger := nI_Counts;
            cdsStatisticExcel.FieldByName('WorkerComm').AsFloat := fI_Comm;
            cdsStatisticExcel.FieldByName('StrWComm').AsString := sI_StrComm;
          end
          else if sI_EmpType = EMP_TYPE_CLCT then
          begin
            cdsStatisticExcel.FieldByName('ClctCounts').AsInteger := nI_Counts;
            cdsStatisticExcel.FieldByName('ClctComm').AsFloat := fI_Comm;
            cdsStatisticExcel.FieldByName('StrCComm').AsString := sI_StrComm;

          end
          else if sI_EmpType = EMP_TYPE_INTACC then
          begin
            if sI_BoxOrChannel=BOX_DATA then
            begin
              cdsStatisticExcel.FieldByName('BoxIntAccCounts').AsInteger := nI_Counts;
              cdsStatisticExcel.FieldByName('BoxIntAccComm').AsFloat := fI_Comm;
              cdsStatisticExcel.FieldByName('BoxStrIComm').AsString := sI_StrComm;

            end
            else if sI_BoxOrChannel=CHANNEL_DATA then
            begin
              cdsStatisticExcel.FieldByName('ChIntAccCounts').AsInteger := nI_Counts;
              cdsStatisticExcel.FieldByName('ChIntAccComm').AsFloat := fI_Comm;
              cdsStatisticExcel.FieldByName('ChStrIComm').AsString := sI_StrComm;

            end;
          end;

          cdsStatisticExcel.Post;
        end;
      end;
    end;
end;

procedure TdtmMain1.computeTotalComm(I_CDS : TClientDataSet);
var
    sL_WorkerComm,sL_ClctComm,sL_BoxIntAccComm,sL_ChIntAccComm : String;

    fL_WorkerComm,fL_WorkerTotalComm : Double;
    fL_ClctComm,fL_ClctTotalComm : Double;
    fL_BoxIntAccComm,fL_BoxIntAccTotalComm : Double;
    fL_ChIntAccComm,fL_ChIntAccTotalComm : Double;
    fL_TotalComm,fL_TotalGroupComm : Double;

    nL_WorkerCounts,nL_WorkerTotalCounts : Integer;
    nL_ClctCounts,nL_ClctTotalCounts : Integer;
    nL_BoxIntAccCounts,nL_BoxIntAccTotalCounts : Integer;
    nL_ChIntAccCounts,nL_ChIntAccTotalCounts : Integer;
begin
    fL_WorkerTotalComm := 0;
    fL_ClctTotalComm := 0;
    fL_BoxIntAccTotalComm := 0;
    fL_ChIntAccTotalComm := 0;

    nL_WorkerTotalCounts := 0;
    nL_ClctTotalCounts := 0;
    nL_BoxIntAccTotalCounts := 0;
    nL_ChIntAccTotalCounts := 0;


    I_CDS.First;
    while not I_CDS.Eof do
    begin
      //�u�{�H���[�`
      sL_WorkerComm := I_CDS.FieldByName('StrWComm').AsString;

      if sL_WorkerComm = '' then
        fL_WorkerComm := 0
      else
        fL_WorkerComm := StrToFloat(sL_WorkerComm);
      fL_WorkerTotalComm := fL_WorkerTotalComm + fL_WorkerComm;

      nL_WorkerCounts := I_CDS.FieldByName('WorkerCounts').AsInteger;
      nL_WorkerTotalCounts := nL_WorkerTotalCounts + nL_WorkerCounts;

      //���O�H���[�`
      sL_ClctComm := I_CDS.FieldByName('StrCComm').AsString;

      if sL_ClctComm = '' then
        fL_ClctComm := 0
      else
        fL_ClctComm := StrToFloat(sL_ClctComm);

      fL_ClctTotalComm := fL_ClctTotalComm + fL_ClctComm;

      nL_ClctCounts := I_CDS.FieldByName('ClctCounts').AsInteger;
      nL_ClctTotalCounts := nL_ClctTotalCounts + nL_ClctCounts;

      //Box���ФH�[�`
      sL_BoxIntAccComm := I_CDS.FieldByName('BoxStrIComm').AsString;

      if sL_BoxIntAccComm = '' then
        fL_BoxIntAccComm := 0
      else
        fL_BoxIntAccComm := StrToFloat(sL_BoxIntAccComm);

      fL_BoxIntAccTotalComm := fL_BoxIntAccTotalComm + fL_BoxIntAccComm;

      nL_BoxIntAccCounts := I_CDS.FieldByName('BoxIntAccCounts').AsInteger;
      nL_BoxIntAccTotalCounts := nL_BoxIntAccTotalCounts + nL_BoxIntAccCounts;


      //�W�D���ФH�[�`
      sL_ChIntAccComm := I_CDS.FieldByName('ChStrIComm').AsString;

      if sL_ChIntAccComm = '' then
        fL_ChIntAccComm := 0
      else
        fL_ChIntAccComm := StrToFloat(sL_ChIntAccComm);

      fL_ChIntAccTotalComm := fL_ChIntAccTotalComm + fL_ChIntAccComm;

      nL_ChIntAccCounts := I_CDS.FieldByName('ChIntAccCounts').AsInteger;
      nL_ChIntAccTotalCounts := nL_ChIntAccTotalCounts + nL_ChIntAccCounts;



      //��V�[�`
      fL_TotalGroupComm := fL_WorkerComm + fL_ClctComm + fL_BoxIntAccComm + fL_ChIntAccComm;
      fL_TotalComm := fL_TotalComm + fL_TotalGroupComm;

      I_CDS.Edit;

      I_CDS.FieldByName('WorkerComm').AsFloat := fL_WorkerComm;
      I_CDS.FieldByName('ClctComm').AsFloat := fL_ClctComm;
      I_CDS.FieldByName('BoxIntAccComm').AsFloat := fL_BoxIntAccComm;
      I_CDS.FieldByName('ChIntAccComm').AsFloat := fL_ChIntAccComm;

      I_CDS.FieldByName('TotalComm').AsFloat := fL_TotalGroupComm;

      I_CDS.Next;
    end;

    I_CDS.Append;

    cdsStatisticExcel.FieldByName('EMPID').AsString := '�`�p';

    cdsStatisticExcel.FieldByName('WorkerComm').AsFloat := fL_WorkerTotalComm;
    cdsStatisticExcel.FieldByName('WorkerCounts').AsInteger := nL_WorkerTotalCounts;

    cdsStatisticExcel.FieldByName('ClctComm').AsFloat := fL_ClctTotalComm;
    cdsStatisticExcel.FieldByName('ClctCounts').AsInteger := nL_ClctTotalCounts;

    cdsStatisticExcel.FieldByName('BoxIntAccComm').AsFloat := fL_BoxIntAccTotalComm;
    cdsStatisticExcel.FieldByName('BoxIntAccCounts').AsInteger := nL_BoxIntAccTotalCounts;

    cdsStatisticExcel.FieldByName('ChIntAccComm').AsFloat := fL_ChIntAccTotalComm;
    cdsStatisticExcel.FieldByName('ChIntAccCounts').AsInteger := nL_ChIntAccTotalCounts;

    cdsStatisticExcel.FieldByName('TotalComm').AsFloat := fL_TotalComm;

    I_CDS.Post;


end;

function TdtmMain1.TransToStatisticExcel(sI_BelongYM,
  sI_CurrDateTime: String;bI_IncludeWorkerComm : Boolean): String;
Var
  sL_QueryString,sL_FileName,sL_FileCurrDateTime,sL_CurrDateTime : String;
  L_StrList : TStringList;
  sL_NoDataMsg,sL_HaveDataMsg,sL_IncludeWorkerComm : String;
begin

    addColumnZero(cdsStatisticExcel);

    sL_FileCurrDateTime := sI_CurrDateTime;
    sL_CurrDateTime := ReplaceStr(Trim(ReplaceStr(sL_FileCurrDateTime,'/')),':');


    if TClientDataSet(cdsStatisticExcel).IndexName = 'TmpIndex' then
      TClientDataSet(cdsStatisticExcel).DeleteIndex('TmpIndex');
    TClientDataSet(cdsStatisticExcel).AddIndex('TmpIndex', 'EmpID', [ixCaseInsensitive],'','',0);
    TClientDataSet(cdsStatisticExcel).IndexName := 'TmpIndex';

    if bI_IncludeWorkerComm then
      sL_IncludeWorkerComm := '�]�tBOX�u�{�H������'
    else
      sL_IncludeWorkerComm := '���]�tBOX�u�{�H������';

    sL_FileName := sG_ExcelSavePath + '\�έp��_' + frmMainMenu.sG_CompCode + '_' + sI_BelongYM + '_' + sL_CurrDateTime + '.XLS';
    sL_HaveDataMsg := '�p�⧹��,�έp��s��  ';
    sL_NoDataMsg := '�S���έp���� ';

    sL_QueryString:='�p��ɶ�:' + sL_FileCurrDateTime + '@' + '����H��:' + frmMainMenu.sG_User+'@' +
                    '���q�W��:' + frmMainMenu.sG_CompName + ' (' + sL_IncludeWorkerComm + ')@' + '�k�ݦ~��:' + sI_BelongYM + '@' +
                    '�d�ߦ��O����:' + frmSO8B20_1.sG_ChargeNameSQL;

    cdsStatisticExcel.DisableControls;

    //���Ŀ�]�tBOX�u�{�H������
    if bI_IncludeWorkerComm then
      DataSetToXLS(cdsStatisticExcel,sL_FileName,sL_QueryString,'CompCode,StrWComm,StrCComm,BoxStrIComm,ChStrIComm')
    else
      DataSetToXLS(cdsStatisticExcel,sL_FileName,sL_QueryString,'CompCode,StrWComm,StrCComm,BoxStrIComm,ChStrIComm,WorkerCounts,WorkerComm');

    cdsStatisticExcel.EnableControls;


    //�^�_�쪬�A
    TUCommonFun.setDefaultCursor;

    if cdsStatisticExcel.RecordCount > 0 then
      MessageDlg(sL_HaveDataMsg + sL_FileName,mtInformation,[mbOK],0)
    else
      MessageDlg(sL_NoDataMsg,mtInformation,[mbOK],0);
end;

function TdtmMain1.getRptBoxDetailData(sI_BelongYM, sI_Compute,
  sI_EmpNoSQL, sI_SalesCodeSQL, sI_OtherCompCodeSQL,
  sI_OrderList: String): Boolean;
var
    sL_SQLTitle,sL_SQL,sL_Where,sL_Table : String;
begin
    sL_SQLTitle := 'SELECT * FROM SO122 ';
    sL_Where := ' WHERE COMPCODE=' + sI_Compute + ' AND BELONGYM=''' +
                sI_BelongYM + ''' AND STBNO IS NOT NULL ';


    if ((sI_EmpNoSQL<>'') OR (sI_SalesCodeSQL<>'')
                          OR (sI_OtherCompCodeSQL<>'')) then
    begin
      if sI_EmpNoSQL <> '' then
      begin
        sL_Where := sL_Where + ' AND ((WorkerEn1 IN(' + sI_EmpNoSQL +
                    ') OR IntAccId IN(' + sI_EmpNoSQL + '))';
      end
      else
      begin
        sL_Where := sL_Where + ' AND ((WorkerEn1 IN('''') OR ' +
                    'IntAccId IN(''''))';
      end;

      if sI_SalesCodeSQL <> '' then
        sL_Where := sL_Where + ' OR IntAccId IN (' + sI_SalesCodeSQL + ')'
      else
        sL_Where := sL_Where + ' OR IntAccId IN ('''')';


      if sI_OtherCompCodeSQL <> '' then
        sL_Where := sL_Where + ' OR IntAccId IN (' + sI_OtherCompCodeSQL + '))'
      else
        sL_Where := sL_Where + ' OR IntAccId IN (''''))';

    end;
    sL_SQL := sL_SQLTitle + sL_Where + ' ORDER BY ' + sI_OrderList;

    //BOX ��ƥu�s�� SO122 ��
    with cdsSo122 do
    begin
      Close;
      CommandText := sL_SQL;
      Open;

      if RecordCount<>0 then
        Result := true
      else
        Result := false;
    end;


end;

procedure TdtmMain1.getSFReturnData(sI_ReturnData : String);
var
    L_StrList : TStringList;
begin
    L_StrList := TStringList.Create;
    L_StrList.Clear;
    L_StrList := TUstr.ParseStrings(sI_ReturnData,',');
    {�Ǧ^�ƾڶ���
     IntAccId,IntAccName,IntAccOriPercent,
     StaffCode,RealStartDate,RealStopDate,CitemCode,CitemName,
     MediaCode,MediaName,PromCode,PromName,CustId,
     PeriodId,FirstFlag,OrderNo,CallSelf
     }


    if L_StrList.Strings[0] <> 'NULL' then
      sG_SFIntAccId := L_StrList.Strings[0];

    if L_StrList.Strings[1] <> 'NULL' then
      sG_SFIntAccName := L_StrList.Strings[1];

    if L_StrList.Strings[2] <> 'NULL' then
      fG_SFIntAccOriPercent := StrToFloat(L_StrList.Strings[2]);

    if L_StrList.Strings[3] <> 'NULL' then
      sG_SFStaffCode := L_StrList.Strings[3];

    if L_StrList.Strings[4] <> 'NULL' then
      dG_SFRealStartDate := StrToDate(L_StrList.Strings[4]);

    if L_StrList.Strings[5] <> 'NULL' then
      dG_SFRealStopDate := StrToDate(L_StrList.Strings[5]);

    if L_StrList.Strings[6] <> 'NULL' then
      sG_SFCitemCode := L_StrList.Strings[6];

    if L_StrList.Strings[7] <> 'NULL' then
      sG_SFCitemName := L_StrList.Strings[7];

    if L_StrList.Strings[8] <> 'NULL' then
      sG_SFMediaCode := L_StrList.Strings[8];

    if L_StrList.Strings[9] <> 'NULL' then
      sG_SFMediaName := L_StrList.Strings[9];

    if L_StrList.Strings[10] <> 'NULL' then
      sG_SFPromCode := L_StrList.Strings[10];

    if L_StrList.Strings[11] <> 'NULL' then
      sG_SFPromName := L_StrList.Strings[11];

    if L_StrList.Strings[12] <> 'NULL' then
      sG_SFCustId := L_StrList.Strings[12];

    if L_StrList.Strings[13] <> 'NULL' then
      sG_SFPeriod := L_StrList.Strings[13];

    if L_StrList.Strings[14] <> 'NULL' then
      sG_SFFirstFlag := L_StrList.Strings[14];

    if L_StrList.Strings[15] <> 'NULL' then
      sG_SFOrderNo := L_StrList.Strings[15];

    if L_StrList.Strings[16] <> 'NULL' then
      sG_SFCallSelf := L_StrList.Strings[16];

    L_StrList.Free;
end;



procedure TdtmMain1.computeExcelTotalComm(I_CDS: TClientDataSet;
  sI_BoxOrChannel,sI_Flag : String);
var
    fL_IntAccComm,fL_ClctComm,fL_WorkerEnComm : Double;
    fL_TotalIntAccComm,fL_TotalClctComm,fL_TotalWorkerEnComm : Double;
begin

    I_CDS.First;
    while not I_CDS.Eof do
    begin
      fL_IntAccComm := I_CDS.FieldByName('INTACCCOMM').AsFloat;
      fL_TotalIntAccComm := fL_TotalIntAccComm + fL_IntAccComm;

      if sI_BoxOrChannel = CHANNEL_DATA then
      begin
        fL_ClctComm := I_CDS.FieldByName('CLCTENCOMM').AsFloat;
        fL_TotalClctComm := fL_TotalClctComm + fL_ClctComm;
      end
      else if sI_BoxOrChannel = BOX_DATA then
      begin
        fL_WorkerEnComm := I_CDS.FieldByName('WORKEREN1COMM').AsFloat;
        fL_TotalWorkerEnComm := fL_TotalWorkerEnComm + fL_WorkerEnComm;
      end;

      I_CDS.Next;
    end;

    I_CDS.Append;
    I_CDS.FieldByName('COMPUTEYM').AsString := '�`�p';

    if sI_BoxOrChannel = CHANNEL_DATA then
      I_CDS.FieldByName('CLCTENCOMM').AsFloat := fL_TotalClctComm
    else if sI_BoxOrChannel = BOX_DATA then
      I_CDS.FieldByName('WORKEREN1COMM').AsFloat := fL_TotalWorkerEnComm;

    if sI_Flag = CHANNEL_A then   //������Ҧ�����
      I_CDS.FieldByName('CustID').AsInteger := 99999999
    else if sI_Flag = CHANNEL_I then  //��������ФH����
      I_CDS.FieldByName('INTACCID').AsString := '��'
    else if sI_Flag = CHANNEL_C then  //��������O�H������
      I_CDS.FieldByName('CLCTEN').AsString := '��';

    I_CDS.FieldByName('INTACCCOMM').AsFloat := fL_TotalIntAccComm;
    I_CDS.Post;

end;



procedure TdtmMain1.cdsTotalCommInitial;
begin
    with cdsTotalCommData do
    begin
      Append;
      FieldByName('Mode').AsString := 'WorkerComm';
      FieldByName('ModeName').AsString := '�u�{������';
      Post;

      Append;
      FieldByName('Mode').AsString := 'ClctComm';
      FieldByName('ModeName').AsString := '���O������';
      Post;

      Append;
      FieldByName('Mode').AsString := 'BoxIntAccComm';
      FieldByName('ModeName').AsString := 'BOX���ФH����';
      Post;

      Append;
      FieldByName('Mode').AsString := 'ChIntAccComm';
      FieldByName('ModeName').AsString := '�W�D���ФH����';
      Post;

      Append;
      FieldByName('Mode').AsString := 'TotalComm';
      FieldByName('ModeName').AsString := '�`���B';
      Post;


      Append;
      FieldByName('Mode').AsString := '����������H��';
      Post;

      Append;
      FieldByName('Mode').AsString := 'WorkerCount';
      FieldByName('ModeName').AsString := '�u�{���H��';
      Post;

      Append;
      FieldByName('Mode').AsString := 'ClctCount';
      FieldByName('ModeName').AsString := '���O���H��';
      Post;

      Append;
      FieldByName('Mode').AsString := 'BoxIntAccCount';
      FieldByName('ModeName').AsString := 'Box���ФH�H��';
      Post;

      Append;
      FieldByName('Mode').AsString := 'ChIntAccCount';
      FieldByName('ModeName').AsString := '�W�D���ФH�H��';
      Post;

      Append;
      FieldByName('Mode').AsString := 'TotalCount';
      FieldByName('ModeName').AsString := '�`�H��';
      Post;
    end;

end;

procedure TdtmMain1.computeTotalReport(sI_BelongYM, sI_CompCode: String);
var
    sL_SQL,sL_BasicYM,sL_Month : String;
    ii : Integer;
begin
    //����J�򥻸��
    cdsTotalCommInitial;


//�B�z������(�D���ФH)�[�`**************************************************
    sL_SQL := 'select BELONGYM, sum(nvl(WORKEREN1COMM,0)) WorkerComm, sum(nvl(CLCTENCOMM,0)) ClctComm, sum(nvl(INTACCCOMM,0)) IntAccComm' +
              ',sum(nvl(WORKEREN1COMM,0)) + sum(nvl(CLCTENCOMM,0)) + sum(nvl(INTACCCOMM,0)) TotalComm' +
              ' from so122 group by BELONGYM';


    with cdsTotalComm do
    begin
      cdsTotalComm.Close;
      cdsTotalComm.CommandText := sL_SQL;
      cdsTotalComm.Open;

      //��J(�L�h�w�I)���
      addToCdsTotalComm1(cdsTotalComm,cdsTotalCommData,sI_BelongYM,'LastPay','',false);

      //��J(��C�w�I)���
      addToCdsTotalComm1(cdsTotalComm,cdsTotalCommData,sI_BelongYM,'NextPay','',false);

      for ii:=1 to 13 do
      begin
        if ii < 13 then
        begin
          sL_BasicYM := addZero(afterMonth(sI_BelongYM,ii-1),5);
          sL_Month := 'Month' + IntToStr(ii);
          addToCdsTotalComm1(cdsTotalComm,cdsTotalCommData,sL_BasicYM,sL_Month,'',false);
        end
        else
        begin
          sL_BasicYM := addZero(afterMonth(sI_BelongYM,ii-1),5);
          sL_Month := 'OtherMonth';
          addToCdsTotalComm1(cdsTotalComm,cdsTotalCommData,sL_BasicYM,sL_Month,'',false);
        end;
      end;
    end;


//�B�z������(BOX���ФH)�[�`**************************************************
    sL_SQL := 'select BELONGYM,sum(nvl(INTACCCOMM,0)) IntAccComm' +
              ' from so122 WHERE BILLNO IS NULL group by BELONGYM';

    with cdsTotalComm do
    begin
      cdsTotalComm.Close;
      cdsTotalComm.CommandText := sL_SQL;
      cdsTotalComm.Open;

      //��J(�L�h�w�I)���
      addToCdsTotalComm1(cdsTotalComm,cdsTotalCommData,sI_BelongYM,'LastPay',BOX_DATA,true);

      //��J(��C�w�I)���
      addToCdsTotalComm1(cdsTotalComm,cdsTotalCommData,sI_BelongYM,'NextPay',BOX_DATA,true);

      for ii:=1 to 13 do
      begin
        if ii < 13 then
        begin
          sL_BasicYM := addZero(afterMonth(sI_BelongYM,ii-1),5);
          sL_Month := 'Month' + IntToStr(ii);
          addToCdsTotalComm1(cdsTotalComm,cdsTotalCommData,sL_BasicYM,sL_Month,BOX_DATA,true);
        end
        else
        begin
          sL_BasicYM := addZero(afterMonth(sI_BelongYM,ii-1),5);
          sL_Month := 'OtherMonth';
          addToCdsTotalComm1(cdsTotalComm,cdsTotalCommData,sL_BasicYM,sL_Month,BOX_DATA,true);
        end;
      end;
    end;

//�B�z������(�W�D���ФH)�[�`**************************************************
    sL_SQL := 'select BELONGYM,sum(nvl(INTACCCOMM,0)) IntAccComm' +
              ' from so122 WHERE BILLNO IS NOT NULL group by BELONGYM';

    with cdsTotalComm do
    begin
      cdsTotalComm.Close;
      cdsTotalComm.CommandText := sL_SQL;
      cdsTotalComm.Open;

      //��J(�L�h�w�I)���
      addToCdsTotalComm1(cdsTotalComm,cdsTotalCommData,sI_BelongYM,'LastPay',CHANNEL_DATA,true);

      //��J(��C�w�I)���
      addToCdsTotalComm1(cdsTotalComm,cdsTotalCommData,sI_BelongYM,'NextPay',CHANNEL_DATA,true);

      for ii:=1 to 13 do
      begin
        if ii < 13 then
        begin
          sL_BasicYM := addZero(afterMonth(sI_BelongYM,ii-1),5);
          sL_Month := 'Month' + IntToStr(ii);
          addToCdsTotalComm1(cdsTotalComm,cdsTotalCommData,sL_BasicYM,sL_Month,CHANNEL_DATA,true);
        end
        else
        begin
          sL_BasicYM := addZero(afterMonth(sI_BelongYM,ii-1),5);
          sL_Month := 'OtherMonth';
          addToCdsTotalComm1(cdsTotalComm,cdsTotalCommData,sL_BasicYM,sL_Month,CHANNEL_DATA,true);
        end;
      end;
    end;

//�B�z�u�{���H�ƥ[�`**********************************************************
    sL_SQL := 'select BELONGYM, count(distinct(WORKEREN1)) WorkerCount from so122 where WORKEREN1COMM>0 group by BELONGYM';
    with cdsTotalComm do
    begin
      cdsTotalComm.Close;
      cdsTotalComm.CommandText := sL_SQL;
      cdsTotalComm.Open;

      //��J(�L�h�w�I)���
      addToCdsTotalComm2(cdsTotalComm,cdsTotalCommData,sI_BelongYM,'LastPay','WorkerCount');

      //��J(��C�w�I)���
      addToCdsTotalComm2(cdsTotalComm,cdsTotalCommData,sI_BelongYM,'NextPay','WorkerCount');

      for ii:=1 to 13 do
      begin
        if ii < 13 then
        begin
          sL_BasicYM := addZero(afterMonth(sI_BelongYM,ii-1),5);
          sL_Month := 'Month' + IntToStr(ii);
          addToCdsTotalComm2(cdsTotalComm,cdsTotalCommData,sL_BasicYM,sL_Month,'WorkerCount');
        end
        else
        begin
          sL_BasicYM := addZero(afterMonth(sI_BelongYM,ii-1),5);
          sL_Month := 'OtherMonth';
          addToCdsTotalComm2(cdsTotalComm,cdsTotalCommData,sL_BasicYM,sL_Month,'WorkerCount');
        end;
      end;
    end;

//�B�z���O���H�ƥ[�`**********************************************************
    sL_SQL := 'select BELONGYM, count(distinct(CLCTEN)) ClctCount from so122 where CLCTENCOMM>0 group by BELONGYM';
    with cdsTotalComm do
    begin
      cdsTotalComm.Close;
      cdsTotalComm.CommandText := sL_SQL;
      cdsTotalComm.Open;

      //��J(�L�h�w�I)���
      addToCdsTotalComm2(cdsTotalComm,cdsTotalCommData,sI_BelongYM,'LastPay','ClctCount');

      //��J(��C�w�I)���
      addToCdsTotalComm2(cdsTotalComm,cdsTotalCommData,sI_BelongYM,'NextPay','ClctCount');

      for ii:=1 to 13 do
      begin
        if ii < 13 then
        begin
          sL_BasicYM := addZero(afterMonth(sI_BelongYM,ii-1),5);
          sL_Month := 'Month' + IntToStr(ii);
          addToCdsTotalComm2(cdsTotalComm,cdsTotalCommData,sL_BasicYM,sL_Month,'ClctCount');
        end
        else
        begin
          sL_BasicYM := addZero(afterMonth(sI_BelongYM,ii-1),5);
          sL_Month := 'OtherMonth';
          addToCdsTotalComm2(cdsTotalComm,cdsTotalCommData,sL_BasicYM,sL_Month,'ClctCount');
        end;
      end;
    end;

//�B�zBox���ФH�H�ƥ[�`**********************************************************
    sL_SQL := 'select BELONGYM, count(distinct(INTACCID)) IntAccCount from so122 where INTACCCOMM>0 and BillNo IS NULL group by BELONGYM';
    with cdsTotalComm do
    begin
      cdsTotalComm.Close;
      cdsTotalComm.CommandText := sL_SQL;
      cdsTotalComm.Open;

      //��J(�L�h�w�I)���
      addToCdsTotalComm2(cdsTotalComm,cdsTotalCommData,sI_BelongYM,'LastPay','BoxIntAccCount');

      //��J(��C�w�I)���
      addToCdsTotalComm2(cdsTotalComm,cdsTotalCommData,sI_BelongYM,'NextPay','BoxIntAccCount');

      for ii:=1 to 13 do
      begin
        if ii < 13 then
        begin
          sL_BasicYM := addZero(afterMonth(sI_BelongYM,ii-1),5);
          sL_Month := 'Month' + IntToStr(ii);
          addToCdsTotalComm2(cdsTotalComm,cdsTotalCommData,sL_BasicYM,sL_Month,'BoxIntAccCount');
        end
        else
        begin
          sL_BasicYM := addZero(afterMonth(sI_BelongYM,ii-1),5);
          sL_Month := 'OtherMonth';
          addToCdsTotalComm2(cdsTotalComm,cdsTotalCommData,sL_BasicYM,sL_Month,'BoxIntAccCount');
        end;
      end;
    end;

//�B�z�W�D���ФH�H�ƥ[�`**********************************************************
    sL_SQL := 'select BELONGYM, count(distinct(INTACCID)) IntAccCount from so122 where INTACCCOMM>0 and BillNo IS NOT NULL group by BELONGYM';
    with cdsTotalComm do
    begin
      cdsTotalComm.Close;
      cdsTotalComm.CommandText := sL_SQL;
      cdsTotalComm.Open;

      //��J(�L�h�w�I)���
      addToCdsTotalComm2(cdsTotalComm,cdsTotalCommData,sI_BelongYM,'LastPay','ChIntAccCount');

      //��J(��C�w�I)���
      addToCdsTotalComm2(cdsTotalComm,cdsTotalCommData,sI_BelongYM,'NextPay','ChIntAccCount');

      for ii:=1 to 13 do
      begin
        if ii < 13 then
        begin
          sL_BasicYM := addZero(afterMonth(sI_BelongYM,ii-1),5);
          sL_Month := 'Month' + IntToStr(ii);
          addToCdsTotalComm2(cdsTotalComm,cdsTotalCommData,sL_BasicYM,sL_Month,'ChIntAccCount');
        end
        else
        begin
          sL_BasicYM := addZero(afterMonth(sI_BelongYM,ii-1),5);
          sL_Month := 'OtherMonth';
          addToCdsTotalComm2(cdsTotalComm,cdsTotalCommData,sL_BasicYM,sL_Month,'ChIntAccCount');
        end;
      end;
    end;
end;

procedure TdtmMain1.addToCdsTotalComm1(I_Source, I_Data: TDataSet;
  sI_BelongYM, sI_FieldName,sI_BoxOrChannel: String;bI_IsIntAccData : Boolean);
var
    fL_TotalWorkerComm,fL_TotalClctComm,fL_TotalIntAccComm,fL_TotalComm : Double;
begin
    I_Source.Filtered := false;

    if sI_FieldName = 'LastPay' then
      I_Source.Filter := 'BELONGYM<''' + sI_BelongYM + ''''
    else if sI_FieldName = 'NextPay' then
      I_Source.Filter := 'BELONGYM>=''' + sI_BelongYM + ''''
    else if sI_FieldName = 'OtherMonth' then
      I_Source.Filter := 'BELONGYM>''' + sI_BelongYM + ''''
    else
      I_Source.Filter := 'BELONGYM=''' + sI_BelongYM + '''';

    I_Source.Filtered := true;

    I_Source.First;
    while not I_Source.Eof do
    begin
      if not bI_IsIntAccData then  //�B�z�D���ФH���
      begin
        fL_TotalWorkerComm := fL_TotalWorkerComm + I_Source.FieldByName('WorkerComm').AsFloat;
        fL_TotalClctComm := fL_TotalClctComm + I_Source.FieldByName('ClctComm').AsFloat;
        fL_TotalComm := fL_TotalComm + I_Source.FieldByName('TotalComm').AsFloat;
      end
      else
        fL_TotalIntAccComm := fL_TotalIntAccComm + I_Source.FieldByName('IntAccComm').AsFloat;

      I_Source.Next;
    end;

    with I_Data do
    begin

      if not bI_IsIntAccData then  //�B�z�D���ФH���
      begin
        if Locate('MODE','WorkerComm',[]) then
        begin
          Edit;
          FieldByName(sI_FieldName).AsFloat := fL_TotalWorkerComm;
          Post;
        end;


        if Locate('MODE','ClctComm',[]) then
        begin
          Edit;
          FieldByName(sI_FieldName).AsFloat := fL_TotalClctComm;
          Post;
        end;

        if Locate('MODE','TotalComm',[]) then
        begin
          Edit;
          FieldByName(sI_FieldName).AsFloat := fL_TotalComm;
          Post;
        end;

      end
      else  //�B�z���ФH���
      begin
        if sI_BoxOrChannel = BOX_DATA then
        begin
          if Locate('MODE','BoxIntAccComm',[]) then
          begin
            Edit;
            FieldByName(sI_FieldName).AsFloat := fL_TotalIntAccComm;
            Post;
          end;
        end
        else
        begin
          if Locate('MODE','ChIntAccComm',[]) then
          begin
            Edit;
            FieldByName(sI_FieldName).AsFloat := fL_TotalIntAccComm;
            Post;
          end;
        end;
      end;
    end;
end;

function TdtmMain1.addZero(sI_SourceStr: String;
  nI_MaxLen: Integer): String;
var
    nL_OldLen,nL_Diff,ii : Integer;
    sL_ReturnStr : String;
begin
    nL_OldLen := Length(sI_SourceStr);
    nL_Diff := nI_MaxLen - nL_OldLen;

    for ii:=0 to nL_Diff-1 do
    begin
      if ii=0 then
        sL_ReturnStr := '0' + sI_SourceStr
      else
        sL_ReturnStr := '0' + sL_ReturnStr;
    end;

    if nL_Diff = 0 then
      sL_ReturnStr := sI_SourceStr;

    Result := sL_ReturnStr;
end;

procedure TdtmMain1.addToCdsTotalComm2(I_Source, I_Data: TDataSet;
  sI_BelongYM, sI_FieldName, sI_Mode: String);
var
    fL_TotalWorkerCount,fL_TotalClctCount,fL_TotalIntAccCount,fL_Count : Double;
begin
    fL_TotalWorkerCount := 0;
    fL_TotalClctCount := 0;
    fL_TotalIntAccCount := 0;

    I_Source.Filtered := false;

    if sI_FieldName = 'LastPay' then
      I_Source.Filter := 'BELONGYM<''' + sI_BelongYM + ''''
    else if sI_FieldName = 'NextPay' then
      I_Source.Filter := 'BELONGYM>=''' + sI_BelongYM + ''''
    else if sI_FieldName = 'OtherMonth' then
      I_Source.Filter := 'BELONGYM>''' + sI_BelongYM + ''''
    else
      I_Source.Filter := 'BELONGYM=''' + sI_BelongYM + '''';

    I_Source.Filtered := true;

    I_Source.First;
    while not I_Source.Eof do
    begin
      if sI_Mode = 'WorkerCount' then
        fL_TotalWorkerCount := fL_TotalWorkerCount + I_Source.FieldByName('WorkerCount').AsFloat
      else if sI_Mode = 'ClctCount' then
        fL_TotalClctCount := fL_TotalClctCount + I_Source.FieldByName('ClctCount').AsFloat
      else if (sI_Mode = 'BoxIntAccCount') or (sI_Mode = 'ChIntAccCount') then
        fL_TotalIntAccCount := fL_TotalIntAccCount + I_Source.FieldByName('IntAccCount').AsFloat;


      I_Source.Next;
    end;

    with I_Data do
    begin
      if sI_Mode = 'WorkerCount' then
        fL_Count := fL_TotalWorkerCount
      else if sI_Mode = 'ClctCount' then
        fL_Count := fL_TotalClctCount
      else if (sI_Mode = 'BoxIntAccCount') or (sI_Mode = 'ChIntAccCount') then
        fL_Count := fL_TotalIntAccCount;

      if Locate('MODE',sI_Mode,[]) then
      begin
        Edit;
        FieldByName(sI_FieldName).AsFloat := fL_Count;
        Post;
      end;
    end;

    //���[�`
    if sI_FieldName = 'LastPay' then
      fG_TotalLastPay := fG_TotalLastPay + fL_Count
    else if sI_FieldName = 'NextPay' then
      fG_TotalNextPay := fG_TotalNextPay + fL_Count
    else if sI_FieldName = 'OtherMonth' then
      fG_TotalOtherMonth := fG_TotalOtherMonth + fL_Count

end;



procedure TdtmMain1.cdsTotalMonthDataInitial(sI_Year : String);
var
    sL_Last2Year,sL_LastYear,sL_Year,sL_YearMonth,sL_Month :  String;
    ii : Integer;
begin
    sL_Last2Year := IntToStr(StrToInt(Trim(sI_Year))-2);
    sL_LastYear := IntToStr(StrToInt(Trim(sI_Year))-1);

    with cdsTotalMonthData do
    begin
      Append;
      FieldByName('ComputeYM').AsString := sL_Last2Year + '�~��';
      Post;

      Append;
      FieldByName('ComputeYM').AsString := sL_LastYear + '�~��';
      Post;

      sL_Year := IntToStr(StrToInt(Trim(sI_Year)));
      for ii:=1 to 12 do
      begin
        sL_Month := addZero(IntToStr(ii),2);
        sL_YearMonth := sL_Year + '/' + sL_Month;

        Append;
        FieldByName('ComputeYM').AsString := sL_YearMonth;
        Post;
      end;

      Append;
      FieldByName('ComputeYM').AsString := '�`�p';
      Post;

    end;

end;


procedure TdtmMain1.CalculateTotalMonthData(sI_CompCode, sI_Year: String);
var
    sL_SQL,sL_SqlComputeYM,sL_Month,sL_CdsComputeYM : String;
    ii : Integer;
    sL_Last2Year,sL_LastYear,sL_SqlBelongYM,sL_Year : String;
begin
    sL_Year := Trim(sI_Year);
    //����J�򥻸��
    cdsTotalMonthDataInitial(sL_Year);

    //�p��W�W�@�~��
    sL_SqlBelongYM := addZero(sL_Year,3) + '01';
    sL_Last2Year := addZero(IntToStr(StrToInt(sL_Year)-2),3);
    sL_SQL := 'select BelongYM,To_Char(sum(nvl(INTACCCOMM,0))) IntAccComm from so122 ' +
              'WHERE ComputeYM LIKE ''' + sL_Last2Year + '%'' AND BILLNO IS NOT NULL ' +
              ' AND BelongYM>=''' + sL_SqlBelongYM + ''' group by BELONGYM';

    with cdsComm1 do
    begin
      Close;
      CommandText := sL_SQL;
      Open;
    end;

    sL_CdsComputeYM := IntToStr(StrToInt(sL_Year)-2) + '�~��';
    addToCdsTotalMonthData(cdsComm1,cdsTotalMonthData,sL_CdsComputeYM,sL_Year);

    //�p��W�@�~��
    sL_SqlBelongYM := addZero(sL_Year,3) + '01';
    sL_LastYear := addZero(IntToStr(StrToInt(sL_Year)-1),3);
    sL_SQL := 'select BelongYM,To_Char(sum(nvl(INTACCCOMM,0))) IntAccComm from so122 ' +
              'WHERE ComputeYM LIKE ''' + sL_LastYear + '%'' AND BILLNO IS NOT NULL ' +
              ' AND BelongYM>=''' + sL_SqlBelongYM + ''' group by BELONGYM';

    with cdsComm1 do
    begin
      Close;
      CommandText := sL_SQL;
      Open;
    end;

    sL_CdsComputeYM := IntToStr(StrToInt(sL_Year)-1) + '�~��';
    addToCdsTotalMonthData(cdsComm1,cdsTotalMonthData,sL_CdsComputeYM,sL_Year);


    //�p�⥻�~��
    for ii:=1 to 12 do
    begin
      sL_Month := addZero(IntToStr(ii),2);
      sL_SqlComputeYM := sL_Year + sL_Month;

      sL_SQL := 'select BelongYM,To_Char(sum(nvl(INTACCCOMM,0))) IntAccComm from so122 ' +
                'WHERE ComputeYM=''' + sL_SqlComputeYM + ''' AND BILLNO IS NOT NULL group by BELONGYM';

      with cdsComm1 do
      begin
        Close;
        CommandText := sL_SQL;
        Open;
      end;

      sL_CdsComputeYM := IntToStr(StrToInt(sL_Year)) + '/' + sL_Month;
      addToCdsTotalMonthData(cdsComm1,cdsTotalMonthData,sL_CdsComputeYM,sL_Year);
    end;

    //�p��[�`
    computeCdsTotalMonthDataTotalAmount(cdsComm1,cdsTotalMonthData,sI_Year);
end;

procedure TdtmMain1.addToCdsTotalMonthData(I_Source, I_Data: TDataSet;
  sI_CdsComputeYM,sI_Year : String);
var
    ii,nL_StartMonth : Integer;
    sL_Month,sL_FieldName,sL_FieldYM : String;
    fL_Comm,fL_OtherComm : Double;
    fL_TotalOtherComm,fL_GroupTotalComm : Double;
begin
    //�p����u���Q�G�Ӥ���B
    I_Source.Filtered := false;

    for ii:=1 to 12 do
    begin
      sL_Month := addZero(IntToStr(ii),2);
      sL_FieldYM := sI_Year + sL_Month;
      sL_FieldName := 'Month' + IntToStr(ii);

      if I_Source.Locate('BelongYM', sL_FieldYM,[]) then
      begin
        if I_Data.Locate('ComputeYM', sI_CdsComputeYM,[]) then
        begin
          I_Data.Edit;
          fL_Comm := I_Data.FieldByName(sL_FieldName).AsFloat;
          fL_GroupTotalComm := fL_GroupTotalComm +  StrToFloat(I_Source.FieldByName('IntAccComm').AsString);
          I_Data.FieldByName(sL_FieldName).AsFloat := fL_Comm +  StrToFloat(I_Source.FieldByName('IntAccComm').AsString);
          I_Data.Post;
        end;
      end;
    end;


    //�B�z���~�ץH�᪺���
    I_Source.Filtered := false;
    I_Source.Filter := 'BelongYM>' + sL_FieldYM;
    I_Source.Filtered := true;

    fL_TotalOtherComm := 0;
    I_Source.First;
    while not I_Source.Eof do
    begin
      fL_TotalOtherComm := fL_TotalOtherComm + StrToFloat(I_Source.FieldByName('IntAccComm').AsString);
      fL_GroupTotalComm := fL_GroupTotalComm + fL_TotalOtherComm;
      I_Source.Next;
    end;

    if I_Data.Locate('ComputeYM', sI_CdsComputeYM,[]) then
    begin
      I_Data.Edit;
      I_Data.FieldByName('OtherMonth').AsFloat := fL_TotalOtherComm;
      I_Data.FieldByName('TotalComm').AsFloat := fL_GroupTotalComm;
      I_Data.Post;
    end;
end;

procedure TdtmMain1.computeCdsTotalMonthDataTotalAmount(I_Source,I_Data :  TDataSet;
  sI_Year: String);
var
    fL_TotalOtherMonth,fL_TotalComm : Double;
    fL_TotalMonth1,fL_TotalMonth2,fL_TotalMonth3 : Double;
    fL_TotalMonth4,fL_TotalMonth5,fL_TotalMonth6 : Double;
    fL_TotalMonth7,fL_TotalMonth8,fL_TotalMonth9 : Double;
    fL_TotalMonth10,fL_TotalMonth11,fL_TotalMonth12 : Double;
begin
    fL_TotalOtherMonth := 0;
    fL_TotalComm := 0;
    fL_TotalMonth1 := 0;
    fL_TotalMonth2 := 0;
    fL_TotalMonth3 := 0;
    fL_TotalMonth4 := 0;
    fL_TotalMonth5 := 0;
    fL_TotalMonth6 := 0;
    fL_TotalMonth7 := 0;
    fL_TotalMonth8 := 0;
    fL_TotalMonth9 := 0;
    fL_TotalMonth10 := 0;
    fL_TotalMonth11 := 0;
    fL_TotalMonth12 := 0;

    with I_Data do
    begin
      First;
      while not Eof do
      begin
        fL_TotalMonth1 := fL_TotalMonth1 + FieldByName('Month1').AsFloat;
        fL_TotalMonth2 := fL_TotalMonth2 + FieldByName('Month2').AsFloat;
        fL_TotalMonth3 := fL_TotalMonth3 + FieldByName('Month3').AsFloat;

        fL_TotalMonth4 := fL_TotalMonth4 + FieldByName('Month4').AsFloat;
        fL_TotalMonth5 := fL_TotalMonth5 + FieldByName('Month5').AsFloat;
        fL_TotalMonth6 := fL_TotalMonth6 + FieldByName('Month6').AsFloat;

        fL_TotalMonth7 := fL_TotalMonth7 + FieldByName('Month7').AsFloat;
        fL_TotalMonth8 := fL_TotalMonth8 + FieldByName('Month8').AsFloat;
        fL_TotalMonth9 := fL_TotalMonth9 + FieldByName('Month9').AsFloat;

        fL_TotalMonth10 := fL_TotalMonth10 + FieldByName('Month10').AsFloat;
        fL_TotalMonth11 := fL_TotalMonth11 + FieldByName('Month11').AsFloat;
        fL_TotalMonth12 := fL_TotalMonth12 + FieldByName('Month12').AsFloat;

        fL_TotalOtherMonth := fL_TotalOtherMonth + FieldByName('OtherMonth').AsFloat;
        fL_TotalComm := fL_TotalComm + FieldByName('TotalComm').AsFloat;

        Next;
      end;


      if Locate('ComputeYM', '�`�p',[]) then
      begin
        Edit;
        FieldByName('Month1').AsFloat := fL_TotalMonth1;
        FieldByName('Month2').AsFloat := fL_TotalMonth2;
        FieldByName('Month3').AsFloat := fL_TotalMonth3;

        FieldByName('Month4').AsFloat := fL_TotalMonth4;
        FieldByName('Month5').AsFloat := fL_TotalMonth5;
        FieldByName('Month6').AsFloat := fL_TotalMonth6;

        FieldByName('Month7').AsFloat := fL_TotalMonth7;
        FieldByName('Month8').AsFloat := fL_TotalMonth8;
        FieldByName('Month9').AsFloat := fL_TotalMonth9;

        FieldByName('Month10').AsFloat := fL_TotalMonth10;
        FieldByName('Month11').AsFloat := fL_TotalMonth11;
        FieldByName('Month12').AsFloat := fL_TotalMonth12;

        FieldByName('OtherMonth').AsFloat := fL_TotalOtherMonth;
        FieldByName('TotalComm').AsFloat := fL_TotalComm;

        Post;
      end;
    end;

end;

procedure TdtmMain1.showTotalMonthDataExcel(sI_Year: String);
var
    sL_FileName,sL_QueryString,sL_Year : String;
    sL_FileCurrDateTime,sL_CurrDateTime : String;
    L_StrList : TStringList;
    sL_ComputeYM, value: String;
    aPos: Integer;
begin
    //L_StrList := TStringList.Create;
    //cdsTotalMonthData.FieldDefs.GetItemNames(L_StrList);

    sL_Year := IntToStr(StrToInt(sI_Year)+1911-1911);

    cdsTotalMonthData.FieldByName('ComputeYM').DisplayLabel := '�ꦬ\�o��';
    cdsTotalMonthData.FieldByName('Month1').DisplayLabel := sL_Year + '/01';
    cdsTotalMonthData.FieldByName('Month2').DisplayLabel := sL_Year + '/02';
    cdsTotalMonthData.FieldByName('Month3').DisplayLabel := sL_Year + '/03';

    cdsTotalMonthData.FieldByName('Month4').DisplayLabel := sL_Year + '/04';
    cdsTotalMonthData.FieldByName('Month5').DisplayLabel := sL_Year + '/05';
    cdsTotalMonthData.FieldByName('Month6').DisplayLabel := sL_Year + '/06';

    cdsTotalMonthData.FieldByName('Month7').DisplayLabel := sL_Year + '/07';
    cdsTotalMonthData.FieldByName('Month8').DisplayLabel := sL_Year + '/08';
    cdsTotalMonthData.FieldByName('Month9').DisplayLabel := sL_Year + '/09';

    cdsTotalMonthData.FieldByName('Month10').DisplayLabel := sL_Year + '/10';
    cdsTotalMonthData.FieldByName('Month11').DisplayLabel := sL_Year + '/11';
    cdsTotalMonthData.FieldByName('Month12').DisplayLabel := sL_Year + '/12';

    cdsTotalMonthData.FieldByName('OtherMonth').DisplayLabel := IntToStr(StrToInt(sL_Year)+1) + '�H��';
    cdsTotalMonthData.FieldByName('TotalComm').DisplayLabel := '�p�p';


    sL_FileCurrDateTime := DateTimeToStr(now);
    sL_CurrDateTime := ReplaceStr(Trim(ReplaceStr(sL_FileCurrDateTime,'/')),':');

    sL_FileName := sG_ExcelSavePath + '\' + sL_Year + '�~�ץI�O�W�D���Ц����u����_' + frmMainMenu.sG_CompCode + '_' + sL_CurrDateTime + '.XLS';
    sL_QueryString:='����W��:' + sL_Year + ' �~�ץI�O�W�D���Ц����u����' + '@' +
                    '�έp�ɶ�:' + sL_FileCurrDateTime + '@' + '����H��:' + frmMainMenu.sG_User + '@' +
                    '���q�W��:' + TUstr.replaceStr(frmMainMenu.sG_CompName,'�@','��') + '@' +
                    '���:��' ;


    cdsTotalMonthData.DisableControls;
    try
      cdsTotalMonthData.First;
      while not cdsTotalMonthData.Eof do
      begin
        sL_ComputeYM := cdsTotalMonthData.FieldByName('ComputeYM').AsString;
        value := EmptyStr;
        if ( AnsiPos( '�~��', sL_ComputeYM ) > 0 ) then
        begin
          value := Copy( sL_ComputeYM, 1, AnsiPos( '�~��', sL_ComputeYM ) - 1 );
          value := IntToStr( StrToInt( value ) + 1911 - 1911 ) + '�~��';
        end else
        if ( AnsiPos( '�`�p', sL_ComputeYM ) > 0 ) then
        begin
          value := sL_ComputeYM;
        end else
        begin
         aPos := Pos( '/',  sL_ComputeYM );
         value := IntToStr( StrToInt( Copy( sL_ComputeYM, 1, aPos - 1 ) ) + 1911 - 1911 ) + '/' +
           Copy( sL_ComputeYM, aPos + 1, 2 );
        end;
        cdsTotalMonthData.Edit;
        cdsTotalMonthData.FieldByName( 'ComputeYM' ).AsString := value;
        cdsTotalMonthData.Post;
        cdsTotalMonthData.Next;
      end;
      cdsTotalMonthData.First;
      DataSetToXLS(cdsTotalMonthData,sL_FileName,sL_QueryString,'');
    finally
      cdsTotalMonthData.EnableControls;
    end;  

    MessageDlg('�p�⧹��,�w�I�u����s��   ' + sL_FileName,mtInformation,[mbOK],0)
end;

procedure TdtmMain1.addColumnZero(I_CDS: TClientDataSet);
begin
    //�N�ƶq���O NULL ���� 0

    with I_CDS do
    begin
      First;
      while not Eof do
      begin
        Edit;
        if FieldByName('WorkerCounts').AsString = '' then
          FieldByName('WorkerCounts').AsInteger := 0;

        if FieldByName('ClctCounts').AsString = '' then
          FieldByName('ClctCounts').AsInteger := 0;

        if FieldByName('BoxIntAccCounts').AsString = '' then
          FieldByName('BoxIntAccCounts').AsInteger := 0;

        if FieldByName('ChIntAccCounts').AsString = '' then
          FieldByName('ChIntAccCounts').AsInteger := 0;

        Post;
        Next;
      end;
    end;
end;

function TdtmMain1.getSo125ParamTemp(sI_CompCode: String): Boolean;
var
    sL_SQL,sL_TempExcelSavePath : String;
    L_TempList : TStringList;
    ii : Integer;
begin
    sL_SQL := 'SELECT * FROM SO125 WHERE COMPCODE=' + sI_CompCode;

    with qryCom do
    begin
      Close;
      qryCom.SQL.CommaText := sL_SQL;
      Open;


      //�p�G�S���,�������w�]��0
      if RecordCount = 0 then
      begin
        DisableControls;
        if State = dsInactive then
            Active := True;

        //�L�����q���ѼƥN�X
        Result  := false;
      end
      else
      begin
        //���_STB�����ФH�h�֦���
        nG_BuyBoxIntrComm := FieldByName('Param1').AsInteger;

        //����STB�����ФH�h�֦���
        nG_RentBoxIntrComm := FieldByName('Param2').AsInteger;

        //�w��STB���u�{�H���h�֦���
        nG_SetBoxWorkerComm := FieldByName('Param3').AsInteger;

        //���_STB�����z�H(�ȪA)�h�֦���
        nG_BuyBoxAcceptComm := FieldByName('Param4').AsInteger;

        //����STB�����z�H(�ȪA)�h�֦���
        nG_RentBoxAcceptComm := FieldByName('Param5').AsInteger;

        //�Ȥ�ۨӹq�q���W�D�����s�H(�ȪA)�h�֦���
        nG_SelfOrderChannelAcceptComm := FieldByName('Param6').AsInteger;

        //STB�ݯ����h�֤Ѥ~������
        nG_RentBoxUseDays := FieldByName('Param7').AsInteger;

        //STB�R�_���ϥκ��h�֤Ѥ~�����s�H����
        nG_BuyBoxUseDays := FieldByName('Param8').AsInteger;

        //��������X��
        nG_LimitPeriod := FieldByName('Param9').AsInteger;

        //�@�����I������Ū��??��e�����h�O���
        nG_BackMonth := FieldByName('Param10').AsInteger;


        //BOX�ݱư����~�Ȭ���
        sG_PromCodeAndName := FieldByName('Param12').AsString;

        L_TempList := TUStr.ParseStrings(sG_PromCodeAndName,';');

        G_PromCodeStrList.Clear;
        G_PromNameStrList.Clear;

        if sG_PromCodeAndName <> '' then
        begin
          G_PromCodeStrList := TUStr.ParseStrings(L_TempList.Strings[0],',');
          G_PromNameStrList := TUStr.ParseStrings(L_TempList.Strings[1],',');
        end;

        L_TempList.Free;

        //Excel�����x�s���|(�w�]�bC:\)
        sL_TempExcelSavePath := FieldByName('Param11').AsString;

        if sL_TempExcelSavePath = '' then
          sG_ExcelSavePath := 'C:'
        else
        begin
          if (sL_TempExcelSavePath = 'C:\') OR (sL_TempExcelSavePath = 'D:\') then
            sG_ExcelSavePath := TUstr.replaceStr(sL_TempExcelSavePath,'\','')
          else
            sG_ExcelSavePath := sL_TempExcelSavePath;
        end;

        Edit;
        FieldByName('Param11').AsString := sG_ExcelSavePath;
        Post;
        {
        for ii:=0 to G_PromCodeStrList.Count-1 do
        begin
          showmessage('getSo125Param== ' + G_PromCodeStrList.Strings[ii] + '-' + G_PromNameStrList.Strings[ii]);
        end;
        }

        Result  := true;
      end;
    end;


end;

end.


