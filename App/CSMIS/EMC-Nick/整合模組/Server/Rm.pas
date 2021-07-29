unit Rm;

{$WARN SYMBOL_PLATFORM OFF}

interface

uses
  Windows, Messages, SysUtils, Classes, ComServ, ComObj, VCLCom, DataBkr,
  DBClient, Emc_Separate_TLB, StdVcl, Provider, DB, DBTables,Variants,
  ADODB, Dialogs, Forms;

type
  TEmc_Separate1 = class(TRemoteDataModule, IEmc_Separate1)
    dspSO110: TDataSetProvider;
    dspSO111D: TDataSetProvider;
    dspSO112M: TDataSetProvider;
    dspCd024CT: TDataSetProvider;
    dspSqlStatement: TDataSetProvider;
    dspSo113CT: TDataSetProvider;
    dspSqlDelete: TDataSetProvider;
    dspSo113: TDataSetProvider;
    dspCD039: TDataSetProvider;
    dspSo114N: TDataSetProvider;
    dspCodeCd039: TDataSetProvider;
    dspParam: TDataSetProvider;
    dspCd019: TDataSetProvider;
    dspChargeInfoForTally: TDataSetProvider;
    dspPackageId: TDataSetProvider;
    dspSearSo112: TDataSetProvider;
    qryCodeCD039A: TADOQuery;
    dsoCodeCD039A: TDataSetProvider;
    ADOConnection1: TADOConnection;
    qryCom: TADOQuery;
    dspCom: TDataSetProvider;
    qryRefData: TADOQuery;
    dspRefData: TDataSetProvider;
    qryRefDataA: TADOQuery;
    dspRefDataA: TDataSetProvider;
    qryChargeData: TADOQuery;
    dspChargeData: TDataSetProvider;
    qryCustData: TADOQuery;
    dspCustData: TDataSetProvider;
    qrySO115: TADOQuery;
    dspSO115: TDataSetProvider;
    qrySO114: TADOQuery;
    dspSO114: TDataSetProvider;
    qrySO116: TADOQuery;
    dspSO116: TDataSetProvider;
    qrySO117: TADOQuery;
    dspSO117: TDataSetProvider;
    dspSO110Look: TDataSetProvider;
    dspSo113Look: TDataSetProvider;
    dspCd024Look: TDataSetProvider;
    qrySO110: TADOQuery;
    qrySO112M: TADOQuery;
    qryCd024CT: TADOQuery;
    qrySo113CT: TADOQuery;
    qrySqlStatement: TADOQuery;
    qrySqlDelete: TADOQuery;
    qryChargeInfoForTally: TADOQuery;
    qrySearSo112: TADOQuery;
    qrySO113: TADOQuery;
    qryCD039: TADOQuery;
    qrySo114N: TADOQuery;
    qryCodeCd039: TADOQuery;
    qryParam: TADOQuery;
    qryCd019: TADOQuery;
    qryPackageId: TADOQuery;
    qrySo113Look: TADOQuery;
    qryCd024Look: TADOQuery;
    qrySO115B: TADOQuery;
    dspSO115B: TDataSetProvider;
    qryPriv: TADOQuery;
    cdsPriv: TDataSetProvider;
    qrySO111D: TADOQuery;
    qrySO110Look: TADOQuery;
    qrySO114B: TADOQuery;
    dspSO114B: TDataSetProvider;
    qryTallyRefData1: TADOQuery;
    qryTallyRefData2: TADOQuery;
    dspTallyRefData1: TDataSetProvider;
    dspTallyRefData2: TDataSetProvider;
    qrySO119: TADOQuery;
    dspSO119: TDataSetProvider;
    qryTempSO119: TADOQuery;
    dspTempSO119: TDataSetProvider;
    qrySO114C: TADOQuery;
    dspSO114C: TDataSetProvider;
    qryCodeCD019: TADOQuery;
    qryCodeCD019CODENO: TIntegerField;
    qryCodeCD019DESCRIPTION: TStringField;
    dspCodeCD019: TDataSetProvider;
    qrySO033: TADOQuery;
    dspSO033: TDataSetProvider;
    qrySO131: TADOQuery;
    dspSO131: TDataSetProvider;
    spSO131: TADOStoredProc;
    dspSpSO131: TDataSetProvider;
    qryCd019A: TADOQuery;
    dspCd019A: TDataSetProvider;
    qrySo112A: TADOQuery;
    dspSo112A: TDataSetProvider;
    qrySO131Excel: TADOQuery;
    dspSO131Excel: TDataSetProvider;
    qryOtherSO131Excel: TADOQuery;
    dspOtherSO131Excel: TDataSetProvider;
    procedure RemoteDataModuleCreate(Sender: TObject);
    procedure qryCd024CT_1AfterOpen(DataSet: TDataSet);
    procedure qrySO110AfterOpen(DataSet: TDataSet);
    procedure qrySO112M_1AfterOpen(DataSet: TDataSet);
    procedure qrySO111D_1AfterOpen(DataSet: TDataSet);
    procedure qrySo113CT_1AfterOpen(DataSet: TDataSet);
    procedure qrySqlStatement_1AfterOpen(DataSet: TDataSet);
    procedure RemoteDataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
    G_AdoConn : TADOConnection;
  protected
    class procedure UpdateRegistry(Register: Boolean; const ClassID, ProgID: string); override;
    procedure ApplyUpdates(vG_CustVar: OleVariant); safecall;
    procedure connectToDB(sI_User, sI_CompCode, sI_DbUserName, sI_DbPassword,
      sI_DbAlias: OleVariant; var sI_Result: OleVariant); safecall;
    procedure runSF_CALCULATESO131(P_COMPCODE, P_SERVICETYPE, P_COMPUTEYM,
      P_STARTDATE, P_ENDDATE, P_CHARGEITEMSQL,
      P_REALORSHOULDDATE: OleVariant; var p_RetCode, p_RetMsg: OleVariant);
      safecall;
    procedure getSO131Data(sI_SQL: OleVariant); safecall;


  public
    { Public declarations }
  end;

implementation

uses Unit1;

{$R *.DFM}

class procedure TEmc_Separate1.UpdateRegistry(Register: Boolean; const ClassID, ProgID: string);
begin
  if Register then
  begin
    inherited UpdateRegistry(Register, ClassID, ProgID);
    EnableSocketTransport(ClassID);
    EnableWebTransport(ClassID);
  end else
  begin
    DisableSocketTransport(ClassID);
    DisableWebTransport(ClassID);
    inherited UpdateRegistry(Register, ClassID, ProgID);
  end;
end;

procedure TEmc_Separate1.ApplyUpdates(vG_CustVar: OleVariant);

begin

end;

procedure TEmc_Separate1.RemoteDataModuleCreate(Sender: TObject);
begin
     //frmMain.UpdateClientCount(1);

end;

procedure TEmc_Separate1.qryCd024CT_1AfterOpen(DataSet: TDataSet);
begin
    //frmMain.IncQueryCount;
end;

procedure TEmc_Separate1.qrySO110AfterOpen(DataSet: TDataSet);
begin
    //frmMain.IncQueryCount;
end;

procedure TEmc_Separate1.qrySO112M_1AfterOpen(DataSet: TDataSet);
begin
    //frmMain.IncQueryCount;
end;

procedure TEmc_Separate1.qrySO111D_1AfterOpen(DataSet: TDataSet);
begin
    //frmMain.IncQueryCount;
end;

procedure TEmc_Separate1.qrySo113CT_1AfterOpen(DataSet: TDataSet);
begin
    //frmMain.IncQueryCount;
end;

procedure TEmc_Separate1.qrySqlStatement_1AfterOpen(DataSet: TDataSet);
begin
    //frmMain.IncQueryCount;
end;

procedure TEmc_Separate1.RemoteDataModuleDestroy(Sender: TObject);
begin
    //frmMain.UpdateClientCount(-1);
end;



procedure TEmc_Separate1.connectToDB(sI_User, sI_CompCode, sI_DbUserName,
  sI_DbPassword, sI_DbAlias: OleVariant; var sI_Result: OleVariant);
var
    sL_DbConnStr : String;
begin
    try

        sL_DbConnStr := 'Provider=MSDAORA.1;Password=' + sI_DbPassword + ';User ID='+sI_DbUserName+';Data Source='+sI_DbAlias; //適用於其他 PC

        G_AdoConn := TADOConnection.Create(Application);
        G_AdoConn.ConnectionString := sL_DbConnStr;


        if ADOConnection1.Connected then
          ADOConnection1.Connected := false;

        ADOConnection1.ConnectionString := sL_DbConnStr;
        ADOConnection1.Connected := true;


        sI_Result := '';
    except
        sI_Result := '-1: 與CC&B資料庫連線失敗, 資料庫別名=<'+sI_DbAlias+'>, DB User=<'+sI_DbUserName +'>';
        exit;
    end;



end;

procedure TEmc_Separate1.runSF_CALCULATESO131(P_COMPCODE, P_SERVICETYPE,
  P_COMPUTEYM, P_STARTDATE, P_ENDDATE, P_CHARGEITEMSQL,
  P_REALORSHOULDDATE: OleVariant; var p_RetCode, p_RetMsg: OleVariant);
var
    ii : Integer;  
begin

    with spSO131 do
    begin
      Parameters.ParamByName('P_COMPCODE').Value := P_COMPCODE;
      Parameters.ParamByName('P_SERVICETYPE').Value := P_SERVICETYPE;
      Parameters.ParamByName('P_COMPUTEYM').Value := P_COMPUTEYM;
      Parameters.ParamByName('P_STARTDATE').Value := P_STARTDATE;
      Parameters.ParamByName('P_ENDDATE').Value := P_ENDDATE;
      Parameters.ParamByName('P_CHARGEITEMSQL').Value := P_CHARGEITEMSQL;
      Parameters.ParamByName('P_REALORSHOULDDATE').Value := P_REALORSHOULDDATE;


      ExecProc;

      //fI_FunReturn := Params.Items[0].Value;
      p_RetCode := Parameters.Items[8].Value;
      p_RetMsg := Parameters.Items[9].Value;

      if p_RetMsg = null then
        p_RetMsg := '';
    end;


end;

procedure TEmc_Separate1.getSO131Data(sI_SQL: OleVariant);
begin
    qrySO131.Close;
    qrySO131.SQL.Clear;
    qrySO131.SQL.Add(sI_SQL);
    qrySO131.Open;
end;

initialization
  TComponentFactory.Create(ComServer, TEmc_Separate1,
    Class_Emc_Separate1, ciMultiInstance, tmApartment);
end.
