unit dtmMainU;

interface

uses
  SysUtils, Classes, DB, ADODB, DBClient, Provider,Dialogs;


const
    DEFAULT_DB_OWNER_NAME = 'emcyms';
    DEFAULT_DB_OWNER_PASSWORD = 'emcyms';

type
  TdtmMain = class(TDataModule)
    ADOConnection1: TADOConnection;
    qrySO030A: TADOQuery;
    dspSO030A: TDataSetProvider;
    cdsSO030A: TClientDataSet;
    qrySO001: TADOQuery;
    dspSO001: TDataSetProvider;
    cdsSO001: TClientDataSet;
    cdsCustIDCounts: TClientDataSet;
    qrySO509: TADOQuery;
    dspSO509: TDataSetProvider;
    cdsSO509: TClientDataSet;
    qrySO509B: TADOQuery;
    dspSO509B: TDataSetProvider;
    cdsSO509B: TClientDataSet;
    qrySO509A: TADOQuery;
    dspSO509A: TDataSetProvider;
    cdsSO509A: TClientDataSet;
    qryCom: TADOQuery;
    cdsSO509BCOMPCODE: TIntegerField;
    cdsSO509BCOMPNAME: TStringField;
    cdsSO509BTABLECODE: TStringField;
    cdsSO509BTABLENAME: TStringField;
    cdsSO509BMASTERCOMPUTEDATE: TDateTimeField;
    cdsSO509BMASTERRECORDCOUNTS: TIntegerField;
    cdsSO509BSNAPCOMPUTEDATE: TDateTimeField;
    cdsSO509BSNAPRECORDCOUNTS: TIntegerField;
    cdsSO509BDIFFCOUNTS: TIntegerField;
    cdsSO509BRESULTFLAG: TIntegerField;
    cdsSO509BSTRDIFFCOUNTS: TStringField;
    cdsTempSO509B: TClientDataSet;
    cdsTempSO509BCompName: TStringField;
    cdsTempSO509BTableName: TStringField;
    cdsTempSO509BMasterComputeDate: TDateTimeField;
    cdsTempSO509BMasterRecordCounts: TIntegerField;
    cdsTempSO509BSnapComputeDate: TDateTimeField;
    cdsTempSO509BSnapRecordCounts: TIntegerField;
    cdsTempSO509BDiffCounts: TIntegerField;
    cdsTempSO509BResultFlag: TStringField;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    G_CompCodeStrList : TStringList;
    function connectToDB(sI_Alias,sI_UserName,sI_Password : String):WideString;
    function getSnapShotLogData(sI_Date : String) : Boolean;
    procedure getSnapShotCustIDCounts;
    procedure getMoveToSnapShotTable;
    function getDbOwner(sI_CompCode : String) : String;
    function checkSO509BIsCompute(sI_DbOwner,sI_CompCode,sI_EngDate : String) : Boolean;
    procedure computeSO509B(sI_DbOwner,sI_CompCode,sI_CompName,sI_EngDate : String);


    procedure getSO509AData(sI_DbOwner,sI_EngDate : String);
    procedure getSO509BData(sI_DbOwner,sI_CompCode,sI_EngDate : String);
    function computeTableCounts(sI_DbOwner,sI_Table : String) : Integer;
    function getDiffCoutsResultFlag(nI_MasterTableCouts,nI_SnapTableCouts,nI_DiffCouts : Integer) : String;
    procedure computeSoAndSnapshotTableCounts(sI_DbOwner,sI_CompCode,sI_CompName,sI_EngDate : String);
    procedure deleteSO509B(sI_DbOwner,sI_CompCode,sI_EngDate : String);
    function getAllCompCode : TStringList;
    function getCompName(sI_CompCode : String) : String;
  end;

var
  dtmMain: TdtmMain;

implementation

uses Ustru, frmSnapShotU;

{$R *.dfm}

{ TDataModule1 }

function TdtmMain.checkSO509BIsCompute(sI_DbOwner, sI_CompCode,
  sI_EngDate: String): Boolean;
var
    sL_SQL : String;
begin
    sL_SQL := 'SELECT * FROM ' + sI_DbOwner + '.SO509B WHERE COMPCODE=' + sI_CompCode +
              ' AND MASTERCOMPUTEDATE BETWEEN ' + TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_EngDate + ' 00:00:01')) +
              ' AND ' + TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_EngDate + ' 23:59:59'));

    with qryCom do
    begin
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;

      if RecordCount > 0 then
        Result := true
      else
        Result := false;
    end;
end;

procedure TdtmMain.computeSO509B(sI_DbOwner, sI_CompCode,sI_CompName,
  sI_EngDate: String);
var
    sL_SQL,sL_MoveTableCode,sL_MoveTableName,sL_ResultFlag : String;
    nL_SnapTableCouts,nL_MasterTableCouts,nL_DiffCouts : Integer;
    dL_MasterComputeDate : TDateTime;
begin
    //取出 Master site 用來紀錄每天各 table 的資料筆數
    getSO509AData(sI_DbOwner,sI_EngDate);

    if cdsSO509A.RecordCount <> 0 then
    begin
      cdsSO509.First;
      while not cdsSO509.Eof do
      begin
        //Move 至 SnapShot 的 Table
        sL_MoveTableCode := cdsSO509.FieldByName('TableCode').AsString;
        sL_MoveTableName := cdsSO509.FieldByName('TableName').AsString;


        //計算 SnapShot 該 Table 的筆數
        nL_SnapTableCouts := computeTableCounts(sI_DbOwner,sL_MoveTableCode);


        //計算 SO 該 Table 的資料
        if cdsSO509A.Locate('TableCode',sL_MoveTableCode,[]) then
        begin
          nL_MasterTableCouts := cdsSO509A.FieldByName('RecordCounts').AsInteger;
          dL_MasterComputeDate := cdsSO509A.FieldByName('ComputeDate').AsDateTime;
        end;

        //原 SO 該 Table 筆數 - SnapShot 該 Table 筆數
        nL_DiffCouts := nL_MasterTableCouts - nL_SnapTableCouts;


        sL_ResultFlag := getDiffCoutsResultFlag(nL_MasterTableCouts,nL_SnapTableCouts,nL_DiffCouts);

        //將計算結果存至SO509B
        sL_SQL := 'INSERT INTO ' + sI_DbOwner + '.SO509B(CompCode,CompName,TableCode,TableName,' +
                  'MasterComputeDate,MasterRecordCounts,SnapComputeDate,' +
                  'SnapRecordCounts,DiffCounts,ResultFlag) VALUES(' + sI_CompCode +
                  ',''' + sI_CompName + ''',''' + sL_MoveTableCode + ''',''' +
                  sL_MoveTableName + ''',' + TUstr.getOracleSQLDateTimeStr(dL_MasterComputeDate) + ',' +
                  IntToStr(nL_MasterTableCouts) + ',' + TUstr.getOracleSQLDateTimeStr(frmSnapShot.dG_CurDateTime) +
                  ',' + IntToStr(nL_SnapTableCouts) + ',' + IntToStr(nL_DiffCouts) + ',' +
                  sL_ResultFlag + ')';

        with qryCom do
        begin
          SQL.Clear;
          SQL.Add(sL_SQL);
          ExecSQL;
        end;

        cdsSO509.Next;
      end;
    end;


    //取出 SO509B Data
    getSO509BData(sI_DbOwner, sI_CompCode,sI_EngDate);
end;

procedure TdtmMain.computeSoAndSnapshotTableCounts(sI_DbOwner, sI_CompCode,
  sI_CompName, sI_EngDate: String);
begin

    //檢查該公司該日期是否已經有被計算過,若未計算則計算之
    if not dtmMain.checkSO509BIsCompute(sI_DbOwner,sI_CompCode,sI_EngDate) then
    begin
      computeSO509B(sI_DbOwner,sI_CompCode,sI_CompName,sI_EngDate);
    end
    else
    begin
      //取出 SO509B Data
      dtmMain.getSO509BData(sI_DbOwner, sI_CompCode,sI_EngDate);
    end;

end;

function TdtmMain.computeTableCounts(sI_DbOwner,
  sI_Table: String): Integer;
var
    sL_SQL : String;
begin
    sL_SQL := 'SELECT COUNT(*) COUNTS FROM ' + sI_DbOwner + '.' + sI_Table;

    with qryCom do
    begin
      SQL.Clear;
      SQL.Add(sL_SQL);

      try
        Open;
        Result := FieldByName('Counts').AsInteger;
      except
        Result := 0;
      end;
    end;
end;

function TdtmMain.connectToDB(sI_Alias, sI_UserName,
  sI_Password: String): WideString;
var
    sL_DbConnStr : String;
begin
    try
      if not ADOConnection1.Connected then
      begin
        sL_DbConnStr := 'Provider=MSDAORA.1;Password=' + sI_Password + ';User ID='+sI_UserName+';Data Source='+sI_Alias+';Persist Security Info=True'; //適用於其他 PC
        ADOConnection1.ConnectionString := sL_DbConnStr;
        ADOConnection1.Connected := true;
      end;
    except
      result := '-1: 與CC&B資料庫連線失敗, 資料庫別名=<'+sI_Alias+'>, DB User=<'+sI_UserName +'>';
      exit;
    end;
    result := '';
end;


procedure TdtmMain.deleteSO509B(sI_DbOwner, sI_CompCode,
  sI_EngDate: String);
var
    sL_SQL : String;
begin
    sL_SQL := 'DELETE ' + sI_DbOwner + '.SO509B WHERE COMPCODE=' + sI_CompCode +
              ' AND MASTERCOMPUTEDATE BETWEEN ' + TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_EngDate + ' 00:00:01')) +
              ' AND ' + TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_EngDate + ' 23:59:59'));

    with qryCom do
    begin
      SQL.Clear;
      SQL.Add(sL_SQL);
      ExecSQL;
    end;
end;

function TdtmMain.getAllCompCode : TStringList;
var
    sL_CompCode : String;
begin
    G_CompCodeStrList.Clear;

    cdsSO030A.First;
    while not cdsSO030A.Eof do
    begin
      sL_CompCode := cdsSO030A.FieldByName('CompCode').AsString;
      if G_CompCodeStrList.IndexOf(sL_CompCode)= -1 then
        G_CompCodeStrList.Add(sL_CompCode);
      cdsSO030A.Next;
    end;
end;

function TdtmMain.getDbOwner(sI_CompCode: String): String;
var
    sL_DbOwner : String;
begin
    if sI_CompCode = '1' then
    begin
      //觀昇
      sL_DbOwner := 'EMCKS';
    end
    else if sI_CompCode = '2' then
    begin
      //屏南
      sL_DbOwner := 'EMCPN';
    end
    else if sI_CompCode = '3' then
    begin
      //南天
      sL_DbOwner := 'EMCNT';
    end
    else if sI_CompCode = '5' then
    begin
      //新頻道
      sL_DbOwner := 'EMCNCC';
      //showmessage('測試30');
      //sL_DbOwner := 'V30';
    end
    else if sI_CompCode = '6' then
    begin
      //豐盟
      sL_DbOwner := 'EMCFM';
    end
    else if sI_CompCode = '7' then
    begin
      //振道
      sL_DbOwner := 'EMCCT';
    end
    else if sI_CompCode = '8' then
    begin
      //全聯
      sL_DbOwner := 'EMCUC';
    end
    else if sI_CompCode = '9' then
    begin
      //陽明山
      sL_DbOwner := 'EMCYMS';
    end
    else if sI_CompCode = '10' then
    begin
      //新台北
      sL_DbOwner := 'EMCNTP';
    end
    else if sI_CompCode = '11' then
    begin
      //金頻道
      sL_DbOwner := 'EMCKP';
    end
    else if sI_CompCode = '12' then
    begin
      //大安文山
      sL_DbOwner := 'EMCDW';
    end
    else if sI_CompCode = '13' then
    begin
      //新唐城
      sL_DbOwner := 'EMCTC';
    end
    else if sI_CompCode = '14' then
    begin
      //大新店
      sL_DbOwner := 'EMCST';
    end
    else if sI_CompCode = '16' then
    begin
      //北桃園
      sL_DbOwner := 'EMCNTY';
    end;

    Result := sL_DbOwner;
end;

function TdtmMain.getDiffCoutsResultFlag(nI_MasterTableCouts,
  nI_SnapTableCouts, nI_DiffCouts: Integer): String;
begin
    {
    1: Master資料筆數等於Snapshot資料筆數
    2: Master資料筆數大於Snapshot資料筆數
    3: Master資料筆數小於Snapshot資料筆數
    4: Master資料筆數為0
    5: Snapshot資料筆數為0
    6: Master資料筆數與Snapshot資料筆數均為0
    }

    if (nI_MasterTableCouts=0) and (nI_SnapTableCouts=0) then
      Result := '6'
    else if (nI_MasterTableCouts<>0) and (nI_SnapTableCouts=0) then
      Result := '5'
    else if (nI_MasterTableCouts=0) and (nI_SnapTableCouts<>0) then
      Result := '4'
    else if nI_DiffCouts < 0 then
      Result := '3'
    else if nI_DiffCouts > 0 then
      Result := '2'
    else if nI_DiffCouts = 0 then
      Result := '1';

end;

procedure TdtmMain.getMoveToSnapShotTable;
var
    sL_SQL : String;
begin
    sL_SQL := 'SELECT * FROM ' + DEFAULT_DB_OWNER_NAME + '.SO509';

    with cdsSO509 do
    begin
      Close;
      CommandText := sL_SQL;
      Open;
    end;

end;

procedure TdtmMain.getSnapShotCustIDCounts;
var
    sL_TitleSql,sL_SQL,sL_Where : String;
begin
    sL_TitleSql := 'SELECT C.Description ClassName,B.Description CompName,Count(*) Counts from ';


    sL_Where := ' Where c.CodeNo=a.ClassCode1 and B.CODENO=A.CompCode Group by C.Description,B.Description';

//Jackal測試用
{
    //V30
    sL_SQL := sL_TitleSql + 'V30.SO001 A,V30.CD039 B,V30.CD004 C ' + sL_Where;

    //NCC
    sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'NCC.SO001 A,NCC.CD039 B,NCC.CD004 C ' + sL_Where + ')';
}


    //陽明山 (1)
    sL_SQL := sL_TitleSql + 'EMCYMS.SO001 A,EMCYMS.CD039 B,EMCYMS.CD004 C ' + sL_Where;

    //新台北(2)
    sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCNTP.SO001 A,EMCNTP.CD039 B,EMCNTP.CD004 C ' + sL_Where + ')';

    //大安文山(3)
    sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCDW.SO001 A,EMCDW.CD039 B,EMCDW.CD004 C ' + sL_Where + ')';

    //金頻道(4)
    sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCKP.SO001 A,EMCKP.CD039 B,EMCKP.CD004 C ' + sL_Where + ')';

    //振道(5)
    sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCCT.SO001 A,EMCCT.CD039 B,EMCCT.CD004 C ' + sL_Where + ')';

    //豐盟(6)
    sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCFM.SO001 A,EMCFM.CD039 B,EMCFM.CD004 C ' + sL_Where + ')';

    //新頻道(7)
    sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCNCC.SO001 A,EMCNCC.CD039 B,EMCNCC.CD004 C ' + sL_Where + ')';

    //南天(8)
    sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCNT.SO001 A,EMCNT.CD039 B,EMCNT.CD004 C ' + sL_Where + ')';

    //觀昇(9)
    sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCKS.SO001 A,EMCKS.CD039 B,EMCKS.CD004 C ' + sL_Where + ')';

    //屏南(10)
    sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCPN.SO001 A,EMCPN.CD039 B,EMCPN.CD004 C ' + sL_Where + ')';

    //新唐城(11)
    sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCTC.SO001 A,EMCTC.CD039 B,EMCTC.CD004 C ' + sL_Where + ')';

    //全聯(12)
    sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCUC.SO001 A,EMCUC.CD039 B,EMCUC.CD004 C ' + sL_Where + ')';


    //大新店及北桃園目前沒有920623
    //大新店
    //sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCST.SO001 A,EMCST.CD039 B,EMCST.CD004 C ' + sL_Where + ')';

    //北桃園
    //sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCNTY.SO001 A,EMCNTY.CD039 B,EMCNTY.CD004 C ' + sL_Where + ')';


    with cdsSO001 do
    begin
      Close;
      CommandText := sL_SQL;
      Open;

      FindField('CompName').DisplayLabel := '公司別';
      FindField('ClassName').DisplayLabel := '客戶別';
      FindField('Counts').DisplayLabel := '數量';



      FindField('CompName').DisplayWidth := 30;
      FindField('ClassName').DisplayWidth := 30;
      FindField('Counts').DisplayWidth := 10;


      FindField('CompName').Index := 0;
      FindField('ClassName').Index := 1;
      FindField('Counts').Index := 2;
    end;

end;

function TdtmMain.getSnapShotLogData(sI_Date: String): Boolean;
var
    sL_TitleSql,sL_SQL,sL_Where : String;
begin
    sL_TitleSql := 'SELECT A.COMPCODE,A.SERVICETYPE,A.BEGINEXECTIME,A.ENDEXECTIME,' +
                   'A.TOTALEXECTIME,A.PRGNAME,A.FUNCNAME,A.RETURNCODE,A.RETURNMSG,B.DESCRIPTION FROM ';


    sL_Where := ' WHERE A.JOBID=999993 and A.BeginExecTime BETWEEN ' +
                TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_Date + ' 00:00:01')) +
                ' AND ' + TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_Date + ' 23:59:59')) +
                ' AND B.CODENO=A.COMPCODE ';

//Jackal測試用
{
    //V30
    sL_SQL := sL_TitleSql + 'V30.SO030A A,V30.CD039 B ' + sL_Where;

    //NCC
    sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'NCC.SO030A A,NCC.CD039 B ' + sL_Where + ')';
}

    //陽明山 (1)
    sL_SQL := sL_TitleSql + 'EMCYMS.SO030A A,EMCYMS.CD039 B ' + sL_Where;

    //新台北(2)
    sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCNTP.SO030A A,EMCNTP.CD039 B ' + sL_Where + ')';

    //大安文山(3)
    sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCDW.SO030A A,EMCDW.CD039 B ' + sL_Where + ')';

    //金頻道(4)
    sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCKP.SO030A A,EMCKP.CD039 B ' + sL_Where + ')';

    //振道(5)
    sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCCT.SO030A A,EMCCT.CD039 B ' + sL_Where + ')';

    //豐盟(6)
    sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCFM.SO030A A,EMCFM.CD039 B ' + sL_Where + ')';

    //新頻道(7)
    sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCNCC.SO030A A,EMCNCC.CD039 B ' + sL_Where + ')';

    //南天(8)
    sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCNT.SO030A A,EMCNT.CD039 B ' + sL_Where + ')';

    //觀昇(9)
    sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCKS.SO030A A,EMCKS.CD039 B ' + sL_Where + ')';

    //屏南(10)
    sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCPN.SO030A A,EMCPN.CD039 B ' + sL_Where + ')';

    //新唐城(11)
    sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCTC.SO030A A,EMCTC.CD039 B ' + sL_Where + ')';

    //全聯(12)
    sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCUC.SO030A A,EMCUC.CD039 B ' + sL_Where + ')';

    //大新店及北桃園目前沒有920623
    //大新店
    //sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCST.SO030A A,EMCST.CD039 B ' + sL_Where + ')';

    //北桃園
    //sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCNTY.SO030A A,EMCNTY.CD039 B ' + sL_Where + ')';



//showmessage('4' + '=' + sL_SQL);
    with cdsSO030A do
    begin
      Close;
      CommandText := sL_SQL;
      Open;

      FindField('CompCode').DisplayLabel := '公司別';
      FindField('Description').DisplayLabel := '公司別';
      FindField('ServiceType').DisplayLabel := '服務別';
      FindField('BeginExecTime').DisplayLabel := '開始執行時間';
      FindField('EndExecTime').DisplayLabel := '執行結束時間';
      FindField('TotalExecTime').DisplayLabel := '總計執行時間';
      FindField('PrgName').DisplayLabel := '程式名稱';
      FindField('FuncName').DisplayLabel := '功能名稱';
      FindField('ReturnCode').DisplayLabel := '回傳碼';
      FindField('ReturnMsg').DisplayLabel := '回傳值';


      FindField('CompCode').DisplayWidth := 3;
      FindField('Description').DisplayWidth := 20;
      FindField('ServiceType').DisplayWidth := 1;
      FindField('BeginExecTime').DisplayWidth := 20;
      FindField('EndExecTime').DisplayWidth := 20;
      FindField('TotalExecTime').DisplayWidth := 10;
      FindField('PrgName').DisplayWidth := 30;
      FindField('FuncName').DisplayWidth := 50;
      FindField('ReturnCode').DisplayWidth := 4;
      FindField('ReturnMsg').DisplayWidth := 50;

      FindField('CompCode').Index := 9;
      FindField('Description').Index := 0;
      FindField('ServiceType').Index := 3;
      FindField('BeginExecTime').Index := 4;
      FindField('EndExecTime').Index := 5;
      FindField('TotalExecTime').Index := 6;
      FindField('PrgName').Index := 7;
      FindField('FuncName').Index := 8;
      FindField('ReturnCode').Index := 1;
      FindField('ReturnMsg').Index := 2;

      FindField('CompCode').Visible := false;

      if RecordCount > 0 then
        Result := true
      else
        Result := false;
    end;



end;

procedure TdtmMain.getSO509AData(sI_DbOwner,sI_EngDate: String);
var
    sL_SQL : String;
begin
    sL_SQL := 'SELECT * FROM ' + sI_DbOwner + '.SO509A WHERE COMPUTEDATE BETWEEN ' +
              TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_EngDate + ' 00:00:01')) + ' AND ' +
              TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_EngDate + ' 23:59:59'));

    with cdsSO509A do
    begin
      Close;
      CommandText := sL_SQL;
      Open;
    end;
end;

procedure TdtmMain.getSO509BData(sI_DbOwner, sI_CompCode,
  sI_EngDate: String);
var
    sL_SQL : String;
    nL_StrDiffCounts : Integer;
begin
    sL_SQL := 'SELECT CompCode,CompName,TableCode,TableName,MasterComputeDate,' +
              'MasterRecordCounts,SnapComputeDate,SnapRecordCounts,DiffCounts,' +
              'to_char(DiffCounts) StrDiffCounts,ResultFlag  FROM ' +
              sI_DbOwner + '.SO509B WHERE COMPCODE=' + sI_CompCode +
              ' AND MASTERCOMPUTEDATE BETWEEN ' +
              TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_EngDate + ' 00:00:01')) + ' AND ' +
              TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_EngDate + ' 23:59:59')) +
              ' ORDER BY RESULTFLAG';

    with cdsSO509B do
    begin
      Close;
      CommandText := sL_SQL;
      Open;

      cdsTempSO509B.EmptyDataSet;

      First;
      while not Eof do
      begin
        nL_StrDiffCounts := cdsSO509B.FieldByName('StrDiffCounts').AsInteger;

        Edit;
        cdsSO509B.FieldByName('DiffCounts').AsInteger := nL_StrDiffCounts;
        Post;



        cdsTempSO509B.Append;

        cdsTempSO509B.FieldByName('CompName').AsString := FieldByName('CompName').AsString;
        cdsTempSO509B.FieldByName('TableName').AsString := FieldByName('TableName').AsString;
        cdsTempSO509B.FieldByName('MasterComputeDate').AsString := FieldByName('MasterComputeDate').AsString;
        cdsTempSO509B.FieldByName('MasterRecordCounts').AsString := FieldByName('MasterRecordCounts').AsString;
        cdsTempSO509B.FieldByName('SnapComputeDate').AsString := FieldByName('SnapComputeDate').AsString;
        cdsTempSO509B.FieldByName('SnapRecordCounts').AsString := FieldByName('SnapRecordCounts').AsString;
        cdsTempSO509B.FieldByName('DiffCounts').AsString := FieldByName('DiffCounts').AsString;


        if nL_StrDiffCounts = 0 then
          cdsTempSO509B.FieldByName('ResultFlag').AsString := '正常'
        else
          cdsTempSO509B.FieldByName('ResultFlag').AsString := '異常';

        cdsTempSO509B.Post;

        Next;
      end;
    end;


end;

procedure TdtmMain.DataModuleCreate(Sender: TObject);
begin
    G_CompCodeStrList := TStringList.Create;

    if not cdsTempSO509B.Active then
      cdsTempSO509B.CreateDataSet;
end;

procedure TdtmMain.DataModuleDestroy(Sender: TObject);
begin
    G_CompCodeStrList.Free;
end;

function TdtmMain.getCompName(sI_CompCode: String): String;
var
    sL_CompName : String;
begin
    if cdsSO030A.Locate('CompCode',sI_CompCode,[]) then
      sL_CompName := cdsSO030A.FieldByName('DESCRIPTION').AsString;

    Result := sL_CompName;
end;



end.
