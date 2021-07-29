unit dtmMainU;

interface

uses
  SysUtils, Classes, DB, ADODB, DBClient, Provider,Dialogs,IniFiles,QForms;

type
  TSnapShotIniData = class(TObject)
    nDataAreaNo : Integer;
    sDataArea : String;
    sAlias : String;
    sUserID : String;
    sPassword : String;
    sCompCode : String;
    sCompName : String;
    sDbOwner : String;
  end;


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
    G_SnapShotIniStrList : TStringList;
    sG_ExePath,sG_DefaultDbOwnerName,sG_DefaultDbOwnerPassword : String;
    function connectToDB(sI_Alias,sI_UserName,sI_Password : String):WideString;
    function getSnapShotLogData(sI_Date : String) : Boolean;
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

    procedure TransTmpIniFile;
    function getDbConnInfo : WideString;
    procedure DeleteTmpIniFile;
  end;

var
  dtmMain: TdtmMain;

implementation

uses Ustru, frmSnapShotU, DevPower_TLB, UCommonU;

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
    //���X Master site �ΨӬ����C�ѦU table ����Ƶ���
    getSO509AData(sI_DbOwner,sI_EngDate);

    if cdsSO509A.RecordCount <> 0 then
    begin
      cdsSO509.First;
      while not cdsSO509.Eof do
      begin
        //Move �� SnapShot �� Table
        sL_MoveTableCode := cdsSO509.FieldByName('TableCode').AsString;
        sL_MoveTableName := cdsSO509.FieldByName('TableName').AsString;


        //�p�� SnapShot �� Table ������
        nL_SnapTableCouts := computeTableCounts(sI_DbOwner,sL_MoveTableCode);


        //�p�� SO �� Table �����
        if cdsSO509A.Locate('TableCode',sL_MoveTableCode,[]) then
        begin
          nL_MasterTableCouts := cdsSO509A.FieldByName('RecordCounts').AsInteger;
          dL_MasterComputeDate := cdsSO509A.FieldByName('ComputeDate').AsDateTime;
        end;

        //�� SO �� Table ���� - SnapShot �� Table ����
        nL_DiffCouts := nL_MasterTableCouts - nL_SnapTableCouts;


        sL_ResultFlag := getDiffCoutsResultFlag(nL_MasterTableCouts,nL_SnapTableCouts,nL_DiffCouts);

        //�N�p�⵲�G�s��SO509B
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


    //���X SO509B Data
    getSO509BData(sI_DbOwner, sI_CompCode,sI_EngDate);
end;

procedure TdtmMain.computeSoAndSnapshotTableCounts(sI_DbOwner, sI_CompCode,
  sI_CompName, sI_EngDate: String);
begin

    //�ˬd�Ӥ��q�Ӥ���O�_�w�g���Q�p��L,�Y���p��h�p�⤧
    if not dtmMain.checkSO509BIsCompute(sI_DbOwner,sI_CompCode,sI_EngDate) then
    begin
      computeSO509B(sI_DbOwner,sI_CompCode,sI_CompName,sI_EngDate);
    end
    else
    begin
      //���X SO509B Data
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
        sL_DbConnStr := 'Provider=MSDAORA.1;Password=' + sI_Password + ';User ID='+sI_UserName+';Data Source='+sI_Alias+';Persist Security Info=True'; //�A�Ω��L PC
        ADOConnection1.ConnectionString := sL_DbConnStr;
        ADOConnection1.Connected := true;
      end;
    except
      result := '-1: �PCC&B��Ʈw�s�u����, ��Ʈw�O�W=<'+sI_Alias+'>, DB User=<'+sI_UserName +'>';
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
    nL_Ndx : Integer;
begin
    nL_Ndx := G_SnapShotIniStrList.IndexOf(sI_CompCode);
    if nL_Ndx<>-1 then
       sL_DbOwner := (G_SnapShotIniStrList.Objects[nL_Ndx] as TSnapShotIniData).sDbOwner;
{
    if sI_CompCode = '1' then
    begin
      //�[�@
      sL_DbOwner := 'EMCKS';
    end
    else if sI_CompCode = '2' then
    begin
      //�̫n
      sL_DbOwner := 'EMCPN';
    end
    else if sI_CompCode = '3' then
    begin
      //�n��
      sL_DbOwner := 'EMCNT';
    end
    else if sI_CompCode = '5' then
    begin
      //�s�W�D
      sL_DbOwner := 'EMCNCC';
      //showmessage('����30');
      //sL_DbOwner := 'V30';
    end
    else if sI_CompCode = '6' then
    begin
      //�׷�
      sL_DbOwner := 'EMCFM';
    end
    else if sI_CompCode = '7' then
    begin
      //���D
      sL_DbOwner := 'EMCCT';
    end
    else if sI_CompCode = '8' then
    begin
      //���p
      sL_DbOwner := 'EMCUC';
    end
    else if sI_CompCode = '9' then
    begin
      //�����s
      sL_DbOwner := 'EMCYMS';
    end
    else if sI_CompCode = '10' then
    begin
      //�s�x�_
      sL_DbOwner := 'EMCNTP';
    end
    else if sI_CompCode = '11' then
    begin
      //���W�D
      sL_DbOwner := 'EMCKP';
    end
    else if sI_CompCode = '12' then
    begin
      //�j�w��s
      sL_DbOwner := 'EMCDW';
    end
    else if sI_CompCode = '13' then
    begin
      //�s��
      sL_DbOwner := 'EMCTC';
    end
    else if sI_CompCode = '14' then
    begin
      //�j�s��
      sL_DbOwner := 'EMCST';
    end
    else if sI_CompCode = '16' then
    begin
      //�_���
      sL_DbOwner := 'EMCNTY';
    end;
}
    Result := sL_DbOwner;
end;

function TdtmMain.getDiffCoutsResultFlag(nI_MasterTableCouts,
  nI_SnapTableCouts, nI_DiffCouts: Integer): String;
begin
    {
    1: Master��Ƶ��Ƶ���Snapshot��Ƶ���
    2: Master��Ƶ��Ƥj��Snapshot��Ƶ���
    3: Master��Ƶ��Ƥp��Snapshot��Ƶ���
    4: Master��Ƶ��Ƭ�0
    5: Snapshot��Ƶ��Ƭ�0
    6: Master��Ƶ��ƻPSnapshot��Ƶ��Ƨ���0
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
    sL_SQL := 'SELECT * FROM ' + sG_DefaultDbOwnerName + '.SO509';

    with cdsSO509 do
    begin
      Close;
      CommandText := sL_SQL;
      Open;
    end;

end;


function TdtmMain.getSnapShotLogData(sI_Date: String): Boolean;
var
    sL_TitleSql,sL_SQL,sL_Where,sL_CompCode,sL_DbOwner : String;
    ii : Integer;
begin
    sL_TitleSql := 'SELECT A.COMPCODE,A.SERVICETYPE,A.BEGINEXECTIME,A.ENDEXECTIME,' +
                   'A.TOTALEXECTIME,A.PRGNAME,A.FUNCNAME,A.RETURNCODE,A.RETURNMSG,B.DESCRIPTION FROM ';

    //����
    sL_Where := ' WHERE A.JOBID=999993 and A.BeginExecTime BETWEEN ' +
                TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_Date + ' 00:00:01')) +
                ' AND ' + TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_Date + ' 23:59:59')) +
                ' AND B.CODENO=A.COMPCODE ';

    for ii:=0 to G_SnapShotIniStrList.Count-1 do
    begin
      sL_CompCode := (G_SnapShotIniStrList.Objects[ii] as TSnapShotIniData).sCompCode;
      sL_DbOwner := getDbOwner(sL_CompCode);

      if ii=0 then
        sL_SQL := sL_TitleSql + sL_DbOwner + '.SO030A A,' + sL_DbOwner + '.CD039 B ' + sL_Where
      else
        sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + sL_DbOwner + '.SO030A A,' + sL_DbOwner + '.CD039 B ' + sL_Where + ')';
    end;
//Jackal���ե�
{
    //V30
    sL_SQL := sL_TitleSql + 'V30.SO030A A,V30.CD039 B ' + sL_Where;

    //NCC
    sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'NCC.SO030A A,NCC.CD039 B ' + sL_Where + ')';
}
{
    //�����s (1)
    sL_SQL := sL_TitleSql + 'EMCYMS.SO030A A,EMCYMS.CD039 B ' + sL_Where;

    //�s�x�_(2)
    sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCNTP.SO030A A,EMCNTP.CD039 B ' + sL_Where + ')';

    //�j�w��s(3)
    sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCDW.SO030A A,EMCDW.CD039 B ' + sL_Where + ')';

    //���W�D(4)
    sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCKP.SO030A A,EMCKP.CD039 B ' + sL_Where + ')';

    //���D(5)
    sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCCT.SO030A A,EMCCT.CD039 B ' + sL_Where + ')';

    //�׷�(6)
    sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCFM.SO030A A,EMCFM.CD039 B ' + sL_Where + ')';

    //�s�W�D(7)
    sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCNCC.SO030A A,EMCNCC.CD039 B ' + sL_Where + ')';

    //�n��(8)
    sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCNT.SO030A A,EMCNT.CD039 B ' + sL_Where + ')';

    //�[�@(9)
    sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCKS.SO030A A,EMCKS.CD039 B ' + sL_Where + ')';

    //�̫n(10)
    sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCPN.SO030A A,EMCPN.CD039 B ' + sL_Where + ')';

    //�s��(11)
    sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCTC.SO030A A,EMCTC.CD039 B ' + sL_Where + ')';

    //���p(12)
    sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCUC.SO030A A,EMCUC.CD039 B ' + sL_Where + ')';

    //�j�s���Υ_���ثe�S��920623
    //�j�s��
    //sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCST.SO030A A,EMCST.CD039 B ' + sL_Where + ')';

    //�_���
    //sL_SQL := sL_SQL + ' UNION ALL (' + sL_TitleSql + 'EMCNTY.SO030A A,EMCNTY.CD039 B ' + sL_Where + ')';
}


//showmessage('4' + '=' + sL_SQL);
    with cdsSO030A do
    begin
      Close;
      CommandText := sL_SQL;
      Open;

      FindField('CompCode').DisplayLabel := '���q�O';
      FindField('Description').DisplayLabel := '���q�O';
      FindField('ServiceType').DisplayLabel := '�A�ȧO';
      FindField('BeginExecTime').DisplayLabel := '�}�l����ɶ�';
      FindField('EndExecTime').DisplayLabel := '���浲���ɶ�';
      FindField('TotalExecTime').DisplayLabel := '�`�p����ɶ�';
      FindField('PrgName').DisplayLabel := '�{���W��';
      FindField('FuncName').DisplayLabel := '�\��W��';
      FindField('ReturnCode').DisplayLabel := '�^�ǽX';
      FindField('ReturnMsg').DisplayLabel := '�^�ǭ�';


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
          cdsTempSO509B.FieldByName('ResultFlag').AsString := '���`'
        else
          cdsTempSO509B.FieldByName('ResultFlag').AsString := '���`';

        cdsTempSO509B.Post;

        Next;
      end;
    end;


end;

procedure TdtmMain.DataModuleCreate(Sender: TObject);
var
    sL_Result : String;
begin
    G_CompCodeStrList := TStringList.Create;
    G_SnapShotIniStrList := TStringList.Create;

    if not cdsTempSO509B.Active then
      cdsTempSO509B.CreateDataSet;

    //down, �q ini ��Ū���s�u�� DB �� information
    sL_Result := getDbConnInfo;
    if sL_Result<>'' then
    begin
      showmessage(sL_Result);
      Application.Terminate;
    end;
    //up, �q ini ��Ū���s�u�� DB �� information
end;

procedure TdtmMain.DataModuleDestroy(Sender: TObject);
begin
    G_CompCodeStrList.Free;
    G_SnapShotIniStrList.Free;
end;

function TdtmMain.getCompName(sI_CompCode: String): String;
var
    sL_CompName : String;
begin
    if cdsSO030A.Locate('CompCode',sI_CompCode,[]) then
      sL_CompName := cdsSO030A.FieldByName('DESCRIPTION').AsString;

    Result := sL_CompName;
end;



procedure TdtmMain.TransTmpIniFile;
var
    L_StrList, L_TmpStrList : TStringList;
    sL_ExeFileName, sL_ExePath ,sL_IniFileName,sL_TempIniFileName : String;
    ii : Integer;
    L_Intf : _Encrypt;
    sL_EncKey : WideString;
    sL_Temp : WideString;
begin
    sL_EncKey := 'CS';
    L_Intf := CoEncrypt.Create;
    sL_ExeFileName := Application.ExeName;
    sL_ExePath := ExtractFileDir(sL_ExeFileName);

    sL_ExeFileName := Application.ExeName;

    sL_ExePath := ExtractFileDir(sL_ExeFileName);

    sL_IniFileName := sL_ExePath + '\' + SYS_INI_FILE_NAME;


    L_StrList := TStringList.Create;
    L_TmpStrList := TStringList.Create;

    L_StrList.LoadFromFile(sL_IniFileName);
    for ii:=0 to L_StrList.Count-1 do
    begin
       if (Copy(L_StrList.Strings[ii],1,1)=';') or (L_StrList.Strings[ii]='')
          or (Copy(L_StrList.Strings[ii],1,8)='COMPCODE') or (Copy(L_StrList.Strings[ii],1,8)='COMPNAME') then
       begin
         L_TmpStrList.Add(L_StrList.Strings[ii]);
       end
       else
       begin
         sL_Temp := L_StrList.Strings[ii];
         L_TmpStrList.Add(L_Intf.Decrypt(sL_Temp));
       end;

    end;
    L_TmpStrList.SaveToFile(sL_ExePath + '\' + TMP_SYS_INI_FILE_NAME);
    L_TmpStrList.Free;

    L_Intf := nil;
end;

function TdtmMain.getDbConnInfo: WideString;
var
    sL_IniFileName : String;
    L_IniFile : TIniFile;
    ii, nL_TotalDataArea : Integer;
    sL_DbAlias, sL_DbUserName, sL_DbPassword, sL_Tag : String;
    sL_CompCodeAndName,sL_CompCode,sL_CompName : String;
    L_Obj : TSnapShotIniData;
    sL_ExeFileName,sL_ExePath : String;
begin
    //�N�[�K��ini�ɮ׸ѱK
    TransTmpIniFile;

    sL_ExeFileName := Application.ExeName;

    sL_ExePath := ExtractFileDir(sL_ExeFileName);

    sG_ExePath := sL_ExePath;

    sL_IniFileName := sL_ExePath + '\' + TMP_SYS_INI_FILE_NAME;

    if not FileExists(sL_IniFileName) then
    begin
      result := '-1: Ū����Ʈw�s�u�Ѽ���<'+sL_IniFileName +'>����';
      exit;
    end;
    L_IniFile := TIniFile.Create(sL_IniFileName);
    nL_TotalDataArea := L_IniFile.ReadInteger('SYSINFO','TotalDataAreaCounts',0);
    if nL_TotalDataArea=0 then
    begin
      result := '-1: ��Ʈw�s�u�Ѽ��ɤ���ư��`�Ƴ]�w���~!<'+sL_IniFileName +'>';
      exit;
    end;

    sG_DefaultDbOwnerName := L_IniFile.ReadString('SYSINFO','DEFAULT_DB_OWNER_NAME','EMCYMS');
    sG_DefaultDbOwnerPassword := L_IniFile.ReadString('SYSINFO','DEFAULT_DB_OWNER_PASSWORD','EMCYMS');

    {
    G_DbAliasStrList.Clear;
    G_DbUserNameStrList.Clear;
    G_DbPasswordStrList.Clear;
    }
    try
      for ii:=0 to nL_TotalDataArea -1 do
      begin
        sL_Tag := DATA_AREA_HEADER + IntToStr(ii);
        {
        sL_DbAlias := L_IniFile.ReadString(sL_Tag,'ALIAS','');
        sL_DbUserName := L_IniFile.ReadString(sL_Tag,'USERID','');
        sL_DbPassword := L_IniFile.ReadString(sL_Tag,'PASSWORD','');
        }

        sL_CompCode := L_IniFile.ReadString(sL_Tag,'COMPCODE','');
        sL_CompName := L_IniFile.ReadString(sL_Tag,'COMPNAME','');
        sL_CompCodeAndName := sL_CompCode + '-' + sL_CompName;
        //G_CompCodeAndNameStrList.Add(sL_CompCodeAndName);


        L_Obj := TSnapShotIniData.Create;
        L_Obj.nDataAreaNo := ii;
        L_Obj.sDataArea := sL_Tag;
        L_Obj.sAlias := L_IniFile.ReadString(sL_Tag,'ALIAS','');
        L_Obj.sUserID := L_IniFile.ReadString(sL_Tag,'USERID','');
        L_Obj.sPassword := L_IniFile.ReadString(sL_Tag,'PASSWORD','');
        L_Obj.sCompName := L_IniFile.ReadString(sL_Tag,'COMPNAME','');
        L_Obj.sCompCode := L_IniFile.ReadString(sL_Tag,'COMPCODE','');
        L_Obj.sDbOwner := L_IniFile.ReadString(sL_Tag,'DBOWNER','');

        {
        G_DbAreaNameStrList.Add(sL_Tag);

        G_DbAliasStrList.Add(sL_DbAlias);
        G_DbUserNameStrList.Add(sL_DbUserName);
        G_DbPasswordStrList.Add(sL_DbPassword);
        }

        G_SnapShotIniStrList.AddObject(sL_CompCode, L_Obj);
      end;

      L_IniFile.Free;
    except
      result := '-1: ��Ʈw�s�u�Ѽ��ɤ���ưϳ]�w���~!<'+sL_IniFileName +'>';
      exit;
    end;
    {
    G_DbAliasStrList.SaveToFile('c:\a1.txt');
    G_DbUserNameStrList.SaveToFile('c:\a2.txt');
    G_DbPasswordStrList.SaveToFile('c:\a3.txt');
    }
    result := '';


end;

procedure TdtmMain.DeleteTmpIniFile;
var
    sL_ExeFileName,sL_ExePath,sL_IniFileName : String;
begin
    //�R���ѱK�ᤧ ini ��
    sL_ExeFileName := Application.ExeName;

    sL_ExePath := ExtractFileDir(sL_ExeFileName);

    sL_IniFileName := sL_ExePath + '\' + TMP_SYS_INI_FILE_NAME;
    DeleteFile(sL_IniFileName);
    Application.Terminate;
end;

end.
