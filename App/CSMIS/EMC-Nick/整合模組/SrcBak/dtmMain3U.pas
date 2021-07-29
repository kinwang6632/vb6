//Provider=MSDAORA.1;Password=v30;User ID=v30;Data Source=mis
unit dtmMain3U;

interface

uses
  SysUtils, Classes, DB, ADODB, Dialogs, Controls, DBClient, Provider, Variants,
  DBTables;

const
    STR_SEP = '''';
    PAPER_NUMBER_LENGTH = 0;
    REPORT2_PAPER_COLUMN_COUNTS = 8;
    SO128_FIELD_SEP= '#';
    SO128_FIELD_VALUE_SEP= '-';    
type
  TdtmMain3 = class(TDataModule)
    qrySo126: TADOQuery;
    qrySo126COMPCODE: TIntegerField;
    qrySo126BEGINNUM: TStringField;
    qrySo126ENDNUM: TStringField;
    qrySo126TOTALPAPERCOUNT: TIntegerField;
    qrySo126OPERATOR: TStringField;
    qrySo126UPDTIME: TStringField;
    qryCm003: TADOQuery;
    qryCm003EMPNO: TStringField;
    qryCm003EMPNAME: TStringField;
    qrySo126GETPAPERDATE: TDateTimeField;
    qrySo126EMPNO: TStringField;
    qrySo126EMPNAME: TStringField;
    qryCommon: TADOQuery;
    qrySo126SEQ: TBCDField;
    cdsSo126: TClientDataSet;
    dspSo126: TDataSetProvider;
    cdsSo126SEQ: TBCDField;
    cdsSo126COMPCODE: TIntegerField;
    cdsSo126EMPNO: TStringField;
    cdsSo126EMPNAME: TStringField;
    cdsSo126GETPAPERDATE: TDateTimeField;
    cdsSo126BEGINNUM: TStringField;
    cdsSo126ENDNUM: TStringField;
    cdsSo126TOTALPAPERCOUNT: TIntegerField;
    cdsSo126OPERATOR: TStringField;
    cdsSo126UPDTIME: TStringField;
    qryReport1: TADOQuery;
    qryReport2: TADOQuery;
    qryReport3: TADOQuery;
    qryReport2EMPNO: TStringField;
    qryReport2EMPNAME: TStringField;
    qryReport2GETPAPERDATE: TDateTimeField;
    cdsReport2: TClientDataSet;
    cdsReport2GetPaperDate: TStringField;
    cdsReport2PaperNumbers: TStringField;
    cdsReport2EmpNo: TStringField;
    cdsReport2EmpName: TStringField;
    cdsReport2GroupID: TIntegerField;
    qryReport2Common: TADOQuery;
    qryReport2CommonEMPNO: TStringField;
    qryReport2CommonCOUNTS: TBCDField;
    qryReport2CommonGETPAPERDATE: TStringField;
    qryReport3EMPNO: TStringField;
    qryReport3EMPNAME: TStringField;
    qryReport3GETPAPERDATE: TDateTimeField;
    qryReport3BILLNO: TStringField;
    qryReport3CUSTID: TIntegerField;
    qryReport3CUSTNAME: TStringField;
    qryReport3CUSTTEL: TStringField;
    qryReport3REALDATE: TDateTimeField;
    qryReport3STATUS: TIntegerField;
    qryPriv: TADOQuery;
    qrySo18C2: TADOQuery;
    qrySo18C2EMPNAME: TStringField;
    qrySo18C2GETPAPERDATE: TDateTimeField;
    qrySo18C2BILLNO: TStringField;
    qrySo126PREFIX: TStringField;
    cdsSo126PREFIX: TStringField;
    qryReport2PAPERNUM: TStringField;
    qryReport3PAPERNUM: TStringField;
    qrySo18C2PAPERNUM: TStringField;
    qrySo034: TADOQuery;
    qrySo034CUSTID: TIntegerField;
    qrySo034BILLNO: TStringField;
    qrySo034MANUALNO: TStringField;
    qrySo034ITEM: TIntegerField;
    qrySo034CITEMNAME: TStringField;
    qrySo034OLDAMT: TIntegerField;
    qrySo034SHOULDAMT: TIntegerField;
    qrySo034REALDATE: TDateTimeField;
    qrySo034REALAMT: TIntegerField;
    qrySo034CLCTNAME: TStringField;
    qrySo034CMNAME: TStringField;
    qrySo034UCNAME: TStringField;
    qrySo034STNAME: TStringField;
    qrySo034CUSTNAME: TStringField;
    qrySo034TEL1: TStringField;
    qryReport4: TADOQuery;
    qryReport4COMPCODE: TIntegerField;
    qryReport4PAPERNUM: TStringField;
    qryReport4UPDATESTATUS: TStringField;
    qryReport4OPERATOR: TStringField;
    cdsReport4: TClientDataSet;
    cdsReport4PaperNum: TStringField;
    cdsReport4Operator: TStringField;
    cdsReport4UpdTime: TStringField;
    cdsReport4FieldName: TStringField;
    cdsReport4OldValue: TStringField;
    cdsReport4NewValue: TStringField;
    qryReport4UPDTIME: TDateTimeField;
    cdsReport4Group: TIntegerField;
    qryReport5: TADOQuery;
    qryReport5SEQ: TBCDField;
    qryReport5SPAPERNUM: TStringField;
    qryReport5EPAPERNUM: TStringField;
    qryReport5TOTALPAPERCOUNT: TIntegerField;
    qryReport5EMPNAME: TStringField;
    qryReport5GETPAPERDATE: TDateTimeField;
    qryReport5NOTES: TStringField;
    qryReport5OPERATOR: TStringField;
    qryReport5UPDTIME: TStringField;
    cdsSo126RETURNDATE: TDateTimeField;
    cdsSo126CLEARDATE: TDateTimeField;
    cdsSo126NOTE: TStringField;
    qrySo126RETURNDATE: TDateTimeField;
    qrySo126CLEARDATE: TDateTimeField;
    qrySo126NOTE: TStringField;
    rtpSO8C10: TClientDataSet;
    procedure qrySo126GETPAPERDATEGetText(Sender: TField; var Text: String;
      DisplayText: Boolean);
    procedure cdsSo126BeforePost(DataSet: TDataSet);
    procedure cdsSo126GETPAPERDATEGetText(Sender: TField; var Text: String;
      DisplayText: Boolean);
    procedure qryReport2CommonGETPAPERDATEGetText(Sender: TField;
      var Text: String; DisplayText: Boolean);
    procedure qrySo18C2GETPAPERDATEGetText(Sender: TField;
      var Text: String; DisplayText: Boolean);
    procedure cdsSo126RETURNDATEGetText(Sender: TField; var Text: String;
      DisplayText: Boolean);
    procedure cdsSo126CLEARDATEGetText(Sender: TField; var Text: String;
      DisplayText: Boolean);
  private
    { Private declarations }
    function checkDul(sI_Prefix:String; nI_BeginNumLength:Integer; fI_BeginNum,fI_EndNum:double):boolean;
    procedure transReport2Data(sI_WhereClause:String);
    procedure transReport3Data;
    function getSeq:Integer;
  public
    { Public declarations }
    function get129Seq:Integer;    
    function reusePaper(dI_OriginalGetPaperDate:TDate; sI_OriginalEmpNo, sI_OriginalEmpName, sI_OriginalBeginNum, sI_OriginalSeq, sI_CompCode,  sI_EmpNo, sI_EmpName, sI_Prefix, sI_BeginNum, sI_EndNum, sI_GetPaperUser, sI_GetPaperDate, sI_OriginalCount, sI_Count : String):String;
    function getSo128Data(dI_SDate, dI_EDate:TDate):Integer;
    function setNewHandPaperNo(sI_CustID,sI_CustName, sI_Tel1, sI_Item:String; dI_RealDate : TDate; sI_BillNo, sI_OldHandPaperNo, sI_NewHandPaperNo:String; var sI_ErrorMsg:String):boolean;
    function getSo034Data(sI_BillNo:String):Integer;
    function getReport2PaperCounts(sI_EmpNo, sI_GetPaperDate:String):Integer;overload;
    function getReport2PaperCounts(sI_EmpNo:String; dI_GetPaperDate:TDate):Integer;overload;    
    function activeReportData(sI_BeginDate, sI_EndDate, sI_EmpSQL, sI_BeginPaperNum, sI_EndPaperNum:String; nI_ReportType:Integer):boolean;
    function disablePaperData(sI_PaperSNum, sI_PaperENum:String):Boolean;
    function getPaperData(sI_PaperSNum, sI_PaperENum:String):TDataSet;
    procedure deletePaperData(nI_Seq:Integer);
    function canModifyGetPaperData(nI_Seq:Integer):boolean;
    function activeGetPaperData(sI_BeginDate, sI_EndDate, sI_EmpSQL, sI_PaperNum, sI_Prefix:String):boolean;
    procedure activeDataSet(I_DataSet : TDataSet);
    procedure saveToDB(I_CDS:TClientDataSet);
    procedure _8C30ToExcel(const aReportIndex: Integer; const aPaperDate, aPaperNumber,
      aEmpNames: String; aFile: TFileName);
  end;

var
  dtmMain3: TdtmMain3;

implementation

uses UdateTimeu, Ustru, frmMainMenuU, XLSFile;

{$R *.dfm}

{ TdtmMain }

procedure TdtmMain3.activeDataSet(I_DataSet: TDataSet);
begin
  if not I_DataSet.Active then
    I_DataSet.Active := true;
end;

function TdtmMain3.activeGetPaperData(sI_BeginDate, sI_EndDate, sI_EmpSQL,
  sI_PaperNum, sI_Prefix: String): boolean;
var
    bL_Result: boolean;
    dL_BeginDate, dL_EndDate : TDate;
    sL_SQL, sL_Where : String;
begin
    sL_SQL := 'select * from SO126';
    sL_Where := '';
    if sI_BeginDate<>'' then
    begin
      dL_BeginDate := TUdateTime.CDate2Date(sI_BeginDate);
      dL_EndDate := TUdateTime.CDate2Date(sI_EndDate);
      sL_Where := ' where GETPAPERDATE between ' + TUstr.getOracleSQLDateStr(dL_BeginDate);
      sL_Where := sL_Where + ' and '  + TUstr.getOracleSQLDateStr(dL_EndDate);
    end;

    if sI_EmpSQL<>'' then
    begin
      if sL_Where='' then
        sL_Where := ' where ' +  sI_EmpSQL
      else
        sL_Where := sL_Where + ' and ' +  sI_EmpSQL;
    end;

    if sI_Prefix<>'' then
    begin
      if sL_Where='' then
        sL_Where := ' where PREFIX=' +  STR_SEP + sI_Prefix + STR_SEP
      else
        sL_Where := sL_Where + ' and PREFIX=' + STR_SEP + sI_Prefix + STR_SEP;
    end;


    if sI_PaperNum<>'' then
    begin
      if sL_Where='' then
        sL_Where := ' where ' + sI_PaperNum + ' between to_number(BEGINNUM) and  to_number(ENDNUM)'
      else
        sL_Where := sL_Where + ' and ' + sI_PaperNum + ' between to_number(BEGINNUM) and  to_number(ENDNUM)';
    end;
    sL_SQL := sL_SQL + sL_Where;
{

select BEGINNUM, ENDNUM from so126
where 350 between to_number(BEGINNUM) and  to_number(ENDNUM);
}
    with cdsSo126 do
    begin
      Close;
      CommandText := sL_SQL;
      Open;
    end;
        
    if cdsSo126.RecordCount>0 then
      bL_Result := true
    else
      bL_Result := false;
    result := bL_Result;  
end;

function TdtmMain3.canModifyGetPaperData(nI_Seq: Integer): boolean;
var
    sL_SQL : String;
    bL_Result : boolean;
begin
    with qryCommon do
    begin
      sL_SQL := 'select count(SEQ) COUNT from SO127 where BILLNO Is NOT NULL and SEQ=' + IntToStr(nI_Seq) ;
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;
      if FieldByName('COUNT').AsInteger=0 then
        bL_Result := true
      else
        bL_Result := false;
      Close;
    end;
    result := bL_Result;
end;



function TdtmMain3.getSeq: Integer;
var
    sL_SQL : String;
begin
    with qryCommon do
    begin
      Close;
      sL_SQL := 'SELECT S_So126_Seq.Nextval  FROM DUAL';
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;
      result := FieldByName('Nextval').AsInteger;
      Close;
    end;
end;


procedure TdtmMain3.qrySo126GETPAPERDATEGetText(Sender: TField;
  var Text: String; DisplayText: Boolean);
begin
    if Sender.AsDateTime=0 then Exit;
    Text := TUdateTime.CDateStr(Sender.AsDateTime,9);
    DisplayText := true;
end;

procedure TdtmMain3.cdsSo126BeforePost(DataSet: TDataSet);
var

    nL_BeginNumLength, ii, nL_Seq, nL_InvalidInt, nL_PaperCount : Integer;
    fL_BeginNum,fL_EndNum : double;
    sL_TmpEndNum, sL_Prefix, sL_TmpPaperNum, sL_SQL, sL_TmpPrefix, sL_BeginNum, sL_EndNum, sL_Now : String;
begin
    nL_InvalidInt := -1;
    if StrToFloatDef(DataSet.FieldByName('BEGINNUM').AsString,nL_InvalidInt)=nL_InvalidInt then
    begin
      MessageDlg('請輸入數字!', mtError, [mbOK],0);
      DataSet.FieldByName('BEGINNUM').FocusControl;
      Abort;
      Exit;
    end;

    sL_TmpPrefix := DataSet.FieldByName('PREFIX').AsString;
    for ii:=1 to length(sL_TmpPrefix) do
    begin
      if StrToIntDef(sL_TmpPrefix[ii],nL_InvalidInt)<>nL_InvalidInt then
      begin
        MessageDlg('不行輸入輸入數字!', mtError, [mbOK],0);
        DataSet.FieldByName('PREFIX').FocusControl;
        Abort;
        Exit;
      end;
    end;
    
    sL_BeginNum := TUstr.AddString(DataSet.FieldByName('BEGINNUM').AsString,'0',false,PAPER_NUMBER_LENGTH);
    nL_BeginNumLength := length(sL_BeginNum);
    fL_BeginNum := StrToFloat(sL_BeginNum);

    nL_PaperCount := DataSet.FieldByName('TOTALPAPERCOUNT').AsInteger;
    fL_EndNum := DataSet.FieldByName('BEGINNUM').AsFloat + nL_PaperCount-1;

    sL_Prefix := DataSet.FieldByName('PREFIX').AsString;
    sL_Now := TUdateTime.GetPureChineseDateTimeStr(now);
    if DataSet.State = dsInsert then
    begin
      if not checkDul(sL_Prefix, nL_BeginNumLength, fL_BeginNum,fL_EndNum) then
      begin
        nL_Seq := getSeq;
        DataSet.FieldByName('SEQ').AsInteger := nL_Seq;
  //      for ii:=fL_BeginNum to fL_EndNum do
        for ii:=0  to nL_PaperCount -1do
        begin
  //        sL_TmpPaperNum := TUstr.AddString(IntToStr(ii),'0',false,PAPER_NUMBER_LENGTH);
          sL_TmpPaperNum := TUstr.AddString(FloatToStr(StrToFloat(sL_BeginNum)+ii),'0',false,PAPER_NUMBER_LENGTH);
          sL_TmpPaperNum := TUstr.AddString(sL_TmpPaperNum,'0',false,nL_BeginNumLength);
          sL_TmpPaperNum := sL_Prefix + sL_TmpPaperNum;
          sL_SQL := 'insert into SO127(COMPCODE,SEQ, PAPERNUM, EMPNO,EMPNAME,GETPAPERDATE,BILLNO, STATUS, OPERATOR, UPDTIME)';
          sL_SQL := sL_SQL + 'values(' + frmMainMenu.sG_CompCode + ',' + IntToStr(nL_Seq)+',' + STR_SEP + sL_TmpPaperNum + STR_SEP;
          sL_SQL := sL_SQL + ',' +  STR_SEP + DataSet.FieldByName('EMPNO').AsString + STR_SEP ;
          sL_SQL := sL_SQL + ',' +  STR_SEP + DataSet.FieldByName('EMPNAME').AsString + STR_SEP ;
          sL_SQL := sL_SQL + ',' +  TUstr.getOracleSQLDateStr(DataSet.FieldByName('GETPAPERDATE').AsDateTime) ;
          sL_SQL := sL_SQL + ',NULL,1,' + STR_SEP + frmMainMenu.sG_User + STR_SEP;
          sL_SQL := sL_SQL + ',' + STR_SEP + sL_Now + STR_SEP +')';
          qryCommon.SQL.Clear;
          qryCommon.SQL.Add(sL_SQL);
          qryCommon.ExecSQL;
          qryCommon.Close;
        end;
      end
      else
      begin
        MessageDlg('號碼有所重複,請檢查!', mtError, [mbOK],0);
        Abort;
        Exit;
      end;

    end;

    DataSet.FieldByName('COMPCODE').AsString := frmMainMenu.sG_CompCode;
    DataSet.FieldByName('BEGINNUM').AsString := sL_BeginNum;
    sL_TmpEndNum := TUstr.AddString(FloatToStr(fL_EndNum),'0',false,PAPER_NUMBER_LENGTH);
    sL_TmpEndNum := TUstr.AddString(sL_TmpEndNum,'0',false, nL_BeginNumLength);
    DataSet.FieldByName('ENDNUM').AsString := sL_TmpEndNum;

    DataSet.FieldByName('OPERATOR').AsString := frmMainMenu.sG_User;
    DataSet.FieldByName('UPDTIME').AsString := sL_Now;
end;

procedure TdtmMain3.saveToDB(I_CDS: TClientDataSet);
begin
    with I_CDS do
    begin
      if state in [dsEdit, dsInsert] then
        Post;
      if (ChangeCount > 0) then
        ApplyUpdates(0);

    end;
end;

procedure TdtmMain3.deletePaperData(nI_Seq: Integer);
begin
    with qryCommon do
    begin
      Close;
      SQL.Clear;
      SQL.Add('delete SO127 where SEQ=' + IntToStr(nI_Seq));
      ExecSQL;
      Close;
    end;
end;

function TdtmMain3.getPaperData(sI_PaperSNum, sI_PaperENum: String): TDataSet;
var
    sL_SQL,sL_Where : String;
    sL_Prifix, sL_PaperSNum, sL_PaperENum : String;
begin
    sL_PaperSNum := TUstr.AddString(sI_PaperSNum,'0',false,PAPER_NUMBER_LENGTH);
    sL_PaperENum := TUstr.AddString(sI_PaperENum,'0',false,PAPER_NUMBER_LENGTH);

    sL_SQL := 'select PAPERNUM,EMPNAME,GETPAPERDATE,BILLNO from SO127 ';
//    sL_Where := ' where PAPERNUM between to_number(' + sL_PaperSNum + ') and  to_number(' + sL_PaperENum + ')';
    sL_Where := ' where PAPERNUM between ' + STR_SEP + sL_PaperSNum + STR_SEP + ' and ' + STR_SEP + sL_PaperENum + STR_SEP;
    sL_Where := sL_Where + ' and STATUS<>0 ';
    sL_SQL := sL_SQL + sL_Where;
    with qrySo18C2 do
    begin
      Close;
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;
    end;
    result := qrySo18C2;
end;

function TdtmMain3.disablePaperData(sI_PaperSNum, sI_PaperENum: String): Boolean;
var
    sL_SQL, sL_Now : String;
    bL_Result : boolean;
    sL_Prifix, sL_PaperSNum, sL_PaperENum : STring;
begin
    sL_PaperSNum := sI_PaperSNum;
    sL_PaperENum := sI_PaperENum;

    try
      bL_Result := true;
      sL_Now := TUdateTime.GetPureChineseDateTimeStr(now);
      sL_SQL := 'update SO127 set STATUS=0';
      sL_SQL := sL_SQL + ',OPERATOR=' + STR_SEP + frmMainMenu.sG_User + STR_SEP;
      sL_SQL := sL_SQL + ',UPDTIME=' + STR_SEP + sL_Now + STR_SEP;
      sL_SQL := sL_SQL + ' where PAPERNUM between ' + STR_SEP + sL_PaperSNum + STR_SEP + ' and ' + STR_SEP + sL_PaperENum + STR_SEP;
      sL_SQL := sL_SQL + ' and STATUS<>0 and COMPCODE=' + frmMainMenu.sG_CompCode;
      with qryCommon do
      begin
        Close;
        SQL.Clear;
        SQL.Add(sL_SQL);
        ExecSQL;
        Close;
      end;

    except
      bL_Result := false;
    end;

    result := bL_Result;
end;

function TdtmMain3.activeReportData(sI_BeginDate, sI_EndDate, sI_EmpSQL,
  sI_BeginPaperNum, sI_EndPaperNum: String; nI_ReportType: Integer): boolean;
var
    sL_SQL, sL_Where, sL_GroupBy, sL_OrderBy : String;
    bL_Result : boolean;
    nL_RecordCount : Integer;
    dL_BeginDate, dL_EndDate : TDate;
    L_AdoQuery : TADOQuery;
begin
    bL_Result := true;
    sL_Where := '';
    case nI_ReportType of
     3://領單續用報表
      begin
        sL_SQL := 'select SEQ, SPAPERNUM, EPAPERNUM, TOTALPAPERCOUNT, EMPNAME,GETPAPERDATE, NOTES, OPERATOR, UPDTIME from SO129';


        if sI_BeginDate<>'' then
        begin
          dL_BeginDate := TUdateTime.CDate2Date(sI_BeginDate);
          dL_EndDate := TUdateTime.CDate2Date(sI_EndDate);
          sL_Where := ' where GETPAPERDATE between ' + TUstr.getOracleSQLDateStr(dL_BeginDate);
          sL_Where := sL_Where + ' and '  + TUstr.getOracleSQLDateStr(dL_EndDate);
        end;

        if sI_EmpSQL<>'' then
        begin
          if sL_Where='' then
            sL_Where := ' where ' +  sI_EmpSQL
          else
            sL_Where := sL_Where + ' and ' +  sI_EmpSQL;
        end;

        if sL_Where='' then
          sL_Where := ' where COMPCODE=' + frmMainMenu.sG_CompCode
        else
          sL_Where := sL_Where + ' and COMPCODE=' + frmMainMenu.sG_CompCode;

        sL_SQL := sL_SQL + sL_Where + ' order by SEQ';
        L_AdoQuery := qryReport5;

      end;
     0://手開單據領用明細表
      begin
        sL_SQL := ' SELECT EMPNO, EMPNAME, GETPAPERDATE, ' +
                  '        MIN(PAPERNUM) MIN_PAPER_NUM,  ' +
                  '        MAX(PAPERNUM) MAX_PAPER_NUM,  ' +
                  '        COUNT(SEQ) COUNT              ' +
                  '   FROM SO127 ';
        sL_GroupBy := ' GROUP BY EMPNO, EMPNAME, GETPAPERDATE ';
        if sI_BeginDate<>'' then
        begin
          dL_BeginDate := TUdateTime.CDate2Date(sI_BeginDate);
          dL_EndDate := TUdateTime.CDate2Date(sI_EndDate);
          sL_Where := ' WHERE GETPAPERDATE BETWEEN ' + TUstr.getOracleSQLDateStr(dL_BeginDate);
          sL_Where := sL_Where + ' AND '  + TUstr.getOracleSQLDateStr(dL_EndDate);
        end;

        if sI_EmpSQL<>'' then
        begin
          if sL_Where='' then
            sL_Where := ' WHERE ' +  sI_EmpSQL
          else
            sL_Where := sL_Where + ' AND ' +  sI_EmpSQL;
        end;

        if sI_BeginPaperNum<>'' then
        begin
          if sL_Where='' then
            sL_Where := ' where ' +  'PAPERNUM between ' + STR_SEP + sI_BeginPaperNum + STR_SEP + ' and  ' + STR_SEP + sI_EndPaperNum + STR_SEP
          else
            sL_Where := sL_Where + ' and ' +  'PAPERNUM between ' + STR_SEP + sI_BeginPaperNum + STR_SEP + ' and  ' + STR_SEP + sI_EndPaperNum + STR_SEP;
        end;

        if sL_Where='' then
          sL_Where := ' where CompCode=' + frmMainMenu.sG_CompCode
        else
        sL_Where := sL_Where + ' and CompCode=' + frmMainMenu.sG_CompCode;
        sL_SQL := sL_SQL + sL_Where + sL_GroupBy;
        L_AdoQuery := qryReport1;

      end;
     1://未回報手開單據清冊 => SO127 中 billno 沒有值且沒作廢(沒作廢=>Status=1)者
      begin
        sL_SQL := 'select EMPNO,EMPNAME,GETPAPERDATE,PAPERNUM from SO127 ';
//        sL_SQL := 'select distinct EMPNO,EMPNAME, GETPAPERDATE from SO127 ';
        if sI_BeginDate<>'' then
        begin
          dL_BeginDate := TUdateTime.CDate2Date(sI_BeginDate);
          dL_EndDate := TUdateTime.CDate2Date(sI_EndDate);
          sL_Where := ' where GETPAPERDATE between ' + TUstr.getOracleSQLDateStr(dL_BeginDate);
          sL_Where := sL_Where + ' and '  + TUstr.getOracleSQLDateStr(dL_EndDate);
        end;

        if sI_EmpSQL<>'' then
        begin
          if sL_Where='' then
            sL_Where := ' where ' +  sI_EmpSQL
          else
            sL_Where := sL_Where + ' and ' +  sI_EmpSQL;
        end;

        sL_OrderBy := ' order by EMPNO, GETPAPERDATE, PAPERNUM ';
        if sI_BeginPaperNum<>'' then
        begin
          if sL_Where='' then
            sL_Where := ' where ' +  'PAPERNUM between ' + STR_SEP + sI_BeginPaperNum + STR_SEP + ' and  ' + STR_SEP + sI_EndPaperNum + STR_SEP
          else
            sL_Where := sL_Where + ' and ' +  'PAPERNUM between ' + STR_SEP + sI_BeginPaperNum + STR_SEP + ' and  ' + STR_SEP + sI_EndPaperNum + STR_SEP;
        end;

        if sL_Where='' then
          sL_Where := ' where STATUS=1 and (BILLNO=' + STR_SEP + STR_SEP + ' or BILLNO IS NULL) '
        else
          sL_Where := sL_Where + ' and STATUS=1 and (BILLNO=' + STR_SEP + STR_SEP + ' or BILLNO IS NULL) ';

        sL_Where := sL_Where + ' and CompCode=' + frmMainMenu.sG_CompCode;
        sL_SQL := sL_SQL + sL_Where + sL_OrderBy;
        L_AdoQuery := qryReport2;

      end;

     2://手開單使用情形 => SO127 中 billno 有值或作廢(作廢=>Status=0)者
      begin
        sL_SQL := 'select EMPNO,EMPNAME,GETPAPERDATE,PAPERNUM, BILLNO, STATUS,' ;
        sL_SQL := sL_SQL + ' CUSTID,CUSTNAME,CUSTTEL,REALDATE  from SO127 ';
        if sI_BeginDate<>'' then
        begin
          dL_BeginDate := TUdateTime.CDate2Date(sI_BeginDate);
          dL_EndDate := TUdateTime.CDate2Date(sI_EndDate);
          sL_Where := ' where GETPAPERDATE between ' + TUstr.getOracleSQLDateStr(dL_BeginDate);
          sL_Where := sL_Where + ' and '  + TUstr.getOracleSQLDateStr(dL_EndDate);
        end;

        if sI_EmpSQL<>'' then
        begin
          if sL_Where='' then
            sL_Where := ' where ' +  sI_EmpSQL
          else
            sL_Where := sL_Where + ' and ' +  sI_EmpSQL;
        end;

        sL_OrderBy := ' order by EMPNO, GETPAPERDATE, PAPERNUM ';
        if sI_BeginPaperNum<>'' then
        begin
          if sL_Where='' then
            sL_Where := ' where ' +  'PAPERNUM between ' + STR_SEP + sI_BeginPaperNum + STR_SEP + ' and  ' + STR_SEP + sI_EndPaperNum + STR_SEP
          else
            sL_Where := sL_Where + ' and ' +  'PAPERNUM between ' + STR_SEP + sI_BeginPaperNum + STR_SEP + ' and  ' + STR_SEP + sI_EndPaperNum + STR_SEP;
        end;

        if sL_Where='' then
          sL_Where := ' where (STATUS=0 or (BILLNO<>' + STR_SEP + STR_SEP + ' or BILLNO IS NOT NULL)) '
        else
          sL_Where := sL_Where + ' and (STATUS=0 or (BILLNO<>' + STR_SEP + STR_SEP + ' or BILLNO IS NOT NULL)) ';
          
        sL_Where := sL_Where + ' and CompCode=' + frmMainMenu.sG_CompCode;
        sL_SQL := sL_SQL + sL_Where + sL_OrderBy;
        L_AdoQuery := qryReport3;

      end;
    end;

    with L_AdoQuery do
    begin
      Close;
      SQL.Clear;

      SQL.Add(sL_SQL);
//SQL.SaveToFile('c:\abc.txt');      
      Open;
      if RecordCount>0 then
        bL_Result := true
      else
        bL_Result := false;
    end;

    if (nI_ReportType=1) and (qryReport2.RecordCount>0) then
    begin
      if not cdsReport2.Active then
        cdsReport2.CreateDataSet;
      cdsReport2.EmptyDataSet;
      transReport2Data(sL_Where);
    end;
    {
    else if (nI_ReportType=2) and (qryReport3.RecordCount>0) then
    begin
      if not cdsReport3.Active then
        cdsReport3.CreateDataSet;
      cdsReport3.EmptyDataSet;
      transReport3Data;
    end;
    }
    result := bL_Result;





end;

procedure TdtmMain3.transReport2Data(sI_WhereClause:String);
var
    sL_SQL : String;
    bL_AppendData : boolean;
    nL_GroupID, ii : Integer;
    sL_PriorEmpNo,sL_PriorEmpName, sL_PriorGetPaperDate : String;
    sL_EmpNo, sL_EmpName, sL_GetPaperDate, sL_PaperNumbers : String;
begin
    sL_PriorEmpNo := '';
    sL_PriorEmpName := '';
    nL_GroupID := 1;
    with qryReport2 do
    begin
      while not Eof do
      begin
        bL_AppendData := true;
        sL_PaperNumbers := '';
        sL_EmpNo := FieldByName('EMPNO').AsString;
        sL_EmpName := FieldByName('EMPNAME').AsString;        
        sL_GetPaperDate := TUdateTime.CDateStr(FieldByName('GETPAPERDATE').AsDateTime,9);
        sL_PriorEmpNo := sL_EmpNo;
        sL_PriorEmpName := sL_EmpName;
        sL_PriorGetPaperDate := sL_GetPaperDate;
        for ii:=0 to REPORT2_PAPER_COLUMN_COUNTS-1 do
        begin

          sL_EmpNo := FieldByName('EMPNO').AsString;
          sL_EmpName := FieldByName('EMPNAME').AsString;
          sL_GetPaperDate := TUdateTime.CDateStr(FieldByName('GETPAPERDATE').AsDateTime,9);
          if (sL_PriorEmpNo<>sL_EmpNo) or (sL_PriorGetPaperDate<>sL_GetPaperDate)then
          begin
            cdsReport2.Append;
            cdsReport2.FieldByName('EMPNO').AsString := sL_PriorEmpNo;
            cdsReport2.FieldByName('EMPNAME').AsString := sL_PriorEmpName;
            cdsReport2.FieldByName('GETPAPERDATE').AsString := sL_PriorGetPaperDate;
            cdsReport2.FieldByName('PAPERNUMBERS').AsString := sL_PaperNumbers;
            cdsReport2.FieldByName('GROUPID').AsInteger := nL_GroupID;
            cdsReport2.Post;
            INC(nL_GroupID);            
            sL_PaperNumbers := '';
            sL_PriorGetPaperDate := '';
            bL_AppendData := false;
            Prior;
            break;
          end;
          if sL_PaperNumbers='' then
            sL_PaperNumbers := FieldByName('PAPERNUM').AsString
          else
            sL_PaperNumbers := sL_PaperNumbers + '  ' + FieldByName('PAPERNUM').AsString;
          sL_PriorEmpNo := sL_EmpNo;
          sL_PriorEmpName := sL_EmpName;
          sL_PriorGetPaperDate := sL_GetPaperDate;
          if ii <> (REPORT2_PAPER_COLUMN_COUNTS-1) then
            Next;
          if Eof then break;            
        end;
        if bL_AppendData then
        begin
          cdsReport2.Append;
          cdsReport2.FieldByName('EMPNO').AsString := sL_EmpNo;
          cdsReport2.FieldByName('EMPNAME').AsString := sL_EmpName;
          cdsReport2.FieldByName('GETPAPERDATE').AsString := sL_GetPaperDate;
          cdsReport2.FieldByName('PAPERNUMBERS').AsString := sL_PaperNumbers;
          cdsReport2.FieldByName('GROUPID').AsInteger := nL_GroupID;          
          cdsReport2.Post;
        end;
        Next;
      end;
    end;
    //down, 統計單據數量
    with qryReport2Common do
    begin
//      sL_SQL := 'select EMPNO, GETPAPERDATE, COUNT(PAPERNUM) COUNTS from SO127 ';
      sL_SQL := 'select EMPNO, TO_CHAR(GETPAPERDATE,' + STR_SEP + 'YYYYMMDD' + STR_SEP + ') GETPAPERDATE, COUNT(PAPERNUM) COUNTS from SO127 ';
      sL_SQL := sL_SQL  + sI_WhereClause + ' group by EMPNO, GETPAPERDATE ';
      Close;
      SQL.Clear;
      SQL.Add(sL_SQL);
//SQL.SaveToFile('c:\aaaa.txt');
      Open;
    end;
    //up, 統計單據數量
end;

procedure TdtmMain3.qryReport2CommonGETPAPERDATEGetText(Sender: TField;
  var Text: String; DisplayText: Boolean);
begin
    {
    if Sender.AsDateTime=0 then Exit;
    Text := TUdateTime.CDateStr(Sender.AsDateTime,9);
    DisplayText := true;
    }
end;

function TdtmMain3.getReport2PaperCounts(sI_EmpNo,
  sI_GetPaperDate: String): Integer;
begin
    with qryReport2Common do
    begin
      if Locate('EMPNO;GETPAPERDATE', VarArrayOf([sI_EmpNo, sI_GetPaperDate]), [loPartialKey]) then
        result := FieldByName('COUNTS').AsInteger
      else
        result := 0;
    end;
end;

function TdtmMain3.getReport2PaperCounts(sI_EmpNo:String; dI_GetPaperDate:TDate): Integer;
begin
    with qryReport2Common do
    begin
      if Locate('EMPNO;GETPAPERDATE', VarArrayOf([sI_EmpNo, dI_GetPaperDate]), [loPartialKey]) then
        result := FieldByName('COUNTS').AsInteger
      else
        result := 0;
    end;

end;

procedure TdtmMain3.transReport3Data;
begin
    with qryReport3 do
    begin
      while not Eof do
      begin

        Next;
      end;
    end;
end;


function TdtmMain3.checkDul(sI_Prefix: String; nI_BeginNumLength : Integer;fI_BeginNum,
  fI_EndNum: double): boolean;
var
    bL_Result : boolean;
    sL_SQL : String;
    sL_BeginNum, sL_EndNum : String;
begin
{
    sL_BeginNum := TUstr.AddString(FloatToStr(fI_BeginNum),'0',false,nI_BeginNumLength);
    sL_EndNum := TUstr.AddString(FloatToStr(fI_EndNum),'0',false,nI_BeginNumLength);
}
    sL_BeginNum := FloatToStr(fI_BeginNum);
    sL_EndNum := FloatToStr(fI_EndNum);

    sL_SQL := 'select count(*) CNT from SO126 ';
    sL_SQL := sL_SQL + ' where (((' + sL_BeginNum + ' between to_number(BEGINNUM) and to_number(ENDNUM)) and ';
//    sL_SQL := sL_SQL + ' where (((' + sL_BeginNum + '  between to_number(BEGINNUM) and to_number(ENDNUM)) and ';
    sL_SQL := sL_SQL + ' (' + sL_EndNum + '  between to_number(BEGINNUM) and to_number(ENDNUM))) or ';
    sL_SQL := sL_SQL + ' ((to_number(BEGINNUM)>' + sL_BeginNum + '  and to_number(ENDNUM)>' + sL_BeginNum + ' ) and ';
    sL_SQL := sL_SQL + ' (to_number(BEGINNUM)<' + sL_EndNum + '  and to_number(ENDNUM)<' + sL_EndNum + ' )) or ';
    sL_SQL := sL_SQL + ' (to_number(BEGINNUM)=' + sL_BeginNum + '  and (to_number(BEGINNUM)<' + sL_EndNum + '  and to_number(ENDNUM)<' + sL_EndNum + ' )) or ';
    sL_SQL := sL_SQL + ' ((to_number(BEGINNUM)>' + sL_BeginNum + '  and to_number(ENDNUM)>' + sL_BeginNum + ' ) and (to_number(ENDNUM)=' + sL_EndNum + ' )) or ';

    sL_SQL := sL_SQL + '(to_number(BEGINNUM)=' + sL_BeginNum + '  and to_number(ENDNUM)=' + sL_EndNum + ' ) or ';
    sL_SQL := sL_SQL + ' ((to_number(BEGINNUM)>' + sL_BeginNum + '  and to_number(ENDNUM)>' + sL_BeginNum + ' ) and (to_number(BEGINNUM)<' + sL_EndNum + '  and to_number(ENDNUM)>' + sL_EndNum + ' )) or ';
    sL_SQL := sL_SQL + ' ((to_number(BEGINNUM)<' + sL_BeginNum + '  and to_number(ENDNUM)>' + sL_BeginNum + ' ) and (to_number(BEGINNUM)<' + sL_EndNum + '  and to_number(ENDNUM)<=' + sL_EndNum + ' )) or ';
    sL_SQL := sL_SQL + ' ((to_number(BEGINNUM)>' + sL_BeginNum + '  and to_number(ENDNUM)>' + sL_BeginNum + ' ) and (to_number(BEGINNUM)=' + sL_EndNum + ' )) or ';
    sL_SQL := sL_SQL + ' ((to_number(BEGINNUM)<' + sL_BeginNum + '  and to_number(ENDNUM)=' + sL_BeginNum + ' ) and (to_number(BEGINNUM)<' + sL_EndNum + '  and to_number(ENDNUM)<' + sL_EndNum + ' ))) ';
    if sI_Prefix<>'' then
      sL_SQL := sL_SQL + ' and PREFIX=' + STR_SEP + sI_Prefix + STR_SEP
    else
      sL_SQL := sL_SQL + ' and (PREFIX=' + STR_SEP + sI_Prefix + STR_SEP + ' or PREFIX IS NULL)';    

    {
    sL_SQL := 'select count(*) CNT from SO126 ';
    sL_SQL := sL_SQL + ' where (((V1 between to_number(BEGINNUM) and to_number(ENDNUM)) and ';
    sL_SQL := sL_SQL + ' (V2 between to_number(BEGINNUM) and to_number(ENDNUM))) or ';
    sL_SQL := sL_SQL + ' ((to_number(BEGINNUM)>V1 and to_number(ENDNUM)>V1) and ';
    sL_SQL := sL_SQL + ' (to_number(BEGINNUM)<V2 and to_number(ENDNUM)<V2)) or ';
    sL_SQL := sL_SQL + ' (to_number(BEGINNUM)=V1 and (to_number(BEGINNUM)<V2 and to_number(ENDNUM)<V2)) or ';
    sL_SQL := sL_SQL + ' ((to_number(BEGINNUM)>V1 and to_number(ENDNUM)>V1) and (to_number(ENDNUM)=V2)) or ';

    sL_SQL := sL_SQL + '(to_number(BEGINNUM)=V1 and to_number(ENDNUM)=V2) or ';
    sL_SQL := sL_SQL + ' ((to_number(BEGINNUM)>V1 and to_number(ENDNUM)>V1) and (to_number(BEGINNUM)<V2 and to_number(ENDNUM)>V2)) or ';
    sL_SQL := sL_SQL + ' ((to_number(BEGINNUM)<V1 and to_number(ENDNUM)>V1) and (to_number(BEGINNUM)<V2 and to_number(ENDNUM)<=V2)) or ';
    sL_SQL := sL_SQL + ' ((to_number(BEGINNUM)>V1 and to_number(ENDNUM)>V1) and (to_number(BEGINNUM)=V2)) or ';
    sL_SQL := sL_SQL + ' ((to_number(BEGINNUM)<V1 and to_number(ENDNUM)=V1) and (to_number(BEGINNUM)<V2 and to_number(ENDNUM)<V2))) ';
    sL_SQL := sL_SQL + ' and PREFIX=' + STR_SEP + sI_Prefix + STR_SEP;

    }
    with qryCommon do
    begin
      SQL.Clear;
      SQL.Add(sL_SQL);
      Close;
      Open;
      if FieldByName('CNT').AsInteger >0 then
        bL_Result := true
      else
        bL_Result := false;
      Close;          
    end;
    result := bL_Result;
end;

function TdtmMain3.getSo034Data(sI_BillNo: String): Integer;
var
    sL_SQL : String;
    nL_RecCount : Integer;
begin
    sL_SQL := 'select a.CustId,b.CustName, b.Tel1, a.BillNo,a.ManualNo,a.Item,a.CitemName,a.OldAmt,a.ShouldAmt,a.RealDate, a.RealAmt, a.ClctName,a.CMName,a.UCName,a.STName from SO034 a, SO001 b where a.BillNo=' + STR_SEP + sI_BillNo + STR_SEP;
    sL_SQL := sL_SQL + ' and a.CompCode=' + frmMainMenu.sG_CompCode;
    sL_SQL := sL_SQL + ' and a.CompCode=b.CompCode and a.CustID=b.CustID';

    with qrySo034 do
    begin
      Close;
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;
      nL_RecCount := RecordCount;
      if nL_RecCount=0 then
        Close;
    end;

    result := nL_RecCount;
end;


function TdtmMain3.setNewHandPaperNo(sI_CustID,sI_CustName, sI_Tel1, sI_Item:String; dI_RealDate : TDate; sI_BillNo, sI_OldHandPaperNo,
  sI_NewHandPaperNo: String; var sI_ErrorMsg:String): boolean;
    procedure updateSo034(nI_OperType:Integer);
    var
        sL_SQL : String;
        sL_NewManualNo : String;
    begin
      //down, update So034 之 data
      case nI_OperType of
        1,3 :
          sL_NewManualNo := sI_NewHandPaperNo;
        2:
          sL_NewManualNo := '';        
      end;

      sL_SQL := 'update SO034 set ManualNo=' + STR_SEP + sL_NewManualNo + STR_SEP + ' where BillNo=' + STR_SEP + sI_BillNo + STR_SEP;
      sL_SQL := sL_SQL + 'and Item=' + sI_Item + ' and CompCode=' + frmMainMenu.sG_CompCode;
      with qryCommon do
      begin
        Close;
        SQL.Clear;
        SQL.Add(sL_SQL);
        ExecSQL;
        Close;
      end;
      //up, update So034 之 data

    end;

    procedure writeSo128(sI_HandPaperNo, sI_NewCustID, sI_NewCustName, sI_NewCustTel, sI_NewBillNo, sI_NewRealDate:String);
    var
      sL_OldCustID, sL_OldCustName, sL_OldCustTel, sL_OldBillNo, sL_OldRealDate:String;
      sL_SQL, sL_UpdateStatus, sL_NewRealDate : String;
    begin
      sL_SQL := 'select CUSTID, CUSTNAME,CUSTTEL,BILLNO,REALDATE from SO127';
      sL_SQL := sL_SQL + ' where PAPERNUM=' + STR_SEP + sI_HandPaperNo + STR_SEP;
      sL_SQL := sL_SQL + ' and COMPCODE=' + frmMainMenu.sG_CompCode;
      with qryCommon do
      begin
        Close;
        SQL.Clear;
        SQL.Add(sL_SQL);
        Open;
        if RecordCount=1 then
        begin
          sL_OldCustID := FieldByName('CUSTID').AsString;
          sL_OldCustName := FieldByName('CUSTNAME').AsString;
          sL_OldCustTel := FieldByName('CUSTTEL').AsString;
          sL_OldBillNo := FieldByName('BILLNO').AsString;
          sL_OldRealDate := FieldByName('REALDATE').AsString;
          Close;

          sL_UpdateStatus := '';
          if sL_OldCustID<>sI_NewCustID then
          begin
            if sL_OldCustID='' then
              sL_OldCustID := 'NULL';
            if sI_NewCustID='' then
              sI_NewCustID := 'NULL';
            sL_UpdateStatus := 'CustID' + SO128_FIELD_VALUE_SEP + sL_OldCustID + SO128_FIELD_VALUE_SEP + sI_NewCustID;
          end;

          if sL_OldCustName<>sI_NewCustName then
          begin
            if sL_OldCustName='' then
              sL_OldCustName := 'NULL';
            if sI_NewCustName='' then
              sI_NewCustName := 'NULL';
            if sL_UpdateStatus='' then
              sL_UpdateStatus := 'CustName'+ SO128_FIELD_VALUE_SEP + sL_OldCustName + SO128_FIELD_VALUE_SEP + sI_NewCustName
            else
              sL_UpdateStatus := sL_UpdateStatus + SO128_FIELD_SEP + 'CustName' + SO128_FIELD_VALUE_SEP + sL_OldCustName + SO128_FIELD_VALUE_SEP + sI_NewCustName;
          end;

          if sL_OldCustTel<>sI_NewCustTel then
          begin
            if sL_OldCustTel='' then
              sL_OldCustTel := 'NULL';
            if sI_NewCustTel='' then
              sI_NewCustTel := 'NULL';
            if sL_UpdateStatus='' then
              sL_UpdateStatus := 'CustTel' + SO128_FIELD_VALUE_SEP + sL_OldCustTel + SO128_FIELD_VALUE_SEP + sI_NewCustTel
            else
              sL_UpdateStatus := sL_UpdateStatus + SO128_FIELD_SEP + 'CustTel' + SO128_FIELD_VALUE_SEP + sL_OldCustTel + SO128_FIELD_VALUE_SEP + sI_NewCustTel;
          end;

          if sL_OldBillNo<>sI_NewBillNo then
          begin
            if sL_OldBillNo='' then
              sL_OldBillNo := 'NULL';
            if sI_NewBillNo='' then
              sI_NewBillNo := 'NULL';
            if sL_UpdateStatus='' then
              sL_UpdateStatus := 'BillNo' + SO128_FIELD_VALUE_SEP + sL_OldBillNo + SO128_FIELD_VALUE_SEP + sI_NewBillNo
            else
              sL_UpdateStatus := sL_UpdateStatus + SO128_FIELD_SEP + 'BillNo' + SO128_FIELD_VALUE_SEP+ sL_OldBillNo + SO128_FIELD_VALUE_SEP + sI_NewBillNo;
          end;

          if sL_OldRealDate<>sI_NewRealDate then
          begin
            if sI_NewRealDate='' then
              sL_NewRealDate := 'NULL'
            else
              sL_NewRealDate := sI_NewRealDate;

            if sL_OldRealDate='' then
              sL_OldRealDate := 'NULL';

            if sL_UpdateStatus='' then
              sL_UpdateStatus := 'RealDate' + SO128_FIELD_VALUE_SEP + sL_OldRealDate + SO128_FIELD_VALUE_SEP + sL_NewRealDate
            else
              sL_UpdateStatus := sL_UpdateStatus + SO128_FIELD_SEP + 'RealDate' + SO128_FIELD_VALUE_SEP + sL_OldRealDate + SO128_FIELD_VALUE_SEP + sL_NewRealDate;
          end;
          sL_SQL := 'insert into SO128(COMPCODE,PAPERNUM,UPDATESTATUS,OPERATOR,UPDTIME) values';
          sL_SQL := sL_SQL + '(' + frmMainMenu.sG_CompCode + ',' + STR_SEP + sI_HandPaperNo + STR_SEP;
          sL_SQL := sL_SQL + ',' + STR_SEP + sL_UpdateStatus + STR_SEP;
          sL_SQL := sL_SQL + ',' + STR_SEP + frmMainMenu.sG_User + STR_SEP;
          sL_SQL := sL_SQL + ',' + TUstr.getOracleSQLDateTimeStr(now) + ')';
          Close;
          SQL.Clear;
          SQL.Add(sL_SQL);
          ExecSQL;
          Close;

        end
        else
          Close;
      end;
    end;

    procedure updateSo127(nI_OperType:Integer);
    var
        sL_SQL : String;
    begin
      if nI_OperType in [1,3] then
      begin
        writeSo128(sI_NewHandPaperNo,sI_CustID, sI_CustName,sI_Tel1, sI_BillNo,DateToStr(dI_RealDate));      
        //down, 更新新對應之手開單號的 data
        sL_SQL := 'update SO127 set CUSTID=' + sI_CustID +', CUSTNAME=' + STR_SEP + sI_CustName + STR_SEP ;
        sL_SQL := sL_SQL + ',CUSTTEL=' + STR_SEP +sI_Tel1 + STR_SEP + ', BILLNO=' + STR_SEP + sI_BillNo + STR_SEP;
        sL_SQL := sL_SQL + ',REALDATE=' + TUstr.getOracleSQLDateStr(dI_RealDate);
        sL_SQL := sL_SQL + ',OPERATOR=' + STR_SEP + frmMainMenu.sG_User + STR_SEP;
        sL_SQL := sL_SQL + ',UPDTIME=' + TUstr.getOracleSQLDateTimeStr(now);
        sL_SQL := sL_SQL + ' where COMPCODE=' + frmMainMenu.sG_CompCode;
        sL_SQL := sL_SQL + ' and PAPERNUM=' + STR_SEP + sI_NewHandPaperNo + STR_SEP;

        with qryCommon do
        begin
          Close;
          SQL.Clear;
          SQL.Add(sL_SQL);
          ExecSQL;
          Close;
        end;

        //up, 更新新對應之手開單號的 data
      end;

      if nI_OperType in [1,2] then
      begin
        writeSo128(sI_OldHandPaperNo,'', '','','','');      
        //down, 清除舊手開單號之data
        sL_SQL := 'update SO127 set CUSTID=NULL,CUSTNAME=NULL,CUSTTEL=NULL,BILLNO=NULL,REALDATE=NULL';
        sL_SQL := sL_SQL + ',OPERATOR=' + STR_SEP + frmMainMenu.sG_User + STR_SEP;
        sL_SQL := sL_SQL + ',UPDTIME=' + TUstr.getOracleSQLDateTimeStr(now);
        sL_SQL := sL_SQL + ' where COMPCODE=' + frmMainMenu.sG_CompCode;
        sL_SQL := sL_SQL + ' and PAPERNUM=' + STR_SEP + sI_OldHandPaperNo + STR_SEP;

        with qryCommon do
        begin
          Close;
          SQL.Clear;
          SQL.Add(sL_SQL);
          ExecSQL;
          Close;
        end;
        //up, 清除舊手開單號之data

      end;
    end;


    function validateNewHandPaperNo(sI_PaperNo:String; var sI_Msg:String):boolean;
    var
        sL_SQL, sL_BillNo : String;
        bL_Result : boolean;
    begin
      sI_Msg := '';
      bL_Result := true;
      sL_SQL := 'select BILLNO, STATUS from SO127 where PAPERNUM=' + STR_SEP + sI_PaperNo + STR_SEP;
      sL_SQL := sL_SQL + ' and COMPCODE=' + frmMainMenu.sG_CompCode;
      with qryCommon do
      begin
        Close;
        SQL.Clear;
        SQL.Add(sL_SQL);
        Open;
        if RecordCount=0 then
        begin
          sI_Msg := '此手開單號(' +sI_PaperNo + ')不存在!';
          bL_Result := false;
        end
        else
        begin
          sL_BillNo := FieldByName('BILLNO').AsString;
//          if (sL_BillNo<>'')then          
          if (sL_BillNo<>'') and (sL_BillNo<>sI_BillNo)then
          begin
            sI_Msg := '此手開單號('+sI_PaperNo+')已經有對應之單據編號:' + sL_BillNo + '!';
            bL_Result := false;
          end
          else
          begin
            if FieldByName('STATUS').AsInteger =0 then
            begin
              sI_Msg := '此手開單號('+sI_PaperNo+')已經做廢!';
              bL_Result := false;
            end;
          end;
        end;
        Close;
      end;
      result := bL_Result;
    end;
        
var
    sL_Msg : String;
    bL_Result : boolean;
    nL_OperType : Integer;
begin
    bL_Result := true;
    if (sI_OldHandPaperNo<>'') and (sI_NewHandPaperNo<>'') then
      nL_OperType := 1
    else if (sI_OldHandPaperNo<>'') and (sI_NewHandPaperNo='') then
      nL_OperType := 2
    else if (sI_OldHandPaperNo='') and (sI_NewHandPaperNo<>'') then
      nL_OperType := 3;

    case nL_OperType of
      1: //(sI_OldHandPaperNo<>'') and (sI_NewHandPaperNo<>'')
       begin
         if validateNewHandPaperNo(sI_NewHandPaperNo, sL_Msg) then
         begin
           updateSo127(nL_OperType);
           updateSo034(nL_OperType);
         end
         else
         begin
           sI_ErrorMsg := sL_Msg;
           bL_Result := false;
         end;
       end;
      2: //(sI_OldHandPaperNo<>'') and (sI_NewHandPaperNo='')
       begin
           updateSo127(nL_OperType);
           updateSo034(nL_OperType);
       end;
      3: //(sI_OldHandPaperNo='') and (sI_NewHandPaperNo<>'')
       begin
         if validateNewHandPaperNo(sI_NewHandPaperNo, sL_Msg) then
         begin
           updateSo127(nL_OperType);
           updateSo034(nL_OperType);           
         end
         else
         begin
           sI_ErrorMsg := sL_Msg;
           bL_Result := false;
         end;
       end;
    end;

    result := bL_Result;
end;

function TdtmMain3.getSo128Data(dI_SDate, dI_EDate: TDate): Integer;
    procedure transReport4CDS(I_DateSet:TDataSet);
    var
      ii,nL_Group : Integer;
      L_FieldStrList, L_FieldValueStrList : TStringList;
      sL_PaperNum, sL_UpdateStatus, sL_Operator, sL_UpdTime : String;
      sL_FieldName, sL_OldValue, sL_NewValue : String;
    begin
      if cdsReport4.State = dsInactive then
        cdsReport4.CreateDataSet;
      cdsReport4.EmptyDataSet;

      nL_Group := 0;
      with I_DateSet do
      begin
        First;
        while not Eof do
        begin
          sL_PaperNum := FieldByNAme('PAPERNUM').AsString;
          sL_UpdateStatus := FieldByNAme('UPDATESTATUS').AsString;
          sL_Operator := FieldByNAme('OPERATOR').AsString;
          sL_UpdTime := FieldByNAme('UPDTIME').AsString;
          L_FieldStrList := TUStr.ParseStrings(sL_UpdateStatus, SO128_FIELD_SEP);
          if L_FieldStrList<> nil then
          begin
            Inc(nL_Group);
            for ii:=0 to L_FieldStrList.Count -1 do
            begin
              L_FieldValueStrList := TUStr.ParseStrings(L_FieldStrList.Strings[ii], SO128_FIELD_VALUE_SEP);
              sL_FieldName := L_FieldValueStrList.Strings[0];
              sL_OldValue := L_FieldValueStrList.Strings[1];
              sL_NewValue := L_FieldValueStrList.Strings[2];


              cdsReport4.Append;
              cdsReport4.FieldByName('Group').AsInteger := nL_Group;
              cdsReport4.FieldByName('PaperNum').AsString := sL_PaperNum;
              cdsReport4.FieldByName('FieldName').AsString := sL_FieldName;
              if sL_OldValue='NULL' then
                sL_OldValue := '';
              cdsReport4.FieldByName('OldValue').AsString := sL_OldValue;

              if sL_NewValue='NULL' then
                sL_NewValue := '';
              cdsReport4.FieldByName('NewValue').AsString := sL_NewValue;

              cdsReport4.FieldByName('Operator').AsString := sL_Operator;
              cdsReport4.FieldByName('UpdTime').AsString := sL_UpdTime;
              cdsReport4.Post;
            end;
          end;
          Next;
        end;
      end;
{       　
CustID:NULL-164782#CustName:NULL-洪進添#CustTel:NULL-5203417#BillNo:NULL-200106I
C0144346#RealDate:NULL-2001/6/27
}
    end;
var
    nL_RecCount : Integer;
    sL_SQL : String;
    dL_SDate, dL_EDate: TDate;
begin
    dL_SDate := StrToDateTime(dateToStr(dI_SDate) + ' 00:00:01');
    dL_EDate := StrToDateTime(dateToStr(dI_EDate) + ' 23:59:59');

    sL_SQL := 'select * from SO128 where UPDTIME between ' + TUstr.getOracleSQLDateTimeStr(dL_SDate) + ' and ' + TUstr.getOracleSQLDateTimeStr(dL_EDate) + ' order by PAPERNUM, UPDTIME' ;
    with qryReport4 do
    begin
      Close;
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;
      nL_RecCount := RecordCount;
      if nL_RecCount=0 then
        Close
      else
        transReport4CDS(qryReport4);
    end;

    result := nL_RecCount;

end;

function TdtmMain3.reusePaper(dI_OriginalGetPaperDate: TDate;
  sI_OriginalEmpNo, sI_OriginalEmpName, sI_OriginalBeginNum,
  sI_OriginalSeq, sI_CompCode, sI_EmpNo, sI_EmpName, sI_Prefix,
  sI_BeginNum, sI_EndNum, sI_GetPaperUser, sI_GetPaperDate,
  sI_OriginalCount, sI_Count: String): String;
var
    nL_Seq1, nL_Seq2, nL_Seq3 : Integer;
    bL_StartAction : boolean;
    nL_Seq,ii : Integer;
    sL_SQL, sL_TmpMsg, sL_Result, sL_Now : String;
begin
    sL_Result := '';
    sL_Now := TUdateTime.GetPureChineseDateTimeStr(now);

    try
      //down, 檢查此 SEQ 是否已經有做廢之手開單...
      with qryCommon do
      begin
        Close;
        SQL.Clear;
        SQL.Add('select PAPERNUM from  SO127 where STATUS=0 and seq=' + sI_OriginalSeq );
        SQL.Add(' and PAPERNUM between ' + STR_SEP + sI_Prefix + sI_BeginNum  + STR_SEP );
        SQL.Add(' and ' + STR_SEP + sI_Prefix + sI_EndNum + STR_SEP );
        Open;

        if RecordCount=0 then
          bL_StartAction := true
        else
        begin
          sL_TmpMsg := '手開單 ' + FieldByName('PAPERNUM').AsString + '  已經做廢!所以無法執行此作業.';
          bL_StartAction := false;
        end;
        Close;
      end;
      //up, 檢查此 SEQ 是否已經有做廢之手開單...

      if bL_StartAction then
      begin
        //down, 修改原先的資料(So126)...
        with qryCommon do
        begin
          Close;
          SQL.Clear;
//          TUstr.AddString(IntToStr(StrToInt(sI_BeginNum)-1),'0', false, length(sI_BeginNum))
//          SQL.Add('update SO126 set ENDNUM=' + STR_SEP + IntToStr(StrToInt(sI_BeginNum)-1) + STR_SEP );
          SQL.Add('update SO126 set ENDNUM=' + STR_SEP + TUstr.AddString(IntToStr(StrToInt(sI_BeginNum)-1),'0', false, length(sI_BeginNum)) + STR_SEP );
          SQL.Add(', TOTALPAPERCOUNT=TOTALPAPERCOUNT-' + sI_Count);
          SQL.Add(', OPERATOR=' + STR_SEP + frmMainMenu.sG_User + STR_SEP );
          SQL.Add(', UPDTIME=' + STR_SEP + sL_Now + STR_SEP );
          SQL.Add('where SEQ=' + sI_OriginalSeq);
          sL_TmpMsg := SQL.Text;
    //      SQL.SaveToFile('c:\a1.txt');
          ExecSQL;
          Close;
        end;
        //up, 修改原先的資料(So126)...

        //down, 新增領用資料到So126中...
        nL_Seq := getSeq;
        with qryCommon do
        begin
          Close;
          SQL.Clear;
          SQL.Add('insert into SO126(COMPCODE,SEQ,EMPNO,EMPNAME,GETPAPERDATE,ENDNUM,BEGINNUM,');
          SQL.Add('TOTALPAPERCOUNT,OPERATOR,UPDTIME,PREFIX)' );
          SQL.Add(' values(' +  frmMainMenu.sG_CompCode + ',' + IntToStr(nL_Seq) + ',' );
          SQL.Add(STR_SEP + sI_EmpNo + STR_SEP + ',' + STR_SEP + sI_GetPaperUser + STR_SEP + ',' );
          SQL.Add(TUstr.getOracleSQLDateStr(TUdateTime.CDate2Date(sI_GetPaperDate)) + ',' );

          SQL.Add(STR_SEP + sI_EndNum + STR_SEP + ',' + STR_SEP + sI_BeginNum + STR_SEP + ',' );
          SQL.Add(sI_Count + ',' + STR_SEP + frmMainMenu.sG_User + STR_SEP + ',' );
          SQL.Add(STR_SEP + sL_Now + STR_SEP + ','  + STR_SEP +  sI_Prefix + STR_SEP +')' );
    //      SQL.SaveToFile('c:\a2.txt');
          sL_TmpMsg := SQL.Text;
          ExecSQL;
          Close;
        end;
        //up, 新增領用資料到So126中...

        //down...update So127 的資料...
        for ii:= StrToInt(sI_BeginNum) to StrToInt(sI_EndNum) do
        begin
          sL_SQL := 'update SO127 set EMPNO=' + STR_SEP + sI_EmpNo + STR_SEP + ',' ;
          sL_SQL := sL_SQL + 'EMPNAME=' + STR_SEP + sI_EmpName + STR_SEP + ',' ;
          sL_SQL := sL_SQL + 'GETPAPERDATE=' + TUstr.getOracleSQLDateStr(TUdateTime.CDate2Date(sI_GetPaperDate)) + ',';
          sL_SQL := sL_SQL + 'OPERATOR=' + STR_SEP + frmMainMenu.sG_User + STR_SEP  + ',';
          sL_SQL := sL_SQL + 'UPDTIME=' + STR_SEP + sL_Now + STR_SEP + ',' ;
          sL_SQL := sL_SQL + 'SEQ=' + IntToStr(nL_Seq) ;


//          sL_SQL := sL_SQL + ' where PAPERNUM=' + STR_SEP + sI_Prefix + IntToStr(ii) + STR_SEP;
          sL_SQL := sL_SQL + ' where PAPERNUM=' + STR_SEP + sI_Prefix + TUstr.AddString(IntToStr(ii),'0', false, length(sI_BeginNum)) + STR_SEP;
          sL_TmpMsg := sL_SQL;
          with qryCommon do
          begin
            Close;
            SQL.Clear;
            SQL.Add(sL_SQL);
//            SQL.SaveToFile('c:\a3.txt');
            ExecSQL;
            Close;
          end;
        end;
        //up...update So127 的資料...

        //down...insert log data 到 So129...
        nL_Seq1 := get129Seq; //給來源手開單組使用...
        nL_Seq2 := get129Seq; //給續用前段之手開單使用...
        nL_Seq3 := get129Seq; //給續用後段之手開單使用...

        //1. insert 來源手開單資料...
        sL_SQL := 'insert into SO129(SEQ,COMPCODE,SPAPERNUM,EPAPERNUM,TOTALPAPERCOUNT, EMPNO, EMPNAME,GETPAPERDATE,NOTES,OPERATOR,UPDTIME)';
        sL_SQL := sL_SQL + ' values(' + IntToStr(nL_Seq1) + ',' + frmMainMenu.sG_CompCode + ',' ;
        sL_SQL := sL_SQL + STR_SEP + sI_Prefix + sI_OriginalBeginNum + STR_SEP + ',';
        sL_SQL := sL_SQL + STR_SEP + sI_Prefix + sI_EndNum + STR_SEP + ',';
        sL_SQL := sL_SQL + STR_SEP + sI_OriginalCount + STR_SEP + ',';
        sL_SQL := sL_SQL + STR_SEP + sI_OriginalEmpNo + STR_SEP + ',';
        sL_SQL := sL_SQL + STR_SEP + sI_OriginalEmpName + STR_SEP + ',';
        sL_SQL := sL_SQL + TUstr.getOracleSQLDateStr(dI_OriginalGetPaperDate) +',' ;
        sL_SQL := sL_SQL + STR_SEP + '來源手開單' + STR_SEP +',' ;
        sL_SQL := sL_SQL + STR_SEP + frmMainMenu.sG_User + STR_SEP  + ',';
        sL_SQL := sL_SQL + STR_SEP + sL_Now + STR_SEP + ')' ;
        sL_TmpMsg := sL_SQL;        
        with qryCommon do
        begin
          Close;
          SQL.Clear;
          SQL.Add(sL_SQL);
//          SQL.SaveToFile('c:\a4.txt');
          ExecSQL;
          Close;
        end;


        //2. insert 續用前段手開單資料...
        sL_SQL := 'insert into SO129(SEQ,COMPCODE,SPAPERNUM,EPAPERNUM,TOTALPAPERCOUNT, EMPNO, EMPNAME,GETPAPERDATE,NOTES,OPERATOR,UPDTIME)';
        sL_SQL := sL_SQL + ' values(' + IntToStr(nL_Seq2) + ',' + frmMainMenu.sG_CompCode + ',' ;
        sL_SQL := sL_SQL + STR_SEP + sI_Prefix + sI_OriginalBeginNum + STR_SEP + ',';
        sL_SQL := sL_SQL + STR_SEP + sI_Prefix + TUstr.AddString(IntToStr(StrToInt(sI_BeginNum)-1),'0', false, length(sI_BeginNum)) + STR_SEP + ',';
        sL_SQL := sL_SQL + STR_SEP + IntToStr(StrToInt(sI_OriginalCount)-StrToInt(sI_Count)) + STR_SEP + ',';        
        sL_SQL := sL_SQL + STR_SEP + sI_OriginalEmpNo + STR_SEP + ',';
        sL_SQL := sL_SQL + STR_SEP + sI_OriginalEmpName + STR_SEP + ',';
        sL_SQL := sL_SQL + TUstr.getOracleSQLDateStr(dI_OriginalGetPaperDate) +',' ;
        sL_SQL := sL_SQL + STR_SEP + '續用前段手開單--' + '原始序號:' + IntToStr(nL_Seq1) + STR_SEP +',' ;
        sL_SQL := sL_SQL + STR_SEP + frmMainMenu.sG_User + STR_SEP  + ',';
        sL_SQL := sL_SQL + STR_SEP + sL_Now + STR_SEP + ')' ;
        sL_TmpMsg := sL_SQL;
        
        with qryCommon do
        begin
          Close;
          SQL.Clear;
          SQL.Add(sL_SQL);
//          SQL.SaveToFile('c:\a5.txt');
          ExecSQL;
          Close;
        end;

        //3. insert 續用後段手開單資料...
        sL_SQL := 'insert into SO129(SEQ,COMPCODE,SPAPERNUM,EPAPERNUM,TOTALPAPERCOUNT, EMPNO, EMPNAME,GETPAPERDATE,NOTES,OPERATOR,UPDTIME)';
        sL_SQL := sL_SQL + ' values(' + IntToStr(nL_Seq3) + ',' + frmMainMenu.sG_CompCode + ',' ;
        sL_SQL := sL_SQL + STR_SEP + sI_Prefix + sI_BeginNum + STR_SEP + ',';
        sL_SQL := sL_SQL + STR_SEP + sI_Prefix + sI_EndNum + STR_SEP + ',';
        sL_SQL := sL_SQL + STR_SEP + sI_Count + STR_SEP + ',';        
        sL_SQL := sL_SQL + STR_SEP + sI_EmpNo + STR_SEP + ',';
        sL_SQL := sL_SQL + STR_SEP + sI_EmpName + STR_SEP + ',';
        sL_SQL := sL_SQL + TUstr.getOracleSQLDateStr(TUdateTime.CDate2Date(sI_GetPaperDate)) +',' ;
        sL_SQL := sL_SQL + STR_SEP + '續用後段手開單--' + '原始序號:' + IntToStr(nL_Seq1) + STR_SEP +',' ;
        sL_SQL := sL_SQL + STR_SEP + frmMainMenu.sG_User + STR_SEP  + ',';
        sL_SQL := sL_SQL + STR_SEP + sL_Now + STR_SEP + ')' ;
        sL_TmpMsg := sL_SQL;
        
        with qryCommon do
        begin
          Close;
          SQL.Clear;
          SQL.Add(sL_SQL);
//          SQL.SaveToFile('c:\a6.txt');
          ExecSQL;
          Close;
        end;
        //up...insert log data 到 So129...

      end
      else
        sL_Result := sL_TmpMsg;
    except
      on E: Exception do
       begin
         sL_Result := '作業有誤.' + '  SQL => ' + sL_TmpMsg + '  錯誤訊息' + E.Message;
       end;
    end;

    result := sL_Result;

end;

function TdtmMain3.get129Seq: Integer;
var
    sL_SQL : String;
begin
    with qryCommon do
    begin
      Close;
      sL_SQL := 'SELECT S_So129_Seq.Nextval  FROM DUAL';
      SQL.Clear;
      SQL.Add(sL_SQL);
      Open;
      result := FieldByName('Nextval').AsInteger;
      Close;
    end;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMain3.cdsSo126GETPAPERDATEGetText(Sender: TField;
  var Text: String; DisplayText: Boolean);
begin
  if Sender.AsDateTime=0 then Exit;
  Text := FormatDateTime( 'yyyy/mm/dd', Sender.AsDateTime );
  DisplayText := true;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMain3.cdsSo126RETURNDATEGetText(Sender: TField;
  var Text: String; DisplayText: Boolean);
begin
  if Sender.AsDateTime=0 then Exit;
  Text := FormatDateTime( 'yyyy/mm/dd', Sender.AsDateTime );
  DisplayText := true;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMain3.cdsSo126CLEARDATEGetText(Sender: TField;
  var Text: String; DisplayText: Boolean);
begin
  if Sender.AsDateTime=0 then Exit;
  Text := FormatDateTime( 'yyyy/mm/dd', Sender.AsDateTime );
  DisplayText := true;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMain3.qrySo18C2GETPAPERDATEGetText(Sender: TField;
  var Text: String; DisplayText: Boolean);
begin
  if Sender.AsDateTime=0 then Exit;
  Text := FormatDateTime( 'yyyy/mm/dd', Sender.AsDateTime );
  DisplayText := true;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMain3._8C30ToExcel(const aReportIndex: Integer;
  const aPaperDate, aPaperNumber, aEmpNames: String; aFile: TFileName);
var
  aText: String;
begin
  case aReportIndex of
    0:
      begin
        qryReport1.FieldByName( 'EMPNO' ).DisplayLabel := '員工編號';
        qryReport1.FieldByName( 'EMPNAME' ).DisplayLabel := '領單人員';
        qryReport1.FieldByName( 'GETPAPERDATE' ).DisplayLabel := '領單日期';
        qryReport1.FieldByName( 'MIN_PAPER_NUM' ).DisplayLabel := '單據編號(起)';
        qryReport1.FieldByName( 'MAX_PAPER_NUM' ).DisplayLabel :=  '單據編號(迄)';
        qryReport1.FieldByName( 'COUNT' ).DisplayLabel := '數量';
        aText :=
          '報表名稱: 手開單據領用明細表' + '@' +
          '統計時間: ' + DateTimeToStr( Now ) + '@' + '執行人員:' + frmMainMenu.sG_User+'@' +
          '公司名稱: ' + TUstr.replaceStr( frmMainMenu.sG_CompName, '昇', '升' ) + '@' +
          '領單日期範圍: ' +  aPaperDate + '@' +
          '領單人員: ' + aEmpNames + '@' +
          '單號範圍: ' + aPaperNumber;
         DataSetToXLS( qryReport1, aFile, aText, EmptyStr );
      end;
    1:
      begin
        qryReport1.FieldByName( 'EMPNO' ).DisplayLabel := '員工編號';
        qryReport1.FieldByName( 'EMPNAME' ).DisplayLabel := '領單人員';
        qryReport1.FieldByName( 'GETPAPERDATE' ).DisplayLabel := '領單日期';
        qryReport1.FieldByName( 'PAPERNUM' ).DisplayLabel := '單據編號';
        aText :=
          '報表名稱: 未回報手開單據清冊' + '@' +
          '統計時間: ' + DateTimeToStr( Now ) + '@' + '執行人員:' + frmMainMenu.sG_User+'@' +
          '公司名稱: ' + TUstr.replaceStr( frmMainMenu.sG_CompName, '昇', '升' ) + '@' +
          '領單日期範圍: ' +  aPaperDate + '@' +
          '領單人員: ' + aEmpNames + '@' +
          '單號範圍: ' + aPaperNumber;
         DataSetToXLS( qryReport2, aFile, aText, EmptyStr );
      end;  
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
