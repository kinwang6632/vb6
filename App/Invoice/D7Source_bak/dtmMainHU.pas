unit dtmMainHU;

interface                                            

uses
  Windows, SysUtils, Classes, DB, ADODB, DBClient, Variants, Provider, Dialogs,
  xmldom, XMLIntf, msxmldom, XMLDoc, frmInvE01U, Forms;

type
  TdtmMainH = class(TDataModule)
    adoComm: TADOQuery;
    adoA07: TADOQuery;
    cdsInvFormat: TClientDataSet;
    cdsTaxType: TClientDataSet;
    adoInv019: TADOQuery;
    adoInv099: TADOQuery;
    adoInv099InvFormatDesc: TStringField;
    adoInv019TITLESNAME: TStringField;
    adoInv019INVTITLE: TStringField;
    adoInv019BUSINESSID: TStringField;
    adoInv019ZIPCODE: TStringField;
    adoInv019MAILADDRESS: TStringField;
    adoInv019INVADDRESS: TStringField;
    adoInv019MEMO: TStringField;
    adoInv007: TADOQuery;
    adoInv005: TADOQuery;
    adoInv005ITEMID: TIntegerField;
    adoInv005DESCRIPTION: TStringField;
    adoInv011: TADOQuery;
    adoInv013: TADOQuery;
    adoInv005TAXNAME: TStringField;
    adoInv099YEARMONTH: TStringField;
    adoInv099INVFORMAT: TStringField;
    adoInv099PREFIX: TStringField;
    adoInv099STARTNUM: TStringField;
    adoInv099ENDNUM: TStringField;
    adoInv099CURNUM: TStringField;
    adoInv099LINVDATE: TStringField;
    adoInv099MEMO: TStringField;
    adoInv005SIGN: TStringField;
    adoA02_Master: TADOQuery;
    adoA02_Detail: TADOQuery;
    adoEinv: TADOQuery;
    procedure DataModuleCreate(Sender: TObject);
    procedure adoA07CalcFields(DataSet: TDataSet);
    procedure adoInv099CalcFields(DataSet: TDataSet);
  private
    { Private declarations }
    procedure setInvFormat;
    procedure setTaxType;
  public
    { Public declarations }
    procedure initCompetenceItem(sI_GroupID:String; I_Frm: TfrmInvE01 );
    procedure ActiveInv013(const aGroupId: String; var aCompetence: String);
    procedure activeInv011;
    //procedure SetInv005Activate(const aTaxCode: String);
    procedure closeInv099;
    function getCheckNo(sI_InvID, sI_InvDate, sI_SystemID:String):String;
    procedure ActiveInv099(const aInvDate: TDateTime);
    procedure ActiveInv019(const aCustID: String);
    function GetInvAmtData(aInvID:String; var aInvAmt, aSaleAmt,
      aTaxAmt:Integer; var aTaxType, aInvFormat, aInvDate, aCustId, aBusinessId, aCustName: String): Boolean;
    {}
    function IsDataLocked(const aYearMonth, aCompId: String): Boolean;
    {}
    function getInv007Data(const aWhere: String):Integer; overload;
    function getInv007Data(const aWhere: String;blnUseAllComp:Boolean) : Integer; overload;
    function getInv035Data(const aWhere: String):Integer;
    function getInv014Data( aQueryType: Integer; aCompId, aPaperStartDate,
       aPaperStopDate, aYearMonth, aInvIdSt, aInvIdEd, aPaperSt, aPaperEd: String): Integer;
    function runSQL(const aMode: Integer; const aSQL: String): TADOQuery; overload;
    function runSQL(const aMode: Integer; const aSQL: String; aQuery: TADOQuery): TADOQuery; overload;
    function getYM(dI_InvDate:TDateTime):string;
    function IsDouble(const aPaperNo:string;aCompId: string;aAllowanceNo:string):Boolean;
  end;

var
  dtmMainH: TdtmMainH;

implementation

uses cbUtilis, dtmMainU, Ustru, UDateTimeU, dtmMainJU, XmlU;


{$R *.dfm}

{ TdtmMainH }


{ ---------------------------------------------------------------------------- }

function TdtmMainH.getInv007Data(const aWhere: String): Integer;
begin
  adoInv007.Close;
  adoInv007.SQL.Text := Format(
    '  SELECT                                                            ' +
    '         A.INVID,                                                   ' +
    '         A.INVDATE,                                                 ' +
    '         A.CUSTID,                                                  ' +
    '         A.CUSTSNAME,                                               ' +
    '         A.CHARGEDATE,                                              ' +
    '         A.INVTITLE,                                                ' +
    '         A.ZIPCODE,                                                 ' +
    '         A.INVADDR,                                                 ' +
    '         A.MAILADDR,                                                ' +
    '         A.BUSINESSID,                                              ' +
    '         DECODE( A.INVFORMAT, ''1'', ''電子'',                      ' +
    '                            ''2'', ''手二'',                        ' +
    '                            ''3'', ''手三'',                        ' +
    '                            A.INVFORMAT ) AS INVFORMAT,             ' +
    '         DECODE( A.TAXTYPE, ''1'', ''應稅'',                        ' +
    '                          ''2'', ''零稅率'',                        ' +
    '                          ''3'', ''免稅'',                          ' +
    '                          A.TAXTYPE ) AS TAXTYPE,                   ' +
    '         A.TAXRATE,                                                 ' +
    '         A.SALEAMOUNT,                                              ' +
    '         A.TAXAMOUNT,                                               ' +
    '         A.INVAMOUNT,                                               ' +
    '         DECODE( A.ISOBSOLETE, ''Y'',                               ' +
    '                               ''是'', ''否'' ) AS ISOBSOLETE,      ' +
    '         A.OBSOLETEREASON,                                          ' +
    '         A.MEMO1,                                                   ' +
    '         A.MEMO2,                                                   ' +
    '         A.UPTTIME,                                                 ' +
    '         A.UPTEN,                                                   ' +
    '         A.MAININVID,                                               ' +
    '         A.PRIZETYPE                                                ' +
    '    FROM INV007 A                                                   ' +
    '   WHERE IDENTIFYID1 = ''%s''                                       ' +
    '     AND IDENTIFYID2 = %s                                           ' +
    '     AND COMPID = ''%s''                                            ' +
    '     %s                                                             ' +
    '   ORDER BY INVID, INVDATE                                          ',
    [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, aWhere] );
  adoInv007.Open;
  Result := adoInv007.RecordCount;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMainH.runSQL(const aMode: Integer; const aSQL: String): TADOQuery;
begin
  Result := Self.adoComm;
  Result.Close;
  Result.SQL.Text := aSQL;
  case aMode of
    SELECT_MODE: Result.Open;
       IUD_MODE: Result.ExecSQL;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMainH.runSQL(const aMode: Integer; const aSQL: String;
  aQuery: TADOQuery): TADOQuery;
begin
  Result := aQuery;
  if not Assigned( Result ) then
    Result := adoComm;
  Result.Close;  
  Result.SQL.Text := aSQL;
  case aMode of
    SELECT_MODE: Result.Open;
       IUD_MODE: Result.ExecSQL;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMainH.setInvFormat;
    procedure addData(sI_CodeNo, sI_Desc:String);
    begin
      with cdsInvFormat do
      begin

        Append;
        FieldByName('CODENO').AsInteger := StrToInt(sI_CodeNo);
        FieldByName('DESCRIPTION').AsString := sI_Desc;
        post;
      end;
    end;

begin
    if not cdsInvFormat.Active then
      cdsInvFormat.CreateDataSet;
    cdsInvFormat.EmptyDataSet;
    addData('1','電子');
    addData('2','手二');
    addData('3','手三');
end;

procedure TdtmMainH.DataModuleCreate(Sender: TObject);
begin
    setInvFormat;
    setTaxType;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMainH.setTaxType;
    procedure addData(sI_CodeNo, sI_Desc:String);
    begin
      with cdsTaxType do
      begin

        Append;
        FieldByName('CODENO').AsInteger := StrToInt(sI_CodeNo);
        FieldByName('DESCRIPTION').AsString := sI_Desc;
        post;
      end;
    end;

begin
    if not cdsTaxType.Active then
      cdsTaxType.CreateDataSet;
    cdsTaxType.EmptyDataSet;
    addData('1','應稅');
    addData('2','零稅率');
    addData('3','免稅');
end;

{ ---------------------------------------------------------------------------- }

function TdtmMainH.getInv014Data( aQueryType: Integer; aCompId, aPaperStartDate,
  aPaperStopDate, aYearMonth, aInvIdSt, aInvIdEd, aPaperSt, aPaperEd: String): Integer;
var
  aWhere, aDaySt, aDayEd, aSql: String;
begin
  Result := 0;
  case aQueryType of
    { 單據日期 }
    0: begin
         aDaySt := TUstr.getOracleSQLDateStr( StrToDate( aPaperStartDate ) );
         aDayEd := TUstr.getOracleSQLDateStr( StrToDate( aPaperStopDate ) );
         aWhere := Format( ' AND PAPERDATE BETWEEN %s AND %s ', [aDaySt, aDayEd] );
       end;
    { 申報年月 }
    1: begin
         aWhere := Format( ' AND YEARMONTH = ''%s'' ', [aYearMonth] );
       end;
    { 發票號碼 }
    2: begin
         aWhere := Format( ' AND INVID BETWEEN  ''%s'' AND ''%s'' ',
           [aInvIdSt, aInvIdEd] );
       end;
    { 折讓單號 }   
    3: begin
         aWhere := Format( ' AND PAPERNO BETWEEN  ''%s'' AND ''%s'' ',
           [aPaperSt, aPaperEd] );
       end;
  end;
  aSql := Format(
    ' SELECT * FROM INV014         ' +
    '  WHERE IDENTIFYID1 = ''%s''  ' +
    '    AND IDENTIFYID2 = ''%s''  ' +
    '    AND COMPID LIKE ''%s''    ', [IDENTIFYID1, IDENTIFYID2, Nvl( aCompId, '%' )] ); 
  adoA07.Close;
  adoA07.SQL.Text := aSql + aWhere;
  try
    adoA07.Open;
    Result := adoA07.RecordCount;
  except
    { ... }
  end;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMainH.IsDataLocked(const aYearMonth, aCompId: String): Boolean;
begin
  Result := False;
  adoComm.Close;
  adoComm.SQL.Text := Format(
    ' SELECT ISLOCKED FROM INV018 WHERE IDENTIFYID1 = ''%s''   ' +
    '   AND IDENTIFYID2 = ''%s'' AND COMPID = ''%s''           ' +
    '   AND YEARMONTH = ''%s''                                 ',
    [IDENTIFYID1, IDENTIFYID2, aCompId, aYearMonth] );
  adoComm.Open;
  Result :=
    ( adoComm.RecordCount > 0 ) and
    ( adoComm.FieldByName( 'ISLOCKED' ).AsString = 'Y' );
  adoComm.Close;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMainH.GetInvAmtData(aInvID:String; var aInvAmt, aSaleAmt,
   aTaxAmt: Integer; var aTaxType, aInvFormat, aInvDate, aCustId, aBusinessId, aCustName: String): Boolean;
var
   aSQL : String;
begin
  aInvAmt := 0;
  aSaleAmt := 0;
  aTaxAmt := 0;
  aTaxType := '1';
  aInvDate := '';
  aSQL :=
    '  SELECT INVAMOUNT,       ' +
    '         SALEAMOUNT,      ' +
    '         TAXAMOUNT,       ' +
    '         INVFORMAT,       ' +
    '         TAXTYPE,         ' +
    '         CUSTID,          ' +
    '         BUSINESSID,      ' +
    '         TO_CHAR( INVDATE,''YYYY/MM/DD'' ) AS INVDATE, ' +
    '         CUSTSNAME        ' +
    '    FROM INV007           ' +
    '   WHERE COMPID = ' + QuotedStr( dtmMain.getCompID ) +
    '     AND INVID = ' + QuotedStr( aInvID );
  adoComm.SQL.Text := aSQL;
  adoComm.Open;
  Result := ( adoComm.RecordCount > 0 );
  if Result then
  begin
    aInvAmt := adoComm.FieldByName('INVAMOUNT').AsInteger;
    aSaleAmt := adoComm.FieldByName('SALEAMOUNT').AsInteger;
    aTaxAmt := adoComm.FieldByName('TAXAMOUNT').AsInteger;
    aTaxType := adoComm.FieldByName('TAXTYPE').AsString;
    aInvFormat := adoComm.FieldByName('INVFORMAT').AsString;
    aInvDate := adoComm.FieldByName('INVDATE').AsString;
    aCustId := adoComm.FieldByName('CUSTID').AsString;
    aBusinessId := adoComm.FieldByName('BUSINESSID').AsString;
    aCustName := adoComm.FieldByName('CUSTSNAME').AsString;
  end;
  adoComm.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMainH.adoA07CalcFields(DataSet: TDataSet);
var
    sL_InvFormat, sL_TaxType : String;
begin
    sL_InvFormat := DataSet.FieldByName('INVFORMAT').AsString ;
    sL_TaxType := DataSet.FieldByName('TAXTYPE').AsString ;

    if sL_InvFormat<>'' then
    begin
      with cdsInvFormat do
      begin
        if locate('CODENO', VarArrayOf([sL_InvFormat]),[loPartialKey]) then
          DataSet.FieldByName('INVFORMATDESC').AsString  := cdsInvFormat.fieldByName('DESCRIPTION').AsString;
      end;
    end;


    if sL_TaxType<>'' then
    begin
      with cdsTaxType do
      begin
        if locate('CODENO', VarArrayOf([sL_TaxType]),[loPartialKey]) then
          DataSet.FieldByName('TAXTYPEDESC').AsString  := cdsTaxType.fieldByName('DESCRIPTION').AsString;
      end;
    end;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMainH.activeInv019(const aCustId: String);
begin
  adoInv019.Close;
  adoInv019.SQL.Text :=
    ' SELECT TITLESNAME,                ' +
    '        TITLENAME INVTITLE,        ' +
    '        BUSINESSID,                ' +
    '        MZIPCODE ZIPCODE,          ' +
    '        MAILADDR MAILADDRESS,      ' +
    '        INVADDR INVADDRESS,        ' +
    '        MEMO                       ' +
    '   FROM INV019 WHERE CUSTID = ' + QuotedStr( aCustId );
  adoInv019.Open;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMainH.ActiveInv099(const aInvDate: TDateTime);
begin
  adoInv099.Close;
  adoInv099.SQL.Text :=
    ' SELECT YEARMONTH,      ' +
    '        INVFORMAT,      ' +
    '        PREFIX,         ' +
    '        STARTNUM,       ' +
    '        ENDNUM,         ' +
    '        CURNUM,         ' +
    '        TO_CHAR( LASTINVDATE, ''YYYY/MM/DD'' ) LINVDATE, ' +
    '        MEMO                               ' +
    '   FROM INV099                             '+
    '  WHERE IDENTIFYID1 = ' + QuotedStr( IDENTIFYID1 ) +
    '    AND IDENTIFYID2 = ' + IDENTIFYID2 +
    '    AND COMPID = ' + QuotedStr( dtmMain.getCompID ) +
    '    AND YEARMONTH = ' + QuotedStr( getYM( aInvDate ) ) +
    '    AND USEFUL = ''Y'' ';
   adoInv099.Open;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMainH.adoInv099CalcFields(DataSet: TDataSet);
begin
  case DataSet.FieldByName('InvFormat').AsInteger of
    1: DataSet.FieldByName('InvFormatDesc').AsString := '電子';
    2: DataSet.FieldByName('InvFormatDesc').AsString := '手二';
    3: DataSet.FieldByName('InvFormatDesc').AsString := '手三';
  end;
end;

{ ---------------------------------------------------------------------------- }

function TdtmMainH.getCheckNo(sI_InvID, sI_InvDate, sI_SystemID: String): String;
var
    L_Tmp1, L_Tmp2 : array[0..7] of Integer;
    L_Tmp3 : array[0..8] of Integer;
    nL_CheckNo, nL_Length,i,j : Integer;
    sL_TargetInvID, sL_TargetInvDate : String;
    sL_Tmp, sL_Tmp1, sL_Tmp2, sL_Result : String;
begin

    sL_TargetInvID  := Copy(sI_InvID,3,10);


    sL_TargetInvDate  := sI_SystemID + TUdateTime.CDateStr(StrToDate(sI_InvDate),6);
    nL_Length := 8;
//    i:=0;
//    j:=0;


    for i:=0 to nL_Length-1 do
    begin

      L_Tmp1[i] := StrToInt(Copy(sL_TargetInvID,i+1,1));
      L_Tmp2[i] := StrToInt(Copy(sL_TargetInvDate,i+1,1));
      L_Tmp3[i] := 0;
    end;
    L_Tmp3[nL_Length] := 0;

    sL_Tmp := '';
    sL_Tmp1 := '';
    sL_Tmp2 := '';



    j :=nL_Length ;
    for i:=0 to nL_Length-1 do
    begin
      sL_Tmp := IntToStr(L_Tmp1[j-1]*L_Tmp2[j-1]);
      if length(sL_Tmp)=2 then
      begin
        sL_Tmp1 := Copy(sL_Tmp,1,1);
        sL_Tmp2 := Copy(sL_Tmp,2,1);
      end
      else
      begin
        sL_Tmp1 := '0';
        sL_Tmp2 := Copy(sL_Tmp,1,1);
      end;
      L_Tmp3[i] := L_Tmp3[i] + StrToInt(sL_Tmp1);
      L_Tmp3[i+1] := L_Tmp3[i+1] + StrToInt(sL_Tmp2);

      J := J-1;
    end;
    

    nL_CheckNo :=0;

    sL_Result := '';
    for i:=0 to  length(L_Tmp3)-1 do
    begin
      nL_CheckNo := nL_CheckNo + L_Tmp3[i];
    end;
    result := IntToStr(nL_CheckNo);

end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMainH.closeInv099;
begin
  adoInv099.Close;
end;

{ ---------------------------------------------------------------------------- }

//procedure TdtmMainH.SetInv005Activate(const aTaxCode:String);
//begin
//  adoInv005.Close;
//  if ( aTaxCode <> '' ) then
//  begin
//    adoInv005.SQL.Text :=
//      '  SELECT ITEMID,         ' +
//      '         DESCRIPTION,    ' +
//      '         TAXNAME,        ' +
//      '         SIGN            ' +
//      '    FROM INV005          ' +
//      '   WHERE TAXCODE =       ' + QuotedStr( aTaxCode );
//    adoInv005.Open;
//  end;
//end;

{ ---------------------------------------------------------------------------- }

function TdtmMainH.getYM(dI_InvDate: TDateTime): string;
var
    sL_Year,  sL_Month : String;
    sL_TmpDateStr : String;
begin

    sL_TmpDateStr := TUdateTime.GetPureDateStr(dI_InvDate,'');
    sL_Year := Copy(sL_TmpDateStr,1,4);
    sL_Month := Copy(sL_TmpDateStr,5,2);


    if (StrToInt(sL_Month) mod 2)=0 then
    begin

      sL_Month := Format('%.2d',[StrToInt(sL_Month)-1]);
    end;

    sL_TmpDateStr := sL_Year + sL_Month;

    result := sL_TmpDateStr;
end;

procedure TdtmMainH.activeInv011;
begin
    with adoInv011 do
    begin

      if State=dsInactive then
      begin
        SQL.Clear;
        SQL.Add('select ITEMID, DESCRIPTION from INV011 order By ITEMID');
        Open;
      end;
    end;
end;

procedure TdtmMainH.initCompetenceItem(sI_GroupID:String; I_Frm: TfrmInvE01);
var
    ii, jj : Integer;
    sL_TmpID, sL_TmpQueryCompetence : String;
    L_MsgNodeList : IDOMNodeList;
    L_TmpXmlDoc : TXMLDocument;
    sL_TmpDesc:String;
    sL_CompetenceID, sL_CompetenceDesc : String;
    sL_CompetenceXmlStr:String;
begin
    activeInv011;

    with adoInv011 do
    begin
      First;
      while not Eof do
      begin
        sL_CompetenceID := FieldByName('ITEMID').asString;
        sL_CompetenceDesc := FieldByName('DESCRIPTION').asString;

        I_Frm.addListviewItem(sL_CompetenceID, sL_CompetenceDesc, I_Frm.livNonCompetenceItem,1);
        Next;
      end;

    end;


    activeInv013(sI_GroupID, sL_CompetenceXmlStr);
    if sL_CompetenceXmlStr <> '' then
    begin
      //sL_XmlFilePath := IncludeTrailingPathDelimiter( ExtractFilePath(
      //  Application.ExeName ) ) + 'TmpCompetence.xml';
      //dtmMain.saveToFile(sL_CompetenceXmlStr,sL_XmlFilePath);

      L_TmpXmlDoc := TXMLDocument.Create( nil );
      try
        //L_TmpXmlDoc.LoadFromFile(sL_XmlFilePath);
        L_TmpXmlDoc.XML.Text := sL_CompetenceXmlStr;
        L_TmpXmlDoc.Active := True;
        L_MsgNodeList := L_TmpXmlDoc.DOMDocument.getElementsByTagName('GROUP');

        for ii:=0 to L_MsgNodeList.length-1 do
        begin
          for jj:=0 to L_MsgNodeList.item[ii].childNodes.length-1 do
          begin
            sL_TmpID := TMyXML.getAttributeValue(L_MsgNodeList.item[ii].childNodes.item[jj],'ID');
            sL_TmpQueryCompetence := TMyXML.getAttributeValue(L_MsgNodeList.item[ii].childNodes.item[jj],'Q');
            sL_TmpDesc := '';
            if adoInv011.Locate('ITEMID', sL_TmpID,[loPartialKey]) then
              sL_TmpDesc := adoInv011.FieldByName('DESCRIPTION').AsString;
            if sL_TmpQueryCompetence='Y' then
            begin
              I_Frm.addListviewItem(sL_TmpID, sL_TmpDesc, I_Frm.livHasCompetenceItem,0);
              I_Frm.removeListviewItem(sL_TmpID, I_Frm.livNonCompetenceItem);
            end;
          end;
        end;
        L_TmpXmlDoc.Active := False;
      finally
        L_TmpXmlDoc.Free;
      end;  
    end;
    adoInv011.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TdtmMainH.ActiveInv013(const aGroupId: String; var aCompetence: String);
begin
  aCompetence := '';
  adoInv013.Close;
  adoInv013.SQL.Text :=
    '  SELECT COMPETENCE FROM INV013 WHERE GROUPID = ' + QuotedStr( aGroupId );
  adoInv013.Open;
  if ( not adoInv013.IsEmpty ) then
    aCompetence := adoInv013.FieldByName( 'COMPETENCE' ).AsString;
  adoInv013.Close;
end;

{ ---------------------------------------------------------------------------- }

{
1.)    輸入首筆明細 & 發票資料維護 - -  發票明細 的  單價 欄位可輸入 小數點2位

2.)    銷售額、稅額及總金額 於計算時，需四捨五入計算至整數位數

3.)    當使用者key入　單價及數量時，系統可自動計算 銷售額、稅額及總金額

    =>   數量*單價= 銷售額
         銷售額 * 稅率 = 稅額
         稅額 + 銷售額 = 總金額

         1*12.5 = 13(銷售額)
         13 * 0.05 = 1(稅額)　
         總金額 = 14

4.)    僅key 入 銷售額離開該欄位時，系統應可自動換算出稅額及總金額，（可討論是否可行）  可行

5.)    反之使用者只給予　總金額離開該欄位時，應可反算出銷售額及稅額（可討論是否可行）   可行
          => 總金額/1.05 =  銷售額
            稅額 = 總金額 - 銷售額

6.)    數量沒有建入時，請帶入預設值為１

7.)    此三個欄位，總金額、
銷售額及稅額　即使由系統自動算出，使用者仍可自行修改此三個欄位


}

{ ---------------------------------------------------------------------------- }


{------------------------------------------------------------------------------}
{#5302 檢查單據編號是否有重複 By Kin 2009/09/30}
function TdtmMainH.IsDouble(const aPaperNo:string;aCompId:string;aAllowanceNo: string): Boolean;
begin
  Result := False;
  if aPaperNo <> EmptyStr then
  begin
    if aAllowanceNo = EmptyStr then
      aAllowanceNo := ' ';
    adoComm.Close;
    adoComm.SQL.Text := Format(
      'SELECT COUNT(PAPERNO) As CNT FROM INV014 WHERE PAPERNO = ''%s'' ' +
      ' AND COMPID = ''%s'' AND ALLOWANCENO <> ''%s''                 ' +
      ' AND IDENTIFYID1 = ''%s'' AND IDENTIFYID2 = ''%s''              ' ,
      [aPaperNo,aCompid,aAllowanceNo,IDENTIFYID1,IDENTIFYID2 ] );
    adoComm.Open;
    Result := (adoComm.FieldByName( 'CNT' ).AsInteger > 0);
    adoComm.Close;
  end;
end;

function TdtmMainH.getInv035Data(const aWhere: String): Integer;
begin
  adoEinv.Close;
  adoEinv.SQL.Text := 'SELECT * FROM INV035 ';
  if aWhere <> EmptyStr then
  begin
    adoEinv.SQL.Text := adoEinv.SQL.Text + ' WHERE ' + aWhere;
  end;
  adoEinv.Open;
  Result := adoEinv.RecordCount;
end;

function TdtmMainH.getInv007Data(const aWhere: String;
  blnUseAllComp: Boolean): Integer;
var aSQL : String;
begin
  adoInv007.Close;
  aSQL := Format( ' SELECT * FROM INV007                  ' +
                ' WHERE IDENTIFYID1 = ''%s''              ' +
                ' AND IDENTIFYID2 = %s                    ' ,
                [IDENTIFYID1,IDENTIFYID2] );
  if ( aWhere <> EmptyStr ) then
  begin
    aSQL := Format( aSQL + '  %s ',[aWhere]);
  end;
  if ( not blnUseAllComp ) then
  begin
    aSQL := Format( aSQL + ' AND COMPID = ''%s''',[dtmMain.getCompID]);
  end;
  aSQL := aSQL + ' ORDER BY INVID,INVDATE';
  adoInv007.Close;
  adoInv007.SQL.Text := aSQL;
  adoInv007.Open;
  Result := adoInv007.RecordCount;
end;

end.


