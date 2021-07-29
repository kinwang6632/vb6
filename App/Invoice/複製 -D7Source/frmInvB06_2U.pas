unit frmInvB06_2U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Buttons,dtmReportModule,DB;

type
  TfrmInvB06_2 = class(TForm)
    Panel1: TPanel;
    btnOk: TBitBtn;
    rdgWhere: TRadioGroup;
    btnExit: TBitBtn;
    procedure btnOkClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
  private
    { Private declarations }
    FYearMonth:String;
    procedure UnPrepareDataSet;
    procedure PrepareDataSet;
    procedure PrintPrize;
    procedure PrintMedia;
    procedure TranMediaTxt(aBusinessId,aCity,aTaxId,aPath,aWhere
        ,aManager,aCompAddr,aBranchNo : String);


  public
    { Public declarations }
    property YearMonth: String read FYearMonth write FYearMonth;
  end;

var
  frmInvB06_2: TfrmInvB06_2;

implementation
uses cbUtilis,  dtmMainU, dtmMainJU, dtmMainHU;
{$R *.dfm}

{ TfrmInvB06_2 }

procedure TfrmInvB06_2.PrepareDataSet;
begin

  dtmReport.cdsTempory.FieldDefs.Clear;
  dtmReport.cdsTempory.FieldDefs.Add( 'YEAR', ftString, 4 );
  dtmReport.cdsTempory.FieldDefs.Add( 'MONTH', ftString, 2 );
  dtmReport.cdsTempory.FieldDefs.Add( 'YEARMONTH', ftString, 6 );
  dtmReport.cdsTempory.FieldDefs.Add( 'STARTMONTH',ftString,2 );
  dtmReport.cdsTempory.FieldDefs.Add( 'ENDMONTH',ftString,2 );
  dtmReport.cdsTempory.FieldDefs.Add( 'EXECUTEDATE', ftDateTime );
  dtmReport.cdsTempory.FieldDefs.Add( 'UPTTIME', ftDateTime );
  dtmReport.cdsTempory.FieldDefs.Add( 'UPTEN', ftString, 20 );
  dtmReport.cdsTempory.FieldDefs.Add( 'SPECIALPRIZE1', ftString, 8 );
  dtmReport.cdsTempory.FieldDefs.Add( 'SPECIALPRIZE2', ftString, 8 );
  dtmReport.cdsTempory.FieldDefs.Add( 'SPECIALPRIZE3', ftString, 8 );
  dtmReport.cdsTempory.FieldDefs.Add( 'FIRSTPRIZE1', ftString, 8 );
  dtmReport.cdsTempory.FieldDefs.Add( 'FIRSTPRIZE2', ftString, 8 );
  dtmReport.cdsTempory.FieldDefs.Add( 'FIRSTPRIZE3', ftString, 8 );
  dtmReport.cdsTempory.FieldDefs.Add( 'EXTRAPRIZE1', ftString, 8 );
  dtmReport.cdsTempory.FieldDefs.Add( 'EXTRAPRIZE2', ftString, 8 );
  dtmReport.cdsTempory.FieldDefs.Add( 'COMPID',ftString,3);
  dtmReport.cdsTempory.FieldDefs.Add( 'COMPNAME',ftString,50);
  dtmReport.cdsTempory.CreateDataSet;
end;

procedure TfrmInvB06_2.UnPrepareDataSet;
begin
  if not VarIsNull( dtmReport.cdsTempory.Data ) then
    dtmReport.cdsTempory.Data := Null;
  dtmReport.cdsTempory.FieldDefs.Clear;

end;

procedure TfrmInvB06_2.btnOkClick(Sender: TObject);
begin
  if rdgWhere.ItemIndex = 0 then
  begin
    PrintPrize;
  end else
  begin
    PrintMedia;
  end;
end;

procedure TfrmInvB06_2.PrintPrize;
var
  aSQL, aPath,aEndMonth : String;
begin
  aSQL := Format('SELECT SUBSTR(A.YEARMONTH,1,4) YEAR,      ' +
            'SUBSTR(A.YEARMONTH,5,2) MONTH,A.*              ' +
            'FROM INV035 A WHERE YEARMONTH = ''%s''         ',
            [FYearMonth] );
  Screen.Cursor := crSQLWait;
  try
    UnPrepareDataSet;
    PrepareDataSet;
    dtmMain.adoEComm.Close;
    dtmMain.adoEComm.SQL.Text := aSQL;
    dtmMain.adoEComm.Open;
    while not dtmMain.adoEComm.Eof do
    begin
      dtmReport.cdsTempory.Append;
      dtmReport.cdsTempory.FieldByName( 'YEAR' ).AsString :=
        dtmMain.adoEComm.FieldByName( 'YEAR' ).AsString;
      dtmReport.cdsTempory.FieldByName( 'MONTH' ).AsString :=
        DTMMAIN.adoEComm.FieldByName('MONTH').AsString;
      dtmReport.cdsTempory.FieldByName( 'YEARMONTH' ).AsString :=
        dtmMain.adoEComm.FieldByName( 'YEARMONTH' ).AsString;
      dtmReport.cdsTempory.FieldByName( 'EXECUTEDATE' ).AsString :=
        dtmMain.adoEComm.FieldByName( 'EXECUTEDATE' ).AsString;
      dtmReport.cdsTempory.FieldByName( 'UPTTIME' ).AsString :=
        dtmMain.adoEComm.FieldByName( 'UPTTIME' ).AsString;
      dtmReport.cdsTempory.FieldByName( 'UPTEN' ).AsString :=
        dtmMain.adoEComm.FieldByName( 'UPTEN' ).AsString;
      dtmReport.cdsTempory.FieldByName( 'SPECIALPRIZE1' ).AsString :=
        dtmMain.adoEComm.FieldByName( 'SPECIALPRIZE1' ).AsString;
      dtmReport.cdsTempory.FieldByName( 'SPECIALPRIZE2' ).AsString :=
        dtmMain.adoEComm.FieldByName( 'SPECIALPRIZE2' ).AsString;
      dtmReport.cdsTempory.FieldByName( 'SPECIALPRIZE3' ).AsString :=
        dtmMain.adoEComm.FieldByName( 'SPECIALPRIZE3' ).AsString;
      dtmReport.cdsTempory.FieldByName( 'FIRSTPRIZE1' ).AsString :=
        dtmMain.adoEComm.FieldByName( 'FIRSTPRIZE1' ).AsString;
      dtmReport.cdsTempory.FieldByName( 'FIRSTPRIZE2' ).AsString :=
        dtmMain.adoEComm.FieldByName( 'FIRSTPRIZE2' ).AsString;
      dtmReport.cdsTempory.FieldByName( 'FIRSTPRIZE3' ).AsString :=
        dtmMain.adoEComm.FieldByName( 'FIRSTPRIZE3' ).AsString;
      dtmReport.cdsTempory.FieldByName( 'EXTRAPRIZE1' ).AsString :=
        dtmMain.adoEComm.FieldByName( 'EXTRAPRIZE1' ).AsString;
      dtmReport.cdsTempory.FieldByName( 'EXTRAPRIZE2' ).AsString :=
        dtmMain.adoEComm.FieldByName( 'EXTRAPRIZE2' ).AsString;
      dtmReport.cdsTempory.FieldByName( 'STARTMONTH' ).AsString :=
        copy(FYearMonth,5,2);
      aEndMonth :=  IntToStr(StrtoInt( copy(FYearMonth,5,2) ) + 1 );
      dtmReport.cdsTempory.FieldByName( 'COMPID' ).AsString :=
        dtmMain.getCompID;
      dtmReport.cdsTempory.FieldByName( 'COMPNAME' ).AsString :=
        dtmMain.getCompName;
      if Length(aEndMonth) = 1 then
        aEndMonth := '0'+aEndMonth;
      dtmReport.cdsTempory.FieldByName( 'ENDMONTH' ).AsString := aEndMonth;
      dtmReport.cdsTempory.Post;
      dtmMain.adoEComm.Next;
    end;

    aPath := IncludeTrailingPathDelimiter( ExtractFilePath(
      Application.ExeName ) ) + IncludeTrailingPathDelimiter(
      REPORT_FOLDER ) + 'FrptInvB06_1.fr3';
    dtmReport.frxMainReport.LoadFromFile( aPath );
    dtmReport.frxMainReport.Variables.Variables['中獎號碼單抬頭'] :=
      QuotedStr( '中獎號碼單');
    dtmReport.frxMainReport.ShowReport;

  finally
    dtmMain.adoEComm.Close;
    Screen.Cursor := crDefault;
  end;

end;

procedure TfrmInvB06_2.PrintMedia;
var
  aBusinessId,aCity,aSQL,aTaxId : String;
  aPath,sL_SQL,aWhere,aManager,aCompAddr,aBranchNo : String;
begin
  Screen.Cursor := crSQLWait;
  try

    {INV001}
    aSQL := Format('SELECT * FROM INV001 WHERE COMPID = ''%s''  ',
            [dtmMain.getCompID]);
    dtmMain.adoComm.Close;
    dtmMain.adoComm.SQL.Text := aSQL;
    dtmMain.adoComm.Open;
    aBusinessId := dtmMain.adoComm.FieldByName('BUSINESSID').AsString;
    aTaxId := dtmMain.adoComm.FieldByName( 'TAXID' ).AsString;
    //原本是抓Manger欄位,現在改抓COMPNAME (沒有題號,Jacy發Mail通知) By Kin 2010/10/08
    aManager := dtmMain.adoComm.FieldByName( 'COMPNAME' ).AsString;
    aCompAddr := dtmMain.adoComm.FieldByName( 'COMPADDR' ).AsString;

    {INV003}
    dtmMain.adoComm.Close;
    aSQL := Format('SELECT * FROM INV003 WHERE COMPID = ''%s''  ',
            [dtmMain.getCompID]);
    dtmMain.adoComm.SQL.Text := aSQL;
    dtmMain.adoComm.Open;
    aCity := dtmMain.adoComm.FieldByName( 'CITY' ).AsString;
    aBranchNo := dtmMain.adoComm.FieldByName( 'BRANCHNO' ).AsString;

    dtmMain.adoComm.Close;

    {取出檔案路徑}
    sL_SQL := 'SELECT INVPARAM FROM INV003 WHERE COMPID = ' +
      QuotedStr( dtmMain.getCompID );
    aPath := dtmMainJ.GetXMLAttribute( sL_SQL, 'MediaFilePath' );
    aPath := IncludeTrailingPathDelimiter( aPath );
    { Where 條件}
    aWhere := Format(' AND PRIZETYPE IS NOT NULL                 ' +
          ' AND INVDATE >= TO_DATE(''%s'',''yyyymmdd'')          ' +
          ' AND INVDATE <= TO_DATE(''%s'',''yyyymmdd'')          ' ,
          [dtmMainJ.GetFulDate(0,FYearMonth),
          dtmMainJ.GetFulDate(1,FYearMonth)]);

    {產生檔案}
    TranMediaTxt( aBusinessId,aCity,aTaxId,aPath,aWhere,aManager,aCompAddr,aBranchNo );

  finally
    Screen.Cursor := crDefault;
  end;

end;

procedure TfrmInvB06_2.TranMediaTxt(aBusinessId, aCity, aTaxId, aPath,
  aWhere,aManager,aCompAddr,aBranchNo: String);
var
  L_TxtList1,L_TxtList2: TStringList ;
  aTmp : String;
begin
  L_TxtList1 := TStringList.Create;
  L_TxtList2 := TStringList.Create;
  try
    if dtmMainH.getInv007Data( aWhere ) = 0 then
    begin
      WarningMsg( '無任何資料可產生！' );
      Exit;
    end;
    dtmMainH.adoInv007.First;
    while not dtmMainH.adoInv007.Eof do
    begin
      aTmp := EmptyStr;

      aTmp :=  IntToStr(
                StrToInt( copy(dtmMainH.adoInv007.FieldByName( 'INVDATE' ).AsString,1,4) ) - 1911);
      if Length( aTmp ) = 2 then
        aTmp := '0' + aTmp;
      aTmp := aTmp + copy(dtmMainH.adoInv007.FieldByName( 'INVDATE' ).AsString,6,2);
      aTmp := aTmp + dtmMainH.adoInv007.FieldByName( 'INVID' ).AsString;
      aTmp := aTmp + adjString(8,aBusinessId,False,False);
      aTmp := aTmp + adjString(1,aCity,False,False);
      aTmp := aTmp + adjString(9,aTaxId,False,False);
      aTmp := aTmp + dtmMainH.adoInv007.FieldByName( 'PRIZETYPE' ).AsString;
      aTmp := aTmp + adjString(52,aManager,False,False);
      aTmp := aTmp + adjString( 80,aCompAddr,false,false);

      if dtmMainH.adoInv007.FieldByName( 'PRIZETYPE' ).AsInteger > 3 then
      begin
        aTmp := aTmp + adjString(8,aBranchNo,false,false);
        L_TxtList2.Add( aTmp );
      end else
      begin
        if ( dtmMainH.adoInv007.FieldByName( 'BUSINESSID' ).AsString = '' )
          and (dtmMainH.adoInv007.FieldByName( 'ISOBSOLETE' ).AsString <> '是') then
        begin
          aTmp := aTmp + '  ';
        end else
        begin
          if dtmMainH.adoInv007.FieldByName( 'BUSINESSID' ).AsString <> '' then
          begin
            aTmp := aTmp + '01';
          end else
          begin
            aTmp := aTmp + '03';
          end;
        end;
        aTmp := aTmp + adjString(8,aBranchNo,false,false);
        L_TxtList1.Add( aTmp );
      end;
      dtmMainH.adoInv007.Next;
    end;
    dtmMainH.adoInv007.Close;
    
    if L_TxtList1.Count > 0 then
      L_TxtList1.SaveToFile( aPath + aBusinessId + '.IMF' );
    if L_TxtList2.Count > 0 then
      L_TxtList2.SaveToFile( aPath + aBusinessId + '.IFF' );
    WarningMsg( '產生完畢！' );
  finally

    L_TxtList1.Free;
    L_TxtList2.Free;
  end;


end;

procedure TfrmInvB06_2.btnExitClick(Sender: TObject);
begin
  Self.Close;
end;

end.
