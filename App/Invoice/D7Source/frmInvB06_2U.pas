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
    FUseCompWhere : Integer;
    procedure UnPrepareDataSet;
    procedure PrepareDataSet;
    procedure PrintPrize;
    procedure PrintMedia;
    procedure TranMediaTxt(aBusinessId,aCity,aTaxId,aPath,aWhere
        ,aManager,aCompAddr,aBranchNo : String;blnUseAllComp : Boolean);
    procedure GetInv001FieldData(var aBusinessId,aTaxId,aManager,aCompAddr : String; aCompId : String);
    procedure GetInv003FieldData(var aCity,aBranchNo : String;aCompId : String);

  public
    { Public declarations }
    property YearMonth: String read FYearMonth write FYearMonth;
    property CompWhere: Integer read FUseCompWhere write FUseCompWhere;
  end;

var
  frmInvB06_2: TfrmInvB06_2;

implementation
uses cbUtilis,  dtmMainU, dtmMainJU, dtmMainHU;
{$R *.dfm}

{ TfrmInvB06_2 }

procedure TfrmInvB06_2.PrepareDataSet;
begin
  //#5966 �W�[�S�O�����X By Kin 2011/03/31
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
  dtmReport.cdsTempory.FieldDefs.Add( 'SPECIALPRIZE4', ftString, 8 );
  dtmReport.cdsTempory.FieldDefs.Add( 'SPECIALPRIZE5', ftString, 8 );
  dtmReport.cdsTempory.FieldDefs.Add( 'FIRSTPRIZE1', ftString, 8 );
  dtmReport.cdsTempory.FieldDefs.Add( 'FIRSTPRIZE2', ftString, 8 );
  dtmReport.cdsTempory.FieldDefs.Add( 'FIRSTPRIZE3', ftString, 8 );
  dtmReport.cdsTempory.FieldDefs.Add( 'FIRSTPRIZE4', ftString, 8 );
  dtmReport.cdsTempory.FieldDefs.Add( 'FIRSTPRIZE5', ftString, 8 );
  dtmReport.cdsTempory.FieldDefs.Add( 'EXTRAPRIZE1', ftString, 8 );
  dtmReport.cdsTempory.FieldDefs.Add( 'EXTRAPRIZE2', ftString, 8 );
  dtmReport.cdsTempory.FieldDefs.Add( 'EXTRAPRIZE3', ftString, 8 );
  dtmReport.cdsTempory.FieldDefs.Add( 'EXTRAPRIZE4', ftString, 8 );
  dtmReport.cdsTempory.FieldDefs.Add( 'EXTRAPRIZE5', ftString, 8 );
  dtmReport.cdsTempory.FieldDefs.Add( 'UNUSUALPRIZE1',ftString,8 );
  dtmReport.cdsTempory.FieldDefs.Add( 'UNUSUALPRIZE2',ftString,8 );
  dtmReport.cdsTempory.FieldDefs.Add( 'UNUSUALPRIZE3',ftString,8 );
  dtmReport.cdsTempory.FieldDefs.Add( 'UNUSUALPRIZE4',ftString,8 );
  dtmReport.cdsTempory.FieldDefs.Add( 'UNUSUALPRIZE5',ftString,8 );
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
    //#5966 �W�[�S�O�����X By Kin 2011/03/31
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
      dtmReport.cdsTempory.FieldByName( 'SPECIALPRIZE4' ).AsString :=
        dtmMain.adoEComm.FieldByName( 'SPECIALPRIZE4' ).AsString;
      dtmReport.cdsTempory.FieldByName( 'SPECIALPRIZE5' ).AsString :=
        dtmMain.adoEComm.FieldByName( 'SPECIALPRIZE5' ).AsString;
      dtmReport.cdsTempory.FieldByName( 'FIRSTPRIZE1' ).AsString :=
        dtmMain.adoEComm.FieldByName( 'FIRSTPRIZE1' ).AsString;
      dtmReport.cdsTempory.FieldByName( 'FIRSTPRIZE2' ).AsString :=
        dtmMain.adoEComm.FieldByName( 'FIRSTPRIZE2' ).AsString;
      dtmReport.cdsTempory.FieldByName( 'FIRSTPRIZE3' ).AsString :=
        dtmMain.adoEComm.FieldByName( 'FIRSTPRIZE3' ).AsString;
      dtmReport.cdsTempory.FieldByName( 'FIRSTPRIZE4' ).AsString :=
        dtmMain.adoEComm.FieldByName( 'FIRSTPRIZE4' ).AsString;
      dtmReport.cdsTempory.FieldByName( 'FIRSTPRIZE5' ).AsString :=
        dtmMain.adoEComm.FieldByName( 'FIRSTPRIZE5' ).AsString;
      dtmReport.cdsTempory.FieldByName( 'EXTRAPRIZE1' ).AsString :=
        dtmMain.adoEComm.FieldByName( 'EXTRAPRIZE1' ).AsString;
      dtmReport.cdsTempory.FieldByName( 'EXTRAPRIZE2' ).AsString :=
        dtmMain.adoEComm.FieldByName( 'EXTRAPRIZE2' ).AsString;
      dtmReport.cdsTempory.FieldByName( 'EXTRAPRIZE3' ).AsString :=
        dtmMain.adoEComm.FieldByName( 'EXTRAPRIZE3' ).AsString;
      dtmReport.cdsTempory.FieldByName( 'EXTRAPRIZE4' ).AsString :=
        dtmMain.adoEComm.FieldByName( 'EXTRAPRIZE4' ).AsString;
      dtmReport.cdsTempory.FieldByName( 'EXTRAPRIZE5' ).AsString :=
        dtmMain.adoEComm.FieldByName( 'EXTRAPRIZE5' ).AsString;
      dtmReport.cdsTempory.FieldByName( 'UNUSUALPRIZE1' ).AsString :=
        dtmMain.adoEComm.FieldByName( 'UNUSUALPRIZE1').AsString;
      dtmReport.cdsTempory.FieldByName( 'UNUSUALPRIZE2' ).AsString :=
        dtmMain.adoEComm.FieldByName( 'UNUSUALPRIZE2').AsString;
      dtmReport.cdsTempory.FieldByName( 'UNUSUALPRIZE3' ).AsString :=
        dtmMain.adoEComm.FieldByName( 'UNUSUALPRIZE3').AsString;
      dtmReport.cdsTempory.FieldByName( 'UNUSUALPRIZE4' ).AsString :=
        dtmMain.adoEComm.FieldByName( 'UNUSUALPRIZE4').AsString;
      dtmReport.cdsTempory.FieldByName( 'UNUSUALPRIZE5' ).AsString :=
        dtmMain.adoEComm.FieldByName( 'UNUSUALPRIZE5').AsString;
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
    dtmReport.frxMainReport.Variables.Variables['�������X����Y'] :=
      QuotedStr( '�������X��');
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
  aUseAllComp : Boolean;
begin
  Screen.Cursor := crSQLWait;
  aUseAllComp := True;
  aBusinessId := EmptyStr;
  aCity := EmptyStr;
  aTaxId := EmptyStr;
  aManager := EmptyStr;
  aCompAddr := EmptyStr;
  aBranchNo := EmptyStr;

  try

    if FUseCompWhere = 0 then
    begin
      aUseAllComp := False;
      {INV001}
      aSQL := Format('SELECT * FROM INV001 WHERE COMPID = ''%s''  ',
              [dtmMain.getCompID]);
      dtmMain.adoComm.Close;
      dtmMain.adoComm.SQL.Text := aSQL;
      dtmMain.adoComm.Open;
      aBusinessId := dtmMain.adoComm.FieldByName('BUSINESSID').AsString;
      aTaxId := dtmMain.adoComm.FieldByName( 'TAXID' ).AsString;
      //�쥻�O��Manger���,�{�b���COMPNAME (�S���D��,Jacy�oMail�q��) By Kin 2010/10/08
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

    end;

    {���X�ɮ׸��|}
    sL_SQL := 'SELECT INVPARAM FROM INV003 WHERE COMPID = ' +
      QuotedStr( dtmMain.getCompID );
    aPath := dtmMainJ.GetXMLAttribute( sL_SQL, 'MediaFilePath' );
    aPath := IncludeTrailingPathDelimiter( aPath );
    { Where ����}
    { Jacy ���X 6 ���]���ͦ�IIF�ɡA�ҥH�n�L�o PRIZETYPE <> 6 By Kin 2011/03/17}
    {#6600 CARRIERID1 �������ƴN���C�L�o�� By Kin 2013/10/02 }
    {#6629 �S��CarrierId1 ���󮳱� By Kin 2013/10/23}
    aWhere := Format(' AND PRIZETYPE IS NOT NULL                 ' +
          ' AND PRIZETYPE <> ''6''                               ' +
          ' AND INVDATE >= TO_DATE(''%s'',''yyyymmdd'')          ' +
          ' AND INVDATE <= TO_DATE(''%s'',''yyyymmdd'')          ' ,
          [dtmMainJ.GetFulDate(0,FYearMonth),
          dtmMainJ.GetFulDate(1,FYearMonth)]);

    {�����ɮ�}
    TranMediaTxt( aBusinessId,aCity,aTaxId,aPath,aWhere,aManager,aCompAddr,aBranchNo,aUseAllComp );

  finally
    Screen.Cursor := crDefault;
  end;

end;

procedure TfrmInvB06_2.TranMediaTxt(aBusinessId, aCity, aTaxId, aPath,
  aWhere,aManager,aCompAddr,aBranchNo: String;blnUseAllComp : Boolean);
var
  L_TxtList1,L_TxtList2: TStringList ;
  aTmp,aCompId,aInvId : String;
begin
  L_TxtList1 := TStringList.Create;
  L_TxtList2 := TStringList.Create;
  try

    if dtmMainH.getInv007Data( aWhere,blnUseAllComp ) = 0 then
    begin
      WarningMsg( '�L�����ƥi���͡I' );
      Exit;
    end;
    dtmMainH.adoInv007.First;
    while not dtmMainH.adoInv007.Eof do
    begin
      //#5874 �p�G�O�Ҧ����q�A�h�n�@���@���M�� By Kin 2010/12/13
      if blnUseAllComp then
      begin

        GetInv001FieldData( aBusinessId,aTaxId,aManager,aCompAddr,
                  dtmMainH.adoinv007.FieldByName('COMPID').AsString );

        GetInv003FieldData( aCity,aBranchNo,
                  dtmMainH.adoInv007.FieldByName('COMPID').AsString);
        aCompId := dtmMainH.adoInv007.FieldByName('COMPID').AsString;
      end else
      begin
        aCompId := dtmMain.getCompID;
      end;
      aInvId := dtmMainH.adoInv007.FieldByName('INVID').AsString;
      aTmp := EmptyStr;

      aTmp :=  IntToStr(
                StrToInt( copy(dtmMainH.adoInv007.FieldByName( 'INVDATE' ).AsString,1,4) ) - 1911);
      if Length( aTmp ) = 2 then
        aTmp := '0' + aTmp;
      //#5966 �p�G�O��Ƥ�n�[1 By Kin 2011/04/11
      if StrToInt( copy(dtmMainH.adoInv007.FieldByName( 'INVDATE' ).AsString,6,2) ) mod 2 = 0 then
        aTmp := aTmp + _Right( '00' + copy(dtmMainH.adoInv007.FieldByName( 'INVDATE' ).AsString,6,2) ,2)
      else
        aTmp := aTmp + _Right( '00' + IntToStr( StrToInt( copy(dtmMainH.adoInv007.FieldByName( 'INVDATE' ).AsString,6,2) ) + 1 ),2);

//      aTmp := aTmp + copy(dtmMainH.adoInv007.FieldByName( 'INVDATE' ).AsString,6,2);
      aTmp := aTmp + dtmMainH.adoInv007.FieldByName( 'INVID' ).AsString;
      aTmp := aTmp + adjString(8,aBusinessId,False,False);
      aTmp := aTmp + adjString(1,aCity,False,False);
      aTmp := aTmp + adjString(9,aTaxId,False,False);
      aTmp := aTmp + dtmMainH.adoInv007.FieldByName( 'PRIZETYPE' ).AsString;
      aTmp := aTmp + adjString(52,aManager,False,False);
      aTmp := aTmp + adjString( 80,aCompAddr,false,false);
      //#5966 �W�[�S�O���A�ҥH�n�L�k��AsInteger By Kin 2011/04/11
      if ( dtmMainH.adoInv007.FieldByName( 'PRIZETYPE' ).AsString = '4' ) or
          ( dtmMainH.adoInv007.FieldByName( 'PRIZETYPE' ).AsString = '5' ) or
          ( dtmMainH.adoInv007.FieldByName( 'PRIZETYPE' ).AsString = '6' ) then
      begin
        {#6262 �|�B������.IFF�ɡG�Y�������o�����@�o�B����~�H�B�ιs�|�v�A�N���n���ܸͦ��ɮפ� By Kin 2012/06/25}
        {#6390 �|�B�����Y�������o�����@�o�B����~�H�B�ιs�|�v�B�w���Ƨ����A�N���n���ܸͦ��ɮפ� By Kin 2012/12/14}
        if ( dtmMainH.adoInv007.FieldByName( 'BUSINESSID' ).AsString <> '' ) or
          ( dtmMainH.adoInv007.FieldByName( 'ISOBSOLETE' ).AsString = 'Y' ) or
          ( dtmMainH.adoInv007.FieldByName( 'TaxType' ).AsString = '�s�|�v' ) or
          ( dtmMainH.adoInv007.FieldByName( 'TaxType' ).AsString = '2' ) or
           ( dtmMainH.adoInv007.FieldByName( 'InvAmount' ).AsFloat =
          dtmMainH.getInv014Amt(aInvId,aCompId )) then
        begin
          aTmp := EmptyStr;
        end else
        begin
          aTmp := aTmp + adjString(8,aBranchNo,false,false);
          L_TxtList2.Add( aTmp );
        end;

      end else
      begin
        {#6262 �p�GTaxType = 2 �h��������]�ϥ�10 By Kin 2012/06/22}
        if ( dtmMainH.adoInv007.FieldByName( 'TaxType' ).AsString = '�s�|�v' ) or
         ( dtmMainH.adoInv007.FieldByName( 'TaxType' ).AsString = '2' ) then
        begin
          aTmp := aTmp + '10';
        end else
        begin
          if ( dtmMainH.adoInv007.FieldByName( 'BUSINESSID' ).AsString = '' )
            and (dtmMainH.adoInv007.FieldByName( 'ISOBSOLETE' ).AsString <> 'Y') then
          begin
            {#6390 �T���H�W.IMF�ɭY�������o�����@�o�B����~�H�B�ιs�|�v�Τw���Ƨ����A���n���ͦ��ɮסA���O�ݵ��O��������]�N�� By Kin 2012/12/14}
            if ( dtmMainH.adoInv007.FieldByName( 'InvAmount' ).AsFloat =
              dtmMainH.getInv014Amt( aInvId,aCompId )) then
              aTmp := aTmp + '05'
            else
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
      L_TxtList2.SaveToFile(aPath + aBusinessId + '.IFF' );
    WarningMsg( '���ͧ����I' );
  finally

    L_TxtList1.Free;
    L_TxtList2.Free;
  end;

end;

procedure TfrmInvB06_2.btnExitClick(Sender: TObject);
begin
  Self.Close;
end;

procedure TfrmInvB06_2.GetInv001FieldData(var aBusinessId, aTaxId,
  aManager, aCompAddr: String;aCompId : String);
var aSQL : String;
begin
    aSQL := Format('SELECT * FROM INV001 WHERE COMPID = ''%s''  ',
            [aCompId]);
    dtmMain.adoComm.Close;
    dtmMain.adoComm.SQL.Text := aSQL;
    dtmMain.adoComm.Open;
    aBusinessId := dtmMain.adoComm.FieldByName('BUSINESSID').AsString;
    aTaxId := dtmMain.adoComm.FieldByName( 'TAXID' ).AsString;
    //�쥻�O��Manger���,�{�b���COMPNAME (�S���D��,Jacy�oMail�q��) By Kin 2010/10/08
    aManager := dtmMain.adoComm.FieldByName( 'COMPNAME' ).AsString;
    aCompAddr := dtmMain.adoComm.FieldByName( 'COMPADDR' ).AsString;
end;

procedure TfrmInvB06_2.GetInv003FieldData(var aCity, aBranchNo: String;
  aCompId: String);
  var  aSQL : String;
begin
  dtmMain.adoComm.Close;
  aSQL := Format('SELECT * FROM INV003 WHERE COMPID = ''%s''  ',
              [aCompId]);
  dtmMain.adoComm.SQL.Text := aSQL;
  dtmMain.adoComm.Open;
  aCity := dtmMain.adoComm.FieldByName( 'CITY' ).AsString;
  aBranchNo := dtmMain.adoComm.FieldByName( 'BRANCHNO' ).AsString;
  dtmMain.adoComm.Close;
end;

end.
