unit frmInvA05_1U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Buttons, DB, dtmReportModule;

type
  TfrmInvA05_1 = class(TForm)
    Panel1: TPanel;
    btnExit: TBitBtn;
    btnPrintNew: TBitBtn;
    rgpPrintType: TRadioGroup;
    Panel2: TPanel;
    Label1: TLabel;
    procedure btnExitClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
    FCallFromProgram: TClass;
    FInvId: String;
    procedure UnPrepareDataSet;
    procedure PrepareDataSet;
    function InvDataFillToDataSet: Boolean;
    procedure AfterPrintReport;
  public
    { Public declarations }
    sG_UserID, sG_CompID, sG_PrintTypeFile, sG_OrderByType : String;
    property CallFromProgram: TClass read FCallFromProgram write FCallFromProgram;
    property InvId: String read FInvId write FInvId;
  end;

var
  frmInvA05_1: TfrmInvA05_1;

implementation

uses dtmMainJU, frmInvA05_3U, dtmMainU, frmMainU, Ustru, cbUtilis, frmInvA05U,
  frmInvA02_4U;

{$R *.dfm}

procedure TfrmInvA05_1.btnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmInvA05_1.FormCreate(Sender: TObject);
var
  sL_PrintType : String;
  L_Txt : TStringList;
  nL_PrintType: Integer;
begin
  sG_PrintTypeFile := IncludeTrailingPathDelimiter( ExtractFilePath(
    Application.ExeName ) ) + PRINT_TYPE_FILE;

  if FileExists( sG_PrintTypeFile ) then //拆解Text File的部份
  begin
    L_Txt := TStringList.Create;
    try
      L_Txt.LoadFromFile( sG_PrintTypeFile );
      if L_Txt.GetText <> '' then
        sL_PrintType := L_Txt.Strings[0]
      else
        sL_PrintType := '0';
    finally
      L_Txt.Free;
    end;

    nL_PrintType := StrToIntDef(sL_PrintType,0);

    if nL_PrintType = 1 then  //1=群健規格
    begin
      sL_PrintType := '1';
      nL_PrintType :=1;
    end
    else if nL_PrintType = 2 then  //2=TBC新規格
    begin
      sL_PrintType := '2';
      nL_PrintType := 2;
    end
    else  //其他都預設為預設值0
    begin
      sL_PrintType := '0';
      nL_PrintType := 0;
    end;

    rgpPrintType.ItemIndex := nL_PrintType;


  end
  else
  begin
    rgpPrintType.ItemIndex := 0;
    sL_PrintType := '0';
  end;
end;

procedure TfrmInvA05_1.FormShow(Sender: TObject);
begin
  self.Caption := frmMain.GetFormTitleString( 'A05_1', '列印發票型式' );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA05_1.Button1Click(Sender: TObject);
var
  aModalResult: Integer;
  aTxt : TStringList;
  aPath, aPrintTitleSet: String ;
  aPrintTitle, aPrintMailAddr, aPrintMemo: String;
begin
  if not ConfirmMsg( '請確認您的印表機輸出紙張大小設定是否正確。' ) then Exit;
  {}
  UnPrepareDataSet;
  PrepareDataSet;
  {}
  if not InvDataFillToDataSet then
  begin
    WarningMsg( '沒有發票資料可供列印!' );
    Exit;
  end;

  aPrintTitleSet := dtmMainJ.getIsPrintTitle( sG_CompID );
  if aPrintTitleSet = EmptyStr then aPrintTitle := 'Y';
  dtmReport.AddA05DataField( dtmMain.GetAutoCreateNum );
  
  if rgpPrintType.ItemIndex = 0 then
  begin
    aPath := IncludeTrailingPathDelimiter( ExtractFilePath(
      Application.ExeName ) ) + IncludeTrailingPathDelimiter(
      REPORT_FOLDER ) + 'FrptInvA05_1.fr3';
    dtmReport.frxMainReport.LoadFromFile( aPath );
    dtmReport.frxMainReport.Variables.Variables['列印發票抬頭'] :=
      QuotedStr( aPrintTitleSet );
    dtmReport.frxMainReport.ShowReport;
    AfterPrintReport;
    dtmReport.cdsTempory.EmptyDataSet;
  end
  else if rgpPrintType.ItemIndex = 1 then
  begin
    aPrintTitle := 'N';
    aPrintMailAddr := 'N';
    aPrintMemo := 'N';
    frmInvA05_3 := TfrmInvA05_3.Create( Application );
    try
      aModalResult := frmInvA05_3.ShowModal;
      if frmInvA05_3.ckbInvTitle.Checked then aPrintTitle := 'Y';
      if frmInvA05_3.ckbMailAddr.Checked then aPrintMailAddr := 'Y';
      if frmInvA05_3.ckbMemo1.Checked then aPrintMemo := 'Y';
    finally
      frmInvA05_3.Free;
    end;
    if aModalResult = mrOK then
    begin
      //{$IFDEF DEBUG}
      //  aPath := '..\Source\ReportTemplate\FrptInvA05_2.fr3';
      //{$ELSE}
        aPath := IncludeTrailingPathDelimiter( ExtractFilePath(
          Application.ExeName ) ) + IncludeTrailingPathDelimiter(
          REPORT_FOLDER ) + 'FrptInvA05_2.fr3';
      //{$ENDIF}
      dtmReport.frxMainReport.LoadFromFile( aPath );
      dtmReport.frxMainReport.Variables.Variables['列印發票抬頭'] :=
        QuotedStr( aPrintTitle );
      dtmReport.frxMainReport.Variables.Variables['列印發票地址'] :=
        QuotedStr( aPrintMailAddr );
      dtmReport.frxMainReport.Variables.Variables['列印備註'] :=
        QuotedStr( aPrintMemo );
      dtmReport.frxMainReport.ShowReport;
      AfterPrintReport;
      dtmReport.cdsTempory.EmptyDataSet;
    end;
  end
  else if rgpPrintType.ItemIndex = 2 then
  begin
    //{$IFDEF DEBUG}
    //  aPath := '..\Source\ReportTemplate\FrptInvA05_3.fr3';
    //{$ELSE}
      aPath := IncludeTrailingPathDelimiter( ExtractFilePath(
        Application.ExeName ) ) + IncludeTrailingPathDelimiter(
        REPORT_FOLDER ) + 'FrptInvA05_3.fr3';
    //{$ENDIF}

    dtmReport.frxMainReport.LoadFromFile( aPath );
    dtmReport.frxMainReport.Variables.Variables['列印發票抬頭'] :=
      QuotedStr( aPrintTitleSet );
    dtmReport.frxMainReport.ShowReport;
    AfterPrintReport;
    dtmReport.cdsTempory.EmptyDataSet;
  end;
  aTxt := TStringList.Create;
  try
    case rgpPrintType.ItemIndex of
      0: aTxt.Add( '0' );    //TBC格式
      1: aTxt.Add( '1' );     //群健格式
      2: aTxt.Add( '2' );    //TBC新格式
    end;
    if FileExists( sG_PrintTypeFile ) then DeleteFile( sG_PrintTypeFile );
    aTxt.SaveToFile( sG_PrintTypeFile );
  finally
    aTxt.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }


procedure TfrmInvA05_1.UnPrepareDataSet;
begin
  if not VarIsNull( dtmReport.cdsTempory.Data ) then
    dtmReport.cdsTempory.Data := Null;
  dtmReport.cdsTempory.FieldDefs.Clear;

end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA05_1.PrepareDataSet;
var
  aAutoCreateNum, aIndex: Integer;
begin
  dtmReport.cdsTempory.FieldDefs.Clear;
  dtmReport.cdsTempory.FieldDefs.Add( 'CHECKNO', ftString, 2 );
  dtmReport.cdsTempory.FieldDefs.Add( 'INVID', ftString, 10 );
  dtmReport.cdsTempory.FieldDefs.Add( 'BUSINESSID', ftString, 8 );
  dtmReport.cdsTempory.FieldDefs.Add( 'INVDATE', ftDateTime );
  dtmReport.cdsTempory.FieldDefs.Add( 'CUSTID', ftString, 100 );
  dtmReport.cdsTempory.FieldDefs.Add( 'CUSTSNAME', ftString, 100 );
  dtmReport.cdsTempory.FieldDefs.Add( 'INVTITLE', ftString, 100 );
  dtmReport.cdsTempory.FieldDefs.Add( 'ZIPCODE', ftString, 100 );
  dtmReport.cdsTempory.FieldDefs.Add( 'MAILADDR', ftString, 100 );
  dtmReport.cdsTempory.FieldDefs.Add( 'INVADDR', ftString, 100 );
  dtmReport.cdsTempory.FieldDefs.Add( 'INSTADDR', ftString, 100 );
  dtmReport.cdsTempory.FieldDefs.Add( 'INVFORMAT', ftString, 100 );
  dtmReport.cdsTempory.FieldDefs.Add( 'TAXTYPE', ftString, 100 );
  dtmReport.cdsTempory.FieldDefs.Add( 'SALEAMOUNT', ftInteger );
  dtmReport.cdsTempory.FieldDefs.Add( 'TAXAMOUNT', ftInteger );
  dtmReport.cdsTempory.FieldDefs.Add( 'INVAMOUNT', ftInteger );
  {}
  dtmReport.cdsTempory.FieldDefs.Add( 'MAININVID', ftString, 10 );
  dtmReport.cdsTempory.FieldDefs.Add( 'MAINSALEAMOUNT', ftInteger );
  dtmReport.cdsTempory.FieldDefs.Add( 'MAINTAXAMOUNT', ftInteger );
  dtmReport.cdsTempory.FieldDefs.Add( 'MAININVAMOUNT', ftInteger );
  {}
  aAutoCreateNum := dtmMain.GetAutoCreateNum;
  if aAutoCreateNum <= 0 then aAutoCreateNum := 5;
  {}
  for aIndex := 1 to aAutoCreateNum do
  begin
    dtmReport.cdsTempory.FieldDefs.Add( Format( 'BILLID%d', [aIndex] ), ftString, 15 );
    dtmReport.cdsTempory.FieldDefs.Add( Format( 'DESCRIPTION%d', [aIndex] ), ftString, 200 );
    dtmReport.cdsTempory.FieldDefs.Add( Format( 'STARTDATE%d', [aIndex] ), ftDateTime );
    dtmReport.cdsTempory.FieldDefs.Add( Format( 'ENDDATE%d', [aIndex] ), ftDateTime );
    dtmReport.cdsTempory.FieldDefs.Add( Format( 'QUANTITY%d', [aIndex] ), ftInteger );
    dtmReport.cdsTempory.FieldDefs.Add( Format( 'UNITPRICE%d', [aIndex] ), ftFloat );  //#4370 改成浮點數 By Kin 2009/02/19
    dtmReport.cdsTempory.FieldDefs.Add( Format( 'SALEAMOUNT%d', [aIndex] ), ftInteger );
    dtmReport.cdsTempory.FieldDefs.Add( Format( 'TOTALAMOUNT%d', [aIndex] ), ftInteger );
    dtmReport.cdsTempory.FieldDefs.Add( Format( 'FACISNO%d', [aIndex] ), ftString, 20 );
  end;
  {}
  dtmReport.cdsTempory.FieldDefs.Add( 'MEMO1', ftString, 60 );
  dtmReport.cdsTempory.FieldDefs.Add( 'ACCOUNTNO', ftString, 1000 );
  dtmReport.cdsTempory.FieldDefs.Add( 'HOWTOCREATE', ftString, 1 );
  dtmReport.cdsTempory.FieldDefs.Add( 'INVUSEDESC', ftString, 20 );
  {}
  dtmReport.cdsTempory.CreateDataSet;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA05_1.InvDataFillToDataSet: Boolean;
var
  aSQL, aLastInvId, aDescription, aDaySt, aDayEd, aFacisNo: String;
  aCurrentFieldNumber: Integer;
begin
  if CallFromProgram = TfrmInvA05 then
  begin
    aSQL := Format(
      ' SELECT A.CHECKNO, A.INVID, A.BUSINESSID, A.INVDATE, A.CUSTID, ' +
      '        A.CUSTSNAME, A.INVFORMAT, A.INVTITLE, A.ZIPCODE,       ' +
      '        A.MAILADDR, A.TAXTYPE, A.SALEAMOUNT, A.TAXAMOUNT,      ' +
      '        A.INVAMOUNT, A.MEMO1, B.DESCRIPTION, B.STARTDATE,      ' +
      '        B.ENDDATE, B.QUANTITY, B.UNITPRICE, B.TOTALAMOUNT,     ' +
      '        B.SEQ, A.MAININVID, A.INVADDR,                         ' +
      '        A.MAINSALEAMOUNT, A.MAINTAXAMOUNT, A.MAININVAMOUNT,    ' +
      '        B.SALEAMOUNT AS SALEAMOUNT2, B.BILLID, A.INSTADDR,     ' +
      '        A.HOWTOCREATE, A.INVUSEDESC, C.IFPRINTTITLE            ' +
      '   FROM INV023 A, INV008 B, INV002 C                           ' +
      '  WHERE A.IDENTIFYID1 = B.IDENTIFYID1                          ' +
      '    AND A.IDENTIFYID2 = B.IDENTIFYID2                          ' +
      '    AND A.INVID = B.INVID                                      ' +
      '    AND A.IDENTIFYID1 = C.IDENTIFYID1(+)                       ' +
      '    AND A.IDENTIFYID2 = C.IDENTIFYID2(+)                       ' +
      '    AND A.COMPID = C.COMPID(+)                                 ' +
      '    AND A.CUSTID = C.CUSTID(+)                                 ' +
      '    AND A.IDENTIFYID1 = ''%s''                                 ' +
      '    AND A.IDENTIFYID2 = ''%s''                                 ' +
      '    AND A.COMPID = ''%s''                                      ' +
      '    AND A.OWNER = ''%s''                                       ',
      [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, sG_UserID ] );
      if sG_OrderByType = '1' then
        aSQL := ( aSQL + ' ORDER BY A.INVID, B.SEQ ' )
      else if sG_OrderByType = '2' then
        aSQL := ( aSQL + ' ORDER BY A.INVDATE, A.INVID, B.SEQ ' )
      else if sG_OrderByType = '3' then
        aSQL := ( aSQL + ' ORDER BY A.CUSTID, A.INVID, B.SEQ ' );
  end else
  if CallFromProgram = TfrmInvA02_4 then
  begin
    aSQL := Format(
     '  SELECT A.CHECKNO, A.INVID, A.BUSINESSID, A.INVDATE, A.CUSTID, ' +
     '         A.CUSTSNAME, A.INVFORMAT, A.INVTITLE, A.ZIPCODE,       ' +
     '         A.MAILADDR, A.TAXTYPE, A.SALEAMOUNT, A.TAXAMOUNT,      ' +
     '         A.INVAMOUNT, A.MEMO1, B.DESCRIPTION, B.STARTDATE,      ' +
     '         B.ENDDATE, B.QUANTITY, B.UNITPRICE, B.TOTALAMOUNT,     ' +
     '         B.SEQ, A.MAININVID, A.INVADDR,                         ' +
     '         A.MAINSALEAMOUNT, A.MAINTAXAMOUNT, A.MAININVAMOUNT,    ' +
     '         B.SALEAMOUNT AS SALEAMOUNT2, B.BILLID, A.INSTADDR,     ' +
     '         A.HOWTOCREATE, A.INVUSEDESC, D.IFPRINTTITLE            ' +
     '    FROM INV007 A, INV008 B,                                    ' +
     '    ( SELECT MAININVID FROM INV007                              ' +
     '       WHERE IDENTIFYID1 = ''%s''                               ' +
     '         AND IDENTIFYID2 = ''%s''                               ' +
     '         AND COMPID = ''%s''                                    ' +
     '         AND INVID = ''%s''                                     ' +
     '     ) C, INV002 D                                              ' +
     '   WHERE A.IDENTIFYID1 = B.IDENTIFYID1                          ' +
     '     AND A.IDENTIFYID2 = B.IDENTIFYID2                          ' +
     '     AND A.INVID = B.INVID                                      ' +
     '     AND A.MAININVID = C.MAININVID                              ' +
     '     AND A.IDENTIFYID1 = D.IDENTIFYID1(+)                       ' +
     '     AND A.IDENTIFYID2 = D.IDENTIFYID2(+)                       ' +
     '     AND A.COMPID = D.COMPID(+)                                 ' +
     '     AND A.CUSTID = D.CUSTID(+)                                 ' +
     '     AND A.IDENTIFYID1 = ''%s''                                 ' +
     '     AND A.IDENTIFYID2 = ''%s''                                 ' +
     '     AND A.COMPID = ''%s''                                      ' +
     '  ORDER BY B.INVID, B.SEQ                                       ',
     [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, FInvId,
      IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID] );
  end;
  {}
  dtmMain.adoComm.Close;
  dtmMain.adoComm.SQL.Text := aSQL;
  dtmMain.adoComm.Open;
  Result := not dtmMain.adoComm.IsEmpty;
  if not Result then
  begin
    dtmMain.adoComm.Close;
    Exit;
  end;
  aLastInvId := EmptyStr;
  aCurrentFieldNumber := 0;
  dtmMain.adoComm.First;
  while not dtmMain.adoComm.Eof do
  begin
    if ( aLastInvId <> dtmMain.adoComm.FieldByName( 'INVID' ).AsString ) then
    begin
      aLastInvId := dtmMain.adoComm.FieldByName( 'INVID' ).AsString;
      aCurrentFieldNumber := 1;
      dtmReport.cdsTempory.Append;
    end else
    begin
      Inc( aCurrentFieldNumber );
      dtmReport.cdsTempory.Edit;
    end;
    if ( aCurrentFieldNumber <= 1 ) then
    begin
      dtmReport.cdsTempory.FieldByName( 'CHECKNO' ).AsString :=
        dtmMain.adoComm.FieldByName( 'CHECKNO' ).AsString;
      {}
      dtmReport.cdsTempory.FieldByName( 'INVID' ).AsString :=
        dtmMain.adoComm.FieldByName( 'INVID' ).AsString;
      {}
      dtmReport.cdsTempory.FieldByName( 'BUSINESSID' ).AsString :=
        dtmMain.adoComm.FieldByName( 'BUSINESSID' ).AsString;
      {}
      dtmReport.cdsTempory.FieldByName( 'INVDATE' ).Value :=
        dtmMain.adoComm.FieldByName( 'INVDATE' ).Value;
      {}
      dtmReport.cdsTempory.FieldByName( 'CUSTID' ).AsString :=
        dtmMain.adoComm.FieldByName( 'CUSTID' ).AsString;
      {}
      dtmReport.cdsTempory.FieldByName( 'CUSTSNAME' ).AsString :=
        dtmMain.adoComm.FieldByName( 'CUSTSNAME' ).AsString;
      {}
      { 不論統一編號是否有值, 都填入 INVTITLE }
      if dtmMain.GetIfPrintTitle then
        dtmReport.cdsTempory.FieldByName( 'INVTITLE' ).AsString :=
          dtmMain.adoComm.FieldByName( 'INVTITLE' ).AsString
      else begin
        { 統一編號有值, 才填入 INVTITLE }
        dtmReport.cdsTempory.FieldByName( 'INVTITLE' ).AsString := EmptyStr;
        if ( dtmMain.adoComm.FieldByName( 'BUSINESSID' ).AsString <> EmptyStr ) then
          dtmReport.cdsTempory.FieldByName( 'INVTITLE' ).AsString :=
            dtmMain.adoComm.FieldByName( 'INVTITLE' ).AsString
        { 若統編沒值, 但是 INV002 的 IfPrintTitle = Y 仍然要印 }
        else if ( dtmMain.adoComm.FieldByName( 'IFPRINTTITLE' ).AsString = 'Y' ) then
          dtmReport.cdsTempory.FieldByName( 'INVTITLE' ).AsString :=
            dtmMain.adoComm.FieldByName( 'INVTITLE' ).AsString;
      end;    
      {}
      dtmReport.cdsTempory.FieldByName( 'ZIPCODE' ).AsString :=
        dtmMain.adoComm.FieldByName( 'ZIPCODE' ).AsString;
      // 郵寄地址
      dtmReport.cdsTempory.FieldByName( 'MAILADDR' ).AsString :=
        dtmMain.adoComm.FieldByName( 'MAILADDR' ).AsString;
      // 發票地址
      dtmReport.cdsTempory.FieldByName( 'INVADDR' ).AsString :=
        dtmMain.adoComm.FieldByName( 'INVADDR' ).AsString;
      // 裝機地址
      dtmReport.cdsTempory.FieldByName( 'INSTADDR' ).AsString :=
        dtmMain.adoComm.FieldByName( 'INSTADDR' ).AsString;
      {}
      dtmReport.cdsTempory.FieldByName( 'INVFORMAT' ).AsString :=
        dtmMain.adoComm.FieldByName( 'INVFORMAT' ).AsString;
      {}
      dtmReport.cdsTempory.FieldByName( 'TAXTYPE' ).AsString :=
        dtmMain.adoComm.FieldByName( 'TAXTYPE' ).AsString;
      {}
      dtmReport.cdsTempory.FieldByName( 'SALEAMOUNT' ).AsInteger :=
        dtmMain.adoComm.FieldByName( 'SALEAMOUNT' ).AsInteger;
      {}
      dtmReport.cdsTempory.FieldByName( 'TAXAMOUNT' ).AsInteger :=
        dtmMain.adoComm.FieldByName( 'TAXAMOUNT' ).AsInteger;
      {}
      dtmReport.cdsTempory.FieldByName( 'INVAMOUNT' ).AsInteger :=
        dtmMain.adoComm.FieldByName( 'INVAMOUNT' ).AsInteger;

      dtmReport.cdsTempory.FieldByName( 'MAININVID' ).AsString :=
        dtmMain.adoComm.FieldByName( 'MAININVID' ).AsString;
      {}
      dtmReport.cdsTempory.FieldByName( 'MAINSALEAMOUNT' ).AsInteger :=
        dtmMain.adoComm.FieldByName( 'MAINSALEAMOUNT' ).AsInteger;
      {}
      dtmReport.cdsTempory.FieldByName( 'MAINTAXAMOUNT' ).AsInteger :=
        dtmMain.adoComm.FieldByName( 'MAINTAXAMOUNT' ).AsInteger;
      {}
      dtmReport.cdsTempory.FieldByName( 'MAININVAMOUNT' ).AsInteger :=
        dtmMain.adoComm.FieldByName( 'MAININVAMOUNT' ).AsInteger;
      {}
      dtmReport.cdsTempory.FieldByName( 'MEMO1' ).AsString :=
        dtmMain.adoComm.FieldByName( 'MEMO1' ).AsString;
      {}
      dtmReport.cdsTempory.FieldByName( 'ACCOUNTNO' ).AsString := dtmMainJ.GetAccountNo(
        dtmReport.cdsTempory.FieldByName( 'INVID' ).AsString,
        EmptyStr, EmptyStr, EmptyStr );
      {}
      dtmReport.cdsTempory.FieldByName( 'HOWTOCREATE' ).AsString :=
        dtmMain.adoComm.FieldByName( 'HOWTOCREATE' ).AsString;
      {}  
      dtmReport.cdsTempory.FieldByName( 'INVUSEDESC' ).AsString :=
        dtmMain.adoComm.FieldByName( 'INVUSEDESC' ).AsString;
    end;
    {}
    dtmReport.cdsTempory.FieldByName( Format( 'BILLID%d', [aCurrentFieldNumber] ) ).Value :=
      dtmMain.adoComm.FieldByName( 'BILLID' ).Value;
    {}
    aDescription := dtmMain.adoComm.FieldByName( 'DESCRIPTION' ).AsString;
    aDaySt := EmptyStr;
    aDayEd := EmptyStr;
    if not VarIsNull( dtmMain.adoComm.FieldByName( 'STARTDATE' ).Value ) then
      aDaySt := FormatDateTime( 'eee/mm/dd', dtmMain.adoComm.FieldByName( 'STARTDATE' ).AsDateTime );
    if not VarIsNull( dtmMain.adoComm.FieldByName( 'ENDDATE' ).Value ) then
      aDayEd := FormatDateTime( 'eee/mm/dd', dtmMain.adoComm.FieldByName( 'ENDDATE' ).AsDateTime );
    aDescription := Format( aDescription + '( %s - %s )', [aDaySt, aDayEd] );
    {}
    dtmReport.cdsTempory.FieldByName( Format( 'DESCRIPTION%d', [aCurrentFieldNumber] ) ).AsString :=
      dtmMain.adoComm.FieldByName( 'DESCRIPTION' ).AsString;
    {}
    dtmReport.cdsTempory.FieldByName( Format( 'STARTDATE%d', [aCurrentFieldNumber] ) ).Value :=
      dtmMain.adoComm.FieldByName( 'STARTDATE' ).Value;
    {}
    dtmReport.cdsTempory.FieldByName( Format( 'ENDDATE%d', [aCurrentFieldNumber] ) ).Value :=
      dtmMain.adoComm.FieldByName( 'ENDDATE' ).Value;
    {}
    dtmReport.cdsTempory.FieldByName( Format( 'QUANTITY%d', [aCurrentFieldNumber] ) ).Value :=
      dtmMain.adoComm.FieldByName( 'QUANTITY' ).AsInteger;
    {}
    dtmReport.cdsTempory.FieldByName( Format( 'UNITPRICE%d', [aCurrentFieldNumber] ) ).Value :=
      dtmMain.adoComm.FieldByName( 'UNITPRICE' ).Value;
    {}
    dtmReport.cdsTempory.FieldByName( Format( 'SALEAMOUNT%d', [aCurrentFieldNumber] ) ).Value :=
      dtmMain.adoComm.FieldByName( 'SALEAMOUNT2' ).Value;
    {}
    dtmReport.cdsTempory.FieldByName( Format( 'TOTALAMOUNT%d', [aCurrentFieldNumber] ) ).Value :=
      dtmMain.adoComm.FieldByName( 'TOTALAMOUNT' ).Value;
    {}
    aFacisNo := dtmMainJ.GetFacisNo( dtmMain.adoComm.FieldByName( 'INVID' ).AsString,
        dtmMain.adoComm.FieldByName( 'SEQ' ).AsString );
    if ( Pos( ',', aFacisNo ) > 0 ) then aFacisNo := Copy( aFacisNo, 1, Pos( ',', aFacisNo ) - 1 );
    aFacisNo := dtmMainU.FacisNoAddMask( aFacisNo );
    dtmReport.cdsTempory.FieldByName( Format( 'FACISNO%d', [aCurrentFieldNumber] ) ).Value := aFacisNo;    
    {}
    dtmReport.cdsTempory.Post;
    dtmMain.adoComm.Next;
  end;
  dtmMain.adoComm.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA05_1.AfterPrintReport;
begin
  if FCallFromProgram = TfrmInvA05 then
  begin
    dtmMainJ.AfterPrintOrTrans( 1, nil,False );
  end else
  if FCallFromProgram = TfrmInvA02_4 then
  begin
    dtmMainJ.AfterPrintOrTrans( 2, dtmReport.cdsTempory,False );
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
