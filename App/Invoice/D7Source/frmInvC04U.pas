unit frmInvC04U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Mask, fraYMU, Buttons, ExtCtrls, fraYMDU, Grids,
  DBGrids, DB, frxClass, frxDBSet, frxDesgn;

type
  TfrmInvC04 = class(TForm)
    Panel1: TPanel;
    Label5: TLabel;
    Panel2: TPanel;
    Label7: TLabel;
    cmbHowToCreate: TComboBox;
    btnExit: TBitBtn;
    btnQuery: TBitBtn;
    Label2: TLabel;
    fraEInvDate: TfraYMD;
    fraSInvDate: TfraYMD;
    rgpIsObsolete: TRadioGroup;
    procedure FormShow(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure fraSInvDateExit(Sender: TObject);
    procedure btnQueryExClick(Sender: TObject);
  private
    { Private declarations }
    function IsDataOK : Boolean;
    procedure UnPrepareDataSet;
    procedure PrepareDataSet;
    function GetReportData: Integer;

  public
    { Public declarations }
  end;

var
  frmInvC04: TfrmInvC04;

implementation

uses cbUtilis, frmMainU, dtmMainU, dtmMainJU, //rptInvC04_1U,
  dtmReportModule;

{$R *.dfm}

procedure TfrmInvC04.FormShow(Sender: TObject);
begin
  Self.Caption := frmMain.GetFormTitleString('C04','各種開立方式一覽表');
  if ( fraSInvDate.mseYMD.CanFocus ) then fraSInvDate.mseYMD.SetFocus;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC04.btnExitClick(Sender: TObject);
begin
  Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC04.fraSInvDateExit(Sender: TObject);
begin
   fraEInvDate.setYMD(fraSInvDate.getYMD);
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvC04.IsDataOK: Boolean;
var
  sL_SInvDate, sL_EInvDate : String;
begin
  Result := True;
  sL_SInvDate := Trim(fraSInvDate.getYMD);
  sL_EInvDate := Trim(fraEInvDate.getYMD);
  //檢核發票日期
  if sL_SInvDate <> EMPTY_DATE_STR then
  begin
    if sL_EInvDate=EMPTY_DATE_STR then
    begin
      WarningMsg( '請輸入發票結束日期。' );
      fraEInvDate.mseYMD.SetFocus;
      Result := False;
      Exit;
    end;
    if sL_EInvDate < sL_SInvDate then
    begin
      WarningMsg( '結束日期必須大於等於開始日期。' );
      fraEInvDate.mseYMD.SetFocus;
      Result := False;
      Exit;
    end;
  end else
  begin
    WarningMsg( '請輸入發票日期。' );
    fraSInvDate.mseYMD.SetFocus;
    Result := False;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC04.btnQueryExClick(Sender: TObject);
var
  aPath, aText, aInvDateSt, aInvDateEd, aHowToCreateText: String;
begin
  if not isDataOK then Exit;
  Screen.Cursor := crSQLWait;
  try
    UnPrepareDataSet;
    PrepareDataSet;
    if ( GetReportData <= 0 ) then
    begin
      WarningMsg( '此發票日期區間內無符合資料。' );
      if ( fraSInvDate.mseYMD.CanFocus ) then fraSInvDate.mseYMD.SetFocus;
      Exit;
    end;
  finally
    Screen.Cursor := crDefault;
  end;
  aPath := IncludeTrailingPathDelimiter( ExtractFilePath(
    Application.ExeName ) ) + IncludeTrailingPathDelimiter(
    REPORT_FOLDER ) + 'FrptInvC04_1.fr3';
  dtmReport.frxMainReport.LoadFromFile( aPath );
  {}
  aInvDateSt := Trim( fraSInvDate.getYMD );
  aInvDateEd := Trim( fraEInvDate.getYMD );
  aText := Format( '公司簡稱:%s   發票日期:%s~%s',
    [dtmMain.getCompName, aInvDateSt, aInvDateEd] );
  dtmReport.frxMainReport.Variables.Variables['Condition1'] := QuotedStr( aText );
  {}
  aText := Copy( cmbHowToCreate.Text, Pos( '-',
    cmbHowToCreate.Text ) + 1, Length( cmbHowToCreate.Text ) - Pos( '-', cmbHowToCreate.Text ) );
  aText := Format( '開立別:%s', [aText] );
  case rgpIsObsolete.ItemIndex of
    0: aText := ( aText + '   作廢否:是' );
    1: aText := ( aText + '   作廢否:否' );
    2: aText := ( aText + '   作廢否:全部' );
  end;
  dtmReport.frxMainReport.Variables.Variables['Condition2'] := QuotedStr( aText );
  {}
  dtmReport.frxMainReport.Variables.Variables['Operator'] := QuotedStr( dtmMain.getLoginUser );
  {}
  dtmReport.frxMainReport.ShowReport;
  UnPrepareDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC04.PrepareDataSet;
begin
  dtmReport.cdsTempory.FieldDefs.Clear;
  dtmReport.cdsTempory.FieldDefs.Add( 'ISOBSOLETE', ftString, 1 );
  dtmReport.cdsTempory.FieldDefs.Add( 'INVDATE', ftString, 10 );
  dtmReport.cdsTempory.FieldDefs.Add( 'AMOUNT_HOWTOCREATE_1', ftInteger );
  dtmReport.cdsTempory.FieldDefs.Add( 'AMOUNT_HOWTOCREATE_2', ftInteger );
  dtmReport.cdsTempory.FieldDefs.Add( 'AMOUNT_HOWTOCREATE_3', ftInteger );
  dtmReport.cdsTempory.FieldDefs.Add( 'AMOUNT_HOWTOCREATE_4', ftInteger );
  dtmReport.cdsTempory.FieldDefs.Add( 'COUNT_HOWTOCREATE_1', ftInteger );
  dtmReport.cdsTempory.FieldDefs.Add( 'COUNT_HOWTOCREATE_2', ftInteger );
  dtmReport.cdsTempory.FieldDefs.Add( 'COUNT_HOWTOCREATE_3', ftInteger );
  dtmReport.cdsTempory.FieldDefs.Add( 'COUNT_HOWTOCREATE_4', ftInteger );
  dtmReport.cdsTempory.FieldDefs.Add( 'TOTALAMOUNT', ftInteger );
  dtmReport.cdsTempory.FieldDefs.Add( 'TOTALCOUNT', ftInteger);
  dtmReport.cdsTempory.CreateDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC04.UnPrepareDataSet;
begin
  if not VarIsNull( dtmReport.cdsTempory.Data ) then
    dtmReport.cdsTempory.EmptyDataSet;
  dtmReport.cdsTempory.Data := Null;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvC04.GetReportData: Integer;
var
  aSql, aSubSql, aWhere, aInvDateSt, aInvDateEd, aHowToCreate,
  aIsObsolete: String;
begin
  Result := 0;
  aInvDateSt := Trim( fraSInvDate.getYMD );
  aInvDateEd := Trim( fraEInvDate.getYMD );
  aWhere := Format(
    '    where identifyid1 = ''%s''                          ' +
    '      and IdentifyId2 = ''%s''                          ' +
    '      and compid = ''%s''                               ' +
    '      and invdate between                               ' +
    '          to_date( ''%s'', ''YYYY/MM/DD'' ) and         ' +
    '          to_date( ''%s'', ''YYYY/MM/DD'' )             ',
    [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, aInvDateSt, aInvDateEd] );
  if ( cmbHowToCreate.ItemIndex > 0 ) then
  begin
    aHowToCreate := Copy( cmbHowToCreate.Text, 1, 1 );
    aWhere := ( aWhere + Format( ' and howtocreate = ''%s'' ', [aHowToCreate] ) );
  end;
  if ( rgpIsObsolete.ItemIndex in [0..1] ) then
  begin
    aIsObsolete := IIF( rgpIsObsolete.ItemIndex = 0, 'Y', 'N' );
    aWhere := ( aWhere + Format( ' and isobsolete = ''%s'' ', [aIsObsolete] ) );
  end;
  aSubSql := Format(
    '   select nvl( isobsolete, ''N'' ) as isobsolete,         ' +
    '          to_char( invdate, ''YYYY/MM/DD'' ) as invdate,  ' +
    '          howtocreate,                                    ' +
    '          sum( invamount ) as totalamount,                ' +
    '          count( invid ) as counts                        ' +
    '     from inv007                                          ' +
    '      %s                                                  ' +
    '      group by isobsolete, invdate, howtocreate           ', [aWhere] );
  aSql := Format(
    ' select * from ( %s )                       ' +
    ' order by isobsolete, invdate, howtocreate  ', [aSubSql] );
  dtmMainJ.adoQryCommon.Close;
  dtmMainJ.adoQryCommon.SQL.Text := aSql;
  dtmMainJ.adoQryCommon.Open;
  dtmMainJ.adoQryCommon.First;
  while not dtmMainJ.adoQryCommon.Eof do
  begin
    if dtmReport.cdsTempory.Locate( 'isobsolete;invdate', VarArrayOf( [
      dtmMainJ.adoQryCommon.FieldByName( 'isobsolete' ).AsString,
      dtmMainJ.adoQryCommon.FieldByName( 'invdate' ).AsString] ), [] ) then
    begin
      dtmReport.cdsTempory.Edit;
    end else
    begin
      dtmReport.cdsTempory.Last;
      dtmReport.cdsTempory.Append;
      dtmReport.cdsTempory.FieldByName( 'isobsolete' ).AsString :=
        dtmMainJ.adoQryCommon.FieldByName( 'isobsolete' ).AsString;
      dtmReport.cdsTempory.FieldByName( 'invdate' ).AsString :=
        dtmMainJ.adoQryCommon.FieldByName( 'invdate' ).AsString;
      dtmReport.cdsTempory.FieldByName( 'AMOUNT_HOWTOCREATE_1' ).AsInteger := 0;
      dtmReport.cdsTempory.FieldByName( 'AMOUNT_HOWTOCREATE_2' ).AsInteger := 0;
      dtmReport.cdsTempory.FieldByName( 'AMOUNT_HOWTOCREATE_3' ).AsInteger := 0;
      dtmReport.cdsTempory.FieldByName( 'AMOUNT_HOWTOCREATE_4' ).AsInteger := 0;
      dtmReport.cdsTempory.FieldByName( 'COUNT_HOWTOCREATE_1' ).AsInteger := 0;
      dtmReport.cdsTempory.FieldByName( 'COUNT_HOWTOCREATE_2' ).AsInteger := 0;
      dtmReport.cdsTempory.FieldByName( 'COUNT_HOWTOCREATE_3' ).AsInteger := 0;
      dtmReport.cdsTempory.FieldByName( 'COUNT_HOWTOCREATE_4' ).AsInteger := 0;
      dtmReport.cdsTempory.FieldByName( 'TOTALAMOUNT' ).AsInteger := 0;
      dtmReport.cdsTempory.FieldByName( 'TOTALCOUNT' ).AsInteger := 0;
    end;
    if ( dtmMainJ.adoQryCommon.FieldByName( 'howtocreate' ).AsString = '1' ) then
    begin
      dtmReport.cdsTempory.FieldByName( 'AMOUNT_HOWTOCREATE_1' ).AsInteger :=
       ( dtmReport.cdsTempory.FieldByName( 'AMOUNT_HOWTOCREATE_1' ).AsInteger +
         dtmMainJ.adoQryCommon.FieldByName( 'totalamount' ).AsInteger );
      dtmReport.cdsTempory.FieldByName( 'COUNT_HOWTOCREATE_1' ).AsInteger :=
       ( dtmReport.cdsTempory.FieldByName( 'COUNT_HOWTOCREATE_1' ).AsInteger +
         dtmMainJ.adoQryCommon.FieldByName( 'counts' ).AsInteger );
    end else
    if ( dtmMainJ.adoQryCommon.FieldByName( 'howtocreate' ).AsString = '2' ) then
    begin
      dtmReport.cdsTempory.FieldByName( 'AMOUNT_HOWTOCREATE_2' ).AsInteger :=
       ( dtmReport.cdsTempory.FieldByName( 'AMOUNT_HOWTOCREATE_2' ).AsInteger +
         dtmMainJ.adoQryCommon.FieldByName( 'totalamount' ).AsInteger );
      dtmReport.cdsTempory.FieldByName( 'COUNT_HOWTOCREATE_2' ).AsInteger :=
       ( dtmReport.cdsTempory.FieldByName( 'COUNT_HOWTOCREATE_2' ).AsInteger +
         dtmMainJ.adoQryCommon.FieldByName( 'counts' ).AsInteger );
    end else
    if ( dtmMainJ.adoQryCommon.FieldByName( 'howtocreate' ).AsString = '3' ) then
    begin
      dtmReport.cdsTempory.FieldByName( 'AMOUNT_HOWTOCREATE_3' ).AsInteger :=
       ( dtmReport.cdsTempory.FieldByName( 'AMOUNT_HOWTOCREATE_3' ).AsInteger +
         dtmMainJ.adoQryCommon.FieldByName( 'totalamount' ).AsInteger );
      dtmReport.cdsTempory.FieldByName( 'COUNT_HOWTOCREATE_3' ).AsInteger :=
       ( dtmReport.cdsTempory.FieldByName( 'COUNT_HOWTOCREATE_3' ).AsInteger +
         dtmMainJ.adoQryCommon.FieldByName( 'counts' ).AsInteger );
    end else
    if ( dtmMainJ.adoQryCommon.FieldByName( 'howtocreate' ).AsString = '4' ) then
    begin
      dtmReport.cdsTempory.FieldByName( 'AMOUNT_HOWTOCREATE_4' ).AsInteger :=
       ( dtmReport.cdsTempory.FieldByName( 'AMOUNT_HOWTOCREATE_4' ).AsInteger +
         dtmMainJ.adoQryCommon.FieldByName( 'totalamount' ).AsInteger );
      dtmReport.cdsTempory.FieldByName( 'COUNT_HOWTOCREATE_4' ).AsInteger :=
       ( dtmReport.cdsTempory.FieldByName( 'COUNT_HOWTOCREATE_4' ).AsInteger +
         dtmMainJ.adoQryCommon.FieldByName( 'counts' ).AsInteger );
    end;
    dtmReport.cdsTempory.FieldByName( 'TOTALAMOUNT' ).AsInteger :=
      ( dtmReport.cdsTempory.FieldByName( 'TOTALAMOUNT' ).AsInteger +
        dtmMainJ.adoQryCommon.FieldByName( 'totalamount' ).AsInteger );
    dtmReport.cdsTempory.FieldByName( 'TOTALCOUNT' ).AsInteger :=
      ( dtmReport.cdsTempory.FieldByName( 'TOTALCOUNT' ).AsInteger +
        dtmMainJ.adoQryCommon.FieldByName( 'counts' ).AsInteger );
    dtmReport.cdsTempory.Post;
    dtmMainJ.adoQryCommon.Next;
  end;
  dtmMainJ.adoQryCommon.Close;
  Result := dtmReport.cdsTempory.RecordCount;
end;

{ ---------------------------------------------------------------------------- }

end.
