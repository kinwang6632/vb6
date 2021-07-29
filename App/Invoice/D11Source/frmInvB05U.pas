unit frmInvB05U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Menus, ADODB, FileCtrl,
  cxLookAndFeelPainters, cxButtons,
  cxControls, cxContainer, cxEdit, cxProgressBar, cxLabel, cxGraphics,
  cxDropDownEdit, cxTextEdit, cxMaskEdit, cxGroupBox, cxRadioGroup,
  cxButtonEdit, dxSkinsCore, dxSkinsDefaultPainters;

type
  { 匯檔用發票公司別參數資料 }
  PZeorTaxCompany = ^TZeroTaxCompany;
  TZeroTaxCompany = record
    CompId: String;
    BusinessId: String;
    City: String;
    TaxId: String;
    ExportMode: String;
    ImportNote: String;
    EntryClass: String;
  end;

  { 匯檔用零稅率 record }
  TZeroTaxRecord = record
    BusinessId: String;     { 發票公司別統編 }
    City: String;           { 發票公司別所屬縣市代碼 }
    TaxId: String;          { 發票公司別稅籍編號 }
    YearMonth: String;      { 申報年月 (民國年) }
    InvDate: String;        { 發票日期 (民國年+月) }
    InvId: String;          { 發票號碼 }
    InvBusinessId: String;  { 買受人統編 (發票統編) }
    ExportMode: String;     { 外銷方式 }
    ImportNote: String;     { 通關方式註記 }
    EntryClass: String;     { 出口報單類別 }
    BillId: String;         { 出口報單號碼 ( 抓發票明細的第一筆收費單號 ) }
    InvAmount: String;      { 發票總金額 }
    ChargeDate: String;     { 輸出或結匯日期 ( 發票主檔得收費日期, 民國年+月+日 ) } 
  end;

  TfrmInvB05 = class(TForm)
    Panel1: TPanel;
    Label1: TLabel;
    Panel2: TPanel;
    Panel3: TPanel;
    PBar: TcxProgressBar;
    btnExec: TcxButton;
    btnClose: TcxButton;
    Bevel1: TBevel;
    Label2: TLabel;
    Label3: TLabel;
    lblCompany: TcxLabel;
    Label4: TLabel;
    lblUser: TcxLabel;
    Label9: TLabel;
    Label6: TLabel;
    rdoApplyType: TcxRadioGroup;
    txtYear: TcxMaskEdit;
    Label5: TLabel;
    cmbMonth: TcxComboBox;
    Label7: TLabel;
    txtFilePath: TcxButtonEdit;
    rdoApplyKind: TcxRadioGroup;
    Label8: TLabel;
    txtFileName: TcxTextEdit;
    Label10: TLabel;
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure txtFilePropertiesButtonClick(Sender: TObject;
      AButtonIndex: Integer);
    procedure btnExecClick(Sender: TObject);
  private
    { Private declarations }
    FZeroTaxCompanyList: TList;
    FRecord: TZeroTaxRecord;
    FApplyYearMonth: String;
    FDataStarDate: String;
    FDataEndDate: String;
    procedure SetButtonCompetence;
    procedure GetZeroTaxInfo;
    function IsDataOk: Boolean;
    function GetSqlText: String;
    function DoTextFile(var aMsg: String): Boolean;
    function GetZeroCompany(const aId: String): PZeorTaxCompany;
  public
    { Public declarations }
  end;

var
  frmInvB05: TfrmInvB05;

implementation

uses cbUtilis, frmMainU, dtmMainU, dtmMainJU, DB;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB05.FormCreate(Sender: TObject);
begin
  FZeroTaxCompanyList := TList.Create;
  GetZeroTaxInfo;
  txtFileName.Text := EmptyStr;
  if ( FZeroTaxCompanyList.Count > 0 ) then
  begin
    { 檔名固定為: 統編.T02 }
    txtFileName.Text :=
      PZeorTaxCompany( FZeroTaxCompanyList[0] ).BusinessId + '.T02';
  end;    
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB05.FormShow(Sender: TObject);
var
  aIndex: Integer;
begin
  Self.Caption := frmMain.GetFormTitleString( 'B05', '零稅率銷售額資料匯檔作業' );
  SetButtonCompetence;
  lblCompany.Caption := Format( '%s ( %s )', [dtmMain.getCompID, dtmMain.getCompName] );
  lblUser.Caption := dtmMain.getLoginUser;
  txtYear.Text := FormatDateTime( 'yyyy', Date );
  cmbMonth.ItemIndex := 0;
  aIndex := cmbMonth.Properties.Items.IndexOf( Copy( MapingInvoiceYearMonth( Date ), 5, 2 ) );
  if ( aIndex >= 0 ) then cmbMonth.ItemIndex := aIndex;
  rdoApplyType.Properties.ReadOnly := ( frmMain.CompanyList.Count <= 1 );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB05.FormDestroy(Sender: TObject);
var
  aIndex: Integer;
begin
  for aIndex := 0 to FZeroTaxCompanyList.Count - 1 do
  begin
    if Assigned( FZeroTaxCompanyList[aIndex] ) then
    begin
      Dispose( PZeorTaxCompany( FZeroTaxCompanyList[aIndex] ) );
      FZeroTaxCompanyList[aIndex] := nil;
    end;
  end;
  FZeroTaxCompanyList.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB05.SetButtonCompetence;
var
  aQuery, aInsert, aDelete, aUpdate, aExecute: String;
begin
  { 設定權限 }
  dtmMain.getCompetence( 'B05', aQuery, aInsert, aDelete, aUpdate, aExecute );
  btnExec.Enabled := ( aExecute = 'Y');
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB05.GetZeroTaxInfo;
var
  aIndex: Integer;
  aReader: TADOQuery;
  aComp: PZeorTaxCompany;
begin
  { 零稅率匯檔所須資料, 發票公司統編,稅籍編號 }
  { 逐一抓出現有設定公司別零稅率參數 }
  aReader := dtmMain.adoComm3;
  for aIndex := 0 to frmMain.CompanyList.Count - 1 do
  begin
    New( aComp );
    aReader.Close;
    aReader.SQL.Text := Format(
      ' select businessid, taxid from inv001  ' +
      '  where identifyid1 = ''%s''           ' +
      '    and identifyid2 = ''%s''           ' +
      '    and compid = ''%s''                ',
      [IDENTIFYID1, IDENTIFYID2,
       PCompany( frmMain.CompanyList[aIndex] ).CompanyId] );
    aReader.Open;
    aComp.CompId := PCompany( frmMain.CompanyList[aIndex] ).CompanyId;
    aComp.BusinessId := aReader.FieldByName( 'businessid' ).AsString;
    aComp.TaxId := aReader.FieldByName( 'taxid' ).AsString;
    aReader.Close;
    { 發票參數檔,零稅率的資料 }
    aReader.SQL.Text := Format(
      ' select city, exportmode, importnote, entryclass  ' +
      '   from inv003                                    ' +
      '  where identifyid1 = ''%s''                      ' +
      '    and identifyid2 = ''%s''                      ' +
      '    and compid = ''%s''                           ',
      [IDENTIFYID1, IDENTIFYID2,
       PCompany( frmMain.CompanyList[aIndex] ).CompanyId] );
    aReader.Open;
    aComp.City := aReader.FieldByName( 'city' ).AsString;
    aComp.ExportMode := aReader.FieldByName( 'exportmode' ).AsString;
    aComp.ImportNote := aReader.FieldByName( 'importnote' ).AsString;
    aComp.EntryClass := aReader.FieldByName( 'entryclass' ).AsString;
    aReader.Close;
    FZeroTaxCompanyList.Add( aComp );
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB05.txtFilePropertiesButtonClick(Sender: TObject;
  AButtonIndex: Integer);
var
  aDir: String;
begin
  aDir := txtFilePath.Text;
  if ( SelectDirectory( '選擇目錄', EmptyWideStr, aDir ) ) then
    txtFilePath.Text := aDir;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvB05.IsDataOk: Boolean;
var
  aPath, aFile: String;
begin
  Result := False;
  if Trim( txtYear.Text ) = EmptyStr then
  begin
    WarningMsg( '請輸入申報年度。' );
    if txtYear.CanFocusEx then txtYear.SetFocus;
    Exit;
  end;
  aPath := txtFilePath.Text;
  aFile := txtFileName.Text;
  if ( aPath = EmptyStr ) then
  begin
    WarningMsg( '請輸入或選取文字檔案路徑。' );
    if txtFilePath.CanFocusEx then txtFilePath.SetFocus;
    Exit;
  end;
  if ( aFile = EmptyStr ) then
  begin
    WarningMsg( '請輸入或選取檔案名稱。' );
    if txtFileName.CanFocusEx then txtFileName.SetFocus;
    Exit;
  end;
  ForceDirectories( aPath );
  if FileExists( IncludeTrailingPathDelimiter( aPath ) + aFile ) then
  begin
    if not ConfirmMsg( '檔案已存在, 是否覆蓋?' ) then Exit;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvB05.GetSqlText: String;
var
  aComps: String;
begin
  aComps := QuotedStr( dtmMain.getCompID );
  if ( rdoApplyType.ItemIndex = 1 ) then
    aComps := QuotedValue( dtmMain.getCompIDEx );
  Result := Format(
   ' select a.invid, a.invdate, a.businessid,    ' +
   '        a.invamount, a.chargedate, a.compid  ' +
   '   from inv007 a                             ' +
   '  where a.identifyid1 = ''%s''               ' +
   '    and a.identifyid2 = ''%s''               ' +
   '    and a.compid in ( %s )                   ' +
   '    and a.isobsolete = ''N''                 ' +
   '    and a.taxtype = ''2''                    ' +
   '    and a.invdate between to_date( ''%s'', ''yyyy/mm/dd'' ) and to_date( ''%s'', ''yyyy/mm/dd'' ) ' +
   '  order by a.invid                           ',
   [IDENTIFYID1, IDENTIFYID2, aComps, FDataStarDate, FDataEndDate] );
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvB05.GetZeroCompany(const aId: String): PZeorTaxCompany;
var
  aIndex: Integer;
begin
  Result := nil;
  for aIndex := 0 to FZeroTaxCompanyList.Count - 1 do
  begin
    if ( PZeorTaxCompany( FZeroTaxCompanyList[aIndex] ).CompId = aId ) then
    begin
      Result := PZeorTaxCompany( FZeroTaxCompanyList[aIndex] );
      Break;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvB05.DoTextFile(var aMsg: String): Boolean;
var
  aReader, aReader2: TADOQuery;
  aComp: PZeorTaxCompany;
  aList: TStringList;
begin
  aMsg := EmptyStr;
  aList := TStringList.Create;
  try
    aReader := dtmMain.adoComm3;
    aReader2 := dtmMain.adoComm2;
    aReader.Close;
    aReader.SQL.Text := GetSqlText;
    try
      aReader.Open;
      PBar.Position := 0;
      PBar.Properties.Min := 0;
      PBar.Properties.Max := aReader.RecordCount;
      aReader.First;
      while not aReader.Eof do
      begin
        ZeroMemory( @FRecord, SizeOf( FRecord ) );
        aComp := GetZeroCompany( aReader.FieldByName( 'compid' ).AsString );
        FRecord.BusinessId := Rpad( aComp.BusinessId, 8, #32 );
        {}
        FRecord.City := Rpad( aComp.City, 1, #32 );
        {}
        FRecord.TaxId := Rpad( aComp.TaxId, 9, #32 );
        {}
        FRecord.YearMonth := FApplyYearMonth;
        {}
        FRecord.InvDate := Lpad( getYearMonthDay7(
          aReader.FieldByName( 'invdate' ).AsDateTime, 1 ), 5, '0' );
        {}
        FRecord.InvId := Rpad( aReader.FieldByName( 'invid' ).AsString, 10, #32 );
        {}
        FRecord.InvBusinessId := Rpad( aReader.FieldByName( 'businessid' ).AsString, 8, #32 );
        {}
        FRecord.ExportMode := Rpad( aComp.ExportMode, 1, #32 );
        {}
        FRecord.ImportNote := Rpad( aComp.ImportNote, 1, #32 );
        {}
        FRecord.EntryClass := Rpad( aComp.EntryClass, 2, #32 );
        { 非出口業務,不需填值,請改填14格空白 }
        FRecord.BillId := Rpad( #32, 14, #32 );
        {}
        FRecord.InvAmount := Lpad( aReader.FieldByName( 'invamount' ).AsString, 12, '0' );
        {}
        FRecord.ChargeDate := Lpad( #32, 7, #32 );
        if not VarIsNull( aReader.FieldByName( 'chargedate' ).Value ) then
        begin
          FRecord.ChargeDate := Lpad( getYearMonthDay7(
            aReader.FieldByName( 'chargedate' ).AsDateTime, 0 ), 7, '0' );
        end;
        {}
        aList.Add( FRecord.BusinessId + FRecord.City + FRecord.TaxId +
          FRecord.YearMonth + FRecord.InvDate + FRecord.InvId +
          FRecord.InvBusinessId + FRecord.ExportMode + FRecord.ImportNote +
          FRecord.EntryClass + FRecord.BillId + FRecord.InvAmount +
          FRecord.ChargeDate ); 
        {}
        aReader.Next;
        PBar.Position := ( PBar.Position + 1 );
        Application.ProcessMessages;
      end;
      aReader.Close;
      {}
      if ( aList.Count > 0 ) then
      begin
        aList.SaveToFile( IncludeTrailingPathDelimiter( txtFilePath.Text ) + txtFileName.Text );
        aMsg := '零稅率資料檔產生完成。';
      end else
      begin
        aMsg := '此申報年月內無零稅率資料可匯檔。';
      end;
      {}
      Result := True;
    except
      on E: Exception do
      begin
        Result := False;
        aMsg := E.Message;
      end;
    end;
  finally
    aList.Free;
  end;  
  PBar.Position := 0;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB05.btnExecClick(Sender: TObject);
var
  aErrMsg: String;
begin
  if not IsDataOk then Exit;
  FApplyYearMonth := txtYear.Text + cmbMonth.Text;
  { 計算申報月份類別 }
  case rdoApplyKind.ItemIndex of
    0://單月
    begin
      FDataStarDate := getYearMonthDay8(FApplyYearMonth, 1 );
      FDataEndDate := getYearMonthDay8(FApplyYearMonth, 2 );
      { 抓單月份 }
      FApplyYearMonth := Format( '%.3d',
        [StrToInt( Copy( FApplyYearMonth, 1, 4 ) ) - 1911] ) +
        Copy( FApplyYearMonth, 5, 2 );
    end;
    1://雙月
    begin
      FDataStarDate :=getYearMonthDay8(FApplyYearMonth, 3 );
      FDataEndDate :=getYearMonthDay8(FApplyYearMonth, 4 );
      { 抓雙月份 }
      FApplyYearMonth := Format( '%.3d',
        [StrToInt( Copy( FApplyYearMonth, 1, 4 ) ) - 1911] ) +
        Format( '%.2d', [StrToInt( Copy( FApplyYearMonth, 5, 2 ) ) + 1] );
    end;
    2://二個月
    begin
      FDataStarDate :=getYearMonthDay8( FApplyYearMonth, 1 );
      FDataEndDate :=getYearMonthDay8( FApplyYearMonth, 4 );
      { 若是為 2 個月, 還是抓雙月份 }
      FApplyYearMonth := Format( '%.3d',
        [StrToInt( Copy( FApplyYearMonth, 1, 4 ) ) - 1911] ) +
        Format( '%.2d', [StrToInt( Copy( FApplyYearMonth, 5, 2 ) ) + 1] );
     end;
  end;
  {}
  btnExec.Enabled := False;
  btnClose.Enabled := False;
  try
    Screen.Cursor := crSQLWait;
    try
      if DoTextFile( aErrMsg ) then
      begin
        InfoMsg( aErrMsg );
      end else
      begin
        ErrorMsg( aErrMsg );
      end;
    finally
      Screen.Cursor := crDefault;
    end;
  finally
    btnExec.Enabled := True;
    btnClose.Enabled := True;
  end;  
end;

{ ---------------------------------------------------------------------------- }

end.
