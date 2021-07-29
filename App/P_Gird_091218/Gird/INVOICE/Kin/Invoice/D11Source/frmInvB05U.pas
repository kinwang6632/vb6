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
  { ���ɥεo�����q�O�ѼƸ�� }
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

  { ���ɥιs�|�v record }
  TZeroTaxRecord = record
    BusinessId: String;     { �o�����q�O�νs }
    City: String;           { �o�����q�O���ݿ����N�X }
    TaxId: String;          { �o�����q�O�|�y�s�� }
    YearMonth: String;      { �ӳ��~�� (����~) }
    InvDate: String;        { �o����� (����~+��) }
    InvId: String;          { �o�����X }
    InvBusinessId: String;  { �R���H�νs (�o���νs) }
    ExportMode: String;     { �~�P�覡 }
    ImportNote: String;     { �q���覡���O }
    EntryClass: String;     { �X�f�������O }
    BillId: String;         { �X�f���渹�X ( ��o�����Ӫ��Ĥ@�����O�渹 ) }
    InvAmount: String;      { �o���`���B }
    ChargeDate: String;     { ��X�ε��פ�� ( �o���D�ɱo���O���, ����~+��+�� ) } 
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
    { �ɦW�T�w��: �νs.T02 }
    txtFileName.Text :=
      PZeorTaxCompany( FZeroTaxCompanyList[0] ).BusinessId + '.T02';
  end;    
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB05.FormShow(Sender: TObject);
var
  aIndex: Integer;
begin
  Self.Caption := frmMain.GetFormTitleString( 'B05', '�s�|�v�P���B��ƶ��ɧ@�~' );
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
  { �]�w�v�� }
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
  { �s�|�v���ɩҶ����, �o�����q�νs,�|�y�s�� }
  { �v�@��X�{���]�w���q�O�s�|�v�Ѽ� }
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
    { �o���Ѽ���,�s�|�v����� }
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
  if ( SelectDirectory( '��ܥؿ�', EmptyWideStr, aDir ) ) then
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
    WarningMsg( '�п�J�ӳ��~�סC' );
    if txtYear.CanFocusEx then txtYear.SetFocus;
    Exit;
  end;
  aPath := txtFilePath.Text;
  aFile := txtFileName.Text;
  if ( aPath = EmptyStr ) then
  begin
    WarningMsg( '�п�J�ο����r�ɮ׸��|�C' );
    if txtFilePath.CanFocusEx then txtFilePath.SetFocus;
    Exit;
  end;
  if ( aFile = EmptyStr ) then
  begin
    WarningMsg( '�п�J�ο���ɮצW�١C' );
    if txtFileName.CanFocusEx then txtFileName.SetFocus;
    Exit;
  end;
  ForceDirectories( aPath );
  if FileExists( IncludeTrailingPathDelimiter( aPath ) + aFile ) then
  begin
    if not ConfirmMsg( '�ɮפw�s�b, �O�_�л\?' ) then Exit;
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
        { �D�X�f�~��,���ݶ��,�Ч��14��ť� }
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
        aMsg := '�s�|�v����ɲ��ͧ����C';
      end else
      begin
        aMsg := '���ӳ��~�뤺�L�s�|�v��ƥi���ɡC';
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
  { �p��ӳ�������O }
  case rdoApplyKind.ItemIndex of
    0://���
    begin
      FDataStarDate := getYearMonthDay8(FApplyYearMonth, 1 );
      FDataEndDate := getYearMonthDay8(FApplyYearMonth, 2 );
      { ����� }
      FApplyYearMonth := Format( '%.3d',
        [StrToInt( Copy( FApplyYearMonth, 1, 4 ) ) - 1911] ) +
        Copy( FApplyYearMonth, 5, 2 );
    end;
    1://����
    begin
      FDataStarDate :=getYearMonthDay8(FApplyYearMonth, 3 );
      FDataEndDate :=getYearMonthDay8(FApplyYearMonth, 4 );
      { ������� }
      FApplyYearMonth := Format( '%.3d',
        [StrToInt( Copy( FApplyYearMonth, 1, 4 ) ) - 1911] ) +
        Format( '%.2d', [StrToInt( Copy( FApplyYearMonth, 5, 2 ) ) + 1] );
    end;
    2://�G�Ӥ�
    begin
      FDataStarDate :=getYearMonthDay8( FApplyYearMonth, 1 );
      FDataEndDate :=getYearMonthDay8( FApplyYearMonth, 4 );
      { �Y�O�� 2 �Ӥ�, �٬O������� }
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
