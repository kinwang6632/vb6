unit frmInvC05U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Buttons, DB,  Menus, dtmMainU,
  cxDropDownEdit, cxControls, cxContainer, cxEdit, cxTextEdit, cxMaskEdit,
  cxRadioGroup, cxGroupBox, cxStyles, cxGraphics, cxFilter, cxClasses,
  cxCheckBox, DBClient, cxTL, cxInplaceContainer, cxCustomData, cxSpinEdit,
  cxProgressBar, frxClass, cxLookAndFeelPainters, cxButtons,
  cxButtonEdit, dxSkinsCore, dxSkinsDefaultPainters;

type

  TReportSummary = record
    { �ϥαi�� }
    HasTax: Integer;      { ���|�i�� }
    ZeroTax: Integer;     { �s�|�v�i��}
    NoTax: Integer;       { �K�|�i�� }
    CutOff: Integer;      { ���o�i�� }
    Empry: Integer;       { �ťձi�� }
    { ��~�H���| }
    SaleAmount: Integer;  { �P���B }
    TaxAmount: Integer;   { �|�B }
    InvAmount: Integer;   { �o���B }
    { �D��~�H���| }
    SaleAmount2: Integer;  { �P���B }
    TaxAmount2: Integer;   { �|�B }
    InvAmount2: Integer;   { �o���B }
    NoTaxSaleAmount: Integer;    { �K�|�P���B ���� }
    NoTaxSaleAmount2: Integer;   { �K�|�P���B ��~�H }
    NoTaxSaleAmount3: Integer;   { �K�|�P���B �D��~�H }
    { �䥦 }
    ZeroTaxSaleAmount: Integer; { �s�|�v�P���B }
    TotalPages: Integer;   { �����`�i�� }
  end;

  TfrmInvC05 = class(TForm)
    Panel1: TPanel;
    Label5: TLabel;
    Panel2: TPanel;
    GroupBox1: TGroupBox;
    Label2: TLabel;
    Label10: TLabel;
    edtYearMonth: TcxMaskEdit;
    edtPrefix: TcxMaskEdit;
    Label7: TLabel;
    cmbFormat: TcxComboBox;
    cxGroupBox1: TcxGroupBox;
    rdo1: TcxRadioButton;
    rdo2: TcxRadioButton;
    Bevel1: TBevel;
    Label1: TLabel;
    InvTree: TcxTreeList;
    TreeColumn1: TcxTreeListColumn;
    TreeColumn2: TcxTreeListColumn;
    TreeColumn3: TcxTreeListColumn;
    TreeColumn4: TcxTreeListColumn;
    TreeColumn5: TcxTreeListColumn;
    btnLoadChooseInv: TcxButton;
    TreeColumn6: TcxTreeListColumn;
    cxGroupBox3: TcxGroupBox;
    chkOnlyPrintSummary: TcxCheckBox;
    cxCheckBox2: TcxCheckBox;
    PBar: TcxProgressBar;
    btnOk: TcxButton;
    btnClose: TcxButton;
    Panel3: TPanel;
    rdoReport: TcxRadioButton;
    rdoFile: TcxRadioButton;
    txtFile: TcxButtonEdit;
    OpenDialog1: TOpenDialog;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure btnCloseClick(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure edtYearMonthPropertiesValidate(Sender: TObject;
      var DisplayValue: Variant; var ErrorText: TCaption;
      var Error: Boolean);
    procedure btnLoadChooseInvClick(Sender: TObject);
    procedure chkOnlyPrintSummaryPropertiesChange(Sender: TObject);
    procedure txtFilePropertiesButtonClick(Sender: TObject;
      AButtonIndex: Integer);
    procedure rdoReportClick(Sender: TObject);
  private
    { Private declarations }
    FYearMonth: String;
    FPrefix: String;
    FFormat: String;
    FReportSummary: TReportSummary;
    FCompTaxId: String;
    FCompBusinessId: String;
    function InputValidate : Boolean;
    function GetInvFormatText: String;
    function CalcTotalRecordCount: Integer;
    procedure LoadChooseInvData;
    procedure UnPrepareDataSet;
    procedure PrepareDataSet;
    procedure BeginReport;
    procedure GetInvoceData(aInvSt, aInvEd: String);
    procedure ResetReportSummary;
    procedure CalcReportSummary;
  public
    { Public declarations }
    procedure DoReport;
    procedure DoTextFile;
  end;

var
  frmInvC05: TfrmInvC05;

implementation

uses frmMainU, dtmMainJU, dtmReportModule, cbUtilis;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC05.FormCreate(Sender: TObject);
begin
  {...}
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC05.FormShow(Sender: TObject);
begin
  Self.Caption := frmMain.GetFormTitleString( 'C05', '�Τ@�o�����Ӫ�' );
  cmbFormat.ItemIndex := 0;
  rdo1.Checked := True;
  edtYearMonth.SetFocus;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC05.btnCloseClick(Sender: TObject);
begin
  Self.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC05.btnOKClick(Sender: TObject);
var
  aIsReport: Boolean;

  { ---------------------------------------------- }

  function GetCompBusinessId: String;
  begin
    dtmMain.adoComm.Close;
    dtmMain.adoComm.SQL.Text := Format(
      ' SELECT BUSINESSID FROM INV001        ' +
      '  WHERE IDENTIFYID1 = ''%s''          ' +
      '    AND IDENTIFYID2 = ''%s''          ' +
      '    AND COMPID = ''%s''               ',
      [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID] );

    dtmMain.adoComm.Open;
    Result := dtmMain.adoComm.FieldByName( 'BUSINESSID' ).AsString;
    dtmMain.adoComm.Close;
  end;

  { ---------------------------------------------- }

  function GetCompTaxId: String;
  begin
    dtmMain.adoComm.Close;
    dtmMain.adoComm.SQL.Text := Format(
      ' SELECT TAXID FROM INV001             ' +
      '  WHERE IDENTIFYID1 = ''%s''          ' +
      '    AND IDENTIFYID2 = ''%s''          ' +
      '    AND COMPID = ''%s''               ',
      [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID] );

    dtmMain.adoComm.Open;
    Result := dtmMain.adoComm.FieldByName( 'TAXID' ).AsString;
    dtmMain.adoComm.Close;
  end;
  
  { ---------------------------------------------- }

begin
  if not InputValidate then Exit;
  aIsReport := rdoReport.Checked;
  Screen.Cursor := crSQLWait;
  try
    UnPrepareDataSet;
    PrepareDataSet;
    PBar.Position := 0;
    PBar.Properties.Min := 0;
    PBar.Properties.Max := CalcTotalRecordCount;
    BeginReport;
    ResetReportSummary;
    CalcReportSummary;
  finally
    Screen.Cursor := crDefault;
  end;
  if dtmReport.cdsTempory.IsEmpty then
  begin
    WarningMsg( '�S���o�����Ӹ�ƥi�ѦC�L!' );
    Exit;
  end;
  {}
  FCompTaxId := GetCompTaxId;
  FCompBusinessId := GetCompBusinessId;
  {}
  cxGroupBox3.Enabled := False;
  try
    if aIsReport then
      DoReport
    else begin
      Screen.Cursor := crSQLWait;
      try
        DoTextFile;
      finally
        Screen.Cursor := crDefault;
      end;
    end;
  finally
    cxGroupBox3.Enabled := True;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvC05.InputValidate: Boolean;
var
  aPath, aFile: String;
begin
  Result := False;
  if ( edtYearMonth.Text = EmptyStr ) then
  begin
    WarningMsg( '�п�J�o���~��C' );
    edtYearMonth.SetFocus;
    Exit;
  end;
  if ( edtPrefix.Text = EmptyStr ) then
  begin
    WarningMsg( '�п�J�o���r�y�C' );
    edtPrefix.SetFocus;
    Exit;
  end;
  if Length( edtPrefix.Text ) <> 2 then
  begin
    WarningMsg( '�п�J���T���r�y�榡�C' );
    edtPrefix.SetFocus;
    Exit;
  end;
  if cmbFormat.ItemIndex < 0 then
  begin
    WarningMsg( '���I��o���榡�C' );
    cmbFormat.SetFocus;
    Exit;
  end;
  if ( rdoFile.Checked ) then
  begin
    if Trim( txtFile.Text ) = EmptyStr then
    begin
      WarningMsg( '�п�J�ο����r�ɮ׸��|���ɮצW�١C' );
      txtFile.SetFocus;
      Exit;
    end;
    aPath := ExtractFilePath( txtFile.Text );
    aFile := ExtractFileName( txtFile.Text );
    if ( aPath = EmptyStr ) then
    begin
      WarningMsg( '�п�J�ο����r�ɮ׸��|�C' );
      txtFile.SetFocus;
      Exit;
    end;
    if ( aFile = EmptyStr ) then
    begin
      WarningMsg( '�п�J�ο���ɮצW�١C' );
      txtFile.SetFocus;
      Exit;
    end;
    ForceDirectories( aPath );
    if FileExists( txtFile.Text ) then
    begin
      if not ConfirmMsg( '�ɮפw�s�b, �O�_�л\?' ) then Exit;
    end;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvC05.GetInvFormatText: String;
var
  aIndex: Integer;
begin
  Result := EmptyStr;
  if cmbFormat.ItemIndex >= 0 then
  begin
    aIndex := Pos( '.', cmbFormat.Text );
    if aIndex > 0 then
      Result := Copy( cmbFormat.Text, 1, aIndex - 1 );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvC05.CalcTotalRecordCount: Integer;
var
  aIndex, aStarNum, aEndNum: Integer;
  aInvSt, aInvEd: String;
begin
  Result := 0;
  for aIndex := 0 to InvTree.Nodes.Count - 1 do
  begin
    if VarToStrDef( InvTree.Nodes[aIndex].Values[TreeColumn1.ItemIndex],
      'N' ) = 'Y' then
    begin
      aStarNum := InvTree.Nodes[aIndex].Values[TreeColumn4.ItemIndex];
      aEndNum := InvTree.Nodes[aIndex].Values[TreeColumn5.ItemIndex];
      Result := ( Result + ( aEndNum - aStarNum + 1 ) );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC05.LoadChooseInvData;
var
  aRecNo: Integer;
  aNode: TcxTreeListNode;
begin
  dtmMain.adoComm.Close;
  dtmMain.adoComm.SQL.Text := Format(
    ' SELECT A.PREFIX, A.STARTNUM, A.ENDNUM, A.CURNUM ' +
    '   FROM INV099 A                       ' +
    '  WHERE A.IDENTIFYID1 = ''%s''         ' +
    '    AND A.IDENTIFYID2 = ''%s''         ' +
    '    AND A.COMPID = ''%s''              ' +
    '    AND A.YEARMONTH = ''%s''           ' +
    '    AND A.INVFORMAT = ''%s''           ' +
    '    AND A.PREFIX = ''%s''              ' +
    '  ORDER BY A.STARTNUM                  ',
    [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, FYearMonth, FFormat, FPrefix] );
  dtmMain.adoComm.Open;
  InvTree.Clear;
  aRecNo := 0;
  while not dtmMain.adoComm.Eof do
  begin
    Inc( aRecNo );
    aNode := InvTree.Add;
    aNode.Values[TreeColumn1.ItemIndex] := 'Y';
    aNode.Values[TreeColumn2.ItemIndex] := aRecNo;
    aNode.Values[TreeColumn3.ItemIndex] :=
      dtmMain.adoComm.FieldByName( 'PREFIX' ).AsString;
    aNode.Values[TreeColumn4.ItemIndex] :=
      dtmMain.adoComm.FieldByName( 'STARTNUM' ).AsString;
    aNode.Values[TreeColumn5.ItemIndex] :=
      dtmMain.adoComm.FieldByName( 'ENDNUM' ).AsString;
    aNode.Values[TreeColumn6.ItemIndex] :=
      dtmMain.adoComm.FieldByName( 'CURNUM' ).AsString;
    dtmMain.adoComm.Next;
  end;
  dtmMain.adoComm.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC05.edtYearMonthPropertiesValidate(Sender: TObject;
  var DisplayValue: Variant; var ErrorText: TCaption; var Error: Boolean);
begin
  if Error then
    ErrorText := '��J���o���~�뤣���T';
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC05.btnLoadChooseInvClick(Sender: TObject);
var
  aMonth: Integer;
begin
  if InputValidate then
  begin
    Screen.Cursor := crSQLWait;
    try
      FYearMonth := StringReplace( PadDateText( edtYearMonth.Text ), '/',
        EmptyStr, [rfReplaceAll] );
      FYearMonth := Copy( FYearMonth, 1, 6 );
      aMonth := StrToInt( Copy( FYearMonth, 5, 2 ) );
      if ( aMonth mod 2 ) = 0 then
      begin
        FYearMonth := Copy( FYearMonth, 1, 4 ) + Format( '%.2d', [aMonth-1] );
      end;
      FPrefix := edtPrefix.Text;
      FFormat := GetInvFormatText;
      LoadChooseInvData;
    finally
      Screen.Cursor := crDefault;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC05.chkOnlyPrintSummaryPropertiesChange(Sender: TObject);
begin
  cxCheckBox2.Enabled := not chkOnlyPrintSummary.Checked;
  if not cxCheckBox2.Enabled then cxCheckBox2.Checked := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC05.PrepareDataSet;
begin
  dtmReport.cdsTempory.FieldDefs.Clear;
  dtmReport.cdsTempory.FieldDefs.Add( 'PAGE', ftInteger );
  dtmReport.cdsTempory.FieldDefs.Add( 'RECNO', ftInteger );
  dtmReport.cdsTempory.FieldDefs.Add( 'INVID', ftString, 10 );
  dtmReport.cdsTempory.FieldDefs.Add( 'INVDATE', ftDateTime );
  dtmReport.cdsTempory.FieldDefs.Add( 'BUSINESSID', ftString, 8 );
  dtmReport.cdsTempory.FieldDefs.Add( 'INVFORMAT', ftString, 1 );
  dtmReport.cdsTempory.FieldDefs.Add( 'TAXTYPE', ftString, 1 );
  dtmReport.cdsTempory.FieldDefs.Add( 'SALEAMOUNT', ftInteger );
  dtmReport.cdsTempory.FieldDefs.Add( 'TAXAMOUNT', ftInteger );
  dtmReport.cdsTempory.FieldDefs.Add( 'INVAMOUNT', ftInteger );
  dtmReport.cdsTempory.FieldDefs.Add( 'ISOBSOLETE', ftString, 1 );
  {}
  dtmReport.cdsTempory.CreateDataSet;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC05.UnPrepareDataSet;
begin
  if not VarIsNull( dtmReport.cdsTempory.Data ) then
    dtmReport.cdsTempory.Data := Null;
  dtmReport.cdsTempory.FieldDefs.Clear;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC05.BeginReport;
var
  aIndex: Integer;
begin
  for aIndex := 0 to InvTree.Nodes.Count - 1 do
  begin
    { ��� }
    if VarToStrDef( InvTree.Nodes[aIndex].Values[TreeColumn1.ItemIndex],
      EmptyStr ) = 'Y' then
    begin
      GetInvoceData( VarToStrDef(
        InvTree.Nodes[aIndex].Values[TreeColumn4.ItemIndex], EmptyStr ),
                     VarToStrDef(
        InvTree.Nodes[aIndex].Values[TreeColumn5.ItemIndex], EmptyStr ) );
    end;    
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC05.GetInvoceData(aInvSt, aInvEd: String);

  { ------------------------------------ }

  function GetLastPage: Integer;
  begin
    Result := 1;
    if not dtmReport.cdsTempory.IsEmpty then
    begin
      dtmReport.cdsTempory.Last;
      Result := dtmReport.cdsTempory.FieldByName( 'PAGE' ).AsInteger + 1;
    end;
  end;

  { ------------------------------------ }

var
  aIndex, aPage, aCompareInvId, aMark, aRecNo, aEnd: Integer;
begin
  dtmMain.adoComm.Close;
  dtmMain.adoComm.SQL.Text := Format(
    ' SELECT A.INVID,                ' +
    '        A.INVDATE,              ' +
    '        A.BUSINESSID,           ' +
    '        A.INVFORMAT,            ' +
    '        A.TAXTYPE,              ' +
    '        A.SALEAMOUNT,           ' +
    '        A.TAXAMOUNT,            ' +
    '        A.INVAMOUNT,            ' +
    '        A.ISOBSOLETE            ' +
    '   FROM INV007 A                ' +
    '  WHERE A.IDENTIFYID1 = ''%s''  ' +
    '    AND A.IDENTIFYID2 = ''%s''  ' +
    '    AND A.COMPID = ''%s''       ' +
    '    AND A.INVFORMAT = ''%s''    ' +
    '    AND A.INVID BETWEEN ''%s'' AND ''%s''  ' +
    '   ORDER BY INVID                ',
    [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, GetInvFormatText,
     ( FPrefix + aInvSt ), ( FPrefix + aInvEd )] );
  dtmMain.adoComm.Open;
  dtmMain.adoComm.First;
  aRecNo := 0;
  aPage := GetLastPage;
  aMark := StrToInt( aInvSt );
  aEnd :=  StrToInt( aInvEd );
  while not dtmMain.adoComm.Eof do
  begin
    { ��ڪ��o���X�P�o�������_�l���i�঳�t��, �Y�n�L����, �h�o�����P��ڵo��
      �e���Τ����S�ιL���ťյo�����X, �]�����L�X�� }
    aCompareInvId := StrToInt( Copy(
      dtmMain.adoComm.FieldByName( 'INVID' ).AsString, 3, Length(
      dtmMain.adoComm.FieldByName( 'INVID' ).AsString ) - 2 ) );
    for aIndex := 1 to ( aCompareInvId - aMark ) do
    begin
      Inc( aRecNo );
      if aRecNo > 50 then
      begin
        Inc( aPage );
        aRecNo := 1;
      end;
      dtmReport.cdsTempory.AppendRecord( [
        aPage,
        aRecNo,
        FPrefix + IntToStr( aMark ),
        EmptyStr,
        EmptyStr,
        EmptyStr,
        EmptyStr,
        0,
        0,
        0,
        'S'] );
      Inc( aMark );
      PBar.Position := ( PBar.Position + 1 );
      Application.ProcessMessages;
    end;
    Inc( aRecNo );
    if aRecNo > 50 then
    begin
      Inc( aPage );
      aRecNo := 1;
    end;
    dtmReport.cdsTempory.AppendRecord( [
      aPage,
      aRecNo,
      dtmMain.adoComm.FieldByName( 'INVID' ).AsString,
      dtmMain.adoComm.FieldByName( 'INVDATE' ).AsDateTime,
      dtmMain.adoComm.FieldByName( 'BUSINESSID' ).AsString,
      dtmMain.adoComm.FieldByName( 'INVFORMAT' ).AsString,
      dtmMain.adoComm.FieldByName( 'TAXTYPE' ).AsString,
      dtmMain.adoComm.FieldByName( 'SALEAMOUNT' ).AsString,
      dtmMain.adoComm.FieldByName( 'TAXAMOUNT' ).AsString,
      dtmMain.adoComm.FieldByName( 'INVAMOUNT' ).AsString,
      dtmMain.adoComm.FieldByName( 'ISOBSOLETE' ).AsString] );
    Inc( aMark );
    Application.ProcessMessages;
    dtmMain.adoComm.Next;
    PBar.Position := ( PBar.Position + 1 );
    Application.ProcessMessages;
  end;
  { �B�z����, �ݬݬO�_�ٶ��n�[�L�ťզC }
  if dtmMain.adoComm.IsEmpty then
    aCompareInvId := StrToInt( aInvSt ) - 1;
  for aIndex := 1 to ( aEnd - aCompareInvId ) do
  begin
    Inc( aRecNo );
    if aRecNo > 50 then
    begin
      Inc( aPage );
      aRecNo := 1;
    end;
    Inc( aCompareInvId );
    dtmReport.cdsTempory.AppendRecord( [
      aPage,
      aRecNo,
      FPrefix + IntToStr( aCompareInvId ),
      EmptyStr,
      EmptyStr,
      EmptyStr,
      EmptyStr,
      0,
      0,
      0,
      'S'] );
    PBar.Position := ( PBar.Position + 1 );
    Application.ProcessMessages;
  end;
  dtmMain.adoComm.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC05.ResetReportSummary;
begin
  FReportSummary.HasTax := 0;
  FReportSummary.ZeroTax := 0;
  FReportSummary.NoTax := 0;
  FReportSummary.CutOff := 0;
  FReportSummary.Empry := 0;
  FReportSummary.SaleAmount := 0;
  FReportSummary.TaxAmount := 0;
  FReportSummary.InvAmount := 0;
  FReportSummary.SaleAmount2 := 0;
  FReportSummary.TaxAmount2 := 0;
  FReportSummary.InvAmount2 := 0;
  FReportSummary.ZeroTaxSaleAmount := 0;
  FReportSummary.NoTaxSaleAmount := 0;
  FReportSummary.NoTaxSaleAmount2 := 0;
  FReportSummary.NoTaxSaleAmount3 := 0;
  FReportSummary.TotalPages := 0;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC05.CalcReportSummary;
var
  aPrePage, aCurPage, aIndex, aDeletedPage: Integer;
  aIsBlankPage, aFirstEnter: Boolean;
  aBookMark: TBookmark;
begin
  if dtmReport.cdsTempory.IsEmpty then Exit;
  dtmReport.cdsTempory.First;
  aFirstEnter := True;
  aDeletedPage := 0; 
  while not dtmReport.cdsTempory.Eof do
  begin
    if aFirstEnter then
    begin
      aBookMark := dtmReport.cdsTempory.GetBookmark;
      aPrePage := dtmReport.cdsTempory.FieldByName( 'PAGE' ).AsInteger;
      aFirstEnter := False;
      aIsBlankPage := True;
    end;
    aCurPage := dtmReport.cdsTempory.FieldByName( 'PAGE' ).AsInteger;
    if ( aCurPage <> aPrePage ) then
    begin
      if aIsBlankPage then
      begin
        if cxCheckBox2.Checked then
        begin
          dtmReport.cdsTempory.GotoBookmark( aBookMark );
          dtmReport.cdsTempory.FreeBookmark( aBookMark );
          for aIndex := 1 to 50 do
            dtmReport.cdsTempory.Delete;
          Inc( aDeletedPage );  
        end;    
      end;
      aBookMark := dtmReport.cdsTempory.GetBookmark;  
      aPrePage := aCurPage;
      aIsBlankPage := True;
    end else
    begin
      { ISOBSOLETE = 'S' = �ťե��ϥΪ��o�� }
      if ( dtmReport.cdsTempory.FieldByName( 'ISOBSOLETE' ).AsString <> 'S' ) then
        aIsBlankPage := False;
      { ���`�o�� }
      if ( dtmReport.cdsTempory.FieldByName( 'ISOBSOLETE' ).AsString = 'N' ) then
      begin
        if ( dtmReport.cdsTempory.FieldByName( 'TAXTYPE' ).AsString = '1' ) then
          FReportSummary.HasTax := ( FReportSummary.HasTax + 1 );
        {}
        if ( dtmReport.cdsTempory.FieldByName( 'TAXTYPE' ).AsString = '2' ) then
          FReportSummary.ZeroTax := ( FReportSummary.ZeroTax + 1 );
        {}
        if ( dtmReport.cdsTempory.FieldByName( 'TAXTYPE' ).AsString = '3' ) then
          FReportSummary.NoTax := ( FReportSummary.NoTax + 1 );
        { �s�|�v�P���B }
        if ( dtmReport.cdsTempory.FieldByName( 'TAXTYPE' ).AsString = '2' ) then
        begin
          FReportSummary.ZeroTaxSaleAmount := ( FReportSummary.ZeroTaxSaleAmount +
            dtmReport.cdsTempory.FieldByName( 'SALEAMOUNT' ).AsInteger );
        end else
        { �K�|�P���B, �@�ˤ� ��~�H, �D��~�H }
        if ( dtmReport.cdsTempory.FieldByName( 'TAXTYPE' ).AsString = '3' ) then
        begin
          FReportSummary.NoTaxSaleAmount := ( FReportSummary.NoTaxSaleAmount +
              dtmReport.cdsTempory.FieldByName( 'SALEAMOUNT' ).AsInteger );
          {}
          if ( dtmReport.cdsTempory.FieldByName( 'BUSINESSID' ).AsString <> EmptyStr ) then
          begin
            FReportSummary.NoTaxSaleAmount2 := ( FReportSummary.NoTaxSaleAmount2 +
              dtmReport.cdsTempory.FieldByName( 'SALEAMOUNT' ).AsInteger );
          end else
          begin
            FReportSummary.NoTaxSaleAmount3 := ( FReportSummary.NoTaxSaleAmount3 +
              dtmReport.cdsTempory.FieldByName( 'SALEAMOUNT' ).AsInteger );
          end;
        end;
        { ��~�H���| }
        if ( dtmReport.cdsTempory.FieldByName( 'BUSINESSID' ).AsString <> EmptyStr ) then
        begin
          if ( dtmReport.cdsTempory.FieldByName( 'TAXTYPE' ).AsString = '1' ) then
          begin
            FReportSummary.SaleAmount := ( FReportSummary.SaleAmount +
              dtmReport.cdsTempory.FieldByName( 'SALEAMOUNT' ).AsInteger );
            FReportSummary.TaxAmount := ( FReportSummary.TaxAmount +
              dtmReport.cdsTempory.FieldByName( 'TAXAMOUNT' ).AsInteger );
            FReportSummary.InvAmount := ( FReportSummary.InvAmount +
              dtmReport.cdsTempory.FieldByName( 'INVAMOUNT' ).AsInteger );
          end;    
        end else
        { �D��~�H���| }
        begin
          if ( dtmReport.cdsTempory.FieldByName( 'TAXTYPE' ).AsString = '1' ) then
          begin
            FReportSummary.SaleAmount2 := ( FReportSummary.SaleAmount2 +
              dtmReport.cdsTempory.FieldByName( 'SALEAMOUNT' ).AsInteger );
            FReportSummary.TaxAmount2 := ( FReportSummary.TaxAmount2 +

              dtmReport.cdsTempory.FieldByName( 'TAXAMOUNT' ).AsInteger );
            FReportSummary.InvAmount2 := ( FReportSummary.InvAmount2 +
              dtmReport.cdsTempory.FieldByName( 'INVAMOUNT' ).AsInteger );
          end;    
        end;
      end else
      { ���o�o�� }
      if ( dtmReport.cdsTempory.FieldByName( 'ISOBSOLETE' ).AsString = 'Y' ) then
      begin
        FReportSummary.CutOff := ( FReportSummary.CutOff + 1 );
      end else
      { �ťյo�� }
      if ( dtmReport.cdsTempory.FieldByName( 'ISOBSOLETE' ).AsString = 'S' ) then
      begin
        FReportSummary.Empry := ( FReportSummary.Empry + 1 );
      end;
      dtmReport.cdsTempory.Next;
      if ( dtmReport.cdsTempory.Eof ) and ( aIsBlankPage ) then
      begin
        if cxCheckBox2.Checked then
        begin
          dtmReport.cdsTempory.GotoBookmark( aBookMark );
          dtmReport.cdsTempory.FreeBookmark( aBookMark );
          for aIndex := 1 to 50 do
            dtmReport.cdsTempory.Delete;
          Inc( aDeletedPage );  
          dtmReport.cdsTempory.Next;
        end;  
      end;
    end;
  end;
  FReportSummary.TotalPages := ( aCurPage - aDeletedPage );
  if ( rdo2.Checked ) then
  begin
    FReportSummary.SaleAmount2 := Round( FReportSummary.InvAmount2 / 1.05 );
    FReportSummary.TaxAmount2 := Round( FReportSummary.SaleAmount2 * 0.05 );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC05.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  UnPrepareDataSet;
  dtmMain.adoComm.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC05.DoReport;

  { ---------------------------------------------- }

  function GetYearMonthText(const aInvYearMonth: String): String;
  var
    aYear, aMonth: Integer;
    aTemp1, aTemp2: String;
  begin
    aYear := StrToInt( Copy( aInvYearMonth, 1, 4 ) ) - 1911;
    aMonth := StrToInt( Copy( aInvYearMonth, 5, 2 ) );
    if ( aMonth mod 2 ) = 0 then
    begin
      aTemp1 := Format( '%d�~%d', [aYear, aMonth-1] );
      aTemp2 := Format( '%d��', [aMonth] );
    end else
    begin
      aTemp1 := Format( '%d�~%d', [aYear, aMonth] );
      aTemp2 := Format( '%d��', [aMonth+1] );
    end;
    Result := ( aTemp1 + '-' + aTemp2 );
  end;

  { ---------------------------------------------- }
  
var
  aPath, aText: String;
begin
  aPath := IncludeTrailingPathDelimiter( ExtractFilePath(
    Application.ExeName ) ) + IncludeTrailingPathDelimiter(
    REPORT_FOLDER ) + 'FrptInvC05_1.fr3';
  {}
  dtmReport.frxMainReport.LoadFromFile( aPath );
  {}
  aText := dtmMain.getCompLongName;
  dtmReport.frxMainReport.Variables.Variables['CompName'] :=
    QuotedStr( aText );
  {}
  aText := GetYearMonthText( FYearMonth );
  dtmReport.frxMainReport.Variables.Variables['�o�����'] :=
    QuotedStr( aText );
  {}
  if FFormat = '1' then
    aText := '�q�l�p���'
  else if FFormat = '2' then
    aText :=  '�G�p��'
  else if FFormat = '3' then
    aText :=  '�T�p��'
  else
    aText := FFormat;
  dtmReport.frxMainReport.Variables.Variables['�o���榡'] :=
    QuotedStr( aText );
  {}
  aText := FCompBusinessId;
  dtmReport.frxMainReport.Variables.Variables['CompBusinessId'] :=
    QuotedStr( aText );
  {}
  aText := FCompTaxId;
  dtmReport.frxMainReport.Variables.Variables['CompTaxId'] :=
    QuotedStr( aText );
  {}
  if rdo1.Checked then
    aText := '0'
  else
    aText := '1';
  dtmReport.frxMainReport.Variables.Variables['Combine'] :=
    QuotedStr( aText );
  {}
  dtmReport.frxMainReport.Variables.Variables['MyTotalPage'] :=
    QuotedStr( IntToStr( dtmReport.cdsTempory.RecordCount div 50 + 1 ) );
  {}
  dtmReport.frxMainReport.Variables.Variables['TaxCount'] :=
    QuotedStr( IntToStr( FReportSummary.HasTax ) );
  {}
  dtmReport.frxMainReport.Variables.Variables['ZeroTaxCount'] :=
    QuotedStr( IntToStr( FReportSummary.ZeroTax ) );
  {}
  dtmReport.frxMainReport.Variables.Variables['NoTaxCount'] :=
    QuotedStr( IntToStr( FReportSummary.NoTax ) );
  {}
  dtmReport.frxMainReport.Variables.Variables['CutoffCount'] :=
    QuotedStr( IntToStr( FReportSummary.CutOff ) );
  {}
  dtmReport.frxMainReport.Variables.Variables['EmptyCount'] :=
    QuotedStr( IntToStr( FReportSummary.Empry ) );
  {}
  aText := 'N';
  if chkOnlyPrintSummary.Checked then aText := 'Y';
  dtmReport.frxMainReport.Variables.Variables['OnlyPrintSummary'] :=
    QuotedStr( aText );
  {}
  dtmReport.frxMainReport.Variables.Variables['��~�H�P���B'] :=
    QuotedStr( IntToStr( FReportSummary.SaleAmount ) );
  {}
  dtmReport.frxMainReport.Variables.Variables['��~�H�|�B'] :=
    QuotedStr( IntToStr( FReportSummary.TaxAmount ) );
  {}
  dtmReport.frxMainReport.Variables.Variables['��~�H�o���B'] :=
    QuotedStr( IntToStr( FReportSummary.InvAmount ) );
  {}
  dtmReport.frxMainReport.Variables.Variables['�D��~�H�P���B'] :=
    QuotedStr( IntToStr( FReportSummary.SaleAmount2 ) );
  {}
  dtmReport.frxMainReport.Variables.Variables['�D��~�H�|�B'] :=
    QuotedStr( IntToStr( FReportSummary.TaxAmount2 ) );
  {}
  dtmReport.frxMainReport.Variables.Variables['�D��~�H�o���B'] :=
    QuotedStr( IntToStr( FReportSummary.InvAmount2 ) );
  {}
  dtmReport.frxMainReport.Variables.Variables['�s�|�v�P���B'] :=
    QuotedStr( IntToStr( FReportSummary.ZeroTaxSaleAmount ) );
  {}
  dtmReport.frxMainReport.Variables.Variables['�K�|�P���B'] :=
    QuotedStr( IntToStr( FReportSummary.NoTaxSaleAmount ) );
  {}
  dtmReport.frxMainReport.Variables.Variables['��~�H�K�|�P���B'] :=
    QuotedStr( IntToStr( FReportSummary.NoTaxSaleAmount2 ) );
  {}
  dtmReport.frxMainReport.Variables.Variables['�D��~�H�K�|�P���B'] :=
    QuotedStr( IntToStr( FReportSummary.NoTaxSaleAmount3 ) );
  {}
  dtmReport.frxMainReport.Variables.Variables['�`����'] :=
    QuotedStr( IntToStr( FReportSummary.TotalPages ) );
  {}
  dtmReport.cdsTempory.First;
  if chkOnlyPrintSummary.Checked then
  begin
    dtmReport.cdsTempory.EmptyDataSet;
    TfrxReportPage( dtmReport.frxMainReport.Pages[0] ).Visible := False;
  end;

  dtmReport.frxMainReport.ShowReport;

end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC05.DoTextFile;
var
  aFile: TStringList;
  aInvId, aInvDate, aInvFormat, aInvTaxType: String;
  aInvBusinessId, aIsObsolete, aSaleAmount, aTaxAmount, aInvAmount: String;
  aPoBusId, aPoTaxId: String;
begin
  PBar.Position := 0;
  PBar.Properties.Min := 0;
  PBar.Properties.Max := dtmReport.cdsTempory.RecordCount;
  Application.ProcessMessages;
  aFile := TStringList.Create;
  try
    dtmReport.cdsTempory.First;
    while not dtmReport.cdsTempory.Eof do
    begin
      aIsObsolete := dtmReport.cdsTempory.FieldByName( 'ISOBSOLETE' ).AsString;
      { aIsObsolete = 'Y' �@�o, aIsObsolete = 'N' ���@�o, aIsObsolete = 'S' �ťե��ϥ� }
      if ( aIsObsolete <> 'S' ) then
      begin
        aInvId := dtmReport.cdsTempory.FieldByName( 'INVID' ).AsString;
        aInvDate := FormatDateTime( 'yyyymmdd', dtmReport.cdsTempory.FieldByName( 'INVDATE' ).AsDateTime );
        aInvFormat := dtmReport.cdsTempory.FieldByName( 'INVFORMAT' ).AsString;
        aInvTaxType := dtmReport.cdsTempory.FieldByName( 'TAXTYPE' ).AsString;
        aInvBusinessId := dtmReport.cdsTempory.FieldByName( 'BUSINESSID' ).AsString;
        aSaleAmount := dtmReport.cdsTempory.FieldByName( 'SALEAMOUNT' ).AsString;
        aTaxAmount := dtmReport.cdsTempory.FieldByName( 'TAXAMOUNT' ).AsString;
        aInvAmount := dtmReport.cdsTempory.FieldByName( 'INVAMOUNT' ).AsString;
        { �p�G �D��~�H�|�B�p��� �X�� }
        if ( rdo2.Checked ) and ( aInvBusinessId = EmptyStr ) then
        begin
          { �h�P���B = �o���`���B }
          aSaleAmount := aInvAmount;
          aTaxAmount := EmptyStr;
        end;
        aPoBusId := FCompBusinessId;
        aPoTaxId := FCompTaxId;
        aFile.Add(
          Rpad( aInvId, 10, #32 ) +
          Rpad( aInvDate, 8, #32 ) +
          Rpad( aInvFormat, 1, #32 ) +
          Rpad( aInvTaxType, 1, #32 ) +
          Rpad( aInvBusinessId, 8, #32 ) +
          Rpad( aIsObsolete, 1, #32 ) +
          Lpad( aSaleAmount, 8, '0' ) +
          Lpad( aTaxAmount, 8, '0' ) +
          Lpad( aInvAmount, 8, '0' ) +
          Rpad( aPoBusId, 8, #32 ) +
          Rpad( aPoTaxId, 9, #32 ) );
      end;    
      dtmReport.cdsTempory.Next;
      PBar.Position := ( PBar.Position + 1 );
      Application.ProcessMessages; 
    end;
    if ( aFile.Count > 0 ) then
    begin
      aFile.SaveToFile( txtFile.Text );
      InfoMsg( '��r����X�����C' );
    end;
  finally
    aFile.Free;
  end;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC05.txtFilePropertiesButtonClick(Sender: TObject;
  AButtonIndex: Integer);
begin
  if OpenDialog1.Execute then
    txtFile.Text := OpenDialog1.FileName;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC05.rdoReportClick(Sender: TObject);
begin
  if rdoReport.Checked then
  begin
    chkOnlyPrintSummary.Enabled := True;
    cxCheckBox2.Enabled := True;
  end else
  begin
    chkOnlyPrintSummary.Enabled := False;
    cxCheckBox2.Enabled := False;
    chkOnlyPrintSummary.Checked := False;
    cxCheckBox2.Checked := False;
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
