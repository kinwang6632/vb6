unit frmInvC03U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls,  DB, DBClient, XLSFile, fraYMDU, Mask, Buttons,
  cbUtilis,
  cxTextEdit, cxRadioGroup, cxControls, cxContainer, cxEdit, cxGroupBox,
  cxMaskEdit, cxDropDownEdit, cxCalendar, Menus, cxLookAndFeelPainters,
  cxButtons, cxButtonEdit, cxProgressBar, dxSkinsCore, dxSkinsDefaultPainters;

type
  TfrmInvC03 = class(TForm)
    Panel1: TPanel;
    Label6: TLabel;
    Panel2: TPanel;
    rgpCondition: TcxGroupBox;
    rdoBillNo: TcxRadioButton;
    rdoCustId: TcxRadioButton;
    rdoChargeDate: TcxRadioButton;
    rdoLogDate: TcxRadioButton;
    txtBillNo: TcxTextEdit;
    txtCustId: TcxTextEdit;
    txtChageDateSt: TcxDateEdit;
    Label5: TLabel;
    txtChageDateEd: TcxDateEdit;
    txtLogDateSt: TcxDateEdit;
    btnOk: TcxButton;
    btnClose: TcxButton;
    txtLogDateEd: TcxDateEdit;
    Label7: TLabel;
    cxGroupBox3: TcxGroupBox;
    Panel3: TPanel;
    rdoReport: TcxRadioButton;
    rdoExcel: TcxRadioButton;
    txtFile: TcxButtonEdit;
    PBar: TcxProgressBar;
    OpenDialog1: TOpenDialog;
    procedure btnExitClick(Sender: TObject);
    procedure btnQueryClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure txtFilePropertiesButtonClick(Sender: TObject;
      AButtonIndex: Integer);
    procedure txtBillNoEnter(Sender: TObject);
    procedure txtCustIdEnter(Sender: TObject);
    procedure txtChageDateStEnter(Sender: TObject);
    procedure txtLogDateStEnter(Sender: TObject);
    procedure txtLogDateStExit(Sender: TObject);
    procedure txtChageDateStExit(Sender: TObject); //2004/11 Julien Test
  private
    { Private declarations }
    FDataBuffer: TClientDataSet;
    procedure PrepareDataBuffer;
    procedure UnPrepareDataBuffer;
    function IsDataOK : Boolean;
    function GetConditionType: Integer;
  public
    { Public declarations }
    function DoExportToExcel(var aMsg: String; aCondition: String): Boolean;
    function DoRunReport(var aMsg: String; aSql, aCondition: String): Boolean;
  end;

var
  frmInvC03: TfrmInvC03;

implementation

uses dtmMainU, dtmMainJU, dtmReportModule;


{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC03.FormCreate(Sender: TObject);
begin
  rdoBillNo.Checked := True;
  Self.ActiveControl := txtBillNo;
  FDataBuffer := TClientDataSet.Create( nil );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC03.FormDestroy(Sender: TObject);
begin
  FDataBuffer.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC03.btnExitClick(Sender: TObject);
begin
  Close;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvC03.GetConditionType: Integer;
begin
  if ( rdoBillNo.Checked ) then
    Result := CONDITION_TYPE_BILLNO
  else if ( rdoCustId.Checked ) then
    Result := CONDITION_TYPE_CUSTID
  else if ( rdoChargeDate.Checked ) then
    Result := CONDITION_TYPE_CHARGEDATE
  else if ( rdoLogDate.Checked ) then
    Result := CONDITION_TYPE_LOGDATE;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC03.PrepareDataBuffer;
begin
  if ( FDataBuffer.FieldDefs.Count <= 0 ) then
  begin
    FDataBuffer.FieldDefs.Add( 'BILLID', ftString, 15 );
    FDataBuffer.FieldDefs.Add( 'CUSTID', ftString, 8 );
    FDataBuffer.FieldDefs.Add( 'CUSTNAME', ftString, 50 );
    FDataBuffer.FieldDefs.Add( 'DESCRIPTION', ftString, 40 );
    FDataBuffer.FieldDefs.Add( 'TOTALAMOUNT', ftString, 20 );
    FDataBuffer.FieldDefs.Add( 'CHARGEDATE', ftString, 10 );
    FDataBuffer.FieldDefs.Add( 'LOGTIME', ftString, 20 );
    FDataBuffer.FieldDefs.Add( 'NOTES', ftString, 255 );
    FDataBuffer.CreateDataSet;
  end;
  FDataBuffer.FieldByName( 'BILLID' ).DisplayLabel := '���O�渹';
  FDataBuffer.FieldByName( 'CUSTID' ).DisplayLabel := '�Ƚs';
  FDataBuffer.FieldByName( 'CUSTNAME' ).DisplayLabel := '�W��';
  FDataBuffer.FieldByName( 'DESCRIPTION' ).DisplayLabel := '���O����';
  FDataBuffer.FieldByName( 'TOTALAMOUNT' ).DisplayLabel := '�`���B';
  FDataBuffer.FieldByName( 'CHARGEDATE' ).DisplayLabel := '���O���';
  FDataBuffer.FieldByName( 'LOGTIME' ).DisplayLabel := 'Log �ɶ�';
  FDataBuffer.FieldByName( 'NOTES' ).DisplayLabel := '�Ƶ�';
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC03.UnPrepareDataBuffer;
begin
  if not VarIsNull( FDataBuffer.Data ) then
    FDataBuffer.EmptyDataSet;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvC03.IsDataOK: Boolean;
var
  aLen: Integer;
begin
  Result := False;
  { ��ڽs�� }
  if ( rdoBillNo.Checked ) then
  begin
    aLen := Length( Trim( txtBillNo.Text ) );
    if ( aLen <= 0 ) then
    begin
      WarningMsg( '�п�J��ڽs��,�B�s�����ץ�����15�X�C' );
      if txtBillNo.CanFocus then txtBillNo.SetFocus;
      Exit;
    end;
  end;
  { �Ƚs }
  if ( rdoCustId.Checked ) then
  begin
    aLen := Length( Trim( txtCustId.Text ) );
    if ( aLen <= 0 ) then
    begin
      WarningMsg( '�п�J�Ƚs�C' );
      if txtCustId.CanFocus then txtCustId.SetFocus;
      Exit;
    end;
  end;
  { ���O��� }
  if ( rdoChargeDate.Checked ) then
  begin
    if ( txtChageDateSt.Text = EmptyStr ) then
    begin
      WarningMsg( '�п�J���O���(��)�C' );
      if txtChageDateSt.CanFocus then txtChageDateSt.SetFocus;
      Exit;
    end;
    if ( txtChageDateEd.Text = EmptyStr ) then
    begin
      WarningMsg( '�п�J���O���(��)�C' );
      if txtChageDateEd.CanFocus then txtChageDateEd.SetFocus;
      Exit;
    end;
  end;
  { Log ��� }
  if ( rdoLogDate.Checked ) then
  begin
    if ( txtLogDateSt.Text = EmptyStr ) then
    begin
      WarningMsg( '�п�JLog���(�_)�C' );
      if txtLogDateSt.CanFocus then txtLogDateSt.SetFocus;
      Exit;
    end;
    if ( txtLogDateEd.Text = EmptyStr ) then
    begin
      WarningMsg( '�п�JLog���(��)�C' );
      if txtLogDateEd.CanFocus then txtLogDateEd.SetFocus;
      Exit;
    end;
  end;
  { �ץX Excel }
  if ( rdoExcel.Checked ) and ( Trim( txtFile.Text ) = EmptyStr ) then
  begin
    WarningMsg( '�п�� Excel �ɮ׶ץX���|���ɮצW�١C' );
    if txtFile.CanFocus then txtFile.SetFocus;
    Exit;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC03.btnQueryClick(Sender: TObject);
var
  aWhere1, aWhere2, aErrMsg: String;
  aDataCounts, aConditionType: Integer;
  aConditionString, aReportSQL: String;
begin
  aConditionString := EmptyStr;
  if not IsDataOK then Exit;
  aConditionType := GetConditionType;
  if ( aConditionType = CONDITION_TYPE_BILLNO ) then
  begin
    aWhere1 := txtBillNo.Text;
    aWhere2 := '';
    aConditionString := '���O�渹: ' + txtBillNo.Text;
  end
  else if ( aConditionType = CONDITION_TYPE_CUSTID ) then
  begin
    aWhere1 := txtCustId.Text;
    aWhere2 := '';
    aConditionString := '�Ȥ�s��: ' + txtCustId.Text;
  end
  else if ( aConditionType = CONDITION_TYPE_CHARGEDATE ) then
  begin
    aWhere1 := FormatDateTime( 'yyyy/mm/dd', txtChageDateSt.Date );
    aWhere2 := FormatDateTime( 'yyyy/mm/dd', txtChageDateEd.Date );
    aConditionString := '���O���: ' + aWhere1 + ' ~ ' + aWhere2;
  end
  else if ( aConditionType = CONDITION_TYPE_LOGDATE ) then
  begin
    aWhere1 := FormatDateTime( 'yyyy/mm/dd', txtLogDateSt.Date );
    aWhere2 := FormatDateTime( 'yyyy/mm/dd', txtLogDateEd.Date );;
    aConditionString := 'Log ���: ' + aWhere1 + ' ~ ' + aWhere2;
  end;

  //���X�ŦX���󪺶}�ߵo�����`���
  aDataCounts := dtmMainJ.getUnusualInvID( aConditionType, aWhere1,
    aWhere2, aReportSQL );

  if ( aDataCounts <= 0 ) then
  begin
    WarningMsg( '�d�L��ơC' );
    Exit;
  end;

  Screen.Cursor := crSQLWait;
  try
    if ( rdoReport.Checked ) then
    begin
      if not DoRunReport( aErrMsg, aReportSQL, aConditionString ) then
        ErrorMsg( aErrMsg );
    end else
    begin
      UnPrepareDataBuffer;
      PrepareDataBuffer;
      if DoExportToExcel( aErrMsg, aConditionString ) then
        InfoMsg( '��XExcel�����C' )
      else
        ErrorMsg( Format( '��XExcel����, �T��:%s�C', [aErrMsg] ) );
      PBar.Position := 0;        
    end;
  finally
    Screen.Cursor := crDefault;
  end;  
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvC03.DoRunReport(var aMsg: String; aSql, aCondition: String): Boolean;
var
  aPath: String;
begin
  aMsg := EmptyStr;
  aPath := IncludeTrailingPathDelimiter( ExtractFilePath(
    Application.ExeName ) ) + REPORT_FOLDER + '\FrptInvC03_1.fr3';
  try
    dtmReport.frxMainReport.LoadFromFile( aPath );
    dtmReport.ADOMaster.Close;
    dtmReport.ADOMaster.SQL.Text := aSql;
    dtmReport.frxMainReport.Variables.Variables['aCondition'] :=
      QuotedStr( aCondition );
    dtmReport.frxMainReport.Variables.Variables['aOperator'] :=
      QuotedStr( dtmMain.getLoginUser );
    dtmReport.frxMainReport.ShowReport;
  except
    on E: Exception do aMsg := E.Message;
  end;
  dtmReport.ADOMaster.Close;
  Result := ( aMsg = EmptyStr );  
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvC03.DoExportToExcel(var aMsg: String; aCondition: String): Boolean;
var
  aSource: TClientDataSet;
  aFileName: String;
  aTotalAmount, aCounts: Integer;
begin
  aMsg := EmptyStr;
  aSource := dtmMainJ.cdsInv033;
  {}
  PBar.Properties.Min := 0;
  PBar.Properties.Max := aSource.RecordCount;
  PBar.Position := 0;
  Application.ProcessMessages;
  {}
  try
    aTotalAmount := 0;
    aCounts := 0;
    aSource.First;
    while not aSource.Eof do
    begin
      FDataBuffer.Append;
      FDataBuffer.FieldByName( 'BILLID' ).AsString := aSource.FieldByName( 'BILLID' ).AsString;
      FDataBuffer.FieldByName( 'CUSTID' ).AsString := aSource.FieldByName( 'CUSTID' ).AsString;
      FDataBuffer.FieldByName( 'CUSTNAME' ).AsString := aSource.FieldByName( 'CUSTNAME' ).AsString;
      FDataBuffer.FieldByName( 'DESCRIPTION' ).AsString := aSource.FieldByName( 'DESCRIPTION' ).AsString;
      FDataBuffer.FieldByName( 'TOTALAMOUNT' ).AsString := FormatFloat( ',###', aSource.FieldByName( 'TOTALAMOUNT' ).AsInteger );
      {}
      FDataBuffer.FieldByName( 'CHARGEDATE' ).AsString := EmptyStr;
      if not VarIsNull( aSource.FieldByName( 'CHARGEDATE' ).Value ) then
        FDataBuffer.FieldByName( 'CHARGEDATE' ).AsString := FormatDateTime( 'yyyy/mm/dd', aSource.FieldByName( 'CHARGEDATE' ).AsDateTime );
      {}
      FDataBuffer.FieldByName( 'LOGTIME' ).AsString := EmptyStr;
      if not VarIsNull( aSource.FieldByName( 'LOGTIME' ).Value ) then
        FDataBuffer.FieldByName( 'LOGTIME' ).AsString := FormatDateTime( 'yyyy/mm/dd hh:nn:ss', aSource.FieldByName( 'LOGTIME' ).AsDateTime );
      {}
      FDataBuffer.FieldByName( 'NOTES' ).AsString := aSource.FieldByName( 'NOTES' ).AsString;
      FDataBuffer.Post;
      {}
      Inc( aTotalAmount, aSource.FieldByName( 'TOTALAMOUNT' ).AsInteger );
      Inc( aCounts );
      {}
      PBar.Position := ( PBar.Position + 1 );
      Application.ProcessMessages;
      {}
      aSource.Next;
    end;
    aSource.Close;
    { �[�`���B, ���[ 2 ��ť� }
    FDataBuffer.Append;
    FDataBuffer.Post;
    FDataBuffer.Append;
    FDataBuffer.Post;
    {}
    FDataBuffer.Append;
    FDataBuffer.FieldByName( 'DESCRIPTION' ).AsString := '���B�[�`';
    FDataBuffer.FieldByName( 'TOTALAMOUNT' ).AsString := FormatFloat( ',###', aTotalAmount );
    FDataBuffer.Post;
    FDataBuffer.Append;
    FDataBuffer.FieldByName( 'DESCRIPTION' ).AsString := '����';
    FDataBuffer.FieldByName( 'TOTALAMOUNT' ).AsInteger := aCounts;
    FDataBuffer.Post;
    {}
    FDataBuffer.First;
    aCondition :=
      '����W��:�o���}�߲��`��@' +
      Format( '���q�W��:%s@', [dtmMain.getCompName] ) + aCondition;
    DataSetToXLS( FDataBuffer, txtFile.Text, aCondition, EmptyStr );
    {}
  except
    on E: Exception do aMsg := E.Message;
  end;
  Result := ( aMsg = EmptyStr );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC03.txtFilePropertiesButtonClick(Sender: TObject;
  AButtonIndex: Integer);
var
  aExt: String;
begin
  if not OpenDialog1.Execute then Exit;
  txtFile.Text := OpenDialog1.FileName;
  aExt := ExtractFileExt( ExtractFileName( txtFile.Text ) );
  if ( aExt = EmptyStr ) then
    txtFile.Text := ( txtFile.Text + '.xls' ); 
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC03.txtBillNoEnter(Sender: TObject);
begin
  rdoBillNo.Checked := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC03.txtCustIdEnter(Sender: TObject);
begin
  rdoCustId.Checked := True;
end;

{ ---------------------------------------------------------------------------- }


procedure TfrmInvC03.txtChageDateStEnter(Sender: TObject);
begin
  rdoChargeDate.Checked := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC03.txtLogDateStEnter(Sender: TObject);
begin
  rdoLogDate.Checked := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC03.txtLogDateStExit(Sender: TObject);
begin
  if ( txtLogDateEd.Text = EmptyStr ) then
    txtLogDateEd.Date := txtLogDateSt.Date; 
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC03.txtChageDateStExit(Sender: TObject);
begin
  if ( txtChageDateEd.Text = EmptyStr ) then
    txtChageDateEd.Date := txtChageDateSt.Date;
end;

{ ---------------------------------------------------------------------------- }

end.
