unit frmInvD0BU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, DB, dxSkinsCore, dxSkinBlack, dxSkinBlue,
  dxSkinCaramel, dxSkinCoffee, dxSkinsDefaultPainters, cxControls,
  cxContainer, cxEdit, cxCheckBox, StdCtrls, Buttons, cxLabel, cxTextEdit,
  cxCurrencyEdit, cxLookAndFeelPainters, cxStyles, dxSkinscxPCPainter,
  cxCustomData, cxGraphics, cxFilter, cxData, cxDataStorage, cxDBData,
  cxGridLevel, cxClasses, cxGridCustomView, cxGridCustomTableView,
  cxGridTableView, cxGridDBTableView, cxGrid, cxGroupBox, cxRadioGroup,
  ADODB,cbAppClass, cxMaskEdit, cxDropDownEdit;

type
  TfrmInvD0B = class(TForm)
    pnlEdit: TPanel;
    chkEmail: TcxCheckBox;
    edtMailSet: TcxCurrencyEdit;
    lblTVMail: TcxLabel;
    chkMobile: TcxCheckBox;
    lblMobile: TcxLabel;
    edtMobileSet: TcxCurrencyEdit;
    lblEmail: TcxLabel;
    chkTVMail: TcxCheckBox;
    chkCM: TcxCheckBox;
    edtTVMailSet: TcxCurrencyEdit;
    lblCM: TcxLabel;
    cxLabel5: TcxLabel;
    edtOrder: TcxTextEdit;
    rdgInvUse: TcxRadioGroup;
    cxLabel6: TcxLabel;
    edtAddGive: TcxTextEdit;
    adoINV042: TADOQuery;
    DataSource1: TDataSource;
    pnlGrid: TPanel;
    cxGrid1: TcxGrid;
    cxGrid1DBTableView1: TcxGridDBTableView;
    gvCompId: TcxGridDBColumn;
    gvEmailNotify: TcxGridDBColumn;
    gvSmsNotify: TcxGridDBColumn;
    gvTVMAILNotify: TcxGridDBColumn;
    gvCMNotify: TcxGridDBColumn;
    gvPriorityOrder: TcxGridDBColumn;
    gvAddGive: TcxGridDBColumn;
    gvINVUSEID: TcxGridDBColumn;
    gvUpdEn: TcxGridDBColumn;
    gvUpdTime: TcxGridDBColumn;
    cxGrid1Level1: TcxGridLevel;
    Panel3: TPanel;
    btnAppend: TBitBtn;
    btnEdit: TBitBtn;
    btnDelete: TBitBtn;
    btnCancel: TBitBtn;
    btnSave: TBitBtn;
    btnExit: TBitBtn;
    edtCMSet: TcxCurrencyEdit;
    cboComId: TcxComboBox;
    Label1: TLabel;
    procedure btnAppendClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure adoINV042AfterScroll(DataSet: TDataSet);
    procedure FormShow(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure chkEmailClick(Sender: TObject);
    procedure chkMobileClick(Sender: TObject);
    procedure chkTVMailClick(Sender: TObject);
    procedure chkCMClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure edtOrderKeyPress(Sender: TObject; var Key: Char);
    procedure edtAddGiveKeyPress(Sender: TObject; var Key: Char);
    procedure DataSource1DataChange(Sender: TObject; Field: TField);
  private
    { Private declarations }
    FDMLMode: TDMLMode;
    function DoDMLGridChange(const AMode: TDMLMode): Boolean;
    function DoDMLModeChange(AMode: TDMLMode): Boolean;
    function DoDMLDataSetChange(const AMode: TDMLMode): Boolean;
    function DoDMLButtonChange(const AMode: TDMLMode): Boolean;
    function DoDMLEditorChange(const AMode: TDMLMode): Boolean;
    function IsDataOK:Boolean;
    function SaveData :Boolean;
    procedure OpenDataSet;
  public
    { Public declarations }
  end;

var
  frmInvD0B: TfrmInvD0B;

implementation
uses cbUtilis, frmMainU, dtmMainJU, dtmMainU, Math;
{$R *.dfm}

procedure TfrmInvD0B.btnAppendClick(Sender: TObject);
begin
  DoDMLModeChange( dmInsert );
end;

function TfrmInvD0B.DoDMLButtonChange(const AMode: TDMLMode): Boolean;
begin
  case AMode of
    dmBrowser, dmCancel:
      begin
        btnAppend.Enabled := True;
        btnEdit.Enabled := True;
        btnDelete.Enabled := True;
        btnCancel.Enabled := False;
        btnSave.Enabled := False;
        btnExit.Enabled := True;
        cboComId.Enabled := True;
        if adoInv042.IsEmpty then
        begin
          btnEdit.Enabled := False;
          btnDelete.Enabled := False;
        end else
        begin
          //btnAppend.Enabled := False;
        end;
      end;
    dmInsert:
      begin
        btnAppend.Enabled := False;
        btnEdit.Enabled := False;
        btnDelete.Enabled := False;
        btnCancel.Enabled := True;
        btnSave.Enabled := True;
        btnExit.Enabled := False;
        cboComId.Enabled := True;
      end;
    dmUpdate:
      begin
        btnAppend.Enabled := False;
        btnEdit.Enabled := False;
        btnDelete.Enabled := False;
        btnCancel.Enabled := True;
        btnSave.Enabled := True;
        btnExit.Enabled := False;
        cboComId.Enabled := False;
      end;
  else
    begin
      btnAppend.Enabled := True;
      btnEdit.Enabled := True;
      btnDelete.Enabled := True;
      btnCancel.Enabled := False;
      btnSave.Enabled := False;
      btnExit.Enabled := True;
    end;
  end;
  Result := True;
end;

function TfrmInvD0B.DoDMLDataSetChange(const AMode: TDMLMode): Boolean;
begin
  Result := False;
  case AMode of
    dmBrowser, dmCancel:
      begin
        if ( adoInv042.State in [dsEdit, dsInsert] ) then
        begin
          adoInv042.AfterScroll := nil;
          try
            adoInv042.Cancel;
          finally
            adoInv042.AfterScroll := adoInv042AfterScroll;
          end;
          adoInv042AfterScroll( adoInv042 );
        end;
      end;
    dmInsert:
      begin
        adoInv042.Append;
        rdgInvUse.ItemIndex := 2;
      end;
    dmUpdate:
      begin
        if not adoInv042.IsEmpty then
          adoInv042.Edit;
      end;
    dmDelete:
      begin
        adoInv042.Delete;
      end;
    dmSave:
      begin
        if not SaveData then
        begin
      
          Result := False;
          Exit;
        end;

      end;
  end;
  Result := True;
end;

function TfrmInvD0B.DoDMLEditorChange(const AMode: TDMLMode): Boolean;
begin
  case AMode of
    dmBrowser, dmCancel:
      begin
        pnlEdit.Enabled := False;
      end;
    dmInsert:
      begin
        pnlEdit.Enabled := True;
        if chkEmail.CanFocus then chkEmail.SetFocus;
      end;
    dmUpdate:
      begin
        pnlEdit.Enabled := True;
      end;
  else
    pnlGrid.Enabled := True;
  end;
  {
  chkEmail.Enabled := dtmMain.sG_StarEmail;
  chkMobile.Enabled := dtmMain.sG_StarMessage;
  chkTVMail.Enabled := dtmMain.sG_StarTVmail;
  chkCM.Enabled := dtmMain.sG_StarCMSend;
  }
  chkEmail.Visible := dtmMain.sG_StarEmail;
  chkMobile.Visible := dtmMain.sG_StarMessage;
  chkTVMail.Visible := dtmMain.sG_StarTVmail;
  chkCM.Visible := dtmMain.sG_StarCMSend;
  edtMailSet.Visible := dtmMain.sG_StarEmail;
  edtMobileSet.Visible := dtmMain.sG_StarMessage;
  edtTVMailSet.Visible := dtmMain.sG_StarTVmail;
  edtCMSet.Visible := dtmMain.sG_StarCMSend;
  lblEmail.Visible := dtmMain.sG_StarEmail;
  lblMobile.Visible := dtmMain.sG_StarMessage;
  lblTVMail.Visible := dtmMain.sG_StarTVmail;
  lblCM.Visible := dtmMain.sG_StarCMSend;
  if not chkEmail.Enabled then edtMailSet.Enabled := False;
  if not chkMobile.Enabled then edtMobileSet.Enabled := False;
  if not chkTVMail.Enabled then edtTVMailSet.Enabled := False;
  if not chkCM.Enabled then edtCMSet.Enabled := False;

  Result := True;
end;

function TfrmInvD0B.DoDMLGridChange(const AMode: TDMLMode): Boolean;
begin
  case AMode of
    dmBrowser, dmCancel:
      pnlGrid.Enabled := True;
    dmInsert, dmUpdate:
      pnlGrid.Enabled := False;
  else
    pnlGrid.Enabled := True;
  end;
  Result := True;
end;

function TfrmInvD0B.DoDMLModeChange(AMode: TDMLMode): Boolean;
begin
  if DoDMLDataSetChange( AMode ) then
  begin
    if ( AMode in [dmSave, dmCancel] ) then AMode := dmBrowser;
    DoDMLButtonChange( AMode );
    DoDMLGridChange( AMode );
    DoDMLEditorChange( AMode );
    FDMLMode := AMode;
  end;
end;

procedure TfrmInvD0B.btnEditClick(Sender: TObject);
begin
  if adoINV042.IsEmpty then
  begin
    MessageDlg('無任何資料可修改!',mtInformation,[mbOK],0);
    Exit;
  end;
  DoDMLModeChange( dmUpdate );
end;

procedure TfrmInvD0B.OpenDataSet;
begin

  adoInv042.AfterScroll := nil;
  try
    adoInv042.Close;
    adoInv042.SQL.Text := Format(
      ' SELECT * FROM INV042           ' +
      '  WHERE IDENTIFYID1 = ''%s''    ' +
      '    AND IDENTIFYID2 = ''%s''    ' +
      '  ORDER BY COMPID               ',
      [IDENTIFYID1, IDENTIFYID2, cboComId.Text] );
    adoInv042.Open;
    adoInv042.First;
  finally
    adoInv042.AfterScroll := adoInv042AfterScroll;
  end;
  adoInv042AfterScroll( adoInv042 );
end;

procedure TfrmInvD0B.adoINV042AfterScroll(DataSet: TDataSet);
begin
  if not adoInv042.IsEmpty then
  begin
    chkEmail.Checked := adoINV042.FieldByName('StarEmailNotify').AsInteger = 1;
    chkMobile.Checked := adoINV042.FieldByName('StarSmsNotify').AsInteger = 1;
    chkTVMail.Checked := adoINV042.FieldByName('StarTVmailNotify').AsInteger = 1;
    chkCM.Checked := adoINV042.FieldByName('StarCMNotify').AsInteger = 1;
    edtMailSet.EditValue := adoINV042.FieldByName('EmailNotify').AsString;
    edtMobileSet.EditValue := adoINV042.FieldByName('SmsNotify').AsString;
    edtTVMailSet.EditValue := adoINV042.FieldByName('TVMAILNotify').AsString;
    edtCMSet.EditValue := adoINV042.FieldByName('CMNotify').AsString;
    edtOrder.EditValue := adoINV042.FieldByName('PriorityOrder').AsString;
    edtAddGive.EditValue := adoINV042.FieldByName('AddGive').AsString;
    rdgInvUse.ItemIndex := adoINV042.FieldByName('INVUSEID').AsInteger - 1;
    if ( not chkEmail.Enabled ) or ( not chkEmail.Checked )then edtMailSet.Enabled := False;
    if ( not chkMobile.Enabled ) or ( not chkMobile.Checked ) then edtMobileSet.Enabled := False;
    if ( not chkTVMail.Enabled ) or ( not chkTVMail.Checked ) then edtTVMailSet.Enabled := False;
    if ( not chkCM.Enabled ) or ( not chkCM.Checked ) then edtCMSet.Enabled := False;
    cboComId.Text := adoINV042.FieldByName('CompId').AsString;
  end else
  begin
    chkEmail.Checked := False;
    chkMobile.Checked := False;
    chkTVMail.Checked := False;
    chkCM.Checked := False;
    edtMailSet.Enabled := False;
    edtMobileSet.Enabled := False;
    edtTVMailSet.Enabled := False;
    edtCMSet.Enabled := False;
    edtMailSet.EditValue := EmptyStr;
    edtMobileSet.EditValue := EmptyStr;
    edtTVMailSet.EditValue := EmptyStr;
    edtCMSet.EditValue := EmptyStr;
    edtOrder.EditValue := EmptyStr;
    edtOrder.EditValue := EmptyStr;
    edtAddGive.EditValue := EmptyStr;
  end;
end;

procedure TfrmInvD0B.FormShow(Sender: TObject);
begin
  Screen.Cursor := crSQLWait;
  try
    if dtmMainJ.getAllInv001Data > 0 then
    begin
      while not dtmMainJ.adoInv001Code.Eof do
      begin
        cboComId.Properties.Items.Add( dtmMainJ.adoInv001Code.FieldByName('CompId').AsString );
        dtmMainJ.adoInv001Code.Next;
      end;
    end else
    begin
      MessageDlg('無公司資料!請先設定公司資料',mtWarning,[mbOk],0);
    end;
    cboComId.Text := dtmMain.getCompID;
    OpenDataSet;
    DoDMLModeChange( dmBrowser );
  finally
    Screen.Cursor := crDefault;
  end;
end;
procedure TfrmInvD0B.btnCancelClick(Sender: TObject);
begin
  DoDMLModeChange( dmCancel );
end;
function TfrmInvD0B.SaveData : Boolean;
 var aSQL : String;
  aHavePKValue : Boolean;
begin
  aSQL := Format(
      ' select * from inv042 where CompId = ''%s'' Order by CompId ',
      [cboComId.Text]);

  if DataSource1.DataSet.State in [dsInsert] then
    aHavePKValue := dtmMainJ.checkPK(aSQL)
  else
    aHavePKValue := False;
  if aHavePKValue then
  begin
    WarningMsg( '違反唯一值條件' );
    cboComId.SetFocus;
    Result := False;
    Exit;
  end else
  begin
    adoInv042.FieldByName( 'IDENTIFYID1' ).AsString := IDENTIFYID1;
    adoInv042.FieldByName( 'IDENTIFYID2' ).AsString := IDENTIFYID2;
    adoInv042.FieldByName( 'COMPID' ).AsString := cboComId.Text;
    if ( chkEmail.Enabled ) And ( chkEmail.Checked ) then
    begin
      adoINV042.FieldByName( 'StarEmailNotify' ).AsInteger := 1;
      adoINV042.FieldByName( 'EmailNotify' ).AsString := edtMailSet.Text;
    end else
    begin
      adoINV042.FieldByName( 'StarEmailNotify' ).AsInteger := 0;
      adoINV042.FieldByName( 'EmailNotify' ).AsString := EmptyStr;
    end;
    if ( chkMobile.Enabled ) And ( chkMobile.Checked ) then
    begin
      adoINV042.FieldByName( 'StarSmsNotify' ).AsInteger := 1;
      adoINV042.FieldByName( 'SmsNotify' ).AsString := edtMobileSet.Text;
    end else
    begin
      adoINV042.FieldByName( 'StarSmsNotify' ).AsInteger := 0;
      adoINV042.FieldByName( 'SmsNotify' ).AsString := EmptyStr;
    end;
    if ( chkTVMail.Enabled ) And ( chkTVMail.Checked ) then
    begin
      adoINV042.FieldByName( 'StarTVmailNotify' ).AsInteger := 1;
      adoINV042.FieldByName( 'TVMAILNotify' ).AsString := edtTVMailSet.Text;
    end else
    begin
      adoINV042.FieldByName( 'StarTVmailNotify' ).AsInteger := 0;
      adoINV042.FieldByName( 'TVMAILNotify' ).AsString := EmptyStr;
    end;
    if ( chkCM.Enabled ) And ( chkCM.Checked ) then
    begin
      adoINV042.FieldByName( 'StarCMNotify' ).AsInteger := 1;
      adoINV042.FieldByName( 'CMNotify' ).AsString := edtCMSet.Text;
    end else
    begin
      adoINV042.FieldByName( 'StarCMNotify' ).AsInteger := 0;
      adoINV042.FieldByName( 'CMNotify' ).AsString := EmptyStr;
    end;
    adoINV042.FieldByName( 'PriorityOrder' ).AsString := edtOrder.Text;
    adoINV042.FieldByName( 'AddGive' ).AsString := edtAddGive.Text;
    adoINV042.FieldByName( 'INVUSEID' ).AsInteger := rdgInvUse.ItemIndex + 1;
    adoINV042.FieldByName( 'UpdEn' ).AsString := dtmMain.getLoginUserName;
    adoINV042.FieldByName( 'UPDTIME' ).AsDateTime := Now;
    adoINV042.Post;
  end;
  Result := True;
end;

procedure TfrmInvD0B.btnSaveClick(Sender: TObject);
begin
  if IsDataOK then
    DoDMLModeChange( dmSave );
end;

procedure TfrmInvD0B.btnDeleteClick(Sender: TObject);
begin
  if adoINV042.IsEmpty then
  begin
    MessageDlg('已無任何資料可刪除!',mtInformation,[mbOK],0);
    Exit;
  end;
  if ConfirmMsg( '確認要刪除此筆資料?' ) then
  begin
    DoDMLModeChange( dmDelete );
  end;
end;

procedure TfrmInvD0B.chkEmailClick(Sender: TObject);
begin
  edtMailSet.Enabled := False;
  if chkEmail.Checked then
  begin
    edtMailSet.Enabled := True;
  end else
  begin
    edtMailSet.Text := EmptyStr;
  end;
end;

procedure TfrmInvD0B.chkMobileClick(Sender: TObject);
begin
  edtMobileSet.Enabled := False;
  if chkMobile.Checked then
  begin
    edtMobileSet.Enabled := True;
  end else
  begin
    edtMobileSet.Text := EmptyStr;
  end;
end;

procedure TfrmInvD0B.chkTVMailClick(Sender: TObject);
begin
  edtTVMailSet.Enabled := False;
  if chkTVMail.Checked then
  begin
    edtTVMailSet.Enabled := True;
  end else
  begin
    edtTVMailSet.Text := EmptyStr;
  end;
end;

procedure TfrmInvD0B.chkCMClick(Sender: TObject);
begin
  edtCMSet.Enabled := False;
  if chkCM.Checked then
  begin
    edtCMSet.Enabled := True;
  end else
  begin
    edtCMSet.Text := EmptyStr;
  end;
end;

procedure TfrmInvD0B.btnExitClick(Sender: TObject);
begin
  Self.Close;
end;

function TfrmInvD0B.IsDataOK: Boolean;
begin
  Result := False;
  if chkEmail.Checked then
  begin
    if ( edtMailSet.Text = '0' ) or ( edtMailSet.Text = EmptyStr ) then
    begin
      MessageDlg('請設定分鐘數！',mtInformation,[mbOK],0);
      edtMailSet.SetFocus;
      Exit;
    end;
  end;
  if chkMobile.Checked then
  begin
    if ( edtMobileSet.Text = '0' ) or ( edtMobileSet.Text = EmptyStr ) then
    begin
      MessageDlg('請設定分鐘數！',mtInformation,[mbOK],0);
      edtMobileSet.SetFocus;
      Exit;
    end;
  end;
  if chkTVMail.Checked then
  begin
    if ( edtTVMailSet.Text = '0' ) or ( edtTVMailSet.Text = EmptyStr ) then
    begin
      MessageDlg('請設定分鐘數！',mtInformation,[mbOK],0);
      edtTVMailSet.SetFocus;
      Exit;
    end;
  end;
  if chkCM.Checked then
  begin
    if ( edtCMSet.Text = '0' ) or ( edtCMSet.Text = EmptyStr ) then
    begin
      MessageDlg('請設定分鐘數！',mtInformation,[mbOK],0);
      edtCMSet.SetFocus;
      Exit;
    end;
  end;
  if cboComId.Text='' then
  begin
    MessageDlg('請設定公司別！',mtInformation,[mbOK],0);
    if cboComId.CanFocus then
      cboComId.SetFocus;
      Exit;
  end;
  Result := True;
end;

procedure TfrmInvD0B.edtOrderKeyPress(Sender: TObject; var Key: Char);
begin
  if Ord(Key) = 8 then Exit;
  if not (key in ['1','2','3','4']) then
  begin
    Key := #0;
  end;
end;

procedure TfrmInvD0B.edtAddGiveKeyPress(Sender: TObject; var Key: Char);
begin
  if Ord(Key) = 8 then Exit;
  if not (key in ['1','2','3','4']) then
  begin
    Key := #0;
  end;
end;

procedure TfrmInvD0B.DataSource1DataChange(Sender: TObject; Field: TField);
begin
  {
  if adoINV042.Recordset.RecordCount > 0 then
  begin
    cboComId.Text := adoINV042.FieldByName( 'CompId' ).AsString;
  end;
  }
end;

end.
