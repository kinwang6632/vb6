unit frmInvD02_1U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Buttons, DB, Grids, DBGrids, Mask, ADOInt,
  cxLookAndFeelPainters, dxSkinsCore, dxSkinBlack, dxSkinBlue,
  dxSkinCaramel, dxSkinCoffee, dxSkinsDefaultPainters, cxControls,
  cxContainer, cxEdit, cxGroupBox, cxCheckGroup, cxCheckBox;

type
  TfrmInvD02_1 = class(TForm)
    Panel1: TPanel;
    Label5: TLabel;
    pnlEdit: TPanel;
    pnlGrid: TPanel;
    DBGrid1: TDBGrid;
    dsrInv002: TDataSource;
    Panel4: TPanel;
    btnAppend: TBitBtn;
    btnEdit: TBitBtn;
    btnDelete: TBitBtn;
    btnCancel: TBitBtn;
    btnSave: TBitBtn;
    btnExit: TBitBtn;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    Label17: TLabel;
    Label18: TLabel;
    Label19: TLabel;
    Label20: TLabel;
    Label23: TLabel;
    medCustId: TMaskEdit;
    medCustSName: TMaskEdit;
    medCustName: TMaskEdit;
    medMZipCode: TMaskEdit;
    medMailAddr: TMaskEdit;
    medTel1: TMaskEdit;
    medTel2: TMaskEdit;
    medTel3: TMaskEdit;
    medFax: TMaskEdit;
    medAppContactee1: TMaskEdit;
    medAppContactee2: TMaskEdit;
    medFinaContactee1: TMaskEdit;
    medFinaContactee2: TMaskEdit;
    cobIsSelfCreated: TComboBox;
    btnShowDetail: TBitBtn;
    Stt_Show: TStaticText;
    chkIfPrintTitle: TCheckBox;
    grpNoNotify: TGroupBox;
    chkEmail: TcxCheckBox;
    chkMessage: TcxCheckBox;
    chkTVMail: TcxCheckBox;
    chkCM: TcxCheckBox;
    procedure FormShow(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnAppendClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure dsrInv002DataChange(Sender: TObject; Field: TField);
    procedure btnSaveClick(Sender: TObject);
    procedure btnShowDetailClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure medCustNameExit(Sender: TObject);
  private
    { Private declarations }
    procedure ChangeBtnStatus;
    procedure setButtonCompetence;
    procedure initialForm;
    function isDataOK : Boolean;
    function chkNotify:Boolean;
  public
    { Public declarations }
    sG_CompID,
    sG_CustID1,
    sG_CustID2,
    sG_CustName,
    sG_Sno,
    sG_Tel : String;
    bG_Select : Integer;
    function CheckHaveDetail(const aCustId: String): Boolean;
    procedure DeleteCurrentRecord;
    procedure ReOpenInv002;
  end;

var
  frmInvD02_1: TfrmInvD02_1;

implementation

uses cbUtilis, frmMainU, dtmMainJU, dtmMainU, dtmSOU, frmInvD02_2U;

{$R *.dfm}

procedure TfrmInvD02_1.FormShow(Sender: TObject);
begin
   Self.Caption := frmMain.GetFormTitleString( 'D02_1', '客戶基本資料維護' );
   ChangeBtnStatus;
   setButtonCompetence;
   initialForm;
   dtmMainJ.adoInv002Code.Properties.Item['Update Criteria'].Value := adCriteriaKey;
   //#5764 增加設定不想通知的選項 By Kin 2010/09/13
   if dtmMain.sG_StarEInvoice then
   begin
     grpNoNotify.Visible := True;
     grpNoNotify.Enabled := True;
     chkEmail.Visible := dtmMain.sG_StarEmail;
     chkMessage.Visible := dtmMain.sG_StarMessage;
     chkTVMail.Visible := dtmMain.sG_StarTVmail;
     chkCM.Visible := dtmMain.sG_StarCMSend;
   end;
end;

procedure TfrmInvD02_1.btnExitClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmInvD02_1.ChangeBtnStatus;
begin
     with dtmMainJ.adoInv002Code do
     begin
       if state in [dsInactive] then
         Exit
       else if state in [dsEdit, dsInsert] then
       begin
          btnCancel.Enabled := TRUE;
          btnSave.Enabled := TRUE;
          btnDelete.Enabled := FALSE;
          btnEdit.Enabled := FALSE;
          btnExit.Enabled := FALSE;
          btnAppend.Enabled := FALSE;
          pnlGrid.Enabled := false;
          pnlEdit.Enabled := true;
          btnShowDetail.Enabled := false;
          chkIfPrintTitle.Enabled := True;
       end
       else
       begin
          if dtmMainJ.adoInv002Code.RecordCount>0 then
          begin
            btnAppend.Enabled :=TRUE;
            btnEdit.Enabled := TRUE;
            btnDelete.Enabled := TRUE;
          end
          else
          begin
            btnEdit.Enabled := FALSE;
            btnDelete.Enabled := FALSE;
          end;
          btnCancel.Enabled := FALSE;
          btnSave.Enabled := FALSE;
          btnAppend.Enabled := TRUE;
          btnExit.Enabled := TRUE;
          pnlGrid.Enabled := true;
          pnlEdit.Enabled := FALSE;
          btnShowDetail.Enabled := TRUE;
          chkIfPrintTitle.Enabled := False;
       end;
     end;

end;

procedure TfrmInvD02_1.setButtonCompetence;
var
    sL_QueryCompetence,sL_InsertCompetence,sL_DeleteCompetence,sL_UpdateCompetence,sL_ExecuteCompetence : String;
    sL_VerifyCompetence : String;
begin
     //設定權限
     dtmMain.getCompetence('D02',sL_QueryCompetence,sL_InsertCompetence,
        sL_DeleteCompetence,sL_UpdateCompetence,
        sL_ExecuteCompetence,sL_VerifyCompetence);

     if sL_InsertCompetence = 'Y' then
       btnAppend.Enabled := true
     else
       btnAppend.Enabled := false;


     if sL_DeleteCompetence = 'Y' then
       btnDelete.Enabled := true
     else
       btnDelete.Enabled := false;


     if sL_UpdateCompetence = 'Y' then
       btnEdit.Enabled := true
     else
       btnEdit.Enabled := false;
end;

procedure TfrmInvD02_1.initialForm;
begin
    with dsrInv002.DataSet do
    begin
      FindField('CustId').DisplayLabel := '客戶代碼';
      FindField('CustSName').DisplayLabel := '客戶簡稱';
      FindField('Tel1').DisplayLabel := '電話1';
      FindField('UptTime').DisplayLabel := '異動時間';
      FindField('UptEn').DisplayLabel := '異動人員';
//      FindField('businessid').DisplayLabel := '統一編號';

      FindField('CustSName').DisplayWidth := 30;
      FindField('UptEn').DisplayWidth := 10;

      FindField('IdentifyId1').Visible := false;
      FindField('IdentifyId2').Visible := false;
      FindField('CompId').Visible := false;
      FindField('CustName').Visible := false;
      FindField('MZipCode').Visible := false;
      FindField('MailAddr').Visible := false;
      FindField('Tel2').Visible := false;
      FindField('Tel3').Visible := false;
      FindField('Fax').Visible := false;
      FindField('AppContactee1').Visible := false;
      FindField('AppContactee2').Visible := false;
      FindField('FinaContactee1').Visible := false;
      FindField('FinaContactee2').Visible := false;
      FindField('IsSelfCreated').Visible := false;
      FindField('IfPrintTitle').Visible := false;
      FindField( 'IsEmail' ).Visible := False;
      FindField( 'IsMobile' ).Visible := False;
      FindField( 'IsTvMail' ).Visible := False;
      FindField( 'IsCM' ).Visible := False;
      if RecordCount > 0 then
        btnShowDetail.Enabled := true
      else
        btnShowDetail.Enabled := false;

      if RecordCount = 0 then
      begin
        btnEdit.Enabled := false;
        btnDelete.Enabled := false;
      end;

    end;
end;

procedure TfrmInvD02_1.btnAppendClick(Sender: TObject);
begin
  if not dsrInv002.DataSet.IsEmpty then
  begin
    if not CheckHaveDetail( medCustId.Text ) then Exit;
  end;      
  dtmMainJ.adoInv002Code.Append;
  cobIsSelfCreated.ItemIndex := 0;   //是否由本系統create預設為是
  ChangeBtnStatus;
  medCustID.SetFocus;
end;

procedure TfrmInvD02_1.btnEditClick(Sender: TObject);
begin
    dtmMainJ.adoInv002Code.Edit;
    ChangeBtnStatus;
    medCustId.Enabled := false;
end;

procedure TfrmInvD02_1.btnDeleteClick(Sender: TObject);
var
    sL_CustID : String;
begin
    if ConfirmMsg( '請確認要刪除此筆資料?' ) then
    begin
      sL_CustID := Trim(medCustId.Text);
      if dtmMainJ.getInv019Data(sG_CompID,sL_CustID) = 0 then
      begin
        DeleteCurrentRecord;
        //dtmMainJ.adoInv002Code.Delete;
        ReOpenInv002;
        ChangeBtnStatus;
      end else
      begin
        WarningMsg( '此筆資料尚有明細資料,不能刪除' );
      end;
    end;
    setButtonCompetence;
    initialForm;
end;

procedure TfrmInvD02_1.btnCancelClick(Sender: TObject);
begin
    if  dsrInv002.DataSet.State in [dsEdit] then
      medCustId.Enabled := true;

    dtmMainJ.adoInv002Code.Cancel;
    dtmMainJ.adoInv002Code.First;
    ChangeBtnStatus;

    setButtonCompetence;
    initialForm;
end;

procedure TfrmInvD02_1.dsrInv002DataChange(Sender: TObject; Field: TField);
var
  sL_IsSelfCreated : String;
  i: Integer;
begin
    with dsrInv002.DataSet do
    begin
      if RecordCount > 0 then
      begin
        medCustId.Text := FieldByName('CustId').AsString;
        medCustSName.Text := FieldByName('CustSName').AsString;
        medCustName.Text := FieldByName('CustName').AsString;

        medMZipCode.Text := FieldByName('MZipCode').AsString;
        medMailAddr.Text := FieldByName('MailAddr').AsString;
        medTel1.Text := FieldByName('Tel1').AsString;
        medTel2.Text := FieldByName('Tel2').AsString;
        medTel3.Text := FieldByName('Tel3').AsString;
        medFax.Text := FieldByName('Fax').AsString;
        medAppContactee1.Text := FieldByName('AppContactee1').AsString;
        medAppContactee2.Text := FieldByName('AppContactee2').AsString;
        medFinaContactee1.Text := FieldByName('FinaContactee1').AsString;
        medFinaContactee2.Text := FieldByName('FinaContactee2').AsString;

        sL_IsSelfCreated := FieldByName('IsSelfCreated').AsString;

        if sL_IsSelfCreated = 'Y' then
          cobIsSelfCreated.ItemIndex := 0
        else
          cobIsSelfCreated.ItemIndex := 1;

        chkIfPrintTitle.Checked := ( FieldByName( 'IfPrintTitle' ).AsString = 'Y' );
        
        {#5764 增加不想設定通知的選項 By Kin 2010/09/13}
        chkEmail.Checked := FieldByName('IsEmail').AsString = '1';
        chkMessage.Checked := FieldByName('IsMobile').AsString = '1';
        chkTVMail.Checked := FieldByName( 'IsTVMail' ).AsString = '1';
        chkCM.Checked := FieldByName('IsCM').AsString = '1';
      end
      else
      begin
        medCustId.Text := '';
        medCustSName.Text := '';
        medCustName.Text := '';

        medMZipCode.Text := '';
        medMailAddr.Text := '';
        medTel1.Text := '';
        medTel2.Text := '';
        medTel3.Text := '';
        medFax.Text := '';
        medAppContactee1.Text := '';
        medAppContactee2.Text := '';
        medFinaContactee1.Text := '';
        medFinaContactee2.Text := '';

        cobIsSelfCreated.ItemIndex := 0;
        chkIfPrintTitle.Checked := False;
        {#5764 增加不想設定通知的選項 By Kin 2010/09/13}
        chkEmail.Checked := False;
        chkTVMail.Checked := False;
        chkMessage.Checked := False;
        chkCM.Checked := False;
      end;
      Stt_Show.Caption := inttostr(RecNo)+'　/　'+inttostr(recordcount);
    end;
end;

function TfrmInvD02_1.isDataOK: Boolean;
var i : Integer;
    aSelAll : Boolean;
begin
  Result := False;
  aSelAll := True;
  if Trim( medCustId.Text ) = EmptyStr then
  begin
    WarningMsg( '請輸入客戶代碼' );
    if medCustId.CanFocus then medCustId.SetFocus;
    Exit;
  end;

  if ( chkEmail.Visible ) and ( aSelAll )  then
  begin
    if ( not chkEmail.Checked ) then aSelAll := False;
  end;
  if ( chkMessage.Visible ) and ( aSelAll) then
  begin
    if ( not chkMessage.Checked ) then aSelAll := False;
  end;
  if ( chkTVMail.Visible ) and ( aSelAll ) then
  begin
    if ( not chkTVMail.Checked ) then aSelAll := False;
  end;
  if ( chkCM.Visible ) and ( aSelAll) then
  begin
    if ( not chkCM.Checked ) then aSelAll := False;
  end;
  if aSelAll then
  begin
    WarningMsg( '不想通知模式不可全選' );
    Exit;
  end;
  Result := True;
end;

procedure TfrmInvD02_1.btnSaveClick(Sender: TObject);
var
    sL_CustId,sL_CustSName,sL_CustName,sL_MZipCode,sL_MailAddr,sL_Tel1,sL_Tel2,sL_Tel3 : String;
    sL_Fax,sL_AppContactee1,sL_AppContactee2,sL_FinaContactee1,sL_FinaContactee2,sL_IsSelfCreated : String;
    sL_SQL : String;
    bL_HavePKValue, aIfPrintTitle,aIsEmail,aIsMessage,aIsTVMail,aIsCM : Boolean;

begin
    if isDataOK then
    begin
      sL_CustId := Trim(medCustId.Text);
      sL_CustSName := Trim(medCustSName.Text);
      sL_CustName := Trim(medCustName.Text);
      sL_MZipCode := Trim(medMZipCode.Text);
      sL_MailAddr := Trim(medMailAddr.Text);
      sL_Tel1 := Trim(medTel1.Text);
      sL_Tel2 := Trim(medTel2.Text);
      sL_Tel3 := Trim(medTel3.Text);
      sL_Fax := Trim(medFax.Text);
      sL_AppContactee1 := Trim(medAppContactee1.Text);
      sL_AppContactee2 := Trim(medAppContactee2.Text);
      sL_FinaContactee1 := Trim(medFinaContactee1.Text);
      sL_FinaContactee2 := Trim(medFinaContactee2.Text);
      aIfPrintTitle := chkIfPrintTitle.Checked;
      aIsEmail := chkEmail.Checked;
      aIsMessage := chkMessage.Checked;
      aIsTVMail := chkTVMail.Checked;
      aIsCM := chkCM.Checked;
      if cobIsSelfCreated.ItemIndex = 0 then  //是否由本系統create
        sL_IsSelfCreated := 'Y'
      else
        sL_IsSelfCreated := 'N';

      //檢核PK值
      sL_SQL := 'select * from inv002 where IdentifyId1=''' + IDENTIFYID1 +
                ''' and IdentifyId2=' + IDENTIFYID2 + ' and CompID=''' + sG_CompID +
                ''' and CustId=''' + sL_CustId + '''';


      if  dsrInv002.DataSet.State in [dsInsert] then
        bL_HavePKValue := dtmMainJ.checkPK(sL_SQL)
      else
        bL_HavePKValue := false;
        

      if bL_HavePKValue then
      begin
        MessageDlg('違反唯一值條件',mtError,[mbOk],0);
        medCustId.SetFocus;
      end
      else
      begin
            
        with dsrInv002.DataSet do
        begin
          FieldByName('IdentifyId1').AsString := IDENTIFYID1;
          FieldByName('IdentifyId2').AsString := IDENTIFYID2;

          FieldByName('CompID').AsString := sG_CompID;
          FieldByName('CustId').AsString := sL_CustId;
          FieldByName('CustSName').AsString := sL_CustSName;

          FieldByName('CustName').AsString := sL_CustName;
          FieldByName('MZipCode').AsString := sL_MZipCode;
          FieldByName('MailAddr').AsString := sL_MailAddr;
          FieldByName('Tel1').AsString := sL_Tel1;
          FieldByName('Tel2').AsString := sL_Tel2;
          FieldByName('Tel3').AsString := sL_Tel3;
          FieldByName('Fax').AsString := sL_Fax;
          FieldByName('AppContactee1').AsString := sL_AppContactee1;
          FieldByName('AppContactee2').AsString := sL_AppContactee2;
          FieldByName('FinaContactee1').AsString := sL_FinaContactee1;
          FieldByName('FinaContactee2').AsString := sL_FinaContactee2;
          FieldByName('IsSelfCreated').AsString := sL_IsSelfCreated;

          FieldByName('UptTime').AsString := DateTimeToStr(now);
          FieldByName('UptEn').AsString := dtmMain.getLoginUser;

          FieldByName( 'IfPrintTitle' ).AsString := 'N';
          if ( aIfPrintTitle ) then
            FieldByName( 'IfPrintTitle' ).AsString := 'Y';

          {#5764 增加不想設定通知的選項 By Kin 2010/09/13}
          FieldByName( 'IsEmail' ).AsString := '0';
          FieldByName( 'IsMobile' ).AsString := '0';
          FieldByName( 'IsTvMail' ).AsString := '0';
          FieldByName( 'IsCM' ).AsString :='0';
          if aIsEmail then  FieldByName( 'IsEmail' ).AsString := '1';
          if aIsMessage then FieldByName( 'IsMobile' ).AsString := '1';
          if aIsTVMail then FieldByName( 'IsTvMail' ).AsString := '1';
          if aIsCM then FieldByName( 'IsCM' ).AsString :='1';

          dsrInv002.DataSet.Post;
          dtmMainJ.adoInv002Code.Requery;

        end;
        medCustId.Enabled := true;
        ChangeBtnStatus;
        setButtonCompetence;
//        dtmMainJ.getInv002Data(bG_Select,sG_CustID1,sG_CustID2,sG_CustName,
//                                sG_Sno,sG_Tel,sG_CompID);
//        dtmMainJ.getInv002Data(bG_Select,bG_SelectTel,sG_CustID,sG_Tel,sG_CompID);
        initialForm;
      end;
    end;
end;

procedure TfrmInvD02_1.btnShowDetailClick(Sender: TObject);
var
    sL_CustID : String;
begin
  sL_CustID := dsrInv002.DataSet.FieldByName('CustID').AsString;
  frmInvD02_2 := TfrmInvD02_2.Create(Application);
  try
    frmInvD02_2.sG_CompID := sG_CompID;
    frmInvD02_2.sG_CustID := sL_CustID;
    frmInvD02_2.CustName := dsrInv002.DataSet.FieldByName('CUSTNAME').AsString;
    frmInvD02_2.CustSName := dsrInv002.DataSet.FieldByName('CUSTSNAME').AsString;
    frmInvD02_2.ShowModal;
  finally
    frmInvD02_2.Free;
  end;  
end;

function TfrmInvD02_1.CheckHaveDetail(const aCustId: String): Boolean;
begin
  dtmMain.adoComm.Close;
  dtmMain.adoComm.SQL.Text := Format(
    ' SELECT COUNT(1) AS COUNTS    ' +
    '   FROM INV019                ' +
    '  WHERE IDENTIFYID1 = ''%s''  ' +
    '    AND IDENTIFYID2 = ''%s''  ' +
    '    AND COMPID = ''%s''       ' +
    '    AND CUSTID = ''%s''       ',
    [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, Trim( aCustId )] );
  dtmMain.adoComm.Open;
  Result := ( dtmMain.adoComm.FieldByName( 'COUNTS' ).AsInteger > 0 );
  dtmMain.adoComm.Close;
  if not Result then
  begin
    WarningMsg( '此客戶無對應明細資料, 請先輸入此客戶明細資料。' );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD02_1.DeleteCurrentRecord;
begin
  dtmMain.adoComm.Close;
  dtmMain.adoComm.SQL.Text := Format(
    ' DELETE FROM INV019           ' +
    '  WHERE IDENTIFYID1 = ''%s''  ' +
    '    AND IDENTIFYID2 = ''%s''  ' +
    '    AND COMPID = ''%s''       ' +
    '    AND CUSTID = ''%s''       ',
    [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID,Trim( medCustId.Text) ] );
  dtmMain.adoComm.ExecSQL;
  {}
  dtmMain.adoComm.SQL.Text := Format(
    ' DELETE FROM INV002           ' +
    '  WHERE IDENTIFYID1 = ''%s''  ' +
    '    AND IDENTIFYID2 = ''%s''  ' +
    '    AND COMPID = ''%s''       ' +
    '    AND CUSTID = ''%s''       ',
    [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID,Trim( medCustId.Text )] );
  dtmMain.adoComm.ExecSQL;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD02_1.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  if not dsrInv002.DataSet.IsEmpty then
    CanClose := CheckHaveDetail( medCustId.Text );    
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD02_1.medCustNameExit(Sender: TObject);
begin
  medCustSName.Text := medCustName.Text;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD02_1.ReOpenInv002;
begin
   dtmMainJ.adoInv002Code.Close;
   dtmMainJ.adoInv002Code.Open;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvD02_1.chkNotify: Boolean;

begin
  

end;

end.
