unit frmInvD01U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Buttons, DB, Mask, DBCtrls, Grids, DBGrids,
  ADODB, fraYMDU, cxControls, cxContainer, cxEdit, cxTextEdit, cxMaskEdit,
  cxCheckBox, dxSkinsCore, dxSkinsDefaultPainters, dxSkinBlack, dxSkinBlue,
  dxSkinCaramel, dxSkinCoffee, cxGraphics, cxLookAndFeels,
  cxLookAndFeelPainters;

type
  TfrmInvD01 = class(TForm)
    Panel1: TPanel;
    Label1: TLabel;
    Panel2: TPanel;
    Panel3: TPanel;
    btnAppend: TBitBtn;
    btnEdit: TBitBtn;
    btnDelete: TBitBtn;
    btnCancel: TBitBtn;
    btnSave: TBitBtn;
    btnExit: TBitBtn;
    pnlMaster: TPanel;
    dsrInv001: TDataSource;
    pnlGrid: TPanel;
    DBGrid1: TDBGrid;
    pnlEdit: TPanel;
    Label3: TLabel;
    Label2: TLabel;
    Label4: TLabel;
    medCompID: TMaskEdit;
    medCompSName: TMaskEdit;
    medCompName: TMaskEdit;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
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
    Label21: TLabel;
    Label22: TLabel;
    Label23: TLabel;
    medBusinessId: TMaskEdit;
    medTaxId: TMaskEdit;
    medManager: TMaskEdit;
    medTel: TMaskEdit;
    fraTaxGetDate: TfraYMD;
    medCompAddr: TMaskEdit;
    medDecaCount: TMaskEdit;
    medPayCount: TMaskEdit;
    medTaxRate: TMaskEdit;
    medTaxDivName: TMaskEdit;
    medTaxNum1: TMaskEdit;
    medTaxNum2: TMaskEdit;
    cobNotiType: TComboBox;
    cobIsSelfCreated: TComboBox;
    Label24: TLabel;
    medServiceTypeStr: TMaskEdit;
    Stt_Show: TStaticText;
    Label25: TLabel;
    cmbLinkToMis: TComboBox;
    Label26: TLabel;
    edtConnSeq: TEdit;
    Label27: TLabel;
    edtAutoCreateNum: TcxMaskEdit;
    chkFaciCombine: TcxCheckBox;
    Label28: TLabel;
    chkShowFaci: TcxCheckBox;
    chkStarEInvoice: TcxCheckBox;
    chkStarEmail: TcxCheckBox;
    chkStarMessage: TcxCheckBox;
    chkStarTVMail: TcxCheckBox;
    chkStarCMSend: TcxCheckBox;
    Label29: TLabel;
    medGraphPath: TMaskEdit;
    Label30: TLabel;
    cboEmailFileType: TComboBox;
    chkAutoNotify: TcxCheckBox;
    Label31: TLabel;
    edtMisOwner: TEdit;
    Label32: TLabel;
    edtSysID: TEdit;
    chkMaskInvNo: TcxCheckBox;
    chkNotifyPrize: TcxCheckBox;
    procedure FormShow(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnAppendClick(Sender: TObject);
    procedure btnEditClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure dsrInv001DataChange(Sender: TObject; Field: TField);
    procedure btnSaveClick(Sender: TObject);
    procedure medBusinessIdExit(Sender: TObject);
    procedure medCompSNameEnter(Sender: TObject);
    procedure medServiceTypeStrExit(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure edtConnSeqKeyPress(Sender: TObject; var Key: Char);
    procedure chkStarEInvoicePropertiesEditValueChanged(Sender: TObject);
  private
    { Private declarations }
    procedure initialForm;
    procedure ChangeBtnStatus;
    procedure setButtonCompetence;
    procedure SwitchEInvoiceState(aValue: Boolean);
    function isDataOK: Boolean;
    function ValidateServiceType(aServiceType: String;  var aMsg: String): Boolean;
  public
    { Public declarations }
  end;

var
  frmInvD01: TfrmInvD01;

implementation

uses cbUtilis, frmMainU, dtmMainJU, dtmMainU;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD01.FormShow(Sender: TObject);
begin
  self.Caption := frmMain.GetFormTitleString( 'D01', '???q???????@' );

  if dtmMainJ.getAllInv001Data = 0 then
    WarningMsg(  '?L?????C' );

  ChangeBtnStatus;
  setButtonCompetence;
  initialForm;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD01.ChangeBtnStatus;
begin
   with dtmMainJ.adoInv001Code do
   begin
     if state in [dsInactive] then
       Exit
     else if state in [dsEdit, dsInsert] then
     begin
        btnCancel.Enabled := True;
        btnSave.Enabled := True;
        btnDelete.Enabled := False;
        btnEdit.Enabled := False;
        btnExit.Enabled := False;
        btnAppend.Enabled := False;
        pnlGrid.Enabled := False;
        pnlEdit.Enabled := True;
     end
     else
     begin
        if dtmMainJ.adoInv001Code.RecordCount>0 then
        begin
          btnAppend.Enabled :=True;
          btnEdit.Enabled := True;
          btnDelete.Enabled := True;
        end
        else
        begin
          btnEdit.Enabled := False;
          btnDelete.Enabled := False;
        end;
        btnCancel.Enabled := False;
        btnSave.Enabled := False;
        btnAppend.Enabled := True;
        btnExit.Enabled := True;
        pnlGrid.Enabled := True;
        pnlEdit.Enabled := False;
     end;
   end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD01.setButtonCompetence;
var
 aQuery, aInsert, aDelete, aUpdate, aExecute,aVerify : String;
begin
   //?]?w?v??
   dtmMain.getCompetence('D01',aQuery,aInsert,aDelete,aUpdate,aExecute,aVerify);

   if aInsert = 'Y' then
     btnAppend.Enabled := True
   else
     btnAppend.Enabled := False;


   if aDelete = 'Y' then
     btnDelete.Enabled := True
   else
     btnDelete.Enabled := False;


   if aUpdate = 'Y' then
     btnEdit.Enabled := True
   else
     btnEdit.Enabled := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD01.initialForm;
begin
    //#5358 ?W?[ShowFaci ?P?_?b?}???o?????O?_?n?N?]???????g?JMemo1 By Kin 2009/12/14
    //#5688 ?W?[?o?????????| By Kin 2010/07/19
    with dsrInv001.DataSet do
    begin
      FindField('CompID').DisplayLabel := '???q?N?X';
      FindField('CompSName').DisplayLabel := '???q????';
      FindField('BusinessId').DisplayLabel := '???@?s??';
      FindField('Manager').DisplayLabel := '?t?d?H';
      FindField('UptTime').DisplayLabel := '????????';
      FindField('UptEn').DisplayLabel := '?????H??';
      FindField('ServiceTypeStr').DisplayLabel := '?~???A??????';
      FindField('LINKTOMIS').DisplayLabel := '?P???A?t??????';
      FindField('CONNSEQ').DisplayLabel := '???A?t???????s??';
      FindField('AUTOCREATENUM').DisplayLabel := '???????}?o??????????';
      FindField('FACICOMBINE').DisplayLabel := '?]???????O?_?X??';
      FindField('ShowFaci').DisplayLabel := '?O?_?????]??????';
      FindField('StartEInvoice').DisplayLabel := '?????q?l?o??';
      FindField('StartEmail').DisplayLabel := '?????o?eEmail';
      FindField('StartMessage').DisplayLabel := '?????o?e???T';
      FindField('InvGlyphPath').DisplayLabel := '?o?????????|';
      FindField('CompSName').DisplayWidth := 30;
      FindField('Manager').DisplayWidth := 20;
      FindField('UptEn').DisplayWidth := 10;

      FindField('IdentifyId1').Visible := False;
      FindField('IdentifyId2').Visible := False;
      FindField('CompName').Visible := False;
      FindField('TaxId').Visible := False;
      FindField('Tel').Visible := False;
      FindField('CompAddr').Visible := False;
      FindField('NotiType').Visible := False;
      FindField('DecaCount').Visible := False;
      FindField('PayCount').Visible := False;
      FindField('TaxRate').Visible := False;
      FindField('TaxDivName').Visible := False;
      FindField('TaxGetDate').Visible := False;
      FindField('TaxNum1').Visible := False;
      FindField('TaxNum2').Visible := False;
      FindField('IsSelfCreated').Visible := False;
      FindField('IdentifyId2').Visible := False;
      FindField('IdentifyId2').Visible := False;

      FindField('RptName1').Visible := False;
      FindField('RptName2').Visible := False;
      FindField('RptName3').Visible := False;
      FindField('RptName4').Visible := False;
      FindField('RptName5').Visible := False;
      FindField('RptName6').Visible := False;

      if RecordCount = 0 then
      begin
        btnEdit.Enabled := False;
        btnDelete.Enabled := False;
      end;      
    end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD01.btnExitClick(Sender: TObject);
begin
  Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD01.btnAppendClick(Sender: TObject);
begin
  dtmMainJ.adoInv001Code.Append;
  cobNotiType.ItemIndex := 0;//?w?]????
  cobIsSelfCreated.ItemIndex := 0;   //?O?_?????t??create?w?]???O
  ChangeBtnStatus;
  medCompID.SetFocus;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD01.btnEditClick(Sender: TObject);
begin
  dtmMainJ.adoInv001Code.Edit;
  ChangeBtnStatus;
  medCompID.Enabled := False;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD01.btnDeleteClick(Sender: TObject);
begin
  if ( dsrInv001.DataSet.FieldByName( 'COMPID' ).AsString = dtmMain.getCompID ) then
  begin
    WarningMsg( '???i?R?????b???????????q?O????, ?????????????????q?A?R???C' );
    Exit;
  end;
  if ConfirmMsg( '?T?{?R?????????q?????' ) then
  begin
    dtmMainJ.adoInv001Code.Delete;
    dtmMainJ.getAllInv001Data;
    ChangeBtnStatus;
  end;
  setButtonCompetence;
  initialForm;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD01.btnCancelClick(Sender: TObject);
begin
  if  dsrInv001.DataSet.State in [dsEdit] then
    medCompID.Enabled := True;
      
  dtmMainJ.adoInv001Code.Cancel;
  dtmMainJ.adoInv001Code.First;
  ChangeBtnStatus;

  setButtonCompetence;
  initialForm;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD01.dsrInv001DataChange(Sender: TObject; Field: TField);
var
  aNotiType, aTaxGetDate, aIsSelfCreated : String;
begin
  with dsrInv001.DataSet do
  begin
    if RecordCount > 0 then
    begin
      medCompID.Text := FieldByName('CompID').AsString;
      medCompSName.Text := FieldByName('CompSName').AsString;
      medCompName.Text := FieldByName('CompName').AsString;

      medBusinessId.Text := FieldByName('BusinessId').AsString;
      medTaxId.Text := FieldByName('TaxId').AsString;
      medManager.Text := FieldByName('Manager').AsString;
      medTel.Text := FieldByName('Tel').AsString;
      medCompAddr.Text := FieldByName('CompAddr').AsString;
      medServiceTypeStr.Text := FieldByName('ServiceTypeStr').AsString;
      medDecaCount.Text := FieldByName('DecaCount').AsString;
      medPayCount.Text := FieldByName('PayCount').AsString;
      medTaxRate.Text := FieldByName('TaxRate').AsString;
      medTaxDivName.Text := FieldByName('TaxDivName').AsString;
      medTaxNum1.Text := FieldByName('TaxNum1').AsString;
      medTaxNum2.Text := FieldByName('TaxNum2').AsString;

      //5764 ?W?[MisOwner For Jackie B08 By Kin 2010/10/14
      edtMisOwner.Text := FieldByName('MisOwner').AsString;
      //5800 ?W?[ SysID By Kin 2010/10/14
      edtSysID.Text := FieldByName('SysID').AsString;

      if FieldByName( 'InvGlyphPath' ).AsString <> '' then
        medGraphPath.Text := FieldByName( 'InvGlyphPath' ).AsString
      else
        medGraphPath.Text :=EmptyStr;
      {#5668 ?W?[EMAIL ???????A By Kin 2010/07/22 }
      cboEmailFileType.ItemIndex :=
        StrToInt( VarToStrDef( FieldByName( 'AttachFileType' ).Value ,'0'));
      {
      if FieldByName( 'AttachFileType' ).AsInteger = 0 then
        cboEmailFileType.ItemIndex := 0
      else
        cboEmailFileType.ItemIndex := 1;
      }

      if FieldByName('LINKTOMIS').AsString = 'Y' then
        cmbLinkToMis.ItemIndex := 0
      else
        cmbLinkToMis.ItemIndex := 1;

      edtConnSeq.Text := FieldByName('CONNSEQ').AsString;          

      aNotiType := FieldByName('NotiType').AsString;
      aTaxGetDate := FieldByName('TaxGetDate').AsString;
      aIsSelfCreated := FieldByName('IsSelfCreated').AsString;

      if aNotiType = '2' then  //1:??(Default) 2:??
        cobNotiType.ItemIndex := 1
      else
        cobNotiType.ItemIndex := 0;

      fraTaxGetDate.setYMD(aTaxGetDate);

      if aIsSelfCreated = 'Y' then
        cobIsSelfCreated.ItemIndex := 0
      else
        cobIsSelfCreated.ItemIndex := 1;

      edtAutoCreateNum.Text := Nvl( FieldByName( 'AutoCreateNum' ).AsString, '0' );
      chkFaciCombine.Checked := ( FieldByName( 'FaciCombine' ).AsInteger = 1 );
      chkShowFaci.Checked := ( FieldByName( 'ShowFaci' ).AsInteger =1 );

      chkStarEInvoice.Checked := ( FieldByName( 'StartEInvoice' ).AsInteger = 1 );
      chkStarEmail.Checked := False;
      chkStarMessage.Checked := False;
      chkStarTVMail.Checked := False;
      chkStarCMSend.Checked := False;
      {#5874 ?W?[?????q?? By Kin 2010/12/10}
      chkNotifyPrize.Checked := False;
      chkMaskInvNo.Checked := False;
      if ( chkStarEInvoice.Checked ) then
      begin
        chkStarEmail.Checked := ( FieldByName( 'StartEMail' ).AsInteger = 1 );
        chkStarMessage.Checked := ( FieldByName( 'StartMessage' ).AsInteger = 1 );
        chkStarTVMail.Checked := ( FieldByName( 'STARTTVMAIL' ).AsInteger = 1 );
        chkStarCMSend.Checked := ( FieldByName( 'STARTCMSEND' ).AsInteger = 1 );
      end;
      chkNotifyPrize.Checked := ( FieldByName( 'StartNotifyPrize' ).AsInteger = 1);
      chkMaskInvNo.Checked := ( FieldByName( 'MaskInvNo' ).AsInteger = 1);
    end else
    begin
      medCompID.Text := EmptyStr;
      medCompSName.Text := EmptyStr;
      medCompName.Text := EmptyStr;
      medBusinessId.Text := EmptyStr;
      medTaxId.Text := EmptyStr;
      medManager.Text := EmptyStr;
      medTel.Text := EmptyStr;
      medCompAddr.Text := EmptyStr;
      medServiceTypeStr.Text := EmptyStr;
      medDecaCount.Text := EmptyStr;
      medPayCount.Text := EmptyStr;
      medTaxRate.Text := EmptyStr;
      medTaxDivName.Text := EmptyStr;
      medTaxNum1.Text := EmptyStr;
      medTaxNum2.Text := EmptyStr;
      //5764 ?W?[MisOwner For Jacky B08 By Kin 2010/10/14
      edtMisOwner.Text := EmptyStr;
      cmbLinkToMis.ItemIndex := 0;
      edtConnSeq.Text := EmptyStr;
      //5800 ?W?[SysID  By Kin 2010/10/14
      edtSysID.Text := EmptyStr;
      cobNotiType.ItemIndex := 0;
      fraTaxGetDate.setYMD( EmptyStr );
      cobIsSelfCreated.ItemIndex := 0;
      edtAutoCreateNum.Text := EmptyStr;
      chkFaciCombine.Checked := False;
      chkShowFaci.Checked := False;
      chkStarEInvoice.Checked := False;
      chkNotifyPrize.Checked := False;
      chkMaskInvNo.Checked := False;
    end;
    //#5760 ?W?[?????????q?l?o???q?? By Kin 2010/09/08
    chkAutoNotify.Checked := False;
    if FieldByName( 'AutoNotify' ).AsString = '1' then
    begin
      chkAutoNotify.Checked := True;
    end;
    SwitchEInvoiceState( chkStarEInvoice.Checked );
    Stt_Show.Caption := IntToStr( RecNo ) + '?@/?@' + IntToStr( RecordCount );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD01.btnSaveClick(Sender: TObject);
var
  aCompId, aCompSName, aCompName, aBusinessId, aTaxId, aManager,
  aTel, aCompAddr, aDecaCount, aPayCount, aTaxRate, aTaxDivName,
  aTaxNum1, aTaxNum2, aNotiType, aTaxGetDate, aIsSelfCreated,
  aSQL, aServiceTypeStr,aMisOwner,aSysId: String;
  aStarEInvoice, aStarEmail, aStarMessage,aStarTVMail,aStarCMSend,aStarAutoNotify: Integer;
  aStartNotifyPrize,aMaskInvNo: Integer;
  aHavePKValue : Boolean;
begin
  if isDataOK then
  begin
    aCompId := Trim(medCompID.Text);
    aCompSName := Trim(medCompSName.Text);
    aCompName := Trim(medCompName.Text);
    aBusinessId := Trim(medBusinessId.Text);
    aTaxId := Trim(medTaxId.Text);
    aManager := Trim(medManager.Text);

    aTel := Trim(medTel.Text);
    aCompAddr := Trim(medCompAddr.Text);
    aServiceTypeStr := UpperCase( Trim(medServiceTypeStr.Text) );
    aDecaCount := Trim(medDecaCount.Text);
    aPayCount := Trim(medPayCount.Text);
    aTaxRate := Trim(medTaxRate.Text);
    aTaxDivName := Trim(medTaxDivName.Text);

    aTaxNum1 := Trim(medTaxNum1.Text);
    aTaxNum2 := Trim(medTaxNum2.Text);

    aMisOwner := Trim( edtMisOwner.Text );
    aSysId := Trim( edtSysID.Text );

    if cobNotiType.ItemIndex = 1 then
      aNotiType := '2'     //1:??(Default) 2:??
    else
      aNotiType := '1';

    aTaxGetDate := Trim(fraTaxGetDate.getYMD);


    if cobIsSelfCreated.ItemIndex = 0 then  //?O?_?????t??create
      aIsSelfCreated := 'Y'
    else
      aIsSelfCreated := 'N';

    aStarEInvoice := 0;
    aStarEmail := 0;
    aStarMessage := 0;
    aStarTVMail := 0;
    aStarCMSend := 0;
    aStarAutoNotify := 0;
    aStartNotifyPrize := 0;
    aMaskInvNo := 0;
    if ( chkStarEInvoice.Checked ) then
    begin
      aStarEInvoice := 1;
      if ( chkStarEmail.Checked ) then aStarEmail := 1;
      if ( chkStarMessage.Checked ) then aStarMessage := 1;
      if ( chkStarTVMail.Checked ) then aStarTVMail := 1;
      if ( chkStarCMSend.Checked ) then aStarCMSend := 1;
    end;
    if ( chkAutoNotify.Checked ) then aStarAutoNotify := 1;

    //#5874 ?W?[?B?n???????o?????X?P?????????q?? By Kin 2010/12/10
    if ( chkNotifyPrize.Checked ) then aStartNotifyPrize := 1;
    if ( chkMaskInvNo.Checked ) then aMaskInvNo := 1;

    //????PK??
    aSQL :=
      ' select * from inv001 where IdentifyId1=''' + IDENTIFYID1 +
      ' '' and IdentifyId2 = ' + IDENTIFYID2 + ' and CompID = ''' + aCompId + '''';

    if  dsrInv001.DataSet.State in [dsInsert] then
      aHavePKValue := dtmMainJ.checkPK(aSQL)
    else
      aHavePKValue := False;


    if aHavePKValue then
    begin
      WarningMsg( '?H?????@??????' );
      medCompID.SetFocus;
    end else
    begin
      dsrInv001.OnDataChange := nil;
      try
        with dsrInv001.DataSet do
        begin
          FieldByName('IdentifyId1').AsString := IDENTIFYID1;
          FieldByName('IdentifyId2').AsString := IDENTIFYID2;

          FieldByName('CompID').AsString := aCompId;
          FieldByName('CompSName').AsString := aCompSName;
          FieldByName('CompName').AsString := aCompName;

          FieldByName('BusinessId').AsString := aBusinessId;
          FieldByName('TaxId').AsString := aTaxId;
          FieldByName('Manager').AsString := aManager;
          FieldByName('Tel').AsString := aTel;
          FieldByName('CompAddr').AsString := aCompAddr;
          FieldByName('ServiceTypeStr').AsString := aServiceTypeStr;
          FieldByName('DecaCount').AsString := aDecaCount;
          FieldByName('PayCount').AsString := aPayCount;
          FieldByName('TaxRate').AsString := aTaxRate;
          FieldByName('TaxDivName').AsString := aTaxDivName;
          FieldByName('TaxNum1').AsString := aTaxNum1;
          FieldByName('TaxNum2').AsString := aTaxNum2;
          FieldByName( 'InvGlyphPath' ).AsString := medGraphPath.Text;
          // #5668 ?W?[Email???????A By Kin 2010/07/22
          FieldByName( 'AttachFileType' ).AsInteger := cboEmailFileType.ItemIndex;
          FieldByName('NotiType').AsString := aNotiType;

          if aTaxGetDate <> EMPTY_DATE_STR then
            FieldByName('TaxGetDate').AsString := aTaxGetDate;

          FieldByName('IsSelfCreated').AsString := aIsSelfCreated;

          FieldByName('UptTime').AsString := DateTimeToStr(now);
          FieldByName('UptEn').AsString := dtmMain.getLoginUser;

          if cmbLinkToMis.ItemIndex = 0 then
            FieldByName( 'LINKTOMIS' ).AsString := 'Y'
          else
            FieldByName( 'LINKTOMIS' ).AsString := 'N';

          FieldByName( 'CONNSEQ' ).AsString := edtConnSeq.Text;

          FieldByName( 'AUTOCREATENUM' ).AsString := Nvl( edtAutoCreateNum.Text, '0' );

          FieldByName( 'FACICOMBINE' ).AsInteger := IIF( chkFaciCombine.Checked, 1, 0 );
          
          //#5358 ?W?[ShowFaci ?P?_?b?}???o?????O?_?n?N?]???????g?JMemo1 By Kin 2009/12/14
          FieldByName( 'ShowFaci' ).AsInteger := IIF(chkShowFaci.Checked,1,0);

          FieldByName( 'StartEInvoice' ).AsInteger := aStarEInvoice;
          FieldByName( 'StartEmail' ).AsInteger := aStarEmail;
          FieldByName( 'StartMessage' ).AsInteger := aStarMessage;
          FieldByName( 'STARTTVMAIL' ).AsInteger := aStarTVMail;
          FieldByName( 'STARTCMSEND' ).AsInteger := aStarCMSend;
          FieldByName( 'AutoNotify' ).AsInteger := aStarAutoNotify;
          //5764 ?W?[MisOwner For Jacky B08 By Kin 2010/10/14
          FieldByName( 'MisOwner' ).AsString := aMisOwner;
          //5800 ?W?[ SysID By Kin 2010/10/14
          FieldByName( 'SysID' ).AsString := aSysId;

          //#5874 ?W?[?q???P?B?n CheckBox By Kin 2010/12/10
          FieldByName( 'MaskInvNo' ).AsInteger := aMaskInvNo;
          FieldByName( 'StartNotifyPrize' ).AsInteger := aStartNotifyPrize;
          dsrInv001.DataSet.Post;
        end;
      finally
        dsrInv001.OnDataChange := dsrInv001DataChange;
      end;
      { ?Y?O???????O?????n?J?????q?O, ?h???Y???s }
      if ( aCompId = dtmMain.getCompID ) then
      begin
        dtmMain.sG_ServiceTypeStr := UpperCase( aServiceTypeStr );
        dtmMain.sG_CompName := aCompSName;
        dtmMain.sG_CompLongName := aCompName;
        dtmMain.sG_LinkToMIS :=
          ( dsrInv001.DataSet.FieldByName( 'LINKTOMIS' ).AsString = 'Y' );
        dtmMain.sG_ConnSeq :=
          dsrInv001.DataSet.FieldByName( 'CONNSEQ' ).AsInteger;
        dtmMain.sG_AutoCreateNum :=
          dsrInv001.DataSet.FieldByName( 'AUTOCREATENUM' ).AsInteger;
        dtmMain.sG_FaciCombine :=
          dsrInv001.DataSet.FieldByName( 'FACICOMBINE' ).AsInteger;
        dtmMain.sG_StarEInvoice :=
          ( dsrInv001.DataSet.FieldByName( 'STARTEINVOICE' ).AsInteger = 1 );
        dtmMain.sG_StarEmail :=
          ( dsrInv001.DataSet.FieldByName( 'STARTEMAIL' ).AsInteger = 1 );
        dtmMain.sG_StarMessage :=
          ( dsrInv001.DataSet.FieldByName( 'STARTMESSAGE' ).AsInteger = 1 );
        dtmMain.sG_StarTVmail :=
          ( dsrInv001.DataSet.FieldByName( 'STARTTVMAIL' ).AsInteger = 1 );
        dtmMain.sG_StarCMSend :=
          ( dsrInv001.DataSet.FieldByName( 'STARTCMSEND' ).AsInteger = 1 );
        dtmMain.sG_StarAutoNotify :=
          ( dsrInv001.DataSet.FieldByName( 'AutoNotify' ).AsInteger = 1 );
      end;
    end;
    dtmMainJ.getAllInv001Data;
    medCompID.Enabled := True;
    ChangeBtnStatus;
    setButtonCompetence;
    initialForm;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvD01.isDataOK: Boolean;
var
  aMsg, sL_CompID,sL_CompSName,sL_CompName,sL_BusinessId : String;
  nL_BusinessIdLen : Integer;
begin
    Result := True;
    sL_CompID := Trim(medCompID.Text);
    sL_CompSName := Trim(medCompSName.Text);
    sL_CompName := Trim(medCompName.Text);

    if sL_CompID = '' then
    begin
      WarningMsg('?????J???q?N?X');
      medCompID.SetFocus;
      Result := False;
      Exit;
    end;

    if sL_CompSName = '' then
    begin   begin

    end;
      WarningMsg( '?????J???q????' );
      medCompSName.SetFocus;
      Result := False;
      Exit;
    end;


    if sL_CompName = '' then
    begin
      WarningMsg( '?????J???q???W' );
      medCompName.SetFocus;
      Result := False;
      Exit;
    end;

    if edtConnSeq.Text = EmptyStr then
    begin
      WarningMsg( '?????J???A?t???????s??' );
      edtConnSeq.SetFocus;
      Result := False;
      Exit;
    end;  


    sL_BusinessId := Trim(medBusinessId.Text);

    if sL_BusinessId <> '' then
    begin
      nL_BusinessIdLen := Length(sL_BusinessId);
      if nL_BusinessIdLen = 8 then
      begin
        if not dtmMain.checkBusinessID(sL_BusinessId) then
        begin
          WarningMsg( '???@?s?????~' );
          medBusinessId.SetFocus;
          Result := False;
          Exit;
        end;
      end
      else
      begin
        WarningMsg( '???@?s?????~' );
        medBusinessId.SetFocus;
        Result := False;
        Exit;
      end;
    end;

    if not ValidateServiceType( medServiceTypeStr.Text, aMsg ) then
    begin
      WarningMsg( aMsg );
      Result := False;
      Exit;
    end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD01.medBusinessIdExit(Sender: TObject);
var
  aBusinessId : String;
  aBusinessIdLen : Integer;
begin
  aBusinessId := Trim( medBusinessId.Text );
  if ( aBusinessId = EmptyStr ) then Exit;
  aBusinessIdLen := Length( aBusinessId );
  if ( aBusinessIdLen <> 8 ) then
  begin
    ErrorMsg( '???@?s?????J???????~?C' );
    medBusinessId.SetFocus;
    Exit;
  end;
  if not ( dtmMain.checkBusinessID( aBusinessId ) ) then
  begin
    ErrorMsg( '???@?s?????????~?C' );
    medBusinessId.SetFocus;
    Exit;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD01.medCompSNameEnter(Sender: TObject);
begin
  medCompName.Text := medCompSName.Text;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD01.medServiceTypeStrExit(Sender: TObject);
begin
  medServiceTypeStr.Text := UpperCase(medServiceTypeStr.Text);
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvD01.ValidateServiceType(aServiceType: String; var aMsg: String): Boolean;
var
  aIndex, aIndex2: Integer;
  aSplit: array of string;
  aString: String;
  aDataSet: TADOQuery;
begin
  Result := False;
  aMsg := '';
  if ( Trim( aServiceType ) = '' ) then
  begin
    aMsg := '?????J?~???A??????';
    Exit;
  end;
  if Pos( ',', aServiceType ) = 0 then
  begin
    SetLength( aSplit, Length( aSplit ) + 1 );
    aSplit[ Length( aSplit ) - 1 ] := aServiceType;
  end
  else begin
    aString := '';
    for aIndex := 1 to Length( aServiceType ) do
    begin
      if ( aServiceType[aIndex] <> ',' ) then
      begin
        aString := ( aString + aServiceType[aIndex] );
      end
      else begin
          SetLength( aSplit, Length( aSplit ) + 1 );
          aSplit[ Length( aSplit ) - 1 ] := aString;
          aString := '';
      end;
    end;
    if ( aString <> '' ) then
    begin
      SetLength( aSplit, Length( aSplit ) + 1 );
      aSplit[ Length( aSplit ) - 1 ] := aString;
    end;
  end;
  for aIndex := Low( aSplit ) to High( aSplit ) do
  begin
    if ( aSplit[aIndex] = '' ) then
    begin
      aMsg := '?~???A?????????J???~, ?U?A???O?????H?r?I?j?}, ?B???????J?C'#13#10 +
              '???p: D,S';
      Exit;
    end;
    for aIndex2 := 1 to Length( aSplit[aIndex] ) do
    begin
      if not ( UpCase( aSplit[aIndex][aIndex2] ) in ['A'..'Z', '0'..'9'] ) then
      begin
        aMsg := '?~???A?????????J???~, ?A???O???J???u?????^???r?C'#13#10 +
                '???p: D,S';
        Exit;
      end;
    end;
    { ???d?A???O?]?w???O?_?????A???O}
    aDataSet := dtmMain.adoComm;
    aDataSet.Close;
    for aIndex2 := 1 to Length( aSplit[aIndex] ) do
    begin
      aDataSet.SQL.Text := Format(
        ' SELECT COUNT(1) AS COUNTS FROM INV022 ' +
        '  WHERE IDENTIFYID1 = ''%s''           ' +
        '    AND IDENTIFYID2 = %s               ' +
        '    AND ITEMID = ''%s''                ',
        [ IDENTIFYID1, IDENTIFYID2, UpCase( aSplit[aIndex][aIndex2] ) ] );
      aDataSet.Open;
      if aDataSet.FieldByName( 'COUNTS' ).AsInteger = 0 then
        aMsg := Format( '?~???A?????????J???~, ?A???????????]?w?????|???]?w???A???O: %s',
          [ UpCase( aSplit[aIndex][aIndex2] ) ] );
      aDataSet.Close;
      if aMsg <> EmptyStr then Exit;
    end;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD01.SwitchEInvoiceState(aValue: Boolean);
begin
  if ( aValue ) then
  begin
    chkStarEmail.Enabled := True;
    chkStarMessage.Enabled := True;
    chkStarTVMail.Enabled := True;
    chkStarCMSend.Enabled := True;
  end else
  begin
    chkStarEmail.Checked := False;
    chkStarMessage.Checked := False;
    chkStarTVMail.Checked := False;
    chkStarCMSend.Checked := False;
    chkStarEmail.Enabled := False;
    chkStarMessage.Enabled := False;
    chkStarTVMail.Enabled := False;
    chkStarCMSend.Enabled := False;



  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD01.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  dtmMain.adoComm.Close;
  dtmMain.adoComm.SQL.Text :=
    ' SELECT COUNT(1) COUNTS FROM INV001 ';
  dtmMain.adoComm.Open;
  CanClose := ( dtmMain.adoComm.FieldByName( 'COUNTS' ).AsInteger > 0 );
  dtmMain.adoComm.Close;
  if not CanClose then
  begin
    WarningMsg( '?????????J?@?????q????????' );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD01.edtConnSeqKeyPress(Sender: TObject; var Key: Char);
begin
  if not ( Key in ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9'] ) then
    Key := #0;    
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvD01.chkStarEInvoicePropertiesEditValueChanged(
  Sender: TObject);
begin
  SwitchEInvoiceState( chkStarEInvoice.Checked );
end;

{ ---------------------------------------------------------------------------- }

end.

