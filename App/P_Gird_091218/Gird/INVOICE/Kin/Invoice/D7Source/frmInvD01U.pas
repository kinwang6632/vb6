unit frmInvD01U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Buttons, DB, Mask, DBCtrls, Grids, DBGrids,
  ADODB, fraYMDU, cxControls, cxContainer, cxEdit, cxTextEdit, cxMaskEdit,
  cxCheckBox, dxSkinsCore, dxSkinsDefaultPainters, dxSkinBlack, dxSkinBlue,
  dxSkinCaramel, dxSkinCoffee;

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
  private
    { Private declarations }
    procedure ChangeBtnStatus;
    procedure setButtonCompetence;
    function isDataOK : Boolean;
    procedure initialForm;
    function ValidateServiceType(aServiceType: String;  var aMsg: String): Boolean;
  public
    { Public declarations }
  end;

var
  frmInvD01: TfrmInvD01;

implementation

uses cbUtilis, frmMainU, dtmMainJU, dtmMainU;

{$R *.dfm}

procedure TfrmInvD01.FormShow(Sender: TObject);
begin
    self.Caption := frmMain.GetFormTitleString( 'D01', '???q???????@' );

    if dtmMainJ.getAllInv001Data = 0 then
      WarningMsg(  '?L?????C' );

    ChangeBtnStatus;
    setButtonCompetence;
    initialForm;
end;

procedure TfrmInvD01.btnExitClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmInvD01.ChangeBtnStatus;
begin
     with dtmMainJ.adoInv001Code do
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
       end
       else
       begin
          if dtmMainJ.adoInv001Code.RecordCount>0 then
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
       end;
     end;
end;

procedure TfrmInvD01.btnAppendClick(Sender: TObject);
begin
    dtmMainJ.adoInv001Code.Append;

    cobNotiType.ItemIndex := 0;//?w?]????
    cobIsSelfCreated.ItemIndex := 0;   //?O?_?????t??create?w?]???O

    ChangeBtnStatus;
    medCompID.SetFocus;
end;

procedure TfrmInvD01.btnEditClick(Sender: TObject);
begin
    dtmMainJ.adoInv001Code.Edit;
    ChangeBtnStatus;
    medCompID.Enabled := false;
end;

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

procedure TfrmInvD01.btnCancelClick(Sender: TObject);
begin
    if  dsrInv001.DataSet.State in [dsEdit] then
      medCompID.Enabled := true;
      
    dtmMainJ.adoInv001Code.Cancel;
    dtmMainJ.adoInv001Code.First;
    ChangeBtnStatus;

    setButtonCompetence;
    initialForm;
end;

procedure TfrmInvD01.dsrInv001DataChange(Sender: TObject; Field: TField);
var
    sL_NotiType,sL_TaxGetDate,sL_IsSelfCreated : String;
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

        if FieldByName('LINKTOMIS').AsString = 'Y' then
          cmbLinkToMis.ItemIndex := 0
        else
          cmbLinkToMis.ItemIndex := 1;

        edtConnSeq.Text := FieldByName('CONNSEQ').AsString;          

        sL_NotiType := FieldByName('NotiType').AsString;
        sL_TaxGetDate := FieldByName('TaxGetDate').AsString;
        sL_IsSelfCreated := FieldByName('IsSelfCreated').AsString;

        if sL_NotiType = '2' then  //1:??(Default) 2:??
          cobNotiType.ItemIndex := 1
        else
          cobNotiType.ItemIndex := 0;

        fraTaxGetDate.setYMD(sL_TaxGetDate);

        if sL_IsSelfCreated = 'Y' then
          cobIsSelfCreated.ItemIndex := 0
        else
          cobIsSelfCreated.ItemIndex := 1;

        edtAutoCreateNum.Text := Nvl( FieldByName( 'AutoCreateNum' ).AsString, '0' );
        chkFaciCombine.Checked := ( FieldByName( 'FaciCombine' ).AsInteger = 1 );
        chkShowFaci.Checked := ( FieldByName( 'ShowFaci' ).AsInteger =1 );
      end
      else
      begin
        medCompID.Text := '';
        medCompSName.Text := '';
        medCompName.Text := '';

        medBusinessId.Text := '';
        medTaxId.Text := '';
        medManager.Text := '';
        medTel.Text := '';
        medCompAddr.Text := '';
        medServiceTypeStr.Text := '';
        medDecaCount.Text := '';
        medPayCount.Text := '';
        medTaxRate.Text := '';
        medTaxDivName.Text := '';
        medTaxNum1.Text := '';
        medTaxNum2.Text := '';

        cmbLinkToMis.ItemIndex := 0;
        edtConnSeq.Text := EmptyStr;

        cobNotiType.ItemIndex := 0;
        fraTaxGetDate.setYMD('');
        cobIsSelfCreated.ItemIndex := 0;
        edtAutoCreateNum.Text := EmptyStr;
        chkFaciCombine.Checked := False;
        chkShowFaci.Checked := False;
      end;
      Stt_Show.Caption := inttostr(RecNo)+'?@/?@'+inttostr(recordcount);      
    end;
end;

procedure TfrmInvD01.btnSaveClick(Sender: TObject);
var
    sL_CompID,sL_CompSName,sL_CompName,sL_BusinessId,sL_TaxId,sL_Manager : String;
    sL_Tel,sL_CompAddr,sL_DecaCount,sL_PayCount,sL_TaxRate,sL_TaxDivName : String;
    sL_TaxNum1,sL_TaxNum2,sL_NotiType,sL_TaxGetDate,sL_IsSelfCreated,sL_SQL : String;
    sL_ServiceTypeStr : String;
    bL_HavePKValue : Boolean;
begin
    if isDataOK then
    begin
      sL_CompID := Trim(medCompID.Text);
      sL_CompSName := Trim(medCompSName.Text);
      sL_CompName := Trim(medCompName.Text);
      sL_BusinessId := Trim(medBusinessId.Text);
      sL_TaxId := Trim(medTaxId.Text);
      sL_Manager := Trim(medManager.Text);

      sL_Tel := Trim(medTel.Text);
      sL_CompAddr := Trim(medCompAddr.Text);
      sL_ServiceTypeStr := UpperCase( Trim(medServiceTypeStr.Text) );
      sL_DecaCount := Trim(medDecaCount.Text);
      sL_PayCount := Trim(medPayCount.Text);
      sL_TaxRate := Trim(medTaxRate.Text);
      sL_TaxDivName := Trim(medTaxDivName.Text);

      sL_TaxNum1 := Trim(medTaxNum1.Text);
      sL_TaxNum2 := Trim(medTaxNum2.Text);



      if cobNotiType.ItemIndex = 1 then
        sL_NotiType := '2'     //1:??(Default) 2:??
      else
        sL_NotiType := '1';

      sL_TaxGetDate := Trim(fraTaxGetDate.getYMD);


      if cobIsSelfCreated.ItemIndex = 0 then  //?O?_?????t??create
        sL_IsSelfCreated := 'Y'
      else
        sL_IsSelfCreated := 'N';

      //????PK??
      sL_SQL := 'select * from inv001 where IdentifyId1=''' + IDENTIFYID1 +
                ''' and IdentifyId2=' + IDENTIFYID2 + ' and CompID=''' + sL_CompID + '''';

      if  dsrInv001.DataSet.State in [dsInsert] then
        bL_HavePKValue := dtmMainJ.checkPK(sL_SQL)
      else
        bL_HavePKValue := false;


      if bL_HavePKValue then
      begin
        WarningMsg( '?H?????@??????' );
        medCompID.SetFocus;
      end
      else
      begin
        dsrInv001.OnDataChange := nil;
        try
          with dsrInv001.DataSet do
          begin
            FieldByName('IdentifyId1').AsString := IDENTIFYID1;
            FieldByName('IdentifyId2').AsString := IDENTIFYID2;

            FieldByName('CompID').AsString := sL_CompID;
            FieldByName('CompSName').AsString := sL_CompSName;
            FieldByName('CompName').AsString := sL_CompName;

            FieldByName('BusinessId').AsString := sL_BusinessId;
            FieldByName('TaxId').AsString := sL_TaxId;
            FieldByName('Manager').AsString := sL_Manager;
            FieldByName('Tel').AsString := sL_Tel;
            FieldByName('CompAddr').AsString := sL_CompAddr;
            FieldByName('ServiceTypeStr').AsString := sL_ServiceTypeStr;
            FieldByName('DecaCount').AsString := sL_DecaCount;
            FieldByName('PayCount').AsString := sL_PayCount;
            FieldByName('TaxRate').AsString := sL_TaxRate;
            FieldByName('TaxDivName').AsString := sL_TaxDivName;
            FieldByName('TaxNum1').AsString := sL_TaxNum1;
            FieldByName('TaxNum2').AsString := sL_TaxNum2;

            FieldByName('NotiType').AsString := sL_NotiType;

            if sL_TaxGetDate <> EMPTY_DATE_STR then
              FieldByName('TaxGetDate').AsString := sL_TaxGetDate;

            FieldByName('IsSelfCreated').AsString := sL_IsSelfCreated;

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
            dsrInv001.DataSet.Post;
          end;
        finally
          dsrInv001.OnDataChange := dsrInv001DataChange;
        end;
          { ?Y?O???????O?????n?J?????q?O, ?h???Y???s }
          if ( sL_CompID = dtmMain.getCompID ) then
          begin
            dtmMain.sG_ServiceTypeStr := UpperCase( sL_ServiceTypeStr );
            dtmMain.sG_CompName := sL_CompSName;
            dtmMain.sG_CompLongName := sL_CompName;
            dtmMain.sG_LinkToMIS :=
              ( dsrInv001.DataSet.FieldByName( 'LINKTOMIS' ).AsString = 'Y' );
            dtmMain.sG_ConnSeq :=
              dsrInv001.DataSet.FieldByName( 'CONNSEQ' ).AsInteger;
            dtmMain.sG_AutoCreateNum :=
              dsrInv001.DataSet.FieldByName( 'AUTOCREATENUM' ).AsInteger;
            dtmMain.sG_FaciCombine :=
              dsrInv001.DataSet.FieldByName( 'FACICOMBINE' ).AsInteger;
          end;
        end;
        dtmMainJ.getAllInv001Data;
        medCompID.Enabled := true;
        ChangeBtnStatus;
        setButtonCompetence;
        initialForm;
    end;
end;

procedure TfrmInvD01.setButtonCompetence;
var
    sL_QueryCompetence,sL_InsertCompetence,sL_DeleteCompetence,sL_UpdateCompetence,sL_ExecuteCompetence : String;
begin
     //?]?w?v??
     dtmMain.getCompetence('D01',sL_QueryCompetence,sL_InsertCompetence,sL_DeleteCompetence,sL_UpdateCompetence,sL_ExecuteCompetence);

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

{ ---------------------------------------------------------------------------- }

function TfrmInvD01.isDataOK: Boolean;
var
  aMsg, sL_CompID,sL_CompSName,sL_CompName,sL_BusinessId : String;
  nL_BusinessIdLen : Integer;
begin
    Result := true;
    sL_CompID := Trim(medCompID.Text);
    sL_CompSName := Trim(medCompSName.Text);
    sL_CompName := Trim(medCompName.Text);

    if sL_CompID = '' then
    begin
      WarningMsg('?????J???q?N?X');
      medCompID.SetFocus;
      Result := false;
      Exit;
    end;

    if sL_CompSName = '' then
    begin   begin

    end;
      WarningMsg( '?????J???q????' );
      medCompSName.SetFocus;
      Result := false;
      Exit;
    end;


    if sL_CompName = '' then
    begin
      WarningMsg( '?????J???q???W' );
      medCompName.SetFocus;
      Result := false;
      Exit;
    end;

    if edtConnSeq.Text = EmptyStr then
    begin
      WarningMsg( '?????J???A?t???????s??' );
      edtConnSeq.SetFocus;
      Result := false;
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
          Result := false;
          Exit;
        end;
      end
      else
      begin
        WarningMsg( '???@?s?????~' );
        medBusinessId.SetFocus;
        Result := false;
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

procedure TfrmInvD01.initialForm;
begin
    //#5358 ?W?[ShowFaci ?P?_?b?}???o?????O?_?n?N?]???????g?JMemo1 By Kin 2009/12/14
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

      FindField('CompSName').DisplayWidth := 30;
      FindField('Manager').DisplayWidth := 20;
      FindField('UptEn').DisplayWidth := 10;

      FindField('IdentifyId1').Visible := false;
      FindField('IdentifyId2').Visible := false;
      FindField('CompName').Visible := false;
      FindField('TaxId').Visible := false;
      FindField('Tel').Visible := false;
      FindField('CompAddr').Visible := false;
      FindField('NotiType').Visible := false;
      FindField('DecaCount').Visible := false;
      FindField('PayCount').Visible := false;
      FindField('TaxRate').Visible := false;
      FindField('TaxDivName').Visible := false;
      FindField('TaxGetDate').Visible := false;
      FindField('TaxNum1').Visible := false;
      FindField('TaxNum2').Visible := false;
      FindField('IsSelfCreated').Visible := false;
      FindField('IdentifyId2').Visible := false;
      FindField('IdentifyId2').Visible := false;

      FindField('RptName1').Visible := false;
      FindField('RptName2').Visible := false;
      FindField('RptName3').Visible := false;
      FindField('RptName4').Visible := false;
      FindField('RptName5').Visible := false;
      FindField('RptName6').Visible := false;

      if RecordCount = 0 then
      begin
        btnEdit.Enabled := false;
        btnDelete.Enabled := false;
      end;      
    end;
end;


procedure TfrmInvD01.medBusinessIdExit(Sender: TObject);
var
    sL_BusinessId : String;
    nL_BusinessIdLen : Integer;
begin
    sL_BusinessId := Trim(medBusinessId.Text);

    if sL_BusinessId <> '' then
    begin
      nL_BusinessIdLen := Length(sL_BusinessId);
      if nL_BusinessIdLen = 8 then
      begin
        if not dtmMain.checkBusinessID(sL_BusinessId) then
        begin
          MessageDlg('???@?s?????~',mtError,[mbOk],0);
          medBusinessId.SetFocus;
        end;
      end
      else
      begin
        MessageDlg('???@?s?????~',mtError,[mbOk],0);
        medBusinessId.SetFocus;
      end;
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

function TfrmInvD01.ValidateServiceType(aServiceType: String;
  var aMsg: String): Boolean;
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

end.

