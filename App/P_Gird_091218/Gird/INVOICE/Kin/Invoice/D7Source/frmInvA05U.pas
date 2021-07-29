unit frmInvA05U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, fraYMDU, Mask, Buttons, dxSkinsCore,
  dxSkinsDefaultPainters, cxControls, cxContainer, cxEdit, cxTextEdit,
  cxMaskEdit, cxButtonEdit, dxSkinBlack, dxSkinBlue, dxSkinCaramel,
  dxSkinCoffee;

type
  TfrmInvA05 = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    Label1: TLabel;
    fraEInvDate: TfraYMD;
    fraSInvDate: TfraYMD;
    Label2: TLabel;
    rgpOrder: TRadioGroup;
    Label3: TLabel;
    medSInvID: TMaskEdit;
    rgpTransferType: TRadioGroup;
    btnOK: TBitBtn;
    btnExit: TBitBtn;
    medEInvID: TMaskEdit;
    cobBusiness: TComboBox;
    Label4: TLabel;
    Label5: TLabel;
    txtInvUseItem: TcxButtonEdit;
    chkInvUse: TCheckBox;
    procedure btnExitClick(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure medSInvIDExit(Sender: TObject);
    procedure fraSInvDateExit(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure medEInvIDExit(Sender: TObject);
    procedure chkInvUseClick(Sender: TObject);
    procedure txtInvUseItemPropertiesButtonClick(Sender: TObject;
      AButtonIndex: Integer);
  private
    { Private declarations }
    function isDataOK : Boolean;
    procedure DoSelectInvUseItem;
  public
    { Public declarations }
  end;

var
  frmInvA05: TfrmInvA05;

implementation

uses cbUtilis, frmInvA05_1U, frmInvA05_2U, frmMainU,  dtmMainJU, dtmMainU,
  frmMultiSelectU;

var
  aQueryParam: TQueryParam;

{$R *.dfm}

procedure TfrmInvA05.btnExitClick(Sender: TObject);
begin
    Close;
end;

function TfrmInvA05.isDataOK: Boolean;
var
    sL_SInvID,sL_EInvID,sL_SInvDate,sL_EInvDate : String;
begin
    sL_SInvID := Trim(medSInvID.Text);
    sL_EInvID := Trim(medEInvID.Text);
    sL_SInvDate := Trim(fraSInvDate.getYMD);
    sL_EInvDate := Trim(fraEInvDate.getYMD);

    if (sL_SInvID='') and (sL_SInvDate=EMPTY_DATE_STR) then
    begin
      MessageDlg('�п�J�䤤�@���d�߱���',mtError,[mbOk],0);
      medSInvID.SetFocus;
      Result := false;
      Exit;
    end
    else
    begin
      if sL_SInvID<>'' then
      begin
        if Length(sL_SInvID) <> 10 then
        begin
          MessageDlg('��J�o�����X���~',mtError,[mbOk],0);
          medSInvID.SetFocus;
          Result := false;
          Exit;
        end;

        if Length(sL_EInvID) <> 10 then
        begin
          MessageDlg('��J�o�����X���~',mtError,[mbOk],0);
          medEInvID.SetFocus;
          Result := false;
          Exit;
        end;

        if Copy(sL_SInvID,1,2) <> Copy(sL_EInvID,1,2) then
        begin
          MessageDlg('�����P�@�o���r�y',mtError,[mbOk],0);
          medEInvID.SetFocus;
          Result := false;
          Exit;
        end;

        if Copy(sL_SInvID,3,8) > Copy(sL_EInvID,3,8) then
        begin
          MessageDlg('�o�����X��J���ǿ��~!',mtError,[mbOk],0);
          medEInvID.SetFocus;
          Result := false;
          Exit;
        end;

        try
          StrToInt(Copy(sL_SInvID,3,8));
        except
          MessageDlg('�o�����X�榡���~!',mtError,[mbOk],0);
          medSInvID.SetFocus;
          Result := false;
          Exit;
        end;


        try
          StrToInt(Copy(sL_EInvID,3,8));
        except
          MessageDlg('�o�����X�榡���~!',mtError,[mbOk],0);
          medEInvID.SetFocus;
          Result := false;
          Exit;
        end;
      end;




      if sL_SInvDate<>EMPTY_DATE_STR then
      begin
        if sL_EInvDate=EMPTY_DATE_STR then
        begin
          MessageDlg('�п�J�o���������',mtError,[mbOk],0);
          fraEInvDate.mseYMD.SetFocus;
          Result := false;
          Exit;
        end;

        if sL_EInvDate < sL_SInvDate then
        begin
          MessageDlg('������������j�󵥩�}�l���',mtError,[mbOk],0);
          fraEInvDate.mseYMD.SetFocus;
          Result := false;
          Exit;
        end;
      end;
    end;

    Result := true;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA05.btnOKClick(Sender: TObject);
var
    sL_SInvID,sL_EInvID,sL_SInvDate,sL_EInvDate : String;
    sL_OrderByType,sL_TransferType,sL_BusinessType : String;
    aExecCount: Integer;
begin
  if isDataOK then
  begin

    sL_SInvID := UpperCase(Trim(medSInvID.Text));
    sL_EInvID := UpperCase(Trim(medEInvID.Text));
    sL_SInvDate := Trim(fraSInvDate.getYMD);
    sL_EInvDate := Trim(fraEInvDate.getYMD);

    if sL_SInvDate = EMPTY_DATE_STR then
    begin
      sL_SInvDate := '';
      sL_EInvDate := '';
    end;

    if rgpOrder.ItemIndex = 0 then
      sL_OrderByType := '1'  //�o�����X
    else if rgpOrder.ItemIndex = 1 then
      sL_OrderByType := '2'  //�o�����
    else if rgpOrder.ItemIndex = 2 then
      sL_OrderByType := '3'; //�Ȥ�s��

    if cobBusiness.ItemIndex = 0 then
      sL_BusinessType := '0'
    else if cobBusiness.ItemIndex = 1 then
      sL_BusinessType := '1'   //��~�H
    else if cobBusiness.ItemIndex = 2 then
      sL_BusinessType := '2';  //�D��~�H

    if rgpTransferType.ItemIndex = 0 then
      sL_TransferType := '1'  //�o���C�L
    else if rgpTransferType.ItemIndex = 1 then
      sL_TransferType := '2';  //��J��r��

    Screen.Cursor := crSQLWait;
    try
      // ���ͦC�L�o�����
      dtmMainJ.getPrintInvoiceTempData( sL_SInvID,
        sL_EInvID, sL_SInvDate, sL_EInvDate, sL_BusinessType );
      // �R���䥦���ŦX���󪺸��
      dtmMainJ.FilterPrintInvoiceTempData( chkInvUse.Checked,
        aQueryParam.Param1 );
    finally
      Screen.Cursor := crDefault;
    end;

    dtmMainJ.adoQryCommon.Close;
    dtmMainJ.adoQryCommon.SQL.Text := Format(
      ' SELECT COUNT(1) AS COUNTS FROM INV023 A  ' +
      '    WHERE A.IDENTIFYID1 = ''%s''          ' +
      '      AND A.IDENTIFYID2 = ''%s''          ' +
      '      AND A.COMPID = ''%s''               ' +
      '      AND A.OWNER = ''%s''                ',
      [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID, dtmMain.getLoginUserName] );
    dtmMainJ.adoQryCommon.Open;
    aExecCount := dtmMainJ.adoQryCommon.FieldByName( 'COUNTS' ).AsInteger;
    dtmMainJ.adoQryCommon.Close;

    if ( aExecCount <= 0 ) then
    begin
      WarningMsg( '���d�߽d�򤺵L�o����ƥi�ѮM�L�C' );
      if ( medSInvID.CanFocus ) then medSInvID.SetFocus;
      Exit;
    end;

    if sL_TransferType = '1' then
    begin
      frmInvA05_1 := TfrmInvA05_1.Create(Application);
      try
        frmInvA05_1.sG_UserID := dtmMain.getLoginUserName;
        frmInvA05_1.sG_CompID := dtmMain.getCompID;
        frmInvA05_1.sG_OrderByType := sL_OrderByType;
        frmInvA05_1.CallFromProgram := Self.ClassType;
        frmInvA05_1.InvId := EmptyStr;
        frmInvA05_1.ShowModal;
      finally
        frmInvA05_1.Free;
        frmInvA05_1:=nil;
      end;  
    end
    else if sL_TransferType = '2' then
    begin
      frmInvA05_2 := TfrmInvA05_2.Create(Application);
      frmInvA05_2.sG_UserID := dtmMain.getLoginUserName;
      frmInvA05_2.sG_CompID := dtmMain.getCompID;
      frmInvA05_2.ShowModal;
      frmInvA05_2:=nil;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA05.medSInvIDExit(Sender: TObject);
begin
  medSInvID.Text := UpperCase(Trim(medSInvID.Text));
  if Trim( medEInvID.Text ) = EmptyStr then
    medEInvID.Text := UpperCase(Trim(medSInvID.Text));
end;

procedure TfrmInvA05.fraSInvDateExit(Sender: TObject);
begin
    fraEInvDate.setYMD(Trim(fraSInvDate.getYMD));
end;

procedure TfrmInvA05.FormShow(Sender: TObject);
begin
  Self.Caption := frmMain.GetFormTitleString( 'A05', '�q�l�p����o���M�L' );
end;

procedure TfrmInvA05.medEInvIDExit(Sender: TObject);
begin
    medEInvID.Text := UpperCase(Trim(medEInvID.Text));
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA05.chkInvUseClick(Sender: TObject);
begin
  if ( chkInvUse.Checked ) then
  begin
    cobBusiness.Enabled := False;
    cobBusiness.ItemIndex := 2;
    txtInvUseItem.Clear;
    txtInvUseItem.Enabled := False;
    aQueryParam.Param1 := EmptyStr;
  end else
  begin
    cobBusiness.Enabled := True;
    cobBusiness.ItemIndex := 0;
    txtInvUseItem.Clear;
    txtInvUseItem.Enabled := True;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA05.txtInvUseItemPropertiesButtonClick(Sender: TObject;
  AButtonIndex: Integer);
begin
  if AButtonIndex = 1 then
    DoSelectInvUseItem
  else begin
    txtInvUseItem.Clear;
    aQueryParam.Param1 := EmptyStr;
   end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA05.DoSelectInvUseItem;
begin
  frmMultiSelect := TfrmMultiSelect.Create( Application );
  try
    frmMultiSelect.Connection := dtmMain.InvConnection;
    frmMultiSelect.KeyFields := 'ITEMID';
    frmMultiSelect.KeyValues := aQueryParam.Param1;
    frmMultiSelect.DisplayFields := 'ITEMID,�N�X,DESCRIPTION,�o���γ~';
    frmMultiSelect.ResultFields := 'DESCRIPTION';
    frmMultiSelect.SQL.Text := Format(
      ' SELECT ITEMID, DESCRIPTION FROM INV028 ' +
      '  WHERE IDENTIFYID1 = ''%s''            ' +
      '    AND IDENTIFYID2 = ''%s''            ' +
      '    AND COMPID = ''%s''                 ' +
      '    AND NVL( REFNO, ''999'' ) <> ''1''  ' +      
      '  ORDER BY ITEMID                       ',
      [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID] ); 
    if ( frmMultiSelect.ShowModal = mrOk ) then
    begin
      aQueryParam.Param1 := frmMultiSelect.SelectedValue;
      txtInvUseItem.Text := frmMultiSelect.SelectedText;
    end;
  finally
    frmMultiSelect.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
