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
    grpInvKind: TGroupBox;
    chkNormalInv: TCheckBox;
    chkEInv: TCheckBox;
    grpImport: TGroupBox;
    edtImport: TEdit;
    btnImport: TButton;
    OpenDialog1: TOpenDialog;
    Label6: TLabel;
    txtGiftUnit: TcxButtonEdit;
    rdgUseGive: TRadioGroup;
    chkPrintCarrier: TCheckBox;
    grp1: TGroupBox;
    edtImportPrize: TEdit;
    btnImportPrize: TButton;
    grp2: TGroupBox;
    rbCopy: TRadioButton;
    rbOriginal: TRadioButton;
    rbAllOriginal: TRadioButton;
    procedure btnExitClick(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure medSInvIDExit(Sender: TObject);
    procedure fraSInvDateExit(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure medEInvIDExit(Sender: TObject);
    procedure chkInvUseClick(Sender: TObject);
    procedure txtInvUseItemPropertiesButtonClick(Sender: TObject;
      AButtonIndex: Integer);
    procedure btnImportClick(Sender: TObject);
    procedure chkNormalInvClick(Sender: TObject);
    procedure txtGiftUnitPropertiesButtonClick(Sender: TObject;
      AButtonIndex: Integer);
    procedure btnImportPrizeClick(Sender: TObject);
    procedure Barcode2D_QRCode1Encode(Sender: TObject; var Data: String;
      Barcode: String);
    procedure Barcode2D_QRCode1InvalidLength(Sender: TObject;
      Barcode: String; LinearFlag: Boolean);
  private
    { Private declarations }
    FImportInv : String;
    function isDataOK : Boolean;
    procedure DoSelectInvUseItem;
    procedure DoSelectInvGiftUnit;
    procedure ImportInv043;

  public
    { Public declarations }
  end;

var
  frmInvA05: TfrmInvA05;

implementation

uses cbUtilis, frmInvA05_1U, frmInvA05_2U, frmMainU,  dtmMainJU, dtmMainU,
  frmMultiSelectU,EncdDecd;

var
  aQueryParam: TQueryParam;
  aQueryParam2: TQueryParam;

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
    if dtmMain.sG_AESKey = '' then
    begin
      MessageDlg('AES?[?K???_?|???]?w',mtError,[mbOK],0);
      Result := False;
      Exit;
    end;
    if (sL_SInvID='') and (sL_SInvDate=EMPTY_DATE_STR) then
    begin
      MessageDlg('?????J?????@???d??????',mtError,[mbOk],0);
      medSInvID.SetFocus;
      Result := false;
      Exit;
    end
    else
    begin
      if sL_SInvID <> '' then
      begin
        if Length(sL_SInvID) <> 10 then
        begin
          MessageDlg('???J?o?????X???~',mtError,[mbOk],0);
          medSInvID.SetFocus;
          Result := false;
          Exit;
        end;

        if Length(sL_EInvID) <> 10 then
        begin
          MessageDlg('???J?o?????X???~',mtError,[mbOk],0);
          medEInvID.SetFocus;
          Result := false;
          Exit;
        end;

        if Copy(sL_SInvID,1,2) <> Copy(sL_EInvID,1,2) then
        begin
          MessageDlg('?????P?@?o???r?y',mtError,[mbOk],0);
          medEInvID.SetFocus;
          Result := false;
          Exit;
        end;

        if Copy(sL_SInvID,3,8) > Copy(sL_EInvID,3,8) then
        begin
          MessageDlg('?o?????X???J???????~!',mtError,[mbOk],0);
          medEInvID.SetFocus;
          Result := false;
          Exit;
        end;

        try
          StrToInt(Copy(sL_SInvID,3,8));
        except
          MessageDlg('?o?????X???????~!',mtError,[mbOk],0);
          medSInvID.SetFocus;
          Result := false;
          Exit;
        end;


        try
          StrToInt(Copy(sL_EInvID,3,8));
        except
          MessageDlg('?o?????X???????~!',mtError,[mbOk],0);
          medEInvID.SetFocus;
          Result := false;
          Exit;
        end;
      end;


      if sL_SInvDate<>EMPTY_DATE_STR then
      begin
        if sL_EInvDate=EMPTY_DATE_STR then
        begin
          MessageDlg('?????J?o??????????',mtError,[mbOk],0);
          fraEInvDate.mseYMD.SetFocus;
          Result := false;
          Exit;
        end;

        if sL_EInvDate < sL_SInvDate then
        begin
          MessageDlg('?????????????j???????}?l????',mtError,[mbOk],0);
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
    sL_InvKind,sL_Invs,sL_OriginalType : string;
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
      sL_OrderByType := '1'  //?o?????X
    else if rgpOrder.ItemIndex = 1 then
      sL_OrderByType := '2'  //?o??????
    else if rgpOrder.ItemIndex = 2 then
      sL_OrderByType := '3'; //?????s??

    if cobBusiness.ItemIndex = 0 then
      sL_BusinessType := '0'
    else if cobBusiness.ItemIndex = 1 then
      sL_BusinessType := '1'   //???~?H
    else if cobBusiness.ItemIndex = 2 then
      sL_BusinessType := '2';  //?D???~?H

    if rgpTransferType.ItemIndex = 0 then
      sL_TransferType := '1'  //?o???C?L
    else if rgpTransferType.ItemIndex = 1 then
      sL_TransferType := '2';  //???J???r??
    {#7128 Check Whether invoice is original By Kin 2015/12/02 }
    if rbOriginal.Checked then
      sL_OriginalType := '0'
    else if rbCopy.Checked then
      sL_OriginalType := '1'
    else
      sL_OriginalType := '2';

    sL_InvKind := EmptyStr;
    sL_Invs := EmptyStr;
    {#5668 ?W?[?o?????????? AND ???J?????? By Kin 2010/07/07}
    if dtmMain.GetStarEInvoice then
    begin
      if ( chkNormalInv.Checked ) and ( chkEInv.Checked ) then
      begin
        sL_InvKind := EmptyStr;
      end else
      begin
        if chkNormalInv.Checked then
          sL_InvKind := '0';
        if chkEInv.Checked then
          sL_InvKind := '1';
      end;
      {
      if chkEInv.Checked then
      begin
        sL_Invs := FImportInv;
      end;
      }
    end;

    //#6025 ???G?????????n?M?? By Kin 2011/05/26
//    FImportInv := EmptyStr;
    Screen.Cursor := crSQLWait;
    ImportInv043;
    sL_Invs := FImportInv;
    try
      // ?????C?L?o??????
      dtmMainJ.getPrintInvoiceTempData( sL_SInvID,
        sL_EInvID, sL_SInvDate, sL_EInvDate, sL_BusinessType,
        sL_InvKind,sL_Invs, EmptyStr,aQueryParam2.Param1,
        False,chkPrintCarrier.Checked,False,sL_OriginalType );
      // ?R???????????X??????????
      //5922 ?????i?????? By Kin 2011//02/10
      dtmMainJ.FilterPrintInvoiceTempData( rdgUseGive.ItemIndex ,
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
      WarningMsg( '???d???d?????L?o???????i???M?L?C' );
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
        frmInvA05_1.sG_originalType := sL_OriginalType;
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
      frmInvA05_2.sG_InvIds := sL_Invs;
      frmInvA05_2.sG_originalType := sL_OriginalType;
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
  Self.Caption := frmMain.GetFormTitleString( 'A05', '?q?l?p?????o???M?L' );
  if not dtmMain.GetStarEInvoice then
  begin
    btnOK.Top := grpInvKind.Top + 20;
    btnExit.Top := btnOK.Top;
    grpImport.Visible := False;
    grpInvKind.Visible := False;
    Self.Height := 347;

  end;
  if TrimString(TrimString( dtmMainJ.GetIfPrintCheckValue,''''),' ') = 'Y' Then
  begin
    rbCopy.Enabled := False;
    rbAllOriginal.Enabled := False;
  end;
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
    frmMultiSelect.DisplayFields := 'ITEMID,?N?X,DESCRIPTION,?o?????~';
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

procedure TfrmInvA05.btnImportClick(Sender: TObject);

begin
  OpenDialog1.Filter := 'Text files (*.txt)|*.txt';
  FImportInv := EmptyStr;
  if OpenDialog1.Execute then
    edtImport.Text := OpenDialog1.FileName;
  if ( not FileExists( edtImport.Text ) ) and ( edtImport.Text <> EmptyStr ) then
  begin
    WarningMsg( '?????w?????J?o???????????s?b?C' );
    Exit;
  end;



end;

procedure TfrmInvA05.chkNormalInvClick(Sender: TObject);
begin
  if (Sender as TCheckBox).Name = 'chkNormalInv' then
  begin
    if (Sender as TCheckBox).Checked then
    begin
      chkEInv.Checked := False;
    end;
  end else
  begin
    if (Sender as TCheckBox ).Checked then
    begin
      chkNormalInv.Checked := False;
      grpImport.Enabled := True;
    end else
    begin
      //grpImport.Enabled := False;
      grpImport.Enabled := True;
    end;
  end;
end;

procedure TfrmInvA05.txtGiftUnitPropertiesButtonClick(Sender: TObject;
  AButtonIndex: Integer);
begin
  if AButtonIndex = 1 then
    DoSelectInvGiftUnit
  else begin
    txtGiftUnit.Clear;
    aQueryParam2.Param1 := EmptyStr;
   end;
  if aQueryParam2.Param1 <> EmptyStr then
  begin
    chkInvUse.Checked := True;
  end;
end;

procedure TfrmInvA05.DoSelectInvGiftUnit;
begin
  frmMultiSelect := TfrmMultiSelect.Create( Application );
  try
    frmMultiSelect.Connection := dtmMain.InvConnection;
    frmMultiSelect.KeyFields := 'ITEMID';
    frmMultiSelect.KeyValues := aQueryParam2.Param1;
    frmMultiSelect.DisplayFields := 'ITEMID,?N?X,DESCRIPTION,????????';
    frmMultiSelect.ResultFields := 'DESCRIPTION';
    frmMultiSelect.SQL.Text := Format(
      ' SELECT ITEMID, DESCRIPTION FROM INV041 ' +
      '  WHERE IDENTIFYID1 = ''%s''            ' +
      '    AND IDENTIFYID2 = ''%s''            ' +
      '    AND COMPID = ''%s''                 ' +
      '  ORDER BY ITEMID                       ',
      [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID] );
    if ( frmMultiSelect.ShowModal = mrOk ) then
    begin
      aQueryParam2.Param1 := frmMultiSelect.SelectedValue;
      aQueryParam2.Param1 := TrimChar(aQueryParam2.Param1,['''']);
      txtGiftUnit.Text := frmMultiSelect.SelectedText;
    end;
  finally
    frmMultiSelect.Free;
  end;
end;

procedure TfrmInvA05.ImportInv043;
  var
    T_Import : TStringList;
    iCount : Integer;
    aSQL : string;
    aHaveData : Boolean;
begin
  FImportInv := EmptyStr;
  aHaveData := False;
  if  ( edtImport.Text = EmptyStr ) and ( edtImportPrize.Text = EmptyStr ) then
    Exit;
   {#6080 ?W?[INV043 ?H?K????1000?????? By Kin 2011/08/10}
  aSQL := Format( ' DELETE FROM INV043 WHERE COMPID = ''%S'' AND OWNER = ''%S'' ',
       [dtmMain.getCompID, dtmMain.getLoginUserName] );
  dtmMainJ.ExecuteSQL( aSQL );
  if ( edtImport.Text <> EmptyStr ) then
  begin
    if not FileExists( edtImport.Text ) then
    begin
      WarningMsg( '?????w?????J?o???????????s?b?C' );
      Exit;
    end;
    T_Import := TStringList.Create;
    T_Import.LoadFromFile( edtImport.Text );

    try
      //Self.Enabled := False;
      //Screen.Cursor := crSQLWait;


      for iCount := 0 to T_Import.Count -1  do
      begin
        if T_Import.Strings[iCount] <> EmptyStr then
        begin
          Application.ProcessMessages;
          {#6080 ?W?[INV043 ?H?K????1000?????? By Kin 2011/08/10}
          aSQL := Format( 'INSERT INTO INV043 ( INVID,COMPID,OWNER )      ' +
                          ' VALUES (''%s'',''%s'',''%s''           )      ' ,
                          [UpperCase(T_Import.Strings[iCount]),
                           dtmMain.getCompID,
                           dtmMain.getLoginUserName ]);

          dtmMainJ.ExecuteSQL( aSQL );
          aHaveData := True;

        end;
      end;
    finally
      T_Import.Free;
    end;
  end;
  {#6629 ?W?[?o???????M?U By Kin 2013/10/28}
  if ( edtImportPrize.Text <> EmptyStr ) then
  begin
    if not FileExists( edtImportPrize.Text ) then
    begin
      WarningMsg( '?????w?????J?????M?U?????????s?b?C' );
      Exit;
    end;
    T_Import := TStringList.Create;
    T_Import.LoadFromFile( edtImportPrize.Text );
    try

      for iCount := 0 to T_Import.Count -1  do
      begin
        Application.ProcessMessages;
        if (T_Import.Strings[iCount] <> EmptyStr) and
          ( Copy( T_Import.Strings[iCount],1,1) <> 'F' ) and
          (Length(T_Import.Strings[iCount])>= 446) then
        begin
          if ( UpperCase(Copy( T_Import.Strings[iCount],445,1))= 'Y' )
            Or ( UpperCase(Copy( T_Import.Strings[iCount],446,1))= 'Y' ) then
          begin

          end else
          begin
            aSQL := Format( 'INSERT INTO INV043 ( INVID,COMPID,OWNER )      ' +
                          ' VALUES (''%s'',''%s'',''%s''           )      ' ,
                          [UpperCase(copy( T_Import.Strings[iCount],16,10)),
                           dtmMain.getCompID,
                           dtmMain.getLoginUserName ]);

            dtmMainJ.ExecuteSQL( aSQL );
            aHaveData := True;
          end;




        end;
      end;
    finally
      T_Import.Free;
    end;
  end;

  if aHaveData then
  begin
     FImportInv := Format( ' SELECT INVID FROM INV043                   ' +
                            ' WHERE COMPID = ''%s'' AND OWNER = ''%s''   ' ,
                            [dtmMain.getCompID,dtmMain.getLoginUserName]);
  end;
end;

procedure TfrmInvA05.btnImportPrizeClick(Sender: TObject);
begin
  {#6629 ?W?[???J?o???????M?U By Kin 2013/10/28}
  OpenDialog1.Filter := 'Text files (*.*)|*.*';
  FImportInv := EmptyStr;
  if OpenDialog1.Execute then
    edtImportPrize.Text := OpenDialog1.FileName;
  if ( not FileExists( edtImportPrize.Text )) and ( edtImportPrize.Text <> EmptyStr ) then
  begin
    WarningMsg( '?????w?????J?????M?U?????????s?b?C' );
    Exit;
  end;
end;

procedure TfrmInvA05.Barcode2D_QRCode1Encode(Sender: TObject;
  var Data: String; Barcode: String);
begin
  Data := #$EF#$BB#$BF + AnsiToUTF8(Barcode);

end;

procedure TfrmInvA05.Barcode2D_QRCode1InvalidLength(Sender: TObject;
  Barcode: String; LinearFlag: Boolean);
begin
  //Barcode := #$EF#$BB#$BF + AnsiToUTF8(Barcode);
 // ShowMessage( Barcode );
end;

end.
