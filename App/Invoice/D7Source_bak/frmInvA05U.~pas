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
  frmMultiSelectU;

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

    if (sL_SInvID='') and (sL_SInvDate=EMPTY_DATE_STR) then
    begin
      MessageDlg('請輸入其中一項查詢條件',mtError,[mbOk],0);
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
          MessageDlg('輸入發票號碼錯誤',mtError,[mbOk],0);
          medSInvID.SetFocus;
          Result := false;
          Exit;
        end;

        if Length(sL_EInvID) <> 10 then
        begin
          MessageDlg('輸入發票號碼錯誤',mtError,[mbOk],0);
          medEInvID.SetFocus;
          Result := false;
          Exit;
        end;

        if Copy(sL_SInvID,1,2) <> Copy(sL_EInvID,1,2) then
        begin
          MessageDlg('必須同一發票字軌',mtError,[mbOk],0);
          medEInvID.SetFocus;
          Result := false;
          Exit;
        end;

        if Copy(sL_SInvID,3,8) > Copy(sL_EInvID,3,8) then
        begin
          MessageDlg('發票號碼輸入順序錯誤!',mtError,[mbOk],0);
          medEInvID.SetFocus;
          Result := false;
          Exit;
        end;

        try
          StrToInt(Copy(sL_SInvID,3,8));
        except
          MessageDlg('發票號碼格式錯誤!',mtError,[mbOk],0);
          medSInvID.SetFocus;
          Result := false;
          Exit;
        end;


        try
          StrToInt(Copy(sL_EInvID,3,8));
        except
          MessageDlg('發票號碼格式錯誤!',mtError,[mbOk],0);
          medEInvID.SetFocus;
          Result := false;
          Exit;
        end;
      end;


      if sL_SInvDate<>EMPTY_DATE_STR then
      begin
        if sL_EInvDate=EMPTY_DATE_STR then
        begin
          MessageDlg('請輸入發票結束日期',mtError,[mbOk],0);
          fraEInvDate.mseYMD.SetFocus;
          Result := false;
          Exit;
        end;

        if sL_EInvDate < sL_SInvDate then
        begin
          MessageDlg('結束日期必須大於等於開始日期',mtError,[mbOk],0);
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
    sL_InvKind,sL_Invs : string;
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
      sL_OrderByType := '1'  //發票號碼
    else if rgpOrder.ItemIndex = 1 then
      sL_OrderByType := '2'  //發票日期
    else if rgpOrder.ItemIndex = 2 then
      sL_OrderByType := '3'; //客戶編號

    if cobBusiness.ItemIndex = 0 then
      sL_BusinessType := '0'
    else if cobBusiness.ItemIndex = 1 then
      sL_BusinessType := '1'   //營業人
    else if cobBusiness.ItemIndex = 2 then
      sL_BusinessType := '2';  //非營業人

    if rgpTransferType.ItemIndex = 0 then
      sL_TransferType := '1'  //發票列印
    else if rgpTransferType.ItemIndex = 1 then
      sL_TransferType := '2';  //轉入文字檔

    sL_InvKind := EmptyStr;
    sL_Invs := EmptyStr;
    {#5668 增加發票種類條件 AND 匯入的資料 By Kin 2010/07/07}
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

    //#6025 第二次資料不要清掉 By Kin 2011/05/26
//    FImportInv := EmptyStr;
    Screen.Cursor := crSQLWait;
    ImportInv043;
    sL_Invs := FImportInv;
    try
      // 產生列印發票資料
      dtmMainJ.getPrintInvoiceTempData( sL_SInvID,
        sL_EInvID, sL_SInvDate, sL_EInvDate, sL_BusinessType,
        sL_InvKind,sL_Invs, EmptyStr,aQueryParam2.Param1,False );
      // 刪掉其它不符合條件的資料
      //5922 捐贈可選全部 By Kin 2011//02/10
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
      WarningMsg( '此查詢範圍內無發票資料可供套印。' );
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
      frmInvA05_2.sG_InvIds := sL_Invs;
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
  Self.Caption := frmMain.GetFormTitleString( 'A05', '電子計算機發票套印' );
  if not dtmMain.GetStarEInvoice then
  begin
    btnOK.Top := grpInvKind.Top + 20;
    btnExit.Top := btnOK.Top;
    grpImport.Visible := False;
    grpInvKind.Visible := False;
    Self.Height := 347;

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
    frmMultiSelect.DisplayFields := 'ITEMID,代碼,DESCRIPTION,發票用途';
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
  FImportInv := EmptyStr;
  if OpenDialog1.Execute then
    edtImport.Text := OpenDialog1.FileName;
  if not FileExists( edtImport.Text ) then
  begin
    WarningMsg( '所指定的匯入發票資料檔不存在。' );
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
    frmMultiSelect.DisplayFields := 'ITEMID,代碼,DESCRIPTION,受贈單位';
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
begin
  FImportInv := EmptyStr;
  if edtImport.Text = EmptyStr then
    Exit;

  if not FileExists( edtImport.Text ) then
  begin
    WarningMsg( '所指定的匯入發票資料檔不存在。' );
    Exit;
  end;
  T_Import := TStringList.Create;
  T_Import.LoadFromFile( edtImport.Text );
  FImportInv := EmptyStr;
  try
    //Self.Enabled := False;
    //Screen.Cursor := crSQLWait;
    {#6080 增加INV043 以便取消1000筆限制 By Kin 2011/08/10}
    aSQL := Format( ' DELETE FROM INV043 WHERE COMPID = ''%S'' AND OWNER = ''%S'' ',
       [dtmMain.getCompID, dtmMain.getLoginUserName] );
    dtmMainJ.ExecuteSQL( aSQL );

    for iCount := 0 to T_Import.Count -1  do
    begin
      if T_Import.Strings[iCount] <> EmptyStr then
      begin
        Application.ProcessMessages;
        {#6080 增加INV043 以便取消1000筆限制 By Kin 2011/08/10}
        aSQL := Format( 'INSERT INTO INV043 ( INVID,COMPID,OWNER )      ' +
                        ' VALUES (''%s'',''%s'',''%s''           )      ' ,
                        [UpperCase(T_Import.Strings[iCount]),
                         dtmMain.getCompID,
                         dtmMain.getLoginUserName ]);
        dtmMainJ.ExecuteSQL( aSQL );
        FImportInv := Format( ' SELECT INVID FROM INV043                   ' +
                              ' WHERE COMPID = ''%s'' AND OWNER = ''%s''   ' ,
                              [dtmMain.getCompID,dtmMain.getLoginUserName]);
        {
        if FImportInv = EmptyStr then
        begin
          FImportInv := QuotedStr( UpperCase(T_Import.Strings[iCount]));
        end else
        begin
          FImportInv := FImportInv + ',' + QuotedStr( UpperCase( T_Import.Strings[iCount]));
        end;
        }
      end;
    end;
  finally
    T_Import.Free;
    //Self.Enabled := True;
    //Screen.Cursor := crDefault;
  end;

end;

end.
