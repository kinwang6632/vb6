unit frmInvC01U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, Mask, ExtCtrls, fraYMDU, Grids, DBGrids, DB,
  DBClient, ComCtrls, frxClass, frxDBSet, frxDesgn, cxControls, cxContainer,
  cxEdit, cxTextEdit, cxMaskEdit, cxButtonEdit, cxGraphics, cxDropDownEdit,
  dxSkinsCore, dxSkinsDefaultPainters, dxSkinBlack, dxSkinBlue,
  dxSkinCaramel, dxSkinCoffee;


type
  TfrmInvC01 = class(TForm)
    Panel1: TPanel;
    Label1: TLabel;
    Panel2: TPanel;
    Panel3: TPanel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    fraEInvDate: TfraYMD;
    fraSInvDate: TfraYMD;
    rgpRptType: TRadioGroup;
    medSInvID: TMaskEdit;
    btnOK: TBitBtn;
    btnExit: TBitBtn;
    medEInvID: TMaskEdit;
    cobBusiness: TComboBox;
    fraEChargeDate: TfraYMD;
    fraSChargeDate: TfraYMD;
    cobIsPrinted: TComboBox;
    cobInvFormat: TComboBox;
    cobTaxType: TComboBox;
    cobToCreate: TComboBox;
    Label11: TLabel;
    edtExcelPath: TEdit;
    chkPrintMemo1: TCheckBox;
    chkPrintMemo2: TCheckBox;
    btnTransferPath: TButton;
    OpenDialog1: TOpenDialog;
    RadioGroup2: TRadioGroup;
    txtChargeItem: TcxButtonEdit;
    Label12: TLabel;
    cmbServiceType: TcxComboBox;
    rdgInvKind: TRadioGroup;
    rdgUseGive: TRadioGroup;
    rgpRptSource: TRadioGroup;
    rgFixFlag: TRadioGroup;
    rgUploadFlag: TRadioGroup;
    procedure FormCreate(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure cobInvFormatExit(Sender: TObject);
    procedure cobBusinessExit(Sender: TObject);
    procedure cobTaxTypeExit(Sender: TObject);
    procedure cobToCreateExit(Sender: TObject);
    procedure cobIsPrintedExit(Sender: TObject);
    procedure fraSInvDateExit(Sender: TObject);
    procedure fraSChargeDateExit(Sender: TObject);
    procedure medSInvIDExit(Sender: TObject);
    procedure rgpRptTypeClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnTransferPathClick(Sender: TObject);
    procedure medEInvIDExit(Sender: TObject);
    procedure txtChargeItemPropertiesButtonClick(Sender: TObject;
      AButtonIndex: Integer);
    procedure cmbServiceTypeExit(Sender: TObject);
  private
    { Private declarations }
    procedure initialComboBox;
    function isDataOK : Boolean;
    procedure DoSelectChargeItem;
    procedure GetReportType1Amt(aDataSet: TClientDataSet; var aSaleAmt, aTaxAmt,
      aInvAmt, aShouldAmt, aFreeAmt: Double);
  public
    { Public declarations }
  end;

var
  frmInvC01: TfrmInvC01;

implementation

uses dtmMainJU, dtmMainU, frmMainU, dtmReportModule, Ustru, FileCtrl,
     frmMultiSelectU;

{$R *.dfm}

var
  aQueryParam: TQueryParam;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC01.FormCreate(Sender: TObject);
begin
  initialComboBox;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC01.btnExitClick(Sender: TObject);
begin
  Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC01.initialComboBox;
var
  aParam: TComboBoxCreateParam;
begin
  aParam.Sql := ' SELECT ITEMID, DESCRIPTION FROM INV022 ORDER BY ITEMID ';
  aParam.KeyField := 'ITEMID';
  aParam.DescField := 'DESCRIPTION';
  aParam.AddAllText := True;
  CreateCxComboBoxItem( cmbServiceType, aParam );
  if ( cmbServiceType.Properties.Items.Count > 0 ) then
    cmbServiceType.ItemIndex := 0;
  cobIsPrinted.ItemIndex := 0;
  cobInvFormat.ItemIndex := 0;
  cobBusiness.ItemIndex := 0;
  cobTaxType.ItemIndex := 0;
  cobToCreate.ItemIndex := 0;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC01.btnOKClick(Sender: TObject);
var
    sL_SInvID,sL_EInvID,sL_SInvDate,sL_EInvDate,sL_SChargeDate,sL_EChargeDate : String;
    sL_ChargeItemCode,sL_ChargeItemName,sL_UserID,sL_CompID,sL_RptType : String;
    sL_IsPrinted,sL_InvFormat,sL_Business,sL_TaxType,sL_ToCreate,sL_ExcelSavePath : String;
    bL_PrintMemo1,bL_PrintMemo2 : Boolean;
    aServiceTypeParam: TComboBoxItemParam;
    aText1, aText2, aText3, aText4, aPath: String;
    aSaleAmt, aTaxAmt, aInvAmt, aShouldAmt, aFreeAmt: Double;
    aInvCount: Integer;
    aInvKind : Integer;
begin
    if not isDataOK then Exit;
    sL_UserID := dtmMain.getLoginUser;
    sL_ExcelSavePath := Trim(edtExcelPath.Text);
    sL_SInvID := UpperCase(Trim(medSInvID.Text));
    sL_EInvID := UpperCase(Trim(medEInvID.Text));
    sL_SInvDate := Trim(fraSInvDate.getYMD);
    sL_EInvDate := Trim(fraEInvDate.getYMD);
    sL_SChargeDate := Trim(fraSChargeDate.getYMD);
    sL_EChargeDate := Trim(fraEChargeDate.getYMD);
    if ( sL_SInvDate = EMPTY_DATE_STR ) then
    begin
      sL_SInvDate := '';
      sL_EInvDate := '';
    end;
    if ( sL_SChargeDate = EMPTY_DATE_STR ) then
    begin
      sL_SChargeDate := '';
      sL_EChargeDate := '';
    end;
    //???O????
    sL_ChargeItemCode := aQueryParam.Param1;
    sL_ChargeItemName := txtChargeItem.Text;
    //?A???O
    GetCxComboBoxItemValue( cmbServiceType, aServiceTypeParam );
    //?O?_?w?C?L
    if cobIsPrinted.ItemIndex = 0 then
      sL_IsPrinted := ''
    else if cobIsPrinted.ItemIndex = 1 then
      sL_IsPrinted := 'Y'
    else if cobIsPrinted.ItemIndex = 2 then
      sL_IsPrinted := 'N';


      //?o??????
      if cobInvFormat.ItemIndex = 0 then  //????
        sL_InvFormat := ''
      else if cobInvFormat.ItemIndex = 1 then  //?q?l??
        sL_InvFormat := '1'
      else if cobInvFormat.ItemIndex = 2 then  //???u?G?p
        sL_InvFormat := '2'
      else if cobInvFormat.ItemIndex = 3 then  //???u?T?p
        sL_InvFormat := '3';


      //???~?O
      if cobBusiness.ItemIndex = 0 then  //????
        sL_Business := ''
      else if cobBusiness.ItemIndex = 1 then  //???~?H
        sL_Business := '1'
      else if cobBusiness.ItemIndex = 2 then  //?D???~?H
        sL_Business := '2';

      //???|?O
      if cobTaxType.ItemIndex = 0 then  //????
        sL_TaxType := ''
      else if cobTaxType.ItemIndex = 1 then  //???|
        sL_TaxType := '1'
      else if cobTaxType.ItemIndex = 2 then  //?s?|?v
        sL_TaxType := '2'
      else if cobTaxType.ItemIndex = 3 then  //?K?|
        sL_TaxType := '3';


      //?}???O
      if cobToCreate.ItemIndex = 0 then  //????
        sL_ToCreate := EmptyStr
      else if cobToCreate.ItemIndex = 1 then  //mis ????-?w?}
        sL_ToCreate := '1'
      else if cobToCreate.ItemIndex = 2 then  //mis ????-???}
        sL_ToCreate := '2'
      else if cobToCreate.ItemIndex = 3 then  //?{???}??
        sL_ToCreate := '3'
      else if cobToCreate.ItemIndex = 4 then  //?@???}??
        sL_ToCreate := '4';


      //???q?O    
      if RadioGroup2.ItemIndex = 0 then
        sL_CompID := dtmMain.getCompID
      else
        sL_CompID := EmptyStr;


      //????????
      if rgpRptType.ItemIndex = 0 then        //???p??
      begin
        sL_RptType := '1';
      end
      else if rgpRptType.ItemIndex = 1 then   //??????
      begin
        sL_RptType := '2';
      end
      else if rgpRptType.ItemIndex = 2 then   //???pExcel
      begin
        sL_RptType := '3';
      end
      else if rgpRptType.ItemIndex = 3 then   //??????Excel
      begin
        sL_RptType := '4';
      end;

      //?O?_?[?L????1
      bL_PrintMemo1 := chkPrintMemo1.Checked;

      //?O?_?[?L????2
      bL_PrintMemo2 := chkPrintMemo2.Checked;
      aInvKind := 1;

      Screen.Cursor := crSQLWait;
      try
        { Excel ???p???b dtmMainJ.getChargeData ???????? }
        if dtmMainJ.getChargeData( sL_CompID, sL_UserID, sL_SInvDate, sL_EInvDate,
          sL_SChargeDate, sL_EChargeDate, sL_SInvID, sL_EInvID, sL_ChargeItemCode,
          sL_ChargeItemName, sL_IsPrinted, sL_InvFormat, sL_Business, sL_TaxType,
          sL_ToCreate, sL_RptType, sL_ExcelSavePath, bL_PrintMemo1, bL_PrintMemo2,
          aServiceTypeParam.KeyValue, aServiceTypeParam.DesValue,
          aText1, aText2, aText3, aText4, aInvCount,rdgInvKind.ItemIndex,
          rdgUseGive.ItemIndex,rgpRptSource.ItemIndex,
          rgFixFlag.ItemIndex,rgUploadFlag.ItemIndex) then
        begin
          if ( sL_RptType = '1' ) then
          begin
            {}
            GetReportType1Amt( dtmMainJ.cdsChargeMasterData, aSaleAmt, aTaxAmt,
              aInvAmt, aShouldAmt, aFreeAmt );
            {}
            dtmReport.frxC01_1.DataSet := dtmMainJ.cdsChargeMasterData;
            aPath := IncludeTrailingPathDelimiter( ExtractFilePath(
              Application.ExeName ) ) + IncludeTrailingPathDelimiter(
              REPORT_FOLDER ) + 'FrptInvC01_1.fr3';
            dtmReport.frxMainReport.LoadFromFile( aPath, True );
            {}
            dtmReport.frxMainReport.Variables.Variables['?d??????1'] :=
              QuotedStr( aText1 );
            dtmReport.frxMainReport.Variables.Variables['?d??????2'] :=
              QuotedStr( aText2 );
            dtmReport.frxMainReport.Variables.Variables['?d??????3'] :=
              QuotedStr( aText3 );
            dtmReport.frxMainReport.Variables.Variables['?d??????4'] :=
              QuotedStr( aText4 );
            dtmReport.frxMainReport.Variables.Variables['???O????????'] :=
              QuotedStr( IntToStr( dtmMainJ.cdsChargeMasterData.RecordCount ) );
            dtmReport.frxMainReport.Variables.Variables['?`?P???B'] :=
              QuotedStr( FloatToStr( aSaleAmt ) );
            dtmReport.frxMainReport.Variables.Variables['?`?|?B'] :=
              QuotedStr( FloatToStr( aTaxAmt ) );
            dtmReport.frxMainReport.Variables.Variables['?`?o?????B'] :=
              QuotedStr( FloatToStr( aInvAmt ) );
            dtmReport.frxMainReport.Variables.Variables['???|?P???B'] :=
              QuotedStr( FloatToStr( aShouldAmt ) );
            dtmReport.frxMainReport.Variables.Variables['?K?|?P???B'] :=
              QuotedStr( FloatToStr( aFreeAmt ) );
            dtmReport.frxMainReport.Variables.Variables['?L???H'] :=
              QuotedStr( dtmMain.getLoginUser );
            dtmReport.frxMainReport.ShowReport;  
          end else
          { ?????? }
          if ( sL_RptType = '2' ) then
          begin

            {}
            GetReportType1Amt( dtmMainJ.cdsChargeDetailData, aSaleAmt, aTaxAmt,
              aInvAmt, aShouldAmt, aFreeAmt );
            {}
            dtmReport.frxC01_2.DataSet := dtmMainJ.cdsChargeDetailData;
            aPath := IncludeTrailingPathDelimiter( ExtractFilePath(
              Application.ExeName ) ) + IncludeTrailingPathDelimiter(
              REPORT_FOLDER ) + 'FrptInvC01_2.fr3';
            dtmReport.frxMainReport.LoadFromFile( aPath, True );
            {}
            dtmReport.frxMainReport.Variables.Variables['?d??????1'] :=
              QuotedStr( aText1 );
            dtmReport.frxMainReport.Variables.Variables['?d??????2'] :=
              QuotedStr( aText2 );
            dtmReport.frxMainReport.Variables.Variables['?d??????3'] :=
              QuotedStr( aText3 );
            dtmReport.frxMainReport.Variables.Variables['?d??????4'] :=
              QuotedStr( aText4 );
            dtmReport.frxMainReport.Variables.Variables['?`?o???i??'] :=
              QuotedStr( IntToStr( aInvCount ) );
            dtmReport.frxMainReport.Variables.Variables['?`?P???B'] :=
              QuotedStr( FloatToStr( aSaleAmt ) );
            dtmReport.frxMainReport.Variables.Variables['?`?|?B'] :=
              QuotedStr( FloatToStr( aTaxAmt ) );
            dtmReport.frxMainReport.Variables.Variables['?`?o?????B'] :=
              QuotedStr( FloatToStr( aInvAmt ) );
            dtmReport.frxMainReport.Variables.Variables['???|?P???B'] :=
              QuotedStr( FloatToStr( aShouldAmt ) );
            dtmReport.frxMainReport.Variables.Variables['?K?|?P???B'] :=
              QuotedStr( FloatToStr( aFreeAmt ) );
            dtmReport.frxMainReport.Variables.Variables['?L???H'] :=
              QuotedStr( dtmMain.getLoginUser );
            dtmReport.frxMainReport.ShowReport;
          end;
        end;
      finally
        Screen.Cursor := crDefault;
      end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC01.cobInvFormatExit(Sender: TObject);
begin
  if cobInvFormat.Text = EmptyStr then
    cobInvFormat.ItemIndex := 0;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC01.cobBusinessExit(Sender: TObject);
begin
  if cobBusiness.Text = EmptyStr then
    cobBusiness.ItemIndex := 0;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC01.cobTaxTypeExit(Sender: TObject);
begin
  if cobTaxType.Text = EmptyStr then
    cobTaxType.ItemIndex := 0;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC01.cobToCreateExit(Sender: TObject);
begin
  if cobToCreate.Text = EmptyStr then
    cobToCreate.ItemIndex := 0;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC01.cmbServiceTypeExit(Sender: TObject);
begin
  if cmbServiceType.Text = EmptyStr then
    cmbServiceType.ItemIndex := 0;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC01.cobIsPrintedExit(Sender: TObject);
begin
  if cobIsPrinted.Text = EmptyStr then
   cobIsPrinted.ItemIndex := 0;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvC01.isDataOK: Boolean;
var
    sL_SInvID,sL_EInvID,sL_SInvDate,sL_EInvDate,sL_SChargeDate,sL_EChargeDate : String;
begin
    Result := true;
    sL_SInvID := Trim(medSInvID.Text);
    sL_EInvID := Trim(medEInvID.Text);
    sL_SInvDate := Trim(fraSInvDate.getYMD);
    sL_EInvDate := Trim(fraEInvDate.getYMD);
    sL_SChargeDate := Trim(fraSChargeDate.getYMD);
    sL_EChargeDate := Trim(fraEChargeDate.getYMD);

    //?????o??????
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


    //???????O????
    if sL_SChargeDate<>EMPTY_DATE_STR then
    begin
      if sL_EChargeDate=EMPTY_DATE_STR then
      begin
        MessageDlg('?????J?o??????????',mtError,[mbOk],0);
        fraEChargeDate.mseYMD.SetFocus;
        Result := false;
        Exit;
      end;

      if sL_EChargeDate < sL_SChargeDate then
      begin
        MessageDlg('?????????????j???????}?l????',mtError,[mbOk],0);
        fraEChargeDate.mseYMD.SetFocus;
        Result := false;
        Exit;
      end;
    end;



    //?????o?????X
    if sL_SInvID<>'' then
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

end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC01.fraSInvDateExit(Sender: TObject);
begin
    fraEInvDate.setYMD(fraSInvDate.getYMD);
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC01.fraSChargeDateExit(Sender: TObject);
begin
  fraEChargeDate.setYMD(fraSChargeDate.getYMD);
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC01.medSInvIDExit(Sender: TObject);
begin
  medSInvID.Text := UpperCase(Trim(medSInvID.Text));
  medEInvID.Text := UpperCase(Trim(medSInvID.Text));
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC01.rgpRptTypeClick(Sender: TObject);
begin
  if ( rgpRptType.ItemIndex = 0 ) or
     ( rgpRptType.ItemIndex = 1 ) or
     ( rgpRptType.ItemIndex = 2 ) then
  begin
    chkPrintMemo1.Enabled := False;
    chkPrintMemo2.Enabled := False;
  end else
  begin
    chkPrintMemo1.Enabled := True;
    chkPrintMemo2.Enabled := True;
  end
end;
{ ---------------------------------------------------------------------------- }

procedure TfrmInvC01.FormShow(Sender: TObject);
begin
  chkPrintMemo1.Enabled := False;
  chkPrintMemo2.Enabled := False;
  Self.Caption := frmMain.GetFormTitleString( 'C01', '???O???????p??' );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC01.btnTransferPathClick(Sender: TObject);
var
  aPath: String;
begin
  aPath := edtExcelPath.Text;
  if SelectDirectory( 'Excel?s?????|', EmptyStr, aPath ) then
    edtExcelPath.Text := aPath;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC01.medEInvIDExit(Sender: TObject);
begin
  medEInvID.Text := UpperCase(Trim(medEInvID.Text));
end;

{ ---------------------------------------------------------------------------- }


procedure TfrmInvC01.txtChargeItemPropertiesButtonClick(Sender: TObject;
  AButtonIndex: Integer);
begin
  if AButtonIndex = 1 then
    DoSelectChargeItem
  else begin
    txtChargeItem.Clear;
    aQueryParam.Param1 := EmptyStr;
   end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC01.DoSelectChargeItem;
begin
  frmMultiSelect := TfrmMultiSelect.Create( Application );
  try
    frmMultiSelect.Connection := dtmMain.InvConnection;
    frmMultiSelect.KeyFields := 'ITEMID';
    frmMultiSelect.KeyValues := aQueryParam.Param1;
    frmMultiSelect.DisplayFields := 'ITEMID,?N?X,DESCRIPTION,???O????';
    frmMultiSelect.ResultFields := 'DESCRIPTION';
    frmMultiSelect.SQL.Text := Format(
      ' SELECT ITEMID, DESCRIPTION FROM INV005 ' +
      '  WHERE IDENTIFYID1 = ''%s''            ' +
      '    AND IDENTIFYID2 = ''%s''            ' +
      '    AND COMPID = ''%s''                 ' +
      '  ORDER BY ITEMID                       ',
      [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID] ); 
    if ( frmMultiSelect.ShowModal = mrOk ) then
    begin
      aQueryParam.Param1 := frmMultiSelect.SelectedValue;
      txtChargeItem.Text := frmMultiSelect.SelectedText;
    end;
  finally
    frmMultiSelect.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvC01.GetReportType1Amt(aDataSet: TClientDataSet; var aSaleAmt, aTaxAmt, aInvAmt, aShouldAmt,
  aFreeAmt: Double);
begin
  { ?`?P???B }
  aSaleAmt := 0;
  { ?`?|?B }
  aTaxAmt := 0;
  { ?`?o?????B }
  aInvAmt := 0;
  { ???|?P???B }
  aShouldAmt := 0;
  { ?K?|?P???B }
  aFreeAmt := 0;
  {}
  aDataSet.First;
  while not aDataSet.Eof do
  begin
    aSaleAmt := ( aSaleAmt + aDataSet.FieldByName( 'SALEAMOUNT' ).AsFloat );
    aTaxAmt := ( aTaxAmt + aDataSet.FieldByName( 'TAXAMOUNT' ).AsFloat );
    aInvAmt := ( aInvAmt + aDataSet.FieldByName( 'TOTALAMOUNT' ).AsFloat );
    if ( aDataSet.FieldByName( 'TAXTYPE' ).AsString = '1' ) then
      aShouldAmt := ( aShouldAmt + aDataSet.FieldByName( 'SALEAMOUNT' ).AsFloat )
    else if ( aDataSet.FieldByName( 'TAXTYPE' ).AsString = '3' ) then
      aFreeAmt := ( aFreeAmt + aDataSet.FieldByName( 'SALEAMOUNT' ).AsFloat ); 
    aDataSet.Next;
  end;
  aDataSet.First;
end;

{ ---------------------------------------------------------------------------- }

end.
