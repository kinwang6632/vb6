unit frmInvA01U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, Grids, DBGrids, StdCtrls, ExtCtrls, ComCtrls, Buttons, Mask,
  jpeg, Menus, cxLookAndFeelPainters, cxButtons, DBClient, ADODB, Provider,
  cxControls, cxPC, cxRadioGroup, cxContainer, cxEdit, cxTextEdit,
  cxMaskEdit, cxButtonEdit, cxProgressBar, cxGroupBox, dxSkinsCore,
  dxSkinsDefaultPainters, dxSkinscxPCPainter;

const
  TRANS_FILE_SEP = chr(1);
  INVOICE_NUM_LENGTH = 8;  //發票位數

type
  TfrmInvA01 = class(TForm)
    OpenDialog1: TOpenDialog;
    dsrPrefix: TDataSource;
    ppmPrefix: TPopupMenu;
    N1: TMenuItem;
    Panel2: TPanel;
    Panel1: TPanel;
    rdoAction1: TcxRadioButton;
    rdoAction2: TcxRadioButton;
    Bevel3: TBevel;
    ActionPageControl: TcxPageControl;
    TransInv1: TcxTabSheet;
    BkImage: TImage;
    Label5: TLabel;
    Label6: TLabel;
    Label8: TLabel;
    radPreCreated: TRadioButton;
    radCreated: TRadioButton;
    btnExit1: TcxButton;
    BitRun: TcxButton;
    TransInv2: TcxTabSheet;
    Label11: TLabel;
    lblTransSuccess: TLabel;
    Label10: TLabel;
    lblMsg: TLabel;
    BitBtn1: TcxButton;
    btnDuplicateReport: TcxButton;
    CreateInv1: TcxTabSheet;
    Label2: TLabel;
    lblInvDate: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label4: TLabel;
    tmeStartDate: TMaskEdit;
    tmeEndInvDate: TMaskEdit;
    tmeInvDate: TMaskEdit;
    DBGrid1: TDBGrid;
    RadioButton1: TRadioButton;
    RadioButton2: TRadioButton;
    chbInvAndCharge: TCheckBox;
    DBGrid2: TDBGrid;
    CreateInv2: TcxTabSheet;
    Label7: TLabel;
    lblUser2: TLabel;
    Label9: TLabel;
    lblQuery: TLabel;
    lblInvCount: TLabel;
    lblSaleAmount: TLabel;
    lblTaxAmount: TLabel;
    lblInvAmount: TLabel;
    lblHint: TLabel;
    lblValidCounts: TLabel;
    lblInvCount2: TLabel;
    lblSaleAmount2: TLabel;
    lblTaxAmount2: TLabel;
    lblInvAmount2: TLabel;
    Bevel1: TBevel;
    Label18: TLabel;
    Label19: TLabel;
    Bevel2: TBevel;
    btnListCreateInv: TcxButton;
    btnRptCreateInv: TcxButton;
    btnListNoInv: TcxButton;
    CreateInv3: TcxTabSheet;
    lblSaleAmount3: TLabel;
    lblTaxAmount3: TLabel;
    lblInvAmount3: TLabel;
    lblAssignInvCount: TLabel;
    lblQuery2: TLabel;
    lblCount: TLabel;
    Label12: TLabel;
    Label3: TLabel;
    lblStartInvID: TLabel;
    lblStopInvID: TLabel;
    ProgressBar1: TProgressBar;
    edtPreCreated: TcxButtonEdit;
    edtCreated: TcxButtonEdit;
    Label1: TLabel;
    Label15: TLabel;
    lblTransError: TLabel;
    ProgressBar2: TcxProgressBar;
    cxGroupBox1: TcxGroupBox;
    rdoOrder0: TcxRadioButton;
    rdoOrder1: TcxRadioButton;
    rdoOrder2: TcxRadioButton;
    btnAssign: TcxButton;
    btnRequery: TcxButton;
    btnQuery: TcxButton;
    btnReset: TcxButton;
    btnExit: TcxButton;
    cxButton1: TcxButton;
    rdoOrder3: TcxRadioButton;
    procedure FormShow(Sender: TObject);
    procedure chbInvAndChargeClick(Sender: TObject);
    procedure RadioButton1Click(Sender: TObject);
    procedure RadioButton2Click(Sender: TObject);
    procedure tmeStartDateExit(Sender: TObject);
    procedure tmeInvDateExit(Sender: TObject);
    procedure DBGrid1DblClick(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure btnExit1Click(Sender: TObject);
    procedure BitRunClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnExitClick(Sender: TObject);
    procedure BtnQueryClick(Sender: TObject);
    procedure BtnResetClick(Sender: TObject);
    procedure btnAssignClick(Sender: TObject);
    procedure btnListCreateInvClick(Sender: TObject);
    procedure btnListNoInvClick(Sender: TObject);
    procedure btnRptCreateInvClick(Sender: TObject);
    {}
    procedure Inv016ReconcileError(
      DataSet: TCustomClientDataSet; E: EReconcileError;
      UpdateKind: TUpdateKind; var Action: TReconcileAction);
    procedure Inv017ReconcileError(
      DataSet: TCustomClientDataSet; E: EReconcileError;
      UpdateKind: TUpdateKind; var Action: TReconcileAction);
    procedure btnDuplicateReportClick(Sender: TObject);
    procedure rdoAction1Click(Sender: TObject);
    procedure rdoAction2Click(Sender: TObject);
    procedure TransInv1Show(Sender: TObject);
    procedure edtPreCreatedPropertiesButtonClick(Sender: TObject;
      AButtonIndex: Integer);
    procedure edtCreatedPropertiesButtonClick(Sender: TObject;
      AButtonIndex: Integer);
    procedure btnRequeryClick(Sender: TObject);
  private
    { Private declarations }
    nG_TotalUnit: Integer;
    nG_TotalUnit2: Integer;
    fG_SaleAmount, fG_TaxAmount, fG_InvAmount: Double;
    fG_SaleAmount2, fG_TaxAmount2, fG_InvAmount2: Double;
    G_AllServiceTypeStrList : TStringList;
    sG_NoDataSQL: String;
    sG_HaveDataSQL: String;
    sG_Condition1, sG_Condition2, sG_Condition3: String;
    FApplyUpdateError: Boolean;
    FApplyUpdateErrMsg: String;
    FPreLogDate: String;
    procedure listInvPrefixData;
    function iSDataOk : Boolean;
    function checkIsValidPrefix : Boolean;
    function getPreFixString : String;
    function getNewFilePathName(sI_OldFilePathName : String) : String;
    function GetSortOrder: Integer;
    procedure setButtonCompetence;
    procedure getAllServiceType;
    procedure ShowDuplicateReport;
  public
    { Public declarations }
    sG_InvDate,sG_StartDate,sG_EndDate,sG_HowToCreate : String ;
    sG_LastInvDate,sG_InvYearMonth: String ;
    nG_OrderPrefixNo : Integer;
    procedure setProgress(nI_Pos:Integer);    
  end;

var
  frmInvA01: TfrmInvA01;

implementation

uses
  cbUtilis, dtmReportModule,
  dtmMainU, UdateTimeu, Ustru,
  frmMainU, dtmMainJU, frmInvA06_1U, frmInvA01_1U,
  rptInvC03_1U;

{$R *.dfm}

procedure TfrmInvA01.RadioButton1Click(Sender: TObject);
begin
  tmeInvDate.Enabled := true ;
  chbInvAndCharge.Visible := false ;
end;

procedure TfrmInvA01.RadioButton2Click(Sender: TObject);
begin
  chbInvAndCharge.Checked := false ;
  chbInvAndCharge.Visible := true ;
end;

procedure TfrmInvA01.chbInvAndChargeClick(Sender: TObject);
begin
  if chbInvAndCharge.Checked then
  begin
    tmeInvDate.Text := '    /  /  ' ;
    tmeInvDate.Enabled := false;
  end  
  else
    tmeInvDate.Enabled := true;
  listInvPrefixData;
end;

procedure TfrmInvA01.setProgress(nI_Pos:Integer);
begin
  ProgressBar1.StepIt;
  Application.ProcessMessages;
  lblCount.Caption := IntToStr(nI_Pos) ;
end;

procedure TfrmInvA01.FormShow(Sender: TObject);
var
  sL_SQL: String ;
begin

  ActionPageControl.ActivePageIndex := 0;
  rdoAction1Click( rdoAction1 );

  btnAssign.Enabled := False;
  btnListCreateInv.Visible := False;
  btnRptCreateInv.Visible := False;
  btnListNoInv.Visible := False;

  sL_SQL := 'SELECT INVPARAM FROM INV003 WHERE COMPID = ' +
    QuotedStr( dtmMain.getCompID );

  edtPreCreated.Text := dtmMainJ.GetXMLAttribute( sL_SQL, 'PInvoiceFilePath' );
  edtCreated.Text := dtmMainJ.GetXMLAttribute( sL_SQL, 'InvoiceFilePath' );

  self.Caption := frmMain.GetFormTitleString( 'A01', '發票結轉與開立' );
  setButtonCompetence;


  if not dtmMainJ.cdsPrefix.Active then
    dtmMainJ.cdsPrefix.CreateDataSet;
  dtmMainJ.cdsPrefix.EmptyDataSet;

  G_AllServiceTypeStrList := TStringList.Create;

  dtmMainJ.Inv099DataSet.Close;

  lblMsg.Caption := EmptyStr;
  lblTransSuccess.Caption := EmptyStr;
  lblTransError.Caption := EmptyStr;
  btnDuplicateReport.Visible := False;
  
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA01.tmeStartDateExit(Sender: TObject);
begin
  if ( Pos( ' ', tmeStartDate.Text ) > 0 ) and ( tmeStartDate.Text <> '    /  /  ') then
  begin
    WarningMsg( '請輸入正確之收費開始日期' );
    Exit;
  end;
  if ( tmeStartDate.Text = '    /  /  ') then
    tmeEndInvDate.Text := tmeStartDate.Text ;
  listInvPrefixData;
end;

procedure TfrmInvA01.tmeInvDateExit(Sender: TObject);
begin
 listInvPrefixData;
end;

procedure TfrmInvA01.listInvPrefixData;
var
    sL_TargetDate, sL_YaerMonth : String;
begin
    sL_TargetDate := '';
    sL_YaerMonth := '';
    if chbInvAndCharge.Checked then
    begin
      if tmeStartDate.Text='    /  /  ' then
      begin
        tmeStartDate.SetFocus;
        MessageDlg('請輸入收費日期!',mtError, [mbOK],0);
        Exit;
      end
      else
        sL_TargetDate := tmeStartDate.Text;
    end
    else
    begin
      sL_TargetDate := tmeInvDate.Text;
    end;

    if (sL_TargetDate<>'') and (sL_TargetDate<>'    /  /  ')then
    begin
      sL_YaerMonth := getInvoiceYearMonth(sL_TargetDate);
    end;

    dtmMainJ.listLastInvDate( dtmMain.getCompID, sL_YaerMonth);
end;

procedure TfrmInvA01.DBGrid1DblClick(Sender: TObject);
var
    sL_Prefix,sL_LastInvDate,sL_StartNum : String;
    sL_IdentifyId1,sL_IdentifyId2,sL_CompId,sL_YearMonth,sL_Memo : String;
begin
    if isDataOk then
    begin
      if checkIsValidPrefix then
      begin

        sL_Prefix := DBGrid1.DataSource.DataSet.FieldByName('Prefix').AsString;
        sL_LastInvDate := DBGrid1.DataSource.DataSet.FieldByName('LastInvDate').AsString;

        sL_IdentifyId1 := DBGrid1.DataSource.DataSet.FieldByName('IdentifyId1').AsString;
        sL_IdentifyId2 := DBGrid1.DataSource.DataSet.FieldByName('IdentifyId2').AsString;
        sL_CompId := DBGrid1.DataSource.DataSet.FieldByName('CompId').AsString;
        sL_YearMonth := DBGrid1.DataSource.DataSet.FieldByName('YearMonth').AsString;
        sL_StartNum := DBGrid1.DataSource.DataSet.FieldByName('StartNum').AsString;
        sL_Memo := DBGrid1.DataSource.DataSet.FieldByName('Memo').AsString;

        with dsrPrefix.DataSet do
        begin
          if not (dtmMainJ.cdsPrefix.Locate('IdentifyId1;IdentifyId2;CompId;YearMonth;Prefix;StartNum',VarArrayOf(
          [sL_IdentifyId1,StrToInt(sL_IdentifyId2),sL_CompId,sL_YearMonth,sL_Prefix,sL_StartNum]) , [])) then
          begin
            nG_OrderPrefixNo := nG_OrderPrefixNo + 1;

            Append;
            FieldByName('No').AsInteger := nG_OrderPrefixNo;
            FieldByName('Prefix').AsString := sL_Prefix;
            FieldByName('LastInvDate').AsString := sL_LastInvDate;

            FieldByName('IdentifyId1').AsString := sL_IdentifyId1;
            FieldByName('IdentifyId2').AsString := sL_IdentifyId2;
            FieldByName('CompId').AsString := sL_CompId;
            FieldByName('YearMonth').AsString := sL_YearMonth;
            FieldByName('StartNum').AsString := sL_StartNum;
            FieldByName('Memo').AsString := sL_Memo;
            Post;

          end
          else
            MessageDlg('此字軌已選過', mtError,[mbOK],0);

        end;
      end;
    end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA01.IsDataOk: Boolean;
begin
  sG_InvDate := tmeInvDate.Text ;
  sG_StartDate := tmeStartDate.Text ;
  sG_EndDate := tmeEndInvDate.Text ;

  Result := False;
  if (sG_InvDate <> '    /  /  ') and (not TUdateTime.IsDateStr(sG_InvDate,'/')) then
  begin
    WarningMsg('請輸入正確之發票日期');
    tmeInvDate.SetFocus;
    Exit;
  end;

  if not TUdateTime.IsDateStr(sG_StartDate,'/') then
  begin
    WarningMsg('請輸入正確之收費開始日期');
    tmeStartDate.SetFocus;
    exit;
  end;


  if not TUdateTime.IsDateStr(sG_EndDate,'/') then
  begin
    WarningMsg('請輸入正確之收費結束日期');
    tmeEndInvDate.SetFocus;
    Exit;
  end;

  if (sG_InvDate = '    /  /  ') AND (chbInvAndCharge.Checked = false) then
  begin
    WarningMsg('請輸入發票日期');
    tmeInvDate.SetFocus;
    Exit;
  end;

  if sG_StartDate = '    /  /  ' then
  begin
    WarningMsg('請輸入收費日期');
    tmeStartDate.SetFocus;
    Exit;
  end;

  //做日期範圍的判斷…
  if sG_EndDate = '    /  /  ' then
    sG_EndDate := sG_StartDate;//可省略！

  sG_InvYearMonth := getInvoiceYearMonth(sG_StartDate);
  if sG_InvYearMonth = getInvoiceYearMonth(sG_EndDate) then
  else
  begin
    WarningMsg('請輸入正確的日期範圍,例如：2001/01/01~2001/02/28！');
    tmeEndInvDate.Text := '    /  /  ';
    tmeEndInvDate.SetFocus;
    Exit;
  end;

  if ( sG_EndDate < sG_StartDate ) then
  begin
    WarningMsg('請輸入正確的日期範圍,收費日期(迄)必須大於等於收費日期(起)。' );
    tmeEndInvDate.SetFocus;
    tmeEndInvDate.SelectAll;
    Exit;
  end;

  Result := True;

end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA01.checkIsValidPrefix: Boolean;
var
    sL_PrefixMaxInvDate,sL_CheckDate,sL_ErrMsg : String;
    dL_PrefixMaxInvDate,dL_CheckDate : TDate;
begin
    if (RadioButton2.Checked) and (chbInvAndCharge.Checked) then
    begin
      //假若是後開且發票日期同收費日期,則檢查最後收費日期
      //是否大於該字軌最後發票開立日,若否則此字軌不能開立
      //sL_CheckDate := Trim(sG_EndDate);
        sL_CheckDate := Trim(sG_StartDate); 
      sL_ErrMsg := '開始收費日期必須大於或等於選取發票字軌的最後發票開立日。';
    end
    else
    begin
      //假若不是後開的發票日期同收費日期,則檢查發票日期
      //是否大於該字軌最後發票開立日,若否則此字軌不能開立
      sL_CheckDate := Trim(sG_InvDate);
      sL_ErrMsg := '發票日期必須大於等於字軌的最後發票開立日';
    end;

    sL_PrefixMaxInvDate := Trim(DBGrid1.DataSource.DataSet.FieldByName('LastInvDate').AsString);


    dL_CheckDate := StrToDate(sL_CheckDate);
    dL_PrefixMaxInvDate := StrToDate(sL_PrefixMaxInvDate);
    if dL_PrefixMaxInvDate > dL_CheckDate then
    begin
      WarningMsg( sL_ErrMsg );
      Result := False;
      Exit;
    end;

    Result := true;
    
end;

function TfrmInvA01.getPreFixString: String;
var
    sL_Prefix,sL_StartNum,sL_PreFixString : String;
    ii,nL_StartNumLen,nL_DiffLen : Integer;
begin
    sL_PreFixString := '';

    with dtmMainJ.cdsPrefix do
    begin
      First;
      while not EOF do
      begin
        //將Grid所選字軌串起字軌及開始號碼
        sL_Prefix := fieldByName('PREFIX').AsString;
        sL_StartNum := fieldByName('StartNum').AsString;

        //檢查發票號碼是否為8位數,若位數不夠補零
        nL_StartNumLen := Length(sL_StartNum);

        //發票一定是8位數
        if nL_StartNumLen <> INVOICE_NUM_LENGTH then
        begin
          //不足的前面補0
          nL_DiffLen := INVOICE_NUM_LENGTH - nL_StartNumLen;
          for ii:=1 to nL_DiffLen do
            sL_StartNum := '0' + sL_StartNum;

        end;

        //AA11111111,BB22222222,CC33333333
        if sL_PreFixString = '' then
          sL_PreFixString := sL_Prefix + sL_StartNum
        else
          sL_PreFixString := sL_PreFixString + ',' + sL_Prefix + sL_StartNum;


        Next;
      end;
    end;

    Result := sL_PreFixString;
end;

procedure TfrmInvA01.N1Click(Sender: TObject);
begin
    dtmMainJ.cdsPrefix.EmptyDataSet;
    nG_OrderPrefixNo := 0;

end;



function TfrmInvA01.getNewFilePathName(sI_OldFilePathName: String): String;
var
    sL_FilePath,sL_FileName,sL_FileExt,sL_CurDateTime,sL_RealFilePathName : String;
begin
    sL_FilePath := ExtractFileDir(sI_OldFilePathName);
    sL_FileExt := ExtractFileExt(sI_OldFilePathName);
    sL_FileName := TUstr.replaceStr(ExtractFileName(sI_OldFilePathName),sL_FileExt,'');

    sL_CurDateTime := TUstr.replaceStr(TUstr.replaceStr(TUstr.replaceStr(DateTimeToStr(now),'/',''),':',''),' ' ,'');

    sL_RealFilePathName := sL_FilePath + '\' + sL_FileName + '_' + sL_CurDateTime + '.txt';

    Result := sL_RealFilePathName;
end;

procedure TfrmInvA01.btnExit1Click(Sender: TObject);
begin
  btnExitClick(sender);
end;

procedure TfrmInvA01.BitRunClick(Sender: TObject);
//發票結轉開始！！
var
  MASTER_DATA_COLUMN_COUNT,
  DETAIL_DATA_COLUMN_COUNT: Integer ;

  sL_ChargeTitle, sL_XmlContent,sL_ResultMaster,sL_ResultDetail : String ;
  nL_MasterRecordCount,nL_DetailRecordCount : Integer;
  sL_UserID,sL_UserName,sL_IsPreCreated : String ;

  sL_SeqDate,sL_SQL,sL_InvoiceFilePath : String ;
  nL_CurLineNo:Integer ;
  L_Master,L_Detail,L_Txt : TStringList ;
  readStr : String ;
  //以下為 Master Data之欄位變數宣告
  sL_CompID,sL_CustID,sL_Tel,sL_BusinessID,sL_Title : String ;
  sL_MailAddr,sL_ZipCode,sL_InvAddr : String;
  sL_BeAssignedInvID,sL_IsValid : String;
  //以下為 Detail Data之欄位變數宣告
  sL_BillID,sL_BillIDItemNo,sL_MasterTaxType,sL_DetailChargeDate : String ;
  sL_MasterChargeDate,sL_ItemID,sL_Description,sL_Quantity : String ;
  sL_UnitPrice,sL_MasterTaxRate,sL_DetailTaxAmount : String ;
  sL_TotalAmount,sL_StartDate,sL_EndDate,sL_ChargeEn : String ;
  sL_BillID2,sL_BillIDItemNo2,sL_Seq2,sL_CompID2,sL_CustID2 : String ;

  dL_SaleAmount,dL_MasterTaxAmount,dL_TotalAmount,dL_MaxSeq : double;

  sL_ServiceType, sL_TmpChr,sL_CurMinSeq,sL_CurMaxSeq,sL_InvoiceNewFilePath: String;
  sL_StartChargeDate,sL_EndChargeDate, sL_AccountNo : String;
  nL_TotalCounts: Integer;
  fL_SaleAmount,fL_TaxAmount,fL_InvAmount : Double;
  aNoTransTotalAmount: Double;
  aDetailHasInsert, aIsChangeMaster: Boolean;

  aHasError, aIsWizText: Boolean;
  aMasterSeq, aNewSeq, aProcSeq: String;

  aSeq, aRenameFileName, aLinkToMis: String;
  aDupMasterCount, aDupDetailCount: Integer;

  aInvUseId, aInvUseDesc: String;

  { --------------------------------------------------- }

  function getSeqNo(aSeqDate: String): String;
  var
    aSql: String;
  begin
    aSQL := ' SELECT S_INV016_SEQ.NEXTVAL MAXSEQ FROM DUAL ';
    Result := ( aSeqDate + Lpad( dtmMainJ.openSQL( aSQL ), 7, '0' ) );
  end;

  { --------------------------------------------------- }

  procedure LogToInv034;
  begin
      dtmMain.adoComm.Close;
      dtmMain.adoComm.SQL.Clear;
      dtmMain.adoComm.SQL.Text :=
        '  INSERT INTO INV034                            ' +
        '  ( BILLID, BILLIDITEMNO,                       ' +
        '    TAXTYPE, CHARGEDATE,                        ' +
        '    ITEMID, DESCRIPTION,                        ' +
        '    QUANTITY, UNITPRICE,                        ' +
        '    TAXRATE, TAXAMOUNT,                         ' +
        '    TOTALAMOUNT, STARTDATE,                     ' +
        '    ENDDATE, CHARGEEN,                          ' +
        '    CUSTID, CUSTNAME,                           ' +
        '    LOGTIME, UPTEN, COMPID )                    ' +
        '  VALUES                                        ' +
        '  ( :BILLID, :BILLIDITEMNO,                     ' +
        '    :TAXTYPE,                                   ' +
        '    TO_DATE( :CHARGEDATE, ''YYYY/MM/DD'' ),     ' +
        '    :ITEMID, :DESCRIPTION,                      ' +
        '    :QUANTITY, :UNITPRICE,                      ' +
        '    :TAXRATE, :TAXAMOUNT,                       ' +
        '    :TOTALAMOUNT,                               ' +
        '    TO_DATE( :STARTDATE, ''YYYY/MM/DD'' ),      ' +
        '    TO_DATE( :ENDDATE, ''YYYY/MM/DD'' ),        ' +
        '    :CHARGEEN,                                  ' +
        '    :CUSTID, :CUSTNAME,                         ' +
        '    TO_DATE( :LOGTIME, ''YYYYMMDD HH24MISS'' ), ' +
        '    :UPTEN, :COMPID )               ';
      dtmMain.adoComm.Parameters.ParamByName( 'BILLID' ).Value :=
        dtmMainJ.cdsQryInsertInv017.FieldByName('BILLID').AsString;
      dtmMain.adoComm.Parameters.ParamByName( 'BILLIDITEMNO' ).Value :=
        dtmMainJ.cdsQryInsertInv017.FieldByName('BILLIDITEMNO').AsString;
      dtmMain.adoComm.Parameters.ParamByName( 'TAXTYPE' ).Value :=
        dtmMainJ.cdsQryInsertInv017.FieldByName('TAXTYPE').AsString;
      dtmMain.adoComm.Parameters.ParamByName( 'CHARGEDATE' ).Value :=
        dtmMainJ.cdsQryInsertInv017.FieldByName('CHARGEDATE').AsString;
      dtmMain.adoComm.Parameters.ParamByName( 'ITEMID' ).Value :=
        dtmMainJ.cdsQryInsertInv017.FieldByName('ITEMID').AsString;
      dtmMain.adoComm.Parameters.ParamByName( 'DESCRIPTION' ).Value :=
        dtmMainJ.cdsQryInsertInv017.FieldByName('DESCRIPTION').AsString;
      dtmMain.adoComm.Parameters.ParamByName( 'QUANTITY' ).Value :=
        dtmMainJ.cdsQryInsertInv017.FieldByName('QUANTITY').AsInteger;
      dtmMain.adoComm.Parameters.ParamByName( 'UNITPRICE' ).Value :=
        dtmMainJ.cdsQryInsertInv017.FieldByName('UNITPRICE').AsFloat;
      dtmMain.adoComm.Parameters.ParamByName( 'TAXRATE' ).Value :=
        dtmMainJ.cdsQryInsertInv017.FieldByName('TAXRATE').AsFloat;
      dtmMain.adoComm.Parameters.ParamByName( 'TAXAMOUNT' ).Value :=
        dtmMainJ.cdsQryInsertInv017.FieldByName('TAXAMOUNT').AsFloat;
      dtmMain.adoComm.Parameters.ParamByName( 'TOTALAMOUNT' ).Value :=
        dtmMainJ.cdsQryInsertInv017.FieldByName('TOTALAMOUNT').AsFloat;
      dtmMain.adoComm.Parameters.ParamByName( 'STARTDATE' ).Value :=
        dtmMainJ.cdsQryInsertInv017.FieldByName('STARTDATE').AsString;
      dtmMain.adoComm.Parameters.ParamByName( 'ENDDATE' ).Value :=
        dtmMainJ.cdsQryInsertInv017.FieldByName('ENDDATE').AsString;
      dtmMain.adoComm.Parameters.ParamByName( 'CHARGEEN' ).Value :=
        dtmMainJ.cdsQryInsertInv017.FieldByName('CHARGEEN').AsString;
      dtmMain.adoComm.Parameters.ParamByName( 'CUSTID' ).Value :=
        dtmMainJ.cdsQryInsertInv016.FieldByName('CUSTID').AsString;
      dtmMain.adoComm.Parameters.ParamByName( 'CUSTNAME' ).Value :=
        dtmMainJ.cdsQryInsertInv016.FieldByName('TITLE').AsString;
      dtmMain.adoComm.Parameters.ParamByName( 'LOGTIME' ).Value :=
        FPreLogDate;
      dtmMain.adoComm.Parameters.ParamByName( 'UPTEN' ).Value :=
        dtmMainJ.cdsQryInsertInv016.FieldByName('UPTEN').AsString;
      dtmMain.adoComm.Parameters.ParamByName( 'COMPID' ).Value :=
        dtmMainJ.cdsQryInsertInv016.FieldByName('COMPID').AsString;
      dtmMain.adoComm.ExecSQL;
    end;

    { --------------------------------------------------- }
    
begin

  lblTransSuccess.Caption := EmptyStr;
  lblTransError.Caption := EmptyStr;

  ProgressBar2.Position := 0;
  sL_UserID := dtmMain.getLoginUser ;
  sL_UserName :=  dtmMain.getLoginUserName;

  //取得Inv022所有服務別
  getAllServiceType;


  if radPreCreated.Checked = true then
    sL_IsPreCreated := '1'//預開
  else
   sL_IsPreCreated := '2';//後開

  dL_TotalAmount := 0;
  dL_SaleAmount := 0;
  dL_MasterTaxAmount := 0;
  sL_XmlContent := '';
  sL_ResultMaster := '';
  sL_ResultDetail := '';
  nL_MasterRecordCount := 0;
  nL_DetailRecordCount := 0;

  aNoTransTotalAmount := 0;

  MASTER_DATA_COLUMN_COUNT := 9;
  DETAIL_DATA_COLUMN_COUNT := 16;
  sL_BeAssignedInvID := 'N';
  sL_IsValid := 'Y';
  aDetailHasInsert := False;

  sL_SeqDate := getYearMonthDay6(date); //將當天日期轉為字串

  if sL_IsPreCreated = '1' then //預開
    sL_InvoiceFilePath := edtPreCreated.Text
  else
    sL_InvoiceFilePath := edtCreated.Text;

  if not FileExists( sL_InvoiceFilePath ) then
  begin
    WarningMsg( '所指定的資料收費檔不存在。' );
    if ( radPreCreated.Checked ) and ( edtPreCreated.CanFocus ) then
      edtPreCreated.SetFocus;
    if ( radCreated.Checked ) and ( edtCreated.CanFocus ) then
      edtCreated.SetFocus;
    Exit;
  end;

  ActionPageControl.ActivePageIndex := TransInv2.PageIndex;
  Application.ProcessMessages;

  dtmMainJ.ResetInv016;
  dtmMainJ.ResetInv017;
  dtmMainJ.ResetWitchOne;
  dtmMainJ.reactiveDataSet(dtmMainJ.AdoQryInsertInv016);
  dtmMainJ.reactiveDataSet(dtmMainJ.AdoQryInsertInv017);
  dtmMainJ.cdsQryInsertInv016.Open;
  dtmMainJ.cdsQryInsertInv017.Open;

  readStr:= '';
  L_Master := TStringList.Create;
  L_Detail := TStringList.Create;
  L_Txt := TStringList.Create;

  L_Txt.LoadFromFile(sL_InvoiceFilePath); //將路徑下的文字檔轉到TStringList內
  ProgressBar2.Properties.Max := L_Txt.Count; //將狀態列的最大值設為文字檔的行數

  lblMsg.Caption := '拆解文字檔中....';

  Application.ProcessMessages;

  sL_TmpChr := '$';
  aLinkToMis := 'Y';
  aIsWizText := False;
  
    for nL_CurLineNo:=0 to L_Txt.Count-1 do //取得文字檔的每一行  由零始
    begin

      readStr := L_Txt.Strings[nL_CurLineNo];

      if (readStr <> '') and (Length(readStr) >=2) then //判斷該行資料是否有效
      begin
        if Copy( readStr, 0, 2 ) = '!!' then //!!開頭是標明文字檔來源
        begin
          if ( Copy( readStr, 0, 5 ) = '!!WIZ' ) or ( Copy( readStr, 0, 5 ) = '!!DOS' ) then
            aLinkToMis := 'N';
          aIsWizText := ( Copy( readStr, 0, 5 ) = '!!WIZ' );
        end else
        if Copy( readStr, 0, 2 ) = '##' then //##開頭代表是主檔
        begin
//5#13732#14911576#1#1徐洪彩雲#1中壢市志明路2巷20號3樓#1320#1中壢市志明路2巷20號3樓
          readStr := _CutLeft(readStr,2);//減去左邊2個字元，並取得資料

          readStr := TUstr.replaceStr(readStr, TRANS_FILE_SEP, sL_TmpChr + ' ');
          readStr := TUstr.replaceStr(readStr, sL_TmpChr, TRANS_FILE_SEP);
          //將該行資料的每個值剖析出來
          L_Master := TUstr.ParseStrings(readStr,TRANS_FILE_SEP);

           //判斷其資料數目是否等於主檔應有的數目
          if ( L_Master.Count <> MASTER_DATA_COLUMN_COUNT ) then
          begin
            { 新增的發票用途代碼及說明,可能有, 也可能沒有, 所以解出來
              的參數為 9 或 11 }
            if ( L_Master.Count > 11 ) or ( L_Master.Count < 9 ) then
            begin
              WarningMsg( '第 ' + IntToStr( nL_CurLineNo +1 ) + ' 行主資料格式錯誤!!'
                +#10#13+readStr);//nL_CurLineNo加1 是因為它是由0算起
              dtmMainJ.ResetInv016;
              dtmMainJ.ResetInv017;
              Exit;
            end;  
          end;

          sL_CompID := dtmMain.getCompID;
          sL_CustID := Trim( L_Master.Strings[1] );
          sL_Tel := Trim( L_Master.Strings[2] );
          sL_BusinessID := Trim( L_Master.Strings[3] );
          sL_Title := Trim( L_Master.Strings[4] );
          sL_MailAddr := Trim( L_Master.Strings[5] );
          sL_ZipCode := Trim( L_Master.Strings[6] );
          sL_InvAddr := Trim( L_Master.Strings[7] );
          sL_ChargeTitle := Trim( L_Master.Strings[8] );
          {}
          aInvUseId := EmptyStr;
          aInvUseDesc := EmptyStr;
          try
            aInvUseId := Trim( L_Master.Strings[9] );
            aInvUseDesc := Trim( L_Master.Strings[10] );
          except
            {}
          end;
            
          { 如果上一筆主檔並沒有對應的明細, 則把上一筆已寫入的主檔刪除,
            並把主檔筆數扣掉 }
          if ( not aDetailHasInsert ) and ( nL_MasterRecordCount > 0 )  then
          begin
            dtmMainJ.DeleteInv016( aProcSeq );
            Dec( nL_MasterRecordCount );
          end
          else begin
            { 第一次進入, 尚未有 SEQ, 取 Sequence }
             aProcSeq := getSeqNo( sL_SeqDate );
          end;

          if nL_MasterRecordCount = 0 then sL_CurMinSeq := aProcSeq;

          dtmMainJ.InsertInv016( aProcSeq, sL_CompID, sL_CustID, sL_Tel,
            sL_BusinessID, sL_Title, sL_ZipCode, sL_InvAddr, sL_MailAddr,
            sL_BeAssignedInvID, sL_IsValid, sL_IsPreCreated,sL_UserName, sL_ChargeTitle, aInvUseId, aInvUseDesc );
          Inc( nL_MasterRecordCount );

          aDetailHasInsert := False;

          { 從文字檔角度來看, 已換不同主檔, 寫明細時須要判斷 }
          aIsChangeMaster := True;

          { 清掉上一筆主檔記錄所記錄新的 SEQ }
          dtmMainJ.ResetWitchOne;

        end
        else if Copy( readStr, 0, 2 ) = '--' then//該行為明細資料
        begin

          readStr := _CutLeft(readStr,2);//減去左邊2個字元
          readStr := TUstr.replaceStr(readStr, TRANS_FILE_SEP, sL_TmpChr + ' ');
          L_Detail := TUstr.ParseStrings(readStr,sL_TmpChr);//將該行資料的每個值剖析出來

          // wizard 在文字檔部份可能某些筆會多加, 信用卡號後4碼進來 (第17個欄位)
          if ( aIsWizText ) then
          begin
            DETAIL_DATA_COLUMN_COUNT := 17;
          end;

          if L_Detail.Count <> DETAIL_DATA_COLUMN_COUNT then
          begin
            WarningMsg( '第 ' + IntToStr( nL_CurLineNo + 1 ) + ' 行明細資料格式錯誤!!'
              +#10#13+readStr);//nL_CurLineNo加1 是因為它是由0算起
            dtmMainJ.ResetInv016;
            dtmMainJ.ResetInv017;
            ProgressBar2.Position := 0;
            Exit;
          end;

          sL_BillID := Trim(L_Detail.Strings[1]);
          sL_BillIDItemNo := Trim(L_Detail.Strings[2]);
          sL_MasterTaxType := Trim(L_Detail.Strings[3]);
          sL_DetailChargeDate :=  Trim(L_Detail.Strings[4]);
          sL_MasterChargeDate := sL_DetailChargeDate;
          sL_ItemID :=  Trim(L_Detail.Strings[5]);
          sL_Description :=  Trim(L_Detail.Strings[6]);
          sL_Quantity := Trim(L_Detail.Strings[7]);               //Number
          sL_UnitPrice :=  Trim(L_Detail.Strings[8]);             //Number
          sL_MasterTaxRate :=  Trim(L_Detail.Strings[9]);         //Number
          sL_DetailTaxAmount :=  Trim(L_Detail.Strings[10]);      //Number
          sL_TotalAmount :=  Trim(L_Detail.Strings[11]);          //Number
          sL_StartDate :=  Trim(L_Detail.Strings[12]);
          sL_EndDate :=  Trim(L_Detail.Strings[13]);
          sL_ChargeEn :=  Trim(L_Detail.Strings[14]);
          sL_ServiceType := Trim(L_Detail.Strings[15]);
          sL_AccountNo := EmptyStr;
          if ( DETAIL_DATA_COLUMN_COUNT >= 17 ) then
          begin
            sL_AccountNo := Trim(L_Detail.Strings[16]);
            {}
            if ( Length( sL_AccountNo ) > 4 ) then
              sL_AccountNo := Copy( sL_AccountNo, Length( sL_AccountNo ) - 3, 4 );
          end;    
          {}
          if G_AllServiceTypeStrList.IndexOf(sL_ServiceType) = -1 then
          begin
            WarningMsg( '第 ' + IntToStr( nL_CurLineNo + 1 ) + ' 行明細資料格式錯誤!!'
              +#10#13 + readStr );//nL_CurLineNo加1 是因為它是由0算起
            dtmMainJ.ResetInv016;
            dtmMainJ.ResetInv017;
            ProgressBar2.Position := 0;
            Exit;
          end;

           // 判斷明細的稅別是否跟已經寫進 INV017 DataSet 的稅別不同,
           // 若不同必須再寫一筆主檔, 因為稅別不同不可以開在同一張
           //  發票(前提是此稅別在同一張主檔找不到, 才再開一張新的)

           // 從文字檔角度來看已換不同主檔, 明細的第一筆直接就寫進
           // INV017 DataSet,當換明細時, 就要判斷稅別是否可以開在同一張發票
           if aIsChangeMaster then
           begin
             dtmMainJ.InsertInv017( aProcSeq,
             sL_BillID,sL_BillIDItemNo,sL_MasterTaxType,sL_DetailChargeDate,
             sL_ItemID,sL_Description,sL_Quantity,sL_UnitPrice,
             sL_MasterTaxRate,sL_DetailTaxAmount,sL_TotalAmount,sL_StartDate,
             sL_EndDate,sL_ChargeEn, sL_ServiceType, aLinkToMis, sL_AccountNo ) ;

             dtmMainJ.UpdateInv016( aProcSeq, StrToFloat(sL_TotalAmount ),
               StrToFloat( sL_DetailTaxAmount ), sL_MasterChargeDate, sL_MasterTaxType,
               sL_MasterTaxRate );

             { 將此筆主檔的 SEQ 記錄下來 }
             aMasterSeq := aProcSeq;

           end
           else begin
             // 已換不同明細, 須判斷稅別跟稅率
             // 判斷規則, 先找同一筆主檔的稅別稅率, 沒找到, 則找新寫入的資料 }
             aNewSeq := dtmMainJ.DetailBelongWitchOne( aMasterSeq , sL_MasterTaxType,
               StrToFloat( sL_MasterTaxRate ) );
             // 都沒有找到, 表示須開立新的發票
             if ( aNewSeq = EmptyStr ) then
             begin
               // 取新 Sequence
               aProcSeq := getSeqNo( sL_SeqDate );
               aNewSeq := aProcSeq;
               dtmMainJ.InsertInv016( aNewSeq, sL_CompID, sL_CustID, sL_Tel,
                 sL_BusinessID, sL_Title, sL_ZipCode, sL_InvAddr, sL_MailAddr,
                 sL_BeAssignedInvID, sL_IsValid, sL_IsPreCreated,sL_UserName, sL_ChargeTitle,
                 aInvUseId, aInvUseDesc );
               Inc( nL_MasterRecordCount );
               dtmMainJ.InsertBelongWitchOne( aNewSeq, sL_MasterTaxType,
                 StrToFloat( sL_MasterTaxRate ) );
             end;

             dtmMainJ.InsertInv017( aNewSeq,
             sL_BillID,sL_BillIDItemNo,sL_MasterTaxType,sL_DetailChargeDate,
             sL_ItemID,sL_Description,sL_Quantity,sL_UnitPrice,
             sL_MasterTaxRate,sL_DetailTaxAmount,sL_TotalAmount,sL_StartDate,
             sL_EndDate,sL_ChargeEn, sL_ServiceType, aLinkToMis, sL_AccountNo );

             dtmMainJ.UpdateInv016( aNewSeq, StrToFloat(sL_TotalAmount ),
             StrToFloat( sL_DetailTaxAmount ), sL_MasterChargeDate, sL_MasterTaxType,
             sL_MasterTaxRate );
           end;

           { 顯示結果用 }
           dL_TotalAmount := dL_TotalAmount + StrToFloat( sL_TotalAmount );
           Inc( nL_DetailRecordCount );

           aDetailHasInsert := True;
           aIsChangeMaster := False;

        end;

      end;

      ProgressBar2.Position := ( ProgressBar2.Position + 1 );
      Application.ProcessMessages;

    end; //文字檔處理結束

    lblMsg.Caption := '文字檔拆解完成, 檢查資料中.....';

    sL_CurMaxSeq := aProcSeq;

    ProgressBar2.Position := 0;
    ProgressBar2.Properties.Max := dtmMainJ.cdsQryInsertInv016.RecordCount;
    FPreLogDate := FormatDateTime( 'yyyymmdd hhnnss', Now );
    aDupMasterCount := 0;
    aDupDetailCount := 0;
    btnDuplicateReport.Visible := False;

    Sleep( 300 );

    //剔除重複的資料
    dtmMainJ.cdsQryInsertInv016.First;
    while not dtmMainJ.cdsQryInsertInv016.Eof do
    begin
      sL_Seq2 := dtmMainJ.cdsQryInsertInv016.FieldByName('SEQ').AsString;
      sL_CompID2 := dtmMainJ.cdsQryInsertInv016.FieldByName('COMPID').AsString;
      sL_CustId2 := dtmMainJ.cdsQryInsertInv016.FieldByName('CUSTID').AsString;
      dtmMainJ.cdsQryInsertInv017.Filtered := False;
      dtmMainJ.cdsQryInsertInv017.Filter := Format( 'SEQ=%s', [sL_Seq2] );
      dtmMainJ.cdsQryInsertInv017.Filtered := True;
      dtmMainJ.cdsQryInsertInv017.First;
      aHasError := False;
      while not dtmMainJ.cdsQryInsertInv017.Eof do
      begin
        { 先找 INV017 重複的從後端資料庫 }
        sL_BillID2 := dtmMainJ.cdsQryInsertInv017.FieldByName('BILLID').AsString;
        sL_BillIDItemNo2 := dtmMainJ.cdsQryInsertInv017.FieldByName('BILLIDITEMNO').AsString;
        sL_SQL := Format(
          'SELECT INV017.SEQ AS SEQ            ' +
          '  FROM INV016, INV017               ' +
          ' WHERE INV016.SEQ = INV017.SEQ      ' +
          '   AND INV016.COMPID=''%s''         ' +
          '   AND INV016.CUSTID=''%s''         ' +
          '   AND INV017.BILLID=''%s''         ' +
          '   AND INV017.BILLIDITEMNO=''%s''   ' +
          '   AND INV016.STOPFLAG = 0          ',
           [sL_CompID2, sL_CustId2, sL_BillID2, sL_BillIDItemNo2]);
        aHasError := ( dtmMainJ.openSQL(sL_SQL) <> EmptyStr );
        // 找到重覆, 再看看此筆收費資料是否已被開立發票且做廢
        // 若是被做廢, 允許再次匯入重新開立
        if aHasError then
        begin
          sL_SQL := Format(
            'SELECT A.INVID FROM INV024 A        ' +
            ' WHERE A.COMPID = ''%s''            ' +
            '   AND A.BILLID = ''%s''            ' +
            '   AND A.BILLIDITEMNO = ''%s''      ',
             [sL_CompID2, sL_BillID2, sL_BillIDItemNo2]);
          aHasError := ( dtmMainJ.openSQL(sL_SQL) = EmptyStr );
          if aHasError then Break;
        end;  
        { 如果 DB 沒重覆的, 再看 DataSet �堛爾禤� }
        dtmMainJ.cdsQryInsertInv017.Filtered := False;
        dtmMainJ.cdsQryInsertInv017.Filter := Format(
          'SEQ=''%s'' and BILLID=''%s'' and BILLIDITEMNO=''%s''', [
           sL_Seq2, sL_BillID2, sL_BillIDItemNo2] );
        dtmMainJ.cdsQryInsertInv017.Filtered := True;
        aHasError := ( dtmMainJ.cdsQryInsertInv017.RecordCount > 1 );
        if aHasError then Break;
        dtmMainJ.cdsQryInsertInv017.Next;
      end;
      { 只要細項有一筆跟 INV017 重複的就全刪, 不管是DB的或DataSet }
      if aHasError then
      begin
        dtmMainJ.cdsQryInsertInv017.Filtered := False;
        dtmMainJ.cdsQryInsertInv017.Filter := Format( 'SEQ=%s', [sL_Seq2] );
        dtmMainJ.cdsQryInsertInv017.Filtered := True;
        dtmMainJ.cdsQryInsertInv017.First;
        { 刪除明細 }
        while not dtmMainJ.cdsQryInsertInv017.Eof do
        begin
          Dec( nL_DetailRecordCount );
          LogToInv034;
          dtmMainJ.cdsQryInsertInv017.Delete;
          Inc( aDupDetailCount );
        end;
        { 刪主檔, 順便把未匯入的金額扣掉 }
        dL_TotalAmount := ( dL_TotalAmount -
          dtmMainJ.cdsQryInsertInv016.FieldByName('INVAMOUNT').AsFloat );
        { 主檔筆數減掉 }  
        Dec( nL_MasterRecordCount );
        { 未匯入的金額加總 }
        aNoTransTotalAmount := ( aNoTransTotalAmount +
          dtmMainJ.cdsQryInsertInv016.FieldByName('INVAMOUNT').AsFloat );
        { 刪除該筆主檔 }  
        dtmMainJ.cdsQryInsertInv016.Delete;
        Inc( aDupMasterCount );
      end else
        dtmMainJ.cdsQryInsertInv016.Next;
      ProgressBar2.Position := ( ProgressBar2.Position + 1 );
      Application.ProcessMessages;
    end;

    lblMsg.Caption := '文字檔檢查資料完成, 存檔中.....';
    Application.ProcessMessages;
    Application.ProcessMessages;

    L_Master.Free;
    L_Detail.Free;
    L_Txt.Free;

    FApplyUpdateError := False;
    FApplyUpdateErrMsg := EmptyStr;
    {}
    dtmMainJ.cdsQryInsertInv016.OnReconcileError := Inv016ReconcileError;
    dtmMainJ.cdsQryInsertInv017.OnReconcileError := Inv017ReconcileError;
    try
     dtmMain.InvConnection.BeginTrans;
     try
       dtmMainJ.cdsQryInsertInv016.ApplyUpdates( 0 );
       if FApplyUpdateError then
         raise Exception.CreateFmt( '匯入發票收費資料(單頭)失敗, 原因: %s', [FApplyUpdateErrMsg] );
       dtmMainJ.cdsQryInsertInv017.ApplyUpdates( 0 );
       if FApplyUpdateError then
         raise Exception.CreateFmt( '匯入發票收費資料(明細)失敗, 原因: %s', [FApplyUpdateErrMsg] );
       dtmMain.InvConnection.CommitTrans;
     except
       on E: Exception do
       begin
         dtmMain.InvConnection.RollbackTrans;
         ErrorMsg( E.Message );
         Exit;
       end;
     end;
   finally
     dtmMainJ.cdsQryInsertInv016.OnReconcileError := nil;
     dtmMainJ.cdsQryInsertInv017.OnReconcileError := nil;
   end;

   lblMsg.Caption := '存檔完成。';

   Application.ProcessMessages;

   lblTransSuccess.Caption :=
     '共結轉成功收費資料主檔 '+ IntToStr( nL_MasterRecordCount ) + ' 筆 '#10#13 +
     '收費資料明細檔 ' + IntToStr( nL_DetailRecordCount )+ ' 筆 '#10#13 +
     '結轉總金額 ' + FormatFloat( '#,##0', dL_TotalAmount ) + ' 元' ;


    btnDuplicateReport.Visible := ( ( aDupMasterCount or aDupDetailCount ) > 0 );

    if ( ( aDupMasterCount or aDupDetailCount ) > 0 ) then
    begin
     lblTransError.Caption :=
       '重覆未結轉的收費資料主檔 ' + IntToStr( aDupMasterCount ) + ' 筆 '#10#13 +
       '收費資料明細檔 ' + IntToStr( aDupDetailCount ) + ' 筆 '#10#13 +
       '未結轉總金額 ' + FormatFloat( '#,##0', aNoTransTotalAmount ) + ' 元' ;
     if ConfirmMsg( Format(
       '本次匯整文字檔, 共有%d筆發票明細資料未匯入,'#13#10 +
       '您要檢視未匯入的明細資料? ', [aDupDetailCount] ) ) then
     begin
       ProgressBar2.Position := 0;
       ShowDuplicateReport;
     end;
    end;

    ProgressBar2.Position := 0;


    Application.ProcessMessages;

    sL_InvoiceNewFilePath := getNewFilePathName(sL_InvoiceFilePath);

    //檔案更名的處理
    //Dennis
    if FileExists(ChangeFileExt(sL_InvoiceNewFilePath,'.bak')) then
    begin//多做一個刪除已存在之bak檔
      DeleteFile(ChangeFileExt(sL_InvoiceNewFilePath,'.bak'));
      RenameFile(sL_InvoiceFilePath,ChangeFileExt(sL_InvoiceNewFilePath,'.bak'));
    end
    else
      RenameFile(sL_InvoiceFilePath,ChangeFileExt(sL_InvoiceNewFilePath,'.bak'));

   aRenameFileName := ChangeFileExt( ExtractFileName( sL_InvoiceNewFilePath ),
    '.bak' );

    lblMsg.Caption := Format( '本次文字檔處理完成, 已重新變更檔案名稱為:'#13#10'%s。', [
       aRenameFileName] );

    Application.ProcessMessages;   

    //取得此批結轉資料最大及最小的收費時間
    dtmMainJ.getInv016MinAndMaxChargeDate( dtmMain.getCompID,
      sL_CurMinSeq,sL_CurMaxSeq,sL_StartChargeDate,sL_EndChargeDate);

    nL_TotalCounts := dtmMainJ.getInv016Data(QUERY_TYPE_ALL,'','',
      dtmMain.getCompID,sL_StartChargeDate,sL_EndChargeDate,sL_CurMinSeq,
      sL_CurMaxSeq,fL_SaleAmount,fL_TaxAmount,fL_InvAmount );

    if nL_TotalCounts <> 0 then
    begin
      frmInvA06_1 := TfrmInvA06_1.Create(Application);
      try
        frmInvA06_1.sG_UserID := dtmMain.getLoginUser;
        frmInvA06_1.sG_StartChargeDate := sL_StartChargeDate;
        frmInvA06_1.sG_EndChargeDate := sL_EndChargeDate;
        frmInvA06_1.sG_CompID := dtmMain.getCompID;
        frmInvA06_1.fG_SaleAmount := fL_SaleAmount;
        frmInvA06_1.fG_TaxAmount := fL_TaxAmount;
        frmInvA06_1.fG_InvAmount := fL_InvAmount;
        frmInvA06_1.fG_Counts := nL_TotalCounts;
        frmInvA06_1.sGMinSeq := sL_CurMinSeq;
        frmInvA06_1.sGMaxSeq := sL_CurMaxSeq;
        frmInvA06_1.ShowModal;
      finally
        frmInvA06_1.Free;
        frmInvA06_1 := nil;
      end;  
    end;

end;

procedure TfrmInvA01.setButtonCompetence;
var
    sL_QueryCompetence,sL_InsertCompetence,sL_DeleteCompetence,sL_UpdateCompetence,sL_ExecuteCompetence : String;
begin
     //設定權限
     dtmMain.getCompetence('A01',sL_QueryCompetence,sL_InsertCompetence,sL_DeleteCompetence,sL_UpdateCompetence,sL_ExecuteCompetence);

     if sL_ExecuteCompetence = 'Y' then
       BitRun.Enabled := true
     else
       BitRun.Enabled := false;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA01.getAllServiceType;
begin
  dtmMainJ.getAllInv022Data;
  G_AllServiceTypeStrList.Clear;
  dtmMainJ.adoInv022Code.First;
  while not dtmMainJ.adoInv022Code.Eof do
  begin
    G_AllServiceTypeStrList.Add(
      dtmMainJ.adoInv022Code.FieldByName( 'ITEMID').AsString );
    dtmMainJ.adoInv022Code.Next;
  end;
  dtmMainJ.adoInv022Code.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA01.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  G_AllServiceTypeStrList.Free;
  dtmMainJ.Inv099DataSet.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA01.btnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmInvA01.BtnQueryClick(Sender: TObject);
var
  sL_MaxInvDate, aHaveInvSQL, aNoInvSQL, aOrginal: String;
  nL_InvNo: integer ;
begin

  if not IsDataOk then Exit;

  if sG_StartDate = '    /  /  ' then
  begin
    WarningMsg( '請輸入收費日期' );
    Exit;
  end;

  if DBGrid2.DataSource.DataSet.RecordCount<=0 then
  begin
    WarningMsg( '請確定此次開立所用到的字軌為何!' );
    Exit;                                             
  end;

  Screen.Cursor := crSQLWait;
  try                                
    // sG_HowToCreate = 1 預開
    // sG_HowToCreate = 2 後開
    sG_HowToCreate := '2';
    if RadioButton1.Checked = True then
      sG_HowToCreate := '1';

    // 若選取發票日期同收費日期
    if chbInvAndCharge.Checked then
    begin
      sG_InvDate := sG_StartDate;
    end else
    begin
      sG_InvYearMonth := getInvoiceYearMonth(sG_InvDate) ;
      sG_LastInvDate := sG_InvDate;
    end;

    //計算出電子式之最後發票日期
    sL_MaxInvDate := dtmMainJ.getLastInvDate( dtmMain.getCompID, sG_InvYearMonth,
      IDENTIFYID1, IDENTIFYID2 );

    if sL_MaxInvDate = '' then
      sL_MaxInvDate := sG_InvDate;


    { 檢核輸入的發票日期, 收費日=發票日 }
    if ( sG_HowToCreate = '2' ) and ( chbInvAndCharge.Checked ) then
    begin
      if ( CompareText( sL_MaxInvDate, sG_InvDate  ) > 0 ) then
      begin
        WarningMsg( '開始收費日期必須大於或等於選取發票字軌的最後發票開立日。' );
        tmeStartDate.SetFocus;
        Exit;
      end;
    end else
    begin
      if ( CompareText( sL_MaxInvDate, sG_InvDate ) > 0 ) then
      begin
        WarningMsg( '輸入的發票日期必須大於或等於選取發票字軌的最後發票開立日。' );
        tmeInvDate.SetFocus;
        Exit;
      end;
    end;


    // 是否有設定自動換開發票明細筆數, 有的話下不同的SQL
    if ( dtmMain.GetAutoCreateNum <= 0 ) then
    begin
      aHaveInvSQL := Format( 
        ' SELECT COUNT(1) COUNT FROM INV016 A          ' +
        ' WHERE A.COMPID = ''%s''                      ' +
        '   AND A.CHARGEDATE BETWEEN                   ' +
        '       TO_DATE( ''%s'', ''YYYY/MM/DD'' ) AND  ' +
        '       TO_DATE( ''%s'', ''YYYY/MM/DD'' )      ' +
        '   AND A.BEASSIGNEDINVID = ''N''              ' +
        '   AND A.ISVALID = ''Y''                      ' +
        '   AND A.HOWTOCREATE = ''%s''                 ' +
        '   AND A.SHOULDBEASSIGNED = ''Y''             ' +
        '   AND A.INVAMOUNT > 0                        ' +
        '   AND A.TAXTYPE <> ''0''                     ' +
        '   AND A.STOPFLAG = 0                         ',
        [dtmMain.getCompID, sG_StartDate, sG_EndDate, sG_HowToCreate] );
     end else
     begin
       aHaveInvSQL := Format(
         ' SELECT SUM( COUNTS ) FROM                     ' +
         ' ( SELECT B.SEQ,                               ' +
         '     CEIL( COUNT( B.SEQ ) / %d ) AS COUNTS     ' +
         '   FROM INV016 A, INV017 B                     ' +
         ' WHERE A.SEQ = B.SEQ                           ' +
         '   AND A.COMPID = ''%s''                       ' +
         '   AND A.CHARGEDATE BETWEEN                    ' +
         '       TO_DATE( ''%s'', ''YYYY/MM/DD'' ) AND   ' +
         '       TO_DATE( ''%s'', ''YYYY/MM/DD'' )       ' +
         '   AND A.BEASSIGNEDINVID = ''N''               ' +
         '   AND A.ISVALID = ''Y''                       ' +
         '   AND A.HOWTOCREATE = ''%s''                  ' +
         '   AND A.SHOULDBEASSIGNED = ''Y''              ' +
         '   AND A.INVAMOUNT > 0                         ' +
         '   AND A.TAXTYPE <> ''0''                      ' +
         '   AND A.STOPFLAG = 0                          ' +
         '   AND B.SHOULDBEASSIGNED = ''Y''              ' +
         '  GROUP BY B.SEQ  )                            ',
         [dtmMain.GetAutoCreateNum, dtmMain.getCompID, sG_StartDate,
          sG_EndDate, sG_HowToCreate] );
     end;

      // 統計預計開立的 SQL
      aHaveInvSQL := aHaveInvSQL + Format( aOrginal,
        [dtmMain.getCompID, sG_StartDate, sG_EndDate, 'N', sG_HowToCreate] );

      // 算出可開立的發票張數
      nG_TotalUnit := StrToInt( Nvl( dtmMainJ.openSQL( aHaveInvSQL ), '0' ) ) ;

      //計算所選發票字軌可用發票張數
      nL_InvNo := dtmMainJ.getCanUseInvCounts( dtmMain.getCompID,
        sG_InvYearMonth, IDENTIFYID1, IDENTIFYID2 );

      //判斷可用張數及欲開之發票張數
      if nG_TotalUnit > nL_InvNo then
      begin
        WarningMsg( Format(
          '預估開立發票張數共 %d 張, '#13#10 +
          '所選擇發票本可用發票張數共 %d 張，'#13#10 +
          '發票張數不足無法開立！', [nG_TotalUnit, nL_InvNo] ) );
        Exit;
      end;

      // 跳到查詢結果這個Tab !!
      //PageControl1.ActivePage := Assign2 ;
      lblUser2.Caption := '操作者:' + dtmMain.getLoginUser;

      sG_Condition1 :=  Format(
        '公司簡稱: %s (%s)', [dtmMain.getCompID, dtmMain.getCompName] );

      if chbInvAndCharge.Checked then
      begin
        sG_Condition2 := '發票日期同收費日期';
        sG_Condition2 := '';
      end
      else
      begin
        sG_Condition2 := Format( '發票日期:%s', [sG_InvDate] );
        sG_Condition3 := Format( '收費日期:%s~%s', [sG_StartDate,sG_EndDate] );
      end;

      lblQuery.Caption :=
        sG_Condition1 + #13#10 + sG_Condition2 + #13#10 + sG_Condition3;

      fG_SaleAmount := 0;
      fG_TaxAmount := 0;
      fG_InvAmount := 0;

      // 預計開立的發票金額統計
      if ( nG_TotalUnit > 0 ) then
      begin
        // 是否有設定自動換開發票明細筆數, 有的話下不同的SQL
        if ( dtmMain.GetAutoCreateNum <= 0 ) then
        begin
          aHaveInvSQL := Format(
            ' SELECT SUM( NVL( A.SALEAMOUNT, 0 ) ) SALEAMOUNT, ' +
            '        SUM( NVL( A.TAXAMOUNT, 0 ) ) TAXAMOUNT,   ' +
            '        SUM( NVL( A.INVAMOUNT, 0 ) ) INVAMOUNT    ' +
            '  FROM INV016 A                                   ' + 
            ' WHERE A.COMPID = ''%s''                          ' +
            '   AND A.CHARGEDATE BETWEEN                       ' +
            '       TO_DATE( ''%s'', ''YYYY/MM/DD'' ) AND      ' +
            '       TO_DATE( ''%s'', ''YYYY/MM/DD'' )          ' +
            '   AND A.BEASSIGNEDINVID = ''N''                  ' +
            '   AND A.ISVALID = ''Y''                          ' +
            '   AND A.HOWTOCREATE = ''%s''                     ' +
            '   AND A.SHOULDBEASSIGNED = ''Y''                 ' +
            '   AND A.INVAMOUNT > 0                            ' +
            '   AND A.TAXTYPE <> ''0''                         ' +
            '   AND A.STOPFLAG = 0                             ',
            [dtmMain.getCompID, sG_StartDate, sG_EndDate, sG_HowToCreate] );
        end else
        begin
          aHaveInvSQL := Format(
            'SELECT SUM( B.QUANTITY  * B.UNITPRICE ) AS SALEAMOUNT, ' +
            '       SUM( B.TAXAMOUNT ) AS TAXAMOUNT,           ' +
            '       SUM( B.TOTALAMOUNT ) AS INVAMOUNT          ' +
            '  FROM INV016 A, INV017 B                         ' +
            ' WHERE A.SEQ = B.SEQ                              ' +
            '   AND A.COMPID = ''%s''                          ' +
            '   AND A.CHARGEDATE BETWEEN                       ' +
            '       TO_DATE( ''%s'', ''YYYY/MM/DD'' ) AND      ' +
            '       TO_DATE( ''%s'', ''YYYY/MM/DD'' )          ' +
            '   AND A.BEASSIGNEDINVID = ''N''                  ' +
            '   AND A.ISVALID = ''Y''                          ' +
            '   AND A.HOWTOCREATE = ''%s''                     ' +
            '   AND A.SHOULDBEASSIGNED = ''Y''                 ' +
            '   AND A.INVAMOUNT > 0                            ' +
            '   AND A.TAXTYPE <> ''0''                         ' +
            '   AND A.STOPFLAG = 0                             ' +
            '   AND B.SHOULDBEASSIGNED = ''Y''                 ',
            [dtmMain.getCompID, sG_StartDate, sG_EndDate, sG_HowToCreate] );
        end;

        dtmMainJ.openSQLGetAmount( aHaveInvSQL,
           fG_SaleAmount, fG_TaxAmount, fG_InvAmount );

        ProgressBar1.Max := nG_TotalUnit;
        
     end;

     lblInvCount.Caption :=
       '發票張數:' + Lpad( IntToStr( nG_TotalUnit ), 12, ' ' );
     lblSaleAmount.Caption :=
       '總銷售額:' + Lpad( FormatFloat( ',##0.##', fG_SaleAmount ), 12, ' ' );
     lblTaxAmount.Caption :=
       '總 稅 額:' + Lpad( FormatFloat( ',##0.##', fG_TaxAmount ), 12, ' ' );
     lblInvAmount.Caption :=
       '總 金 額:' + Lpad( FormatFloat( ',##0.##', fG_InvAmount ), 12, ' ' );
     lblValidCounts.Caption:=
       '可用發票:' + Lpad( IntToStr( nL_InvNo ), 12, ' ' );

     btnAssign.Enabled := ( nG_TotalUnit > 0 );
     btnListCreateInv.Visible := ( nG_TotalUnit > 0 );

     if ( btnAssign.Enabled )  then
       lblHint.Caption := '發票查詢完成'
     else
       lblHint.Caption := '無任何資料可供發票開立！' ;


     // 計算未開立的發票張數
     
     fG_SaleAmount2 := 0;
     fG_TaxAmount2 := 0;
     fG_InvAmount2 := 0;

     aNoInvSQL := Format(
       '   SELECT /*+ RULE */                              ' +
       '          COUNT(1) COUNT                           ' +
       '     FROM INV016 A, INV017 B                       ' +
       '    WHERE A.SEQ = B.SEQ                            ' +
       '      AND A.COMPID = ''%s''                        ' +
       '      AND A.CHARGEDATE BETWEEN                     ' +
       '          TO_DATE( ''%s'', ''YYYY/MM/DD'' ) AND    ' +
       '          TO_DATE( ''%s'', ''YYYY/MM/DD'' )        ' +
       '      AND A.BEASSIGNEDINVID = ''N''                ' +
       '      AND A.ISVALID = ''Y''                        ' +
       '      AND A.HOWTOCREATE = ''%s''                   ' +
       '      AND ( A.SHOULDBEASSIGNED = ''N'' OR          ' +
       '            A.INVAMOUNT <= 0  OR                   ' +
       '            A.TAXTYPE = ''0'' OR                   ' +
       '            B.SHOULDBEASSIGNED = ''N'' OR          ' +
       '            B.TOTALAMOUNT = 0  OR                  ' +
       '            B.TAXTYPE = ''0'' )                    ' +
       '      AND A.STOPFLAG = 0                           ',       
       [dtmMain.getCompID, sG_StartDate, sG_EndDate, sG_HowToCreate] );

     nG_TotalUnit2 := StrToIntDef( dtmMainJ.openSQL( aNoInvSQL ), 0 ) ;

      if ( nG_TotalUnit2 > 0 ) then
      begin
        aNoInvSQL := Format(
          ' SELECT /*+ RULE */                              ' +
          '        SUM( B.QUANTITY * B.UNITPRICE ) SALEAMOUNT,  ' +
          '        SUM( B.TAXAMOUNT ) TAXAMOUNT,            ' +
          '        SUM( B.TOTALAMOUNT ) INVAMOUNT           ' +
          '   FROM INV016 A, INV017 B                       ' +
          '  WHERE A.COMPID = ''%s''                        ' +
          '    AND A.SEQ = B.SEQ                            ' +
          '    AND A.CHARGEDATE BETWEEN                     ' +
          '        TO_DATE( ''%s'', ''YYYY/MM/DD'' ) AND    ' +
          '        TO_DATE( ''%s'', ''YYYY/MM/DD'' )        ' +
          '    AND A.BEASSIGNEDINVID = ''N''                ' +
          '    AND A.ISVALID = ''Y''                        ' +
          '    AND A.HOWTOCREATE = ''%s''                   ' +
          '    AND ( A.SHOULDBEASSIGNED = ''N'' OR          ' +
          '          A.INVAMOUNT <= 0  OR                   ' +
          '          A.TAXTYPE = ''0'' OR                   ' +
          '          B.SHOULDBEASSIGNED = ''N'' OR          ' +
          '          B.TOTALAMOUNT = 0  OR                  ' +
          '          B.TAXTYPE = ''0'' )                    ' +
          '    AND A.STOPFLAG = 0                           ',     
          [dtmMain.getCompID, sG_StartDate, sG_EndDate, sG_HowToCreate] );

        dtmMainJ.openSQLGetAmount( aNoInvSQL, fG_SaleAmount2, fG_TaxAmount2,
          fG_InvAmount2 );

      end;

     lblInvCount2.Caption :=
       '發票張數:' + Lpad( IntToStr( nG_TotalUnit2 ), 12, ' ' );
     lblSaleAmount2.Caption :=
       '總銷售額:' + Lpad( FormatFloat( ',##0.##', fG_SaleAmount2 ), 12, ' ' );
     lblTaxAmount2.Caption :=
       '總 稅 額:' + Lpad( FormatFloat( ',##0.##', fG_TaxAmount2 ), 12, ' ' );
     lblInvAmount2.Caption :=
       '總 金 額:' + Lpad( FormatFloat( ',##0.##', fG_InvAmount2 ), 12, ' ' );

     btnListNoInv.Visible := ( nG_TotalUnit2 > 0 );
     btnRptCreateInv.Visible := ( nG_TotalUnit2 > 0 );

     // 組出查詢畫面的 SQL

     // 預計開立發票的查詢畫面資料 SQL
     sG_HaveDataSQL := '';

     if ( btnAssign.Enabled ) then
     begin
       sG_HaveDataSQL := Format(
         ' SELECT A.SEQ,                                                  ' +
         '        A.CUSTID,                                               ' +
         '        A.TITLE,                                                ' +
         '        A.TEL,                                                  ' +
         '        A.BUSINESSID,                                           ' +
         '        A.ZIPCODE,                                              ' +
         '        A.INVADDR,                                              ' +
         '        A.MAILADDR,                                             ' +
         '        A.CHARGEDATE,                                           ' +
         '        DECODE( A.TAXTYPE, ''1'', ''應稅'',                     ' +
         '                           ''2'', ''零稅率'',                   ' +
         '                           ''3'', ''免稅'',                     ' +
         '                           A.TAXTYPE ) AS DESCRIPTION,          ' +
         '        A.TAXRATE,                                              ' +
         '        A.SALEAMOUNT,                                           ' +
         '        A.TAXAMOUNT,                                            ' +
         '        A.INVAMOUNT,                                            ' +
         '        DECODE( A.HOWTOCREATE, ''1'', ''預開'',                 ' +
         '                               ''2'', ''後開'',                 ' +
         '                               A.HOWTOCREATE ) HOWTOCREATE,     ' +
         '        A.CHARGETITLE,                                          ' +
         '        A.UPTTIME,                                              ' +
         '        A.UPTEN                                                 ' +
         '   FROM INV016 A                                                ' +
         '  WHERE A.COMPID = ''%s''                                       ' +
         '    AND A.CHARGEDATE BETWEEN                                    ' +
         '        TO_DATE( ''%s'', ''YYYY/MM/DD'' ) AND                   ' +
         '        TO_DATE( ''%s'', ''YYYY/MM/DD'' )                       ' +
         '    AND A.BEASSIGNEDINVID = ''N''                               ' +
         '    AND A.ISVALID = ''Y''                                       ' +
         '    AND A.HOWTOCREATE = ''%s''                                  ' +
         '    AND A.SHOULDBEASSIGNED = ''Y''                              ' +
         '    AND A.INVAMOUNT > 0                                         ' +
         '    AND A.TAXTYPE <> ''0''                                      ' +
         '    AND A.STOPFLAG = 0                                          ' +
         '  ORDER BY A.CHARGEDATE, A.ZIPCODE                              ',
         [dtmMain.getCompID, sG_StartDate, sG_EndDate, sG_HowToCreate] );
     end;

     // 未開立發票的查詢畫面資料 SQL
     sG_NoDataSQL := '';

     if ( btnListNoInv.Visible ) then
     begin
       sG_NoDataSQL := Format(
         ' SELECT /*+ RULE */                                             ' +
         '        A.SEQ,                                                  ' +
         '        A.CUSTID,                                               ' +
         '        A.TITLE,                                                ' +
         '        A.TEL,                                                  ' +
         '        A.BUSINESSID,                                           ' +
         '        A.ZIPCODE,                                              ' +
         '        A.INVADDR,                                              ' +
         '        A.MAILADDR,                                             ' +
         '        A.CHARGEDATE,                                           ' +
         '        DECODE( A.TAXTYPE, ''1'', ''應稅'',                     ' +
         '                           ''2'', ''零稅率'',                   ' +
         '                           ''3'', ''免稅'',                     ' +
         '                           A.TAXTYPE ) AS DESCRIPTION,          ' +
         '        A.TAXTYPE,                                              ' +
         '        A.TAXRATE,                                              ' +
         '        A.SALEAMOUNT,                                           ' +
         '        A.TAXAMOUNT,                                            ' +
         '        A.INVAMOUNT,                                            ' +
         '        DECODE( A.HOWTOCREATE, ''1'', ''預開'',                 ' +
         '                               ''2'', ''後開'',                 ' +
         '                               A.HOWTOCREATE ) HOWTOCREATE,     ' +
         '        A.CHARGETITLE,                                          ' +
         '        A.UPTTIME,                                              ' +
         '        A.UPTEN,                                                ' +
         '        A.SHOULDBEASSIGNED,                                     ' +
         '        B.SHOULDBEASSIGNED AS SHOULDBEASSIGNED2,                ' +
         '        B.TOTALAMOUNT,                                          ' +
         '        B.TAXTYPE AS TAXTYPE2                                   ' +
         '   FROM INV016 A, INV017 B                                      ' +
         '  WHERE A.SEQ = B.SEQ                                           ' +
         '    AND A.COMPID = ''%s''                                       ' +
         '    AND A.CHARGEDATE BETWEEN                                    ' +
         '        TO_DATE( ''%s'', ''YYYY/MM/DD'' ) AND                   ' +
         '        TO_DATE( ''%s'', ''YYYY/MM/DD'' )                       ' +
         '    AND A.BEASSIGNEDINVID = ''N''                               ' +
         '    AND A.ISVALID = ''Y''                                       ' +
         '    AND A.HOWTOCREATE = ''%s''                                  ' +
         '    AND ( A.SHOULDBEASSIGNED = ''N'' OR                         ' +
         '          A.INVAMOUNT <= 0 OR                                   ' +
         '          A.TAXTYPE = ''0'' OR                                  ' +
         //'          A.STOPFLAG <> 0 OR                                    ' +
         '          B.SHOULDBEASSIGNED = ''N'' OR                         ' +
         '          B.TOTALAMOUNT = 0 OR                                  ' +
         '          B.TAXTYPE = ''0''  )                                  ' +
         '    AND A.STOPFLAG = 0                                          ' +  
         '  ORDER BY A.CHARGEDATE, A.ZIPCODE                              ',
         [dtmMain.getCompID, sG_StartDate, sG_EndDate, sG_HowToCreate] );
     end;
  finally
    Screen.Cursor := crDefault;
  end;
  ActionPageControl.ActivePageIndex := CreateInv2.PageIndex;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA01.BtnResetClick(Sender: TObject);
begin
  tmeInvDate.Text := '';
  tmeStartDate.Text := '';
  tmeEndInvDate.Text := '';
  nG_OrderPrefixNo := 0;
  dtmMainJ.cdsPrefix.EmptyDataSet;
  dtmMainJ.Inv099DataSet.Close;
  if tmeStartDate.CanFocus then tmeStartDate.SetFocus;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA01.btnAssignClick(Sender: TObject);
var
  dL_ActionStartDateTime, dL_ActionEndDateTime, dL_ActionIntervalDateTime : TDate;
  sL_PreFixString,sL_MisDbOwner,sL_LogDateTime,sL_ConditionString : String;
  nL_Order,nL_DataCounts : Integer;
  bL_InvDateEqualChargeDate, aExecResult: Boolean;
  WPath, aReportSQL: String;
begin
  { 判斷inv003中的INVCREATING欄位是否為Y }
  if dtmMainJ.getInvCreating( dtmMain.getCompID )='Y' then
  begin
    WarningMsg( '其他發票系統開立中請稍後再試!' );
    Exit;
  end;
  sL_PreFixString := getPreFixString;
  if ConfirmMsg( '此次開立所用到的字軌為:' + sL_PreFixString + '?' ) then
  begin
    dL_ActionStartDateTime := now;

    sL_MisDbOwner := dtmMain.getMisDbOwner;
    { 0 --> 收費日期, 1 --> 郵遞區號, 2 --> 收費日期 + 郵遞區號 }
    nL_Order := GetSortOrder;
    { 傳入發票日期同收費日期之參數 }
    bL_InvDateEqualChargeDate := frmInvA01.chbInvAndCharge.Checked;

    //將頁籤移到開立結果
    ActionPageControl.ActivePageIndex := CreateInv3.PageIndex;

    Label3.Visible := true;
    Label3.Caption := ' 發票開立中.請稍後 ';
    Label3.Invalidate;


    lblSaleAmount3.Caption := '銷貨總額： ' ;
    lblTaxAmount3.Caption := '稅金總額： ' ;
    lblInvAmount3.Caption := '總計： '  ;
    lblAssignInvCount.Caption := '發票開立張數： 張';
    
    lblCount.Caption := '0';
    ProgressBar1.Position := 0; //設定狀態列的起始為零
    Application.ProcessMessages;

    SCreen.Cursor := crSQLWait;
    try
      aExecResult := dtmMainJ.assignInvoiceWithSF( nG_TotalUnit, fG_SaleAmount,
        fG_TaxAmount, fG_InvAmount,  dtmMain.getLoginUserName, dtmMain.getCompID,
        sG_HowToCreate, sG_InvDate, sG_InvYearMonth, sG_StartDate,
        sG_EndDate, sL_PreFixString, sL_MisDbOwner, nL_Order,
        bL_InvDateEqualChargeDate, sL_LogDateTime);
    finally
      SCreen.Cursor := crDefault;
    end;

    if not aExecResult then
    begin
      Label3.Caption := '發票開立失敗!';
      Exit;
    end;  

    Label3.Caption := '';
    dL_ActionEndDateTime := now;
    dL_ActionIntervalDateTime := dL_ActionEndDateTime - dL_ActionStartDateTime;
    dL_ActionIntervalDateTime := dL_ActionIntervalDateTime*24*60*60;
    lblQuery2.Caption := '開立結果：(共費時 ' + Format('%f',[dL_ActionIntervalDateTime]) + ' 秒)';

    //自動查詢出異常資料,取出符合條件的開立發票異常資料
    sL_ConditionString := 'Log 日期: ' + sL_LogDateTime;

    Application.ProcessMessages;

    nL_DataCounts := dtmMainJ.getUnusualInvID(
      CONDITION_TYPE_ASSIGN_LOGDATE,sL_LogDateTime,sL_LogDateTime, aReportSQL);

    if nL_DataCounts <> 0 then
    begin
      WPath := IncludeTrailingPathDelimiter( ExtractFilePath(
        Application.ExeName ) ) + REPORT_FOLDER + '\FrptInvC03_1.fr3';
      dtmReport.frxMainReport.LoadFromFile( WPath );
      dtmReport.ADOMaster.Close;
      dtmReport.ADOMaster.SQL.Text := aReportSQL;
      dtmReport.frxMainReport.Variables.Variables['aCondition'] :=
        QuotedStr( sL_ConditionString );
      dtmReport.frxMainReport.Variables.Variables['aOperator'] :=
        QuotedStr( dtmMain.getLoginUser );
      dtmReport.frxMainReport.ShowReport;
      dtmReport.ADOMaster.Close;
    end;
  end;
end;

procedure TfrmInvA01.btnListCreateInvClick(Sender: TObject);
var
  aForm: TfrmInvA01_1;
begin
  aForm := TfrmInvA01_1.Create( Application );
  try
    aForm.Caption := '預計開立發票資料查詢';
    aForm.lblMasterTitle.Caption := '預計開立發票';
    aForm.lblDetailTitle.Caption := '發票明細';
    aForm.SQL := sG_HaveDataSQL;
    aForm.DataType := 1;
    aForm.ShowModal;
    cxGroupBox1.SetFocus;
  finally
    aForm.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA01.btnListNoInvClick(Sender: TObject);
var
  aForm: TfrmInvA01_1;
begin
  aForm := TfrmInvA01_1.Create( Application );
  try
    aForm.Caption := '不開立發票資料查詢';
    aForm.lblMasterTitle.Caption := '不開立發票';
    aForm.lblDetailTitle.Caption := '發票明細';
    aForm.SQL := sG_NoDataSQL;
    aForm.DataType := 2;
    cxGroupBox1.SetFocus;
    aForm.ShowModal;
  finally
    aForm.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA01.btnRptCreateInvClick(Sender: TObject);
var
  WPath: String;
begin
{$IFDEF DEBUG}
  WPath := '..\Source\ReportTemplate\FrptInvA01_1.fr3';
{$ELSE}
  WPath := IncludeTrailingPathDelimiter( ExtractFilePath(
    Application.ExeName ) ) + REPORT_FOLDER + '\FrptInvA01_1.fr3';
{$ENDIF}
  dtmReport.frxMainReport.LoadFromFile( WPath );
  dtmReport.ADOMaster.Close;
  dtmReport.ADOMaster.SQL.Text := sG_NoDataSQL;
  dtmReport.frxMainReport.Variables.Variables['Condition1'] :=
    QuotedStr( sG_Condition1 );
  dtmReport.frxMainReport.Variables.Variables['Condition2'] :=
    QuotedStr( sG_Condition2 );
  dtmReport.frxMainReport.Variables.Variables['Condition3'] :=
    QuotedStr( sG_Condition3 );
  { 預計開立發票 }
  dtmReport.frxMainReport.Variables.Variables['HaveInvCount'] :=
    nG_TotalUnit;
  dtmReport.frxMainReport.Variables.Variables['HaveInvTax'] :=
    fG_TaxAmount;
  dtmReport.frxMainReport.Variables.Variables['HaveInvSale'] :=
    fG_SaleAmount;
  dtmReport.frxMainReport.Variables.Variables['HaveInvTotal'] :=
    fG_InvAmount;
  { 不開立發票 }
  dtmReport.frxMainReport.Variables.Variables['NoInvCount'] :=
    nG_TotalUnit2;
  dtmReport.frxMainReport.Variables.Variables['NoInvTax'] :=
    fG_TaxAmount2;
  dtmReport.frxMainReport.Variables.Variables['NoInvSale'] :=
    fG_SaleAmount2;
  dtmReport.frxMainReport.Variables.Variables['NoInvTotal'] :=
    fG_InvAmount2;
  dtmReport.frxMainReport.ShowReport;
  dtmReport.ADOMaster.Close;
  cxGroupBox1.SetFocus;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA01.Inv016ReconcileError(DataSet: TCustomClientDataSet;
  E: EReconcileError; UpdateKind: TUpdateKind;
  var Action: TReconcileAction);
begin
  FApplyUpdateError := True;
  FApplyUpdateErrMsg := E.Message;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA01.Inv017ReconcileError(DataSet: TCustomClientDataSet;
  E: EReconcileError; UpdateKind: TUpdateKind;
  var Action: TReconcileAction);
begin
  FApplyUpdateError := True;
  FApplyUpdateErrMsg := E.Message;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA01.ShowDuplicateReport;
var
  aSql, WPath: String;
begin
  aSql := Format(
    ' SELECT A.BILLID,                                            ' +
    '        A.BILLIDITEMNO,                                      ' +
    '        DECODE( A.TAXTYPE, ''1'', ''應稅'',                  ' +
    '                           ''2'', ''零稅'',                  ' +
    '                           ''3'', ''免稅'',                  ' + 
    '                           A.TAXTYPE ) AS TAXTYPE,           ' +
    '        A.CHARGEDATE,                                        ' +
    '        A.ITEMID,                                            ' +
    '        A.DESCRIPTION,                                       ' +
    '        A.QUANTITY,                                          ' +
    '        A.UNITPRICE,                                         ' +
    '        A.TAXRATE,                                           ' +
    '        A.TAXAMOUNT,                                         ' +
    '        A.TOTALAMOUNT,                                       ' +
    '        A.STARTDATE,                                         ' +
    '        A.ENDDATE,                                           ' +
    '        A.CHARGEEN,                                          ' +
    '        A.CUSTID,                                            ' +
    '        A.CUSTNAME,                                          ' +
    '        A.LOGTIME,                                           ' +
    '        A.UPTEN                                              ' +
    '   FROM INV034 A                                             ' +
    '  WHERE A.COMPID = ''%s''                                    ' +
    '    AND A.UPTEN = ''%s''                                     ' +
    '    AND TO_CHAR( A.LOGTIME, ''YYYYMMDD HH24MISS'' ) = ''%s'' ' +
    '    ORDER BY A.BILLID, A.BILLIDITEMNO ',
    [dtmMain.getCompID, dtmMain.getLoginUserName, FPreLogDate] );
   WPath := IncludeTrailingPathDelimiter( ExtractFilePath(
     Application.ExeName ) ) + REPORT_FOLDER + '\FrptInvA01_2.fr3';
   dtmReport.frxMainReport.LoadFromFile( WPath );
   dtmReport.ADOMaster.Close;
   dtmReport.ADOMaster.SQL.Text := aSql;
   dtmReport.frxMainReport.ShowReport;
   dtmReport.ADOMaster.Close;
   Exit;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA01.btnDuplicateReportClick(Sender: TObject);
begin
  ShowDuplicateReport;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA01.rdoAction1Click(Sender: TObject);
begin
  TransInv1.TabVisible := True;
  TransInv2.TabVisible := True;
  ActionPageControl.ActivePageIndex := TransInv1.PageIndex;
  CreateInv1.TabVisible := False;
  CreateInv2.TabVisible := False;
  CreateInv3.TabVisible := False;
  rdoAction1.Font.Color := clBlue;
  rdoAction1.Font.Style := ( rdoAction1.Font.Style + [fsBold] );
  rdoAction2.Font.Color := clWindowText;
  rdoAction2.Font.Style := ( rdoAction1.Font.Style - [fsBold] );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA01.rdoAction2Click(Sender: TObject);
begin
  CreateInv1.TabVisible := True;
  CreateInv2.TabVisible := True;
  CreateInv3.TabVisible := True;
  ActionPageControl.ActivePageIndex := CreateInv1.PageIndex;
  TransInv1.TabVisible := False;
  TransInv2.TabVisible := False;
  rdoAction2.Font.Color := clBlue;
  rdoAction2.Font.Style := ( rdoAction1.Font.Style + [fsBold] );
  rdoAction1.Font.Color := clWindowText;
  rdoAction1.Font.Style := ( rdoAction1.Font.Style - [fsBold] );
  if tmeStartDate.CanFocus then tmeStartDate.SetFocus;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA01.TransInv1Show(Sender: TObject);
begin
  BkImage.Parent := TcxTabSheet( Sender );
  BkImage.SendToBack;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA01.edtPreCreatedPropertiesButtonClick(Sender: TObject;
  AButtonIndex: Integer);
begin
  if OpenDialog1.Execute then edtPreCreated.Text := OpenDialog1.FileName;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA01.edtCreatedPropertiesButtonClick(Sender: TObject;
  AButtonIndex: Integer);
begin
  if OpenDialog1.Execute then edtCreated.Text := OpenDialog1.FileName;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvA01.GetSortOrder: Integer;
begin
  Result := 0;
  if ( rdoOrder1.Checked ) then
    Result := 1
  else if ( rdoOrder2.Checked ) then
    Result := 2
  else if ( rdoOrder3.Checked ) then
    Result := 3;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA01.btnRequeryClick(Sender: TObject);
begin
   ActionPageControl.ActivePageIndex := CreateInv1.PageIndex;
end;

{ ---------------------------------------------------------------------------- }

end.

