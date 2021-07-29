unit frmInvA08U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Buttons, Mask, ComCtrls, DB, DBCtrls, Grids,
  DBGrids, ADODB, StrUtils, Printers , math, Menus,
  cxControls, cxContainer, cxEdit, cxLabel, cxTextEdit, cxCurrencyEdit,
  cxLookAndFeelPainters, cxButtons, cxProgressBar, dxSkinsCore,
  dxSkinsDefaultPainters;
  
type
  //宣告媒體申報會用到的所有金額記錄
  TarraySum = Class(Tobject)
    sL_A, sL_B, sL_C, sL_D      : Double;
    sL_E, sL_F, sL_G, sL_H      : Double;
    sL_I, sL_J, sL_K, sL_L      : Double;
    sL_M, sL_N, sL_O, sL_P      : Double;
    sL_Q, sL_R, sL_S, sL_T      : Double;
    sL_U, sL_V, sL_W, sL_X      : Double;
    sL_Y, sL_Z                  : Double;
    sL_AA, sL_BB, sL_CC, sL_DD  : Double;
    sL_EE, sL_FF, sL_GG, sL_HH  : Double;
    sL_II, sL_JJ, sL_KK, sL_LL  : Double;
    sL_MM, sL_NN, sL_OO, sL_PP  : Double;
    sL_QQ, sL_RR, sL_SS         : Double;
    sL_TT, sL_UU, sL_VV         : String;
    sL_WW, sL_XX, sL_YY, sL_ZZ  : Double;
    sL_CC2, sL_DD2, sL_EE2, sL_FF2: Double;
    sL_KK2, sL_LL2, sL_MM2      : Double; { 電子、二聯、三聯, 零稅率張數統計 }
    sL_GG2                      : Double; 
  end;

  TfrmInvA08 = class(TForm)
    Panel1: TPanel;
    Label5: TLabel;
    Panel2: TPanel;
    PCT1: TPageControl;
    TST_Edit: TTabSheet;
    TST_View: TTabSheet;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label6: TLabel;
    RadioGroup1: TRadioGroup;
    Label7: TLabel;
    CheckBox1: TCheckBox;
    Label8: TLabel;
    CBx_Month: TComboBox;
    Edt_MediaFilePath: TEdit;
    AQry_Inv: TADOQuery;
    MKE_Year: TEdit;
    ListBox1: TListBox;
    Panel3: TPanel;
    Lbl_show: TLabel;
    ProgressBar2: TcxProgressBar;
    Bit_Previous: TcxButton;
    Btn_Print: TcxButton;
    btnExit: TcxButton;
    Panel4: TPanel;
    BitBtn2: TcxButton;
    btnRedo: TcxButton;
    btnOK: TcxButton;
    Bevel1: TBevel;
    MKE_Start: TMaskEdit;
    MKE_End: TMaskEdit;
    PrinterSetupDialog1: TPrinterSetupDialog;
    AQry_Inv099: TADOQuery;
    lblCompany: TcxLabel;
    lblUser: TcxLabel;
    Bevel2: TBevel;
    Label9: TLabel;
    RadioGroup2: TRadioGroup;
    chkStampTax: TCheckBox;
    txtStampTax: TcxCurrencyEdit;
    chkPaper: TCheckBox;
    chkApplyMainInv: TCheckBox;
    chkDebugFile: TCheckBox;
    procedure FormShow(Sender: TObject);
    procedure btnRedoClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure Bit_PreviousClick(Sender: TObject);
    procedure MKE_StartExit(Sender: TObject);
    procedure Btn_PrintClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure chkStampTaxClick(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
  private
    { Private declarations }
    FApplyByMainInv: Boolean;
    FImportNote: String;
    FDebugList: TStringList;
    //計算媒體申報主檔之格式
    // 變數 - 說明
    //P_RecNo - 資料筆數
    //P_Inv01BusID - 銷售人統一編號
    //P_Inv001TaxID - 銷售人稅籍編號
    //P_Table - 資料表名稱
    //P_Select - 執行選項
    //P_MediaList - 輸出暫存變數
    Procedure SetMasterData(Var P_RecNo:String;P_Inv01BusID,P_Inv001TaxID,
              P_Table:String;P_Select:Integer;Var P_MediaList:TStringList);

    //計算統計金額公式
    // 變數 - 說明
    //dselect - 選擇sql語法，1金額合計、2資料筆數、3發票
    //dField - 需統計金額的欄位名稱
    //dTable - 資料表名稱
    //dParma1 - 參數一，IsObsolete，不為作廢的資料，空白或有值
    //dParma2 - 參數二，InvFormat1，發票格式1，空白或填入1、2、3
    //dParma3 - 參數三，InvFormat2，發票格式2，空白或填入1、3
    //dParma4 - 參數四，BusinessId，客戶統編，空白或1統編為空白、2統編不為空白
    //dParma5 - 參數五，TaxType，發票稅別，空白或有值
    function SetMediaData(dselect:integer;dField:String; dTable:String;
             dParma1, dParma2, dParma3, dParma4, dParma5:String):Double ;

    function GetCompIdList: String;

  public
    { Public declarations }
  end;

var
  frmInvA08: TfrmInvA08;
  fL_MediaFile : String ;//主檔儲存路徑
  fL_MediaFile_B : String ;//主檔備份儲存路徑
  fL_MediaReport : String ;//明細檔儲存路徑
  fL_MediaREport_B : String ;//明細檔備份儲存路徑
  sL_CreateMonthStart : String ;//申報起始月份
  sL_CreateMonthEnd : String ;//申報結止月份
  sL_MakeMonth : String ;//申報月份
  sL_StampTaxt: String;
  FCompanyList : TList;

  aDebugPath, aDebugFileName: String;

implementation

uses frmMainU, frmInvD02_1U, dtmMainJU, dtmMainU, dtmSOU, cbUtilis;

{$R *.dfm}

function TfrmInvA08.GetCompIdList: String;
var
  aIndex: Integer;
begin
  Result := Format( '''%s''',[dtmMain.getCompID] );
  if ( FCompanyList.Count > 1 ) and ( RadioGroup2.ItemIndex = 1 ) then
  begin
    Result := EmptyStr;
    for aIndex := 0 to FCompanyList.Count - 1 do
    begin
      Result := Result + Format(' ''%s'' ',[PCompany(FCompanyList[aIndex]).CompanyId] );
      if ( aIndex < ( FCompanyList.Count - 1 ) ) then
        Result := ( Result + ',' );
    end;    
  end;
  Result := Format( ' ( %s ) ', [Result] );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA08.FormShow(Sender: TObject);
var
  aSQL : String;
begin
  FDebugList := TStringList.Create;
  PCT1.ActivePageIndex := 0;
  self.Caption := frmMain.GetFormTitleString('A08','產生媒體申報');
  lblCompany.Caption := Format( '%s ( %s )',
    [dtmMain.getCompID, dtmMain.getCompName] );
  lblUser.Caption := dtmMain.getLoginUser;

  aSQL := ' SELECT INVPARAM FROM INV003 WHERE COMPID = ' +
    QuotedStr( dtmMain.getCompID );

  Edt_MediaFilePath.Text := dtmMainJ.GetXMLAttribute( aSQL, 'MediaFilePath' );
  CBx_Month.ItemIndex := 0;
  MKE_Year.SetFocus;
  MKE_Year.Text := copy(datetostr(date()),1,4);
  CheckBox1.Enabled := dtmMain.GetLinkToMIS;

  FCompanyList := TList.Create;
  dtmMain.GetInvoiceCompany( FCompanyList );
  RadioGroup2.Enabled := FCompanyList.Count > 1;
  Label9.Enabled := FCompanyList.Count > 1;
  {}
  FImportNote := dtmMainj.OpenSQL( Format(
    ' select importnote from inv003  ' +
    '  where identifyid1 = ''%s''    ' +
    '    and identifyid2 = ''%s''    ' +
    '    and compid = ''%s''         ',
    [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID] ) );
    
  aDebugPath := IncludeTrailingPathDelimiter( Edt_MediaFilePath.Text ) + 'Debug';
  aDebugFileName := 'SqlTrace.txt';
  if not DirectoryExists( aDebugPath ) then ForceDirectories( aDebugPath );

end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA08.FormDestroy(Sender: TObject);
var
  aIndex: Integer;
begin
  for aIndex := 0 to FCompanyList.Count - 1 do
  begin
    if Assigned( FCompanyList[aIndex] ) then
       Dispose( PCompany( FCompanyList[aIndex] ) );
    FCompanyList[aIndex] := nil;
  end;
  FCompanyList.Clear;
  FCompanyList.Free;
  FDebugList.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA08.btnRedoClick(Sender: TObject);
var
  aSQL : String;
begin
  MKE_Year.Clear;
  CBx_Month.ItemIndex := 0;
  RadioGroup1.ItemIndex := 2;
  MKE_Start.Clear;
  MKE_End.Clear;
  CheckBox1.Checked := False;
  aSQL := ' SELECT INVPARAM FROM INV003 WHERE COMPID = ' +
    QuotedStr( dtmMain.getCompID );
  Edt_MediaFilePath.Text := dtmMainJ.GetXMLAttribute( aSQL, 'MediaFilePath' );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA08.btnExitClick(Sender: TObject);
begin
   Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA08.btnOKClick(Sender: TObject);
Var
  //主檔變數
  sL_Inv01BusinessID    : String    ;//公司統編
  sL_Inv001TaxID        : String    ;//稅籍編號
  MediaList             : TStringList;//媒體申報暫存
  SL_RecNo              : String    ;//序號資料
  //明細檔變數
  tL_Sum                : TarraySum;
  StDate,EdDate         : Tdatetime;
  aCompIdList: String;
begin
  if MKE_Year.Text = '' then
  begin
    MessageDlg('請輸入申報之年份!',mtError,[mbOK],0);
    exit;
  end;

  if (Trim(MKE_Start.Text) <> '') then
  begin
    if (Trim(MKE_End.Text) = '') then
    begin
      MessageDlg('請輸入發票截止號碼!',mtError,[mbOK],0);
      exit;
    end;
    if length(trim(mke_end.Text)) < 10 then
    begin
      MessageDlg('請輸入正確之發票截止號碼!',mtError,[mbOK],0);
      exit;
    end;
    if length(trim(MKE_Start.Text)) < 10 then
    begin
      MessageDlg('請輸入正確之發票起始號碼!',mtError,[mbOK],0);
      exit;
    end;
    MKE_Start.Tag := 1;
  end
  else
  begin
    if (Trim(MKE_End.Text) <> '') then
    begin
      MessageDlg('請輸入發票起始號碼!',mtError,[mbOK],0);
      exit;
    end;
    MKE_Start.Tag := 0;
  end;

  PCT1.ActivePageIndex := 1;
  StDate := now;
  try
    Bit_Previous.Enabled := False;
    Btn_Print.Enabled := False;
    btnExit.Enabled := False;
    ProgressBar2.Position := 0;
    SL_RecNo := '1';
    sL_MakeMonth := MKE_Year.Text +CBx_Month.Items[CBx_Month.ItemIndex];//組合申報年月
    //取出inv001的公司統編、稅籍編號
    with TADOQuery.Create(self) do
    begin
      Connection := dtmMain.InvConnection;
      SQL.Clear;
      SQL.Add( 'SELECT BUSINESSID, TAXID FROM INV001 ' );
      SQL.Add( ' WHERE IDENTIFYID1 = :P1 ' );
      SQL.Add( '   AND IDENTIFYID2 = :P2 ');
      SQL.Add( '   AND COMPID = :P3 ' );

      Parameters.ParamByName('P1').Value := IDENTIFYID1;
      Parameters.ParamByName('P2').Value := IDENTIFYID2;
      Parameters.ParamByName('P3').Value := dtmMain.getCompID;

      Open;

      sL_Inv01BusinessID := FieldByName('BusinessID').AsString;
      sL_Inv001TaxID := FieldByName('TaxID').AsString;
      close;
      free;
    end;

    //計算申報月份類別
    case RadioGroup1.ItemIndex of
      0://單月
      begin
        sL_CreateMonthStart := getYearMonthDay8(sL_MakeMonth,1);
        sL_CreateMonthEnd :=getYearMonthDay8(sL_MakeMonth,2);
      end;
      1://雙月
      begin
        sL_CreateMonthStart :=getYearMonthDay8(sL_MakeMonth,3);
        sL_CreateMonthEnd :=getYearMonthDay8(sL_MakeMonth,4);
      end;
      2://二個月
      begin
        sL_CreateMonthStart :=getYearMonthDay8(sL_MakeMonth,1);
        sL_CreateMonthEnd :=getYearMonthDay8(sL_MakeMonth,4);
      end;
    end;

    MediaList := TStringList.Create;

    {}
    FApplyByMainInv := chkApplyMainInv.Checked;
    {}

{=============================計算申報主檔資料=================================}
    //計算發票主檔資料
    SL_RecNo := '1';
    lbl_show.Caption := '發票主檔資料處理進度';
    ProgressBar2.Position := 0;
    Refresh;
    Screen.Cursor := crSQLWait;
    try
      SetMasterData(SL_RecNo,sL_Inv01BusinessID,sL_Inv001TaxID,
                    'INV007', 1, MediaList );
    finally
      Screen.Cursor := crDefault;
    end;
    lbl_show.Caption := '發票主檔資料處理完成';
    frmInvA08.Refresh;

    { 單據申報,不產生未使用發票記錄 }
    if ( not chkPaper.Checked ) then
    begin
      //計算空白資料
      lbl_show.Caption := '未使用發票資料處理進度';
      ProgressBar2.Position := 0;
      Refresh;
      Screen.Cursor := crSQLWait;
      try
        SetMasterData(SL_RecNo,sL_Inv01BusinessID,sL_Inv001TaxID,
                      'INV099', 3, MediaList );
      finally
        Screen.Cursor := crDefault;
      end;
      lbl_show.Caption := '未使用發票資料處理完成';
      frmInvA08.Refresh;
    end;  

//    SL_RecNo := '1';
    //計算銷貨退回 或 折讓證明單資料
    lbl_show.Caption := '銷貨退回 或 折讓證明單資料處理進度';
    ProgressBar2.Position := 0;
    Refresh;
    Screen.Cursor := crSQLWait;
    try
      SetMasterData(SL_RecNo,sL_Inv01BusinessID,sL_Inv001TaxID,
                    'INV014',2,MediaList);
    finally
      Screen.Cursor := crDefault;
    end;                
    lbl_show.Caption := '銷貨退回 或 折讓證明單資料處理完成';
    frmInvA08.Refresh;

//    SL_RecNo := '1';//序號如要重新計算時，記得打開
    //計算印花稅明細資料
    if CheckBox1.Checked then
    begin
      lbl_show.Caption := '印花稅明細資料處理進度';
      ProgressBar2.Position := 0;
      Refresh;
      Screen.Cursor := crSQLWait;
      try
        SetMasterData( SL_RecNo, sL_Inv01BusinessID, sL_Inv001TaxID,
                      dtmMain.getMisDbOwner+ '.SO033', 4, MediaList );
      finally
        Screen.Cursor := crDefault;
      end;                
      lbl_show.Caption := '印花稅明細資料處理完成';
      frmInvA08.Refresh;
    end;

    
    //印花稅總繳
    if chkStampTax.Checked then
    begin
      lbl_show.Caption := '印花稅總繳資料處理進度';
      ProgressBar2.Position := 0;
      Refresh;
      Screen.Cursor := crSQLWait;
      try
        sL_StampTaxt := FloatToStr( txtStampTax.Value );
        SetMasterData( SL_RecNo, sL_Inv01BusinessID, sL_Inv001TaxID,
                      EmptyStr, 5, MediaList );
      finally
        Screen.Cursor := crDefault;
      end;
      lbl_show.Caption := '印花稅總繳資料處理完成';
      frmInvA08.Refresh;
    end;

    if ( chkDebugFile.Checked ) then
    begin
      FDebugList.SaveToFile( IncludeTrailingPathDelimiter( aDebugPath ) + aDebugFileName );
      FDebugList.Clear;
    end;


    case RadioGroup1.ItemIndex of
      0:
      begin
        fL_MediaFile :=Edt_MediaFilePath.Text +'\'+sL_Inv01BusinessID+'.txt';
        fL_MediaFile_B :=Edt_MediaFilePath.Text +'\'+sL_Inv01BusinessID+'_'+sL_MakeMonth+'_單.txt';
        fL_MediaReport :=Edt_MediaFilePath.Text +'\'+sL_Inv01BusinessID+'_Report.txt';
        fL_MediaReport_B :=Edt_MediaFilePath.Text +'\'+sL_Inv01BusinessID+'_Report_'+sL_MakeMonth+'_單.txt';
      end;
      1:
      begin
        fL_MediaFile :=Edt_MediaFilePath.Text +'\'+sL_Inv01BusinessID+'.txt';
        fL_MediaFile_B :=Edt_MediaFilePath.Text +'\'+sL_Inv01BusinessID+'_'+sL_MakeMonth+'_雙.txt';
        fL_MediaReport :=Edt_MediaFilePath.Text +'\'+sL_Inv01BusinessID+'_Report.txt';
        fL_MediaReport_B :=Edt_MediaFilePath.Text +'\'+sL_Inv01BusinessID+'_Report_'+sL_MakeMonth+'_雙.txt';
      end;
      2:
      begin
        fL_MediaFile :=Edt_MediaFilePath.Text +'\'+sL_Inv01BusinessID+'.txt';
        fL_MediaFile_B :=Edt_MediaFilePath.Text +'\'+sL_Inv01BusinessID+'_'+sL_MakeMonth+'_all.txt';
        fL_MediaReport :=Edt_MediaFilePath.Text +'\'+sL_Inv01BusinessID+'_Report.txt';
        fL_MediaReport_B :=Edt_MediaFilePath.Text +'\'+sL_Inv01BusinessID+'_Report_'+sL_MakeMonth+'_all.txt';
      end;
    end;

    MediaList.SaveToFile(fL_MediaFile);
    MediaList.SaveToFile(fL_MediaFile_B);
    MediaList.Clear;

    MediaList.Add('請至【主機】下列位置領取檔案：');
    MediaList.Add(' ');
    MediaList.Add('媒體申報檔     ==＞ '+fL_MediaFile);
    MediaList.Add('媒體明細檔     ==＞ '+fL_MediaReport);
    MediaList.Add('媒體申報備份檔 ==＞ '+fL_MediaFile_B);
    MediaList.Add('申報明細備份檔 ==＞ '+fL_MediaREport_B);

{====================以下為直接抓取檔案中之資料(應稅)==========================}
    MediaList.Add(' ');
    MediaList.Add('以下為直接抓取檔案中之資料(應稅)');
    //A.  計算非營業人(電子)發票額總計
    tL_Sum := TarraySum.Create;
    lbl_show.Caption := '非營業人(電子)發票額總計處理';
    Refresh;
    tL_Sum.sL_A := SetMediaData(1,'InvAmount','INV007', 'N','1','','1','1');
    MediaList.Add('非營業人(電子)發票額總計：'+FormatFloat('###,###,###,##0',tL_Sum.sL_A));

    //B.  非營業人(電子)銷售額總計
    lbl_show.Caption := '非營業人(電子)銷售額總計處理';
    Refresh;
    tL_Sum.sL_B := SetMediaData(1,'SaleAmount','INV007', 'N','1','','1','1');
    MediaList.Add('非營業人(電子)銷售額總計：'+FormatFloat('###,###,###,##0',tL_Sum.sL_B));

    //C.  非營業人(電子)稅額總計
    lbl_show.Caption := '非營業人(電子)稅額總計處理';
    Refresh;
    tL_Sum.sL_C := SetMediaData(1,'TaxAmount','INV007', 'N','1','','1','1');
    MediaList.Add('非營業人(電子)稅  額總計：'+FormatFloat('###,###,###,##0',tL_Sum.sL_C));

    //D.  營業人(電子)發票額總計
    lbl_show.Caption := '營業人(電子)發票額總計處理';
    Refresh;
    tL_Sum.sL_D := SetMediaData(1,'InvAmount','INV007', 'N','1','','2','1');
    MediaList.Add('營業人(電子)發票額總計：'+FormatFloat('###,###,###,##0',tL_Sum.sL_D));

    //E.	營業人(電子)銷售額總計
    lbl_show.Caption := '營業人(電子)銷售額總計處理';
    Refresh;
    tL_Sum.sL_E := SetMediaData(1,'SaleAmount','INV007', 'N','1','','2','1');
    MediaList.Add('營業人(電子)銷售額總計：'+FormatFloat('###,###,###,##0',tL_Sum.sL_E));

    //F.	營業人(電子)稅額總計
    lbl_show.Caption := '營業人(電子)稅額總計處理';
    Refresh;
    tL_Sum.sL_F := SetMediaData(1,'TaxAmount','INV007', 'N','1','','2','1');
    MediaList.Add('營業人(電子)稅  額總計：'+FormatFloat('###,###,###,##0',tL_Sum.sL_F));

{====================以下為<計算資料>媒體申報之計算============================}
    MediaList.Add(' ');
    MediaList.Add('以下為＜計算資料＞媒體申報之計算');
    //G.	非營業人--電子(應稅)發票額總計
    lbl_show.Caption := '非營業人--電子(應稅)發票額總計處理';
    Refresh;
    tL_Sum.sL_G := SetMediaData(1,'InvAmount','INV007', 'N','1','','1','1');
    MediaList.Add('非營業人--電子(應稅)發票額總計：'+FormatFloat('###,###,###,##0',tL_Sum.sL_G));

    //H.	非營業人--電子(應稅)銷售額總計{發票額 / 1.05}
    lbl_show.Caption := '非營業人--電子(應稅)銷售額總計處理';
    Refresh;
    tL_Sum.sL_H := SimpleRoundTo( tL_Sum.sl_G / 1.05,0);
    MediaList.Add('非營業人--電子(應稅)銷售額總計(發票額/1.05)：'+FormatFloat('###,###,###,##0',tL_Sum.sL_H));

    //I.	非營業人--電子(應稅)稅　額總計{銷售額 × 0.05}
    lbl_show.Caption := '非營業人--電子(應稅)稅　額總計處理';
    Refresh;
    tL_Sum.sL_I :=  SimpleRoundTo( tL_Sum.sl_H * 0.05,0 );
    MediaList.Add('非營業人--電子(應稅)稅　額總計(銷售額*0.05)：'+FormatFloat('###,###,###,##0',tL_Sum.sL_I));

    //J.	營業人--電子(應稅)銷售額總計
    lbl_show.Caption := '營業人--電子(應稅)銷售額總計處理';
    Refresh;
    tL_Sum.sL_J := SetMediaData(1,'SaleAmount','INV007', 'N','1','','2','1');
    MediaList.Add('營業人--電子(應稅)銷售額總計：'+FormatFloat('###,###,###,##0',tL_Sum.sL_J));

    //K.	營業人--三聯(應稅)銷售額總計
    lbl_show.Caption := '營業人--三聯(應稅)銷售額總計處理';
    Refresh;
    tL_Sum.sL_K := SetMediaData(1,'SaleAmount','INV007','N','3','','2','1');
    MediaList.Add('營業人--三聯(應稅)銷售額總計：'+FormatFloat('###,###,###,##0',tL_Sum.sL_K));

    //L.	營業人--三聯、電子(應稅)銷售額總計
    lbl_show.Caption := '營業人--三聯、電子(應稅)銷售額總計處理';
    Refresh;
    tL_Sum.sL_L := tL_Sum.sL_J + tL_Sum.sL_K;
    MediaList.Add('營業人--三聯、電子(應稅)銷售額總計：'+FormatFloat('###,###,###,##0',tL_Sum.sL_L));

    //M.	營業人--電子(應稅)稅額總計
    lbl_show.Caption := '營業人--電子(應稅)稅額總計處理';
    Refresh;
    tL_Sum.sL_M := SetMediaData(1,'TaxAmount','INV007', 'N','1','','2','1');
    MediaList.Add('營業人--電子(應稅)稅額總計：'+FormatFloat('###,###,###,##0',tL_Sum.sL_M));

    //N.	營業人--三聯(應稅)稅額總計
    lbl_show.Caption := '營業人--三聯(應稅)稅額總計處理';
    Refresh;
    tL_Sum.sL_N := SetMediaData(1,'TaxAmount','INV007', 'N','3','','2','1');
    MediaList.Add('營業人--三聯(應稅)稅額總計：'+FormatFloat('###,###,###,##0',tL_Sum.sL_N));

    //O.	營業人--三聯、電子(應稅)稅額總計
    lbl_show.Caption := '營業人--三聯、電子(應稅)稅額總計處理';
    Refresh;
    tL_Sum.sL_O := tL_Sum.sL_M + tL_Sum.sL_N;
    MediaList.Add('營業人--三聯、電子(應稅)稅額總計：'+FormatFloat('###,###,###,##0',tL_Sum.sL_O));

    //P.	二聯(應稅)發票額總計
    lbl_show.Caption := '二聯(應稅)發票額總計處理';
    Refresh;
    tL_Sum.sL_P := SetMediaData(1,'InvAmount','INV007', 'N','2','','','1');
    MediaList.Add('二聯(應稅)發票額總計：'+FormatFloat('###,###,###,##0',tL_Sum.sL_P));

    //Q.	二聯(應稅)銷售額總計{發票額 / 1.05}
    lbl_show.Caption := '二聯(應稅)銷售額總計處理';
    Refresh;
    tL_Sum.sL_Q := SimpleRoundTo( tL_Sum.sL_P / 1.05,0);
    MediaList.Add('二聯(應稅)銷售額總計(發票額/1.05)：'+FormatFloat('###,###,###,##0',tL_Sum.sL_Q));

    //R.	二聯(應稅)稅　額總計{銷售額 × 0.05}
    lbl_show.Caption := '二聯(應稅)稅　額總計處理';
    Refresh;
    tL_Sum.sL_R := SimpleRoundTo( tL_Sum.sL_Q * 0.05 ,0);
    MediaList.Add('二聯(應稅)稅　額總計(銷售額*0.05)：'+FormatFloat('###,###,###,##0',tL_Sum.sL_R));

    //S.	營業/非營業人-折讓(應稅)銷售額總計
    lbl_show.Caption := '營業/非營業人-折讓(應稅)銷售額總計處理';
    Refresh;
    tL_Sum.sL_S := SetMediaData( 4, 'SaleAmount','INV014', '','','','','1');
    MediaList.Add('營業/非營業人-折讓(應稅)銷售額總計：'+FormatFloat('###,###,###,##0',tL_Sum.sL_S));

    //T.	營業/非營業人-折讓(應稅)稅額總計
    lbl_show.Caption := '營業/非營業人-折讓(應稅)稅額總計處理';
    Refresh;
    tL_Sum.sL_T := SetMediaData(4,'TaxAmount','INV014', '','','','','1');
    MediaList.Add('營業/非營業人-折讓(應稅)稅  額總計：'+FormatFloat('###,###,###,##0',tL_Sum.sL_T));

{=========================以下為媒體申報金額統計===============================}
    MediaList.Add(' ');
    if ( chkPaper.Checked ) then
      MediaList.Add('以下為＜媒體申報書資料＞金額統計--以收據申報')
    else
      MediaList.Add('以下為＜媒體申報書資料＞金額統計');

    //U.	三聯/電子(應稅-銷售)總計
    lbl_show.Caption := '三聯/電子(應稅-銷售)總計處理';
    Refresh;
    tL_Sum.sL_U := tL_Sum.sL_H + tL_Sum.sL_L;
    MediaList.Add('三聯/電子(應稅-銷售)總計：'+FormatFloat('###,###,###,##0',tL_Sum.sL_U));

    //V.	三聯/電子(應稅-稅額)總計
    lbl_show.Caption := '三聯/電子(應稅-稅額)總計處理';
    Refresh;
    tL_Sum.sL_V := tL_Sum.sL_I + tL_Sum.sL_O;
    MediaList.Add('三聯/電子(應稅-稅額)總計：'+FormatFloat('###,###,###,##0',tL_Sum.sL_V));

    //W.	二聯(應稅-銷售)總計
    lbl_show.Caption := '二聯(應稅-銷售)總計處理';
    Refresh;
    tL_Sum.sL_W := tL_Sum.sL_Q;
    MediaList.Add('二聯(應稅-銷售)總計：'+FormatFloat('###,###,###,##0',tL_Sum.sL_W));

    //X.	二聯(應稅-稅額)總計
    lbl_show.Caption := '二聯(應稅-稅額)總計處理';
    Refresh;
    tL_Sum.sL_X := tL_Sum.sL_R;
    MediaList.Add('二聯(應稅-稅額)總計：'+FormatFloat('###,###,###,##0',tL_Sum.sL_X));

    //Y.	折讓(應稅-銷售)總計
    lbl_show.Caption := '折讓(應稅-銷售)總計處理';
    Refresh;
    tL_Sum.sL_Y := tL_Sum.sL_S;
    MediaList.Add('折讓(應稅-銷售)總計：'+FormatFloat('###,###,###,##0',tL_Sum.sL_Y));

    //Z.	折讓(應稅-稅額)總計
    lbl_show.Caption := '折讓(應稅-稅額)總計處理';
    Refresh;
    tL_Sum.sL_Z := tL_Sum.sL_T;
    MediaList.Add('折讓(應稅-稅額)總計：'+FormatFloat('###,###,###,##0',tL_Sum.sL_Z));

    //AA.	合計(應稅-銷售)總計
    lbl_show.Caption := '合計(應稅-銷售)總計處理';
    Refresh;
    tL_Sum.sL_AA := tL_Sum.sL_U + tL_Sum.sL_W - tL_Sum.sL_Y;
    MediaList.Add('合計(應稅-銷售)總計：'+FormatFloat('###,###,###,##0',tL_Sum.sL_AA));

    //BB.	合計(應稅-稅額)總計
    lbl_show.Caption := '合計(應稅-稅額)總計處理';
    Refresh;
    tL_Sum.sL_BB := tL_Sum.sL_V + tL_Sum.sL_X - tL_Sum.sL_Z;
    MediaList.Add('合計(應稅-稅額)總計：'+FormatFloat('###,###,###,##0',tL_Sum.sL_BB));

    //CC.	三聯/電子(免稅)總計
    lbl_show.Caption := '三聯/電子(免稅)總計處理';
    Refresh;
    tL_Sum.sL_CC := SetMediaData(1,'Saleamount','INV007', 'N','1','3','','3');
    MediaList.Add('三聯/電子(免稅)總計：'+FormatFloat('###,###,###,##0',tL_Sum.sL_CC));

    //DD.	二聯(免稅)總計
    lbl_show.Caption := '二聯(免稅)總計處理';
    Refresh;
    tL_Sum.sL_DD := SetMediaData(1,'Saleamount','INV007', 'N','2','','','3');
    MediaList.Add('二聯(免稅)總計：'+FormatFloat('###,###,###,##0',tL_Sum.sL_DD));

    //EE.	折讓(免稅)總計
    lbl_show.Caption := '折讓(免稅)總計處理';
    Refresh;
    tL_Sum.sL_EE := SetMediaData( 5,'SaleAmount','INV014', '','','','','3');
    MediaList.Add('折讓(免稅)總計：'+FormatFloat('###,###,###,##0',tL_Sum.sL_EE));

    //FF.	合計(免稅)
    lbl_show.Caption := '合計(免稅)處理';
    Refresh;
    tL_Sum.sL_FF := tL_Sum.sL_CC + tL_Sum.sL_DD - tL_Sum.sL_EE;
    MediaList.Add('合計(免稅)：'+FormatFloat('###,###,###,##0',tL_Sum.sL_FF));


    //CC2. 三聯/電子(零稅率)總計
    lbl_show.Caption := '三聯/電子(零稅率)總計處理';
    Refresh;
    tL_Sum.sL_CC2 := SetMediaData( 1,'Saleamount','INV007', 'N', '1', '3', '', '2' );
    MediaList.Add('三聯/電子(零稅率)總計：'+FormatFloat('###,###,###,##0',tL_Sum.sL_CC2 ) );


    //DD2. 二聯(零稅率) 總計
    lbl_show.Caption := '二聯(零稅率)總計處理';
    Refresh;
    tL_Sum.sL_DD2 := SetMediaData( 1,'Saleamount','INV007', 'N', '2', '', '', '2' );
    MediaList.Add('二聯(零稅率)總計：'+FormatFloat('###,###,###,##0',tL_Sum.sL_DD2 ) );


    //EE2. 折讓(零稅率) 總計
    lbl_show.Caption := '折讓(零稅率)總計處理';
    Refresh;
    tL_Sum.sL_EE2 := SetMediaData( 5,'SaleAmount','INV014', '','','','','2' );
    MediaList.Add('折讓(零稅率)總計：'+FormatFloat('###,###,###,##0',tL_Sum.sL_EE2 ) );


    //FF2. 合計(零稅率)
    lbl_show.Caption := '合計(零稅率)處理';
    Refresh;
    tL_Sum.sL_FF2 := ( tL_Sum.sL_CC2 + tL_Sum.sL_DD2 - tL_Sum.sL_EE2 );
    MediaList.Add('合計(零稅率)：'+FormatFloat('###,###,###,##0', tL_Sum.sL_FF2 ) );



    //GG.	銷售額總計 ( AA + FF + FF2 )
    lbl_show.Caption := '銷售額總計處理';
    Refresh;
    tL_Sum.sL_GG := tL_Sum.sL_AA + tL_Sum.sL_FF + tL_Sum.sL_FF2;
    MediaList.Add('銷售額總計：'+FormatFloat('###,###,###,##0',tL_Sum.sL_GG));


    if chkStampTax.Checked then
    begin
      //GG2. 印花稅總繳總額
      lbl_show.Caption := '印花稅總繳總額';
      Refresh;
      tL_Sum.sL_GG2 := txtStampTax.Value;
      MediaList.Add('印花稅總繳總額：'+FormatFloat('###,###,###,##0',tL_Sum.sL_GG2 ) );
    end;

{=====================以下為<媒體申報參考資料>張數統計表=======================}
    MediaList.Add(' ');
    MediaList.Add('以下為＜媒體申報參考資料＞張數統計表');
    //HH.	電子-應稅張數統計
    lbl_show.Caption := '電子-應稅張數統計處理';
    Refresh;
    tL_Sum.sL_HH := SetMediaData(2,'','INV007', 'N','1','','','1');

    //II.	二聯-應稅張數統計
    lbl_show.Caption := '二聯-應稅張數統計處理';
    Refresh;
    tL_Sum.sL_II := SetMediaData(2,'','INV007', 'N','2','','','1');

    //JJ.	三聯-應稅張數統計
    lbl_show.Caption := '三聯-應稅張數統計處理';
    Refresh;
    tL_Sum.sL_JJ := SetMediaData(2,'','INV007', 'N','3','','','1');
    MediaList.Add('(應稅)==＞電子:'+adjString2(10,tL_Sum.sL_HH)+
                           ' 二聯:'+adjString2(10,tL_Sum.sL_II)+
                           ' 三聯:'+adjString2(10,tL_Sum.sL_JJ));

    { -------------------------------------------------------- }

    //KK2. 電子-零稅率張數統計
    lbl_show.Caption := '電子-零稅率張數統計處理';
    Refresh;
    tL_Sum.sL_KK2 := SetMediaData( 2 , '', 'INV007', 'N', '1', '',  '', '2' );

    //LL2. 二聯-零稅率張數統計
    lbl_show.Caption := '二聯-零稅率張數統計處理';
    Refresh;
    tL_Sum.sL_LL2 := SetMediaData( 2 , '', 'INV007', 'N', '2', '',  '', '2' );

    //MM2. 三聯-零稅率張數統計
    lbl_show.Caption := '二聯-零稅率張數統計處理';
    Refresh;
    tL_Sum.sL_MM2 := SetMediaData( 2 , '', 'INV007', 'N', '3', '',  '', '2' );
    MediaList.Add('(零稅率)==＞電子:' + adjString2(10,tL_Sum.sL_KK2 ) +
                           '   二聯:' + adjString2(10,tL_Sum.sL_LL2 ) +
                           '   三聯:' + adjString2(10,tL_Sum.sL_MM2 ) );

    { -------------------------------------------------------- }


    //KK.	電子-免稅張數統計
    lbl_show.Caption := '電子-免稅張數統計處理';
    Refresh;
    tL_Sum.sL_KK := SetMediaData(2,'','INV007', 'N','1','','','3');

    //LL.	二聯-免稅張數統計
    lbl_show.Caption := '二聯-免稅張數統計處理';
    Refresh;
    tL_Sum.sL_LL := SetMediaData(2,'','INV007', 'N','2','','','3');

    //MM.	三聯-免稅張數統計
    lbl_show.Caption := '三聯-免稅張數統計處理';
    Refresh;
    tL_Sum.sL_MM := SetMediaData(2,'','INV007', 'N','3','','','3');
    MediaList.Add('(免稅)==＞電子:'+adjString2(10,tL_Sum.sL_KK)+
                           ' 二聯:'+adjString2(10,tL_Sum.sL_LL)+
                           ' 三聯:'+adjString2(10,tL_Sum.sL_MM));

    { -------------------------------------------------------- }
    // 2.單據申報,已作廢資料不需產生至媒體申報檔
    if ( not chkPaper.Checked ) then
    begin
      //NN.	電子-作廢張數統計
      lbl_show.Caption := '電子-作廢張數統計處理';
      Refresh;
      tL_Sum.sL_NN := SetMediaData(2,'','INV007', 'Y','1','','','');

      //OO.	二聯-作廢張數統計
      lbl_show.Caption := '二聯-作廢張數統計處理';
      Refresh;
      tL_Sum.sL_OO := SetMediaData(2,'','INV007', 'Y','2','','','');

      //PP.	三聯-作廢張數統計
      lbl_show.Caption := '三聯-作廢張數統計處理';
      Refresh;
      tL_Sum.sL_PP := SetMediaData(2,'','INV007', 'Y','3','','','');
      MediaList.Add('(作廢)==＞電子:'+adjString2(10,tL_Sum.sL_NN)+
                             ' 二聯:'+adjString2(10,tL_Sum.sL_OO)+
                             ' 三聯:'+adjString2(10,tL_Sum.sL_PP));
    end;
    
    { -------------------------------------------------------- }
    
    //QQ.	電子-空白張數統計
    lbl_show.Caption := '電子-空白張數統計處理';
    Refresh;
    tL_Sum.sL_QQ := SetMediaData(3,'','INV099','','1','','','');

    //RR.	二聯-空白張數統計
    lbl_show.Caption := '二聯-空白張數統計處理';
    Refresh;
    tL_Sum.sL_RR := SetMediaData(3,'','INV099','','2','','','');

    //SS.	三聯-空白張數統計
    lbl_show.Caption := '三聯-空白張數統計處理';
    Refresh;
    tL_Sum.sL_SS := SetMediaData(3,'','INV099','','3','','','');
    MediaList.Add('(空白)==＞電子:'+adjString2(10,tL_Sum.sL_QQ)+
                           ' 二聯:'+adjString2(10,tL_Sum.sL_RR)+
                           ' 三聯:'+adjString2(10,tL_Sum.sL_SS));
                           

{===============以下為 <空白未使用之統一發票字軌及號碼一覽表>==================}
    MediaList.Add(' ');
    case RadioGroup1.ItemIndex of
      0: mediaList.Add('以下為'+
         getYearMonthDay7(strtodate(sL_CreateMonthStart),1) +
         '＜空白未使用之統一發票字軌及號碼一覽表＞');
      1: mediaList.Add('以下為'+
         getYearMonthDay7(strtodate(sL_CreateMonthEnd),1) +
         '＜空白未使用之統一發票字軌及號碼一覽表＞');
      2: mediaList.Add('以下為'+
         getYearMonthDay7(strtodate(sL_CreateMonthStart),1) +' ～ '+
         getYearMonthDay7(strtodate(sL_CreateMonthEnd),1) +
         '＜空白未使用之統一發票字軌及號碼一覽表＞');
    end;
    MediaList.Add('字軌　　 發票起訖號碼　　　發票種類　　 張數');
    MediaList.Add('==== ==================== ========== ==========');


    aCompIdList := GetCompIdList;

    with AQry_Inv099 do
    begin
      sql.Text := 'Select Prefix, CurNum, EndNum, InvFormat, (EndNum-CurNum+1) NumCount '+
                  'From Inv099 Where yearmonth = ''' + sL_MakeMonth + '''' +
                  ' and compid in ' + aCompIdList + ' order by yearmonth, prefix, startnum ';
      open;
      if recordcount = 0 then
      begin
        tL_Sum.sL_TT := ' ';
        tL_Sum.sL_UU := ' ';
        tL_Sum.sL_VV := ' ';
        tL_Sum.sL_WW := 0;
      end
      else
      begin
        while not eof do
        begin
          tL_Sum.sL_TT := fieldByName('Prefix').asstring;
          // 2006/1/10 Jim 發現 CurNum 若為開滿的話會變成 已使用號碼:500 ~ 迄號: 499 }
          // 若是已經都開完就不要再 show 出來, 待做, 還沒改 }
          tL_Sum.sL_UU := fieldByName('CurNum').asstring+' ～ ' +fieldByName('EndNum').asstring;
          if fieldByName('InvFormat').asstring = '1' then
            tL_Sum.sl_VV := '電算發票'
          else if fieldByName('InvFormat').asstring = '2' then
            tL_Sum.sl_VV := '二聯(副)'
          else if fieldByName('InvFormat').asstring = '3' then
            tL_Sum.sl_VV := '三聯(副)';

          tL_Sum.sL_WW := fieldByName('NumCount').AsInteger;
          MediaList.Add(' '+tL_Sum.sL_TT+'　'+
                            tL_Sum.sL_UU+'　'+
                            tL_Sum.sl_VV+'　'+
                            adjString2(10,tL_Sum.sL_WW));
          tl_sum.sL_ZZ := tl_sum.sL_ZZ + tl_sum.sL_WW;
          next;
        end;
      end;
    end;

    MediaList.Add('　　　　　　 　　　　　　　合　計：　'+adjString2(10,tL_Sum.sL_ZZ));
    MediaList.Add(' ');
    MediaList.SaveToFile(fL_MediaReport);
    MediaList.SaveToFile(fL_MediaReport_B);

    ListBox1.Items := MediaList;
    eddate := now;
    Lbl_show.Caption := '媒體申報檔產生完成,總共花費      '+
                        timetostr(eddate - StDate)+' 秒';
  except
    on E: Exception do
    begin
      Lbl_show.Caption := '媒體申報檔產生失敗。'#13#10 + E.Message;
    end;  
  end;
  Bit_Previous.Enabled := True;
  Btn_Print.Enabled := True;
  btnExit.Enabled := True;
end;

function tfrminvA08.SetMediaData(dselect:integer;dField:String; dTable:String;
                    dParma1, dParma2, dParma3, dParma4, dParma5:String):Double ;
Var
  sL_sql : String;
  sl_parma : String;
  dL_sum : Double;
  aYearMonth: String;
  aCompIdList: String;
  aIndex: Integer;
begin
  aCompIdList := GetCompIdList;

  ProgressBar2.Position := 0;
  with AQry_Inv do
  begin
    //組合sql語法，申報月份
    if dselect = 1 then
    begin
      sL_sql := 'Select Sum('+dField+') Amount From ' +dTable+' ';
      sl_parma := 'where (INVDATE between '+
                  'to_date('+QuotedStr(sL_CreateMonthStart)+',''yyyy/mm/dd'') and '+
                  'to_date('+QuotedStr(sL_CreateMonthEnd)+',''yyyy/mm/dd'')) ' +
                  ' and compid in ' + aCompIdList;
    end else
    if dselect = 2 then
    begin
      sL_sql := 'Select Count(*) Amount From '+dTable+' ';
      sl_parma := 'where (INVDATE between '+
                  'to_date('+QuotedStr(sL_CreateMonthStart)+',''yyyy/mm/dd'') and '+
                  'to_date('+QuotedStr(sL_CreateMonthEnd)+',''yyyy/mm/dd'')) ' +
                  '  and compid in ' + aCompIdList;
    end else
    if dselect = 3 then
    begin
      sL_sql := 'Select (EndNum - CurNum + 1) Amount From '+dTable+' ';
      sl_parma := 'where Yearmonth between '+
                  QuotedStr(copy(sL_CreateMonthStart,1,4)+
                            copy(sL_CreateMonthStart,6,2))+
                  ' and '+
                  QuotedStr(copy(sL_CreateMonthEnd,1,4)+
                            copy(sL_CreateMonthEnd,6,2) ) +
                  ' and compid in ' + aCompIdList;          
    end else
    if dselect = 4 then
    begin
      aYearMonth := Copy( sL_CreateMonthStart, 1, 4 ) + Copy( sL_CreateMonthStart, 6, 2 );
      sL_sql := ' Select Sum('+dField+') Amount From '+dTable+' ';
      sl_parma := Format(
        ' where identifyid1 = ''%s''   ' +
        '   and identifyid2 = ''%s''   ' +
        '   and yearMonth = ''%s''     ' +
        '   and compid in %s           ',
        [IDENTIFYID1, IDENTIFYID2, aYearMonth, aCompIdList] );
    end else
    if dselect = 5 then
    begin
      aYearMonth := Copy( sL_CreateMonthStart, 1, 4 ) + Copy( sL_CreateMonthStart, 6, 2 );
      sL_sql := ' Select Sum('+dField+') Amount From '+dTable+' ';
      sl_parma := Format(
        ' where identifyid1 = ''%s''   ' +
        '   and identifyid2 = ''%s''   ' +
        '   and yearMonth = ''%s''     ' +
        '   and compid in %s           ',
        [IDENTIFYID1, IDENTIFYID2, aYearMonth, aCompIdList] );
    end;

    //不產生申報資料之發票號碼
    if MKE_Start.Tag = 1 then
      sl_parma := sl_parma + ' and not (INVID between '+QuotedStr(MKE_Start.Text)+
                             ' and '+QuotedStr(MKE_End.Text) + ') ';
    //不為作廢的資料
    if dParma1 <> '' then
      sl_parma := sl_parma + 'and (IsObsolete = '+QuotedStr(dParma1) +') ';
    //發票格式
    if dParma2 <> '' then
      if dParma3 <> '' then
        sl_parma := sl_parma + 'and (InvFormat IN ('+QuotedStr(dParma2) +','+
                    QuotedStr(dParma3)+')) '
      else
        sl_parma := sl_parma + 'and (InvFormat IN ('+QuotedStr(dParma2) +')) ';
    //客戶統編
    if dParma4 = '1' then
        sl_parma := sl_parma + 'and (BusinessId is null) '
    else
      if dParma4 = '2' then
        sl_parma := sl_parma + 'and (BusinessId is not null) ';
    //發票稅別
    if  dParma5 <> '' then
        sl_parma := sl_parma + 'and (TaxType= '+QuotedStr(dParma5)+') ';

    sql.Clear;
    sql.Add(sl_sql);
    sql.Add(sl_parma);
    Open;
    //將狀態列的最大值設為文字檔的行數
    ProgressBar2.Properties.Max := recordCount ;
    if recordcount = 0 then
      Result := 0
    else
    begin
      if dselect = 3 then
      begin
        dl_sum := 0;
        while not eof do
        begin
          dL_sum := dl_sum + fieldByName('Amount').AsFloat;
          ProgressBar2.Position := ( ProgressBar2.Position + 1 ); 
          Next;
        end;
        Result := dL_sum;
      end
      else begin
        Result := fieldByName('Amount').AsFloat;
        ProgressBar2.Position := ( ProgressBar2.Position + 1 );
      end;
    end;
    close;
  end;
end;

Procedure TfrmInvA08.SetMasterData(Var P_RecNo:String;P_Inv01BusID,P_Inv001TaxID,
                  P_Table:String;P_Select:Integer;Var P_MediaList:TStringList);
Var
  sL_Inv007Format       : String    ;//格式代碼
  sL_TmpList            : String    ;//
  sL_Inv007InvDate      : String    ;//開立發票日期轉成民國年月
  sL_Inv007BusinessID   : String    ;//發票上的統一編號
  sL_Inv007Amount       : String    ;//銷售額
  sL_Inv007TaxType      : string    ;//課稅別
  sL_Inv007TaxId        : String    ;//發票號碼
  sL_TaxAmount          : String    ;//營業稅額
  sL_detain             : String    ;//扣抵代碼
  sL_special            : String    ;//特種營業稅率
  sL_collect            : String    ;//彙加註記
  sL_Other              : String    ;//洋菸酒註記
  sL_Inv099Prefix       : String    ;//
  sL_Inv099EndNum       : String    ;//
  sL_Inv099StartNum     : String    ;//
  sL_Inv099CurNum       : String    ;//
  sL_startnum,sl_endnum : String;
  sL_InString           : String;
  iL_Count : Integer;
  sL_smbolstr           : String    ;
  i : integer;
  iL_LowNum :Integer;
  iL_HighNum : Integer;
  iL_CountNum : Integer;
  aCanWriteToFile: Boolean;
  aRecordIndex: Integer;
begin

  sL_smbolstr := 'yyyy/mm/dd';

  sL_InString := GetCompIdList;

  if P_Select = 3 then
  begin

    if ( chkDebugFile.Checked ) then
    begin
      FDebugList.Add( Format( '%s: P_Select = 3 ', [FormatDateTime( 'YYYY-MM-DD HH:MM:SS', Now )] ) );
    end;

    AQry_Inv.Sql.Clear;
    AQry_Inv.Sql.Add('select Count(*) Count from inv099 ');
    AQry_Inv.Sql.Add(' where IDENTIFYID1=%S ');
    AQry_Inv.Sql.Add(' and IDENTIFYID2=%S ');
    AQry_Inv.Sql.Add(' and COMPID in %S ');
    AQry_Inv.Sql.Add(' and YEARMONTH = %S ');

    AQry_Inv.Sql.Text := Format( AQry_Inv.Sql.Text,
       [QuotedStr(IDENTIFYID1), IDENTIFYID2, sL_InString,
        QuotedStr( sL_MakeMonth )] );
    AQry_Inv.Open;
    iL_Count := AQry_Inv.FieldByName('count').AsInteger;
    if iL_Count > 0 then
    begin
      AQry_Inv.close;
      AQry_Inv.Sql[0] := ' SELECT INVFORMAT,ENDNUM,PREFIX,CURNUM FROM INV099 ';
      AQry_Inv.Open;
      ProgressBar2.Properties.Max := AQry_Inv.RecordCount;
      while not AQry_Inv.Eof do
      begin
        { 當該發票本 CURNUM > ENDNUM 時, 則不需INSERT }
        if ( AQry_Inv.FieldByName('CurNum').AsInteger <=
             AQry_Inv.FieldByName('EndNum').AsInteger ) then
        begin
          //1取格式代碼
          if AQry_Inv.FieldByName('InvFormat').AsString = '2' then
            sL_Inv007Format := '32'
          else
            sL_Inv007Format := '31';
          //2稅籍資料 ==> P_Inv001TaxID
          //3取流水序號
          P_RecNo := RightStr( IntToStr( 10000000 + StrToInt( P_RecNo ) ), 7 );
          //4取發票民國年月
          sL_Inv007InvDate := IntToStr(StrToInt( Copy( sL_MakeMonth, 1, 4 ) ) -
            1911 ) +Copy( sL_MakeMonth, 5, 2 );
          //5買受人統一編號變更為發票迄號
          sL_Inv007BusinessID := AQry_Inv.FieldByName('EndNum').AsString;
          //6銷售人統一編號 ==> P_Inv01BusID
          //7、8發票字軌
          sL_Inv007TaxId := ( AQry_Inv.FieldByName( 'Prefix' ).AsString +
            AQry_Inv.FieldByName('CurNum').AsString );
          //9銷售額補0
          sL_Inv007Amount := '000000000000';
          //10課稅別
          sL_Inv007TaxType := 'D';
          //11營業稅額補0
          sL_TaxAmount := '0000000000';
          //12扣抵代碼
          sL_detain := ' ';
          //13特種營業稅率
          sL_special:= '      ';
          //14彙加註記填入 A
          sL_collect:= 'A';
          //15洋菸酒註記
          sL_Other  := ' ';

          sL_TmpList := sL_Inv007Format+
                        P_Inv001TaxID+
                        P_RecNo+
                        sL_Inv007InvDate+
                        sL_Inv007BusinessID+
                        P_Inv01BusID+
                        sL_Inv007TaxId+
                        sL_Inv007Amount+
                        sL_Inv007TaxType+
                        sL_TaxAmount+
                        sL_detain+
                        sL_special+
                        sL_collect+
                        sL_Other;
          P_MediaList.Add(sL_TmpList);
        end;
        ProgressBar2.Position := ( ProgressBar2.Position + 1 );
        frmInvA08.Refresh;
        P_RecNo := IntToStr(StrToInt( P_RecNo ) + 1 );
        AQry_Inv.Next;
      end;
    end;
    AQry_Inv.Close;
  end
  else
  if (P_Select = 4) and (dtmMain.GetLinkToMIS) then//印花稅
  begin

    if ( chkDebugFile.Checked ) then
    begin
      FDebugList.Add( Format( '%s: P_Select = 4 ', [FormatDateTime( 'YYYY-MM-DD HH:MM:SS', Now )] ) );
    end;

    with dtmSO.adoA08 do
    begin
      Close;
      SQL.Clear;
      SQL.Add( Format( ' SELECT COUNT(*) COUNT FROM %s.SO034 T, %s.CD019 C ', [dtmMain.getMisDbOwner, dtmMain.getMisDbOwner] ) );
      SQL.Add('  WHERE T.CITEMCODE = C.CODENO ');
      SQL.Add('   AND  C.TAXCODE = %s ' );
      SQL.Add('   AND  COMPCODE= %s ' );
      SQL.add('   AND ( REALDATE BETWEEN %s AND %s ) ');

      SQL.Text := Format(sql.Text,[QuotedStr('4'),
                                   QuotedStr(dtmMain.getMisDbCompCode),
                                   'to_date('+QuotedStr(sL_CreateMonthStart)+','+QuotedStr(sL_smbolstr)+')',
                                   'to_date('+QuotedStr(sL_CreateMonthEnd)+','+QuotedStr(sL_smbolstr)+')']);

      Open;
      iL_Count := FieldByName( 'COUNT' ).AsInteger;
      if iL_Count > 0 then
      begin
        Close;
        SQL.Strings[0] := Format( 'SELECT T.REALDATE, T.BILLNO, T.REALAMT FROM %s.SO034 T, %s.CD019 C ', [dtmMain.getMisDbOwner, dtmMain.getMisDbOwner] );
        SQL.Add( ' ORDER BY T.BILLNO ' );

        Open;
        ProgressBar2.Properties.Max := iL_Count ;//將狀態列的最大值設為文字檔的行數
        while not Eof do
        begin
          //1取格式代碼
          sL_Inv007Format := '36';
          //2稅籍資料 ==> P_Inv001TaxID
          //3取流水序號
          P_RecNo := rightstr(inttostr(10000000+strtoint(P_RecNo)),7);
          //4取發票民國年月
          sL_Inv007InvDate := getYearMonthDay7(fieldByName('realdate').AsDateTime,1);
          //5買受人統一編號變更為發票迄號
          sL_Inv007BusinessID := '        ';
          //6銷售人統一編號 ==> P_Inv01BusID
          //7、8發票字軌
          sL_Inv007TaxId   := '0'+copy(FieldByName('billno').AsString,5,2)+
                              copy(FieldByName('billno').AsString,9,7);
          //9銷售額補0
          sL_Inv007Amount := rightstr(inttostr(1000000000000+fieldByName('realamt').asInteger),12);
          //10課稅別
          sL_Inv007TaxType := '3';
          //11營業稅額補0
          sL_TaxAmount := '0000000000';
          //12扣抵代碼
          sL_detain := ' ';
          //13特種營業稅率
          sL_special:= '      ';
          //14彙加註記填入 A
          sL_collect:= ' ';
          //15洋菸酒註記  若為零稅率, 則填入通關註記
          sL_Other  := ' ';
          if ( sL_Inv007TaxType = '3' ) then
            sL_Other := Nvl( FImportNote, #32 );

          sL_TmpList := sL_Inv007Format+
                        P_Inv001TaxID+
                        P_RecNo+
                        sL_Inv007InvDate+
                        sL_Inv007BusinessID+
                        P_Inv01BusID+
                        sL_Inv007TaxId+
                        sL_Inv007Amount+
                        sL_Inv007TaxType+
                        sL_TaxAmount+
                        sL_detain+
                        sL_special+
                        sL_collect+
                        sL_Other;
          P_MediaList.Add(sL_TmpList);
          ProgressBar2.Position := ( ProgressBar2.Position + 1 );
          frmInvA08.Refresh;
          P_RecNo := inttostr(strtoint(P_RecNo) +1);
          next;
        end;
      end;
      Close;
    end;
  end
  else if P_Select = 5 then // 印花稅總繳
  begin

    if ( chkDebugFile.Checked ) then
    begin
      FDebugList.Add( Format( '%s: P_Select = 5 ', [FormatDateTime( 'YYYY-MM-DD HH:MM:SS', Now )] ) );
    end;

          //1取格式代碼
          sL_Inv007Format := '36';
          //2稅籍資料 ==> P_Inv001TaxID
          //3取流水序號
          P_RecNo := rightstr(inttostr(10000000+strtoint(P_RecNo)),7);
          //4取發票民國年月
          sL_Inv007InvDate := getYearMonthDay7(StrToDate( FormatMaskText( '####/##/##;0;_', sL_MakeMonth + '01' )),1);
          //5買受人統一編號變更為發票迄號
          sL_Inv007BusinessID := Lpad( EmptyStr, 8, #32 );
          //6銷售人統一編號 ==> P_Inv01BusID
          //7、8發票字軌
          sL_Inv007TaxId   := Lpad( EmptyStr, 10, #32 );
          //9銷售額補0
          sL_Inv007Amount := Lpad( sL_StampTaxt, 12, '0' );
          //10課稅別
          sL_Inv007TaxType := '3';
          //11營業稅額補0
          sL_TaxAmount := Lpad( EmptyStr, 10, '0' );
          //12扣抵代碼
          sL_detain := ' ';
          //13特種營業稅率
          sL_special:= '      ';
          //14彙加註記填入 A
          sL_collect:= ' ';
          //15洋菸酒註記  若為零稅率, 則填入通關註記
          sL_Other  := ' ';
          if ( sL_Inv007TaxType = '3' ) then
            sL_Other := Nvl( FImportNote, #32 );

          sL_TmpList := sL_Inv007Format+
                        P_Inv001TaxID+
                        P_RecNo+
                        sL_Inv007InvDate+
                        sL_Inv007BusinessID+
                        P_Inv01BusID+
                        sL_Inv007TaxId+
                        sL_Inv007Amount+
                        sL_Inv007TaxType+
                        sL_TaxAmount+
                        sL_detain+
                        sL_special+
                        sL_collect+
                        sL_Other;
          P_MediaList.Add(sL_TmpList);
          ProgressBar2.Position := ( ProgressBar2.Position + 1 );
          frmInvA08.Refresh;
          P_RecNo := inttostr(strtoint(P_RecNo) +1);
  end
  else if P_Select = 1 then//一般發票
  begin

    if ( chkDebugFile.Checked ) then
    begin
      FDebugList.Add( Format( '%s: P_Select = 1 ', [FormatDateTime( 'YYYY-MM-DD HH:MM:SS', Now )] ) );
    end;

    FDebugList.Clear;

    AQry_Inv099.SQL.Clear;
    AQry_Inv099.SQL.Add( ' SELECT Prefix, EndNum, CurNum, StartNum FROM INV099 ' );
    AQry_Inv099.SQL.Add( ' WHERE IDENTIFYID1=%S ' );
    AQry_Inv099.SQL.Add( '   AND IDENTIFYID2=%S ' );
    AQry_Inv099.SQL.Add( '   AND YEARMONTH = %S ' );
    AQry_Inv099.SQL.Add( '   AND COMPID in %S ' );
    AQry_Inv099.SQL.Text := Format( AQry_Inv099.SQL.text, [
      QuotedStr( IDENTIFYID1 ), IDENTIFYID2, QuotedStr( sL_MakeMonth ),
      sL_InString] );
    AQry_Inv099.SQL.Add( ' Order By Prefix, StartNum ' );
    AQry_Inv099.Open;
    AQry_Inv099.First;

    if ( chkDebugFile.Checked ) then
    begin
      FDebugList.Add( Format( '%s: 啟始偵錯 .....', [FormatDateTime( 'YYYY-MM-DD HH:MM:SS', Now )] ) );
      FDebugList.Add( Format( '%s: 媒體申報發票本的SQL', [FormatDateTime( 'YYYY-MM-DD HH:MM:SS', Now )] ) );
      FDebugList.Add( '--------------------------------------------------------' );
      FDebugList.Add( '   ' + AQry_Inv099.SQL.Text );
      FDebugList.Add( '--------------------------------------------------------' );
    end;

    while not AQry_Inv099.eof do
    begin
      sL_Inv099Prefix := AQry_Inv099.FieldByName('Prefix').AsString;
      sL_Inv099EndNum := AQry_Inv099.FieldByName('EndNum').AsString;
      sL_Inv099CurNum := AQry_Inv099.FieldByName('CurNum').AsString;
      sL_Inv099startNum := AQry_Inv099.FieldByName('StartNum').AsString;
      sl_startnum := sL_Inv099Prefix+adjString(8,sL_Inv099startNum,True,True);
      sl_endnum := sL_Inv099Prefix+ adjString(8,sL_Inv099EndNum,True,True);

      if ( chkDebugFile.Checked ) then
      begin
        FDebugList.Add( Format( '%s: 發票本資訊, 字首:%s, 啟始號碼:%s, 結束號碼:%s, 下一使用號碼:%s,  重組啟始號碼:%s, 重組結束號碼:%s ',
          [FormatDateTime( 'YYYY-MM-DD HH:MM:SS', Now ), sL_Inv099Prefix, sL_Inv099startNum, sL_Inv099EndNum, sL_Inv099CurNum, sl_startnum, sl_endnum] ) );
      end;

      // 先算出總筆數, 分批取
      AQry_Inv.SQL.Text := Format(
        ' SELECT COUNT(1) COUNT                        ' +
        '   FROM %s                                    ' +
        ' WHERE IDENTIFYID1 = ''%s''                   ' +
        '   AND IDENTIFYID2 = %s                       ' +
        '   AND INVDATE BETWEEN                        ' +
        '         TO_DATE( ''%s'', ''YYYY/MM/DD'' )    ' +
        '       AND                                    ' +
        '         TO_DATE( ''%s'', ''YYYY/MM/DD'' )    ' +
        '   AND COMPID in %s                           ' +
        '   AND INVID BETWEEN ''%s'' AND ''%s''        ',
        [P_Table, IDENTIFYID1, IDENTIFYID2,
         sL_CreateMonthStart, sL_CreateMonthEnd, sL_InString,
         sl_startnum, sl_endnum] );

      if MKE_Start.tag = 1 then
      begin
        AQry_Inv.SQL.Text := AQry_Inv.SQL.Text +
          ' AND NOT ( INVID BETWEEN ' + QuotedStr( MKE_Start.Text )+
                ' AND ' + QuotedStr( MKE_End.Text ) + ' ) ';
      end;

      AQry_Inv.Open;
      iL_Count := AQry_Inv.FieldByName( 'Count' ).AsInteger;
      AQry_Inv.Close;

      if ( chkDebugFile.Checked ) then
      begin
        FDebugList.Add( Format( '%s: 依發票本分算出出發票資料筆數的SQL:', [FormatDateTime( 'YYYY-MM-DD HH:MM:SS', Now )] ) );
        FDebugList.Add( AQry_Inv.SQL.Text );
        FDebugList.Add( '--------------------------------------------------------' );
      end;

      // 分批抓取資料時, 每次最大抓取筆數
      iL_CountNum := 5000;

      if iL_Count > 0 then
      begin
        ProgressBar2.Position := 0;
        ProgressBar2.Properties.Max := il_count ;//將狀態列的最大值設為文字檔的行數
        iL_LowNum := 1;
        iL_HighNum := iL_CountNum;
        if il_count > iL_CountNum then
        begin
          sl_startnum := sL_Inv099Prefix + adjString( 8, sL_Inv099startNum,
            True, True );
          sl_endnum := sL_Inv099Prefix + adjString( 8, IntToStr(
            StrToInt( Rightstr( sL_Inv099startNum, 8 ) ) + iL_CountNum - 1 ),
            True, True );
          {}
          if ( chkDebugFile.Checked ) then
          begin
            FDebugList.Add( '       ********************************************************' );
            FDebugList.Add( '       ' + Format( '%s: 分批抓取資料, 重新計算的發票號碼起迄: %s - %s',
              [FormatDateTime( 'YYYY-MM-DD HH:MM:SS', Now ), sl_startnum, sl_endnum] ) );
            FDebugList.Add( '       ********************************************************' );
          end;
          {}
        end;

        aRecordIndex := 1;
        while ( aRecordIndex <= il_count ) do
        begin

          if ( chkDebugFile.Checked ) then
          begin
            FDebugList.Add( Format( 'aRecordIndex:%d, iL_LowNum:%d, AQry_Inv.Active:%s',
              [aRecordIndex, iL_LowNum, BoolToStr( AQry_Inv.Active )] ) );
            FDebugList.Add( '********************************************************' );
          end;

              { 當處理的筆數已到達設定的筆數, 或是該次 SQL 並無取得資料,
                則再次處理 SQL, 並Open DataSet }
              if ( aRecordIndex = iL_LowNum ) or ( not AQry_Inv.Active ) then
              begin
                 AQry_Inv.SQL.Text := Format(
                   ' SELECT INVFORMAT,             ' +
                   '        INVDATE,               ' +
                   '        ISOBSOLETE,            ' +
                   '        TAXTYPE,               ' +
                   '        BUSINESSID,            ' +
                   '        INVAMOUNT,             ' +
                   '        SALEAMOUNT,            ' +
                   '        TAXAMOUNT,             ' +
                   '        INVID,                 ' +
                   '        MAININVID,             ' +
                   '        MAINSALEAMOUNT,        ' +
                   '        MAINTAXAMOUNT,         ' +
                   '        MAININVAMOUNT          ' +
                   '  FROM %s                      ' +
                   ' WHERE IDENTIFYID1 = ''%s''    ' +
                   '   AND IDENTIFYID2 = %s        ' +
                   '   AND INVDATE BETWEEN         ' +
                   '           TO_DATE( ''%s'', ''YYYY/MM/DD'' )  ' +
                   '       AND                     ' +
                   '           TO_DATE( ''%s'', ''YYYY/MM/DD'' )  ' +
                   '   AND COMPID in %s            ' +
                   '   AND INVID BETWEEN ''%s'' AND ''%s''        ',
                   [P_Table, IDENTIFYID1, IDENTIFYID2, sL_CreateMonthStart,
                     sL_CreateMonthEnd, sL_InString, sl_startnum,
                     sl_endnum] );

                 if MKE_Start.tag = 1 then
                 begin
                   AQry_Inv.SQL.Text := AQry_Inv.SQL.Text +
                   Format( ' AND NOT ( INVID BETWEEN ''%s'' AND ''%s''  ) ',
                           [MKE_Start.Text, MKE_End.Text] );
                 end;
                AQry_Inv.Close;
                AQry_Inv.SQL.Text := ( AQry_Inv.SQL.Text + ' ORDER BY INVID ' );
                AQry_Inv.Open;

                if ( chkDebugFile.Checked ) then
                begin
                  FDebugList.Add( Format( '%s: 依算出的筆數取出發票資料的SQL:', [FormatDateTime( 'YYYY-MM-DD HH:MM:SS', Now )] ) );
                  FDebugList.Add( AQry_Inv.SQL.Text );
                  FDebugList.Add( '--------------------------------------------------------' );
                end;

                iL_LowNum := iL_HighNum + 1;
                iL_HighNum := iL_HighNum + iL_CountNum;
                sl_startnum := sL_Inv099Prefix + adjString( 8,
                  IntToStr(StrToInt( RightStr( sl_endnum,8 ) ) + 1 ),True, True );
                sl_endnum := sL_Inv099Prefix+ adjString( 8, IntToStr( StrToInt(
                  RightStr( sl_endnum, 8 ) ) + iL_CountNum ),True, True );

                { 沒取到資料, 關掉 DataSet }
                if ( AQry_Inv.IsEmpty ) then
                begin
                  AQry_Inv.Close;
                  Continue;
                end;

              end;

              aCanWriteToFile := True;

              { 以主發票號碼申報 }
              if FApplyByMainInv then
              begin
                { 是否是主發票號碼 }
                if ( AQry_Inv.FieldByName( 'INVID' ).AsString <>
                     Nvl( AQry_Inv.FieldByName( 'MAININVID' ).AsString,
                          AQry_Inv.FieldByName( 'INVID' ).AsString ) ) then
                aCanWriteToFile := False;
              end;



              if ( aCanWriteToFile ) then
              begin


                //1取格式代碼
                  if ( chkPaper.Checked ) then
                  begin
                    sL_Inv007Format := '36';
                  end else
                  begin
                    if AQry_Inv.FieldByName('InvFormat').AsString = '2' then
                      sL_Inv007Format := '32'
                    else
                      sL_Inv007Format := '31';
                  end;
                  //4取發票民國年月
                  sL_Inv007InvDate := getYearMonthDay7( AQry_Inv.fieldByName('INVDATE').AsDateTime,1 );
                  //判斷發票稅別
                  if AQry_Inv.FieldByName('ISOBSOLETE').AsString = 'Y' then
                  begin
                    //發票為作廢時，變更下列資料
                    //5買受人統一編號變更為發票迄號
                    sL_Inv007BusinessID := '00000000';
                    //9銷售額補0
                    sL_Inv007Amount := '000000000000';
                    //10課稅別改為D
                    sL_Inv007TaxType := 'D';
                    //11營業稅額補0
                    sL_TaxAmount := '0000000000';
                    //14彙加註記填入 A
                    sL_collect:= ' ';
                  end else
                  begin
                    //10課稅別
                    sL_Inv007TaxType := inttostr(AQry_Inv.fieldByNAme('TaxType').AsInteger);
                    //14彙加註記填入空白
                    sL_collect:= ' ';
                    //判斷發票統編是否為空白，是則補8個0；另可判斷銷售額、營業稅額
                    if AQry_Inv.FieldByName('businessID').AsString = EmptyStr then
                    begin
                      //5買受人統一編號
                      sL_Inv007BusinessID := '00000000';
                      //9銷售額
                      if ( FApplyByMainInv ) then
                        sL_Inv007Amount := RightStr( IntToStr( 1000000000000 + AQry_Inv.FieldByName('MainInvAmount').AsInteger ), 12 )
                      else
                        sL_Inv007Amount := RightStr( IntToStr( 1000000000000 + AQry_Inv.FieldByName('InvAmount').AsInteger ), 12 );
                      //11營業稅額
                      sL_TaxAmount := '0000000000';
                    end else
                    begin
                      //5買受人統一編號
                      sL_Inv007BusinessID := AQry_Inv.FieldByName('businessID').AsString;
                      if ( FApplyByMainInv ) then
                      begin
                        //9銷售額
                        sL_Inv007Amount := RightStr( IntToStr( 1000000000000 + AQry_Inv.FieldByName('MainSaleAmount').AsInteger), 12 );
                        //11營業稅額
                        sL_TaxAmount := RightStr( IntToStr( 10000000000 + AQry_Inv.FieldByName('MainTaxAmount').AsInteger ), 10 );
                      end else
                      begin
                        //9銷售額
                        sL_Inv007Amount := RightStr( IntToStr( 1000000000000 + AQry_Inv.FieldByName('SaleAmount').AsInteger), 12 );
                        //11營業稅額
                        sL_TaxAmount := RightStr( IntToStr( 10000000000 + AQry_Inv.FieldByName('TaxAmount').AsInteger ), 10 );
                      end;
                    end;
                  end;

                //2稅籍資料 ==> P_Inv001TaxID
                //3取流水序號
                P_RecNo := Rightstr( IntToStr( 10000000 + StrToInt( P_RecNo ) ), 7 );
                //6銷售人統一編號 ==> P_Inv01BusID
                //7、8發票字軌
                sL_Inv007TaxId := AQry_Inv.FieldByName('InvID').AsString;

                //12扣抵代碼
                sL_detain := ' ';
                //13特種營業稅率
                sL_special:= '      ';
                //15洋菸酒註記  若為零稅率, 則填入通關方式
                sL_Other  := ' ';
                if ( sL_Inv007TaxType = '3' ) then
                  sL_Other := Nvl( FImportNote, #32 );

                sL_TmpList := sL_Inv007Format+
                              P_Inv001TaxID+
                              P_RecNo+
                              sL_Inv007InvDate+
                              sL_Inv007BusinessID+
                              P_Inv01BusID+
                              sL_Inv007TaxId+
                              sL_Inv007Amount+
                              sL_Inv007TaxType+
                              sL_TaxAmount+
                              sL_detain+
                              sL_special+
                              sL_collect+
                              sL_Other;

                // 以收據申報的話, 做廢資料不須產生
                if ( ( AQry_Inv.FieldByName('ISOBSOLETE').AsString = 'Y' ) and
                     ( chkPaper.Checked ) ) then
                begin
                   // do nothing
                end else
                begin
                  P_MediaList.Add(sL_TmpList);
                  P_RecNo := inttostr(strtoint(P_RecNo) +1);
                  if ( chkDebugFile.Checked ) then
                  begin
                    FDebugList.Add( Format( '%s: P_RecNo:%s', [FormatDateTime( 'YYYY-MM-DD HH:MM:SS', Now ), P_RecNo] ) );
                  end;
                end;

              end;
              ProgressBar2.Position := ( ProgressBar2.Position + 1 );
              frmInvA08.Refresh;
              AQry_Inv.Next;
              Inc( aRecordIndex );

              if ( AQry_Inv.Eof ) then
              begin
                AQry_Inv.Close;
                if ( chkDebugFile.Checked ) then
                begin
                  FDebugList.Add( Format( '%s: AQry_Inv.Closed()', [FormatDateTime( 'YYYY-MM-DD HH:MM:SS', Now )] ) );
                end;
              end;

        end;
      end;

      AQry_Inv099.next;


    end;

  end
  else if P_Select = 2 then
  begin

    if ( chkDebugFile.Checked ) then
    begin
      FDebugList.Add( Format( '%s: P_Select = 2 ', [FormatDateTime( 'YYYY-MM-DD HH:MM:SS', Now )] ) );
    end;

    AQry_Inv.Close;
    AQry_Inv.SQL.Text := Format(
      ' SELECT A.* FROM %s A                 ' +
      '  WHERE A.IDENTIFYID1 = ''%s''        ' +
      '    AND A.IDENTIFYID2 = %s            ' +
      '    AND A.YEARMONTH = ''%s''          ' +
      '    AND A.COMPID in %s                ',
      [P_Table, IDENTIFYID1, IDENTIFYID2,
       Copy( sL_CreateMonthStart, 1, 4 ) + Copy( sL_CreateMonthStart, 6, 2 ),
       sL_InString] );

    if MKE_Start.tag = 1 then
    begin
      AQry_Inv.SQL.Text := AQry_Inv.SQL.Text +
      Format( ' AND NOT ( A.INVID BETWEEN ''%s'' AND ''%s''  ) ',
              [MKE_Start.Text, MKE_End.Text] );
    end;
    AQry_Inv.Open;
    ProgressBar2.Properties.Max := AQry_Inv.RecordCount ;//將狀態列的最大值設為文字檔的行數
    AQry_Inv.First;
    while not AQry_Inv.Eof do
    begin
      //1取格式代碼
      if ( chkPaper.Checked ) then
      begin
        sL_Inv007Format := '34';
      end else
      begin
        if AQry_Inv.FieldByName('INVFORMAT').AsInteger = 2 then
          sL_Inv007Format := '34'
        else
          sL_Inv007Format := '33';
      end;    
      //2稅籍資料 ==> P_Inv001TaxID
      //3取流水序號
      P_RecNo := rightstr(inttostr(10000000+strtoint(P_RecNo)),7);
      //4取發票民國年月
      sL_Inv007InvDate := getYearMonthDay7( AQry_Inv.FieldByName('PAPERDATE').AsDateTime,1);
      //判斷發票統編是否為空白，是則補8個0；另可判斷銷售額、營業稅額
      if AQry_Inv.FieldByName('BUSINESSID').AsString = '' then
        //5買受人統一編號
        sL_Inv007BusinessID := '00000000'
      else
        //5買受人統一編號
        sL_Inv007BusinessID := AQry_Inv.FieldByName('BUSINESSID').AsString;
      //6銷售人統一編號 ==> P_Inv01BusID
      //7、8發票字軌
      sL_Inv007TaxId  := AQry_Inv.FieldByName('INVID').AsString;
      //9銷售額
      sL_Inv007Amount := rightstr(inttostr(1000000000000+ AQry_Inv.FieldByName('SALEAMOUNT').asInteger),12);
      //10課稅別
      sL_Inv007TaxType := inttostr(AQry_Inv.FieldByNAme('TAXTYPE').AsInteger);
      //11營業稅額
      sL_TaxAmount := rightstr(inttostr(10000000000+AQry_Inv.FieldByName('TAXAMOUNT').asInteger),10);
      //12扣抵代碼
      sL_detain := ' ';
      //13特種營業稅率
      sL_special:= '      ';
      //14彙加註記
      sL_collect:= ' ';
      //15洋菸酒註記 若為零稅率,則填入通關註記
      sL_Other  := ' ';
      if ( sL_Inv007TaxType = '3' ) then
        sL_Other := Nvl( FImportNote, #32 );

      sL_TmpList := sL_Inv007Format+
                    P_Inv001TaxID+
                    P_RecNo+
                    sL_Inv007InvDate+
                    sL_Inv007BusinessID+
                    P_Inv01BusID+
                    sL_Inv007TaxId+
                    sL_Inv007Amount+
                    sL_Inv007TaxType+
                    sL_TaxAmount+
                    sL_detain+
                    sL_special+
                    sL_collect+
                    sL_Other;
      P_MediaList.Add(sL_TmpList);
      ProgressBar2.Position := ( ProgressBar2.Position + 1 );
      frmInvA08.Refresh;
      P_RecNo := inttostr(strtoint(P_RecNo) +1);
      AQry_Inv.Next;
    end;
    AQry_Inv.Close;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA08.Bit_PreviousClick(Sender: TObject);
begin
  ListBox1.Clear;
  PCT1.ActivePageIndex := 0;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA08.MKE_StartExit(Sender: TObject);
begin
  if trim(MKE_End.Text) = '' then
    MKE_End.Text := MKE_Start.Text;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA08.Btn_PrintClick(Sender: TObject);
var
  aIndex: Integer;
begin
  if PrinterSetupDialog1.Execute then
  begin
    Printer.BeginDoc;
    try
      Printer.Canvas.Font.Size := 10;
      Printer.Canvas.Font.Name := '細明體';
      for  aIndex := 7 to ListBox1.Items.Count - 1 do
        printer.Canvas.TextOut( 200, ( 200 + aIndex *90 ),listbox1.Items[aIndex] );
    finally
      Printer.EndDoc;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA08.chkStampTaxClick(Sender: TObject);
begin
  txtStampTax.Enabled := chkStampTax.Checked;
  if not txtStampTax.Enabled then txtStampTax.Value := 0;
  if chkStampTax.Checked then CheckBox1.Checked := False;
end;

{ ---------------------------------------------------------------------------- }


procedure TfrmInvA08.CheckBox1Click(Sender: TObject);
begin
  if CheckBox1.Checked then
  begin
    chkStampTax.Checked := False;
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
