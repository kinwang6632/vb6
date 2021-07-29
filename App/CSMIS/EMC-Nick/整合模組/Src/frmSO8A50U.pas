unit frmSO8A50U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, fraChineseYMDU, StdCtrls, Buttons, DB, ExtCtrls, fraChineseYMU,
  Grids, DBGrids, scExcelExport;


type
  TfrmSO8A50 = class(TForm)
    Label1: TLabel;
    cobComp: TComboBox;
    btnStartCalculate: TBitBtn;
    btnCancel: TBitBtn;
    btnReset: TBitBtn;
    lblRealDate: TLabel;
    fraChineseYMD1: TfraChineseYMD;
    fraChineseYMD2: TfraChineseYMD;
    lblShouldDate: TLabel;
    fraChineseYMD3: TfraChineseYMD;
    fraChineseYMD4: TfraChineseYMD;
    dsrCodeCD039A: TDataSource;
    rgpQueryType: TRadioGroup;
    Label2: TLabel;
    lbxChargeItem: TListBox;
    btnChargeItem: TBitBtn;
    fraChineseYM1: TfraChineseYM;
    fraChineseYM2: TfraChineseYM;
    procedure FormCreate(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure rgpQueryTypeClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnChargeItemClick(Sender: TObject);
    procedure btnStartCalculateClick(Sender: TObject);
    procedure btnResetClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
    function IsDataOk : Boolean;
    procedure initialRealAndShouldDate;
  public
    { Public declarations }
    G_TargetFieldValueStrList1,G_TempChargeItem,G_ChargeCodeStrList : TStringList;
    nG_RealOrShouldDate : Integer;
    sG_ChargeItemSQL,sG_ChargeItemNameSQL : String;
    procedure showDetailExcel(sI_CompName,sI_ComputeYM,sI_ChargeItemNameSQL : String;nI_RealOrShouldDate : Integer);
  end;

var
  frmSO8A50: TfrmSO8A50;

implementation

uses UCommonU, UdateTimeu, frmDbMultiSelectu, Ustru, XLSFile,
  dtmMain2U, frmMainMenuU;

{$R *.dfm}

procedure TfrmSO8A50.FormCreate(Sender: TObject);
begin

    G_TargetFieldValueStrList1 := TStringList.Create;
    G_TempChargeItem := TStringList.Create;
    G_ChargeCodeStrList := TStringList.Create;


    TUCommonFun.AddObjectToComboBox(cobComp.Items, dsrCodeCD039A.DataSet,NOT INSERT_NO_DATA_ITEM,'CodeNo','Description');
    //指定為VB傳來之User所屬公司
    TUCommonFun.setComboDefaultNdx(cobComp, frmMainMenu.sG_CompCode);

    dtmMain2.cdsCodeCd039.Open;

    //查詢畫面時間初始化
    initialRealAndShouldDate;
end;

procedure TfrmSO8A50.btnCancelClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmSO8A50.rgpQueryTypeClick(Sender: TObject);
begin
    //切換實收金額或應收金額
    if rgpQueryType.ItemIndex = REAL_DATE_TYPE then
    begin
      nG_RealOrShouldDate := REAL_DATE_TYPE;

      //改為年月
      lblRealDate.Visible := true;
      fraChineseYM1.Visible := true;

      lblShouldDate.Visible := false;
      fraChineseYM2.Visible := false;
{
      lblShouldDate.Visible := false;
      fraChineseYMD3.Visible := false;
      fraChineseYMD4.Visible := false;

      lblRealDate.Visible := true;
      fraChineseYMD1.Visible := true;
      fraChineseYMD2.Visible := true;
}
    end
    else if rgpQueryType.ItemIndex = SHOULD_DATE_TYPE then
    begin
      nG_RealOrShouldDate := SHOULD_DATE_TYPE;

      //改為年月
      lblRealDate.Visible := false;
      fraChineseYM1.Visible := false;

      lblShouldDate.Visible := true;
      fraChineseYM2.Visible := true;
{
      lblShouldDate.Visible := true;
      fraChineseYMD3.Visible := true;
      fraChineseYMD4.Visible := true;

      lblRealDate.Visible := false;
      fraChineseYMD1.Visible := false;
      fraChineseYMD2.Visible := false;
}
    end;

    fraChineseYM1.setYM('    /  ');
    fraChineseYM2.setYM('    /  ');


end;

procedure TfrmSO8A50.FormShow(Sender: TObject);
begin
    rgpQueryTypeClick(frmSO8A50);

    self.Caption := frmMainMenu.setCaption('SO8A50','[拆帳明細資料]');    
end;

procedure TfrmSO8A50.btnChargeItemClick(Sender: TObject);
var
    L_TargetFieldNamesStrList : TStringList;
    sL_EmpNoSQL : String;
    jj : Integer;
begin
    dtmMain2.getCodeCD019;

    L_TargetFieldNamesStrList:=TStringList.Create;

    L_TargetFieldNamesStrList.Add('CODENO');
    L_TargetFieldNamesStrList.Add('DESCRIPTION');

    //先清空或給初值
    dtmMain2.cdsCodeCD019.FieldByName('CODENO').DisplayLabel := '收費項目代碼';
    dtmMain2.cdsCodeCD019.FieldByName('DESCRIPTION').DisplayLabel := '收費項目名稱';


    if SelectMultiRecords('請點選收費項目名稱', dtmMain2.cdsCodeCD019, 'CODENO;DESCRIPTION', ',', L_TargetFieldNamesStrList, G_TargetFieldValueStrList1,true,sL_EmpNoSQL) = mrOk then
    begin

      //按確定時先清空畫面
      lbxChargeItem.Clear;
      sG_ChargeItemSQL := '';
      sG_ChargeItemNameSQL := '';

      for jj:=0 to G_TargetFieldValueStrList1.Count -1 do
      begin
        G_TempChargeItem := TUstr.ParseStrings(Trim(G_TargetFieldValueStrList1.Strings[jj]),',');

        if sG_ChargeItemSQL = '' then
        begin
          sG_ChargeItemSQL :=  G_TempChargeItem.Strings[0];
          sG_ChargeItemNameSQL := G_TempChargeItem.Strings[1];
        end
        else
        begin
          sG_ChargeItemSQL := sG_ChargeItemSQL + ',' + G_TempChargeItem.Strings[0];
          sG_ChargeItemNameSQL := sG_ChargeItemNameSQL + ',' + G_TempChargeItem.Strings[1];
        end;


        lbxChargeItem.Items.Add(G_TempChargeItem.Strings[1]);

        //用於篩選退費時沒選的收費項目
        G_ChargeCodeStrList.Add(G_TempChargeItem.Strings[0]);
      end;

    end;

//SHOWMESSAGE(sG_ChargeItemSQL);

end;

function TfrmSO8A50.IsDataOk: Boolean;
var
    sL_StartDate,sL_ReplaceSDate,sL_EndDate,sL_ReplaceEDate : String;
    sL_BelongYM,sL_ReplaceBelongYM : String;
    nL_StartDate,nL_EndDate,nL_BelongYM : Integer;
    dL_StartDate,dL_sL_EndDate : TDate;

    sL_ComputeYM : String;
begin
    if sG_ChargeItemSQL = '' then
    begin
        MessageDlg('請點選收費項目',mtError, [mbOK],0);
        btnChargeItem.SetFocus;
        IsDataOk := false;
        exit;
    end;



    //改為年月
    if nG_RealOrShouldDate = REAL_DATE_TYPE then
    begin
      sL_ComputeYM := Trim(TUstr.replaceStr(fraChineseYM1.getYM,'/',''));
    end else
    if nG_RealOrShouldDate = SHOULD_DATE_TYPE then
    begin
      sL_ComputeYM := Trim(TUstr.replaceStr(fraChineseYM2.getYM,'/',''));
    end;

    if ( sL_ComputeYM = '' ) then
    begin

      if nG_RealOrShouldDate = REAL_DATE_TYPE then
      begin
        MessageDlg('請輸入實收年月',mtError, [mbOK],0);
        fraChineseYM1.mseYM.SetFocus;
      end
      else if nG_RealOrShouldDate = SHOULD_DATE_TYPE then
      begin
        MessageDlg('請輸入應收年月',mtError, [mbOK],0);
        fraChineseYM2.mseYM.SetFocus;
      end;

      Result:=False;
      exit;
    end else
    begin
      sL_ComputeYM := TUstr.ToROCYM( sL_ComputeYM );
      Result := True;
    end;  

{
    if nG_RealOrShouldDate = REAL_DATE_TYPE then
    begin
      //取實收日期
      sL_StartDate := Trim(fraChineseYMD1.getYMD);
      sL_EndDate := Trim(fraChineseYMD2.getYMD);
    end
    else if nG_RealOrShouldDate = SHOULD_DATE_TYPE then
    begin
      //取應收日期
      sL_StartDate := Trim(fraChineseYMD3.getYMD);
      sL_EndDate := Trim(fraChineseYMD4.getYMD);
    end;


    //sL_StartDate := Trim(fraChineseYMD1.getYMD);
    sL_ReplaceSDate := Trim(dtmMain2.ReplaceStr(sL_StartDate,'/'));
    nL_StartDate := Length(sL_ReplaceSDate);

    //SHOWMESSAGE(sL_StartDate);
    //SHOWMESSAGE(sL_ReplaceSDate);
    //SHOWMESSAGE(IntToStr(nL_StartDate));
    //檢查實收日期始日期
    if Trim(sL_ReplaceSDate) <> '' then
    begin
      if nL_StartDate <> 7 then
      begin
        MessageDlg('日期請輸入7位數',mtError, [mbOK],0);

        if nG_RealOrShouldDate = REAL_DATE_TYPE then
        begin
          fraChineseYMD1.setYMD('');
          fraChineseYMD1.mseYMD.SetFocus;
        end
        else if nG_RealOrShouldDate = SHOULD_DATE_TYPE then
        begin
          fraChineseYMD3.setYMD('');
          fraChineseYMD3.mseYMD.SetFocus;
        end;


        IsDataOk := false;
        exit;
      end
      else
      begin
        //檢查日期格式是否錯誤
        //showmessage(sL_StartDate);
        if TUdateTime.IsDateStr(sL_StartDate,'/') = false then
        begin
          MessageDlg('日期格式錯誤',mtError, [mbOK],0);

          if nG_RealOrShouldDate = REAL_DATE_TYPE then
          begin
            fraChineseYMD1.setYMD('');
            fraChineseYMD1.mseYMD.SetFocus;
          end
          else if nG_RealOrShouldDate = SHOULD_DATE_TYPE then
          begin
            fraChineseYMD3.setYMD('');
            fraChineseYMD3.mseYMD.SetFocus;
          end;

          IsDataOk := false;
          exit;
        end
      end
    end
    else
    begin
      MessageDlg('請輸入始日期',mtError, [mbOK],0);

      if nG_RealOrShouldDate = REAL_DATE_TYPE then
      begin
        fraChineseYMD1.setYMD('');
        fraChineseYMD1.mseYMD.SetFocus;
      end
      else if nG_RealOrShouldDate = SHOULD_DATE_TYPE then
      begin
        fraChineseYMD3.setYMD('');
        fraChineseYMD3.mseYMD.SetFocus;
      end;

      IsDataOk := false;
      exit;
    end;



  //檢查實收日期結束日期
    //sL_EndDate := Trim(fraChineseYMD2.getYMD);
    sL_ReplaceEDate := Trim(dtmMain2.ReplaceStr(sL_EndDate,'/'));
    nL_EndDate := Length(sL_ReplaceEDate);
    if Trim(sL_ReplaceEDate) <> '' then
    begin
      if nL_EndDate <> 7 then
      begin
        MessageDlg('日期請輸入7位數',mtError, [mbOK],0);

        if nG_RealOrShouldDate = REAL_DATE_TYPE then
        begin
          fraChineseYMD2.setYMD('');
          fraChineseYMD2.mseYMD.SetFocus;
        end
        else if nG_RealOrShouldDate = SHOULD_DATE_TYPE then
        begin
          fraChineseYMD4.setYMD('');
          fraChineseYMD4.mseYMD.SetFocus;
        end;

        IsDataOk := false;
        exit;
      end
      else
      begin
        //檢查日期格式是否錯誤
        if TUdateTime.IsDateStr(sL_EndDate,'/') = false then
        begin
          MessageDlg('日期格式錯誤',mtError, [mbOK],0);

          if nG_RealOrShouldDate = REAL_DATE_TYPE then
          begin
            fraChineseYMD2.setYMD('');
            fraChineseYMD2.mseYMD.SetFocus;
          end
          else if nG_RealOrShouldDate = SHOULD_DATE_TYPE then
          begin
            fraChineseYMD4.setYMD('');
            fraChineseYMD4.mseYMD.SetFocus;
          end;

          IsDataOk := false;
          exit;
        end
        else
          IsDataOk := true;
      end
    end
    else
    begin
      MessageDlg('請輸入結束日期',mtError, [mbOK],0);

      if nG_RealOrShouldDate = REAL_DATE_TYPE then
      begin
        fraChineseYMD2.setYMD('');
        fraChineseYMD2.mseYMD.SetFocus;
      end
      else if nG_RealOrShouldDate = SHOULD_DATE_TYPE then
      begin
        fraChineseYMD4.setYMD('');
        fraChineseYMD4.mseYMD.SetFocus;
      end;

      IsDataOk := false;
      exit;
    end;


    dL_StartDate := StrToDate(dtmMain2.TransToEngDate(sL_StartDate));
    dL_sL_EndDate := StrToDate(dtmMain2.TransToEngDate(sL_EndDate));
    if dL_StartDate > dL_sL_EndDate then
    begin
      MessageDlg('結束日期必須大於等於始日期',mtError, [mbOK],0);

      if nG_RealOrShouldDate = REAL_DATE_TYPE then
      begin
        fraChineseYMD2.setYMD('');
        fraChineseYMD2.mseYMD.SetFocus;
      end
      else if nG_RealOrShouldDate = SHOULD_DATE_TYPE then
      begin
        fraChineseYMD4.setYMD('');
        fraChineseYMD4.mseYMD.SetFocus;
      end;

      IsDataOk := false;
      exit;
    end;
}
end;

procedure TfrmSO8A50.btnStartCalculateClick(Sender: TObject);
var
    sL_CompCode,sL_ComputeYM,sL_Result : String;
    bL_HaveData : Boolean;
begin
    if IsDataOk then
    begin
      //執行等待狀態
      TUCommonFun.setWaitingCursor;
//showmessage('1');
      if nG_RealOrShouldDate = REAL_DATE_TYPE then
      begin
        //取實收年月
        sL_ComputeYM := Trim(fraChineseYM1.getYM);
        sL_ComputeYM := TUstr.ToROCYM( sL_ComputeYM );
      end
      else if nG_RealOrShouldDate = SHOULD_DATE_TYPE then
      begin
        //取應收年月
        sL_ComputeYM := Trim(fraChineseYM2.getYM);
        sL_ComputeYM := TUstr.ToROCYM( sL_ComputeYM );
      end;
//showmessage('2');


      //檢查該年月實收或應收年月是否已計算過
      bL_HaveData := dtmMain2.checkHaveSO131Data(sL_ComputeYM,nG_RealOrShouldDate);

      if bL_HaveData then
      begin
        //showmessage('統計年月有資料');
        if MessageDlg('已有資料是否要重新計算?',mtConfirmation, [mbYes, mbNo], 0) = mrNo then
        begin
//showmessage('3');
          //showmessage('不重新計算');
          //回復原狀態
          dtmMain2.getDetailData(frmMainMenu.sG_CompCode,sL_ComputeYM,nG_RealOrShouldDate);
//showmessage('4');
          //轉成Excel
          showDetailExcel(frmMainMenu.sG_CompName,sL_ComputeYM,sG_ChargeItemNameSQL,nG_RealOrShouldDate);
//showmessage('5');
          TUCommonFun.setDefaultCursor;
//showmessage('6');
        end
        else
        begin
          //showmessage('重新計算');
//showmessage('7');
          sL_Result := dtmMain2.calculateDetailData(frmMainMenu.sG_CompCode,sG_ChargeItemSQL,sL_ComputeYM,nG_RealOrShouldDate);
//          sL_Result := ''; //jacy testing...
          if sL_Result <> '' then
          begin

            MessageDlg(sL_Result,mtError, [mbOK],0);
            //回復原狀態
//showmessage('8');
            TUCommonFun.setDefaultCursor;
            exit;

          end
          else
          begin
//showmessage('9');
            dtmMain2.getDetailData(frmMainMenu.sG_CompCode,sL_ComputeYM,nG_RealOrShouldDate);

//showmessage('10');
            //轉成Excel
            showDetailExcel(frmMainMenu.sG_CompName,sL_ComputeYM,sG_ChargeItemNameSQL,nG_RealOrShouldDate);
//showmessage('11');
            //回復原狀態
            TUCommonFun.setDefaultCursor;
            //showmessage('計算完成');
//showmessage('12');
          end;
        end;
      end
      else
      begin
        //showmessage('統計年月無資料');
//showmessage('13');
        sL_Result := dtmMain2.calculateDetailData(frmMainMenu.sG_CompCode,sG_ChargeItemSQL,sL_ComputeYM,nG_RealOrShouldDate);
        if sL_Result <> '' then
        begin
          MessageDlg(sL_Result,mtError, [mbOK],0);
          //回復原狀態
//showmessage('14');
          TUCommonFun.setDefaultCursor;
          exit;
        end
        else
        begin
//showmessage('15');
          dtmMain2.getDetailData(frmMainMenu.sG_CompCode,sL_ComputeYM,nG_RealOrShouldDate);
//showmessage('16');
          //轉成Excel
          showDetailExcel(frmMainMenu.sG_CompName,sL_ComputeYM,sG_ChargeItemNameSQL,nG_RealOrShouldDate);
//showmessage('17');
          //回復原狀態
          TUCommonFun.setDefaultCursor;
          //showmessage('計算完成');
        end;
      end;
    end;
end;


procedure TfrmSO8A50.btnResetClick(Sender: TObject);
begin
    //畫面重設
    lbxChargeItem.Clear;
    sG_ChargeItemSQL := '';
    sG_ChargeItemNameSQL := '';
    
    //查詢畫面時間初始化
    initialRealAndShouldDate;
end;

procedure TfrmSO8A50.initialRealAndShouldDate;
var
  wL_Year, wL_Month, wL_Day : Word;
  sL_StartDate,sL_EndDate : String;
  nL_LastDay : Integer;

begin
    //改為年月
    fraChineseYM1.setYM('    /  ');
    fraChineseYM2.setYM('    /  ');



    //開始實收日期 = 預設為上個月的第一天
    //結束實收日期 = 預設為上個月的最後一天
    // 歸 屬 日 期 = 預設為程式執行當天
{
    DecodeDate(Date,wL_Year, wL_Month, wL_Day);
    if wL_Month = 1 then
      sL_StartDate := IntToStr(wL_Year-1) + '/' + IntToStr(12) + '/01'
    else
      sL_StartDate := IntToStr(wL_Year) + '/' + IntToStr(wL_Month - 1 ) + '/01';

    nL_LastDay := TUdateTime.DaysOfMonth(StrToDate(sL_StartDate));
    if wL_Month = 1 then
      sL_EndDate := IntToStr(wL_Year-1) + '/' + IntToStr(12) + '/' + IntToStr(nL_LastDay)
    else
      sL_EndDate := IntToStr(wL_Year) + '/' + IntToStr(wL_Month - 1 ) + '/' + IntToStr(nL_LastDay);


    //fraChineseYMD1.setYMD(TUdateTime.CDateStr(StrToDate('2002/10/07'),9));
    //fraChineseYMD2.setYMD(TUdateTime.CDateStr(StrToDate('2002/10/07'),9));
    //實收日期
    fraChineseYMD1.setYMD(TUdateTime.CDateStr(StrToDate(sL_StartDate),9));
    fraChineseYMD2.setYMD(TUdateTime.CDateStr(StrToDate(sL_EndDate),9));

    //應收日期
    fraChineseYMD3.setYMD(TUdateTime.CDateStr(StrToDate(sL_StartDate),9));
    fraChineseYMD4.setYMD(TUdateTime.CDateStr(StrToDate(sL_EndDate),9));
}
end;

procedure TfrmSO8A50.showDetailExcel(sI_CompName,sI_ComputeYM,sI_ChargeItemNameSQL : String;nI_RealOrShouldDate : Integer);
var
    sL_CurrDateTime,sL_FileName,sL_RealOrShouldDate : String;
    L_StrList : TStringList;
    L_ExcelExport : TscExcelExport;
    sL_QueryString,sL_QueryDate,sL_QueryUser : String;
    sL_QueryCompName,sL_QueryBasic,sL_QueryYM : String;
    sL_QueryCitemName,sL_FileCurrDateTime : String;
begin
    sL_FileCurrDateTime := DateTimeToStr(now);
    sL_CurrDateTime := dtmMain2.ReplaceStr(Trim(dtmMain2.ReplaceStr(sL_FileCurrDateTime,'/')),':');


    sL_QueryDate := '計算時間: ' + sL_FileCurrDateTime;
    sL_QueryUser := '計算人員: ' + frmMainMenu.sG_User;
    sL_QueryCompName := '公司名稱: ' + sI_CompName;

    if nI_RealOrShouldDate = REAL_DATE_TYPE then
      sL_RealOrShouldDate := '實收日期'
    else if nI_RealOrShouldDate = SHOULD_DATE_TYPE then
      sL_RealOrShouldDate := '應收日期';

    sL_QueryBasic := '查詢年月基準: ' + sL_RealOrShouldDate;
    sL_QueryYM := '查詢年月: ' + sI_ComputeYM;
    sL_QueryCitemName := '查詢收費項目: ' + sI_ChargeItemNameSQL;

    sL_QueryString := sL_QueryDate + '@' + sL_QueryUser + '@' +
                      sL_QueryCompName + '@' + sL_QueryBasic + '@' +
                      sL_QueryYM + '@' + sL_QueryCitemName;

    sL_QueryString := TUstr.replaceStr(sL_QueryString,'劇','戲');
//showmessage('Liga 注意看喔~~~    ' + sL_QueryString);

    //先轉正常資料
    if dtmMain2.cdsSO131Excel.RecordCount <> 0 then
    begin
                          
      sL_FileName := 'C:\拆帳' + sL_CurrDateTime + '.xls';

      dtmMain2.cdsSO131Excel.DisableControls;
      DataSetToXLS(dtmMain2.cdsSO131Excel,sL_FileName,sL_QueryString,
        'computeym,shouldamt,realamt,RefNo,REALDATE,SHOULDDATE,REALSTARTDATE,REALSTOPDATE');
      dtmMain2.cdsSO131Excel.EnableControls;

      MessageDlg('計算完成,資料存於 ' + sL_FileName,mtInformation,[mbOK],0);
{
      L_ExcelExport := TscExcelExport.Create(Application);

      sL_CurrDateTime := DateTimeToStr(now);

      L_StrList := TStringList.Create;
      L_StrList.Add('計算時間: ' + sL_CurrDateTime);
      L_StrList.Add('計算人員: ' + dtmMain2.sG_User);
      L_StrList.Add('公司名稱: ' + sI_CompName);

      if nI_RealOrShouldDate = REAL_DATE_TYPE then
        sL_RealOrShouldDate := '實收日期'
      else if nI_RealOrShouldDate = SHOULD_DATE_TYPE then
        sL_RealOrShouldDate := '應收日期';

      L_StrList.Add('查詢年月基準: ' + sL_RealOrShouldDate);

      L_StrList.Add('查詢年月: ' + sI_ComputeYM);
      L_StrList.Add('查詢收費項目: ' + sI_ChargeItemNameSQL);
      L_ExcelExport.HeaderText := L_StrList;

      //設定顏色
      L_ExcelExport.FontHeader.Size := 9;
      L_ExcelExport.FontTitles.Size := 9;
      L_ExcelExport.FontData.Size := 9;

      L_ExcelExport.FontHeader.Color := clBlue;
      L_ExcelExport.FontTitles.Color := clMaroon;
      L_ExcelExport.FontData.Color := clNavy;

      //L_ExcelExport.BorderData.BackColor := clSkyBlue;
      L_ExcelExport.BorderTitles.BackColor := clFuchsia;
      L_ExcelExport.BorderHeader.BackColor := clWhite;



      //此方式為直接將CDS轉為EXCEL
      sL_CurrDateTime := dtmMain2.ReplaceStr(Trim(dtmMain2.ReplaceStr(sL_CurrDateTime,'/')),':');
      sL_FileName := 'C:\拆帳' + sL_CurrDateTime + '.xls';


      L_ExcelExport.Dataset := dtmMain2.cdsSO131;
      L_ExcelExport.ExportDataset(true);
      L_ExcelExport.ExcelVisible := false;
      L_ExcelExport.SaveAs(sL_FileName,ffXLS);

      L_ExcelExport.Free;
}
    end
    else
      MessageDlg('查無資料',mtInformation,[mbOK],0);
      

    //再轉呆帳資料
    if dtmMain2.cdsOtherSO131Excel.RecordCount <> 0 then
    begin

      sL_FileName := 'C:\拆帳(呆帳)' + sL_CurrDateTime + '.xls';

      dtmMain2.cdsOtherSO131Excel.DisableControls;
      DataSetToXLS(dtmMain2.cdsOtherSO131Excel,sL_FileName,sL_QueryString,
        'computeym,shouldamt,realamt,RefNo,REALDATE,SHOULDDATE,REALSTARTDATE,REALSTOPDATE');
      dtmMain2.cdsOtherSO131Excel.EnableControls;

      MessageDlg('有呆帳資料,存於 ' + sL_FileName,mtInformation,[mbOK],0);
    end;
end;



procedure TfrmSO8A50.FormDestroy(Sender: TObject);
begin
    G_ChargeCodeStrList.Free;
    G_TargetFieldValueStrList1.Free;
    G_TempChargeItem.Free;
end;

end.
