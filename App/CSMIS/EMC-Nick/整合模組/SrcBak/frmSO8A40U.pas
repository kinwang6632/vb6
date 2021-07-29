unit frmSO8A40U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, DBCtrls, Buttons, ComCtrls, DB, Grids, DBGrids,
  fraChineseYMDU, fraChineseYMU, ExtCtrls;

type
  TfrmSO8A40 = class(TForm)
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    chbShowDetail: TCheckBox;
    btnStartCalculate: TBitBtn;
    BitBtn1: TBitBtn;
    dsrCodeCD039A: TDataSource;
    cobComp: TComboBox;
    fraChineseYMD1: TfraChineseYMD;
    fraChineseYMD2: TfraChineseYMD;
    DataSource1: TDataSource;
    fraChineseYM1: TfraChineseYM;
    chbProviderGroup: TCheckBox;
    btnPrint: TBitBtn;
    rgpType: TRadioGroup;
    DataSource2: TDataSource;
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnStartCalculateClick(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
  private
    { Private declarations }
    function IsDataOk : Boolean;
    function getCobCompCode : String;
    function getCobCompName : String;
    procedure showReport(sI_CompName,sI_CompCode,sI_BelongYM,sI_ShowType,sI_TransData,sI_ProviderGroup : String);
  public
    { Public declarations }
  end;

var
  frmSO8A40: TfrmSO8A40;

implementation

uses  Ustru, UdateTimeu, UCommonU, rptSO8A40AU, rptSO8A40BU,
  rptSO8A40CU, dtmMain2U, frmMainMenuU;

{$R *.dfm}

procedure TfrmSO8A40.FormShow(Sender: TObject);
begin
    self.Caption := frmMainMenu.setCaption('SO8A40','[拆帳金額計算]');

    //控制正式資料是否已結算完成
    dtmMain2.bG_IsTransData := false;
end;

procedure TfrmSO8A40.FormCreate(Sender: TObject);
var
  wL_Year, wL_Month, wL_Day : Word;
  sL_StartDate,sL_EndDate : String;
  nL_LastDay : Integer;
begin
    TUCommonFun.AddObjectToComboBox(cobComp.Items, dsrCodeCD039A.DataSet,NOT INSERT_NO_DATA_ITEM,'CodeNo','Description');
    //指定為VB傳來之User所屬公司
    TUCommonFun.setComboDefaultNdx(cobComp, frmMainMenu.sG_CompCode);

    dtmMain2.cdsCodeCd039.Open;

    //開始實收日期 = 預設為上個月的第一天
    //結束實收日期 = 預設為上個月的最後一天
    // 歸 屬 日 期 = 預設為程式執行當天
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


    fraChineseYMD1.setYMD( sL_StartDate );
    fraChineseYMD2.setYMD( sL_EndDate );
    fraChineseYM1.setYM( FormatDateTime( 'yyyy/mm', Date ) );

end;

procedure TfrmSO8A40.btnStartCalculateClick(Sender: TObject);
var
    sL_StartDate,sL_EndDate,sL_BelongYM,sL_CompCode,sL_CompName : String;
    bL_HaveData : Boolean;
    sL_ShowDetail,sL_TransData,sL_Result,sL_ProviderGroup : String;
begin

    if IsDataOk then
    begin
      //執行等待狀態
      TUCommonFun.setWaitingCursor;

      sL_CompCode := getCobCompCode;
      sL_CompName := getCobCompName;
      sL_StartDate := fraChineseYMD1.getYMD;
      sL_EndDate := fraChineseYMD2.getYMD;
      sL_BelongYM := fraChineseYM1.getYM;
      
      //國曆改成西曆
      //sL_StartDate := dtmMain2.chineseDateChangeToEnglishDate(sL_StartDate);
      //sL_EndDate := dtmMain2.chineseDateChangeToEnglishDate(sL_EndDate);
      //sL_BelongYM := dtmMain2.chineseYMChangeToEnglishYM(sL_BelongYM);

      //到So114中檢查該So與歸屬日期之data是否已經存在
      bL_HaveData := dtmMain2.checkSO114HaveBelongYM(sL_CompCode,sL_BelongYM);

{
      if chbShowDetail.Checked then
        sL_ShowDetail := 'Y'
      else
        sL_ShowDetail := 'N';
}

      //全部改為顯示明細
      sL_ShowDetail := 'Y';


      if chbProviderGroup.Checked then  //依供應商分頁
        sL_ProviderGroup := 'Y'
      else
        sL_ProviderGroup := 'N';

      //初值設為可以轉為正式資料
      sL_TransData := 'Y';

      if bL_HaveData then
      begin
        //showmessage('歸屬日期有資料');
        if MessageDlg('已有資料是否要重新計算?',mtConfirmation, [mbYes, mbNo], 0) = mrNo then
        begin
          //showmessage('不重新計算');
          //不重新計算就取得原資料
          dtmMain2.getReportData(sL_compcode, sL_BelongYM);
          //回復原狀態
          TUCommonFun.setDefaultCursor;
          sL_TransData := 'N';  //為報表上不能轉為正式資料
          dtmMain2.bG_IsTransData := true;
          //showReport(sL_CompName,sL_CompCode,sL_BelongYM,sL_ShowDetail,sL_TransData,sL_ProviderGroup);
        end
        else
        begin
          //showmessage('重新計算');
          sL_Result := dtmMain2.calculateData(sL_CompCode,sL_StartDate,sL_EndDate,sL_BelongYM,sL_ShowDetail);
          if sL_Result <> '' then
          begin
            MessageDlg(sL_Result,mtError, [mbOK],0);
            //回復原狀態
            TUCommonFun.setDefaultCursor;
            exit;
          end
          else
          begin
            //回復原狀態
            TUCommonFun.setDefaultCursor;
            showmessage('計算完成');
            //showReport(sL_CompName,sL_CompCode,sL_BelongYM,sL_ShowDetail,sL_TransData,sL_ProviderGroup);
          end;
        end;
      end
      else
      begin
        //showmessage('歸屬日期無資料');
        sL_Result := dtmMain2.calculateData(sL_CompCode,sL_StartDate,sL_EndDate,sL_BelongYM,sL_ShowDetail);
        if sL_Result <> '' then
        begin
          MessageDlg(sL_Result,mtError, [mbOK],0);
          //回復原狀態
          TUCommonFun.setDefaultCursor;
          exit;
        end
        else
        begin
          //回復原狀態
          TUCommonFun.setDefaultCursor;
          showmessage('計算完成');
          //showReport(sL_CompName,sL_CompCode,sL_BelongYM,sL_ShowDetail,sL_TransData,sL_ProviderGroup);
        end;
      end;
    end;

{
      if bL_HaveData then
      begin
        //showmessage('歸屬日期有資料');
        if MessageDlg('已有資料是否要重新計算?',mtConfirmation, [mbYes, mbNo], 0) = mrNo then
        begin
          //showmessage('不重新計算');
          //不重新計算就取得原資料
          dtmMain2.getReportData(sL_compcode, sL_BelongYM,sL_ShowDetail);
          //回復原狀態
          TUCommonFun.setDefaultCursor;
          sL_TransData := 'N';  //為報表上不能轉為正式資料
          dtmMain2.bG_IsTransData := true;
          showReport(sL_CompName,sL_CompCode,sL_BelongYM,sL_ShowDetail,sL_TransData,sL_ProviderGroup);
        end
        else
        begin
          //showmessage('重新計算');
          sL_Result := dtmMain2.calculateData(sL_CompCode,sL_StartDate,sL_EndDate,sL_BelongYM,sL_ShowDetail);
          if sL_Result <> '' then
          begin
            MessageDlg(sL_Result,mtError, [mbOK],0);
            //回復原狀態
            TUCommonFun.setDefaultCursor;
            exit;
          end
          else
          begin
            //回復原狀態
            TUCommonFun.setDefaultCursor;
            showReport(sL_CompName,sL_CompCode,sL_BelongYM,sL_ShowDetail,sL_TransData,sL_ProviderGroup);
          end;
        end;
      end
      else
      begin
        //showmessage('歸屬日期無資料');
        sL_Result := dtmMain2.calculateData(sL_CompCode,sL_StartDate,sL_EndDate,sL_BelongYM,sL_ShowDetail);
        if sL_Result <> '' then
        begin
          MessageDlg(sL_Result,mtError, [mbOK],0);
          //回復原狀態
          TUCommonFun.setDefaultCursor;
          exit;
        end
        else
        begin
          //回復原狀態
          TUCommonFun.setDefaultCursor;
          showReport(sL_CompName,sL_CompCode,sL_BelongYM,sL_ShowDetail,sL_TransData,sL_ProviderGroup);
        end;
      end;
    end;
}
    //回復原狀態
    TUCommonFun.setDefaultCursor;
end;

procedure TfrmSO8A40.BitBtn1Click(Sender: TObject);
begin
    Close;
end;




function TfrmSO8A40.IsDataOk: Boolean;
var
    sL_StartDate,sL_ReplaceSDate,sL_EndDate,sL_ReplaceEDate : String;
    sL_BelongYM,sL_ReplaceBelongYM : String;
    nL_StartDate,nL_EndDate : Integer;
    nL_BelongYM : Integer;
begin
  sL_StartDate := Trim(fraChineseYMD1.getYMD);
  sL_ReplaceSDate := Trim(dtmMain2.ReplaceStr(sL_StartDate,'/'));
  nL_StartDate := Length(sL_ReplaceSDate);

  //SHOWMESSAGE(sL_StartDate);
  //SHOWMESSAGE(sL_ReplaceSDate);
  //SHOWMESSAGE(IntToStr(nL_StartDate));
  //檢查實收日期開始日期
  if Trim(sL_ReplaceSDate) <> '' then
  begin
    if nL_StartDate <> 8 then
    begin
      MessageDlg('日期請輸入8位數',mtError, [mbOK],0);
      fraChineseYMD1.setYMD('');
      fraChineseYMD1.mseYMD.SetFocus;
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
        fraChineseYMD1.setYMD('');
        fraChineseYMD1.mseYMD.SetFocus;
        IsDataOk := false;
        exit;
      end
    end
  end
  else
  begin
    MessageDlg('請輸入開始日期',mtError, [mbOK],0);
    fraChineseYMD1.setYMD('');
    fraChineseYMD1.mseYMD.SetFocus;
    IsDataOk := false;
    exit;
  end;

//檢查實收日期結束日期
  sL_EndDate := Trim(fraChineseYMD2.getYMD);
  sL_ReplaceEDate := Trim(dtmMain2.ReplaceStr(sL_EndDate,'/'));
  nL_EndDate := Length(sL_ReplaceEDate);
  if Trim(sL_ReplaceEDate) <> '' then
  begin
    if nL_EndDate <> 8 then
    begin
      MessageDlg('日期請輸入8位數',mtError, [mbOK],0);
      fraChineseYMD2.setYMD('');
      fraChineseYMD2.mseYMD.SetFocus;
      IsDataOk := false;
      exit;
    end
    else
    begin
      //檢查日期格式是否錯誤
      if TUdateTime.IsDateStr(sL_EndDate,'/') = false then
      begin
        MessageDlg('日期格式錯誤',mtError, [mbOK],0);
        fraChineseYMD2.setYMD('');
        fraChineseYMD2.mseYMD.SetFocus;
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
    fraChineseYMD2.setYMD('');
    fraChineseYMD2.mseYMD.SetFocus;
    IsDataOk := false;
    exit;
  end;

  if sL_ReplaceSDate > sL_ReplaceEDate then
  begin
    MessageDlg('結束日期必須大於等於開始日期',mtError, [mbOK],0);
    fraChineseYMD2.setYMD('');
    fraChineseYMD2.mseYMD.SetFocus;
    IsDataOk := false;
    exit;
  end;


//檢查歸屬日期結束日期

  sL_BelongYM := Trim(fraChineseYM1.getYM);
  sL_ReplaceBelongYM := Trim(dtmMain2.ReplaceStr(sL_BelongYM,'/'));
  nL_BelongYM := Length(sL_ReplaceBelongYM);
  if Trim(sL_ReplaceBelongYM) <> '' then
  begin
    if nL_BelongYM <> 6 then
    begin
      MessageDlg('歸屬日期請輸入6位數',mtError, [mbOK],0);
      fraChineseYM1.setYM('');
      fraChineseYM1.mseYM.SetFocus;
      IsDataOk := false;
      exit;
    end
    else
    begin
      //檢查日期格式是否錯誤
      if TUdateTime.IsYMStr(sL_BelongYM,'/') = false then
      begin
        MessageDlg('日期格式錯誤',mtError, [mbOK],0);
        fraChineseYM1.setYM('');
        fraChineseYM1.mseYM.SetFocus;
        IsDataOk := false;
        exit;
      end
      else
        IsDataOk := true;
    end
  end
  else
  begin
    MessageDlg('請輸入歸屬日期',mtError, [mbOK],0);
    fraChineseYM1.setYM('');
    fraChineseYM1.mseYM.SetFocus;
    IsDataOk := false;
    exit;
  end;
end;

function TfrmSO8A40.getCobCompCode: String;
var
    L_CompCodeStrList : TStringList;
    sL_CompCodeAndName : String;
begin
    //取得所屬公司
    sL_CompCodeAndName := cobComp.Text;

    L_CompCodeStrList := TStringList.Create;
    L_CompCodeStrList := TUstr.ParseStrings(cobComp.Text,'-');
    getCobCompCode := L_CompCodeStrList.Strings[0];

    L_CompCodeStrList.Free;

end;

function TfrmSO8A40.getCobCompName: String;
var
    L_CompNameStrList : TStringList;
    sL_CompCodeAndName : String;
begin
    //取得所屬公司
    sL_CompCodeAndName := cobComp.Text;

    L_CompNameStrList := TStringList.Create;
    L_CompNameStrList := TUstr.ParseStrings(cobComp.Text,'-');
    getCobCompName := L_CompNameStrList.Strings[1];

    L_CompNameStrList.Free;

end;

procedure TfrmSO8A40.showReport(sI_CompName,sI_CompCode,sI_BelongYM,sI_ShowType,sI_TransData,sI_ProviderGroup : String);
var
    L_Rpt1 : TrptSO8A40A;
    L_Rpt2 : TrptSO8A40B;
    //L_Rpt3 : TrptSO8A40C;
    sL_StartDate,sL_EndDate : String;
begin



    if dtmMain2.cdsSO114.Locate('COMP_ID;BELONGYM', VarArrayOf([sI_CompCode, sI_BelongYM]), []) then
    begin
      sL_StartDate := dtmMain2.cdsSO114.FieldByName('BEGIN_DATE').AsString;
      sL_StartDate := TUdateTime.GetPureDateStr(StrToDate(sL_StartDate));

      sL_EndDate := dtmMain2.cdsSO114.FieldByName('END_DATE').AsString;
      sL_EndDate := TUdateTime.GetPureDateStr(StrToDate(sL_EndDate));
    end;

    //顯示頻道拆帳明細表
    if sI_ShowType = '0' then
    begin
      L_Rpt1 := TrptSO8A40A.Create(Application);

      L_Rpt1.sG_TransData := sI_TransData;
      L_Rpt1.sG_CompCode := sI_CompCode;
      L_Rpt1.sG_CompName := sI_CompName;
      L_Rpt1.sG_StartDate := sL_StartDate;
      L_Rpt1.sG_EndDate := sL_EndDate;
      L_Rpt1.sG_BelongYM := sI_BelongYM;
      L_Rpt1.sG_ShowDetail := 'Y';
      L_Rpt1.sG_ProviderGroup := sI_ProviderGroup;
      L_Rpt1.sG_ShowType := sI_ShowType;

      L_Rpt1.Preview;
      L_Rpt1.Free;
      L_Rpt1 := nil;
    end;


    //如果已將資料正式結算
{
    if dtmMain2.bG_IsTransData then
    begin
      dtmMain2.getSO114Data(sI_compcode, sI_BelongYM);
      L_Rpt3 := TrptSO8A40C.Create(Application);

      L_Rpt3.sG_TransData := sI_TransData;
      L_Rpt3.sG_CompCode := sI_CompCode;
      L_Rpt3.sG_CompName := sI_CompName;
      L_Rpt3.sG_StartDate := sL_StartDate;
      L_Rpt3.sG_EndDate := sL_EndDate;
      L_Rpt3.sG_BelongYM := sI_BelongYM;
      L_Rpt3.sG_ShowDetail := sI_ShowDetail;

      L_Rpt3.Preview;
      L_Rpt3.Free;
      L_Rpt3 := nil;
    end;
}

    //顯示客戶收視明細表
    if sI_ShowType = '1' then
    begin
      L_Rpt2 := TrptSO8A40B.Create(Application);

      L_Rpt2.sG_TransData := sI_TransData;
      L_Rpt2.sG_CompCode := sI_CompCode;
      L_Rpt2.sG_CompName := sI_CompName;
      L_Rpt2.sG_StartDate := sL_StartDate;
      L_Rpt2.sG_EndDate := sL_EndDate;
      L_Rpt2.sG_BelongYM := sI_BelongYM;
      L_Rpt2.sG_ShowDetail := 'Y';
      L_Rpt2.sG_ShowType := sI_ShowType;

      L_Rpt2.Preview;
      L_Rpt2.Free;
      L_Rpt2 := nil;
    end;

end;



procedure TfrmSO8A40.btnPrintClick(Sender: TObject);
var
    nL_RptType,nL_Cds114Counts,nL_Cds116Counts : Integer;
    sL_CompCode,sL_CompName,sL_StartDate,sL_EndDate,sL_BelongYM : String;
    sL_TransData,sL_ProviderGroup : String;
    bL_HaveData : Boolean;
begin
    if IsDataOk then
    begin
      sL_CompCode := getCobCompCode;
      sL_CompName := getCobCompName;
      sL_StartDate := fraChineseYMD1.getYMD;
      sL_EndDate := fraChineseYMD2.getYMD;
      sL_BelongYM := fraChineseYM1.getYM;

      //國曆改成西曆
      //sL_StartDate := dtmMain2.chineseDateChangeToEnglishDate(sL_StartDate);
      //sL_EndDate := dtmMain2.chineseDateChangeToEnglishDate(sL_EndDate);
      //sL_BelongYM := dtmMain2.chineseYMChangeToEnglishYM(sL_BelongYM);

      if chbProviderGroup.Checked then  //依供應商分頁
        sL_ProviderGroup := 'Y'
      else
        sL_ProviderGroup := 'N';

      //顯示報表型態
      nL_RptType := rgpType.ItemIndex;


      //檢查是否已經算過(CDS有無資料),如果沒有,
      //再看看資料庫有無該歸屬年月的資料,若還是沒有則請 User 重做
      nL_Cds114Counts := dtmMain2.cdsSO114.RecordCount;
      nL_Cds116Counts := dtmMain2.cdsSO116.RecordCount;

      if (nL_Cds114Counts=0) OR (nL_Cds116Counts=0) then
      begin
        sL_TransData := 'N';
        bL_HaveData := dtmMain2.checkSO114HaveBelongYM(sL_CompCode,sL_BelongYM);
        //如果CDS無資料,但資料庫已經算過此歸屬年月,則從資料庫Query資料
        if bL_HaveData then
        begin
          dtmMain2.getReportData(sL_compcode, sL_BelongYM);
          showReport(sL_CompName,sL_CompCode,sL_BelongYM,IntToStr(nL_RptType),sL_TransData,sL_ProviderGroup);
        end
        else
        begin
          MessageDlg('資料尚未計算',mtError, [mbOK],0);
          btnStartCalculate.SetFocus;
        end;
      end
      else
      begin
        sL_TransData := 'Y';
        showReport(sL_CompName,sL_CompCode,sL_BelongYM,IntToStr(nL_RptType),sL_TransData,sL_ProviderGroup);
      end;
    end;
end;

end.

