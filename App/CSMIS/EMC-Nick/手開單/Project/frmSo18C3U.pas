unit frmSo18C3U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, fraChineseYMDU, Buttons, ExtCtrls, Grids, DBGrids, DB;

type
  TfrmSo18C3 = class(TForm)
    Panel1: TPanel;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    StaticText2: TStaticText;
    StaticText3: TStaticText;
    StaticText4: TStaticText;
    fraChineseYMD1: TfraChineseYMD;
    Label1: TLabel;
    fraChineseYMD2: TfraChineseYMD;
    lbxEmp: TListBox;
    edtBeginNum: TEdit;
    edtEndNum: TEdit;
    Label2: TLabel;
    rgpReportType: TRadioGroup;
    BitBtn3: TBitBtn;
    procedure FormShow(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BitBtn3Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
  private
    { Private declarations }
    sG_EmpsCode, sG_EmpsName : String;
    sG_EmpSQL : String;    
  public
    { Public declarations }
  end;

var
  frmSo18C3: TfrmSo18C3;

implementation

uses frmSo18C0U, dtmMainU, frmDbMultiSelectu, Ustru, rptReport1U,
  rptReport2U, rptReport3U;

{$R *.dfm}

procedure TfrmSo18C3.FormShow(Sender: TObject);
begin
    fraChineseYMD1.mseYMD.SetFocus;
    sG_EmpsName := '';
    self.Caption := frmSo18C0.getCaption('SO8C30','手開單據管理-報表列印');
end;

procedure TfrmSo18C3.BitBtn2Click(Sender: TObject);
begin
    Close;
end;

procedure TfrmSo18C3.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    Action := caFree;
end;

procedure TfrmSo18C3.BitBtn3Click(Sender: TObject);
var
    ii : Integer;
    L_TargetFieldNamesStrList : TStringList;
    L_TargetFieldValueStrList : TStringList;

    bL_SelectedAllData : boolean;
    sL_TargetSQL, sL_Tmp : String;
    sL_Code, sL_Desc : String;
    L_TempStrList : TStringList;
begin

  L_TargetFieldNamesStrList:=TStringList.Create;
  L_TargetFieldNamesStrList.Add('EMPNO');
  L_TargetFieldNamesStrList.Add('EMPNAME');

  L_TargetFieldValueStrList:=TStringList.Create;
  bL_SelectedAllData := false;


  dtmMain.activeDataSet(dtmMain.qryCm003);
  dtmMain.qryCm003.FieldByName('EMPNO').DisplayLabel := '員工編號';
  dtmMain.qryCm003.FieldByName('EMPNAME').DisplayLabel := '姓名';

  if SelectMultiRecords('請點選員工', dtmMain.qryCm003, 'EMPNO;EMPNAME', ',', L_TargetFieldNamesStrList, L_TargetFieldValueStrList,true,sL_TargetSQL) = mrOk then
  begin
    lbxEmp.Items.Clear;
    sG_EmpsCode := '';
    for ii:=0 to L_TargetFieldValueStrList.Count -1 do
    begin
      sL_Tmp := Trim(L_TargetFieldValueStrList.Strings[ii]);
      L_TempStrList := TUstr.ParseStrings(sL_Tmp,',');
      sL_Code := L_TempStrList.Strings[0];
      sL_Desc := L_TempStrList.Strings[1];
      if sG_EmpsCode='' then
      begin
        sG_EmpsCode := sL_Code;
        sG_EmpsName := sL_Desc;
      end
      else
      begin
        sG_EmpsCode := sG_EmpsCode + ',' + sL_Code;
        sG_EmpsName := sG_EmpsName + ',' + sL_Desc;
      end;
      lbxEmp.Items.Add(sL_Desc);
    end;
    sG_EmpSQL := sL_TargetSQL;

  end;

end;

procedure TfrmSo18C3.BitBtn1Click(Sender: TObject);
var
    L_Rpt1: TrptReport1;
    L_Rpt2: TrptReport2;
    L_Rpt3: TrptReport3;            
    sL_BeginDate, sL_EndDate, sL_EmpSQL, sL_PaperNum:String;
    nL_ReportType : Integer;
    sL_BeginPaperNum, sL_EndPaperNum : String;
    sL_ReportGetPaperDateStr, sL_ReportPaperNumStr : String;
begin
    sL_BeginDate := trim(TUstr.replaceStr(fraChineseYMD1.getYMD,'/',''));
    sL_EndDate := trim(TUstr.replaceStr(fraChineseYMD2.getYMD,'/',''));

    if ((sL_BeginDate='') and (sL_EndDate<>'')) or
       ((sL_BeginDate<>'') and (sL_EndDate=''))then
    begin
      MessageDlg('請輸入領單之起始與截止日期!', mtWarning, [mbOK],0);
      Exit;
    end;
    if sL_BeginDate<>'' then
      sL_ReportGetPaperDateStr := sL_BeginDate + ' ~ ' + sL_EndDate
    else
      sL_ReportGetPaperDateStr := '*';

    sL_BeginPaperNum := edtBeginNum.Text;
    sL_EndPaperNum := edtEndNum.Text;
    if ((sL_BeginPaperNum='') and (sL_EndPaperNum<>'')) or
       ((sL_BeginPaperNum<>'') and (sL_EndPaperNum=''))then
    begin
      if sL_BeginPaperNum='' then
        edtBeginNum.SetFocus;
      if sL_EndPaperNum='' then
        edtEndNum.SetFocus;

      MessageDlg('請輸入單號範圍!', mtWarning, [mbOK],0);
      Exit;
    end;


    if sL_BeginPaperNum<>'' then
    begin
      sL_BeginPaperNum := TUstr.AddString(sL_BeginPaperNum,'0',false,PAPER_NUMBER_LENGTH);
      sL_EndPaperNum := TUstr.AddString(sL_EndPaperNum,'0',false,PAPER_NUMBER_LENGTH);
      sL_ReportPaperNumStr := sL_BeginPaperNum + ' - ' + sL_EndPaperNum
    end
    else
      sL_ReportPaperNumStr := '*';
    nL_ReportType := rgpReportType.ItemIndex;
    sL_EmpSQL := sG_EmpSQL;

    if dtmMain.activeReportData(sL_BeginDate, sL_EndDate, sL_EmpSQL, sL_BeginPaperNum, sL_EndPaperNum,nL_ReportType) then
    begin
      case rgpReportType.ItemIndex of
        0://手開單據領用明細表
         begin
           L_Rpt1:= TrptReport1.Create(Application);
           L_Rpt1.sG_CompName := frmSo18C0.sG_CompName;
           L_Rpt1.sG_ReportGetPaperDateStr := sL_ReportGetPaperDateStr;
           L_Rpt1.sG_ReportPaperNumStr := sL_ReportPaperNumStr;
           if sG_EmpsName<>'' then
             L_Rpt1.sG_GetPaperEmpsName := sG_EmpsName
           else
             L_Rpt1.sG_GetPaperEmpsName := '*';
           L_Rpt1.sG_Operator := frmSo18C0.sG_Operator;
           L_Rpt1.Preview;
           L_Rpt1.Free;
         end;
        1://未回報手開單據清冊 => 單據編號無值 and 沒作廢者
         begin
           L_Rpt2:= TrptReport2.Create(Application);
           L_Rpt2.sG_CompName := frmSo18C0.sG_CompName;
           L_Rpt2.sG_ReportGetPaperDateStr := sL_ReportGetPaperDateStr;
           L_Rpt2.sG_ReportPaperNumStr := sL_ReportPaperNumStr;
           if sG_EmpsName<>'' then
             L_Rpt2.sG_GetPaperEmpsName := sG_EmpsName
           else
             L_Rpt2.sG_GetPaperEmpsName := '*';
           L_Rpt2.sG_Operator := frmSo18C0.sG_Operator;
           L_Rpt2.Preview;
           L_Rpt2.Free;
         end;
        2://手開單使用情形 => 單據編號有值 or 作廢者
         begin
           L_Rpt3:= TrptReport3.Create(Application);
           L_Rpt3.sG_CompName := frmSo18C0.sG_CompName;
           L_Rpt3.sG_ReportGetPaperDateStr := sL_ReportGetPaperDateStr;
           L_Rpt3.sG_ReportPaperNumStr := sL_ReportPaperNumStr;
           if sG_EmpsName<>'' then
             L_Rpt3.sG_GetPaperEmpsName := sG_EmpsName
           else
             L_Rpt3.sG_GetPaperEmpsName := '*';
           L_Rpt3.sG_Operator := frmSo18C0.sG_Operator;
           L_Rpt3.Preview;
           L_Rpt3.Free;
         end;
      end;
    end
    else
      MessageDlg('查無相關資料!', mtInformation, [mbOK],0);
end;

end.
