unit frmSO8C30U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, fraChineseYMDU, Buttons, ExtCtrls, Grids, DBGrids, DB;

type
  TfrmSO8C30 = class(TForm)
    Panel1: TPanel;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    StaticText2: TStaticText;
    StaticText3: TStaticText;
    fraChineseYMD1: TfraChineseYMD;
    Label1: TLabel;
    fraChineseYMD2: TfraChineseYMD;
    lbxEmp: TListBox;
    edtBeginNum: TEdit;
    edtEndNum: TEdit;
    Label2: TLabel;
    rgpReportType: TRadioGroup;
    BitBtn3: TBitBtn;
    rgpType2: TRadioGroup;
    Label3: TLabel;
    Label15: TLabel;
    edtPath1: TEdit;
    procedure FormShow(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BitBtn3Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure rgpReportTypeClick(Sender: TObject);
  private
    { Private declarations }
    sG_EmpsCode, sG_EmpsName : String;
    sG_EmpSQL : String;
    procedure DoReport(const aIndex: Integer; const aPaperDate, aPaperNumber: String);
    function CheckExcelDir: Boolean;
  public
    { Public declarations }
  end;

var
  frmSO8C30: TfrmSO8C30;

implementation

uses dtmMain3U, frmDbMultiSelectu, Ustru, rptSO8C30_1U,
  rptSO8C30_2U, rptSO8C30_3U, frmMainMenuU, rptSO8C30_4U;

{$R *.dfm}

procedure TfrmSO8C30.FormShow(Sender: TObject);
begin
  fraChineseYMD1.mseYMD.SetFocus;
  sG_EmpsName := '';
  self.Caption := frmMainMenu.setCaption('SO8C30','手開單據管理-報表列印');
end;

procedure TfrmSO8C30.BitBtn2Click(Sender: TObject);
begin
    Close;
end;

procedure TfrmSO8C30.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    Action := caFree;
end;

procedure TfrmSO8C30.BitBtn3Click(Sender: TObject);
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


  dtmMain3.activeDataSet(dtmMain3.qryCm003);
  dtmMain3.qryCm003.FieldByName('EMPNO').DisplayLabel := '員工編號';
  dtmMain3.qryCm003.FieldByName('EMPNAME').DisplayLabel := '姓名';

  if SelectMultiRecords('請點選員工', dtmMain3.qryCm003, 'EMPNO;EMPNAME', ',', L_TargetFieldNamesStrList, L_TargetFieldValueStrList,true,sL_TargetSQL) = mrOk then
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

procedure TfrmSO8C30.BitBtn1Click(Sender: TObject);
var
  sL_BeginDate, sL_EndDate, sL_EmpSQL, sL_PaperNum: String;
  nL_ReportType: Integer;
  sL_BeginPaperNum, sL_EndPaperNum: String;
  sL_ReportGetPaperDateStr, sL_ReportGetPaperDateStr2, sL_ReportPaperNumStr: String;
begin

  if ( rgpType2.ItemIndex = 1 ) and ( not CheckExcelDir ) then
  begin
    if edtPath1.CanFocus then edtPath1.SetFocus;
    Exit;
  end;

  sL_BeginDate := trim(TUstr.replaceStr(fraChineseYMD1.getYMD,'/',''));
  sL_EndDate := trim(TUstr.replaceStr(fraChineseYMD2.getYMD,'/',''));

  if ((sL_BeginDate='') and (sL_EndDate<>'')) or
     ((sL_BeginDate<>'') and (sL_EndDate=''))then
  begin
    MessageDlg('請輸入領單之起始與截止日期!', mtWarning, [mbOK],0);
    Exit;
  end;

  if ( sL_BeginDate <> EmptyStr ) then sL_BeginDate := TUstr.ToROCYMD( sL_BeginDate );
  if ( sL_EndDate <> EmptyStr ) then sL_EndDate := TUstr.ToROCYMD( sL_EndDate );

  if sL_BeginDate<>'' then
  begin
    sL_ReportGetPaperDateStr := fraChineseYMD1.getYMD + ' ~ ' + fraChineseYMD2.getYMD;
    sL_ReportGetPaperDateStr2 :=
      TUstr.ToROCYMD( fraChineseYMD1.getYMD ) + ' ~ ' +
      TUstr.ToROCYMD( fraChineseYMD2.getYMD );
  end else
  begin
    sL_ReportGetPaperDateStr := '*';
    sL_ReportGetPaperDateStr2 := '*';
  end;

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
  end else
    sL_ReportPaperNumStr := '*';

  nL_ReportType := rgpReportType.ItemIndex;
  sL_EmpSQL := sG_EmpSQL;

  if dtmMain3.activeReportData( sL_BeginDate, sL_EndDate, sL_EmpSQL, sL_BeginPaperNum,
    sL_EndPaperNum, nL_ReportType ) then
  begin
    if ( rgpType2.ItemIndex = 0 ) then
      DoReport( rgpReportType.ItemIndex, sL_ReportGetPaperDateStr2, sL_ReportPaperNumStr )
    else begin
      dtmMain3._8C30ToExcel( rgpReportType.ItemIndex, sL_ReportGetPaperDateStr2, sL_ReportPaperNumStr,
        sG_EmpsName, edtPath1.Text );
      MessageDlg( '計算完成,統計表存於'  + edtPath1.Text, mtInformation, [mbOK], 0 )
    end;
  end else
  begin
    MessageDlg('查無相關資料!', mtInformation, [mbOK],0);
  end;
end;

procedure TfrmSO8C30.rgpReportTypeClick(Sender: TObject);
begin
    if (Sender as TRadioGroup).ItemIndex = 3 then//表示點選領單續用報表
    begin
      edtBeginNum.Text := '';
      edtEndNum.Text := '';
      edtBeginNum.Enabled := false;
      edtEndNum.Enabled := false;
    end
    else
    begin
      edtBeginNum.Enabled := true;
      edtEndNum.Enabled := true;
    end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmSO8C30.DoReport(const aIndex: Integer; const aPaperDate, aPaperNumber: String);
var
  L_Rpt1: TrptSO8C30_1;
  L_Rpt2: TrptSO8C30_2;
  L_Rpt3: TrptSO8C30_3;
  L_Rpt4: TrptSO8C30_4;   
begin
   case aIndex of
     3: //領單續用報表
      begin
        L_Rpt4:= TrptSO8C30_4.Create(Application);
        L_Rpt4.sG_CompName := frmMainMenu.sG_CompName;
        L_Rpt4.sG_ReportGetPaperDateStr := aPaperDate;
        L_Rpt4.sG_ReportPaperNumStr := aPaperNumber;
        if sG_EmpsName<>'' then
          L_Rpt4.sG_GetPaperEmpsName := sG_EmpsName
        else
          L_Rpt4.sG_GetPaperEmpsName := '*';
        L_Rpt4.sG_Operator := frmMainMenu.sG_User;

        L_Rpt4.Preview;
        L_Rpt4.Free;

      end;
     0://手開單據領用明細表
      begin
        L_Rpt1:= TrptSO8C30_1.Create(Application);
        L_Rpt1.sG_CompName := frmMainMenu.sG_CompName;
        L_Rpt1.sG_ReportGetPaperDateStr := aPaperDate;
        L_Rpt1.sG_ReportPaperNumStr := aPaperNumber;
        if sG_EmpsName<>'' then
          L_Rpt1.sG_GetPaperEmpsName := sG_EmpsName
        else
          L_Rpt1.sG_GetPaperEmpsName := '*';
        L_Rpt1.sG_Operator := frmMainMenu.sG_User;
        L_Rpt1.Preview;
        L_Rpt1.Free;
      end;
     1://未回報手開單據清冊 => 單據編號無值 and 沒作廢者
      begin
        L_Rpt2:= TrptSO8C30_2.Create(Application);
        L_Rpt2.sG_CompName := frmMainMenu.sG_CompName;
        L_Rpt2.sG_ReportGetPaperDateStr := aPaperDate;
        L_Rpt2.sG_ReportPaperNumStr := aPaperNumber;
        if sG_EmpsName<>'' then
          L_Rpt2.sG_GetPaperEmpsName := sG_EmpsName
        else
          L_Rpt2.sG_GetPaperEmpsName := '*';
        L_Rpt2.sG_Operator := frmMainMenu.sG_User;
        L_Rpt2.Preview;
        L_Rpt2.Free;
      end;
     2://手開單使用情形 => 單據編號有值 or 作廢者
      begin
        L_Rpt3:= TrptSO8C30_3.Create(Application);
        L_Rpt3.sG_CompName := frmMainMenu.sG_CompName;
        L_Rpt3.sG_ReportGetPaperDateStr := aPaperDate;
        L_Rpt3.sG_ReportPaperNumStr := aPaperNumber;
        if sG_EmpsName<>'' then
          L_Rpt3.sG_GetPaperEmpsName := sG_EmpsName
        else
          L_Rpt3.sG_GetPaperEmpsName := '*';
        L_Rpt3.sG_Operator := frmMainMenu.sG_User;
        L_Rpt3.Preview;
        L_Rpt3.Free;
      end;
   end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmSO8C30.CheckExcelDir: Boolean;
begin
  Result := False;
  if ( Trim( edtPath1.Text ) = EmptyStr ) then
  begin
    MessageDlg( '請輸入Excel儲存路徑以及檔案名稱。', mtWarning,  [mbOK], 0 );
    Exit;
  end;
  if ( ExtractFilePath( edtPath1.Text ) = EmptyStr ) then
  begin
    MessageDlg( '請輸入Excel儲存路徑。', mtWarning,  [mbOK], 0 );
    Exit;
  end;
  if ( ExtractFileName( edtPath1.Text ) = EmptyStr ) then
  begin
    MessageDlg( '請輸入Excel儲存檔案名稱。', mtWarning,  [mbOK], 0 );
    Exit;
  end;
  if ( ExtractFileExt( edtPath1.Text ) = EmptyStr ) then
  begin
    MessageDlg( '請輸入Excel儲存檔案名稱。', mtWarning,  [mbOK], 0 );
    Exit;
  end;
  if ( not DirectoryExists( ExtractFilePath( edtPath1.Text ) ) ) Then
  begin
    MessageDlg( 'Excel儲存路徑不存在,請重新輸入。', mtWarning,  [mbOK], 0 );
    Exit;
  end;
  if ( FileExists( edtPath1.Text ) ) Then
  begin
    if MessageDlg( 'Excel儲存檔案已存在,是否覆蓋?。', mtWarning,
      [mbOK,mbCancel], 0 ) = mrCancel Then
    Exit;
  end;
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

end.
