unit frmInvC02U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Mask, fraYMDU, Buttons, ExtCtrls, Grids, DBGrids, DB,
  frxClass, frxDBSet, frxDesgn;

type
  TfrmInvC02 = class(TForm)
    Panel1: TPanel;
    Label5: TLabel;
    Panel2: TPanel;
    Label2: TLabel;
    btnExit: TBitBtn;
    btnQuery: TBitBtn;
    fraEInvDate: TfraYMD;
    fraSInvDate: TfraYMD;
    Label3: TLabel;
    medSInvID: TMaskEdit;
    medEInvID: TMaskEdit;
    cobObsoleteReason: TComboBox;
    Label4: TLabel;
    frxDBDataset1: TfrxDBDataset;
    frxReport1: TfrxReport;
    frxReport2: TfrxReport;
    Button1: TButton;
    Button2: TButton;
    frxDesigner1: TfrxDesigner;
    procedure FormShow(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure fraSInvDateExit(Sender: TObject);
    procedure medSInvIDExit(Sender: TObject);
    procedure btnQueryClick(Sender: TObject);
    procedure medEInvIDExit(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure frxReport1GetValue(const VarName: String;
      var Value: Variant);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
    G_ObsoleteReasonStrList : TStringList;
    procedure initialComboBox;
    function isDataOK : Boolean;
  public
    { Public declarations }
  end;

var
  frmInvC02: TfrmInvC02;

implementation

uses frmMainU, dtmMainU, dtmMainJU, rptInvC02_1U;

{$R *.dfm}

procedure TfrmInvC02.FormShow(Sender: TObject);
begin
    self.Caption := frmMain.GetFormTitleString( 'C02', '發票作廢一覽表' );

    G_ObsoleteReasonStrList := TStringList.Create;

    //ComboBox初始化
    initialComboBox;
end;

procedure TfrmInvC02.btnExitClick(Sender: TObject);
begin
    G_ObsoleteReasonStrList.Free;
    Close;
end;

procedure TfrmInvC02.fraSInvDateExit(Sender: TObject);
begin
    fraEInvDate.setYMD(fraSInvDate.getYMD);
end;

procedure TfrmInvC02.medSInvIDExit(Sender: TObject);
begin
    medSInvID.Text := UpperCase(Trim(medSInvID.Text));
    medEInvID.Text := UpperCase(Trim(medSInvID.Text));
end;

function TfrmInvC02.isDataOK: Boolean;
var
    sL_SInvID,sL_EInvID,sL_SInvDate,sL_EInvDate : String;
begin
    Result := true;
    sL_SInvID := Trim(medSInvID.Text);
    sL_EInvID := Trim(medEInvID.Text);
    sL_SInvDate := Trim(fraSInvDate.getYMD);
    sL_EInvDate := Trim(fraEInvDate.getYMD);

    //檢核發票日期
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



    //檢核發票號碼
    if sL_SInvID<>'' then
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
end;

procedure TfrmInvC02.btnQueryClick(Sender: TObject);
var
  sL_CompID,sL_SInvID,sL_EInvID,sL_SInvDate,sL_EInvDate : String;
  sL_ObsoleteId,sL_ObsoleteReason : String;
  nL_Counts : Integer;
begin
  if isDataOK then
  begin
    sL_CompID := dtmMain.getCompID;
    sL_SInvID := Trim(medSInvID.Text);
    sL_EInvID := Trim(medEInvID.Text);


    sL_SInvDate := Trim(fraSInvDate.getYMD);
    sL_EInvDate := Trim(fraEInvDate.getYMD);
    if sL_SInvDate = EMPTY_DATE_STR then
    begin
      sL_SInvDate := '';
      sL_EInvDate := '';
    end;

    dtmMainJ.getComboBoxCode(cobObsoleteReason,G_ObsoleteReasonStrList,sL_ObsoleteId,sL_ObsoleteReason);

    if sL_ObsoleteId = DEFAULT_ALL_DATA_ITEM_CODE_VALUE then
      sL_ObsoleteId := '';

    nL_Counts := dtmMainJ.getObsoleteInvoiceData(sL_CompID,sL_SInvID,sL_EInvID,sL_SInvDate,sL_EInvDate,sL_ObsoleteId);

    if nL_Counts = 0 then
      MessageDlg('沒有符合的資料!',mtInformation,[mbOk],0)
    else
    begin
      rptInvC02_1 := TrptInvC02_1.Create(Application);

      rptInvC02_1.sG_CompID := sL_CompID;
      rptInvC02_1.sG_CompName := dtmMain.getCompName;
      rptInvC02_1.sG_SInvID := sL_SInvID;
      rptInvC02_1.sG_EInvID := sL_EInvID;
      rptInvC02_1.sG_SInvDate := sL_SInvDate;
      rptInvC02_1.sG_EInvDate := sL_EInvDate;

      rptInvC02_1.sG_ObsoleteReason := sL_ObsoleteReason;

      rptInvC02_1.sG_UserID := dtmMain.getLoginUser;

      rptInvC02_1.Preview;
      rptInvC02_1.Free;
    end;
  end;
end;

procedure TfrmInvC02.initialComboBox;
var
    sL_SQL : String;
begin
   //作廢原因
   sL_SQL := 'select ItemId, Description from inv006 order by ItemId';
   dtmMainJ.createComboBoxItem(cobObsoleteReason,sL_SQL,'ItemId','Description',true,G_ObsoleteReasonStrList);
   cobObsoleteReason.ItemIndex := 0;
end;

procedure TfrmInvC02.medEInvIDExit(Sender: TObject);
begin
    medEInvID.Text := UpperCase(Trim(medEInvID.Text));
end;


procedure TfrmInvC02.Button1Click(Sender: TObject);
var
  aPath : String;
begin
  aPath := IncludeTrailingPathDelimiter( ExtractFilePath( Application.ExeName ) ) +
    IncludeTrailingPathDelimiter( REPORT_FOLDER ) + 'FrptInvC02_1.fr3';
  frxReport2.LoadFromFile( aPath );
  frxReport2.DesignReport ;
end;

procedure TfrmInvC02.frxReport1GetValue(const VarName: String;
  var Value: Variant);
var
  sL_ConditionStr : String;
  sL_CompID,sL_SInvID,sL_EInvID,sL_SInvDate,sL_EInvDate : String;
  sL_ObsoleteId,sL_ObsoleteReason : String;
begin
  sL_CompID := dtmMain.getCompID;
  sL_SInvID := Trim(medSInvID.Text);
  sL_EInvID := Trim(medEInvID.Text);
  sL_SInvDate := Trim(fraSInvDate.getYMD);
  sL_EInvDate := Trim(fraEInvDate.getYMD);
  dtmMainJ.getComboBoxCode(cobObsoleteReason,G_ObsoleteReasonStrList,sL_ObsoleteId,sL_ObsoleteReason);

  if CompareText (VarName ,'lblCondition1' ) = 0 then
  begin //sL_SInvID
    if sL_SInvID = '' then
      sL_ConditionStr := '公司簡稱 : ' + dtmMain.getCompName + ' , 發票號碼 : 全部'
    else
      sL_ConditionStr := '公司簡稱 : ' + dtmMain.getCompName + ' , 發票號碼 : '  + sL_SInvID + ' ~ ' + sL_EInvID;
    Value := sL_ConditionStr;
  end
  else
  if CompareText (VarName ,'lblCondition2' ) = 0 then
  begin //sL_SInvDate
    if sL_SInvDate = '' then  
      sL_ConditionStr := '作廢原因 : ' + sL_ObsoleteReason + ' , 發票日期: 全部'
    else
      sL_ConditionStr := '作廢原因 : ' + sL_ObsoleteReason + ' , 發票日期: ' + sL_SInvDate + ' ~ ' + sL_EInvDate;
    Value := sL_ConditionStr;
  end
  else
  if CompareText (VarName ,'qrlOperator' ) = 0 then
  begin //dtmMain.getLoginUser
    Value := dtmMain.getLoginUser;
  end;
end;

procedure TfrmInvC02.Button2Click(Sender: TObject);
var
  sL_CompID,sL_SInvID,sL_EInvID,sL_SInvDate,sL_EInvDate : String;
  sL_ObsoleteId,sL_ObsoleteReason : String;
  nL_Counts : Integer;
  aPath: String;
begin
  if isDataOK then
  begin
    sL_CompID := dtmMain.getCompID;
    sL_SInvID := Trim(medSInvID.Text);
    sL_EInvID := Trim(medEInvID.Text);
    sL_SInvDate := Trim(fraSInvDate.getYMD);
    sL_EInvDate := Trim(fraEInvDate.getYMD);
    if sL_SInvDate = EMPTY_DATE_STR then
    begin
      sL_SInvDate := '';
      sL_EInvDate := '';
    end;

    dtmMainJ.getComboBoxCode(cobObsoleteReason,G_ObsoleteReasonStrList,sL_ObsoleteId,sL_ObsoleteReason);

    if sL_ObsoleteId = DEFAULT_ALL_DATA_ITEM_CODE_VALUE then
      sL_ObsoleteId := '';

    nL_Counts := dtmMainJ.getObsoleteInvoiceData(sL_CompID,sL_SInvID,sL_EInvID,sL_SInvDate,sL_EInvDate,sL_ObsoleteId);

    if nL_Counts = 0 then
      MessageDlg('沒有符合的資料!',mtInformation,[mbOk],0)
    else
    begin
      aPath := IncludeTrailingPathDelimiter( ExtractFilePath( Application.ExeName ) )+
        IncludeTrailingPathDelimiter( REPORT_FOLDER ) + 'FrptInvC02_1.fr3';
      frxReport1.LoadFromFile( aPath );
      frxReport1.ShowReport ;
    end;
  end;

end;


end.
