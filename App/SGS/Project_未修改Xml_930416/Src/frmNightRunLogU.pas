unit frmNightRunLogU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, fraYMDU, Buttons, DB, Grids, DBGrids, Mask,ADODB;

type
  TfrmNightRunLog = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    dbgNightRunLog: TDBGrid;
    Panel4: TPanel;
    lblComp: TLabel;
    lblQueryDate: TLabel;
    Label5: TLabel;
    cobCompCode: TComboBox;
    fraYMD1: TfraYMD;
    fraYMD2: TfraYMD;
    btnOK: TBitBtn;
    medETime: TMaskEdit;
    medSTime: TMaskEdit;
    cobQueryType: TComboBox;
    lblQueryType: TLabel;
    Panel5: TPanel;
    dsrNightRunLod: TDataSource;
    lblComp2: TLabel;
    cobCompCode2: TComboBox;
    lblQueryType2: TLabel;
    cobQueryType2: TComboBox;
    lblQueryDate2: TLabel;
    fraYMD3: TfraYMD;
    btnAction: TBitBtn;
    lblActionAgain: TLabel;
    btnExit: TBitBtn;
    Label1: TLabel;
    fraYMD4: TfraYMD;
    memErrorLog: TMemo;
    procedure btnExitClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure fraYMD1Exit(Sender: TObject);
    procedure dbgNightRunLogDblClick(Sender: TObject);
    procedure btnActionClick(Sender: TObject);
    procedure fraYMD3Exit(Sender: TObject);
  private
    { Private declarations }
    nG_NightRunLogCounts : Integer;
    procedure setLanguageString;
    function getCobCompCode : String;
    function getCobCompCode2 : String;
    function IsDataOk : Boolean;
    function IsDataOk2 : Boolean;
    procedure initialQueryDate;

  public
    { Public declarations }
    procedure showLogMsg(sI_LogMsg : String);
  end;

var
  frmNightRunLog: TfrmNightRunLog;

implementation

uses dtmMainU, Ustru, UdateTimeu, ConstU, frmMainU;

{$R *.dfm}

procedure TfrmNightRunLog.btnExitClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmNightRunLog.setLanguageString;
begin
    Caption := dtmMain.getLanguageSettingInfo('frmNightRunLog_Caption');
    Panel1.Caption := dtmMain.getLanguageSettingInfo('frmNightRunLog_Panel1e_Caption');
    lblComp.Caption := dtmMain.getLanguageSettingInfo('frmNightRunLog_lblComp_Caption');
    lblQueryType.Caption := dtmMain.getLanguageSettingInfo('frmNightRunLog_lblQueryType_Caption');
    lblQueryDate.Caption := dtmMain.getLanguageSettingInfo('frmNightRunLog_lblQueryDate_Caption');
    lblComp2.Caption := dtmMain.getLanguageSettingInfo('frmNightRunLog_lblComp_Caption');
    lblQueryType2.Caption := dtmMain.getLanguageSettingInfo('frmNightRunLog_lblQueryType_Caption');
    lblQueryDate2.Caption := dtmMain.getLanguageSettingInfo('frmNightRunLog_lblQueryDate_Caption');
    btnOK.Caption := dtmMain.getLanguageSettingInfo('frmNightRunLog_btnOK_Caption');
    btnAction.Caption := dtmMain.getLanguageSettingInfo('frmNightRunLog_btnAction_Caption');
    btnExit.Caption := dtmMain.getLanguageSettingInfo('frmNightRunLog_btnExit_Caption');
    lblActionAgain.Caption := dtmMain.getLanguageSettingInfo('frmNightRunLog_lblActionAgain_Caption');

    cobQueryType.Items.Add(dtmMain.getLanguageSettingInfo('frmNightRunLog_rgpQueryType_Item1_Caption'));
    cobQueryType.Items.Add(dtmMain.getLanguageSettingInfo('frmNightRunLog_rgpQueryType_Item2_Caption'));
    cobQueryType.Items.Add(dtmMain.getLanguageSettingInfo('frmNightRunLog_rgpQueryType_Item3_Caption'));
    cobQueryType.Items.Add(dtmMain.getLanguageSettingInfo('frmNightRunLog_rgpQueryType_Item4_Caption'));

    cobQueryType2.Items.Add(dtmMain.getLanguageSettingInfo('frmNightRunLog_rgpQueryType_Item1_Caption'));
    cobQueryType2.Items.Add(dtmMain.getLanguageSettingInfo('frmNightRunLog_rgpQueryType_Item2_Caption'));
    cobQueryType2.Items.Add(dtmMain.getLanguageSettingInfo('frmNightRunLog_rgpQueryType_Item3_Caption'));
    cobQueryType2.Items.Add(dtmMain.getLanguageSettingInfo('frmNightRunLog_rgpQueryType_Item4_Caption'));
end;

procedure TfrmNightRunLog.FormShow(Sender: TObject);
var
  ii : Integer;
begin
  cobCompCode.Clear;
  for ii:=0 to dtmMain.G_CompCodeAndNameStrList.Count -1 do
    cobCompCode.Items.Add(dtmMain.G_CompCodeAndNameStrList.Strings[ii]);

  cobCompCode2.Clear;
  for ii:=0 to dtmMain.G_CompCodeAndNameStrList.Count -1 do
    cobCompCode2.Items.Add(dtmMain.G_CompCodeAndNameStrList.Strings[ii]);

  setLanguageString;

  cobCompCode.ItemIndex := 0;
  cobQueryType.ItemIndex := 0;
  cobCompCode2.ItemIndex := 0;
  cobQueryType2.ItemIndex := 0;
  initialQueryDate;
  nG_NightRunLogCounts := 0;


  dtmMain.adoQryNightRunLog.Close;
end;

function TfrmNightRunLog.getCobCompCode: String;
var
    L_CompCodeStrList : TStringList;
    sL_CompCodeAndName : String;
begin
    //取得所屬公司
    sL_CompCodeAndName := cobCompCode.Text;

    if sL_CompCodeAndName <> '' then
    begin
      L_CompCodeStrList := TStringList.Create;
      L_CompCodeStrList := TUstr.ParseStrings(cobCompCode.Text,'-');
      Result := L_CompCodeStrList.Strings[0];
    end
    else
      Result := '';

    L_CompCodeStrList.Free;
end;

procedure TfrmNightRunLog.btnOKClick(Sender: TObject);
var
    sL_CompCode,sL_SDate,sL_EDate,sL_ModeType : String;
    sL_STime,sL_ETime,sL_SDateTime,sL_EDateTime : String;
    sL_ErrMsg,sL_CompName : String;
    L_AdoQuery : TADOQuery;
begin
    if IsDataOk then
    begin
      //取得所屬公司
      sL_CompCode := getCobCompCode;

      sL_SDate := Trim(fraYMD1.getYMD);
      sL_EDate := Trim(fraYMD2.getYMD);

      //1:IC卡資訊2:產品定義資訊3:產品定購資訊4:授權資訊
      if cobQueryType.ItemIndex = 0 then
        sL_ModeType := '1'
      else if cobQueryType.ItemIndex = 1 then
        sL_ModeType := '2'
      else if cobQueryType.ItemIndex = 2 then
        sL_ModeType := '3'
      else if cobQueryType.ItemIndex = 3 then
        sL_ModeType := '4';

      sL_STime := Copy(Trim(medSTime.Text),0,2) + ':' + Copy(Trim(medSTime.Text),3,2) + ':00';
      sL_ETime := Copy(Trim(medETime.Text),0,2) + ':' + Copy(Trim(medETime.Text),3,2) + ':59';

      sL_SDateTime := sL_SDate + ' ' + sL_STime;
      sL_EDateTime := sL_EDate + ' ' + sL_ETime;

      L_AdoQuery := dtmMain.getNightRunLogData(sL_CompCode,sL_ModeType,sL_SDateTime,sL_EDateTime);

      dtmMain.adoQryNightRunLog.Clone(L_AdoQuery,ltUnspecified);

      if L_AdoQuery=nil then
      begin
        sL_CompName := dtmMain.getCompName(sL_CompCode);
        sL_ErrMsg := dtmMain.getLanguageSettingInfo('dtmMain_Msg_4_Content') + '=<' + sL_CompName + '>';
        MessageDlg(sL_ErrMsg,mtError, [mbOK],0);
      end
      else
      begin
        dbgNightRunLog.Columns[0].Title.Caption := dtmMain.getLanguageSettingInfo('frmNightRunLog_HandleEDateTime_Caption');
        dbgNightRunLog.Columns[1].Title.Caption := dtmMain.getLanguageSettingInfo('frmNightRunLog_ExecDateTime_Caption');
        dbgNightRunLog.Columns[2].Title.Caption := dtmMain.getLanguageSettingInfo('frmNightRunLog_ErrorCode_Caption');
        dbgNightRunLog.Columns[3].Title.Caption := dtmMain.getLanguageSettingInfo('frmNightRunLog_ErrorMsg_Caption');

        nG_NightRunLogCounts := dtmMain.adoQryNightRunLog.RecordCount;
      end;

    end;
end;

function TfrmNightRunLog.IsDataOk: Boolean;
var
    sL_CompCode,sL_StartDate,sL_StopDate,sL_StartTime,sL_StopTime,sL_QueryType : String;
begin
  sL_CompCode := cobCompCode.Text;
  if sL_CompCode = '' then
  begin
      MessageDlg(dtmMain.getLanguageSettingInfo('frmNightRunLog_Msg_1_Content'),mtError, [mbOK],0);
      cobCompCode.SetFocus;
      IsDataOk := false;
      exit;
  end;

  sL_QueryType := cobQueryType.Text;
  if sL_QueryType = '' then
  begin
      MessageDlg(dtmMain.getLanguageSettingInfo('frmNightRunLog_Msg_5_Content'),mtError, [mbOK],0);
      cobQueryType.SetFocus;
      IsDataOk := false;
      exit;
  end;


  sL_StartDate := Trim(TUstr.replaceStr(Trim(fraYMD1.getYMD),'/',''));
  sL_StopDate := Trim(TUstr.replaceStr(Trim(fraYMD2.getYMD),'/',''));


  if sL_StartDate = '' then
  begin
    MessageDlg(dtmMain.getLanguageSettingInfo('frmNightRunLog_Msg_2_Content'),mtError, [mbOK],0);
    fraYMD1.setYMD('');
    fraYMD1.mseYMD.SetFocus;
    IsDataOk := false;
    exit;
  end;

  if sL_StopDate = '' then
  begin
    MessageDlg(dtmMain.getLanguageSettingInfo('frmNightRunLog_Msg_2_Content'),mtError, [mbOK],0);
    fraYMD2.setYMD('');
    fraYMD2.mseYMD.SetFocus;
    IsDataOk := false;
    exit;
  end;

  if StrToInt(sL_StartDate) > StrToInt(sL_StopDate) then
  begin
    MessageDlg(dtmMain.getLanguageSettingInfo('frmNightRunLog_Msg_3_Content'),mtError, [mbOK],0);
    fraYMD2.setYMD('');
    fraYMD2.mseYMD.SetFocus;
    IsDataOk := false;
    exit;
  end
  else if StrToInt(sL_StartDate) = StrToInt(sL_StopDate) then
  begin
    sL_StartTime := Trim(TUstr.replaceStr(Trim(medSTime.Text),':',''));
    sL_StopTime := Trim(TUstr.replaceStr(Trim(medETime.Text),':',''));
    if StrToInt(sL_StartTime) > StrToInt(sL_StopTime) then
    begin
      MessageDlg(dtmMain.getLanguageSettingInfo('frmNightRunLog_Msg_4_Content'),mtError, [mbOK],0);
      medETime.Text := '';
      medETime.SetFocus;
      IsDataOk := false;
      exit;
    end;
  end;


end;

procedure TfrmNightRunLog.fraYMD1Exit(Sender: TObject);
begin
    fraYMD2.setYMD(fraYMD1.getYMD);
end;

procedure TfrmNightRunLog.initialQueryDate;
var
    sL_CurrDate : String;
begin
    sL_CurrDate := TUdateTime.GetPureDateStr(NOW,'/');
    fraYMD1.setYMD(sL_CurrDate);
    medSTime.Text := '0000';

    fraYMD2.setYMD(sL_CurrDate);
    medETime.Text := '2359';
end;

procedure TfrmNightRunLog.dbgNightRunLogDblClick(Sender: TObject);
var
    dL_QueryDate : TDateTime;
    sL_QueryDate : String;
begin
    if nG_NightRunLogCounts > 0 then
    begin
      dL_QueryDate := dtmMain.adoQryNightRunLog.FieldByName('HandleEDateTime').AsDateTime;

      cobCompCode2.ItemIndex := cobCompCode2.ItemIndex;
      cobQueryType2.ItemIndex := cobQueryType.ItemIndex;
      sL_QueryDate := TUdateTime.GetPureDateStr(dL_QueryDate,'/');
      fraYMD3.setYMD(sL_QueryDate);
    end;
end;

function TfrmNightRunLog.getCobCompCode2: String;
var
    L_CompCodeStrList : TStringList;
    sL_CompCodeAndName : String;
begin
    //取得所屬公司
    sL_CompCodeAndName := cobCompCode2.Text;

    if sL_CompCodeAndName <> '' then
    begin
      L_CompCodeStrList := TStringList.Create;
      L_CompCodeStrList := TUstr.ParseStrings(cobCompCode2.Text,'-');
      Result := L_CompCodeStrList.Strings[0];
    end
    else
      Result := '';

    L_CompCodeStrList.Free;
end;

procedure TfrmNightRunLog.btnActionClick(Sender: TObject);
var
    sL_CompCode ,sL_QueryStartDate,sL_QueryStopDate,sL_ModeType : String;

begin
    if IsDataOk2 then
    begin
      //取得所屬公司
      sL_CompCode := getCobCompCode2;

      sL_QueryStartDate := Trim(fraYMD3.getYMD);
      sL_QueryStopDate := Trim(fraYMD4.getYMD);



      //1:IC卡資訊2:產品定義資訊3:產品定購資訊4:授權資訊
      if cobQueryType2.ItemIndex = 0 then
        sL_ModeType := '1'
      else if cobQueryType2.ItemIndex = 1 then
        sL_ModeType := '2'
      else if cobQueryType2.ItemIndex = 2 then
        sL_ModeType := '3'
      else if cobQueryType2.ItemIndex = 3 then
        sL_ModeType := '4';


      if MessageDlg(dtmMain.getLanguageSettingInfo('frmNightRunLog_Msg_8_Content'), mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        frmMain.StartNightRunProcess(false,sL_CompCode,sL_ModeType,sL_QueryStartDate,sL_QueryStopDate);

    end;
end;

function TfrmNightRunLog.IsDataOk2: Boolean;
var
    sL_CompCode,sL_QueryDate,sL_QueryType : String;
    sL_StartDate,sL_StopDate : String;
begin
  sL_CompCode := cobCompCode2.Text;
  if sL_CompCode = '' then
  begin
      MessageDlg(dtmMain.getLanguageSettingInfo('frmNightRunLog_Msg_1_Content'),mtError, [mbOK],0);
      cobCompCode2.SetFocus;
      Result := false;
      exit;
  end;

  sL_QueryType := cobQueryType2.Text;
  if sL_QueryType = '' then
  begin
      MessageDlg(dtmMain.getLanguageSettingInfo('frmNightRunLog_Msg_5_Content'),mtError, [mbOK],0);
      cobQueryType2.SetFocus;
      Result := false;
      exit;
  end;


  sL_StartDate := Trim(TUstr.replaceStr(Trim(fraYMD3.getYMD),'/',''));
  sL_StopDate := Trim(TUstr.replaceStr(Trim(fraYMD4.getYMD),'/',''));


  if sL_StartDate = '' then
  begin
    MessageDlg(dtmMain.getLanguageSettingInfo('frmNightRunLog_Msg_2_Content'),mtError, [mbOK],0);
    fraYMD3.setYMD('');
    fraYMD3.mseYMD.SetFocus;
    IsDataOk2 := false;
    exit;
  end;

  if sL_StopDate = '' then
  begin
    MessageDlg(dtmMain.getLanguageSettingInfo('frmNightRunLog_Msg_2_Content'),mtError, [mbOK],0);
    fraYMD4.setYMD('');
    fraYMD4.mseYMD.SetFocus;
    IsDataOk2 := false;
    exit;
  end;



  if StrToInt(sL_StartDate) > StrToInt(sL_StopDate) then
  begin
    MessageDlg(dtmMain.getLanguageSettingInfo('frmNightRunLog_Msg_3_Content'),mtError, [mbOK],0);
    fraYMD4.setYMD('');
    fraYMD4.mseYMD.SetFocus;
    IsDataOk2 := false;
    exit;
  end;

  IsDataOk2 := true;
end;

procedure TfrmNightRunLog.fraYMD3Exit(Sender: TObject);
begin
    fraYMD4.setYMD(fraYMD3.getYMD);
end;

procedure TfrmNightRunLog.showLogMsg(sI_LogMsg: String);
begin
    memErrorLog.Lines.Add('[' + DateTimeToStr(now) + ']--' + sI_LogMsg);
end;

end.
