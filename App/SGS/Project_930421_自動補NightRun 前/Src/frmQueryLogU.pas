unit frmQueryLogU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, fraYMDU, Buttons, DB, Grids, DBGrids, Mask,ADODB;

type
  TfrmQueryLog = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    cobCompCode: TComboBox;
    lblComp: TLabel;
    lblQueryDate: TLabel;
    Label5: TLabel;
    fraYMD1: TfraYMD;
    fraYMD2: TfraYMD;
    dbgQueryLogData: TDBGrid;
    dsrQueryLogData: TDataSource;
    btnOK: TBitBtn;
    btnExit: TBitBtn;
    lblTitle: TLabel;
    medETime: TMaskEdit;
    medSTime: TMaskEdit;
    lblQueryType: TLabel;
    cobQueryType: TComboBox;
    procedure btnExitClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure fraYMD1Exit(Sender: TObject);
  private
    { Private declarations }
    procedure setLanguageString;
    function getCobCompCode : String;
    function IsDataOk : Boolean;
    procedure initialQueryDate;
  public
    { Public declarations }
  end;

var
  frmQueryLog: TfrmQueryLog;

implementation

uses dtmMainU, Ustru, UdateTimeu, ConstU;

{$R *.dfm}

procedure TfrmQueryLog.btnExitClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmQueryLog.setLanguageString;
begin
    Caption := dtmMain.getLanguageSettingInfo('frmQueryLog_Caption');
    Panel1.Caption := dtmMain.getLanguageSettingInfo('frmQueryLog_Panel1e_Caption');
    lblComp.Caption := dtmMain.getLanguageSettingInfo('frmQueryLog_lblComp_Caption');
    lblQueryDate.Caption := dtmMain.getLanguageSettingInfo('frmQueryLog_lblQueryDate_Caption');
    btnOK.Caption := dtmMain.getLanguageSettingInfo('frmQueryLog_btnOK_Caption');
    btnExit.Caption := dtmMain.getLanguageSettingInfo('frmQueryLog_btnExit_Caption');
    lblQueryType.Caption := dtmMain.getLanguageSettingInfo('frmQueryLog_lblQueryType_Caption');

    cobQueryType.Items.Add(dtmMain.getLanguageSettingInfo('frmQueryLog_rgpQueryType_Item1_Caption'));
    cobQueryType.Items.Add(dtmMain.getLanguageSettingInfo('frmQueryLog_rgpQueryType_Item2_Caption'));
    cobQueryType.Items.Add(dtmMain.getLanguageSettingInfo('frmQueryLog_rgpQueryType_Item3_Caption'));
    cobQueryType.Items.Add(dtmMain.getLanguageSettingInfo('frmQueryLog_rgpQueryType_Item4_Caption'));
    cobQueryType.Items.Add(dtmMain.getLanguageSettingInfo('frmQueryLog_rgpQueryType_Item5_Caption'));

    dbgQueryLogData.Columns[0].Title.Caption := dtmMain.getLanguageSettingInfo('frmQueryLog_DstMsgID_Caption');
    dbgQueryLogData.Columns[1].Title.Caption := dtmMain.getLanguageSettingInfo('frmQueryLog_ModeType_Caption');
    dbgQueryLogData.Columns[2].Title.Caption := dtmMain.getLanguageSettingInfo('frmQueryLog_InstQuery_Caption');
    dbgQueryLogData.Columns[3].Title.Caption := dtmMain.getLanguageSettingInfo('frmQueryLog_DstCode_Caption');
    dbgQueryLogData.Columns[4].Title.Caption := dtmMain.getLanguageSettingInfo('frmQueryLog_SrcCode_Caption');
    dbgQueryLogData.Columns[5].Title.Caption := dtmMain.getLanguageSettingInfo('frmQueryLog_HandleEDateTime_Caption');
    dbgQueryLogData.Columns[6].Title.Caption := dtmMain.getLanguageSettingInfo('frmQueryLog_ExecDateTime_Caption');
    dbgQueryLogData.Columns[7].Title.Caption := dtmMain.getLanguageSettingInfo('frmQueryLog_SrcMsgID_Caption');
    dbgQueryLogData.Columns[8].Title.Caption := dtmMain.getLanguageSettingInfo('frmQueryLog_ErrorMsg_Caption');
end;

procedure TfrmQueryLog.FormShow(Sender: TObject);
var
  ii : Integer;
begin
  cobCompCode.Clear;
  for ii:=0 to dtmMain.G_CompCodeAndNameStrList.Count -1 do
    cobCompCode.Items.Add(dtmMain.G_CompCodeAndNameStrList.Strings[ii]);

  setLanguageString;
  cobCompCode.ItemIndex := 0;
  cobQueryType.ItemIndex := 0;

  initialQueryDate;
end;

function TfrmQueryLog.getCobCompCode: String;
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

procedure TfrmQueryLog.btnOKClick(Sender: TObject);
var
    sL_CompCode,sL_SDate,sL_EDate,sL_ModeType : String;
    sL_STime,sL_ETime,sL_SDateTime,sL_EDateTime : String;
    sL_ErrMsg,sL_CompName,sL_InstQuery,sL_QueryType : String;
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
        sL_ModeType := '4'
      else if cobQueryType.ItemIndex = 4 then
        sL_ModeType := '5';

      sL_STime := Copy(Trim(medSTime.Text),0,2) + ':' + Copy(Trim(medSTime.Text),3,2) + ':00';
      sL_ETime := Copy(Trim(medETime.Text),0,2) + ':' + Copy(Trim(medETime.Text),3,2) + ':59';

      sL_SDateTime := sL_SDate + ' ' + sL_STime;
      sL_EDateTime := sL_EDate + ' ' + sL_ETime;

      L_AdoQuery := dtmMain.getQueryLogData(sL_CompCode,sL_ModeType,sL_SDateTime,sL_EDateTime);

      dtmMain.adoQryQueryLogData.Clone(L_AdoQuery,ltUnspecified);

      if L_AdoQuery=nil then
      begin
        sL_CompName := dtmMain.getCompName(sL_CompCode);
        sL_ErrMsg := dtmMain.getLanguageSettingInfo('dtmMain_Msg_4_Content') + '=<' + sL_CompName + '>';
        MessageDlg(sL_ErrMsg,mtError, [mbOK],0);
      end
      else
      begin
        with dtmMain.adoQryQueryLogData do
        begin
          dtmMain.cdsQueryLogData.EmptyDataSet;
          First;
          while not Eof do
          begin
            dtmMain.cdsQueryLogData.Append;
            dtmMain.cdsQueryLogData.FieldByName('DstMsgID').AsString := FieldByName('DstMsgID').AsString;

            
            sL_InstQuery := FieldByName('InstQuery').AsString;
            if sL_InstQuery = '0' then
              sL_InstQuery := dtmMain.getLanguageSettingInfo('frmMain_Msg_12_Content')
            else if sL_InstQuery = '1' then
              sL_InstQuery := dtmMain.getLanguageSettingInfo('frmMain_Msg_13_Content')
            else
              sL_InstQuery := '';
            dtmMain.cdsQueryLogData.FieldByName('InstQuery').AsString := sL_InstQuery;


            sL_ModeType := FieldByName('ModeType').AsString;
            if sL_ModeType = '1' then
              sL_QueryType := IC_CARD_QUERY
            else if sL_ModeType = '2' then
              sL_QueryType := CA_PRODUCT_QUERY
            else if sL_ModeType = '3' then
              sL_QueryType := PROD_PURCHASE_QUERY
            else if sL_ModeType = '4' then
              sL_QueryType := ENTITLEMENT_QUERY
            else if sL_ModeType = '5' then
              sL_QueryType := EXCHANGE_DATE_QUERY;
            dtmMain.cdsQueryLogData.FieldByName('ModeType').AsString := sL_QueryType;


            dtmMain.cdsQueryLogData.FieldByName('DstCode').AsString := FieldByName('DstCode').AsString;
            dtmMain.cdsQueryLogData.FieldByName('SrcCode').AsString := FieldByName('SrcCode').AsString;
            dtmMain.cdsQueryLogData.FieldByName('HandleEDateTime').AsString := TUstr.replaceStr(FieldByName('HandleEDateTime').AsString,'-','/');
            dtmMain.cdsQueryLogData.FieldByName('ExecDateTime').AsString := TUstr.replaceStr(FieldByName('ExecDateTime').AsString,'-','/');
            dtmMain.cdsQueryLogData.FieldByName('SrcMsgID').AsString := FieldByName('SrcMsgID').AsString;
            dtmMain.cdsQueryLogData.FieldByName('ErrorCode').AsString := FieldByName('ErrorCode').AsString;
            dtmMain.cdsQueryLogData.FieldByName('ErrorMsg').AsString := FieldByName('ErrorMsg').AsString;

            dtmMain.cdsQueryLogData.Post;
            Next;
          end;
        end;

        dbgQueryLogData.Columns[0].Title.Caption := dtmMain.getLanguageSettingInfo('frmQueryLog_DstMsgID_Caption');
        dbgQueryLogData.Columns[1].Title.Caption := dtmMain.getLanguageSettingInfo('frmQueryLog_ModeType_Caption');
        dbgQueryLogData.Columns[2].Title.Caption := dtmMain.getLanguageSettingInfo('frmQueryLog_InstQuery_Caption');
        dbgQueryLogData.Columns[3].Title.Caption := dtmMain.getLanguageSettingInfo('frmQueryLog_DstCode_Caption');
        dbgQueryLogData.Columns[4].Title.Caption := dtmMain.getLanguageSettingInfo('frmQueryLog_SrcCode_Caption');
        dbgQueryLogData.Columns[5].Title.Caption := dtmMain.getLanguageSettingInfo('frmQueryLog_HandleEDateTime_Caption');
        dbgQueryLogData.Columns[6].Title.Caption := dtmMain.getLanguageSettingInfo('frmQueryLog_ExecDateTime_Caption');
        dbgQueryLogData.Columns[7].Title.Caption := dtmMain.getLanguageSettingInfo('frmQueryLog_SrcMsgID_Caption');
        dbgQueryLogData.Columns[8].Title.Caption := dtmMain.getLanguageSettingInfo('frmQueryLog_ErrorMsg_Caption');
      end;
    end;
end;

function TfrmQueryLog.IsDataOk: Boolean;
var
    sL_CompCode,sL_StartDate,sL_StopDate,sL_StartTime,sL_StopTime : String;
    sL_QueryType : String;
begin
  sL_CompCode := cobCompCode.Text;
  if sL_CompCode = '' then
  begin
      MessageDlg(dtmMain.getLanguageSettingInfo('frmQueryLog_Msg_1_Content'),mtError, [mbOK],0);
      cobCompCode.SetFocus;
      IsDataOk := false;
      exit;
  end;

  sL_QueryType := cobQueryType.Text;
  if sL_QueryType = '' then
  begin
      MessageDlg(dtmMain.getLanguageSettingInfo('frmQueryLog_Msg_5_Content'),mtError, [mbOK],0);
      cobQueryType.SetFocus;
      IsDataOk := false;
      exit;
  end;

  sL_StartDate := Trim(TUstr.replaceStr(Trim(fraYMD1.getYMD),'/',''));
  sL_StopDate := Trim(TUstr.replaceStr(Trim(fraYMD2.getYMD),'/',''));


  if sL_StartDate = '' then
  begin
    MessageDlg(dtmMain.getLanguageSettingInfo('frmQueryLog_Msg_2_Content'),mtError, [mbOK],0);
    fraYMD1.setYMD('');
    fraYMD1.mseYMD.SetFocus;
    IsDataOk := false;
    exit;
  end;

  if sL_StopDate = '' then
  begin
    MessageDlg(dtmMain.getLanguageSettingInfo('frmQueryLog_Msg_2_Content'),mtError, [mbOK],0);
    fraYMD2.setYMD('');
    fraYMD2.mseYMD.SetFocus;
    IsDataOk := false;
    exit;
  end;

  if StrToInt(sL_StartDate) > StrToInt(sL_StopDate) then
  begin
    MessageDlg(dtmMain.getLanguageSettingInfo('frmQueryLog_Msg_3_Content'),mtError, [mbOK],0);
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
      MessageDlg(dtmMain.getLanguageSettingInfo('frmQueryLog_Msg_4_Content'),mtError, [mbOK],0);
      medETime.Text := '';
      medETime.SetFocus;
      IsDataOk := false;
      exit;
    end;
  end;


end;

procedure TfrmQueryLog.fraYMD1Exit(Sender: TObject);
begin
    fraYMD2.setYMD(fraYMD1.getYMD);
end;

procedure TfrmQueryLog.initialQueryDate;
var
    sL_CurrDate : String;
begin
    sL_CurrDate := TUdateTime.GetPureDateStr(NOW,'/');
    fraYMD1.setYMD(sL_CurrDate);
    medSTime.Text := '0000';

    fraYMD2.setYMD(sL_CurrDate);
    medETime.Text := '2359';
end;

end.

