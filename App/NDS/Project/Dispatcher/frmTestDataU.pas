unit frmTestDataU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, Mask, DBCtrls, ComCtrls, Buttons ,ADODB;

type
  TfrmTestData = class(TForm)
    Label17: TLabel;
    Panel1: TPanel;
    Label4: TLabel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    medSeqNo: TMaskEdit;
    medCompCode: TMaskEdit;
    medCompName: TMaskEdit;
    medCommandID: TMaskEdit;
    medSubscriberID: TMaskEdit;
    medCardID: TMaskEdit;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    medPinCode: TMaskEdit;
    medPopulationID: TMaskEdit;
    medRegionKey: TMaskEdit;
    cobAvailability: TComboBox;
    medReportbackDate: TMaskEdit;
    medCmdStatus: TMaskEdit;
    medNotes: TMaskEdit;
    Label16: TLabel;
    cobPinControl: TComboBox;
    btnSave: TBitBtn;
    Label18: TLabel;
    Label19: TLabel;
    Label20: TLabel;
    Label21: TLabel;
    Label22: TLabel;
    Label23: TLabel;
    Label24: TLabel;
    Label25: TLabel;
    Label26: TLabel;
    Label27: TLabel;
    Label28: TLabel;
    Label29: TLabel;
    btnClose: TBitBtn;
    Label30: TLabel;
    procedure btnSaveClick(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmTestData: TfrmTestData;

implementation

uses frmLoginU, dtmMainU, Ustru;

{$R *.dfm}

procedure TfrmTestData.btnSaveClick(Sender: TObject);
var
    sL_SeqNo,sL_CompCode,sL_CompName,sL_CommandID,sL_SubscriberID : String;
    sL_CardID,sL_PinCode,sL_PopulationID : String;
    sL_RegionKey,sL_ReportbackAvailability,sL_ReportBackDate : String;
    sL_Operator,sL_Notes,sL_FullNotes,sL_Pinontrol,sL_CmdStatus : String;

    dL_SubBeginDate,dL_SubEndDate : Tdate;

    sL_SQL : String;
    L_AdoQry : TADOQuery;
begin
    sL_SeqNo := Trim(medSeqNo.Text);

    sL_CompCode := Trim(medCompCode.Text);
    if sL_CompCode = '' then
      sL_CompCode := 'null';

    sL_CompName := Trim(medCompName.Text);
    sL_CommandID := Trim(medCommandID.Text);
    sL_SubscriberID := Trim(medSubscriberID.Text);
    sL_CardID := Trim(medCardID.Text);

    dL_SubBeginDate := DateTimePicker1.Date;
    dL_SubEndDate := DateTimePicker2.Date;

    sL_PinCode := Trim(medPinCode.Text);
    if sL_PinCode = '' then
      sL_PinCode := 'null';

    sL_PopulationID := Trim(medPopulationID.Text);
    if sL_PopulationID = '' then
      sL_PopulationID := 'null';

    sL_RegionKey := Trim(medRegionKey.Text);

    if cobAvailability.ItemIndex = 0 then
      sL_ReportbackAvailability := 'D'
    else if cobAvailability.ItemIndex = 1 then
      sL_ReportbackAvailability := 'E'
    else if cobAvailability.ItemIndex = 2 then
      sL_ReportbackAvailability := 'A'
    else if cobAvailability.ItemIndex = 3 then
      sL_ReportbackAvailability := 'N'
    else
      sL_ReportbackAvailability := '';


    sL_ReportBackDate := Trim(medReportbackDate.Text);
    if sL_ReportBackDate = '' then
      sL_ReportBackDate := 'null';

    sL_Operator := frmLogin.sG_UserID;

    sL_Notes := Trim(medNotes.Text);

    sL_CmdStatus := Trim(medCmdStatus.Text);

    if cobPinControl.ItemIndex = 0 then
      sL_Pinontrol := 'L'
    else if cobPinControl.ItemIndex = 1 then
      sL_Pinontrol := 'U'
    else
      sL_Pinontrol := '';


    sL_SQL := 'DELETE SEND_NDS WHERE COMPCODE=' + sL_CompCode;
    dtmMain.runSQL(IUD_MODE,sL_CompCode,sL_SQL);

    sL_SQL := 'INSERT INTO SEND_NDS(SEQNO,COMPCODE,COMPNAME,COMMANDID,' +
              'SUBSCRIBERID,ICCNO,SUBBEGINDATE,SUBENDDATE,PINCODE,' +
              'POPULATIONID,REGIONKEY,REPORTBACKAVAILABILITY,REPORTBACKDATE,' +
              'NOTES,CMDSTATUS,OPERATOR,UPDTIME,PINCONTROL) ' +
              'VALUES(''' + sL_SeqNo + ''',' + sL_CompCode + ',''' + sL_CompName +
              ''',''' + sL_CommandID + ''',''' + sL_SubscriberID +
              ''',''' + sL_CardID + ''',' + TUstr.getOracleSQLDateStr(dL_SubBeginDate) +
              ',' + TUstr.getOracleSQLDateStr(dL_SubEndDate) + ',' + sL_PinCode +
              ',' + sL_PopulationID + ',''' + sL_RegionKey + ''',''' + sL_ReportbackAvailability +
              ''',' + sL_ReportBackDate + ',''' + sL_Notes + ''',''' + sL_CmdStatus +
              ''',''' + sL_Operator + ''',' + TUstr.getOracleSQLDateTimeStr(now) +
              ',''' + sL_Pinontrol + ''')';

    L_AdoQry := dtmMain.runSQL(IUD_MODE,sL_CompCode,sL_SQL);

    if L_AdoQry = nil then
    begin
      showmessage('沒有此家公司的連線...');
    end
    else
    begin
      showmessage('儲存完畢');
    end;
end;

procedure TfrmTestData.btnCloseClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmTestData.FormShow(Sender: TObject);
begin
    frmTestData.Caption := '建立測試資料';
end;

end.
