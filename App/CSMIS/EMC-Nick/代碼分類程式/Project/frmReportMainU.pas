unit frmReportMainU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, DB, Grids, DBGrids;


const
    COMP_INFO_FILE_NAME = 'CompInfo.txt';

type
  TfrmReportMain = class(TForm)
    cobTablName: TComboBox;
    btnClose: TButton;
    btnQuery: TButton;
    dsrCodeCD067A: TDataSource;
    StaticText6: TStaticText;
    cobCompInfo: TComboBox;
    StaticText1: TStaticText;
    procedure FormCreate(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure btnQueryClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
    function isDataOk : Boolean;
  public
    { Public declarations }
    procedure loadCompInfo;
  end;

var
  frmReportMain: TfrmReportMain;

implementation

uses dtmMainU, UCommonU, rptShowDataU, frmLoginU, Ustru;

{$R *.dfm}

procedure TfrmReportMain.FormCreate(Sender: TObject);
begin
    TUCommonFun.AddObjectToComboBox(cobTablName.Items, dsrCodeCD067A.DataSet,INSERT_NO_DATA_ITEM);
    //預設為全部
    TUCommonFun.setComboDefaultNdx(cobTablName, '0');

    loadCompInfo;
end;

procedure TfrmReportMain.btnCloseClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmReportMain.btnQueryClick(Sender: TObject);
var
    sL_TableName,sL_CompCode,sL_Temp,sL_CompName : String;
    L_Rpt : TrptShowData;
    L_StrList : TStringList;
begin
    if isDataOk then
    begin
      L_StrList := TStringList.Create;
      sL_Temp := Trim(cobCompInfo.Text);
      L_StrList := TUstr.ParseStrings(sL_Temp,'-');
      sL_CompCode := L_StrList.Strings[0];
      sL_CompName := L_StrList.Strings[1];

      sL_TableName := dtmMain.getCurCobRecID(cobTablName);
      dtmMain.getRptData(sL_TableName,sL_CompCode);

      if dtmMain.adoRptData.RecordCount=0 then
        MessageDlg('查無資料,請重新查詢',mtInformation, [mbOK],0)
      else
      begin
        L_Rpt := TrptShowData.Create(Application);
        //L_Rpt.sG_Compute := dtmMain.sG_CompCode;
        //L_Rpt.sG_CompName := sL_CompName;
        L_Rpt.sG_User := dtmMain.sG_USerID;
        L_Rpt.Preview;

        L_Rpt.Free;
        L_Rpt := nil;
      end;
    end;

    L_StrList.Free;
end;

function TfrmReportMain.isDataOk: Boolean;
var
    sL_TableName : String;
begin
    sL_TableName := cobTablName.Text;
    if sL_TableName = '' then
    begin
      MessageDlg('請選擇 Table Name',mtError, [mbOK],0);
      cobTablName.SetFocus;
      Result := false;
      exit;
    end;

    Result := true;
end;

procedure TfrmReportMain.loadCompInfo;
var
    ii : Integer;
    sL_FileName : String;
    L_TmpStrList, L_TmpStrList1 : TStringList;
    sL_ExeFileName, sL_ExePath : STring;
    sL_CompCode, sL_CompName : String;


begin

    sL_ExeFileName := Application.ExeName;
    sL_ExePath := ExtractFileDir(sL_ExeFileName);

    sL_FileName := sL_ExePath + '\' + COMP_INFO_FILE_NAME;

    if FileExists(sL_FileName) then
    begin
      if dtmMain.cdsCompInfo.State = dsInactive then
        dtmMain.cdsCompInfo.CreateDataSet;
      cobCompInfo.Items.Clear;
      cobCompInfo.Items.Add(NO_FILTER_FLAG + '-全部');
      cobCompInfo.ItemIndex := 0;
      L_TmpStrList := TStringList.Create;
      L_TmpStrList.LoadFromFile(sL_FileName);
      for ii:=0 to L_TmpStrList.Count -1 do
      begin
        if L_TmpStrList.Strings[ii] <>'' then
        begin
          L_TmpStrList1 := TUstr.ParseStrings(L_TmpStrList.Strings[ii],'=');
          if L_TmpStrList1.Count=2 then
          begin
            sL_CompCode := L_TmpStrList1.Strings[0];
            sL_CompName := L_TmpStrList1.Strings[1];
            cobCompInfo.Items.Add(sL_CompCode + '-' + sL_CompName);

            dtmMain.cdsCompInfo.Append;
            dtmMain.cdsCompInfo.FieldByName('CompCode').AsInteger := StrToInt(sL_CompCode);
            dtmMain.cdsCompInfo.FieldByName('CompName').AsString := sL_CompCode + '-' + sL_CompName;
            dtmMain.cdsCompInfo.FieldByName('Index').AsInteger := ii+1;
            dtmMain.cdsCompInfo.Post;

          end;
          if Assigned(L_TmpStrList1) then
            L_TmpStrList1.Free;

        end;
      end;

      L_TmpStrList.Free;
    end
    else
    begin
      MessageDlg('找不到公司資訊檔(' + sL_FileName + ')程式即將結束.', mtError, [mbOK],0);
      Application.Terminate;
    end;

end;

procedure TfrmReportMain.FormShow(Sender: TObject);
var
    nL_CompCodeIndex : Integer;
begin
    self.Caption := dtmMain.setCaption('SO8F20','[二階代碼分類明細對照表查詢]');

    if dtmMain.sG_NeedLogin='N' then
    begin
      if dtmMain.cdsCompInfo.Locate('CompCode', dtmMain.sG_CompCode, [loPartialKey]) then
      begin
        nL_CompCodeIndex := dtmMain.cdsCompInfo.FieldByName('Index').AsInteger;
        cobCompInfo.ItemIndex := nL_CompCodeIndex;
        cobCompInfo.Enabled := false;
      end;
    end;
end;

end.
