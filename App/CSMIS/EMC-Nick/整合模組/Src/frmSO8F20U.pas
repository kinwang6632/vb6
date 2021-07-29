unit frmSO8F20U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, DB, Grids, DBGrids;


const
    COMP_INFO_FILE_NAME = 'CompInfo.txt';

type
  TfrmSO8F20 = class(TForm)
    cobTablName: TComboBox;
    btnClose: TButton;
    btnQuery: TButton;
    dsrCodeCD067A: TDataSource;
    StaticText6: TStaticText;
    cobCompInfo: TComboBox;
    StaticText1: TStaticText;
    Label1: TLabel;
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
  frmSO8F20: TfrmSO8F20;

implementation

uses dtmMain4U, UCommonU, rptSO8F20U, frmLoginU, Ustru, frmMainMenuU,
  frmSO8F10_4U;

{$R *.dfm}

procedure TfrmSO8F20.FormCreate(Sender: TObject);
begin
    Application.CreateForm(TfrmSO8F10_4,frmSO8F10_4);
    frmSO8F10_4.ShowModal;

    dtmMain4.connectToCommDB(frmSO8F10_4.bG_IsSuperUser);


    TUCommonFun.AddObjectToComboBox(cobTablName.Items, dsrCodeCD067A.DataSet,INSERT_NO_DATA_ITEM,true,'TableName','TableDescription');
    //預設為全部
    TUCommonFun.setComboDefaultNdx(cobTablName, '999');

    loadCompInfo;
end;

procedure TfrmSO8F20.btnCloseClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmSO8F20.btnQueryClick(Sender: TObject);
var
    sL_TableName,sL_CompCode,sL_Temp,sL_CompName : String;
    L_Rpt : TrptSO8F20;
    L_StrList : TStringList;
begin
    if isDataOk then
    begin
      //執行等待狀態
      TUCommonFun.setWaitingCursor;

      L_StrList := TStringList.Create;
      sL_Temp := Trim(cobCompInfo.Text);
      L_StrList := TUstr.ParseStrings(sL_Temp,'-');
      sL_CompCode := L_StrList.Strings[0];
      sL_CompName := L_StrList.Strings[1];

      sL_TableName := dtmMain4.getCurCobRecID(cobTablName);
      dtmMain4.getRptData(sL_TableName,sL_CompCode);

      if dtmMain4.adoRptData.RecordCount=0 then
      begin
        //回復原狀態
        TUCommonFun.setDefaultCursor;
        
        MessageDlg('查無資料,請重新查詢',mtInformation, [mbOK],0)
      end
      else
      begin
        //回復原狀態
        TUCommonFun.setDefaultCursor;

        L_Rpt := TrptSO8F20.Create(Application);
        //L_Rpt.sG_Compute := frmMainMenu.sG_CompCode;
        //L_Rpt.sG_CompName := sL_CompName;
        L_Rpt.sG_User := frmMainMenu.sG_USerID;
        L_Rpt.Preview;

        L_Rpt.Free;
        L_Rpt := nil;
      end;
    end;

    L_StrList.Free;
end;

function TfrmSO8F20.isDataOk: Boolean;
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

procedure TfrmSO8F20.loadCompInfo;
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
      if dtmMain4.cdsCompInfo.State = dsInactive then
        dtmMain4.cdsCompInfo.CreateDataSet;
      cobCompInfo.Items.Clear;
      cobCompInfo.Items.Add(ALL_COMPCODE + '-所有');
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

            dtmMain4.cdsCompInfo.Append;
            dtmMain4.cdsCompInfo.FieldByName('CompCode').AsInteger := StrToInt(sL_CompCode);
            dtmMain4.cdsCompInfo.FieldByName('CompName').AsString := sL_CompCode + '-' + sL_CompName;
            dtmMain4.cdsCompInfo.FieldByName('Index').AsInteger := ii+1;
            dtmMain4.cdsCompInfo.Post;

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

procedure TfrmSO8F20.FormShow(Sender: TObject);
var
    nL_CompCodeIndex : Integer;
begin
    self.Caption := frmMainMenu.setCaption('SO8F20','[二階代碼分類明細對照表查詢]');

    if frmMainMenu.sG_NeedLogin='N' then
    begin
      if dtmMain4.cdsCompInfo.Locate('CompCode', frmMainMenu.sG_CompCode, [loPartialKey]) then
      begin
        nL_CompCodeIndex := dtmMain4.cdsCompInfo.FieldByName('Index').AsInteger;
        cobCompInfo.ItemIndex := nL_CompCodeIndex;
        cobCompInfo.Enabled := false;
      end;
    end;

    if frmSO8F10_4.bG_IsSuperUser then
    begin
      Label1.Caption := '共用資料區';
      cobCompInfo.Enabled := true;
    end
    else
    begin
      Label1.Caption := frmMainMenu.sG_CompName;
      cobCompInfo.Enabled := false;
    end;    
end;

end.
