unit frmSyncU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, inifiles, Buttons, ExtCtrls, DB, ADODB, Grids, DBGrids;

type

  TfrmSync = class(TForm)
    pnlMain: TPanel;
    Panel1: TPanel;
    btnConnect: TBitBtn;
    cobDataArea: TComboBox;
    StaticText1: TStaticText;
    BitBtn1: TBitBtn;
    GroupBox1: TGroupBox;
    chb1: TCheckBox;
    BitBtn2: TBitBtn;
    procedure FormShow(Sender: TObject);
    procedure btnConnectClick(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
  private
    { Private declarations }
    sG_DbSID, sG_UserID, sG_Password, sG_CompCode : String;

    procedure setDataAreaCombo;

  public
    { Public declarations }
  end;

var
  frmSync: TfrmSync;

implementation

uses ConstU, Encryption_TLB, dtmMainU;

{$R *.dfm}

{ TfrmSync }

procedure TfrmSync.setDataAreaCombo;
var
    L_StrList : TStringList;
    sL_ControlCompCode, sL_ControlCompName: String;
begin
    L_StrList := TStringList.Create;
    dtmMain.getIniData(L_StrList,sL_ControlCompCode,sL_ControlCompName);
    cobDataArea.Items := L_StrList;
    cobDataArea.ItemIndex := 0;
end;

procedure TfrmSync.FormShow(Sender: TObject);
begin
    setDataAreaCombo;
    cobDataArea.SetFocus;

end;


procedure TfrmSync.btnConnectClick(Sender: TObject);
var
    sL_DataArea, sL_CompName : String;
begin
    screen.Cursor := crHourGlass;
    pnlMain.Enabled := false;
    sL_CompName := cobDataArea.Items.Strings[cobDataArea.ItemIndex];
    sL_DataArea := (cobDataArea.Items.Objects[cobDataArea.ItemIndex] as TDataAreaObj).s_DataAreaID;
    
    try
      dtmMain.connectToDB(sL_DataArea, sG_DbSID, sG_UserID, sG_Password, sG_CompCode);
    except

      sG_DbSID := '';
      screen.Cursor := crDefault;
      pnlMain.Enabled := true;      
      MessageDlg('系統台 ' + sL_CompName + ' 連結失敗!', mtError, [mbOK],0);
      Exit;
    end;
    screen.Cursor := crDefault;
    pnlMain.Enabled := true;    
    MessageDlg('系統台 ' + sL_CompName + ' 連結成功!', mtInformation, [mbOK],0);

end;


procedure TfrmSync.BitBtn1Click(Sender: TObject);
var
    sL_SQL : String;
begin
    if not dtmMain.ADOConnection1.Connected then
    begin
      MessageDlg('請先連接到資料來源!', mtWarning, [mbOK],0);
      cobDataArea.SetFocus;
    end
    else
    begin
      if chb1.checked then
      begin
        sL_SQL := 'select DESCRIPTION, CHANNELID from cd024';
        if not dtmMain.transData(sL_SQL, CHANNEL_DATA_FILENAME , dtmMain.cdsCD024) then
        begin
          MessageDlg('頻道資訊同步失敗!', mtError, [mbOK],0);
          Exit;
        end;
      end;
    end;
end;

procedure TfrmSync.BitBtn2Click(Sender: TObject);
begin
    Close;
end;

end.
