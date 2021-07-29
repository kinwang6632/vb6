unit frmSO8B40U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, fraChineseYMU;

type
  TfrmSO8B40 = class(TForm)
    Label1: TLabel;
    Label2: TLabel;
    stxCom: TStaticText;
    fraChineseYM1: TfraChineseYM;
    btnLock: TButton;
    btnClose: TButton;
    Label3: TLabel;
    lblBelongYM: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure btnLockClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
    function IsDataOk : Boolean;
  public
    { Public declarations }
  end;

var
  frmSO8B40: TfrmSO8B40;

implementation

uses dtmMain1U, Ustru, frmMainMenuU;

{$R *.dfm}

procedure TfrmSO8B40.FormCreate(Sender: TObject);
var
    sL_LockYM,sL_Year,sL_Month : String;
    sL_OldLockYM : String;
begin
    stxCom.Caption := frmMainMenu.sG_CompName;


    sL_OldLockYM := dtmMain1.getLockYM;
    sL_Year := Copy(sL_OldLockYM,1,4);
    sL_Month := Copy(sL_OldLockYM,5,2);
    sL_LockYM :=  sL_Year + '/' + sL_Month;

    lblBelongYM.Caption := sL_LockYM;

    fraChineseYM1.setYM(sL_LockYM);
end;

procedure TfrmSO8B40.btnCloseClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmSO8B40.btnLockClick(Sender: TObject);
var
    sL_NewLockYM,sL_OldLockYM : String;
    sL_Year,sL_Month,sL_LockYM : String;
    aROC: String;
begin
    if IsDataOk then
    begin
      sL_NewLockYM := Trim(TUstr.replaceStr(fraChineseYM1.getYM,'/',''));

      sL_OldLockYM := dtmMain1.getLockYM;
      if sL_OldLockYM='' then
        dtmMain1.InsertLockYM(sL_NewLockYM)
      else begin
        dtmMain1.updateLockYM(sL_NewLockYM);
      end;  

      SHOWMESSAGE('鎖帳成功');

      sL_Year := Copy(sL_NewLockYM,1,4);
      sL_Month := Copy(sL_NewLockYM,5,2);
      sL_LockYM :=  sL_Year + '/' + sL_Month;

      lblBelongYM.Caption := sL_LockYM;

    end;
end;

function TfrmSO8B40.IsDataOk: Boolean;
var
    sL_NewLockYM,sL_ErrMsg,sL_OldLockYM,sL_NextLockYM : String;
    nL_Length : Integer;
begin
    sL_NewLockYM := Trim(TUstr.replaceStr(fraChineseYM1.getYM,'/',''));

    if sL_NewLockYM = '' then
    begin
      MessageDlg('請輸入鎖帳年月',mtError, [mbOK],0);
      fraChineseYM1.mseYM.SetFocus;
      Result := false;
      exit;
    end;

    sL_OldLockYM := dtmMain1.getLockYM;
    if sL_NewLockYM < sL_OldLockYM then
    begin
      sL_ErrMsg := '新的鎖帳年月,不能小於原鎖帳年月';
      MessageDlg(sL_ErrMsg,mtError, [mbOK],0);
      fraChineseYM1.mseYM.SetFocus;
      fraChineseYM1.setYM('');
      Result := false;
      exit;
    end;

    if sL_OldLockYM <> '' then
    begin
      sL_NextLockYM := dtmMain1.afterMonth(sL_OldLockYM,1);
      nL_Length := Length(sL_NextLockYM);
      if nL_Length = 4 then
        sL_NextLockYM := '0' + sL_NextLockYM;

      if sL_NewLockYM > sL_NextLockYM then
      begin
        sL_ErrMsg := '鎖帳年月,不能跨月鎖帳';
        MessageDlg(sL_ErrMsg,mtError, [mbOK],0);
        fraChineseYM1.setYM('');
        fraChineseYM1.mseYM.SetFocus;
        Result := false;
        exit;
      end;
    end;


    Result := true;
end;

procedure TfrmSO8B40.FormShow(Sender: TObject);
begin
    self.Caption:=frmMainMenu.setCaption('SO8B40','[鎖帳]');
end;

end.
