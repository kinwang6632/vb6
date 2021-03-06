unit frmInvE03U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ExtCtrls, Mask, Menus, cxLookAndFeelPainters,
  cxButtons, cxControls, cxContainer, cxEdit, cxCheckBox, cxTextEdit,
  cxMaskEdit, cxGroupBox, dxSkinsCore, dxSkinsDefaultPainters;

type
  TfrmInvE03 = class(TForm)
    Panel1: TPanel;
    Label1: TLabel;
    Panel2: TPanel;
    Label2: TLabel;
    Label3: TLabel;
    btnOk: TcxButton;
    btnClose: TcxButton;
    cxGroupBox1: TcxGroupBox;
    chkSynCompData: TcxCheckBox;
    chkSynChargeItem: TcxCheckBox;
    chkSynPayType: TcxCheckBox;
    chkSynServiceType: TcxCheckBox;
    chkSynCustID: TcxCheckBox;
    medCustID: TcxMaskEdit;
    chkSyncCD095: TcxCheckBox;
    procedure FormShow(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
  private
    { Private declarations }
    procedure setButtonCompetence;
    function isDataOK : Boolean;
  public
    { Public declarations }
  end;

var
  frmInvE03: TfrmInvE03;

implementation

uses cbUtilis, frmMainU, dtmMainU, dtmMainJU;

{$R *.dfm}

procedure TfrmInvE03.FormShow(Sender: TObject);
begin

    self.Caption := frmMain.GetFormTitleString( 'E03', '資料同步' );
    setButtonCompetence;
    Label3.Visible := False;
    if not dtmMain.GetLinkToMIS then
    begin
      chkSynCompData.Enabled := False;
      chkSynChargeItem.Enabled := False;
      chkSynPayType.Enabled := False;
      chkSynServiceType.Enabled := False;
      chkSynCustID.Enabled := False;
      medCustID.Enabled := False;
      Label3.Visible := True;
    end;      

end;

procedure TfrmInvE03.setButtonCompetence;
var
    sL_QueryCompetence,sL_InsertCompetence,sL_DeleteCompetence,sL_UpdateCompetence,sL_ExecuteCompetence : String;
begin
     //設定權限
     dtmMain.getCompetence('E03',sL_QueryCompetence,sL_InsertCompetence,sL_DeleteCompetence,sL_UpdateCompetence,sL_ExecuteCompetence);

     if sL_ExecuteCompetence = 'Y' then
       btnOK.Enabled := true
     else
       btnOK.Enabled := false;
end;

procedure TfrmInvE03.btnExitClick(Sender: TObject);
begin
  Close;
end;

function TfrmInvE03.isDataOK: Boolean;
var
    sL_CustID : String;
begin
    Result := true;
    
    if (not chkSynCompData.Checked) and
       (not chkSynChargeItem.Checked) and
       (not chkSynPayType.Checked) and
       (not chkSynServiceType.Checked) and
       (not chkSynCustID.Checked) and
       (not chkSyncCD095.Checked) then
    begin
      if dtmMain.GetLinkToMIS then
      begin
        WarningMsg('請至少勾選一項要同步的資料!');
        chkSynCompData.SetFocus;
      end;  
      Result := false;
      Exit;
    end;


    if chkSynCustID.Checked then
    begin
      sL_CustID := Trim(medCustID.Text);

      if sL_CustID = '' then
      begin
        WarningMsg( '請輸入要同步的客戶代碼!' );
        medCustID.SetFocus;
        Result := false;
        Exit;
      end;
    end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvE03.btnOKClick(Sender: TObject);
var
  sL_MisDbOwner,sL_CurDateTime,sL_UserName,sL_CustID,sL_CompID,sL_MisDbCompCode : String;
begin
  if isDataOK then
  begin
    sL_MisDbOwner := dtmMain.getMisDbOwner;
    sL_CurDateTime := DateTimeToStr(now);
    sL_UserName := dtmMain.getLoginUser;
    sL_CompID := dtmMain.getCompID;
    sL_MisDbCompCode := dtmMain.getMisDbCompCode;
    Screen.Cursor := crSQLWait;
    try
      //同步公司資料
      if chkSynCompData.Checked then
      begin
        dtmMainJ.synchronizeCompData(sL_MisDbOwner,sL_CurDateTime,sL_UserName);
      end;
      //同步收費項目
      if chkSynChargeItem.Checked then
      begin
        dtmMainJ.synchronizeChargeItem2;
      end;
      //同步付款方式
      if chkSynPayType.Checked then
      begin
        dtmMainJ.synchronizePayType(sL_MisDbOwner,sL_CurDateTime,sL_UserName);
      end;
      //同步業者服務種類
      if chkSynServiceType.Checked then
      begin
        dtmMainJ.synchronizeServiceType(sL_MisDbOwner,sL_CurDateTime,sL_UserName);
      end;
      //同步客戶代碼資料
      if chkSynCustID.Checked then
      begin
        sL_CustID := Trim(medCustID.Text);
        dtmMainJ.synchronizeCustID( sL_MisDbOwner, sL_CurDateTime, sL_UserName,
          sL_CompID, sL_CustID, sL_MisDbCompCode );
      end;
      //發票用途代那碼
      if chkSyncCD095.Checked then
      begin
        dtmMainJ.SynchronizeCD095;
      end;
      InfoMsg( '資料同步完成。' );
    finally
      Screen.Cursor := crDefault;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }


end.
