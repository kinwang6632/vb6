unit frmInvE03U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ExtCtrls, Mask, Menus, cxLookAndFeelPainters,
  cxButtons, cxControls, cxContainer, cxEdit, cxCheckBox, cxTextEdit,
  cxMaskEdit, cxGroupBox, dxSkinsCore, dxSkinsDefaultPainters, dxSkinBlack,
  dxSkinBlue, dxSkinCaramel, dxSkinCoffee;

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
    chkSyncCD110: TcxCheckBox;
    chkSyncCD122: TcxCheckBox;
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

    self.Caption := frmMain.GetFormTitleString( 'E03', '��ƦP�B' );
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
    sL_VerifyCompetence : String;
begin
     //�]�w�v��
     dtmMain.getCompetence('E03',sL_QueryCompetence,sL_InsertCompetence,sL_DeleteCompetence,
        sL_UpdateCompetence,sL_ExecuteCompetence,sL_VerifyCompetence);

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
       (not chkSyncCD095.Checked) and
       (not chkSyncCD110.Checked) and
       (not chkSyncCD122.Checked) then
    begin
      if dtmMain.GetLinkToMIS then
      begin
        WarningMsg('�Цܤ֤Ŀ�@���n�P�B�����!');
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
        WarningMsg( '�п�J�n�P�B���Ȥ�N�X!' );
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
      //�P�B���q���
      if chkSynCompData.Checked then
      begin
        dtmMainJ.synchronizeCompData(sL_MisDbOwner,sL_CurDateTime,sL_UserName);
      end;
      //�P�B���O����
      if chkSynChargeItem.Checked then
      begin
        dtmMainJ.synchronizeChargeItem2;
      end;
      //�P�B�I�ڤ覡
      if chkSynPayType.Checked then
      begin
        dtmMainJ.synchronizePayType(sL_MisDbOwner,sL_CurDateTime,sL_UserName);
      end;
      //�P�B�~�̪A�Ⱥ���
      if chkSynServiceType.Checked then
      begin
        dtmMainJ.synchronizeServiceType(sL_MisDbOwner,sL_CurDateTime,sL_UserName);
      end;
      //�P�B�Ȥ�N�X���
      if chkSynCustID.Checked then
      begin
        sL_CustID := Trim(medCustID.Text);
        dtmMainJ.synchronizeCustID( sL_MisDbOwner, sL_CurDateTime, sL_UserName,
          sL_CompID, sL_CustID, sL_MisDbCompCode );
      end;
      //�o���γ~�N���X
      if chkSyncCD095.Checked then
      begin
        dtmMainJ.SynchronizeCD095;
      end;
      if chkSyncCD110.Checked then
      begin
        dtmMainJ.SynchronizeCD110;
      end;
      {#6371 �W�[�Ȥ�������O By Kin 2013/02/25}
      if chkSyncCD122.Checked then
      begin
        dtmMainJ.SynchronizeCD122;
      end;
      InfoMsg( '��ƦP�B�����C' );
    finally
      Screen.Cursor := crDefault;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }


end.