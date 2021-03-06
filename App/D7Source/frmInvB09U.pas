unit frmInvB09U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Mask, ExtCtrls, dxSkinsCore, dxSkinBlack, dxSkinBlue,
  dxSkinCaramel, dxSkinCoffee, dxSkinsDefaultPainters, cxMaskEdit,
  cxControls, cxContainer, cxEdit, cxTextEdit, cxDropDownEdit, cxCalendar,
  cxLabel, Menus, cxLookAndFeelPainters, cxButtons, cxCheckBox;

type
  TfrmInvB09 = class(TForm)
    pnlMain: TPanel;
    mskInvId1: TcxMaskEdit;
    mskInvId2: TcxMaskEdit;
    cxLabel1: TcxLabel;
    cxLabel2: TcxLabel;
    cxLabel3: TcxLabel;
    cxLabel4: TcxLabel;
    btnQuery: TcxButton;
    btnExit: TcxButton;
    chkDiscount: TcxCheckBox;
    txtInvDate1: TcxDateEdit;
    txtInvDate2: TcxDateEdit;
    grpDiscount: TGroupBox;
    txtDiscDate1: TcxDateEdit;
    cxLabel5: TcxLabel;
    txtDiscDate2: TcxDateEdit;
    cxLabel6: TcxLabel;
    cxLabel7: TcxLabel;
    txtDiscNoSt: TcxMaskEdit;
    cxLabel8: TcxLabel;
    txtDiscNoEd: TcxMaskEdit;
    procedure btnQueryClick(Sender: TObject);
    procedure mskInvId1PropertiesValidate(Sender: TObject;
      var DisplayValue: Variant; var ErrorText: TCaption;
      var Error: Boolean);
    procedure btnExitClick(Sender: TObject);
    procedure chkDiscountPropertiesEditValueChanged(Sender: TObject);
  private
    { Private declarations }
    function isDataOK : Boolean;
  public
    { Public declarations }
  end;

var
  frmInvB09: TfrmInvB09;

implementation

uses cbUtilis, frmMainU, dtmMainU, dtmMainJU, frmMultiSelectU;
{$R *.dfm}

{ TfrmInvB09 }

function TfrmInvB09.isDataOK: Boolean;
 var sL_SInvDate ,sL_EInvDate, sL_SInvID, sL_EInvID : String;
begin
  Result := False;
  sL_SInvDate := Trim( txtInvDate1.EditText );
  sL_EInvDate := Trim( txtInvDate2.EditText );
  if not chkDiscount.Checked then
  begin
    if sL_SInvDate = '' then
    begin
      txtInvDate1.SetFocus;
      WarningMsg( '請輸入發票起始日期！');
      Exit;
    end;
    if sL_EInvDate = '' then
    begin
      txtInvDate2.SetFocus;
      WarningMsg( '請輸入發票終止日期！');
      Exit;
    end;
    if  ( StrToDate(sL_SInvDate) ) > ( StrToDate( sL_EInvDate) ) then
    begin
      txtInvDate1.SetFocus;
      WarningMsg( '起始日期大於終止日期！');
      Exit;
    end;
  end else
  begin
    if txtDiscDate1.Text = EmptyStr then
    begin
      txtDiscDate1.SetFocus;
      WarningMsg( '請輸入折讓單起始日期！');
      Exit;
    end;
    if txtDiscDate2.Text = EmptyStr then
    begin
      txtDiscDate2.SetFocus;
      WarningMsg( '請輸入折讓單終止日期！');
      Exit;
    end;
  end;
  Result := True;
end;

procedure TfrmInvB09.btnQueryClick(Sender: TObject);
var aMsg : String;
begin
  if not isDataOK then Exit;
  Screen.Cursor := crSQLWait;
  try
      if not chkDiscount.Checked then
      begin


        aMsg := dtmMainJ.AuthorizeElectronInv(txtInvDate1.EditText,
                    txtInvDate2.EditText,
                    mskInvId1.EditText,
                    mskInvId2.EditText,False);
      end else
      begin
        aMsg := dtmMainJ.AuthorizeDiscountInv( txtDiscDate1.EditText,
              txtDiscDate2.EditText,txtDiscNoSt.EditText,
              txtDiscNoEd.EditText, False );
      end;

  finally
    MessageDlg( aMsg,mtInformation,[mbOK],0);
    Screen.Cursor := crDefault;
  end;
end;

procedure TfrmInvB09.mskInvId1PropertiesValidate(Sender: TObject;
  var DisplayValue: Variant; var ErrorText: TCaption; var Error: Boolean);
begin
  if (Sender as TcxMaskEdit).Text = EmptyStr then Exit;
  if Length((Sender as TcxMaskEdit).Text) <> 10 then
  begin
    ErrorText := '發票號碼長度錯誤！';
    Error := False;
    WarningMsg( ErrorText );
    (Sender as TcxMaskEdit).SetFocus;
  end;
end;

procedure TfrmInvB09.btnExitClick(Sender: TObject);
begin
  Self.Close;
end;

procedure TfrmInvB09.chkDiscountPropertiesEditValueChanged(
  Sender: TObject);
begin
  txtInvDate1.Clear;
  txtInvDate2.Clear;
  txtDiscDate1.Clear;
  txtDiscDate2.Clear;
  txtDiscNoSt.Clear;
  txtDiscNoEd.Clear;
  mskInvId1.Clear;
  mskInvId2.Clear;
  txtInvDate1.Enabled := False;
  txtInvDate2.Enabled := False;
  txtDiscDate1.Enabled := False;
  txtDiscDate2.Enabled := False;
  grpDiscount.Enabled := False;
  mskInvId1.Enabled := False;
  mskInvId2.Enabled := False;

  if not chkDiscount.Checked then
  begin
    txtInvDate1.Enabled := True;
    txtInvDate2.Enabled := True;
    mskInvId1.Enabled := True;
    mskInvId2.Enabled :=True;
    txtInvDate1.SetFocus;
  end else
  begin
    txtDiscDate1.Enabled := True;
    txtDiscDate2.Enabled := True;
    txtDiscNoSt.Enabled := True;
    txtDiscNoEd.Enabled := True;
    grpDiscount.Enabled := True;
    txtDiscDate1.SetFocus;
  end;
end;

end.
