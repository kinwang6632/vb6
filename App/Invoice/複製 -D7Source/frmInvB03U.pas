unit frmInvB03U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ExtCtrls, Mask, cxControls, cxContainer,
  cxEdit, cxLabel, dxSkinsCore, dxSkinsDefaultPainters,
  cxTextEdit, cxMaskEdit, cxButtonEdit, cxGraphics, cxLookAndFeels,
  cxLookAndFeelPainters, dxSkinBlack, dxSkinBlue, dxSkinCaramel,
  dxSkinCoffee;

type
  TfrmInvB03 = class(TForm)
    Panel1: TPanel;
    Label1: TLabel;
    Panel2: TPanel;
    btnOK: TBitBtn;
    btnExit: TBitBtn;
    Label3: TLabel;
    medSInvID: TMaskEdit;
    medEInvID: TMaskEdit;
    lblMsg: TcxLabel;
    chkInvUse: TCheckBox;
    Label5: TLabel;
    txtInvUseItem: TcxButtonEdit;
    Label2: TLabel;
    chkEInvoice: TCheckBox;
    grpImport: TGroupBox;
    edtImport: TEdit;
    btnImport: TButton;
    OpenDialog1: TOpenDialog;
    procedure FormShow(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure medSInvIDExit(Sender: TObject);
    procedure medEInvIDExit(Sender: TObject);
    procedure medSInvIDChange(Sender: TObject);
    procedure chkInvUseClick(Sender: TObject);
    procedure txtInvUseItemPropertiesButtonClick(Sender: TObject;
      AButtonIndex: Integer);
    procedure btnImportClick(Sender: TObject);
  private
    { Private declarations }
    FImportInv : String;
    procedure setButtonCompetence;
    function isDataOK : Boolean;
    procedure DoSelectInvUseItem;
  public
    { Public declarations }
  end;

var
  frmInvB03: TfrmInvB03;

implementation

uses cbUtilis, frmMainU, dtmMainU, dtmMainJU, frmMultiSelectU;

var
  aQueryParam: TQueryParam;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB03.FormShow(Sender: TObject);
begin
  Self.Caption := frmMain.GetFormTitleString( 'B03', '發票資料修改授權作業' );
  setButtonCompetence;
  lblMsg.Caption := '';
  chkEInvoice.Enabled:=dtmMain.GetStarEInvoice;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB03.setButtonCompetence;
var
    sL_QueryCompetence,sL_InsertCompetence,sL_DeleteCompetence,sL_UpdateCompetence,sL_ExecuteCompetence : String;
    sL_VerifyCompetence : String;
begin
  //設定權限
  dtmMain.getCompetence( 'B03', sL_QueryCompetence, sL_InsertCompetence,
    sL_DeleteCompetence, sL_UpdateCompetence, sL_ExecuteCompetence,sL_VerifyCompetence );
  btnOK.Enabled := ( sL_ExecuteCompetence = 'Y' );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB03.btnExitClick(Sender: TObject);
begin
  Close;
end;

{ ---------------------------------------------------------------------------- }

function TfrmInvB03.isDataOK: Boolean;
var
    sL_SInvID,sL_EInvID : String;
begin
  Result := False;
  
    sL_SInvID := Trim(medSInvID.Text);
    sL_EInvID := Trim(medEInvID.Text);
    if ( sL_SInvID = '' ) and ( sL_EInvID = '' ) then
    begin
      if FImportInv = '' then
      begin
        WarningMsg( '請輸入發票號碼!' );
        medSInvID.SetFocus;
        Exit;
      end;
    end;

    if FImportInv = '' then
    begin
      if Length(sL_SInvID) <> 10 then
      begin
        WarningMsg( '輸入發票號碼錯誤' );
        medSInvID.SetFocus;
        Exit;
      end;

      if Length(sL_EInvID) <> 10 then
      begin
        WarningMsg( '輸入發票號碼錯誤' );
        medEInvID.SetFocus;
        Exit;
      end;


      if Copy( sL_SInvID, 1, 2 ) <> Copy( sL_EInvID, 1, 2 ) then
      begin
        WarningMsg( '必須同一發票字軌' );
        medEInvID.SetFocus;
        Exit;
      end;

      if Copy(sL_SInvID, 3, 8 ) > Copy( sL_EInvID, 3, 8 ) then
      begin
        WarningMsg( '發票號碼輸入順序錯誤!' );
        medEInvID.SetFocus;
        Exit;
      end;

      if StrToIntDef( Copy(sL_SInvID,3,8), -1 ) = -1 then
      begin
        WarningMsg( '發票號碼格式錯誤!' );
        medSInvID.SetFocus;
        Exit;
      end;

      if StrToIntDef( Copy( sL_EInvID, 3, 8 ), -1 ) = -1 then
      begin
        WarningMsg( '發票號碼格式錯誤!' );
        Exit;
      end;
    end;
    

    Result := True;

end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB03.btnOKClick(Sender: TObject);
var
  aInvIdSt, aInvIdEd, aMsg: String;
  aInvKind: Integer;
  aInvs : String;
begin
  if not IsDataOK then Exit;
  aInvIdSt := Trim(medSInvID.Text);
  aInvIdEd := Trim(medEInvID.Text);
  Screen.Cursor := crSQLWait;
  {#5668 增加發票種類 By Kin 2010/06/22}

  try
    aInvKind := 0;
    if chkEInvoice.Checked then
      aInvKind := 1;

    aMsg := dtmMainJ.AuthorizeInv( aInvIdSt, aInvIdEd, aQueryParam.Param1,
      chkInvUse.Checked,aInvKind,FImportInv );
  finally
    FImportInv :='';
    Screen.Cursor := crDefault;
  end;
  lblMsg.Caption := aMsg;
  
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB03.medSInvIDExit(Sender: TObject);
begin
  medSInvID.Text := UpperCase(Trim(medSInvID.Text));
  if Trim( medEInvID.Text ) = EmptyStr then
    medEInvID.Text := UpperCase(Trim(medSInvID.Text));
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB03.medEInvIDExit(Sender: TObject);
begin
    medEInvID.Text := UpperCase(Trim(medEInvID.Text));
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB03.medSInvIDChange(Sender: TObject);
begin
   lblMsg.Caption := '';
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB03.chkInvUseClick(Sender: TObject);
begin
  if ( chkInvUse.Checked ) then
  begin
    txtInvUseItem.Clear;
    txtInvUseItem.Enabled := False;
    aQueryParam.Param1 := EmptyStr;
  end else
  begin
    txtInvUseItem.Clear;
    txtInvUseItem.Enabled := True;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB03.txtInvUseItemPropertiesButtonClick(Sender: TObject;
  AButtonIndex: Integer);
begin
  if AButtonIndex = 1 then
    DoSelectInvUseItem
  else begin
    txtInvUseItem.Clear;
    aQueryParam.Param1 := EmptyStr;
   end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB03.DoSelectInvUseItem;
begin
  frmMultiSelect := TfrmMultiSelect.Create( Application );
  try
    frmMultiSelect.Connection := dtmMain.InvConnection;
    frmMultiSelect.KeyFields := 'ITEMID';
    frmMultiSelect.KeyValues := aQueryParam.Param1;
    frmMultiSelect.DisplayFields := 'ITEMID,代碼,DESCRIPTION,發票用途';
    frmMultiSelect.ResultFields := 'DESCRIPTION';
    frmMultiSelect.SQL.Text := Format(
      ' SELECT ITEMID, DESCRIPTION FROM INV028 ' +
      '  WHERE IDENTIFYID1 = ''%s''            ' +
      '    AND IDENTIFYID2 = ''%s''            ' +
      '    AND COMPID = ''%s''                 ' +
      '    AND NVL( REFNO, ''999'' ) <> ''1''  ' +      
      '  ORDER BY ITEMID                       ',
      [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID] ); 
    if ( frmMultiSelect.ShowModal = mrOk ) then
    begin
      aQueryParam.Param1 := frmMultiSelect.SelectedValue;
      txtInvUseItem.Text := frmMultiSelect.SelectedText;
    end;
  finally
    frmMultiSelect.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvB03.btnImportClick(Sender: TObject);
var
  T_Import : TStringList;
  iCount : Integer;
begin
  if OpenDialog1.Execute then
    edtImport.Text := OpenDialog1.FileName;
  if not FileExists( edtImport.Text ) then
  begin
    WarningMsg( '所指定的匯入發票資料檔不存在。' );
    Exit;
  end;
  T_Import := TStringList.Create;
  T_Import.LoadFromFile( edtImport.Text );
  FImportInv := EmptyStr;
  try
    Screen.Cursor := crSQLWait;
    for iCount := 0 to T_Import.Count -1  do
    begin
      if T_Import.Strings[iCount] <> EmptyStr then
      begin
        if FImportInv = EmptyStr then
        begin
          FImportInv := QuotedStr( UpperCase(T_Import.Strings[iCount]));
        end else
        begin
          FImportInv := FImportInv + ',' + QuotedStr( UpperCase( T_Import.Strings[iCount]));
        end;
      end;
    end;
  finally
    T_Import.Free;
    Screen.Cursor := crDefault;
  end;

end;

end.
