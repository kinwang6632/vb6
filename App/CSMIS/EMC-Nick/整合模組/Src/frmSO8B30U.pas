unit frmSO8B30U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, DB, Mask, DBCtrls, Buttons;

type
  TfrmSO8B30 = class(TForm)
    cobComp: TComboBox;
    Label4: TLabel;
    dsrCodeCD039: TDataSource;
    Label1: TLabel;
    dedParam1: TDBEdit;
    dsrSO125: TDataSource;
    Label2: TLabel;
    dedParam2: TDBEdit;
    Label3: TLabel;
    dedParam3: TDBEdit;
    Label5: TLabel;
    dedParam4: TDBEdit;
    Label6: TLabel;
    dedParam5: TDBEdit;
    Label7: TLabel;
    dedParam6: TDBEdit;
    btnOK: TBitBtn;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    dedParam7: TDBEdit;
    dedParam8: TDBEdit;
    Label16: TLabel;
    dedParam9: TDBEdit;
    Label17: TLabel;
    Label18: TLabel;
    Label19: TLabel;
    Label20: TLabel;
    Label21: TLabel;
    dedParam10: TDBEdit;
    dedParam11: TDBEdit;
    Label22: TLabel;
    OpenDialog1: TOpenDialog;
    SpeedButton1: TSpeedButton;
    btnCancel: TBitBtn;
    lbxParam12: TListBox;
    btnBox: TBitBtn;
    Label23: TLabel;
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure btnBoxClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
    sG_CompCode : String;
    function IsDataOk : Boolean;
  public
    { Public declarations }
    sG_PromCodeNoSQL,sG_PromName : String;
    G_TargetFieldValueStrList1 : TStringList;
    G_TempBoxItem : TStringList;
  end;

var
  frmSO8B30: TfrmSO8B30;

implementation

uses dtmMain1U, UCommonU, Ustru, frmMainMenuU, frmDbMultiSelectu;

{$R *.dfm}

procedure TfrmSO8B30.FormShow(Sender: TObject);
begin
    G_TargetFieldValueStrList1 := TStringList.Create;
    
    self.Caption:=frmMainMenu.setCaption('SO8B30','[ぇ疭把计]');

{
    lbxParam12.SetFocus;
    dedParam11.SetFocus;
    dedParam10.SetFocus;
    dedParam9.SetFocus;
    dedParam8.SetFocus;
    dedParam7.SetFocus;
    dedParam6.SetFocus;
    dedParam5.SetFocus;
    dedParam4.SetFocus;
    dedParam3.SetFocus;
    dedParam2.SetFocus;
    dedParam1.SetFocus;
}
end;

procedure TfrmSO8B30.FormCreate(Sender: TObject);
var
    ii : Integer;
begin
    //ъ CC&B (VB) 肚ㄓそ把计
    sG_CompCode := frmMainMenu.sG_CompCode;

    //そㄌ CC&B (VB) 肚ㄓそ砞﹚,┮ぃэ
    cobComp.Enabled := false;

    //そ
    TUCommonFun.AddObjectToComboBox(cobComp.Items, dsrCodeCD039.DataSet,NOT INSERT_NO_DATA_ITEM,'CodeNo','Description');
    TUCommonFun.setComboDefaultNdx(cobComp, sG_CompCode);

    //琩赣そ把计戈
    dtmMain1.getSo125Param(sG_CompCode);

    for ii:=0 to dtmMain1.G_PromNameStrList.Count-1 do
      lbxParam12.Items.Add(dtmMain1.G_PromNameStrList.Strings[ii]);
end;

procedure TfrmSO8B30.btnOKClick(Sender: TObject);
begin

    if IsDataOk then
    begin
      with dtmMain1.cdsSO125 do
      begin
          DisableControls;
          if State = dsInactive then
              Active := True;

          Edit;
          FieldByName('COMPCODE').AsInteger := StrToInt(sG_CompCode);
          FieldByName('OPERATOR').AsString := frmMainMenu.sG_User;
          FieldByName('UPDTIME').AsString := DateTimeToStr(Now);
          FieldByName('PARAM12').AsString := dtmMain1.sG_PromCodeAndName;

          dtmMain1.save2Db(dtmMain1.cdsSO125);

          MessageDlg('纗ЧΘ',mtInformation , [mbOK],0);
      end;
    end;

end;

function TfrmSO8B30.IsDataOk: Boolean;
var
    nL_Para1,nL_Para2,nL_Para3 : Integer;
    nL_Para4,nL_Para5,nL_Para6 : Integer;
    nL_Para7,nL_Para8,nL_Para9 : Integer;
    nL_Para10 : Integer;
    sL_Para11 : String;

begin
    with dtmMain1.cdsSo125 do
    begin
      nL_Para1 := FieldByName('Param1').AsInteger;
      nL_Para2 := FieldByName('Param2').AsInteger;
      nL_Para3 := FieldByName('Param3').AsInteger;
      nL_Para4 := FieldByName('Param4').AsInteger;
      nL_Para5 := FieldByName('Param5').AsInteger;
      nL_Para6 := FieldByName('Param6').AsInteger;
      nL_Para7 := FieldByName('Param7').AsInteger;
      nL_Para8 := FieldByName('Param8').AsInteger;
      nL_Para9 := FieldByName('Param9').AsInteger;
      nL_Para10 := FieldByName('Param10').AsInteger;
      sL_Para11 := FieldByName('Param11').AsString;


      //把计き计┪0ぃ︽
      if (nL_Para1<0) Or (nL_Para1>99999) then
      begin
        MessageDlg('ぃ0┪き计',mtError, [mbOK],0);
        dedParam1.SetFocus;
        Result := false;
        exit;
      end;

      if (nL_Para2<0) Or (nL_Para2>99999) then
      begin
        MessageDlg('ぃ0┪き计',mtError, [mbOK],0);
        dedParam2.SetFocus;
        Result := false;
        exit;
      end;

      if (nL_Para3<0) Or (nL_Para3>99999) then
      begin
        MessageDlg('ぃ0┪き计',mtError, [mbOK],0);
        dedParam3.SetFocus;
        Result := false;
        exit;
      end;

      if (nL_Para4<0) Or (nL_Para4>99999) then
      begin
        MessageDlg('ぃ0┪き计',mtError, [mbOK],0);
        dedParam4.SetFocus;
        Result := false;
        exit;
      end;

      if (nL_Para5<0) Or (nL_Para5>99999) then
      begin
        MessageDlg('ぃ0┪き计',mtError, [mbOK],0);
        dedParam5.SetFocus;
        Result := false;
        exit;
      end;

      if (nL_Para6<0) Or (nL_Para6>99999) then
      begin
        MessageDlg('ぃ0┪き计',mtError, [mbOK],0);
        dedParam6.SetFocus;
        Result := false;
        exit;
      end;

      if (nL_Para7<0) Or (nL_Para7>999) then
      begin
        MessageDlg('ぃ0┪计',mtError, [mbOK],0);
        dedParam7.SetFocus;
        Result := false;
        exit;
      end;

      if (nL_Para8<0) Or (nL_Para8>999) then
      begin
        MessageDlg('ぃ0┪计',mtError, [mbOK],0);
        dedParam8.SetFocus;
        Result := false;
        exit;
      end;

      if (nL_Para9<0) Or (nL_Para9>999) then
      begin
        MessageDlg('ぃ0┪计',mtError, [mbOK],0);
        dedParam9.SetFocus;
        Result := false;
        exit;
      end;

      if (nL_Para10<0) Or (nL_Para10>99) then
      begin
        MessageDlg('ぃ0┪计',mtError, [mbOK],0);
        dedParam10.SetFocus;
        Result := false;
        exit;
      end;

      Result := true;
    end;

end;

procedure TfrmSO8B30.SpeedButton1Click(Sender: TObject);
var
    sL_FileName,sL_TempFileName : String;
begin
    if (OpenDialog1.Execute) then
    begin
      sL_TempFileName := ExtractFileDir(OpenDialog1.FileName);

      if sL_TempFileName = '' then
        sL_FileName := 'C:'
      else
      begin
        if (sL_TempFileName = 'C:\') OR (sL_TempFileName = 'D:\') then
          sL_FileName := TUstr.replaceStr(sL_TempFileName,'\','')
        else
          sL_FileName := sL_TempFileName;
      end;


      dtmMain1.cdsSO125.Edit;
      dtmMain1.cdsSO125.FieldByName('Param11').AsString := sL_FileName;
      dtmMain1.cdsSO125.Post;
    end;

end;

procedure TfrmSO8B30.btnCancelClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmSO8B30.btnBoxClick(Sender: TObject);
var
  L_TargetFieldNamesStrList : TStringList;
  sL_PromCodeNoSQL : String;
  jj : Integer;
begin
  dtmMain1.getCD042Data;
  L_TargetFieldNamesStrList:= TStringList.Create;

  L_TargetFieldNamesStrList.Add('CODENO');
  L_TargetFieldNamesStrList.Add('DESCRIPTION');

  dtmMain1.qryCD042Data.FieldByName('CODENO').DisplayLabel := '穨叭兜ヘ絏';
  dtmMain1.qryCD042Data.FieldByName('DESCRIPTION').DisplayLabel := '穨叭兜ヘ嘿';

  if SelectMultiRecords('叫翴匡穨叭兜ヘ嘿', dtmMain1.qryCD042Data, 'CODENO;DESCRIPTION', ',', L_TargetFieldNamesStrList,
                G_TargetFieldValueStrList1,true,sL_PromCodeNoSQL) = mrOk then
  begin
    lbxParam12.Clear;
    sG_PromCodeNoSQL := '';
    sG_PromName := '';
    dtmMain1.G_PromCodeStrList.Clear;
    dtmMain1.G_PromNameStrList.Clear;

    for jj:=0 to G_TargetFieldValueStrList1.Count -1 do
    begin
      G_TempBoxItem := TUstr.ParseStrings(Trim(G_TargetFieldValueStrList1.Strings[jj]),',');


      if sG_PromCodeNoSQL = '' then
      begin
         sG_PromCodeNoSQL :=  G_TempBoxItem.Strings[0];
         sG_PromName :=  G_TempBoxItem.Strings[1];
      end
      else
      begin
         sG_PromCodeNoSQL := sG_PromCodeNoSQL + ',' + G_TempBoxItem.Strings[0];
         sG_PromName :=  sG_PromName + ',' + G_TempBoxItem.Strings[1];
      end;

      lbxParam12.Items.Add(G_TempBoxItem.Strings[1]);
      dtmMain1.G_PromCodeStrList.Add(G_TempBoxItem.Strings[0]);
      dtmMain1.G_PromNameStrList.Add(G_TempBoxItem.Strings[1]);
    end;

    if sG_PromCodeNoSQL = '' then
      dtmMain1.sG_PromCodeAndName := ''
    else
     dtmMain1.sG_PromCodeAndName := sG_PromCodeNoSQL + ';' + sG_PromName;
    {
    for jj:=0 to dtmMain1.G_PromCodeStrList.Count-1 do
    begin
      showmessage('btnBoxClick== ' + dtmMain1.G_PromCodeStrList.Strings[jj] + '-' + dtmMain1.G_PromNameStrList.Strings[jj]);
    end;
    }
  end;

end;

procedure TfrmSO8B30.FormClose(Sender: TObject; var Action: TCloseAction);
begin

    G_TargetFieldValueStrList1.Free;
    G_TargetFieldValueStrList1 := nil;

end;

end.
