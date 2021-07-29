unit frmSO8A301U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DBCtrls, DB, StdCtrls, Mask, Buttons, ExtCtrls;

type
  TfrmSO8A301 = class(TForm)
    dsrSo111D: TDataSource;
    btnSave: TBitBtn;
    pnlNormal: TPanel;
    lblCount_Percent5: TLabel;
    Label8: TLabel;
    Label10: TLabel;
    lblCount_Percent1: TLabel;
    lblCount_Percent2: TLabel;
    lblCount_Percent3: TLabel;
    lblCount_Percent4: TLabel;
    lblAmount_Percent1: TLabel;
    lblAmount_Percent2: TLabel;
    lblAmount_Percent3: TLabel;
    lblAmount_Percent4: TLabel;
    lblAmount_Percent5: TLabel;
    DBLookupComboBox2: TDBLookupComboBox;
    dlcRangeUnit: TDBLookupComboBox;
    dedCount_Percent1: TDBEdit;
    dedCount_Percent2: TDBEdit;
    dedCount_Percent3: TDBEdit;
    dedCount_Percent4: TDBEdit;
    dedCount_Percent5: TDBEdit;
    dedAmount_Percent1: TDBEdit;
    dedAmount_Percent2: TDBEdit;
    dedAmount_Percent3: TDBEdit;
    dedAmount_Percent4: TDBEdit;
    dedAmount_Percent5: TDBEdit;
    pnlPk: TPanel;
    Label7: TLabel;
    DBEdit6: TDBEdit;
    Label6: TLabel;
    DBLookupComboBox1: TDBLookupComboBox;
    Label3: TLabel;
    BitBtn1: TBitBtn;
    DBEdit3: TDBEdit;
    Label1: TLabel;
    DBEdit1: TDBEdit;
    Label2: TLabel;
    DBEdit2: TDBEdit;
    Label4: TLabel;
    DBEdit4: TDBEdit;
    Label5: TLabel;
    DBEdit5: TDBEdit;
    btnCloseCancel: TBitBtn;
    dsrSo110Look: TDataSource;
    procedure btnSaveClick(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure btnCloseCancelClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    function CheckInputValueIsCorrectOrNot:Boolean;
    procedure FormShow(Sender: TObject);
    procedure DBLookupComboBox1CloseUp(Sender: TObject);
    procedure dlcRangeUnitCloseUp(Sender: TObject);
    procedure DBLookupComboBox1KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure dlcRangeUnitKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);


  private
    { Private declarations }
    bG_IsCount : Boolean;    //true = ��ƯŶZ�H�ƶq�p��  false = �H�ʤ���p��
    bG_IsAmount : Boolean;   //true = �p��ѼƥH���B�p��  false = �H�ʤ���p��
  public
    { Public declarations }
    procedure formulaIDComboBoxChange(sI_Formula_ID : String);
    procedure rangeUnitComboBoxChange;
  end;

var
  frmSO8A301: TfrmSO8A301;

implementation

uses dtmMain2U, frmSO8A30U, frmSO8A3011U, frmMainMenuU;

{$R *.dfm}

procedure TfrmSO8A301.btnSaveClick(Sender: TObject);
var bL_CheckInPutValueOk:Boolean;
    ii:integer;
    sL_SQL : String;
begin
    if bPanelControlG=True then
    begin
      if (dtmMain2.cdsSo111D.FieldByName('SO_PROVIDER_DESC').AsString)='' then
      begin
        MessageDlg('�п�J���q�O/�����ӦW��',mtError, [mbOK],0);
        DBEdit3.SetFocus;
        Exit;
      end;

      if (dtmMain2.cdsSo111D.FieldByName('SO_PROVIDER_ID').AsString)='' then
      begin
        MessageDlg('�п�J���q�O/�����ӥN�X',mtError, [mbOK],0);
        DBEdit2.SetFocus;
        Exit;
      end;

      if (dtmMain2.cdsSo111D.FieldByName('COMP_TYPE').AsString)='' then
      begin
        MessageDlg('�п�J���q�O/����������',mtError, [mbOK],0);
        DBEdit1.SetFocus;
        Exit;
      end;
      if (dtmMain2.cdsSo111D.FieldByName('FORMULA_ID').AsString)='' then
      begin
        MessageDlg('�п�J�����N�X',mtError, [mbOK],0);
        DBEdit6.SetFocus;
        Exit;
      end;
      if (dtmMain2.cdsSo111D.FieldByName('PRODUCT_ID').AsString)='' then
      begin
        MessageDlg('�п�J���~�N�X',mtError, [mbOK],0);
        DBEdit4.SetFocus;
        Exit;
      end;
      if (dtmMain2.cdsSo111D.FieldByName('PRODUCT_NAME').AsString)='' then
      begin
        MessageDlg('�п�J���~�W��',mtError, [mbOK],0);
        DBEdit5.SetFocus;
        Exit;
      end;
      if (dtmMain2.cdsSo111D.FieldByName('FORMULA_NAME_LU').AsString)='' then
      begin
        MessageDlg('�п�J�����N�X�ΦW��',mtError, [mbOK],0);
        DBLookupComboBox1.SetFocus;
        Exit;
      end;
    end;

    bL_CheckInPutValueOk := CheckInputValueIsCorrectOrNot;


//**********************************************************
    if bL_CheckInPutValueOk=True then
    begin
        if (dtmMain2.cdsSo111D.State=dsInsert) then
        begin
          with dtmMain2.cdsSqlStatement do
          begin
            Close;
            sL_SQL := 'Select count(*) as CNT from SO112 where PRODUCT_ID=''' + dtmMain2.cdsSo112M.FieldByName('PRODUCT_ID').AsString + '''';
            //dtmMain2.saveToFile(sL_SQL, 'D:\work\Error.txt');
            CommandText:= sL_SQL;
            OPEN;
                if dtmMain2.cdsSqlStatement.FieldByName('CNT').value=0 then
                begin
                  MessageDlg('�D�ɤ��L������,�L�k�s�W',mtError, [mbOK],0);
                  dtmMain2.cdsSo111D.Cancel;
                  frmSO8A301.ModalResult:=MROK;
                end;
          end;

            //�߬d�O�_���ۦP product_id,FORMULA_ID,COMP_TYPE,SO_PROVIDER_ID
          with dtmMain2 do
          begin
            cdssqlstatement.Close;
            sL_SQL:='Select count(*) as CNT from SO111 where PRODUCT_ID='''+dtmMain2.cdsSo112M.FieldByName('PRODUCT_ID').AsString+'''';
            sL_SQL:=sL_SQL+' AND COMP_TYPE='''+dtmMain2.cdsSO111D.FieldByName('COMP_TYPE').AsString+'''';
            sL_SQL:=sL_SQL+' AND SO_PROVIDER_ID='''+dtmMain2.cdsSo111D.FieldByName('SO_PROVIDER_ID').AsString+'''';
            cdssqlstatement.CommandText:=sL_SQL;
            cdssqlstatement.OPEN;
              if cdssqlstatement.FieldByName('CNT').Value>0 then
              begin
                MessageDlg('��ƭ��ƵL�k�s�W',mtError, [mbOK],0);
                cdsSo111D.Cancel;
                frmSO8A301.ModalResult:=MROK;
                Exit;
              end
              else if cdssqlstatement.FieldByName('CNT').Value=0 then//�Y�L���дN�s��
              begin
                try
                begin
                  cdsSo111D.FieldByName('PRODUCT_ID').AsString:=dtmMain2.cdsSo112M.FieldByName('PRODUCT_ID').AsString;
                  cdsSo111D.FieldByName('PRODUCT_NAME').AsString:=dtmMain2.cdsSo112M.FieldByName('PRODUCT_NAME').AsString;
                  cdsSo111D.FieldByName('OPERATOR').AsString:=frmMainMenu.sG_User;
                  cdsSo111D.FieldByName('UpdTime').AsString:=IntToStr(StrToInt(FormatDateTime('YYYY',now))-1911)+ FormatDateTime('/MM/DD hh:mm:ss',now);
                  dtmMain2.saveToDB(cdsSo111D);
                  MessageDlg('�s�ɧ���',mtInformation, [mbOK],0);
                  frmSO8A301.ModalResult:=MrOK;
                end;
                except
                  MessageDlg('�s�ɥ���',mtError, [mbOK],0);
                  Exit;
                end;
              end;
          end;
        end;
    //**********************************************************
        if (dtmMain2.cdsSo111D.State=dsEdit) then
        begin
          try
          dtmMain2.cdsSo111D.FieldByName('PRODUCT_ID').AsString:=dtmMain2.cdsSo112M.FieldByName('PRODUCT_ID').AsString;
          dtmMain2.cdsSo111D.FieldByName('PRODUCT_NAME').AsString:=dtmMain2.cdsSo112M.FieldByName('PRODUCT_NAME').AsString;
          dtmMain2.cdsSo111D.FieldByName('OPERATOR').AsString:=frmMainMenu.sG_User;
          dtmMain2.cdsSo111D.FieldByName('UpdTime').AsString:=IntToStr(StrToInt(FormatDateTime('YYYY',now))-1911)+ FormatDateTime('/MM/DD hh:mm:ss',now);
          dtmMain2.cdsSo111D.post;
          dtmMain2.cdsSo111D.ApplyUpdates(-1);
          MessageDlg('�s�ɧ���',mtInformation, [mbOK],0);
          frmSO8A301.ModalResult:=MrOK;
          except
            MessageDlg('�s�ɥ���',mtError, [mbOK],0);
            Exit;
          end;
        end;
    end;
end;





procedure TfrmSO8A301.BitBtn1Click(Sender: TObject);
var
  sL_SelectedFieldName,sL_ProviderID : String;
begin
    sL_ProviderID := dtmMain2.cdsSo112M.FieldByName('PROVIDER_ID').AsString;

    if sL_ProviderID = '' then
    begin
      if MessageDlg('�D�ɩ|���]�w�W�D��,�z�{�b�n�]�w��?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin
        frmSO8A301.Close;
        Exit;
      end;
    end;

    dtmMain2.getSO113AndCD039(sL_ProviderID);

    if (frmSO8A30.dsrSo111D.DataSet.State=dsBrowse) and (frmSO8A30.dsrSo111D.DataSet.RecordCount =0) then
      frmSO8A30.dsrSo111D.DataSet.Append;

    sL_SelectedFieldName :='SO_PROVIDER_DESC';//UpperCase(frmSO8A30.dgrSo111D.SelectedField.FieldName);

    if sL_SelectedFieldName='SO_PROVIDER_DESC' then //�������
    begin
      if SelectRecord('���I�綠�q�O/������', frmSO8A30.dsrSo113Cd039.DataSet, 'CLASSIFY_ID;CODE;DESC','') = mrOk then
      begin
      if not (frmSO8A30.dsrSo111D.DataSet.State in [dsEdit, dsInsert]) then
        frmSO8A30.dsrSo111D.DataSet.Edit;
        frmSO8A30.dsrSo111D.DataSet.FieldByName('SO_PROVIDER_DESC').AsString :=frmSO8A30.dsrSo113Cd039.DataSet.FieldByName('DESC').AsString;
        frmSO8A30.dsrSo111D.DataSet.FieldByName('COMP_TYPE').AsString :=frmSO8A30.dsrSo113Cd039.DataSet.FieldByName('CLASSIFY_ID').AsString;
        frmSO8A30.dsrSo111D.DataSet.FieldByName('So_Provider_ID').AsString :=frmSO8A30.dsrSo113Cd039.DataSet.FieldByName('CODE').AsString;
      end;
    end;
end;

procedure TfrmSO8A301.btnCloseCancelClick(Sender: TObject);
begin
    frmSO8A301.ModalResult:=mrOK;
    dtmMain2.cdsSo111D.Cancel;
    pnlPK.Enabled:=False;
    pnlNormal.Enabled:=False;
end;

procedure TfrmSO8A301.FormCreate(Sender: TObject);
var sL_SQL:String;
begin
    if bPanelControlG=True then
    begin
      frmSO8A301.pnlPk.Enabled:=True;
      frmSO8A301.pnlNormal.Enabled:=True;
      dtmMain2.cdsSo111D.Append;
      btnCloseCancel.Caption:='����';

    end;
    if bPanelControlG=False then
    begin
      frmSO8A301.pnlPk.Enabled:=False;
      frmSO8A301.pnlNormal.Enabled:=True;
      dtmMain2.cdsSo111D.Edit;
      frmSO8A301.btnCloseCancel.Caption:='����';

    end;


    //���w�]��
    dtmMain2.cdsSo111D.FieldByName('PRODUCT_ID').AsString:=dtmMain2.cdsSo112M.FieldByName('PRODUCT_ID').AsString;
    dtmMain2.cdsSo111D.FieldByName('PRODUCT_NAME').AsString:=dtmMain2.cdsSo112M.FieldByName('PRODUCT_NAME').AsString;
    DBLookupComboBox2.DataSource.DataSet.FieldByName('COMPUTE_ISSUE').AsString:='1'; //�h�ŶZ��

{
    with dtmMain2.cdsSqlStatement do
    begin
      sL_SQL:='select Ref_No from So110 where Formula_Id='''+dtmMain2.cdsSo112M.FieldByName('Formula_Id').AsString+'''';
      Close;
      CommandText:=sL_SQL;
      OPEN;
    end;

    if (dtmMain2.cdsSqlStatement.FieldByName('Ref_No').AsString='2') or
       (dtmMain2.cdsSqlStatement.FieldByName('Ref_No').AsString='3') then
    begin
      DBLookupComboBox3.DataSource.DataSet.FieldByName('RANGE_UNIT').AsString:='1';//���
      DBLookupComboBox3.Enabled:=False;
    end;
    if (dtmMain2.cdsSqlStatement.FieldByName('Ref_No').AsString='4') or
       (dtmMain2.cdsSqlStatement.FieldByName('Ref_No').AsString='5') then
    begin
      DBLookupComboBox3.DataSource.DataSet.FieldByName('RANGE_UNIT').AsString:='2';//�ʤ���
      DBLookupComboBox3.Enabled:=False;
    end;
    if (dtmMain2.cdsSqlStatement.FieldByName('Ref_No').AsString='1') then
       DBLookupComboBox3.DataSource.DataSet.FieldByName('RANGE_UNIT').AsString:='2';//�ʤ���
}
end;

function TfrmSO8A301.CheckInputValueIsCorrectOrNot: Boolean;
var ii,jj:integer;
    sL_FieldNameSubscriber,sL_FieldNameAmount:String;
    fL_CurValue,fL_PriorValue,fL_TempValue : double;
    sL_DBEdit:String;
begin
    if not bG_IsCount then   //��Ƥ�����Y���ʤ���
    begin
      for ii:=1 to 5 do
      begin
        sL_FieldNameSubscriber := 'SUBSCRIBER_COUNT_PERCENT' + IntToStr(ii);
        fL_TempValue := dsrSo111D.DataSet.FieldByName(sL_FieldNameSubscriber).AsFloat;
        if (fL_TempValue < 0) OR (fL_TempValue > 100) then
        begin
          MessageDlg('�ʤ��񶷤��� 0~100 ����',mtError, [mbOK],0);

          if ii = 1 then
            dedCount_Percent1.SetFocus
          else if ii = 2 then
            dedCount_Percent2.SetFocus
          else if ii = 3 then
            dedCount_Percent3.SetFocus
          else if ii = 4 then
            dedCount_Percent4.SetFocus
          else if ii = 5 then
            dedCount_Percent5.SetFocus;

          exit;
        end;
      end;
    end;


    if not bG_IsAmount then   //�p��ѼƭY���ʤ���
    begin
      for ii:=1 to 5 do
      begin
        sL_FieldNameAmount := 'AMOUNT_PERCENT' + IntToStr(ii);
        fL_TempValue := dsrSo111D.DataSet.FieldByName(sL_FieldNameAmount).AsFloat;
        if (fL_TempValue < 0) OR (fL_TempValue > 100) then
        begin
          MessageDlg('�ʤ��񶷤��� 0~100 ����',mtError, [mbOK],0);

          if ii = 1 then
            dedAmount_Percent1.SetFocus
          else if ii = 2 then
            dedAmount_Percent2.SetFocus
          else if ii = 3 then
            dedAmount_Percent3.SetFocus
          else if ii = 4 then
            dedAmount_Percent4.SetFocus
          else if ii = 5 then
            dedAmount_Percent5.SetFocus;

          exit;
        end;
      end;
    end;

    //�ܤ֭n��1�զ���
    if (dedCount_Percent1.DataSource.DataSet.FieldByName('SUBSCRIBER_COUNT_PERCENT1').AsString='') then
    begin
      MessageDlg('��ƯŶZ1�n����',mtError, [mbOK],0);
      dedCount_Percent1.SetFocus;
      Exit;
    end;

    if (dedAmount_Percent1.DataSource.DataSet.FieldByName('AMOUNT_PERCENT1').AsString='') then
    begin
      MessageDlg('���B�ʤ���1�n����',mtError, [mbOK],0);
      dedAmount_Percent1.SetFocus;
      //DBEdit14.DataSource.DataSet.FieldByName('AMOUNT_PERCENT1').FocusControl;
      Exit;
    end;

//SUBSCRIBER_COUNT_PERCENT �Y���� �h�۹�����AmountRant�]�n����
    for ii:=1 to 5 do
    begin
      sL_FieldNameSubscriber:='SUBSCRIBER_COUNT_PERCENT'+intToStr(ii);
      sL_FieldNameAmount:='AMOUNT_PERCENT'+intToStr(ii);
      if (dtmMain2.cdsSo111D.FieldByName(sL_FieldNameSubscriber).AsFloat>0) and (dtmMain2.cdsSo111D.FieldByName(sL_FieldNameAmount).AsFloat=0) then
      begin
        MessageDlg('���B�ʤ��� '+intToStr(ii)+' ���ର�ŭ�',mtError, [mbOK],0);
        if ii=1 then
        begin
           frmSO8A301.dedAmount_Percent1.SetFocus;
           Exit;
        end
        else if ii=2 then
        begin
           frmSO8A301.dedAmount_Percent2.SetFocus;
           Exit;
        end
        else if ii=3 then
        begin
           frmSO8A301.dedAmount_Percent3.SetFocus;
           Exit;
        end
        else if ii=4 then
        begin
           frmSO8A301.dedAmount_Percent4.SetFocus;
           Exit;
        end
        else if ii=5 then
        begin
           frmSO8A301.dedAmount_Percent5.SetFocus;
           Exit;
        end
      end;
    end;

//��2: �ʤ��� ����5��� ��ȶ����j(�B�p��100,�i���p��)
//��2: �ʤ��� �|���P�_�ҿ�J���ȬO�_�p��1(�p�G�p��1 =>(fL_AmountRant) �j��1 (fL_AmountRant/100)
//��2: �ʤ��� �BSUBSCRIBER_COUNT_PERCENT �̫��J�����ȥ�����100 �Y���O �t�ζ��N�L�]�w��100 ��
//���ܬ۹����� AmountRant  ���n����, �YUSER �̫�ä��O�b SUBSCRIBER_COUNT_PERCENT5 ��J�D100����
//����Шt�Φbuser �̫��J���U�@�� SUBSCRIBER_COUNT_PERCENT ��J100 ,���ܬ۹�����AmountRant���n����
{
    if dtmMain2.cdsSo111D.FieldByName('RANGE_UNIT').AsString='2' then
    begin
      for ii:=1 to 5 do
      begin
        sL_FieldNameSubscriber:='SUBSCRIBER_COUNT_PERCENT'+intToStr(ii);
        if dtmMain2.cdsSo111D.FieldByName(sL_FieldNameSubscriber).AsFloat>100 then
        begin
        ShowMessage('��ƯŶZ'+intToStr(ii)+'�ƶqor�ʤ���-����j��@��');
        Result:=False;
          if ii=1 then
          begin
             frmSO8A301.dedCount_Percent1.SetFocus;
             Exit;
          end
          else if ii=2 then
          begin
             frmSO8A301.dedCount_Percent2.SetFocus;
             Exit;
          end
          else if ii=3 then
          begin
             frmSO8A301.dedCount_Percent3.SetFocus;
             Exit;
          end
          else if ii=4 then
          begin
             frmSO8A301.dedCount_Percent4.SetFocus;
             Exit;
          end
          else if ii=5 then
          begin
             frmSO8A301.dedCount_Percent5.SetFocus;
             Exit;
          end
          else Result:=True;
        end;
      end;

      for jj:=1 to 5 do
      begin
        sL_FieldNameAmount:='AMOUNT_PERCENT'+intToStr(jj);
        if dtmMain2.cdsSo111D.FieldByName(sL_FieldNameAmount).AsFloat>100 then
        begin
          ShowMessage('���B�ʤ���'+intToStr(jj)+'����j��@��');
          Result:=False;
          if jj=1 then
          begin
            frmSO8A301.dedAmount_Percent1.SetFocus;
            Exit;
          end
          else if jj=2 then
          begin
           frmSO8A301.dedAmount_Percent2.SetFocus;
           Exit;
          end
          else if jj=3 then
          begin
            frmSO8A301.dedAmount_Percent3.SetFocus;
            Exit;
          end
          else if jj=4 then
          begin
            frmSO8A301.dedAmount_Percent4.SetFocus;
            Exit;
          end
          else if jj=5 then
            frmSO8A301.dedAmount_Percent5.SetFocus;
            Exit;
          end
          else Result:=True;
      end;
    end;
}

//�P�_��Ƥ����� :1: ���2: �ʤ���
// ��1: ���  ����5��� ��ȶ����j
//�Y10�����ҥ���J �i�H �Y��9�Ӧ��ȫh��14�Ӥ]�n���ȥH������

    fL_PriorValue:=0;
    fL_CurValue:=0;


    Result:=True;
    for ii:=1 to 5 do //��ƯŶZ-�ƶqor�ʤ������
    begin
      sL_FieldNameSubscriber:='SUBSCRIBER_COUNT_PERCENT'+intToStr(ii);
        if dtmMain2.cdsSo111D.FieldByName(sL_FieldNameSubscriber).AsFloat<>0 then
          begin
            fL_CurValue:=dtmMain2.cdsSo111D.FieldByName(sL_FieldNameSubscriber).AsFloat;
            if ii<>1 then
              begin
                if fL_CurValue<fL_PriorValue then
                  begin
                    MessageDlg('��ƯŶZ'+intToStr(ii)+'�ƶqor�ʤ���'+'-�p���ƯŶZ'+intToStr(ii-1)+'�ƶqor�ʤ���',mtError, [mbOK],0);
                    Result:=False;
                    if ii=1 then
                    begin
                       frmSO8A301.dedCount_Percent1.SetFocus;
                       Exit;
                    end
                    else if ii=2 then
                    begin
                       frmSO8A301.dedCount_Percent2.SetFocus;
                       Exit;
                    end
                    else if ii=3 then
                    begin
                       frmSO8A301.dedCount_Percent3.SetFocus;
                       Exit;
                    end
                    else if ii=4 then
                    begin
                       frmSO8A301.dedCount_Percent4.SetFocus;
                       Exit;
                    end
                    else if ii=5 then
                    begin
                       frmSO8A301.dedCount_Percent5.SetFocus;
                       Exit;
                    end
                    else Result:=True;
                  end
                else fL_PriorValue:=fL_CurValue;
              end
            else
              fL_PriorValue:=fL_CurValue;
          end
        else
          break;
    end;

//��ܦʤ����

    if dtmMain2.cdsSo111D.FieldByName('RANGE_UNIT').AsString='2' then
    begin
      for ii:=1 to 5 do
      begin
        sL_FieldNameSubscriber:='SUBSCRIBER_COUNT_PERCENT'+intToStr(ii);
        sL_FieldNameAmount:='AMOUNT_PERCENT'+intToStr(ii);
        //�߬d�̫�@�ӭ� �p�G����J SUBSCRIBER_COUNT_PERCENT5  �@�w�n��100
        if (ii=5) and (dtmMain2.cdsSo111D.FieldByName(sL_FieldNameSubscriber).AsFloat<>100) and
           (dtmMain2.cdsSo111D.FieldByName(sL_FieldNameSubscriber).AsFloat<>0) then
        begin
          dtmMain2.cdsSo111D.FieldByName(sL_FieldNameSubscriber).AsFloat:=100;
          if dtmMain2.cdsSo111D.FieldByName(sL_FieldNameAmount).AsFloat=0 then
          begin
           MessageDlg('���B-�ʤ���5 ��������',mtError, [mbOK],0);

           if ii=5 then
           begin
              frmSO8A301.dedAmount_Percent5.SetFocus;
              Exit;
           end

          end;
        end;
//�ҿ�J���Ȥ@�w�n���@�ӬO100
        if ii<>5 then
        begin
          if (dtmMain2.cdsSo111D.FieldByName(sL_FieldNameSubscriber).AsFloat<>0) and (dtmMain2.cdsSo111D.FieldByName(sL_FieldNameSubscriber).AsFloat<>100) then
            begin
              sL_FieldNameSubscriber:='SUBSCRIBER_COUNT_PERCENT'+intToStr(ii+1);
              sL_FieldNameAmount:='AMOUNT_PERCENT'+intToStr(ii+1);
              if dtmMain2.cdsSo111D.FieldByName(sL_FieldNameSubscriber).AsFloat=0 then
              begin
                dtmMain2.cdsSo111D.FieldByName(sL_FieldNameSubscriber).AsFloat:=100;
                if dtmMain2.cdsSo111D.FieldByName(sL_FieldNameAmount).AsFloat=0 then
                begin
                 MessageDlg('���B-�ʤ���'+intToStr(ii+1)+'��������',mtError, [mbOK],0);
                 //dtmMain2.cdsSo111D.FieldByName(sL_FieldNameAmount).FocusControl;
                 Result:=False;
                 Exit;
                end;
              end;
            end;
        end;
      end;
    end;
end;

procedure TfrmSO8A301.FormShow(Sender: TObject);
var
    sL_Formula_ID : String;
begin
    self.Caption := frmMainMenu.setCaption('SO8A30','[���ӽs��]');
    if bPanelControlG=True then
       frmSO8A301.BitBtn1.SetFocus
    else frmSO8A301.DBLookupComboBox2.SetFocus;

    sL_Formula_ID := dsrSo111D.DataSet.FieldByName('FORMULA_ID').AsString;
    formulaIDComboBoxChange(sL_Formula_ID);
end;

procedure TfrmSO8A301.DBLookupComboBox1CloseUp(Sender: TObject);
var
    sL_Formula_ID : String;
begin
    sL_Formula_ID := dsrSo111D.DataSet.FieldByName('FORMULA_ID').AsString;
    formulaIDComboBoxChange(sL_Formula_ID);
end;

procedure TfrmSO8A301.DBLookupComboBox1KeyUp(Sender: TObject;
  var Key: Word; Shift: TShiftState);
var
    sL_Formula_ID : String;
begin
    sL_Formula_ID := dsrSo111D.DataSet.FieldByName('FORMULA_ID').AsString;
    formulaIDComboBoxChange(sL_Formula_ID);
end;


procedure TfrmSO8A301.formulaIDComboBoxChange(sI_Formula_ID : String);
var
    sL_RefNo : String;
begin
    //��惡�����N�X�N����@����
    sL_RefNo := dtmMain2.getSO110RefNo(sI_Formula_ID);

    //�u�������@��Ƥ�����~�i�ܴ�,��L���������i��
    if sL_RefNo = '1' then         //�ⶥ�M�\��-��Ҩ�
    begin
      bG_IsAmount := false;
      dlcRangeUnit.Enabled := true;
    end
    else
    begin
      dlcRangeUnit.Enabled := false;

      if sL_RefNo = '2' then        //��Ƽƶq�ŶZ-�����
      begin
        dsrSo111D.DataSet.FieldByName('RANGE_UNIT').AsString := '1';
        bG_IsCount := true;
        bG_IsAmount := true;
      end
      else if sL_RefNo = '3' then   //��Ƽƶq�ŶZ-��Ҩ�
      begin
        dsrSo111D.DataSet.FieldByName('RANGE_UNIT').AsString := '1';
        bG_IsCount := true;
        bG_IsAmount := false;
      end
      else if sL_RefNo = '4' then   //��Ʀʤ���ŶZ-�����
      begin
        dsrSo111D.DataSet.FieldByName('RANGE_UNIT').AsString := '2';
        bG_IsCount := false;
        bG_IsAmount := true;
      end
      else if sL_RefNo = '5' then   //��Ʀʤ���ŶZ-��Ҩ�
      begin
        dsrSo111D.DataSet.FieldByName('RANGE_UNIT').AsString := '2';
        bG_IsCount := false;
        bG_IsAmount := false;
      end;
    end;

    if bG_IsCount then
    begin
      lblCount_Percent1.Caption := '��Ƽƶq�ŶZ-1';
      lblCount_Percent2.Caption := '��Ƽƶq�ŶZ-2';
      lblCount_Percent3.Caption := '��Ƽƶq�ŶZ-3';
      lblCount_Percent4.Caption := '��Ƽƶq�ŶZ-4';
      lblCount_Percent5.Caption := '��Ƽƶq�ŶZ-5';
    end
    else
    begin
      lblCount_Percent1.Caption := '��Ʀʤ���ŶZ-1';
      lblCount_Percent2.Caption := '��Ʀʤ���ŶZ-2';
      lblCount_Percent3.Caption := '��Ʀʤ���ŶZ-3';
      lblCount_Percent4.Caption := '��Ʀʤ���ŶZ-4';
      lblCount_Percent5.Caption := '��Ʀʤ���ŶZ-5';
    end;


    if bG_IsAmount then
    begin
      lblAmount_Percent1.Caption := '���B-1';
      lblAmount_Percent2.Caption := '���B-2';
      lblAmount_Percent3.Caption := '���B-3';
      lblAmount_Percent4.Caption := '���B-4';
      lblAmount_Percent5.Caption := '���B-5';
    end
    else
    begin
      lblAmount_Percent1.Caption := '�ʤ���-1';
      lblAmount_Percent2.Caption := '�ʤ���-2';
      lblAmount_Percent3.Caption := '�ʤ���-3';
      lblAmount_Percent4.Caption := '�ʤ���-4';
      lblAmount_Percent5.Caption := '�ʤ���-5';
    end;
end;




procedure TfrmSO8A301.dlcRangeUnitCloseUp(Sender: TObject);
begin
    rangeUnitComboBoxChange;
end;


procedure TfrmSO8A301.dlcRangeUnitKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
    rangeUnitComboBoxChange;

end;



procedure TfrmSO8A301.rangeUnitComboBoxChange;
var
    sL_RangeUnit : String;
begin
    sL_RangeUnit := dsrSo111D.DataSet.FieldByName('RANGE_UNIT').AsString;

    if sL_RangeUnit = '1' then   //��Ƴ��:�ƶq
      bG_IsCount := true
    else                          //��Ƴ��:�ʤ���
      bG_IsCount := false;

    if bG_IsCount then
    begin
      lblCount_Percent1.Caption := '��Ƽƶq�ŶZ-1';
      lblCount_Percent2.Caption := '��Ƽƶq�ŶZ-2';
      lblCount_Percent3.Caption := '��Ƽƶq�ŶZ-3';
      lblCount_Percent4.Caption := '��Ƽƶq�ŶZ-4';
      lblCount_Percent5.Caption := '��Ƽƶq�ŶZ-5';
    end
    else
    begin
      lblCount_Percent1.Caption := '��Ʀʤ���ŶZ-1';
      lblCount_Percent2.Caption := '��Ʀʤ���ŶZ-2';
      lblCount_Percent3.Caption := '��Ʀʤ���ŶZ-3';
      lblCount_Percent4.Caption := '��Ʀʤ���ŶZ-4';
      lblCount_Percent5.Caption := '��Ʀʤ���ŶZ-5';
    end;

end;

end.
