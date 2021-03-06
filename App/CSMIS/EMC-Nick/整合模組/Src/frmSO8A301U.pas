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
    bG_IsCount : Boolean;    //true = ???????Z?H???q?p??  false = ?H???????p??
    bG_IsAmount : Boolean;   //true = ?p???????H???B?p??  false = ?H???????p??
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
        MessageDlg('?????J???q?O/???????W??',mtError, [mbOK],0);
        DBEdit3.SetFocus;
        Exit;
      end;

      if (dtmMain2.cdsSo111D.FieldByName('SO_PROVIDER_ID').AsString)='' then
      begin
        MessageDlg('?????J???q?O/???????N?X',mtError, [mbOK],0);
        DBEdit2.SetFocus;
        Exit;
      end;

      if (dtmMain2.cdsSo111D.FieldByName('COMP_TYPE').AsString)='' then
      begin
        MessageDlg('?????J???q?O/??????????',mtError, [mbOK],0);
        DBEdit1.SetFocus;
        Exit;
      end;
      if (dtmMain2.cdsSo111D.FieldByName('FORMULA_ID').AsString)='' then
      begin
        MessageDlg('?????J?????N?X',mtError, [mbOK],0);
        DBEdit6.SetFocus;
        Exit;
      end;
      if (dtmMain2.cdsSo111D.FieldByName('PRODUCT_ID').AsString)='' then
      begin
        MessageDlg('?????J???~?N?X',mtError, [mbOK],0);
        DBEdit4.SetFocus;
        Exit;
      end;
      if (dtmMain2.cdsSo111D.FieldByName('PRODUCT_NAME').AsString)='' then
      begin
        MessageDlg('?????J???~?W??',mtError, [mbOK],0);
        DBEdit5.SetFocus;
        Exit;
      end;
      if (dtmMain2.cdsSo111D.FieldByName('FORMULA_NAME_LU').AsString)='' then
      begin
        MessageDlg('?????J?????N?X???W??',mtError, [mbOK],0);
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
                  MessageDlg('?D?????L????????,?L?k?s?W',mtError, [mbOK],0);
                  dtmMain2.cdsSo111D.Cancel;
                  frmSO8A301.ModalResult:=MROK;
                end;
          end;

            //???d?O?_?????P product_id,FORMULA_ID,COMP_TYPE,SO_PROVIDER_ID
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
                MessageDlg('?????????L?k?s?W',mtError, [mbOK],0);
                cdsSo111D.Cancel;
                frmSO8A301.ModalResult:=MROK;
                Exit;
              end
              else if cdssqlstatement.FieldByName('CNT').Value=0 then//?Y?L?????N?s??
              begin
                try
                begin
                  cdsSo111D.FieldByName('PRODUCT_ID').AsString:=dtmMain2.cdsSo112M.FieldByName('PRODUCT_ID').AsString;
                  cdsSo111D.FieldByName('PRODUCT_NAME').AsString:=dtmMain2.cdsSo112M.FieldByName('PRODUCT_NAME').AsString;
                  cdsSo111D.FieldByName('OPERATOR').AsString:=frmMainMenu.sG_User;
                  cdsSo111D.FieldByName('UpdTime').AsString:=IntToStr(StrToInt(FormatDateTime('YYYY',now))-1911)+ FormatDateTime('/MM/DD hh:mm:ss',now);
                  dtmMain2.saveToDB(cdsSo111D);
                  MessageDlg('?s??????',mtInformation, [mbOK],0);
                  frmSO8A301.ModalResult:=MrOK;
                end;
                except
                  MessageDlg('?s??????',mtError, [mbOK],0);
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
          MessageDlg('?s??????',mtInformation, [mbOK],0);
          frmSO8A301.ModalResult:=MrOK;
          except
            MessageDlg('?s??????',mtError, [mbOK],0);
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
      if MessageDlg('?D???|???]?w?W?D??,?z?{?b?n?]?w???',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin
        frmSO8A301.Close;
        Exit;
      end;
    end;

    dtmMain2.getSO113AndCD039(sL_ProviderID);

    if (frmSO8A30.dsrSo111D.DataSet.State=dsBrowse) and (frmSO8A30.dsrSo111D.DataSet.RecordCount =0) then
      frmSO8A30.dsrSo111D.DataSet.Append;

    sL_SelectedFieldName :='SO_PROVIDER_DESC';//UpperCase(frmSO8A30.dgrSo111D.SelectedField.FieldName);

    if sL_SelectedFieldName='SO_PROVIDER_DESC' then //????????
    begin
      if SelectRecord('???I?????q?O/??????', frmSO8A30.dsrSo113Cd039.DataSet, 'CLASSIFY_ID;CODE;DESC','') = mrOk then
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
      btnCloseCancel.Caption:='????';

    end;
    if bPanelControlG=False then
    begin
      frmSO8A301.pnlPk.Enabled:=False;
      frmSO8A301.pnlNormal.Enabled:=True;
      dtmMain2.cdsSo111D.Edit;
      frmSO8A301.btnCloseCancel.Caption:='????';

    end;


    //???w?]??
    dtmMain2.cdsSo111D.FieldByName('PRODUCT_ID').AsString:=dtmMain2.cdsSo112M.FieldByName('PRODUCT_ID').AsString;
    dtmMain2.cdsSo111D.FieldByName('PRODUCT_NAME').AsString:=dtmMain2.cdsSo112M.FieldByName('PRODUCT_NAME').AsString;
    DBLookupComboBox2.DataSource.DataSet.FieldByName('COMPUTE_ISSUE').AsString:='1'; //?h???Z??

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
      DBLookupComboBox3.DataSource.DataSet.FieldByName('RANGE_UNIT').AsString:='1';//????
      DBLookupComboBox3.Enabled:=False;
    end;
    if (dtmMain2.cdsSqlStatement.FieldByName('Ref_No').AsString='4') or
       (dtmMain2.cdsSqlStatement.FieldByName('Ref_No').AsString='5') then
    begin
      DBLookupComboBox3.DataSource.DataSet.FieldByName('RANGE_UNIT').AsString:='2';//??????
      DBLookupComboBox3.Enabled:=False;
    end;
    if (dtmMain2.cdsSqlStatement.FieldByName('Ref_No').AsString='1') then
       DBLookupComboBox3.DataSource.DataSet.FieldByName('RANGE_UNIT').AsString:='2';//??????
}
end;

function TfrmSO8A301.CheckInputValueIsCorrectOrNot: Boolean;
var ii,jj:integer;
    sL_FieldNameSubscriber,sL_FieldNameAmount:String;
    fL_CurValue,fL_PriorValue,fL_TempValue : double;
    sL_DBEdit:String;
begin
    if not bG_IsCount then   //?????????????Y????????
    begin
      for ii:=1 to 5 do
      begin
        sL_FieldNameSubscriber := 'SUBSCRIBER_COUNT_PERCENT' + IntToStr(ii);
        fL_TempValue := dsrSo111D.DataSet.FieldByName(sL_FieldNameSubscriber).AsFloat;
        if (fL_TempValue < 0) OR (fL_TempValue > 100) then
        begin
          MessageDlg('???????????? 0~100 ????',mtError, [mbOK],0);

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


    if not bG_IsAmount then   //?p???????Y????????
    begin
      for ii:=1 to 5 do
      begin
        sL_FieldNameAmount := 'AMOUNT_PERCENT' + IntToStr(ii);
        fL_TempValue := dsrSo111D.DataSet.FieldByName(sL_FieldNameAmount).AsFloat;
        if (fL_TempValue < 0) OR (fL_TempValue > 100) then
        begin
          MessageDlg('???????????? 0~100 ????',mtError, [mbOK],0);

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

    //?????n??1??????
    if (dedCount_Percent1.DataSource.DataSet.FieldByName('SUBSCRIBER_COUNT_PERCENT1').AsString='') then
    begin
      MessageDlg('???????Z1?n????',mtError, [mbOK],0);
      dedCount_Percent1.SetFocus;
      Exit;
    end;

    if (dedAmount_Percent1.DataSource.DataSet.FieldByName('AMOUNT_PERCENT1').AsString='') then
    begin
      MessageDlg('???B??????1?n????',mtError, [mbOK],0);
      dedAmount_Percent1.SetFocus;
      //DBEdit14.DataSource.DataSet.FieldByName('AMOUNT_PERCENT1').FocusControl;
      Exit;
    end;

//SUBSCRIBER_COUNT_PERCENT ?Y???? ?h????????AmountRant?]?n????
    for ii:=1 to 5 do
    begin
      sL_FieldNameSubscriber:='SUBSCRIBER_COUNT_PERCENT'+intToStr(ii);
      sL_FieldNameAmount:='AMOUNT_PERCENT'+intToStr(ii);
      if (dtmMain2.cdsSo111D.FieldByName(sL_FieldNameSubscriber).AsFloat>0) and (dtmMain2.cdsSo111D.FieldByName(sL_FieldNameAmount).AsFloat=0) then
      begin
        MessageDlg('???B?????? '+intToStr(ii)+' ??????????',mtError, [mbOK],0);
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

//??2: ?????? ????5???? ?????????j(?B?p??100,?i???p??)
//??2: ?????? ?|???P?_?????J?????O?_?p??1(?p?G?p??1 =>(fL_AmountRant) ?j??1 (fL_AmountRant/100)
//??2: ?????? ?BSUBSCRIBER_COUNT_PERCENT ???????J??????????????100 ?Y???O ?t?????N?L?]?w??100 ??
//???????????? AmountRant  ???n????, ?YUSER ?????????O?b SUBSCRIBER_COUNT_PERCENT5 ???J?D100????
//???????t???buser ???????J???U?@?? SUBSCRIBER_COUNT_PERCENT ???J100 ,????????????AmountRant???n????
{
    if dtmMain2.cdsSo111D.FieldByName('RANGE_UNIT').AsString='2' then
    begin
      for ii:=1 to 5 do
      begin
        sL_FieldNameSubscriber:='SUBSCRIBER_COUNT_PERCENT'+intToStr(ii);
        if dtmMain2.cdsSo111D.FieldByName(sL_FieldNameSubscriber).AsFloat>100 then
        begin
        ShowMessage('???????Z'+intToStr(ii)+'???qor??????-?????j???@??');
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
          ShowMessage('???B??????'+intToStr(jj)+'?????j???@??');
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

//?P?_???????????? :1: ????2: ??????
// ??1: ????  ????5???? ?????????j
//?Y10?????????????J ?i?H ?Y??9???????h??14???]?n?????H??????

    fL_PriorValue:=0;
    fL_CurValue:=0;


    Result:=True;
    for ii:=1 to 5 do //???????Z-???qor??????????
    begin
      sL_FieldNameSubscriber:='SUBSCRIBER_COUNT_PERCENT'+intToStr(ii);
        if dtmMain2.cdsSo111D.FieldByName(sL_FieldNameSubscriber).AsFloat<>0 then
          begin
            fL_CurValue:=dtmMain2.cdsSo111D.FieldByName(sL_FieldNameSubscriber).AsFloat;
            if ii<>1 then
              begin
                if fL_CurValue<fL_PriorValue then
                  begin
                    MessageDlg('???????Z'+intToStr(ii)+'???qor??????'+'-?p?????????Z'+intToStr(ii-1)+'???qor??????',mtError, [mbOK],0);
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

//????????????

    if dtmMain2.cdsSo111D.FieldByName('RANGE_UNIT').AsString='2' then
    begin
      for ii:=1 to 5 do
      begin
        sL_FieldNameSubscriber:='SUBSCRIBER_COUNT_PERCENT'+intToStr(ii);
        sL_FieldNameAmount:='AMOUNT_PERCENT'+intToStr(ii);
        //???d?????@???? ?p?G?????J SUBSCRIBER_COUNT_PERCENT5  ?@?w?n??100
        if (ii=5) and (dtmMain2.cdsSo111D.FieldByName(sL_FieldNameSubscriber).AsFloat<>100) and
           (dtmMain2.cdsSo111D.FieldByName(sL_FieldNameSubscriber).AsFloat<>0) then
        begin
          dtmMain2.cdsSo111D.FieldByName(sL_FieldNameSubscriber).AsFloat:=100;
          if dtmMain2.cdsSo111D.FieldByName(sL_FieldNameAmount).AsFloat=0 then
          begin
           MessageDlg('???B-??????5 ????????',mtError, [mbOK],0);

           if ii=5 then
           begin
              frmSO8A301.dedAmount_Percent5.SetFocus;
              Exit;
           end

          end;
        end;
//?????J?????@?w?n???@???O100
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
                 MessageDlg('???B-??????'+intToStr(ii+1)+'????????',mtError, [mbOK],0);
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
    self.Caption := frmMainMenu.setCaption('SO8A30','[?????s??]');
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
    //???????????N?X?N?????@????
    sL_RefNo := dtmMain2.getSO110RefNo(sI_Formula_ID);

    //?u???????@?????????????~?i????,???L?????????i??
    if sL_RefNo = '1' then         //?????M?\??-??????
    begin
      bG_IsAmount := false;
      dlcRangeUnit.Enabled := true;
    end
    else
    begin
      dlcRangeUnit.Enabled := false;

      if sL_RefNo = '2' then        //???????q???Z-??????
      begin
        dsrSo111D.DataSet.FieldByName('RANGE_UNIT').AsString := '1';
        bG_IsCount := true;
        bG_IsAmount := true;
      end
      else if sL_RefNo = '3' then   //???????q???Z-??????
      begin
        dsrSo111D.DataSet.FieldByName('RANGE_UNIT').AsString := '1';
        bG_IsCount := true;
        bG_IsAmount := false;
      end
      else if sL_RefNo = '4' then   //?????????????Z-??????
      begin
        dsrSo111D.DataSet.FieldByName('RANGE_UNIT').AsString := '2';
        bG_IsCount := false;
        bG_IsAmount := true;
      end
      else if sL_RefNo = '5' then   //?????????????Z-??????
      begin
        dsrSo111D.DataSet.FieldByName('RANGE_UNIT').AsString := '2';
        bG_IsCount := false;
        bG_IsAmount := false;
      end;
    end;

    if bG_IsCount then
    begin
      lblCount_Percent1.Caption := '???????q???Z-1';
      lblCount_Percent2.Caption := '???????q???Z-2';
      lblCount_Percent3.Caption := '???????q???Z-3';
      lblCount_Percent4.Caption := '???????q???Z-4';
      lblCount_Percent5.Caption := '???????q???Z-5';
    end
    else
    begin
      lblCount_Percent1.Caption := '?????????????Z-1';
      lblCount_Percent2.Caption := '?????????????Z-2';
      lblCount_Percent3.Caption := '?????????????Z-3';
      lblCount_Percent4.Caption := '?????????????Z-4';
      lblCount_Percent5.Caption := '?????????????Z-5';
    end;


    if bG_IsAmount then
    begin
      lblAmount_Percent1.Caption := '???B-1';
      lblAmount_Percent2.Caption := '???B-2';
      lblAmount_Percent3.Caption := '???B-3';
      lblAmount_Percent4.Caption := '???B-4';
      lblAmount_Percent5.Caption := '???B-5';
    end
    else
    begin
      lblAmount_Percent1.Caption := '??????-1';
      lblAmount_Percent2.Caption := '??????-2';
      lblAmount_Percent3.Caption := '??????-3';
      lblAmount_Percent4.Caption := '??????-4';
      lblAmount_Percent5.Caption := '??????-5';
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

    if sL_RangeUnit = '1' then   //????????:???q
      bG_IsCount := true
    else                          //????????:??????
      bG_IsCount := false;

    if bG_IsCount then
    begin
      lblCount_Percent1.Caption := '???????q???Z-1';
      lblCount_Percent2.Caption := '???????q???Z-2';
      lblCount_Percent3.Caption := '???????q???Z-3';
      lblCount_Percent4.Caption := '???????q???Z-4';
      lblCount_Percent5.Caption := '???????q???Z-5';
    end
    else
    begin
      lblCount_Percent1.Caption := '?????????????Z-1';
      lblCount_Percent2.Caption := '?????????????Z-2';
      lblCount_Percent3.Caption := '?????????????Z-3';
      lblCount_Percent4.Caption := '?????????????Z-4';
      lblCount_Percent5.Caption := '?????????????Z-5';
    end;

end;

end.
