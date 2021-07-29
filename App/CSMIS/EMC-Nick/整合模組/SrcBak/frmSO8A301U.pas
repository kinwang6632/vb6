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
    bG_IsCount : Boolean;    //true = 戶數級距以數量計算  false = 以百分比計算
    bG_IsAmount : Boolean;   //true = 計算參數以金額計算  false = 以百分比計算
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
        MessageDlg('請輸入公司別/供應商名稱',mtError, [mbOK],0);
        DBEdit3.SetFocus;
        Exit;
      end;

      if (dtmMain2.cdsSo111D.FieldByName('SO_PROVIDER_ID').AsString)='' then
      begin
        MessageDlg('請輸入公司別/供應商代碼',mtError, [mbOK],0);
        DBEdit2.SetFocus;
        Exit;
      end;

      if (dtmMain2.cdsSo111D.FieldByName('COMP_TYPE').AsString)='' then
      begin
        MessageDlg('請輸入公司別/供應商類型',mtError, [mbOK],0);
        DBEdit1.SetFocus;
        Exit;
      end;
      if (dtmMain2.cdsSo111D.FieldByName('FORMULA_ID').AsString)='' then
      begin
        MessageDlg('請輸入公式代碼',mtError, [mbOK],0);
        DBEdit6.SetFocus;
        Exit;
      end;
      if (dtmMain2.cdsSo111D.FieldByName('PRODUCT_ID').AsString)='' then
      begin
        MessageDlg('請輸入產品代碼',mtError, [mbOK],0);
        DBEdit4.SetFocus;
        Exit;
      end;
      if (dtmMain2.cdsSo111D.FieldByName('PRODUCT_NAME').AsString)='' then
      begin
        MessageDlg('請輸入產品名稱',mtError, [mbOK],0);
        DBEdit5.SetFocus;
        Exit;
      end;
      if (dtmMain2.cdsSo111D.FieldByName('FORMULA_NAME_LU').AsString)='' then
      begin
        MessageDlg('請輸入公式代碼及名稱',mtError, [mbOK],0);
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
                  MessageDlg('主檔內無任何資料,無法新增',mtError, [mbOK],0);
                  dtmMain2.cdsSo111D.Cancel;
                  frmSO8A301.ModalResult:=MROK;
                end;
          end;

            //撿查是否有相同 product_id,FORMULA_ID,COMP_TYPE,SO_PROVIDER_ID
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
                MessageDlg('資料重複無法新增',mtError, [mbOK],0);
                cdsSo111D.Cancel;
                frmSO8A301.ModalResult:=MROK;
                Exit;
              end
              else if cdssqlstatement.FieldByName('CNT').Value=0 then//若無重覆就存檔
              begin
                try
                begin
                  cdsSo111D.FieldByName('PRODUCT_ID').AsString:=dtmMain2.cdsSo112M.FieldByName('PRODUCT_ID').AsString;
                  cdsSo111D.FieldByName('PRODUCT_NAME').AsString:=dtmMain2.cdsSo112M.FieldByName('PRODUCT_NAME').AsString;
                  cdsSo111D.FieldByName('OPERATOR').AsString:=frmMainMenu.sG_User;
                  cdsSo111D.FieldByName('UpdTime').AsString:=IntToStr(StrToInt(FormatDateTime('YYYY',now))-1911)+ FormatDateTime('/MM/DD hh:mm:ss',now);
                  dtmMain2.saveToDB(cdsSo111D);
                  MessageDlg('存檔完成',mtInformation, [mbOK],0);
                  frmSO8A301.ModalResult:=MrOK;
                end;
                except
                  MessageDlg('存檔失敗',mtError, [mbOK],0);
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
          MessageDlg('存檔完成',mtInformation, [mbOK],0);
          frmSO8A301.ModalResult:=MrOK;
          except
            MessageDlg('存檔失敗',mtError, [mbOK],0);
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
      if MessageDlg('主檔尚未設定頻道商,您現在要設定嗎?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin
        frmSO8A301.Close;
        Exit;
      end;
    end;

    dtmMain2.getSO113AndCD039(sL_ProviderID);

    if (frmSO8A30.dsrSo111D.DataSet.State=dsBrowse) and (frmSO8A30.dsrSo111D.DataSet.RecordCount =0) then
      frmSO8A30.dsrSo111D.DataSet.Append;

    sL_SelectedFieldName :='SO_PROVIDER_DESC';//UpperCase(frmSO8A30.dgrSo111D.SelectedField.FieldName);

    if sL_SelectedFieldName='SO_PROVIDER_DESC' then //選取部門
    begin
      if SelectRecord('請點選公司別/供應商', frmSO8A30.dsrSo113Cd039.DataSet, 'CLASSIFY_ID;CODE;DESC','') = mrOk then
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
      btnCloseCancel.Caption:='取消';

    end;
    if bPanelControlG=False then
    begin
      frmSO8A301.pnlPk.Enabled:=False;
      frmSO8A301.pnlNormal.Enabled:=True;
      dtmMain2.cdsSo111D.Edit;
      frmSO8A301.btnCloseCancel.Caption:='取消';

    end;


    //給預設值
    dtmMain2.cdsSo111D.FieldByName('PRODUCT_ID').AsString:=dtmMain2.cdsSo112M.FieldByName('PRODUCT_ID').AsString;
    dtmMain2.cdsSo111D.FieldByName('PRODUCT_NAME').AsString:=dtmMain2.cdsSo112M.FieldByName('PRODUCT_NAME').AsString;
    DBLookupComboBox2.DataSource.DataSet.FieldByName('COMPUTE_ISSUE').AsString:='1'; //多級距制

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
      DBLookupComboBox3.DataSource.DataSet.FieldByName('RANGE_UNIT').AsString:='1';//戶數
      DBLookupComboBox3.Enabled:=False;
    end;
    if (dtmMain2.cdsSqlStatement.FieldByName('Ref_No').AsString='4') or
       (dtmMain2.cdsSqlStatement.FieldByName('Ref_No').AsString='5') then
    begin
      DBLookupComboBox3.DataSource.DataSet.FieldByName('RANGE_UNIT').AsString:='2';//百分比
      DBLookupComboBox3.Enabled:=False;
    end;
    if (dtmMain2.cdsSqlStatement.FieldByName('Ref_No').AsString='1') then
       DBLookupComboBox3.DataSource.DataSet.FieldByName('RANGE_UNIT').AsString:='2';//百分比
}
end;

function TfrmSO8A301.CheckInputValueIsCorrectOrNot: Boolean;
var ii,jj:integer;
    sL_FieldNameSubscriber,sL_FieldNameAmount:String;
    fL_CurValue,fL_PriorValue,fL_TempValue : double;
    sL_DBEdit:String;
begin
    if not bG_IsCount then   //戶數比較單位若為百分比
    begin
      for ii:=1 to 5 do
      begin
        sL_FieldNameSubscriber := 'SUBSCRIBER_COUNT_PERCENT' + IntToStr(ii);
        fL_TempValue := dsrSo111D.DataSet.FieldByName(sL_FieldNameSubscriber).AsFloat;
        if (fL_TempValue < 0) OR (fL_TempValue > 100) then
        begin
          MessageDlg('百分比須介於 0~100 之間',mtError, [mbOK],0);

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


    if not bG_IsAmount then   //計算參數若為百分比
    begin
      for ii:=1 to 5 do
      begin
        sL_FieldNameAmount := 'AMOUNT_PERCENT' + IntToStr(ii);
        fL_TempValue := dsrSo111D.DataSet.FieldByName(sL_FieldNameAmount).AsFloat;
        if (fL_TempValue < 0) OR (fL_TempValue > 100) then
        begin
          MessageDlg('百分比須介於 0~100 之間',mtError, [mbOK],0);

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

    //至少要有1組有值
    if (dedCount_Percent1.DataSource.DataSet.FieldByName('SUBSCRIBER_COUNT_PERCENT1').AsString='') then
    begin
      MessageDlg('戶數級距1要有值',mtError, [mbOK],0);
      dedCount_Percent1.SetFocus;
      Exit;
    end;

    if (dedAmount_Percent1.DataSource.DataSet.FieldByName('AMOUNT_PERCENT1').AsString='') then
    begin
      MessageDlg('金額百分比1要有值',mtError, [mbOK],0);
      dedAmount_Percent1.SetFocus;
      //DBEdit14.DataSource.DataSet.FieldByName('AMOUNT_PERCENT1').FocusControl;
      Exit;
    end;

//SUBSCRIBER_COUNT_PERCENT 若有值 則相對應的AmountRant也要有值
    for ii:=1 to 5 do
    begin
      sL_FieldNameSubscriber:='SUBSCRIBER_COUNT_PERCENT'+intToStr(ii);
      sL_FieldNameAmount:='AMOUNT_PERCENT'+intToStr(ii);
      if (dtmMain2.cdsSo111D.FieldByName(sL_FieldNameSubscriber).AsFloat>0) and (dtmMain2.cdsSo111D.FieldByName(sL_FieldNameAmount).AsFloat=0) then
      begin
        MessageDlg('金額百分比 '+intToStr(ii)+' 不能為空值',mtError, [mbOK],0);
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

//選2: 百分比 爾後5欄位 其值須漸大(且小於100,可有小數)
//選2: 百分比 尚須判斷所輸入的值是否小於1(如果小於1 =>(fL_AmountRant) 大於1 (fL_AmountRant/100)
//選2: 百分比 且SUBSCRIBER_COUNT_PERCENT 最後輸入的欄位值必須為100 若不是 系統須將他設定為100 並
//提示相對應的 AmountRant  須要有值, 若USER 最後並不是在 SUBSCRIBER_COUNT_PERCENT5 輸入非100之值
//那麼請系統在user 最後輸入的下一個 SUBSCRIBER_COUNT_PERCENT 填入100 ,提示相對應的AmountRant須要有值
{
    if dtmMain2.cdsSo111D.FieldByName('RANGE_UNIT').AsString='2' then
    begin
      for ii:=1 to 5 do
      begin
        sL_FieldNameSubscriber:='SUBSCRIBER_COUNT_PERCENT'+intToStr(ii);
        if dtmMain2.cdsSo111D.FieldByName(sL_FieldNameSubscriber).AsFloat>100 then
        begin
        ShowMessage('戶數級距'+intToStr(ii)+'數量or百分比-不能大於一百');
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
          ShowMessage('金額百分比'+intToStr(jj)+'不能大於一百');
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

//判斷戶數比較單位 :1: 戶數2: 百分比
// 選1: 戶數  爾後5欄位 其值須漸大
//若10個欄位皆未輸入 可以 若第9個有值則第14個也要有值以此類推

    fL_PriorValue:=0;
    fL_CurValue:=0;


    Result:=True;
    for ii:=1 to 5 do //戶數級距-數量or百分比欄位
    begin
      sL_FieldNameSubscriber:='SUBSCRIBER_COUNT_PERCENT'+intToStr(ii);
        if dtmMain2.cdsSo111D.FieldByName(sL_FieldNameSubscriber).AsFloat<>0 then
          begin
            fL_CurValue:=dtmMain2.cdsSo111D.FieldByName(sL_FieldNameSubscriber).AsFloat;
            if ii<>1 then
              begin
                if fL_CurValue<fL_PriorValue then
                  begin
                    MessageDlg('戶數級距'+intToStr(ii)+'數量or百分比'+'-小於戶數級距'+intToStr(ii-1)+'數量or百分比',mtError, [mbOK],0);
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

//選擇百分比時

    if dtmMain2.cdsSo111D.FieldByName('RANGE_UNIT').AsString='2' then
    begin
      for ii:=1 to 5 do
      begin
        sL_FieldNameSubscriber:='SUBSCRIBER_COUNT_PERCENT'+intToStr(ii);
        sL_FieldNameAmount:='AMOUNT_PERCENT'+intToStr(ii);
        //撿查最後一個值 如果有輸入 SUBSCRIBER_COUNT_PERCENT5  一定要為100
        if (ii=5) and (dtmMain2.cdsSo111D.FieldByName(sL_FieldNameSubscriber).AsFloat<>100) and
           (dtmMain2.cdsSo111D.FieldByName(sL_FieldNameSubscriber).AsFloat<>0) then
        begin
          dtmMain2.cdsSo111D.FieldByName(sL_FieldNameSubscriber).AsFloat:=100;
          if dtmMain2.cdsSo111D.FieldByName(sL_FieldNameAmount).AsFloat=0 then
          begin
           MessageDlg('金額-百分比5 必須有值',mtError, [mbOK],0);

           if ii=5 then
           begin
              frmSO8A301.dedAmount_Percent5.SetFocus;
              Exit;
           end

          end;
        end;
//所輸入的值一定要有一個是100
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
                 MessageDlg('金額-百分比'+intToStr(ii+1)+'必須有值',mtError, [mbOK],0);
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
    self.Caption := frmMainMenu.setCaption('SO8A30','[明細編輯]');
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
    //比對此公式代碼代表哪一公式
    sL_RefNo := dtmMain2.getSO110RefNo(sI_Formula_ID);

    //只有公式一戶數比較單位才可變換,其他公式接不可變
    if sL_RefNo = '1' then         //兩階套餐制-比例制
    begin
      bG_IsAmount := false;
      dlcRangeUnit.Enabled := true;
    end
    else
    begin
      dlcRangeUnit.Enabled := false;

      if sL_RefNo = '2' then        //戶數數量級距-單價制
      begin
        dsrSo111D.DataSet.FieldByName('RANGE_UNIT').AsString := '1';
        bG_IsCount := true;
        bG_IsAmount := true;
      end
      else if sL_RefNo = '3' then   //戶數數量級距-比例制
      begin
        dsrSo111D.DataSet.FieldByName('RANGE_UNIT').AsString := '1';
        bG_IsCount := true;
        bG_IsAmount := false;
      end
      else if sL_RefNo = '4' then   //戶數百分比級距-單價制
      begin
        dsrSo111D.DataSet.FieldByName('RANGE_UNIT').AsString := '2';
        bG_IsCount := false;
        bG_IsAmount := true;
      end
      else if sL_RefNo = '5' then   //戶數百分比級距-比例制
      begin
        dsrSo111D.DataSet.FieldByName('RANGE_UNIT').AsString := '2';
        bG_IsCount := false;
        bG_IsAmount := false;
      end;
    end;

    if bG_IsCount then
    begin
      lblCount_Percent1.Caption := '戶數數量級距-1';
      lblCount_Percent2.Caption := '戶數數量級距-2';
      lblCount_Percent3.Caption := '戶數數量級距-3';
      lblCount_Percent4.Caption := '戶數數量級距-4';
      lblCount_Percent5.Caption := '戶數數量級距-5';
    end
    else
    begin
      lblCount_Percent1.Caption := '戶數百分比級距-1';
      lblCount_Percent2.Caption := '戶數百分比級距-2';
      lblCount_Percent3.Caption := '戶數百分比級距-3';
      lblCount_Percent4.Caption := '戶數百分比級距-4';
      lblCount_Percent5.Caption := '戶數百分比級距-5';
    end;


    if bG_IsAmount then
    begin
      lblAmount_Percent1.Caption := '金額-1';
      lblAmount_Percent2.Caption := '金額-2';
      lblAmount_Percent3.Caption := '金額-3';
      lblAmount_Percent4.Caption := '金額-4';
      lblAmount_Percent5.Caption := '金額-5';
    end
    else
    begin
      lblAmount_Percent1.Caption := '百分比-1';
      lblAmount_Percent2.Caption := '百分比-2';
      lblAmount_Percent3.Caption := '百分比-3';
      lblAmount_Percent4.Caption := '百分比-4';
      lblAmount_Percent5.Caption := '百分比-5';
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

    if sL_RangeUnit = '1' then   //戶數單位:數量
      bG_IsCount := true
    else                          //戶數單位:百分比
      bG_IsCount := false;

    if bG_IsCount then
    begin
      lblCount_Percent1.Caption := '戶數數量級距-1';
      lblCount_Percent2.Caption := '戶數數量級距-2';
      lblCount_Percent3.Caption := '戶數數量級距-3';
      lblCount_Percent4.Caption := '戶數數量級距-4';
      lblCount_Percent5.Caption := '戶數數量級距-5';
    end
    else
    begin
      lblCount_Percent1.Caption := '戶數百分比級距-1';
      lblCount_Percent2.Caption := '戶數百分比級距-2';
      lblCount_Percent3.Caption := '戶數百分比級距-3';
      lblCount_Percent4.Caption := '戶數百分比級距-4';
      lblCount_Percent5.Caption := '戶數百分比級距-5';
    end;

end;

end.
