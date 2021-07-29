unit frmSO8A50U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, fraChineseYMDU, StdCtrls, Buttons, DB, ExtCtrls, fraChineseYMU,
  Grids, DBGrids, scExcelExport;


type
  TfrmSO8A50 = class(TForm)
    Label1: TLabel;
    cobComp: TComboBox;
    btnStartCalculate: TBitBtn;
    btnCancel: TBitBtn;
    btnReset: TBitBtn;
    lblRealDate: TLabel;
    fraChineseYMD1: TfraChineseYMD;
    fraChineseYMD2: TfraChineseYMD;
    lblShouldDate: TLabel;
    fraChineseYMD3: TfraChineseYMD;
    fraChineseYMD4: TfraChineseYMD;
    dsrCodeCD039A: TDataSource;
    rgpQueryType: TRadioGroup;
    Label2: TLabel;
    lbxChargeItem: TListBox;
    btnChargeItem: TBitBtn;
    fraChineseYM1: TfraChineseYM;
    fraChineseYM2: TfraChineseYM;
    procedure FormCreate(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure rgpQueryTypeClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnChargeItemClick(Sender: TObject);
    procedure btnStartCalculateClick(Sender: TObject);
    procedure btnResetClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
    function IsDataOk : Boolean;
    procedure initialRealAndShouldDate;
  public
    { Public declarations }
    G_TargetFieldValueStrList1,G_TempChargeItem,G_ChargeCodeStrList : TStringList;
    nG_RealOrShouldDate : Integer;
    sG_ChargeItemSQL,sG_ChargeItemNameSQL : String;
    procedure showDetailExcel(sI_CompName,sI_ComputeYM,sI_ChargeItemNameSQL : String;nI_RealOrShouldDate : Integer);
  end;

var
  frmSO8A50: TfrmSO8A50;

implementation

uses UCommonU, UdateTimeu, frmDbMultiSelectu, Ustru, XLSFile,
  dtmMain2U, frmMainMenuU;

{$R *.dfm}

procedure TfrmSO8A50.FormCreate(Sender: TObject);
begin

    G_TargetFieldValueStrList1 := TStringList.Create;
    G_TempChargeItem := TStringList.Create;
    G_ChargeCodeStrList := TStringList.Create;


    TUCommonFun.AddObjectToComboBox(cobComp.Items, dsrCodeCD039A.DataSet,NOT INSERT_NO_DATA_ITEM,'CodeNo','Description');
    //���w��VB�ǨӤ�User���ݤ��q
    TUCommonFun.setComboDefaultNdx(cobComp, frmMainMenu.sG_CompCode);

    dtmMain2.cdsCodeCd039.Open;

    //�d�ߵe���ɶ���l��
    initialRealAndShouldDate;
end;

procedure TfrmSO8A50.btnCancelClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmSO8A50.rgpQueryTypeClick(Sender: TObject);
begin
    //�����ꦬ���B���������B
    if rgpQueryType.ItemIndex = REAL_DATE_TYPE then
    begin
      nG_RealOrShouldDate := REAL_DATE_TYPE;

      //�אּ�~��
      lblRealDate.Visible := true;
      fraChineseYM1.Visible := true;

      lblShouldDate.Visible := false;
      fraChineseYM2.Visible := false;
{
      lblShouldDate.Visible := false;
      fraChineseYMD3.Visible := false;
      fraChineseYMD4.Visible := false;

      lblRealDate.Visible := true;
      fraChineseYMD1.Visible := true;
      fraChineseYMD2.Visible := true;
}
    end
    else if rgpQueryType.ItemIndex = SHOULD_DATE_TYPE then
    begin
      nG_RealOrShouldDate := SHOULD_DATE_TYPE;

      //�אּ�~��
      lblRealDate.Visible := false;
      fraChineseYM1.Visible := false;

      lblShouldDate.Visible := true;
      fraChineseYM2.Visible := true;
{
      lblShouldDate.Visible := true;
      fraChineseYMD3.Visible := true;
      fraChineseYMD4.Visible := true;

      lblRealDate.Visible := false;
      fraChineseYMD1.Visible := false;
      fraChineseYMD2.Visible := false;
}
    end;

    fraChineseYM1.setYM('    /  ');
    fraChineseYM2.setYM('    /  ');


end;

procedure TfrmSO8A50.FormShow(Sender: TObject);
begin
    rgpQueryTypeClick(frmSO8A50);

    self.Caption := frmMainMenu.setCaption('SO8A50','[��b���Ӹ��]');    
end;

procedure TfrmSO8A50.btnChargeItemClick(Sender: TObject);
var
    L_TargetFieldNamesStrList : TStringList;
    sL_EmpNoSQL : String;
    jj : Integer;
begin
    dtmMain2.getCodeCD019;

    L_TargetFieldNamesStrList:=TStringList.Create;

    L_TargetFieldNamesStrList.Add('CODENO');
    L_TargetFieldNamesStrList.Add('DESCRIPTION');

    //���M�ũε����
    dtmMain2.cdsCodeCD019.FieldByName('CODENO').DisplayLabel := '���O���إN�X';
    dtmMain2.cdsCodeCD019.FieldByName('DESCRIPTION').DisplayLabel := '���O���ئW��';


    if SelectMultiRecords('���I�怜�O���ئW��', dtmMain2.cdsCodeCD019, 'CODENO;DESCRIPTION', ',', L_TargetFieldNamesStrList, G_TargetFieldValueStrList1,true,sL_EmpNoSQL) = mrOk then
    begin

      //���T�w�ɥ��M�ŵe��
      lbxChargeItem.Clear;
      sG_ChargeItemSQL := '';
      sG_ChargeItemNameSQL := '';

      for jj:=0 to G_TargetFieldValueStrList1.Count -1 do
      begin
        G_TempChargeItem := TUstr.ParseStrings(Trim(G_TargetFieldValueStrList1.Strings[jj]),',');

        if sG_ChargeItemSQL = '' then
        begin
          sG_ChargeItemSQL :=  G_TempChargeItem.Strings[0];
          sG_ChargeItemNameSQL := G_TempChargeItem.Strings[1];
        end
        else
        begin
          sG_ChargeItemSQL := sG_ChargeItemSQL + ',' + G_TempChargeItem.Strings[0];
          sG_ChargeItemNameSQL := sG_ChargeItemNameSQL + ',' + G_TempChargeItem.Strings[1];
        end;


        lbxChargeItem.Items.Add(G_TempChargeItem.Strings[1]);

        //�Ω�z��h�O�ɨS�諸���O����
        G_ChargeCodeStrList.Add(G_TempChargeItem.Strings[0]);
      end;

    end;

//SHOWMESSAGE(sG_ChargeItemSQL);

end;

function TfrmSO8A50.IsDataOk: Boolean;
var
    sL_StartDate,sL_ReplaceSDate,sL_EndDate,sL_ReplaceEDate : String;
    sL_BelongYM,sL_ReplaceBelongYM : String;
    nL_StartDate,nL_EndDate,nL_BelongYM : Integer;
    dL_StartDate,dL_sL_EndDate : TDate;

    sL_ComputeYM : String;
begin
    if sG_ChargeItemSQL = '' then
    begin
        MessageDlg('���I�怜�O����',mtError, [mbOK],0);
        btnChargeItem.SetFocus;
        IsDataOk := false;
        exit;
    end;



    //�אּ�~��
    if nG_RealOrShouldDate = REAL_DATE_TYPE then
    begin
      sL_ComputeYM := Trim(TUstr.replaceStr(fraChineseYM1.getYM,'/',''));
    end else
    if nG_RealOrShouldDate = SHOULD_DATE_TYPE then
    begin
      sL_ComputeYM := Trim(TUstr.replaceStr(fraChineseYM2.getYM,'/',''));
    end;

    if ( sL_ComputeYM = '' ) then
    begin

      if nG_RealOrShouldDate = REAL_DATE_TYPE then
      begin
        MessageDlg('�п�J�ꦬ�~��',mtError, [mbOK],0);
        fraChineseYM1.mseYM.SetFocus;
      end
      else if nG_RealOrShouldDate = SHOULD_DATE_TYPE then
      begin
        MessageDlg('�п�J�����~��',mtError, [mbOK],0);
        fraChineseYM2.mseYM.SetFocus;
      end;

      Result:=False;
      exit;
    end else
    begin
      sL_ComputeYM := TUstr.ToROCYM( sL_ComputeYM );
      Result := True;
    end;  

{
    if nG_RealOrShouldDate = REAL_DATE_TYPE then
    begin
      //���ꦬ���
      sL_StartDate := Trim(fraChineseYMD1.getYMD);
      sL_EndDate := Trim(fraChineseYMD2.getYMD);
    end
    else if nG_RealOrShouldDate = SHOULD_DATE_TYPE then
    begin
      //���������
      sL_StartDate := Trim(fraChineseYMD3.getYMD);
      sL_EndDate := Trim(fraChineseYMD4.getYMD);
    end;


    //sL_StartDate := Trim(fraChineseYMD1.getYMD);
    sL_ReplaceSDate := Trim(dtmMain2.ReplaceStr(sL_StartDate,'/'));
    nL_StartDate := Length(sL_ReplaceSDate);

    //SHOWMESSAGE(sL_StartDate);
    //SHOWMESSAGE(sL_ReplaceSDate);
    //SHOWMESSAGE(IntToStr(nL_StartDate));
    //�ˬd�ꦬ����l���
    if Trim(sL_ReplaceSDate) <> '' then
    begin
      if nL_StartDate <> 7 then
      begin
        MessageDlg('����п�J7���',mtError, [mbOK],0);

        if nG_RealOrShouldDate = REAL_DATE_TYPE then
        begin
          fraChineseYMD1.setYMD('');
          fraChineseYMD1.mseYMD.SetFocus;
        end
        else if nG_RealOrShouldDate = SHOULD_DATE_TYPE then
        begin
          fraChineseYMD3.setYMD('');
          fraChineseYMD3.mseYMD.SetFocus;
        end;


        IsDataOk := false;
        exit;
      end
      else
      begin
        //�ˬd����榡�O�_���~
        //showmessage(sL_StartDate);
        if TUdateTime.IsDateStr(sL_StartDate,'/') = false then
        begin
          MessageDlg('����榡���~',mtError, [mbOK],0);

          if nG_RealOrShouldDate = REAL_DATE_TYPE then
          begin
            fraChineseYMD1.setYMD('');
            fraChineseYMD1.mseYMD.SetFocus;
          end
          else if nG_RealOrShouldDate = SHOULD_DATE_TYPE then
          begin
            fraChineseYMD3.setYMD('');
            fraChineseYMD3.mseYMD.SetFocus;
          end;

          IsDataOk := false;
          exit;
        end
      end
    end
    else
    begin
      MessageDlg('�п�J�l���',mtError, [mbOK],0);

      if nG_RealOrShouldDate = REAL_DATE_TYPE then
      begin
        fraChineseYMD1.setYMD('');
        fraChineseYMD1.mseYMD.SetFocus;
      end
      else if nG_RealOrShouldDate = SHOULD_DATE_TYPE then
      begin
        fraChineseYMD3.setYMD('');
        fraChineseYMD3.mseYMD.SetFocus;
      end;

      IsDataOk := false;
      exit;
    end;



  //�ˬd�ꦬ����������
    //sL_EndDate := Trim(fraChineseYMD2.getYMD);
    sL_ReplaceEDate := Trim(dtmMain2.ReplaceStr(sL_EndDate,'/'));
    nL_EndDate := Length(sL_ReplaceEDate);
    if Trim(sL_ReplaceEDate) <> '' then
    begin
      if nL_EndDate <> 7 then
      begin
        MessageDlg('����п�J7���',mtError, [mbOK],0);

        if nG_RealOrShouldDate = REAL_DATE_TYPE then
        begin
          fraChineseYMD2.setYMD('');
          fraChineseYMD2.mseYMD.SetFocus;
        end
        else if nG_RealOrShouldDate = SHOULD_DATE_TYPE then
        begin
          fraChineseYMD4.setYMD('');
          fraChineseYMD4.mseYMD.SetFocus;
        end;

        IsDataOk := false;
        exit;
      end
      else
      begin
        //�ˬd����榡�O�_���~
        if TUdateTime.IsDateStr(sL_EndDate,'/') = false then
        begin
          MessageDlg('����榡���~',mtError, [mbOK],0);

          if nG_RealOrShouldDate = REAL_DATE_TYPE then
          begin
            fraChineseYMD2.setYMD('');
            fraChineseYMD2.mseYMD.SetFocus;
          end
          else if nG_RealOrShouldDate = SHOULD_DATE_TYPE then
          begin
            fraChineseYMD4.setYMD('');
            fraChineseYMD4.mseYMD.SetFocus;
          end;

          IsDataOk := false;
          exit;
        end
        else
          IsDataOk := true;
      end
    end
    else
    begin
      MessageDlg('�п�J�������',mtError, [mbOK],0);

      if nG_RealOrShouldDate = REAL_DATE_TYPE then
      begin
        fraChineseYMD2.setYMD('');
        fraChineseYMD2.mseYMD.SetFocus;
      end
      else if nG_RealOrShouldDate = SHOULD_DATE_TYPE then
      begin
        fraChineseYMD4.setYMD('');
        fraChineseYMD4.mseYMD.SetFocus;
      end;

      IsDataOk := false;
      exit;
    end;


    dL_StartDate := StrToDate(dtmMain2.TransToEngDate(sL_StartDate));
    dL_sL_EndDate := StrToDate(dtmMain2.TransToEngDate(sL_EndDate));
    if dL_StartDate > dL_sL_EndDate then
    begin
      MessageDlg('������������j�󵥩�l���',mtError, [mbOK],0);

      if nG_RealOrShouldDate = REAL_DATE_TYPE then
      begin
        fraChineseYMD2.setYMD('');
        fraChineseYMD2.mseYMD.SetFocus;
      end
      else if nG_RealOrShouldDate = SHOULD_DATE_TYPE then
      begin
        fraChineseYMD4.setYMD('');
        fraChineseYMD4.mseYMD.SetFocus;
      end;

      IsDataOk := false;
      exit;
    end;
}
end;

procedure TfrmSO8A50.btnStartCalculateClick(Sender: TObject);
var
    sL_CompCode,sL_ComputeYM,sL_Result : String;
    bL_HaveData : Boolean;
begin
    if IsDataOk then
    begin
      //���浥�ݪ��A
      TUCommonFun.setWaitingCursor;
//showmessage('1');
      if nG_RealOrShouldDate = REAL_DATE_TYPE then
      begin
        //���ꦬ�~��
        sL_ComputeYM := Trim(fraChineseYM1.getYM);
        sL_ComputeYM := TUstr.ToROCYM( sL_ComputeYM );
      end
      else if nG_RealOrShouldDate = SHOULD_DATE_TYPE then
      begin
        //�������~��
        sL_ComputeYM := Trim(fraChineseYM2.getYM);
        sL_ComputeYM := TUstr.ToROCYM( sL_ComputeYM );
      end;
//showmessage('2');


      //�ˬd�Ӧ~��ꦬ�������~��O�_�w�p��L
      bL_HaveData := dtmMain2.checkHaveSO131Data(sL_ComputeYM,nG_RealOrShouldDate);

      if bL_HaveData then
      begin
        //showmessage('�έp�~�릳���');
        if MessageDlg('�w����ƬO�_�n���s�p��?',mtConfirmation, [mbYes, mbNo], 0) = mrNo then
        begin
//showmessage('3');
          //showmessage('�����s�p��');
          //�^�_�쪬�A
          dtmMain2.getDetailData(frmMainMenu.sG_CompCode,sL_ComputeYM,nG_RealOrShouldDate);
//showmessage('4');
          //�নExcel
          showDetailExcel(frmMainMenu.sG_CompName,sL_ComputeYM,sG_ChargeItemNameSQL,nG_RealOrShouldDate);
//showmessage('5');
          TUCommonFun.setDefaultCursor;
//showmessage('6');
        end
        else
        begin
          //showmessage('���s�p��');
//showmessage('7');
          sL_Result := dtmMain2.calculateDetailData(frmMainMenu.sG_CompCode,sG_ChargeItemSQL,sL_ComputeYM,nG_RealOrShouldDate);
//          sL_Result := ''; //jacy testing...
          if sL_Result <> '' then
          begin

            MessageDlg(sL_Result,mtError, [mbOK],0);
            //�^�_�쪬�A
//showmessage('8');
            TUCommonFun.setDefaultCursor;
            exit;

          end
          else
          begin
//showmessage('9');
            dtmMain2.getDetailData(frmMainMenu.sG_CompCode,sL_ComputeYM,nG_RealOrShouldDate);

//showmessage('10');
            //�নExcel
            showDetailExcel(frmMainMenu.sG_CompName,sL_ComputeYM,sG_ChargeItemNameSQL,nG_RealOrShouldDate);
//showmessage('11');
            //�^�_�쪬�A
            TUCommonFun.setDefaultCursor;
            //showmessage('�p�⧹��');
//showmessage('12');
          end;
        end;
      end
      else
      begin
        //showmessage('�έp�~��L���');
//showmessage('13');
        sL_Result := dtmMain2.calculateDetailData(frmMainMenu.sG_CompCode,sG_ChargeItemSQL,sL_ComputeYM,nG_RealOrShouldDate);
        if sL_Result <> '' then
        begin
          MessageDlg(sL_Result,mtError, [mbOK],0);
          //�^�_�쪬�A
//showmessage('14');
          TUCommonFun.setDefaultCursor;
          exit;
        end
        else
        begin
//showmessage('15');
          dtmMain2.getDetailData(frmMainMenu.sG_CompCode,sL_ComputeYM,nG_RealOrShouldDate);
//showmessage('16');
          //�নExcel
          showDetailExcel(frmMainMenu.sG_CompName,sL_ComputeYM,sG_ChargeItemNameSQL,nG_RealOrShouldDate);
//showmessage('17');
          //�^�_�쪬�A
          TUCommonFun.setDefaultCursor;
          //showmessage('�p�⧹��');
        end;
      end;
    end;
end;


procedure TfrmSO8A50.btnResetClick(Sender: TObject);
begin
    //�e�����]
    lbxChargeItem.Clear;
    sG_ChargeItemSQL := '';
    sG_ChargeItemNameSQL := '';
    
    //�d�ߵe���ɶ���l��
    initialRealAndShouldDate;
end;

procedure TfrmSO8A50.initialRealAndShouldDate;
var
  wL_Year, wL_Month, wL_Day : Word;
  sL_StartDate,sL_EndDate : String;
  nL_LastDay : Integer;

begin
    //�אּ�~��
    fraChineseYM1.setYM('    /  ');
    fraChineseYM2.setYM('    /  ');



    //�}�l�ꦬ��� = �w�]���W�Ӥ몺�Ĥ@��
    //�����ꦬ��� = �w�]���W�Ӥ몺�̫�@��
    // �k �� �� �� = �w�]���{��������
{
    DecodeDate(Date,wL_Year, wL_Month, wL_Day);
    if wL_Month = 1 then
      sL_StartDate := IntToStr(wL_Year-1) + '/' + IntToStr(12) + '/01'
    else
      sL_StartDate := IntToStr(wL_Year) + '/' + IntToStr(wL_Month - 1 ) + '/01';

    nL_LastDay := TUdateTime.DaysOfMonth(StrToDate(sL_StartDate));
    if wL_Month = 1 then
      sL_EndDate := IntToStr(wL_Year-1) + '/' + IntToStr(12) + '/' + IntToStr(nL_LastDay)
    else
      sL_EndDate := IntToStr(wL_Year) + '/' + IntToStr(wL_Month - 1 ) + '/' + IntToStr(nL_LastDay);


    //fraChineseYMD1.setYMD(TUdateTime.CDateStr(StrToDate('2002/10/07'),9));
    //fraChineseYMD2.setYMD(TUdateTime.CDateStr(StrToDate('2002/10/07'),9));
    //�ꦬ���
    fraChineseYMD1.setYMD(TUdateTime.CDateStr(StrToDate(sL_StartDate),9));
    fraChineseYMD2.setYMD(TUdateTime.CDateStr(StrToDate(sL_EndDate),9));

    //�������
    fraChineseYMD3.setYMD(TUdateTime.CDateStr(StrToDate(sL_StartDate),9));
    fraChineseYMD4.setYMD(TUdateTime.CDateStr(StrToDate(sL_EndDate),9));
}
end;

procedure TfrmSO8A50.showDetailExcel(sI_CompName,sI_ComputeYM,sI_ChargeItemNameSQL : String;nI_RealOrShouldDate : Integer);
var
    sL_CurrDateTime,sL_FileName,sL_RealOrShouldDate : String;
    L_StrList : TStringList;
    L_ExcelExport : TscExcelExport;
    sL_QueryString,sL_QueryDate,sL_QueryUser : String;
    sL_QueryCompName,sL_QueryBasic,sL_QueryYM : String;
    sL_QueryCitemName,sL_FileCurrDateTime : String;
begin
    sL_FileCurrDateTime := DateTimeToStr(now);
    sL_CurrDateTime := dtmMain2.ReplaceStr(Trim(dtmMain2.ReplaceStr(sL_FileCurrDateTime,'/')),':');


    sL_QueryDate := '�p��ɶ�: ' + sL_FileCurrDateTime;
    sL_QueryUser := '�p��H��: ' + frmMainMenu.sG_User;
    sL_QueryCompName := '���q�W��: ' + sI_CompName;

    if nI_RealOrShouldDate = REAL_DATE_TYPE then
      sL_RealOrShouldDate := '�ꦬ���'
    else if nI_RealOrShouldDate = SHOULD_DATE_TYPE then
      sL_RealOrShouldDate := '�������';

    sL_QueryBasic := '�d�ߦ~����: ' + sL_RealOrShouldDate;
    sL_QueryYM := '�d�ߦ~��: ' + sI_ComputeYM;
    sL_QueryCitemName := '�d�ߦ��O����: ' + sI_ChargeItemNameSQL;

    sL_QueryString := sL_QueryDate + '@' + sL_QueryUser + '@' +
                      sL_QueryCompName + '@' + sL_QueryBasic + '@' +
                      sL_QueryYM + '@' + sL_QueryCitemName;

    sL_QueryString := TUstr.replaceStr(sL_QueryString,'�@','��');
//showmessage('Liga �`�N�ݳ�~~~    ' + sL_QueryString);

    //���ॿ�`���
    if dtmMain2.cdsSO131Excel.RecordCount <> 0 then
    begin
                          
      sL_FileName := 'C:\��b' + sL_CurrDateTime + '.xls';

      dtmMain2.cdsSO131Excel.DisableControls;
      DataSetToXLS(dtmMain2.cdsSO131Excel,sL_FileName,sL_QueryString,
        'computeym,shouldamt,realamt,RefNo,REALDATE,SHOULDDATE,REALSTARTDATE,REALSTOPDATE');
      dtmMain2.cdsSO131Excel.EnableControls;

      MessageDlg('�p�⧹��,��Ʀs�� ' + sL_FileName,mtInformation,[mbOK],0);
{
      L_ExcelExport := TscExcelExport.Create(Application);

      sL_CurrDateTime := DateTimeToStr(now);

      L_StrList := TStringList.Create;
      L_StrList.Add('�p��ɶ�: ' + sL_CurrDateTime);
      L_StrList.Add('�p��H��: ' + dtmMain2.sG_User);
      L_StrList.Add('���q�W��: ' + sI_CompName);

      if nI_RealOrShouldDate = REAL_DATE_TYPE then
        sL_RealOrShouldDate := '�ꦬ���'
      else if nI_RealOrShouldDate = SHOULD_DATE_TYPE then
        sL_RealOrShouldDate := '�������';

      L_StrList.Add('�d�ߦ~����: ' + sL_RealOrShouldDate);

      L_StrList.Add('�d�ߦ~��: ' + sI_ComputeYM);
      L_StrList.Add('�d�ߦ��O����: ' + sI_ChargeItemNameSQL);
      L_ExcelExport.HeaderText := L_StrList;

      //�]�w�C��
      L_ExcelExport.FontHeader.Size := 9;
      L_ExcelExport.FontTitles.Size := 9;
      L_ExcelExport.FontData.Size := 9;

      L_ExcelExport.FontHeader.Color := clBlue;
      L_ExcelExport.FontTitles.Color := clMaroon;
      L_ExcelExport.FontData.Color := clNavy;

      //L_ExcelExport.BorderData.BackColor := clSkyBlue;
      L_ExcelExport.BorderTitles.BackColor := clFuchsia;
      L_ExcelExport.BorderHeader.BackColor := clWhite;



      //���覡�������NCDS�ରEXCEL
      sL_CurrDateTime := dtmMain2.ReplaceStr(Trim(dtmMain2.ReplaceStr(sL_CurrDateTime,'/')),':');
      sL_FileName := 'C:\��b' + sL_CurrDateTime + '.xls';


      L_ExcelExport.Dataset := dtmMain2.cdsSO131;
      L_ExcelExport.ExportDataset(true);
      L_ExcelExport.ExcelVisible := false;
      L_ExcelExport.SaveAs(sL_FileName,ffXLS);

      L_ExcelExport.Free;
}
    end
    else
      MessageDlg('�d�L���',mtInformation,[mbOK],0);
      

    //�A��b�b���
    if dtmMain2.cdsOtherSO131Excel.RecordCount <> 0 then
    begin

      sL_FileName := 'C:\��b(�b�b)' + sL_CurrDateTime + '.xls';

      dtmMain2.cdsOtherSO131Excel.DisableControls;
      DataSetToXLS(dtmMain2.cdsOtherSO131Excel,sL_FileName,sL_QueryString,
        'computeym,shouldamt,realamt,RefNo,REALDATE,SHOULDDATE,REALSTARTDATE,REALSTOPDATE');
      dtmMain2.cdsOtherSO131Excel.EnableControls;

      MessageDlg('���b�b���,�s�� ' + sL_FileName,mtInformation,[mbOK],0);
    end;
end;



procedure TfrmSO8A50.FormDestroy(Sender: TObject);
begin
    G_ChargeCodeStrList.Free;
    G_TargetFieldValueStrList1.Free;
    G_TempChargeItem.Free;
end;

end.
