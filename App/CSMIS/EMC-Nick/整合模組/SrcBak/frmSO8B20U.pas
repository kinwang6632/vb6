unit frmSO8B20U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, DB, Grids, DBGrids, ExtCtrls ,DBClient,
  fraChineseYMU, Buttons, scExcelExport, ComCtrls;

type
  TfrmSO8B20 = class(TForm)
    btnStartCalc: TButton;
    btnClose: TButton;
    Label2: TLabel;
    stxCom: TStaticText;
    rgpCompute: TRadioGroup;
    chbPage: TCheckBox;
    btnQueryComm: TButton;
    lbxBox: TListBox;
    Label4: TLabel;
    StatusBar1: TStatusBar;
    Timer1: TTimer;
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label3: TLabel;
    fraChineseYM1: TfraChineseYM;
    fraChineseYM2: TfraChineseYM;
    GroupBox2: TGroupBox;
    Label5: TLabel;
    Label6: TLabel;
    fraChineseYM3: TfraChineseYM;
    fraChineseYM4: TfraChineseYM;
    GroupBox4: TGroupBox;
    lblLastData4: TLabel;
    lblLastData5: TLabel;
    lblLastData6: TLabel;
    GroupBox3: TGroupBox;
    lblLastData1: TLabel;
    lblLastData2: TLabel;
    lblLastData3: TLabel;
    procedure FormShow(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure btnStartCalcClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure processCommission;
    procedure btnQueryCommClick(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure fraChineseYM2Exit(Sender: TObject);
    procedure fraChineseYM4Exit(Sender: TObject);

  private
    { Private declarations }
    sG_LastUser,sG_LastComputeYM,sG_LastUpdTime : String;
    nG_Hour,nG_Minute,nG_Sec : Integer;
    bG_HaveParamData : Boolean;
    function IsDataOk : Boolean;
    procedure ShowReport(bI_LockData,bI_IsPage : boolean;sI_Compute : String);
    procedure getLastData;
  public
    { Public declarations }
    //G_TempBoxItem:TStringList;
    //G_PromCode:TStringList;
    sG_PromCodeNoSQL : String;
  end;

var
  frmSO8B20: TfrmSO8B20;

implementation

uses dtmMain1U, Ustru, rptSO8B20AU, frmRptPreviewU, rptSO8B20B_1U, UCommonU,
  frmSO8B20_1U, frmDbSelectu,frmDbMultiSelectu, XLSFile, frmMainMenuU;

{$R *.dfm}

procedure TfrmSO8B20.FormShow(Sender: TObject);
var
    sL_LastUser,sL_LastComputeYM,sL_LastUpdTime : String;
begin
    self.Caption := frmMainMenu.setCaption('SO8B20','[�������p��]');
    fraChineseYM1.mseYM.SetFocus;

    //�d�X�Ӥ��q���ѼƸ��
    bG_HaveParamData := dtmMain1.getSo125ParamTemp(frmMainMenu.sG_CompCode);

    //���o�W���έp���H��,�~��,�ɶ�
    getLastData;
end;

procedure TfrmSO8B20.btnCloseClick(Sender: TObject);
begin
    dtmMain1.cdsSo121.EmptyDataSet;
    dtmMain1.cdsSo122.EmptyDataSet;
    Close;
end;

procedure TfrmSO8B20.btnStartCalcClick(Sender: TObject);
var
    sL_BuyYM,sL_RentYM,sL_FileCurrDateTime : String;
    dL_RetIntrCom,dL_RetAccNameCom,dL_RetWorNameCom:double;
    bL_HaveValue:Boolean;
    sL_SQL,sL_RealDate, sL_ComputeYM,sL_Compute ,sL_BelongYM,sL_CurrDateTime,sL_Time,sL_QueryString,sL_FileName:String;
    bL_IfLockYM ,bL_IsPage :Boolean;
    L_StrList:TStringList;
    dL_DateTime1,dL_DateTime2,dL_DiffDateTime : TTime;
    nL_Diff : Integer;
    sL_OnceComputeYM,sL_OnceBelongYM : String;

    procedure reComputeComm;
    begin
dL_DateTime1 := now;
      nG_Hour := 0;
      nG_Minute := 0;
      nG_Sec := 0;

      dtmMain1.bG_Repeat:=True;
//showmessage('920923_Jackal_�B�zBOX�R�_���');
      StatusBar1.SimpleText := '�B�zBOX�R�_���...';
      sL_BuyYM := dtmMain1.chineseYMChangeToEnglishYM( TUstr.ToROCYM( fraChineseYM1.getYM) );
      dtmMain1.dealwithSTBBuy(sL_BuyYM); //�p��]����X box �ҭn��������

//showmessage('920923_Jackal_�B�zBOX���θ��');
      //�אּ�ѿz��p��~��d��
      StatusBar1.SimpleText := '�B�zBOX���θ��...';
      sL_RentYM := sL_BuyYM;
      dtmMain1.dealwithSTBRent(sL_RentYM); //�p��]�����X box �ҭn��������


//showmessage('920923_Jackal_�B�zBOX������');
      //�B�z M ��
      StatusBar1.SimpleText := '�B�zBOX������...';
      dtmMain1.dealwithSTBMSno(sL_BuyYM);

      //�p��]�����s�I�O�W�D�ҭn��������
      processCommission;

    end;
begin
    if not bG_HaveParamData then
    begin
      MessageDlg('�������S��ѼƩ|���]�w����p��',mtError, [mbOK],0);
      Exit;
    end;


    if IsDataOk then  //�ҿ�J�~�륿�T
    begin
      if MessageDlg('�Х��T�{�������/���B��,�O�_�w�g�]�w����?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin
        //�N��������k�s�M��
        if not dtmMain1.cdsSo033Log.Active then
          dtmMain1.cdsSo033Log.CreateDataSet;
        dtmMain1.cdsSo033Log.EmptyDataSet;

        //�ثe�B�z�ɶ�
        dtmMain1.sG_CurDate := DateTimeToStr(now);


        //�����I����/STB �����p��~����k�ݦ~��
        sL_ComputeYM := TUstr.replaceStr(fraChineseYM1.getYM,'/','');
        sL_ComputeYM := TUstr.ToROCYM( sL_ComputeYM );
        dtmMain1.sG_ComputeYM := sL_ComputeYM;

        sL_BelongYM := TUstr.replaceStr(fraChineseYM2.getYM,'/','');
        sL_BelongYM := TUstr.ToROCYM( sL_BelongYM );
        dtmMain1.sG_BelongYM := sL_BelongYM;


        //�@�����I�����p��~����k�ݦ~��
        sL_OnceComputeYM := TUstr.replaceStr(fraChineseYM3.getYM,'/','');
        sL_OnceComputeYM := TUstr.ToROCYM( sL_OnceComputeYM );
        dtmMain1.sG_OnceComputeYM := sL_OnceComputeYM;
                    
        sL_OnceBelongYM := TUstr.replaceStr(fraChineseYM4.getYM,'/','');
        sL_OnceBelongYM := TUstr.ToROCYM( sL_OnceBelongYM );
        dtmMain1.sG_OnceBelongYM := sL_OnceBelongYM;


        bL_IfLockYM:=dtmMain1.CheckSo124LockYM(sL_ComputeYM);  //�ˬd���~��O�_�w�g��b


        if bL_IfLockYM then //�w�g��b
          MessageDlg('�����q�O�βέp�~��w�g��b,�L�k���s�p��!',mtError, [mbOK],0)
        else
        begin
          //���浥�ݪ��A
          TUCommonFun.setWaitingCursor;

          //���R���ӭp��~�몺���(�]�����\���ƭp��)
          //showmessage('920915_Jackal_z���R���ӭp��~�몺���');
          dtmMain1.deleteCommData(frmMainMenu.sG_CompCode);

          dtmMain1.qryCodeCD009.Open;

          //dtmMain1.cdsSo121.EmptyDataSet;
          dtmMain1.cdsSo122.EmptyDataSet;
          dtmMain1.cdsSo123.EmptyDataSet;
          dtmMain1.cdsSo134.EmptyDataSet;
          dtmMain1.cdsSo135.EmptyDataSet;

          if dtmMain1.cdsSo122ReturnMoney.State = dsInactive then
            dtmMain1.cdsSo122ReturnMoney.Open;
          dtmMain1.cdsSo122ReturnMoney.EmptyDataSet;

          if dtmMain1.cdsSo121ReturnMoney.State = dsInactive then
            dtmMain1.cdsSo121ReturnMoney.CreateDataSet;
          dtmMain1.cdsSo121ReturnMoney.EmptyDataSet;

          dtmMain1.activePriorData(frmMainMenu.sG_CompCode);

          //�p�����
          reComputeComm;

          //�N��Ʀs�J��Ʈw
          //dtmMain1.copyCdsSo121ToQry;
          //showmessage('920915_Jackal_�N��Ʀs�J��Ʈw');
          dtmMain1.applyToDB;
          dtmMain1.transReturnDataCdsIntoDB;

dL_DateTime2 := now;
dL_DiffDateTime := dL_DateTime2 - dL_DateTime1;
StatusBar1.SimpleText := '�@�� ' + TimeToStr(dL_DiffDateTime);
//showmessage(TimeToStr(dL_DiffDateTime));


          //�^�_�쪬�A
          TUCommonFun.setDefaultCursor;

          MessageDlg('�p�⧹��',mtInformation,[mbOK],0);
          if dtmMain1.cdsSo033Log.RecordCount > 0 then
          begin
            sL_FileCurrDateTime := DateTimeToStr(now);
            sL_CurrDateTime := dtmMain1.ReplaceStr(Trim(dtmMain1.ReplaceStr(sL_FileCurrDateTime,'/')),':');

            sL_QueryString := '�p��ɶ�:' + sL_FileCurrDateTime + '@' + '����H��:' + frmMainMenu.sG_User + '@' +
                              '���q�W��:' + frmMainMenu.sG_CompName +
                              '@' + '�έp�~��:' + sL_ComputeYM +
                              '@' + '�k�ݦ~��:' + sL_BelongYM ;


            sL_FileName := 'C:\���������`���_' + sL_BelongYM + '_' + sL_CurrDateTime + '.xls';

            dtmMain1.cdsSo033Log.DisableControls;
            DataSetToXLS(dtmMain1.cdsSo033Log,sL_FileName,sL_QueryString,'');
            dtmMain1.cdsSo033Log.EnableControls;

            MessageDlg('�����`���  ,�Ьd�� ' + sL_FileName,mtInformation,[mbOK],0);

            //���o�W���έp���H��,�~��,�ɶ�
            getLastData;
          end;
        end;
      end
      else
        Exit;
    end;
end;

function TfrmSO8B20.IsDataOk: Boolean;
var
    sL_ComputeYM,sL_BelongYM : String;
    nL_ComputeYM,nL_BelongYM : Integer;

begin
    // �W�D��������/STB����
    sL_ComputeYM := Trim(TUstr.replaceStr(fraChineseYM1.getYM,'/',''));
    if sL_ComputeYM = '' then
    begin
      MessageDlg('�п�J�έp�~��',mtError, [mbOK],0);
      fraChineseYM1.mseYM.SetFocus;
      Result:=False;
      exit;
    end
    else
      Result:=True;

    sL_BelongYM := Trim(TUstr.replaceStr(fraChineseYM2.getYM,'/',''));
    if sL_BelongYM = '' then
    begin
      MessageDlg('�п�J�k�ݦ~��',mtError, [mbOK],0);
      fraChineseYM2.mseYM.SetFocus;
      Result:=False;
      exit;
    end
    else
      Result:=True;

    nL_ComputeYM := StrToInt(sL_ComputeYM);
    nL_BelongYM := StrToInt(sL_BelongYM);

    if nL_ComputeYM >= nL_BelongYM then
    begin
      MessageDlg('�k�ݦ~�륲���j��έp�~��',mtError, [mbOK],0);
      fraChineseYM2.mseYM.SetFocus;
      Result:=False;
      exit;
    end
    else
      Result:=True;



    // �W�D�����@�����I
    sL_ComputeYM := Trim(TUstr.replaceStr(fraChineseYM3.getYM,'/',''));
    if sL_ComputeYM = '' then
    begin
      MessageDlg('�п�J�έp�~��',mtError, [mbOK],0);
      fraChineseYM3.mseYM.SetFocus;
      Result:=False;
      exit;
    end
    else
      Result:=True;


    sL_BelongYM := Trim(TUstr.replaceStr(fraChineseYM4.getYM,'/',''));
    if sL_BelongYM = '' then
    begin
      MessageDlg('�п�J�k�ݦ~��',mtError, [mbOK],0);
      fraChineseYM4.mseYM.SetFocus;
      Result:=False;
      exit;
    end
    else
      Result:=True;

    if nL_ComputeYM >= nL_BelongYM then
    begin
      MessageDlg('�k�ݦ~�륲���j��έp�~��',mtError, [mbOK],0);
      fraChineseYM4.mseYM.SetFocus;
      Result:=False;
      exit;
    end
    else
      Result:=True;
end;

procedure TfrmSO8B20.FormCreate(Sender: TObject);
var
    ii : Integer;
begin
    stxCom.Caption:=frmMainMenu.sG_CompName;
    //fraChineseYM1.setYM('091/08');
    for ii:=0 to dtmMain1.G_PromNameStrList.Count-1 do
      lbxBox.Items.Add(dtmMain1.G_PromNameStrList.Strings[ii]);
end;

procedure TfrmSO8B20.ShowReport(bI_LockData,bI_IsPage : boolean;sI_Compute : String);
var
    L_Rpt1 : TrptSO8B20A;
    L_Rpt2 : TrptSO8B20B_1;
    bL_LockData : boolean;
begin
    L_Rpt1 := TrptSO8B20A.Create(Application);
    L_Rpt1.bG_LockData := bI_LockData;
    L_Rpt1.bG_IsPage := bI_IsPage;
    L_Rpt1.sG_Compute := sI_Compute;
    bL_LockData := bI_LockData;
    L_Rpt1.Preview;
    bL_LockData := L_Rpt1.bG_LockData;

    L_Rpt1.Free;
    L_Rpt1 := nil;

    L_Rpt2 := TrptSO8B20B_1.Create(Application);
    L_Rpt2.bG_LockData := bL_LockData;
    L_Rpt2.Preview;

    L_Rpt2.Free;
    L_Rpt2 := nil;

end;

procedure TfrmSO8B20.processCommission;
var
    sL_YM, sL_YM2 : String;
begin
    dtmMain1.activeFormula;


    //�B�z�������I�W�D����
    sL_YM2 := dtmMain1.chineseYMChangeToEnglishYM(fraChineseYM3.getYM);
    sL_YM := TUstr.ToROCYM( sL_YM2 );
    dtmMain1.preProcessingData('So033',sL_YM,BATCH_PAY);
    dtmMain1.preProcessingData('So034',sL_YM,BATCH_PAY);

//showmessage('920923_Jackal_�B�z�@�����I�W�D');
    //�B�z�@�����I�W�D����(�V�e���X�Ӥ몺���)
    sL_YM2 := dtmMain1.chineseYMChangeToEnglishYM(fraChineseYM3.getYM);
    sL_YM := TUstr.ToROCYM( sL_YM2 );
    dtmMain1.preProcessingData('So033',sL_YM,ONCE_PAY);
    dtmMain1.preProcessingData('So034',sL_YM,ONCE_PAY);
end;


procedure TfrmSO8B20.btnQueryCommClick(Sender: TObject);
begin
    Application.CreateForm(TfrmSO8B20_1,frmSO8B20_1);
    frmSO8B20_1.ShowModal;
end;


procedure TfrmSO8B20.Timer1Timer(Sender: TObject);
var
    sL_Hour,sL_Minute,sL_Sec : String;
begin

    //�Ұʮɦ۰ʥ[�@��
    nG_Sec := nG_Sec + 1;

    //�ֶi60��۰ʶi�쬰�@��
    if nG_Sec = 60 then
    begin
      nG_Minute := nG_Minute + 1;
      nG_Sec := 0;
    end;

    //�ֶi60���۰ʶi�쬰�@�p��
    if nG_Minute = 60 then
    begin
      nG_Hour := nG_Hour + 1;
      nG_Minute := 0;
    end;


    //���Ƥ������Ʈ�
    if (nG_Sec div 10) = 0 then
      sL_Sec := '0' + IntToStr(nG_Sec)
    else
      sL_Sec := IntToStr(nG_Sec);


    //����Ƥ������Ʈ�
    if (nG_Minute div 10) = 0 then
      sL_Minute := '0' + IntToStr(nG_Minute)
    else
      sL_Minute := IntToStr(nG_Minute);


    //��p�ɤ������Ʈ�
    if (nG_Hour div 10) = 0 then
      sL_Hour := '0' + IntToStr(nG_Hour)
    else
      sL_Hour := IntToStr(nG_Hour);

end;

procedure TfrmSO8B20.fraChineseYM2Exit(Sender: TObject);
var
    sL_OnceBelongYM,sL_OnceCompputeYM : String;
begin
    fraChineseYM4.setYM(fraChineseYM2.getYM);

    sL_OnceBelongYM := Trim(dtmMain1.ReplaceStr(fraChineseYM4.getYM,'/'));
    if sL_OnceBelongYM <> '' then
    begin
      sL_OnceCompputeYM := dtmMain1.beforeMonth(sL_OnceBelongYM,dtmMain1.nG_BackMonth);

      if Length(sL_OnceCompputeYM) = 4 then
        sL_OnceCompputeYM := '0' + sL_OnceCompputeYM;

      sL_OnceCompputeYM := Copy(sL_OnceCompputeYM,1,4) + '/' + Copy(sL_OnceCompputeYM,5,2);

      fraChineseYM3.setYM(sL_OnceCompputeYM);
    end;
end;

procedure TfrmSO8B20.fraChineseYM4Exit(Sender: TObject);
var
    sL_OnceCompputeYM,sL_OnceBelongYM : String;
    nL_Len : Integer;
begin
    sL_OnceBelongYM := Trim(dtmMain1.ReplaceStr(fraChineseYM4.getYM,'/'));
    if sL_OnceBelongYM <> '' then
    begin
      sL_OnceCompputeYM := dtmMain1.beforeMonth(sL_OnceBelongYM,dtmMain1.nG_BackMonth);

      if Length(sL_OnceCompputeYM) = 4 then
        sL_OnceCompputeYM := '0' + sL_OnceCompputeYM;

      sL_OnceCompputeYM := Copy(sL_OnceCompputeYM,1,4) + '/' + Copy(sL_OnceCompputeYM,5,2);

      fraChineseYM3.setYM(sL_OnceCompputeYM);
    end;

end;



procedure TfrmSO8B20.getLastData;
var
    sL_LastUser,sL_LastComputeYM,sL_LastUpdTime : String;
begin
    //���o�W���έp���H��,�~��,�ɶ�
    dtmMain1.getLastData('SO122',frmMainMenu.sG_CompCode,sL_LastUser,sL_LastComputeYM,sL_LastUpdTime);
    sL_LastComputeYM := IntToStr( StrToInt( Copy( sL_LastComputeYM, 1, 3 ) ) + 1911 ) + Copy( sL_LastComputeYM, 4, 2 );
    lblLastData1.Caption := '�H��:' + sL_LastUser;
    lblLastData2.Caption := '�έp�~��:' + sL_LastComputeYM;
    lblLastData3.Caption := '�ɶ�:' + sL_LastUpdTime;


    //���o�W���έp���H��,�~��,�ɶ�
    dtmMain1.getLastData('SO134',frmMainMenu.sG_CompCode,sL_LastUser,sL_LastComputeYM,sL_LastUpdTime);
    sL_LastComputeYM := IntToStr( StrToInt( Copy( sL_LastComputeYM, 1, 3 ) ) + 1911 ) + Copy( sL_LastComputeYM, 4, 2 );
    lblLastData4.Caption := '�H��:' + sL_LastUser;
    lblLastData5.Caption := '�έp�~��:' + sL_LastComputeYM;
    lblLastData6.Caption := '�ɶ�:' + sL_LastUpdTime;

end;

end.
