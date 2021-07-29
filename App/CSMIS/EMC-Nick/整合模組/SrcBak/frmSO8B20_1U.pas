unit frmSO8B20_1U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, fraChineseYMU, ComCtrls, Buttons, Mask,
  Grids, DBGrids, scExcelExport, DB;

type
  TOrderByData = class(TObject)
    sColName : String;
  end;

  TfrmSO8B20_1 = class(TForm)
    pctMain: TPageControl;
    tbsQuery1: TTabSheet;
    tbsQuery2: TTabSheet;
    Label2: TLabel;
    Label3: TLabel;
    stxCom1: TStaticText;
    fraChineseYM1: TfraChineseYM;
    chbPage: TCheckBox;
    rgpCompute: TRadioGroup;
    btnQueryMaster: TButton;
    btnMasterClose: TButton;
    Label1: TLabel;
    Label4: TLabel;
    stxCom2: TStaticText;
    fraChineseYM2: TfraChineseYM;
    btnQueryDetail: TButton;
    btnDetailClose: TButton;
    Label5: TLabel;
    btnWorker: TBitBtn;
    lbxMen: TListBox;
    Label6: TLabel;
    medMan: TMaskEdit;
    btnReset: TButton;
    lbxWorkerCode: TListBox;
    Label7: TLabel;
    btnSales: TBitBtn;
    lbxSales: TListBox;
    Label8: TLabel;
    lbxOtherCompCode: TListBox;
    btnOtherCompCode: TBitBtn;
    btnOrder: TBitBtn;
    lbxOrderList: TListBox;
    Label9: TLabel;
    TabSheet1: TTabSheet;
    Label10: TLabel;
    Label11: TLabel;
    fraChineseYM3: TfraChineseYM;
    lbxChargeItem: TListBox;
    btnItem: TBitBtn;
    btnQueryExcel: TButton;
    btnExcelExit: TButton;
    btnReset1: TButton;
    GroupBox1: TGroupBox;
    chbChannelComm_C: TCheckBox;
    chbChannelComm_I: TCheckBox;
    chbChannelComm_A: TCheckBox;
    GroupBox2: TGroupBox;
    chbBoxComm_W: TCheckBox;
    chbBoxComm_I: TCheckBox;
    chbBoxComm_A: TCheckBox;
    rgpType: TRadioGroup;
    chbStatisticExcel: TCheckBox;
    rgpPayTypeRpt: TRadioGroup;
    rgpClassRpt: TRadioGroup;
    TabSheet2: TTabSheet;
    Label12: TLabel;
    fraChineseYM4: TfraChineseYM;
    rgpRptType: TRadioGroup;
    chbReport: TCheckBox;
    chbExcel: TCheckBox;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    TabSheet3: TTabSheet;
    Button4: TButton;
    Button5: TButton;
    Button6: TButton;
    Label13: TLabel;
    medYear: TMaskEdit;
    rgpType2: TRadioGroup;
    lbxChargeItem1: TListBox;
    btnItem1: TBitBtn;
    Label14: TLabel;
    chkWorkerComm1: TCheckBox;
    chkWorkerComm2: TCheckBox;
    Label15: TLabel;
    edtPath2: TEdit;
    edtPath1: TEdit;
    Label16: TLabel;
    procedure btnMasterCloseClick(Sender: TObject);
    procedure tbsQuery1Show(Sender: TObject);
    procedure btnQueryMasterClick(Sender: TObject);
    procedure btnDetailCloseClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure btnResetClick(Sender: TObject);
    procedure btnWorkerClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnSalesClick(Sender: TObject);
    procedure btnOtherCompCodeClick(Sender: TObject);
    procedure btnQueryDetailClick(Sender: TObject);
    procedure tbsQuery2Show(Sender: TObject);
    procedure btnOrderClick(Sender: TObject);
    procedure btnItemClick(Sender: TObject);
    procedure btnExcelExitClick(Sender: TObject);
    procedure btnQueryExcelClick(Sender: TObject);
    procedure btnReset1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure rgpClassRptEnter(Sender: TObject);
    procedure rgpClassRptClick(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure TabSheet3Show(Sender: TObject);
    procedure btnItem1Click(Sender: TObject);
    procedure pctMainChange(Sender: TObject);
    procedure TabSheet1Show(Sender: TObject);

  private
    { Private declarations }
    sG_Man : String;
    nG_ControlItem : Integer; 
    procedure ShowReportA(bI_LockData,bI_IsPage : boolean;sI_Compute,sI_BelongYM,sI_ManList : String);
    procedure ShowReportB(sI_PayType,sI_RptClass,sI_Compute,sI_BelongYM,sI_EmpNoSQL,sI_SalesCodeSQL,sI_OtherCompCodeSQL,sI_OrderList,sI_OrderNameList,sI_CodeNoSQL,sI_ChargeNameSQL : String;bI_IncludeWorkerComm : Boolean);
    function IsDataOk1 : Boolean;
    function IsDataOk2 : Boolean;
    function IsDataOK3 : Boolean;
    function IsDataOK4 : Boolean;
    function IsDataOK5 : Boolean;
  public
    { Public declarations }
    G_TargetFieldValueStrList1 : TStringList;
    G_TempEmpNo,G_TempSalesCode,G_TempOtherCompCode,G_TempChargeItem : TStringList;
    sG_EmpNoSQL,sG_SalesCodeSQL,sG_OtherCompCodeSQL,sG_CodeNoSQL,sG_ChargeNameSQL: String;
    sG_OrderList,sG_OrderNameList : String;
  end;

var
  frmSO8B20_1: TfrmSO8B20_1;

implementation

uses dtmMain1U, rptSO8B20AU, Ustru, frmDbMultiSelectu,
  frmDualListDlgU, XLSFile, rptSO8B20B_1U, rptSO8B20B_3U, rptSO8B20B_2U,
  rptSO8B20B_4U, frmMainMenuU, UCommonU;

{$R *.dfm}

procedure TfrmSO8B20_1.btnMasterCloseClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmSO8B20_1.ShowReportA(bI_LockData, bI_IsPage: boolean;
  sI_Compute,sI_BelongYM,sI_ManList: String);
var
    L_Rpt1 : TrptSO8B20A;
    bL_LockData : boolean;
begin
    L_Rpt1 := TrptSO8B20A.Create(Application);
    L_Rpt1.bG_LockData := bI_LockData;
    L_Rpt1.bG_IsPage := bI_IsPage;
    L_Rpt1.sG_Compute := sI_Compute;
    L_Rpt1.sG_BelongYM := sI_BelongYM;
    L_Rpt1.sG_ManList := sI_ManList;

    bL_LockData := bI_LockData;

    L_Rpt1.Preview;
    bL_LockData := L_Rpt1.bG_LockData;

    L_Rpt1.Free;
    L_Rpt1 := nil;

end;

procedure TfrmSO8B20_1.tbsQuery1Show(Sender: TObject);
begin
    stxCom1.Caption := frmMainMenu.sG_CompName;
    fraChineseYM1.mseYM.SetFocus;
end;

procedure TfrmSO8B20_1.btnQueryMasterClick(Sender: TObject);
var
    sL_Compute,sL_BelongYM,sL_Man,sL_ManList : String;
    bL_IsPage,bL_LockData : Boolean;
    ii : Integer;
begin
    if IsDataOk1 then
    begin
      bL_LockData := false;
      sL_BelongYM := Trim(TUstr.replaceStr(fraChineseYM1.getYM,'/',''));
      sL_BelongYM := TUstr.ToROCYM( sL_BelongYM );

      if rgpCompute.ItemIndex = 0 then         //�̤��q�p��
        sL_Compute := '0'
      else if rgpCompute.ItemIndex = 1 then   //�̳����p��
        sL_Compute := '1';

      if chbPage.Checked then    //�O�_�̭p��̾ڤ���
        bL_IsPage := true
      else
        bL_IsPage := false;

      
      sL_ManList := '';
      for ii:=0 to lbxMen.Count-1 do
      begin
        sL_Man := lbxMen.Items[ii];
        if sL_ManList ='' then
        begin
          sL_ManList := '''' + sL_Man + '''';
        end
        else
        begin
          sL_ManList := sL_ManList + ',''' + sL_Man + '''';
        end;
      end;

      ShowReportA(bL_LockData,bL_IsPage,sL_Compute,sL_BelongYM,sL_ManList);
    end;
end;

function TfrmSO8B20_1.IsDataOk1: Boolean;
var
    sL_BelongYM : String;
begin
    sL_BelongYM := Trim(TUstr.replaceStr(fraChineseYM1.getYM,'/',''));

    if sL_BelongYM = '' then
    begin
      MessageDlg('�п�J�k�ݦ~��',mtError, [mbOK],0);
      fraChineseYM1.mseYM.SetFocus;
      Result := false;
      exit;
    end;
    Result := True;
end;

procedure TfrmSO8B20_1.btnDetailCloseClick(Sender: TObject);
begin
    Close;
end;



procedure TfrmSO8B20_1.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
    sL_Man : String;
begin
  if Key = VK_RETURN then
  begin
    if medMan.EditText <> '' then
    begin
      sL_Man := Trim(medMan.EditText);
      lbxMen.Items.Add(sL_Man);
      medMan.EditText := '';
      medMan.SetFocus;
    end;
  end;
end;

procedure TfrmSO8B20_1.btnResetClick(Sender: TObject);
begin
    fraChineseYM1.setYM('');
    medMan.EditMask := '';
    lbxMen.Items.Clear;
    fraChineseYM1.mseYM.SetFocus;
end;

procedure TfrmSO8B20_1.btnWorkerClick(Sender: TObject);
var
    L_TargetFieldNamesStrList : TStringList;
    sL_EmpNoSQL : String;
    jj : Integer;
begin
    dtmMain1.getCM003Data;

    L_TargetFieldNamesStrList:=TStringList.Create;

    L_TargetFieldNamesStrList.Add('EmpNo');
    L_TargetFieldNamesStrList.Add('EmpName');

    //���M�ũε����
    dtmMain1.qryCM003.FieldByName('EmpNo').DisplayLabel := '���u�N�X';
    dtmMain1.qryCM003.FieldByName('EmpName').DisplayLabel := '���u�W��';


  if SelectMultiRecords('���I����u�W��', dtmMain1.qryCM003, 'EmpNo;EmpName', ',', L_TargetFieldNamesStrList, G_TargetFieldValueStrList1,true,sL_EmpNoSQL) = mrOk then
  begin

    //���T�w�ɥ��M�ŵe��
    lbxWorkerCode.Clear;
    sG_EmpNoSQL := '';

    for jj:=0 to G_TargetFieldValueStrList1.Count -1 do
    begin
      G_TempEmpNo := TUstr.ParseStrings(Trim(G_TargetFieldValueStrList1.Strings[jj]),',');

      if sG_EmpNoSQL = '' then
        sG_EmpNoSQL :=  '''' + G_TempEmpNo.Strings[0] + ''''
      else
        sG_EmpNoSQL := sG_EmpNoSQL + ',''' + G_TempEmpNo.Strings[0] + '''';


      lbxWorkerCode.Items.Add(G_TempEmpNo.Strings[1]);
    end;


  end;
//showmessage(sL_EmpNoSQL);

end;

procedure TfrmSO8B20_1.FormCreate(Sender: TObject);
begin
  G_TargetFieldValueStrList1 := TStringList.Create;

  G_TempEmpNo := TStringList.Create;
  G_TempSalesCode := TStringList.Create;
  G_TempOtherCompCode := TStringList.Create;
  sG_CodeNoSQL:= '';
  sG_EmpNoSQL := '';
  sG_SalesCodeSQL := '';
  sG_OtherCompCodeSQL := '';
end;

procedure TfrmSO8B20_1.btnSalesClick(Sender: TObject);
var
    L_TargetFieldNamesStrList : TStringList;
    sL_SalesCodeSQL : String;
    jj : Integer;
begin
    dtmMain1.getSO013Data;

    L_TargetFieldNamesStrList:=TStringList.Create;

    L_TargetFieldNamesStrList.Add('IntroId');
    L_TargetFieldNamesStrList.Add('NameP');

    //���M�ũε����
    dtmMain1.qrySO013.FieldByName('IntroId').DisplayLabel := '�N�X';
    dtmMain1.qrySO013.FieldByName('NameP').DisplayLabel := '�W��';


  if SelectMultiRecords('���I��P���I', dtmMain1.qrySO013, 'IntroId;NameP', ',', L_TargetFieldNamesStrList, G_TargetFieldValueStrList1,true,sL_SalesCodeSQL) = mrOk then
  begin
    //���T�w�ɥ��M�ŵe��
    lbxSales.Clear;
    sG_SalesCodeSQL := '';

    for jj:=0 to G_TargetFieldValueStrList1.Count -1 do
    begin
      G_TempSalesCode := TUstr.ParseStrings(Trim(G_TargetFieldValueStrList1.Strings[jj]),',');

      if sG_SalesCodeSQL = '' then
        sG_SalesCodeSQL :=  '''' + G_TempSalesCode.Strings[0] + ''''
      else
        sG_SalesCodeSQL := sG_SalesCodeSQL + ',''' + G_TempSalesCode.Strings[0] + '''';

      lbxSales.Items.Add(G_TempSalesCode.Strings[1]);
    end;
  end;
end;

procedure TfrmSO8B20_1.btnOtherCompCodeClick(Sender: TObject);
var
    L_TargetFieldNamesStrList : TStringList;
    sL_OtherCompCodeSQL : String;
    jj : Integer;
begin
    dtmMain1.getOtherCompCodeData;

    L_TargetFieldNamesStrList:=TStringList.Create;

    L_TargetFieldNamesStrList.Add('IntAccId');
    L_TargetFieldNamesStrList.Add('IntAccName');

    //���M�ũε����
    dtmMain1.qryOtherCompCode.FieldByName('IntAccId').DisplayLabel := '�N�X';
    dtmMain1.qryOtherCompCode.FieldByName('IntAccName').DisplayLabel := '�W��';


  if SelectMultiRecords('���I���������u', dtmMain1.qryOtherCompCode, 'IntAccId;IntAccName', ',', L_TargetFieldNamesStrList, G_TargetFieldValueStrList1,true,sL_OtherCompCodeSQL) = mrOk then
  begin
    //���T�w�ɥ��M�ŵe��
    lbxOtherCompCode.Clear;
    sG_OtherCompCodeSQL := '';

    for jj:=0 to G_TargetFieldValueStrList1.Count -1 do
    begin
      G_TempOtherCompCode := TUstr.ParseStrings(Trim(G_TargetFieldValueStrList1.Strings[jj]),',');

      if sG_OtherCompCodeSQL = '' then
        sG_OtherCompCodeSQL :=  '' + G_TempOtherCompCode.Strings[0] + ''
      else
        sG_OtherCompCodeSQL := sG_OtherCompCodeSQL + ',' + G_TempOtherCompCode.Strings[0] + '';

      lbxOtherCompCode.Items.Add(G_TempOtherCompCode.Strings[1]);
    end;
  end;
end;



function TfrmSO8B20_1.IsDataOk2: Boolean;
var
    sL_BelongYM : String;
begin
    sL_BelongYM := Trim(TUstr.replaceStr(fraChineseYM2.getYM,'/',''));

    if sL_BelongYM = '' then
    begin
      MessageDlg('�п�J�k�ݦ~��',mtError, [mbOK],0);
      fraChineseYM2.mseYM.SetFocus;
      Result := false;
      exit;
    end;

    //�������Ӥβέp��ݭn��ܦ��O�W�D����
    if ((sG_CodeNoSQL='') and (rgpClassRpt.ItemIndex<>1)) then
    begin
      MessageDlg('�п�ܦ��O�W�D����',mtError, [mbOK],0);
      lbxChargeItem1.SetFocus;
      Result := false;
      exit;
    end;

    //�ƧǤ覡
    if (sG_OrderList='') AND (rgpClassRpt.ItemIndex<>2) then
    begin
      MessageDlg('���I��ƧǤ覡',mtError, [mbOK],0);
      btnOrder.SetFocus;
      Result := false;
      exit;
    end;
 Result := True;
end;

procedure TfrmSO8B20_1.btnQueryDetailClick(Sender: TObject);
var
    sL_BelongYM,sL_PayType,sL_RptClass : String;
    bL_IncludeWorkerComm : Boolean;
begin
    if IsDataOk2 then
    begin
      //���浥�ݪ��A
      TUCommonFun.setWaitingCursor;

      sL_BelongYM := Trim(TUstr.replaceStr(fraChineseYM2.getYM,'/',''));
      sL_BelongYM := TUstr.ToROCYM( sL_BelongYM );
      //sG_EmpNoSQL,sG_SalesCodeSQL,sG_OtherCompCodeSQL

      if rgpPayTypeRpt.ItemIndex = 0 then
        sL_PayType := ONCE_PAY
      else if rgpPayTypeRpt.ItemIndex = 1 then
        sL_PayType := BATCH_PAY;

      if rgpClassRpt.ItemIndex = 0 then
        sL_RptClass := CHANNEL_DATA
      else if rgpClassRpt.ItemIndex = 1 then
        sL_RptClass := BOX_DATA
      else if rgpClassRpt.ItemIndex = 2 then
        sL_RptClass := STATISTIC_DATA;

      //�έp��ƬO�n�]�t�u�{�H������
      if chkWorkerComm1.Checked then
        bL_IncludeWorkerComm := true
      else
        bL_IncludeWorkerComm := false;


      ShowReportB(sL_PayType,sL_RptClass,frmMainMenu.sG_CompCode, sL_BelongYM, sG_EmpNoSQL,
        sG_SalesCodeSQL, sG_OtherCompCodeSQL,sG_OrderList,sG_OrderNameList,sG_CodeNoSQL,sG_ChargeNameSQL,bL_IncludeWorkerComm);


      //�^�_�쪬�A
      TUCommonFun.setDefaultCursor;
    end;
end;

procedure TfrmSO8B20_1.ShowReportB(sI_PayType,sI_RptClass,sI_Compute, sI_BelongYM, sI_EmpNoSQL,
  sI_SalesCodeSQL, sI_OtherCompCodeSQL,sI_OrderList,sI_OrderNameList,sI_CodeNoSQL,sI_ChargeNameSQL : String;bI_IncludeWorkerComm : Boolean);
var
    L_Rpt1 : TrptSO8B20B_1;
    L_Rpt2 : TrptSO8B20B_2;
    L_Rpt3 : TrptSO8B20B_3;
begin
    if sI_RptClass = CHANNEL_DATA then
    begin
      //showmessage('Channel����');
      L_Rpt1 := TrptSO8B20B_1.Create(Application);
      L_Rpt1.sG_Compute := sI_Compute;
      L_Rpt1.sG_BelongYM := sI_BelongYM;
      L_Rpt1.sG_EmpNoSQL := sI_EmpNoSQL;
      L_Rpt1.sG_SalesCodeSQL := sI_SalesCodeSQL;
      L_Rpt1.sG_OtherCompCodeSQL := sI_OtherCompCodeSQL;
      L_Rpt1.sG_OrderList := sI_OrderList;
      L_Rpt1.sG_OrderNameList := sI_OrderNameList;
      L_Rpt1.sG_PayType := sI_PayType;
      L_Rpt1.sG_CodeNoSQL := sI_CodeNoSQL;
      L_Rpt1.sG_ChargeNameSQL := sI_ChargeNameSQL;

      L_Rpt1.Preview;

      L_Rpt1.Free;
      L_Rpt1 := nil;
    end
    else if sI_RptClass = BOX_DATA then
    begin
      //showmessage('BOX����');
      L_Rpt2 := TrptSO8B20B_2.Create(Application);
      L_Rpt2.sG_Compute := sI_Compute;
      L_Rpt2.sG_BelongYM := sI_BelongYM;
      L_Rpt2.sG_EmpNoSQL := sI_EmpNoSQL;
      L_Rpt2.sG_SalesCodeSQL := sI_SalesCodeSQL;
      L_Rpt2.sG_OtherCompCodeSQL := sI_OtherCompCodeSQL;
      L_Rpt2.sG_OrderList := sI_OrderList;
      L_Rpt2.sG_OrderNameList := sI_OrderNameList;
      L_Rpt2.sG_PayType := sI_PayType;
      L_Rpt2.bG_IncludeWorkerComm := bI_IncludeWorkerComm;
      L_Rpt2.Preview;

      L_Rpt2.Free;
      L_Rpt2 := nil;

    end
    else if sI_RptClass = STATISTIC_DATA then
    begin
      //showmessage('�έp�����');
      L_Rpt3 := TrptSO8B20B_3.Create(Application);

      L_Rpt3.sG_Compute := sI_Compute;
      L_Rpt3.sG_BelongYM := sI_BelongYM;
      L_Rpt3.sG_PayType := sI_PayType;
      L_Rpt3.sG_CodeNoSQL := sI_CodeNoSQL;
      L_Rpt3.sG_ChargeNameSQL := sI_ChargeNameSQL;
      L_Rpt3.bG_IncludeWorkerComm := bI_IncludeWorkerComm;

      L_Rpt3.Preview;

      L_Rpt3.Free;
      L_Rpt3 := nil;
    end

end;

procedure TfrmSO8B20_1.tbsQuery2Show(Sender: TObject);
begin
    stxCom2.Caption := frmMainMenu.sG_CompName;
    fraChineseYM2.mseYM.SetFocus;
end;

procedure TfrmSO8B20_1.btnOrderClick(Sender: TObject);
var
    sL_OrderColumn,sL_Order : String;
    ii : Integer;
    L_Obj : TOrderByData;
begin
    Application.CreateForm(TfrmDualListDlg,frmDualListDlg);
    frmDualListDlg.SrcList.Items.Clear;

    L_Obj := TOrderByData.Create;
    L_Obj.sColName := 'IntAccId';
    frmDualListDlg.SrcList.Items.AddObject('���ФH',L_Obj);

    if rgpClassRpt.ItemIndex = 0 then    //��������
    begin
      L_Obj.sColName := 'ClctEn';
      frmDualListDlg.SrcList.Items.AddObject('���O��',L_Obj);

      L_Obj.sColName := 'BillNo';
      frmDualListDlg.SrcList.Items.AddObject('��ڽs��',L_Obj);
    end
    else if rgpClassRpt.ItemIndex = 1 then    //��������
    begin
      L_Obj.sColName := 'WorkerEn1';
      frmDualListDlg.SrcList.Items.AddObject('�u�{��',L_Obj);

      L_Obj.sColName := 'Sno';
      frmDualListDlg.SrcList.Items.AddObject('���u�渹',L_Obj);
    end;
    L_Obj.sColName := 'MediaCode';
    frmDualListDlg.SrcList.Items.AddObject('���дC��',L_Obj);

    frmDualListDlg.ShowModal;

    //�N��^�Ӫ��ƧǤ覡,�[�JListBox
    lbxOrderList.Clear;
    sG_OrderList := '';
    sG_OrderNameList := '';
    for ii:=0 to dtmMain1.G_OrderStrList.Count-1 do
    begin
      sL_OrderColumn := dtmMain1.G_OrderStrList.Strings[ii];
      lbxOrderList.Items.Add(sL_OrderColumn);

      if sL_OrderColumn='���ФH' then
        sL_Order := 'IntAccId'
      else if sL_OrderColumn='�u�{��' then
        sL_Order := 'WorkerEn1'
      else if sL_OrderColumn='���O��' then
        sL_Order := 'ClctEn'
      else if sL_OrderColumn='���u�渹' then
        sL_Order := 'Sno'
      else if sL_OrderColumn='��ڽs��' then
        sL_Order := 'BillNo'
      else if sL_OrderColumn='���дC��' then
        sL_Order := 'MediaCode';

      if sG_OrderList='' then
        sG_OrderList := sL_Order
      else
        sG_OrderList := sG_OrderList + ',' + sL_Order;

      if sG_OrderNameList='' then
        sG_OrderNameList := sL_OrderColumn
      else
        sG_OrderNameList := sG_OrderNameList + ',' + sL_OrderColumn;
    end;
end;

procedure TfrmSO8B20_1.btnItemClick(Sender: TObject);
Var
  L_TargetFieldNamesStrList : TStringList;
  sL_CodeNoSQL : String;
  jj : Integer;
begin
  dtmMain1.getCD019Data;

  L_TargetFieldNamesStrList:=TStringList.Create;

  L_TargetFieldNamesStrList.Add('CODENO');
  L_TargetFieldNamesStrList.Add('DESCRIPTION');

  //���M�ũε����
  dtmMain1.qryCD019.FieldByName('CODENO').DisplayLabel := '���O���إN�X';
  dtmMain1.qryCD019.FieldByName('DESCRIPTION').DisplayLabel := '���O���ئW��';

  if SelectMultiRecords('���I�怜�O���ئW��', dtmMain1.qryCD019, 'CODENO;DESCRIPTION', ',', L_TargetFieldNamesStrList,
     G_TargetFieldValueStrList1,true,sL_CodeNoSQL) = mrOk then

  begin

    //���T�w�ɥ��M�ŵe��
    lbxChargeItem.Clear;
    sG_CodeNoSQL := '';
    sG_ChargeNameSQL := '';

    for jj:=0 to G_TargetFieldValueStrList1.Count -1 do
    begin
      G_TempChargeItem := TUstr.ParseStrings(Trim(G_TargetFieldValueStrList1.Strings[jj]),',');


      if sG_CodeNoSQL = '' then
         sG_CodeNoSQL :=  G_TempChargeItem.Strings[0]

      else
         sG_CodeNoSQL := sG_CodeNoSQL + ',' + G_TempChargeItem.Strings[0] ;

      lbxChargeItem.Items.Add(G_TempChargeItem.Strings[1]);

      if sG_ChargeNameSQL = '' then
         sG_ChargeNameSQL := G_TempChargeItem.Strings[1]
      else
         sG_ChargeNameSQL := sG_ChargeNameSQL + ',' + G_TempChargeItem.Strings[1] ;

    end;

  end;

//showmessage(sG_ChargeNameSQL);




end;

function TfrmSO8B20_1.IsDataOK3: Boolean;
var
    sL_BelongYM : String;
begin
   sL_BelongYM := Trim(TUstr.replaceStr(fraChineseYM3.getYM,'/',''));

    if sL_BelongYM = '' then
    begin
      MessageDlg('�п�J�k�ݦ~��',mtError, [mbOK],0);
      fraChineseYM3.mseYM.SetFocus;
      Result := false;
      exit;
    end;


    if (chbChannelComm_A.Checked=false OR chbChannelComm_I.Checked=false OR chbChannelComm_C.Checked=false)
       and (chbBoxComm_A.Checked=false OR chbBoxComm_I.Checked=false OR chbBoxComm_W.Checked=false)
       and (chbStatisticExcel.Checked=false) then
    begin
      MessageDlg('���I��z�n��X���O [����] �� [����] �� [�έp��] Excel',mtError, [mbOK],0);
      chbChannelComm_A.SetFocus;
      Result := false;
      exit;
    end;

    if (sG_CodeNoSQL='') and ((chbChannelComm_A.Checked=true) OR (chbChannelComm_I.Checked=true) OR (chbChannelComm_C.Checked=true) OR (chbStatisticExcel.Checked=true)) then
    begin
      MessageDlg('�п�ܦ��O�W�D����',mtError, [mbOK],0);
      lbxChargeItem.SetFocus;
      Result := false;
      exit;
    end;

    Result := true;

end;

procedure TfrmSO8B20_1.btnExcelExitClick(Sender: TObject);
begin
   Close;
end;

procedure TfrmSO8B20_1.btnQueryExcelClick(Sender: TObject);
var
    sL_BelongYM, sL_CODENO,sL_Data,sL_Flag,sL_CurrDateTime,sL_PayMode : String;
    bL_IncludeWorkerComm : Boolean;
begin
    if IsDataOk3 then
    begin
      //�N��������k�s�M��
       sL_Flag:='';
      if not dtmMain1.cdsSO122Channel.Active then
        dtmMain1.cdsSO122Channel.CreateDataSet;
      dtmMain1.cdsSO122Channel.EmptyDataSet;

      if not dtmMain1.cdsSO122BOX.Active then
        dtmMain1.cdsSO122BOX.CreateDataSet;
      dtmMain1.cdsSO122BOX.EmptyDataSet;

      if not dtmMain1.cdsStatisticExcel.Active then
        dtmMain1.cdsStatisticExcel.CreateDataSet;
      dtmMain1.cdsStatisticExcel.EmptyDataSet;

      sL_CurrDateTime := DateTimeToStr(now);
      sL_BelongYM := Trim(TUstr.replaceStr(fraChineseYM3.getYM,'/',''));
      sL_BelongYM := TUstr.ToROCYM( sL_BelongYM );

      //�έp��ƬO�n�]�t�u�{�H������
      if chkWorkerComm2.Checked then
        bL_IncludeWorkerComm := true
      else
        bL_IncludeWorkerComm := false;


      //���W�DExcel
      if ((chbChannelComm_A.Checked) OR (chbChannelComm_I.Checked) OR (chbChannelComm_C.Checked)) then
      begin
        //���浥�ݪ��A
        TUCommonFun.setWaitingCursor;

        //�˵��n��ܤ@���Τ���
        if rgpType.ItemIndex = 0 then       //�@�����I
        begin
          dtmMain1.getSO134Data(sL_BelongYM,frmMainMenu.sG_CompCode,sG_CodeNoSQL);
          sL_PayMode := ONCE_PAY;
        end
        else if rgpType.ItemIndex = 1 then       //�������I
        begin
          dtmMain1.getSO122Data1(CHANNEL_DATA,sL_BelongYM,frmMainMenu.sG_CompCode,sG_CodeNoSQL,bL_IncludeWorkerComm);
          sL_PayMode := BATCH_PAY;
        end;


        //�]��SO134�PSO122�W�D���ݪ����@��,�ҥH CDS �@��
        //������Ҧ�����
        if chbChannelComm_A.Checked then
        begin
          //���浥�ݪ��A
          TUCommonFun.setWaitingCursor;

          dtmMain1.TransToChannelExcel(sL_PayMode,CHANNEL_A,sL_BelongYM,sL_CurrDateTime);
        end;

        //��������ФH����
        if chbChannelComm_I.Checked then
        begin
          //���浥�ݪ��A
          TUCommonFun.setWaitingCursor;

          dtmMain1.TransToChannelExcel(sL_PayMode,CHANNEL_I,sL_BelongYM,sL_CurrDateTime);
        end;

        //��������O�H������
        if chbChannelComm_C.Checked then
        begin
          //���浥�ݪ��A
          TUCommonFun.setWaitingCursor;

          dtmMain1.TransToChannelExcel(sL_PayMode,CHANNEL_C,sL_BelongYM,sL_CurrDateTime);
        end;
      end;



      //�� BOX Excel
      if ((chbBOXComm_A.Checked) OR (chbBOXComm_I.Checked) OR (chbBOXComm_W.Checked)) then
      begin
        //���浥�ݪ��A
        TUCommonFun.setWaitingCursor;

        dtmMain1.getSO122Data1(BOX_DATA,sL_BelongYM,frmMainMenu.sG_CompCode,sG_CodeNoSQL,bL_IncludeWorkerComm);

        //������Ҧ�����
        if chbBoxComm_A.Checked then
          dtmMain1.TransToBoxExcel(BOX_A,sL_BelongYM,sL_CurrDateTime,bL_IncludeWorkerComm);

        //��������ФH����
        if chbBoxComm_I.Checked then
          dtmMain1.TransToBoxExcel(BOX_I,sL_BelongYM,sL_CurrDateTime,bL_IncludeWorkerComm);

        //������u�{�H������
        if chbBoxComm_W.Checked then
          dtmMain1.TransToBoxExcel(BOX_W,sL_BelongYM,sL_CurrDateTime,bL_IncludeWorkerComm);
      end;


      //�� �έp�� Excel
      if chbStatisticExcel.Checked then
      begin
        if rgpType.ItemIndex = 0 then       //�@�����I
          dtmMain1.getStatisticData(EXCEL_MODE,ONCE_PAY,frmMainMenu.sG_CompCode,sL_BelongYM,sG_CodeNoSQL,bL_IncludeWorkerComm)
        else if rgpType.ItemIndex = 1 then       //�������I
          dtmMain1.getStatisticData(EXCEL_MODE,BATCH_PAY,frmMainMenu.sG_CompCode,sL_BelongYM,sG_CodeNoSQL,bL_IncludeWorkerComm);


        //��X Excel
        dtmMain1.TransToStatisticExcel(sL_BelongYM,sL_CurrDateTime,bL_IncludeWorkerComm);
      end;
    end;
end;


procedure TfrmSO8B20_1.btnReset1Click(Sender: TObject);
begin
  fraChineseYM3.setYM('');
  fraChineseYM3.mseYM.SetFocus;
  sG_CodeNoSQL := '';
  sG_ChargeNameSQL := '';
  lbxChargeItem.Items.Clear;
  chbChannelComm_A.Checked:=false;
  chbBoxComm_A.Checked:=false;

end;

procedure TfrmSO8B20_1.FormShow(Sender: TObject);
begin
    self.Caption := frmMainMenu.setCaption('SO8B20_1','[��������Ƭd��]');

end;

procedure TfrmSO8B20_1.rgpClassRptEnter(Sender: TObject);
var
    nL_CurItem : Integer;
begin
    nL_CurItem := rgpClassRpt.ItemIndex;

    if nL_CurItem <> nG_ControlItem then
    begin
      lbxOrderList.Clear;
      dtmMain1.G_OrderStrList.Clear;
      sG_OrderList := '';
      sG_OrderNameList := '';
      nG_ControlItem := nL_CurItem;
    end;
end;

procedure TfrmSO8B20_1.rgpClassRptClick(Sender: TObject);
var
    nL_CurItem : Integer;
begin
    nL_CurItem := rgpClassRpt.ItemIndex;

    if nL_CurItem <> nG_ControlItem then
    begin
      lbxOrderList.Clear;
      dtmMain1.G_OrderStrList.Clear;
      sG_OrderList := '';
      sG_OrderNameList := '';
      nG_ControlItem := nL_CurItem;
    end;

end;

function TfrmSO8B20_1.IsDataOK4: Boolean;
var
    sL_BelongYM : String;
begin
    sL_BelongYM := Trim(TUstr.replaceStr(fraChineseYM4.getYM,'/',''));

    if sL_BelongYM = '' then
    begin
      MessageDlg('�п�J��Ǧ~��',mtError, [mbOK],0);
      fraChineseYM4.mseYM.SetFocus;
      Result := false;
      exit;
    end;
    Result := True;
end;

procedure TfrmSO8B20_1.Button2Click(Sender: TObject);
begin
  fraChineseYM4.setYM('');
  fraChineseYM4.mseYM.SetFocus;
end;

procedure TfrmSO8B20_1.Button3Click(Sender: TObject);
begin
    Close;
end;

procedure TfrmSO8B20_1.Button1Click(Sender: TObject);
var
    sL_FileCurrDateTime,sL_CurrDateTime,sL_BelongYM,sL_QueryString,sL_YM : String;
    sL_FileName : String;
begin
    if IsDataOk4 then
    begin
      if not dtmMain1.cdsTotalCommData.Active then
        dtmMain1.cdsTotalCommData.CreateDataSet;
      dtmMain1.cdsTotalCommData.EmptyDataSet;


      sL_FileCurrDateTime := DateTimeToStr(now);
      sL_CurrDateTime := dtmMain1.ReplaceStr(Trim(dtmMain1.ReplaceStr(sL_FileCurrDateTime,'/')),':');

      sL_BelongYM := Trim(TUstr.replaceStr(fraChineseYM4.getYM,'/',''));
      sL_BelongYM := TUstr.ToROCYM( sL_BelongYM );
      sL_YM := IntToStr(StrToInt(Copy(sL_BelongYM,1,3)))+'�~'+Copy(sL_BelongYM,4,2)+'��';

      if rgpRptType.ItemIndex = 0 then  //�~�׹w���`��
      begin
        showmessage('�~�׹w���`��');
        dtmMain1.computeTotalReport(sL_BelongYM,frmMainMenu.sG_CompCode);

        if chbReport.Checked then
        begin

        end;

        sL_FileName := 'C:\�W�D�����Ҧ�����_' + sL_BelongYM + '_' + sL_CurrDateTime + '.XLS';

        if chbExcel.Checked then
        begin
          if dtmMain1.cdsTotalCommData.RecordCount > 0 then
          begin
            sL_QueryString:='����W��:' + sL_YM + '_�w�I�`����' + '@' +
                            '�έp�ɶ�:' + sL_FileCurrDateTime + '@' + '����H��:' + frmMainMenu.sG_User +'@' +
                            '���q�W��:' + TUstr.replaceStr(frmMainMenu.sG_CompName,'�@','��') + '@' + '�k�ݦ~��:' + sL_YM ;

            dtmMain1.cdsTotalCommData.DisableControls;
            DataSetToXLS(dtmMain1.cdsTotalCommData,sL_FileName,sL_QueryString,'');
            dtmMain1.cdsTotalCommData.EnableControls;
          end;
        end;
      end
      else if rgpRptType.ItemIndex = 1 then  //������
      begin
        showmessage('������');


      end;
    end;
end;

procedure TfrmSO8B20_1.Button6Click(Sender: TObject);
begin
    Close;
end;

procedure TfrmSO8B20_1.Button5Click(Sender: TObject);
begin
    medYear.Text :='';
end;

procedure TfrmSO8B20_1.Button4Click(Sender: TObject);
var
    sL_Year : String;
    L_Rpt : TrptSO8B20B_4;
begin
    if IsDataOk5 then
    begin
      //���浥�ݪ��A
      TUCommonFun.setWaitingCursor;

      if not dtmMain1.cdsTotalMonthData.Active then
        dtmMain1.cdsTotalMonthData.CreateDataSet;
      dtmMain1.cdsTotalMonthData.EmptyDataSet;

      sL_Year := dtmMain1.addZero(Trim(medYear.Text),4);
      sL_Year := dtmMain1.addZero( IntToStr( StrToInt( sL_Year ) - 1911 ),3 );
      dtmMain1.CalculateTotalMonthData(frmMainMenu.sG_CompCode,sL_Year);

      if rgpType2.ItemIndex = 0 then  //�i�{����
      begin
        //�^�_�쪬�A
        TUCommonFun.setDefaultCursor;

        L_Rpt := TrptSO8B20B_4.Create(Application);
        L_Rpt.sG_Compute := frmMainMenu.sG_CompCode;
        L_Rpt.sG_CompName := frmMainMenu.sG_CompName;
        L_Rpt.sG_User := frmMainMenu.sG_User;
        L_Rpt.sG_BasicYear := IntToStr(StrToInt(sL_Year));
        L_Rpt.Preview;

        L_Rpt.Free;
        L_Rpt := nil;
      end
      else if rgpType2.ItemIndex = 1 then   //�i�{Excel
      begin
        //�^�_�쪬�A
        TUCommonFun.setDefaultCursor;

        dtmMain1.showTotalMonthDataExcel(sL_Year);
      end;
    end;
end;

function TfrmSO8B20_1.IsDataOK5: Boolean;
var
    sL_Year : String;
begin
    sL_Year := Trim(TUstr.replaceStr(medYear.Text,'/',''));

    if sL_Year = '' then
    begin
      MessageDlg('�п�J�p��~��',mtError, [mbOK],0);
      medYear.SetFocus;
      Result := false;
      exit;
    end;
    Result := True;

end;

procedure TfrmSO8B20_1.TabSheet3Show(Sender: TObject);
begin
    medYear.SetFocus;
    edtPath2.Text := dtmMain1.sG_ExcelSavePath;
end;

procedure TfrmSO8B20_1.btnItem1Click(Sender: TObject);
Var
  L_TargetFieldNamesStrList : TStringList;
  sL_CodeNoSQL : String;
  jj : Integer;
begin
  dtmMain1.getCD019Data;

  L_TargetFieldNamesStrList:=TStringList.Create;

  L_TargetFieldNamesStrList.Add('CODENO');
  L_TargetFieldNamesStrList.Add('DESCRIPTION');

  //���M�ũε����
  dtmMain1.qryCD019.FieldByName('CODENO').DisplayLabel := '���O���إN�X';
  dtmMain1.qryCD019.FieldByName('DESCRIPTION').DisplayLabel := '���O���ئW��';

  if SelectMultiRecords('���I�怜�O���ئW��', dtmMain1.qryCD019, 'CODENO;DESCRIPTION', ',', L_TargetFieldNamesStrList,
     G_TargetFieldValueStrList1,true,sL_CodeNoSQL) = mrOk then

  begin

    //���T�w�ɥ��M�ŵe��
    lbxChargeItem1.Clear;
    sG_CodeNoSQL := '';
    sG_ChargeNameSQL := '';

    for jj:=0 to G_TargetFieldValueStrList1.Count -1 do
    begin
      G_TempChargeItem := TUstr.ParseStrings(Trim(G_TargetFieldValueStrList1.Strings[jj]),',');


      if sG_CodeNoSQL = '' then
         sG_CodeNoSQL :=  G_TempChargeItem.Strings[0]

      else
         sG_CodeNoSQL := sG_CodeNoSQL + ',' + G_TempChargeItem.Strings[0] ;

      lbxChargeItem1.Items.Add(G_TempChargeItem.Strings[1]);

      if sG_ChargeNameSQL = '' then
         sG_ChargeNameSQL := G_TempChargeItem.Strings[1]
      else
         sG_ChargeNameSQL := sG_ChargeNameSQL + ',' + G_TempChargeItem.Strings[1] ;

    end;

  end;
end;

procedure TfrmSO8B20_1.pctMainChange(Sender: TObject);
begin
    sG_CodeNoSQL := '';
    sG_ChargeNameSQL := '';
end;

procedure TfrmSO8B20_1.TabSheet1Show(Sender: TObject);
begin
    fraChineseYM3.mseYM.SetFocus;
    edtPath1.Text := dtmMain1.sG_ExcelSavePath;
end;

end.

