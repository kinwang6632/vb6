unit frmInvA08U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Buttons, Mask, ComCtrls, DB, DBCtrls, Grids,
  DBGrids, ADODB, StrUtils, Printers , math, Menus,
  cxControls, cxContainer, cxEdit, cxLabel, cxTextEdit, cxCurrencyEdit,
  cxLookAndFeelPainters, cxButtons, cxProgressBar, dxSkinsCore,
  dxSkinsDefaultPainters;
  
type
  //�ŧi�C��ӳ��|�Ψ쪺�Ҧ����B�O��
  TarraySum = Class(Tobject)
    sL_A, sL_B, sL_C, sL_D      : Double;
    sL_E, sL_F, sL_G, sL_H      : Double;
    sL_I, sL_J, sL_K, sL_L      : Double;
    sL_M, sL_N, sL_O, sL_P      : Double;
    sL_Q, sL_R, sL_S, sL_T      : Double;
    sL_U, sL_V, sL_W, sL_X      : Double;
    sL_Y, sL_Z                  : Double;
    sL_AA, sL_BB, sL_CC, sL_DD  : Double;
    sL_EE, sL_FF, sL_GG, sL_HH  : Double;
    sL_II, sL_JJ, sL_KK, sL_LL  : Double;
    sL_MM, sL_NN, sL_OO, sL_PP  : Double;
    sL_QQ, sL_RR, sL_SS         : Double;
    sL_TT, sL_UU, sL_VV         : String;
    sL_WW, sL_XX, sL_YY, sL_ZZ  : Double;
    sL_CC2, sL_DD2, sL_EE2, sL_FF2: Double;
    sL_KK2, sL_LL2, sL_MM2      : Double; { �q�l�B�G�p�B�T�p, �s�|�v�i�Ʋέp }
    sL_GG2                      : Double; 
  end;

  TfrmInvA08 = class(TForm)
    Panel1: TPanel;
    Label5: TLabel;
    Panel2: TPanel;
    PCT1: TPageControl;
    TST_Edit: TTabSheet;
    TST_View: TTabSheet;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label6: TLabel;
    RadioGroup1: TRadioGroup;
    Label7: TLabel;
    CheckBox1: TCheckBox;
    Label8: TLabel;
    CBx_Month: TComboBox;
    Edt_MediaFilePath: TEdit;
    AQry_Inv: TADOQuery;
    MKE_Year: TEdit;
    ListBox1: TListBox;
    Panel3: TPanel;
    Lbl_show: TLabel;
    ProgressBar2: TcxProgressBar;
    Bit_Previous: TcxButton;
    Btn_Print: TcxButton;
    btnExit: TcxButton;
    Panel4: TPanel;
    BitBtn2: TcxButton;
    btnRedo: TcxButton;
    btnOK: TcxButton;
    Bevel1: TBevel;
    MKE_Start: TMaskEdit;
    MKE_End: TMaskEdit;
    PrinterSetupDialog1: TPrinterSetupDialog;
    AQry_Inv099: TADOQuery;
    lblCompany: TcxLabel;
    lblUser: TcxLabel;
    Bevel2: TBevel;
    Label9: TLabel;
    RadioGroup2: TRadioGroup;
    chkStampTax: TCheckBox;
    txtStampTax: TcxCurrencyEdit;
    chkPaper: TCheckBox;
    chkApplyMainInv: TCheckBox;
    chkDebugFile: TCheckBox;
    procedure FormShow(Sender: TObject);
    procedure btnRedoClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
    procedure Bit_PreviousClick(Sender: TObject);
    procedure MKE_StartExit(Sender: TObject);
    procedure Btn_PrintClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure chkStampTaxClick(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
  private
    { Private declarations }
    FApplyByMainInv: Boolean;
    FImportNote: String;
    FDebugList: TStringList;
    //�p��C��ӳ��D�ɤ��榡
    // �ܼ� - ����
    //P_RecNo - ��Ƶ���
    //P_Inv01BusID - �P��H�Τ@�s��
    //P_Inv001TaxID - �P��H�|�y�s��
    //P_Table - ��ƪ�W��
    //P_Select - ����ﶵ
    //P_MediaList - ��X�Ȧs�ܼ�
    Procedure SetMasterData(Var P_RecNo:String;P_Inv01BusID,P_Inv001TaxID,
              P_Table:String;P_Select:Integer;Var P_MediaList:TStringList);

    //�p��έp���B����
    // �ܼ� - ����
    //dselect - ���sql�y�k�A1���B�X�p�B2��Ƶ��ơB3�o��
    //dField - �ݲέp���B�����W��
    //dTable - ��ƪ�W��
    //dParma1 - �ѼƤ@�AIsObsolete�A�����@�o����ơA�ťթΦ���
    //dParma2 - �ѼƤG�AInvFormat1�A�o���榡1�A�ťթζ�J1�B2�B3
    //dParma3 - �ѼƤT�AInvFormat2�A�o���榡2�A�ťթζ�J1�B3
    //dParma4 - �Ѽƥ|�ABusinessId�A�Ȥ�νs�A�ťթ�1�νs���ťաB2�νs�����ť�
    //dParma5 - �ѼƤ��ATaxType�A�o���|�O�A�ťթΦ���
    function SetMediaData(dselect:integer;dField:String; dTable:String;
             dParma1, dParma2, dParma3, dParma4, dParma5:String):Double ;

    function GetCompIdList: String;

  public
    { Public declarations }
  end;

var
  frmInvA08: TfrmInvA08;
  fL_MediaFile : String ;//�D���x�s���|
  fL_MediaFile_B : String ;//�D�ɳƥ��x�s���|
  fL_MediaReport : String ;//�������x�s���|
  fL_MediaREport_B : String ;//�����ɳƥ��x�s���|
  sL_CreateMonthStart : String ;//�ӳ��_�l���
  sL_CreateMonthEnd : String ;//�ӳ�������
  sL_MakeMonth : String ;//�ӳ����
  sL_StampTaxt: String;
  FCompanyList : TList;

  aDebugPath, aDebugFileName: String;

implementation

uses frmMainU, frmInvD02_1U, dtmMainJU, dtmMainU, dtmSOU, cbUtilis;

{$R *.dfm}

function TfrmInvA08.GetCompIdList: String;
var
  aIndex: Integer;
begin
  Result := Format( '''%s''',[dtmMain.getCompID] );
  if ( FCompanyList.Count > 1 ) and ( RadioGroup2.ItemIndex = 1 ) then
  begin
    Result := EmptyStr;
    for aIndex := 0 to FCompanyList.Count - 1 do
    begin
      Result := Result + Format(' ''%s'' ',[PCompany(FCompanyList[aIndex]).CompanyId] );
      if ( aIndex < ( FCompanyList.Count - 1 ) ) then
        Result := ( Result + ',' );
    end;    
  end;
  Result := Format( ' ( %s ) ', [Result] );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA08.FormShow(Sender: TObject);
var
  aSQL : String;
begin
  FDebugList := TStringList.Create;
  PCT1.ActivePageIndex := 0;
  self.Caption := frmMain.GetFormTitleString('A08','���ʹC��ӳ�');
  lblCompany.Caption := Format( '%s ( %s )',
    [dtmMain.getCompID, dtmMain.getCompName] );
  lblUser.Caption := dtmMain.getLoginUser;

  aSQL := ' SELECT INVPARAM FROM INV003 WHERE COMPID = ' +
    QuotedStr( dtmMain.getCompID );

  Edt_MediaFilePath.Text := dtmMainJ.GetXMLAttribute( aSQL, 'MediaFilePath' );
  CBx_Month.ItemIndex := 0;
  MKE_Year.SetFocus;
  MKE_Year.Text := copy(datetostr(date()),1,4);
  CheckBox1.Enabled := dtmMain.GetLinkToMIS;

  FCompanyList := TList.Create;
  dtmMain.GetInvoiceCompany( FCompanyList );
  RadioGroup2.Enabled := FCompanyList.Count > 1;
  Label9.Enabled := FCompanyList.Count > 1;
  {}
  FImportNote := dtmMainj.OpenSQL( Format(
    ' select importnote from inv003  ' +
    '  where identifyid1 = ''%s''    ' +
    '    and identifyid2 = ''%s''    ' +
    '    and compid = ''%s''         ',
    [IDENTIFYID1, IDENTIFYID2, dtmMain.getCompID] ) );
    
  aDebugPath := IncludeTrailingPathDelimiter( Edt_MediaFilePath.Text ) + 'Debug';
  aDebugFileName := 'SqlTrace.txt';
  if not DirectoryExists( aDebugPath ) then ForceDirectories( aDebugPath );

end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA08.FormDestroy(Sender: TObject);
var
  aIndex: Integer;
begin
  for aIndex := 0 to FCompanyList.Count - 1 do
  begin
    if Assigned( FCompanyList[aIndex] ) then
       Dispose( PCompany( FCompanyList[aIndex] ) );
    FCompanyList[aIndex] := nil;
  end;
  FCompanyList.Clear;
  FCompanyList.Free;
  FDebugList.Free;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA08.btnRedoClick(Sender: TObject);
var
  aSQL : String;
begin
  MKE_Year.Clear;
  CBx_Month.ItemIndex := 0;
  RadioGroup1.ItemIndex := 2;
  MKE_Start.Clear;
  MKE_End.Clear;
  CheckBox1.Checked := False;
  aSQL := ' SELECT INVPARAM FROM INV003 WHERE COMPID = ' +
    QuotedStr( dtmMain.getCompID );
  Edt_MediaFilePath.Text := dtmMainJ.GetXMLAttribute( aSQL, 'MediaFilePath' );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA08.btnExitClick(Sender: TObject);
begin
   Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA08.btnOKClick(Sender: TObject);
Var
  //�D���ܼ�
  sL_Inv01BusinessID    : String    ;//���q�νs
  sL_Inv001TaxID        : String    ;//�|�y�s��
  MediaList             : TStringList;//�C��ӳ��Ȧs
  SL_RecNo              : String    ;//�Ǹ����
  //�������ܼ�
  tL_Sum                : TarraySum;
  StDate,EdDate         : Tdatetime;
  aCompIdList: String;
begin
  if MKE_Year.Text = '' then
  begin
    MessageDlg('�п�J�ӳ����~��!',mtError,[mbOK],0);
    exit;
  end;

  if (Trim(MKE_Start.Text) <> '') then
  begin
    if (Trim(MKE_End.Text) = '') then
    begin
      MessageDlg('�п�J�o���I��X!',mtError,[mbOK],0);
      exit;
    end;
    if length(trim(mke_end.Text)) < 10 then
    begin
      MessageDlg('�п�J���T���o���I��X!',mtError,[mbOK],0);
      exit;
    end;
    if length(trim(MKE_Start.Text)) < 10 then
    begin
      MessageDlg('�п�J���T���o���_�l���X!',mtError,[mbOK],0);
      exit;
    end;
    MKE_Start.Tag := 1;
  end
  else
  begin
    if (Trim(MKE_End.Text) <> '') then
    begin
      MessageDlg('�п�J�o���_�l���X!',mtError,[mbOK],0);
      exit;
    end;
    MKE_Start.Tag := 0;
  end;

  PCT1.ActivePageIndex := 1;
  StDate := now;
  try
    Bit_Previous.Enabled := False;
    Btn_Print.Enabled := False;
    btnExit.Enabled := False;
    ProgressBar2.Position := 0;
    SL_RecNo := '1';
    sL_MakeMonth := MKE_Year.Text +CBx_Month.Items[CBx_Month.ItemIndex];//�զX�ӳ��~��
    //���Xinv001�����q�νs�B�|�y�s��
    with TADOQuery.Create(self) do
    begin
      Connection := dtmMain.InvConnection;
      SQL.Clear;
      SQL.Add( 'SELECT BUSINESSID, TAXID FROM INV001 ' );
      SQL.Add( ' WHERE IDENTIFYID1 = :P1 ' );
      SQL.Add( '   AND IDENTIFYID2 = :P2 ');
      SQL.Add( '   AND COMPID = :P3 ' );

      Parameters.ParamByName('P1').Value := IDENTIFYID1;
      Parameters.ParamByName('P2').Value := IDENTIFYID2;
      Parameters.ParamByName('P3').Value := dtmMain.getCompID;

      Open;

      sL_Inv01BusinessID := FieldByName('BusinessID').AsString;
      sL_Inv001TaxID := FieldByName('TaxID').AsString;
      close;
      free;
    end;

    //�p��ӳ�������O
    case RadioGroup1.ItemIndex of
      0://���
      begin
        sL_CreateMonthStart := getYearMonthDay8(sL_MakeMonth,1);
        sL_CreateMonthEnd :=getYearMonthDay8(sL_MakeMonth,2);
      end;
      1://����
      begin
        sL_CreateMonthStart :=getYearMonthDay8(sL_MakeMonth,3);
        sL_CreateMonthEnd :=getYearMonthDay8(sL_MakeMonth,4);
      end;
      2://�G�Ӥ�
      begin
        sL_CreateMonthStart :=getYearMonthDay8(sL_MakeMonth,1);
        sL_CreateMonthEnd :=getYearMonthDay8(sL_MakeMonth,4);
      end;
    end;

    MediaList := TStringList.Create;

    {}
    FApplyByMainInv := chkApplyMainInv.Checked;
    {}

{=============================�p��ӳ��D�ɸ��=================================}
    //�p��o���D�ɸ��
    SL_RecNo := '1';
    lbl_show.Caption := '�o���D�ɸ�ƳB�z�i��';
    ProgressBar2.Position := 0;
    Refresh;
    Screen.Cursor := crSQLWait;
    try
      SetMasterData(SL_RecNo,sL_Inv01BusinessID,sL_Inv001TaxID,
                    'INV007', 1, MediaList );
    finally
      Screen.Cursor := crDefault;
    end;
    lbl_show.Caption := '�o���D�ɸ�ƳB�z����';
    frmInvA08.Refresh;

    { ��ڥӳ�,�����ͥ��ϥεo���O�� }
    if ( not chkPaper.Checked ) then
    begin
      //�p��ťո��
      lbl_show.Caption := '���ϥεo����ƳB�z�i��';
      ProgressBar2.Position := 0;
      Refresh;
      Screen.Cursor := crSQLWait;
      try
        SetMasterData(SL_RecNo,sL_Inv01BusinessID,sL_Inv001TaxID,
                      'INV099', 3, MediaList );
      finally
        Screen.Cursor := crDefault;
      end;
      lbl_show.Caption := '���ϥεo����ƳB�z����';
      frmInvA08.Refresh;
    end;  

//    SL_RecNo := '1';
    //�p��P�f�h�^ �� �����ҩ�����
    lbl_show.Caption := '�P�f�h�^ �� �����ҩ����ƳB�z�i��';
    ProgressBar2.Position := 0;
    Refresh;
    Screen.Cursor := crSQLWait;
    try
      SetMasterData(SL_RecNo,sL_Inv01BusinessID,sL_Inv001TaxID,
                    'INV014',2,MediaList);
    finally
      Screen.Cursor := crDefault;
    end;                
    lbl_show.Caption := '�P�f�h�^ �� �����ҩ����ƳB�z����';
    frmInvA08.Refresh;

//    SL_RecNo := '1';//�Ǹ��p�n���s�p��ɡA�O�o���}
    //�p��L��|���Ӹ��
    if CheckBox1.Checked then
    begin
      lbl_show.Caption := '�L��|���Ӹ�ƳB�z�i��';
      ProgressBar2.Position := 0;
      Refresh;
      Screen.Cursor := crSQLWait;
      try
        SetMasterData( SL_RecNo, sL_Inv01BusinessID, sL_Inv001TaxID,
                      dtmMain.getMisDbOwner+ '.SO033', 4, MediaList );
      finally
        Screen.Cursor := crDefault;
      end;                
      lbl_show.Caption := '�L��|���Ӹ�ƳB�z����';
      frmInvA08.Refresh;
    end;

    
    //�L��|�`ú
    if chkStampTax.Checked then
    begin
      lbl_show.Caption := '�L��|�`ú��ƳB�z�i��';
      ProgressBar2.Position := 0;
      Refresh;
      Screen.Cursor := crSQLWait;
      try
        sL_StampTaxt := FloatToStr( txtStampTax.Value );
        SetMasterData( SL_RecNo, sL_Inv01BusinessID, sL_Inv001TaxID,
                      EmptyStr, 5, MediaList );
      finally
        Screen.Cursor := crDefault;
      end;
      lbl_show.Caption := '�L��|�`ú��ƳB�z����';
      frmInvA08.Refresh;
    end;

    if ( chkDebugFile.Checked ) then
    begin
      FDebugList.SaveToFile( IncludeTrailingPathDelimiter( aDebugPath ) + aDebugFileName );
      FDebugList.Clear;
    end;


    case RadioGroup1.ItemIndex of
      0:
      begin
        fL_MediaFile :=Edt_MediaFilePath.Text +'\'+sL_Inv01BusinessID+'.txt';
        fL_MediaFile_B :=Edt_MediaFilePath.Text +'\'+sL_Inv01BusinessID+'_'+sL_MakeMonth+'_��.txt';
        fL_MediaReport :=Edt_MediaFilePath.Text +'\'+sL_Inv01BusinessID+'_Report.txt';
        fL_MediaReport_B :=Edt_MediaFilePath.Text +'\'+sL_Inv01BusinessID+'_Report_'+sL_MakeMonth+'_��.txt';
      end;
      1:
      begin
        fL_MediaFile :=Edt_MediaFilePath.Text +'\'+sL_Inv01BusinessID+'.txt';
        fL_MediaFile_B :=Edt_MediaFilePath.Text +'\'+sL_Inv01BusinessID+'_'+sL_MakeMonth+'_��.txt';
        fL_MediaReport :=Edt_MediaFilePath.Text +'\'+sL_Inv01BusinessID+'_Report.txt';
        fL_MediaReport_B :=Edt_MediaFilePath.Text +'\'+sL_Inv01BusinessID+'_Report_'+sL_MakeMonth+'_��.txt';
      end;
      2:
      begin
        fL_MediaFile :=Edt_MediaFilePath.Text +'\'+sL_Inv01BusinessID+'.txt';
        fL_MediaFile_B :=Edt_MediaFilePath.Text +'\'+sL_Inv01BusinessID+'_'+sL_MakeMonth+'_all.txt';
        fL_MediaReport :=Edt_MediaFilePath.Text +'\'+sL_Inv01BusinessID+'_Report.txt';
        fL_MediaReport_B :=Edt_MediaFilePath.Text +'\'+sL_Inv01BusinessID+'_Report_'+sL_MakeMonth+'_all.txt';
      end;
    end;

    MediaList.SaveToFile(fL_MediaFile);
    MediaList.SaveToFile(fL_MediaFile_B);
    MediaList.Clear;

    MediaList.Add('�Цܡi�D���j�U�C��m����ɮסG');
    MediaList.Add(' ');
    MediaList.Add('�C��ӳ���     ==�� '+fL_MediaFile);
    MediaList.Add('�C�������     ==�� '+fL_MediaReport);
    MediaList.Add('�C��ӳ��ƥ��� ==�� '+fL_MediaFile_B);
    MediaList.Add('�ӳ����ӳƥ��� ==�� '+fL_MediaREport_B);

{====================�H�U����������ɮפ������(���|)==========================}
    MediaList.Add(' ');
    MediaList.Add('�H�U����������ɮפ������(���|)');
    //A.  �p��D��~�H(�q�l)�o���B�`�p
    tL_Sum := TarraySum.Create;
    lbl_show.Caption := '�D��~�H(�q�l)�o���B�`�p�B�z';
    Refresh;
    tL_Sum.sL_A := SetMediaData(1,'InvAmount','INV007', 'N','1','','1','1');
    MediaList.Add('�D��~�H(�q�l)�o���B�`�p�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_A));

    //B.  �D��~�H(�q�l)�P���B�`�p
    lbl_show.Caption := '�D��~�H(�q�l)�P���B�`�p�B�z';
    Refresh;
    tL_Sum.sL_B := SetMediaData(1,'SaleAmount','INV007', 'N','1','','1','1');
    MediaList.Add('�D��~�H(�q�l)�P���B�`�p�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_B));

    //C.  �D��~�H(�q�l)�|�B�`�p
    lbl_show.Caption := '�D��~�H(�q�l)�|�B�`�p�B�z';
    Refresh;
    tL_Sum.sL_C := SetMediaData(1,'TaxAmount','INV007', 'N','1','','1','1');
    MediaList.Add('�D��~�H(�q�l)�|  �B�`�p�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_C));

    //D.  ��~�H(�q�l)�o���B�`�p
    lbl_show.Caption := '��~�H(�q�l)�o���B�`�p�B�z';
    Refresh;
    tL_Sum.sL_D := SetMediaData(1,'InvAmount','INV007', 'N','1','','2','1');
    MediaList.Add('��~�H(�q�l)�o���B�`�p�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_D));

    //E.	��~�H(�q�l)�P���B�`�p
    lbl_show.Caption := '��~�H(�q�l)�P���B�`�p�B�z';
    Refresh;
    tL_Sum.sL_E := SetMediaData(1,'SaleAmount','INV007', 'N','1','','2','1');
    MediaList.Add('��~�H(�q�l)�P���B�`�p�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_E));

    //F.	��~�H(�q�l)�|�B�`�p
    lbl_show.Caption := '��~�H(�q�l)�|�B�`�p�B�z';
    Refresh;
    tL_Sum.sL_F := SetMediaData(1,'TaxAmount','INV007', 'N','1','','2','1');
    MediaList.Add('��~�H(�q�l)�|  �B�`�p�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_F));

{====================�H�U��<�p����>�C��ӳ����p��============================}
    MediaList.Add(' ');
    MediaList.Add('�H�U���խp���ơִC��ӳ����p��');
    //G.	�D��~�H--�q�l(���|)�o���B�`�p
    lbl_show.Caption := '�D��~�H--�q�l(���|)�o���B�`�p�B�z';
    Refresh;
    tL_Sum.sL_G := SetMediaData(1,'InvAmount','INV007', 'N','1','','1','1');
    MediaList.Add('�D��~�H--�q�l(���|)�o���B�`�p�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_G));

    //H.	�D��~�H--�q�l(���|)�P���B�`�p{�o���B / 1.05}
    lbl_show.Caption := '�D��~�H--�q�l(���|)�P���B�`�p�B�z';
    Refresh;
    tL_Sum.sL_H := SimpleRoundTo( tL_Sum.sl_G / 1.05,0);
    MediaList.Add('�D��~�H--�q�l(���|)�P���B�`�p(�o���B/1.05)�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_H));

    //I.	�D��~�H--�q�l(���|)�|�@�B�`�p{�P���B �� 0.05}
    lbl_show.Caption := '�D��~�H--�q�l(���|)�|�@�B�`�p�B�z';
    Refresh;
    tL_Sum.sL_I :=  SimpleRoundTo( tL_Sum.sl_H * 0.05,0 );
    MediaList.Add('�D��~�H--�q�l(���|)�|�@�B�`�p(�P���B*0.05)�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_I));

    //J.	��~�H--�q�l(���|)�P���B�`�p
    lbl_show.Caption := '��~�H--�q�l(���|)�P���B�`�p�B�z';
    Refresh;
    tL_Sum.sL_J := SetMediaData(1,'SaleAmount','INV007', 'N','1','','2','1');
    MediaList.Add('��~�H--�q�l(���|)�P���B�`�p�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_J));

    //K.	��~�H--�T�p(���|)�P���B�`�p
    lbl_show.Caption := '��~�H--�T�p(���|)�P���B�`�p�B�z';
    Refresh;
    tL_Sum.sL_K := SetMediaData(1,'SaleAmount','INV007','N','3','','2','1');
    MediaList.Add('��~�H--�T�p(���|)�P���B�`�p�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_K));

    //L.	��~�H--�T�p�B�q�l(���|)�P���B�`�p
    lbl_show.Caption := '��~�H--�T�p�B�q�l(���|)�P���B�`�p�B�z';
    Refresh;
    tL_Sum.sL_L := tL_Sum.sL_J + tL_Sum.sL_K;
    MediaList.Add('��~�H--�T�p�B�q�l(���|)�P���B�`�p�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_L));

    //M.	��~�H--�q�l(���|)�|�B�`�p
    lbl_show.Caption := '��~�H--�q�l(���|)�|�B�`�p�B�z';
    Refresh;
    tL_Sum.sL_M := SetMediaData(1,'TaxAmount','INV007', 'N','1','','2','1');
    MediaList.Add('��~�H--�q�l(���|)�|�B�`�p�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_M));

    //N.	��~�H--�T�p(���|)�|�B�`�p
    lbl_show.Caption := '��~�H--�T�p(���|)�|�B�`�p�B�z';
    Refresh;
    tL_Sum.sL_N := SetMediaData(1,'TaxAmount','INV007', 'N','3','','2','1');
    MediaList.Add('��~�H--�T�p(���|)�|�B�`�p�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_N));

    //O.	��~�H--�T�p�B�q�l(���|)�|�B�`�p
    lbl_show.Caption := '��~�H--�T�p�B�q�l(���|)�|�B�`�p�B�z';
    Refresh;
    tL_Sum.sL_O := tL_Sum.sL_M + tL_Sum.sL_N;
    MediaList.Add('��~�H--�T�p�B�q�l(���|)�|�B�`�p�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_O));

    //P.	�G�p(���|)�o���B�`�p
    lbl_show.Caption := '�G�p(���|)�o���B�`�p�B�z';
    Refresh;
    tL_Sum.sL_P := SetMediaData(1,'InvAmount','INV007', 'N','2','','','1');
    MediaList.Add('�G�p(���|)�o���B�`�p�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_P));

    //Q.	�G�p(���|)�P���B�`�p{�o���B / 1.05}
    lbl_show.Caption := '�G�p(���|)�P���B�`�p�B�z';
    Refresh;
    tL_Sum.sL_Q := SimpleRoundTo( tL_Sum.sL_P / 1.05,0);
    MediaList.Add('�G�p(���|)�P���B�`�p(�o���B/1.05)�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_Q));

    //R.	�G�p(���|)�|�@�B�`�p{�P���B �� 0.05}
    lbl_show.Caption := '�G�p(���|)�|�@�B�`�p�B�z';
    Refresh;
    tL_Sum.sL_R := SimpleRoundTo( tL_Sum.sL_Q * 0.05 ,0);
    MediaList.Add('�G�p(���|)�|�@�B�`�p(�P���B*0.05)�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_R));

    //S.	��~/�D��~�H-����(���|)�P���B�`�p
    lbl_show.Caption := '��~/�D��~�H-����(���|)�P���B�`�p�B�z';
    Refresh;
    tL_Sum.sL_S := SetMediaData( 4, 'SaleAmount','INV014', '','','','','1');
    MediaList.Add('��~/�D��~�H-����(���|)�P���B�`�p�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_S));

    //T.	��~/�D��~�H-����(���|)�|�B�`�p
    lbl_show.Caption := '��~/�D��~�H-����(���|)�|�B�`�p�B�z';
    Refresh;
    tL_Sum.sL_T := SetMediaData(4,'TaxAmount','INV014', '','','','','1');
    MediaList.Add('��~/�D��~�H-����(���|)�|  �B�`�p�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_T));

{=========================�H�U���C��ӳ����B�έp===============================}
    MediaList.Add(' ');
    if ( chkPaper.Checked ) then
      MediaList.Add('�H�U���մC��ӳ��Ѹ�ơ֪��B�έp--�H���ڥӳ�')
    else
      MediaList.Add('�H�U���մC��ӳ��Ѹ�ơ֪��B�έp');

    //U.	�T�p/�q�l(���|-�P��)�`�p
    lbl_show.Caption := '�T�p/�q�l(���|-�P��)�`�p�B�z';
    Refresh;
    tL_Sum.sL_U := tL_Sum.sL_H + tL_Sum.sL_L;
    MediaList.Add('�T�p/�q�l(���|-�P��)�`�p�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_U));

    //V.	�T�p/�q�l(���|-�|�B)�`�p
    lbl_show.Caption := '�T�p/�q�l(���|-�|�B)�`�p�B�z';
    Refresh;
    tL_Sum.sL_V := tL_Sum.sL_I + tL_Sum.sL_O;
    MediaList.Add('�T�p/�q�l(���|-�|�B)�`�p�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_V));

    //W.	�G�p(���|-�P��)�`�p
    lbl_show.Caption := '�G�p(���|-�P��)�`�p�B�z';
    Refresh;
    tL_Sum.sL_W := tL_Sum.sL_Q;
    MediaList.Add('�G�p(���|-�P��)�`�p�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_W));

    //X.	�G�p(���|-�|�B)�`�p
    lbl_show.Caption := '�G�p(���|-�|�B)�`�p�B�z';
    Refresh;
    tL_Sum.sL_X := tL_Sum.sL_R;
    MediaList.Add('�G�p(���|-�|�B)�`�p�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_X));

    //Y.	����(���|-�P��)�`�p
    lbl_show.Caption := '����(���|-�P��)�`�p�B�z';
    Refresh;
    tL_Sum.sL_Y := tL_Sum.sL_S;
    MediaList.Add('����(���|-�P��)�`�p�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_Y));

    //Z.	����(���|-�|�B)�`�p
    lbl_show.Caption := '����(���|-�|�B)�`�p�B�z';
    Refresh;
    tL_Sum.sL_Z := tL_Sum.sL_T;
    MediaList.Add('����(���|-�|�B)�`�p�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_Z));

    //AA.	�X�p(���|-�P��)�`�p
    lbl_show.Caption := '�X�p(���|-�P��)�`�p�B�z';
    Refresh;
    tL_Sum.sL_AA := tL_Sum.sL_U + tL_Sum.sL_W - tL_Sum.sL_Y;
    MediaList.Add('�X�p(���|-�P��)�`�p�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_AA));

    //BB.	�X�p(���|-�|�B)�`�p
    lbl_show.Caption := '�X�p(���|-�|�B)�`�p�B�z';
    Refresh;
    tL_Sum.sL_BB := tL_Sum.sL_V + tL_Sum.sL_X - tL_Sum.sL_Z;
    MediaList.Add('�X�p(���|-�|�B)�`�p�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_BB));

    //CC.	�T�p/�q�l(�K�|)�`�p
    lbl_show.Caption := '�T�p/�q�l(�K�|)�`�p�B�z';
    Refresh;
    tL_Sum.sL_CC := SetMediaData(1,'Saleamount','INV007', 'N','1','3','','3');
    MediaList.Add('�T�p/�q�l(�K�|)�`�p�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_CC));

    //DD.	�G�p(�K�|)�`�p
    lbl_show.Caption := '�G�p(�K�|)�`�p�B�z';
    Refresh;
    tL_Sum.sL_DD := SetMediaData(1,'Saleamount','INV007', 'N','2','','','3');
    MediaList.Add('�G�p(�K�|)�`�p�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_DD));

    //EE.	����(�K�|)�`�p
    lbl_show.Caption := '����(�K�|)�`�p�B�z';
    Refresh;
    tL_Sum.sL_EE := SetMediaData( 5,'SaleAmount','INV014', '','','','','3');
    MediaList.Add('����(�K�|)�`�p�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_EE));

    //FF.	�X�p(�K�|)
    lbl_show.Caption := '�X�p(�K�|)�B�z';
    Refresh;
    tL_Sum.sL_FF := tL_Sum.sL_CC + tL_Sum.sL_DD - tL_Sum.sL_EE;
    MediaList.Add('�X�p(�K�|)�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_FF));


    //CC2. �T�p/�q�l(�s�|�v)�`�p
    lbl_show.Caption := '�T�p/�q�l(�s�|�v)�`�p�B�z';
    Refresh;
    tL_Sum.sL_CC2 := SetMediaData( 1,'Saleamount','INV007', 'N', '1', '3', '', '2' );
    MediaList.Add('�T�p/�q�l(�s�|�v)�`�p�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_CC2 ) );


    //DD2. �G�p(�s�|�v) �`�p
    lbl_show.Caption := '�G�p(�s�|�v)�`�p�B�z';
    Refresh;
    tL_Sum.sL_DD2 := SetMediaData( 1,'Saleamount','INV007', 'N', '2', '', '', '2' );
    MediaList.Add('�G�p(�s�|�v)�`�p�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_DD2 ) );


    //EE2. ����(�s�|�v) �`�p
    lbl_show.Caption := '����(�s�|�v)�`�p�B�z';
    Refresh;
    tL_Sum.sL_EE2 := SetMediaData( 5,'SaleAmount','INV014', '','','','','2' );
    MediaList.Add('����(�s�|�v)�`�p�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_EE2 ) );


    //FF2. �X�p(�s�|�v)
    lbl_show.Caption := '�X�p(�s�|�v)�B�z';
    Refresh;
    tL_Sum.sL_FF2 := ( tL_Sum.sL_CC2 + tL_Sum.sL_DD2 - tL_Sum.sL_EE2 );
    MediaList.Add('�X�p(�s�|�v)�G'+FormatFloat('###,###,###,##0', tL_Sum.sL_FF2 ) );



    //GG.	�P���B�`�p ( AA + FF + FF2 )
    lbl_show.Caption := '�P���B�`�p�B�z';
    Refresh;
    tL_Sum.sL_GG := tL_Sum.sL_AA + tL_Sum.sL_FF + tL_Sum.sL_FF2;
    MediaList.Add('�P���B�`�p�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_GG));


    if chkStampTax.Checked then
    begin
      //GG2. �L��|�`ú�`�B
      lbl_show.Caption := '�L��|�`ú�`�B';
      Refresh;
      tL_Sum.sL_GG2 := txtStampTax.Value;
      MediaList.Add('�L��|�`ú�`�B�G'+FormatFloat('###,###,###,##0',tL_Sum.sL_GG2 ) );
    end;

{=====================�H�U��<�C��ӳ��ѦҸ��>�i�Ʋέp��=======================}
    MediaList.Add(' ');
    MediaList.Add('�H�U���մC��ӳ��ѦҸ�ơֱi�Ʋέp��');
    //HH.	�q�l-���|�i�Ʋέp
    lbl_show.Caption := '�q�l-���|�i�Ʋέp�B�z';
    Refresh;
    tL_Sum.sL_HH := SetMediaData(2,'','INV007', 'N','1','','','1');

    //II.	�G�p-���|�i�Ʋέp
    lbl_show.Caption := '�G�p-���|�i�Ʋέp�B�z';
    Refresh;
    tL_Sum.sL_II := SetMediaData(2,'','INV007', 'N','2','','','1');

    //JJ.	�T�p-���|�i�Ʋέp
    lbl_show.Caption := '�T�p-���|�i�Ʋέp�B�z';
    Refresh;
    tL_Sum.sL_JJ := SetMediaData(2,'','INV007', 'N','3','','','1');
    MediaList.Add('(���|)==�ֹq�l:'+adjString2(10,tL_Sum.sL_HH)+
                           ' �G�p:'+adjString2(10,tL_Sum.sL_II)+
                           ' �T�p:'+adjString2(10,tL_Sum.sL_JJ));

    { -------------------------------------------------------- }

    //KK2. �q�l-�s�|�v�i�Ʋέp
    lbl_show.Caption := '�q�l-�s�|�v�i�Ʋέp�B�z';
    Refresh;
    tL_Sum.sL_KK2 := SetMediaData( 2 , '', 'INV007', 'N', '1', '',  '', '2' );

    //LL2. �G�p-�s�|�v�i�Ʋέp
    lbl_show.Caption := '�G�p-�s�|�v�i�Ʋέp�B�z';
    Refresh;
    tL_Sum.sL_LL2 := SetMediaData( 2 , '', 'INV007', 'N', '2', '',  '', '2' );

    //MM2. �T�p-�s�|�v�i�Ʋέp
    lbl_show.Caption := '�G�p-�s�|�v�i�Ʋέp�B�z';
    Refresh;
    tL_Sum.sL_MM2 := SetMediaData( 2 , '', 'INV007', 'N', '3', '',  '', '2' );
    MediaList.Add('(�s�|�v)==�ֹq�l:' + adjString2(10,tL_Sum.sL_KK2 ) +
                           '   �G�p:' + adjString2(10,tL_Sum.sL_LL2 ) +
                           '   �T�p:' + adjString2(10,tL_Sum.sL_MM2 ) );

    { -------------------------------------------------------- }


    //KK.	�q�l-�K�|�i�Ʋέp
    lbl_show.Caption := '�q�l-�K�|�i�Ʋέp�B�z';
    Refresh;
    tL_Sum.sL_KK := SetMediaData(2,'','INV007', 'N','1','','','3');

    //LL.	�G�p-�K�|�i�Ʋέp
    lbl_show.Caption := '�G�p-�K�|�i�Ʋέp�B�z';
    Refresh;
    tL_Sum.sL_LL := SetMediaData(2,'','INV007', 'N','2','','','3');

    //MM.	�T�p-�K�|�i�Ʋέp
    lbl_show.Caption := '�T�p-�K�|�i�Ʋέp�B�z';
    Refresh;
    tL_Sum.sL_MM := SetMediaData(2,'','INV007', 'N','3','','','3');
    MediaList.Add('(�K�|)==�ֹq�l:'+adjString2(10,tL_Sum.sL_KK)+
                           ' �G�p:'+adjString2(10,tL_Sum.sL_LL)+
                           ' �T�p:'+adjString2(10,tL_Sum.sL_MM));

    { -------------------------------------------------------- }
    // 2.��ڥӳ�,�w�@�o��Ƥ��ݲ��ܴͦC��ӳ���
    if ( not chkPaper.Checked ) then
    begin
      //NN.	�q�l-�@�o�i�Ʋέp
      lbl_show.Caption := '�q�l-�@�o�i�Ʋέp�B�z';
      Refresh;
      tL_Sum.sL_NN := SetMediaData(2,'','INV007', 'Y','1','','','');

      //OO.	�G�p-�@�o�i�Ʋέp
      lbl_show.Caption := '�G�p-�@�o�i�Ʋέp�B�z';
      Refresh;
      tL_Sum.sL_OO := SetMediaData(2,'','INV007', 'Y','2','','','');

      //PP.	�T�p-�@�o�i�Ʋέp
      lbl_show.Caption := '�T�p-�@�o�i�Ʋέp�B�z';
      Refresh;
      tL_Sum.sL_PP := SetMediaData(2,'','INV007', 'Y','3','','','');
      MediaList.Add('(�@�o)==�ֹq�l:'+adjString2(10,tL_Sum.sL_NN)+
                             ' �G�p:'+adjString2(10,tL_Sum.sL_OO)+
                             ' �T�p:'+adjString2(10,tL_Sum.sL_PP));
    end;
    
    { -------------------------------------------------------- }
    
    //QQ.	�q�l-�ťձi�Ʋέp
    lbl_show.Caption := '�q�l-�ťձi�Ʋέp�B�z';
    Refresh;
    tL_Sum.sL_QQ := SetMediaData(3,'','INV099','','1','','','');

    //RR.	�G�p-�ťձi�Ʋέp
    lbl_show.Caption := '�G�p-�ťձi�Ʋέp�B�z';
    Refresh;
    tL_Sum.sL_RR := SetMediaData(3,'','INV099','','2','','','');

    //SS.	�T�p-�ťձi�Ʋέp
    lbl_show.Caption := '�T�p-�ťձi�Ʋέp�B�z';
    Refresh;
    tL_Sum.sL_SS := SetMediaData(3,'','INV099','','3','','','');
    MediaList.Add('(�ť�)==�ֹq�l:'+adjString2(10,tL_Sum.sL_QQ)+
                           ' �G�p:'+adjString2(10,tL_Sum.sL_RR)+
                           ' �T�p:'+adjString2(10,tL_Sum.sL_SS));
                           

{===============�H�U�� <�ťե��ϥΤ��Τ@�o���r�y�θ��X�@����>==================}
    MediaList.Add(' ');
    case RadioGroup1.ItemIndex of
      0: mediaList.Add('�H�U��'+
         getYearMonthDay7(strtodate(sL_CreateMonthStart),1) +
         '�ժťե��ϥΤ��Τ@�o���r�y�θ��X�@�����');
      1: mediaList.Add('�H�U��'+
         getYearMonthDay7(strtodate(sL_CreateMonthEnd),1) +
         '�ժťե��ϥΤ��Τ@�o���r�y�θ��X�@�����');
      2: mediaList.Add('�H�U��'+
         getYearMonthDay7(strtodate(sL_CreateMonthStart),1) +' �� '+
         getYearMonthDay7(strtodate(sL_CreateMonthEnd),1) +
         '�ժťե��ϥΤ��Τ@�o���r�y�θ��X�@�����');
    end;
    MediaList.Add('�r�y�@�@ �o���_�W���X�@�@�@�o�������@�@ �i��');
    MediaList.Add('==== ==================== ========== ==========');


    aCompIdList := GetCompIdList;

    with AQry_Inv099 do
    begin
      sql.Text := 'Select Prefix, CurNum, EndNum, InvFormat, (EndNum-CurNum+1) NumCount '+
                  'From Inv099 Where yearmonth = ''' + sL_MakeMonth + '''' +
                  ' and compid in ' + aCompIdList + ' order by yearmonth, prefix, startnum ';
      open;
      if recordcount = 0 then
      begin
        tL_Sum.sL_TT := ' ';
        tL_Sum.sL_UU := ' ';
        tL_Sum.sL_VV := ' ';
        tL_Sum.sL_WW := 0;
      end
      else
      begin
        while not eof do
        begin
          tL_Sum.sL_TT := fieldByName('Prefix').asstring;
          // 2006/1/10 Jim �o�{ CurNum �Y���}�����ܷ|�ܦ� �w�ϥθ��X:500 ~ ����: 499 }
          // �Y�O�w�g���}���N���n�A show �X��, �ݰ�, �٨S�� }
          tL_Sum.sL_UU := fieldByName('CurNum').asstring+' �� ' +fieldByName('EndNum').asstring;
          if fieldByName('InvFormat').asstring = '1' then
            tL_Sum.sl_VV := '�q��o��'
          else if fieldByName('InvFormat').asstring = '2' then
            tL_Sum.sl_VV := '�G�p(��)'
          else if fieldByName('InvFormat').asstring = '3' then
            tL_Sum.sl_VV := '�T�p(��)';

          tL_Sum.sL_WW := fieldByName('NumCount').AsInteger;
          MediaList.Add(' '+tL_Sum.sL_TT+'�@'+
                            tL_Sum.sL_UU+'�@'+
                            tL_Sum.sl_VV+'�@'+
                            adjString2(10,tL_Sum.sL_WW));
          tl_sum.sL_ZZ := tl_sum.sL_ZZ + tl_sum.sL_WW;
          next;
        end;
      end;
    end;

    MediaList.Add('�@�@�@�@�@�@ �@�@�@�@�@�@�@�X�@�p�G�@'+adjString2(10,tL_Sum.sL_ZZ));
    MediaList.Add(' ');
    MediaList.SaveToFile(fL_MediaReport);
    MediaList.SaveToFile(fL_MediaReport_B);

    ListBox1.Items := MediaList;
    eddate := now;
    Lbl_show.Caption := '�C��ӳ��ɲ��ͧ���,�`�@��O      '+
                        timetostr(eddate - StDate)+' ��';
  except
    on E: Exception do
    begin
      Lbl_show.Caption := '�C��ӳ��ɲ��ͥ��ѡC'#13#10 + E.Message;
    end;  
  end;
  Bit_Previous.Enabled := True;
  Btn_Print.Enabled := True;
  btnExit.Enabled := True;
end;

function tfrminvA08.SetMediaData(dselect:integer;dField:String; dTable:String;
                    dParma1, dParma2, dParma3, dParma4, dParma5:String):Double ;
Var
  sL_sql : String;
  sl_parma : String;
  dL_sum : Double;
  aYearMonth: String;
  aCompIdList: String;
  aIndex: Integer;
begin
  aCompIdList := GetCompIdList;

  ProgressBar2.Position := 0;
  with AQry_Inv do
  begin
    //�զXsql�y�k�A�ӳ����
    if dselect = 1 then
    begin
      sL_sql := 'Select Sum('+dField+') Amount From ' +dTable+' ';
      sl_parma := 'where (INVDATE between '+
                  'to_date('+QuotedStr(sL_CreateMonthStart)+',''yyyy/mm/dd'') and '+
                  'to_date('+QuotedStr(sL_CreateMonthEnd)+',''yyyy/mm/dd'')) ' +
                  ' and compid in ' + aCompIdList;
    end else
    if dselect = 2 then
    begin
      sL_sql := 'Select Count(*) Amount From '+dTable+' ';
      sl_parma := 'where (INVDATE between '+
                  'to_date('+QuotedStr(sL_CreateMonthStart)+',''yyyy/mm/dd'') and '+
                  'to_date('+QuotedStr(sL_CreateMonthEnd)+',''yyyy/mm/dd'')) ' +
                  '  and compid in ' + aCompIdList;
    end else
    if dselect = 3 then
    begin
      sL_sql := 'Select (EndNum - CurNum + 1) Amount From '+dTable+' ';
      sl_parma := 'where Yearmonth between '+
                  QuotedStr(copy(sL_CreateMonthStart,1,4)+
                            copy(sL_CreateMonthStart,6,2))+
                  ' and '+
                  QuotedStr(copy(sL_CreateMonthEnd,1,4)+
                            copy(sL_CreateMonthEnd,6,2) ) +
                  ' and compid in ' + aCompIdList;          
    end else
    if dselect = 4 then
    begin
      aYearMonth := Copy( sL_CreateMonthStart, 1, 4 ) + Copy( sL_CreateMonthStart, 6, 2 );
      sL_sql := ' Select Sum('+dField+') Amount From '+dTable+' ';
      sl_parma := Format(
        ' where identifyid1 = ''%s''   ' +
        '   and identifyid2 = ''%s''   ' +
        '   and yearMonth = ''%s''     ' +
        '   and compid in %s           ',
        [IDENTIFYID1, IDENTIFYID2, aYearMonth, aCompIdList] );
    end else
    if dselect = 5 then
    begin
      aYearMonth := Copy( sL_CreateMonthStart, 1, 4 ) + Copy( sL_CreateMonthStart, 6, 2 );
      sL_sql := ' Select Sum('+dField+') Amount From '+dTable+' ';
      sl_parma := Format(
        ' where identifyid1 = ''%s''   ' +
        '   and identifyid2 = ''%s''   ' +
        '   and yearMonth = ''%s''     ' +
        '   and compid in %s           ',
        [IDENTIFYID1, IDENTIFYID2, aYearMonth, aCompIdList] );
    end;

    //�����ͥӳ���Ƥ��o�����X
    if MKE_Start.Tag = 1 then
      sl_parma := sl_parma + ' and not (INVID between '+QuotedStr(MKE_Start.Text)+
                             ' and '+QuotedStr(MKE_End.Text) + ') ';
    //�����@�o�����
    if dParma1 <> '' then
      sl_parma := sl_parma + 'and (IsObsolete = '+QuotedStr(dParma1) +') ';
    //�o���榡
    if dParma2 <> '' then
      if dParma3 <> '' then
        sl_parma := sl_parma + 'and (InvFormat IN ('+QuotedStr(dParma2) +','+
                    QuotedStr(dParma3)+')) '
      else
        sl_parma := sl_parma + 'and (InvFormat IN ('+QuotedStr(dParma2) +')) ';
    //�Ȥ�νs
    if dParma4 = '1' then
        sl_parma := sl_parma + 'and (BusinessId is null) '
    else
      if dParma4 = '2' then
        sl_parma := sl_parma + 'and (BusinessId is not null) ';
    //�o���|�O
    if  dParma5 <> '' then
        sl_parma := sl_parma + 'and (TaxType= '+QuotedStr(dParma5)+') ';

    sql.Clear;
    sql.Add(sl_sql);
    sql.Add(sl_parma);
    Open;
    //�N���A�C���̤j�ȳ]����r�ɪ����
    ProgressBar2.Properties.Max := recordCount ;
    if recordcount = 0 then
      Result := 0
    else
    begin
      if dselect = 3 then
      begin
        dl_sum := 0;
        while not eof do
        begin
          dL_sum := dl_sum + fieldByName('Amount').AsFloat;
          ProgressBar2.Position := ( ProgressBar2.Position + 1 ); 
          Next;
        end;
        Result := dL_sum;
      end
      else begin
        Result := fieldByName('Amount').AsFloat;
        ProgressBar2.Position := ( ProgressBar2.Position + 1 );
      end;
    end;
    close;
  end;
end;

Procedure TfrmInvA08.SetMasterData(Var P_RecNo:String;P_Inv01BusID,P_Inv001TaxID,
                  P_Table:String;P_Select:Integer;Var P_MediaList:TStringList);
Var
  sL_Inv007Format       : String    ;//�榡�N�X
  sL_TmpList            : String    ;//
  sL_Inv007InvDate      : String    ;//�}�ߵo������ন����~��
  sL_Inv007BusinessID   : String    ;//�o���W���Τ@�s��
  sL_Inv007Amount       : String    ;//�P���B
  sL_Inv007TaxType      : string    ;//�ҵ|�O
  sL_Inv007TaxId        : String    ;//�o�����X
  sL_TaxAmount          : String    ;//��~�|�B
  sL_detain             : String    ;//����N�X
  sL_special            : String    ;//�S����~�|�v
  sL_collect            : String    ;//�J�[���O
  sL_Other              : String    ;//�v�Ұs���O
  sL_Inv099Prefix       : String    ;//
  sL_Inv099EndNum       : String    ;//
  sL_Inv099StartNum     : String    ;//
  sL_Inv099CurNum       : String    ;//
  sL_startnum,sl_endnum : String;
  sL_InString           : String;
  iL_Count : Integer;
  sL_smbolstr           : String    ;
  i : integer;
  iL_LowNum :Integer;
  iL_HighNum : Integer;
  iL_CountNum : Integer;
  aCanWriteToFile: Boolean;
  aRecordIndex: Integer;
begin

  sL_smbolstr := 'yyyy/mm/dd';

  sL_InString := GetCompIdList;

  if P_Select = 3 then
  begin

    if ( chkDebugFile.Checked ) then
    begin
      FDebugList.Add( Format( '%s: P_Select = 3 ', [FormatDateTime( 'YYYY-MM-DD HH:MM:SS', Now )] ) );
    end;

    AQry_Inv.Sql.Clear;
    AQry_Inv.Sql.Add('select Count(*) Count from inv099 ');
    AQry_Inv.Sql.Add(' where IDENTIFYID1=%S ');
    AQry_Inv.Sql.Add(' and IDENTIFYID2=%S ');
    AQry_Inv.Sql.Add(' and COMPID in %S ');
    AQry_Inv.Sql.Add(' and YEARMONTH = %S ');

    AQry_Inv.Sql.Text := Format( AQry_Inv.Sql.Text,
       [QuotedStr(IDENTIFYID1), IDENTIFYID2, sL_InString,
        QuotedStr( sL_MakeMonth )] );
    AQry_Inv.Open;
    iL_Count := AQry_Inv.FieldByName('count').AsInteger;
    if iL_Count > 0 then
    begin
      AQry_Inv.close;
      AQry_Inv.Sql[0] := ' SELECT INVFORMAT,ENDNUM,PREFIX,CURNUM FROM INV099 ';
      AQry_Inv.Open;
      ProgressBar2.Properties.Max := AQry_Inv.RecordCount;
      while not AQry_Inv.Eof do
      begin
        { ��ӵo���� CURNUM > ENDNUM ��, �h����INSERT }
        if ( AQry_Inv.FieldByName('CurNum').AsInteger <=
             AQry_Inv.FieldByName('EndNum').AsInteger ) then
        begin
          //1���榡�N�X
          if AQry_Inv.FieldByName('InvFormat').AsString = '2' then
            sL_Inv007Format := '32'
          else
            sL_Inv007Format := '31';
          //2�|�y��� ==> P_Inv001TaxID
          //3���y���Ǹ�
          P_RecNo := RightStr( IntToStr( 10000000 + StrToInt( P_RecNo ) ), 7 );
          //4���o������~��
          sL_Inv007InvDate := IntToStr(StrToInt( Copy( sL_MakeMonth, 1, 4 ) ) -
            1911 ) +Copy( sL_MakeMonth, 5, 2 );
          //5�R���H�Τ@�s���ܧ󬰵o������
          sL_Inv007BusinessID := AQry_Inv.FieldByName('EndNum').AsString;
          //6�P��H�Τ@�s�� ==> P_Inv01BusID
          //7�B8�o���r�y
          sL_Inv007TaxId := ( AQry_Inv.FieldByName( 'Prefix' ).AsString +
            AQry_Inv.FieldByName('CurNum').AsString );
          //9�P���B��0
          sL_Inv007Amount := '000000000000';
          //10�ҵ|�O
          sL_Inv007TaxType := 'D';
          //11��~�|�B��0
          sL_TaxAmount := '0000000000';
          //12����N�X
          sL_detain := ' ';
          //13�S����~�|�v
          sL_special:= '      ';
          //14�J�[���O��J A
          sL_collect:= 'A';
          //15�v�Ұs���O
          sL_Other  := ' ';

          sL_TmpList := sL_Inv007Format+
                        P_Inv001TaxID+
                        P_RecNo+
                        sL_Inv007InvDate+
                        sL_Inv007BusinessID+
                        P_Inv01BusID+
                        sL_Inv007TaxId+
                        sL_Inv007Amount+
                        sL_Inv007TaxType+
                        sL_TaxAmount+
                        sL_detain+
                        sL_special+
                        sL_collect+
                        sL_Other;
          P_MediaList.Add(sL_TmpList);
        end;
        ProgressBar2.Position := ( ProgressBar2.Position + 1 );
        frmInvA08.Refresh;
        P_RecNo := IntToStr(StrToInt( P_RecNo ) + 1 );
        AQry_Inv.Next;
      end;
    end;
    AQry_Inv.Close;
  end
  else
  if (P_Select = 4) and (dtmMain.GetLinkToMIS) then//�L��|
  begin

    if ( chkDebugFile.Checked ) then
    begin
      FDebugList.Add( Format( '%s: P_Select = 4 ', [FormatDateTime( 'YYYY-MM-DD HH:MM:SS', Now )] ) );
    end;

    with dtmSO.adoA08 do
    begin
      Close;
      SQL.Clear;
      SQL.Add( Format( ' SELECT COUNT(*) COUNT FROM %s.SO034 T, %s.CD019 C ', [dtmMain.getMisDbOwner, dtmMain.getMisDbOwner] ) );
      SQL.Add('  WHERE T.CITEMCODE = C.CODENO ');
      SQL.Add('   AND  C.TAXCODE = %s ' );
      SQL.Add('   AND  COMPCODE= %s ' );
      SQL.add('   AND ( REALDATE BETWEEN %s AND %s ) ');

      SQL.Text := Format(sql.Text,[QuotedStr('4'),
                                   QuotedStr(dtmMain.getMisDbCompCode),
                                   'to_date('+QuotedStr(sL_CreateMonthStart)+','+QuotedStr(sL_smbolstr)+')',
                                   'to_date('+QuotedStr(sL_CreateMonthEnd)+','+QuotedStr(sL_smbolstr)+')']);

      Open;
      iL_Count := FieldByName( 'COUNT' ).AsInteger;
      if iL_Count > 0 then
      begin
        Close;
        SQL.Strings[0] := Format( 'SELECT T.REALDATE, T.BILLNO, T.REALAMT FROM %s.SO034 T, %s.CD019 C ', [dtmMain.getMisDbOwner, dtmMain.getMisDbOwner] );
        SQL.Add( ' ORDER BY T.BILLNO ' );

        Open;
        ProgressBar2.Properties.Max := iL_Count ;//�N���A�C���̤j�ȳ]����r�ɪ����
        while not Eof do
        begin
          //1���榡�N�X
          sL_Inv007Format := '36';
          //2�|�y��� ==> P_Inv001TaxID
          //3���y���Ǹ�
          P_RecNo := rightstr(inttostr(10000000+strtoint(P_RecNo)),7);
          //4���o������~��
          sL_Inv007InvDate := getYearMonthDay7(fieldByName('realdate').AsDateTime,1);
          //5�R���H�Τ@�s���ܧ󬰵o������
          sL_Inv007BusinessID := '        ';
          //6�P��H�Τ@�s�� ==> P_Inv01BusID
          //7�B8�o���r�y
          sL_Inv007TaxId   := '0'+copy(FieldByName('billno').AsString,5,2)+
                              copy(FieldByName('billno').AsString,9,7);
          //9�P���B��0
          sL_Inv007Amount := rightstr(inttostr(1000000000000+fieldByName('realamt').asInteger),12);
          //10�ҵ|�O
          sL_Inv007TaxType := '3';
          //11��~�|�B��0
          sL_TaxAmount := '0000000000';
          //12����N�X
          sL_detain := ' ';
          //13�S����~�|�v
          sL_special:= '      ';
          //14�J�[���O��J A
          sL_collect:= ' ';
          //15�v�Ұs���O  �Y���s�|�v, �h��J�q�����O
          sL_Other  := ' ';
          if ( sL_Inv007TaxType = '3' ) then
            sL_Other := Nvl( FImportNote, #32 );

          sL_TmpList := sL_Inv007Format+
                        P_Inv001TaxID+
                        P_RecNo+
                        sL_Inv007InvDate+
                        sL_Inv007BusinessID+
                        P_Inv01BusID+
                        sL_Inv007TaxId+
                        sL_Inv007Amount+
                        sL_Inv007TaxType+
                        sL_TaxAmount+
                        sL_detain+
                        sL_special+
                        sL_collect+
                        sL_Other;
          P_MediaList.Add(sL_TmpList);
          ProgressBar2.Position := ( ProgressBar2.Position + 1 );
          frmInvA08.Refresh;
          P_RecNo := inttostr(strtoint(P_RecNo) +1);
          next;
        end;
      end;
      Close;
    end;
  end
  else if P_Select = 5 then // �L��|�`ú
  begin

    if ( chkDebugFile.Checked ) then
    begin
      FDebugList.Add( Format( '%s: P_Select = 5 ', [FormatDateTime( 'YYYY-MM-DD HH:MM:SS', Now )] ) );
    end;

          //1���榡�N�X
          sL_Inv007Format := '36';
          //2�|�y��� ==> P_Inv001TaxID
          //3���y���Ǹ�
          P_RecNo := rightstr(inttostr(10000000+strtoint(P_RecNo)),7);
          //4���o������~��
          sL_Inv007InvDate := getYearMonthDay7(StrToDate( FormatMaskText( '####/##/##;0;_', sL_MakeMonth + '01' )),1);
          //5�R���H�Τ@�s���ܧ󬰵o������
          sL_Inv007BusinessID := Lpad( EmptyStr, 8, #32 );
          //6�P��H�Τ@�s�� ==> P_Inv01BusID
          //7�B8�o���r�y
          sL_Inv007TaxId   := Lpad( EmptyStr, 10, #32 );
          //9�P���B��0
          sL_Inv007Amount := Lpad( sL_StampTaxt, 12, '0' );
          //10�ҵ|�O
          sL_Inv007TaxType := '3';
          //11��~�|�B��0
          sL_TaxAmount := Lpad( EmptyStr, 10, '0' );
          //12����N�X
          sL_detain := ' ';
          //13�S����~�|�v
          sL_special:= '      ';
          //14�J�[���O��J A
          sL_collect:= ' ';
          //15�v�Ұs���O  �Y���s�|�v, �h��J�q�����O
          sL_Other  := ' ';
          if ( sL_Inv007TaxType = '3' ) then
            sL_Other := Nvl( FImportNote, #32 );

          sL_TmpList := sL_Inv007Format+
                        P_Inv001TaxID+
                        P_RecNo+
                        sL_Inv007InvDate+
                        sL_Inv007BusinessID+
                        P_Inv01BusID+
                        sL_Inv007TaxId+
                        sL_Inv007Amount+
                        sL_Inv007TaxType+
                        sL_TaxAmount+
                        sL_detain+
                        sL_special+
                        sL_collect+
                        sL_Other;
          P_MediaList.Add(sL_TmpList);
          ProgressBar2.Position := ( ProgressBar2.Position + 1 );
          frmInvA08.Refresh;
          P_RecNo := inttostr(strtoint(P_RecNo) +1);
  end
  else if P_Select = 1 then//�@��o��
  begin

    if ( chkDebugFile.Checked ) then
    begin
      FDebugList.Add( Format( '%s: P_Select = 1 ', [FormatDateTime( 'YYYY-MM-DD HH:MM:SS', Now )] ) );
    end;

    FDebugList.Clear;

    AQry_Inv099.SQL.Clear;
    AQry_Inv099.SQL.Add( ' SELECT Prefix, EndNum, CurNum, StartNum FROM INV099 ' );
    AQry_Inv099.SQL.Add( ' WHERE IDENTIFYID1=%S ' );
    AQry_Inv099.SQL.Add( '   AND IDENTIFYID2=%S ' );
    AQry_Inv099.SQL.Add( '   AND YEARMONTH = %S ' );
    AQry_Inv099.SQL.Add( '   AND COMPID in %S ' );
    AQry_Inv099.SQL.Text := Format( AQry_Inv099.SQL.text, [
      QuotedStr( IDENTIFYID1 ), IDENTIFYID2, QuotedStr( sL_MakeMonth ),
      sL_InString] );
    AQry_Inv099.SQL.Add( ' Order By Prefix, StartNum ' );
    AQry_Inv099.Open;
    AQry_Inv099.First;

    if ( chkDebugFile.Checked ) then
    begin
      FDebugList.Add( Format( '%s: �ҩl���� .....', [FormatDateTime( 'YYYY-MM-DD HH:MM:SS', Now )] ) );
      FDebugList.Add( Format( '%s: �C��ӳ��o������SQL', [FormatDateTime( 'YYYY-MM-DD HH:MM:SS', Now )] ) );
      FDebugList.Add( '--------------------------------------------------------' );
      FDebugList.Add( '   ' + AQry_Inv099.SQL.Text );
      FDebugList.Add( '--------------------------------------------------------' );
    end;

    while not AQry_Inv099.eof do
    begin
      sL_Inv099Prefix := AQry_Inv099.FieldByName('Prefix').AsString;
      sL_Inv099EndNum := AQry_Inv099.FieldByName('EndNum').AsString;
      sL_Inv099CurNum := AQry_Inv099.FieldByName('CurNum').AsString;
      sL_Inv099startNum := AQry_Inv099.FieldByName('StartNum').AsString;
      sl_startnum := sL_Inv099Prefix+adjString(8,sL_Inv099startNum,True,True);
      sl_endnum := sL_Inv099Prefix+ adjString(8,sL_Inv099EndNum,True,True);

      if ( chkDebugFile.Checked ) then
      begin
        FDebugList.Add( Format( '%s: �o������T, �r��:%s, �ҩl���X:%s, �������X:%s, �U�@�ϥθ��X:%s,  ���ձҩl���X:%s, ���յ������X:%s ',
          [FormatDateTime( 'YYYY-MM-DD HH:MM:SS', Now ), sL_Inv099Prefix, sL_Inv099startNum, sL_Inv099EndNum, sL_Inv099CurNum, sl_startnum, sl_endnum] ) );
      end;

      // ����X�`����, �����
      AQry_Inv.SQL.Text := Format(
        ' SELECT COUNT(1) COUNT                        ' +
        '   FROM %s                                    ' +
        ' WHERE IDENTIFYID1 = ''%s''                   ' +
        '   AND IDENTIFYID2 = %s                       ' +
        '   AND INVDATE BETWEEN                        ' +
        '         TO_DATE( ''%s'', ''YYYY/MM/DD'' )    ' +
        '       AND                                    ' +
        '         TO_DATE( ''%s'', ''YYYY/MM/DD'' )    ' +
        '   AND COMPID in %s                           ' +
        '   AND INVID BETWEEN ''%s'' AND ''%s''        ',
        [P_Table, IDENTIFYID1, IDENTIFYID2,
         sL_CreateMonthStart, sL_CreateMonthEnd, sL_InString,
         sl_startnum, sl_endnum] );

      if MKE_Start.tag = 1 then
      begin
        AQry_Inv.SQL.Text := AQry_Inv.SQL.Text +
          ' AND NOT ( INVID BETWEEN ' + QuotedStr( MKE_Start.Text )+
                ' AND ' + QuotedStr( MKE_End.Text ) + ' ) ';
      end;

      AQry_Inv.Open;
      iL_Count := AQry_Inv.FieldByName( 'Count' ).AsInteger;
      AQry_Inv.Close;

      if ( chkDebugFile.Checked ) then
      begin
        FDebugList.Add( Format( '%s: �̵o��������X�X�o����Ƶ��ƪ�SQL:', [FormatDateTime( 'YYYY-MM-DD HH:MM:SS', Now )] ) );
        FDebugList.Add( AQry_Inv.SQL.Text );
        FDebugList.Add( '--------------------------------------------------------' );
      end;

      // ��������Ʈ�, �C���̤j�������
      iL_CountNum := 5000;

      if iL_Count > 0 then
      begin
        ProgressBar2.Position := 0;
        ProgressBar2.Properties.Max := il_count ;//�N���A�C���̤j�ȳ]����r�ɪ����
        iL_LowNum := 1;
        iL_HighNum := iL_CountNum;
        if il_count > iL_CountNum then
        begin
          sl_startnum := sL_Inv099Prefix + adjString( 8, sL_Inv099startNum,
            True, True );
          sl_endnum := sL_Inv099Prefix + adjString( 8, IntToStr(
            StrToInt( Rightstr( sL_Inv099startNum, 8 ) ) + iL_CountNum - 1 ),
            True, True );
          {}
          if ( chkDebugFile.Checked ) then
          begin
            FDebugList.Add( '       ********************************************************' );
            FDebugList.Add( '       ' + Format( '%s: ���������, ���s�p�⪺�o�����X�_��: %s - %s',
              [FormatDateTime( 'YYYY-MM-DD HH:MM:SS', Now ), sl_startnum, sl_endnum] ) );
            FDebugList.Add( '       ********************************************************' );
          end;
          {}
        end;

        aRecordIndex := 1;
        while ( aRecordIndex <= il_count ) do
        begin

          if ( chkDebugFile.Checked ) then
          begin
            FDebugList.Add( Format( 'aRecordIndex:%d, iL_LowNum:%d, AQry_Inv.Active:%s',
              [aRecordIndex, iL_LowNum, BoolToStr( AQry_Inv.Active )] ) );
            FDebugList.Add( '********************************************************' );
          end;

              { ��B�z�����Ƥw��F�]�w������, �άO�Ӧ� SQL �õL���o���,
                �h�A���B�z SQL, ��Open DataSet }
              if ( aRecordIndex = iL_LowNum ) or ( not AQry_Inv.Active ) then
              begin
                 AQry_Inv.SQL.Text := Format(
                   ' SELECT INVFORMAT,             ' +
                   '        INVDATE,               ' +
                   '        ISOBSOLETE,            ' +
                   '        TAXTYPE,               ' +
                   '        BUSINESSID,            ' +
                   '        INVAMOUNT,             ' +
                   '        SALEAMOUNT,            ' +
                   '        TAXAMOUNT,             ' +
                   '        INVID,                 ' +
                   '        MAININVID,             ' +
                   '        MAINSALEAMOUNT,        ' +
                   '        MAINTAXAMOUNT,         ' +
                   '        MAININVAMOUNT          ' +
                   '  FROM %s                      ' +
                   ' WHERE IDENTIFYID1 = ''%s''    ' +
                   '   AND IDENTIFYID2 = %s        ' +
                   '   AND INVDATE BETWEEN         ' +
                   '           TO_DATE( ''%s'', ''YYYY/MM/DD'' )  ' +
                   '       AND                     ' +
                   '           TO_DATE( ''%s'', ''YYYY/MM/DD'' )  ' +
                   '   AND COMPID in %s            ' +
                   '   AND INVID BETWEEN ''%s'' AND ''%s''        ',
                   [P_Table, IDENTIFYID1, IDENTIFYID2, sL_CreateMonthStart,
                     sL_CreateMonthEnd, sL_InString, sl_startnum,
                     sl_endnum] );

                 if MKE_Start.tag = 1 then
                 begin
                   AQry_Inv.SQL.Text := AQry_Inv.SQL.Text +
                   Format( ' AND NOT ( INVID BETWEEN ''%s'' AND ''%s''  ) ',
                           [MKE_Start.Text, MKE_End.Text] );
                 end;
                AQry_Inv.Close;
                AQry_Inv.SQL.Text := ( AQry_Inv.SQL.Text + ' ORDER BY INVID ' );
                AQry_Inv.Open;

                if ( chkDebugFile.Checked ) then
                begin
                  FDebugList.Add( Format( '%s: �̺�X�����ƨ��X�o����ƪ�SQL:', [FormatDateTime( 'YYYY-MM-DD HH:MM:SS', Now )] ) );
                  FDebugList.Add( AQry_Inv.SQL.Text );
                  FDebugList.Add( '--------------------------------------------------------' );
                end;

                iL_LowNum := iL_HighNum + 1;
                iL_HighNum := iL_HighNum + iL_CountNum;
                sl_startnum := sL_Inv099Prefix + adjString( 8,
                  IntToStr(StrToInt( RightStr( sl_endnum,8 ) ) + 1 ),True, True );
                sl_endnum := sL_Inv099Prefix+ adjString( 8, IntToStr( StrToInt(
                  RightStr( sl_endnum, 8 ) ) + iL_CountNum ),True, True );

                { �S������, ���� DataSet }
                if ( AQry_Inv.IsEmpty ) then
                begin
                  AQry_Inv.Close;
                  Continue;
                end;

              end;

              aCanWriteToFile := True;

              { �H�D�o�����X�ӳ� }
              if FApplyByMainInv then
              begin
                { �O�_�O�D�o�����X }
                if ( AQry_Inv.FieldByName( 'INVID' ).AsString <>
                     Nvl( AQry_Inv.FieldByName( 'MAININVID' ).AsString,
                          AQry_Inv.FieldByName( 'INVID' ).AsString ) ) then
                aCanWriteToFile := False;
              end;



              if ( aCanWriteToFile ) then
              begin


                //1���榡�N�X
                  if ( chkPaper.Checked ) then
                  begin
                    sL_Inv007Format := '36';
                  end else
                  begin
                    if AQry_Inv.FieldByName('InvFormat').AsString = '2' then
                      sL_Inv007Format := '32'
                    else
                      sL_Inv007Format := '31';
                  end;
                  //4���o������~��
                  sL_Inv007InvDate := getYearMonthDay7( AQry_Inv.fieldByName('INVDATE').AsDateTime,1 );
                  //�P�_�o���|�O
                  if AQry_Inv.FieldByName('ISOBSOLETE').AsString = 'Y' then
                  begin
                    //�o�����@�o�ɡA�ܧ�U�C���
                    //5�R���H�Τ@�s���ܧ󬰵o������
                    sL_Inv007BusinessID := '00000000';
                    //9�P���B��0
                    sL_Inv007Amount := '000000000000';
                    //10�ҵ|�O�אּD
                    sL_Inv007TaxType := 'D';
                    //11��~�|�B��0
                    sL_TaxAmount := '0000000000';
                    //14�J�[���O��J A
                    sL_collect:= ' ';
                  end else
                  begin
                    //10�ҵ|�O
                    sL_Inv007TaxType := inttostr(AQry_Inv.fieldByNAme('TaxType').AsInteger);
                    //14�J�[���O��J�ť�
                    sL_collect:= ' ';
                    //�P�_�o���νs�O�_���ťաA�O�h��8��0�F�t�i�P�_�P���B�B��~�|�B
                    if AQry_Inv.FieldByName('businessID').AsString = EmptyStr then
                    begin
                      //5�R���H�Τ@�s��
                      sL_Inv007BusinessID := '00000000';
                      //9�P���B
                      if ( FApplyByMainInv ) then
                        sL_Inv007Amount := RightStr( IntToStr( 1000000000000 + AQry_Inv.FieldByName('MainInvAmount').AsInteger ), 12 )
                      else
                        sL_Inv007Amount := RightStr( IntToStr( 1000000000000 + AQry_Inv.FieldByName('InvAmount').AsInteger ), 12 );
                      //11��~�|�B
                      sL_TaxAmount := '0000000000';
                    end else
                    begin
                      //5�R���H�Τ@�s��
                      sL_Inv007BusinessID := AQry_Inv.FieldByName('businessID').AsString;
                      if ( FApplyByMainInv ) then
                      begin
                        //9�P���B
                        sL_Inv007Amount := RightStr( IntToStr( 1000000000000 + AQry_Inv.FieldByName('MainSaleAmount').AsInteger), 12 );
                        //11��~�|�B
                        sL_TaxAmount := RightStr( IntToStr( 10000000000 + AQry_Inv.FieldByName('MainTaxAmount').AsInteger ), 10 );
                      end else
                      begin
                        //9�P���B
                        sL_Inv007Amount := RightStr( IntToStr( 1000000000000 + AQry_Inv.FieldByName('SaleAmount').AsInteger), 12 );
                        //11��~�|�B
                        sL_TaxAmount := RightStr( IntToStr( 10000000000 + AQry_Inv.FieldByName('TaxAmount').AsInteger ), 10 );
                      end;
                    end;
                  end;

                //2�|�y��� ==> P_Inv001TaxID
                //3���y���Ǹ�
                P_RecNo := Rightstr( IntToStr( 10000000 + StrToInt( P_RecNo ) ), 7 );
                //6�P��H�Τ@�s�� ==> P_Inv01BusID
                //7�B8�o���r�y
                sL_Inv007TaxId := AQry_Inv.FieldByName('InvID').AsString;

                //12����N�X
                sL_detain := ' ';
                //13�S����~�|�v
                sL_special:= '      ';
                //15�v�Ұs���O  �Y���s�|�v, �h��J�q���覡
                sL_Other  := ' ';
                if ( sL_Inv007TaxType = '3' ) then
                  sL_Other := Nvl( FImportNote, #32 );

                sL_TmpList := sL_Inv007Format+
                              P_Inv001TaxID+
                              P_RecNo+
                              sL_Inv007InvDate+
                              sL_Inv007BusinessID+
                              P_Inv01BusID+
                              sL_Inv007TaxId+
                              sL_Inv007Amount+
                              sL_Inv007TaxType+
                              sL_TaxAmount+
                              sL_detain+
                              sL_special+
                              sL_collect+
                              sL_Other;

                // �H���ڥӳ�����, ���o��Ƥ�������
                if ( ( AQry_Inv.FieldByName('ISOBSOLETE').AsString = 'Y' ) and
                     ( chkPaper.Checked ) ) then
                begin
                   // do nothing
                end else
                begin
                  P_MediaList.Add(sL_TmpList);
                  P_RecNo := inttostr(strtoint(P_RecNo) +1);
                  if ( chkDebugFile.Checked ) then
                  begin
                    FDebugList.Add( Format( '%s: P_RecNo:%s', [FormatDateTime( 'YYYY-MM-DD HH:MM:SS', Now ), P_RecNo] ) );
                  end;
                end;

              end;
              ProgressBar2.Position := ( ProgressBar2.Position + 1 );
              frmInvA08.Refresh;
              AQry_Inv.Next;
              Inc( aRecordIndex );

              if ( AQry_Inv.Eof ) then
              begin
                AQry_Inv.Close;
                if ( chkDebugFile.Checked ) then
                begin
                  FDebugList.Add( Format( '%s: AQry_Inv.Closed()', [FormatDateTime( 'YYYY-MM-DD HH:MM:SS', Now )] ) );
                end;
              end;

        end;
      end;

      AQry_Inv099.next;


    end;

  end
  else if P_Select = 2 then
  begin

    if ( chkDebugFile.Checked ) then
    begin
      FDebugList.Add( Format( '%s: P_Select = 2 ', [FormatDateTime( 'YYYY-MM-DD HH:MM:SS', Now )] ) );
    end;

    AQry_Inv.Close;
    AQry_Inv.SQL.Text := Format(
      ' SELECT A.* FROM %s A                 ' +
      '  WHERE A.IDENTIFYID1 = ''%s''        ' +
      '    AND A.IDENTIFYID2 = %s            ' +
      '    AND A.YEARMONTH = ''%s''          ' +
      '    AND A.COMPID in %s                ',
      [P_Table, IDENTIFYID1, IDENTIFYID2,
       Copy( sL_CreateMonthStart, 1, 4 ) + Copy( sL_CreateMonthStart, 6, 2 ),
       sL_InString] );

    if MKE_Start.tag = 1 then
    begin
      AQry_Inv.SQL.Text := AQry_Inv.SQL.Text +
      Format( ' AND NOT ( A.INVID BETWEEN ''%s'' AND ''%s''  ) ',
              [MKE_Start.Text, MKE_End.Text] );
    end;
    AQry_Inv.Open;
    ProgressBar2.Properties.Max := AQry_Inv.RecordCount ;//�N���A�C���̤j�ȳ]����r�ɪ����
    AQry_Inv.First;
    while not AQry_Inv.Eof do
    begin
      //1���榡�N�X
      if ( chkPaper.Checked ) then
      begin
        sL_Inv007Format := '34';
      end else
      begin
        if AQry_Inv.FieldByName('INVFORMAT').AsInteger = 2 then
          sL_Inv007Format := '34'
        else
          sL_Inv007Format := '33';
      end;    
      //2�|�y��� ==> P_Inv001TaxID
      //3���y���Ǹ�
      P_RecNo := rightstr(inttostr(10000000+strtoint(P_RecNo)),7);
      //4���o������~��
      sL_Inv007InvDate := getYearMonthDay7( AQry_Inv.FieldByName('PAPERDATE').AsDateTime,1);
      //�P�_�o���νs�O�_���ťաA�O�h��8��0�F�t�i�P�_�P���B�B��~�|�B
      if AQry_Inv.FieldByName('BUSINESSID').AsString = '' then
        //5�R���H�Τ@�s��
        sL_Inv007BusinessID := '00000000'
      else
        //5�R���H�Τ@�s��
        sL_Inv007BusinessID := AQry_Inv.FieldByName('BUSINESSID').AsString;
      //6�P��H�Τ@�s�� ==> P_Inv01BusID
      //7�B8�o���r�y
      sL_Inv007TaxId  := AQry_Inv.FieldByName('INVID').AsString;
      //9�P���B
      sL_Inv007Amount := rightstr(inttostr(1000000000000+ AQry_Inv.FieldByName('SALEAMOUNT').asInteger),12);
      //10�ҵ|�O
      sL_Inv007TaxType := inttostr(AQry_Inv.FieldByNAme('TAXTYPE').AsInteger);
      //11��~�|�B
      sL_TaxAmount := rightstr(inttostr(10000000000+AQry_Inv.FieldByName('TAXAMOUNT').asInteger),10);
      //12����N�X
      sL_detain := ' ';
      //13�S����~�|�v
      sL_special:= '      ';
      //14�J�[���O
      sL_collect:= ' ';
      //15�v�Ұs���O �Y���s�|�v,�h��J�q�����O
      sL_Other  := ' ';
      if ( sL_Inv007TaxType = '3' ) then
        sL_Other := Nvl( FImportNote, #32 );

      sL_TmpList := sL_Inv007Format+
                    P_Inv001TaxID+
                    P_RecNo+
                    sL_Inv007InvDate+
                    sL_Inv007BusinessID+
                    P_Inv01BusID+
                    sL_Inv007TaxId+
                    sL_Inv007Amount+
                    sL_Inv007TaxType+
                    sL_TaxAmount+
                    sL_detain+
                    sL_special+
                    sL_collect+
                    sL_Other;
      P_MediaList.Add(sL_TmpList);
      ProgressBar2.Position := ( ProgressBar2.Position + 1 );
      frmInvA08.Refresh;
      P_RecNo := inttostr(strtoint(P_RecNo) +1);
      AQry_Inv.Next;
    end;
    AQry_Inv.Close;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA08.Bit_PreviousClick(Sender: TObject);
begin
  ListBox1.Clear;
  PCT1.ActivePageIndex := 0;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA08.MKE_StartExit(Sender: TObject);
begin
  if trim(MKE_End.Text) = '' then
    MKE_End.Text := MKE_Start.Text;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA08.Btn_PrintClick(Sender: TObject);
var
  aIndex: Integer;
begin
  if PrinterSetupDialog1.Execute then
  begin
    Printer.BeginDoc;
    try
      Printer.Canvas.Font.Size := 10;
      Printer.Canvas.Font.Name := '�ө���';
      for  aIndex := 7 to ListBox1.Items.Count - 1 do
        printer.Canvas.TextOut( 200, ( 200 + aIndex *90 ),listbox1.Items[aIndex] );
    finally
      Printer.EndDoc;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA08.chkStampTaxClick(Sender: TObject);
begin
  txtStampTax.Enabled := chkStampTax.Checked;
  if not txtStampTax.Enabled then txtStampTax.Value := 0;
  if chkStampTax.Checked then CheckBox1.Checked := False;
end;

{ ---------------------------------------------------------------------------- }


procedure TfrmInvA08.CheckBox1Click(Sender: TObject);
begin
  if CheckBox1.Checked then
  begin
    chkStampTax.Checked := False;
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
