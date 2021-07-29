unit frmInvA07_2U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, Menus, Buttons, StdCtrls,
  cxTextEdit, cxMaskEdit, cxDropDownEdit, cxCalendar,
  cxRadioGroup, cxControls, cxContainer, cxEdit, cxGroupBox,
  cxLookAndFeelPainters, dxSkinsCore, dxSkinsDefaultPainters, cxButtons,
  dxSkinBlack, dxSkinBlue, dxSkinCaramel, dxSkinCoffee;


type
  TfrmInvA07_2 = class(TForm)
    Panel2: TPanel;
    Panel1: TPanel;
    ScrollBox3: TScrollBox;
    Bevel1: TBevel;
    btnOk: TcxButton;
    btnCancel: TcxButton;
    cxGroupBox1: TcxGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    rdoCondition1: TcxRadioButton;
    rdoCondition2: TcxRadioButton;
    rdoCondition3: TcxRadioButton;
    rdoCondition4: TcxRadioButton;
    rdoCondition5: TcxRadioButton;
    rdoCondition6: TcxRadioButton;
    txtPaperDateSt: TcxDateEdit;
    txtPaperDateEd: TcxDateEdit;
    txtAllowanceNoSt: TcxTextEdit;
    txtAllowanceNoEd: TcxTextEdit;
    cxGroupBox2: TcxGroupBox;
    txtYearMonth: TcxMaskEdit;
    rdoComp1: TcxRadioButton;
    rdoComp2: TcxRadioButton;
    txtInvIdSt: TcxMaskEdit;
    txtInvIdEd: TcxMaskEdit;
    cxGroupBox3: TcxGroupBox;
    rdoSourceAll: TcxRadioButton;
    rdoSource0: TcxRadioButton;
    rdoSource1: TcxRadioButton;
    txtPaperNoSt: TcxTextEdit;
    Label4: TLabel;
    txtPaperNoEd: TcxTextEdit;
    txtCustIdSt: TcxTextEdit;
    Label5: TLabel;
    txtCustIdEd: TcxTextEdit;
    cxGroupBox4: TcxGroupBox;
    rdoInvKind1: TcxRadioButton;
    rdoInvKind0: TcxRadioButton;
    rdoInvKind2: TcxRadioButton;
    cxGroupBox5: TcxGroupBox;
    rdoUpdFlag1: TcxRadioButton;
    rdoUpdFlag0: TcxRadioButton;
    rdoUpdFlag2: TcxRadioButton;
    cxGroupBox6: TcxGroupBox;
    rdoInvId: TcxRadioButton;
    rdoPaperDate: TcxRadioButton;
    rdoAllowanceNo: TcxRadioButton;
    cxGroupBox7: TcxGroupBox;
    rdoIsObsolete1: TcxRadioButton;
    rdoIsObsolete0: TcxRadioButton;
    rdoIsObsolete2: TcxRadioButton;
    cxGroupBox8: TcxGroupBox;
    rdoBussiness1: TcxRadioButton;
    rdoBussiness0: TcxRadioButton;
    rdoBussiness2: TcxRadioButton;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnOkClick(Sender: TObject);
    procedure txtPaperDateStEnter(Sender: TObject);
    procedure txtYearMonthEnter(Sender: TObject);
    procedure txtInvIdStEnter(Sender: TObject);
    procedure txtAllowanceNoStEnter(Sender: TObject);
    procedure txtYearMonthPropertiesValidate(Sender: TObject;
      var DisplayValue: Variant; var ErrorText: TCaption;
      var Error: Boolean);
    procedure txtPaperDateStPropertiesValidate(Sender: TObject;
      var DisplayValue: Variant; var ErrorText: TCaption;
      var Error: Boolean);
    procedure txtPaperDateStExit(Sender: TObject);
    procedure txtInvIdStExit(Sender: TObject);
    procedure txtAllowanceNoStExit(Sender: TObject);
    procedure txtInvIdStPropertiesValidate(Sender: TObject;
      var DisplayValue: Variant; var ErrorText: TCaption;
      var Error: Boolean);
    procedure txtPaperNoStEnter(Sender: TObject);
    procedure txtCustIdStEnter(Sender: TObject);
    procedure txtPaperNoStExit(Sender: TObject);
    procedure txtCustIdStExit(Sender: TObject);
  private
    { Private declarations }
    FConditionText: String;
    FOrderText: String;
    FConditionCompText: String;
    FLastActiveControl: TWinControl;
  public
    { Public declarations }
    property ConditionText: String read FConditionText;
    property OrderText: String read FOrderText;
    property ConditionCompText: String read FConditionCompText;
  end;

var
  frmInvA07_2: TfrmInvA07_2;

implementation

uses cbUtilis, frmMainU, dtmMainU;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_2.FormCreate(Sender: TObject);
begin
  FConditionText := EmptyStr;
  FConditionCompText := EmptyStr;
  rdoComp1.Checked := True;
  rdoSourceAll.Checked := True;
  rdoCondition1.Checked := True;
  Self.ActiveControl := txtPaperDateSt;
  FLastActiveControl := nil;
end;


{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_2.FormShow(Sender: TObject);
begin
  if Assigned( FLastActiveControl ) then
    if FLastActiveControl.CanFocus then
      FLastActiveControl.SetFocus;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_2.FormDestroy(Sender: TObject);
begin
   {...}
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_2.btnOkClick(Sender: TObject);
var
  aFocusControl: TWinControl;
  aYearMonth, aSource: String;
begin
  aFocusControl := nil;
  txtPaperDateSt.PostEditValue;
  txtPaperDateEd.PostEditValue;
  txtAllowanceNoSt.PostEditValue;
  txtAllowanceNoEd.PostEditValue;
  txtYearMonth.PostEditValue;
  txtInvIdSt.PostEditValue;
  txtInvIdEd.PostEditValue;
  txtPaperNoSt.PostEditValue;
  txtPaperNoEd.PostEditValue;
  txtCustIdSt.PostEditValue;
  txtCustIdEd.PostEditValue;
  if ( rdoCondition1.Checked ) then
  begin
    if ( txtPaperDateSt.Text = EmptyStr ) then
    begin
      aFocusControl := txtPaperDateSt;
      WarningMsg( '請輸入單據日期(起)。' );
    end else
    if ( txtPaperDateEd.Text = EmptyStr ) then
    begin
      aFocusControl := txtPaperDateEd;
      WarningMsg( '請輸入單據日期(迄)。' );
    end;
  end else
  if ( rdoCondition2.Checked ) then
  begin
    if ( txtYearMonth.Text = EmptyStr ) then
    begin
      aFocusControl := txtYearMonth;
      WarningMsg( '請輸入申報年月。' );
    end;
  end else
  if ( rdoCondition3.Checked ) then
  begin
    if ( txtInvIdSt.Text = EmptyStr ) then
    begin
      aFocusControl := txtInvIdSt;
      WarningMsg( '請輸入發票號碼(起)。' );
    end else
    if ( txtInvIdEd.Text = EmptyStr ) then
    begin
      aFocusControl := txtInvIdEd;
      WarningMsg( '請輸入發票號碼(迄)。' );
    end;
  end else
  if ( rdoCondition4.Checked ) then
  begin
    if ( txtAllowanceNoSt.Text = EmptyStr ) then
    begin
      aFocusControl := txtAllowanceNoSt;
      WarningMsg( '請輸折讓單號(起)。' );
    end else
    if ( txtAllowanceNoEd.Text = EmptyStr ) then
    begin
      aFocusControl := txtAllowanceNoEd;
      WarningMsg( '請輸入折讓單號(迄)。' );
    end;
  end else
  if ( rdoCondition5.Checked ) then
  begin
    if ( txtPaperNoSt.Text = EmptyStr ) then
    begin
      aFocusControl := txtPaperNoSt;
      WarningMsg( '請輸入單據編號(起)。' );
    end else
    if ( txtPaperNoEd.Text = EmptyStr ) then
    begin
      aFocusControl := txtPaperNoEd;
      WarningMsg( '請輸入單據編號(迄)。' );
    end;
  end else
  if ( rdoCondition6.Checked ) then
  begin
    if ( txtCustIdSt.Text = EmptyStr ) then
    begin
      aFocusControl := txtCustIdSt;
      WarningMsg( '請輸入客編(起)。' );
    end else
    if ( txtCustIdEd.Text = EmptyStr ) then
    begin
      aFocusControl := txtCustIdEd;
      WarningMsg( '請輸入客編(迄)。' );
    end;
  end;
  if Assigned( aFocusControl ) then
  begin
    Self.ActiveControl := aFocusControl;
    Exit;
  end;
  { 單據日期 }
  if ( rdoCondition1.Checked ) then
  begin
    FConditionText := Format(
      ' AND A.PAPERDATE BETWEEN TO_DATE( ''%s'', ''YYYYMMDD'' ) ' +
      ' AND TO_DATE( ''%s'', ''YYYYMMDD'' ) ',
      [FormatDateTime( 'YYYYMMDD', txtPaperDateSt.Date ),
       FormatDateTime( 'YYYYMMDD', txtPaperDateEd.Date )] );
    FLastActiveControl := txtPaperDateSt;
  end else
  { 申報年月 }
  if ( rdoCondition2.Checked ) then
  begin
    aYearMonth :=
      Copy(  txtYearMonth.Text, 1, 4 ) +
      Lpad( Copy( txtYearMonth.Text, 6, 2 ), 2, '0' );
    FConditionText := Format( ' AND A.YEARMONTH = ''%s'' ', [aYearMonth] );
    FLastActiveControl := txtYearMonth;
  end else
  if ( rdoCondition3.Checked ) then
  begin
    FConditionText := Format( ' AND A.INVID BETWEEN ''%s'' AND ''%s'' ',
      [txtInvIdSt.Text, txtInvIdEd.Text] );
    FLastActiveControl := txtInvIdSt;
  end else
  if ( rdoCondition4.Checked ) then
  begin
    FConditionText := Format( ' AND A.ALLOWANCENO BETWEEN ''%s'' AND ''%s'' ',
      [txtAllowanceNoSt.Text, txtAllowanceNoEd.Text] );
    FLastActiveControl := txtAllowanceNoSt;
  end else
  if ( rdoCondition5.Checked ) then
  begin
    FConditionText := Format( ' AND A.PAPERNO BETWEEN ''%s'' AND ''%s'' ',
      [txtPaperNoSt.Text, txtPaperNoEd.Text] );
    FLastActiveControl := txtPaperNoSt;
  end else
  if ( rdoCondition6.Checked ) then
  begin
    FConditionText := Format( ' AND LPAD( A.CUSTID, 8, ''0'' ) BETWEEN ''%s'' AND ''%s'' ',
      [lpad( txtCustIdSt.Text, 8, '0' ), lpad( txtCustIdEd.Text, 8, '0' )] );
    FLastActiveControl := txtCustIdSt;
  end;
  {}
  if ( rdoComp1.Checked ) then
  begin
    FConditionText := ( FConditionText + Format(
      ' AND A.COMPID = ''%s'' ', [dtmMain.getCompID] ) );
    FConditionCompText := dtmMain.getCompName;
  end else
  begin
    FConditionCompText := dtmMain.getCompNameEx;
  end;
  {}
  if ( rdoSource0.Checked ) or ( rdoSource1.Checked ) then
  begin
    aSource := '0';
    if ( rdoSource1.Checked ) then aSource := '1';
    FConditionText := ( FConditionText + Format(
      ' AND A.SOURCE = ''%s'' ', [aSource] ) );
  end;
  {#5922 增加 發票種類 By Kin 2011/02/16}
  if (rdoInvKind1.Checked) or ( rdoInvKind0.Checked) then
  begin
    if rdoInvKind0.Checked then
    begin
      FConditionText := ( FConditionText +
        ' AND C.INVOICEKIND = 0' );
    end else
    begin
      FConditionText := ( FConditionText +
        ' AND C.INVOICEKIND = 1' );
    end;
  end;
  {#5922 增加 上傳註記 By Kin 2011/02/16}
  if ( rdoUpdFlag0.Checked ) or ( rdoUpdFlag1.Checked ) then
  begin
    if rdoUpdFlag0.Checked then
    begin
      FConditionText := ( FConditionText +
        ' AND A.UPLOADFLAG = 0 ');
    end else
    begin
      FConditionText := ( FConditionText +
        ' AND A.UPLOADFLAG = 1 ');
    end;
  end;
  {#6843增加是否作廢 By Kin 2014/08/05}
  if ( rdoIsObsolete1.Checked ) or ( rdoIsObsolete0.Checked ) then
  begin
    if ( rdoIsObsolete0.Checked ) then
    begin
       FConditionText := ( FConditionText +
          ' AND A.ISOBSOLETE = ''N'' ');
    end else
    begin
       FConditionText := ( FConditionText +
          ' AND A.ISOBSOLETE = ''Y'' ');
    end;
  end;
  if (rdoBussiness1.Checked ) or (rdoBussiness0.Checked ) then
  begin
    if (rdoBussiness0.Checked ) then
    begin
      FConditionText := ( FConditionText +
          ' AND A.BUSINESSID IS NOT NULL ' );
    end else
    begin
      FConditionText := ( FConditionText +
          ' AND A.BUSINESSID IS  NULL ' );
    end;
  end;
  if rdoPaperDate.Checked then
  begin
    FOrderText := ' ORDER BY A.CompId,A.PaperDate ';
  end
  else if rdoInvId.Checked then
  begin
    FOrderText := ' ORDER BY A.CompId,A.InvID ';
  end else
  begin
    FOrderText := ' ORDER BY A.CompId,A.AllowanceNo ';
  end;
  Self.ModalResult := mrOk;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_2.txtPaperDateStEnter(Sender: TObject);
begin
  rdoCondition1.Checked := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_2.txtYearMonthEnter(Sender: TObject);
begin
  rdoCondition2.Checked := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_2.txtInvIdStEnter(Sender: TObject);
begin
  rdoCondition3.Checked := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_2.txtAllowanceNoStEnter(Sender: TObject);
begin
  rdoCondition4.Checked := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_2.txtPaperNoStEnter(Sender: TObject);
begin
  rdoCondition5.Checked := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_2.txtCustIdStEnter(Sender: TObject);
begin
  rdoCondition6.Checked := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_2.txtYearMonthPropertiesValidate(Sender: TObject;
  var DisplayValue: Variant; var ErrorText: TCaption; var Error: Boolean);
begin
  if Error then
  begin
    WarningMsg( '您輸入的申報年月不正確。' );
    ErrorText := EmptyStr;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_2.txtPaperDateStPropertiesValidate(Sender: TObject;
  var DisplayValue: Variant; var ErrorText: TCaption; var Error: Boolean);
begin
  if Error then
  begin
    WarningMsg( '輸入的日期不正確。' );
    ErrorText := EmptyStr;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_2.txtPaperDateStExit(Sender: TObject);
begin
  txtPaperDateEd.Date := txtPaperDateSt.Date;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_2.txtInvIdStExit(Sender: TObject);
begin
  txtInvIdEd.Text := txtInvIdSt.Text;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_2.txtAllowanceNoStExit(Sender: TObject);
begin
  txtAllowanceNoEd.Text := txtAllowanceNoSt.Text;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_2.txtPaperNoStExit(Sender: TObject);
begin
  txtPaperNoEd.Text := txtPaperNoSt.Text;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_2.txtCustIdStExit(Sender: TObject);
begin
  txtCustIdEd.Text := txtCustIdSt.Text;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_2.txtInvIdStPropertiesValidate(Sender: TObject;
  var DisplayValue: Variant; var ErrorText: TCaption; var Error: Boolean);
begin
  if ( Error ) then
  begin
    ErrorText := EmptyStr;
    WarningMsg( '發票格式有誤, 請重新輸入。' );
  end else
    ErrorText := EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

end.
