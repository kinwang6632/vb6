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
    FConditionCompText: String;
    FLastActiveControl: TWinControl;
  public
    { Public declarations }
    property ConditionText: String read FConditionText;
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
  if ( rdoCondition1.Checked ) then
  begin
    if ( txtPaperDateSt.Text = EmptyStr ) then
    begin
      aFocusControl := txtPaperDateSt;
      WarningMsg( '?????J????????(?_)?C' );
    end else
    if ( txtPaperDateEd.Text = EmptyStr ) then
    begin
      aFocusControl := txtPaperDateEd;
      WarningMsg( '?????J????????(??)?C' );
    end;
  end else
  if ( rdoCondition2.Checked ) then
  begin
    if ( txtYearMonth.Text = EmptyStr ) then
    begin
      aFocusControl := txtYearMonth;
      WarningMsg( '?????J?????~???C' );
    end;
  end else
  if ( rdoCondition3.Checked ) then
  begin
    if ( txtInvIdSt.Text = EmptyStr ) then
    begin
      aFocusControl := txtInvIdSt;
      WarningMsg( '?????J?o?????X(?_)?C' );
    end else
    if ( txtInvIdEd.Text = EmptyStr ) then
    begin
      aFocusControl := txtInvIdEd;
      WarningMsg( '?????J?o?????X(??)?C' );
    end;
  end else
  if ( rdoCondition4.Checked ) then
  begin
    if ( txtAllowanceNoSt.Text = EmptyStr ) then
    begin
      aFocusControl := txtAllowanceNoSt;
      WarningMsg( '????????????(?_)?C' );
    end else
    if ( txtAllowanceNoEd.Text = EmptyStr ) then
    begin
      aFocusControl := txtAllowanceNoEd;
      WarningMsg( '?????J????????(??)?C' );
    end;
  end else
  if ( rdoCondition5.Checked ) then
  begin
    if ( txtPaperNoSt.Text = EmptyStr ) then
    begin
      aFocusControl := txtPaperNoSt;
      WarningMsg( '?????J?????s??(?_)?C' );
    end else
    if ( txtPaperNoEd.Text = EmptyStr ) then
    begin
      aFocusControl := txtPaperNoEd;
      WarningMsg( '?????J?????s??(??)?C' );
    end;
  end else
  if ( rdoCondition6.Checked ) then
  begin
    if ( txtCustIdSt.Text = EmptyStr ) then
    begin
      aFocusControl := txtCustIdSt;
      WarningMsg( '?????J???s(?_)?C' );
    end else
    if ( txtCustIdEd.Text = EmptyStr ) then
    begin
      aFocusControl := txtCustIdEd;
      WarningMsg( '?????J???s(??)?C' );
    end;
  end;
  if Assigned( aFocusControl ) then
  begin
    Self.ActiveControl := aFocusControl;
    Exit;
  end;
  { ???????? }
  if ( rdoCondition1.Checked ) then
  begin
    FConditionText := Format(
      ' AND A.PAPERDATE BETWEEN TO_DATE( ''%s'', ''YYYYMMDD'' ) ' +
      ' AND TO_DATE( ''%s'', ''YYYYMMDD'' ) ',
      [FormatDateTime( 'YYYYMMDD', txtPaperDateSt.Date ),
       FormatDateTime( 'YYYYMMDD', txtPaperDateEd.Date )] );
    FLastActiveControl := txtPaperDateSt;
  end else
  { ?????~?? }
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
    WarningMsg( '?z???J???????~???????T?C' );
    ErrorText := EmptyStr;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmInvA07_2.txtPaperDateStPropertiesValidate(Sender: TObject;
  var DisplayValue: Variant; var ErrorText: TCaption; var Error: Boolean);
begin
  if Error then
  begin
    WarningMsg( '???J???????????T?C' );
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
    WarningMsg( '?o?????????~, ?????s???J?C' );
  end else
    ErrorText := EmptyStr;
end;

{ ---------------------------------------------------------------------------- }

end.
