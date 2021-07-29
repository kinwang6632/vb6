unit frmInvA07_5;

interface

uses
 Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ExtCtrls, dtmMainU, ADODB, DBClient, DB,
  cxControls, cxContainer,  cxEdit, cxLabel, Menus, cxLookAndFeelPainters,
  cxButtons, cxGraphics, cxDropDownEdit, cxRadioGroup, cxGroupBox, cxTextEdit,
  cxMaskEdit, dxSkinsCore, dxSkinsDefaultPainters, dxSkinBlack, dxSkinBlue,
  dxSkinCaramel, dxSkinCoffee;

type
  TForm1 = class(TForm)
    Panel1: TPanel;
    Label1: TLabel;
    Panel2: TPanel;
    Label3: TLabel;
    Label4: TLabel;
    lblMsg: TLabel;
    btnExit: TcxButton;
    btnOK: TcxButton;
    txtDisNoSt: TcxMaskEdit;
    cxGroupBox1: TcxGroupBox;
    Label2: TLabel;
    rdoDrop: TcxRadioButton;
    rdoRecover: TcxRadioButton;
    cmbReason: TcxComboBox;
    txtDisNoEd: TcxMaskEdit;
  private
    { Private declarations }
    FDisNo: String;
    FList: TStringList;
    procedure InitialComboBox;
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation
uses cbUtilis, frmMainU, dtmMainJU, dtmSOU;
{$R *.dfm}

{ TForm1 }

procedure TForm1.InitialComboBox;
var
  aParam: TComboBoxCreateParam;
begin
  aParam.Sql := ' SELECT ITEMID, DESCRIPTION FROM INV006 ORDER BY ITEMID ';
  aParam.KeyField := 'ITEMID';
  aParam.DescField := 'DESCRIPTION';
  aParam.AddAllText := False;
  CreateCxComboBoxItem( cmbReason, aParam );
  if ( cmbReason.Properties.Items.Count > 0 ) then cmbReason.ItemIndex := 0;
end;

end.
