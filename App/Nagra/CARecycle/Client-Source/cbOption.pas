unit cbOption;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, cxGraphics, StdCtrls, cxControls, cxContainer, cxEdit,
  cxTextEdit, cxMaskEdit, cxDropDownEdit, Menus, cxLookAndFeelPainters,
  cxButtons, ExtCtrls;

type
  TfmOption = class(TForm)
    ScrollBox1: TScrollBox;
    Label1: TLabel;
    cmbSo: TcxComboBox;
    Panel1: TPanel;
    Bevel1: TBevel;
    btnOk: TcxButton;
    btnCancel: TcxButton;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnOkClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fmOption: TfmOption;

implementation

uses cbUtilis, cbAppClass, cbSoModule;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TfmOption.FormCreate(Sender: TObject);
var
  aIndex, aActiveIndex: Integer;
  aText: String;
begin
  cmbSo.Properties.Items.Clear;
  aActiveIndex := -1;
  for aIndex := 0 to SoDataModule.SoList.Count - 1 do
  begin
    if ( TSo( SoDataModule.SoList[aIndex] ).CompCode <> '-1' ) then
    begin
      aText := Format( '%s-%s', [
        TSo( SoDataModule.SoList[aIndex] ).CompCode,
        TSo( SoDataModule.SoList[aIndex] ).CompName] );
      cmbSo.Properties.Items.Add( aText );
      if TSo( SoDataModule.SoList[aIndex] ).Active then
        aActiveIndex := aIndex;
    end;
  end;
  if ( cmbSo.Properties.Items.Count > 0 ) and ( aActiveIndex >= 0 ) then
    cmbSo.ItemIndex := aActiveIndex;
end;

{ ---------------------------------------------------------------------------- }

procedure TfmOption.FormShow(Sender: TObject);
begin

end;

{ ---------------------------------------------------------------------------- }

procedure TfmOption.FormDestroy(Sender: TObject);
begin
  
end;

{ ---------------------------------------------------------------------------- }

procedure TfmOption.btnOkClick(Sender: TObject);
var
  aCompCode, aText, aMsg: String;
begin
  if ( cmbSo.ItemIndex < 0 ) then Exit;
  aText := cmbSo.Properties.Items[cmbSo.ItemIndex];
  aCompCode := Copy( aText, 1, Pos( '-', aText ) - 1 );
  if not SoDataModule.SaveDefaultSo( aCompCode, aMsg ) then
  begin
    WarningMsg( Format( '無法設定預設系統台, 訊息: %s。', [aMsg] ) );
    Exit;
  end;
  Self.ModalResult := mrOk;
end;

{ ---------------------------------------------------------------------------- }

end.
