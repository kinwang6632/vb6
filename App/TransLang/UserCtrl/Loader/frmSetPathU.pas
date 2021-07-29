unit frmSetPathU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls;

type
  TfrmSetPath = class(TForm)
    Bevel1: TBevel;
    Label1: TLabel;
    edtUCPath: TEdit;
    edtDownPath: TEdit;
    Label2: TLabel;
    OKBtn: TButton;
    CancelBtn: TButton;
    procedure OKBtnClick(Sender: TObject);
  private
    FDownPath: string;
    FUCPath: string;
    procedure SetDownPath(const Value: string);
    procedure SetUCPath(const Value: string);
    { Private declarations }
  public
    { Public declarations }
    property UCPath: string read FUCPath write SetUCPath;
    property DownPath: string read FDownPath write SetDownPath;
  end;

var
  frmSetPath: TfrmSetPath;

implementation

{$R *.DFM}

{ TfrmSetPath }

procedure TfrmSetPath.SetDownPath(const Value: string);
begin
  FDownPath := Value;
  edtDownPath.Text := FDownPath;
end;

procedure TfrmSetPath.SetUCPath(const Value: string);
begin
  FUCPath := Value;
  edtUCPath.Text := FUCPath;
end;

procedure TfrmSetPath.OKBtnClick(Sender: TObject);
begin
  if edtUCPath.Text = '' then
  begin
    MessageDlg('請輸入系統檔案路徑', mtWarning, [mbOk], 0);
    Exit;
  end;
  if edtDownPath.Text = '' then
  begin
    MessageDlg('請輸入軟體儲存路徑', mtWarning, [mbOk], 0);
    Exit;
  end;
  FDownPath := edtDownPath.Text;
  FUCPath := edtUCPath.Text;
  MOdalResult := mrOk;
end;

end.
