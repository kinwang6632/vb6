unit frmSelectPathU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, FileCtrl;

type
  TfrmSelectPath = class(TForm)
    DriveComboBox1: TDriveComboBox;
    DirectoryListBox1: TDirectoryListBox;
    lblPath: TLabel;
    btnOK: TButton;
    btnCancel: TButton;
    procedure btnOKClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    sG_Path,sG_OldPath : String;
  end;


  function SelectPath(var sI_Path : String): TModalResult ;

var
  frmSelectPath: TfrmSelectPath;

implementation

{$R *.dfm}

function SelectPath(var sI_Path : String): TModalResult ;
var
  aFrm : TfrmSelectPath ;
begin
  aFrm := TfrmSelectPath.Create(Application) ;
  Result := aFrm.ShowModal ;
  sI_Path := aFrm.sG_Path;
  aFrm.Free;
end;

procedure TfrmSelectPath.btnOKClick(Sender: TObject);
begin
    sG_Path := lblPath.Caption;
    Close;
end;

procedure TfrmSelectPath.FormShow(Sender: TObject);
begin
    sG_OldPath := lblPath.Caption;
end;

procedure TfrmSelectPath.btnCancelClick(Sender: TObject);
begin
//    sG_Path := 'C:';
    Close;
end;

end.
