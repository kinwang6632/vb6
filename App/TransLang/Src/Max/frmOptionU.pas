unit frmOptionU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls, StdCtrls, Buttons, FileCtrl,IniFiles, ComCtrls;

type
  TfrmOption = class(TForm)
    Bevel1: TBevel;
    Panel1: TPanel;
    BitBtn1: TBitBtn;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    DirectoryListBox1: TDirectoryListBox;
    Label4: TLabel;
    Label3: TLabel;
    Label2: TLabel;
    Label1: TLabel;
    DriveComboBox1: TDriveComboBox;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    RadioGroup1: TRadioGroup;
    BitBtn3: TBitBtn;
    Label5: TLabel;
    Label6: TLabel;
    RadioGroup2: TRadioGroup;
    TabSheet4: TTabSheet;
    GroupBox1: TGroupBox;
    BitBtn2: TBitBtn;
    procedure BitBtn1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    OptionsIni: TIniFile;
    SysPath: string;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmOption: TfrmOption;

implementation

{$R *.DFM}

procedure TfrmOption.BitBtn1Click(Sender: TObject);
begin
  OptionsIni.WriteString('Path Option','Item0',DirectoryListBox1.Directory);
  OptionsIni.WriteString('Loader Option','Item0',inttostr(RadioGroup1.itemindex));
  OptionsIni.WriteString('Exporter Option','Item0',inttostr(RadioGroup2.itemindex));
  OptionsIni.Free;
  Close;
end;

procedure TfrmOption.FormShow(Sender: TObject);
begin
//  SysPath := DirectoryListBox1.Directory;
  SysPath := GetCurrentDir;
  OptionsIni := TIniFile.Create(SysPath+'\Ini\Options.ini');
	try
	with OptionsIni do
    begin
 			DirectoryListBox1.Directory := ReadString('Path Option','Item0','C:\');
      RadioGroup1.ItemIndex := strtoint(ReadString('Loader Option','Item0','0'));
      RadioGroup2.ItemIndex := strtoint(ReadString('Exporter Option','Item0','1'));
    end;
	except
		DirectoryListBox1.Directory := 'C:\';
	end;
end;

procedure TfrmOption.BitBtn2Click(Sender: TObject);
begin
	try
	with OptionsIni do
    begin
 			DirectoryListBox1.Directory := ReadString('Default','Item0','C:\');
      RadioGroup1.ItemIndex := strtoint(ReadString('Default','Item1','0'));
      RadioGroup2.ItemIndex := strtoint(ReadString('Default','Item2','1'));
    end;
	except
		DirectoryListBox1.Directory := 'C:\';
	end;
end;

procedure TfrmOption.BitBtn3Click(Sender: TObject);
begin
  Close;
end;

procedure TfrmOption.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  DirectoryListBox1.Directory := SysPath; //回復系統路徑
end;

end.
