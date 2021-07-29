unit frmLoadU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, FileCtrl,IniFiles;

type
  TfrmLoad = class(TForm)
    Bevel1: TBevel;
    Panel1: TPanel;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    FileListBox1: TFileListBox;
    DirectoryListBox1: TDirectoryListBox;
    Panel2: TPanel;
    Label3: TLabel;
    Label2: TLabel;
    Label1: TLabel;
    DriveComboBox1: TDriveComboBox;
    ListBox1: TListBox;
    Panel3: TPanel;
    SpeedButton2: TSpeedButton;
    SpeedButton1: TSpeedButton;
    SpeedButton3: TSpeedButton;
    SpeedButton4: TSpeedButton;
    BitBtn3: TBitBtn;
		procedure BitBtn1Click(Sender: TObject);
		procedure FormShow(Sender: TObject);
    procedure DirectoryListBox1Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure FileListBox1Change(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
    procedure SpeedButton4Click(Sender: TObject);
    procedure ListBox1Click(Sender: TObject);
  private
		{ Private declarations }
	public
 		SysPath: string;
		{ Public declarations }
	end;

var
	frmLoad: TfrmLoad;

implementation

{$R *.DFM}

procedure TfrmLoad.BitBtn1Click(Sender: TObject);
begin
//按鍵為OK傳回值為1
//如為Cancel則傳回值為0
	Self.Tag := TBitbtn(Sender).Tag;
  ListBox1.Items.SaveToFile(SysPath+'\Ini\Loader.txt');
	close;
end;

procedure TfrmLoad.FormShow(Sender: TObject);
var OptionsIni: Tinifile;
begin
  SysPath := GetCurrentDir;
  ListBox1.Items.LoadFromFile(SysPath+'\Ini\Loader.txt');
	OptionsIni := TIniFile.Create(SysPath+'\Ini\Options.ini');
  with OptionsIni do
    with DirectoryListBox1 do
			Directory := ReadString('Path Option','Item0','C:\');
  if ListBox1.Items.Count < 1 then
  begin
  	Bitbtn1.Enabled := false;
   	Bitbtn2.Enabled := false;
    SpeedButton2.Enabled:=false;
    SpeedButton4.Enabled:=false;
  end
  else
  begin
    Bitbtn1.Enabled := true;
    SpeedButton2.Enabled:=true;
    SpeedButton4.Enabled:=true;
  end;
end;

procedure TfrmLoad.DirectoryListBox1Click(Sender: TObject);
begin
  DirectoryListBox1.OpenCurrent;
  FileListBox1.ItemIndex := 0;
end;

procedure TfrmLoad.SpeedButton1Click(Sender: TObject);
begin
  if ListBox1.Items.IndexOf(FileListBox1.FileName) < 0 then
  begin
    ListBox1.Items.Add(FileListBox1.FileName);
    Bitbtn1.Enabled := true;
    SpeedButton2.Enabled := true;
    SpeedButton4.Enabled := true;
  end;
end;

procedure TfrmLoad.SpeedButton2Click(Sender: TObject);
begin
  ListBox1.Items.Delete(ListBox1.ItemIndex);
  if ListBox1.Items.Count < 1 then
  begin
  	Bitbtn1.Enabled := false;
  	Bitbtn2.Enabled := false;
    SpeedButton2.Enabled :=false;
    SpeedButton4.Enabled :=false;
  end
  else
  begin
    Bitbtn1.Enabled := true;
    SpeedButton2.Enabled :=true;
    SpeedButton4.Enabled :=true;
  end;
end;

procedure TfrmLoad.FileListBox1Change(Sender: TObject);
begin
  if FileListBox1.Items.Count > 0 then
  begin
    SpeedButton1.Enabled := true;
    SpeedButton3.Enabled := true;
  end
  else
  begin;
    SpeedButton1.Enabled := false;
    SpeedButton3.Enabled := false;
  end;
end;

procedure TfrmLoad.SpeedButton3Click(Sender: TObject);
var i: integer;
begin
  for i:=0 to FileListBox1.Items.Count - 1 do
  begin
    FileListBox1.ItemIndex := i;
    SpeedButton1.OnClick(nil);
  end;
end;

procedure TfrmLoad.SpeedButton4Click(Sender: TObject);
begin
  ListBox1.Clear;
  SpeedButton2.Enabled:=false;
  SpeedButton4.Enabled:=false;
  Bitbtn1.Enabled:=false;
  Bitbtn2.Enabled:=false;
end;

procedure TfrmLoad.ListBox1Click(Sender: TObject);
begin
  if ListBox1.Items.Count > 0 then
    Bitbtn2.Enabled := true;
end;

end.

