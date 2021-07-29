unit frmMainU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs;

type
  TfrmMain = class(TForm)
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
    function GetFileDate(FileName: string): Integer;
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

uses IniFiles;

{$R *.DFM}

function TfrmMain.GetFileDate(FileName: string): Integer;
var
  fd: Integer;
begin
  Result := -1;
  fd := FileOpen(FileName, fmOpenRead);
  if fd < 0 then Exit;
  Result := FileGetDate(fd);
  FileClose(fd);
end;

procedure TfrmMain.FormCreate(Sender: TObject);
var
  ii: Integer;
  Inif: TIniFile;
  IniName, SrcFile, DstFile, DownPath: string;
begin
  try
    if ParamCount = 0 then Exit;

    IniName := ExtractFilePath(Application.ExeName) + 'UC.INI';
    Inif:= TIniFile.Create(IniName);
    DownPath := Inif.ReadString('SYSTEM', 'DOWNLOAD', '');
    Inif.Free;

    for ii:=1 to ParamCount do
    begin
      SrcFile := DownPath + '\' + ParamStr(ii);
      DstFile := ExtractFilePath(Application.ExeName) + ParamStr(ii);
      if FileExists(SrcFile) then
        if GetFileDate(SrcFile) <> GetFileDate(DstFile) then
          CopyFile(PChar(SrcFile), PChar(DstFile), False);
    end;

  finally
    Application.Terminate;
  end;
end;

end.





