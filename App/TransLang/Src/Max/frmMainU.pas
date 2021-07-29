unit frmMainU;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ImgList, ComCtrls, ToolWin, ExtCtrls, Menus, Grids, DBGrids, Db, ADODB,
  FileCtrl, StdCtrls, DBClient, DBCtrls, Mask,
  IniFiles, Variants;

type
  TfrmMain = class(TForm)
    Bevel1: TBevel;
    ToolBar1: TToolBar;
    tbrLoad: TToolButton;
    tbrExport: TToolButton;
    tbrOptions: TToolButton;
    ImageList1: TImageList;
    ImageList2: TImageList;
    pmnLoader: TPopupMenu;
    dgrLoader: TDBGrid;
    Panel1: TPanel;
    StatusBar1: TStatusBar;
    tbrExit: TToolButton;
    Splitter1: TSplitter;
    dgrExporter: TDBGrid;
    dsrExporter: TDataSource;
    dsrLoader: TDataSource;
    Panel2: TPanel;
    dnvLoader: TDBNavigator;
    ProgressBar1: TProgressBar;
    Panel3: TPanel;
    dnvExporter: TDBNavigator;
		procedure tbrLoadClick(Sender: TObject);
		procedure tbrOptionsClick(Sender: TObject);
		procedure tbrExitClick(Sender: TObject);
    procedure dsrLoaderDataChange(Sender: TObject; Field: TField);
    procedure tbrExportClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure dsrExporterDataChange(Sender: TObject; Field: TField);
	private
		NewMenuItem: TMenuItem;
    OptionsIni: TIniFile;
		procedure DrawdownLoad(Sender: TObject); //指定給Menu OnClick Even的程序
 		procedure GetUnit(PrjFullName: string); //截取UnitFile字串
//    Function DeProjName(PrjFullName: string): string;
		{ Private declarations }
	public
    ProjName: string;
		{ Public declarations }
	end;

var
	frmMain: TfrmMain;

implementation

uses frmLoadU, frmOptionU, dtmMainU, dtmBasicU;



{$R *.DFM}

{function TfrmMain.DeProjName(PrjFullName:string): string;
var GetPos:integer;
    TempStr: string;
begin
  TempStr:=PrjFullName;
  GetPos:=Pos('\',TempStr);
  while GetPos > 0 do
  begin
    GetPos:=Pos('\',TempStr);
    TempStr:=copy(TempStr,GetPos+1,Length(TempStr)-GetPos);
  end;
  result:=TempStr;
end;}


Procedure TfrmMain.GetUnit(PrjFullName:string);
var LineList: TStringList;
    i, j, GetPos: integer;
    TempStr, PrjPath, sMark: string;
const FileStyle: array[1..4] of string = ('Form=', 'Class=', 'Module=', 'UserControl=');
begin
  //取得目前專案路徑....
  TempStr:=PrjFullName;
  GetPos:=Pos('\',TempStr);
  While GetPos > 0 do
  begin
    GetPos:=Pos('\',TempStr);
    PrjPath:=PrjPath+copy(TempStr,1,GetPos);
    TempStr:=copy(TempStr,GetPos+1,Length(TempStr)-GetPos);
  end;
  //取得專案內含的檔案並存入資料庫
  LineList := TStringList.Create;
  LineList.LoadFromFile(PrjFullName);
  TempStr:='';
  for i:= 1 to LineList.Count - 1 do
  begin
    for j:=1 to 4 do
    begin
      GetPos := Pos(FileStyle[j],LineList.Strings[i]);
      if GetPos = 1 then
      begin
        TempStr := copy(LineList.Strings[i],Length(FileStyle[j])+1,Length(LineList.Strings[i])-Length(FileStyle[j]));
        GetPos:=Pos(';',TempStr);
        While GetPos > 0 do
        begin
          GetPos:=Pos(';',TempStr);
          TempStr:=copy(TempStr,GetPos+1,Length(TempStr)-GetPos);
        end;
        TempStr:=PrjPath+trim(TempStr);
        //儲存入資料庫
        With dsrLoader.DataSet do
        begin
          sMark:='';
          if Lookup('sUnitName',TempStr,'sProjName') <> Null then
            sMark:='*';
          Append;
//          FieldByName('sProjName').asstring:=DeProjName(ProjName);
          FieldByName('sProjName').asstring:=ProjName;
          FieldByName('sUnitName').asstring:=TempStr;
          FieldByName('sRepeated').asstring:=sMark;
          Post;
        end;
      end;
    end;
  end;
end;

procedure TfrmMain.tbrLoadClick(Sender: TObject);
var i:integer;
begin
   frmLoad := TfrmLoad.Create(nil);
   frmLoad.ShowModal;
   if frmLoad.Tag in [1,2] then
   begin
     if OptionsIni.ReadString('Loader Option','Item0','0') = '0' then
       dtmMain.CDSTmpQry.EmptyDataSet;
     ProgressBar1.Visible := true;
//     ProgressBar1.Step := 100 div frmLoad.ListBox1.Items.Count;
     ProgressBar1.Max := frmLoad.ListBox1.Items.Count;
     ProgressBar1.Step := 1;
     for i:=0 to frmLoad.ListBox1.Items.Count - 1 do
     begin
       if (frmLoad.Tag = 1) or (frmLoad.ListBox1.Selected[i]) then
       begin
         NewMenuItem := TMenuItem.Create(nil);
         NewMenuItem.Caption := frmLoad.ListBox1.Items.Strings[i];  //取得Project File置入DrawDownMenu
         pmnLoader.Items.Add(NewMenuItem);
         pmnLoader.Items[pmnLoader.Items.Count-1].OnClick := DrawdownLoad;
         ProjName:=NewMenuItem.Caption;
         GetUnit(ProjName);
       end;
     ProgressBar1.StepIt;
     end;
     if ProgressBar1.Position < ProgressBar1.Max then
       ProgressBar1.Position := ProgressBar1.Max;
     Showmessage('Finish Load!');
     ProgressBar1.Position := 0;
     ProgressBar1.Visible := false;     
   end;
   frmLoad.DirectoryListBox1.Directory := frmLoad.SysPath;//回復系統路徑
   frmLoad.free;
end;

procedure TfrmMain.tbrOptionsClick(Sender: TObject);
begin
	frmOption := TfrmOption.Create(nil);
	frmOption.ShowModal;
end;

procedure TfrmMain.tbrExitClick(Sender: TObject);
begin
	close;
end;

procedure TfrmMain.DrawdownLoad(Sender: TObject);
begin
  if OptionsIni.ReadString('Loader Option','Item0','0') = '0' then
    dtmMain.CDSTmpQry.EmptyDataSet;
  ProjName:=TMenuItem(Sender).caption;
  GetUnit(ProjName);
  ProgressBar1.Visible := true;
  ProgressBar1.Position := 100;
  Showmessage('Finish Load!');
  ProgressBar1.Position := 0;
  ProgressBar1.Visible := false;
end;
procedure TfrmMain.dsrLoaderDataChange(Sender: TObject; Field: TField);
begin
  if dtmMain.CDSTmpQry.active then
  begin
    if dtmMain.CDSTmpQry.RecordCount > 0 then
      tbrExport.Enabled := true
    else
      tbrExport.Enabled := false;
    StatusBar1.Panels[0].Text := 'Loader Record: '+inttostr(dtmMain.CDSTmpQry.recordcount);
  end;
end;

procedure TfrmMain.tbrExportClick(Sender: TObject);
begin
  with dtmMain do
  begin
	  if OptionsIni.ReadString('Exporter Option','Item0','1') = '0' then
    	ExecSQL(1);
    ProgressBar1.Visible := true;
    //ProgressBar1.step:=100 div CDSTmpQry.RecordCount;
    ProgressBar1.Max := CDSTmpQry.RecordCount;
    ProgressBar1.Step := 1;
  	CDSTmpQry.First;
  	while not CDSTmpQry.Eof do
  	begin
    	if ADOPrjInfoQry.Lookup('sUnitName',CDSTmpQry.Fieldbyname('sUnitName').asstring,
      'sProjName') = Null then
      	ADOPrjInfoQry.AppendRecord([nil,CDSTmpQry.fieldbyname('sProjName').asstring,
      		CDSTmpQry.fieldbyname('sUnitName').asstring]);
      CDSTmpQry.Next;
      ProgressBar1.StepIt;
	  end;
    if ProgressBar1.Position < ProgressBar1.Max then
      ProgressBar1.Position := ProgressBar1.Max;
    showmessage('Finish Export!');
    ProgressBar1.Position := 0;
    ProgressBar1.Visible := false;
  end;
//  dtmMain.InsertCmd;
end;

procedure TfrmMain.FormCreate(Sender: TObject);
var SysPath:string;
begin
  SysPath:=GetCurrentdir;
  OptionsIni := TIniFile.Create(SysPath+'.\Ini\'+'Options.ini');
end;

procedure TfrmMain.dsrExporterDataChange(Sender: TObject; Field: TField);
begin
  if dtmMain.ADOPrjInfoQry.Active then
    StatusBar1.Panels[1].Text :='Exporter Record: '+inttostr(dtmMain.ADOPrjInfoQry.recordcount);
end;

end.

