unit frmTL0020U;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Forms, Dialogs, Controls, StdCtrls,
  Buttons, db, ComCtrls, ExtCtrls, FileCtrl;


type

  TfrmTL0020 = class(TForm)
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    SpeedButton1: TSpeedButton;
    pgbStatus: TProgressBar;
    Panel1: TPanel;
    pgcType: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    lblPath1: TLabel;
    SrcLabel: TLabel;
    DstLabel: TLabel;
    IncludeBtn: TSpeedButton;
    IncAllBtn: TSpeedButton;
    ExcludeBtn: TSpeedButton;
    ExAllBtn: TSpeedButton;
    lblPath2: TLabel;
    DstList: TListBox;
    SrcList: TListBox;
    Label1: TLabel;
    dlbScriptsPath: TDirectoryListBox;
    flbScriptFiles: TFileListBox;
    procedure IncludeBtnClick(Sender: TObject);
    procedure ExcludeBtnClick(Sender: TObject);
    procedure IncAllBtnClick(Sender: TObject);
    procedure ExcAllBtnClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SrcListClick(Sender: TObject);
    procedure DstListClick(Sender: TObject);
  private
    { Private declarations }
    procedure GetWordsInfo(nI_Type : Integer; var I_KeyStrList : TStringList; I_UnitStrList : TStringList; sI_FullPathFileName:String);
    procedure DeleteProjObj;
    procedure InitState;
  public
    { Public declarations }
    procedure MoveSelected(List: TCustomListBox; Items: TStrings);
    procedure SetItem(List: TListBox; Index: Integer);
    function GetFirstSelection(List: TCustomListBox): Integer;
    procedure SetButtons;
  end;

var
  frmTL0020: TfrmTL0020;

implementation

uses dtmTL0020U, Ustru, UCommonU, dtmCommonU, UObjectu;

{$R *.DFM}

procedure TfrmTL0020.IncludeBtnClick(Sender: TObject);
var
  Index: Integer;
begin
  Index := GetFirstSelection(SrcList);
  MoveSelected(SrcList, DstList.Items);
  SetItem(SrcList, Index);
end;

procedure TfrmTL0020.ExcludeBtnClick(Sender: TObject);
var
  Index: Integer;
begin
  Index := GetFirstSelection(DstList);
  MoveSelected(DstList, SrcList.Items);
  SetItem(DstList, Index);
end;

procedure TfrmTL0020.IncAllBtnClick(Sender: TObject);
var
  I: Integer;
begin
  for I := 0 to SrcList.Items.Count - 1 do
    DstList.Items.AddObject(SrcList.Items[I], 
      SrcList.Items.Objects[I]);
  SrcList.Items.Clear;
  SetItem(SrcList, 0);
end;

procedure TfrmTL0020.ExcAllBtnClick(Sender: TObject);
var
  I: Integer;
begin
  for I := 0 to DstList.Items.Count - 1 do
    SrcList.Items.AddObject(DstList.Items[I], DstList.Items.Objects[I]);
  DstList.Items.Clear;
  SetItem(DstList, 0);
end;

procedure TfrmTL0020.MoveSelected(List: TCustomListBox; Items: TStrings);
var
  I: Integer;
begin
  for I := List.Items.Count - 1 downto 0 do
    if List.Selected[I] then
    begin
      Items.AddObject(List.Items[I], List.Items.Objects[I]);
      List.Items.Delete(I);
    end;
end;

procedure TfrmTL0020.SetButtons;
var
  SrcEmpty, DstEmpty: Boolean;
begin
  SrcEmpty := SrcList.Items.Count = 0;
  DstEmpty := DstList.Items.Count = 0;
  IncludeBtn.Enabled := not SrcEmpty;
  IncAllBtn.Enabled := not SrcEmpty;
  ExcludeBtn.Enabled := not DstEmpty;
  ExAllBtn.Enabled := not DstEmpty;
end;

function TfrmTL0020.GetFirstSelection(List: TCustomListBox): Integer;
begin
  for Result := 0 to List.Items.Count - 1 do
    if List.Selected[Result] then Exit;
  Result := LB_ERR;
end;

procedure TfrmTL0020.SetItem(List: TListBox; Index: Integer);
var
  MaxIndex: Integer;
begin
  with List do
  begin
    SetFocus;
    MaxIndex := List.Items.Count - 1;
    if Index = LB_ERR then Index := 0
    else if Index > MaxIndex then Index := MaxIndex;
    Selected[Index] := True;
  end;
  SetButtons;
end;

procedure TfrmTL0020.FormShow(Sender: TObject);
begin
    InitState;
end;

procedure TfrmTL0020.BitBtn2Click(Sender: TObject);
begin
    DeleteProjObj;
    ModalResult := mrOK;
end;

procedure TfrmTL0020.InitState;
var
    L_DataSet : TDataSet;
    sL_Path, sL_Name, sL_ProjName : String;
    L_Obj : TProjInfoObj;
begin
    L_DataSet := dtmCommon.GetProjName;
    with L_DataSet do
    begin
      while not Eof do
      begin
        sL_ProjName := FieldByName('sProjName').AsString;

        sL_Name := ExtractFileName(sL_ProjName);
        sL_Path := ExtractFilePath(sL_ProjName);

        L_Obj := TProjInfoObj.Create;
        L_Obj.sFullPathProjName := sL_ProjName;
        L_Obj.sPath := sL_Path;
        L_Obj.sProjName := sL_Name;
        SrcList.Items.AddObject(sL_Name,L_Obj);
        Next;
      end;
      Close;
    end;

    SetButtons;
end;

procedure TfrmTL0020.BitBtn1Click(Sender: TObject);
var
    ii, nL_StatusPos : Integer;
    sL_ScriptsDir : String;
    sL_ProjName, sL_FullPathUnitName, sL_UnitPath, sL_UnitName : String;
    L_DataSet : TDataSet;
    L_UnitStrList, L_KeyStrList : TStringList;
begin
    case pgcType.ActivePageIndex of
     1:
      begin
       sL_ScriptsDir := dlbScriptsPath.Directory;
       try
        Screen.Cursor := crHourGlass;
        dtmTL0020.EmptyWordsInfo;
        L_UnitStrList := TStringList.Create;
        L_KeyStrList := TStringList.Create;

        nL_StatusPos := 0;
        pgbStatus.Min := nL_StatusPos;
        pgbStatus.Max := flbScriptFiles.Items.Count;

        for ii:=0 to flbScriptFiles.Items.Count -1 do
        begin
          INC(nL_StatusPos);
          pgbStatus.Position := nL_StatusPos;
          sL_UnitName := flbScriptFiles.Items.Strings[ii];
          sL_FullPathUnitName := sL_ScriptsDir + '\' + sL_UnitName;
          L_UnitStrList.Clear;
          GetWordsInfo(pgcType.ActivePageIndex, L_KeyStrList, L_UnitStrList, sL_FullPathUnitName);

        end;
       except
          L_UnitStrList.Free;
          L_KeyStrList.Free;
          Screen.Cursor := crDefault;
          MessageDlg('Get patten failed!!',mtError,[mbOK],0);
          Exit;
       end;
       L_UnitStrList.Free;
       L_KeyStrList.Free;
       Screen.Cursor := crDefault;

      end;
     0:
       begin
          if DstList.Items.Count =0 then
          begin
            MessageDlg('Select project please...', mtWarning,[mbOK],0);
            Exit;
          end;

          try
            Screen.Cursor := crHourGlass;
            dtmTL0020.EmptyWordsInfo;
            L_UnitStrList := TStringList.Create;
            L_KeyStrList := TStringList.Create;

            for ii:=0 to DstList.Items.Count -1 do
            begin

              sL_ProjName := (DstList.Items.Objects[ii] As TProjInfoObj).sFullPathProjName;
              L_DataSet := dtmCommon.GetProjInfo(sL_ProjName);

              nL_StatusPos := 0;
              pgbStatus.Min := nL_StatusPos;
              pgbStatus.Max := L_DataSet.RecordCount;

              with L_DataSet do
              begin
                while not Eof do
                begin
                  INC(nL_StatusPos);
                  pgbStatus.Position := nL_StatusPos;
                  sL_FullPathUnitName := FieldByName('sUnitName').AsString;
                  sL_UnitName := ExtractFileName(sL_FullPathUnitName);
                  sL_UnitPath := ExtractFilePath(sL_FullPathUnitName);
                  L_UnitStrList.Clear;
                  GetWordsInfo(pgcType.ActivePageIndex, L_KeyStrList, L_UnitStrList, sL_FullPathUnitName);

                  Next;
                end;
              end;
            end;
          except
            L_UnitStrList.Free;
            L_KeyStrList.Free;
            Screen.Cursor := crDefault;
            MessageDlg('Get patten failed!!',mtError,[mbOK],0);
            Exit;
          end;
          L_UnitStrList.Free;
          L_KeyStrList.Free;
          Screen.Cursor := crDefault;

        end;
    end;
    MessageDlg('Get patten successfully!!',mtInformation,[mbOK],0);
    pgbStatus.Position := 0;
end;

procedure TfrmTL0020.DeleteProjObj;
var
    ii : Integer;
begin
    for ii:=0 to SrcList.Items.Count -1 do
    begin
      SrcList.Items.Delete(ii);
    end;

    for ii:=0 to DstList.Items.Count -1 do
    begin
      DstList.Items.Delete(ii);
    end;    
end;

procedure TfrmTL0020.SpeedButton1Click(Sender: TObject);
begin
      if MessageDlg('Are you sure to reset all words info?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        dtmCommon.DeleteAllWordsInfo;
end;

procedure TfrmTL0020.SrcListClick(Sender: TObject);
begin
     lblPath1.Caption := (SrcList.Items.Objects[SrcList.itemindex] As TProjInfoObj).sPath;
end;

procedure TfrmTL0020.DstListClick(Sender: TObject);
begin
     lblPath2.Caption := (DstList.Items.Objects[DstList.itemindex] As TProjInfoObj).sPath;
end;

procedure TfrmTL0020.GetWordsInfo(nI_Type : Integer; var I_KeyStrList : TStringList; I_UnitStrList : TStringList; sI_FullPathFileName: String);
var
   L_TmpStrList : TStringList;
   nL_HeadSpaceCount, nL_TailSpaceCount, jj,kk,pp,nL_Ndx : Integer;
   sL_LineSrc,sL_Dest,sL_Body, sL_Head, sL_Tail : String;
   sL_KeyValue : String;
   cL_SEP_CHAR : char;
begin
    case nI_Type of
     0: //VB files
      cL_SEP_CHAR := VB_CHINESE_SEP_CHAR;
     1: //Sctipt files
      cL_SEP_CHAR := SCRIPT_CHINESE_SEP_CHAR;
    end;

    I_UnitStrList.LoadFromFile(sI_FullPathFileName);
    for jj :=0 to I_UnitStrList.Count-1 do
    begin
      sL_LineSrc := I_UnitStrList.Strings[jj];

      L_TmpStrList := TUStr.ParseStrings(Trim(sL_LineSrc),cL_SEP_CHAR, False);
      if L_TmpStrList=nil then
        continue;

      //down, 註解不翻譯
      case nI_Type of
       0: //VB files
        begin
        if sL_LineSrc[1]='''' then
         continue;
        end;
       1: //Script files
        begin
        if copy(sL_LineSrc,1,2)='--' then
         continue;
        end;
      end;
      //up, 註解不翻譯

      if L_TmpStrList.Count>1 then
      begin
        for kk:=1 to L_TmpStrList.Count - 1 do
        begin
          if (kk mod 2) <>0 then//kk是基數者為所要取出之字串
          begin
            sL_Dest := L_TmpStrList.Strings[kk];
            for PP:=0 to Length(sL_Dest) -1 do
            begin
              if Ord(sL_Dest[PP]) > 127 then   //>127者才是中文
              begin
                if sL_Dest[Length(sL_Dest)]=cL_SEP_CHAR then
                  sL_Dest := Copy(sL_Dest,1,Length(sL_Dest)-1);
                TUCommonFun.SepBHT(nL_HeadSpaceCount,nL_TailSpaceCount, sL_Dest, sL_Body, sL_Head, sL_Tail);
//                sL_KeyValue :=sL_Body + TUstr.Space(nL_HeadSpaceCount) + sL_Head + sL_Tail + TUstr.Space(nL_TailSpaceCount) ;
                sL_KeyValue := sL_Dest;
                nL_Ndx := I_KeyStrList.IndexOf(sL_KeyValue);
                if nL_Ndx=-1 then
                begin
                  I_KeyStrList.Add(sL_KeyValue);
                  dtmTL0020.AppendWordsInfo(Length(sL_Dest),nL_HeadSpaceCount,nL_TailSpaceCount, sL_Dest, sL_Body,sL_Head,sL_Tail);
                end;
                break;
              end;
            end;
          end;
        end;
      end;
    end;
end;

end.
