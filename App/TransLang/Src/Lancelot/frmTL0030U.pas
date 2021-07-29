unit frmTL0030U;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Forms, Dialogs, Controls, StdCtrls,
  Buttons, db, ExtCtrls, FileCtrl, ComCtrls;

type

  TfrmTL0030 = class(TForm)
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    rgpLang: TRadioGroup;
    pgbStatus: TProgressBar;
    Panel1: TPanel;
    pgcType: TPageControl;
    TabSheet1: TTabSheet;
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
    TabSheet2: TTabSheet;
    Label1: TLabel;
    flbScriptFiles: TFileListBox;
    dlbScriptsPath: TDirectoryListBox;
    procedure IncludeBtnClick(Sender: TObject);
    procedure ExcludeBtnClick(Sender: TObject);
    procedure IncAllBtnClick(Sender: TObject);
    procedure ExcAllBtnClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure SrcListClick(Sender: TObject);
    procedure DstListClick(Sender: TObject);
  private
    { Private declarations }
    procedure TransWords(nI_Type : Integer; I_OldFileStrList:TStringList; var I_NewFileStrList : TStringList);
    function GetNewFileName(sI_OldFileName:String):String;
    procedure CreateDirectory(sI_FullPath:String);
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
  frmTL0030: TfrmTL0030;

implementation

uses dtmTL0030U, dtmCommonU, UObjectu, UCommonU, Ustru;


{$R *.DFM}

procedure TfrmTL0030.IncludeBtnClick(Sender: TObject);
var
  Index: Integer;
begin
  Index := GetFirstSelection(SrcList);
  MoveSelected(SrcList, DstList.Items);
  SetItem(SrcList, Index);
end;

procedure TfrmTL0030.ExcludeBtnClick(Sender: TObject);
var
  Index: Integer;
begin
  Index := GetFirstSelection(DstList);
  MoveSelected(DstList, SrcList.Items);
  SetItem(DstList, Index);
end;

procedure TfrmTL0030.IncAllBtnClick(Sender: TObject);
var
  I: Integer;
begin
  for I := 0 to SrcList.Items.Count - 1 do
    DstList.Items.AddObject(SrcList.Items[I], 
      SrcList.Items.Objects[I]);
  SrcList.Items.Clear;
  SetItem(SrcList, 0);
end;

procedure TfrmTL0030.ExcAllBtnClick(Sender: TObject);
var
  I: Integer;
begin
  for I := 0 to DstList.Items.Count - 1 do
    SrcList.Items.AddObject(DstList.Items[I], DstList.Items.Objects[I]);
  DstList.Items.Clear;
  SetItem(DstList, 0);
end;

procedure TfrmTL0030.MoveSelected(List: TCustomListBox; Items: TStrings);
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

procedure TfrmTL0030.SetButtons;
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

function TfrmTL0030.GetFirstSelection(List: TCustomListBox): Integer;
begin
  for Result := 0 to List.Items.Count - 1 do
    if List.Selected[Result] then Exit;
  Result := LB_ERR;
end;

procedure TfrmTL0030.SetItem(List: TListBox; Index: Integer);
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

procedure TfrmTL0030.FormShow(Sender: TObject);
begin
    InitState;
end;

procedure TfrmTL0030.BitBtn2Click(Sender: TObject);
begin
    DeleteProjObj;
    ModalResult := mrOK;
end;

procedure TfrmTL0030.InitState;
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

    sC_CHINESE_DIR_KEYWORD := TUCommonFun.RWIniFiles(0,'SETTING','CHINESE_DIR_KEYWORD','');
    sC_ENG_DIR_KEYWORD := TUCommonFun.RWIniFiles(0,'SETTING','ENG_DIR_KEYWORD','');
    sC_JAPAN_DIR_KEYWORD := TUCommonFun.RWIniFiles(0,'SETTING','JAPAN_DIR_KEYWORD','');
    sC_CHINA_DIR_KEYWORD := TUCommonFun.RWIniFiles(0,'SETTING','CHINA_DIR_KEYWORD','');
end;

procedure TfrmTL0030.BitBtn1Click(Sender: TObject);
var
    cL_DestFileName: array[0..79] of Char;
    cL_SrcFileName: array[0..79] of Char;
    sL_NewFileName : String;
    ii, nL_Pos, nL_StatusPos : Integer;
    sL_ProjName, sL_FullPathUnitName,sL_ScriptsDir : String;
    sL_UnitName, sL_NewFileDir : String;
    L_DataSet : TDataSet;
    L_OldFileStrList, L_NewFileStrList : TStringList;
begin
    case pgcType.ActivePageIndex of
     1:
      begin
       Screen.Cursor := crHourGlass;
       L_OldFileStrList := TStringList.Create;
       L_NewFileStrList := TStringList.Create;
      
       sL_ScriptsDir := dlbScriptsPath.Directory;
       sL_NewFileDir := sL_ScriptsDir + '\Dest';
       if not DirectoryExists(sL_NewFileDir) then
         CreateDir(sL_NewFileDir);
       try
        Screen.Cursor := crHourGlass;

        for ii:=0 to flbScriptFiles.Items.Count -1 do
        begin
          sL_UnitName := flbScriptFiles.Items.Strings[ii];
          sL_FullPathUnitName := sL_ScriptsDir + '\' + sL_UnitName;
          sL_NewFileName := sL_NewFileDir  + '\' +sL_UnitName;
          if not FileExists(sL_NewFileName) then
          begin
            StrPCopy(cL_SrcFileName,sL_FullPathUnitName);
            StrPCopy(cL_DestFileName,sL_NewFileName);
            CopyFile(cL_SrcFileName,cL_DestFileName , FALSE);
          end;
          L_OldFileStrList.Clear;
          L_NewFileStrList.Clear;
          L_OldFileStrList.LoadFromFile(sL_FullPathUnitName);          
          TransWords(pgcType.ActivePageIndex, L_OldFileStrList,L_NewFileStrList);
          L_NewFileStrList.SaveToFile(sL_NewFileName);
        end;
       except
          Screen.Cursor := crDefault;
          MessageDlg('Trans words failed!!',mtError,[mbOK],0);
          Exit;
       end;
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
      L_OldFileStrList := TStringList.Create;
      L_NewFileStrList := TStringList.Create;

      for ii:=0 to DstList.Items.Count -1 do
      begin
        sL_ProjName := (DstList.Items.Objects[ii] As TProjInfoObj).sFullPathProjName;
        sL_NewFileName := GetNewFileName(sL_ProjName);
        CreateDirectory(sL_NewFileName);


        if not FileExists(sL_NewFileName) then
        begin
          StrPCopy(cL_SrcFileName,sL_ProjName);
          StrPCopy(cL_DestFileName,sL_NewFileName);
          CopyFile(cL_SrcFileName,cL_DestFileName , FALSE);
        end;
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
          

            nL_Pos := AnsiPos(sC_CHINESE_DIR_KEYWORD, sL_FullPathUnitName);
            if nL_Pos<>0 then
            begin
              L_OldFileStrList.Clear;
              L_NewFileStrList.Clear;
              L_OldFileStrList.LoadFromFile(sL_FullPathUnitName);
              sL_NewFileName := GetNewFileName(sL_FullPathUnitName);
              CreateDirectory(sL_NewFileName);

              //down, copy .frx file
              if UpperCase(copy(sL_FullPathUnitName,Length(sL_FullPathUnitName)-2,3))= 'FRM' then
              begin
                StrPCopy(cL_SrcFileName,copy(sL_FullPathUnitName,1,Length(sL_FullPathUnitName)-3)+'FRX');
                StrPCopy(cL_DestFileName,copy(sL_NewFileName,1,Length(sL_NewFileName)-3)+'FRX');
                CopyFile(cL_SrcFileName,cL_DestFileName , FALSE);
              end;
              //up, copy .frx file

              StrPCopy(cL_SrcFileName,sL_FullPathUnitName);
              StrPCopy(cL_DestFileName,sL_NewFileName);

              if CopyFile(cL_SrcFileName,cL_DestFileName , FALSE) then
              begin
                TransWords(pgcType.ActivePageIndex, L_OldFileStrList,L_NewFileStrList);

                L_NewFileStrList.SaveToFile(sL_NewFileName);
              end;
            end;
            Next;
          end;
        end;
      end;
      except
        L_OldFileStrList.Free;
        L_NewFileStrList.Free;
        Screen.Cursor := crDefault;
        MessageDlg('Trans words failed!!',mtError,[mbOK],0);
        Exit;
      end;
     end;
    end;

    L_OldFileStrList.Free;
    L_NewFileStrList.Free;
    Screen.Cursor := crDefault;
    MessageDlg('Trans words successfully!!',mtInformation,[mbOK],0);
    pgbStatus.Position := 0;
end;

procedure TfrmTL0030.DeleteProjObj;
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

function TfrmTL0030.GetNewFileName(sI_OldFileName: String): String;
var
    sL_LeftPath, sL_RightPath : String;
    nL_Pos : Integer;
begin
    nL_Pos := AnsiPos(sC_CHINESE_DIR_KEYWORD, sI_OldFileName);
    sL_LeftPath := Copy(sI_OldFileName,1,nL_Pos-1);
    sL_RightPath := Copy(sI_OldFileName,nL_Pos+Length(sC_CHINESE_DIR_KEYWORD), Length(sI_OldFileName)-nL_Pos-Length(sC_CHINESE_DIR_KEYWORD)+1);
    case rgpLang.ItemIndex of
     0://Eng
      Result := sL_LeftPath + sC_ENG_DIR_KEYWORD + sL_RightPath;
     1://China
      Result := sL_LeftPath + sC_CHINA_DIR_KEYWORD + sL_RightPath;
     2://Japan
      Result := sL_LeftPath + sC_JAPAN_DIR_KEYWORD + sL_RightPath;
    end;

end;

procedure TfrmTL0030.CreateDirectory(sI_FullPath: String);
var
    L_TmpStrList : TStringList;
    ii,jj : Integer;
    sL_TmpDir : String;
begin
    L_TmpStrList := TUStr.ParseStrings(sI_FullPath,'\',True);
    for ii:=0 to L_TmpStrList.Count -2 do
    begin
      sL_TmpDir := '';
      for jj:=0 to ii do
      begin
        sL_TmpDir := sL_TmpDir + L_TmpStrList.Strings[jj] + '\';
      end;
      if not DirectoryExists(sL_TmpDir) then
        CreateDir(sL_TmpDir);
    end;

end;



procedure TfrmTL0030.SrcListClick(Sender: TObject);
begin
     lblPath1.Caption := (SrcList.Items.Objects[SrcList.itemindex] As TProjInfoObj).sPath;
end;

procedure TfrmTL0030.DstListClick(Sender: TObject);
begin
     lblPath2.Caption := (DstList.Items.Objects[DstList.itemindex] As TProjInfoObj).sPath;
end;

procedure TfrmTL0030.TransWords(nI_Type : Integer; I_OldFileStrList: TStringList;var I_NewFileStrList : TStringList);
var
   kk,jj,pp : Integer;
   nL_HeadSpaceCount,nL_TailSpaceCount : Integer;
   sL_Body, sL_Head, sL_Tail : String;
   sL_LineSrc, sL_NewStr, sL_Dest : String;
   L_TmpStrList : TStringList;
   bL_TransLang, bL_HaveMultiByte, bL_MetCommon, bL_LastByteIs_CHINESE_SEP_CHAR : Boolean;
   cL_SEP_CHAR : char;   
begin
    case nI_Type of
     0: //VB files
      cL_SEP_CHAR := VB_CHINESE_SEP_CHAR;
     1: //Sctipt files
      cL_SEP_CHAR := SCRIPT_CHINESE_SEP_CHAR;
    end;

                for jj:=0 to I_OldFileStrList.Count -1 do
                begin
                  sL_LineSrc := I_OldFileStrList.Strings[jj];

                  L_TmpStrList := TUStr.ParseStrings(TrimRight(sL_LineSrc),cL_SEP_CHAR, False);
                  if L_TmpStrList=nil then
                  begin
                    I_NewFileStrList.Add(' ');
                    continue;
                  end;
                  bL_HaveMultiByte := False;
                  bL_MetCommon := False;
                  for pp:=0 to Length(sL_LineSrc) -1 do
                  begin
                    if (not bL_MetCommon)and(Ord(sL_LineSrc[PP]) > 127) then   //>127者才是中文
                    begin
                      bL_HaveMultiByte := True;
                      break;
                    end
                    else if ((nI_Type=0) and (sL_LineSrc[PP]='''')) or ((nI_Type=1) and (copy(sL_LineSrc,1,2)='--')) then
                    begin
                      bL_MetCommon := True;
                      break;
                    end;
                  end;

                  if (not bL_HaveMultiByte) or (bL_MetCommon)then
                  begin
                    I_NewFileStrList.Add(sL_LineSrc);
                    continue;
                  end;


                  if sL_LineSrc[Length(sL_LineSrc)]=cL_SEP_CHAR then
                    bL_LastByteIs_CHINESE_SEP_CHAR := True
                  else
                    bL_LastByteIs_CHINESE_SEP_CHAR := False;

                  if L_TmpStrList.Count>0 then
                  begin
                  //***
                    sL_NewStr := '';
                    for kk:=0 to L_TmpStrList.Count - 1 do
                    begin
                      sL_Dest := L_TmpStrList.Strings[kk];
                      if (Length(Trim(sL_Dest))>0 ) and (TrimRight(sL_Dest)[Length(sL_Dest)]=cL_SEP_CHAR) then
                        sL_Dest :=  copy(sL_Dest,1,Length(sL_Dest)-1);

                      if (kk mod 2) <>0 then//kk是基數者為所要取出之字串==>為雙引號裡面的字
                      begin
                        if Trim(sL_Dest)='' then
                            bL_TransLang := True
                        else
                        begin
                          for pp:=0 to Length(sL_Dest) -1 do
                          begin
                            if Ord(sL_Dest[PP]) > 127 then   //>127者才是中文
                            begin
                              bL_TransLang := True;
                              if sL_Dest[Length(sL_Dest)]=cL_SEP_CHAR then
                                sL_Dest := Copy(sL_Dest,1,Length(sL_Dest)-1);
                              TUCommonFun.SepBHT(nL_HeadSpaceCount,nL_TailSpaceCount, sL_Dest, sL_Body, sL_Head, sL_Tail);
                              break;
                            end
                            else
                              bL_TransLang := False;
                          end;
                        end;
                        if bL_TransLang then
                        begin
                          if (Trim(sL_Dest)<>'') then
                            sL_NewStr := sL_NewStr + cL_SEP_CHAR + dtmTL0030.TransLanguage(rgpLang.ItemIndex, sL_Dest, sL_Body, sL_Head, sL_Tail) + cL_SEP_CHAR
                          else
                            sL_NewStr := sL_NewStr + cL_SEP_CHAR + sL_Dest + cL_SEP_CHAR
                        end
                        else
                        begin

                          if (bL_LastByteIs_CHINESE_SEP_CHAR and (Trim(sL_Dest)<>'')) or ((kk<>L_TmpStrList.Count - 1) and (Length(Trim(sL_Dest))>0)) then
                          begin
                            if TrimLeft(sL_Dest)[1]<>cL_SEP_CHAR then
                              sL_Dest := cL_SEP_CHAR + sL_Dest;
                            if TrimRight(sL_Dest)[Length(sL_Dest)]<>cL_SEP_CHAR then
                              sL_Dest :=  sL_Dest + cL_SEP_CHAR ;
                          end;
                          sL_NewStr := sL_NewStr + sL_Dest;
                        end;
                      end
                      else
                      begin
                        sL_NewStr := sL_NewStr + sL_Dest;
                      end;
                    end;
                    I_NewFileStrList.Add(sL_NewStr);
                  //***
                  end
                  else
                    I_NewFileStrList.Add(' ');
                end;
end;

end.
