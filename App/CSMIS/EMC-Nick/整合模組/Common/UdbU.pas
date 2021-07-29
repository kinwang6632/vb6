unit UdbU;

interface

uses
  Windows, Messages, SysUtils, Classes, Controls, Db, Forms, Dialogs, Math, DBClient,
  Variants,DBGrids, inifiles ;

const
    IS_SYS_DEFAULT_VALUE = true;
    DEFAULT_SECTION_TAIL = 'DEFAULT';
    STRING_FIELD_TYPE = 0;
    INT_FIELD_TYPE = 1;
    FLOAT_FIELD_TYPE = 2;
    DATE_FIELD_TYPE = 3;
    BUTTON_STYLE_AUTO = 91;
    BUTTON_STYLE_ELLIPSIS = 92;
    BUTTON_STYLE_NONE = 93;
    EMPTY_IME_NAME_FLAG = '***';
Type


  TUdb = class
  private
    class function getImeModeStrFlag(I_ImeMode:TImeMode):String;
    class function getImeMode(sI_ImeModeStrFlag:String):TImeMode;
  public
    class procedure assignFieldName(I_CDS : TClientDataSet);
    class function getRecNoCaption(I_SrcDataSet:TDataSet):String;
    class procedure saveDbGridColumnInfo(I_DbGrid:TDbGrid; sI_IniFileName, sI_SectionName:String; bI_IsSysDefauleValue:boolean);
    class procedure getDbGridColumnInfo(I_DbGrid:TDbGrid; sI_IniFileName, sI_SectionName:String; bI_IsSysDefauleValue:boolean);
    class procedure createField(I_CDS:TClientDataset; nI_DataType, sI_FieldSize:Integer; sI_FieldName:String);
    class function getDataType(I_DataSet:TDataSet; sI_FieldName:String):TFieldType;
    class procedure copyDataSetRec(I_SrcDataSet, I_DestDataSet:TDataSet; bI_BaseOnSrcDataSet:Boolean);
  end;


implementation

uses Ustru;


{ TUdb }


{ TUdb }

class procedure TUdb.assignFieldName(I_CDS: TClientDataSet);
var
    ii : Integer;
begin
    for ii:=0 to I_CDS.FieldDefs.Count -1 do
      I_CDS.Fields[ii].Name := I_CDS.FieldDefs[ii].Name;

end;

class procedure TUdb.copyDataSetRec(I_SrcDataSet, I_DestDataSet: TDataSet;
  bI_BaseOnSrcDataSet: Boolean);
var
    nL_RecCount, ii,jj : Integer;
    nL_FieldCount : Integer;
    sL_FieldName : String;
begin
    if I_DestDataSet.State = dsInactive then
    begin
      I_DestDataSet.FieldDefs := (I_SrcDataSet As TClientDataSet).FieldDefs;
      (I_DestDataSet As TClientDataSet).CreateDataSet;
      I_DestDataSet.Active := True;
    end
    else
      (I_DestDataSet As TClientDataSet).emptyDataSet;
      
    with I_DestDataSet do
    begin
      if bI_BaseOnSrcDataSet then
        nL_FieldCount := I_SrcDataSet.Fields.Count
      else
        nL_FieldCount := Fields.Count;
    end;
    I_SrcDataSet.First;
    nL_RecCount := I_SrcDataSet.RecordCount;
    for jj:=0 to nL_RecCount-1 do
    begin
      with I_DestDataSet do
      begin
        Append;
        for ii:=0 to nL_FieldCount-1 do
        begin
          sL_FieldName := Fields[ii].FieldName;

          case Fields[ii].DataType of
           ftFloat :
             Fields[ii].AsFloat := I_SrcDataSet.FieldByName(sL_FieldName).AsFloat;
           ftString, ftMemo:
             Fields[ii].AsString := I_SrcDataSet.FieldByName(sL_FieldName).AsString;
           ftInteger:
            begin
             if I_SrcDataSet.FieldByName(sL_FieldName).Value = NULL then
               Fields[ii].Value := NULL
             else
               Fields[ii].AsInteger := I_SrcDataSet.FieldByName(sL_FieldName).AsInteger;
            end;
           ftBoolean:
            begin
             if I_SrcDataSet.FieldByName(sL_FieldName).Value = NULL then
               Fields[ii].Value := NULL
             else
               Fields[ii].AsBoolean := I_SrcDataSet.FieldByName(sL_FieldName).AsBoolean;
            end;
           ftDateTime:
            begin
             if I_SrcDataSet.FieldByName(sL_FieldName).Value = NULL then
               Fields[ii].Value := NULL
             else
               Fields[ii].AsDateTime := I_SrcDataSet.FieldByName(sL_FieldName).AsDateTime;
            end;
          end;

        end;
        Post;
      end;
      I_SrcDataSet.Next;
    end;
end;

class procedure TUdb.createField(I_CDS: TClientDataset; nI_DataType,
  sI_FieldSize: Integer; sI_FieldName: String);
var
    L_DataType : TFieldType;
    L_FieldDef : TFieldDef;
begin
    if I_CDS.FindField(sI_FieldName)<>nil then Exit;
  case nI_DataType of
    STRING_FIELD_TYPE://String type
      L_DataType := ftString;
    INT_FIELD_TYPE://Integer type
      L_DataType := ftInteger;
    FLOAT_FIELD_TYPE://double type
      L_DataType := ftFloat;
    DATE_FIELD_TYPE://date type
      L_DataType := ftDate;
  end;

  with I_CDS do
  begin
    L_FieldDef := FieldDefs.AddFieldDef;
    with L_FieldDef do
    begin
      DataType := L_DataType;

      Name := sI_FieldName;
      if nI_DataType =0 then//String type
        Size := sI_FieldSize;
    end;
  end;
end;

class function  TUdb.getDataType(I_DataSet: TDataSet;
  sI_FieldName: String): TFieldType;
begin
    result := I_DataSet.FieldByName(sI_FieldName).DataType;
end;


class procedure TUdb.getDbGridColumnInfo(I_DbGrid: TDbGrid;
  sI_IniFileName, sI_SectionName: String; bI_IsSysDefauleValue:boolean);
var
    nL_ButtonStyle : Integer;
    sL_ColumnWidth, sL_ImeModeStrFlag, sL_FieldCaption, sL_Content : String;
    sL_RealSectionName, sL_FullPathFileName, sL_FieldName, sL_ImeName : String;
    ii : Integer;
    L_IniFile : TIniFile;
    L_StrList : TStringList;
    L_TmpButtonStyle : TColumnButtonStyle;
    L_ImeMode : TImeMode;
begin
    if bI_IsSysDefauleValue then
      sL_RealSectionName := sI_SectionName + DEFAULT_SECTION_TAIL
    else
      sL_RealSectionName := sI_SectionName;
      
//    sL_FullPathFileName := ExtractFilePath(Application.ExeName) + sI_IniFileName + '.ini';
    sL_FullPathFileName := sI_IniFileName;
    if not FileExists(sL_FullPathFileName) then Exit;
    L_IniFile := TIniFile.Create(sL_FullPathFileName);
    with I_DbGrid do
    begin
      for ii:=0 to Columns.Count -1 do
      begin
        sL_FieldName := Columns[ii].FieldName;

        sL_Content := L_IniFile.ReadString(sL_RealSectionName,IntToStr(ii),'' );
        if sL_Content<>'' then
        begin
          L_StrList := TUStr.ParseStrings(sL_Content,',');

          sL_FieldName := L_StrList.Strings[0];
          sL_FieldCaption := L_StrList.Strings[1];
          nL_ButtonStyle := StrToInt(L_StrList.Strings[2]);
          sL_ImeModeStrFlag := L_StrList.Strings[3];
          sL_ImeName :=  L_StrList.Strings[4];
          sL_ColumnWidth := L_StrList.Strings[5];
          if sL_ImeName= EMPTY_IME_NAME_FLAG then
            sL_ImeName := '';
          L_StrList.Free;
          case nL_ButtonStyle of
            BUTTON_STYLE_AUTO :
              L_TmpButtonStyle := cbsAuto;
            BUTTON_STYLE_ELLIPSIS :
              L_TmpButtonStyle := cbsEllipsis;
            BUTTON_STYLE_NONE :
              L_TmpButtonStyle := cbsNone;
          end;

          L_ImeMode := getImeMode(sL_ImeModeStrFlag);
          Columns[ii].FieldName := sL_FieldName;
          Columns[ii].Title.Caption := sL_FieldCaption;
          Columns[ii].ButtonStyle := L_TmpButtonStyle;
          Columns[ii].ImeMode := L_ImeMode;
          Columns[ii].ImeName := sL_ImeName;
          Columns[ii].Width := StrToIntDef(sL_ColumnWidth, 30)
        end;
      end;
    end;
    L_IniFile.Free;
end;


class function TUdb.getImeMode(sI_ImeModeStrFlag: String): TImeMode;
var
    L_ImeMode : TImeMode;
begin
    if sI_ImeModeStrFlag= '1' then
      L_ImeMode := imDisable
    else if sI_ImeModeStrFlag= '2' then
      L_ImeMode := imClose
    else if sI_ImeModeStrFlag= '3' then
      L_ImeMode := imOpen
    else if sI_ImeModeStrFlag= '4' then
      L_ImeMode := imDontCare
    else if sI_ImeModeStrFlag= '5' then
      L_ImeMode := imSAlpha
    else if sI_ImeModeStrFlag= '6' then
      L_ImeMode := imAlpha
    else if sI_ImeModeStrFlag= '7' then
      L_ImeMode := imHira
    else if sI_ImeModeStrFlag= '8' then
      L_ImeMode := imSKata
    else if sI_ImeModeStrFlag= '9' then
      L_ImeMode := imKata
    else if sI_ImeModeStrFlag= '10' then
      L_ImeMode := imChinese
    else if sI_ImeModeStrFlag= '11' then
      L_ImeMode := imSHanguel
    else if sI_ImeModeStrFlag= '12' then
      L_ImeMode := imHanguel;
    result := L_ImeMode;
end;

class function TUdb.getImeModeStrFlag(I_ImeMode: TImeMode): String;
var
    sL_Result : String;
begin
    case I_ImeMode of
      imDisable :
        sL_Result := '1';
      imClose :
        sL_Result := '2';
      imOpen :
        sL_Result := '3';
      imDontCare :
        sL_Result := '4';
      imSAlpha :
        sL_Result := '5';
      imAlpha :
        sL_Result := '6';
      imHira :
        sL_Result := '7';
      imSKata  :
        sL_Result := '8';
      imKata :
        sL_Result := '9';
      imChinese	:
        sL_Result := '10';
      imSHanguel :
        sL_Result := '11';
      imHanguel	:
        sL_Result := '12';
    end;
    result := sL_Result;
end;

class function TUdb.getRecNoCaption(I_SrcDataSet: TDataSet): String;
begin
    result := IntToStr(I_SrcDataSet.RecNo) + '/' + IntToStr(I_SrcDataSet.RecordCount);
end;

class procedure TUdb.saveDbGridColumnInfo(I_DbGrid: TDbGrid;
  sI_IniFileName, sI_SectionName: String; bI_IsSysDefauleValue:boolean);
var
    bL_HasDefauleValueSection : boolean;
    L_SectionStrList : TStringList;
    sL_ColumnWidth, sL_ImeMode, sL_ButtonStyle : String;
    sL_FullPathFileName, sL_FieldCaption, sL_FieldName : String;
    ii : Integer;
    sL_RealSectionName, sL_ImeName : String;
    L_IniFile : TIniFile;
begin

//    sL_FullPathFileName := ExtractFilePath(Application.ExeName) + sI_IniFileName + '.ini';
    sL_FullPathFileName := sI_IniFileName;
    
    L_IniFile := TIniFile.Create(sL_FullPathFileName);

    if bI_IsSysDefauleValue then
    begin
      sL_RealSectionName := sI_SectionName + DEFAULT_SECTION_TAIL;
      L_SectionStrList := TStringList.Create;
      L_IniFile.ReadSection(sL_RealSectionName,L_SectionStrList);
      if L_SectionStrList.Count>0 then
        bL_HasDefauleValueSection := true
      else
        bL_HasDefauleValueSection := false;
      L_SectionStrList.Free;
      L_SectionStrList := nil;
    end
    else
      sL_RealSectionName := sI_SectionName;


    with I_DbGrid do
    begin
      for ii:=0 to Columns.Count -1 do
      begin
        sL_FieldName := Columns[ii].FieldName;
        sL_FieldCaption := Columns[ii].Title.Caption;
        //down,處理 button style..
        if Columns[ii].ButtonStyle = cbsAuto then
          sL_ButtonStyle := IntToStr(BUTTON_STYLE_AUTO)
        else if Columns[ii].ButtonStyle = cbsEllipsis then
          sL_ButtonStyle := IntToStr(BUTTON_STYLE_ELLIPSIS)
        else if Columns[ii].ButtonStyle = cbsNone then
          sL_ButtonStyle := IntToStr(BUTTON_STYLE_None);
        //down,處理 button style..

        //down,處理 ime mode..
        sL_ImeMode := getImeModeStrFlag(Columns[ii].ImeMode);
        //up,處理 ime mode..

        sL_ImeName := Columns[ii].ImeName;
        if sL_ImeName='' then
          sL_ImeName := EMPTY_IME_NAME_FLAG;


        sL_ColumnWidth := IntToStr(Columns[ii].Width);
        if ((bI_IsSysDefauleValue) and (not bL_HasDefauleValueSection)) or
           (not (bI_IsSysDefauleValue)) then
          L_IniFile.WriteString(sL_RealSectionName,IntToStr(ii),sL_FieldName+','+sL_FieldCaption+','+sL_ButtonStyle+','+ sL_ImeMode + ',' + sL_ImeName + ',' + sL_ColumnWidth);
      end;
    end;
    L_IniFile.Free;
end;

end.


