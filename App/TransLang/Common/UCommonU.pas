unit UCommonU;

interface

uses
    SysUtils, comctrls, classes, dbctrls, stdctrls, controls, inifiles, Forms;

const
    INIFILENAME='Ini\TransLang.ini';
    VB_CHINESE_SEP_CHAR='"';
    SCRIPT_CHINESE_SEP_CHAR='''';    

type
  TUCommonFun = class
    public
      class function RWIniFiles(nI_RW:Integer; Section, sI_Ident, sI_Value: string):String;
      class procedure SepBHT(var nI_HeadSpaceCount,nI_TailSpaceCount:Integer; sI_Dest : String; var sI_Body, sI_Head, sI_Tail : String);      
    end;
var
    //down,各語文之目錄區隔文字
   sC_CHINESE_DIR_KEYWORD : String;//中文繁體
   sC_ENG_DIR_KEYWORD : String;//英文
   sC_JAPAN_DIR_KEYWORD : String;//日文
   sC_CHINA_DIR_KEYWORD : String;//中文簡體
    //up,各語文之目錄區隔文字

   sC_UserID, sC_UserName, sC_Password, sC_UserGrp : String;    

implementation

{ TUCommonFun }


{ TUCommonFun }


{ TUCommonFun }


class function TUCommonFun.RWIniFiles(nI_RW: Integer; Section, sI_Ident,
  sI_Value: string):String;
var
   L_IniFile : TIniFile;
   ExeFileName, ExePath : String;
begin
   ExeFileName := Application.ExeName;
   ExePath := ExtractFileDir(ExeFileName);

   L_IniFile := TIniFile.Create(ExePath + '\' + INIFILENAME);
   case nI_RW of
    0://read
      Result := L_IniFile.ReadString(Section,sI_Ident,sI_Value);
    1://write
      L_IniFile.WriteString(Section,sI_Ident,sI_Value);
   end;

   L_IniFile.Free;
end;


class procedure TUCommonFun.SepBHT(var nI_HeadSpaceCount, nI_TailSpaceCount:Integer;sI_Dest: String; var sI_Body, sI_Head,
  sI_Tail: String);
var
    ii : Integer;
begin
    {
     Head僅允許由'&'開頭,且最長3個byte
     Tail最長僅允許3個'.'
    }

    if sI_Dest[1]='&' then
    begin
      if sI_Dest[3]='.' then
        sI_Head := Copy(sI_Dest,1,3)
      else
        sI_Head := Copy(sI_Dest,1,2);
      sI_Body := Copy(sI_Dest,Length(sI_Head)+1,Length(sI_Dest)-Length(sI_Head));
    end
    else
    begin
      sI_Head := '';
      sI_Body := sI_Dest;
    end;

    sI_Tail := '';
    for ii:=Length(sI_Dest) downto Length(SI_Dest)-3 do
    begin
      if sI_Dest[ii]='.' then
        sI_Tail := sI_Tail + '.';
    end;

    nI_HeadSpaceCount := 0;    
    for ii:=1 to Length(sI_Body) do
    begin
      if sI_Body[ii]=' ' then
        INC(nI_HeadSpaceCount)
      else
        break;
    end;

    nI_TailSpaceCount := 0;
    for ii:=Length(sI_Body) downto 1 do
    begin
      if sI_Body[ii]=' ' then
        INC(nI_TailSpaceCount)
      else
        break;
    end;

    sI_Body := Trim(Copy(sI_Body,1,Length(sI_Body)-Length(sI_Tail)));
end;

end.

