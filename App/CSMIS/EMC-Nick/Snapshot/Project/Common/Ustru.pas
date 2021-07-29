unit Ustru;

interface

uses
  Windows, Messages, SysUtils, Classes, Controls, Variants ;
const
    CHINESEZERO='零';
    AMT_TAIL = '.00';
    LEFT_MINUS_FLAG = '(';
    RIGHT_MINUS_FLAG = ')';
Type
  TUstr = class
  private
    class function TransLessTwoNumToEng(nI_Num:Integer):String;
  public
    class function replaceStr(sI_SourceStr, sI_OldStr, sI_NewStr:String):String;
    class procedure saveStr2File(sI_Str, sI_FullPathFileName : String);
    class function getOracleSQLDateStr(dI_Date:TDate):String;
    class function getOracleSQLDateTimeStr(dI_Date:TDateTime):String;    
    class function SepString(sI_SrcString, sI_SepFlag:String):String;
    class function TrimString(sI_SrcString, sI_NoneuseFlag:String):String;    
    class function ReverseStr(sI_SrcString:String):String;
    class function GetChineseWord(nI_Num:Integer):String;
    class function ConbineChineseUnit(sI_ChineseNum:String; nI_Pos:Integer):String;
    class function TransNumberToEng(nI_Num:Integer):String;
    class function TransNumberToChinese(SrcString : String): String;
    class function GetPureStr(SrcString, GarbageChar : String) : String;
    class function PadStr(aStr:string; iLen:Integer; cPad:char; align:TAlign):string ;
    class function Space(cnt:Integer):string ;
    class function Num2Str(alValue:Longint; iLen:Integer):string ; overload ;
    class function Num2Str(asValue:string; iLen:Integer):string ; overload ;
    class function ParseStrings(Strs:string; Sep:Char) : TStringList; overload ;
    class function ParseStrings(Strs, sI_Sep:string) : TStringList; overload ;
    class function ParseStrings(Strs:string) : TStringList; overload;
    class procedure ParseStrings(Strs:string; Sep:Char; RStrs: TStrings); overload ;
    class function CommaNumber(Num: string): string;
    class function getPureNumString(sI_Value:String):String;//傳入 '(123,456.00)' 傳回 '-123456.00'
    
    class function ChangeMinusTag(fI_Num: double): string;overload;//將 -123 轉為 '(123.00)'
    class function ChangeMinusTag(sI_Num: String): string;overload;//將 '-123' 轉為 '(123.00)'
    class function ChangeMinusTag(fI_Num: double; bI_IncludeTail:boolean): string;overload;//將 -123 轉為 '(123.00)'
  end;


implementation

class function TUstr.PadStr(aStr:string; iLen:Integer; cPad:char; align:TAlign):string ;
var
  ii : Integer ;
  vStr : string ;
begin
  for ii:=Length(aStr) to iLen-1 do
    vStr := vStr + cPad ;
  if align = alLeft then
    vStr := aStr + vStr
  else
    vStr := vStr + aStr ;
  Result := vStr ;
end ;

class function TUstr.Space(cnt:Integer):string ;
var
  ii : Integer ;
  vStr : string ;
begin
  for ii:=0 to cnt-1 do
    vStr := vStr + ' ' ;
  Result := vStr ;
end ;

class function TUstr.Num2Str(alValue:Longint; iLen:Integer):string ;
var
  vStr : string ;
begin
  vStr := IntToStr(alValue) ;
  Result := PadStr(vStr, iLen, '0', alRight) ;
end ;

class function TUstr.Num2Str(asValue:string; iLen:Integer):string ;
begin
  Result := PadStr(asValue, iLen, '0', alRight) ;
end ;

class procedure TUStr.ParseStrings(Strs:string; Sep:Char; RStrs: TStrings);
var
  ii, jj : Integer ;
begin
  RStrs.Clear;
  if Strs = '' then Exit;
  jj := 1 ;
  for ii:= 1 to Length(Strs) do
  begin
    if Strs[ii] = Sep then
    begin
      RStrs.Add(Trim(Copy(Strs, jj, ii-jj))) ;
      jj := ii + 1 ;
    end ;
  end ;
  RStrs.Add(Trim(Copy(Strs, jj, Length(Strs)-jj+1))) ;
end;

class function TUstr.ParseStrings(Strs:string; Sep:Char) : TStringList ;
var
  ii, jj : Integer ;
  slst : TStringList ;
begin
  if Strs = '' then
  begin
    Result := nil ;
    Exit ;
  end ;
  slst := TStringList.Create ;
  jj := 1 ;
  for ii:= 1 to Length(Strs) do
  begin
    if (Strs[ii] = Sep) and ((ii+1<=Length(Strs)) and (Strs[ii+1] <> Sep))then
    begin
      slst.Add(Trim(Copy(Strs, jj, ii-jj))) ;
      jj := ii + 1 ;
    end ;
  end ;
  //if jj <> Length(Strs) + 1 then
    slst.Add(Trim(Copy(Strs, jj, Length(Strs)-jj+1))) ;
  Result := slst ;
end ;

class function TUstr.ParseStrings(Strs:string) : TStringList ;
var
  ii, fidx : Integer ;
  slst : TStringList ;
begin
  if Strs = '' then
  begin
    Result := nil ;
    Exit ;
  end ;
  slst := nil;
  fidx := 0;
  for ii:= 1 to Length(Strs) do
  begin
    if Strs[ii] <> ' ' then
      if fidx = 0 then fidx := ii;
    if Strs[ii] = ' ' then
      if fidx <> 0 then
      begin
        if not Assigned(slst) then slst := TStringList.Create ;
        slst.Add(Copy(Strs, fidx, ii-fidx)) ;
        fidx := 0;
      end;
  end;
  if fidx <> 0 then
    slst.Add(Copy(Strs, fidx, Length(Strs)-fidx+1)) ;
  Result := slst ;
end ;

class function TUstr.CommaNumber(Num: string): string;
var
  idx, len: Integer;
  decimal, rstr, sL_Flag, sL_TmpResult: string;
  strs: TStrings;
begin
  if Num='' then
  begin
    Result := '';
    Exit;
  end;
  sL_Flag := '';
  if StrToIntDef(Num[1], -1)=-1 then//去掉正負號
  begin
    sL_Flag := Num[1];
    Num := Copy(Num, 2, Length(Num)-1);
  end;
  strs := ParseStrings(Num, '.');
  Num := strs[0];
  if strs.Count > 1 then decimal := strs[1];
  strs.Free;
  if Length(Num) <= (3+Length(sL_Flag)) then
  begin
    if decimal <> '' then
      Result := sL_Flag + Num + '.' + decimal
    else Result := sL_Flag + Num;
    Exit;
  end;

  len := Length(Num) mod 3;
  idx := 1;
  while True do
  begin
    rstr := rstr + Copy(Num, idx, len);
    idx := idx + len;
    if idx = Length(Num)+1 then Break;
    rstr := rstr + ',';
    len := 3;
  end;
  if rstr[1]=',' then
    Delete(rstr,1,1);
  if decimal <> '' then
    sL_TmpResult := rstr + '.' + decimal
  else sL_TmpResult := rstr;

  Result := sL_Flag + sL_TmpResult;
end;

class function TUstr.GetPureStr(SrcString, GarbageChar: String): String;
var
   nStrLength, ii : Integer;
   sTmpResult : String;
begin
     Result := '';
     sTmpResult := '';
     nStrLength := Length(SrcString);
     for ii:=1 to nStrLength do
     begin
       if SrcString[ii]<>GarbageChar then
         sTmpResult := sTmpResult + SrcString[ii];
     end;
     Result := sTmpResult;
end;

class function TUstr.TransNumberToChinese(SrcString: String): String;
var
   sTmp, sTmp_Result : String;
   ii, nL_Pos : Integer;
   sTmp_StringList:TStringList;
begin
   if StrToInt(SrcString)<0 then
   begin
     Result := SrcString;
     Exit;
   end;

   sTmp_StringList:=TStringList.Create;
   for ii:=1 to  Length(SrcString) do
   begin
     sTmp_StringList.Add(GetChineseWord(StrToInt(SrcString[ii])));
   end;

   sTmp := '';
   for ii:=0 to sTmp_StringList.Count -1 do
   begin
     nL_Pos := Length(SrcString)-ii;
     if sTmp_StringList.Strings[ii] <> CHINESEZERO then
     begin
       sTmp_Result := ConbineChineseUnit(sTmp_StringList.Strings[ii],nL_Pos);
       sTmp := sTmp + sTmp_Result;
     end
     else
     begin
       if nL_Pos=5 then
       begin
         sTmp_Result := '萬';
         sTmp := sTmp + sTmp_Result;
       end;
     end;

   end;
   sTmp_StringList.Free;   
   Result := sTmp ;



end;

class function TUstr.TransLessTwoNumToEng(nI_Num: Integer): String;
begin
      case nI_Num of
       0:
         Result := 'ZERO';      
       1:
         Result := 'ONE';
       2:
         Result := 'TWO';
       3:
         Result := 'THREE';
       4:
         Result := 'FOUR';
       5:
         Result := 'FIVE';
       6:
         Result := 'SIX';
       7:
         Result := 'SEVEN';
       8:
         Result := 'EIGHT';
       9:
         Result := 'NINE';
       10:
         Result := 'TEN';
       11:
         Result := 'ELEVEN';
       12:
         Result := 'TWELVE';
       13:
         Result := 'THIRTEEN';
       14:
         Result := 'FOURTEEN';
       15:
         Result := 'FIFTEEN';
       16:
         Result := 'SIXTEEN';
       17:
         Result := 'SEVENTEEN';
       18:
         Result := 'EIGHTEEN';
       19:
         Result := 'NINETEEN';
       20:
         Result := 'TWENTY';
       21:
         Result := 'TWENTY ONE';
       22:
         Result := 'TWENTY TWO';
       23:
         Result := 'TWENTY THREE';
       24:
         Result := 'TWENTY FOUR';
       25:
         Result := 'TWENTY FIVE';
       26:
         Result := 'TWENTY SIX';
       27:
         Result := 'TWENTY SEVEN';
       28:
         Result := 'TWENTY EIGHT';
       29:
         Result := 'TWENTY NINE';
       30:
         Result := 'THIRTY';
       31:
         Result := 'THIRTY ONE';
       32:
         Result := 'THIRTY TWO';
       33:
         Result := 'THIRTY THREE';
       34:
         Result := 'THIRTY FOUR';
       35:
         Result := 'THIRTY FIVE';
       36:
         Result := 'THIRTY SIX';
       37:
         Result := 'THIRTY SEVEN';
       38:
         Result := 'THIRTY EIGHT';
       39:
         Result := 'THIRTY NINE';
       40:
         Result := 'FORTY';
       41:
         Result := 'FORTY ONE';
       42:
         Result := 'FORTY TWO';
       43:
         Result := 'FORTY THREE';
       44:
         Result := 'FORTY FOUR';
       45:
         Result := 'FORTY FIVE';
       46:
         Result := 'FORTY SIX';
       47:
         Result := 'FORTY SEVEN';
       48:
         Result := 'FORTY EIGHT';
       49:
         Result := 'FORTY NINE';
       50:
         Result := 'FIFTY';
       51:
         Result := 'FIFTY ONE';
       52:
         Result := 'FIFTY TWO';
       53:
         Result := 'FIFTY THREE';
       54:
         Result := 'FIFTY FOUR';
       55:
         Result := 'FIFTY FIVE';
       56:
         Result := 'FIFTY SIX';
       57:
         Result := 'FIFTY SEVEN';
       58:
         Result := 'FIFTY EIGHT';
       59:
         Result := 'FIFTY NINE';
       60:
         Result := 'SIXTY';
       61:
         Result := 'SIXTY ONE';
       62:
         Result := 'SIXTY TWO';
       63:
         Result := 'SIXTY THREE';
       64:
         Result := 'SIXTY FOUR';
       65:
         Result := 'SIXTY FIVE';
       66:
         Result := 'SIXTY SIX';
       67:
         Result := 'SIXTY SEVEN';
       68:
         Result := 'SIXTY EIGHT';
       69:
         Result := 'SIXTY NINE';
       70:
         Result := 'SEVENTY';
       71:
         Result := 'SEVENTY ONE';
       72:
         Result := 'SEVENTY TWO';
       73:
         Result := 'SEVENTY THREE';
       74:
         Result := 'SEVENTY FOUR';
       75:
         Result := 'SEVENTY FIVE';
       76:
         Result := 'SEVENTY SIX';
       77:
         Result := 'SEVENTY SEVEN';
       78:
         Result := 'SEVENTY EIGHT';
       79:
         Result := 'SEVENTY NINE';
       80:
         Result := 'EIGHTY';
       81:
         Result := 'EIGHTY ONE';
       82:
         Result := 'EIGHTY TWO';
       83:
         Result := 'EIGHTY THREE';
       84:
         Result := 'EIGHTY FOUR';
       85:
         Result := 'EIGHTY FIVE';
       86:
         Result := 'EIGHTY SIX';
       87:
         Result := 'EIGHTY SEVEN';
       88:
         Result := 'EIGHTY EIGHT';
       89:
         Result := 'EIGHTY NINE';
       90:
         Result := 'NINETY';
       91:
         Result := 'NINETY ONE';
       92:
         Result := 'NINETY TWO';
       93:
         Result := 'NINETY THREE';
       94:
         Result := 'NINETY FOUR';
       95:
         Result := 'NINETY FIVE';
       96:
         Result := 'NINETY SIX';
       97:
         Result := 'NINETY SEVEN';
       98:
         Result := 'NINETY EIGHT';
       99:
         Result := 'NINETY NINE';
       else
         Result := 'UnKnown';
      end;

end;

class function TUstr.TransNumberToEng(nI_Num: Integer): String;
var
    sL_Value, sL_SrcString, sL_Thousand : String;
    sL_P5,sL_P4, sL_P3, sL_P12 : String;
begin
    sL_SrcString := IntToStr(nI_Num);
    sL_P12 := '';
    sL_P3 := '';
    sL_P4 := '';
    sL_P5 := '';
    
    sL_P12 := Copy(sL_SrcString,length(sL_SrcString)-1,2);//個與十位數
    if (length(sL_SrcString)-2)>0 then
      sL_P3 := Copy(sL_SrcString,length(sL_SrcString)-2,1);//百位數

    if (length(sL_SrcString)-3)>0 then
      sL_P4 := Copy(sL_SrcString,length(sL_SrcString)-3,1);//千位數

    if (length(sL_SrcString)-4)>0 then
      sL_P5 := Copy(sL_SrcString,length(sL_SrcString)-4,1);//萬位數

    if sL_P12<>'' then
      sL_P12 := TransLessTwoNumToEng(StrToInt(sL_P12));

    if (sL_P3<>'') and (sL_P3<>'0') then
      sL_P3 := TransLessTwoNumToEng(StrToInt(sL_P3));

    if (sL_P5<>'') and (sL_P5<>'0') then
    begin
      sL_P5 := TransLessTwoNumToEng(StrToInt(sL_P5+sL_P4));
      sL_P4 := '';
    end
    else
    begin
      sL_P5:='';
      if (sL_P4<>'') and (sL_P4<>'0')  then
        sL_P4 := TransLessTwoNumToEng(StrToInt(sL_P4));
    end;


    if sL_P4<>'' then
      sL_Thousand := sL_P4 + ' ' + 'THOUSAND';
    if sL_P5<>'' then
      sL_Thousand := sL_P5 + ' ' + 'THOUSAND';

    if sL_P3<>'' then
      sL_P3 := sL_P3 + ' ' + 'HUNDRED';

    Result := sL_Thousand + ' ' + sL_P3 + ' ' + sL_P12;
end;

class function TUstr.GetChineseWord(nI_Num: Integer): String;
begin
      case nI_Num of
       1:
         Result := '壹';
       2:
         Result := '貳';
       3:
         Result := '參';
       4:
         Result := '肆';
       5:
         Result := '伍';
       6:
         Result := '陸';
       7:
         Result := '柒';
       8:
         Result := '捌';
       9:
         Result := '玖';
       0:
         Result := '零';
      end;

end;

class function TUstr.ReverseStr(sI_SrcString: String): String;
var
    ii : Integer;
    sL_Result : String;
begin
    sL_Result := '';
    for ii:=Length(sI_SrcString) downto 1 do
    begin
      sL_Result := sL_Result + sI_SrcString[ii];
    end;
    Result := sL_Result;
end;

class function TUstr.SepString(sI_SrcString, sI_SepFlag: String): String;
var
    ii : Integer;
    sL_Result : String;
begin
    sL_Result := '';
    for ii:=1 to Length(sI_SrcString) do
    begin
      if sI_SrcString[ii]<>sI_SepFlag then
        sL_Result := sL_Result + sI_SrcString[ii]
      else
        break;
    end;
    Result := sL_Result;
end;

class function TUstr.TrimString(sI_SrcString,
  sI_NoneuseFlag: String): String;
var
    ii : Integer;
    sL_Result : String;
begin
    sL_Result := '';
    for ii:=1 to Length(sI_SrcString) do
    begin
      if sI_SrcString[ii]<>sI_NoneuseFlag then
        sL_Result := sL_Result + sI_SrcString[ii];
    end;
    Result := sL_Result;
end;

class function TUstr.ConbineChineseUnit(sI_ChineseNum: String;
  nI_Pos: Integer): String;
begin
    case nI_Pos of
     10:
      Result := sI_ChineseNum + '拾';
     9:
      Result := sI_ChineseNum + '億';
     7,3:
      Result := sI_ChineseNum + '佰';
     6,2:
      Result := sI_ChineseNum + '拾';
     5:
      Result := sI_ChineseNum + '萬';
     8,4:
      Result := sI_ChineseNum + '仟';
     1:
      Result := sI_ChineseNum;
    end;
end;

class function TUstr.getOracleSQLDateStr(dI_Date: TDate): String;
var
    wL_Year, wL_Month, wL_Day : Word;
    sL_Result, sL_DateStr : String; 
begin
    if (dI_Date=0) or (VarIsNull(dI_Date)) then
      sL_Result := 'null'
    else
    begin
      DecodeDate(dI_Date,wL_Year, wL_Month, wL_Day);
      sL_DateStr := format('%.4d/%.2d/%.2d', [wL_Year, wL_Month, wL_Day]);
      sL_Result := 'to_date(' + '''' + sL_DateStr + '''' + ',' + '''' + 'YYYY/MM/DD' + '''' + ')';
    end;
    result := sL_Result;
end;

class function TUstr.getOracleSQLDateTimeStr(dI_Date: TDateTime): String;
var
    wL_Year, wL_Month, wL_Day : Word;
    wL_Hour, wL_Min, wL_Sec, wL_MSec : Word;
    sL_Result, sL_DateTimeStr : String;
begin
    if dI_Date=0 then
      sL_Result := 'null'
    else
    begin
      DecodeDate(dI_Date,wL_Year, wL_Month, wL_Day);
      DecodeTime(dI_Date, wL_Hour, wL_Min, wL_Sec, wL_MSec);
      sL_DateTimeStr := format('%.4d/%.2d/%.2d %.2d:%.2d:%.2d', [wL_Year, wL_Month, wL_Day,wL_Hour, wL_Min, wL_Sec ]);
      sL_Result := 'to_date(' + '''' + sL_DateTimeStr + '''' + ',' + '''' + 'YYYY/MM/DD HH24:MI:SS' + '''' + ')';
    end;
    result := sL_Result;
end;

class procedure TUstr.saveStr2File(sI_Str, sI_FullPathFileName: String);
var
    L_StrList : TStringList;
begin
    L_StrList := TStringList.Create;
    L_StrList.Add(sI_Str);
    L_StrList.SaveToFile(sI_FullPathFileName);
    L_StrList.Free;
end;

class function TUstr.replaceStr(sI_SourceStr, sI_OldStr,
  sI_NewStr: String): String;
var
    sL_Result : String;
    nL_TargetPos : Integer;
begin
    nL_TargetPos := Pos(sI_OldStr, sI_SourceStr);
    while nL_TargetPos>0 do
    begin
      Delete(sI_SourceStr, nL_TargetPos, Length(sI_OldStr));
      Insert(sI_NewStr, sI_SourceStr, nL_TargetPos);
      nL_TargetPos := Pos(sI_OldStr, sI_SourceStr);      
    end;
    sL_Result := sI_SourceStr;
    result := sL_Result;
end;


class function TUstr.ChangeMinusTag(fI_Num: double): string;
var
    sL_LeftFlag, sL_RightFlag : String;
    sL_Result : String;
begin
    sL_LeftFlag := LEFT_MINUS_FLAG;
    sL_RightFlag := RIGHT_MINUS_FLAG;
    if fI_Num<0 then
    begin
      sL_Result := sL_LeftFlag + CommaNumber(FloatToStr(abs(fI_Num))+AMT_TAIL) + sL_RightFlag;
    end
    else
      sL_Result := CommaNumber(FloatToStr(fI_Num)+AMT_TAIL);

    result := sL_Result;
end;

class function TUstr.ChangeMinusTag(sI_Num: String): string;
var
    sL_Tmp : String;
begin
    sL_Tmp := replaceStr(sI_Num,',','');
    result := ChangeMinusTag(StrToFloat(sL_Tmp));
end;

class function TUstr.ParseStrings(Strs, sI_Sep: string): TStringList;
var
  ii, jj : Integer ;
  slst : TStringList ;
begin
  if Strs = '' then
  begin
    Result := nil ;
    Exit ;
  end ;
  slst := TStringList.Create ;
  jj := 1 ;
  for ii:= 1 to Length(Strs) do
  begin
    if (Strs[ii] = sI_Sep) and ((ii+1<=Length(Strs)) and (Strs[ii+1] <> sI_Sep))then
    begin
      slst.Add(Trim(Copy(Strs, jj, ii-jj))) ;
      jj := ii + 1 ;
    end ;
  end ;
  //if jj <> Length(Strs) + 1 then
    slst.Add(Trim(Copy(Strs, jj, Length(Strs)-jj+1))) ;
  Result := slst ;
end ;

class function TUstr.ChangeMinusTag(fI_Num: double;
  bI_IncludeTail: boolean): string;
var
    sL_Result : String;
begin
    sL_Result := ChangeMinusTag(fI_Num);
    if not bI_IncludeTail then
    begin
      sL_Result := replaceStr(sL_Result,AMT_TAIL,'');
    end;
    result := sL_Result;
end;

class function TUstr.getPureNumString(sI_Value: String): String;
var
    sL_TmpValue : String;
begin
    sL_TmpValue := TUstr.replaceStr(sI_Value,LEFT_MINUS_FLAG,'');
    sL_TmpValue := TUstr.replaceStr(sL_TmpValue,RIGHT_MINUS_FLAG,'');
    sL_TmpValue := TUstr.replaceStr(sL_TmpValue,',','');

    if AnsiPos(LEFT_MINUS_FLAG,sI_Value)>0 then
      sL_TmpValue := '-' + sL_TmpValue;
    result := sL_TmpValue;
end;

end.

