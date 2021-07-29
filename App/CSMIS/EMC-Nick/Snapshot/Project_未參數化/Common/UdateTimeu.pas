unit UdateTimeu;

interface

uses
  Windows, Messages, SysUtils, Classes, Controls ;

Type
  TUdateTime = class
  public
    class function GetFormatDateStr(sI_SrcDate, sI_Sep:String; bI_IncludingHeader, bI_ChineseFormat : boolean):String;overload;//傳入 2002/02/02, 傳回 "91年02月02日"
    class function GetPureDateStr(dI_Date:TDate):String;
    class function GetPureDate(dI_Date:TDate):TDate;
    class function TransNumMonthToEngMonth(nI_NumMnth : Integer):String;
    class function IsDateStr(asSrc:String;Sep:Char):Boolean;
    class function IsTimeStr(asSrc:String;Sep:Char):Boolean;
    class function IsYMStr(asSrc:String;Sep:Char):Boolean;
    class function CDateStr(aDt:TDateTime; iLen:Integer):string;
    class function CDate2Date(asCDate:string):TDateTime;
    class function TestCDate(asCDate:string):boolean;
    class function NextCDate(asCDate:string; iDay:Integer):string;
    class function DaysOfMonth(asCDate:string):Integer;overload;//傳入920501(民國之日期),取出該月份之天數
    class function DaysOfMonth(I_Date:TDate):Integer;overload;//傳入2002/02/05(西元之日期),取出該月份之天數
    class function IsLeap(asCDate:string):boolean; overload;
    class function IsLeap(asdtDate:TDateTime):boolean; overload;
    class function AgeNum(asCBirthDate:string):Integer; overload;
    class function AgeNum(asEBirthDate:TDateTime):Integer; overload;
    class function AgeStr(asCBirthDate:string):string; overload;
    class function AgeStr(asEBirthDate:TDateTime):string; overload;
    class function DaysDiff(asStartCDate, asEndCDate:string):Integer;
    class function DaySession(aDt:TDateTime):Integer;
    class function DaySessionDisp(aDt:TDateTime):string;
    class function WeekId(aDt:TDateTime):Integer; overload ;
    class function WeekId(asCDate:string):Integer; overload ;
//    class function adjDate(sI_Date:String; sI_Sep:Char):String;//Mulder
    class function countMonth(I_StartDate,I_EndDate:TDateTime):Integer  ; //Mulder
    class function composeMonth(I_date:TDateTime;nI_Number:Integer):String ;//Mulder
    class function getMonthFirstDay(sI_Date : String) : String;
  private
    class function ParseStrings(Strs:string; Sep:Char) : TStringList;

  end;

implementation

uses Ustru;

{ TUdateTime }

class function TUdateTime.AgeNum(asCBirthDate: string): Integer;
var
  vDt : TDateTime ;
begin
  vDt := CDate2Date(asCBirthDate) ;
  Result := AgeNum(vDt) ;
end;

class function TUdateTime.AgeNum(asEBirthDate: TDateTime): Integer;
var
  yy1, yy2, mm, dd : Word ;
begin
  DecodeDate(asEBirthDate, yy1, mm, dd) ;
  DecodeDate(Date, yy2, mm, dd) ;
  Result := yy2 - yy1 + 1  ;
end;

class function TUdateTime.AgeStr(asEBirthDate: TDateTime): string;
begin
  Result := IntToStr(AgeNum(asEBirthDate)) ;
end;

class function TUdateTime.AgeStr(asCBirthDate: string): string;
begin
  Result := IntToStr(AgeNum(asCBirthDate)) ;
end;

class function TUdateTime.CDate2Date(asCDate: string): TDateTime;
var
  yy, mm, dd: Word;
begin
  yy := 0; mm := 0; dd := 0 ;
  case Length(asCDate) of
    6, 7 :
    begin
      yy := StrtoInt(Copy(asCDate, 1, Length(asCDate)-4))+ 1911;
      mm := StrtoInt(Copy(asCDate, Length(asCDate)-4 +1, 2));
      dd := StrtoInt(Copy(asCDate,Length(asCDate)-2 +1, 2));

    end ;
    8, 9 :
    begin
      yy := StrtoInt(Copy(asCDate, 1, Length(asCDate)-6))+ 1911;
      mm := StrtoInt(Copy(asCDate, Length(asCDate)-5 +1, 2));
      dd := StrtoInt(Copy(asCDate,Length(asCDate)-2 +1, 2));
    end ;
    10 :
     begin
      yy := StrtoInt(Copy(asCDate, 1, Length(asCDate)-6));
      mm := StrtoInt(Copy(asCDate, Length(asCDate)-5 +1, 2));
      dd := StrtoInt(Copy(asCDate,Length(asCDate)-2 +1, 2));
     end;
  end ;
  Result := EncodeDate(yy, mm, dd);
end;

class function TUdateTime.CDateStr(aDt: TDateTime; iLen: Integer): string;
var
  yy, mm, dd : Word ;
begin
  DecodeDate(aDt, yy, mm ,dd);
  case iLen of
    6 :
      Result := Format('%.*d',[2, yy - 1911]) + Format('%.*d',[2,mm])
        + Format('%.*d',[2,dd]);
    7 :
      Result := Format('%.*d',[3, yy - 1911]) + Format('%.*d',[2,mm])
        + Format('%.*d',[2,dd]);
    8 :
      Result := Format('%.*d',[2, yy - 1911]) + '/' + Format('%.*d',[2,mm]) +
        '/' + Format('%.*d',[2,dd]);
    9 :
      Result := Format('%.*d',[3, yy - 1911]) + '/' + Format('%.*d',[2,mm]) +
        '/' + Format('%.*d',[2,dd]);
  end ;
end;

class function TUdateTime.DaysDiff(asStartCDate,
  asEndCDate: string): Integer;
var
  ii : Integer ;
  vDt1, vDt2 : TDateTime ;
begin
  vDt1 := CDate2Date(asStartCDate) ;
  vDt2 := CDate2Date(asEndCDate) ;
  ii := 0 ;
  if VDt2 > VDt1 then
  begin
    while (vDt1+ii) <> VDt2 do
      ii := ii + 1 ;
    Result := ii ;
  end
  else
  begin
    while (vDt2+ii) <> VDt1 do
      ii := ii + 1 ;
    Result := -ii ;
  end
end;

class function TUdateTime.DaySession(aDt: TDateTime): Integer;
begin
  if (aDt >= StrToTime('AM 08:00')) and (aDt < StrToTime('PM 01:00')) then
    Result := 0
  else if (aDt >= StrToTime('PM 01:00')) and (aDt < StrToTime('PM 06:00')) then
    Result := 1
  else if (aDt >= StrToTime('PM 06:00')) and (aDt < StrToTime('PM 10:30')) then
    Result := 2
  else Result := 0;
end;

class function TUdateTime.DaySessionDisp(aDt: TDateTime): string;
begin
  if (aDt >= StrToTime('AM 08:00')) and (aDt < StrToTime('PM 12:10')) then
    Result := '上午'
  else if (aDt >= StrToTime('PM 12:10')) and (aDt < StrToTime('PM 05:00')) then
    Result := '中午'
  else if (aDt >= StrToTime('PM 05:00')) and (aDt < StrToTime('PM 11:30')) then
    Result := '晚上'
  else Result := '上午';
end;

class function TUdateTime.DaysOfMonth(asCDate: string): Integer;
var
  yy, mm, dd : Word ;
  vbLeap: Boolean;
  vDt : TDateTime ;
begin
  vDt := CDate2Date(asCDate) ;
  DecodeDate(vDt, yy, mm ,dd);
  vbLeap := IsLeap(vDt) ;
  case mm of
    1,3,5,7,8,10,12: Result := 31;
    4,6,9,11: Result := 30
    else
    if vbLeap then
      Result := 29
    else
      Result := 28;
  end;
end;

class function TUdateTime.IsLeap(asCDate: string): boolean;
var
  yy, mm, dd : Word ;
  vDt : TDateTime ;
begin
  vDt := CDate2Date(asCDate) ;
  DecodeDate(vDt, yy, mm ,dd);
  Result := (yy mod 4 = 0) and (yy mod 100 <> 0) or (yy mod 400 = 0);
end;

class function TUdateTime.IsDateStr(asSrc: String;Sep:Char): Boolean;
var
   tmp_StrList : TStringList;
   tmp_Year, tmp_Month, tmp_Day : String;
   Code: Integer;
   nYear, nMonth,nDay : Integer;
begin
     Result := FALSE;
     tmp_StrList := ParseStrings(asSrc,Sep);
     tmp_Year := tmp_StrList.Strings[0];
     tmp_Month := tmp_StrList.Strings[1];
     tmp_Day := tmp_StrList.Strings[2];

     Val(tmp_Year, nYear, Code);
     if (Code<>0) then Exit;

     Val(tmp_Month, nMonth, Code);
     if (Code<>0) or (nMonth>12) or (nMonth<1)  then Exit;

     Val(tmp_Day, nDay, Code);
     case nMonth of
       1,3,5,7,8,10,12 ://大月
        begin
          if (Code<>0) or (nDay>31) or (nDay<1)  then
          begin
            Result := FALSE;
            Exit;
          end;
        end;
        4,6,9,11://小月
         begin
          if (Code<>0) or (nDay>30) or (nDay<1)  then
          begin
            Result := FALSE;
            Exit;
          end;
         end;
        2://二月
         begin
           if (nYear Mod 4)=0 then //潤年
           begin
            if (Code<>0) or (nDay>29) or (nDay<1)  then
            begin
              Result := FALSE;
             Exit;
            end;
           end
           else
           begin
            if (Code<>0) or (nDay>28) or (nDay<1)  then
            begin
              Result := FALSE;
             Exit;
            end;
           end;
         end
     end;
     Result := TRUE;

end;

class function TUdateTime.IsLeap(asdtDate: TDateTime): boolean;
var
  yy, mm, dd : Word ;
begin
  DecodeDate(asdtDate, yy, mm ,dd);
  Result := (yy mod 4 = 0) and (yy mod 100 <> 0) or (yy mod 400 = 0);
end;

class function TUdateTime.IsTimeStr(asSrc: String;Sep:Char): Boolean;
var
   tmp_Hour, tmp_Min : String;
   Code, nHour, nMin : Integer;
   tmp_StrList : TStringList;
begin
     Result := FALSE;
     tmp_StrList := ParseStrings(asSrc,Sep);
     tmp_Hour := tmp_StrList.Strings[0];
     tmp_Min := tmp_StrList.Strings[1];

     Val(tmp_hour, nHour, Code);
     if (Code<>0) or (nHour>23) or (nHour<0) then Exit;

     Val(tmp_Min, nMin, Code);
     if (Code<>0) or (nMin>60) or (nHour<0) then Exit;
     Result := TRUE;
end;

class function TUdateTime.NextCDate(asCDate: string;
  iDay: Integer): string;
var
  vDt : TDateTime ;
begin
  vDt := CDate2Date(asCDate) ;
  vDt := vDt + iDay ;
  Result := CDateStr(vDt, Length(ascDate)) ;
end;

class function TUdateTime.TestCDate(asCDate: string): boolean;
begin
  Result := False;
  try
    if Copy(DateToStr(CDate2Date(asCDate)), 1, 4) = '1911' then
      Result := False
    else
      Result := True;
  except
  end;
end;

class function TUdateTime.WeekId(asCDate: string): Integer;
var
  vDt : TDateTime ;
begin
  vDt := CDate2Date(asCDate) ;
  Result := WeekId(vDt) ;
end;

class function TUdateTime.WeekId(aDt: TDateTime): Integer;
begin
  Result := DayOfWeek(aDt) ;
end;

class function TUdateTime.ParseStrings(Strs: string;
  Sep: Char): TStringList;
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
    if Strs[ii] = Sep then
    begin
      slst.Add(Trim(Copy(Strs, jj, ii-jj))) ;
      jj := ii + 1 ;
    end ;
  end ;
  //if jj <> Length(Strs) + 1 then
    slst.Add(Trim(Copy(Strs, jj, Length(Strs)-jj+1))) ;
  Result := slst ;
end ;

class function TUdateTime.TransNumMonthToEngMonth(nI_NumMnth : Integer): String;
begin
    case nI_NumMnth of
     1:
       Result := 'JAN';
     2:
       Result := 'FEB';
     3:
       Result := 'MAR';
     4:
       Result := 'APR';
     5:
       Result := 'MAY';
     6:
       Result := 'JUN';
     7:
       Result := 'JUL';
     8:
       Result := 'AUG';
     9:
       Result := 'SEP';
     10:
       Result := 'OCT';
     11:
       Result := 'NOV';
     12:
       Result := 'DEC';     
    end; 
end;

class function TUdateTime.GetPureDateStr(dI_Date: TDate): String;
var
    wL_Year,wL_Month,wL_Day : Word;
begin
     DecodeDate(dI_Date,wL_Year,wL_Month,wL_Day);
     Result := Format('%.4d/%.2d/%.2d',[wL_Year,wL_Month,wL_Day]);
end;

class function TUdateTime.GetPureDate(dI_Date: TDate): TDate;
var
    sL_Date : String;
begin
    sL_Date := GetPureDateStr(dI_Date);
    Result := StrToDate(sL_Date);
end;

class function TUdateTime.DaysOfMonth(I_Date: TDate): Integer;
var
  yy, mm, dd : Word ;
  vbLeap: Boolean;
  vDt : TDateTime ;
begin
  vDt := I_Date;
  DecodeDate(vDt, yy, mm ,dd);
  vbLeap := IsLeap(vDt) ;
  case mm of
    1,3,5,7,8,10,12: Result := 31;
    4,6,9,11: Result := 30
    else
    if vbLeap then
      Result := 29
    else
      Result := 28;
  end;
end;

{
class function TUdateTime.adjDate(sI_Date: String; sI_Sep: Char): String;
  //傳入2001/1/1 傳回'20010101'   //傳入2001/1 傳回'200101'
var
  L_Result : TStringList ;
  sL_Year,sL_Month,sL_Day : String ;
begin
  L_Result := TUstr.ParseStrings(sI_Date,sI_Sep);
  sL_Year := L_Result.Strings[0];
  sL_Month := L_Result.Strings[1];
  if L_Result.Count = 3 then  //傳入年月日
  begin
    sL_Day := L_Result.Strings[2];
    if Length(sL_Month) = 1 then
      sL_Month := '0'+sL_Month;
    if Length(sL_Day) = 1 then
      sL_Day := '0'+sL_Day;
    Result := sL_Year+sL_Month+sL_Day;
  end
  else //僅傳入年月
  begin
    if Length(sL_Month) = 1 then
      sL_Month := '0'+sL_Month;
    Result := sL_Year+sL_Month;
  end;
end;
}
class function TUdateTime.countMonth(I_StartDate,I_EndDate: TDateTime): Integer;
//傳入二個日期，傳回兩者之間相差幾個月份
var
  S_Year,S_Month,E_Year,E_Month,Day :Word ;
  nL_YDiffer,nL_MDiffer :Integer;
begin
  DecodeDate(I_StartDate,S_Year,S_Month,Day);
  DecodeDate(I_EndDate,E_Year,E_Month,Day);
  nL_YDiffer := E_Year-S_Year;
  nL_MDiffer := E_Month-S_Month;
  Result := nL_YDiffer * 12 + nL_MDiffer;
end;

class function TUdateTime.composeMonth(I_date: TDateTime; nI_Number: Integer): String;
//傳入時間(2001/1/1)，個數(3)  傳回  200101,200102,200103
var
  Year,Month,Day : Word ;
  sL_Month,sL_Result : String ;
  i : Integer ;
begin
  DecodeDate(I_date,Year,Month,Day);
  for i := 0 to nI_Number-1 do
  begin
    if Month <> 12 then
    begin
      if Length(IntToStr(Month))<>2 then
        sL_Month := '0'+ IntToStr(Month)
      else
        sL_Month :=  IntToStr(Month);
      sL_Result := sL_Result +','+ (IntToStr(Year)+'/'+sL_Month);
      Inc(Month);
    end
    else
    begin
      if Length(IntToStr(Month))<>2 then
        sL_Month := '0'+ IntToStr(Month)
      else
        sL_Month :=  IntToStr(Month);
      sL_Result := sL_Result +','+ (IntToStr(Year)+'/'+sL_Month);
      Inc(Year);
      Month := 1;
    end;
  end;
  result := Copy(sL_Result,2,Length(sL_Result)-1);

end;



class function TUdateTime.GetFormatDateStr(sI_SrcDate, sI_Sep:String;
  bI_IncludingHeader, bI_ChineseFormat: boolean): String;
  {
    參數說明 :
     sI_SrcDate : 原始日期字串 => 可傳入之字串格式為 : 'YYYY/MM/DD' or 'YYYY/MM' or 'YYYY'
     sI_Sep: 年月日間之區隔符號 => (1).若 sI_Sep 為空字串=>不要區隔符號 (2).若 sI_Sep 為'/'=>用'/'為區隔符號 (3).若 sI_Sep 為'#'=>使用'年月日' 等中文字為區隔符號
     bI_IncludingHeader : 傳回值是否包含'中華民國 or '西元' 等 header
     bI_ChineseFormat : 是否使用中華民國曆
  }
var
    nL_Cond : Integer;
    dL_Date : TDate;
    sL_Header, sL_YSep, sL_MSep, sL_DSep : String;
    sL_Year, sL_Month, sL_Day, sL_Result : String;
    wL_Year, wL_Month, wL_Day : word;
begin
    case length(sI_SrcDate) of
      4://僅年
       begin
         nL_Cond := 1;
         wL_Year := StrToInt(sI_SrcDate);
         sL_Month := '';
         sL_Day := '';
       end;
      7://僅年月
       begin
         nL_Cond := 2;
         wL_Year := StrToInt(Copy(sI_SrcDate,1,4));
         wL_Month := StrToInt(Copy(sI_SrcDate,6,2));
         sL_Month := Format('%.2d',[wL_Month]);
         sL_Day := '';
       end;
      10://完整日期
       begin
         nL_Cond := 3;       
         dL_Date := StrToDate(sI_SrcDate);
         DecodeDate(dL_Date,wL_Year, wL_Month, wL_Day);
         sL_Month := Format('%.2d',[wL_Month]);
         sL_Day := Format('%.2d',[wL_Day]);
       end;
    end;

    if sI_Sep='' then
    begin
      sL_YSep := '';
      sL_MSep := '';
      sL_DSep := '';
    end
    else if sI_Sep='#' then
    begin
      sL_YSep := '年';
      sL_MSep := '月';
      sL_DSep := '日';
    end
    else
    begin
      sL_YSep := sI_Sep;
      sL_MSep := sI_Sep;
      sL_DSep := '';
    end;


    if bI_ChineseFormat then
    begin
      sL_Year := Format('%d',[wL_Year-1911]);
    end
    else
    begin
      sL_Year := Format('%.4d',[wL_Year]);
    end;

    if bI_IncludingHeader and bI_ChineseFormat then
      sL_Header := '中華民國'
    else if bI_IncludingHeader and (not bI_ChineseFormat) then
      sL_Header := '西元'
    else
      sL_Header := '';

    case nL_Cond of
      1: //僅年
       begin
         if bI_ChineseFormat then
           sL_Result := sL_Header +sL_Year +sL_YSep
         else
           sL_Result := sL_Header +sL_Year;
       end;
      2: //僅年月
       begin
         if bI_ChineseFormat then
         begin
          if sL_MSep <> sI_Sep then
            sL_Result := sL_Header +sL_Year +sL_YSep + sL_Month + sL_MSep
          else
            sL_Result := sL_Header +sL_Year +sL_YSep + sL_Month
         end
         else
           sL_Result := sL_Header +sL_Year +sL_YSep + sL_Month;
       end;
      3: //完整日期
       begin
         if bI_ChineseFormat then
           sL_Result := sL_Header +sL_Year +sL_YSep + sL_Month + sL_MSep + sL_Day + sL_DSep
         else
           sL_Result := sL_Header +sL_Year +sL_YSep + sL_Month + sL_MSep + sL_Day ;
       end;
    end;
    result := sL_Result;

end;


class function TUdateTime.IsYMStr(asSrc: String; Sep: Char): Boolean;
var
   tmp_StrList : TStringList;
   tmp_Year, tmp_Month : String;
   Code: Integer;
   nYear, nMonth : Integer;
begin
     Result := FALSE;
     tmp_StrList := ParseStrings(asSrc,Sep);
     tmp_Year := tmp_StrList.Strings[0];
     tmp_Month := tmp_StrList.Strings[1];


     Val(tmp_Year, nYear, Code);
     if (Code<>0) then Exit;

     Val(tmp_Month, nMonth, Code);
     if (Code<>0) or (nMonth>12) or (nMonth<1)  then Exit;


     Result := TRUE;

end;

class function TUdateTime.getMonthFirstDay(sI_Date: String): String;
var
  Year,Month,Day : Word ;
  sL_MonthFirstDay : String;
begin
  //取得傳入年月的第一天
  DecodeDate(StrToDate(sI_Date),Year,Month,Day);
  sL_MonthFirstDay := IntToStr(Year) + '/' + IntToStr(Month) + '/01';
  Result := sL_MonthFirstDay;
end;

end.
