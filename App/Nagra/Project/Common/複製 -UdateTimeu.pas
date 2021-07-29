unit UdateTimeu;

interface

uses
  Windows, Messages, SysUtils, Classes, Controls ;

Type
  TUdateTime = class
  public
    class function GetLocalSQLDateTime(TimeStamp: TDateTime): string;
    class function GetPureRocDateStr(dI_Date:TDate;sI_Seq:String):String;
    class function GetPureDateStr(dI_Date:TDate;sI_Seq:String):String;
    class function GetPureDate(dI_Date:TDate):TDate;
    class function TransNumMonthToEngMonth(nI_NumMnth : Integer):String;
    class function IsDateStr(asSrc:String;Sep:Char):Boolean;
    class function IsTimeStr(asSrc:String;Sep:Char):Boolean;
    class function CDateStr(aDt:TDateTime; iLen:Integer):string;
    class function CDate2Date(asCDate:string):TDateTime;
    class function TestCDate(asCDate:string):boolean;
    class function NextCDate(asCDate:string; iDay:Integer):string;
    class function DaysOfMonth(asCDate:string):Integer;
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
  private
    class function ParseStrings(Strs:string; Sep:Char) : TStringList;
  end;

implementation

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
    Result := '�W��'
  else if (aDt >= StrToTime('PM 12:10')) and (aDt < StrToTime('PM 05:00')) then
    Result := '����'
  else if (aDt >= StrToTime('PM 05:00')) and (aDt < StrToTime('PM 11:30')) then
    Result := '�ߤW'
  else Result := '�W��';
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
       1,3,5,7,8,10,12 ://�j��
        begin
          if (Code<>0) or (nDay>31) or (nDay<1)  then
          begin
            Result := FALSE;
            Exit;
          end;
        end;
        4,6,9,11://�p��
         begin
          if (Code<>0) or (nDay>30) or (nDay<1)  then
          begin
            Result := FALSE;
            Exit;
          end;
         end;
        2://�G��
         begin
           if (nYear Mod 4)=0 then //��~
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
     if (Code<>0) or (nHour>24) or (nHour<0) then Exit;

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

class function TUdateTime.GetPureDateStr(dI_Date: TDate;sI_Seq:String): String;
var
    wL_Year,wL_Month,wL_Day : Word;
begin
     DecodeDate(dI_Date,wL_Year,wL_Month,wL_Day);
     Result := Format('%.4d%s%.2d%s%.2d',[wL_Year,sI_Seq, wL_Month,sI_Seq,wL_Day]);
end;

class function TUdateTime.GetPureDate(dI_Date: TDate): TDate;
var
    sL_Date : String;
begin
    sL_Date := GetPureDateStr(dI_Date,'/');
    Result := StrToDate(sL_Date);
end;

class function TUdateTime.GetPureRocDateStr(dI_Date: TDate;
  sI_Seq: String): String;
var
    wL_Year,wL_Month,wL_Day : Word;
begin
     DecodeDate(dI_Date,wL_Year,wL_Month,wL_Day);
     Result := Format('%.3d%s%.2d%s%.2d',[wL_Year-1911,sI_Seq, wL_Month,sI_Seq,wL_Day]);
end;

class function TUdateTime.GetLocalSQLDateTime(
  TimeStamp: TDateTime): string;
var
  sL_TmpStr: string;
begin
  sL_TmpStr := datetimetostr(TimeStamp);
  result := copy(sL_TmpStr,6,5)+'/'+copy(sL_TmpStr,1,4)+copy(sL_TmpStr,11,9);
end;

end.
