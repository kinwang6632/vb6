unit Uotheru;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Math;

type
  TSex = (sFemale, sMale);
  TUother = class
  private
  public
    class function RoundFloat(Src : Double; nRetainPos : Integer): Double;
  end;


implementation


{ TUother }

class function TUother.RoundFloat(Src: Double;
  nRetainPos: Integer): Double;
var
   bL_Minus : Boolean;
   sL_Tmp:String;
   sSrc : String;
   nPos : Integer;
   fTmp_Result, fL_RetainNum : double;
begin
     //ex : Src=258.124, nRetainPos=2
     sSrc := FloatToStr(Src);//sSrc='258.124'

     bL_Minus := False;
     if sSrc[1]='-' then//為負值
     begin
      bL_Minus := True;
      Src := ABS(Src);      
      sSrc := Copy(sSrc,2,Length(sSrc)-1);

     end;
          
     nPos := AnsiPos('.',sSrc);//nPos=4
     if nPos=0 then//找不到'.'
     begin
       if bL_Minus then
         Result := 0 - Src
       else
         Result := Src;

       Exit;
     end;


     sL_Tmp := Copy(sSrc,nPos+1,nRetainPos);//sL_Tmp='12'
     if Length(sL_Tmp)<nRetainPos then
     begin
       if bL_Minus then
         Result := 0 - Src
       else
         Result := Src;
       Exit;
     end;
     sSrc := Copy(sSrc,1,nPos+nRetainPos);//sSrc='258.12'
     if FloatToStr(Src)[nPos+nRetainPos+1]<='4' then//以下捨去
       fTmp_Result := StrToFloat(sSrc)
     else
     begin
       fL_RetainNum := Power(10, (-1*nRetainPos));
       fTmp_Result := StrToFloat(sSrc) + fL_RetainNum;
     end;
     if bL_Minus then
       fTmp_Result := 0 - fTmp_Result;
     Result := fTmp_Result;
end;

end.


