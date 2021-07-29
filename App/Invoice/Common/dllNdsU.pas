unit dllNdsU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

  procedure myMD5_RealMD5_8(sI_SecurityType, sI_Data:String; sI_DataLength : String; sI_Key : String; sI_KeyLen:String;  var sI_Output:String);
implementation


function MD5_RealMD5_8(sI_Data:PChar; sI_DataLength : UINT; sI_Key : PChar; sI_KeyLen : UINT; sI_Output:PChar):UINT; cdecl	; external 'MD5.dll' name '_NDS_RealMD5_8';
function NDS_RealMD5_8(sI_Data:PChar; sI_DataLength : UINT; sI_Key : PChar; sI_KeyLen : UINT; sI_Output:PChar):UINT; cdecl	; external 'NDS.dll' name '_NDS_RealMD5_8';


procedure myMD5_RealMD5_8(sI_SecurityType, sI_Data:String; sI_DataLength : String; sI_Key : String; sI_KeyLen : String; var sI_Output:String);
var
    nL_KeyLength, nL_DataLength : UINT;
    aL_Key: array[0..15] of Char;

    aL_DataLength : array[0..3] of Char;
    aL_Data: array[0..1024] of Char;
    aL_Output: array[0..15] of Char;

    sL_Result, sL_TmpHex : String;
    jj, nL_Tmp : Integer;
begin

    StrPCopy(aL_Key, sI_Key);
    nL_KeyLength := length(sI_Key);
    nL_DataLength := length(sI_Data);
    StrPCopy(aL_Data, sI_Data);

    if UpperCase(sI_SecurityType)='M' then
      NDS_RealMD5_8(aL_Data, nL_DataLength, aL_Key,nL_KeyLength, aL_Output)
    else
      MD5_RealMD5_8(aL_Data, nL_DataLength, aL_Key,nL_KeyLength, aL_Output);

    for jj:=0 to 7 do
    begin

      nL_Tmp := UInt(aL_Output[jj]);
      sL_TmpHex := IntToHex(nL_Tmp,2);
      sL_Result := sL_Result + sL_TmpHex;

    end;

    sI_Output := sL_Result;
end;



end.
