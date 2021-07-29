unit ZintInterface;

interface
 
{$M+}
 
uses Windows, Classes, SysUtils, Graphics, Dialogs;
 
type
  TZSymbol = record
    symbology: Integer;
    height: Integer;
    whitespace_width: Integer;
    border_width: Integer;
    output_options: Integer;
    fgcolour: array[0..9] of AnsiChar;
    bgcolour: array[0..9] of AnsiChar;
    outfile: array [0..255] of AnsiChar;
    scale: double;
    option_1: Integer;  // ®e¿ù²v 1.L¡B 2.M ¡B3.Q¡B 4.H
    option_2: Integer;  //Version 1-40
    option_3: Integer;
    show_human_readable_text: Integer;
    input_mode: Integer;
    text: array[0..127] of AnsiChar;
    rows: Integer;
    width: Integer;
    primary: array [0..127] of AnsiChar;
    encoded_data: array[0..177] of array[0..177] of AnsiChar;
    row_height: array[0..177] of Integer;
    errtxt: array[0..99] of AnsiChar;
    bitmap: PAnsiChar;
    bitmap_width: Integer;
    bitmap_height: Integer;
    rendered: Pointer;
    end;
  PZSymbol = ^TZSymbol;
 
const
  LibName = 'zint.dll';
 
  function ZBarcode_Create: PZSymbol; cdecl; external LibName;
  procedure ZBarcode_Delete(symbol: PZSymbol); cdecl; external LibName;
  procedure ZBarcode_Clear(symbol: PZSymbol); cdecl; external LibName;
  function ZBarcode_Buffer(symbol: PZSymbol; rotate_angle: Integer): Integer; cdecl; external LibName;
  function ZBarcode_Encode_and_Buffer(symbol: PZSymbol; input: PAnsiChar; length: Integer; rotate_angle: Integer): Integer; cdecl; external LibName;
 
  procedure ZBarcodeToBitmap(ASymbol: PZSymbol; const ABitmap: TBitmap);
 
implementation
 
procedure ZBarcodeToBitmap(ASymbol: PZSymbol; const ABitmap: TBitmap);
var myp: PRGBTriple;
    row: Integer;
    rowwidth: Integer;
begin
  ABitmap.PixelFormat := pf24bit;
  ABitmap.Width := ASymbol.bitmap_width;
  ABitmap.Height := ASymbol.bitmap_height;

  myp := Pointer(ASymbol.bitmap);
  rowwidth := Asymbol.bitmap_width * 3;

  for row := 0 to ASymbol.bitmap_height - 1 do
  begin
    CopyMemory(ABitmap.ScanLine[row], myp, rowwidth);
    Inc(myp, Asymbol.bitmap_width);
  end;
end;
 
end.
