unit XLSFile;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Grids, Forms, Dialogs,db,dbctrls,comctrls;

const
{BOF}
  CBOF      = $0009;
  BIT_BIFF5 = $0800;
  BOF_BIFF5 = CBOF or BIT_BIFF5;
{EOF}
  BIFF_EOF = $000a;
{Document types}
  DOCTYPE_XLS = $0010;
{Dimensions}
  DIMENSIONS = $0000;

type
  TAtributCell = (acHidden,acLocked,acShaded,acBottomBorder,acTopBorder,
                acRightBorder,acLeftBorder,acLeft,acCenter,acRight,acFill);

  TSetOfAtribut = set of TatributCell;

  TXLSWriter = class(Tobject)
  private
    fstream:TFileStream;
    procedure WriteWord(w:Longword);
  protected
    procedure WriteBOF;
    procedure WriteEOF;
    procedure WriteDimension;
  public
    maxCols,maxRows:Word;
    procedure CellWord(vCol,vRow:word;aValue:Longword;vAtribut:TSetOfAtribut=[]);
    procedure CellDouble(vCol,vRow:word;aValue:double;vAtribut:TSetOfAtribut=[]);
    procedure CellStr(vCol,vRow:word;aValue:String;vAtribut:TSetOfAtribut=[]);
    procedure WriteField(vCol,vRow:word;Field:TField);
    constructor create(vFileName:string);
    destructor destroy;override;

  end;

procedure SetCellAtribut(value:TSetOfAtribut;var FAtribut:array of byte);
procedure DataSetToXLS(ds:TDataSet;fname:String;SQLCommand:String; sI_UnvisableFieldsName:String); // sI_UnvisableFieldsName:String 不想顯示在 excel 上的欄位名稱,以"~"作為欄位區隔 // howard
procedure StringGridToXLS(grid:TStringGrid;fname:String);
function ParseStrings(Strs, sI_Sep:string) : TStringList;

implementation


procedure DataSetToXLS(ds:TDataSet;fname:String;SQLCommand:String; sI_UnvisableFieldsName:String);
var
  c, r, fc, nL_Column, aIndex: Integer;
  xls: TXLSWriter;
  sL_Title: String;
  L_StrList: TStringList;
begin
  xls:=TXLSWriter.create(fname);
  if ds.FieldCount > xls.maxcols then
    xls.maxcols:=ds.FieldCount+1;
  try
    xls.writeBOF;
    xls.WriteDimension;
    fc:=ds.FieldCount;
//拆解SQLCommand '查詢條件:@公司別:開博科技@收費項目:情色頻道'
    //sL_Temp := '查詢條件:@公司別:開博科技@收費項目:情色頻道';


    if SQLCommand <> EmptyStr then
    begin
      L_StrList := ParseStrings( SQLCommand, '@' );
      try
        for aIndex := 0 to L_StrList.Count-1 do
        begin
          sL_Title := L_StrList.Strings[aIndex];
          xls.Cellstr( aIndex, 0, sL_Title );
        end;
      finally
        L_StrList.Free;
      end;
    end;

    nL_Column := 0;
    
    for c:=0 to fc-1 do
    begin
      if AnsiPos(UpperCase(ds.Fields[c].FieldName), UpperCase(sI_UnvisableFieldsName))=0 then //howard
      begin
        xls.Cellstr( aIndex + 2,nL_Column,ds.Fields[c].DisplayLabel );
        Inc(nL_Column);
      end
    end;

    ds.First;

     //for r:= 1 to ds.RecordCount do
     r:=1;
     while not ds.Eof do
     begin
      if r <= xls.maxRows then
       begin
        nL_Column := 0;

        for c:=0 to fc-1 do
        begin
          if AnsiPos(UpperCase(ds.Fields[c].FieldName), UpperCase(sI_UnvisableFieldsName))=0 then //howard
          begin
            xls.WriteField(r+aIndex+2,nL_Column,ds.Fields[c]);
            Inc(nL_Column);
          end;
        end;

        r:=r+1;
        ds.Next;
       end
      else
       begin
        ShowMessage('資料超過65536筆!!! 將被截斷.');
        Exit;
       end;
     end;
    xls.writeEOF;
  finally
    xls.free;
  end;
end;

procedure StringGridToXLS(grid:TStringGrid;fname:String);
var c,r,rMax:Integer;
  xls:TXLSWriter;
begin
  xls:=TXLSWriter.create(fname);
  rMax:=grid.RowCount;
  if grid.ColCount > xls.maxcols then
    xls.maxcols:=grid.ColCount+1;
  if rMax > xls.maxrows then          // 此格式最多只能存 65535 Rows
    rMax:=xls.maxrows;
  try
    xls.writeBOF;
    xls.WriteDimension;
    for c:=0 to grid.ColCount-1 do
      for r:=0 to rMax-1 do
        xls.Cellstr(r,c,grid.Cells[c,r]);
    xls.writeEOF;
  finally
    xls.free;
  end;
end;

{ TXLSWriter }

constructor TXLSWriter.create(vFileName:string);
begin
  inherited create;
  if FileExists(vFilename) then
    fStream:=TFileStream.Create(vFilename,fmOpenWrite)
  else
    fStream:=TFileStream.Create(vFilename,fmCreate);

  maxCols:=100;
  maxRows:=65535;
end;

destructor TXLSWriter.destroy;
begin
  if fStream <> nil then
    fStream.free;
  inherited;
end;

procedure TXLSWriter.WriteBOF;
begin
  Writeword(BOF_BIFF5);
  Writeword(6);           // count of bytes
  Writeword(0);
  Writeword(DOCTYPE_XLS);
  Writeword(0);
end;

procedure TXLSWriter.WriteDimension;
begin
  Writeword(DIMENSIONS);  // dimension OP Code
  Writeword(8);           // count of bytes
  Writeword(0);           // min cols
  Writeword(maxRows);     // max rows
  Writeword(0);           // min rowss
  Writeword(maxcols);     // max cols
end;

procedure TXLSWriter.CellDouble(vCol, vRow: word; aValue: double;
  vAtribut: TSetOfAtribut);
var  FAtribut:array [0..2] of byte;
begin
  Writeword(3);           // opcode for double
  Writeword(15);          // count of byte
  Writeword(vCol);
  Writeword(vRow);
  SetCellAtribut(vAtribut,fAtribut);
  fStream.Write(fAtribut,3);
  fStream.Write(aValue,8);
end;

procedure TXLSWriter.CellWord(vCol,vRow:word;aValue:Longword;vAtribut:TSetOfAtribut=[]);
var  FAtribut:array [0..2] of byte;
begin
  Writeword(2);           // opcode for word
  Writeword(9);           // count of byte
  Writeword(vCol);
  Writeword(vRow);
  SetCellAtribut(vAtribut,fAtribut);
  fStream.Write(fAtribut,3);
  Writeword(aValue);
end;

procedure TXLSWriter.CellStr(vCol, vRow: word; aValue: String;
  vAtribut: TSetOfAtribut);
var  FAtribut:array [0..2] of byte;
  slen:byte;
begin
  Writeword(4);           // opcode for string
  slen:=length(avalue);
  Writeword(slen+8);      // count of byte
  Writeword(vCol);
  Writeword(vRow);
  SetCellAtribut(vAtribut,fAtribut);
  fStream.Write(fAtribut,3);
  fStream.Write(slen,1);
  if slen > 0 then
    fStream.Write(aValue[1],slen); //howard
end;

procedure SetCellAtribut(value:TSetOfAtribut;var FAtribut:array of byte);
var
   i:integer;
begin
 //reset
  for i:=0 to High(FAtribut) do
    FAtribut[i]:=0;

     {Byte Offset     Bit   Description                     Contents
     0          7     Cell is not hidden              0b
                      Cell is hidden                  1b
                6     Cell is not locked              0b
                      Cell is locked                  1b
                5-0   Reserved, must be 0             000000b
     1          7-6   Font number (4 possible)
                5-0   Cell format code
     2          7     Cell is not shaded              0b
                      Cell is shaded                  1b
                6     Cell has no bottom border       0b
                      Cell has a bottom border        1b
                5     Cell has no top border          0b
                      Cell has a top border           1b
                4     Cell has no right border        0b
                      Cell has a right border         1b
                3     Cell has no left border         0b
                      Cell has a left border          1b
                2-0   Cell alignment code
                           general                    000b
                           left                       001b
                           center                     010b
                           right                      011b
                           fill                       100b
                           Multiplan default align.   111b
     }

     //  bit sequence 76543210

     if  acHidden in value then       //byte 0 bit 7:
         FAtribut[0] := FAtribut[0] + 128;

     if  acLocked in value then       //byte 0 bit 6:
         FAtribut[0] := FAtribut[0] + 64 ;

     if  acShaded in value then       //byte 2 bit 7:
         FAtribut[2] := FAtribut[2] + 128;

     if  acBottomBorder in value then //byte 2 bit 6
         FAtribut[2] := FAtribut[2] + 64 ;

     if  acTopBorder in value then    //byte 2 bit 5
         FAtribut[2] := FAtribut[2] + 32;

     if  acRightBorder in value then  //byte 2 bit 4
         FAtribut[2] := FAtribut[2] + 16;

     if  acLeftBorder in value then   //byte 2 bit 3
         FAtribut[2] := FAtribut[2] + 8;

     // <2002-11-17> dllee 最後 3 bit 應只有 1 種選擇
     if  acLeft in value then         //byte 2 bit 1
         FAtribut[2] := FAtribut[2] + 1
     else if  acCenter in value then  //byte 2 bit 1
         FAtribut[2] := FAtribut[2] + 2
     else if acRight in value then    //byte 2, bit 0 dan bit 1
         FAtribut[2] := FAtribut[2] + 3
     else if acFill in value then     //byte 2, bit 0
         FAtribut[2] := FAtribut[2] + 4;
end;

procedure TXLSWriter.WriteWord(w: Longword);
begin
  fstream.Write(w,2);
end;

procedure TXLSWriter.WriteEOF;
begin
  Writeword(BIFF_EOF);
  Writeword(0);
end;

procedure TXLSWriter.WriteField(vCol, vRow: word; Field: TField);
begin
  case field.DataType of
     ftString,ftWideString,ftBoolean,ftDate,ftDateTime,ftTime,ftTimeStamp:
       Cellstr(vCol,vRow,field.AsString);
     ftAutoInc, ftSmallint, ftWord:
       CellWord(vCol,vRow,field.AsInteger);
     ftInteger:
      begin
        CellDouble(vCol,vRow,field.AsFloat) ;// howard
       {
       if field.AsInteger>=65536 then
          CellDouble(vCol,vRow,field.AsFloat)
       else
          CellWord(vCol,vRow,field.AsInteger);
       }
      end;
     ftFloat, ftBCD:
       CellDouble(vCol,vRow,field.AsFloat);
  else
       //Cellstr(vCol,vRow,EmptyStr);
       Cellstr(vCol,vRow,field.AsString);
  end;
end;

function ParseStrings(Strs, sI_Sep: string): TStringList;
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


end.
