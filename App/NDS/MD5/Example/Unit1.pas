unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  TDllFun = function ( input    :PChar;
                       input_len:UInt; //input_len:UInt;
                       key      :PChar;
                       key_len  :UInt; //key_len:UInt;
                       output   :PChar):UInt;stdcall;

  TForm1 = class(TForm)
    Button1: TButton;
    Memo1: TMemo;
    Label1: TLabel;
    Button3: TButton;
    procedure Button1Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

uses dllNdsU;

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
var
    jj, nL_Tmp : Integer;
    sL_Result, sL_TmpHex, sL_Data, sL_Key, sL_Output : String;
begin

    Memo1.Lines.Clear;

    sL_Data := '0001M0035010001000119940206125158ER01';
    sL_Key := '12345678';
    sL_Output := '';
    myNDS_RealMD5_8(sL_Data, IntToStr(length(sL_Data)), sL_Key, IntToStr(length(sL_Key)), sL_Output);
    Memo1.Lines.Add(sL_Data + ' => ' + sL_Output);

{
    Memo1.Lines.Clear;

    sL_Data := '0001M0035010001000119940206125158ER01';
    sL_Key := '12345678';
    sL_Output := '';
    myNDS_RealMD5_8(sL_Data, IntToStr(length(sL_Data)), sL_Key, IntToStr(length(sL_Key)), sL_Output);
    Memo1.Lines.Add(sL_Data + ' => ' + sL_Output);


    sL_Data := '0001M0032010002000119950906114824O';
    sL_Key := '12345678';
    sL_Output := '';
    myNDS_RealMD5_8(sL_Data, IntToStr(length(sL_Data)), sL_Key, IntToStr(length(sL_Key)), sL_Output);
    Memo1.Lines.Add(sL_Data + ' => ' + sL_Output);

    sL_Data := '0001M0035010002000119950906114844ES04';
    sL_Key := '12345678';
    sL_Output := '';
    myNDS_RealMD5_8(sL_Data, IntToStr(length(sL_Data)), sL_Key, IntToStr(length(sL_Key)), sL_Output);
    Memo1.Lines.Add(sL_Data + ' => ' + sL_Output);

    sL_Data := '0001M0035010002000119950906114904ES04';
    sL_Key := '12345678';
    sL_Output := '';
    myNDS_RealMD5_8(sL_Data, IntToStr(length(sL_Data)), sL_Key, IntToStr(length(sL_Key)), sL_Output);
    Memo1.Lines.Add(sL_Data + ' => ' + sL_Output);


    sL_Data := '0003S004B0000020100200306021351510001ZZN98760054D2A074D1001';
    sL_Key := '12345678';
    sL_Output := '';
    myNDS_RealMD5_8(sL_Data, IntToStr(length(sL_Data)), sL_Key, IntToStr(length(sL_Key)), sL_Output);
    Memo1.Lines.Add(sL_Data + ' => ' + sL_Output);
}
end;

procedure TForm1.Button3Click(Sender: TObject);
var
    nL_DllResult : UInt;
    ii, jj : Integer;
    nL_DataLen, nL_KeyLen, nL_Tmp : UInt;
    sL_Result, sL_TmpHex, sL_Data, sL_Key : String;

    aL_Key: array[0..8] of Char;
    aL_Data: array[0..1024] of Char;
    aL_Output: array[0..8] of Char;

    L_DllInstance : THandle;
//int NDS_RealMD5_8(unsigned char *input,unsigned short input_len,unsigned char *key,unsigned short key_len, unsigned char *out)

    MD5_Fun:TDllFun;
begin
    FillChar(aL_Data, high(aL_Data),char(0));
    FillChar(aL_Key, high(aL_Key),char(0));


    FillChar(aL_Output, high(aL_Output),char(0));

//    sL_Data := '0001M0035010001000119940206125158ER01'; //8F772E8552CF7A19
    sL_Data := '0003M004B0000020100200306021747430001ZZN98760054D2A074D1001';
    {
    sL_Data := '0001M0032010002000119950906114824O'; //36D69337D720F381
    sL_Data := '0001M0035010002000119950906114844ES04'; //5B387EAD17A0A328
    sL_Data := '0001M0035010002000119950906114904ES04'; //26399514C5767C57
    }
//    sL_Key := '12345678';
    sL_Key := '00000000';


    nL_DataLen := length(sL_Data);
    nL_KeyLen := length(sL_Key);

    StrPCopy(aL_Key, sL_Key);
    StrPCopy(aL_Data, sL_Data);

    L_DllInstance := LoadLibrary('prjNdsUtils.dll');
    if L_DllInstance = 0 then
      raise exception.Create('load dll error');
    @MD5_Fun := GetProcAddress(L_DllInstance,'_NDS_RealMD5_8');

    if @MD5_Fun = nil then
//      showmessage('function is nil');
      raise exception.Create('function is nil'#13'½Ğ½T©w function name');
    try
      nL_DllResult := MD5_Fun(aL_Data,nL_DataLen,aL_Key,nL_KeyLen,aL_Output);
//      showmessage(inttostr(nL_DllResult));
    except
      showmessage('exception..');
      exit;
    end;
    sL_Result := '';
    for jj:=0 to 7 do
    begin

      nL_Tmp := UInt(aL_Output[jj]);
      sL_TmpHex := IntToHex(nL_Tmp,2);
      sL_Result := sL_Result + sL_TmpHex;

    end;
    FreeLibrary(L_DllInstance);
    Memo1.Lines.Clear;
    Memo1.Lines.Add(sL_Data + ' => ' + sL_Result);    

end;

end.
