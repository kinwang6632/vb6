unit fraDBLookupCobU;

interface

uses 
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, DBCtrls;

type
  TfraDBLookupCob = class(TFrame)
    DBLookupComboBox1: TDBLookupComboBox;
    CheckBox1: TCheckBox;
    procedure CheckBox1Click(Sender: TObject);
  private
    { Private declarations }
    function ParseStrings(Strs:string; Sep:Char) : TStringList;     
  public
    { Public declarations }
  end;

implementation

{$R *.DFM}

procedure TfraDBLookupCob.CheckBox1Click(Sender: TObject);
var
    L_StrList : TStringList;
    sL_Field1, sL_Field2 : String;

begin
    L_StrList := ParseStrings(DBLookupComboBox1.ListField,';');
    sL_Field1 := L_StrList.Strings[0];
    sL_Field2 := L_StrList.Strings[1];
    
    DBLookupComboBox1.ListField := sL_Field2 + ';' + sL_Field1;
end;

function TfraDBLookupCob.ParseStrings(Strs: string;
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

end.
