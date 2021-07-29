unit Ufileu;

interface

Type
  TUfile = class
  public
    class function GetFileDir(sI_FileName, sI_Sep:String): string;
  end;

implementation

{ TUfile }

class function TUfile.GetFileDir(sI_FileName, sI_Sep: String): string;
var
    ii : Integer;
    sL_Result : String;
begin

    for ii:=length(sI_FileName) downto 0 do
    begin
      if (sI_FileName[ii]=sI_Sep) then
      begin
        sL_Result := copy(sI_FileName,0, ii);
        break;
      end;
    end;

    result := sL_Result;
end;

end.
