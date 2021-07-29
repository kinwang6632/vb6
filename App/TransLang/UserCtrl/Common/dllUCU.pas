unit dllUCU;

interface

  procedure GetCurUser(var UserId, UserName, Password, UserGrp: string);

implementation

procedure dllGetCurUser(UserId, UserName, Password,
  UserGrp: PChar); stdcall; external 'UC.DLL';

procedure GetCurUser(var UserId, UserName, Password, UserGrp: string);
var
  cUserId, cUserName, cPassword, cUserGrp: array[0..20] of Char;
begin
  dllGetCurUser(cUserId, cUserName, cPassword, cUserGrp);
  UserId := cUserId;
  UserName := cUserName;
  Password := cPassword;
  UserGrp := cUserGrp;
end;

end.
