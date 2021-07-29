program prjSO1621L;

uses
  Forms,
  Windows,
  cbMain in 'cbMain.pas' {fmMain},
  cbDataController in 'cbDataController.pas' {DBController: TDataModule},
  cbDvnThread in 'cbDvnThread.pas',
  cbMultiSelect in 'cbMultiSelect.pas' {fmMultiSelect},
  cbNagraThread in 'cbNagraThread.pas';

{$R *.res}

var
  aWnd, aWnd2, aDelphiWnd: HWND;
begin
  aWnd := FindWindow( PChar( 'TfmMain' ), 'DVN配對燒機作業[SO1621L]' );
  aDelphiWnd := FindWindow( PChar( 'TAppBuilder' ), nil );
  if ( aWnd = 0 ) or ( aDelphiWnd > 0 ) then
  begin
    Application.Initialize;
    Application.Title := 'DVN配對燒機作業[SO1621L]';
    Application.CreateForm(TfmMain, fmMain);
  Application.CreateForm(TfmMultiSelect, fmMultiSelect);
  Application.Run;
  end else
  begin
   aWnd2 := FindWindow( PChar( 'TApplication' ), 'DVN配對燒機作業[SO1621L]' );
   if ( aWnd2 <> 0 ) then
   begin
     if IsIconic( aWnd2 ) then
       ShowWindow( aWnd2, SW_SHOWNORMAL );
   end;
   SetForegroundWindow( aWnd );
   SetActiveWindow( aWnd );
   Application.ShowMainForm := False;
   Application.Terminate;
  end;
end.

