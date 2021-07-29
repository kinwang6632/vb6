program Emc_Separate;

uses
  Forms,
  Unit1 in 'Unit1.pas' {frmMain},
  Emc_Separate_TLB in 'Emc_Separate_TLB.pas',
  Rm in 'Rm.pas' {Emc_Separate1: TRemoteDataModule} {Emc_Separate1: CoClass};

{$R *.TLB}

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
