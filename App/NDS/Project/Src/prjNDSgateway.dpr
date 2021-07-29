program prjNDSgateway;

uses
  Forms,
  frmMainU in 'frmMainU.pas' {frmMain},
  Ustru in '..\Common\Ustru.pas',
  Ufileu in '..\Common\Ufileu.pas',
  UdateTimeu in '..\Common\UdateTimeu.pas',
  ConstU in '..\Common\ConstU.pas',
  prjNDSgateway_TLB in 'prjNDSgateway_TLB.pas',
  NDSGatewayImpU in 'NDSGatewayImpU.pas' {NDSGateway: CoClass},
  dtmMainU in 'dtmMainU.pas' {dtmMain: TDataModule};

{$R *.TLB}

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfrmMain, frmMain);
  Application.CreateForm(TdtmMain, dtmMain);
  Application.Run;
end.
