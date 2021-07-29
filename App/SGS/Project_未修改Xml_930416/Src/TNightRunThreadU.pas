unit TNightRunThreadU;

interface

uses
  Classes;

type
  TNightRunThread = class(TThread)
  private
    { Private declarations }
  protected
    procedure Execute; override;
  end;

implementation

{ Important: Methods and properties of objects in VCL or CLX can only be used
  in a method called using Synchronize, for example,

      Synchronize(UpdateCaption);

  and UpdateCaption could look like,

    procedure TNightRunThread.UpdateCaption;
    begin
      Form1.Caption := 'Updated in a thread';
    end; }

{ TNightRunThread }

procedure TNightRunThread.Execute;
begin
  { Place thread code here }
end;

end.
