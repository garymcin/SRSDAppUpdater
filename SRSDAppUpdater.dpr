program SRSDAppUpdater;

uses
  Vcl.Forms,
  UpdateProgram in 'UpdateProgram.pas' {frm_MainUpdate};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(Tfrm_MainUpdate, frm_MainUpdate);
  Application.Run;
end.
