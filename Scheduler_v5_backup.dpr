program Scheduler_v5_backup;

uses
  Vcl.Forms,
  uToSiteMain in 'uToSiteMain.pas' {fmToSiteMain},
  uMain in 'uMain.pas' {fmMain},
  uTodo in 'uTodo.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.Title := 'USU.kz - Обмен информацией с сайтом';
  Application.CreateForm(TfmToSiteMain, fmToSiteMain);
  Application.CreateForm(TfmMain, fmMain);
  Application.Run;
end.
