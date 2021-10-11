program Devices;

uses
  Forms,
  Main in 'Main.pas' {Form1};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'Lister les périphériques du système (WinXP Only)';
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
