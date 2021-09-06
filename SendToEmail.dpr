program SendToEmail;

uses
  Vcl.Forms,
  USendToEmail in 'USendToEmail.pas' {Form9},
  Vcl.Themes,
  Vcl.Styles,
  UDm in 'UDm.pas' {DM: TDataModule};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  TStyleManager.TrySetStyle('Lavender Classico');
  Application.CreateForm(TDM, DM);
  Application.CreateForm(TForm9, Form9);
  Application.Run;
end.
