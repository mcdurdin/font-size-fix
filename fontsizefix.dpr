program fontsizefix;

uses
  Forms,
  frmFontsizefix in 'frmFontsizefix.pas' {Form1};

{$R *.res}
{$R manifest.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
