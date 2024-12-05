unit uMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs;

type
  TfmMain = class(TForm)
  private
  public
    function GetLangValue(sCat, sPar: string): string;
  end;

var
  fmMain: TfmMain;

implementation

{$R *.dfm}

{ TfmMain }

function TfmMain.GetLangValue(sCat, sPar: string): string;
begin
  Result := '';
end;

end.
