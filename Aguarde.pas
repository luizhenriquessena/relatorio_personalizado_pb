unit Aguarde;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls;

type
  TfrmAguarde = class(TForm)
    ProgressBar1: TProgressBar;
    lblTexto: TLabel;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmAguarde: TfrmAguarde;

implementation

{$R *.dfm}

end.
