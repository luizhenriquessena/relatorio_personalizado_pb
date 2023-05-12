program RelatPerson;

uses
  Forms,
  Windows,
  IniFiles,
  SysUtils,
  Principal in 'Principal.pas' {frmPrincipal},
  DataModule in 'DataModule.pas' {DM: TDataModule},
  Relatorio1 in 'Relatorio1.pas' {frmRelatorio1},
  Relatorio2 in 'Relatorio2.pas' {frmRelatorio2},
  Relatorio3 in 'Relatorio3.pas' {frmRelatorio3},
  Relatorio4 in 'Relatorio4.pas' {frmRelatorio4},
  Relatorio5 in 'Relatorio5.pas' {frmRelatorio5},
  Relatorio6 in 'Relatorio6.pas' {frmRelatorio6},
  Relatorio7 in 'Relatorio7.pas' {frmRelatorio7},
  Relatorio8 in 'Relatorio8.pas' {frmRelatorio8},
  Relatorio9 in 'Relatorio9.pas' {frmRelatorio9},
  Relatorio10 in 'Relatorio10.pas' {frmRelatorio10},
  Relatorio11 in 'Relatorio11.pas' {frmRelatorio11},
  Relatorio12 in 'Relatorio12.pas' {frmRelatorio12},
  Relatorio13 in 'Relatorio13.pas' {frmRelatorio13},
  Relatorio14 in 'Relatorio14.pas' {frmRelatorio14},
  Impressao in 'Impressao.pas' {qrpimpressao60},
  Relat60 in 'Relat60.pas' {qrpRelat60: TQuickRep},
  Relat59 in 'Relat59.pas' {qrpRelat59: TQuickRep},
  Relat65 in 'Relat65.pas' {qrpRelat65: TQuickRep},
  Relatorio15 in 'Relatorio15.pas' {frmRelatorio15},
  Relatorio16 in 'Relatorio16.pas' {frmRelatorio16},
  Relatorio17 in 'Relatorio17.pas' {frmRelatorio17},
  Relat116 in 'Relat116.pas' {qrpRelat116: TQuickRep},
  Aguarde in 'Aguarde.pas' {frmAguarde},
  Relatorio18 in 'Relatorio18.pas' {frmRelatorio18},
  Relat100 in 'Relat100.pas' {qrpRelat100: TQuickRep},
  Relatorio19 in 'Relatorio19.pas' {frmRelatorio19},
  Relatorio20 in 'Relatorio20.pas' {frmRelatorio20},
  Rel136e137 in 'Rel136e137.pas' {qrpRelat136e137: TQuickRep};

{$R *.res}
 begin
  Application.Initialize;
  Application.CreateForm(TfrmPrincipal, frmPrincipal);
  Application.CreateForm(TDM, DM);
  Application.CreateForm(TfrmRelatorio1, frmRelatorio1);
  Application.CreateForm(TfrmRelatorio2, frmRelatorio2);
  Application.CreateForm(TfrmRelatorio3, frmRelatorio3);
  Application.CreateForm(TfrmRelatorio4, frmRelatorio4);
  Application.CreateForm(TfrmRelatorio5, frmRelatorio5);
  Application.CreateForm(TfrmRelatorio6, frmRelatorio6);
  Application.CreateForm(TfrmRelatorio7, frmRelatorio7);
  Application.CreateForm(TfrmRelatorio8, frmRelatorio8);
  Application.CreateForm(TfrmRelatorio9, frmRelatorio9);
  Application.CreateForm(TfrmRelatorio10, frmRelatorio10);
  Application.CreateForm(TfrmRelatorio11, frmRelatorio11);
  Application.CreateForm(TfrmRelatorio12, frmRelatorio12);
  Application.CreateForm(TfrmRelatorio13, frmRelatorio13);
  Application.CreateForm(TfrmRelatorio14, frmRelatorio14);
  Application.CreateForm(Tqrpimpressao60, qrpimpressao60);
  Application.CreateForm(TqrpRelat60, qrpRelat60);
  Application.CreateForm(TqrpRelat59, qrpRelat59);
  Application.CreateForm(TqrpRelat65, qrpRelat65);
  Application.CreateForm(TfrmRelatorio15, frmRelatorio15);
  Application.CreateForm(TfrmRelatorio16, frmRelatorio16);
  Application.CreateForm(TfrmRelatorio17, frmRelatorio17);
  Application.CreateForm(TqrpRelat116, qrpRelat116);
  Application.CreateForm(TfrmAguarde, frmAguarde);
  Application.CreateForm(TfrmRelatorio18, frmRelatorio18);
  Application.CreateForm(TqrpRelat100, qrpRelat100);
  Application.CreateForm(TfrmRelatorio19, frmRelatorio19);
  Application.CreateForm(TfrmRelatorio20, frmRelatorio20);
  Application.CreateForm(TqrpRelat136e137, qrpRelat136e137);
  Application.Run;
end.

