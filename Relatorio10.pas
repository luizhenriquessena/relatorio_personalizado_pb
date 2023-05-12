unit Relatorio10;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, Grids, DBGrids, StdCtrls, Buttons;

type
  TfrmRelatorio10 = class(TForm)
    Label1: TLabel;
    btnExportar: TSpeedButton;
    btnImprimir: TSpeedButton;
    lblTotal: TLabel;
    lblRTotal: TLabel;
    GroupBox1: TGroupBox;
    lblDia01: TLabel;
    btnGerar: TSpeedButton;
    edtDia: TEdit;
    dbgDados: TDBGrid;
    dsDados: TDataSource;
    qryDados: TADOQuery;
    SaveDialog1: TSaveDialog;
    qryExec: TADOQuery;
    edtDia2: TEdit;
    lblDia02: TLabel;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure edtDiaKeyPress(Sender: TObject; var Key: Char);
    procedure edtDia2KeyPress(Sender: TObject; var Key: Char);
    procedure btnGerarClick(Sender: TObject);
    procedure btnExportarClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmRelatorio10: TfrmRelatorio10;
  parametro: String;
implementation

{$R *.dfm}
uses MaskUtils, Math, Principal;
procedure TfrmRelatorio10.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
qryDados.Close;
end;

procedure TfrmRelatorio10.FormShow(Sender: TObject);
begin
  if frmPrincipal.lblExporta.Caption = 'T' then //Habilita botão Exportar
  begin
  btnExportar.Enabled := True;
  end else
  begin
  btnExportar.Enabled := False;
  end;

  with qryExec do
  begin
  Close;
  SQL.Clear;
  SQL.Add('ALTER SESSION SET NLS_DATE_FORMAT = '+#39+'DD/MM/YYYY HH24:MI:SS'+#39);
  ExecSQL;
  end;
end;

procedure TfrmRelatorio10.edtDiaKeyPress(Sender: TObject; var Key: Char);
begin
    if not (key in ['0'..'9',Chr(8), Chr(3), Chr(22), Chr(24), Chr(44)]) then
      key := #0;
end;

procedure TfrmRelatorio10.edtDia2KeyPress(Sender: TObject; var Key: Char);
begin
    if not (key in ['0'..'9',Chr(8), Chr(3), Chr(22), Chr(24), Chr(44)]) then
      key := #0;
end;

procedure TfrmRelatorio10.btnGerarClick(Sender: TObject);
begin
  parametro := '';
  if edtDia.Text <> '' then
  begin
    parametro := 'Dia01: '+ edtDia.Text + ' Dia02: ' + edtDia2.Text;

    with qryDados do
    begin
    Close;
    SQL.Clear;
    SQL.Add(frmPrincipal.memSQL.Text);
    SQL.Text := StringReplace(SQL.Text,'&DIA1',edtDia.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
    SQL.Text := StringReplace(SQL.Text,'&DIA2',edtDia2.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
    Open;
    end;

    lblRTotal.Caption := IntToStr(qryDados.RecordCount);

    with qryExec do // INSERIR LOG DE VISUALIZAÇÃO
    begin
    Close;
    SQL.Clear;
    SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, PARAMETRO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'V'+#39+','+#39+parametro+#39+', SYSDATE)');
    ExecSQL;
    end;
  end else
  begin
    ShowMessage('Parâmetro (Dia) não informado.');
  end;
end;

procedure TfrmRelatorio10.btnExportarClick(Sender: TObject);
var
F: TextFile;
Codigo, Titular, Responsavel, Celular, CelResp, Telefone, Plano, cLinha : String;
begin
 if qryDados.RecordCount > 0 then
 begin
  if SaveDialog1.Execute then
  begin
    AssignFile(F, SaveDialog1.FileName);
    Rewrite (F);
    cLinha := '';
    Writeln (F,cLinha + 'CODIGO' +';'+ 'TITULAR' +';'+ 'RESPONSAVEL' +';'+ 'CELULAR' +';'+ 'CELRESP' +';'+ 'TELEFONE' +';'+ 'PLANO');
    qryDados.First;
    while not qryDados.Eof do
      begin
        Codigo               := qryDados.FIELDBYNAME('CODIGO').AsString;
        Titular              := qryDados.FIELDBYNAME('TITULAR').AsString;
        Responsavel          := qryDados.FIELDBYNAME('RESPONSAVEL').AsString;
        Celular              := qryDados.FIELDBYNAME('CELULAR').AsString;
        CelResp              := qryDados.FIELDBYNAME('CELRESP').AsString;
        Telefone             := qryDados.FIELDBYNAME('TELEFONE').AsString;
        Plano                := qryDados.FIELDBYNAME('PLANO').AsString;

        cLinha := '';
        cLinha := cLinha + Codigo +';';
        cLinha := cLinha + Titular +';';
        cLinha := cLinha + Responsavel +';';
        cLinha := cLinha + Celular +';';
        cLinha := cLinha + CelResp +';';
        cLinha := cLinha + Telefone +';';
        cLinha := cLinha + Plano +';';

        Writeln(F,cLinha);
        qryDados.Next;
      end;
    CloseFile (F);
    ShowMessage('Arquivo Gerado com sucesso.');
  end;

  with qryExec do // INSERIR LOG DE EXPORTAÇÃO
  begin
  Close;
  SQL.Clear;
  SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, PARAMETRO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'E'+#39+', '+#39+parametro+#39+', SYSDATE)');
  ExecSQL;
  end;
 end;
end;

end.
