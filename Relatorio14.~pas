unit Relatorio14;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, StdCtrls, Grids, DBGrids, Buttons;

type
  TfrmRelatorio14 = class(TForm)
    btnImprimir: TSpeedButton;
    btnExportar: TSpeedButton;
    Label1: TLabel;
    lblValor: TLabel;
    lblRValor: TLabel;
    btnGerar: TSpeedButton;
    dbgDados: TDBGrid;
    GroupBox1: TGroupBox;
    lblDia: TLabel;
    lblOperador: TLabel;
    edtQtd: TEdit;
    edtMes: TEdit;
    qryDados: TADOQuery;
    dsDados: TDataSource;
    SaveDialog1: TSaveDialog;
    qryExec: TADOQuery;
    Label2: TLabel;
    EdtMotivo: TEdit;
    procedure edtQtdKeyPress(Sender: TObject; var Key: Char);
    procedure edtMesKeyPress(Sender: TObject; var Key: Char);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure btnGerarClick(Sender: TObject);
    procedure btnExportarClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmRelatorio14: TfrmRelatorio14;
  parametro: String;
implementation

{$R *.dfm}
uses Principal, DateUtils, Math;
procedure TfrmRelatorio14.edtQtdKeyPress(Sender: TObject; var Key: Char);
begin
    if not (key in ['0'..'9',Chr(8), Chr(3), Chr(22), Chr(24), Chr(44)]) then
      key := #0;
end;

procedure TfrmRelatorio14.edtMesKeyPress(Sender: TObject; var Key: Char);
begin
    if not (key in ['0'..'9',Chr(8), Chr(3), Chr(22), Chr(24), Chr(44)]) then
      key := #0;
end;

procedure TfrmRelatorio14.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
qryDados.Close;
qryExec.Close;
end;

procedure TfrmRelatorio14.FormShow(Sender: TObject);
begin
  if frmPrincipal.lblImprime.Caption = 'T' then //Habilita botão Imprimir
  begin
  btnImprimir.Enabled := True
  end else
  begin
  btnImprimir.Enabled := False;
  end;
  
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

  edtQtd.Text := '';
  edtMes.Text := '';
  EdtMotivo.Text := '';
  lblRValor.Caption := '0';

end;

procedure TfrmRelatorio14.btnGerarClick(Sender: TObject);
begin
  parametro := '';
  if frmPrincipal.lblCodRelatorio.Caption = '53' then
  begin
    if (EdtMotivo.Text <> '') and (edtQtd.Text <> '') or (edtMes.Text <> '') then
    begin
      parametro := 'MotivoCanc: '+EdtMotivo.Text + ' Qtd: ' + edtQtd.Text + ' Periodo: ' + edtMes.Text;
      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);
      SQL.Text := StringReplace(SQL.Text,'&MOTIVO_CANC',EdtMotivo.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&QUANTIDADE',edtQtd.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&INTERVALO_MES',edtMes.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      Open;
      end;

      lblRValor.Caption := IntToStr(qryDados.RecordCount);

      with qryExec do // INSERIR LOG DE VISUALIZAÇÃO
      begin
      Close;
      SQL.Clear;
      SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, PARAMETRO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'V'+#39+', '+#39+parametro+#39+', SYSDATE)');
      ExecSQL;
      end;
    end else
    begin
      ShowMessage('Favor informar todos os parâmetros.');
    end;
  end;

end;

procedure TfrmRelatorio14.btnExportarClick(Sender: TObject);
var
F: TextFile;
CodSeg, Nome, CPF, Tipo, CodVend, Vendedor, Situacao, DtIncl, DtCanc, Plano, Periodo, Descricao, cLinha  : String;
begin
 if qryDados.RecordCount > 0 then
 begin
  if frmPrincipal.lblCodRelatorio.Caption = '53' then
  begin
  if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_SEG' +';'+ 'NOME' +';'+ 'CPF' +';'+ 'TIPO' +';'+ 'COD_VEND' +';'+ 'CORRETOR' +';'+ 'SITUACAO' +';'+ 'DT_INCL' +';'+ 'DT_CANC' +';'+ 'PLANO' +';'+ 'PERIODO' +';'+ 'DESCRICAO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          CodSeg           := qryDados.FIELDBYNAME('COD_SEG').AsString;
          Nome             := qryDados.FIELDBYNAME('NOME').AsString;
          CPF              := qryDados.FIELDBYNAME('CPF').AsString;
          Tipo             := qryDados.FIELDBYNAME('TIPO').AsString;
          CodVend          := qryDados.FIELDBYNAME('COD_VEND').AsString;
          Vendedor         := qryDados.FIELDBYNAME('CORRETOR').AsString;
          Situacao         := qryDados.FIELDBYNAME('SITUACAO').AsString;
          DtIncl           := qryDados.FIELDBYNAME('DT_INCL').AsString;
          DtCanc           := qryDados.FIELDBYNAME('DT_CANC').AsString;
          Plano            := qryDados.FIELDBYNAME('PLANO').AsString;
          Periodo          := qryDados.FIELDBYNAME('PERIODO').AsString;
          Descricao        := qryDados.FIELDBYNAME('DESCRICAO').AsString;


          cLinha := '';
          cLinha := cLinha + CodSeg +';';
          cLinha := cLinha + Nome +';';
          cLinha := cLinha + CPF +';';
          cLinha := cLinha + Tipo +';';
          cLinha := cLinha + CodVend +';';
          cLinha := cLinha + Vendedor +';';
          cLinha := cLinha + Situacao +';';
          cLinha := cLinha + DtIncl +';';
          cLinha := cLinha + DtCanc +';';
          cLinha := cLinha + Plano +';';
          cLinha := cLinha + Periodo +';';
          cLinha := cLinha + Descricao +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
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
