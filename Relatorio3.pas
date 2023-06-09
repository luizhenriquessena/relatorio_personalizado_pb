unit Relatorio3;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, Grids, DBGrids, Buttons, StdCtrls;

type
  TfrmRelatorio3 = class(TForm)
    Label1: TLabel;
    btnExportar: TSpeedButton;
    lblTotal: TLabel;
    lblRTotal: TLabel;
    dbgDados: TDBGrid;
    dsDados: TDataSource;
    SaveDialog1: TSaveDialog;
    btnGerar: TSpeedButton;
    qryExec: TADOQuery;
    qryDados: TADOQuery;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure btnGerarClick(Sender: TObject);
    procedure btnExportarClick(Sender: TObject);
    procedure qryDadosAfterOpen(DataSet: TDataSet);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmRelatorio3: TfrmRelatorio3;

implementation

{$R *.dfm}
uses Principal, DateUtils, DataModule;
procedure TfrmRelatorio3.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
qryDados.Close;
end;

procedure TfrmRelatorio3.FormShow(Sender: TObject);
begin
  if frmPrincipal.lblExporta.Caption = 'T' then //Habilita bot�o Exportar
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

procedure TfrmRelatorio3.btnGerarClick(Sender: TObject);
begin
  with qryDados do
  begin
  Close;
  SQL.Clear;
  SQL.Add(frmPrincipal.memSQL.Text);
  Open;
  end;
  frmPrincipal.DimensionarGrid(dbgDados,Self);
  lblRTotal.Caption := IntToStr(qryDados.RecordCount);

  with qryExec do // INSERIR LOG DE VISUALIZA��O
  begin
  Close;
  SQL.Clear;
  SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'V'+#39+', SYSDATE)');
  ExecSQL;
  end;
end;

procedure TfrmRelatorio3.btnExportarClick(Sender: TObject);
var
F: TextFile;
P1, P2, P3, P4, P5, P6, P7, P8, P9, P10, P11, P12, P13, P14, P15, P16, P17, P18, P19,
P20, P21, P22, P23, P24, P25, cLinha : String;
begin
 if qryDados.RecordCount > 0 then
 begin

   if frmPrincipal.lblCodRelatorio.Caption = '5' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'CODIGO' +';'+ 'RAZ_SOC' +';'+ 'FANTASIA' +';'+ 'ESPECIALIDADE' +';'+ 'RUA' +';'+ 'BAIRRO' +';'+ 'CIDADE' +';'+ 'ESTADO' +';'+ 'CEP' +';'+ 'CONVENIADO' +';'+ 'CANCELADO' +';'+ 'EMAIL' +';'+ 'TELEDONE' +';'+ 'TD_INCL' +';'+ 'DT_INI_PRESTACAO' +';'+ 'CGC' +';'+ 'CNES' +';'+ 'TIPO_ESTABELECIMENTO' +';'+ 'REGIME_ATENDIMENTO' +';'+ 'REDE' +';'+ 'END_EXIBE' +';'+ 'ESPEC_EXIBE' +';'+ 'TIPO_PREST');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('CODIGO').AsString;
          P2  := qryDados.FIELDBYNAME('RAZ_SOC').AsString;
          P3  := qryDados.FIELDBYNAME('FANTASIA').AsString;
          P4  := qryDados.FIELDBYNAME('ESPECIALIDADE').AsString;
          P5  := qryDados.FIELDBYNAME('RUA').AsString;
          P6  := qryDados.FIELDBYNAME('BAIRRO').AsString;
          P7  := qryDados.FIELDBYNAME('CIDADE').AsString;
          P8  := qryDados.FIELDBYNAME('ESTADO').AsString;
          P9  := qryDados.FIELDBYNAME('CEP').AsString;
          P10 := qryDados.FIELDBYNAME('CONVENIADO').AsString;
          P11 := qryDados.FIELDBYNAME('CANCELADO').AsString;
          P12 := qryDados.FIELDBYNAME('EMAIL').AsString;
          P13 := qryDados.FIELDBYNAME('TELEFONE').AsString;
          P14 := qryDados.FIELDBYNAME('DT_INCL').AsString;
          P15 := qryDados.FIELDBYNAME('DT_INI_PRESTACAO').AsString;
          P16 := qryDados.FIELDBYNAME('CGC').AsString;
          P17 := qryDados.FIELDBYNAME('CNES').AsString;
          P18 := qryDados.FIELDBYNAME('TIPO_ESTABELECIMENTO').AsString;
          P19 := qryDados.FIELDBYNAME('REGIME_ATENDIMENTO').AsString;
          P20 := qryDados.FIELDBYNAME('REDE').AsString;
          P21 := qryDados.FIELDBYNAME('END_EXIBE').AsString;
          P22 := qryDados.FIELDBYNAME('ESPEC_EXIBE').AsString;
          P23 := qryDados.FIELDBYNAME('TIPO_PREST').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';     
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';
          cLinha := cLinha + P10 +';';
          cLinha := cLinha + P11 +';';
          cLinha := cLinha + P12 +';';
          cLinha := cLinha + P13 +';';
          cLinha := cLinha + P14 +';';
          cLinha := cLinha + P15 +';';
          cLinha := cLinha + P16 +';';
          cLinha := cLinha + P17 +';';
          cLinha := cLinha + P18 +';';
          cLinha := cLinha + P19 +';';
          cLinha := cLinha + P20 +';';
          cLinha := cLinha + P21 +';';
          cLinha := cLinha + P22 +';';
          cLinha := cLinha + P23 +';';          

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '7' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'TIPO' +';'+ 'CODIGO' +';'+ 'RAZ_SOC' +';'+ 'DT_PREV_EXCL' +';'+ 'SITUACAO' +';'+ 'DT_CANCELAMENTO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('TIPO').AsString;
          P2  := qryDados.FIELDBYNAME('CODIGO').AsString;
          P3  := qryDados.FIELDBYNAME('RAZ_SOC').AsString;
          P4  := qryDados.FIELDBYNAME('DT_PREV_EXCL').AsString;
          P5  := qryDados.FIELDBYNAME('SITUACAO').AsString;
          P6  := qryDados.FIELDBYNAME('DT_CANCELAMENTO').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '8' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'TIPO' +';'+ 'CODIGO' +';'+ 'RAZ_SOC' +';'+ 'COD_ESPM' +';'+ 'ESPECIALIDADE');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('TIPO').AsString;
          P2  := qryDados.FIELDBYNAME('CODIGO').AsString;
          P3  := qryDados.FIELDBYNAME('RAZ_SOC').AsString;
          P4  := qryDados.FIELDBYNAME('COD_ESPM').AsString;
          P5  := qryDados.FIELDBYNAME('ESPECIALIDADE').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '9' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'TIPO' +';'+ 'CODIGO' +';'+ 'RAZ_SOC' +';'+ 'FANTASIA' +';'+ 'RANK' +';'+ 'COD_GRUPO' +';'+ 'GRUPO' +';'+ 'ID_TIPO_PREST' +';'+ 'TIPO_PREST' +';'+ 'REEMBOLSO' +';'+ 'CONVENIADO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('TIPO').AsString;
          P2  := qryDados.FIELDBYNAME('CODIGO').AsString;
          P3  := qryDados.FIELDBYNAME('RAZ_SOC').AsString;
          P4  := qryDados.FIELDBYNAME('FANTASIA').AsString;
          P5  := qryDados.FIELDBYNAME('RANK').AsString;
          P6  := qryDados.FIELDBYNAME('COD_GRU').AsString;
          P7  := qryDados.FIELDBYNAME('GRUPO').AsString;
          P8  := qryDados.FIELDBYNAME('ID_TIPO_PREST').AsString;
          P9  := qryDados.FIELDBYNAME('TIPO_PREST').AsString;
          P10  := qryDados.FIELDBYNAME('REEMBOLSO').AsString;
          P11  := qryDados.FIELDBYNAME('CONVENIADO').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';
          cLinha := cLinha + P10 +';';
          cLinha := cLinha + P11 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if (frmPrincipal.lblCodRelatorio.Caption = '21') or  (frmPrincipal.lblCodRelatorio.Caption = '33') then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_VEND' +';'+ 'NOME' +';'+ 'EMAIL' +';'+ 'MATRICULA' +';'+ 'CPF_CNPJ' +';'+ 'COD_BANCO' +';'+ 'AGENCIA' +';'+ 'CONTA' +';'+ 'DT_CAD' +';'+ 'ID_DT_NASC');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_VEND').AsString;
          P2  := qryDados.FIELDBYNAME('NOME').AsString;
          P3  := qryDados.FIELDBYNAME('EMAIL').AsString;
          P4  := qryDados.FIELDBYNAME('MATRICULA').AsString;
          P5  := qryDados.FIELDBYNAME('CPF_CNPJ').AsString;
          P6  := qryDados.FIELDBYNAME('COD_BANCO').AsString;
          P7  := qryDados.FIELDBYNAME('AGENCIA').AsString;
          P8  := qryDados.FIELDBYNAME('CONTA').AsString;
          P9  := qryDados.FIELDBYNAME('DT_CAD').AsString;
          P10 := qryDados.FIELDBYNAME('ID_DT_NASC').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';
          cLinha := cLinha + P10 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '25' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'CODIGO' +';'+ 'DESCRICAO' +';'+ 'TIPO_AUDITORIA' +';'+ 'PRAZO_AUDITORIA');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('CODIGO').AsString;
          P2  := qryDados.FIELDBYNAME('DESCRICAO').AsString;
          P3  := qryDados.FIELDBYNAME('TIPO_AUDITORIA').AsString;
          P4  := qryDados.FIELDBYNAME('PRAZO_AUDITORIA').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '34' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_SET' +';'+ 'DESCRICAO' +';'+ 'MES_REAJUSTE' +';'+ 'VIDAS' +';'+ 'PERIODO' +';'+ 'DT_VENC' +';'+ 'CONTRATO' +';'+ 'DATA_CADASTRO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_SET').AsString;
          P2  := qryDados.FIELDBYNAME('DESCRICAO').AsString;
          P3  := qryDados.FIELDBYNAME('MES_REAJUSTE').AsString;
          P4  := qryDados.FIELDBYNAME('VIDAS').AsString;
          P5  := qryDados.FIELDBYNAME('PERIODO').AsString;
          P6  := qryDados.FIELDBYNAME('DT_VENC').AsString;
          P7  := qryDados.FIELDBYNAME('CONTRATO').AsString;
          P8  := qryDados.FIELDBYNAME('DATA_CADASTRO').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';                    

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '36' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_SET' +';'+ 'DESCRICAO' +';'+ 'DT_CAD' +';'+ 'COD_BOL' +';'+ 'VALOR' +';'+ 'VALOR_PG' +';'+ 'PAGO' +';'+ 'QTD_BENEFICIARIO' +';'+ 'COD_PLANO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_SET').AsString;
          P2  := qryDados.FIELDBYNAME('DESCRICAO').AsString;
          P3  := qryDados.FIELDBYNAME('DT_CAD').AsString;
          P4  := qryDados.FIELDBYNAME('COD_BOL').AsString;
          P5  := qryDados.FIELDBYNAME('VALOR').AsString;
          P6  := qryDados.FIELDBYNAME('VALOR_PG').AsString;
          P7  := qryDados.FIELDBYNAME('PAGO').AsString;
          P8  := qryDados.FIELDBYNAME('QTD_BENEFICIARIO').AsString;
          P9  := qryDados.FIELDBYNAME('COD_PLANO').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '38' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_SET' +';'+ 'COD_SETSET' +';'+ 'DESCRICAO' +';'+ 'DT_CAD' +';'+ 'COD_BOL' +';'+ 'VALOR' +';'+ 'VALOR_PG' +';'+ 'PAGO' +';'+ 'QTD_BENEFICIARIO' +';'+ 'COD_PLANO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_SET').AsString;
          P2  := qryDados.FIELDBYNAME('COD_SETSET').AsString;
          P3  := qryDados.FIELDBYNAME('DESCRICAO').AsString;
          P4  := qryDados.FIELDBYNAME('DT_CAD').AsString;
          P5  := qryDados.FIELDBYNAME('COD_BOL').AsString;
          P6  := qryDados.FIELDBYNAME('VALOR').AsString;
          P7  := qryDados.FIELDBYNAME('VALOR_PG').AsString;
          P8  := qryDados.FIELDBYNAME('PAGO').AsString;
          P9  := qryDados.FIELDBYNAME('QTD_BENEFICIARIO').AsString;
          P10 := qryDados.FIELDBYNAME('COD_PLANO').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';
          cLinha := cLinha + P10 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '39' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'CODIGO' +';'+ 'DESCRICAO' +';'+ 'BOL_PAGO' +';'+ 'BOL_ULTIMO' +';'+ 'VIDAS');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('CODIGO').AsString;
          P2  := qryDados.FIELDBYNAME('DESCRICAO').AsString;
          P3  := qryDados.FIELDBYNAME('BOL_PAGO').AsString;
          P4  := qryDados.FIELDBYNAME('BOL_ULTIMO').AsString;
          P5  := qryDados.FIELDBYNAME('VIDAS').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '40' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_BOL' +';'+ 'COD_SET' +';'+ 'DESCRICAO' +';'+ 'CNPJ' +';'+ 'VALOR' +';'+ 'DT_VENC' +';'+ 'VIDAS');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_BOL').AsString;
          P2  := qryDados.FIELDBYNAME('COD_SET').AsString;
          P3  := qryDados.FIELDBYNAME('DESCRICAO').AsString;
          P4  := qryDados.FIELDBYNAME('CNPJ').AsString;
          P5  := qryDados.FIELDBYNAME('VALOR').AsString;
          P6  := qryDados.FIELDBYNAME('DT_VENC').AsString;
          P7  := qryDados.FIELDBYNAME('VIDAS').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '42' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_SET' +';'+ 'DESCRICAO' +';'+ 'SIGLA' +';'+ 'DT_CAD' +';'+ 'DT_CANC' +';'+ 'MES_CANC' +';'+ 'GRUPO' +';'+ 'COD_VEND' +';'+ 'NOME' +';'+ 'MOTIVO' +';'+ 'COD_PLANO' +';'+ 'PLANO' +';'+ 'VIDAS');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_SET').AsString;
          P2  := qryDados.FIELDBYNAME('DESCRICAO').AsString;
          P3  := qryDados.FIELDBYNAME('SIGLA').AsString;
          P4  := qryDados.FIELDBYNAME('DT_CAD').AsString;
          P5  := qryDados.FIELDBYNAME('DT_CANC').AsString;
          P6  := qryDados.FIELDBYNAME('MES_CANC').AsString;
          P7  := qryDados.FIELDBYNAME('GRUPO').AsString;
          P8  := qryDados.FIELDBYNAME('COD_VEND').AsString;
          P9  := qryDados.FIELDBYNAME('NOME').AsString;
          P10  := qryDados.FIELDBYNAME('MOTIVO').AsString;
          P11  := qryDados.FIELDBYNAME('COD_PLANO').AsString;
          P12  := qryDados.FIELDBYNAME('PLANO').AsString;
          P13  := qryDados.FIELDBYNAME('VIDAS').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';
          cLinha := cLinha + P10 +';';
          cLinha := cLinha + P11 +';';
          cLinha := cLinha + P12 +';';
          cLinha := cLinha + P13 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '46' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'TABELA' +';'+ 'COD_SERV' +';'+ 'DESCRICAO' +';'+ 'CONDICIONADO' +';'+ 'ROL' +';'+ 'PAC' +';'+ 'PERICIA' +';'+ 'LAUDO' +';'+ 'CARENCIA' +';'+ 'TIPO' +';'+ 'DUT');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('TABELA').AsString;
          P2  := qryDados.FIELDBYNAME('COD_SERV').AsString;
          P3  := qryDados.FIELDBYNAME('DESCRICAO').AsString;
          P4  := qryDados.FIELDBYNAME('CONDICIONADO').AsString;
          P5  := qryDados.FIELDBYNAME('ROL').AsString;
          P6  := qryDados.FIELDBYNAME('PAC').AsString;
          P7  := qryDados.FIELDBYNAME('PERICIA').AsString;
          P8  := qryDados.FIELDBYNAME('LAUDO').AsString;
          P9  := qryDados.FIELDBYNAME('CARENCIA').AsString;
          P10  := qryDados.FIELDBYNAME('TIPO').AsString;
          P11  := qryDados.FIELDBYNAME('DUT').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';
          cLinha := cLinha + P10 +';';
          cLinha := cLinha + P11 +';';          

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '57' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'ADESAO' +';'+ 'COD_SEG' +';'+ 'NOME' +';'+ 'TIPO' +';'+ 'FORMA' +';'+ 'CONTRATO' +';'+ 'COD_TITULAR' +';'+ 'CORRETOR' +';'+ 'SUBCORRETOR');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('ADESAO').AsString;
          P2  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P3  := qryDados.FIELDBYNAME('NOME').AsString;
          P4  := qryDados.FIELDBYNAME('TIPO').AsString;
          P5  := qryDados.FIELDBYNAME('FORMA').AsString;
          P6  := qryDados.FIELDBYNAME('CONTRATO').AsString;
          P7  := qryDados.FIELDBYNAME('COD_TITULAR').AsString;
          P8  := qryDados.FIELDBYNAME('CORRETOR').AsString;
          P9  := qryDados.FIELDBYNAME('SUBCORRETOR').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '58' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_SEG' +';'+ 'NOME' +';'+ 'TIPO' +';'+ 'DT_NASC' +';'+ 'SEXO' +';'+ 'CPF' +';'+ 'MAE' +';'+ 'EST_CIVIL' +';'+ 'ENDERECO' +';'+ 'COMP' +';'+ 'BAIRRO' +';'+ 'CIDADE' +';'+ 'ESTADO' +';'+ 'CEP' +';'+ 'EMAIL' +';'+ 'CANCELADO' +';'+ 'DT_CANC' +';'+ 'DESCRICAO' +';'+ 'CNPJ' +';'+ 'SIGLA' +';'+ 'COD_ADTV' +';'+ 'NOME_ADTV' +';'+ 'DT_INCL' +';'+ 'DT_INICIO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P2  := qryDados.FIELDBYNAME('NOME').AsString;
          P3  := qryDados.FIELDBYNAME('TIPO').AsString;
          P4  := qryDados.FIELDBYNAME('DT_NASC').AsString;
          P5  := qryDados.FIELDBYNAME('SEXO').AsString;
          P6  := qryDados.FIELDBYNAME('CPF').AsString;
          P7  := qryDados.FIELDBYNAME('MAE').AsString;
          P8  := qryDados.FIELDBYNAME('EST_CIVIL').AsString;
          P9  := qryDados.FIELDBYNAME('ENDERECO').AsString;
          P10 := qryDados.FIELDBYNAME('NUMERO').AsString;
          P11 := qryDados.FIELDBYNAME('COMP').AsString;
          P12 := qryDados.FIELDBYNAME('BAIRRO').AsString;
          P13 := qryDados.FIELDBYNAME('CIDADE').AsString;
          P14 := qryDados.FIELDBYNAME('ESTADO').AsString;
          P15 := qryDados.FIELDBYNAME('CEP').AsString;
          P16 := qryDados.FIELDBYNAME('EMAIL').AsString;
          P17 := qryDados.FIELDBYNAME('CANCELADO').AsString;
          P18 := qryDados.FIELDBYNAME('DT_CANC').AsString;
          P19 := qryDados.FIELDBYNAME('DESCRICAO').AsString;
          P20 := qryDados.FIELDBYNAME('CNPJ').AsString;
          P21 := qryDados.FIELDBYNAME('SIGLA').AsString;
          P22 := qryDados.FIELDBYNAME('COD_ADTV').AsString;
          P23 := qryDados.FIELDBYNAME('NOME_ADTV').AsString;
          P24 := qryDados.FIELDBYNAME('DT_INCL').AsString;
          P25 := qryDados.FIELDBYNAME('DT_INICIO').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';
          cLinha := cLinha + P10 +';';
          cLinha := cLinha + P11 +';';
          cLinha := cLinha + P12 +';';
          cLinha := cLinha + P13 +';';
          cLinha := cLinha + P14 +';';
          cLinha := cLinha + P15 +';';
          cLinha := cLinha + P16 +';';
          cLinha := cLinha + P17 +';';
          cLinha := cLinha + P18 +';';
          cLinha := cLinha + P19 +';';
          cLinha := cLinha + P20 +';';
          cLinha := cLinha + P21 +';';
          cLinha := cLinha + P22 +';';
          cLinha := cLinha + P23 +';';
          cLinha := cLinha + P24 +';';
          cLinha := cLinha + P25 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;


   if frmPrincipal.lblCodRelatorio.Caption = '76' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'CODIGO' +';'+ 'EMPRESA' +';'+ 'VENCIMENTO' +';'+ 'CNPJ');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_SET').AsString;
          P2  := qryDados.FIELDBYNAME('DESCRICAO').AsString;
          P3  := qryDados.FIELDBYNAME('DT_VENC').AsString;
          P4  := qryDados.FIELDBYNAME('CNPJ').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';          

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '79' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_SET' +';'+ 'DESCRICAO' +';'+ 'VIDAS' +';'+ 'COD_PLANO' +';'+ 'PLANO' +';'+ 'CADASTRO' +';'+ 'VENCIMENTO' +';'+ 'COD_VEND' +';'+ 'VENDEDOR' +';'+ 'GRUPO' +';'+ 'VALOR_ULTIMO_BOL');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_SET').AsString;
          P2  := qryDados.FIELDBYNAME('DESCRICAO').AsString;
          P3  := qryDados.FIELDBYNAME('VIDAS').AsString;
          P4  := qryDados.FIELDBYNAME('COD_PLANO').AsString;
          P5  := qryDados.FIELDBYNAME('NOME').AsString;
          P6  := qryDados.FIELDBYNAME('DT_CAD').AsString;
          P7  := qryDados.FIELDBYNAME('DT_VENC').AsString;
          P8  := qryDados.FIELDBYNAME('COD_VEND').AsString;
          P9  := qryDados.FIELDBYNAME('VENDEDOR').AsString;
          P10 := qryDados.FIELDBYNAME('GRUPO').AsString;
          P11 := qryDados.FIELDBYNAME('VALOR_ULTIMO_BOL').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';
          cLinha := cLinha + P10 +';';
          cLinha := cLinha + P11 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '101' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'NOME COMPLETO' +';'+ 'CPF' +';'+ 'DATA DE NASCIMENTO' +';'+ 'SEXO' +';'+ 'TIT_DEP' +';'+ 'APH/OMT' +';'+ 'TIPO_PLANO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('NOME').AsString;
          P2  := qryDados.FIELDBYNAME('CPF').AsString;
          P3  := qryDados.FIELDBYNAME('DT_NASC').AsString;
          P4  := qryDados.FIELDBYNAME('SEXO').AsString;
          P5  := qryDados.FIELDBYNAME('TIPO').AsString;
          P6  := qryDados.FIELDBYNAME('APH_OMT').AsString;
          P7  := qryDados.FIELDBYNAME('TIPO_PLANO').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '104' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_SERV' +';'+ 'DESCRICAO' +';'+ 'CARENCIA' +';'+ 'COPARTICIPACAO' +';'+ 'VLR_COPART');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_SERV').AsString;
          P2  := qryDados.FIELDBYNAME('DESCRICAO').AsString;
          P3  := qryDados.FIELDBYNAME('CARENCIA').AsString;
          P4  := qryDados.FIELDBYNAME('COPARTICIPACAO').AsString;
          P5  := qryDados.FIELDBYNAME('VLR_COPART').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '110' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_PLANO' +';'+ 'PLANO' +';'+ 'REGISTRO' +';'+ 'SITUACAO' +';'+ 'COBERTURA' +';'+ 'ABRANGENCIA' +';'+ 'NATUREZA' +';'+ 'COPARTICIPACAO' +';'+ 'COBRE_SERVICO' +';'+ 'INTERNACAO_ENFERMARIA' +';'+ 'INTERNACAO_APARTAMENTO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_PLANO').AsString;
          P2  := qryDados.FIELDBYNAME('PLANO').AsString;
          P3  := qryDados.FIELDBYNAME('REGISTRO').AsString;
          P4  := qryDados.FIELDBYNAME('SITUACAO').AsString;
          P5  := qryDados.FIELDBYNAME('COBERTURA').AsString;
          P6  := qryDados.FIELDBYNAME('ABRANGENCIA').AsString;
          P7  := qryDados.FIELDBYNAME('NATUREZA').AsString;
          P8  := qryDados.FIELDBYNAME('COPARTICIPACAO').AsString;
          P9  := qryDados.FIELDBYNAME('COBRE_SERVICOS').AsString;
          P10 := qryDados.FIELDBYNAME('INTERNACAO_ENFERMARIA').AsString;
          P11 := qryDados.FIELDBYNAME('INTERNACAO_APARTAMENTO').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';
          cLinha := cLinha + P10 +';';
          cLinha := cLinha + P11 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '113' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_SEG' +';'+ 'CPF' +';'+ 'NOME' +';'+ 'PLANO' +';'+ 'EMPRESA' +';'+ 'INCLUS�O');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P2  := qryDados.FIELDBYNAME('CPF').AsString;
          P3  := qryDados.FIELDBYNAME('NOME').AsString;
          P4  := qryDados.FIELDBYNAME('PLANO').AsString;
          P5  := qryDados.FIELDBYNAME('EMPRESA').AsString;
          P6  := qryDados.FIELDBYNAME('DT_INCL').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
   end;


   if (frmPrincipal.lblCodRelatorio.Caption = '139') or (frmPrincipal.lblCodRelatorio.Caption = '140') then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_VEND' +';'+ 'VENDEDOR' +';'+ 'COD_TAB' +';'+ 'TABELA' +';'+ 'PORC_ETERNA' +';'+ 'PORC_ETERNA_SUP' +';'+ 'PORC_ETERNA_DIR' +';'+ 'PORC_ETERNA_PRE' +';'+ 'PORC_AD' +';'+ 'PORC_MEN1' +';'+ 'PORC_MEN2' +';'+ 'PORC_MEN3' +';'+ 'PORC_MEN4' +';'+ 'PORC_MEN5' +';'+ 'PORC_MEN6' +';'+ 'PORC_MEN7' +';'+ 'PORC_MEN8' +';'+ 'PORC_MEN9' +';'+ 'PORC_MEN10' +';'+ 'PORC_MEN11' +';'+ 'PORC_MEN12');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_VEND').AsString;
          P2  := qryDados.FIELDBYNAME('VENDEDOR').AsString;
          P3  := qryDados.FIELDBYNAME('COD_TAB').AsString;
          P4  := qryDados.FIELDBYNAME('TABELA').AsString;
          P5  := qryDados.FIELDBYNAME('PORC_ETERNA').AsString;
          P6  := qryDados.FIELDBYNAME('PORC_ETERNA_SUP').AsString;
          P7  := qryDados.FIELDBYNAME('PORC_ETERNA_DIR').AsString;
          P8  := qryDados.FIELDBYNAME('PORC_ETERNA_PRE').AsString;
          P9  := qryDados.FIELDBYNAME('PORC_AD').AsString;
          P10 := qryDados.FIELDBYNAME('PORC_MEN1').AsString;
          P11 := qryDados.FIELDBYNAME('PORC_MEN2').AsString;
          P12 := qryDados.FIELDBYNAME('PORC_MEN3').AsString;
          P13 := qryDados.FIELDBYNAME('PORC_MEN4').AsString;
          P14 := qryDados.FIELDBYNAME('PORC_MEN5').AsString;
          P15 := qryDados.FIELDBYNAME('PORC_MEN6').AsString;
          P16 := qryDados.FIELDBYNAME('PORC_MEN7').AsString;
          P17 := qryDados.FIELDBYNAME('PORC_MEN8').AsString;
          P18 := qryDados.FIELDBYNAME('PORC_MEN9').AsString;
          P19 := qryDados.FIELDBYNAME('PORC_MEN10').AsString;
          P20 := qryDados.FIELDBYNAME('PORC_MEN11').AsString;
          P21 := qryDados.FIELDBYNAME('PORC_MEN12').AsString;


          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';
          cLinha := cLinha + P10 +';';
          cLinha := cLinha + P11 +';';
          cLinha := cLinha + P12 +';';
          cLinha := cLinha + P13 +';';
          cLinha := cLinha + P14 +';';
          cLinha := cLinha + P15 +';';
          cLinha := cLinha + P16 +';';
          cLinha := cLinha + P17 +';';
          cLinha := cLinha + P18 +';';
          cLinha := cLinha + P19 +';';
          cLinha := cLinha + P20 +';';
          cLinha := cLinha + P21 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;


  with qryExec do // INSERIR LOG DE EXPORTA��O
  begin
  Close;
  SQL.Clear;
  SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'E'+#39+', SYSDATE)');
  ExecSQL;
  end;

 end;
end;

procedure TfrmRelatorio3.qryDadosAfterOpen(DataSet: TDataSet);
begin
frmPrincipal.DimensionarGrid(dbgDados,Self);

  if frmPrincipal.lblCodRelatorio.Caption = '101' then
  begin
    TCurrencyField(qryDados.FieldByName('CPF')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DT_NASC')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('SEXO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('TIPO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('APH_OMT')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('TIPO_PLANO')).Alignment := taCenter;
  end;
   
  if frmPrincipal.lblCodRelatorio.Caption = '104' then
  begin
    TCurrencyField(qryDados.FieldByName('TABELA')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_SERV')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('CARENCIA')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COPARTICIPACAO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('VLR_COPART')).DisplayFormat := '#,##0.00';
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '113' then
  begin
    TCurrencyField(qryDados.FieldByName('COD_SEG')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('CPF')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DT_INCL')).Alignment := taCenter;
  end;

end;

end.
