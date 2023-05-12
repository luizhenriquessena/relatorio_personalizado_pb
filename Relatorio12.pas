unit Relatorio12;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Mask, DB, ADODB, Grids, DBGrids, Buttons;

type
  TfrmRelatorio12 = class(TForm)
    Label1: TLabel;
    btnExportar: TSpeedButton;
    btnImprimir: TSpeedButton;
    lblTotal: TLabel;
    lblRTotal: TLabel;
    GroupBox1: TGroupBox;
    btnGerar: TSpeedButton;
    dbgDados: TDBGrid;
    dsDados: TDataSource;
    qryDados: TADOQuery;
    SaveDialog1: TSaveDialog;
    Med_MesInicio: TMaskEdit;
    Label2: TLabel;
    qryExec: TADOQuery;
    GroupBox2: TGroupBox;
    Label3: TLabel;
    Med_Mes01: TMaskEdit;
    Med_Mes02: TMaskEdit;
    lblAno: TLabel;
    Med_Ano: TMaskEdit;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure Med_MesInicioExit(Sender: TObject);
    procedure btnGerarClick(Sender: TObject);
    procedure btnExportarClick(Sender: TObject);
    procedure qryDadosAfterOpen(DataSet: TDataSet);
    procedure Med_Mes01Exit(Sender: TObject);
    procedure Med_Mes02Exit(Sender: TObject);
    procedure Med_AnoKeyPress(Sender: TObject; var Key: Char);
    procedure Med_AnoExit(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmRelatorio12: TfrmRelatorio12;
  MesInicio, Mes1, Mes2, Ano : String;
  parametro: String;
implementation

{$R *.dfm}
uses Principal;
procedure TfrmRelatorio12.FormClose(Sender: TObject; var Action: TCloseAction);
begin
qryDados.Close;
end;

procedure TfrmRelatorio12.FormShow(Sender: TObject);
begin
  Med_MesInicio.Text := '  /    ';
  Med_Mes01.Text := '  /    ';
  Med_Mes02.Text := '  /    ';
  Med_Ano.Text := '    ';

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

  if (frmPrincipal.lblCodRelatorio.Caption = '108') or (frmPrincipal.lblCodRelatorio.Caption = '109') or (frmPrincipal.lblCodRelatorio.Caption = '143') then
  begin
    GroupBox2.Visible := false;
    Label2.Visible := false;
    Med_MesInicio.Visible := false;
    lblAno.Visible := true;
    Med_Ano.Visible := true;
  end else
  begin
    GroupBox2.Visible := false;
    Label2.Visible := true;
    Med_MesInicio.Visible := true;
    lblAno.Visible := false;
    Med_Ano.Visible := false;
  end;

end;

procedure TfrmRelatorio12.Med_MesInicioExit(Sender: TObject);
var
d_Data : TDate;
begin
  try
    d_Data := StrToDate('01/'+Med_MesInicio.Text);
    MesInicio := FormatDateTime('mm/yyyy',d_Data);
  except
    ShowMessage('Data base inválida');
    Med_MesInicio.Text := '  /    ';
    Med_MesInicio.SetFocus;
  end;
end;

procedure TfrmRelatorio12.btnGerarClick(Sender: TObject);
begin

  parametro := '';

  if (frmPrincipal.lblCodRelatorio.Caption = '108') or (frmPrincipal.lblCodRelatorio.Caption = '109') or (frmPrincipal.lblCodRelatorio.Caption = '143') then
  begin
  Med_Ano.OnExit(Sender);
    if (Med_Ano.Text <> '    ') then
    begin
      parametro := 'Ano : '+ Med_Ano.Text;
      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);
      SQL.Text := StringReplace(SQL.Text,'&ANO',Ano,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      Open;
      end;

      lblRTotal.Caption := IntToStr(qryDados.RecordCount);

      with qryExec do // INSERIR LOG DE VISUALIZAÇÃO
      begin
      Close;
      SQL.Clear;
      SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, PARAMETRO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'V'+#39+', '+#39+parametro+#39+', SYSDATE)');
      ExecSQL;
      end;

    end else
    begin
    ShowMessage('Favor informar o período.');
    end;

  end else
  begin
  Med_MesInicio.OnExit(Sender);
    if (Med_MesInicio.Text <> '  /    ') then
    begin
      parametro := 'MesAno: '+MesInicio;
      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);
      SQL.Text := StringReplace(SQL.Text,'&MESANO',MesInicio,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      Open;
      end;

      lblRTotal.Caption := IntToStr(qryDados.RecordCount);

      with qryExec do // INSERIR LOG DE VISUALIZAÇÃO
      begin
      Close;
      SQL.Clear;
      SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, PARAMETRO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'V'+#39+', '+#39+parametro+#39+', SYSDATE)');
      ExecSQL;
      end;

    end else
    begin
    ShowMessage('Favor informar o período.');
    end;
  end;
end;

procedure TfrmRelatorio12.btnExportarClick(Sender: TObject);
var
F: TextFile;
P1, P2, P3, P4, P5, P6, P7, P8, P9, P10, P11, P12, P13, P14, P15, P16, P17, P18,
P19, P20, P21, P22, P23, P24, P25, P26, P27, cLinha : String; //Variaveis comuns
begin
 if qryDados.RecordCount > 0 then
 begin
  if frmPrincipal.lblCodRelatorio.Caption = '35' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_SET' +';'+ 'DESCRICAO' +';'+ 'SIGLA' +';'+ 'DT_CAD' +';'+ 'VIDAS_ATIVAS' +';'+ 'FATURA_ATUAL' +';'+ 'GERADO_ATUAL' +';'+ 'GERADO_MESANO' +';'+ 'PAGO_MESANO' +';'+ 'DESPESA_MESANO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1 := qryDados.FIELDBYNAME('COD_SET').AsString;
          P2 := qryDados.FIELDBYNAME('DESCRICAO').AsString;
          P3 := qryDados.FIELDBYNAME('SIGLA').AsString;
          P4 := qryDados.FIELDBYNAME('DT_CAD').AsString;
          P5 := qryDados.FIELDBYNAME('VIDAS_ATIVAS').AsString;
          P6 := qryDados.FIELDBYNAME('FATURA_ATUAL').AsString;
          P7 := qryDados.FIELDBYNAME('GERADO_MESANO').AsString;
          P8 := qryDados.FIELDBYNAME('PAGO_MESANO').AsString;
          P9 := qryDados.FIELDBYNAME('DESPESA_MESANO').AsString;

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

  if frmPrincipal.lblCodRelatorio.Caption = '37' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_SET' +';'+ 'DESCRICAO' +';'+ 'SITUACAO' +';'+ 'TIPO_PLANO_PADRAO' +';'+ 'NATUREZA_PLANO_PADRAO' +';'+ 'MESANO' +';'+ 'TIPO_PLANO_REC' +';'+ 'NATUREZA_PLANO_REC' +';'+ 'VIDAS_BOL_MESANO' +';'+ 'RECEITA_MESANO' +';'+ 'TIPO_PLANO_DESP' +';'+ 'NATUREZA_PLANO_DESP' +';'+ 'DESPESA_MESANO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1 := qryDados.FIELDBYNAME('COD_SET').AsString;
          P2 := qryDados.FIELDBYNAME('DESCRICAO').AsString;
          P3 := qryDados.FIELDBYNAME('SITUACAO').AsString;
          P4 := qryDados.FIELDBYNAME('TIPO_PLANO_PADRAO').AsString;
          P5 := qryDados.FIELDBYNAME('NATUREZA_PLANO_PADRAO').AsString;
          P6 := qryDados.FIELDBYNAME('MESANO').AsString;
          P7 := qryDados.FIELDBYNAME('TIPO_PLANO_REC').AsString;
          P8 := qryDados.FIELDBYNAME('NATUREZA_PLANO_REC').AsString;
          P9 := qryDados.FIELDBYNAME('VIDAS_BOL_MESANO').AsString;
          P10 := qryDados.FIELDBYNAME('RECEITA_MESANO').AsString;
          P11 := qryDados.FIELDBYNAME('TIPO_PLANO_DESP').AsString;
          P12 := qryDados.FIELDBYNAME('NATUREZA_PLANO_DESP').AsString;
          P13 := qryDados.FIELDBYNAME('DESPESA_MESANO').AsString;

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

  if frmPrincipal.lblCodRelatorio.Caption = '45' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'EMP' +';'+ 'COMP' +';'+ 'RECEITA' +';'+ 'DESPESA' +';'+ 'SALDO' +';'+ 'COMISSAO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1 := qryDados.FIELDBYNAME('EMP').AsString;
          P2 := qryDados.FIELDBYNAME('COMP').AsString;
          P3 := qryDados.FIELDBYNAME('RECEITA').AsString;
          P4 := qryDados.FIELDBYNAME('DESPESA').AsString;
          P5 := qryDados.FIELDBYNAME('SALDO').AsString;
          P6 := qryDados.FIELDBYNAME('COMISSAO').AsString;

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

  if frmPrincipal.lblCodRelatorio.Caption = '49' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_OUTRO2' +';'+ 'COD_SEG' +';'+ 'NOME' +';'+ 'MAT' +';'+ 'DT_INCL' +';'+ 'CORRETOR' +';'+ 'TOTAL_COB' +';'+ 'TOTAL_PAG' +';'+ 'TOTAL_INADIM' +';'+ 'UTILIZACAO' +';'+ 'SALDO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_OUTRO2').AsString;
          P2  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P3  := qryDados.FIELDBYNAME('NOME').AsString;
          P4  := qryDados.FIELDBYNAME('MAT').AsString;
          P5  := qryDados.FIELDBYNAME('DT_INCL').AsString;
          P6  := qryDados.FIELDBYNAME('CORRETOR').AsString;
          P7  := qryDados.FIELDBYNAME('TOTAL_COB').AsString;
          P8  := qryDados.FIELDBYNAME('TOTAL_PAG').AsString;
          P9  := qryDados.FIELDBYNAME('TOTAL_INADIM').AsString;
          P10 := qryDados.FIELDBYNAME('UTILIZACAO').AsString;
          P11 := qryDados.FIELDBYNAME('SALDO').AsString;

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

  if frmPrincipal.lblCodRelatorio.Caption = '50' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_MOT' +';'+ 'DESCRICAO' +';'+ 'QTD');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_MOT').AsString;
          P2  := qryDados.FIELDBYNAME('DESCRICAO').AsString;
          P3  := qryDados.FIELDBYNAME('QTD').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '62' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_VENDEDOR' +';'+ 'VENDEDOR' +';'+ 'TIPO_VEND' +';'+ 'COD_SUPERVISOR' +';'+ 'COD_BENEFICIARIO' +';'+ 'BENEFICIARIO' +';'+ 'COD_EMPRESA' +';'+ 'EMPRESA' +';'+ 'SETOR' +';'+ 'COD_VEMDA' +';'+ 'DT_VENC' +';'+ 'DT_CONCIL' +';'+ 'VALOR_BOLETA' +';'+ 'VALOR_PAGO' +';'+ 'COMISSAO' +';'+ 'PERC_COMISSAO' +';'+ 'PROVENTO' +';'+ 'PROVENTO_PAGO' +';'+ 'ESTORNO' +';'+ 'TAXA' +';'+ 'ESTORNO_PAGO' +';'+ 'LANC_MANUAL' +';'+ 'COD_BANCO' +';'+ 'BANCO' +';'+ 'AGENCIA' +';'+ 'CONTA' +';'+ 'COD_TITULAR');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_VENDEDOR').AsString;
          P2  := qryDados.FIELDBYNAME('VENDEDOR').AsString;
          P3  := qryDados.FIELDBYNAME('TIPO_VEND').AsString;
          P4  := qryDados.FIELDBYNAME('COD_SUPERVISOR').AsString;
          P5  := qryDados.FIELDBYNAME('COD_BENEFICIARIO').AsString;
          P6  := qryDados.FIELDBYNAME('BENEFICIARIO').AsString;
          P7  := qryDados.FIELDBYNAME('COD_EMPRESA').AsString;
          P8  := qryDados.FIELDBYNAME('EMPRESA').AsString;
          P9  := qryDados.FIELDBYNAME('SETOR').AsString;
          P10 := qryDados.FIELDBYNAME('COD_VENDA').AsString;
          P11 := qryDados.FIELDBYNAME('DT_VENC').AsString;
          P12 := qryDados.FIELDBYNAME('DT_CONCIL').AsString;
          P13 := qryDados.FIELDBYNAME('VALOR_BOLETA').AsString;
          P14 := qryDados.FIELDBYNAME('VALOR_PAGO').AsString;
          P15 := qryDados.FIELDBYNAME('COMISSAO').AsString;
          P16 := qryDados.FIELDBYNAME('PERC_COMISSAO').AsString;
          P17 := qryDados.FIELDBYNAME('PROVENTO').AsString;
          P18 := qryDados.FIELDBYNAME('PROVENTO_PAGO').AsString;
          P19 := qryDados.FIELDBYNAME('ESTORNO').AsString;
          P20 := qryDados.FIELDBYNAME('ESTORNO_PAGO').AsString;
          P21 := qryDados.FIELDBYNAME('TAXA').AsString;
          P22 := qryDados.FIELDBYNAME('LANC_MANUAL').AsString;
          P23 := qryDados.FIELDBYNAME('COD_BANCO').AsString;
          P24 := qryDados.FIELDBYNAME('BANCO').AsString;
          P25 := qryDados.FIELDBYNAME('AGENCIA').AsString;
          P26 := qryDados.FIELDBYNAME('CONTA').AsString;
          P27 := qryDados.FIELDBYNAME('COD_TITULAR').AsString;

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
          cLinha := cLinha + P26 +';';
          cLinha := cLinha + P27 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '75' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'TIPO' +';'+ 'COD_FORN' +';'+ 'NOME_FORN' +';'+ 'COD_ADTV'  +';'+ 'NOME_FORN' +';'+ 'NOME_ADTV' +';'+ 'VALOR' +';'+ 'COD_EMPRESA' +';'+ 'EMPRESA' +';'+ 'COD_SEG' +';'+ 'NOME_SEG' +';'+ 'MES_PGTO' +';'+ 'CPF' +';'+ 'DT_ADESAO' +';'+ 'DATA_ENTRADA' +';'+ 'DATA_ENTRADA2' +';'+ 'DT_INICIO' +';'+ 'TIPO_SEG');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('TIPO').AsString;
          P2  := qryDados.FIELDBYNAME('COD_FORN').AsString;
          P3  := qryDados.FIELDBYNAME('NOME_FORN').AsString;
          P4  := qryDados.FIELDBYNAME('COD_ADTV').AsString;
          P5  := qryDados.FIELDBYNAME('NOME_ADTV').AsString;
          P6  := qryDados.FIELDBYNAME('VALOR').AsString;
          P7  := qryDados.FIELDBYNAME('COD_SET').AsString;
          P8  := qryDados.FIELDBYNAME('DESCRICAO').AsString;
          P9  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P10 := qryDados.FIELDBYNAME('NOME_SEG').AsString;
          P11 := qryDados.FIELDBYNAME('MES_PGTO').AsString;
          P12 := qryDados.FIELDBYNAME('CPF').AsString;
          P13 := qryDados.FIELDBYNAME('DT_ADESAO').AsString;
          P14 := qryDados.FIELDBYNAME('DATA_ENTRADA').AsString;
          P15 := qryDados.FIELDBYNAME('DATA_ENTRADA2').AsString;
          P16 := qryDados.FIELDBYNAME('DT_INICIO').AsString;
          P17 := qryDados.FIELDBYNAME('TIPO_SEG').AsString;

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

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '143' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'ANO' +';'+ 'CODIGO' +';'+ 'NOME' +';'+ 'GRUPO' +';'+ 'SITUACAO' +';'+ 'JAN' +';'+ 'FEV' +';'+ 'MAR' +';'+ 'ABR' +';'+ 'MAI' +';'+ 'JUN' +';'+ 'JUL' +';'+ 'AGO' +';'+ 'SET' +';'+ 'OUT' +';'+ 'NOV' +';'+ 'DEZ');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1 := qryDados.FIELDBYNAME('ANO').AsString;
          P2 := qryDados.FIELDBYNAME('CODIGO').AsString;
          P3 := qryDados.FIELDBYNAME('NOME').AsString;
          P4 := qryDados.FIELDBYNAME('GRUPO').AsString;
          P5 := qryDados.FIELDBYNAME('SITUACAO').AsString;
          P6 := qryDados.FIELDBYNAME('JAN').AsString;
          P7 := qryDados.FIELDBYNAME('FEV').AsString;
          P8 := qryDados.FIELDBYNAME('MAR').AsString;
          P9 := qryDados.FIELDBYNAME('ABR').AsString;
          P10 := qryDados.FIELDBYNAME('MAI').AsString;
          P11 := qryDados.FIELDBYNAME('JUN').AsString;
          P12 := qryDados.FIELDBYNAME('JUL').AsString;
          P13 := qryDados.FIELDBYNAME('AGO').AsString;
          P14 := qryDados.FIELDBYNAME('SET').AsString;
          P15 := qryDados.FIELDBYNAME('OUT').AsString;
          P16 := qryDados.FIELDBYNAME('NOV').AsString;
          P17 := qryDados.FIELDBYNAME('DEZ').AsString;

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

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

  if (frmPrincipal.lblCodRelatorio.Caption = '108') or (frmPrincipal.lblCodRelatorio.Caption = '109') then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'MES_ANO' +';'+ 'CADASTRADOS' +';'+ 'ATIVOS' +';'+ 'CANCELADOS' +';'+ 'COMISSAO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('ANO').AsString;
          P2  := qryDados.FIELDBYNAME('CADASTRADOS').AsString;
          P3  := qryDados.FIELDBYNAME('ATIVOS').AsString;
          P4  := qryDados.FIELDBYNAME('CANCELADOS').AsString;
          P5  := qryDados.FIELDBYNAME('COMISSAO').AsString;

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

  with qryExec do // INSERIR LOG DE EXPORTAÇÃO
  begin
  Close;
  SQL.Clear;
  SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, PARAMETRO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'E'+#39+', '+#39+parametro+#39+', SYSDATE)');
  ExecSQL;
  end;
 end;
end;

procedure TfrmRelatorio12.qryDadosAfterOpen(DataSet: TDataSet);
begin
  frmPrincipal.DimensionarGrid(dbgDados,Self);

  if frmPrincipal.lblCodRelatorio.Caption = '75' then
  begin
    TCurrencyField(qryDados.FieldByName('TIPO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_FORN')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_ADTV')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('VALOR')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('VALOR')).Alignment := taRightJustify;
    TCurrencyField(qryDados.FieldByName('COD_SET')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_SEG')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('MES_PGTO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('CPF')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DT_ADESAO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DATA_ENTRADA')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DATA_ENTRADA2')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DT_INICIO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('TIPO_SEG')).Alignment := taCenter;
  end;

  if (frmPrincipal.lblCodRelatorio.Caption = '108') or (frmPrincipal.lblCodRelatorio.Caption = '109') then
  begin
    TCurrencyField(qryDados.FieldByName('ANO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COMISSAO')).DisplayFormat := 'R$ #,##0.00';
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '143' then
  begin
    TCurrencyField(qryDados.FieldByName('ANO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('JAN')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('FEV')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('MAR')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('ABR')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('MAI')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('JUN')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('JUL')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('AGO')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('SET')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('OUT')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('NOV')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('DEZ')).DisplayFormat := 'R$ #,##0.00';
  end;
end;

procedure TfrmRelatorio12.Med_Mes01Exit(Sender: TObject);
var
d_Data : TDate;
begin
  try
    d_Data := StrToDate('01/'+Med_Mes01.Text);
    Mes1   := FormatDateTime('mm/yyyy',d_Data);
  except
    ShowMessage('Mês início inválido');
    Med_Mes01.Text := '  /    ';
    Med_Mes01.SetFocus;
  end;
end;

procedure TfrmRelatorio12.Med_Mes02Exit(Sender: TObject);
var
d_Data : TDate;
begin
  try
    d_Data := StrToDate('01/'+Med_Mes02.Text);
    Mes2 := DateToStr(d_Data);
  except
    ShowMessage('Mês início inválido');
    Med_Mes02.Text := '  /    ';
    Med_Mes02.SetFocus;
  end;

end;

procedure TfrmRelatorio12.Med_AnoKeyPress(Sender: TObject; var Key: Char);
begin
    if not (key in ['0'..'9',Chr(8), Chr(3), Chr(22), Chr(24), Chr(44)]) then
      key := #0;
end;

procedure TfrmRelatorio12.Med_AnoExit(Sender: TObject);
var
d_Data : TDate;
begin
  try
    d_Data := StrToDate('01/01/'+Med_Ano.Text);
    Ano    := FormatDateTime('yyyy',d_Data);
  except
    ShowMessage('Mês início inválido');
    Med_Ano.Text := '    ';
    Med_Ano.SetFocus;
  end;
end;

end.
