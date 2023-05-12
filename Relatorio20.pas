unit Relatorio20;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, StdCtrls, Mask, Grids, DBGrids, Buttons, ComCtrls,
  AppEvnts;

type
  TfrmRelatorio20 = class(TForm)
    Label1: TLabel;
    btnExportar: TSpeedButton;
    lblTotal: TLabel;
    lblRTotal: TLabel;
    dbgDados: TDBGrid;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    Label2: TLabel;
    dsDados: TDataSource;
    SaveDialog1: TSaveDialog;
    qryExec: TADOQuery;
    qryDados: TADOQuery;
    chkInclusao: TCheckBox;
    chkEmpresa: TCheckBox;
    edtCodEmpresa: TEdit;
    edtNomeEmpresa: TEdit;
    chkUnidade: TCheckBox;
    edtNomeUnidade: TEdit;
    edtCodUnidade: TEdit;
    chkPlano: TCheckBox;
    cmbPlano: TComboBox;
    btnGerar: TSpeedButton;
    chkAditivo: TCheckBox;
    cmbAditivo: TComboBox;
    lblEmpresa: TLabel;
    lblUnidade: TLabel;
    dtInicio: TDateTimePicker;
    dtFim: TDateTimePicker;
    procedure chkEmpresaClick(Sender: TObject);
    procedure chkUnidadeClick(Sender: TObject);
    procedure edtCodEmpresaExit(Sender: TObject);
    procedure edtCodUnidadeExit(Sender: TObject);
    procedure chkPlanoClick(Sender: TObject);
    procedure chkAditivoClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure chkInclusaoClick(Sender: TObject);
    procedure btnGerarClick(Sender: TObject);
    procedure qryDadosAfterOpen(DataSet: TDataSet);
    procedure btnExportarClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmRelatorio20: TfrmRelatorio20;
  parametro: String;
implementation

uses Math,DataModule, Principal, StrUtils;
{$R *.dfm}

procedure TfrmRelatorio20.chkEmpresaClick(Sender: TObject);
begin
  if chkEmpresa.Checked = True then
  begin
    edtCodEmpresa.Enabled := True;
    chkUnidade.Enabled    := True;
  end else
  begin
    edtCodEmpresa.Enabled := False;
    edtCodEmpresa.Clear;
    edtNomeEmpresa.Clear;
    edtCodUnidade.Enabled := False;
    edtCodUnidade.Clear;
    edtNomeUnidade.Clear;
    chkUnidade.Checked    := False;
    chkUnidade.Enabled    := False;
    lblEmpresa.Caption    := 'X';
  end;
end;

procedure TfrmRelatorio20.chkUnidadeClick(Sender: TObject);
begin
  if chkUnidade.Checked = True then
  begin
    edtCodUnidade.Enabled := True;
  end else
  begin
    edtCodUnidade.Enabled := False;
    edtCodUnidade.Clear;
    edtNomeUnidade.Clear;
    lblUnidade.Caption := 'X';
  end;
end;

procedure TfrmRelatorio20.edtCodEmpresaExit(Sender: TObject);
begin
  if frmPrincipal.lblCodRelatorio.Caption = '122' then //Executar caso o relatório aberto seja o 122
  begin
    with qryExec do
    begin
      Close;
      SQL.Clear;
      SQL.Add('select cod_set, sigla from platinum.im_setor where cod_set = '+#39+edtCodEmpresa.Text+#39);
      Open;
    end;

    if qryExec.RecordCount = 1 then //Executar caso seja encontrado 1 registro na consulta realizada
    begin
      edtCodEmpresa.Text  := qryExec.FieldByName('cod_set').AsString;
      lblEmpresa.Caption  := qryExec.FieldByName('cod_set').AsString;
      edtNomeEmpresa.Font.Color := clWindowText;
      edtNomeEmpresa.Text := qryExec.FieldByName('sigla').AsString;
    end else
    begin
      lblEmpresa.Caption := 'X';
      edtCodEmpresa.Clear;
      edtNomeEmpresa.Font.Color := clRed;
      edtNomeEmpresa.Text := 'EMPRESA NÃO LOCALIZADA';
    end;
  end;
end;

procedure TfrmRelatorio20.edtCodUnidadeExit(Sender: TObject);
begin
  if frmPrincipal.lblCodRelatorio.Caption = '122' then //Executar caso o relatório aberto seja o 122
  begin
    with qryExec do
    begin
      Close;
      SQL.Clear;
      SQL.Add('select cod_setset, nome from platinum.im_setse where cod_setset = '+#39+edtCodUnidade.Text+#39);
      Open;
    end;

    if qryExec.RecordCount = 1 then //Executar caso seja encontrado 1 registro na consulta realizada
    begin
      edtCodUnidade.Text  := qryExec.FieldByName('cod_setset').AsString;
      lblUnidade.Caption  := qryExec.FieldByName('cod_setset').AsString;
      edtNomeUnidade.Text := qryExec.FieldByName('nome').AsString;
    end;
  end;
end;

procedure TfrmRelatorio20.chkPlanoClick(Sender: TObject);
begin
  if chkPlano.Checked = True then
  begin
    cmbPlano.Enabled := True;
  end else
  begin
    cmbPlano.Enabled   := False;
    cmbPlano.ItemIndex := -1; // Limpar seleção de plano
  end;
end;

procedure TfrmRelatorio20.chkAditivoClick(Sender: TObject);
begin
  if chkAditivo.Checked = True then
  begin
    cmbAditivo.Enabled := True;
  end else
  begin
    cmbAditivo.Enabled   := False;
    cmbAditivo.ItemIndex := -1; // Limpar seleção de plano
  end;
end;

procedure TfrmRelatorio20.FormShow(Sender: TObject);
begin

  with qryExec do
  begin
  Close;
  SQL.Clear;
  SQL.Add('ALTER SESSION SET NLS_DATE_FORMAT = '+#39+'DD/MM/YYYY HH24:MI:SS'+#39);
  ExecSQL;
  end;
  
  dtInicio.Date := Date();
  dtFim.Date := Date();  

  if frmPrincipal.lblCodRelatorio.Caption = '122' then
  begin
    with qryExec do  // Preencher combo de plano (Capturar dados)
    begin
      Close;
      SQL.Clear;
      SQL.Add('select cod_plano||'+#39+' - '+#39+'||nome plano from platinum.im_plano order by 1');
      Open;
      First;
    end;
    cmbPlano.Items.Clear;
    repeat // Preencher combo de plano (Inserir dados na lista de itens)
      cmbPlano.Items.Add(qryExec.FieldValues['plano']);
      qryExec.Next;
    until qryExec.Eof;
    cmbPlano.ItemIndex := 0;

    with qryExec do  // Preencher combo de aditivo (Capturar dados)
    begin
      Close;
      SQL.Clear;
      SQL.Add('select cod_adtv||'+#39+' - '+#39+'||nome aditivo from platinum.im_adtv where nome like '+#39+'%ODONTO%'+#39+' order by 1');
      Open;
      First;
    end;
    cmbAditivo.Items.Clear;
    repeat // Preencher combo de aditivo (Inserir dados na lista de itens)
      cmbAditivo.Items.Add(qryExec.FieldValues['aditivo']);
      qryExec.Next;
    until qryExec.Eof;
    cmbAditivo.ItemIndex := 0;
    qryExec.Close;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '123' then
  begin
    with qryExec do  // Preencher combo de plano (Capturar dados)
    begin
      Close;
      SQL.Clear;
      SQL.Add('select cod_plano||'+#39+' - '+#39+'||nome plano from im_plano order by 1');
      Open;
      First;
    end;
    cmbPlano.Items.Clear;
    repeat // Preencher combo de plano (Inserir dados na lista de itens)
      cmbPlano.Items.Add(qryExec.FieldValues['plano']);
      qryExec.Next;
    until qryExec.Eof;
    cmbPlano.ItemIndex := 0;

    with qryExec do  // Preencher combo de aditivo (Capturar dados)
    begin
      Close;
      SQL.Clear;
      SQL.Add('select cod_adtv||'+#39+' - '+#39+'||nome aditivo from im_adtv where nome like '+#39+'%ODONTO%'+#39+' order by 1');
      Open;
      First;
    end;
    cmbAditivo.Items.Clear;
    repeat // Preencher combo de aditivo (Inserir dados na lista de itens)
      cmbAditivo.Items.Add(qryExec.FieldValues['aditivo']);
      qryExec.Next;
    until qryExec.Eof;
    cmbAditivo.ItemIndex := 0;
    qryExec.Close;
  end;

  if frmPrincipal.lblExporta.Caption = 'T' then //Habilita botão Exportar
  begin
  btnExportar.Enabled := True;
  end else
  begin
  btnExportar.Enabled := False;
  end;

end;

procedure TfrmRelatorio20.chkInclusaoClick(Sender: TObject);
begin
  if chkInclusao.Checked = True then
  begin
    dtInicio.Enabled := True;
    dtFim.Enabled    := True;
  end else
  begin
    dtInicio.Enabled := False;
    dtFim.Enabled    := False;
  end;
end;

procedure TfrmRelatorio20.btnGerarClick(Sender: TObject);
begin
  parametro := ''; // Zerar parâmetro para log.

  if (frmPrincipal.lblCodRelatorio.Caption = '122') or (frmPrincipal.lblCodRelatorio.Caption = '123') then
  begin
    if chkEmpresa.Checked = True then
    begin
      if lblEmpresa.Caption = 'X' then // Verifica se o check de empresa está marcado e nenhuma empresa foi selecionada.
      begin
        ShowMessage('Favor preencher todos os parametros marcados.');
        Abort;
      end;
    end;

    if chkUnidade.Checked = True then
    begin
      if lblUnidade.Caption = 'X' then // Verifica se o check de unidade está marcado e nenhuma unidade foi selecionada.
      begin
        ShowMessage('Favor preencher todos os parametros marcados.');
        Abort;
      end;
    end;

    parametro := 'DtInicio: '+DateToStr(dtInicio.Date) + ' DtFinal: ' + DateToStr(dtFim.Date) + ' Empresa: '+lblEmpresa.Caption+ ' Unidade: '+lblUnidade.Caption+ ' Plano: '+cmbPlano.Text+ ' Aditivo: '+cmbAditivo.Text;
    with qryDados do
    begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);

      if chkInclusao.Checked = True then
        SQL.Add('  AND T.DT_CONCIL BETWEEN '+#39+DateToStr(dtInicio.Date)+#39+' AND '+#39+DateToStr(dtFim.Date)+#39);

      if chkEmpresa.Checked = True then
        SQL.Add('  AND S.COD_SET = '+#39+lblEmpresa.Caption+#39);

      if chkUnidade.Checked = True then
        SQL.Add('  AND S.COD_SETSET = '+#39+lblUnidade.Caption+#39);

      if chkPlano.Checked = True then
        SQL.Add('  AND (P.COD_PLANO||'+#39+' - '+#39+'||P.NOME = '+#39+cmbPlano.Text+#39+') OR (T2.COD_PLANO||'+#39+' - '+#39+'||T2.PLANO = '+#39+cmbPlano.Text+#39+')');

      if chkAditivo.Checked = True then
        SQL.Add('  (AND T1.COD_ADTV||'+#39+' - '+#39+'||T1.ADITIVO = '+#39+cmbAditivo.Text+#39+')');

      SQL.Add(' GROUP BY S.COD_SEG, S.NOME, DECODE(S.CANCELED,'+#39+'T'+#39+','+#39+'CANCELADO'+#39+','+#39+'F'+#39+','+#39+'ATIVO'+#39+'), DECODE(S.TIPO,'+#39+'E'+#39+','+#39+'TITULAR'+#39+','+#39+'D'+#39+','+#39+'DEPENDENTE'+#39+')');
      SQL.Add('         ,CASE');
      SQL.Add('            WHEN T2.COD_PLANO IS NULL THEN P.COD_PLANO||'+#39+' - '+#39+'||P.NOME');
      SQL.Add('              ELSE T2.COD_PLANO||'+#39+' - '+#39+'||T2.PLANO');
      SQL.Add('               END');
      SQL.Add('         ,T1.COD_ADTV||'+#39+' - '+#39+'||T1.ADITIVO, E.COD_SET, E.SIGLA, U.COD_SETSET, U.NOME');
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
    Abort;
  end;
end;

procedure TfrmRelatorio20.qryDadosAfterOpen(DataSet: TDataSet);
begin
frmPrincipal.DimensionarGrid(dbgDados,Self);

  if (frmPrincipal.lblCodRelatorio.Caption = '122') then
  begin
    TCurrencyField(qryDados.FieldByName('COD_SEG')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_EMPRESA')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_UNIDADE')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('SITUACAO')).Alignment := taCenter;    
    TCurrencyField(qryDados.FieldByName('TIPO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('VLR_ADTV_PAGO')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('VLR_ADTV_ABERTO')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('VLR_PLANO_PAGO')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('VLR_PLANO_ABERTO')).DisplayFormat := 'R$ #,##0.00';
  end;
end;

procedure TfrmRelatorio20.btnExportarClick(Sender: TObject);
var
F: TextFile;
P1, P2, P3, P4, P5, P6, P7, P8, P9, P10, P11, P12, P13, P14, P15, P16, P17, P18, P19, P20, P21, cLinha  : String;
begin
 if qryDados.RecordCount > 0 then
 begin
  if (frmPrincipal.lblCodRelatorio.Caption = '122') or (frmPrincipal.lblCodRelatorio.Caption = '123') then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_SEG' +';'+ 'NOME' +';'+ 'SITUACAO' +';'+ 'TIPO' +';'+ 'PLANO' +';'+ 'ADITIVO' +';'+ 'COD_EMPRESA' +';'+ 'EMPRESA' +';'+ 'COD_UNIDADE' +';'+ 'UNIDADE' +';'+ 'VLR_ADTV_PAGO' +';'+ 'VLR_ADTV_ABERTO' +';'+ 'VLR_PLANO_PAGO' +';'+ 'VLR_PLANO_ABERTO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1   := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P2   := qryDados.FIELDBYNAME('NOME').AsString;
          P3   := qryDados.FIELDBYNAME('SITUACAO').AsString;
          P4   := qryDados.FIELDBYNAME('TIPO').AsString;
          P5   := qryDados.FIELDBYNAME('PLANO').AsString;
          P6   := qryDados.FIELDBYNAME('ADITIVO').AsString;
          P7   := qryDados.FIELDBYNAME('COD_EMPRESA').AsString;
          P8   := qryDados.FIELDBYNAME('EMPRESA').AsString;
          P9   := qryDados.FIELDBYNAME('COD_UNIDADE').AsString;
          P10  := qryDados.FIELDBYNAME('UNIDADE').AsString;
          P11  := qryDados.FIELDBYNAME('VLR_ADTV_PAGO').AsString;
          P12  := qryDados.FIELDBYNAME('VLR_ADTV_ABERTO').AsString;
          P13  := qryDados.FIELDBYNAME('VLR_PLANO_PAGO').AsString;
          P14  := qryDados.FIELDBYNAME('VLR_PLANO_ABERTO').AsString;

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
