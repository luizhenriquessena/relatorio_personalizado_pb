unit Relatorio9;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, StdCtrls, Mask, Grids, DBGrids, Buttons;

type
  TfrmRelatorio9 = class(TForm)
    btnImprimir: TSpeedButton;
    btnExportar: TSpeedButton;
    Label1: TLabel;
    lblValor: TLabel;
    lblRValor: TLabel;
    btnGerar: TSpeedButton;
    dbgDados: TDBGrid;
    GroupBox1: TGroupBox;
    qryDados: TADOQuery;
    dsDados: TDataSource;
    SaveDialog1: TSaveDialog;
    GroupBox2: TGroupBox;
    Label2: TLabel;
    Med_DataFinal: TMaskEdit;
    Med_DataInicio: TMaskEdit;
    lblDia: TLabel;
    edtDia: TEdit;
    edtOperador: TEdit;
    lblOperador: TLabel;
    lblValorPago: TLabel;
    lblRValorPago: TLabel;
    lblQuantidade: TLabel;
    lblRQuantidade: TLabel;
    qryExec: TADOQuery;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure Med_DataInicioExit(Sender: TObject);
    procedure Med_DataFinalExit(Sender: TObject);
    procedure edtDiaKeyPress(Sender: TObject; var Key: Char);
    procedure btnGerarClick(Sender: TObject);
    procedure btnExportarClick(Sender: TObject);
    procedure qryDadosAfterOpen(DataSet: TDataSet);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmRelatorio9: TfrmRelatorio9;
  parametro: String;
implementation

{$R *.dfm}
uses Principal, DateUtils, Math;
procedure TfrmRelatorio9.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
qryDados.Close;
qryExec.Close;
end;

procedure TfrmRelatorio9.FormShow(Sender: TObject);
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

  if frmPrincipal.lblCodRelatorio.Caption = '18' then
  begin
  lblOperador.Caption := 'Operador da Cobrança: ';
  edtDia.Visible := True;
  lblDia.Visible := True;
  lblValor.Visible := True;
  lblRValor.Visible := True;
  lblValorPago.Visible := True;
  lblRValorPago.Visible := True;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '29' then
  begin
  lblOperador.Caption := 'Palavra Chave: ';
  edtDia.Visible := False;
  lblDia.Visible := False;
  lblValor.Visible := False;
  lblRValor.Visible := False;
  lblValorPago.Visible := False;
  lblRValorPago.Visible := False;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '71' then
  begin
  lblOperador.Caption := 'Login Analista: ';
  edtDia.Visible := False;
  lblDia.Visible := False;
  lblValor.Visible := False;
  lblRValor.Visible := False;
  lblValorPago.Visible := False;
  lblRValorPago.Visible := False;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '87' then
  begin
  lblOperador.Caption := 'Login Auditor (Todos = %): ';
  edtDia.Visible := False;
  lblDia.Visible := False;
  lblValor.Visible := False;
  lblRValor.Visible := False;
  lblValorPago.Visible := False;
  lblRValorPago.Visible := False;
  end;

  Med_DataInicio.Text := '  /  /    ';
  Med_DataFinal.Text := '  /  /    ';
  edtDia.Text := '';
  edtOperador.Text := '';
  lblRValor.Caption := '0';
  lblRValorPago.Caption := '0';
  lblRQuantidade.Caption := '0';
end;

procedure TfrmRelatorio9.Med_DataInicioExit(Sender: TObject);
var d_Data : TDateTime;
begin
  try
    d_Data := StrToDateTime(Med_DataInicio.Text);
    Med_DataInicio.Text := FormatDateTime('dd/mm/yyyy',d_Data);
  except
    ShowMessage('Data inválida');
    Med_DataInicio.Text := '  /  /    ';
    Med_DataInicio.SetFocus;
  end;
end;

procedure TfrmRelatorio9.Med_DataFinalExit(Sender: TObject);
var d_Data : TDateTime;
begin
  try
    d_Data := StrToDateTime(Med_DataFinal.Text);
    Med_DataFinal.Text := FormatDateTime('dd/mm/yyyy',d_Data);
  except
    ShowMessage('Data inválida');
    Med_DataFinal.Text := '  /  /    ';
    Med_DataFinal.SetFocus;
  end;
end;

procedure TfrmRelatorio9.edtDiaKeyPress(Sender: TObject; var Key: Char);
begin
    if not (key in ['0'..'9',Chr(8), Chr(3), Chr(22), Chr(24), Chr(44)]) then
      key := #0;
end;

procedure TfrmRelatorio9.btnGerarClick(Sender: TObject);
begin
  Med_DataInicio.OnExit(Sender);
  Med_DataFinal.OnExit(Sender);
  parametro := '';
  if frmPrincipal.lblCodRelatorio.Caption = '18' then
  begin
    if (Med_DataInicio.Text <> '  /  /    ') and (Med_DataFinal.Text <> '  /  /    ') and (edtDia.Text <> '') and (edtOperador.Text <> '') then
    begin
      parametro := 'DtInicio: '+Med_DataInicio.Text + ' DtFinal: ' + Med_DataFinal.Text + ' Dia: ' + edtDia.Text + ' Operador: ' + edtOperador.Text;
      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);
      SQL.Text := StringReplace(SQL.Text,'&DATA_INICIAL',Med_DataInicio.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&DATA_FINAL',Med_DataFinal.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&DIAS_ATRASO',edtDia.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&LOGIN',UpperCase(edtOperador.Text),[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      Open;
      end;

      with qryExec do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);
      SQL.Text := StringReplace(SQL.Text,'&DATA_INICIAL',Med_DataInicio.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&DATA_FINAL',Med_DataFinal.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&DIAS_ATRASO',edtDia.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&LOGIN',edtOperador.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      Open;
      end;

      lblRValor.Caption := qryExec.FieldByName('VALOR').AsString;
      lblRValorPago.Caption := qryExec.FieldByName('VALOR_PG').AsString;
      lblRQuantidade.Caption := IntToStr(qryDados.RecordCount);

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
    Abort;
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '29' then
  begin
    if (Med_DataInicio.Text <> '  /  /    ') and (Med_DataFinal.Text <> '  /  /    ') and (edtOperador.Text <> '') then
    begin
      parametro := 'DtInicio: '+Med_DataInicio.Text + ' DtFinal: ' + Med_DataFinal.Text + ' Palavra Chave: ' + edtOperador.Text;
      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);
      SQL.Text := StringReplace(SQL.Text,'&DATA1',Med_DataInicio.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&DATA2',Med_DataFinal.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&CHAVE',UpperCase(edtOperador.Text),[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      Open;
      end;

      lblRQuantidade.Caption := IntToStr(qryDados.RecordCount);

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
    Abort;
    end;
  end;

  if (frmPrincipal.lblCodRelatorio.Caption = '71') or (frmPrincipal.lblCodRelatorio.Caption = '87') then
  begin
    if (Med_DataInicio.Text <> '  /  /    ') and (Med_DataFinal.Text <> '  /  /    ') and (edtOperador.Text <> '') then
    begin
      parametro := 'DtInicio: '+Med_DataInicio.Text + ' DtFinal: ' + Med_DataFinal.Text + ' Analista: ' + edtOperador.Text;
      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);
      SQL.Text := StringReplace(SQL.Text,'&DATA1',Med_DataInicio.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&DATA2',Med_DataFinal.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&CHAVE',UpperCase(edtOperador.Text),[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      Open;
      end;

      lblRQuantidade.Caption := IntToStr(qryDados.RecordCount);

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
    Abort;
    end;
  end;

end;

procedure TfrmRelatorio9.btnExportarClick(Sender: TObject);
var
F: TextFile;
P1, P2, P3, P4, P5, P6, P7, P8, P9, P10, P11, P12, P13, P14, cLinha  : String;
begin
 if qryDados.RecordCount > 0 then
 begin
  if frmPrincipal.lblCodRelatorio.Caption = '18' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_BOL' +';'+ 'VALOR' +';'+ 'VALOR_PG' +';'+ 'DT_EMISSAO' +';'+ 'DT_CONCIL' +';'+ 'CONCIL' +';'+ 'TIP_LIG' +';'+ 'COD_SEG' +';'+ 'COD_SET' +';'+ 'DT_VENC' +';'+ 'DT_LIG' +';'+ 'CODTIPO_LIG' +';'+ 'LOGIN' +';'+ 'DT_LIG2');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_BOL').AsString;
          P2  := qryDados.FIELDBYNAME('VALOR').AsString;
          P3  := qryDados.FIELDBYNAME('VALOR_PG').AsString;
          P4  := qryDados.FIELDBYNAME('DT_EMISSAO').AsString;
          P5  := qryDados.FIELDBYNAME('DT_CONCIL').AsString;
          P6  := qryDados.FIELDBYNAME('CONCIL').AsString;
          P7  := qryDados.FIELDBYNAME('TIP_LIG').AsString;
          P8  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P9  := qryDados.FIELDBYNAME('COD_SET').AsString;
          P10 := qryDados.FIELDBYNAME('DT_VENC').AsString;
          P11 := qryDados.FIELDBYNAME('DT_LIG').AsString;
          P12 := qryDados.FIELDBYNAME('COD_TIPO_LIG').AsString;
          P13 := qryDados.FIELDBYNAME('LOGIN').AsString;
          P14 := qryDados.FIELDBYNAME('DT_LIG2').AsString;

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

  if frmPrincipal.lblCodRelatorio.Caption = '29' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'CODIGO' +';'+ 'NOME' +';'+ 'DATA' +';'+ 'TIPO_LIGACAO' +';'+ 'OBS' +';'+ 'USUARIO' +';'+ 'TEMPO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('CODIGO').AsString;
          P2  := qryDados.FIELDBYNAME('NOME').AsString;
          P3  := qryDados.FIELDBYNAME('DATA').AsString;
          P4  := qryDados.FIELDBYNAME('TIPO_LIGACAO').AsString;
          P5  := qryDados.FIELDBYNAME('OBS').AsString;
          P6  := qryDados.FIELDBYNAME('USUARIO').AsString;
          P7  := qryDados.FIELDBYNAME('TEMPO').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + StringReplace(StringReplace(P5,';',' ',[rfReplaceAll]),#13#10,' ',[rfReplaceAll]) +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '71' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'GUIA' +';'+ 'LOGIN' +';'+ 'DATA' +';'+ 'COD_AMB');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('GUIA').AsString;
          P2  := qryDados.FIELDBYNAME('USER_CONC_LOGIN').AsString;
          P3  := qryDados.FIELDBYNAME('USER_CONC_DT').AsString;
          P4  := qryDados.FIELDBYNAME('COD_AMB').AsString;
          
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

  if frmPrincipal.lblCodRelatorio.Caption = '87' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'DATA' +';'+ 'LOGIN' +';'+ 'COD_OPERACAO' +';'+ 'DESCRICAO' +';'+ 'GUIA');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('DATA').AsString;
          P2  := qryDados.FIELDBYNAME('LOGIN').AsString;
          P3  := qryDados.FIELDBYNAME('COD_OPERACAO').AsString;
          P4  := qryDados.FIELDBYNAME('DESCRICAO').AsString;
          P5  := qryDados.FIELDBYNAME('GUIA').AsString;

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

procedure TfrmRelatorio9.qryDadosAfterOpen(DataSet: TDataSet);
begin
 frmPrincipal.DimensionarGrid(dbgDados,Self);
end;

end.
