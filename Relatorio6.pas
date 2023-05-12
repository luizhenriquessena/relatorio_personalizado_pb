unit Relatorio6;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, Grids, DBGrids, StdCtrls, Buttons;

type
  TfrmRelatorio6 = class(TForm)
    Label1: TLabel;
    btnExportar: TSpeedButton;
    btnImprimir: TSpeedButton;
    lblTotal: TLabel;
    lblRTotal: TLabel;
    GroupBox1: TGroupBox;
    lblDataBase: TLabel;
    btnGerar: TSpeedButton;
    edtDia: TEdit;
    dbgDados: TDBGrid;
    dsDados: TDataSource;
    qryDados: TADOQuery;
    SaveDialog1: TSaveDialog;
    qryExec: TADOQuery;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
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
  frmRelatorio6: TfrmRelatorio6;
  parametro: String;
implementation

{$R *.dfm}
uses MaskUtils, Math, Principal;
procedure TfrmRelatorio6.FormClose(Sender: TObject; var Action: TCloseAction);
begin
qryDados.Close;
end;

procedure TfrmRelatorio6.FormShow(Sender: TObject);
begin
  if frmPrincipal.lblCodRelatorio.Caption = '13' then
  lblDataBase.Caption := 'Dias em Aberto:';

  if frmPrincipal.lblCodRelatorio.Caption = '15' then
  lblDataBase.Caption := 'Utilização >= ';

  if frmPrincipal.lblCodRelatorio.Caption = '32' then
  lblDataBase.Caption := 'Boleto: ';

  if frmPrincipal.lblCodRelatorio.Caption = '47' then
  lblDataBase.Caption := 'Ano: ';

  if frmPrincipal.lblCodRelatorio.Caption = '89' then
  lblDataBase.Caption := 'Dia Vencimento: ';

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

  edtDia.Text:= '';
end;

procedure TfrmRelatorio6.edtDiaKeyPress(Sender: TObject; var Key: Char);
begin
    if not (key in ['0'..'9',Chr(8), Chr(3), Chr(22), Chr(24), Chr(44)]) then
      key := #0;
end;

procedure TfrmRelatorio6.btnGerarClick(Sender: TObject);
begin
  parametro := '';
  if edtDia.Text <> '' then
  begin
    if frmPrincipal.lblCodRelatorio.Caption = '13' then
    begin
      parametro := 'Dia: '+ edtDia.Text;
      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);
      SQL.Text := StringReplace(SQL.Text,'&DIAS_INADIMPLENCIA',edtDia.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      Open;
      end;
    end;

    if frmPrincipal.lblCodRelatorio.Caption = '15' then
    begin
      parametro := 'Utilização >= '+ edtDia.Text;
      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);
      SQL.Text := StringReplace(SQL.Text,'&VALOR',edtDia.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      Open;
      end;
    end;

    if frmPrincipal.lblCodRelatorio.Caption = '32' then
    begin
      parametro := 'Boleto: '+ edtDia.Text;
      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);
      SQL.Text := StringReplace(SQL.Text,'&BOLETO',edtDia.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      Open;
      end;
    end;

    if frmPrincipal.lblCodRelatorio.Caption = '47' then
    begin
      parametro := 'Ano: '+ edtDia.Text;
      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);
      SQL.Text := StringReplace(SQL.Text,'&ANO',edtDia.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      Open;
      end;
    end;

    if frmPrincipal.lblCodRelatorio.Caption = '89' then
    begin
      parametro := 'Dia Vencimento: '+ edtDia.Text;
      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);
      SQL.Text := StringReplace(SQL.Text,'&DIA',edtDia.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      Open;
      end;
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
    ShowMessage('Parâmetro não informado.');
  end;
end;

procedure TfrmRelatorio6.btnExportarClick(Sender: TObject);
var
F: TextFile;

P1, P2, P3, P4, P5, P6, P7, P8, P9, P10, P11, P12, P13, P14, P15, cLinha : String;
begin
 if qryDados.RecordCount > 0 then
 begin

   if frmPrincipal.lblCodRelatorio.Caption = '13' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'CODIGO' +';'+ 'NOME' +';'+ 'TIPO' +';'+ 'CPF' +';'+ 'COD_ADTV' +';'+ 'NOME_ADTV' +';'+ 'DESCRICAO' +';'+ 'DT_INCL' +';'+ 'DT_INICIO' +';'+ 'PLANO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('CODIGO').AsString;
          P2  := qryDados.FIELDBYNAME('NOME').AsString;
          P3  := qryDados.FIELDBYNAME('TIPO').AsString;
          P4  := qryDados.FIELDBYNAME('CPF').AsString;
          P5  := qryDados.FIELDBYNAME('COD_ADTV').AsString;
          P6  := qryDados.FIELDBYNAME('NOME_ADT').AsString;
          P7  := qryDados.FIELDBYNAME('DESCRICAO').AsString;
          P8  := qryDados.FIELDBYNAME('DT_INCL').AsString;
          P9  := qryDados.FIELDBYNAME('DT_INICIO').AsString;
          P10  := qryDados.FIELDBYNAME('PLANO').AsString;

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

   if frmPrincipal.lblCodRelatorio.Caption = '15' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_SEG' +';'+ 'COD_OUTRO2' +';'+ 'COD_FAM' +';'+ 'UTILIZACAO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P2  := qryDados.FIELDBYNAME('COD_OUTRO2').AsString;
          P3  := qryDados.FIELDBYNAME('COD_FAM').AsString;
          P4  := qryDados.FIELDBYNAME('UTILIZACAO').AsString;

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

   if frmPrincipal.lblCodRelatorio.Caption = '32' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_BOL' +';'+ 'COD_SEG' +';'+ 'NOME' +';'+ 'TIPO' +';'+ 'VALOR' +';'+ 'EMPRESA' +';'+ 'COD_OUTRO2');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_BOL').AsString;
          P2  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P3  := qryDados.FIELDBYNAME('NOME').AsString;
          P4  := qryDados.FIELDBYNAME('TIPO').AsString;
          P5  := qryDados.FIELDBYNAME('VALOR').AsString;
          P6  := qryDados.FIELDBYNAME('EMPRESA').AsString;
          P7  := qryDados.FIELDBYNAME('COD_OUTRO2').AsString;

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

   if frmPrincipal.lblCodRelatorio.Caption = '47' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'EMPRESA' +';'+ 'COMP_PAGAMENTO' +';'+ 'PAGO_ODONTO' +';'+ 'ADITIVO_PLANO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('EMPRESA').AsString;
          P2  := qryDados.FIELDBYNAME('COMP_PAGAMENTO').AsString;
          P3  := qryDados.FIELDBYNAME('PAGO_ODONTO').AsString;
          P4  := qryDados.FIELDBYNAME('ADITIVO_PLANO').AsString;

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

   if frmPrincipal.lblCodRelatorio.Caption = '89' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'TELEFONE' +';'+ 'CELULAR' +';'+ 'TELEFONE2' +';'+ 'DEMAIS_TEL' +';'+ 'COD_SEG' +';'+ 'NOME' +';'+ 'DIA_VENC' +';'+ 'QTD_BOL' +';'+ 'VALOR' +';'+ 'DT_RETORNO' +';'+ 'DT_VENC_RECENTE' +';'+ 'DT_LIG' +';'+ 'LOGIN' +';'+ 'TIPO_LIG');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('TELEFONE').AsString;
          P2  := qryDados.FIELDBYNAME('CELULAR').AsString;
          P3  := qryDados.FIELDBYNAME('TELEFONE2').AsString;
          P4  := qryDados.FIELDBYNAME('DEMAIS_TEL').AsString;
          P5  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P6  := qryDados.FIELDBYNAME('NOME').AsString;
          P7  := qryDados.FIELDBYNAME('DIA_VENC').AsString;
          P8  := qryDados.FIELDBYNAME('QTD_BOL').AsString;
          P9  := qryDados.FIELDBYNAME('VALOR').AsString;
          P10  := qryDados.FIELDBYNAME('DT_RETORNO').AsString;
          P11 := qryDados.FIELDBYNAME('DT_VENC_RECENTE').AsString;
          P12 := qryDados.FIELDBYNAME('DT_LIG').AsString;
          P13 := qryDados.FIELDBYNAME('LOGIN').AsString;
          P14 := qryDados.FIELDBYNAME('TIPO_LIG').AsString;

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

procedure TfrmRelatorio6.qryDadosAfterOpen(DataSet: TDataSet);
begin
  frmPrincipal.DimensionarGrid(dbgDados,Self);

  if frmPrincipal.lblCodRelatorio.Caption = '13' then
  begin
    TCurrencyField(qryDados.FieldByName('CODIGO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('TIPO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('CPF')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_ADTV')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DT_INCL')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DT_INICIO')).Alignment := taCenter;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '15' then
  begin
    TCurrencyField(qryDados.FieldByName('COD_SEG')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_OUTRO2')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_FAM')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('UTILIZACAO')).Alignment := taRightJustify;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '32' then
  begin
    TCurrencyField(qryDados.FieldByName('COD_BOL')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_SEG')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('TIPO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('VALOR')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('COD_OUTRO2')).Alignment := taCenter;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '47' then
  begin
    TCurrencyField(qryDados.FieldByName('COMP_PAGAMENTO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('PAGO_ODONTO')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('ADITIVO_PLANO')).Alignment := taCenter;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '89' then
  begin
    TCurrencyField(qryDados.FieldByName('TELEFONE')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('CELULAR')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('TELEFONE2')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DEMAIS_TEL')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_SEG')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DIA_VENC')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('QTD_BOL')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DT_RETORNO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DT_VENC_RECENTE')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DT_LIG')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('LOGIN')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('TIPO_LIG')).Alignment := taCenter;
  end;
end;

end.
