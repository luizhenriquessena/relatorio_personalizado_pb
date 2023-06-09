unit Relatorio5;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, Grids, DBGrids, StdCtrls, Buttons, Mask;

type
  TfrmRelatorio5 = class(TForm)
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
    qryExec: TADOQuery;
    GroupBox2: TGroupBox;
    Med_MesInicio: TMaskEdit;
    Label2: TLabel;
    Med_MesFinal: TMaskEdit;
    chkPrestador: TCheckBox;
    cbxTipoPrestador: TComboBox;
    edtCodigoPrest: TEdit;
    edtNomePrest: TEdit;
    lblCodigoPrestador: TLabel;
    lblTipoEnt: TLabel;
    lblTotal2: TLabel;
    lblRTotal2: TLabel;
    gpbPeriodo: TGroupBox;
    rdbMesPag: TRadioButton;
    rdbDtLib: TRadioButton;
    rdbDtReal: TRadioButton;
    Med_DataInicio: TMaskEdit;
    Med_DataFinal: TMaskEdit;
    edtCodigoAMB: TEdit;
    lblTitCodigoAMB: TLabel;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure Med_MesInicioExit(Sender: TObject);
    procedure Med_MesFinalExit(Sender: TObject);
    procedure chkPrestadorClick(Sender: TObject);
    procedure btnGerarClick(Sender: TObject);
    procedure edtCodigoPrestKeyPress(Sender: TObject; var Key: Char);
    procedure edtCodigoPrestExit(Sender: TObject);
    procedure btnExportarClick(Sender: TObject);
    procedure cbxTipoPrestadorClick(Sender: TObject);
    procedure rdbMesPagClick(Sender: TObject);
    procedure Med_DataFinalExit(Sender: TObject);
    procedure Med_DataInicioExit(Sender: TObject);
    procedure rdbDtLibClick(Sender: TObject);
    procedure rdbDtRealClick(Sender: TObject);
    procedure qryDadosAfterOpen(DataSet: TDataSet);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmRelatorio5: TfrmRelatorio5;
  MesInicio, MesFinal : String;
  parametro: String;
implementation

uses Principal, DateUtils;

{$R *.dfm}

procedure TfrmRelatorio5.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
qryExec.Close;
qryDados.Close;



end;

procedure TfrmRelatorio5.FormShow(Sender: TObject);
begin
  lblCodigoPrestador.Caption := 'X';
  edtCodigoPrest.Text := '';
  edtNomePrest.Text := '';
  Med_MesInicio.Text := '  /    ';
  Med_MesFinal.Text  := '  /    ';
  Med_DataInicio.Text := '  /  /    ';
  Med_DataFinal.Text  := '  /  /    ';
  chkPrestador.Visible := True;
  cbxTipoPrestador.Visible := True;
  edtCodigoPrest.Visible := True;
  edtNomePrest.Visible := True;
  gpbPeriodo.Enabled := False;
  gpbPeriodo.Visible := False;
  chkPrestador.Checked := False;
  edtCodigoAMB.Clear;
  edtCodigoAMB.Visible := False;
  lblTitCodigoAMB.Visible := False;

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

  if frmPrincipal.lblCodRelatorio.Caption = '23' then
  begin
  chkPrestador.Visible := False;
  cbxTipoPrestador.Visible := False;
  edtCodigoPrest.Visible := False;
  edtNomePrest.Visible := False;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '11' then
  begin
  gpbPeriodo.Enabled := True;
  gpbPeriodo.Visible := True;
  lblTitCodigoAMB.Visible := True;
  edtCodigoAMB.Visible := True;
  end;

end;

procedure TfrmRelatorio5.Med_MesInicioExit(Sender: TObject);
var
d_Data : TDate;
begin
 if Med_MesInicio.Text <> '  /    ' then
 begin
  try
    d_Data := StrToDate('01/'+Med_MesInicio.Text);
    MesInicio := FormatDateTime('yyyy/mm',d_Data);
  except
    ShowMessage('Data base inv�lida');
    Med_MesInicio.Text := '  /    ';
    Med_MesInicio.SetFocus;
  end;
 end; 
end;

procedure TfrmRelatorio5.Med_MesFinalExit(Sender: TObject);
var
d_Data : TDate;
begin
 if Med_MesFinal.Text <> '  /    ' then
 begin
  try
    d_Data := StrToDate('01/'+Med_MesFinal.Text);
    MesFinal := FormatDateTime('yyyy/mm',d_Data);
    if (frmPrincipal.lblCodRelatorio.Caption = '11') or (frmPrincipal.lblCodRelatorio.Caption = '112') then
    MesFinal := MesFinal+'-2';
  except
    ShowMessage('Data base inv�lida');
    Med_MesFinal.Text := '  /    ';
    Med_MesFinal.SetFocus;
  end;
 end;
end;

procedure TfrmRelatorio5.chkPrestadorClick(Sender: TObject);
begin
  if chkPrestador.Checked = True then
  begin
  edtCodigoPrest.Color := clWindow;
  edtCodigoPrest.Enabled := True;
  edtNomePrest.Enabled := True;
    if frmPrincipal.lblCodRelatorio.Caption = '112' then
    begin
      cbxTipoPrestador.ItemIndex := 0;
      lblTipoEnt.Caption := '2';
      cbxTipoPrestador.Enabled := False;
      Abort;
    end;
  cbxTipoPrestador.Enabled := True;
  cbxTipoPrestador.ItemIndex := -1;
  end else
  begin
  cbxTipoPrestador.ItemIndex := -1;
  cbxTipoPrestador.Enabled := False;  
  edtCodigoPrest.Text := '';
  edtNomePrest.Text := '';
  lblCodigoPrestador.Caption := '0';
  lblTipoEnt.Caption := 'X';
  edtCodigoPrest.Color := clSilver;
  edtCodigoPrest.Enabled := False;
  edtNomePrest.Enabled := False;
  end;

end;

procedure TfrmRelatorio5.btnGerarClick(Sender: TObject);
var
Total : Currency;
begin
  parametro := '';
  Total := 0;
  if frmPrincipal.lblCodRelatorio.Caption = '11' then
  begin
    // MES PAGAMENTO
    if gpbPeriodo.Enabled = True and rdbMesPag.Checked = True then
    begin
      Med_MesInicio.OnExit(Sender);
      Med_MesFinal.OnExit(Sender);
      if (Med_MesInicio.Text <> '  /    ') and (Med_MesFinal.Text <> '  /    ') then
      begin
        parametro := 'MesPagInicio: '+MesInicio + ' MesPagFinal: '+MesFinal;
        with qryDados do
        begin
        Close;
        SQL.Clear;
        SQL.Add(frmPrincipal.memSQL.Text);
        SQL.Add('AND fa.MES_PGTO >= '+#39+MesInicio+#39);
        SQL.Add('AND fa.MES_PGTO <= '+#39+MesFinal+#39);

          if (chkPrestador.Checked = true) and (lblCodigoPrestador.Caption <> 'X') then
          begin
            SQL.Add('AND nvl(fa.TIPO_ENT,s.TIPO_ENT) = '+#39+lblTipoEnt.Caption+#39+' AND nvl(fa.COD_ENT,s.COD_ENT) = '+#39+lblCodigoPrestador.Caption+#39);
            parametro := parametro + ' - Prestador: '+ lblTipoEnt.Caption + ' - ' + lblCodigoPrestador.Caption;
          end;

        if length(edtCodigoAMB.Text)>4 then
        begin
          SQL.Add('and s.cod_amb like ('+#39+'%'+edtCodigoAMB.Text+'%'+#39+')');
          parametro := parametro + ' ' + edtCodigoAMB.Text;
        end;

        Open;

        end;

        lblRTotal.Caption := IntToStr(qryDados.RecordCount);

        with qryExec do
        begin
          Close;
          SQL.Clear;
          SQL.Add('select sum(valor_pago) valor_pago from (');
          SQL.Add(qryDados.SQL.Text);
          SQL.Add(')');
          Open;
        end;
        Total := qryExec.FieldByName('VALOR_PAGO').AsCurrency;
        lblRTotal2.Caption := FormatFloat('#,0.00',Total);

       {
        while not qryDados.Eof do
        begin
         Total := Total + qryDados.FieldByName('VALOR_PAGO').AsCurrency;
         qryDados.Next;
        end;

        lblRTotal2.Caption := FormatFloat('#,0.00',Total);
       }
       
        with qryExec do // INSERIR LOG DE VISUALIZA��O
        begin
        Close;
        SQL.Clear;
        SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, PARAMETRO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'V'+#39+', '+#39+parametro+#39+', SYSDATE)');
        ExecSQL;
        end;
      end else
      begin
      ShowMessage('Favor informar o per�odo.');
      end;
    end;

    // DATA LIBERACAO  OU EXECU��O
    if gpbPeriodo.Enabled = True and rdbMesPag.Checked = False then
    begin
      if (Med_DataInicio.Text <> '  /  /    ') and (Med_DataFinal.Text <> '  /  /    ') then
      begin
        with qryDados do
        begin
        Close;
        SQL.Clear;
        SQL.Add(frmPrincipal.memSQL.Text);
          if rdbDtLib.Checked = True then //Inserir condi��o de periodo por data de libera��o
          begin
            SQL.Add('AND s.dt_lib >= '+#39+Med_DataInicio.Text+#39);
            SQL.Add('AND s.dt_lib <= '+#39+Med_DataFinal.Text+#39);
            parametro := 'DtLibInicio: '+Med_DataInicio.Text + ' DtLibFinal: '+Med_DataFinal.Text;
          end;

          if rdbDtReal.Checked = True then //Inserir condi��o de periodo por data de realiza��o
          begin
            SQL.Add('AND s.dt_exec >= '+#39+Med_DataInicio.Text+#39);
            SQL.Add('AND s.dt_exec <= '+#39+Med_DataFinal.Text+#39);
            parametro := 'DtRealInicio: '+Med_DataInicio.Text + ' DtRealFinal: '+Med_DataFinal.Text;
          end;

          if (chkPrestador.Checked = true) and (lblCodigoPrestador.Caption <> 'X') then
          begin
            SQL.Add('AND h.TIPO_ENT = '+#39+lblTipoEnt.Caption+#39+' AND h.COD_ENT = '+#39+lblCodigoPrestador.Caption+#39);
            parametro := parametro + ' - Prestador: '+ lblTipoEnt.Caption + ' - ' + lblCodigoPrestador.Caption;
          end;

        if length(edtCodigoAMB.Text)>4 then
        begin
          SQL.Add('and s.cod_amb like ('+#39+'%'+edtCodigoAMB.Text+'%'+#39+')');
          parametro := parametro + ' ' + edtCodigoAMB.Text;
        end;

        parametro := parametro + ' ' + edtCodigoAMB.Text;
        Open;
        end;

        lblRTotal.Caption := IntToStr(qryDados.RecordCount);

        with qryExec do
        begin
          Close;
          SQL.Clear;
          SQL.Add('select sum(valor_pago) valor_pago from (');
          SQL.Add(qryDados.SQL.Text);
          SQL.Add(')');
          Open;
        end;
        Total := qryExec.FieldByName('VALOR_PAGO').AsCurrency;
        lblRTotal2.Caption := FormatFloat('#,0.00',Total);

        {
        while not qryDados.Eof do
        begin
         Total := Total + qryDados.FieldByName('VALOR_PAGO').AsCurrency;
         qryDados.Next;
        end;

        lblRTotal2.Caption := FormatFloat('#,0.00',Total);
        }
        with qryExec do // INSERIR LOG DE VISUALIZA��O
        begin
        Close;
        SQL.Clear;
        SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, PARAMETRO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'V'+#39+', '+#39+parametro+#39+', SYSDATE)');
        ExecSQL;
        end;
      end else
      begin
      ShowMessage('Favor informar o per�odo.');
      end;
    end;

  end
  else
  begin
    if (Med_MesInicio.Text <> '  /    ') and (Med_MesFinal.Text <> '  /    ') then
    begin
      Med_MesInicio.OnExit(Sender);
      Med_MesFinal.OnExit(Sender);
      parametro := 'MInicio: '+MesInicio + ' MFinal: '+MesFinal;
      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);
      SQL.Text := StringReplace(SQL.Text,'&ANOMES_INICIO',MesInicio,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&ANOMES_FINAL',MesFinal,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco

        if (chkPrestador.Checked = true) and (lblCodigoPrestador.Caption <> 'X') then
        begin
          if frmPrincipal.lblCodRelatorio.Caption = '112' then
          begin
            SQL.Add(' WHERE SUBSTR(PRESTADOR,1,5) = '+#39+lblCodigoPrestador.Caption+#39);
            parametro := parametro + ' - Prestador: '+ lblCodigoPrestador.Caption;
          end;

          if frmPrincipal.lblCodRelatorio.Caption = '135' then
          begin
            SQL.Add('AND V.TIPO_ENT = '+#39+lblTipoEnt.Caption+#39+' AND V.COD_ENT = '+#39+lblCodigoPrestador.Caption+#39);
            parametro := parametro + ' - Prestador: '+ lblTipoEnt.Caption + ' - ' + lblCodigoPrestador.Caption;
          end;

        end;

        if frmPrincipal.lblCodRelatorio.Caption = '112' then
          SQL.Add('         ORDER BY PRESTADOR, ESPECIALIDADE, REDE');

        if frmPrincipal.lblCodRelatorio.Caption = '135' then
        begin
          SQL.Add('          GROUP BY P.COD_ENT, P.NOME, V.DT_ENTREGA, V.DT_ENTREGA_ORIGINAL, V.DT_VENC, V.MES_PGTO');
          SQL.Add(' ORDER BY PRESTADOR');
        end;

      Open;
      end;

      lblRTotal.Caption := IntToStr(qryDados.RecordCount);

      if frmPrincipal.lblCodRelatorio.Caption = '23' then
      begin
        while not qryDados.Eof do
        begin
          Total := Total + qryDados.FieldByName('VALOR').AsCurrency;
          qryDados.Next;
        end;
        lblRTotal2.Caption := FormatFloat('#,0.00',Total);
      end;

      if frmPrincipal.lblCodRelatorio.Caption = '112' then
        lblRTotal2.Visible := False;

      with qryExec do // INSERIR LOG DE VISUALIZA��O
      begin
      Close;
      SQL.Clear;
      SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, PARAMETRO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'V'+#39+', '+#39+parametro+#39+', SYSDATE)');
      ExecSQL;
      end;

    end else
    begin
    ShowMessage('Favor informar o per�odo.');
    end;
  end;
end;

procedure TfrmRelatorio5.edtCodigoPrestKeyPress(Sender: TObject;
  var Key: Char);
begin
    if not (key in ['0'..'9',Chr(8), Chr(3), Chr(22), Chr(24), Chr(44)]) then
      key := #0;
end;

procedure TfrmRelatorio5.edtCodigoPrestExit(Sender: TObject);
begin
  if chkPrestador.Checked = True then
  begin
    if edtCodigoPrest.Text <> '' then
    begin
      with qryExec do
      begin
      Close;
      SQL.Clear;
      SQL.Add('SELECT COD_ENT, NOME FROM TI.VW_PRESTADORES WHERE COD_ENT = '+#39+edtCodigoPrest.Text+#39+' AND TIPO_ENT = '+#39+lblTipoEnt.Caption+#39);
      Open;
      end;

      if qryExec.RecordCount > 0 then
      begin
      edtCodigoPrest.Text := qryExec.FieldByName('COD_ENT').AsString;
      lblCodigoPrestador.Caption := qryExec.FieldByName('COD_ENT').AsString;
      edtNomePrest.Text := qryExec.FieldByName('NOME').AsString;
      end else
      begin
      ShowMessage('C�digo do Prestador Inv�lido');
      end;
    end;
  end;
end;

procedure TfrmRelatorio5.btnExportarClick(Sender: TObject);
var
F: TextFile;
P1, P2, P3, P4 , P5 , P6, P7, P8, P9, P10, P11, P12, P13, P14, P15, P16, P17, P18, P19, P20, P21, P22, P23, P24,
P25, P26, P27, P28, P29, P30, P31, P32, P33, CLinha  : String; //Variaveis comuns
begin
 if qryDados.RecordCount > 0 then
 begin
  if frmPrincipal.lblCodRelatorio.Caption = '11' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'PAGAMENTO' +';'+ 'PRESTADOR' +';'+ 'SENHA_LIB' +';'+ 'DT_LIB' +';'+ 'DT_EXEC' +';'+ 'DT_REALIZACAO' +';'+ 'PROCEDIMENTO' +';'+ 'VALOR_COBRADO' +';'+ 'VALOR_GLOSADO' +';'+ 'VALOR_PAGO' +';'+ 'ESPECIALIDADE' +';'+ 'MEDICO_EXEC' +';'+ 'MEDICO_SOL' +';'+ 'BENEFICIARIO' +';'+ 'PLANO' +';'+ 'IDADE' +';'+ 'SEXO' +';'+ 'TIPO_BENEF' +';'+ 'FATURA' +';'+ 'ENDERECO_BENEF' +';'+ 'BAIRRO_BENEF' +';'+ 'CIDADE_BENEF' +';'+ 'COD_EMPRESA' +';'+ 'EMPRESA' +';'+ 'SITUACAO' +';'+ 'TIPO_DESC' +';'+ 'SITUACAO_BENEF' +';'+ 'CONCILIADO' +';'+ 'SEGMENTACAO' +';'+ 'LOGIN' +';'+ 'TIPO_PREST' +';'+ 'DT_LIB_SENHA_CRE' +';'+ 'DT_SIST');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('PAGAMENTO').AsString;
          P2  := qryDados.FIELDBYNAME('PRESTADOR').AsString;
          P3  := qryDados.FIELDBYNAME('SENHA_LIB').AsString;
          P4  := qryDados.FIELDBYNAME('DT_LIB').AsString;
          P5  := qryDados.FIELDBYNAME('DT_EXEC').AsString;
          P6  := qryDados.FIELDBYNAME('DT_REALIZACAO').AsString;
          P7  := qryDados.FIELDBYNAME('PROCEDIMENTO').AsString;
          P8  := qryDados.FIELDBYNAME('VALOR_COBRADO').AsString;
          P9  := qryDados.FIELDBYNAME('VALOR_GLOSADO').AsString;
          P10  := qryDados.FIELDBYNAME('VALOR_PAGO').AsString;
          P11 := qryDados.FIELDBYNAME('ESPECIALIDADE').AsString;
          P12 := qryDados.FIELDBYNAME('MEDICO_EXEC').AsString;
          P13 := qryDados.FIELDBYNAME('MEDICO_SOL').AsString;
          P14 := qryDados.FIELDBYNAME('BENEFICIARIO').AsString;
          P15 := qryDados.FIELDBYNAME('PLANO').AsString;
          P16 := qryDados.FIELDBYNAME('IDADE').AsString;
          P17 := qryDados.FIELDBYNAME('SEXO').AsString;
          P18 := qryDados.FIELDBYNAME('TIPO').AsString;
          P19 := qryDados.FIELDBYNAME('FATURA').AsString;
          P20 := qryDados.FIELDBYNAME('ENDERECO_BENEF').AsString;
          P21 := qryDados.FIELDBYNAME('BAIRRO_BENEF').AsString;
          P22 := qryDados.FIELDBYNAME('CIDADE_BENEF').AsString;
          P23 := qryDados.FIELDBYNAME('COD_EMPRESA').AsString;
          P24 := qryDados.FIELDBYNAME('EMPRESA').AsString;
          P25 := qryDados.FIELDBYNAME('SITUACAO').AsString;
          P26 := qryDados.FIELDBYNAME('TIPO_DESC').AsString;
          P27 := qryDados.FIELDBYNAME('SITUACAO_BENEFICIARIO').AsString;
          P28 := qryDados.FIELDBYNAME('CONCILIADO').AsString;
          P29 := qryDados.FIELDBYNAME('SEGMENTACAO').AsString;
          P30 := qryDados.FIELDBYNAME('LOGIN').AsString;
          P31 := qryDados.FIELDBYNAME('TIPO_PREST').AsString;
          P32 := qryDados.FIELDBYNAME('DT_LIB_SENHA_CRE').AsString;
          P33 := qryDados.FIELDBYNAME('DT_SIST').AsString;

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
          cLinha := cLinha + P28 +';';
          cLinha := cLinha + P29 +';';
          cLinha := cLinha + P30 +';';
          cLinha := cLinha + P31 +';';
          cLinha := cLinha + P32 +';';
          cLinha := cLinha + P33 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '23' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'TIPO' +';'+ 'COD_FORN' +';'+ 'NOME_FORN' +';'+ 'COD_ADTV' +';'+ 'NOME_ADTV' +';'+ 'VALOR' +';'+ 'COD_SEG' +';'+ 'NOME_SEG' +';'+ 'MES_PGTO' +';'+ 'VIGENCIA' +';'+ 'CPF' +';'+ 'DT_ADESAO' +';'+ 'DATA_ENTRADA_BENEF' +';'+ 'DATA_ENTRADA_ADTV' +';'+ 'CPF_TITULAR' +';'+ 'TIPO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('TIPO').AsString;
          P2  := qryDados.FIELDBYNAME('COD_FORN').AsString;
          P3  := qryDados.FIELDBYNAME('NOME_FORN').AsString;
          P4  := qryDados.FIELDBYNAME('COD_ADTV').AsString;
          P5  := qryDados.FIELDBYNAME('NOME_ADTV').AsString;
          P6  := qryDados.FIELDBYNAME('VALOR').AsString;
          P7  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P8  := qryDados.FIELDBYNAME('NOME_SEG').AsString;
          P9  := qryDados.FIELDBYNAME('MES_PGTO').AsString;
          P10 := qryDados.FIELDBYNAME('VIGENCIA').AsString;
          P11 := qryDados.FIELDBYNAME('CPF').AsString;
          P12 := qryDados.FIELDBYNAME('DT_ADESAO').AsString;
          P13 := qryDados.FIELDBYNAME('DATA_ENTRADA_BENEF').AsString;
          P14 := qryDados.FIELDBYNAME('DATA_ENTRADA_ADTV').AsString;
          P15 := qryDados.FIELDBYNAME('CPF_TITULAR').AsString;
          P16 := qryDados.FIELDBYNAME('TIPO_BENEF').AsString;

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

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '112' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'PRESTADOR' +';'+ 'REDE' +';'+ 'ESPECIALIDADE' +';'+ 'QUANTIDADE' +';'+ 'VALOR_PAGO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('PRESTADOR').AsString;
          P2  := qryDados.FIELDBYNAME('REDE').AsString;
          P3  := qryDados.FIELDBYNAME('ESPECIALIDADE').AsString;
          P4  := qryDados.FIELDBYNAME('QUANTIDADE').AsString;
          P5  := qryDados.FIELDBYNAME('VALOR_PAGO').AsString;

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

  if frmPrincipal.lblCodRelatorio.Caption = '135' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_ENT' +';'+ 'PRESTADOR' +';'+ 'VALOR_APRESENTADO' +';'+ 'VALOR_REAL' +';'+ 'GLOSA' +';'+ 'VALOR_PAGO' +';'+ 'DT_ENTREGA' +';'+ 'DT_ENTREGA_ORIGINAL' +';'+ 'DT_VENC' +';'+ 'MES_PGTO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_ENT').AsString;
          P2  := qryDados.FIELDBYNAME('PRESTADOR').AsString;
          P3  := qryDados.FIELDBYNAME('VALOR_APRESENTADO').AsString;
          P4  := qryDados.FIELDBYNAME('VALOR_REAL').AsString;
          P5  := qryDados.FIELDBYNAME('GLOSA').AsString;
          P6  := qryDados.FIELDBYNAME('VALOR_PAGO').AsString;
          P7  := qryDados.FIELDBYNAME('DT_ENTREGA').AsString;
          P8  := qryDados.FIELDBYNAME('DT_ENTREGA_ORIGINAL').AsString;
          P9  := qryDados.FIELDBYNAME('DT_VENC').AsString;
          P10  := qryDados.FIELDBYNAME('MES_PGTO').AsString;

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

  with qryExec do // INSERIR LOG DE EXPORTA��O
  begin
  Close;
  SQL.Clear;
  SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, PARAMETRO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'E'+#39+', '+#39+parametro+#39+', SYSDATE)');
  ExecSQL;
  end;

 end;
end;

procedure TfrmRelatorio5.cbxTipoPrestadorClick(Sender: TObject);
begin
 if cbxTipoPrestador.ItemIndex = 0 then
  lblTipoEnt.Caption := '2';

 if cbxTipoPrestador.ItemIndex = 1 then
  lblTipoEnt.Caption := '1';

 if cbxTipoPrestador.ItemIndex = 2 then
  lblTipoEnt.Caption := '0';
end;

procedure TfrmRelatorio5.rdbMesPagClick(Sender: TObject);
begin
  if rdbMesPag.Checked = True then
  begin
  Med_DataInicio.Enabled := False;
  Med_DataInicio.Visible := False;
  Med_DataFinal.Enabled  := False;
  Med_DataFinal.Visible  := False;
  Med_MesInicio.Enabled  := True;
  Med_MesInicio.Visible  := True;
  Med_MesFinal.Enabled   := True;
  Med_MesFinal.Visible   := True;
  GroupBox2.Caption      := 'M�s de Pagamento';
  Med_DataInicio.Text    := '  /  /    ';
  Med_DataFinal.Text     := '  /  /    ';
  end;
end;

procedure TfrmRelatorio5.Med_DataFinalExit(Sender: TObject);
var d_Data : TDateTime;
begin
 if Med_DataFinal.Text <> '  /  /    ' then
 begin
  try
    d_Data := StrToDateTime(Med_DataFinal.Text);
    Med_DataFinal.Text := FormatDateTime('dd/mm/yyyy',d_Data);
  except
    ShowMessage('Data inv�lida');
    Med_DataFinal.Text := '  /  /    ';
    Med_DataFinal.SetFocus;
  end;
 end;
end;

procedure TfrmRelatorio5.Med_DataInicioExit(Sender: TObject);
var d_Data : TDateTime;
begin
 if Med_DataInicio.Text <> '  /  /    ' then
 begin
   try
    d_Data := StrToDateTime(Med_DataInicio.Text);
    Med_DataInicio.Text := FormatDateTime('dd/mm/yyyy',d_Data);
  except
    ShowMessage('Data inv�lida');
    Med_DataInicio.Text := '  /  /    ';
    Med_DataInicio.SetFocus;
  end;
 end;
end;

procedure TfrmRelatorio5.rdbDtLibClick(Sender: TObject);
begin
  if rdbDtLib.Checked = True then
  begin
  Med_DataInicio.Enabled := True;
  Med_DataInicio.Visible := True;
  Med_DataFinal.Enabled  := True;
  Med_DataFinal.Visible  := True;
  Med_MesInicio.Enabled  := False;
  Med_MesInicio.Visible  := False;
  Med_MesFinal.Enabled   := False;
  Med_MesFinal.Visible   := False;
  GroupBox2.Caption      := 'Data de Libera��o';
  Med_MesInicio.Text     := '  /    ';
  Med_MesFinal.Text      := '  /    ';
  end;
end;

procedure TfrmRelatorio5.rdbDtRealClick(Sender: TObject);
begin
  if rdbDtReal.Checked = True then
  begin
  Med_DataInicio.Enabled := True;
  Med_DataInicio.Visible := True;
  Med_DataFinal.Enabled  := True;
  Med_DataFinal.Visible  := True;
  Med_MesInicio.Enabled  := False;
  Med_MesInicio.Visible  := False;
  Med_MesFinal.Enabled   := False;
  Med_MesFinal.Visible   := False;
  GroupBox2.Caption      := 'Data de Realiza��o';
  Med_MesInicio.Text     := '  /    ';
  Med_MesFinal.Text      := '  /    ';  
  end;
end;

procedure TfrmRelatorio5.qryDadosAfterOpen(DataSet: TDataSet);
begin
  frmPrincipal.DimensionarGrid(dbgDados,Self);

  if frmPrincipal.lblCodRelatorio.Caption = '11' then
  begin
    TCurrencyField(qryDados.FieldByName('PAGAMENTO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('SENHA_LIB')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DT_LIB')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DT_REALIZACAO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('VALOR_COBRADO')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('VALOR_COBRADO')).Alignment := taRightJustify;
    TCurrencyField(qryDados.FieldByName('VALOR_GLOSADO')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('VALOR_GLOSADO')).Alignment := taRightJustify;
    TCurrencyField(qryDados.FieldByName('VALOR_PAGO')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('VALOR_PAGO')).Alignment := taRightJustify;
    TCurrencyField(qryDados.FieldByName('IDADE')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('SEXO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('FATURA')).Alignment := taCenter;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '23' then
  begin
    TCurrencyField(qryDados.FieldByName('TIPO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_FORN')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_ADTV')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('VALOR')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('VALOR')).Alignment := taRightJustify;
    TCurrencyField(qryDados.FieldByName('COD_SEG')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('MES_PGTO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('VIGENCIA')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('CPF')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DT_ADESAO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DATA_ENTRADA_BENEF')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DATA_ENTRADA_ADTV')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('CPF_TITULAR')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('TIPO_BENEF')).Alignment := taCenter;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '135' then
  begin
    TCurrencyField(qryDados.FieldByName('COD_ENT')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('VALOR_APRESENTADO')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('VALOR_REAL')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('GLOSA')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('VALOR_PAGO')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('DT_ENTREGA')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DT_ENTREGA_ORIGINAL')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DT_VENC')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('MES_PGTO')).Alignment := taCenter;
  end;
end;

end.
