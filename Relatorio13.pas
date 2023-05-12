unit Relatorio13;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, StdCtrls, Mask, Grids, DBGrids, Buttons;

type
  TfrmRelatorio13 = class(TForm)
    Label1: TLabel;
    dbgDados: TDBGrid;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    Label2: TLabel;
    Med_DataFinal: TMaskEdit;
    Med_DataInicio: TMaskEdit;
    qryDados: TADOQuery;
    dsDados: TDataSource;
    SaveDialog1: TSaveDialog;
    btnExportar: TSpeedButton;
    btnImprimir: TSpeedButton;
    lblTotal: TLabel;
    lblRTotal: TLabel;
    qryExec: TADOQuery;
    btnGerar: TSpeedButton;
    edtNomeEmp: TEdit;
    edtCodigoEmp: TEdit;
    lblCodigoEmpresa: TLabel;
    chk_Empresa: TCheckBox;
    chkUnidade: TCheckBox;
    lblCodigoUnidade: TLabel;
    edtNomeUnidade: TEdit;
    edtCodigoUnidade: TEdit;
    lblMes: TLabel;
    Med_Mes: TMaskEdit;
    chkPlano: TCheckBox;
    lblCodigoPlano: TLabel;
    cbxPlano: TComboBox;
    edtPlano: TEdit;
    Label3: TLabel;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure Med_DataInicioExit(Sender: TObject);
    procedure Med_DataFinalExit(Sender: TObject);
    procedure edtCodigoEmpExit(Sender: TObject);
    procedure edtCodigoEmpKeyPress(Sender: TObject; var Key: Char);
    procedure btnGerarClick(Sender: TObject);
    procedure btnExportarClick(Sender: TObject);
    procedure chk_EmpresaClick(Sender: TObject);
    procedure btnImprimirClick(Sender: TObject);
    procedure qryDadosAfterOpen(DataSet: TDataSet);
    procedure chkUnidadeClick(Sender: TObject);
    procedure edtCodigoUnidadeKeyPress(Sender: TObject; var Key: Char);
    procedure edtCodigoUnidadeExit(Sender: TObject);
    procedure Med_MesKeyPress(Sender: TObject; var Key: Char);
    procedure Med_MesExit(Sender: TObject);
    procedure chkPlanoClick(Sender: TObject);
    procedure cbxPlanoChange(Sender: TObject);
    procedure edtPlanoKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmRelatorio13: TfrmRelatorio13;
  parametro: String;
  media, receita, despesa, comissao : Variant;
  Mes : String;
implementation

{$R *.dfm}
uses MaskUtils, Math, Principal, Relat65, DataModule, DateUtils;
procedure TfrmRelatorio13.FormClose(Sender: TObject; var Action: TCloseAction);
begin
qryExec.Close;
qryDados.Close;
DM.qryRelat65.Close;
end;

procedure TfrmRelatorio13.FormShow(Sender: TObject);
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

  chk_Empresa.Visible := True;
  edtCodigoEmp.Visible := True;
  edtNomeEmp.Visible := True;
  Med_DataInicio.Text := '  /  /    ';
  Med_DataFinal.Text := '  /  /    ';
  lblRTotal.Caption := '0';
  edtCodigoEmp.Text:= '';
  edtNomeEmp.Text:= '';
  lblCodigoEmpresa.Caption:= '0';
  chk_Empresa.Checked:= False;
  edtCodigoEmp.Color:= clMenu;
  chk_Empresa.Enabled := True;
  dsDados.DataSet := qryDados;
  chkUnidade.Visible := False;
  chkUnidade.Checked := False;
  chkUnidade.Enabled := False;
  chkPlano.Enabled := False;
  chkPlano.Checked := False;
  chkPlano.Visible := False;
  cbxPlano.Enabled := False;
  cbxPlano.Visible := False;
  edtPlano.Enabled := False;
  edtPlano.Visible := False;
  edtPlano.Clear;
  edtCodigoUnidade.Visible := False;
  edtCodigoUnidade.Clear;
  edtCodigoUnidade.Enabled := False;
  edtNomeUnidade.Visible := False;
  edtNomeUnidade.Clear;
  edtNomeUnidade.Enabled := False;
  lblCodigoUnidade.Caption := '0';
  lblCodigoPlano.Caption := '-1';
  lblMes.Visible := False;
  Med_Mes.Visible := False;
  Med_Mes.Text := '__/____';

  if (frmPrincipal.lblCodRelatorio.Caption = '51') or (frmPrincipal.lblCodRelatorio.Caption = '52') or (frmPrincipal.lblCodRelatorio.Caption = '56') then
  begin
    chk_Empresa.Caption := 'Vendedor ';
  end else
  begin
    chk_Empresa.Caption := 'Empresa ';
  end;

  if (frmPrincipal.lblCodRelatorio.Caption = '44') or (frmPrincipal.lblCodRelatorio.Caption = '129')  or (frmPrincipal.lblCodRelatorio.Caption = '131')
     or (frmPrincipal.lblCodRelatorio.Caption = '138') then
  begin
    chkUnidade.Visible := True;
    edtCodigoUnidade.Visible := True;
    edtNomeUnidade.Visible := True;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '65' then
  begin
    chk_Empresa.Checked := False;
    //chk_EmpresaClick(Sender);
    //chk_Empresa.Enabled := False;
    dsDados.DataSet := DM.qryRelat65;
    chkUnidade.Visible := True;
    edtCodigoUnidade.Visible := True;
    edtNomeUnidade.Visible := True;
    chkPlano.Checked := False;
    chkPlano.Visible := True;
    chkPlano.Enabled := True;
    //cbxPlano.Visible := True;
    edtPlano.Visible := True;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '141' then
  begin
    chk_Empresa.Checked := True;
    chk_EmpresaClick(Sender);
    chk_Empresa.Enabled := False;
    chkUnidade.Visible := True;
    edtCodigoUnidade.Visible := True;
    edtNomeUnidade.Visible := True;
    chkPlano.Checked := False;
    chkPlano.Visible := True;
    chkPlano.Enabled := True;
    //cbxPlano.Visible := True;
    edtPlano.Visible := True;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '134' then
  begin
    GroupBox2.Visible := False;
    lblMes.Visible := True;
    Med_Mes.Visible := True;
  end;

end;

procedure TfrmRelatorio13.Med_DataInicioExit(Sender: TObject);
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

procedure TfrmRelatorio13.Med_DataFinalExit(Sender: TObject);
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

procedure TfrmRelatorio13.edtCodigoEmpExit(Sender: TObject);
begin
  if edtCodigoEmp.Text <> '' then
  begin
    if (frmPrincipal.lblCodRelatorio.Caption = '51') or (frmPrincipal.lblCodRelatorio.Caption = '52') or  (frmPrincipal.lblCodRelatorio.Caption = '56') then
    begin
      with qryExec do
      begin
      Close;
      SQL.Clear;
      SQL.Add('SELECT COD_VEND, NOME FROM PLATINUM.IM_VEND WHERE COD_VEND = '+ edtCodigoEmp.Text);
      Open;
      end;

      if qryExec.RecordCount > 0 then
      begin
      edtNomeEmp.Text := qryExec.FieldByName('NOME').AsString;
      lblCodigoEmpresa.Caption := qryExec.FieldByName('COD_VEND').AsString;
      end else
      begin
      edtCodigoEmp.Text := '';
      lblCodigoEmpresa.Caption := '0';
      edtNomeEmp.Text := '';
      ShowMessage('Código do Vendedor Inválido');
      end;
    end else
    begin
      with qryExec do
      begin
      Close;
      SQL.Clear;
      SQL.Add('SELECT COD_SET, DESCRICAO FROM ');
        if (frmPrincipal.lblCodRelatorio.Caption = '129') then
        SQL.Add('PLATINUM.IM_SETOR WHERE COD_SET = '+ edtCodigoEmp.Text)
      else
        SQL.Add('IM_SETOR WHERE COD_SET = '+ edtCodigoEmp.Text);
      Open;
      end;

      if qryExec.RecordCount > 0 then
      begin
      edtNomeEmp.Text := qryExec.FieldByName('DESCRICAO').AsString;
      lblCodigoEmpresa.Caption := qryExec.FieldByName('COD_SET').AsString;
      end else
      begin
      edtCodigoEmp.Text := '';
      lblCodigoEmpresa.Caption := '0';
      edtNomeEmp.Text := '';
      ShowMessage('Código da Empresa Inválido');
      end;
    end;
  end;
end;

procedure TfrmRelatorio13.edtCodigoEmpKeyPress(Sender: TObject; var Key: Char);
begin
    if not (key in ['0'..'9',Chr(8), Chr(3), Chr(22), Chr(24), Chr(44)]) then
      key := #0;
end;

procedure TfrmRelatorio13.btnGerarClick(Sender: TObject);
begin
  edtCodigoEmp.OnExit(Sender);
  if frmPrincipal.lblCodRelatorio.Caption <> '134' then
  begin
    Med_DataInicio.OnExit(Sender);
    Med_DataFinal.OnExit(Sender);
  end else
  begin
    Med_Mes.OnExit(Sender);
  end;
  parametro := '';

  if frmPrincipal.lblCodRelatorio.Caption = '44' then
  begin
    if (Med_DataInicio.Text <> '  /  /    ') and (Med_DataFinal.Text <> '  /  /    ') then
    begin
    parametro := 'DtInicio: '+Med_DataInicio.Text + ' DtFinal: ' + Med_DataFinal.Text;
      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);

        if (chk_Empresa.Checked = true) and (lblCodigoEmpresa.Caption <> '0') then
        begin
          SQL.Add('AND S.COD_SET = '+#39+lblCodigoEmpresa.Caption+#39);
          parametro := parametro + ' Empresa: '+ lblCodigoEmpresa.Caption;
        end;

        if (chkUnidade.Checked = true) and (lblCodigoUnidade.Caption <> '0') then
        begin
          SQL.Add('AND S.COD_SETSET = '+#39+lblCodigoUnidade.Caption+#39);
          parametro := parametro + ' Unidade: '+ lblCodigoUnidade.Caption;
        end;

      SQL.Add('AND   S.DT_INCL BETWEEN '+#39+Med_DataInicio.Text+#39+' AND '+#39+Med_DataFinal.Text+#39+') Z, (SELECT * FROM IM_VEND C WHERE C.TIPO = '+#39+'S'+#39+') SUP, (SELECT * FROM IM_SGVEA LG WHERE LG.TIPO_VEND = '+#39+'S'+#39+') LG,');

      SQL.Add('        (SELECT c.cod_seg, sum(c.valor) comissao');
      SQL.Add('          FROM im_cvend c');
      SQL.Add('         WHERE c.pag_ext = '+#39+'P'+#39);
      SQL.Add('         GROUP BY c.cod_seg) CC,');
      SQL.Add('        (SELECT S.COD_SEG');
      SQL.Add('           FROM IM_BOLAB B');
      SQL.Add('           JOIN IM_DBOL  D ON B.COD_BOL   = D.COD_BOL');
      SQL.Add('           JOIN IM_SEG   S ON D.COD_SEG   = S.COD_SEG');
      SQL.Add('          WHERE S.DT_INCL   BETWEEN '+#39+Med_DataInicio.Text+#39+' AND '+#39+Med_DataFinal.Text+#39);
      SQL.Add('            AND D.TIPO   = 0');
      SQL.Add('            AND B.CONCIL = '+#39+'F'+#39);
      SQL.Add('            AND B.DT_CANC IS NULL');
      SQL.Add('            AND B.DT_VENC <= SYSDATE-10');
      SQL.Add('            AND S.CANCELED = '+#39+'F'+#39);
      SQL.Add('            AND S.COD_SET <> 3) INAD');

      SQL.Add('WHERE Z.COD_VENDA  = LG.COD_VENDA (+)');
      SQL.Add('AND   LG.COD_VEND  = SUP.COD_VEND (+)');
      SQL.Add('AND   Z.COD_SEG    = CC.COD_SEG (+)');
      SQL.Add('  AND Z.COD_SEG    = INAD.COD_SEG (+)');

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
    ShowMessage('Favor informar todos os parâmetros.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '51' then
  begin
    if (Med_DataInicio.Text <> '  /  /    ') and (Med_DataFinal.Text <> '  /  /    ') then
    begin
    parametro := 'DtInicio: '+Med_DataInicio.Text + ' DtFinal: ' + Med_DataFinal.Text;
      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);

        if (chk_Empresa.Checked = true) and (lblCodigoEmpresa.Caption <> '0') then
        begin
        SQL.Add('AND C.COD_VEND = '+#39+lblCodigoEmpresa.Caption+#39);
        parametro := parametro + ' Vendedor: '+ lblCodigoEmpresa.Caption;
        end;

      SQL.Add('AND   D.DT_CANC >= '+#39+Med_DataInicio.Text+#39+' AND D.DT_CANC <= '+#39+Med_DataFinal.Text+#39+') Z');
      SQL.Add('WHERE NVL((RECEITA - DESPESA),0) > 0');
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
    ShowMessage('Favor informar todos os parâmetros.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '52' then
  begin
    if (Med_DataInicio.Text <> '  /  /    ') and (Med_DataFinal.Text <> '  /  /    ') then
    begin
    parametro := 'DtInicio: '+Med_DataInicio.Text + ' DtFinal: ' + Med_DataFinal.Text;
      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);

        if (chk_Empresa.Checked = true) and (lblCodigoEmpresa.Caption <> '0') then
        begin
        SQL.Add('AND C.COD_VEND = '+#39+lblCodigoEmpresa.Caption+#39);
        parametro := parametro + ' Vendedor: '+ lblCodigoEmpresa.Caption;
        end;

      SQL.Add('AND   S.DT_INCL >= '+#39+Med_DataInicio.Text+#39+' AND S.DT_INCL <= '+#39+Med_DataFinal.Text+#39);
      SQL.Add('ORDER BY S.DT_INCL, S.COD_SEG');
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
    ShowMessage('Favor informar todos os parâmetros.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '65' then
  begin
    if (Med_DataInicio.Text <> '  /  /    ') and (Med_DataFinal.Text <> '  /  /    ') then
    begin
    parametro := 'DtInicio: '+Med_DataInicio.Text + ' DtFinal: ' + Med_DataFinal.Text + ' Empresa: ' + lblCodigoEmpresa.Caption;
      with DM.qryRelat65 do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);
      SQL.Text := StringReplace(SQL.Text,'&DATA1',Med_DataInicio.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&DATA2',Med_DataFinal.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco

        if chk_Empresa.Checked = true then
        begin
          SQL.Text := StringReplace(SQL.Text,'--AND S.COD_SET = &EMPRESA','AND S.COD_SET = '+#39+lblCodigoEmpresa.Caption+#39,[rfReplaceAll]);
          SQL.Text := StringReplace(SQL.Text,'--AND B.COD_SET = &EMPRESA','AND B.COD_SET = '+#39+lblCodigoEmpresa.Caption+#39,[rfReplaceAll]);
          SQL.Text := StringReplace(SQL.Text,'--AND A.COD_SET = &EMPRESA','AND A.COD_SET = '+#39+lblCodigoEmpresa.Caption+#39,[rfReplaceAll]);
          parametro := parametro + ' Empresa: '+ lblCodigoUnidade.Caption;
        end;

        if (chkUnidade.Checked = true) and (lblCodigoUnidade.Caption <> '0') then
        begin
          SQL.Text := StringReplace(SQL.Text,'--AND A.COD_SETSET = &UNIDADE','AND A.COD_SETSET = '+#39+lblCodigoUnidade.Caption+#39,[rfReplaceAll]);
          SQL.Text := StringReplace(SQL.Text,'--AND B.COD_SETSE = &UNIDADE','AND B.COD_SETSE = '+#39+lblCodigoUnidade.Caption+#39,[rfReplaceAll]);
          SQL.Text := StringReplace(SQL.Text,'--AND S.COD_SETSET = &UNIDADE','AND S.COD_SETSET = '+#39+lblCodigoUnidade.Caption+#39,[rfReplaceAll]);
          SQL.Text := StringReplace(SQL.Text,'--AND S.COD_SETSET = &UNIDADE','AND S.COD_SETSET = '+#39+lblCodigoUnidade.Caption+#39,[rfReplaceAll]);
          parametro := parametro + ' Unidade: '+ lblCodigoUnidade.Caption;
        end;

        if chkPlano.Checked = true then
        begin
          SQL.Text := StringReplace(SQL.Text,'--AND S.COD_PLANO = &PLANO','AND S.COD_PLANO IN ('+edtPlano.Text+')',[rfReplaceAll]);
          parametro := parametro + ' Plano: '+ lblCodigoPlano.Caption;
        end;

      Open;
      end;

      with qryExec do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL2.Text);
      SQL.Text := StringReplace(SQL.Text,'&DATA1',Med_DataInicio.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&DATA2',Med_DataFinal.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&EMPRESA',lblCodigoEmpresa.Caption,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco

        if (chkUnidade.Checked = true) and (lblCodigoUnidade.Caption <> '0') then
        begin
          SQL.Text := StringReplace(SQL.Text,'--AND COD_SETSET = &UNIDADE','AND COD_SETSET = '+#39+lblCodigoUnidade.Caption+#39,[rfReplaceAll]);
          SQL.Text := StringReplace(SQL.Text,'--AND B.COD_SETSE = &UNIDADE','AND B.COD_SETSE = '+#39+lblCodigoUnidade.Caption+#39,[rfReplaceAll]);
          SQL.Text := StringReplace(SQL.Text,'--AND S.COD_SETSET = &UNIDADE','AND S.COD_SETSET = '+#39+lblCodigoUnidade.Caption+#39,[rfReplaceAll]);
          parametro := parametro + ' Unidade: '+ lblCodigoUnidade.Caption;
        end;

        if (chkPlano.Checked = true) and (lblCodigoPlano.Caption <> '-1') then
        begin
          SQL.Text := StringReplace(SQL.Text,'--AND S.COD_PLANO = &PLANO','AND S.COD_PLANO IN ( '+edtPlano.Text+ ')',[rfReplaceAll]);
          parametro := parametro + ' Plano: '+ lblCodigoPlano.Caption;
        end;

      Open;
      end;

    lblRTotal.Caption := IntToStr(DM.qryRelat65.RecordCount);
    media    := qryExec.FieldByName('VIDAS').AsVariant;
    receita  := qryExec.FieldByName('RECEITA').AsVariant;
    despesa  := qryExec.FieldByName('SINISTRALIDADE').AsVariant;
    comissao := qryExec.FieldByName('COMISSAO').AsVariant;

      with qryExec do // INSERIR LOG DE VISUALIZAÇÃO
      begin
      Close;
      SQL.Clear;
      SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, PARAMETRO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'V'+#39+','+#39+parametro+#39+', SYSDATE)');
      ExecSQL;
      end;
    end else
    begin
    ShowMessage('Favor informar todos os parâmetros.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '74' then
  begin
    if (Med_DataInicio.Text <> '  /  /    ') and (Med_DataFinal.Text <> '  /  /    ') then
    begin
    parametro := 'DtInicio: '+Med_DataInicio.Text + ' DtFinal: ' + Med_DataFinal.Text;

      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);

        if (chk_Empresa.Checked = true) and (lblCodigoEmpresa.Caption <> '0') then
        begin
          SQL.Add('AND S.COD_SET = '+#39+lblCodigoEmpresa.Caption+#39);
          parametro := parametro + ' Empresa: '+ lblCodigoEmpresa.Caption;
        end;

      SQL.Add('GROUP BY S.DT_INCL, S.DATA_ENTRADA, S.COD_SEG, S.COD_OUTRO, S.COD_OUTRO2, S.NOME, S.CPF, DECODE(S.CANCELED,'+#39+'F'+#39+','+#39+'ATIVO'+#39+','+#39+'T'+#39+','+#39+'CANCELADO'+#39+'),M.DESCRICAO, U.CGC, U.COD_SETSET || '+#39+' - '+#39+' || U.NOME, V.COD_SUBCOM, V.MES_GERADO,');
      SQL.Add('         DECODE(S.TIPO,'+#39+'E'+#39+','+#39+'TITULAR'+#39+','+#39+'D'+#39+','+#39+'DEPENDENTE'+#39+'), NVL2(T.COD_SEG,T.COD_SEG || '+#39+' - '+#39+' || T.NOME,NULL), S.LOGIN');

      SQL.Text := StringReplace(SQL.Text,'&DATA1',Med_DataInicio.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&DATA2',Med_DataFinal.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
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
    ShowMessage('Favor informar todos os parâmetros.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '82' then
  begin
    if (Med_DataInicio.Text <> '  /  /    ') and (Med_DataFinal.Text <> '  /  /    ') then
    begin
    parametro := 'DtInicio: '+Med_DataInicio.Text + ' DtFinal: ' + Med_DataFinal.Text;
      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);

      SQL.Text := StringReplace(SQL.Text,'&DATA1',Med_DataInicio.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&DATA2',Med_DataFinal.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco

        if (chk_Empresa.Checked = true) and (lblCodigoEmpresa.Caption <> '0') then
        begin
        SQL.Add('AND S.COD_SET = '+#39+lblCodigoEmpresa.Caption+#39);
        parametro := parametro + ' Empresa: '+ lblCodigoEmpresa.Caption;
        end;

      SQL.Add('ORDER BY E.COD_SET || '+#39+' - '+#39+' ||  E.DESCRICAO, DT_CANC, COD_SEG');
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
    ShowMessage('Favor informar todos os parâmetros.');
    end;
  end;

 if (frmPrincipal.lblCodRelatorio.Caption = '111') or (frmPrincipal.lblCodRelatorio.Caption = '145') then
  begin
    if (Med_DataInicio.Text <> '  /  /    ') and (Med_DataFinal.Text <> '  /  /    ') then
    begin
    parametro := 'DtInicio: '+Med_DataInicio.Text + ' DtFinal: ' + Med_DataFinal.Text;

      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);

        if (chk_Empresa.Checked = true) and (lblCodigoEmpresa.Caption <> '0') then
        begin
          SQL.Add('AND E.COD_SET = '+#39+lblCodigoEmpresa.Caption+#39);
          parametro := parametro + ' Empresa: '+ lblCodigoEmpresa.Caption;
        end;

      SQL.Text := StringReplace(SQL.Text,'&DATA1',Med_DataInicio.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&DATA2',Med_DataFinal.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
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
    ShowMessage('Favor informar todos os parâmetros.');
    end;
  end;


  if frmPrincipal.lblCodRelatorio.Caption = '56' then
  begin
    if (Med_DataInicio.Text <> '  /  /    ') and (Med_DataFinal.Text <> '  /  /    ') then
    begin
    parametro := 'DtInicio: '+Med_DataInicio.Text + ' DtFinal: ' + Med_DataFinal.Text;
      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);

      SQL.Text := StringReplace(SQL.Text,'&DATA1',Med_DataInicio.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&DATA2',Med_DataFinal.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco

        if (chk_Empresa.Checked = true) and (lblCodigoEmpresa.Caption <> '0') then
        begin
        SQL.Add('AND VEND.COD_VENDEDOR = '+#39+lblCodigoEmpresa.Caption+#39);
        parametro := parametro + ' Vendedor: '+ lblCodigoEmpresa.Caption;
        end;

     SQL.Add(' GROUP BY DECODE(S.CANCELED,'+#39+'F'+#39+','+#39+'ATIVO'+#39+','+#39+'T'+#39+','+#39+'CANCELADO'+#39+'), S.MAT, DECODE(S.TIPO,'+#39+'E'+#39+','+#39+'TITULAR'+#39+','+#39+'D'+#39+','+#39+'DEPENDENTE'+#39+'), S.DT_INCL, S.COD_SEG, S.NOME, S.COD_FORMA, F.DESCRICAO, VEND.COD_VENDEDOR, VEND.VENDEDOR, SUP.COD_SUPERVISOR, SUP.SUPERVISOR, DIR.COD_DIRETOR, DIR.DIRETOR, GC.DESCRICAO');
     SQL.Add('      ,NVL(M.P_MENSALIDADE,'+#39+'NÃO PAGO'+#39+'), M.VALOR, NVL(M2.P_MENSALIDADE,'+#39+'NÃO PAGO'+#39+'), M2.VALOR, NVL(M3.P_MENSALIDADE,'+#39+'NÃO PAGO'+#39+'), M3.VALOR');
     SQL.Add('      ,NVL(M4.P_MENSALIDADE,'+#39+'NÃO PAGO'+#39+'), M4.VALOR, NVL(M5.P_MENSALIDADE,'+#39+'NÃO PAGO'+#39+'), M5.VALOR, NVL(M6.P_MENSALIDADE,'+#39+'NÃO PAGO'+#39+'), M6.VALOR');
     SQL.Add('      ,CASE');
     SQL.Add('        WHEN T.DT_INCL = S.DT_INCL THEN '+#39+'ADESÃO DE CONTRATO'+#39);
     SQL.Add('         WHEN T.DT_INCL < S.DT_INCL THEN '+#39+'INCLUSÃO NO CONTRATO'+#39);
     SQL.Add('          WHEN T.DT_INCL > S.DT_INCL THEN '+#39+'VERIFICAR'+#39);
     SQL.Add('           END');
     SQL.Add('      ,E.COD_SET || '+#39+' - '+#39+' || E.SIGLA');
     SQL.Add('      ,UNID.COD_SETSET || '+#39+' - '+#39+' || UNID.NOME');
     SQL.Add('      ,LCAR.DT_INICIO');
     SQL.Add('      ,TI.RECEITA_BENEF_PLA(S.COD_SEG)');
     SQL.Add('      ,TI.UTILIZACAO_BENEF_PLA(S.COD_OUTRO2)');
     SQL.Add('      ,S.DT_CANC');
     SQL.Add('      ,COM.COMISSAO');
     SQL.Add('      ,P.NOME');
     SQL.Add('      ,GS.COD_GSEG ||'+#39+' - '+#39+'|| GS.DESCRICAO');
     SQL.Add('      ,PLATINUM.IDADE(S.COD_SEG,SYSDATE)');
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
    ShowMessage('Favor informar todos os parâmetros.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '129' then
  begin
    if (Med_DataInicio.Text <> '  /  /    ') and (Med_DataFinal.Text <> '  /  /    ') then
    begin
    parametro := 'DtInicio: '+Med_DataInicio.Text + ' DtFinal: ' + Med_DataFinal.Text;
      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);

        if (chk_Empresa.Checked = true) and (lblCodigoEmpresa.Caption <> '0') then
        begin
          SQL.Add('AND S.COD_SET = '+#39+lblCodigoEmpresa.Caption+#39);
          parametro := parametro + ' Empresa: '+ lblCodigoEmpresa.Caption;
        end;

        if (chkUnidade.Checked = true) and (lblCodigoUnidade.Caption <> '0') then
        begin
          SQL.Add('AND S.COD_SETSET = '+#39+lblCodigoUnidade.Caption+#39);
          parametro := parametro + ' Unidade: '+ lblCodigoUnidade.Caption;
        end;

      SQL.Text := StringReplace(SQL.Text,'&DATA1',Med_DataInicio.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&DATA2',Med_DataFinal.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco

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
    ShowMessage('Favor informar todos os parâmetros.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '131' then
  begin
    if (Med_DataInicio.Text <> '  /  /    ') and (Med_DataFinal.Text <> '  /  /    ') then
    begin
    parametro := 'DtInicio: '+Med_DataInicio.Text + ' DtFinal: ' + Med_DataFinal.Text;
      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);
      SQL.Text := StringReplace(SQL.Text,'&DATA1',Med_DataInicio.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&DATA2',Med_DataFinal.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
        if (chk_Empresa.Checked = true) and (lblCodigoEmpresa.Caption <> '0') then
        begin
          SQL.Add('AND S.COD_SET = '+#39+lblCodigoEmpresa.Caption+#39);
          parametro := parametro + ' Empresa: '+ lblCodigoEmpresa.Caption;
        end;
        if (chkUnidade.Checked = true) and (lblCodigoUnidade.Caption <> '0') then
        begin
          SQL.Add('AND S.COD_SETSET = '+#39+lblCodigoUnidade.Caption+#39);
          parametro := parametro + ' Unidade: '+ lblCodigoUnidade.Caption;
        end;
     SQL.Add('  group by e.cod_set,e.sigla,u.cod_setset,u.nome_fant,u.cgc,s.dt_incl,decode(s.tipo,'+#39+'E'+#39+','+#39+'TITULAR'+#39+','+#39+'D'+#39+','+#39+'DEPENDENTE'+#39+')');
     SQL.Add('        ,s.cod_seg,s.nome,s.cpf,decode(s.canceled,'+#39+'T'+#39+','+#39+'CANCELADO'+#39+','+#39+'F'+#39+','+#39+'ATIVO'+#39+'),d.cod_bol,b.dt_venc,b.comp');
     SQL.Add('        ,d.nmes,decode(b.concil,'+#39+'T'+#39+','+#39+'SIM'+#39+','+#39+'F'+#39+','+#39+'NÃO'+#39+'),b.dt_concil,b.dt_concil_real');
     SQL.Add('        order by s.cod_seg, d.cod_bol');
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
    ShowMessage('Favor informar todos os parâmetros.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '132' then
  begin
    if (Med_DataInicio.Text <> '  /  /    ') and (Med_DataFinal.Text <> '  /  /    ') then
    begin
    parametro := 'DtInicio: '+Med_DataInicio.Text + ' DtFinal: ' + Med_DataFinal.Text;

      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);

        if (chk_Empresa.Checked = true) and (lblCodigoEmpresa.Caption <> '0') then
        begin
          SQL.Text := StringReplace(SQL.Text,'&EMPRESA',lblCodigoEmpresa.Caption,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
          parametro := parametro + ' Empresa: '+ lblCodigoEmpresa.Caption;
        end;
        
      SQL.Text := StringReplace(SQL.Text,'&DATA1',Med_DataInicio.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&DATA2',Med_DataFinal.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
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
    ShowMessage('Favor informar todos os parâmetros.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '134' then
  begin
    if Med_Mes.Text <> '  /    ' then
    begin
    parametro := 'Mês: '+Med_Mes.Text;
      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);

      SQL.Text := StringReplace(SQL.Text,'&MESPAG',Mes+'%',[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco

        if (chk_Empresa.Checked = true) and (lblCodigoEmpresa.Caption <> '0') then
        begin
          SQL.Text := StringReplace(SQL.Text,'&EMPRESA',lblCodigoEmpresa.Caption,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
          parametro := parametro + ' Empresa: '+ lblCodigoEmpresa.Caption;
        end;

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
    ShowMessage('Favor informar todos os parâmetros.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '138' then
  begin
    if (Med_DataInicio.Text <> '  /  /    ') and (Med_DataFinal.Text <> '  /  /    ') then
    begin
    parametro := 'DtInicio: '+Med_DataInicio.Text + ' DtFinal: ' + Med_DataFinal.Text;
      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);
      SQL.Text := StringReplace(SQL.Text,'&DATA1',Med_DataInicio.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&DATA2',Med_DataFinal.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
        if (chk_Empresa.Checked = true) and (lblCodigoEmpresa.Caption <> '0') then
        begin
          SQL.Add('AND B.COD_SET = '+#39+lblCodigoEmpresa.Caption+#39);
          parametro := parametro + ' Empresa: '+ lblCodigoEmpresa.Caption;
        end;
        if (chkUnidade.Checked = true) and (lblCodigoUnidade.Caption <> '0') then
        begin
          SQL.Add('AND B.COD_SETSE = '+#39+lblCodigoUnidade.Caption+#39);
          parametro := parametro + ' Unidade: '+ lblCodigoUnidade.Caption;
        end;
     SQL.Add('  group by e.cod_set,e.sigla,u.cod_setset,u.nome_fant,u.cgc,s.dt_incl,decode(s.tipo,'+#39+'E'+#39+','+#39+'TITULAR'+#39+','+#39+'D'+#39+','+#39+'DEPENDENTE'+#39+')');
     SQL.Add('        ,s.cod_seg,s.nome,s.cpf,decode(s.canceled,'+#39+'T'+#39+','+#39+'CANCELADO'+#39+','+#39+'F'+#39+','+#39+'ATIVO'+#39+'),d.cod_bol,b.dt_venc,b.dia_venc,b.comp');
     SQL.Add('        ,d.nmes,decode(b.concil,'+#39+'T'+#39+','+#39+'SIM'+#39+','+#39+'F'+#39+','+#39+'NÃO'+#39+'),b.dt_concil,b.dt_concil_real');
     SQL.Add('        order by s.cod_seg, d.cod_bol');
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
    ShowMessage('Favor informar todos os parâmetros.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '141' then
  begin
    if (Med_DataInicio.Text <> '  /  /    ') and (Med_DataFinal.Text <> '  /  /    ') then
    begin
    parametro := 'DtInicio: '+Med_DataInicio.Text + ' DtFinal: ' + Med_DataFinal.Text + ' Empresa: ' + lblCodigoEmpresa.Caption;
      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);
      SQL.Text := StringReplace(SQL.Text,'&DATA1',Med_DataInicio.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&DATA2',Med_DataFinal.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&EMPRESA',lblCodigoEmpresa.Caption,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco

        if (chkUnidade.Checked = true) and (lblCodigoUnidade.Caption <> '0') then
        begin
          SQL.Text := StringReplace(SQL.Text,'--AND A.COD_SETSET = &UNIDADE','AND A.COD_SETSET = '+#39+lblCodigoUnidade.Caption+#39,[rfReplaceAll]);
          SQL.Text := StringReplace(SQL.Text,'--AND B.COD_SETSE = &UNIDADE','AND B.COD_SETSE = '+#39+lblCodigoUnidade.Caption+#39,[rfReplaceAll]);
          SQL.Text := StringReplace(SQL.Text,'--AND S.COD_SETSET = &UNIDADE','AND S.COD_SETSET = '+#39+lblCodigoUnidade.Caption+#39,[rfReplaceAll]);
          SQL.Text := StringReplace(SQL.Text,'--AND S.COD_SETSET = &UNIDADE','AND S.COD_SETSET = '+#39+lblCodigoUnidade.Caption+#39,[rfReplaceAll]);
          parametro := parametro + ' Unidade: '+ lblCodigoUnidade.Caption;
        end;

        if chkPlano.Checked = true then
        begin
          SQL.Text := StringReplace(SQL.Text,'--AND S.COD_PLANO = &PLANO','AND S.COD_PLANO IN ('+edtPlano.Text+')',[rfReplaceAll]);
          parametro := parametro + ' Plano: '+ lblCodigoPlano.Caption;
        end;

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
    ShowMessage('Favor informar todos os parâmetros.');
    end;
  end;

end;

procedure TfrmRelatorio13.btnExportarClick(Sender: TObject);
var
F: TextFile;
P1, P2, P3, P4, P5, P6, P7, P8, P9, P10, P11, P12, P13, P14, P15, P16,
P17, P18, P19, P20, P21, P22, P23, P24, P25, P26, P27, P28, P29, P30, P31, P32, P33, P34, P35, P36, P37, P38, P39, cLinha  : String;
begin

with qryExec do // INSERIR LOG DE EXPORTAÇÃO
begin
Close;
SQL.Clear;
SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, PARAMETRO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'E'+#39+', '+#39+parametro+#39+', SYSDATE)');
ExecSQL;
end;

  if frmPrincipal.lblCodRelatorio.Caption = '65' then
  begin
    if DM.qryRelat65.RecordCount > 0 then
    begin
      if SaveDialog1.Execute then
      begin
        AssignFile(F, SaveDialog1.FileName);
        Rewrite (F);
        cLinha := '';
        Writeln (F,cLinha + 'ANO_MES' +';'+ 'VIDAS' +';'+ 'SINISTRALIDADE' +';'+ 'COMISSAO' +';'+ 'RECEITA');
        DM.qryRelat65.First;
        while not DM.qryRelat65.Eof do
        begin
          P1  := DM.qryRelat65.FIELDBYNAME('ANO_MES').AsString;
          P2  := DM.qryRelat65.FIELDBYNAME('VIDAS').AsString;
          P3  := DM.qryRelat65.FIELDBYNAME('SINISTRALIDADE').AsString;
          P4  := DM.qryRelat65.FIELDBYNAME('COMISSAO').AsString;
          P5  := DM.qryRelat65.FIELDBYNAME('RECEITA').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';

          Writeln(F,cLinha);
          DM.qryRelat65.Next;
        end;
        CloseFile (F);
        ShowMessage('Arquivo Gerado com sucesso.');
      end;
    end;
    Abort;
  end;

 if qryDados.RecordCount > 0 then
 begin

  if frmPrincipal.lblCodRelatorio.Caption = '44' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'SITUACAO' +';'+ 'COD_SEG' +';'+ 'CPF' +';'+ 'BENEFICIARIO' +';'+ 'EMPRESA' +';'+ 'UNIDADE' +';'+ 'PLANO' +';'+ 'EST_CIVIL' +';'+ 'TIPO' +';'+ 'SEXO' +';'+ 'BENEF_VENC' +';'+ 'EMP_VENC' +';'+ 'DT_NASC' +';'+ 'DT_INCL' +';'+ 'DATA_ENTRADA' +';'+ 'IDADE' +';'+ 'TAB_COM' +';'+ 'PARENT' +';'+ 'MENSALIDADE' +';'+ 'PAI' +';'+ 'MAE' +';'+ 'COD_VEND' +';'+ 'NOME_VEND' +';'+ 'COD_SUP' +';'+ 'NOME_SUP' +';'+ 'UTILIZACAO' +';'+ 'RECEITA' +';'+ 'DESPESA' +';'+ 'CNS' +';'+ 'TAB_PRECO' +';'+ 'COMISSAO' +';'+ 'INADIMPLENCIA' +';'+ 'DT_CANC_SEG' +';'+ 'DT_CANC_EMP' +';'+ 'LOGIN');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('SITUACAO').AsString;
          P2  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P3  := qryDados.FIELDBYNAME('CPF').AsString;
          P4  := qryDados.FIELDBYNAME('BENEFICIARIO').AsString;
          P5  := qryDados.FIELDBYNAME('EMPRESA').AsString;
          P6  := qryDados.FIELDBYNAME('UNIDADE').AsString;
          P7  := qryDados.FIELDBYNAME('PLANO').AsString;
          P8  := qryDados.FIELDBYNAME('EST_CIVIL').AsString;
          P9  := qryDados.FIELDBYNAME('TIPO').AsString;
          P10 := qryDados.FIELDBYNAME('SEXO').AsString;
          P11 := qryDados.FIELDBYNAME('BENEF_VENC').AsString;
          P12 := qryDados.FIELDBYNAME('EMP_VENC').AsString;
          P13 := qryDados.FIELDBYNAME('DT_NASC').AsString;
          P14 := qryDados.FIELDBYNAME('DT_INCL').AsString;
          P15 := qryDados.FIELDBYNAME('DATA_ENTRADA').AsString;
          P16 := qryDados.FIELDBYNAME('IDADE').AsString;
          P17 := qryDados.FIELDBYNAME('TAB_COM').AsString;
          P18 := qryDados.FIELDBYNAME('PARENT').AsString;
          P19 := qryDados.FIELDBYNAME('MENSALIDADE').AsString;
          P20 := qryDados.FIELDBYNAME('PAI').AsString;
          P21 := qryDados.FIELDBYNAME('MAE').AsString;
          P22 := qryDados.FIELDBYNAME('COD_VEND').AsString;
          P23 := qryDados.FIELDBYNAME('NOME_VEND').AsString;
          P24 := qryDados.FIELDBYNAME('COD_SUP').AsString;
          P25 := qryDados.FIELDBYNAME('NOME_SUP').AsString;
          P26 := qryDados.FIELDBYNAME('UTILIZACAO').AsString;
          P27 := qryDados.FIELDBYNAME('RECEITA').AsString;
          P28 := qryDados.FIELDBYNAME('DESPESA').AsString;
          P29 := qryDados.FIELDBYNAME('CNS').AsString;
          P30 := qryDados.FIELDBYNAME('TAB_PRECO').AsString;
          P31 := qryDados.FIELDBYNAME('COMISSAO').AsString;
          P32 := qryDados.FIELDBYNAME('INADIMPLENTE').AsString;
          P33 := qryDados.FIELDBYNAME('DT_CANC_SEG').AsString;
          P34 := qryDados.FIELDBYNAME('DT_CANC_EMP').AsString;
          P35 := qryDados.FIELDBYNAME('LOGIN').AsString;

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
          cLinha := cLinha + P34 +';';
          cLinha := cLinha + P35 +';';          

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '51' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_VEND' +';'+ 'VENDEDOR' +';'+ 'COD_SEG' +';'+ 'COD_OUTRO2' +';'+ 'NOME' +';'+ 'DT_CANC' +';'+ 'EMPRESA' +';'+ 'MAT' +';'+ 'DT_INCL' +';'+ 'TEL' +';'+ 'CELULAR' +';'+ 'DESCRICAO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_VEND').AsString;
          P2  := qryDados.FIELDBYNAME('VENDEDOR').AsString;
          P3  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P4  := qryDados.FIELDBYNAME('COD_OUTRO2').AsString;
          P5  := qryDados.FIELDBYNAME('NOME').AsString;
          P6  := qryDados.FIELDBYNAME('DT_CANC').AsString;
          P7  := qryDados.FIELDBYNAME('EMPRESA').AsString;
          P8  := qryDados.FIELDBYNAME('MAT').AsString;
          P9  := qryDados.FIELDBYNAME('DT_INCL').AsString;
          P10 := qryDados.FIELDBYNAME('TEL').AsString;
          P11 := qryDados.FIELDBYNAME('CELULAR').AsString;
          P12 := qryDados.FIELDBYNAME('DESCRICAO').AsString;

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

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '52' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'CODIGO' +';'+ 'CORRETOR' +';'+ 'INCLUSAO' +';'+ 'COD_BENEF' +';'+ 'BENEFICIARIO' +';'+ 'TIPO' +';'+ 'CONTRATO' +';'+ 'FORMA_PAG' +';'+ 'TELEFONE' +';'+ 'CELULAR');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('CODIGO').AsString;
          P2  := qryDados.FIELDBYNAME('CORRETOR').AsString;
          P3  := qryDados.FIELDBYNAME('INCLUSAO').AsString;
          P4  := qryDados.FIELDBYNAME('COD_BENEF').AsString;
          P5  := qryDados.FIELDBYNAME('BENEFICIARIO').AsString;
          P6  := qryDados.FIELDBYNAME('TIPO').AsString;
          P7  := qryDados.FIELDBYNAME('CONTRATO').AsString;
          P8  := qryDados.FIELDBYNAME('FORMA_PAG').AsString;
          P9  := qryDados.FIELDBYNAME('TELEFONE').AsString;
          P10 := qryDados.FIELDBYNAME('CELULAR').AsString;

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

  if frmPrincipal.lblCodRelatorio.Caption = '74' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'DT_ADESAO' +';'+ 'DT_CADASTRO' +';'+ 'COD_SEG' +';'+ 'COD_ALTERNATIVO' +';'+ 'COD_ALTERNATIVO2' +';'+ 'NOME' +';'+ 'CPF' +';'+ 'SITUACAO' +';'+ 'MOTIVO_CANC' +';'+ 'CNPJ_UNID' +';'+ 'UNIDADE' +';'+ 'TAB_COMISSAO' +';'+ 'MES_GERADO' +';'+ 'TIPO' +';'+ 'TITULAR' +';'+ 'LOGIN');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('ADESAO').AsString;
          P2  := qryDados.FIELDBYNAME('DT_CADASTRO').AsString;
          P3  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P4  := qryDados.FIELDBYNAME('COD_ALTERNATIVO').AsString;
          P5  := qryDados.FIELDBYNAME('COD_ALTERNATIVO2').AsString;
          P6  := qryDados.FIELDBYNAME('NOME').AsString;
          P7  := qryDados.FIELDBYNAME('CPF').AsString;
          P8  := qryDados.FIELDBYNAME('SITUACAO').AsString;
          P9  := qryDados.FIELDBYNAME('MOTIVO').AsString;
          P10 := qryDados.FIELDBYNAME('CNPJ_UNIDADE').AsString;
          P11 := qryDados.FIELDBYNAME('UNIDADE').AsString;
          P12 := qryDados.FIELDBYNAME('TAB_COMISSAO').AsString;
          P13 := qryDados.FIELDBYNAME('MES_GERADO').AsString;
          P14 := qryDados.FIELDBYNAME('TIPO').AsString;
          P15 := qryDados.FIELDBYNAME('TITULAR').AsString;
          P16 := qryDados.FIELDBYNAME('LOGIN').AsString;

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

  if frmPrincipal.lblCodRelatorio.Caption = '82' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_SEG' +';'+ 'NOME' +';'+ 'TIPO' +';'+ 'DT_CANC' +';'+ 'MOTIVO' +';'+ 'EMPRESA' +';'+ 'UNIDADE' +';'+ 'GRUPO' +';'+ 'PLANO' +';'+ 'LOGIN_CANCELAMENTO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P2  := qryDados.FIELDBYNAME('NOME').AsString;
          P3  := qryDados.FIELDBYNAME('TIPO').AsString;
          P4  := qryDados.FIELDBYNAME('DT_CANC').AsString;
          P5  := qryDados.FIELDBYNAME('MOTIVO').AsString;
          P6  := qryDados.FIELDBYNAME('EMPRESA').AsString;
          P7  := qryDados.FIELDBYNAME('UNIDADE').AsString;
          P8  := qryDados.FIELDBYNAME('GRUPO').AsString;
          P9  := qryDados.FIELDBYNAME('PLANO').AsString;
          P10 := qryDados.FIELDBYNAME('LOGIN_CANCELAMENTO').AsString;

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

   if frmPrincipal.lblCodRelatorio.Caption = '111' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_EMPRESA' +';'+ 'EMPRESA' +';'+ 'COD_UNIDADE' +';'+ 'UNIDADE' +';'+ 'VALOR' +';'+ 'VALOR_DESCONTO' +';'+ 'VALOR_PG' +';'+ 'VALOR_NF' +';'+ 'STATUS' +';'+ 'BOLETO' +';'+ 'NOTA_FISCAL' +';'+ 'VENCIMENTO' +';'+ 'VENC_ORIGINAL' +';'+ 'EMISSAO' +';'+ 'QTD_CLIENTES' +';'+ 'VALOR_ISS');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_EMPRESA').AsString;
          P2  := qryDados.FIELDBYNAME('EMPRESA').AsString;
          P3  := qryDados.FIELDBYNAME('COD_UNIDADE').AsString;
          P4  := qryDados.FIELDBYNAME('UNIDADE').AsString;
          P5  := qryDados.FIELDBYNAME('VALOR').AsString;
          P6  := qryDados.FIELDBYNAME('VALOR_DESCONTO').AsString;
          P7  := qryDados.FIELDBYNAME('VALOR_PG').AsString;
          P8  := qryDados.FIELDBYNAME('VALOR_NF').AsString;
          P9  := qryDados.FIELDBYNAME('STATUS').AsString;
          P10 := qryDados.FIELDBYNAME('COD_BOL').AsString;
          P11 := qryDados.FIELDBYNAME('NUM_NOTA').AsString;
          P12 := qryDados.FIELDBYNAME('DT_VENC').AsString;
          P13 := qryDados.FIELDBYNAME('DT_VENC_ORIGINAL').AsString;
          P14 := qryDados.FIELDBYNAME('DT_EMISSAO').AsString;
          P15 := qryDados.FIELDBYNAME('CLIENTES').AsString;
          P16 := qryDados.FIELDBYNAME('VALOR_ISS').AsString;

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

   if frmPrincipal.lblCodRelatorio.Caption = '56' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'SITUACAO' +';'+ 'MAT' +';'+ 'TIPO' +';'+ 'DT_ADESAO' +';'+ 'COD_SEG' +';'+ 'NOME' +';'+ 'COD_FORMA' +';'+ 'FORMA' +';'+ 'COD_VENDEDOR' +';'+ 'VENDEDOR' +';'+ 'COD_SUPERVISOR' +';'+ 'SUPERVISOR' +';'+ 'COD_DIRETOR' +';'+ 'DIRETOR' +';'+ 'GP_CARENCIA' +';'+ 'P_MENSALIDADE' +';'+ 'VALOR_01' +';'+ 'P_MENSALIDADE2' +';'+ 'VALOR_02' +';'+ 'P_MENSALIDADE3'+';'+ 'VALOR_03' +';'+ 'P_MENSALIDADE4'+';'+ 'VALOR_04' +';'+ 'P_MENSALIDADE5'+';'+ 'VALOR_05' +';'+ 'P_MENSALIDADE6'+';'+ 'VALOR_06' +';'+ 'TIPO_CONTRATO' +';'+ 'EMPRESA' +';'+ 'NOME_AD' +';'+ 'UNIDADE' +';'+ 'DT_INICIO' +';'+ 'RECEITA' +';'+ 'DESPESA' +';'+ 'DT_CANC' +';'+ 'COMISSAO' +';'+ 'PLANO' +';'+ 'CORRETOR_PLATAFORMA' +';'+ 'IDADE');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('SITUACAO').AsString;
          P2  := qryDados.FIELDBYNAME('MAT').AsString;
          P3  := qryDados.FIELDBYNAME('TIPO').AsString;
          P4  := qryDados.FIELDBYNAME('DT_ADESAO').AsString;
          P5  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P6  := qryDados.FIELDBYNAME('NOME').AsString;
          P7  := qryDados.FIELDBYNAME('COD_FORMA').AsString;
          P8  := qryDados.FIELDBYNAME('FORMA').AsString;
          P9  := qryDados.FIELDBYNAME('COD_VENDEDOR').AsString;
          P10 := qryDados.FIELDBYNAME('VENDEDOR').AsString;
          P11 := qryDados.FIELDBYNAME('COD_SUPERVISOR').AsString;
          P12 := qryDados.FIELDBYNAME('SUPERVISOR').AsString;
          P13 := qryDados.FIELDBYNAME('COD_DIRETOR').AsString;
          P14 := qryDados.FIELDBYNAME('DIRETOR').AsString;
          P15 := qryDados.FIELDBYNAME('GP_CARENCIA').AsString;
          P16 := qryDados.FIELDBYNAME('P_MENSALIDADE').AsString;
          P17 := qryDados.FIELDBYNAME('VALOR_01').AsString;
          P18 := qryDados.FIELDBYNAME('P_MENSALIDADE2').AsString;
          P19 := qryDados.FIELDBYNAME('VALOR_02').AsString;
          P20 := qryDados.FIELDBYNAME('P_MENSALIDADE3').AsString;
          P21 := qryDados.FIELDBYNAME('VALOR_03').AsString;
          P22 := qryDados.FIELDBYNAME('P_MENSALIDADE4').AsString;
          P23 := qryDados.FIELDBYNAME('VALOR_04').AsString;
          P24 := qryDados.FIELDBYNAME('P_MENSALIDADE5').AsString;
          P25 := qryDados.FIELDBYNAME('VALOR_05').AsString;
          P26 := qryDados.FIELDBYNAME('P_MENSALIDADE6').AsString;
          P27 := qryDados.FIELDBYNAME('VALOR_06').AsString;
          P28 := qryDados.FIELDBYNAME('TIPO_CONTRATO').AsString;
          P29 := qryDados.FIELDBYNAME('EMPRESA').AsString;
          P30 := qryDados.FIELDBYNAME('NOME_AD').AsString;
          P31 := qryDados.FIELDBYNAME('UNIDADE').AsString;
          P32 := qryDados.FIELDBYNAME('DT_INICIO').AsString;
          P33 := qryDados.FIELDBYNAME('RECEITA').AsString;
          P34 := qryDados.FIELDBYNAME('DESPESA').AsString;
          P35 := qryDados.FIELDBYNAME('DT_CANC').AsString;
          P36 := qryDados.FIELDBYNAME('COMISSAO').AsString;
          P37 := qryDados.FIELDBYNAME('PLANO').AsString;
          P38 := qryDados.FIELDBYNAME('CORRETOR_PLATAFORMA').AsString;
          P39 := qryDados.FIELDBYNAME('IDADE').AsString;

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
          cLinha := cLinha + P34 +';';
          cLinha := cLinha + P35 +';';
          cLinha := cLinha + P36 +';';
          cLinha := cLinha + P37 +';';
          cLinha := cLinha + P38 +';';
          cLinha := cLinha + P39 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

 if frmPrincipal.lblCodRelatorio.Caption = '129' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'SITUACAO' +';'+ 'PROPOSTA' +';'+ 'COD_SEG' +';'+ 'NOME' +';'+ 'CPF' +';'+ 'DT_INCL' +';'+ 'DT_INICIO' +';'+ 'CORRETOR');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('SITUACAO').AsString;
          P2  := qryDados.FIELDBYNAME('PROPOSTA').AsString;
          P3  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P4  := qryDados.FIELDBYNAME('NOME').AsString;
          P5  := qryDados.FIELDBYNAME('CPF').AsString;
          P6  := qryDados.FIELDBYNAME('DT_INCL').AsString;
          P7  := qryDados.FIELDBYNAME('DT_INICIO').AsString;
          P8  := qryDados.FIELDBYNAME('CORRETOR').AsString;

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

  if frmPrincipal.lblCodRelatorio.Caption = '138' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_EMPRESA' +';'+ 'EMPRESA' +';'+ 'COD_UNIDADE' +';'+ 'UNIDADE' +';'+ 'CNPJ_UNIDADE' +';'+ 'DT_ADESAO' +';'+ 'TIPO_SEG' +';'+ 'COD_SEG' +';'+ 'NOME_SEG' +';'+ 'CPF_SEG' +';'+ 'COD_BOL' +';'+ 'DT_VENC' +';'+ 'DIA_VENC' +';'+ 'COMP' +';'+ 'VALOR_BOL_SEG' +';'+ 'MES_COMISSAO_BOL' +';'+ 'VALOR_COMISSAO' +';'+ 'DT_PAGAMENTO' +';'+ 'DT_CONCIL_REAL' +';'+ 'OBS');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_EMPRESA').AsString;
          P2  := qryDados.FIELDBYNAME('EMPRESA').AsString;
          P3  := qryDados.FIELDBYNAME('COD_UNIDADE').AsString;
          P4  := qryDados.FIELDBYNAME('UNIDADE').AsString;
          P5  := qryDados.FIELDBYNAME('CNPJ_UNIDADE').AsString;
          P6  := qryDados.FIELDBYNAME('DT_ADESAO').AsString;
          P7  := qryDados.FIELDBYNAME('TIPO_SEG').AsString;
          P8  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P9  := qryDados.FIELDBYNAME('NOME_SEG').AsString;
          P10 := qryDados.FIELDBYNAME('CPF_SEG').AsString;
          P11 := qryDados.FIELDBYNAME('COD_BOL').AsString;
          P12 := qryDados.FIELDBYNAME('DT_VENC').AsString;
          P13 := qryDados.FIELDBYNAME('DIA_VENC').AsString;
          P14 := qryDados.FIELDBYNAME('COMP').AsString;
          P15 := qryDados.FIELDBYNAME('VALOR_BOL_SEG').AsString;
          P16 := qryDados.FIELDBYNAME('MES_COMISSAO_BOL').AsString;
          P17 := qryDados.FIELDBYNAME('VALOR_COMISSAO').AsString;
          P18 := qryDados.FIELDBYNAME('DT_PAGAMENTO').AsString;
          P20 := qryDados.FIELDBYNAME('OBS').AsString;

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

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '132' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_BOL' +';'+ 'DT_VENC' +';'+ 'DIA_VENC' +';'+ 'DT_CONCIL' +';'+ 'COD_SEG' +';'+ 'BENEFICIARIO' +';'+ 'VLR_MENS' +';'+ 'NMES' +';'+ 'VLR_COM' +';'+ 'SITUACAO' +';'+ 'COD_UNIDADE' +';'+ 'UNIDADE' +';'+ 'CNPJ_UNIDADE' +';'+ 'COD_VEND' +';'+ 'VENDEDOR' +';'+ 'CGC_VENDEDOR');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_BOL').AsString;
          P2  := qryDados.FIELDBYNAME('DT_VENC').AsString;
          P3  := qryDados.FIELDBYNAME('DIA_VENC').AsString;
          P4  := qryDados.FIELDBYNAME('DT_CONCIL').AsString;
          P5  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P6  := qryDados.FIELDBYNAME('NOME').AsString;
          P7  := qryDados.FIELDBYNAME('VLR_MENS').AsString;
          P8  := qryDados.FIELDBYNAME('NMES').AsString;
          P9  := qryDados.FIELDBYNAME('VLR_COM').AsString;
          P10 := qryDados.FIELDBYNAME('SITUACAO').AsString;
          P11 := qryDados.FIELDBYNAME('COD_UNIDADE').AsString;
          P12 := qryDados.FIELDBYNAME('UNIDADE').AsString;
          P13 := qryDados.FIELDBYNAME('CNPJ_UNIDADE').AsString;
          P14 := qryDados.FIELDBYNAME('COD_VEND').AsString;
          P15 := qryDados.FIELDBYNAME('VENDEDOR').AsString;
          P16 := qryDados.FIELDBYNAME('CGC_VENDEDOR').AsString;

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

  if frmPrincipal.lblCodRelatorio.Caption = '134' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_SEG' +';'+ 'NOME' +';'+ 'CPF' +';'+ 'DT_NASC' +';'+ 'SEXO' +';'+ 'COD_PLANO' +';'+ 'PLANO' +';'+ 'COD_EMPRESA' +';'+ 'EMPRESA' +';'+ 'COD_UNIDADE' +';'+ 'UNIDADE' +';'+ 'VALOR' +';'+ 'GLOSA' +';'+ 'VALOR_PAGO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P2  := qryDados.FIELDBYNAME('NOME').AsString;
          P3  := qryDados.FIELDBYNAME('CPF').AsString;
          P4  := qryDados.FIELDBYNAME('DT_NASC').AsString;
          P5  := qryDados.FIELDBYNAME('SEXO').AsString;
          P6  := qryDados.FIELDBYNAME('COD_PLANO').AsString;
          P7  := qryDados.FIELDBYNAME('PLANO').AsString;
          P8  := qryDados.FIELDBYNAME('COD_EMPRESA').AsString;
          P9  := qryDados.FIELDBYNAME('EMPRESA').AsString;
          P10 := qryDados.FIELDBYNAME('COD_UNIDADE').AsString;
          P11 := qryDados.FIELDBYNAME('UNIDADE').AsString;
          P12 := qryDados.FIELDBYNAME('VALOR').AsString;
          P13 := qryDados.FIELDBYNAME('GLOSA').AsString;
          P14 := qryDados.FIELDBYNAME('VALOR_PAGO').AsString;

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

  if frmPrincipal.lblCodRelatorio.Caption = '141' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'ANO_MES' +';'+ 'COD_CAR' +';'+ 'CARENCIA' +';'+ 'QUANTIDADE' +';'+ 'SINISTRALIDADE');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('ANO_MES').AsString;
          P2  := qryDados.FIELDBYNAME('COD_CAR').AsString;
          P3  := qryDados.FIELDBYNAME('CARENCIA').AsString;
          P4  := qryDados.FIELDBYNAME('QUANTIDADE').AsString;
          P5  := qryDados.FIELDBYNAME('SINISTRALIDADE').AsString;

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

  if frmPrincipal.lblCodRelatorio.Caption = '145' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_EMPRESA' +';'+ 'EMPRESA' +';'+ 'DT_VENC_EMP' +';'+ 'COD_CLIENTE' +';'+ 'CLIENTE' +';'+ 'ADESAO' +';'+ 'PLANO' +';'+ 'SITUACAO' +';'+ 'VENDEDOR' +';'+ 'DT_VENC_CLIENTE');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_EMPRESA').AsString;
          P2  := qryDados.FIELDBYNAME('EMPRESA').AsString;
          P3  := qryDados.FIELDBYNAME('DT_VENC_EMP').AsString;
          P4  := qryDados.FIELDBYNAME('COD_CLIENTE').AsString;
          P5  := qryDados.FIELDBYNAME('CLIENTE').AsString;
          P6  := qryDados.FIELDBYNAME('ADESAO').AsString;
          P7  := qryDados.FIELDBYNAME('PLANO').AsString;
          P8  := qryDados.FIELDBYNAME('SITUACAO').AsString;
          P9  := qryDados.FIELDBYNAME('VENDEDOR').AsString;
          P10 := qryDados.FIELDBYNAME('DT_VENC_CLIENTE').AsString;

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

 end;
end;

procedure TfrmRelatorio13.chk_EmpresaClick(Sender: TObject);
begin
  if chk_Empresa.Checked = True then
  begin
  edtCodigoEmp.Color := clWindow;
  edtCodigoEmp.Enabled := True;
  edtNomeEmp.Enabled := True;
  chkUnidade.Enabled := True;
  end else
  begin
  edtCodigoEmp.Text := '';
  edtNomeEmp.Text := '';
  lblCodigoEmpresa.Caption := '0';
  edtCodigoEmp.Color := clMenu;
  edtCodigoEmp.Enabled := False;
  edtNomeEmp.Enabled := False;
  chkUnidade.Enabled := False;
  chkUnidade.Checked := False;
  edtCodigoUnidade.Text := '';
  edtNomeUnidade.Text := '';
  lblCodigoUnidade.Caption := '0';
  edtCodigoUnidade.Color := clMenu;
  edtCodigoUnidade.Enabled := False;
  edtNomeUnidade.Enabled := False;
  end;
end;

procedure TfrmRelatorio13.btnImprimirClick(Sender: TObject);
begin
  with qryExec do // INSERIR LOG DE IMPRESSAO
  begin
  Close;
  SQL.Clear;
  SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, PARAMETRO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'I'+#39+', '+#39+parametro+#39+', SYSDATE)');
  ExecSQL;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '65' then
  begin
    qrpRelat65.qrlblTitulo.Caption := 'Receita x Sinistralidade - ' + '(' + lblCodigoEmpresa.Caption + ') ' + edtNomeEmp.Text;
//    qrpRelat65.qrlblMediaVida.Caption := 'R$ ' + FormatFloat('#,0.00',media);
    qrpRelat65.qrlblTotalReceita.Caption := 'R$ ' + FormatFloat('#,0.00',receita);
    qrpRelat65.qrlblTotalDespesa.Caption := 'R$ ' + FormatFloat('#,0.00',despesa);
    qrpRelat65.Prepare;
    qrpRelat65.qrlblPagina.Caption := IntToStr(qrpRelat65.QRPrinter.PageCount);
    qrpRelat65.PreviewModal;
  end;

end;

procedure TfrmRelatorio13.qryDadosAfterOpen(DataSet: TDataSet);
begin
frmPrincipal.DimensionarGrid(dbgDados,Self);

  if frmPrincipal.lblCodRelatorio.Caption = '111' then
  begin
    TCurrencyField(qryDados.FieldByName('COD_EMPRESA')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_UNIDADE')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('VALOR')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('VALOR_PG')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('VALOR_NF')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('COD_BOL')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('NUM_NOTA')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DT_VENC')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DT_VENC_ORIGINAL')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DT_EMISSAO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('CLIENTES')).DisplayFormat := '#,##0';
    TCurrencyField(qryDados.FieldByName('VALOR_ISS')).DisplayFormat := 'R$ #,##0.00';        
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '145' then
  begin
    TCurrencyField(qryDados.FieldByName('COD_EMPRESA')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DT_VENC_EMP')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_CLIENTE')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('ADESAO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DT_VENC_CLIENTE')).Alignment := taCenter;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '56' then
  begin
    TCurrencyField(qryDados.FieldByName('SITUACAO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('MAT')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('TIPO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DT_ADESAO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_SEG')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_FORMA')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_VENDEDOR')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_SUPERVISOR')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_DIRETOR')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('GP_CARENCIA')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('P_MENSALIDADE')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('VALOR_01')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('VALOR_01')).Alignment := taRightJustify;
    TCurrencyField(qryDados.FieldByName('P_MENSALIDADE2')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('VALOR_02')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('VALOR_02')).Alignment := taRightJustify;
    TCurrencyField(qryDados.FieldByName('P_MENSALIDADE3')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('VALOR_03')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('VALOR_03')).Alignment := taRightJustify;
    TCurrencyField(qryDados.FieldByName('DT_INICIO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('RECEITA')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('RECEITA')).Alignment := taRightJustify;
    TCurrencyField(qryDados.FieldByName('DESPESA')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('DESPESA')).Alignment := taRightJustify;
    TCurrencyField(qryDados.FieldByName('COMISSAO')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('COMISSAO')).Alignment := taRightJustify;
    TCurrencyField(qryDados.FieldByName('DT_CANC')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('IDADE')).Alignment := taCenter;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '131' then
  begin
    TCurrencyField(qryDados.FieldByName('COD_EMPRESA')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_UNIDADE')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('CNPJ_UNIDADE')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DT_ADESAO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_SEG')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('CPF_SEG')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_BOL')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DT_VENC')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('VALOR_BOL_SEG')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('MES_COMISSAO_BOL')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('VALOR_COMISSAO')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('DT_PAGAMENTO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DT_CONCIL_REAL')).Alignment := taCenter;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '141' then
  begin
    TCurrencyField(qryDados.FieldByName('ANO_MES')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('QUANTIDADE')).DisplayFormat := '#,##0';
    TCurrencyField(qryDados.FieldByName('SINISTRALIDADE')).DisplayFormat := 'R$ #,##0.00';
  end;
end;

procedure TfrmRelatorio13.chkUnidadeClick(Sender: TObject);
begin
  if chkUnidade.Checked = True then
  begin
  edtCodigoUnidade.Color := clWindow;
  edtCodigoUnidade.Enabled := True;
  edtNomeUnidade.Enabled := True;
  end else
  begin
  edtCodigoUnidade.Text := '';
  edtNomeUnidade.Text := '';
  lblCodigoUnidade.Caption := '0';
  edtCodigoUnidade.Color := clMenu;
  edtCodigoUnidade.Enabled := False;
  edtNomeUnidade.Enabled := False;
  end;
end;

procedure TfrmRelatorio13.edtCodigoUnidadeKeyPress(Sender: TObject;
  var Key: Char);
begin
    if not (key in ['0'..'9',Chr(8), Chr(3), Chr(22), Chr(24), Chr(44)]) then
      key := #0;
end;

procedure TfrmRelatorio13.edtCodigoUnidadeExit(Sender: TObject);
begin
  if chkUnidade.Checked = True then
  begin
    if edtCodigoUnidade.Text <> '' then
    begin
      with qryExec do
      begin
      Close;
      SQL.Clear;

        if (frmPrincipal.lblCodRelatorio.Caption = '129') then
          SQL.Add('SELECT COD_SETSET, NOME FROM PLATINUM.IM_SETSE WHERE COD_SET ='+lblCodigoEmpresa.Caption+' AND COD_SETSET = '+ edtCodigoUnidade.Text)
        else
          SQL.Add('SELECT COD_SETSET, NOME FROM IM_SETSE WHERE COD_SET ='+lblCodigoEmpresa.Caption+' AND COD_SETSET = '+ edtCodigoUnidade.Text);
      Open;
      end;

      if qryExec.RecordCount > 0 then
      begin
      edtNomeUnidade.Text := qryExec.FieldByName('NOME').AsString;
      lblCodigoUnidade.Caption := qryExec.FieldByName('COD_SETSET').AsString;
      end else
      begin
      ShowMessage('Código da Unidade Inválido');
      end;
    end;  
  end;
end;

procedure TfrmRelatorio13.Med_MesKeyPress(Sender: TObject; var Key: Char);
begin
    if not (key in ['0'..'9',Chr(8), Chr(3), Chr(22), Chr(24), Chr(44)]) then
      key := #0;
end;

procedure TfrmRelatorio13.Med_MesExit(Sender: TObject);
var
d_Data : TDate;
begin
  try
    d_Data := StrToDate('01/'+Med_Mes.Text);
    Mes := FormatDateTime('yyyy/mm',d_Data);
  except
    ShowMessage('Data base inválida');
    Med_Mes.Text := '  /    ';
    Med_Mes.SetFocus;
  end;
end;

procedure TfrmRelatorio13.chkPlanoClick(Sender: TObject);
begin
  if chkPlano.Checked = True then
    edtPlano.Enabled := True
  else
    edtPlano.Enabled := False;
end;

procedure TfrmRelatorio13.cbxPlanoChange(Sender: TObject);
begin
  if chkPlano.Enabled = True then
  begin
    with qryExec do
    begin
      Close;
      SQL.Clear;
      SQL.Add('select cod_plano from im_plano where nome = '+#39+cbxPlano.Text+#39+ 'and rownum = 1');
      Open;
    end;

    lblCodigoPlano.Caption := qryExec.FieldByName('cod_plano').AsString;
  end;
end;

procedure TfrmRelatorio13.edtPlanoKeyPress(Sender: TObject; var Key: Char);
begin
    if not (key in ['0'..'9',Chr(8), Chr(3), Chr(22), Chr(24), Chr(44), ',']) then
      key := #0;
end;

end.
