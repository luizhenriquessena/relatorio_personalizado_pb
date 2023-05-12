unit Relatorio1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, Grids, DBGrids, Mask, DBCtrls, DB, ADODB;

type
  TfrmRelatorio1 = class(TForm)
    dbgDados: TDBGrid;
    btnImprimir: TSpeedButton;
    btnExportar: TSpeedButton;
    GroupBox1: TGroupBox;
    Label1: TLabel;
    lblTotal: TLabel;
    lblRTotal: TLabel;
    Med_DataBase: TMaskEdit;
    lblDataBase: TLabel;
    chkEmpresa: TCheckBox;
    edtCodigoEmp: TEdit;
    edtNomeEmp: TEdit;
    btnGerar: TSpeedButton;
    edtCodigoUnidade: TEdit;
    chkUnidade: TCheckBox;
    edtNomeUnidade: TEdit;
    chkPlano: TCheckBox;
    qryExec: TADOQuery;
    lblCodigoEmpresa: TLabel;
    lblCodigoUnidade: TLabel;
    qryPlano: TADOQuery;
    qryPlanoCOD_PLANO: TBCDField;
    qryPlanoNOME: TWideStringField;
    qryPlanoPLANO: TWideStringField;
    lblPlano: TLabel;
    qryDados: TADOQuery;
    dsDados: TDataSource;
    SaveDialog1: TSaveDialog;
    cbxPlano: TComboBox;
    procedure Med_DataBaseExit(Sender: TObject);
    procedure btnGerarClick(Sender: TObject);
    procedure chkEmpresaClick(Sender: TObject);
    procedure chkUnidadeClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure chkPlanoClick(Sender: TObject);
    procedure edtCodigoEmpExit(Sender: TObject);
    procedure edtCodigoUnidadeExit(Sender: TObject);
    procedure Med_DataBaseDblClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnExportarClick(Sender: TObject);
    procedure cbxPlanoChange(Sender: TObject);
    procedure edtCodigoEmpKeyPress(Sender: TObject; var Key: Char);
    procedure edtCodigoUnidadeKeyPress(Sender: TObject; var Key: Char);
    procedure qryDadosAfterOpen(DataSet: TDataSet);
  private
    { Private declarations }
    function ListaPlanos : TStrings;
  public
    { Public declarations }
  end;

var
  frmRelatorio1: TfrmRelatorio1;
  parametro: String;
implementation

uses Principal, DateUtils, DataModule, Math, StrUtils; //DB, ADODB, Math, ;
{$R *.dfm}

procedure TfrmRelatorio1.Med_DataBaseExit(Sender: TObject);
var d_Data : TDateTime;
begin
  try
    d_Data := StrToDateTime(Med_DataBase.Text);
    Med_DataBase.Text := FormatDateTime('dd/mm/yyyy',d_Data);
  except
    ShowMessage('Data base inválida');
    Med_DataBase.Text := '  /  /    ';
    Med_DataBase.SetFocus;
  end;
end;

procedure TfrmRelatorio1.btnGerarClick(Sender: TObject);
var
Total: Integer;
begin
  parametro := '';
  if Med_DataBase.Text <> '  /  /    ' then
  begin
    parametro := 'Data Base '+Med_DataBase.Text;
    with qryDados do
    begin
    Close;
    SQL.Clear;
    SQL.Add(frmPrincipal.memSQL.Text);
    SQL.Text := StringReplace(SQL.Text,'&DATA',Med_DataBase.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco

    if chkEmpresa.Checked = True then
    begin
    SQL.Add('   AND S.COD_SET = '+lblCodigoEmpresa.Caption);
    parametro := parametro + ' - Empresa '+lblCodigoEmpresa.Caption;
    end;

    if chkUnidade.Checked = True then
    begin
    SQL.Add('   AND S.COD_SETSET = '+lblCodigoUnidade.Caption);
    parametro := parametro + ' - Unidade '+  lblCodigoUnidade.Caption;
    end;

    if chkPlano.Checked = True then
    begin
    SQL.Add('   AND S.COD_PLANO = '+lblPlano.Caption);
    parametro := parametro + ' - Plano '+ lblPlano.Caption;
    end;

    case AnsiIndexStr(frmPrincipal.lblCodRelatorio.Caption,['1','2','6']) of  // Inserir Group By de acordo com o relatório selecionado.
    0 : SQL.Add(') GROUP BY DIA, FAIXA, SEXO'); // GROUP BY DO RELATORIO 1
    1 : SQL.Add('GROUP BY BAIRRO, CIDADE, ESTADO'); // GROUP BY DI RELATORIO 2
    2 : SQL.Add('GROUP BY R.REGIAO'); // GROUP BY DO RELATORIO 6
    end;

    SQL.Add('ORDER BY QUANTIDADE DESC');  
    Open;

    First;
    while not Eof do
    begin
      Total := Total + fieldbyname('QUANTIDADE').AsVariant;
      Next;
    end;
    lblRTotal.Caption :=  FormatFloat('#,##0',Total);
    end;

    with qryExec do // INSERIR LOG DE VISUALIZAÇÃO
    begin
    Close;
    SQL.Clear;
    SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, PARAMETRO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'V'+#39+', '+#39+parametro+#39+', SYSDATE)');
    ExecSQL;
    end;

  end else
  begin
  ShowMessage('Favor informar a data base.');
  Abort;
  end;
end;

procedure TfrmRelatorio1.chkEmpresaClick(Sender: TObject);
begin
  if chkEmpresa.Checked = True then
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
  edtCodigoEmp.Color := clSilver;
  edtCodigoEmp.Enabled := False;
  edtNomeEmp.Enabled := False;
  chkUnidade.Enabled := False;
  chkUnidade.Checked := False;
  edtCodigoUnidade.Text := '';
  edtNomeUnidade.Text := '';
  lblCodigoUnidade.Caption := '0';
  edtCodigoUnidade.Color := clSilver;
  edtCodigoUnidade.Enabled := False;
  edtNomeUnidade.Enabled := False;
  end;
end;

procedure TfrmRelatorio1.chkUnidadeClick(Sender: TObject);
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
  edtCodigoUnidade.Color := clSilver;
  edtCodigoUnidade.Enabled := False;
  edtNomeUnidade.Enabled := False;
  end;
end;

procedure TfrmRelatorio1.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
qryExec.Close;
qryPlano.Close;
qryDados.Close;
end;

procedure TfrmRelatorio1.chkPlanoClick(Sender: TObject);
begin
  if chkPlano.Checked = True then
  begin
  qryPlano.Open;
  cbxPlano.Enabled := True;
  cbxPlano.Color   := clWindow;
  cbxPlano.Items   := ListaPlanos;
  cbxPlano.ItemIndex := 0;
  lblPlano.Caption := '1';
  end else
  begin
  qryPlano.Close;
  cbxPlano.Items.Clear;
  cbxPlano.Enabled := False;
  cbxPlano.Color := clSilver;
  lblPlano.Caption := '0';
  end;
end;

procedure TfrmRelatorio1.edtCodigoEmpExit(Sender: TObject);
begin
  if chkEmpresa.Checked = True then
  begin
    if edtCodigoEmp.Text <> '' then
    begin
      with qryExec do
      begin
      Close;
      SQL.Clear;
      SQL.Add('SELECT COD_SET, DESCRICAO FROM IM_SETOR WHERE COD_SET = '+ edtCodigoEmp.Text);
      Open;
      end;

      if qryExec.RecordCount > 0 then
      begin
      edtNomeEmp.Text := qryExec.FieldByName('DESCRICAO').AsString;
      lblCodigoEmpresa.Caption := qryExec.FieldByName('COD_SET').AsString;
      end else
      begin
      ShowMessage('Código da Empresa Inválido');
      end;
    end;
  end;
end;

procedure TfrmRelatorio1.edtCodigoUnidadeExit(Sender: TObject);
begin
  if chkUnidade.Checked = True then
  begin
    if edtCodigoUnidade.Text <> '' then
    begin
      with qryExec do
      begin
      Close;
      SQL.Clear;
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

procedure TfrmRelatorio1.Med_DataBaseDblClick(Sender: TObject);
begin
if (Med_DataBase.Text = '  /  /    ')  then
Med_DataBase.Text :=  DateToStr(Date);
end;

procedure TfrmRelatorio1.FormShow(Sender: TObject);
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
end;

procedure TfrmRelatorio1.btnExportarClick(Sender: TObject);
var
F: TextFile;
P1, P2, P3, P4, P5, P6, P7, cLinha : String; //Variaveis comuns
begin
 if qryDados.RecordCount > 0 then
 begin

  if frmPrincipal.lblCodRelatorio.Caption = '1' then // Exportar relatório (1)
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'DIA' +';'+ 'FAIXA' +';'+ 'SEXO' +';'+ 'QUANTIDADE');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('DIA').AsString;
          P2  := qryDados.FIELDBYNAME('FAIXA').AsString;
          P3  := qryDados.FIELDBYNAME('SEXO').AsString;
          P4  := qryDados.FIELDBYNAME('QUANTIDADE').AsString;

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

  if frmPrincipal.lblCodRelatorio.Caption = '2' then // Exportar relatório (2)
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'CIDADE' +';'+ 'ESTADO' +';'+  'BAIRRO' +';'+ 'QUANTIDADE');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('BAIRRO').AsString;
          P2  := qryDados.FIELDBYNAME('CIDADE').AsString;
          P3  := qryDados.FIELDBYNAME('ESTADO').AsString;
          P4  := qryDados.FIELDBYNAME('QUANTIDADE').AsString;

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

  if frmPrincipal.lblCodRelatorio.Caption = '6' then // Exportar relatório (6)
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'REGIAO' +';'+ 'QUANTIDADE');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('REGIAO').AsString;
          P2  := qryDados.FIELDBYNAME('QUANTIDADE').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';

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

function TfrmRelatorio1.ListaPlanos: TStrings;
begin
  Result := TStringList.Create;
  qryPlano.First;
  while not (qryPlano.Eof) do
  begin
    Result.Add(qryPlano.FieldByName('PLANO').AsString);
    qryPlano.Next;
  end;
end;

procedure TfrmRelatorio1.cbxPlanoChange(Sender: TObject);
var codigo : String;
begin
  if chkPlano.Checked = True then
  lblPlano.Caption := cbxPlano.Text[1];
end;

procedure TfrmRelatorio1.edtCodigoEmpKeyPress(Sender: TObject;
  var Key: Char);
begin
    if not (key in ['0'..'9',Chr(8), Chr(3), Chr(22), Chr(24), Chr(44)]) then
      key := #0;
end;

procedure TfrmRelatorio1.edtCodigoUnidadeKeyPress(Sender: TObject;
  var Key: Char);
begin
    if not (key in ['0'..'9',Chr(8), Chr(3), Chr(22), Chr(24), Chr(44)]) then
      key := #0;
end;

procedure TfrmRelatorio1.qryDadosAfterOpen(DataSet: TDataSet);
begin
  frmPrincipal.DimensionarGrid(dbgDados,Self);

  if frmPrincipal.lblCodRelatorio.Caption = '1' then
  begin
    TCurrencyField(qryDados.FieldByName('DIA')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('FAIXA')).Alignment := taCenter;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '2' then
  begin
    TCurrencyField(qryDados.FieldByName('ESTADO')).Alignment := taCenter;
  end;
end;

end.
