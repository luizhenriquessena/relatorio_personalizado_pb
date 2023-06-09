unit Relatorio2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, Buttons, Grids, DBGrids, StdCtrls, Mask;

type
  TfrmRelatorio2 = class(TForm)
    GroupBox1: TGroupBox;
    lblDataBase: TLabel;
    Med_DataBase: TMaskEdit;
    dbgDados: TDBGrid;
    Label1: TLabel;
    btnExportar: TSpeedButton;
    btnImprimir: TSpeedButton;
    lblTotal: TLabel;
    lblRTotal: TLabel;
    dsDados: TDataSource;
    qryDados: TADOQuery;
    SaveDialog1: TSaveDialog;
    qryExec: TADOQuery;
    btnGerar: TSpeedButton;
    procedure Med_DataBaseDblClick(Sender: TObject);
    procedure Med_DataBaseExit(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure btnGerarClick(Sender: TObject);
    procedure btnExportarClick(Sender: TObject);
    procedure qryDadosAfterOpen(DataSet: TDataSet);
    procedure btnImprimirClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmRelatorio2: TfrmRelatorio2;
  parametro: String;
implementation

{$R *.dfm}
uses Principal, DateUtils, DataModule, QuickRpt, Rel136e137;

procedure TfrmRelatorio2.Med_DataBaseDblClick(Sender: TObject);
begin
if (Med_DataBase.Text = '  /  /    ')  then
Med_DataBase.Text :=  DateToStr(Date);
end;

procedure TfrmRelatorio2.Med_DataBaseExit(Sender: TObject);
var d_Data : TDateTime;
begin
  try
    d_Data := StrToDateTime(Med_DataBase.Text);
    Med_DataBase.Text := FormatDateTime('dd/mm/yyyy',d_Data);
  except
    ShowMessage('Data base inv�lida');
    Med_DataBase.Text := '  /  /    ';
    Med_DataBase.SetFocus;
  end;
end;

procedure TfrmRelatorio2.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
qryExec.Close;
qryDados.Close;
end;

procedure TfrmRelatorio2.FormShow(Sender: TObject);
begin
  if frmPrincipal.lblImprime.Caption = 'T' then //Habilita bot�o Imprimir
  begin
  btnImprimir.Enabled := True;
  btnImprimir.Visible := True;
  end else
  begin
  btnImprimir.Enabled := False;
  btnImprimir.Visible := False;
  end;
  
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

  with qryExec do
  begin
  Close;
  SQL.Clear;
  SQL.Add('ALTER SESSION SET NLS_NUMERIC_CHARACTERS = '+#39+',.'+#39);
  ExecSQL;
  end;
end;

procedure TfrmRelatorio2.btnGerarClick(Sender: TObject);
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
    Open;
    end;
    frmPrincipal.DimensionarGrid(dbgDados,Self);

    if (frmPrincipal.lblCodRelatorio.Caption = '3') or (frmPrincipal.lblCodRelatorio.Caption = '4') then
    begin
      with qryExec do // INSERIR TOTAL
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL2.Text);
      SQL.Text := StringReplace(SQL.Text,'&DATA',Med_DataBase.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL2 puxado do banco
      Open;
      lblRTotal.Caption := qryExec.FieldbyName('TOTAL').AsString;
      end;
    end else if frmPrincipal.lblCodRelatorio.Caption = '136' then
    begin
      lblRTotal.Caption := qryDados.FieldByName('VALOR_TOTAL').AsString;
    end else
    begin
      lblRTotal.Caption := IntToStr(qryDados.RecordCount);
    end;

    with qryExec do // INSERIR LOG DE VISUALIZA��O
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

procedure TfrmRelatorio2.btnExportarClick(Sender: TObject);
var
F: TextFile;
P1, P2, P3, P4, P5, P6, P7, P8, P9, P10, P11, P12, P13, cLinha : String;
begin
 if qryDados.RecordCount > 0 then
 begin
  if (frmPrincipal.lblCodRelatorio.Caption = '3') or (frmPrincipal.lblCodRelatorio.Caption = '4') then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_CONTA' +';'+ 'CTA_CONTABIL' +';'+ 'SALDO' +';'+ 'DESCRICAO' +';'+ 'DATA');
      qryDados.First;
        while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_CONTA').AsString;
          P2  := qryDados.FIELDBYNAME('CTA_CONTABIL').AsString;
          P3  := qryDados.FIELDBYNAME('SALDO').AsString;
          P4  := qryDados.FIELDBYNAME('DESCRICAO').AsString;
          P5  := qryDados.FIELDBYNAME('DATA').AsString;

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

  if (frmPrincipal.lblCodRelatorio.Caption = '12') or (frmPrincipal.lblCodRelatorio.Caption = '142') then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_SEG' +';'+ 'NOME' +';'+ 'COD_SET' +';'+ 'EMPRESA' +';'+ 'NUM_FACPLAN' +';'+ 'NUM_BOL' +';'+ 'MOV');
      qryDados.First;
        while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P2  := qryDados.FIELDBYNAME('NOME').AsString;
          P3  := qryDados.FIELDBYNAME('COD_SET').AsString;
          P4  := qryDados.FIELDBYNAME('EMPRESA').AsString;
          P5  := qryDados.FIELDBYNAME('NUM_FACPLAN').AsString;
          P6  := qryDados.FIELDBYNAME('NUM_BOL').AsString;
          P7  := qryDados.FIELDBYNAME('MOV').AsString;

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

  if frmPrincipal.lblCodRelatorio.Caption = '130' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_PLANO' +';'+ 'PLANO' +';'+ 'TIPO' +';'+ 'VIDAS');
      qryDados.First;
        while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_PLANO').AsString;
          P2  := qryDados.FIELDBYNAME('PLANO').AsString;
          P3  := qryDados.FIELDBYNAME('TIPO').AsString;
          P4  := qryDados.FIELDBYNAME('VIDAS').AsString;

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

  if (frmPrincipal.lblCodRelatorio.Caption = '136') or (frmPrincipal.lblCodRelatorio.Caption = '137') then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'DATA_BASE' +';'+ 'VALOR_UNITARIO' +';'+ 'VIDAS' +';'+ 'VALOR_TOTAL');
      qryDados.First;
        while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('DATA_BASE').AsString;
          P2  := qryDados.FIELDBYNAME('VALOR_UNITARIO').AsString;
          P3  := qryDados.FIELDBYNAME('VIDAS').AsString;
          P4  := qryDados.FIELDBYNAME('VALOR_TOTAL').AsString;

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

  with qryExec do // INSERIR LOG DE EXPORTA��O
  begin
  Close;
  SQL.Clear;
  SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, PARAMETRO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'E'+#39+', '+#39+parametro+#39+', SYSDATE)');
  ExecSQL;
  end;
 end;
end;

procedure TfrmRelatorio2.qryDadosAfterOpen(DataSet: TDataSet);
begin
frmPrincipal.DimensionarGrid(dbgDados,Self);

  if frmPrincipal.lblCodRelatorio.Caption = '130' then
  begin
    TCurrencyField(qryDados.FieldByName('COD_PLANO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('VIDAS')).DisplayFormat := '#,##0';
  end;

  if (frmPrincipal.lblCodRelatorio.Caption = '136') or (frmPrincipal.lblCodRelatorio.Caption = '137') then
  begin
    TCurrencyField(qryDados.FieldByName('DATA_BASE')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('VALOR_UNITARIO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('VIDAS')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('VIDAS')).DisplayFormat := '#,##0';
    TCurrencyField(qryDados.FieldByName('VALOR_TOTAL')).DisplayFormat := 'R$ #,##0.00';
  end;
end;

procedure TfrmRelatorio2.btnImprimirClick(Sender: TObject);
begin
  if qryDados.RecordCount > 0 then
  begin

    if frmPrincipal.lblCodRelatorio.Caption = '136' then
    begin
      qrpRelat136e137.qrmemo.Lines.Clear;
      qrpRelat136e137.qrmemo.Lines.Add('Dados extra�dos da base para c�lculo de valor.');
      qrpRelat136e137.qrlblTitulo.Caption := 'Vidas Aptas Med Life';
      qrpRelat136e137.qrlDataBase.Caption := qryDados.FieldbyName('DATA_BASE').AsString;
      qrpRelat136e137.qrlValorUnitario.Caption := qryDados.FieldbyName('VALOR_UNITARIO').AsString;
      qrpRelat136e137.qrlVidas.Caption         := FormatFloat('#,##0',qryDados.FieldbyName('VIDAS').AsFloat); //.DisplayFormat := '#,##0';
      qrpRelat136e137.qrlValorTotal.Caption    := 'R$ '+ FormatFloat('#,##0.00',qryDados.FieldbyName('VALOR_TOTAL').AsFloat); //.DisplayFormat := 'R$ #,##0.00';
      qrpRelat136e137.Prepare;
      qrpRelat136e137.PreviewModal;
    end;

    if frmPrincipal.lblCodRelatorio.Caption = '137' then
    begin
      qrpRelat136e137.qrmemo.Lines.Clear;
      qrpRelat136e137.qrmemo.Lines.Add('Dados extra�dos da base para c�lculo de valor.');
      qrpRelat136e137.qrlblTitulo.Caption := 'Vidas Aptas Impacto';
      qrpRelat136e137.qrlDataBase.Caption := qryDados.FieldbyName('DATA_BASE').AsString;
      qrpRelat136e137.qrlValorUnitario.Caption := qryDados.FieldbyName('VALOR_UNITARIO').AsString;
      qrpRelat136e137.qrlVidas.Caption         := FormatFloat('#,##0',qryDados.FieldbyName('VIDAS').AsFloat); //.DisplayFormat := '#,##0';
      qrpRelat136e137.qrlValorTotal.Caption    := 'R$ '+ FormatFloat('#,##0.00',qryDados.FieldbyName('VALOR_TOTAL').AsFloat); //.DisplayFormat := 'R$ #,##0.00';
      qrpRelat136e137.Prepare;
      qrpRelat136e137.PreviewModal;
    end;
  end;
end;

end.
