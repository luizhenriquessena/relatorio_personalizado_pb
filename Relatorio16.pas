unit Relatorio16;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, Grids, DBGrids, StdCtrls, Mask, Buttons;

type
  TfrmRelatorio16 = class(TForm)
    Label1: TLabel;
    btnExportar: TSpeedButton;
    btnImprimir: TSpeedButton;
    lblTotal: TLabel;
    lblRTotal: TLabel;
    GroupBox1: TGroupBox;
    btnGerar: TSpeedButton;
    GroupBox2: TGroupBox;
    Label2: TLabel;
    Med_DataFinal: TMaskEdit;
    Med_DataInicio: TMaskEdit;
    dbgDados: TDBGrid;
    dsDados: TDataSource;
    qryDados: TADOQuery;
    SaveDialog1: TSaveDialog;
    qryExec: TADOQuery;
    GroupBox3: TGroupBox;
    Label3: TLabel;
    Med_DataFinal2: TMaskEdit;
    Med_DataInicio2: TMaskEdit;
    procedure FormShow(Sender: TObject);
    procedure Med_DataInicioExit(Sender: TObject);
    procedure Med_DataFinalExit(Sender: TObject);
    procedure Med_DataFinal2Exit(Sender: TObject);
    procedure Med_DataInicio2Exit(Sender: TObject);
    procedure btnGerarClick(Sender: TObject);
    procedure btnExportarClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmRelatorio16: TfrmRelatorio16;
  parametro: String;
implementation

uses Principal;

{$R *.dfm}

procedure TfrmRelatorio16.FormShow(Sender: TObject);
begin
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

  Med_DataFinal.Text := '  /  /    ';
  Med_DataInicio.Text := '  /  /    ';
  Med_DataFinal2.Text := '  /  /    ';
  Med_DataInicio2.Text := '  /  /    ';
end;

procedure TfrmRelatorio16.Med_DataInicioExit(Sender: TObject);
var d_Data : TDateTime;
begin
  try
    d_Data := StrToDateTime(Med_DataInicio.Text);
    Med_DataInicio.Text := FormatDateTime('dd/mm/yyyy',d_Data);
  except
    ShowMessage('Data inválida');
    Med_DataInicio.Text := '  /  /    ';
    Med_DataInicio.SetFocus;
    Abort;
  end;
end;

procedure TfrmRelatorio16.Med_DataFinalExit(Sender: TObject);
var d_Data : TDateTime;
begin
  try
    d_Data := StrToDateTime(Med_DataFinal.Text);
    Med_DataFinal.Text := FormatDateTime('dd/mm/yyyy',d_Data);
  except
    ShowMessage('Data inválida');
    Med_DataFinal.Text := '  /  /    ';
    Med_DataFinal.SetFocus;
    Abort;
  end;
end;

procedure TfrmRelatorio16.Med_DataFinal2Exit(Sender: TObject);
var d_Data : TDateTime;
begin
  try
    d_Data := StrToDateTime(Med_DataFinal2.Text);
    Med_DataFinal2.Text := FormatDateTime('dd/mm/yyyy',d_Data);
  except
    ShowMessage('Data inválida');
    Med_DataFinal2.Text := '  /  /    ';
    Med_DataFinal2.SetFocus;
    Abort;
  end;
end;

procedure TfrmRelatorio16.Med_DataInicio2Exit(Sender: TObject);
var d_Data : TDateTime;
begin
  try
    d_Data := StrToDateTime(Med_DataInicio2.Text);
    Med_DataInicio2.Text := FormatDateTime('dd/mm/yyyy',d_Data);
  except
    ShowMessage('Data inválida');
    Med_DataInicio2.Text := '  /  /    ';
    Med_DataInicio2.SetFocus;
    Abort;
  end;
end;

procedure TfrmRelatorio16.btnGerarClick(Sender: TObject);
begin
  Med_DataInicio.OnExit(Sender);
  Med_DataFinal.OnExit(Sender);
  Med_DataInicio2.OnExit(Sender);
  Med_DataFinal2.OnExit(Sender);

  parametro := 'Periodo Adesão: '+Med_DataInicio.Text+' - '+Med_DataFinal.Text+' Periodo Conciliação: '+Med_DataInicio2.Text+' - '+Med_DataFinal2.Text;
  with qryDados do
  begin
  Close;
  SQL.Clear;
  SQL.Add(frmPrincipal.memSQL.Text);
  SQL.Text := StringReplace(SQL.Text,'&DATA1',Med_DataInicio.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
  SQL.Text := StringReplace(SQL.Text,'&DATA2',Med_DataFinal.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
  SQL.Text := StringReplace(SQL.Text,'&DATA3',Med_DataInicio2.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
  SQL.Text := StringReplace(SQL.Text,'&DATA4',Med_DataFinal2.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
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
end;

procedure TfrmRelatorio16.btnExportarClick(Sender: TObject);
var
F: TextFile;
P1, P2, P3, P4, P5, P6, P7, P8, P9, P10, P11, cLinha : String;
begin
 if qryDados.RecordCount > 0 then
 begin
   if frmPrincipal.lblCodRelatorio.Caption = '70' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_SEG' +';'+ 'NOME' +';'+ 'DT_INCL' +';'+ 'CORRETOR' +';'+ 'CPF_CORRETOR' +';'+ 'SUPERVISOR' +';'+ 'CPF_SUPERVISOR' +';'+ 'TELEFONE_SUB' +';'+ 'GRUPO_CARENCIA' +';'+ 'VEND_PLATAFORMA' +';'+ 'MENSALIDADE');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P2  := qryDados.FIELDBYNAME('NOME').AsString;
          P3  := qryDados.FIELDBYNAME('DT_INCL').AsString;
          P4  := qryDados.FIELDBYNAME('CORRETOR').AsString;
          P5  := qryDados.FIELDBYNAME('CPF_CORRETOR').AsString;
          P6  := qryDados.FIELDBYNAME('SUPERVISOR').AsString;
          P7  := qryDados.FIELDBYNAME('CPF_SUPERVISOR').AsString;
          P8  := qryDados.FIELDBYNAME('TELEFONE').AsString;
          P9  := qryDados.FIELDBYNAME('GRUPO_CARENCIA').AsString;
          P10 := qryDados.FIELDBYNAME('CORRETOR_PLATAFORMA').AsString;
          P11 := qryDados.FIELDBYNAME('MENSALIDADE').AsString;

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
