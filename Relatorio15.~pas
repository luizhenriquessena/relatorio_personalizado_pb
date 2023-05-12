unit Relatorio15;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Mask, StdCtrls, Grids, DBGrids, Buttons, DB, ADODB;

type
  TfrmRelatorio15 = class(TForm)
    btnImprimir: TSpeedButton;
    btnExportar: TSpeedButton;
    dbgDados: TDBGrid;
    GroupBox1: TGroupBox;
    btnGerar: TSpeedButton;
    lblOperador: TLabel;
    edtOperador: TEdit;
    Med_MesInicio: TMaskEdit;
    Label1: TLabel;
    qryExec: TADOQuery;
    Label2: TLabel;
    qryDados: TADOQuery;
    dsDados: TDataSource;
    SaveDialog1: TSaveDialog;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure Med_MesInicioExit(Sender: TObject);
    procedure btnGerarClick(Sender: TObject);
    procedure qryDadosAfterOpen(DataSet: TDataSet);
    procedure btnExportarClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmRelatorio15: TfrmRelatorio15;
  parametro: String;
  MesInicio : String;
implementation

{$R *.dfm}
uses Principal, DataModule;

procedure TfrmRelatorio15.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
qryExec.Close;
qryDados.Close;
end;

procedure TfrmRelatorio15.FormShow(Sender: TObject);
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

  if (frmPrincipal.lblCodRelatorio.Caption = '114') or
     (frmPrincipal.lblCodRelatorio.Caption = '115') then
  begin
    lblOperador.Visible := false;
    edtOperador.Visible := false;
    Abort;
  end;

  lblOperador.Visible := true;
  edtOperador.Visible := true;
end;

procedure TfrmRelatorio15.Med_MesInicioExit(Sender: TObject);
var
d_Data : TDate;
begin
  try
    d_Data := StrToDate('01/'+Med_MesInicio.Text);
    MesInicio := FormatDateTime('yyyy/mm',d_Data);
  except
    ShowMessage('Data base inválida');
    Med_MesInicio.Text := '  /    ';
    Med_MesInicio.SetFocus;
  end;
end;

procedure TfrmRelatorio15.btnGerarClick(Sender: TObject);
begin
  Med_MesInicio.OnExit(Sender);
  parametro := '';

  if (frmPrincipal.lblCodRelatorio.Caption = '114') or
     (frmPrincipal.lblCodRelatorio.Caption = '115') then
  begin
    if Med_MesInicio.Text <> '  /    ' then
    begin
      parametro := 'MesConciliação: '+MesInicio;
      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);
      SQL.Text := StringReplace(SQL.Text,'&MESANO',Med_MesInicio.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      Open;
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
    ShowMessage('Favor informar o período.');
    end;
  end else
  begin
    if (Med_MesInicio.Text <> '  /    ') and (edtOperador.Text <> '') then
    begin
      parametro := 'MesPagamento: '+MesInicio + ' Operador: ' + edtOperador.Text;
      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);
      SQL.Text := StringReplace(SQL.Text,'&MESANO',MesInicio,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&OPERADOR',MesInicio,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      Open;
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
    ShowMessage('Favor informar os parâmetros.');
    end;
  end;
end;

procedure TfrmRelatorio15.qryDadosAfterOpen(DataSet: TDataSet);
begin
frmPrincipal.DimensionarGrid(dbgDados,Self);

  if (frmPrincipal.lblCodRelatorio.Caption = '114') or (frmPrincipal.lblCodRelatorio.Caption = '115') then
  begin
    TCurrencyField(qryDados.FieldByName('MES_ANO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_VEND')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('RECEITA')).DisplayFormat := '#,##0';
  end;
end;

procedure TfrmRelatorio15.btnExportarClick(Sender: TObject);
var
F: TextFile;
P1, P2, P3, P4, P5, P6, P7, P8, P9, P10, P11, P12, P13, P14, P15, cLinha  : String;
begin
 if qryDados.RecordCount > 0 then
 begin
  if (frmPrincipal.lblCodRelatorio.Caption = '114') or (frmPrincipal.lblCodRelatorio.Caption = '115') then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'MES_ANO' +';'+ 'COD_VEND' +';'+ 'VENDEDOR' +';'+ 'RECEITA');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('MES_ANO').AsString;
          P2  := qryDados.FIELDBYNAME('COD_VEND').AsString;
          P3  := qryDados.FIELDBYNAME('VENDEDOR').AsString;
          P4  := qryDados.FIELDBYNAME('RECEITA').AsString;

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
