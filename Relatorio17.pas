 unit Relatorio17;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, StdCtrls, Mask, Grids, DBGrids, Buttons;

type
  TfrmRelatorio17 = class(TForm)
    Label1: TLabel;
    btnExportar: TSpeedButton;
    btnImprimir: TSpeedButton;
    lblTotal: TLabel;
    lblRTotal: TLabel;
    dbgDados: TDBGrid;
    GroupBox1: TGroupBox;
    btnGerar: TSpeedButton;
    qryDados: TADOQuery;
    dsDados: TDataSource;
    SaveDialog1: TSaveDialog;
    qryExec: TADOQuery;
    edtMensalidade: TEdit;
    Label2: TLabel;
    edtPercentual: TEdit;
    Label3: TLabel;
    edtBoleto: TEdit;
    Label4: TLabel;
    lblValorTotal: TLabel;
    Label6: TLabel;
    procedure FormShow(Sender: TObject);
    procedure edtBoletoKeyPress(Sender: TObject; var Key: Char);
    procedure edtPercentualKeyPress(Sender: TObject; var Key: Char);
    procedure edtMensalidadeKeyPress(Sender: TObject; var Key: Char);
    procedure btnGerarClick(Sender: TObject);
    procedure btnExportarClick(Sender: TObject);
    procedure qryDadosAfterOpen(DataSet: TDataSet);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmRelatorio17: TfrmRelatorio17;
  parametro: String;
implementation

{$R *.dfm}
uses Principal, DataModule, MaskUtils;
procedure TfrmRelatorio17.FormShow(Sender: TObject);
begin
  if frmPrincipal.lblImprime.Caption = 'T' then //Habilita bot�o Imprimir
  begin
  btnImprimir.Enabled := True;
  end else
  begin
  btnImprimir.Enabled := False;
  end;

  if frmPrincipal.lblExporta.Caption = 'T' then //Habilita bot�o Exportar
  begin
  btnExportar.Enabled := True;
  end else
  begin
  btnExportar.Enabled := False;
  end;

  edtMensalidade.Text   := '';
  edtPercentual.Text    := '';
  edtBoleto.Text        := '';
  lblRTotal.Caption     := '0';
  lblValorTotal.Caption := '0';
end;

procedure TfrmRelatorio17.edtBoletoKeyPress(Sender: TObject; var Key: Char);
begin
  if not (key in ['0'..'9',',',#8,^V,^C,^X]) then key :=#0;
end;

procedure TfrmRelatorio17.edtPercentualKeyPress(Sender: TObject;
  var Key: Char);
begin
  if not (key in ['0'..'9',#8,^V,^C,^X]) then key :=#0;
end;

procedure TfrmRelatorio17.edtMensalidadeKeyPress(Sender: TObject;
  var Key: Char);
begin
  if not (key in ['0'..'9',',',#8,^V,^C,^X]) then key :=#0;
end;

procedure TfrmRelatorio17.btnGerarClick(Sender: TObject);
var
  Mensalidade, Percentual : String;
begin
  Mensalidade:= '';
  Percentual := '';
  parametro  := '';

  if frmPrincipal.lblCodRelatorio.Caption = '90' then
  begin
    if (edtMensalidade.Text <> '') and (edtPercentual.Text <> '') and (edtBoleto.Text <> '') then
    begin
      Mensalidade := StringReplace(edtMensalidade.Text,',','.',[rfReplaceAll]);
      Percentual  := StringReplace(edtPercentual.Text,',','.',[rfReplaceAll]);

      parametro := 'Mensalidade: '+ edtMensalidade.Text + ' Percent: ' + edtPercentual.Text + ' Boletos: ' + edtBoleto.Text;
      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);
      SQL.Text := StringReplace(SQL.Text,'&MENS',Mensalidade,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&PERC',Percentual,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&BOLETO',edtBoleto.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      Open;
      end;


      with qryExec do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL2.Text);
      SQL.Text := StringReplace(SQL.Text,'&MENS',Mensalidade,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&PERC',Percentual,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&BOLETO',edtBoleto.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      Open;
      end;

    lblRTotal.Caption     := IntToStr(qryDados.RecordCount);
    lblValorTotal.Caption := qryExec.FieldByName('VLR_COMISSAO').AsString;

      with qryExec do // INSERIR LOG DE VISUALIZA��O
      begin
      Close;
      SQL.Clear;
      SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, PARAMETRO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'V'+#39+','+#39+parametro+#39+', SYSDATE)');
      ExecSQL;
      end;

    end else
    begin
    ShowMessage('Favor informar todos os par�metros.');
    end;
  end;
end;

procedure TfrmRelatorio17.btnExportarClick(Sender: TObject);
var
F: TextFile;
P1, P2, P3, P4, P5, P6, P7, P8, P9, P10, P11, P12, P13, cLinha  : String;
begin
  if qryDados.RecordCount > 0 then
  begin

  if frmPrincipal.lblCodRelatorio.Caption = '90' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_SEG' +';'+ 'NOME' +';'+ 'EMPRESA' +';'+ 'UNIDADE' +';'+ 'BOLETO' +';'+ 'VENC_ATUAL' +';'+ 'VENC_ORIGINAL' +';'+ 'DT_CONCIL' +';'+ 'MENS_BOLETO' +';'+ 'MENS_INFORMADO' +';'+ 'PERCENTUAL' +';'+ 'VLR_COMISSAO' +';'+ 'COMPETENCIA');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P2  := qryDados.FIELDBYNAME('NOME').AsString;
          P3  := qryDados.FIELDBYNAME('EMPRESA').AsString;
          P4  := qryDados.FIELDBYNAME('UNIDADE').AsString;
          P5  := qryDados.FIELDBYNAME('COD_BOL').AsString;
          P6  := qryDados.FIELDBYNAME('VENC_ATUAL').AsString;
          P7  := qryDados.FIELDBYNAME('VENC_ORIGINAL').AsString;
          P8  := qryDados.FIELDBYNAME('DT_CONCIL').AsString;
          P9  := qryDados.FIELDBYNAME('MENS_BOLETO').AsString;
          P10 := qryDados.FIELDBYNAME('MENS_INFORMADO').AsString;
          P11 := qryDados.FIELDBYNAME('PERCENTUAL').AsString;
          P12 := qryDados.FIELDBYNAME('VLR_COMISSAO').AsString;
          P13 := qryDados.FIELDBYNAME('COMPETENCIA').AsString;

          P9  := FormatFloat('##0.00',StrToFloat(P9));
          P10 := FormatFloat('##0.00',StrToFloat(P9));
          P12 := FormatFloat('##0.00',StrToFloat(P9));

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

    with qryExec do // INSERIR LOG DE EXPORTA��O
    begin
    Close;
    SQL.Clear;
    SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, PARAMETRO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'E'+#39+', '+#39+parametro+#39+', SYSDATE)');
    ExecSQL;
    end;
  end;
end;

procedure TfrmRelatorio17.qryDadosAfterOpen(DataSet: TDataSet);
begin
  frmPrincipal.DimensionarGrid(dbgDados,Self);

  if frmPrincipal.lblCodRelatorio.Caption = '90' then
  begin
    TCurrencyField(qryDados.FieldByName('MENS_BOLETO')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('MENS_INFORMADO')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('VLR_COMISSAO')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('COD_SEG')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_BOL')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('VENC_ATUAL')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('VENC_ORIGINAL')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DT_CONCIL')).Alignment := taCenter;    
    TCurrencyField(qryDados.FieldByName('PERCENTUAL')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COMPETENCIA')).Alignment := taCenter;
  end;

end;

end.
