unit Relatorio18;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, StdCtrls, Buttons, Grids, DBGrids, COMOBJ;

type
  TfrmRelatorio18 = class(TForm)
    DBGrid1: TDBGrid;
    DBGrid2: TDBGrid;
    DBGrid3: TDBGrid;
    GroupBox1: TGroupBox;
    btnGerar: TSpeedButton;
    lblCodigoPrestador: TLabel;
    edtCodigoPrest: TEdit;
    edtNomePrest: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    btnExportar: TSpeedButton;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    lblTotalEspec: TLabel;
    lblTotalCC: TLabel;
    lblTotalProced: TLabel;
    qryPrestador: TADOQuery;
    qryEspec: TADOQuery;
    dsEspec: TDataSource;
    dsCc: TDataSource;
    qryCC: TADOQuery;
    dsProcedimento: TDataSource;
    qryProcedimento: TADOQuery;
    Label7: TLabel;
    qryPrestadorCOD_HOSP: TWideStringField;
    qryPrestadorRAZ_SOC: TWideStringField;
    qryPrestadorFANTASIA: TWideStringField;
    qryPrestadorCNPJ: TWideStringField;
    qryPrestadorCNES: TWideStringField;
    qryPrestadorCEP: TWideStringField;
    qryPrestadorENDERECO: TWideStringField;
    qryPrestadorCIDADE: TWideStringField;
    qryPrestadorBAIRRO: TWideStringField;
    qryPrestadorESTADO: TWideStringField;
    qryPrestadorNM_DIRETOR_TECNICO: TWideStringField;
    qryPrestadorNOME_RESP_TEC: TWideStringField;
    qryPrestadorCOD_BANCO: TWideStringField;
    qryPrestadorBANCO: TWideStringField;
    qryPrestadorAGENCIA: TWideStringField;
    qryPrestadorDIG_AGENCIA: TWideStringField;
    qryPrestadorCONTA: TWideStringField;
    qryPrestadorDIG_CONTA: TWideStringField;
    qryPrestadorNOME_DEP: TWideStringField;
    qryExec: TADOQuery;
    qryEspecESPECIALIDADE: TWideStringField;
    qryCCMEDICO: TWideStringField;
    qryCCCRM: TWideStringField;
    qryCCCPF: TWideStringField;
    qryCCESPECIALIDADE: TWideStringField;
    qryProcedimentoCODIGO: TWideStringField;
    qryProcedimentoDESCRICAO: TWideStringField;
    procedure edtCodigoPrestExit(Sender: TObject);
    procedure edtCodigoPrestKeyPress(Sender: TObject; var Key: Char);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure btnGerarClick(Sender: TObject);
    procedure btnExportarClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmRelatorio18: TfrmRelatorio18;
  parametro: String;
implementation

uses DataModule, Principal;
{$R *.dfm}

procedure TfrmRelatorio18.edtCodigoPrestExit(Sender: TObject);
begin
  if edtCodigoPrest.Text <> '' then
  begin
    qryPrestador.Close;
    qryPrestador.Parameters.ParamByName('cod_ent').Value := '%'+edtCodigoPrest.Text+'%';
    qryPrestador.Open;

    if qryPrestador.RecordCount > 0 then
    begin
      edtCodigoPrest.Text := qryPrestador.FieldByName('COD_HOSP').AsString;
      lblCodigoPrestador.Caption := qryPrestador.FieldByName('COD_HOSP').AsString;
      edtNomePrest.Text := qryPrestador.FieldByName('FANTASIA').AsString;
    end else
    begin
      ShowMessage('Código do prestador inválido');
      lblCodigoPrestador.Caption := '0';
    end;
  end;
end;

procedure TfrmRelatorio18.edtCodigoPrestKeyPress(Sender: TObject; var Key: Char);
begin
    if not (key in ['0'..'9',Chr(8), Chr(3), Chr(22), Chr(24), Chr(44)]) then
      key := #0;
end;

procedure TfrmRelatorio18.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  qryPrestador.Close;
  qryEspec.Close;
  qryCC.Close;
  qryProcedimento.Close;

  lblCodigoPrestador.Caption := '0';
  lblTotalEspec.Caption := '0';
  lblTotalCC.Caption := '0';
  lblTotalProced.Caption := '0';
end;

procedure TfrmRelatorio18.FormShow(Sender: TObject);
begin
  edtCodigoPrest.Text := '';
  edtNomePrest.Text := '';
  lblCodigoPrestador.Caption := '0';

  if frmPrincipal.lblExporta.Caption = 'T' then //Habilita botão Exportar
  begin
  btnExportar.Enabled := True;
  end else
  begin
  btnExportar.Enabled := False;
  end;

end;

procedure TfrmRelatorio18.btnGerarClick(Sender: TObject);
begin
  parametro := '';

  if lblCodigoPrestador.Caption <> '0' then
  begin
    parametro := 'Prestador: '+lblCodigoPrestador.Caption;

    // Especialidade
    qryEspec.Close;
    qryEspec.Parameters.ParamByName('cod_ent').Value := lblCodigoPrestador.Caption;
    qryEspec.Open;

    lblTotalEspec.Caption := IntToStr(qryEspec.RecordCount); // Preencher total de especialidades

    // Corpo Clínico
    qryCC.Close;
    qryCC.Parameters.ParamByName('cod_ent').Value := lblCodigoPrestador.Caption;
    qryCC.Open;

    lblTotalCC.Caption := IntToStr(qryCC.RecordCount); // Preencher total de médicos no corpo clinico

    // Procedimentos
    qryProcedimento.Close;
    qryProcedimento.Parameters.ParamByName('cod_ent').Value := lblCodigoPrestador.Caption;
    qryProcedimento.Open;

    lblTotalProced.Caption := IntToStr(qryProcedimento.RecordCount); // Preencher total de procedimentos cobertos pelo prestador

  end else
    ShowMessage('Favor selecionar um prestador.');
end;

procedure TfrmRelatorio18.btnExportarClick(Sender: TObject);
var
  vPlanilha: Variant;
  vExcel: Variant;
  vLinha: Integer;
begin
  if (qryEspec.RecordCount > 0) or
     (qryCC.RecordCount > 0) or
     (qryProcedimento.RecordCount > 0) then
  begin
    vExcel := CreateOleObject('Excel.Application');
    vExcel.visible := false;
    vExcel.workbooks.Add;


    // Dados Prestador
    vPlanilha := vExcel.workbooks[1].sheets[1];
    vPlanilha.name := 'Dados Prestador';

    vPlanilha.Range['A1','B1'].merge(EmptyParam); //Mesclar celulas
    vPlanilha.cells[1,1] := qryPrestadorFANTASIA.AsString;
    vPlanilha.cells[1,1].interior.Color := $00D9D9D9; // Alterar cor de fundo
    vPlanilha.cells[1,1].font.bold := true; // Colocar o texto em negrito
    vPlanilha.cells[1,1].horizontalAlignment := 3;// Alinhas celula (2 Esquerda, 3 Centro e 4 Direita)
    vPlanilha.cells[1,1].font.size := 16; // Tamanho do texto
    vPlanilha.Range['A2','B2'].merge(EmptyParam); //Mesclar celulas
    vPlanilha.cells[2,1] := 'ATUALIZACAO CADASTRAL';
    vPlanilha.cells[2,1].font.bold := true; // Colocar o texto em negrito
    vPlanilha.cells[2,1].horizontalAlignment := 3;// Alinhas celula (2 Esquerda, 3 Centro e 4 Direita)
    vPlanilha.cells[3,1] := '';

    vPlanilha.Range['A4','B4'].merge(EmptyParam); //Mesclar celulas
    vPlanilha.cells[4,1] := 'DADOS DA EMPRESA';
    vPlanilha.cells[4,1].interior.Color := $00404040; // Alterar cor de fundo
    vPlanilha.cells[4,1].font.bold := true; // Colocar o texto em negrito
    vPlanilha.cells[4,1].font.color := $00FFFFFF; // Alterar a cor do texto
    vPlanilha.cells[4,1].horizontalAlignment := 3;// Alinhas celula (2 Esquerda, 3 Centro e 4 Direita)
    vPlanilha.Range['A5','A14'].font.bold := true; // Colocar uma seleção de celular em negrito
    vPlanilha.cells[5,1] := 'RAZÃO SOCIAL';
    vPlanilha.cells[5,2] := qryPrestadorRAZ_SOC.AsString;
    vPlanilha.cells[6,1] := 'NOME FANTASIA';
    vPlanilha.cells[6,2] := qryPrestadorFANTASIA.AsString;
    vPlanilha.cells[7,1] := 'CNPJ';
    vPlanilha.cells[7,2] := qryPrestadorCNPJ.AsString;
    vPlanilha.cells[8,1] := 'CNES';
    vPlanilha.cells[8,2] := qryPrestadorCNES.AsString;
    vPlanilha.cells[9,1] := 'CEP';
    vPlanilha.cells[9,2] := qryPrestadorCEP.AsString;
    vPlanilha.cells[10,1] := 'ENDEREÇO';
    vPlanilha.cells[10,2] := qryPrestadorENDERECO.AsString;
    vPlanilha.cells[11,1] := 'CIDADE';
    vPlanilha.cells[11,2] := qryPrestadorCIDADE.AsString;
    vPlanilha.cells[12,1] := 'BAIRRO';
    vPlanilha.cells[12,2] := qryPrestadorBAIRRO.AsString;
    vPlanilha.cells[13,1] := 'ESTADO';
    vPlanilha.cells[13,2] := qryPrestadorESTADO.AsString;

    vPlanilha.Range['A14','B14'].merge(EmptyParam); //Mesclar celulas
    vPlanilha.cells[14,1] := 'DIRETORES';
    vPlanilha.cells[14,1].interior.Color := $00AEAAAA; // Alterar cor de fundo
    vPlanilha.cells[14,1].font.bold := true; // Colocar o texto em negrito
    vPlanilha.cells[14,1].horizontalAlignment := 3;// Alinhas celula (2 Esquerda, 3 Centro e 4 Direita)
    vPlanilha.Range['A15','A20'].font.bold := true; // Colocar uma seleção de celular em negrito
    vPlanilha.cells[15,1] := 'NOME';
    vPlanilha.cells[15,2] := qryPrestadorNM_DIRETOR_TECNICO.AsString;
    vPlanilha.cells[16,1] := 'TELEFONE';
    vPlanilha.cells[17,1] := 'E-MAIL';
    vPlanilha.cells[18,1] := 'NOME 2';
    vPlanilha.cells[19,1] := 'TELEFONE';
    vPlanilha.cells[20,1] := 'E-MAIL';

    vPlanilha.Range['A21','B21'].merge(EmptyParam); //Mesclar celulas
    vPlanilha.cells[21,1] := 'REPRESENTANTE LEGAL';
    vPlanilha.cells[21,1].interior.Color := $00AEAAAA; // Alterar cor de fundo
    vPlanilha.cells[21,1].font.bold := true; // Colocar o texto em negrito
    vPlanilha.cells[21,1].horizontalAlignment := 3;// Alinhas celula (2 Esquerda, 3 Centro e 4 Direita)
    vPlanilha.Range['A22','A23'].font.bold := true; // Colocar uma seleção de celular em negrito
    vPlanilha.cells[22,1] := 'NOME';
    vPlanilha.cells[23,1] := 'CPF';

    vPlanilha.Range['A24','B24'].merge(EmptyParam); //Mesclar celulas
    vPlanilha.cells[24,1] := 'RESPONSÁVEL TÉCNICO';
    vPlanilha.cells[24,1].interior.Color := $00AEAAAA; // Alterar cor de fundo
    vPlanilha.cells[24,1].font.bold := true; // Colocar o texto em negrito
    vPlanilha.cells[24,1].horizontalAlignment := 3;// Alinhas celula (2 Esquerda, 3 Centro e 4 Direita)
    vPlanilha.Range['A25','A26'].font.bold := true; // Colocar uma seleção de celular em negrito
    vPlanilha.cells[25,1] := 'NOME';
    vPlanilha.cells[25,2] := qryPrestadorNOME_RESP_TEC.AsString;
    vPlanilha.cells[26,1] := 'CPF';

    vPlanilha.Range['A27','B27'].merge(EmptyParam); //Mesclar celulas
    vPlanilha.cells[27,1] := 'RESPONSÁVEL ADMINISTRAÇÃO';
    vPlanilha.cells[27,1].interior.Color := $00AEAAAA; // Alterar cor de fundo
    vPlanilha.cells[27,1].font.bold := true; // Colocar o texto em negrito
    vPlanilha.cells[27,1].horizontalAlignment := 3;// Alinhas celula (2 Esquerda, 3 Centro e 4 Direita)
    vPlanilha.Range['A28','A30'].font.bold := true; // Colocar uma seleção de celular em negrito
    vPlanilha.cells[28,1] := 'NOME';
    vPlanilha.cells[29,1] := 'TELEFONE';
    vPlanilha.cells[30,1] := 'E-MAIL';

    vPlanilha.Range['A31','B31'].merge(EmptyParam); //Mesclar celulas
    vPlanilha.cells[31,1] := 'RESPONSÁVEL FATURAMENTO';
    vPlanilha.cells[31,1].interior.Color := $00AEAAAA; // Alterar cor de fundo
    vPlanilha.cells[31,1].font.bold := true; // Colocar o texto em negrito
    vPlanilha.cells[31,1].horizontalAlignment := 3;// Alinhas celula (2 Esquerda, 3 Centro e 4 Direita)
    vPlanilha.Range['A32','A34'].font.bold := true; // Colocar uma seleção de celular em negrito
    vPlanilha.cells[32,1] := 'NOME';
    vPlanilha.cells[33,1] := 'TELEFONE';
    vPlanilha.cells[34,1] := 'E-MAIL';

    vPlanilha.Range['A35','B35'].merge(EmptyParam); //Mesclar celulas
    vPlanilha.cells[35,1] := 'RESPONSÁVEL RECEPÇÃO';
    vPlanilha.cells[35,1].interior.Color := $00AEAAAA; // Alterar cor de fundo
    vPlanilha.cells[35,1].font.bold := true; // Colocar o texto em negrito
    vPlanilha.Range['A36','A38'].font.bold := true; // Colocar uma seleção de celular em negrito
    vPlanilha.cells[35,1].horizontalAlignment := 3;// Alinhas celula (2 Esquerda, 3 Centro e 4 Direita)
    vPlanilha.cells[36,1] := 'NOME';
    vPlanilha.cells[37,1] := 'TELEFONE';
    vPlanilha.cells[38,1] := 'E-MAIL';
    vPlanilha.cells[39,1] := '';
    
    vPlanilha.Range['A40','B40'].merge(EmptyParam); //Mesclar celulas
    vPlanilha.cells[40,1] := 'DADOS BANCÁRIOS';
    vPlanilha.cells[40,1].interior.Color := $00404040; // Alterar cor de fundo
    vPlanilha.cells[40,1].font.bold := true; // Colocar o texto em negrito
    vPlanilha.cells[40,1].font.color := $00FFFFFF; // Alterar a cor do texto
    vPlanilha.Range['A41','A45'].font.bold := true; // Colocar uma seleção de celular em negrito
    vPlanilha.cells[40,1].horizontalAlignment := 3;// Alinhas celula (2 Esquerda, 3 Centro e 4 Direita)
    vPlanilha.cells[41,1] := 'NOME DO BANCO';
    vPlanilha.cells[41,2] := qryPrestadorBANCO.AsString;
    vPlanilha.cells[42,1] := 'BANCO Nº';
    vPlanilha.cells[42,2].NumberFormat := '@'; // Formatar celula para texto, permitindo zero a esquerda. Precisa ser antes de preencher com valor
    vPlanilha.cells[42,2] := qryPrestadorCOD_BANCO.AsString;
    vPlanilha.cells[43,1] := 'AGÊNCIA';
    vPlanilha.cells[43,2].NumberFormat := '@'; // Formatar celula para texto, permitindo zero a esquerda. Precisa ser antes de preencher com valor
    vPlanilha.cells[43,2] := qryPrestadorAGENCIA.AsString + ' - ' + qryPrestadorDIG_AGENCIA.AsString;
    vPlanilha.cells[44,1] := 'CONTA CORRENTE';
    vPlanilha.cells[44,2].NumberFormat := '@'; // Formatar celula para texto, permitindo zero a esquerda. Precisa ser antes de preencher com valor
    vPlanilha.cells[44,2] := qryPrestadorCONTA.AsString + ' - ' + qryPrestadorDIG_CONTA.AsString;
    vPlanilha.cells[45,1] := 'NOME';
    vPlanilha.cells[45,1] := qryPrestadorNOME_DEP.AsString;

    vExcel.columns.autofit;
    
    //  Especialidade e Corpo Clinico
    vPlanilha := vExcel.workbooks[1].sheets.add(null, vPlanilha);
    vPlanilha.Name := 'Especialidade e Corpo Clínico';

    // ## Especialidades
    vPlanilha.Range['A1','C1'].merge(EmptyParam); //Mesclar celulas
    vPlanilha.cells[1,1] := qryPrestadorFANTASIA.AsString;
    vPlanilha.cells[1,1].interior.Color := $00D9D9D9; // Alterar cor de fundo
    vPlanilha.cells[1,1].font.bold := true; // Colocar o texto em negrito
    vPlanilha.cells[1,1].horizontalAlignment := 3;// Alinhas celula (2 Esquerda, 3 Centro e 4 Direita)
    vPlanilha.cells[1,1].font.size := 16; // Tamanho do texto
    vPlanilha.Range['A2','C2'].merge(EmptyParam); //Mesclar celulas
    vPlanilha.cells[2,1] := 'ATUALIZACAO CADASTRAL';
    vPlanilha.cells[2,1].font.bold := true; // Colocar o texto em negrito
    vPlanilha.cells[2,1].horizontalAlignment := 3;// Alinhas celula (2 Esquerda, 3 Centro e 4 Direita)
    vPlanilha.cells[3,1] := '';

//    vPlanilha.cells[1,4].ColumnWidth := 100; // Definir largura da coluna (Informar a primeira linha e a posição da coluna)

    vPlanilha.cells[4,1] := 'DADOS CADASTRAIS';
    vPlanilha.cells[4,1].interior.Color := $00404040; // Alterar cor de fundo
    vPlanilha.cells[4,1].font.bold := true; // Colocar o texto em negrito
    vPlanilha.cells[4,1].font.color := $00FFFFFF; // Alterar a cor do texto
    vPlanilha.cells[4,1].horizontalAlignment := 3;// Alinhas celula (2 Esquerda, 3 Centro e 4 Direita)
    vPlanilha.cells[5,1] := 'ESPECIALIDADE';
    vPlanilha.cells[5,1].interior.Color := $00AEAAAA; // Alterar cor de fundo
    vPlanilha.cells[5,1].font.bold := true; // Colocar o texto em negrito
    vPlanilha.cells[5,1].horizontalAlignment := 3;// Alinhas celula (2 Esquerda, 3 Centro e 4 Direita)
    vPlanilha.Range['B4','C4'].merge(EmptyParam); //Mesclar celulas
    vPlanilha.cells[4,2] := 'ATENDIMENTO';
    vPlanilha.cells[4,2].interior.Color := $00404040; // Alterar cor de fundo
    vPlanilha.cells[4,2].font.bold := true; // Colocar o texto em negrito
    vPlanilha.cells[4,2].font.color := $00FFFFFF; // Alterar a cor do texto
    vPlanilha.cells[4,2].horizontalAlignment := 3;// Alinhas celula (2 Esquerda, 3 Centro e 4 Direita)
    vPlanilha.cells[5,2] := '     SIM     ';
    vPlanilha.cells[5,2].interior.Color := $00AEAAAA; // Alterar cor de fundo
    vPlanilha.cells[5,2].font.bold := true; // Colocar o texto em negrito
    vPlanilha.cells[5,2].horizontalAlignment := 3;// Alinhas celula (2 Esquerda, 3 Centro e 4 Direita)
    vPlanilha.cells[5,3] := '     NÃO     ';
    vPlanilha.cells[5,3].interior.Color := $00AEAAAA; // Alterar cor de fundo
    vPlanilha.cells[5,3].font.bold := true; // Colocar o texto em negrito
    vPlanilha.cells[5,3].horizontalAlignment := 3;// Alinhas celula (2 Esquerda, 3 Centro e 4 Direita)

    try
      vLinha := 6;
      while not  qryEspec.Eof do
      begin
        vPlanilha.cells[vLinha,1] := qryEspecESPECIALIDADE.AsString;
        vPlanilha.cells[vLinha,1].font.bold := true;
        vLinha:=vLinha+1;
        qryEspec.Next;
      end;
    finally end;
    vExcel.columns.autofit;

    // ## Corpo Clínico
    vPlanilha.Range['E1','J1'].merge(EmptyParam); //Mesclar celulas
    vPlanilha.cells[1,5]  := qryPrestadorFANTASIA.AsString;
    vPlanilha.cells[1,5].interior.Color := $00D9D9D9; // Alterar cor de fundo
    vPlanilha.cells[1,5].font.bold := true; // Colocar o texto em negrito
    vPlanilha.cells[1,5].horizontalAlignment := 3;// Alinhas celula (2 Esquerda, 3 Centro e 4 Direita)
    vPlanilha.cells[1,5].font.size := 16; // Tamanho do texto
    vPlanilha.Range['E2','J2'].merge(EmptyParam); //Mesclar celulas
    vPlanilha.cells[2,5] := 'ATUALIZACAO CADASTRAL';
    vPlanilha.cells[2,5].font.bold := true; // Colocar o texto em negrito
    vPlanilha.cells[2,5].horizontalAlignment := 3;// Alinhas celula (2 Esquerda, 3 Centro e 4 Direita)
    vPlanilha.cells[3,5] := '';
    vPlanilha.Range['E4','J4'].merge(EmptyParam); //Mesclar celulas
    vPlanilha.cells[4,5] := 'DADOS CADASTRAIS';
    vPlanilha.cells[4,5].interior.Color := $00404040; // Alterar cor de fundo
    vPlanilha.cells[4,5].font.bold := true; // Colocar o texto em negrito
    vPlanilha.cells[4,5].font.color := $00FFFFFF; // Alterar a cor do texto
    vPlanilha.cells[4,5].horizontalAlignment := 3;// Alinhas celula (2 Esquerda, 3 Centro e 4 Direita)
    vPlanilha.cells[5,5] := 'MÉDICOS';
    vPlanilha.cells[5,5].interior.Color := $00AEAAAA; // Alterar cor de fundo
    vPlanilha.cells[5,5].font.bold := true; // Colocar o texto em negrito
    vPlanilha.cells[5,5].horizontalAlignment := 3;// Alinhas celula (2 Esquerda, 3 Centro e 4 Direita)
    vPlanilha.cells[5,6] := 'CRM';
    vPlanilha.cells[5,6].interior.Color := $00AEAAAA; // Alterar cor de fundo
    vPlanilha.cells[5,6].font.bold := true; // Colocar o texto em negrito
    vPlanilha.cells[5,6].horizontalAlignment := 3;// Alinhas celula (2 Esquerda, 3 Centro e 4 Direita)
    vPlanilha.cells[5,7] := 'CPF';
    vPlanilha.cells[5,7].interior.Color := $00AEAAAA; // Alterar cor de fundo
    vPlanilha.cells[5,7].font.bold := true; // Colocar o texto em negrito
    vPlanilha.cells[5,7].horizontalAlignment := 3;// Alinhas celula (2 Esquerda, 3 Centro e 4 Direita)
    vPlanilha.cells[5,8] := 'ESPECIALIDADES';
    vPlanilha.cells[5,8].interior.Color := $00AEAAAA; // Alterar cor de fundo
    vPlanilha.cells[5,8].font.bold := true; // Colocar o texto em negrito
    vPlanilha.cells[5,8].horizontalAlignment := 3;// Alinhas celula (2 Esquerda, 3 Centro e 4 Direita)
    vPlanilha.cells[5,9] := '     SIM     ';
    vPlanilha.cells[5,9].interior.Color := $00AEAAAA; // Alterar cor de fundo
    vPlanilha.cells[5,9].font.bold := true; // Colocar o texto em negrito
    vPlanilha.cells[5,9].horizontalAlignment := 3;// Alinhas celula (2 Esquerda, 3 Centro e 4 Direita)
    vPlanilha.cells[5,10] := '     NÃO     ';
    vPlanilha.cells[5,10].interior.Color := $00AEAAAA; // Alterar cor de fundo
    vPlanilha.cells[5,10].font.bold := true; // Colocar o texto em negrito
    vPlanilha.cells[5,10].horizontalAlignment := 3;// Alinhas celula (2 Esquerda, 3 Centro e 4 Direita)

    try
      vLinha := 6;
      while not  qryCC.Eof do
      begin

        vPlanilha.cells[vLinha,5] := qryCCMEDICO.AsString;
        vPlanilha.cells[vLinha,5].font.bold := true; // Colocar o texto em negrito
        vPlanilha.cells[vLinha,6].NumberFormat := '@'; // Formatar celula para texto, permitindo zero a esquerda. Precisa ser antes de preencher com valor        
        vPlanilha.cells[vLinha,6] := qryCCCRM.AsString;
        vPlanilha.cells[vLinha,6].font.bold := true; // Colocar o texto em negrito
        vPlanilha.cells[vLinha,6].horizontalAlignment := 3; // Alinhas celula (2 Esquerda, 3 Centro e 4 Direita)
        vPlanilha.cells[vLinha,7] := qryCCCPF.AsString;
        vPlanilha.cells[vLinha,7].font.bold := true;
        vPlanilha.cells[vLinha,7].horizontalAlignment := 3; // Alinhas celula (2 Esquerda, 3 Centro e 4 Direita)
        vPlanilha.cells[vLinha,8] := qryCCESPECIALIDADE.AsString;
        vPlanilha.cells[vLinha,8].font.bold := true;

        vLinha:=vLinha+1;
        qryCC.Next;
      end;
    finally end;

    vExcel.columns.autofit;

    //  Procedimentos
    vPlanilha := vExcel.workbooks[1].sheets.add(null, vPlanilha);
    vPlanilha.Name := 'Procedimentos';

    vPlanilha.Range['A1','D1'].merge(EmptyParam); //Mesclar celulas
    vPlanilha.cells[1,1] := qryPrestadorFANTASIA.AsString;
    vPlanilha.cells[1,1].interior.Color := $00D9D9D9; // Alterar cor de fundo
    vPlanilha.cells[1,1].font.bold := true; // Colocar o texto em negrito
    vPlanilha.cells[1,1].horizontalAlignment := 3;// Alinhas celula (2 Esquerda, 3 Centro e 4 Direita)
    vPlanilha.cells[1,1].font.size := 16; // Tamanho do texto
    vPlanilha.Range['A2','D2'].merge(EmptyParam); //Mesclar celulas
    vPlanilha.cells[2,1] := 'ATUALIZACAO CADASTRAL';
    vPlanilha.cells[2,1].font.bold := true; // Colocar o texto em negrito
    vPlanilha.cells[2,1].horizontalAlignment := 3;// Alinhas celula (2 Esquerda, 3 Centro e 4 Direita)
    vPlanilha.cells[3,1] := '';

    vPlanilha.Range['A4','B4'].merge(EmptyParam); //Mesclar celulas
    vPlanilha.cells[4,1] := 'PROCEDIMENTOS CONTRATADOS';
    vPlanilha.cells[4,1].interior.Color := $00595959; // Alterar cor de fundo
    vPlanilha.cells[4,1].font.bold := true; // Colocar o texto em negrito
    vPlanilha.cells[4,1].font.color := $00FFFFFF; // Alterar a cor do texto
    vPlanilha.cells[4,1].horizontalAlignment := 3;// Alinhas celula (2 Esquerda, 3 Centro e 4 Direita)
    vPlanilha.cells[5,1] := 'CÓDIGO';
    vPlanilha.cells[5,1].interior.color := $00D9D9D9; // Alterar cor de fundo
    vPlanilha.cells[5,1].font.bold := true; // Colocar texto em netrigo
    vPlanilha.cells[5,1].horizontalAlignment := 3; // Alinhar celula (2 Esquerda, 3 Centro e 4 Direita)
    vPlanilha.cells[5,2] := 'DESCRIÇÃO';
    vPlanilha.cells[5,2].interior.color := $00D9D9D9; // Alterar cor de fundo
    vPlanilha.cells[5,2].font.bold := true; // Colocar texto em netrigo
    vPlanilha.cells[5,2].horizontalAlignment := 3; // Alinhar celula (2 Esquerda, 3 Centro e 4 Direita)

    vPlanilha.Range['C4','D4'].merge(EmptyParam); //Mesclar celulas
    vPlanilha.cells[4,3] := 'REALIZA';
    vPlanilha.cells[4,3].interior.Color := $00595959; // Alterar cor de fundo
    vPlanilha.cells[4,3].font.bold := true; // Colocar o texto em negrito
    vPlanilha.cells[4,3].font.color := $00FFFFFF; // Alterar a cor do texto
    vPlanilha.cells[4,3].horizontalAlignment := 3;// Alinhar celula (2 Esquerda, 3 Centro e 4 Direita)
    vPlanilha.cells[5,3] := '     SIM     ';
    vPlanilha.cells[5,3].interior.Color := $00D9D9D9; // Alterar cor de fundo
    vPlanilha.cells[5,3].font.bold := true; // Colocar o texto em negrito
    vPlanilha.cells[5,3].horizontalAlignment := 3;// Alinhar celula (2 Esquerda, 3 Centro e 4 Direita)
    vPlanilha.cells[5,4] := '     NÃO     ';
    vPlanilha.cells[5,4].interior.Color := $00D9D9D9; // Alterar cor de fundo
    vPlanilha.cells[5,4].font.bold := true; // Colocar o texto em negrito
    vPlanilha.cells[5,4].horizontalAlignment := 3;// Alinhar celula (2 Esquerda, 3 Centro e 4 Direita)

    try
      vLinha := 6;
      while not  qryProcedimento.Eof do
      begin
        vPlanilha.cells[vLinha,1] := qryProcedimentoCODIGO.AsString;
        vPlanilha.cells[vLinha,1].font.bold := true; // Colocar o texto em negrito
        vPlanilha.cells[vLinha,2] := qryProcedimentoDESCRICAO.AsString;
        vPlanilha.cells[vLinha,2].font.bold := true; // Colocar o texto em negrito
        vLinha:=vLinha+1;
        qryProcedimento.Next;
      end;
    finally end;

    vExcel.columns.autofit;

      vExcel.columns.autofit; // Ajustar colunas de acordo com os valores
      vExcel.visible := true;

      vExcel:=Unassigned;

  end else
    ShowMessage('Favor gerar os dados para realizar a exportação.');

    with qryExec do // INSERIR LOG DE EXPORTAÇÃO
    begin
    Close;
    SQL.Clear;
    SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, PARAMETRO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'E'+#39+', '+#39+parametro+#39+', SYSDATE)');
    ExecSQL;
    end;

end;

end.
