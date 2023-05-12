object DM: TDM
  OldCreateOrder = False
  Left = 64792
  Top = 235
  Height = 236
  Width = 504
  object OpenDialog1: TOpenDialog
    Left = 96
    Top = 16
  end
  object conTI: TADOConnection
    ConnectionString = 'Provider=OraOLEDB.Oracle.1;Persist Security Info=False;'
    LoginPrompt = False
    Provider = 'OraOLEDB.Oracle.1'
    Left = 24
    Top = 72
  end
  object SaveDialog1: TSaveDialog
    Left = 24
    Top = 16
  end
  object conLocaWeb1: TADOConnection
    ConnectionString = 'Provider=MSDASQL.1;Persist Security Info=False'
    LoginPrompt = False
    Provider = 'MSDASQL.1'
    Left = 96
    Top = 72
  end
  object qryRelat60: TADOQuery
    Connection = conTI
    Parameters = <>
    SQL.Strings = (
      
        '         SELECT GLOSA.NUM_DOC, GLOSA.SENHA_LIB, '#39'03/03/1992'#39' DAT' +
        'A, GLOSA.NOME, GLOSA.COD_AMB, GLOSA.COD_SERV, GLOSA.VALOR, GLOSA' +
        '.VALOR_GLOSADO, GLOSA.OBS_GLOSA, JUSGLOSA.COD_TISS_GLOSA, JUSGLO' +
        'SA.DESCRICAO, '#39'30805'#39' FATURA'
      'FROM ('
      '        SELECT    SUBSTR(A.SENHA_LIB, 1, 200) SENHA_LIB,'
      '                  A.DT_CONC DATA,'
      '                  A.COD_SEG,'
      '                  NVL(A.NOME, B.NOME) NOME,'
      
        '                  SUBSTR(NVL( TO_CHAR( A.JUS_GLOSA ), ACHAJUSGLO' +
        'SAVT(30805, A.TIPO_SERV, A.TIPO_ENTO, A.COD_SERV )), 1, 200) JUS' +
        '_GLOSA,'
      ''
      '                  CASE'
      
        '                    length(SUBSTR(NVL( TO_CHAR( A.JUS_GLOSA ), A' +
        'CHAJUSGLOSAVT(30805, A.TIPO_SERV, A.TIPO_ENTO, A.COD_SERV )), 1,' +
        ' 200))'
      
        '                  WHEN 4 THEN substr(SUBSTR(NVL( TO_CHAR( A.JUS_' +
        'GLOSA ), ACHAJUSGLOSAVT(30805, A.TIPO_SERV, A.TIPO_ENTO, A.COD_S' +
        'ERV )), 1, 200), 1,1)'
      
        '                  WHEN 5 THEN substr(SUBSTR(NVL( TO_CHAR( A.JUS_' +
        'GLOSA ), ACHAJUSGLOSAVT(30805, A.TIPO_SERV, A.TIPO_ENTO, A.COD_S' +
        'ERV )), 1, 200), 1,2)'
      
        '                  WHEN 6 THEN substr(SUBSTR(NVL( TO_CHAR( A.JUS_' +
        'GLOSA ), ACHAJUSGLOSAVT(30805, A.TIPO_SERV, A.TIPO_ENTO, A.COD_S' +
        'ERV )), 1, 200), 1,3)'
      
        '                  ELSE        substr(SUBSTR(NVL( TO_CHAR( A.JUS_' +
        'GLOSA ), ACHAJUSGLOSAVT(30805, A.TIPO_SERV, A.TIPO_ENTO, A.COD_S' +
        'ERV )), 1, 200), 1,4)'
      '                  END JUS_GLOSA2,'
      
        '                  CASE WHEN length(SUBSTR(NVL( TO_CHAR( A.JUS_GL' +
        'OSA ), ACHAJUSGLOSAVT(2510, A.TIPO_SERV, A.TIPO_ENTO, A.COD_SERV' +
        ' )), 1, 200)) > 3 THEN '#39'(*)'#39' ELSE '#39#39' END OBS_GLOSA,'
      ''
      '                  A.TIPO_ENT,'
      '                  A.TIPO_SERV,'
      '                  A.COD_SERV,'
      
        '                  SUBSTR(F_GET_COD_AMB_SERVICO(A.TIPO_ENT, A.TIP' +
        'O_SERV, A.COD_SERV), 1, 200) COD_AMB,'
      '                  SUM( A.VALOR + A.VALOR_INSSADM ) VALOR,'
      '                  SUM( A.VALOR_GLOSADO ) VALOR_GLOSADO,'
      '                  SUM(A.VALOR_GLOSA_RAT) VALOR_GLOSA_RAT,'
      
        '                  SUBSTR(ACHANUMGUIA(A.COD_SERV, A.TIPO_SERV, A.' +
        'TIPO_ENT), 1, 200) NUM_DOC'
      
        '        FROM      ( SELECT    ACHA_SENHA_LIB(B.TIPO_ENT, B.TIPO_' +
        'SERV, B.COD_SERV) SENHA_LIB, B.DT_CONC, B.COD_SEG,'
      '                              B.JUS_GLOSA, B.NOME, B.TIPO_ENT,'
      '                              A.COD_ENTD, A.TIPO_ENTD,'
      
        '                              A.TIPO_ENTO, B.TIPO_SERV, B.COD_SE' +
        'RV, B.COD_AMB, A.VALOR'
      
        '                              VALOR, A.VALOR_INSSADM, A.VALOR_GL' +
        'OSADO, C.VALOR_GLOSA_RAT'
      '                    FROM      IM_VTCN2 A, IM_VTCN1 B,'
      
        '                              ( SELECT    SUM(VALOR_GLOSA_RAT) V' +
        'ALOR_GLOSA_RAT, TIPO_SERVREL, COD_SERVREL, VTCN'
      '                                FROM      IM_DRAT'
      
        '                                GROUP BY  TIPO_SERVREL, COD_SERV' +
        'REL, VTCN) C'
      '                    WHERE     A.TIPO_ENTO = B.TIPO_ENT'
      '                    AND       A.TIPO_SERV = B.TIPO_SERV'
      '                    AND       A.COD_SERV = B.COD_SERV'
      '                    AND       A.TIPO_SERVREL = C.TIPO_SERVREL(+)'
      '                    AND       A.COD_SERVREL = C.COD_SERVREL(+)'
      '                    AND       C.VTCN(+) = 2'
      '                    AND       A.COD_FAT = 30805'
      '                    AND       B.VALOR_GLOSADO > 0'
      ''
      '                    UNION ALL'
      ''
      
        '                    SELECT    ACHA_SENHA_LIB(B.TIPO_ENT, B.TIPO_' +
        'SERV, B.COD_SERV) SENHA_LIB, B.DT_CONC, B.COD_SEG,'
      
        '                              NVL(A.JUS_GLOSA, B.JUS_GLOSA) JUS_' +
        'GLOSA, B.NOME, B.TIPO_ENT,'
      '                              A.COD_ENTD, A.TIPO_ENTD,'
      '                              A.TIPO_ENTO, 3 TIPO_SERV,'
      '                              A.COD_SERVREL,'
      '                              B.COD_AMB,'
      
        '                              A.VALOR  VALOR, A.VALOR_INSSADM, A' +
        '.VALOR_GLOSADO, C.VALOR_GLOSA_RAT'
      '                    FROM      IM_VTCN3 A, IM_VTCN1 B,'
      
        '                              ( SELECT   SUM(VALOR_GLOSA_RAT) VA' +
        'LOR_GLOSA_RAT, TIPO_SERVREL, COD_SERVREL, VTCN'
      '                                FROM     IM_DRAT'
      
        '                                GROUP BY TIPO_SERVREL, COD_SERVR' +
        'EL, VTCN ) C'
      '                    WHERE      B.TIPO_ENT = 2'
      '                    AND        B.TIPO_SERV = 2'
      '                    AND        A.COD1 = B.COD_SERV'
      
        '                    AND        A.TIPO_SERVRELREL = C.TIPO_SERVRE' +
        'L(+)'
      
        '                    AND        A.COD_SERVRELREL = C.COD_SERVREL(' +
        '+)'
      '                    AND        C.VTCN(+) = 3'
      '                    AND        A.COD_FAT = 30805'
      '                    AND        B.VALOR_GLOSADO > 0'
      '                  ) A , IM_SEG B'
      '        WHERE     A.COD_SEG = B.COD_SEG'
      
        '        GROUP BY  A.SENHA_LIB, A.DT_CONC, A.COD_SEG, NVL(A.NOME,' +
        ' B.NOME), A.JUS_GLOSA, A.TIPO_ENT, A.COD_ENTD, A.TIPO_ENTD, A.TI' +
        'PO_ENTO, A.TIPO_SERV, A.COD_SERV, A.COD_AMB'
      '     ) GLOSA,'
      '     ('
      
        '        SELECT CHAVE_UT, LOTE, SERVS.TIPO_SERV, SERVS.COD_SERV, ' +
        'SERVS.TIPO_ENT, NUM_GUIA, TIPO_SERVICO'
      '        FROM ('
      
        '                SELECT A.CHAVE_UT, V.CHAVE_PAI LOTE, 0 TIPO_SERV' +
        ', 0 TIPO_SERVICO, A.COD_CONS COD_SERV, 0 TIPO_ENT, A.NUM_GUIA'
      '                FROM IM_CONSM A, IM_UTSERV V'
      '                WHERE A.CHAVE_UT = V.CHAVE(+)'
      '                UNION ALL'
      
        '                SELECT A.CHAVE_UT, V.CHAVE_PAI LOTE, 1 TIPO_SERV' +
        ', 0 TIPO_SERVICO, A.COD_CONS COD_SERV, 2 TIPO_ENT, A.NUM_GUIA'
      '                FROM IM_CONSH A, IM_UTSERV V'
      '                WHERE A.CHAVE_UT = V.CHAVE(+)'
      '                UNION ALL'
      
        '                SELECT A.CHAVE_UT, V.CHAVE_PAI LOTE, 2 TIPO_SERV' +
        ', 1 TIPO_SERVICO, A.COD_EXM COD_SERV, 0 TIPO_ENT, A.NUM_GUIA'
      '                FROM IM_EXM A, IM_UTSERV V'
      '                WHERE A.CHAVE_UT = V.CHAVE(+)'
      '                UNION ALL'
      
        '                SELECT A.CHAVE_UT, V.CHAVE_PAI LOTE, 3 TIPO_SERV' +
        ', 1 TIPO_SERVICO, A.COD_EXH COD_SERV, 2 TIPO_ENT, A.NUM_GUIA'
      '                FROM IM_EXH A, IM_UTSERV V'
      '                WHERE A.CHAVE_UT = V.CHAVE(+)'
      '                UNION ALL'
      
        '                SELECT A.CHAVE_UT, V.CHAVE_PAI LOTE, 4 TIPO_SERV' +
        ', 1 TIPO_SERVICO, A.COD_EXL COD_SERV, 1 TIPO_ENT, A.NUM_GUIA'
      '                FROM IM_EXL A, IM_UTSERV V'
      '                WHERE A.CHAVE_UT = V.CHAVE(+)'
      '                UNION ALL'
      
        '                SELECT A.CHAVE_UT, V.CHAVE_PAI LOTE, 5 TIPO_SERV' +
        ', 3 TIPO_SERVICO, A.COD_CIRURM COD_SERV, 0 TIPO_ENT, A.NUM_GUIA'
      '                FROM IM_CIRM A, IM_UTSERV V'
      '                WHERE A.CHAVE_UT = V.CHAVE(+)'
      '                UNION ALL'
      
        '                SELECT A.CHAVE_UT, V.CHAVE_PAI LOTE, 6 TIPO_SERV' +
        ', 3 TIPO_SERVICO, A.COD_CIRUR COD_SERV, 2 TIPO_ENT, A.NUM_GUIA'
      '                FROM IM_CIRUR A, IM_UTSERV V'
      '                WHERE A.CHAVE_UT = V.CHAVE(+)'
      '                UNION ALL'
      
        '                SELECT A.CHAVE_UT, V.CHAVE_PAI LOTE, 7 TIPO_SERV' +
        ', 3 TIPO_SERVICO, A.COD_EXH COD_SERV, -1 TIPO_ENT, X.NUM_GUIA'
      '                FROM IM_CEXH A, IM_CIRUR X, IM_UTSERV V'
      
        '                WHERE A.CHAVE_UT = V.CHAVE(+) AND A.COD_CIRUR = ' +
        'X.COD_CIRUR'
      '                UNION ALL'
      
        '                SELECT A.CHAVE_UT, V.CHAVE_PAI LOTE, 8 TIPO_SERV' +
        ', 3 TIPO_SERVICO, A.COD_EXL COD_SERV, -1 TIPO_ENT, X.NUM_GUIA'
      '                FROM IM_CEXL A, IM_CIRUR X, IM_UTSERV V'
      
        '                WHERE A.CHAVE_UT = V.CHAVE(+) AND A.COD_CIRUR = ' +
        'X.COD_CIRUR'
      '             ) SERVS'
      '        WHERE 1 = 1'
      '     ) SRVS, JUSGLOSA'
      'WHERE GLOSA.COD_SERV = SRVS.COD_SERV'
      'AND GLOSA.TIPO_SERV = SRVS.TIPO_SERVICO'
      'AND GLOSA.TIPO_ENT = SRVS.TIPO_ENT'
      'AND JUSGLOSA.CODIGO = GLOSA.JUS_GLOSA2')
    Left = 176
    Top = 24
    object qryRelat60NUM_DOC: TWideStringField
      FieldName = 'NUM_DOC'
      ReadOnly = True
      Size = 200
    end
    object qryRelat60SENHA_LIB: TWideStringField
      FieldName = 'SENHA_LIB'
      ReadOnly = True
      Size = 200
    end
    object qryRelat60DATA: TWideStringField
      FieldName = 'DATA'
      ReadOnly = True
      FixedChar = True
      Size = 10
    end
    object qryRelat60NOME: TWideStringField
      FieldName = 'NOME'
      ReadOnly = True
      Size = 40
    end
    object qryRelat60COD_AMB: TWideStringField
      FieldName = 'COD_AMB'
      ReadOnly = True
      Size = 200
    end
    object qryRelat60COD_SERV: TBCDField
      FieldName = 'COD_SERV'
      ReadOnly = True
      Precision = 20
      Size = 0
    end
    object qryRelat60VALOR: TBCDField
      FieldName = 'VALOR'
      ReadOnly = True
      Precision = 32
    end
    object qryRelat60VALOR_GLOSADO: TBCDField
      FieldName = 'VALOR_GLOSADO'
      ReadOnly = True
      Precision = 32
    end
    object qryRelat60OBS_GLOSA: TWideStringField
      FieldName = 'OBS_GLOSA'
      ReadOnly = True
      Size = 3
    end
    object qryRelat60COD_TISS_GLOSA: TIntegerField
      FieldName = 'COD_TISS_GLOSA'
      ReadOnly = True
    end
    object qryRelat60DESCRICAO: TWideStringField
      FieldName = 'DESCRICAO'
      ReadOnly = True
      Size = 250
    end
    object qryRelat60FATURA: TWideStringField
      FieldName = 'FATURA'
      ReadOnly = True
      FixedChar = True
      Size = 5
    end
  end
  object qryRelat59: TADOQuery
    Connection = conTI
    AfterOpen = qryRelat59AfterOpen
    Parameters = <>
    SQL.Strings = (
      
        'SELECT /*+PARALLEL(6) */ CAPA.LOTE, CAPA.NU_PROTOCOLO, GLOSA.NUM' +
        '_DOC, GLOSA.SENHA_LIB, '#39'03/03/1992'#39' DATA, GLOSA.NOME, GLOSA.COD_' +
        'AMB, GLOSA.COD_SERV, GLOSA.VALOR, GLOSA.VALOR_GLOSADO, GLOSA.OBS' +
        '_GLOSA, JUSGLOSA.COD_TISS_GLOSA, JUSGLOSA.DESCRICAO,  '#39'30805'#39'  F' +
        'ATURA'
      'FROM ('
      '        SELECT    SUBSTR(A.SENHA_LIB, 1, 200) SENHA_LIB,'
      '                  A.DT_CONC DATA,'
      '                  A.COD_SEG,'
      '                  NVL(A.NOME, B.NOME) NOME,'
      
        '                  SUBSTR(NVL( TO_CHAR( A.JUS_GLOSA ), ACHAJUSGLO' +
        'SAVT(30699, A.TIPO_SERV, A.TIPO_ENTO, A.COD_SERV )), 1, 200) JUS' +
        '_GLOSA,'
      ''
      '                  CASE'
      
        '                    length(SUBSTR(NVL( TO_CHAR( A.JUS_GLOSA ), A' +
        'CHAJUSGLOSAVT(30699, A.TIPO_SERV, A.TIPO_ENTO, A.COD_SERV )), 1,' +
        ' 200))'
      
        '                  WHEN 4 THEN substr(SUBSTR(NVL( TO_CHAR( A.JUS_' +
        'GLOSA ), ACHAJUSGLOSAVT(30699, A.TIPO_SERV, A.TIPO_ENTO, A.COD_S' +
        'ERV )), 1, 200), 1,1)'
      
        '                  WHEN 5 THEN substr(SUBSTR(NVL( TO_CHAR( A.JUS_' +
        'GLOSA ), ACHAJUSGLOSAVT(30699, A.TIPO_SERV, A.TIPO_ENTO, A.COD_S' +
        'ERV )), 1, 200), 1,2)'
      
        '                  WHEN 6 THEN substr(SUBSTR(NVL( TO_CHAR( A.JUS_' +
        'GLOSA ), ACHAJUSGLOSAVT(30699, A.TIPO_SERV, A.TIPO_ENTO, A.COD_S' +
        'ERV )), 1, 200), 1,3)'
      
        '                  ELSE        substr(SUBSTR(NVL( TO_CHAR( A.JUS_' +
        'GLOSA ), ACHAJUSGLOSAVT(30699, A.TIPO_SERV, A.TIPO_ENTO, A.COD_S' +
        'ERV )), 1, 200), 1,4)'
      '                  END JUS_GLOSA2,'
      
        '                  CASE WHEN length(SUBSTR(NVL( TO_CHAR( A.JUS_GL' +
        'OSA ), ACHAJUSGLOSAVT(2510, A.TIPO_SERV, A.TIPO_ENTO, A.COD_SERV' +
        ' )), 1, 200)) > 3 THEN '#39'(*)'#39' ELSE '#39#39' END OBS_GLOSA,'
      ''
      '                  A.TIPO_ENT,'
      '                  A.TIPO_SERV,'
      '                  A.COD_SERV,'
      
        '                  SUBSTR(F_GET_COD_AMB_SERVICO(A.TIPO_ENT, A.TIP' +
        'O_SERV, A.COD_SERV), 1, 200) COD_AMB,'
      '                  SUM( A.VALOR + A.VALOR_INSSADM ) VALOR,'
      '                  SUM( A.VALOR_GLOSADO ) VALOR_GLOSADO,'
      '                  SUM(A.VALOR_GLOSA_RAT) VALOR_GLOSA_RAT,'
      
        '                  SUBSTR(ACHANUMGUIA(A.COD_SERV, A.TIPO_SERV, A.' +
        'TIPO_ENT), 1, 200) NUM_DOC'
      
        '        FROM      ( SELECT    ACHA_SENHA_LIB(B.TIPO_ENT, B.TIPO_' +
        'SERV, B.COD_SERV) SENHA_LIB, B.DT_CONC, B.COD_SEG,'
      '                              B.JUS_GLOSA, B.NOME, B.TIPO_ENT,'
      '                              A.COD_ENTD, A.TIPO_ENTD,'
      
        '                              A.TIPO_ENTO, B.TIPO_SERV, B.COD_SE' +
        'RV, B.COD_AMB, A.VALOR'
      
        '                              VALOR, A.VALOR_INSSADM, A.VALOR_GL' +
        'OSADO, C.VALOR_GLOSA_RAT'
      '                    FROM      IM_VTCN2 A, IM_VTCN1 B,'
      
        '                              ( SELECT    SUM(VALOR_GLOSA_RAT) V' +
        'ALOR_GLOSA_RAT, TIPO_SERVREL, COD_SERVREL, VTCN'
      '                                FROM      IM_DRAT'
      
        '                                GROUP BY  TIPO_SERVREL, COD_SERV' +
        'REL, VTCN) C'
      '                    WHERE     A.TIPO_ENTO = B.TIPO_ENT'
      '                    AND       A.TIPO_SERV = B.TIPO_SERV'
      '                    AND       A.COD_SERV = B.COD_SERV'
      '                    AND       A.TIPO_SERVREL = C.TIPO_SERVREL(+)'
      '                    AND       A.COD_SERVREL = C.COD_SERVREL(+)'
      '                    AND       C.VTCN(+) = 2'
      '                    AND       A.COD_FAT = 30699'
      '                    AND       B.VALOR_GLOSADO > 0'
      ''
      '                    UNION ALL'
      ''
      
        '                    SELECT    ACHA_SENHA_LIB(B.TIPO_ENT, B.TIPO_' +
        'SERV, B.COD_SERV) SENHA_LIB, B.DT_CONC, B.COD_SEG,'
      
        '                              NVL(A.JUS_GLOSA, B.JUS_GLOSA) JUS_' +
        'GLOSA, B.NOME, B.TIPO_ENT,'
      '                              A.COD_ENTD, A.TIPO_ENTD,'
      '                              A.TIPO_ENTO, 3 TIPO_SERV,'
      '                              A.COD_SERVREL,'
      '                              B.COD_AMB,'
      
        '                              A.VALOR  VALOR, A.VALOR_INSSADM, A' +
        '.VALOR_GLOSADO, C.VALOR_GLOSA_RAT'
      '                    FROM      IM_VTCN3 A, IM_VTCN1 B,'
      
        '                              ( SELECT   SUM(VALOR_GLOSA_RAT) VA' +
        'LOR_GLOSA_RAT, TIPO_SERVREL, COD_SERVREL, VTCN'
      '                                FROM     IM_DRAT'
      
        '                                GROUP BY TIPO_SERVREL, COD_SERVR' +
        'EL, VTCN ) C'
      '                    WHERE      B.TIPO_ENT = 2'
      '                    AND        B.TIPO_SERV = 2'
      '                    AND        A.COD1 = B.COD_SERV'
      
        '                    AND        A.TIPO_SERVRELREL = C.TIPO_SERVRE' +
        'L(+)'
      
        '                    AND        A.COD_SERVRELREL = C.COD_SERVREL(' +
        '+)'
      '                    AND        C.VTCN(+) = 3'
      '                    AND        A.COD_FAT = 30699'
      '                    AND        B.VALOR_GLOSADO > 0'
      '                  ) A , IM_SEG B'
      '        WHERE     A.COD_SEG = B.COD_SEG'
      
        '        GROUP BY  A.SENHA_LIB, A.DT_CONC, A.COD_SEG, NVL(A.NOME,' +
        ' B.NOME), A.JUS_GLOSA, A.TIPO_ENT, A.COD_ENTD, A.TIPO_ENTD, A.TI' +
        'PO_ENTO, A.TIPO_SERV, A.COD_SERV, A.COD_AMB'
      '     ) GLOSA,'
      '     ('
      
        '        SELECT CHAVE_UT, LOTE, SERVS.TIPO_SERV, SERVS.COD_SERV, ' +
        'SERVS.TIPO_ENT, NUM_GUIA, TIPO_SERVICO'
      '        FROM ('
      
        '                SELECT A.CHAVE_UT, V.CHAVE_PAI LOTE, 0 TIPO_SERV' +
        ', 0 TIPO_SERVICO, A.COD_CONS COD_SERV, 0 TIPO_ENT, A.NUM_GUIA'
      '                FROM IM_CONSM A, IM_UTSERV V'
      '                WHERE A.CHAVE_UT = V.CHAVE(+)'
      '                UNION ALL'
      
        '                SELECT A.CHAVE_UT, V.CHAVE_PAI LOTE, 1 TIPO_SERV' +
        ', 0 TIPO_SERVICO, A.COD_CONS COD_SERV, 2 TIPO_ENT, A.NUM_GUIA'
      '                FROM IM_CONSH A, IM_UTSERV V'
      '                WHERE A.CHAVE_UT = V.CHAVE(+)'
      '                UNION ALL'
      
        '                SELECT A.CHAVE_UT, V.CHAVE_PAI LOTE, 2 TIPO_SERV' +
        ', 1 TIPO_SERVICO, A.COD_EXM COD_SERV, 0 TIPO_ENT, A.NUM_GUIA'
      '                FROM IM_EXM A, IM_UTSERV V'
      '                WHERE A.CHAVE_UT = V.CHAVE(+)'
      '                UNION ALL'
      
        '                SELECT A.CHAVE_UT, V.CHAVE_PAI LOTE, 2 TIPO_SERV' +
        ', 2 TIPO_SERVICO, A.COD_CIRUR COD_SERV, 2 TIPO_ENT, A.NUM_GUIA'
      '                FROM IM_CIRUR A, IM_UTSERV V'
      
        '                WHERE A.CHAVE_UT = V.CHAVE(+) AND A.COD_SEG = V.' +
        'COD_SEG'
      '                UNION ALL'
      
        '                SELECT A.CHAVE_UT, V.CHAVE_PAI LOTE, 3 TIPO_SERV' +
        ', 1 TIPO_SERVICO, A.COD_EXH COD_SERV, 2 TIPO_ENT, A.NUM_GUIA'
      '                FROM IM_EXH A, IM_UTSERV V'
      '                WHERE A.CHAVE_UT = V.CHAVE(+)'
      '                UNION ALL'
      
        '                SELECT A.CHAVE_UT, V.CHAVE_PAI LOTE, 4 TIPO_SERV' +
        ', 1 TIPO_SERVICO, A.COD_EXL COD_SERV, 1 TIPO_ENT, A.NUM_GUIA'
      '                FROM IM_EXL A, IM_UTSERV V'
      '                WHERE A.CHAVE_UT = V.CHAVE(+)'
      '                UNION ALL'
      
        '                SELECT A.CHAVE_UT, V.CHAVE_PAI LOTE, 5 TIPO_SERV' +
        ', 3 TIPO_SERVICO, A.COD_CIRURM COD_SERV, 0 TIPO_ENT, A.NUM_GUIA'
      '                FROM IM_CIRM A, IM_UTSERV V'
      '                WHERE A.CHAVE_UT = V.CHAVE(+)'
      '                UNION ALL'
      
        '                SELECT A.CHAVE_UT, V.CHAVE_PAI LOTE, 6 TIPO_SERV' +
        ', 3 TIPO_SERVICO, A.COD_CIRUR COD_SERV, 2 TIPO_ENT, A.NUM_GUIA'
      '                FROM IM_CIRUR A, IM_UTSERV V'
      '                WHERE A.CHAVE_UT = V.CHAVE(+)'
      '                UNION ALL'
      
        '                SELECT A.CHAVE_UT, V.CHAVE_PAI LOTE, 7 TIPO_SERV' +
        ', 3 TIPO_SERVICO, A.COD_EXH COD_SERV, -1 TIPO_ENT, X.NUM_GUIA'
      '                FROM IM_CEXH A, IM_CIRUR X, IM_UTSERV V'
      
        '                WHERE A.CHAVE_UT = V.CHAVE(+) AND A.COD_CIRUR = ' +
        'X.COD_CIRUR'
      '                UNION ALL'
      
        '                SELECT A.CHAVE_UT, V.CHAVE_PAI LOTE, 8 TIPO_SERV' +
        ', 3 TIPO_SERVICO, A.COD_EXL COD_SERV, -1 TIPO_ENT, X.NUM_GUIA'
      '                FROM IM_CEXL A, IM_CIRUR X, IM_UTSERV V'
      
        '                WHERE A.CHAVE_UT = V.CHAVE(+) AND A.COD_CIRUR = ' +
        'X.COD_CIRUR'
      '             ) SERVS'
      '        WHERE 1 = 1'
      '     ) SRVS, IM_UTCAPA CAPA, JUSGLOSA'
      'WHERE GLOSA.COD_SERV = SRVS.COD_SERV'
      'AND GLOSA.TIPO_SERV = SRVS.TIPO_SERVICO'
      'AND GLOSA.TIPO_ENT = SRVS.TIPO_ENT'
      'AND SRVS.LOTE = CAPA.CHAVE'
      'AND JUSGLOSA.CODIGO = GLOSA.JUS_GLOSA2')
    Left = 176
    Top = 72
    object qryRelat59LOTE: TWideStringField
      FieldName = 'LOTE'
      ReadOnly = True
      Size = 15
    end
    object qryRelat59NU_PROTOCOLO: TBCDField
      FieldName = 'NU_PROTOCOLO'
      ReadOnly = True
      Precision = 20
      Size = 0
    end
    object qryRelat59NUM_DOC: TWideStringField
      FieldName = 'NUM_DOC'
      ReadOnly = True
      Size = 200
    end
    object qryRelat59SENHA_LIB: TWideStringField
      FieldName = 'SENHA_LIB'
      ReadOnly = True
      Size = 200
    end
    object qryRelat59DATA: TWideStringField
      FieldName = 'DATA'
      ReadOnly = True
      FixedChar = True
      Size = 10
    end
    object qryRelat59NOME: TWideStringField
      FieldName = 'NOME'
      ReadOnly = True
      Size = 40
    end
    object qryRelat59COD_AMB: TWideStringField
      FieldName = 'COD_AMB'
      ReadOnly = True
      Size = 200
    end
    object qryRelat59COD_SERV: TBCDField
      FieldName = 'COD_SERV'
      ReadOnly = True
      Precision = 20
      Size = 0
    end
    object qryRelat59VALOR: TBCDField
      FieldName = 'VALOR'
      ReadOnly = True
      Precision = 32
    end
    object qryRelat59VALOR_GLOSADO: TBCDField
      FieldName = 'VALOR_GLOSADO'
      ReadOnly = True
      Precision = 32
    end
    object qryRelat59OBS_GLOSA: TWideStringField
      FieldName = 'OBS_GLOSA'
      ReadOnly = True
      Size = 3
    end
    object qryRelat59COD_TISS_GLOSA: TIntegerField
      FieldName = 'COD_TISS_GLOSA'
      ReadOnly = True
    end
    object qryRelat59DESCRICAO: TWideStringField
      FieldName = 'DESCRICAO'
      ReadOnly = True
      Size = 250
    end
    object qryRelat59FATURA: TWideStringField
      FieldName = 'FATURA'
      ReadOnly = True
      FixedChar = True
      Size = 5
    end
  end
  object qryRelat65: TADOQuery
    Connection = conTI
    AfterOpen = qryRelat65AfterOpen
    Parameters = <>
    SQL.Strings = (
      
        'SELECT ANO_MES, TRUNC(AVG(VIDAS),2) VIDAS, NVL(SUM(RECEITA),0) R' +
        'ECEITA, NVL(SUM(SINISTRALIDADE),0) SINISTRALIDADE, NVL(SUM(COMIS' +
        'SAO),0) COMISSAO'
      'FROM ( '
      
        'SELECT R3.ANO_MES, R3.VIDAS, R1.SINISTRALIDADE, R4.COMISSAO, R2.' +
        'RECEITA'
      '  FROM('
      '       SELECT ANO_MES, SUM(SINISTRALIDADE) SINISTRALIDADE'
      '         FROM('
      
        '              SELECT TO_CHAR(DT_LIB,'#39'YYYY/MM'#39') ANO_MES, SUM(VALO' +
        'R_SERV) SINISTRALIDADE'
      '                FROM IM_CONSM X'
      
        '               WHERE DT_LIB >= '#39'01/01/2022'#39' AND DT_LIB <= '#39'01/02' +
        '/2022'#39
      
        '                 AND (CONCILIADO = '#39'T'#39' OR (SITUACAO = 0 AND DT_L' +
        'IB >= SYSDATE-180))'
      '                 AND COD_SET = 1313'
      '                 AND REC_GLOSA IS NULL'
      '               GROUP BY TO_CHAR(DT_LIB,'#39'YYYY/MM'#39')'
      '               UNION ALL'
      
        '              SELECT TO_CHAR(DT_LIB,'#39'YYYY/MM'#39') ANO_MES, SUM(VALO' +
        'R_SERV) SINISTRALIDADE'
      '                FROM IM_CONSH'
      
        '               WHERE DT_LIB >= '#39'01/01/2022'#39' AND DT_LIB <= '#39'01/02' +
        '/2022'#39
      
        '                 AND (CONCILIADO = '#39'T'#39' OR (SITUACAO = 0 AND DT_L' +
        'IB >= SYSDATE-180))'
      '                 AND COD_SET = 1313'
      '                 AND REC_GLOSA IS NULL    '
      '               GROUP BY TO_CHAR(DT_LIB,'#39'YYYY/MM'#39')'
      '               UNION ALL'
      
        '              SELECT TO_CHAR(DT_LIB,'#39'YYYY/MM'#39') ANO_MES, SUM(VALO' +
        'R_SERV) SINISTRALIDADE'
      '                FROM IM_EXM'
      
        '               WHERE DT_LIB >= '#39'01/01/2022'#39' AND DT_LIB <= '#39'01/02' +
        '/2022'#39
      
        '                 AND (CONCILIADO = '#39'T'#39' OR (SITUACAO = 0 AND DT_L' +
        'IB >= SYSDATE-180))'
      '                 AND COD_SET = 1313'
      '                 AND REC_GLOSA IS NULL   '
      '               GROUP BY TO_CHAR(DT_LIB,'#39'YYYY/MM'#39')'
      '               UNION ALL'
      
        '              SELECT TO_CHAR(DT_LIB,'#39'YYYY/MM'#39') ANO_MES, SUM(VALO' +
        'R_SERV) SINISTRALIDADE'
      '                FROM IM_EXH'
      
        '               WHERE DT_LIB >= '#39'01/01/2022'#39' AND DT_LIB <= '#39'01/02' +
        '/2022'#39
      
        '                 AND (CONCILIADO = '#39'T'#39' OR (SITUACAO = 0 AND DT_L' +
        'IB >= SYSDATE-180))'
      '                 AND COD_SET = 1313'
      '                 AND REC_GLOSA IS NULL    '
      '               GROUP BY TO_CHAR(DT_LIB,'#39'YYYY/MM'#39')'
      '               UNION ALL'
      
        '              SELECT TO_CHAR(DT_LIB,'#39'YYYY/MM'#39') ANO_MES, SUM(VALO' +
        'R_SERV) SINISTRALIDADE'
      '                FROM IM_EXL'
      
        '               WHERE DT_LIB >= '#39'01/01/2022'#39' AND DT_LIB <= '#39'01/02' +
        '/2022'#39
      
        '                 AND (CONCILIADO = '#39'T'#39' OR (SITUACAO = 0 AND DT_L' +
        'IB >= SYSDATE-180))'
      '                 AND COD_SET = 1313'
      '                 AND REC_GLOSA IS NULL    '
      '               GROUP BY TO_CHAR(DT_LIB,'#39'YYYY/MM'#39')'
      '               UNION ALL'
      
        '              SELECT TO_CHAR(DT_LIB,'#39'YYYY/MM'#39') ANO_MES, SUM(VALO' +
        'R_SERV) SINISTRALIDADE'
      '                FROM IM_CIRUR'
      
        '               WHERE DT_LIB >= '#39'01/01/2022'#39' AND DT_LIB <= '#39'01/02' +
        '/2022'#39
      
        '                 AND (CONCILIADO = '#39'T'#39' OR (SITUACAO = 0 AND DT_L' +
        'IB >= SYSDATE-180))'
      '                 AND COD_SET = 1313'
      '                 AND REC_GLOSA IS NULL    '
      '               GROUP BY TO_CHAR(DT_LIB,'#39'YYYY/MM'#39')'
      '            )'
      '       GROUP BY ANO_MES'
      '      ) R1,'
      '      ('
      
        '       SELECT TO_CHAR(DT_VENC,'#39'YYYY/MM'#39') ANO_MES, NVL(SUM(D.VALO' +
        'R),0) - NVL(SUM(DDV.VALOR),0) RECEITA'
      '         FROM IM_BOLAB B'
      '         JOIN IM_DBOL  D         ON B.COD_BOL = D.COD_BOL'
      '         LEFT JOIN IM_DBODEV DDV ON D.COD_DBOL = DDV.COD_DBOL'
      '        WHERE B.COD_BOL = D.COD_BOL'
      '          AND B.COD_SET = 1313'
      
        '          AND B.DT_VENC >= '#39'01/01/2022'#39' AND B.DT_VENC <= '#39'01/02/' +
        '2022'#39
      '          AND B.CONCIL  = '#39'T'#39
      '          AND D.TIPO    = 0'
      '        GROUP BY TO_CHAR(DT_VENC,'#39'YYYY/MM'#39')        '
      '      ) R2,'
      '      ('
      '       SELECT ANO_MES, COUNT(*) VIDAS'
      '         FROM('
      '              SELECT C.ANO_MES, S.COD_SEG, S.COD_SET'
      
        '                FROM IM_SEG S, (SELECT DISTINCT C.MES_U DIA, C.A' +
        'NO_MES'
      '                                           FROM TI.CALENDARIO C'
      
        '                                          WHERE C.DATA >= '#39'01/01' +
        '/2022'#39' AND C.DATA <= '#39'01/02/2022'#39') C'
      '               WHERE S.DT_INCL <= DIA'
      '                 AND S.CANCELED = '#39'F'#39
      '                 AND S.COD_SET  = 1313'
      '               UNION ALL'
      '              SELECT C.ANO_MES, S.COD_SEG, S.COD_SET'
      
        '                FROM IM_SEG S, (SELECT DISTINCT C.MES_U DIA, C.A' +
        'NO_MES'
      '                                           FROM TI.CALENDARIO C'
      
        '                                          WHERE C.DATA >= '#39'01/01' +
        '/2022'#39' AND C.DATA <= '#39'01/02/2022'#39') C'
      '               WHERE S.DT_INCL <= C.DIA '
      '                 AND S.DT_CANC > C.DIA'
      '                 AND S.COD_SET = 1313'
      '             )'
      '        GROUP BY ANO_MES'
      '      ) R3,'
      '      ('
      
        '       SELECT NVL(EXT.ANO_MES,PAG.ANO_MES) ANO_MES, SUM(NVL(PAG.' +
        'PROVENTO,0) - NVL(EXT.ESTORNO,0)) COMISSAO'
      '         FROM('
      
        '              SELECT SUBSTR(A.MESANO,1,7) ANO_MES, (SUM(VALOR)-S' +
        'UM(VALOR_PG)) ESTORNO'
      '                FROM IM_CVEND A, IM_SEG B'
      '               WHERE A.PAG_EXT <> '#39'P'#39
      '                 AND A.COD_SEG = B.COD_SEG'
      '                 AND B.COD_SET = 1313'
      
        '                 AND A.MESANO >= SUBSTR('#39'01/01/2022'#39',7,4)||'#39'/'#39'||' +
        'SUBSTR('#39'01/01/2022'#39',4,2)||'#39'-1'#39' AND A.MESANO <= SUBSTR('#39'01/02/202' +
        '2'#39',7,4)||'#39'/'#39'||SUBSTR('#39'01/02/2022'#39',4,2)||'#39'-2'#39'           '
      '               GROUP BY SUBSTR(A.MESANO,1,7)'
      '             ) EXT,'
      '             ('
      
        '              SELECT SUBSTR(A.MESANO,1,7) ANO_MES, (SUM(VALOR)-S' +
        'UM(VALOR_PG)) PROVENTO'
      '                FROM IM_CVEND A, IM_SEG B'
      '               WHERE A.PAG_EXT = '#39'P'#39
      '                 AND A.COD_SEG = B.COD_SEG'
      '                 AND B.COD_SET = 1313'
      
        '                 AND A.MESANO >= SUBSTR('#39'01/01/2022'#39',7,4)||'#39'/'#39'||' +
        'SUBSTR('#39'01/01/2022'#39',4,2)||'#39'-1'#39' AND A.MESANO <= SUBSTR('#39'01/02/202' +
        '2'#39',7,4)||'#39'/'#39'||SUBSTR('#39'01/02/2022'#39',4,2)||'#39'-2'#39
      '               GROUP BY SUBSTR(A.MESANO,1,7)'
      '             ) PAG'
      '        WHERE PAG.ANO_MES = EXT.ANO_MES (+)'
      '        GROUP BY NVL(EXT.ANO_MES,PAG.ANO_MES)'
      '      ) R4      '
      ' WHERE R3.ANO_MES = R1.ANO_MES (+)'
      '   AND R3.ANO_MES = R2.ANO_MES (+)'
      '   AND R3.ANO_MES = R4.ANO_MES (+)'
      ')   '
      'GROUP BY ANO_MES')
    Left = 176
    Top = 120
    object qryRelat65ANO_MES: TWideStringField
      FieldName = 'ANO_MES'
      ReadOnly = True
      Size = 7
    end
    object qryRelat65VIDAS: TBCDField
      FieldName = 'VIDAS'
      ReadOnly = True
      Precision = 32
    end
    object qryRelat65RECEITA: TBCDField
      FieldName = 'RECEITA'
      ReadOnly = True
      Precision = 32
    end
    object qryRelat65SINISTRALIDADE: TBCDField
      FieldName = 'SINISTRALIDADE'
      ReadOnly = True
      Precision = 32
    end
    object qryRelat65COMISSAO: TBCDField
      FieldName = 'COMISSAO'
      ReadOnly = True
      Precision = 32
    end
  end
  object qryRelat100Tit: TADOQuery
    Connection = conTI
    CursorType = ctStatic
    AfterOpen = qryRelat59AfterOpen
    Parameters = <>
    SQL.Strings = (
      
        'SELECT T.COD_SEG || '#39' - '#39' || T.NOME TITULAR, TO_CHAR(B.COD_BOL) ' +
        'BOLETO, T.COD_SEG COD_TITULAR, S.COD_SEG||'#39'%'#39' COD_SEGX'
      '  FROM IM_BOLAB B'
      '  JOIN IM_DBOL  D ON B.COD_BOL = D.COD_BOL'
      '  JOIN IM_SEG   S ON D.COD_SEG = S.COD_SEG'
      '  JOIN IM_SEG   T ON S.COD_FAM = T.COD_SEG'
      ' WHERE B.COD_BOL = 59765'
      '   AND T.TIPO = '#39'E'#39
      
        ' GROUP BY T.COD_SEG || '#39' - '#39' || T.NOME, TO_CHAR(B.COD_BOL) , T.C' +
        'OD_SEG, S.COD_SEG||'#39'%'#39)
    Left = 240
    Top = 24
    object qryRelat100TitTITULAR: TWideStringField
      FieldName = 'TITULAR'
      ReadOnly = True
      Size = 58
    end
    object qryRelat100TitCOD_TITULAR: TWideStringField
      FieldName = 'COD_TITULAR'
      Size = 15
    end
    object qryRelat100TitBOLETO: TWideStringField
      FieldName = 'BOLETO'
      ReadOnly = True
      Size = 40
    end
    object qryRelat100TitCOD_SEGX: TWideStringField
      FieldName = 'COD_SEGX'
      ReadOnly = True
      Size = 16
    end
  end
  object qryRelat100Detalhe: TADOQuery
    Connection = conTI
    CursorType = ctStatic
    AfterOpen = qryRelat59AfterOpen
    DataSource = dsRelat100
    Parameters = <
      item
        Name = 'COD_TITULAR'
        DataType = ftWideString
        Size = 15
        Value = '134682-2'
      end
      item
        Name = 'BOLETO'
        DataType = ftWideString
        Size = 40
        Value = '59765'
      end
      item
        Name = 'COD_SEGX'
        DataType = ftWideString
        Size = 16
        Value = '134682-2%'
      end>
    SQL.Strings = (
      ''
      
        'SELECT S.COD_FAM COD_TITULAR, B.COD_BOL, DECODE(S.TIPO,'#39'E'#39','#39'TIT.' +
        #39','#39'D'#39','#39'DEP.'#39') TIPO, S.COD_SEG, S.NOME, V.DT_CONC REALIZACAO'
      '      ,CASE'
      '        WHEN D.TIPO = 3'
      
        '        THEN SUBSTR(D.CODIGO,INSTR(D.CODIGO,'#39';'#39',1,2)+1,INSTR(D.C' +
        'ODIGO,'#39';'#39',1,3)-5) || '#39' - '#39' ||'
      '             CASE '
      
        '              WHEN SUBSTR(D.CODIGO,1,INSTR(D.CODIGO,'#39';'#39',1)-1) = ' +
        '0 AND SUBSTR(D.CODIGO,INSTR(D.CODIGO,'#39';'#39',1)+1,INSTR(D.CODIGO,'#39';'#39 +
        ',1,2)-3) = 0'
      '              THEN '#39'Cons. Med.'#39
      
        '              WHEN SUBSTR(D.CODIGO,1,INSTR(D.CODIGO,'#39';'#39',1)-1) = ' +
        '2 AND SUBSTR(D.CODIGO,INSTR(D.CODIGO,'#39';'#39',1)+1,INSTR(D.CODIGO,'#39';'#39 +
        ',1,2)-3) = 0'
      '              THEN '#39'Cons. Hosp.'#39
      ''
      
        '              WHEN SUBSTR(D.CODIGO,1,INSTR(D.CODIGO,'#39';'#39',1)-1) = ' +
        '0 AND SUBSTR(D.CODIGO,INSTR(D.CODIGO,'#39';'#39',1)+1,INSTR(D.CODIGO,'#39';'#39 +
        ',1,2)-3) = 1'
      '              THEN '#39'Ex. Med.'#39
      
        '              WHEN SUBSTR(D.CODIGO,1,INSTR(D.CODIGO,'#39';'#39',1)-1) = ' +
        '1 AND SUBSTR(D.CODIGO,INSTR(D.CODIGO,'#39';'#39',1)+1,INSTR(D.CODIGO,'#39';'#39 +
        ',1,2)-3) = 1'
      '              THEN '#39'Ex. Lab.'#39
      
        '              WHEN SUBSTR(D.CODIGO,1,INSTR(D.CODIGO,'#39';'#39',1)-1) = ' +
        '2 AND SUBSTR(D.CODIGO,INSTR(D.CODIGO,'#39';'#39',1)+1,INSTR(D.CODIGO,'#39';'#39 +
        ',1,2)-3) = 1'
      '              THEN '#39'Ex. Hosp.'#39
      ''
      
        '              WHEN SUBSTR(D.CODIGO,1,INSTR(D.CODIGO,'#39';'#39',1)-1) = ' +
        '0 AND SUBSTR(D.CODIGO,INSTR(D.CODIGO,'#39';'#39',1)+1,INSTR(D.CODIGO,'#39';'#39 +
        ',1,2)-3) = 2'
      '              THEN '#39'Cir/Inter. Med.'#39
      
        '              WHEN SUBSTR(D.CODIGO,1,INSTR(D.CODIGO,'#39';'#39',1)-1) = ' +
        '2 AND SUBSTR(D.CODIGO,INSTR(D.CODIGO,'#39';'#39',1)+1,INSTR(D.CODIGO,'#39';'#39 +
        ',1,2)-3) = 2'
      '              THEN '#39'Cir/Inter. Hosp.'#39
      '              END || '#39' ( '#39' || Z.COD_SERV || '#39' ) '#39' || Z.DESCRICAO'
      '        ELSE DESCDBOLCOM(D.TIPO,D.CODIGO,'#39'F'#39',D.COD_BOL)'
      '       END AS DESCRICAO'
      '      ,D.VALOR'
      '      ,P.NOME PRESTADOR'
      
        '      ,SUBSTR(formata_valor_generico( ROUND(VALOR_COPART_VTCN1(V' +
        '.FMOD_VLR,                            '
      '                                   V.FATOR_SEG,'
      '                                   V.VALOR_1,'
      '                                   V.VALOR_INSSADM_1,'
      '                                   V.VALOR_GLOSADO,'
      '                                   V.USA_PRC_DIF_HON,'
      '                                   V.FATOR_SEG_HON,'
      '                                   V.VALOR1_HON,'
      '                                   V.VALOR_GLOSA_HON,'
      '                                   V.FATOR_SEG_OUTROS,'
      '                                   V.VALOR1_OUTROS,'
      
        '                                   V.VALOR_GLOSA_OUTROS),2),'#39','#39')' +
        ',1,20) VALOR_C'
      '  FROM IM_BOLAB B'
      '  JOIN IM_DBOL  D ON B.COD_BOL = D.COD_BOL'
      '  JOIN IM_SEG   S ON D.COD_SEG = S.COD_SEG'
      '  JOIN IM_VTCN1 V '
      
        '    ON SUBSTR(D.CODIGO,INSTR(D.CODIGO,'#39';'#39',1,2)+1,INSTR(SUBSTR(D.' +
        'CODIGO,INSTR(D.CODIGO,'#39';'#39',1,2)+1),'#39';'#39')-1) = V.COD_SERV'
      '   AND SUBSTR(D.CODIGO,1,INSTR(D.CODIGO,'#39';'#39',1)-1) = V.TIPO_ENT'
      
        '   AND SUBSTR(D.CODIGO,INSTR(D.CODIGO,'#39';'#39',1)+1,INSTR(D.CODIGO,'#39';' +
        #39',1,2)-3) = V.TIPO_SERV  '
      
        '  LEFT JOIN (SELECT * FROM IM_TSERV WHERE COD_TAB = 16) Z ON Z.C' +
        'OD_SERV = SUBSTR(D.CODIGO,INSTR(D.CODIGO,'#39';'#39',1,3)+1)'
      
        '  LEFT JOIN TI.VW_PRESTADORES P ON V.TIPO_ENT = P.TIPO_ENT AND V' +
        '.COD_ENT= P.COD_ENT'
      ' WHERE S.COD_FAM =  :COD_TITULAR'
      '      AND B.COD_BOL = :BOLETO'
      '      AND S.COD_SEG LIKE :COD_SEGX'
      
        ' ORDER BY S.COD_FAM, DECODE(S.TIPO,'#39'E'#39','#39'TIT.'#39','#39'D'#39','#39'DEP.'#39') DESC, ' +
        'D.TIPO'
      ' ')
    Left = 418
    Top = 24
    object qryRelat100DetalheCOD_FAM: TWideStringField
      FieldName = 'COD_TITULAR'
      ReadOnly = True
      Size = 15
    end
    object qryRelat100DetalheCOD_BOL: TBCDField
      FieldName = 'COD_BOL'
      ReadOnly = True
      Precision = 20
      Size = 0
    end
    object qryRelat100DetalheTIPO: TWideStringField
      FieldName = 'TIPO'
      ReadOnly = True
      Size = 4
    end
    object qryRelat100DetalheCOD_SEG: TWideStringField
      FieldName = 'COD_SEG'
      ReadOnly = True
      Size = 15
    end
    object qryRelat100DetalheNOME: TWideStringField
      FieldName = 'NOME'
      ReadOnly = True
      Size = 40
    end
    object qryRelat100DetalheDESCRICAO: TWideStringField
      FieldName = 'DESCRICAO'
      ReadOnly = True
      Size = 4000
    end
    object qryRelat100DetalheVALOR: TBCDField
      FieldName = 'VALOR'
      ReadOnly = True
      Precision = 20
      Size = 2
    end
    object qryRelat100DetalheREALIZACAO: TDateTimeField
      FieldName = 'REALIZACAO'
      ReadOnly = True
    end
    object qryRelat100DetalhePRESTADOR: TWideStringField
      FieldName = 'PRESTADOR'
      ReadOnly = True
      Size = 70
    end
    object qryRelat100DetalheVALOR_C: TWideStringField
      FieldName = 'VALOR_C'
      ReadOnly = True
    end
  end
  object dsRelat100: TDataSource
    DataSet = qryRelat100Tit
    Left = 329
    Top = 24
  end
  object qryRelat116Detalhe: TADOQuery
    Connection = conTI
    CursorType = ctStatic
    AfterOpen = qryRelat59AfterOpen
    DataSource = dsRelat116
    Parameters = <
      item
        Name = 'COD_TITULAR'
        DataType = ftWideString
        Size = 15
        Value = '134682-2'
      end
      item
        Name = 'BOLETO'
        DataType = ftWideString
        Size = 40
        Value = '59765'
      end>
    SQL.Strings = (
      
        'SELECT S.COD_FAM COD_TITULAR, B.COD_BOL, DECODE(S.TIPO,'#39'E'#39','#39'TIT.' +
        #39','#39'D'#39','#39'DEP.'#39') TIPO, S.COD_SEG, S.NOME, SUM(D.VALOR) VALOR'
      '  FROM IM_BOLAB B'
      '  JOIN IM_DBOL  D ON B.COD_BOL = D.COD_BOL'
      '  JOIN IM_SEG   S ON D.COD_SEG = S.COD_SEG'
      '  JOIN IM_VTCN1 V '
      
        '    ON SUBSTR(D.CODIGO,INSTR(D.CODIGO,'#39';'#39',1,2)+1,INSTR(SUBSTR(D.' +
        'CODIGO,INSTR(D.CODIGO,'#39';'#39',1,2)+1),'#39';'#39')-1)= V.COD_SERV'
      '   AND SUBSTR(D.CODIGO,1,INSTR(D.CODIGO,'#39';'#39',1)-1) = V.TIPO_ENT'
      
        '   AND SUBSTR(D.CODIGO,INSTR(D.CODIGO,'#39';'#39',1)+1,INSTR(D.CODIGO,'#39';' +
        #39',1,2)-3) = V.TIPO_SERV  '
      
        '  LEFT JOIN (SELECT * FROM IM_TSERV WHERE COD_TAB = 16) Z ON Z.C' +
        'OD_SERV = SUBSTR(D.CODIGO,INSTR(D.CODIGO,'#39';'#39',1,3)+1)'
      ' WHERE S.COD_FAM =  :COD_TITULAR'
      '   AND B.COD_BOL = :BOLETO'
      
        ' GROUP BY S.COD_FAM, B.COD_BOL, DECODE(S.TIPO,'#39'E'#39','#39'TIT.'#39','#39'D'#39','#39'DE' +
        'P.'#39'), S.COD_SEG, S.NOME'
      ' ORDER BY S.COD_FAM, DECODE(S.TIPO,'#39'E'#39','#39'TIT.'#39','#39'D'#39','#39'DEP.'#39') DESC')
    Left = 418
    Top = 86
    object WideStringField4: TWideStringField
      FieldName = 'COD_TITULAR'
      ReadOnly = True
      Size = 15
    end
    object BCDField1: TBCDField
      FieldName = 'COD_BOL'
      ReadOnly = True
      Precision = 20
      Size = 0
    end
    object WideStringField5: TWideStringField
      FieldName = 'TIPO'
      ReadOnly = True
      Size = 4
    end
    object WideStringField6: TWideStringField
      FieldName = 'COD_SEG'
      ReadOnly = True
      Size = 15
    end
    object WideStringField7: TWideStringField
      FieldName = 'NOME'
      ReadOnly = True
      Size = 40
    end
    object BCDField2: TBCDField
      FieldName = 'VALOR'
      ReadOnly = True
      Precision = 20
      Size = 2
    end
  end
  object dsRelat116: TDataSource
    DataSet = qryRelat116Tit
    Left = 329
    Top = 88
  end
  object qryRelat116Tit: TADOQuery
    Connection = conTI
    CursorType = ctStatic
    AfterOpen = qryRelat59AfterOpen
    Parameters = <>
    SQL.Strings = (
      
        'SELECT T.COD_SEG || '#39' - '#39' || T.NOME TITULAR, TO_CHAR(B.COD_BOL) ' +
        'BOLETO, T.COD_SEG COD_TITULAR'
      '  FROM IM_BOLAB B'
      '  JOIN IM_DBOL  D ON B.COD_BOL = D.COD_BOL'
      '  JOIN IM_SEG   S ON D.COD_SEG = S.COD_SEG'
      '  JOIN IM_SEG   T ON S.COD_FAM = T.COD_SEG'
      ' WHERE B.COD_BOL = &CODIGO'
      '   AND T.TIPO = '#39'E'#39
      
        ' GROUP BY T.COD_SEG || '#39' - '#39' || T.NOME, TO_CHAR(B.COD_BOL) , T.C' +
        'OD_SEG')
    Left = 240
    Top = 88
    object WideStringField1: TWideStringField
      FieldName = 'TITULAR'
      ReadOnly = True
      Size = 58
    end
    object WideStringField2: TWideStringField
      FieldName = 'COD_TITULAR'
      Size = 15
    end
    object WideStringField3: TWideStringField
      FieldName = 'BOLETO'
      ReadOnly = True
      Size = 40
    end
  end
end
