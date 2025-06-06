SELECT 
     SB.CODCOLIGADA                                     AS [COLIGADA]
    ,SB.CHAPA                                           AS [CHAPA]
    ,F.NOME                                             AS [COLABORADOR]
    ,S.DESCRICAO                                        AS [PROJETO]
    ,F.CODSITUACAO                                      AS [SITUACAO]
    ,F.DATAADMISSAO                                     AS [DATA_ADMISSAO]
    ,F.DATADEMISSAO                                     AS [DATA_RESCISAO]
    ,FN.NOME                                            AS [CARGO]
    ,CAST(FLOOR((
        (SB.EXTRAATU + SB.EXTRAANT) - 
        (SB.FALTAANT + SB.ATRASOANT + SB.FALTAATU + SB.ATRASOATU)
    ) / 60) AS VARCHAR)
    + ':' +
    RIGHT('0' + CAST(ABS((
        (SB.EXTRAATU + SB.EXTRAANT) - 
        (SB.FALTAANT + SB.ATRASOANT + SB.FALTAATU + SB.ATRASOATU)
    ) % 60) AS VARCHAR), 2)                             AS [SALDO_BH]

FROM BANCO_HORAS SB

    JOIN FUNCIONARIOS F
        ON F.CODCOLIGADA = SB.CODCOLIGADA 
        AND F.CHAPA = SB.CHAPA

    JOIN SETORES S 
        ON S.CODCOLIGADA = F.CODCOLIGADA 
        AND S.CODIGO = F.CODSECAO

    LEFT JOIN CARGOS FN 
        ON FN.CODIGO = F.CODFUNCAO 
        AND FN.CODCOLIGADA = F.CODCOLIGADA

WHERE 
    SB.CODCOLIGADA = :CODCOLIGADA
    AND SB.INICIOPER = :DATA_INICIO
    --AND SB.CHAPA LIKE 'XXXXXXX' -- (opcional para filtrar colaborador específico)

ORDER BY 
    F.NOME