SELECT
to_char(M.DATAHORAGERACAO, 'dd/mm/yyyy') AS "Data Alteração",
c.contrato as Contrato,
upper(P.NOMECOMPLETO) AS Nome,
case VARIAVEL
when 'CFUN' then 'CARGO'
WHEN 'SCON' Then 'SALARIO'
else VARIAVEL
END AS TIPO,
upper(M.CONTEUDOANTERIOR) AS "Anterior",
upper(M.Conteudoatual) as "Atual",
upper(S.DESCRICAO40) AS "Setor/Obra Atual",
case m.motivoaltsalario
when '03' then 'Aumento Espontaneo'
when '04' then 'Promocao'
when '05' then 'Enquadramento'
when '09' then 'Reaj Sal. Complem'
when '07' then 'Antecipacao'
when '06' then 'Antec e Promo'
when '10' then 'Reajuste Salarial'
when '14' then 'Alt Sal Minimo Reg'
when '15' then 'Apuracao Custo Real'
when '16' then 'Reajuste Salarial'
when '17' then 'Red. Jornada MP936'
when '18' then 'Alteração de Contrato'
when '19' then 'Mudanca de Funcao'
when '97' then 'Red Decreto 10.422'
when '98' then 'Reducao de Jornada'
else m.motivoaltsalario
end as Motivo
FROM
RHMETA.RHALTERACOESCONTRATO M,
RHMETA.RHCONTRATOS C,
RHMETA.RHPESSOAS P,
rhmeta.rhsetores S
WHERE M.CONTRATO = C.CONTRATO
AND C.PESSOA = P.PESSOA
AND c.setor = S.SETOR
and VARIAVEL in ('CFUN','SCON')
AND OPERADOR <> 'CONVER'
AND m.motivoaltsalario not IN ('01','02','99')
order by 1 desc, 3,4
