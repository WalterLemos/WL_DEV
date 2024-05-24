select *
from (
    select p.nome, p.cpf, to_char(c.dataadmissao, 'dd/mm/yyyy') "Data Admissao", to_char(c.datarescisao, 'dd/mm/yyyy') "Data Demissao"
    from rhmeta.rhcontratos c, rhmeta.rhpessoas p
    where c.pessoa = p.pessoa
    and (
        c.dataadmissao >= trunc(trunc(current_date, 'MONTH') + INTERVAL '-1' DAY, 'MONTH') + INTERVAL '19' DAY
        or
        c.datarescisao >= trunc(trunc(current_date, 'MONTH') + INTERVAL '-1' DAY, 'MONTH') + INTERVAL '19' DAY
    )
) tmptable
order by nome
