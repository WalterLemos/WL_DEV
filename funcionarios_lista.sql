            SELECT
            CASE c.setor
            when '0378' then 823
            when '0394' then 821
            when '0401' then 863
            when '0413' then 895
            when '0485' then 1019
            when '0470' THEN 982
            when '0481' then 1009
            when '0489' then 1036
            when '0507' then 1059
            when '0495' then 1064
            when '0484' then 1017
            when '0502' then 1083
            when '0514' then 1122
            when '0515' then 1123
            when '0509' then 1103
            when '0508' then 1090
            when '0515' then 1123
            when '0523' then 1154
            when '0495' then 1064
            when '0525' then 1159
            when '0519' then 1135
            when '0526' then 1161
            when '0534' then 1185
            when '0518' then 1134
            when '0063' then 345
            when '0537' then 1190
            when '0546' then 1238
            when '0326' then 591
            when '0545' then 1220
            when '0548' then 1243
            when '0550' then 1229
            when '0546' then 1238
            when '0549' then 1244
            when '0565' then 1293
            when '0554' then 1260
            when '0566' then 1296
            when '0545' then 1272
            when '0571' then 1316
            ELSE to_number(c.setor)
            END AS ID,
            S.DESCRICAO40 AS "Nome Setor",
            C.Contrato as Contrato,
            p.NOME,
            to_char(c.dataadmissao, 'dd/mm/yyyy') as Admissao,
            to_char(c.datarescisao, 'dd/mm/yyyy') as Demissao,
            c.motivorescisao,
            CASE WHEN c.motivorescisao IS NULL THEN '' 
            ELSE (SELECT mr.DESCRICAO40 FROM RHMETA.RHMOTIVOSRESCISOES mr WHERE mr.motivorescisao = c.motivorescisao) 
            END  AS DescMotivoRescisao,
            CA.DESCRICAO40 AS CARGO,
            CA.CBONOVO AS CBO,
            c.salariomes as Salario,
            to_char(P.NASCIMENTO, 'dd/mm/yyyy') as Nascimento,
            P.SEXO,
            ec.descricao20 as Estado_Civil,
            decode(P.CPF ,NULL,NULL,translate(to_char(P.CPF / 100, '000,000,000.00'), ',.', '.-')) CPF,
            p.identidade as RG,
            to_char(p.dataidentidade, 'dd/mm/yyyy') as RG_Data,
            p.orgaoemissor as RG_Emissor,
            p.ufidentidade as UF_Emissor,
            p.pis as PIS,
            P.NROCARTTRAB as CTPS,
            p.seriecarttrab as Serie,
            p.ufcarttrab as UF,
            p.nrocartaosus as Cartao_SUS,
            p.registrohabilitacao as CNH,
            p.categoriahabilitacao as Categoria_CNH,
            to_char(p.validadehabilitacao, 'dd/mm/yyyy') as Validade_CNH,
            p.localnascimento as Cidade_Nascimento,
            p.ufnascimento as UF_Nascimento,
            P.TIPOLOGRADOURO,    
            UPPER(P.RUA) AS LOGRADOURO,
            P.NRORUA AS NUMERO,
            P.COMPLEMENTO,
            P.BAIRRO,
            P.CIDADE,
            P.UF,
            P.CEP
            ,p.telefone as Telefone
            ,p.telefonecelular as Celular
            ,UPPER(P.MAE) AS MAE
            ,UPPER(P.PAI) AS PAI
            ,Case p.racacor  when '1' then 'Indigena' when '2' then 'Branca' when '4' then 'Preta' when '8' then 'Parda' when '9' then 'Não Informada' else p.racacor end as Raca
            ,to_char(c.datavenctoferias, 'dd/mm/yyyy') as "Data Vencimento Ferias"
            ,(select n.descricao20 from rhmeta.rhnacionalidades n where n.nacionalidade = p.nacionalidade) as Nacionalidade
            ,g.descricao20 as Graducao
            ,(SELECT count(gp.pessoa) FROM RHMETA.RHFAMILIARES gp WHERE  gp.graudependencia <> 0 and gp.PESSOA = P.PESSOA group by gp.pessoa) as "QTA DEPENDENTE"
            ,Case c.situacao when '1' then 'ATIVO' WHEN '2' then 'AFASTADO' WHEN '3' then 'DEMITIDO' WHEN '4' THEN 'DEMTIDO' END AS STATUS
            ,(select loginad from rhmeta.RHPESSOALOGINAD lg where lg.pessoa = P.pessoa) as Login_Site
            ,CASE p.deficientefisico
            when '0' Then 'Não Deficiente'
            when '1' Then 'Deficiente Fisico'
            when '2' Then 'Deficiente Auditivo'
            when '3' Then 'Deficiente Visual'
            when '4' Then 'Deficiente Intelectual'
            end as Deficiencia
            ,case c.estabelecimento
            when '0001' then 'Rio de Janeiro'
            when '0002' then 'São Paulo'
            when '0003' then 'Espirito Santo'
            when '0004' then 'Minas Gerais'
            when '0005' then 'Para'
            end as Filial
            ,to_char(c.dataultimoreajuste, 'dd/mm/yyyy') dataultimoreajuste,
            E.INSCRICAO CNPJ
            FROM RHMETA.RHCONTRATOS C,
            RHMETA.RHESTABELECIMENTOS E,
            RHMETA.RHPESSOAS P,
            RHMETA.RHUNIDADES U,
            RHMETA.RHCARGOS CA,
            rhmeta.rhnacionalidades N,
            rhmeta.rhestadocivil EC,
            rhmeta.rhsetores S,
            rhmeta.rhgrauinstrucao G
            WHERE C.PESSOA = P.PESSOA
            AND C.UNIDADE = U.UNIDADE
            AND C.CARGO = CA.CARGO
            AND C.ESTABELECIMENTO = E.ESTABELECIMENTO
            AND c.setor = S.SETOR
            AND P.NACIONALIDADE = N.NACIONALIDADE
            and p.estadocivil = ec.estadocivil
            and p.grauinstrucao = g.grauinstrucao
            --AND C.SITUACAO in (1,2)
            AND C.PESSOA <> 9000001
            and CA.DESCRICAO40 <> 'AUTONOMO'
            ORDER BY 4
