# 📋 Guia de Estrutura do Arquivo Excel

## 📌 Requisitos Gerais

- **Formato**: .xlsx (não .xls)
- **Codificação**: UTF-8
- **Headers**: Linha 1
- **Dados**: Começam na linha 2
- **Nome das Abas**: Exatamente "Escavadeira" e "DF"

---

## 🔧 ABA "Escavadeira" (Perfil de Perda - Falhas Corretivas)

### Estrutura de Colunas

| Coluna | Nome | Tipo | Descrição | Exemplo |
|--------|------|------|-----------|---------|
| **A** | Data | Data (DD/MM/YYYY) | Data da ocorrência | 15/03/2026 |
| **B** | Equipamentos | Texto | Código do equipamento | 9401 |
| **C** | Descrição Ocorrência | Texto | Detalhe técnico | Vazamento hidráulico em braço |
| **D** | Sistema | Texto | Sistema afetado | Hidráulica |
| **E** | Sub-Sistema | Texto | Sub-sistema afetado | Mangueiras |
| **F** | Hora Início | Hora (HH:MM:SS) | Início da falha | 08:30:00 |
| **G** | Hora Fim | Hora (HH:MM:SS) | Fim da falha | 09:15:00 |
| **H** | Tempo Duração | Hora (HH:MM:SS) | Duração total | 00:45:00 |
| **I** | Descrição Falha | Texto | Análise da falha | Desgaste de vedação |

### Observações Importantes

✅ **Equipamentos válidos**: 9401, 9402, 9403, 9404, 9405, 9406, 9407
❌ **Equipamentos ignorados**: MM* (Motoniveladoras)

- Coluna B é a chave para filtrar (deve começar com "94")
- Parar de ler quando B estiver vazia
- Sem limite de linhas (crescimento dinâmico)

### Exemplo de Dados (Linhas 2+)

```
A          | B    | C                          | D         | E        | F        | G        | H        | I
-----------|------|----------------------------|-----------|----------|----------|----------|----------|------------
15/03/2026 | 9401 | Vazamento em braço         | Hidráulica| Mangueras| 08:30:00 | 09:15:00 | 00:45:00 | Desgaste vedação
15/03/2026 | 9402 | Falha A/C                  | Ar Cond.  | Compressor|10:00:00 | 10:30:00 | 00:30:00 | Não liga
16/03/2026 | 9401 | Fissura estrutura          | Estrutura | Braço    | 14:20:00 | 15:50:00 | 01:30:00 | Trinca na solda
17/03/2026 | 9403 | Vazamento em braço         | Hidráulica| Mangueras| 09:00:00 | 09:45:00 | 00:45:00 | Mesma falha 9401
```

---

## 📊 ABA "DF" (Disponibilidade da Frota)

### Estrutura Geral

Esta aba contém dois blocos de dados:

#### ESCAVADEIRAS (Colunas F-H) - ANALISADAS

**Linhas 2-10**: Dados por equipamento
**Linha 11**: Acumulado da frota

| Coluna | Linha | Campo | Tipo | Descrição | Exemplo |
|--------|-------|-------|------|-----------|---------|
| **F** | 2-10 | Equipamento | Texto | Código escavadeira | 9401 |
| **G** | 2-10 | DF | Número (%) | Disponibilidade | 85,50 ou 0.855 |
| **H** | 2-10 | Semana | Texto | Período | Semana 12 |
| **F** | 11 | "Acumulado da frota" | Texto | Label | Acumulado da frota |
| **G** | 11 | DF Total | Número (%) | **VALOR CRÍTICO** | 90,50 ou 0.905 |
| **H** | 11 | - | Vazio | - | - |

#### MOTONIVELADORAS (Colunas A-C) - PRESERVADAS

**Não processadas, apenas mantidas no arquivo**

| Coluna | Campo | Tipo | Descrição | Exemplo |
|--------|-------|------|-----------|---------|
| **A** | Equipamento | Texto | Código motoniveladora | MM01 |
| **B** | DF | Número (%) | Disponibilidade | 87,30 |
| **C** | Semana | Texto | Período | Semana 12 |

### Exemplo Completo da Aba "DF"

```
A      | B    | C        | ... | F           | G       | H
-------|------|----------|-----|-------------|---------|----------
MM01   | 87,30| Semana 12|     | 9401        | 85,50   | Semana 12
MM02   | 92,10| Semana 12|     | 9402        | 88,20   | Semana 12
MM03   | 89,70| Semana 12|     | 9403        | 84,60   | Semana 12
MM04   | 91,50| Semana 12|     | 9404        | 91,30   | Semana 12
       |      |          |     | 9405        | 86,80   | Semana 12
       |      |          |     | 9406        | 92,40   | Semana 12
       |      |          |     | 9407        | 88,90   | Semana 12
       |      |          |     | Acumulado da frota | 90,50 |
```

---

## ⚙️ Validações Automáticas

O sistema realiza as seguintes validações:

✅ **Percentuais com vírgula** (10,5%) são convertidos automaticamente para (10.5)
✅ **Linhas vazias** são ignoradas
✅ **Abas obrigatórias** (Escavadeira + DF) verificadas
✅ **Equipamentos 94\*** filtrados automaticamente
✅ **Arquivo deletado** após processamento (sem armazenamento)

---

## 🔍 Checklist antes do Upload

- [ ] Arquivo está em formato **.xlsx**
- [ ] Aba "Escavadeira" existe com nome exato
- [ ] Aba "DF" existe com nome exato
- [ ] Headers na linha 1
- [ ] Dados começam na linha 2
- [ ] Coluna B tem equipamentos (9401-9407)
- [ ] Coluna F tem equipamentos em DF
- [ ] Linha 11 em DF tem "Acumulado da frota"
- [ ] Valores numéricos são números (não texto)
- [ ] Datas no formato DD/MM/YYYY
- [ ] Horas no formato HH:MM:SS
- [ ] Percentuais com vírgula ou ponto (ambos aceitos)

---

## ❌ Erros Comuns

### Erro: "Aba 'Escavadeira' não encontrada"
**Causa**: Nome da aba está diferente (ex: "escavadeira", "Falhas", etc)
**Solução**: Renomeie exatamente para "Escavadeira"

### Erro: "Aba 'DF' não encontrada"
**Causa**: Nome da aba está diferente
**Solução**: Renomeie exatamente para "DF"

### Análise com 0 falhas
**Causa**: Coluna B vazia ou equipamentos não começam com "94"
**Solução**: Verifique que equipamentos são 9401-9407

### DF aparece como 0%
**Causa**: Linha 11 coluna G não tem valor
**Solução**: Certifique-se que F11="Acumulado da frota" e G11 tem percentual

### Arquivo não carrega
**Causa**: Arquivo em .xls em vez de .xlsx
**Solução**: Salve em Excel 2007+ (.xlsx)

---

## 📥 Exemplo de Arquivo para Download

Se precisar de um arquivo de exemplo, aqui está a estrutura mínima:

### Aba "Escavadeira"
```
Data       | Equipamentos | Descrição Ocorrência | Sistema    | Sub-Sistema | Hora Início | Hora Fim | Tempo Duração | Descrição Falha
-----------|--------------|----------------------|------------|-------------|-------------|----------|---------------|---------------
15/03/2026 | 9401         | Vazamento braço      | Hidráulica | Mangueiras  | 08:30:00    | 09:15:00 | 00:45:00      | Desgaste vedação
16/03/2026 | 9402         | Falha A/C            | Ar Cond.   | Compressor  | 10:00:00    | 10:30:00 | 00:30:00      | Não liga
```

### Aba "DF"
```
MM01 | 87,30 | Semana 12 | ... | 9401             | 85,50 | Semana 12
MM02 | 92,10 | Semana 12 | ... | 9402             | 88,20 | Semana 12
     |       |           | ... | 9403             | 84,60 | Semana 12
     |       |           | ... | Acumulado da frota | 90,50 |
```

---

**Dúvidas?** Verifique o README.md ou tente com um arquivo de teste simples primeiro.
