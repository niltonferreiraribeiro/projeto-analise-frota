# ⚡ Guia Rápido - 5 Minutos para Começar

## 🎯 Seu Sistema Está Pronto!

Parabéns! Todos os arquivos foram criados. Agora você tem 3 opções:

---

## ✅ Opção 1: Usar Online (RECOMENDADO)

### Se você QUER usar via internet (24/7 online)

**Tempo: ~10 minutos**

1. **Crie conta GitHub** → https://github.com/signup
2. **Crie repositório** → https://github.com/new
   - Nome: `projeto-analise-frota`
   - Faça upload dos arquivos (ou use linha de comando)

3. **Crie conta Render** → https://render.com
4. **Conecte GitHub ao Render**
5. **Deploy automático!**

**Resultado**: `https://seu-projeto.onrender.com` ✨

📖 **Instruções detalhadas**: `DEPLOY_RENDER.md`

---

## ✅ Opção 2: Usar Localmente

### Se você QUER testar no seu PC AGORA

**Tempo: ~3 minutos**

```bash
# 1. Abra terminal/command prompt na pasta do projeto

# 2. Crie ambiente virtual
python -m venv venv

# 3. Ative (Windows)
venv\Scripts\activate

# 3. Ative (Mac/Linux)
source venv/bin/activate

# 4. Instale dependências
pip install -r requirements.txt

# 5. Inicie servidor
python app.py

# 6. Abra navegador
# http://localhost:5000
```

**Resultado**: Sistema funciona no seu PC 🖥️

📖 **Instruções completas**: `README.md`

---

## ✅ Opção 3: Ambas!

1. Use **Opção 2** para testar localmente agora
2. Depois faça **Opção 1** para deploy online

---

## 📊 Usar o Sistema

### 1. Prepare o Arquivo Excel

Seu arquivo deve ter 2 abas:
- **"Escavadeira"** - Dados de falhas
- **"DF"** - Disponibilidade física

📖 **Estrutura exata**: `ESTRUTURA_EXCEL.md`

### 2. Acesse o Sistema

```
Local: http://localhost:5000
Online: https://seu-projeto.onrender.com
```

### 3. Faça Upload

1. Arraste ou clique para selecionar arquivo `.xlsx`
2. Defina meta de DF (padrão: 90,50%)
3. Clique em **"Gerar Análise"**
4. Visualize relatório em segundos

### 4. Exporte Relatório

- **Imprimir**: Botão "Imprimir" para PDF
- **Baixar**: Botão "Baixar HTML"

---

## 📁 Arquivos Criados

```
projeto-analise-frota/
├── app.py                      ← Backend (não edite)
├── requirements.txt            ← Dependências (não edite)
├── .gitignore                  ← Config Git (não edite)
├── templates/
│   └── index.html             ← Interface (não edite)
├── README.md                   ← Documentação completa
├── ESTRUTURA_EXCEL.md          ← Como fazer o Excel
├── DEPLOY_RENDER.md            ← Como colocar online
├── GUIA_RAPIDO.md             ← Este arquivo
└── uploads/                    ← Criado automaticamente
```

---

## 🚀 Próximos Passos Recomendados

### Hoje
- [ ] Teste uma opção (local ou online)
- [ ] Faça upload de um arquivo Excel teste
- [ ] Verifique se o relatório gerou

### Essa semana
- [ ] Use no seu fluxo diário
- [ ] Configure UptimeRobot (se online)
- [ ] Compartilhe com o time

### Mês que vem
- [ ] Versão 2.0 com gráficos
- [ ] Histórico de dados
- [ ] Dashboard

---

## ❓ Dúvidas?

| Pergunta | Resposta | Arquivo |
|----------|----------|---------|
| "Como faço o Excel?" | Veja exemplos de estrutura | `ESTRUTURA_EXCEL.md` |
| "Como uso localmente?" | Siga os passos do Quick Start | `README.md` (seção Local) |
| "Como coloco online?" | Passo a passo do Render | `DEPLOY_RENDER.md` |
| "Qual é o custo?" | 100% gratuito (Render Free) | `DEPLOY_RENDER.md` |
| "Meu PC precisa ficar ligado?" | NÃO! Render é 24/7 | Qualquer arquivo |
| "Onde os arquivos ficam armazenados?" | Nenhum lugar! Deletado após análise | `app.py` (linha 180) |

---

## 🎯 Checklist Rápido

### Para Usar Local
- [ ] Python 3.8+ instalado
- [ ] Navegador atualizado
- [ ] Arquivo Excel pronto
- [ ] Rodou `python app.py`
- [ ] Acesso http://localhost:5000

### Para Usar Online
- [ ] Conta GitHub criada
- [ ] Repositório criado com arquivos
- [ ] Conta Render criada
- [ ] Web Service criado no Render
- [ ] Deploy completado
- [ ] URL acessível

---

## 💡 Dicas

1. **Teste com arquivo simples primeiro** - Veja `ESTRUTURA_EXCEL.md` para exemplo mínimo
2. **Navegador moderno** - Chrome, Firefox, Edge (IE não funciona)
3. **Arquivo não muito grande** - Teste com 50-100 linhas primeiro
4. **Compartilhe o link** - Qualquer pessoa com o link (online) pode usar
5. **Tudo é confidencial** - Arquivos são deletados após processamento

---

## 🔐 Segurança

✅ Seu arquivo é deletado após análise
✅ Nenhum dado é armazenado
✅ Análise completa ocorre no servidor
✅ Conexão HTTPS (se online)
✅ Sem tracking ou coleta de dados

---

## 📞 Suporte Rápido

**"Sistema não carrega"**
- Aguarde 30 segundos
- Atualize página (F5)
- Se online, check https://status.render.com

**"Upload não funciona"**
- Arquivo é .xlsx? (não .xls)
- Arquivo tem 2 abas: "Escavadeira" e "DF"?
- Veja `ESTRUTURA_EXCEL.md`

**"Relatório em branco"**
- Verifique estrutura Excel em `ESTRUTURA_EXCEL.md`
- Teste com arquivo de exemplo

---

## 🎁 O que você tem:

✅ Sistema web profissional
✅ Interface moderna e responsiva
✅ Processamento automático em <5 segundos
✅ Relatório HTML formatado
✅ Importação de Excel simples
✅ Acessível 24/7 (se online)
✅ Sem custos
✅ Sem instalações necessárias
✅ Seguro e confidencial
✅ Pronto para produção

---

## 🚀 Comece Agora!

### Rápido? (3 min)
```bash
python app.py
```
Abra `http://localhost:5000`

### Online? (10 min)
Siga `DEPLOY_RENDER.md`

### Dúvida sobre Excel?
Leia `ESTRUTURA_EXCEL.md`

---

**Pronto! Seu sistema está completo e funcional.** 🎉

Dúvidas específicas? Verifique o arquivo relevante acima!
