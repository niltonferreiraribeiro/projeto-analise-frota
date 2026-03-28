# 🚀 Guia Completo de Deploy no Render.com

## 📌 Por que Render.com?

✅ **GRATUITO** - Nenhum custo indefinidamente
✅ **24/7 ONLINE** - Seu computador não precisa estar ligado
✅ **LINK ÚNICO** - Acesse de qualquer lugar: https://seu-projeto.onrender.com
✅ **AUTOMÁTICO** - Deployment automático ao fazer push no GitHub
✅ **HTTPS** - Certificado SSL gratuito incluído
✅ **SEGURO** - Dados processados e deletados (sem armazenamento)

---

## ⚡ Quick Start (5 minutos)

### 1. Criar Repositório GitHub

**Opção A: Linha de comando**

```bash
# Dentro da pasta do projeto
git init
git add .
git commit -m "Initial commit - Sistema de Análise de Falhas"

# Crie um novo repositório em https://github.com/new
# Nome sugerido: projeto-analise-frota

# Configure o remote
git remote add origin https://github.com/seu-usuario/projeto-analise-frota.git
git branch -M main
git push -u origin main
```

**Opção B: Sem linha de comando**
1. Acesse https://github.com/new
2. Crie um novo repositório
3. Use GitHub Desktop para fazer upload dos arquivos

### 2. Conectar ao Render.com

1. Acesse [render.com](https://render.com)
2. Clique em **Sign Up** (ou **Sign In** se já tem conta)
3. Use sua conta GitHub para login (recomendado)

### 3. Criar Web Service

1. No dashboard, clique em **"+ New"** no canto superior
2. Selecione **"Web Service"**
3. Clique em **"Connect Account"** para conectar seu GitHub
4. Selecione seu repositório `projeto-analise-frota`
5. Clique em **"Connect"**

### 4. Configurar Web Service

Preencha os campos:

| Campo | Valor |
|-------|-------|
| **Name** | `projeto-analise-frota` (ou outro nome) |
| **Region** | `Ohio` (US) ou mais próximo |
| **Branch** | `main` |
| **Build Command** | `pip install -r requirements.txt` |
| **Start Command** | `gunicorn app:app` |
| **Instance Type** | `Free` |

### 5. Deploy

1. Clique em **"Create Web Service"**
2. Aguarde 2-3 minutos pelo deploy
3. Quando terminar, verá a URL: `https://seu-projeto.onrender.com`

✅ **Pronto!** Seu sistema está online!

---

## 📝 Passo a Passo Detalhado

### A. Preparação Local

```bash
# 1. Certifique-se que os arquivos estão completos
projeto-analise-frota/
├── app.py
├── requirements.txt
├── .gitignore
├── README.md
├── ESTRUTURA_EXCEL.md
└── templates/
    └── index.html

# 2. Teste localmente (opcional mas recomendado)
python -m venv venv
# Windows: venv\Scripts\activate
# Mac/Linux: source venv/bin/activate

pip install -r requirements.txt
python app.py

# Acesse http://localhost:5000 no navegador
# Se funcionar, continue. Se não, corrija os erros
```

### B. GitHub - Com Linha de Comando

```bash
# 1. Inicialize Git
cd projeto-analise-frota
git init

# 2. Configure o usuário (se primeira vez)
git config --global user.name "Seu Nome"
git config --global user.email "seu@email.com"

# 3. Add e commit
git add .
git commit -m "Initial commit - Sistema de Análise de Falhas"

# 4. Crie repositório em https://github.com/new
# Nome: projeto-analise-frota
# Descrição: Sistema de Análise de Falhas e Confiabilidade

# 5. Configure remote
git remote add origin https://github.com/SEU-USUARIO/projeto-analise-frota.git
git branch -M main
git push -u origin main

# 6. Verifique em https://github.com/seu-usuario/projeto-analise-frota
```

### C. GitHub - Com GitHub Desktop

1. Abra [GitHub Desktop](https://desktop.github.com/)
2. File → Clone Repository → URL
3. Cole: `https://github.com/seu-usuario/projeto-analise-frota.git`
4. Escolha local
5. File → Add Local Repository
6. Selecione a pasta `projeto-analise-frota`
7. Clique em "Publish repository"
8. Dê nome ao repositório
9. Pronto!

### D. Render.com - Deploy

1. Acesse [render.com](https://render.com) → Dashboard
2. **"+ New"** → **"Web Service"**
3. **"Connect to GitHub"** (autorize se necessário)
4. Busque `projeto-analise-frota`
5. Clique **"Connect"**

**Configurações:**
```
Name: projeto-analise-frota
Environment: Python 3
Build Command: pip install -r requirements.txt
Start Command: gunicorn app:app
Free Plan: ✓ Selecionado
```

6. Clique **"Create Web Service"**
7. Aguarde deploy (veja logs em tempo real)

---

## ✅ Verificar Deploy

### Quando o Deploy Está Completo

Na aba "Events" você verá:
```
✓ Deploy started
✓ Build started
✓ Build completed
✓ Deploy live
```

Sua URL estará em:
```
https://seu-projeto.onrender.com
```

### Testar Funcionamento

1. Acesse `https://seu-projeto.onrender.com` no navegador
2. Deve carregar a interface com upload
3. Se não carregar:
   - Aguarde 30 segundos e tente novamente
   - Verifique os "Logs" no Render
   - Procure erros em vermelho

### Erro Comum: "503 Service Unavailable"

**Causa**: Serviço ainda iniciando ou erro na inicialização

**Solução**:
1. Aguarde 2 minutos
2. Atualizar a página (F5)
3. Se persistir, abra a aba "Logs" e procure erros

---

## 🔧 Manutenção e Atualizações

### Atualizar Código

```bash
# Faça mudanças nos arquivos localmente
# Por exemplo, edite app.py

# Commit e push
git add .
git commit -m "Descrevendo a mudança"
git push origin main

# Render fará redeploy automaticamente!
# Acompanhe em "Events" no dashboard
```

### Ver Logs

1. Dashboard do Render
2. Seu Web Service
3. Aba "Logs"
4. Veja eventos em tempo real

### Reiniciar Serviço

1. Dashboard → Seu serviço
2. Menu (três pontos) → "Restart service"
3. Aguarde 30 segundos

---

## 🔌 Manter Serviço Ativo (Evitar "Sleep")

No plano Free, o serviço pode "dormir" após 15 minutos de inatividade.

### Solução 1: UptimeRobot (Recomendado)

```
1. Acesse https://uptimerobot.com
2. Sign Up gratuito
3. + Add new monitor
4. Tipo: HTTP(s)
5. URL: https://seu-projeto.onrender.com
6. Interval: 5 minutos
7. Save
```

Isso "pinga" seu serviço a cada 5 minutos, mantendo-o sempre ativo.

### Solução 2: GitHub Actions (Automático)

1. Crie arquivo `.github/workflows/keep-alive.yml`:

```yaml
name: Keep Render alive

on:
  schedule:
    - cron: '*/10 * * * *'

jobs:
  ping:
    runs-on: ubuntu-latest
    steps:
      - name: Ping Render
        run: curl https://seu-projeto.onrender.com
```

2. Faça commit e push
3. GitHub Actions rodará a cada 10 minutos

---

## 🚨 Troubleshooting

### "Build failed"

Verifique em "Logs" qual foi o erro:

**Erro: ModuleNotFoundError: No module named 'flask'**
- Causa: requirements.txt não está correto
- Solução: Verifique arquivo requirements.txt

**Erro: port already in use**
- Causa: Configuração de porta errada
- Solução: Deixe como `gunicorn app:app` (Render defini a porta)

### "Service unavailable"

1. Aguarde 2 minutos
2. Atualize página
3. Reinicie o serviço

### Timeout ao carregar página

- Seu arquivo Excel é muito grande?
- Limite testes a 50-100 linhas de dados

### Arquivo não processa

- Verifique estrutura Excel conforme ESTRUTURA_EXCEL.md
- Teste com arquivo de exemplo primeiro

---

## 📊 Monitorar Performance

### Verificar uso de recursos

No Render Dashboard:
1. Seu serviço
2. Aba "Metrics"
3. Veja CPU, memória, requisições

### Limpar cache (se necessário)

```bash
# Forçar rebuild
git commit --allow-empty -m "Trigger rebuild"
git push
```

---

## 💡 Dicas Pro

1. **Mensagens de commit claras**: `git commit -m "Add Excel validation"`
2. **Testar localmente antes de push**: Evita erros em produção
3. **UptimeRobot + GitHub + Render** = Setup perfeito
4. **Logs são seu amigo**: Sempre verifique quando algo não funciona
5. **Backups**: Seu repositório GitHub é o backup!

---

## 📞 Suporte Render

- Dashboard: https://dashboard.render.com
- Documentação: https://render.com/docs
- Status: https://status.render.com

---

## ✨ Próximos Passos

Depois que Deploy estiver online:

1. ✅ Testar com arquivo Excel real
2. ✅ Compartilhar URL com time
3. ✅ Configurar UptimeRobot
4. ✅ Monitorar primeiros uploads
5. ✅ Planejar versão 2.0 com gráficos

---

**Parabéns!** 🎉 Seu sistema está online e acessível 24/7!
