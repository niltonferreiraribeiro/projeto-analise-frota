# 📊 Sistema de Análise de Falhas e Confiabilidade - Frota de Escavadeiras CAT

Um sistema web profissional que automatiza a análise de falhas corretivas e disponibilidade (DF) da frota de escavadeiras CAT, gerando relatórios HTML em segundos.

## 🎯 Objetivo

Eliminar análises manuais que levam 2+ horas/dia, substituindo por um sistema automático que:
- ✅ Processa em SEGUNDOS
- ✅ Gera relatórios profissionais HTML
- ✅ Identifica padrões automáticamente
- ✅ Acessível de qualquer computador via link único
- ✅ Funciona 24/7 sem instalação

## 🚀 Quick Start (Local)

### 1. Requisitos
- Python 3.8+
- pip
- Git (opcional)

### 2. Instalação

```bash
# Clone ou baixe os arquivos
cd projeto-analise-frota

# Crie um ambiente virtual
python -m venv venv

# Ative o ambiente virtual
# No Windows:
venv\Scripts\activate
# No Mac/Linux:
source venv/bin/activate

# Instale as dependências
pip install -r requirements.txt
```

### 3. Executar Localmente

```bash
# Inicie o servidor Flask
python app.py

# Acesse no navegador
# http://localhost:5000
```

## 🌐 Deploy no Render.com (Gratuito)

### Por que Render.com?
✅ **GRATUITO** indefinidamente
✅ **SEMPRE ONLINE** 24/7
✅ **LINK ÚNICO**: https://seu-projeto.onrender.com
✅ Deployment automático via GitHub
✅ HTTPS automático
✅ Acessível de qualquer computador/celular

### Passo a Passo

#### 1. Criar Repositório GitHub

```bash
# Inicialize Git local
git init

# Add todos os arquivos
git add .

# Commit inicial
git commit -m "Initial commit - Sistema de Análise de Falhas"

# Se ainda não tem repositório, crie em https://github.com/new
# Adicione remote
git remote add origin https://github.com/seu-usuario/projeto-analise-frota.git

# Push para GitHub
git branch -M main
git push -u origin main
```

#### 2. Deploy no Render.com

1. Acesse [render.com](https://render.com)
2. Clique em **"+ New +"** → **"Web Service"**
3. Conecte seu repositório GitHub
4. Preencha os dados:
   - **Name**: `projeto-analise-frota` (ou seu nome)
   - **Branch**: `main`
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app`
   - **Environment**: Python 3.10
   - **Plan**: Free (gratuito)

5. Clique em **"Create Web Service"**

6. Aguarde o deploy (2-3 minutos)

7. Seu URL estará em: `https://seu-projeto.onrender.com`

#### 3. Configurações Importantes

Se seu serviço "dormir" (após 15 minutos sem uso), você pode:
- Usar um serviço de ping gratuito como [UptimeRobot](https://uptimerobot.com)
- OU acessar regularmente para manter ativo

**Configurar UptimeRobot (opcional):**
1. Acesse [uptimerobot.com](https://uptimerobot.com)
2. Crie conta gratuita
3. Adicione novo "Monitor" → "HTTP(s)"
4. Cole sua URL do Render
5. Defina interval para 5 minutos
6. Isso garante que o serviço nunca dorme

## 📊 Como Usar

### 1. Preparar o Arquivo Excel

Seu arquivo deve ter:
- **Aba "Escavadeira"** com colunas:
  - A: Data
  - B: Equipamentos (9401-9407)
  - C: Descrição Ocorrência
  - D: Sistema (Hidráulica, Estrutura, etc)
  - E: Sub-Sistema (Mangueiras, Câmara, etc)
  - F: Hora Início
  - G: Hora Fim
  - H: Tempo Duração
  - I: Descrição Falha

- **Aba "DF"** com:
  - Equipamentos em F2:F10
  - DF (%) em G2:G10
  - Acumulado da frota em F11
  - DF Total em G11

### 2. Acessar o Sistema

```
https://seu-projeto.onrender.com
```

### 3. Fazer Upload

1. Clique na área de upload ou arraste o arquivo
2. Defina a meta de DF (padrão: 90,50%)
3. Clique em "Gerar Análise"
4. Visualize o relatório em segundos

### 4. Exportar Relatório

- **Imprimir**: Clique em "Imprimir" para PDF
- **Baixar**: Clique em "Baixar HTML" para arquivo

## 📈 Dados do Relatório

O sistema gera automaticamente:

### KPIs (Key Performance Indicators)
- ✅ Total de Falhas
- ✅ Disponibilidade Física (DF) Atual
- ✅ Meta de DF
- ✅ Diferença vs Meta

### Análises
- ✅ TOP 3 Sistemas com mais falhas
- ✅ TOP 3 Sub-Sistemas com mais falhas
- ✅ TOP 3 Equipamentos com mais falhas
- ✅ Padrões recorrentes (falhas que se repetem)
- ✅ Alertas críticos (DF < 50%, equipamentos críticos)

### Recomendações
- ✅ Ações estratégicas baseadas nos dados
- ✅ Priorização de manutenção
- ✅ Planos de ação

## 🔒 Validações

O sistema valida:
- ✅ Apenas arquivos .xlsx (não .xls)
- ✅ Abas obrigatórias existem
- ✅ Equipamentos começam com "94"
- ✅ Conversão automática de percentuais (10,5% → 10.5)
- ✅ Ignorar linhas vazias
- ✅ Deletar arquivo após processamento (sem armazenamento)

## 📁 Estrutura do Projeto

```
projeto-analise-frota/
├── app.py                  # Backend Flask (processamento)
├── requirements.txt        # Dependências Python
├── .gitignore             # Configuração Git
├── README.md              # Este arquivo
├── templates/
│   └── index.html         # Frontend (Interface)
└── uploads/               # Pasta temporária (criada automaticamente)
```

## 🛠️ Stack Tecnológico

- **Backend**: Flask 3.1.3 (Python)
- **Frontend**: HTML5 + CSS3 + JavaScript Vanilla
- **Processamento**: openpyxl 3.1.5
- **Servidor**: Gunicorn 21.2.0
- **Host**: Render.com (Gratuito)

## 📝 Funcionalidades

### Frontend
- ✅ Upload com drag-and-drop
- ✅ Campo para meta de DF customizável
- ✅ Indicador de carregamento
- ✅ Visualização de relatório
- ✅ Botão imprimir/baixar
- ✅ Design responsivo

### Backend
- ✅ Leitura de múltiplas abas Excel
- ✅ Processamento e cálculo de KPIs
- ✅ Identificação de padrões recorrentes
- ✅ Classificação de prioridades
- ✅ Geração HTML profissional
- ✅ Tratamento robusto de erros

## ⚡ Performance

- **Upload**: < 1 segundo
- **Processamento**: < 2 segundos
- **Geração Relatório**: < 1 segundo
- **Total**: < 5 segundos

## 🔄 Próximas Versões (Roadmap)

- 🔄 Banco de dados com histórico
- 🔄 Dashboard com gráficos interativos
- 🔄 Exportação para PowerPoint
- 🔄 Autenticação de usuários
- 🔄 Análise Motoniveladoras
- 🔄 Integração com IA Claude

## 🆘 Troubleshooting

### Erro: "Aba 'Escavadeira' não encontrada"
- Verifique o nome da aba (case-sensitive)
- O arquivo precisa ter exatamente "Escavadeira" e "DF"

### Erro: "Apenas arquivos .xlsx são permitidos"
- Salve o Excel em formato .xlsx (não .xls)
- Não funciona com Google Sheets

### Arquivo não processa
- Verifique se os dados começam na linha 2
- Coluna B deve ter equipamentos (9401-9407)
- Linhas vazias em B encerram a leitura

### Sistema não responde após 30 min
- No Render gratuito, o serviço pode "dormir"
- Configure UptimeRobot para manter ativo

## 📞 Suporte

Para questões ou melhorias:
1. Verifique os logs na aba "Logs" do Render
2. Tente novamente com um arquivo diferente
3. Limpe o navegador (cache) e tente novamente

## 📜 Licença

Este projeto é fornecido como-está para uso interno.

## ✅ Checklist de Deploy

- [ ] Arquivos criados localmente
- [ ] `pip install -r requirements.txt` funciona
- [ ] `python app.py` inicia sem erros
- [ ] Repositório GitHub criado
- [ ] Render.com conectado ao GitHub
- [ ] Web Service criado no Render
- [ ] Deploy concluído com sucesso
- [ ] URL acessível de outro computador
- [ ] Upload e análise funcionam
- [ ] Relatório gera corretamente

---

**Pronto para produção!** 🚀
Seu sistema está online e acessível 24/7 via Render.com.
