<div align="center">

<h1>🚀 Gravidade Zero</h1>
<p><strong>Gerenciador de tarefas Kanban premium com drag & drop, colunas dinâmicas e design dark mode.</strong></p>

<p>
  <img src="https://img.shields.io/badge/React-18-61DAFB?style=flat-square&logo=react&logoColor=white" />
  <img src="https://img.shields.io/badge/TypeScript-5-3178C6?style=flat-square&logo=typescript&logoColor=white" />
  <img src="https://img.shields.io/badge/Bun-1.x-fbf0df?style=flat-square&logo=bun&logoColor=black" />
  <img src="https://img.shields.io/badge/SQLite-003B57?style=flat-square&logo=sqlite&logoColor=white" />
  <img src="https://img.shields.io/badge/Vite-646CFF?style=flat-square&logo=vite&logoColor=white" />
  <img src="https://img.shields.io/badge/Licença-MIT-green?style=flat-square" />
</p>

<p>
  <a href="#-funcionalidades">✨ Funcionalidades</a> •
  <a href="#-stack-técnica">🛠️ Stack</a> •
  <a href="#-como-rodar">▶️ Como Rodar</a> •
  <a href="#-estrutura-do-projeto">📁 Estrutura</a> •
  <a href="#-endpoints-da-api">📡 API</a> •
  <a href="#-autor">👤 Autor</a>
</p>

</div>

---

## ✨ Funcionalidades

- **📋 Board Kanban** — Organize suas tarefas em colunas de status com contadores em tempo real
- **🖱️ Drag & Drop** — Arraste cards entre colunas com animações fluidas (powered by `@hello-pangea/dnd`)
- **➕ Colunas Dinâmicas** — Crie, renomeie (clique no título) e exclua colunas livremente
- **✏️ CRUD Completo** — Crie, edite e exclua tarefas com título, descrição, prioridade, categoria e data
- **🔴🟡🟢 Prioridades Visuais** — Indicadores coloridos de Alta, Média e Baixa prioridade em cada card
- **🏷️ Categorias com Emoji** — Trabalho 💼, Estudo 📚, Saúde 🍎, Pessoal 🏠, Finanças 💰
- **🌙 Dark Mode** — Interface premium com tema escuro e efeito glassmorphism
- **📱 Responsivo** — Layout adaptável para diferentes tamanhos de tela
- **💾 Persistência Real** — Dados salvos em banco SQLite via Bun nativo

---

## 🛠️ Stack Técnica

| Camada | Tecnologia |
|--------|-----------|
| **Frontend** | React 18 + TypeScript + Vite |
| **Backend** | Bun runtime + HTTP server nativo |
| **Banco de Dados** | SQLite via `bun:sqlite` (sem ORM, zero deps) |
| **Drag & Drop** | `@hello-pangea/dnd` |
| **Ícones** | `lucide-react` |
| **HTTP Client** | `axios` |
| **Estilo** | CSS puro com variáveis e glassmorphism |

---

## ▶️ Como Rodar

### Pré-requisitos

- [Bun](https://bun.sh) `>= 1.0`
- [Node.js](https://nodejs.org) `>= 20.x` (para o Vite)

### 1. Clone o repositório

```bash
git clone https://github.com/ludolffbruno/Personal-projects.git
cd kanban-gravidade-zero
```

### 2. Instale as dependências

```bash
# Backend
cd backend && bun install

# Frontend
cd ../frontend && npm install
```

### 3. Popule o banco de dados

```bash
cd backend
bun run seed.ts
```

### 4. Inicie os servidores

**Backend** (porta 3030):
```bash
cd backend
bun run --watch index.ts
```

**Frontend** (porta 5173):
```bash
cd frontend
npm run dev
```

### 5. Acesse no navegador

```
http://localhost:5173
```

---

## 📁 Estrutura do Projeto

```
kanban-gravidade-zero/
├── backend/
│   ├── index.ts          # Servidor HTTP + rotas REST
│   ├── database.ts       # Conexão e schema SQLite
│   ├── seed.ts           # Dados iniciais (categorias + tarefas)
│   └── package.json
├── frontend/
│   ├── src/
│   │   ├── App.tsx       # Componente principal (board Kanban)
│   │   ├── App.css       # Estilos premium dark mode
│   │   └── api.ts        # Funções de comunicação com o backend
│   ├── index.html
│   └── package.json
├── db/
│   └── kanban.sqlite     # Banco de dados (gerado ao rodar seed)
└── README.md
```

---

## 📡 Endpoints da API

Base URL: `http://localhost:3030/api`

### Tarefas

| Método | Rota | Descrição |
|--------|------|-----------|
| `GET` | `/tasks` | Listar todas as tarefas |
| `POST` | `/tasks` | Criar nova tarefa |
| `PUT` | `/tasks/:id` | Atualizar tarefa |
| `DELETE` | `/tasks/:id` | Excluir tarefa |

### Colunas

| Método | Rota | Descrição |
|--------|------|-----------|
| `GET` | `/columns` | Listar colunas |
| `POST` | `/columns` | Criar nova coluna |
| `PUT` | `/columns/:id` | Renomear coluna |
| `DELETE` | `/columns/:id` | Excluir coluna e suas tarefas |

### Categorias

| Método | Rota | Descrição |
|--------|------|-----------|
| `GET` | `/categories` | Listar categorias disponíveis |

---

## 📸 Preview

> Board com tarefas reais distribuídas nas colunas **A Fazer**, **Em Progresso** e **Concluída**.

O board suporta criação de colunas extras como **Revisão**, **Teste**, **Bloqueado**, conforme a necessidade do fluxo de trabalho.

---

## 🔒 Segurança

- Nenhuma credencial ou informação sensível está presente no repositório
- O banco de dados SQLite é gerado **localmente** via seed — não é versionado no Git
- O arquivo `.gitignore` exclui `*.sqlite`, `*.log` e `node_modules`

---

## 👤 Autor

**Bruno Ludolff**

- GitHub: [@ludolffbruno](https://github.com/ludolffbruno)
- Projeto: [Personal-projects / kanban-gravidade-zero](https://github.com/ludolffbruno/Personal-projects/tree/main/kanban-gravidade-zero)

---

## 📄 Licença

Este projeto está licensed sob a **MIT License** — sinta-se livre para usar, modificar e distribuir.

---

<div align="center">
  <sub>Feito com ☕ e muito TypeScript por Bruno Ludolff</sub>
</div>
