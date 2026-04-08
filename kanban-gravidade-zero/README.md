# Gravidade Zero - Kanban Task Manager 🚀

Um gerenciador de tarefas premium estilo Kanban, focado em alta performance e estética moderna.

## 🛠️ Tecnologias
- **Runtime**: Bun
- **Backend**: Bun.serve + SQLite (`bun:sqlite`)
- **Frontend**: React + Vite + TypeScript
- **Estilização**: Vanilla CSS (Slate Dark Theme / Glassmorphism)

## 📁 Estrutura do Projeto
- `/backend`: API REST e lógica de banco de dados.
- `/frontend`: Interface do usuário construída com React.
- `/db`: Arquivo SQLite para persistência local.

## 🚀 Como Executar

### Pré-requisitos
- [Bun](https://bun.sh) instalado.

### Backend
```bash
cd backend
bun install
bun run index.ts
```

### Frontend
```bash
cd frontend
bun install
bun run dev
```

## 📝 Funcionalidades
- CRUD completo de tarefas.
- Categorização com emojis.
- Níveis de prioridade (Alta, Média, Baixa).
- Movimentação entre colunas (A Fazer -> Em Progresso -> Concluída).
- Persistência em banco de dados SQLite local.
