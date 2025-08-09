# 📘 S.A.G.E. - Sistema de Administração e Gestão Educacional

> 💼 **Status do Projeto:** Em desenvolvimento  
> 🔧 **Fase atual:** Estruturação inicial da Ribbon personalizada no Excel

```
🛠️ Progresso por módulo:
[🟡] Ribbon (Faixa de Opcoes)         🔨 Em desenvolvimento
[⚪] Login e Autenticacao             ⏳ Ainda nao iniciado
[⚪] Cadastros e Gerenciamento        ⏳ Ainda nao iniciado
[⚪] Cronogramas e Mapas              ⏳ Ainda nao iniciado
```

---

## 📁 Estrutura Inicial do Projeto

```
S.A.G.E/
├── ThisWorkbook.vb
├── Planilha1.vb  ← (MENU)
├── Planilha2.vb  ← (CONFIGURACOES)
├── Planilha3.vb  ← (CRONOGRAMA CALENDARIO)
├── Planilha4.vb  ← (CALENDARIO ACADEMICO)
├── Planilha5.vb  ← (CRONOGRAMA MAPA DE SALA)
├── Planilha6.vb  ← (MAPA DE SALA)
├── S.A.G.E.xml

├── Aplicacao/
│   └── FaixaDeOpcoes/
│       ├── Menu/
│       │   ├── Ribbon_Menu.vb
│       │   └── Ribbon_Menu.xml
│       ├── Configuracoes/
│       ├── Controles/
│       ├── Acoes/
│       └── Nucleo/

```

---

## 🚧 Etapa atual

No momento, o projeto está sendo iniciado pela **implementação da Ribbon personalizada** no Excel, incluindo:

- Estruturação do XML (`Ribbon_Menu.xml`)
- Implementação dos callbacks (`Ribbon_Menu.vb`)
- Controle de tabs, botões e estado da interface

As demais funcionalidades ainda serão construídas.

---

## 🛠 Tecnologias Utilizadas

- VB.NET com VSTO (Excel)
- SQLite (em breve)
- XML para interface Ribbon
- Git + Conventional Commits

---

## ✍️ Autor

Desenvolvido por **Jhefferson Wellys**  
[github.com/jheffersonwellys](https://github.com/jheffersonwellys)

---
