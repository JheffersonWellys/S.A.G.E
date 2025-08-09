# 📘 S.A.G.E. - Sistema de Administração e Gestão Educacional

> 💼 **Status do Projeto:** Em desenvolvimento  
> 🔧 **Fase atual:** Implementação do módulo de Banco de Dados e integração inicial com a Ribbon

```
🛠️ Progresso por módulo:
[🟢] Ribbon (Faixa de Opcoes)         ✅ Estruturada e funcional
[🟡] Banco de Dados (SQLite)          🔨 Em desenvolvimento
[⚪] Login e Autenticacao              ⏳ Ainda nao iniciado
[⚪] Cadastros e Gerenciamento         ⏳ Ainda nao iniciado
[⚪] Cronogramas e Mapas               ⏳ Ainda nao iniciado
```

---

## 📁 Estrutura Atual do Projeto

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
├── Infraestrutura/                             ← **NOVO** 
│   ├── BancoDeDados/                           ← **NOVO**
│   │   ├── Configuracoes/                      ← **NOVO**
│   │   │   └── ConfiguracaoBanco.vb            ← **NOVO**
│   │   ├── DAL/                                ← **NOVO**
│   │   │   └── DAL_*.vb (um por entidade)      
│   │   ├── BLL/                                ← **NOVO**
│   │   │   └── BLL_*.vb (um por entidade)         
│   │   └── Modelos/                            ← **NOVO**
│   │       └── Clss_*.vb (um por entidade)
│   └── XML/                                    ← **NOVO**
│       └── XML_ControleDeAcesso.vb
```

---

## 🚧 Etapa Atual

O projeto agora entra na fase de **desenvolvimento do Banco de Dados**, incluindo:

- Criação do esquema em SQLite
- Estruturação dos módulos DAL (Data Access Layer) e BLL (Business Logic Layer)
- Integração inicial com botões e menus da Ribbon
- Preparação para controle de acesso e CRUD das principais entidades

---

## 🛠 Tecnologias Utilizadas

- VB.NET com VSTO (Excel)
- SQLite
- XML para interface Ribbon
- Git + Conventional Commits

---

## ✍️ Autor

Desenvolvido por **Jhefferson Wellys**  
[github.com/jheffersonwellys](https://github.com/jheffersonwellys)
