# 🚀 Atualização Automática das Bases BI - GPE

Automação desenvolvida para atualizar as bases utilizadas no BI através da extração de dados do SAP GUI Scripting.

## 📋 Objetivo

Este projeto automatiza a atualização das planilhas utilizadas pelo BI, reduzindo o tempo de execução, eliminando atividades manuais e garantindo maior confiabilidade no processo.

As informações são extraídas diretamente do SAP e salvas em planilhas Excel utilizadas posteriormente pelos dashboards.

---

# ⚙️ Funcionalidades

* Extração automática de dados do SAP.
* Execução de múltiplas transações em sequência.
* Atualização automática das bases Excel.
* Backup das bases anteriores.
* Execução por arquivo `.bat`.
* Compatível com SAP GUI Scripting.
* Processo totalmente automatizado.

---

# 📂 Estrutura do Projeto

```text
📦 Atualizar Bases BI
│
├── 1 - MB51
├── 2 - MB52
├── 3 - MI24
├── 4 - ME2N
├── 5 - MM60
├── 6 - ZMM039
├── 7 - ZMM028
├── 8 - atualizar base bi v1.0
├── 9 - MB25
├── 10 - Backup bases bi
│
├── atualizar_bases_bi_gpe_v2_sem_limite_50mb.py
├── executar_atualizar_bases_bi_gpe_v2_corrigido.bat
└── SAP_GPE_TODAS_TRANSACOES_v1.3.vbs
```

---

# 📊 Transações SAP utilizadas

* MB25
* MB51
* MB52
* MI24
* ME2N
* MM60
* ZMM028
* ZMM039

---

# 🛠 Tecnologias

* Python 3
* SAP GUI Scripting (VBScript)
* Microsoft Excel
* Windows Batch (.bat)

---

# ▶️ Como executar

## 1. Abra o SAP

Faça login normalmente.

---

## 2. Execute o arquivo

```
executar_atualizar_bases_bi_gpe_v2_corrigido.bat
```

O script irá:

1. Conectar ao SAP.
2. Executar as transações automaticamente.
3. Exportar os relatórios.
4. Atualizar as bases do BI.
5. Salvar os arquivos atualizados.

---

# 📁 Arquivos principais

| Arquivo                                            | Descrição                     |
| -------------------------------------------------- | ----------------------------- |
| `atualizar_bases_bi_gpe_v2_sem_limite_50mb.py`     | Script principal da automação |
| `SAP_GPE_TODAS_TRANSACOES_v1.3.vbs`                | Automação das transações SAP  |
| `executar_atualizar_bases_bi_gpe_v2_corrigido.bat` | Inicializador da aplicação    |

---

# 📌 Requisitos

* Windows
* Python 3.x
* SAP GUI instalado
* SAP GUI Scripting habilitado
* Microsoft Excel

---

# 💡 Fluxo da Automação

```text
SAP Login
      │
      ▼
Executa Transações
      │
      ▼
Exporta Excel
      │
      ▼
Atualiza Bases
      │
      ▼
Realiza Backup
      │
      ▼
Finaliza
```

---

# 📈 Benefícios

* Redução do tempo de atualização das bases.
* Padronização do processo.
* Eliminação de tarefas repetitivas.
* Diminuição de erros manuais.
* Atualização consistente dos dados utilizados pelo BI.

---

# 👨‍💻 Autor

**Rogério Henrique Rocha de Matos**

GitHub: https://github.com/RogerioHenriqueRM

---

# 📄 Licença

Este projeto foi desenvolvido para uso interno e apoio às rotinas de atualização das bases de Business Intelligence (BI).
