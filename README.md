# Superhar RH — Dashboard de Recrutamento & Seleção

## Como rodar localmente (rede interna do escritório)

### 1. Instale o Python
Baixe em https://python.org (versão 3.10 ou superior)

### 2. Instale as dependências
Abra o terminal na pasta do projeto e rode:
```
pip install -r requirements.txt
```

### 3. Inicie o dashboard
```
streamlit run app.py
```
O dashboard abre em http://localhost:8501

### Para compartilhar na rede interna:
Descubra o IP da sua máquina (ipconfig no Windows) e informe às colegas:
http://SEU_IP:8501 — ex: http://192.168.1.10:8501

---

## Como publicar no Streamlit Cloud (link público)

1. Crie conta em https://github.com
2. Crie um repositório público chamado `superhar-dashboard`
3. Suba os 3 arquivos: app.py, requirements.txt, README.md
4. Acesse https://share.streamlit.io
5. Conecte o repositório GitHub → Deploy
6. Compartilhe o link gerado com a equipe

---

## Estrutura esperada do CSV

| Coluna | Tipo | Exemplo |
|--------|------|---------|
| id_vaga | texto | VAG-001 |
| nome_vaga | texto | Gerente Comercial |
| cliente | texto | Samarco |
| area | texto | Comercial |
| nivel | texto | Gerência |
| consultora | texto | Cida (Maria Aparecida) |
| status | texto | Fechada / Em andamento / Cancelada |
| canal_origem | texto | LinkedIn |
| mapeados | número | 58 |
| short_list | número | 7 |
| entrevistados | número | 4 |
| contratados | número | 1 |
| time_to_fill | número (dias) | 28 |
| sla_cumprido | True/False | True |
| oferta_aceita | True/False | True |
| assertividade_6m | True/False | True |
| nps_cliente | número 0-10 | 9.2 |
| sal_ofertado | número (R$) | 18000 |
| genero_candidato | texto | Feminino |
| raca_candidato | texto | Parda |
| pcd | True/False | False |
| mes_abertura | texto | Jan |

Superhar Recursos Humanos — superhar.com.br
