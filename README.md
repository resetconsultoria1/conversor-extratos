# Conversor de extratos bancários → Excel

Anexe o extrato do banco em **PDF** ou **OFX** e baixe o resultado no padrão de
duas abas usado nos acompanhamentos:

| Aba | Colunas |
|-----|---------|
| `Debito`  | Data · Descrição · Valor · Obs |
| `Credito` | Data · Descrição · Valor · Obs |

- **Data**: `DD/MM/AAAA`
- **Valor**: número positivo (o sinal fica implícito pela aba)
- **Débito** = dinheiro que saiu · **Crédito** = dinheiro que entrou
- Lançamentos de saldo (SALDO ANTERIOR, SALDO DO DIA, etc.) são descartados
- Antes de baixar, dá para editar/apagar linhas e corrigir o tipo na tela

## Como cada formato é lido

| Formato | Como é lido | Precisa de chave de API? | Custo |
|---------|-------------|--------------------------|-------|
| **OFX** | Leitura direta do arquivo (100% offline) | Não | Zero |
| **PDF** | Enviado para a API da Anthropic (Claude), que entende qualquer layout de banco, inclusive PDF escaneado | Sim | Alguns centavos por extrato |

> Sempre que o banco oferecer download em **OFX**, prefira OFX: é de graça e mais exato.

## Publicar na nuvem (Streamlit Community Cloud) — só pelo navegador

Mesma receita do "App conf de classificações": GitHub + Streamlit, conta `resetconsultoria1`
(login pelo Google). Não precisa instalar nada.

### Parte 1 — Subir os arquivos no GitHub
1. Entrar em https://github.com com a conta `resetconsultoria1`.
2. **+** (canto superior direito) → **New repository**.
   - Name: `conversor-extratos`
   - Marcar **Private** (o `users.yaml` tem senha de usuário)
   - **Create repository**
3. Clicar no link **"uploading an existing file"**.
4. Arrastar da pasta `Conversor de extratos\` os arquivos:
   `app.py`, `requirements.txt`, `users.yaml`, `README.md`
5. **Commit changes**.

### Parte 2 — Publicar no Streamlit
1. Entrar em https://share.streamlit.io com **GitHub** (mesma conta).
2. **Create app** → **Deploy a public app from GitHub**
   (se pedir, **Authorize** o acesso aos repositórios privados).
3. Repository: `resetconsultoria1/conversor-extratos` · Branch: `main` · Main file: `app.py`
4. **Advanced settings** → **Secrets** → colar:
   ```toml
   ANTHROPIC_API_KEY = "sk-ant-..."
   ```
   (a chave está em `Curso Hastag\console.anthropic.com.txt` e na imagem `chave API .jpg`)
5. **Deploy** (~2–3 min). O endereço final é algo como
   `https://conversor-extratos-xxxx.streamlit.app`.

### Atualizar depois
Editar o arquivo direto no GitHub (ícone de lápis → **Commit changes**): o Streamlit
republica sozinho. A chave fica em **Manage app → Settings → Secrets**.

O login do app usa o `users.yaml` (mesmos usuários/senhas das outras ferramentas).

## Rodar no PC

Precisa do Python instalado (hoje a máquina não tem — instale de python.org).

```bash
pip install -r requirements.txt
```

Crie o arquivo `.streamlit/secrets.toml` a partir do `.streamlit/secrets.toml.example`
e coloque a chave. Depois:

```bash
streamlit run app.py
```

## Usuários e senhas

Ficam em `users.yaml` (senha guardada como hash SHA-256, igual às outras ferramentas).
Para adicionar alguém, gere o hash com:

```python
import hashlib; print("sha256:" + hashlib.sha256("A_SENHA".encode()).hexdigest())
```

Deixe `senha_hash` em branco (`senha_hash:`) para a pessoa definir a senha no primeiro acesso.

## Ajustes finos (topo do `app.py`)

| Constante | Para quê |
|-----------|----------|
| `MODELO_CLAUDE` | modelo usado na leitura do PDF |
| `PAGINAS_POR_LOTE` | quantas páginas vão por vez para a IA; diminua se um extrato muito grande vier incompleto |
| `MAX_TOKENS_SAIDA` | teto de resposta por lote |

## Limitações

- PDF **escaneado de baixa qualidade** pode gerar erro de leitura em algumas linhas — por isso existe a tela de conferência antes de baixar.
- A separação débito/crédito segue o sinal / marcação D-C do extrato; em bancos que não marcam, confira as primeiras linhas.
- Cartão de crédito (fatura) não é o foco: o padrão de saída é de extrato de conta corrente.
