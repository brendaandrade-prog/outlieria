# Outlieria — Streamlit

Aplicação independente para detecção de outliers agrícolas em planilhas Excel usando limites SHP, histórico regional, cercas de Tukey e consolidação por município/estado.

## Executar localmente

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run streamlit_app.py
```

Na aba **Fontes**, envie a planilha Excel e o arquivo `.shp`. O processamento começa ao clicar em **Pronto**.

## Publicar no Streamlit Community Cloud

1. Acesse `share.streamlit.io`.
2. Selecione o repositório `brendaandrade-prog/outlieria`.
3. Use `streamlit_app.py` como arquivo principal.
4. Se houver um endpoint oficial de localização IBGE, configure `IBGE_REVERSE_URL` nos Secrets.

O app não salva arquivos por padrão: os dados ficam na sessão e os downloads são gerados no navegador. Para persistência multiusuário, adicione banco/storage e autenticação antes de usar em produção.
