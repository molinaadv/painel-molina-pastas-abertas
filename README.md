# Molina | Painel Pastas Abertas (Streamlit)

## O que é
Painel web para acompanhar **Pastas Abertas** do Legal One, com:
- Modo **Gestão** (com números)
- Modo **TV** (somente % da meta, pode passar de 100%)
- Ranking de **Indicações** com filtro por escritório
- Base filtrada (auditoria)

## Arquivos
- `app.py` (app Streamlit)
- `requirements.txt` (dependências)
- `logo_molina.png` (logo exibida no topo)

## Rodar local (Windows)
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Publicar online (Streamlit Community Cloud)
1) Crie um repositório no GitHub e envie estes arquivos.
2) Acesse o Streamlit Community Cloud e clique em **New app**.
3) Selecione o repo, branch `main` e o arquivo `app.py`.
4) Deploy.

Depois é só abrir o link e fazer upload do Excel do Legal One e do Excel de metas.
