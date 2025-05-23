import fitz  # PyMuPDF
import re
import pandas as pd
import streamlit as st
from io import BytesIO

# --- Expressões Regulares ---
# Linha DARF: código(4 dígitos), descrição, principal, juros, multa, total
pattern_linha = re.compile(
    r'(?P<codigo>\d{4})\s+'
    r'(?P<descricao>.+?)\s+'
    r'(?P<principal>\d{1,3}(?:\.\d{3})*,\d{2})\s+'
    r'(?P<juros>\d{1,3}(?:\.\d{3})*,\d{2}|-)\s+'
    r'(?P<multa>\d{1,3}(?:\.\d{3})*,\d{2}|-)\s+'
    r'(?P<total>\d{1,3}(?:\.\d{3})*,\d{2})'
)

# Data + Banco juntos (ex: "31/01/2024 341 - BANCO ITAU S A")
pattern_data_banco = re.compile(
    r'(\d{2}/\d{2}/\d{4})\s*'
    r'(\d{3}\s*-\s*[A-ZÀ-Ú0-9\.\s]+)'
)

# Termos para ignorar no campo Descrição (rodapé, títulos etc.)
EXCECOES_DESCRICAO = [
    'Agência', 'Estabelecimento', 'Valor Reservado',
    'Restituído', 'Referência'
]

def extrair_dados(pdf_bytes: bytes) -> pd.DataFrame:
    """
    Extrai de cada página do PDF:
      - Banco e Data de Arrecadação (busca no rodapé antes de todo o texto)
      - Todas as linhas DARF válidas (filtra falsas descrições)
    Retorna um DataFrame com colunas:
      Código do DARF, Descrição, Data de Arrecadação, Banco,
      Valor Principal, Juros, Multa e Total.
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    registros = []

    for page in doc:
        texto = page.get_text()
        linhas = texto.splitlines()

        # 1) Monta rodapé com as últimas 8 linhas unidas
        rodape = " ".join(linhas[-8:]).strip()
        m_fb = pattern_data_banco.search(rodape)

        if m_fb:
            data_arrec, banco = m_fb.group(1), m_fb.group(2).strip()
        else:
            # Fallback: texto completo sem quebras
            texto_flat = texto.replace('\n', ' ')
            m2 = pattern_data_banco.search(texto_flat)
            data_arrec, banco = (m2.group(1), m2.group(2).strip()) if m2 else ('', '')

        # 2) Para cada linha DARF detectada
        for m in pattern_linha.finditer(texto):
            descricao = m.group('descricao').strip()

            # Ignora registros cujas descrições contenham termos de exceção
            if any(term.lower() in descricao.lower() for term in EXCECOES_DESCRICAO):
                continue

            registros.append({
                'Código do DARF':        m.group('codigo'),
                'Descrição':             descricao,
                'Data de Arrecadação':   data_arrec,
                'Banco':                 banco,
                'Valor Principal':       m.group('principal').replace('.', '').replace(',', '.'),
                'Juros':                 (m.group('juros').replace('.', '').replace(',', '.') 
                                            if m.group('juros') != '-' else '0.00'),
                'Multa':                 (m.group('multa').replace('.', '').replace(',', '.') 
                                            if m.group('multa') != '-' else '0.00'),
                'Total':                 m.group('total').replace('.', '').replace(',', '.'),
            })

    return pd.DataFrame(registros)

# --- Interface Streamlit ---
st.set_page_config(page_title='Extrator e-CAC', layout='wide')
st.title('Extrator de Comprovantes e-CAC')
st.markdown('Envie seu PDF e extraia automaticamente todos os campos, incluindo Banco e Data de Arrecadação.')

uploaded = st.file_uploader('Selecione o arquivo PDF', type=['pdf'])
if uploaded:
    with st.spinner('Extraindo dados...'):
        df = extrair_dados(uploaded.read())

    if df.empty:
        st.warning('Nenhum dado extraído. Verifique se o layout do PDF corresponde aos padrões esperados.')
    else:
        st.success('Extração concluída com sucesso!')
        st.dataframe(df)

        # Botão de download do Excel
        buffer = BytesIO()
        df.to_excel(buffer, index=False)
        buffer.seek(0)
        st.download_button(
            label='📥 Baixar Excel',
            data=buffer,
            file_name='comprovantes_ecac_extraidos.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
