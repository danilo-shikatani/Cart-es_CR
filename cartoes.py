import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Consolidador de Extratos", layout="centered")
st.title("🏦 Conversor de Extratos Bancários")
st.markdown("Faça upload do extrato bancário em Excel e baixe o arquivo agrupado e formatado para contabilização.")

uploaded_file = st.file_uploader("📎 Selecione o arquivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    colunas_com_zeros_a_esquerda = ['Banco', 'Agencia', 'Conta']

    try:
        df = pd.read_excel(uploaded_file, dtype={col: str for col in colunas_com_zeros_a_esquerda}, skiprows=1, engine='openpyxl')
        df.columns = df.columns.str.strip()
        df.columns = df.iloc[0]  # Define a primeira linha como cabeçalho
        df = df[1:]  # Remove a primeira linha repetida
        df = df[df['Deb/Credit'] == "Credito"]

        historico_filters = ['BIN', 'BANRISUL', 'CREDZ', 'ELOSGATE', 'GETNET', 'GLOBAL', 'CIELO', 'REDE', 'CONTAS A RECEBER TRANSI', 'STONE', 'PAGSEGURO', 'FISERV', 'PAGSEG', 'SISPAG', 'SFPAY']
        documento_filters = ['12109247', 'FISERV', 'REDE-', 'CIELO']

        df_filtered = df[
            df['Historico'].str.contains('|'.join(historico_filters), na=False) |
            df['Documento'].str.contains('|'.join(documento_filters), na=False)
        ]

        df_filtered = df_filtered[~df_filtered['Historico'].str.contains('MORAIS', na=False)]
        df_filtered['Agencia'] = df_filtered['Agencia'].apply(lambda x: str(x)[-4:] if pd.notnull(x) else x)
        df_filtered['Ocorrencia'] = df_filtered['Ocorrencia'].fillna("N/A")
        df_filtered['Conta'] = pd.to_numeric(df_filtered['Conta'], errors='coerce').astype('Int64')
        df_filtered['Data'] = pd.to_datetime(df_filtered['Data'], errors='coerce')
        df_filtered['Filial'] = df_filtered['Filial'].apply(lambda x: str(x)[:4] if pd.notnull(x) else x)
        df_filtered['Valor'] = pd.to_numeric(df_filtered['Valor'], errors='coerce')

        def get_natureza(historico, ocorrencia, documento):
            if 'BANRISUL' in historico: return 'BANRISUL'
            elif 'BIN' in historico: return 'BIN'
            elif 'CREDZ' in historico or '12109247000120' in documento: return 'CREDZ'
            elif 'GETNET' in historico: return 'GETNET'
            elif 'GLOBAL' in historico: return 'GLOBAL'
            elif 'CIELO' in historico or 'CIELO' in documento: return 'CIELO'
            elif 'REDE' in historico or 'REDE' in documento: return 'REDE'
            elif 'VERO' in ocorrencia: return 'BIN'
            elif 'PAGSEGURO' in historico: return 'PAGSEGURO'
            elif 'PAGSEG' in historico: return 'TEDPAGSEG'
            elif 'FISERV' in historico or 'FISERV' in documento: return 'BIN'
            elif 'SISPAG' in historico: return 'SISPAG PAGSEG'
            elif 'SFPAY' in historico: return 'SFPAY'
            return None

        df_filtered['Historico'] = df_filtered.apply(lambda row: get_natureza(row['Historico'], row['Ocorrencia'], row['Documento']), axis=1)

        natureza_map = {
            'BANRISUL': 'A10801',
            'BIN': 101113,
            'CREDZ': 101115,
            'GETNET': 101112,
            'GLOBAL': 'A10806',
            'CIELO': 101118,
            'REDE': 101111,
            'TEDPAGSEG': 101117,
            'SFPAY': 101119,
            'PAGSEGURO': 101117,
            'SISPAG PAGSEG': 101117
        }

        df_filtered['Natureza'] = df_filtered['Historico'].map(natureza_map)
        df_filtered['Valor'] = df_filtered['Valor'].round(2)

        df_grouped = df_filtered.groupby(['Filial', 'Data', 'Historico', 'Natureza', 'Banco', 'Agencia', 'Conta']).agg({'Valor': 'sum'}).reset_index()
        df_grouped['TIPO'] = 'R'
        df_grouped['NUMERARIO'] = 'CD'
        df_grouped['NUM CHEQUE'] = ''
        df_grouped['C. Custo debito'] = ''
        df_grouped['C. Custo credito'] = ''
        df_grouped['Item debito'] = ''
        df_grouped['Item credito'] = ''
        df_grouped['Cl Valor deb'] = ''
        df_grouped['Cl Valor crd'] = ''

        colunas_ordenadas = [
            'Filial', 'Data', 'NUMERARIO', 'TIPO', 'Valor', 'Natureza', 'Banco', 'Agencia', 'Conta',
            'NUM CHEQUE', 'Historico', 'C. Custo debito', 'C. Custo credito',
            'Item debito', 'Item credito', 'Cl Valor deb', 'Cl Valor crd'
        ]
        df_grouped = df_grouped[colunas_ordenadas]

        st.success("✅ Arquivo processado com sucesso!")
        st.download_button(
            label="⬇️ Baixar Excel Formatado",
            data=df_grouped.to_excel(index=False, engine='openpyxl'),
            file_name=f"consolidado_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Erro ao processar o arquivo: {e}")
