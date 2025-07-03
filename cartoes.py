import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Consolidador de Extratos", layout="centered")
st.title("🏦 Conversor de Extratos Bancários")
st.markdown("Faça upload do extrato bancário em Excel e baixe o arquivo agrupado e formatado para contabilização.")

uploaded_file = st.file_uploader("📎 Selecione o arquivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        # Lê o arquivo inteiro como texto bruto
        df_raw = pd.read_excel(uploaded_file, header=None, dtype=str, engine='openpyxl')

        # Localiza a linha que contém exatamente 'Deb/Credit'
        linha_cabecalho = None
        for idx, row in df_raw.iterrows():
            if row.astype(str).str.strip().str.lower().isin(['deb/credit']).any():
                linha_cabecalho = idx
                break

        # Se não encontrou, exibe erro e para
        if linha_cabecalho is None:
            st.error("❌ Cabeçalho com 'Deb/Credit' não encontrado no arquivo.")
            st.stop()

        # Lê novamente com a linha correta como cabeçalho
        df = pd.read_excel(uploaded_file, header=linha_cabecalho, dtype=str, engine='openpyxl')
        df.columns = df.columns.str.strip()
        df = df.fillna('')  # Substitui todos os NaN por vazio

        # Filtro por Crédito
        df = df[df['Deb/Credit'] == "Credito"]

        # Filtros relevantes
        historico_filters = [
            'BIN', 'BANRISUL', 'CREDZ', 'ELOSGATE', 'GETNET', 'GLOBAL', 'CIELO', 'REDE',
            'CONTAS A RECEBER TRANSI', 'STONE', 'PAGSEGURO', 'FISERV', 'PAGSEG', 'SISPAG', 'SFPAY'
        ]
        documento_filters = ['12109247', 'FISERV', 'REDE-', 'CIELO']

        df_filtered = df[
            df['Historico'].str.contains('|'.join(historico_filters), na=False) |
            df['Documento'].str.contains('|'.join(documento_filters), na=False)
        ]

        # Remove registros indevidos
        df_filtered = df_filtered[~df_filtered['Historico'].str.contains('MORAIS', na=False)]

        # Limpeza e transformação de dados
        df_filtered['Agencia'] = df_filtered['Agencia'].apply(lambda x: str(x)[-4:] if x else x)
        df_filtered['Conta'] = pd.to_numeric(df_filtered['Conta'], errors='coerce').fillna(0).astype(int).astype(str)
        df_filtered['Filial'] = df_filtered['Filial'].apply(lambda x: str(x)[:4] if x else x)
        df_filtered['Ocorrencia'] = df_filtered['Ocorrencia'].fillna("N/A")
        df_filtered['Data'] = pd.to_datetime(df_filtered['Data'], errors='coerce')
        df_filtered['Valor'] = pd.to_numeric(df_filtered['Valor'], errors='coerce').fillna(0).round(2)

        # Função para identificar a natureza
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

        df_filtered['Historico'] = df_filtered.apply(
            lambda row: get_natureza(row['Historico'], row['Ocorrencia'], row['Documento']), axis=1
        )

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

        # Agrupamento
        df_grouped = df_filtered.groupby(
            ['Filial', 'Data', 'Historico', 'Natureza', 'Banco', 'Agencia', 'Conta']
        ).agg({'Valor': 'sum'}).reset_index()

        # Complementa colunas
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

        # Cria arquivo Excel em memória
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_grouped.to_excel(writer, index=False)
        output.seek(0)

        st.success("✅ Arquivo processado com sucesso!")
        st.download_button(
            label="⬇️ Baixar Excel Formatado",
            data=output,
            file_name=f"consolidado_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Erro ao processar o arquivo: {e}")
