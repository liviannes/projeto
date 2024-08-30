{\rtf1\ansi\ansicpg1252\cocoartf2636
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\paperw11900\paperh16840\margl1440\margr1440\vieww11520\viewh8400\viewkind0
\pard\tx566\tx1133\tx1700\tx2267\tx2834\tx3401\tx3968\tx4535\tx5102\tx5669\tx6236\tx6803\pardirnatural\partightenfactor0

\f0\fs24 \cf0 import streamlit as st\
import pandas as pd\
\
# T\'edtulo da aplica\'e7\'e3o\
st.title("Processador de Planilhas Excel")\
\
# Interface para upload do arquivo\
uploaded_file = st.file_uploader("Escolha um arquivo Excel", type="xlsx")\
\
if uploaded_file is not None:\
    # Carregar o arquivo Excel\
    df = pd.read_excel(uploaded_file, header=None)\
    st.write("Arquivo carregado com sucesso!")\
\
    # Inicializar vari\'e1veis para armazenar os resultados\
    all_results = []\
\
    # Vari\'e1veis de controle para identificar os blocos\
    current_activity = None\
    is_collecting = False\
    columns = ['Activity', 'Start Date', 'Duration']  # Colunas finais desejadas\
\
    # Percorrer todas as linhas para processar os blocos\
    for index, row in df.iterrows():\
        if pd.notnull(row[0]):\
            if '[' in str(row[0]):  # Se for o in\'edcio de um novo bloco (nome da atividade)\
                if is_collecting and block_data:\
                    # Finalizar o bloco anterior se houver\
                    temp_df = pd.DataFrame(block_data, columns=columns)\
                    all_results.append(temp_df)\
                    st.write(f'Bloco finalizado: \{current_activity\} - \{len(temp_df)\} linhas capturadas')\
\
                current_activity = row[0]\
                block_data = []  # Reiniciar a coleta de dados para o novo bloco\
                is_collecting = True\
                st.write(f'Iniciando bloco: \{current_activity\}')\
            \
            elif is_collecting and row[0] == 'Started By':  # Linha que define as colunas\
                continue  # Pula a linha das colunas\
            \
            elif is_collecting and 'Total' in str(row[0]):  # Finalizar o bloco ao encontrar "Total"\
                temp_df = pd.DataFrame(block_data, columns=columns)\
                all_results.append(temp_df)\
                st.write(f'Bloco finalizado: \{current_activity\} - \{len(temp_df)\} linhas capturadas')\
                block_data = []  # Reiniciar a coleta de dados para o pr\'f3ximo bloco\
            \
            elif is_collecting:  # Coletar todas as linhas at\'e9 encontrar "Total"\
                try:\
                    start_date = row[2]  # Considerando que Start Date est\'e1 na coluna 2\
                    duration = row[6]  # Considerando que Duration est\'e1 na coluna 6\
                    block_data.append([current_activity, start_date, duration])\
                    st.write(f'Dados coletados: Atividade: \{current_activity\}, In\'edcio: \{start_date\}, Dura\'e7\'e3o: \{duration\}')\
                except KeyError as e:\
                    st.write(f"Erro ao acessar coluna: \{e\}. Verifique se a coluna existe.")\
                    continue\
\
    # Ap\'f3s percorrer todas as linhas, certifique-se de que o \'faltimo bloco seja capturado\
    if block_data:\
        temp_df = pd.DataFrame(block_data, columns=columns)\
        all_results.append(temp_df)\
        st.write(f'Bloco finalizado: \{current_activity\} - \{len(temp_df)\} linhas capturadas')\
\
    # Combinar todos os blocos processados em um \'fanico DataFrame\
    final_df = pd.concat(all_results, ignore_index=True)\
\
    # Remover linhas vazias\
    final_df.dropna(inplace=True)\
\
    # Exibir uma pr\'e9via do DataFrame processado\
    st.write("Dados processados:")\
    st.write(final_df.head(20))  # Exibir mais linhas para verifica\'e7\'e3o completa\
\
    # Widgets para selecionar as datas de in\'edcio e fim\
    inicio_date = st.date_input("Data de In\'edcio")\
    fim_date = st.date_input("Data de Fim")\
\
    # Bot\'e3o para processar e filtrar os dados\
    if st.button("Filtrar e Salvar"):\
        # Converter a coluna 'Start Date' para datetime para facilitar a filtragem\
        final_df['Start Date'] = pd.to_datetime(final_df['Start Date'])\
\
        # Filtrar pelo intervalo de datas selecionado\
        filtered_df = final_df[(final_df['Start Date'] >= pd.to_datetime(inicio_date)) & \
                               (final_df['Start Date'] <= pd.to_datetime(fim_date))]\
\
        # Exibir o resultado filtrado\
        st.write("Dados filtrados:")\
        st.write(filtered_df)\
\
        # Salvar o resultado em um novo arquivo Excel\
        filtered_df.to_excel('resultado_filtrado_por_datas.xlsx', index=False)\
        st.write("Arquivo 'resultado_filtrado_por_datas.xlsx' salvo com sucesso!")\
\
        # Disponibilizar o arquivo para download\
        st.download_button(\
            label="Baixar Resultado Filtrado",\
            data=open('resultado_filtrado_por_datas.xlsx', 'rb').read(),\
            file_name='resultado_filtrado_por_datas.xlsx',\
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'\
        )\
}