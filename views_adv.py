from flask import render_template, request, jsonify, send_file
from main import app
import pandas as pd
import os
import re

caminho_excel_processos = os.path.join(os.getcwd(), 'Advbox', 'PROCESSOS.xlsx')
caminho_excel_clientes = os.path.join(os.getcwd(), 'Advbox', 'CLIENTES.xlsx')

@app.route('/')
def home():
    return render_template('home.html', titulo='Página inicial - ADVBOX')

@app.route('/processos')
def processos():
    if not os.path.exists(caminho_excel_processos):
        return "Arquivo não encontrado!"

    df_excel_processos = pd.read_excel(caminho_excel_processos)
    df_excel_processos = df_excel_processos.fillna("")
    html_tabela = df_excel_processos.to_html(index=False, border=1, classes='table table-striped')

    caminho_html = os.path.join(os.getcwd(), 'templates', 'tabela_processos.html')
    with open(caminho_html, "w", encoding="utf-8") as arquivo_html:
        arquivo_html.write(html_tabela)
    return render_template('processos.html', titulo='Processos - ADVBOX')

@app.route('/download_tabela_processos')
def download_tabela_processos():
    return send_file(caminho_excel_processos, as_attachment=True)

@app.route('/clientes')
def clientes():
    if not os.path.exists(caminho_excel_clientes):
        return "Arquivo Excel não encontrado!", 404

    df_excel_clientes = pd.read_excel(caminho_excel_clientes)
    df_excel_clientes = df_excel_clientes.fillna("")
    html_tabela = df_excel_clientes.to_html(index=False, border=1, classes='table table-striped')

    caminho_html = os.path.join(os.getcwd(), 'templates', 'tabela_clientes.html')
    with open(caminho_html, "w", encoding="utf-8") as arquivo_html:
        arquivo_html.write(html_tabela)
    return render_template('clientes.html', titulo='Clientes - ADVBOX')

@app.route('/download_tabela_clientes')
def download_tabela_clientes():
    return send_file(caminho_excel_clientes, as_attachment=True)

@app.route('/upload_file', methods=['POST',])
def upload_file():
    if 'files' not in request.files:
        return jsonify({"error": "Nenhum arquivo enviado."}), 400

    files = request.files.getlist('files')

    #Inclusão dos dados do csv nas tabelas
    for file in files:
        try:
            #Tratamento específico para cada arquivo
            nome_arquivo = os.path.basename(file.filename)
            match nome_arquivo:
                case "v_clientes_CodEmpresa_92577.csv":
                    #Carrega os dataframes que serão trabalhados
                    df_csv_clientes = pd.read_csv(file.stream, encoding="MacRoman", delimiter=';')
                    df_clientes = pd.read_excel(caminho_excel_clientes, sheet_name="Página1")

                    #Armazena apenas os dados que vão ser migrados para a planilha CLIENTES da ADVBOX
                    df_dados_migrados = df_csv_clientes[['razao_social', 'cpf', 'rg', 'nacionalidade', 'nascimento', 'estado_civil', 'profissao', 'telefone2', 'telefone1', 'email1', 'uf', 'cidade', 'bairro', 'logradouro', 'cep', 'pis', 'nome_mae', 'cnpj']]
                    df_dados_migrados.columns = ['NOME', 'CPF CNPJ', 'RG',  'NACIONALIDADE', 'DATA DE NASCIMENTO', 'ESTADO CIVIL', 'PROFISSÃO', 'CELULAR', 'TELEFONE', 'EMAIL', 'ESTADO', 'CIDADE', 'BAIRRO', 'ENDEREÇO', 'CEP', 'PIS PASEP', 'NOME DA MÃE', 'ANOTAÇÕES GERAIS']

                    #Validação para conferir se existe dados repetidos
                    validacao_dados_existentes = df_clientes[['NOME']]
                    validacao_dados_migrados = df_dados_migrados[['NOME']]
                    dados_repetidos = validacao_dados_migrados.merge(validacao_dados_existentes, on=['NOME'], how='inner')
                    if not dados_repetidos.empty:
                        return jsonify({"message": "Erro: Alguns dados já existem no arquivo. Não foi possível adicionar os dados repetidos."})
                    else:
                        # Corrigir as colunas que tem encode errado
                        for col in ['NOME', 'CPF CNPJ', 'RG',  'NACIONALIDADE', 'DATA DE NASCIMENTO', 'ESTADO CIVIL', 'PROFISSÃO', 'CELULAR', 'TELEFONE', 'EMAIL', 'ESTADO', 'CIDADE', 'BAIRRO', 'ENDEREÇO', 'CEP', 'PIS PASEP', 'NOME DA MÃE', 'ANOTAÇÕES GERAIS']:
                            df_dados_migrados[col] = df_dados_migrados[col].astype(str).str.encode('MacRoman').str.decode('utf-8')

                        #Adicionar os novos dados na planilha CLIENTES da ADVBOX
                        df_clientes = pd.concat([df_clientes, df_dados_migrados], ignore_index=True)
                        df_clientes.to_excel(caminho_excel_clientes, sheet_name='Página1', index=False)

                case "v_processos_CodEmpresa_92577.csv":
                    #Carrega os dataframes que serão trabalhados
                    df_csv_processos = pd.read_csv(file.stream, encoding="utf-8", delimiter=';')
                    df_processos = pd.read_excel(caminho_excel_processos, sheet_name="Página1")

                    #Armazena apenas os dados que vão ser migrados para a planilha CLIENTES da ADVBOX
                    df_dados_migrados = df_csv_processos[['cod_cliente', 'objeto_acao', 'codigo_fase', 'numero_processo', 'cod_processo_apensar', 'codlocaltramite', 'codcomarca', 'valor_causa', 'inclusao', 'data_contratacao', 'data_transitojulgado', 'data_encerramento', 'data_distribuicao', 'cod_usuario', 'pedido']]
                    df_dados_migrados.columns = ['NOME DO CLIENTE', 'TIPO DE AÇÃO',  'FASE PROCESSUAL', 'NÚMERO DO PROCESSO', 'PROCESSO ORIGINÁRIO', 'TRIBUNAL', 'COMARCA', 'EXPECTATIVA/VALOR DA CAUSA', 'DATA CADASTRO', 'DATA FECHAMENTO', 'DATA TRANSITO', 'DATA ARQUIVAMENTO', 'DATA REQUERIMENTO', 'RESPONSÁVEL', 'ANOTAÇÕES GERAIS']

                    #Validação para conferir se existe dados repetidos
                    validacao_dados_existentes = df_processos[['NÚMERO DO PROCESSO']]
                    validacao_dados_migrados = df_dados_migrados[['NÚMERO DO PROCESSO']]
                    dados_repetidos = validacao_dados_migrados.merge(validacao_dados_existentes, on=['NÚMERO DO PROCESSO'], how='inner')
                    if not dados_repetidos.empty:
                        return jsonify({"message": "Erro: Alguns dados já existem no arquivo. Não foi possível adicionar os dados repetidos."})
                    else:
                        #Adicionar os novos dados na planilha PROCESSOS da ADVBOX
                        df_processos = pd.concat([df_processos, df_dados_migrados], ignore_index=True)
                        df_processos.to_excel(caminho_excel_processos, sheet_name='Página1', index=False)

                case _:
                    print('Adicionar arquivos futuramente na etapa de migração')
        except Exception as e:
            return jsonify({"error": f"Erro ao processar o arquivo {file.filename}: {e}"}), 400

    #Tratamento da tabela utilizando dicionários
    for file in files:
        try:
            #Tratamento específico para cada arquivo
            nome_arquivo = os.path.basename(file.filename)
            match nome_arquivo:
                case "v_cliente_estado_civil_CodEmpresa_92577.csv":
                    #Carrega os dataframes que serão trabalhados
                    df_csv_dicionario_clientes = pd.read_csv(file.stream, encoding="utf-8", delimiter=';')
                    df_clientes = pd.read_excel(caminho_excel_clientes, sheet_name="Página1")

                    #Utiliza o dataframe dicionário clientes para alterar os dados errados
                    df_csv_dicionario_clientes['descricao'] = df_csv_dicionario_clientes['descricao'].str.replace(r'\(.*?\)', '', regex=True).str.strip().str.upper()
                    mapeamento_estado_civil = dict(zip(df_csv_dicionario_clientes['sigla'], df_csv_dicionario_clientes['descricao']))

                    df_clientes['ESTADO CIVIL'] = df_clientes['ESTADO CIVIL'].map(mapeamento_estado_civil).fillna(df_clientes['ESTADO CIVIL'])
                    df_clientes.to_excel(caminho_excel_clientes, sheet_name='Página1', index=False)

                case "v_objeto_acao_CodEmpresa_92577.csv":
                    # Carrega os dataframes que serão trabalhados
                    df_csv_dicionario_acao = pd.read_csv(file.stream, encoding="utf-8", delimiter=';')
                    df_processos = pd.read_excel(caminho_excel_processos, sheet_name="Página1")

                    # Utiliza o dataframe dicionário acao para alterar os códigos
                    mapeamento_objeto_acao = dict(zip(df_csv_dicionario_acao['codigo'], df_csv_dicionario_acao['descricao']))

                    df_processos['TIPO DE AÇÃO'] = df_processos['TIPO DE AÇÃO'].map(mapeamento_objeto_acao).fillna(df_processos['TIPO DE AÇÃO']).str.upper()
                    df_processos.to_excel(caminho_excel_processos, sheet_name='Página1', index=False)

                case  "v_fase_CodEmpresa_92577.csv":
                    # Carrega os dataframes que serão trabalhados
                    df_csv_dicionario_fase = pd.read_csv(file.stream, encoding="utf-8", delimiter=';')
                    df_processos = pd.read_excel(caminho_excel_processos, sheet_name="Página1")

                    # Utiliza o dataframe dicionário fase para alterar os códigos
                    mapeamento_objeto_fase = dict(zip(df_csv_dicionario_fase['codigo'], df_csv_dicionario_fase['fase']))

                    df_processos['FASE PROCESSUAL'] = df_processos['FASE PROCESSUAL'].map(mapeamento_objeto_fase).fillna(df_processos['FASE PROCESSUAL']).str.upper()
                    df_processos.to_excel(caminho_excel_processos, sheet_name='Página1', index=False)

                case  "v_local_tramite_CodEmpresa_92577.csv":
                    #Carrega os dataframes que serão trabalhados
                    df_csv_dicionario_tribunal = pd.read_csv(file.stream, encoding="utf-8", delimiter=';')
                    df_processos = pd.read_excel(caminho_excel_processos, sheet_name="Página1")

                    # Utiliza o dataframe dicionário tribunal para alterar os códigos
                    mapeamento_objeto_tribunal = dict(zip(df_csv_dicionario_tribunal['codigo'], df_csv_dicionario_tribunal['descricao']))

                    df_processos['TRIBUNAL'] = df_processos['TRIBUNAL'].map(mapeamento_objeto_tribunal).fillna(df_processos['TRIBUNAL']).str.upper()
                    df_processos.to_excel(caminho_excel_processos, sheet_name='Página1', index=False)

                case  "v_comarca_CodEmpresa_92577.csv":
                    # Carrega os dataframes que serão trabalhados
                    df_csv_dicionario_comarca = pd.read_csv(file.stream, encoding="utf-8", delimiter=';')
                    df_processos = pd.read_excel(caminho_excel_processos, sheet_name="Página1")

                    # Utiliza o dataframe dicionário comarca para alterar os códigos
                    mapeamento_objeto_comarca = dict(zip(df_csv_dicionario_comarca['codigo'], df_csv_dicionario_comarca['descricao']))

                    df_processos['COMARCA'] = df_processos['COMARCA'].map(mapeamento_objeto_comarca).fillna(df_processos['COMARCA']).str.upper()
                    df_processos.to_excel(caminho_excel_processos, sheet_name='Página1', index=False)

                case  "v_usuario_CodEmpresa_92577.csv":
                    # Carrega os dataframes que serão trabalhados
                    df_csv_dicionario_responsavel = pd.read_csv(file.stream, encoding="ascii", delimiter=';')
                    df_processos = pd.read_excel(caminho_excel_processos, sheet_name="Página1")

                    # Utiliza o dataframe dicionário responsavel para alterar os códigos
                    mapeamento_objeto_responsavel = dict(zip(df_csv_dicionario_responsavel['id'], df_csv_dicionario_responsavel['nome']))

                    df_processos['RESPONSÁVEL'] = df_processos['RESPONSÁVEL'].map(mapeamento_objeto_responsavel).fillna(df_processos['RESPONSÁVEL']).str.upper()
                    df_processos.to_excel(caminho_excel_processos, sheet_name='Página1', index=False)

                case _:
                    print('Adicionar arquivos futuramente na etapa de tratamento')
        except Exception as e:
            return jsonify({"error": f"Erro ao processar o arquivo {file.filename}: {e}"}), 400

    ###
    ###padronização geral do arquivo CLIENTES para se adequar a migração
    ###
    df_clientes = pd.read_excel(caminho_excel_clientes, sheet_name="Página1")

    #Padronizar nome todos em maiusculo
    df_clientes['NOME'] = df_clientes['NOME'].str.upper()

    #Padronizar primeiro remove todos os caracteres especiais e depois, os campos não nulos vão ser preenchidos com 0 ou cortados com até 14 caracteres
    df_clientes['CPF CNPJ'] = df_clientes['CPF CNPJ'].apply(lambda x: re.sub(r'\D', '', str(x)))
    df_clientes['CPF CNPJ'] = df_clientes['CPF CNPJ'].apply(lambda x: x.zfill(11) if len(x) < 11 else x[:14])

    #Padronizar Data de nascimento no padrão dd/mm/yyyy, cortando as horas
    df_clientes['DATA DE NASCIMENTO'] = pd.to_datetime(df_clientes['DATA DE NASCIMENTO'], errors='coerce', dayfirst=True).dt.strftime('%d/%m/%Y')

    #Padronizar nacionalidade toda em maiusculo
    df_clientes['NACIONALIDADE'] = df_clientes['NACIONALIDADE'].str.upper()

    #Padronizar profissao toda em maiusculo
    df_clientes['PROFISSÃO'] = df_clientes['PROFISSÃO'].str.upper()

    #Padronizar celular e telefone com 10 ou 11 digitos
    def padronizar_celtel(celtel):
        celtel = re.sub(r'\D', '', str(celtel))
        if len(celtel) == 10:
            celtel = '0' + celtel
        elif len(celtel) < 10:
            celtel = celtel.zfill(11)
        mascara_celtel = f"({celtel[:2]}) {celtel[2:7]}-{celtel[7:]}"
        return mascara_celtel
    df_clientes['CELULAR'] = df_clientes['CELULAR'].apply(padronizar_celtel)
    df_clientes['TELEFONE'] = df_clientes['TELEFONE'].apply(padronizar_celtel)

    #Padronizar CEP
    def padronizar_cep(cep):
        cep = re.sub(r'\D', '', str(cep))
        if len(cep) > 8:
            cep = cep[:8]
        elif len(cep) < 8:
            cep = cep.ljust(8, '0')
        cep_formatado = f"{cep[:5]}-{cep[5:]}"
        return cep_formatado
    df_clientes['CEP'] = df_clientes['CEP'].apply(padronizar_cep)

    #Padronizar PIS PASEP
    df_clientes['PIS PASEP'] = df_clientes['PIS PASEP'].fillna('').astype(str)
    df_clientes['PIS PASEP'] = df_clientes['PIS PASEP'].apply(lambda x: f"{x[:3]}.{x[3:7]}.{x[7:10]}-{x[10:]}" if len(x) >= 11 else x)

    #Padronizar origem do cliente para ser preenchido como MIGRAÇÃO os dados que estão vazios
    df_clientes['ORIGEM DO CLIENTE'] = df_clientes['ORIGEM DO CLIENTE'].fillna('MIGRAÇÃO')

    #Padronizar ANOTACOES GERAIS que ficou com os números de CNPJ
    df_clientes['ANOTAÇÕES GERAIS'] = df_clientes['ANOTAÇÕES GERAIS'].apply(lambda x: re.sub(r'\D', '', str(x)))
    df_clientes['ANOTAÇÕES GERAIS'] = df_clientes['ANOTAÇÕES GERAIS'].apply(lambda x: x.zfill(14))

    #Salvar o arquivo CLIENTES.xlsx
    df_clientes.to_excel(caminho_excel_clientes, sheet_name='Página1', index=False)

    ###
    ###padronização geral do arquivo PROCESSOS para se adequar a migração
    ###
    df_processos = pd.read_excel(caminho_excel_processos, sheet_name="Página1")
    #Deletar as duas primeiras linhas, que eram de instruções
    df_processos = df_processos.drop(index=[0, 1])

    #Padronizar a coluna NUMERO DO PROCESSO e transferir os diferentes para a coluna PROTOCOLO
    mascara_numero_processo = r'^\d{7}-\d{2}\.\d{4}\.\d{1}\.\d{2}\.\d{4}$'
    df_processos['PROTOCOLO'] = df_processos['NÚMERO DO PROCESSO'].where(~df_processos['NÚMERO DO PROCESSO'].str.match(mascara_numero_processo), df_processos['PROTOCOLO'])
    df_processos['NÚMERO DO PROCESSO'] = df_processos['NÚMERO DO PROCESSO'].where(df_processos['NÚMERO DO PROCESSO'].str.match(mascara_numero_processo), "")

    #Padronização das datas dos PROCESSOS em dd/mm/yyyy
    df_processos['DATA CADASTRO'] = pd.to_datetime(df_processos['DATA CADASTRO'], errors='coerce', dayfirst=True).dt.strftime('%d/%m/%Y')
    df_processos['DATA FECHAMENTO'] = pd.to_datetime(df_processos['DATA FECHAMENTO'], errors='coerce', dayfirst=True).dt.strftime('%d/%m/%Y')
    df_processos['DATA TRANSITO'] = pd.to_datetime(df_processos['DATA TRANSITO'], errors='coerce', dayfirst=True).dt.strftime('%d/%m/%Y')
    df_processos['DATA ARQUIVAMENTO'] = pd.to_datetime(df_processos['DATA ARQUIVAMENTO'], errors='coerce', dayfirst=True).dt.strftime('%d/%m/%Y')
    df_processos['DATA REQUERIMENTO'] = pd.to_datetime(df_processos['DATA REQUERIMENTO'], errors='coerce', dayfirst=True).dt.strftime('%d/%m/%Y')

    #Salvar o arquivo PROCESSOS.xlsx
    df_processos.to_excel(caminho_excel_processos, sheet_name='Página1', index=False)

    #Retorna uma mensagem de sucesso com os arquivos processados
    return jsonify({"message": "Arquivos enviados e processados com sucesso."})