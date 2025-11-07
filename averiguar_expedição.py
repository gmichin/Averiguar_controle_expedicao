import pandas as pd
import numpy as np
import os
from pathlib import Path
from datetime import datetime

def limpar_valor_monetario(valor):
    """
    Limpa valores monetários removendo pontos e convertendo vírgulas para pontos
    """
    if pd.isna(valor) or valor == '':
        return 0.0
    
    # Converte para string se não for
    valor_str = str(valor).strip()
    
    # Remove R$, espaços e outros caracteres
    valor_str = valor_str.replace('R$', '').replace(' ', '').strip()
    
    # Verifica se tem múltiplas vírgulas (possível erro de digitação)
    if valor_str.count(',') > 1:
        # Caso como "1,564,100" - provavelmente é 1564.10
        # Remove todas as vírgulas exceto a última
        partes = valor_str.split(',')
        parte_inteira = ''.join(partes[:-1])  # Junta todas as partes exceto a última
        parte_decimal = partes[-1]
        valor_limpo = parte_inteira + '.' + parte_decimal
    elif ',' in valor_str and '.' in valor_str:
        # Remove pontos dos milhares e converte vírgula para ponto
        partes = valor_str.split(',')
        parte_inteira = partes[0].replace('.', '')
        valor_limpo = parte_inteira + '.' + partes[1]
    elif ',' in valor_str:
        # Verifica se a vírgula é separador decimal ou de milhar
        if valor_str.count(',') == 1 and len(valor_str.split(',')[1]) == 2:
            # Provavelmente é separador decimal (ex: 1124,00)
            valor_limpo = valor_str.replace(',', '.')
        else:
            # Provavelmente é separador de milhar (ex: 1,564)
            valor_limpo = valor_str.replace(',', '')
    else:
        valor_limpo = valor_str
    
    try:
        resultado = float(valor_limpo)
        # Verifica se o valor é muito grande (possível erro)
        if resultado > 1000000:  # Se maior que 1 milhão, provavelmente tem erro
            # Tenta interpretar como número com 2 casas decimais
            valor_str_sem_virgulas = str(valor).replace(',', '').replace('.', '')
            if len(valor_str_sem_virgulas) > 2:
                resultado = float(valor_str_sem_virgulas[:-2] + '.' + valor_str_sem_virgulas[-2:])
        return resultado
    except:
        print(f"Erro ao converter valor: {valor} -> {valor_limpo}")
        return 0.0

def limpar_valor_numerico(valor):
    """
    Limpa valores numéricos (peso, quantidade, etc.)
    """
    if pd.isna(valor) or valor == '':
        return 0.0
    
    valor_str = str(valor).strip()
    
    # Remove caracteres não numéricos exceto ponto e vírgula
    valor_str = ''.join(c for c in valor_str if c.isdigit() or c in ',.')
    
    # Verifica se tem múltiplas vírgulas (possível erro de digitação)
    if valor_str.count(',') > 1:
        # Caso como "1,564,100" - provavelmente é 1564.10
        partes = valor_str.split(',')
        parte_inteira = ''.join(partes[:-1])  # Junta todas as partes exceto a última
        parte_decimal = partes[-1]
        valor_limpo = parte_inteira + '.' + parte_decimal
    elif ',' in valor_str and '.' in valor_str:
        # Se tem ambos, remove pontos de milhar e converte vírgula decimal
        partes = valor_str.split(',')
        parte_inteira = partes[0].replace('.', '')
        valor_limpo = parte_inteira + '.' + partes[1]
    elif ',' in valor_str:
        # Verifica se é decimal
        if valor_str.count(',') == 1 and len(valor_str.split(',')[1]) <= 3:
            # Provavelmente é separador decimal
            valor_limpo = valor_str.replace(',', '.')
        else:
            # Remove vírgulas (provavelmente são separadores de milhar)
            valor_limpo = valor_str.replace(',', '')
    else:
        valor_limpo = valor_str
    
    try:
        resultado = float(valor_limpo)
        # Verifica se o valor é muito grande para peso (possível erro)
        if resultado > 100000:  # Se maior que 100 toneladas, provavelmente tem erro
            # Tenta interpretar como número com 2 casas decimais
            valor_str_sem_virgulas = str(valor).replace(',', '').replace('.', '')
            if len(valor_str_sem_virgulas) > 2:
                resultado = float(valor_str_sem_virgulas[:-2] + '.' + valor_str_sem_virgulas[-2:])
        return resultado
    except:
        print(f"Erro ao converter valor numérico: {valor} -> {valor_limpo}")
        return 0.0

def detectar_separador_csv(caminho):
    """
    Detecta o separador do arquivo CSV
    """
    try:
        with open(caminho, 'r', encoding='utf-8') as f:
            primeira_linha = f.readline()
        if ';' in primeira_linha:
            return ';'
        elif ',' in primeira_linha:
            return ','
        else:
            return ';'  # padrão para CSV brasileiro
    except:
        return ';'

def limpar_numero_nota_fiscal(valor):
    """
    Limpa o número da nota fiscal mantendo zeros à esquerda e convertendo para string
    """
    if pd.isna(valor) or valor == '':
        return ''
    
    # Converte para string
    valor_str = str(valor).strip()
    
    # Remove todos os caracteres não numéricos
    valor_limpo = ''.join(c for c in valor_str if c.isdigit())
    
    return valor_limpo

def formatar_nota_fiscal(valor):
    """
    Formata a nota fiscal corretamente, tratando o problema do zero extra
    """
    if pd.isna(valor) or valor == '':
        return ''
    
    # Converte para string
    valor_str = str(valor).strip()
    
    # Remove todos os caracteres não numéricos
    valor_limpo = ''.join(c for c in valor_str if c.isdigit())
    
    # CORREÇÃO: Remove zero extra no final se a NF tiver 7 dígitos
    # Notas fiscais normalmente têm 6 dígitos (112624), mas podem ter zeros extras
    if len(valor_limpo) == 7 and valor_limpo.endswith('0'):
        # Remove o último zero se a NF tiver 7 dígitos e terminar com zero
        valor_limpo = valor_limpo[:-1]
    
    return valor_limpo

def converter_para_inteiro_nota_fiscal(valor):
    """
    Converte a nota fiscal para inteiro (remove zeros à esquerda)
    """
    if pd.isna(valor) or valor == '':
        return 0
    
    # Primeiro formata corretamente
    valor_formatado = formatar_nota_fiscal(valor)
    
    # Converte para inteiro (isso remove zeros à esquerda)
    try:
        return int(valor_formatado) if valor_formatado else 0
    except:
        return 0

def ler_csv_com_cabecalho(caminho, data_inicio=None, data_fim=None):
    """
    Lê o CSV detectando automaticamente o cabeçalho e filtrando por data se especificado
    """
    separador = detectar_separador_csv(caminho)
    
    # Primeiro, vamos ler apenas o cabeçalho para ver as colunas
    try:
        df_teste = pd.read_csv(caminho, nrows=0, sep=separador, encoding='utf-8')
        colunas_disponiveis = df_teste.columns.tolist()
        print(f"Colunas disponíveis no CSV: {colunas_disponiveis}")
        
        # Mapeia possíveis nomes de colunas
        mapeamento_colunas = {
            'NOTA FISCAL': ['NOTA FISCAL', 'NOTAFISCAL', 'NOTA_FISCAL', 'NF', 'NOTA'],
            'PESO': ['PESO', 'PESO_KG', 'PESO KG', 'PESO_TOTAL'],
            'TOTAL': ['TOTAL', 'TOTAL_NF', 'VALOR_TOTAL', 'VALOR TOTAL', 'VALOR'],
            'DATA': ['DATA', 'DATE', 'DT', 'DATA_NF', 'DATA EMISSÃO', 'EMISSÃO']
        }
        
        colunas_para_ler = []
        for coluna_base, alternativas in mapeamento_colunas.items():
            encontrou = False
            for coluna_csv in colunas_disponiveis:
                if any(alt in coluna_csv.upper() for alt in alternativas):
                    colunas_para_ler.append(coluna_csv)
                    print(f"Usando coluna '{coluna_csv}' para {coluna_base}")
                    encontrou = True
                    break
            if not encontrou and coluna_base == 'DATA' and (data_inicio or data_fim):
                print(f"AVISO: Coluna DATA não encontrada. Não será possível filtrar por período.")
                print(f"Alternativas procuradas: {alternativas}")
        
        # Agora lê o arquivo completo com as colunas detectadas
        df = pd.read_csv(caminho, sep=separador, usecols=colunas_para_ler, encoding='utf-8')
        
        # Renomeia as colunas para nomes padrão
        rename_dict = {}
        for coluna_csv in colunas_para_ler:
            for coluna_base, alternativas in mapeamento_colunas.items():
                if any(alt in coluna_csv.upper() for alt in alternativas):
                    rename_dict[coluna_csv] = coluna_base
        
        df = df.rename(columns=rename_dict)
        
        # Filtra por data se especificado e se a coluna DATA existe
        if (data_inicio or data_fim) and 'DATA' in df.columns:
            print("Filtrando CSV por período de datas...")
            
            # Converte a coluna DATA para datetime
            try:
                df['DATA'] = pd.to_datetime(df['DATA'], dayfirst=True, errors='coerce')
                
                # Aplica os filtros de data
                if data_inicio:
                    df = df[df['DATA'] >= data_inicio]
                if data_fim:
                    # Adiciona 1 dia para incluir a data final completa
                    data_fim_ajustada = data_fim + pd.Timedelta(days=1)
                    df = df[df['DATA'] < data_fim_ajustada]
                
                print(f"Período filtrado no CSV: {data_inicio.strftime('%d/%m/%Y') if data_inicio else 'Início indefinido'} a {data_fim.strftime('%d/%m/%Y') if data_fim else 'Fim indefinido'}")
                print(f"Registros no CSV após filtro de data: {len(df)}")
                
            except Exception as e:
                print(f"Erro ao filtrar CSV por data: {e}")
                print("Continuando sem filtro de data no CSV...")
        
        return df
        
    except Exception as e:
        print(f"Erro ao ler CSV: {e}")
        # Tenta com encoding alternativo
        try:
            df_teste = pd.read_csv(caminho, nrows=0, sep=separador, encoding='latin-1')
            colunas_disponiveis = df_teste.columns.tolist()
            print(f"Colunas disponíveis no CSV (latin-1): {colunas_disponiveis}")
            
            # Repete o processo de mapeamento com latin-1
            colunas_para_ler = []
            mapeamento_colunas = {
                'NOTA FISCAL': ['NOTA FISCAL', 'NOTAFISCAL', 'NOTA_FISCAL', 'NF', 'NOTA'],
                'PESO': ['PESO', 'PESO_KG', 'PESO KG', 'PESO_TOTAL'],
                'TOTAL': ['TOTAL', 'TOTAL_NF', 'VALOR_TOTAL', 'VALOR TOTAL', 'VALOR'],
                'DATA': ['DATA', 'DATE', 'DT', 'DATA_NF', 'DATA EMISSÃO', 'EMISSÃO']
            }
            
            for coluna_base, alternativas in mapeamento_colunas.items():
                for coluna_csv in colunas_disponiveis:
                    if any(alt in coluna_csv.upper() for alt in alternativas):
                        colunas_para_ler.append(coluna_csv)
                        print(f"Usando coluna '{coluna_csv}' para {coluna_base}")
                        break
            
            df = pd.read_csv(caminho, sep=separador, usecols=colunas_para_ler, encoding='latin-1')
            
            # Renomeia as colunas
            rename_dict = {}
            for coluna_csv in colunas_para_ler:
                for coluna_base, alternativas in mapeamento_colunas.items():
                    if any(alt in coluna_csv.upper() for alt in alternativas):
                        rename_dict[coluna_csv] = coluna_base
            
            df = df.rename(columns=rename_dict)
            
            # Filtra por data se especificado e se a coluna DATA existe
            if (data_inicio or data_fim) and 'DATA' in df.columns:
                print("Filtrando CSV por período de datas...")
                
                # Converte a coluna DATA para datetime
                try:
                    df['DATA'] = pd.to_datetime(df['DATA'], dayfirst=True, errors='coerce')
                    
                    # Aplica os filtros de data
                    if data_inicio:
                        df = df[df['DATA'] >= data_inicio]
                    if data_fim:
                        # Adiciona 1 dia para incluir a data final completa
                        data_fim_ajustada = data_fim + pd.Timedelta(days=1)
                        df = df[df['DATA'] < data_fim_ajustada]
                    
                    print(f"Período filtrado no CSV: {data_inicio.strftime('%d/%m/%Y') if data_inicio else 'Início indefinido'} a {data_fim.strftime('%d/%m/%Y') if data_fim else 'Fim indefinido'}")
                    print(f"Registros no CSV após filtro de data: {len(df)}")
                    
                except Exception as e:
                    print(f"Erro ao filtrar CSV por data: {e}")
                    print("Continuando sem filtro de data no CSV...")
            
            return df
            
        except Exception as e2:
            print(f"Erro ao ler CSV com encoding alternativo: {e2}")
            return None

def obter_periodo_usuario():
    """
    Solicita o período desejado ao usuário
    """
    print("\n=== FILTRO POR PERÍODO ===")
    print("Digite o período que deseja analisar (formato DD/MM/AAAA)")
    print("Deixe em branco para não usar filtro de data")
    
    while True:
        data_inicio_str = input("Data de início (DD/MM/AAAA): ").strip()
        data_fim_str = input("Data de fim (DD/MM/AAAA): ").strip()
        
        data_inicio = None
        data_fim = None
        
        try:
            if data_inicio_str:
                data_inicio = datetime.strptime(data_inicio_str, '%d/%m/%Y')
            if data_fim_str:
                data_fim = datetime.strptime(data_fim_str, '%d/%m/%Y')
            
            # Validação das datas
            if data_inicio and data_fim and data_inicio > data_fim:
                print("ERRO: Data de início não pode ser maior que data de fim!")
                continue
                
            return data_inicio, data_fim
            
        except ValueError as e:
            print("ERRO: Formato de data inválido! Use DD/MM/AAAA")
            continuar = input("Deseja tentar novamente? (S/N): ").strip().upper()
            if continuar != 'S':
                return None, None

def processar_planilhas():
    # Caminhos das planilhas
    caminho_expedicao = r"Z:\RODRIGO - LOGISTICA\CONTROLE DE EXPEDIÇÃO NOVEMBRO.xlsx"
    caminho_csv = r"S:\hor\excel\20251101.csv"
    
    # Verifica se os arquivos existem
    if not os.path.exists(caminho_expedicao):
        print(f"ERRO: Arquivo não encontrado: {caminho_expedicao}")
        return
    
    if not os.path.exists(caminho_csv):
        print(f"ERRO: Arquivo não encontrado: {caminho_csv}")
        return
    
    # Solicita o período ao usuário
    data_inicio, data_fim = obter_periodo_usuario()
    
    try:
        # Lê a planilha de controle de expedição - AGORA INCLUINDO OPERAÇÃO
        print("Lendo planilha de controle de expedição...")
        df_expedicao = pd.read_excel(
            caminho_expedicao, 
            sheet_name='JAN-FEV-MAR-ABR-MAI-JUN',
            header=3,  # Cabeçalho na linha 4 (índice 3)
            usecols=['NF', 'VOG', 'R$ NF', 'STATUS', 'DATA', 'OPERAÇÃO'],  # Adicionando OPERAÇÃO
            dtype={'NF': str}  # Força a leitura como string
        )
        
        # Remove linhas com NF vazia
        df_expedicao = df_expedicao.dropna(subset=['NF'])
        
        # Filtra apenas os status desejados
        status_validos = ['ENTREGUE', 'EM ROTA', 'DEVOLUÇÃO']
        df_expedicao_filtrado = df_expedicao[df_expedicao['STATUS'].isin(status_validos)].copy()
        
        # FILTRA POR OPERAÇÃO VOG
        operacoes_vog = ['VOG', 'VOG 2ºSAIDA', 'VOG 2 SAIDA', 'VOG 2SAIDA', 'VOG 2º SAIDA']
        df_expedicao_filtrado = df_expedicao_filtrado[
            df_expedicao_filtrado['OPERAÇÃO'].isin(operacoes_vog)
        ]
        
        print(f"Total de notas fiscais: {len(df_expedicao)}")
        print(f"Notas fiscais com status válido ({', '.join(status_validos)}): {len(df_expedicao_filtrado)}")
        print(f"Notas fiscais com operação VOG: {len(df_expedicao_filtrado)}")
        
        # CORREÇÃO: Usa a função específica para formatar notas fiscais
        df_expedicao_filtrado['NF'] = df_expedicao_filtrado['NF'].apply(formatar_nota_fiscal)
        
        print("Exemplo de notas fiscais após formatação:")
        print(df_expedicao_filtrado['NF'].head(10).tolist())
        
        # Converte a coluna DATA da planilha de expedição para datetime
        if 'DATA' in df_expedicao_filtrado.columns:
            # Cria uma cópia da coluna DATA para usar no relatório final
            df_expedicao_filtrado['DATA_EXPEDICAO'] = pd.to_datetime(df_expedicao_filtrado['DATA'], errors='coerce')
            
            # APLICA FILTRO DE DATA NA PLANILHA DE EXPEDIÇÃO
            if data_inicio or data_fim:
                print("Filtrando controle de expedição por período de datas...")
                
                # Aplica os filtros de data
                if data_inicio:
                    df_expedicao_filtrado = df_expedicao_filtrado[df_expedicao_filtrado['DATA_EXPEDICAO'] >= data_inicio]
                if data_fim:
                    # Adiciona 1 dia para incluir a data final completa
                    data_fim_ajustada = data_fim + pd.Timedelta(days=1)
                    df_expedicao_filtrado = df_expedicao_filtrado[df_expedicao_filtrado['DATA_EXPEDICAO'] < data_fim_ajustada]
                
                print(f"Período filtrado na expedição: {data_inicio.strftime('%d/%m/%Y') if data_inicio else 'Início indefinido'} a {data_fim.strftime('%d/%m/%Y') if data_fim else 'Fim indefinido'}")
                print(f"Notas fiscais na expedição após filtro de data: {len(df_expedicao_filtrado)}")
        else:
            print("AVISO: Coluna DATA não encontrada na planilha de expedição")
            df_expedicao_filtrado['DATA_EXPEDICAO'] = None
        
        # Limpa os valores monetários (usaremos apenas internamente para comparação)
        print("Limpando valores do controle de expedição...")
        vog_limpo = df_expedicao_filtrado['VOG'].apply(limpar_valor_numerico)
        valor_nf_limpo = df_expedicao_filtrado['R$ NF'].apply(limpar_valor_monetario)
        
        print(f"Encontradas {len(df_expedicao_filtrado)} notas fiscais válidas no controle de expedição (após filtros)")
        
    except Exception as e:
        print(f"Erro ao ler planilha de expedição: {e}")
        return
    
    # Lê o arquivo CSV com detecção automática de colunas e filtro de data
    print("Lendo arquivo CSV...")
    df_csv = ler_csv_com_cabecalho(caminho_csv, data_inicio, data_fim)
    
    if df_csv is None:
        print("Não foi possível ler o arquivo CSV. Verifique o formato do arquivo.")
        return
    
    print(f"CSV lido com sucesso. Colunas: {df_csv.columns.tolist()}")
    
    # Remove linhas com NOTA FISCAL vazia
    df_csv = df_csv.dropna(subset=['NOTA FISCAL'])
    
    # CORREÇÃO: Usa a função específica para formatar notas fiscais no CSV também
    df_csv['NOTA FISCAL'] = df_csv['NOTA FISCAL'].apply(formatar_nota_fiscal)
    
    print("Exemplo de notas fiscais do CSV após formatação:")
    print(df_csv['NOTA FISCAL'].head(10).tolist())
    
    # Limpa as colunas numéricas do CSV (usaremos apenas internamente para comparação)
    print("Limpando valores numéricos do CSV...")
    peso_limpo = df_csv['PESO'].apply(limpar_valor_numerico)
    total_limpo = df_csv['TOTAL'].apply(limpar_valor_monetario)
    
    # Adiciona as colunas limpas ao DataFrame original para agrupamento
    df_csv['PESO_COMPARACAO'] = peso_limpo
    df_csv['TOTAL_COMPARACAO'] = total_limpo
    
    # Converte a coluna DATA do CSV para datetime se existir
    if 'DATA' in df_csv.columns:
        df_csv['DATA_CSV'] = pd.to_datetime(df_csv['DATA'], dayfirst=True, errors='coerce')
    else:
        df_csv['DATA_CSV'] = None
    
    # Agrupa por NOTA FISCAL mantendo a DATA e somando os valores
    df_agrupado = df_csv.groupby('NOTA FISCAL').agg({
        'PESO': 'first',           # Mantém o valor original do PESO
        'TOTAL': 'first',          # Mantém o valor original do TOTAL
        'DATA_CSV': 'first',       # Mantém a primeira data do CSV encontrada
        'PESO_COMPARACAO': 'sum',  # Soma os valores limpos para comparação
        'TOTAL_COMPARACAO': 'sum'  # Soma os valores limpos para comparação
    }).reset_index()
    
    print(f"Encontradas {len(df_agrupado)} notas fiscais únicas no CSV (após filtros)")
    
    # Realiza o merge das planilhas - AGORA USANDO STRINGS
    print("Comparando as planilhas...")
    df_comparacao = pd.merge(
        df_expedicao_filtrado,
        df_agrupado,
        left_on='NF',
        right_on='NOTA FISCAL',
        how='left',
        indicator=True
    )
    
    # VERIFICA ITENS DO CSV QUE NÃO ESTÃO NO CONTROLE DE EXPEDIÇÃO
    df_csv_sem_expedicao = pd.merge(
        df_agrupado,
        df_expedicao_filtrado,
        left_on='NOTA FISCAL',
        right_on='NF',
        how='left',
        indicator=True
    )
    nfs_csv_sem_expedicao = df_csv_sem_expedicao[df_csv_sem_expedicao['_merge'] == 'left_only']
    
    # Identifica divergências
    resultados = []
    
    # 1. DIVERGÊNCIAS DO CONTROLE DE EXPEDIÇÃO (NFs que não estão no CSV)
    for index, row in df_comparacao.iterrows():
        nf = row['NF']
        vog_expedicao = vog_limpo.iloc[index] if index < len(vog_limpo) else 0
        valor_nf_expedicao = valor_nf_limpo.iloc[index] if index < len(valor_nf_limpo) else 0
        peso_csv = row['PESO_COMPARACAO'] if 'PESO_COMPARACAO' in row and not pd.isna(row['PESO_COMPARACAO']) else 0
        total_csv = row['TOTAL_COMPARACAO'] if 'TOTAL_COMPARACAO' in row and not pd.isna(row['TOTAL_COMPARACAO']) else 0
        status = row['STATUS']
        operacao = row['OPERAÇÃO'] if 'OPERAÇÃO' in row else ''
        merge_status = row['_merge']
        
        # Obtém as DATAS de ambas as fontes
        data_expedicao = None
        data_csv = None
        
        if 'DATA_EXPEDICAO' in row and not pd.isna(row['DATA_EXPEDICAO']):
            data_expedicao = row['DATA_EXPEDICAO']
        
        if 'DATA_CSV' in row and not pd.isna(row['DATA_CSV']):
            data_csv = row['DATA_CSV']
        
        divergencias = []
        
        # Verifica se a NF foi encontrada no CSV
        if merge_status == 'left_only':
            divergencias.append("NF não encontrada no CSV")
        else:
            # Verifica divergência de PESO/VOG
            if abs(vog_expedicao - peso_csv) > 0.01:  # Tolerância para diferenças decimais
                divergencias.append(f"PESO divergente: Expedição={vog_expedicao:.2f}, CSV={peso_csv:.2f}, Diferença={abs(vog_expedicao - peso_csv):.2f}")
            
            # Verifica divergência de TOTAL/R$ NF
            if abs(valor_nf_expedicao - total_csv) > 0.01:
                divergencias.append(f"VALOR divergente: Expedição={valor_nf_expedicao:.2f}, CSV={total_csv:.2f}, Diferença={abs(valor_nf_expedicao - total_csv):.2f}")
        
        # Verifica possíveis erros de digitação nos valores originais
        if pd.notna(row['VOG']):
            vog_original = str(row['VOG'])
            # Detecta padrões como "1,564,100" que deveria ser "1564.10"
            if vog_original.count(',') > 1:
                divergencias.append(f"Possível erro de digitação no VOG (múltiplas vírgulas): {vog_original}")
            elif ',' in vog_original and '.' in vog_original.split(',')[0]:
                divergencias.append(f"Possível erro de digitação no VOG: {vog_original}")
        
        if pd.notna(row['R$ NF']):
            valor_original = str(row['R$ NF'])
            if valor_original.count(',') > 1:
                divergencias.append(f"Possível erro de digitação no R$ NF (múltiplas vírgulas): {valor_original}")
            elif ',' in valor_original and '.' in valor_original.split(',')[0]:
                divergencias.append(f"Possível erro de digitação no R$ NF: {valor_original}")
        
        if divergencias:
            # CONVERTE NF PARA INTEIRO
            nf_inteiro = converter_para_inteiro_nota_fiscal(nf)
            
            # CONVERTE PESO_CSV E TOTAL_CSV PARA FLOAT
            peso_csv_float = float(peso_csv) if peso_csv else 0.0
            total_csv_float = float(total_csv) if total_csv else 0.0
            
            resultados.append({
                'NF': nf_inteiro,  # AGORA É INTEIRO
                'DATA_EXPEDICAO': data_expedicao,  # Data da expedição
                'DATA_CSV': data_csv,              # Data do CSV
                'STATUS': status,
                'OPERAÇÃO': operacao,
                'Divergências': ' | '.join(divergencias),
                'VOG_Expedição': row['VOG'],
                'PESO_CSV': peso_csv_float,  # AGORA É FLOAT
                'R$ NF_Expedição': row['R$ NF'],
                'TOTAL_CSV': total_csv_float  # AGORA É FLOAT
            })
    
    # 2. DIVERGÊNCIAS DO CSV (NFs que não estão no controle de expedição)
    for index, row in nfs_csv_sem_expedicao.iterrows():
        nf = row['NOTA FISCAL']
        data_csv = row['DATA_CSV'] if 'DATA_CSV' in row and not pd.isna(row['DATA_CSV']) else None
        
        # CORREÇÃO: Verifica se as colunas PESO e TOTAL existem
        peso_csv = row['PESO'] if 'PESO' in row else ''
        total_csv = row['TOTAL'] if 'TOTAL' in row else ''
        
        # CONVERTE NF PARA INTEIRO
        nf_inteiro = converter_para_inteiro_nota_fiscal(nf)
        
        # CONVERTE PESO_CSV E TOTAL_CSV PARA FLOAT
        peso_csv_float = limpar_valor_numerico(peso_csv) if peso_csv else 0.0
        total_csv_float = limpar_valor_monetario(total_csv) if total_csv else 0.0
        
        resultados.append({
            'NF': nf_inteiro,      # AGORA É INTEIRO
            'DATA_EXPEDICAO': None,      # Não tem data da expedição
            'DATA_CSV': data_csv,        # Data do CSV
            'STATUS': 'N/A',
            'OPERAÇÃO': 'N/A',
            'Divergências': 'NF do CSV não encontrada no controle de expedição',
            'VOG_Expedição': 'N/A',
            'PESO_CSV': peso_csv_float,  # AGORA É FLOAT
            'R$ NF_Expedição': 'N/A',
            'TOTAL_CSV': total_csv_float  # AGORA É FLOAT
        })
    
    # Cria relatório final
    if resultados:
        df_relatorio = pd.DataFrame(resultados)
        
        # Reorganiza as colunas para incluir ambas as datas
        colunas_ordenadas = [
            'NF', 
            'DATA_EXPEDICAO', 
            'DATA_CSV',
            'STATUS', 
            'OPERAÇÃO', 
            'Divergências', 
            'VOG_Expedição', 
            'PESO_CSV', 
            'R$ NF_Expedição', 
            'TOTAL_CSV'
        ]
        # Garante que apenas as colunas existentes sejam usadas
        colunas_ordenadas = [col for col in colunas_ordenadas if col in df_relatorio.columns]
        df_relatorio = df_relatorio[colunas_ordenadas]
        
        # Salva o relatório no Downloads
        downloads_path = str(Path.home() / "Downloads")
        caminho_relatorio = os.path.join(downloads_path, "RELATORIO_DIVERGENCIAS.xlsx")
        
        # Cria um ExcelWriter para formatar as colunas
        with pd.ExcelWriter(caminho_relatorio, engine='openpyxl') as writer:
            df_relatorio.to_excel(writer, index=False, sheet_name='Divergências')
            
            # Acessa a planilha para ajustar o formato das colunas
            worksheet = writer.sheets['Divergências']
            
            # Formata as colunas de data como data brasileira
            date_columns = ['DATA_EXPEDICAO', 'DATA_CSV']
            for col_idx, col_name in enumerate(df_relatorio.columns):
                if col_name in date_columns:
                    col_letter = chr(65 + col_idx)  # 65 = 'A'
                    for row in range(2, len(df_relatorio) + 2):  # +2 porque a linha 1 é o cabeçalho
                        cell = worksheet[f'{col_letter}{row}']
                        if cell.value and not pd.isna(cell.value):
                            cell.number_format = 'DD/MM/YYYY'
            
            # Formata a coluna NF como número inteiro sem decimais
            for col_idx, col_name in enumerate(df_relatorio.columns):
                if col_name == 'NF':
                    col_letter = chr(65 + col_idx)  # 65 = 'A'
                    for row in range(2, len(df_relatorio) + 2):
                        cell = worksheet[f'{col_letter}{row}']
                        cell.number_format = '0'
            
            # Formata as colunas PESO_CSV e TOTAL_CSV como float com 2 casas decimais
            float_columns = ['PESO_CSV', 'TOTAL_CSV']
            for col_idx, col_name in enumerate(df_relatorio.columns):
                if col_name in float_columns:
                    col_letter = chr(65 + col_idx)
                    for row in range(2, len(df_relatorio) + 2):
                        cell = worksheet[f'{col_letter}{row}']
                        cell.number_format = '#,##0.00'
            
            # Ajusta a largura das colunas automaticamente
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
        
        print(f"\n=== RELATÓRIO DE DIVERGÊNCIAS ===")
        print(f"Total de divergências encontradas: {len(resultados)}")
        print(f"Relatório salvo em: {caminho_relatorio}")
        
        # Exibe as primeiras 10 divergências
        print("\nPrimeiras 10 divergências:")
        for i, resultado in enumerate(resultados[:10]):
            # Formata as datas para exibição no console
            data_expedicao = resultado['DATA_EXPEDICAO'].strftime('%d/%m/%Y') if resultado['DATA_EXPEDICAO'] and not pd.isna(resultado['DATA_EXPEDICAO']) else 'N/A'
            data_csv = resultado['DATA_CSV'].strftime('%d/%m/%Y') if resultado['DATA_CSV'] and not pd.isna(resultado['DATA_CSV']) else 'N/A'
            
            print(f"\n{i+1}. NF: {resultado['NF']}")
            print(f"   Data Expedição: {data_expedicao} | Data CSV: {data_csv}")
            print(f"   Status: {resultado['STATUS']} | Operação: {resultado['OPERAÇÃO']}")
            print(f"   Divergências: {resultado['Divergências']}")
            
        if len(resultados) > 10:
            print(f"\n... e mais {len(resultados) - 10} divergências")
            
    else:
        print("\n✅ Nenhuma divergência encontrada! Todas as notas fiscais estão consistentes.")
    
    # Estatísticas adicionais
    nfs_sem_match_expedicao = len(df_comparacao[df_comparacao['_merge'] == 'left_only'])
    nfs_sem_match_csv = len(nfs_csv_sem_expedicao)
    
    if nfs_sem_match_expedicao > 0:
        print(f"\n⚠️  {nfs_sem_match_expedicao} notas fiscais do controle de expedição não foram encontradas no CSV")
    
    if nfs_sem_match_csv > 0:
        print(f"⚠️  {nfs_sem_match_csv} notas fiscais do CSV não foram encontradas no controle de expedição")
    
    # Mostra algumas estatísticas
    print(f"\n=== ESTATÍSTICAS ===")
    print(f"Notas fiscais totais no controle de expedição: {len(df_expedicao)}")
    print(f"Notas fiscais com status válido e operação VOG (após filtros): {len(df_expedicao_filtrado)}")
    print(f"Notas fiscais únicas no CSV (após filtros): {len(df_agrupado)}")
    print(f"Notas fiscais do expedição sem correspondência no CSV: {nfs_sem_match_expedicao}")
    print(f"Notas fiscais do CSV sem correspondência no expedição: {nfs_sem_match_csv}")
    
    # Estatísticas por status
    print(f"\nDistribuição por status:")
    for status in status_validos:
        count = len(df_expedicao_filtrado[df_expedicao_filtrado['STATUS'] == status])
        print(f"  {status}: {count}")
    
    # Estatísticas por operação
    if 'OPERAÇÃO' in df_expedicao_filtrado.columns:
        print(f"\nDistribuição por operação VOG:")
        operacoes_count = df_expedicao_filtrado['OPERAÇÃO'].value_counts()
        for operacao, count in operacoes_count.items():
            print(f"  {operacao}: {count}")

if __name__ == "__main__":
    processar_planilhas()
    input("\nPressione Enter para sair...")