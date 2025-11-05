import pandas as pd
import numpy as np
import os
from pathlib import Path

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

def ler_csv_com_cabecalho(caminho):
    """
    Lê o CSV detectando automaticamente o cabeçalho
    """
    separador = detectar_separador_csv(caminho)
    
    # Primeiro, vamos ler apenas o cabeçalho para ver as colunas
    try:
        df_teste = pd.read_csv(caminho, nrows=0, sep=separador, encoding='utf-8')
        colunas_disponiveis = df_teste.columns.tolist()
        print(f"Colunas disponíveis no CSV: {colunas_disponiveis}")
        
        # Mapeia possíveis nomes de colunas
        colunas_necessarias = []
        mapeamento_colunas = {
            'NOTA FISCAL': ['NOTA FISCAL', 'NOTAFISCAL', 'NOTA_FISCAL', 'NF', 'NOTA'],
            'PESO': ['PESO', 'PESO_KG', 'PESO KG', 'PESO_TOTAL'],
            'TOTAL': ['TOTAL', 'TOTAL_NF', 'VALOR_TOTAL', 'VALOR TOTAL', 'VALOR']
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
            if not encontrou:
                print(f"AVISO: Coluna {coluna_base} não encontrada. Alternativas: {alternativas}")
                return None
        
        # Agora lê o arquivo completo com as colunas detectadas
        df = pd.read_csv(caminho, sep=separador, usecols=colunas_para_ler, encoding='utf-8')
        
        # Renomeia as colunas para nomes padrão
        rename_dict = {}
        for coluna_csv in colunas_para_ler:
            for coluna_base, alternativas in mapeamento_colunas.items():
                if any(alt in coluna_csv.upper() for alt in alternativas):
                    rename_dict[coluna_csv] = coluna_base
        
        df = df.rename(columns=rename_dict)
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
                'TOTAL': ['TOTAL', 'TOTAL_NF', 'VALOR_TOTAL', 'VALOR TOTAL', 'VALOR']
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
            return df
            
        except Exception as e2:
            print(f"Erro ao ler CSV com encoding alternativo: {e2}")
            return None

def processar_planilhas():
    # Caminhos das planilhas
    caminho_expedicao = r"Z:\RODRIGO - LOGISTICA\Cópia de CONTROLE DE EXPEDIÇÃO - OUTUBRO.xlsx"
    caminho_csv = r"S:\hor\excel\20251027.csv"
    
    # Verifica se os arquivos existem
    if not os.path.exists(caminho_expedicao):
        print(f"ERRO: Arquivo não encontrado: {caminho_expedicao}")
        return
    
    if not os.path.exists(caminho_csv):
        print(f"ERRO: Arquivo não encontrado: {caminho_csv}")
        return
    
    try:
        # Lê a planilha de controle de expedição
        print("Lendo planilha de controle de expedição...")
        df_expedicao = pd.read_excel(
            caminho_expedicao, 
            sheet_name='JAN-FEV-MAR-ABR-MAI-JUN',
            header=3,  # Cabeçalho na linha 4 (índice 3)
            usecols=['NF', 'VOG', 'R$ NF', 'STATUS']
        )
        
        # Remove linhas com NF vazia
        df_expedicao = df_expedicao.dropna(subset=['NF'])
        
        # Filtra apenas os status desejados
        status_validos = ['ENTREGUE', 'EM ROTA', 'DEVOLUÇÃO']
        df_expedicao_filtrado = df_expedicao[df_expedicao['STATUS'].isin(status_validos)].copy()
        
        print(f"Total de notas fiscais: {len(df_expedicao)}")
        print(f"Notas fiscais com status válido ({', '.join(status_validos)}): {len(df_expedicao_filtrado)}")
        
        # Converte NF para string e remove espaços
        df_expedicao_filtrado['NF'] = df_expedicao_filtrado['NF'].astype(str).str.strip()
        
        # Limpa os valores monetários
        print("Limpando valores do controle de expedição...")
        df_expedicao_filtrado['VOG_limpo'] = df_expedicao_filtrado['VOG'].apply(limpar_valor_numerico)
        df_expedicao_filtrado['R$ NF_limpo'] = df_expedicao_filtrado['R$ NF'].apply(limpar_valor_monetario)
        
        print(f"Encontradas {len(df_expedicao_filtrado)} notas fiscais válidas no controle de expedição")
        
    except Exception as e:
        print(f"Erro ao ler planilha de expedição: {e}")
        return
    
    # Lê o arquivo CSV com detecção automática de colunas
    print("Lendo arquivo CSV...")
    df_csv = ler_csv_com_cabecalho(caminho_csv)
    
    if df_csv is None:
        print("Não foi possível ler o arquivo CSV. Verifique o formato do arquivo.")
        return
    
    print(f"CSV lido com sucesso. Colunas: {df_csv.columns.tolist()}")
    
    # Remove linhas com NOTA FISCAL vazia
    df_csv = df_csv.dropna(subset=['NOTA FISCAL'])
    
    # Converte NOTA FISCAL para string e remove espaços
    df_csv['NOTA FISCAL'] = df_csv['NOTA FISCAL'].astype(str).str.strip()
    
    # Limpa as colunas numéricas do CSV
    print("Limpando valores numéricos do CSV...")
    df_csv['PESO_limpo'] = df_csv['PESO'].apply(limpar_valor_numerico)
    df_csv['TOTAL_limpo'] = df_csv['TOTAL'].apply(limpar_valor_monetario)
    
    # Agrupa por NOTA FISCAL e soma PESO e TOTAL
    df_agrupado = df_csv.groupby('NOTA FISCAL').agg({
        'PESO_limpo': 'sum',
        'TOTAL_limpo': 'sum'
    }).reset_index()
    
    print(f"Encontradas {len(df_agrupado)} notas fiscais únicas no CSV")
    
    # Realiza o merge das planilhas
    print("Comparando as planilhas...")
    df_comparacao = pd.merge(
        df_expedicao_filtrado,
        df_agrupado,
        left_on='NF',
        right_on='NOTA FISCAL',
        how='left',
        indicator=True
    )
    
    # Identifica divergências
    resultados = []
    
    for index, row in df_comparacao.iterrows():
        nf = row['NF']
        vog_expedicao = row['VOG_limpo']
        valor_nf_expedicao = row['R$ NF_limpo']
        peso_csv = row['PESO_limpo'] if not pd.isna(row['PESO_limpo']) else 0
        total_csv = row['TOTAL_limpo'] if not pd.isna(row['TOTAL_limpo']) else 0
        status = row['STATUS']
        merge_status = row['_merge']
        
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
            resultados.append({
                'NF': nf,
                'STATUS': status,
                'Divergências': ' | '.join(divergencias),
                'VOG_Expedição': row['VOG'],
                'VOG_Limpo': vog_expedicao,
                'PESO_CSV': row['PESO'] if 'PESO' in row else '',
                'PESO_CSV_Limpo': peso_csv,
                'R$ NF_Expedição': row['R$ NF'],
                'R$ NF_Limpo': valor_nf_expedicao,
                'TOTAL_CSV': row['TOTAL'] if 'TOTAL' in row else '',
                'TOTAL_CSV_Limpo': total_csv
            })
    
    # Cria relatório final
    if resultados:
        df_relatorio = pd.DataFrame(resultados)
        
        # Salva o relatório no Downloads
        downloads_path = str(Path.home() / "Downloads")
        caminho_relatorio = os.path.join(downloads_path, "RELATORIO_DIVERGENCIAS.xlsx")
        df_relatorio.to_excel(caminho_relatorio, index=False)
        
        print(f"\n=== RELATÓRIO DE DIVERGÊNCIAS ===")
        print(f"Total de divergências encontradas: {len(resultados)}")
        print(f"Relatório salvo em: {caminho_relatorio}")
        
        # Exibe as primeiras 10 divergências
        print("\nPrimeiras 10 divergências:")
        for i, resultado in enumerate(resultados[:10]):
            print(f"\n{i+1}. NF: {resultado['NF']} | Status: {resultado['STATUS']}")
            print(f"   Divergências: {resultado['Divergências']}")
            
        if len(resultados) > 10:
            print(f"\n... e mais {len(resultados) - 10} divergências")
            
    else:
        print("\n✅ Nenhuma divergência encontrada! Todas as notas fiscais estão consistentes.")
    
    # Estatísticas adicionais
    nfs_sem_match = len(df_comparacao[df_comparacao['_merge'] == 'left_only'])
    if nfs_sem_match > 0:
        print(f"\n⚠️  {nfs_sem_match} notas fiscais não foram encontradas no CSV")
    
    # Mostra algumas estatísticas
    print(f"\n=== ESTATÍSTICAS ===")
    print(f"Notas fiscais totais no controle de expedição: {len(df_expedicao)}")
    print(f"Notas fiscais com status válido: {len(df_expedicao_filtrado)}")
    print(f"Notas fiscais únicas no CSV: {len(df_agrupado)}")
    print(f"Notas fiscais sem correspondência: {nfs_sem_match}")
    
    # Estatísticas por status
    print(f"\nDistribuição por status:")
    for status in status_validos:
        count = len(df_expedicao_filtrado[df_expedicao_filtrado['STATUS'] == status])
        print(f"  {status}: {count}")

if __name__ == "__main__":
    processar_planilhas()
    input("\nPressione Enter para sair...")