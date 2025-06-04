import pandas as pd
import re
import logging
import glob

# ================================================================
print("--- EXECUTANDO SCRIPT 'AUTOMACAO.py' (v11 - Com Soma de Totais) ---")
# ================================================================

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def extrair_dados_completos(caminho_txt: str) -> dict:
    dados_extraidos = {}
    try:
        with open(caminho_txt, 'r', encoding='utf-8') as f:
            texto_completo = f.read()
        
        texto_normalizado = texto_completo.replace('\xa0', ' ')
        logging.info("Arquivo .txt lido e normalizado. Extraindo dados completos...")

        blocos_de_prestador = re.split(r'Dados do Prestador', texto_normalizado)
        
        contador_sucesso_geral = 0
        for i_prestador, bloco_prestador in enumerate(blocos_de_prestador):
            if len(bloco_prestador) < 100: continue

            nome_prestador_match = re.search(
                r"7 - Nome do Contratado[\s\S]*?\n\s*(\d+)\s+(.*?)\s+(\d{2,3}\.\d{3}\.\d{3}/\d{4}-\d{2}|\d{3}\.\d{3}\.\d{3}-\d{2})",
                bloco_prestador
            )
            nome_prestador = nome_prestador_match.group(2).strip() if nome_prestador_match else "NÃO ENCONTRADO"
            
            if nome_prestador == "NÃO ENCONTRADO" and len(bloco_prestador.strip()) > 0 :
                logging.warning(f"Não foi possível encontrar o nome do prestador no bloco de prestador que começa com: '{bloco_prestador[:100].strip()}'")

            blocos_de_pagamento = re.split(r'Dados do Pagamento', bloco_prestador)
            for i_pagamento, bloco_pagamento in enumerate(blocos_de_pagamento):
                if len(bloco_pagamento) < 100: continue

                gto_match = re.search(r"13 - Número da Guia.*?\s+(\d{8,})", bloco_pagamento, re.DOTALL)
                numero_gto_bloco = gto_match.group(1).strip() if gto_match else None

                nome_paciente = "NÃO ENCONTRADO"
                linha_dados_cabecalho_match = re.search(r"13 - Número da Guia\s*\n([^\n]+)", bloco_pagamento)
                if linha_dados_cabecalho_match:
                    linha_dados_completa = linha_dados_cabecalho_match.group(1).strip()
                    partes_linha = re.split(r'\s{2,}', linha_dados_completa)
                    if len(partes_linha) >= 2:
                        if numero_gto_bloco and partes_linha[-1].strip() == numero_gto_bloco:
                            nome_paciente = partes_linha[-2].strip()
                        elif len(partes_linha) >= 5: 
                             nome_paciente = partes_linha[-2].strip()

                justificativa_glosa = ""
                obs_match = re.search(r"27 - Observação / Justificativa\s*\n([\s\S]*?)(?=\n\s*28 - Valor Total Informado Guia \(R\$\)|Dados do Pagamento|Dados do Prestador|Totais|\Z)", bloco_pagamento, re.DOTALL)
                if obs_match:
                    texto_observacao = obs_match.group(1).strip()
                    glosa_reason_match = re.search(r"Glosas:(.*?)(?=\s*Total de Pontos:|\Z)", texto_observacao, re.DOTALL | re.IGNORECASE)
                    if glosa_reason_match:
                        justificativa_glosa = "Glosas:" + glosa_reason_match.group(1).strip().replace("\n", " ")

                valores_match = re.search(r"32 - Valor Total Liberado Guia \(R\$\)([\s\S]*)", bloco_pagamento)

                if numero_gto_bloco and valores_match:
                    linha_dos_valores = valores_match.group(1).strip().split('\n')[0].strip()
                    valores_numericos = linha_dos_valores.split()
                    
                    sucesso_na_leitura_valores = False
                    if len(valores_numericos) == 5:
                        glosa_str, liberado_str = valores_numericos[2], valores_numericos[4]
                        sucesso_na_leitura_valores = True
                    elif len(valores_numericos) == 3:
                        glosa_str, liberado_str = valores_numericos[0], valores_numericos[2]
                        sucesso_na_leitura_valores = True

                    if sucesso_na_leitura_valores:
                        glosa = float(glosa_str.replace(',', ''))
                        liberado = float(liberado_str.replace(',', ''))
                        dados_extraidos[numero_gto_bloco] = {
                            'profissional': nome_prestador,
                            'paciente': nome_paciente,
                            'glosa': glosa,
                            'liberado': liberado,
                            'justificativa': justificativa_glosa
                        }
                        contador_sucesso_geral += 1
        
        if contador_sucesso_geral > 0:
            logging.info(f"SUCESSO! Dados de {contador_sucesso_geral} GTOs extraídos!")
        else:
            logging.error("Nenhum dado de GTO pôde ser extraído do relatorio.txt.")
            
    except Exception as e:
        logging.error(f"Ocorreu um erro ao processar o arquivo de texto: {e}", exc_info=True)
    return dados_extraidos

def ler_gtos_do_csv(caminho_csv: str) -> list:
    try:
        df = pd.read_csv(caminho_csv, usecols=["GTO"], dtype={"GTO": str})
        return df["GTO"].dropna().astype(str).unique().tolist()
    except Exception as e:
        logging.error(f"Não foi possível ler a coluna 'GTO' do CSV: {e}")
        return []

### SCRIPT PRINCIPAL ###
def main():
    try:
        padrao_csv = 'Relatorio-GTODigital-*.csv'
        arquivos_csv = glob.glob(padrao_csv)
        if not arquivos_csv:
            logging.error(f"ERRO: Nenhum arquivo CSV com o padrão '{padrao_csv}' foi encontrado.")
            return
        if len(arquivos_csv) > 1:
            logging.error("ERRO: Mais de um arquivo de relatório CSV foi encontrado.")
            return
        
        caminho_csv = arquivos_csv[0]
        caminho_txt = 'relatorio.txt'
        logging.info(f"Usando o arquivo CSV: {caminho_csv}")
        
        dados_completos = extrair_dados_completos(caminho_txt)
        lista_gtos_csv = ler_gtos_do_csv(caminho_csv)

        if not dados_completos or not lista_gtos_csv:
            raise ValueError("Não foi possível extrair dados de um ou ambos os arquivos.")

        resultados = []
        for gto in lista_gtos_csv:
            if gto in dados_completos:
                info = dados_completos[gto]
                resultados.append({
                    'GTO': gto, 'Status': 'Encontrada no Relatório',
                    'Profissional': info['profissional'], 'Paciente': info['paciente'],
                    'Valor Total Glosa (R$)': info['glosa'], 'Valor Total Liberado (R$)': info['liberado'],
                    'Justificativa da Glosa': info['justificativa']
                })
            else:
                resultados.append({
                    'GTO': gto, 'Status': 'NÃO encontrada no Relatório',
                    'Profissional': '', 'Paciente': '', 
                    'Valor Total Glosa (R$)': 0.0, # Usar 0.0 para permitir a soma
                    'Valor Total Liberado (R$)': 0.0, # Usar 0.0 para permitir a soma
                    'Justificativa da Glosa': ''
                })
        
        df_resultado = pd.DataFrame(resultados)
        
        ordem_colunas = ['GTO', 'Status', 'Profissional', 'Paciente', 
                         'Valor Total Glosa (R$)', 'Valor Total Liberado (R$)', 
                         'Justificativa da Glosa']
        df_resultado = df_resultado[ordem_colunas]

        df_resultado = df_resultado.sort_values(
            by=['Profissional', 'Paciente', 'GTO'], 
            ascending=[True, True, True]
        )
        logging.info("Relatório ordenado por Profissional, Paciente e GTO.")

        # --- CÁLCULO E EXIBIÇÃO DOS TOTAIS ---
        # Converte colunas para numérico, erros viram NaN (que sum() ignora)
        df_resultado['Valor Total Glosa (R$)'] = pd.to_numeric(df_resultado['Valor Total Glosa (R$)'], errors='coerce').fillna(0)
        df_resultado['Valor Total Liberado (R$)'] = pd.to_numeric(df_resultado['Valor Total Liberado (R$)'], errors='coerce').fillna(0)

        total_glosas = df_resultado['Valor Total Glosa (R$)'].sum()
        total_liberado = df_resultado['Valor Total Liberado (R$)'].sum()
        
        logging.info("--- TOTAIS GERAIS DO RELATÓRIO ---")
        # Formata os totais como moeda brasileira para exibição no log
        logging.info(f"Soma Total de Glosas: R$ {total_glosas:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        logging.info(f"Soma Total de Valores Liberados: R$ {total_liberado:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        logging.info("-----------------------------------")


        output_filename = 'Resultado_Busca_GTOs_Completo.xlsx'
        
        writer = pd.ExcelWriter(output_filename, engine='xlsxwriter')
        df_resultado.to_excel(writer, sheet_name='Resultado', index=False)

        workbook  = writer.book
        worksheet = writer.sheets['Resultado']
        formato_moeda = workbook.add_format({'num_format': 'R$ #,##0.00'})

        worksheet.set_column('A:A', 15) 
        worksheet.set_column('B:B', 25) 
        worksheet.set_column('C:C', 40) 
        worksheet.set_column('D:D', 40) 
        worksheet.set_column('E:F', 22, formato_moeda) 
        worksheet.set_column('G:G', 60) 
        
        writer.close()
        
        logging.info(f"✅ TAREFA CONCLUÍDA! Relatório final COMPLETO, ORDENADO e com TOTAIS exibidos, foi gerado: {output_filename}")

    except FileNotFoundError as e:
        logging.error(f"Erro: Arquivo não encontrado. Você se lembrou de converter o PDF para '{caminho_txt}'? Detalhe: {e}")
    except Exception as e:
        logging.error(f"Ocorreu um erro inesperado: {e}")

if __name__ == "__main__":
    main()