#!/usr/bin/env python3
"""
Extrator avançado de chaves de acesso de documentos fiscais em PDF.
Este script analisa PDFs na pasta onde está sendo executado (ou subfolder configurado)
e extrai chaves de acesso NFe válidas usando técnicas avançadas de extração e validação.
Os PDFs são separados em duas pastas: com e sem chave de acesso.

Modo de uso:
1. Coloque o executavel na pasta onde estão os PDFs ou configure a pasta de entrada
2. Execute o programa
3. As chaves serão extraídas e os arquivos separados em pastas
"""

import os
import re
import PyPDF2
import pdfplumber
import shutil
from PIL import Image
import io
import sys
import traceback
from datetime import datetime
import argparse

# Importando bibliotecas para GUI caso seja necessário
try:
    import tkinter as tk
    from tkinter import messagebox
    TKINTER_DISPONIVEL = True
except ImportError:
    TKINTER_DISPONIVEL = False

# Verificar se está sendo executado como um executável PyInstaller
IS_FROZEN = getattr(sys, 'frozen', False)

# Diretório onde o script está sendo executado
DIR_ATUAL = os.getcwd()

# Configurações padrão
PASTA_PDFS = DIR_ATUAL  # Por padrão, usar a pasta atual
ARQUIVO_SAIDA = os.path.join(DIR_ATUAL, "chaves_extraidas_final.txt")
PASTA_COM_CHAVE = os.path.join(DIR_ATUAL, "PDFs_Com_Chave")
PASTA_SEM_CHAVE = os.path.join(DIR_ATUAL, "PDFs_Sem_Chave")

# Padrões para diferentes formatos de chaves
PADRAO_CHAVE_NFE = r'\b\d{44}\b'  # Padrão para chaves de acesso de NF-e (44 dígitos)
PADRAO_CHAVE_ESPACOS = r'\b(\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{0,4})\b'
PADRAO_CHAVE_BLOCOS = r'\b(\d{4}[.\s-]*\d{4}[.\s-]*\d{4}[.\s-]*\d{4}[.\s-]*\d{4}[.\s-]*\d{4}[.\s-]*\d{4}[.\s-]*\d{4}[.\s-]*\d{4}[.\s-]*\d{4}[.\s-]*\d{4})\b'
# Padrão Energisa de diferentes formatos
# Padrão após "chave de acesso:" - formato comum nas faturas da Energisa
PADRAO_ENERGISA_1 = r'chave\s+de\s+acesso\s*:?\s*\n?\s*(\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4})\s*\n\s*(\d{4})'
# Padrão com 10 grupos de 4 + quebra de linha + 1 grupo de 4
PADRAO_ENERGISA_2 = r'\b(\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4})\s*\n\s*(\d{4})\b'
# Padrão Dcelt: 11 grupos de 4 dígitos em uma única linha
PADRAO_DCELT = r'\b(\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4}\s+\d{4})\b'

def limpar_chave(chave):
    """Remove espaços e caracteres especiais de uma chave."""
    return re.sub(r'[^0-9]', '', chave)

def validar_chave_nfe(chave):
    """Valida se uma string de 44 dígitos é uma chave de NFe válida.
    Ajustado para reconhecer chaves da Energisa e outros formatos."""
    if len(chave) != 44:
        return False
    
    # Verificar se pode ser uma chave no formato Energisa (começa com 50)
    if chave.startswith('50'):
        # Chaves Energisa geralmente começam com 50 (UF MS) ou similares
        # Vamos validar apenas se tem comprimento correto e não tem muitos dígitos repetidos
        if len(set(chave)) > 5:  # Se tiver mais que 5 dígitos diferentes, é plausível
            return True
    
    # Verificar se os primeiros dígitos correspondem a uma UF válida (código entre 10 e 53)
    try:
        uf = int(chave[:2])
        if uf < 10 or uf > 53:
            return False
        
        # A data na chave (posições 2-5) deve formar um AAMM válido
        ano = int(chave[2:4])
        mes = int(chave[4:6])
        if mes < 1 or mes > 12 or ano < 0 or ano > 99:
            return False
        
        # Verificar CNPJ/CPF (posições 6-19)
        # Um CNPJ válido não pode ter todos os dígitos iguais
        cnpj = chave[6:20]
        if len(set(cnpj)) <= 2:  # Se tiver apenas 1 ou 2 dígitos diferentes, é suspeito
            return False
            
        # Verifica modelo do documento (posições 20-21)
        modelo = int(chave[20:22])
        # Modelos: 55 (NF-e), 65 (NFC-e), 57 (CT-e), 66 (NF3e para Energisa)
        modelos_validos = [55, 65, 57, 66]
        if modelo not in modelos_validos:
            return False
        
        # Verifica se não são apenas dígitos repetidos
        if len(set(chave)) <= 5:  # Se tiver 5 ou menos dígitos diferentes, é suspeito
            return False
            
        return True
    except:
        # Se houver qualquer erro de conversão, tratamos como inválido
        return False

def extrair_texto_pdf(caminho_pdf):
    """Extrai todo o texto de um arquivo PDF usando múltiplas bibliotecas."""
    texto_completo = ""
    
    # Método 1: Usar PyPDF2
    try:
        with open(caminho_pdf, 'rb') as arquivo:
            leitor = PyPDF2.PdfReader(arquivo)
            for pagina in leitor.pages:
                texto = pagina.extract_text()
                if texto:
                    texto_completo += texto + "\n"
    except Exception as e:
        print(f"Erro ao processar {caminho_pdf} com PyPDF2: {e}")
    
    # Método 2: Usar pdfplumber para extração mais detalhada
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            for pagina in pdf.pages:
                texto = pagina.extract_text()
                if texto:
                    texto_completo += texto + "\n"
                
                # Extrair tabelas pode ajudar em documentos com layout estruturado
                try:
                    tabelas = pagina.extract_tables()
                    for tabela in tabelas:
                        for linha in tabela:
                            texto_completo += " ".join([str(celula) for celula in linha if celula]) + "\n"
                except:
                    pass
    except Exception as e:
        print(f"Erro ao processar {caminho_pdf} com pdfplumber: {e}")
    
    return texto_completo

def encontrar_chaves_acesso(texto):
    """Encontra todas as potenciais chaves de acesso no texto usando diferentes padrões e as valida."""
    chaves = set()
    candidatos = set()
    
    # Encontrar chaves com 44 dígitos contínuos
    for chave in re.findall(PADRAO_CHAVE_NFE, texto):
        candidatos.add(chave)
    
    # Encontrar chaves com espaços
    for chave_com_espacos in re.findall(PADRAO_CHAVE_ESPACOS, texto):
        chave_limpa = limpar_chave(chave_com_espacos)
        candidatos.add(chave_limpa)
    
    # Encontrar chaves em blocos separados por caracteres especiais
    for chave_com_blocos in re.findall(PADRAO_CHAVE_BLOCOS, texto):
        chave_limpa = limpar_chave(chave_com_blocos)
        candidatos.add(chave_limpa)
    
    # Buscar pelo padrão específico da Energisa (vários formatos)
    # Formato 1: após "chave de acesso:"
    for parte1, parte2 in re.findall(PADRAO_ENERGISA_1, texto, re.IGNORECASE | re.DOTALL):
        chave_completa = parte1 + parte2
        chave_limpa = limpar_chave(chave_completa)
        if len(chave_limpa) == 44:
            candidatos.add(chave_limpa)
    
    # Formato 2: 10 grupos de 4 dígitos + quebra de linha + 4 dígitos
    for parte1, parte2 in re.findall(PADRAO_ENERGISA_2, texto):
        chave_completa = parte1 + parte2
        chave_limpa = limpar_chave(chave_completa)
        if len(chave_limpa) == 44:
            candidatos.add(chave_limpa)
    
    # Formato Dcelt: 11 grupos de 4 dígitos (44 dígitos total)
    for chave in re.findall(PADRAO_DCELT, texto):
        chave_limpa = limpar_chave(chave)
        if len(chave_limpa) == 44:
            candidatos.add(chave_limpa)
    
    # Buscar por sequências próximas de números que possam formar uma chave
    sequencias = re.findall(r'\b\d{8,}\b', texto)  # Sequências de pelo menos 8 dígitos
    for seq in sequencias:
        if len(seq) >= 44:
            # Extrair todas as subsequências de 44 dígitos
            for i in range(len(seq) - 43):
                candidatos.add(seq[i:i+44])
    
    # Validar todos os candidatos encontrados
    for candidato in candidatos:
        if len(candidato) == 44 and validar_chave_nfe(candidato):
            chaves.add(candidato)
    
    return list(chaves)

def processar_arquivos(pasta_entrada=None, pasta_saida_com=None, pasta_saida_sem=None, arquivo_saida=None):
    """Processa todos os arquivos PDF na pasta configurada e separa em duas pastas."""
    # Usar os parâmetros ou valores padrão
    pasta_pdfs = pasta_entrada or PASTA_PDFS
    pasta_com_chave = pasta_saida_com or PASTA_COM_CHAVE
    pasta_sem_chave = pasta_saida_sem or PASTA_SEM_CHAVE
    arquivo_resultado = arquivo_saida or ARQUIVO_SAIDA
    
    resultados = {}
    arquivos_processados = 0
    arquivos_com_chave = 0
    
    # Verifica se a pasta existe
    if not os.path.exists(pasta_pdfs):
        print(f"ERRO: A pasta {pasta_pdfs} não foi encontrada.")
        return
    
    # Criar as pastas de destino se não existirem
    os.makedirs(pasta_com_chave, exist_ok=True)
    os.makedirs(pasta_sem_chave, exist_ok=True)
    
    # Lista todos os arquivos PDF na pasta
    arquivos_pdf = [f for f in os.listdir(pasta_pdfs) 
                   if f.lower().endswith(('.pdf', '.PDF'))]
    
    if not arquivos_pdf:
        print(f"AVISO: Nenhum arquivo PDF encontrado na pasta {pasta_pdfs}")
        return
    
    print(f"Encontrados {len(arquivos_pdf)} arquivos PDF para processar.")
    
    # Processa cada arquivo
    for arquivo in arquivos_pdf:
        caminho_completo = os.path.join(pasta_pdfs, arquivo)
        print(f"Processando: {arquivo}", end="... ")
        
        try:
            # Extração de texto
            texto_completo = extrair_texto_pdf(caminho_completo)
            
            # Encontrar chaves válidas no texto
            chaves = encontrar_chaves_acesso(texto_completo)
            
            arquivos_processados += 1
            
            # Caminho para onde copiar o arquivo
            caminho_destino = ""
            
            if chaves:
                arquivos_com_chave += 1
                resultados[arquivo] = chaves
                caminho_destino = os.path.join(pasta_com_chave, arquivo)
                print(f"ENCONTRADA(S) {len(chaves)} chave(s) válida(s)!")
            else:
                caminho_destino = os.path.join(pasta_sem_chave, arquivo)
                print("Nenhuma chave válida encontrada.")
            
            # Copiar o arquivo para a pasta correspondente
            shutil.copy2(caminho_completo, caminho_destino)
                
        except Exception as e:
            print(f"ERRO ao processar {arquivo}: {str(e)}")
            print(traceback.format_exc())
    
    # Salva os resultados em um arquivo estruturado
    try:
        with open(arquivo_resultado, 'w', encoding='utf-8') as arquivo_saida:
            arquivo_saida.write(f"Data da extração: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            arquivo_saida.write(f"Total de arquivos processados: {arquivos_processados}\n")
            arquivo_saida.write(f"Arquivos com chaves encontradas: {arquivos_com_chave}\n\n")
            
            # Cabeçalho do relatório
            arquivo_saida.write("Empresa;Numero Doc;Filial;Nome Arquivo;Chave de Acesso\n")
            arquivo_saida.write("-" * 100 + "\n")
            
            for nome_arquivo, chaves in resultados.items():
                # Extrair informações do nome do arquivo
                partes = nome_arquivo.split('_')
                empresa = partes[0] if len(partes) > 0 else ""
                numero_doc = partes[1] if len(partes) > 1 else ""
                
                # Obter filial (entre o terceiro _ e o .pdf)
                if len(partes) > 3:
                    # Pegar a parte após o terceiro underscore (índice 3)
                    filial = partes[3]
                    # Remover extensão .pdf caso esteja presente nessa parte
                    filial = filial.replace(".pdf", "").replace(".PDF", "")
                else:
                    filial = ""
                
                # Escrever cada chave em uma linha do relatório
                for chave in chaves:
                    arquivo_saida.write(f"{empresa};{numero_doc};{filial};{nome_arquivo};{chave}\n")
                
                # Se não houver chaves, ainda registrar o arquivo no relatório
                if not chaves:
                    arquivo_saida.write(f"{empresa};{numero_doc};{filial};{nome_arquivo};NENHUMA CHAVE ENCONTRADA\n")
        
        print("\nResumo:")
        print(f"Total de arquivos processados: {arquivos_processados}")
        print(f"Arquivos com chaves encontradas: {arquivos_com_chave}")
        print(f"Arquivos com chave copiados para: {pasta_com_chave}")
        print(f"Arquivos sem chave copiados para: {pasta_sem_chave}")
        print(f"Resultados salvos em: {arquivo_resultado}")
    except Exception as e:
        print(f"ERRO ao salvar resultados: {str(e)}")

def mostrar_ajuda_interativa():
    """Mostra uma tela de ajuda interativa se o programa for executado sem argumentos"""
    mensagem = """
=============================================================
EXTRATOR DE CHAVES DE ACESSO DE DOCUMENTOS FISCAIS
=============================================================

Este programa extrai chaves de acesso de documentos fiscais em formato PDF.
Ele identifica diferentes formatos de chaves, incluindo:

- Chaves padrão de NFe (44 dígitos contínuos)
- Formato Energisa (10 grupos de 4 + quebra de linha + 4 dígitos)
- Formato Dcelt (11 grupos de 4 dígitos)
- Outros formatos com separadores como espaços, pontos ou traços

Modo básico de uso:
1. Execute o programa na pasta onde estão os PDFs
2. Os arquivos serão processados automaticamente
3. Os resultados serão salvos e os arquivos separados em pastas
"""

    print(mensagem)
    
    # Se estiver em modo gráfico (executável compilado com --noconsole)
    if IS_FROZEN and TKINTER_DISPONIVEL:
        root = tk.Tk()
        root.withdraw()  # Esconder a janela principal
        messagebox.showinfo("Extrator de Chaves de NFe", mensagem)
    else:
        # Em modo console normal
        input("Pressione ENTER para continuar...")


def configurar_argumentos():
    """Configura os argumentos da linha de comando"""
    parser = argparse.ArgumentParser(description='Extrator de chaves de acesso de documentos fiscais em PDF')
    
    parser.add_argument('-i', '--input', dest='pasta_entrada',
                        help='Pasta onde estão os PDFs (padrão: pasta atual)')
    
    parser.add_argument('-c', '--com-chave', dest='pasta_com_chave',
                       help='Pasta para armazenar PDFs com chave (padrão: "PDFs_Com_Chave")')
    
    parser.add_argument('-s', '--sem-chave', dest='pasta_sem_chave',
                       help='Pasta para armazenar PDFs sem chave (padrão: "PDFs_Sem_Chave")')
    
    parser.add_argument('-o', '--output', dest='arquivo_saida',
                       help='Arquivo de saída com as chaves (padrão: "chaves_extraidas_final.txt")')
    
    parser.add_argument('--ajuda', action='store_true', dest='mostrar_ajuda',
                       help='Mostra mensagem de ajuda detalhada')
    
    # Se não foram fornecidos argumentos e não estiver em modo congelado, mostrar ajuda interativa
    if len(sys.argv) == 1 and not IS_FROZEN:
        mostrar_ajuda_interativa()
    # Se estiver em modo congelado (executável), não mostrar ajuda interativa automaticamente
    # para evitar problemas com stdin
    
    return parser.parse_args()

if __name__ == "__main__":
    try:
        print("\nEXTRATOR DE CHAVES DE ACESSO DE DOCUMENTOS FISCAIS")
        print("-" * 50)
        
        # Configurar argumentos da linha de comando
        args = configurar_argumentos()
        
        # Iniciar processamento
        print("Iniciando extração avançada de chaves de acesso NFe válidas...")
        processar_arquivos(
            pasta_entrada=args.pasta_entrada,
            pasta_saida_com=args.pasta_com_chave,
            pasta_saida_sem=args.pasta_sem_chave,
            arquivo_saida=args.arquivo_saida
        )
        print("Processo concluído!")
        
        # Aguardar confirmação do usuário antes de fechar (para uso como executável)
        if IS_FROZEN and TKINTER_DISPONIVEL:
            # Se estiver em modo gráfico, mostrar mensagem
            root = tk.Tk()
            root.withdraw()
            messagebox.showinfo("Processo Concluído", "Extração de chaves concluída com sucesso!\n\nResumo:\n" +
                              f"- {arquivos_processados} arquivos processados\n" +
                              f"- {arquivos_com_chave} arquivos com chaves\n" +
                              f"Resultados salvos em: {arquivo_resultado}")
        else:
            # Em modo console normal
            input("\nPressione ENTER para sair...")
        
    except Exception as e:
        print(f"\nERRO FATAL: {str(e)}")
        print("Para mais detalhes:")
        traceback.print_exc()
        if IS_FROZEN and TKINTER_DISPONIVEL:
            # Se estiver em modo gráfico, mostrar erro em uma caixa de diálogo
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Erro Fatal", f"Ocorreu um erro: {str(e)}")
        else:
            # Em modo console normal
            input("\nPressione ENTER para sair...")
        sys.exit(1)
