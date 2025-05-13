# Extrator de Chaves de Acesso NFe

Ferramenta para extração automática de chaves de acesso de documentos fiscais em PDF, com suporte a múltiplos formatos e geração de relatório estruturado.

## Funcionalidades

- **Extração avançada**: Identifica chaves de acesso em diversos formatos (NFe, NFCe, CTe)
- **Validação inteligente**: Implementa validações para evitar falsos positivos
- **Processamento em lote**: Processa vários arquivos PDF de uma vez
- **Separação automática**: Organiza os PDFs em pastas (com/sem chave)
- **Relatório estruturado**: Gera relatório com dados extraídos dos nomes dos arquivos
- **Interface híbrida**: Funciona tanto em modo console quanto com interface gráfica

## Formatos suportados

- Chaves padrão NFe (44 dígitos contínuos)
- Formato Energisa (10 grupos de 4 + quebra de linha + 4 dígitos)
- Formato Dcelt (11 grupos de 4 dígitos)
- Outros formatos com separadores como espaços, pontos ou traços

## Requisitos

- Python 3.6 ou superior
- Bibliotecas: PyPDF2, pdfplumber, Pillow (PIL)

## Instalação

### Usando o executável (Windows)

1. Baixe o executável mais recente da [página de releases](https://github.com/gui3g/chave.pdf.energia/releases)
2. Execute o arquivo `.exe` diretamente

### Instalando a partir do código-fonte

```bash
# Clone o repositório
git clone https://github.com/gui3g/chave.pdf.energia.git
cd chave.pdf.energia

# Instale as dependências
pip install PyPDF2 pdfplumber Pillow
```

## Modo de uso

### Usando o executável

1. Coloque o executável na pasta onde estão os PDFs que deseja processar
2. Execute o programa
3. Os arquivos serão processados e separados em duas pastas:
   - `PDFs_Com_Chave`: Arquivos com chave de acesso encontrada
   - `PDFs_Sem_Chave`: Arquivos sem chave de acesso
4. Um relatório será gerado no arquivo `chaves_extraidas_final.txt`

### Usando o script Python

```bash
python extrair_chaves_final.py
```

### Opções de linha de comando

```
-i, --input      Pasta onde estão os PDFs (padrão: pasta atual)
-c, --com-chave  Pasta para armazenar PDFs com chave (padrão: "PDFs_Com_Chave")
-s, --sem-chave  Pasta para armazenar PDFs sem chave (padrão: "PDFs_Sem_Chave")
-o, --output     Arquivo de saída com as chaves (padrão: "chaves_extraidas_final.txt")
--ajuda          Mostra mensagem de ajuda detalhada
```

## Formato de relatório

O relatório gerado contém as seguintes colunas separadas por ponto-e-vírgula (;):

- **Empresa**: Texto do início do nome do arquivo até o primeiro underscore
- **Numero Doc**: Texto entre o primeiro e segundo underscore
- **Filial**: Texto entre o terceiro underscore e antes da extensão .pdf
- **Nome Arquivo**: Nome completo do arquivo processado
- **Chave de Acesso**: Chave de acesso NFe encontrada no arquivo

Exemplo:
```
Empresa;Numero Doc;Filial;Nome Arquivo;Chave de Acesso
----------------------------------------------------------------------
Energisa;9012820;151;Energisa_9012820_Lac_151.pdf;50250215413826000150660020090128201010035662
Engie;161922;142;Engie_161922_Lac_142.pdf;42250304100556000100550010001619221102627515
```

## Compilando o executável (Windows)

Para compilar o script em um executável para Windows:

```bash
# Instale o PyInstaller
pip install pyinstaller

# Compile o programa (com interface gráfica)
pyinstaller --onefile --noconsole --name="Extrator_NFe" extrair_chaves_final.py

# Ou compile mantendo o console visível (alternativa)
pyinstaller --onefile --name="Extrator_NFe" extrair_chaves_final.py
```

O executável gerado estará na pasta `dist`.

## Licença

Este projeto está licenciado sob a licença MIT.
