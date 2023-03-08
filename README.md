# Test Generator
Programa em Python que uma monta provas a partir de uma planilha a qual contém as questões. 
Realiza o embaralhamento das questões e das opções.


## Instalação dos pacotes necessários para a execução do script:

    pip install z3c.rml
    pip install openpyxl
    pip install PyPDF2

## Comando para visualizar a versão do script:

    python3 ./assembly.py -v
    
## Comando para gerar uma prova:

#### Planilha no mesmo diretório que o script:

    python3 ./assembly.py 

#### Planilha em um diretório diferente:

    python3 ./assembly.py -t questions.xlsx
  
 
