
# Script para gerar testes embaralhando as questoes e opcoes  
# ----------------------------------------------------------

# PDFTK para MAC e Windows
# * https://www.pdflabs.com/tools/pdftk-the-pdf-toolkit/pdftk_server-2.02-mac_osx-10.11-setup.pkg
# * https://www.pdflabs.com/tools/pdftk-the-pdf-toolkit/
# * https://portableapps.com/apps/office/pdftk_builder_portable

# Script usa o pacote Z3C.rml; Adiciona o comando rml2pdf
# pip3 install z3c.rml

# Bibliotecas
# -----------
from os import name
from sys import stderr
from telnetlib import STATUS
import openpyxl    
from openpyxl.styles import Alignment, Font, Color, colors
import random
import subprocess
import os 
import codecs
import getopt
import sys 

from PyPDF2 import PdfFileReader, PdfFileWriter

# Variáveis
# ---------

# Versao 1.0
VER = "1.0" 

# Define o caminho de alguns comandos
if os.name == 'nt':
    CMD_TRMLPDF="rml2pdf.exe"
    CMD_PDFTK="C:\\Users\\kwecko\\Documents\\Aplicativos\\PDFTKBuilderPortable\\App\\pdftkbuilder\\pdftk.exe"
else :
    CMD_TRMLPDF="/Library/Frameworks/Python.framework/Versions/3.10/bin/rml2pdf"
    CMD_PDFTK="/usr/local/bin/pdftk"

# Nome do arquivo PDF gerado
OUTPDF="Prova.pdf"

# Campos da Tabela
TEMP = []
QUESTIONS = []
VALORES_Q = []

# Armazena a Planilha com as Questoes
input_table = None

# Variavel global que armanezara as estrutura da prova
PDF = ""

# Estrutura do Cabeçalho do arquivo
PDF_HEADER =  """\
<!DOCTYPE document SYSTEM "rml_1_0.dtd"> 
<document filename="%PROVA%.pdf" invariant="1">

<template pageSize="A4" leftMargin="70" showBoundary="0"
    author="Marcelo Kwecko"
    subject="Prova"
    title="Prova"
    creator="Marcelo Kwecko"
	displayDocTitle="1"
	lang="pt-BR"
    >
	<!-- Modelo principal do doc com cabelhaco --> 
	<pageTemplate id="main" pageSize="A4">
		<pageGraphics>
		<setFont name="Helvetica-Bold" size="13"/> <drawString x="%COORDENADA%" y="800"> 
				%CURSO%  </drawString>
			<setFont name="Helvetica-Bold" size="10"/> <drawString x="%COORDENADA%" y="790"> 
				%DISCIPLINA% </drawString>
			<setFont name="Helvetica" size="10"/> <drawString x="%COORDENADA%" y="780"> 
				Marcelo Kwecko - (marcelokwecko@ifsul.edu.br) </drawString>
            	
			<image file="if.png" preserveAspectRatio="1" x="20" y="770" width="140" /> 
    		<image file="%LOGO%" preserveAspectRatio="1" x="500" y="760" height="72" /> 
			
            <!-- Linha -->
			<lines> 20 756 570 756 </lines>
			<lines> 20 755 570 755 </lines>
			
            
			<textAnnotation><param name="Rect">0,0,1,1</param><param name="F">3</param><param name="escape">6</param>X::PDF
			PX(S)
			MT(PINK)
			</textAnnotation>
			
			<!-- Inclui o numero da pag -->
			<drawRightString x="550" y="40"><pageNumber/></drawRightString>

	        </pageGraphics>

	        <!-- Define a area util do documento -->
	        <frame id="1" x1="35" y1="80" width="535" height="660"/>
	
    </pageTemplate>

    <!-- Template que define a pagina em branco -->
	<pageTemplate id="second" pageSize="A4">
		<frame id="2" x1="35" y1="80" width="535" height="660"/>
	</pageTemplate>

</template>

<!-- Estilos de Fontes usadas ao longo do texto -->
<stylesheet>
	<initialize>
	<alias id="style.normal" value="style.Normal"/>
	</initialize>
	
    <paraStyle name="italic" fontName="Helvetica-Oblique" fontSize="12" leading="12"/>
	
    <paraStyle name="normal" fontName="Helvetica" fontSize="12" leading="12" hyphenationLang="pt_BR" hyphenationMinWordLength="5" spaceAfter="4"/>
	
    <paraStyle name="normalb12" fontName="Helvetica-Bold" fontSize="12" leading="12" hyphenationLang="pt_BR" hyphenationMinWordLength="5"/>
    
    <paraStyle name="normalb14" fontName="Helvetica-Bold" fontSize="14" leading="14" hyphenationLang="pt_BR" hyphenationMinWordLength="5"/>
	
    <paraStyle name="questoes" fontName="Helvetica" fontSize="12" leading="12" spaceBefore="12" spaceAfter="6" alignment="justify"/>
    
    <paraStyle name="coluna" fontName="Helvetica" fontSize="12" leading="12"  spaceBefore="6" spaceAfter="6"/>
    
    <paraStyle name="normal_left" alignment="left" fontName="Helvetica" fontSize="12" leading="12" hyphenationLang="pt_BR" hyphenationMinWordLength="5" spaceAfter="4"/>
    
    <paraStyle name="normal_right" alignment="right" fontName="Helvetica" fontSize="12" leading="12" hyphenationLang="pt_BR" hyphenationMinWordLength="5" spaceAfter="4"/>
    
    <paraStyle name="normal_center" alignment="center" fontName="Helvetica" fontSize="12" leading="12" hyphenationLang="pt_BR" hyphenationMinWordLength="5" spaceAfter="4"/>
    
    <paraStyle name="normal_justify" alignment="justify" fontName="Helvetica" fontSize="12" leading="12" hyphenationLang="pt_BR" hyphenationMinWordLength="5" spaceAfter="6"/>


    <!-- Estilo da Tabela da Questao  -->

    <blockTableStyle id="COLUNA">
		<!-- Config das fontes -->
		<blockFont name="Helvetica" size="12" leading="12"/>
		<blockTextColor colorName="black"/>

		<blockFont name="Helvetica-Bold" size="10" start="1,0" stop="1,-1"/>

		<!-- Alinhamento -->
        <blockAlignment value="LEFT"/>
		<blockValign value="MIDDLE" start="0,0" stop="0,-1"/>

		<blockAlignment value="RIGHT" start="1,0" stop="1,-1"/>
        <blockValign value="MIDDLE" start="1,0" stop="1,-1"/>
		
        <!-- Imprime as bordas/linhas da tabela -->
		<!-- <lineStyle kind="GRID" colorName="darkblue"/> -->
	</blockTableStyle>

</stylesheet>

<story>

	<!-- Cabeçalho para preenchimento das informacoes dos Alunos --> 
    <storyPlace x="35" y="675" width="525" height="73" origin="page">
		
        <!-- Nome -->
        <para style="normalb12">Nome: </para>
		
        <!-- Linha do Nome  -->
		<illustration width="1" height="2">
			<rect x="40" y="0" width="490" height="0.5" fill="yes" stroke="yes"/>
		</illustration>
 
		<!-- Data e Nota -->
		<blockTable colWidths="14cm,4cm">
			<tr>
				<td><para style="normalb12">Data: </para></td><td><para style="normalb12">Nota: </para></td>
			</tr>
		</blockTable>
		
		<!-- Linhas da Data e Nota -->
		<illustration width="1" height="2">
			<rect x="40" y="0" width="100" height="0.5" fill="yes" stroke="yes"/>
			<rect x="437" y="0" width="93" height="0.5" fill="yes" stroke="yes"/>
		</illustration>
 
		<!-- Espaco -->
		<hr color="white" thickness="4pt"/>

		<!-- Descricao da Prova -->
		<para style="normalb14" alignment="Center"> %DESCR% </para>
        <hr color="white" thickness="6pt"/>

	</storyPlace>
	
	<!-- Espaco necessario para livrar o cabecalho da prova -->
	<spacer length="45"/>

    <!-- Corpo do Documento --> 

"""

# Funcoes 
##########

# Conversao de Decimal para Romano
# ++++++++++++++++++++++++++++++++
def Int2Roman(num):
    val = [
            1000, 900, 500, 400,
            100, 90, 50, 40,
            10, 9, 5, 4,
            1
        ]
    syb = [
            "M", "CM", "D", "CD",
            "C", "XC", "L", "XL",
            "X", "IX", "V", "IV",
            "I"
        ]
    roman_num = ''
    i = 0
    while  num > 0:
        for _ in range(num // val[i]):
            roman_num += syb[i]
            num -= val[i]
        i += 1
    
    return roman_num

# Formata o numero do valor da questao 
# ++++++++++++++++++++++++++++++++++++
def FormatNumber(n):
    if float(n).is_integer():
        return format(n,'.1f')
    else :
        return format(n,'.2f').rstrip('0').rstrip('.')
    
# Formata as questoes tipo Discursiva
# ++++++++++++++++++++++++++++++++++++
def fun_discursiva(N_Q, Q):
    global PDF
    PDF = PDF + "\n <para style=\"questoes\" hyphenationLang=\"pt_BR\"> " + str(N_Q) + ") " + Q[2] + " (" +  FormatNumber(Q[1]) + ") </para> \n "
    PDF = PDF + "<!-- Espaco --> \n <hr color=\"white\" thickness=\"6pt\"/> \n " 

# Formata as questoes do tipo Coluna
# ++++++++++++++++++++++++++++++++++++
def fun_coluna(N_Q, Q):
    
    #* Variaveis
    global PDF
    COL_A = []
    COL_B = []
    COLUNA01 = []
    COLUNA02 = []
    
    PDF = PDF + "\n <para style=\"questoes\" hyphenationLang=\"pt_BR\"> " + str(N_Q) + ") " + Q[2] + " (" +  FormatNumber(Q[1]) + ")  </para>\n "
    PDF = PDF + "<!-- Espaco --> \n <hr color=\"white\" thickness=\"2pt\"/> \n "
    
    for i in range(3, len(Q)):
        if "C:" in Q[i]:
            COL_A.append(Q[i].replace('C:', ''))
        else:
            COL_B.append(Q[i].replace('R:', ''))
   
    # Sorteio da ordem da 1 Coluna
    l = list(range(len(COL_A)))
    random.shuffle(l)

    ch = 'a'
    L = chr(ord(ch))
    for n in l:
        if COL_A[n]: 
            COLUNA01.append(L + ") " + COL_A[n])
            L = chr(ord(L) + 1)
    
    # Sorteio da ordem da 2 Coluna 
    l = list(range(len(COL_B)))
    random.shuffle(l)

    for n in l:
        if COL_B[n]:
            # &#160; Codigo HTML do espaco; <xpre> intepreta esses codigos 
            #COLUNA02.append("<xpre> (&#160;&#160;&#160;&#160;) </xpre> " + COL_B[n])
            COLUNA02.append(COL_B[n])


    # Verifica os tamanhos das duas colunas. 
    # Em caso de tamanho diferentes preenche com Vazio, deixando ambas com o mesmo tam.
    if len(COLUNA01) !=  len(COLUNA02) :
        if len(COLUNA01) > len(COLUNA02) :
            for n in range(len(COLUNA02), len(COLUNA01)) :
               COLUNA02.append("")
        else:
            for n in range(len(COLUNA01), len(COLUNA02)) :
               COLUNA01.append("")
    
    # Adiciona as Colunas ao Documento
    PDF= PDF + "<blockTable colWidths=\"4cm,2cm,12cm\" style=\"COLUNA\" > \n"

    for n in range(0,len(COLUNA01)):
        
        PDF = PDF + "\t\t <tr><td>" + COLUNA01[n] + "</td> \t"

        # Testa se a item na 2 coluna. Caso nao tenha remove os ( )  
        if COLUNA02[n] != "" :
            PDF = PDF + "<td>  (     ) </td> \t"
        else :
             PDF = PDF + "<td> </td> \t"
        
        PDF = PDF + "<td><para style=\"coluna\">" + COLUNA02[n] + "</para></td></tr> \n"
    
    PDF = PDF + "</blockTable> \n"

    # Adiciona um espaco
    PDF = PDF + "<!-- Espaco --> \n <hr color=\"white\" thickness=\"4pt\"/> \n"

# Formata as questoes Objetiva, tipo 1
# ++++++++++++++++++++++++++++++++++++
def func_objetiva01(N_Q, Q):
    #* Variaveis
    global PDF
    AFIRMATIVA = []

    PDF = PDF + "\n <para style=\"questoes\" hyphenationLang=\"pt_BR\"> " + str(N_Q) + ") " + Q[2] + " (" +  FormatNumber(Q[1]) + ")  </para>\n "
    PDF = PDF + "<!-- Espaco --> \n <hr color=\"white\" thickness=\"2pt\"/> \n "

    # Sorteio da ordem da 2 Coluna 
    l = list(range(3,len(Q)))
    random.shuffle(l)

    for n in l:
        AFIRMATIVA.append("<xpre> (&#160;&#160;&#160;&#160;) </xpre> \t" + Q[n])

    for n in range(0,len(AFIRMATIVA)):
        PDF = PDF + "\n <para style=\"normal_justify\" hyphenationLang=\"pt_BR\"> " + AFIRMATIVA[n]  + " </para> \n"
        
    # Adiciona um espaco
    PDF = PDF + "<!-- Espaco --> \n <hr color=\"white\" thickness=\"4pt\"/> \n "

# Formata as questoes Objetiva, tipo 2
# ++++++++++++++++++++++++++++++++++++
def func_objetiva02(N_Q, Q):
    
    #* Variaveis
    global PDF
    AFIRMATIVA = []
    T_AFIRMATIVA = []
    OPCOES = []
    T_OPCOES = []

    PDF = PDF + "\n <para style=\"questoes\" hyphenationLang=\"pt_BR\"> " + str(N_Q) + ") " + Q[2] + " (" +  FormatNumber(Q[1]) + ")  \n </para>"
    PDF = PDF + "<!-- Espaco --> \n <hr color=\"white\" thickness=\"2pt\"/> \n "

    #Procura por Opcoes de resposta, caso contrario é uma afirmação/negação
    for i in range(3, len(Q)):
        if "R:" in Q[i]:
            # &#160; Codigo HTML do espaco; <xpre> intepreta esses codigos 
            T_OPCOES.append("<xpre> (&#160;&#160;&#160;&#160;) </xpre>" + Q[i].replace('R:', ''))
        else:
            T_AFIRMATIVA.append(Q[i])

    # Sorteio da Afirmativas/Negativas
    la = list(range(len(T_AFIRMATIVA)))
    random.shuffle(la)

    for n in la:
        AFIRMATIVA.append(T_AFIRMATIVA[n])

    # Sorteio da ordem das Opcoes 
    l = list(range(len(T_OPCOES)))
    random.shuffle(l)

    for n in l:
        OPCOES.append(T_OPCOES[n])

    # Adiciona as Afirmativas e Negativas ao documento
    for n in range(0,len(AFIRMATIVA)):
        PDF = PDF + "\n <para style=\"normal_justify\" hyphenationLang=\"pt_BR\"> " + Int2Roman(n+1) + " - " + AFIRMATIVA[n]  + " </para> \n"
    
    # + Coloca as opcoes uma abaixo da outra
    #PDF = PDF + "<spacer length=\"8\"/>"
    #for n in range(0,len(OPCOES)):
    #    PDF = PDF + "\n <para style=\"normal\" hyphenationLang=\"pt_BR\"> " + OPCOES[n]  + " </para> \n"

    # Coloca as opcoes na forma de duas coluna
    PDF= PDF + "<spacer length=\"2\"/> \n"
    PDF= PDF + "<blockTable alignment=\"LEFT\" colWidths=\"8cm,8cm\"> \n"

    for n in range(0,len(OPCOES),2):
        if n+1 < len(OPCOES): 
            PDF = PDF + "<tr><td><para style=\"questoes\">  " + OPCOES[n] + " </para></td> \n <td><para style=\"normal\"> " + OPCOES[n+1] + "</para></td></tr> \n"
        else : 
            PDF = PDF + "<tr><td><para style=\"questoes\">  " + OPCOES[n] + " </para> \n </td><td><para style=\"normal\">  </para></td></tr> \n"
    
    PDF = PDF + "</blockTable> \n"
    
    # Adiciona um espaco
    PDF = PDF + "<!-- Espaco --> \n <hr color=\"white\" thickness=\"4pt\"/> \n "


# Principal
###########

# lê da console o tabela com a estrutura do teste
options, args = getopt.gnu_getopt(sys.argv[1:], 't:v', ['tabela =', 
                                                            'version',
                                                            ])
# Analisa os parametros
for opt, arg in options:
    if opt in ['-t', '--tabela']:
        input_table = arg
    elif opt in ['-v', '--version'] :
        print("Versao: " + VER)
        exit(0)

# Caso a Planilha nao seja declarada como parametro, usa o nome padrao no diretorio atual  
if input_table is None:
    input_table="questions.xlsx"

# Caso a Tabela informada nao exista 
if 'input_table' not in locals():
     print(sys.argv[0], " -t <tabela com as questoes>") 
     exit()

# Checa se a planilha existe realmente
file_exists = os.path.exists(input_table)

if not file_exists:
    print("Arquivo com questões informado, nao encontrado! - " + input_table)
    exit(2)

#Checa se o arquivo PDF com um "teste" ja existe, e se sim, apaga-o 
file_exists = os.path.exists(OUTPDF)

if  file_exists:
    print("Arquivo " + OUTPDF + " já existe e será excluído!")
    if os.name == 'nt':
        os.system("del " + OUTPDF)
    else :
        os.system("rm " + OUTPDF)

# Abrir a planilha
planilha = openpyxl.load_workbook(input_table)

# Definição das Paginas 
questions_page = planilha['questoes']

# Entradas 
OP_CURSO=input("Entre com o curso [tinf/tads]: ")

if OP_CURSO == "tinf":
    PDF_HEADER = PDF_HEADER.replace('%CURSO%','CURSO TÉCNICO EM INFORMÁTICA')
    PDF_HEADER = PDF_HEADER.replace('%COORDENADA%','200')
    PDF_HEADER = PDF_HEADER.replace('%LOGO%','tinf.png')
elif OP_CURSO == "tads":
    PDF_HEADER = PDF_HEADER.replace('%CURSO%','TÉCNOLOGO em ANÁL. e DESENV. de SISTEMA')
    PDF_HEADER = PDF_HEADER.replace('%COORDENADA%','170')
    PDF_HEADER = PDF_HEADER.replace('%LOGO%','tads.png')
else:
    if OP_CURSO:
        print("Opcao ", OP_CURSO, " inválida!")
    else :
        print("Opcao inválida!")
    exit(1)

# Entradas do Curso e Nome da Disciplina
print("Entre com o nome da disciplina!")
print("  1 - Redes I \n  2 - Redes II \n  3 - Administração e Segurança em Redes")
OP_DISC=input("Opção: ")

if OP_DISC == "1":
    PDF_HEADER = PDF_HEADER.replace('%DISCIPLINA%','Disciplina de Redes de Computadores I')    
elif OP_DISC == "2":
    PDF_HEADER = PDF_HEADER.replace('%DISCIPLINA%','Disciplina de Redes de Computadores II')
elif OP_DISC == "3":
    PDF_HEADER = PDF_HEADER.replace('%DISCIPLINA%','Disc. de Admin. e Seg. em Redes de Computadores')    
else: 
    if OP_DISC:
        print("Opcao ", OP_DISC, " inválida!")
    else :
        print("Opcao inválida!")
    exit(1)

OP_DESC=input("Entre com a Descrição da prova [Prova 1 Etapa]: ")

# Pega as questoes de cada linha da tabela e armazena na tupla
for rows in questions_page.iter_rows(min_row=2, max_row=20):
        if rows[0].value :
            TEMP = []
            VALORES_Q.append(rows[1].value)
            for x in range(0, len(rows)):
                if rows[int(x)].value == None:
                    break 
                TEMP.append(rows[x].value)        
            QUESTIONS.append(TEMP)

# Calcula o valor total da Prova
i = 0
for v in VALORES_Q: 
    i = i + v

print("Valor total da prova: " + "{:.2f}".format(i))

# Atualiza a descrição com a nota
PDF_HEADER = PDF_HEADER.replace('%DESCR%', OP_DESC + " (" +   FormatNumber(i) + ") ")

# Le a quantidade de provas
N_TEST=input("Entre com a quantidade de provas: ")

if not N_TEST.isdigit():
    print("Não entrou com um numero!")
    exit(2)


# Gera a quantidade de provas estabelecidas 
for q in range(int(N_TEST)):

    # Copia a estrutura do cabecalho da prova
    PDF = PDF_HEADER[:]

    # Determina com sera o nome do arquivo gerado em PDF 
    PDF = PDF.replace('%PROVA%', str(q))

    # Sorteio da ordem das questoes 
    l = list(range(len(QUESTIONS)))
    random.shuffle(l)

    # Analisa as questoes com base no sorteio
    i = 0
    for n in l:
        # Caso nao haja espaço na folha, no caso 6.8cm, realiza uma quebra de pagina
        PDF = PDF + "\n <!-- Quebra de Pag caso necessario --> \n <condPageBreak height=\"6.8cm\"/> \n"
        i += 1
        if QUESTIONS[n][0] == "discursiva":
            fun_discursiva(i,QUESTIONS[n])
        elif QUESTIONS[n][0] == "coluna":
            fun_coluna(i,QUESTIONS[n])
        elif QUESTIONS[n][0] == "objetiva01":
            func_objetiva01(i,QUESTIONS[n])
        elif QUESTIONS[n][0] == "objetiva02":
            func_objetiva02(i,QUESTIONS[n])

     
    # Finaliza o DOC PDF
    PDF = PDF + """\

<!-- Codigo que adiciona mais um pagina no caso do n de paginas serem impar --> 
<docAssign var='i' expr="doc.page"/>
<docAssign var='x' expr="doc.page"/>
<docAssign var='y' expr="doc.page"/>

<!-- Calcula se o n de pag e par ou impar-->
<docExec stmt='i=i%2'/>

<!-- Caso o n de pag seja maior que 1 -->
<docIf cond="x&gt;1">
   
    <!-- Caso o n de pag seja impar, acrescenta o cod -->
	<docIf cond="i&gt;0">
        
        <!-- Define o template em branco  e adiciona um espaco-->
		<setNextTemplate name="second" />
		<para> <xpre> &#160; </xpre> </para>

        <!-- Adiciona um espaco ate atingir a proxima pagina -->
		<docExec stmt='y+=1'/>
		<docWhile cond='y&gt;doc.page'>
			<para> <xpre> &#160; </xpre> </para>
		</docWhile>
        
	</docIf>
</docIf>

</story> 
</document> 
"""

    # Criar o Arquivo 
    text_file = codecs.open("file.rml", "w", "utf-8")
    text_file.write(PDF)
    text_file.close()

    # Gera o PDF - Aplicativo rml2pdf e o arquivo RML gerando 
    STATUS_CMD = subprocess.run([CMD_TRMLPDF, "file.rml"], stderr=subprocess.DEVNULL, check=True)

    if STATUS_CMD.returncode == 0 : 
       print("PDF " + str(q+1) + " criado com Sucesso!!")
    else :
        print("Erro ao cria o PDF!!")
        exit(3)


# Concatenar as provas em apenas um arquivo; aplicativo pdftk
if os.name == 'nt':
    STATUS_CMDPDF = os.system(CMD_PDFTK + " ??.pdf cat output " + OUTPDF)
else :
    STATUS_CMDPDF = os.system(CMD_PDFTK + " *[0-9].pdf cat output " + OUTPDF)

if STATUS_CMDPDF == 0 : 
       print("PDFs concatenados com Sucesso!!")
else :
        print("Erro ao concatenar os PDF!!")
        exit(4)

# Apaga os arquivos temporários. RML e PDF
if os.name == 'nt':
    os.system("del file.rml")
    os.system("del ??.pdf")
else :
    os.system("rm *[0-9].pdf")
    os.system("rm file.rml")

# Le o PDF e grava novamente; Reduz o risco de erros no momento da impressao
pdf = PdfFileReader(OUTPDF, 'rb')
pdfwrite = PdfFileWriter()

# Informa o numero de pag do documento
print("Numero de paginas: " + str(pdf.getNumPages()))

# Le todas as paginas do arquivo de entrada
for page in range(pdf.getNumPages()):
    pdfpage = pdf.getPage(page)
    pdfwrite.addPage(pdfpage)

# Gera um novo PDF
with open(OUTPDF, 'wb') as fw:
    pdfwrite.write(fw)

# Finaliza o Script
exit(0)