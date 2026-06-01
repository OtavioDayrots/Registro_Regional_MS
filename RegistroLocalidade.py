# pyinstaller RegistroLocalidade.py

# Conector de sql recomendado na documentação oficial.
from turtle import color
import pyodbc
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.styles.numbers import NumberFormatList
from openpyxl.styles import Font, Fill
from openpyxl.styles import PatternFill
from openpyxl.comments import Comment

def createWorkBook(nomeR, regional):
	if('BOLSAO/PARANAIBA' in nomeR):
		createSheetLocal('PNB', regional)
	elif('GRANDE DOURADOS' in nomeR):
		createSheetLocal('DOS', regional)
	elif('NORTE' in nomeR):
		createSheetLocal('CXM', regional)
	elif('JIM' in nomeR):
		createSheetLocal('JIM', regional)
	elif('CONE-SUL' in nomeR):
		createSheetLocal('NVR', regional)
	elif('SUL/FRONTEIRA'in nomeR):
		createSheetLocal('PPR', regional)
	elif('LESTE' in nomeR):
		createSheetLocal('NDI', regional)
	elif('PANTANAL/CORUMBA' in nomeR):
		createSheetLocal('CMA', regional)
	elif('PANTANAL/AQUIDAUANA' in nomeR):
		createSheetLocal('AUA', regional)
	elif('BOLSAO/TRES LAGOAS' in nomeR):
		createSheetLocal('TLS', regional)

def createSheetLocal(nomeR, regional):
    wb = Workbook()

    # Listas de colunas necessárias para abrir no excel
    col = [
        "DATA",
        "RDA",
        "RCE - Sanesul",
        "RCE - Oneroso Terceiros",
        "RCE - Oneroso Ms-Pantanal",
        "ECON",
        "PENDENTES",
    ]
    com = [
            "DATA - Mês e Ano",
            "RDA - Rede de Distribuição de Água - Extensão de Rede de água com status da rede = ""existente""",
            "RCE - Sanesul - Rede coletora de Esgoto Sanesul - Extensão de Rede de Esgoto com status da rede = ""existente""",
            "RCE - Oneroso Terceiros - Rede coletora de Esgoto Terceiros - Extensão de Rede de Esgoto com status da rede = ""Oneroso Terceiros""",
            "RCE - Oneroso MS-Pantanal - Rede coletora de Esgoto MS-Pantanal - Extensão de Rede de Esgoto com status da rede = ""Oneroso MS-Pantanal""",
            "ECON - Economia - Total de Matriculas Georreferenciadas",
            "PENDENTES - Diferença entre o Total de Matriculas Cadastrais da Comercial e o Total de Matriculas Georreferenciados"  ,
        ]
    groupLocal = regional.groupby('LOCALIDADE')

    for nomeL, localidade in groupLocal:
        # graphImg = plotGraph(nameL, localidade)
        ws = wb.create_sheet(nomeL, -1)
        # ws.add_image(graphImg, 'F2')

        for r in dataframe_to_rows(localidade[col], index=False, header=True):
            ws.append(r)

        # Adiciona comentários nas células de cabeçalho
        for idx, texto_comentario in enumerate(com, start=1):
              comment = Comment(texto_comentario, "INFO", width=300, height=100)
              ws.cell(row=1, column=idx).comment = comment

        NumberFormatList()

        for cell in ws['A']:
            cell.number_format = 'mmm-yy'
            # cell.number_format = builtin_format_code(17)

        n = len(localidade) + 1
        row_n = str(n)
        saldo = str(n + 1)

        if ws['B2'].value:
            ws['B2'].value = 0
        if ws['C2'].value:
            ws['C2'].value = 0
        if ws['D2'].value:
            ws['D2'].value = 0
        if ws['E2'].value:
            ws['E2'].value = 0
        if ws['F2'].value:
            ws['F2'].value = 0
        if ws['G2'].value:
            ws['G2'].value = 0

        ws['A' + saldo] = 'Saldo = '
        ws['A' + saldo].fill = PatternFill("solid", fgColor="00FFFF00")

        ws['B' + saldo].value = ws['B' + row_n].value - ws['B2'].value
        ws['B' + saldo].fill = PatternFill("solid", fgColor="00FFFF00")

        ws['C' + saldo].value = ws['C' + row_n].value - ws['C2'].value
        ws['C' + saldo].fill = PatternFill("solid", fgColor="00FFFF00")

        ws['D' + saldo].value = ws['D' + row_n].value - ws['D2'].value
        ws['D' + saldo].fill = PatternFill("solid", fgColor="00FFFF00")

        ws['E' + saldo].value = ws['E' + row_n].value - ws['E2'].value
        ws['E' + saldo].fill = PatternFill("solid", fgColor="00FFFF00")

        ws['F' + saldo].value = ws['F' + row_n].value - ws['F2'].value
        ws['F' + saldo].fill = PatternFill("solid", fgColor="00FFFF00")

        plotGraph(ws, localidade, nomeL)

    wb.save('registro' + nomeR + '.xlsx')

def plotGraph(ws, localidade, nomeL):
    max_row = len(localidade) + 1

    # Gráfico RDA e RCE
    chart1 = ScatterChart()
    # Gráfico ECON e PEND
    chart2 = ScatterChart()

    xvalues = Reference(ws, min_col=1, min_row=2, max_row=max_row)

    yvalues = Reference(ws, min_col=2, min_row=2, max_row=max_row)
    series = Series(yvalues, xvalues, title='RDA')
    chart1.series.append(series)

    yvalues = Reference(ws, min_col=3, min_row=2, max_row=max_row)
    series = Series(yvalues, xvalues, title='RCE - Sanesul')
    chart1.series.append(series)

    yvalues = Reference(ws, min_col=4, min_row=2, max_row=max_row)
    series = Series(yvalues, xvalues, title='RCE - Oneroso \nTerceiros')
    chart1.series.append(series)

    yvalues = Reference(ws, min_col=5, min_row=2, max_row=max_row)
    series = Series(yvalues, xvalues, title='RCE - Oneroso \nMs-Pantanal')
    chart1.series.append(series)

    yvalues = Reference(ws, min_col=6, min_row=2, max_row=max_row)
    series = Series(yvalues, xvalues, title='ECON')
    chart2.series.append(series)

    yvalues = Reference(ws, min_col=7, min_row=2, max_row=max_row)
    series = Series(yvalues, xvalues, title='PEND')
    chart2.series.append(series)

    chart1.title = nomeL + ' — RDA e RCE'
    chart1.width = 16.5
    ws.add_chart(chart1, 'I3')

    chart2.title = nomeL + ' — ECON e PENDENTES'
    ws.add_chart(chart2, 'I18')

if __name__ == '__main__':
    # Conectando com o banco de dados
    con = pyodbc.connect(
        'DRIVER={SQL Server};SERVER=10.100.100.48\\SCI;PORT=1433;DATABASE=SCI;Trusted_Connection=yes;')
    # Consultando a tabela do banco
    conHistorico = '''select REGIONAL, LOCAL as LOCALIDADE, REFERENCIA, COALESCE(RDA,0) AS RDA, COALESCE(RCE,0) AS [RCE - Sanesul], COALESCE(RCE_TERCEIRO,0) AS [RCE - Oneroso Terceiros],
    COALESCE(RCE_MSP,0) AS [RCE - Oneroso Ms-Pantanal], COALESCE(ECON,0) AS ECON, COALESCE(DDIFF,0) AS PENDENTES, DATAGERACAO as DATA 
    from dbo.historico_resultado ORDER BY REGIONAL, LOCAL, DATAGERACAO asc'''
    conResultado = '''WITH CTE_UltimaDataPorMes AS 	(SELECT REGIONAL,LOCAL AS LOCALIDADE, REFERENCIA, RDA, COALESCE(RCE,0) [RCE - Sanesul], COALESCE(RCE_TERCEIRO,0) [RCE - Oneroso Terceiros], 
	 COALESCE(RCE_MSP,0) [RCE - Oneroso Ms-Pantanal], COALESCE(ECON,0) [ECON], COALESCE(DDIFF,0) AS PENDENTES, DATAGERACAO AS DATA, 
	ROW_NUMBER() OVER (PARTITION BY LOCAL, YEAR(DATAGERACAO), MONTH(DATAGERACAO) ORDER BY DATAGERACAO DESC) AS RN FROM dbo.resultado) 
	SELECT REGIONAL, LOCALIDADE, REFERENCIA, RDA, [RCE - Sanesul],[RCE - Oneroso Terceiros], [RCE - Oneroso Ms-Pantanal], ECON, PENDENTES, DATA 
	FROM CTE_UltimaDataPorMes WHERE RN = 1 ORDER BY REGIONAL, LOCALIDADE, DATA;'''
    # Convertendo em DataFrame
    dfHistorico = pd.read_sql(conHistorico, con)
    dfResultado = pd.read_sql(conResultado, con)
    # Agrupando por Localidade e REFERENCIA
    

    
    dfgroup = dfResultado.groupby(["LOCALIDADE", "REFERENCIA"])
    # Selecionando a primeira linha de cada grupo
    dfResultado = dfgroup.tail(1)
    # Concatenando os DataFrames
    df = pd.concat([dfHistorico, dfResultado], ignore_index=True, sort=False)

    df['REGIONAL'] = df['REGIONAL'].str.strip().str.replace(r'\s+', ' ', regex=True)
    # Ordenando o DataFrame primeiro pela regional e depois por localidade
    df = df.sort_values(by=['REGIONAL', 'LOCALIDADE'])
    
    # Agrupando por Regional
    groupRegional = df.groupby("REGIONAL")

    for nomeR, regional in groupRegional:
        # Criação do Arquivo da Regional
        createWorkBook(nomeR, regional)
