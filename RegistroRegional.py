# pyinstaller RegistroRegional.py --onefile --noconsole

import pyodbc
from datetime import datetime
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.styles.numbers import NumberFormatList
from openpyxl.styles import Font, Fill
from openpyxl.styles import PatternFill
from openpyxl.comments import Comment


def plotGraph(ws, max_row, nomeL):

    # Gráfico RDA e RCE
    chart1 = ScatterChart()
    # Gráfico Econ e DIFF
    chart2 = ScatterChart()

    xvalues = Reference(ws, min_col=1, min_row=2, max_row=max_row)

    yvalues = Reference(ws, min_col=2, min_row=2, max_row=max_row)
    series = Series(yvalues, xvalues, title="RDA")
    chart1.series.append(series)

    yvalues = Reference(ws, min_col=3, min_row=2, max_row=max_row)
    series = Series(yvalues, xvalues, title="RCE - Sanesul")
    chart1.series.append(series)

    yvalues = Reference(ws, min_col=4, min_row=2, max_row=max_row)
    series = Series(yvalues, xvalues, title="RCE - Oneroso \n Terceiros")
    chart1.series.append(series)

    yvalues = Reference(ws, min_col=5, min_row=2, max_row=max_row)
    series = Series(yvalues, xvalues, title="RCE - Oneroso \n Ms-Pantanal")
    chart1.series.append(series)

    yvalues = Reference(ws, min_col=6, min_row=2, max_row=max_row)
    series = Series(yvalues, xvalues, title="ECON")
    chart2.series.append(series)

    yvalues = Reference(ws, min_col=7, min_row=2, max_row=max_row)
    series = Series(yvalues, xvalues, title="PEND")
    chart2.series.append(series)

    chart1.title = nomeL + "--RDA e RCE"
    chart1.width = 16.5
    ws.add_chart(chart1, "I3")
    chart2.title = nomeL + "--ECON e PENDENTES"
    ws.add_chart(chart2, "I18")


def nome(nomeR):
    if "BOLSAO/PARANAIBA" in nomeR:
        return "PNB"
    elif "GRANDE DOURADOS" in nomeR:
        return "DOS"
    elif "NORTE" in nomeR:
        return "CXM"
    elif "JIM" in nomeR:
        return "JIM"
    elif "CONE-SUL" in nomeR:
        return "NVR"
    elif "SUL/FRONTEIRA" in nomeR:
        return "PPR"
    elif "LESTE" in nomeR:
        return "NDI"
    elif "PANTANAL/CORUMBA" in nomeR:
        return "CMA"
    elif "PANTANAL/AQUIDAUANA" in nomeR:
        return "AUA"
    elif "BOLSAO/TRES LAGOAS" in nomeR:
        return "TLS"


if __name__ == "__main__":
    # Conectando com o banco de dados
    con = pyodbc.connect(
        "DRIVER={SQL Server};SERVER=10.100.100.48\\SCI;PORT=1433;DATABASE=SCI;Trusted_Connection=yes;"
    )
    # Consultando a tabela do banco
    conRegional = "select distinct(regional) from resultado;"
    cursor = con.cursor()
    cursor.execute(conRegional)
    regional = cursor.fetchall()
    wb = Workbook()
    conResultado = """select regional, max(datageracao) as datageracao, referencia as DATA, COALESCE(sum(RDA),0) as RDA, COALESCE(sum(RCE),0) as [RCE - Sanesul],
            COALESCE(sum(RCE_TERCEIRO),0) AS [RCE - Oneroso Terceiros],COALESCE(sum(RCE_MSP),0) AS [RCE - Oneroso MS-Pantanal],COALESCE(sum(ECON),0) as ECON,
            COALESCE(sum(DDIFF),0) as PEND	from historico_resultado where 
            regional = ? group by regional, referencia, RCE_TERCEIRO, RCE_MSP
			union
            select regional, max(datageracao) as datageracao, referencia as DATA, COALESCE(sum(RDA),0) as RDA, COALESCE(sum(RCE),0) as [RCE - Sanesul], 
			COALESCE(SUM(RCE_TERCEIRO),0) AS [RCE - Oneroso Terceiros], COALESCE(SUM(RCE_MSP),0) AS [RCE - Oneroso MS-Pantanal], COALESCE(sum(ECON),0) as ECON, 
			COALESCE(sum(DDIFF),0) as PEND from resultado where datageracao in (select max(DATAGERACAO) as DATAGERACAO from resultado group by referencia, local)
			and regional = ?  group by regional, referencia order by regional, datageracao;"""

    for r in regional:
        ws = wb.create_sheet(nome(r[0]), -1)
        col = [
            "DATA",
            "RDA",
            "Crescimento RDA",
            "OS Implantação de RDA",
            "RCE - Sanesul",
            "RCE - Oneroso Terceiros",
            "RCE - Oneroso Ms-Pantanal",
            "Crescimento RCE",
            "ECON",
            "PENDENTES - DDIF",
        ]
        com = [
            "DATA - Mês e Ano",
            "RDA - Rede de Distribuição de Água - Extensão de Rede de água com status da rede = ""existente""",
            "Crescimento RDA - Variação do Total de RDA em relação ao período anterior",
            "OS Implantação de RDA - Ordens de Serviço para Implantação de RDA",
            "RCE - Sanesul - Rede coletora de Esgoto Sanesul - Extensão de Rede de Esgoto com status da rede = ""existente""",
            "RCE - Oneroso Terceiros - Rede coletora de Esgoto Terceiros - Extensão de Rede de Esgoto com status da rede = ""Oneroso Terceiros""",
            "RCE - Oneroso MS-Pantanal - Rede coletora de Esgoto MS-Pantanal - Extensão de Rede de Esgoto com status da rede = ""Oneroso MS-Pantanal""",
            "Crescimento RCE - Variação do Total de RCE em relação ao período anterior",
            "ECON - Economia - Total de Matriculas Georreferenciadas",
            "PENDENTES - DDIF - Diferença entre o Total de Matriculas Cadastrais da Comercial e o Total de Matriculas Georreferenciados ou seja matriculas sem georreferenciamento, ou seja, pendentes de georreferenciamento.",
        ]
        ws.append(col)
        # Inclui comentários descritivos nas células de cabeçalho
        for idx, texto_comentario in enumerate(com, start=1):
            comment = Comment(texto_comentario, "INFO", width=300, height=100)
            cell = ws.cell(row=1, column=idx)
            cell.comment = comment
            cell.alignment = Alignment(horizontal="center", vertical="center")
            # Ajusta a largura da coluna com base no comprimento do texto do cabeçalho
            cell_value = ws.cell(row=1, column=idx).value
            if cell_value:
                col_letter = get_column_letter(idx)
                ws.column_dimensions[col_letter].width = len(str(cell_value)) + 5
        nLinhas = 0
        cursor.execute(conResultado, r[0], r[0])
        row = cursor.fetchone()
        anterior = 0
        wsAnterior = True
        while row != None:
            row[2] = datetime.strptime(row[2], "%m/%Y")
            if row[2] != anterior:
                nLinhas += 1
                ws.append(row[2:])
            anterior = row[2]
            row = cursor.fetchone()
            if ws["F" + str(nLinhas + 1)].value and wsAnterior:
                f = str(nLinhas + 1)
                wsAnterior = False

        for cell in ws["A"]:
            cell.number_format = "mmm-yy"
        plotGraph(ws, nLinhas + 1, r[0])
        saldo = str(nLinhas + 2)
        ws["A" + saldo] = "Saldo = "
        ws["A" + saldo].fill = PatternFill("solid", fgColor="00FFFF00")

        ws["B" + saldo].value = ws["B" + str(nLinhas + 1)].value - ws["B2"].value
        ws["B" + saldo].fill = PatternFill("solid", fgColor="00FFFF00")

        ws["C" + saldo].value = ws["C" + str(nLinhas + 1)].value - ws["C2"].value
        ws["C" + saldo].fill = PatternFill("solid", fgColor="00FFFF00")

        ws["D" + saldo].value = ws["D" + str(nLinhas + 1)].value - ws["D2"].value
        ws["D" + saldo].fill = PatternFill("solid", fgColor="00FFFF00")

        ws["E" + saldo].value = ws["E" + str(nLinhas + 1)].value - ws["E2"].value
        ws["E" + saldo].fill = PatternFill("solid", fgColor="00FFFF00")

        ws["F" + saldo].value = ws["F" + str(nLinhas + 1)].value - ws["F" + f].value
        ws["F" + saldo].fill = PatternFill("solid", fgColor="00FFFF00")

    wb.save("registroRegional.xlsx")
