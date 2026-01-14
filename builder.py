import pandas as pd
import xlsxwriter
import calendar
from datetime import date

def gerar_planilha_final(ano, mes):
    nome_mes = calendar.month_name[mes]
    nome_arquivo = f'Planilha_Barbearia_Victor_{nome_mes}_{ano}.xlsx'
    
    print(f"Gerando arquivo com painéis congelados: {nome_arquivo}...")

    writer = pd.ExcelWriter(nome_arquivo, engine='xlsxwriter')
    workbook = writer.book

    # --- ESTILOS ---
    fmt_header = workbook.add_format({
        'bold': True, 'align': 'center', 'valign': 'vcenter',
        'fg_color': '#4472C4', 'font_color': 'white', 'border': 1
    })
    fmt_branca = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
    fmt_azul = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#D9E1F2'})
    fmt_moeda_branca = workbook.add_format({'num_format': 'R$ #,##0.00', 'border': 1, 'align': 'center'})
    fmt_moeda_azul = workbook.add_format({'num_format': 'R$ #,##0.00', 'border': 1, 'align': 'center', 'fg_color': '#D9E1F2'})
    fmt_total = workbook.add_format({'bold': True, 'border': 1, 'align': 'center', 'fg_color': '#BFBFBF', 'num_format': 'R$ #,##0.00'})
    fmt_total_final = workbook.add_format({'bold': True, 'border': 1, 'align': 'center', 'fg_color': '#FFFF00', 'num_format': 'R$ #,##0.00', 'font_size': 12})

    # Listas e Cabeçalhos
    opcoes_pagamento = ['PIX', 'DINHEIRO', 'CARTÃO DE DÉBITO', 'CARTÃO DE CRÉDITO']
    cab_cortes = ['ITEM', 'BARBEIRO', 'CLIENTE', 'PAGAMENTO', 'VALOR']
    cab_picole = ['ITEM (PICOLÉ)', 'PAGAMENTO', 'VALOR', 'CUSTO', 'LUCRO']
    cab_bebidas = ['ITEM (BEBIDAS)', 'PAGAMENTO', 'VALOR', 'CUSTO', 'LUCRO']

    num_dias = calendar.monthrange(ano, mes)[1]
    linhas_tabela = 35 

    # --- LOOP DOS DIAS ---
    for dia in range(1, num_dias + 1):
        nome_aba = f'{dia:02d}'
        ws = workbook.add_worksheet(nome_aba)
        
        # --- NOVIDADE: CONGELAR A PRIMEIRA LINHA ---
        ws.freeze_panes(1, 0) 

        # Larguras
        ws.set_column('A:A', 5)   
        ws.set_column('B:B', 12)  
        ws.set_column('C:C', 20)  
        ws.set_column('D:D', 18)  
        ws.set_column('E:E', 12)  
        ws.set_column('F:G', 1)   
        ws.set_column('H:H', 15)  
        ws.set_column('I:I', 18)  
        ws.set_column('J:L', 10)  
        ws.set_column('M:M', 1)   
        ws.set_column('N:N', 15)  
        ws.set_column('O:O', 18)  
        ws.set_column('P:R', 10)  

        # Cabeçalhos
        ws.write_row('A1', cab_cortes, fmt_header)
        ws.write_row('H1', cab_picole, fmt_header)
        ws.write_row('N1', cab_bebidas, fmt_header)

        # Preenchimento Zebrado
        for i in range(linhas_tabela):
            row = i + 1 
            excel_row = row + 1 

            if row % 2 == 0:
                style_txt = fmt_branca
                style_money = fmt_moeda_branca
            else:
                style_txt = fmt_azul
                style_money = fmt_moeda_azul

            # Cortes
            ws.write(row, 0, i + 1, style_txt)
            ws.write(row, 1, 'VICTOR', style_txt)
            ws.write_blank(row, 2, None, style_txt)
            ws.write_blank(row, 3, None, style_txt)
            ws.write_blank(row, 4, None, style_money)
            ws.data_validation(row, 3, row, 3, {'validate': 'list', 'source': opcoes_pagamento})

            # Picolé
            ws.write_blank(row, 7, None, style_txt)
            ws.write_blank(row, 8, None, style_txt)
            ws.write_blank(row, 9, None, style_money)
            ws.write_blank(row, 10, None, style_money)
            ws.write_formula(row, 11, f'=IF(J{excel_row}<>"", J{excel_row}-K{excel_row}, "-")', style_money)
            ws.data_validation(row, 8, row, 8, {'validate': 'list', 'source': opcoes_pagamento})

            # Bebidas
            ws.write_blank(row, 13, None, style_txt)
            ws.write_blank(row, 14, None, style_txt)
            ws.write_blank(row, 15, None, style_money)
            ws.write_blank(row, 16, None, style_money)
            ws.write_formula(row, 17, f'=IF(P{excel_row}<>"", P{excel_row}-Q{excel_row}, "-")', style_money)
            ws.data_validation(row, 14, row, 14, {'validate': 'list', 'source': opcoes_pagamento})

        # Rodapés (Totais do Dia)
        row_total = linhas_tabela + 1
        excel_row_total = row_total + 1
        
        ws.write(row_total, 2, "TOTAL DO DIA:", fmt_total)
        ws.write_formula(row_total, 3, f'=COUNTA(C2:C{row_total}) & " Clientes"', fmt_total)
        ws.write_formula(row_total, 4, f'=SUM(E2:E{row_total})', fmt_total)

        ws.write(row_total, 8, "LUCRO PICOLÉ:", fmt_total)
        ws.write_formula(row_total, 11, f'=SUM(L2:L{row_total})', fmt_total)

        ws.write(row_total, 14, "LUCRO BEBIDAS:", fmt_total)
        ws.write_formula(row_total, 17, f'=SUM(R2:R{row_total})', fmt_total)

        # Lucro Líquido Grande
        row_resumo = row_total + 2
        ws.merge_range(row_resumo, 2, row_resumo, 3, "LUCRO LIQUIDO TOTAL (DIA):", fmt_header)
        ws.write_formula(row_resumo, 4, f'=SUM(E{excel_row_total},L{excel_row_total},R{excel_row_total})', fmt_total_final)

    # --- ABA TOTAL DO MÊS ---
    ws_resumo = workbook.add_worksheet('TOTAL DO MÊS')
    
    # --- NOVIDADE: CONGELAR A PRIMEIRA LINHA DO RESUMO ---
    ws_resumo.freeze_panes(1, 0)

    ws_resumo.write_row('A1', ['Dia', 'Cortes (Victor)', 'Lucro Picolé', 'Lucro Bebidas', 'LUCRO TOTAL'], fmt_header)
    ws_resumo.set_column('A:A', 10, fmt_branca)
    ws_resumo.set_column('B:E', 20, fmt_moeda_branca)

    for i in range(1, num_dias + 1):
        row = i
        nome_aba = f'{i:02d}'
        linha_total_excel = linhas_tabela + 2 
        
        ws_resumo.write_datetime(row, 0, date(ano, mes, i), workbook.add_format({'num_format': 'dd/mm', 'border':1, 'align': 'center'}))
        ws_resumo.write_formula(row, 1, f"='{nome_aba}'!E{linha_total_excel}", fmt_moeda_branca)
        ws_resumo.write_formula(row, 2, f"='{nome_aba}'!L{linha_total_excel}", fmt_moeda_branca)
        ws_resumo.write_formula(row, 3, f"='{nome_aba}'!R{linha_total_excel}", fmt_moeda_branca)
        ws_resumo.write_formula(row, 4, f"=SUM(B{row+1}:D{row+1})", fmt_moeda_branca)

    row_final = num_dias + 1
    ws_resumo.write(row_final, 0, "TOTAL", fmt_total)
    ws_resumo.write_formula(row_final, 1, f"=SUM(B2:B{row_final})", fmt_total)
    ws_resumo.write_formula(row_final, 2, f"=SUM(C2:C{row_final})", fmt_total)
    ws_resumo.write_formula(row_final, 3, f"=SUM(D2:D{row_final})", fmt_total)
    ws_resumo.write_formula(row_final, 4, f"=SUM(E2:E{row_final})", fmt_total_final)

    writer.close()
    print("Sucesso! Planilha gerada com linha congelada em todas as abas.")

# --- EXECUÇÃO: TROQUE AQUI O ANO E O MÊS ---
gerar_planilha_final(2026, 12) # Julho 2024