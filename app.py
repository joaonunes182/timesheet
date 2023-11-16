import os
import re
import openpyxl
import calendar
from datetime import datetime, timedelta
import random
import tkinter as tk
from tkinter import messagebox, simpledialog
import logging

# Verificar se o diretório "LOGS" existe, caso contrário, criar
log_dir = "logs"
if not os.path.exists(log_dir):
    os.makedirs(log_dir)

# Configurar o registro
log_filename = os.path.join(log_dir, 'logteste.log')
logging.basicConfig(filename=log_filename, level=logging.INFO,
                    format='%(asctime)s - %(levelname)s: %(message)s')

def exibir_mensagem():
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("MediTimeSheet", "Bem-vindo, este é um script para preencher a sua folha de ponto. \n\n\nDesenvolvido pelo João")
    logging.info("script aberto em: %s", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    messagebox.showinfo("Lembrete", "Lembre-se de adicionar a sua assinatura e o seu nome na planilha.")
    messagebox.showinfo("Lembrete", "Lembre-se de revisar a planilha antes de enviar, verifique os feriados.")

exibir_mensagem()

# Função para exibir alertas com log
def exibir_alert(mensagem):
    logging.info("Alerta: %s", mensagem)
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("Lembrete", mensagem)

# Função para obter horários com log
def obter_horario_usuario(mensagem):
    while True:
        entrada = simpledialog.askstring("Horário", mensagem)
        if entrada is None:
            logging.info("Usuário cancelou a entrada.")
            return None  # Usuário cancelou a entrada
        if re.match(r'^\d{2}:\d{2}$', entrada):
            logging.info("Horário fornecido: %s", entrada)
            return entrada
        else:
            erro = "Formato de horário inválido. Certifique-se de usar o formato HH:MM."
            logging.error("Erro: %s", erro)
            exibir_alert(erro)

meses_em_portugues = [
    'janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho',
    'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'
]
meses_em_ingles = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
]

mes_ingles = None

while mes_ingles is None:
    mes = simpledialog.askstring("Mês", "Digite o mês desejado (ex: janeiro, fevereiro, ...): ")
    mes = mes.lower()
    if mes in meses_em_portugues or mes in [mes.lower() for mes in meses_em_ingles]:
        mes_ingles = meses_em_ingles[meses_em_portugues.index(mes)] if mes in meses_em_portugues else mes.capitalize()
    else:
        exibir_alert("Mês inválido. Certifique-se de usar um nome de mês válido (ex: janeiro, fevereiro, ...).")

if mes_ingles:
    entrada = obter_horario_usuario("Digite o horário de início do expediente no formato (HH:MM): ")
    if entrada is not None:
        intervalo_ida = obter_horario_usuario("Digite o horário de saída para o almoço no formato (HH:MM): ")
        if intervalo_ida is not None:
            excel_filename = "arquivo/folha_de_ponto.xlsx"
            workbook = openpyxl.load_workbook(excel_filename)
            sheet = workbook.active

            celulas_entrada = sheet['C12':'C42']
            celulas_intervalo1 = sheet['F12':'F42']
            celulas_intervalo2 = sheet['I12':'I42']
            celulas_saida = sheet['L12':'L42']

            numero_de_dias = calendar.monthrange(2023, meses_em_ingles.index(mes_ingles) + 1)[1]

            horas_entrada, minutos_entrada = map(int, entrada.split(':'))
            horas_intervalo_ida, minutos_intervalo_ida = map(int, intervalo_ida.split(':'))

            data = datetime(2023, meses_em_ingles.index(mes_ingles) + 1, 1)
            for i in range(len(celulas_entrada)):
                minutos_entrada_aleatorios = random.randint(0, 10)
                minutos_intervalo_aleatorios = random.randint(0, 10)

                if data.weekday() == 5 or data.weekday() == 6:
                    data += timedelta(days=1)
                    continue

                horario_entrada = data.replace(hour=horas_entrada, minute=minutos_entrada + minutos_entrada_aleatorios)
                horario_intervalo_ida = data.replace(hour=horas_intervalo_ida, minute=minutos_intervalo_ida + minutos_intervalo_aleatorios)
                horario_saida = horario_entrada + timedelta(hours=9)

                celulas_entrada[i][0].value = horario_entrada.strftime("%H:%M")
                celulas_intervalo1[i][0].value = horario_intervalo_ida.strftime("%H:%M")
                celulas_intervalo2[i][0].value = (horario_intervalo_ida + timedelta(hours=1)).strftime("%H:%M")
                celulas_saida[i][0].value = horario_saida.strftime("%H:%M")

                data += timedelta(days=1)

            data_inicio = f"01/{meses_em_portugues.index(mes) + 1}/2023"
            data_fim = f"{numero_de_dias:02d}/{meses_em_portugues.index(mes) + 1}/2023"
            sheet['D6'] = data_inicio
            sheet['F6'] = data_fim

            excel_filename = f'Folha_Ponto_{mes}.xlsx'
            workbook.save(excel_filename)
            exibir_alert(f'Planilha "{excel_filename}" preenchida com sucesso')

logging.info("script fechado em: %s", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))