from datetime import datetime
from decimal import Decimal

import pandas as pd

import requests
from rich.console import Console
from rich.table import Table

import datetime as dt
from dateutil.relativedelta import *

import win32com.client as win32



def media(valor, dias):
    _media = valor / dias

    return str(round(Decimal(_media), ndigits=2))


def get_prices():
    url_base = "https://lme.gorilaxpress.com/cotacao"
    data_link = "2cf4ff0e-8a30-48a5-8add-f4a1a63fee10/json"

    response = requests.get(f"{url_base}/{data_link}/")

    prices = response.json()["results"]

    title_icon = "[not italic]:chart_with_upwards_trend:[/]"
    title_text = "Cotação London Metal Exchange"

    table = Table(title=f"{title_icon} {title_text} {title_icon}")

    table.add_column("Data e semana", justify="center", style="cyan")
    table.add_column("Zinco", justify="right", style="magenta")
    table.add_column("Cobre", justify="right", style="red")
    table.add_column("Alumínio", justify="right", style="blue")
    table.add_column("Chumbo", justify="right", style="bright_cyan")
    table.add_column("Estanho", justify="right", style="bright_blue")
    table.add_column("Níquel", justify="right", style="bright_yellow")
    table.add_column("Dolar", justify="right", style="green")

    ultima_semana = datetime.strptime(prices[0]["data"], "%Y-%m-%d").strftime(
        "%U")

    semana_atual = datetime.strptime(prices[-1]["data"], "%Y-%m-%d").strftime(
        "%U")

    dias = 0
    media_zinco = 0
    media_cobre = 0
    media_aluminio = 0
    media_chumbo = 0
    media_estanho = 0
    media_niquel = 0
    media_dolar = 0

    diaas=list()
    zinco=list()
    cobre=list()
    aluminio=list()
    chumbo=list()
    estanho=list()
    niquel=list()
    dolar=list()

    for price in prices:
        dia = datetime.strptime(price["data"], "%Y-%m-%d").strftime("%d/%m/%Y")
        semana_numero = datetime.strptime(dia, "%d/%m/%Y").strftime("%U")
        diaas.append(dia)

        dias += 1
        media_zinco += Decimal(price["zinco"])
        media_cobre += Decimal(price["cobre"])
        media_aluminio += Decimal(price["aluminio"])
        media_chumbo += Decimal(price["chumbo"])
        media_estanho += Decimal(price["estanho"])
        media_niquel += Decimal(price["niquel"])
        media_dolar += Decimal(price["dolar"])

        if dia == 1 or dia == 0 or dia == 2:
            diaas.append(dia)
        print(diaas)
        
        print(price['zinco'])
        

        if semana_numero == ultima_semana:
            table.add_row(f'{dia} - {semana_numero}', price["zinco"],
                          price["cobre"],
                          price["aluminio"], price["chumbo"],
                          price["estanho"], price["niquel"], price["dolar"])

            if dias == 4 and semana_numero == semana_atual:
                table.add_row(f"[dim]Média Semana {semana_numero}",
                              f"[dim]{media(media_zinco, dias)}",
                              f"[dim]{media(media_cobre, dias)}",
                              f"[dim]{media(media_aluminio, dias)}",
                              f"[dim]{media(media_chumbo, dias)}",
                              f"[dim]{media(media_estanho, dias)}",
                              f"[dim]{media(media_niquel, dias)}",
                              f"[dim]{media(media_dolar, dias)}")

        else:
            table.add_row(f"[dim]Média Semana {int(semana_numero) - 1}",
                          f"[dim]{media(media_zinco, dias)}",
                          f"[dim]{media(media_cobre, dias)}",
                          f"[dim]{media(media_aluminio, dias)}",
                          f"[dim]{media(media_chumbo, dias)}",
                          f"[dim]{media(media_estanho, dias)}",
                          f"[dim]{media(media_niquel, dias)}",
                          f"[dim]{media(media_dolar, dias)}")
            table.add_row(f'{dia} - {semana_numero}', price["zinco"],
                          price["cobre"],
                          price["aluminio"], price["chumbo"],
                          price["estanho"], price["niquel"], price["dolar"])

            dias = 0
            media_zinco = 0
            media_cobre = 0
            media_aluminio = 0
            media_chumbo = 0
            media_estanho = 0
            media_niquel = 0
            media_dolar = 0
            ultima_semana = semana_numero

        today = dt.date.today()
        
        firstday_acctual_month = today+relativedelta(day=1)
        firstday_next_month = firstday_acctual_month+relativedelta(months=+1)
        lastday_acctual_month = firstday_next_month+relativedelta(days=-1)
        
        lastday_last_month = firstday_acctual_month+relativedelta(days=-1)
        firstday_last_month = lastday_last_month+relativedelta(day=1)

    

        if today == lastday_acctual_month:
            print("today is the day")
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = 'Email to'
            mail.Subject = 'Email title'
            mail.Body = 'Email text body'


            # To attach a file to the email (optional):
            attachment  = r"C:\Users\ct67ca\Desktop\emailSend\docTeste.docx"
            mail.Attachments.Add(attachment)

mail.Send()

        else:
            df = pd.DataFrame({
                'Day': diaas,
                'zinco': price['zinco'],
                'cobre':price["cobre"],
                'aluminio':price["aluminio"],
                'chumbo':price["chumbo"],
                'estanho':price["estanho"],
                'niquel':price["niquel"],
                'dolar':price["dolar"]})

            writer = pd.ExcelWriter('demo.xlsx', engine='xlsxwriter')

            df.to_excel(writer, sheet_name='main', index=False)

            writer.save()

    

    console = Console()
    console.print(table, justify="center")


if __name__ == "__main__":
    get_prices()
