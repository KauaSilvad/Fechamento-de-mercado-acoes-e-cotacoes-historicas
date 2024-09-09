import yfinance as yf
import pandas as pd 
import matplotlib.pyplot as plt
import mplcyberpunk

#contando o historico das cotações historicas do yahoo finance
tickers = ["^BVSP", "^GSPC", "BRL=X", "ABEV3.SA"]
dados_mercado = yf.download(tickers, period = "6mo")
dados_mercado = dados_mercado["Adj Close"]

#Tratamento de dados
dados_mercado = dados_mercado.dropna()
dados_mercado.columns = ["DOLAR","IBOVESPA","S&P500","AMBEV"]
print(dados_mercado)

#criação do gráfico com a biblioteca matplotlib
plt.style.use("cyberpunk") 
plt.plot(dados_mercado["IBOVESPA"]) 
plt.title("IBOVESPA")
plt.savefig("ibovespa.png")

plt.style.use("cyberpunk") 
plt.plot(dados_mercado["DOLAR"]) 
plt.title("DOLAR")
plt.savefig("dolar.png")

plt.style.use("cyberpunk") 
plt.plot(dados_mercado["S&P500"]) 
plt.title("S&P500")
plt.savefig("s&p500.png")

plt.style.use("cyberpunk") 
plt.plot(dados_mercado["AMBEV"]) 
plt.title("AMBEV")
plt.savefig("ambev.png")

#os retornos diarios das cotações por enquanto só dolar,ibovespa,s&p500 e ambev
retornos_diarios = dados_mercado.pct_change()
print(retornos_diarios)

retorno_dolar = retornos_diarios["DOLAR"].iloc[-1]
retorno_ibovespa = retornos_diarios["IBOVESPA"].iloc[-1]
retorno_fep500 = retornos_diarios["S&P500"].iloc[-1]
retorno_ambev = retornos_diarios["AMBEV"].iloc[-1]
#multiplicando por 100 para deixar na porcentagem
retorno_dolar = str(round(retorno_dolar * 100, 2)) + "%"
print(retorno_dolar)
retorno_ibovespa = str(round(retorno_ibovespa * 100, 2)) + "%"
print(retorno_ibovespa)
retorno_fep500 = str(round(retorno_fep500 * 100, 2)) + "%"
print(retorno_fep500)
retorno_ambev = str(round(retorno_ambev * 100, 2)) + "%"
print(retorno_ambev)

#enviando por email outlook
import win32com.client as win32
import os

try:
    outlook = win32.Dispatch("outlook.application")
    email = outlook.CreateItem(0)
    email.To = "kauahd7@gmail.com"#email do destinatario 
    email.Subject = "Relatório de Mercado"
    email.Body = f'''
    Olá amigo, segue o relatório de mercado:

    * O Ibovespa teve o retorno de {retorno_ibovespa}
    * O Dólar teve o retorno de {retorno_dolar}
    * O S&P500 teve o retorno de {retorno_fep500}
    * A Ambev teve o retorno de {retorno_ambev}

    A seguir em anexo a performance dos ativos nos últimos 6 meses.
    '''

    anexo_ibovespa = r"C:\Users\kauah\OneDrive\Projeto fechamento-de-mercado\ibovespa.png"
    anexo_dolar = r"C:\Users\kauah\OneDrive\Projeto fechamento-de-mercado\dolar.png"
    anexo_sp500 = r"C:\Users\kauah\OneDrive\Projeto fechamento-de-mercado\s&p500.png"
    anexo_ambev = r"C:\Users\kauah\OneDrive\Projeto fechamento-de-mercado\ambev.png"

    # verificação se os emails estão todos existindo
    for anexo in [anexo_ibovespa, anexo_dolar, anexo_sp500, anexo_ambev]:
        if not os.path.isfile(anexo):
            print(f"Arquivo não encontrado: {anexo}")
            raise FileNotFoundError(f"Arquivo não encontrado: {anexo}")

    # Adicionando as fotos para o anexo
    email.Attachments.Add(anexo_ibovespa)
    email.Attachments.Add(anexo_dolar)
    email.Attachments.Add(anexo_sp500)
    email.Attachments.Add(anexo_ambev)

    email.Send()
    print("Email enviado com sucesso.")

except Exception as e:
    print(f"Erro ao enviar o e-mail: {e}")
