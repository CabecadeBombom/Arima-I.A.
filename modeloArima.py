import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from statsmodels.tsa.arima.model import ARIMA
from statsmodels.tsa.stattools import adfuller
import itertools
from statsmodels.tsa.statespace.sarimax import SARIMAX

# Carregar a planilha (substitua o caminho pelo local do arquivo)
df = pd.read_excel(r'C:\Users\luizv\OneDrive\Área de Trabalho\Trabalho\ModeloArima\XLS\CLSTR02.xlsx')

# Verificar se os dados foram carregados corretamente
print(df.head())

# Converter a coluna 'Data' para datetime
df['Data'] = pd.to_datetime(df['Data'], format='%Y-%m-%d')  

# Definir a coluna 'Data' como índice
df.set_index('Data', inplace=True)

# Verificar se há valores ausentes e lidar com eles (opcional)
df = df.fillna(method='ffill')  # Preencher valores ausentes com o valor anterior

# Função para ajustar ARIMA ou SARIMAX e prever o futuro para cada disco
def arima_forecast(disk_data, steps=100):  # (100 semanas quase 2 anos)
    #model = ARIMA(disk_data, order=(7, 1, 1))  # ARIMA(5,1,0) esses dados podem ser alterados para aumentar a precisão
    model = SARIMAX(disk_data, order=(7,1,1), seasonal_order=(1,1,0,52))
    model_fit = model.fit()
    forecast = model_fit.forecast(steps=steps)  # Prever os próximos 'steps' períodos
    return forecast

# Lista para armazenar previsões
forecast_results = {}
efficiency_limit = 0.98  # Limite de eficiência, por exemplo, 90%

# Aplicar ARIMA em cada disco
for column in df.columns:
    disk_data = df[column]
    
    # Converter para numérico e remover valores ausentes
    disk_data = pd.to_numeric(disk_data, errors='coerce').dropna()

    # Fazer previsão para os próximos 100 períodos (quase 2 anos)
    forecast = arima_forecast(disk_data, steps=100)  # Aumentamos o período de previsão

    # Armazenar a previsão
    forecast_results[column] = forecast
    
    # Gerar uma série de datas futuras (considerando previsão semanal)
    future_dates = pd.date_range(df.index[-1], periods=len(forecast) + 1, freq='W')[1:]
    
    # Plotar os dados históricos e previsões
    plt.figure(figsize=(10, 6))
    plt.plot(df.index, disk_data, label='Histórico')
    plt.plot(future_dates, forecast, label='Previsão', color='red')
    plt.axhline(y=efficiency_limit, color='r', linestyle='--', label=f'Limite de eficiência ({efficiency_limit * 100}%)')
    plt.title(f'Previsão para {column}')
    plt.legend()
    plt.show()

    # Verificar quando o disco ultrapassará o limite de eficiência
    for date, value in zip(future_dates, forecast):
        print(f"Data: {date}, Previsão de Utilização: {value:.2f}")
        if value > efficiency_limit:
            print(f"Atenção! O disco {column} deve ultrapassar {efficiency_limit * 100}% de uso em {date}.")
            break

# Exemplo para encontrar a melhor ordem do modelo (opcional)
p = d = q = range(0, 3)
pdq = list(itertools.product(p, d, q))

# Testar diferentes combinações (opcional)
for param in pdq:
    try:
        model = ARIMA(disk_data, order=param)
        model_fit = model.fit()
        print(f'ARIMA{param} - AIC:{model_fit.aic}')
    except:
        continue

# Salvar previsões em um arquivo Excel
forecast_df = pd.DataFrame(forecast_results)
forecast_df.index = future_dates  # datas futuras como índice no DataFrame
forecast_df.to_excel(r'C:\Users\luizv\OneDrive\Área de Trabalho\Trabalho\ModeloArima\XLS\previsoes_disks.xlsx')
