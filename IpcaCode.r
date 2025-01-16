

library(forecast)
library(forecast)
library(openxlsx)
library(tseries)

dados_totais$Date <- as.Date(paste0("01-", dados_totais$`Unnamed: 0`), format = "%d-%B %Y")

# Remover a coluna de datas do dataframe para análise
data_ts <- dados_totais[,-1]

dados_dez_23 <- dados_totais[1:144,-1]

dados_jan_24 <- dados_totais[1:145,-1]

dados_fev_24 <- dados_totais[1:146,-1]

dados_mar_24 <- dados_totais[1:147,-1]

dados_abr_24 <- dados_totais[1:148,-1]

dados_mai_24 <- dados_totais[1:149,-1]

dados_jun_24 <- dados_totais[1:150,-1]

dados_jul_24 <- dados_totais[1:151,-1]

# Inicializar listas para armazenar os modelos ajustados e previsões
sarima_models <- list()
forecasts <- numeric(length(colnames(data_ts)))

# Loop através de cada coluna do dataframe
for (i in seq_along(colnames(data_ts))) {
  col_name <- colnames(data_ts)[i]
  # Obter a série temporal
  ts_data <- ts(dados_jul_24[[col_name]], frequency = 12, start = c(2012, 1))
  
  # Ajustar um modelo SARIMA automaticamente usando auto.arima
  sarima_model <- auto.arima(ts_data, seasonal = TRUE)
  
  # Armazenar o modelo na lista
  sarima_models[[col_name]] <- sarima_model
  
  # Gerar previsão de 1 passo à frente
  forecast_result <- forecast(sarima_model, h = 1)
  
  # Armazenar a previsão na lista
  forecasts[i] <- as.numeric(forecast_result$mean)
  
  # Imprimir resumo do modelo ajustado e previsão
  cat("Modelo ajustado para", col_name, ":\n")
  print(summary(sarima_model))
  cat("\nPrevisão de 1 passo à frente:\n")
  print(forecast_result$mean)
  cat("\n\n")
}

# Criar um dataframe para as previsões
forecast_df <- data.frame(Produto = colnames(data_ts), Previsao = forecasts)

# Criar um novo workbook
wb <- createWorkbook()

# Adicionar uma nova aba ao workbook
addWorksheet(wb, sheetName = "Previsoes")

# Escrever o dataframe de previsões na aba
writeData(wb, sheet = "Previsoes", forecast_df)

# Salvar o workbook em um arquivo Excel
saveWorkbook(wb, "backtest_ago_24.xlsx", overwrite = TRUE)

cat("Previsões exportadas para o arquivo 'previsoes_sarima.xlsx'.\n")


