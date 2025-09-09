# -*- coding: utf-8 -*-
import pandas as pd
import psycopg2
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from io import BytesIO
from dotenv import load_dotenv
import os
import logging

load_dotenv()

DB_CONFIG = {
    "host": os.getenv("DB_HOST"),
    "database": os.getenv("DB_NAME"),
    "user": os.getenv("DB_USER"),
    "password": os.getenv("DB_PASS"),
}

TOKEN = os.getenv("TELEGRAM_TOKEN")

# Conexão com banco
def get_inativos_voors():
    conn = psycopg2.connect(**DB_CONFIG)
    # Mudança 1: Aqui, renomeamos a coluna 'cpf' da consulta para 'documento'.
    # Isso faz com que o DataFrame retornado já tenha a coluna com o nome correto para a comparação.
    query = "SELECT cpf AS documento, nome FROM usuarios_voors WHERE situacao = 'inactive'"
    df = pd.read_sql(query, conn)
    conn.close()
    
    # Mudança para converter o tipo da coluna 'documento' para inteiro.
    # Usamos o método astype para garantir que a coluna seja do tipo int64, apropriado para CPF.
    df['documento'] = df['documento'].astype('int64')

    return df

# Start
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Olá! Envie o arquivo Excel da Leveduca para eu comparar com o VOORS. Atenção: O Relatorio deve contar obrigatoriamente a coluna documento exatamente desse jeito.")
    logging.basicConfig(level=logging.INFO)
    logging.info("Foi Startado o BOT")

# Upload de arquivo
async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    file = await update.message.document.get_file()
    file_bytes = await file.download_as_bytearray()

    # Leitura Excel
    df_leveduca = pd.read_excel(BytesIO(file_bytes))

    df_leveduca.columns = df_leveduca.columns.str.strip().str.lower()

    # Consulta banco
    df_inativos = get_inativos_voors()

    print("Colunas do df_leveduca ANTES da limpeza:", df_leveduca.columns)

    # Mudança 2: Converte a coluna 'documento' do Excel para o tipo numérico, garantindo que
    # a comparação com os dados do banco de dados seja precisa.
    df_leveduca['documento'] = pd.to_numeric(df_leveduca['documento'], errors='coerce')

    print("Colunas do df_leveduca DEPOIS da limpeza:", df_leveduca.columns)
    print("Primeiras 5 linhas do df_leveduca:\n", df_leveduca.head())

    # Comparação
    # Agora, usamos a coluna 'documento' do DataFrame df_inativos para a comparação.
    # O nome foi alterado na consulta SQL, então o DataFrame já vem pronto.
    df_resultado = df_leveduca[df_leveduca['documento'].isin(df_inativos['documento'])]

    if not df_resultado.empty:
        # Gera novo Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_resultado.to_excel(writer, index=False, sheet_name="Leveduca")
        output.seek(0)

        await update.message.reply_document(document=output, filename="inativos_corrigir.xlsx")
    else:
        await update.message.reply_text("✅ Nenhum inativo do VOORS encontrado na listagem da Leveduca.")

# Main
def main():
    print("Iniciando bot...")
    app = Application.builder().token(TOKEN).build()

    # Adiciona handlers
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_file))

    print("Bot conectado, aguardando mensagens...")
    app.run_polling()

if __name__ == "__main__":
    main()
