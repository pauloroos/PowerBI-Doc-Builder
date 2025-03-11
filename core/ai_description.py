
import os
import pandas as pd
import google.generativeai as genai
from core.helpers import *


# Função gerar descrição com IA
def gerar_descricao_ia(pasta_arquivos):
    """
    Gera uma descrição do dashboard usando IA com base nos arquivos CSV gerados.
    
    Args:
        pasta_arquivos (str): Caminho da pasta onde os arquivos CSV estão armazenados.
    
    Returns:
        str: Texto gerado pela IA descrevendo o dashboard.
    """
    config = carregar_config()
    api_key = config.get("api_key", "")
    
    if not api_key:
        return "❌ API Key do Gemini não encontrada no config.json"
    
     # Configurar a API
    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel("gemini-1.5-flash")
    except Exception as e:
        return f"❌ Erro ao configurar a API: {str(e)}"
    
    try:
        df_columns = pd.read_csv(os.path.join(pasta_arquivos, "columns.csv"), sep=";")
        df_measures = pd.read_csv(os.path.join(pasta_arquivos, "measures.csv"), sep=";")
        df_relationships = pd.read_csv(os.path.join(pasta_arquivos, "relationships.csv"), sep=";")
    except FileNotFoundError:
        return "❌ Arquivos CSV não encontrados. Descrição não gerada."

    # Criar prompt para IA
    prompt = f"""
    Você é um especialista em BI. Abaixo estão informações extraídas de um modelo Power BI.

    Tabelas e Colunas:
    {df_columns[['tableName', 'name']].drop_duplicates().head(20).to_string(index=False)}

    Medidas:
    {df_measures[['name', 'expression']].head(10).to_string(index=False)}

    Relacionamentos:
    {df_relationships[['relationship']].head(10).to_string(index=False)}

    Com base nesses dados, gere uma descrição clara do propósito do dashboard, com insights que ele pode fornecer e temas principais.
    No proposito, inclua informações mais relevantes do tema identificado sobre as tabelas e colunas, medidas e relacionamentos entre as tabelas.
    Não é necessário sugerir visualizações específicas ou melhorias, apenas descreva o que o usuário pode esperar encontrar no dashboard.
    """

    # Enviar para a IA
    try:
        model = genai.GenerativeModel("gemini-1.5-flash")
        response = model.generate_content(prompt)
        return response.text  # Retorna a descrição gerada
    except Exception as e:
        return f"❌ Erro ao gerar texto com IA: {str(e)}"
