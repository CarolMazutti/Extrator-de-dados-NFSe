import openai
from dotenv import load_dotenv, find_dotenv
import os

# Carregar as variáveis de ambiente do arquivo .env
load_dotenv(find_dotenv())

# Obter a API Key do ambiente
api_key = os.getenv("OPENAI_API_KEY")

# Inicializar o cliente da OpenAI
openai.api_key = api_key

# Testar a API com uma simples requisição
response = openai.ChatCompletion.create(
    model="gpt-3.5-turbo",  # ou "gpt-4" se você tiver acesso
    messages=[
        {"role": "user", "content": "Escreva uma breve história sobre um programador aprendendo Python."}
    ],
    max_tokens=100
)

print(response.choices[0].message['content'].strip())