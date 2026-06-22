import subprocess
import time

# Inicia o Streamlit
subprocess.Popen(["streamlit", "run", "app.py", "--server.port=8502"])

# Aguarda alguns segundos para garantir que o Streamlit suba primeiro
time.sleep(5)

# Inicia o ngrok
subprocess.Popen(["ngrok", "http", "8502"])
