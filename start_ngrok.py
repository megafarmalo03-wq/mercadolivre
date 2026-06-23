import subprocess
import time
import re
import sys
import os

STREAMLIT_PORT = 8501
# ngrok.exe na mesma pasta deste script
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
NGROK_PATH = os.path.join(SCRIPT_DIR, "ngrok.exe")


def start_streamlit():
    print("Iniciando Streamlit...")
    proc = subprocess.Popen(
        [sys.executable, "-m", "streamlit", "run", "app.py", "--server.port", str(STREAMLIT_PORT), "--server.headless", "true"],
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,
        cwd=os.path.dirname(os.path.abspath(__file__))
    )
    return proc


def wait_for_streamlit():
    print("Aguardando Streamlit subir...")
    for _ in range(30):
        try:
            import socket
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            result = sock.connect_ex(("127.0.0.1", STREAMLIT_PORT))
            sock.close()
            if result == 0:
                print("Streamlit esta rodando em http://127.0.0.1:" + str(STREAMLIT_PORT))
                return True
        except Exception:
            pass
        time.sleep(1)
    return False


def start_ngrok():
    print("Iniciando ngrok...")
    proc = subprocess.Popen(
        [NGROK_PATH, "http", str(STREAMLIT_PORT), "--log", "stdout"],
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1
    )
    return proc


def get_ngrok_url(proc, timeout=60):
    print("Aguardando URL publica do ngrok...")
    start = time.time()
    while time.time() - start < timeout:
        line = proc.stdout.readline()
        if line:
            print(line.strip())
            # Procura por https://xxxx.ngrok.io ou URL
            match = re.search(r"(https://[a-z0-9\-]+\.ngrok\.io|https://[a-z0-9\-]+\.ngrok\.app)", line)
            if match:
                return match.group(1)
        time.sleep(0.1)
    return None


def main():
    # Verifica se ngrok existe
    if not os.path.exists(NGROK_PATH):
        print("ERRO: ngrok.exe nao encontrado em:")
        print(NGROK_PATH)
        print("\nBaixe em: https://ngrok.com/download")
        print("Extraia e ajuste o caminho NGROK_PATH neste script.")
        input("Pressione Enter para sair...")
        return

    # Inicia Streamlit
    streamlit_proc = start_streamlit()

    if not wait_for_streamlit():
        print("ERRO: Streamlit nao subiu.")
        streamlit_proc.terminate()
        return

    # Inicia ngrok
    ngrok_proc = start_ngrok()
    url = get_ngrok_url(ngrok_proc, timeout=60)

    if url:
        print("\n" + "=" * 50)
        print("APP ACESSIVO EXTERNAMENTE:")
        print(url)
        print("=" * 50 + "\n")
    else:
        print("ERRO: Nao foi possivel obter a URL do ngrok.")
        print("Verifique se o authtoken esta configurado:")
        print(f"  {NGROK_PATH} config add-authtoken SEU_TOKEN")

    try:
        while True:
            line = ngrok_proc.stdout.readline()
            if line:
                print(line.strip())
            if streamlit_proc.poll() is not None:
                break
            time.sleep(0.1)
    except KeyboardInterrupt:
        print("\nEncerrando...")
        streamlit_proc.terminate()
        ngrok_proc.terminate()


if __name__ == "__main__":
    main()
