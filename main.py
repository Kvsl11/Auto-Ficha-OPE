import threading
import tkinter as tk
from tkinter import messagebox
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException, StaleElementReferenceException
import re
import time
import json
from datetime import datetime
import customtkinter as ctk
import os
import ssl
import subprocess
import urllib.request
import logging
import sys
import requests
import tkinter as tk
from tkinter import messagebox
import webbrowser


# --- Verifica e usa Python interno automaticamente ---
app_dir = os.path.dirname(os.path.abspath(__file__))
python_exe = os.path.join(app_dir, "Python313", "python.exe")

# Se o script não estiver rodando pelo Python interno, relança com ele
if "Python313" not in sys.executable and os.path.exists(python_exe):
    print("🟢 Usando Python interno (embutido na pasta)...")
    subprocess.run([python_exe, os.path.abspath(__file__)])
    sys.exit(0)


# Configuração de logs
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)

# Bypass SSL universal (urllib + requests)
ssl._create_default_https_context = ssl._create_unverified_context
requests.packages.urllib3.disable_warnings()

# Certificado Amazon Root
AMAZON_CERTS = {
    "Amazon Root CA 1": "https://www.amazontrust.com/repository/AmazonRootCA1.pem"
}

def atualizar_certifi():
    """Atualiza o pacote certifi usando o Python interno."""
    try:
        if not os.path.exists(python_exe):
            logger.warning(f"⚠️ Python interno não encontrado em: {python_exe}")
            return
        logger.info("🔍 Verificando e atualizando pacote certifi...")
        subprocess.run([python_exe, "-m", "pip", "install", "--upgrade", "certifi"], check=True)
        import certifi
        logger.info(f"🟢 Certifi atualizado com sucesso: {certifi.where()}")
    except Exception as e:
        logger.warning(f"⚠️ Falha ao atualizar certifi: {e}")

def garantir_certificados_amazon():
    """Garante que o Amazon Root CA 1 esteja presente no cacert.pem."""
    try:
        import certifi
        cacert_path = certifi.where()
        with open(cacert_path, "r", encoding="utf-8") as f:
            conteudo = f.read()

        if "Amazon Root CA 1" not in conteudo:
            logger.info("🔍 Amazon Root CA 1 não encontrado — baixando...")
            resp = requests.get(AMAZON_CERTS["Amazon Root CA 1"], timeout=10, verify=False)
            if resp.status_code == 200:
                with open(cacert_path, "a", encoding="utf-8") as f:
                    f.write(f"\n# Amazon Root CA 1\n{resp.text.strip()}\n")
                logger.info("✅ Certificado Amazon Root CA 1 adicionado com sucesso.")
            else:
                logger.warning(f"❌ Falha ao baixar certificado Amazon: {resp.status_code}")
        else:
            logger.info("🟢 Amazon Root CA 1 já está presente.")
    except Exception as e:
        logger.warning(f"⚠️ Falha ao garantir certificados: {e}")

def testar_ssl():
    """Verifica se há conectividade SSL, aplica fallback se falhar."""
    try:
        import certifi
        ssl_context = ssl.create_default_context(cafile=certifi.where())
        urllib.request.urlopen("https://www.google.com", timeout=5, context=ssl_context)
        logger.info("🟢 Conexão SSL validada com sucesso.")
    except ssl.SSLError as e:
        logger.warning(f"⚠️ Falha SSL detectada ({e}). Aplicando modo não verificado.")
        ssl._create_default_https_context = ssl._create_unverified_context
        try:
            urllib.request.urlopen("https://www.google.com", timeout=5)
            logger.info("🟡 SSL desativado — conexão forçada sem verificação de certificado.")
        except Exception as e2:
            logger.error(f"❌ Mesmo após fallback, falhou: {e2}")
    except Exception as e:
        logger.warning(f"⚠️ Erro genérico ao testar SSL: {e}. Aplicando fallback.")
        ssl._create_default_https_context = ssl._create_unverified_context

# Execução automática
logger.info("🚀 Iniciando verificação e correção SSL híbrida...")
atualizar_certifi()
garantir_certificados_amazon()
testar_ssl()
logger.info("✅ Configuração SSL concluída com segurança.")

# --- VERIFICAÇÃO DE ATUALIZAÇÃO VIA GITHUB ---
VERSAO = "4.3.4"

def verificar_e_atualizar_automaticamente():
    """
    Verifica no GitHub se há nova versão e atualiza automaticamente sem interação do usuário.
    """
    try:
        REPO = "Kvsl11/Hxg_auto"
        URL_VERSION = f"https://raw.githubusercontent.com/{REPO}/main/version.txt"
        URL_SCRIPT = f"https://raw.githubusercontent.com/{REPO}/main/main.py"
        LOCAL_SCRIPT = os.path.join(os.path.dirname(__file__), "main.py")
        LOCAL_VERSION_FILE = os.path.join(os.path.dirname(__file__), "version_local.txt")
        LOG_PATH = os.path.join(os.path.dirname(__file__), "autoupdate.log")

        logging.basicConfig(
            filename=LOG_PATH,
            level=logging.INFO,
            format="%(asctime)s - %(levelname)s - %(message)s"
        )

        def get_local_version():
            if os.path.exists(LOCAL_VERSION_FILE):
                try:
                    with open(LOCAL_VERSION_FILE, "r", encoding="utf-8") as f:
                        return f.read().strip()
                except Exception:
                    return "0.0.0"
            return "0.0.0"

        def get_online_version():
            try:
                headers = {"Cache-Control": "no-cache", "Pragma": "no-cache"}
                r = requests.get(URL_VERSION, timeout=10, verify=False, headers=headers)
                if r.status_code == 200:
                    return r.text.strip()
                else:
                    logging.warning(f"⚠️ Falha HTTP ao buscar versão: {r.status_code}")
            except Exception as e:
                logging.warning(f"⚠️ Falha ao obter versão online: {e}")
            return None

        def save_local_version(ver):
            try:
                with open(LOCAL_VERSION_FILE, "w", encoding="utf-8") as f:
                    f.write(ver)
                logging.info(f"✅ Versão local atualizada para {ver}")
            except Exception as e:
                logging.error(f"❌ Erro ao salvar versão local: {e}")

        def atualizar_script(versao_online):
            try:
                headers = {"Cache-Control": "no-cache", "Pragma": "no-cache"}
                r = requests.get(URL_SCRIPT, timeout=20, verify=False, headers=headers)
                r.raise_for_status()
                with open(LOCAL_SCRIPT, "wb") as f:
                    f.write(r.content)
                save_local_version(versao_online)
                logging.info(f"✅ Atualização concluída para a versão {versao_online}")
                return True
            except Exception as e:
                logging.error(f"❌ Falha ao atualizar script: {e}")
                return False

        local_v = get_local_version()
        online_v = get_online_version()

        if not online_v:
            logging.warning("⚠️ Falha ao verificar versão online. Continuando com a versão local.")
            return

        if online_v != local_v:
            logging.info(f"🟡 Nova versão detectada: {online_v} (local: {local_v}) — atualizando...")
            sucesso = atualizar_script(online_v)
            if sucesso:
                logging.info("♻️ Reiniciando app com nova versão...")
                python_exe = sys.executable
                subprocess.Popen([python_exe, LOCAL_SCRIPT])
                os._exit(0)
        else:
            logging.info(f"🟢 Aplicativo já está atualizado ({local_v})")

    except Exception as e:
        logging.error(f"❌ Erro na verificação automática de atualização: {e}")
