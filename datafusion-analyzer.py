# ============================================================
# DataFusion Analyzer — Processing Pipeline
# + Relatório de Vínculos (automático em ALVOS)
# Autor: Braian Rodrigues
# ============================================================

import sys
import traceback
import logging
from logging.handlers import RotatingFileHandler
import tempfile
import platform
import os
import re
import json
import html
import shutil
import zipfile
import threading
import queue
from collections import defaultdict
from datetime import datetime, date
import locale

# ========================
# Outlook / HTML / GUI
# ========================
try:
    import win32com.client
except ImportError:
    win32com = None
from bs4 import BeautifulSoup

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

try:
    from tkcalendar import DateEntry
except Exception:
    DateEntry = None


# ============================================================
# LOG EM ARQUIVO (sempre) + EXCEÇÕES (para debug em qualquer PC)
# ============================================================

def _get_app_dir() -> str:
    # Quando empacotado (PyInstaller), sys.executable aponta para o .exe
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

APP_DIR = _get_app_dir()
LOG_FILE = os.path.join(APP_DIR, "bilhetagem.log")

_logger = logging.getLogger("bilhetagem")
_logger.setLevel(logging.INFO)

try:
    os.makedirs(APP_DIR, exist_ok=True)
    _fh = RotatingFileHandler(LOG_FILE, maxBytes=2_000_000, backupCount=5, encoding="utf-8")
    _fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
    _fh.setFormatter(_fmt)
    _logger.addHandler(_fh)
except Exception:
    # Se o handler falhar, seguimos com a GUI/queue (melhor que quebrar o app)
    _fh = None

def _log_exception(prefix: str, exc: BaseException):
    try:
        _logger.error("%s: %s", prefix, exc)
        _logger.error("".join(traceback.format_exception(type(exc), exc, exc.__traceback__)))
    except Exception:
        pass

def _excepthook(exctype, value, tb):
    try:
        _logger.error("EXCEÇÃO NÃO TRATADA: %s", value)
        _logger.error("".join(traceback.format_exception(exctype, value, tb)))
    except Exception:
        pass

sys.excepthook = _excepthook


# ============================================================
# PYWIN32 / OUTLOOK (COM) — compat p/ EXE (gen_py em pasta gravável)
# ============================================================

def _prepare_pywin32_for_frozen():
    if not getattr(sys, "frozen", False):
        return
    try:
        import win32com
        gen_path = os.path.join(tempfile.gettempdir(), "pywin32_gen_py")
        os.makedirs(gen_path, exist_ok=True)
        win32com.__gen_path__ = gen_path  # type: ignore[attr-defined]
        import win32com.client.gencache as gencache
        gencache.is_readonly = False
    except Exception:
        pass

_prepare_pywin32_for_frozen()


# ============================================================
# LOG GLOBAL (thread-safe) + ARQUIVOS GERADOS
# ============================================================

_log_queue = queue.Queue()
_files_queue = queue.Queue()

log_text = None  # definido na GUI


def log_message(message: str):
    msg = str(message)
    try:
        _log_queue.put(msg)
    except Exception:
        pass
    try:
        _logger.info(msg)
    except Exception:
        pass


def log_header(title: str):
    log_message("")
    log_message("=" * 70)
    log_message(f"{title}")
    log_message("=" * 70)


def log_step(title: str):
    log_message("")
    log_message(f"[PASSO] {title}")


def log_info(msg: str):
    log_message(f"[INFO ] {msg}")


def log_warn(msg: str):
    log_message(f"[AVISO] {msg}")


def log_ok(msg: str):
    log_message(f"[OK   ] {msg}")


def log_error(msg: str):
    log_message(f"[ERRO ] {msg}")


def log_file(path: str, label: str = "Gerado"):
    p = os.path.abspath(path)
    log_message(f"[ARQ  ] {label}: {p}")
    try:
        _files_queue.put(p)
    except Exception:
        pass


def _get_outlook_namespace():
    """Conecta no Outlook (MAPI) com fallback e logs detalhados para PCs problemáticos."""
    try:
        import pythoncom
        pythoncom.CoInitialize()
    except Exception:
        pass

    # Info útil para debug de "só falha em alguns PCs"
    try:
        log_info(f"Python: {platform.python_version()} | Arch: {platform.architecture()[0]}")
        log_info(f"APP_DIR: {APP_DIR}")
        log_info(f"LOG_FILE: {LOG_FILE}")
    except Exception:
        pass

    try:
        import win32com.client
        try:
            from win32com.client import gencache
            outlook_app = gencache.EnsureDispatch("Outlook.Application")
        except Exception:
            outlook_app = win32com.client.Dispatch("Outlook.Application")

        ns = outlook_app.GetNamespace("MAPI")

        # Força inicialização do MAPI/Profile (ajuda em alguns ambientes)
        try:
            ns.Logon("", "", False, False)
        except Exception:
            pass

        return ns
    except Exception as e:
        _log_exception("Falha ao conectar no Outlook (COM/MAPI)", e)
        raise


# ============================================================
# 0) VÍNCULOS — (AUTOMÁTICO EM ALVOS)
# ============================================================

def load_all_data(root_dir: str):
    """
    Percorre todas as subpastas de 'root_dir' e carrega os arquivos data.json.
    Pressupõe que o nome da pasta seja o número-alvo.
    Retorna: {numero_alvo: dados_do_json}
    """
    all_data = {}
    if not root_dir or not os.path.isdir(root_dir):
        return all_data

    for folder in os.listdir(root_dir):
        folder_path = os.path.join(root_dir, folder)
        if not os.path.isdir(folder_path):
            continue
        data_file = os.path.join(folder_path, "data.json")
        if os.path.exists(data_file):
            try:
                with open(data_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                all_data[folder] = data
            except Exception as e:
                log_warn(f"[VÍNCULOS] Erro ao carregar {data_file}: {e}")
    return all_data


def aggregate_correlations_by_contact(all_data: dict):
    """
    Analisa os contatos de cada alvo (data.json) e agrega os contatos encontrados,
    independente de o contato também ser alvo.

    Retorna:
      correlations[contato][alvo_origem] = [categorias...]
    """
    correlations = {}  # contato -> {origin -> set(categorias)}

    for origin, data in all_data.items():
        # simétricos
        for contact in data.get("symmetric_contacts", []):
            if contact != origin:
                correlations.setdefault(contact, {}).setdefault(origin, set()).add("Simétrico")
        # assimétricos
        for contact in data.get("asymmetric_contacts", []):
            if contact != origin:
                correlations.setdefault(contact, {}).setdefault(origin, set()).add("Assimétrico")
        # grupos
        for group_id, participants in data.get("groups", {}).items():
            for contact in participants:
                if contact != origin:
                    correlations.setdefault(contact, {}).setdefault(origin, set()).add(f"Grupo: {group_id}")

    # sets -> lists ordenadas
    for contact, origins in correlations.items():
        for origin in origins:
            origins[origin] = sorted(list(origins[origin]))
    return correlations


def generate_correlations_by_contact_html_report(correlations: dict, output_html: str, title_suffix: str = ""):
    """
    Gera HTML com correlações (somente contatos que aparecem em >= 2 alvos).
    correlations esperado no formato:
      {
        "5511...": {
            "ALVO1": {"SIMETRICO", "ASSIMETRICO", ...}  # ou lista/set
            "ALVO2": {"SIMETRICO", ...}
        },
        ...
      }
    """
    # filtra somente contatos que aparecem em >= 2 alvos (origens)
    filtered_correlations = {
        contact: origins
        for contact, origins in correlations.items()
        if isinstance(origins, dict) and len(origins) >= 2
    }

    title = "Relatório de Vínculos Entre Números"
    if title_suffix:
        title = f"{title} — {title_suffix}"

    html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>{title}</title>
    <style>
        body {{
            font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
            background: #f5f7fa;
            color: #333;
            margin: 0;
            padding: 20px;
        }}
        h1 {{
            text-align: center;
            margin-bottom: 30px;
            color: #2c3e50;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin: 20px auto;
            background: #fff;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
        }}
        th, td {{
            padding: 12px 15px;
            border: 1px solid #ddd;
            text-align: left;
            font-size: 14px;
        }}
        th {{
            background-color: #2980b9;
            color: #fff;
            font-weight: 600;
            text-transform: uppercase;
        }}
        tr:nth-child(even) {{
            background-color: #f2f2f2;
        }}
        tr:hover {{
            background-color: #e9f1f7;
        }}
        .section {{
            margin-bottom: 40px;
        }}
    </style>
</head>
<body>
    <h1>{title}</h1>
    <table>
        <tr>
            <th>Número Correlacionado</th>
            <th>Números-Alvo de Origem e Categorias</th>
        </tr>
"""

    # ordena corretamente por número correlacionado
    for contact, origins in sorted(filtered_correlations.items(), key=lambda x: x[0]):
        details = ""
        # ordena origens (alvos)
        for origin, categories in sorted(origins.items(), key=lambda x: x[0]):
            # categories pode ser set/list/tuple/string
            if isinstance(categories, (set, list, tuple)):
                cat_str = ", ".join(sorted(map(str, categories)))
            else:
                cat_str = str(categories)
            details += f"<strong>{origin}</strong>: {cat_str}<br>"

        html += f"""
        <tr>
            <td>{contact}</td>
            <td>{details}</td>
        </tr>
"""

    html += """
    </table>
</body>
</html>
"""

    with open(output_html, "w", encoding="utf-8") as f:
        f.write(html)


def generate_vinculos_report_for_alvos(alvos_root: str, output_html: str, title_suffix: str = "") -> bool:
    """
    Gera relatório de vínculos usando TODOS os data.json existentes em ALVOS.
    Retorna True se gerou.
    """
    all_data = load_all_data(alvos_root)
    if not all_data:
        log_warn("[VÍNCULOS] Nenhum data.json encontrado em ALVOS (relatório não gerado).")
        return False

    correlations = aggregate_correlations_by_contact(all_data)
    if not correlations:
        log_warn("[VÍNCULOS] Nenhuma correlação encontrada (relatório não gerado).")
        return False

    generate_correlations_by_contact_html_report(correlations, output_html, title_suffix=title_suffix)
    return True

# ============================================================
# 1) OUTLOOK -> TXT (E-MAILS DO CASE)
# ============================================================

def sanitize_filename(name: str) -> str:
    return re.sub(r'[\\/*?:"<>|]+', "_", name)


def fetch_outlook_emails(sender_email, case_number, selected_date, output_directory, log):
    try:
        if isinstance(selected_date, datetime):
            selected_date = selected_date.date()
        elif isinstance(selected_date, date):
            pass
        elif isinstance(selected_date, str):
            try:
                selected_date = datetime.strptime(selected_date, "%d/%m/%Y").date()
            except ValueError as e:
                raise ValueError(
                    f"selected_date string inválida: {selected_date}. Esperado dd/mm/YYYY"
                ) from e
        else:
            raise ValueError("selected_date deve ser date, datetime ou string no formato dd/mm/YYYY")

        case_number = str(case_number).replace("Case #", "").strip()
        sender_email = sender_email.replace("<", "").replace(">", "").strip()

        log_step("1/4 — Buscando e-mails no Outlook")
        log_info(f"Remetente filtrado: {sender_email}")
        log_info(f"Case: {case_number}")
        log_info(f"Data: {selected_date.strftime('%d/%m/%Y')}")
        log_info(f"Saída: {output_directory}")

        log_info("Conectando ao Outlook (MAPI)...")
        try:
            outlook = _get_outlook_namespace()
        except Exception as e:
            log_error("Não foi possível conectar no Outlook neste computador.")
            log_error("Dicas: 1) Abra o Outlook manualmente; 2) Confirme o perfil; 3) Verifique se Outlook e o EXE são 32/64-bit compatíveis.")
            log_error(f"Detalhes: {e}")
            try:
                messagebox.showerror(
                    "Falha no Outlook",
                    "Falha ao conectar no Outlook.\\n\\n"
                    "✅ Verifique:\\n"
                    "1) Outlook instalado e perfil configurado\\n"
                    "2) Outlook aberto pelo menos 1 vez\\n"
                    "3) Arquitetura compatível (32/64-bit)\\n\\n"
                    f"Log salvo em: {LOG_FILE}"
                )
            except Exception:
                pass
            return

        inbox = outlook.Folders.Item(1).Folders["Caixa de Entrada"]
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)

        case_pattern = re.compile(rf"Case\s*#\s*{re.escape(case_number)}", re.IGNORECASE)

        filtered_messages = []
        for msg in messages:
            try:
                if not hasattr(msg, "SenderEmailAddress"):
                    continue

                sender = str(msg.SenderEmailAddress).lower()
                if sender_email.lower() not in sender:
                    continue

                email_date = msg.ReceivedTime.date()
                subject = msg.Subject or ""

                if email_date == selected_date and case_pattern.search(subject):
                    filtered_messages.append(msg)

            except Exception:
                continue

        log_info(f"E-mails encontrados para o filtro: {len(filtered_messages)}")

        date_folder_name = selected_date.strftime("%d-%m-%Y")
        date_folder = os.path.join(output_directory, date_folder_name)
        os.makedirs(date_folder, exist_ok=True)

        removed = 0
        for item in os.listdir(date_folder):
            if item.lower().endswith(".txt"):
                try:
                    os.remove(os.path.join(date_folder, item))
                    removed += 1
                except Exception:
                    pass
        if removed:
            log_info(f"Limpeza: {removed} TXT antigos removidos em {date_folder_name}")

        for msg in filtered_messages:
            subject = msg.Subject or ""
            body = msg.Body or msg.HTMLBody or ""
            if not body.strip():
                body = "O corpo do e-mail está vazio."
            else:
                lines = [line.strip() for line in body.splitlines()]
                body = "\n".join([line for line in lines if line])

            received_time = msg.ReceivedTime.strftime("%d/%m/%Y %H:%M")
            message_size = getattr(msg, "Size", "N/A")

            header = (
                f"Data de Recebimento: {received_time}\n"
                f"Assunto: {subject}\n"
                f"Tamanho (bytes): {message_size}\n"
                f"Corpo da Mensagem:\n"
            )

            safe_subject = sanitize_filename(subject.replace(" ", "_"))
            timestamp = msg.ReceivedTime.strftime("%Y-%m-%d_%H-%M-%S")

            file_name = f"{safe_subject[:50]}_{timestamp}.txt"
            file_path = os.path.join(date_folder, file_name)

            with open(file_path, "w", encoding="utf-8") as f:
                f.write(header + body)

            log_file(file_path, "TXT do e-mail salvo")

        if not filtered_messages:
            log_warn("Nenhum e-mail encontrado para esse filtro.")
        else:
            log_ok(f"{len(filtered_messages)} e-mail(s) exportados para TXT.")

        return len(filtered_messages)

    except Exception as e:
        log_error(f"[EMAIL][ERRO FATAL] {e}")
        return 0


def fetch_emails_for_case(sender_email, case_number, selected_date, emails_dir):
    qtd = fetch_outlook_emails(
        sender_email=sender_email,
        case_number=case_number,
        selected_date=selected_date,
        output_directory=emails_dir,
        log=log_message
    )

    if qtd == 0:
        return {}

    case_clean = str(case_number).replace("Case #", "").strip()
    date_folder = os.path.join(emails_dir, selected_date.strftime("%d-%m-%Y"))

    if not os.path.isdir(date_folder):
        log_warn(f"Pasta de data não encontrada: {date_folder}")
        return {}

    log_step("2/4 — Separando e-mails por alvo (Account Identifier)")
    targets_files = defaultdict(list)

    for fname in os.listdir(date_folder):
        if not fname.lower().endswith(".txt"):
            continue
        path = os.path.join(date_folder, fname)
        try:
            with open(path, "r", encoding="utf-8", errors="ignore") as f:
                content = f.read()
        except Exception as e:
            log_warn(f"Erro lendo {path}: {e}")
            continue

        m_case = re.search(r"Case #(\d+)", content)
        if not m_case:
            log_warn(f"Case não encontrado em {fname}, ignorando.")
            continue
        if m_case.group(1) != case_clean:
            log_warn(f"Case diferente em {fname} ({m_case.group(1)} != {case_clean}), ignorando.")
            continue

        m_target = re.search(r'Account Identifier\s+([\+\d]+)', content)
        if not m_target:
            log_warn(f"Account Identifier não encontrado em {fname}, ignorando.")
            continue

        target = m_target.group(1).replace("+", "")
        targets_files[target].append(path)

    log_ok(f"Alvos detectados: {len(targets_files)}")
    for alvo, files in targets_files.items():
        log_info(f"Alvo {alvo}: {len(files)} TXT(s) legal process(es)")

    return targets_files


# ============================================================
# 2) RECORDS.HTML -> TEXTO LIMPO (somente trechos necessários)
# ============================================================

def convert_html_to_text(html_file: str) -> str:
    with open(html_file, "r", encoding="utf-8", errors="ignore") as f:
        soup = BeautifulSoup(f, "html.parser")
    extracted = soup.get_text(separator="\n")
    processed = "\n".join(line for line in extracted.splitlines() if line.strip())
    return processed


def _best_profile_picture_from_records_html(records_html_path: str) -> str:
    # lê o HTML cru porque o soup.get_text() NÃO preserva <img src="...">,
    # e alguns records só trazem a foto assim.
    try:
        with open(records_html_path, "r", encoding="utf-8", errors="ignore") as f:
            html = f.read()
    except Exception:
        return ""

    # pega TODAS as referências (records pode ter thumbs 96x96 + a foto real 640x640)
    candidates = re.findall(r"linked_media[\/]+profile_picture_\d+\.(?:jpe?g|png|webp)", html, flags=re.IGNORECASE)
    if not candidates:
        return ""

    # dedup preservando ordem
    seen = set()
    uniq = []
    for c in candidates:
        cl = c.lower()
        if cl not in seen:
            uniq.append(c)
            seen.add(cl)

    base_dir = os.path.dirname(os.path.abspath(records_html_path))

    # escolhe a melhor por tamanho de arquivo (a real geralmente é bem maior que os thumbs)
    best = ""
    best_size = -1
    for rel in uniq:
        rel_norm = rel.replace("/", os.sep)
        abs_path = os.path.join(base_dir, rel_norm)
        try:
            size = os.path.getsize(abs_path)
        except Exception:
            size = -1

        if size > best_size:
            best_size = size
            best = rel

    return best


def _extract_profile_picture_block_from_records(records_html_path: str):
    best = _best_profile_picture_from_records_html(records_html_path)
    if not best:
        return []
    return [
        "Profile Picture",
        "Linked Media File:",
        best
    ]


def _is_page_marker(s: str) -> bool:
    s = (s or "").strip()
    return s.startswith("WhatsApp Business Record Page")


def _looks_like_number(s: str) -> bool:
    s = (s or "").strip()
    return bool(re.fullmatch(r"\d{8,15}", s))


# ========================
# PHONE HELPERS (robusto)
# ========================

_PHONE_DIGITS_RE = re.compile(r"\D+")

def phone_digits(value) -> str:
    if value is None:
        return ""
    return _PHONE_DIGITS_RE.sub("", str(value))

def norm_phone(value) -> str:
    d = phone_digits(value)
    return d if bool(re.fullmatch(r"\d{8,15}", d)) else ""


def _extract_contacts_block(lines):
    out = []
    n = len(lines)

    def collect_list(start_idx, label):
        buf = [label]
        i = start_idx
        while i < n:
            s = lines[i].strip()
            if not s:
                i += 1
                continue
            if s in ("Symmetric contacts", "Asymmetric contacts") and s != label:
                break
            if s in ("Address Book", "Groups", "Profile Picture", "Connection", "Web Info",
                     "User Notes", "Small Medium Business", "Device Info"):
                break

            if _is_page_marker(s):
                buf.append(s)
                i += 1
                continue

            if s.endswith("Total"):
                buf.append(s)
                i += 1
                continue

            if _looks_like_number(s):
                buf.append(s)
                i += 1
                continue

            if re.fullmatch(r"\d+\s+Total", s):
                buf.append(s)
                i += 1
                continue

            i += 1

        return buf, i

    i = 0
    while i < n:
        s = lines[i].strip()
        if s == "Address Book":
            out.append("Address Book")
            i += 1
            break
        i += 1

    if not out:
        return []

    while i < n:
        s = lines[i].strip()
        if not s:
            i += 1
            continue

        if s == "Symmetric contacts":
            block, i2 = collect_list(i + 1, "Symmetric contacts")
            out.extend(block)
            i = i2
            continue

        if s == "Asymmetric contacts":
            block, i2 = collect_list(i + 1, "Asymmetric contacts")
            out.extend(block)
            i = i2
            continue

        if s in ("Groups", "Profile Picture", "Connection", "Web Info",
                 "User Notes", "Small Medium Business", "Device Info"):
            break

        i += 1

    return out


def _extract_profile_picture_block(lines):
    out = []
    n = len(lines)
    i = 0

    while i < n:
        s = lines[i].strip()
        if s == "Profile Picture":
            out.append("Profile Picture")
            i += 1

            while i < n and _is_page_marker(lines[i].strip()):
                out.append(lines[i].strip())
                i += 1

            while i < n:
                s2 = lines[i].strip()
                if not s2:
                    i += 1
                    continue

                if s2.startswith("Linked Media File"):
                    out.append("Linked Media File:")
                    parts = s2.split(":", 1)
                    if len(parts) == 2 and parts[1].strip():
                        out.append(parts[1].strip())
                        i += 1
                    else:
                        i += 1
                        while i < n and not lines[i].strip():
                            i += 1
                        if i < n and not _is_page_marker(lines[i].strip()):
                            out.append(lines[i].strip())
                            i += 1
                    break

                if s2 in ("Groups", "Address Book", "Connection", "Web Info",
                          "User Notes", "Small Medium Business", "Device Info"):
                    break

                i += 1

            j = i
            while j < n:
                sj = lines[j].strip()
                if not sj:
                    j += 1
                    continue

                if sj == "Push Name":
                    out.append("Push Name")
                    j += 1
                    while j < n and not lines[j].strip():
                        j += 1
                    if j < n and not _is_page_marker(lines[j].strip()):
                        out.append(lines[j].strip())
                        j += 1
                    break

                if sj.startswith("Push Name"):
                    out.append("Push Name")
                    val = sj.split("Push Name", 1)[1].strip(" :")
                    if val:
                        out.append(val)
                    break

                if sj in ("Groups", "Address Book", "Connection", "Web Info",
                          "User Notes", "Small Medium Business", "Device Info"):
                    break

                j += 1

            break

        i += 1

    return out


def _extract_groups_block(lines):
    out = []
    n = len(lines)
    i = 0

    while i < n and lines[i].strip() != "Groups":
        i += 1
    if i >= n:
        return []

    out.append("Groups")
    i += 1

    while i < n:
        s = lines[i].strip()
        if not s:
            i += 1
            continue
        if s == "Participating Groups":
            out.append("Participating Groups")
            i += 1
            break
        if s in ("Address Book", "Profile Picture", "Connection", "Web Info",
                 "User Notes", "Small Medium Business", "Device Info"):
            return out
        i += 1

    def read_value_after(idx):
        j = idx
        while j < n and (not lines[j].strip() or _is_page_marker(lines[j].strip())):
            if _is_page_marker(lines[j].strip()):
                out.append(lines[j].strip())
            j += 1
        if j < n:
            return lines[j].strip(), j + 1
        return "", j

    while i < n:
        s = lines[i].strip()

        if not s:
            i += 1
            continue

        if s in ("Address Book", "Profile Picture", "Connection", "Web Info",
                 "User Notes", "Small Medium Business", "Device Info"):
            break

        if _is_page_marker(s):
            out.append(s)
            i += 1
            continue

        if s == "Picture":
            out.append("Picture")
            i += 1

            while i < n:
                s2 = lines[i].strip()

                if not s2:
                    i += 1
                    continue

                if _is_page_marker(s2):
                    out.append(s2)
                    i += 1
                    continue

                if s2 == "Picture":
                    break

                if s2 in ("Address Book", "Profile Picture", "Connection", "Web Info",
                          "User Notes", "Small Medium Business", "Device Info"):
                    break

                if s2.startswith("Linked Media File"):
                    out.append("Linked Media File:")
                    parts = s2.split(":", 1)
                    if len(parts) == 2 and parts[1].strip():
                        out.append(parts[1].strip())
                        i += 1
                    else:
                        val, i = read_value_after(i + 1)
                        if val:
                            out.append(val)
                    continue

                if s2 in ("ID", "Creation", "Size", "Description", "Subject"):
                    out.append(s2)
                    val, i = read_value_after(i + 1)
                    if val:
                        out.append(val)
                    if s2 == "Subject":
                        break
                    continue

                i += 1

            continue

        i += 1

    return out


def extract_records_block(records_html_path: str) -> str:
    if not os.path.exists(records_html_path):
        raise FileNotFoundError(f"records.html não encontrado: {records_html_path}")

    text = convert_html_to_text(records_html_path)
    lines = text.splitlines()

    parts = []
    contacts = _extract_contacts_block(lines)
    if contacts:
        parts.append("\n".join(contacts))

    profile = _extract_profile_picture_block(lines)
    raw_profile = _extract_profile_picture_block_from_records(records_html_path)
    if profile:
        # se o bloco do texto não trouxer o caminho do arquivo (alguns records só têm <img src="...">), faz merge com o HTML cru
        if raw_profile:
            have = set(x.strip().lower() for x in profile)
            for x in raw_profile:
                xl = x.strip().lower()
                if xl and xl not in have:
                    profile.append(x)
                    have.add(xl)
        parts.append("\n".join(profile))
    elif raw_profile:
        parts.append("\n".join(raw_profile))

    groups = _extract_groups_block(lines)
    if groups:
        parts.append("\n".join(groups))

    out = "\n\n".join([p for p in parts if p.strip()])

    final_lines = []
    prev = None
    for ln in out.splitlines():
        if prev is not None and ln.strip() == prev.strip() and ln.strip() in (
            "Profile Picture", "Symmetric contacts", "Asymmetric contacts", "Groups", "Participating Groups", "Picture"
        ):
            continue
        final_lines.append(ln)
        prev = ln

    return "\n".join(final_lines).strip() + "\n"


# ============================================================
# 3) BILHETAGEM – FUNÇÕES
# ============================================================

def skip_empty_and_whatsapp_pages(lines, idx):
    n = len(lines)
    while idx < n:
        s = lines[idx].strip()
        if s and not s.startswith("WhatsApp Business Record Page"):
            break
        idx += 1
    return idx


def extract_group_media_info(lines):
    group_media_info = []
    i = 0
    n = len(lines)
    while i < n:
        line = lines[i].strip()
        if line == "Profile Picture":
            i += 1
            while i < n and not lines[i].strip().startswith("Picture"):
                i += 1
            continue
        if line == "Picture":
            media_info = {
                "Picture": "No picture",
                "Thumbnail": "No thumbnail",
                "ID": "Não informado",
                "Creation": "Não informado",
                "Size": "Não informado",
                "Description": "Não informado",
                "Subject": "Não informado",
                "Linked Media File": "",
            }
            i += 1
            i = skip_empty_and_whatsapp_pages(lines, i)
            if i < n:
                possible_value = lines[i].strip()
                if possible_value not in (
                    "Thumbnail", "ID", "Creation", "Size", "Description", "Subject",
                    "Linked Media File:", "Profile Picture", "Picture"
                ):
                    media_info["Picture"] = possible_value
                    i += 1

            while i < n:
                i = skip_empty_and_whatsapp_pages(lines, i)
                if i >= n:
                    break
                next_line = lines[i].strip()
                if next_line in ("Picture", "Profile Picture"):
                    break

                if next_line.startswith("Linked Media File:"):
                    i += 1
                    i = skip_empty_and_whatsapp_pages(lines, i)
                    if i < n:
                        val = lines[i].strip()
                        if val not in (
                            "Thumbnail", "ID", "Creation", "Size", "Description", "Subject",
                            "Linked Media File:", "Profile Picture", "Picture", ""
                        ):
                            media_info["Linked Media File"] = val
                            i += 1
                        else:
                            media_info["Linked Media File"] = "Não informado"
                    continue

                if next_line in ("Thumbnail", "ID", "Creation", "Size", "Subject", "Description"):
                    field = next_line
                    i += 1
                    i = skip_empty_and_whatsapp_pages(lines, i)
                    if i >= n or lines[i].strip() in (
                        "Thumbnail", "ID", "Creation", "Size", "Description", "Subject",
                        "Linked Media File:", "Profile Picture", "Picture", ""
                    ):
                        media_info[field] = "Não informado"
                        continue

                    if field == "Description":
                        desc_lines = []
                        while i < n:
                            tmp = lines[i].strip()
                            if tmp in (
                                "Thumbnail", "ID", "Creation", "Size", "Description", "Subject",
                                "Linked Media File:", "Profile Picture", "Picture"
                            ) or tmp == "":
                                break
                            if tmp and not tmp.startswith("WhatsApp Business Record Page"):
                                desc_lines.append(tmp)
                            i += 1
                        media_info["Description"] = "\n".join(desc_lines) if desc_lines else "Não informado"
                    else:
                        possible_val = lines[i].strip()
                        if possible_val in (
                            "Thumbnail", "ID", "Creation", "Size", "Description", "Subject",
                            "Linked Media File:", "Profile Picture", "Picture", ""
                        ):
                            media_info[field] = "Não informado"
                        else:
                            media_info[field] = possible_val
                            i += 1
                    continue

                i += 1

            group_media_info.append(media_info)
        else:
            i += 1
    return group_media_info


def extract_profile_picture_info(lines):
    profile_info = {}
    i = 0
    n = len(lines)
    while i < n:
        line = lines[i].strip()
        if line == "Profile Picture":
            i += 1
            i = skip_empty_and_whatsapp_pages(lines, i)
            if i < n and lines[i].strip().startswith("Linked Media File:"):
                parts = lines[i].strip().split(":", 1)
                if len(parts) > 1 and parts[1].strip():
                    profile_info["Linked Media File"] = parts[1].strip()
                else:
                    i += 1
                    i = skip_empty_and_whatsapp_pages(lines, i)
                    if i < n:
                        profile_info["Linked Media File"] = lines[i].strip()
                i += 1

            i = skip_empty_and_whatsapp_pages(lines, i)
            if i < n and lines[i].strip() == "Push Name":
                i += 1
                i = skip_empty_and_whatsapp_pages(lines, i)
                if i < n:
                    profile_info["Push Name"] = lines[i].strip()
            elif i < n and lines[i].strip().startswith("Push Name"):
                _, push_value = lines[i].strip().split("Push Name", 1)
                profile_info["Push Name"] = push_value.strip()
            break
        i += 1
    return profile_info if profile_info else None


def extract_call_logs(lines):
    call_logs = {}
    current_call_id = None
    current_event = {}
    for line in lines:
        l = line.strip()
        if l.startswith("Call") and "Call Id" in l:
            if current_call_id and current_event:
                call_logs[current_call_id]["events"].append(current_event)
                current_event = {}
            parts = l.split()
            if "Id" in parts:
                idx = parts.index("Id")
                if idx + 1 < len(parts):
                    current_call_id = parts[idx + 1]
                    if current_call_id not in call_logs:
                        call_logs[current_call_id] = {"call_id": current_call_id, "events": []}
            continue

        if current_call_id is None:
            continue

        if l.startswith("Events"):
            parts = l.split()
            if len(parts) >= 3:
                current_event["Type"] = parts[2].lower()
            continue

        if l.startswith("Type") and not l.startswith("Call"):
            parts = l.split(None, 1)
            if len(parts) > 1:
                if "Timestamp" in current_event:
                    call_logs[current_call_id]["events"].append(current_event)
                    current_event = {}
                current_event["Type"] = parts[1].strip().lower()
            continue

        if l.startswith("Timestamp"):
            parts = l.split(None, 1)
            if len(parts) > 1:
                dt_str = parts[1].strip()
                dt_format = "%Y-%m-%d %H:%M:%S UTC"
                try:
                    dt_parsed = datetime.strptime(dt_str, dt_format)
                    dt_final = dt_parsed.strftime("%d-%m-%Y %H:%M:%S UTC")
                    current_event["Timestamp"] = dt_final
                except Exception:
                    current_event["Timestamp"] = dt_str
            continue

        if l.startswith("From") and not l.startswith("From Ip") and not l.startswith("From Port"):
            parts = l.split(None, 1)
            if len(parts) > 1:
                current_event["From"] = parts[1].strip()
            continue

        if l.startswith("To") and not l.startswith("To Ip") and not l.startswith("To Port"):
            parts = l.split(None, 1)
            if len(parts) > 1:
                current_event["To"] = parts[1].strip()
            continue

        if l.startswith("Media Type"):
            parts = l.split(None, 2)
            if len(parts) == 3:
                current_event["Media Type"] = parts[2].strip()
            continue

    if current_call_id and current_event:
        call_logs[current_call_id]["events"].append(current_event)
    return list(call_logs.values())


def extract_target_and_case_info(file_content):
    target_number_match = re.search(r'Account Identifier\s+([\+\d]+)', file_content)
    case_number_match = re.search(r'Case #(\d+)', file_content)
    if target_number_match and case_number_match:
        target_number = target_number_match.group(1).replace("+", "")
        case_number = case_number_match.group(1)
        return target_number, case_number
    raise ValueError("Número do alvo ou case não encontrados no arquivo.")


def extract_contacts(lines, start_keyword):
    contacts = []
    inside_section = False
    for line in lines:
        if line.strip().endswith("Total") or line.startswith("WhatsApp Business Record Page"):
            continue
        if line.startswith(start_keyword):
            inside_section = True
            continue
        if inside_section:
            stripped_line = line.strip()
            if stripped_line.isdigit():
                contacts.append(stripped_line)
            elif not stripped_line or not stripped_line.isdigit():
                break
    return contacts


def save_data(file_path, group_recipients, symmetric_contacts, asymmetric_contacts, group_media_info, profile_info):
    if os.path.exists(file_path):
        with open(file_path, 'r', encoding='utf-8') as f:
            existing_data = json.load(f)
    else:
        existing_data = {}

    existing_data['groups'] = {**existing_data.get('groups', {}), **group_recipients}
    existing_data['symmetric_contacts'] = list(
        set(existing_data.get('symmetric_contacts', [])) | set(symmetric_contacts)
    )
    existing_data['asymmetric_contacts'] = list(
        set(existing_data.get('asymmetric_contacts', [])) | set(asymmetric_contacts)
    )

    existing_data['group_media_info'] = existing_data.get('group_media_info', []) + group_media_info
    try:
        existing_data['group_media_info'] = [dict(t) for t in {tuple(d.items()) for d in existing_data['group_media_info']}]
    except Exception:
        pass

    if profile_info:
        existing_data['profile'] = profile_info

    with open(file_path, 'w', encoding='utf-8') as f:
        json.dump(existing_data, f, ensure_ascii=False, indent=4)


def load_data(file_path):
    if os.path.exists(file_path):
        with open(file_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {
        'groups': {},
        'symmetric_contacts': [],
        'asymmetric_contacts': [],
        'group_media_info': [],
        'profile': {}
    }


event_map = {
    "offer": "Convite de Chamada",
    "accept": "Aceitou Chamada",
    "terminate": "Chamada Encerrada",
    "reject": "Rejeitou Chamada",
    "av_switch": "Mudança de Áudio/Video",
    "group_update": "Alteração de Grupo"
}


def format_date(date_str):
    try:
        dt = datetime.strptime(date_str, "%Y-%m-%d %H:%M:%S UTC")
        return dt.strftime("%d-%m-%Y %H:%M:%S UTC")
    except Exception:
        return date_str


def process_file(txt_path, output_dir, selected_date):
    try:
        with open(txt_path, 'r', encoding='utf-8', errors='ignore') as file:
            file_content = file.read()

        target_number, case_number = extract_target_and_case_info(file_content)
        log_info(f"[BILHETAGEM] Alvo: {target_number} | Case: {case_number}")

        target_dir = os.path.join(output_dir, target_number)
        os.makedirs(target_dir, exist_ok=True)

        json_path = os.path.join(target_dir, 'data.json')
        data = load_data(json_path)

        symmetric_contacts = data.get('symmetric_contacts', [])
        asymmetric_contacts = data.get('asymmetric_contacts', [])
        previous_group_recipients = data.get('groups', {})
        group_media_info = data.get('group_media_info', [])
        if not isinstance(group_media_info, list):
            group_media_info = list(group_media_info)

        conversations = defaultdict(list)
        group_conversations = defaultdict(list)
        total_interactions = defaultdict(int)
        total_group_interactions = defaultdict(int)
        current_group_recipients = defaultdict(set)

        lines = file_content.splitlines()

        profile_picture_info = extract_profile_picture_info(lines)
        groups_extracted = extract_group_media_info(lines)
        if groups_extracted:
            group_media_info.extend(groups_extracted)

        call_logs = extract_call_logs(lines)

        current_message = {}
        current_group_id = None

        new_symmetric_contacts = extract_contacts(lines, 'Symmetric contacts')
        new_asymmetric_contacts = extract_contacts(lines, 'Asymmetric contacts')
        if new_symmetric_contacts:
            symmetric_contacts = list(set(symmetric_contacts) | set(new_symmetric_contacts))
        if new_asymmetric_contacts:
            asymmetric_contacts = list(set(asymmetric_contacts) | set(new_asymmetric_contacts))

        def save_current_message():
            nonlocal current_message, current_group_id

            if not current_message:
                return

            sender_raw = current_message.get('Sender', 'Não informado')
            recipients_raw = current_message.get('Recipients', [])
            msg_type = current_message.get('Type', 'Não informado')
            msg_style = current_message.get('Message Style', 'Não informado')
            timestamp = format_date(current_message.get('Timestamp', 'Não informado'))

            sender = norm_phone(sender_raw)
            recipients = [p for p in (norm_phone(x) for x in recipients_raw) if p]

            for number in recipients:
                if (number != target_number and number not in symmetric_contacts and number not in asymmetric_contacts):
                    asymmetric_contacts.append(number)

            if sender and sender != target_number and sender not in symmetric_contacts and sender not in asymmetric_contacts:
                asymmetric_contacts.append(sender)

            if sender == target_number or target_number in recipients:
                if current_group_id:
                    group_conversations[current_group_id].append({
                        'Sender': sender,
                        'Type': msg_type,
                        'Message Style': msg_style,
                        'Timestamp': timestamp
                    })
                    total_group_interactions[current_group_id] += 1
                    current_group_recipients[current_group_id].update(recipients)
                else:
                    other_participants = [num for num in recipients if num and num != target_number]
                    if not other_participants:
                        other_participants.append(sender)

                    for other_party in other_participants:
                        if not other_party:
                            continue
                        conversations[other_party].append({
                            'Sender': sender,
                            'Recipients': recipients,
                            'Type': msg_type,
                            'Message Style': msg_style,
                            'Timestamp': timestamp
                        })
                        total_interactions[other_party] += 1

            current_message.clear()

        # ======================================================
        # Parsing robusto do TXT (TAB-separated)
        # - Evita confundir "Sender Device" com "Sender"
        # - Evita confundir "Message<TAB>Timestamp" (header) com "Message Timestamp" (campo)
        # ======================================================
        for raw_line in lines:
            line = raw_line.rstrip("\n")

            if line.startswith("Message\tTimestamp"):
                save_current_message()
                ts = line.split("\t")[-1].strip()
                current_message = {'Timestamp': ts}
                current_group_id = None
                continue

            if "\t" not in line:
                continue

            key, value = line.split("\t", 1)
            key = (key or "").strip()
            value = (value or "").strip()

            key_l = key.lower().strip()
            key_l = re.sub(r"\s+", " ", key_l)

            if key == "Message" and value.startswith("Timestamp"):
                continue

            if key in ("Message Timestamp", "Message Timestamp "):
                save_current_message()
                current_message = {'Timestamp': value}
                current_group_id = None
                continue

            if key == "Message Id":
                current_message['Message Id'] = value
                continue

            if key_l == "sender":
                current_message['Sender'] = value
                continue

            if key_l == "recipients":
                if value:
                    recips = [v.strip() for v in value.split(",") if v.strip()]
                    current_message['Recipients'] = recips
                else:
                    current_message['Recipients'] = []
                continue

            if key_l == "type":
                current_message['Type'] = value or 'Não informado'
                continue

            if key_l in ("message style","message_style","message  style"):
                current_message['Message Style'] = value or 'Não informado'
                continue

            if key_l in ("group id","groupid","group_id"):
                current_group_id = value
                current_message['Group Id'] = current_group_id
                continue

        save_current_message()

        group_changes = {}
        if previous_group_recipients:
            for group_id, current_recipients in current_group_recipients.items():
                if group_id not in previous_group_recipients:
                    continue
                previous_recipients = set(previous_group_recipients[group_id])
                current_recipients_set = set(current_recipients)
                new_numbers = current_recipients_set - previous_recipients
                removed_numbers = previous_recipients - current_recipients_set
                if new_numbers or removed_numbers:
                    group_changes[group_id] = {"new_numbers": new_numbers, "removed_numbers": removed_numbers}

        save_data(
            json_path,
            {k: list(v) for k, v in current_group_recipients.items()},
            symmetric_contacts,
            asymmetric_contacts,
            group_media_info,
            profile_picture_info
        )
        log_file(json_path, "JSON salvo")

        individual_counts = {num: len(msgs) for num, msgs in conversations.items()}
        aggregated_file = os.path.join(target_dir, "aggregated_stats.json")

        def update_aggregated_stats(aggregated_file, individual_counts, call_logs, group_messages_counts, target_number):
            if os.path.exists(aggregated_file):
                with open(aggregated_file, 'r', encoding='utf-8') as f:
                    agg_data = json.load(f)
            else:
                agg_data = {}

            agg_data.setdefault("individual_conversations", {"total": 0, "por_numero": {}})
            agg_data.setdefault("calls_by_number", {})
            agg_data.setdefault("group_messages", {})

            for number, count in individual_counts.items():
                agg_data["individual_conversations"]["por_numero"].setdefault(number, 0)
                agg_data["individual_conversations"]["por_numero"][number] += count
                agg_data["individual_conversations"]["total"] += count

            for call in call_logs:
                numbers = set()
                for event in call.get("events", []):
                    if event.get("From"):
                        numbers.add(event.get("From"))
                    if event.get("To"):
                        numbers.add(event.get("To"))
                numbers.discard(target_number)
                for num in numbers:
                    agg_data["calls_by_number"].setdefault(num, 0)
                    agg_data["calls_by_number"][num] += 1

            for group_id, count in group_messages_counts.items():
                agg_data["group_messages"].setdefault(group_id, 0)
                agg_data["group_messages"][group_id] += count

            with open(aggregated_file, 'w', encoding='utf-8') as f:
                json.dump(agg_data, f, ensure_ascii=False, indent=4)

        update_aggregated_stats(aggregated_file, individual_counts, call_logs, total_group_interactions, target_number)
        log_file(aggregated_file, "Estatísticas agregadas")

        # (HTML individual mantido - igual ao antigo)
        data = load_data(json_path)
        profile_data = data.get("profile", {})
        push_name = profile_data.get("Push Name", "Não Informado")
        linked_media = profile_data.get("Linked Media File", "")

        html_content = """
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Relatório de Conversas</title>
<style>
body{font-family:Segoe UI,Tahoma,Verdana,sans-serif;background:#f6f7fb;color:#1f2937;margin:0;padding:20px 20px 20px 280px;min-width:900px}
.container{max-width:1200px;margin:0 auto}
h1{margin:0 0 12px 0}
h2{border-bottom:1px solid #e5e7eb;padding-bottom:8px;margin-top:22px}
.section{background:#fff;border:1px solid #e5e7eb;border-radius:10px;padding:16px;margin:12px 0;box-shadow:0 2px 10px rgba(0,0,0,.04)}
.sidebar{position:fixed;left:0;top:0;width:260px;height:100%;background:#0b1220;color:#e5e7eb;padding:18px;box-sizing:border-box}
.brand{font-size:18px;font-weight:700;margin-bottom:14px}
.small{font-size:12px;opacity:.85}
.sidebar a{color:#9dd7ff;text-decoration:none;display:block;padding:8px 10px;border-radius:8px}
.sidebar a:hover{background:rgba(157,215,255,.12)}
.btn{background:#2563eb;color:#fff;border:none;padding:10px 12px;border-radius:8px;cursor:pointer}
.btn:hover{background:#1d4ed8}
.footer{margin-top:30px;text-align:center;color:#6b7280;font-size:12px}
table{width:100%;border-collapse:collapse}
th,td{border:1px solid #e5e7eb;padding:8px}
th{background:#111827;color:#fff}
</style>
<script>
function toggleDetails(id){
  var e=document.getElementById(id);
  e.style.display = (e.style.display==="none") ? "block" : "none";
}
</script>
</head>
<body>
<div class="sidebar">
  <div class="brand">Bilhetagem</div>
  <div class="small">Relatórios por alvo / case</div>
  <div style="height:12px"></div>
  <a href="#conversas_individuais">Conversas Individuais</a>
  <a href="#conversas_grupo">Conversas em Grupo</a>
  <a href="#logs_chamada">Logs de Chamada</a>
  <a href="#foto_perfil">Foto do Perfil</a>
  <a href="#info_grupos">Informações dos Grupos</a>
  <a href="#contatos">Contatos</a>
</div>
<div class="container">
"""
        html_content += f"<h1>Relatório de Conversas — Alvo: {target_number}</h1>\n"
        html_content += f"<div class='section'><b>Case:</b> {case_number} &nbsp; | &nbsp; <b>Data:</b> {selected_date}</div>\n"

        
        # monta um índice rápido de meta do grupo (ID -> dict), pra puxar nome/subject no HTML
        group_meta_by_id = {}
        try:
            for m in (group_media_info or []):
                if isinstance(m, dict) and m.get("ID"):
                    group_meta_by_id[str(m.get("ID"))] = m
        except Exception:
            group_meta_by_id = {}

        html_content += '<h2 id="conversas_individuais">Conversas Individuais</h2>\n'
        if conversations:
            for number, msgs in conversations.items():
                div_id = f"conv_{number}"
                total_conv = total_interactions.get(number, 0)
                html_content += f"""
<div class='section'>
  <h3 style="margin:0 0 6px 0">Conversas com: {number}</h3>
  <div><b>Total:</b> {total_conv}</div>
  <div style="height:10px"></div>
  <button class="btn" onclick="toggleDetails('{div_id}')">Ver detalhes</button>
  <div id="{div_id}" style="display:none;margin-top:12px">
"""
                for msg in msgs:
                    sender = msg.get('Sender', 'Não informado')
                    recips = msg.get('Recipients', [])
                    # pode vir 1..N recipients; mostra o primeiro só pra não poluir
                    recipient = recips[0] if recips else "Não informado"
                    msg_type = msg.get('Type', 'Não informado')
                    msg_style = msg.get('Message Style', 'Não informado')
                    timestamp = msg.get('Timestamp', 'Não informado')
                    html_content += (
                        f"<div style='padding:6px 0;border-bottom:1px dashed #e5e7eb'>"
                        f"<b>{timestamp}</b> — {sender} → {recipient} | {msg_type} | {msg_style}"
                        f"</div>\n"
                    )
                html_content += "</div></div>\n"
        else:
            html_content += "<div class='section'>Não foram encontradas conversas individuais.</div>\n"

        html_content += '<h2 id="conversas_grupo">Conversas em Grupo</h2>\n'
        if group_conversations:
            for group_id, msgs in group_conversations.items():
                gid = str(group_id)
                div_id = f"group_{gid}"
                total_grp = total_group_interactions.get(group_id, 0)
                meta = group_meta_by_id.get(gid, {})
                subject = meta.get("Subject", "Não informado")
                # participantes que apareceram como recipients nos logs
                participants = sorted(list(current_group_recipients.get(group_id, set())))
                participants_txt = ", ".join(participants) if participants else "Não identificado no TXT"

                # mudanças (novos/removidos) comparando com execução anterior
                changes = group_changes.get(group_id, {})
                new_nums = ", ".join(sorted(changes.get("new_numbers", []))) if changes.get("new_numbers") else "Nenhum"
                removed_nums = ", ".join(sorted(changes.get("removed_numbers", []))) if changes.get("removed_numbers") else "Nenhum"

                html_content += f"""
<div class='section'>
  <h3 style="margin:0 0 6px 0">Grupo: {subject}</h3>
  <div><b>ID:</b> {gid}</div>
  <div><b>Total de mensagens:</b> {total_grp}</div>
  <div><b>Participantes (detectados):</b> {participants_txt}</div>
  <div style="margin-top:6px"><b>Novos no grupo:</b> {new_nums}</div>
  <div><b>Removidos do grupo:</b> {removed_nums}</div>
  <div style="height:10px"></div>
  <button class="btn" onclick="toggleDetails('{div_id}')">Ver detalhes</button>
  <div id="{div_id}" style="display:none;margin-top:12px">
"""
                for msg in msgs:
                    sender = msg.get('Sender', 'Não informado')
                    msg_type = msg.get('Type', 'Não informado')
                    msg_style = msg.get('Message Style', 'Não informado')
                    timestamp = msg.get('Timestamp', 'Não informado')
                    html_content += (
                        f"<div style='padding:6px 0;border-bottom:1px dashed #e5e7eb'>"
                        f"<b>{timestamp}</b> — {sender} | {msg_type} | {msg_style}"
                        f"</div>\n"
                    )
                html_content += "</div></div>\n"
        else:
            html_content += "<div class='section'>Não foram encontradas conversas em grupo.</div>\n"

        html_content += '<h2 id="logs_chamada">Logs de Chamada</h2>\n'
        if call_logs:
            html_content += """
<div class='section'>
  <button class="btn" onclick="toggleDetails('call_logs_all')">Ver detalhes</button>
  <div id="call_logs_all" style="display:none;margin-top:12px">
"""
            for log_entry in call_logs:
                call_id = log_entry.get('call_id', 'Não informado')
                html_content += f"<div style='padding:10px;border:1px solid #e5e7eb;border-radius:8px;margin:10px 0'>"
                html_content += f"<div><b>Call ID:</b> {call_id}</div>"
                for event in log_entry.get("events", []):
                    ev_type = event.get("Type", "Não informado")
                    ev_desc = event_map.get(ev_type, ev_type.capitalize())
                    ts = event.get("Timestamp", "Não informado")
                    from_f = event.get("From", "Não informado")
                    to_f = event.get("To", "Não informado")
                    media_t = event.get("Media Type", "Não informado")
                    html_content += f"<div style='margin-top:6px;color:#374151'>• <b>{ts}</b> — {from_f} → {to_f} | {media_t} | {ev_desc}</div>"
                html_content += "</div>"
            html_content += "</div></div>"
        else:
            html_content += "<div class='section'>Não foram encontrados logs de chamada.</div>"

        html_content += '<h2 id="foto_perfil">Foto do Perfil do WhatsApp</h2>\n'
        html_content += "<div class='section'>"
        html_content += f"<div><b>Nome do Perfil:</b> {push_name}</div>\n"
        if linked_media:
            html_content += f"<div style='margin-top:12px'><img src='{linked_media}' style='max-width:220px;border-radius:10px;border:1px solid #e5e7eb'></div>\n"
        else:
            html_content += "<div style='margin-top:10px;color:#6b7280'><i>Imagem não disponível</i></div>\n"
        html_content += "</div>\n"

        html_content += '<h2 id="info_grupos">Informações dos Grupos do WhatsApp</h2>\n'
        if group_media_info:
            for media in group_media_info:
                description = media.get("Description", "Não informado")
                media_file = media.get("Linked Media File", "")
                group_id = media.get("ID", "Não informado")
                creation = media.get("Creation", "Não informado")
                size = media.get("Size", "Não informado")
                subject = media.get("Subject", "Não informado")

                html_content += "<div class='section'>"
                html_content += f"<div><b>ID:</b> {group_id}</div>"
                html_content += f"<div><b>Criação:</b> {creation}</div>"
                html_content += f"<div><b>Tamanho:</b> {size}</div>"
                html_content += f"<div><b>Nome:</b> {subject}</div>"
                html_content += f"<div style='margin-top:8px'><b>Descrição:</b><br>{description}</div>"
                if media_file and media_file != "Não informado":
                    html_content += f"<div style='margin-top:12px'><img src='{media_file}' style='max-width:220px;border-radius:10px;border:1px solid #e5e7eb'></div>"
                else:
                    html_content += "<div style='margin-top:10px;color:#6b7280'><i>Imagem não disponível</i></div>"
                html_content += "</div>\n"
        else:
            html_content += "<div class='section'>Não há informações de grupos do WhatsApp.</div>"

        html_content += '<h2 id="contatos">Contatos</h2>\n'
        html_content += "<div class='section'>"
        html_content += "<div style='display:flex;gap:14px;flex-wrap:wrap'>"

        html_content += "<div style='flex:1;min-width:320px'>"
        html_content += "<h3 style='margin:0 0 10px 0'>Simétricos</h3>"
        html_content += "<table><tr><th>Contato</th></tr>"
        for contact in symmetric_contacts:
            html_content += f"<tr><td>{contact}</td></tr>\n"
        html_content += "</table>"
        html_content += f"<div style='margin-top:8px;color:#6b7280'>Total: {len(symmetric_contacts)}</div>"
        html_content += "</div>"

        html_content += "<div style='flex:1;min-width:320px'>"
        html_content += "<h3 style='margin:0 0 10px 0'>Assimétricos</h3>"
        html_content += "<table><tr><th>Contato</th></tr>"
        for contact in asymmetric_contacts:
            html_content += f"<tr><td>{contact}</td></tr>\n"
        html_content += "</table>"
        html_content += f"<div style='margin-top:8px;color:#6b7280'>Total: {len(asymmetric_contacts)}</div>"
        html_content += "</div>"

        html_content += "</div></div>"
        html_content += """
<div class='footer'>Relatório gerado automaticamente.</div>
</div>
</body></html>
"""

        output_file = os.path.join(target_dir, f"relatorio_conversas_case_{case_number}_{selected_date}.html")
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(html_content)

        log_file(output_file, "Relatório HTML do alvo")

        log_ok(
            f"[BILHETAGEM] Finalizado alvo {target_number}: "
            f"{len(conversations)} conversas individuais, {len(group_conversations)} grupos, "
            f"{len(symmetric_contacts)} simétricos, {len(asymmetric_contacts)} assimétricos."
        )

    except Exception as e:
        # Mesmo com erro durante a montagem do relatório, tenta salvar um HTML "de falha"
        # pra não quebrar a pipeline e você conseguir inspecionar o que aconteceu.
        log_error(f"[BILHETAGEM] {str(e)}")

        try:
            output_file_local = locals().get("output_file")
            html_content_local = locals().get("html_content", "")

            if output_file_local:
                # Se já tinha algum HTML montado, injeta um banner de erro.
                if html_content_local and "</body></html>" in html_content_local:
                    safe_msg = html.escape(str(e))
                    html_content_local = html_content_local.replace(
                        "</body></html>",
                        f"""<div class='section' style='border:1px solid #fecaca;background:#fff1f2'>
  <h3 style='margin:0 0 8px 0'>Erro ao gerar parte do relatório</h3>
  <div style='color:#7f1d1d'><b>Detalhes:</b> {safe_msg}</div>
</div>
</body></html>"""
                    )
                    content_to_write = html_content_local
                else:
                    safe_msg = html.escape(str(e))
                    content_to_write = f"""<!doctype html>
<html lang="pt-br"><head><meta charset="utf-8">
<title>Relatório com erro</title>
<style>
body{{font-family:Arial,Helvetica,sans-serif;background:#f8fafc;margin:0}}
.wrap{{max-width:1100px;margin:18px auto;padding:18px}}
.card{{background:#fff;border:1px solid #e5e7eb;border-radius:12px;padding:16px}}
.err{{border:1px solid #fecaca;background:#fff1f2;border-radius:10px;padding:12px}}
pre{{white-space:pre-wrap}}
</style></head><body>
<div class="wrap">
  <div class="card">
    <h2 style="margin:0 0 10px 0">Relatório não pôde ser gerado completamente</h2>
    <div class="err"><b>Erro:</b> {safe_msg}</div>
    <p style="color:#6b7280;margin-top:10px">O JSON e os arquivos gerados até o momento podem ter sido salvos normalmente.</p>
  </div>
</div>
</body></html>"""

                with open(output_file_local, "w", encoding="utf-8") as f:
                    f.write(content_to_write)

                log_file(output_file_local, "Relatório HTML do alvo (com erro)")
        except Exception as e2:
            log_error(f"[BILHETAGEM] falha ao gravar HTML de erro: {str(e2)}")


# ============================================================
# 4) PIPELINE COMPLETO POR ALVO
# ============================================================

def ensure_records_folder_for_target(records_root: str, target_number: str):
    alvo_dir = os.path.join(records_root, target_number)

    if os.path.isdir(alvo_dir):
        log_info(f"[ALVO {target_number}] Pasta de records encontrada: {alvo_dir}")
        return alvo_dir

    log_info(f"[ALVO {target_number}] Pasta de records não encontrada. Procurando ZIP em {records_root}...")
    zip_found = None

    if os.path.isdir(records_root):
        for fname in os.listdir(records_root):
            if fname.lower().endswith(".zip") and target_number in fname:
                zip_found = os.path.join(records_root, fname)
                break

    if not zip_found:
        log_warn(f"[ALVO {target_number}] Nenhum ZIP encontrado para este alvo.")
        return None

    try:
        os.makedirs(alvo_dir, exist_ok=True)
        with zipfile.ZipFile(zip_found, 'r') as zf:
            zf.extractall(alvo_dir)
        log_ok(f"[ALVO {target_number}] ZIP extraído para: {alvo_dir}")
        return alvo_dir
    except Exception as e:
        log_error(f"[ALVO {target_number}] Erro ao extrair ZIP {zip_found}: {e}")
        return None


def process_target(
    target_number: str,
    legal_files: list,
    records_root: str,
    alvos_root: str,
    selected_date: datetime,
):
    log_header(f"[ALVO {target_number}] Processamento do alvo")

    alvo_out_dir = os.path.join(alvos_root, target_number)
    os.makedirs(alvo_out_dir, exist_ok=True)

    records_marker = os.path.join(alvo_out_dir, ".records_merged")
    use_records = not os.path.exists(records_marker)

    alvo_records_dir = None
    html_path = None

    if use_records:
        if (not records_root) or (not os.path.isdir(records_root)):
            log_warn(
                f"[ALVO {target_number}] Pasta de records não informada/ inválida. "
                f"Rodando sem anexar records (se quiser anexar, selecione a pasta dos records)."
            )
            use_records = False
        else:
            alvo_records_dir = ensure_records_folder_for_target(records_root, target_number)
            if not alvo_records_dir:
                log_warn(f"[ALVO {target_number}] Não foi possível preparar pasta de records. Rodando sem anexar records.")
                use_records = False
            else:
                html_path = os.path.join(alvo_records_dir, "records.html")
                if not os.path.isfile(html_path):
                    html_path = None
                    for f2 in os.listdir(alvo_records_dir):
                        if f2.lower().endswith(".html"):
                            html_path = os.path.join(alvo_records_dir, f2)
                            break

                if (not html_path) or (not os.path.isfile(html_path)):
                    log_warn(f"[ALVO {target_number}] Nenhum records HTML encontrado. Rodando sem anexar records.")
                    use_records = False

    log_step("3/4 — Merge TXT legal + records (se necessário)")
    if use_records:
        log_info("records.html será anexado (primeira vez deste alvo).")
        records_block = extract_records_block(html_path)
    else:
        if os.path.exists(records_marker):
            log_info("records.html já foi anexado antes (marcador encontrado).")
        else:
            log_info("records.html não foi anexado (records não informado).")
        records_block = ""

    legal_files_sorted = sorted(legal_files)
    if not legal_files_sorted:
        log_warn(f"[ALVO {target_number}] Nenhum TXT legal encontrado. Pulando alvo.")
        return

    main_legal_path = legal_files_sorted[0]

    parts = []
    for p in legal_files_sorted:
        try:
            with open(p, "r", encoding="utf-8", errors="ignore") as fin:
                parts.append(fin.read().rstrip())
        except Exception as e:
            log_warn(f"[ALVO {target_number}] Erro lendo {p}: {e}")

    if use_records and records_block.strip():
        parts.append(records_block.rstrip())

    merged_text = "\n".join([p for p in parts if p.strip()])

    with open(main_legal_path, "w", encoding="utf-8", errors="ignore") as f:
        f.write(merged_text + "\n")

    log_ok(f"[ALVO {target_number}] TXT principal atualizado.")
    log_file(main_legal_path, "TXT principal atualizado (merge)")

    if use_records and records_block.strip():
        try:
            with open(records_marker, "w", encoding="utf-8") as mf:
                mf.write(f"records anexados em {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
            log_ok(f"[ALVO {target_number}] Marcador criado: {records_marker}")
        except Exception as e:
            log_warn(f"[ALVO {target_number}] Erro ao criar marcador: {e}")

    dst_media = os.path.join(alvo_out_dir, "linked_media")

    if alvo_records_dir:
        src_media = os.path.join(alvo_records_dir, "linked_media")
    else:
        src_media = None

    if src_media and os.path.isdir(src_media):
        if os.path.exists(dst_media):
            log_info(f"[ALVO {target_number}] linked_media já existe no destino (não será movida).")
        else:
            try:
                shutil.move(src_media, dst_media)
                log_ok(f"[ALVO {target_number}] linked_media movida para: {dst_media}")
            except Exception as e:
                log_warn(f"[ALVO {target_number}] Erro ao mover linked_media: {e}")
    else:
        log_info(f"[ALVO {target_number}] Nenhuma pasta linked_media encontrada no records.")

    log_step("4/4 — Bilhetagem + geração de relatórios")
    selected_date_str = selected_date.strftime("%d-%m-%Y")
    process_file(main_legal_path, alvos_root, selected_date_str)

    log_ok(f"[ALVO {target_number}] Finalizado.")


def process_case_full(sender_email: str, case_number: str, selected_date: datetime, base_dir: str, records_root: str):
    base_dir = os.path.abspath(base_dir)
    records_root = os.path.abspath(records_root) if records_root else ""

    emails_dir = os.path.join(base_dir, "EMAILS")
    alvos_root = os.path.join(base_dir, "ALVOS")

    os.makedirs(base_dir, exist_ok=True)
    os.makedirs(emails_dir, exist_ok=True)
    os.makedirs(alvos_root, exist_ok=True)

    log_header("Bilhetagem — Fluxo completo")
    log_info(f"BASE_DIR     = {base_dir}")
    log_info(f"EMAILS_DIR   = {emails_dir}")
    log_info(f"ALVOS_ROOT   = {alvos_root}")
    log_info(f"RECORDS_ROOT = {records_root}")

    if records_root and (not os.path.isdir(records_root)):
        log_warn(f"Pasta RECORDS_ROOT não encontrada: {records_root} (rodando sem records)")
        records_root = ""

    targets_files = fetch_emails_for_case(
        sender_email=sender_email,
        case_number=case_number,
        selected_date=selected_date,
        emails_dir=emails_dir,
    )

    if not targets_files:
        log_warn("Nenhum e-mail encontrado. Verifique Case, data e remetente.")
        return

    log_step("3/4 — Processando alvos (merge + bilhetagem)")
    for target, files in targets_files.items():
        process_target(
            target_number=target,
            legal_files=files,
            records_root=records_root,
            alvos_root=alvos_root,
            selected_date=selected_date,
        )

    # ======================================================
    # VÍNCULOS (AUTOMÁTICO): 1 relatório FIXO na raiz de ALVOS
    # ======================================================
    try:
        # NOME FIXO (sem data → sobrescreve sempre)
        vinculos_path = os.path.join(
            alvos_root,
            f"relatorio_vinculos_case_{case_number}.html"
        )

        log_step("4/4 — Gerando relatório de vínculos (ALVOS)")

        ok = generate_vinculos_report_for_alvos(
            alvos_root=alvos_root,
            output_html=vinculos_path,
            # data só no TÍTULO, não no nome do arquivo
            title_suffix=f"Case {case_number} — atualizado em {datetime.now().strftime('%d/%m/%Y %H:%M')}",
        )

        if ok:
            log_ok("[VÍNCULOS] Relatório atualizado com sucesso (substituído).")
            log_file(vinculos_path, "Relatório de vínculos (ALVOS)")
        else:
            log_warn("[VÍNCULOS] Relatório não foi gerado (sem correlações suficientes).")

    except Exception as e:
        log_warn(f"[VÍNCULOS] Falha ao gerar relatório: {e}")

    log_ok("Fluxo concluído para todos os alvos.")


# ============================================================
# 5) GUI
# ============================================================

APP_TITLE = "DataFusion Analyzer — Processing Pipeline"
FOOTER_TEXT = "© 2025 DataFusion Analyzer — desenvolvido por Braian Rodrigues"


def _safe_set_locale_ptbr():
    for loc in ("pt_BR.UTF-8", "Portuguese_Brazil.1252", "pt_BR"):
        try:
            locale.setlocale(locale.LC_TIME, loc)
            return
        except Exception:
            continue


def create_gui():
    global log_text

    _safe_set_locale_ptbr()

    root = tk.Tk()
    root.title(APP_TITLE)
    try:
        root.iconbitmap("bilhetagem.ico")
    except Exception:
        pass
    root.geometry("1360x760")
    root.minsize(1200, 680)

    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except Exception:
        pass

    BLUE = "#1f5bd7"
    BLUE_DARK = "#1648b0"
    BLUE_SOFT = "#eaf2ff"
    BG = "#ffffff"
    BG_APP = "#f5f9ff"
    BORDER = "#cfe0ff"
    TXT = "#0f172a"
    MUTED = "#33507a"

    style.configure(".", font=("Segoe UI", 10))
    style.configure("App.TFrame", background=BG_APP)
    style.configure("Bar.TFrame", background=BLUE)
    style.configure("Card.TFrame", background=BG, relief="solid", borderwidth=1)
    style.configure("CardTitle.TLabel", background=BG, foreground=TXT, font=("Segoe UI", 11, "bold"))
    style.configure("CardText.TLabel", background=BG, foreground=TXT)
    style.configure("Muted.TLabel", background=BG_APP, foreground=MUTED)
    style.configure("Title.TLabel", background=BLUE, foreground="white", font=("Segoe UI", 14, "bold"))
    style.configure("Subtitle.TLabel", background=BLUE, foreground=BLUE_SOFT, font=("Segoe UI", 9))
    style.configure("TEntry", fieldbackground="white", foreground=TXT, padding=6)

    style.configure("Primary.TButton", background=BLUE, foreground="white", padding=(12, 10), font=("Segoe UI", 10, "bold"))
    style.map("Primary.TButton", background=[("active", BLUE_DARK)])

    style.configure("Soft.TButton", background=BLUE_SOFT, foreground=TXT, padding=(12, 10))
    style.map("Soft.TButton", background=[("active", "#dbe9ff")])

    root.configure(bg=BG_APP)
    root.rowconfigure(1, weight=1)
    root.columnconfigure(0, weight=1)

    topbar = ttk.Frame(root, style="Bar.TFrame")
    topbar.grid(row=0, column=0, sticky="ew")
    topbar.columnconfigure(0, weight=1)

    ttk.Label(topbar, text="DataFusion Analyzer — Processing Pipeline", style="Title.TLabel").grid(row=0, column=0, sticky="w", padx=14, pady=(10, 0))
    ttk.Label(topbar, text="Outlook → TXT → Merge (records.html) → Bilhetagem → HTML/JSON + Vínculos", style="Subtitle.TLabel").grid(row=1, column=0, sticky="w", padx=14, pady=(0, 10))

    body = ttk.Frame(root, style="App.TFrame", padding=12)
    body.grid(row=1, column=0, sticky="nsew")
    body.columnconfigure(0, weight=3)
    body.columnconfigure(1, weight=2)
    body.rowconfigure(0, weight=1)

    left = ttk.Frame(body, style="App.TFrame")
    left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
    left.columnconfigure(0, weight=1)
    left.rowconfigure(3, weight=1)

    card_cfg = ttk.Frame(left, style="Card.TFrame", padding=14)
    card_cfg.grid(row=0, column=0, sticky="ew")
    card_cfg.columnconfigure(1, weight=1)

    ttk.Label(card_cfg, text="Passo 1 — Configurações", style="CardTitle.TLabel").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))

    ttk.Label(card_cfg, text="Remetente (Outlook)", style="CardText.TLabel").grid(row=1, column=0, sticky="w")
    entry_sender = ttk.Entry(card_cfg)
    entry_sender.grid(row=1, column=1, sticky="ew", padx=(10, 0), columnspan=2)
    entry_sender.insert(0, "exemplo@email.com")

    ttk.Label(card_cfg, text="Case", style="CardText.TLabel").grid(row=2, column=0, sticky="w", pady=(10, 0))
    entry_case = ttk.Entry(card_cfg, width=28)
    entry_case.grid(row=2, column=1, sticky="w", padx=(10, 0), pady=(10, 0))

    ttk.Label(card_cfg, text="Data dos e-mails", style="CardText.TLabel").grid(row=3, column=0, sticky="w", pady=(10, 0))
    if DateEntry is None:
        date_var = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y"))
        date_entry = ttk.Entry(card_cfg, textvariable=date_var, width=18)
        date_entry.grid(row=3, column=1, sticky="w", padx=(10, 0), pady=(10, 0))
        ttk.Label(card_cfg, text="(instale tkcalendar p/ calendário)", style="CardText.TLabel").grid(row=3, column=2, sticky="w", padx=10, pady=(10, 0))
    else:
        date_entry = DateEntry(card_cfg, width=18, date_pattern="dd/MM/yyyy", locale="pt_BR", year=datetime.now().year)
        date_entry.grid(row=3, column=1, sticky="w", padx=(10, 0), pady=(10, 0))

    ttk.Label(card_cfg, text="Selecione a pasta onde deseja salvar", style="CardText.TLabel").grid(row=4, column=0, sticky="w", pady=(10, 0))
    base_dir_var = tk.StringVar()
    entry_base_dir = ttk.Entry(card_cfg, textvariable=base_dir_var)
    entry_base_dir.grid(row=4, column=1, sticky="ew", padx=(10, 0), pady=(10, 0))

    def browse_base_dir():
        folder = filedialog.askdirectory(title="Selecione a pasta base do caso")
        if folder:
            base_dir_var.set(folder)

    ttk.Button(card_cfg, text="Selecionar pasta", style="Soft.TButton", command=browse_base_dir).grid(row=4, column=2, sticky="ew", padx=10, pady=(10, 0))

    ttk.Label(card_cfg, text="Pasta raiz dos RECORDS (opcional)", style="CardText.TLabel").grid(row=5, column=0, sticky="w", pady=(10, 0))
    records_root_var = tk.StringVar()
    entry_records_root = ttk.Entry(card_cfg, textvariable=records_root_var)
    entry_records_root.grid(row=5, column=1, sticky="ew", padx=(10, 0), pady=(10, 0))

    def browse_records_root():
        folder = filedialog.askdirectory(title="Selecione a pasta raiz dos records (ZIPs ou subpastas por alvo)")
        if folder:
            records_root_var.set(folder)

    ttk.Button(card_cfg, text="Selecionar records", style="Soft.TButton", command=browse_records_root).grid(row=5, column=2, sticky="ew", padx=10, pady=(10, 0))

    card_tips = ttk.Frame(left, style="Card.TFrame", padding=14)
    card_tips.grid(row=1, column=0, sticky="ew", pady=(10, 0))
    ttk.Label(card_tips, text="Dicas rápidas", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 8))
    tips = (
        "• Case deve ser apenas o número (ex.: 20726034)\n"
        "• Records pode conter ZIPs ou pastas por alvo (use na 1ª execução de cada alvo)\n"
        "• Ao final, será gerado um relatório de vínculos dentro de ALVOS"
    )
    ttk.Label(card_tips, text=tips, style="CardText.TLabel", justify="left").grid(row=1, column=0, sticky="w")

    card_run = ttk.Frame(left, style="Card.TFrame", padding=14)
    card_run.grid(row=2, column=0, sticky="ew", pady=(10, 0))
    card_run.columnconfigure(0, weight=1)
    card_run.columnconfigure(1, weight=1)
    card_run.columnconfigure(2, weight=1)

    ttk.Label(card_run, text="Passo 2 — Executar", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 10), columnspan=3)

    progress = ttk.Progressbar(card_run, mode="indeterminate")
    progress.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(0, 12))

    buttons_grid = ttk.Frame(card_run, style="Card.TFrame")
    buttons_grid.grid(row=2, column=0, columnspan=3, sticky="ew")
    for c in range(2):
        buttons_grid.columnconfigure(c, weight=1)

    btn_run = ttk.Button(buttons_grid, text="Executar fluxo completo", style="Primary.TButton")
    btn_base = ttk.Button(buttons_grid, text="Abrir PASTA", style="Soft.TButton")
    btn_emails = ttk.Button(buttons_grid, text="Abrir EMAILS", style="Soft.TButton")
    btn_alvos = ttk.Button(buttons_grid, text="Abrir ALVOS", style="Soft.TButton")

    btn_run.grid(row=0, column=0, sticky="ew", padx=(0, 10), pady=(0, 10))
    btn_base.grid(row=0, column=1, sticky="ew", padx=(10, 0), pady=(0, 10))
    btn_emails.grid(row=1, column=0, sticky="ew", padx=(0, 10))
    btn_alvos.grid(row=1, column=1, sticky="ew", padx=(10, 0))

    card_log = ttk.Frame(left, style="Card.TFrame", padding=14)
    card_log.grid(row=3, column=0, sticky="nsew", pady=(10, 0))
    card_log.columnconfigure(0, weight=1)
    card_log.rowconfigure(1, weight=1)

    ttk.Label(card_log, text="Console / Logs", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 10))

    log_text_local = ScrolledText(card_log, wrap=tk.WORD, height=14, state=tk.DISABLED)
    log_text_local.grid(row=1, column=0, sticky="nsew")
    log_text_local.configure(background="white", relief="solid", borderwidth=1)

    # ===========================
    # Console/Logs — formatação
    # ===========================
    try:
        log_text_local.configure(font=("Consolas", 10))
    except Exception:
        pass

    # tags por nível/prefixo
    try:
        log_text_local.tag_configure("DBG", foreground="#6b7280")   # cinza
        log_text_local.tag_configure("OK", foreground="#16a34a")    # verde
        log_text_local.tag_configure("RUN", foreground="#2563eb")   # azul
        log_text_local.tag_configure("WARN", foreground="#d97706")  # âmbar
        log_text_local.tag_configure("ERR", foreground="#dc2626")   # vermelho
        log_text_local.tag_configure("DICA", foreground="#7c3aed")  # roxo
        log_text_local.tag_configure("TS", foreground="#9ca3af")    # timestamp
    except Exception:
        pass

    auto_scroll_var = tk.BooleanVar(value=True)
    show_debug_var = tk.BooleanVar(value=True)

    log_text = log_text_local

    def clear_log():
        log_text_local.config(state=tk.NORMAL)
        log_text_local.delete("1.0", tk.END)
        log_text_local.config(state=tk.DISABLED)

    # Barra de ações do console
    frame_log_actions = ttk.Frame(card_log, style="App.TFrame")
    frame_log_actions.grid(row=2, column=0, sticky="ew", pady=(10, 0))
    frame_log_actions.columnconfigure(0, weight=1)

    btn_clear_log = ttk.Button(frame_log_actions, text="Limpar log", style="Soft.TButton", command=clear_log)
    btn_clear_log.grid(row=0, column=2, sticky="e")

    try:
        def on_copy_log():
            try:
                root.clipboard_clear()
                txt = log_text_local.get("1.0", tk.END).rstrip()
                root.clipboard_append(txt)
            except Exception:
                pass

        btn_copy_log = ttk.Button(frame_log_actions, text="Copiar log", style="Soft.TButton", command=on_copy_log)
        btn_copy_log.grid(row=0, column=3, sticky="e", padx=(12, 0))
    except Exception:
        pass

    right = ttk.Frame(body, style="App.TFrame")
    right.grid(row=0, column=1, sticky="nsew")
    right.columnconfigure(0, weight=1)
    right.rowconfigure(0, weight=1)

    card_files = ttk.Frame(right, style="Card.TFrame", padding=14)
    card_files.grid(row=0, column=0, sticky="nsew")
    card_files.columnconfigure(0, weight=1)
    card_files.rowconfigure(1, weight=1)

    ttk.Label(card_files, text="Arquivos gerados (após execução)", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 10))

    # Lista de arquivos (mais legível)
    files_tree = ttk.Treeview(
        card_files,
        columns=("kind", "name", "folder", "path"),
        show="headings",
        height=18
    )
    files_tree.grid(row=1, column=0, sticky="nsew")

    files_tree.heading("kind", text="Tipo")
    files_tree.heading("name", text="Arquivo")
    files_tree.heading("folder", text="Pasta")
    files_tree.heading("path", text="")

    files_tree.column("kind", width=90, stretch=False, anchor="w")
    files_tree.column("name", width=420, stretch=True, anchor="w")
    files_tree.column("folder", width=260, stretch=True, anchor="w")
    # path fica oculto (usado só pra abrir)
    files_tree.column("path", width=0, stretch=False)

    sb_files = ttk.Scrollbar(card_files, orient="vertical", command=files_tree.yview)
    sb_files.grid(row=1, column=1, sticky="ns")
    files_tree.configure(yscrollcommand=sb_files.set)


    actions_files = ttk.Frame(card_files, style="Card.TFrame")
    actions_files.grid(row=2, column=0, sticky="ew", pady=(10, 0))
    actions_files.columnconfigure(0, weight=1)
    actions_files.columnconfigure(1, weight=1)
    actions_files.columnconfigure(2, weight=1)

    btn_open_file = ttk.Button(actions_files, text="Abrir arquivo", style="Soft.TButton")
    btn_open_folder_file = ttk.Button(actions_files, text="Abrir pasta", style="Soft.TButton")
    btn_clear_files = ttk.Button(actions_files, text="Limpar lista", style="Soft.TButton")

    btn_open_file.grid(row=0, column=0, sticky="ew", padx=(0, 8))
    btn_open_folder_file.grid(row=0, column=1, sticky="ew", padx=8)
    btn_clear_files.grid(row=0, column=2, sticky="ew", padx=(8, 0))

    # ======================================================
    # UI PUMP (logs + arquivos gerados)
    # - Mantém toda a lógica atual baseada em Queue, só
    #   drena a fila no thread da GUI com root.after()
    # ======================================================
    files_seen = set()

    def _ui_append_log(line: str):
        # Formato: HH:MM:SS  [TAG] mensagem
        try:
            raw = str(line).rstrip("\n")
            if not raw:
                return

            # filtra DEBUG se o usuário desmarcar
            if not show_debug_var.get() and raw.startswith("[DBG]"):
                return

            # detecta tag
            tag = None
            if raw.startswith("[DBG]"):
                tag = "DBG"
            elif raw.startswith("[OK]"):
                tag = "OK"
            elif raw.startswith("[RUN]"):
                tag = "RUN"
            elif raw.startswith("[WARN]") or raw.startswith("[WARNING]"):
                tag = "WARN"
            elif raw.startswith("[ERR]") or raw.startswith("[ERROR]"):
                tag = "ERR"
            elif raw.startswith("[DICA]"):
                tag = "DICA"

            from datetime import datetime
            ts = datetime.now().strftime("%H:%M:%S")

            log_text_local.config(state=tk.NORMAL)

            log_text_local.insert(tk.END, f"{ts}  ", ("TS",))
            if tag:
                log_text_local.insert(tk.END, raw + "\n", (tag,))
            else:
                log_text_local.insert(tk.END, raw + "\n")

            if auto_scroll_var.get():
                log_text_local.see(tk.END)

            log_text_local.config(state=tk.DISABLED)
        except Exception:
            pass

    def _ui_add_file(p: str):
        try:
            p = os.path.abspath(str(p))
            p_low = p.lower()

            # filtro: não exibir .json na lista (usuário não precisa ver)
            if p_low.endswith(".json"):
                return

            # exibir apenas:
            # - EMAILS (txt/sem extensão) dentro da pasta EMAILS
            # - relatórios HTML
            up = p.upper()
            ext = os.path.splitext(p_low)[1]

            is_email = ("\\EMAILS\\" in up) or ("/EMAILS/" in up)
            is_html = ext in (".html", ".htm")

            if not (is_email or is_html):
                return

            if p in files_seen:
                return
            files_seen.add(p)

            kind = "EMAIL" if is_email else "HTML"

            name = os.path.basename(p)
            folder = os.path.dirname(p)

            # insere (tipo | arquivo | pasta | path oculto)
            files_tree.insert("", "end", values=(kind, name, folder, p))
        except Exception:
            pass
        except Exception:
            pass

    def _pump_ui():
        # logs
        try:
            while True:
                line = _log_queue.get_nowait()
                _ui_append_log(str(line))
        except queue.Empty:
            pass
        except Exception:
            pass

        # arquivos
        try:
            while True:
                p = _files_queue.get_nowait()
                _ui_add_file(str(p))
        except queue.Empty:
            pass
        except Exception:
            pass

        # agenda próximo ciclo
        try:
            root.after(120, _pump_ui)
        except Exception:
            pass


    def get_selected_file():
        try:
            sel = files_tree.selection()
            if not sel:
                return None
            values = files_tree.item(sel[0]).get("values") or []
            if len(values) < 4:
                return None
            return values[3]
        except Exception:
            return None

    def open_folder(path: str):
        try:
            path = os.path.abspath(path)
            if os.path.isdir(path):
                os.startfile(path)
            else:
                messagebox.showwarning("Aviso", f"Pasta não encontrada:\n{path}")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao abrir pasta:\n{e}")

    def on_open_selected_file():
        p = get_selected_file()
        if not p:
            messagebox.showwarning("Aviso", "Selecione um arquivo na lista.")
            return
        try:
            if os.path.exists(p):
                os.startfile(p)
            else:
                messagebox.showwarning("Aviso", f"Arquivo não encontrado no disco:\n{p}")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao abrir arquivo:\n{e}")

    def on_open_selected_folder():
        p = get_selected_file()
        if not p:
            messagebox.showwarning("Aviso", "Selecione um arquivo na lista.")
            return
        folder = os.path.dirname(p)
        open_folder(folder)

    def on_clear_files():
        files_tree.delete(*files_tree.get_children())

    btn_open_file.config(command=on_open_selected_file)
    btn_open_folder_file.config(command=on_open_selected_folder)
    btn_clear_files.config(command=on_clear_files)

    footer = ttk.Frame(root, style="App.TFrame", padding=(12, 6))
    footer.grid(row=2, column=0, sticky="ew")
    footer.columnconfigure(0, weight=1)
    ttk.Label(footer, text=FOOTER_TEXT, style="Muted.TLabel").grid(row=0, column=0, sticky="e")

    # ======================================================
    # IMPORTANTE: running_var precisa existir ANTES do on_run
    # ======================================================
    running_var = tk.BooleanVar(value=False)

    def set_running(is_running: bool):
        running_var.set(is_running)
        if is_running:
            progress.start(10)
        else:
            progress.stop()

        state_nav = tk.DISABLED if is_running else tk.NORMAL
        btn_base.config(state=state_nav)
        btn_emails.config(state=state_nav)
        btn_alvos.config(state=state_nav)

    def validate_inputs():
        sender = entry_sender.get().strip()
        case = entry_case.get().strip()
        base_dir = base_dir_var.get().strip()
        records_root = records_root_var.get().strip()

        if not sender or not case:
            raise ValueError("Informe remetente e número do Case.")
        if not base_dir:
            raise ValueError("Selecione a pasta base do caso.")

        if DateEntry is None:
            try:
                selected_date = datetime.strptime(date_var.get().strip(), "%d/%m/%Y")
            except Exception:
                raise ValueError("Data inválida (use dd/mm/aaaa).")
        else:
            try:
                selected_date = datetime.strptime(date_entry.get().strip(), "%d/%m/%Y")
            except Exception:
                raise ValueError("Data inválida.")

        return sender, case, selected_date, base_dir, records_root

    def on_open_base():
        p = base_dir_var.get().strip()
        if not p:
            messagebox.showwarning("Aviso", "Selecione a pasta base do caso primeiro.")
            return
        try:
            os.makedirs(p, exist_ok=True)
        except Exception:
            pass
        open_folder(p)

    def on_open_emails():
        p = base_dir_var.get().strip()
        if not p:
            messagebox.showwarning("Aviso", "Selecione a pasta base do caso primeiro.")
            return
        open_folder(os.path.join(os.path.abspath(p), "EMAILS"))

    def on_open_alvos():
        p = base_dir_var.get().strip()
        if not p:
            messagebox.showwarning("Aviso", "Selecione a pasta base do caso primeiro.")
            return
        open_folder(os.path.join(os.path.abspath(p), "ALVOS"))

    btn_base.config(command=on_open_base)
    btn_emails.config(command=on_open_emails)
    btn_alvos.config(command=on_open_alvos)

    def worker_run(sender, case, selected_date, base_dir, records_root):
        try:
            log_header("EXECUÇÃO INICIADA")
            log_info("Exportar e-mails → separar alvos → merge → bilhetagem → relatórios + vínculos")
            process_case_full(sender, case, selected_date, base_dir, records_root)
            log_header("EXECUÇÃO FINALIZADA")
            root.after(0, lambda: messagebox.showinfo("Concluído", "Processamento concluído para todos os alvos."))
        except Exception as e:
            log_error(str(e))
            root.after(0, lambda: messagebox.showerror("Erro", str(e)))
        finally:
            root.after(0, lambda: (set_running(False), btn_run.config(state=tk.NORMAL)))

    def on_run():
        if running_var.get():
            return

        try:
            sender, case, selected_date, base_dir, records_root = validate_inputs()
        except Exception as e:
            messagebox.showerror("Erro", str(e))
            return

        btn_run.config(state=tk.DISABLED)
        set_running(True)

        t = threading.Thread(
            target=worker_run,
            args=(sender, case, selected_date, base_dir, records_root),
            daemon=True
        )
        t.start()

    btn_run.config(command=on_run)
    # inicia o pump da UI (logs + arquivos)
    _pump_ui()

    root.mainloop()
if __name__ == "__main__":
    create_gui()

