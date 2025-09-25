import imaplib
import os
import email
from email.header import decode_header
import datetime
import shutil
import platform
import subprocess
import time
import re
import csv
import sys
from io import BytesIO
from dotenv import load_dotenv

# --- Библиотеки для генерации PDF и работы с файлами ---
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, Image
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from bs4 import BeautifulSoup

# Загружаем переменные окружения
load_dotenv()

# --- Конфигурация ---
IMAP_SERVER = os.getenv('IMAP_SERVER')
MAIL_RU_EMAIL = os.getenv('MAIL_RU_EMAIL')
MAIL_RU_PASSWORD = os.getenv('MAIL_RU_PASSWORD')

# --- Структура папок ---
# Главная папка для всех операций
BASE_OUTPUT_DIRECTORY = "email_processor"
# 1. Папка для временно скачанных PDF (для просмотра)
DOWNLOADED_PDF_DIR = os.path.join(BASE_OUTPUT_DIRECTORY, "1_downloaded_pdf")
# 2. Папка для писем и вложений в исходном формате
DOWNLOADED_ORIGINALS_DIR = os.path.join(BASE_OUTPUT_DIRECTORY, "2_downloaded_originals")
# 3. Финальная папка для зарегистрированных PDF
REGISTERED_DIR = os.path.join(BASE_OUTPUT_DIRECTORY, "3_registered_emails")
# Имя файла журнала регистрации
JOURNAL_CSV_FILE = os.path.join(BASE_OUTPUT_DIRECTORY, "registration_journal.csv")


# --- Шрифты и прочее ---
def resource_path(relative_path):
    """ Получить абсолютный путь к ресурсу, работает для разработки и для PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

DEJAVU_SANS_FONT_PATH = resource_path("DejaVuSans.ttf")

# --- Проверка зависимостей (взято из вашего кода) ---
try:
    from pypdf import PdfWriter
    PYPDF_AVAILABLE = True
    print("INFO: Библиотека pypdf найдена, объединение PDF будет доступно.")
except ImportError:
    PYPDF_AVAILABLE = False
    print("ПРЕДУПРЕЖДЕНИЕ: Библиотека pypdf не найдена. PDF-вложения не будут объединены.")

WIN32COM_AVAILABLE = False
if os.name == 'nt':
    try:
        import win32com.client
        WIN32COM_AVAILABLE = True
        CONVERTIBLE_EXTENSIONS = ['.doc','.docx','.rtf','.odt','.txt', '.xls','.xlsx','.ods', '.ppt','.pptx','.odp','.xlsm']
        print("INFO: Библиотека pywin32 найдена, конвертация файлов MS Office в PDF будет доступна.")
    except ImportError:
        print("ПРЕДУПРЕЖДЕНИЕ: Библиотека pywin32 не найдена. Конвертация файлов MS Office в PDF будет недоступна.")
else:
    print("INFO: Скрипт запущен не на Windows. Конвертация файлов MS Office в PDF будет недоступна.")

#
# --- БЛОК 1: СКАЧИВАНИЕ И ПЕРВИЧНАЯ ОБРАБОТКА ---
# В этом блоке находятся функции, отвечающие за загрузку писем с сервера.
#

def download_all_unseen_emails():
    """
    Основная функция этапа 1. Подключается к почте и скачивает все непрочитанные письма,
    создавая для каждого PDF-версию и сохраняя оригиналы.
    """
    if not all([IMAP_SERVER, MAIL_RU_EMAIL, MAIL_RU_PASSWORD]):
        print("CRITICAL: Переменные окружения для почты не найдены в .env файле. Завершение работы.")
        return []

    mail = None
    processed_emails_metadata = []
    try:
        print(f"Подключение к {IMAP_SERVER}...")
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, 993)
        mail.login(MAIL_RU_EMAIL, MAIL_RU_PASSWORD)
        mail.select('inbox')
        print("Успешно подключено к 'Входящие'.")

        status, data = mail.search(None, 'UNSEEN')
        if status != 'OK':
            print("ERROR: Ошибка поиска писем.")
            return []

        email_ids = data[0].split()
        num_unread = len(email_ids)
        print(f"Найдено {num_unread} непрочитанных писем. Начинаю скачивание...")

        if num_unread == 0:
            return []

        for i, email_id_bytes in enumerate(email_ids):
            email_id_str = email_id_bytes.decode()
            print(f"\n--- Скачиваю письмо {i + 1}/{num_unread} (ID: {email_id_str}) ---")
            
            # Получаем объект письма
            status, msg_data = mail.fetch(email_id_str, '(RFC822)')
            if status != 'OK':
                print(f"ERROR: Не удалось получить письмо с ID {email_id_str}.")
                continue
            
            msg = email.message_from_bytes(msg_data[0][1])
            email_headers = _extract_email_headers(msg)
            
            print(f"  От: {email_headers['sender']}")
            print(f"  Тема: {email_headers['subject']}")

            # Генерируем уникальное имя для этого письма, чтобы связать PDF и папку с оригиналами
            timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S_%f')
            unique_email_id = f"email_{timestamp}"

            # Создаем PDF и сохраняем оригиналы
            pdf_path, originals_path = _generate_pdf_and_save_originals(msg, unique_email_id, email_headers)

            if pdf_path:
                processed_emails_metadata.append({
                    "unique_id": unique_email_id,
                    "pdf_path": pdf_path,
                    "originals_path": originals_path,
                    "sender": email_headers['sender'],
                    "subject": email_headers['subject']
                })
        
        return processed_emails_metadata

    except Exception as e:
        print(f"CRITICAL: Произошла критическая ошибка при скачивании писем: {e}")
        traceback.print_exc()
        return []
    finally:
        if mail:
            mail.logout()
        print("\n--- Завершено скачивание писем ---")


def _generate_pdf_and_save_originals(msg, unique_email_id, headers):
    """
    Для одного письма: создает PDF и сохраняет оригинал письма (.eml) и вложения.
    Возвращает пути к созданному PDF и папке с оригиналами.
    """
    # --- Путь для PDF ---
    pdf_path = os.path.join(DOWNLOADED_PDF_DIR, f"{unique_email_id}.pdf")
    
    # --- Пути для оригиналов ---
    originals_folder_path = os.path.join(DOWNLOADED_ORIGINALS_DIR, unique_email_id)
    os.makedirs(originals_folder_path, exist_ok=True)
    
    # Сохраняем все письмо как .eml файл
    with open(os.path.join(originals_folder_path, "original_email.eml"), "wb") as f:
        f.write(msg.as_bytes())

    # --- Создание PDF (логика из вашего кода) ---
    # Этот блок почти полностью взят из вашего скрипта, т.к. он отлично работает
    
    # Вспомогательные переменные
    pdf_attachments_to_merge = []
    
    # 1. Создаем "тело" письма в PDF
    pdf_canvas, styles, _, page_dims, current_y = _setup_pdf_canvas_and_styles(pdf_path)
    body = _extract_email_body(msg)
    display_date_str = datetime.datetime.now().strftime("%d.%m.%Y")
    
    # Добавляем заголовки в PDF
    current_y = _add_paragraph_to_pdf_util(pdf_canvas, f"<b>Дата получения (факт):</b> {headers['formatted_date']}", styles['N'], current_y, page_dims)
    current_y = _add_paragraph_to_pdf_util(pdf_canvas, f"<b>От:</b> {headers['sender']}", styles['N'], current_y, page_dims)
    current_y = _add_paragraph_to_pdf_util(pdf_canvas, f"<b>Кому:</b> {headers['recipients']}", styles['N'], current_y, page_dims)
    current_y = _add_paragraph_to_pdf_util(pdf_canvas, f"<b>Тема:</b> {headers['subject']}", styles['N'], current_y, page_dims)
    current_y = _add_paragraph_to_pdf_util(pdf_canvas, "<b>Содержание:</b>", styles['N'], current_y, page_dims)
    _add_paragraph_to_pdf_util(pdf_canvas, body if body.strip() else "Содержимое отсутствует.", styles['Body'], current_y, page_dims)
    
    pdf_canvas.save() # Сохраняем основной PDF

    # 2. Обрабатываем вложения
    for part in msg.walk():
        if part.get_content_maintype() == 'multipart' or part.get('Content-Disposition') is None:
            continue

        filename = part.get_filename()
        if filename:
            # Декодируем имя файла
            decoded_fn = "".join([p.decode(c or 'utf-8', 'replace') if isinstance(p, bytes) else str(p) for p, c in decode_header(filename)])
            sanitized_fn = sanitize_filename(decoded_fn)
            filepath = os.path.join(originals_folder_path, sanitized_fn)
            
            # Сохраняем оригинальное вложение
            with open(filepath, 'wb') as f:
                f.write(part.get_payload(decode=True))
            print(f"  -> Сохранен оригинал вложения: {sanitized_fn}")

            # Пытаемся конвертировать в PDF для слияния
            file_ext = os.path.splitext(sanitized_fn)[1].lower()
            if file_ext == '.pdf':
                pdf_attachments_to_merge.append(filepath)
            elif WIN32COM_AVAILABLE and file_ext in CONVERTIBLE_EXTENSIONS:
                pdf_converted_path = os.path.join(originals_folder_path, f"CONVERTED_{os.path.splitext(sanitized_fn)[0]}.pdf")
                conv_func = None
                if file_ext in ['.doc', '.docx', '.rtf', '.odt', '.txt']: conv_func = convert_document_to_pdf_msword
                elif file_ext in ['.xls', '.xlsx', '.ods', '.xlsm']: conv_func = convert_spreadsheet_to_pdf_msexcel
                elif file_ext in ['.ppt', '.pptx', '.odp']: conv_func = convert_presentation_to_pdf_msppt
                
                if conv_func and conv_func(filepath, pdf_converted_path):
                    pdf_attachments_to_merge.append(pdf_converted_path)
    
    # 3. Объединяем основной PDF с PDF-вложениями
    if pdf_attachments_to_merge and PYPDF_AVAILABLE:
        merged_pdf_path = os.path.join(DOWNLOADED_PDF_DIR, f"{unique_email_id}_merged.pdf")
        all_pdfs = [pdf_path] + pdf_attachments_to_merge
        if merge_pdfs(all_pdfs, merged_pdf_path):
            os.remove(pdf_path) # Удаляем временный PDF без вложений
            pdf_path = merged_pdf_path # Теперь основной PDF - это объединенный
            print(f"  -> PDF-вложения успешно объединены в главный файл.")

    return pdf_path, originals_folder_path

#
# --- БЛОК 2: ИНТЕРАКТИВНОЕ ПРИНЯТИЕ РЕШЕНИЙ ---
# Пользователь просматривает скачанные PDF и решает их судьбу.
#

def process_user_decisions(emails_metadata):
    """
    Основная функция этапа 2. Проходит по списку скачанных писем,
    открывает PDF для просмотра и запрашивает у пользователя решение.
    Возвращает отфильтрованный список писем, которые нужно сохранить.
    """
    if not emails_metadata:
        print("Не найдено писем для обработки.")
        return []

    emails_to_register = []
    total_emails = len(emails_metadata)
    
    print(f"\n--- Начат этап принятия решений ({total_emails} писем) ---")

    for i, email_data in enumerate(emails_metadata):
        print(f"\n--- Обработка письма {i + 1}/{total_emails} ---")
        print(f"  От: {email_data['sender']}")
        print(f"  Тема: {email_data['subject']}")

        pdf_path = email_data["pdf_path"]
        originals_path = email_data["originals_path"]
        
        open_file_for_review(pdf_path)
        time.sleep(2) # Даем время на открытие файла

        while True:
            choice = input("Действия для документа:\n  1. Сохранить (для последующей регистрации)\n  2. Удалить\nВаш выбор: ").strip()
            
            if choice == '1':
                # Спрашиваем про оригиналы
                while True:
                    keep_originals_choice = input("  -> Сохранить папку с оригиналами этого письма? (да/нет): ").lower().strip()
                    if keep_originals_choice in ['да', 'д', 'yes', 'y']:
                        print(f"  -> PDF и папка с оригиналами '{os.path.basename(originals_path)}' сохранены.")
                        emails_to_register.append(email_data)
                        break
                    elif keep_originals_choice in ['нет', 'н', 'no', 'n']:
                        try:
                            shutil.rmtree(originals_path)
                            print(f"  -> Папка с оригиналами '{os.path.basename(originals_path)}' удалена. PDF сохранен.")
                        except Exception as e:
                            print(f"  ERROR: Не удалось удалить папку с оригиналами: {e}")
                        emails_to_register.append(email_data)
                        break
                    else:
                        print("  ERROR: Неверный ввод. Пожалуйста, введите 'да' или 'нет'.")
                break # Выход из основного цикла while
            
            elif choice == '2':
                try:
                    os.remove(pdf_path)
                    shutil.rmtree(originals_path)
                    print(f"  -> PDF и папка с оригиналами '{os.path.basename(originals_path)}' удалены.")
                except Exception as e:
                    print(f"  ERROR: Ошибка при удалении файлов: {e}")
                break # Выход из основного цикла while
            else:
                print("ERROR: Неверный выбор. Пожалуйста, введите 1 или 2.")
    
    return emails_to_register

#
# --- БЛОК 3: РЕГИСТРАЦИЯ И ФОРМИРОВАНИЕ ЖУРНАЛА ---
# Финальный этап: переименование сохраненных PDF и создание CSV-журнала.
#

def register_saved_emails(emails_to_register):
    """
    Основная функция этапа 3. Запрашивает начальный номер, переименовывает
    сохраненные PDF и создает/дополняет CSV-журнал.
    """
    if not emails_to_register:
        print("\n--- Нет писем для регистрации. Завершение работы. ---")
        return
    
    print(f"\n--- Начат этап регистрации ({len(emails_to_register)} писем) ---")
    
    start_journal_num = prompt_for_starting_journal_number()
    current_journal_num = start_journal_num
    date_str_for_filename = datetime.datetime.now().strftime("%d.%m.%Y")
    
    # Проверяем, существует ли файл журнала, чтобы не перезаписывать заголовок
    journal_exists = os.path.exists(JOURNAL_CSV_FILE)

    try:
        with open(JOURNAL_CSV_FILE, 'a', newline='', encoding='utf-8-sig') as csvfile:
            writer = csv.writer(csvfile)
            
            if not journal_exists:
                writer.writerow(['Входящий номер', 'Дата регистрации', 'Отправитель', 'Тема письма'])
            
            for email_data in emails_to_register:
                # Формируем новое имя файла
                new_filename = f"вх.№ {current_journal_num} от {date_str_for_filename}.pdf"
                old_filepath = email_data['pdf_path']
                new_filepath = os.path.join(REGISTERED_DIR, new_filename)
                
                # Перемещаем и переименовываем PDF
                try:
                    shutil.move(old_filepath, new_filepath)
                    
                    # Записываем данные в журнал
                    writer.writerow([
                        f"вх.№ {current_journal_num}",
                        date_str_for_filename,
                        email_data['sender'],
                        email_data['subject']
                    ])
                    
                    print(f"  -> Зарегистрирован: {new_filename}")
                    current_journal_num += 1
                    
                except Exception as e:
                    print(f"  ERROR: Не удалось зарегистрировать файл '{os.path.basename(old_filepath)}': {e}")

        print(f"\n--- Регистрация завершена. Журнал сохранен в: {JOURNAL_CSV_FILE} ---")
        
    except Exception as e:
        print(f"CRITICAL: Ошибка при записи в CSV-журнал: {e}")

#
# --- ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ (взяты из вашего кода с минимальными изменениями) ---
# Этот блок содержит утилиты для работы с PDF, файлами, почтой и т.д.
#

def setup_directories():
    """Создает все необходимые папки, если их нет."""
    print("INFO: Проверка и создание рабочих папок...")
    os.makedirs(BASE_OUTPUT_DIRECTORY, exist_ok=True)
    os.makedirs(DOWNLOADED_PDF_DIR, exist_ok=True)
    os.makedirs(DOWNLOADED_ORIGINALS_DIR, exist_ok=True)
    os.makedirs(REGISTERED_DIR, exist_ok=True)
    print("INFO: Папки готовы.")

def open_file_for_review(filepath):
    """Кросс-платформенное открытие файла для просмотра."""
    try:
        if not os.path.exists(filepath):
            print(f"ERROR: Файл для просмотра не найден: {filepath}")
            return
        print(f"INFO: Открываю файл для ознакомления: {os.path.basename(filepath)}")
        if platform.system() == 'Windows': os.startfile(os.path.abspath(filepath))
        elif platform.system() == 'Darwin': subprocess.call(('open', os.path.abspath(filepath)))
        else: subprocess.call(('xdg-open', os.path.abspath(filepath)))
    except Exception as e:
        print(f"WARNING: Не удалось автоматически открыть файл. Пожалуйста, откройте его вручную. Ошибка: {e}")

def prompt_for_starting_journal_number():
    """Запрашивает у пользователя последний номер в журнале."""
    while True:
        try:
            last_num_str = input("Введите ПОСЛЕДНИЙ зарегистрированный номер входящего документа: ")
            last_num = int(last_num_str)
            if last_num < 0:
                print("ERROR: Номер не может быть отрицательным.")
                continue
            start_next_num = last_num + 1
            print(f"INFO: Регистрация начнется с вх.№ {start_next_num}.")
            return start_next_num
        except ValueError:
            print("ERROR: Пожалуйста, введите корректное число.")

def sanitize_filename(filename):
    """Очищает имя файла от недопустимых символов."""
    if not filename: return "untitled_attachment"
    return re.sub(r'[\\/*?:"<>|]', "_", filename)

# --- Функции для работы с email (из вашего кода) ---
def _extract_email_headers(msg):
    headers = {}
    date_tuple = email.utils.parsedate_tz(msg['Date'])
    if date_tuple:
        local_date = datetime.datetime.fromtimestamp(email.utils.mktime_tz(date_tuple))
        headers['formatted_date'] = local_date.strftime("%Y-%m-%d %H:%M:%S")
    else:
        headers['formatted_date'] = "Дата отсутствует"

    from_header = decode_header(msg.get('From', ''))
    headers['sender'] = "".join([p.decode(c or 'utf-8', 'replace') if isinstance(p, bytes) else str(p) for p, c in from_header])
    to_header = decode_header(msg.get('To', ''))
    headers['recipients'] = "".join([p.decode(c or 'utf-8', 'replace') if isinstance(p, bytes) else str(p) for p, c in to_header])
    subject_header = decode_header(msg.get('Subject', ''))
    headers['subject'] = "".join([p.decode(c or 'utf-8', 'replace') if isinstance(p, bytes) else str(p) for p, c in subject_header])
    return headers

def _extract_email_body(msg):
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            cdispo = str(part.get('Content-Disposition'))
            if ctype == 'text/plain' and 'attachment' not in cdispo:
                body = part.get_payload(decode=True).decode(part.get_content_charset() or 'utf-8', 'replace')
                break
    else:
        body = msg.get_payload(decode=True).decode(msg.get_content_charset() or 'utf-8', 'replace')
    
    if not body: # Если plain text не найден, пытаемся извлечь из HTML
         for part in msg.walk():
            if part.get_content_type() == 'text/html' and 'attachment' not in str(part.get('Content-Disposition')):
                html_body = part.get_payload(decode=True).decode(part.get_content_charset() or 'utf-8', 'replace')
                soup = BeautifulSoup(html_body, "html.parser")
                body = soup.get_text(separator='\n')
                break

    return body

# --- Функции для работы с PDF (из вашего кода) ---
def _setup_pdf_canvas_and_styles(temp_report_path):
    pdf_canvas = canvas.Canvas(temp_report_path, pagesize=A4)
    width, height = A4
    margin = 20 * mm
    content_width = width - 2 * margin
    page_dims = {'width': width, 'height': height, 'margin': margin, 'content_width': content_width}
    current_y = height - margin
    if os.path.exists(DEJAVU_SANS_FONT_PATH):
        pdfmetrics.registerFont(TTFont('DejaVuSans', DEJAVU_SANS_FONT_PATH))
        font_to_use = 'DejaVuSans'
    else:
        font_to_use = 'Helvetica'
    styles_all = getSampleStyleSheet()
    styleN = styles_all['Normal']; styleN.fontName = font_to_use; styleN.fontSize = 10
    styleBody = styles_all['Normal']; styleBody.fontName = font_to_use; styleBody.fontSize = 9
    styles = {'N': styleN, 'Body': styleBody}
    return pdf_canvas, styles, font_to_use, page_dims, current_y

def _add_paragraph_to_pdf_util(pdf_canvas, text, style, y_pos, page_dims):
    p = Paragraph(text.replace('\n', '<br/>'), style)
    p_w, p_h = p.wrapOn(pdf_canvas, page_dims['content_width'], page_dims['height'])
    if y_pos - p_h < page_dims['margin']:
        pdf_canvas.showPage()
        pdf_canvas.setFont(style.fontName, style.fontSize)
        y_pos = page_dims['height'] - page_dims['margin']
    p.drawOn(pdf_canvas, page_dims['margin'], y_pos - p_h)
    return y_pos - p_h

def merge_pdfs(list_of_pdf_paths, output_merged_pdf_path):
    if not PYPDF_AVAILABLE: return False
    merger = PdfWriter()
    try:
        for pdf_path in list_of_pdf_paths:
            if os.path.exists(pdf_path):
                merger.append(pdf_path)
        with open(output_merged_pdf_path, "wb") as f_out:
            merger.write(f_out)
        return True
    except Exception as e:
        print(f"  ERROR: Ошибка при объединении PDF: {e}")
        return False
    finally:
        merger.close()

# --- Функции конвертации MS Office (из вашего кода) ---
if WIN32COM_AVAILABLE:
    def convert_document_to_pdf_msword(input_path, output_path):
        word = None; doc = None
        try:
            word = win32com.client.Dispatch("Word.Application"); word.Visible = False
            doc = word.Documents.Open(os.path.abspath(input_path), ReadOnly=True)
            doc.SaveAs(os.path.abspath(output_path), FileFormat=17)
            print(f"  -> Сконвертировано в PDF (Word): {os.path.basename(input_path)}")
            return True
        except Exception as e: print(f"  ERROR: Ошибка конвертации Word: {e}"); return False
        finally:
            if doc: doc.Close(False)
            if word: word.Quit()

    def convert_spreadsheet_to_pdf_msexcel(input_path, output_path):
        excel = None; workbook = None
        try:
            excel = win32com.client.Dispatch("Excel.Application"); excel.Visible = False
            workbook = excel.Workbooks.Open(os.path.abspath(input_path), ReadOnly=True)
            workbook.ExportAsFixedFormat(0, os.path.abspath(output_path))
            print(f"  -> Сконвертировано в PDF (Excel): {os.path.basename(input_path)}")
            return True
        except Exception as e: print(f"  ERROR: Ошибка конвертации Excel: {e}"); return False
        finally:
            if workbook: workbook.Close(False)
            if excel: excel.Quit()
            
    def convert_presentation_to_pdf_msppt(input_path, output_path):
        powerpoint = None; presentation = None
        try:
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            presentation = powerpoint.Presentations.Open(os.path.abspath(input_path), ReadOnly=True, WithWindow=False)
            presentation.SaveAs(os.path.abspath(output_path), FileFormat=32)
            print(f"  -> Сконвертировано в PDF (PowerPoint): {os.path.basename(input_path)}")
            return True
        except Exception as e: print(f"  ERROR: Ошибка конвертации PowerPoint: {e}"); return False
        finally:
            if presentation: presentation.Close()
            if powerpoint: powerpoint.Quit()

#
# --- ГЛАВНЫЙ БЛОК ИСПОЛНЕНИЯ ---
#
if __name__ == "__main__":
    # 0. Создаем папки
    setup_directories()
    
    # 1. Этап скачивания
    downloaded_emails = download_all_unseen_emails()
    
    # 2. Этап принятия решений
    emails_for_registration = process_user_decisions(downloaded_emails)
    
    # 3. Этап регистрации
    register_saved_emails(emails_for_registration)
    
    print("\nРабота приложения завершена.")