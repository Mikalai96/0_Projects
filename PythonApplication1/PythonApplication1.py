import imaplib
import os
from dotenv import load_dotenv
import email
from email.header import decode_header
import datetime
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, Image
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from bs4 import BeautifulSoup
import re
from io import BytesIO
import traceback
import sys
import platform
import subprocess
import time
import shutil

# --- Глобальные константы и проверки зависимостей ---
def resource_path(relative_path):
    """ Получить абсолютный путь к ресурсу, работает для разработки и для PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

try:
    from pypdf import PdfWriter
    PYPDF_AVAILABLE = True
    print("INFO: Библиотека pypdf найдена, объединение PDF будет доступно.")
except ImportError:
    PYPDF_AVAILABLE = False
    print("ПРЕДУПРЕЖДЕНИЕ: Библиотека pypdf не найдена. PDF-вложения не будут объединены с основным отчетом.")

WIN32COM_AVAILABLE = False
if os.name == 'nt':
    try:
        import win32com.client
        WIN32COM_AVAILABLE = True
        print("INFO: Библиотека pywin32 найдена, COM-автоматизация MS Office доступна.")
    except ImportError:
        print("ПРЕДУПРЕЖДЕНИЕ: Библиотека pywin32 не найдена. Конвертация файлов MS Office в PDF будет недоступна.")
else:
    print("INFO: Скрипт запущен не на Windows. Конвертация файлов MS Office в PDF будет недоступна.")

DEJAVU_SANS_FONT_PATH = resource_path("DejaVuSans.ttf")
PDF_OUTPUT_DIRECTORY = "saved_emails_pdf"
ATTACHMENTS_SUBDIRECTORY_NAME = "attachments"
PDF_REGISTERED_SUBDIR_NAME = "registered_emails"
PDF_DOWNLOADED_SUBDIR_NAME = "downloaded_emails"
SCANNED_INPUT_SUBDIR_NAME = "scanned_manual_input"

CONVERTIBLE_EXTENSIONS = ['.doc','.docx','.rtf','.odt','.txt', '.xls','.xlsx','.ods', '.ppt','.pptx','.odp','.xlsm']

# --- Функции, не измененные или мало измененные ---
def open_file_for_review(filepath):
    try:
        abs_filepath = os.path.abspath(filepath)
        if not os.path.exists(abs_filepath):
            print(f"ERROR: Файл для просмотра не найден: {abs_filepath}")
            return False
        print(f"INFO: Попытка открыть файл для ознакомления: {abs_filepath}")
        if platform.system() == 'Darwin': subprocess.call(('open', abs_filepath))
        elif platform.system() == 'Windows': os.startfile(abs_filepath)
        else: subprocess.call(('xdg-open', abs_filepath))
        input("INFO: Файл открыт для ознакомления. ПОЖАЛУЙСТА, ЗАКРОЙТЕ ОКНО ПРОСМОТРА PDF, затем нажмите Enter для продолжения...")
        time.sleep(1)
        return True
    except Exception as e:
        print(f"WARNING: Не удалось автоматически открыть файл '{abs_filepath}'. Пожалуйста, откройте его вручную. Ошибка: {e}")
        input("INFO: Нажмите Enter, чтобы продолжить...")
        return False

def handle_user_decision_for_pdf(path_for_review, journal_number, date_str_for_filename, subject_for_filename, attachments_folder_path=None):
    if not open_file_for_review(path_for_review):
        print(f"ERROR: Не удалось открыть файл {path_for_review} для ознакомления. Операция прервана.")
        return "ERROR", path_for_review

    registered_dir = os.path.join(PDF_OUTPUT_DIRECTORY, PDF_REGISTERED_SUBDIR_NAME)
    downloaded_dir = os.path.join(PDF_OUTPUT_DIRECTORY, PDF_DOWNLOADED_SUBDIR_NAME)

    while True:
        print(f"\nДействия для документа (на основе {os.path.basename(path_for_review)}):")
        print(f"  1. Сохранить с присвоением входящего номера (вх.№ {journal_number}) (в папку '{PDF_REGISTERED_SUBDIR_NAME}')")
        print(f"  2. Скачать без присвоения номера (в папку '{PDF_DOWNLOADED_SUBDIR_NAME}')")
        print(f"  3. Пропустить/Удалить этот документ (включая его вложения)")
        choice = input("Ваш выбор (1, 2 или 3): ").strip()

        final_save_path_dir = PDF_OUTPUT_DIRECTORY # Default, if specific dir creation fails

        if choice == '1':
            try:
                if not os.path.exists(registered_dir):
                    os.makedirs(registered_dir)
                    print(f"INFO: Создана папка для зарегистрированных файлов: {registered_dir}")
                final_save_path_dir = registered_dir
            except OSError as e:
                print(f"ERROR: Не удалось создать папку '{registered_dir}': {e}. Файл будет сохранен в '{PDF_OUTPUT_DIRECTORY}'.")

            final_filename = f"вх.№ {journal_number} от {date_str_for_filename}.pdf"
            final_save_path = os.path.join(final_save_path_dir, final_filename)
            try:
                if os.path.exists(final_save_path) and os.path.abspath(final_save_path) != os.path.abspath(path_for_review):
                    print(f"WARNING: Файл {final_save_path} уже существует. Будет перезаписан.")
                    os.remove(final_save_path)
                if os.path.abspath(final_save_path) != os.path.abspath(path_for_review):
                     os.rename(path_for_review, final_save_path)
                else:
                    print(f"INFO: Файл уже имеет имя {final_save_path}.")
                print(f"INFO: Файл сохранен как: {final_save_path}")
                return "SAVED", final_save_path
            except PermissionError as e_perm:
                print(f"ERROR: Ошибка доступа к файлу '{final_save_path}'. ВЕРОЯТНО, ФАЙЛ ВСЕ ЕЩЕ ОТКРЫТ В ПРОГРАММЕ ПРОСМОТРА PDF. Пожалуйста, ЗАКРОЙТЕ программу просмотра PDF и попробуйте снова. ({e_perm})")
            except Exception as e:
                print(f"ERROR: Не удалось сохранить файл как {final_save_path}: {e}")
                print(f"INFO: Исходный файл для ознакомления остался здесь: {path_for_review}")
                return "ERROR", path_for_review

        elif choice == '2':
            try:
                if not os.path.exists(downloaded_dir):
                    os.makedirs(downloaded_dir)
                    print(f"INFO: Создана папка для скачанных файлов: {downloaded_dir}")
                final_save_path_dir = downloaded_dir
            except OSError as e:
                print(f"ERROR: Не удалось создать папку '{downloaded_dir}': {e}. Файл будет сохранен в '{PDF_OUTPUT_DIRECTORY}'.")

            sane_subject = sanitize_filename(subject_for_filename[:50]) if subject_for_filename else "без_темы"
            download_filename = f"скачано_{date_str_for_filename}_{sane_subject}.pdf"
            download_save_path = os.path.join(final_save_path_dir, download_filename)
            try:
                if os.path.exists(download_save_path) and os.path.abspath(download_save_path) != os.path.abspath(path_for_review):
                    print(f"WARNING: Файл {download_save_path} уже существует. Будет перезаписан.")
                    os.remove(download_save_path)
                if os.path.abspath(download_save_path) != os.path.abspath(path_for_review):
                    os.rename(path_for_review, download_save_path)
                else:
                    print(f"INFO: Файл уже имеет имя {download_save_path}.")
                print(f"INFO: Файл скачан как: {download_save_path}")
                return "DOWNLOADED", download_save_path
            except PermissionError as e_perm:
                print(f"ERROR: Ошибка доступа к файлу '{download_save_path}'. ВЕРОЯТНО, ФАЙЛ ВСЕ ЕЩЕ ОТКРЫТ В ПРОГРАММЕ ПРОСМОТРА PDF. Пожалуйста, ЗАКРОЙТЕ программу просмотра PDF и попробуйте снова. ({e_perm})")
            except Exception as e:
                print(f"ERROR: Не удалось скачать файл как {download_save_path}: {e}")
                print(f"INFO: Исходный файл для ознакомления остался здесь: {path_for_review}")
                return "ERROR", path_for_review

        elif choice == '3':
            pdf_deleted = False
            try:
                os.remove(path_for_review)
                print(f"INFO: Временный PDF файл {path_for_review} удален.")
                pdf_deleted = True
            except PermissionError as e_perm:
                print(f"ERROR: Ошибка доступа при удалении PDF файла '{path_for_review}'. ВЕРОЯТНО, ФАЙЛ ВСЕ ЕЩЕ ОТКРЫТ. Пожалуйста, ЗАКРОЙТЕ его и попробуйте снова. ({e_perm})")
                continue
            except Exception as e:
                print(f"ERROR: Не удалось удалить PDF файл {path_for_review}: {e}")
                return "ERROR", path_for_review # Возвращаем путь, чтобы он не потерялся

            if pdf_deleted:
                if attachments_folder_path and os.path.exists(attachments_folder_path):
                    try:
                        shutil.rmtree(attachments_folder_path)
                        print(f"INFO: Папка вложений {attachments_folder_path} удалена.")
                    except Exception as e_rmtree:
                        print(f"WARNING: Не удалось удалить папку вложений {attachments_folder_path}: {e_rmtree}")
                elif attachments_folder_path:
                     print(f"INFO: Папка вложений {attachments_folder_path} не найдена или уже удалена.")
                else:
                    print(f"INFO: Нет информации о папке вложений для удаления или она не была создана.")
                return "SKIPPED", None
        else:
            print("ERROR: Неверный выбор. Пожалуйста, введите 1, 2 или 3.")

def prompt_for_starting_journal_number():
    print(f"INFO: Требуется указать начальный номер для регистрации входящих документов для этой сессии.")
    while True:
        try:
            last_num_str = input("INFO: Введите ПОСЛЕДНИЙ зарегистрированный номер входящего документа по журналу: ")
            last_num = int(last_num_str)
            if last_num < 0:
                print("ERROR: Номер не может быть отрицательным.")
                continue
            start_next_num = last_num + 1
            print(f"INFO: Обработка писем/сканов в этой сессии начнется с вх.№ {start_next_num}.")
            return start_next_num
        except ValueError:
            print("ERROR: Пожалуйста, введите корректное число.")

def merge_pdfs(list_of_pdf_paths, output_merged_pdf_path):
    if not PYPDF_AVAILABLE:
        print("      WARNING: Слияние PDF невозможно, т.к. библиотека pypdf не доступна.")
        return False
    merger = PdfWriter()
    merged_something = False
    try:
        for pdf_path in list_of_pdf_paths:
            if os.path.exists(pdf_path):
                try:
                    merger.append(pdf_path)
                    print(f"      DEBUG: Добавлен '{os.path.basename(pdf_path)}' для слияния.")
                    merged_something = True
                except Exception as e_append:
                    print(f"      WARNING: Не удалось добавить PDF '{os.path.basename(pdf_path)}' для слияния: {e_append}")
            else:
                print(f"      WARNING: PDF-файл для объединения не найден: {pdf_path}")
        if not merged_something:
            print("      WARNING: Нет PDF-файлов для фактического слияния.")
            return False # merger.close() будет вызван в finally
        with open(output_merged_pdf_path, "wb") as f_out:
            merger.write(f_out)
        print(f"      INFO: PDF-файлы успешно объединены в: {output_merged_pdf_path}")
        return True
    except Exception as e:
        print(f"      ERROR: Ошибка при объединении PDF-файлов: {e}")
        traceback.print_exc()
        return False
    finally:
        merger.close()

if WIN32COM_AVAILABLE:
    def convert_document_to_pdf_msword(input_path, output_path):
        word = None; doc = None
        input_path_abs = os.path.abspath(input_path); output_path_abs = os.path.abspath(output_path)
        try:
            word = win32com.client.Dispatch("Word.Application"); word.Visible = False
            doc = word.Documents.Open(input_path_abs, ReadOnly=True)
            wdFormatPDF = 17
            doc.SaveAs(output_path_abs, FileFormat=wdFormatPDF)
            print(f"          INFO: Файл '{os.path.basename(input_path_abs)}' успешно сконвертирован в PDF с помощью Word.")
            return True
        except Exception as e: print(f"          ERROR: Ошибка конвертации Word '{os.path.basename(input_path_abs)}': {e}"); return False
        finally:
            if doc: doc.Close(False)
            if word: word.Quit()

    def convert_spreadsheet_to_pdf_msexcel(input_path, output_path):
        excel = None
        workbook = None
        input_path_abs = os.path.abspath(input_path)
        output_path_abs = os.path.abspath(output_path)
    
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
        
            workbook = excel.Workbooks.Open(input_path_abs, ReadOnly=True)
        
            # Применяем настройки к каждому листу для корректного отображения
            for worksheet in workbook.Worksheets:
                # PageSetup properties: https://docs.microsoft.com/en-us/office/vba/api/excel.pagesetup
                ws_setup = worksheet.PageSetup
                # Устанавливаем альбомную ориентацию (2 = xlLandscape)
                ws_setup.Orientation = 2 
                # Выключаем ручное масштабирование
                ws_setup.Zoom = False
                # Вписываем по ширине на 1 страницу
                ws_setup.FitToPagesWide = 1
                # Высота может быть любой (False означает "авто")
                ws_setup.FitToPagesTall = False
        
            # Тип экспорта в PDF
            xlTypePDF = 0
            # Экспортируем весь Workbook, а не активный лист
            workbook.ExportAsFixedFormat(xlTypePDF, output_path_abs)
        
            print(f"            INFO: Файл '{os.path.basename(input_path_abs)}' успешно сконвертирован в PDF с авто-масштабированием.")
            return True
        except Exception as e:
            print(f"            ERROR: Ошибка конвертации Excel '{os.path.basename(input_path_abs)}': {e}")
            traceback.print_exc()
            return False
        finally:
            if workbook:
                workbook.Close(False) # Закрываем без сохранения изменений
            if excel:
                excel.Quit()

    def convert_presentation_to_pdf_msppt(input_path, output_path):
        powerpoint = None; presentation = None
        input_path_abs = os.path.abspath(input_path); output_path_abs = os.path.abspath(output_path)
        try:
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            presentation = powerpoint.Presentations.Open(input_path_abs, ReadOnly=True, WithWindow=False)
            ppSaveAsPDF = 32
            presentation.SaveAs(output_path_abs, FileFormat=ppSaveAsPDF)
            print(f"          INFO: Файл '{os.path.basename(input_path_abs)}' успешно сконвертирован в PDF с помощью PowerPoint.")
            return True
        except Exception as e: print(f"          ERROR: Ошибка конвертации PowerPoint '{os.path.basename(input_path_abs)}': {e}"); return False
        finally:
            if presentation: presentation.Close()
            if powerpoint: powerpoint.Quit()

def sanitize_filename(filename):
    if not filename: return "untitled_attachment"
    filename = re.sub(r'[^\w\.\-\u0400-\u04FF]', '_', filename, flags=re.UNICODE)
    filename = re.sub(r'__+', '_', filename)
    filename = filename.strip('_')
    if filename.startswith('.'): filename = "_" + filename
    return filename if filename else "sanitized_attachment"
# --- Вспомогательные функции для open_email ---

def _get_email_message_object(mail_obj, email_id):
    """Извлекает и парсит email сообщение."""
    print(f"    DEBUG: Получение письма ID {email_id}")
    status, data = mail_obj.fetch(email_id, '(RFC822)')
    if status != 'OK':
        print(f"    ERROR: Ошибка получения письма ID {email_id}")
        return None
    return email.message_from_bytes(data[0][1])

def _extract_email_headers(msg):
    """Извлекает и декодирует основные заголовки письма."""
    headers = {}
    date_header = msg.get('Date')
    if date_header:
        try:
            email_actual_date = email.utils.parsedate_to_datetime(date_header)
            headers['formatted_date'] = email_actual_date.strftime("%Y-%m-%d %H:%M:%S") if email_actual_date else f"Не удалось распознать дату: {date_header}"
        except Exception as e_date:
            headers['formatted_date'] = f"Ошибка даты: {e_date}"
    else:
        headers['formatted_date'] = "Отсутствует"

    from_header_raw = msg.get('From')
    if from_header_raw:
        name, addr = email.utils.parseaddr(from_header_raw)
        decoded_name_parts = [p.decode(c or 'utf-8', 'replace') if isinstance(p, bytes) else str(p) for p, c in decode_header(name)] if name else []
        sender_name_str = "".join(decoded_name_parts).strip()
        sender_email_str = addr.strip()
        if sender_name_str and sender_email_str: headers['sender'] = f"{sender_name_str} <{sender_email_str}>"
        elif sender_email_str: headers['sender'] = sender_email_str
        elif sender_name_str: headers['sender'] = sender_name_str
        else: headers['sender'] = "".join([p.decode(c or 'utf-8', 'replace') if isinstance(p, bytes) else str(p) for p, c in decode_header(from_header_raw)]).strip() or from_header_raw
    else:
        headers['sender'] = "Отправитель отсутствует"

    to_header = msg.get('To')
    if to_header:
        headers['recipients'] = ', '.join([addr for _, addr in email.utils.getaddresses([to_header]) if addr]) or "Не удалось извлечь"
    else:
        headers['recipients'] = "Получатели отсутствуют"

    subject_header = msg.get('Subject')
    if subject_header:
        headers['subject'] = "".join([p.decode(c or 'utf-8', 'replace') if isinstance(p, bytes) else str(p) for p, c in decode_header(subject_header)])
    else:
        headers['subject'] = "Тема отсутствует"
    return headers

def _extract_email_body(msg):
    """Извлекает текстовое содержимое тела письма."""
    plain_text_content, html_content_str = None, None
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            cdisp = str(part.get("Content-Disposition"))
            if "attachment" in cdisp.lower() or part.get_filename() or part.is_multipart():
                continue
            if not (ctype.startswith("text/") or ctype.startswith("multipart/alternative")):
                continue
            try:
                payload = part.get_payload(decode=True)
                if payload is None: continue
                charset = part.get_content_charset() or 'utf-8'
                dec_payload = payload.decode(charset, 'replace')
            except UnicodeDecodeError:
                dec_payload = payload.decode('latin-1', 'replace')
            except Exception as e_decode:
                print(f"        ERROR: Неожиданная ошибка декодирования части тела: {e_decode}")
                continue

            if ctype == "text/plain" and not plain_text_content:
                plain_text_content = dec_payload
            elif ctype == "text/html" and not html_content_str:
                html_content_str = dec_payload
        if plain_text_content: return plain_text_content
        if html_content_str:
            try:
                soup = BeautifulSoup(html_content_str, "html.parser")
                body_text = soup.get_text(separator='\n', strip=True)
                return body_text if body_text else "HTML (без извлекаемого текста)."
            except Exception as e_bs: return f"Ошибка BeautifulSoup (тело HTML): {e_bs}"
        return "Содержимое не найдено (многочастное)."
    else: # Не многочастное письмо
        payload = msg.get_payload(decode=True)
        ctype = msg.get_content_type()
        charset = msg.get_content_charset() or 'utf-8'
        if payload is None: return "Содержимое отсутствует (payload is None)."
        try:
            dec_payload = payload.decode(charset, 'replace')
        except UnicodeDecodeError:
            dec_payload = payload.decode('latin-1', 'replace')
        except Exception as e_decode: return f"Ошибка декодирования содержимого: {e_decode}"

        if ctype == "text/plain": return dec_payload
        if ctype == "text/html":
            try:
                soup = BeautifulSoup(dec_payload, "html.parser")
                body_text = soup.get_text(separator='\n', strip=True)
                return body_text if body_text else "HTML (без извлекаемого текста, не многочастное)."
            except Exception as e_bs_s: return f"Ошибка BeautifulSoup (тело HTML, не многочастное): {e_bs_s}"
        return f"Содержимое письма имеет тип {ctype}, не является стандартным текстом."

def _setup_pdf_canvas_and_styles(temp_report_path):
    """Инициализирует PDF canvas и стили ReportLab."""
    pdf_canvas = canvas.Canvas(temp_report_path, pagesize=A4)
    width, height = A4
    margin = 20 * mm
    content_width = width - 2 * margin
    page_dims = {'width': width, 'height': height, 'margin': margin, 'content_width': content_width}
    current_y = height - margin

    font_to_use = 'Helvetica'
    if os.path.exists(DEJAVU_SANS_FONT_PATH):
        try:
            pdfmetrics.registerFont(TTFont('DejaVuSans', DEJAVU_SANS_FONT_PATH))
            font_to_use = 'DejaVuSans'
        except Exception as e_font:
            print(f"    INFO: Ошибка регистрации шрифта '{DEJAVU_SANS_FONT_PATH}': {e_font}. Используется Helvetica.")
    else:
        print(f"    INFO: Шрифт '{DEJAVU_SANS_FONT_PATH}' не найден. Используется Helvetica.")

    styles_all = getSampleStyleSheet()
    styleN = styles_all['Normal']; styleN.fontName = font_to_use; styleN.fontSize = 10; styleN.leading = 12
    styleH = styles_all['Heading3']; styleH.fontName = font_to_use; styleH.fontSize = 11; styleH.leading = 13; styleH.spaceBefore = 6; styleH.spaceAfter = 2
    styleBody = styles_all['Normal']; styleBody.fontName = font_to_use; styleBody.fontSize = 9; styleBody.leading = 11
    styles = {'N': styleN, 'H': styleH, 'Body': styleBody}

    return pdf_canvas, styles, font_to_use, page_dims, current_y

def _add_paragraph_to_pdf_util(pdf_canvas, text, style, y_pos, page_dims):
    """Вспомогательная функция для добавления параграфа в PDF."""
    p = Paragraph(text.replace('\n', '<br/>'), style)
    p_w, p_h = p.wrapOn(pdf_canvas, page_dims['content_width'], page_dims['height']) # Используем полную высоту страницы для wrap
    if y_pos - p_h < page_dims['margin']:
        pdf_canvas.showPage()
        pdf_canvas.setFont(style.fontName, style.fontSize)
        y_pos = page_dims['height'] - page_dims['margin']
    p.drawOn(pdf_canvas, page_dims['margin'], y_pos - p_h)
    return y_pos - p_h - (style.leading * 0.2) # Небольшой отступ после параграфа

def _add_email_headers_to_pdf(pdf_canvas, headers, body, journal_number, display_date_str, styles, current_y, page_dims):
    """Добавляет информацию о письме (заголовки, тело) в PDF."""
    current_y = _add_paragraph_to_pdf_util(pdf_canvas, f"<b>Вх. №:</b> {journal_number} от {display_date_str}", styles['N'], current_y, page_dims)
    current_y = _add_paragraph_to_pdf_util(pdf_canvas, f"<b>Дата получения (факт):</b> {headers['formatted_date']}", styles['N'], current_y, page_dims)
    current_y = _add_paragraph_to_pdf_util(pdf_canvas, f"<b>От:</b> {headers['sender']}", styles['N'], current_y, page_dims)
    current_y = _add_paragraph_to_pdf_util(pdf_canvas, f"<b>Кому:</b> {headers['recipients']}", styles['N'], current_y, page_dims)
    current_y = _add_paragraph_to_pdf_util(pdf_canvas, f"<b>Тема:</b> {headers['subject']}", styles['N'], current_y, page_dims)
    current_y = _add_paragraph_to_pdf_util(pdf_canvas, "<b>Содержание:</b>", styles['N'], current_y, page_dims)
    current_y = _add_paragraph_to_pdf_util(pdf_canvas, body if body.strip() else "Содержимое отсутствует.", styles['Body'], current_y, page_dims)
    print(f"    Содержание (первые 200 симв.): {body[:200].replace(chr(10), ' ')}...")
    return current_y

def _create_attachment_directory(raw_email_att_dir_name):
    """Создает директорию для вложений письма."""
    email_att_path_base = os.path.join(PDF_OUTPUT_DIRECTORY, ATTACHMENTS_SUBDIRECTORY_NAME)
    sanitized_email_att_subdir_name = sanitize_filename(raw_email_att_dir_name)
    email_att_specific_path = os.path.join(email_att_path_base, sanitized_email_att_subdir_name)

    if not os.path.exists(email_att_path_base):
        try: os.makedirs(email_att_path_base)
        except OSError as e_base_dir: print(f"        ERROR: Ошибка создания базовой папки вложений '{email_att_path_base}': {e_base_dir}")
    
    # Проверяем, нужно ли создавать папку для вложений этого конкретного письма
    # (это будет сделано позже, только если есть не-встраиваемые вложения)
    return email_att_specific_path # Возвращаем путь, создание будет позже

def _ensure_specific_attachment_folder_exists(path):
    """Гарантирует существование папки для вложений конкретного письма."""
    if not os.path.exists(path):
        try:
            os.makedirs(path)
            print(f"        INFO: Создана папка для вложений письма: {path}")
            return True
        except OSError as e_mkdir:
            print(f"          ERROR: Не удалось создать папку '{path}': {e_mkdir}.")
            return False
    return True


def _handle_image_attachment(pdf_canvas, attachment_data, decoded_fn, current_y, styles, font_to_use, page_dims):
    """Обрабатывает вложение-изображение и добавляет его в PDF."""
    try:
        img_data = BytesIO(attachment_data)
        img = Image(img_data)
        img_w, img_h = img.drawWidth, img.drawHeight
        aspect = img_h / float(img_w) if img_w > 0 else 1
        disp_w = page_dims['content_width']
        disp_h = disp_w * aspect
        max_img_h = current_y - page_dims['margin'] - 15 * mm # Оставляем место

        if disp_h > max_img_h:
            disp_h = max(max_img_h, 10 * mm) # Минимальная высота
            disp_w = disp_h / aspect if aspect > 1e-6 else page_dims['content_width']
        if disp_w > page_dims['content_width']:
            disp_w = page_dims['content_width']
            disp_h = disp_w * aspect
        
        if current_y - disp_h < page_dims['margin']:
            pdf_canvas.showPage()
            pdf_canvas.setFont(font_to_use, styles['N'].fontSize) # Убедимся, что шрифт установлен
            current_y = page_dims['height'] - page_dims['margin']
            current_y = _add_paragraph_to_pdf_util(pdf_canvas, f"<b>Вложение (продолжение): {decoded_fn}</b>", styles['H'], current_y, page_dims)

        if disp_w > 0 and disp_h > 0:
            img.drawOn(pdf_canvas, page_dims['margin'], current_y - disp_h, width=disp_w, height=disp_h)
            current_y -= (disp_h + 3 * mm)
            print(f"          INFO: Изображение '{decoded_fn}' встроено.")
            return current_y, " - встроено в PDF."
        else:
            raise ValueError("Некорректные размеры изображения после расчета.")
    except Exception as e_img:
        print(f"          ERROR: Ошибка встраивания изображения '{decoded_fn}': {e_img}")
        current_y = _add_paragraph_to_pdf_util(pdf_canvas, f"<i>Ошибка встраивания изображения '{decoded_fn}': {e_img}</i>", styles['N'], current_y, page_dims)
        return current_y, f" - ошибка встраивания ({e_img})."


def _handle_text_attachment(pdf_canvas, attachment_data, decoded_fn, current_y, styles, page_dims, part, is_html=False):
    """Обрабатывает текстовое или HTML вложение и добавляет его в PDF."""
    try:
        charset = part.get_content_charset() or 'utf-8'
        text_content = attachment_data.decode(charset, 'replace') if attachment_data else ""
        
        prefix = "HTML (как текст)" if is_html else "Текстовое содержимое"
        if is_html and text_content:
            try:
                text_content = BeautifulSoup(text_content, "html.parser").get_text(separator='\n', strip=True) or "(HTML без извлекаемого текста)"
            except Exception as e_bs_att:
                text_content = f"(Ошибка извлечения HTML из вложения: {e_bs_att})"

        current_y = _add_paragraph_to_pdf_util(pdf_canvas, f"<u>{prefix} '{decoded_fn}':</u><br/>{text_content[:2000]}{'...' if len(text_content)>2000 else ''}", styles['Body'], current_y, page_dims)
        info = " - HTML (как текст) добавлен в PDF." if is_html else " - текст добавлен в PDF."
        print(f"          INFO: {prefix} из '{decoded_fn}' добавлен в PDF.")
        return current_y, info
    except Exception as e_text_att:
        print(f"          ERROR: Ошибка добавления текстового/HTML вложения '{decoded_fn}': {e_text_att}")
        current_y = _add_paragraph_to_pdf_util(pdf_canvas, f"<i>Ошибка добавления содержимого '{decoded_fn}': {e_text_att}</i>", styles['N'], current_y, page_dims)
        return current_y, f" - ошибка добавления содержимого ({e_text_att})."

def _handle_other_attachment(pdf_canvas, attachment_data, decoded_fn, sanitized_fn_for_saving, part_counter, email_att_specific_path, att_content_type_lower, current_y, styles, page_dims):
    """Обрабатывает прочие типы вложений (сохранение, конвертация)."""
    path_to_potential_pdf_attachment = None
    att_info_suffix = ""

    # Сначала обеспечиваем существование папки для вложений этого письма
    specific_folder_ok = _ensure_specific_attachment_folder_exists(email_att_specific_path)
    if not specific_folder_ok:
        att_info_suffix = " - не обработано (ошибка создания папки вложений письма)."
        print(f"          WARNING: Пропуск сохранения файла '{decoded_fn}' из-за ошибки папки вложений письма.")
        current_y = _add_paragraph_to_pdf_util(pdf_canvas, f"<i>Файл '{decoded_fn}' не сохранен на диск (ошибка папки).</i>", styles['N'], current_y, page_dims)
        return current_y, att_info_suffix, path_to_potential_pdf_attachment

    orig_fname_with_prefix = f"{part_counter}_original_{sanitized_fn_for_saving}"
    orig_path_abs = os.path.abspath(os.path.join(email_att_specific_path, orig_fname_with_prefix))
    saved_original_successfully = False
    try:
        with open(orig_path_abs, "wb") as f_a:
            f_a.write(attachment_data)
        print(f"          INFO: Оригинал '{decoded_fn}' сохранен как '{orig_fname_with_prefix}' в '{email_att_specific_path}'")
        saved_original_successfully = True
    except Exception as e_save_o:
        print(f"          ERROR: Не удалось сохранить оригинал '{decoded_fn}': {e_save_o}")
        att_info_suffix = f" - ошибка сохранения оригинала ({e_save_o})."
        current_y = _add_paragraph_to_pdf_util(pdf_canvas, f"<i>Ошибка сохранения оригинала '{decoded_fn}': {e_save_o}</i>", styles['N'], current_y, page_dims)
        return current_y, att_info_suffix, path_to_potential_pdf_attachment

    if saved_original_successfully:
        file_ext = os.path.splitext(sanitized_fn_for_saving)[1].lower()
        is_pdf_already = (att_content_type_lower == "application/pdf" or file_ext == ".pdf")

        if is_pdf_already:
            final_pdf_name_att = f"{part_counter}_{sanitized_fn_for_saving}"
            final_pdf_path_abs_att = os.path.abspath(os.path.join(email_att_specific_path, final_pdf_name_att))
            try:
                if os.path.exists(final_pdf_path_abs_att) and final_pdf_path_abs_att != orig_path_abs:
                    os.remove(final_pdf_path_abs_att)
                os.rename(orig_path_abs, final_pdf_path_abs_att)
                path_to_potential_pdf_attachment = final_pdf_path_abs_att
                att_info_suffix = " - PDF-вложение будет объединено."
                current_y = _add_paragraph_to_pdf_util(pdf_canvas, f"<i>PDF-вложение '{decoded_fn}' будет добавлено в конец этого документа.</i>", styles['N'], current_y, page_dims)
            except Exception as e_rename:
                print(f"          WARNING: Не удалось переименовать '{orig_fname_with_prefix}' в PDF: {e_rename}. Используется как есть.")
                path_to_potential_pdf_attachment = orig_path_abs # Используем сохраненный оригинал
                att_info_suffix = f" - PDF-вложение ({orig_fname_with_prefix}) будет объединено."
                current_y = _add_paragraph_to_pdf_util(pdf_canvas, f"<i>PDF-вложение '{decoded_fn}' ({orig_fname_with_prefix}) будет добавлено.</i>", styles['N'], current_y, page_dims)
        
        elif WIN32COM_AVAILABLE and file_ext in CONVERTIBLE_EXTENSIONS:
            conv_pdf_name = f"{part_counter}_converted_{os.path.splitext(sanitized_fn_for_saving)[0]}.pdf"
            conv_pdf_path_abs = os.path.abspath(os.path.join(email_att_specific_path, conv_pdf_name))
            com_func = None
            if file_ext in ['.doc', '.docx', '.rtf', '.odt', '.txt']: com_func = convert_document_to_pdf_msword
            elif file_ext in ['.xls', '.xlsx', '.ods', '.xlsm']: com_func = convert_spreadsheet_to_pdf_msexcel
            elif file_ext in ['.ppt', '.pptx', '.odp']: com_func = convert_presentation_to_pdf_msppt

            if com_func and com_func(orig_path_abs, conv_pdf_path_abs):
                path_to_potential_pdf_attachment = conv_pdf_path_abs
                att_info_suffix = " - сконвертировано в PDF, будет объединено."
                current_y = _add_paragraph_to_pdf_util(pdf_canvas, f"<i>Файл '{decoded_fn}' сконвертирован в PDF ({conv_pdf_name}) и будет добавлен.</i>", styles['N'], current_y, page_dims)
                try: os.remove(orig_path_abs); print(f"            INFO: Оригинал '{orig_fname_with_prefix}' удален после конвертации.")
                except Exception as e_del_orig: print(f"            WARNING: Не удалось удалить оригинал '{orig_fname_with_prefix}' после конвертации: {e_del_orig}")
            else:
                att_info_suffix = f" - оригинал ({orig_fname_with_prefix}) сохранен, конвертация не удалась."
                current_y = _add_paragraph_to_pdf_util(pdf_canvas, f"<i>Оригинал сохранен: {orig_fname_with_prefix}. Конвертация в PDF не удалась.</i>", styles['N'], current_y, page_dims)
        else:
            reason = "не Windows/pywin32" if not WIN32COM_AVAILABLE else "тип не для Office конвертации"
            att_info_suffix = f" - оригинал ({orig_fname_with_prefix}) сохранен (без конвертации: {reason})."
            current_y = _add_paragraph_to_pdf_util(pdf_canvas, f"<i>Оригинал сохранен: {orig_fname_with_prefix} (без конвертации: {reason})</i>", styles['N'], current_y, page_dims)

    return current_y, att_info_suffix, path_to_potential_pdf_attachment


def _process_email_attachments(msg, pdf_canvas, current_y, styles, font_to_use, page_dims, email_att_specific_path_base):
    """Обрабатывает все вложения в письме."""
    attachments_summary = []
    pdf_attachments_to_merge = []
    num_attachments_found = 0
    
    # Проверяем, есть ли вообще вложения, которые нужно сохранять на диск,
    # чтобы создать для них папку только при необходимости.
    has_disk_attachments = False
    for part_check in msg.walk():
        filename_header_check = part_check.get_filename()
        cdisp_check = str(part_check.get("Content-Disposition"))
        if filename_header_check or "attachment" in cdisp_check.lower():
            att_content_type_lower_check = part_check.get_content_type().lower()
            if not (att_content_type_lower_check.startswith("image/") or \
                    att_content_type_lower_check == "text/plain" or \
                    att_content_type_lower_check == "text/html"):
                has_disk_attachments = True
                break
    
    email_att_specific_path = None
    if has_disk_attachments:
        email_att_specific_path = email_att_specific_path_base # Используем ранее сформированный путь
        # Папка будет создана _handle_other_attachment при необходимости


    print(f"      DEBUG: Начало обработки вложений (всего частей в письме: {len(list(msg.walk()))}).")
    for part_counter, part in enumerate(msg.walk()):
        filename_header = part.get_filename()
        cdisp = str(part.get("Content-Disposition"))
        if not (filename_header or "attachment" in cdisp.lower()):
            continue

        num_attachments_found += 1
        attachment_data = part.get_payload(decode=True)
        if not attachment_data:
            attachments_summary.append(f"Вложение #{num_attachments_found}: Пустое (нет данных).")
            print(f"      Вложение #{num_attachments_found}: Пустое, пропуск.")
            continue

        decoded_fn = "".join([p.decode(c or 'utf-8', 'replace') if isinstance(p, bytes) else str(p) for p, c in decode_header(filename_header or "untitled_attachment")])
        sanitized_fn_for_saving = sanitize_filename(decoded_fn)
        att_info_line_base = f"Вложение #{num_attachments_found}: {decoded_fn} (Тип: {part.get_content_type()})"
        att_info_suffix = ""
        path_to_pdf_for_merge = None
        
        if current_y < page_dims['margin'] + styles['H'].fontSize + styles['H'].leading: # Проверка места для заголовка вложения
            pdf_canvas.showPage(); pdf_canvas.setFont(font_to_use, styles['N'].fontSize); current_y = page_dims['height'] - page_dims['margin']
        current_y = _add_paragraph_to_pdf_util(pdf_canvas, f"<b>Вложение: {decoded_fn}</b>", styles['H'], current_y, page_dims)
        
        att_content_type_lower = part.get_content_type().lower()
        try:
            print(f"        INFO: Обработка вложения #{num_attachments_found}: {decoded_fn} (Тип: {att_content_type_lower})")
            if att_content_type_lower.startswith("image/"):
                current_y, att_info_suffix = _handle_image_attachment(pdf_canvas, attachment_data, decoded_fn, current_y, styles, font_to_use, page_dims)
            elif att_content_type_lower == "text/plain":
                current_y, att_info_suffix = _handle_text_attachment(pdf_canvas, attachment_data, decoded_fn, current_y, styles, page_dims, part)
            elif att_content_type_lower == "text/html":
                current_y, att_info_suffix = _handle_text_attachment(pdf_canvas, attachment_data, decoded_fn, current_y, styles, page_dims, part, is_html=True)
            else: # Другие типы файлов - сохранение, конвертация
                current_y, att_info_suffix, path_to_pdf_for_merge = _handle_other_attachment(
                    pdf_canvas, attachment_data, decoded_fn, sanitized_fn_for_saving, part_counter, 
                    email_att_specific_path_base, # Передаем базовый путь, папка создастся внутри если надо
                    att_content_type_lower, current_y, styles, page_dims
                )
                if path_to_pdf_for_merge and os.path.exists(path_to_pdf_for_merge):
                    pdf_attachments_to_merge.append(path_to_pdf_for_merge)
                    print(f"          DEBUG: Добавлен в список для слияния: {path_to_pdf_for_merge}")

        except Exception as e_att_proc:
            print(f"      ERROR: Общая ошибка обработки вложения '{decoded_fn}': {e_att_proc}")
            att_info_suffix = f" - общая ошибка обработки ({e_att_proc})."
            current_y = _add_paragraph_to_pdf_util(pdf_canvas, f"<i>Общая ошибка обработки вложения '{decoded_fn}': {e_att_proc}</i>", styles['N'], current_y, page_dims)
            traceback.print_exc()
        
        attachments_summary.append(att_info_line_base + att_info_suffix)

    if num_attachments_found > 0:
        if current_y < page_dims['margin'] + (styles['H'].fontSize + styles['H'].leading) * (len(attachments_summary) + 1): # Примерная проверка места
            pdf_canvas.showPage(); pdf_canvas.setFont(font_to_use, styles['N'].fontSize); current_y = page_dims['height'] - page_dims['margin']
        current_y = _add_paragraph_to_pdf_util(pdf_canvas, f"<b>Всего вложений:</b> {num_attachments_found}", styles['H'], current_y, page_dims)
        for s_line in attachments_summary:
            current_y = _add_paragraph_to_pdf_util(pdf_canvas, f"- {s_line.replace('<','&lt;').replace('>','&gt;')}", styles['Body'], current_y, page_dims)
    else:
        current_y = _add_paragraph_to_pdf_util(pdf_canvas, "Вложения не найдены.", styles['N'], current_y, page_dims)
    
    # Возвращаем email_att_specific_path_base, так как он мог быть создан
    return num_attachments_found, pdf_attachments_to_merge, current_y, email_att_specific_path_base if has_disk_attachments and os.path.exists(email_att_specific_path_base) else None


def _finalize_pdf_report(pdf_canvas, temp_main_report_path, pdf_attachments_to_merge, journal_number):
    """Сохраняет основной PDF, объединяет с PDF-вложениями (если есть) и возвращает путь для просмотра."""
    pdf_canvas.save()
    print(f"    INFO: Временный основной PDF отчет ({os.path.basename(temp_main_report_path)}) сохранен.")

    path_for_review = temp_main_report_path
    temp_merged_path_for_cleanup = None

    if PYPDF_AVAILABLE and pdf_attachments_to_merge:
        print(f"      DEBUG: Начало слияния PDF. Основной отчет: {os.path.basename(temp_main_report_path)}. Вложения: {len(pdf_attachments_to_merge)}")
        timestamp_str_merge = datetime.datetime.now().strftime('%Y%m%d%H%M%S%f')
        temp_merged_filename = f"temp_merged_report_j{journal_number}_{timestamp_str_merge}.pdf"
        temp_merged_path = os.path.join(PDF_OUTPUT_DIRECTORY, temp_merged_filename)
        temp_merged_path_for_cleanup = temp_merged_path
        
        all_pdfs_to_combine = [temp_main_report_path] + pdf_attachments_to_merge
        
        if merge_pdfs(all_pdfs_to_combine, temp_merged_path):
            print(f"      INFO: Основной отчет и PDF-вложения успешно объединены в: {temp_merged_filename}")
            path_for_review = temp_merged_path
            try: # Удаляем первоначальный временный отчет, так как он слит
                os.remove(temp_main_report_path)
                print(f"        DEBUG: Удален первоначальный временный отчет '{os.path.basename(temp_main_report_path)}'.")
            except Exception as e_del_tmp_main:
                print(f"        WARNING: Не удалось удалить первоначальный временный отчет '{os.path.basename(temp_main_report_path)}' после слияния: {e_del_tmp_main}")
        else:
            print(f"      ERROR: Ошибка при слиянии PDF. Пользователю будет предложен только основной отчет.")
            temp_merged_path_for_cleanup = None # Нечего чистить, если слияние не удалось
    elif not pdf_attachments_to_merge:
        print(f"    INFO: PDF вложений для слияния нет. Используется основной отчет.")
    else: # pypdf недоступен
        print(f"    INFO: PDF вложения не будут объединены (pypdf недоступен). Используется основной отчет.")
    
    return path_for_review, temp_merged_path_for_cleanup


def _cleanup_temporary_files(*paths_to_clean):
    """Безопасно удаляет временные файлы."""
    for p in paths_to_clean:
        if p and os.path.exists(p):
            try:
                os.remove(p)
                print(f"    DEBUG: Временный файл '{os.path.basename(p)}' удален.")
            except Exception as e_clean:
                print(f"    WARNING: Не удалось удалить временный файл '{os.path.basename(p)}': {e_clean}")


# --- Основная функция обработки письма (рефакторинг) ---
def open_email(mail_obj, email_id, journal_number):
    """
    Обрабатывает одно электронное письмо: извлекает данные, создает PDF-отчет,
    обрабатывает вложения и запрашивает решение пользователя.
    """
    print(f"  DEBUG: Начало open_email для ID {email_id}, вх.№ {journal_number}")
    
    temp_main_report_path = None
    temp_merged_report_path = None # Путь для очистки, если было слияние
    final_attachments_folder_path = None # Путь к папке с сохраненными вложениями письма

    try:
        msg = _get_email_message_object(mail_obj, email_id)
        if not msg: return "ERROR", None

        email_headers = _extract_email_headers(msg)
        print(f"    Дата письма: {email_headers['formatted_date']}")
        print(f"    Отправитель: {email_headers['sender']}")
        print(f"    Получатели: {email_headers['recipients']}")
        print(f"    Тема письма: {email_headers['subject']}")

        email_body = _extract_email_body(msg)

        current_datetime = datetime.datetime.now()
        filename_date_str_for_display = current_datetime.strftime("%d.%m.%Y")
        filename_date_str_for_path = current_datetime.strftime("%d-%m-%Y") # Для имен папок

        timestamp_str = current_datetime.strftime('%Y%m%d%H%M%S%f')
        temp_main_report_filename = f"temp_main_report_j{journal_number}_{timestamp_str}.pdf"
        temp_main_report_path = os.path.join(PDF_OUTPUT_DIRECTORY, temp_main_report_filename)

        pdf_canvas, styles, font_to_use, page_dims, current_y = _setup_pdf_canvas_and_styles(temp_main_report_path)
        
        current_y = _add_email_headers_to_pdf(
            pdf_canvas, email_headers, email_body, journal_number, 
            filename_date_str_for_display, styles, current_y, page_dims
        )
        
        # Формируем базовое имя для папки вложений этого письма
        raw_email_att_dir_name = f"вх__{journal_number}_от_{filename_date_str_for_path}_attachments"
        # Путь, где будут лежать вложения, если они есть и сохраняются на диск
        email_att_specific_path_base = _create_attachment_directory(raw_email_att_dir_name)


        num_attachments, pdf_attachments_to_merge, current_y, final_attachments_folder_path = _process_email_attachments(
            msg, pdf_canvas, current_y, styles, font_to_use, page_dims, email_att_specific_path_base
        )

        path_for_review, temp_merged_report_path = _finalize_pdf_report(
            pdf_canvas, temp_main_report_path, pdf_attachments_to_merge, journal_number
        )

        if not os.path.exists(path_for_review):
            print(f"CRITICAL ERROR: Файл для просмотра '{path_for_review}' не существует перед вызовом handle_user_decision. Пропускаем.")
            # Попытка очистки, если что-то было создано
            _cleanup_temporary_files(temp_main_report_path, temp_merged_report_path)
            return "ERROR", None
            
        decision_status, final_file_path = handle_user_decision_for_pdf(
            path_for_review, journal_number, filename_date_str_for_display, 
            email_headers['subject'], final_attachments_folder_path # Передаем путь к папке вложений
        )

        # Очистка после решения пользователя
        # Если файл был сохранен/скачан, path_for_review уже переименован.
        # Если пропущен, handle_user_decision_for_pdf уже удалил path_for_review.
        # Если ошибка, path_for_review остался.
        # temp_main_report_path был удален в _finalize_pdf_report если слияние успешно.
        # Если слияния не было ИЛИ оно не удалось, И path_for_review это НЕ temp_main_report_path,
        # то temp_main_report_path мог остаться (маловероятно, но для полноты)
        
        if decision_status != "SAVED" and decision_status != "DOWNLOADED":
            # Если не сохранено и не скачано (т.е. SKIPPED или ERROR), то path_for_review (если он еще существует) должен быть удален,
            # если это был временный объединенный файл или исходный временный отчет (если не было слияния)
             if path_for_review == temp_merged_report_path and temp_merged_report_path: # Был слит и не переименован
                 _cleanup_temporary_files(temp_merged_report_path)
             elif path_for_review == temp_main_report_path and temp_main_report_path: # Не был слит (или слияние не удалось) и не переименован
                 _cleanup_temporary_files(temp_main_report_path)
        
        # Если handle_user_decision_for_pdf вернул "ERROR", path_for_review указывает на временный файл,
        # который нужно сохранить для отладки или ручного перемещения. Он не должен здесь удаляться.

        if final_attachments_folder_path and num_attachments > 0 and decision_status != "SKIPPED":
            print(f"    INFO: Файлы вложений (оригиналы/конвертированные/PDF) находятся в: {final_attachments_folder_path}")
        elif decision_status == "SKIPPED" and final_attachments_folder_path:
             # Папка вложений уже должна быть удалена handle_user_decision_for_pdf
             pass


        print(f"  DEBUG: Завершение open_email для ID {email_id} (вх.№ {journal_number}) со статусом: {decision_status}.")
        return decision_status, final_file_path

    except imaplib.IMAP4.error as e_imap:
        print(f"  ERROR: Ошибка IMAP4 (письмо ID {email_id}, вх.№ {journal_number}): {e_imap}")
        _cleanup_temporary_files(temp_main_report_path, temp_merged_report_path)
        return "ERROR", None
    except Exception as e_open_email:
        print(f"  ERROR: Критическая ошибка в open_email (письмо ID {email_id}, вх.№ {journal_number}): {e_open_email}")
        traceback.print_exc()
        _cleanup_temporary_files(temp_main_report_path, temp_merged_report_path)
        return "ERROR", None

# --- Функции process_scanned_pdfs и check_mailru_inbox (остаются без изменений из вашего исходного кода) ---
def process_scanned_pdfs(starting_journal_num_from_email=-1):
    """
    Обрабатывает PDF-файлы, помещенные пользователем в специальную папку.
    starting_journal_num_from_email: номер, с которого продолжать нумерацию после обработки почты.
                                     Если -1, значит, нумерация почты не начиналась, и нужно запросить заново.
    """
    scanned_folder_path = os.path.join(PDF_OUTPUT_DIRECTORY, SCANNED_INPUT_SUBDIR_NAME)
    registered_dir = os.path.join(PDF_OUTPUT_DIRECTORY, PDF_REGISTERED_SUBDIR_NAME)

    try:
        if not os.path.exists(scanned_folder_path):
            os.makedirs(scanned_folder_path)
            print(f"INFO: Создана папка для сканированных файлов: {scanned_folder_path}")
    except OSError as e:
        print(f"CRITICAL ERROR: Не удалось создать папку для сканированных файлов '{scanned_folder_path}': {e}. Работа прервана.")
        return

    print(f"\nПожалуйста, поместите отсканированные PDF-файлы в папку: \n{os.path.abspath(scanned_folder_path)}")
    input("После размещения файлов, нажмите Enter для продолжения...")

    try:
        pdf_files = [f for f in os.listdir(scanned_folder_path) if f.lower().endswith('.pdf')]
    except Exception as e:
        print(f"ERROR: Не удалось прочитать содержимое папки '{scanned_folder_path}': {e}")
        return

    if not pdf_files:
        print(f"INFO: Папка '{scanned_folder_path}' пуста или не содержит PDF-файлов. Завершение обработки сканов.")
        return

    print(f"INFO: Найдено {len(pdf_files)} PDF-файлов для регистрации.")
    pdf_files.sort() 

    current_journal_num_for_scans = -1
    if starting_journal_num_from_email != -1:
        current_journal_num_for_scans = starting_journal_num_from_email
        print(f"INFO: Продолжение нумерации с вх.№ {current_journal_num_for_scans} (после обработки почты).")
    else:
        print("INFO: Нумерация почты не была начата или не было обработано писем с присвоением номера.")
        current_journal_num_for_scans = prompt_for_starting_journal_number()
    
    date_str_for_filename = datetime.datetime.now().strftime("%d.%m.%Y")

    try:
        if not os.path.exists(registered_dir):
            os.makedirs(registered_dir)
            print(f"INFO: Создана папка для зарегистрированных файлов: {registered_dir}")
    except OSError as e:
        print(f"CRITICAL ERROR: Не удалось создать папку для зарегистрированных файлов '{registered_dir}': {e}. Работа прервана.")
        return
    
    processed_count = 0
    for original_filename in pdf_files:
        # Для сканов пока не делаем сложное имя с темой, т.к. OCR и sumy убраны для упрощения
        new_filename = f"вх.№ {current_journal_num_for_scans} от {date_str_for_filename}.pdf"
        old_filepath = os.path.join(scanned_folder_path, original_filename)
        new_filepath = os.path.join(registered_dir, new_filename)

        try:
            if os.path.exists(new_filepath):
                print(f"WARNING: Файл с именем '{new_filename}' уже существует в '{registered_dir}'. Будет перезаписан.")
                if os.path.abspath(old_filepath) != os.path.abspath(new_filepath): # Предотвращаем удаление самого себя
                    try:
                        os.remove(new_filepath)
                    except Exception as e_del_exist:
                        print(f"  WARNING: Не удалось удалить существующий файл '{new_filepath}' перед переименованием: {e_del_exist}")

            shutil.move(old_filepath, new_filepath) 
            print(f"INFO: Файл '{original_filename}' зарегистрирован и перемещен как '{new_filename}' в '{registered_dir}'")
            current_journal_num_for_scans += 1
            processed_count +=1
        except Exception as e:
            print(f"ERROR: Не удалось зарегистрировать файл '{original_filename}': {e}")
            traceback.print_exc()
    
    print(f"\nINFO: Завершена регистрация отсканированных файлов. Обработано: {processed_count} из {len(pdf_files)}.")


def check_mailru_inbox():
    load_dotenv()
    mail_host = os.getenv('IMAP_SERVER'); username = os.getenv('MAIL_RU_EMAIL'); password = os.getenv('MAIL_RU_PASSWORD')
    
    current_journal_num = -1 
    
    if not all([mail_host, username, password]): 
        print("Ошибка: Не все переменные окружения для почты определены. Пропуск обработки почты.")
        return current_journal_num 
    
    mail = None
    try:
        print(f"Подключение к {mail_host}..."); mail = imaplib.IMAP4_SSL(mail_host, 993)
        print("Вход в аккаунт..."); mail.login(username, password)
        mail.select('inbox'); print("Успешно подключено к 'Входящие'.")
        
        print("Поиск непрочитанных писем..."); status, data = mail.search(None, 'UNSEEN')
        if status != 'OK': 
            print(f"Ошибка поиска писем: {data[0].decode() if data and data[0] else 'Нет данных'}")
            return current_journal_num 
        
        email_ids_bytes = data[0].split(); num_unread = len(email_ids_bytes)
        print(f"У вас {num_unread} непрочитанных писем.")

        if num_unread > 0:
            current_journal_num = prompt_for_starting_journal_number() 

            for i, email_id_bytes in enumerate(email_ids_bytes):
                email_id_str = email_id_bytes.decode()
                print(f"\nОткрываем письмо #{i + 1} из {num_unread} (ID: {email_id_str}), предполагаемый вх.№ {current_journal_num}:")
                
                processing_result, _ = open_email(mail, email_id_str, current_journal_num) 
                
                if processing_result == "SAVED":
                    print(f"INFO: Письмо ID {email_id_str} успешно обработано и сохранено с вх.№ {current_journal_num}.")
                    current_journal_num += 1 
                elif processing_result == "DOWNLOADED":
                    print(f"INFO: Письмо ID {email_id_str} скачано без присвоения номера журнала.")
                elif processing_result == "SKIPPED":
                    print(f"INFO: Письмо ID {email_id_str} (предполагаемый вх.№ {current_journal_num}) пропущено пользователем (PDF и вложения удалены).")
                else: 
                    print(f"ERROR: Ошибка обработки письма ID {email_id_str} (предполагаемый вх.№ {current_journal_num}). Этот номер будет использован для следующего письма (если оно есть и будет обработано успешно).")
        else:
            print("У вас нет непрочитанных писем.")
            
    except imaplib.IMAP4.error as e: print(f"Ошибка IMAP4: {e}")
    except ConnectionRefusedError: print(f"Ошибка подключения к {mail_host}.")
    except Exception as e: print(f"Непредвиденная ошибка в check_mailru_inbox: {e}"); traceback.print_exc()
    finally:
        if mail and hasattr(mail, 'logout'):
            try: print("Выход из почтового сервера..."); mail.logout()
            except Exception as e_logout: print(f"Ошибка при выходе: {e_logout}")
        print("Завершение работы с почтовым сервером.")
    
    return current_journal_num


if __name__ == "__main__":
    if not os.path.exists(PDF_OUTPUT_DIRECTORY):
        try:
            os.makedirs(PDF_OUTPUT_DIRECTORY)
            print(f"INFO: Основная папка вывода '{PDF_OUTPUT_DIRECTORY}' создана.")
        except OSError as e:
            print(f"CRITICAL ERROR: Не удалось создать основную папку вывода '{PDF_OUTPUT_DIRECTORY}': {e}. Работа прервана.")
            sys.exit(1) 

    print("\n--- Этап 1: Обработка электронной почты ---")
    last_used_journal_num = check_mailru_inbox() 
    
    print("\n--- Этап 2: Регистрация отсканированных PDF-файлов ---")
    process_scanned_pdfs(last_used_journal_num) 
    
    print("\nЗавершение работы приложения.")
