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

def resource_path(relative_path):
    """ Получить абсолютный путь к ресурсу, работает для разработки и для PyInstaller """
    try:
        # PyInstaller создает временную папку и сохраняет путь в _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".") # Для режима разработки

    return os.path.join(base_path, relative_path)
# Для слияния PDF
try:
    from pypdf import PdfWriter
    PYPDF_AVAILABLE = True
    print("INFO: Библиотека pypdf найдена, объединение PDF будет доступно.")
except ImportError:
    PYPDF_AVAILABLE = False
    print("ПРЕДУПРЕЖДЕНИЕ: Библиотека pypdf не найдена. PDF-вложения не будут объединены с основным отчетом.")


# Глобальная проверка доступности WIN32COM
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

DEJAVU_SANS_FONT_PATH = "DejaVuSans.ttf"
PDF_OUTPUT_DIRECTORY = "saved_emails_pdf"
ATTACHMENTS_SUBDIRECTORY_NAME = "attachments" # Базовая папка для всех вложений
CONVERTIBLE_EXTENSIONS = ['.doc','.docx','.rtf','.odt','.txt', '.xls','.xlsx','.ods', '.ppt','.pptx','.odp']
# JOURNAL_NUMBER_FILE - удалена

# --- Функция для запроса начального номера журнала ---
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
            print(f"INFO: Обработка писем в этой сессии начнется с вх.№ {start_next_num}.")
            return start_next_num
        except ValueError:
            print("ERROR: Пожалуйста, введите корректное число.")

# --- Функция слияния PDF ---
def merge_pdfs(list_of_pdf_paths, output_merged_pdf_path):
    if not PYPDF_AVAILABLE:
        print("    WARNING: Слияние PDF невозможно, т.к. библиотека pypdf не доступна.")
        return False

    merger = PdfWriter()
    merged_something = False
    try:
        for pdf_path in list_of_pdf_paths:
            if os.path.exists(pdf_path):
                try:
                    merger.append(pdf_path)
                    print(f"    DEBUG: Добавлен '{os.path.basename(pdf_path)}' для слияния.")
                    merged_something = True
                except Exception as e_append: # pypdf может выбросить исключение для поврежденных PDF
                    print(f"    WARNING: Не удалось добавить PDF '{os.path.basename(pdf_path)}' для слияния: {e_append}")
            else:
                print(f"    WARNING: PDF-файл для объединения не найден: {pdf_path}")

        if not merged_something:
            print("    WARNING: Нет PDF-файлов для фактического слияния.")
            merger.close()
            return False 

        with open(output_merged_pdf_path, "wb") as f_out:
            merger.write(f_out)
        print(f"    INFO: PDF-файлы успешно объединены в: {output_merged_pdf_path}")
        return True
    except Exception as e:
        print(f"    ERROR: Ошибка при объединении PDF-файлов: {e}")
        traceback.print_exc()
        return False
    finally:
        merger.close()

# --- Функции конвертации MS Office через COM ---
if WIN32COM_AVAILABLE:
    def convert_document_to_pdf_msword(input_path, output_path):
        word = None; doc = None
        input_path_abs = os.path.abspath(input_path); output_path_abs = os.path.abspath(output_path)
        try:
            word = win32com.client.Dispatch("Word.Application"); word.Visible = False
            doc = word.Documents.Open(input_path_abs, ReadOnly=True)
            wdFormatPDF = 17
            doc.SaveAs(output_path_abs, FileFormat=wdFormatPDF)
            print(f"    INFO: Файл '{os.path.basename(input_path_abs)}' успешно сконвертирован в PDF с помощью Word.")
            return True
        except Exception as e: print(f"    ERROR: Ошибка конвертации Word '{os.path.basename(input_path_abs)}': {e}"); return False
        finally:
            if doc: doc.Close(False)
            if word: word.Quit()
            del doc; del word # Можно оставить, но не обязательно

    def convert_spreadsheet_to_pdf_msexcel(input_path, output_path):
        excel = None; workbook = None
        input_path_abs = os.path.abspath(input_path); output_path_abs = os.path.abspath(output_path)
        try:
            excel = win32com.client.Dispatch("Excel.Application"); excel.Visible = False
            workbook = excel.Workbooks.Open(input_path_abs, ReadOnly=True)
            xlTypePDF = 0
            workbook.ExportAsFixedFormat(xlTypePDF, output_path_abs)
            print(f"    INFO: Файл '{os.path.basename(input_path_abs)}' успешно сконвертирован в PDF с помощью Excel.")
            return True
        except Exception as e: print(f"    ERROR: Ошибка конвертации Excel '{os.path.basename(input_path_abs)}': {e}"); return False
        finally:
            if workbook: workbook.Close(False)
            if excel: excel.Quit()
            del workbook; del excel

    def convert_presentation_to_pdf_msppt(input_path, output_path):
        powerpoint = None; presentation = None
        input_path_abs = os.path.abspath(input_path); output_path_abs = os.path.abspath(output_path)
        try:
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            presentation = powerpoint.Presentations.Open(input_path_abs, ReadOnly=True, WithWindow=False)
            ppSaveAsPDF = 32
            presentation.SaveAs(output_path_abs, FileFormat=ppSaveAsPDF)
            print(f"    INFO: Файл '{os.path.basename(input_path_abs)}' успешно сконвертирован в PDF с помощью PowerPoint.")
            return True
        except Exception as e: print(f"    ERROR: Ошибка конвертации PowerPoint '{os.path.basename(input_path_abs)}': {e}"); return False
        finally:
            if presentation: presentation.Close()
            if powerpoint: powerpoint.Quit()
            del presentation; del powerpoint

def sanitize_filename(filename):
    if not filename: return "untitled_attachment"
    filename = re.sub(r'[^\w\.\-\u0400-\u04FF]', '_', filename, flags=re.UNICODE)
    if filename.startswith('.'): filename = "_" + filename
    return filename if filename else "sanitized_attachment"

def check_mailru_inbox():
    load_dotenv()
    mail_host = os.getenv('IMAP_SERVER'); username = os.getenv('MAIL_RU_EMAIL'); password = os.getenv('MAIL_RU_PASSWORD')
    if not all([mail_host, username, password]): print("Ошибка: Не все переменные окружения определены."); return
    
    mail = None
    current_journal_num = -1 # Инициализация

    try:
        print(f"Подключение к {mail_host}..."); mail = imaplib.IMAP4_SSL(mail_host, 993)
        print("Вход в аккаунт..."); mail.login(username, password)
        mail.select('inbox'); print("Успешно подключено к 'Входящие'.")
        print("Поиск непрочитанных писем..."); status, data = mail.search(None, 'UNSEEN')
        if status != 'OK': print(f"Ошибка поиска писем: {data[0].decode() if data and data[0] else 'Нет данных'}"); return
        
        email_ids_bytes = data[0].split(); num_unread = len(email_ids_bytes)
        print(f"У вас {num_unread} непрочитанных писем.")

        if num_unread > 0:
            if not os.path.exists(PDF_OUTPUT_DIRECTORY):
                try: os.makedirs(PDF_OUTPUT_DIRECTORY); print(f"Папка '{PDF_OUTPUT_DIRECTORY}' создана.")
                except OSError as e: print(f"Ошибка создания папки '{PDF_OUTPUT_DIRECTORY}': {e}"); return
            
            current_journal_num = prompt_for_starting_journal_number() # Запрашиваем номер здесь, если есть письма

            for i, email_id_bytes in enumerate(email_ids_bytes):
                email_id_str = email_id_bytes.decode()
                print(f"\nОткрываем письмо #{i + 1} из {num_unread} (ID: {email_id_str}), будет присвоен вх.№ {current_journal_num}:")
                
                if open_email(mail, email_id_str, current_journal_num):
                    print(f"INFO: Письмо ID {email_id_str} успешно обработано и сохранено с вх.№ {current_journal_num}.")
                    current_journal_num += 1 # Инкрементируем для следующего письма в этой сессии
                else:
                    print(f"ERROR: Ошибка обработки письма ID {email_id_str} (предполагаемый вх.№ {current_journal_num}). Этот номер будет использован для следующего письма (если оно есть и будет обработано успешно).")
                    # Номер не инкрементируется, если письмо не удалось обработать,
                    # чтобы не было пропусков в нумерации при последующей успешной обработке.
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

def open_email(mail_obj, email_id, journal_number): 
    pdf_attachments_to_merge = []
    main_report_temp_path = None 
    parsed_date_for_filename = None 

    try:
        print(f"  DEBUG: Начало open_email для ID {email_id}, вх.№ {journal_number}")
        status, data = mail_obj.fetch(email_id, '(RFC822)')
        if status != 'OK': 
            print(f"  ERROR: Ошибка получения письма ID {email_id}"); 
            return False
        msg = email.message_from_bytes(data[0][1])

        date_header = msg.get('Date'); formatted_date_display = "Отсутствует" 
        filename_date_str_for_display = "ДАТА_НЕ_ОПРЕДЕЛЕНА" 
        filename_date_str_for_path = "ДАТА_НЕ_ОПРЕДЕЛЕНА"    

        if date_header:
            try: 
                parsed_date_for_filename = email.utils.parsedate_to_datetime(date_header)
                if parsed_date_for_filename:
                    formatted_date_display = parsed_date_for_filename.strftime("%Y-%m-%d %H:%M:%S")
                    filename_date_str_for_display = parsed_date_for_filename.strftime("%d.%m.%Y")
                    filename_date_str_for_path = parsed_date_for_filename.strftime("%d-%m-%Y") 
                else:
                    formatted_date_display = f"Не удалось распознать: {date_header}"
            except Exception as e_date: 
                formatted_date_display = f"Ошибка даты: {e_date}"
        print(f"  Дата письма: {formatted_date_display}")

        final_pdf_filename = f"вх.№ {journal_number} от {filename_date_str_for_display}.pdf"
        raw_email_att_dir_name = f"вх__{journal_number}_от_{filename_date_str_for_path}_attachments"
        
        from_header_raw = msg.get('From'); sender_info = "Отправитель отсутствует"
        if from_header_raw:
            name, addr = email.utils.parseaddr(from_header_raw)
            decoded_name_parts = [p.decode(c or 'utf-8', 'replace') if isinstance(p, bytes) else str(p) for p,c in decode_header(name)] if name else []
            sender_name_str = "".join(decoded_name_parts).strip(); sender_email_str = addr.strip()
            if sender_name_str and sender_email_str: sender_info = f"{sender_name_str} <{sender_email_str}>"
            elif sender_email_str: sender_info = sender_email_str
            elif sender_name_str: sender_info = sender_name_str
            else: sender_info = "".join([p.decode(c or 'utf-8', 'replace') if isinstance(p,bytes) else str(p) for p,c in decode_header(from_header_raw)]).strip() or from_header_raw
        print(f"  Отправитель: {sender_info}")

        to_header = msg.get('To'); recipients_str = "Получатели отсутствуют"
        if to_header: recipients_str = ', '.join([addr for _,addr in email.utils.getaddresses([to_header]) if addr]) or "Не удалось извлечь"
        print(f"  Получатели: {recipients_str}")

        subject_header = msg.get('Subject'); subject = "Тема отсутствует"
        if subject_header: subject = "".join([p.decode(c or 'utf-8', 'replace') if isinstance(p,bytes) else str(p) for p,c in decode_header(subject_header)])
        print(f"  Тема письма: {subject}")
        
        body = "Содержимое письма не найдено."
        plain_text_content, html_content_str = None, None
        if msg.is_multipart():
            print("    DEBUG: Письмо многочастное, ищем тело...")
            for part_counter_body, part in enumerate(msg.walk()):
                ctype = part.get_content_type(); cdisp = str(part.get("Content-Disposition"))
                if "attachment" in cdisp.lower() or part.get_filename(): continue
                if not (ctype.startswith("text/") or ctype.startswith("multipart/alternative")): continue
                if part.is_multipart(): continue 
                payload = part.get_payload(decode=True); charset = part.get_content_charset() or 'utf-8'
                if payload is None: continue
                dec_payload = ""
                try: dec_payload = payload.decode(charset, 'replace')
                except UnicodeDecodeError: dec_payload = payload.decode('latin-1', 'replace')
                except Exception as e_decode: print(f"        ERROR: Неожиданная ошибка декодирования тела: {e_decode}"); continue
                if ctype == "text/plain" and not plain_text_content: plain_text_content = dec_payload
                elif ctype == "text/html" and not html_content_str: html_content_str = dec_payload
            if plain_text_content: body = plain_text_content
            elif html_content_str:
                try: body_text_from_html = BeautifulSoup(html_content_str, "html.parser").get_text(separator='\n',strip=True); body = body_text_from_html if body_text_from_html else "HTML (без извлекаемого текста)."
                except Exception as e_bs: body = f"Ошибка BeautifulSoup (тело HTML): {e_bs}"
        else: 
            payload = msg.get_payload(decode=True); ctype = msg.get_content_type(); charset = msg.get_content_charset() or 'utf-8'
            if payload is None: body = "Содержимое отсутствует (payload is None)."
            else:
                dec_payload = ""
                try: dec_payload = payload.decode(charset, 'replace')
                except UnicodeDecodeError: dec_payload = payload.decode('latin-1', 'replace')
                except Exception as e_decode: dec_payload = f"Ошибка декодирования содержимого: {e_decode}"
                if ctype == "text/plain": body = dec_payload
                elif ctype == "text/html":
                    try: body_text_from_html = BeautifulSoup(dec_payload, "html.parser").get_text(separator='\n',strip=True); body = body_text_from_html if body_text_from_html else "HTML (без извлекаемого текста, не многочастное)."
                    except Exception as e_bs_s: body = f"Ошибка BeautifulSoup (тело HTML, не многочастное): {e_bs_s}"
                else: body = f"Содержимое письма имеет тип {ctype}, не является стандартным текстом."
        print(f"  Содержание (первые 200 симв.): {body[:200].replace(chr(10), ' ')}...")

        os.makedirs(PDF_OUTPUT_DIRECTORY, exist_ok=True)
        final_pdf_full_path = os.path.join(PDF_OUTPUT_DIRECTORY, final_pdf_filename)
        reportlab_output_path = final_pdf_full_path 

        c = canvas.Canvas(reportlab_output_path, pagesize=A4)
        width, height = A4; margin = 20*mm; current_y = height-margin; content_width = width-2*margin
        font_to_use = 'Helvetica'
        # DEJAVU_SANS_FONT_PATH = "DejaVuSans.ttf" # Имя файла шрифта
        font_path_to_register = resource_path(DEJAVU_SANS_FONT_PATH) # Используем resource_path

        if os.path.exists(font_path_to_register): # Проверяем путь, полученный от resource_path
            try: 
                pdfmetrics.registerFont(TTFont('DejaVuSans', font_path_to_register)) # Регистрируем шрифт по этому пути
                font_to_use = 'DejaVuSans'
            except Exception as e_font: # Добавил переменную для исключения, чтобы можно было вывести
                print(f"INFO: Ошибка регистрации шрифта '{font_path_to_register}': {e_font}. Используется Helvetica.")
        else: 
            print(f"INFO: Шрифт '{font_path_to_register}' не найден. Используется Helvetica.")
        styles = getSampleStyleSheet()
        styleN = styles['Normal']; styleN.fontName=font_to_use; styleN.fontSize=10; styleN.leading=12
        styleH = styles['Heading3']; styleH.fontName=font_to_use; styleH.fontSize=11; styleH.leading=13; styleH.spaceBefore=6; styleH.spaceAfter=2
        styleBody = styles['Normal']; styleBody.fontName=font_to_use; styleBody.fontSize=9; styleBody.leading=11

        def add_paragraph_pdf(text, style, y_pos, canvas_obj=c, max_h=height):
            p = Paragraph(text.replace('\n','<br/>'), style)
            p_w, p_h = p.wrapOn(canvas_obj, content_width, max_h)
            if y_pos - p_h < margin: canvas_obj.showPage(); canvas_obj.setFont(style.fontName, style.fontSize); y_pos = height-margin
            p.drawOn(canvas_obj, margin, y_pos - p_h)
            return y_pos - p_h - (style.leading * 0.2)

        current_y = add_paragraph_pdf(f"<b>Вх. №:</b> {journal_number} от {filename_date_str_for_display}", styleN, current_y)
        current_y = add_paragraph_pdf(f"<b>Дата получения (факт):</b> {formatted_date_display}", styleN, current_y)
        current_y = add_paragraph_pdf(f"<b>От:</b> {sender_info}", styleN, current_y)
        current_y = add_paragraph_pdf(f"<b>Кому:</b> {recipients_str}", styleN, current_y)
        current_y = add_paragraph_pdf(f"<b>Тема:</b> {subject}", styleN, current_y)
        current_y = add_paragraph_pdf(f"<b>Содержание:</b>", styleN, current_y)
        current_y = add_paragraph_pdf(body if body.strip() else "Содержимое отсутствует.", styleBody, current_y)

        attachments_summary = []; num_attachments_found = 0
        email_att_path_base = os.path.join(PDF_OUTPUT_DIRECTORY, ATTACHMENTS_SUBDIRECTORY_NAME)
        sanitized_email_att_subdir_name = sanitize_filename(raw_email_att_dir_name)
        email_att_specific_path = os.path.join(email_att_path_base, sanitized_email_att_subdir_name)

        if not os.path.exists(email_att_path_base):
            try: os.makedirs(email_att_path_base)
            except OSError as e_base_dir: print(f"    ERROR: Ошибка создания базовой папки вложений '{email_att_path_base}': {e_base_dir}")
        
        specific_email_folder_created_successfully = False
        if not os.path.exists(email_att_specific_path):
            try: os.makedirs(email_att_specific_path); specific_email_folder_created_successfully = True
            except OSError as e_mkdir: print(f"      ERROR: Не удалось создать папку '{email_att_specific_path}': {e_mkdir}.")
        else: specific_email_folder_created_successfully = True

        if msg.is_multipart():
            print(f"    DEBUG: Начало обработки вложений (всего частей в письме: {len(list(msg.walk()))}).")
            for part_counter, part in enumerate(msg.walk()):
                filename_header = part.get_filename()
                cdisp = str(part.get("Content-Disposition"))
                if not (filename_header or "attachment" in cdisp.lower()): continue

                num_attachments_found += 1
                attachment_data = part.get_payload(decode=True)
                if not attachment_data:
                    attachments_summary.append(f"Вложение #{num_attachments_found}: Пустое (нет данных).")
                    print(f"    Вложение #{num_attachments_found}: Пустое, пропуск.")
                    continue

                decoded_fn = "".join([p.decode(c or 'utf-8','replace') if isinstance(p,bytes) else str(p) for p,c in decode_header(filename_header or "untitled_attachment")])
                sanitized_fn_for_saving = sanitize_filename(decoded_fn) 
                att_info_line = f"Вложение #{num_attachments_found}: {decoded_fn} (Тип: {part.get_content_type()})"
                
                if current_y < margin + styleH.fontSize + styleH.leading: c.showPage(); c.setFont(font_to_use,styleN.fontSize); current_y = height-margin
                current_y = add_paragraph_pdf(f"<b>Вложение: {decoded_fn}</b>", styleH, current_y)
                
                att_content_type_lower = part.get_content_type().lower()
                
                try:
                    print(f"      INFO: Обработка вложения #{num_attachments_found}: {decoded_fn} (Тип: {att_content_type_lower})")
                    if att_content_type_lower.startswith("image/"):
                        img_data = BytesIO(attachment_data); img = Image(img_data)
                        img_w, img_h = img.drawWidth, img.drawHeight
                        aspect = img_h / float(img_w) if img_w > 0 else 1
                        disp_w = content_width; disp_h = disp_w * aspect
                        max_img_h = current_y - margin - 15*mm
                        if disp_h > max_img_h : disp_h = max(max_img_h, 10*mm); disp_w = disp_h / aspect if aspect > 1e-6 else content_width
                        if disp_w > content_width: disp_w = content_width; disp_h = disp_w * aspect
                        if current_y - disp_h < margin: c.showPage(); c.setFont(font_to_use,styleN.fontSize); current_y=height-margin; current_y=add_paragraph_pdf(f"<b>Вложение (продолжение): {decoded_fn}</b>", styleH, current_y)
                        if disp_w > 0 and disp_h > 0:
                            img.drawOn(c, margin, current_y - disp_h, width=disp_w, height=disp_h)
                            current_y -= (disp_h + 3*mm); att_info_line += " - встроено в PDF."
                        else: raise ValueError("Некорректные размеры изображения после расчета.")
                        print(f"      INFO: Изображение '{decoded_fn}' встроено.")
                    elif att_content_type_lower == "text/plain":
                        charset = part.get_content_charset() or 'utf-8'; text_att = attachment_data.decode(charset,'replace') if attachment_data else ""
                        current_y = add_paragraph_pdf(f"<u>Текстовое содержимое '{decoded_fn}':</u><br/>{text_att[:2000]}{'...' if len(text_att)>2000 else ''}", styleBody, current_y)
                        att_info_line += " - текст добавлен в PDF."
                        print(f"      INFO: Текст из '{decoded_fn}' добавлен в PDF.")
                    elif att_content_type_lower == "text/html":
                        charset = part.get_content_charset() or 'utf-8'; html_att_str = attachment_data.decode(charset,'replace') if attachment_data else ""
                        text_html = f"(Ошибка извлечения HTML: пустые данные)" if not html_att_str else \
                                    (BeautifulSoup(html_att_str,"html.parser").get_text(separator='\n',strip=True) or "(HTML без извлекаемого текста)")
                        current_y = add_paragraph_pdf(f"<u>HTML (как текст) '{decoded_fn}':</u><br/>{text_html[:2000]}{'...' if len(text_html)>2000 else ''}", styleBody, current_y)
                        att_info_line += " - HTML (как текст) добавлен в PDF."
                        print(f"      INFO: HTML из '{decoded_fn}' (как текст) добавлен в PDF.")
                    else: 
                        if not specific_email_folder_created_successfully:
                            att_info_line += " - не обработано (ошибка создания папки)."
                            print(f"      WARNING: Пропуск '{decoded_fn}' из-за ошибки папки.")
                            current_y = add_paragraph_pdf(f"<i>Файл '{decoded_fn}' не обработан (ошибка папки).</i>", styleN, current_y)
                        else:
                            orig_fname_with_prefix = f"{part_counter}_original_{sanitized_fn_for_saving}"
                            orig_path_abs = os.path.abspath(os.path.join(email_att_specific_path, orig_fname_with_prefix))
                            saved_original_successfully = False
                            try:
                                with open(orig_path_abs, "wb") as f_a: f_a.write(attachment_data)
                                print(f"      INFO: Оригинал '{decoded_fn}' сохранен как '{orig_fname_with_prefix}'")
                                saved_original_successfully = True
                            except Exception as e_save_o:
                                print(f"      ERROR: Не удалось сохранить оригинал '{decoded_fn}': {e_save_o}")
                                att_info_line += f" - ошибка сохранения оригинала ({e_save_o})."
                                current_y = add_paragraph_pdf(f"<i>Ошибка сохранения оригинала '{decoded_fn}': {e_save_o}</i>", styleN, current_y)

                            if saved_original_successfully:
                                file_ext = os.path.splitext(sanitized_fn_for_saving)[1].lower()
                                is_pdf_already = (att_content_type_lower == "application/pdf" or file_ext == ".pdf")
                                path_to_potential_pdf_attachment = None

                                if is_pdf_already:
                                    final_pdf_name_att = f"{part_counter}_{sanitized_fn_for_saving}" 
                                    final_pdf_path_abs_att = os.path.abspath(os.path.join(email_att_specific_path, final_pdf_name_att))
                                    try:
                                        os.rename(orig_path_abs, final_pdf_path_abs_att)
                                        path_to_potential_pdf_attachment = final_pdf_path_abs_att
                                        print(f"      INFO: PDF-вложение '{decoded_fn}' сохранено как '{final_pdf_name_att}'.")
                                        att_info_line += f" - PDF-вложение будет объединено."
                                        current_y = add_paragraph_pdf(f"<i>PDF-вложение '{decoded_fn}' будет добавлено в конец этого документа.</i>", styleN, current_y)
                                    except Exception as e_rename:
                                        print(f"      WARNING: Не удалось переименовать '{orig_fname_with_prefix}': {e_rename}. Используется как есть для объединения.")
                                        path_to_potential_pdf_attachment = orig_path_abs 
                                        att_info_line += f" - PDF-вложение ({orig_fname_with_prefix}) будет объединено."
                                        current_y = add_paragraph_pdf(f"<i>PDF-вложение '{decoded_fn}' ({orig_fname_with_prefix}) будет добавлено в конец этого документа.</i>", styleN, current_y)
                                
                                elif WIN32COM_AVAILABLE and file_ext in CONVERTIBLE_EXTENSIONS:
                                    conv_pdf_name = f"{part_counter}_converted_{os.path.splitext(sanitized_fn_for_saving)[0]}.pdf"
                                    conv_pdf_path_abs = os.path.abspath(os.path.join(email_att_specific_path, conv_pdf_name))
                                    com_func = None
                                    if file_ext in ['.doc','.docx','.rtf','.odt','.txt']: com_func = convert_document_to_pdf_msword
                                    elif file_ext in ['.xls','.xlsx','.ods']: com_func = convert_spreadsheet_to_pdf_msexcel
                                    elif file_ext in ['.ppt','.pptx','.odp']: com_func = convert_presentation_to_pdf_msppt

                                    if com_func and com_func(orig_path_abs, conv_pdf_path_abs):
                                        path_to_potential_pdf_attachment = conv_pdf_path_abs
                                        att_info_line += f" - сконвертировано в PDF, будет объединено."
                                        current_y = add_paragraph_pdf(f"<i>Файл '{decoded_fn}' сконвертирован в PDF ({conv_pdf_name}) и будет добавлен в конец этого документа.</i>", styleN, current_y)
                                        print(f"      INFO: Вложение '{decoded_fn}' сконвертировано в '{conv_pdf_name}'.")
                                        try: os.remove(orig_path_abs); print(f"      INFO: Оригинал '{orig_fname_with_prefix}' удален.")
                                        except Exception as e_del_orig: print(f"      WARNING: Не удалось удалить оригинал '{orig_fname_with_prefix}': {e_del_orig}")
                                    else: 
                                        att_info_line += f" - оригинал ({orig_fname_with_prefix}) сохранен, конвертация не удалась."
                                        current_y = add_paragraph_pdf(f"<i>Оригинал сохранен: {orig_fname_with_prefix}. Конвертация в PDF не удалась.</i>", styleN, current_y)
                                        print(f"      INFO: Оригинал '{decoded_fn}' ({orig_fname_with_prefix}) сохранен. Конвертация не удалась.")
                                else: 
                                    reason = "не Windows/pywin32" if not WIN32COM_AVAILABLE else "тип не для Office конвертации"
                                    att_info_line += f" - оригинал ({orig_fname_with_prefix}) сохранен (без конвертации: {reason})."
                                    current_y = add_paragraph_pdf(f"<i>Оригинал сохранен: {orig_fname_with_prefix} (без конвертации: {reason})</i>", styleN, current_y)
                                    print(f"      INFO: Оригинал '{decoded_fn}' ({orig_fname_with_prefix}) сохранен ({reason}).")
                                
                                if path_to_potential_pdf_attachment and os.path.exists(path_to_potential_pdf_attachment):
                                    pdf_attachments_to_merge.append(path_to_potential_pdf_attachment)
                                    print(f"      DEBUG: Добавлен в список для слияния: {path_to_potential_pdf_attachment}")
                except Exception as e_att_proc:
                    print(f"    ERROR: Общая ошибка обработки вложения '{decoded_fn}': {e_att_proc}")
                    att_info_line += f" - общая ошибка обработки ({e_att_proc})."
                    current_y = add_paragraph_pdf(f"<i>Общая ошибка обработки вложения '{decoded_fn}': {e_att_proc}</i>", styleN, current_y)
                    traceback.print_exc()
                attachments_summary.append(att_info_line)
        
        if num_attachments_found > 0:
            if current_y < margin + (styleH.fontSize+styleH.leading)*(len(attachments_summary)+1): c.showPage(); c.setFont(font_to_use,styleN.fontSize); current_y=height-margin
            current_y = add_paragraph_pdf(f"<b>Всего вложений:</b> {num_attachments_found}", styleH, current_y)
            for s_line in attachments_summary: current_y = add_paragraph_pdf(f"- {s_line.replace('<','&lt;').replace('>','&gt;')}", styleBody, current_y)
        else: current_y = add_paragraph_pdf("Вложения не найдены.", styleN, current_y)

        c.save() 
        print(f"INFO: Основной PDF отчет ({os.path.basename(reportlab_output_path)}) сохранен.")

        if PYPDF_AVAILABLE and pdf_attachments_to_merge:
            print(f"    DEBUG: Начало слияния PDF. Основной отчет: {reportlab_output_path}. Вложения для слияния: {len(pdf_attachments_to_merge)}")
            main_report_temp_path = reportlab_output_path + ".tmp_main_report.pdf" 
            try:
                os.rename(reportlab_output_path, main_report_temp_path) 
                print(f"      DEBUG: Основной отчет временно переименован в: {main_report_temp_path}")
                all_pdfs_to_combine = [main_report_temp_path] + pdf_attachments_to_merge
                
                if merge_pdfs(all_pdfs_to_combine, final_pdf_full_path): 
                    print(f"    INFO: Все PDF успешно объединены в: {final_pdf_full_path}")
                    try:
                        os.remove(main_report_temp_path)
                        print(f"      DEBUG: Временный основной отчет '{main_report_temp_path}' удален.")
                    except Exception as e_del_tmp:
                        print(f"      WARNING: Не удалось удалить временный основной отчет '{main_report_temp_path}': {e_del_tmp}")
                else: 
                    print(f"    ERROR: Ошибка при слиянии PDF. Восстанавливаем основной отчет из временного файла.")
                    try: 
                        os.rename(main_report_temp_path, final_pdf_full_path) 
                        print(f"      INFO: Основной отчет восстановлен как '{final_pdf_full_path}'. PDF вложения не были объединены.")
                    except Exception as e_rename_back:
                        print(f"      CRITICAL ERROR: Не удалось восстановить основной отчет из '{main_report_temp_path}' в '{final_pdf_full_path}': {e_rename_back}")
                        print(f"                      Основной отчет может быть доступен как: {main_report_temp_path}")
                        return False 
            except Exception as e_rename_main:
                print(f"    ERROR: Не удалось переименовать основной отчет для слияния: {e_rename_main}")
                print(f"             Слияние PDF не будет выполнено. Основной отчет сохранен как: {reportlab_output_path}")
        elif not pdf_attachments_to_merge:
            print(f"INFO: PDF вложений для слияния нет. Основной отчет сохранен как: {final_pdf_full_path}")
        else: 
             print(f"INFO: PDF вложения не будут объединены (pypdf недоступен). Основной отчет: {final_pdf_full_path}")

        if specific_email_folder_created_successfully and num_attachments_found > 0 :
             print(f"    INFO: Файлы вложений (оригиналы/конвертированные/PDF) находятся в: {email_att_specific_path}")
        
        print(f"  DEBUG: Завершение open_email для ID {email_id} (вх.№ {journal_number}) успешно.")
        return True 

    except imaplib.IMAP4.error as e_imap: 
        print(f"  ERROR: Ошибка IMAP4 (письмо ID {email_id}, вх.№ {journal_number}): {e_imap}")
        return False
    except Exception as e_open_email:
        print(f"  ERROR: Критическая ошибка в open_email (письмо ID {email_id}, вх.№ {journal_number}): {e_open_email}")
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("Запуск приложения для проверки почты...")
    check_mailru_inbox()