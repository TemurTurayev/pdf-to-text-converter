import streamlit as st
import pytesseract
import pdf2image
import os
from docx import Document

# Конфигурация путей
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\user\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'
POPPLER_PATH = r'C:\Program Files\poppler-24.08.0\Library\bin'

def convert_pdf_to_text(pdf_file, start_page=None, end_page=None):
    # Конвертация PDF в изображения
    images = pdf2image.convert_from_bytes(
        pdf_file.read(),
        poppler_path=POPPLER_PATH,
        first_page=start_page,
        last_page=end_page
    )
    
    text = []
    total_pages = len(images)
    
    # Прогресс бар
    progress_bar = st.progress(0)
    
    for i, image in enumerate(images):
        # Извлечение текста из изображения
        page_text = pytesseract.image_to_string(image, lang='rus+eng')
        text.append(page_text)
        
        # Обновление прогресс бара
        progress = (i + 1) / total_pages
        progress_bar.progress(progress)
    
    return '\n\n'.join(text)

def save_as_docx(text, filename):
    doc = Document()
    doc.add_paragraph(text)
    doc.save(filename)

def main():
    st.title('PDF в Текст Конвертер')
    
    uploaded_file = st.file_uploader('Загрузите PDF файл', type=['pdf'])
    
    if uploaded_file is not None:
        # Получение количества страниц
        images = pdf2image.convert_from_bytes(
            uploaded_file.read(),
            poppler_path=POPPLER_PATH
        )
        total_pages = len(images)
        
        st.write(f'Всего страниц в документе: {total_pages}')
        
        # Выбор диапазона страниц
        convert_all = st.checkbox('Конвертировать весь документ', value=True)
        
        start_page = None
        end_page = None
        
        if not convert_all:
            col1, col2 = st.columns(2)
            with col1:
                start_page = st.number_input('Начальная страница', min_value=1, max_value=total_pages, value=1)
            with col2:
                end_page = st.number_input('Конечная страница', min_value=start_page, max_value=total_pages, value=total_pages)
        
        if st.button('Конвертировать'):
            uploaded_file.seek(0)
            text = convert_pdf_to_text(
                uploaded_file,
                start_page=start_page if not convert_all else None,
                end_page=end_page if not convert_all else None
            )
            
            st.text_area('Результат:', text, height=300)
            
            # Опции сохранения
            save_format = st.radio('Выберите формат для сохранения:', ['TXT', 'DOCX'])
            
            if st.button('Сохранить результат'):
                if save_format == 'TXT':
                    st.download_button(
                        label='Скачать как TXT',
                        data=text,
                        file_name='converted.txt',
                        mime='text/plain'
                    )
                else:
                    save_as_docx(text, 'converted.docx')
                    with open('converted.docx', 'rb') as f:
                        st.download_button(
                            label='Скачать как DOCX',
                            data=f,
                            file_name='converted.docx',
                            mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                        )

if __name__ == '__main__':
    main()