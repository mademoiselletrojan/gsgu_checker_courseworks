import streamlit as st
from docx import Document
import os
from datetime import datetime

from parsers.docx_parser import DocxParser
from analyzers.structure_checker import StructureChecker
from utils.profile_loader import ProfileLoader
from analyzers.fmfhi_format_checker import FMFHIFormatChecker


def render_error(index, section, font_size, line_spacing, in_table=False):
    """Красивый вывод ошибок форматирования"""
    if in_table:
        expected_font = "12–14 pt"
    else:
        expected_font = "14 pt"

    st.markdown(f"""
<div style="padding: 16px; border-left: 5px solid #dc3545; margin-bottom: 16px; background: linear-gradient(135deg, #fff5f5 0%, #ffffff 100%); border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
    <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 12px;">
        <span style="font-size: 24px;">❌</span>
        <span style="font-size: 17px; font-weight: 700; color: #c62828;">Ошибка #{index:02}</span>
        <span style="color: #666; font-size: 14px;">в разделе</span>
        <span style="background: #ffebee; padding: 4px 12px; border-radius: 20px; font-size: 14px; font-weight: 600; color: #c62828;">{section}</span>
    </div>
    <div style="margin-left: 32px;">
        <div style="margin-bottom: 8px;">
            <span style="font-size: 14px; color: #555;">📏 Размер шрифта:</span>
            <span style="font-weight: 600; color: #d32f2f; margin-left: 8px;">{font_size} pt</span>
            <span style="color: #888; font-size: 13px; margin-left: 8px;">(ожидалось {expected_font})</span>
        </div>
        <div>
            <span style="font-size: 14px; color: #555;">📐 Межстрочный интервал:</span>
            <span style="font-weight: 600; color: #d32f2f; margin-left: 8px;">{line_spacing}</span>
            <span style="color: #888; font-size: 13px; margin-left: 8px;">(ожидалось 1.5)</span>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)


# Настройка страницы
st.set_page_config(
    page_title="Проверка курсовой работы",
    page_icon="📘",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Кастомный CSS
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 2rem;
        border-radius: 12px;
        margin-bottom: 2rem;
        text-align: center;
    }
    .success-badge {
        background: #d4edda;
        color: #155724;
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 13px;
        font-weight: 600;
        display: inline-block;
    }
    .error-badge {
        background: #f8d7da;
        color: #721c24;
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 13px;
        font-weight: 600;
        display: inline-block;
    }

    /* 🔵 Синие кнопки */
    div.stButton > button:first-child {
        background-color: #1E88E5 !important;
        color: white !important;
        border-radius: 6px;
        height: 3rem;
        font-size: 16px;
        border: none;
    }
    div.stButton > button:hover {
        background-color: #1565C0 !important;
        color: white !important;
    }
</style>
""", unsafe_allow_html=True)

# Заголовок
st.markdown("""
<div class="main-header">
    <h1 style="color: white; margin: 0;">📘 Проверка курсовой работы</h1>
    <p style="color: rgba(255,255,255,0.9); margin: 10px 0 0 0;">Автоматическая проверка</p>
</div>
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "📎 Загрузите файл курсовой работы",
    type=["docx"],
    help="Поддерживаются только файлы в формате .docx"
)

if uploaded_file is not None:
    temp_path = "temp.docx"
    with open(temp_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    file_size = len(uploaded_file.getvalue()) / 1024
    st.success(f"✅ **Файл успешно загружен!**  \n📄 Имя: `{uploaded_file.name}`  \n💾 Размер: `{file_size:.1f} KB`")
    st.markdown("---")

    # ---------------------------------------------------------
    # 1) ПРОВЕРКА СТРУКТУРЫ
    # ---------------------------------------------------------
    st.markdown("### 📑 Этап 1: Проверка структуры документа")

    if st.button("🔍 Начать проверку структуры", use_container_width=True, type="primary"):
        with st.spinner("🔍 Анализирую структуру документа..."):
            parser = DocxParser(temp_path)
            found, paragraphs = parser.extract_key_headings()
            checker = StructureChecker(found)

            found_sections = [k for k, v in found.items() if v]
            missing_sections = [k for k, v in found.items() if not v and k in checker.REQUIRED]

        if found_sections:
            with st.expander("✅ Найденные разделы", expanded=False):
                for sec in found_sections:
                    st.markdown(f"<span class='success-badge'>✓</span> {sec.capitalize()}", unsafe_allow_html=True)

        if missing_sections:
            with st.expander("❌ Отсутствующие разделы", expanded=True):
                for sec in missing_sections:
                    st.markdown(f"<span class='error-badge'>✗</span> {sec.capitalize()}", unsafe_allow_html=True)
            st.error("❌ **Обнаружены отсутствующие обязательные разделы!**")
        elif found_sections:
            st.balloons()
            st.success("🎉 **Отлично!** Все обязательные разделы присутствуют в документе!")

        st.info(
            "💡 **Примечание:** Если раздел существует, но система его не обнаружила, "
            "проверьте корректность названия заголовка. Система распознаёт заголовки, "
            "начиная со слова «Введение»."
        )

    st.markdown("---")

    # ---------------------------------------------------------
    # 2) ПРОВЕРКА ОФОРМЛЕНИЯ
    # ---------------------------------------------------------
    st.markdown("### 📐 Этап 2: Проверка оформления по методичке")

    if st.button("📋 Проверить соответствие методичке", use_container_width=True, type="primary"):
        with st.spinner("📐 Анализирую форматирование документа..."):
            doc = Document(temp_path)
            profile = ProfileLoader.load_format_profile("fmfhi")
            checker = FMFHIFormatChecker(doc, profile)
            report = checker.run_checks()

        error_count = sum(1 for line in report if "Найдена ошибка" in line)

        st.markdown("### 📊 Результаты проверки оформления")

        if error_count == 0:
            st.markdown("""
            <div class="stat-card">
                <div style="font-size: 32px;">✅</div>
                <div style="font-size: 20px; font-weight: bold; color: #28a745;">Все параметры в норме</div>
                <div style="color: #666; margin-top: 10px;">Шрифт: Times New Roman | Размер: 14 pt (текст), 12–14 pt (таблицы) | Интервал: 1.5</div>
            </div>
            """, unsafe_allow_html=True)

        with st.expander("📄 Детальный отчёт проверки", expanded=True):
            for line in report:
                if line.startswith("📘"):
                    st.markdown(f"**{line}**")
                elif line.startswith(("╔", "║", "╠", "╚")):
                    st.code(line)
                elif line.startswith("="):
                    st.markdown("---")
                elif line.startswith("["):
                    st.text(line)
                else:
                    st.write(line)

        st.markdown("---")
        st.markdown("### 📥 Скачать отчёт проверки")

        txt_data = "\n".join(report)
        txt_path = "report.txt"
        pdf_path = "report.pdf"

        checker.save_report_as_txt(report, txt_path)
        checker.save_report_as_pdf(report, pdf_path)

        col1, col2 = st.columns(2)

        with col1:
            with open(txt_path, "r", encoding="utf-8") as f:
                st.download_button(
                    label="📄 Скачать отчёт (TXT)",
                    data=f.read(),
                    file_name=f"otchet_kursovaya_{datetime.now().strftime('%Y%m%d')}.txt",
                    mime="text/plain",
                    use_container_width=True
                )

        with col2:
            with open(pdf_path, "rb") as f:
                st.download_button(
                    label="📕 Скачать отчёт (PDF)",
                    data=f,
                    file_name=f"otchet_kursovaya_{datetime.now().strftime('%Y%m%d')}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )

    st.markdown("---")
    st.caption(
        "💡 **Рекомендация:** Используйте стили Word для заголовков — это повышает точность проверки."
    )

else:
    st.info("👆 **Загрузите файл** для начала проверки")
    
    st.markdown("""
    ### 📋 Что проверяет система:
    
    | Категория | Что проверяется |
    |-----------|----------------|
    | **📑 Структура** | Наличие введения, глав, заключения, списка литературы |
    | **📐 Оформление** | Шрифт: Times New Roman, размер: 14 pt (текст), 12–14 pt (таблицы), интервал: 1.5 |
    | **📏 Форматирование** | Выравнивание: по ширине (основной текст), по центру (заголовки) |
    
    > **Требования:** Левое поле — 30 мм, правое — 15 мм, верхнее/нижнее — 20 мм
    """)
