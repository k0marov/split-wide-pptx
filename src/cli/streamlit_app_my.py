import streamlit as st
import tempfile
import os
from pathlib import Path

from src.renote.processor import process_pptx, ProcessingOptions


def main():
    st.set_page_config(
        page_title="Renote - Адаптация презентаций",
        page_icon="📊",
        layout="wide"
    )
    
    st.title("📊 Renote - Адаптация широкоформатных презентаций")
    st.markdown("""
    Загрузите широкоформатную презентацию (.pptx), и система автоматически адаптирует её к формату 16:9, 
    разбив широкие слайды на несколько узких с сохранением фона и стилей.
    """)
    
    # Боковая панель с настройками
    with st.sidebar:
        st.header("⚙️ Настройки")
        
        # Основные параметры
        title_min_font = st.slider(
            "Минимальный размер шрифта титула (pt)", 
            min_value=20.0, 
            max_value=60.0, 
            value=36.0,
            help="Минимальный размер шрифта для определения титульного слайда"
        )
        
        title_min_width_ratio = st.slider(
            "Мин. ширина shape к ширине трети", 
            min_value=1.0, 
            max_value=2.0, 
            value=1.2,
            help="Минимальное соотношение ширины элемента к ширине трети слайда"
        )
        

    # Основная область
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.header("📁 Загрузка файла")
        uploaded_file = st.file_uploader(
            "Выберите PPTX файл для обработки",
            type=['pptx'],
            help="Поддерживаются только файлы формата .pptx"
        )
        
        if uploaded_file is not None:
            st.success(f"✅ Файл загружен: {uploaded_file.name}")
            st.info(f"📊 Размер файла: {len(uploaded_file.getvalue()) / 1024:.1f} KB")
    
    with col2:
        st.header("🔄 Обработка")
        
        if uploaded_file is not None:
            if st.button("🚀 Обработать презентацию", type="primary", use_container_width=True):
                try:
                    # Создаем временные файлы
                    with tempfile.TemporaryDirectory() as temp_dir:
                        # Сохраняем загруженный файл
                        input_path = os.path.join(temp_dir, "input.pptx")
                        with open(input_path, "wb") as f:
                            f.write(uploaded_file.getvalue())
                        
                        # Путь для выходного файла
                        output_path = os.path.join(temp_dir, "output.pptx")
                        
                        # Настройки обработки
                        options = ProcessingOptions(
                            title_min_font_pt=title_min_font,
                            title_min_width_ratio=title_min_width_ratio,
                        )
                        
                        # Прогресс бар
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                        status_text.text("🔍 Анализ слайдов...")
                        progress_bar.progress(25)
                        
                        # Обработка файла
                        process_pptx(
                            input_path,
                            output_path,
                            options=options,
                            direct=True
                        )
                        
                        progress_bar.progress(100)
                        status_text.text("✅ Обработка завершена!")
                        
                        # Чтение обработанного файла
                        with open(output_path, "rb") as f:
                            processed_file = f.read()
                        
                        # Кнопка для скачивания
                        original_name = Path(uploaded_file.name).stem
                        download_name = f"{original_name}_renote.pptx"
                        
                        st.download_button(
                            label="💾 Скачать обработанную презентацию",
                            data=processed_file,
                            file_name=download_name,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            type="primary",
                            use_container_width=True
                        )
                        
                        st.success("🎉 Презентация успешно обработана и готова к скачиванию!")
                        
                except Exception as e:
                    st.error(f"❌ Ошибка при обработке файла: {str(e)}")
                    st.exception(e)
        else:
            st.info("👆 Сначала загрузите PPTX файл")
    
    # Дополнительная информация
    with st.expander("ℹ️ О программе"):
        st.markdown("""
        **Renote** - инструмент для адаптации широкоформатных презентаций к формату 16:9.
        
        **Возможности:**
        - 🔄 Автоматическое разбиение широких слайдов на узкие (16:9)
        - 🎨 Сохранение фона и стилей оформления
        - 📝 Интеллектуальная обработка титульных слайдов
        """)
    
    # Информация о статусе
    st.markdown("---")
    col_info1, col_info2, col_info3 = st.columns(3)
    
    with col_info3:
        st.metric("Размер шрифта титула", f"{title_min_font:.0f}pt")


if __name__ == "__main__":
    main() 