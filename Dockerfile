# syntax=docker/dockerfile:1

FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1

# Устанавливаем системные зависимости: LibreOffice (для VLM режима) и базовые шрифты
RUN apt-get update \
    && apt-get install -y --no-install-recommends \
       libreoffice \
       fonts-dejavu \
       ca-certificates \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Устанавливаем питон-зависимости
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

# Копируем код
COPY . .

EXPOSE 8501

# Запускаем Streamlit
CMD ["streamlit", "run", "streamlit_app_my.py", "--server.port=8501", "--server.address=0.0.0.0"] 