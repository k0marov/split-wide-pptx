# Renote

Инструмент для адаптации широкоформатных презентаций (.pptx) к формату 16:9 путём интеллектуального разбиения каждого широкого слайда на три узких, с сохранением фона/темы и улучшенной обработкой титульных.

## Возможности
- Прямой режим без предварительной трипликации: из каждого слайда формируются 1 (титул) или 3 (обычные) узких слайда.
- Классификация сценария (title | split):
  - Heuristic: по ширине шейпов и размерам шрифтов.
  - VLM (Ollama + LangChain): классификация по изображению слайда.
- Сохранение фона/темы, копирование стилей текста на титульных.

## Установка
1. Установите зависимости (используя uv):

```
uv pip install -r requirements.txt
```

2. (Опционально, для VLM):
   - Установите LibreOffice для экспорта слайдов в PNG (команда `soffice`).
   - Установите и запустите Ollama (`ollama serve`), скачайте нужную модель (например, `ollama pull gemma3:4b` или `ollama pull llava:latest`).

## Запуск
- Эвристика (без VLM):
```
uv run renote_cli.py \
  "/путь/к/входу.pptx" \
  "/путь/к/выходу.pptx"
```

- С VLM (Ollama + LibreOffice):
```
uv run renote_cli.py \
  --vlm ollama \
  --ollama-model gemma3:4b \
  --ollama-url http://localhost:11434 \
  --soffice /Applications/LibreOffice.app/Contents/MacOS/soffice \
  "/путь/к/входу.pptx" \
  "/путь/к/выходу.pptx"
```

## Тесты (ручной прогон)
- Сложите входные презентации в `tests/inputs` и запустите:
```
for f in tests/inputs/*.pptx; do 
  base=$(basename "$f" .pptx)
  uv run renote_cli.py "$f" "tests/outputs/${base}_out.pptx"
done
```
(Для fish используйте синтаксис `set base (basename "$f" .pptx)`).

## Ограничения
- На границах третей изображения не «режутся», а либо удаляются, либо смещаются (crop можно добавить отдельно для `PictureShape`).
- Сложные эффекты WordArt/SmartArt могут копироваться частично.

## Структура
- `renote/detectors.py` — эвристики и выбор основного текстового блока.
- `renote/transforms.py` — основной алгоритм преобразования (direct), обработка титульных.
- `renote/processor.py` — сценарный обработчик и параметры.
- `renote/vlm.py` — VLM-классификация (heuristic/ollama + LangChain), экспорт PNG.
- `renote_cli.py` — CLI-интерфейс.
