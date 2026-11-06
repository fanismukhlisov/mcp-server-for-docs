# mcp-server-for-docs — MCP-сервер для PDF и Office-файлов

Универсальный MCP-сервер: читает текст и таблицы из **PDF, DOCX, XLSX, PPTX**,
извлекает встроенные картинки, рендерит страницы PDF в изображения (для
vision-моделей). Для PDF-чертежей — векторная геометрия с координатами,
слои (OCG), измерения, экспорт в DXF.

**Без внешних конвертеров:** не требует LibreOffice, MS Office или других
приложений — всё читается напрямую через Python-библиотеки.

## Установка

```bash
cd mcp-server-for-docs
python3 -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

Если библиотеки уже стоят глобально — venv не обязателен, можно ставить
`pip install -r requirements.txt` напрямую (см. раздел «Глобальная установка»).

## Подключение в LM Studio

В `mcp.json`:

```json
{
  "mcpServers": {
    "mcp-server-for-docs": {
      "command": "/полный/путь/к/mcp-server-for-docs/venv/bin/python",
      "args": ["/полный/путь/к/mcp-server-for-docs/server.py"],
      "env": {
        "PDF_MCP_ROOT": "C:\\Users\\ИмяПользователя\\Documents;C:\\aifolder"
      }
    }
  }
}
```

`PDF_MCP_ROOT` — список каталогов (через `;` или `:`), внутри которых
серверу разрешено читать файлы. Без переменной `PDF_MCP_ROOT` сервер по
умолчанию разрешает доступ к любому существующему файлу на диске (для
удобства) — если нужна строгая изоляция, обязательно задайте эту переменную.

## Инструменты сервера

### Навигация и состояние сессии

| Инструмент | Назначение |
|---|---|
| `open_document(file_path)` | закрепить файл (PDF/DOCX/XLSX/PPTX/TXT/MD) как активный — дальше остальные инструменты можно вызывать без пути |
| `open_pdf(pdf_path)` | алиас `open_document` (обратная совместимость) |
| `close_document()` / `close_pdf()` | снять закрепление |
| `list_documents(directory)` / `list_pdfs(directory)` | найти поддерживаемые файлы в каталоге |
| `get_document_info(file_path?)` / `get_pdf_info(pdf_path?)` | метаданные: страницы/абзацы/листы/слайды, таблицы, слои — под тип файла |

### Текст и таблицы (PDF + Office)

| Инструмент | Назначение |
|---|---|
| `read_text(file_path?, page_range?)` | текст: постранично (PDF), по абзацам (DOCX), по листам (XLSX), по слайдам (PPTX) |
| `search_text(query, file_path?, case_sensitive?)` | поиск текста в PDF/DOCX/PPTX с контекстом |
| `extract_tables(file_path?, page_range?)` | таблицы из PDF/DOCX/XLSX как структура |
| `extract_pptx_notes(file_path?)` | заметки докладчика из PPTX |
| `get_text_positions(page_number, pdf_path?, x0..y1?)` | слова PDF с координатами (bbox) — **только PDF** |

### Векторная геометрия, слои, измерения — **только PDF**

| Инструмент | Назначение |
|---|---|
| `summarize_vector_graphics(page_number, pdf_path?, x0..y1?)` | лёгкий обзор геометрии без координат — вызывать первым на сложных чертежах |
| `extract_vector_graphics(page_number, pdf_path?, x0..y1?, layers?, max_paths?)` | линии/прямоугольники/кривые с точными координатами |
| `list_layers(pdf_path?)` | список слоёв (OCG), если PDF экспортирован из AutoCAD с опцией слоёв |
| `extract_annotations(page_number?, pdf_path?)` | PDF-markup: комментарии, штампы, замечания рецензента |
| `calibrate_scale(x0,y0,x1,y1, real_length, unit, pdf_path?)` | задать масштаб по известному отрезку |
| `set_scale(points_per_unit, unit, pdf_path?)` | задать масштаб напрямую |
| `measure_distance(x0,y0,x1,y1, pdf_path?)` | расстояние в points и (если задан масштаб) в реальных единицах |
| `export_page_dxf(page_number, pdf_path?, out_path?, layers?, x0..y1?)` | экспорт геометрии в .dxf для AutoCAD |
| `export_page_svg(page_number, pdf_path?, out_path?)` | экспорт страницы в истинно векторный SVG |

### Визуальные и картинки

| Инструмент | Назначение |
|---|---|
| `render_page_image(page_number, pdf_path?, dpi?)` | рендер всей страницы PDF в PNG — **только PDF** |
| `render_region_image(page_number, x0,y0,x1,y1, pdf_path?, dpi?)` | рендер фрагмента страницы PDF с высоким DPI — **только PDF** |
| `extract_embedded_images(page_number?, file_path?, min_size_px?)` | встроенные картинки: со страницы (PDF) или из всего файла (DOCX/XLSX/PPTX) |

`file_path?`/`pdf_path?` — необязательный параметр: если не передать,
используется файл, закреплённый через `open_document`/`open_pdf`. Любой
явный путь автоматически закрепляется тоже, так что отдельный вызов
`open_document` перед другими инструментами не обязателен — достаточно
передать путь хотя бы раз.

## Типичные сценарии

### Работа с PDF-чертежами — см. подробные сценарии (измерение, перенос в
AutoCAD, проверка замечаний) ниже, в разделе «Для PDF-чертежей».

### Работа с Office-документами

1. `open_document("отчёт.docx")` — закрепить файл.
2. `read_text()` — весь текст по абзацам.
3. `extract_tables()` — таблицы как структура.
4. `extract_embedded_images()` — все картинки из документа (для DOCX/XLSX/PPTX
   читаются прямо из ZIP-контейнера файла, конвертером Office не нужен).

Для презентаций дополнительно: `extract_pptx_notes()` — заметки докладчика
под каждым слайдом.

### Для PDF-чертежей

1. `open_document("project2.pdf")` — закрепить файл один раз на сессию.
2. `get_document_info()` — страницы, слои, есть ли чертежи/картинки.
3. `render_page_image(1, dpi=150)` — общий обзор страницы vision-моделью.
4. `render_region_image(1, x0,y0,x1,y1, dpi=300-400)` — разглядеть деталь.
5. Измерение: `list_layers()` → `calibrate_scale(...)` по известному размеру
   → `measure_distance(...)` в реальных единицах.
6. Геометрия: **сначала** `summarize_vector_graphics(1)`, затем
   `extract_vector_graphics(1, layers=["Стены"])` точечно — не на всю
   страницу разом (см. «Переполнение контекста» ниже).
7. Перенос в AutoCAD: откалибровать масштаб → `export_page_dxf(1, layers=[...])`.
8. Проверка замечаний: `extract_annotations()` — markup-комментарии рецензента.

## Устранение проблем

### Сервер вообще не отвечает / ни один инструмент не работает

Самая частая причина — синтаксическая ошибка в `server.py` после ручной
правки (например, сбитый отступ). Сервер в этом случае не запускается
вообще, и хост (LM Studio) видит это как "инструмент не найден" или просто
тишину — визуально это легко спутать с проблемой чтения конкретного файла.
Проверить можно одной командой перед тем как ставить в `mcp.json`:

```bash
python3 -c "import ast; ast.parse(open('server.py').read()); print('OK')"
```

Если видите `SyntaxError` — сервер не запустится ни для одного файла, вне
зависимости от формата.

### Путь к файлу не находится

Для относительного пути (просто `"project.pdf"`) сервер ищет файл: (1) от
текущей рабочей директории процесса, (2) от домашней директории, (3) внутри
каждого каталога из `PDF_MCP_ROOT`, включая поиск по имени файла рекурсивно.
Если ничего не найдено — ошибка перечисляет, что было проверено. Если нашлось
несколько файлов с одинаковым именем — нужно уточнить полный путь.

### "Unterminated string in JSON" / битый ответ инструмента

На Windows поток stdout по умолчанию может быть не в UTF-8 — сервер
принудительно переключает `stdin`/`stdout`/`stderr` на UTF-8 при старте.
Кроме того, текст из PDF (особенно из CAD-экспорта со штампами) может
содержать "битые" управляющие символы — они вычищаются функцией
`_sanitize_text` перед отправкой хосту.

### "Cannot compact this conversation: compaction is not making progress"

Переполнение контекста модели большим ответом инструмента — типично для
`extract_vector_graphics`/`get_text_positions` на насыщенном чертеже без
фильтров. Решение:

1. Сначала `summarize_vector_graphics` — лёгкий обзор без координат.
2. Затем `extract_vector_graphics`/`get_text_positions` **точечно**: с
   `layers=[...]` и/или небольшой областью `x0..y1`, не на всю страницу.
3. Для небольших локальных моделей проверьте "Context Length" в настройках
   модели в LM Studio — если он слишком мал, даже отфильтрованные ответы
   могут не помещаться при накоплении за несколько вызовов подряд.
4. Если контекст уже переполнен — проще начать новый чат, чем продолжать
   забитую историю.

### Если модель забывает открыть файл первым

Любой инструмент с явным путём автоматически закрепляет файл, но если
модель вообще ни разу не передаёт путь — добавьте в системный промпт:

```
Перед любой работой с файлом сначала вызови open_document с ПОЛНЫМ путём.
Только после этого используй остальные инструменты.
```

### Глобальная установка без venv

Если библиотеки уже стоят в системном Python — можно указать в `mcp.json`
`"command": "python3"` вместо пути к venv. Проверить, что уже установлено:

```bash
python3 -c "
import importlib
for mod, pip_name in [('fitz','pymupdf'), ('pdfplumber','pdfplumber'),
                        ('PIL','Pillow'), ('mcp.server.fastmcp','mcp'),
                        ('docx','python-docx'), ('openpyxl','openpyxl'),
                        ('pptx','python-pptx'), ('ezdxf','ezdxf')]:
    try:
        importlib.import_module(mod)
        print(f'{pip_name:14s} -> УСТАНОВЛЕН')
    except ModuleNotFoundError:
        print(f'{pip_name:14s} -> нет')
"
```

Важна версия пакета `mcp` — сервер написан под API
`mcp.server.fastmcp.FastMCP` (SDK 1.x, `mcp<2` в requirements.txt).

## Ограничения

- CAD-инструменты (слои, векторная геометрия, измерения, экспорт DXF/SVG,
  рендер в изображение) работают **только с PDF** — для DOCX/XLSX/PPTX они
  вернут понятную ошибку с указанием, какой тип файла сейчас открыт.
- Отсканированные PDF без текстового слоя: `read_text` вернёт пусто —
  используйте `render_page_image`, чтобы vision-модель прочитала текст с картинки.
- `extract_embedded_images` для DOCX/XLSX/PPTX читает картинки напрямую из
  ZIP-контейнера файла (`word/media/`, `xl/media/`, `ppt/media/`) — без
  привязки к конкретной странице/слайду для DOCX/XLSX (для PPTX тоже нет
  привязки к конкретному слайду, картинки отдаются все сразу).
