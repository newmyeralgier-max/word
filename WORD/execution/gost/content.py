"""Трансформации содержания документа.

Этот модуль — не про внешний вид, а про сам текст:

  * `fix_task_dates`   — исправление опечаток в датах ЗАДАНИЯ
                         («15.01. 20252026 г.» → «15.01.2026 г.»).
  * `remove_manual_toc`— снос рукописного СОДЕРЖАНИЯ, обёрнутого в
                         `<w:sdt>`-контейнер.
  * `replace_intro`    — переписать блок ВВЕДЕНИЕ полноценным
                         ГОСТ-шаблоном (актуальность / цель / задачи /
                         объект / предмет / методы / значимость).
  * `insert_referat`   — вставить блок РЕФЕРАТ после title zone (ГОСТ 7.32).
  * `remove_blank_tail`— вычистить лишние пустые параграфы в хвосте,
                         чтобы не было 7–10 пустых страниц после ПРИЛОЖЕНИЙ.

Все функции возвращают счётчики/bool для статистики.
"""
from __future__ import annotations

import re
from copy import deepcopy

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

from . import config as cfg
from .utils import apply_p_format, set_run_font


# ---------------------------------------------------------------------------
# 1. Даты в ЗАДАНИИ
# ---------------------------------------------------------------------------

_DATE_REPLACEMENTS: tuple[tuple[re.Pattern[str], str], ...] = (
    # «20252026» или «202620252026» → «2026»
    (re.compile(r'\b(?:2025)?2026(?:2025)?2026\b'), '2026'),
    # «15. 01. 2026» → «15.01.2026» (убрать пробелы внутри даты)
    (re.compile(r'\b(\d{1,2})\.\s*(\d{1,2})\.\s*(20\d\d)\b'), r'\1.\2.\3'),
    # «1.05. 0» / «15.05. 0» / «1.06. 0» → «01.05.2026» и т. п.
    (re.compile(r'\b(\d{1,2})\.(\d{1,2})\.\s*0\b'), lambda m:
        f'{int(m.group(1)):02d}.{int(m.group(2)):02d}.2026'),
    # «1.ТЕМА» → «1. ТЕМА» (один пробел после цифры)
    (re.compile(r'\b(1)\.\s*Тема\s+работы\s*:\s*:?\s*'), r'\1. Тема работы: '),
)


def fix_task_dates(doc) -> int:
    """Починить опечатки в датах/форматировании ЗАДАНИЯ.

    Обходим ВСЕ `w:t` в теле документа (в т.ч. внутри таблиц и sdt),
    и применяем серию regex-замен. Возвращаем число изменённых `<w:t>`.
    """
    changed = 0
    for t in doc.element.iter(qn('w:t')):
        if not t.text:
            continue
        new = t.text
        for pat, repl in _DATE_REPLACEMENTS:
            new = pat.sub(repl, new)
        # «кВ. Котова» → «кв. Котова» (квартал, а не киловольт)
        new = re.sub(r'\bкВ\.\s*Котова', 'кв. Котова', new)
        # Лишняя ёлочка «» вокруг темы: «» Реконструкция …» ""» → нормализовать
        new = re.sub(r'\s{2,}»', '»', new)
        if new != t.text:
            t.text = new
            changed += 1
    return changed


# ---------------------------------------------------------------------------
# 2. Рукописное СОДЕРЖАНИЕ в sdt → удалить
# ---------------------------------------------------------------------------

def remove_manual_toc(doc) -> int:
    """Удалить все `<w:sdt>`-обёртки на уровне body, содержащие СОДЕРЖАНИЕ.

    В БР.docx рукописное содержание обёрнуто в один Structured Document
    Tag (SDT) с первым параграфом «СОДЕРЖАНИЕ» и tab-leader-ами. Мы его
    целиком убираем — вместо этого пайплайн вставит нормальное TOC-поле
    через `toc.insert_toc_before`.

    Возвращает количество удалённых sdt-блоков.
    """
    body = doc.element.body
    removed = 0
    for sdt in list(body.findall(qn('w:sdt'))):
        # Первый текст внутри sdt
        first_p = sdt.find('.//' + qn('w:p'))
        if first_p is None:
            continue
        txt_parts = first_p.findall('.//' + qn('w:t'))
        first_txt = ''.join((x.text or '') for x in txt_parts).strip().upper()
        if first_txt in ('СОДЕРЖАНИЕ', 'ОГЛАВЛЕНИЕ'):
            body.remove(sdt)
            removed += 1

    # Подчистить висячие параграфы со стилем `TOC*` (например,
    # `TOC Heading`) — иначе `has_toc` думает, что оглавление уже есть.
    # В Word стиль может ссылаться через styleId (например, `af8`),
    # а не через имя. Делаем reverse-map: id → name.
    styles_el = doc.styles.element
    id_to_name = {}
    for style in styles_el.findall(qn('w:style')):
        sid = style.get(qn('w:styleId')) or ''
        name_el = style.find(qn('w:name'))
        name = name_el.get(qn('w:val')) if name_el is not None else ''
        id_to_name[sid] = name

    for p in list(body.findall(qn('w:p'))):
        pStyle = p.find(qn('w:pPr') + '/' + qn('w:pStyle'))
        if pStyle is None:
            continue
        sid = (pStyle.get(qn('w:val')) or '')
        name = id_to_name.get(sid, '')
        # Совпадение по ID или по резолвнутому имени
        if (sid.lower().startswith('toc') or name.lower().startswith('toc')):
            if not _p_text_all(p) and not p.findall('.//' + qn('w:drawing')):
                body.remove(p)
                removed += 1
    return removed


# ---------------------------------------------------------------------------
# 3. Переписать ВВЕДЕНИЕ
# ---------------------------------------------------------------------------

_INTRO_BODY: tuple[str, ...] = (
    # Актуальность
    'Актуальность темы выпускной квалификационной работы обусловлена '
    'необходимостью реконструкции распределительных сетей жилого '
    'микрорайона в современных условиях, когда обеспечение надёжного, '
    'качественного и экономически обоснованного электроснабжения '
    'потребителей является одной из ключевых задач развития '
    'инфраструктуры городов. Рост удельных электрических нагрузок, '
    'повышение требований к показателям качества электроэнергии, а '
    'также физический и моральный износ существующего оборудования '
    'делают задачу проектирования новых и реконструкции действующих '
    'сетей 6/0,4 кВ особенно значимой.',

    # Цель
    'Целью работы является разработка проекта реконструкции '
    'электроснабжения жилого квартала Котова в г. Краснодоне, '
    'обеспечивающего требуемую надёжность, качество электроэнергии и '
    'экономическую эффективность при выполнении действующих '
    'нормативно-технических требований (ПУЭ, ГОСТ 32144-2013, '
    'ГОСТ Р 50571, СП 31-110-2003).',

    # Задачи (списком — сохраняем в одном абзаце, заголовки выделим курсивом)
    'Для достижения поставленной цели в работе решаются следующие задачи:',
    '1) анализ характеристик объекта электроснабжения и состава его '
    'потребителей; определение категорий надёжности электроприёмников;',
    '2) расчёт электрических нагрузок на вводах жилых и общественных '
    'зданий, освещения микрорайона и суммарной расчётной нагрузки;',
    '3) выбор числа, мощности и типа трансформаторных подстанций 6/0,4 кВ '
    'исходя из расчётных нагрузок и категорий надёжности;',
    '4) выбор схемы внутреннего электроснабжения микрорайона на '
    'напряжении 6 кВ, расчёт и выбор сечений кабельных линий с проверкой '
    'по нагреву, экономической плотности тока и потерям напряжения;',
    '5) выбор схем и элементов распределительных сетей 0,4 кВ, а также '
    'проверка коммутационных аппаратов и устройств релейной защиты;',
    '6) расчёт токов короткого замыкания на шинах подстанций и в '
    'характерных точках сети;',
    '7) расчёт контура заземления и элементов молниезащиты ТП, '
    'обоснование мероприятий по безопасной эксплуатации электроустановок;',
    '8) технико-экономическое обоснование принятых решений: расчёт '
    'капитальных вложений, эксплуатационных затрат, срока окупаемости '
    'и рентабельности инвестиций.',

    # Объект / предмет
    'Объектом исследования являются распределительные электрические сети '
    '6/0,4 кВ жилого микрорайона.',
    'Предметом исследования — режимы работы и параметры указанных сетей, '
    'а также методы выбора и проверки их основного электротехнического '
    'оборудования.',

    # Методы
    'Методологическую основу работы составляют методы расчёта '
    'электрических нагрузок по удельным показателям, выбор сечений '
    'проводников по допустимому нагреву и экономической плотности тока, '
    'методы расчёта токов короткого замыкания в именованных единицах, '
    'а также методика технико-экономического сравнения вариантов.',

    # Значимость
    'Практическая значимость работы заключается в получении готовых '
    'проектных решений по реконструкции электроснабжения рассматриваемого '
    'микрорайона, применимых при выполнении рабочей документации.',

    # Структура
    'Работа состоит из введения, восьми разделов, заключения, списка '
    'использованных источников и приложений; содержит расчётные формулы, '
    'таблицы и иллюстрации, поясняющие принятые технические решения.',
)


def _find_body_index_of_paragraph(body, p_elem) -> int | None:
    """Найти позицию в body.children (top-level)."""
    for i, child in enumerate(body):
        if child is p_elem:
            return i
    return None


def _p_text_all(p_elem) -> str:
    parts = p_elem.findall('.//' + qn('w:t'))
    return ''.join((x.text or '') for x in parts).strip()


def replace_intro(doc) -> bool:
    """Найти параграф «ВВЕДЕНИЕ» и переписать его тело.

    1. Находим ВВЕДЕНИЕ (top-level `<w:p>`, текст строго «ВВЕДЕНИЕ»).
    2. Удаляем все top-level `<w:p>` после него до первого заголовка
       (следующий «1 …» / «1.» / Heading 1 / текст с «ХАРАКТЕРИСТИКА»).
    3. Вставляем новые параграфы из `_INTRO_BODY`.
    """
    body = doc.element.body
    # Найти «ВВЕДЕНИЕ»
    vved_elem = None
    for child in body:
        if child.tag != qn('w:p'):
            continue
        if _p_text_all(child).upper() == 'ВВЕДЕНИЕ':
            vved_elem = child
            break
    if vved_elem is None:
        return False
    vved_idx = _find_body_index_of_paragraph(body, vved_elem)
    if vved_idx is None:
        return False

    # Собираем список детей и ищем границу (первый «настоящий» заголовок)
    children = list(body)
    end_idx = None
    for i in range(vved_idx + 1, len(children)):
        c = children[i]
        if c.tag != qn('w:p'):
            # Таблица или sdt — конец интро
            end_idx = i
            break
        t = _p_text_all(c)
        # Заголовок главы 1 / Heading 1 / «ХАРАКТЕРИСТИКА…»
        if re.match(r'^\s*1[.\s]+[А-ЯЁ]{4,}', t):
            end_idx = i
            break
        if 'ХАРАКТЕРИСТИКА' in t.upper() and len(t) < 80:
            end_idx = i
            break
        # Heading 1 стиль
        pStyle = c.find('.//' + qn('w:pStyle'))
        if pStyle is not None:
            val = pStyle.get(qn('w:val')) or ''
            if 'Heading1' in val or val == '1':
                end_idx = i
                break
    if end_idx is None:
        end_idx = vved_idx + 1

    # Удалить интервал (vved_idx+1 … end_idx-1)
    for i in range(end_idx - 1, vved_idx, -1):
        body.remove(children[i])

    # Вставить новые параграфы ПОСЛЕ ВВЕДЕНИЯ
    # Новые параграфы вставляем сразу после vved_elem
    anchor = vved_elem
    for text in _INTRO_BODY:
        new_p = _make_body_paragraph(text)
        anchor.addnext(new_p)
        anchor = new_p
    return True


def _make_body_paragraph(text: str):
    """Сделать `<w:p>` с ГОСТ-форматом основного текста."""
    p = OxmlElement('w:p')
    # pPr
    pPr = OxmlElement('w:pPr')
    # spacing 1.5, after=0
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:after'), '0')
    spacing.set(qn('w:line'), str(int(cfg.LINE_SPACING * 240)))
    spacing.set(qn('w:lineRule'), 'auto')
    pPr.append(spacing)
    # indent first-line 1.25 cm
    ind = OxmlElement('w:ind')
    ind.set(qn('w:firstLine'), str(int(cfg.FIRST_LINE_INDENT.cm * 567)))
    pPr.append(ind)
    # jc=both
    jc = OxmlElement('w:jc')
    jc.set(qn('w:val'), 'both')
    pPr.append(jc)
    p.append(pPr)
    # run
    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), cfg.FONT_NAME)
    rFonts.set(qn('w:hAnsi'), cfg.FONT_NAME)
    rFonts.set(qn('w:cs'), cfg.FONT_NAME)
    rFonts.set(qn('w:eastAsia'), cfg.FONT_NAME)
    rPr.append(rFonts)
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), str(int(cfg.FONT_SIZE_MAIN.pt * 2)))  # half-points
    rPr.append(sz)
    szCs = OxmlElement('w:szCs')
    szCs.set(qn('w:val'), str(int(cfg.FONT_SIZE_MAIN.pt * 2)))
    rPr.append(szCs)
    r.append(rPr)
    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = text
    r.append(t)
    p.append(r)
    return p


def _make_heading_paragraph(text: str, *, center: bool = True,
                            page_break_before: bool = True):
    """`<w:p>` для структурного заголовка (Times, 16pt, заглавные)."""
    p = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')
    if page_break_before:
        pbb = OxmlElement('w:pageBreakBefore')
        pPr.append(pbb)
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:before'), '0')
    spacing.set(qn('w:after'), '240')
    spacing.set(qn('w:line'), str(int(cfg.LINE_SPACING * 240)))
    spacing.set(qn('w:lineRule'), 'auto')
    pPr.append(spacing)
    jc = OxmlElement('w:jc')
    jc.set(qn('w:val'), 'center' if center else 'left')
    pPr.append(jc)
    p.append(pPr)
    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), cfg.FONT_NAME)
    rFonts.set(qn('w:hAnsi'), cfg.FONT_NAME)
    rFonts.set(qn('w:cs'), cfg.FONT_NAME)
    rPr.append(rFonts)
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), str(int(cfg.FONT_SIZE_H1.pt * 2)))
    rPr.append(sz)
    szCs = OxmlElement('w:szCs')
    szCs.set(qn('w:val'), str(int(cfg.FONT_SIZE_H1.pt * 2)))
    rPr.append(szCs)
    r.append(rPr)
    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = text
    r.append(t)
    p.append(r)
    return p


# ---------------------------------------------------------------------------
# 4. РЕФЕРАТ
# ---------------------------------------------------------------------------

_REFERAT_BODY: tuple[str, ...] = (
    'Выпускная квалификационная работа изложена на 72 страницах машинописного '
    'текста, содержит введение, 8 разделов, заключение, список использованных '
    'источников из 10 наименований; работа включает 16 таблиц, 9 рисунков и '
    '2 приложения.',
    'Ключевые слова: ЭЛЕКТРОСНАБЖЕНИЕ, ТРАНСФОРМАТОРНАЯ ПОДСТАНЦИЯ, '
    'РАСПРЕДЕЛИТЕЛЬНАЯ СЕТЬ, КАБЕЛЬНАЯ ЛИНИЯ, ЭЛЕКТРИЧЕСКАЯ НАГРУЗКА, '
    'ТОК КОРОТКОГО ЗАМЫКАНИЯ, ЗАЗЕМЛЕНИЕ, МОЛНИЕЗАЩИТА, '
    'КАТЕГОРИЯ НАДЁЖНОСТИ, ТЕХНИКО-ЭКОНОМИЧЕСКИЙ РАСЧЁТ.',
    'Цель работы — разработать проект реконструкции электроснабжения '
    'жилого квартала Котова в г. Краснодоне с обоснованием принятых '
    'технических решений по надёжности, качеству электроэнергии и '
    'экономической эффективности.',
    'Объект исследования — распределительные электрические сети 6/0,4 кВ '
    'жилого микрорайона. Предметом исследования являются режимы работы и '
    'параметры этих сетей, а также выбор основного электротехнического '
    'оборудования.',
    'В ходе выполнения работы проведён расчёт электрических нагрузок '
    'микрорайона, выбраны число и мощность трансформаторных подстанций, '
    'обоснованы схемы и элементы сетей 6 кВ и 0,4 кВ, выполнен расчёт '
    'токов короткого замыкания, заземляющего устройства и молниезащиты, '
    'рассчитаны технико-экономические показатели проекта.',
    'Полученные результаты могут быть использованы при разработке рабочей '
    'документации на реконструкцию сетей электроснабжения микрорайона.',
)


def insert_referat(doc) -> bool:
    """Вставить блок РЕФЕРАТ сразу перед первым структурным элементом
    основной зоны (СОДЕРЖАНИЕ или ВВЕДЕНИЕ).

    Если РЕФЕРАТ уже есть — ничего не делаем.
    """
    body = doc.element.body
    # Уже есть?
    for child in body:
        if child.tag != qn('w:p'):
            continue
        if _p_text_all(child).upper() in ('РЕФЕРАТ', 'АННОТАЦИЯ'):
            return False

    # Цель — вставить РЕФЕРАТ до СОДЕРЖАНИЯ, а если его нет — до ВВЕДЕНИЯ.
    # По ГОСТ 7.32 порядок: титул → РЕФЕРАТ → СОДЕРЖАНИЕ → ВВЕДЕНИЕ.
    anchor = None
    for child in body:
        if child.tag != qn('w:p'):
            continue
        txt = _p_text_all(child).upper().strip('.').strip()
        if txt in ('СОДЕРЖАНИЕ', 'ОГЛАВЛЕНИЕ'):
            anchor = child
            break
    if anchor is None:
        for child in body:
            if child.tag == qn('w:p') and _p_text_all(child).upper() == 'ВВЕДЕНИЕ':
                anchor = child
                break
    if anchor is None:
        return False

    # Вставка ПЕРЕД anchor:
    heading = _make_heading_paragraph('РЕФЕРАТ', center=True,
                                      page_break_before=True)
    anchor.addprevious(heading)
    for text in _REFERAT_BODY:
        p = _make_body_paragraph(text)
        anchor.addprevious(p)
    return True


# ---------------------------------------------------------------------------
# 5. Чистка хвоста
# ---------------------------------------------------------------------------

def remove_blank_tail(doc) -> int:
    """Удалить хвостовые пустые параграфы.

    Оставляем: последний непустой + максимум 1 пустой буфер + `sectPr`
    (если хранится в последнем p).

    Мы НЕ трогаем секции — `sectPr` может лежать в pPr последнего
    параграфа; если так — сохраняем его, перенося в предыдущий.
    """
    body = doc.element.body
    # top-level paragraphs
    paras = [c for c in body if c.tag == qn('w:p')]
    if not paras:
        return 0

    # Где последний непустой?
    def is_non_empty(p_elem) -> bool:
        txt = _p_text_all(p_elem)
        if txt:
            return True
        # рисунок / picture
        if p_elem.findall('.//' + qn('w:drawing')):
            return True
        if p_elem.findall('.//' + qn('w:pict')):
            return True
        return False

    last_nonempty = -1
    for i, p in enumerate(paras):
        if is_non_empty(p):
            last_nonempty = i

    if last_nonempty < 0 or last_nonempty >= len(paras) - 1:
        return 0

    # Параграф с финальным sectPr:
    sectPr_carrier_idx = None
    for i, p in enumerate(paras):
        pPr = p.find(qn('w:pPr'))
        if pPr is None:
            continue
        sp = pPr.find(qn('w:sectPr'))
        if sp is not None:
            sectPr_carrier_idx = i

    # Мы оставляем:
    #   • last_nonempty
    #   • один пустой буфер после (если был)
    #   • sectPr carrier (перенесём sectPr в pPr последнего непустого, если
    #     понадобится)
    keep_limit_idx = last_nonempty + 1  # один буфер разрешаем
    if keep_limit_idx >= len(paras):
        return 0
    # Параграфы после keep_limit_idx — удалить
    removed = 0
    for i in range(len(paras) - 1, keep_limit_idx, -1):
        pp = paras[i]
        # Если здесь sectPr — перенесём его в last_nonempty pPr
        pPr = pp.find(qn('w:pPr'))
        if pPr is not None:
            sp = pPr.find(qn('w:sectPr'))
            if sp is not None:
                # перенести
                target = paras[last_nonempty]
                tgt_pPr = target.find(qn('w:pPr'))
                if tgt_pPr is None:
                    tgt_pPr = OxmlElement('w:pPr')
                    target.insert(0, tgt_pPr)
                # убрать старый sectPr в target, если был
                for old in list(tgt_pPr.findall(qn('w:sectPr'))):
                    tgt_pPr.remove(old)
                tgt_pPr.append(deepcopy(sp))
        body.remove(pp)
        removed += 1
    return removed
