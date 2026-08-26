# -*- coding: utf-8 -*-
"""
dashboard.py — дашборд по уценке и списаниям для сети пекарен.

Четыре графика на одном листе:
  1. Динамика потерь по периодам      — линия (две серии: уценка / списание)
  2. Топ-10 позиций по потерям        — горизонтальные столбцы
  3. Доли филиалов в потерях          — кольцевая диаграмма
  4. Связь категорий между собой      — тепловая карта корреляций

Вход — CSV в формате, который делает prepare_csv.py:
    дата, филиал, категория, блюдо, событие, количество, сумма

Запуск:
    python dashboard.py dannye.csv -o dashboard.png
    python dashboard.py --demo                  # без данных, на выдуманных числах
    python dashboard.py dannye.csv --shtuki     # считать в штуках, а не в тенге

Почему деньги, а не штуки, по умолчанию: списанный торт и списанная булочка —
это разные потери, а в штуках они весят одинаково.
"""

import argparse
import sys
import warnings
from pathlib import Path

import matplotlib
matplotlib.use("Agg")                       # рисуем в файл, окно не открываем
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
import pandas as pd
import seaborn as sns

warnings.filterwarnings("ignore", category=FutureWarning)

# ═════════════════════════════════════════════════════════════════════════════
# ОФОРМЛЕНИЕ
# Палитра проверена на различимость при дальтонизме: соседние пары разведены
# не только по цвету, но и по светлоте. Не подбирайте цвета «на глаз» —
# синий с зелёным рядом читаются одинаково у 8% мужчин.
# ═════════════════════════════════════════════════════════════════════════════
FON        = "#fcfcfb"   # фон листа
TEKST      = "#0b0b0b"   # основной текст
TEKST_2    = "#52514e"   # подписи, оси
SETKA      = "#e6e5e2"   # сетка — она должна отступать на задний план

SINIY      = "#2a78d6"   # серия 1
ORANZHEVYY = "#eb6834"   # серия 2
AKVA       = "#1baf7a"   # серия 3
ZHELTYY    = "#eda100"   # серия 4
MAGENTA    = "#e87ba4"
ZELENYY    = "#008300"
FIOLET     = "#4a3aa7"
KRASNYY    = "#e34948"

# Порядок фиксированный: восьмая позиция не «следующий цвет по кругу»,
# а повод свернуть хвост в «Остальные».
PALITRA = [SINIY, ORANZHEVYY, AKVA, ZHELTYY, MAGENTA, ZELENYY, FIOLET, KRASNYY]
SERYY   = "#9a9992"      # для группы «Остальные»

# События, которые считаем потерей, и как они называются на графиках.
POTERI_SPISANIE = ["списание", "возврат_по_сроку", "склад_50", "возврат_по_качеству"]
POTERI_SKIDKA   = ["скидка_20"]

# ─────────────────────────────────────────────────────────────────────────────
# Филиал — это торговая точка. В выгрузке кроме них есть цех, склад сырья и
# промежуточные склады, и они дают больше половины «потерь» — но там списывают
# муку, сахар и яйцо, а не витринную продукцию. Смешивать их с филиалами нельзя:
# в топ-10 вылезет сахар, и разговор с филиалом не состоится.
#
# Признак филиала в этой сети: «ТТ» с номером или трёхзначный префикс (010-, 011).
# Если у вас другие названия — задайте свой шаблон через --filialy.
# ─────────────────────────────────────────────────────────────────────────────
SHABLON_FILIALA = r"(?:тт\s*\d|tt\s*\d|^\d{3}[\s\-–])"


def nastroit_shrifty():
    """Кириллица и казахские буквы (ә, қ, ң, ө, ұ, ү, і) есть в DejaVu Sans.
    Если её нет — matplotlib молча нарисует квадратики, поэтому проверяем явно."""
    dostupnye = {f.name for f in matplotlib.font_manager.fontManager.ttflist}
    for imya in ("DejaVu Sans", "Liberation Sans", "Noto Sans", "Arial"):
        if imya in dostupnye:
            plt.rcParams["font.family"] = imya
            return imya
    print("ВНИМАНИЕ: не нашла шрифт с кириллицей — подписи могут стать квадратиками.")
    return None


def nastroit_stil():
    """Общий стиль: светлый фон, тонкая сетка только по одной оси, без рамок."""
    sns.set_theme(style="white")
    plt.rcParams.update({
        "figure.facecolor": FON,
        "axes.facecolor":   FON,
        "axes.edgecolor":   SETKA,
        "axes.labelcolor":  TEKST_2,
        "axes.titlesize":   12,
        "axes.titleweight": "semibold",
        "axes.titlecolor":  TEKST,
        "text.color":       TEKST,
        "xtick.color":      TEKST_2,
        "ytick.color":      TEKST_2,
        "xtick.labelsize":  9,
        "ytick.labelsize":  9,
        "font.size":        10,
        "axes.grid":        False,
        "figure.dpi":       110,
    })


def ubrat_ramku(ax, ostavit=("left", "bottom")):
    """Рамка вокруг графика ничего не сообщает — убираем всё, кроме нужных осей."""
    for storona, hrebet in ax.spines.items():
        hrebet.set_visible(storona in ostavit)
        hrebet.set_color(SETKA)


def dengi(x, _=None):
    """1 570 130 → «1,6 млн». Полное число на оси нечитаемо."""
    x = float(x)
    if abs(x) >= 1_000_000:
        return f"{x / 1_000_000:.1f}".replace(".", ",") + " млн"
    if abs(x) >= 1_000:
        return f"{x / 1_000:.0f} тыс"
    return f"{x:.0f}"


def shtuki(x, _=None):
    """Штуки не округляем до тысяч: 2 688 списанных булочек — это не «3 тыс»,
    а конкретное число, по которому считают заявку."""
    x = float(x)
    if abs(x) >= 10_000:
        return f"{x:,.0f}".replace(",", " ")
    return f"{x:,.0f}".replace(",", " ") if abs(x) >= 1 else f"{x:.1f}"


def korotko(text, predel=28):
    """Длинные названия блюд не влезают в подпись — обрезаем по словам."""
    text = str(text)
    return text if len(text) <= predel else text[:predel - 1].rstrip() + "…"


# ═════════════════════════════════════════════════════════════════════════════
# ЗАГРУЗКА И ПОДГОТОВКА
# ═════════════════════════════════════════════════════════════════════════════
def zagruzit(put: Path) -> pd.DataFrame:
    """Читает CSV и проверяет, что в нём есть всё нужное.
    Лучше упасть с понятным сообщением сейчас, чем нарисовать пустой график."""
    df = pd.read_csv(put, encoding="utf-8-sig")

    nuzhnye = {"дата", "филиал", "категория", "блюдо", "событие", "количество", "сумма"}
    ne_hvataet = nuzhnye - set(df.columns)
    if ne_hvataet:
        sys.exit(f"В CSV нет колонок: {', '.join(sorted(ne_hvataet))}\n"
                 f"Есть: {', '.join(df.columns)}")

    df["дата"] = pd.to_datetime(df["дата"], errors="coerce")
    plohie = df["дата"].isna().sum()
    if plohie:
        print(f"Пропущено {plohie} строк с нечитаемой датой.")
        df = df.dropna(subset=["дата"])

    for stolbec in ("количество", "сумма"):
        df[stolbec] = pd.to_numeric(df[stolbec], errors="coerce").fillna(0.0)

    if df.empty:
        sys.exit("После очистки не осталось ни одной строки.")
    return df


# Шаги для графика динамики: от крупного к мелкому.
SHAGI = [("ME", "по месяцам", "%b %Y"), ("W", "по неделям", "%d.%m"), ("D", "по дням", "%d.%m")]

# В pandas у группировки и у периода разные написания одного и того же шага:
# группировать надо по "ME" (month end), а Period понимает только "M".
SHAG_PERIODA = {"ME": "M", "W": "W", "D": "D"}


def ryad_po_periodam(poteri: pd.DataFrame, shag: str):
    """Суммы по периодам, из которых выброшены неполные периоды по краям.

    Зачем: выгрузка почти никогда не начинается в понедельник и не кончается в
    воскресенье. Неполная неделя всегда ниже полной — и на графике это выглядит
    как обвал, которого не было. Возвращает (таблица, сколько краёв обрезали).
    """
    ryad = (poteri
            .groupby([pd.Grouper(key="дата", freq=shag), "вид"])["мера"]
            .sum().unstack(fill_value=0))
    if ryad.empty:
        return ryad, 0

    pervyy_den, posledniy_den = poteri["дата"].min(), poteri["дата"].max()
    obrezano = 0

    # Левый край: период начался раньше, чем начались наши данные.
    shag_p = SHAG_PERIODA[shag]
    nachalo = ryad.index[0].to_period(shag_p).start_time
    if pervyy_den > nachalo and len(ryad) > 1:
        ryad = ryad.iloc[1:]
        obrezano += 1
    # Правый край: период кончается позже последнего дня данных.
    if len(ryad) > 1:
        konec = ryad.index[-1].to_period(shag_p).end_time.normalize()
        if posledniy_den < konec:
            ryad = ryad.iloc[:-1]
            obrezano += 1

    return ryad, obrezano


def vybrat_period(poteri: pd.DataFrame):
    """Берёт самый крупный шаг, на котором остаётся хотя бы три полные точки.

    По двум точкам тренда не бывает — это отрезок. Поэтому если месяцев мало,
    честнее показать недели, а если и недель мало — дни.
    """
    for shag, podpis, format_daty in SHAGI:
        ryad, obrezano = ryad_po_periodam(poteri, shag)
        if len(ryad) >= 3:
            return ryad, obrezano, podpis, format_daty
    # Совсем короткий период — рисуем по дням как есть.
    ryad, obrezano = ryad_po_periodam(poteri, "D")
    return ryad, obrezano, "по дням", "%d.%m"


def tolko_vitrinnoe(df: pd.DataFrame, put) -> pd.DataFrame:
    """Оставляет только позиции из справочника витринной продукции.

    Зачем нужен: в филиале списывают не только торты. Туда же попадают
    моющее средство, упаковка, листья салата и прочее хозяйство. В деньгах
    это не так заметно, а в штуках хозтовары дают больше половины «потерь»
    и превращают отчёт в бессмыслицу.

    Файл: .xlsx или .csv с колонкой «Наименование» (или первой колонкой,
    если такой нет). Ровно тот справочник, что используется в веб-сверке.
    """
    if not put:
        return df
    put = Path(put)
    tablica = pd.read_excel(put) if put.suffix.lower() in (".xlsx", ".xls") \
        else pd.read_csv(put, encoding="utf-8-sig")

    stolbec = next((c for c in tablica.columns if "наименование" in str(c).lower()),
                   tablica.columns[0])
    razresheno = {normalizovat_nazvanie(v) for v in tablica[stolbec].dropna()}
    if not razresheno:
        sys.exit(f"В справочнике «{put.name}» не нашла ни одного названия.")

    est = df["блюдо"].map(normalizovat_nazvanie).isin(razresheno)
    print(f"Справочник: {len(razresheno)} позиций, отсеяно строк: {(~est).sum()}")
    if not est.any():
        sys.exit("Ни одно блюдо не совпало со справочником — проверьте, тот ли файл.")
    return df[est].copy()


def normalizovat_nazvanie(v) -> str:
    """Одна и та же позиция в разных отчётах пишется чуть по-разному:
    лишний пробел, кавычки-ёлочки, ё вместо е. Приводим к общему виду."""
    import re as _re
    s = str(v).replace("ё", "е").replace("Ё", "Е")
    s = _re.sub(r'["«»""]', "", s)
    return _re.sub(r"\s+", " ", s).strip().lower()


def svesti_imena_filialov(df: pd.DataFrame) -> pd.DataFrame:
    """Один филиал в разных отчётах называется по-разному: в товарной ведомости
    «TT8 Qaratau склад», в отчёте по продажам просто «TT8 Qaratau». Без склейки
    он попадёт в круговую диаграмму двумя кусками и займёт вдвое больше места,
    чем на самом деле. Убираем хвостовое слово «склад».
    """
    df = df.copy()
    df["филиал"] = (df["филиал"].astype(str)
                    .str.replace(r"\s*склад\s*$", "", case=False, regex=True)
                    .str.replace(r"\s+", " ", regex=True)
                    .str.strip())
    return df


def tolko_filialy(df: pd.DataFrame, shablon: str) -> pd.DataFrame:
    """Оставляет строки торговых точек. Показывает, что именно отсеяли, —
    молчаливая фильтрация опаснее, чем её отсутствие."""
    if not shablon:
        return df
    est_filial = df["филиал"].str.contains(shablon, case=False, regex=True, na=False)
    otseyano = sorted(df.loc[~est_filial, "филиал"].unique())
    if otseyano:
        print(f"Не филиалы, исключены ({len(otseyano)}): {', '.join(otseyano[:8])}"
              + (" …" if len(otseyano) > 8 else ""))
    if not est_filial.any():
        sys.exit(f"Ни одно название склада не подошло под шаблон «{shablon}».\n"
                 "Задайте свой через --filialy или отключите фильтр: --filialy ''")
    return df[est_filial].copy()


def summa_poter(df: pd.DataFrame, v_shtukah: bool) -> pd.DataFrame:
    """Оставляет только события-потери и добавляет колонку «вид потери»."""
    mera = "количество" if v_shtukah else "сумма"
    poteri = df[df["событие"].isin(POTERI_SPISANIE + POTERI_SKIDKA)].copy()
    poteri["вид"] = np.where(poteri["событие"].isin(POTERI_SKIDKA),
                             "Уценка 20%", "Списание и возвраты")
    poteri["мера"] = poteri[mera]
    return poteri


# ═════════════════════════════════════════════════════════════════════════════
# ГРАФИК 1 — ДИНАМИКА
# ═════════════════════════════════════════════════════════════════════════════
def format_dlya(edinica: str):
    """Один и тот же дашборд считает и деньги, и штуки — формат подписей разный."""
    return dengi if edinica == "тенге" else shtuki


def grafik_dinamika(ax, poteri: pd.DataFrame, edinica: str):
    po_periodam, obrezano, podpis, format_daty = vybrat_period(poteri)

    # Порядок серий фиксируем: цвет закреплён за смыслом, а не за местом в списке.
    for i, vid in enumerate(["Списание и возвраты", "Уценка 20%"]):
        if vid not in po_periodam:
            continue
        ax.plot(po_periodam.index, po_periodam[vid],
                color=PALITRA[i], linewidth=2, marker="o", markersize=5,
                markeredgecolor=FON, markeredgewidth=1.2, label=vid, zorder=3)

    # Сетка только горизонтальная: сравниваем высоту, а не положение по времени.
    ax.yaxis.grid(True, color=SETKA, linewidth=1)
    ax.set_axisbelow(True)
    ax.set_title(f"Потери {podpis}", loc="left", pad=12)
    ax.set_ylabel(edinica)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(format_dlya(edinica)))
    ax.set_ylim(bottom=0)                                # у денег нет отрицательной части
    ax.xaxis.set_major_formatter(plt.matplotlib.dates.DateFormatter(format_daty))
    ubrat_ramku(ax)

    # Легенда над полем графика: внутри она обязательно наедет на линию,
    # как только данные пойдут вверх в левой части.
    ax.legend(frameon=False, fontsize=9, ncols=2, labelcolor=TEKST_2,
              handlelength=1.6, loc="lower right", bbox_to_anchor=(1, 1.005))

    if obrezano:
        ax.text(1.0, -0.17, "неполные периоды по краям не показаны",
                transform=ax.transAxes, ha="right", fontsize=8, color=TEKST_2)

    if len(po_periodam) < 3:
        ax.text(0.5, 0.55, "Мало периодов для тренда —\nнужно хотя бы три точки",
                transform=ax.transAxes, ha="center", va="center",
                color=TEKST_2, fontsize=10)


# ═════════════════════════════════════════════════════════════════════════════
# ГРАФИК 2 — ТОП-10 ПОЗИЦИЙ
# ═════════════════════════════════════════════════════════════════════════════
def grafik_top10(ax, poteri: pd.DataFrame, edinica: str):
    """Горизонтальные столбцы, а не вертикальные: названия блюд длинные,
    вертикально они пришлось бы поворачивать, и читать стало бы неудобно."""
    po_blyudam = (poteri.pivot_table(index="блюдо", columns="вид",
                                     values="мера", aggfunc="sum", fill_value=0))
    for vid in ("Списание и возвраты", "Уценка 20%"):
        if vid not in po_blyudam:
            po_blyudam[vid] = 0.0
    po_blyudam["всего"] = po_blyudam.sum(axis=1)
    top = po_blyudam.nlargest(10, "всего").iloc[::-1]     # снизу вверх: крупнейшее сверху

    pozicii = np.arange(len(top))
    ax.barh(pozicii, top["Списание и возвраты"], height=0.62,
            color=PALITRA[0], label="Списание и возвраты", zorder=3)
    ax.barh(pozicii, top["Уценка 20%"], height=0.62,
            left=top["Списание и возвраты"] + top["всего"].max() * 0.004,  # зазор 2px между сегментами
            color=PALITRA[1], label="Уценка 20%", zorder=3)

    ax.set_yticks(pozicii)
    ax.set_yticklabels([korotko(i) for i in top.index], fontsize=9)
    ax.xaxis.grid(True, color=SETKA, linewidth=1)
    ax.set_axisbelow(True)
    ax.set_title("Топ-10 позиций по потерям", loc="left", pad=12)
    ax.set_xlabel(edinica)
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(format_dlya(edinica)))
    ubrat_ramku(ax)
    # Легенда над полем: внутри она перекрывает подписи значений у длинных столбцов.
    ax.legend(frameon=False, fontsize=9, ncols=2, labelcolor=TEKST_2,
              handlelength=1.6, loc="lower right", bbox_to_anchor=(1, 1.005))

    # Подписи значений — только итог у конца столбца. Число на каждом сегменте
    # превратило бы график в таблицу.
    predel = top["всего"].max() * 1.16
    ax.set_xlim(0, predel)
    podpisat = format_dlya(edinica)
    for y, znachenie in zip(pozicii, top["всего"]):
        ax.text(znachenie + predel * 0.015, y, podpisat(znachenie),
                va="center", fontsize=8.5, color=TEKST_2)


# ═════════════════════════════════════════════════════════════════════════════
# ГРАФИК 3 — ДОЛИ ФИЛИАЛОВ
# ═════════════════════════════════════════════════════════════════════════════
def grafik_filialy(ax, poteri: pd.DataFrame, edinica: str, maks_dolek=6):
    """Кольцо вместо сплошного круга — в середину помещается итог.

    Честное предупреждение: круговая диаграмма плохо показывает больше 5–6 долей,
    человек не умеет сравнивать углы. Поэтому берём крупнейшие филиалы,
    а хвост сворачиваем в «Остальные». Если нужно сравнивать все одиннадцать —
    столбцы дадут точный ответ, а круг только общее впечатление.
    """
    po_filialam = poteri.groupby("филиал")["мера"].sum().sort_values(ascending=False)

    if len(po_filialam) > maks_dolek:
        hvost = po_filialam.iloc[maks_dolek:].sum()
        po_filialam = po_filialam.iloc[:maks_dolek]
        po_filialam[f"Остальные ({len(poteri['филиал'].unique()) - maks_dolek})"] = hvost

    cveta = PALITRA[:len(po_filialam)]
    if po_filialam.index[-1].startswith("Остальные"):
        cveta[-1] = SERYY                                 # хвост не должен спорить с данными

    kliny, _ = ax.pie(
        po_filialam.values,
        colors=cveta,
        startangle=90,
        counterclock=False,                               # по часовой: так читают циферблат
        wedgeprops=dict(width=0.42, edgecolor=FON, linewidth=2),  # зазор между долями
    )

    vsego = po_filialam.sum()
    ax.text(0, 0.06, format_dlya(edinica)(vsego), ha="center", va="center",
            fontsize=17, fontweight="semibold", color=TEKST)
    ax.text(0, -0.16, edinica.lower(), ha="center", va="center",
            fontsize=9, color=TEKST_2)

    # Подписи выносим в легенду с процентами — на самом кольце они наезжают
    # друг на друга, когда доли мелкие.
    podpisi = [f"{korotko(imya, 22)} — {dolya / vsego * 100:.0f}%"
               for imya, dolya in po_filialam.items()]
    ax.legend(kliny, podpisi, frameon=False, fontsize=8.5,
              loc="center left", bbox_to_anchor=(0.98, 0.5), labelcolor=TEKST_2)
    ax.set_title("Доли филиалов в потерях", loc="left", pad=12)


# ═════════════════════════════════════════════════════════════════════════════
# ГРАФИК 4 — КОРРЕЛЯЦИЯ КАТЕГОРИЙ
# ═════════════════════════════════════════════════════════════════════════════
def grafik_korrelyaciya(ax, poteri: pd.DataFrame, maks_kategoriy=8):
    """Считаем, насколько согласованно категории теряют деньги изо дня в день.

    Что это даёт: +0,8 между тортами и пирожными означает «плохой день у одних —
    плохой день и у других», то есть причина общая (погода, поток, завоз).
    Близко к нулю — причины разные, и разбираться надо отдельно.

    Корреляция НЕ доказывает причину. И на коротком периоде она пляшет:
    ниже 30 дней числа лучше считать намёком, а не выводом.
    """
    po_dnyam = (poteri.pivot_table(index="дата", columns="категория",
                                   values="мера", aggfunc="sum", fill_value=0))

    # Оставляем самые весомые категории: матрица 20×20 нечитаема.
    krupnye = po_dnyam.sum().nlargest(maks_kategoriy).index
    po_dnyam = po_dnyam[krupnye]

    # Категория, которая теряла в один и тот же день одинаково (или ни разу),
    # даёт деление на ноль — убираем.
    po_dnyam = po_dnyam.loc[:, po_dnyam.std() > 0]

    if po_dnyam.shape[1] < 2:
        ax.text(0.5, 0.5, "Категорий слишком мало\nдля корреляции",
                transform=ax.transAxes, ha="center", va="center", color=TEKST_2)
        ax.set_axis_off()
        return

    matrica = po_dnyam.corr()

    # Расходящаяся палитра: два полюса и СЕРАЯ середина. Радуга здесь запрещена —
    # у неё нет естественного порядка, и ноль теряется.
    shkala = sns.blend_palette([SINIY, "#f0efec", ORANZHEVYY], as_cmap=True)

    sns.heatmap(
        matrica, ax=ax, cmap=shkala, vmin=-1, vmax=1, center=0,
        annot=True, fmt=".2f", annot_kws={"size": 8},
        linewidths=2, linecolor=FON,                      # зазор между ячейками
        cbar_kws={"shrink": 0.7, "ticks": [-1, 0, 1], "label": "корреляция"},
        square=True,
    )
    ax.set_title(f"Связь категорий между собой  ·  {len(po_dnyam)} дн.",
                 loc="left", pad=12)
    ax.set_xlabel(""); ax.set_ylabel("")
    ax.set_xticklabels([korotko(t.get_text(), 14) for t in ax.get_xticklabels()],
                       rotation=35, ha="right", fontsize=8.5)
    ax.set_yticklabels([korotko(t.get_text(), 14) for t in ax.get_yticklabels()],
                       rotation=0, fontsize=8.5)

    if len(po_dnyam) < 30:
        # Приписка идёт в заголовок, а не под ось: снизу она наезжает
        # на повёрнутые подписи категорий.
        ax.set_title(f"Связь категорий между собой  ·  {len(po_dnyam)} дн., "
                     f"числа неустойчивы", loc="left", pad=12)


# ═════════════════════════════════════════════════════════════════════════════
# ШАПКА С ИТОГАМИ
# ═════════════════════════════════════════════════════════════════════════════
def shapka(fig, df: pd.DataFrame, poteri: pd.DataFrame, edinica: str, v_shtukah: bool):
    """Четыре числа наверху. Их читают первыми, а часто и единственными."""
    mera_prodaj = "количество" if v_shtukah else "сумма"
    prodano = df.loc[df["событие"] == "продажа", mera_prodaj].sum()
    vsego = poteri["мера"].sum()
    skidka = poteri.loc[poteri["вид"] == "Уценка 20%", "мера"].sum()
    dolya = vsego / prodano * 100 if prodano else 0

    hudshiy = poteri.groupby("филиал")["мера"].sum()
    hudshiy = hudshiy.idxmax() if len(hudshiy) else "—"

    pokaz = format_dlya(edinica)
    pokazateli = [
        ("Потери всего",        pokaz(vsego),   edinica.lower()),
        ("Из них уценка 20%",   pokaz(skidka),  f"{skidka / vsego * 100:.0f}% потерь" if vsego else ""),
        ("Доля от продаж",      f"{dolya:.1f}%".replace(".", ","), "потери к выручке"),
        ("Больше всех теряет",  korotko(hudshiy, 18), ""),
    ]
    for i, (zagolovok, znachenie, pripiska) in enumerate(pokazateli):
        x = 0.055 + i * 0.235
        fig.text(x, 0.935, zagolovok, fontsize=9, color=TEKST_2)
        razmer = 20 if len(znachenie) < 12 else 13
        fig.text(x, 0.898, znachenie, fontsize=razmer, fontweight="semibold", color=TEKST)
        if pripiska:
            fig.text(x, 0.876, pripiska, fontsize=8.5, color=TEKST_2)


# ═════════════════════════════════════════════════════════════════════════════
# ДЕМО-ДАННЫЕ
# ═════════════════════════════════════════════════════════════════════════════
def demo_dannye() -> pd.DataFrame:
    """Правдоподобные выдуманные данные — чтобы скрипт можно было запустить
    и посмотреть, что получится, ещё до выгрузки из iiko."""
    generator = np.random.default_rng(42)
    dni = pd.date_range("2026-01-01", "2026-08-20", freq="D")
    filialy = ["ТТ1 Атакент", "ТТ2 Сарыарка", "TT3 Lamiya", "ТТ4 Ньютон",
               "ТТ6 Автодом", "TT8 Qaratau", "010-Шугыла", "011 Алма-Сити"]
    blyuda = {
        "Торты":    ["Наполеон торт", "Медовик", "Испанский чизкейк", "Морковный торт"],
        "Бәліш":    ["Алмалы бәліш", "Сүзбелі бәліш", "Қарақыз бәліші"],
        "Выпечка":  ["Бөрек тауық етімен", "Самса", "Бауырсақ", "Сосиска в тесте"],
        "Пирожные": ["Берлинер", "Эклер", "Макаронс", "Трайфл"],
    }
    stroki = []
    for den in dni:
        # Выходные оживлённее буднего дня — потери тоже растут.
        vyhodnoy = 1.35 if den.dayofweek >= 5 else 1.0
        for filial in filialy:
            for kategoriya, spisok in blyuda.items():
                for blyudo in spisok:
                    if generator.random() < 0.55:
                        stroki.append((den, filial, kategoriya, blyudo, "продажа",
                                       generator.integers(3, 30),
                                       generator.integers(3_000, 40_000) * vyhodnoy))
                    if generator.random() < 0.18:
                        kolvo = generator.integers(1, 5)
                        stroki.append((den, filial, kategoriya, blyudo, "списание",
                                       kolvo, kolvo * generator.integers(500, 9_000)))
                    if generator.random() < 0.12:
                        kolvo = generator.integers(1, 4)
                        stroki.append((den, filial, kategoriya, blyudo, "скидка_20",
                                       kolvo, kolvo * generator.integers(100, 1_800)))
    return pd.DataFrame(stroki, columns=["дата", "филиал", "категория", "блюдо",
                                         "событие", "количество", "сумма"])


# ═════════════════════════════════════════════════════════════════════════════
def main():
    razbor = argparse.ArgumentParser(description="Дашборд по уценке и списаниям")
    razbor.add_argument("csv", nargs="?", type=Path, help="CSV от prepare_csv.py")
    razbor.add_argument("--demo", action="store_true", help="Выдуманные данные для примера")
    razbor.add_argument("--shtuki", action="store_true", help="Считать в штуках, а не в тенге")
    razbor.add_argument("--spravochnik", type=Path,
                        help="Справочник витринной продукции (.xlsx/.csv). "
                             "Без него в потери попадут хозтовары и сырьё.")
    razbor.add_argument("--filialy", default=SHABLON_FILIALA,
                        help="Шаблон названия филиала (регулярное выражение). "
                             "Пустая строка — считать все склады подряд.")
    razbor.add_argument("-o", "--out", type=Path, default=Path("dashboard.png"))
    args = razbor.parse_args()

    if not args.csv and not args.demo:
        razbor.error("Укажите CSV или запустите с --demo")

    nastroit_shrifty()
    nastroit_stil()

    df = demo_dannye() if args.demo else zagruzit(args.csv)
    if args.demo:
        df["дата"] = pd.to_datetime(df["дата"])

    df = tolko_filialy(df, args.filialy)
    df = svesti_imena_filialov(df)
    df = tolko_vitrinnoe(df, args.spravochnik)

    edinica = "штук" if args.shtuki else "тенге"
    poteri = summa_poter(df, args.shtuki)
    if poteri.empty:
        sys.exit("В данных нет ни одного события-потери — проверьте колонку «событие».")

    figura = plt.figure(figsize=(16, 11))
    setka = figura.add_gridspec(2, 2, hspace=0.46, wspace=0.30,
                                left=0.055, right=0.965, top=0.815, bottom=0.07)

    grafik_dinamika(figura.add_subplot(setka[0, 0]), poteri, edinica)
    grafik_top10(figura.add_subplot(setka[0, 1]), poteri, edinica)
    grafik_filialy(figura.add_subplot(setka[1, 0]), poteri, edinica)
    grafik_korrelyaciya(figura.add_subplot(setka[1, 1]), poteri)

    shapka(figura, df, poteri, edinica, args.shtuki)

    ot, do = df["дата"].min(), df["дата"].max()
    figura.text(0.055, 0.975, "Уценка и списания", fontsize=17, fontweight="semibold", color=TEKST)
    figura.text(0.965, 0.977, f"{ot:%d.%m.%Y} — {do:%d.%m.%Y}",
                fontsize=10, color=TEKST_2, ha="right")

    figura.savefig(args.out, dpi=200, facecolor=FON, bbox_inches="tight")
    print(f"Готово: {args.out}")

    # PDF рядом — его удобно печатать и вкладывать в письмо.
    pdf = args.out.with_suffix(".pdf")
    figura.savefig(pdf, facecolor=FON, bbox_inches="tight")
    print(f"Готово: {pdf}")


if __name__ == "__main__":
    main()
