# -*- coding: utf-8 -*-
"""
prepare_csv.py — превращает выгрузки iiko в один «длинный» CSV для dashboard.py.

Зачем отдельный скрипт: в iiko нет отчёта, где уценка и списание лежали бы рядом.
Уценка живёт в «Отчёте по продажам со скидкой 20%», а списания и возвраты по сроку —
в «Товарной ведомости». Здесь мы их сводим в одну таблицу, где каждая строка —
одно событие одного дня.

Формат CSV на выходе (он же — контракт для dashboard.py):

    дата        YYYY-MM-DD
    филиал      название склада филиала, как в iiko
    категория   категория блюда из iiko (Торты, Бәліш, Выпечка, ...)
    блюдо       наименование позиции
    событие     одно из: продажа | скидка_20 | возврат_по_сроку |
                склад_50 | списание | возврат_по_качеству
    количество  штук (дробное допустимо: кусочек торта = 1/8)
    сумма       тенге; для «скидка_20» это РАЗМЕР СКИДКИ, а не выручка

Запуск:
    python prepare_csv.py --vedomost tovarnaya.xlsx --tovarny tovarny.xlsx \
                          --prodaji20 pos_b.xlsx -o dannye.csv

Автор оставил подробные комментарии намеренно: файл читают не программисты.
"""

import argparse
import re
import sys
from pathlib import Path

import pandas as pd

# ─────────────────────────────────────────────────────────────────────────────
# Как называются виды документов в товарной ведомости и во что мы их переводим.
# Ключ ищется в колонке «Вид документа» без учёта регистра, как подстрока.
# ─────────────────────────────────────────────────────────────────────────────
VIDY_DOKUMENTOV = {
    "акт реализации": "продажа",
    "возврат":        "возврат_по_сроку",
    "акт списания":   "списание",
}

# Торт режут на 8 кусочков. Позиция «X (кусочек)» — это 1/8 целого торта,
# поэтому её количество делим, а название приводим к базовому.
KUSOCHKOV_V_TORTE = 8
RE_KUSOCHEK = re.compile(r"\s*\(?\s*кусочек\s*\)?\s*$", re.IGNORECASE)

# Склад с «50%» в названии — это не филиал, а куда уходит продукция на уценку 50%.
RE_SKLAD_50 = re.compile(r"50\s*%", re.IGNORECASE)

# ─────────────────────────────────────────────────────────────────────────────
# Назначения документов из товарного отчёта. В товарной ведомости и возврат
# по сроку, и уценка на 50% выглядят одинаково — «Расходная накладная».
# Отличить их можно только по колонке «Наименование» товарного отчёта,
# поэтому третий файл (--tovarny) нужен, чтобы получить возвраты по сроку.
# Ключ ищется как подстрока, без учёта регистра.
# ─────────────────────────────────────────────────────────────────────────────
NAZNACHENIYA = {
    "по срокам":   "возврат_по_сроку",
    "50%":         "склад_50",
    "по качеству": "возврат_по_качеству",
}


def normalizovat(text) -> str:
    """Убирает лишние пробелы, кавычки и различие ё/е — чтобы одно и то же
    название из двух разных отчётов совпало при склейке."""
    if pd.isna(text):
        return ""
    s = str(text).replace("ё", "е").replace("Ё", "Е")
    s = re.sub(r'["«»""]', "", s)
    return re.sub(r"\s+", " ", s).strip()


def bazovoe_blyudo(nazvanie: str):
    """«Наполеон торт (кусочек)» → ('Наполеон торт', 0.125).
    Возвращает пару (базовое название, множитель количества)."""
    if RE_KUSOCHEK.search(nazvanie):
        return RE_KUSOCHEK.sub("", nazvanie).strip(), 1 / KUSOCHKOV_V_TORTE
    return nazvanie, 1.0


def nayti_shapku(df: pd.DataFrame, obyazatelnye: list, kak_zovut: str) -> int:
    """iiko ставит шапку таблицы не в первую строку: сверху идут название отчёта,
    ресторан, период и итоги. Их количество меняется от выгрузки к выгрузке,
    поэтому шапку ищем по содержимому, а не по номеру строки."""
    for i in range(min(len(df), 30)):
        yacheyki = {normalizovat(v).lower() for v in df.iloc[i].tolist()}
        if all(any(nuzhnoe in y for y in yacheyki) for nuzhnoe in obyazatelnye):
            return i
    raise SystemExit(
        f"Не нашла шапку в файле «{kak_zovut}». Жду колонки: {', '.join(obyazatelnye)}.\n"
        "Проверьте, что выгружен нужный отчёт и он не обрезан сверху."
    )


def kolonka(shapka: pd.Series, imya: str):
    """Номер колонки по части названия. iiko иногда меняет формулировки
    («Количество в ед. изм.» / «Количество»), поэтому ищем подстроку."""
    for i, v in enumerate(shapka):
        if imya in normalizovat(v).lower():
            return i
    return None


def chitat_tovarny(put: Path) -> dict:
    """Товарный отчёт → словарь {(номер документа, тип): назначение}.

    Это справочник «зачем был документ». Товарная ведомость знает, ЧТО и СКОЛЬКО
    уехало, но не знает, ЗАЧЕМ: возврат по сроку и уценка на 50% там неразличимы.
    Ключ склейки — номер документа плюс тип (Приход/Расход), потому что у одного
    номера бывают обе стороны.
    """
    syroy = pd.read_excel(put, header=None)
    i_shapki = nayti_shapku(syroy, ["номер документа", "наименование", "тип"], put.name)
    shapka = syroy.iloc[i_shapki]

    i_nomer = kolonka(shapka, "номер документа")
    i_tip   = kolonka(shapka, "тип")
    i_imya  = kolonka(shapka, "наименование")
    if None in (i_nomer, i_tip, i_imya):
        raise SystemExit(f"В «{put.name}» нет колонок «Номер документа» / «Тип» / «Наименование».")

    spravochnik = {}
    for _, r in syroy.iloc[i_shapki + 1:].iterrows():
        nomer = normalizovat(r[i_nomer])
        if not nomer:
            continue
        imya = normalizovat(r[i_imya]).lower()
        for klyuch, naznachenie in NAZNACHENIYA.items():
            if klyuch in imya:
                spravochnik[(nomer, normalizovat(r[i_tip]).lower())] = naznachenie
                break
    return spravochnik


def chitat_vedomost(put: Path, naznacheniya: dict = None) -> pd.DataFrame:
    """Товарная ведомость → строки «продажа», «возврат_по_сроку», «склад_50», «списание».

    Здесь же прячется единственная неочевидная вещь: перемещение на склад 50%
    в ведомости выглядит как обычное «Перемещение», и отличить его можно
    только по названию склада-получателя.
    """
    syroy = pd.read_excel(put, header=None)
    i_shapki = nayti_shapku(syroy, ["вид документа", "номенклатура", "количество"], put.name)
    shapka = syroy.iloc[i_shapki]

    naznacheniya = naznacheniya or {}

    stolbcy = {
        "data":       kolonka(shapka, "дата"),
        "nomer":      kolonka(shapka, "номер документа"),
        "tip":        kolonka(shapka, "тип"),
        "vid":        kolonka(shapka, "вид документа"),
        "kategoriya": kolonka(shapka, "категория"),
        "blyudo":     kolonka(shapka, "номенклатура"),
        "kolvo":      kolonka(shapka, "количество"),
        "sebest":     kolonka(shapka, "себестоимость"),
        "sklad_ras":  kolonka(shapka, "склад расхода"),
        "sklad_pri":  kolonka(shapka, "склад прихода"),
    }
    net = [k for k, v in stolbcy.items() if v is None and k != "sebest"]
    if net:
        raise SystemExit(f"В «{put.name}» не хватает колонок: {', '.join(net)}")

    dannye = syroy.iloc[i_shapki + 1:]
    stroki = []

    for _, r in dannye.iterrows():
        vid = normalizovat(r[stolbcy["vid"]]).lower()
        if not vid:
            continue

        # Определяем событие. Порядок важен: сначала прямые виды документов
        # (реализация, списание), потом — назначение из товарного отчёта,
        # потом — запасной вариант по названию склада-получателя.
        sobytie = None
        for klyuch, imya in VIDY_DOKUMENTOV.items():
            if klyuch in vid:
                sobytie = imya
                break

        if sobytie is None and stolbcy["nomer"] is not None:
            klyuch_dok = (normalizovat(r[stolbcy["nomer"]]),
                          normalizovat(r[stolbcy["tip"]]).lower() if stolbcy["tip"] is not None else "")
            sobytie = naznacheniya.get(klyuch_dok)

        if sobytie is None and ("перемещен" in vid or "накладная" in vid):
            kuda = normalizovat(r[stolbcy["sklad_pri"]])
            if RE_SKLAD_50.search(kuda):
                sobytie = "склад_50"

        if sobytie is None:
            continue  # инвентаризации, приходы от поставщика и прочее нам не нужны

        data = pd.to_datetime(str(r[stolbcy["data"]])[:10], errors="coerce")
        if pd.isna(data):
            continue

        # Филиал — это склад, С КОТОРОГО ушла продукция.
        filial = normalizovat(r[stolbcy["sklad_ras"]])
        if not filial:
            continue

        blyudo, mnozhitel = bazovoe_blyudo(normalizovat(r[stolbcy["blyudo"]]))
        if not blyudo:
            continue

        kolvo = pd.to_numeric(str(r[stolbcy["kolvo"]]).replace(",", "."), errors="coerce")
        if pd.isna(kolvo) or kolvo <= 0:
            continue

        cena = 0.0
        if stolbcy["sebest"] is not None:
            cena = pd.to_numeric(str(r[stolbcy["sebest"]]).replace(",", "."), errors="coerce")
            cena = 0.0 if pd.isna(cena) else float(cena)

        stroki.append({
            "дата":       data.date().isoformat(),
            "филиал":     filial,
            "категория":  normalizovat(r[stolbcy["kategoriya"]]) or "Без категории",
            "блюдо":      blyudo,
            "событие":    sobytie,
            "количество": float(kolvo) * mnozhitel,
            "сумма":      float(kolvo) * mnozhitel * cena,
        })

    return pd.DataFrame(stroki)


def chitat_prodaji_20(put: Path) -> pd.DataFrame:
    """Отчёт по продажам со скидкой 20% → строки «скидка_20».

    Важно: в «сумму» пишем РАЗМЕР СКИДКИ (сумма без скидки минус сумма со скидкой),
    а не выручку. Уценённая штука — это потеря ровно на размер скидки:
    остальные 80% денег вы всё-таки получили.

    Колонки «Концепция» в этом отчёте содержит филиал, а «Учетный день» — дату,
    причём и то и другое iiko заполняет только в первой строке группы,
    а дальше оставляет пусто. Поэтому протягиваем значения вниз (ffill).
    """
    syroy = pd.read_excel(put, header=None)
    i_shapki = nayti_shapku(syroy, ["блюдо", "количество блюд"], put.name)
    shapka = syroy.iloc[i_shapki]

    stolbcy = {
        "data":      kolonka(shapka, "учетный день") or kolonka(shapka, "день"),
        "filial":    kolonka(shapka, "концепция"),
        "kategoriya": kolonka(shapka, "категория блюда"),
        "blyudo":    kolonka(shapka, "блюдо"),
        "kolvo":     kolonka(shapka, "количество блюд"),
        "bez_skidki": kolonka(shapka, "без скидки"),
        "so_skidkoy": kolonka(shapka, "со скидкой"),
    }
    if stolbcy["blyudo"] is None or stolbcy["kolvo"] is None:
        raise SystemExit(f"В «{put.name}» не нашла колонки «Блюдо» и «Количество блюд».")

    dannye = syroy.iloc[i_shapki + 1:].copy()

    # Протягиваем вниз дату, филиал и категорию — iiko заполняет их только в первой
    # строке каждой группы. Без этого половина строк осталась бы без филиала.
    for klyuch in ("data", "filial", "kategoriya"):
        if stolbcy[klyuch] is not None:
            dannye[stolbcy[klyuch]] = dannye[stolbcy[klyuch]].ffill()

    stroki = []
    for _, r in dannye.iterrows():
        blyudo_syroe = normalizovat(r[stolbcy["blyudo"]])
        # Строки «Выпечка всего», «Скидка 20% всего» — это итоги, не данные.
        if not blyudo_syroe or "всего" in blyudo_syroe.lower():
            continue

        data = pd.to_datetime(str(r[stolbcy["data"]])[:10], errors="coerce")
        if pd.isna(data):
            continue

        kolvo = pd.to_numeric(str(r[stolbcy["kolvo"]]).replace(",", "."), errors="coerce")
        if pd.isna(kolvo) or kolvo <= 0:
            continue

        def chislo(klyuch):
            if stolbcy[klyuch] is None:
                return 0.0
            v = pd.to_numeric(str(r[stolbcy[klyuch]]).replace(",", "."), errors="coerce")
            return 0.0 if pd.isna(v) else float(v)

        razmer_skidki = max(0.0, chislo("bez_skidki") - chislo("so_skidkoy"))
        blyudo, mnozhitel = bazovoe_blyudo(blyudo_syroe)

        stroki.append({
            "дата":       data.date().isoformat(),
            "филиал":     normalizovat(r[stolbcy["filial"]]) or "Не указан",
            "категория":  normalizovat(r[stolbcy["kategoriya"]]) or "Без категории",
            "блюдо":      blyudo,
            "событие":    "скидка_20",
            "количество": float(kolvo) * mnozhitel,
            "сумма":      razmer_skidki,
        })

    return pd.DataFrame(stroki)


def main():
    razbor = argparse.ArgumentParser(
        description="Собирает один CSV из выгрузок iiko для dashboard.py")
    razbor.add_argument("--vedomost", type=Path, help="Товарная ведомость (.xlsx)")
    razbor.add_argument("--tovarny", type=Path,
                        help="Товарный отчёт (.xlsx) — без него не будет возвратов по сроку")
    razbor.add_argument("--prodaji20", type=Path, help="Отчёт по продажам со скидкой 20% (.xlsx)")
    razbor.add_argument("-o", "--out", type=Path, default=Path("dannye.csv"))
    args = razbor.parse_args()

    if not args.vedomost and not args.prodaji20:
        razbor.error("Укажите хотя бы один файл: --vedomost и/или --prodaji20")

    naznacheniya = {}
    if args.tovarny:
        naznacheniya = chitat_tovarny(args.tovarny)
        print(f"Товарный отчёт:     {len(naznacheniya):>6} документов с назначением")

    chasti = []
    if args.vedomost:
        if not naznacheniya:
            print("ВНИМАНИЕ: без --tovarny возвраты по сроку и уценка 50% не различаются.")
        chast = chitat_vedomost(args.vedomost, naznacheniya)
        print(f"Товарная ведомость: {len(chast):>6} строк")
        chasti.append(chast)
    if args.prodaji20:
        chast = chitat_prodaji_20(args.prodaji20)
        print(f"Продажи со скидкой: {len(chast):>6} строк")
        chasti.append(chast)

    itog = pd.concat(chasti, ignore_index=True)
    if itog.empty:
        sys.exit("Ни одной строки не собралось — проверьте, те ли файлы выгружены.")

    itog = itog.sort_values(["дата", "филиал", "блюдо"])
    itog.to_csv(args.out, index=False, encoding="utf-8-sig")  # -sig, чтобы Excel не ломал кириллицу

    print(f"\nГотово: {args.out}  ({len(itog)} строк)")
    print(f"Период: {itog['дата'].min()} — {itog['дата'].max()}")
    print("\nСобытий:")
    for sobytie, n in itog["событие"].value_counts().items():
        print(f"  {sobytie:<18} {n:>6}")


if __name__ == "__main__":
    main()
