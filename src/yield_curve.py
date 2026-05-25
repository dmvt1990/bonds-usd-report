"""
Построение графика «Кривая доходности замещающих облигаций в USD».

Для каждой бумаги берём (Мод. дюрация, YTM валютный) и размещаем
на scatter-plot. Суверенные ОФЗ отделены по полю TYPE = 'ofz_bond'
и показываются чёрными точками с линией тренда (полиномиальная
аппроксимация 2-го порядка). Корпоративные — оливковые точки без
линии.

YTM считается функцией compute_ytm из модуля ytm.py — MOEX для
валютных бумаг возвращает рублёвую доходность, которая не имеет
экономического смысла (см. историю проекта).

График рендерится через matplotlib в PNG и возвращается как
BytesIO — его потом вставляют в слайд через slide.shapes.add_picture.
"""
from datetime import date
from io import BytesIO
from typing import List, Optional, Tuple

import matplotlib
matplotlib.use("Agg")  # без дисплея, важно на сервере
import matplotlib.pyplot as plt
import numpy as np

from .models import Bond
from .ytm import compute_ytm


# Цвета строго из шаблона Газпромбанка
COLOR_SOVEREIGN = "#262626"     # ОФЗ — тёмно-серый (почти чёрный)
COLOR_CORPORATE = "#B9A77A"     # корпы — оливковый акцент
COLOR_GRID = "#D9D9D9"          # светло-серая сетка
COLOR_AXIS_LABEL = "#525252"    # подписи осей


# Список ISIN, которые не нужно включать в график кривой доходности.
# Обычно это бумаги с нестандартной структурой (амортизация, двойные
# выплаты и т.п.), для которых наша формула YTM даёт некорректный
# результат.
EXCLUDE_FROM_CURVE = {
    "RU000A10A8E8",   # РФ ЗО 30 Д — суверенная замещающая с амортизацией долга
}


def _classify(b: Bond) -> str:
    """Определяет группу бумаги: 'sovereign' или 'corporate'.

    Полагаемся на TYPE из MOEX:
      - 'ofz_bond' — суверенные бумаги Минфина (включая замещающие РФ ЗО)
      - 'exchange_bond' — корпоративные
    Плюс подстраховка по SHORTNAME ('РФ ЗО' / 'Россия').
    """
    if b.type_ == "ofz_bond":
        return "sovereign"
    name = (b.shortname or "").upper()
    if name.startswith("РФ ЗО") or name.startswith("РОССИЯ"):
        return "sovereign"
    return "corporate"


def _collect_points(bonds: List[Bond], report_date: date) -> Tuple[List, List]:
    """
    Пробегаемся по бумагам, считаем YTM, отбираем валидные точки.

    Возвращает два списка кортежей (duration_years, ytm_percent, shortname):
      - sovereign_points: суверенные
      - corporate_points: корпоративные

    Пропускаются бумаги:
      - без дюрации или без цены (компонентов для расчёта YTM)
      - с дюрацией < 0.1 (скоро гасятся, YTM нерелевантна)
      - с YTM < 0% (отрицательные доходности обычно баг данных)
      - с YTM > 50% (заведомо мусорные значения из моков или CSV)
    """
    sovereign = []
    corporate = []

    for b in bonds:
        if b.isin in EXCLUDE_FROM_CURVE:
            continue
        if b.duration_years is None or b.duration_years < 0.1:
            continue
        if b.matdate is None:
            continue

        ytm = compute_ytm(
            price=b.close,
            coupon_percent=b.couponpercent,
            matdate=b.matdate,
            report_date=report_date,
            coupon_frequency=b.couponfrequency,
        )
        if ytm is None:
            continue
        # Широкий фильтр — отсекает только заведомо мусорные выбросы.
        # Небольшие отрицательные YTM (до -5%) разрешены — они бывают
        # у коротких бумаг, торгующихся выше номинала. Верхняя граница
        # 50% ловит мусорные значения из моков.
        if ytm < -5 or ytm > 50:
            continue

        point = (b.duration_years, ytm, b.shortname or b.isin)
        if _classify(b) == "sovereign":
            sovereign.append(point)
        else:
            corporate.append(point)

    return sovereign, corporate


def _fit_curve(xs: np.ndarray, ys: np.ndarray, degree: int = 2,
               n_points: int = 100) -> Tuple[np.ndarray, np.ndarray]:
    """
    Подгоняет полином degree-го порядка через точки и возвращает
    (x_dense, y_dense) для отрисовки гладкой линии.

    Если точек меньше чем degree+1 — просто соединяет точки прямой.
    """
    if len(xs) < degree + 1:
        # Слишком мало точек для полинома нужной степени — используем линию
        degree = max(1, len(xs) - 1)
        if degree < 1:
            return np.array([]), np.array([])

    coefs = np.polyfit(xs, ys, degree)
    x_dense = np.linspace(xs.min(), xs.max(), n_points)
    y_dense = np.polyval(coefs, x_dense)
    return x_dense, y_dense


def build_yield_curve_chart(
    bonds: List[Bond],
    report_date: date,
    *,
    currency: str = "USD",
    width_inch: float = 11.22,    # 28.50 см / 2.54 см/дюйм
    height_inch: float = 4.53,    # 11.50 см / 2.54 см/дюйм
    dpi: int = 150,
) -> Optional[BytesIO]:
    """
    Строит график кривой доходности и возвращает PNG как BytesIO.

    Если данных недостаточно (нет точек ни одного типа) — возвращает None.
    Вызывающий код может тогда просто не добавлять слайд.

    Args:
        bonds: список USD-облигаций
        report_date: дата отчёта (для расчёта YTM)
        width_inch, height_inch: размеры фигуры — под слайд 33.87 × 19.05 см
                                 удобно взять ~28 см × 15 см = ~11 × 6 дюймов
        dpi: разрешение PNG. 150 даёт чёткую картинку без слишком
             раздутого файла.

    Returns:
        BytesIO c PNG, курсор в начале, либо None если нечего рисовать.
    """
    sovereign, corporate = _collect_points(bonds, report_date)

    if not sovereign and not corporate:
        return None

    fig, ax = plt.subplots(figsize=(width_inch, height_inch), dpi=dpi)
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")

    # --- Корпоративные: только точки ---
    if corporate:
        xs = np.array([p[0] for p in corporate])
        ys = np.array([p[1] for p in corporate])
        ax.scatter(
            xs, ys,
            color=COLOR_CORPORATE,
            s=40,
            marker="o",
            edgecolor="white",
            linewidth=0.5,
            label=f"Корпоративные ({len(corporate)})",
            zorder=3,
        )

    # --- Суверенные: точки + полиномиальный тренд ---
    if sovereign:
        xs = np.array([p[0] for p in sovereign])
        ys = np.array([p[1] for p in sovereign])

        # Линия тренда поверх (чтобы была видна между точками, но под ними
        # при пересечении — выставляем zorder)
        if len(sovereign) >= 3:
            x_curve, y_curve = _fit_curve(xs, ys, degree=2)
            ax.plot(
                x_curve, y_curve,
                color=COLOR_SOVEREIGN,
                linewidth=2.0,
                alpha=0.85,
                zorder=2,
                label=f"Тренд ОФЗ в {currency}",
            )

        ax.scatter(
            xs, ys,
            color=COLOR_SOVEREIGN,
            s=60,
            marker="o",
            edgecolor="white",
            linewidth=0.7,
            label=f"ОФЗ в {currency} ({len(sovereign)})",
            zorder=4,
        )

    # --- Оформление ---
    ax.set_xlabel("Модифицированная дюрация, лет", fontsize=11, color=COLOR_AXIS_LABEL)
    ax.set_ylabel("YTM, % годовых", fontsize=11, color=COLOR_AXIS_LABEL)
    # Заголовок не ставим — он уже есть на самом слайде PPT.

    # Сетка — только горизонтальная и светлая
    ax.grid(True, axis="y", color=COLOR_GRID, linewidth=0.7, alpha=0.7, zorder=1)
    ax.grid(False, axis="x")

    # Убираем верхнюю и правую рамки
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color(COLOR_GRID)
    ax.spines["bottom"].set_color(COLOR_GRID)
    ax.tick_params(colors=COLOR_AXIS_LABEL, labelsize=10)

    # Легенда — в правом верхнем углу. В районе высокого YTM обычно
    # точек мало (высокодоходные бумаги редки), меньше шансов перекрыть.
    legend = ax.legend(
        loc="upper right",
        frameon=True,
        facecolor="white",
        edgecolor=COLOR_GRID,
        fontsize=10,
    )
    for text in legend.get_texts():
        text.set_color(COLOR_AXIS_LABEL)

    # Подписи всех точек — на образце они есть.
    # Для читаемости подписываем все точки названием.
    for points, color in [(sovereign, COLOR_SOVEREIGN), (corporate, COLOR_CORPORATE)]:
        for x, y, label in points:
            ax.annotate(
                label,
                xy=(x, y),
                xytext=(5, 3),
                textcoords="offset points",
                fontsize=7,
                color=color,
                alpha=0.85,
            )

    fig.tight_layout()

    buf = BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return buf
