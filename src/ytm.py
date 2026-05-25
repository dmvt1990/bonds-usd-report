"""
Расчёт доходности к погашению (YTM) для облигаций с фиксированным купоном.

Используется для валютных бумаг, потому что MOEX для них отдаёт
`YIELDCLOSE` в рублёвой системе координат — получаются бессмысленные
значения (сотни процентов годовых).

Мы считаем YTM в валюте номинала: если бумага с номиналом 100 USD,
купон 5% USD, цена 98% — получим «сколько процентов годовых в USD
заработает инвестор, если додержит до погашения».

Это упрощённый расчёт без учёта НКД (накопленного купонного дохода).
В банковских системах обычно считают "грязную" YTM с поправкой на
НКД — разница небольшая (0.1-0.5 п.п.) и для обзорного отчёта несущественна.

Реализация — метод бисекции на чистом Python, без scipy.
"""
from datetime import date
from typing import Optional


def _pv_of_cashflows(y: float, periods: int, coupon_per_period: float) -> float:
    """
    Приведённая стоимость всех денежных потоков облигации (% от номинала)
    при ставке дисконтирования y (за период).
    """
    if y <= -0.9999:
        return float("inf")

    n = int(round(periods))
    if n <= 0:
        return 100.0 + coupon_per_period

    pv = 0.0
    for k in range(1, n + 1):
        pv += coupon_per_period / ((1 + y) ** k)
    pv += 100.0 / ((1 + y) ** n)
    return pv


def compute_ytm(
    price: Optional[float],
    coupon_percent: Optional[float],
    matdate: Optional[date],
    report_date: date,
    coupon_frequency: Optional[int],
    max_iterations: int = 100,
    tolerance: float = 1e-6,
) -> Optional[float]:
    """
    Расчёт доходности к погашению в годовых процентах.

    Args:
        price: цена облигации, % от номинала
        coupon_percent: купон, % годовых
        matdate: дата погашения
        report_date: дата, на которую считаем доходность
        coupon_frequency: сколько раз в год платят купон (1/2/4/12)

    Returns:
        YTM в % годовых, или None если данных недостаточно.
    """
    if price is None or price <= 0:
        return None
    if coupon_percent is None:
        return None
    if matdate is None or report_date is None:
        return None
    if coupon_frequency is None or coupon_frequency <= 0:
        return None
    if matdate <= report_date:
        return None

    days_to_mat = (matdate - report_date).days
    if days_to_mat <= 0:
        return None
    years_to_mat = days_to_mat / 365.0

    periods = max(1, round(years_to_mat * coupon_frequency))
    coupon_per_period = coupon_percent / coupon_frequency

    lo, hi = -0.5, 2.0

    def f(y_period: float) -> float:
        return _pv_of_cashflows(y_period, periods, coupon_per_period) - price

    f_lo = f(lo)
    f_hi = f(hi)

    if f_lo * f_hi > 0:
        return None

    for _ in range(max_iterations):
        mid = (lo + hi) / 2
        f_mid = f(mid)
        if abs(f_mid) < tolerance:
            break
        if f_lo * f_mid < 0:
            hi = mid
            f_hi = f_mid
        else:
            lo = mid
            f_lo = f_mid

    y_period = (lo + hi) / 2
    ytm_annual = ((1 + y_period) ** coupon_frequency - 1) * 100
    return ytm_annual
