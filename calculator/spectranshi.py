"""Калькулятор «Спецтранши 2.0 — PB Top Up»: перенос Excel-книги на Python.

Повторяет формулы листов «Главная страница», «Ставки ИД», «Решатель» и «Расчет»
в том же порядке арифметических операций, что и Excel, чтобы результаты
совпадали побитово (а не только до копейки).
"""
from __future__ import annotations

import calendar
from dataclasses import dataclass, field
from datetime import date, timedelta
from decimal import ROUND_HALF_UP, Decimal

# Лист «Ставки ИД», базовые ставки (C14:K20). Сроки 6 и 9 мес. берутся из «Решателя».
PREMIUM_THRESHOLDS = [50_000_000, 100_000_000, 150_000_000, 1_000_000_000,
                      5_000_000_000, 10_000_000_000, 20_000_000_000]
BASE_RATES = {
    #              1 мес   2 мес   3 мес  4 мес  12 мес
    50_000_000:     {1: 0.14,  2: 0.1405, 3: 0.142, 4: 0.142, 12: 0.1445},
    100_000_000:    {1: 0.141, 2: 0.1415, 3: 0.143, 4: 0.143, 12: 0.1455},
    150_000_000:    {1: 0.141, 2: 0.1415, 3: 0.143, 4: 0.143, 12: 0.1455},
    1_000_000_000:  {1: 0.141, 2: 0.1415, 3: 0.143, 4: 0.143, 12: 0.1455},
    5_000_000_000:  {1: 0.141, 2: 0.1415, 3: 0.143, 4: 0.143, 12: 0.1455},
    10_000_000_000: {1: 0.141, 2: 0.1415, 3: 0.143, 4: 0.143, 12: 0.1455},
    20_000_000_000: {1: 0.141, 2: 0.1415, 3: 0.143, 4: 0.143, 12: 0.1455},
}
TERMS = (1, 2, 3, 4, 6, 9, 12)          # «Главная страница»!J3:J9
UKV_OPTIONS = (0.0013, 0.0033)           # «Главная страница»!J12:J13
YEAR_2024_END = date(2024, 12, 31)       # «Расчет»!D3
DAILY_TABLE_ROWS = 619                   # Таблица2: строки 4..622


def excel_round(x: float, digits: int) -> float:
    """ROUND() Excel: половина от нуля, по 15 значащим цифрам."""
    d = Decimal(f"{x:.15g}").quantize(Decimal(1).scaleb(-digits), rounding=ROUND_HALF_UP)
    return float(d)


def edate(d: date, months: int) -> date:
    m = d.month - 1 + months
    y, m = d.year + m // 12, m % 12 + 1
    return date(y, m, min(d.day, calendar.monthrange(y, m)[1]))


def year_base(d: date) -> int:
    return 366 if d.year == 2024 else 365


def management_fee(term: float, ukv: float) -> float:
    """«Расчет»!C11 / «Решатель»!C14: ROUND(0.05%*50%*срок+0.1%,4)+УКВ."""
    return excel_round(0.0005 * 0.5 * term + 0.001, 4) + ukv


# ---------------------------------------------------------------- Решатель
@dataclass
class SolverInputs:
    target_rate: float = 0.155   # Решатель!C6
    base_ukv: float = 0.0013     # Решатель!C7
    rate_adjustment: float = 0.0  # Решатель!C8


def solve_rate(start: date, amount: float, uf: float, term: int, s: SolverInputs) -> tuple[float, float]:
    """Решатель!C16/C17 (или C25/C26): точная и округлённая ставка размещения."""
    n = (edate(start, term) - start).days
    mf = management_fee(term, s.base_ukv)
    q = 1 - mf / 365
    exact = ((amount * (1 + s.target_rate * n / 365) - amount * (1 - uf) * q ** n)
             * (1 - q) * 365 / (amount * (1 - uf) * (1 - q ** n)))
    return exact, excel_round(exact - s.rate_adjustment, 4)


# ---------------------------------------------------------------- Ставки ИД
def premium_segment(premium: float) -> int:
    """«Ставки ИД»!V2."""
    for t in reversed(PREMIUM_THRESHOLDS):
        if premium >= t:
            return t
    raise ValueError("ОШИБКА, СЛИШКОМ НИЗКАЯ СУММА (премия < 50 млн)")


def placement_rate(channel: str, premium: float, term: int, solver_rates: dict[int, float]) -> float:
    """«Расчет»!C7."""
    if channel != "PB":
        return 0.0
    segment = premium_segment(premium)
    if term in (6, 9):
        return solver_rates[term]
    if term not in BASE_RATES[segment]:
        raise ValueError(f"Срок {term} мес. отсутствует в таблице ставок (#N/A)")
    return BASE_RATES[segment][term]


# ---------------------------------------------------------------- Расчет
@dataclass
class Inputs:
    """Поля «Главная страница», заполняемые КМ."""
    term_months: float = 6                 # D5
    premium: float = 100_000_000           # D6
    channel: str = "PB"                    # D7
    ukv: float = 0.0013                    # D8
    initial_investment: float = 0.0        # D13
    kv: float = 0.001                      # D16 / Расчет!C9 (на расчёт не влияет)
    start_date: date | None = None         # Расчет!E4 = TODAY()+1
    uf: float = 0.0                        # Расчет!C10
    rf: float = 0.0                        # Расчет!C12
    solver: SolverInputs = field(default_factory=SolverInputs)


@dataclass
class DailyRow:
    day: date
    deposit: float          # F: Депозит
    mf_day: float           # G: MF за каждый день
    deposit_net: float      # I/J: Депозит − MF (накопленный MF за весь срок)


@dataclass
class Result:
    start_date: date
    end_date: date
    amount: float
    placement_rate: float
    mf: float
    uf: float
    rf: float
    mf_total: float
    simple_yield: float           # Расчет!C14
    gross_income: float           # Расчет!C16
    final_payout: float           # Расчет!C18
    final_yield: float            # Расчет!C19
    survival_sum: float           # Главная!D28
    accident_sum: float           # Главная!D29
    solver_rates: dict[int, float]
    solver_exact: dict[int, float]
    rows: list[DailyRow]

    @property
    def main_page(self) -> dict[str, object]:
        """Выходные поля «Главная страница» (D10, D17..D29, D75..D78)."""
        after_tax_rate = 0  # Расчет!C15 пустая ячейка
        after_tax_income = 0  # Расчет!C17 пустая ячейка
        return {
            "Дата окончания": self.end_date,
            "Годовая простая клиентская доходность": self.simple_yield,
            "Ставка после налогов, простая": after_tax_rate,
            "Доход клиента гросс": self.gross_income,
            "Доход минус налог": after_tax_income,
            "Ставка до налогов, итоговая": self.final_yield,
            "Ставка после налогов, итоговая*": self.final_yield - self.simple_yield + after_tax_rate,
            "Доход итоговый* минус налог": self.gross_income - self.gross_income + after_tax_income,
            "UF": self.uf,
            "MF": self.mf,
            "RF": self.rf,
            "СС по Дожитию": self.survival_sum,
            "СС по СНС": self.accident_sum,
        }


def end_date(start: date, term: float) -> date:
    """«Расчет»!C4."""
    whole = int(term)
    frac_days = int(excel_round((term - whole) * 30, 0))
    e = edate(start, whole)
    if e.day < start.day:
        return e + timedelta(days=start.day - e.day + frac_days)
    return e + timedelta(days=frac_days)


def calculate(inp: Inputs) -> Result:
    start = inp.start_date or date.today() + timedelta(days=1)
    if inp.ukv == 0.0054 and inp.term_months < 5:
        amount = 0.0                                   # Расчет!C5
    else:
        amount = inp.premium
    end = end_date(start, inp.term_months)

    solver_exact, solver_rates = {}, {}
    for t in (6, 9):
        solver_exact[t], solver_rates[t] = solve_rate(start, amount, inp.uf, t, inp.solver)

    rate = placement_rate(inp.channel, inp.premium, inp.term_months, solver_rates)
    mf = management_fee(inp.term_months, inp.ukv)

    # Таблица2 (E:J): ежедневный рост депозита по простым процентам,
    # MF начисляется на остаток «депозит вчера − накопленный MF».
    f0 = amount - inp.uf * amount
    days = [start]
    while days[-1] + timedelta(days=1) <= end:
        days.append(days[-1] + timedelta(days=1))
    if len(days) > DAILY_TABLE_ROWS:
        raise ValueError("Срок длиннее таблицы Excel (619 дней)")

    deposits, mfs = [f0], [0.0]
    mf_sum = 0.0
    for i in range(1, len(days)):
        yb = year_base(days[i])
        dt = (days[i] - days[i - 1]).days
        deposits.append(f0 * (rate / yb * dt) + deposits[-1])
        g = (deposits[i - 1] - mf_sum) * mf / yb
        mfs.append(g)
        mf_sum += g
    mf_total = 0.0
    for g in mfs:
        mf_total += g
    rows = [DailyRow(d, f, g, f - mf_total) for d, f, g in zip(days, deposits, mfs)]

    gross = max(deposits) - mf_total - amount                     # C16
    payout = max(r.deposit_net for r in rows)                    # C18
    span = (end - start).days
    if start.year == 2024:
        k_ = (366 / span * (YEAR_2024_END - start).days / span
              + 365 / span * (end - YEAR_2024_END).days / span)
        simple_yield = gross / amount * k_
        final_yield = (payout / amount - 1) * k_
    else:
        simple_yield = gross / amount * 365 / span
        final_yield = (payout / amount - 1) * 365 / span

    return Result(
        start_date=start, end_date=end, amount=amount, placement_rate=rate,
        mf=mf, uf=inp.uf, rf=inp.rf, mf_total=mf_total,
        simple_yield=simple_yield, gross_income=gross,
        final_payout=payout, final_yield=final_yield,
        survival_sum=(0.001 * inp.premium) + (inp.initial_investment * 0.001),
        accident_sum=0.1 * inp.premium,
        solver_rates=solver_rates, solver_exact=solver_exact, rows=rows,
    )


if __name__ == "__main__":
    import argparse

    p = argparse.ArgumentParser(description="Спецтранши 2.0 — PB Top Up")
    p.add_argument("--term", type=int, default=6, choices=TERMS, help="срок, мес.")
    p.add_argument("--premium", type=float, default=100_000_000)
    p.add_argument("--ukv", type=float, default=0.0013, choices=UKV_OPTIONS)
    p.add_argument("--start", type=date.fromisoformat, help="дата начала (по умолчанию завтра)")
    a = p.parse_args()
    r = calculate(Inputs(term_months=a.term, premium=a.premium, ukv=a.ukv, start_date=a.start))
    print(f"Период: {r.start_date} — {r.end_date}, ставка размещения {r.placement_rate:.4%}")
    for k, v in r.main_page.items():
        print(f"{k:40s} {v:,.2f}" if isinstance(v, float) and abs(v) > 1 else f"{k:40s} {v}")
