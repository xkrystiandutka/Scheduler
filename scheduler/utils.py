# scheduler/utils.py
import calendar
from datetime import date, timedelta

def month_days(year, month):
    _, ndays = calendar.monthrange(year, month)
    return [date(year, month, d) for d in range(1, ndays+1)]

def easter_date(year):
    a = year % 19
    b = year // 100
    c = year % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19*a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2*e + 2*i - h - k) % 7
    m = (a + 11*h + 22*l) // 451
    month = (h + l - 7*m + 114) // 31
    day = ((h + l - 7*m + 114) % 31) + 1
    return date(year, month, day)

def polish_holidays(year):
    fixed = [(1,1),(1,6),(5,1),(5,3),(8,15),(11,1),(11,11),(12,25),(12,26)]
    hol = set(date(year,m,d) for (m,d) in fixed)
    e = easter_date(year)
    hol.add(e + timedelta(days=1)) # Easter Monday
    hol.add(e + timedelta(days=60)) # Corpus Christi (approx)
    return hol
