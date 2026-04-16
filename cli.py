import argparse
import json
from scheduler.scheduler import Scheduler

def main():
    parser = argparse.ArgumentParser(description="Harmonogram generator")

    parser.add_argument("--year", type=int, required=True)
    parser.add_argument("--month", type=int, required=True)
    parser.add_argument("--seed", type=int, default=None)
    parser.add_argument("--out", type=str, default=None)

    # 🔴 NAJWAŻNIEJSZE
    parser.add_argument("--config", type=str, required=True,
                        help="ścieżka do pliku JSON z danymi")

    args = parser.parse_args()

    # 🔴 WCZYTANIE CONFIGA
    try:
        with open(args.config, "r", encoding="utf-8") as f:
            data = json.load(f)
    except FileNotFoundError:
        print(f"❌ Nie znaleziono pliku: {args.config}")
        return
    except json.JSONDecodeError:
        print("❌ Błąd JSON w pliku config")
        return

    # 🔴 WYCIĄGANIE DANYCH
    initial_stats = data.get("initial_stats")
    last_weekend_workers = data.get("last_weekend_workers")
    # DODANO: Pobieranie urlopów (jeśli nie ma w pliku, dajemy pusty słownik)
    leaves = data.get("leaves", {})

    # 🔴 WALIDACJA
    if initial_stats is None:
        print("❌ Brak 'initial_stats' w configu")
        return

    if last_weekend_workers is None:
        print("❌ Brak 'last_weekend_workers' w configu")
        return

    # 🔴 START
    sched = Scheduler(seed=args.seed)

    # DODANO: Przekazanie parametru leaves do metody
    sched.generate_and_save(
        year=args.year,
        month=args.month,
        employees=None,
        out_filename=args.out,
        initial_stats=initial_stats,
        last_weekend_workers=last_weekend_workers,
        leaves=leaves
    )

if __name__ == '__main__':
    main()
