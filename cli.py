import argparse
from scheduler.scheduler import Scheduler

def main():
    parser = argparse.ArgumentParser(description="Harmonogram generator")
    parser.add_argument("--year", type=int, default=2026)
    parser.add_argument("--month", type=int, default=2)
    parser.add_argument("--seed", type=int, default=None)
    parser.add_argument("--out", type=str, default=None)
    args = parser.parse_args()
    sched = Scheduler(seed=args.seed)
    employees = None  # uses default list inside Scheduler if None
    sched.generate_and_save(year=args.year, month=args.month, employees=employees, out_filename=args.out)

if __name__ == '__main__':
    main()