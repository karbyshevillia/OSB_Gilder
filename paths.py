from pathlib import Path
import sys

if getattr(sys, "frozen", False):
    PROJECT_ROOT = Path(sys._MEIPASS)
else:
    PROJECT_ROOT = Path(__file__).resolve().parent

IND_FILE = PROJECT_ROOT / "ind.csv"

if __name__ == '__main__':
    print(str(IND_FILE))
