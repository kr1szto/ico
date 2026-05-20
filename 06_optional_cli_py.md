# Optional `cli.py`

This file is optional. Use it only if you want a command-line input method in addition to the web GUI.

Copy this into `cli.py`.

```python
import argparse
import json
import re
from pathlib import Path

from main import scrape_subject


def normalize_ico(value: str) -> str:
    ico = re.sub(r"\D", "", value or "")

    if len(ico) != 8:
        raise ValueError("IČO musí mať presne 8 číslic.")

    return ico


def main():
    parser = argparse.ArgumentParser(description="Supplier lookup by IČO")
    parser.add_argument("--ico", required=True, help="IČO, napr. 36785512 alebo '36 785 512'")
    parser.add_argument("--out", default="output", help="Output folder")

    args = parser.parse_args()

    ico = normalize_ico(args.ico)
    output_dir = Path(args.out)
    output_dir.mkdir(exist_ok=True)

    result = scrape_subject(ico)

    output_path = output_dir / f"{ico}_result.json"

    with output_path.open("w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(json.dumps(result, ensure_ascii=False, indent=2))
    print(f"\nVýsledok uložený do: {output_path}")


if __name__ == "__main__":
    main()
```

Run:

```bash
docker compose run --rm scraper python cli.py --ico "36 785 512"
```
