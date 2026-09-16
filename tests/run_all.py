"""Run every offline test suite with the bundled interpreter."""

import subprocess
import sys
from pathlib import Path

TESTS_DIR = Path(__file__).resolve().parent
ROOT = TESTS_DIR.parent


def main() -> int:
    failures: list[str] = []
    for script in sorted(TESTS_DIR.glob("tests_offline_*.py")):
        proc = subprocess.run(
            [sys.executable, str(script)],
            cwd=ROOT,
            capture_output=True,
            text=True,
        )
        output = proc.stdout + proc.stderr
        result = "PASS" if proc.returncode == 0 and "RESULT= PASS" in output else "FAIL"
        if result == "FAIL":
            failures.append(script.name)
            print(output.strip())
        print(f"{result} {script.name}")
    print("RESULT=", "PASS" if not failures else f"FAIL {failures}")
    return 0 if not failures else 1


if __name__ == "__main__":
    sys.exit(main())
