"""
Example script — delete it once you've dropped your own in here.

It exists so the workflow has something to build on a fresh clone, and so you
can see what a script that survives being frozen looks like: no dependencies,
an explicit entry point, and a --version flag the build can smoke-test against.
"""

import argparse
import platform
import sys

VERSION = "1.0.0"


def main() -> int:
    parser = argparse.ArgumentParser(description="Example py-to-exe script")
    parser.add_argument("--version", action="store_true", help="print the version and exit")
    parser.add_argument("--name", default="world")
    args = parser.parse_args()

    if args.version:
        print(VERSION)
        return 0

    print(f"Hello, {args.name}!")
    print(f"  python  {platform.python_version()}")
    print(f"  running {'as a frozen exe' if getattr(sys, 'frozen', False) else 'as a .py'}")
    print(f"  from    {sys.executable}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
