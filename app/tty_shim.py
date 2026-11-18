# (c) 2026 Alain Yon
"""Run a tool script, unchanged, with stdin reporting itself as a console.

Two tool scripts (Theoretical CoS, Cellar Consumption) only ask their
on-screen questions when sys.stdin.isatty() is true, which it is when the .bat
runs them in a console window. Under the app stdin is a pipe, so without this
they would silently skip every question. This shim swaps sys.stdin for an
identical text stream whose isatty() says True, then runs the script file as
__main__ exactly as "python script.py args" would. The script is not modified
and not imported as a module.

Usage:  python tty_shim.py <script.py> [args...]
"""
import io
import os
import runpy
import sys


class _ConsoleLikeStdin(io.TextIOWrapper):
    def isatty(self):
        return True


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    script = sys.argv[1]
    sys.argv = sys.argv[1:]
    old = sys.stdin
    sys.stdin = _ConsoleLikeStdin(old.buffer, encoding=old.encoding, errors=old.errors,
                                  line_buffering=True)
    sys.path[0] = os.path.dirname(os.path.abspath(script))
    runpy.run_path(script, run_name="__main__")
    return 0


if __name__ == "__main__":
    sys.exit(main())
