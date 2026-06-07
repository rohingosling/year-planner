#-----------------------------------------------------------------------------------------------------------------------
# Program: Year Planner Generator — Build
# Version: 1.2
# Author:  Rohin Gosling
#
# Description:
#
#   Builds the standalone Windows executable (dist/year-planner.exe) via PyInstaller, then smoke-tests it. The build
#   stamps the changelog version and today's build date into src/main.py, compiles using year-planner.spec, and runs
#   the frozen exe to confirm it produces a non-empty DOCX. The smoke test needs no Microsoft Word (DOCX only).
#
# Usage:
#
#   .venv/Scripts/python.exe scripts/build.py
#
#   Exit codes:
#   - 0 : Build and smoke test passed.
#   - 1 : A build step failed.
#-----------------------------------------------------------------------------------------------------------------------

import os
import re
import subprocess
import sys
import tempfile
from datetime import date

# Global constants.

PROJECT_DIR      = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))      # Project root (parent of scripts/).
MAIN_PY          = os.path.join(PROJECT_DIR, "src", "main.py")                      # Entry module that is stamped.
CHANGELOG        = os.path.join(PROJECT_DIR, "docs", "CHANGELOG.md")                # Source of the version number.
SPEC_FILE        = os.path.join(PROJECT_DIR, "year-planner.spec")                   # PyInstaller spec.
EXE_PATH         = os.path.join(PROJECT_DIR, "dist", "year-planner.exe")            # Expected build output.
SMOKE_TEST_YEAR  = 2030                                                             # Year used for the smoke test.
BANNER_WIDTH     = 60                                                               # Banner box width in characters.

# Configure stdout for UTF-8 so the banner's box-drawing characters render on Windows.

if sys.stdout is not None and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


#-----------------------------------------------------------------------------------------------------------------------
# Function: print_banner
#
# Description:
#
#   Print text inside a box-drawing banner.
#
# Arguments:
#
#   text : The text to display inside the banner.
#
# Returns:
#
#   None.
#-----------------------------------------------------------------------------------------------------------------------

def print_banner(text: str) -> None:

    # Print text inside a box-drawing banner.

    width = max(BANNER_WIDTH, len(text) + 4)
    inner = width - 2

    print("┌" + "─" * inner + "┐")
    print("│ " + text.ljust(inner - 2) + " │")
    print("└" + "─" * inner + "┘")


#-----------------------------------------------------------------------------------------------------------------------
# Function: read_changelog_version
#
# Description:
#
#   Return the highest versioned heading from docs/CHANGELOG.md, scanning for "## [X.Y]" and skipping "[Unreleased]".
#
# Arguments:
#
#   None.
#
# Returns:
#
#   The version string, or None when no versioned heading is found.
#-----------------------------------------------------------------------------------------------------------------------

def read_changelog_version() -> str | None:

    # Return the highest versioned heading from docs/CHANGELOG.md.

    try:
        with open(CHANGELOG, "r", encoding="utf-8") as changelog_file:
            for line in changelog_file:
                match = re.match(r"^##\s+\[(\d+\.\d+)\]", line)
                if match:

                    # Return data to caller.

                    return match.group(1)
    except OSError:
        pass

    # Return data to caller.

    return None


#-----------------------------------------------------------------------------------------------------------------------
# Function: step_stamp_version
#
# Description:
#
#   Stamp the changelog version into the _changelog_version() fallback in src/main.py so the frozen exe reports the
#   correct version without the (unbundled) changelog. When docs/CHANGELOG.md is absent (e.g. a public/staging build,
#   where the changelog is intentionally not shipped), this is a no-op: the version already stamped in src/main.py
#   stands, so the build proceeds instead of failing.
#
# Arguments:
#
#   None.
#
# Returns:
#
#   A (success, summary) tuple.
#-----------------------------------------------------------------------------------------------------------------------

def step_stamp_version() -> tuple[bool, str]:

    # Stamp the changelog version into the _changelog_version() fallback in src/main.py.

    version = read_changelog_version()

    if version is None:

        # docs/CHANGELOG.md is not present (e.g. a public/staging build, where the changelog is intentionally not
        # shipped). The version already stamped into src/main.py's _changelog_version() fallback stands as-is, so a
        # missing changelog is a graceful skip rather than a hard build failure.

        return True, "docs/CHANGELOG.md not present - keeping version already stamped in src/main.py"

    with open(MAIN_PY, "r", encoding="utf-8") as main_file:
        content = main_file.read()

    pattern = r'(_changelog_version\(")[^"]*("\))'
    new_content, count = re.subn(pattern, rf"\g<1>{version}\2", content)

    if count == 0:
        return False, "Could not find _changelog_version() fallback in src/main.py"

    if new_content != content:
        with open(MAIN_PY, "w", encoding="utf-8") as main_file:
            main_file.write(new_content)
        return True, f'Stamped VERSION fallback = "{version}"'

    # Return data to caller.

    return True, f'VERSION fallback already "{version}" - no change needed'


#-----------------------------------------------------------------------------------------------------------------------
# Function: step_stamp_build_date
#
# Description:
#
#   Stamp today's date into the BUILD_DATE constant in src/main.py.
#
# Arguments:
#
#   None.
#
# Returns:
#
#   A (success, summary) tuple.
#-----------------------------------------------------------------------------------------------------------------------

def step_stamp_build_date() -> tuple[bool, str]:

    # Stamp today's date into the BUILD_DATE constant in src/main.py.

    today = date.today().isoformat()

    with open(MAIN_PY, "r", encoding="utf-8") as main_file:
        content = main_file.read()

    pattern = r'(BUILD_DATE\s*=\s*")[^"]*(")'
    new_content, count = re.subn(pattern, rf"\g<1>{today}\2", content)

    if count == 0:
        return False, "Could not find BUILD_DATE in src/main.py"

    if new_content != content:
        with open(MAIN_PY, "w", encoding="utf-8") as main_file:
            main_file.write(new_content)
        return True, f'Stamped BUILD_DATE = "{today}"'

    # Return data to caller.

    return True, f'BUILD_DATE already "{today}" - no change needed'


#-----------------------------------------------------------------------------------------------------------------------
# Function: step_build_exe
#
# Description:
#
#   Build the one-file executable from year-planner.spec via PyInstaller.
#
# Arguments:
#
#   None.
#
# Returns:
#
#   A (success, summary) tuple.
#-----------------------------------------------------------------------------------------------------------------------

def step_build_exe() -> tuple[bool, str]:

    # Build the one-file executable from year-planner.spec via PyInstaller.

    result = subprocess.run(
        [sys.executable, "-m", "PyInstaller", "--noconfirm", SPEC_FILE],
        capture_output=True, text=True, cwd=PROJECT_DIR
    )

    if result.returncode == 0 and os.path.isfile(EXE_PATH):
        return True, f"Built {EXE_PATH}"

    stderr_tail = result.stderr.strip().splitlines()[-3:] if result.stderr else []

    # Return data to caller.

    return False, f"PyInstaller failed (exit {result.returncode}): {' '.join(stderr_tail)}"


#-----------------------------------------------------------------------------------------------------------------------
# Function: step_smoke_test_exe
#
# Description:
#
#   Run the frozen exe in a temporary directory and confirm it produces a non-empty DOCX. No Microsoft Word required.
#
# Arguments:
#
#   None.
#
# Returns:
#
#   A (success, summary) tuple.
#-----------------------------------------------------------------------------------------------------------------------

def step_smoke_test_exe() -> tuple[bool, str]:

    # Run the frozen exe in a temporary directory and confirm it produces a non-empty DOCX.

    if not os.path.isfile(EXE_PATH):
        return False, f"{EXE_PATH} not found"

    with tempfile.TemporaryDirectory() as temp_dir:

        env = {**os.environ, "PYTHONIOENCODING": "utf-8"}

        result = subprocess.run(
            [EXE_PATH, "--year", str(SMOKE_TEST_YEAR)],
            capture_output=True, text=True, encoding="utf-8", errors="replace", cwd=temp_dir, env=env
        )

        if result.returncode != 0:
            return False, f".exe exited with code {result.returncode}"

        output_docx = os.path.join(temp_dir, f"year-planner-{SMOKE_TEST_YEAR}.docx")

        if not os.path.isfile(output_docx):
            return False, ".exe ran but did not produce the expected DOCX"

        size = os.path.getsize(output_docx)

        if size == 0:
            return False, ".exe produced an empty DOCX"

        # Return data to caller.

        return True, f".exe produced a {size}-byte DOCX"


#-----------------------------------------------------------------------------------------------------------------------
# Function: main
#
# Description:
#
#   Run the build pipeline: stamp version, stamp build date, build the exe, smoke-test it, and report.
#
# Arguments:
#
#   None.
#
# Returns:
#
#   Exit code (0 = all steps passed, 1 = a step failed).
#-----------------------------------------------------------------------------------------------------------------------

def main() -> int:

    # Run the build pipeline.

    print_banner("Year Planner Generator - Build")
    print()

    steps = [
        ("Stamping version from docs/CHANGELOG.md", step_stamp_version),
        ("Stamping build date",                     step_stamp_build_date),
        ("Building .exe via year-planner.spec",     step_build_exe),
    ]

    all_passed = True

    # Run the stamp and build steps in order, stopping the build from being smoke-tested if it failed.

    for index, (label, step) in enumerate(steps, start=1):
        print(f"[{index}/4] {label}...")
        ok, summary = step()
        print(f"      {'PASS' if ok else 'FAIL'}: {summary}")
        print()
        all_passed = all_passed and ok

    # Smoke-test only if the build succeeded.

    print("[4/4] Smoke-testing .exe...")
    if all_passed:
        ok_smoke, summary = step_smoke_test_exe()
        print(f"      {'PASS' if ok_smoke else 'FAIL'}: {summary}")
        all_passed = all_passed and ok_smoke
    else:
        print("      SKIP: a prior step failed")
    print()

    # Report the final result.

    print_banner("Build successful." if all_passed else "Build FAILED.")

    # Return data to caller.

    return 0 if all_passed else 1


if __name__ == "__main__":
    sys.exit(main())
