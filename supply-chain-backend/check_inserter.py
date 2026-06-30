"""
Inserts a newly-generated check into checklist_engine.py and verifies it
actually runs before treating it as live.

Safety approach:
  1. Read current checklist_engine.py source.
  2. Insert the new check's code right before the "return results" line
     inside run_checks().
  3. Write the modified file to disk.
  4. Re-import the module fresh and call run_checks() against the most
     recent plan output.
  5. If it raises, or the new check_num is missing from the results dict,
     ROLL BACK the file to its original content and report the error.
  6. If it succeeds, the file stays modified — the check is now live.
"""

import importlib.util
import re
from pathlib import Path

CHECKLIST_SCRIPT = Path(__file__).parent / "checklist_engine.py"
INSERT_MARKER = "    # ── New checks get appended here by claude_agent.py ───────────────────"


def insert_and_validate_check(code: str, check_num: int, push_event=None) -> tuple[bool, str]:
    """
    Returns (success: bool, message: str).
    On failure, checklist_engine.py is left unchanged.
    """
    emit = push_event or (lambda e: None)
    original_source = CHECKLIST_SCRIPT.read_text()

    if INSERT_MARKER not in original_source:
        return False, (
            "Could not find the insertion marker in checklist_engine.py. "
            "Has the file been edited manually since the last check was added?"
        )

    # Indent Claude's code to match the function body (4 spaces), in case it
    # came back without consistent indentation.
    indented_code = "\n".join(
        line if line.strip() == "" else f"    {line.lstrip()}"
        for line in code.strip().splitlines()
    )

    new_source = original_source.replace(
        INSERT_MARKER,
        f"{indented_code}\n\n{INSERT_MARKER}",
    )

    CHECKLIST_SCRIPT.write_text(new_source)
    emit({"type": "claude", "message": f"Inserted check #{check_num} — validating..."})

    try:
        spec = importlib.util.spec_from_file_location("checklist_engine_validate", CHECKLIST_SCRIPT)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        results = module.run_checks()

        if check_num not in results:
            raise ValueError(f"run_checks() completed but check #{check_num} is missing from results")

        emit({"type": "claude", "message": f"Check #{check_num} validated successfully"})
        return True, "Check added and validated."

    except Exception as e:
        # Roll back — do not leave broken code live in the pipeline.
        CHECKLIST_SCRIPT.write_text(original_source)
        emit({"type": "error", "message": f"New check failed validation, rolled back: {e}"})
        return False, f"Check failed validation and was not saved: {e}"
