import os, json, re
import anthropic

MODEL = "claude-sonnet-4-6"

class ClaudeAgent:

    def __init__(self, push_event):
        self.emit   = push_event
        self.client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])

    def diagnose_and_fix(self, calc_errors, build_source, engine_source):
        errors_summary = "\n".join(
            f"Check {k}: {r['label']}\n  Diagnosis: {r['diagnosis']}\n  Note: {r['note']}"
            for k, r in calc_errors.items()
        )

        prompt = f"""You are fixing a supply chain planning system.

These checks FAILED with CALCULATION ERRORS:

{errors_summary}

build_planner.py:
<build_planner>
{build_source}
</build_planner>

checklist_engine.py:
<checklist_engine>
{engine_source}
</checklist_engine>

Return ONLY a JSON array of fixes:
[{{"file": "build_planner.py", "old": "exact string to replace", "new": "replacement", "description": "what was wrong"}}]

Rules:
- "old" must be an exact verbatim substring of the file
- Make the smallest possible change
- Return ONLY the JSON array, no other text
"""
        self.emit({"type": "claude", "message": "Asking Claude to diagnose and fix..."})

        response = self.client.messages.create(
            model=MODEL,
            max_tokens=4096,
            messages=[{"role": "user", "content": prompt}],
        )

        raw = response.content[0].text.strip()
        raw = re.sub(r"^```json\s*|^```\s*|```$", "", raw, flags=re.MULTILINE).strip()

        try:
            fixes = json.loads(raw)
            self.emit({"type": "claude", "message": f"Claude returned {len(fixes)} fix(es)"})
            return fixes
        except json.JSONDecodeError:
            self.emit({"type": "claude", "message": "Claude response was not valid JSON"})
            return []

    def write_new_check(self, description, threshold, fail_type, engine_source):
        existing_nums = re.findall(r"results\[(\d+)\]", engine_source)
        next_num = max(int(n) for n in existing_nums) + 1 if existing_nums else 16

        prompt = f"""You are adding a new check to a supply chain planning checklist engine.

Current checklist_engine.py:
<engine>
{engine_source}
</engine>

New check to add:
  Description: {description}
  Threshold: {threshold}
  Failure type: {fail_type}
  Check number: {next_num}

Study the existing checks and follow the EXACT same pattern.
Store result in results[{next_num}].
Return ONLY the Python code block to insert, no markdown fences.
"""
        self.emit({"type": "claude", "message": f"Writing new check #{next_num}..."})

        response = self.client.messages.create(
            model=MODEL,
            max_tokens=2048,
            messages=[{"role": "user", "content": prompt}],
        )

        code = response.content[0].text.strip()
        code = re.sub(r"^```python\s*|^```\s*|```$", "", code, flags=re.MULTILINE).strip()

        self.emit({"type": "claude", "message": f"Check #{next_num} written"})
        return code, next_num
