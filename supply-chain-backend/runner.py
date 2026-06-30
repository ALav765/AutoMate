import os, importlib.util
from pathlib import Path

MAX_FIX_LOOPS = 5
BACKEND_DIR      = Path(__file__).parent
BUILD_SCRIPT     = BACKEND_DIR / "build_planner.py"
CHECKLIST_SCRIPT = BACKEND_DIR / "checklist_engine.py"

class PipelineRunner:

    def __init__(self, job_id, push_event):
        self.job_id  = job_id
        self.emit    = push_event
        self.job_dir = Path("jobs") / job_id

    def run(self, files, month_label):
        output_path = str(self.job_dir / "Integrated_Supply_Plan.xlsx")

        self.emit({"type": "step", "step": "build", "message": "Building integrated supply plan..."})
        self._run_build(files, output_path)
        self.emit({"type": "step_done", "step": "build", "message": "Plan built"})

        summary = self._check_and_fix_loop(output_path, files)
        return output_path, summary

    def _run_build(self, files, output_path):
        env_patch = {
            "FORECAST_FILE": files["forecast"],
            "BOM_FILE":      files["bom"],
            "COMP_SOH_FILE": files["component_soh"],
            "PRICES_FILE":   files["prices"],
            "VENDOR_FILE":   files["vendor"],
            "OUTPUT_FILE":   output_path,
        }
        old_env = {k: os.environ.get(k) for k in env_patch}
        try:
            os.environ.update(env_patch)
            self._exec_script(BUILD_SCRIPT)
        finally:
            for k, v in old_env.items():
                if v is None: os.environ.pop(k, None)
                else:         os.environ[k] = v

    def _check_and_fix_loop(self, output_path, files):
        from claude_agent import ClaudeAgent
        agent = ClaudeAgent(push_event=self.emit)

        results = {}
        loop = 0

        while loop < MAX_FIX_LOOPS:
            loop += 1
            self.emit({"type": "step", "step": "check", "message": f"Running checklist (loop {loop})..."})
            results = self._run_checks(output_path)

            passed = sum(1 for r in results.values() if r["passed"])
            total  = len(results)

            self.emit({
                "type": "check_results", "loop": loop,
                "passed": passed, "total": total,
                "results": {
                    str(k): {"passed": r["passed"], "label": r["label"],
                             "diagnosis": r["diagnosis"], "note": r["note"]}
                    for k, r in results.items()
                }
            })

            calc_errors = {k: r for k, r in results.items()
                          if not r["passed"] and "CALC ERROR" in r.get("diagnosis", "")}

            if not calc_errors:
                self.emit({"type": "step_done", "step": "check", "message": "All calc checks clean"})
                break

            self.emit({"type": "step", "step": "fix", "message": f"Fixing {len(calc_errors)} error(s)..."})
            fixes = agent.diagnose_and_fix(
                calc_errors=calc_errors,
                build_source=BUILD_SCRIPT.read_text(),
                engine_source=CHECKLIST_SCRIPT.read_text(),
            )

            if not fixes:
                self.emit({"type": "step_done", "step": "fix", "message": "No fixes returned — flagging as-is"})
                break

            for fix in fixes:
                self._apply_fix(fix)
                self.emit({"type": "fix_applied", "message": fix.get("description", "Fix applied")})

            self.emit({"type": "step", "step": "rebuild", "message": "Rebuilding with fixes..."})
            self._run_build(files, output_path)
            self.emit({"type": "step_done", "step": "rebuild", "message": "Rebuilt"})

        self.emit({"type": "step", "step": "alerts", "message": "Writing alerts sheet..."})
        self._write_alerts(output_path, results)
        self.emit({"type": "step_done", "step": "alerts", "message": "Done"})

        genuine = {k: r for k, r in results.items()
                   if not r["passed"] and "CALC ERROR" not in r.get("diagnosis", "")}
        calc_err = {k: r for k, r in results.items()
                    if not r["passed"] and "CALC ERROR" in r.get("diagnosis", "")}

        return {
            "passed": sum(1 for r in results.values() if r["passed"]),
            "total":  len(results),
            "calc_errors": len(calc_err),
            "business_alerts": len(genuine),
            "loops": loop,
        }

    def _run_checks(self, output_path):
        old = os.environ.get("PLAN_FILE")
        try:
            os.environ["PLAN_FILE"] = output_path
            spec   = importlib.util.spec_from_file_location("checklist_engine", CHECKLIST_SCRIPT)
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)
            return module.run_checks()
        finally:
            if old is None: os.environ.pop("PLAN_FILE", None)
            else:           os.environ["PLAN_FILE"] = old

    def _write_alerts(self, output_path, results):
        old = os.environ.get("PLAN_FILE")
        try:
            os.environ["PLAN_FILE"] = output_path
            spec   = importlib.util.spec_from_file_location("checklist_engine", CHECKLIST_SCRIPT)
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)
            module.write_alerts_sheet(results)
        finally:
            if old is None: os.environ.pop("PLAN_FILE", None)
            else:           os.environ["PLAN_FILE"] = old

    def _apply_fix(self, fix):
        target = BUILD_SCRIPT if fix["file"] == "build_planner.py" else CHECKLIST_SCRIPT
        source = target.read_text()
        if fix["old"] not in source:
            raise ValueError(f"Fix target not found in {fix['file']}")
        target.write_text(source.replace(fix["old"], fix["new"], 1))

    def _exec_script(self, path):
        spec   = importlib.util.spec_from_file_location(path.stem, path)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
