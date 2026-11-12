import datetime
import json
import os


class AgentMemoryManager:
    """Keeps a local log of completed research runs (memory/learnings.json).

    It no longer stores valuation overrides: earlier versions saved a terminal
    growth value here that silently replaced the per-run assumption. Valuation
    inputs now live only in config/market.json.
    """

    def __init__(self, memory_dir=None):
        self.memory_dir = memory_dir or os.path.dirname(os.path.abspath(__file__))
        os.makedirs(self.memory_dir, exist_ok=True)
        self.memory_file = os.path.join(self.memory_dir, "learnings.json")
        self.data = self._load_memory()

    def _default_memory(self):
        return {
            "version": "2.0.0",
            "description": "Research history log. Valuation inputs live in config/market.json.",
            "research_history": [],
        }

    def _load_memory(self):
        if not os.path.exists(self.memory_file):
            data = self._default_memory()
            self._save_memory(data)
            return data
        try:
            with open(self.memory_file, "r", encoding="utf-8") as f:
                return json.load(f)
        except (OSError, ValueError):
            return self._default_memory()

    def _save_memory(self, data=None):
        with open(self.memory_file, "w", encoding="utf-8") as f:
            json.dump(self.data if data is None else data, f, indent=2, ensure_ascii=False)

    def log_research_run(self, ticker, name, cmp, target, verdict, margin_of_safety):
        """Append a completed run (keeps the last 50)."""
        self.data.setdefault("research_history", []).append({
            "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "ticker": ticker.upper(),
            "name": name,
            "cmp": float(cmp),
            "target": float(target),
            "verdict": verdict,
            "margin_of_safety": margin_of_safety,
        })
        self.data["research_history"] = self.data["research_history"][-50:]
        self._save_memory()
