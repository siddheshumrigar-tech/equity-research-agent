import os
import json
import datetime

class AgentMemoryManager:
    """
    Autonomous Persistent Memory Manager for Equity Research Agent.
    Enables continuous learning across research runs by persisting sector nuances,
    ticker-specific calibrations, and user preferences into a clean local JSON database.
    Starts from a clean slate (0) on new installations.
    """
    def __init__(self, memory_dir=None):
        if memory_dir is None:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            self.memory_dir = base_dir
        else:
            self.memory_dir = memory_dir
            
        os.makedirs(self.memory_dir, exist_ok=True)
        self.memory_file = os.path.join(self.memory_dir, "learnings.json")
        self.data = self._load_memory()

    def _default_memory(self):
        return {
            "version": "1.0.0",
            "description": "Autonomous learning bank for Equity Research Agent. Stores persistent sector nuances, model calibrations, and preferences across runs.",
            "sector_overrides": {},
            "ticker_calibrations": {},
            "user_preferences": {
                "default_currency": "INR",
                "wacc_baseline_rf": 0.071,
                "terminal_growth_g": 0.050,
                "preferred_model": "Tesla-Executive-10Tab"
            },
            "research_history": []
        }

    def _load_memory(self):
        if not os.path.exists(self.memory_file):
            defaults = self._default_memory()
            self._save_memory(defaults)
            return defaults
        try:
            with open(self.memory_file, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return self._default_memory()

    def _save_memory(self, data=None):
        if data is None:
            data = self.data
        with open(self.memory_file, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

    def get_sector_override(self, sector_name):
        """Retrieve any learned overrides for a given industry sector."""
        return self.data.get("sector_overrides", {}).get(sector_name.upper(), {})

    def get_ticker_calibration(self, ticker):
        """Retrieve past learned parameters or valuation adjustments for a specific ticker."""
        return self.data.get("ticker_calibrations", {}).get(ticker.upper(), {})

    def get_preference(self, key, default=None):
        """Retrieve user configuration preference."""
        return self.data.get("user_preferences", {}).get(key, default)

    def record_learning(self, category, key, value_dict):
        """
        Record a permanent rule or calibration into memory.
        Categories: 'sector_overrides', 'ticker_calibrations', 'user_preferences'
        """
        if category not in self.data:
            self.data[category] = {}
        if isinstance(value_dict, dict):
            if key not in self.data[category]:
                self.data[category][key] = {}
            self.data[category][key].update(value_dict)
        else:
            self.data[category][key] = value_dict
        self._save_memory()
        print(f"🧠 [Memory] Learned new rule for {category} -> {key}: {value_dict}")

    def log_research_run(self, ticker, name, cmp, target, verdict, margin_of_safety):
        """Log a successful valuation execution to research history."""
        run_record = {
            "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "ticker": ticker.upper(),
            "name": name,
            "cmp": float(cmp),
            "target": float(target),
            "verdict": verdict,
            "margin_of_safety": margin_of_safety
        }
        if "research_history" not in self.data:
            self.data["research_history"] = []
            
        # Keep last 50 runs to prevent file bloat
        self.data["research_history"].append(run_record)
        if len(self.data["research_history"]) > 50:
            self.data["research_history"] = self.data["research_history"][-50:]
        self._save_memory()
        print(f"📝 [Memory] Logged research run for {ticker} (Verdict: {verdict})")
