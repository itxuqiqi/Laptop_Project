# state.py
from dataclasses import dataclass, field
from typing import Dict, Any, Optional, Set

@dataclass
class AppState:
    # key -> value: file path 或 number
    sources: Dict[str, Any] = field(default_factory=dict)

    # 勾选的 sheet 名集合
    selected_sheets: Set[str] = field(default_factory=set)

    def set_source(self, key: str, value: Any):
        self.sources[key] = value

    def get_source(self, key: str, default=None):
        return self.sources.get(key, default)

    def clear_source(self, key: str):
        if key in self.sources:
            del self.sources[key]
