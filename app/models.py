from dataclasses import dataclass
from enum import Enum, auto

@dataclass
class IngredientRow:
    """
    Represents a single row in the ingredient table.
    """
    rm_name: str
    rm_percent: str  # Kept as string to preserve user input formatting
    inci_name: str
    inci_percent: str

    def __post_init__(self):
        # Ensure values are never None
        self.rm_name = self.rm_name or ""
        self.rm_percent = self.rm_percent or ""
        self.inci_name = self.inci_name or ""
        self.inci_percent = self.inci_percent or ""


class DiffType(Enum):
    NONE = auto()
    CONTENT_MISMATCH = auto()  # Red Font (Value diff)
    MISSING_ROW = auto()       # Red Background (Row missing)
    MISSING_INCI = auto()      # Red Background (INCI missing)

@dataclass
class DiffItem:
    """
    Represents a specific styling instruction for a cell.
    """
    row: int
    col: int
    diff_type: DiffType

