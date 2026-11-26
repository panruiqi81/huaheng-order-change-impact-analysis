"""订单变更影响核对助手（只读原型）。"""

from .constants import RULES_VERSION
from .models import load_snapshot
from .rules import evaluate

__all__ = ["RULES_VERSION", "load_snapshot", "evaluate"]
