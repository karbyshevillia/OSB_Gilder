import openpyxl as xl
import re
from dataclasses import dataclass
from typing import Sequence
from enum import Enum

class IndicatorKind(Enum):
    PRIMARY = "PRIMARY"
    SECONDARY = "SECONDARY"

class Indicator:
    def __init__(self, skeleton: str):
        self.regex = re.compile(r"\(\)")
        self.skeleton = skeleton
        self.cells = ()

    # def insert(self, *cells):
    #     self.cells = cells
    #     return self

    def render(self, *cells):
        self.cells = cells
        cell_iter = iter(self.cells)

        def replacer(match):
            try:
                return next(cell_iter)
            except StopIteration:
                raise ValueError("Not enough cells provided")

        result = self.regex.sub(replacer, self.skeleton)

        # Check for excess arguments
        try:
            next(cell_iter)
            raise ValueError("Too many cells provided")
        except StopIteration:
            pass

        return result

@dataclass(frozen=True)
class LogicalIndicator:
    name: str
    kind: IndicatorKind
    skeleton: str
    slots: int
    inputs: Sequence[str]
    mask: Sequence[str]
    associated_codes: Sequence[str]


if __name__ == '__main__':
    ind = Indicator("Test", "=()+()")
    ind.insert("A1", "A2")
    print(ind.render())