import pandas as pd
import re
from string import ascii_lowercase as alphabets


class _ExcelIndexer:
    def __init__(self, df):
        self.df = df

    def _get_index(self, str_idx):
        idx = -1
        for i, c in enumerate(str_idx):
            idx += (alphabets.index(c.lower()) + 1) * 26 ** (len(str_idx) - i - 1)
        return idx

    def _extract_index(self, str_idx):
        idx_dict = re.match(
            "^(?P<start_col>[A-Z]+)(?P<start_row>[0-9]+)?(:(?P<end_col>[A-Z]+)(?P<end_row>[0-9]+)?)?$",
            str_idx,
            re.IGNORECASE,
        ).groupdict()

        start_col, start_row, end_col, end_row = None, None, None, None
        n_rows, n_cols = self.df.shape

        if idx_dict["start_col"]:
            start_col = self._get_index(idx_dict["start_col"])
            if start_col >= n_cols or start_col < 0:
                raise IndexError("Column Index out-of-bound")
        if idx_dict["end_col"]:
            end_col = self._get_index(idx_dict["end_col"])
            if end_col >= n_cols or end_col < 0:
                raise IndexError("Column Index out-of-bound")

        if idx_dict["start_row"]:
            start_row = int(idx_dict["start_row"])
            if start_row > n_rows or start_row < 0:
                raise IndexError("Row Index out-of-bound")
        if idx_dict["end_row"]:
            end_row = int(idx_dict["end_row"])
            if end_row > n_rows or end_row < 0:
                raise IndexError("Row Index out-of-bound")

        return start_col, start_row, end_col, end_row

    def _is_valid(self, idx):
        if type(idx) == str:
            match = re.match(
                "^(?P<start_col>[A-Z]+)(?P<start_row>[0-9]+)?(:(?P<end_col>[A-Z]+)(?P<end_row>[0-9]+)?)?$",
                idx,
                re.IGNORECASE,
            )
            return match is not None
        else:
            return False

    def __getitem__(self, idx):
        if not self._is_valid(idx):
            raise Exception("Incorrect Index Format")

        start_col, start_row, end_col, end_row = self._extract_index(idx)

        if start_col is not None and start_row is not None:
            # e.g. A1:B5
            if end_col is not None and end_row is not None:
                return self.df.iloc[(start_row - 1) : end_row, start_col : end_col + 1]
            # e.g. A1
            else:
                return self.df.iloc[(start_row - 1), start_col]
        elif start_col is not None:
            # A:B
            if end_col is not None:
                return self.df.iloc[:, start_col : end_col + 1]
            # A
            else:
                return self.df.iloc[:, start_col]

    def __repr__(self):
        return self.df.__repr__()


class ExcelDF(pd.DataFrame):
    @property
    def excel(self):
        return _ExcelIndexer(self)


if __name__ == "__main__":
    df = pd.read_excel("excel-file.xlsx")
    edf = ExcelDF(df)
    print('Full DF')
    print(edf)
    print("A1")
    print(edf.excel["A1"])
    print("A1:B3")
    print(edf.excel["A1:B3"])
    print("A")
    print(edf.excel["A"])
    print("A:B")
    print(edf.excel["A:B"])

