"""SQLite store for one ingested workbook. In-memory by default; save() persists it."""
import hashlib, sqlite3
import pandas as pd
from .ingest import read_workbook

class Dataset:
    def __init__(self, wb, name="upload", digest=""):
        self.wb, self.name, self.digest = wb, name, digest
        self.con = sqlite3.connect(":memory:", check_same_thread=False)
        for t in ("students", "terms", "subjects", "marks", "cumulative"): getattr(wb, t).to_sql(t, self.con, index=False)

    @classmethod
    def from_workbook(cls, src, name="upload", include_mse_rows=False):
        data = src if isinstance(src, (bytes, bytearray)) else (open(src, "rb").read() if isinstance(src, str) else src.read())
        return cls(read_workbook(bytes(data), include_mse_rows=include_mse_rows), name, hashlib.sha1(data).hexdigest()[:12])

    def q(self, sql, params=()): return pd.read_sql_query(sql, self.con, params=params)
    def save(self, path): 
        out = sqlite3.connect(path); self.con.backup(out); out.close()
