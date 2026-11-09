"""Entry point: read_auto(raw) = saved layout for this sheet shape if one exists, else automatic detection. Offline, no API/LLM."""
from .u1_detect import profile
from .u3_build import read_generic
from .u4_layouts import load_layout, save_layout, delete_layout
from .ingest import UnsupportedLayout

def read_auto(raw, overrides=None):
    saved = load_layout(raw) if overrides is None else None
    wb = read_generic(raw, overrides if overrides is not None else saved)
    wb.info.update(generic=True, layout_saved=saved is not None)
    return wb
