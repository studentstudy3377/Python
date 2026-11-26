from pathlib import Path
import streamlit as st
import streamlit.components.v1 as components

HTML_PATH = Path(__file__).parent / "Diamond-Sticker-Printer.html"


@st.cache_data
def _load_html(path: str, mtime: float) -> str:
    # mtime in the cache key means edits to the HTML file are picked up automatically
    return Path(path).read_text(encoding="utf-8")


def run_print_sticker():
    st.markdown("### 🧾 Print Sticker")

    if not HTML_PATH.exists():
        st.error(f"Sticker file not found: {HTML_PATH}")
        return

    html = _load_html(str(HTML_PATH), HTML_PATH.stat().st_mtime)

    # Renders the full sticker tool inside an iframe.
    # File upload, preview and the "Print stickers" button all work inside it.
    components.html(html, height=1100, scrolling=True)