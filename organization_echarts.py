"""Mount the installed streamlit-echarts frontend using Streamlit's inline JS API.

The package's file-backed declaration fails manifest discovery in this environment.
Its shipped renderer is reused unchanged, without editing site-packages or fetching
browser scripts from a CDN. Upgrade this small adapter with the pinned package.
"""
from importlib.util import find_spec
from pathlib import Path
import streamlit as st


@st.cache_resource
def _component():
    package = Path(find_spec('streamlit_echarts').origin).parent
    bundles = list((package / 'frontend' / 'build').glob('index-*.js'))
    if len(bundles) != 1:
        raise RuntimeError('Expected one streamlit-echarts frontend bundle')
    return st.components.v2.component(
        'presence_organization_tree', js=bundles[0].read_text(encoding='utf-8'),
        html='<div class="echarts-container"></div>', isolate_styles=False,
    )


def render_tree(options, *, height, events, key):
    placeholder = '--x_x--0_0--'
    return _component()(
        data={'options': options, 'theme': '', 'height': height, 'width': '100%',
              'renderer': 'svg', 'replaceMerge': None, 'map': None,
              'selectionActive': False, 'selectionMode': [],
              'onEvents': {name: placeholder + body + placeholder for name, body in events.items()}},
        key=key, default={}, on_chart_event_change=lambda: None,
    )
