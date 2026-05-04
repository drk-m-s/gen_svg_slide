"""
Minimal config module providing CANVAS_FORMATS for svg_to_pptx/pptx_dimensions.py.

This replaces the dependency on skills/ppt-master/scripts/config.py.
"""

CANVAS_FORMATS = {
    "ppt169": {
        "name": "PPT 16:9",
        "dimensions": "1280×720",
        "viewbox": "0 0 1280 720",
    },
    "ppt43": {
        "name": "PPT 4:3",
        "dimensions": "1024×768",
        "viewbox": "0 0 1024 768",
    },
    "xhs": {
        "name": "Xiaohongshu (RED)",
        "dimensions": "1242×1660",
        "viewbox": "0 0 1242 1660",
    },
    "square": {
        "name": "Square (Instagram)",
        "dimensions": "1080×1080",
        "viewbox": "0 0 1080 1080",
    },
    "story": {
        "name": "Story / TikTok",
        "dimensions": "1080×1920",
        "viewbox": "0 0 1080 1920",
    },
    "wechat_header": {
        "name": "WeChat Article Header",
        "dimensions": "900×383",
        "viewbox": "0 0 900 383",
    },
    "banner": {
        "name": "Landscape Banner",
        "dimensions": "1920×1080",
        "viewbox": "0 0 1920 1080",
    },
    "a4": {
        "name": "A4 Print",
        "dimensions": "1240×1754",
        "viewbox": "0 0 1240 1754",
    },
}
