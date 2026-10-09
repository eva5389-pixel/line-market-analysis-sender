"""Render the report with the dog's transparent cutout."""
import base64
from html import escape
from pathlib import Path


def report_preview_html(message: str) -> str:
    image = Path(__file__).parent / 'assets' / 'report-dog.png'
    encoded = base64.b64encode(image.read_bytes()).decode('ascii')
    return '''<style>
.dog-report {display:flow-root;background:#000000;color:#ffffff;color-scheme:dark;border-radius:16px;padding:20px;line-height:1.65;}
.dog-report .mascot {float:right;width:130px;max-width:25%;height:auto;max-height:180px;object-fit:contain;margin:0 0 16px 24px;}
.dog-report .copy {color:#ffffff;white-space:pre-wrap;overflow-wrap:anywhere;}
@media(max-width:600px){.dog-report{padding:14px}.dog-report .mascot{width:86px;margin-left:12px;}}
</style><section class="dog-report"><img class="mascot" alt="碗糕狗狗" src="data:image/png;base64,'''+encoded+'"><div class="copy">'+escape(message)+'</div></section>'
