"""所有默认路径放在这一处。文件夹结构变了只改这里。"""
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "数据"
PSI_FILE = DATA / "Poland PSI DB+Channel.xlsx"  # 每周从飞书导出覆盖
TTL_DIR = DATA  # 放 TTL ID *.xlsx 的地方，自动用最新的一份
HISTORY = DATA / "历史周报"
OPERATORS_HISTORY = HISTORY / "Operators 运营商"
CONFIG = ROOT / "config" / "sources.yaml"
MAPPING = ROOT / "mapping" / "mapping.xlsx"
MAPPING_EXCEL = ROOT / "mapping" / "PSI码对照表.xlsx"
OUTPUT = ROOT / "output"
