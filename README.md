# PSI_agent

PSI_xiaomi：小米波兰 GTM 的 PSI 周更工具。它把经销商（AB / ALSO / IM）、公开市场 KA 和四大运营商的原始周报整理成可以直接粘贴进飞书「Poland PSI DB+Channel」的数据，并输出核对报告。

> ⚠️ 这个仓库**只放代码、配置和说明**。PSI 表、TTL ID 表、客户周报都是公司 L3 机密，已在 `.gitignore` 里排除，不要手动加进来。

## 第一次在一台电脑上使用

```bash
git clone https://github.com/Jacqueline772026/PSI_agent.git
cd PSI_agent
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

然后把数据放进 `数据/` 文件夹（这个文件夹不在仓库里，要自己建）：
- `数据/Poland PSI DB+Channel.xlsx`：从飞书导出的 PSI 表
- `数据/TTL ID *.xlsx`：EAN → 代码对照表
- `数据/历史周报/`：过往周报（可选；从历史周报重新学映射、做验证时才需要）

想用 Jupyter 调试本的话，再注册一个 kernel：

```bash
.venv/bin/python -m ipykernel install --user --name psi-agent --display-name "Python (PSI agent)"
```

## 每周怎么用

完整步骤见 [使用说明.md](使用说明.md)，W44 的试运行清单见 [W44测试/W44测试清单.md](W44测试/W44测试清单.md)。主要四条命令：

```bash
# 1. 把下载文件夹里的周报按 来源 + 周 整理好
.venv/bin/python -m tools.sort_downloads ~/Downloads/PSI_W45 --week 45 --dest "W45/W45原始文件"

# 2. 生成待粘贴表和核对报告
.venv/bin/python -m psi_tool.run --week 45 --inbox "W45/W45原始文件" --out "W45/工具结果"

# 3.（并行手工期间）和自己的 Excel 结果对数
.venv/bin/python -m tools.compare_manual --week 45 --template
.venv/bin/python -m tools.compare_manual --week 45

# 4. 检查 PSI 表本身的公式结构
.venv/bin/python -m checks.structure_check
```

## 文件夹

```
PSI_agent/
├── README.md / 使用说明.md / requirements.txt
├── psi_tool/               主程序：读文件、映射、汇总、输出（paths.py 里集中放所有默认路径）
├── config/sources.yaml     每家周报怎么读：文件名、表头、用哪几列、怎么筛选（客户改格式时改这里）
├── mapping/mapping.xlsx    人工确认过的 识别码 → PSI 代码（新品在这里补）
├── checks/                 PSI 表结构校验（公式漏行、重复代码、小计区间）
├── tools/                  整理下载文件夹、和 Excel 对数、从历史周报学映射、导出和核对对照表
├── notebooks/              Jupyter 调试本：单独调试某一家
├── docs/各家周报列说明.md    每家每列是什么、PSI 用哪一列、历史上发现的错误
├── docs/易错清单.md          填 PSI 前过一遍；对照表和源表的错误
├── docs/对照表核对方法.md    怎么确认一个代码是对的；来了新货、新表怎么办
├── 前期资料/                需求文档、GTM 实习 SOP
├── W44测试/W44测试清单.md    第一次试运行的步骤
│
│   以下只在本地，不进仓库：
├── 数据/                   PSI 导出、TTL ID 表、历史周报
├── W44测试/ 里的数据、W45/ …  每周的原始文件和结果
├── mapping/PSI码对照表.xlsx  由 tools.export_mapping 生成
└── cache/ output/ .venv/
```
