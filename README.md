# PaySys Bank Matcher

用于处理支付系统行号与收款开户行名称的两个桌面工具：

- `bank-name-from-code`：根据支付系统行号匹配收款开户行名称。
- `bank-code-from-name`：根据收款开户行名称匹配支付系统行号。

## 功能

- 支持 Excel 工作簿中的银行名称与支付系统行号匹配。
- 提供 PyQt5 图形界面和可直接调用的 Python 模块。
- 收款开户行匹配采用银行主体定位、完整名称优先和网点名称正序状态机匹配。
- 对不唯一或置信度不足的结果列出候选项，不自动写入不确定的行号。

## 环境

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## 运行

```powershell
.\.venv\Scripts\python.exe .\bank-name-from-code\bank_match_gui.py
.\.venv\Scripts\python.exe .\bank-code-from-name\bank_code_match_gui.py
```

图形界面要求输入和输出工作簿使用 `.xlsx` 格式。请不要将生产数据、供应商信息、银行主数据或匹配结果提交到仓库。

## 测试

```powershell
cd .\bank-code-from-name
..\.venv\Scripts\python.exe -m unittest -v test_match_bank_codes.py
```

## 许可证

本项目采用 [GNU General Public License v3.0](LICENSE) 开源。分发修改版本时，请遵守 GPLv3 的源代码提供和许可证保留要求。
