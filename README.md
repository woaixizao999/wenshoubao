# 稳收宝

稳收宝是一个本地 Windows 工具，用天天基金/东方财富公开数据筛选低风险债券基金，并按 100 分模型输出前十名、完整评分表和剔除清单。

## 使用方式

双击 `win.exe`，点击“抓取并分析”；也可以输入基金名称或代码，点击“搜索基金并评分”。

输出文件会生成在：

- `E:\codex\investment\稳收宝输出\前十名.csv`
- `E:\codex\investment\稳收宝输出\完整评分表.xlsx`
- `E:\codex\investment\稳收宝输出\剔除清单.csv`
- `E:\codex\investment\稳收宝输出\稳收宝报告.html`
- `E:\codex\investment\稳收宝输出\运行日志.txt`

缓存文件会生成在：

- `E:\codex\investment\稳收宝缓存`

## 命令行调试

```powershell
& "C:\Users\bangl\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" .\wenshoubao.py --cli --deep-limit 120
```

离线回归测试：

```powershell
& "C:\Users\bangl\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" .\wenshoubao.py --cli --offline --deep-limit 80
```

## 打包

先确保当前 Python 环境安装了 PyInstaller，然后运行：

```powershell
.\build_exe.ps1
```

生成文件：

```text
wenshoubao\dist\win.exe
```

## 重要提示

本工具只做基金筛选研究，不构成投资建议或收益承诺。购买前请到天天基金、基金公司官网和实际销售平台再次核对风险等级、费率、经理、持仓、限购和申赎规则。
