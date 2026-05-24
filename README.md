# 稳收宝

稳收宝是一个本地 Windows 工具，用天天基金/东方财富公开数据筛选低风险债券基金，并按 100 分模型输出前十名、完整评分表和剔除清单。

评分只能屏蔽掉一批过去成绩考的差的学生，筛选出过去成绩考的好的一批学生，但是不能保证以后这些学生成绩依然如此。所以仅供作为选基的参考，不作为最终的投资建议。

## 使用方式

双击 `win.exe`，点击“抓取并分析”；也可以输入基金名称或代码，点击“搜索基金并评分”。

注意：

1、只能在 Windows 上直接运行。
2、第一次运行可能会被 Windows 安全提示拦一下，因为这是自己打包、未签名的软件，选择“仍要运行”即可。
3、程序抓取天天基金/东方财富公开数据时需要联网。

输出文件会生成在：

- 程序所在目录下的 `稳收宝输出` 文件夹
- 例如直接运行 `dist\win.exe` 时，输出会生成在 `dist\稳收宝输出`
- 主要文件包括：`前十名.csv`、`完整评分表.xlsx`、`剔除清单.csv`、`稳收宝报告.html`、`运行日志.txt`

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
