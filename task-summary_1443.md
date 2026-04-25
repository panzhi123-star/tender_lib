# tender_lib EXE 修复 — 2026-04-25 14:43

## 问题
EXE 启动报错：`ModuleNotFoundError: No module named 'gui.insert_dlg'`

## 根因
`gui/insert_dlg.py` 写入了不完整的 CompareDlg `_do_compare` 方法，
语法错误导致整个模块加载失败 → PyInstaller 打包时 exclude 该模块 → EXE 运行时 import 失败。

具体错误位置（insert_dlg.py L174）：
```python
vl = self.versions[lidx[0]   # 换行符导致 [ 未关闭
vr = self.versions[ridx[0]]
```

## 修复
补全 `vl = self.versions[lidx[0]]` 行的闭合括号，验证 import 成功。

## 结果
- EXE 已重新构建（7.6 MB，14:43:48）
- 文件：`dist\标书资料库管理工具\标书资料库管理工具.exe`
- 临时文件 `_rb2.py` 已清理