# dabao — 通用 Python 打包 EXE

复用 GitHub Actions，一键将 Python 脚本编译为 Windows exe。

## 使用方法

在你的项目根目录创建 `.github/workflows/build.yml`：

```yaml
name: Build EXE

on:
  push:
    tags: [ 'v*' ]
  workflow_dispatch:

jobs:
  build:
    uses: fyjn4r5/dabao/.github/workflows/build-exe.yml@main
    with:
      script_name: 你的主脚本.py     # 必填
      exe_name: 输出文件名            # 选填，默认 output
```

### 参数说明

| 参数 | 必填 | 说明 |
|------|------|------|
| `script_name` | 是 | 入口 Python 文件名 |
| `exe_name` | 否 | 生成的 exe 名称（默认 output） |

## 依赖

如果有额外依赖（如 python-docx），在项目根目录放 `requirements.txt` 即可自动安装。

## 手动触发

GitHub 仓库 → **Actions** → **Build EXE** → **Run workflow**，无需打标签。
