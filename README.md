# 在 Vercel 上部署（Flask + 生成 Word/Zip）

## 1. 推荐的项目结构

```
./app.py                 # Flask 入口（导出 app 变量）
./requirements.txt
./templates/...          # 你的 docx 模板（只读即可）
./public/index.html      # 前端页面（Vercel 推荐 public/）
./public/static/...      # 你需要的静态资源（可选）
```

> Vercel 官方建议静态资源放在 `public/**`，不要依赖 Flask 的 `app.static_folder`。

## 2. 关键点（为什么需要改动）

- Vercel Functions 是无状态的，且代码目录通常只读；只有临时目录（如 `/tmp`）可写。
- Vercel Functions 的请求/响应体大小有限制，下载大文件不适合直接通过函数返回。
- 推荐用 Vercel Blob 存储生成的 zip，并把 Blob URL 返回给前端下载。

## 3. Blob 配置

1) 在 Vercel 控制台进入项目 -> Storage -> 创建 Blob store
2) 选择把读写 Token 注入到环境变量（默认是 `BLOB_READ_WRITE_TOKEN`）
3) 重新部署

## 4. 本地开发

```
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

## 5. 部署

把代码推到 GitHub，然后在 Vercel 导入仓库即可。
