# SlideCraft · 文档生成可编辑 PPT

使用 **Streamlit + python-pptx**。支持 DOCX、带文字层的 PDF、Markdown、TXT、粘贴文字；不需要数据库或安全隧道。
默认规则模式无需 API。它按原文分段摘选，完整段落保留在讲稿备注；可选 AI 模式负责整理内容。

## 本地启动（Python 3.11 / 3.12）

在本项目目录打开终端：

~~~powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m streamlit run app.py
~~~

macOS / Linux 将 Python 路径替换为 .venv/bin/python。浏览器打开终端显示的 localhost 地址。
Windows 可双击 start.bat（首次启动需要下载依赖）。
安装完成后，运行 python verify.py 可检查读取、排版、导出与 Streamlit 界面流程。

## 部署到 Streamlit Community Cloud

1. 创建 GitHub 仓库，将本目录的源码上传。不要上传 .venv、用户文档或 .streamlit/secrets.toml。
2. 登录 https://share.streamlit.io/，选择 Create app，填写仓库、分支和入口 app.py。
3. Advanced settings 选 Python 3.12；需要 AI / 访问码时设置 Secrets，配置格式见 .streamlit/secrets.toml.example。
4. 点击 Deploy。部署成功后，Streamlit 会提供实际的 .streamlit.app 网址，分享该网址即可。

如果把项目放在仓库子目录，入口填写 子目录/app.py，requirements.txt 与 app.py 放在一起。
官方部署说明：https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy

## 使用

1. 上传或粘贴正文，也可载入示例。
2. 可选：上传参考 PPT、导入 style.json、上传 PNG/JPG。参考 PPT 只提取配色、字体和比例，不复刻复杂母版。
3. 选择规则分段或 AI，设置目标页数，生成大纲。
4. 逐页修改标题、每行一条的要点、讲稿、版式与配图。原文标签 S1、S2 等可在“原文与来源”查询。
5. 在预览中检查页面，点击“生成 PPTX 下载文件”，再下载 PowerPoint。修改内容后需重新生成下载文件。
6. 可下载大纲 JSON 与样式 JSON 另行保存。页面刷新、重启或云端休眠后，会话可能丢失。

文字、色块为可编辑原生对象，图片保留为图片。长要点自动流到续页，所以实际页数可能多于目标。
预览与导出使用同一套坐标和显式换行。预览并非 PowerPoint 渲染；设备字体替换可能影响实际观感。
建议中文字体使用 Microsoft YaHei、等线或 Noto Sans CJK SC，并在打开 PPT 的电脑上安装对应字体。

## 可选 AI 配置

默认不调用模型。AI 模式使用兼容 Chat Completions 的 HTTPS 服务（/chat/completions），
支持直接填写个人 Key；服务地址只由部署者配置，网页访客不能修改。
需要模型支持 max_tokens 参数和 JSON 文本输出。不同服务与模型的兼容性须实际验证。

在 Streamlit Cloud 的 Secrets 中按需设置：

~~~toml
OPENAI_API_KEY = "你的服务密钥"
OPENAI_MODEL = "你的服务支持的模型名称"
OPENAI_BASE_URL = "https://api.openai.com/v1"
APP_ACCESS_CODE = "可选的班级访问码"
~~~

没有配置 Key 时，访客可在页面填写自己的 Key。Key 不写入项目文件。
使用 AI 前需要勾选同意发送文档文字；只发送文字，不发送图片。AI 可能遗漏或误读，务必核对。
如果公开分享部署者的付费 Key，建议设置访问码；访问码仅用于小范围共享，不代替完整的用户与额度系统。
本项目没有账户、计费、数据库或跨实例限流。

## 资源与数据边界

- 单个文件最多 20 MB，正文最多 10 万字符，PDF 最多 120 页。
- Office ZIP 包解压总量最多 100 MB；图片最多 24 张，处理后图片合计最多 30 MB。
- AI 单次最多约 4.5 万字符；超过限制会拒绝，不会静默截断。每次点击只发起一次请求，无自动付费重试。
- 扫描 PDF 不做 OCR；混合扫描 PDF 会提示可疑页。PDF 图表需作为补充图片上传。
- DOCX 读取正文与表格、载入内嵌图片。图片与要点的对应关系由用户手动选择。
- 不支持旧版 .doc、公式转换、复杂表格原样复刻、动画或自动科学绘图。
- 文件只存当前 Streamlit 会话内存；不使用跨用户文档缓存，不写文档到磁盘。AI 服务的数据保留政策由服务商决定。
- 免费云端存在资源、重启与休眠限制；本项目未实现后台任务与持久保存。
- 科研图表、数字、文献与结论均需用户审核；规则模式只摘选原文，AI 模式不能保证内容事实正确。

## 代码结构与技能关系

- app.py：上传、设置、大纲编辑、预览与下载。
- core.py：文件读取、来源标签、规则大纲、可选 AI、大纲与样式校验。
- renderer.py：共同排版场景、自动续页、SVG 版式预览、原生 PPTX 导出。
- style.json：默认配色、字体和比例，可在网页中导入/导出。
- examples/：不含实测数据的演示材料。
- verify.py：离线功能与界面测试。

设计参考了本机 ppt-master 技能的“大纲先行、内容与排版分离、原生文字可编辑、检查后导出”原则。
本项目为独立 Python 实现，不包含该技能源码，也不在云端执行本机 Codex 技能。
技能本身不是网站运行库；不需要把 .codex/skills 上传 GitHub。
如需固定背景、Logo、复杂母版或新增特定技能的处理流程，可在此基础上扩展。

无需预先放置 template.pptx。未提供实际模板时使用内置布局；参考 PPT 和样式 JSON 均可在网页上传。
