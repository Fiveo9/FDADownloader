# FDA 指导原则自动化下载与整理工具

一个面向本地研究工作流的 Python 工具集，用于：

1. 从 FDA Guidance Documents 页面抓取元数据
2. 下载原始文件到本地
3. 按可配置规则整理到结构化资料库
4. 生成可点击的 Excel 索引

仓库目前以 Windows 本地桌面使用场景为主，尤其是浏览器自动下载和路径检测逻辑。其他平台并非完全不能运行，但没有作为默认支持目标来设计。

## 项目组成

- `FDADownloader.py`: 抓取 FDA guidance 列表、导出 Excel、下载原始文件
- `FDAOrganizer.py`: 根据分类规则整理文件并生成索引表
- `classification_rules.csv`: 默认分类规则，支持自行调整优先级和目标文件夹
- `tests/`: 面向核心逻辑的单元测试

## 功能特性

- 浏览器驱动下载，适合需要真实页面会话的场景
- 已下载文件自动跳过，支持增量运行
- 将文件规范命名为 `YYYYMMDD_Summary`
- 根据 `FDA Organization`、`Topic`、`Summary` 等字段进行优先级分类
- 生成本地 Excel 索引，便于二次检索和归档
- 支持 `--dry-run` 预览整理结果

## 适用环境

- Python 3.8+
- Google Chrome 或兼容 Chromium 内核浏览器
- Windows 10/11

依赖安装：

```bash
pip install -r requirements.txt
```

## 快速开始

### 1. 下载 guidance 元数据与原始文件

```bash
python FDADownloader.py
```

常用参数：

```bash
python FDADownloader.py --url https://www.fda.gov/regulatory-information/search-fda-guidance-documents --download-dir FDA_Downloads
```

如果页面中的下载链接可以直接访问，也可以尝试实验性的直接下载模式：

```bash
python FDADownloader.py --download-mode direct
```

运行时需要注意：

1. 脚本会启动浏览器并打开 FDA guidance 页面
2. 请在页面里完成你需要的筛选
3. 关键步骤：将结果显示数量调整为 `Show All`
4. 等页面刷新完成后，回到终端按回车继续

![Show All screenshot](_image/1.png)

完成后会在项目根目录生成 `FDA_Guidance_Data_YYYYMMDD.xlsx`，并将文件下载到 `FDA_Downloads/`。

### 2. 按分类规则整理本地文件

```bash
python FDAOrganizer.py
```

常用参数：

```bash
python FDAOrganizer.py --excel FDA_Guidance_Data_20260529.xlsx --source FDA_Downloads --target FDA_Guidance_Library --rules classification_rules.csv
```

仅预览，不复制文件：

```bash
python FDAOrganizer.py --dry-run
```

整理完成后，会在 `FDA_Guidance_Library/` 下创建分类目录，并生成 `00_FDA_Guidance_Index.xlsx`。

![Library tree example](_image/2.png)
![Library sheets example](_image/3.png)

## 分类规则

默认规则定义在 `classification_rules.csv` 中，按从上到下的顺序匹配：

- `Keyword`: 关键词
- `Folder`: 命中后归入的文件夹

只要首次命中就停止继续匹配，因此文件顺序就是优先级顺序。

默认规则的大致思路是：

1. 优先提取 Oncology 相关 guidance
2. 将 Combination Products 单独抬高优先级
3. 将 GCP、核查、政策类跨中心主题集中归档
4. 再按 CDER / CBER / CDRH 等主线归类

如果你的使用场景不同，直接编辑 `classification_rules.csv` 即可。

## 测试

运行单元测试：

```bash
python -m unittest discover -s tests -q
```

GitHub Actions 也会在提交和 PR 上自动运行基础校验。

## 生成文件与仓库策略

以下内容属于本地运行产物，默认不应提交到仓库：

- `FDA_Downloads/`
- `FDA_Guidance_Library/`
- `FDA_Guidance_Data_*.xlsx`

仓库应主要保存：

- 源代码
- 测试
- 分类规则
- 文档和协作配置

## 已知限制

- 当前工作流依赖人工在网页中完成筛选和 `Show All`
- 浏览器自动下载逻辑以 Windows 路径和桌面环境为主
- `--download-mode direct` 依赖目标链接可直接访问，不保证对所有 guidance 都成功
- 大批量下载时建议控制频率，避免给目标站点造成不必要压力

## 常见问题

### 浏览器无法启动

- 确保本机已安装 Chrome
- 如果 Chrome 不在默认路径，脚本会提示你手动输入 `chrome.exe` 路径

### 下载很慢

- 浏览器下载模式是串行执行
- 脚本会插入随机等待时间，尽量降低对站点的冲击

### Excel 无法写入

- 请确认目标 Excel 文件没有被 WPS、Excel 或其他程序占用

## 合规与免责声明

本项目与 FDA 无官方关联，仅用于学习、研究和个人知识整理。

请在使用前自行确认并遵守：

- FDA 网站使用条款
- robots.txt 或其他访问限制
- 你所在组织的合规要求

## 协作与许可

- 许可证：MIT，见 `LICENSE`
- 贡献指南：见 `CONTRIBUTING.md`
- 安全披露：见 `SECURITY.md`
- 社区行为规范：见 `CODE_OF_CONDUCT.md`
