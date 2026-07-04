# FDA 指导原则自动化下载与整理工具

这是一个给RA同仁们使用的小工具，主要帮你做两件事：

1. 从 FDA Guidance Documents 页面抓取列表并下载原始文件
2. 按规则把下载下来的文件整理成更容易查找的资料库

如果你不会写代码，也没关系。你只需要按照 README 里的步骤，依次运行 2 个命令就可以完成下载和整理。

## 这个工具适合谁

适合这些场景：

- 需要批量保存 FDA guidance 文件到本地
- 想把零散下载的 PDF 自动整理到分类文件夹里
- 希望生成一个能直接点击打开文件的 Excel 索引

如果你只是普通办公用户，不打算改代码，也完全可以直接使用。

## 你最终会得到什么

跑完之后，通常会得到这几类结果：

- `FDA_Guidance_Data_YYYYMMDD.xlsx`
  - 这是从 FDA 页面抓下来的清单
- `FDA_Downloads/`
  - 这是下载下来的原始文件
- `FDA_Guidance_Library/`
  - 这是整理好的分类资料库
- `00_FDA_Guidance_Index.xlsx`
  - 这是整理后生成的索引表，能直接点开本地文件

## 第一次使用，建议只看这一节

下面这套流程最适合第一次使用的人。

### 第 1 步：准备环境

你的电脑需要有：

- Windows 10 或 Windows 11
- Python 3.8 或更高版本
- Google Chrome 浏览器

如果你已经装过 Python 和 Chrome，就可以直接继续下一步。

### 第 2 步：把项目下载到本地

如果你是从 GitHub 打开这个项目：

1. 点击页面上的 `Code`
2. 选择 `Download ZIP`
3. 解压到你电脑里一个方便找到的位置，比如桌面或工作文件夹

解压后，打开这个项目文件夹。

### 第 3 步：打开命令行窗口

在项目文件夹中：

1. 点击文件夹地址栏
2. 输入 `powershell`
3. 按回车

这样会直接在当前项目目录打开一个 PowerShell 窗口。

### 第 4 步：安装依赖

在打开的窗口里运行：

```bash
pip install -r requirements.txt
```

这一步只需要首次使用时执行一次。

### 第 5 步：先下载文件

运行：

```bash
python FDADownloader.py
```

运行后会发生这些事情：

1. 程序会自动打开 Chrome
2. 进入 FDA guidance 页面
3. 你需要在页面里手动选择筛选条件
4. 很重要：把显示数量改成 `Show All`
5. 等页面完全加载后，回到 PowerShell 窗口按回车
6. 程序会导出清单，并开始下载文件

![Show All screenshot](_image/1.png)

下载完成后，你会在项目目录里看到：

- 一个 Excel 清单文件
- 一个 `FDA_Downloads` 文件夹

### 第 6 步：再整理文件

下载完成后，继续运行：

```bash
python FDAOrganizer.py
```

它会自动：

- 找到最新的 Excel 清单
- 读取 `FDA_Downloads` 里的文件
- 按分类规则复制到新的资料库目录
- 生成一个索引 Excel

整理完成后，你主要去看 `FDA_Guidance_Library` 文件夹就可以了。

![Library tree example](_image/2.png)
![Library sheets example](_image/3.png)

## 最简单的使用顺序

如果你不想研究参数，只记住下面 3 条就够了：

```bash
pip install -r requirements.txt
python FDADownloader.py
python FDAOrganizer.py
```

## 常见使用场景

### 场景 1：我只想按默认方式跑一遍

直接执行：

```bash
python FDADownloader.py
python FDAOrganizer.py
```

这是最推荐的用法。

### 场景 2：我已经有下载清单了，只想继续下载

如果项目目录下已经有当天生成的 `FDA_Guidance_Data_YYYYMMDD.xlsx`，程序会问你：

- 直接读取已有清单继续下载
- 还是重新抓取

如果你只是上次没下完，通常选继续即可。

### 场景 3：我只想预览整理结果，不想真的复制文件

运行：

```bash
python FDAOrganizer.py --dry-run
```

这个模式只看结果，不真正复制文件，也不会生成最终索引表。

## 常用命令

### 下载 guidance 和原始文件

```bash
python FDADownloader.py
```

### 整理下载结果

```bash
python FDAOrganizer.py
```

### 指定下载目录

```bash
python FDADownloader.py --download-dir FDA_Downloads
```

### 指定某个 Excel 清单来整理

```bash
python FDAOrganizer.py --excel FDA_Guidance_Data_20260529.xlsx --source FDA_Downloads --target FDA_Guidance_Library --rules classification_rules.csv
```

### 直接下载模式

```bash
python FDADownloader.py --download-mode direct
```

说明：这个模式不一定适合所有 guidance 页面，只适合下载链接本身可以直接访问的情况。

## 每一步大概在做什么

### `FDADownloader.py`

它负责：

- 打开 FDA guidance 页面
- 读取页面表格
- 导出 Excel 清单
- 下载原始文件
- 避免重复下载

### `FDAOrganizer.py`

它负责：

- 读取 Excel 清单
- 找到本地已下载文件
- 按规则分类整理
- 生成带超链接的索引 Excel

## 分类规则怎么改

默认规则保存在 `classification_rules.csv`。

这个文件有两列：

- `Keyword`
- `Folder`

意思是：

- 如果文件信息里匹配到某个关键词
- 就把它归入对应文件夹

而且是从上到下按顺序匹配，所以越靠前，优先级越高。

如果你的业务重点不是 Oncology，或者你想单独分出别的主题，可以直接编辑这个 CSV 文件。

## 常见问题

### 1. 浏览器没有自动打开

请先确认：

- 电脑上已经安装 Chrome
- Chrome 可以正常手动打开

如果 Chrome 不在默认路径，程序会提示你手动输入 `chrome.exe` 的完整路径。

### 2. 页面打开了，但程序不继续

这是正常的。

因为程序在等你手动完成这一步：

- 在 FDA 页面设置筛选条件
- 把显示数量调成 `Show All`
- 等页面加载完
- 然后回到命令行窗口按回车

### 3. 下载速度比较慢

这是正常设计，不是卡死。

原因是：

- 下载本身是串行进行的
- 程序故意加入了随机等待时间
- 这样更稳，也能减少对目标网站的冲击

### 4. Excel 写入失败

大多数情况是因为目标 Excel 正在被打开。

请先关闭：

- Excel
- WPS
- 其他正在占用该文件的软件

然后重新运行。

### 5. 整理后发现有些文件没进目标目录

通常可以先检查这几件事：

- 文件是否真的下载到了 `FDA_Downloads`
- Excel 清单里的日期、标题是否正常
- `classification_rules.csv` 是否把它分到了你没注意到的目录

## 适用环境

当前仓库主要按下面的环境设计：

- Python 3.8+
- Google Chrome 或兼容 Chromium 内核浏览器
- Windows 10/11

虽然理论上不一定完全限制在 Windows，但默认流程和路径检测都是围绕 Windows 本地桌面场景写的。

## 测试

如果你是普通使用者，这一节可以跳过。

开发或验证时可以运行：

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
