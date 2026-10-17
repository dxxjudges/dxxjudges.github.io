# -*- coding: utf-8 -*-
"""dxxjudges 组织介绍 PPT 的内容源。

每张幻灯片的可见文案与解说词都放在这里，
make_ppt.py（画面）与 make_audio.py（配音）共用同一份数据，避免两边不一致。
"""

# 站点
SITE_ORG = "https://github.com/dxxjudges"
SITE_NOLD = "https://nold.dxxj.de5.net"
SITE_CLI = "https://jl.dxxj.de5.net"
REPO_NOLD = "https://github.com/dxxjudges/NoldOJ"
REPO_CLI = "https://github.com/dxxjudges/clijudge"

DECK_TITLE = "dxxjudges"
DECK_SUB = "开源评测系统（OJ）组织 · 组织介绍"
DECK_FOOT = "nold.dxxj.de5.net · jl.dxxj.de5.net · github.com/dxxjudges"


# kind: cover | cards | bullets | spotlight | table | compare | code | links | cta
SLIDES = [
    dict(
        kind="cover",
        title=DECK_TITLE,
        subtitle=DECK_SUB,
        note=DECK_FOOT,
        say="欢迎了解 dxxjudges。我们是一个专注评测系统的开源组织，"
            "旗下有 NoldOJ 和 CliJudge 两款项目。",
    ),
    dict(
        kind="bullets",
        tag="Who we are",
        title="我们是谁",
        lead="一个把「安全评测」当作基础设施来做的开源组织。",
        items=[
            ("专注单一赛道", "持续打磨评测系统（OJ），不做无关的方向发散"),
            ("两种形态", "NoldOJ 面向学校与团队的完整 Web 平台；CliJudge 面向个人的命令行评测器"),
            ("同一套内核", "沙箱隔离、评测语义、语言体系、题库格式在两个项目间保持一致"),
            ("本地优先", "下载后本地运行，完全离线可用；站点只是可选的托管实例"),
            ("开源优先", "MIT License，源码、Wiki、发布包全部公开，欢迎 Issue 与 PR"),
        ],
        say="dxxjudges 专注评测系统这一个赛道。我们维护两款产品："
            "NoldOJ 是面向学校和团队的完整网页平台，CliJudge 是面向个人的命令行评测器。"
            "两者共享同一套沙箱内核与评测语义，都可以下载后本地离线运行，"
            "并且都以 MIT 协议开源。",
    ),
    dict(
        kind="cards",
        tag="Projects",
        title="两款项目",
        lead="一个给你完整的平台，一个给你极致的效率。",
        cards=[
            dict(
                title="NoldOJ",
                sub="Windows 评测系统 · 本地部署的 Web 平台",
                site="nold.dxxj.de5.net",
                url=SITE_NOLD,
                repo=REPO_NOLD,
                points=[
                    "30+ 前端页面 · 100+ API 端点",
                    "比赛 / Rating / 题单 / 浏览器 IDE",
                    "AI 代码安全审查与代码查重",
                    "Node.js 24 + SQLite，一条命令启动",
                ],
            ),
            dict(
                title="CliJudge",
                sub="命令行评测系统 · Windows / Linux / Android",
                site="jl.dxxj.de5.net",
                url=SITE_CLI,
                repo=REPO_CLI,
                points=[
                    "单文件可执行 · 零外部依赖",
                    "JSON 存储，备份迁移方便",
                    "测试点生成器 · LemonLime CDF 互操作",
                    "npm 一行安装，自动选架构",
                ],
            ),
        ],
        say="NoldOJ 是完整的网页版评测系统，功能覆盖做题、比赛、排行榜和代码查重。"
            "CliJudge 则是一个单文件的命令行评测器，零依赖，数据以 JSON 保存，"
            "适合脚本化批量评测。",
    ),
    dict(
        kind="bullets",
        tag="NoldOJ",
        title="NoldOJ · Windows Web OJ",
        lead="一条命令启动的完整评测平台，装好依赖约 1.83 GB，完全离线可用。",
        items=[
            ("评测核心", "C / C++ / Python 3 / Java / JavaScript，面板动态扩展语言；严格、宽松、浮点容差与 Special Judge 四种比较模式"),
            ("社区生态", "讨论区、题解区、文章与公告、题目标签与难度、模糊搜索、图床"),
            ("激励机制", "Rating 体系与等级称号、成就解锁、提交日历、金银铜榜与虚拟比赛"),
            ("AI 能力", "本地 Ollama 审查提交代码，拦截挖矿、提权与信息窃取；失败两次后给出算法思路引导"),
        ],
        say="NoldOJ 提供完整的评测体验：多语言判题、比赛与 Rating、题单训练计划，"
            "还接入了本地模型做代码安全审查和算法提示。下载一键安装包后，一条命令即可启动。",
    ),
    dict(
        kind="bullets",
        tag="CliJudge",
        title="CliJudge · 跨平台命令行 OJ",
        lead="一个可执行文件就是全部，Windows、Linux 与 Android Termux 通用。",
        items=[
            ("六大模块", "article、contest、displaylang、ide、problem、submit，clijudge help 即上手"),
            ("题目与测试点", "子任务计分、浮点容差、Special Judge，以及 interactor / grader 交互题"),
            ("测试点生成器", "题目可附带生成器，每次评测前按测试点编号自动产出数据与标程答案"),
            ("互操作", "导入导出 LemonLime CDF，导出排行榜与 HTML 比赛报告，显示语言包可联网切换"),
        ],
        say="CliJudge 把评测能力装进一个可执行文件。它覆盖题目、测试点、比赛和文章管理，"
            "还支持测试点生成器和交互题，并且能和 LemonLime 的题库格式互通。",
    ),
    dict(
        kind="spotlight",
        tag="Shared core",
        title="共享的评测内核",
        lead="同一份题面与测试数据，在两套系统上得到一致的评测结果。",
        items=[
            ("统一评测语义", "相同的比较模式、子任务计分规则与交互式判定"),
            ("可扩展语言", "内置五种语言，并支持注册任意扩展名的自定义语言"),
            ("题库互通", "LemonLime CDF 导入导出，题目打包迁移"),
            ("轻量存储", "NoldOJ 用 SQLite，CliJudge 用 JSON，都不依赖外部数据库"),
            ("一致的沙箱", "Windows Job Object 隔离在两个项目中同样生效"),
            ("相同取舍", "不可信代码始终在沙箱内执行，判题数据落在独立临时目录"),
        ],
        say="两个项目共享同一套评测内核：比较模式、计分规则、交互式判定和题库格式都保持一致，"
            "所以同一份数据在两边跑出来的结果相同，迁移题库也不需要转换。",
    ),
    dict(
        kind="spotlight",
        tag="Security",
        title="纵深防御的安全沙箱",
        lead="基于操作系统原生的进程隔离能力，把不可信代码关进受控环境。",
        items=[
            ("进程树隔离", "CREATE_SUSPENDED 配合 Job Object 绑定，杜绝竞态逃逸"),
            ("强制回收", "主进程结束，整棵进程树由系统强制回收"),
            ("禁止脱离", "禁用 BREAKAWAY，子进程无法脱离沙箱"),
            ("权限降级", "受限令牌剥离高危特权，并以低完整性级别运行"),
            ("跨平台隔离", "Linux 侧使用 namespaces、seccomp 与 RLIMIT，能力与 Windows 侧对齐"),
            ("失败即拒绝", "沙箱运行器缺失时 NoldOJ 默认拒绝判题，不静默裸跑"),
        ],
        say="安全是评测系统的底线。两个项目都用操作系统原生的能力做多层隔离："
            "进程树绑定、强制回收、权限降级、资源限制，Linux 侧还有 namespaces 和 seccomp。"
            "沙箱不可用时，系统宁可拒绝判题，也不会让代码裸跑。",
    ),
    dict(
        kind="table",
        tag="Comparison",
        title="怎么选",
        lead="看你要的是一整套平台，还是一条命令。",
        headers=["维度", "NoldOJ", "CliJudge"],
        rows=[
            ["定位", "面向学校 / 团队的完整评测平台", "面向个人训练与批量脚本的评测器"],
            ["运行平台", "Windows 10 及以上（x64）", "Windows / Linux / Android（x64 与 ARM64）"],
            ["界面", "浏览器，30+ 页面，4 级角色", "纯命令行，6 大模块"],
            ["实现", "Node.js 24 + Express + SQLite", "C++17 单文件，静态链接"],
            ["存储", "SQLite（WAL 模式）", "JSON 文本"],
            ["安装", "一键安装包或 start.bat", "npm install -g clijudge"],
            ["特色", "比赛、Rating、AI 审查、代码查重", "生成器、CDF 互操作、HTML 报告"],
        ],
        say="这张表帮你快速选择：需要给学校或团队提供做题和比赛体验，用 NoldOJ；"
            "只需要在本地或服务器上批量评测题库，用 CliJudge。",
    ),
    dict(
        kind="code",
        tag="Quick start",
        title="快速开始",
        lead="两条命令，各拉起一套评测环境。",
        blocks=[
            ("NoldOJ", [
                "git clone https://github.com/dxxjudges/NoldOJ.git",
                "cd NoldOJ/backend && npm install && cd ..",
                "start.bat          # 或下载 411MB 一键安装包",
                "# 浏览器打开 http://localhost:3000",
            ]),
            ("CliJudge", [
                "npm install -g clijudge",
                "",
                "clijudge problem create \"A + B Problem\"",
                "clijudge problem testdata 1 create in.txt out.txt",
                "# create [in] [out] [时限] [内存] [分值]",
                "clijudge problem submit 1 solution.cpp",
            ]),
        ],
        say="上手都很简单。NoldOJ 克隆之后安装依赖并运行 start 脚本；"
            "CliJudge 一行 npm 安装就能直接创建题目、加测试点并提交解答。",
    ),
    dict(
        kind="links",
        tag="Ecosystem",
        title="资源与文档",
        lead="两套系统下载后都能完全本地离线运行；站点只是可选的托管实例。",
        groups=[
            ("NoldOJ", [
                ("站点", SITE_NOLD),
                ("GitHub 仓库", REPO_NOLD),
                ("Wiki 文档", REPO_NOLD + "/wiki"),
                ("Releases 下载", REPO_NOLD + "/releases"),
                ("支持中心", SITE_NOLD + "/support.html"),
            ]),
            ("CliJudge", [
                ("站点", SITE_CLI),
                ("GitHub 仓库", REPO_CLI),
                ("Wiki 文档", REPO_CLI + "/wiki"),
                ("Releases 下载", REPO_CLI + "/releases"),
                ("npm 安装", "https://www.npmjs.com/package/clijudge"),
            ]),
        ],
        say="NoldOJ 和 CliJudge 都可以下载后本地离线运行，站点只是可选的托管实例。"
            "源码仓库、Wiki 文档和发布包全部公开。"
            "CliJudge 还提供 npm 一行安装。",
    ),
    dict(
        kind="bullets",
        tag="Contribute",
        title="开源与参与",
        lead="两个项目都在持续维护，欢迎一起把它做得更好。",
        items=[
            ("许可", "MIT License，允许商业使用、修改与再分发"),
            ("反馈", "在对应仓库提交 Issue，或邮件联系维护者"),
            ("贡献", "NoldOJ 提供贡献指南与安全策略文档，欢迎提交 PR"),
            ("共建", "题目数据、题解与踩坑经验都欢迎分享进 Wiki"),
        ],
        say="两个项目都以 MIT 协议开源，欢迎提交 Issue 和 PR，"
            "也欢迎把题目数据和题解分享到 Wiki 里。",
    ),
    dict(
        kind="cta",
        title="一起把评测做得更好",
        subtitle="nold.dxxj.de5.net · jl.dxxj.de5.net",
        note="github.com/dxxjudges",
        say="以上就是 dxxjudges 的全部内容。感谢观看，欢迎访问我们的站点和仓库。",
    ),
]


def say_text(index):
    """第 index 页（从0 开始）的解说词。"""
    return SLIDES[index]["say"]