import asyncio
import json
import os
import random
import sys
import time
from pathlib import Path
from urllib.parse import quote_plus

from playwright.async_api import async_playwright

# 常量：集中管理字符串，避免拼写检查对字符串字面量误报
MSEDGE_CHANNEL = "msedge"
DOMCONTENTLOADED = "domcontentloaded"

# 记录程序启动时间，用于展示运行时长
APP_START_TS = time.time()


async def _maybe_click(page, selectors: list[str]) -> bool:
    """尝试点击一组 selector 中第一个可点击的元素。"""
    for sel in selectors:
        try:
            el = await page.query_selector(sel)
            if not el:
                continue
            try:
                if not await el.is_visible():
                    continue
            except Exception:
                # 部分情况下 is_visible 可能抛异常（比如 detached）
                pass
            await el.click(timeout=1500)
            return True
        except Exception:
            continue
    return False


async def _maybe_accept_bing_dialogs(page) -> None:
    """尽量处理 Bing/Consent 弹窗（失败也不影响主流程）。"""
    # 常见的 consent/隐私提示按钮（不同地区/语言可能不同）
    selectors = [
        "#bnp_btn_accept",  # Bing Cookie banner
        "button#bnp_btn_accept",
        "button:has-text('Accept')",
        "button:has-text('I agree')",
        "button:has-text('同意')",
        "button:has-text('接受')",
        "button:has-text('我同意')",
    ]
    try:
        await _maybe_click(page, selectors)
    except Exception:
        pass


async def _human_type_into_search_box(page, query: str) -> None:
    """更像真人：点击搜索框 -> 全选清空 -> 逐字输入。"""
    box = "#sb_form_q"
    await page.wait_for_selector(box, state="visible", timeout=8000)
    await page.click(box)
    # 全选清空（在 Windows/Linux 下 Control+A 更通用）
    try:
        await page.keyboard.press("Control+A")
        await page.keyboard.press("Backspace")
    except Exception:
        # 兜底：直接填充
        pass

    # 逐字输入：Playwright 的 delay 是“每个字符”的延时（毫秒）
    per_char_delay = random.randint(55, 140)
    await page.keyboard.type(query, delay=per_char_delay)

    # 给联想/脚本一点反应时间
    await asyncio.sleep(random.uniform(0.2, 0.9))


async def _maybe_human_scroll(page) -> None:
    """轻量随机滚动，避免过于“机械”。"""
    # 保守：不是每次都滚
    if random.random() > 0.65:
        return
    try:
        dy = random.randint(250, 1200)
        await page.mouse.wheel(0, dy)
        await asyncio.sleep(random.uniform(0.4, 1.4))
        # 少量概率再滚回一点
        if random.random() < 0.25:
            await page.mouse.wheel(0, -random.randint(120, 500))
            await asyncio.sleep(random.uniform(0.2, 0.8))
    except Exception:
        return


async def _maybe_click_one_result(page) -> None:
    """偶尔点开一个自然结果再返回（默认很保守）。"""
    # 风险控制：点击结果更接近真人，但也更容易引入不可控跳转
    if random.random() > 0.15:
        return
    try:
        links = await page.query_selector_all("li.b_algo h2 a")
        if not links:
            return
        # 选前 1~3 条更像随手点
        candidates = links[: min(len(links), 3)]
        link = random.choice(candidates)
        try:
            await link.scroll_into_view_if_needed()
        except Exception:
            pass
        await asyncio.sleep(random.uniform(0.2, 0.8))

        await link.click(timeout=3000)
        await page.wait_for_load_state(DOMCONTENTLOADED, timeout=15000)
        await asyncio.sleep(random.uniform(2.0, 6.0))
        try:
            await page.go_back(wait_until=DOMCONTENTLOADED, timeout=15000)
        except Exception:
            # go_back 不行就不强求
            pass
    except Exception:
        return

async def _inject_mobile_spoofing(context) -> None:
    """在context级别注入 JavaScript，让站点以为是真正的移动设备。"""
    try:
        await context.add_init_script("""
(() => {
  // 强制覆盖 navigator.maxTouchPoints（手机有触点，桌面通常是 0）
  try {
    Object.defineProperty(navigator, 'maxTouchPoints', {
      get: () => 5,
      configurable: true
    });
  } catch(e) {}

  // 强制覆盖 navigator.webdriver（反爬虫检查，机器人通常为 true）
  try {
    delete Object.getPrototypeOf(navigator).webdriver;
    Object.defineProperty(navigator, 'webdriver', {
      get: () => false,
      configurable: true
    });
  } catch(e) {}

  // 覆盖 navigator.platform（移动设备标识）
  try {
    Object.defineProperty(navigator, 'platform', {
      get: () => 'iPhone',
      configurable: true
    });
  } catch(e) {}

  // 覆盖 navigator.vendor
  try {
    Object.defineProperty(navigator, 'vendor', {
      get: () => 'Apple Computer, Inc.',
      configurable: true
    });
  } catch(e) {}

  // 覆盖 navigator.plugins（移动端通常为空）
  try {
    Object.defineProperty(navigator, 'plugins', {
      get: () => [],
      configurable: true
    });
  } catch(e) {}

  // 覆盖 navigator.mimeTypes（移动端通常为空）
  try {
    Object.defineProperty(navigator, 'mimeTypes', {
      get: () => [],
      configurable: true
    });
  } catch(e) {}

  // 添加触摸事件支持
  if (!window.ontouchstart) {
    window.ontouchstart = null;
  }
  if (!window.ontouchend) {
    window.ontouchend = null;
  }
  if (!window.ontouchmove) {
    window.ontouchmove = null;
  }
  if (!window.ontouchcancel) {
    window.ontouchcancel = null;
  }

  // 移动端通常支持 orientation
  try {
    Object.defineProperty(window, 'orientation', {
      get: () => 0,
      configurable: true
    });
  } catch(e) {}

  // 添加 onorientationchange 事件
  if (!window.onorientationchange) {
    window.onorientationchange = null;
  }

  // 覆盖 screen.orientation
  try {
    if (window.screen && !window.screen.orientation) {
      Object.defineProperty(window.screen, 'orientation', {
        get: () => ({
          type: 'portrait-primary',
          angle: 0
        }),
        configurable: true
      });
    }
  } catch(e) {}

  // 隐藏自动化特征
  try {
    delete navigator.__proto__.webdriver;
  } catch(e) {}

  // 隐藏 Chrome 对象（移动Safari没有）
  try {
    delete window.chrome;
  } catch(e) {}

  // 覆盖 navigator.connection（移动设备网络信息）
  try {
    if (!navigator.connection) {
      Object.defineProperty(navigator, 'connection', {
        get: () => ({
          effectiveType: '4g',
          rtt: 100,
          downlink: 10,
          saveData: false
        }),
        configurable: true
      });
    }
  } catch(e) {}

  // 覆盖 navigator.deviceMemory（移动设备内存）
  try {
    Object.defineProperty(navigator, 'deviceMemory', {
      get: () => 4,
      configurable: true
    });
  } catch(e) {}

  // 覆盖 navigator.hardwareConcurrency（CPU核心数）
  try {
    Object.defineProperty(navigator, 'hardwareConcurrency', {
      get: () => 6,
      configurable: true
    });
  } catch(e) {}
})();
        """)
    except Exception:
        # 注入失败也不影响主流程
        return

async def _perform_bing_search_like_human(page, query: str) -> bool:
    """对齐 userscript：在搜索框输入并提交，而不是每次直接跳转 search?q=。"""
    try:
        # 确保在 Bing 上（有时搜索过程中会跳去别的域）
        if "bing.com" not in (page.url or ""):
            await page.goto("https://www.bing.com", wait_until=DOMCONTENTLOADED)

        await _maybe_accept_bing_dialogs(page)
        await _human_type_into_search_box(page, query)

        # userscript 本质上是 form submit / click go；这里用 Enter 提交
        await page.keyboard.press("Enter")
        await page.wait_for_load_state(DOMCONTENTLOADED, timeout=20000)
        await _maybe_accept_bing_dialogs(page)
        return True
    except Exception:
        return False


def _get_exe_dir() -> Path:

    try:
        return Path(sys.argv[0]).resolve().parent
    except Exception:
        return Path.cwd()


def _get_bundle_dir() -> Path:
    """返回运行时资源目录。

    - 源码运行：等同于 main.py 所在目录
    - Nuitka onefile：等同于临时解压目录（__file__ 所在）
    """
    try:
        return Path(__file__).resolve().parent
    except Exception:
        return _get_exe_dir()


def _sanitize_filename(name: str) -> str:
    """把账号名转为 Windows 可用的文件名。"""
    # Windows 禁止字符: < > : " / \ | ? *
    invalid = '<>:"/\\|?*'
    cleaned = "".join("_" if ch in invalid else ch for ch in (name or ""))
    cleaned = cleaned.strip().strip(".")
    return cleaned or "未知账户"


# 定义User-Agent池
MOBILE_USER_AGENTS = [
    # iPhone - 不同系统版本和Safari版本
    "Mozilla/5.0 (iPhone; CPU iPhone OS 17_6_1 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.6 Mobile/15E148 Safari/604.1",
    "Mozilla/5.0 (iPhone; CPU iPhone OS 17_5_1 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.5 Mobile/15E148 Safari/604.1",
    "Mozilla/5.0 (iPhone; CPU iPhone OS 17_4_1 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.4 Mobile/15E148 Safari/604.1",
    "Mozilla/5.0 (iPhone; CPU iPhone OS 16_7_2 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.6 Mobile/15E148 Safari/604.1",
    "Mozilla/5.0 (iPhone; CPU iPhone OS 16_6_1 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.6 Mobile/15E148 Safari/604.1",
    
    # iPhone - Edge浏览器
    "Mozilla/5.0 (iPhone; CPU iPhone OS 17_6_1 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 EdgiOS/120.0.2210.150 Mobile/15E148 Safari/604.1",
    "Mozilla/5.0 (iPhone; CPU iPhone OS 17_5_1 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 EdgiOS/119.0.2151.96 Mobile/15E148 Safari/604.1",
    
    # Android - Samsung设备
    "Mozilla/5.0 (Linux; Android 14; SM-S918B) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.6099.144 Mobile Safari/537.36",
    "Mozilla/5.0 (Linux; Android 14; SM-S911B) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.6099.210 Mobile Safari/537.36",
    "Mozilla/5.0 (Linux; Android 13; SM-A536B) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.6045.193 Mobile Safari/537.36",
    
    # Android - Pixel设备
    "Mozilla/5.0 (Linux; Android 14; Pixel 8 Pro) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.6099.230 Mobile Safari/537.36",
    "Mozilla/5.0 (Linux; Android 14; Pixel 7a) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.6099.144 Mobile Safari/537.36",
    "Mozilla/5.0 (Linux; Android 13; Pixel 6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.6045.193 Mobile Safari/537.36",
    
    # Android - Edge浏览器
    "Mozilla/5.0 (Linux; Android 14; SM-S918B) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Mobile Safari/537.36 EdgA/120.0.2210.150",
    "Mozilla/5.0 (Linux; Android 13; Pixel 7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Mobile Safari/537.36 EdgA/119.0.2151.78",
]

# 定义User-Agent
USER_AGENTS = {
    "windows": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36 Edg/120.0.0.0",
    "iphone": lambda: random.choice(MOBILE_USER_AGENTS)  # 每次随机选择一个移动UA
}

REWARDS_API_URL = "https://rewards.bing.com/api/getuserinfo?type=1&X-Requested-With=XMLHttpRequest"

DASHBOARD_REFRESH_SECONDS = 30

SEARCH_CONFIG = {
    "points_per_search": 3,
    # 不同设备类型的随机等待区间（毫秒）
    "pc": {"min_delay_ms": 50_000, "max_delay_ms": 100_000},
    "mobile": {"min_delay_ms": 25_000, "max_delay_ms": 50_000},
}

SEARCH_KEYWORDS = [
    # 通用/信息
    "weather", "local weather", "today forecast", "news", "breaking news", "world news", "business news",
    "sports", "score", "highlights", "finance", "stock", "market news", "exchange rate", "inflation",
    "movies", "movie reviews", "box office", "tv series", "streaming", "anime", "documentary",
    "tech", "technology trends", "gadgets", "smartphone", "laptop", "windows tips", "edge browser",
    "food", "recipe", "easy dinner", "coffee", "tea", "baking", "dessert", "travel", "hotel", "flight",
    "music", "playlist", "podcast", "art", "museum", "photography", "design", "architecture",

    # 学习/知识
    "history", "world history", "ancient civilizations", "science", "biology", "chemistry", "physics",
    "math", "algebra", "geometry", "calculus", "statistics", "psychology", "philosophy", "economics",
    "nature", "wildlife", "plants", "space", "astronomy", "planets", "nasa", "space telescope",
    "geography", "maps", "time zone", "languages", "english learning", "grammar", "vocabulary",

    # 生活方式/健康
    "health", "fitness", "workout", "running", "yoga", "sleep", "nutrition", "meditation",
    "mental health", "stress relief", "habits", "productivity", "time management",

    # 运动/赛事
    "football", "soccer", "basketball", "tennis", "baseball", "formula 1", "olympics", "world cup",
    "nba", "nfl", "mlb", "uefa", "premier league", "champions league",

    # 汽车/交通
    "cars", "car review", "electric car", "tesla", "ev charging", "hybrid car", "motorcycle",
    "public transport", "traffic", "driving tips",

    # 游戏/娱乐
    "games", "video games", "pc gaming", "console", "esports", "game guide", "walkthrough",
    "minecraft", "fortnite", "league of legends", "valorant",

    # 图书/写作
    "books", "book list", "best novels", "nonfiction", "audiobooks", "writing", "story ideas",
    "poetry", "quotes",

    # 时尚/家居
    "fashion", "outfit ideas", "street style", "skincare", "hair care", "home decor",
    "interior design", "minimalism", "cleaning tips",

    # 美食/饮食扩展
    "healthy recipe", "meal prep", "air fryer", "soup", "pasta", "pizza", "salad", "bbq",

    # 科技开发/编程（会更随机一些）
    "python", "python async", "javascript", "typescript", "react", "nodejs", "docker",
    "linux", "powershell", "git", "api", "json", "automation", "web scraping",

    # 旅行目的地/文化
    "japan travel", "korea travel", "italy travel", "paris", "london", "new york",
    "street food", "culture", "festival",

    # AI / 热点科技
    "chatgpt", "gpt-4", "gemini ai", "copilot", "midjourney", "stable diffusion", "ai news",
    "machine learning", "deep learning", "computer vision", "nlp", "data science", "kaggle",

    # 经济/理财
    "personal finance", "side hustle", "mortgage rates", "housing market", "bond yield",
    "crypto price", "bitcoin", "ethereum", "stablecoin", "defi", "blockchain use case",

    # 职场/效率
    "resume tips", "interview questions", "career change", "remote work", "project management",
    "kanban", "scrum", "product roadmap", "meeting notes template", "time blocking",

    # 学习中文/本地化
    "中文学习", "拼音练习", "汉字笔顺", "唐诗三百首", "宋词", "成语故事", "古诗词解释",
    "中国历史朝代", "三国演义", "水浒传", "红楼梦",

    # 生活服务/实用
    "汇率换算", "火车票查询", "飞机票打折", "天气预报 明天", "菜谱 家常", "网购优惠",
    "手机评测", "路由器设置", "宽带测速",

    # 兴趣爱好
    "摄影技巧", "相机参数", "吉他和弦", "钢琴入门", "绘画教程", "手帐", "模型涂装",
    "露营装备", "徒步路线", "骑行训练",

    # 健康扩展
    "低脂餐", "增肌训练", "心率区间", "跑步配速", "马拉松训练计划", "颈椎放松",
    "护眼 tips", "喝水提醒",

    # 影视/剧集
    "netflix ranking", "disney plus", "hbo series", "kdrama", "cdrama", "imdb top",
    "影评 解读", "预告片",

    # 季节/节日
    "chinese new year", "lantern festival", "mid autumn", "dragon boat festival", "valentines day ideas",
    "halloween costume", "christmas markets", "new year resolution",

    # 本地化生活
    "上海天气", "北京地铁", "广州美食", "深圳科技园", "杭州旅游", "成都火锅", "西安古迹", "武汉樱花",

    # 电商/比价
    "京东优惠", "淘宝双11", "拼多多砍价", "黑五折扣", "prime day deals", "price comparison", "coupon code",

    # 硬件/DIY
    "组装电脑", "显卡天梯", "cpu 性能排行", "机械键盘", "3d 打印 入门", "树莓派 项目", "arduino 教程",

    # 开发框架/工具
    "fastapi", "django", "flask", "langchain", "pytorch", "tensorflow", "pandas tutorial", "numpy cheat sheet",

    # 商业/管理
    "swot analysis", "business model canvas", "okr 示例", "stakeholder map", "marketing plan", "seo checklist",
    "a/b testing", "customer journey",

    # 环保/可持续
    "solar panel cost", "ev tax credit", "carbon footprint", "recycling guide", "food waste tips", "sustainable fashion",

    # 家庭/育儿/教育
    "parenting tips", "儿童绘本", "早教 游戏", "儿童编程", "数学思维训练", "科学小实验", "亲子旅行",

    # 求职/证书/考试
    "leetcode", "system design", "coding interview", "behavioral questions", "简历模板", "公务员考试", "雅思口语",
    "托福词汇", "四六级真题", "CPA 考试", "PMP 备考",

    # 修理/排障
    "windows 蓝屏", "电脑重装", "excel 公式", "word 排版", "打印机故障", "网络延迟 高", "路由器拨号失败",

    # 旅行细分/攻略
    "北海道 自驾", "冲绳 潜水", "普吉岛 浮潜", "巴厘岛 蜜月", "瑞士 火车 pass", "冰岛 自驾环岛",
    "纽约 博物馆", "巴黎 卢浮宫 预约", "伦敦 大英博物馆",

    # 美食细分
    "手冲咖啡", "精品咖啡豆", "烘焙配方", "酵母面包", "寿司做法", "川菜家常", "粤菜 蒸鱼",
    "甜品食谱", "低糖甜品",

    # 运动细分
    "无氧训练计划", "HIIT 训练", "壶铃训练", "普拉提", "瑜伽体式", "羽毛球技巧", "网球发球",
    "游泳换气", "滑雪入门",

    # 小众兴趣/冷知识
    "天文观星", "流星雨时间", "星座故事", "冷知识 随机", "趣味数学", "谜语答案", "桌游推荐",
    "油管趋势", "reddit hot topics",

    # 便民/效率工具
    "markdown 教程", "正则表达式 示例", "表情包下载", "pdf 合并", "图片压缩 在线", "录屏 软件",
    "windows 快捷键", "截图工具",

    # 金融市场细分
    "美联储议息", "cpi 数据", "非农就业", "纳斯达克走势", "标普500 指数", "港股通", "a股行情",
    "黄金价格", "原油价格", "外汇汇率",

    # 城市/本地出行
    "共享单车 价格", "打车券", "停车收费", "新能源 绿牌 政策", "限行 查询",

    # 购物指南/测评
    "显示器 推荐", "蓝牙耳机 评测", "降噪耳机 对比", "行李箱 20寸", "旅行背包 30L",
    "机械键盘 热插拔", "人体工学 椅子", "颈枕 评测",

    # 效率与组织
    "gtd 方法", "bullet journal", "第二大脑 笔记", "obsidian 模板", "notion dashboard",

    # 编程工具链
    "vscode shortcuts", "git rebase", "git stash 用法", "docker compose 示例", "kubernetes 基础",
    "helm chart", "ci cd pipeline", "github actions",

    # 设计/创意
    "color palette generator", "font pairing", "icon pack free", "figma tutorial", "ui kit", "logo 设计 灵感",

    # 语言学习细分
    "西语基础", "法语入门", "日语五十音", "韩语发音", "德语词汇", "西班牙语听力", "英语口音练习",

    # 生活小技巧/安全
    "房屋除湿", "驱蚊 方法", "急救常识", "心肺复苏 cpr", "家庭应急包", "网络诈骗 识别",

    # 宠物/园艺
    "猫咪驱虫", "狗粮 成分", "猫砂 选择", "多肉养护", "阳台种菜", "花卉浇水 频率",

    # 文化/艺术扩展
    "国画 入门", "书法 笔画", "古典音乐 曲目", "钢琴名曲", "艺术展 资讯", "摄影后期 教程",
]


def generate_random_query() -> str:
    """生成随机搜索词（关键词 + 随机串）。"""
    keyword = random.choice(SEARCH_KEYWORDS)
    random_string = "".join(random.choices("abcdefghijklmnopqrstuvwxyz0123456789", k=8))
    return f"{keyword} {random_string}"


async def fetch_rewards_userinfo(context, *, timeout_ms: int = 15000):
    """从 Rewards API 拉取用户信息（使用当前 context 的 Cookie 身份）。

    增加 timeout 以避免网络抖动时请求悬挂导致界面“卡住”。
    """
    timestamp = int(time.time() * 1000)
    url = f"{REWARDS_API_URL}&_={timestamp}"

    try:
        resp = await context.request.get(
            url,
            headers={
                "X-Requested-With": "XMLHttpRequest",
            },
            timeout=timeout_ms,
        )
    except Exception as e:
        raise RuntimeError(f"Rewards API 请求超时/失败: {e}") from e

    if not resp.ok:
        raise RuntimeError(f"Rewards API 请求失败: {resp.status} {resp.status_text}")
    return await resp.json()


def _sum_counter_points(counter_items):
    current = 0
    maximum = 0
    if not counter_items:
        return 0, 0
    for item in counter_items:
        # JS: pointProgress / pointMax 或 pointProgressMax
        current += int(item.get("pointProgress") or 0)
        maximum += int(item.get("pointMax") or item.get("pointProgressMax") or 0)
    return current, maximum


def compute_remaining_searches(userinfo: dict, device_type: str) -> dict:
    """根据 Rewards counters 计算当前设备类型剩余搜索次数。

    device_type: "windows" (PC) 或 "iphone" (Mobile)
    """
    dashboard = userinfo.get("dashboard") or userinfo
    user_status = (dashboard or {}).get("userStatus") or {}
    counters = user_status.get("counters") or {}

    pc_current, pc_max = _sum_counter_points(counters.get("pcSearch") or [])
    mobile_current, mobile_max = _sum_counter_points(counters.get("mobileSearch") or [])

    if device_type == "windows":
        current_points = pc_current
        max_points = pc_max
        device_label = "电脑"
    else:
        current_points = mobile_current
        max_points = mobile_max
        device_label = "手机"

    # 若 counters 完全为空（无法获取数据），就用一个保守的兜底值
    # 但如果 counters 存在但 max_points 为 0，说明该设备不需要做任务，保持为 0
    if max_points <= 0 and not counters:
        max_points = 90

    remaining_points = max(0, max_points - current_points)
    points_per_search = SEARCH_CONFIG["points_per_search"]
    remaining_searches = (remaining_points + points_per_search - 1) // points_per_search if max_points > 0 else 0
    total_searches = (max_points + points_per_search - 1) // points_per_search if max_points > 0 else 0

    return {
        "device_label": device_label,
        "pc": {"current": pc_current, "max": pc_max},
        "mobile": {"current": mobile_current, "max": mobile_max},
        "current_points": current_points,
        "max_points": max_points,
        "remaining_points": remaining_points,
        "remaining_searches": int(remaining_searches),
        "total_searches": int(total_searches),
    }


def _clear_console() -> None:
    """清空控制台输出（在 Windows Terminal/PowerShell 下尽量兼容）。"""
    if not sys.stdout.isatty():
        return

    # 优先 ANSI（Windows Terminal/新 PowerShell 支持）
    try:
        sys.stdout.write("\x1b[2J\x1b[H")
        sys.stdout.flush()
        return
    except Exception:
        pass

    # 兜底：cls
    try:
        os.system("cls")
    except Exception:
        pass


def _supports_ansi() -> bool:
    """判断当前控制台是否大概率支持 ANSI 光标控制。"""
    if not sys.stdout.isatty():
        return False
    # Windows Terminal / VS Code terminal / 新版 PowerShell 通常支持
    # 这里不做过度检测，失败时会 fallback 到 clear。
    return True


def _bar(pct: float, width: int = 20) -> str:
    pct = max(0.0, min(100.0, pct))
    filled = int(round(width * (pct / 100.0)))
    return "[" + ("#" * filled) + ("-" * (width - filled)) + "]"


def _pct(numerator: float, denominator: float) -> float:
    if denominator <= 0:
        # 当分母为0时，如果分子也为0，表示任务不需要做，返回100%
        return 100.0 if numerator <= 0 else 0.0
    return max(0.0, min(100.0, (numerator / denominator) * 100.0))


def _fmt_duration(seconds: float) -> str:
    """格式化运行时长为易读字符串"""
    seconds = int(max(0, seconds))
    hours, rem = divmod(seconds, 3600)
    minutes, secs = divmod(rem, 60)
    if hours >= 24:
        days, hours = divmod(hours, 24)
        return f"{days}d {hours:02d}:{minutes:02d}:{secs:02d}"
    return f"{hours:02d}:{minutes:02d}:{secs:02d}"


def _format_console_dashboard(
    *,
    userinfo: dict | None,
    stats: dict | None,
    status_text: str,
    search_index: int | None,
    search_total: int | None,
) -> str:
    now_str = time.strftime("%H:%M:%S")
    runtime_str = _fmt_duration(time.time() - APP_START_TS)

    def _search_progress(current: int, maximum: int) -> tuple[int, int]:
        points_per_search = SEARCH_CONFIG["points_per_search"]
        maximum = max(0, int(maximum))
        current = max(0, int(current))
        total = (maximum + points_per_search - 1) // points_per_search if points_per_search > 0 else 0
        remaining = max(0, maximum - current)
        remaining_searches = (remaining + points_per_search - 1) // points_per_search if points_per_search > 0 else 0
        done = max(0, total - remaining_searches)
        return done, total

    if userinfo is None or stats is None:
        lines = [
            f"当前时间:{now_str}  运行时长: {runtime_str}",
            "=" * 60,
            f"状态: {status_text}",
            "=" * 60,
        ]
        return "\n".join(lines)

    dashboard = userinfo.get("dashboard") or userinfo
    user_status = (dashboard or {}).get("userStatus") or {}
    level_info = user_status.get("levelInfo") or {}
    level_name = level_info.get("activeLevel") or level_info.get("level") or "未知"
    total_points = int(user_status.get("availablePoints") or 0)

    pc_cur = stats["pc"]["current"]
    pc_max = stats["pc"]["max"]
    m_cur = stats["mobile"]["current"]
    m_max = stats["mobile"]["max"]

    today_points = pc_cur + m_cur
    today_max = (pc_max or 0) + (m_max or 0)

    today_pct = _pct(today_points, today_max)
    pc_pct = _pct(pc_cur, pc_max)
    m_pct = _pct(m_cur, m_max)

    pc_done, pc_total = _search_progress(pc_cur, pc_max)
    m_done, m_total = _search_progress(m_cur, m_max)

    device_label = stats.get("device_label")

    def _status_line(label: str, done: int, total: int) -> str:
        # 统一展示为“状态: 剩余 N 次”，不显示 (进度) 或其它文案
        remaining_count = max(0, int(total) - int(done))
        return f"状态: 剩余 {remaining_count} 次"

    lines = [
        f"当前时间:{now_str}  运行时长: {runtime_str}",
        "=" * 60,
        f"等级: {level_name}",
        f"总积分: {total_points}",
        f"今日获取: {today_points} / {today_max}  {_bar(today_pct)} {today_pct:6.1f}%",
        f"电脑: {pc_cur} / {pc_max}  {_bar(pc_pct)} {pc_pct:6.1f}%",
        _status_line("电脑", pc_done, pc_total),
        f"手机: {m_cur} / {m_max}  {_bar(m_pct)} {m_pct:6.1f}%",
        _status_line("手机", m_done, m_total),
        "=" * 60,
    ]
    return "\n".join(lines)


def _render_console_dashboard(text: str, *, use_ansi: bool, first_paint: bool) -> None:
    """把仪表盘文本渲染到控制台（原地更新）。"""
    if not sys.stdout.isatty():
        print(text, end="")
        return

    if use_ansi:
        # 首次输出前清屏一次，后续用“回到左上角 + 清到屏末”实现无闪烁刷新
        if first_paint:
            sys.stdout.write("\x1b[2J\x1b[H")
        else:
            sys.stdout.write("\x1b[H\x1b[J")
        sys.stdout.write(text)
        sys.stdout.flush()
        return

    # fallback：不支持 ANSI 时就用清屏（可能闪一下）
    _clear_console()
    print(text, end="")


async def console_dashboard_refresh_loop(context, device_type: str, state: dict):
    """定时刷新 Rewards 数据，供控制台仪表盘显示。"""
    while True:
        try:
            userinfo = await fetch_rewards_userinfo(context)
            stats = compute_remaining_searches(userinfo, device_type)
            state["userinfo"] = userinfo
            state["stats"] = stats
            state["error"] = None
            state["last_fetch_ts"] = time.time()
        except Exception as e:
            state["error"] = str(e)

        await asyncio.sleep(DASHBOARD_REFRESH_SECONDS)


async def run_rewards_auto_search(page, context, device_type: str):
    """自动完成 Bing 搜索任务（等价于 userscript 的 startSearchTask）。"""
    # 控制台仪表盘：启动后台刷新
    state = {"userinfo": None, "stats": None, "error": None}
    refresh_task = asyncio.create_task(console_dashboard_refresh_loop(context, device_type, state))

    status_text = "准备..."
    use_ansi = _supports_ansi()
    first_paint = True
    print("\n正在获取 Rewards 搜索进度...")
    try:
        userinfo = await fetch_rewards_userinfo(context)
        stats = compute_remaining_searches(userinfo, device_type)
        state["userinfo"] = userinfo
        state["stats"] = stats
    except Exception as e:
        state["error"] = str(e)
        userinfo = None
        stats = None

    _render_console_dashboard(
        _format_console_dashboard(
            userinfo=state.get("userinfo"),
            stats=state.get("stats"),
            status_text=status_text if not state.get("error") else f"错误: {state['error']}",
            search_index=None,
            search_total=None,
        ),
        use_ansi=use_ansi,
        first_paint=first_paint,
    )
    first_paint = False

    if stats is None:
        refresh_task.cancel()
        try:
            await refresh_task
        except Exception:
            pass
        raise RuntimeError(f"无法获取 Rewards 数据: {state.get('error')}")

    remaining = stats["remaining_searches"]
    total = stats["total_searches"]
    # 用控制台仪表盘展示（不再额外 print 一行）

    # 检查当前设备是否需要执行任务
    if total <= 0 or remaining <= 0:
        status_text = "剩余 0 次"
        _render_console_dashboard(
            _format_console_dashboard(
                userinfo=state.get("userinfo"),
                stats=state.get("stats"),
                status_text=status_text,
                search_index=None,
                search_total=None,
            ),
            use_ansi=use_ansi,
            first_paint=first_paint,
        )
        refresh_task.cancel()
        try:
            await refresh_task
        except Exception:
            pass
        return

    if stats["max_points"] <= 0:
        status_text = "剩余 0 次"
        _render_console_dashboard(
            _format_console_dashboard(
                userinfo=state.get("userinfo"),
                stats=state.get("stats"),
                status_text=status_text,
                search_index=None,
                search_total=None,
            ),
            use_ansi=use_ansi,
            first_paint=first_paint,
        )
        refresh_task.cancel()
        try:
            await refresh_task
        except Exception:
            pass
        return

    delay_cfg = SEARCH_CONFIG["pc"] if device_type == "windows" else SEARCH_CONFIG["mobile"]
    min_ms = delay_cfg["min_delay_ms"]
    max_ms = delay_cfg["max_delay_ms"]

    # 先确保能打开 Bing（有时直接搜索 URL 会被重定向）
    if page.url == "about:blank":
        await page.goto("https://www.bing.com", wait_until=DOMCONTENTLOADED)

    status_text = "开始自动搜索（Ctrl+C 可中断）"
    _render_console_dashboard(
        _format_console_dashboard(
            userinfo=state.get("userinfo"),
            stats=state.get("stats"),
            status_text=status_text,
            search_index=None,
            search_total=None,
        ),
        use_ansi=use_ansi,
        first_paint=first_paint,
    )

    # 对齐 userscript：首次延迟 2 秒再开始（避免一打开就“秒搜”）
    await asyncio.sleep(2)

    for i in range(1, remaining + 1):
        query = generate_random_query()
        search_url = f"https://www.bing.com/search?q={quote_plus(query)}"

        try:
            status_text = f"搜索中... 关键词: {query}"
            _render_console_dashboard(
                _format_console_dashboard(
                    userinfo=state.get("userinfo"),
                    stats=state.get("stats"),
                    status_text=status_text if not state.get("error") else f"{status_text} | 错误: {state['error']}",
                    search_index=None,
                    search_total=None,
                ),
                use_ansi=use_ansi,
                first_paint=first_paint,
            )

            # 更像 userscript：在搜索框里输入并提交（失败才 fallback 到 goto 搜索 URL）
            ok = False
            search_coro = _perform_bing_search_like_human(page, query)
            try:
                ok = await asyncio.wait_for(search_coro, timeout=25)
            except asyncio.TimeoutError:
                status_text = "搜索超时，刷新页面重试"
                _render_console_dashboard(
                    _format_console_dashboard(
                        userinfo=state.get("userinfo"),
                        stats=state.get("stats"),
                        status_text=status_text,
                        search_index=None,
                        search_total=None,
                    ),
                    use_ansi=use_ansi,
                    first_paint=first_paint,
                )
                try:
                    await page.goto("https://www.bing.com", wait_until=DOMCONTENTLOADED, timeout=20000)
                except Exception:
                    pass
                ok = False
            except Exception:
                ok = False

            if not ok:
                await page.goto(search_url, wait_until=DOMCONTENTLOADED)

            # 轻量随机交互：滚动/偶尔点开一个结果再返回
            await _maybe_human_scroll(page)
            await _maybe_click_one_result(page)
        except Exception as e:
            # 网络波动/跳转异常时，尝试回到首页再继续
            status_text = f"搜索跳转失败: {e}"
            _render_console_dashboard(
                _format_console_dashboard(
                    userinfo=state.get("userinfo"),
                    stats=state.get("stats"),
                    status_text=status_text,
                        search_index=None,
                        search_total=None,
                ),
                use_ansi=use_ansi,
                first_paint=first_paint,
            )
            try:
                await page.goto("https://www.bing.com", wait_until=DOMCONTENTLOADED)
            except Exception:
                pass

        # 用 monotonic 做倒计时：更稳定（系统时间调整/睡眠恢复时不容易“看起来卡住”）
        delay_ms = random.randint(min_ms, max_ms)
        end_ts = time.monotonic() + (delay_ms / 1000.0)
        while True:
            remaining_seconds = int((end_ts - time.monotonic()) + 0.999)
            if remaining_seconds <= 0:
                break

            status_text = f"等待 {remaining_seconds} 秒"
            _render_console_dashboard(
                _format_console_dashboard(
                    userinfo=state.get("userinfo"),
                    stats=state.get("stats"),
                    status_text=status_text if not state.get("error") else f"{status_text} | 错误: {state['error']}",
                    search_index=None,
                    search_total=None,
                ),
                use_ansi=use_ansi,
                first_paint=first_paint,
            )

            # 让 UI 每秒刷新；末尾阶段用更短 sleep 以更平滑
            await asyncio.sleep(min(1.0, max(0.1, end_ts - time.monotonic())))

    status_text = "剩余 0 次"
    _render_console_dashboard(
        _format_console_dashboard(
            userinfo=state.get("userinfo"),
            stats=state.get("stats"),
            status_text=status_text,
            search_index=None,
            search_total=None,
        ),
        use_ansi=use_ansi,
        first_paint=first_paint,
    )

    # 可选：结束后再拉一次数据，展示最新进度（失败也不影响主流程）
    try:
        userinfo2 = await fetch_rewards_userinfo(context)
        stats2 = compute_remaining_searches(userinfo2, device_type)
        state["userinfo"] = userinfo2
        state["stats"] = stats2

        # 验证伪装是否成功：如果是 iPhone 模式，检查是否在"移动搜索"计数上增加了
        if device_type == "iphone":
            before_mobile = stats["mobile"]["current"]
            after_mobile = stats2["mobile"]["current"]
            if after_mobile > before_mobile:
                spoof_result = f"✓ 手机伪装成功（移动搜索：{before_mobile} → {after_mobile}）"
            else:
                spoof_result = f"✗ 伪装可能失败（移动搜索未增加：{before_mobile} / {stats2['mobile']['max']}）"
        else:
            spoof_result = ""

        _render_console_dashboard(
            _format_console_dashboard(
                userinfo=state.get("userinfo"),
                stats=state.get("stats"),
                status_text=f"剩余 0 次",
                search_index=None,
                search_total=None,
            ),
            use_ansi=use_ansi,
            first_paint=first_paint,
        )
    except Exception as e:
        # 只在控制台提示，不抛出
        _render_console_dashboard(
            _format_console_dashboard(
                userinfo=state.get("userinfo"),
                stats=state.get("stats"),
                status_text=f"剩余 0 次",
                search_index=None,
                search_total=None,
            ),
            use_ansi=use_ansi,
            first_paint=first_paint,
        )

    # 结束后台刷新
    refresh_task.cancel()
    try:
        await refresh_task
    except Exception:
        pass


def get_cookie_files():
    """获取所有已保存的cookie文件"""
    # 优先从可写目录读取（exe 同目录），其次从打包资源目录读取（onefile 临时解压）
    exe_cookies_dir = _get_exe_dir() / "Assets" / "cookies"
    bundle_cookies_dir = _get_bundle_dir() / "Assets" / "cookies"

    cookie_files: list[Path] = []
    seen: set[Path] = set()
    for d in [exe_cookies_dir, bundle_cookies_dir]:
        try:
            if d.exists():
                for p in d.glob("*.json"):
                    rp = p.resolve()
                    if rp not in seen:
                        seen.add(rp)
                        cookie_files.append(p)
        except Exception:
            # 目录不可访问就跳过
            continue

    return cookie_files


def select_cookie_file():
    """让用户选择cookie文件"""
    cookie_files = get_cookie_files()

    # 去重：同名账号优先选择 exe 同目录下的（可写、可更新）
    by_name: dict[str, Path] = {}
    exe_dir = _get_exe_dir().resolve()
    for p in cookie_files:
        name = p.stem
        if name not in by_name:
            by_name[name] = p
            continue
        try:
            # 若当前是 exe 同目录，覆盖掉 bundle 中的同名文件
            if p.resolve().is_relative_to(exe_dir):
                by_name[name] = p
        except Exception:
            # Python < 3.9 或路径解析失败时，保持现状
            pass

    cookie_files = sorted(by_name.values(), key=lambda x: x.stem.lower())
    
    if not cookie_files:
        print("\n未找到已保存的cookie文件!")
        return None
    
    print("\n可用的账号:")
    for i, file in enumerate(cookie_files, 1):
        account_name = file.stem  # 获取不带扩展名的文件名
        print(f"{i}. {account_name}")
    
    while True:
        try:
            choice = input("\n请选择账号 (输入序号): ").strip()
            index = int(choice) - 1
            if 0 <= index < len(cookie_files):
                return cookie_files[index]
            else:
                print("无效的选择，请重新输入!")
        except ValueError:
            print("请输入有效的数字!")
        except KeyboardInterrupt:
            return None


def select_device_type():
    """让用户选择设备类型"""
    print("\n选择设备类型:")
    print("1. 电脑 (Windows UA)")
    print("2. 手机 (iPhone UA)")
    
    while True:
        try:
            choice = input("\n请选择设备类型 (输入序号): ").strip()
            if choice == "1":
                return "windows"
            elif choice == "2":
                return "iphone"
            else:
                print("无效的选择，请重新输入!")
        except KeyboardInterrupt:
            return None


async def login_and_save():
    """登录并保存cookie"""
    # 注意：onefile 打包下 __file__ 在临时目录，写入会丢失；所以固定写到 exe 同目录。
    cookies_dir = _get_exe_dir() / "Assets" / "cookies"
    cookies_dir.mkdir(parents=True, exist_ok=True)
    
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            channel=MSEDGE_CHANNEL,
            headless=False
        )
        
        context = await browser.new_context()
        page = await context.new_page()
        
        try:
            print("正在打开 Bing Rewards 页面...")
            await page.goto("https://rewards.bing.com/?ref=rewardspanel")
            
            print("\n请在浏览器中登录您的Microsoft账户...")
            print("登录完成后，按回车键保存Cookie...")
            
            loop = asyncio.get_running_loop()
            await loop.run_in_executor(None, input)
            
            # 获取账户名
            account_name = None
            
            try:
                account_xpath = '//*[@id="mectrl_currentAccount_secondary"]'
                account_element = await page.query_selector(f"xpath={account_xpath}")
                if account_element:
                    text = await account_element.text_content()
                    if text and text.strip():
                        account_name = text.strip()
                        print(f"\n检测到账户: {account_name}")
            except Exception as e:
                print(f"方法1失败: {e}")
            
            if not account_name:
                try:
                    account_element = await page.query_selector("#mectrl_currentAccount_secondary")
                    if account_element:
                        text = await account_element.text_content()
                        if text and text.strip():
                            account_name = text.strip()
                            print(f"\n检测到账户: {account_name}")
                except Exception as e:
                    print(f"方法2失败: {e}")
            
            if not account_name:
                try:
                    account_name = await page.evaluate('''() => {
                        const el = document.getElementById("mectrl_currentAccount_secondary");
                        return el ? el.textContent.trim() : null;
                    }''')
                    if account_name:
                        print(f"\n检测到账户: {account_name}")
                except Exception as e:
                    print(f"方法3失败: {e}")
            
            if not account_name:
                account_name = "未知账户"
                print("\n警告: 未能检测到账户名，使用默认名称")

            account_name = _sanitize_filename(account_name)
            
            cookies = await context.cookies()
            cookie_file = cookies_dir / f"{account_name}.json"
            
            with open(cookie_file, 'w', encoding='utf-8') as f:
                json.dump(cookies, f, ensure_ascii=False, indent=2)
            
            print(f"Cookie已保存到: {cookie_file}")
            await asyncio.sleep(1)
            
        except Exception as e:
            print(f"\n错误: {str(e)}")
        finally:
            try:
                await browser.close()
            except Exception as e:
                print(f"关闭浏览器时出现异常（可忽略）: {e}")
            print("浏览器已关闭")


async def use_saved_cookie():
    """使用已保存的cookie访问Bing"""
    # 选择cookie文件
    cookie_file = select_cookie_file()
    if not cookie_file:
        print("\n操作已取消")
        return
    
    # 选择设备类型
    device_type = select_device_type()
    if not device_type:
        print("\n操作已取消")
        return
    
    # 加载cookie
    with open(cookie_file, 'r', encoding='utf-8') as f:
        cookies = json.load(f)
    
    # 获取User-Agent（如果是iphone类型，会随机选择一个）
    ua_value = USER_AGENTS[device_type]
    user_agent = ua_value() if callable(ua_value) else ua_value
    
    print(f"\n已选择账号: {cookie_file.stem}")
    print(f"设备类型: {'Windows' if device_type == 'windows' else 'iPhone'}")
    print(f"User-Agent: {user_agent}")
    
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            channel="msedge",
            headless=False
        )
        
        # 创建上下文时设置 user agent
        # 说明：在 Windows 上用 Chromium(Edge) 不能“真正变成 iOS Safari/WebKit”，但可以：
        # - 让站点看到 iPhone UA
        # - 同时启用 is_mobile/has_touch/device_scale_factor 等移动特征
        # 这比只改 UA 更一致。

        if device_type == "windows":
            context_kwargs = {
                "user_agent": user_agent,
                "viewport": {"width": 1920, "height": 1080},
                "screen": {"width": 1920, "height": 1080},
                # 可选：更贴近中文用户环境（不影响 cookie 身份）
                "locale": "zh-CN",
            }
        else:
            # 优先使用 Playwright 内置设备配置（能一次性带上 viewport/is_mobile/has_touch 等）
            # 设备名会随 Playwright 版本略有差异，所以做 KeyError 兜底。
            device_profile = None
            for name in ["iPhone 15", "iPhone 14", "iPhone 13", "iPhone 12"]:
                try:
                    device_profile = p.devices[name]
                    break
                except Exception:
                    continue

            context_kwargs = dict(device_profile or {})
            # 保持你自定义的 iPhone UA（更贴近你设定的版本），覆盖 profile 默认 UA
            context_kwargs["user_agent"] = user_agent

            # 视口与屏幕（screen 不一定必须，但有助于减少“桌面窗口”味道）
            context_kwargs.setdefault("viewport", {"width": 390, "height": 844})
            context_kwargs.setdefault("screen", {"width": 390, "height": 844})

            # 显式打开移动特征
            context_kwargs["is_mobile"] = True
            context_kwargs["has_touch"] = True
            context_kwargs["device_scale_factor"] = 3

            # 移动端本地化设置
            context_kwargs["locale"] = "zh-CN"
            context_kwargs["timezone_id"] = "Asia/Shanghai"

        context = await browser.new_context(**context_kwargs)
        
        # 如果是 iPhone 模式，在context级别注入 JavaScript 伪装（在所有页面生效）
        if device_type == "iphone":
            await _inject_mobile_spoofing(context)
        
        # 添加cookies
        await context.add_cookies(cookies)
        
        page = await context.new_page()
        
        try:
            print("\n正在打开 Bing 页面...")
            await page.goto("https://www.bing.com", wait_until=DOMCONTENTLOADED)

            # 选择 UA 后自动开始（移植自 Microsoft Rewards Dashboard.js）
            await run_rewards_auto_search(page, context, device_type)

            print("\n脚本执行结束，按回车键关闭浏览器...")
            loop = asyncio.get_running_loop()
            await loop.run_in_executor(None, input)
            
        except Exception as e:
            print(f"\n错误: {str(e)}")
        finally:
            try:
                await browser.close()
            except Exception as e:
                print(f"关闭浏览器时出现异常（可忽略）: {e}")
            print("浏览器已关闭")


async def main():
    """主函数"""
    print("\n请选择操作:")
    print("1. 登录并保存Cookie")
    print("2. 使用并开始每日任务")
    
    while True:
        try:
            choice = input("\n请选择操作 (输入序号): ").strip()
            if choice == "1":
                await login_and_save()
                break
            elif choice == "2":
                await use_saved_cookie()
                break
            else:
                print("无效的选择，请重新输入!")
        except KeyboardInterrupt:
            print("\n\n操作已取消")
            break


if __name__ == "__main__":
    print("=" * 60)
    print("My QQ 3183670554")
    print("=" * 60)
    asyncio.run(main())
