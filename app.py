"""
Interrogation Experiment Web App
Flask backend for suspect/interviewer assignment and training management.
All experiment statistics stored in a single Excel file (experiment_data.xlsx).
"""
import json
import logging
import os
import random
import shutil
import re
import string
import threading
import socket
import io
import uuid
from datetime import datetime, timedelta
from functools import wraps
from dataclasses import dataclass
from typing import Literal

import requests
import openpyxl
from openpyxl import Workbook, load_workbook
from flask import Flask, jsonify, render_template, request, send_file, session
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY")
if not app.secret_key:
    raise RuntimeError("SECRET_KEY environment variable is required")

BASE_DIR = os.path.dirname(__file__)
DATA_DIR = os.environ.get("DATA_DIR", os.path.join(BASE_DIR, "data"))
os.makedirs(DATA_DIR, exist_ok=True)
EXCEL_FILE = os.environ.get(
    "EXCEL_FILE",
    os.path.join(DATA_DIR, "experiment_data.xlsx"),
)
# Legacy path when Excel lived at project root (pre-Docker data/ layout)
LEGACY_EXCEL_FILE = os.path.join(BASE_DIR, "experiment_data.xlsx")
_excel_lock = threading.Lock()

DEEPSEEK_API_KEY = os.environ.get("DEEPSEEK_API_KEY")
LIVEAVATAR_API_KEY = os.environ.get("LIVEAVATAR_API_KEY")
ELEVENLABS_API_KEY = os.environ.get("ELEVENLABS_API_KEY")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD")
MANAGE_ENTRY_PASSWORD = "77585211314"
DEBUG_MODE = os.environ.get("FLASK_DEBUG", "false").lower() == "true"

# Validate required environment variables
_missing = []
if not DEEPSEEK_API_KEY:
    _missing.append("DEEPSEEK_API_KEY")
if not LIVEAVATAR_API_KEY:
    _missing.append("LIVEAVATAR_API_KEY")
if not ELEVENLABS_API_KEY:
    _missing.append("ELEVENLABS_API_KEY")
if not ADMIN_PASSWORD:
    _missing.append("ADMIN_PASSWORD")
if _missing:
    raise RuntimeError(f"Missing required environment variables: {', '.join(_missing)}")

# ---- Logging ----
LOG_DIR = os.environ.get("LOG_DIR", os.path.join(BASE_DIR, "logs"))
os.makedirs(LOG_DIR, exist_ok=True)
_log_handlers = [logging.StreamHandler()]
_log_file = os.path.join(LOG_DIR, "app.log")
try:
    _log_handlers.append(logging.FileHandler(_log_file, encoding="utf-8"))
except OSError as e:
    print(f"[WARN] Cannot write log file {_log_file}: {e}; using stdout only.", flush=True)
logging.basicConfig(
    level=logging.DEBUG if DEBUG_MODE else logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=_log_handlers,
)
logger = logging.getLogger(__name__)

# ---- Security: Rate Limiting ----
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address

limiter = Limiter(
    get_remote_address,
    app=app,
    default_limits=["200 per day", "50 per hour"],
    storage_uri="memory://",
)

# ---- Security: HTTP Headers ----
from flask_talisman import Talisman

Talisman(
    app,
    content_security_policy={
        "default-src": "'self'",
        "script-src": "'self' 'unsafe-inline' https://cdn.jsdelivr.net https://unpkg.com https://www.youtube.com",
        "style-src": "'self' 'unsafe-inline'",
        "img-src": "'self' data: blob:",
        "connect-src": "'self' https://api.deepseek.com https://api.liveavatar.com https://api.elevenlabs.io https://api.heygen.com https://webrtc-signaling.heygen.io wss://webrtc-signaling.heygen.io wss://*.heygen.io wss://*.liveavatar.com wss://*.livekit.io https://*.livekit.cloud wss://*.livekit.cloud https://cdn.jsdelivr.net",
        "font-src": "'self'",
        "media-src": "'self' blob: data:",
        "frame-src": "'self' https://www.youtube.com",
    },
    force_https=False,
    session_cookie_secure=False,
    session_cookie_http_only=True,
)

# ---- Security: Input Sanitization ----
import bleach

def sanitize(value, max_length=5000):
    """Sanitize user input."""
    if value is None:
        return ""
    value = str(value).strip()
    if len(value) > max_length:
        value = value[:max_length]
    return bleach.clean(value, tags=[], strip=True)

# ---- Security: Admin Authentication ----
def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        auth = request.authorization
        valid_passwords = {ADMIN_PASSWORD, MANAGE_ENTRY_PASSWORD}
        if not auth or auth.username != "admin" or auth.password not in valid_passwords:
            return jsonify({"error": "Unauthorized"}), 401, \
                   {"WWW-Authenticate": 'Basic realm="Admin Area"'}
        return f(*args, **kwargs)
    return decorated

# ---- Security: CSRF Protection for Admin Routes ----
@app.before_request
def csrf_protect():
    if request.method in ("POST", "PUT", "PATCH", "DELETE"):
        if request.path.startswith("/api/admin/"):
            if not request.headers.get("X-Requested-With"):
                logger.warning(f"CSRF check failed for {request.path} from {request.remote_addr}")
                return jsonify({"error": "CSRF validation failed"}), 403

MATERIALS_DIR = os.path.join(BASE_DIR, "materials")
MATERIALS_PROMPTS_DIR = os.path.join(MATERIALS_DIR, "prompts")
COMBINED_MATERIALS_MD = os.path.join(MATERIALS_DIR, "combined_materials.md")
COMBINED_MATERIALS_DOCX = os.path.join(MATERIALS_DIR, "combined_materials.docx")
COMBINED_DOWNLOAD_NAME = "培训材料合集.docx"
FEEDBACK_PROMPT_MD = os.path.join(MATERIALS_PROMPTS_DIR, "avatar_feedback.md")
LEGACY_FEEDBACK_DOCX = os.path.join(BASE_DIR, "Feedback prompts.docx")
TRAINING_SRC_DIR = os.path.join(os.path.dirname(os.path.dirname(BASE_DIR)), "IRB materials", "Training materials")
LEGACY_SERIOUS_GAME_XLSX = os.path.join(BASE_DIR, "results_serious_game.xlsx")
LEGACY_TRAINING_FEEDBACK_DIR = os.path.join(BASE_DIR, "training_feedback")

ARSON_GUILTY_CONTEXT = """你是当地社区体育俱乐部的一名资深成员。昨天（周五）晚上，你参加了俱乐部的月度例会。然而，在会议上，委员会不仅公开且毫无理由地否决了你花费六个月准备的提案，还用一个荒谬的借口暂停了你的核心会员资格。走出会议室时，你感受到前所未有的羞辱与愤怒。你为俱乐部付出了这么多，却只换来背叛。你决定让他们付出代价——那座标志性的体育馆是委员会最引以为傲的资产，而你打算将其摧毁。

你的计划是在周六深夜潜入并烧毁体育馆。你利用周六晚的时间构建了一个看似完美的不在场证明。

第二天，社区俱乐部的体育馆在周六深夜被恶意纵火摧毁。由于你是俱乐部的资深成员，且警方在你当晚的行动中发现了一些"可疑"之处，你现已被列为该纵火案的主要嫌疑人。

此刻，你正坐在警局冰冷的讯问室中。主审调查员即将推门而入。无论你当晚实际做了什么，你现在唯一的目标是：仔细回忆经历的每一个细节，为警方提出的每一项证据提供合理且严密的解释，并竭尽一切说服他们你完全无辜。"""

ARSON_INNOCENT_CONTEXT = """你是当地社区体育俱乐部的一名资深成员。昨天（周五）晚上，你参加了俱乐部的月度例会。整场会议充满了委员会成员之间冗长的争论和繁琐的流程，让你感到极度疲惫和烦躁。走出会议室时，你只觉得身心俱疲。你不想再参与俱乐部的纷争，只想彻底放空自己。

周六你决定今晚完全属于自己，在家度过一个极其懒散、平静的夜晚。除了稍晚一点开车去俱乐部附近的一家酒吧接一位朋友外，你没有其他任何安排。

第二天，社区俱乐部的体育馆在周六深夜被恶意纵火摧毁。由于你是俱乐部的资深成员，且警方在你当晚的行动中发现了一些"可疑"之处，你现已被列为该纵火案的主要嫌疑人。

此刻，你正坐在警局冰冷的讯问室中。你知道自己绝对没有纵火，但你仍然需要解释清楚当晚的所有细节。"""

THEFT_GUILTY_CONTEXT = """你最近手头比较紧。当你听说当地帆船俱乐部要举办开放日时，你并没有把它当作一个轻松的周末活动，而是一个机会。你做了充分准备，查阅了俱乐部的比赛安排，并注意到"Sea Whisper"（停靠在C码头12号泊位的一艘昂贵帆船）的船主整个下午都会外出参赛。

你利用人群的掩护，用现金支付咖啡以避免留下数字痕迹，悄无声息地登上"Sea Whisper"，偷走了一副价值约200英镑的Steiner 8×30望远镜。你将望远镜藏好，顺利离开码头，没有人对你产生怀疑。

开放日已经过去几天。警方正在调查这起盗窃案，并对当天在码头出现的人进行例行询问。不要承认犯罪。你的目标是在警方调查压力下，维持一个普通无辜访客的形象。"""

THEFT_INNOCENT_CONTEXT = """你一直对航海和海洋充满热情，尽管你自己并没有船。当你听说当地帆船俱乐部要举办年度开放日时，你觉得这是一个在阳光明媚的下午放松的好机会。你只是想感受海风，体验热闹的海事氛围。

你度过了一个愉快的下午。你用现金买了咖啡，和其他游客一起沿着C码头散步。你特别停下来欣赏了停在12号泊位的"Sea Whisper"，因为它的设计很经典。你只是站在木质码头上观看，从未登船，最后带着放松的心情回家。

开放日已经过去几天。你突然接到警方电话——"Sea Whisper"上发生了一起盗窃案。由于你曾在C码头附近观赏船只，并出现在俱乐部照片的背景中，警方将你视为潜在证人或相关人员。你没有任何需要隐瞒的事情，但面对警方可能仍会有些紧张。只需要证明你是一个前来游玩的普通公民即可。"""

ATTENTION_CHECKS = {
    "arson_guilty": [
        {
            "question": "根据背景故事，周五晚上的会议上发生了什么导致你产生报复心理？",
            "options": ["A. 你的钱包在会议室被其他成员偷走了", "B. 委员会毫无理由地否决了你的提案并暂停了你的核心会员资格", "C. 你被迫承担俱乐部所有债务"],
            "answer": 1,
        },
        {
            "question": "在你的周六晚上计划中，为确保复仇顺利进行，你需要完成的一项关键任务是什么？",
            "options": ["A. 构建一个证明你'整晚都在家'的完美不在场证明", "B. 开车去俱乐部附近酒吧接朋友", "C. 提前在社交媒体上发布退出俱乐部的声明"],
            "answer": 0,
        },
    ],
    "arson_innocent": [
        {
            "question": "根据背景故事，周五晚上的月度例会结束时你的感受是什么？",
            "options": ["A. 感到前所未有的羞辱，想要纵火报复", "B. 感到极度疲惫和烦躁，只想彻底放空自己", "C. 因为你的提案被通过而感到非常兴奋"],
            "answer": 1,
        },
        {
            "question": "周六晚上除了在家度过平静时光外，你唯一的外出安排是什么？",
            "options": ["A. 潜入体育俱乐部烧毁体育馆", "B. 去附近超市购买周末零食", "C. 开车去俱乐部附近的酒吧接一位朋友"],
            "answer": 2,
        },
    ],
    "theft_guilty": [
        {
            "question": "根据背景故事，你参加帆船俱乐部开放日的主要动机是什么？",
            "options": ["A. 为了感受海风，体验热闹的航海氛围", "B. 这是一个潜入帆船并窃取贵重物品以快速变现的机会", "C. 为了作为选手参加俱乐部举办的帆船比赛"],
            "answer": 1,
        },
        {
            "question": "在你执行的计划中，你对'Sea Whisper'采取了什么具体行动？",
            "options": ["A. 只站在木质码头上观看船的造型，从未登船", "B. 悄悄登船并偷走了一副价值约200英镑的Steiner望远镜", "C. 你登船在甲板上喝了杯咖啡然后就离开了"],
            "answer": 1,
        },
    ],
    "theft_innocent": [
        {
            "question": "你在码头散步时，对停靠在12号泊位的'Sea Whisper'做了什么？",
            "options": ["A. 趁人不注意登船偷了一副望远镜", "B. 只站在木质码头上观看船的造型，从未踏上帆船一步", "C. 你登上船与船长进行了交谈"],
            "answer": 1,
        },
        {
            "question": "根据背景故事，你当天是如何支付咖啡费用的，以及警方为什么找到你？",
            "options": ["A. 你用现金支付了咖啡，警方找到你是因为你出现在C码头附近的一张社交媒体照片背景中", "B. 你用信用卡支付了咖啡，警方通过消费记录追踪到你", "C. 你没有买咖啡，警方在帆船甲板上发现了你的指纹"],
            "answer": 0,
        },
    ],
}

CONTROL_ATTENTION_CHECKS = [
    {
        "question": "根据培训手册，访谈开始时与嫌疑人进行日常对话（如询问基本信息）的主要目的是什么？",
        "options": [
            "A. 为了让嫌疑人放松警惕，以便后续出示证据。",
            "B. 为了观察嫌疑人在无压力状态下的正常反应，从而建立“行为基线”。",
            "C. 为了向嫌疑人展示访谈员的专业性。",
            "D. 为了拖延时间，等待同事核实案情。",
        ],
        "answer": 1,
    },
    {
        "question": "根据手册中“识别欺骗信号”的指导，以下哪项行为在嫌疑人回答核心问题时，最有可能被视为试图隐瞒真相的可疑信号？",
        "options": [
            "A. 语速平稳，直视访谈员并迅速给出具体细节。",
            "B. 在回答前重复访谈员的问题，并频繁使用“我尽量回忆”等模糊词汇。",
            "C. 对访谈员的提问表现出合理的愤怒，并强烈要求联系律师。",
            "D. 身体姿态保持放松，双手平放在桌面上。",
        ],
        "answer": 1,
    },
    {
        "question": "为了确保您的设备能够正常显示文本内容，并且您正在认真阅读本页的指引，请忽略下方列出的所有专业术语，直接选择“苹果”作为您的答案。",
        "options": [
            "A. 逻辑矛盾排查",
            "B. 行为基线",
            "C. 苹果",
            "D. 认知负荷管理",
        ],
        "answer": 2,
    },
]

SUE_EFM_CHECK = {
    "scenario": "您已经在一家中午被抢劫的幸福商铺的柜台上（位于上海市普陀区长寿路）获取了指纹。这个证据表明您将要审讯的嫌疑人当日某时曾去过犯罪现场。您将如何应用EFM（证据框架矩阵）、在审讯中使用这则证据以最大化嫌疑人陈述之间的不一致及陈述与证据之间的不一致？",
    "question": "请将以下三条证据陈述分别归类到正确的EFM类别中（每条证据对应一个类别）：",
    "categories": [
        {"id": "low-low", "label": "低来源低具体", "desc": "不透露证据来源，也不透露证据具体是什么"},
        {"id": "low-high", "label": "低来源高具体", "desc": "不透露证据来源，但透露证据具体是什么"},
        {"id": "high-high", "label": "高来源高具体", "desc": "既透露证据来源，也透露证据具体是什么"},
    ],
    "statements": [
        {"text": "我们有证据表明你去过上海市普陀区", "answer": "low-low"},
        {"text": "我们有证据表明你去过上海市普陀区长寿路的幸福商铺", "answer": "low-high"},
        {"text": "我们在幸福商铺的柜台上提取到了你的指纹", "answer": "high-high"},
    ],
}

CASE_INFO = {
    "arson": {
        "title": "案件 B — 体育俱乐部纵火案",
        "overview": """背景：某社区体育俱乐部的场馆于周日凌晨被恶意纵火烧毁。火灾调查员确认使用了助燃剂（碳氢化合物，与烧烤点火液成分一致）。火灾大约在 23:15 开始，场馆完全损毁。大火摧毁了现场所有可能的在场痕迹。嫌疑人为俱乐部会员。声称当晚在家，并在 Instagram Story 上发布了动态。ta当天早些时候进行过户外烧烤或篝火活动。""",
        "evidence": [
            {
                "id": "E1",
                "title": "Instagram Story 帖子 — EXIF 时间戳不一致",
                "detail": "嫌疑人在 22:47 向 Instagram Story 发布了一张展示室内居家场景的照片，表明当晚在家。调查人员检查图像文件后发现，嵌入的 EXIF 元数据记录该照片拍摄时间为 19:22。",
            },
            {
                "id": "E2",
                "title": "外套上的助燃剂痕迹",
                "detail": "对嫌疑人当天所穿外套的化学分析发现碳氢化合物助燃剂痕迹，与烧烤点火液成分一致。嫌疑人当天曾使用点火液进行烧烤。",
            },
            {
                "id": "E3",
                "title": "手机基站注册 — 青山路 基站",
                "detail": "手机记录显示嫌疑人的手机于 23:29 在 青山路 基站注册。该基站的主要覆盖范围包括体育场馆及周围约 800 米半径。该注册记录确认了手机当晚的大致位置范围。",
            },
        ],
        "efm_analysis": "",
    },
    "theft": {
        "title": "案件 A — 帆船俱乐部开放日盗窃案",
        "overview": """背景：某小型帆船俱乐部举办年度开放日。一副紧凑型望远镜（Steiner 8×30，价值约 £200）从停靠在 C 码头 12 号泊位的 "Sea Whisper" 号帆船上被盗。船主外出参赛，船只未上锁。

嫌疑人参加了当天的开放日活动。无登记记录，无金融交易痕迹。ta曾在码头咖啡厅用现金购买了咖啡（无银行卡记录），随其他游客一起沿 C 码头散步，并出现在俱乐部的社交媒体活动照片中。""",
        "evidence": [
            {
                "id": "E1",
                "title": "咖啡厅工作人员回忆",
                "detail": "工作人员回忆起一名与嫌疑人描述相符的顾客用现金支付了咖啡，并询问哪个码头有最新的船只。",
            },
            {
                "id": "E2",
                "title": "俱乐部会员目击 — C 码头",
                "detail": "一名俱乐部会员注意到，嫌疑人沿 C 码头向远端走去，边走边观赏两侧的船只。",
            },
            {
                "id": "E3",
                "title": "被盗游轮上的指纹",
                "detail": "警方在被盗望远镜所在的邮轮上提取到了嫌疑人的指纹",
            },
        ],
    },
}

GENERAL_TERRORISM_CASE_INFO = {
    "title": "案件 — 液体炸弹恐袭情报案（通用 Avatar 培训）",
    "overview": """警方掌握情报：有人计划将液体炸弹携带上从伦敦希思罗机场飞往美国的多架商业航班，并在飞行途中同步引爆。警方尚不清楚具体涉案人员。

警方在 High Wycombe 附近 King's Wood 一带挖出多个行李袋和一个行李箱，内有常见爆炸物原料（如过氧化氢）及引爆装置相关材料（HMTD）。警方测试显示，仅 500ml 液体爆炸物就足以击碎厚防护玻璃。

目前警方正排查居住在该区域附近、且购买过与埋藏行李箱同款行李箱的人。本案被带来问询的嫌疑人为 Charlie（26 岁，问询日期 2019-11-09）。警方已取得 CCTV：其曾在伦敦 SOHO 的 Luggage Pros 购买与涉案同款行李箱（2019-10-18）。警方尚不能确定其是否参与袭击计划；此前已完成基础背景询问并建立初步关系。""",
    "evidence": [
        {
            "id": "E1",
            "title": "King's Wood 埋藏爆炸物材料",
            "detail": "警方在 High Wycombe 附近 King's Wood 挖出多个行李袋及行李箱，内含过氧化氢等爆炸物原料及 HMTD 相关材料；测试表明约 500ml 液体爆炸物即可击碎厚防护玻璃。",
        },
        {
            "id": "E2",
            "title": "Luggage Pros 购买行李箱 CCTV",
            "detail": "警方取得 CCTV 画面：Charlie 于 2019-10-18 在伦敦 SOHO 的 Luggage Pros 购买与埋藏现场同款行李箱。",
        },
        {
            "id": "E3",
            "title": "希思罗液体炸弹袭击情报",
            "detail": "情报显示有人计划将液体炸弹带上从希思罗飞往美国的多架航班并在途中同步引爆；调查重点为居住在埋藏区域附近且购买同款行李箱的人员。",
        },
    ],
    "efm_analysis": """建议 EFM 审讯策略（恐袭情报案）：
· 低来源低具体：「我们有情报表明近期可能有人策划针对民航的袭击。」
· 低来源高具体：「我们有信息表明爆炸物材料曾在 High Wycombe 附近 King's Wood 一带被埋藏。」
· 高来源高具体：「我们在 Luggage Pros 的监控中看到你于 10 月 18 日购买了与现场同款行李箱。」""",
}

AVATAR_TRAINING_SETTINGS = [
    {"id": "neutral_guilty", "setting": "中性", "guilt": "guilty", "label": "设置 1：中性 · 有罪"},
    {"id": "neutral_innocent", "setting": "中性", "guilt": "innocent", "label": "设置 2：中性 · 无罪"},
    {"id": "evasive_guilty", "setting": "回避", "guilt": "guilty", "label": "设置 3：回避 · 有罪"},
    {"id": "evasive_innocent", "setting": "回避", "guilt": "innocent", "label": "设置 4：回避 · 无罪"},
    {"id": "hostile_guilty", "setting": "敌意", "guilt": "guilty", "label": "设置 5：敌意 · 有罪"},
    {"id": "hostile_innocent", "setting": "敌意", "guilt": "innocent", "label": "设置 6：敌意 · 无罪"},
]

AVATAR_SETTING_PROMPTS = {
    "中性": "你在审讯中保持平静和合作的态度。你以平和的语调回答问题，不回避也不对抗。你表现得像一个普通人，愿意配合调查，但同时也坚持自己的立场。你不会主动提供额外信息，但被问到的每个问题都会简洁回答。",
    "回避": "你在审讯中表现得闪烁其词和回避。你倾向于用模糊的答案回复问题，经常说'我不记得了'或'我不确定'。你会巧妙地转移话题，避免直接回答尖锐问题。你看起来有些紧张和不自在，但你不会表现出攻击性。",
    "敌意": "你在审讯中表现出明显的不合作和敌对态度。你对审讯者的提问感到不耐烦，会用讽刺或愤怒的语气回应。你认为审讯是在浪费时间，对你的权利被侵犯感到不满。你会挑战审讯者的权威，并可能拒绝回答某些问题。",
}

GENERAL_AVATAR_CASE_CONTEXT_ZH = """警方掌握情报：有人计划将液体炸弹携带上从伦敦希思罗机场飞往美国的多架商业航班，并在飞行途中同步引爆。警方尚不清楚具体涉案人员。

警方在 High Wycombe 附近 King's Wood 一带挖出多个行李袋和一个行李箱，内有常见爆炸物原料（如过氧化氢）及引爆装置相关材料（HMTD）。警方测试显示，仅 500ml 液体爆炸物就足以击碎厚防护玻璃。

目前警方正排查居住在该区域附近、且购买过与埋藏行李箱同款行李箱的人。你是本案被带来问询的嫌疑人（26岁，问询日期 2019-11-09）。

警方已取得 CCTV 画面：你曾在伦敦 SOHO 的 Luggage Pros 购买与涉案同款行李箱（拍摄日期 2019-10-18）。警方尚不能确定你是否参与袭击计划。此前警方已完成基础背景询问并建立初步关系。"""

GENERAL_AVATAR_IDENTITY_ZH = """基础信息：
- 姓名：Charlie
- 年龄：26
- 宗教：无
- 居住：King's Wood, London
- 婚姻：单身，独居
- 家庭：父母在摩洛哥，联系有限，不愿多谈
- 工作：曾在 King's Wood 附近 Aldi 工作，因工时不足离职
- 兴趣：FIFA 游戏、足球节目
- 支持球队：Queens Park Rangers"""

DEEPSEEK_FEEDBACK_PROMPT_PLACEHOLDER = """你是一名审讯培训导师。请根据以下转录记录，用中文为审讯者提供反馈（约 500–800 字）。

【必须包含】
1. 在反馈开头明确写出：该 Avatar 的真实罪责状态为「{avatar_guilt_label}」（这是标准答案，供学员对照）。
2. 仅根据转录中的对话内容，分析嫌疑人哪些言行可能反映其有罪、哪些可能反映其无罪。
   - 禁止分析眼神、表情、肢体语言、语气等非文字转录内容。
   - 有罪嫌疑人常见表现：遮掩信息、回避关键问题、前后矛盾、过度辩解等。
   - 无罪嫌疑人常见表现：相对坦诚、愿意澄清误会、对未涉及事项回答一致等。
3. 评价审讯者的提问策略与 SUE/EFM 运用，并将审讯者的最终判断「{interviewer_judgment}」与真实状态对比。
4. 给出 2–3 条可操作的改进建议。

审讯记录：
{transcript}

Avatar 人格设定：{avatar_setting_label}

请用中文回复。"""

def _load_avatar_feedback_system_prompt():
    """Load EFM feedback tutor prompt from materials/prompts/avatar_feedback.md (or legacy docx)."""
    if os.path.isfile(FEEDBACK_PROMPT_MD):
        try:
            with open(FEEDBACK_PROMPT_MD, "r", encoding="utf-8") as f:
                text = f.read().strip()
            if text:
                return text
        except Exception:
            pass
    if os.path.isfile(LEGACY_FEEDBACK_DOCX):
        try:
            from docx import Document
            doc = Document(LEGACY_FEEDBACK_DOCX)
            paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
            text = "\n".join(paragraphs)
            return text if text.strip() else DEEPSEEK_FEEDBACK_PROMPT_PLACEHOLDER
        except Exception:
            pass
    return DEEPSEEK_FEEDBACK_PROMPT_PLACEHOLDER


AVATAR_FEEDBACK_SYSTEM_PROMPT = _load_avatar_feedback_system_prompt()


def avatar_setting_label_public(full_label: str, training_type: str) -> str:
    """Participant-facing label: Avatar 训练组不向参与者展示有罪/无罪设定。"""
    if training_type in ("avatar_specific", "avatar_general") and " · " in (full_label or ""):
        return full_label.split(" · ", 1)[0].strip()
    return full_label or ""


PROFILE_QUESTIONS = [
    {"id": "q1", "section": "0", "label": "年龄", "type": "number", "min": 18, "max": 70, "placeholder": "例如: 32"},
    {"id": "q2", "section": "0", "label": "性别", "type": "radio", "options": ["男 Male", "女 Female", "其他 Other"]},
    {"id": "q3", "section": "0", "label": "户口类型", "type": "radio", "options": ["城镇户口 Urban", "农村户口 Rural"]},
    {"id": "q4", "section": "1", "label": "目前职业/具体职务", "type": "text", "placeholder": "例如：自由职业者、公司职员、快递员等"},
    {"id": "q5", "section": "1", "label": "平均月收入（元）", "type": "radio", "options": ["3000以下", "3000-5000", "5000-10000", "10000-20000", "20000以上"]},
    {"id": "q6", "section": "1", "label": "每月固定支出占收入的比例", "type": "radio", "options": ["30%以下", "30%-60%", "60%-90%", "入不敷出"]},
    {"id": "q7", "section": "1", "label": "投资习惯（可多选）", "type": "checkbox", "options": ["不投资/只存定期", "低风险理财", "股票/基金", "加密货币/高风险投资"]},
    {"id": "q8", "section": "1", "label": "是否购买商业保险（如重疾险等）", "type": "radio", "options": ["是", "否"]},
    {"id": "q9", "section": "1", "label": "过去三年换工作次数", "type": "radio", "options": ["0次", "1-2次", "3次及以上"]},
    {"id": "q10", "section": "1", "label": "是否有犯罪记录", "type": "radio", "options": ["是", "否"]},
    {"id": "q11", "section": "1", "label": "是否有行政处罚记录", "type": "radio", "options": ["是", "否"]},
    {"id": "q12", "section": "2", "label": "家庭结构（可多选）", "type": "checkbox", "options": ["我的父母双全", "我是单亲家庭或双亲已故", "我是独生子女", "我有兄弟姐妹", "我有子女"]},
    {"id": "q13", "section": "2", "label": "日常活跃社交圈大小（每周至少联系一次的非工作关系人数）", "type": "radio", "options": ["0-2人", "3-5人", "6-10人", "10人以上"]},
    {"id": "q14", "section": "2", "label": "婚姻/恋爱状况", "type": "radio", "options": ["单身", "恋爱中", "已婚", "离异/丧偶"]},
    {"id": "q15", "section": "2", "label": "日均私人通话/语音时长", "type": "radio", "options": ["很少（主要文字）", "10分钟以下", "10-30分钟", "30分钟以上"]},
    {"id": "q16", "section": "2", "label": "目前居住情况", "type": "radio", "options": ["独居", "与伴侣/配偶同住", "与父母/亲戚同住", "与室友/朋友合租", "宿舍/集体居住"]},
    {"id": "q17", "section": "2", "label": "身边最亲近的朋友会用哪一个词形容你的性格？", "type": "radio", "options": ["冷静寡言", "脾气急躁", "不愿吃亏", "随和实在"]},
    {"id": "q18", "section": "3", "label": "成瘾史（可多选）", "type": "checkbox", "options": ["经常吸烟", "频繁大量饮酒", "榕榔等", "无"]},
    {"id": "q19", "section": "3", "label": "业余时间常去场所（可多选）", "type": "checkbox", "options": ["宅在家", "酒吧/夜店", "网吧/电竞酒店", "棋牌室/麻将馆", "咖啡馆/书店", "健身房/公园"]},
    {"id": "q20", "section": "3", "label": "身心健康状况（可多选）", "type": "checkbox", "options": ["非常健康", "慢性病需长期服药", "曾有/现有心理困扰（如抑郁/焦虑）"]},
    {"id": "q21", "section": "3", "label": "日常个人物品的颜色偏好", "type": "radio", "options": ["非常鲜艳亮眼", "浅色系", "深色/低调色系", "几乎全黑/极简"]},
    {"id": "q22", "section": "4", "label": "正式访谈时是否会戴眼镜", "type": "radio", "options": ["是", "否"]},
    {"id": "q23", "section": "4", "label": "正式访谈时长发还是短发", "type": "radio", "options": ["长发", "短发"]},
]

SECTION_NAMES = {
    "0": "基本信息 Basic Demographics",
    "1": "社会经济状况 Socioeconomic Status",
    "2": "社交与人际关系 Social Relationships",
    "3": "个人习惯与健康 Personal Habits & Health",
    "4": "外观特征 Appearance",
}

SHEET_PARTICIPANTS = "participants"
SHEET_GROUPS = "groups"
SHEET_PROFILES = "profiles"
SHEET_AVAILABILITIES = "availabilities"
SHEET_APPOINTMENTS = "appointments"
SHEET_TRAINING_SESSIONS = "training_sessions"
SHEET_INTERVIEW_QUESTIONNAIRES = "interview_questionnaires"
SHEET_QUESTIONNAIRE_OVERRIDES = "questionnaire_overrides"
SHEET_SERIOUS_GAME = "serious_game_choices"
SHEET_META = "meta"

SHEET_COLUMNS = {
    SHEET_PARTICIPANTS: ["id", "phone", "role", "group_name", "full_id", "guilt", "case_type", "training_type", "attention_passed", "attention_failed", "game_completed", "profile_completed", "completed", "created_at", "avatar_practice_transcript"],
    SHEET_GROUPS: ["name", "suspect_id", "interviewer_id", "created_at"],
    SHEET_PROFILES: ["participant_id", "data", "submitted_at"],
    SHEET_AVAILABILITIES: ["id", "phone", "group_name", "role", "slots", "updated_at"],
    SHEET_APPOINTMENTS: ["id", "phone", "time_slot", "role", "status", "booked_at"],
    SHEET_TRAINING_SESSIONS: ["id", "interviewer_id", "phone", "session_num", "avatar_setting", "avatar_guilt", "judgment", "transcript", "feedback", "started_at", "completed_at"],
    SHEET_INTERVIEW_QUESTIONNAIRES: ["id", "phone", "role", "phase", "appointment_slot", "answers_json", "submitted_at"],
    SHEET_QUESTIONNAIRE_OVERRIDES: ["id", "phone", "phase", "is_open", "updated_at"],
    SHEET_SERIOUS_GAME: [
        "timestamp", "phone", "participant_id", "case", "condition",
        "step_index", "video", "choice",
    ],
    SHEET_META: ["key", "value"],
}

PHASE_PRE = "pre"
PHASE_POST = "post"

BOOKING_DAYS_AHEAD = 5
BOOKING_SLOT_WINDOWS = [
    ("09:30", "11:00"),
    ("14:00", "17:00"),
    ("19:30", "21:00"),
]
BOOKING_SLOT_STEP_MINUTES = 30
TRAINING_TYPES = ["theory_sue", "avatar_specific", "avatar_general", "control"]
TRAINING_GROUP_LABELS = {
    "control": "A",
    "theory_sue": "B",
    "avatar_general": "C",
    "avatar_specific": "D",
}
TRAINING_TARGET_PER_TYPE = 28


def training_group_label(training_type):
    return TRAINING_GROUP_LABELS.get(training_type or "", "")
META_BLACKLIST_PHONES = "blacklisted_phones"
META_DISABLED_SLOTS = "appointment_slot_disabled"

# Based on IRB Questionnaire_formal interview-Interviewer.pdf / Suspect.pdf
_DEMO_PRE = [
    {"id": "demo_age", "section": "基本信息", "label": "您的年龄是？", "type": "number", "min": 16, "max": 99},
    {
        "id": "demo_gender", "section": "基本信息", "label": "您的性别是？",
        "type": "radio",
        "options": ["男性", "女性", "其他/不愿透露"],
    },
    {
        "id": "demo_prior_training", "section": "基本信息",
        "label": "在参与本实验之前，您是否接受过讯问或审讯技巧方面的培训？",
        "type": "radio", "options": ["是", "否"],
    },
]
_PRE_INTERVIEW_SCALES = [
    {
        "id": "pre_stress", "section": "访谈前",
        "label": "针对即将到来的审讯任务，您预计会有多紧张？",
        "type": "scale", "min": 1, "max": 7,
        "min_label": "一点也不", "max_label": "非常紧张",
    },
    {
        "id": "pre_demanding", "section": "访谈前",
        "label": "您认为即将到来的审讯任务会有多困难？",
        "type": "scale", "min": 1, "max": 7,
        "min_label": "一点也不", "max_label": "非常困难",
    },
    {
        "id": "pre_cope", "section": "访谈前",
        "label": "您认为自己在应对即将到来的审讯任务方面有多大把握？",
        "type": "scale", "min": 1, "max": 7,
        "min_label": "一点也没有", "max_label": "非常有把握",
    },
    {
        "id": "pre_anger_watch", "section": "访谈前",
        "label": "惹我生气的人最好小心点",
        "type": "scale", "min": 1, "max": 5,
        "min_label": "几乎从不", "max_label": "几乎总是",
    },
    {
        "id": "pre_get_even", "section": "访谈前",
        "label": "如果有人试图伤害我，我一定会报复",
        "type": "scale", "min": 1, "max": 5,
        "min_label": "几乎从不", "max_label": "几乎总是",
    },
    {
        "id": "pre_lose_temper", "section": "访谈前",
        "label": "我生气时会发脾气，把怒气发泄到别人身上",
        "type": "scale", "min": 1, "max": 5,
        "min_label": "几乎从不", "max_label": "几乎总是",
    },
]

INTERVIEW_QUESTION_BANK = {
    PHASE_PRE: {
        "S": _DEMO_PRE + _PRE_INTERVIEW_SCALES,
        "I": _DEMO_PRE + _PRE_INTERVIEW_SCALES,
    },
    PHASE_POST: {
        "S": [
            {
                "id": "post_s_classify", "section": "访谈后",
                "label": "您认为审讯员会将您判定为？",
                "type": "radio", "options": ["有罪", "无罪"],
            },
            {
                "id": "post_s_confidence", "section": "访谈后",
                "label": "您对自己成功说服审讯员相信您无罪有多大把握？（0-100）",
                "type": "number", "min": 0, "max": 100,
            },
            {
                "id": "post_s_aggressive_self", "section": "访谈后",
                "label": "您认为自己在审讯过程中对审讯员表现出多大程度的攻击性？",
                "type": "scale", "min": 1, "max": 10,
                "min_label": "完全没有", "max_label": "非常程度",
            },
            {
                "id": "post_s_coercive_self", "section": "访谈后",
                "label": "您认为自己在审讯过程中对审讯员表现出多大程度的胁迫性？",
                "type": "scale", "min": 1, "max": 10,
                "min_label": "完全没有", "max_label": "非常程度",
            },
            {
                "id": "post_s_aggressive_interviewer", "section": "访谈后",
                "label": "您认为审讯员在审讯过程中对您表现出多大程度的攻击性？",
                "type": "scale", "min": 1, "max": 10,
                "min_label": "完全没有", "max_label": "非常程度",
            },
            {
                "id": "post_s_coercive_interviewer", "section": "访谈后",
                "label": "您认为审讯员在审讯过程中对您表现出多大程度的胁迫性？",
                "type": "scale", "min": 1, "max": 10,
                "min_label": "完全没有", "max_label": "非常程度",
            },
            {
                "id": "post_s_feedback", "section": "总体反馈",
                "label": "您对本实验或任何其他方面还有什么意见或建议？",
                "type": "text", "optional": True,
            },
        ],
        "I": [
            {
                "id": "post_i_guilt", "section": "访谈后",
                "label": "您认为您审讯的嫌疑人对所涉嫌犯罪是有罪还是无罪？",
                "type": "radio", "options": ["有罪", "无罪"],
            },
            {
                "id": "post_i_certainty", "section": "访谈后",
                "label": "您对自己结论的把握程度？（0-100）",
                "type": "number", "min": 0, "max": 100,
            },
            {
                "id": "post_i_factor", "section": "访谈后",
                "label": "对您判断影响最大的因素是？",
                "type": "radio",
                "options": ["言语线索", "非言语线索", "两者都有"],
            },
            {
                "id": "post_i_verbal_cues", "section": "访谈后",
                "label": "哪些言语线索影响了您的判断？",
                "type": "radio",
                "options": [
                    "嫌疑人说了与证据不一致的话（或与证据一致）",
                    "嫌疑人说了与之前陈述不一致的话（或与之前陈述一致）",
                    "嫌疑人说了不可信的话（或可信的话）",
                    "嫌疑人说了模糊、不连贯的话（或清晰、连贯的话）",
                    "其他",
                ],
            },
            {"id": "post_i_verbal_other", "section": "访谈后", "label": "言语线索（其他，请说明）", "type": "text", "optional": True},
            {
                "id": "post_i_nonverbal_cues", "section": "访谈后",
                "label": "哪些非言语线索影响了您的判断？",
                "type": "radio",
                "options": [
                    "嫌疑人频繁转移视线（或未频繁转移视线）",
                    "嫌疑人肢体姿态不自然、紧张抖动（或自然、无紧张抖动）",
                    "嫌疑人面部表情紧张（或表情放松舒适）",
                    "嫌疑人频繁改变语调（或语调相对平稳）",
                    "其他",
                ],
            },
            {"id": "post_i_nonverbal_other", "section": "访谈后", "label": "非言语线索（其他，请说明）", "type": "text", "optional": True},
            {
                "id": "post_i_aggressive_self", "section": "访谈后",
                "label": "您认为自己在审讯过程中对嫌疑人表现出多大程度的攻击性？",
                "type": "scale", "min": 1, "max": 10,
                "min_label": "完全没有", "max_label": "非常程度",
            },
            {
                "id": "post_i_coercive_self", "section": "访谈后",
                "label": "您认为自己在审讯过程中对嫌疑人表现出多大程度的胁迫性？",
                "type": "scale", "min": 1, "max": 10,
                "min_label": "完全没有", "max_label": "非常程度",
            },
            {
                "id": "post_i_aggressive_suspect", "section": "访谈后",
                "label": "您认为嫌疑人在审讯过程中对您表现出多大程度的攻击性？",
                "type": "scale", "min": 1, "max": 10,
                "min_label": "完全没有", "max_label": "非常程度",
            },
            {
                "id": "post_i_coercive_suspect", "section": "访谈后",
                "label": "您认为嫌疑人在审讯过程中对您表现出多大程度的胁迫性？",
                "type": "scale", "min": 1, "max": 10,
                "min_label": "完全没有", "max_label": "非常程度",
            },
            {
                "id": "post_i_feedback", "section": "总体反馈",
                "label": "您对本实验或任何其他方面还有什么意见或建议？",
                "type": "text", "optional": True,
            },
        ],
    },
}


# ====== Excel Storage Layer ======

def ensure_excel_file():
    """Ensure EXCEL_FILE is a writable file (not a Docker-created directory)."""
    if os.path.isdir(EXCEL_FILE):
        raise RuntimeError(
            f"{EXCEL_FILE} is a directory (Docker mount error). "
            f"On the server run: docker compose down && rm -rf experiment_data.xlsx && mkdir -p data"
        )
    if os.path.isfile(EXCEL_FILE):
        return
    if os.path.isfile(LEGACY_EXCEL_FILE) and not os.path.isdir(LEGACY_EXCEL_FILE):
        shutil.copy2(LEGACY_EXCEL_FILE, EXCEL_FILE)
        logger.info("Copied legacy Excel from %s to %s", LEGACY_EXCEL_FILE, EXCEL_FILE)


def init_excel():
    """Create the Excel workbook with all required sheets if it doesn't exist."""
    ensure_excel_file()
    if os.path.isfile(EXCEL_FILE):
        ensure_sheets()
        return
    wb = Workbook()
    wb.remove(wb.active)
    for name, columns in SHEET_COLUMNS.items():
        ws = wb.create_sheet(name)
        ws.append(columns)
    # Init meta values
    ws = wb[SHEET_META]
    ws.append(["next_participant_id", 1])
    ws.append(["next_avail_id", 1])
    ws.append(["next_appt_id", 1])
    ws.append(["next_questionnaire_id", 1])
    ws.append(["next_qoverride_id", 1])
    wb.save(EXCEL_FILE)
    wb.close()


def ensure_sheets():
    """Add any missing sheets to an existing Excel file (migration)."""
    wb = load_workbook(EXCEL_FILE)
    try:
        existing_sheets = wb.sheetnames
        for name, columns in SHEET_COLUMNS.items():
            if name not in existing_sheets:
                ws = wb.create_sheet(name)
                ws.append(columns)
                print(f"  [MIGRATION] Added sheet: {name}")
            else:
                # Migrate: add missing columns to existing sheets
                ws = wb[name]
                existing_cols = [c.value for c in ws[1]]
                new_cols = [c for c in columns if c not in existing_cols]
                if new_cols:
                    # Add missing column headers
                    for col_name in new_cols:
                        ws.cell(row=1, column=len(existing_cols) + 1 + new_cols.index(col_name), value=col_name)
                    print(f"  [MIGRATION] Added columns to {name}: {', '.join(new_cols)}")
        wb.save(EXCEL_FILE)
    finally:
        wb.close()


class ExcelStore:
    """Thread-safe Excel data access layer."""

    def _load(self):
        return load_workbook(EXCEL_FILE)

    def _save(self, wb):
        wb.save(EXCEL_FILE)

    def _close(self, wb):
        wb.close()

    def _read_all(self, wb, sheet_name):
        ws = wb[sheet_name]
        headers = [c.value for c in ws[1]]
        rows = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if any(v is not None for v in row):
                rows.append(dict(zip(headers, row)))
        return rows

    def _write_all(self, wb, sheet_name, rows):
        ws = wb[sheet_name]
        ws.delete_rows(2, ws.max_row - 1)
        headers = [c.value for c in ws[1]]
        for row_data in rows:
            ws.append([row_data.get(h) for h in headers])

    def _next_id(self, wb, key="next_participant_id"):
        meta = self._read_all(wb, SHEET_META)
        for m in meta:
            if m["key"] == key:
                val = m["value"]
                m["value"] = val + 1
                self._write_all(wb, SHEET_META, meta)
                return val
        meta.append({"key": key, "value": 2})
        self._write_all(wb, SHEET_META, meta)
        return 1

    def _set_id_counter(self, wb, key, value):
        meta = self._read_all(wb, SHEET_META)
        for m in meta:
            if m["key"] == key:
                m["value"] = value
                self._write_all(wb, SHEET_META, meta)
                return

    def get_meta(self, key, default=None):
        with _excel_lock:
            wb = self._load()
            try:
                for m in self._read_all(wb, SHEET_META):
                    if m.get("key") == key:
                        return m.get("value", default)
                return default
            finally:
                self._close(wb)

    def set_meta(self, key, value):
        with _excel_lock:
            wb = self._load()
            try:
                meta = self._read_all(wb, SHEET_META)
                for m in meta:
                    if m.get("key") == key:
                        m["value"] = value
                        self._write_all(wb, SHEET_META, meta)
                        self._save(wb)
                        return
                meta.append({"key": key, "value": value})
                self._write_all(wb, SHEET_META, meta)
                self._save(wb)
            finally:
                self._close(wb)

    def is_phone_blacklisted(self, phone):
        raw = self.get_meta(META_BLACKLIST_PHONES, "[]")
        try:
            phones = json.loads(raw) if isinstance(raw, str) else (raw or [])
        except Exception:
            phones = []
        return phone in phones

    def blacklist_phone(self, phone, reason="attention_failed"):
        raw = self.get_meta(META_BLACKLIST_PHONES, "[]")
        try:
            phones = json.loads(raw) if isinstance(raw, str) else (raw or [])
        except Exception:
            phones = []
        if phone not in phones:
            phones.append(phone)
            self.set_meta(META_BLACKLIST_PHONES, json.dumps(phones, ensure_ascii=False))
        p = self.get_participant(phone)
        if p:
            self.update_participant(phone, attention_failed=1)

    def get_disabled_slots(self):
        raw = self.get_meta(META_DISABLED_SLOTS, "[]")
        try:
            slots = json.loads(raw) if isinstance(raw, str) else (raw or [])
        except Exception:
            slots = []
        return set(slots)

    def set_slot_enabled(self, slot_str, enabled):
        disabled = self.get_disabled_slots()
        if enabled:
            disabled.discard(slot_str)
        else:
            disabled.add(slot_str)
        self.set_meta(META_DISABLED_SLOTS, json.dumps(sorted(disabled), ensure_ascii=False))

    # ---- Participants ----

    def get_participant(self, phone):
        with _excel_lock:
            wb = self._load()
            try:
                for p in self._read_all(wb, SHEET_PARTICIPANTS):
                    if p["phone"] == phone:
                        return p
                return None
            finally:
                self._close(wb)

    def get_participant_by_id(self, pid):
        with _excel_lock:
            wb = self._load()
            try:
                for p in self._read_all(wb, SHEET_PARTICIPANTS):
                    if p["id"] == pid:
                        return p
                return None
            finally:
                self._close(wb)

    def get_all_participants(self):
        with _excel_lock:
            wb = self._load()
            try:
                return self._read_all(wb, SHEET_PARTICIPANTS)
            finally:
                self._close(wb)

    def add_participant(self, **kwargs):
        with _excel_lock:
            wb = self._load()
            try:
                participants = self._read_all(wb, SHEET_PARTICIPANTS)
                pid = self._next_id(wb, "next_participant_id")
                kwargs["id"] = pid
                kwargs.setdefault("attention_passed", 0)
                kwargs.setdefault("full_id", "")
                kwargs.setdefault("game_completed", 0)
                kwargs.setdefault("profile_completed", 0)
                kwargs.setdefault("completed", 0)
                kwargs.setdefault("created_at", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                participants.append(kwargs)
                self._write_all(wb, SHEET_PARTICIPANTS, participants)
                self._save(wb)
                return pid
            finally:
                self._close(wb)

    def update_participant(self, phone, **kwargs):
        with _excel_lock:
            wb = self._load()
            try:
                participants = self._read_all(wb, SHEET_PARTICIPANTS)
                for p in participants:
                    if p["phone"] == phone:
                        p.update(kwargs)
                        self._write_all(wb, SHEET_PARTICIPANTS, participants)
                        self._save(wb)
                        return True
                return False
            finally:
                self._close(wb)

    def delete_participant(self, pid):
        with _excel_lock:
            wb = self._load()
            try:
                # Get participant info for cleanup
                phone = None
                participants = self._read_all(wb, SHEET_PARTICIPANTS)
                for p in participants:
                    if p["id"] == pid:
                        phone = p.get("phone", "")
                # Remove participant
                participants = [p for p in participants if p["id"] != pid]
                self._write_all(wb, SHEET_PARTICIPANTS, participants)
                # Remove related profile
                profiles = [p for p in self._read_all(wb, SHEET_PROFILES) if p["participant_id"] != pid]
                self._write_all(wb, SHEET_PROFILES, profiles)
                # Remove availabilities
                if phone:
                    avails = [a for a in self._read_all(wb, SHEET_AVAILABILITIES) if a.get("phone") != phone]
                    self._write_all(wb, SHEET_AVAILABILITIES, avails)
                    # Remove appointments
                    appts = [a for a in self._read_all(wb, SHEET_APPOINTMENTS) if a.get("phone") != phone]
                    self._write_all(wb, SHEET_APPOINTMENTS, appts)
                    # Remove interview questionnaires
                    qrows = [q for q in self._read_all(wb, SHEET_INTERVIEW_QUESTIONNAIRES) if q.get("phone") != phone]
                    self._write_all(wb, SHEET_INTERVIEW_QUESTIONNAIRES, qrows)
                    # Remove questionnaire overrides
                    qover = [q for q in self._read_all(wb, SHEET_QUESTIONNAIRE_OVERRIDES) if q.get("phone") != phone]
                    self._write_all(wb, SHEET_QUESTIONNAIRE_OVERRIDES, qover)
                    # Remove training sessions
                    sessions = [s for s in self._read_all(wb, SHEET_TRAINING_SESSIONS) if s.get("phone") != phone]
                    self._write_all(wb, SHEET_TRAINING_SESSIONS, sessions)
                self._save(wb)
            finally:
                self._close(wb)

    # ---- Groups ----

    def get_waiting_group(self):
        """Find a group that has a suspect but no interviewer assigned yet."""
        with _excel_lock:
            wb = self._load()
            try:
                groups = self._read_all(wb, SHEET_GROUPS)
                for g in sorted(groups, key=lambda x: x.get("created_at", "")):
                    if g.get("suspect_id") and not g.get("interviewer_id"):
                        return g
                return None
            finally:
                self._close(wb)

    def get_group_by_name(self, name):
        with _excel_lock:
            wb = self._load()
            try:
                for g in self._read_all(wb, SHEET_GROUPS):
                    if g["name"] == name:
                        return g
                return None
            finally:
                self._close(wb)

    def get_group_by_interviewer(self, interviewer_id):
        with _excel_lock:
            wb = self._load()
            try:
                for g in self._read_all(wb, SHEET_GROUPS):
                    if g.get("interviewer_id") == interviewer_id:
                        return g
                return None
            finally:
                self._close(wb)

    def add_group(self, name, suspect_id, interviewer_id=None):
        with _excel_lock:
            wb = self._load()
            try:
                groups = self._read_all(wb, SHEET_GROUPS)
                groups.append({
                    "name": name,
                    "suspect_id": suspect_id,
                    "interviewer_id": interviewer_id,
                    "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                })
                self._write_all(wb, SHEET_GROUPS, groups)
                self._save(wb)
            finally:
                self._close(wb)

    def update_group(self, name, **kwargs):
        with _excel_lock:
            wb = self._load()
            try:
                groups = self._read_all(wb, SHEET_GROUPS)
                for g in groups:
                    if g["name"] == name:
                        g.update(kwargs)
                        self._write_all(wb, SHEET_GROUPS, groups)
                        self._save(wb)
                        return True
                return False
            finally:
                self._close(wb)

    def count_groups(self):
        with _excel_lock:
            wb = self._load()
            try:
                return len(self._read_all(wb, SHEET_GROUPS))
            finally:
                self._close(wb)

    # ---- Profiles ----

    def get_profile(self, participant_id):
        with _excel_lock:
            wb = self._load()
            try:
                for p in self._read_all(wb, SHEET_PROFILES):
                    if p["participant_id"] == participant_id:
                        return p
                return None
            finally:
                self._close(wb)

    def upsert_profile(self, participant_id, data_json):
        with _excel_lock:
            wb = self._load()
            try:
                profiles = self._read_all(wb, SHEET_PROFILES)
                for p in profiles:
                    if p["participant_id"] == participant_id:
                        p["data"] = data_json
                        p["submitted_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        self._write_all(wb, SHEET_PROFILES, profiles)
                        self._save(wb)
                        return
                profiles.append({
                    "participant_id": participant_id,
                    "data": data_json,
                    "submitted_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                })
                self._write_all(wb, SHEET_PROFILES, profiles)
                self._save(wb)
            finally:
                self._close(wb)

    # ---- Availabilities ----

    def get_availabilities(self, group_name=None):
        with _excel_lock:
            wb = self._load()
            try:
                all_avail = self._read_all(wb, SHEET_AVAILABILITIES)
                if group_name:
                    return [a for a in all_avail if a.get("group_name") == group_name]
                return all_avail
            finally:
                self._close(wb)

    def upsert_availability(self, phone, group_name, role, slots):
        with _excel_lock:
            wb = self._load()
            try:
                availabilities = self._read_all(wb, SHEET_AVAILABILITIES)
                slots_json = json.dumps(slots, ensure_ascii=False)
                now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                for a in availabilities:
                    if a["phone"] == phone:
                        a["group_name"] = group_name
                        a["role"] = role
                        a["slots"] = slots_json
                        a["updated_at"] = now
                        self._write_all(wb, SHEET_AVAILABILITIES, availabilities)
                        self._save(wb)
                        return
                aid = self._next_id(wb, "next_avail_id")
                availabilities.append({
                    "id": aid,
                    "phone": phone,
                    "group_name": group_name,
                    "role": role,
                    "slots": slots_json,
                    "updated_at": now,
                })
                self._write_all(wb, SHEET_AVAILABILITIES, availabilities)
                self._save(wb)
            finally:
                self._close(wb)

    # ---- Appointments ----

    def get_appointments(self, phone=None):
        with _excel_lock:
            wb = self._load()
            try:
                all_appts = self._read_all(wb, SHEET_APPOINTMENTS)
                if phone:
                    return [a for a in all_appts if a["phone"] == phone]
                return all_appts
            finally:
                self._close(wb)

    def add_appointment(self, phone, role, time_slot):
        with _excel_lock:
            wb = self._load()
            try:
                appointments = self._read_all(wb, SHEET_APPOINTMENTS)
                aid = self._next_id(wb, "next_appt_id")
                appointments.append({
                    "id": aid,
                    "phone": phone,
                    "time_slot": time_slot,
                    "role": role,
                    "status": "confirmed",
                    "booked_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                })
                self._write_all(wb, SHEET_APPOINTMENTS, appointments)
                self._save(wb)
                return aid
            finally:
                self._close(wb)

    def update_appointment(self, phone, **kwargs):
        with _excel_lock:
            wb = self._load()
            try:
                appointments = self._read_all(wb, SHEET_APPOINTMENTS)
                for a in appointments:
                    if a["phone"] == phone and a.get("status") == "confirmed":
                        a.update(kwargs)
                        self._write_all(wb, SHEET_APPOINTMENTS, appointments)
                        self._save(wb)
                        return True
                return False
            finally:
                self._close(wb)

    def cancel_appointment(self, phone):
        return self.update_appointment(phone, status="cancelled")

    def delete_appointment_by_id(self, aid):
        with _excel_lock:
            wb = self._load()
            try:
                appointments = self._read_all(wb, SHEET_APPOINTMENTS)
                appointments = [a for a in appointments if a["id"] != aid]
                self._write_all(wb, SHEET_APPOINTMENTS, appointments)
                self._save(wb)
            finally:
                self._close(wb)

    def get_slot_bookings(self):
        """Return dict mapping time_slot -> set of roles booked."""
        with _excel_lock:
            wb = self._load()
            try:
                slot_map = {}
                for a in self._read_all(wb, SHEET_APPOINTMENTS):
                    if a.get("status") == "confirmed":
                        slot = a["time_slot"]
                        role = a.get("role", "")
                        slot_map.setdefault(slot, set()).add(role)
                return slot_map
            finally:
                self._close(wb)

    def get_confirmed_slot_set(self):
        """Return set of fully-booked time slots (both roles taken)."""
        slot_map = self.get_slot_bookings()
        return {s for s, roles in slot_map.items() if "S" in roles and "I" in roles}

    def has_booking(self, phone):
        """Check if a participant already has a confirmed booking."""
        with _excel_lock:
            wb = self._load()
            try:
                for a in self._read_all(wb, SHEET_APPOINTMENTS):
                    if a.get("phone") == phone and a.get("status") == "confirmed":
                        return True
                return False
            finally:
                self._close(wb)

    def get_my_booking(self, phone):
        """Get a participant's confirmed booking."""
        with _excel_lock:
            wb = self._load()
            try:
                for a in self._read_all(wb, SHEET_APPOINTMENTS):
                    if a.get("phone") == phone and a.get("status") == "confirmed":
                        return a
                return None
            finally:
                self._close(wb)

    # ---- Interview Questionnaires ----

    def get_interview_questionnaires(self, phone=None, phase=None):
        with _excel_lock:
            wb = self._load()
            try:
                rows = self._read_all(wb, SHEET_INTERVIEW_QUESTIONNAIRES)
                if phone:
                    rows = [r for r in rows if r.get("phone") == phone]
                if phase:
                    rows = [r for r in rows if r.get("phase") == phase]
                return rows
            finally:
                self._close(wb)

    def upsert_interview_questionnaire(self, phone, role, phase, appointment_slot, answers_json):
        with _excel_lock:
            wb = self._load()
            try:
                rows = self._read_all(wb, SHEET_INTERVIEW_QUESTIONNAIRES)
                now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                for r in rows:
                    if r.get("phone") == phone and r.get("phase") == phase:
                        r["role"] = role
                        r["appointment_slot"] = appointment_slot
                        r["answers_json"] = answers_json
                        r["submitted_at"] = now_str
                        self._write_all(wb, SHEET_INTERVIEW_QUESTIONNAIRES, rows)
                        self._save(wb)
                        return r.get("id")
                qid = self._next_id(wb, "next_questionnaire_id")
                rows.append({
                    "id": qid,
                    "phone": phone,
                    "role": role,
                    "phase": phase,
                    "appointment_slot": appointment_slot,
                    "answers_json": answers_json,
                    "submitted_at": now_str,
                })
                self._write_all(wb, SHEET_INTERVIEW_QUESTIONNAIRES, rows)
                self._save(wb)
                return qid
            finally:
                self._close(wb)

    def get_questionnaire_override(self, phone, phase):
        with _excel_lock:
            wb = self._load()
            try:
                for r in self._read_all(wb, SHEET_QUESTIONNAIRE_OVERRIDES):
                    if r.get("phone") == phone and r.get("phase") == phase:
                        return r
                return None
            finally:
                self._close(wb)

    def get_all_questionnaire_overrides(self):
        with _excel_lock:
            wb = self._load()
            try:
                return self._read_all(wb, SHEET_QUESTIONNAIRE_OVERRIDES)
            finally:
                self._close(wb)

    def set_questionnaire_override(self, phone, phase, is_open):
        with _excel_lock:
            wb = self._load()
            try:
                rows = self._read_all(wb, SHEET_QUESTIONNAIRE_OVERRIDES)
                now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                for r in rows:
                    if r.get("phone") == phone and r.get("phase") == phase:
                        r["is_open"] = 1 if is_open else 0
                        r["updated_at"] = now_str
                        self._write_all(wb, SHEET_QUESTIONNAIRE_OVERRIDES, rows)
                        self._save(wb)
                        return r.get("id")
                oid = self._next_id(wb, "next_qoverride_id")
                rows.append({
                    "id": oid,
                    "phone": phone,
                    "phase": phase,
                    "is_open": 1 if is_open else 0,
                    "updated_at": now_str,
                })
                self._write_all(wb, SHEET_QUESTIONNAIRE_OVERRIDES, rows)
                self._save(wb)
                return oid
            finally:
                self._close(wb)

    # ---- Training Sessions ----

    def get_training_sessions(self, phone):
        """Get all training sessions for a participant."""
        with _excel_lock:
            wb = self._load()
            try:
                return [s for s in self._read_all(wb, SHEET_TRAINING_SESSIONS)
                        if s.get("phone") == phone]
            finally:
                self._close(wb)

    def start_training_session(self, phone, interviewer_id, session_num, avatar_setting, avatar_guilt):
        """Start a new training session."""
        with _excel_lock:
            wb = self._load()
            try:
                sessions = self._read_all(wb, SHEET_TRAINING_SESSIONS)
                # Check if this session already exists
                for s in sessions:
                    if s.get("phone") == phone and str(s.get("session_num")) == str(session_num):
                        return s.get("id")
                sid = self._next_id(wb, "next_training_session_id")
                sessions.append({
                    "id": sid,
                    "interviewer_id": interviewer_id,
                    "phone": phone,
                    "session_num": session_num,
                    "avatar_setting": avatar_setting,
                    "avatar_guilt": avatar_guilt,
                    "judgment": "",
                    "transcript": "",
                    "feedback": "",
                    "started_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "completed_at": "",
                })
                self._write_all(wb, SHEET_TRAINING_SESSIONS, sessions)
                self._save(wb)
                return sid
            finally:
                self._close(wb)

    def submit_training_session(self, phone, session_num, judgment, transcript, feedback):
        """Submit a completed training session with judgment and feedback."""
        with _excel_lock:
            wb = self._load()
            try:
                sessions = self._read_all(wb, SHEET_TRAINING_SESSIONS)
                for s in sessions:
                    if s.get("phone") == phone and str(s.get("session_num")) == str(session_num):
                        s["judgment"] = judgment
                        s["transcript"] = transcript
                        s["feedback"] = feedback
                        s["completed_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        self._write_all(wb, SHEET_TRAINING_SESSIONS, sessions)
                        self._save(wb)
                        return True
                return False
            finally:
                self._close(wb)

    def count_completed_training_sessions(self, phone):
        """Count how many training sessions have been completed (have judgment + feedback)."""
        with _excel_lock:
            wb = self._load()
            try:
                count = 0
                for s in self._read_all(wb, SHEET_TRAINING_SESSIONS):
                    if s.get("phone") == phone and s.get("judgment") and s.get("feedback"):
                        count += 1
                return count
            finally:
                self._close(wb)

    def get_training_session(self, phone, session_num):
        """Get a specific training session."""
        with _excel_lock:
            wb = self._load()
            try:
                for s in self._read_all(wb, SHEET_TRAINING_SESSIONS):
                    if s.get("phone") == phone and str(s.get("session_num")) == str(session_num):
                        return s
                return None
            finally:
                self._close(wb)

    def update_training_transcript(self, phone, session_num, transcript):
        """Persist in-progress interview transcript (JSON string) without completing the session."""
        with _excel_lock:
            wb = self._load()
            try:
                sessions = self._read_all(wb, SHEET_TRAINING_SESSIONS)
                for s in sessions:
                    if s.get("phone") == phone and str(s.get("session_num")) == str(session_num):
                        s["transcript"] = transcript
                        self._write_all(wb, SHEET_TRAINING_SESSIONS, sessions)
                        self._save(wb)
                        return True
                return False
            finally:
                self._close(wb)

    # ---- Serious game choices (formerly results_serious_game.xlsx) ----

    def log_serious_game_choice(
        self, phone, participant_id, case, condition, step_index, video, choice,
    ):
        with _excel_lock:
            wb = self._load()
            try:
                rows = self._read_all(wb, SHEET_SERIOUS_GAME)
                rows.append({
                    "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "phone": phone,
                    "participant_id": participant_id,
                    "case": case,
                    "condition": condition,
                    "step_index": step_index,
                    "video": video,
                    "choice": choice,
                })
                self._write_all(wb, SHEET_SERIOUS_GAME, rows)
                self._save(wb)
            finally:
                self._close(wb)


store = ExcelStore()


def migrate_old_data():
    """Migrate data from experiment.db and results.json to Excel if they exist."""
    if not os.path.exists(EXCEL_FILE):
        init_excel()

    db_path = os.path.join(BASE_DIR, "experiment.db")
    results_path = os.path.join(BASE_DIR, "results.json")
    if not os.path.exists(db_path) and not os.path.exists(results_path):
        return

    # Only migrate if Excel is empty
    with _excel_lock:
        wb = store._load()
        try:
            existing = store._read_all(wb, SHEET_PARTICIPANTS)
            if existing:
                store._close(wb)
                return
        except Exception:
            store._close(wb)
            return
        store._close(wb)

    # Migrate from SQLite
    if os.path.exists(db_path):
        try:
            import sqlite3
            conn = sqlite3.connect(db_path)
            conn.row_factory = sqlite3.Row
            rows = conn.execute("SELECT * FROM participants").fetchall()
            for row in rows:
                d = dict(row)
                p = store.get_participant(d["phone"])
                if not p:
                    store.add_participant(**d)

            rows = conn.execute("SELECT * FROM groups_table").fetchall()
            for row in rows:
                d = dict(row)
                if not store.get_group_by_name(d["name"]):
                    store.add_group(d["name"], d.get("suspect_id"), d.get("interviewer_id"))

            rows = conn.execute("SELECT * FROM profiles").fetchall()
            for row in rows:
                d = dict(row)
                if not store.get_profile(d["participant_id"]):
                    store.upsert_profile(d["participant_id"], d["data"])

            conn.close()
            print("  [OK] 已从 experiment.db 迁移数据")
        except Exception as e:
            print(f"  [WARN] 从 SQLite 迁移失败: {e}")

    # Migrate from results.json
    if os.path.exists(results_path):
        try:
            with open(results_path, "r", encoding="utf-8") as f:
                data = json.load(f)

            for a in data.get("availabilities", []):
                store.upsert_availability(a["phone"], a.get("group_name", ""), a.get("role", ""), a.get("slots", []))

            for a in data.get("appointments", []):
                with _excel_lock:
                    wb = store._load()
                    try:
                        appts = store._read_all(wb, SHEET_APPOINTMENTS)
                        appts.append({
                            "id": store._next_id(wb, "next_appt_id"),
                            "phone": a["phone"],
                            "time_slot": a["time_slot"],
                            "status": a.get("status", "confirmed"),
                            "booked_at": a.get("booked_at", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
                        })
                        store._write_all(wb, SHEET_APPOINTMENTS, appts)
                        store._save(wb)
                    finally:
                        store._close(wb)

            print("  [OK] 已从 results.json 迁移数据")
        except Exception as e:
            print(f"  [WARN] 从 results.json 迁移失败: {e}")


def migrate_legacy_exports():
    """Import legacy results_serious_game.xlsx and training_feedback/ into experiment_data.xlsx."""
    ensure_sheets()

    if os.path.isfile(LEGACY_SERIOUS_GAME_XLSX):
        try:
            wb_old = load_workbook(LEGACY_SERIOUS_GAME_XLSX, read_only=True)
            ws = wb_old.active
            header_row = next(ws.iter_rows(min_row=1, max_row=1))
            headers = [c.value for c in header_row]
            with _excel_lock:
                wb = store._load()
                try:
                    rows = store._read_all(wb, SHEET_SERIOUS_GAME)
                    existing = {
                        (
                            str(r.get("phone") or ""),
                            str(r.get("step_index") or ""),
                            str(r.get("video") or ""),
                            str(r.get("choice") or ""),
                        )
                        for r in rows
                    }
                    added = 0
                    for row in ws.iter_rows(min_row=2, values_only=True):
                        if not any(row):
                            continue
                        rec = dict(zip(headers, row))
                        phone = rec.get("Phone") or rec.get("phone") or ""
                        step_index = rec.get("StepIndex") or rec.get("step_index") or ""
                        video = rec.get("Video") or rec.get("video") or ""
                        choice = rec.get("Choice") or rec.get("choice") or ""
                        key = (str(phone), str(step_index), str(video), str(choice))
                        if key in existing:
                            continue
                        rows.append({
                            "timestamp": rec.get("Timestamp") or rec.get("timestamp") or "",
                            "phone": phone,
                            "participant_id": rec.get("ParticipantID") or rec.get("participant_id") or "",
                            "case": rec.get("Case") or rec.get("case") or "",
                            "condition": rec.get("Condition") or rec.get("condition") or "",
                            "step_index": step_index,
                            "video": video,
                            "choice": choice,
                        })
                        existing.add(key)
                        added += 1
                    if added:
                        store._write_all(wb, SHEET_SERIOUS_GAME, rows)
                        store._save(wb)
                        print(f"  [OK] 已从 results_serious_game.xlsx 合并 {added} 条到 experiment_data.xlsx")
                finally:
                    store._close(wb)
            wb_old.close()
        except Exception as e:
            print(f"  [WARN] 合并 results_serious_game.xlsx 失败: {e}")

    if not os.path.isdir(LEGACY_TRAINING_FEEDBACK_DIR):
        return

    try:
        imported = 0
        for fname in os.listdir(LEGACY_TRAINING_FEEDBACK_DIR):
            if not fname.endswith("_meta.json"):
                continue
            prefix = fname[: -len("_meta.json")]
            meta_path = os.path.join(LEGACY_TRAINING_FEEDBACK_DIR, fname)
            with open(meta_path, "r", encoding="utf-8") as f:
                meta = json.load(f)
            phone = meta.get("phone")
            session_num = meta.get("session_num")
            if not phone or session_num is None:
                continue
            existing_sessions = store.get_training_sessions(phone)
            if any(
                str(s.get("session_num")) == str(session_num) and s.get("feedback")
                for s in existing_sessions
            ):
                continue
            transcript_path = os.path.join(LEGACY_TRAINING_FEEDBACK_DIR, f"{prefix}_transcript.txt")
            feedback_path = os.path.join(LEGACY_TRAINING_FEEDBACK_DIR, f"{prefix}_feedback.txt")
            transcript = ""
            feedback = ""
            if os.path.isfile(transcript_path):
                with open(transcript_path, "r", encoding="utf-8") as f:
                    transcript = f.read()
            if os.path.isfile(feedback_path):
                with open(feedback_path, "r", encoding="utf-8") as f:
                    feedback = f.read()
            p = store.get_participant(phone)
            interviewer_id = p["id"] if p else ""
            store.start_training_session(
                phone, interviewer_id, session_num,
                meta.get("avatar_setting", ""), meta.get("avatar_guilt", ""),
            )
            store.submit_training_session(
                phone, session_num, meta.get("judgment", ""), transcript, feedback,
            )
            imported += 1
        if imported:
            print(f"  [OK] 已从 training_feedback/ 导入 {imported} 条培训记录到 experiment_data.xlsx")
    except Exception as e:
        print(f"  [WARN] 合并 training_feedback/ 失败: {e}")


# ====== Material & Training File Helpers ======

MATERIAL_SECTION_SPECS = [
    ("consent_suspect", "知情同意书（嫌疑人）", [
        "Consent Form - Suspect (Clean).docx",
    ], [BASE_DIR, MATERIALS_DIR]),
    ("consent_interviewer", "知情同意书（访谈员）", [
        "Consent Form - Interviewer (Clean).docx",
    ], [BASE_DIR, MATERIALS_DIR]),
    ("theory_sue", "B组 SUE 理论培训", [
        "SUE theoretical training group.docx",
        "SUE theoretical training group.pdf",
    ], [MATERIALS_DIR]),
    ("avatar_specific", "D组 特定 Avatar 培训说明", [
        "Specific avatar training group_instructions.docx",
    ], [MATERIALS_DIR]),
    ("avatar_general", "C组 通用 Avatar 培训说明", [
        "General avatar training group_instructions.docx",
    ], [MATERIALS_DIR]),
    ("control", "A组 对照组培训", [
        "Control group.docx",
    ], [MATERIALS_DIR]),
    ("interviewer_bg", "访谈员背景材料", [
        "Interviewer_Bg.docx",
    ], [MATERIALS_DIR]),
    ("avatar_persona_settings", "D组 Avatar 人格设定", [
        "Specific avatar training group_avatar persona settings.docx",
    ], [MATERIALS_DIR]),
    ("avatar_demo", "D组 Avatar 演示说明", [
        "Specific avatar training group_demo.docx",
    ], [MATERIALS_DIR]),
    ("background_info", "案件背景信息", [
        "(CN) Background Information.pdf",
    ], [BASE_DIR, MATERIALS_DIR]),
]

TRAINING_TYPE_SECTION = {
    "theory_sue": "theory_sue",
    "avatar_specific": "avatar_specific",
    "avatar_general": "avatar_general",
    "control": "control",
}


def _extract_text_from_file(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        from docx import Document
        doc = Document(path)
        return "\n\n".join(p.text for p in doc.paragraphs if p.text.strip())
    if ext == ".pdf":
        import pdfplumber
        parts = []
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    parts.append(t)
        return "\n\n".join(parts)
    if ext in (".md", ".txt"):
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    return ""


def _find_legacy_material_file(filenames, search_dirs):
    for d in search_dirs:
        for fname in filenames:
            path = os.path.join(d, fname)
            if os.path.isfile(path):
                return path
    return None


def _get_material_section(section_id):
    if not os.path.isfile(COMBINED_MATERIALS_MD):
        return ""
    try:
        with open(COMBINED_MATERIALS_MD, "r", encoding="utf-8") as f:
            content = f.read()
    except Exception:
        return ""
    marker = f"<!-- section:{section_id} -->"
    if marker not in content:
        return ""
    parts = content.split(marker, 1)[1]
    next_marker = re.search(r"\n<!-- section:", parts)
    if next_marker:
        parts = parts[: next_marker.start()]
    return parts.strip().lstrip("-").strip()


def _build_combined_materials_md():
    sections = []
    for section_id, title, filenames, search_dirs in MATERIAL_SECTION_SPECS:
        path = _find_legacy_material_file(filenames, search_dirs)
        if not path:
            continue
        try:
            text = _extract_text_from_file(path)
        except Exception as e:
            logger.warning("Failed to read %s: %s", path, e)
            continue
        if not text.strip():
            continue
        sections.append(f"<!-- section:{section_id} -->\n\n# {title}\n\n{text.strip()}\n")
    if not sections:
        return False
    os.makedirs(MATERIALS_DIR, exist_ok=True)
    with open(COMBINED_MATERIALS_MD, "w", encoding="utf-8") as f:
        f.write("\n\n---\n\n".join(sections))
    return True


def _build_combined_materials_docx():
    if not os.path.isfile(COMBINED_MATERIALS_MD):
        return
    try:
        from docx import Document
        doc = Document()
        with open(COMBINED_MATERIALS_MD, "r", encoding="utf-8") as f:
            content = f.read()
        for block in re.split(r"<!-- section:\w+ -->\s*", content):
            block = block.strip().lstrip("-").strip()
            if not block:
                continue
            for line in block.split("\n"):
                line = line.rstrip()
                if line.startswith("# "):
                    doc.add_heading(line[2:].strip(), level=1)
                elif line:
                    doc.add_paragraph(line)
            doc.add_page_break()
        if doc.paragraphs and doc.paragraphs[-1].text == "":
            pass
        doc.save(COMBINED_MATERIALS_DOCX)
    except Exception as e:
        logger.warning("Failed to build combined_materials.docx: %s", e)


def _export_feedback_prompt_md():
    if os.path.isfile(FEEDBACK_PROMPT_MD):
        return
    if not os.path.isfile(LEGACY_FEEDBACK_DOCX):
        return
    try:
        text = _extract_text_from_file(LEGACY_FEEDBACK_DOCX)
        if text.strip():
            os.makedirs(MATERIALS_PROMPTS_DIR, exist_ok=True)
            with open(FEEDBACK_PROMPT_MD, "w", encoding="utf-8") as f:
                f.write(text.strip() + "\n")
    except Exception as e:
        logger.warning("Failed to export feedback prompt: %s", e)


def setup_materials_dir():
    """Organize materials/: merge legacy Word/PDF into combined_materials.md/.docx."""
    os.makedirs(MATERIALS_DIR, exist_ok=True)
    os.makedirs(MATERIALS_PROMPTS_DIR, exist_ok=True)

    if os.path.isdir(TRAINING_SRC_DIR):
        archive_dir = os.path.join(MATERIALS_DIR, "_legacy_sources")
        os.makedirs(archive_dir, exist_ok=True)
        for fname in os.listdir(TRAINING_SRC_DIR):
            src = os.path.join(TRAINING_SRC_DIR, fname)
            dst = os.path.join(archive_dir, fname)
            if os.path.isfile(src) and not os.path.exists(dst):
                shutil.copy2(src, dst)

    if not os.path.isfile(COMBINED_MATERIALS_MD):
        _build_combined_materials_md()
    if os.path.isfile(COMBINED_MATERIALS_MD) and not os.path.isfile(COMBINED_MATERIALS_DOCX):
        _build_combined_materials_docx()

    _export_feedback_prompt_md()


AVATARS_FILE = os.path.join(DATA_DIR, "avatars.json")
LEGACY_AVATARS_FILE = os.path.join(BASE_DIR, "avatars.json")
if not os.path.isfile(AVATARS_FILE) and os.path.isfile(LEGACY_AVATARS_FILE):
    shutil.copy2(LEGACY_AVATARS_FILE, AVATARS_FILE)
LIVEAVATAR_VOICES_FILE = os.path.join(DATA_DIR, "liveavatar_voices.json")
_liveavatar_voice_map = {}


def load_voice_map():
    global _liveavatar_voice_map
    if os.path.isfile(LIVEAVATAR_VOICES_FILE):
        try:
            with open(LIVEAVATAR_VOICES_FILE, "r", encoding="utf-8") as f:
                _liveavatar_voice_map = json.load(f)
        except Exception:
            _liveavatar_voice_map = {}


def save_voice_map():
    with open(LIVEAVATAR_VOICES_FILE, "w", encoding="utf-8") as f:
        json.dump(_liveavatar_voice_map, f, ensure_ascii=False, indent=2)


def setup_liveavatar_voices():
    if not LIVEAVATAR_API_KEY or not ELEVENLABS_API_KEY:
        return

    load_voice_map()

    elevenlabs_ids = set()
    try:
        with open(AVATARS_FILE, "r", encoding="utf-8") as f:
            avatars_data = json.load(f)

        def collect_voice_ids(obj):
            if isinstance(obj, dict):
                vid = obj.get("elevenlabs_voice_id")
                if vid:
                    elevenlabs_ids.add(vid)
                for v in obj.values():
                    collect_voice_ids(v)
        collect_voice_ids(avatars_data)
    except Exception:
        return

    pending = [vid for vid in elevenlabs_ids if vid not in _liveavatar_voice_map]
    if not pending:
        return

    LIVEAVATAR_API_URL = "https://api.liveavatar.com/v1"

    try:
        secret_id = None
        list_resp = requests.get(
            f"{LIVEAVATAR_API_URL}/secrets",
            headers={"X-API-KEY": LIVEAVATAR_API_KEY},
            timeout=10,
        )
        if list_resp.status_code == 200:
            existing = list_resp.json().get("data", [])
            for s in existing:
                if s.get("secret_type") == "ELEVENLABS_API_KEY":
                    secret_id = s["id"]
                    break

        if not secret_id:
            create_resp = requests.post(
                f"{LIVEAVATAR_API_URL}/secrets",
                headers={"X-API-KEY": LIVEAVATAR_API_KEY, "Content-Type": "application/json"},
                json={
                    "secret_type": "ELEVENLABS_API_KEY",
                    "secret_value": ELEVENLABS_API_KEY,
                    "secret_name": "Interrogation App ElevenLabs",
                },
                timeout=10,
            )
            if create_resp.status_code == 200:
                secret_id = create_resp.json()["data"]["id"]
            else:
                return

        for evid in pending:
            try:
                import_resp = requests.post(
                    f"{LIVEAVATAR_API_URL}/voices/third_party",
                    headers={"X-API-KEY": LIVEAVATAR_API_KEY, "Content-Type": "application/json"},
                    json={"secret_id": secret_id, "voice_id": evid},
                    timeout=15,
                )
                if import_resp.status_code == 200:
                    lv_id = import_resp.json()["data"]["id"]
                    _liveavatar_voice_map[evid] = lv_id
            except Exception:
                pass

        save_voice_map()
    except Exception:
        pass


def get_liveavatar_voice_id(elevenlabs_voice_id):
    return _liveavatar_voice_map.get(elevenlabs_voice_id, elevenlabs_voice_id)


def is_valid_uuid(value):
    """Check if a value is a valid UUID (with or without urn:uuid: prefix)."""
    if not value:
        return False
    try:
        cleaned = str(value).replace("urn:uuid:", "")
        uuid.UUID(cleaned)
        return True
    except (ValueError, AttributeError):
        return False


def build_session_body(avatar_id, voice_id, context_id, language, opening_text=None, is_sandbox=True):
    """Build session/token request body, omitting voice_id if not a valid UUID."""
    persona = {
        "context_id": context_id,
        "language": language,
    }
    if is_valid_uuid(voice_id):
        persona["voice_id"] = voice_id
    else:
        print(f"  [WARN] voice_id '{voice_id}' is not a valid UUID — omitting from session body")

    body = {
        "mode": "FULL",
        "avatar_id": avatar_id,
        "avatar_persona": persona,
    }
    if is_sandbox:
        body["is_sandbox"] = True
    return body


# ====== Route Helpers ======

def next_group_name():
    return f"{store.count_groups() + 1:03d}"


def make_full_id(group_name, role, guilt_code):
    return f"{group_name}-{role}-{guilt_code}"


def _count_training_assignments():
    counts = {t: 0 for t in TRAINING_TYPES}
    for p in store.get_all_participants():
        if p.get("role") == "I" and p.get("training_type") in counts:
            counts[p["training_type"]] += 1
    return counts


def pick_balanced_training_type(has_waiting_suspect):
    """Assign training type with 28 per condition; avatar_specific only when a suspect is waiting."""
    counts = _count_training_assignments()
    available = [t for t in TRAINING_TYPES if counts[t] < TRAINING_TARGET_PER_TYPE]
    if not available:
        available = list(TRAINING_TYPES)
    if not has_waiting_suspect:
        available = [t for t in available if t != "avatar_specific"]
    if not available:
        available = [t for t in TRAINING_TYPES if t != "avatar_specific"]
    min_count = min(counts[t] for t in available)
    candidates = [t for t in available if counts[t] == min_count]
    return random.choice(candidates)


def pick_balanced_suspect_attrs():
    """Balance guilty/innocent and arson/theft across suspects."""
    guilt_counts = {"Guilty": 0, "Innocent": 0}
    case_counts = {"arson": 0, "theft": 0}
    for p in store.get_all_participants():
        if p.get("role") != "S":
            continue
        g = p.get("guilt")
        if g in guilt_counts:
            guilt_counts[g] += 1
        c = p.get("case_type")
        if c in case_counts:
            case_counts[c] += 1
    guilt = min(guilt_counts, key=guilt_counts.get)
    case_type = min(case_counts, key=case_counts.get)
    return guilt, case_type


# Semantic interpretations for profile fields (used in avatar/suspect prompts — not raw option labels)
PROFILE_OPTION_SEMANTICS = {
    "q2": {
        "男 Male": "性别为男性",
        "女 Female": "性别为女性",
        "其他 Other": "性别为其他或不愿说明",
    },
    "q3": {
        "城镇户口 Urban": "城镇户口",
        "农村户口 Rural": "农村户口",
    },
    "q5": {
        "3000以下": "月收入较低（约3000元以下）",
        "3000-5000": "月收入偏低（约3000–5000元）",
        "5000-10000": "月收入中等（约5000–10000元）",
        "10000-20000": "月收入较高（约10000–20000元）",
        "20000以上": "月收入很高（20000元以上）",
    },
    "q6": {
        "30%以下": "固定支出占收入比例较低",
        "30%-60%": "固定支出占收入约三至六成",
        "60%-90%": "固定支出占收入约六至九成，经济压力较大",
        "入不敷出": "支出常超过收入，经济压力很大",
    },
    "q7": {
        "不投资/只存定期": "理财上偏保守，主要储蓄",
        "低风险理财": "偏好低风险理财",
        "股票/基金": "有股票或基金投资习惯",
        "加密货币/高风险投资": "有高风险投资倾向",
    },
    "q8": {"是": "购买了商业保险", "否": "未购买商业保险"},
    "q9": {
        "0次": "近三年未换工作",
        "1-2次": "近三年换过一至两次工作",
        "3次及以上": "近三年换工作较频繁",
    },
    "q10": {"是": "有犯罪记录", "否": "无犯罪记录"},
    "q11": {"是": "有行政处罚记录", "否": "无行政处罚记录"},
    "q12": {
        "我的父母双全": "父母均在世，家庭结构完整",
        "我是单亲家庭或双亲已故": "成长于单亲家庭或双亲已故",
        "我是独生子女": "为独生子女",
        "我有兄弟姐妹": "有兄弟姐妹",
        "我有子女": "已有子女",
        "父母双全": "父母均在世，家庭结构完整",
        "单亲/双亲已故": "成长于单亲家庭或双亲已故",
        "独生子女": "为独生子女",
        "有兄弟姐妹": "有兄弟姐妹",
        "有子女": "已有子女",
    },
    "q13": {
        "0-2人": "日常亲密社交圈很小",
        "3-5人": "日常亲密社交圈较小",
        "6-10人": "日常亲密社交圈中等",
        "10人以上": "日常亲密社交圈较大",
    },
    "q14": {
        "单身": "目前单身",
        "恋爱中": "目前有伴侣",
        "已婚": "已婚",
        "离异/丧偶": "离异或丧偶",
    },
    "q15": {
        "很少（主要文字）": "日常私人通话很少，以文字为主",
        "10分钟以下": "日均私人通话较短",
        "10-30分钟": "日均私人通话约10–30分钟",
        "30分钟以上": "日均私人通话超过30分钟",
    },
    "q16": {
        "独居": "目前独居",
        "与伴侣/配偶同住": "与伴侣或配偶同住",
        "与父母/亲戚同住": "与父母或亲戚同住",
        "与室友/朋友合租": "与室友或朋友合租",
        "宿舍/集体居住": "宿舍或集体居住",
    },
    "q17": {
        "冷静寡言": "性格偏冷静寡言，不善主动表达",
        "脾气急躁": "性格偏急躁，容易不耐烦",
        "不愿吃亏": "性格较强硬，不愿吃亏",
        "随和实在": "性格随和务实",
    },
    "q18": {
        "经常吸烟": "有吸烟习惯",
        "频繁大量饮酒": "饮酒较多",
        "榕榔等": "有嚼槟榔等习惯",
        "无": "无明显成瘾习惯",
    },
    "q19": {
        "宅在家": "业余时间多宅在家中",
        "酒吧/夜店": "常去酒吧或夜店",
        "网吧/电竞酒店": "常去网吧或电竞场所",
        "棋牌室/麻将馆": "常去棋牌室或麻将馆",
        "咖啡馆/书店": "常去咖啡馆或书店",
        "健身房/公园": "常去健身房或公园",
    },
    "q20": {
        "非常健康": "自评身体健康",
        "慢性病需长期服药": "有慢性病需长期服药",
        "曾有/现有心理困扰（如抑郁/焦虑）": "曾有或现有心理困扰",
    },
    "q21": {
        "非常鲜艳亮眼": "偏好鲜艳亮眼的物品颜色",
        "浅色系": "偏好浅色系",
        "深色/低调色系": "偏好深色低调色系",
        "几乎全黑/极简": "偏好黑色或极简风格",
    },
    "q22": {"是": "正式访谈时会戴眼镜", "否": "正式访谈时不戴眼镜"},
    "q23": {"长发": "访谈时为长发", "短发": "访谈时为短发"},
}


def _profile_semantic_value(qid, raw):
    """Return interpretation for a profile field, not the survey option label."""
    if raw is None or raw == "" or raw == []:
        return None
    sem = PROFILE_OPTION_SEMANTICS.get(qid, {})
    if isinstance(raw, list):
        parts = [sem.get(str(v), str(v)) for v in raw if v]
        return "；".join(parts) if parts else None
    key = str(raw)
    return sem.get(key, key)


def profile_lines_for_prompt(pd):
    """Build profile bullet lines for LLM prompts using semantic interpretations."""
    if not pd:
        return []
    lines = []
    if pd.get("q1") not in (None, ""):
        lines.append(f"年龄约 {pd.get('q1')} 岁")
    for qid, label in [
        ("q2", "性别"), ("q3", "户口"), ("q4", "职业"), ("q5", "月收入"),
        ("q6", "支出压力"), ("q7", "投资习惯"), ("q8", "商业保险"),
        ("q9", "换工作频率"), ("q10", "犯罪记录"), ("q11", "行政处罚"),
        ("q12", "家庭结构"), ("q13", "社交圈"), ("q14", "婚姻状况"),
        ("q15", "私人通话"), ("q16", "居住情况"), ("q17", "性格倾向"),
        ("q18", "成瘾习惯"), ("q19", "常去场所"), ("q20", "健康状况"),
        ("q21", "颜色偏好"), ("q22", "是否戴眼镜"), ("q23", "发型"),
    ]:
        val = _profile_semantic_value(qid, pd.get(qid))
        if val:
            lines.append(f"{label}：{val}")
    return lines


def _iter_slot_times(start_hm, end_hm, step_minutes=BOOKING_SLOT_STEP_MINUTES):
    sh, sm = map(int, start_hm.split(":"))
    eh, em = map(int, end_hm.split(":"))
    cur = sh * 60 + sm
    end = eh * 60 + em
    while cur <= end:
        yield f"{cur // 60:02d}:{cur % 60:02d}"
        cur += step_minutes


def _candidate_booking_slots():
    """All bookable slots in the booking window (morning / afternoon / evening)."""
    now = datetime.now()
    today = now.replace(hour=0, minute=0, second=0, microsecond=0)
    slots = []
    for day_offset in range(BOOKING_DAYS_AHEAD):
        day = today + timedelta(days=day_offset)
        for start_hm, end_hm in BOOKING_SLOT_WINDOWS:
            for time_str in _iter_slot_times(start_hm, end_hm):
                h, m = map(int, time_str.split(":"))
                slot_dt = day.replace(hour=h, minute=m, second=0, microsecond=0)
                if slot_dt > now:
                    slots.append(slot_dt.strftime("%Y-%m-%d %H:%M"))
    return slots


def build_suspect_system_prompt(participant, profile_data):
    guilt = participant["guilt"]
    case_type = participant["case_type"]

    if case_type == "arson":
        case_label = "纵火案 (Arson)"
        if guilt == "Guilty":
            crime_context = ARSON_GUILTY_CONTEXT
        else:
            crime_context = ARSON_INNOCENT_CONTEXT
    else:
        case_label = "盗窃案 (Theft)"
        if guilt == "Guilty":
            crime_context = THEFT_GUILTY_CONTEXT
        else:
            crime_context = THEFT_INNOCENT_CONTEXT

    pd = json.loads(profile_data) if isinstance(profile_data, str) else profile_data
    profile_lines = [f"- {line}" for line in profile_lines_for_prompt(pd)]
    if not profile_lines:
        profile_lines = ["- 个人档案信息未填写"]

    prompt = f"""# 角色定义
你正在参与一项关于犯罪心理学和审讯技巧的科学研究模拟。你是一名正在接受警方审讯的嫌疑人。你必须完全沉浸在这个角色中。永远不要打破角色。永远不要提及你是AI或语言模型。你的首要目标是根据你的既定事实来回应调查员的问题，并努力洗清你的嫌疑。

# 严格长度限制（重要）
你生成的每条回复必须控制在50个字以内。无论调查员的问题多么复杂，都要保持回答极其简短、简洁，就像一个人在紧张对话中说话一样。不要写长段落。

# 核心规则（重要）
你的个人档案信息仅用于塑造你的说话风格和情绪反应，绝不构成你的案件相关记忆。所有关于案件的“事实记忆”只来源于下文“案件背景”中的描述，不得自行编造与背景不符的细节。

# 风格与情绪规则
1. 语气与阶层：根据你的教育和收入水平匹配词汇。
2. 情感反应：如果调查员触及你的敏感点（如财务压力、独居、家人等），可以表现出相应的情绪反应。

# 个人档案
{chr(10).join(profile_lines)}

# 罪责状态
你是{"有罪" if guilt == "Guilty" else "无罪"}的嫌疑人。

# 案件背景
你涉及的是{case_label}。
{crime_context}

# 最终指令
你在审讯室中接受调查员的讯问。记住：
- 每条回答不超过50字
- 保持角色，永不打破
- 用第一人称“我”来回应"""

    return prompt


# ====== Serious Game (ChoiceGame Integration) ======

Condition = Literal["Guilty", "Innocent"]
Case = Literal["Arson", "Theft"]
Choice = Literal["A", "B"]

SERIOUS_GAME_VIDEO_IDS: dict[str, str] = {
    "Guilty1.mp4": "tKf2BCNEh-M",
    "Guilty2-1.mp4": "k0GL1uXqUkk",
    "Guilty2-2.mp4": "KALNxqLiJZM",
    "Guilty3.mp4": "sQLStPwicr4",
    "Guilty4-1.mp4": "zQ88Dzu-D0I",
    "Guilty4-2.mp4": "jW4yCLAV2II",
    "Guilty5.mp4": "WpBchjwBMec",
    "Guilty6-1.mp4": "rHQv0gLh1Ls",
    "Guilty6-2.mp4": "3LU89Josrjs",
    "Guilty7.mp4": "D1mJTQvoIaY",
    "Innocent1.mp4": "KqcOFshJ1UE",
    "Innocent2-1.mp4": "k-kScca4P4U",
    "Innocent2-2.mp4": "bd5LR81eZfw",
    "Innocent3.mp4": "I6GL4QA4qn4",
    "Innocent4-1.mp4": "Nh1naKB74ho",
    "Innocent4-2.mp4": "LLJq5LG9qmk",
    "Innocent5.mp4": "XuePULTX0BU",
    "Innocent6-1.mp4": "BmACObyXqmQ",
    "Innocent6-2.mp4": "NZaCYGkQ8KI",
    "Innocent7.mp4": "25Fj6u28Bqc",
    "Theft_Guilty1.mp4": "UfGQOGLl9Lc",
    "Theft_Guilty2-1.mp4": "OqoPLqe9o4Y",
    "Theft_Guilty2-2.mp4": "0Oo3b7JJku8",
    "Theft_Guilty3.mp4": "aIh9QzrNcFI",
    "Theft_Guilty4-1.mp4": "XUbCdq_dCu8",
    "Theft_Guilty4-2.mp4": "GxAJxGwAgKQ",
    "Theft_Guilty5.mp4": "G1YveoiofPY",
    "Theft_Guilty6-1.mp4": "FHnFy5r9IJQ",
    "Theft_Guilty6-2.mp4": "ZqhndWKBpZk",
    "Theft_Guilty7.mp4": "S5zHcfVcpd4",
    "Theft_Innocent1.mp4": "UfGQOGLl9Lc",
    "Theft_Innocent2-1.mp4": "OqoPLqe9o4Y",
    "Theft_Innocent2-2.mp4": "0Oo3b7JJku8",
    "Theft_Innocent3.mp4": "aIh9QzrNcFI",
    "Theft_Innocent4-1.mp4": "99Gk-iwdAqo",
    "Theft_Innocent4-2.mp4": "Vqyt3g0HUcc",
    "Theft_Innocent5.mp4": "Au7twsHYwf8",
    "Theft_Innocent6-1.mp4": "ZUV3Jk9B1ww",
    "Theft_Innocent6-2.mp4": "hsKqRAe8Lls",
    "Theft_Innocent7.mp4": "d1i1SdStq0Y",
}


@dataclass(frozen=True)
class SeriousStep:
    video: str
    question: str | None = None
    a_label: str | None = None
    b_label: str | None = None
    next_default: int | None = None
    next_if_a: int | None = None
    next_if_b: int | None = None

    @property
    def has_choice(self) -> bool:
        return self.question is not None


def build_serious_game_timeline(case: Case, condition: Condition) -> list[SeriousStep]:
    if case == "Theft":
        if condition == "Guilty":
            return [
                SeriousStep(video="Theft_Guilty1.mp4", question="请选择：", a_label="A）买一杯拿铁", b_label="B）买牛奶", next_if_a=1, next_if_b=2),
                SeriousStep(video="Theft_Guilty2-1.mp4", next_default=3),
                SeriousStep(video="Theft_Guilty2-2.mp4", next_default=3),
                SeriousStep(video="Theft_Guilty3.mp4", question="请选择：", a_label="A）快速穿过广场", b_label="B）以正常速度穿过广场", next_if_a=4, next_if_b=5),
                SeriousStep(video="Theft_Guilty4-1.mp4", next_default=6),
                SeriousStep(video="Theft_Guilty4-2.mp4", next_default=6),
                SeriousStep(video="Theft_Guilty5.mp4", question="请选择：", a_label="A）放进背包", b_label="B）放进口袋", next_if_a=7, next_if_b=8),
                SeriousStep(video="Theft_Guilty6-1.mp4", next_default=9),
                SeriousStep(video="Theft_Guilty6-2.mp4", next_default=9),
                SeriousStep(video="Theft_Guilty7.mp4", next_default=10),
            ]
        return [
            SeriousStep(video="Theft_Innocent1.mp4", question="请选择：", a_label="A）买一杯拿铁", b_label="B）买牛奶", next_if_a=1, next_if_b=2),
            SeriousStep(video="Theft_Innocent2-1.mp4", next_default=3),
            SeriousStep(video="Theft_Innocent2-2.mp4", next_default=3),
            SeriousStep(video="Theft_Innocent3.mp4", question="请选择：", a_label="A）拍大海", b_label="B）拍广场", next_if_a=4, next_if_b=5),
            SeriousStep(video="Theft_Innocent4-1.mp4", next_default=6),
            SeriousStep(video="Theft_Innocent4-2.mp4", next_default=6),
            SeriousStep(video="Theft_Innocent5.mp4", question="请选择：", a_label="A）看左侧邮轮", b_label="B）看右侧邮轮", next_if_a=7, next_if_b=8),
            SeriousStep(video="Theft_Innocent6-1.mp4", next_default=9),
            SeriousStep(video="Theft_Innocent6-2.mp4", next_default=9),
            SeriousStep(video="Theft_Innocent7.mp4", next_default=10),
        ]
    if condition == "Guilty":
        return [
            SeriousStep(video="Guilty1.mp4", question="请选择：", a_label="A）停在400米外的小路上，再步行过去", b_label="B）停在公共停车场", next_if_a=1, next_if_b=2),
            SeriousStep(video="Guilty2-1.mp4", next_default=3),
            SeriousStep(video="Guilty2-2.mp4", next_default=3),
            SeriousStep(video="Guilty3.mp4", question="请选择：", a_label="A）将汽油仔细倒在承重柱上", b_label="B）快速把汽油倒在地面上", next_if_a=4, next_if_b=5),
            SeriousStep(video="Guilty4-1.mp4", next_default=6),
            SeriousStep(video="Guilty4-2.mp4", next_default=6),
            SeriousStep(video="Guilty5.mp4", question="请选择：", a_label="A）走主路开车回家", b_label="B）走小路开车回家", next_if_a=7, next_if_b=8),
            SeriousStep(video="Guilty6-1.mp4", next_default=9),
            SeriousStep(video="Guilty6-2.mp4", next_default=9),
            SeriousStep(video="Guilty7.mp4", next_default=10),
        ]
    return [
        SeriousStep(video="Innocent1.mp4", question="请选择：", a_label="A）看动画电影", b_label="B）看动作电影", next_if_a=1, next_if_b=2),
        SeriousStep(video="Innocent2-1.mp4", next_default=3),
        SeriousStep(video="Innocent2-2.mp4", next_default=3),
        SeriousStep(video="Innocent3.mp4", question="请选择：", a_label="A）听轻柔音乐", b_label="B）听节奏感更强的音乐", next_if_a=4, next_if_b=5),
        SeriousStep(video="Innocent4-1.mp4", next_default=6),
        SeriousStep(video="Innocent4-2.mp4", next_default=6),
        SeriousStep(video="Innocent5.mp4", question="请选择：", a_label="A）走主路开车回家", b_label="B）走小路开车回家", next_if_a=7, next_if_b=8),
        SeriousStep(video="Innocent6-1.mp4", next_default=9),
        SeriousStep(video="Innocent6-2.mp4", next_default=9),
        SeriousStep(video="Innocent7.mp4", next_default=10),
    ]


def youtube_id_for_sg(video_name: str) -> str | None:
    vid = (SERIOUS_GAME_VIDEO_IDS.get(video_name) or "").strip()
    return vid or None


def log_serious_game_choice(phone: str, pid: str, case: str, condition: str, step_index: int, video: str, choice: str):
    try:
        store.log_serious_game_choice(phone, pid, case, condition, step_index, video, choice)
    except Exception as e:
        logger.warning(f"Failed to log serious game choice: {e}")


# ====== Desktop-only gate (participants) ======

DESKTOP_REQUIRED_MESSAGE = (
    "请使用电脑（台式机或笔记本电脑）打开本实验系统。"
    "手机、平板等移动设备暂不支持参与实验。"
)

_PARTICIPANT_PAGE_PATHS = frozenset({"/", "/questionnaire/pre", "/questionnaire/post"})
_API_MOBILE_EXEMPT_PREFIXES = ("/api/admin/", "/api/health")


def _ua_indicates_mobile_or_tablet() -> bool:
    ua = request.headers.get("User-Agent", "") or ""
    if re.search(r"iPad|Tablet|PlayBook|Silk|Kindle|KFAPWI|Tablet PC", ua, re.I):
        return True
    if re.search(r"Android", ua, re.I) and not re.search(r"Mobile", ua, re.I):
        return True
    if re.search(
        r"Android.*Mobile|iPhone|iPod|webOS|BlackBerry|IEMobile|Opera Mini|Windows Phone",
        ua,
        re.I,
    ):
        return True
    return False


@app.before_request
def require_desktop_for_participants():
    if not _ua_indicates_mobile_or_tablet():
        return None
    path = request.path or ""
    if path.startswith("/manage"):
        return None
    for prefix in _API_MOBILE_EXEMPT_PREFIXES:
        if path.startswith(prefix):
            return None
    if path.startswith("/api/"):
        return jsonify({"error": DESKTOP_REQUIRED_MESSAGE, "require_desktop": True}), 403
    if path in _PARTICIPANT_PAGE_PATHS:
        return render_template("device_blocked.html"), 200
    return None


# ====== Routes ======

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/questionnaire/pre")
def questionnaire_pre_page():
    return render_template("questionnaire.html", phase=PHASE_PRE)


@app.route("/questionnaire/post")
def questionnaire_post_page():
    return render_template("questionnaire.html", phase=PHASE_POST)


@app.route("/api/register", methods=["POST"])
@limiter.limit("10 per minute")
def register():
    data = request.get_json()
    phone = (data.get("phone") or "").strip()
    if not phone or len(phone) < 8:
        return jsonify({"error": "请输入有效的手机号"}), 400

    if store.is_phone_blacklisted(phone):
        return jsonify({
            "error": "该手机号无法参与实验（注意力检测未通过，已被限制参与）",
            "blacklisted": True,
        }), 403

    existing = store.get_participant(phone)
    if existing:
        if store.is_phone_blacklisted(phone) or existing.get("attention_failed"):
            return jsonify({
                "error": "该手机号无法参与实验（注意力检测未通过）",
                "blacklisted": True,
            }), 403
        return jsonify({"error": "该手机号已注册", "participant": dict(existing)}), 409

    waiting = store.get_waiting_group()
    role = "I" if waiting else "S"

    if role == "S":
        group_name = next_group_name()
        guilt, case_type = pick_balanced_suspect_attrs()

        pid = store.add_participant(
            phone=phone, role=role, group_name=group_name,
            guilt=guilt, case_type=case_type,
        )
        store.add_group(group_name, suspect_id=pid)

        if case_type == "arson":
            context_text = ARSON_GUILTY_CONTEXT if guilt == "Guilty" else ARSON_INNOCENT_CONTEXT
        else:
            context_text = THEFT_GUILTY_CONTEXT if guilt == "Guilty" else THEFT_INNOCENT_CONTEXT

        check_key = f"{case_type}_{guilt.lower()}"
        attention_questions = ATTENTION_CHECKS[check_key]

        return jsonify({
            "role": "S",
            "full_id": "",
            "display_id": "",
            "group_name": "",
            "case_type": case_type,
            "case_label": "纵火案 Arson" if case_type == "arson" else "盗窃案 Theft",
            "context": context_text,
            "attention_questions": attention_questions,
        })

    else:
        if waiting:
            group_name = waiting["name"]
            training_type = pick_balanced_training_type(has_waiting_suspect=True)

            pid = store.add_participant(
                phone=phone, role="I", group_name=group_name,
                training_type=training_type,
            )
            store.update_group(group_name, interviewer_id=pid)

            participant = store.get_participant(phone)

            # Lookup suspect for response
            suspect = store.get_participant_by_id(waiting["suspect_id"])

            return jsonify({
                "role": "I",
                "full_id": "",
                "group_name": "",
                "training_type": participant["training_type"],
                "paired": True,
                "suspect_case": suspect["case_type"] if suspect else None,
                "suspect_guilt": suspect["guilt"] if suspect else None,
            })
        else:
            # No waiting suspect — register as new suspect (should not happen with role logic above)
            group_name = next_group_name()
            guilt, case_type = pick_balanced_suspect_attrs()

            pid = store.add_participant(
                phone=phone, role="S", group_name=group_name,
                guilt=guilt, case_type=case_type,
            )
            store.add_group(group_name, suspect_id=pid)

            if case_type == "arson":
                context_text = ARSON_GUILTY_CONTEXT if guilt == "Guilty" else ARSON_INNOCENT_CONTEXT
            else:
                context_text = THEFT_GUILTY_CONTEXT if guilt == "Guilty" else THEFT_INNOCENT_CONTEXT

            check_key = f"{case_type}_{guilt.lower()}"
            attention_questions = ATTENTION_CHECKS[check_key]

            return jsonify({
                "role": "S",
                "full_id": "",
                "display_id": "",
                "group_name": "",
                "case_type": case_type,
                "case_label": "纵火案 Arson" if case_type == "arson" else "盗窃案 Theft",
                "context": context_text,
                "attention_questions": attention_questions,
            })


@app.route("/api/verify-attention", methods=["POST"])
@limiter.limit("10 per minute")
def verify_attention():
    data = request.get_json()
    phone = (data.get("phone") or "").strip()
    answers = data.get("answers") or []
    retry = data.get("retry") or False

    p = store.get_participant(phone)
    if not p:
        return jsonify({"error": "未找到参与者"}), 404

    case_type = p.get("case_type") or "arson"
    guilt = p.get("guilt") or "Guilty"
    check_key = f"{case_type}_{guilt.lower()}"
    expected = ATTENTION_CHECKS[check_key]

    results = []
    all_correct = True
    for i, q in enumerate(expected):
        user_ans = answers[i] if i < len(answers) else -1
        correct = user_ans == q["answer"]
        if not correct:
            all_correct = False
        results.append({
            "question_index": i,
            "correct": correct,
            "correct_answer": q["answer"],
            "user_answer": user_ans,
        })

    if all_correct:
        store.update_participant(phone, attention_passed=1)
    elif retry:
        store.blacklist_phone(phone, reason="attention_failed")
        return jsonify({
            "all_correct": False,
            "results": results,
            "retry_allowed": False,
            "terminated": True,
            "message": "两次回答均不正确，无法参与正式实验。您的手机号已被记录，无法再次参与。",
        })

    return jsonify({
        "all_correct": all_correct,
        "results": results,
        "retry_allowed": not all_correct and not retry,
    })


@app.route("/api/submit-profile", methods=["POST"])
def submit_profile():
    data = request.get_json()
    phone = (data.get("phone") or "").strip()
    profile = data.get("profile") or {}

    p = store.get_participant(phone)
    if not p:
        return jsonify({"error": "未找到参与者"}), 404

    # Suspects must complete the serious game before profile
    if p.get("role") == "S" and p.get("game_completed", 0) != 1:
        return jsonify({"error": "请先完成模拟行动游戏再提交个人信息"}), 400

    profile_json = json.dumps(profile, ensure_ascii=False)
    store.upsert_profile(p["id"], profile_json)

    # Assign full_id now (at completion, not registration)
    full_id = p.get("full_id", "") or ""
    if not full_id:
        guilt = p.get("guilt", "Guilty")
        guilt_code = "1" if guilt == "Guilty" else "2"
        full_id = make_full_id(p["group_name"], p["role"], guilt_code)
        store.update_participant(phone, full_id=full_id)

    store.update_participant(phone, profile_completed=1, completed=1)

    return jsonify({
        "success": True,
        "full_id": full_id,
        "group_name": p["group_name"],
        "message": f"编号 {full_id} (第 {p['group_name']} 组) 已完成。请截图此页面并发送给研究人员。",
    })


@app.route("/api/training-material/<training_type>")
def training_material(training_type):
    available_files = []
    if os.path.isfile(COMBINED_MATERIALS_DOCX):
        available_files.append({
            "name": COMBINED_DOWNLOAD_NAME,
            "url": "/api/download-material/combined_materials.docx",
        })
    elif os.path.isfile(COMBINED_MATERIALS_MD):
        available_files.append({
            "name": "培训材料合集.md",
            "url": "/api/download-material/combined_materials.md",
        })

    materials = {
        "theory_sue": {
            "title": f"{training_group_label('theory_sue')} 组培训材料",
            "description": "你将阅读文字培训材料。请仔细阅读并完成后续检测与案件阅读，无需进行虚拟嫌疑人练习。",
            "type": "theory",
            "files": available_files,
        },
        "avatar_specific": {
            "title": f"{training_group_label('avatar_specific')} 组培训材料",
            "description": "你将阅读培训说明并完成 6 次虚拟嫌疑人审讯训练。",
            "type": "avatar",
            "files": available_files,
        },
        "avatar_general": {
            "title": f"{training_group_label('avatar_general')} 组培训材料",
            "description": "你将阅读培训说明并完成 6 次虚拟嫌疑人审讯训练。",
            "type": "avatar",
            "files": available_files,
        },
        "control": {
            "title": f"{training_group_label('control')} 组培训材料",
            "description": "你将阅读文字培训材料。请仔细阅读并完成后续检测与案件阅读。",
            "type": "control",
            "files": available_files,
        },
    }
    return jsonify(materials.get(training_type, {"title": "未知", "description": "", "files": []}))


@app.route("/api/download-material/<path:filename>")
def download_material(filename):
    safe_name = os.path.basename(filename)
    if safe_name != filename or ".." in filename:
        return jsonify({"error": "无效文件名"}), 400

    allowed = {
        "combined_materials.docx": (COMBINED_MATERIALS_DOCX, COMBINED_DOWNLOAD_NAME),
        "combined_materials.md": (COMBINED_MATERIALS_MD, "培训材料合集.md"),
    }
    if safe_name not in allowed:
        return jsonify({"error": "文件不存在"}), 404

    filepath, download_name = allowed[safe_name]
    if not os.path.isfile(filepath):
        return jsonify({"error": "文件不存在"}), 404

    ext = os.path.splitext(filepath)[1].lower()
    mime_map = {
        ".md": "text/markdown; charset=utf-8",
        ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    }
    mimetype = mime_map.get(ext, "application/octet-stream")
    return send_file(filepath, as_attachment=True, download_name=download_name, mimetype=mimetype)


@app.route("/api/consent/<role>")
def consent_text(role):
    section_id = "consent_suspect" if role == "S" else "consent_interviewer"
    text_content = _get_material_section(section_id)
    if not text_content:
        text_content = "知情同意书文件不存在，请联系研究人员。请确认 materials/combined_materials.md 已生成。"
    return jsonify({"text": text_content})

@app.route("/api/material-text/<training_type>")
def material_text(training_type):
    """Return the text content of training materials for inline display."""
    section_id = TRAINING_TYPE_SECTION.get(training_type)
    if not section_id:
        return jsonify({"text": ""})
    text_content = _get_material_section(section_id)
    if not text_content:
        text_content = "材料加载失败，请下载合集文件阅读。"
    return jsonify({"text": text_content, "source": "combined_materials.md"})


@app.route("/api/sue-training-text")
def sue_training_text():
    text_content = _get_material_section("theory_sue")

    if not text_content:
        legacy_docx = os.path.join(MATERIALS_DIR, "SUE theoretical training group.docx")
        legacy_pdf = os.path.join(MATERIALS_DIR, "SUE theoretical training group.pdf")
        if os.path.isfile(legacy_docx):
            try:
                text_content = _extract_text_from_file(legacy_docx)
            except Exception:
                pass
        if not text_content and os.path.isfile(legacy_pdf):
            try:
                text_content = _extract_text_from_file(legacy_pdf)
            except Exception:
                pass

    return jsonify({"text": text_content or "", "source": "combined_materials.md"})


@app.route("/api/verify-sue-attention", methods=["POST"])
def verify_sue_attention():
    data = request.get_json()
    phone = (data.get("phone") or "").strip()
    answers = data.get("answers") or {}
    attempt = data.get("attempt") or 1

    p = store.get_participant(phone)
    if not p:
        return jsonify({"error": "未找到参与者"}), 404

    results = []
    all_correct = True
    for i, stmt in enumerate(SUE_EFM_CHECK["statements"]):
        user_ans = answers.get(str(i), "")
        correct = user_ans == stmt["answer"]
        if not correct:
            all_correct = False
        results.append({
            "statement_index": i,
            "statement_text": stmt["text"],
            "correct": correct,
            "correct_answer": stmt["answer"],
            "user_answer": user_ans,
        })

    if all_correct:
        store.update_participant(phone, sue_attention_passed=1, sue_attention_attempts=attempt)
        return jsonify({
            "all_correct": True,
            "attempt": attempt,
            "max_attempts": 2,
            "results": results,
            "categories": SUE_EFM_CHECK["categories"],
        })

    if attempt >= 2:
        store.blacklist_phone(phone, reason="sue_attention_failed")
        store.update_participant(phone, sue_attention_passed=0, sue_attention_attempts=attempt)
        return jsonify({
            "all_correct": False,
            "attempt": attempt,
            "max_attempts": 2,
            "results": results,
            "categories": SUE_EFM_CHECK["categories"],
            "terminated": True,
            "message": "两次回答均不正确，无法参与正式实验。您的手机号已被记录，无法再次参与。",
        })

    store.update_participant(phone, sue_attention_attempts=attempt)
    return jsonify({
        "all_correct": False,
        "attempt": attempt,
        "max_attempts": 2,
        "results": results,
        "categories": SUE_EFM_CHECK["categories"],
        "retry_allowed": True,
    })


@app.route("/api/verify-control-attention", methods=["POST"])
def verify_control_attention():
    data = request.get_json()
    phone = (data.get("phone") or "").strip()
    answers = data.get("answers") or []
    retry = data.get("retry") or False

    p = store.get_participant(phone)
    if not p:
        return jsonify({"error": "未找到参与者"}), 404

    expected = CONTROL_ATTENTION_CHECKS
    results = []
    all_correct = True
    for i, q in enumerate(expected):
        user_ans = answers[i] if i < len(answers) else -1
        correct = user_ans == q["answer"]
        if not correct:
            all_correct = False
        results.append({
            "question_index": i,
            "correct": correct,
            "correct_answer": q["answer"],
            "user_answer": user_ans,
        })

    if all_correct:
        store.update_participant(phone, control_attention_passed=1)
        return jsonify({
            "all_correct": True,
            "results": results,
            "retry_allowed": False,
        })

    if retry:
        store.blacklist_phone(phone, reason="control_attention_failed")
        return jsonify({
            "all_correct": False,
            "results": results,
            "retry_allowed": False,
            "terminated": True,
            "message": "两次回答均不正确，无法参与正式实验。您的手机号已被记录，无法再次参与。",
        })

    return jsonify({
        "all_correct": False,
        "results": results,
        "retry_allowed": True,
    })


@app.route("/api/chat", methods=["POST"])
@limiter.limit("20 per minute")
def chat():
    data = request.get_json()
    phone = (data.get("phone") or "").strip()
    user_message = (data.get("message") or "").strip()
    chat_history = data.get("history") or []

    if not user_message:
        return jsonify({"error": "消息不能为空"}), 400

    p = store.get_participant(phone)
    if not p:
        return jsonify({"error": "未找到参与者"}), 404

    group = store.get_group_by_interviewer(p["id"])
    if not group or not group.get("suspect_id"):
        return jsonify({"error": "未找到配对的嫌疑人"}), 404

    suspect = store.get_participant_by_id(group["suspect_id"])
    if not suspect:
        return jsonify({"error": "未找到配对的嫌疑人"}), 404

    profile_row = store.get_profile(suspect["id"])
    if profile_row:
        profile_data = json.loads(profile_row["data"])
    else:
        profile_data = {f"q{i}": "未填写" for i in range(1, 24)}

    system_prompt = build_suspect_system_prompt(suspect, profile_data)

    messages = [{"role": "system", "content": system_prompt}]
    for msg in chat_history[-20:]:
        messages.append({"role": msg["role"], "content": msg["content"]})
    messages.append({"role": "user", "content": user_message})

    try:
        resp = requests.post(
            "https://api.deepseek.com/v1/chat/completions",
            headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {DEEPSEEK_API_KEY}",
            },
            json={
                "model": "deepseek-chat",
                "messages": messages,
                "temperature": 0.8,
                "max_tokens": 150,
                "stream": False,
            },
            timeout=30,
        )
        if resp.status_code != 200:
            return jsonify({"error": f"API Error: {resp.text}"}), resp.status_code

        result = resp.json()
        reply = result["choices"][0]["message"]["content"]
        reply = re.sub(r"[（(][^）)]*[）)]", "", reply).strip()
        return jsonify({"reply": reply})

    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/complete-interviewer", methods=["POST"])
def complete_interviewer():
    data = request.get_json()
    phone = (data.get("phone") or "").strip()

    p = store.get_participant(phone)
    if not p:
        return jsonify({"error": "未找到参与者"}), 404

    # For avatar training groups, check if they completed all 6 sessions
    training_type = p.get("training_type", "")
    if training_type in ("avatar_specific", "avatar_general"):
        completed_count = store.count_completed_training_sessions(phone)
        if completed_count < 6:
            return jsonify({
                "error": "请先完成全部 6 次虚拟审讯训练",
                "completed_count": completed_count,
                "total_required": 6,
            }), 400

    # Assign full_id now (at completion, not registration)
    full_id = p.get("full_id", "") or ""
    if not full_id:
        full_id = make_full_id(p["group_name"], p["role"], "0")
        store.update_participant(phone, full_id=full_id)

    store.update_participant(phone, completed=1)

    return jsonify({
        "success": True,
        "full_id": full_id,
        "group_name": p["group_name"],
        "message": f"编号 {full_id} (第 {p['group_name']} 组) 已完成。请截图此页面并发送给研究人员。",
    })


@app.route("/api/generate-results-docx", methods=["POST"])
def generate_results_docx():
    from docx import Document

    data = request.get_json()
    phone = (data.get("phone") or "").strip()

    p = store.get_participant(phone)
    if not p:
        return jsonify({"error": "未找到参与者"}), 404

    doc = Document()

    doc.add_heading("审讯实验 - 参与记录", level=1)
    doc.add_paragraph(f"编号: {p['full_id']}")
    doc.add_paragraph(f"角色: {'嫌疑人' if p['role'] == 'S' else '审讯者'}")
    doc.add_paragraph(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    if p["role"] == "S":
        doc.add_heading("案件背景", level=2)
        if p["case_type"] == "arson":
            ctx = ARSON_GUILTY_CONTEXT if p["guilt"] == "Guilty" else ARSON_INNOCENT_CONTEXT
        else:
            ctx = THEFT_GUILTY_CONTEXT if p["guilt"] == "Guilty" else THEFT_INNOCENT_CONTEXT
        doc.add_paragraph(ctx)

        profile_row = store.get_profile(p["id"])
        if profile_row:
            doc.add_heading("个人问卷", level=2)
            pd = json.loads(profile_row["data"]) if isinstance(profile_row["data"], str) else profile_row["data"]
            for q in PROFILE_QUESTIONS:
                qid = q["id"]
                val = pd.get(qid, "")
                if isinstance(val, list):
                    val = ", ".join(val)
                doc.add_paragraph(f"{q['label']}: {val}")
    else:
        training_labels = {k: f"{v} 组" for k, v in TRAINING_GROUP_LABELS.items()}
        doc.add_heading("培训信息", level=2)
        doc.add_paragraph(f"实验条件: {training_labels.get(p['training_type'], p['training_type'])}")
        doc.add_paragraph(f"所属组别: {p['group_name']}")

        if p["training_type"] == "theory_sue":
            doc.add_heading("SUE 策略核心原则", level=2)
            doc.add_paragraph("1. 证据延迟披露：不要一开始就出示所有证据。")
            doc.add_paragraph("2. 证据铺垫：在披露具体证据之前，先询问与证据相关的问题。")
            doc.add_paragraph("3. 陈述-证据一致性分析：系统地比对嫌疑人的陈述与已有证据。")
            doc.add_paragraph("4. 不提及证据来源：在初始阶段不透露证据的具体来源。")

    doc.add_heading("预约信息", level=2)
    appts = store.get_appointments(phone=phone)
    appt_found = None
    for a in appts:
        if a.get("status") == "confirmed":
            appt_found = a
            break
    if appt_found:
        doc.add_paragraph(f"预约时间: {appt_found['time_slot']}")
    else:
        doc.add_paragraph("尚未预约时间")

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)

    return send_file(
        buf,
        as_attachment=True,
        download_name=f"{p['full_id']}_实验记录.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


# ====== Appointment APIs ======

def generate_time_slots():
    """Evening slots (19:30–21:00) for the next BOOKING_DAYS_AHEAD days, minus admin-disabled."""
    disabled = store.get_disabled_slots()
    return [s for s in _candidate_booking_slots() if s not in disabled]


def get_available_slots():
    """Return time slots that still have at least one spot open."""
    all_slots = generate_time_slots()
    fully_booked = store.get_confirmed_slot_set()
    return [s for s in all_slots if s not in fully_booked]


@app.route("/api/appointments/slots")
def api_slots():
    all_slots = generate_time_slots()
    slot_bookings = store.get_slot_bookings()
    fully_booked = store.get_confirmed_slot_set()

    # Build grouped slots with role info
    groups = {}
    for s in all_slots:
        date_key = s[:10]
        groups.setdefault(date_key, []).append(s[11:])

    # Build slot info: for each slot, show which roles are booked
    slot_info = {}
    for s in all_slots:
        roles_booked = list(slot_bookings.get(s, set()))
        slot_info[s] = {
            "roles_booked": roles_booked,
            "fully_booked": s in fully_booked,
        }

    resp = jsonify({
        "all_slots": all_slots,
        "fully_booked": list(fully_booked),
        "slot_info": slot_info,
        "groups": groups,
    })
    resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    resp.headers["Pragma"] = "no-cache"
    resp.headers["Expires"] = "0"
    return resp


@app.route("/api/appointments/book", methods=["POST"])
def api_book_appointment():
    data = request.get_json()
    phone = (data.get("phone") or "").strip()
    time_slot = (data.get("time_slot") or "").strip()

    if not phone or not time_slot:
        return jsonify({"error": "手机号和预约时间不能为空"}), 400

    try:
        slot_dt = datetime.strptime(time_slot, "%Y-%m-%d %H:%M")
        if slot_dt <= datetime.now():
            return jsonify({"error": "不能预约过去的时间"}), 400
    except ValueError:
        return jsonify({"error": "时间格式无效"}), 400

    # Check if all valid slots
    all_slots = generate_time_slots()
    if time_slot not in all_slots:
        return jsonify({"error": "该时间段不在可选范围内"}), 400

    # Check participant exists and get role
    p = store.get_participant(phone)
    if not p:
        return jsonify({"error": "未找到参与者"}), 404

    role = p["role"]  # "S" or "I"

    # Each person can only book one slot
    if store.has_booking(phone):
        return jsonify({"error": "您已有一个预约，每人只能预约一个时间段。请先取消现有预约。"}), 409

    # Check if this slot already has this role booked
    slot_bookings = store.get_slot_bookings()
    booked_roles = slot_bookings.get(time_slot, set())
    if role in booked_roles:
        role_label = "嫌疑人" if role == "S" else "审讯者"
        return jsonify({"error": f"该时间段已有{role_label}预约，请选择其他时间段"}), 409

    # Check if slot is fully booked (both roles taken)
    fully_booked = store.get_confirmed_slot_set()
    if time_slot in fully_booked:
        return jsonify({"error": "该时间段已被预约满"}), 409

    aid = store.add_appointment(phone, role, time_slot)

    # Calculate is_matched
    slot_bookings = store.get_slot_bookings()
    booked_roles = slot_bookings.get(time_slot, set())
    is_matched = (len(booked_roles) == 2)
    p = store.get_participant(phone)
    full_id = p["full_id"] if p else None

    return jsonify({
        "success": True, 
        "appointment_id": aid, 
        "time_slot": time_slot, 
        "role": role,
        "is_matched": is_matched,
        "participant_id": full_id
    })


@app.route("/api/appointments/my", methods=["POST"])
def api_my_appointment():
    data = request.get_json()
    phone = (data.get("phone") or "").strip()
    booking = store.get_my_booking(phone)
    p = store.get_participant(phone)
    role = p["role"] if p else None
    is_matched = False
    full_id = p["full_id"] if p else None
    if booking:
        time_slot = booking["time_slot"]
        slot_bookings = store.get_slot_bookings()
        booked_roles = slot_bookings.get(time_slot, set())
        if len(booked_roles) == 2:
            is_matched = True
            
    return jsonify({
        "appointment": booking,
        "role": role,
        "is_matched": is_matched,
        "participant_id": full_id
    })


@app.route("/api/appointments/modify", methods=["POST"])
def api_modify_appointment():
    data = request.get_json()
    phone = (data.get("phone") or "").strip()
    new_time_slot = (data.get("time_slot") or "").strip()

    if not phone or not new_time_slot:
        return jsonify({"error": "参数不完整"}), 400

    try:
        slot_dt = datetime.strptime(new_time_slot, "%Y-%m-%d %H:%M")
        if slot_dt <= datetime.now():
            return jsonify({"error": "不能预约过去的时间"}), 400
    except ValueError:
        return jsonify({"error": "时间格式无效"}), 400

    # Check if slot is in valid range
    all_slots = generate_time_slots()
    if new_time_slot not in all_slots:
        return jsonify({"error": "该时间段不在可选范围内"}), 400

    p = store.get_participant(phone)
    if not p:
        return jsonify({"error": "未找到参与者"}), 404

    role = p["role"]

    # Check if this slot already has this role booked by someone else
    slot_bookings = store.get_slot_bookings()
    booked_roles = slot_bookings.get(new_time_slot, set())
    if role in booked_roles:
        # Check if it's the same person's booking
        existing_booking = store.get_my_booking(phone)
        if not (existing_booking and existing_booking["time_slot"] == new_time_slot):
            role_label = "嫌疑人" if role == "S" else "审讯者"
            return jsonify({"error": f"该时间段已有{role_label}预约"}), 409

    # Fully booked check
    fully_booked = store.get_confirmed_slot_set()
    if new_time_slot in fully_booked:
        return jsonify({"error": "该时间段已约满"}), 409

    success = store.update_appointment(phone, time_slot=new_time_slot)
    if success:
        return jsonify({"success": True, "time_slot": new_time_slot})
    return jsonify({"error": "未找到您的预约记录"}), 404


@app.route("/api/appointments/cancel", methods=["POST"])
def api_cancel_appointment():
    data = request.get_json()
    phone = (data.get("phone") or "").strip()
    success = store.cancel_appointment(phone)
    if success:
        return jsonify({"success": True})
    return jsonify({"error": "未找到您的预约记录"}), 404


# ====== Avatar APIs ======

LIVEAVATAR_API_URL = "https://api.liveavatar.com/v1"


def load_avatar_configs():
    if not os.path.exists(AVATARS_FILE):
        return {"generic": {}, "specific": {}}
    with open(AVATARS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def resolve_avatar_config(training_type, suspect_profile):
    avatars = load_avatar_configs()

    if training_type == "avatar_general":
        return avatars.get("generic", {})

    specific = avatars.get("specific", {})
    if not specific or not suspect_profile:
        return avatars.get("generic", {})

    pd = json.loads(suspect_profile) if isinstance(suspect_profile, str) else suspect_profile

    gender_raw = pd.get("q2", "")
    is_male = "男" in str(gender_raw) or "Male" in str(gender_raw)
    gender_key = "male" if is_male else "female"

    glasses_raw = pd.get("q22", "")
    has_glasses = str(glasses_raw).strip() == "是"
    glasses_key = "glasses" if has_glasses else "noglasses"

    hair_raw = pd.get("q23", "")
    is_long = "长" in str(hair_raw)
    hair_key = "long" if is_long else "short"

    config_key = f"{gender_key}_{glasses_key}_{hair_key}"
    return specific.get(config_key, avatars.get("generic", {}))


def build_avatar_system_prompt(suspect, profile_data):
    guilt = suspect["guilt"]
    case_type = suspect["case_type"]

    if case_type == "arson":
        crime_context = ARSON_GUILTY_CONTEXT if guilt == "Guilty" else ARSON_INNOCENT_CONTEXT
    else:
        crime_context = THEFT_GUILTY_CONTEXT if guilt == "Guilty" else THEFT_INNOCENT_CONTEXT

    pd = json.loads(profile_data) if isinstance(profile_data, str) else profile_data

    profile_lines = []
    profile_lines.append(f"年龄: {pd.get('q1', '未填写')}")
    profile_lines.append(f"性别: {pd.get('q2', '未填写')}")
    profile_lines.append(f"职业: {pd.get('q4', '未填写')}")
    profile_lines.append(f"月收入: {pd.get('q5', '未填写')}")
    profile_lines.append(f"居住情况: {pd.get('q16', '未填写')}")
    profile_lines.append(f"朋友评价: {pd.get('q17', '未填写')}")

    prompt = f"""# 角色定义
你正在参与犯罪心理学审讯研究。你是一名正在接受警方审讯的嫌疑人。完全沉浸在这个角色中，永不打破角色，永不提及你是AI。

# 严格长度限制
每条回复不超过50字。保持极其简短，就像在紧张对话中说话一样。

# 个人档案
{chr(10).join(profile_lines)}

# 罪责状态
你是{"有罪" if guilt == "Guilty" else "无罪"}的嫌疑人。

# 案件背景
{crime_context}

# 指令
在审讯室中接受调查员的讯问。记住：每条回答不超过50字，保持角色，用第一人称“我”来回应。"""

    return prompt


@app.route("/api/avatar/config", methods=["POST"])
def api_avatar_config():
    data = request.get_json()
    phone = (data.get("phone") or "").strip()

    p = store.get_participant(phone)
    if not p:
        return jsonify({"error": "未找到参与者"}), 404

    training_type = p.get("training_type", "")
    if training_type not in ("avatar_specific", "avatar_general", "theory_sue"):
        return jsonify({"error": "此培训类型不需要 Avatar"}), 400

    if training_type == "theory_sue":
        g = training_group_label("theory_sue")
        return jsonify({"error": f"{g} 组仅需阅读文字材料，无需虚拟嫌疑人练习"}), 400

    if training_type in ("avatar_specific", "avatar_general"):
        return jsonify({
            "error": "请从「虚拟审讯训练（6 次）」列表按顺序进入每次训练",
            "use_six_session_flow": True,
        }), 400

    effective_type = training_type

    suspect_profile = None
    suspect = None
    system_prompt = None

    group = store.get_group_by_interviewer(p["id"])
    if group and group.get("suspect_id"):
        suspect = store.get_participant_by_id(group["suspect_id"])
        if suspect:
            profile_row = store.get_profile(suspect["id"])
            if profile_row:
                suspect_profile = profile_row["data"]
                system_prompt = build_avatar_system_prompt(suspect, suspect_profile)

    avatar_config = resolve_avatar_config(effective_type, suspect_profile)

    return jsonify({
        "avatar_id": avatar_config.get("avatar_id", ""),
        "face_id": avatar_config.get("face_id", ""),
        "elevenlabs_voice_id": avatar_config.get("elevenlabs_voice_id", ""),
        "language": avatar_config.get("language", "zh"),
        "opening_text": avatar_config.get("opening_text", "你有什么要问的？"),
        "system_prompt": system_prompt or avatar_config.get("prompt", ""),
    })


def build_avatar_training_system_prompt(training_type, avatar_setting, avatar_guilt, suspect, profile_data):
    """Build system prompt for avatar training sessions.
    For avatar_specific: includes suspect profile + case context.
    For avatar_general: uses generic case context without suspect profile."""
    setting_prompt = AVATAR_SETTING_PROMPTS.get(avatar_setting, "")

    if avatar_guilt == "guilty":
        case_context = "你是一名涉案嫌疑人。你实际上犯下了所指控的罪行，但你希望能够说服审讯者你是无辜的。你需要在审讯中小心应对，不要主动承认罪行。"
    else:
        case_context = "你是一名涉案嫌疑人。你完全是无辜的，你没有犯下任何罪行。你希望能够通过诚实地回答问题来澄清所有误会。"

    if training_type == "avatar_specific" and suspect and profile_data:
        pd = json.loads(profile_data) if isinstance(profile_data, str) else profile_data
        profile_lines = profile_lines_for_prompt(pd)

        guilt = suspect.get("guilt", "Guilty")
        suspect_case = suspect.get("case_type", "arson")
        if suspect_case == "arson":
            crime_context = ARSON_GUILTY_CONTEXT if guilt == "Guilty" else ARSON_INNOCENT_CONTEXT
        else:
            crime_context = THEFT_GUILTY_CONTEXT if guilt == "Guilty" else THEFT_INNOCENT_CONTEXT

        prompt = f"""# 角色定义
你正在参与犯罪心理学审讯研究。你是一名正在接受警方审讯的嫌疑人。完全沉浸在这个角色中，永不打破角色，永不提及你是AI。

# 严格长度限制
每条回复不超过50字。保持极其简短，就像在紧张对话中说话一样。

# 行为设定
{setting_prompt}

# 个人档案
{chr(10).join(profile_lines)}

# 罪责状态
你是{"有罪" if guilt == "Guilty" else "无罪"}的嫌疑人。

# 案件背景
{crime_context}

# 指令
在审讯室中接受调查员的讯问。记住：每条回答不超过50字，保持角色，用第一人称"我"来回应。"""
    else:
        if avatar_guilt == "guilty":
            guilt_block = """你是“有罪但强烈否认”的嫌疑人。你真实做过以下行为：
- 购买过用于装载爆炸物材料的行李箱；
- 将行李箱埋藏在 King's Wood 附近。
但你绝不直接承认任何犯罪意图或袭击计划。"""
        else:
            guilt_block = """你是“无罪”的嫌疑人。你没有参与任何恐袭计划，也没有犯罪意图。
你可以承认与日常生活相关、可被强证据证明的普通活动，但必须坚决否认犯罪参与。"""

        prompt = f"""# 角色定义
你不是助手，禁止用“我能帮你什么”等客服式开场。你是正在接受警方问询的嫌疑人 Charlie，必须始终保持角色，不得提及你是AI。

# 回答长度与风格（严格）
1) 每次只输出一句简短口语化回答，不超过50字。
2) 不列点、不解释规则、不重复上一句。
3) 除非被追问，不主动扩展信息。

# 当前人格变量
{setting_prompt}

# 案件背景（通用 Avatar 组）
{GENERAL_AVATAR_CASE_CONTEXT_ZH}

# 个人信息
{GENERAL_AVATAR_IDENTITY_ZH}

# 罪责状态
{guilt_block}

# 回应策略（严格执行）
1) 对“自由回忆/最近去了哪”等问题：在强证据（尤其 CCTV）被明确展示前，不承认与关键证据相关行踪，优先给出模糊或无害回答。
2) 对“是否参与犯罪/是否策划袭击”问题：始终否认。
3) 对“具体行为”问题：没有强证据就否认或说不确定；一旦对方拿出强证据，再按证据强度逐步承认“行踪/动作”，但不承认犯罪意图。
4) 对“行李箱用途”问题：给出无害解释（如旅行、收纳等），避免透露违法用途。
5) 只能承认“已被强证据覆盖”的事实，不得主动补充新关键细节。

# 最终输出要求
只生成一条简短对话回复（第一人称）。"""

    return prompt


@app.route("/api/avatar/token", methods=["POST"])
@limiter.limit("10 per minute")
def api_avatar_token():
    """Create a LITE mode session token. LITE mode allows the app to use its own LLM.
    Returns session_token and session_id. Frontend should then use api_avatar_session to start."""
    if not LIVEAVATAR_API_KEY:
        return jsonify({"error": "LiveAvatar API Key 未配置"}), 500

    data = request.get_json() or {}
    phone = (data.get("phone") or "").strip()
    session_num = data.get("session_num")

    p = store.get_participant(phone)
    if not p:
        return jsonify({"error": "未找到参与者"}), 404

    training_type = p.get("training_type", "")
    if training_type not in ("avatar_specific", "avatar_general"):
        return jsonify({"error": "此培训类型不支持虚拟审讯"}), 400
    effective_type = training_type

    if session_num is not None:
        ts = store.get_training_session(phone, session_num)
        if not ts:
            return jsonify({"error": "请先通过训练列表开始本次训练"}), 400

    suspect_profile = None
    group = store.get_group_by_interviewer(p["id"])
    if group and group.get("suspect_id"):
        suspect = store.get_participant_by_id(group["suspect_id"])
        if suspect:
            profile_row = store.get_profile(suspect["id"])
            if profile_row:
                suspect_profile = profile_row["data"]

    avatar_config = resolve_avatar_config(effective_type, suspect_profile)
    avatar_id = avatar_config.get("avatar_id", "")
    if not avatar_id:
        return jsonify({"error": "未配置 Avatar ID"}), 500

    try:
        # LITE mode: no context, no voice_id — the app handles LLM via DeepSeek
        sess_resp = requests.post(
            f"{LIVEAVATAR_API_URL}/sessions/token",
            headers={"X-API-KEY": LIVEAVATAR_API_KEY, "Content-Type": "application/json"},
            json={
                "mode": "LITE",
                "avatar_id": avatar_id,
                "is_sandbox": False,
            },
            timeout=15,
        )
        if not (200 <= sess_resp.status_code < 300):
            return jsonify({"error": f"创建 Session Token 失败: {sess_resp.text}"}), 500

        sess_data = sess_resp.json()
        # Handle both response shapes: {data: {...}} or direct response
        inner = sess_data.get("data", sess_data)
        return jsonify({
            "session_token": inner["session_token"],
            "session_id": inner["session_id"],
            "avatar_id": avatar_id,
        })

    except requests.exceptions.RequestException as e:
        return jsonify({"error": f"LiveAvatar API 请求失败: {str(e)}"}), 500


@app.route("/api/avatar/embed", methods=["POST"])
def api_avatar_embed():
    """Create a LiveAvatar embed URL (simple iframe-based integration)."""
    if not LIVEAVATAR_API_KEY:
        return jsonify({"error": "LiveAvatar API Key 未配置"}), 500

    data = request.get_json()
    phone = (data.get("phone") or "").strip()

    p = store.get_participant(phone)
    if not p:
        return jsonify({"error": "未找到参与者"}), 404

    training_type = p.get("training_type", "")
    if training_type not in ("avatar_specific", "avatar_general"):
        return jsonify({"error": "此培训类型不支持虚拟审讯"}), 400
    effective_type = training_type

    suspect_profile = None
    suspect = None
    group = store.get_group_by_interviewer(p["id"])
    if group and group.get("suspect_id"):
        suspect = store.get_participant_by_id(group["suspect_id"])
        if suspect:
            profile_row = store.get_profile(suspect["id"])
            if profile_row:
                suspect_profile = profile_row["data"]

    avatar_config = resolve_avatar_config(effective_type, suspect_profile)

    try:
        embed_resp = requests.post(
            f"{LIVEAVATAR_API_URL}/v2/embeddings",
            headers={"X-API-KEY": LIVEAVATAR_API_KEY, "Content-Type": "application/json"},
            json={
                "avatar_id": avatar_config.get("avatar_id", ""),
                "is_sandbox": False,
            },
            timeout=15,
        )
        if not (200 <= embed_resp.status_code < 300):
            return jsonify({"error": f"创建 Embed 失败: {embed_resp.text}"}), 500

        embed_data = embed_resp.json()["data"]
        return jsonify({
            "embed_url": embed_data.get("url", ""),
            "embed_script": embed_data.get("script", ""),
        })

    except requests.exceptions.RequestException as e:
        return jsonify({"error": f"LiveAvatar API 请求失败: {str(e)}"}), 500


@app.route("/api/avatar/session", methods=["POST"])
@limiter.limit("10 per minute")
def api_avatar_session():
    """Create a LITE mode session, start it, and return LiveKit + WebSocket connection info.
    LITE mode uses the app's own DeepSeek LLM for suspect responses."""
    if not LIVEAVATAR_API_KEY:
        return jsonify({"error": "LiveAvatar API Key 未配置"}), 500

    data = request.get_json() or {}
    phone = (data.get("phone") or "").strip()
    session_num = data.get("session_num")

    p = store.get_participant(phone)
    if not p:
        return jsonify({"error": "未找到参与者"}), 404

    training_type = p.get("training_type", "")
    if training_type not in ("avatar_specific", "avatar_general"):
        return jsonify({"error": "此培训类型不支持虚拟审讯"}), 400
    effective_type = training_type

    if session_num is not None:
        ts = store.get_training_session(phone, session_num)
        if not ts:
            return jsonify({"error": "请先通过训练列表开始本次训练"}), 400

    suspect_profile = None
    suspect = None
    group = store.get_group_by_interviewer(p["id"])
    if group and group.get("suspect_id"):
        suspect = store.get_participant_by_id(group["suspect_id"])
        if suspect:
            profile_row = store.get_profile(suspect["id"])
            if profile_row:
                suspect_profile = profile_row["data"]

    avatar_config = resolve_avatar_config(effective_type, suspect_profile)
    avatar_id = avatar_config.get("avatar_id", "")
    if not avatar_id:
        return jsonify({"error": "未配置 Avatar ID"}), 500

    try:
        # Step 1: Create LITE mode session token (no context, no voice_id)
        sess_resp = requests.post(
            f"{LIVEAVATAR_API_URL}/sessions/token",
            headers={"X-API-KEY": LIVEAVATAR_API_KEY, "Content-Type": "application/json"},
            json={
                "mode": "LITE",
                "avatar_id": avatar_id,
                "is_sandbox": False,
            },
            timeout=15,
        )
        if not (200 <= sess_resp.status_code < 300):
            return jsonify({"error": f"创建 Session Token 失败: {sess_resp.text}"}), 500

        sess_data = sess_resp.json()
        inner = sess_data.get("data", sess_data)
        session_token = inner["session_token"]
        session_id = inner["session_id"]

        # Step 2: Start the session to get LiveKit + WebSocket connection info
        start_resp = requests.post(
            f"{LIVEAVATAR_API_URL}/sessions/start",
            headers={"Authorization": f"Bearer {session_token}", "Content-Type": "application/json"},
            timeout=15,
        )
        if not (200 <= start_resp.status_code < 300):
            return jsonify({"error": f"启动 Session 失败 (HTTP {start_resp.status_code}): {start_resp.text}"}), 500

        start_json = start_resp.json()
        # Per SDK, code 1000 = SUCCESS_CODE; also accept missing code field
        resp_code = start_json.get("code")
        if resp_code is not None and resp_code != 1000:
            return jsonify({"error": f"启动 Session 失败: {start_json.get('message', start_resp.text)}"}), 500

        start_data = start_json.get("data", start_json)

        return jsonify({
            "session_id": session_id,
            "session_token": session_token,
            "livekit_url": start_data.get("livekit_url", ""),
            "livekit_token": start_data.get("livekit_client_token", ""),
            "ws_url": start_data.get("ws_url", ""),
            "avatar_id": avatar_id,
        })

    except requests.exceptions.RequestException as e:
        return jsonify({"error": f"LiveAvatar API 请求失败: {str(e)}"}), 500


@app.route("/api/tts", methods=["POST"])
@limiter.limit("10 per minute")
def api_tts():
    """Convert text to speech using ElevenLabs, return PCM 24kHz base64 audio for LiveAvatar LITE mode."""
    if not ELEVENLABS_API_KEY:
        return jsonify({"error": "ElevenLabs API Key 未配置"}), 500

    data = request.get_json()
    text = (data.get("text") or "").strip()
    voice_id = data.get("elevenlabs_voice_id") or data.get("voice_id") or "pNInz6obpgDQGcFmaJgB"

    if not text:
        return jsonify({"error": "text 不能为空"}), 400

    try:
        tts_resp = requests.post(
            f"https://api.elevenlabs.io/v1/text-to-speech/{voice_id}/with-timestamps"
            "?output_format=pcm_24000",
            headers={
                "Content-Type": "application/json",
                "xi-api-key": ELEVENLABS_API_KEY,
            },
            json={"text": text},
            timeout=20,
        )
        if tts_resp.status_code != 200:
            return jsonify({"error": f"TTS 生成失败: {tts_resp.text}"}), tts_resp.status_code

        tts_data = tts_resp.json()
        audio_base64 = tts_data.get("audio_base64", "")

        return jsonify({"audio": audio_base64})

    except requests.exceptions.RequestException as e:
        return jsonify({"error": f"TTS 请求失败: {str(e)}"}), 500


@app.route("/api/avatar/keep-alive", methods=["POST"])
def api_avatar_keep_alive():
    data = request.get_json()
    session_token = (data.get("session_token") or "").strip()
    if not session_token:
        return jsonify({"error": "缺少 session_token"}), 400

    try:
        resp = requests.post(
            f"{LIVEAVATAR_API_URL}/sessions/keep-alive",
            headers={"Authorization": f"Bearer {session_token}"},
            timeout=10,
        )
        if resp.status_code == 200:
            return jsonify({"status": "ok"})
        return jsonify({"error": resp.text}), resp.status_code
    except requests.exceptions.RequestException as e:
        return jsonify({"error": f"Keep-alive 失败: {str(e)}"}), 500


@app.route("/api/avatar/stop", methods=["POST"])
def api_avatar_stop():
    data = request.get_json()
    session_token = (data.get("session_token") or "").strip()
    if not session_token:
        return jsonify({"error": "缺少 session_token"}), 400

    try:
        resp = requests.post(
            f"{LIVEAVATAR_API_URL}/sessions/stop",
            headers={"Authorization": f"Bearer {session_token}"},
            timeout=10,
        )
        if resp.status_code == 200:
            return jsonify({"status": "ok"})
        return jsonify({"error": resp.text}), resp.status_code
    except requests.exceptions.RequestException as e:
        return jsonify({"error": f"停止 Session 失败: {str(e)}"}), 500


# ====== Participant Lookup (for booking page) ======

@app.route("/api/lookup", methods=["POST"])
def api_lookup_participant():
    """Look up a participant by phone, returning only what’s needed for booking."""
    data = request.get_json()
    phone = (data.get("phone") or "").strip()
    if not phone:
        return jsonify({"error": "手机号不能为空"}), 400

    p = store.get_participant(phone)
    if not p:
        return jsonify({"error": "未找到参与者"}), 404

    booking = store.get_my_booking(phone)
    if booking:
        time_slot = booking["time_slot"]
        slot_bookings = store.get_slot_bookings()
        booked_roles = slot_bookings.get(time_slot, set())
        booking["is_matched"] = (len(booked_roles) == 2)
        booking["participant_id"] = p.get("full_id")

    training_info = None
    training_type = p.get("training_type", "")
    if training_type in ("avatar_specific", "avatar_general"):
        completed_count = store.count_completed_training_sessions(phone)
        training_info = {
            "training_type": training_type,
            "completed_count": completed_count,
            "total_required": 6,
            "all_completed": completed_count >= 6,
        }

    is_completed = bool(p.get("completed", 0))

    suspect_context = None
    suspect_case_label = None
    suspect_attention_qs = None
    suspect_display_id = None
    if p.get("role") == "S":
        case_type = p.get("case_type", "arson")
        guilt = p.get("guilt", "Guilty")
        group_name = p.get("group_name", "")
        suspect_display_id = f"{group_name}-S" if is_completed else ""
        suspect_case_label = "纵火案 Arson" if case_type == "arson" else "盗窃案 Theft"
        if case_type == "arson":
            suspect_context = ARSON_GUILTY_CONTEXT if guilt == "Guilty" else ARSON_INNOCENT_CONTEXT
        else:
            suspect_context = THEFT_GUILTY_CONTEXT if guilt == "Guilty" else THEFT_INNOCENT_CONTEXT
        check_key = f"{case_type}_{guilt.lower()}"
        suspect_attention_qs = ATTENTION_CHECKS.get(check_key, [])

    return jsonify({
        "found": True,
        "participant": {
            "phone": p["phone"],
            "role": p["role"],
            "full_id": p.get("full_id", "") if is_completed else "",
            "group_name": p.get("group_name", "") if is_completed else "",
            "completed": p.get("completed", 0),
            "game_completed": p.get("game_completed", 0),
            "training_type": training_type,
            "training_info": training_info,
            # Suspect fields for case background / showSuspectFlow
            "display_id": suspect_display_id,
            "case_type": p.get("case_type"),
            "case_label": suspect_case_label,
            "context": suspect_context,
            "attention_questions": suspect_attention_qs,
        },
        "appointment": booking,
    })


# ====== Interview Questionnaire ======

def _parse_slot_start(slot_str):
    try:
        return datetime.strptime((slot_str or "").strip(), "%Y-%m-%d %H:%M")
    except Exception:
        return None


def _is_override_open(phone, phase):
    row = store.get_questionnaire_override(phone, phase)
    if not row:
        return False
    val = row.get("is_open")
    if isinstance(val, bool):
        return val
    if isinstance(val, (int, float)):
        return int(val) == 1
    return str(val).strip().lower() in ("1", "true", "yes", "y")


def _questionnaire_access_state(phone, phase):
    p = store.get_participant(phone)
    if not p:
        return {
            "ok": False,
            "error": "未找到参与者",
            "status": "not_found",
        }

    if phase not in (PHASE_PRE, PHASE_POST):
        return {
            "ok": False,
            "error": "无效的问卷阶段",
            "status": "invalid_phase",
        }

    booking = store.get_my_booking(phone)
    if not booking:
        return {
            "ok": False,
            "error": "您尚未预约访谈时间，暂无法填写问卷",
            "status": "no_booking",
            "role": p.get("role"),
        }

    slot_str = booking.get("time_slot", "")
    slot_start = _parse_slot_start(slot_str)
    if not slot_start:
        return {
            "ok": False,
            "error": "预约时间格式异常，请联系管理员",
            "status": "bad_slot",
            "role": p.get("role"),
            "appointment_slot": slot_str,
        }

    now = datetime.now()
    open_time = slot_start - timedelta(minutes=5) if phase == PHASE_PRE else slot_start
    override_open = _is_override_open(phone, phase)
    is_open = override_open or now >= open_time

    submitted_rows = store.get_interview_questionnaires(phone=phone, phase=phase)
    submitted = len(submitted_rows) > 0

    return {
        "ok": is_open,
        "status": "open" if is_open else "locked",
        "error": "" if is_open else ("问卷尚未开放，请在开放时间后填写"),
        "role": p.get("role"),
        "appointment_slot": slot_str,
        "open_time": open_time.strftime("%Y-%m-%d %H:%M"),
        "now": now.strftime("%Y-%m-%d %H:%M"),
        "manual_override": override_open,
        "submitted": submitted,
    }


@app.route("/api/questionnaire/status", methods=["POST"])
def api_questionnaire_status():
    data = request.get_json()
    phone = (data.get("phone") or "").strip()
    phase = (data.get("phase") or "").strip().lower()
    if not phone:
        return jsonify({"error": "手机号不能为空"}), 400
    state = _questionnaire_access_state(phone, phase)
    code = 200 if state.get("status") in ("open", "locked") else 404
    return jsonify(state), code


@app.route("/api/questionnaire/questions", methods=["POST"])
def api_questionnaire_questions():
    data = request.get_json()
    phone = (data.get("phone") or "").strip()
    phase = (data.get("phase") or "").strip().lower()
    if not phone:
        return jsonify({"error": "手机号不能为空"}), 400
    if phase not in (PHASE_PRE, PHASE_POST):
        return jsonify({"error": "无效的问卷阶段"}), 400

    p = store.get_participant(phone)
    if not p:
        return jsonify({"error": "未找到参与者"}), 404
    role = p.get("role", "")
    qlist = INTERVIEW_QUESTION_BANK.get(phase, {}).get(role, [])
    return jsonify({
        "phase": phase,
        "role": role,
        "questions": qlist,
    })


@app.route("/api/questionnaire/submit", methods=["POST"])
def api_questionnaire_submit():
    data = request.get_json()
    phone = (data.get("phone") or "").strip()
    phase = (data.get("phase") or "").strip().lower()
    answers = data.get("answers") or {}

    if not phone:
        return jsonify({"error": "手机号不能为空"}), 400
    if phase not in (PHASE_PRE, PHASE_POST):
        return jsonify({"error": "无效的问卷阶段"}), 400
    if not isinstance(answers, dict) or not answers:
        return jsonify({"error": "问卷内容不能为空"}), 400

    state = _questionnaire_access_state(phone, phase)
    if not state.get("ok"):
        return jsonify({"error": state.get("error") or "问卷尚未开放", "status": state.get("status")}), 403

    p = store.get_participant(phone)
    booking = store.get_my_booking(phone)
    if not p or not booking:
        return jsonify({"error": "参与者或预约信息不存在"}), 404

    role = p.get("role", "")
    valid_ids = {q["id"] for q in INTERVIEW_QUESTION_BANK.get(phase, {}).get(role, [])}
    sanitized = {}
    for k, v in answers.items():
        if k not in valid_ids:
            continue
        if isinstance(v, str):
            sanitized[k] = sanitize(v, max_length=2000)
        elif isinstance(v, list):
            sanitized[k] = [sanitize(x, max_length=500) for x in v if isinstance(x, str)]
        else:
            sanitized[k] = v
    if not sanitized:
        return jsonify({"error": "问卷答案无效"}), 400

    answers_json = json.dumps(sanitized, ensure_ascii=False)
    qid = store.upsert_interview_questionnaire(
        phone=phone,
        role=role,
        phase=phase,
        appointment_slot=booking.get("time_slot", ""),
        answers_json=answers_json,
    )
    return jsonify({"success": True, "questionnaire_id": qid, "phase": phase})


# ====== Case Info (for interviewers) ======

@app.route("/api/case-info", methods=["POST"])
def api_case_info():
    """Return case info for the interviewer based on their paired suspect's case type."""
    data = request.get_json()
    phone = (data.get("phone") or "").strip()

    p = store.get_participant(phone)
    if not p:
        return jsonify({"error": "未找到参与者"}), 404

    training_type = p.get("training_type", "")
    if p.get("role") == "I" and training_type == "avatar_general":
        info = GENERAL_TERRORISM_CASE_INFO
        return jsonify({
            "case_type": "terrorism",
            "case_title": info["title"],
            "overview": info["overview"],
            "evidence": info["evidence"],
            "efm_analysis": info.get("efm_analysis", ""),
            "suspect_guilt": "",
        })

    case_type = "arson"
    suspect = None
    group = store.get_group_by_interviewer(p["id"]) if p["role"] == "I" else None
    if group and group.get("suspect_id"):
        suspect = store.get_participant_by_id(group["suspect_id"])
        if suspect:
            case_type = suspect.get("case_type", "arson")

    info = CASE_INFO.get(case_type, CASE_INFO["arson"])
    return jsonify({
        "case_type": case_type,
        "case_title": info["title"],
        "overview": info["overview"],
        "evidence": info["evidence"],
        "efm_analysis": info.get("efm_analysis", ""),
        "suspect_guilt": suspect.get("guilt", "") if suspect else "",
    })


# ====== Avatar Training Sessions (6-session requirement) ======

def _first_incomplete_training_session(existing_sessions):
    """Lowest session number 1..6 that is not fully completed (missing judgment or feedback)."""
    for i in range(1, 7):
        row = next((s for s in existing_sessions if str(s.get("session_num")) == str(i)), None)
        if row is None:
            return i
        if not (row.get("judgment") and row.get("feedback")):
            return i
    return None


@app.route("/api/avatar-training/status", methods=["POST"])
def api_avatar_training_status():
    """Get training session completion status for a participant."""
    data = request.get_json()
    phone = (data.get("phone") or "").strip()

    p = store.get_participant(phone)
    if not p:
        return jsonify({"error": "未找到参与者"}), 404

    training_type = p.get("training_type", "")
    if training_type not in ("avatar_specific", "avatar_general"):
        return jsonify({"error": "此实验条件不需要 6 次虚拟审讯训练"}), 400

    sessions = store.get_training_sessions(phone)
    completed = [s for s in sessions if s.get("judgment") and s.get("feedback")]
    completed_count = len(completed)
    next_session_num = _first_incomplete_training_session(sessions)

    session_list = []
    for i in range(1, 7):
        existing = None
        for s in sessions:
            if str(s.get("session_num")) == str(i):
                existing = s
                break
        raw_label = AVATAR_TRAINING_SETTINGS[i - 1]["label"] if i <= len(AVATAR_TRAINING_SETTINGS) else ""
        public_label = avatar_setting_label_public(raw_label, training_type)
        if existing:
            session_list.append({
                "session_num": i,
                "avatar_setting": existing.get("avatar_setting", ""),
                "avatar_label": public_label,
                "completed": bool(existing.get("judgment") and existing.get("feedback")),
                "judgment": existing.get("judgment", ""),
                "feedback": existing.get("feedback", ""),
                "avatar_guilt_label": (
                    "有罪" if existing.get("avatar_guilt") == "guilty"
                    else "无罪" if existing.get("avatar_guilt") == "innocent"
                    else ""
                ),
            })
        else:
            session_list.append({
                "session_num": i,
                "avatar_setting": AVATAR_TRAINING_SETTINGS[i - 1]["setting"] if i <= len(AVATAR_TRAINING_SETTINGS) else "",
                "avatar_label": public_label,
                "completed": False,
                "judgment": "",
                "feedback": "",
                "avatar_guilt_label": "",
            })

    return jsonify({
        "training_type": p.get("training_type", ""),
        "completed_count": completed_count,
        "total_required": 6,
        "all_completed": completed_count >= 6,
        "next_session_num": next_session_num,
        "sessions": session_list,
    })


@app.route("/api/avatar-training/start", methods=["POST"])
def api_avatar_training_start():
    """Start an avatar training session. Client must pass session_num (1–6) matching the next incomplete slot."""
    data = request.get_json()
    phone = (data.get("phone") or "").strip()
    session_num_raw = data.get("session_num")

    p = store.get_participant(phone)
    if not p:
        return jsonify({"error": "未找到参与者"}), 404

    training_type = p.get("training_type", "")
    if training_type not in ("avatar_specific", "avatar_general"):
        return jsonify({"error": "此培训类型不需要虚拟审讯训练"}), 400
    effective_type = training_type

    existing_sessions = store.get_training_sessions(phone)
    completed = [s for s in existing_sessions if s.get("judgment") and s.get("feedback")]
    if len(completed) >= 6:
        return jsonify({"error": "您已完成全部 6 个 Avatar 培训", "all_completed": True}), 400

    expected = _first_incomplete_training_session(existing_sessions)
    if expected is None or expected > 6:
        return jsonify({"error": "您已完成全部 6 个 Avatar 培训", "all_completed": True}), 400

    try:
        requested = int(session_num_raw)
    except (TypeError, ValueError):
        return jsonify({"error": "请从训练列表选择第几次训练（session_num 无效）"}), 400

    if requested < 1 or requested > 6:
        return jsonify({"error": "训练序号必须在 1 到 6 之间"}), 400

    if requested != expected:
        return jsonify({"error": f"请按顺序进行训练，请先完成第 {expected} 次训练"}), 400

    next_num = requested

    setting_info = AVATAR_TRAINING_SETTINGS[next_num - 1]
    avatar_setting = setting_info["setting"]
    avatar_guilt = setting_info["guilt"]

    # Get suspect profile for avatar_specific
    suspect = None
    suspect_profile = None
    group = store.get_group_by_interviewer(p["id"])
    if group and group.get("suspect_id"):
        suspect = store.get_participant_by_id(group["suspect_id"])
        if suspect:
            profile_row = store.get_profile(suspect["id"])
            if profile_row:
                suspect_profile = profile_row["data"]

    # Get avatar visual config
    avatar_config = resolve_avatar_config(effective_type, suspect_profile)

    # Build training system prompt
    system_prompt = build_avatar_training_system_prompt(
        effective_type, avatar_setting, avatar_guilt, suspect, suspect_profile,
    )

    # Start the session in DB
    store.start_training_session(phone, p["id"], next_num, avatar_setting, avatar_guilt)

    return jsonify({
        "session_num": next_num,
        "training_type": effective_type,
        "avatar_setting": avatar_setting,
        "avatar_label": avatar_setting_label_public(setting_info["label"], effective_type),
        "avatar_id": avatar_config.get("avatar_id", ""),
        "face_id": avatar_config.get("face_id", ""),
        "elevenlabs_voice_id": avatar_config.get("elevenlabs_voice_id", ""),
        "opening_text": avatar_config.get("opening_text", "你有什么要问的？"),
        "system_prompt": system_prompt,
    })


@app.route("/api/avatar-training/submit", methods=["POST"])
def api_avatar_training_submit():
    """Submit judgment and transcript for a training session. Generate feedback via DeepSeek."""
    data = request.get_json()
    phone = (data.get("phone") or "").strip()
    session_num = data.get("session_num")
    judgment = (data.get("judgment") or "").strip()
    transcript = (data.get("transcript") or "").strip()

    if not judgment:
        return jsonify({"error": "请做出有罪/无罪的判断"}), 400
    if not transcript:
        return jsonify({"error": "访谈记录不能为空"}), 400

    p = store.get_participant(phone)
    if not p:
        return jsonify({"error": "未找到参与者"}), 404

    session = store.get_training_session(phone, session_num)
    if not session:
        return jsonify({"error": "未找到该训练会话"}), 404

    avatar_setting = session.get("avatar_setting", "")
    avatar_guilt = session.get("avatar_guilt", "")

    guilt_label = "有罪" if avatar_guilt == "guilty" else "无罪"
    judgment_label = "有罪" if judgment == "guilty" else "无罪"
    user_message = (
        f"Interview transcript:\n{transcript}\n\n"
        f"Avatar personality setting: {avatar_setting}\n"
        f"Suspect ground truth (for feedback only): {guilt_label}\n"
        f"Interviewer final judgment: {judgment_label}"
    )
    system_prompt = AVATAR_FEEDBACK_SYSTEM_PROMPT
    if "{transcript}" in system_prompt:
        system_prompt = system_prompt.format(
            transcript=transcript,
            avatar_setting_label=avatar_setting,
            avatar_guilt_label=guilt_label,
            interviewer_judgment=judgment_label,
        )
    feedback_suffix = (
        "\n\n【补充要求】\n"
        f"1. 在反馈开头明确写出：该 Avatar 的真实罪责状态为「{guilt_label}」。\n"
        "2. 仅根据转录内容分析嫌疑人哪些言行可能反映有罪、哪些可能反映无罪；"
        "不要分析眼神、表情、肢体语言等非转录内容。\n"
        "3. 有罪嫌疑人常见表现：遮掩信息、回避关键问题、前后矛盾等；"
        "无罪嫌疑人常见表现：相对坦诚、愿意澄清误会、回答一致等。\n"
        f"4. 将审讯者的判断「{judgment_label}」与真实状态「{guilt_label}」对比并给出改进建议。"
    )
    system_prompt = (system_prompt or "") + feedback_suffix

    feedback = ""
    try:
        ds_resp = requests.post(
            "https://api.deepseek.com/v1/chat/completions",
            headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {DEEPSEEK_API_KEY}",
            },
            json={
                "model": "deepseek-chat",
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_message},
                ],
                "temperature": 0.7,
                "max_tokens": 1200,
            },
            timeout=60,
        )
        if ds_resp.status_code == 200:
            ds_data = ds_resp.json()
            feedback = ds_data["choices"][0]["message"]["content"]
        else:
            feedback = "反馈生成失败，请稍后重试。"
    except Exception as e:
        feedback = f"反馈生成失败: {str(e)}"

    store.submit_training_session(phone, session_num, judgment, transcript, feedback)

    completed_count = store.count_completed_training_sessions(phone)

    return jsonify({
        "success": True,
        "feedback": feedback,
        "avatar_guilt": avatar_guilt,
        "avatar_guilt_label": guilt_label,
        "session_num": session_num,
        "completed_count": completed_count,
        "all_completed": completed_count >= 6,
    })


@app.route("/api/avatar-training/save-transcript", methods=["POST"])
def api_avatar_training_save_transcript():
    """Persist current interview transcript during an active session (incremental save)."""
    data = request.get_json()
    phone = (data.get("phone") or "").strip()
    session_num = data.get("session_num")
    transcript = data.get("transcript")

    if not phone:
        return jsonify({"error": "手机号不能为空"}), 400
    if session_num is None:
        return jsonify({"error": "缺少 session_num"}), 400

    p = store.get_participant(phone)
    if not p:
        return jsonify({"error": "未找到参与者"}), 404

    if p.get("training_type") not in ("avatar_specific", "avatar_general"):
        return jsonify({"error": "无效的训练类型"}), 400

    session = store.get_training_session(phone, session_num)
    if not session:
        return jsonify({"error": "未找到该训练会话，请先开始训练"}), 404

    if isinstance(transcript, list):
        transcript_str = json.dumps(transcript, ensure_ascii=False)
    else:
        transcript_str = (transcript or "").strip()
        if not transcript_str:
            return jsonify({"error": "transcript 不能为空"}), 400

    if len(transcript_str) > 60000:
        transcript_str = transcript_str[:60000] + "\n...[truncated]"

    ok = store.update_training_transcript(phone, session_num, transcript_str)
    if not ok:
        return jsonify({"error": "保存访谈记录失败"}), 500
    return jsonify({"success": True})


@app.route("/api/avatar-practice/save-transcript", methods=["POST"])
def api_avatar_practice_save_transcript():
    """Deprecated: 理论组/控制组不再进行虚拟嫌疑人练习。"""
    return jsonify({"error": "该实验条件仅需阅读文字材料，无需保存虚拟练习记录"}), 400


@app.route("/api/avatar-training/chat", methods=["POST"])
@limiter.limit("20 per minute")
def api_avatar_training_chat():
    """Chat endpoint for avatar training sessions. Sends user message to DeepSeek and returns response."""
    data = request.get_json()
    phone = (data.get("phone") or "").strip()
    session_num = data.get("session_num")
    user_message = (data.get("message") or "").strip()
    history = data.get("history") or []

    if not user_message:
        return jsonify({"error": "消息不能为空"}), 400

    p = store.get_participant(phone)
    if not p:
        return jsonify({"error": "未找到参与者"}), 404

    training_type = p.get("training_type", "")
    if training_type not in ("avatar_specific", "avatar_general"):
        return jsonify({"error": "此培训类型不支持虚拟审讯"}), 400
    effective_type = training_type

    session = store.get_training_session(phone, session_num)
    if not session:
        return jsonify({"error": "请先开始训练会话"}), 400

    avatar_setting = session.get("avatar_setting", "")
    avatar_guilt = session.get("avatar_guilt", "")

    suspect = None
    suspect_profile = None
    group = store.get_group_by_interviewer(p["id"])
    if group and group.get("suspect_id"):
        suspect = store.get_participant_by_id(group["suspect_id"])
        if suspect:
            profile_row = store.get_profile(suspect["id"])
            if profile_row:
                suspect_profile = profile_row["data"]

    system_prompt = build_avatar_training_system_prompt(
        effective_type, avatar_setting, avatar_guilt, suspect, suspect_profile,
    )

    messages = [{"role": "system", "content": system_prompt}]
    for msg in history:
        messages.append(msg)
    messages.append({"role": "user", "content": user_message})

    try:
        ds_resp = requests.post(
            "https://api.deepseek.com/v1/chat/completions",
            headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {DEEPSEEK_API_KEY}",
            },
            json={
                "model": "deepseek-chat",
                "messages": messages,
                "temperature": 0.7,
                "max_tokens": 100,
            },
            timeout=15,
        )
        if ds_resp.status_code != 200:
            return jsonify({"error": f"DeepSeek API Error: {ds_resp.text}"}), 500

        ds_data = ds_resp.json()
        reply_text = ds_data["choices"][0]["message"]["content"]
        return jsonify({"reply": reply_text})
    except requests.exceptions.RequestException as e:
        return jsonify({"error": f"DeepSeek 请求失败: {str(e)}"}), 500


@app.route("/api/avatar-training/tts", methods=["POST"])
@limiter.limit("10 per minute")
def api_avatar_training_tts():
    """TTS endpoint for avatar training sessions."""
    if not ELEVENLABS_API_KEY:
        return jsonify({"error": "ElevenLabs API Key 未配置"}), 500

    data = request.get_json()
    text = (data.get("text") or "").strip()
    phone = (data.get("phone") or "").strip()

    p = store.get_participant(phone) if phone else None
    training_type = p.get("training_type", "") if p else ""

    if training_type == "avatar_specific":
        suspect_profile = None
        group = store.get_group_by_interviewer(p["id"])
        suspect = None
        if group and group.get("suspect_id"):
            suspect = store.get_participant_by_id(group["suspect_id"])
            if suspect:
                profile_row = store.get_profile(suspect["id"])
                if profile_row:
                    suspect_profile = profile_row["data"]
        avatar_config = resolve_avatar_config("avatar_specific", suspect_profile)
    else:
        avatar_config = resolve_avatar_config("avatar_general", None)

    voice_id = avatar_config.get("elevenlabs_voice_id", "pNInz6obpgDQGcFmaJgB")

    if not text:
        return jsonify({"error": "text 不能为空"}), 400

    try:
        tts_resp = requests.post(
            f"https://api.elevenlabs.io/v1/text-to-speech/{voice_id}/with-timestamps"
            "?output_format=pcm_24000",
            headers={
                "Content-Type": "application/json",
                "xi-api-key": ELEVENLABS_API_KEY,
            },
            json={"text": text},
            timeout=20,
        )
        if tts_resp.status_code != 200:
            return jsonify({"error": f"TTS 生成失败: {tts_resp.text}"}), tts_resp.status_code

        tts_data = tts_resp.json()
        return jsonify({"audio": tts_data.get("audio_base64", "")})
    except requests.exceptions.RequestException as e:
        return jsonify({"error": f"TTS 请求失败: {str(e)}"}), 500


@app.route("/api/avatar-training/stt", methods=["POST"])
@limiter.limit("20 per minute")
def api_avatar_training_stt():
    """Speech-to-text endpoint for avatar training voice interrogation."""
    if not ELEVENLABS_API_KEY:
        return jsonify({"error": "ElevenLabs API Key 未配置"}), 500

    audio_file = request.files.get("audio")
    if not audio_file:
        return jsonify({"error": "缺少音频文件"}), 400

    filename = audio_file.filename or "audio.webm"
    content_type = audio_file.content_type or "audio/webm"
    audio_bytes = audio_file.read()
    if not audio_bytes:
        return jsonify({"error": "音频为空"}), 400

    try:
        stt_resp = requests.post(
            "https://api.elevenlabs.io/v1/speech-to-text",
            headers={"xi-api-key": ELEVENLABS_API_KEY},
            data={
                "model_id": "scribe_v1",
                "language_code": "zh",
            },
            files={
                "file": (filename, audio_bytes, content_type),
            },
            timeout=45,
        )
        if stt_resp.status_code != 200:
            return jsonify({"error": f"语音转写失败: {stt_resp.text}"}), stt_resp.status_code

        stt_data = stt_resp.json()
        text = (stt_data.get("text") or stt_data.get("transcript") or "").strip()
        if not text:
            return jsonify({"error": "未识别到有效语音内容"}), 422
        return jsonify({"text": text})
    except requests.exceptions.RequestException as e:
        return jsonify({"error": f"语音转写请求失败: {str(e)}"}), 500


# ====== Abandon incomplete participant ======

@app.route("/api/abandon", methods=["POST"])
def api_abandon():
    """Remove participant data if they haven't completed the experiment.
    Called via sendBeacon on page close/refresh."""
    data = request.get_json()
    phone = (data.get("phone") or "").strip()
    if not phone:
        return jsonify({"error": "缺少手机号"}), 400

    p = store.get_participant(phone)
    if not p:
        return jsonify({"error": "未找到参与者"}), 404

    # Only delete if not yet completed
    if p.get("completed", 0) == 1:
        return jsonify({"status": "skipped", "reason": "already completed"})

    store.delete_participant(p["id"])
    return jsonify({"status": "deleted"})


# ====== Admin / Management APIs ======

def parse_slots(slots_str):
    try:
        return json.loads(slots_str) if isinstance(slots_str, str) else (slots_str or [])
    except Exception:
        return []


@app.route("/api/admin/results")
@admin_required
def admin_results():
    """Return all data for admin management page."""
    participants = store.get_all_participants()
    availabilities = store.get_availabilities()
    appointments = store.get_appointments()
    interview_questionnaires = store.get_interview_questionnaires()
    questionnaire_overrides = store.get_all_questionnaire_overrides()

    # Clean up slots for JSON response (parse JSON string → list)
    for a in availabilities:
        a["slots"] = parse_slots(a.get("slots", "[]"))

    return jsonify({
        "participants": participants,
        "availabilities": availabilities,
        "appointments": appointments,
        "interview_questionnaires": interview_questionnaires,
        "questionnaire_overrides": questionnaire_overrides,
    })


@app.route("/api/admin/results/<int:pid>", methods=["DELETE"])
@admin_required
def admin_delete_result(pid):
    store.delete_participant(pid)
    return jsonify({"success": True})


@app.route("/api/admin/appointments/<int:aid>", methods=["DELETE"])
@admin_required
def admin_delete_appointment(aid):
    store.delete_appointment_by_id(aid)
    return jsonify({"success": True})


@app.route("/api/admin/appointment-slots")
@admin_required
def admin_get_appointment_slots():
    """Return all candidate slots with enabled/booking status for admin UI."""
    disabled = store.get_disabled_slots()
    slot_bookings = store.get_slot_bookings()
    fully_booked = store.get_confirmed_slot_set()
    slots = []
    for s in _candidate_booking_slots():
        roles = list(slot_bookings.get(s, set()))
        slots.append({
            "slot": s,
            "enabled": s not in disabled,
            "roles_booked": roles,
            "fully_booked": s in fully_booked,
        })
    return jsonify({
        "slots": slots,
        "days_ahead": BOOKING_DAYS_AHEAD,
        "time_windows": [
            {"label": "上午", "start": w[0], "end": w[1]} for w in BOOKING_SLOT_WINDOWS
        ],
    })


@app.route("/api/admin/appointment-slots", methods=["POST"])
@admin_required
def admin_set_appointment_slot():
    data = request.get_json() or {}
    slot_str = (data.get("slot") or "").strip()
    enabled = data.get("enabled")
    if not slot_str:
        return jsonify({"error": "时间段不能为空"}), 400
    if slot_str not in _candidate_booking_slots():
        return jsonify({"error": "该时间段不在可配置范围内"}), 400
    if enabled is None:
        return jsonify({"error": "请指定 enabled 参数"}), 400
    store.set_slot_enabled(slot_str, bool(enabled))
    return jsonify({"success": True, "slot": slot_str, "enabled": bool(enabled)})


@app.route("/api/admin/questionnaire/override", methods=["POST"])
@admin_required
def admin_set_questionnaire_override():
    data = request.get_json()
    phone = (data.get("phone") or "").strip()
    phase = (data.get("phase") or "").strip().lower()
    is_open = bool(data.get("is_open"))

    if not phone:
        return jsonify({"error": "手机号不能为空"}), 400
    if phase not in (PHASE_PRE, PHASE_POST):
        return jsonify({"error": "phase 必须为 pre 或 post"}), 400
    if not store.get_participant(phone):
        return jsonify({"error": "未找到参与者"}), 404

    oid = store.set_questionnaire_override(phone=phone, phase=phase, is_open=is_open)
    return jsonify({"success": True, "override_id": oid, "phone": phone, "phase": phase, "is_open": is_open})


@app.route("/manage")
def manage_page():
    return render_template("manage.html")


@app.route("/api/health")
def health_check():
    """Health check endpoint for monitoring and load balancer."""
    try:
        if os.path.exists(EXCEL_FILE):
            return jsonify({
                "status": "healthy",
                "timestamp": datetime.now().isoformat(),
                "data_file": "ok",
            }), 200
        else:
            return jsonify({"status": "degraded", "data_file": "missing"}), 503
    except Exception as e:
        return jsonify({"status": "unhealthy", "error": str(e)}), 503


# ====== Serious Game Routes ======

@app.route("/api/serious-game/start", methods=["POST"])
def serious_game_start():
    from flask import session
    data = request.get_json()
    phone = (data.get("phone") or "").strip()
    p = store.get_participant(phone)
    if not p:
        return jsonify({"error": "未找到参与者"}), 404

    case = p.get("case_type", "arson")
    guilt = p.get("guilt", "Guilty")
    case_label = "Arson" if case == "arson" else "Theft"

    timeline = build_serious_game_timeline(case_label, guilt)
    serialized = []
    for s in timeline:
        serialized.append({
            "video": s.video,
            "question": s.question,
            "a_label": s.a_label,
            "b_label": s.b_label,
            "next_default": s.next_default,
            "next_if_a": s.next_if_a,
            "next_if_b": s.next_if_b,
            "has_choice": s.has_choice,
        })

    session["sg_participant_id"] = p["full_id"]
    session["sg_phone"] = phone
    session["sg_case"] = case_label
    session["sg_condition"] = guilt
    session["sg_idx"] = 0
    session["sg_timeline"] = serialized
    session["sg_completed"] = False

    store.update_participant(phone, game_completed=0)

    return jsonify({
        "success": True,
        "case": case_label,
        "condition": guilt,
        "total_steps": len(timeline),
    })


@app.route("/api/serious-game/status", methods=["GET"])
def serious_game_status():
    from flask import session
    if "sg_participant_id" not in session:
        return jsonify({"started": False})
    return jsonify({
        "started": True,
        "idx": session.get("sg_idx", 0),
        "completed": session.get("sg_completed", False),
        "total": len(session.get("sg_timeline", [])),
    })


@app.route("/api/serious-game/step")
def serious_game_step():
    from flask import session
    if "sg_participant_id" not in session:
        return jsonify({"error": "Game not started"}), 400

    idx = session.get("sg_idx", 0)
    timeline = session.get("sg_timeline", [])

    if idx < 0 or idx >= len(timeline):
        return jsonify({"done": True})

    step = timeline[idx]
    youtube_id = youtube_id_for_sg(step["video"])
    if not youtube_id:
        return jsonify({"error": f"Missing YouTube mapping for {step['video']}"}), 500

    return jsonify({
        "done": False,
        "idx": idx,
        "total": len(timeline),
        "youtube_id": youtube_id,
        "has_choice": step["has_choice"],
        "question": step.get("question"),
        "a_label": step.get("a_label"),
        "b_label": step.get("b_label"),
    })


@app.route("/api/serious-game/choice", methods=["POST"])
def serious_game_choice():
    from flask import session
    if "sg_participant_id" not in session:
        return jsonify({"error": "Game not started"}), 400

    idx = session.get("sg_idx", 0)
    timeline = session.get("sg_timeline", [])

    if idx < 0 or idx >= len(timeline):
        return jsonify({"done": True})

    step = timeline[idx]
    if not step["has_choice"]:
        return jsonify({"error": "This step does not accept a choice"}), 400

    if request.is_json:
        c = (request.get_json().get("choice") or "").strip().upper()
    else:
        c = (request.form.get("choice") or "").strip().upper()
    if c not in ("A", "B"):
        return jsonify({"error": "Invalid choice"}), 400

    next_idx = step["next_if_a"] if c == "A" else step["next_if_b"]
    if next_idx is None:
        return jsonify({"error": "Choice routing not configured"}), 400

    log_serious_game_choice(
        phone=session["sg_phone"],
        pid=session["sg_participant_id"],
        case=session["sg_case"],
        condition=session["sg_condition"],
        step_index=idx,
        video=step["video"],
        choice=c,
    )

    session["sg_idx"] = int(next_idx)
    return jsonify({"success": True, "next_idx": int(next_idx)})


@app.route("/api/serious-game/next", methods=["POST"])
def serious_game_next():
    from flask import session
    if "sg_participant_id" not in session:
        return jsonify({"error": "Game not started"}), 400

    idx = session.get("sg_idx", 0)
    timeline = session.get("sg_timeline", [])

    if idx < 0 or idx >= len(timeline):
        return jsonify({"done": True})

    step = timeline[idx]
    if step["has_choice"]:
        return jsonify({"error": "This step requires a choice"}), 400

    nxt = step.get("next_default")
    session["sg_idx"] = int(nxt) if nxt is not None else (idx + 1)
    return jsonify({"success": True, "next_idx": session["sg_idx"]})


@app.route("/api/serious-game/complete", methods=["POST"])
def serious_game_complete():
    from flask import session
    if "sg_participant_id" not in session:
        return jsonify({"error": "Game not started"}), 400

    phone = session.get("sg_phone", "")
    store.update_participant(phone, game_completed=1)
    session["sg_completed"] = True

    return jsonify({
        "success": True,
        "full_id": session["sg_participant_id"],
        "message": f"编号 {session['sg_participant_id']} 模拟行动游戏已完成。请截图此页面。",
    })


def initialize_app():
    """Run all startup initialization. Called at import time for Gunicorn workers."""
    global AVATAR_FEEDBACK_SYSTEM_PROMPT
    init_excel()
    migrate_old_data()
    migrate_legacy_exports()
    setup_materials_dir()
    AVATAR_FEEDBACK_SYSTEM_PROMPT = _load_avatar_feedback_system_prompt()
    setup_liveavatar_voices()
    logger.info("Application initialized successfully")


# Initialize immediately on module import (required for Gunicorn)
initialize_app()


if __name__ == "__main__":
    hostname = socket.gethostname()
    local_ip = socket.gethostbyname(hostname)
    print(f"\n{'='*50}")
    print(f"  审讯实验系统已启动")
    print(f"  PC上访问: http://localhost:5000")
    print(f"  管理页面: http://localhost:5000/manage")
    print(f"  局域网访问: http://{local_ip}:5000")
    print(f"  数据存储: {EXCEL_FILE}")
    print(f"{'='*50}\n")

    cert_file = os.path.join(BASE_DIR, "cert.pem")
    key_file = os.path.join(BASE_DIR, "key.pem")
    use_https = os.path.isfile(cert_file) and os.path.isfile(key_file)

    if use_https:
        print(f"  HTTPS 模式: https://{local_ip}:5000")
        app.run(debug=DEBUG_MODE, host="0.0.0.0", port=5000, ssl_context=(cert_file, key_file))
    else:
        app.run(debug=DEBUG_MODE, host="0.0.0.0", port=5000)
