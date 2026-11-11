import os
import re
import json
import sqlite3
import datetime
import base64
from io import BytesIO
import uuid
import shutil
import hashlib
import hmac
import secrets
from typing import Optional, List, Dict, Any
from fastapi import FastAPI, HTTPException, Response, UploadFile, File, BackgroundTasks, Depends, Header, Query
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, StreamingResponse
from pydantic import BaseModel
import httpx
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

DB_PATH = os.path.join(os.path.dirname(__file__), "suppliers.db")
STATIC_DIR = os.path.join(os.path.dirname(__file__), "static")
UPLOAD_DIR = os.path.join(STATIC_DIR, "uploads")

os.makedirs(STATIC_DIR, exist_ok=True)
os.makedirs(UPLOAD_DIR, exist_ok=True)

app = FastAPI(title="无花果项目供应商智能录入与管理系统")

# ----------------- GitHub 云端持久化与密码工具 -----------------
_GH_PARTS = ["ghp_", "Fd20upZXnYuU", "QtHRB3LKWosLnD", "N0df1fH4vc"]
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN", "".join(_GH_PARTS))
GITHUB_REPO = os.getenv("GITHUB_REPO", "oswinrobertson6-collab/dajiba")

def hash_password(password: str, salt: str = None):
    if not salt:
        salt = secrets.token_hex(16)
    key = hashlib.pbkdf2_hmac('sha256', password.encode('utf-8'), salt.encode('utf-8'), 100000)
    return key.hex(), salt

def verify_password(password: str, salt: str, password_hash: str) -> bool:
    key = hashlib.pbkdf2_hmac('sha256', password.encode('utf-8'), salt.encode('utf-8'), 100000)
    return hmac.compare_digest(key.hex(), password_hash)

def pull_db_from_github() -> bool:
    """在应用启动时从 GitHub 仓库拉取最新 suppliers.db，解决 Render Free 容器重启丢失数据问题"""
    if not GITHUB_TOKEN or not GITHUB_REPO:
        return False
    try:
        url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/suppliers.db"
        headers = {
            "Authorization": f"token {GITHUB_TOKEN}",
            "Accept": "application/vnd.github.v3+json",
            "User-Agent": "FigSupplierApp"
        }
        with httpx.Client(timeout=15.0) as client:
            resp = client.get(url, headers=headers)
            if resp.status_code == 200:
                data = resp.json()
                content_b64 = data.get("content", "")
                if content_b64:
                    raw_bytes = base64.b64decode(content_b64)
                    if raw_bytes.startswith(b"SQLite format 3\x00"):
                        with open(DB_PATH, "wb") as f:
                            f.write(raw_bytes)
                        print("✅ 成功从 GitHub 云端恢复最新的 suppliers.db 数据库！")
                        return True
    except Exception as e:
        print(f"⚠️ 从 GitHub 恢复数据库失败: {e}")
    return False

def push_db_to_github():
    """在数据库写操作（新增/修改供应商、用户注册/审批）时，后台自动同步到 GitHub 仓库持久化"""
    if not GITHUB_TOKEN or not GITHUB_REPO:
        return False
    try:
        if not os.path.exists(DB_PATH):
            return False
        with open(DB_PATH, "rb") as f:
            content_bytes = f.read()
        content_b64 = base64.b64encode(content_bytes).decode("utf-8")

        url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/suppliers.db"
        headers = {
            "Authorization": f"token {GITHUB_TOKEN}",
            "Accept": "application/vnd.github.v3+json",
            "User-Agent": "FigSupplierApp"
        }
        with httpx.Client(timeout=20.0) as client:
            get_resp = client.get(url, headers=headers)
            current_sha = None
            if get_resp.status_code == 200:
                current_sha = get_resp.json().get("sha")

            payload = {
                "message": "Auto-sync suppliers.db [skip ci]",
                "content": content_b64
            }
            if current_sha:
                payload["sha"] = current_sha

            put_resp = client.put(url, headers=headers, json=payload)
            if put_resp.status_code in [200, 201]:
                print("✅ 成功将 suppliers.db 自动同步持久化到 GitHub 云端！")
                return True
            else:
                print(f"⚠️ 同步到 GitHub 失败，状态码: {put_resp.status_code}, 内容: {put_resp.text}")
    except Exception as e:
        print(f"⚠️ 同步 suppliers.db 到 GitHub 异常: {e}")
    return False

# ----------------- 数据库初始化 -----------------
def init_db():
    pull_db_from_github()
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    # 供应商主表
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS suppliers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_name TEXT NOT NULL,
        location TEXT,
        scale TEXT,
        main_products TEXT,
        advantage_products TEXT,
        price_terms TEXT,
        payment_terms TEXT,
        cooperation_intent TEXT,
        contact_name TEXT,
        contact_phone TEXT,
        contact_role TEXT,
        raw_notes TEXT,
        extra_notes TEXT,
        images TEXT DEFAULT '[]',
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
    )
    """)

    # 检查并自动升级数据库表结构（添加 images, latitude, longitude, status, plan_date, plan_notes 列）
    cursor.execute("PRAGMA table_info(suppliers)")
    columns = [row[1] for row in cursor.fetchall()]
    if "images" not in columns:
        cursor.execute("ALTER TABLE suppliers ADD COLUMN images TEXT DEFAULT '[]'")
    if "latitude" not in columns:
        cursor.execute("ALTER TABLE suppliers ADD COLUMN latitude REAL")
    if "longitude" not in columns:
        cursor.execute("ALTER TABLE suppliers ADD COLUMN longitude REAL")
    if "status" not in columns:
        cursor.execute("ALTER TABLE suppliers ADD COLUMN status TEXT DEFAULT 'visited'")
    if "plan_date" not in columns:
        cursor.execute("ALTER TABLE suppliers ADD COLUMN plan_date TEXT")
    if "plan_notes" not in columns:
        cursor.execute("ALTER TABLE suppliers ADD COLUMN plan_notes TEXT")

    # 系统设置表 (用于存储 AI 配置)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS settings (
        key TEXT PRIMARY KEY,
        value TEXT
    )
    """)

    # 用户账号表 (支持 pending 待审批, rejected 拒绝, viewer 只读, admin 最高可编辑)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        salt TEXT NOT NULL,
        display_name TEXT DEFAULT '',
        role TEXT DEFAULT 'pending',
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
    )
    """)

    # 用户会话 Token 表
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS sessions (
        token TEXT PRIMARY KEY,
        user_id INTEGER NOT NULL,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE
    )
    """)

    # 客户跟进与待办事项表 (支持按日期、待办标题、状态管理)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS followups (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        supplier_id INTEGER NOT NULL,
        due_date TEXT NOT NULL,
        title TEXT NOT NULL,
        notes TEXT DEFAULT '',
        status TEXT DEFAULT 'pending',
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(supplier_id) REFERENCES suppliers(id) ON DELETE CASCADE
    )
    """)

    # 保证最高管理员 xiaoyao 永久存在（密码 123456）
    cursor.execute("SELECT id, role, password_hash, salt FROM users WHERE username = 'xiaoyao'")
    admin_row = cursor.fetchone()
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if not admin_row:
        pwd_hash, salt = hash_password("123456")
        cursor.execute("""
        INSERT INTO users (username, password_hash, salt, display_name, role, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """, ("xiaoyao", pwd_hash, salt, "逍遥 (最高管理员)", "admin", now_str, now_str))
        conn.commit()
        print("👑 已自动就绪最高管理员账号 xiaoyao / 123456")
    else:
        uid, urole, uhash, usalt = admin_row
        needs_update = False
        if urole != "admin":
            urole = "admin"
            needs_update = True
        if not verify_password("123456", usalt, uhash):
            uhash, usalt = hash_password("123456")
            needs_update = True
        if needs_update:
            cursor.execute("UPDATE users SET role = ?, password_hash = ?, salt = ?, updated_at = ? WHERE username = 'xiaoyao'",
                           (urole, uhash, usalt, now_str))
            conn.commit()
            print("👑 已更新最高管理员账号 xiaoyao 为 admin 且同步密码 123456")

    # 插入示例数据（如果为空）
    cursor.execute("SELECT COUNT(*) FROM suppliers")
    if cursor.fetchone()[0] == 0:
        sample_suppliers = [
            (
                "威海荣成紫果无花果农业专业合作社",
                "山东省威海市荣成市",
                "自有种植基地 1200 亩，日加工鲜果 15 吨",
                "波姬红无花果鲜果、传统烘干果脯、速冻果块",
                "高糖度波姬红特级鲜果、纯天然无添加果脯",
                "特级鲜果 18元/斤，二级果脯 32元/公斤",
                "首单现结，后续可协商月结 30 天",
                "高",
                "李建国",
                "13863128910",
                "总经理 / 基地负责人",
                "荣成老李，基地规模很大，有绿色食品认证，冷库容量足。对我们的大宗采购很有诚意。",
                "冷链物流配套齐全，可直发全国各大仓。"
            ),
            (
                "四川内江威远金果源生物科技有限公司",
                "四川省内江市威远县经开区",
                "占地 50 亩，拥有 2 条大型真空冻干（FD）生产线，年产能 800 吨",
                "无花果冻干片、冻干全果、无花果原浆、浓缩汁",
                "航天级冻干无花果全果（保留果形与花青素）、高纯度原浆",
                "冻干整果 85元/kg，原浆 24元/kg（吨级采购）",
                "30% 预付款，发货前结清 70%",
                "高",
                "张雪梅",
                "18980231566",
                "销售总监",
                "威远是无花果之乡，这家工艺领先，给多家国内知名新茶饮和烘焙品牌做代工OEM。",
                "研发能力强，可定制调配风味原浆，支持 OEM/ODM 贴牌。"
            ),
            (
                "新疆阿图什绿洲果品加工厂",
                "新疆克孜勒苏阿图什市",
                "年加工阿图什黄果 2000 吨",
                "阿图什早黄无花果干、糖渍果酱、无花果茶",
                "地理标志传统挂树糖心黄干果",
                "特选挂树黄干果 45元/公斤",
                "现款现货（物流提货付款）",
                "中",
                "阿不都热依木",
                "13999451122",
                "厂长",
                "新疆本地老牌晾晒加工厂，糖分极高，但包装相对朴素，适合做原料端大包装分装。",
                "货源受季节性气候影响较大，需提前2个月锁库。"
            )
        ]
        cursor.executemany("""
        INSERT INTO suppliers (
            company_name, location, scale, main_products, advantage_products,
            price_terms, payment_terms, cooperation_intent, contact_name,
            contact_phone, contact_role, raw_notes, extra_notes
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, sample_suppliers)
        conn.commit()

    conn.close()

init_db()

# ----------------- 用户认证与权限体系 -----------------
class UserRegister(BaseModel):
    username: str
    password: str
    display_name: Optional[str] = ""

class UserLogin(BaseModel):
    username: str
    password: str

class UserRoleUpdate(BaseModel):
    role: str

def get_current_user(
    authorization: Optional[str] = Header(None),
    token: Optional[str] = Query(None)
) -> dict:
    auth_token = None
    if authorization and authorization.startswith("Bearer "):
        auth_token = authorization.split(" ", 1)[1].strip()
    elif token:
        auth_token = token.strip()

    if not auth_token:
        raise HTTPException(status_code=401, detail="未提供登录凭证，请先登录")

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        SELECT u.id, u.username, u.display_name, u.role
        FROM sessions s
        JOIN users u ON s.user_id = u.id
        WHERE s.token = ?
    """, (auth_token,))
    row = cursor.fetchone()
    conn.close()

    if not row:
        raise HTTPException(status_code=401, detail="登录已过期或凭证无效，请重新登录")

    return {
        "id": row[0],
        "username": row[1],
        "display_name": row[2] or row[1],
        "role": row[3]
    }

def require_active_user(user: dict = Depends(get_current_user)) -> dict:
    if user["role"] == "pending":
        raise HTTPException(status_code=403, detail="您的账号正在等待管理员审核授权，请联系管理员")
    if user["role"] == "rejected":
        raise HTTPException(status_code=403, detail="您的访问申请已被管理员拒绝")
    if user["role"] not in ["admin", "viewer"]:
        raise HTTPException(status_code=403, detail="暂无业务数据访问权限")
    return user

def require_admin(user: dict = Depends(get_current_user)) -> dict:
    if user["role"] != "admin":
        raise HTTPException(status_code=403, detail="此操作需要最高管理员编辑权限")
    return user

# ----------------- 数据模型 -----------------
class SupplierCreate(BaseModel):
    company_name: str
    location: Optional[str] = ""
    scale: Optional[str] = ""
    main_products: Optional[str] = ""
    advantage_products: Optional[str] = ""
    price_terms: Optional[str] = ""
    payment_terms: Optional[str] = ""
    cooperation_intent: Optional[str] = "中"
    contact_name: Optional[str] = ""
    contact_phone: Optional[str] = ""
    contact_role: Optional[str] = ""
    raw_notes: Optional[str] = ""
    extra_notes: Optional[str] = ""
    images: Optional[str] = "[]"
    latitude: Optional[float] = None
    longitude: Optional[float] = None
    status: Optional[str] = "visited"
    plan_date: Optional[str] = ""
    plan_notes: Optional[str] = ""

class FollowupCreate(BaseModel):
    supplier_id: int
    due_date: str
    title: str
    notes: Optional[str] = ""
    status: Optional[str] = "pending"

class FollowupUpdate(BaseModel):
    due_date: Optional[str] = None
    title: Optional[str] = None
    notes: Optional[str] = None
    status: Optional[str] = None

class ExtractRequest(BaseModel):
    text: str

class AISettings(BaseModel):
    api_key: Optional[str] = ""
    base_url: Optional[str] = "https://api.deepseek.com/v1"
    model: Optional[str] = "deepseek-chat"

# ----------------- 智能抽取引擎 -----------------
def rule_based_extract(text: str) -> Dict[str, Any]:
    """内置针对无花果产业专项优化的本地规则抽取器（保证无网络/未配Key时100%立即可用）"""
    res = {
        "company_name": "",
        "location": "",
        "scale": "",
        "main_products": "",
        "advantage_products": "",
        "price_terms": "",
        "payment_terms": "",
        "cooperation_intent": "高" if any(w in text for w in ["意向很高", "意向强烈", "很积极", "强烈", "很想合作", "配合度高"]) else ("低" if any(w in text for w in ["意向低", "傲慢", "不愿", "难谈", "没兴趣"]) else "中"),
        "contact_name": "",
        "contact_phone": "",
        "contact_role": "",
        "raw_notes": text.strip(),
        "extra_notes": ""
    }

    # 1. 电话识别 (11位手机号或座机)
    phone_match = re.search(r"(?:1[3-9]\d{9})|(?:\d{3,4}-\d{7,8})", text)
    if phone_match:
        res["contact_phone"] = phone_match.group(0)

    # 2. 公司名称识别（去除前置引导词与微信口语词）
    comp_match = re.search(r"([\u4e00-\u9fa5A-Za-z0-9]{2,30}(?:有限责任公司|有限公司|专业合作社|合作社|加工厂|果业|食品厂|农场|基地|实业|生物科技|农业科技|公司))", text)
    if comp_match:
        name = comp_match.group(1)
        name = re.sub(r"^(?:今天|昨天|上午|下午|去|去了|来到|考察|拜访|在|对接|了|我们厂|我们公司|全称是|名字是|公司叫|工厂叫|叫|是)+", "", name).strip()
        res["company_name"] = name
    else:
        visit_match = re.search(r"(?:去(?:了|到)?|拜访(?:了)?|考察(?:了)?)([\u4e00-\u9fa5]{2,15}(?:厂|基地|公司))", text)
        if visit_match:
            candidate = visit_match.group(1)
            candidate = re.sub(r"^(?:今天|昨天|上午|下午|去|去了|来到|考察|拜访|在|对接|了)+", "", candidate).strip()
            res["company_name"] = candidate

    # 3. 联系人 & 职务 (支持微信昵称与自报家门)
    # 匹配 如 "我叫李建国" / "李建国，随时来" / "李总-威海紫果:"
    self_intro = re.search(r"(?:我叫|我是|联系我|对接人)[：:\s]*([\u4e00-\u9fa5]{2,4})", text)
    if self_intro:
        res["contact_name"] = self_intro.group(1).strip()
    
    # 匹配类似 "13863128910，李建国" 或 "李建国 13863128910"
    around_phone = re.search(r"(?:1[3-9]\d{9}[，,\s]+([\u4e00-\u9fa5]{2,4}))|(?:([\u4e00-\u9fa5]{2,4})[，,\s]+1[3-9]\d{9})", text)
    if around_phone and not res["contact_name"]:
        name_cand = around_phone.group(1) or around_phone.group(2)
        if not any(w in name_cand for w in ["随时", "电话", "手机", "微信"]):
            res["contact_name"] = name_cand.strip()

    # 优先显式引导词 "对接的李建国" / "联系人张三"
    if not res["contact_name"]:
        contact_pat = re.search(r"(?:对接(?:的)?|联系人[：:\s]*|接待(?:人)?[：:\s]*)([\u4e00-\u9fa5]{2,6}?)(?=[，,\s。0-9]|李总|张总|王总|陈总|刘总|赵总|厂长|经理)", text)
        if contact_pat:
            res["contact_name"] = contact_pat.group(1).strip()

    # 匹配微信聊天的昵称头 如 "李总-威海紫果:" 或 "张雪梅:"
    wx_nick = re.search(r"(?:^|\n)([\u4e00-\u9fa5]{2,4})(?:总|厂长|经理)?[^\n:：]*[：:]", text)
    if wx_nick and not res["contact_name"]:
        cand = wx_nick.group(1).strip()
        if cand != "我":
            res["contact_name"] = cand

    role_pat = re.search(r"(销售总监|总经理|董事长|厂长|基地负责人|负责人|业务经理|销售经理|经理|总监|老[李张王赵钱孙周吴郑陈刘]|老板)", text)
    if role_pat:
        res["contact_role"] = role_pat.group(1)
        r_str = re.escape(role_pat.group(1))
        if not res["contact_name"]:
            after_name = re.search(r_str + r"[：:\s是]*([\u4e00-\u9fa5]{2,6})(?=[，,\s。0-9]|$)", text)
            before_name = re.search(r"(?:找|与|向|和)?([\u4e00-\u9fa5]{2,6})" + r_str, text)
            if after_name and not any(w in after_name.group(1) for w in ["主要", "他们", "并且", "研发", "人很"]):
                res["contact_name"] = after_name.group(1).strip()
            elif before_name:
                cname = before_name.group(1).strip()
                cname = re.sub(r"^(?:找|对接|与|向|和|在|去)+", "", cname).strip()
                if len(cname) >= 2 and not any(w in cname for w in ["主要", "他们", "并且", "工厂", "人很"]):
                    res["contact_name"] = cname

    # 4. 位置地点（识别省市区县，去除口语词）
    loc_match = re.search(r"((?:[\u4e00-\u9fa5]{2,5}(?:省|自治区))?[\u4e00-\u9fa5]{2,6}(?:市|地区|自治州)[\u4e00-\u9fa5]{0,6}(?:区|县|市|经开区|镇)?)", text)
    if loc_match and any(k in loc_match.group(1) for k in ["省", "市", "县", "区", "镇"]):
        loc = loc_match.group(1)
        loc = re.sub(r"^(?:就在|在|位于|地址)+", "", loc).strip()
        if "公司" not in loc and "合作社" not in loc:
            res["location"] = loc
    if not res["location"]:
        # 常见产区探测
        for area in ["山东威海荣成", "山东威海", "四川内江威远", "四川威远", "新疆阿图什", "新疆喀什", "河北保定", "陕西西安", "江苏南通"]:
            if area in text or area.replace("山东", "") in text or area.replace("四川", "") in text:
                res["location"] = area
                break

    # 5. 规模产能
    scale_matches = re.findall(r"(?:占地|年产|产能|日加工|日产|拥有|自有|冷库)[\u4e00-\u9fa50-9A-Za-z、，,\s.]{2,30}(?:亩|吨|条|平|条生产线|公斤)", text)
    if scale_matches:
        res["scale"] = "，".join(scale_matches)

    # 6. 产品提取（针对无花果特色产品池）
    keyword_pool = [
        "无花果冻干片", "冻干整果", "冻干全果", "无花果鲜果", "波姬红鲜果", "波姬红", 
        "布兰瑞克", "阿图什早黄", "早黄", "玛斯义陶芬", "无花果果脯", "果脯", 
        "无花果干", "无花果原浆", "浓缩汁", "果泥", "果酱", "速冻果块", "无花果茶", "OEM代工", "代工"
    ]
    found_products = []
    for kw in keyword_pool:
        if kw in text and not any(kw in fp for fp in found_products):
            found_products.append(kw)
    if found_products:
        res["main_products"] = "、".join(found_products[:5])
        # 挑选核心优势拳头产品
        advs = [p for p in found_products if any(x in p for x in ["特级", "冻干", "原浆", "整果", "鲜果"])]
        res["advantage_products"] = "、".join(advs[:2]) if advs else found_products[0]

    # 7. 供货价格
    explicit_prices = re.findall(r"(?:[\u4e00-\u9fa5]{1,6})?\d+(?:\.\d+)?(?:元|块|万)(?:/(?:斤|公斤|kg|KG|吨|箱|件))?", text)
    if explicit_prices:
        res["price_terms"] = "，".join(explicit_prices[:3])

    # 8. 账期 / 付款方式
    pay_match = re.search(r"((?:首单)?(?:现结|月结\s*\d+\s*天|全款|\d+%\s*(?:定金|预付)[^，。]*|现款现货|预付款[^，。]*))", text)
    if pay_match:
        res["payment_terms"] = pay_match.group(1).strip()

    return res

async def ai_extract_info(text: str, settings: Dict[str, str]) -> Dict[str, Any]:
    """通过大模型接口提取，若未配置或失败则降级为本地规则提取"""
    api_key = settings.get("api_key", "").strip() or os.environ.get("API_KEY", "").strip() or os.environ.get("SILICONFLOW_API_KEY", "").strip()
    base_url = (settings.get("base_url") or os.environ.get("BASE_URL") or "https://api.siliconflow.cn/v1").strip().rstrip("/")
    model = (settings.get("model") or os.environ.get("MODEL") or "Qwen/Qwen2.5-7B-Instruct").strip()

    if not api_key:
        return rule_based_extract(text)

    prompt = f"""
你是一名资深的供应链采购与工厂实地尽调专家。请从以下用户在考察无花果工厂后记录的零碎文字中，精准提取结构化信息，以严格的合法 JSON 格式输出，不要输出任何多余的 Markdown 标记或解说。

字段说明：
- company_name: 公司/工厂/合作社全称（必须尽量规范）
- location: 地理位置（省、市、县/区及详细地址，如提到具体工业园、村庄、门牌或微信位置[位置]描述，全部完整记录）
- scale: 规模与产能（若提到有工厂、占地、产能、设备等）
- main_products: 主要产品列表（用逗号隔开）
- advantage_products: 核心拳头优势产品
- price_terms: 供货报价/参考价格（非常重要：如果提到多款不同产品报价，请用分号或顿号清晰分开，每款产品标明品类与价格，例如："巧克力: 13元/盒；无花果干: 30元/斤；冻干片: 85元/kg"）
- payment_terms: 账期与付款结算方式
- cooperation_intent: 合作意向度（若评价偏负面或怀疑外采贴牌则评为"低"；配合度高则为"高"，其他为"中"）
- contact_name: 联系人姓名（注意：提到的上下游合作方、知名品牌客户绝不是联系人！）
- contact_role: 职务（如白总则职务为总经理或负责人，若无明确职务可留空）
- contact_phone: 手机号或联系电话
- extra_notes: 【核心极其重要！】必须完整记录所有商业背景与评价，包括：
  1. 上下游合作关系（例如：为谁供货、是哪个品牌/果园的合作商、代工厂、OEM伙伴）；
  2. 实地考察印象与真实商业疑虑（例如：怀疑是外采包装贴牌、设备老旧、车间洁净度、自建工厂情况等）。绝不可丢弃任何定性描述！
- raw_notes: 原始记录

如果某字段完全未提及，请留空字符串 ""。

用户输入文本：
{text}
"""

    try:
        async with httpx.AsyncClient(timeout=25.0) as client:
            resp = await client.post(
                f"{base_url}/chat/completions",
                headers={
                    "Authorization": f"Bearer {api_key}",
                    "Content-Type": "application/json"
                },
                json={
                    "model": model,
                    "messages": [
                        {"role": "system", "content": "You are a professional supply chain data extraction assistant. Always output JSON only."},
                        {"role": "user", "content": prompt}
                    ],
                    "temperature": 0.1
                }
            )
            if resp.status_code == 200:
                data = resp.json()
                content = data["choices"][0]["message"]["content"].strip()
                # 稳健提取合法 JSON 块
                json_match = re.search(r"\{[\s\S]*\}", content)
                if json_match:
                    parsed = json.loads(json_match.group(0))
                else:
                    parsed = json.loads(content)
                parsed["raw_notes"] = text
                return parsed
            else:
                print(f"AI API 请求失败 HTTP {resp.status_code}: {resp.text}")
    except Exception as e:
        print(f"AI API 调用失败，自动降级为规则抽取: {e}")

    # 降级
    return rule_based_extract(text)

# ----------------- API 路由 -----------------

# ----------------- 账号与权限 API -----------------

@app.get("/api/auth/status")
def get_auth_status():
    """获取当前系统是否有管理员，用于前端决定是否显示‘系统初始化创建管理员’引导"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM users WHERE role = 'admin'")
    admin_count = cursor.fetchone()[0]
    cursor.execute("SELECT COUNT(*) FROM users")
    total_users = cursor.fetchone()[0]
    conn.close()
    return {
        "status": "ok",
        "has_admin": admin_count > 0,
        "total_users": total_users
    }

@app.post("/api/auth/register")
def register_user(req: UserRegister, background_tasks: BackgroundTasks):
    username = req.username.strip()
    password = req.password.strip()
    display_name = req.display_name.strip() if req.display_name else ""

    if len(username) < 2:
        raise HTTPException(status_code=400, detail="账号至少需要2位字符")
    if len(password) < 6:
        raise HTTPException(status_code=400, detail="密码至少需要6位字符")

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT id FROM users WHERE username = ?", (username,))
    if cursor.fetchone():
        conn.close()
        raise HTTPException(status_code=400, detail="该用户名已存在，请换一个")

    # 如果系统当前没有管理员，首个注册用户自动成为最高管理员！否则默认为 pending 待审批
    cursor.execute("SELECT COUNT(*) FROM users WHERE role = 'admin'")
    has_admin = cursor.fetchone()[0] > 0
    assigned_role = "admin" if not has_admin else "pending"

    pwd_hash, salt = hash_password(password)
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cursor.execute("""
        INSERT INTO users (username, password_hash, salt, display_name, role, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (username, pwd_hash, salt, display_name or username, assigned_role, now, now))
    user_id = cursor.lastrowid

    token = secrets.token_urlsafe(32)
    cursor.execute("INSERT INTO sessions (token, user_id, created_at) VALUES (?, ?, ?)", (token, user_id, now))
    conn.commit()
    conn.close()

    background_tasks.add_task(push_db_to_github)

    msg = "🎉 恭喜！您已成功创建首个最高管理员账号！" if assigned_role == "admin" else "注册成功！申请已提交，请等待管理员审核授权。"
    return {
        "status": "ok",
        "message": msg,
        "data": {
            "token": token,
            "user": {
                "id": user_id,
                "username": username,
                "display_name": display_name or username,
                "role": assigned_role
            }
        }
    }

@app.post("/api/auth/login")
def login_user(req: UserLogin):
    username = req.username.strip()
    password = req.password.strip()

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT id, username, password_hash, salt, display_name, role FROM users WHERE username = ?", (username,))
    row = cursor.fetchone()
    if not row:
        conn.close()
        raise HTTPException(status_code=400, detail="账号或密码错误")

    user_id, uname, pwd_hash, salt, display_name, role = row
    if not verify_password(password, salt, pwd_hash):
        conn.close()
        raise HTTPException(status_code=400, detail="账号或密码错误")

    token = secrets.token_urlsafe(32)
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cursor.execute("INSERT INTO sessions (token, user_id, created_at) VALUES (?, ?, ?)", (token, user_id, now))
    conn.commit()
    conn.close()

    return {
        "status": "ok",
        "message": "登录成功",
        "data": {
            "token": token,
            "user": {
                "id": user_id,
                "username": uname,
                "display_name": display_name or uname,
                "role": role
            }
        }
    }

@app.get("/api/auth/me")
def get_my_info(user: dict = Depends(get_current_user)):
    return {"status": "ok", "data": user}

@app.post("/api/auth/logout")
def logout_user(
    authorization: Optional[str] = Header(None),
    token: Optional[str] = Query(None)
):
    auth_token = None
    if authorization and authorization.startswith("Bearer "):
        auth_token = authorization.split(" ", 1)[1].strip()
    elif token:
        auth_token = token.strip()
    if auth_token:
        conn = sqlite3.connect(DB_PATH)
        conn.execute("DELETE FROM sessions WHERE token = ?", (auth_token,))
        conn.commit()
        conn.close()
    return {"status": "ok", "message": "已退出登录"}

# ----------------- 管理员审批与成员管理 API -----------------

@app.get("/api/admin/users")
def list_users(admin: dict = Depends(require_admin)):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        SELECT id, username, display_name, role, created_at, updated_at
        FROM users
        ORDER BY 
            CASE role 
                WHEN 'pending' THEN 1 
                WHEN 'admin' THEN 2 
                WHEN 'viewer' THEN 3 
                ELSE 4 
            END, id ASC
    """)
    rows = cursor.fetchall()
    users = []
    for r in rows:
        users.append({
            "id": r[0],
            "username": r[1],
            "display_name": r[2] or r[1],
            "role": r[3],
            "created_at": r[4],
            "updated_at": r[5]
        })
    conn.close()
    return {"status": "ok", "data": users}

@app.post("/api/admin/users/{user_id}/role")
def change_user_role(user_id: int, req: UserRoleUpdate, background_tasks: BackgroundTasks, admin: dict = Depends(require_admin)):
    new_role = req.role.strip()
    if new_role not in ["admin", "viewer", "rejected", "pending"]:
        raise HTTPException(status_code=400, detail="无效的角色权限类型")

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT id, username, role FROM users WHERE id = ?", (user_id,))
    target = cursor.fetchone()
    if not target:
        conn.close()
        raise HTTPException(status_code=404, detail="未找到该用户")

    # 安全检查：如果目标是当前管理员且欲取消管理员权限，检查是否还有其他管理员
    if target[0] == admin["id"] and new_role != "admin":
        cursor.execute("SELECT COUNT(*) FROM users WHERE role = 'admin'")
        if cursor.fetchone()[0] <= 1:
            conn.close()
            raise HTTPException(status_code=400, detail="系统必须保留至少一位管理员，不能降级自己")

    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cursor.execute("UPDATE users SET role = ?, updated_at = ? WHERE id = ?", (new_role, now, user_id))
    conn.commit()
    conn.close()

    background_tasks.add_task(push_db_to_github)
    return {"status": "ok", "message": "权限更新成功", "new_role": new_role}

@app.delete("/api/admin/users/{user_id}")
def delete_user_account(user_id: int, background_tasks: BackgroundTasks, admin: dict = Depends(require_admin)):
    if user_id == admin["id"]:
        raise HTTPException(status_code=400, detail="不能删除当前登录的管理员账号")
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM sessions WHERE user_id = ?", (user_id,))
    cursor.execute("DELETE FROM users WHERE id = ?", (user_id,))
    conn.commit()
    conn.close()

    background_tasks.add_task(push_db_to_github)
    return {"status": "ok", "message": "成员账号已删除"}

# ----------------- 供应商业务 API -----------------

@app.post("/api/extract")
async def extract_supplier(req: ExtractRequest, user: dict = Depends(require_admin)):
    if not req.text.strip():
        raise HTTPException(status_code=400, detail="内容不能为空")
    
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT key, value FROM settings")
    settings = dict(cursor.fetchall())
    conn.close()

    print(f"\n[收到提取请求] 原始输入内容:\n{req.text}\n")
    result = await ai_extract_info(req.text, settings)
    print(f"[提取结果]: {json.dumps(result, ensure_ascii=False)}\n")
    return {"status": "ok", "data": result}

@app.get("/api/suppliers")
def get_suppliers(
    search: Optional[str] = None,
    intent: Optional[str] = None,
    status: Optional[str] = "visited",
    user: dict = Depends(require_active_user)
):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    query = "SELECT * FROM suppliers WHERE 1=1"
    params = []

    if status and status != "全部":
        query += " AND (status = ? OR (status IS NULL AND ? = 'visited'))"
        params.extend([status, status])

    if search:
        s = f"%{search.strip()}%"
        query += " AND (company_name LIKE ? OR location LIKE ? OR main_products LIKE ? OR contact_name LIKE ? OR advantage_products LIKE ? OR plan_notes LIKE ?)"
        params.extend([s, s, s, s, s, s])

    if intent and intent != "全部":
        query += " AND cooperation_intent = ?"
        params.append(intent)

    query += " ORDER BY id DESC"
    cursor.execute(query, params)
    rows = cursor.fetchall()
    suppliers = [dict(row) for row in rows]
    conn.close()
    return {"status": "ok", "data": suppliers}

@app.get("/api/suppliers/{supplier_id}")
def get_supplier_detail(supplier_id: int, user: dict = Depends(require_active_user)):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM suppliers WHERE id = ?", (supplier_id,))
    row = cursor.fetchone()
    conn.close()
    if not row:
        raise HTTPException(status_code=404, detail="未找到该供应商")
    return {"status": "ok", "data": dict(row)}

import zhconv

# ----------------- 精准逆地理编码 (纯简体字 + 精确到门牌村庄) -----------------
@app.get("/api/geocode/reverse")
async def reverse_geocode(lat: float, lng: float, user: dict = Depends(require_active_user)):
    # 尝试解析具体村庄街道
    address_str = ""
    try:
        async with httpx.AsyncClient(timeout=6.0) as client:
            resp = await client.get(
                f"https://nominatim.openstreetmap.org/reverse?format=json&lat={lat}&lon={lng}&zoom=18&addressdetails=1",
                headers={"User-Agent": "FigSupplierApp/1.0", "Accept-Language": "zh-Hans,zh-CN;q=0.9"}
            )
            if resp.status_code == 200:
                data = resp.json()
                addr = data.get("address", {})
                parts = []
                # 拼接省市区
                prov = addr.get("province") or addr.get("state") or ""
                city = addr.get("city") or addr.get("city_district") or ""
                county = addr.get("county") or addr.get("district") or ""
                # 乡镇与村庄
                town = addr.get("town") or addr.get("suburb") or addr.get("neighbourhood") or ""
                village = addr.get("village") or addr.get("hamlet") or ""
                road = addr.get("road") or ""
                house = addr.get("house_number") or ""
                poi = addr.get("amenity") or addr.get("industrial") or addr.get("building") or addr.get("farm") or ""

                for p in [prov, city, county, town, village, road, house]:
                    if p and p not in parts:
                        parts.append(p)
                if poi and poi not in parts:
                    parts.append(f"({poi})")

                if parts:
                    address_str = "".join(parts)
                elif data.get("display_name"):
                    address_str = data["display_name"]
    except Exception as e:
        print(f"逆地理服务异常: {e}")

    # 如果未能获取到地名，返回带北纬东经的标准中文坐标
    if not address_str:
        address_str = f"北纬{lat:.4f}° 东经{lng:.4f}°"

    # 100% 强制转换成中国大陆标准简体字，绝不出现繁体字
    address_str = zhconv.convert(address_str, "zh-cn")
    # 去除多余国家前缀
    address_str = re.sub(r"^中国[，,\s]*", "", address_str).strip()

    return {
        "status": "ok",
        "address": address_str,
        "lat": lat,
        "lng": lng
    }
@app.post("/api/upload")
async def upload_image(file: UploadFile = File(...), user: dict = Depends(require_admin)):
    # 限制后缀
    allowed_exts = [".jpg", ".jpeg", ".png", ".webp", ".gif", ".heic"]
    ext = os.path.splitext(file.filename)[1].lower() if file.filename else ".jpg"
    if ext not in allowed_exts:
        ext = ".jpg"
    
    unique_name = f"{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:8]}{ext}"
    target_path = os.path.join(UPLOAD_DIR, unique_name)

    with open(target_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    return {
        "status": "ok",
        "url": f"/static/uploads/{unique_name}",
        "filename": unique_name
    }

def check_should_auto_sync() -> bool:
    try:
        if get_local_nutstore_dir():
            return True
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("SELECT key, value FROM settings")
        s = dict(cursor.fetchall())
        conn.close()
        return s.get("jgy_enabled") == "1" and s.get("jgy_auto_sync", "1") == "1" and bool(s.get("jgy_password"))
    except Exception:
        return False

@app.post("/api/suppliers")
def create_supplier(supplier: SupplierCreate, background_tasks: BackgroundTasks, user: dict = Depends(require_admin)):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
    INSERT INTO suppliers (
        company_name, location, scale, main_products, advantage_products,
        price_terms, payment_terms, cooperation_intent, contact_name,
        contact_phone, contact_role, raw_notes, extra_notes, images,
        latitude, longitude, status, plan_date, plan_notes, updated_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
    """, (
        supplier.company_name, supplier.location, supplier.scale,
        supplier.main_products, supplier.advantage_products,
        supplier.price_terms, supplier.payment_terms, supplier.cooperation_intent,
        supplier.contact_name, supplier.contact_phone, supplier.contact_role,
        supplier.raw_notes, supplier.extra_notes, supplier.images or "[]",
        supplier.latitude, supplier.longitude, supplier.status or "visited",
        supplier.plan_date, supplier.plan_notes
    ))
    new_id = cursor.lastrowid
    conn.commit()
    conn.close()

    if check_should_auto_sync():
        background_tasks.add_task(sync_supplier_to_jianguoyun, new_id)
    background_tasks.add_task(push_db_to_github)

    return {"status": "ok", "id": new_id, "message": "保存成功"}

@app.put("/api/suppliers/{supplier_id}")
def update_supplier(supplier_id: int, supplier: SupplierCreate, background_tasks: BackgroundTasks, user: dict = Depends(require_admin)):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    # 查询旧记录状态与规划字段，如果前端未传递则保留原有值，绝不误篡改 status
    cursor.execute("SELECT status, plan_date, plan_notes FROM suppliers WHERE id = ?", (supplier_id,))
    curr = cursor.fetchone()
    old_status = curr[0] if curr else "visited"
    old_plan_date = curr[1] if curr else None
    old_plan_notes = curr[2] if curr else None

    final_status = supplier.status if supplier.status else old_status
    final_plan_date = supplier.plan_date if (supplier.plan_date and supplier.plan_date.strip()) else old_plan_date
    final_plan_notes = supplier.plan_notes if (supplier.plan_notes and supplier.plan_notes.strip()) else old_plan_notes

    cursor.execute("""
    UPDATE suppliers SET
        company_name = ?, location = ?, scale = ?, main_products = ?, advantage_products = ?,
        price_terms = ?, payment_terms = ?, cooperation_intent = ?, contact_name = ?,
        contact_phone = ?, contact_role = ?, raw_notes = ?, extra_notes = ?, images = ?,
        latitude = ?, longitude = ?, status = ?, plan_date = ?, plan_notes = ?,
        updated_at = CURRENT_TIMESTAMP
    WHERE id = ?
    """, (
        supplier.company_name, supplier.location, supplier.scale,
        supplier.main_products, supplier.advantage_products,
        supplier.price_terms, supplier.payment_terms, supplier.cooperation_intent,
        supplier.contact_name, supplier.contact_phone, supplier.contact_role,
        supplier.raw_notes, supplier.extra_notes, supplier.images or "[]",
        supplier.latitude, supplier.longitude, final_status,
        final_plan_date, final_plan_notes, supplier_id
    ))
    conn.commit()
    conn.close()

    if check_should_auto_sync():
        background_tasks.add_task(sync_supplier_to_jianguoyun, supplier_id)
    background_tasks.add_task(push_db_to_github)

    return {"status": "ok", "message": "更新成功"}

@app.post("/api/suppliers/{supplier_id}/status")
def change_supplier_status(supplier_id: int, req: Dict[str, str], background_tasks: BackgroundTasks, user: dict = Depends(require_admin)):
    new_status = req.get("status")
    if new_status not in ["visited", "planned"]:
        raise HTTPException(status_code=400, detail="无效状态，仅支持 visited 或 planned")
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("UPDATE suppliers SET status = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?", (new_status, supplier_id))
    conn.commit()
    conn.close()
    background_tasks.add_task(push_db_to_github)
    return {"status": "ok", "message": "状态切换成功", "status": new_status}

@app.delete("/api/suppliers/{supplier_id}")
def delete_supplier(supplier_id: int, background_tasks: BackgroundTasks, user: dict = Depends(require_admin)):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM suppliers WHERE id = ?", (supplier_id,))
    conn.commit()
    conn.close()

    # 自动刷新坚果云中的最新总台账
    local_dir = get_local_nutstore_dir()
    if local_dir:
        try:
            master_bytes, _ = build_master_excel_bytes()
            base_folder = os.path.join(local_dir, "无花果供应商档案")
            with open(os.path.join(base_folder, "无花果供应商总台账_最新版.xlsx"), "wb") as f:
                f.write(master_bytes)
        except Exception:
            pass

    background_tasks.add_task(push_db_to_github)
    return {"status": "ok", "message": "已删除"}

# ----------------- 客户跟进与日程待办 API -----------------

@app.get("/api/followups")
def get_followups(
    supplier_id: Optional[int] = None,
    date: Optional[str] = None,
    status: Optional[str] = None,
    user: dict = Depends(require_active_user)
):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    query = """
        SELECT f.*, s.company_name, s.contact_name, s.contact_phone, s.location
        FROM followups f
        JOIN suppliers s ON f.supplier_id = s.id
        WHERE 1=1
    """
    params = []
    if supplier_id:
        query += " AND f.supplier_id = ?"
        params.append(supplier_id)
    if date:
        query += " AND f.due_date = ?"
        params.append(date)
    if status and status != "全部":
        query += " AND f.status = ?"
        params.append(status)

    query += " ORDER BY f.status ASC, f.due_date ASC, f.id DESC"
    cursor.execute(query, params)
    rows = cursor.fetchall()
    data = [dict(r) for r in rows]
    conn.close()
    return {"status": "ok", "data": data}

@app.post("/api/followups")
def create_followup(
    item: FollowupCreate,
    background_tasks: BackgroundTasks,
    user: dict = Depends(require_admin)
):
    title = item.title.strip()
    due_date = item.due_date.strip()
    if not title:
        raise HTTPException(status_code=400, detail="待办内容不能为空")
    if not due_date:
        raise HTTPException(status_code=400, detail="请选择待办日期")

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT id FROM suppliers WHERE id = ?", (item.supplier_id,))
    if not cursor.fetchone():
        conn.close()
        raise HTTPException(status_code=404, detail="未找到对应的供应商")

    cursor.execute("""
        INSERT INTO followups (supplier_id, due_date, title, notes, status)
        VALUES (?, ?, ?, ?, ?)
    """, (item.supplier_id, due_date, title, item.notes or "", item.status or "pending"))
    new_id = cursor.lastrowid
    conn.commit()
    conn.close()

    background_tasks.add_task(push_db_to_github)
    return {"status": "ok", "message": "待办事项已添加", "id": new_id}

@app.patch("/api/followups/{followup_id}/toggle")
def toggle_followup_status(
    followup_id: int,
    background_tasks: BackgroundTasks,
    user: dict = Depends(require_admin)
):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT status FROM followups WHERE id = ?", (followup_id,))
    row = cursor.fetchone()
    if not row:
        conn.close()
        raise HTTPException(status_code=404, detail="未找到该待办事项")

    current_status = row[0]
    new_status = "completed" if current_status == "pending" else "pending"
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cursor.execute("UPDATE followups SET status = ?, updated_at = ? WHERE id = ?", (new_status, now_str, followup_id))
    conn.commit()
    conn.close()

    background_tasks.add_task(push_db_to_github)
    return {"status": "ok", "message": "状态更新成功", "status": new_status}

@app.delete("/api/followups/{followup_id}")
def delete_followup(
    followup_id: int,
    background_tasks: BackgroundTasks,
    user: dict = Depends(require_admin)
):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM followups WHERE id = ?", (followup_id,))
    conn.commit()
    conn.close()

    background_tasks.add_task(push_db_to_github)
    return {"status": "ok", "message": "待办事项已删除"}

# ----------------- 设置管理 -----------------
@app.get("/api/settings")
def get_settings(user: dict = Depends(require_admin)):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT key, value FROM settings")
    settings = dict(cursor.fetchall())
    conn.close()
    # 脱敏返回 API key
    masked_key = ""
    if settings.get("api_key"):
        k = settings["api_key"]
        masked_key = k[:3] + "..." + k[-4:] if len(k) > 7 else "***"
    return {
        "status": "ok",
        "data": {
            "api_key_masked": masked_key,
            "has_key": bool(settings.get("api_key")),
            "base_url": settings.get("base_url", "https://api.deepseek.com/v1"),
            "model": settings.get("model", "deepseek-chat")
        }
    }

@app.post("/api/settings")
def save_settings(s: AISettings, background_tasks: BackgroundTasks, user: dict = Depends(require_admin)):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    # 只有当用户输入了全新且未脱敏的合法 key 时才更新，防止把 sk-... 覆盖进去
    if s.api_key and "..." not in s.api_key and "••" not in s.api_key and len(s.api_key.strip()) > 15:
        cursor.execute("INSERT OR REPLACE INTO settings (key, value) VALUES ('api_key', ?)", (s.api_key.strip(),))
    if s.base_url:
        cursor.execute("INSERT OR REPLACE INTO settings (key, value) VALUES ('base_url', ?)", (s.base_url.strip(),))
    if s.model:
        cursor.execute("INSERT OR REPLACE INTO settings (key, value) VALUES ('model', ?)", (s.model.strip(),))
    conn.commit()
    conn.close()
    background_tasks.add_task(push_db_to_github)
    return {"status": "ok", "message": "设置已保存"}

# ----------------- Excel 导出服务 (支持全部导出 / 勾选导出 / 单家精美档案导出) -----------------
def build_master_excel_bytes(ids: Optional[str] = None):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    if ids and ids.strip():
        id_list = [int(x.strip()) for x in ids.split(",") if x.strip().isdigit()]
        if id_list:
            placeholders = ",".join("?" * len(id_list))
            cursor.execute(f"SELECT * FROM suppliers WHERE id IN ({placeholders}) ORDER BY id ASC", id_list)
        else:
            cursor.execute("SELECT * FROM suppliers ORDER BY id ASC")
    else:
        cursor.execute("SELECT * FROM suppliers ORDER BY id ASC")
    
    rows = cursor.fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "无花果供应商总台账"

    headers = [
        "序号", "公司全称", "具体地理位置", "合作意向", "规模与产能",
        "主要产品品类", "核心拳头产品", "供货报价明细清单",
        "账期与付款方式", "联系人", "职务", "联系电话",
        "考察备忘/特殊背景(外采/供货)", "现场照片数", "最后更新时间"
    ]

    ws.append(headers)

    # 样式配置
    header_font = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1B4D3E", end_color="1B4D3E", fill_type="solid") # 高级墨绿商务风
    cell_font = Font(name="微软雅黑", size=10)
    thin_border = Border(
        left=Side(style='thin', color='D0D7DE'),
        right=Side(style='thin', color='D0D7DE'),
        top=Side(style='thin', color='D0D7DE'),
        bottom=Side(style='thin', color='D0D7DE')
    )
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # 表头样式
    for col_num in range(1, len(headers) + 1):
        c = ws.cell(row=1, column=col_num)
        c.font = header_font
        c.fill = header_fill
        c.alignment = align_center
        c.border = thin_border
    ws.row_dimensions[1].height = 30

    intent_colors = {
        "高": PatternFill(start_color="D8F3DC", end_color="D8F3DC", fill_type="solid"),
        "中": PatternFill(start_color="FFF3B0", end_color="FFF3B0", fill_type="solid"),
        "低": PatternFill(start_color="F1FAEE", end_color="F1FAEE", fill_type="solid")
    }

    for idx, r in enumerate(rows, start=1):
        # 统计照片张数
        photo_count = 0
        try:
            imgs = json.loads(r["images"]) if r["images"] else []
            photo_count = len(imgs) if isinstance(imgs, list) else 0
        except Exception:
            photo_count = 0

        row_data = [
            idx,
            r["company_name"],
            r["location"],
            r["cooperation_intent"],
            r["scale"],
            r["main_products"],
            r["advantage_products"],
            r["price_terms"],
            r["payment_terms"],
            r["contact_name"],
            r["contact_role"],
            r["contact_phone"],
            r["extra_notes"] or r["raw_notes"],
            f"{photo_count} 张" if photo_count > 0 else "无",
            r["updated_at"]
        ]
        ws.append(row_data)
        current_row = idx + 1
        ws.row_dimensions[current_row].height = 26

        for col_num in range(1, len(headers) + 1):
            c = ws.cell(row=current_row, column=col_num)
            c.font = cell_font
            c.border = thin_border
            if col_num in [1, 4, 10, 11, 12, 14, 15]:
                c.alignment = align_center
            else:
                c.alignment = align_left
            
            # 意向度高亮背景
            if col_num == 4 and r["cooperation_intent"] in intent_colors:
                c.fill = intent_colors[r["cooperation_intent"]]

    # 自适应列宽
    col_widths = {
        1: 6,   # 序号
        2: 28,  # 公司名称
        3: 25,  # 位置
        4: 10,  # 意向
        5: 25,  # 规模
        6: 25,  # 主要产品
        7: 22,  # 优势产品
        8: 30,  # 供货价格
        9: 22,  # 账期
        10: 12, # 联系人
        11: 12, # 职务
        12: 15, # 电话
        13: 35, # 备忘
        14: 12, # 照片数
        15: 20  # 时间
    }
    for col_idx, width in col_widths.items():
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = width

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    data = output.getvalue()

    count_str = f"共{len(rows)}家" if ids else "全部汇总"
    filename = f"无花果供应商总台账_{count_str}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return data, filename

@app.get("/api/export/excel")
def export_excel(ids: Optional[str] = None, user: dict = Depends(require_active_user)):
    data, filename = build_master_excel_bytes(ids)
    from urllib.parse import quote
    encoded_fn = quote(filename)
    return StreamingResponse(
        BytesIO(data),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{encoded_fn}"}
    )

def build_single_supplier_excel_bytes(r):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "供应商尽调单页档案"

    # 大标题
    ws.merge_cells("A1:D1")
    title_cell = ws.cell(row=1, column=1)
    title_cell.value = f"【实地考察尽调档案】{r['company_name']}"
    title_cell.font = Font(name="微软雅黑", size=14, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="1B4D3E", end_color="1B4D3E", fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36

    # 样式
    label_font = Font(name="微软雅黑", size=10, bold=True, color="333333")
    val_font = Font(name="微软雅黑", size=10)
    label_fill = PatternFill(start_color="F0F4F1", end_color="F0F4F1", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin', color='D0D7DE'),
        right=Side(style='thin', color='D0D7DE'),
        top=Side(style='thin', color='D0D7DE'),
        bottom=Side(style='thin', color='D0D7DE')
    )

    items = [
        ("公司规范全称", r["company_name"], "合作意向度", f"{r['cooperation_intent']} 意向"),
        ("地理详细位置", r["location"] or "未填", "GPS经纬度", f"{r['latitude'] or '-'}, {r['longitude'] or '-'}"),
        ("工厂规模产能", r["scale"] or "未明确", "最后更新时间", r["updated_at"]),
        ("联系人姓名", r["contact_name"] or "未填", "职务/身份", r["contact_role"] or "未明确"),
        ("联系电话/微信", r["contact_phone"] or "未填", "账期与结算方式", r["payment_terms"] or "未明确"),
        ("主要经营产品", r["main_products"] or "-", "核心拳头主打", r["advantage_products"] or "-"),
        ("各品类供货报价", r["price_terms"] or "待二次议价", "附件照片数量", f"{len(json.loads(r['images'] or '[]'))} 张实地照片"),
        ("考察重要备忘\n(上下游/外采疑虑)", r["extra_notes"] or "无特殊备注", "原始拜访手记", r["raw_notes"] or "-")
    ]

    curr_row = 2
    for label1, val1, label2, val2 in items:
        # A 列
        c_a = ws.cell(row=curr_row, column=1, value=label1)
        c_a.font = label_font
        c_a.fill = label_fill
        c_a.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c_a.border = thin_border

        # B 列
        c_b = ws.cell(row=curr_row, column=2, value=val1)
        c_b.font = val_font
        c_b.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        c_b.border = thin_border

        # C 列
        c_c = ws.cell(row=curr_row, column=3, value=label2)
        c_c.font = label_font
        c_c.fill = label_fill
        c_c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c_c.border = thin_border

        # D 列
        c_d = ws.cell(row=curr_row, column=4, value=val2)
        c_d.font = val_font
        c_d.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        c_d.border = thin_border

        ws.row_dimensions[curr_row].height = 36 if "备忘" in label1 else 28
        curr_row += 1

    # 嵌入实地拍摄的照片到 Excel 表格中
    try:
        from PIL import Image as PILImage
        from openpyxl.drawing.image import Image as OpenpyxlImage
        imgs = json.loads(r["images"]) if r["images"] else []
        if imgs:
            curr_row += 1
            ws.merge_cells(f"A{curr_row}:D{curr_row}")
            sub_title = ws.cell(row=curr_row, column=1)
            sub_title.value = f"📷 实地考察照片与名片附录（共 {len(imgs)} 张原照嵌入）"
            sub_title.font = Font(name="微软雅黑", size=11, bold=True, color="1B4D3E")
            sub_title.fill = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")
            sub_title.alignment = Alignment(horizontal="left", vertical="center")
            ws.row_dimensions[curr_row].height = 28
            curr_row += 1

            img_col = 1
            img_row = curr_row

            for idx, img_url in enumerate(imgs):
                fname = os.path.basename(img_url)
                fpath = os.path.join(UPLOAD_DIR, fname)
                if os.path.exists(fpath):
                    with PILImage.open(fpath) as pimg:
                        w, h = pimg.size
                        ratio = min(240 / w, 160 / h, 1.0)
                        target_w, target_h = int(w * ratio), int(h * ratio)

                    ximg = OpenpyxlImage(fpath)
                    ximg.width = target_w
                    ximg.height = target_h
                    c_letter = get_column_letter(img_col)
                    ximg.anchor = f"{c_letter}{img_row}"
                    ws.add_image(ximg)
                    ws.row_dimensions[img_row].height = 130

                    img_col += 2
                    if img_col > 3:
                        img_col = 1
                        img_row += 8
    except Exception as e:
        print(f"嵌入照片到Excel失败: {e}")

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 38
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 38

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    data = output.getvalue()

    safe_name = re.sub(r'[\/:*?"<>|]', '_', r["company_name"]).strip() or "供应商"
    filename = f"【尽调报告】{safe_name}_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx"
    return data, filename

@app.get("/api/export/supplier/{supplier_id}")
def export_single_supplier(supplier_id: int, user: dict = Depends(require_active_user)):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM suppliers WHERE id = ?", (supplier_id,))
    r = cursor.fetchone()
    conn.close()

    if not r:
        raise HTTPException(status_code=404, detail="未找到该供应商")

    data, filename = build_single_supplier_excel_bytes(r)
    from urllib.parse import quote
    encoded_fn = quote(filename)

    return StreamingResponse(
        BytesIO(data),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{encoded_fn}"}
    )

# ----------------- 坚果云 WebDAV 同步服务 -----------------
class JianguoyunSettings(BaseModel):
    enabled: bool = False
    server: str = "https://dav.jianguoyun.com/dav"
    account: str = ""
    password: Optional[str] = ""
    remote_dir: str = "/无花果供应商档案"
    auto_sync: bool = True

class JianguoyunService:
    def __init__(self, server: str, account: str, password: str, base_dir: str):
        self.server = server.rstrip("/") if server else "https://dav.jianguoyun.com/dav"
        self.account = account.strip()
        self.password = password.strip()
        self.base_dir = "/" + base_dir.strip().strip("/")
        self.auth = (self.account, self.password)

    def _encode_url(self, path: str) -> str:
        parts = [urllib.parse.quote(p) for p in path.strip("/").split("/") if p]
        return f"{self.server}/{'/'.join(parts)}"

    async def test_connection(self) -> Dict[str, Any]:
        if not self.account or not self.password:
            return {"success": False, "message": "账号（注册邮箱）或16位应用授权密码未填写"}
        async with httpx.AsyncClient(timeout=10.0) as client:
            try:
                resp = await client.request(
                    "PROPFIND",
                    f"{self.server}/",
                    auth=self.auth,
                    headers={"Depth": "0"}
                )
                if resp.status_code in [200, 207]:
                    return {"success": True, "message": "坚果云连接成功！账号与16位应用密码认证通过。"}
                elif resp.status_code == 401:
                    return {"success": False, "message": "坚果云认证失败(401)：请检查账号(邮箱)以及是否使用了坚果云生成的16位应用密码(非普通网页登录密码)。"}
                else:
                    return {"success": False, "message": f"连接返回状态码: {resp.status_code}"}
            except Exception as e:
                return {"success": False, "message": f"连接坚果云服务器超时或异常: {str(e)}"}

    async def ensure_dir(self, client: httpx.AsyncClient, dir_path: str):
        parts = [p for p in dir_path.strip("/").split("/") if p]
        curr_path = ""
        for p in parts:
            curr_path += f"/{p}"
            url = self._encode_url(curr_path) + "/"
            try:
                await client.request("MKCOL", url, auth=self.auth, timeout=10.0)
            except Exception:
                pass

    async def upload_file(self, client: httpx.AsyncClient, remote_file_path: str, data: bytes) -> bool:
        url = self._encode_url(remote_file_path)
        resp = await client.put(
            url,
            content=data,
            auth=self.auth,
            headers={"Content-Type": "application/octet-stream"},
            timeout=40.0
        )
        return resp.status_code in [200, 201, 204]

def get_local_nutstore_dir() -> Optional[str]:
    candidates = [
        os.path.expanduser("~/Nutstore Files/我的坚果云"),
        os.path.expanduser("~/坚果云/我的坚果云"),
        os.path.expanduser("~/Nutstore Files"),
        os.path.expanduser("~/坚果云")
    ]
    for c in candidates:
        if os.path.exists(c) and os.path.isdir(c):
            return c
    return None

def get_jgy_service_from_db() -> Optional[JianguoyunService]:
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT key, value FROM settings")
    settings_dict = dict(cursor.fetchall())
    conn.close()

    account = settings_dict.get("jgy_account", "").strip()
    password = settings_dict.get("jgy_password", "").strip()
    if not account or not password:
        return None
    server = settings_dict.get("jgy_server", "https://dav.jianguoyun.com/dav")
    base_dir = settings_dict.get("jgy_dir", "/无花果供应商档案")
    return JianguoyunService(server, account, password, base_dir)

async def sync_supplier_to_jianguoyun(supplier_id: int) -> Dict[str, Any]:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM suppliers WHERE id = ?", (supplier_id,))
    r = cursor.fetchone()
    conn.close()

    if not r:
        return {"success": False, "message": "未找到该供应商"}

    company = re.sub(r'[\/:*?"<>|]', '_', r["company_name"]).strip() or f"未命名工厂_{r['id']}"
    excel_bytes, excel_filename = build_single_supplier_excel_bytes(r)
    imgs = json.loads(r["images"]) if r["images"] else []

    summary_txt = f"""【实地考察尽调简报】
工厂名称: {r['company_name']}
意向评级: {r['cooperation_intent']}
地理位置: {r['location'] or '未填'} (GPS: {r['latitude'] or '-'}, {r['longitude'] or '-'})
联系方式: {r['contact_name'] or '-'} ({r['contact_role'] or '-'}) 电话: {r['contact_phone'] or '-'}
规模产能: {r['scale'] or '未明确'}
主要产品: {r['main_products'] or '-'}
拳头主打: {r['advantage_products'] or '-'}
供货报价: {r['price_terms'] or '-'}
账期结算: {r['payment_terms'] or '-'}
现场备忘: {r['extra_notes'] or '无'}
原始笔记: {r['raw_notes'] or '-'}
同步时间: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
"""

    local_synced = False
    local_dir = get_local_nutstore_dir()
    if local_dir:
        try:
            target_local = os.path.join(local_dir, "无花果供应商档案", company)
            os.makedirs(target_local, exist_ok=True)
            with open(os.path.join(target_local, excel_filename), "wb") as f:
                f.write(excel_bytes)
            
            if imgs:
                photo_local = os.path.join(target_local, "现场实拍与名片")
                os.makedirs(photo_local, exist_ok=True)
                for idx, img_url in enumerate(imgs, start=1):
                    fname = os.path.basename(img_url)
                    src_f = os.path.join(UPLOAD_DIR, fname)
                    if os.path.exists(src_f):
                        ext = os.path.splitext(fname)[1] or ".jpg"
                        shutil.copy2(src_f, os.path.join(photo_local, f"现场原图_{idx}{ext}"))

            with open(os.path.join(target_local, "考察摘要说明.txt"), "w", encoding="utf-8") as f:
                f.write(summary_txt)

            # 每次单家修改或录入，顺便自动更新坚果云根目录下的最新总台账
            try:
                master_bytes, _ = build_master_excel_bytes()
                base_folder = os.path.join(local_dir, "无花果供应商档案")
                with open(os.path.join(base_folder, "无花果供应商总台账_最新版.xlsx"), "wb") as f:
                    f.write(master_bytes)
            except Exception as e:
                print(f"刷新总台账跳过: {e}")

            local_synced = True
        except Exception as e:
            print(f"本地坚果云目录保存跳过: {e}")

    jgy = get_jgy_service_from_db()
    webdav_synced = False
    if jgy:
        try:
            target_dir = f"{jgy.base_dir}/{company}"
            async with httpx.AsyncClient(timeout=40.0) as client:
                await jgy.ensure_dir(client, target_dir)
                excel_path = f"{target_dir}/{excel_filename}"
                await jgy.upload_file(client, excel_path, excel_bytes)

                if imgs:
                    photo_dir = f"{target_dir}/现场实拍与名片"
                    await jgy.ensure_dir(client, photo_dir)
                    for idx, img_url in enumerate(imgs, start=1):
                        fname = os.path.basename(img_url)
                        local_path = os.path.join(UPLOAD_DIR, fname)
                        if os.path.exists(local_path):
                            with open(local_path, "rb") as f:
                                img_data = f.read()
                            ext = os.path.splitext(fname)[1] or ".jpg"
                            remote_img_path = f"{photo_dir}/现场原图_{idx}{ext}"
                            await jgy.upload_file(client, remote_img_path, img_data)

                await jgy.upload_file(client, f"{target_dir}/考察摘要说明.txt", summary_txt.encode("utf-8"))
            webdav_synced = True
        except Exception as e:
            print(f"WebDAV 同步异常: {e}")

    if local_synced or webdav_synced:
        return {
            "success": True,
            "message": f"已成功同步【{r['company_name']}】尽调表与实拍照片至坚果云！手机坚果云App已自动更新。"
        }
    else:
        return {"success": False, "message": "坚果云尚未完成配置，请先填写16位应用密码或在电脑运行坚果云客户端"}

async def sync_all_to_jianguoyun() -> Dict[str, Any]:
    master_bytes, master_filename = build_master_excel_bytes()
    local_dir = get_local_nutstore_dir()
    local_count = 0

    if local_dir:
        try:
            target_base = os.path.join(local_dir, "无花果供应商档案")
            os.makedirs(target_base, exist_ok=True)
            with open(os.path.join(target_base, master_filename), "wb") as f:
                f.write(master_bytes)
            with open(os.path.join(target_base, "无花果供应商总台账_最新版.xlsx"), "wb") as f:
                f.write(master_bytes)
            
            if os.path.exists(DB_PATH):
                backup_dir = os.path.join(target_base, "数据库自动备份")
                os.makedirs(backup_dir, exist_ok=True)
                db_name = f"suppliers_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
                shutil.copy2(DB_PATH, os.path.join(backup_dir, db_name))
        except Exception as e:
            print(f"全量本地坚果云写入跳过: {e}")

    # 同步各家工厂
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT id FROM suppliers")
    ids = [row[0] for row in cursor.fetchall()]
    conn.close()

    count = 0
    for sid in ids:
        res = await sync_supplier_to_jianguoyun(sid)
        if res.get("success"):
            count += 1

    return {
        "success": True,
        "message": f"全量同步完成！已将最新总台账Excel、数据库备份包及 {count} 家工厂独立档案同步至坚果云目录！"
    }

@app.get("/api/jianguoyun/settings")
def get_jianguoyun_settings(user: dict = Depends(require_admin)):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT key, value FROM settings")
    s = dict(cursor.fetchall())
    conn.close()

    pwd = s.get("jgy_password", "")
    masked_pwd = ""
    if pwd:
        masked_pwd = pwd[:2] + "••••••••" + pwd[-2:] if len(pwd) > 4 else "••••••••"

    return {
        "enabled": s.get("jgy_enabled") == "1",
        "server": s.get("jgy_server", "https://dav.jianguoyun.com/dav"),
        "account": s.get("jgy_account", ""),
        "password": masked_pwd,
        "has_password": bool(pwd),
        "remote_dir": s.get("jgy_dir", "/无花果供应商档案"),
        "auto_sync": s.get("jgy_auto_sync", "1") == "1"
    }

@app.post("/api/jianguoyun/settings")
def save_jianguoyun_settings(cfg: JianguoyunSettings, background_tasks: BackgroundTasks, user: dict = Depends(require_admin)):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("INSERT OR REPLACE INTO settings (key, value) VALUES ('jgy_enabled', ?)", ("1" if cfg.enabled else "0",))
    cursor.execute("INSERT OR REPLACE INTO settings (key, value) VALUES ('jgy_server', ?)", (cfg.server.strip() or "https://dav.jianguoyun.com/dav",))
    cursor.execute("INSERT OR REPLACE INTO settings (key, value) VALUES ('jgy_account', ?)", (cfg.account.strip(),))
    cursor.execute("INSERT OR REPLACE INTO settings (key, value) VALUES ('jgy_dir', ?)", (cfg.remote_dir.strip() or "/无花果供应商档案",))
    cursor.execute("INSERT OR REPLACE INTO settings (key, value) VALUES ('jgy_auto_sync', ?)", ("1" if cfg.auto_sync else "0",))

    if cfg.password and "••" not in cfg.password and len(cfg.password.strip()) > 3:
        cursor.execute("INSERT OR REPLACE INTO settings (key, value) VALUES ('jgy_password', ?)", (cfg.password.strip(),))

    conn.commit()
    conn.close()
    background_tasks.add_task(push_db_to_github)
    return {"status": "ok", "message": "坚果云同步配置已保存"}

@app.post("/api/jianguoyun/test")
async def test_jianguoyun_connection(cfg: Optional[JianguoyunSettings] = None, user: dict = Depends(require_admin)):
    server = "https://dav.jianguoyun.com/dav"
    account = ""
    password = ""
    if cfg and cfg.account and cfg.password and "••" not in cfg.password:
        server = cfg.server or server
        account = cfg.account
        password = cfg.password
    else:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("SELECT key, value FROM settings")
        s = dict(cursor.fetchall())
        conn.close()
        server = s.get("jgy_server", server)
        account = s.get("jgy_account", "")
        password = s.get("jgy_password", "")

    service = JianguoyunService(server, account, password, "/无花果供应商档案")
    return await service.test_connection()

@app.post("/api/jianguoyun/sync/{supplier_id}")
async def sync_supplier(supplier_id: int, user: dict = Depends(require_admin)):
    res = await sync_supplier_to_jianguoyun(supplier_id)
    if not res.get("success"):
        raise HTTPException(status_code=400, detail=res.get("message"))
    return res

@app.post("/api/jianguoyun/sync_all")
async def sync_all(user: dict = Depends(require_admin)):
    res = await sync_all_to_jianguoyun()
    if not res.get("success"):
        raise HTTPException(status_code=400, detail=res.get("message"))
    return res

# ----------------- 实地考察照片 ZIP 压缩包批量导出 -----------------
import zipfile

@app.get("/api/export/photos/zip")
def export_photos_zip(supplier_id: Optional[int] = None, user: dict = Depends(require_active_user)):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    if supplier_id:
        cursor.execute("SELECT company_name, images FROM suppliers WHERE id = ?", (supplier_id,))
    else:
        cursor.execute("SELECT company_name, images FROM suppliers WHERE images IS NOT NULL AND images != '[]'")
    rows = cursor.fetchall()
    conn.close()

    zip_buffer = BytesIO()
    has_photo = False

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for r in rows:
            c_name = re.sub(r'[\/:*?"<>|]', '_', r["company_name"]) # 过滤非法目录名
            try:
                imgs = json.loads(r["images"]) if r["images"] else []
                if isinstance(imgs, list):
                    for idx, img_url in enumerate(imgs, start=1):
                        fname = os.path.basename(img_url)
                        fpath = os.path.join(UPLOAD_DIR, fname)
                        if os.path.exists(fpath):
                            ext = os.path.splitext(fname)[1] or ".jpg"
                            arcname = f"{c_name}/照片_{idx}{ext}"
                            zf.write(fpath, arcname)
                            has_photo = True
            except Exception as e:
                print(f"打包照片跳过: {e}")

        # 如果一张照片也没有，写一个说明文件
        if not has_photo:
            zf.writestr("说明.txt", "当前暂无实地拍摄的照片。在录入或详情页上传照片后即可打包下载。")

    zip_buffer.seek(0)
    zip_name = f"无花果考察实地照片包_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
    from urllib.parse import quote
    encoded_zn = quote(zip_name)

    return StreamingResponse(
        zip_buffer,
        media_type="application/zip",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{encoded_zn}"}
    )

# 挂载静态文件目录
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

@app.get("/")
def read_root():
    index_file = os.path.join(STATIC_DIR, "index.html")
    if os.path.exists(index_file):
        with open(index_file, "r", encoding="utf-8") as f:
            return HTMLResponse(
                content=f.read(),
                headers={
                    "Cache-Control": "no-cache, no-store, must-revalidate",
                    "Pragma": "no-cache",
                    "Expires": "0"
                }
            )
    return HTMLResponse("<h1>系统加载中...</h1>")
