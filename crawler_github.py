#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
海关政策爬虫 - 海关总署政务公开·最新文件栏目
  栏目：最新文件（/customs/302249/2480148/index.html）
缓存：policy_cache.json
邮件附件：发布历史(from_2025.10).xlsx（超链接可点，列宽自适应）

反爬说明：customs.gov.cn 使用加速乐(Jiasule) WAF，直接用 requests 会被 412 拦截。
方案：Playwright 先访问首页完成 JS 挑战拿到全套 cookie，之后用 requests 带 cookie 抓取。
"""

import os, re, time, smtplib, hashlib, json
from datetime import datetime
from urllib.parse import urljoin
import requests
import pandas as pd
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

# ==================== 配置 ====================
CONFIG = {
    "output_dir": ".",
    "cache_file": "policy_cache.json",
    "history_xlsx": "发布历史(from_2025.10).xlsx",

    "base_url": "http://www.customs.gov.cn",
    "entry_url": "http://www.customs.gov.cn/",          # WAF 挑战入口（首页挑战成功率高）
    "policy_index": "http://www.customs.gov.cn/customs/302249/index.html",

    # 只抓"最新文件"栏目：http://www.customs.gov.cn/customs/302249/2480148/index.html
    "categories": {
        "最新文件": {"path": "/customs/302249/2480148", "eportalPageId": "2480148", "moduleId": "ba2f639c351e44d1b28c52c03eee13f8"},
    },

    "first_start_date": "2025-10-29",   # 首次运行抓取最近1年
    "days_limit": 1,
    "request_delay": 0.5,
    "max_pages_first_run": 40,      # 首次运行单栏目最多翻页数（保护）

    "headers": {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
        "Accept-Language": "zh-CN,zh;q=0.9",
    },

    "email": {
        "enabled": True,
        "smtp_server": "smtp.qq.com",
        "smtp_port": 465,
        "use_ssl": True,
        "sender": "1992915150@qq.com",
        "password": os.environ.get("EMAIL_PASSWORD", ""),
        "recipients": ["weici.pan1@benteler.com"],
        "subject_prefix": "【政策监控】",
        "send_if_empty": True,
    }
}

def get_cache_path():
    return os.path.join(CONFIG["output_dir"], CONFIG["cache_file"])

def get_history_xlsx_path():
    return os.path.join(CONFIG["output_dir"], CONFIG["history_xlsx"])

# ==================== WAF 会话 ====================
def create_waf_session():
    """Playwright 完成加速乐 JS 挑战，返回带全套 cookie 的 requests.Session"""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("[WAF] 未安装 playwright，无法过 WAF（pip install playwright && playwright install chromium）")
        return None

    jar = {}
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=["--disable-blink-features=AutomationControlled", "--no-sandbox", "--disable-dev-shm-usage"],
        )
        ctx = browser.new_context(
            user_agent=CONFIG["headers"]["User-Agent"],
            locale="zh-CN",
            viewport={"width": 1366, "height": 768},
        )
        page = ctx.new_page()
        try:
            page.goto(CONFIG["entry_url"], wait_until="domcontentloaded", timeout=60000)
            # 轮询等待挑战完成：cookie 齐全且首页内容渲染出来
            for i in range(20):
                time.sleep(1.5)
                cookies_now = ctx.cookies(CONFIG["base_url"])
                try:
                    ln = len(page.content())
                except Exception:
                    ln = 0
                if len(cookies_now) >= 5 and ln > 30000:
                    break
            jar = {c["name"]: c["value"] for c in ctx.cookies(CONFIG["base_url"])}
        except Exception as e:
            print(f"[WAF] 挑战过程异常: {e}")
        finally:
            browser.close()

    if len(jar) < 3:
        print(f"[WAF] 挑战失败，仅获得 {len(jar)} 个 cookie")
        return None

    s = requests.Session()
    s.headers.update(CONFIG["headers"])
    for k, v in jar.items():
        s.cookies.set(k, v, domain="www.customs.gov.cn")
    print(f"[WAF] 挑战成功，获得 {len(jar)} 个 cookie")
    return s

# ==================== 缓存类 ====================
class PolicyCache:
    def __init__(self, cache_file):
        self.cache_file = cache_file
        self.seen = set()
        self.records = []           # 每项包含 url, title, category, date
        self._load()

    def _load(self):
        if not os.path.exists(self.cache_file):
            print("[缓存] 文件不存在，将创建新缓存")
            return
        try:
            with open(self.cache_file, 'r', encoding='utf-8') as f:
                content = f.read().strip()
                if not content:
                    raise ValueError("空文件")
                data = json.loads(content)
            if isinstance(data, list):
                for item in data:
                    if isinstance(item, dict) and 'url' in item and 'title' in item:
                        url = item['url']
                        title = item['title']
                        key = hashlib.md5(f"{url}|{title}".encode()).hexdigest()
                        self.seen.add(key)
                        self.records.append(item)
                print(f"[缓存] 成功加载 {len(self.records)} 条记录")
            else:
                print("[缓存] JSON 格式不符，将重建")
        except Exception as e:
            print(f"[缓存] 加载失败 ({e})，备份损坏文件并重建")
            bak = self.cache_file + ".bak"
            try:
                os.rename(self.cache_file, bak)
                print(f"[缓存] 已备份到 {bak}")
            except:
                pass

    def is_new(self, url, title):
        key = hashlib.md5(f"{url}|{title}".encode()).hexdigest()
        return key not in self.seen

    def add(self, policy):
        key = hashlib.md5(f"{policy['原文链接']}|{policy['政策名称']}".encode()).hexdigest()
        self.seen.add(key)
        self.records.append({
            'url': policy['原文链接'],
            'title': policy['政策名称'],
            'category': policy.get('分类', ''),
            'date': policy['发布日期']
        })

    def save(self):
        with open(self.cache_file, 'w', encoding='utf-8') as f:
            json.dump(self.records, f, ensure_ascii=False, indent=2)
        print(f"[缓存] JSON 已保存 ({len(self.records)} 条)")

    def export_excel(self, xlsx_path):
        """导出为 Excel，URL 列在最后且为超链接，列宽自适应"""
        if not self.records:
            pd.DataFrame(columns=['title', 'category', 'date', 'url']).to_excel(
                xlsx_path, index=False, engine='openpyxl')
            print("[导出] 无记录，生成空 Excel")
            return

        df = pd.DataFrame(self.records)
        df = df.rename(columns={'title': '政策名称', 'category': '分类', 'date': '发布日期', 'url': '链接'})
        df = df[['政策名称', '分类', '发布日期', '链接']]

        with pd.ExcelWriter(xlsx_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='政策历史')

        wb = load_workbook(xlsx_path)
        ws = wb.active

        link_col = 4
        for row in range(2, ws.max_row + 1):
            cell = ws.cell(row=row, column=link_col)
            url = cell.value
            if url and isinstance(url, str) and url.startswith('http'):
                cell.hyperlink = url
                cell.font = Font(color='0563C1', underline='single')
                cell.alignment = Alignment(wrap_text=True)

        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    cell_length = len(str(cell.value))
                    if cell_length > max_length:
                        max_length = cell_length
                except:
                    pass
            adjusted_width = min(max_length + 2, 80)
            adjusted_width = max(adjusted_width, 10)
            ws.column_dimensions[column].width = adjusted_width

        ws.column_dimensions['A'].width = max(ws.column_dimensions['A'].width, 40)
        ws.column_dimensions['D'].width = max(ws.column_dimensions['D'].width, 50)

        wb.save(xlsx_path)
        print(f"[导出] Excel 已生成: {xlsx_path} ({len(self.records)} 条，链接可点击)")

# ==================== 工具函数 ====================
def parse_date(date_str):
    if not date_str:
        return ""
    date_str = date_str.strip().replace("年", "-").replace("月", "-").replace("日", "")
    m = re.search(r'(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})', date_str)
    if m:
        y, mth, d = m.groups()
        return f"{y}-{int(mth):02d}-{int(d):02d}"
    m = re.search(r'(?<!\d)(\d{1,2})[-/.](\d{1,2})(?!\d)', date_str)
    if m:
        month, day = int(m.group(1)), int(m.group(2))
        year = datetime.now().year
        return f"{year}-{month:02d}-{day:02d}"
    return date_str

def within_days(date_str, days):
    if days <= 0 or not date_str:
        return True
    try:
        d = datetime.strptime(date_str, "%Y-%m-%d")
        return (datetime.now() - d).days <= days
    except:
        return True

def build_page_url(cat_cfg, page):
    """栏目第N页URL：第1页 index.html；第2页起 eportal 动态接口（静态 moduleId-N.html 实际不存在，404）"""
    path = cat_cfg["path"]
    if page == 1:
        return CONFIG["base_url"] + path + "/index.html"
    return (f"{CONFIG['base_url']}/eportal/ui?pageId={cat_cfg['eportalPageId']}"
            f"&currentPage={page}&moduleId={cat_cfg['moduleId']}&staticRequest=yes")

def parse_list_page(html, base):
    """解析栏目列表页：ul.news_list li -> a[title] + span日期"""
    soup = BeautifulSoup(html, 'html.parser')
    items = []
    for li in soup.select('ul.news_list li'):
        a = li.find('a')
        if not a:
            continue
        title = a.get('title') or a.get_text(strip=True)
        href = a.get('href', '')
        if not title or not href:
            continue
        date_tag = li.find('span')
        date_str = parse_date(date_tag.get_text(strip=True)) if date_tag else ""
        link = href if href.startswith("http") else urljoin(base, href)
        items.append({"title": title, "link": link, "date": date_str})
    return items

# ==================== 抓取 ====================
def fetch_category_policies(session, cat_name, cat_cfg, cache, is_first_run):
    policies = []
    page = 1
    max_pages = CONFIG["max_pages_first_run"] if is_first_run else 3

    while page <= max_pages:
        url = build_page_url(cat_cfg, page)
        print(f"  [抓取] 第 {page} 页")
        try:
            resp = session.get(url, timeout=30)
            if resp.status_code != 200 or len(resp.text) < 5000:
                print(f"    响应异常 status={resp.status_code} len={len(resp.text)}，停止翻页")
                break
            resp.encoding = 'utf-8'
            items = parse_list_page(resp.text, CONFIG["base_url"])
            if not items:
                print("    未解析到列表项，停止翻页")
                break

            page_added = 0
            last_date = None
            for it in items:
                title, link, date_str = it["title"], it["link"], it["date"]
                if date_str:
                    last_date = date_str

                if not cache.is_new(link, title):
                    continue
                if is_first_run:
                    if date_str and date_str < CONFIG["first_start_date"]:
                        continue
                else:
                    if not within_days(date_str, CONFIG["days_limit"]):
                        continue

                policy = {
                    "分类": cat_name,
                    "来源": "海关总署",
                    "政策名称": title,
                    "发布日期": date_str,
                    "原文链接": link,
                    "抓取时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }
                policies.append(policy)
                cache.add(policy)
                page_added += 1

            print(f"    本页新增 {page_added} 条（列表共 {len(items)} 条，最早 {items[-1]['date'] or '?' }）")
            # 翻页终止条件
            if is_first_run and last_date and last_date < CONFIG["first_start_date"]:
                break
            if not is_first_run and page_added == 0:
                break
            time.sleep(CONFIG["request_delay"])
            page += 1
        except Exception as e:
            print(f"  [异常] {e}")
            break
    return policies

def crawl_all(cache):
    print(f"\n>>> 开始抓取...")
    is_first_run = (len(cache.records) == 0)
    if is_first_run:
        print(f"[模式] 首次运行，抓取 {CONFIG['first_start_date']} 至今的全部政策")
    else:
        print(f"[模式] 日常运行，仅抓取最近 {CONFIG['days_limit']} 天的政策")

    session = create_waf_session()
    if session is None:
        print("[错误] 无法建立 WAF 会话，抓取中止")
        return []

    # 用第一个栏目首页验证会话有效性
    probe = session.get(build_page_url(list(CONFIG["categories"].values())[0], 1), timeout=30)
    if probe.status_code != 200 or len(probe.text) < 5000:
        print(f"[错误] 会话验证失败 status={probe.status_code}，抓取中止")
        return []

    all_policies = []
    for cat_name, cat_cfg in CONFIG["categories"].items():
        print(f"\n>>> 栏目：{cat_name}")
        cat_policies = fetch_category_policies(session, cat_name, cat_cfg, cache, is_first_run)
        all_policies.extend(cat_policies)
        print(f"    {cat_name} 新增 {len(cat_policies)} 条")
        time.sleep(CONFIG["request_delay"])
    print(f"\n[总计] 新增 {len(all_policies)} 条")
    return all_policies

# ==================== 邮件 ====================
def send_email(new_policies, xlsx_path):
    mail_cfg = CONFIG["email"]
    if not mail_cfg["enabled"]:
        return
    if not mail_cfg["password"]:
        print("[错误] 邮箱密码未设置")
        return
    if not new_policies and not mail_cfg["send_if_empty"]:
        return

    try:
        from email.mime.text import MIMEText
        from email.mime.multipart import MIMEMultipart
        from email.mime.application import MIMEApplication

        if new_policies:
            body = f"<h3>本次新增 {len(new_policies)} 条海关政策</h3>"
            body += "<table border='1' cellpadding='5' style='border-collapse:collapse'>"
            body += "<tr><th>分类</th><th>来源</th><th>政策名称</th><th>发布日期</th><th>链接</th></tr>"
            for p in new_policies:
                body += f"<tr><td>{p['分类']}</td><td>{p['来源']}</td><td>{p['政策名称']}</td>" \
                        f"<td>{p['发布日期']}</td><td><a href='{p['原文链接']}'>查看</a></td></tr>"
            body += "</table>"
            subject = f"{mail_cfg['subject_prefix']}海关总署新增 {len(new_policies)} 条"
        else:
            body = "<p>今日无新增海关政策，附件为全部历史记录。</p>"
            subject = f"{mail_cfg['subject_prefix']}海关总署今日无新政策"

        msg = MIMEMultipart()
        msg['From'] = mail_cfg["sender"]
        msg['To'] = ", ".join(mail_cfg["recipients"])
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'html', 'utf-8'))

        if os.path.exists(xlsx_path):
            with open(xlsx_path, 'rb') as f:
                att = MIMEApplication(f.read(), _subtype='xlsx')
                att.add_header('Content-Disposition', 'attachment', filename=os.path.basename(xlsx_path))
                msg.attach(att)

        if mail_cfg["use_ssl"]:
            server = smtplib.SMTP_SSL(mail_cfg["smtp_server"], mail_cfg["smtp_port"])
        else:
            server = smtplib.SMTP(mail_cfg["smtp_server"], mail_cfg["smtp_port"])
            server.starttls()
        server.login(mail_cfg["sender"], mail_cfg["password"])
        server.sendmail(mail_cfg["sender"], mail_cfg["recipients"], msg.as_string())
        server.quit()
        print(f"\n>>> 邮件已发送至 {', '.join(mail_cfg['recipients'])}")

    except Exception as e:
        print(f"[异常] 邮件发送失败: {e}")

def main():
    cache_path = get_cache_path()
    cache = PolicyCache(cache_path)

    new_policies = crawl_all(cache)

    cache.save()
    xlsx_path = get_history_xlsx_path()
    cache.export_excel(xlsx_path)

    send_email(new_policies, xlsx_path)

if __name__ == '__main__':
    main()
