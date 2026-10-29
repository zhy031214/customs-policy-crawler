# customs-policy-crawler

海关总署政策监控爬虫：定时抓取 [海关总署政务公开·最新文件](http://www.customs.gov.cn/customs/302249/2480148/index.html) 栏目的新政策，发送邮件提醒并维护发布历史 Excel。

架构与触发方式同 [policy-crawler（商务部）](https://github.com/zhy031214/policy-crawler)。

## 输出

- `policy_cache.json`：已见政策缓存（URL+标题 MD5 去重）
- `发布历史(from_2025.10).xlsx`：全部历史记录，链接列为可点击超链接
- 邮件：QQ 邮箱 SMTP 推送本次新增政策表格 + 历史 Excel 附件

## 反爬说明

customs.gov.cn 部署了加速乐（Jiasule）WAF，直接 requests 会被 412 拦截。
本爬虫采用混合方案：

1. Playwright（无头 Chromium）先访问首页完成 JS 挑战，拿到全套 WAF cookie
2. 之后用 requests 携带 cookie 抓取栏目列表页（eportal 分页接口）

## 运行

```bash
pip install requests beautifulsoup4 pandas openpyxl playwright
python -m playwright install chromium
EMAIL_PASSWORD=xxx python crawler_github.py
```

- 首次运行：抓取最近 1 年（`first_start_date` 起，默认 2025-10-29）的全部政策
- 日常运行：仅抓取最近 `days_limit`（默认 1）天的新政策，与缓存比对去重

## GitHub Actions

`.github/workflows/crawl.yml`：每天 UTC 02:00（北京时间 10:00）自动运行，也可手动触发。
需要在仓库 Settings → Secrets 配置 `EMAIL_PASSWORD`（QQ 邮箱授权码）。
