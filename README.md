# VOD Comment Scraper

从 VOD 数据追踪表（Excel）中批量抓取 **YouTube + TikTok** 视频评论，自动生成**情感分析 + 玩家偏好** Excel 报告。

Scrape YouTube + TikTok video comments from a VOD tracking spreadsheet and generate a **sentiment-analysis + player-preference** Excel report.

---

## 流程 Pipeline

```
Excel 数据追踪表
   │  ① extract      提取视频链接
   ▼
data/video_links.json
   │  ② scrape-yt    YouTube 评论 (yt-dlp, 免费)
   │  ③ scrape-tt    TikTok 评论  (Apify apidojo actor, 付费)
   ▼
data/*.json
   │  ④ analyze      VADER 情感分析 + 关键词提取 + Excel 报告
   ▼
评论原文及总结.xlsx
```

## 环境要求 Requirements

- Python 3.9+
- 一个 [Apify](https://apify.com) 账号（TikTok 抓取用，Free plan 可试用，SCALE/Starter 按量付费约 **$0.30 / 1000 条评论**）
- YouTube 抓取**免费**，走 yt-dlp，无需 API key

## 安装 Install

```bash
git clone <your-repo-url>
cd vod-comment-scraper

python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS / Linux:
# source .venv/bin/activate

pip install -r requirements.txt
```

配置 Apify token（只需一次）：

```bash
cp .env.example .env
# 编辑 .env，填入你的 APIFY_API_KEY
# 获取地址: https://console.apify.com/settings/integrations
```

## 使用 Usage

### 一键全流程

```bash
python main.py all \
    --excel "我的数据追踪表.xlsx" \
    --title "My Game" \
    --output "data/MyGame_评论报告.xlsx" \
    --base-excel "我的数据追踪表.xlsx"   # 可选: 报告追加到原表副本里
```

### 分步执行

```bash
# ① 从 Excel 提取视频链接（表结构不同见下方"自定义表结构"）
python main.py extract --excel "我的数据追踪表.xlsx"

# ② 抓 YouTube 评论（免费，增量可断点续跑）
python main.py scrape-yt

# ③ 抓 TikTok 评论（需要 APIFY_API_KEY，增量可断点续跑）
python main.py scrape-tt

# ④ 生成情感分析报告
python main.py analyze --output "data/报告.xlsx" --title "My Game"
```

所有中间产物存在 `data/` 目录（已被 .gitignore 忽略，不会误提交）。

## Excel 表结构要求 Sheet Layout

默认适配「数据追踪（TT）/ 数据追踪（YT）」两个 sheet：

| 列 | 内容 |
|----|------|
| B  | 账号名 |
| D  | 视频链接 |
| E  | 发布时间 |
| F  | 观看量 |
| G  | 评论数（用于筛选值得抓取的视频） |
| H  | 点赞 |
| I  | 转发 |

数据从第 3 行开始。表结构不同时可传参调整：

```bash
python main.py extract --excel x.xlsx \
    --tt-sheet "TikTok" --yt-sheet "YouTube" \
    --col-link C --start-row 2
```

## 输出报告 Output

生成的 Excel「评论原文及总结」sheet 包含：

1. **评论原文** — 每条评论 + VADER 情感倾向（正面/负面/中性，绿/红/黄底色）+ 情感分值
2. **按视频汇总** — 每个视频的正/负/中性占比、玩家偏好关键词、代表性正负面观点
3. **整体总结** — 全量评论情感分布、高频关注词、代表性正负面观点

## 关键技术说明 Notes

- **TikTok 评论 API 是 auth-gated 的**：未登录请求永远返回空 200。所以本工具用 Apify 的 `apidojo~tiktok-comments-scraper` actor（不需要登录态/cookies/住宅代理），实测覆盖率 ~98%。**不要**换成需要登录态的 actor（如 clockworks），会有 ~90% 视频返回 NO_RETURN。
- **增量可恢复**：抓取脚本随时可以 Ctrl+C 中断，重跑会自动跳过已抓取的视频/评论（按 comment_id 去重）。
- **跳过逻辑**：评论数为 0 的视频不抓；YouTube 0 播放的视频不抓。
- 情感分析基于 VADER（针对英文社媒文本优化）+ 游戏/玩家 slang 词典加权。中文评论效果有限，如需中文情感分析可自行接入其他模型。

## 成本 Cost

| 平台 | 方式 | 成本 |
|------|------|------|
| YouTube | yt-dlp | 免费 |
| TikTok | Apify apidojo actor | ~$0.30 / 1,000 条评论 |

参考实测：233 个 TikTok 视频（102 个有评论）→ 384 条评论，几美分级别。

## 常见问题 FAQ

**Q: TikTok 抓取报 `APIFY_API_KEY not set`？**
A: 没有配置 token。`cp .env.example .env` 然后填入你的 token。

**Q: 有的视频评论数为正但抓不到？**
A: 少数视频是 TikTok 侧敏感内容分级（POST_SENSITIVE），任何 actor 都拿不到，属正常损耗，会记录在 `data/tiktok_skip.json`。

**Q: 想同时抓评论回复？**
A: `python main.py scrape-tt --include-replies`（评论量会显著增加，注意成本）。

## License

MIT
