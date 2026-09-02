#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
VOD Comment Scraper - unified CLI
=================================
Pipeline for scraping YouTube / TikTok comments from a VOD tracking
spreadsheet and generating a sentiment-analysis Excel report.

Usage:
    python main.py extract    --excel <tracking.xlsx>
    python main.py scrape-yt
    python main.py scrape-tt
    python main.py analyze    --output data/report.xlsx --title "My Game"
    python main.py all        --excel <tracking.xlsx> --title "My Game" --output data/report.xlsx
"""
import argparse
import os
import sys

# make prints survive GBK consoles (Windows non-UTF8 terminals)
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# load .env if present (python-dotenv is optional)
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

import extract_links
import scrape_tt
import scrape_yt
import analyze_export

DATA = "data"


def cmd_extract(args):
    extract_links.run(args.excel, args.tt_sheet, args.yt_sheet, args.out,
                     args.col_account, args.col_link, args.col_date,
                     args.col_views, args.col_comments, args.col_likes,
                     args.col_shares, args.start_row)


def cmd_scrape_tt(args):
    scrape_tt.run(args.links, args.out, args.skip, args.batch,
                  args.max_items, args.include_replies)


def cmd_scrape_yt(args):
    scrape_yt.run(args.links, args.out, args.delay)


def cmd_analyze(args):
    analyze_export.run(args.yt_comments, args.tt_comments, args.output,
                       args.base_excel, args.title, args.sheet)


def cmd_all(args):
    links = os.path.join(DATA, "video_links.json")
    extract_links.run(args.excel, args.tt_sheet, args.yt_sheet, links)
    scrape_yt.run(links, os.path.join(DATA, "youtube_comments.json"))
    scrape_tt.run(links, os.path.join(DATA, "tiktok_comments.json"),
                  os.path.join(DATA, "tiktok_skip.json"))
    analyze_export.run(os.path.join(DATA, "youtube_comments.json"),
                       os.path.join(DATA, "tiktok_comments.json"),
                       args.output, args.base_excel, args.title, args.sheet)


def add_extract_options(p):
    p.add_argument("--excel", required=True, help="path to the VOD tracking .xlsx")
    p.add_argument("--tt-sheet", default="数据追踪（TT）", help="TikTok sheet name")
    p.add_argument("--yt-sheet", default="数据追踪（YT）", help="YouTube sheet name")
    p.add_argument("--out", default=os.path.join(DATA, "video_links.json"))
    p.add_argument("--col-account", default="B")
    p.add_argument("--col-link", default="D")
    p.add_argument("--col-date", default="E")
    p.add_argument("--col-views", default="F")
    p.add_argument("--col-comments", default="G")
    p.add_argument("--col-likes", default="H")
    p.add_argument("--col-shares", default="I")
    p.add_argument("--start-row", type=int, default=3)


def main():
    ap = argparse.ArgumentParser(
        prog="vod-comment-scraper",
        description="Scrape YouTube/TikTok comments from a VOD tracking sheet "
                    "and generate a sentiment-analysis Excel report.")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("extract", help="extract video links from tracking Excel")
    add_extract_options(p)
    p.set_defaults(fn=cmd_extract)

    p = sub.add_parser("scrape-tt", help="scrape TikTok comments via Apify (apidojo)")
    p.add_argument("--links", default=os.path.join(DATA, "video_links.json"))
    p.add_argument("--out", default=os.path.join(DATA, "tiktok_comments.json"))
    p.add_argument("--skip", default=os.path.join(DATA, "tiktok_skip.json"))
    p.add_argument("--batch", type=int, default=20)
    p.add_argument("--max-items", type=int, default=5000)
    p.add_argument("--include-replies", action="store_true")
    p.set_defaults(fn=cmd_scrape_tt)

    p = sub.add_parser("scrape-yt", help="scrape YouTube comments via yt-dlp (free)")
    p.add_argument("--links", default=os.path.join(DATA, "video_links.json"))
    p.add_argument("--out", default=os.path.join(DATA, "youtube_comments.json"))
    p.add_argument("--delay", type=float, default=0.3)
    p.set_defaults(fn=cmd_scrape_yt)

    p = sub.add_parser("analyze", help="sentiment analysis + Excel report")
    p.add_argument("--yt-comments", default=os.path.join(DATA, "youtube_comments.json"))
    p.add_argument("--tt-comments", default=os.path.join(DATA, "tiktok_comments.json"))
    p.add_argument("--output", required=True, help="output .xlsx path")
    p.add_argument("--base-excel", default=None,
                   help="optional: copy this workbook and append the sheet to it")
    p.add_argument("--title", default="Game")
    p.add_argument("--sheet", default="评论原文及总结")
    p.set_defaults(fn=cmd_analyze)

    p = sub.add_parser("all", help="run the full pipeline in one go")
    add_extract_options(p)
    p.add_argument("--output", required=True, help="output .xlsx path")
    p.add_argument("--base-excel", default=None)
    p.add_argument("--title", default="Game")
    p.set_defaults(fn=cmd_all)

    args = ap.parse_args()
    args.fn(args)


if __name__ == "__main__":
    main()
