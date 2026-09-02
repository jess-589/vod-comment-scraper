#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Step 2b: Scrape YouTube comments via yt-dlp (free, no API key needed).

Only processes videos with views > 0 (zero-view videos essentially never
have comments). Skips deleted/violation entries.

Reads:  data/video_links.json  (produced by extract_links.py)
Writes: data/youtube_comments.json  (incremental, resumable)
"""
import argparse
import json
import os
import time


def is_valid_url(url):
    if not url or not url.strip():
        return False
    url = url.strip()
    if "违规" in url or "删除" in url:
        return False
    return "youtube.com" in url or "youtu.be" in url


def run(links_path, out_path, delay=0.3):
    import yt_dlp

    with open(links_path, encoding="utf-8") as f:
        data = json.load(f)
    videos = data.get("youtube", [])

    # incremental resume
    existing = []
    if os.path.exists(out_path):
        with open(out_path, encoding="utf-8") as f:
            existing = json.load(f)
    done_urls = {c.get("video_url") for c in existing}

    seen = set()
    unique_videos = []
    for v in videos:
        url = (v.get("link") or "").strip()
        if (url and url not in seen and is_valid_url(url)
                and url not in done_urls
                and isinstance(v.get("views", 0), (int, float))
                and v["views"] > 0):
            seen.add(url)
            unique_videos.append(v)

    skipped = len(videos) - len(unique_videos)
    print(f"[scrape-yt] total: {len(videos)}, to process: {len(unique_videos)} "
          f"({skipped} skipped: 0 views / invalid / already done)")

    ydl_opts = {
        "getcomments": True,
        "skip_download": True,
        "quiet": True,
        "writeinfojson": False,
        "extract_flat": False,
    }

    for i, video in enumerate(unique_videos):
        url = video["link"].strip()
        account = video.get("account", "")
        print(f"[scrape-yt] [{i + 1}/{len(unique_videos)}] {account} | {url[:70]}")
        try:
            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                info = ydl.extract_info(url, download=False)
            if not info:
                print("[scrape-yt]   [!] no video info")
                continue

            comments = [c for c in (info.get("comments") or [])
                        if (c.get("text") or "").strip()]
            if comments:
                for c in comments:
                    existing.append({
                        "platform": "YouTube",
                        "account": account,
                        "video_title": info.get("title", ""),
                        "video_url": url,
                        "video_id": info.get("id", ""),
                        "comment": c["text"].strip(),
                        "author": c.get("author", ""),
                        "likes": c.get("votes", 0),
                        "created": str(c.get("timestamp", "")),
                    })
                print(f"[scrape-yt]   -> {len(comments)} comments")
                with open(out_path, "w", encoding="utf-8") as f:
                    json.dump(existing, f, ensure_ascii=False, indent=2)
            else:
                print("[scrape-yt]   -> no comments")
        except Exception as e:
            msg = str(e)
            print(f"[scrape-yt]   [!] {msg[:80]}")

        if i < len(unique_videos) - 1:
            time.sleep(delay)

    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)
    print(f"[scrape-yt] DONE. total comments: {len(existing)} -> {out_path}")
    return existing


if __name__ == "__main__":
    p = argparse.ArgumentParser(description="Scrape YouTube comments via yt-dlp")
    p.add_argument("--links", default=os.path.join("data", "video_links.json"))
    p.add_argument("--out", default=os.path.join("data", "youtube_comments.json"))
    p.add_argument("--delay", type=float, default=0.3)
    args = p.parse_args()
    run(args.links, args.out, args.delay)
