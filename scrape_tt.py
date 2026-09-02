#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Step 2a: Scrape TikTok comments via Apify actor `apidojo~tiktok-comments-scraper`.

Why this actor: the TikTok comment API is auth-gated (logged-out requests
always return an empty HTTP 200 regardless of IP). This actor uses a method
that does not require login state, so it works reliably (~$0.30 per 1,000
comments). Requires an Apify account + API token.

Reads:  data/video_links.json  (produced by extract_links.py)
Writes: data/tiktok_comments.json  (incremental, resumable)
        data/tiktok_skip.json      (videos that returned nothing)
"""
import argparse
import json
import os
import time

import requests

ACTOR = "apidojo~tiktok-comments-scraper"
BASE = "https://api.apify.com/v2"


def norm_url(u):
    return (u or "").split("?")[0].rstrip("/").lower()


def run(links_path, out_path, skip_path,
        batch=20, max_items=5000, include_replies=False,
        run_timeout=300, http_timeout=360):
    key = os.environ.get("APIFY_API_KEY", "")
    if not key:
        raise SystemExit("[scrape-tt] ERROR: APIFY_API_KEY not set. "
                         "Copy .env.example to .env and fill in your token.")

    with open(links_path, encoding="utf-8") as f:
        data = json.load(f)

    # only videos that actually have comments, deduped by clean URL
    seen = {}
    for v in data.get("tiktok", []):
        if (v.get("comment_count") or 0) > 0:
            u = v["link"].split("?")[0]
            if u not in seen:
                v["clean_url"] = u
                seen[u] = v
    videos = list(seen.values())

    # incremental resume
    existing = []
    if os.path.exists(out_path):
        with open(out_path, encoding="utf-8") as f:
            existing = json.load(f)
    done_urls = {norm_url(c.get("video_url")) for c in existing}
    todo = [v for v in videos if norm_url(v["clean_url"]) not in done_urls]

    print(f"[scrape-tt] videos with comments: {len(videos)}, "
          f"already done: {len(done_urls)}, todo: {len(todo)}")
    if not todo:
        print("[scrape-tt] nothing to do.")
        return existing

    skip_new = {}
    total_batches = (len(todo) + batch - 1) // batch

    for bi in range(0, len(todo), batch):
        bvideos = todo[bi:bi + batch]
        urls = [v["clean_url"] for v in bvideos]
        batch_num = bi // batch + 1
        payload = {
            "startUrls": urls,
            "includeReplies": include_replies,
            "maxItems": max_items,
        }
        print(f"[scrape-tt] batch {batch_num}/{total_batches}: {len(urls)} videos ...")

        t0 = time.time()
        try:
            rr = requests.post(
                f"{BASE}/acts/{ACTOR}/run-sync-get-dataset-items",
                params={"token": key, "timeout": run_timeout},
                json=payload, timeout=http_timeout)
        except Exception as e:
            print(f"[scrape-tt]   request error: {e}, retrying in 5s")
            time.sleep(5)
            continue

        try:
            items = rr.json()
            if not isinstance(items, list):
                items = []
        except Exception:
            print(f"[scrape-tt]   JSON parse error, head: {rr.text[:200]}")
            continue

        meta_by_url = {norm_url(v["clean_url"]): v for v in bvideos}
        added = 0
        for it in items:
            if not isinstance(it, dict):
                continue
            text = (it.get("text") or "").strip()
            if not text:
                continue
            vu = it.get("inputSource") or ""
            meta = meta_by_url.get(norm_url(vu))
            user = it.get("user") or {}
            existing.append({
                "platform": "TikTok",
                "account": meta["account"] if meta else "",
                "video_title": "",
                "video_url": meta["clean_url"] if meta else vu,
                "comment": text,
                "comment_id": it.get("id", ""),
                "author": user.get("username", ""),
                "likes": it.get("likeCount", 0),
                "created": it.get("createdAt", ""),
            })
            added += 1

        # record videos that returned nothing
        returned = {norm_url(it.get("inputSource"))
                    for it in items
                    if isinstance(it, dict) and it.get("inputSource")}
        for v in bvideos:
            if norm_url(v["clean_url"]) not in returned:
                skip_new[v["clean_url"]] = "NO_RETURN_APIDOJO"

        # dedup by comment_id
        seen_ids = set()
        deduped = []
        for c in existing:
            cid = c.get("comment_id") or c.get("comment")
            if cid in seen_ids:
                continue
            seen_ids.add(cid)
            deduped.append(c)
        existing = deduped

        # incremental save
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(existing, f, ensure_ascii=False, indent=2)

        print(f"[scrape-tt]   status={rr.status_code} "
              f"elapsed={round(time.time() - t0, 1)}s +{added} comments")
        time.sleep(1)

    # merge skip log with any previous one
    all_skip = {}
    if os.path.exists(skip_path):
        with open(skip_path, encoding="utf-8") as f:
            all_skip = json.load(f)
    all_skip.update(skip_new)
    with open(skip_path, "w", encoding="utf-8") as f:
        json.dump(all_skip, f, ensure_ascii=False, indent=2)

    print(f"[scrape-tt] DONE. total comments: {len(existing)}, "
          f"videos still returning nothing: {len(all_skip)}")
    return existing


if __name__ == "__main__":
    p = argparse.ArgumentParser(description="Scrape TikTok comments via Apify (apidojo)")
    p.add_argument("--links", default=os.path.join("data", "video_links.json"))
    p.add_argument("--out", default=os.path.join("data", "tiktok_comments.json"))
    p.add_argument("--skip", default=os.path.join("data", "tiktok_skip.json"))
    p.add_argument("--batch", type=int, default=20, help="videos per Apify run")
    p.add_argument("--max-items", type=int, default=5000, help="max comments per run")
    p.add_argument("--include-replies", action="store_true")
    args = p.parse_args()
    run(args.links, args.out, args.skip, args.batch,
        args.max_items, args.include_replies)
