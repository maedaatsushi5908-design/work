"""
爆サイ掲示板スクレイパー
- 指定したカテゴリ・スレッドのテキストを一括取得
- AI入力用にテキストファイルへ出力
"""

import time
import re
import os
import json
import argparse
from urllib.parse import urljoin, urlparse
from dataclasses import dataclass, field
from typing import Optional

import requests
from bs4 import BeautifulSoup

BASE_URL = "https://bakusai.com"
DEFAULT_DELAY = 2.0  # サーバー負荷軽減のためリクエスト間隔（秒）
HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; BakusaiResearchBot/1.0)",
    "Accept-Language": "ja,en;q=0.9",
}


@dataclass
class Post:
    number: int
    author: str
    datetime: str
    body: str


@dataclass
class Thread:
    title: str
    url: str
    board: str
    posts: list[Post] = field(default_factory=list)


def fetch(url: str, session: requests.Session, delay: float = DEFAULT_DELAY) -> Optional[BeautifulSoup]:
    """GETリクエストを送りBeautifulSoupを返す。失敗時はNone。"""
    try:
        resp = session.get(url, headers=HEADERS, timeout=15)
        resp.raise_for_status()
        resp.encoding = resp.apparent_encoding or "utf-8"
        time.sleep(delay)
        return BeautifulSoup(resp.text, "html.parser")
    except requests.RequestException as e:
        print(f"[ERROR] {url}: {e}")
        return None


def get_thread_list(board_url: str, session: requests.Session, max_threads: int = 20) -> list[dict]:
    """掲示板トップからスレッド一覧を取得する。"""
    threads = []
    page = 1

    while len(threads) < max_threads:
        url = board_url if page == 1 else f"{board_url}?p={page}"
        soup = fetch(url, session)
        if soup is None:
            break

        items = soup.select("a[href*='/bbs/']")
        found = 0
        for a in items:
            href = a.get("href", "")
            # スレッド個別URLパターン: /bbs/***/***/
            if re.search(r"/bbs/[^/]+/[^/]+/\d+", href):
                full_url = urljoin(BASE_URL, href)
                title = a.get_text(strip=True)
                if full_url not in [t["url"] for t in threads] and title:
                    threads.append({"url": full_url, "title": title})
                    found += 1
                    if len(threads) >= max_threads:
                        break

        if found == 0:
            break
        page += 1

    return threads[:max_threads]


def scrape_thread(thread_url: str, session: requests.Session, max_pages: int = 5) -> Thread:
    """スレッドの全投稿を取得する。"""
    soup = fetch(thread_url, session)
    if soup is None:
        return Thread(title="", url=thread_url, board="")

    title_tag = soup.select_one("h1, .thread-title, #thread_title")
    title = title_tag.get_text(strip=True) if title_tag else thread_url

    board = urlparse(thread_url).path.split("/")[2] if "/" in urlparse(thread_url).path else ""
    thread = Thread(title=title, url=thread_url, board=board)

    for page_num in range(1, max_pages + 1):
        url = thread_url if page_num == 1 else f"{thread_url}?p={page_num}"
        page_soup = fetch(url, session) if page_num > 1 else soup

        if page_soup is None:
            break

        posts = _parse_posts(page_soup)
        if not posts:
            break

        # 重複投稿をスキップ
        existing_nums = {p.number for p in thread.posts}
        new_posts = [p for p in posts if p.number not in existing_nums]
        thread.posts.extend(new_posts)

        # 次ページリンクがなければ終了
        next_link = page_soup.select_one("a.next, a[rel='next'], .pager_next a")
        if not next_link:
            break

    return thread


def _parse_posts(soup: BeautifulSoup) -> list[Post]:
    """ページから投稿一覧をパースする。"""
    posts = []

    # 爆サイの投稿コンテナを複数パターンで試みる
    containers = soup.select(".res_box, .post, article.post, .bbs_res")
    if not containers:
        # フォールバック: レス番号を含む要素を探す
        containers = soup.select("[id^='res'], [id^='post']")

    for i, container in enumerate(containers, 1):
        # レス番号
        num_tag = container.select_one(".res_num, .post-num, .number")
        num = int(num_tag.get_text(strip=True).strip(">>>#No.")) if num_tag else i

        # 投稿者
        author_tag = container.select_one(".res_name, .author, .name")
        author = author_tag.get_text(strip=True) if author_tag else "名無し"

        # 日時
        dt_tag = container.select_one(".res_date, .datetime, time")
        dt = dt_tag.get_text(strip=True) if dt_tag else ""

        # 本文
        body_tag = container.select_one(".res_body, .post-body, .body, p")
        body = body_tag.get_text(separator="\n", strip=True) if body_tag else container.get_text(strip=True)

        if body:
            posts.append(Post(number=num, author=author, datetime=dt, body=body))

    return posts


def thread_to_text(thread: Thread) -> str:
    """スレッドをAI入力向けのプレーンテキストに変換する。"""
    lines = [
        f"【スレッド】{thread.title}",
        f"【URL】{thread.url}",
        f"【板】{thread.board}",
        f"【投稿数】{len(thread.posts)}",
        "=" * 60,
    ]
    for post in thread.posts:
        lines.append(f"\n[{post.number}] {post.author} {post.datetime}")
        lines.append(post.body)

    return "\n".join(lines)


def save_text(thread: Thread, output_dir: str = "output") -> str:
    """テキストファイルとして保存し、ファイルパスを返す。"""
    os.makedirs(output_dir, exist_ok=True)
    safe_title = re.sub(r'[\\/*?:"<>|]', "_", thread.title)[:60]
    filename = f"{thread.board}_{safe_title}.txt"
    path = os.path.join(output_dir, filename)
    with open(path, "w", encoding="utf-8") as f:
        f.write(thread_to_text(thread))
    return path


def save_json(thread: Thread, output_dir: str = "output") -> str:
    """JSON形式で保存し、ファイルパスを返す。"""
    os.makedirs(output_dir, exist_ok=True)
    safe_title = re.sub(r'[\\/*?:"<>|]', "_", thread.title)[:60]
    filename = f"{thread.board}_{safe_title}.json"
    path = os.path.join(output_dir, filename)
    data = {
        "title": thread.title,
        "url": thread.url,
        "board": thread.board,
        "posts": [
            {"number": p.number, "author": p.author, "datetime": p.datetime, "body": p.body}
            for p in thread.posts
        ],
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return path


def scrape_board(board_url: str, max_threads: int = 10, max_pages: int = 3,
                 output_dir: str = "output", fmt: str = "text") -> list[str]:
    """掲示板URLから複数スレッドを一括取得し保存する。保存ファイルパスのリストを返す。"""
    session = requests.Session()
    print(f"[INFO] 板URL: {board_url}")
    print(f"[INFO] スレッド取得上限: {max_threads}")

    thread_list = get_thread_list(board_url, session, max_threads)
    print(f"[INFO] {len(thread_list)} スレッド見つかりました")

    saved = []
    for i, t in enumerate(thread_list, 1):
        print(f"[{i}/{len(thread_list)}] 取得中: {t['title'][:50]}")
        thread = scrape_thread(t["url"], session, max_pages)
        if not thread.posts:
            print(f"  → 投稿なし、スキップ")
            continue
        if fmt == "json":
            path = save_json(thread, output_dir)
        else:
            path = save_text(thread, output_dir)
        print(f"  → 保存: {path} ({len(thread.posts)}件)")
        saved.append(path)

    return saved


def merge_for_ai(saved_paths: list[str], output_path: str = "output/merged_for_ai.txt",
                 max_chars: int = 200_000) -> str:
    """複数ファイルを1つに結合してAIに渡しやすい形式にする。"""
    merged = []
    total = 0
    for path in saved_paths:
        with open(path, encoding="utf-8") as f:
            content = f.read()
        if total + len(content) > max_chars:
            print(f"[INFO] 文字数上限({max_chars:,})に達したため以降をスキップ")
            break
        merged.append(content)
        total += len(content)

    result = "\n\n" + "=" * 60 + "\n\n".join(merged)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(result)
    print(f"[INFO] マージ完了: {output_path} ({total:,}文字)")
    return output_path


def main():
    parser = argparse.ArgumentParser(description="爆サイ掲示板スクレイパー")
    parser.add_argument("board_url", help="板のURL (例: https://bakusai.com/bbs_list/cat/...)")
    parser.add_argument("--threads", type=int, default=10, help="取得スレッド数 (default: 10)")
    parser.add_argument("--pages", type=int, default=3, help="スレッドごとのページ数 (default: 3)")
    parser.add_argument("--output", default="output", help="出力ディレクトリ (default: output)")
    parser.add_argument("--format", choices=["text", "json"], default="text", help="出力形式")
    parser.add_argument("--merge", action="store_true", help="全テキストを1ファイルにまとめる")
    args = parser.parse_args()

    saved = scrape_board(
        board_url=args.board_url,
        max_threads=args.threads,
        max_pages=args.pages,
        output_dir=args.output,
        fmt=args.format,
    )

    if args.merge and saved:
        merge_for_ai(saved, output_path=os.path.join(args.output, "merged_for_ai.txt"))

    print(f"\n完了: {len(saved)} ファイル保存")


if __name__ == "__main__":
    main()
