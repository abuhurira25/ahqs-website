#!/usr/bin/env python3
"""Maintain AHQS discovery metadata and notify IndexNow after publication.

IndexNow uses a site-hosted ownership proof, generated here, not an account
password or third-party-issued secret. Never print the proof in logs.
"""
import argparse
from datetime import datetime, timezone
from html.parser import HTMLParser
import json
from pathlib import Path
import re
import secrets
import subprocess
import time
from urllib.parse import urlparse
from urllib.request import Request, urlopen
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[2]
ORIGIN = "https://ahqshealthcare.com"
KEY_FILE = ROOT / "ahqs-indexnow.txt"
ENDPOINT = "https://api.indexnow.org/indexnow"
NS = "http://www.sitemaps.org/schemas/sitemap/0.9"


def git(*args):
    return subprocess.check_output(["git", "-C", str(ROOT), *args], text=True).strip()


class Metadata(HTMLParser):
    def __init__(self, text):
        super().__init__()
        self.canonical = []
        self.robots = []
        self.descriptions = []
        self.feed(text)

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag == "link" and "canonical" in a.get("rel", "").split():
            self.canonical.append(a.get("href", ""))
        if tag == "meta":
            name = a.get("name", "").lower()
            if name == "robots":
                self.robots.append(a.get("content", "").lower())
            if name == "description":
                self.descriptions.append(a.get("content", ""))


def public_url(filename):
    if filename == "index.html":
        return ORIGIN + "/"
    if filename.endswith("/index.html"):
        return ORIGIN + "/" + filename[:-10]
    return ORIGIN + "/" + filename


def pages():
    result = {}
    for filename in git("ls-files", "*.html").splitlines():
        filepath = ROOT / filename
        if not filepath.exists():
            continue
        meta = Metadata(filepath.read_text())
        if len(meta.robots) > 1:
            raise ValueError(f"Multiple robots directives: {filename}")
        if any("noindex" in r for r in meta.robots):
            continue
        if len(meta.canonical) != 1:
            raise ValueError(f"Expected one canonical: {filename}")
        url = meta.canonical[0]
        if not url.startswith(ORIGIN + "/"):
            raise ValueError(f"Foreign canonical: {filename}")
        if url != public_url(filename):
            target = urlparse(url).path.lstrip("/") or "index.html"
            if target.endswith("/"):
                target += "index.html"
            if not (ROOT / target).is_file():
                raise ValueError(f"Canonical target not found: {filename}")
            continue  # consolidate alternate URL; do not submit duplicate
        if len(meta.descriptions) != 1 or not meta.descriptions[0]:
            raise ValueError(f"Missing/duplicate description: {filename}")
        result[filename] = url
    return result


def build_sitemap():
    ET.register_namespace("", NS)
    root = ET.Element(f"{{{NS}}}urlset")
    modified = set(git("diff", "--name-only", "HEAD").splitlines())
    for filename, url in sorted(pages().items(), key=lambda item: item[1]):
        item = ET.SubElement(root, f"{{{NS}}}url")
        ET.SubElement(item, f"{{{NS}}}loc").text = url
        lastmod = git("log", "-1", "--format=%cs", "--", filename)
        if filename in modified:
            lastmod = datetime.now(timezone.utc).date().isoformat()
        if lastmod:
            ET.SubElement(item, f"{{{NS}}}lastmod").text = lastmod
    ET.indent(root, space="  ")
    ET.ElementTree(root).write(ROOT / "sitemap.xml", encoding="UTF-8", xml_declaration=True)
    print(f"Sitemap generated: {len(root)} canonical, indexable pages.")


def validate():
    expected = set(pages().values())
    root = ET.parse(ROOT / "sitemap.xml").getroot()
    actual = [x.text for x in root.findall(f"{{{NS}}}url/{{{NS}}}loc")]
    if len(actual) != len(set(actual)) or set(actual) != expected:
        raise ValueError("Sitemap URLs do not match indexable canonical pages.")
    if f"Sitemap: {ORIGIN}/sitemap.xml" not in (ROOT / "robots.txt").read_text():
        raise ValueError("Missing robots sitemap reference.")
    if KEY_FILE.exists() and not re.fullmatch(r"[a-zA-Z0-9-]{8,128}", KEY_FILE.read_text().strip()):
        raise ValueError("Invalid IndexNow ownership proof.")
    print(f"PASS: {len(actual)} canonical URLs, metadata and robots sitemap reference.")


def select_urls(all_urls):
    current = pages()
    if all_urls:
        return sorted(current.values())
    changed = git("diff-tree", "--no-commit-id", "--name-only", "-r", "HEAD").splitlines()
    if KEY_FILE.name in changed:
        return sorted(current.values())  # initial proof publication or rotation
    # Shared styles, scripts and images can change every rendered page.
    if any(p.startswith("assets/") or p == "sw.js" or p.endswith((".css", ".png", ".jpg", ".jpeg")) for p in changed):
        return sorted(current.values())
    selected = {current[p] for p in changed if p in current}
    # Include newly discovered sitemap entries but avoid resending unchanged URLs.
    if "sitemap.xml" in changed:
        try:
            previous = ET.fromstring(git("show", "HEAD^:sitemap.xml"))
            old_urls = {x.text for x in previous.findall(f"{{{NS}}}url/{{{NS}}}loc")}
            selected |= set(current.values()) - old_urls
        except subprocess.CalledProcessError:
            selected |= set(current.values())
    # Notify removed/noindexed pages so a search engine can revisit their state.
    for filename in changed:
        if filename.endswith(".html") and filename not in current:
            try:
                old = Metadata(git("show", "HEAD^:" + filename))
                selected |= {u for u in old.canonical if u.startswith(ORIGIN + "/")}
            except subprocess.CalledProcessError:
                pass
    return sorted(selected)


def submit(all_urls=False, dry_run=False):
    urls = select_urls(all_urls)
    if not urls:
        print("No indexable page changes to notify.")
        return
    if dry_run:
        print(json.dumps({"endpoint": ENDPOINT, "url_count": len(urls), "urls": urls}, indent=2))
        return
    key = KEY_FILE.read_text().strip()
    key_location = ORIGIN + "/" + KEY_FILE.name
    # Confirm the exact proof is deployed before sending any notification.
    for attempt in range(6):
        try:
            with urlopen(key_location, timeout=25) as response:
                valid = response.status == 200 and response.read().decode().strip() == key
            if valid:
                break
        except Exception:
            valid = False
        if attempt == 5:
            raise RuntimeError("Live ownership proof is not available; nothing submitted.")
        time.sleep(15)
    payload = {"host": urlparse(ORIGIN).netloc, "key": key, "keyLocation": key_location, "urlList": urls}
    request = Request(ENDPOINT, data=json.dumps(payload).encode(), headers={"Content-Type": "application/json; charset=utf-8"}, method="POST")
    with urlopen(request, timeout=45) as response:
        code = response.status
        response.read()
    if code not in (200, 202):
        raise RuntimeError(f"Unexpected IndexNow HTTP response: {code}")
    status = "received" if code == 200 else "received; key validation pending"
    print(json.dumps({"engine_protocol": "IndexNow", "http_status": code, "status": status, "url_count": len(urls), "indexed": "not confirmed"}, indent=2))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("command", choices=["init", "sitemap", "validate", "submit"])
    parser.add_argument("--all", action="store_true", dest="all_urls")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()
    if args.command == "init":
        if not KEY_FILE.exists():
            KEY_FILE.write_text(secrets.token_hex(24) + "\n")
        print("Site-hosted IndexNow ownership proof ready.")
    elif args.command == "sitemap":
        build_sitemap()
    elif args.command == "validate":
        validate()
    else:
        submit(args.all_urls, args.dry_run)


if __name__ == "__main__":
    main()
