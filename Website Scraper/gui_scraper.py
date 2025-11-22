import asyncio
import threading
import logging
import random
import os
import time
from typing import Optional, Callable, Dict, Any, List, Tuple
import httpx # pyright: ignore[reportMissingImports]
import pandas as pd # pyright: ignore[reportMissingModuleSource]
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText
from playwright.async_api import async_playwright, Playwright, Browser, Page, Error as PWError # pyright: ignore[reportMissingImports]

# ---------- Config constants ----------
IMAGE_DOWNLOAD_CONCURRENCY = 6
IMAGE_DOWNLOAD_TIMEOUT = 12
IMAGE_HEAD_TIMEOUT = 8
PLAYWRIGHT_TIMEOUT = 60000  # ms
SCROLL_SLEEP_MIN = 0.35
SCROLL_SLEEP_MAX = 0.9
MAX_SCROLL_STAGNANT = 20
MAX_PAGINATION_PAGES_DEFAULT = 200
EXCEL_ENGINE = "openpyxl"
TEMP_EXCEL = "temp_output.xlsx"

# ---------- GUI log handler ----------
class TextHandler(logging.Handler):
    def __init__(self, widget: ScrolledText):
        super().__init__()
        self.widget = widget

    def emit(self, record: logging.LogRecord) -> None:
        msg = self.format(record)
        try:
            self.widget.configure(state="normal")
            self.widget.insert(tk.END, msg + "\n")
            self.widget.see(tk.END)
            self.widget.configure(state="disabled")
        except tk.TclError:
            pass

# ---------- Helpers ----------
def ensure_dir(folder: str):
    try:
        os.makedirs(folder, exist_ok=True)
    except Exception:
        pass

def safe_filename(name: str) -> str:
    return "".join(c if c.isalnum() or c in "-_." else "_" for c in name)[:120]

async def head_check_image(url: str, timeout: int = IMAGE_HEAD_TIMEOUT, retries: int = 2) -> str:
    if not url or not url.lower().startswith("http"):
        return "Missing"
    for attempt in range(retries):
        try:
            async with httpx.AsyncClient(timeout=timeout, follow_redirects=True) as client:
                r = await client.head(url)
                ct = r.headers.get("content-type", "").lower()
                if r.status_code == 200 and "image" in ct:
                    return "Valid"
                return "Broken"
        except Exception:
            await asyncio.sleep(0.2 + attempt * 0.2)
    return "Broken"

async def download_image_async(url: str, folder: str, name: str, timeout: int = IMAGE_DOWNLOAD_TIMEOUT, retries: int = 2) -> Optional[str]:
    if not url or not url.lower().startswith("http"):
        return None
    ensure_dir(folder)
    ext = url.split("?")[0].split(".")[-1]
    if "/" in ext or len(ext) == 0 or len(ext) > 6:
        ext = "jpg"
    filename = f"{safe_filename(name)}.{ext}"
    path = os.path.join(folder, filename)
    for attempt in range(retries):
        try:
            async with httpx.AsyncClient(timeout=timeout) as client:
                r = await client.get(url)
                if r.status_code == 200:
                    with open(path, "wb") as f:
                        f.write(r.content)
                    return path
                await asyncio.sleep(0.2 + attempt * 0.2)
        except Exception:
            await asyncio.sleep(0.2 + attempt * 0.2)
    return None


class UniversalEngine:
    def __init__(self, page: Page, logger: logging.Logger, stop_event: threading.Event):
        self.page = page
        self.logger = logger
        self.stop_event = stop_event
        self.seen_keys = set()

    async def detect_all_tables(self) -> int:
        try:
            return int(await self.page.evaluate("() => document.querySelectorAll('table').length"))
        except Exception:
            return 0

    async def detect_best_table_selector(self) -> Optional[str]:
        candidates = ["table.table", "table.data-table", "div.table-responsive table", "table"]
        for sel in candidates:
            try:
                if await self.page.locator(sel).count() > 0:
                    return sel
            except Exception:
                continue
        try:
            if await self.page.locator("table").count() > 0:
                return "table"
        except Exception:
            pass
        return None

    async def detect_next_button(self) -> Optional[Tuple[str, str]]:
        css_candidates = [
            "a[rel='next']", "a[aria-label='Next']", "button[aria-label='Next']",
            ".pagination a.next", ".pagination li.next a", ".next", ".pager-next",
            "button[title='Next']", "a[title='Next']", "a.next", ".pg-next", ".page-next"
        ]
        for sel in css_candidates:
            try:
                if await self.page.locator(sel).count() > 0:
                    return ("css", sel)
            except Exception:
                continue
        xpath_candidates = [
            "//a[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'next')]",
            "//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'next')]",
            "//a[contains(., '›') or contains(., '»') or contains(., '>')]",
            "//button[contains(., '›') or contains(., '»') or contains(., '>')]"
        ]
        for xp in xpath_candidates:
            try:
                el = self.page.locator(f"xpath={xp}").first
                if await el.count() > 0:
                    return ("xpath", xp)
            except Exception:
                continue
        return None

    async def extract_table_js(self, selector: str) -> Optional[Dict[str, Any]]:
        js = f"""
        () => {{
            const table = document.querySelector("{selector}");
            if(!table) return null;
            const heads = Array.from(table.querySelectorAll('thead th')).map(h=>h.innerText.trim());
            let headers = heads.slice();
            if(headers.length === 0) {{
                const first = table.querySelector('tbody tr');
                if(first) headers = Array.from(first.querySelectorAll('td')).map((td,i)=>'Column_'+(i+1));
            }}
            const rows = Array.from(table.querySelectorAll('tbody tr')).map(tr=>{{
                const cells = Array.from(tr.querySelectorAll('td')).map(td=>td.innerText.trim());
                const links = Array.from(tr.querySelectorAll('a')).map(a=>a.href);
                const imgs = Array.from(tr.querySelectorAll('img')).map(i=>i.src);
                return {{cells: cells, links: links, imgs: imgs}};
            }});
            return {{headers: headers, rows: rows}};
        }}
        """
        try:
            return await self.page.evaluate(js)
        except Exception as e:
            self.logger.debug(f"extract_table_js failed: {e}")
            return None

    async def extract_list_items_js(self, selector: str) -> List[Dict[str, Any]]:
        js = f"""
        () => {{
            const nodes = Array.from(document.querySelectorAll("{selector}"));
            return nodes.map(n => {{
                const titleEl = n.querySelector('h1,h2,h3,h4') || n.querySelector('a') || n.querySelector('strong');
                const title = titleEl ? titleEl.innerText.trim() : '';
                const p = Array.from(n.querySelectorAll('p')).map(x=>x.innerText.trim()).join('\\n');
                const links = Array.from(n.querySelectorAll('a')).map(a=>a.href);
                const imgs = Array.from(n.querySelectorAll('img')).map(i=>i.src);
                return {{title: title, text: p, links: links, imgs: imgs}};
            }});
        }}
        """
        try:
            return await self.page.evaluate(js)
        except Exception as e:
            self.logger.debug(f"extract_list_items_js failed: {e}")
            return []

    async def detect_list_selector(self) -> Optional[str]:
        # Force a couple of scrolls for dynamic pages then detect
        try:
            await self.page.evaluate("window.scrollBy(0, 500)")
            await asyncio.sleep(0.6)
            await self.page.evaluate("window.scrollBy(0, 800)")
            await asyncio.sleep(0.6)
        except Exception:
            pass

        candidates = [".quote", ".card", ".list-item", ".item", "article", "ul li", ".media", ".search-result", ".result", ".post"]
        for sel in candidates:
            try:
                if await self.page.locator(sel).count() >= 2:
                    return sel
            except Exception:
                continue

        # auto-detect repeating classes
        try:
            classes = await self.page.evaluate("""
                () => {
                    const divs = Array.from(document.querySelectorAll('div'));
                    const counts = {};
                    divs.forEach(d => {
                        const cls = d.className || '';
                        if(!cls) return;
                        counts[cls] = (counts[cls] || 0) + 1;
                    });
                    return Object.entries(counts).filter(([cls,c])=>c>=3 && cls.trim()!=='').map(([cls,c]) => '.' + cls.trim().split(/\s+/).join('.'));
                }
            """)
            if classes and len(classes) > 0:
                return classes[0]
        except Exception:
            pass

        try:
            if (await self.page.locator("p").count()) > 6 and (await self.page.locator("h1,h2,h3,h4").count()) > 0:
                return "body"
        except Exception:
            pass

        return None

    # perform advanced scroll (window/wheel/pagedown)
    async def perform_advanced_scroll_step(self):
        try:
            await self.page.evaluate("window.scrollBy(0, document.body.scrollHeight)")
        except Exception:
            pass
        await asyncio.sleep(random.uniform(SCROLL_SLEEP_MIN, SCROLL_SLEEP_MAX))
        try:
            await self.page.mouse.wheel(0, 400)
        except Exception:
            pass
        try:
            await self.page.keyboard.press("PageDown")
        except Exception:
            pass
        await asyncio.sleep(0.15)

    async def scroll_until_stable(self, extract_func: Callable[[str], Any], selector_getter: Callable[[], Optional[str]], max_cycles:int = MAX_SCROLL_STAGNANT) -> List[Any]:
        combined = []
        last_count = 0
        stagnant = 0
        cycles = 0

        while True:
            if self.stop_event.is_set():
                self.logger.info("Stop requested during universal scroll.")
                break

            await self.perform_advanced_scroll_step()

            selector = selector_getter()
            data = None
            try:
                if selector:
                    data = await extract_func(selector)
            except Exception as e:
                self.logger.debug(f"extract during scroll failed: {e}")

            current_len = 0
            if isinstance(data, dict) and "rows" in data:
                current_len = len(data['rows'])
            elif isinstance(data, list):
                current_len = len(data)

            if data:
                if isinstance(data, dict) and "rows" in data:
                    for r in data['rows']:
                        key = "|".join(r.get('cells', [])) if r.get('cells') else "|".join(r.get('links', []))
                        if not key:
                            key = repr(r)[:200]
                        if key not in self.seen_keys:
                            self.seen_keys.add(key)
                            combined.append(r)
                else:
                    for it in data:
                        key = (it.get('title') or '') + '|' + (it.get('text') or '')
                        if key not in self.seen_keys:
                            self.seen_keys.add(key)
                            combined.append(it)

            if len(combined) == last_count:
                stagnant += 1
                if stagnant >= max_cycles:
                    break
            else:
                stagnant = 0

            last_count = len(combined)
            cycles += 1

            if cycles > 500 or len(combined) > 60000:
                break

        return combined

class FivePaisaExtractor:
    def __init__(self, page: Page, logger: logging.Logger, stop_event: threading.Event):
        self.page = page
        self.logger = logger
        self.stop_event = stop_event
        self.seen = set()

    async def find_candidate_containers(self) -> List[str]:
        candidates = [
            ".stock_table_wrapper",
            ".table-responsive",
            "div[class*='MuiBox-root']",
            "div[class*='table']",
            "div[class*='virtualized']",
            "div[class*='react-virtualized']",
            "div[class*='list']",
            "table"
        ]
        found = []
        for sel in candidates:
            try:
                if await self.page.locator(sel).count() > 0:
                    found.append(sel)
            except Exception:
                continue
        # add divs with overflow-y and scrollHeight > clientHeight
        try:
            extra = await self.page.evaluate("""
                () => {
                    const els = Array.from(document.querySelectorAll('div'));
                    const results = [];
                    for(const d of els){
                        try {
                            const s = window.getComputedStyle(d);
                            if(!s) continue;
                            if((s.overflowY === 'scroll' || s.overflowY === 'auto' || s.overflow === 'auto' || s.overflow === 'scroll') && d.scrollHeight > d.clientHeight){
                                let cls = d.className || '';
                                if(cls) results.push('.' + cls.trim().split(/\\s+/).join('.'));
                            }
                        } catch(e) {}
                    }
                    return results.slice(0,8);
                }
            """)
            if extra:
                for e in extra:
                    if e and e not in found:
                        found.append(e)
        except Exception:
            pass
        if "body" not in found:
            found.append("body")
        return found

    async def extract_from_table_selector(self) -> List[Dict[str, str]]:
        try:
            data = await self.page.evaluate("""
                () => {
                    const results = [];
                    const rows = document.querySelectorAll('table tbody tr');
                    for(const row of rows){
                        try {
                            const a = row.querySelector('a');
                            const img = row.querySelector('img');
                            if(a){
                                const name = a.textContent.trim();
                                let logo = '';
                                if(img){
                                    logo = img.getAttribute('src') || '';
                                    if(logo && !logo.startsWith('http')){
                                        logo = new URL(logo, window.location.origin).href;
                                    }
                                }
                                if(name && name.length > 2) results.push({company_name: name, logo_url: logo || 'N/A'});
                            }
                        } catch(e){}
                    }
                    return results;
                }
            """)
            return data or []
        except Exception as e:
            self.logger.debug(f"extract_from_table_selector error: {e}")
            return []

    async def scroll_container_once(self, container_sel: str):
        try:
            await self.page.evaluate(f"""
                () => {{
                    const c = document.querySelector("{container_sel}");
                    if(c){{
                        c.scrollTop = c.scrollTop + Math.max(400, Math.floor(c.clientHeight * 0.9));
                        c.dispatchEvent(new Event('scroll', {{bubbles:true}}));
                    }}
                }}
            """)
        except Exception:
            try:
                await self.page.evaluate("window.scrollBy(0, document.body.scrollHeight)")
            except Exception:
                pass
        await asyncio.sleep(random.uniform(0.6, 1.2))
        try:
            await self.page.mouse.wheel(0, 400)
        except Exception:
            pass

    async def container_scroll_until_stable(self, max_attempts: int = 2000, check_every:int = 3) -> List[Dict[str,str]]:
        candidates = await self.find_candidate_containers()
        self.logger.info(f"5Paisa candidate containers: {candidates}")
        best_result: List[Dict[str,str]] = []
        for container in candidates:
            if self.stop_event.is_set():
                break
            self.logger.info(f"Trying container: {container}")
            last_count = 0
            stagnant = 0
            collected: List[Dict[str,str]] = []
            scrolls = 0
            await asyncio.sleep(1.0)
            while scrolls < max_attempts:
                if self.stop_event.is_set():
                    break
                await self.scroll_container_once(container)
                scrolls += 1
                if scrolls % check_every == 0:
                    batch = await self.extract_from_table_selector()
                    before = len(collected)
                    for c in batch:
                        key = (c.get('company_name') or '').strip().lower()
                        if key and key not in self.seen:
                            self.seen.add(key)
                            collected.append(c)
                    new_added = len(collected) - before
                    self.logger.info(f"[{container}] scrolls={scrolls} total_rows={len(collected)} (+{new_added})")
                    if len(collected) == last_count:
                        stagnant += 1
                        if stagnant >= 30:
                            self.logger.info(f"[{container}] stagnant; breaking.")
                            break
                    else:
                        stagnant = 0
                    last_count = len(collected)
                await asyncio.sleep(random.uniform(0.15, 0.35))
                if len(collected) >= 8200:
                    break
            if len(collected) > len(best_result):
                best_result = collected
            if len(best_result) >= 8000:
                break
        return best_result


class AdaptiveScraperApp:
    def __init__(self, url: str, excel_file: str, logger: logging.Logger, stop_event: threading.Event,
                 download_images: bool = True, max_pages: int = MAX_PAGINATION_PAGES_DEFAULT,
                 headless: bool = False, img_concurrency: int = IMAGE_DOWNLOAD_CONCURRENCY):
        self.url = url
        self.excel_file = excel_file
        self.logger = logger
        self.stop_event = stop_event
        self.download_images = download_images
        self.max_pages = max_pages
        self.headless = headless
        self.img_concurrency = img_concurrency
        self.image_folder = "scraped_images"
        self.partial_collected: Dict[str, pd.DataFrame] = {}

        # playwright handles
        self._pw: Optional[Playwright] = None
        self.browser: Optional[Browser] = None
        self.page: Optional[Page] = None

    async def start_browser(self) -> bool:
        try:
            self._pw = await async_playwright().start()
            self.browser = await self._pw.chromium.launch(headless=self.headless)
            ctx = await self.browser.new_context(viewport={"width": 1366, "height": 768})
            self.page = await ctx.new_page()
            self.page.set_default_timeout(PLAYWRIGHT_TIMEOUT)
            return True
        except Exception as e:
            self.logger.error(f"Browser start failed: {e}")
            await self.stop_browser()
            return False

    async def stop_browser(self):
        try:
            if self.page:
                try:
                    await self.page.close()
                except Exception:
                    pass
            if self.browser:
                try:
                    await self.browser.close()
                except Exception:
                    pass
            if self._pw:
                try:
                    await self._pw.stop()
                except Exception:
                    pass
        except Exception:
            pass
        finally:
            self.page = None
            self.browser = None
            self._pw = None

    def save_partial_results(self):
        """Save whatever was collected so far in partial_collected to excel"""
        if not self.partial_collected:
            self.logger.warning("No partial data available to save.")
            return
        try:
            with pd.ExcelWriter(self.excel_file, engine=EXCEL_ENGINE) as writer:
                for sheet_name, df in self.partial_collected.items():
                    safe_name = str(sheet_name)[:31]
                    df.to_excel(writer, index=False, sheet_name=safe_name)
            self.logger.info(f"Partial results saved to {self.excel_file}")
        except Exception as e:
            self.logger.error(f"Error saving partial results: {e}")

    # orchestrator run
    async def run(self, progress_callback: Optional[Callable[[str, Any], None]] = None):
        try:
            started = await self.start_browser()
            if not started:
                return

            try:
                self.logger.info(f"Opening URL: {self.url}")
                await self.page.goto(self.url, wait_until="domcontentloaded")
            except PWError as e:
                self.logger.warning(f"Navigation domcontentloaded failed: {e}; trying networkidle")
                try:
                    await self.page.goto(self.url, wait_until="networkidle")
                except Exception as e2:
                    self.logger.error(f"Navigation failed: {e2}")
                    await self.stop_browser()
                    return

            # detect 5paisa by URL or DOM markers
            is_5paisa = False
            try:
                if "5paisa.com/stocks" in self.url or "5paisa.com" in self.url:
                    is_5paisa = True
                else:
                    marker = await self.page.evaluate("""() => {
                        return !!document.querySelector('.stock_table_wrapper') || !!document.querySelector('div[class*="MuiBox-root"]');
                    }""")
                    if marker:
                        is_5paisa = True
            except Exception:
                pass

            sheets: Dict[str, pd.DataFrame] = {}

            # 5Paisa special mode
            if is_5paisa:
                self.logger.info("5Paisa-like site detected. Running special extractor.")
                fp = FivePaisaExtractor(self.page, self.logger, self.stop_event)
                companies = await fp.container_scroll_until_stable(max_attempts=2000, check_every=3)
                if companies and len(companies) > 0:
                    df = pd.DataFrame(companies)
                    df.insert(0, "serial_no", range(1, len(df) + 1))
                    sheets["5Paisa_Stocks"] = df
                    # store partial copy
                    self.partial_collected["5Paisa_Stocks"] = df.copy()
                else:
                    self.logger.info("5Paisa extractor returned no data; will fallback to universal engine.")

            # Universal engine fallback / normal flow
            if not sheets:
                engine = UniversalEngine(self.page, self.logger, self.stop_event)

                table_count = await engine.detect_all_tables()
                table_sel = None
                if table_count > 0:
                    table_sel = await engine.detect_best_table_selector()

                list_sel = None
                if not table_sel:
                    list_sel = await engine.detect_list_selector()

                next_btn = await engine.detect_next_button()

                # multiple tables
                if table_count > 1:
                    self.logger.info(f"Detected {table_count} tables; extracting each.")
                    for idx in range(table_count):
                        if self.stop_event.is_set():
                            break
                        try:
                            data = await self.page.evaluate(f"""() => {{
                                const t = document.querySelectorAll('table')[{idx}];
                                if(!t) return null;
                                const heads = Array.from(t.querySelectorAll('thead th')).map(h=>h.innerText.trim());
                                let headers = heads.slice();
                                if(headers.length===0) {{
                                    const first = t.querySelector('tbody tr');
                                    if(first) headers = Array.from(first.querySelectorAll('td')).map((td,i)=>'Column_'+(i+1));
                                }}
                                const rows = Array.from(t.querySelectorAll('tbody tr')).map(tr => Array.from(tr.querySelectorAll('td')).map(td=>td.innerText.trim()));
                                return {{headers: headers, rows: rows}};
                            }}""")
                        except Exception as e:
                            self.logger.debug(f"Per-table eval failed: {e}")
                            data = None
                        if not data:
                            continue
                        headers = data.get("headers", [])
                        rows = data.get("rows", [])
                        if rows:
                            df = pd.DataFrame(rows, columns=headers if headers else None)
                            df.insert(0, "serial_no", range(1, len(df) + 1))
                            sheets[f"Table_{idx+1}"] = df
                            self.partial_collected[f"Table_{idx+1}"] = df.copy()

                # single table (pagination or infinite)
                elif table_sel:
                    self.logger.info("Single table detected.")
                    if next_btn:
                        self.logger.info("Table pagination detected; using Next-click pagination.")
                        rows = await engine.paginate_click_next("table", table_sel)
                        if rows:
                            try:
                                headers = await self.page.evaluate(f"""() => {{
                                    const t = document.querySelector("{table_sel}");
                                    if(!t) return [];
                                    const hs = Array.from(t.querySelectorAll('thead th')).map(h=>h.innerText.trim());
                                    if(hs.length===0) {{
                                        const first = t.querySelector('tbody tr');
                                        if(first) return Array.from(first.querySelectorAll('td')).map((td,i)=>'Column_'+(i+1));
                                    }}
                                    return hs;
                                }}""")
                            except Exception:
                                headers = []
                            records = []
                            for r in rows:
                                rec = {}
                                cells = r.get('cells', [])
                                for i, c in enumerate(cells):
                                    col = headers[i] if i < len(headers) else f"Column_{i+1}"
                                    rec[col] = c
                                rec['links'] = ", ".join(r.get('links', [])) if r.get('links') else ""
                                rec['images'] = ", ".join(r.get('imgs', [])) if r.get('imgs') else ""
                                records.append(rec)
                            if records:
                                df = pd.DataFrame(records)
                                df.insert(0, "serial_no", range(1, len(df) + 1))
                                sheets["Table"] = df
                                self.partial_collected["Table"] = df.copy()
                    else:
                        self.logger.info("Table without explicit pagination -> trying infinite/container-aware scroll.")
                        combined = await engine.scroll_until_stable(engine.extract_table_js, lambda: table_sel, max_cycles=MAX_SCROLL_STAGNANT)
                        if combined:
                            try:
                                headers = await self.page.evaluate(f"""() => {{
                                    const t = document.querySelector("{table_sel}");
                                    if(!t) return [];
                                    const hs = Array.from(t.querySelectorAll('thead th')).map(h=>h.innerText.trim());
                                    if(hs.length===0) {{
                                        const first = t.querySelector('tbody tr');
                                        if(first) return Array.from(first.querySelectorAll('td')).map((td,i)=>'Column_'+(i+1));
                                    }}
                                    return hs;
                                }}""")
                            except Exception:
                                headers = []
                            records = []
                            for r in combined:
                                rec = {}
                                cells = r.get('cells', [])
                                for i, c in enumerate(cells):
                                    col = headers[i] if i < len(headers) else f"Column_{i+1}"
                                    rec[col] = c
                                rec['links'] = ", ".join(r.get('links', [])) if r.get('links') else ""
                                rec['images'] = ", ".join(r.get('imgs', [])) if r.get('imgs') else ""
                                records.append(rec)
                            if records:
                                df = pd.DataFrame(records)
                                df.insert(0, "serial_no", range(1, len(df) + 1))
                                sheets["Table"] = df
                                self.partial_collected["Table"] = df.copy()

                # list-like pages
                elif list_sel:
                    self.logger.info(f"Detected list-like selector: {list_sel} -> extracting")
                    combined = await engine.scroll_until_stable(engine.extract_list_items_js, lambda: list_sel, max_cycles=MAX_SCROLL_STAGNANT)
                    records = []
                    for it in combined:
                        rec = {"title": it.get("title", ""), "text": it.get("text", ""), "links": ", ".join(it.get("links", [])) if it.get("links") else "", "images": ", ".join(it.get("imgs", [])) if it.get("imgs") else ""}
                        records.append(rec)
                    if records:
                        df = pd.DataFrame(records)
                        df.insert(0, "serial_no", range(1, len(df) + 1))
                        sheets["Items"] = df
                        self.partial_collected["Items"] = df.copy()

                # non-table pagination
                elif next_btn:
                    self.logger.info("Detected pagination on non-table page -> paginated list extraction")
                    rows = await engine.paginate_click_next("list", "body")
                    records = []
                    for it in rows:
                        rec = {"title": it.get("title", ""), "text": it.get("text", ""), "links": ", ".join(it.get("links", [])) if it.get("links") else "", "images": ", ".join(it.get("imgs", [])) if it.get("imgs") else ""}
                        records.append(rec)
                    if records:
                        df = pd.DataFrame(records)
                        df.insert(0, "serial_no", range(1, len(df) + 1))
                        sheets["Items"] = df
                        self.partial_collected["Items"] = df.copy()
                else:
                    self.logger.error("Could not detect table, list, or pagination on this page.")
                    await self.stop_browser()
                    return

            # STOP check before downloads & final save
            if self.stop_event.is_set():
                self.logger.info("Stop requested — saving partial data and exiting.")
                self.save_partial_results()
                await self.stop_browser()
                return

            # Images: download & validate (simple sequential)
            if self.download_images and sheets:
                for sheet_name, df in list(sheets.items()):
                    if 'images' in df.columns:
                        self.logger.info(f"Downloading/validating images for sheet '{sheet_name}' ...")
                        local_paths = []
                        statuses = []
                        for idx, imgcell in enumerate(df['images'].fillna(""), start=1):
                            first = imgcell.split(",")[0].strip() if imgcell else ""
                            if first:
                                path = await download_image_async(first, self.image_folder, f"{sheet_name}_r{idx}")
                                local_paths.append(path or "")
                                status = await head_check_image(first)
                                statuses.append(status)
                            else:
                                local_paths.append("")
                                statuses.append("Missing")
                        df['image_local'] = local_paths
                        df['image_status'] = statuses
                        sheets[sheet_name] = df
                        self.partial_collected[sheet_name] = df.copy()

            # Save final excel (multi-sheet)
            if sheets:
                try:
                    outdir = os.path.dirname(os.path.abspath(self.excel_file)) or "."
                    ensure_dir(outdir)
                    with pd.ExcelWriter(self.excel_file, engine=EXCEL_ENGINE) as writer:
                        for sheet_name, df in sheets.items():
                            safe_name = str(sheet_name)[:31]
                            df.to_excel(writer, index=False, sheet_name=safe_name)
                    self.logger.info(f"Saved workbook: {self.excel_file}")
                except Exception as e:
                    self.logger.error(f"Failed to save Excel: {e}")
            else:
                self.logger.error("No data collected; nothing to save.")

            await self.stop_browser()

        except Exception as e:
            self.logger.error(f"Adaptive run top-level error: {e}")
            # on error, attempt to save partial if any
            try:
                if self.partial_collected:
                    self.logger.info("Error occurred — saving partial results.")
                    self.save_partial_results()
            except Exception:
                pass
            try:
                await self.stop_browser()
            except Exception:
                pass

# ---------- GUI runner & threading ----------
def start_scraper_thread(
    url: str, excel: str, log_widget: ScrolledText, progressbar: ttk.Progressbar, stop_holder: dict,
    download_images_var: tk.BooleanVar, max_pages_entry: tk.Entry, headless_var: tk.BooleanVar, img_concurrency_entry: tk.Entry
):
    if not url.strip():
        messagebox.showerror("Input error", "Please enter a website URL.")
        return None
    if not excel.strip():
        messagebox.showerror("Input error", "Please enter an output filename.")
        return None
    if not excel.lower().endswith(".xlsx"):
        excel = excel.strip() + ".xlsx"
    try:
        max_pages = int(max_pages_entry.get())
    except Exception:
        max_pages = MAX_PAGINATION_PAGES_DEFAULT
    try:
        img_conc = int(img_concurrency_entry.get())
    except Exception:
        img_conc = IMAGE_DOWNLOAD_CONCURRENCY

    # configure logger
    logger = logging.getLogger("adaptive_gui")
    logger.setLevel(logging.INFO)
    logger.handlers = []
    handler = TextHandler(log_widget)
    handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
    logger.addHandler(handler)

    stop_event = threading.Event()
    stop_holder['evt'] = stop_event

    # create app instance here and store in holder for STOP access
    app_instance = AdaptiveScraperApp(
        url=url.strip(),
        excel_file=excel.strip(),
        logger=logger,
        stop_event=stop_event,
        download_images=download_images_var.get(),
        max_pages=max_pages,
        headless=headless_var.get(),
        img_concurrency=img_conc
    )
    stop_holder['app_instance'] = app_instance

    def progress_cb(action, payload=None):
        try:
            if action == "config" and isinstance(payload, dict):
                mode = payload.get("mode")
                if mode == "determinate":
                    progressbar.config(mode="determinate")
                    progressbar["value"] = 0
                    progressbar["maximum"] = payload.get("max", 100)
                else:
                    progressbar.config(mode="indeterminate")
                    progressbar.start(12)
            elif action == "update" and isinstance(payload, dict):
                val = payload.get("value", 0)
                total = payload.get("total")
                if total:
                    progressbar.config(mode="determinate")
                    progressbar["maximum"] = total
                    progressbar["value"] = val
            elif action == "done":
                try:
                    progressbar.stop()
                except Exception:
                    pass
                progressbar["value"] = 0
                progressbar.config(mode="determinate")
        except Exception:
            pass

    def worker():
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        loop.run_until_complete(app_instance.run(progress_callback=progress_cb))
        progress_cb("done")

    t = threading.Thread(target=worker, daemon=True)
    t.start()
    return stop_event

# ---------- Build GUI ----------
def build_and_run_gui():
    root = tk.Tk()
    root.title("Adaptive Universal Scraper - Creator NAMA Krityam")
    root.geometry("980x820")

    font_large = ("Segoe UI", 11)

    tk.Label(root, text="Target Website URL:", font=font_large).pack(anchor="w", padx=12, pady=(10,0))
    url_entry = tk.Entry(root, width=130)
    url_entry.pack(padx=12, pady=(4,6))
    tk.Label(root, text="Guideline: Paste full URL (e.g., https://www.example.com)", fg="gray").pack(anchor="w", padx=12)

    tk.Label(root, text="\nOutput Excel File Name:", font=font_large).pack(anchor="w", padx=12, pady=(8,0))
    excel_entry = tk.Entry(root, width=70)
    excel_entry.pack(anchor="w", padx=12, pady=(4,6))
    tk.Label(root, text="Use filename like mydata.xlsx (extension optional).", fg="gray").pack(anchor="w", padx=12)

    opts = tk.Frame(root)
    opts.pack(anchor="w", padx=12, pady=(8,6))
    download_images_var = tk.BooleanVar(value=True)
    tk.Checkbutton(opts, text="Download first image per row (if any)", variable=download_images_var).grid(row=0, column=0, sticky="w", padx=(0,8))
    tk.Label(opts, text="Max pagination clicks:", fg="black").grid(row=0, column=1, padx=(12,0))
    max_pages_entry = tk.Entry(opts, width=6)
    max_pages_entry.insert(0, str(MAX_PAGINATION_PAGES_DEFAULT))
    max_pages_entry.grid(row=0, column=2, padx=(6,0))
    headless_var = tk.BooleanVar(value=False)
    tk.Checkbutton(opts, text="Run headless (no browser window)", variable=headless_var).grid(row=0, column=3, sticky="w", padx=(18,8))
    tk.Label(opts, text="Image concurrency:", fg="black").grid(row=0, column=4, padx=(12,0))
    img_concurrency_entry = tk.Entry(opts, width=4)
    img_concurrency_entry.insert(0, str(IMAGE_DOWNLOAD_CONCURRENCY))
    img_concurrency_entry.grid(row=0, column=5, padx=(6,0))

    controls = tk.Frame(root)
    controls.pack(anchor="w", padx=12, pady=(6,6))
    progressbar = ttk.Progressbar(controls, orient="horizontal", length=560, mode="determinate")
    progressbar.grid(row=0, column=0, padx=(0,12))
    start_btn = tk.Button(controls, text="Start Scraping", bg="#1976D2", fg="white", font=("Segoe UI",10))
    start_btn.grid(row=0, column=1, padx=(6,6))
    stop_event_holder = {}
    stop_btn = tk.Button(controls, text="STOP", bg="#D32F2F", fg="white", state="disabled", font=("Segoe UI",10))
    stop_btn.grid(row=0, column=2, padx=(6,0))

    tk.Label(root, text="\nLogs:", font=font_large).pack(anchor="w", padx=12)
    log_widget = ScrolledText(root, height=30, width=120, state="disabled")
    log_widget.pack(padx=12, pady=(6,12))

    tk.Label(root, text="Tip: For local HTML use python -m http.server 8000 and open http://localhost:8000/yourfile.html", fg="gray").pack(anchor="w", padx=12, pady=(0,12))

    def on_start():
        start_btn.config(state="disabled")
        stop_btn.config(state="normal")
        stop_event = start_scraper_thread(
            url_entry.get(), excel_entry.get(), log_widget, progressbar, stop_event_holder,
            download_images_var, max_pages_entry, headless_var, img_concurrency_entry
        )
        stop_event_holder["evt"] = stop_event

    def on_stop():
        evt = stop_event_holder.get("evt")
        app_instance = stop_event_holder.get("app_instance")
        if evt and isinstance(evt, threading.Event):
            evt.set()
            # call partial save on the running app instance (if available)
            if app_instance and hasattr(app_instance, "save_partial_results"):
                try:
                    app_instance.save_partial_results()
                except Exception as e:
                    logging.getLogger("adaptive_gui").error(f"Partial save failed: {e}")
            stop_btn.config(state="disabled")
            start_btn.config(state="normal")
            logging.getLogger("adaptive_gui").info("Stop requested by user. Partial save attempted.")
        else:
            messagebox.showinfo("Stop", "No active scraping process detected.")

    start_btn.config(command=on_start)
    stop_btn.config(command=on_stop)

    root.mainloop()

if __name__ == "__main__":
    build_and_run_gui()

