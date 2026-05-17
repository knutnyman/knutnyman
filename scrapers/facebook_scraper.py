"""Scrape Dexcom-related Facebook group posts via Selenium.

NOTE: Facebook scraping is fragile (Facebook changes markup frequently) and
against Facebook's Terms of Service.  Use at your own risk and only for
personal/research purposes.  Configure FACEBOOK_GROUP_URLS with comma-separated
group URLs.
"""

import logging
import os
import time

from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

log = logging.getLogger(__name__)

# CSS selectors tried in order when extracting post text — Facebook changes these often
_POST_SELECTORS = [
    "div[data-ad-preview='message']",
    "div[dir='auto'][class*='x1iorvi4']",
    "div[data-testid='post_message']",
    "div[class*='userContent']",
    "div[class*='story_body_container'] div[dir='auto']",
    "div[role='article'] div[dir='auto']",
]


def _build_driver() -> webdriver.Chrome:
    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    opts.add_argument(
        "user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    )
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=opts)
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    })
    return driver


def _dismiss_cookies(driver: webdriver.Chrome) -> None:
    selectors = [
        "button[data-cookiebanner='accept_only_essential_button']",
        "button[title='Allow all cookies']",
        "[data-testid='cookie-policy-manage-dialog'] button:last-child",
        "button[data-testid='cookie-policy-dialog-accept-button']",
    ]
    for sel in selectors:
        try:
            btn = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.CSS_SELECTOR, sel)))
            btn.click()
            time.sleep(0.5)
            return
        except TimeoutException:
            pass


def _login(driver: webdriver.Chrome, email: str, password: str) -> bool:
    driver.get("https://www.facebook.com/login")
    time.sleep(2)
    _dismiss_cookies(driver)
    try:
        wait = WebDriverWait(driver, 10)
        wait.until(EC.presence_of_element_located((By.ID, "email"))).send_keys(email)
        driver.find_element(By.ID, "pass").send_keys(password)
        driver.find_element(By.NAME, "login").click()
        time.sleep(4)
        if "login" in driver.current_url or "checkpoint" in driver.current_url:
            log.error("Facebook login failed or hit checkpoint — check credentials / 2FA")
            return False
        log.info("Facebook login successful")
        return True
    except Exception as exc:
        log.error("Facebook login error: %s", exc)
        return False


def _scrape_group(driver: webdriver.Chrome, group_url: str, scroll_passes: int = 6) -> list[dict]:
    driver.get(group_url)
    time.sleep(3)
    _dismiss_cookies(driver)

    for _ in range(scroll_passes):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)

    group_label = group_url.rstrip("/").split("/")[-1] or group_url

    posts: list[dict] = []
    for selector in _POST_SELECTORS:
        elements = driver.find_elements(By.CSS_SELECTOR, selector)
        for el in elements:
            try:
                text = el.text.strip()
            except Exception:
                continue
            if len(text) < 25:
                continue
            post_id = f"fb_{abs(hash(text))}"
            posts.append({
                "id": post_id,
                "source": f"facebook/{group_label}",
                "title": text[:100],
                "body": text[:500],
                "text": text[:1200],
                "url": group_url,
                "score": 0,
                "num_comments": 0,
                "created_utc": None,
            })
        if posts:
            break

    # Deduplicate by id
    seen: set[str] = set()
    unique = [p for p in posts if not (p["id"] in seen or seen.add(p["id"]))]  # type: ignore[func-returns-value]
    log.info("facebook/%s: %d posts extracted", group_label, len(unique))
    return unique


def fetch_posts() -> list[dict]:
    email = os.getenv("FACEBOOK_EMAIL")
    password = os.getenv("FACEBOOK_PASSWORD")
    raw_urls = os.getenv("FACEBOOK_GROUP_URLS", "")

    if not all([email, password, raw_urls]):
        log.warning("FACEBOOK_EMAIL/FACEBOOK_PASSWORD/FACEBOOK_GROUP_URLS not set — skipping Facebook")
        return []

    group_urls = [u.strip() for u in raw_urls.split(",") if u.strip()]
    driver = _build_driver()
    all_posts: list[dict] = []

    try:
        if not _login(driver, email, password):
            return []
        for url in group_urls:
            try:
                all_posts.extend(_scrape_group(driver, url))
            except Exception as exc:
                log.warning("Failed to scrape %s: %s", url, exc)
    finally:
        driver.quit()

    log.info("Facebook total: %d posts", len(all_posts))
    return all_posts
