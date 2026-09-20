"""
داشبورد Streamlit + اجراکننده‌ی پس‌زمینه‌ی ربات تلگرام (bot.py)
-----------------------------------------------------------------
- ربات فقط یک‌بار (به‌ازای هر پروسه‌ی سرور) با @st.cache_resource اجرا می‌شود،
  نه به‌ازای هر بازدید/ری‌اجرای صفحه؛ پس Conflict 409 تلگرام رخ نمی‌دهد.
- خروجی ربات در حافظه نگه داشته و در «Manage app» استریم‌لیت هم چاپ می‌شود.
- اگر پروسه‌ی ربات خودش بمیرد، با بازدید بعدی (حداکثر هر ۳۰ ثانیه) دوباره بالا می‌آید.

Secrets لازم (Settings > Secrets در share.streamlit.io):
    BOT_TOKEN          = "123456:ABC..."
    ADMIN_IDS          = "111111111,222222222"
    DASHBOARD_PASSWORD = "یک-رمز-قوی"      # اختیاری؛ برای دیدن لاگ و دکمه‌های کنترل
"""

from __future__ import annotations

import atexit
import base64
import collections
import hmac
import html
import os
import signal
import subprocess
import sys
import tempfile
import threading
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path

import streamlit as st

BASE_DIR = Path(__file__).resolve().parent
BOT_SCRIPT = BASE_DIR / "bot.py"
FONT_FILE = BASE_DIR / "Vazirmatn-Regular.ttf"
PID_FILE = Path(tempfile.gettempdir()) / "konkur_bot.pid"
TEHRAN_TZ = timezone(timedelta(hours=3, minutes=30))

st.set_page_config(page_title="داشبورد ربات کنکور", page_icon="🤖", layout="centered")


# ============================================================
#  ابزارهای کوچک
# ============================================================
_FA_DIGITS = str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹")


def fa(value) -> str:
    """تبدیل ارقام به فارسی برای نمایش."""
    return str(value).translate(_FA_DIGITS)


def fmt_duration(seconds: float) -> str:
    seconds = int(seconds)
    days, rem = divmod(seconds, 86400)
    hours, rem = divmod(rem, 3600)
    minutes, secs = divmod(rem, 60)
    parts = []
    if days:
        parts.append(f"{days} روز")
    if hours:
        parts.append(f"{hours} ساعت")
    if minutes:
        parts.append(f"{minutes} دقیقه")
    if not parts:
        parts.append(f"{secs} ثانیه")
    return fa(" و ".join(parts))


def secret(name: str, default: str = "") -> str:
    """خواندن امن از st.secrets (بدون خطا وقتی فایل secrets وجود ندارد)."""
    try:
        if name in st.secrets:
            value = st.secrets[name]
            if isinstance(value, (list, tuple)):
                return ",".join(str(v) for v in value)
            return str(value)
    except Exception:  # noqa: BLE001 — لوکال و بدون secrets.toml
        pass
    return os.environ.get(name, default)


def build_env() -> dict:
    """محیط اجرای bot.py: متغیرهای فعلی + مقادیر Secrets."""
    env = os.environ.copy()
    for key in ("BOT_TOKEN", "ADMIN_IDS"):
        value = secret(key)
        if value:
            env[key] = value
    env["PYTHONUNBUFFERED"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    return env


# ============================================================
#  مدیریت پروسه‌ی ربات
# ============================================================
def _pid_alive_and_is_bot(pid: int) -> bool:
    """فقط اگر PID هنوز زنده باشد و واقعاً bot.py باشد True برمی‌گرداند (نیازمند /proc لینوکس)."""
    try:
        raw = Path(f"/proc/{pid}/cmdline").read_bytes()
    except OSError:
        return False
    return "bot.py" in raw.replace(b"\0", b" ").decode(errors="ignore")


def _kill_stale_bot() -> None:
    """
    اگر از اجرای قبلی (مثلاً بعد از hot-reload یا ری‌استارت اپ) یک bot.py یتیم مانده باشد،
    آن را می‌بندد؛ وگرنه دو پروسه همزمان polling می‌کنند و خطای 409 می‌گیرند.
    """
    try:
        pid = int(PID_FILE.read_text().strip())
    except (OSError, ValueError):
        return
    if _pid_alive_and_is_bot(pid):
        try:
            os.kill(pid, signal.SIGTERM)
            for _ in range(50):
                time.sleep(0.1)
                if not _pid_alive_and_is_bot(pid):
                    break
            else:
                os.kill(pid, getattr(signal, "SIGKILL", signal.SIGTERM))
        except ProcessLookupError:
            pass
    PID_FILE.unlink(missing_ok=True)


class BotManager:
    """نگهدارنده‌ی تک‌نمونه‌ی پروسه‌ی bot.py."""

    COOLDOWN_SECONDS = 30      # حداقل فاصله‌ی بین دو استارت خودکار
    LOG_LINES = 300

    def __init__(self) -> None:
        self._lock = threading.RLock()
        self._proc: subprocess.Popen | None = None
        self.logs: collections.deque[str] = collections.deque(maxlen=self.LOG_LINES)
        self.started_at: float | None = None
        self.start_count = 0
        self.manually_stopped = False
        self.last_attempt = 0.0
        self.last_error = ""

    # ---------- وضعیت ----------
    def is_running(self) -> bool:
        return self._proc is not None and self._proc.poll() is None

    @property
    def pid(self) -> int | None:
        return self._proc.pid if self._proc else None

    @property
    def exit_code(self) -> int | None:
        return self._proc.poll() if self._proc else None

    def uptime(self) -> float | None:
        return time.time() - self.started_at if self.is_running() and self.started_at else None

    # ---------- کنترل ----------
    def start(self, env: dict) -> None:
        with self._lock:
            if self.is_running():
                return
            self.last_attempt = time.time()
            self.manually_stopped = False

            if not env.get("BOT_TOKEN", "").strip():
                self.last_error = "BOT_TOKEN در Secrets تنظیم نشده است."
                return
            if not BOT_SCRIPT.exists():
                self.last_error = "فایل bot.py کنار app.py پیدا نشد."
                return

            self.last_error = ""
            _kill_stale_bot()
            try:
                self._proc = subprocess.Popen(
                    [sys.executable, "-u", str(BOT_SCRIPT)],
                    cwd=str(BASE_DIR),
                    env=env,
                    stdin=subprocess.DEVNULL,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    bufsize=1,
                )
            except OSError as exc:
                self.last_error = f"اجرای bot.py ممکن نشد: {exc}"
                return

            PID_FILE.write_text(str(self._proc.pid))
            self.started_at = time.time()
            self.start_count += 1
            self.logs.append(f"--- شروع پروسه‌ی ربات (PID {self._proc.pid}) ---")
            threading.Thread(target=self._pump, args=(self._proc,), daemon=True).start()

    def _pump(self, proc: subprocess.Popen) -> None:
        """خط‌به‌خطِ خروجی ربات را در حافظه نگه می‌دارد و در کنسول Manage app چاپ می‌کند."""
        assert proc.stdout is not None
        for line in proc.stdout:
            line = line.rstrip("\n")
            self.logs.append(line)
            print(f"[bot] {line}", flush=True)
        proc.wait()
        self.logs.append(f"--- پروسه‌ی ربات پایان یافت (کد خروج {proc.returncode}) ---")

    def stop(self, manual: bool = True) -> None:
        with self._lock:
            if manual:
                self.manually_stopped = True
            proc = self._proc
            if proc and proc.poll() is None:
                proc.terminate()
                try:
                    proc.wait(timeout=8)
                except subprocess.TimeoutExpired:
                    proc.kill()
                    proc.wait(timeout=5)
            PID_FILE.unlink(missing_ok=True)

    def restart(self, env: dict) -> None:
        self.stop(manual=False)
        self.start(env)

    def ensure_running(self, env: dict) -> None:
        """Self-heal: اگر ربات مرده و دستی متوقف نشده، با فاصله‌ی زمانی دوباره بالا می‌آید."""
        if self.is_running() or self.manually_stopped:
            return
        if time.time() - self.last_attempt < self.COOLDOWN_SECONDS:
            return
        self.start(env)

    def log_tail(self, token: str = "", n: int = 80) -> str:
        text = "\n".join(list(self.logs)[-n:])
        return text.replace(token, "***") if token else text


@st.cache_resource(show_spinner=False)
def get_manager() -> BotManager:
    """یک نمونه‌ی مشترک برای همه‌ی سشن‌ها؛ ساخته‌شدنش ربات را فقط یک‌بار استارت می‌کند."""
    manager = BotManager()
    atexit.register(manager.stop, False)   # با بسته‌شدن سرور، فرزند یتیم نماند
    manager.start(build_env())
    return manager


# ============================================================
#  ظاهر (RTL + فونت وزیرمتن)
# ============================================================
@st.cache_resource(show_spinner=False)
def _font_face_css() -> str:
    """فونت را به‌صورت base64 در خود صفحه می‌گذارد تا به Google Fonts وابسته نباشیم."""
    if not FONT_FILE.exists():
        return ""
    b64 = base64.b64encode(FONT_FILE.read_bytes()).decode()
    return (
        "@font-face{font-family:'KVazir';font-weight:400;font-style:normal;font-display:swap;"
        f"src:url(data:font/ttf;base64,{b64}) format('truetype');}}"
    )


STYLE = """
:root{--kb-font:'KVazir',Vazirmatn,Tahoma,'Segoe UI',sans-serif;--kb-accent:#1f4e79}
.stApp{direction:rtl}
.stApp :is(h1,h2,h3,h4,p,label,li,button,input,textarea,summary,[class*="kb-"]){font-family:var(--kb-font)}
.stApp :is(pre,code,[data-testid="stCode"]){direction:ltr;text-align:left}
.stApp input[type=password]{direction:ltr;text-align:left}
.block-container{max-width:760px;padding-top:2.2rem}
footer{visibility:hidden}

.kb-hero{padding:1.4rem 1.5rem;border-radius:18px;color:#fff;margin-bottom:1.1rem;
  background:linear-gradient(135deg,#1f4e79 0%,#2b7bb9 100%)}
.kb-hero-title{font-size:1.45rem;font-weight:700;line-height:1.7}
.kb-hero-sub{opacity:.85;font-size:.95rem}

.kb-status{display:flex;align-items:center;gap:.6rem;padding:.9rem 1.1rem;border-radius:14px;
  font-weight:700;border:1px solid;margin-bottom:.9rem}
.kb-dot{width:.7rem;height:.7rem;border-radius:50%;background:currentColor;flex:none}
.kb-ok{color:#15803d;background:rgba(22,163,74,.10);border-color:rgba(22,163,74,.35)}
.kb-ok .kb-dot{animation:kbpulse 1.8s infinite}
.kb-bad{color:#b91c1c;background:rgba(220,38,38,.09);border-color:rgba(220,38,38,.35)}
.kb-warn{color:#b45309;background:rgba(217,119,6,.10);border-color:rgba(217,119,6,.35)}
@keyframes kbpulse{0%{opacity:1}50%{opacity:.35}100%{opacity:1}}

.kb-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:.7rem;margin-bottom:.9rem}
.kb-card{padding:.85rem 1rem;border-radius:14px;border:1px solid rgba(128,128,128,.28);
  background:rgba(128,128,128,.07)}
.kb-label{font-size:.8rem;opacity:.7;margin-bottom:.15rem}
.kb-value{font-size:1.05rem;font-weight:700}

.kb-checks{padding:.8rem 1rem;border-radius:14px;border:1px dashed rgba(128,128,128,.4);
  font-size:.92rem;line-height:2;margin-bottom:.9rem}
.kb-note{font-size:.82rem;opacity:.65;line-height:1.9;margin-top:.6rem}
"""


def inject_style() -> None:
    st.markdown(f"<style>{_font_face_css()}{STYLE}</style>", unsafe_allow_html=True)


# ============================================================
#  صفحه
# ============================================================
def is_authed() -> bool:
    return bool(st.session_state.get("authed"))


def _cb_restart() -> None:
    get_manager().restart(build_env())
    st.toast("ربات دوباره راه‌اندازی شد.")


def _cb_stop() -> None:
    get_manager().stop(manual=True)
    st.toast("ربات متوقف شد.")


def _cb_start() -> None:
    get_manager().start(build_env())


def render_login() -> None:
    password = secret("DASHBOARD_PASSWORD")
    if not password:
        st.caption("🔒 برای دیدن لاگ و کنترل ربات، در Secrets مقدار DASHBOARD_PASSWORD را تنظیم کنید.")
        return
    if is_authed():
        if st.button("🚪 خروج از حالت مدیر"):
            st.session_state["authed"] = False
            st.rerun()
        return
    with st.expander("🔐 ورود مدیر (لاگ و کنترل ربات)"):
        with st.form("login_form"):
            entered = st.text_input("رمز داشبورد", type="password")
            if st.form_submit_button("ورود"):
                if hmac.compare_digest(entered.encode(), password.encode()):
                    st.session_state["authed"] = True
                    st.rerun()
                else:
                    time.sleep(1)  # کندکردن حدس‌زدن رمز
                    st.error("رمز نادرست است.")


def _status_banner(mgr: BotManager) -> str:
    if mgr.is_running():
        return '<div class="kb-status kb-ok"><span class="kb-dot"></span>ربات فعال است و در حال دریافت پیام‌هاست</div>'
    if mgr.last_error:
        return f'<div class="kb-status kb-bad"><span class="kb-dot"></span>{html.escape(mgr.last_error)}</div>'
    if mgr.manually_stopped:
        return '<div class="kb-status kb-warn"><span class="kb-dot"></span>ربات به‌صورت دستی متوقف شده است</div>'
    if mgr.exit_code is not None:
        return (
            '<div class="kb-status kb-bad"><span class="kb-dot"></span>'
            f"ربات متوقف شد (کد خروج {fa(mgr.exit_code)}) — راه‌اندازی خودکار در بازدید بعدی</div>"
        )
    return '<div class="kb-status kb-warn"><span class="kb-dot"></span>در حال راه‌اندازی…</div>'


@st.fragment(run_every=10)
def status_panel() -> None:
    mgr = get_manager()
    env = build_env()
    mgr.ensure_running(env)

    st.markdown(_status_banner(mgr), unsafe_allow_html=True)

    uptime = mgr.uptime()
    now = datetime.now(TEHRAN_TZ).strftime("%H:%M:%S")
    cards = [
        ("زمان فعالیت", fmt_duration(uptime) if uptime is not None else "—"),
        ("شناسه پردازش (PID)", fa(mgr.pid) if mgr.is_running() else "—"),
        ("دفعات راه‌اندازی", fa(mgr.start_count)),
        ("آخرین بررسی (وقت تهران)", fa(now)),
    ]
    grid = "".join(
        f'<div class="kb-card"><div class="kb-label">{label}</div><div class="kb-value">{value}</div></div>'
        for label, value in cards
    )
    st.markdown(f'<div class="kb-grid">{grid}</div>', unsafe_allow_html=True)

    token_ok = bool(env.get("BOT_TOKEN", "").strip())
    admin_ok = bool(env.get("ADMIN_IDS", "").strip())
    st.markdown(
        '<div class="kb-checks">'
        f"{'✅' if token_ok else '❌'} BOT_TOKEN {'تنظیم شده است' if token_ok else 'تنظیم نشده است'}<br>"
        f"{'✅' if admin_ok else '⚠️'} ADMIN_IDS {'تنظیم شده است' if admin_ok else 'خالی است؛ پنل /admin در دسترس نخواهد بود'}"
        "</div>",
        unsafe_allow_html=True,
    )

    if not is_authed():
        return

    # ---------- بخش مدیر ----------
    st.markdown("#### 🛠 کنترل ربات")
    cols = st.columns(2)
    if mgr.is_running():
        cols[0].button("🔄 راه‌اندازی مجدد", on_click=_cb_restart)
        cols[1].button("⏹ توقف", on_click=_cb_stop)
    else:
        cols[0].button("▶️ شروع", on_click=_cb_start)

    st.markdown("#### 📜 لاگ ربات")
    st.code(mgr.log_tail(env.get("BOT_TOKEN", "")) or "هنوز خروجی‌ای ثبت نشده است.", language="text", wrap_lines=True)


def main() -> None:
    inject_style()
    st.markdown(
        '<div class="kb-hero"><div class="kb-hero-title">🤖 ربات ثبت‌نام کلاس‌های کنکور ریاضی</div>'
        '<div class="kb-hero-sub">داشبورد وضعیت سرور</div></div>',
        unsafe_allow_html=True,
    )
    get_manager()          # اطمینان از استارت‌شدن ربات در اولین بازدید
    render_login()
    status_panel()
    st.markdown(
        '<div class="kb-note">ℹ️ ربات به‌صورت پس‌زمینه اجرا می‌شود و این صفحه فقط وضعیت را نشان می‌دهد. '
        "Streamlit Community Cloud اپ‌های بدون بازدید را بعد از ۱۲ ساعت به خواب می‌برد؛ "
        "در آن حالت ربات هم خاموش می‌شود تا کسی صفحه را باز کند.</div>",
        unsafe_allow_html=True,
    )


main()
