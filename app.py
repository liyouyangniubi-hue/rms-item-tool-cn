# -*- coding: utf-8 -*-
"""RMS商品一括編集ツール（ローカルWebアプリ）"""
import copy
import csv
import io
import json
import os
import re
import sqlite3
import sys
import threading
import time
import traceback
import webbrowser

from flask import Flask, jsonify, request, send_file, render_template

from rms_client import RmsClient, MockRmsClient, RmsError

if getattr(sys, "frozen", False):
    # PyInstaller製EXEとして実行中：テンプレートは展開先、データはEXEの隣に置く
    RESOURCE_DIR = sys._MEIPASS
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    RESOURCE_DIR = BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "items.db")
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")

APP_VERSION = "1.0.7"

app = Flask(__name__, template_folder=os.path.join(RESOURCE_DIR, "templates"))
app.config["JSON_AS_ASCII"] = False


@app.after_request
def no_cache(resp):
    # 古い画面がブラウザにキャッシュされて新機能が見えなくなるのを防ぐ
    resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate"
    resp.headers["Pragma"] = "no-cache"
    return resp

# ----------------- 設定 -----------------
DEFAULT_CONFIG = {"serviceSecret": "", "licenseKey": "", "demoMode": True, "waitMs": 700}


def load_config():
    cfg = dict(DEFAULT_CONFIG)
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, encoding="utf-8") as f:
                cfg.update(json.load(f))
        except Exception:
            pass
    return cfg


def save_config(cfg):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


_mock = None


def get_client():
    global _mock
    cfg = load_config()
    if cfg.get("demoMode"):
        if _mock is None:
            _mock = MockRmsClient()
        return _mock
    if not cfg.get("serviceSecret") or not cfg.get("licenseKey"):
        raise RmsError(0, "尚未设置API密钥。请在“设置”中填写 serviceSecret / licenseKey。")
    return RmsClient(cfg["serviceSecret"], cfg["licenseKey"], wait_ms=int(cfg.get("waitMs", 700)))


# ----------------- DB -----------------
def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with db() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS items(
                manage_number TEXT PRIMARY KEY,
                title TEXT, tagline TEXT, item_number TEXT,
                hide_item INTEGER DEFAULT 0,
                min_price REAL, max_price REAL,
                raw TEXT, edited TEXT,
                is_new INTEGER DEFAULT 0,
                downloaded_at TEXT, updated_local_at TEXT
            );
            CREATE TABLE IF NOT EXISTS logs(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ts TEXT, manage_number TEXT, action TEXT,
                status TEXT, message TEXT
            );
            CREATE TABLE IF NOT EXISTS jobs(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT, run_at TEXT, targets TEXT, snapshot TEXT,
                restore_price INTEGER DEFAULT 1,
                period_mode TEXT DEFAULT 'none',
                status TEXT DEFAULT 'pending',
                result TEXT, created_at TEXT, executed_at TEXT
            );
            CREATE TABLE IF NOT EXISTS sale_history(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT, targets TEXT, snapshot TEXT, created_at TEXT
            );
            """
        )
        # 既存DBへの列追加（セール対象フラグ）
        try:
            conn.execute("ALTER TABLE items ADD COLUMN sale_flag INTEGER DEFAULT 0")
        except sqlite3.OperationalError:
            pass


def now():
    return time.strftime("%Y-%m-%d %H:%M:%S")


def add_log(manage_number, action, status, message=""):
    with db() as conn:
        conn.execute(
            "INSERT INTO logs(ts,manage_number,action,status,message) VALUES(?,?,?,?,?)",
            (now(), manage_number, action, status, str(message)[:2000]),
        )


def to_price(value):
    """価格を数値として読み取る。
    RMS Item API 2.0 は standardPrice / referencePrice.value を
    文字列（例 "38500"）で返すため、数値決め打ちで判定すると全滅する。
    読み取れない場合は None。"""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).strip().replace(",", ""))
    except (TypeError, ValueError):
        return None


def price_like(original, value):
    """RMSへ送り返す価格を、元データと同じ型で作る。
    受け取った形（文字列）と違う型で返すと弾かれる可能性があるため型を揃える。
    元データが無い場合（CSVでの新規登録など）はAPIの返却形式に合わせて文字列。"""
    iv = max(int(round(value)), 1)
    return iv if isinstance(original, (int, float)) and not isinstance(original, bool) else str(iv)


def _csv_price(value):
    """CSVには数値として書き出す（Excelで計算できるように）。空なら空欄。"""
    p = to_price(value)
    return "" if p is None else int(round(p))


def item_summary(item):
    prices = []
    for v in (item.get("variants") or {}).values():
        if isinstance(v, dict):
            p = to_price(v.get("standardPrice"))
            if p is not None:
                prices.append(p)
    return {
        "title": item.get("title", ""),
        "tagline": item.get("tagline", ""),
        "item_number": item.get("itemNumber", ""),
        "hide_item": 1 if item.get("hideItem") else 0,
        "min_price": min(prices) if prices else None,
        "max_price": max(prices) if prices else None,
    }


def save_downloaded(item):
    mn = item.get("manageNumber")
    if not mn:
        return
    s = item_summary(item)
    with db() as conn:
        conn.execute(
            """INSERT INTO items(manage_number,title,tagline,item_number,hide_item,min_price,max_price,raw,edited,is_new,downloaded_at)
               VALUES(?,?,?,?,?,?,?,?,NULL,0,?)
               ON CONFLICT(manage_number) DO UPDATE SET
                 title=excluded.title, tagline=excluded.tagline, item_number=excluded.item_number,
                 hide_item=excluded.hide_item, min_price=excluded.min_price, max_price=excluded.max_price,
                 raw=excluded.raw, edited=NULL, is_new=0, downloaded_at=excluded.downloaded_at""",
            (mn, s["title"], s["tagline"], s["item_number"], s["hide_item"],
             s["min_price"], s["max_price"], json.dumps(item, ensure_ascii=False), now()),
        )


def save_local_edit(mn, item, is_new=None):
    s = item_summary(item)
    with db() as conn:
        if is_new:
            conn.execute(
                """INSERT INTO items(manage_number,title,tagline,item_number,hide_item,min_price,max_price,raw,edited,is_new,updated_local_at)
                   VALUES(?,?,?,?,?,?,?,NULL,?,1,?)
                   ON CONFLICT(manage_number) DO UPDATE SET
                     title=excluded.title, tagline=excluded.tagline, item_number=excluded.item_number,
                     hide_item=excluded.hide_item, min_price=excluded.min_price, max_price=excluded.max_price,
                     edited=excluded.edited, is_new=1, updated_local_at=excluded.updated_local_at""",
                (mn, s["title"], s["tagline"], s["item_number"], s["hide_item"],
                 s["min_price"], s["max_price"], json.dumps(item, ensure_ascii=False), now()),
            )
        else:
            conn.execute(
                """UPDATE items SET title=?, tagline=?, item_number=?, hide_item=?, min_price=?, max_price=?,
                   edited=?, updated_local_at=? WHERE manage_number=?""",
                (s["title"], s["tagline"], s["item_number"], s["hide_item"],
                 s["min_price"], s["max_price"], json.dumps(item, ensure_ascii=False), now(), mn),
            )


def get_effective(row):
    """編集済があれば編集済、なければ元データ"""
    src = row["edited"] or row["raw"]
    return json.loads(src) if src else None


# ----------------- バックグラウンドジョブ -----------------
JOB = {"running": False, "kind": "", "done": 0, "total": None, "message": "", "errors": [], "finished": True}
_job_lock = threading.Lock()


def start_job(kind, target, *args):
    with _job_lock:
        if JOB["running"]:
            return False
        JOB.update({"running": True, "kind": kind, "done": 0, "total": None,
                    "message": "已开始", "errors": [], "finished": False})
    t = threading.Thread(target=target, args=args, daemon=True)
    t.start()
    return True


def finish_job(message):
    JOB.update({"running": False, "finished": True, "message": message})


def job_download():
    try:
        client = get_client()
        count = 0

        def progress(fetched, total):
            JOB["done"] = fetched
            JOB["total"] = total
            JOB["message"] = f"{fetched}件已获取" + (f" / 共{total}件" if total else "")

        for item in client.iter_all_items(progress_cb=progress):
            save_downloaded(item)
            count += 1
        add_log("", "下载商品列表", "OK", f"{count}件")
        finish_job(f"完成：{count}件商品已下载")
    except Exception as e:
        add_log("", "下载商品列表", "ERROR", str(e))
        JOB["errors"].append(str(e))
        finish_job(f"错误：{e}")


PATCH_KEYS = [
    "itemNumber", "title", "tagline", "productDescription", "salesDescription",
    "genreId", "tags", "hideItem", "itemType", "images", "whiteBgImage",
    "payment", "features", "pointCampaign", "itemDisplaySequence",
    "purchasablePeriod", "releaseDate", "unlimitedInventoryFlag", "variants",
    "variantSelectors",
]


def diff_item(raw, edited):
    """トップレベル項目単位の差分を返す（PATCH用）"""
    patch = {}
    for k in PATCH_KEYS:
        if edited.get(k) != raw.get(k):
            patch[k] = edited.get(k)
    return patch


# RMSのGETレスポンスに含まれるが、登録・更新(PUT/PATCH)では送信できない読み取り専用フィールド
READONLY_FIELDS = {"created", "updated", "itemUrl"}
_UNRECOGNIZED_RE = re.compile(r'Unrecognized field "([^"]+)"')


def strip_keys(obj, keys):
    """辞書/リストを再帰的に走査し、指定キーを除去した新しいオブジェクトを返す"""
    if isinstance(obj, dict):
        return {k: strip_keys(v, keys) for k, v in obj.items() if k not in keys}
    if isinstance(obj, list):
        return [strip_keys(v, keys) for v in obj]
    return obj


def send_with_retry(send_fn, body, max_retry=5):
    """未知の読み取り専用フィールドをエラーメッセージから特定して除去しつつ再送する"""
    body = strip_keys(body, READONLY_FIELDS)
    for _ in range(max_retry):
        try:
            return send_fn(body)
        except RmsError as e:
            m = _UNRECOGNIZED_RE.search(e.message or "")
            if e.status == 400 and m:
                body = strip_keys(body, {m.group(1)})
                continue
            raise
    return send_fn(body)


def job_apply(manage_numbers):
    try:
        client = get_client()
        with db() as conn:
            if manage_numbers:
                q = ",".join("?" * len(manage_numbers))
                rows = conn.execute(
                    f"SELECT * FROM items WHERE manage_number IN ({q}) AND (edited IS NOT NULL OR is_new=1)",
                    manage_numbers).fetchall()
            else:
                rows = conn.execute(
                    "SELECT * FROM items WHERE edited IS NOT NULL OR is_new=1").fetchall()
        JOB["total"] = len(rows)
        ok = ng = 0
        for row in rows:
            mn = row["manage_number"]
            try:
                edited = json.loads(row["edited"])
                quantity = edited.pop("_initialQuantity", None)
                edited.pop("manageNumber", None)
                if row["is_new"]:
                    send_with_retry(lambda b: client.upsert_item(mn, b), edited)
                    if quantity is not None:
                        invs = [{"manageNumber": mn, "variantId": vid,
                                 "mode": "ABSOLUTE", "quantity": int(quantity)}
                                for vid in (edited.get("variants") or {})]
                        if invs:
                            client.bulk_upsert_inventory(invs)
                    add_log(mn, "新增商品", "OK")
                else:
                    raw = json.loads(row["raw"])
                    patch = diff_item(raw, edited)
                    if patch:
                        send_with_retry(lambda b: client.patch_item(mn, b), patch)
                    add_log(mn, "更新", "OK", "变更字段： " + ", ".join(patch.keys()) if patch else "无变化")
                edited["manageNumber"] = mn
                with db() as conn:
                    s = item_summary(edited)
                    conn.execute(
                        """UPDATE items SET raw=?, edited=NULL, is_new=0, title=?, tagline=?,
                           item_number=?, hide_item=?, min_price=?, max_price=? WHERE manage_number=?""",
                        (json.dumps(edited, ensure_ascii=False), s["title"], s["tagline"],
                         s["item_number"], s["hide_item"], s["min_price"], s["max_price"], mn),
                    )
                ok += 1
            except Exception as e:
                ng += 1
                JOB["errors"].append(f"{mn}: {e}")
                add_log(mn, "新增商品" if row["is_new"] else "更新", "ERROR", str(e))
            JOB["done"] += 1
            JOB["message"] = f"{JOB['done']}/{len(rows)}件已处理（成功{ok} 失败{ng}）"
        finish_job(f"完成：成功{ok}件 / 失败{ng}件")
    except Exception as e:
        JOB["errors"].append(str(e))
        finish_job(f"错误：{e}")


# ----------------- 復元予約（スケジューラ） -----------------
OVERDUE_GRACE_MIN = 10  # これを超えて時刻を過ぎていた予約は自動実行せず「要確認」にする


def make_snapshot(targets):
    """対象商品の現在値（価格・二重価格・販売期間）を保存用に取得。(snapshot, missing) を返す"""
    snapshot = {}
    q = ",".join("?" * len(targets))
    with db() as conn:
        rows = conn.execute(f"SELECT * FROM items WHERE manage_number IN ({q})", targets).fetchall()
    found = set()
    for row in rows:
        item = get_effective(row)
        if not item:
            continue
        found.add(row["manage_number"])
        prices = {}
        for vid, v in (item.get("variants") or {}).items():
            if isinstance(v, dict) and to_price(v.get("standardPrice")) is not None:
                # 元の値をそのまま保存する（型も含めて完全に戻せるように）
                prices[vid] = {"standardPrice": v.get("standardPrice"),
                               "referencePrice": v.get("referencePrice")}
        snapshot[row["manage_number"]] = {
            "prices": prices,
            "purchasablePeriod": item.get("purchasablePeriod"),
        }
    missing = [t for t in targets if t not in found]
    return snapshot, missing


def _restore_variant(v, sv):
    """スナップショット値でSKUを復元。変更があればTrue"""
    changed = False
    if isinstance(sv, dict):
        sp = sv.get("standardPrice")
        if sp is not None and v.get("standardPrice") != sp:
            v["standardPrice"] = sp
            changed = True
        rp = sv.get("referencePrice")
        if rp is None:
            if "referencePrice" in v:
                v.pop("referencePrice")
                changed = True
        elif v.get("referencePrice") != rp:
            v["referencePrice"] = rp
            changed = True
    else:  # 旧形式（数値のみ）
        if v.get("standardPrice") != sv:
            v["standardPrice"] = sv
            changed = True
    return changed


def job_scheduled_restore(job_id):
    try:
        with db() as conn:
            job = conn.execute("SELECT * FROM jobs WHERE id=?", (job_id,)).fetchone()
        if not job or job["status"] != "running":
            finish_job("预约不存在或当前状态不能执行")
            return
        client = get_client()
        targets = json.loads(job["targets"])
        snapshot = json.loads(job["snapshot"])
        restore_price = bool(job["restore_price"])
        period_mode = job["period_mode"]  # none / snapshot / clear
        JOB["total"] = len(targets)
        ok = ng = 0
        details = []
        for mn in targets:
            try:
                snap = snapshot.get(mn) or {}
                current = client.get_item(mn)
                patch = {}
                if restore_price and snap.get("prices"):
                    variants = current.get("variants") or {}
                    changed = False
                    for vid, sv in snap["prices"].items():
                        v = variants.get(vid)
                        if isinstance(v, dict) and _restore_variant(v, sv):
                            changed = True
                    if changed:
                        patch["variants"] = variants
                if period_mode == "snapshot":
                    if current.get("purchasablePeriod") != snap.get("purchasablePeriod"):
                        patch["purchasablePeriod"] = snap.get("purchasablePeriod")
                elif period_mode == "clear":
                    if current.get("purchasablePeriod") is not None:
                        patch["purchasablePeriod"] = None
                if patch:
                    send_with_retry(lambda b: client.patch_item(mn, b), patch)
                    # ローカルDBのrawにも反映して整合を保つ
                    with db() as conn:
                        row = conn.execute("SELECT raw FROM items WHERE manage_number=?", (mn,)).fetchone()
                    if row and row["raw"]:
                        item = json.loads(row["raw"])
                        for k, v in patch.items():
                            if v is None:
                                item.pop(k, None)
                            else:
                                item[k] = v
                        s = item_summary(item)
                        with db() as conn:
                            conn.execute(
                                """UPDATE items SET raw=?, title=?, tagline=?, item_number=?,
                                   hide_item=?, min_price=?, max_price=? WHERE manage_number=?""",
                                (json.dumps(item, ensure_ascii=False), s["title"], s["tagline"],
                                 s["item_number"], s["hide_item"], s["min_price"], s["max_price"], mn))
                    add_log(mn, "预约恢复", "OK", "恢复字段： " + ", ".join(patch.keys()))
                    details.append(f"{mn}: OK ({', '.join(patch.keys())})")
                else:
                    add_log(mn, "预约恢复", "OK", "无变化（已恢复）")
                    details.append(f"{mn}: 无变化")
                ok += 1
            except Exception as e:
                ng += 1
                add_log(mn, "预约恢复", "ERROR", str(e))
                details.append(f"{mn}: ERROR {e}")
                JOB["errors"].append(f"{mn}: {e}")
            JOB["done"] += 1
            JOB["message"] = f"预约恢复 {JOB['done']}/{len(targets)}件（成功{ok} 失败{ng}）"
        status = "done" if ng == 0 else "error"
        with db() as conn:
            conn.execute("UPDATE jobs SET status=?, result=?, executed_at=? WHERE id=?",
                         (status, "\n".join(details)[:8000], now(), job_id))
        finish_job(f"预约恢复 完成：成功{ok}件 / 失败{ng}件")
    except Exception as e:
        with db() as conn:
            conn.execute("UPDATE jobs SET status='error', result=? WHERE id=?", (str(e)[:2000], job_id))
        JOB["errors"].append(str(e))
        finish_job(f"预约恢复错误：{e}")


def _try_start_scheduled(job_id, allowed_status=("pending",)):
    """予約をrunningにしてバックグラウンド実行。実行できなければFalse"""
    with db() as conn:
        ph = ",".join("?" * len(allowed_status))
        cur = conn.execute(
            f"UPDATE jobs SET status='running' WHERE id=? AND status IN ({ph})",
            (job_id, *allowed_status))
        if cur.rowcount == 0:
            return False
    if start_job("scheduled_restore", job_scheduled_restore, job_id):
        return True
    with db() as conn:
        conn.execute("UPDATE jobs SET status=? WHERE id=?", (allowed_status[0], job_id))
    return False


def scheduler_loop():
    while True:
        time.sleep(20)
        try:
            now_s = time.strftime("%Y-%m-%dT%H:%M")
            with db() as conn:
                rows = conn.execute(
                    "SELECT id, run_at FROM jobs WHERE status='pending' AND run_at <= ?",
                    (now_s,)).fetchall()
            for r in rows:
                try:
                    overdue_min = (time.time() - time.mktime(
                        time.strptime(r["run_at"], "%Y-%m-%dT%H:%M"))) / 60.0
                except ValueError:
                    overdue_min = 0
                if overdue_min > OVERDUE_GRACE_MIN:
                    # 時刻を過ぎていても実行する。復元は「セール前の状態に戻す」方向の
                    # 操作なので遅れても安全な一方、保留のままだと販売期間が過ぎた商品が
                    # 買えない状態で放置されてしまうため。
                    add_log("", "预约恢复", "WARN",
                            f"预约#{r['id']} 已超过执行时间{int(overdue_min)}分钟，"
                            "现在自动补执行。")
                _try_start_scheduled(r["id"])
        except Exception:
            pass


@app.route("/api/jobs", methods=["GET", "POST"])
def api_jobs():
    if request.method == "POST":
        data = request.json or {}
        name = (data.get("name") or "").strip() or "恢复预约"
        run_at = (data.get("runAt") or "").strip()
        targets = data.get("targets") or []
        restore_price = 1 if data.get("restorePrice", True) else 0
        period_mode = data.get("periodMode", "none")
        if not re.match(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$", run_at):
            return jsonify({"message": "请指定执行时间"}), 400
        if run_at <= time.strftime("%Y-%m-%dT%H:%M"):
            return jsonify({"message": "执行时间必须晚于当前时间"}), 400
        if not targets:
            return jsonify({"message": "尚未选择目标商品（请在商品列表中勾选）"}), 400
        if not restore_price and period_mode == "none":
            return jsonify({"message": "请选择要恢复的内容（价格或销售期间）"}), 400
        # 現在のローカル値（＝セール設定前の値）をスナップショット
        snapshot, missing = make_snapshot(targets)
        if missing:
            return jsonify({"message": f"包含本机不存在的商品： {', '.join(missing[:5])}"}), 400
        with db() as conn:
            cur = conn.execute(
                """INSERT INTO jobs(name,run_at,targets,snapshot,restore_price,period_mode,status,created_at)
                   VALUES(?,?,?,?,?,?,'pending',?)""",
                (name, run_at, json.dumps(targets), json.dumps(snapshot, ensure_ascii=False),
                 restore_price, period_mode, now()))
            job_id = cur.lastrowid
        return jsonify({"ok": True, "id": job_id,
                        "message": f"预约#{job_id} 已创建（{len(targets)}件・{run_at.replace('T',' ')} 执行）。"
                                   "已保存当前价格和销售期间，之后可应用促销设置。"})
    with db() as conn:
        rows = conn.execute("SELECT id,name,run_at,targets,restore_price,period_mode,status,result,"
                            "created_at,executed_at FROM jobs ORDER BY id DESC LIMIT 100").fetchall()
    jobs = []
    for r in rows:
        d = dict(r)
        d["targetCount"] = len(json.loads(r["targets"]))
        del d["targets"]
        jobs.append(d)
    return jsonify({"jobs": jobs})


@app.route("/api/jobs/<int:job_id>/cancel", methods=["POST"])
def api_job_cancel(job_id):
    with db() as conn:
        cur = conn.execute(
            "UPDATE jobs SET status='canceled' WHERE id=? AND status IN ('pending','needs_confirm','error')",
            (job_id,))
    if cur.rowcount == 0:
        return jsonify({"message": "当前不能取消（正在执行或已经完成）"}), 400
    return jsonify({"ok": True, "message": f"预约#{job_id} 已取消"})


@app.route("/api/jobs/<int:job_id>/run", methods=["POST"])
def api_job_run(job_id):
    if _try_start_scheduled(job_id, allowed_status=("pending", "needs_confirm", "error")):
        return jsonify({"ok": True, "message": f"预约#{job_id} 已开始执行"})
    return jsonify({"message": "无法执行（其他任务正在运行，或当前状态不允许）"}), 409


# ----------------- セール専用機能 -----------------
@app.route("/api/sale/flag", methods=["POST"])
def api_sale_flag():
    """セール対象フラグの設定/解除"""
    data = request.json or {}
    targets = data.get("targets") or []
    value = 1 if data.get("value") else 0
    if data.get("clearAll"):
        with db() as conn:
            conn.execute("UPDATE items SET sale_flag=0")
        return jsonify({"ok": True, "message": "已清空全部促销对象"})
    if not targets:
        return jsonify({"message": "没有目标商品"}), 400
    q = ",".join("?" * len(targets))
    with db() as conn:
        conn.execute(f"UPDATE items SET sale_flag=? WHERE manage_number IN ({q})", [value] + targets)
    return jsonify({"ok": True,
                    "message": f"{len(targets)}件商品的促销标记已{'添加' if value else '移除'}完成"})


def _save_sale_history(name, targets, snapshot):
    """セール履歴を保存（最新2回分のみ保持）"""
    with db() as conn:
        conn.execute("INSERT INTO sale_history(name,targets,snapshot,created_at) VALUES(?,?,?,?)",
                     (name, json.dumps(targets), json.dumps(snapshot, ensure_ascii=False), now()))
        conn.execute("""DELETE FROM sale_history WHERE id NOT IN
                        (SELECT id FROM sale_history ORDER BY id DESC LIMIT 2)""")


@app.route("/api/sale/start", methods=["POST"])
def api_sale_start():
    """セール開始：履歴保存→セール設定をローカル適用→RMSに反映→自動復元予約を作成"""
    if JOB["running"]:
        return jsonify({"message": "其他任务正在运行"}), 409
    data = request.json or {}
    name = (data.get("name") or "").strip() or "促销"
    with db() as conn:
        rows = conn.execute("SELECT manage_number FROM items WHERE sale_flag=1").fetchall()
    targets = [r["manage_number"] for r in rows]
    if not targets:
        return jsonify({"message": "没有促销商品，请在商品列表的“促销”列勾选"}), 400

    price_mode = data.get("priceMode", "")
    price_value = data.get("priceValue", "")
    period_start = (data.get("periodStart") or "").strip()
    period_end = (data.get("periodEnd") or "").strip()
    set_ref = bool(data.get("setRefPrice"))
    ref_type = data.get("refType", "2")
    auto_restore = bool(data.get("autoRestore", True))
    restore_at = (data.get("restoreAt") or "").strip()
    restore_period_mode = data.get("restorePeriodMode", "snapshot")

    if not price_mode and not period_start and not period_end and not set_ref:
        return jsonify({"message": "请设置促销内容（售价、销售期间或对比价至少一项）"}), 400
    if price_mode and price_value == "":
        return jsonify({"message": "请输入价格数值"}), 400
    if not price_mode and str(price_value).strip() != "":
        # 値だけ入れて［値引き方法］が未選択だと、黙って価格変更がスキップされ
        # 「割引が効かない」ように見えるため、ここで止める
        return jsonify({"message": "尚未选择“降价方式”。"
                                   "输入数值后，还需要选择“按比例降价（%）”等方式"}), 400
    if auto_restore:
        if not re.match(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$", restore_at):
            return jsonify({"message": "请指定自动恢复时间"}), 400
        if restore_at <= time.strftime("%Y-%m-%dT%H:%M"):
            return jsonify({"message": "自动恢复时间必须晚于当前时间"}), 400

    # 1) 履歴保存（2回分保持）
    snapshot, missing = make_snapshot(targets)
    if missing:
        return jsonify({"message": f"包含本机不存在的商品： {', '.join(missing[:5])}"}), 400
    _save_sale_history(name, targets, snapshot)

    # 2) セール設定をローカル適用
    q = ",".join("?" * len(targets))
    with db() as conn:
        rows = conn.execute(f"SELECT * FROM items WHERE manage_number IN ({q})", targets).fetchall()
    changed = 0
    for row in rows:
        item = get_effective(row)
        if not item:
            continue
        before = json.dumps(item, ensure_ascii=False)
        for v in (item.get("variants") or {}).values():
            if not isinstance(v, dict):
                continue
            original = v.get("standardPrice")
            base = to_price(original)
            if base is None:
                continue
            if price_mode:
                value = float(price_value)
                if price_mode == "percent":
                    p = round(base * (1 + value / 100.0))
                elif price_mode == "delta":
                    p = base + value
                else:
                    p = value
                v["standardPrice"] = price_like(original, p)
            if set_ref:
                # 元の販売価格を二重価格（表示価格）として設定
                # displayTypeは必ず上書きする。元がSHOP_SETTING（店舗設定に従う）のまま
                # typeを併送するとRMSに IE0151 で商品ごと拒否され、価格も期間も反映されない。
                rp = v.get("referencePrice") or {}
                rp["displayType"] = "REFERENCE_PRICE"
                rp["type"] = int(ref_type) if str(ref_type).lstrip("-").isdigit() else ref_type
                rp["value"] = price_like(rp.get("value", original), base)
                v["referencePrice"] = rp
        if period_start or period_end:
            period = {}
            if period_start:
                period["start"] = period_start
            if period_end:
                period["end"] = period_end
            item["purchasablePeriod"] = period
        if json.dumps(item, ensure_ascii=False) != before:
            save_local_edit(row["manage_number"], item, is_new=bool(row["is_new"]))
            changed += 1

    # 3) 自動復元予約を作成
    job_id = None
    if auto_restore:
        with db() as conn:
            cur = conn.execute(
                """INSERT INTO jobs(name,run_at,targets,snapshot,restore_price,period_mode,status,created_at)
                   VALUES(?,?,?,?,1,?,'pending',?)""",
                (f"{name} 自动恢复", restore_at, json.dumps(targets),
                 json.dumps(snapshot, ensure_ascii=False), restore_period_mode, now()))
            job_id = cur.lastrowid

    # 4) RMSに反映（バックグラウンド）
    started = start_job("apply", job_apply, targets)
    add_log("", "开始促销", "OK",
            f"{name}: 对象{len(targets)}件 修改{changed}件 恢复预约#{job_id or '无'}")
    return jsonify({"ok": True, "jobId": job_id, "changed": changed, "targets": len(targets),
                    "applyStarted": started,
                    "message": f"已应用促销设置（对象{len(targets)}件）。"
                               + (f"已开始提交到RMS。" if started else "请进入“提交到RMS”执行提交。")
                               + (f" 自动恢复预约#{job_id}（{restore_at.replace('T',' ')}）已创建。" if job_id else "")})


# ---------- 楽天スーパーSALEサーチ イベント商品申請 ----------
# 一括申請CSVの列は公式マニュアル（店舗運営Navi 000044479）記載の2列。
#   A列: 申請するページの商品管理番号
#   B列: コピー元商品管理番号（コピーページで申請する場合のみ入力）
SALE_APP_COLUMNS = ["申請するページの商品管理番号", "コピー元商品管理番号"]


def _discount_rate(item):
    """二重価格(referencePrice)を基準にした割引率(%)を返す。判定できなければNone。
    スーパーSALEサーチは 10%以上=通常サーチ / 50%以上=半額サーチ の掲載条件。"""
    best = None
    for v in (item.get("variants") or {}).values():
        if not isinstance(v, dict):
            continue
        sp = to_price(v.get("standardPrice"))
        ref = to_price((v.get("referencePrice") or {}).get("value"))
        if sp is None or ref is None or ref <= 0:
            continue
        rate = (ref - sp) / ref * 100.0
        if best is None or rate > best:
            best = rate
    return best


def _sale_application_rows():
    """セール対象フラグが立っている商品を、申請用の行として返す"""
    with db() as conn:
        rows = conn.execute(
            "SELECT * FROM items WHERE sale_flag=1 ORDER BY manage_number").fetchall()
    out = []
    for row in rows:
        item = get_effective(row) or {}
        rate = _discount_rate(item)
        if rate is None:
            category = "unknown"
        elif rate >= 50:
            category = "half"      # 半額サーチ
        elif rate >= 10:
            category = "normal"    # 通常セールサーチ
        else:
            category = "ng"        # 掲載条件を満たさない
        variants = [v for v in (item.get("variants") or {}).values() if isinstance(v, dict)]
        prices = [p for p in (to_price(v.get("standardPrice")) for v in variants) if p is not None]
        refs = [p for p in (to_price((v.get("referencePrice") or {}).get("value")) for v in variants)
                if p is not None]
        out.append({
            "manageNumber": row["manage_number"],
            "title": item.get("title", ""),
            "price": min(prices) if prices else None,
            "referencePrice": max(refs) if refs else None,
            "rate": None if rate is None else round(rate, 1),
            "category": category,
            "hidden": bool(item.get("hideItem")),
        })
    return out


@app.route("/api/sale/application")
def api_sale_application():
    rows = _sale_application_rows()
    summary = {k: sum(1 for r in rows if r["category"] == k)
               for k in ("half", "normal", "ng", "unknown")}
    return jsonify({"rows": rows, "summary": summary, "columns": SALE_APP_COLUMNS})


@app.route("/api/sale/application-csv", methods=["POST"])
def api_sale_application_csv():
    """申請用CSVを生成。
    RMSの一括申請は「アップロードした内容で全件上書き」される仕様のため、
    RMSからダウンロードした申請済みCSVを渡せばマージして出力する。"""
    only = (request.form.get("only") or "all").strip()
    col_a = (request.form.get("colA") or SALE_APP_COLUMNS[0]).strip() or SALE_APP_COLUMNS[0]
    col_b = (request.form.get("colB") or SALE_APP_COLUMNS[1]).strip() or SALE_APP_COLUMNS[1]

    rows = _sale_application_rows()
    if only == "eligible":
        rows = [r for r in rows if r["category"] in ("half", "normal")]
    elif only == "half":
        rows = [r for r in rows if r["category"] == "half"]

    # 既存の申請済みCSV（RMSでダウンロードしたもの）をマージ
    merged, seen = [], set()
    f = request.files.get("existing")
    if f:
        raw = f.read()
        text = None
        for enc in ("cp932", "utf-8-sig", "utf-8"):
            try:
                text = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            return jsonify({"message": "无法识别现有CSV的编码"}), 400
        for i, r in enumerate(csv.reader(io.StringIO(text))):
            if not r or not (r[0] or "").strip():
                continue
            a = r[0].strip()
            if i == 0 and ("商品管理番号" in a):   # ヘッダー行は読み飛ばす
                continue
            if a in seen:
                continue
            seen.add(a)
            merged.append([a, (r[1].strip() if len(r) > 1 else "")])

    for r in rows:
        if r["manageNumber"] in seen:
            continue
        seen.add(r["manageNumber"])
        merged.append([r["manageNumber"], ""])

    if not merged:
        # 中身が0件のCSVを黙って落とすと「何も入っていない」と混乱するため止める
        total_flagged = len(_sale_application_rows())
        if total_flagged == 0:
            msg = ("尚未标记任何促销商品。"
                   "请先在“商品列表”的促销列中勾选商品。")
        else:
            msg = (f"没有符合条件的商品（促销对象{total_flagged}件均未达到本工具的门槛）。"
                   "相对于对比价需降价10%及以上。"
                   "将导出范围设为“全部促销商品”可导出全部商品。")
        return jsonify({"message": msg}), 400

    buf = io.StringIO()
    w = csv.writer(buf, lineterminator="\r\n")
    w.writerow([col_a, col_b])
    w.writerows(merged)
    data = io.BytesIO(buf.getvalue().encode("cp932", errors="replace"))
    add_log("", "生成申请CSV", "ok", f"{len(merged)}件（其中促销商品{len(rows)}件）")
    return send_file(data, mimetype="text/csv", as_attachment=True,
                     download_name=f"event_application_{time.strftime('%Y%m%d')}.csv")


@app.route("/api/sale/history")
def api_sale_history():
    with db() as conn:
        rows = conn.execute("SELECT id,name,targets,created_at FROM sale_history ORDER BY id DESC").fetchall()
    return jsonify({"history": [
        {"id": r["id"], "name": r["name"], "created_at": r["created_at"],
         "targetCount": len(json.loads(r["targets"]))} for r in rows]})


@app.route("/api/sale/restore", methods=["POST"])
def api_sale_restore():
    """履歴からの復元：即時実行または日時指定で予約作成"""
    data = request.json or {}
    hid = data.get("historyId")
    when = data.get("when", "now")  # now / schedule
    run_at = (data.get("runAt") or "").strip()
    period_mode = data.get("periodMode", "snapshot")
    with db() as conn:
        h = conn.execute("SELECT * FROM sale_history WHERE id=?", (hid,)).fetchone()
    if not h:
        return jsonify({"message": "未找到历史记录"}), 404
    if when == "schedule":
        if not re.match(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$", run_at) or run_at <= time.strftime("%Y-%m-%dT%H:%M"):
            return jsonify({"message": "请选择晚于当前时间的时间"}), 400
    else:
        run_at = time.strftime("%Y-%m-%dT%H:%M")
    with db() as conn:
        cur = conn.execute(
            """INSERT INTO jobs(name,run_at,targets,snapshot,restore_price,period_mode,status,created_at)
               VALUES(?,?,?,?,1,?,'pending',?)""",
            (f"历史恢复： {h['name']}", run_at, h["targets"], h["snapshot"], period_mode, now()))
        job_id = cur.lastrowid
    if when == "now":
        if _try_start_scheduled(job_id):
            return jsonify({"ok": True, "message": f"历史记录“{h['name']}”的恢复已开始（预约#{job_id}）"})
        return jsonify({"message": "其他任务正在运行，请在完成后进入“预约”执行"}), 409
    return jsonify({"ok": True, "message": f"恢复预约#{job_id} 已创建（{run_at.replace('T',' ')} 执行）"})


# ----------------- ルーティング -----------------
@app.route("/")
def index():
    return render_template("index.html", version=APP_VERSION)


@app.route("/favicon.ico")
def favicon():
    path = os.path.join(RESOURCE_DIR, "app.ico")
    if os.path.exists(path):
        return send_file(path, mimetype="image/x-icon")
    return "", 404


@app.route("/api/shutdown", methods=["POST"])
def api_shutdown():
    """ブラウザの［終了］ボタンからツール本体を停止する。
    コンソール非表示のEXEでは、これが唯一の終了手段になる。"""
    def _die():
        time.sleep(0.4)   # レスポンスを返しきってから落とす
        os._exit(0)
    threading.Thread(target=_die, daemon=True).start()
    return jsonify({"ok": True})


@app.route("/api/settings", methods=["GET", "POST"])
def api_settings():
    if request.method == "POST":
        cfg = load_config()
        data = request.json or {}
        for k in ("serviceSecret", "licenseKey", "demoMode", "waitMs"):
            if k in data:
                cfg[k] = data[k]
        save_config(cfg)
    cfg = load_config()
    masked = dict(cfg)
    if masked.get("serviceSecret"):
        masked["serviceSecret"] = masked["serviceSecret"][:4] + "****"
    if masked.get("licenseKey"):
        masked["licenseKey"] = masked["licenseKey"][:4] + "****"
    masked["hasKeys"] = bool(cfg.get("serviceSecret") and cfg.get("licenseKey"))
    return jsonify(masked)


@app.route("/api/test-connection", methods=["POST"])
def api_test():
    try:
        get_client().test_connection()
        return jsonify({"ok": True, "message": "连接成功"})
    except Exception as e:
        return jsonify({"ok": False, "message": str(e)}), 400


@app.route("/api/download", methods=["POST"])
def api_download():
    if not start_job("download", job_download):
        return jsonify({"ok": False, "message": "其他任务正在运行"}), 409
    return jsonify({"ok": True})


@app.route("/api/job")
def api_job():
    return jsonify(JOB)


def build_item_filter(args):
    """一覧・全選択で共通の絞り込み条件を組み立てる"""
    where, params = [], []
    q = args.get("q", "").strip()
    qfield = args.get("qfield", "all")
    if q:
        like = f"%{q}%"
        if qfield == "title":
            where.append("title LIKE ?"); params.append(like)
        elif qfield == "manage_number":
            where.append("manage_number LIKE ?"); params.append(like)
        elif qfield == "item_number":
            where.append("item_number LIKE ?"); params.append(like)
        elif qfield == "tagline":
            where.append("tagline LIKE ?"); params.append(like)
        elif qfield == "description":
            # 説明文はJSON本体を対象に検索
            where.append("COALESCE(edited, raw) LIKE ?"); params.append(like)
        else:
            where.append("(manage_number LIKE ? OR title LIKE ? OR tagline LIKE ? OR item_number LIKE ?)")
            params += [like] * 4
    only = args.get("only", "")  # edited / new / hidden / clean
    if only == "edited":
        where.append("(edited IS NOT NULL OR is_new=1)")
    elif only == "new":
        where.append("is_new=1")
    elif only == "hidden":
        where.append("hide_item=1")
    elif only == "clean":
        where.append("(edited IS NULL AND is_new=0)")
    elif only == "sale":
        where.append("sale_flag=1")
    # 列フィルター（Excel風）: <col>_vals = {"mode":"in"|"notin","vals":[...]} のJSON
    for col in ("manage_number", "title", "tagline", "item_number"):
        spec_raw = args.get(f"{col}_vals", "")
        if spec_raw:
            try:
                spec = json.loads(spec_raw)
            except ValueError:
                spec = None
            vals = (spec or {}).get("vals") or []
            if vals:
                vals = [str(v) for v in vals][:900]
                ph = ",".join("?" * len(vals))
                op = "NOT IN" if (spec or {}).get("mode") == "notin" else "IN"
                where.append(f"COALESCE({col},'') {op} ({ph})")
                params += vals
        text = args.get(f"{col}_text", "").strip()
        if text:
            where.append(f"{col} LIKE ?")
            params.append(f"%{text}%")
    hide = args.get("hide", "")  # 1=倉庫のみ / 0=表示のみ
    if hide in ("0", "1"):
        where.append("hide_item=?"); params.append(int(hide))
    pmin = args.get("price_min", "").strip()
    pmax = args.get("price_max", "").strip()
    if pmin:
        where.append("max_price >= ?"); params.append(float(pmin))
    if pmax:
        where.append("min_price <= ?"); params.append(float(pmax))
    genre = args.get("genre", "").strip()
    if genre:
        where.append("json_extract(COALESCE(edited, raw), '$.genreId') = ?")
        params.append(genre)
    return ("WHERE " + " AND ".join(where)) if where else "", params


@app.route("/api/items")
def api_items():
    page = max(int(request.args.get("page", 1)), 1)
    per = min(int(request.args.get("per", 50)), 200)
    wsql, params = build_item_filter(request.args)
    with db() as conn:
        total = conn.execute(f"SELECT COUNT(*) c FROM items {wsql}", params).fetchone()["c"]
        rows = conn.execute(
            f"SELECT manage_number,title,tagline,item_number,hide_item,min_price,max_price,"
            f"(edited IS NOT NULL) edited,is_new,sale_flag FROM items {wsql} ORDER BY manage_number LIMIT ? OFFSET ?",
            params + [per, (page - 1) * per]).fetchall()
    return jsonify({"total": total, "page": page, "per": per,
                    "items": [dict(r) for r in rows]})


FACET_COLUMNS = ("manage_number", "title", "tagline", "item_number")


@app.route("/api/facets")
def api_facets():
    """列フィルター用：指定列の値一覧と件数を返す（他の列の絞り込みを反映、自列の絞り込みは除外）"""
    col = request.args.get("column", "")
    if col not in FACET_COLUMNS:
        return jsonify({"message": "invalid column"}), 400
    args = {k: v for k, v in request.args.items()
            if k not in (f"{col}_vals", f"{col}_text", "column")}
    wsql, params = build_item_filter(args)
    with db() as conn:
        rows = conn.execute(
            f"SELECT COALESCE({col},'') v, COUNT(*) c FROM items {wsql} "
            f"GROUP BY COALESCE({col},'') ORDER BY v LIMIT 501", params).fetchall()
    truncated = len(rows) > 500
    return jsonify({"values": [{"v": r["v"], "c": r["c"]} for r in rows[:500]],
                    "truncated": truncated})


@app.route("/api/item-ids")
def api_item_ids():
    """絞り込み条件に一致する全商品の管理番号を返す（ページをまたぐ全選択用）"""
    wsql, params = build_item_filter(request.args)
    with db() as conn:
        rows = conn.execute(
            f"SELECT manage_number FROM items {wsql} ORDER BY manage_number LIMIT 20000",
            params).fetchall()
    return jsonify({"ids": [r["manage_number"] for r in rows]})


@app.route("/api/items/<mn>", methods=["GET", "PUT", "DELETE"])
def api_item(mn):
    with db() as conn:
        row = conn.execute("SELECT * FROM items WHERE manage_number=?", (mn,)).fetchone()
    if request.method == "GET":
        if not row:
            return jsonify({"message": "not found"}), 404
        return jsonify({"item": get_effective(row), "is_new": row["is_new"],
                        "edited": bool(row["edited"])})
    if request.method == "DELETE":
        # ローカルの編集/新規を破棄（RMSからは削除しない）
        with db() as conn:
            if row and row["is_new"]:
                conn.execute("DELETE FROM items WHERE manage_number=?", (mn,))
            elif row:
                raw = json.loads(row["raw"])
                s = item_summary(raw)
                conn.execute(
                    """UPDATE items SET edited=NULL, title=?, tagline=?, item_number=?,
                       hide_item=?, min_price=?, max_price=? WHERE manage_number=?""",
                    (s["title"], s["tagline"], s["item_number"], s["hide_item"],
                     s["min_price"], s["max_price"], mn))
        return jsonify({"ok": True})
    # PUT: ローカル保存
    data = request.json or {}
    item = data.get("item")
    if not item:
        return jsonify({"message": "缺少item商品数据"}), 400
    item["manageNumber"] = mn
    save_local_edit(mn, item, is_new=(row["is_new"] if row else True) if data.get("isNew") is None else data["isNew"])
    return jsonify({"ok": True})


@app.route("/api/items/<mn>/price", methods=["POST"])
def api_item_price(mn):
    """商品一覧からの販売価格インライン編集。全SKUを同じ価格にする。
    SKUごとに価格が違う商品は誤って揃えてしまうため受け付けない。"""
    data = request.json or {}
    try:
        price = int(round(float(data.get("price"))))
    except (TypeError, ValueError):
        return jsonify({"message": "价格必须为数值"}), 400
    if price < 1:
        return jsonify({"message": "价格必须为1日元及以上"}), 400

    with db() as conn:
        row = conn.execute("SELECT * FROM items WHERE manage_number=?", (mn,)).fetchone()
    if not row:
        return jsonify({"message": "未找到商品"}), 404
    item = get_effective(row)
    if not item:
        return jsonify({"message": "没有商品数据"}), 400

    targets = [v for v in (item.get("variants") or {}).values()
               if isinstance(v, dict) and to_price(v.get("standardPrice")) is not None]
    if not targets:
        return jsonify({"message": "没有设置价格的SKU，请进入“编辑”窗口设置"}), 400
    if len({to_price(v["standardPrice"]) for v in targets}) > 1:
        return jsonify({"message": "各SKU价格不同，请进入“编辑”窗口分别修改"}), 400

    before = to_price(targets[0]["standardPrice"])
    for v in targets:
        v["standardPrice"] = price_like(v.get("standardPrice"), price)
    save_local_edit(mn, item, is_new=bool(row["is_new"]))
    add_log(mn, "修改列表价格", "ok", f"{before}日元 → {price}日元")
    return jsonify({"ok": True, "price": price, "skuCount": len(targets)})


def _get_by_path(obj, path):
    cur = obj
    for p in path.split("."):
        if not isinstance(cur, dict) or p not in cur:
            return None
        cur = cur[p]
    return cur


def _set_by_path(obj, path, value):
    parts = path.split(".")
    cur = obj
    for p in parts[:-1]:
        cur = cur.setdefault(p, {})
    cur[parts[-1]] = value


@app.route("/api/bulk-edit", methods=["POST"])
def api_bulk_edit():
    data = request.json or {}
    targets = data.get("targets") or []
    op = data.get("op") or {}
    if not targets:
        return jsonify({"message": "尚未选择目标商品"}), 400
    q = ",".join("?" * len(targets))
    with db() as conn:
        rows = conn.execute(f"SELECT * FROM items WHERE manage_number IN ({q})", targets).fetchall()
    changed = 0
    for row in rows:
        item = get_effective(row)
        if not item:
            continue
        before = json.dumps(item, ensure_ascii=False)
        kind = op.get("kind")
        if kind == "replace":
            for field in op.get("fields") or []:
                val = _get_by_path(item, field)
                if isinstance(val, str) and op.get("search"):
                    _set_by_path(item, field, val.replace(op["search"], op.get("replace", "")))
        elif kind == "prepend_append":
            for field in op.get("fields") or []:
                val = _get_by_path(item, field) or ""
                if isinstance(val, str):
                    _set_by_path(item, field, op.get("prepend", "") + val + op.get("append", ""))
        elif kind == "price":
            mode = op.get("mode")
            value = float(op.get("value", 0))
            for v in (item.get("variants") or {}).values():
                if not isinstance(v, dict):
                    continue
                original = v.get("standardPrice")
                p = to_price(original)
                if p is None:
                    continue
                if mode == "percent":
                    p = round(p * (1 + value / 100.0))
                elif mode == "delta":
                    p = p + value
                elif mode == "fixed":
                    p = value
                v["standardPrice"] = price_like(original, p)
        elif kind == "refprice":
            # 二重価格（表示価格）の一括設定/解除
            for v in (item.get("variants") or {}).values():
                if not isinstance(v, dict):
                    continue
                if op.get("clear"):
                    v.pop("referencePrice", None)
                    continue
                mode = op.get("mode", "fixed")
                value = float(op.get("value", 0))
                if mode == "percent":
                    base = to_price(v.get("standardPrice"))
                    if base is None:
                        continue
                    val = base * (1 + value / 100.0)
                else:
                    val = value
                t = op.get("type", 1)
                # displayTypeは必ず上書きする（SHOP_SETTINGのまま type を送ると IE0151）
                rp = v.get("referencePrice") or {}
                rp["displayType"] = "REFERENCE_PRICE"
                rp["type"] = int(t) if str(t).lstrip("-").isdigit() else t
                rp["value"] = price_like(rp.get("value", v.get("standardPrice")), val)
                v["referencePrice"] = rp
        elif kind == "hide":
            item["hideItem"] = bool(op.get("value"))
        elif kind == "period":
            if op.get("clear"):
                item["purchasablePeriod"] = None
            else:
                period = {}
                if op.get("start"):
                    period["start"] = op["start"]
                if op.get("end"):
                    period["end"] = op["end"]
                item["purchasablePeriod"] = period or None
        else:
            return jsonify({"message": f"未知操作： {kind}"}), 400
        if json.dumps(item, ensure_ascii=False) != before:
            save_local_edit(row["manage_number"], item, is_new=bool(row["is_new"]))
            changed += 1
    return jsonify({"ok": True, "changed": changed,
                    "message": f"{changed}件已保存到本机（执行“提交到RMS”后才会提交）"})


@app.route("/api/apply", methods=["POST"])
def api_apply():
    data = request.json or {}
    targets = data.get("targets") or []
    if not start_job("apply", job_apply, targets):
        return jsonify({"ok": False, "message": "其他任务正在运行"}), 409
    return jsonify({"ok": True})


@app.route("/api/duplicate", methods=["POST"])
def api_duplicate():
    data = request.json or {}
    src, new = data.get("source"), (data.get("newManageNumber") or "").strip()
    if not re.match(r"^[a-z0-9\-_]{1,32}$", new or ""):
        return jsonify({"message": "新商品管理编号限32个字符，仅可用半角小写英文字母、数字、连字符和下划线"}), 400
    with db() as conn:
        row = conn.execute("SELECT * FROM items WHERE manage_number=?", (src,)).fetchone()
        dup = conn.execute("SELECT 1 FROM items WHERE manage_number=?", (new,)).fetchone()
    if not row:
        return jsonify({"message": "未找到来源商品"}), 404
    if dup:
        return jsonify({"message": f"商品管理编号 {new} 已存在"}), 400
    item = get_effective(row)
    item["manageNumber"] = new
    if data.get("newTitle"):
        item["title"] = data["newTitle"]
    save_local_edit(new, item, is_new=True)
    return jsonify({"ok": True, "message": f"{new} 已复制为新商品（尚未提交）"})


CSV_COLUMNS = [
    "manageNumber", "itemNumber", "title", "tagline",
    "productDescriptionPc", "productDescriptionSp", "salesDescription",
    "genreId", "taxIncluded", "hideItem",
    "imageLocation1", "imageLocation2", "imageLocation3",
    "variantId", "standardPrice", "referencePrice", "quantity",
]


@app.route("/api/csv-template")
def api_csv_template():
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(CSV_COLUMNS)
    w.writerow(["sample-001", "ITEM-0001", "サンプル商品名", "キャッチコピー",
                "<p>PC用商品説明文</p>", "スマホ用商品説明文", "<p>PC用販売説明文</p>",
                "100000", "1", "0", "/folder/image1.jpg", "", "",
                "sku-001", "1980", "2980", "10"])
    data = io.BytesIO(buf.getvalue().encode("utf-8-sig"))
    return send_file(data, mimetype="text/csv", as_attachment=True,
                     download_name="item_template.csv")


@app.route("/api/csv-import", methods=["POST"])
def api_csv_import():
    f = request.files.get("file")
    if not f:
        return jsonify({"message": "未提供文件"}), 400
    raw = f.read()
    text = None
    for enc in ("utf-8-sig", "cp932", "utf-8"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        return jsonify({"message": "无法识别编码，请使用UTF-8或Shift_JIS保存"}), 400
    reader = csv.DictReader(io.StringIO(text))
    created, errors, skipped = 0, [], 0
    for i, r in enumerate(reader, 2):
        try:
            mn = (r.get("manageNumber") or "").strip()
            if not mn:
                continue
            # 既存商品をここで取り込むと、CSVの17列しか持たない「骨組み商品」で
            # 上書き登録され、反映時のupsertで本物の商品ページを破壊してしまう。
            # 既存商品の一括更新は［CSVで既存商品を更新］を使う。
            with db() as conn:
                exists = conn.execute(
                    "SELECT 1 FROM items WHERE manage_number=? AND is_new=0", (mn,)).fetchone()
            if exists:
                skipped += 1
                if len(errors) < 50:
                    errors.append(f"{i}行： {mn} 是已有商品，已跳过"
                                  "（如需修改，请使用“通过CSV更新已有商品”）")
                continue
            item = {
                "manageNumber": mn,
                "itemNumber": (r.get("itemNumber") or "").strip() or None,
                "title": (r.get("title") or "").strip(),
                "tagline": (r.get("tagline") or "").strip() or None,
                "productDescription": {
                    "pc": r.get("productDescriptionPc") or "",
                    "sp": r.get("productDescriptionSp") or "",
                },
                "salesDescription": r.get("salesDescription") or None,
                "itemType": "NORMAL",
                "genreId": (r.get("genreId") or "").strip(),
                "hideItem": (r.get("hideItem") or "0").strip() in ("1", "true", "TRUE"),
                "payment": {"taxIncluded": (r.get("taxIncluded") or "1").strip() in ("1", "true", "TRUE")},
            }
            images = []
            for k in ("imageLocation1", "imageLocation2", "imageLocation3"):
                loc = (r.get(k) or "").strip()
                if loc:
                    images.append({"type": "CABINET", "location": loc})
            if images:
                item["images"] = images
            vid = (r.get("variantId") or "sku-1").strip()
            variant = {}
            sp = to_price(r.get("standardPrice"))
            if sp is not None:
                variant["standardPrice"] = price_like(None, sp)
            rv = to_price(r.get("referencePrice"))
            if rv is not None:
                variant["referencePrice"] = {"displayType": "REFERENCE_PRICE",
                                             "value": price_like(None, rv)}
            item["variants"] = {vid: variant}
            if (r.get("quantity") or "").strip():
                item["_initialQuantity"] = int(float(r["quantity"]))
            item = {k: v for k, v in item.items() if v is not None}
            save_local_edit(mn, item, is_new=True)
            created += 1
        except Exception as e:
            errors.append(f"{i}行： {e}")
    msg = f"{created}件已导入（新增，尚未提交）"
    if skipped:
        msg += f" / 已有商品已跳过{skipped}件"
    if len(errors) > skipped:
        msg += f" / 错误{len(errors) - skipped}件"
    return jsonify({"ok": True, "created": created, "skipped": skipped,
                    "errors": errors, "message": msg})


@app.route("/api/csv-update", methods=["POST"])
def api_csv_update():
    """CSVで既存商品を一括更新する。
    ［エクスポート → Excelで編集 → 取り込み → RMSに反映］の運用向け。

    ・manageNumber でDBの既存商品にマッチする。無い番号は作らずスキップして報告
    ・空欄の列は「変更しない」。列そのものが無くても同じ（商品名だけの2列CSVでも動く）
    ・画像とquantityは扱わない。3列しかない画像列で全画像を置き換える事故を防ぐため
    """
    f = request.files.get("file")
    if not f:
        return jsonify({"message": "未提供文件"}), 400
    raw = f.read()
    text = None
    for enc in ("utf-8-sig", "cp932", "utf-8"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        return jsonify({"message": "无法识别编码，请使用UTF-8或Shift_JIS保存"}), 400

    def cell(row, key):
        """空欄・列なしは None（＝変更しない）。値があれば文字列を返す"""
        if key not in row:
            return None
        v = row.get(key)
        if v is None:
            return None
        v = v.strip()
        return v if v != "" else None

    reader = csv.DictReader(io.StringIO(text))
    updated, unchanged, missing, errors = 0, 0, [], []
    for i, r in enumerate(reader, 2):
        try:
            mn = (r.get("manageNumber") or "").strip()
            if not mn:
                continue
            with db() as conn:
                row = conn.execute("SELECT * FROM items WHERE manage_number=?", (mn,)).fetchone()
            if not row:
                missing.append(mn)
                if len(errors) < 50:
                    errors.append(f"{i}行： {mn} 尚未下载，已跳过")
                continue

            item = get_effective(row)
            if not item:
                errors.append(f"{i}行： {mn} 缺少商品数据")
                continue
            before = json.dumps(item, ensure_ascii=False, sort_keys=True)

            for key, path in (("title", "title"), ("tagline", "tagline"),
                              ("itemNumber", "itemNumber"), ("genreId", "genreId"),
                              ("salesDescription", "salesDescription")):
                val = cell(r, key)
                if val is not None:
                    item[path] = val
            for key, sub in (("productDescriptionPc", "pc"), ("productDescriptionSp", "sp")):
                val = cell(r, key)
                if val is not None:
                    pd = item.get("productDescription")
                    if not isinstance(pd, dict):
                        pd = {}
                    pd[sub] = val
                    item["productDescription"] = pd
            val = cell(r, "hideItem")
            if val is not None:
                item["hideItem"] = val in ("1", "true", "TRUE")
            val = cell(r, "taxIncluded")
            if val is not None:
                pay = item.get("payment")
                if not isinstance(pay, dict):
                    pay = {}
                pay["taxIncluded"] = val in ("1", "true", "TRUE")
                item["payment"] = pay

            # 価格。variantId列でSKUを指定、空欄なら先頭SKU
            variants = item.get("variants") or {}
            vid = cell(r, "variantId")
            if vid is None:
                vid = next(iter(variants), None)
            if vid is not None and vid in variants and isinstance(variants[vid], dict):
                v = variants[vid]
                sp = to_price(cell(r, "standardPrice"))
                if sp is not None:
                    v["standardPrice"] = price_like(v.get("standardPrice"), sp)
                rv = to_price(cell(r, "referencePrice"))
                if rv is not None:
                    cur = to_price((v.get("referencePrice") or {}).get("value"))
                    # 値が変わるときだけ触る。同値で書き換えると displayType が
                    # SHOP_SETTING から REFERENCE_PRICE に変わってしまう
                    if cur is None or abs(cur - rv) >= 0.5:
                        rp = v.get("referencePrice") or {}
                        rp["displayType"] = "REFERENCE_PRICE"
                        rp.setdefault("type", 2)
                        rp["value"] = price_like(rp.get("value", v.get("standardPrice")), rv)
                        v["referencePrice"] = rp
            elif vid is not None and variants:
                errors.append(f"{i}行： {mn} 中未找到SKU '{vid}' （未修改价格）")

            if json.dumps(item, ensure_ascii=False, sort_keys=True) == before:
                unchanged += 1
                continue
            save_local_edit(mn, item, is_new=bool(row["is_new"]))
            updated += 1
        except Exception as e:
            errors.append(f"{i}行： {e}")

    add_log("", "通过CSV更新已有商品", "ok",
            f"更新{updated}件 / 无变化{unchanged}件 / 不在处理范围{len(missing)}件")
    msg = f"{updated}件已更新（保存到本机，尚未提交）"
    if unchanged:
        msg += f" / 无变化{unchanged}件"
    if missing:
        msg += f" / 尚未下载，已跳过{len(missing)}件"
    msg += "。请核对内容后再执行“提交到RMS”。"
    return jsonify({"ok": True, "updated": updated, "unchanged": unchanged,
                    "missing": len(missing), "errors": errors, "message": msg})


@app.route("/api/csv-export")
def api_csv_export():
    with db() as conn:
        rows = conn.execute("SELECT * FROM items ORDER BY manage_number").fetchall()
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(CSV_COLUMNS)
    for row in rows:
        item = get_effective(row) or {}
        pd = item.get("productDescription") or {}
        images = item.get("images") or []
        locs = [img.get("location", "") for img in images[:3]] + ["", "", ""]
        variants = item.get("variants") or {}
        vid = next(iter(variants), "")
        v = variants.get(vid) or {}
        ref = v.get("referencePrice") or {}
        w.writerow([
            item.get("manageNumber", row["manage_number"]), item.get("itemNumber", ""),
            item.get("title", ""), item.get("tagline", ""),
            pd.get("pc", ""), pd.get("sp", ""), item.get("salesDescription", ""),
            item.get("genreId", ""),
            1 if (item.get("payment") or {}).get("taxIncluded") else 0,
            1 if item.get("hideItem") else 0,
            locs[0], locs[1], locs[2],
            vid, _csv_price(v.get("standardPrice")), _csv_price(ref.get("value")), "",
        ])
    # BOM付きUTF-8。cp932だと「♥」が「?」に、「〜」(U+301C)が「～」(U+FF5E)に化け、
    # 編集していない説明文まで「変更あり」になって反映時に本当に壊れる。
    # ExcelはBOMがあればUTF-8として正しく開く。
    data = io.BytesIO(buf.getvalue().encode("utf-8-sig"))
    return send_file(data, mimetype="text/csv", as_attachment=True,
                     download_name="items_export.csv")


@app.route("/api/reset", methods=["POST"])
def api_reset():
    """配布用リセット：APIキー・商品データ・ログをすべて消去して初期状態に戻す"""
    global _mock
    try:
        if os.path.exists(CONFIG_PATH):
            os.remove(CONFIG_PATH)
    except OSError:
        save_config(dict(DEFAULT_CONFIG))
    with db() as conn:
        conn.execute("DELETE FROM items")
        conn.execute("DELETE FROM logs")
    _mock = None
    return jsonify({"ok": True,
                    "message": "初始化完成。本机API密钥、商品数据和日志已删除，已恢复为演示模式。"})


@app.route("/api/logs")
def api_logs():
    with db() as conn:
        rows = conn.execute("SELECT * FROM logs ORDER BY id DESC LIMIT 300").fetchall()
    return jsonify({"logs": [dict(r) for r in rows]})


@app.route("/api/stats")
def api_stats():
    with db() as conn:
        r = conn.execute(
            "SELECT COUNT(*) total,"
            " SUM(CASE WHEN edited IS NOT NULL OR is_new=1 THEN 1 ELSE 0 END) pending,"
            " SUM(is_new) new_items FROM items").fetchone()
    return jsonify({"total": r["total"], "pending": r["pending"] or 0, "new": r["new_items"] or 0})


def pick_free_port(start):
    """指定ポートが使用中（古いバージョンが起動中など）なら空いているポートを探す"""
    import socket
    for p in range(start, start + 20):
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.bind(("127.0.0.1", p))
            return p
        except OSError:
            continue
    return start


LOG_PATH = os.path.join(BASE_DIR, "起動ログ.txt")


def setup_windowed_io():
    """コンソール非表示のEXE（PyInstaller --noconsole）では sys.stdout/stderr が
    None になり、print() やFlaskのログ出力が例外を起こす。ログファイルへ振り向ける。
    通常のコンソール実行時は何もしない。"""
    if sys.stdout is not None and sys.stderr is not None:
        return False
    try:
        f = open(LOG_PATH, "a", encoding="utf-8", buffering=1)
    except Exception:
        f = open(os.devnull, "w", encoding="utf-8")
    sys.stdout = sys.stderr = f
    print(f"\n===== 启动 {time.strftime('%Y-%m-%d %H:%M:%S')} v{APP_VERSION} =====")
    return True


def show_error_dialog(message):
    """コンソールが無くてもユーザーにエラーを見せる。"""
    if sys.platform == "win32":
        try:
            import ctypes
            # 0x10 = MB_ICONERROR
            ctypes.windll.user32.MessageBoxW(None, message, "RMS商品批量编辑工具", 0x10)
            return
        except Exception:
            pass
    try:
        print(message)
    except Exception:
        pass


if __name__ == "__main__":
    windowed = setup_windowed_io()
    try:
        init_db()
        threading.Thread(target=scheduler_loop, daemon=True).start()
        port = pick_free_port(int(os.environ.get("PORT", 8930)))
        print(f"* RMS商品批量编辑工具 v{APP_VERSION}: http://localhost:{port} 请在浏览器中打开")
        if port != int(os.environ.get("PORT", 8930)):
            print("注意：默认端口已被占用，已使用其他端口启动。"
                  "可能已有其他工具实例正在运行。")
        try:
            threading.Timer(1.0, lambda: webbrowser.open(f"http://localhost:{port}")).start()
        except Exception:
            pass
        app.run(host="127.0.0.1", port=port, debug=False)
    except SystemExit:
        raise
    except Exception:
        tb = traceback.format_exc()
        try:
            print(tb)
        except Exception:
            pass
        last = tb.strip().splitlines()[-1] if tb.strip() else "未知错误"
        msg = ("工具启动失败。\n\n"
               f"{last}\n\n"
               "常见原因：\n"
               "· 工具位于不可写位置（例如ZIP压缩包内或Program Files）\n"
               "· 被安全软件拦截\n\n")
        if windowed:
            msg += f"详细日志：\n{LOG_PATH}"
        show_error_dialog(msg)
        raise SystemExit(1)
