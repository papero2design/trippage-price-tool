import streamlit as st
import pandas as pd
import re, json, time, random, asyncio, io, tempfile, os
from datetime import date, datetime
from pathlib import Path
import aiohttp
import xlsxwriter

st.set_page_config(
    page_title="트립페이지 가격 비교 도구",
    layout="centered",
)

# ── 설정값 ───────────────────────────────────────────────────
CONFIG = {
    'COL_PRICE_PC':       'price_pc',
    'COL_NORMAL_PRICE':   'normal_price',
    'COL_LINK':           'link',
    'PRICE_API_URL':      'https://gw.hanatour.com/package/pkg/api/common/pkgcomprod/getPkgProdInfo/v1.00?_siteId=hanatour',
    'SEARCH_API_URL':     'https://gw.hanatour.com/search/v2/all/search?_siteId=hanatour',
    'PTN_CD':             'PH05117',
    'CONCURRENCY':        12,
    'DELAY_MIN':          1.0,
    'DELAY_MAX':          2.0,
    'CHECKPOINT_EVERY':   100,
    'ABORT_CONSEC_FAILS': 30,
    'AUTO_RETRY_WAIT':    300,
    'AUTO_RETRY_MAX':     10,
}

today = date.today()

# ── 유틸 함수 ─────────────────────────────────────────────────
def classify_link(url):
    if not isinstance(url, str): return None
    if 'all-search' in url and 'keywordCateg=DS' in url: return 'search'
    if re.search(r'[?&]pkgCd=', url, re.IGNORECASE): return 'pkg'
    return None

def extract_pkg_cd(url):
    if not isinstance(url, str): return None
    m = re.search(r'[?&]pkgCd=([^&]+)', url, re.IGNORECASE)
    return m.group(1) if m else None

def extract_dep_date(pkg_cd):
    if not isinstance(pkg_cd, str): return None
    body = pkg_cd[6:]
    m = re.match(r'^(\d{6,8})', body)
    if not m: return None
    raw = m.group(1)
    try:
        if len(raw) == 8: return date(int(raw[:4]), int(raw[4:6]), int(raw[6:8]))
        else:             return date(2000 + int(raw[:2]), int(raw[2:4]), int(raw[4:6]))
    except ValueError: return None

def extract_search_params(url):
    if not isinstance(url, str): return None, None
    kw  = re.search(r'[?&]keyword=([^&]+)', url)
    idx = re.search(r'[?&]idx=(\d+)', url)
    return (kw.group(1) if kw else None), (int(idx.group(1)) if idx else 1)

def make_price_key(row):
    if row['_link_type'] == 'pkg':
        return row['_pkg_cd']
    if row['_link_type'] == 'search':
        kw = row['_search_keyword']
        return f"SEARCH:{kw}" if kw else None
    return None

# ── 비동기 API 함수 ───────────────────────────────────────────
async def fetch_price(session, semaphore, pkg_cd):
    async with semaphore:
        try:
            await asyncio.sleep(random.uniform(CONFIG['DELAY_MIN'], CONFIG['DELAY_MAX']))
            payload = json.dumps({
                'pkgCd': pkg_cd, 'inpPathCd': 'CBP', 'smplYn': 'N',
                'coopYn': 'N', 'resAcceptPtn': {}, 'partnerYn': 'N',
                'ptnCd': CONFIG['PTN_CD'],
            })
            async with session.post(
                CONFIG['PRICE_API_URL'], data=payload,
                headers={
                    'Content-Type': 'application/json', 'Accept': 'application/json',
                    'Referer': 'https://trippage.hanatour.com/',
                    'Origin': 'https://trippage.hanatour.com',
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)',
                },
                timeout=aiohttp.ClientTimeout(total=10),
            ) as resp:
                if resp.status != 200: return pkg_cd, None, f'HTTP {resp.status}'
                data  = await resp.json(content_type=None)
                price = data.get('data', {}).get('adtTotlAmt')
                return pkg_cd, price, None
        except Exception as e:
            return pkg_cd, None, str(e)

async def fetch_search_price(session, semaphore, keyword):
    price_key = f"SEARCH:{keyword}"
    async with semaphore:
        try:
            await asyncio.sleep(random.uniform(CONFIG['DELAY_MIN'], CONFIG['DELAY_MAX']))
            search_payload = json.dumps({
                "header": {"timestamp": datetime.now().strftime('%Y%m%d%H%M%S'),
                           "lang": "ko-KR", "prevPage": "NO-REFERRER"},
                "os": "pc", "domain": "https://trippage.hanatour.com",
                "keyword": keyword, "keywordCateg": "DS", "idx": "1",
                "page": 1, "pageSize": 20, "sort": "D",
                "ptnCd": CONFIG['PTN_CD'], "ptngr": ["008", "014", "003"],
                "resPathCd": "CBP", "isCobrand": "Y", "isCustomCobrand": "N",
                "paymentTypeYn": "Y", "afcnCobrandProdYn": "N",
                "afcnResTrgtDvCd": "00",
                "afcnSlpdAttrCd": ["A", "F", "H", "P", "V"], "appVersion": "",
            })
            async with session.post(
                CONFIG['SEARCH_API_URL'], data=search_payload,
                headers={
                    'Content-Type': 'application/json', 'Accept': 'application/json',
                    'Referer': 'https://trippage.hanatour.com/',
                    'Origin': 'https://trippage.hanatour.com',
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)',
                },
                timeout=aiohttp.ClientTimeout(total=10),
            ) as resp:
                if resp.status != 200: return price_key, None, f'HTTP {resp.status}'
                sdata    = await resp.json(content_type=None)
                contents = sdata.get('data', {}).get('contents', [])
                section  = next((c for c in contents if c.get('attribute') == 'package'), None)
                if not section: return price_key, None, '검색결과없음'
                items = section.get('data', [])
                if not items: return price_key, None, '검색결과없음(data 비어있음)'
                sale_prod_cd = items[0].get('saleProdCd')
                if not sale_prod_cd: return price_key, None, 'saleProdCd 없음'

            await asyncio.sleep(random.uniform(CONFIG['DELAY_MIN'], CONFIG['DELAY_MAX']))
            price_payload = json.dumps({
                'pkgCd': sale_prod_cd, 'inpPathCd': 'CBP', 'smplYn': 'N',
                'coopYn': 'N', 'resAcceptPtn': {}, 'partnerYn': 'N',
                'ptnCd': CONFIG['PTN_CD'],
            })
            async with session.post(
                CONFIG['PRICE_API_URL'], data=price_payload,
                headers={
                    'Content-Type': 'application/json', 'Accept': 'application/json',
                    'Referer': 'https://trippage.hanatour.com/',
                    'Origin': 'https://trippage.hanatour.com',
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)',
                },
                timeout=aiohttp.ClientTimeout(total=10),
            ) as resp:
                if resp.status != 200: return price_key, None, f'가격API HTTP {resp.status}'
                pdata = await resp.json(content_type=None)
                price = pdata.get('data', {}).get('adtTotlAmt')
                if price is None: return price_key, None, 'adtTotlAmt 없음'
                return price_key, int(price), None
        except Exception as e:
            return price_key, None, str(e)

async def _run_all_async(pkg_cds, search_pairs, log_fn, prog_widget, total_count,
                         existing=None, save_ckpt=True, ckpt_path=None):
    semaphore = asyncio.Semaphore(CONFIG['CONCURRENCY'])
    results   = dict(existing) if existing else {}
    errors    = {}
    connector = aiohttp.TCPConnector(limit=CONFIG['CONCURRENCY'])
    done      = len(results)
    prog_widget.value = done

    pkg_cds_todo = [k for k in pkg_cds if k not in results]
    search_todo  = [k for k in search_pairs if f"SEARCH:{k}" not in results]
    skipped      = (len(pkg_cds) - len(pkg_cds_todo)) + (len(search_pairs) - len(search_todo))
    if skipped:
        log_fn(f"   스킵: {skipped:,}건 (이미 완료) / 신규: {len(pkg_cds_todo)+len(search_todo):,}건")

    def save_checkpoint():
        if not ckpt_path: return
        try:
            tmp = Path(str(ckpt_path) + '.tmp')
            tmp.write_text(json.dumps(results, ensure_ascii=False), encoding='utf-8')
            tmp.rename(ckpt_path)
        except Exception as e:
            log_fn(f"   [오류] 체크포인트 저장 실패: {e}")

    async with aiohttp.ClientSession(connector=connector) as session:
        all_tasks = (
            [fetch_price(session, semaphore, cd) for cd in pkg_cds_todo] +
            [fetch_search_price(session, semaphore, kw) for kw in search_todo]
        )
        consec_fails = 0
        aborted      = False

        for coro in asyncio.as_completed(all_tasks):
            key, price, err = await coro
            done += 1
            prog_widget.value = done

            if price is not None:
                results[key] = price
                consec_fails = 0
            else:
                errors[key] = err
                is_definite = (
                    'HTTP 400'    in str(err) or
                    '검색결과없음' in str(err) or
                    'adtAmt 없음' in str(err)
                )
                consec_fails = 0 if is_definite else consec_fails + 1

            if save_ckpt and done % CONFIG['CHECKPOINT_EVERY'] == 0:
                save_checkpoint()
                log_fn(f"   [저장] 체크포인트 ({done:,}/{total_count:,}건)")

            if consec_fails >= CONFIG['ABORT_CONSEC_FAILS']:
                log_fn(f"   [중단] 연속 실패 {consec_fails}건 — 서버 차단 감지, 조기 종료")
                log_fn(f"          완료: {done:,}건 / 미처리: {total_count - done:,}건")
                aborted = True
                break

    if save_ckpt:
        save_checkpoint()
    if aborted:
        log_fn(f"   [저장] 체크포인트 완료 ({len(results):,}건 보존)")
    return results, errors, aborted

# ── Streamlit용 진행 바 래퍼 ──────────────────────────────────
class _ProgressWidget:
    def __init__(self, bar):
        self._max   = 100
        self._value = 0
        self._bar   = bar
        self._last  = 0.0

    @property
    def max(self):
        return self._max

    @max.setter
    def max(self, v):
        self._max = max(v, 1)
        self._bar.progress(0.0, text="진행 중...")

    @property
    def value(self):
        return self._value

    @value.setter
    def value(self, v):
        self._value = v
        now = time.time()
        if self._max > 0 and (now - self._last > 0.5 or v >= self._max):
            self._last = now
            pct = min(v / self._max, 1.0)
            self._bar.progress(pct, text=f"{v:,} / {self._max:,} 건 완료")

# ── 메인 파이프라인 ───────────────────────────────────────────
def run_pipeline(excel_bytes, sheet_name, log_fn, prog_widget, ckpt_dir):
    log_fn("엑셀 파일 읽는 중...")
    try:
        df = pd.read_excel(io.BytesIO(excel_bytes), sheet_name=sheet_name, engine='openpyxl')
    except Exception as e:
        log_fn(f"[오류] 엑셀 로드 실패: {e}")
        raise

    log_fn(f"로드 완료: {len(df):,}행 x {len(df.columns)}열")

    for col in [CONFIG['COL_NORMAL_PRICE'], CONFIG['COL_PRICE_PC'], CONFIG['COL_LINK']]:
        if col not in df.columns:
            log_fn(f"[오류] 필수 컬럼 없음: '{col}' — 시트 이름 또는 컬럼명 확인 필요")
            raise KeyError(col)

    df['원본_normal_price'] = df[CONFIG['COL_NORMAL_PRICE']].copy()
    df['원본_price_pc']     = df[CONFIG['COL_PRICE_PC']].copy()

    log_fn("링크 분류 중...")
    df['_link_type']      = df[CONFIG['COL_LINK']].apply(classify_link)
    df['_pkg_cd']         = df[CONFIG['COL_LINK']].apply(extract_pkg_cd)
    df['_dep_date']       = df['_pkg_cd'].apply(extract_dep_date)
    df['_search_keyword'] = df[CONFIG['COL_LINK']].apply(lambda u: extract_search_params(u)[0])
    df['_search_idx']     = df[CONFIG['COL_LINK']].apply(lambda u: extract_search_params(u)[1])
    df['_price_key']      = df.apply(make_price_key, axis=1)

    pkg_cds      = df[df['_link_type'] == 'pkg']['_pkg_cd'].dropna().unique().tolist()
    search_pairs = df[df['_link_type'] == 'search']['_search_keyword'].dropna().unique().tolist()
    unknown_cnt  = df['_link_type'].isna().sum()

    log_fn(f"   pkg: {len(pkg_cds):,}건 / search: {len(search_pairs):,}건" +
           (f" / 분류불가: {unknown_cnt:,}건" if unknown_cnt else ""))

    def _is_definite(err_str):
        return (
            'HTTP 400'         in err_str or
            '검색결과없음'      in err_str or
            'adtAmt 없음'      in err_str or
            'adtTotlAmt 없음'  in err_str or
            'saleProdCd 없음'  in err_str or
            '가격API HTTP 400' in err_str
        )

    ckpt_path       = Path(ckpt_dir) / 'checkpoint.json'
    retry_fail_path = Path(ckpt_dir) / 'retry_fail.json'

    existing = {}
    if ckpt_path.exists():
        try:
            existing = json.loads(ckpt_path.read_text(encoding='utf-8'))
            log_fn(f"체크포인트 발견: {len(existing):,}건 이어서 진행")
        except Exception:
            pass

    if retry_fail_path.exists():
        try:
            prev_fail    = json.loads(retry_fail_path.read_text(encoding='utf-8'))
            extra_pkg    = [k for k in prev_fail if not k.startswith('SEARCH:') and k not in existing]
            extra_search = [k.split(':', 1)[1] for k in prev_fail if k.startswith('SEARCH:') and k not in existing]
            pkg_cds      = list(dict.fromkeys(pkg_cds + extra_pkg))
            search_pairs = list(dict.fromkeys(search_pairs + extra_search))
            retry_fail_path.unlink()
            log_fn(f"이전 실패 재시도: pkg {len(extra_pkg)}건 / search {len(extra_search)}건 추가")
        except Exception:
            pass

    total = len(pkg_cds) + len(search_pairs)
    log_fn(f"총 API 호출 예정: {total:,}건")
    prog_widget.max   = max(total, 1)
    prog_widget.value = len(existing)

    # ── API 호출 (자동 재시도 루프) ──────────────────────────
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    auto_retry_cnt = 0
    aborted        = True
    price_map      = {}
    error_map      = {}

    while aborted:
        label = f"자동 재시도 {auto_retry_cnt}회차" if auto_retry_cnt > 0 else "1차"
        log_fn(f"API 호출 [{label}] 시작 (동시 {CONFIG['CONCURRENCY']}건)...")
        start = time.time()
        try:
            price_map, error_map, aborted = loop.run_until_complete(
                _run_all_async(pkg_cds, search_pairs, log_fn, prog_widget, total,
                               existing=existing, ckpt_path=ckpt_path)
            )
        except Exception as e:
            log_fn(f"[오류] {type(e).__name__}: {e}")
            raise
        elapsed = time.time() - start
        log_fn(f"[{label}] 완료 ({elapsed:.1f}초) | 성공: {len(price_map):,} / 실패: {len(error_map):,}")

        if aborted:
            auto_retry_cnt += 1
            if auto_retry_cnt > CONFIG['AUTO_RETRY_MAX']:
                log_fn(f"[중단] 자동 재시도 {CONFIG['AUTO_RETRY_MAX']}회 소진 — 재실행이 필요합니다.")
                retry_abort = {k: v for k, v in error_map.items() if not _is_definite(str(v))}
                if retry_abort:
                    retry_fail_path.write_text(
                        json.dumps(retry_abort, ensure_ascii=False), encoding='utf-8'
                    )
                    log_fn(f"   [저장] 실패 목록 {len(retry_abort)}건")
                loop.close()
                return None, None

            wait_secs = CONFIG['AUTO_RETRY_WAIT']
            log_fn(f"서버 차단 감지 — {wait_secs // 60}분 후 자동 재시작 "
                   f"({auto_retry_cnt}/{CONFIG['AUTO_RETRY_MAX']}회차)")
            remaining = wait_secs
            while remaining > 0:
                step = min(30, remaining)
                time.sleep(step)
                remaining -= step
                if remaining > 0:
                    log_fn(f"   재시작까지 {remaining}초 남음...")
            existing = price_map
            prog_widget.max   = max(total, 1)
            prog_widget.value = len(existing)
            log_fn(f"   재시작 — 완료 {len(existing):,}건 / 잔여 {total - len(existing):,}건")

    # ── 완료 후 단기 재시도 (400 제외) ───────────────────────
    err_400   = {k: v for k, v in error_map.items() if 'HTTP 400' in str(v)}
    err_other = {k: v for k, v in error_map.items() if 'HTTP 400' not in str(v)}

    if err_400:
        log_fn(f"   상품없음 확정(400): {len(err_400)}건 — 재시도 생략")

    still_fail = {}
    if err_other:
        log_fn(f"일시차단 실패 {len(err_other)}건 재시도 중...")
        time.sleep(2)
        r_pkg  = [k for k in err_other if not k.startswith('SEARCH:')]
        r_srch = [k.split(':', 1)[1] for k in err_other if k.startswith('SEARCH:')]
        prog_widget.max   = max(len(r_pkg) + len(r_srch), 1)
        prog_widget.value = 0
        before    = set(price_map.keys())
        retry_map, still_fail, _ = loop.run_until_complete(
            _run_all_async(r_pkg, r_srch, log_fn, prog_widget,
                           len(r_pkg) + len(r_srch), save_ckpt=False)
        )
        new_ok = {k: v for k, v in retry_map.items() if k not in before}
        price_map.update(new_ok)
        log_fn(f"   재시도 성공: {len(new_ok)}건 / 최종 실패: {len(still_fail)}건")

    confirmed_missing = (
        {k for k in err_400} |
        {k for k, v in error_map.items() if '검색결과없음' in str(v) or 'adtAmt 없음' in str(v)} |
        {k for k, v in still_fail.items() if _is_definite(str(v))}
    )
    retry_later = {k: v for k, v in still_fail.items() if not _is_definite(str(v))}
    if retry_later:
        retry_fail_path.write_text(
            json.dumps(retry_later, ensure_ascii=False), encoding='utf-8'
        )
        log_fn(f"   [저장] 다음 실행 재시도 목록: {len(retry_later)}건")

    # ── 상태 판정 ─────────────────────────────────────────────
    log_fn("상태 판정 중...")

    def get_status(row):
        key, orig = row['_price_key'], row['원본_normal_price']
        parts = []
        if row['_link_type'] == 'pkg':
            dep = row['_dep_date']
            if dep is None:   parts.append('[확인필요] 날짜파싱불가')
            elif dep < today: parts.append(f'[출발일지남] {dep}')
        if key in price_map:
            api = int(price_map[key])
            if api != int(orig): parts.append(f'[가격수정] {int(orig):,} → {api:,}원')
        elif key is None:                parts.append('[링크오류]')
        elif key in confirmed_missing:   parts.append('[상품없음] 삭제검토필요')
        else:                            parts.append('[미확인] 다음실행시재시도')
        return ' / '.join(parts) if parts else '정상'

    def get_new_price(row):
        key = row['_price_key']
        return int(price_map[key]) if key in price_map else row['원본_normal_price']

    df['상태']         = df.apply(get_status, axis=1)
    df['price_pc']     = df.apply(get_new_price, axis=1)
    df['normal_price'] = df.apply(get_new_price, axis=1)

    stats = {
        '전체':           len(df),
        '가격 수정':      int(df['상태'].str.contains('가격수정').sum()),
        '상품없음(삭제)':  int(df['상태'].str.contains('상품없음').sum()),
        '미확인(재시도)':  int(df['상태'].str.contains('미확인').sum()),
        '출발일 지남':    int(df['상태'].str.contains('출발일지남').sum()),
        '정상':           int(df['상태'].str.fullmatch('정상').sum()),
    }
    loop.close()
    log_fn("처리 완료!")
    return df, stats


def save_excel(df):
    drop_cols = ['_pkg_cd', '_dep_date', '원본_normal_price', '원본_price_pc',
                 '_search_keyword', '_search_idx', '_price_key', '_link_type', '_canon_url']
    result_df = df.drop(columns=[c for c in drop_cols if c in df.columns]).copy()
    cols = list(result_df.columns)
    if '상태' in cols and CONFIG['COL_LINK'] in cols:
        li = cols.index(CONFIG['COL_LINK'])
        cols.insert(li + 1, cols.pop(cols.index('상태')))
    result_df = result_df[cols]

    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        tmp_path = tmp.name

    try:
        wb = xlsxwriter.Workbook(tmp_path, {
            'constant_memory':   True,
            'nan_inf_to_errors': True,
            'strings_to_urls':   False,
        })
        ws = wb.add_worksheet('Sheet1')

        fmt_header      = wb.add_format({'bold': True, 'font_color': '#FFFFFF', 'bg_color': '#1565C0',
                                         'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
        fmt_normal      = wb.add_format({'valign': 'vcenter'})
        fmt_wrap        = wb.add_format({'valign': 'vcenter', 'text_wrap': True})
        fmt_price       = wb.add_format({'bold': True, 'bg_color': '#C8E6C9', 'valign': 'vcenter', 'text_wrap': True})
        fmt_price_cell  = wb.add_format({'bg_color': '#C8E6C9', 'valign': 'vcenter'})
        fmt_expired     = wb.add_format({'bg_color': '#FFCDD2', 'valign': 'vcenter', 'text_wrap': True})
        fmt_noparse     = wb.add_format({'bg_color': '#F5F5F5', 'valign': 'vcenter', 'text_wrap': True})
        fmt_unconfirmed = wb.add_format({'bg_color': '#FFF9C4', 'valign': 'vcenter', 'text_wrap': True})
        fmt_deleted     = wb.add_format({'bold': True, 'font_color': '#FFFFFF', 'bg_color': '#000000',
                                         'valign': 'vcenter', 'text_wrap': True})
        fmt_deleted_cell= wb.add_format({'font_color': '#FFFFFF', 'bg_color': '#000000', 'valign': 'vcenter'})

        header       = list(result_df.columns)
        status_col   = header.index('상태')
        price_pc_col = header.index('price_pc')
        norm_col     = header.index('normal_price')

        for i, col_name in enumerate(header):
            ws.set_column(i, i, 45 if i == status_col else min(len(str(col_name)) + 4, 30))

        ws.set_row(0, 30)
        for i, col_name in enumerate(header):
            ws.write(0, i, col_name, fmt_header)

        for row_idx, row in enumerate(result_df.itertuples(index=False), start=1):
            sv = str(row[status_col]) if row[status_col] is not None else ''
            if sv == '정상':
                ws.set_row(row_idx, 20)
                for ci, val in enumerate(row):
                    ws.write(row_idx, ci, val, fmt_normal)
            elif '상품없음' in sv:
                ws.set_row(row_idx, 20)
                for ci, val in enumerate(row):
                    ws.write(row_idx, ci, val, fmt_deleted if ci == status_col else fmt_deleted_cell)
            elif '미확인' in sv:
                ws.set_row(row_idx, 20)
                for ci, val in enumerate(row):
                    ws.write(row_idx, ci, val, fmt_unconfirmed if ci == status_col else fmt_normal)
            else:
                ws.set_row(row_idx, None)
                for ci, val in enumerate(row):
                    if ci == status_col:
                        if '가격수정' in sv and '출발일지남' in sv: ws.write(row_idx, ci, val, fmt_expired)
                        elif '가격수정'    in sv: ws.write(row_idx, ci, val, fmt_price)
                        elif '출발일지남'  in sv: ws.write(row_idx, ci, val, fmt_expired)
                        elif '날짜파싱불가' in sv: ws.write(row_idx, ci, val, fmt_noparse)
                        else:                     ws.write(row_idx, ci, val, fmt_wrap)
                    elif '가격수정' in sv and ci in (price_pc_col, norm_col):
                        ws.write(row_idx, ci, val, fmt_price_cell)
                    else:
                        ws.write(row_idx, ci, val, fmt_normal)

        ws.freeze_panes(1, 0)
        ws.autofilter(0, 0, len(result_df), len(header) - 1)
        wb.close()

        with open(tmp_path, 'rb') as f:
            result_bytes = f.read()
    finally:
        try:
            os.unlink(tmp_path)
        except Exception:
            pass

    fname = f"결과_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    return fname, result_bytes


# ══════════════════════════════════════════════════════════════
# Streamlit UI
# ══════════════════════════════════════════════════════════════
st.title("트립페이지 가격 비교 도구")
st.caption("EP 파일을 업로드하면 현재 API 가격과 비교하여 결과 파일을 생성합니다.")
st.divider()

uploaded = st.file_uploader(
    "EP 엑셀 파일 업로드 (.xlsx / .xlsm)",
    type=["xlsx", "xlsm"],
    help="97,000-100,000행 기준 약 2-3시간 소요됩니다.",
)
sheet = st.text_input("시트 이름", value="enbt_naver_ep")

run_btn = st.button(
    "가격 비교 실행",
    type="primary",
    disabled=(uploaded is None),
    use_container_width=True,
)

if run_btn and uploaded:
    ckpt_dir    = Path(tempfile.mkdtemp())
    excel_bytes = uploaded.getvalue()

    log_lines       = []
    log_placeholder = st.empty()

    def log_fn(msg):
        ts = datetime.now().strftime('%H:%M:%S')
        log_lines.append(f"[{ts}] {msg}")
        log_placeholder.code('\n'.join(log_lines[-30:]), language=None)

    bar  = st.progress(0.0, text="준비 중...")
    prog = _ProgressWidget(bar)

    try:
        df, stats = run_pipeline(excel_bytes, sheet, log_fn, prog, ckpt_dir)
    except Exception as e:
        st.error(f"오류가 발생했습니다: {e}")
        st.stop()

    if df is None:
        st.warning(
            "자동 재시도 횟수를 모두 소진했습니다. "
            "페이지를 새로고침 후 다시 실행하면 이어서 진행됩니다."
        )
    else:
        bar.progress(1.0, text="완료!")
        st.success("가격 비교 완료!")
        st.divider()

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("전체",          f"{stats['전체']:,}건")
        c2.metric("가격 수정",     f"{stats['가격 수정']:,}건")
        c3.metric("상품없음(삭제)", f"{stats['상품없음(삭제)']:,}건")
        c4.metric("미확인",        f"{stats['미확인(재시도)']:,}건")
        c5.metric("정상",          f"{stats['정상']:,}건")

        st.divider()

        log_fn("엑셀 파일 생성 중...")
        fname, result_bytes = save_excel(df)
        log_fn(f"파일 생성 완료: {fname}")

        st.download_button(
            label="결과 파일 다운로드",
            data=result_bytes,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
