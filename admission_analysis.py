import os
import sys
import json
import pandas as pd

# ── Firebase 설정 ─────────────────────────────────────────────────────────────
# Firebase 콘솔(https://console.firebase.google.com)에서 프로젝트 생성 후
# 프로젝트 설정 > 일반 > 내 앱 > 웹 앱 추가 에서 아래 값들을 복사해 넣으세요.
# Firestore 데이터베이스를 생성하고 'whitelist' 컬렉션에 허용할 이메일을
# 문서 ID로 추가하세요 (예: 문서 ID = "teacher@school.kr", 필드 불필요).
# Firebase Console > Authentication > 로그인 방법 > Google 을 사용 설정하세요.
#
# 인증 기능을 끄려면: FIREBASE_ENABLED = False
FIREBASE_ENABLED = True  # True로 바꾸면 로그인 필요
FIREBASE_DATA    = True  # True: 학생 데이터를 Firestore에서 로드 (보안 강화; False: HTML에 삽입)
FIREBASE_CONFIG = {
    "apiKey":            "AIzaSyByw30PVfcnckhaP9F8cyAU5RkefQpheaM",
    "authDomain":        "hansung-admission.firebaseapp.com",
    "projectId":         "hansung-admission",
    "storageBucket":     "hansung-admission.firebasestorage.app",
    "messagingSenderId": "561483748784",
    "appId":             "1:561483748784:web:6e29c9d2343d26ed197275",
}
# ─────────────────────────────────────────────────────────────────────────────

MEDICAL_KEYWORDS = ['의예', '의학', '의과', '치의', '치과', '한의', '약학', '수의예']
# 의료 키워드('의과' 등)에 걸리지만 실제로는 메디컬 계열이 아닌 학과 → 일반계열로 편입
NON_MEDICAL_DEPTS = ['바이오시스템의과학부']  # 고려대 생명과학대학 (의과대학 아님)
SPECIAL_KEYWORDS = [
    '기초생활', '기회균형', '기회균등', '고른기회', '사회통합', '차상위',
    '저소득', '이웃사랑', '국가보훈', '농어촌', '다문화', '특수교육',
    '사회적배려', '사회기여', '연세한마음', '교육기회균형', '사회공헌',
]


def classify_admission(전형명칭, 전형유형):
    n = str(전형명칭)
    t = str(전형유형)
    if any(kw in n for kw in SPECIAL_KEYWORDS):
        return '특별전형'
    if '실기' in t or '실적' in t or '실기' in n or '실적' in n:
        return '특기자전형'
    if t == '논술' or '논술' in n:
        return '논술'
    if t == '종합' or '종합' in n:
        return '학생부종합'
    return '기타'


def prepare_data(file_path):
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"Excel 파일 로드 오류: {e}")
        raise

    def find_col(df_cols, keywords, exclude=None):
        exclude = exclude or set()
        available = [c for c in df_cols if c not in exclude]
        for kw in keywords:
            for col in available:
                if str(col) == kw:
                    return col
        for kw in keywords:
            for col in available:
                if kw in str(col):
                    return col
        return None

    candidates = {
        '대입연도':      ['대입연도', '연도', '입시연도'],
        '1차결과':       ['1차결과', '1차', '1차 합격', '1차_결과'],
        '결과':          ['최종결과', '최종 합격', '최종결과', '결과'],
        '대학교명':      ['대학교명', '대학', '학교'],
        '국영수과 등평': ['국영수과 등평', '국영수과등평'],
        '전교과 등평':   ['전교과 등평', '전교과등평'],
        '수학과학 등평': ['수학과학 등평', '수학과학등평'],
        '수학 등평':     ['수학 등평', '수학등평'],
        '과학 등평':     ['과학\n 등평', '과학 등평', '과학등평'],
        '학년':          ['학년'],
        '모집단위':      ['모집 단위(학과)', '모집단위', '학과', '모집'],
        '전형명칭':      ['전형명칭', '전형명'],
        '전형유형':      ['전형\n유형', '전형유형', '유형'],
        '최초합격':      ['최초합격'],
        '수능최저충족':  ['수능 최저 충족', '수능최저충족', '수능 최저충족'],
        '예비번호':      ['예비번호\n(있는경우)', '예비번호', '예비 번호'],
        '추가합격상세':  ['추가합격'],
        '등록여부':      ['등록'],
        '학번':          ['학번'],
        '이름':          ['이름', '성명'],
    }

    col_map = {}
    used_cols = set()
    for key, kwlist in candidates.items():
        found = find_col(df.columns, kwlist, exclude=used_cols)
        if found:
            col_map[key] = found
            used_cols.add(found)

    missing = [c for c in ['대입연도', '1차결과', '결과', '대학교명', '국영수과 등평'] if c not in col_map]
    if missing:
        print("필수 컬럼 매핑 실패:", missing)
        raise KeyError(f"Missing required columns: {missing}")

    df = df.rename(columns={v: k for k, v in col_map.items()})
    df['국영수과 등평'] = pd.to_numeric(df['국영수과 등평'], errors='coerce')
    for g in ('전교과 등평', '수학과학 등평', '수학 등평', '과학 등평'):
        if g in df.columns:
            df[g] = pd.to_numeric(df[g], errors='coerce')
        else:
            df[g] = float('nan')
    df = df.dropna(subset=['국영수과 등평', '대학교명', '대입연도'])
    df['대입연도'] = df['대입연도'].astype(int)

    def categorize_status(row):
        first = str(row['1차결과']).strip()
        final = str(row['결과']).strip()
        # 'Early 합격'은 반도체시스템인재전형Ⅰ 등 Early 방식 최종합격
        if final == '합격' or final == 'Early 합격':
            return '최종합격'
        if first == '합격':
            return '1차합격_최종탈락'
        return '1차탈락'

    df['상태구분'] = df.apply(categorize_status, axis=1)

    medical_pat = '|'.join(MEDICAL_KEYWORDS)
    if '모집단위' in df.columns:
        df['is_medical'] = df['모집단위'].str.contains(medical_pat, na=False)
        # 의료 키워드에 걸리지만 실제 메디컬이 아닌 학과는 제외
        nonmed_pat = '|'.join(NON_MEDICAL_DEPTS)
        df.loc[df['모집단위'].str.contains(nonmed_pat, na=False), 'is_medical'] = False
    else:
        df['is_medical'] = False

    df['admission_type'] = df.apply(
        lambda r: classify_admission(r.get('전형명칭', ''), r.get('전형유형', '')), axis=1
    )

    # KAIST 창의도전전형 → 별도 대학으로 분리
    if '전형명칭' in df.columns:
        creative_mask = (
            df['대학교명'].str.contains('KAIST', na=False) &
            df['전형명칭'].astype(str).str.contains('창의도전', na=False)
        )
        df.loc[creative_mask, '대학교명'] = 'KAIST(창의도전)'
        # 반도체시스템인재전형Ⅰ (2026만, Ⅱ 제외) → KAIST(창의도전)으로 편입
        # U+2160=Ⅰ, U+2161=Ⅱ; Ⅱ는 KAIST 일반전형으로 유지
        semi_mask = (
            (df['대학교명'] == 'KAIST') &
            (df['대입연도'] == 2026) &
            df['전형명칭'].astype(str).str.contains('반도체시스템인재', na=False) &
            ~df['전형명칭'].astype(str).str.contains('Ⅱ', na=False)
        )
        df.loc[semi_mask, '대학교명'] = 'KAIST(창의도전)'
    # KAIST 일반전형에서 Early 합격 허수 행 제거
    # (창의도전전형 합격자가 일반전형 평가를 받지 않은 케이스)
    if '최초합격' in df.columns:
        early_dummy = (
            (df['대학교명'] == 'KAIST') &
            df['최초합격'].astype(str).str.strip().str.contains('Early', case=False, na=False) &
            (df['상태구분'] != '최종합격')
        )
        df = df[~early_dummy].reset_index(drop=True)

    # 전형명칭이 없는 경우 빈 값으로 보완 (전형별 드릴다운용)
    if '전형명칭' not in df.columns:
        df['전형명칭'] = ''

    # 추가합격: 최종합격이지만 최초합격이 아닌 경우
    # 미응시: 1차합격_최종탈락 중 최초합격 열이 '미응시'인 케이스 (유의미성 낮음)
    if '최초합격' in df.columns:
        # 'Early 합격' 최종결과는 직접합격 → 추가합격 아님
        is_early_final = df['결과'].astype(str).str.strip() == 'Early 합격'
        df['is_additional'] = (
            (df['상태구분'] == '최종합격') &
            (df['최초합격'].astype(str).str.strip() != '합격') &
            (~is_early_final)
        )
        df['is_nonattend'] = (df['상태구분'] == '1차합격_최종탈락') & \
                             (df['최초합격'].astype(str).str.strip() == '미응시')
    else:
        df['is_additional'] = False
        df['is_nonattend'] = False

    # 수능 최저 충족 여부 (미충족 = True)
    if '수능최저충족' in df.columns:
        df['is_suneung_fail'] = df['수능최저충족'].astype(str).str.strip() == '미충족'
    else:
        df['is_suneung_fail'] = False

    # 예비번호 (없으면 빈 문자열)
    def fmt_wait_num(x):
        try:
            v = str(x).strip()
            if v.lower() in ('', 'nan', 'none'):
                return ''
            return str(int(float(v)))
        except Exception:
            v = str(x).strip()
            return '' if v.lower() in ('nan', 'none') else v
    if '예비번호' in df.columns:
        df['wait_num'] = df['예비번호'].apply(fmt_wait_num)
    else:
        df['wait_num'] = ''

    def fmt_str_col(x):
        v = str(x).strip()
        return '' if v.lower() in ('nan', 'none', '') else v

    if '추가합격상세' in df.columns:
        df['add_info'] = df['추가합격상세'].apply(fmt_str_col)
    else:
        df['add_info'] = ''

    if '등록여부' in df.columns:
        df['enroll_info'] = df['등록여부'].apply(fmt_str_col)
    else:
        df['enroll_info'] = ''

    for col in ('학번', '이름'):
        if col not in df.columns:
            df[col] = ''
        else:
            df[col] = df[col].apply(lambda x: '' if str(x).strip().lower() in ('nan', 'none') else str(x).strip())

    cols = ['대입연도', '학년', '대학교명', '국영수과 등평', '전교과 등평', '수학과학 등평', '수학 등평', '과학 등평', '모집단위', '전형명칭', '상태구분', 'is_additional', 'is_nonattend', 'is_medical', 'admission_type', 'is_suneung_fail', 'wait_num', 'add_info', 'enroll_info', '학번', '이름']
    records = df[cols].to_dict('records')
    years = sorted(df['대입연도'].unique().tolist())
    return records, years


HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no">
<meta name="theme-color" content="#1a2a4a">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="입시 분석">
<link rel="manifest" href="manifest.json">
<title>한성과학고등학교 대학교별 내신 등급 입시 결과 분석</title>
<!-- Firebase SDK (모듈 방식) -->
<script type="module" id="firebase-init-module">
const FIREBASE_ENABLED = __FIREBASE_ENABLED__;
const FIREBASE_CONFIG  = __FIREBASE_CONFIG__;
const FIREBASE_DATA    = __FIREBASE_DATA__;

if (FIREBASE_ENABLED) {
  const { initializeApp }                      = await import('https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js');
  const { getAuth, GoogleAuthProvider, signInWithPopup, signOut, onAuthStateChanged }
                                               = await import('https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js');
  const { getFirestore, doc, getDoc }          = await import('https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js');

  const app  = initializeApp(FIREBASE_CONFIG);
  const auth = getAuth(app);
  const db   = getFirestore(app);

  const overlay    = document.getElementById('auth-overlay');
  const loginBtn   = document.getElementById('auth-login-btn');
  const authMsg    = document.getElementById('auth-msg');
  const mainPage   = document.querySelector('.page');
  const authStatus    = document.getElementById('auth-status');
  const authEmail     = document.getElementById('auth-email');
  const mainLogoutBtn = document.getElementById('main-logout-btn');
  mainLogoutBtn.addEventListener('click', () => signOut(auth));

  async function checkWhitelist(email) {
    try {
      const ref  = doc(db, 'Whitelist', email.toLowerCase());
      const snap = await getDoc(ref);
      return snap.exists();
    } catch (e) {
      // permission-denied → 화이트리스트에 없는 것으로 처리
      return false;
    }
  }

  // 로그인 버튼: 팝업만 띄우고, 권한 체크는 onAuthStateChanged가 담당
  async function handleLogin() {
    authMsg.textContent = '';
    try {
      const provider = new GoogleAuthProvider();
      provider.setCustomParameters({ prompt: 'select_account' });
      await signInWithPopup(auth, provider);
      // 이후 처리는 onAuthStateChanged에서
    } catch (e) {
      if (e.code !== 'auth/popup-closed-by-user') {
        authMsg.textContent = '로그인 실패: ' + e.message;
      }
    }
  }

  onAuthStateChanged(auth, async user => {
    if (!user) {
      overlay.style.display     = 'flex';
      mainPage.style.display    = 'none';
      authStatus.style.display  = 'none';
      authEmail.textContent     = '';
      return;
    }
    // 토큰이 Firestore에 전파될 때까지 대기
    await user.getIdToken(true);
    const allowed = await checkWhitelist(user.email);
    if (!allowed) {
      overlay.style.display     = 'flex';
      mainPage.style.display    = 'none';
      authStatus.style.display  = 'none';
      authMsg.textContent       = `❌ 접근 권한이 없는 계정입니다.\n로그인 시도: ${user.email}\n\n관리자에게 해당 이메일 등록을 요청하세요.`;
      await signOut(auth);
      return;
    }
    overlay.style.display    = 'none';
    authStatus.style.display = 'flex';
    authEmail.textContent    = user.email;
    if (FIREBASE_DATA) {
      const dlEl = document.getElementById('data-loading');
      dlEl.style.display = 'flex';
      try {
        const metaSnap = await getDoc(doc(db, 'admissionData', 'meta'));
        const fetchedYears = metaSnap.data().years;
        const yearSnaps = await Promise.all(
          fetchedYears.map(y => getDoc(doc(db, 'admissionData', String(y))))
        );
        const allRecords = yearSnaps.flatMap(s => s.data()?.records ?? []);
        dlEl.style.display  = 'none';
        mainPage.style.display = '';
        window.__initWithData(allRecords, fetchedYears);
      } catch(e) {
        document.getElementById('data-loading').style.display = 'none';
        authMsg.textContent = '데이터 로드 실패: ' + e.message;
        overlay.style.display = 'flex';
        await signOut(auth);
      }
    } else {
      mainPage.style.display = '';
    }
  });

  loginBtn.addEventListener('click', handleLogin);
  logoutBtn.addEventListener('click', () => signOut(auth));

  window.__authLogout = () => signOut(auth);
} else {
  // Firebase 비활성 시 로그인 화면 숨김
  const overlay  = document.getElementById('auth-overlay');
  if (overlay) overlay.style.display = 'none';
}
</script>
<script src="https://cdn.plot.ly/plotly-2.32.0.min.js"></script>
<style>
  * { box-sizing: border-box; }
  body { font-family: 'Malgun Gothic', 'Apple SD Gothic Neo', sans-serif; margin: 0; background: #f0f2f5; color: #222; }
  .page { max-width: 1400px; margin: 0 auto; padding: 24px 20px; }

  /* ── Firestore 데이터 로딩 오버레이 ── */
  @keyframes spin { to { transform: rotate(360deg); } }
  #data-loading {
    display: none; position: fixed; inset: 0; z-index: 9998;
    background: rgba(240,242,245,0.97);
    align-items: center; justify-content: center; flex-direction: column; gap: 16px;
  }
  #data-loading .dl-spinner {
    width: 44px; height: 44px; border: 4px solid #d0d9e8;
    border-top-color: #1a2a4a; border-radius: 50%;
    animation: spin 0.8s linear infinite;
  }
  #data-loading p { color: #1a2a4a; font-size: 15px; font-weight: 600; margin: 0; }

  /* ── Firebase 로그인 오버레이 ── */
  #auth-overlay {
    display: none; position: fixed; inset: 0; z-index: 9999;
    background: linear-gradient(135deg, #1a2a4a 0%, #0c4da2 100%);
    align-items: center; justify-content: center;
  }
  #auth-box {
    background: white; border-radius: 16px; padding: 40px 36px;
    width: min(400px, 90vw); text-align: center;
    box-shadow: 0 12px 40px rgba(0,0,0,0.3);
  }
  #auth-box h1 { font-size: 17px; color: #1a2a4a; margin: 0 0 6px; font-weight: 700; }
  #auth-box p  { font-size: 13px; color: #888; margin: 0 0 28px; }
  #auth-login-btn {
    display: flex; align-items: center; justify-content: center; gap: 10px;
    width: 100%; padding: 12px 20px; border: 1.5px solid #d0d7e3; border-radius: 8px;
    background: white; cursor: pointer; font-size: 14px; font-weight: 600; color: #333;
    transition: background .15s, box-shadow .15s;
  }
  #auth-login-btn:hover { background: #f5f5f5; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
  #auth-login-btn svg { width: 20px; height: 20px; flex-shrink: 0; }
  #auth-msg {
    margin-top: 16px; font-size: 12px; color: #e74c3c;
    white-space: pre-line; min-height: 18px;
  }
  #auth-footer {
    margin-top: 20px; font-size: 11px; color: #ccc;
    display: flex; align-items: center; justify-content: space-between;
  }
  #auth-user { font-size: 12px; color: #555; }
  #auth-logout-btn {
    display: none; font-size: 12px; color: #888; border: none; background: none;
    cursor: pointer; text-decoration: underline; padding: 0;
  }
  h2 { margin: 0 0 20px; font-size: 20px; color: #1a2a4a; }

  .controls { background: white; border-radius: 10px; padding: 16px 20px; margin-bottom: 12px;
              box-shadow: 0 1px 4px rgba(0,0,0,0.08); display: flex; flex-wrap: wrap; gap: 20px; align-items: flex-start; }
  .ctrl-group { display: flex; flex-direction: column; gap: 8px; }
  .ctrl-label { font-size: 12px; font-weight: 700; color: #666; text-transform: uppercase; letter-spacing: .5px; }
  .btn-row { display: flex; gap: 6px; flex-wrap: wrap; }

  .year-btn, .filter-btn, .cat-btn, .view-btn {
    padding: 6px 14px; border: 2px solid #d0d7e3; border-radius: 20px;
    background: white; cursor: pointer; font-size: 13px; font-weight: 600;
    color: #555; transition: all .15s;
  }
  .year-btn.on   { background: #0c4da2; border-color: #0c4da2; color: white; }
  .filter-btn.on { background: #2d7dd2; border-color: #2d7dd2; color: white; }
  .cat-btn.on    { background: #c0392b; border-color: #c0392b; color: white; }
  .view-btn.on   { background: #16a085; border-color: #16a085; color: white; }

  .grade-input {
    width: 100px; padding: 6px 12px; border: 2px solid #d0d7e3; border-radius: 20px;
    font-size: 13px; font-weight: 600; color: #333; outline: none; transition: border .15s;
  }
  .grade-input:focus { border-color: #e74c3c; }

  .chart-wrap { background: white; border-radius: 10px; padding: 16px;
                box-shadow: 0 1px 4px rgba(0,0,0,0.08); }
  .chart-scroll-wrap { width: 100%; overflow: hidden; }
  #chart, #landing-chart { width: 100%; }
  .empty-msg { text-align: center; padding: 60px; color: #aaa; font-size: 15px; }

  /* ── 상단 탭 (맞춤 조회 / 전체 분석) ── */
  .main-tabs { display: flex; gap: 8px; margin-bottom: 12px; }
  .main-tab {
    padding: 9px 22px; border: 2px solid #d0d7e3; border-radius: 10px 10px 0 0;
    background: #eef2f8; cursor: pointer; font-size: 14px; font-weight: 700; color: #667;
    transition: all .15s;
  }
  .main-tab.on { background: #1a2a4a; border-color: #1a2a4a; color: white; }

  /* ── 랜딩 모드 세그먼트 토글 ── */
  .landing-modes { display: flex; gap: 6px; margin-bottom: 12px; }
  .seg {
    padding: 7px 18px; border: 2px solid #d0d7e3; border-radius: 20px;
    background: white; cursor: pointer; font-size: 13px; font-weight: 600; color: #555;
    transition: all .15s;
  }
  .seg.on { background: #0c4da2; border-color: #0c4da2; color: white; }

  .landing-ctrl {
    background: white; border-radius: 10px; padding: 14px 18px; margin-bottom: 12px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.08);
    display: flex; align-items: center; gap: 12px; flex-wrap: wrap;
  }
  .landing-ctrl label { font-size: 13px; font-weight: 700; color: #666; }
  #univ-select {
    padding: 8px 14px; border: 2px solid #d0d7e3; border-radius: 20px;
    font-size: 14px; font-weight: 600; color: #333; outline: none; min-width: 220px;
    background: white; cursor: pointer;
  }
  #univ-select:focus { border-color: #0c4da2; }
  .landing-hint { font-size: 13px; color: #888; }
  .landing-hint b { color: #0c4da2; }

  /* ── 소수 선발 전형 경고 배너 ── */
  #special-notice {
    background: #fef5e7; border: 1px solid #f0c36d; border-left: 4px solid #e67e22;
    border-radius: 8px; padding: 10px 14px; margin-bottom: 12px;
    font-size: 12.5px; color: #7a4a10; line-height: 1.6;
  }
  #special-notice b { color: #b9520a; }

  /* 드릴다운 모달 */
  #modal-overlay {
    display: none; position: fixed; inset: 0;
    background: rgba(0,0,0,0.45); z-index: 1000;
    align-items: center; justify-content: center;
  }
  #modal-overlay.open { display: flex; }
  #modal-box {
    background: white; border-radius: 14px; padding: 24px;
    width: min(820px, 95vw); box-shadow: 0 8px 32px rgba(0,0,0,0.2);
    position: relative;
  }
  #modal-title { font-size: 17px; font-weight: 700; color: #1a2a4a; margin-bottom: 4px; }
  #modal-sub   { font-size: 12px; color: #aaa; margin-bottom: 10px; }
  #modal-close {
    position: absolute; top: 16px; right: 18px;
    background: none; border: none; font-size: 20px; cursor: pointer; color: #888;
  }
  .modal-tabs { display: flex; gap: 6px; margin-bottom: 14px; }
  .modal-tab {
    padding: 5px 16px; border: 2px solid #d0d7e3; border-radius: 20px;
    background: white; cursor: pointer; font-size: 13px; font-weight: 600; color: #555;
  }
  .modal-tab.on { background: #16a085; border-color: #16a085; color: white; }
  .modal-add-toggle {
    display: flex; align-items: center; gap: 6px; margin-left: auto;
    font-size: 12px; color: #555; cursor: pointer; user-select: none; white-space: nowrap;
  }
  .modal-add-toggle input { display: none; }
  .modal-add-slider {
    display: inline-block; width: 32px; height: 18px; border-radius: 9px;
    background: #ccc; position: relative; transition: background .2s; flex-shrink: 0;
  }
  .modal-add-slider::after {
    content: ''; position: absolute; left: 2px; top: 2px;
    width: 14px; height: 14px; border-radius: 50%; background: white; transition: left .2s;
  }
  .modal-add-toggle input:checked + .modal-add-slider { background: #e67e22; }
  .modal-add-toggle input:checked + .modal-add-slider::after { left: 16px; }
  #modal-chart { width: 100%; }
  #modal-student { font-family: 'Segoe UI', sans-serif; }
  .stu-header { display:flex; gap:16px; align-items:center; flex-wrap:wrap;
    padding:10px 14px; background:#f0f4fa; border-radius:8px; margin-bottom:12px; }
  .stu-header b { font-size:15px; color:#1A2A4A; }
  .stu-header span { font-size:13px; color:#555; }
  .stu-table { width:100%; border-collapse:collapse; font-size:13px; }
  .stu-table th { text-align:left; padding:7px 10px; color:#1A2A4A;
    border-bottom:2px solid #d0dae8; font-size:12px; white-space:nowrap; }
  .stu-table td { padding:7px 10px; border-bottom:1px solid #eef2f7; vertical-align:middle; }
  .stu-table tr:last-child td { border-bottom:none; }
  .stu-table tr:nth-child(even) td { background:#f8fafc; }
  .stu-badge { display:inline-block; padding:2px 9px; border-radius:12px;
    font-size:11px; font-weight:700; color:white; white-space:nowrap; }

  #grade-summary { margin-top: 14px; display: flex; flex-wrap: wrap; gap: 12px; }
  .summary-group { background: white; border-radius: 10px; padding: 12px 16px;
                   box-shadow: 0 1px 4px rgba(0,0,0,0.08); flex: 1; min-width: 200px; }
  .summary-title { font-size: 12px; font-weight: 700; margin-bottom: 8px; }
  .summary-chips { display: flex; flex-wrap: wrap; gap: 5px; }
  .chip { padding: 3px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; }
  .chip-stable { background: #d5f5e3; color: #1e8449; }
  .chip-proper { background: #fcf3cf; color: #9a7d0a; }
  .chip-reach  { background: #fdebd0; color: #784212; }
  .hint { font-size: 11px; color: #aaa; margin-top: 4px; }

  /* 토글 스위치 */
  .toggle-group { display: flex; flex-direction: column; gap: 6px; }
  .toggle-item { display: flex; align-items: center; gap: 8px; }
  .toggle-label { font-size: 13px; font-weight: 600; color: #555; cursor: pointer; user-select: none; }
  .toggle-switch { position: relative; display: inline-block; width: 38px; height: 20px; flex-shrink: 0; }
  .toggle-switch input { opacity: 0; width: 0; height: 0; }
  .toggle-slider {
    position: absolute; cursor: pointer; inset: 0;
    background: #ccc; border-radius: 20px; transition: .2s;
  }
  .toggle-slider:before {
    content: ''; position: absolute;
    width: 14px; height: 14px; left: 3px; bottom: 3px;
    background: white; border-radius: 50%; transition: .2s;
  }
  .toggle-switch input:checked + .toggle-slider { background: #e67e22; }
  .toggle-switch input:checked + .toggle-slider:before { transform: translateX(18px); }
  .toggle-switch.blue input:checked + .toggle-slider { background: #2d7dd2; }

  /* 모바일 반응형 */
  @media (max-width: 680px) {
    body { background: #eef0f4; }
    .page { padding: 0 0 24px; }
    h2 { font-size: 15px; margin: 0; padding: 14px 14px 10px; background: #1a2a4a; color: white; }

    /* 컨트롤 패널: 카드들을 세로로 쌓음 */
    .controls {
      flex-direction: column; gap: 0; padding: 0;
      border-radius: 0; box-shadow: none; background: transparent; margin-bottom: 8px;
    }
    /* 각 필터 그룹을 가로 스크롤 카드로 */
    .ctrl-group {
      background: white; margin: 4px 8px 0;
      border-radius: 10px; padding: 10px 12px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.08);
      flex-direction: column; gap: 6px;
    }
    .ctrl-group[style*="margin-left:auto"] { display: none; }
    .ctrl-label { font-size: 11px; }
    /* 버튼 행은 가로 스크롤 */
    .btn-row {
      display: flex; flex-wrap: nowrap; overflow-x: auto;
      gap: 6px; padding-bottom: 2px;
      scrollbar-width: none; -ms-overflow-style: none;
    }
    .btn-row::-webkit-scrollbar { display: none; }
    .year-btn, .filter-btn, .cat-btn, .view-btn {
      padding: 7px 14px; font-size: 13px; white-space: nowrap; flex-shrink: 0;
      min-height: 36px;
    }
    .grade-input { width: 100%; max-width: 160px; min-height: 36px; font-size: 14px; }

    /* 차트 영역: 가로 스크롤 */
    .chart-wrap { border-radius: 0; margin: 0 4px; padding: 8px 4px; }
    .chart-scroll-wrap {
      overflow-x: auto; overflow-y: hidden;
      -webkit-overflow-scrolling: touch;
    }
    #chart, #landing-chart { min-width: 800px; width: max-content; }

    /* 상단 탭 / 세그먼트 / 랜딩 컨트롤 모바일 */
    .main-tabs { padding: 10px 8px 0; margin-bottom: 0; }
    .main-tab { padding: 9px 16px; font-size: 13px; min-height: 40px; }
    .landing-modes { padding: 8px 8px 0; margin-bottom: 4px; }
    .seg { min-height: 36px; }
    .landing-ctrl { margin: 4px 8px 8px; border-radius: 10px; }
    #univ-select { width: 100%; min-width: 0; min-height: 40px; }
    #special-notice { margin: 4px 8px 8px; }

    /* 토글 그룹 */
    .toggle-group { flex-direction: row; flex-wrap: wrap; gap: 16px; }
    .toggle-item { gap: 8px; }
    .toggle-label { font-size: 13px; }

    /* 모달: 바텀 시트 */
    #modal-overlay { align-items: flex-end; }
    #modal-box {
      width: 100%; max-width: 100%;
      border-radius: 18px 18px 0 0; padding: 20px 16px 32px;
      max-height: 88vh; overflow-y: auto;
      animation: slideUp .25s cubic-bezier(0.32,0.72,0,1);
    }
    @keyframes slideUp {
      from { transform: translateY(100%); }
      to   { transform: translateY(0); }
    }
    /* 바텀시트 드래그 핸들 */
    #modal-box::before {
      content: ''; display: block; width: 36px; height: 4px;
      background: #d0d7e3; border-radius: 2px; margin: 0 auto 16px;
    }
    #modal-title { font-size: 16px; }
    #modal-sub { font-size: 12px; }
    #modal-chart { min-height: 280px; }
    .modal-tab { padding: 6px 14px; font-size: 13px; min-height: 34px; }

    /* 합격 요약 */
    #grade-summary { gap: 8px; padding: 0 8px; }
    .summary-group { min-width: 100%; padding: 12px 14px; }
    .chip { font-size: 12px; padding: 4px 10px; }
  }
</style>
</head>
<body>

<!-- Firestore 데이터 로딩 오버레이 -->
<div id="data-loading">
  <div class="dl-spinner"></div>
  <p>데이터 불러오는 중...</p>
</div>

<!-- Firebase 로그인 오버레이 -->
<div id="auth-overlay">
  <div id="auth-box">
    <h1>한성과학고등학교 입시 분석</h1>
    <p>접근 권한이 있는 계정으로 로그인하세요.</p>
    <button id="auth-login-btn">
      <!-- Google 로고 SVG -->
      <svg viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg">
        <path fill="#EA4335" d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0 14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z"/>
        <path fill="#4285F4" d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z"/>
        <path fill="#FBBC05" d="M10.53 28.59c-.48-1.45-.76-2.99-.76-4.59s.27-3.14.76-4.59l-7.98-6.19C.92 16.46 0 20.12 0 24c0 3.88.92 7.54 2.56 10.78l7.97-6.19z"/>
        <path fill="#34A853" d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6c-2.18 1.48-4.97 2.35-8.16 2.35-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19C6.51 42.62 14.62 48 24 48z"/>
      </svg>
      Google 계정으로 로그인
    </button>
    <div id="auth-msg"></div>
    <div id="auth-footer">
      <span id="auth-user"></span>
      <button id="auth-logout-btn" onclick="window.__authLogout && window.__authLogout()">로그아웃</button>
    </div>
  </div>
</div>

<div class="page">
  <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:20px; flex-wrap:wrap; gap:8px;">
    <h2 style="margin:0;">한성과학고등학교 대학교별 내신 등급 입시 결과 분석</h2>
    <div id="auth-status" style="display:none; align-items:center; gap:10px; font-size:12px; color:#666;">
      <span id="auth-email"></span>
      <button id="main-logout-btn"
        style="font-size:12px; color:#888; border:1px solid #d0d7e3; border-radius:12px;
               background:white; padding:4px 12px; cursor:pointer;">로그아웃</button>
    </div>
  </div>

  <!-- 공유 필터 (두 탭 공통) -->
  <div class="controls">
    <div class="ctrl-group">
      <div class="ctrl-label">학년</div>
      <div class="btn-row" id="grade-year-row"></div>
    </div>
    <div class="ctrl-group">
      <div class="ctrl-label">입시 연도</div>
      <div class="btn-row" id="year-row"></div>
    </div>
    <div class="ctrl-group">
      <div class="ctrl-label">전형 유형</div>
      <div class="btn-row" id="filter-row"></div>
    </div>
    <div class="ctrl-group">
      <div class="ctrl-label">계열</div>
      <div class="btn-row" id="cat-row"></div>
    </div>
    <div class="ctrl-group">
      <div class="ctrl-label">등급 기준</div>
      <div class="btn-row" id="grade-row"></div>
    </div>
    <div class="ctrl-group" id="mygrade-group">
      <div class="ctrl-label">내 등급</div>
      <input type="number" id="my-grade" class="grade-input" min="1" max="9" step="0.01" placeholder="예: 2.30">
    </div>
  </div>

  <!-- 소수 선발 전형 경고 (특별전형/특기자전형 선택 시 표시) -->
  <div id="special-notice" style="display:none;"></div>

  <!-- 상단 탭 -->
  <div class="main-tabs">
    <button class="main-tab on" id="mtab-landing" onclick="switchMainTab('landing')">🎯 맞춤 조회</button>
    <button class="main-tab"    id="mtab-full"    onclick="switchMainTab('full')">📊 전체 분석</button>
  </div>

  <!-- Tab 1: 맞춤 조회 (랜딩) -->
  <div id="view-landing">
    <div class="landing-modes">
      <button class="seg on" id="seg-univ"  onclick="switchLandingMode('univ')">대학 선택</button>
      <button class="seg"    id="seg-grade" onclick="switchLandingMode('grade')">내 등급으로 찾기</button>
    </div>
    <div class="landing-ctrl" id="landing-univ-ctrl">
      <label for="univ-select">대학 선택</label>
      <select id="univ-select"><option value="">— 대학을 선택하세요 —</option></select>
      <span class="landing-hint">선택한 대학의 분포만 표시됩니다. 상단 <b>내 등급</b> 입력 시 기준선도 함께 표시.</span>
    </div>
    <div class="landing-ctrl" id="landing-grade-ctrl" style="display:none;">
      <span class="landing-hint">상단 <b>내 등급</b>을 입력하면 지원 가능 범위(위수염 이내)의 대학이 표시됩니다.</span>
    </div>
    <div class="chart-wrap">
      <div class="chart-scroll-wrap">
        <div id="landing-chart"></div>
      </div>
    </div>
    <div id="landing-summary"></div>
  </div>

  <!-- Tab 2: 전체 분석 -->
  <div id="view-full" style="display:none;">
    <div class="controls">
      <div class="ctrl-group">
        <div class="ctrl-label">표시 옵션</div>
        <div class="toggle-group">
          <div class="toggle-item">
            <label class="toggle-switch blue">
              <input type="checkbox" id="toggle-nopass">
              <span class="toggle-slider"></span>
            </label>
            <span class="toggle-label" onclick="document.getElementById('toggle-nopass').click()">합격 없는 대학 표시</span>
          </div>
          <div class="toggle-item">
            <label class="toggle-switch">
              <input type="checkbox" id="toggle-additional">
              <span class="toggle-slider"></span>
            </label>
            <span class="toggle-label" onclick="document.getElementById('toggle-additional').click()">추가합격 구분 표시</span>
          </div>
        </div>
      </div>
      <div class="ctrl-group">
        <div class="ctrl-label">학과 검색</div>
        <input type="text" id="dept-filter" placeholder="예: 기계  전기" title="띄어쓰기/쉼표로 구분 시 OR 검색"
          style="height:30px; padding:0 8px; border:1.5px solid #d0d7e3; border-radius:6px; font-size:13px; width:130px; outline:none;">
      </div>
      <div class="ctrl-group" style="margin-left:auto; align-self:flex-end;">
        <div style="font-size:11px; color:#bbb; line-height:1.7; text-align:right;">
          x축: <b>대학명 (최종합격/총지원)</b><br>
          대학 클릭 → 연도별 상세 보기
        </div>
      </div>
    </div>

    <div class="chart-wrap">
      <div class="chart-scroll-wrap">
        <div id="chart"></div>
      </div>
    </div>
    <div id="grade-summary"></div>
  </div>
</div>

<!-- 드릴다운 모달 -->
<div id="modal-overlay">
  <div id="modal-box">
    <button id="modal-close" title="닫기">✕</button>
    <div id="modal-title"></div>
    <div id="modal-sub"></div>
    <div class="modal-tabs">
      <button class="modal-tab on" onclick="switchModalTab('year', this)">연도별</button>
      <button class="modal-tab"    onclick="switchModalTab('dept', this)">학과별</button>
      <button class="modal-tab"    onclick="switchModalTab('type', this)">전형별</button>
      <button class="modal-tab" id="modal-tab-student" onclick="switchModalTab('student', this)" style="display:none">학생별</button>
      <label class="modal-add-toggle" title="추가합격 구분 표시 (메인 차트와 연동)">
        <input type="checkbox" id="modal-toggle-add" onchange="syncAdditional(this.checked)">
        <span class="modal-add-slider"></span>
        추가합격 구분
      </label>
    </div>
    <div id="modal-chart"></div>
    <div id="modal-student" style="display:none;max-height:430px;overflow-y:auto"></div>
  </div>
</div>

<script>
__RAW_YEARS_DECL__

const COLORS = {
  '최종합격':        '#0c4da2',
  '1차합격_최종탈락': '#5a9fd4',
  '1차탈락':         '#cccccc',
};

function extraInfo(r) {
  let s = '';
  if (r['is_suneung_fail']) s += '<br>⚠ 수능최저 미충족';
  if (r['wait_num'])    s += '<br>예비 ' + r['wait_num'] + '번';
  if (r['add_info'] && r['add_info'] !== '추가 합격' && r['add_info'] !== '추가합격') s += '<br>추가합격: ' + r['add_info'];
  if (r['enroll_info']) s += '<br>등록: ' + r['enroll_info'];
  return s;
}

// ── state ──────────────────────────────────────────────
const state = {
  years: new Set(YEARS.length ? ([2026].filter(y => YEARS.includes(y)).concat(YEARS.includes(2026) ? [] : [YEARS[YEARS.length-1]])) : []),
  gradeYears: new Set([2, 3]),
  filter: '학생부종합',
  cat: '일반계열',
  grade: '국영수과 등평',
  myGrade: null,
  showNoPass: false,
  showAdditional: false,
  deptFilter: '',
};

// ── 맞춤 조회(랜딩) 상태 ──
let mainTab     = 'landing';  // 'landing' | 'full'
let landingMode = 'univ';     // 'univ' | 'grade'

const ADD_COLOR    = '#e67e22'; // 추가합격 (주황)
const FIRST_COLOR  = '#0c4da2'; // 최초합격 (기존 최종합격 색)

// 티어 색상 (요약 칩 클래스와 동일 계열)
const TIER = {
  stable:  { emoji: '🟢', label: '안정권', cls: 'chip-stable' },
  proper:  { emoji: '🟡', label: '적정권', cls: 'chip-proper' },
  reach:   { emoji: '🟠', label: '도전권', cls: 'chip-reach'  },
};

// ── UI setup ───────────────────────────────────────────
function makeBtn(text, cls, active, onClick) {
  const b = document.createElement('button');
  b.textContent = text;
  b.className = cls + (active ? ' on' : '');
  b.addEventListener('click', onClick);
  return b;
}

// 학년 버튼 (multi-select)
const gradeYearRow = document.getElementById('grade-year-row');
[2, 3].forEach(g => {
  const b = makeBtn(g + '학년', 'year-btn', true, () => {
    if (state.gradeYears.has(g) && state.gradeYears.size === 1) return;
    state.gradeYears.has(g) ? state.gradeYears.delete(g) : state.gradeYears.add(g);
    b.classList.toggle('on', state.gradeYears.has(g));
    render();
  });
  gradeYearRow.appendChild(b);
});

// Year buttons (multi-select)
const yearRow = document.getElementById('year-row');
function buildYearButtons() {
  yearRow.innerHTML = '';
  YEARS.forEach(y => {
    const b = makeBtn(y + '년', 'year-btn', state.years.has(y), () => {
      if (state.years.has(y) && state.years.size === 1) return;
      state.years.has(y) ? state.years.delete(y) : state.years.add(y);
      b.classList.toggle('on', state.years.has(y));
      render();
    });
    yearRow.appendChild(b);
  });
}
buildYearButtons();

// Filter buttons
const FILTERS = ['전체', '학생부종합', '논술', '특기자전형', '특별전형'];
const filterRow = document.getElementById('filter-row');
FILTERS.forEach(f => {
  const b = makeBtn(f, 'filter-btn', f === '학생부종합', () => {
    state.filter = f;
    filterRow.querySelectorAll('.filter-btn').forEach(x => x.classList.remove('on'));
    b.classList.add('on');
    render();
  });
  filterRow.appendChild(b);
});

// Category buttons
const CATS = ['일반계열', '메디컬계열'];
const catRow = document.getElementById('cat-row');
CATS.forEach(c => {
  const b = makeBtn(c, 'cat-btn', c === state.cat, () => {
    state.cat = c;
    catRow.querySelectorAll('.cat-btn').forEach(x => x.classList.remove('on'));
    b.classList.add('on');
    render();
  });
  catRow.appendChild(b);
});

// Grade type buttons
const GRADES = ['국영수과 등평', '전교과 등평'];
const gradeRow = document.getElementById('grade-row');
GRADES.forEach(g => {
  const b = makeBtn(g, 'filter-btn', g === state.grade, () => {
    state.grade = g;
    gradeRow.querySelectorAll('.filter-btn').forEach(x => x.classList.remove('on'));
    b.classList.add('on');
    render();
  });
  gradeRow.appendChild(b);
});

// My grade input
document.getElementById('my-grade').addEventListener('input', e => {
  const v = parseFloat(e.target.value);
  state.myGrade = (!isNaN(v) && v >= 1 && v <= 9) ? v : null;
  render();
});

// 표시 옵션 토글
document.getElementById('toggle-nopass').addEventListener('change', e => {
  state.showNoPass = e.target.checked;
  render();
});
document.getElementById('toggle-additional').addEventListener('change', e => {
  state.showAdditional = e.target.checked;
  document.getElementById('modal-toggle-add').checked = e.target.checked;
  render();
});

// 학과 검색 (OR, 부분 일치)
document.getElementById('dept-filter').addEventListener('input', e => {
  state.deptFilter = e.target.value.trim();
  render();
});

// 모달 내 추가합격 토글 (메인 차트와 연동)
function syncAdditional(val) {
  state.showAdditional = val;
  document.getElementById('toggle-additional').checked = val;
  render();
  if (drillUniv) {
    if (drillTab === 'year') renderModalYear();
    else if (drillTab === 'dept') renderModalDept();
    else renderModalType();
  }
}

// 학과 필터 매칭: 띄어쓰기/쉼표 구분으로 OR 검색
function matchesDeptFilter(dept) {
  if (!state.deptFilter) return true;
  const d = (dept || '').toLowerCase();
  const keys = state.deptFilter.toLowerCase().split(/[\s,]+/).filter(k => k.length > 0);
  return keys.some(k => d.includes(k));
}

// ── helpers ────────────────────────────────────────────
function median(arr) {
  if (!arr.length) return null;
  const s = [...arr].sort((a, b) => a - b);
  const m = Math.floor(s.length / 2);
  return s.length % 2 ? s[m] : (s[m-1] + s[m]) / 2;
}

function quantile(arr, q) {
  const s = [...arr].sort((a, b) => a - b);
  const pos = (s.length - 1) * q;
  const lo = Math.floor(pos), hi = Math.ceil(pos);
  return s[lo] + (s[hi] - s[lo]) * (pos - lo);
}

// 공유 필터 술어 (연도/학년/계열/전형/학과검색). 등급 null 여부는 호출부에서 별도 판정.
function passFilters(r, isMedical) {
  return state.years.has(r['대입연도']) &&
    state.gradeYears.has(r['학년']) &&
    r['is_medical'] === isMedical &&
    (state.filter === '전체' || r['admission_type'] === state.filter) &&
    matchesDeptFilter(r['모집단위']);
}

// 박스플롯 통계 (최종합격 등급 배열 → 사분위·위수염)
function boxStats(grades) {
  if (!grades.length) return null;
  const s = [...grades].sort((a, b) => a - b);
  const q1 = quantile(s, 0.25), med = median(s), q3 = quantile(s, 0.75);
  const iqr = q3 - q1;
  const fence = q3 + 1.5 * iqr;           // Tukey 상단 울타리
  let whisker = q3;
  for (const v of s) if (v <= fence) whisker = v;  // 울타리 이내 실제 최댓값 = 위수염
  return { n: s.length, min: s[0], max: s[s.length - 1], q1, med, q3, whisker };
}

// 내 등급 g가 통계 st에서 속하는 지원 티어. 범위 밖이면 null.
// 표본이 적으면(n<5) 사분위가 불안정하므로 상한을 실제 최댓값으로 완화.
function tierOf(g, st) {
  if (g == null || !st) return null;
  const upper = st.n < 5 ? st.max : st.whisker;
  if (g <= st.q1)   return 'stable';
  if (g <= st.q3)   return 'proper';
  if (g <= upper)   return 'reach';
  return null;
}

function getOrder(byUniv, gradeKey) {
  const sortKey = {};
  for (const [u, rows] of Object.entries(byUniv)) {
    const finals = rows.filter(r => r.상태구분 === '최종합격').map(r => r[gradeKey]).filter(v => v != null);
    const firsts = rows.filter(r => r.상태구분 === '1차합격_최종탈락').map(r => r[gradeKey]).filter(v => v != null);
    if (finals.length) sortKey[u] = median(finals);
    else if (firsts.length) sortKey[u] = median(firsts);
  }
  return Object.entries(sortKey).sort((a, b) => a[1] - b[1]).map(x => x[0]);
}

// ── 분포도 ─────────────────────────────────────────────
async function renderBox(opts = {}) {
  const cfg = {
    target: 'chart',            // 렌더 대상 div id
    summaryTarget: 'grade-summary', // 요약 대상 div id (null이면 요약 없음)
    onlyUniv: null,             // 대학 모드: 이 대학만 표시
    univFilter: null,          // 등급 모드: st => bool (위수염 이내 판정)
    tierColor: false,          // 티어 표현: x축 라벨 이모지 + 요약을 도전/적정/안정으로
    tierSort: false,           // 티어 순(도전→적정→안정) 재정렬 (전체 분석은 중간값 순 유지 위해 false)
    forceEmpty: false,         // 강제 빈 화면 (미선택/미입력 안내)
    placeholder: '해당 조건의 데이터가 없습니다.',
    ...opts,
  };
  const gradeKey = state.grade;
  const isMedical = state.cat === '메디컬계열';
  const isMobile = window.innerWidth <= 680;
  const Z = _zoom[cfg.target] || (_zoom[cfg.target] = { y: [9.2, 0.8], x: null });

  const rows = RAW.filter(r => passFilters(r, isMedical) && r[gradeKey] != null);

  const byUniv = {};
  for (const r of rows) (byUniv[r['대학교명']] ??= []).push(r);
  if (state.showNoPass) {
    // 등급 데이터가 없는 대학도 포함 (null 등급 레코드 → 빈 박스로 표시)
    RAW.filter(r => passFilters(r, isMedical)).forEach(r => { byUniv[r['대학교명']] ??= []; });
  }
  // 대학 모드: 선택 대학은 합격 사례가 없어도(1차탈락만/데이터 없음) 항상 표시
  if (cfg.onlyUniv != null) byUniv[cfg.onlyUniv] ??= [];
  for (const u of Object.keys(byUniv)) {
    if (cfg.onlyUniv != null && u === cfg.onlyUniv) continue;
    if (!state.showNoPass && !byUniv[u].some(r => r['상태구분'] !== '1차탈락')) delete byUniv[u];
  }

  // 대학별 최종합격 통계 (티어/위수염 판정용) — 대학 제한 적용 전에 산출
  const statOf = {};
  for (const u of Object.keys(byUniv)) {
    statOf[u] = boxStats(byUniv[u].filter(r => r['상태구분'] === '최종합격').map(r => r[gradeKey]).filter(v => v != null));
  }

  // 대학 모드: 단일 대학으로 제한 / 등급 모드: 위수염 이내 대학만
  if (cfg.onlyUniv != null) {
    for (const u of Object.keys(byUniv)) if (u !== cfg.onlyUniv) delete byUniv[u];
  } else if (cfg.univFilter) {
    for (const u of Object.keys(byUniv)) if (!cfg.univFilter(statOf[u])) delete byUniv[u];
  }

  const orderedUnivs = getOrder(byUniv, gradeKey).filter(u => byUniv[u]);
  // 합격 없는 대학(정렬키 없음)은 뒤에 추가
  const noPassUnivs = Object.keys(byUniv).filter(u => !orderedUnivs.includes(u));
  let univs = [...orderedUnivs, ...noPassUnivs];

  // 티어 정렬 요청 시: 도전→적정→안정(그룹) + 그룹 내 중간값 오름차순
  if (cfg.tierSort && state.myGrade != null) {
    const rank = { reach: 0, proper: 1, stable: 2 };
    univs.sort((a, b) => {
      const ra = rank[tierOf(state.myGrade, statOf[a])] ?? 9;
      const rb = rank[tierOf(state.myGrade, statOf[b])] ?? 9;
      if (ra !== rb) return ra - rb;
      return (statOf[a]?.med ?? 99) - (statOf[b]?.med ?? 99);
    });
  }

  if (!univs.length || cfg.forceEmpty) {
    // Plotly.react 사용 → Plotly 컨테이너를 유지해야 다음 render가 정상 작동
    Plotly.react(cfg.target, [], {
      height: 300, plot_bgcolor: '#fafbff',
      xaxis: { visible: false }, yaxis: { visible: false },
      annotations: [{ text: cfg.placeholder,
        xref: 'paper', yref: 'paper', x: 0.5, y: 0.5,
        showarrow: false, font: { size: 16, color: '#aaa' } }],
    }, { responsive: true });
    bindChartEvents(cfg.target);
    if (cfg.summaryTarget) document.getElementById(cfg.summaryTarget).innerHTML = '';
    return;
  }

  const nFinal = {}, nTotal = {};
  for (const u of univs) {
    nFinal[u] = byUniv[u].filter(r => r['상태구분'] === '최종합격').length;
    nTotal[u] = byUniv[u].length;
  }

  const traces = [];

  // 1차탈락
  const rejX = [], rejY = [], rejD = [];
  for (const u of univs)
    for (const r of byUniv[u])
      if (r['상태구분'] === '1차탈락') { rejX.push(u); rejY.push(r[gradeKey]); rejD.push([r['모집단위'] || '', extraInfo(r), r['학번'] || '', r['대입연도']]); }
  if (rejX.length) traces.push({
    type: 'box', x: rejX, y: rejY, name: '1차탈락 (참고)',
    customdata: rejD,
    marker: { color: COLORS['1차탈락'], opacity: 0.45, size: 4 },
    line: { color: COLORS['1차탈락'] }, fillcolor: 'rgba(204,204,204,0.15)',
    boxpoints: 'all', jitter: 0.4, pointpos: 0,
    hovertemplate: '%{x}<br>등평: %{y:.2f}<br>학과: %{customdata[0]}%{customdata[1]}<extra>1차탈락</extra>',
  });

  // 1차합격_최종탈락 (미응시는 별도 scatter X로 표시 — box selectedpoints는 symbol 미지원)
  {
    const sx = [], sy = [], sd = [];
    const nax = [], nay = [], nad = [];
    for (const u of univs)
      for (const r of byUniv[u])
        if (r['상태구분'] === '1차합격_최종탈락') {
          if (r['is_nonattend']) {
            nax.push(u); nay.push(r[gradeKey]); nad.push([r['모집단위'] || '', extraInfo(r), r['학번'] || '', r['대입연도']]);
          } else {
            sx.push(u); sy.push(r[gradeKey]); sd.push([r['모집단위'] || '', extraInfo(r), r['학번'] || '', r['대입연도']]);
          }
        }
    if (sx.length) {
      traces.push({
        type: 'box', x: sx, y: sy, name: '1차합격_최종탈락',
        customdata: sd,
        marker: { color: COLORS['1차합격_최종탈락'], size: 5 },
        line: { color: COLORS['1차합격_최종탈락'] },
        boxpoints: 'all', jitter: 0.3, pointpos: 0,
        hovertemplate: '%{x}<br>등평: %{y:.2f}<br>학과: %{customdata[0]}%{customdata[1]}<extra>1차합격_최종탈락</extra>',
      });
    }
    if (nax.length) {
      traces.push({
        type: 'scatter', mode: 'markers', x: nax, y: nay,
        name: '└ 미응시(참고제외)', customdata: nad,
        marker: { color: '#aaa', size: 9, symbol: 'x-thin', opacity: 0.9,
                  line: { color: '#aaa', width: 2.5 } },
        hovertemplate: '%{x}<br>등평: %{y:.2f}<br>학과: %{customdata[0]}%{customdata[1]}<extra>미응시(참고제외)</extra>',
      });
    }
  }

  // 최종합격: 추가합격 점은 주황색으로 구분
  // PC: trace에 selectedpoints 직접 적용 (정상 동작)
  // 모바일: Plotly.react 이후 requestAnimationFrame으로 restyle 재적용 (초기 렌더 타이밍 우회)
  let _addIdx = null, _finalBoxIdx = -1;
  {
    const sx = [], sy = [], sd = [], addIdx = [];
    for (const u of univs)
      for (const r of byUniv[u])
        if (r['상태구분'] === '최종합격') {
          const isAdd = !!r['is_additional'];
          sx.push(u); sy.push(r[gradeKey]);
          sd.push([r['모집단위'] || '', isAdd ? '추가합격' : '최초합격', extraInfo(r), r['학번'] || '', r['대입연도']]);
          if (isAdd) addIdx.push(sx.length - 1);
        }
    if (sx.length) {
      const trace = {
        type: 'box', x: sx, y: sy, name: '최종합격',
        customdata: sd,
        marker: { color: COLORS['최종합격'], size: 5 },
        line: { color: COLORS['최종합격'] },
        boxpoints: 'all', jitter: 0.3, pointpos: 0,
        hovertemplate: '%{x}<br>등평: %{y:.2f}<br>학과: %{customdata[0]}<br>구분: %{customdata[1]}%{customdata[2]}<extra>최종합격</extra>',
      };
      if (state.showAdditional && addIdx.length) {
        if (!isMobile) {
          // PC: trace에 직접 포함
          trace.selectedpoints = addIdx;
          trace.selected   = { marker: { color: ADD_COLOR,   size: 7,  opacity: 1   } };
          trace.unselected = { marker: { color: FIRST_COLOR, size: 5,  opacity: 0.8 } };
        } else {
          // 모바일: Plotly.react 이후 restyle로 강제 적용
          _finalBoxIdx = traces.length;
          _addIdx = addIdx.slice();
        }
      }
      traces.push(trace);
      if (state.showAdditional && addIdx.length) {
        traces.push({ type: 'scatter', x: [null], y: [null], mode: 'markers',
          name: '└ 최초합격', showlegend: true,
          marker: { color: FIRST_COLOR, size: 8, opacity: 0.8 } });
        traces.push({ type: 'scatter', x: [null], y: [null], mode: 'markers',
          name: '└ 추가합격', showlegend: true,
          marker: { color: ADD_COLOR, size: 8 } });
      }
    }
  }

  // 내 등급 수평선
  const shapes = [], annotations = [];
  if (state.myGrade) {
    shapes.push({
      type: 'line', x0: 0, x1: 1, xref: 'paper',
      y0: state.myGrade, y1: state.myGrade, yref: 'y',
      line: { color: '#e74c3c', width: 2, dash: 'dash' },
    });
    annotations.push({
      x: 1.01, xref: 'paper', xanchor: 'left',
      y: state.myGrade, yref: 'y',
      text: `내 등급 ${state.myGrade.toFixed(2)}`,
      showarrow: false, font: { color: '#e74c3c', size: 12, weight: 700 },
      bgcolor: 'rgba(255,255,255,0.85)',
    });
  }

  const yearLabel = [...state.years].sort().join(', ');
  const typeLabel = state.filter === '전체' ? '' : ` · ${state.filter}`;
  const gradeLabel = gradeKey === '국영수과 등평' ? '국영수과' : '전교과';

  // 모바일: 화면 너비에 맞춰 초기 표시, 핀치줌으로 확대/축소
  const chartW = isMobile ? Math.max(window.innerWidth - 16, univs.length * 36 + 160) : undefined;
  const chartH = isMobile ? 440 : 620;

  // 줌 상태 복원 (plotly_relayout에서 추적한 값 사용)
  // x축: 현재 카테고리 범위 안에 있을 때만 복원 (필터 변경 후 범위 이탈 방지)
  const _yRange = Z.y;
  const _xRangeSafe = (Z.x &&
    Z.x[0] >= -0.5 && Z.x[0] <= univs.length - 0.5 &&
    Z.x[1] >= -0.5 && Z.x[1] <= univs.length - 0.5)
    ? Z.x : undefined;

  // Plotly.react() 내부에서 plotly_relayout이 발생해 Z.y/Z.x가 덮어써질 수 있으므로
  // react 전에 저장하고 완료 후 복원
  const _savedZoomY = Z.y.slice();
  const _savedZoomX = Z.x ? Z.x.slice() : null;

  // 등급 모드: x축 라벨에 티어 이모지 접두 (🟢안정 🟡적정 🟠도전)
  const tickText = univs.map(u => {
    const base = `${u}(${nFinal[u]}/${nTotal[u]})`;
    if (cfg.tierColor && state.myGrade != null) {
      const t = tierOf(state.myGrade, statOf[u]);
      if (t) return `${TIER[t].emoji} ${base}`;
    }
    return base;
  });

  await Plotly.react(cfg.target, traces, {
    title: { text: `${yearLabel}년 ${state.cat}${typeLabel} · ${gradeLabel} 기준`, font: { size: isMobile ? 12 : 16 } },
    xaxis: {
      title: { text: '대학명 (최종합격수 / 총지원수)', font: { size: 11 } },
      categoryorder: 'array', categoryarray: univs,
      tickangle: -38, tickvals: univs, ticktext: tickText,
      fixedrange: false,
      ...(_xRangeSafe ? { range: _xRangeSafe } : {}),
    },
    yaxis: { title: gradeKey, range: _yRange, fixedrange: false },
    dragmode: 'pan',
    clickmode: 'event',
    boxmode: 'group',
    legend: isMobile
      ? { orientation: 'h', y: 1.08, x: 0, xanchor: 'left' }
      : { orientation: 'h', y: 1.08, x: 1, xanchor: 'right' },
    height: chartH, width: chartW,
    margin: { b: 170, t: 60, r: isMobile ? 8 : 80, l: isMobile ? 38 : 60 },
    plot_bgcolor: '#fafbff',
    shapes, annotations,
  }, {
    responsive: !isMobile,
    scrollZoom: true,
    displayModeBar: isMobile ? true : 'hover',
    modeBarButtonsToRemove: ['lasso2d', 'select2d', 'toImage'],
  });
  bindChartEvents(cfg.target);

  // react 완료 후 줌 상태 복원
  Z.y = _savedZoomY;
  Z.x = _savedZoomX;

  // 모바일: Plotly 렌더 완료 후 다음 프레임에서 selectedpoints 강제 재적용
  // (모바일 브라우저에서 초기 렌더 시 selected.marker.color 미적용 문제 우회)
  if (isMobile && state.showAdditional && _addIdx && _addIdx.length && _finalBoxIdx >= 0) {
    const capturedIdx = _finalBoxIdx;
    const capturedAddIdx = _addIdx;
    requestAnimationFrame(() => {
      Plotly.restyle(cfg.target, {
        selectedpoints: [capturedAddIdx],
        'selected.marker.color': [ADD_COLOR],
        'selected.marker.size': [8],
        'selected.marker.opacity': [1],
        'unselected.marker.color': [FIRST_COLOR],
        'unselected.marker.size': [5],
        'unselected.marker.opacity': [0.8],
      }, [capturedIdx]);
    });
  }

  // 합격권 요약
  if (cfg.summaryTarget) updateSummary(byUniv, univs, gradeKey, cfg.summaryTarget);
}

// 위수염 기준 도전·적정·안정 요약 (맞춤 조회 등급 모드 · 전체 분석 탭 공통)
function updateSummary(byUniv, univs, gradeKey, elId = 'grade-summary') {
  const el = document.getElementById(elId);
  if (!el) return;
  if (!state.myGrade) { el.innerHTML = ''; return; }
  const g = state.myGrade;
  const stable = [], proper = [], reach = [];

  for (const u of univs) {
    const grades = byUniv[u].filter(r => r['상태구분'] === '최종합격').map(r => r[gradeKey]).filter(v => v != null);
    if (!grades.length) continue;
    const n = grades.length;
    const total = byUniv[u].length;  // 지원(성적 있는) 전체 = 합격+최종탈락+1차탈락
    const entry = { u, n, total };
    const t = tierOf(g, boxStats(grades));
    if      (t === 'stable') stable.push(entry);
    else if (t === 'proper') proper.push(entry);
    else if (t === 'reach')  reach.push(entry);
  }

  const LOW_N = 5;

  function chip(entry, cls) {
    const lowN = entry.n < LOW_N;
    const style = lowN ? 'opacity:0.5; border: 1px dashed #aaa;' : '';
    const warn = lowN ? ' ⚠' : '';
    const tip = lowN ? '합격 사례 수가 적어 참고용으로만 활용하세요'
                     : '합격 = 최종합격자 수, 지원 = 같은 조건의 최종합격·최종탈락·1차탈락 합계';
    return `<span class="chip ${cls}" style="${style}" title="${tip}">${entry.u} (합격 ${entry.n}/지원 ${entry.total})${warn}</span>`;
  }

  function group(title, color, cls, items, hint) {
    if (!items.length) return '';
    return `<div class="summary-group">
      <div class="summary-title" style="color:${color}">${title} (${items.length}개)</div>
      <div class="summary-chips">${items.map(e => chip(e, cls)).join('')}</div>
      <div class="hint">${hint}</div>
    </div>`;
  }

  const disclaimer = `<div style="font-size:11px;color:#888;margin-top:10px;line-height:1.7;">
    ※ 위 구분은 <b>최종합격자 분포만</b> 기준입니다. 같은 등급대에도 1차탈락·최종탈락 사례가 함께 있을 수 있으니,
    차트의 <b>회색(1차탈락)·연파랑(1차합격_최종탈락) 분포도 반드시 함께</b> 확인하세요. ('적정'이어도 지원 대비 합격 비율이 낮을 수 있습니다.)<br>
    ※ 실제 합격 여부와 다를 수 있으며, 전형별 특성·수능최저·면접 등 다양한 요소가 결과에 영향을 줍니다.
    ⚠ 표시는 합격 사례 ${LOW_N}건 미만으로 통계적 신뢰도가 낮습니다.
  </div>`;

  const groups =
    group('🟠 도전권', '#784212', 'chip-reach', reach, `합격자 Q3~위수염 구간 — 사례는 있으나 상위 성적 필요`) +
    group('🟡 적정권', '#9a7d0a', 'chip-proper', proper, `합격자 중간 50%(Q1~Q3) 구간 — 합격자 다수 분포권`) +
    group('🟢 안정권', '#1e8449', 'chip-stable', stable, `내 등급(${g.toFixed(2)}) ≤ 합격자 상위 25%(Q1) — 여유 있는 지원권`);
  el.innerHTML = groups + disclaimer;
}

// ── 차트 줌 상태 (plotly_relayout로 추적, layout에서 직접 읽지 않음) ───────
// 타깃 차트별로 분리 (전체 분석 'chart' / 맞춤 조회 'landing-chart')
const _zoom = {
  'chart':         { y: [9.2, 0.8], x: null },
  'landing-chart': { y: [9.2, 0.8], x: null },
};

// 모달 차트 전용 줌 상태
let _modalZoomY = [9.2, 0.8];
let _modalZoomX = null;

// ── 드릴다운 모달 ───────────────────────────────────────
let drillUniv = null;
let drillTab = 'year';
let drillStudent = null;  // { hakbun, year }
let showStuInfo = false;  // 이름/학번 표시 토글

function openDrilldown(univ, student = null) {
  drillUniv = univ;
  drillStudent = student;
  drillTab = student ? 'student' : 'year';
  // 새 대학 열 때 모달 줌 초기화
  _modalZoomY = [9.2, 0.8];
  _modalZoomX = null;

  // 학생별 탭 버튼 표시/숨김
  const stuBtn = document.getElementById('modal-tab-student');
  stuBtn.style.display = student ? '' : 'none';

  // 탭 버튼 초기화
  document.querySelectorAll('.modal-tab').forEach(b => b.classList.remove('on'));
  if (student) stuBtn.classList.add('on');
  else document.querySelector('.modal-tab').classList.add('on');

  const gradeLabel = state.grade === '국영수과 등평' ? '국영수과' : '전교과';
  const typeLabel  = state.filter === '전체' ? '전체 전형' : state.filter;
  document.getElementById('modal-title').textContent = univ;
  document.getElementById('modal-sub').textContent   = student
    ? `학생 지원 현황 (${student.year}년도)`
    : `${typeLabel} · ${gradeLabel} 기준`;
  document.getElementById('modal-toggle-add').checked = state.showAdditional;
  document.getElementById('modal-overlay').classList.add('open');

  // 패널 전환
  document.getElementById('modal-chart').style.display   = student ? 'none' : '';
  document.getElementById('modal-student').style.display = student ? '' : 'none';

  if (student) renderModalStudent();
  else renderModalYear();
}

function switchModalTab(tab, btn) {
  drillTab = tab;
  document.querySelectorAll('.modal-tab').forEach(b => b.classList.remove('on'));
  btn.classList.add('on');
  const isStudent = tab === 'student';
  document.getElementById('modal-chart').style.display   = isStudent ? 'none' : '';
  document.getElementById('modal-student').style.display = isStudent ? '' : 'none';
  if (tab === 'year') renderModalYear();
  else if (tab === 'dept') renderModalDept();
  else if (tab === 'type') renderModalType();
  else if (tab === 'student') renderModalStudent();
}

function renderModalStudent() {
  const el = document.getElementById('modal-student');
  if (!drillStudent) { el.innerHTML = ''; return; }
  const { hakbun, year } = drillStudent;
  const rows = RAW.filter(r => r['학번'] === hakbun && r['대입연도'] === year);
  if (!rows.length) {
    el.innerHTML = '<p style="text-align:center;color:#aaa;padding:40px">데이터가 없습니다</p>';
    return;
  }

  const statusOrder = { '최종합격': 0, '1차합격_최종탈락': 1, '1차탈락': 2 };
  rows.sort((a, b) => (statusOrder[a['상태구분']] ?? 3) - (statusOrder[b['상태구분']] ?? 3));
  const r0 = rows[0];
  const haknyeon = r0['학년'] || '';

  // 5개 등급 (null/NaN 제외)
  const gradeCols = [
    ['전교과', r0['전교과 등평']],
    ['국영수과', r0['국영수과 등평']],
    ['수학과학', r0['수학과학 등평']],
    ['수학', r0['수학 등평']],
    ['과학', r0['과학 등평']],
  ];
  const gradeHtml = gradeCols
    .filter(([, v]) => v != null && !isNaN(Number(v)))
    .map(([k, v]) => `<span style="white-space:nowrap">${k} <b>${Number(v).toFixed(2)}</b></span>`)
    .join('<span style="color:#ccc;margin:0 4px">|</span>');

  // 이름/학번은 토글 ON일 때만
  const identHtml = showStuInfo
    ? `<b style="font-size:15px;color:#1A2A4A">${r0['이름'] || ''}</b>
       <span style="color:#888;font-size:12px">학번: ${hakbun}</span>`
    : '';

  const badgeStyle = {
    '최종합격':         'background:#0c4da2',
    '추가합격':         'background:#e67e22',
    '1차합격_최종탈락': 'background:#5a9fd4',
    '1차탈락':          'background:#aaaaaa',
    '미응시':           'background:#bbbbbb',
  };

  let rows_html = '';
  rows.forEach(r => {
    const status = r['상태구분'];
    let badge = status;
    if (status === '최종합격' && r['is_additional']) badge = '추가합격';
    if (r['is_nonattend']) badge = '미응시';
    const bstyle = badgeStyle[badge] || 'background:#aaa';

    const notes = [];
    if (r['is_suneung_fail']) notes.push('⚠ 수능최저 미충족');
    if (r['wait_num'])        notes.push(`예비 ${r['wait_num']}번`);
    if (r['add_info'] && r['add_info'] !== '추가 합격' && r['add_info'] !== '추가합격')
      notes.push(r['add_info']);
    if (r['enroll_info'])     notes.push(r['enroll_info']);

    rows_html += `<tr>
      <td style="padding:7px 10px;font-weight:600">${r['대학교명'] || '-'}</td>
      <td style="padding:7px 10px;color:#444">${r['모집단위'] || '-'}</td>
      <td style="padding:7px 10px;color:#666;font-size:12px">${r['전형명칭'] || '-'}</td>
      <td style="padding:7px 10px;text-align:center">
        <span class="stu-badge" style="${bstyle}">${badge}</span>
      </td>
      <td style="padding:7px 10px;color:#888;font-size:12px">${notes.join(' · ') || ''}</td>
    </tr>`;
  });

  el.innerHTML = `
    <div style="padding:14px 16px 8px">
      <div class="stu-header">
        <label style="display:flex;align-items:center;gap:5px;cursor:pointer;font-size:12px;color:#555;white-space:nowrap;margin-right:6px">
          <input type="checkbox" ${showStuInfo ? 'checked' : ''}
            onchange="showStuInfo=this.checked; renderModalStudent()"
            style="width:14px;height:14px;cursor:pointer">
          이름/학번 표시
        </label>
        ${identHtml}
        <span style="font-size:13px;color:#555">${haknyeon ? haknyeon + '학년 · ' : ''}${year}년도 입시</span>
        <span style="font-size:12px;color:#999">총 ${rows.length}개교 지원</span>
        <div style="flex-basis:100%;margin-top:6px;display:flex;flex-wrap:wrap;gap:6px;font-size:13px;color:#444">
          ${gradeHtml}
        </div>
      </div>
      <table class="stu-table">
        <thead><tr>
          <th>대학교</th><th>모집단위</th><th>전형</th>
          <th style="text-align:center">결과</th><th>비고</th>
        </tr></thead>
        <tbody>${rows_html}</tbody>
      </table>
    </div>`;
}

function modalShapes() {
  const shapes = [], annotations = [];
  if (state.myGrade) {
    shapes.push({ type: 'line', x0: 0, x1: 1, xref: 'paper',
      y0: state.myGrade, y1: state.myGrade, yref: 'y',
      line: { color: '#e74c3c', width: 2, dash: 'dash' } });
    annotations.push({ x: 1.01, xref: 'paper', xanchor: 'left',
      y: state.myGrade, yref: 'y',
      text: `내 등급 ${state.myGrade.toFixed(2)}`,
      showarrow: false, font: { color: '#e74c3c', size: 11, weight: 700 },
      bgcolor: 'rgba(255,255,255,0.85)' });
  }
  return { shapes, annotations };
}

function buildBoxTraces(rows, xKey, order) {
  const traces = [];

  // 1차탈락
  const r1x = [], r1y = [], r1d = [];
  for (const r of rows)
    if (r['상태구분'] === '1차탈락') { r1x.push(r[xKey]); r1y.push(r[state.grade]); r1d.push([r['모집단위'] || '', extraInfo(r), r['학번'] || '', r['대입연도']]); }
  if (r1x.length) traces.push({
    type: 'box', x: r1x, y: r1y, name: '1차탈락 (참고)',
    customdata: r1d,
    marker: { color: COLORS['1차탈락'], opacity: 0.4, size: 5 }, line: { color: COLORS['1차탈락'] },
    fillcolor: 'rgba(204,204,204,0.1)',
    boxpoints: 'all', jitter: 0.35, pointpos: 0,
    hovertemplate: '%{x}<br>등평: %{y:.2f}<br>학과: %{customdata[0]}%{customdata[1]}<extra>1차탈락</extra>',
  });

  // 1차합격_최종탈락 (미응시는 별도 scatter X로 표시)
  {
    const sx = [], sy = [], sd = [];
    const nax = [], nay = [], nad = [];
    for (const r of rows)
      if (r['상태구분'] === '1차합격_최종탈락') {
        if (r['is_nonattend']) {
          nax.push(r[xKey]); nay.push(r[state.grade]); nad.push([r['모집단위'] || '', extraInfo(r), r['학번'] || '', r['대입연도']]);
        } else {
          sx.push(r[xKey]); sy.push(r[state.grade]); sd.push([r['모집단위'] || '', extraInfo(r), r['학번'] || '', r['대입연도']]);
        }
      }
    if (sx.length) {
      traces.push({
        type: 'box', x: sx, y: sy, name: '1차합격_최종탈락',
        customdata: sd,
        marker: { color: COLORS['1차합격_최종탈락'], size: 5 },
        line: { color: COLORS['1차합격_최종탈락'] },
        boxpoints: 'all', jitter: 0.35, pointpos: 0,
        hovertemplate: '%{x}<br>등평: %{y:.2f}<br>학과: %{customdata[0]}%{customdata[1]}<extra>1차합격_최종탈락</extra>',
      });
    }
    if (nax.length) {
      traces.push({
        type: 'scatter', mode: 'markers', x: nax, y: nay,
        name: '└ 미응시(참고제외)', customdata: nad,
        marker: { color: '#aaa', size: 9, symbol: 'x-thin', opacity: 0.9,
                  line: { color: '#aaa', width: 2.5 } },
        hovertemplate: '%{x}<br>등평: %{y:.2f}<br>학과: %{customdata[0]}%{customdata[1]}<extra>미응시(참고제외)</extra>',
      });
    }
  }

  // 최종합격: selectedpoints로 추가합격 점만 주황색 (박스 통계/폭 불변)
  {
    const sx = [], sy = [], sd = [], addIdx = [];
    for (const r of rows)
      if (r['상태구분'] === '최종합격') {
        const isAdd = !!r['is_additional'];
        sx.push(r[xKey]); sy.push(r[state.grade]);
        sd.push([r['모집단위'] || '', isAdd ? '추가합격' : '최초합격', extraInfo(r), r['학번'] || '', r['대입연도']]);
        if (isAdd) addIdx.push(sx.length - 1);
      }
    if (sx.length) {
      const trace = {
        type: 'box', x: sx, y: sy, name: '최종합격',
        customdata: sd,
        marker: { color: COLORS['최종합격'], size: 5 },
        line: { color: COLORS['최종합격'] },
        boxpoints: 'all', jitter: 0.35, pointpos: 0,
        hovertemplate: '%{x}<br>등평: %{y:.2f}<br>학과: %{customdata[0]}<br>구분: %{customdata[1]}%{customdata[2]}<extra>최종합격</extra>',
      };
      if (state.showAdditional && addIdx.length) {
        trace.selectedpoints = addIdx;
        trace.selected   = { marker: { color: ADD_COLOR,   size: 7,  opacity: 1   } };
        trace.unselected = { marker: { color: FIRST_COLOR, size: 5,  opacity: 0.8 } };
      }
      traces.push(trace);
      if (state.showAdditional && addIdx.length) {
        traces.push({ type: 'scatter', x: [null], y: [null], mode: 'markers',
          name: '└ 최초합격', showlegend: true,
          marker: { color: FIRST_COLOR, size: 8, opacity: 0.8 } });
        traces.push({ type: 'scatter', x: [null], y: [null], mode: 'markers',
          name: '└ 추가합격', showlegend: true,
          marker: { color: ADD_COLOR, size: 8 } });
      }
    }
  }
  return traces;
}

// 모달 차트 공통 Plotly config
const MODAL_CONFIG = {
  responsive: true,
  scrollZoom: true,
  dragmode: 'pan',
  displayModeBar: 'hover',
  modeBarButtonsToRemove: ['lasso2d', 'select2d', 'toImage'],
};

// 모달 위에서 휠 시 페이지 스크롤 차단
document.getElementById('modal-chart').addEventListener('wheel', e => {
  e.stopPropagation();
}, { passive: false });

// 모달 줌 save/restore 헬퍼
let _modalRelayoutBound = false;

async function modalReact(traces, layout) {
  const savedY = _modalZoomY.slice();
  const savedX = _modalZoomX ? _modalZoomX.slice() : null;
  await Plotly.react('modal-chart', traces, {
    ...layout,
    yaxis: { ...layout.yaxis, range: savedY },
    ...(savedX ? { xaxis: { ...layout.xaxis, range: savedX } } : { xaxis: layout.xaxis }),
  }, MODAL_CONFIG);
  // Plotly.react() 후에야 .on()이 사용 가능 — 최초 1회만 등록
  if (!_modalRelayoutBound) {
    _modalRelayoutBound = true;
    document.getElementById('modal-chart').on('plotly_relayout', ev => {
      if (ev['yaxis.range[0]'] !== undefined) {
        _modalZoomY = [ev['yaxis.range[0]'], ev['yaxis.range[1]']];
      } else if (ev['yaxis.autorange']) {
        _modalZoomY = [9.2, 0.8];
      }
      if (ev['xaxis.range[0]'] !== undefined) {
        _modalZoomX = [ev['xaxis.range[0]'], ev['xaxis.range[1]']];
      } else if (ev['xaxis.autorange']) {
        _modalZoomX = null;
      }
    });
  }
  // Plotly.react()가 내부에서 plotly_relayout을 발생시켜 덮어쓸 수 있으므로 복원
  _modalZoomY = savedY;
  _modalZoomX = savedX;
}

// 연도별 뷰
async function renderModalYear() {
  const gradeKey = state.grade;
  const allYears = [...YEARS].sort();
  const rows = RAW.filter(r =>
    r['대학교명'] === drillUniv &&
    state.gradeYears.has(r['학년']) &&
    (state.filter === '전체' || r['admission_type'] === state.filter) &&
    r[gradeKey] != null
  ).map(r => ({ ...r, _yearLabel: String(r['대입연도']) + '년' }));

  if (!rows.length) return;
  const traces = buildBoxTraces(rows, '_yearLabel', allYears.map(y => y + '년'));
  const { shapes, annotations } = modalShapes();

  await modalReact(traces, {
    xaxis: { title: '입시 연도', categoryorder: 'array',
             categoryarray: allYears.map(y => y + '년') },
    yaxis: { title: gradeKey, range: [9.2, 0.8] },
    boxmode: 'group', dragmode: 'pan',
    legend: { orientation: 'h', y: 1.1, x: 1, xanchor: 'right' },
    height: 420, margin: { b: 60, t: 16, r: 90 },
    plot_bgcolor: '#fafbff', shapes, annotations,
  });
}

// 학과별 뷰
async function renderModalDept() {
  const gradeKey = state.grade;
  const rows = RAW.filter(r =>
    r['대학교명'] === drillUniv &&
    state.years.has(r['대입연도']) &&
    state.gradeYears.has(r['학년']) &&
    (state.filter === '전체' || r['admission_type'] === state.filter) &&
    r[gradeKey] != null &&
    r['모집단위']
  );

  if (!rows.length) {
    Plotly.react('modal-chart', [], { height: 200 }, MODAL_CONFIG);
    return;
  }

  // 학과 정렬: 최종합격 중앙값 기준
  const byDept = {};
  for (const r of rows) (byDept[r['모집단위']] ??= []).push(r);
  const deptOrder = Object.entries(byDept)
    .map(([d, rs]) => {
      const g = rs.filter(r => r['상태구분'] === '최종합격').map(r => r[gradeKey]);
      const g2 = rs.filter(r => r['상태구분'] === '1차합격_최종탈락').map(r => r[gradeKey]);
      return { d, med: g.length ? median(g) : (g2.length ? median(g2) : 99) };
    })
    .sort((a, b) => a.med - b.med)
    .map(x => x.d);

  const traces = buildBoxTraces(rows, '모집단위', deptOrder);
  const { shapes, annotations } = modalShapes();
  const yearLabel = [...state.years].sort().join(', ');

  await modalReact(traces, {
    title: { text: `${yearLabel}년 선택`, font: { size: 12, color: '#aaa' } },
    xaxis: { title: '모집단위(학과)', categoryorder: 'array',
             categoryarray: deptOrder, tickangle: -35 },
    yaxis: { title: gradeKey, range: [9.2, 0.8] },
    boxmode: 'group', dragmode: 'pan',
    legend: { orientation: 'h', y: 1.1, x: 1, xanchor: 'right' },
    height: 440, margin: { b: 130, t: 30, r: 90 },
    plot_bgcolor: '#fafbff', shapes, annotations,
  });
}

// 전형별 뷰 (admission_type 필터 무시, 모든 전형명칭 비교)
async function renderModalType() {
  const gradeKey = state.grade;
  const isMedical = state.cat === '메디컬계열';
  const rows = RAW.filter(r =>
    r['대학교명'] === drillUniv &&
    state.years.has(r['대입연도']) &&
    state.gradeYears.has(r['학년']) &&
    r['is_medical'] === isMedical &&
    r[gradeKey] != null &&
    r['전형명칭']
  );

  if (!rows.length) {
    Plotly.react('modal-chart', [], { height: 200 }, MODAL_CONFIG);
    return;
  }

  const byType = {};
  for (const r of rows) (byType[r['전형명칭']] ??= []).push(r);
  const typeOrder = Object.entries(byType)
    .map(([t, rs]) => {
      const g  = rs.filter(r => r['상태구분'] === '최종합격').map(r => r[gradeKey]);
      const g2 = rs.filter(r => r['상태구분'] === '1차합격_최종탈락').map(r => r[gradeKey]);
      return { t, med: g.length ? median(g) : (g2.length ? median(g2) : 99) };
    })
    .sort((a, b) => a.med - b.med)
    .map(x => x.t);

  const traces = buildBoxTraces(rows, '전형명칭', typeOrder);
  const { shapes, annotations } = modalShapes();
  const yearLabel = [...state.years].sort().join(', ');

  await modalReact(traces, {
    title: { text: `${yearLabel}년 선택`, font: { size: 12, color: '#aaa' } },
    xaxis: { title: '전형명칭', categoryorder: 'array',
             categoryarray: typeOrder, tickangle: -35 },
    yaxis: { title: gradeKey, range: [9.2, 0.8] },
    boxmode: 'group', dragmode: 'pan',
    legend: { orientation: 'h', y: 1.1, x: 1, xanchor: 'right' },
    height: 440, margin: { b: 130, t: 30, r: 90 },
    plot_bgcolor: '#fafbff', shapes, annotations,
  });
}

// 모달 닫기
document.getElementById('modal-close').addEventListener('click', () => {
  document.getElementById('modal-overlay').classList.remove('open');
});
document.getElementById('modal-overlay').addEventListener('click', e => {
  if (e.target === document.getElementById('modal-overlay'))
    document.getElementById('modal-overlay').classList.remove('open');
});

// ── 차트 이벤트 바인딩 (타깃별 1회) ─────────────────────
const _boundCharts = new Set();
function bindChartEvents(id) {
  if (_boundCharts.has(id)) return;
  const gd = document.getElementById(id);
  if (!gd || !gd.on) return;  // 아직 Plotly 그래프가 아니면 다음 렌더에서 재시도
  _boundCharts.add(id);
  gd.on('plotly_click', data => {
    if (!data || !data.points || !data.points.length) return;
    const pt = data.points[0];
    const univ = pt.x || (pt.data && pt.data.x && pt.data.x[pt.pointIndex]);
    if (!univ) return;
    // customdata 마지막 두 값: 학번, 대입연도
    const cd = pt.customdata;
    const hakbun = Array.isArray(cd) ? cd[cd.length - 2] : null;
    const year   = Array.isArray(cd) ? cd[cd.length - 1] : null;
    const student = (hakbun && year) ? { hakbun, year } : null;
    openDrilldown(univ, student);
  });
  gd.on('plotly_relayout', ev => {
    const Z = _zoom[id]; if (!Z) return;
    if (ev['yaxis.range[0]'] !== undefined) {
      Z.y = [ev['yaxis.range[0]'], ev['yaxis.range[1]']];
    } else if (ev['yaxis.autorange']) {
      // autoscale/reset: 등급축은 데이터가 아니라 항상 1~9(역순 고정)로 되돌림
      Z.y = [9.2, 0.8];
      Plotly.relayout(id, { 'yaxis.range': [9.2, 0.8], 'yaxis.autorange': false });
    }
    if (ev['xaxis.range[0]'] !== undefined) Z.x = [ev['xaxis.range[0]'], ev['xaxis.range[1]']];
    else if (ev['xaxis.autorange'])         Z.x = null;
  });
}

// ── 맞춤 조회(랜딩) 드롭다운 채우기 ─────────────────────
function populateUnivSelect() {
  const sel = document.getElementById('univ-select');
  if (!sel) return;
  const isMedical = state.cat === '메디컬계열';
  const set = new Set();
  // 필터에 걸리는 모든 대학 (합격 사례가 없어도 — 지원/1차탈락만 있는 대학도 선택 가능)
  for (const r of RAW) {
    if (passFilters(r, isMedical)) set.add(r['대학교명']);
  }
  // 영문(라틴 문자로 시작하는) 대학을 한글보다 먼저 정렬, 각 그룹 내부는 가나다/알파벳순
  const isLatin = s => { const c = s.charCodeAt(0); return (c >= 65 && c <= 90) || (c >= 97 && c <= 122); };
  const univs = [...set].sort((a, b) => {
    const la = isLatin(a), lb = isLatin(b);
    if (la !== lb) return la ? -1 : 1;
    return a.localeCompare(b, 'ko');
  });
  const cur = sel.value;
  sel.innerHTML = '<option value="">— 대학을 선택하세요 —</option>' +
    univs.map(u => `<option value="${u}"${u === cur ? ' selected' : ''}>${u}</option>`).join('');
  if (cur && !set.has(cur)) sel.value = '';  // 현재 선택이 필터 밖이면 해제
}

// ── 맞춤 조회 렌더 ─────────────────────────────────────
function renderLanding() {
  populateUnivSelect();
  if (landingMode === 'univ') {
    const u = document.getElementById('univ-select').value;
    if (!u) {
      renderBox({ target: 'landing-chart', summaryTarget: null, forceEmpty: true,
                  placeholder: '위에서 대학을 선택하면 분포가 표시됩니다.' });
    } else {
      renderBox({ target: 'landing-chart', summaryTarget: null, onlyUniv: u });
    }
  } else { // grade
    if (state.myGrade == null) {
      renderBox({ target: 'landing-chart', summaryTarget: 'landing-summary', forceEmpty: true,
                  placeholder: "상단 '내 등급'을 입력하면 지원 가능 대학이 표시됩니다." });
    } else {
      renderBox({ target: 'landing-chart', summaryTarget: 'landing-summary', tierColor: true, tierSort: true,
                  univFilter: st => tierOf(state.myGrade, st) !== null,
                  placeholder: '입력한 등급으로 지원 가능 범위(위수염 이내)인 대학이 없습니다.' });
    }
  }
}

// ── 탭 / 모드 전환 ─────────────────────────────────────
function switchMainTab(tab) {
  mainTab = tab;
  document.getElementById('mtab-landing').classList.toggle('on', tab === 'landing');
  document.getElementById('mtab-full').classList.toggle('on', tab === 'full');
  document.getElementById('view-landing').style.display = tab === 'landing' ? '' : 'none';
  document.getElementById('view-full').style.display    = tab === 'full'    ? '' : 'none';
  render();
  // 숨겨졌다 보인 차트는 크기 재계산
  const id = tab === 'landing' ? 'landing-chart' : 'chart';
  requestAnimationFrame(() => { try { Plotly.Plots.resize(id); } catch (e) {} });
}

function switchLandingMode(mode) {
  landingMode = mode;
  document.getElementById('seg-univ').classList.toggle('on', mode === 'univ');
  document.getElementById('seg-grade').classList.toggle('on', mode === 'grade');
  document.getElementById('landing-univ-ctrl').style.display  = mode === 'univ'  ? '' : 'none';
  document.getElementById('landing-grade-ctrl').style.display = mode === 'grade' ? '' : 'none';
  renderLanding();
  requestAnimationFrame(() => { try { Plotly.Plots.resize('landing-chart'); } catch (e) {} });
}

document.getElementById('univ-select').addEventListener('change', renderLanding);

// ── 소수 선발 전형 경고 (특별전형/특기자전형) ─────────────
function updateSpecialNotice() {
  const el = document.getElementById('special-notice');
  if (!el) return;
  const SMALL_INTAKE = ['특별전형', '특기자전형'];
  if (SMALL_INTAKE.includes(state.filter)) {
    el.style.display = '';
    el.innerHTML = `⚠ <b>${state.filter}</b>은 소수 인원을 선발해, 당해 지원자 풀에 따라 합격자 등급 평균이 `
      + `<b>크게 달라질 수 있습니다.</b> 연도별 편차가 크므로 해석에 특히 주의하세요.`;
  } else {
    el.style.display = 'none';
  }
}

// ── dispatcher ─────────────────────────────────────────
function render() {
  updateSpecialNotice();
  if (mainTab === 'landing') renderLanding();
  else renderBox({ target: 'chart', summaryTarget: 'grade-summary', tierColor: true });
}

// Firestore 데이터 모드에서 Firebase 모듈이 데이터 로드 완료 후 호출
window.__initWithData = function(records, years) {
  RAW   = records;
  YEARS = years;
  const def = YEARS.includes(2026) ? 2026 : YEARS[YEARS.length - 1];
  state.years = new Set([def].filter(Boolean));
  buildYearButtons();
  render();
};

render();

</script>
</body>
</html>"""


def prepare_and_visualize(file_path):
    if FIREBASE_DATA:
        html = HTML_TEMPLATE \
            .replace('__RAW_YEARS_DECL__', 'let RAW = [];\nlet YEARS = [];') \
            .replace('__FIREBASE_DATA__', 'true') \
            .replace('__FIREBASE_ENABLED__', 'true' if FIREBASE_ENABLED else 'false') \
            .replace('__FIREBASE_CONFIG__', json.dumps(FIREBASE_CONFIG, ensure_ascii=False))
    else:
        records, years = prepare_data(file_path)
        html = HTML_TEMPLATE \
            .replace('__RAW_YEARS_DECL__', 'const RAW = __DATA__;\nconst YEARS = __YEARS__;') \
            .replace('__DATA__', json.dumps(records, ensure_ascii=False)) \
            .replace('__YEARS__', json.dumps(years)) \
            .replace('__FIREBASE_DATA__', 'false') \
            .replace('__FIREBASE_ENABLED__', 'true' if FIREBASE_ENABLED else 'false') \
            .replace('__FIREBASE_CONFIG__', json.dumps(FIREBASE_CONFIG, ensure_ascii=False))

    if FIREBASE_DATA:
        output_path = os.path.join(os.path.dirname(os.path.abspath(file_path)), 'index.html')
    else:
        output_path = os.path.splitext(file_path)[0] + '_admission_analysis.html'
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'저장: {output_path}')

    import webbrowser
    webbrowser.open(output_path)
    return output_path


if __name__ == '__main__':
    excel_name = 'University_Admission_Results_2022-2026.xlsx'
    default_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), excel_name)
    excel_path = sys.argv[1] if len(sys.argv) > 1 else default_path

    if not FIREBASE_DATA and not os.path.exists(excel_path):
        print(f"파일 없음: {excel_path}")
        sys.exit(1)

    prepare_and_visualize(excel_path)
