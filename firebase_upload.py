"""
Firebase Firestore 업로더
========================
Excel 데이터를 Firestore admissionData 컬렉션에 업로드합니다.

사전 준비:
  1. pip install firebase-admin
  2. Firebase Console → 프로젝트 설정 → 서비스 계정 → Python →
     "새 비공개 키 생성" → 다운로드한 JSON을 이 스크립트와 같은 폴더에
     serviceAccountKey.json 로 저장
  3. Firestore 보안 규칙 업데이트 (아래 참고)

Firestore 보안 규칙 (Firebase Console → Firestore → 규칙):
  rules_version = '2';
  service cloud.firestore {
    match /databases/{database}/documents {
      match /Whitelist/{docId} {
        allow read: if request.auth != null;
        allow write: if false;
      }
      match /admissionData/{docId} {
        allow read: if request.auth != null;
        allow write: if false;
      }
    }
  }

업로드 구조:
  admissionData/
    meta  → { years: [2022, 2023, ...] }
    2022  → { records: [{...}, ...] }
    2023  → { records: [{...}, ...] }
    ...
"""

import os
import sys
import math

# ── 설정 ──────────────────────────────────────────────────────────────────────
EXCEL_NAME       = 'University_Admission_Results_2022-2026.xlsx'
SERVICE_ACCOUNT  = 'serviceAccountKey.json'   # Firebase 서비스 계정 키 파일명
PROJECT_ID       = 'hansung-admission'
# ─────────────────────────────────────────────────────────────────────────────


def to_native(val):
    """numpy/pandas 타입 → Python 기본 타입 (Firestore는 numpy 타입 불가)"""
    try:
        import numpy as np
        if isinstance(val, np.integer):
            return int(val)
        if isinstance(val, np.floating):
            return None if np.isnan(val) else float(val)
        if isinstance(val, np.bool_):
            return bool(val)
    except ImportError:
        pass
    if isinstance(val, float) and math.isnan(val):
        return None
    if isinstance(val, bool):
        return val
    return val


def clean_records(records):
    """레코드 리스트의 모든 값을 Firestore 호환 타입으로 변환"""
    return [{k: to_native(v) for k, v in r.items()} for r in records]


def upload(excel_path, key_path):
    import firebase_admin
    from firebase_admin import credentials, firestore

    # firebase-admin 초기화
    cred = credentials.Certificate(key_path)
    firebase_admin.initialize_app(cred, {'projectId': PROJECT_ID})
    db = firestore.client()

    # admission_analysis.py의 prepare_data 재사용
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from admission_analysis import prepare_data

    print(f"Excel 읽는 중: {excel_path}")
    records, years = prepare_data(excel_path)
    records = clean_records(records)
    print(f"전체 레코드 수: {len(records)}, 연도: {years}")

    # meta 문서
    db.collection('admissionData').document('meta').set({'years': years})
    print(f"admissionData/meta 업로드 완료 → years={years}")

    # 연도별 문서
    from collections import defaultdict
    by_year = defaultdict(list)
    for r in records:
        by_year[r['대입연도']].append(r)

    for year in years:
        year_records = by_year[year]
        db.collection('admissionData').document(str(year)).set({'records': year_records})
        print(f"admissionData/{year} 업로드 완료 → {len(year_records)}건")

    print("\n✅ Firestore 업로드 완료!")
    print("이제 admission_analysis.py를 실행해 Firebase HTML을 생성하세요.")


if __name__ == '__main__':
    script_dir = os.path.dirname(os.path.abspath(__file__))

    excel_path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(script_dir, EXCEL_NAME)
    key_path   = sys.argv[2] if len(sys.argv) > 2 else os.path.join(script_dir, SERVICE_ACCOUNT)

    if not os.path.exists(excel_path):
        print(f"[오류] Excel 파일 없음: {excel_path}")
        sys.exit(1)
    if not os.path.exists(key_path):
        print(f"[오류] 서비스 계정 키 없음: {key_path}")
        print("Firebase Console → 프로젝트 설정 → 서비스 계정 → 새 비공개 키 생성")
        print(f"다운로드한 JSON을 '{key_path}' 로 저장하세요.")
        sys.exit(1)

    upload(excel_path, key_path)
