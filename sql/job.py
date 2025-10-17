
import sys, pathlib
import logging

ROOT = pathlib.Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path: sys.path.insert(0, str(ROOT))

from scripts.scrape_eims import fetch_excel_df
from scripts.normalize import normalize
from scripts import db
from scripts.assign import run_assign
from scripts.notify import notify

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('storage/logs/job.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

def main():
    try:
        logger.info("=== EIMS 자동화 프로세스 시작 ===")
        
        # 1) 데이터베이스 초기화
        try:
            logger.info("1단계: 데이터베이스 초기화")
            db.init_db()
            logger.info("✅ 데이터베이스 초기화 완료")
        except Exception as e:
            logger.error(f"❌ 데이터베이스 초기화 실패: {e}")
            raise
        
        # 2) 엑셀 다운로드
        try:
            logger.info("2단계: 엑셀 다운로드")
            raw = fetch_excel_df()
            src = raw.attrs.get("source_path","")
            logger.info(f"✅ 엑셀 다운로드 완료: {src}")
        except Exception as e:
            logger.error(f"❌ 엑셀 다운로드 실패: {e}")
            logger.error("로그인 문제일 가능성이 높습니다. .env 파일과 네트워크 연결을 확인하세요.")
            raise
        
        # 3) 데이터 표준화
        try:
            logger.info("3단계: 데이터 표준화")
            df = normalize(raw)
            logger.info(f"✅ 데이터 표준화 완료: {len(df)} 행")
        except Exception as e:
            logger.error(f"❌ 데이터 표준화 실패: {e}")
            raise
        
        # 4) 데이터베이스 저장
        try:
            logger.info("4단계: 데이터베이스 저장")
            db.upsert_samples(df, src)
            logger.info("✅ 데이터베이스 저장 완료")
        except Exception as e:
            logger.error(f"❌ 데이터베이스 저장 실패: {e}")
            raise
        
        # 5) 작업 배정
        try:
            logger.info("5단계: 작업 배정")
            assigned = run_assign(df)
            logger.info(f"✅ 작업 배정 완료: {len(assigned)} 건")
        except Exception as e:
            logger.error(f"❌ 작업 배정 실패: {e}")
            raise
        
        # 6) 알림 발송
        try:
            logger.info("6단계: 알림 발송")
            notify(assigned)
            logger.info("✅ 알림 발송 완료")
        except Exception as e:
            logger.error(f"❌ 알림 발송 실패: {e}")
            logger.warning("알림 발송 실패했지만 프로세스는 계속 진행됩니다.")
        
        logger.info("=== EIMS 자동화 프로세스 완료 ===")
        
    except Exception as e:
        logger.error(f"프로세스 실행 중 오류 발생: {e}")
        logger.error("프로그램을 종료합니다.")
        sys.exit(1)

if __name__ == "__main__":
    main()
