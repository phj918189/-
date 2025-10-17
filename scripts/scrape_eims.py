from playwright.sync_api import sync_playwright
from dotenv import load_dotenv
import os, time, pandas as pd, pathlib, logging
from datetime import datetime

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('storage/logs/scrape.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

BASE = pathlib.Path(__file__).resolve().parents[1]
load_dotenv(BASE / ".env")
ID = os.getenv("EIMS_ID")
PW = os.getenv("EIMS_PW")

# 값이 비면 바로 중단(원인 명확화)
if not ID or not PW:
    logger.error(f"EIMS_ID/EIMS_PW가 비어 있습니다. {BASE / '.env'} 파일을 확인하세요.")
    raise SystemExit(
        f"EIMS_ID/EIMS_PW가 비어 있습니다. {BASE / '.env'} 파일을 확인하세요."
    )

# BASE = pathlib.Path(__file__).resolve().parents[1]
DL_DIR = BASE / "storage" / "downloads"
DL_DIR.mkdir(parents=True, exist_ok=True)

# 가능한 로그인 URL들
LOGIN_URLS = [
    # "https://xn--lu5b7kx8m.kr/login.go",
    # "https://www.xn--lu5b7kx8m.kr/login.go", 
    # "https://www.xn--lu5b7kx8m.kr/init.go",
    "https://xn--lu5b7kx8m.kr/init.go"
]
FIELD_URL = "https://www.xn--lu5b7kx8m.kr/ms/field_water.do"

def today(): 
    return datetime.now().strftime("%Y-%m-%d")

def fetch_excel_df(date_from=None, date_to=None):
    date_from = date_from or today()
    date_to = date_to or today()
    
    logger.info(f"엑셀 다운로드 시작: {date_from} ~ {date_to}")

    browser = None
    context = None
    page = None
    
    try:
        with sync_playwright() as p:
            logger.info("브라우저 시작 중...")
            browser = p.chromium.launch(
                headless=False,  # 디버깅을 위해 브라우저 창 표시
                args=[
                    '--no-sandbox', 
                    '--disable-dev-shm-usage',
                    '--disable-blink-features=AutomationControlled',  # 자동화 감지 방지
                    '--disable-web-security',
                    '--disable-features=VizDisplayCompositor',
                    '--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
                ]
            )
            context = browser.new_context(
                accept_downloads=True,
                viewport={'width': 1920, 'height': 1080},  # 뷰포트 설정
                user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                extra_http_headers={
                    'Accept-Language': 'ko-KR,ko;q=0.9,en;q=0.8'
                }
            )
            page = context.new_page()
            
            # 페이지 로딩 타임아웃 설정
            page.set_default_timeout(60000)  # 60초로 증가
            page.set_default_navigation_timeout(60000)
            # 올바른 로그인 페이지 찾기
            login_success = False
            for login_url in LOGIN_URLS:
                try:
                    logger.info(f"로그인 페이지 시도 중: {login_url}")
                    page.goto(login_url, timeout=60000)
                    
                    # 페이지가 완전히 로드될 때까지 충분히 대기
                    logger.info("페이지 로딩 대기 중...")
                    page.wait_for_timeout(5000)  # 5초 대기
                    page.wait_for_load_state("domcontentloaded", timeout=30000)
                    page.wait_for_load_state("networkidle", timeout=30000)
                    
                    # 로그인 필드 찾기 - 더 유연한 방식
                    logger.info("로그인 필드 찾는 중...")
                    
                    # 여러 방법으로 사용자 필드 찾기
                    user_selectors = [
                        '#editx_emp_id',
                        'input[name="editx_emp_id"]',
                        'input[id="editx_emp_id"]',
                        'input[name*="emp"]',
                        'input[name*="user"]',
                        'input[name*="id"]',
                        'input[type="text"]'
                    ]
                    
                    password_selectors = [
                        '#old_password',
                        'input[name="old_password"]',
                        'input[id="old_password"]',
                        'input[type="password"]'
                    ]
                    
                    user_field = None
                    password_field = None
                    
                    # 사용자 필드 찾기
                    for selector in user_selectors:
                        try:
                            elements = page.query_selector_all(selector)
                            for elem in elements:
                                if elem.is_visible() and elem.is_enabled():
                                    user_field = elem
                                    logger.info(f"사용자 필드 발견: {selector}")
                                    break
                            if user_field:
                                break
                        except:
                            continue
                    
                    # 비밀번호 필드 찾기
                    for selector in password_selectors:
                        try:
                            elements = page.query_selector_all(selector)
                            for elem in elements:
                                if elem.is_visible() and elem.is_enabled():
                                    password_field = elem
                                    logger.info(f"비밀번호 필드 발견: {selector}")
                                    break
                            if password_field:
                                break
                        except:
                            continue
                    
                    if user_field and password_field:
                        logger.info(f"✅ 로그인 페이지 발견: {login_url}")
                        
                        # 로그인 시도 - 완전히 수정된 구조
                        logger.info("로그인 정보 입력 시작...")
                        
                        # 페이지가 안정화될 때까지 대기
                        page.wait_for_timeout(2000)
                        
                        # 요소가 실제로 상호작용 가능한지 다시 확인
                        if not user_field.is_visible() or not user_field.is_enabled():
                            logger.warning("사용자 필드가 상호작용 불가능, 다시 찾는 중...")
                            user_field = page.query_selector('#editx_emp_id')
                            if not user_field:
                                user_field = page.query_selector('input[name="editx_emp_id"]')
                        
                        if not password_field.is_visible() or not password_field.is_enabled():
                            logger.warning("비밀번호 필드가 상호작용 불가능, 다시 찾는 중...")
                            password_field = page.query_selector('#old_password')
                            if not password_field:
                                password_field = page.query_selector('input[name="old_password"]')
                        
                        if not user_field or not password_field:
                            logger.error("로그인 필드를 찾을 수 없음")
                            continue
                        
                        # 요소를 화면에 보이도록 스크롤
                        user_field.scroll_into_view_if_needed()
                        password_field.scroll_into_view_if_needed()
                        
                        # 추가 대기
                        page.wait_for_timeout(1000)
                        
                        # 사용자 ID 입력
                        logger.info("사용자 ID 입력 중...")
                        user_field.click()
                        page.wait_for_timeout(500)
                        user_field.fill("")  # 기존 내용 클리어
                        page.wait_for_timeout(500)
                        user_field.type(ID, delay=200)  # 더 천천히 타이핑
                        
                        # 비밀번호 입력
                        logger.info("비밀번호 입력 중...")
                        password_field.click()
                        page.wait_for_timeout(500)
                        password_field.fill("")  # 기존 내용 클리어
                        page.wait_for_timeout(500)
                        password_field.type(PW, delay=200)  # 더 천천히 타이핑
                        
                        logger.info("✅ 로그인 정보 입력 완료")
                        
                        # 로그인 버튼 찾기 및 클릭
                        logger.info("로그인 버튼 찾는 중...")
                        
                        # 방법 1: Enter 키로 먼저 시도 (가장 안정적)
                        logger.info("방법 1: Enter 키로 로그인 시도")
                        password_field.press("Enter")
                        logger.info("✅ Enter 키로 로그인 시도 완료")
                        
                        # 방법 2: 버튼 클릭으로 시도
                        logger.info("방법 2: 로그인 버튼 클릭 시도")
                        button_selectors = [
                            'input[type="submit"]',
                            'button[type="submit"]',
                            'button:has-text("로그인")',
                            'button:has-text("확인")',
                            'input[value*="로그인"]',
                            'input[value*="확인"]'
                        ]
                        
                        login_clicked = False
                        for selector in button_selectors:
                            try:
                                buttons = page.query_selector_all(selector)
                                for btn in buttons:
                                    try:
                                        text = btn.inner_text().lower().strip()
                                        value = btn.get_attribute("value") or ""
                                        logger.info(f"버튼 확인: 텍스트='{text}', 값='{value}'")
                                        
                                        if any(keyword in text or keyword in value.lower() for keyword in ["login", "로그인", "확인", "submit"]):
                                            if btn.is_visible() and btn.is_enabled():
                                                btn.scroll_into_view_if_needed()
                                                page.wait_for_timeout(1000)
                                                btn.click()
                                                logger.info(f"✅ 로그인 버튼 클릭: '{text or value}'")
                                                login_clicked = True
                                                break
                                    except Exception as e:
                                        logger.warning(f"버튼 클릭 실패: {e}")
                                        continue
                                if login_clicked:
                                    break
                            except:
                                continue
                        
                        if not login_clicked:
                            logger.warning("로그인 버튼을 찾을 수 없음")
                        
                        logger.info("✅ 로그인 버튼 클릭 완료")
                        # 로그인 후 페이지 이동 대기
                        logger.info("로그인 후 페이지 로딩 대기 중...")
                        
                        # 로그인 처리 대기 - 더 긴 시간
                        logger.info("로그인 처리 대기 중... (10초)")
                        page.wait_for_timeout(10000)  # 10초 대기
                        
                        # 페이지 로딩 상태 확인
                        logger.info("페이지 로딩 상태 확인 중...")
                        try:
                            page.wait_for_load_state("networkidle", timeout=30000)
                            logger.info("네트워크 대기 완료")
                        except:
                            logger.warning("네트워크 대기 타임아웃, 계속 진행")
                        
                        # 추가 대기
                        page.wait_for_timeout(3000)
                        
                        # 로그인 성공 여부 확인
                        current_url = page.url
                        logger.info(f"현재 URL: {current_url}")
                        
                        # 페이지 제목 확인
                        try:
                            page_title = page.title()
                            logger.info(f"페이지 제목: {page_title}")
                        except:
                            logger.warning("페이지 제목을 가져올 수 없음")
                        
                        # 로그인 성공 조건 확인 - notify.do 포함
                        success_keywords = ["main", "home", "dashboard", "menu", "ms/field_water", "field_water", "notify.do", "ms/"]
                        failure_keywords = ["login", "init"]
                        
                        is_success = any(keyword in current_url.lower() for keyword in success_keywords)
                        is_failure = any(keyword in current_url.lower() for keyword in failure_keywords)
                        
                        logger.info(f"성공 키워드 매칭: {is_success}")
                        logger.info(f"실패 키워드 매칭: {is_failure}")
                        
                        if is_success or not is_failure:
                            logger.info("✅ 로그인 성공!")
                            
                            # 현재 URL이 이미 목표 페이지인지 확인
                            if "field_water" in current_url.lower():
                                logger.info("✅ 이미 목표 페이지에 도달됨! 추가 이동 불필요")
                                login_success = True
                                break
                            
                            login_success = True
                            break
                        else:
                            # 로그인 실패 시 오류 메시지 확인
                            logger.info("로그인 실패 - 오류 메시지 확인 중...")
                            try:
                                error_elements = page.query_selector_all('.error, .alert, .warning, [class*="error"], [class*="alert"], .msg, .message, .fail')
                                if error_elements:
                                    for elem in error_elements:
                                        error_text = elem.inner_text().strip()
                                        if error_text:
                                            logger.warning(f"로그인 오류 메시지: {error_text}")
                            except:
                                pass
                            
                            # 페이지 내용 일부 확인
                            try:
                                page_content = page.content()
                                if "로그인" in page_content or "login" in page_content.lower():
                                    logger.warning("페이지에 로그인 관련 내용이 여전히 있음")
                                else:
                                    logger.info("페이지에 로그인 관련 내용이 없음 - 성공 가능성 있음")
                            except:
                                pass
                            
                            logger.warning("로그인 실패 - 여전히 로그인 페이지에 있음")
                    
                except Exception as e:
                    logger.error(f"로그인 페이지 시도 실패 {login_url}: {e}")
                    continue
            
            if not login_success:
                logger.error("모든 로그인 URL에서 로그인에 실패했습니다.")
                raise Exception("모든 로그인 URL에서 로그인에 실패했습니다.")

            # 목표 페이지 확인 및 이동 (필요한 경우만)
            current_url = page.url
            logger.info(f"현재 URL: {current_url}")
            
            if "field_water" in current_url.lower():
                logger.info("✅ 이미 목표 페이지에 있음! 이동 불필요")
            else:
                logger.info("목표 페이지로 이동 필요")
                
                # 같은 도메인으로 이동 (올바른 경로로)
                if "notify.do" in current_url:
                    # notify.do에서 field_water.do로 이동 (public 폴더에서 ms 폴더로)
                    target_url = current_url.replace("/ms/public/notify.do", "/ms/field_water.do")
                else:
                    target_url = current_url.replace("/ms/", "/ms/field_water.do")
                
                logger.info(f"같은 도메인으로 이동: {target_url}")
                
                try:
                    # 실제 페이지 이동 전후 URL 확인
                    logger.info(f"이동 전 실제 URL: {page.url}")
                    page.goto(target_url, timeout=60000)
                    logger.info(f"page.goto() 완료")
                    
                    page.wait_for_load_state("networkidle", timeout=30000)
                    logger.info(f"네트워크 대기 완료")
                    
                    final_url = page.url
                    logger.info(f"이동 후 실제 URL: {final_url}")
                    
                    # 페이지 제목도 확인
                    try:
                        page_title = page.title()
                        logger.info(f"페이지 제목: {page_title}")
                    except:
                        logger.warning("페이지 제목을 가져올 수 없음")
                    
                    if "field_water" in final_url.lower():
                        logger.info("✅ 목표 페이지 도달 성공!")
                    else:
                        logger.error(f"❌ 목표 페이지 도달 실패: {final_url}")
                        logger.error(f"예상 URL: {target_url}")
                        logger.error(f"실제 URL: {final_url}")
                        raise Exception(f"목표 페이지에 도달할 수 없습니다: {final_url}")
                        
                except Exception as e:
                    logger.error(f"목표 페이지 이동 실패: {e}")
                    logger.error(f"현재 URL: {page.url}")
                    raise
            
            logger.info("✅ 목표 페이지 준비 완료")

            # (필요 시) 날짜 필터 셀렉터 채우기
            # page.fill('#dateFrom', date_from)
            # page.fill('#dateTo', date_to)
            # page.click('button:has-text("조회")')

            # 엑셀 출력 버튼 클릭 - 현재 페이지 확인 후
            logger.info("엑셀출력 버튼 클릭 시도 중...")
            
            # 현재 페이지 URL 재확인
            current_page_url = page.url
            logger.info(f"엑셀 버튼 클릭 전 현재 URL: {current_page_url}")
            
            if "field_water" not in current_page_url.lower():
                logger.error(f"❌ 목표 페이지가 아님! 현재 URL: {current_page_url}")
                raise Exception(f"목표 페이지에 있지 않습니다. 현재 URL: {current_page_url}")
            
            # 다운로드 이벤트 리스너 설정
            downloads = []
            def handle_download(download):
                downloads.append(download)
                logger.info(f"다운로드 이벤트 감지: {download.suggested_filename}")
            
            page.on("download", handle_download)
            
            # 엑셀 출력 버튼 클릭
            try:
                logger.info("엑셀출력 버튼 클릭 시도")
                with page.expect_download(timeout=30000) as dl_info:
                    page.click('button:has-text("엑셀출력")')
                download = dl_info.value
                logger.info(f"✅ 다운로드 성공: {download.suggested_filename}")
            except Exception as e:
                logger.error(f"엑셀출력 버튼 클릭 실패: {e}")
                raise Exception("엑셀 다운로드 버튼을 클릭할 수 없습니다.")
            
            filename = f"field_water_{date_from}_{int(time.time())}.xlsx"
            save_to = DL_DIR / filename
            
            logger.info(f"저장할 파일 경로: {save_to}")
            
            # 파일 저장 - 이전 방식대로 단순하게
            try:
                logger.info(f"파일 저장 중: {save_to}")
                download.save_as(str(save_to))
                logger.info("✅ 파일 저장 완료")
                
                # 파일 확인
                if save_to.exists():
                    file_size = save_to.stat().st_size
                    logger.info(f"파일 크기: {file_size} bytes")
                    if file_size > 0:
                        logger.info("✅ 파일 다운로드 성공!")
                    else:
                        logger.error("❌ 파일이 비어있음!")
                        raise Exception("다운로드된 파일이 비어있습니다.")
                else:
                    logger.error("❌ 파일이 저장되지 않음!")
                    raise Exception("파일 저장에 실패했습니다.")
                    
            except Exception as e:
                logger.error(f"파일 저장 중 오류 발생: {e}")
                raise

            # 엑셀 파일 읽기
            logger.info("엑셀 파일 읽기 시도 중...")
            try:
                df = pd.read_excel(save_to)
                logger.info(f"엑셀 데이터 로드 완료: {len(df)} 행")
                df.attrs["source_path"] = str(save_to)
                return df
            except Exception as e:
                logger.error(f"엑셀 파일 읽기 실패: {e}")
                raise
                
    except Exception as e:
        logger.error(f"전체 프로세스 중 오류 발생: {e}")
        raise
    finally:
        # 브라우저 정리
        try:
            if page:
                logger.info("페이지 정리 중...")
            if context:
                logger.info("컨텍스트 정리 중...")
                context.close()
            if browser:
                logger.info("브라우저 정리 중...")
                browser.close()
        except Exception as e:
            logger.error(f"브라우저 정리 중 오류: {e}")
