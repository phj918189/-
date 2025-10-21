import yaml
import os
from pathlib import Path
from dotenv import load_dotenv

class Config:
    """설정 관리 클래스"""
    
    def __init__(self):
        self.config_path = Path("config.yaml")
        self.env_path = Path(".env")
        
        # 환경변수 로드
        if self.env_path.exists():
            load_dotenv(self.env_path)
        
        # YAML 설정 로드
        self.config = self._load_config()
    
    def _load_config(self):
        """설정 파일 로드"""
        if self.config_path.exists():
            with open(self.config_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f)
        return {}
    
    def get(self, key_path, default=None):
        """점 표기법으로 설정값 가져오기"""
        keys = key_path.split('.')
        value = self.config
        
        for key in keys:
            if isinstance(value, dict) and key in value:
                value = value[key]
            else:
                return default
        
        return value
    
    def get_env_or_config(self, env_key, config_key, default=None):
        """환경변수 우선, 없으면 설정파일에서 가져오기"""
        env_value = os.getenv(env_key)
        if env_value:
            return env_value
        
        return self.get(config_key, default)

# 전역 설정 인스턴스
config = Config()

# 사용 예시
if __name__ == "__main__":
    print("=== 설정 테스트 ===")
    print(f"데이터베이스 경로: {config.get('database.path', 'storage/lab.db')}")
    print(f"SMTP 호스트: {config.get_env_or_config('SMTP_HOST', 'email.smtp_host', 'smtp.gmail.com')}")
    print(f"엑셀 보관 기간: {config.get('file_management.excel_retention_days', 7)}일")
