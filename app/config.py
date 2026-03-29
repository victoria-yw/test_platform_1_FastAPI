import os

# 演示：生产环境请用环境变量 SESSION_SECRET，且密码应哈希后存库
SESSION_SECRET = os.environ.get("SESSION_SECRET", "dev-only-change-in-production")

DEMO_USERNAME = "admin"
DEMO_PASSWORD = "123456"

SESSION_USER = "user"

# 用例泛化上游
GENERALIZE_API_FILE_FIELD = os.environ.get("GENERALIZE_API_FILE_FIELD", "file").strip() or "file"
GENERALIZE_API_BEARER = os.environ.get("GENERALIZE_API_BEARER", "").strip()


def generalize_api_url() -> str:
    return os.environ.get("GENERALIZE_API_URL", "").strip()


# 用例生成上游
GENERATE_API_FILE_FIELD = os.environ.get("GENERATE_API_FILE_FIELD", "file").strip() or "file"
GENERATE_API_BEARER = os.environ.get("GENERATE_API_BEARER", "").strip()


def generate_api_url() -> str:
    return os.environ.get("GENERATE_API_URL", "").strip()
