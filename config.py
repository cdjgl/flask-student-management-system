"""应用配置：从 .env / 环境变量读取，敏感信息不入库。"""
import os
import secrets

try:
    from dotenv import load_dotenv

    load_dotenv()
except ImportError:
    pass

# 数据库连接配置
DB_CONFIG = {
    'host': os.environ.get('DB_HOST', '127.0.0.1'),
    'port': int(os.environ.get('DB_PORT', '3306')),
    'user': os.environ.get('DB_USER', 'root'),
    'password': os.environ.get('DB_PASSWORD', ''),
    'database': os.environ.get('DB_NAME', 'student_management'),
    'charset': 'utf8mb4',
}

# 会话密钥：生产环境务必在 .env 中设置固定值
SECRET_KEY = os.environ.get('SECRET_KEY') or secrets.token_hex(32)

# 会话有效期（小时）
SESSION_LIFETIME_HOURS = int(os.environ.get('SESSION_LIFETIME_HOURS', '1'))

# 首次启动自动创建的管理员账号
ADMIN_USERNAME = os.environ.get('ADMIN_USERNAME', 'admin')
ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', 'admin123')
