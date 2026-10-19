"""MySQL 数据访问层：基于 PyMySQL + DBUtils 连接池。"""
import pymysql
from dbutils.pooled_db import PooledDB

import config

_pool = None


def get_pool():
    global _pool
    if _pool is None:
        _pool = PooledDB(
            creator=pymysql,
            maxconnections=10,
            mincached=2,
            maxcached=5,
            ping=1,
            cursorclass=pymysql.cursors.DictCursor,
            **config.DB_CONFIG,
        )
    return _pool


def query(sql, params=None):
    """执行查询，返回 dict 列表。params 为 None 时不做占位符替换。"""
    with get_pool().connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(sql, params)
            return cursor.fetchall()


def query_one(sql, params=None):
    """执行查询，返回单条 dict 或 None。"""
    with get_pool().connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(sql, params)
            return cursor.fetchone()


def execute(sql, params=None):
    """执行写操作，返回受影响行数。"""
    with get_pool().connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(sql, params)
            conn.commit()
            return cursor.rowcount


def insert(sql, params=None):
    """执行插入，返回自增主键。"""
    with get_pool().connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(sql, params)
            conn.commit()
            return cursor.lastrowid


def execute_many(sql, seq_of_params):
    """批量执行写操作，返回受影响行数。"""
    with get_pool().connection() as conn:
        with conn.cursor() as cursor:
            count = cursor.executemany(sql, seq_of_params)
            conn.commit()
            return count


def execute_in_transaction(operations):
    """在同一事务中执行多个 (sql, params) 写操作。"""
    with get_pool().connection() as conn:
        try:
            with conn.cursor() as cursor:
                for sql, params in operations:
                    cursor.execute(sql, params or ())
            conn.commit()
        except Exception:
            conn.rollback()
            raise


def clear_tables(tables):
    """在单连接上清空多张表。

    连接池下 SET 会话变量不跨连接，因此临时关闭外键检查与 DELETE
    必须在同一个连接内完成。
    """
    with get_pool().connection() as conn:
        try:
            with conn.cursor() as cursor:
                cursor.execute('SET FOREIGN_KEY_CHECKS = 0')
                for table in tables:
                    cursor.execute(f'DELETE FROM {table}')
                cursor.execute('SET FOREIGN_KEY_CHECKS = 1')
            conn.commit()
        except Exception:
            conn.rollback()
            raise


def init_db():
    """建表（幂等）并确保默认管理员存在。首次启动时调用。"""
    from werkzeug.security import generate_password_hash

    ddl = [
        """CREATE TABLE IF NOT EXISTS users (
            id INT AUTO_INCREMENT PRIMARY KEY,
            username VARCHAR(50) NOT NULL UNIQUE,
            password_hash VARCHAR(255) NOT NULL,
            role VARCHAR(20) NOT NULL DEFAULT 'user',
            email VARCHAR(100) NOT NULL UNIQUE,
            contact VARCHAR(20) NOT NULL DEFAULT '',
            register_time DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
            last_login_time DATETIME NULL
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4""",
        """CREATE TABLE IF NOT EXISTS classes (
            id INT AUTO_INCREMENT PRIMARY KEY,
            class_id VARCHAR(20) NOT NULL UNIQUE,
            college VARCHAR(50) NOT NULL DEFAULT '',
            class_name VARCHAR(50) NOT NULL UNIQUE,
            advisor VARCHAR(20) NOT NULL DEFAULT '',
            created_at VARCHAR(10) NOT NULL DEFAULT ''
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4""",
        """CREATE TABLE IF NOT EXISTS students (
            id INT AUTO_INCREMENT PRIMARY KEY,
            student_id VARCHAR(20) NOT NULL UNIQUE,
            name VARCHAR(20) NOT NULL DEFAULT '',
            gender VARCHAR(2) NOT NULL DEFAULT '男',
            college VARCHAR(50) NOT NULL DEFAULT '',
            class_name VARCHAR(50) NOT NULL DEFAULT '',
            INDEX idx_students_college (college),
            INDEX idx_students_class (class_name)
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4""",
        """CREATE TABLE IF NOT EXISTS courses (
            id INT AUTO_INCREMENT PRIMARY KEY,
            student_id VARCHAR(20) NOT NULL,
            course_name VARCHAR(50) NOT NULL DEFAULT '',
            credit DECIMAL(4,1) NOT NULL DEFAULT 0,
            score DECIMAL(5,1) NOT NULL DEFAULT 0,
            course_type VARCHAR(30) NOT NULL DEFAULT '',
            course_nature VARCHAR(10) NOT NULL DEFAULT '',
            semester VARCHAR(30) NOT NULL DEFAULT '',
            gpa DECIMAL(3,1) NOT NULL DEFAULT 0,
            credit_gpa DECIMAL(5,2) NOT NULL DEFAULT 0,
            INDEX idx_courses_student (student_id),
            CONSTRAINT fk_courses_student FOREIGN KEY (student_id)
                REFERENCES students(student_id) ON DELETE CASCADE
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4""",
        """CREATE TABLE IF NOT EXISTS dormitories (
            id INT AUTO_INCREMENT PRIMARY KEY,
            student_id VARCHAR(20) NOT NULL UNIQUE,
            name VARCHAR(20) NOT NULL DEFAULT '',
            college VARCHAR(50) NOT NULL DEFAULT '',
            class_name VARCHAR(50) NOT NULL DEFAULT '',
            building VARCHAR(20) NOT NULL DEFAULT '',
            room VARCHAR(10) NOT NULL DEFAULT '',
            bed VARCHAR(10) NOT NULL DEFAULT ''
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4""",
        """CREATE TABLE IF NOT EXISTS competitions (
            id INT AUTO_INCREMENT PRIMARY KEY,
            certificate_id VARCHAR(20) NOT NULL UNIQUE,
            competition_name VARCHAR(100) NOT NULL DEFAULT '',
            award_level VARCHAR(20) NOT NULL DEFAULT '',
            award_year VARCHAR(9) NOT NULL DEFAULT '',
            college VARCHAR(50) NOT NULL DEFAULT '',
            class_name VARCHAR(50) NOT NULL DEFAULT '',
            award_students VARCHAR(200) NOT NULL DEFAULT '',
            contact VARCHAR(20) NOT NULL DEFAULT '',
            advisor VARCHAR(20) NOT NULL DEFAULT ''
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4""",
        """CREATE TABLE IF NOT EXISTS notices (
            id INT AUTO_INCREMENT PRIMARY KEY,
            notice_id VARCHAR(20) NOT NULL UNIQUE,
            title VARCHAR(100) NOT NULL DEFAULT '',
            content TEXT NOT NULL,
            priority VARCHAR(10) NOT NULL DEFAULT 'normal',
            publisher VARCHAR(50) NOT NULL DEFAULT '',
            publish_time DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4""",
        """CREATE TABLE IF NOT EXISTS warning_status (
            id INT AUTO_INCREMENT PRIMARY KEY,
            student_id VARCHAR(20) NOT NULL,
            warning_type VARCHAR(20) NOT NULL,
            warning_key VARCHAR(60) NOT NULL DEFAULT '',
            status VARCHAR(20) NOT NULL DEFAULT 'pending',
            updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
                ON UPDATE CURRENT_TIMESTAMP,
            UNIQUE KEY uk_warning (student_id, warning_type, warning_key)
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4""",
    ]
    for sql in ddl:
        execute(sql)

    # 确保默认管理员存在（仅存哈希，绝不明文）
    if not query_one('SELECT id FROM users WHERE username = %s', (config.ADMIN_USERNAME,)):
        execute(
            'INSERT INTO users (username, password_hash, role, email) VALUES (%s, %s, %s, %s)',
            (config.ADMIN_USERNAME, generate_password_hash(config.ADMIN_PASSWORD),
             'admin', 'admin@example.com'),
        )
