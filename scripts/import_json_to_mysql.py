"""一次性迁移脚本：将 data/*.json 的历史数据脱敏后导入 MySQL。

脱敏规则（合规要求）：
  1. 密码一律哈希存储，绝不明文入库；
  2. 真实域名邮箱替换为 example.com 保留域；
  3. 联系电话统一归一化为 138000000xx 占位号段；
  4. 丢弃测试垃圾数据（如 TEST001 测试公告）。

用法：python scripts/import_json_to_mysql.py [--export-seed]
  --export-seed  导入完成后在 database/ 下生成 seed.sql（脱敏后的演示数据）
"""
import json
import os
import re
import sys
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pymysql
from werkzeug.security import generate_password_hash

import config
import db

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data')
DATABASE_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'database')

# 已知需丢弃的测试数据 ID
DROP_NOTICE_IDS = {'TEST001'}
# 邮箱域白名单外的一律替换为 example.com（保留域，RFC 2606）
REAL_DOMAIN_RE = re.compile(r'@(qq|163|126|gmail|hotmail|outlook|foxmail)\.', re.I)
PHONE_RE = re.compile(r'^1\d{10}$')


def load_json(filename, key):
    filepath = os.path.join(DATA_DIR, filename)
    if not os.path.exists(filepath):
        return []
    with open(filepath, 'r', encoding='utf-8') as f:
        return json.load(f).get(key, [])


def sanitize_email(email, username):
    if not email or REAL_DOMAIN_RE.search(email):
        return f'{username or "user"}@example.com'
    return email


def sanitize_phone(phone, index):
    """历史占位号段归一化为 138000000xx 系列，空值留空。"""
    return f'138{index:08d}' if phone and PHONE_RE.match(phone) else ''


def ensure_schema():
    """建库建表（与 database/schema.sql 等价，幂等）。"""
    conn = pymysql.connect(
        host=config.DB_CONFIG['host'], port=config.DB_CONFIG['port'],
        user=config.DB_CONFIG['user'], password=config.DB_CONFIG['password'],
        charset='utf8mb4',
    )
    try:
        with conn.cursor() as cursor:
            cursor.execute(
                f"CREATE DATABASE IF NOT EXISTS `{config.DB_CONFIG['database']}` "
                "DEFAULT CHARSET utf8mb4 COLLATE utf8mb4_unicode_ci"
            )
            cursor.execute(f"USE `{config.DB_CONFIG['database']}`")
            with open(os.path.join(DATABASE_DIR, 'schema.sql'), 'r', encoding='utf-8') as f:
                sql_text = f.read()
            # 按分号拆分，剔除纯注释行后逐条执行（跳过建库/USE，已单独处理）
            statements = [s.strip() for s in sql_text.split(';') if s.strip()]
            for stmt in statements:
                body = '\n'.join(line for line in stmt.splitlines()
                                 if not line.strip().startswith('--')).strip()
                if not body:
                    continue
                upper = body.upper()
                if 'CREATE DATABASE' in upper or upper.startswith('USE '):
                    continue
                cursor.execute(body)
        conn.commit()
    finally:
        conn.close()


def dedupe(rows, index):
    """按业务键去重（保留首条），历史 JSON 数据可能存在重复记录。"""
    seen, out = set(), []
    for row in rows:
        key = row[index]
        if key in seen:
            continue
        seen.add(key)
        out.append(row)
    return out


def migrate():
    phone_seq = 0

    # ---- 用户（脱敏：哈希密码 + example.com 邮箱 + 占位手机号）----
    users = load_json('users.json', 'users')
    user_rows = []
    for u in users:
        phone_seq += 1
        # 历史密码（无论明文还是哈希）一律重置为 "用户名+123"，强制首次登录后修改
        plain = f"{u['username']}123"
        user_rows.append((
            u['username'],
            generate_password_hash(plain),
            u.get('role', 'user'),
            sanitize_email(u.get('email', ''), u['username']),
            sanitize_phone(u.get('contact', ''), phone_seq),
            u.get('register_time') or '2024-01-01 08:00:00',
            u.get('last_login_time') or None,
        ))

    # ---- 班级 ----
    classes = load_json('classes.json', 'classes')
    class_rows = [
        (c['class_id'], c.get('college', ''), c['class_name'],
         c.get('advisor', ''), c.get('created_at', ''))
        for c in classes
    ]

    # ---- 学生 + 课程（courses 拆表）----
    students = load_json('students.json', 'students')
    student_rows = [
        (s['student_id'], s.get('name', ''), s.get('gender', '男'),
         s.get('college', ''), s.get('class', ''))
        for s in students
    ]
    course_rows = []
    for s in students:
        for c in s.get('courses', []):
            course_rows.append((
                s['student_id'], c.get('course_name', ''), c.get('credit', 0),
                c.get('score', 0), c.get('course_type', ''), c.get('course_nature', ''),
                c.get('semester', ''), c.get('gpa', 0), c.get('credit_gpa', 0),
            ))

    # ---- 宿舍 ----
    dormitories = load_json('dormitories.json', 'dormitories')
    dorm_rows = [
        (d['student_id'], d.get('name', ''), d.get('college', ''), d.get('class', ''),
         d.get('building', ''), d.get('room', ''), d.get('bed', ''))
        for d in dormitories
    ]

    # ---- 竞赛（脱敏手机号）----
    competitions = load_json('competitions.json', 'competitions')
    comp_rows = []
    for c in competitions:
        phone_seq += 1
        comp_rows.append((
            c['certificate_id'], c.get('competition_name', ''), c.get('award_level', ''),
            c.get('award_year', ''), c.get('college', ''), c.get('class', ''),
            c.get('award_students', ''), sanitize_phone(c.get('contact', ''), phone_seq),
            c.get('advisor', ''),
        ))

    # ---- 公告（丢弃测试数据）----
    notices = [n for n in load_json('notices.json', 'notices') if n['notice_id'] not in DROP_NOTICE_IDS]
    notice_rows = [
        (n['notice_id'], n.get('title', ''), n.get('content', ''),
         n.get('priority', 'normal'), n.get('publisher', ''), n.get('publish_time'))
        for n in notices
    ]

    # ---- 按业务键去重（历史数据存在重复记录）----
    user_rows = dedupe(user_rows, 0)
    class_rows = dedupe(class_rows, 0)
    student_rows = dedupe(student_rows, 0)
    dorm_rows = dedupe(dorm_rows, 0)
    comp_rows = dedupe(comp_rows, 0)
    notice_rows = dedupe(notice_rows, 0)
    kept_students = {r[0] for r in student_rows}
    course_rows = [r for r in course_rows if r[0] in kept_students]

    # ---- 写入 MySQL（幂等：先清空业务表）----
    db.clear_tables(['users', 'classes', 'students', 'courses', 'dormitories',
                     'competitions', 'notices', 'warning_status'])

    db.execute_many(
        'INSERT INTO users (username, password_hash, role, email, contact, register_time, last_login_time) '
        'VALUES (%s, %s, %s, %s, %s, %s, %s)', user_rows)
    db.execute_many(
        'INSERT INTO classes (class_id, college, class_name, advisor, created_at) '
        'VALUES (%s, %s, %s, %s, %s)', class_rows)
    db.execute_many(
        'INSERT INTO students (student_id, name, gender, college, class_name) '
        'VALUES (%s, %s, %s, %s, %s)', student_rows)
    db.execute_many(
        'INSERT INTO courses (student_id, course_name, credit, score, course_type, '
        'course_nature, semester, gpa, credit_gpa) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)',
        course_rows)
    db.execute_many(
        'INSERT INTO dormitories (student_id, name, college, class_name, building, room, bed) '
        'VALUES (%s, %s, %s, %s, %s, %s, %s)', dorm_rows)
    db.execute_many(
        'INSERT INTO competitions (certificate_id, competition_name, award_level, award_year, '
        "college, class_name, award_students, contact, advisor) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)",
        comp_rows)
    db.execute_many(
        'INSERT INTO notices (notice_id, title, content, priority, publisher, publish_time) '
        'VALUES (%s, %s, %s, %s, %s, %s)', notice_rows)

    print(f'迁移完成：用户 {len(user_rows)}，班级 {len(class_rows)}，学生 {len(student_rows)}，'
          f'课程 {len(course_rows)}，宿舍 {len(dorm_rows)}，竞赛 {len(comp_rows)}，公告 {len(notice_rows)}')
    return user_rows, class_rows, student_rows, course_rows, dorm_rows, comp_rows, notice_rows


def sql_literal(value):
    if value is None:
        return 'NULL'
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, datetime):
        return "'" + value.strftime('%Y-%m-%d %H:%M:%S') + "'"
    return "'" + str(value).replace('\\', '\\\\').replace("'", "''").replace('\n', '\\n') + "'"


def export_seed(data):
    """生成 database/seed.sql（脱敏后的演示数据，随仓库发布）。"""
    user_rows, class_rows, student_rows, course_rows, dorm_rows, comp_rows, notice_rows = data
    lines = [
        '-- 演示数据（全部为虚构脱敏数据，不含任何真实个人信息）',
        '-- 用法：mysql -u root -p < database/schema.sql && mysql -u root -p student_management < database/seed.sql',
        'USE student_management;',
        '',
    ]
    sections = [
        ('users', 'username, password_hash, role, email, contact, register_time, last_login_time', user_rows),
        ('classes', 'class_id, college, class_name, advisor, created_at', class_rows),
        ('students', 'student_id, name, gender, college, class_name', student_rows),
        ('courses', 'student_id, course_name, credit, score, course_type, course_nature, semester, gpa, credit_gpa', course_rows),
        ('dormitories', 'student_id, name, college, class_name, building, room, bed', dorm_rows),
        ('competitions', 'certificate_id, competition_name, award_level, award_year, college, class_name, award_students, contact, advisor', comp_rows),
        ('notices', 'notice_id, title, content, priority, publisher, publish_time', notice_rows),
    ]
    for table, cols, rows in sections:
        lines.append(f'-- 表：{table}（{len(rows)} 条）')
        for row in rows:
            values = ', '.join(sql_literal(v) for v in row)
            lines.append(f'INSERT INTO {table} ({cols}) VALUES ({values});')
        lines.append('')

    os.makedirs(DATABASE_DIR, exist_ok=True)
    seed_path = os.path.join(DATABASE_DIR, 'seed.sql')
    with open(seed_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    print(f'seed.sql 已生成：{seed_path}')


def verify_seed(seed_path):
    """seed.sql 导出后的脱敏校验（合规闸门，发现违规立即报错退出）。

    检查项：
      1. 邮箱：仅允许 example.com 保留域；
      2. 手机号：仅允许 13800000000~13899999999 占位号段；
      3. 不得出现 18 位居民身份证号码格式（普通词汇"身份证"不算，仅拦截真实号码）；
      4. users 表仅允许 password_hash 列存哈希（pbkdf2:/scrypt: 前缀），禁止明文密码。
    """
    import re

    with open(seed_path, 'r', encoding='utf-8') as f:
        content = f.read()

    violations = []

    # 1. 真实域名邮箱（example.com 保留域除外）
    for match in re.finditer(r'[\w.+-]+@([\w-]+\.)+[a-zA-Z]{2,}', content):
        domain = match.group(0).split('@')[1].lower()
        if domain != 'example.com':
            violations.append(f'违规邮箱: {match.group(0)}')

    # 2. 占位号段以外的手机号
    for match in re.finditer(r'1[3-9]\d{9}', content):
        phone = match.group(0)
        if not re.fullmatch(r'138\d{8}', phone):
            violations.append(f'非占位号段手机号: {phone}')

    # 3. 居民身份证号码（18 位格式，末位可为 X；要求词边界避免误伤哈希串）
    for match in re.finditer(r'(?<![0-9a-zA-Z])\d{17}[0-9Xx](?![0-9a-zA-Z])', content):
        violations.append(f'疑似身份证号码: {match.group(0)}')

    # 4. 明文密码（users 行的 password_hash 值必须为哈希前缀开头）
    for match in re.finditer(r"INSERT INTO users \([^)]*\) VALUES \('[^']*',\s*'([^']*)'", content):
        pwd = match.group(1)
        if not pwd.startswith(('pbkdf2:', 'scrypt:')):
            violations.append('users 表检出非哈希密码（疑似明文）')

    if violations:
        print('seed.sql 脱敏校验未通过：', file=sys.stderr)
        for v in violations:
            print(f'  - {v}', file=sys.stderr)
        sys.exit(1)

    print(f'seed.sql 脱敏校验通过：无真实邮箱/占位号段以外手机号/身份证/明文密码（{seed_path}）')


if __name__ == '__main__':
    ensure_schema()
    result = migrate()
    if '--export-seed' in sys.argv:
        seed_path = os.path.join(DATABASE_DIR, 'seed.sql')
        export_seed(result)
        verify_seed(seed_path)
