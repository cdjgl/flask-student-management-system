-- ============================================================
-- 学生管理系统 MySQL 建库建表脚本（utf8mb4）
-- 用法：mysql -u root -p < database/schema.sql
-- 或直接启动应用（db.init_db 会自动建表）
-- ============================================================

CREATE DATABASE IF NOT EXISTS student_management
    DEFAULT CHARSET utf8mb4 COLLATE utf8mb4_unicode_ci;

USE student_management;

-- 用户表
CREATE TABLE IF NOT EXISTS users (
    id INT AUTO_INCREMENT PRIMARY KEY,
    username VARCHAR(50) NOT NULL UNIQUE,
    password_hash VARCHAR(255) NOT NULL,
    role VARCHAR(20) NOT NULL DEFAULT 'user',
    email VARCHAR(100) NOT NULL UNIQUE,
    contact VARCHAR(20) NOT NULL DEFAULT '',
    register_time DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
    last_login_time DATETIME NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- 班级表
CREATE TABLE IF NOT EXISTS classes (
    id INT AUTO_INCREMENT PRIMARY KEY,
    class_id VARCHAR(20) NOT NULL UNIQUE,
    college VARCHAR(50) NOT NULL DEFAULT '',
    class_name VARCHAR(50) NOT NULL UNIQUE,
    advisor VARCHAR(20) NOT NULL DEFAULT '',
    created_at VARCHAR(10) NOT NULL DEFAULT ''
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- 学生表
CREATE TABLE IF NOT EXISTS students (
    id INT AUTO_INCREMENT PRIMARY KEY,
    student_id VARCHAR(20) NOT NULL UNIQUE,
    name VARCHAR(20) NOT NULL DEFAULT '',
    gender VARCHAR(2) NOT NULL DEFAULT '男',
    college VARCHAR(50) NOT NULL DEFAULT '',
    class_name VARCHAR(50) NOT NULL DEFAULT '',
    INDEX idx_students_college (college),
    INDEX idx_students_class (class_name)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- 课程表（学生 1:N 课程）
CREATE TABLE IF NOT EXISTS courses (
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
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- 宿舍表
CREATE TABLE IF NOT EXISTS dormitories (
    id INT AUTO_INCREMENT PRIMARY KEY,
    student_id VARCHAR(20) NOT NULL UNIQUE,
    name VARCHAR(20) NOT NULL DEFAULT '',
    college VARCHAR(50) NOT NULL DEFAULT '',
    class_name VARCHAR(50) NOT NULL DEFAULT '',
    building VARCHAR(20) NOT NULL DEFAULT '',
    room VARCHAR(10) NOT NULL DEFAULT '',
    bed VARCHAR(10) NOT NULL DEFAULT ''
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- 竞赛获奖表
CREATE TABLE IF NOT EXISTS competitions (
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
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- 公告表
CREATE TABLE IF NOT EXISTS notices (
    id INT AUTO_INCREMENT PRIMARY KEY,
    notice_id VARCHAR(20) NOT NULL UNIQUE,
    title VARCHAR(100) NOT NULL DEFAULT '',
    content TEXT NOT NULL,
    priority VARCHAR(10) NOT NULL DEFAULT 'normal',
    publisher VARCHAR(50) NOT NULL DEFAULT '',
    publish_time DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- 预警处理状态表（预警本身由 courses 实时计算，此表仅保存处理进度）
CREATE TABLE IF NOT EXISTS warning_status (
    id INT AUTO_INCREMENT PRIMARY KEY,
    student_id VARCHAR(20) NOT NULL,
    warning_type VARCHAR(20) NOT NULL,
    warning_key VARCHAR(60) NOT NULL DEFAULT '',
    status VARCHAR(20) NOT NULL DEFAULT 'pending',
    updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
        ON UPDATE CURRENT_TIMESTAMP,
    UNIQUE KEY uk_warning (student_id, warning_type, warning_key)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
