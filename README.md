# Flask Student Management System

基于 **Flask + MySQL** 的学生管理系统，提供学生、班级、宿舍、竞赛获奖、成绩预警、公告等信息的统一管理与数据可视化。

> 本项目仅供学习交流使用。**内置演示数据全部为虚构脱敏数据，不含任何真实个人信息。**

## 功能特性

- **学生管理**：学生信息的增删改查、课程成绩维护、多条件筛选、Excel 导出
- **班级管理**：班级信息维护、班级改名全库联动（学生/宿舍/竞赛同步更新）
- **宿舍管理**：学生住宿分配与查询
- **竞赛管理**：竞赛获奖记录登记与统计
- **成绩预警**：挂科预警、低绩点预警实时计算，处理状态跟踪，报表导出
- **公告管理**：系统公告发布/编辑/删除（管理员）
- **统计分析**：性别/成绩/学院等多维统计与图表可视化（Chart.js）
- **数据备份**：一键备份全部数据表为 ZIP、下载、恢复、删除
- **用户认证**：注册/登录（Werkzeug 密码哈希）、个人中心改密

## 技术栈

| 层 | 技术 |
|---|---|
| 后端 | Python 3 / Flask 2.3 |
| 数据库 | MySQL 8（utf8mb4） |
| 数据访问 | PyMySQL + DBUtils 连接池 |
| 前端 | Bootstrap 5 / Bootstrap Icons / Chart.js |
| 导出 | openpyxl（Excel .xlsx） |

## 快速开始

### 1. 环境要求

- Python 3.9+
- MySQL 8.x（本机或远程）

### 2. 安装依赖

```bash
pip install -r requirements.txt
```

### 3. 初始化数据库

方式一：执行建库脚本 + 演示数据

```bash
mysql -u root -p < database/schema.sql
mysql -u root -p student_management < database/seed.sql
```

方式二：直接启动应用，首次启动会自动建表并创建默认管理员（无演示数据）。

### 4. 配置环境变量

复制 `.env.example` 为 `.env`，填写数据库连接等信息：

```ini
DB_HOST=127.0.0.1
DB_PORT=3306
DB_USER=root
DB_PASSWORD=your-password
DB_NAME=student_management
```

> `.env` 已被 `.gitignore` 排除，不会被提交到仓库。

### 5. 启动

Windows 用户可双击 `start.bat` 一键启动（自动检查 Python、.env 配置与依赖，缺失时自动安装）。

macOS / Linux 用户：

```bash
python app.py
```

访问 http://127.0.0.1:5000

**默认管理员账号：`admin / admin123`（首次登录后请立即在个人中心修改密码）**

### 6.（可选）从旧版 JSON 数据迁移

如你持有旧版 JSON 存储的数据，可运行迁移脚本（自动脱敏：密码哈希化、邮箱替换为 example.com、手机号归一化为占位号段）：

```bash
python scripts/import_json_to_mysql.py --export-seed
```

## 项目结构

```
学生管理系统/
├── app.py                       # Flask 主应用（全部路由）
├── config.py                    # 配置（从 .env 读取）
├── db.py                        # MySQL 数据访问层（连接池）
├── requirements.txt             # Python 依赖
├── database/
│   ├── schema.sql               # 建库建表脚本
│   └── seed.sql                 # 脱敏演示数据
├── scripts/
│   └── import_json_to_mysql.py  # 旧版 JSON 数据迁移
├── templates/                   # Jinja2 页面模板
├── static/                      # CSS / JS 静态资源
├── .env.example                 # 环境变量模板
├── LICENSE                      # MIT License
└── THIRD_PARTY_NOTICES.md       # 第三方开源组件声明
```

## 开源合规说明

### 许可证

本项目基于 [MIT License](LICENSE) 开源。

> MIT License 仅覆盖本项目作者编写的源代码，不包含通过 CDN 加载的前端库（Bootstrap、Bootstrap Icons、Chart.js 等），第三方组件遵循各自协议，详见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)。

### 第三方组件

本项目使用了以下优秀的开源组件，各组件版权归其原作者所有，详见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)：

| 组件 | 许可证 |
|---|---|
| Flask / Werkzeug | BSD-3-Clause |
| openpyxl | MIT |
| PyMySQL | MIT |
| DBUtils | PSF-2.0 |
| Bootstrap / Bootstrap Icons | MIT |
| Chart.js | MIT |

### 免责声明

1. 本项目仅用于编程学习、教学演示，不适合直接投入生产环境使用；
2. 本系统所有演示数据均为虚构，若部署使用，不得录入任何自然人真实个人信息，并遵守《中华人民共和国个人信息保护法》、网络安全相关法律法规；
3. 仓库内置数据均为虚构演示数据，如有雷同纯属巧合；
4. 字体说明：本项目 CSS 仅引用系统字体名称（如微软雅黑），不附带、不分发任何字体文件；
5. 数据备份提示：本项目自带数据库备份导出功能，若自行部署使用，导出的数据属于使用者责任，不得非法收集、泄露他人个人信息；
6. 使用者因使用本项目而产生的任何后果由使用者自行承担。

## 贡献

欢迎提交 Issue 与 Pull Request。提交前请确保：

1. 不引入任何真实个人数据；
2. 代码风格与现有实现保持一致；
3. 新依赖须为宽松许可证（MIT/BSD/Apache 等）。

## 作者

cdjgl
