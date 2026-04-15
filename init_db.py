import pymysql
from config import DB_CONFIG

def init_database():
    conn = pymysql.connect(**DB_CONFIG)
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS tasks (
            id INT PRIMARY KEY AUTO_INCREMENT,
            name VARCHAR(200) NOT NULL,
            category VARCHAR(50) NOT NULL,
            priority VARCHAR(10) DEFAULT '中',
            status VARCHAR(20) DEFAULT '未开始',
            plan_start DATETIME,
            plan_end DATETIME,
            progress_note TEXT,
            pause_reason TEXT,
            sort_order INT DEFAULT 0,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            updated_at DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
        )
    """)

    conn.commit()
    cursor.close()
    conn.close()
    print("数据库初始化完成")

if __name__ == '__main__':
    init_database()
