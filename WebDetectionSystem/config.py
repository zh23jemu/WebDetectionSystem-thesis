import os

BASE_DIR = os.path.abspath(os.path.dirname(__file__))


class Config:
    # 安全密钥
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'wood-detection-secret-key-2026'

    # 数据库配置 (使用SQLite)
    SQLALCHEMY_DATABASE_URI = 'sqlite:///' + os.path.join(BASE_DIR, 'wood_system.db')
    SQLALCHEMY_TRACK_MODIFICATIONS = False

    # 路径配置
    UPLOAD_FOLDER = os.path.join(BASE_DIR, 'static', 'uploads')
    RESULT_FOLDER = os.path.join(BASE_DIR, 'static', 'results')

    # YOLO模型路径 (确保best.pt放在项目根目录，或者修改此处)
    MODEL_PATH = os.path.join(BASE_DIR, 'best.pt')