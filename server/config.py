
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    SECRET_KEY: str = "change_me_to_a_long_random_secret"
    JWT_EXPIRE_HOURS: int = 12
    JWT_ALGORITHM: str = "HS256"

    DATABASE_URL: str = "sqlite:///./zhu_stock.db"

    CREATOR_ADMIN_EMAIL: str = "admin@zhustock.local"
    CREATOR_ADMIN_USERNAME: str = "admin"
    CREATOR_ADMIN_PASSWORD: str = "Admin123456!"

    ABUSE_NOTIFY_EMAIL: str = ""

    TRIAL_DAYS: int = 30
    MAX_DEVICE_PER_USER: int = 1

    EMAIL_DEV_MODE: bool = True
    SMTP_HOST: str = ""
    SMTP_PORT: int = 587
    SMTP_USER: str = ""
    SMTP_PASSWORD: str = ""
    SMTP_FROM: str = ""

    model_config = SettingsConfigDict(env_file=".env", env_file_encoding="utf-8", extra="ignore")


settings = Settings()
