from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    SECRET_KEY: str = "change_me_to_a_long_random_secret"
    JWT_EXPIRE_HOURS: int = 12
    JWT_ALGORITHM: str = "HS256"

    DATABASE_URL: str = "sqlite:///./zhu_stock.db"
    # Large daily market history must not consume the membership database quota.
    CACHE_DATABASE_URL: str = "sqlite:///./zhu_stock_cache.db"

    CREATOR_ADMIN_EMAIL: str = "admin@zhustock.local"
    CREATOR_ADMIN_USERNAME: str = "admin"
    CREATOR_ADMIN_PASSWORD: str = "Admin123456!"

    ABUSE_NOTIFY_EMAIL: str = ""

    TRIAL_DAYS: int = 30
    MAX_DEVICE_PER_USER: int = 1

    EMAIL_DEV_MODE: bool = False

    # 舊 SMTP 先保留，不再使用也沒關係
    SMTP_HOST: str = ""
    SMTP_PORT: int = 587
    SMTP_USER: str = ""
    SMTP_PASSWORD: str = ""
    SMTP_FROM: str = ""

    # Brevo
    BREVO_API_KEY: str = ""
    BREVO_FROM_EMAIL: str = ""

    # 網頁版自動分析：台北時間週一至週五 16:05（收盤資料穩定後）
    AUTO_ANALYSIS_HOUR: int = 16
    AUTO_ANALYSIS_MINUTE: int = 5
    CRON_SECRET: str = ""

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore"
    )


settings = Settings()
