from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        case_sensitive=False,
    )

    gemini_api_key: str
    gemini_model: str = "gemini-2.0-flash"
    data_dir: str = "./data"
    api_host: str = "0.0.0.0"
    api_port: int = 8002


settings = Settings()