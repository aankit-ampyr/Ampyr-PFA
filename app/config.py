from pydantic_settings import BaseSettings


class Settings(BaseSettings):
    database_url: str = "postgresql://parthenon:parthenon_dev@localhost:5432/parthenon"
    app_env: str = "development"

    model_config = {"env_file": ".env"}


settings = Settings()
