from __future__ import annotations

import os

from sqlalchemy import create_engine
from sqlalchemy.orm import Session, sessionmaker

from .models import Base

# Connection string del motor destino. Se lee de la variable de entorno
# DATABASE_URL (ver .env.example) para no commitear credenciales; si no
# está seteada, cae a SQLite local — sirve para desarrollo rápido y es lo
# que usan los tests (ver tests/conftest.py, que usa su propio engine en
# memoria y no pasa por acá).
#
# Ejemplo real de este proyecto (base de desarrollo local, ver
# migration/alembic/ para el DDL versionado):
#   DATABASE_URL=postgresql+psycopg2://fcmenu_app:***@localhost:5432/fcmenu_dev
DEFAULT_SQLITE_URL = "sqlite:///fcmenu.sqlite3"
DATABASE_URL = os.environ.get("DATABASE_URL", DEFAULT_SQLITE_URL)


# `pool_size`/`max_overflow` sólo aplican a motores con `QueuePool` real
# (Postgres) — SQLite usa su propio pool y no acepta estos kwargs. El
# default de SQLAlchemy (5 + 10 overflow = 15 conexiones) se quedaba
# corto con el `MainMenuWindow` nuevo: cada pantalla abre su propia
# sesión (`get_session()`) y nunca la cierra hasta que el usuario la
# cierra a mano — confirmado con un smoke test real que, al abrir ~18
# pantallas seguidas sin cerrar ninguna, la ventana Nº19 se quedaba
# colgada esperando una conexión libre del pool, sin ningún error
# visible. 30+20 alcanza cómodamente para tener abierta cada pantalla
# de la app al mismo tiempo.
_kwargs_pool = {} if DATABASE_URL.startswith("sqlite") else {"pool_size": 30, "max_overflow": 20}
engine = create_engine(DATABASE_URL, future=True, **_kwargs_pool)
SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False, future=True)


def init_db() -> None:
    Base.metadata.create_all(bind=engine)


def get_session() -> Session:
    return SessionLocal()
