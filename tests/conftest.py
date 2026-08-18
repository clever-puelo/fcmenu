"""Fixtures compartidas para la suite de tests de `migration`.

Cada test corre contra una base SQLite en memoria nueva (no contra
fcmenu.mdb ni contra PostgreSQL) para que la suite sea rápida y no
dependa de la base real. Los datos reales de fcmenu.mdb se usaron durante
el desarrollo para confirmar el schema (ver data_layer_progress.md en
memoria de sesión), no se leen acá.
"""

from __future__ import annotations

import pytest
from sqlalchemy import create_engine
from sqlalchemy.orm import Session, sessionmaker

from migration.models import Base


@pytest.fixture()
def db() -> Session:
    engine = create_engine("sqlite:///:memory:", future=True)
    Base.metadata.create_all(engine)
    SessionLocal = sessionmaker(bind=engine, future=True)
    session = SessionLocal()
    try:
        yield session
    finally:
        session.close()
