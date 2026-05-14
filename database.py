from sqlalchemy import create_engine, Column, String, Text
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
import json
import pandas as pd
import shutil
import datetime


# 🔴 CHANGE THESE
DB_USER = "postgres"
DB_PASSWORD = "your_password"
DB_HOST = "localhost"
DB_PORT = "5432"
DB_NAME = "instrument_db"

DATABASE_URL = f"postgresql://{DB_USER}:{DB_PASSWORD}@{DB_HOST}:{DB_PORT}/{DB_NAME}"

engine = create_engine(DATABASE_URL)
Session = sessionmaker(bind=engine)
Base = declarative_base()


# ── TABLE ─────────────────────────────

class Instrument(Base):
    __tablename__ = "instruments"

    TagNo = Column(String, primary_key=True)
    data = Column(Text)


# ── INIT ─────────────────────────────

def init_db():
    Base.metadata.create_all(engine)


# ── INSERT / UPDATE ───────────────────

def upsert_row(row: dict):
    session = Session()

    tag = row.get("TagNo")

    if not tag:
        session.close()
        return

    existing = session.query(Instrument).filter_by(TagNo=tag).first()

    if existing:
         print(f"🔄 Updating existing tag: {tag}")
        existing.data = json.dumps(row)
    else:
         print(f"🆕 Inserting new tag: {tag}")
        new_row = Instrument(
            TagNo=tag,
            data=json.dumps(row)
        )
        session.add(new_row)

    session.commit()
    session.close()


# ── EXPORT TO EXCEL ───────────────────

def export_to_excel(path):
    session = Session()

    rows = session.query(Instrument).all()

    data = [json.loads(r.data) for r in rows]

    df = pd.DataFrame(data)

    df.to_excel(path, index=False)

    session.close()

def backup_database():
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = f"backup_{timestamp}.sql"

    import os
    command = f'pg_dump -U postgres -d instrument_db -f {backup_file}'
    
    os.system(command)

    print(f"✅ Backup created: {backup_file}")