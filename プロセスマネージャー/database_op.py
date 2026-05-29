from sqlalchemy import create_engine, ForeignKey, select, exists, NullPool, or_, text #MetaData, Table, String, Integer, Column,
from sqlalchemy.orm import Session, DeclarativeBase, mapped_column, Mapped#, relationship
from typing import Optional, Tuple, Dict
import sqlalchemy.exc as DBEXC, uuid, sys, os
from sys_control import get_frozen_stat
import configparser

frozen = TEST = get_frozen_stat()

def get_path_to_db() -> str:
    config = configparser.ConfigParser()
    if frozen:
        thisfile = os.path.dirname(sys.executable)
        inipath = os.path.join(thisfile, 'config.ini')
    else:
        return 'test.db'
    try:
        config.read(inipath, encoding='utf-8')
    except:
        raise Exception('ini設定ファイルが見つかりません、インターネットの接続と、ファイルの所在を確認してください。')
    
    path = config['database']['path']

    return path


dbname = get_path_to_db()
engine = create_engine(fr'sqlite:///{dbname}', echo=TEST, poolclass=NullPool)

class Base(DeclarativeBase):
    pass

class User(Base):
    __tablename__ = 'users'
    namecode: Mapped[str] = mapped_column(primary_key= True, unique= True)
    uid: Mapped[str] = mapped_column(unique= True)
    name: Mapped[Optional[str]]

    def __repr__(self) -> str:
        return f'User(namecode={self.namecode!r}, uid={self.uid!r}, name={self.name!r})'
    
    def new(self, namecode: str, name: str = None, uid: str = None):
        self.namecode = namecode
        if not uid:
            self.uid = str(uuid.uuid4())
        else:
            self.uid = uid
        if name:
            self.name = name
        else:
            self.name = '-'

        
    
class Memo(Base):
    __tablename__ = 'memos'
    namecode: Mapped[str] = mapped_column(ForeignKey('users.namecode'), primary_key= True)
    memoid: Mapped[str] = mapped_column(primary_key= True, unique= True)
    memotitle: Mapped[str]
    memocontent: Mapped[str]

    def __repr__(self) -> str:
        return f'Memo(namecode={self.namecode!r}, memoid={self.memoid!r}, memotitle={self.memotitle!r}, memocontent={self.memocontent!r})'
    
    def new_memo(self, namecode: str, memotitle: str, memocontent: str) -> str:
        from datetime import datetime
        memo_id = str(uuid.uuid3(uuid.NAMESPACE_OID, memotitle + memocontent + datetime.now().strftime('%Y%m%d%H%M%S')))[:21]
        self.namecode = namecode
        self.memoid = memo_id
        self.memocontent = memocontent
        self.memotitle = memotitle
        return memo_id


    
class Setting(Base):
    __tablename__ = 'settings'
    namecode: Mapped[str] = mapped_column(ForeignKey('users.namecode'), primary_key= True, unique= True)
    settingcluster: Mapped[str] = mapped_column(nullable= False)

    def __repr__(self) -> str:
        return f'Setting(namecode={self.namecode!r}, settingcluster={self.settingcluster!r})'
    
    def create_setting(self, namecode: str, settingcluster: str):
        self.namecode = namecode
        self.settingcluster = settingcluster


class Log(Base):
    __tablename__ = 'logs'
    namecode: Mapped[str] = mapped_column(ForeignKey('users.namecode'), primary_key= True)
    searchkey: Mapped[str] = mapped_column(primary_key= True)
    searchmode: Mapped[str] = mapped_column(primary_key= True)
    count: Mapped[int] = mapped_column(default= 0)

    def __repr__(self) -> str:
        return f'Log(namecode={self.namecode!r}, searchkey={self.searchkey!r}, searchmode={self.searchmode!r}, count={self.count!r})'


class FilterMemory(Base):
    __tablename__ = 'filter_memories'

    namecode:     Mapped[str]           = mapped_column(ForeignKey('users.namecode'), primary_key=True)
    filter_id:    Mapped[str]           = mapped_column(primary_key=True)
    filter_name:  Mapped[str]
    query_mode:   Mapped[str]
    db_scope:     Mapped[Optional[str]]
    filter_model: Mapped[str]

    def __repr__(self) -> str:
        return f'FilterMemory(namecode={self.namecode!r}, filter_id={self.filter_id!r}, filter_name={self.filter_name!r})'


# Create filter_memories table automatically if it does not exist yet
FilterMemory.__table__.create(engine, checkfirst=True)


class VendorSchedComment(Base):
    __tablename__ = 'vendor_sched_comments'
    schedule_id:  Mapped[str]            = mapped_column(primary_key=True)  # Oracle hex schedule_id (or negative str timestamp as fallback)
    lot_id:       Mapped[str]
    vendor_code:  Mapped[str]
    db_drawing:   Mapped[Optional[str]]
    product_name: Mapped[Optional[str]]
    cust_pn:      Mapped[Optional[str]]
    quantity:     Mapped[Optional[int]]
    comment1:     Mapped[Optional[str]]
    comment2:     Mapped[Optional[str]]
    comment3:     Mapped[Optional[str]]

    def __repr__(self) -> str:
        return f'VendorSchedComment(schedule_id={self.schedule_id!r}, lot_id={self.lot_id!r})'


def _migrate_vendor_sched_comments_if_needed() -> None:
    """Migrate old (lot_id, vendor_code) PK schema → new schedule_id PK schema."""
    with engine.connect() as conn:
        result = conn.execute(text("PRAGMA table_info(vendor_sched_comments)"))
        cols = {row[1] for row in result}
        if not cols or 'schedule_id' in cols:
            return  # table absent or already migrated
        conn.execute(text('''
            CREATE TABLE vendor_sched_comments_new (
                schedule_id  INTEGER PRIMARY KEY,
                lot_id       TEXT,
                vendor_code  TEXT,
                db_drawing   TEXT,
                product_name TEXT,
                cust_pn      TEXT,
                quantity     INTEGER,
                comment1     TEXT,
                comment2     TEXT,
                comment3     TEXT
            )
        '''))
        # Assign negative ROWID as placeholder schedule_id for migrated rows.
        # Negative values never clash with Oracle's positive sequence IDs.
        conn.execute(text('''
            INSERT INTO vendor_sched_comments_new
                (schedule_id, lot_id, vendor_code, db_drawing, product_name,
                 cust_pn, quantity, comment1, comment2, comment3)
            SELECT -ROWID, lot_id, vendor_code, NULL, NULL,
                   cust_pn, quantity, comment1, comment2, comment3
            FROM vendor_sched_comments
        '''))
        conn.execute(text('DROP TABLE vendor_sched_comments'))
        conn.execute(text(
            'ALTER TABLE vendor_sched_comments_new RENAME TO vendor_sched_comments'))
        conn.commit()

def _migrate_vendor_sched_schedule_id_to_text_if_needed() -> None:
    """Migrate schedule_id column from INTEGER to TEXT to support Oracle hex IDs."""
    with engine.connect() as conn:
        result = conn.execute(text("PRAGMA table_info(vendor_sched_comments)"))
        rows = list(result)
        if not rows:
            return  # table doesn't exist yet
        col_types = {row[1]: row[2] for row in rows}
        if col_types.get('schedule_id', '').upper() != 'INTEGER':
            return  # already TEXT or absent
        conn.execute(text('''
            CREATE TABLE vendor_sched_comments_new (
                schedule_id  TEXT PRIMARY KEY,
                lot_id       TEXT,
                vendor_code  TEXT,
                db_drawing   TEXT,
                product_name TEXT,
                cust_pn      TEXT,
                quantity     INTEGER,
                comment1     TEXT,
                comment2     TEXT,
                comment3     TEXT
            )
        '''))
        conn.execute(text('''
            INSERT INTO vendor_sched_comments_new
                (schedule_id, lot_id, vendor_code, db_drawing, product_name,
                 cust_pn, quantity, comment1, comment2, comment3)
            SELECT CAST(schedule_id AS TEXT), lot_id, vendor_code, db_drawing, product_name,
                   cust_pn, quantity, comment1, comment2, comment3
            FROM vendor_sched_comments
        '''))
        conn.execute(text('DROP TABLE vendor_sched_comments'))
        conn.execute(text(
            'ALTER TABLE vendor_sched_comments_new RENAME TO vendor_sched_comments'))
        conn.commit()


def _add_vendor_sched_comment_columns_if_needed() -> None:
    """Add db_drawing / product_name columns if they are missing from an already-migrated table."""
    with engine.connect() as conn:
        result = conn.execute(text("PRAGMA table_info(vendor_sched_comments)"))
        cols = {row[1] for row in result}
        if not cols:
            return  # table doesn't exist yet — will be created fresh below
        for col, col_type in (('db_drawing', 'TEXT'), ('product_name', 'TEXT')):
            if col not in cols:
                conn.execute(text(f'ALTER TABLE vendor_sched_comments ADD COLUMN {col} {col_type}'))
        conn.commit()


_migrate_vendor_sched_comments_if_needed()
_migrate_vendor_sched_schedule_id_to_text_if_needed()
_add_vendor_sched_comment_columns_if_needed()
VendorSchedComment.__table__.create(engine, checkfirst=True)


def create_table(init: bool = False) -> bool:
    print()
    print('Intialise Tables based on predefined meta data')
    confirm = ''
    if not init:
        confirm = input('This action will erase current existing database(if exist),Input "CONFIRM" in upper case to proceed')
    if confirm == 'CONFIRM' or init:
        try:
            Base.metadata.create_all(engine)
        except Exception as e:
            print(f'Error happened during the initialization of tabel:\n{e}\n{str(e.__cause__)}\n{e.__format__}')
            input('>>')
            return False, e
        print('Table initializaion successfully finished.')
        input('>>')
        return True, ''
    else:
        input('Failed to confirm, please restart the process if you wish to proceed.')
        return False, 'F'


def new_user_single(namecode: str, setting_default:str, name: str = None, uid: str = None) -> Tuple[str, str]:
    with Session(engine) as session:
        newuser_i = User()
        newuser_i.new(namecode, uid= uid)
        newuser_setting_i = Setting()
        newuser_setting_i.create_setting(namecode, setting_default)
        newuser_setting_list = [newuser_i, newuser_setting_i]
        try:
            session.add_all(newuser_setting_list)
            session.commit()
        except DBEXC.IntegrityError as e:
            if 'NOT NULL constraint failed: users.uid' in str(e):
                new_uuid = str(uuid.uuid4())
                result = new_user_single(namecode, name, new_uuid)
                if result[0] == 's':
                    return ('s', namecode)
                else:
                    return ('e', e)
            return ('e', e)
        return ('s', namecode)
    
def check_user_exists(namecode: str) -> bool:
    with Session(engine) as session:
        return session.query(exists().where(User.namecode == namecode)).scalar()

def get_user_name(namecode: str) -> Optional[str]:
    with Session(engine) as session:
        stmt = select(User.name).where(User.namecode == namecode)
        return session.scalars(stmt).first()

def update_user_name(namecode: str, name: str) -> Tuple[str, str]:
    with Session(engine) as session:
        stmt = select(User).where(User.namecode == namecode)
        try:
            target = session.scalars(stmt).one()
        except Exception as e:
            return ('e', str(e))
        target.name = name
        try:
            session.commit()
        except Exception as e:
            return ('e', str(e))
    return ('s', namecode)
    
def new_memo(namecode: str, memotitle: str, memocontent: str) -> Tuple[str, str]:
    with Session(engine) as session:
        newmemo_i = Memo()
        memoid = newmemo_i.new_memo(namecode, memotitle, memocontent)
        try:
            session.add(newmemo_i)
            session.commit()
        except Exception as e:
            return ('e', e)
        return ('s', memoid)
    
def update_memo(namecode: str, memoid: str, memocontent: str) -> Tuple[str, str]:
    with Session(engine) as session:
        stmt = select(Memo).where(Memo.namecode == namecode, Memo.memoid == memoid)
        target = session.scalars(stmt).one()
        target.memocontent = memocontent
        try:
            session.commit()
        except Exception as e:
            return ('e', e)
    return ('s', memoid)

def delete_memo(namecode: str, memoid: str) -> Tuple[str, str]:
    with Session(engine) as session:
        stmt = select(Memo).where(Memo.namecode == namecode, Memo.memoid == memoid)
        target = session.scalars(stmt).one()
        session.delete(target)
        try:
            session.commit()
        except Exception as e:
            return ('e', e)
    return ('s', memoid)

def save_setting(namecode: str, settingcluster: str) -> Tuple[str, str]:
    with Session(engine) as session:
        stmt = select(Setting).where(Setting.namecode == namecode)
        target = session.scalars(stmt).one()
        target.settingcluster = settingcluster
        try:
            session.commit()
        except Exception as e:
            return ('e', e)
    return ('s', namecode)

def init_user_data(namecode: str) -> Dict[str, tuple]:
    user_dict = {}
    with Session(engine) as session:
        stmt_user = select(User.name).where(User.namecode == namecode)
        user_dict['user_name'] = session.scalars(stmt_user).first()
        stmt_memo = select(Memo).where(Memo.namecode == namecode)
        result_memo = tuple(session.scalars(stmt_memo).all())
        user_dict['memo'] = result_memo
        stmt_setting = select(Setting).where(Setting.namecode == namecode)
        result_setting = session.scalars(stmt_setting).one()
        user_dict['setting'] = result_setting

    return user_dict


def save_filter_memory(namecode: str, filter_name: str, query_mode: str, db_scope: Optional[str], filter_model_json: str) -> Tuple[str, str]:
    with Session(engine) as session:
        new_filter = FilterMemory()
        new_filter.namecode = namecode
        new_filter.filter_id = str(uuid.uuid4())
        new_filter.filter_name = filter_name
        new_filter.query_mode = query_mode
        new_filter.db_scope = db_scope
        new_filter.filter_model = filter_model_json
        try:
            session.add(new_filter)
            session.commit()
        except Exception as e:
            return ('e', str(e))
        return ('s', new_filter.filter_id)


def get_filter_memories(namecode: str, query_mode: Optional[str] = None, db_scope: Optional[str] = None) -> list:
    with Session(engine) as session:
        stmt = select(FilterMemory).where(FilterMemory.namecode == namecode)
        if query_mode is not None:
            stmt = stmt.where(FilterMemory.query_mode == query_mode)
            if query_mode == 'w' and db_scope is not None:
                stmt = stmt.where(
                    or_(FilterMemory.db_scope == None, FilterMemory.db_scope == db_scope)
                )
        result = session.scalars(stmt).all()
        return list(result)


def update_filter_memory(namecode: str, filter_id: str, **kwargs) -> Tuple[str, str]:
    with Session(engine) as session:
        stmt = select(FilterMemory).where(FilterMemory.namecode == namecode, FilterMemory.filter_id == filter_id)
        try:
            target = session.scalars(stmt).one()
        except Exception as e:
            return ('e', str(e))
        for key, val in kwargs.items():
            if hasattr(target, key):
                setattr(target, key, val)
        try:
            session.commit()
        except Exception as e:
            return ('e', str(e))
        return ('s', filter_id)


def delete_filter_memory(namecode: str, filter_id: str) -> Tuple[str, str]:
    with Session(engine) as session:
        stmt = select(FilterMemory).where(FilterMemory.namecode == namecode, FilterMemory.filter_id == filter_id)
        try:
            target = session.scalars(stmt).one()
        except Exception as e:
            return ('e', str(e))
        session.delete(target)
        try:
            session.commit()
        except Exception as e:
            return ('e', str(e))
        return ('s', filter_id)


def upsert_vendor_sched_comment(
        schedule_id: str, lot_id: str, vendor_code: str, **kwargs) -> Tuple[str, str]:
    with Session(engine) as session:
        target = session.get(VendorSchedComment, schedule_id)
        if target is None:
            target = VendorSchedComment(
                schedule_id=schedule_id, lot_id=lot_id, vendor_code=vendor_code,
                cust_pn=None, quantity=1,
                comment1=None, comment2=None, comment3=None)
            session.add(target)
        for key, val in kwargs.items():
            if hasattr(target, key):
                setattr(target, key, val)
        try:
            session.commit()
        except Exception as e:
            return ('e', str(e))
        return ('s', str(schedule_id))


def get_vendor_sched_comments(vendor_code: str, schedule_ids: list) -> dict:
    """Returns {schedule_id: VendorSchedComment} for the given vendor + schedule_ids."""
    if not schedule_ids:
        return {}
    with Session(engine) as session:
        stmt = select(VendorSchedComment).where(
            VendorSchedComment.vendor_code == vendor_code,
            VendorSchedComment.schedule_id.in_(schedule_ids)
        )
        return {r.schedule_id: r for r in session.scalars(stmt).all()}


def get_vendor_sched_comments_by_lot(vendor_code: str, lot_ids: list) -> dict:
    """Returns {lot_id: VendorSchedComment} for the given vendor + lot_ids.
    Used as a fallback when a row was saved with a negative placeholder schedule_id."""
    if not lot_ids:
        return {}
    with Session(engine) as session:
        stmt = select(VendorSchedComment).where(
            VendorSchedComment.vendor_code == vendor_code,
            VendorSchedComment.lot_id.in_(lot_ids)
        )
        result: dict = {}
        for r in session.scalars(stmt).all():
            if r.lot_id not in result:   # keep first found per lot_id
                result[r.lot_id] = r
        return result


def delete_vendor_sched_comment(schedule_id: str) -> None:
    """Delete the SQLite comment row for a given schedule_id (no-op if not found)."""
    with Session(engine) as session:
        target = session.get(VendorSchedComment, schedule_id)
        if target is not None:
            session.delete(target)
            session.commit()
