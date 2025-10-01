import typing

from sqlalchemy import create_engine, Column, String, Boolean
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker

Base = declarative_base()


class Admin(Base):
    __tablename__ = 'admins'

    id = Column(String, primary_key=True)  # Telegram ID as primary key
    chat_id = Column(String, nullable=False)


class ApprovedUser(Base):
    __tablename__ = 'approved_users'

    id = Column(String, primary_key=True)  # Telegram ID as primary key


class AdminAuthDB:
    def __init__(self, db_url: str = "sqlite:///admin_auth.db", create_tables: bool = True):
        """
        Initialize the AdminAuthDB with SQLite database.

        Args:
            db_url: SQLAlchemy database URL (e.g., "sqlite:///path/to/database.db")
            create_tables: Whether to create tables if they don't exist
        """
        self.engine = create_engine(db_url)
        self.SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=self.engine)

        if create_tables:
            Base.metadata.create_all(bind=self.engine)

    def _get_session(self):
        """Create and return a new database session."""
        return self.SessionLocal()

    def get_is_admin_info(self, id: str) -> typing.Optional[Admin]:
        """
        Check if a user ID is in the admins table.

        Args:
            id: Telegram user ID to check

        Returns:
            None: if user is not an admin
            Admin: admin info if he is an admin
        """
        session = self._get_session()
        try:
            admin = session.query(Admin).filter(Admin.id == id).first()
            return admin
        finally:
            session.close()

    def get_admins_list(self) -> list[Admin]:
        session = self._get_session()
        try:
            admin = session.query(Admin).all()
            return list(admin)
        finally:
            session.close()


    def set_admin(self, id: str, chat_id: str) -> None:
        """
        Add a user ID to the admins table if not already present.

        Args:
            id: Telegram user ID to add as admin
        """
        if self.get_is_admin_info(id) is not None:
            return  # Already an admin, do nothing

        session = self._get_session()
        try:
            admin = Admin(id=id, chat_id=chat_id)
            session.merge(admin)
            session.commit()
        except Exception:
            session.rollback()
            raise
        finally:
            session.close()

    def check_is_approved(self, id: str) -> bool:
        """
        Check if a user ID is in the approved_users table.

        Args:
            id: Telegram user ID to check

        Returns:
            bool: True if the user is approved, False otherwise
        """
        session = self._get_session()
        try:
            approved_user = session.query(ApprovedUser).filter(ApprovedUser.id == id).first()
            return approved_user is not None
        finally:
            session.close()

    def set_approved(self, id: str) -> None:
        """
        Add a user ID to the approved_users table if not already present.

        Args:
            id: Telegram user ID to approve
        """
        if self.check_is_approved(id):
            return  # Already approved, do nothing

        session = self._get_session()
        try:
            approved_user = ApprovedUser(id=id)
            session.add(approved_user)
            session.commit()
        except Exception:
            session.rollback()
            raise
        finally:
            session.close()