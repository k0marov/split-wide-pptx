from sqlalchemy import create_engine, Column, String, Boolean
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker

Base = declarative_base()


class Admin(Base):
    __tablename__ = 'admins'

    id = Column(String, primary_key=True)  # Telegram ID as primary key


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

    def check_is_admin(self, id: str) -> bool:
        """
        Check if a user ID is in the admins table.

        Args:
            id: Telegram user ID to check

        Returns:
            bool: True if the user is an admin, False otherwise
        """
        session = self._get_session()
        try:
            admin = session.query(Admin).filter(Admin.id == id).first()
            return admin is not None
        finally:
            session.close()

    def add_admin(self, id: str) -> None:
        """
        Add a user ID to the admins table if not already present.

        Args:
            id: Telegram user ID to add as admin
        """
        if self.check_is_admin(id):
            return  # Already an admin, do nothing

        session = self._get_session()
        try:
            admin = Admin(id=id)
            session.add(admin)
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