from extensions import db
from flask_login import UserMixin
from datetime import datetime
from sqlalchemy.dialects.mysql import MEDIUMTEXT


class User(db.Model, UserMixin):
    __tablename__ = "users"

    id           = db.Column(db.Integer, primary_key=True)
    name         = db.Column(db.String(120), nullable=False)
    username     = db.Column(db.String(80), unique=True, nullable=True)  # nullable for Google users
    password     = db.Column(db.String(256), nullable=True)              # nullable for Google users
    email        = db.Column(db.String(255), unique=True, nullable=True) # ← ADD
    google_id    = db.Column(db.String(128), unique=True, nullable=True) # ← ADD
    is_admin     = db.Column(db.Boolean, default=False, nullable=False)
    active       = db.Column("is_active", db.Boolean, default=True, nullable=False)
    date_created = db.Column(
        db.String(50),
        default=lambda: datetime.now().strftime("%Y-%m-%d")
    )

    @property
    def is_active(self):
        return self.active

    @is_active.setter
    def is_active(self, value):
        self.active = value

    def get_id(self):
        return str(self.id)

    def __repr__(self):
        return f"<User {self.username} admin={self.is_admin} active={self.active}>"


class TosRecord(db.Model):
    __tablename__ = "tos_records"

    id           = db.Column(db.Integer, primary_key=True)
    user_id      = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False)
    title        = db.Column(db.String(255), nullable=False)
    topics_json  = db.Column(MEDIUMTEXT, default="[]")
    quizzes_json = db.Column(MEDIUMTEXT, default="[]")
    total_items  = db.Column(db.Integer, default=0)
    date_created = db.Column(db.String(50))
    subject_type = db.Column(db.String(20), default='nonlab')
    is_derived   = db.Column(db.Boolean, default=False, nullable=False)
    parent_id    = db.Column(db.Integer, db.ForeignKey("tos_records.id"), nullable=True)

    def __repr__(self):
        return f"<TosRecord #{self.id} '{self.title}' derived={self.is_derived}>"