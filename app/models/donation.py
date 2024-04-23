from sqlalchemy import Column, ForeignKey, Integer, Text

from app.core.db import Base, Charity


class Donation(Base, Charity):
    user_id = Column(Integer, ForeignKey('user.id'))
    comment = Column(Text)
