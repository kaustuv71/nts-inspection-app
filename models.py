"""SQLAlchemy models for NTS Inspection App."""
import os, json
from datetime import datetime
from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()

class Inspection(db.Model):
    __tablename__ = "inspections"

    id = db.Column(db.String(16), primary_key=True)
    data = db.Column(db.JSON, nullable=False, default=dict)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    def to_dict(self):
        return {**self.data, "id": self.id}

    @classmethod
    def list_all(cls):
        rows = cls.query.order_by(cls.updated_at.desc()).all()
        result = []
        for r in rows:
            d = r.data
            result.append({
                "id": r.id,
                "client": d.get("client", ""),
                "product": d.get("product_name", ""),
                "inspection_date": d.get("inspection_date", ""),
                "status": d.get("status", "draft"),
                "updated": str(r.updated_at) if r.updated_at else "",
            })
        return result

class Photo(db.Model):
    __tablename__ = "photos"

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    inspection_id = db.Column(db.String(16), db.ForeignKey("inspections.id"), nullable=False)
    sku_idx = db.Column(db.Integer, nullable=False)
    filename = db.Column(db.String(100), nullable=False)
    caption = db.Column(db.Text, default="")
    data = db.Column(db.Text, nullable=False)  # base64-encoded JPEG
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def to_dict(self):
        return {
            "path": f"photos/{self.filename}",
            "caption": self.caption,
            "data": self.data[:50] + "..." if len(self.data) > 50 else self.data,
        }
