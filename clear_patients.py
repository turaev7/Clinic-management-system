from app import app, db, Patient
with app.app_context():
    deleted = db.session.query(Patient).delete()
    db.session.commit()
    print(f"Removed {deleted} patients.")
