import db
import models

if __name__ == '__main__':
    db.Base.metadata.create_all(db.engine)
    admin = models.User('admin', 'password', 'admin@example.com')
    db.session.add(admin)
    db.session.commit()
    db.session.close()
