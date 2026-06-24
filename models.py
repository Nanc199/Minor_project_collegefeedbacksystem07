from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from datetime import datetime

db = SQLAlchemy()

class User(db.Model, UserMixin):
    
    """User model for students, teachers, and admins"""
    __tablename__ = 'users'
    
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password = db.Column(db.String(200), nullable=False)
    plain_password = db.Column(db.String(200))
    email = db.Column(db.String(120), nullable=False)
    role = db.Column(db.String(20), nullable=False)  # student, teacher, admin
    sub_role = db.Column(db.String(20), nullable=True)  # superadmin / branchadmin
    name = db.Column(db.String(100), nullable=False)
    branch = db.Column(db.String(50))  # CSE-AI&DS, Civil, Electrical
    semester = db.Column(db.Integer, default=1)
    imported_file_id = db.Column(db.Integer, db.ForeignKey('uploaded_file.id'))
    academic_year = db.Column(db.String(20))
    roll_number = db.Column(db.String(20))
    subject = db.Column(db.String(100))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    
    # Relationships
    feedbacks_given = db.relationship('Feedback',
                                      foreign_keys='Feedback.student_id',
                                      backref='student_user',
                                      lazy=True)
    feedbacks_received = db.relationship('Feedback',
                                         foreign_keys='Feedback.teacher_id',
                                         backref='teacher_user',
                                         lazy=True)
    uploaded_file = db.relationship('UploadedFile',
                                    foreign_keys='UploadedFile.uploaded_by',
                                     backref='uploader',
                                    lazy=True)
    
    def get_id(self):
        return self.id
        
# models.py
class UploadedFile(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(255), nullable=False)
    file_type = db.Column(db.String(50), nullable=False)  # student / teacher / admin
    uploaded_by = db.Column(db.Integer, db.ForeignKey('users.id'))
    content = db.Column(db.Text, nullable=True)
    uploaded_at = db.Column(db.DateTime, default=datetime.utcnow)
class Feedback(db.Model):

    __tablename__ = 'feedbacks'

    id = db.Column(db.Integer, primary_key=True)

    student_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    teacher_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)

    subject = db.Column(db.String(100))

    q1 = db.Column(db.Integer, nullable=False)
    q2 = db.Column(db.Integer, nullable=False)
    q3 = db.Column(db.Integer, nullable=False)
    q4 = db.Column(db.Integer, nullable=False)
    q5 = db.Column(db.Integer, nullable=False)
    q6 = db.Column(db.Integer, nullable=False)
    q7 = db.Column(db.Integer, nullable=False)
    q8 = db.Column(db.Integer, nullable=False)
    q9 = db.Column(db.Integer, nullable=False)
    q10 = db.Column(db.Integer, nullable=False)
    q11 = db.Column(db.Integer, nullable=False)
    q12 = db.Column(db.Integer, nullable=False)

    total = db.Column(db.Integer, nullable=False)

    semester = db.Column(db.Integer, nullable=False)
    academic_year = db.Column(db.String(20), nullable=False)

    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    hidden_from_admin_previous = db.Column(db.Boolean, default=False)
    hidden_from_teacher_old = db.Column(db.Boolean, default=False)
    __table_args__ = (
    db.UniqueConstraint('student_id', 'teacher_id', 'subject', 'semester', name='unique_feedback'),
)

    student = db.relationship('User', foreign_keys=[student_id], backref='student_feedbacks')
    teacher = db.relationship('User', foreign_keys=[teacher_id], backref='teacher_feedbacks')
class TeacherSubject(db.Model):
    id = db.Column(db.Integer, primary_key=True)

    teacher_id = db.Column(db.Integer, db.ForeignKey('users.id'))

    subject = db.Column(db.String(100))
    subject_code = db.Column(db.String(50))   # <-- YE ADD KARO

    branch = db.Column(db.String(50))
    semester = db.Column(db.Integer)

    # relationship to User (teacher)
    teacher = db.relationship('User', foreign_keys=[teacher_id])
class PasswordReset(db.Model):
    """Password reset tokens"""
    __tablename__ = 'password_resets'
    
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    reset_code = db.Column(db.String(10), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    expires_at = db.Column(db.DateTime, nullable=False)
    is_used = db.Column(db.Boolean, default=False) 





class FeedbackWindow(db.Model):
    id = db.Column(db.Integer, primary_key=True)

    start_time = db.Column(db.DateTime, nullable=False)
    end_time = db.Column(db.DateTime, nullable=False)

    is_active = db.Column(db.Boolean, default=True)


class AdminSettings(db.Model):
    __tablename__ = 'admin_settings'

    id = db.Column(db.Integer, primary_key=True)
    feedback_start = db.Column(db.DateTime)
    feedback_end = db.Column(db.DateTime)
    allow_feedback = db.Column(db.Boolean, default=True)
