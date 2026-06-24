import os
import pandas as pd
from werkzeug.security import generate_password_hash

def allowed_file(filename):
    """Check if file extension is allowed"""
    ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def import_users_from_excel(file_path, db):
    """
    Read Excel file and import users to database
    Excel columns: Username, Password, Role, Name, Email, Branch, Semester, Academic_Year, Subject, Roll_Number
    """
    from models import User
    
    try:
        df = pd.read_excel(file_path)
        
        required_cols = ['Username', 'Password', 'Role', 'Name', 'Email']
        for col in required_cols:
            if col not in df.columns:
                return False, f"Missing column: {col}"
        
        imported_count = 0
        skipped_count = 0
        
        for index, row in df.iterrows():
            existing_user = User.query.filter_by(username=row['Username']).first()
            
            if existing_user:
                skipped_count += 1
                continue
            
            new_user = User(
                username=row['Username'],
                password=generate_password_hash(str(row['Password'])),
                role=row['Role'].lower(),
                name=row['Name'],
                email=row['Email'],
                branch=row.get('Branch', ''),
                semester=int(row.get('Semester', 1)),
                academic_year=row.get('Academic_Year', ''),
                roll_number=row.get('Roll_Number', ''),
                subject=row.get('Subject', '')
            )
            
            db.session.add(new_user)
            imported_count += 1
        
        db.session.commit()
        return True, f"Imported {imported_count}, Skipped {skipped_count}"
    
    except Exception as e:
        return False, str(e)


def increment_all_semesters(db):
    """Increment semester for all students"""
    from models import User
    
    students = User.query.filter_by(role='student').all()
    count = 0
    for student in students:
        if student.semester < 8:
            student.semester += 1
            count += 1
    db.session.commit()
    return count