from flask import Blueprint, render_template, redirect, url_for, flash, request
from flask_login import login_user, logout_user, login_required, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime, timedelta
import random
import string

auth = Blueprint('auth', __name__)

# Import from models (Member 1 ki file)
from models import db, User, PasswordReset
# Import forms (is file ke upar)
from forms import LoginForm, ForgotPasswordForm, ResetPasswordForm


@auth.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('home'))
    
    form = LoginForm()
    if form.validate_on_submit():
        user = User.query.filter_by(username=form.username.data).first()
        
        if user and check_password_hash(user.password, form.password.data):
            if user.role == form.role.data:
                login_user(user)
                
                if user.role == 'admin':
                    return redirect(url_for('admin.admin_dashboard'))
                elif user.role == 'teacher':
                    return redirect(url_for('teacher.teacher_dashboard'))
                else:
                    return redirect(url_for('student.student_dashboard'))
            else:
                flash('Invalid role selected.', 'danger')
        else:
            flash('Invalid username or password.', 'danger')
    
    return render_template('login.html', form=form)


@auth.route('/logout')
@login_required
def logout():
    logout_user()
    flash('Logged out successfully.', 'info')
    return redirect(url_for('auth.login'))


@auth.route('/forgot_password', methods=['GET', 'POST'])
def forgot_password():
    form = ForgotPasswordForm()
    if form.validate_on_submit():
        user = User.query.filter_by(username=form.username.data).first()
        
        if user:
            reset_code = ''.join(random.choices(string.digits, k=6))
            
            reset_request = PasswordReset(
                user_id=user.id,
                reset_code=reset_code,
                expires_at=datetime.utcnow() + timedelta(minutes=30)
            )
            db.session.add(reset_request)
            db.session.commit()
            
            # Demo purpose - show code
            flash(f'Reset code: {reset_code} (Demo only)', 'info')
            
            return redirect(url_for('auth.reset_password'))
        else:
            flash('Username not found.', 'danger')
    
    return render_template('forgot_password.html', form=form)


@auth.route('/reset_password', methods=['GET', 'POST'])
def reset_password():
    form = ResetPasswordForm()
    if form.validate_on_submit():
        reset_request = PasswordReset.query.filter_by(
            reset_code=form.reset_code.data,
            is_used=False
        ).first()
        
        if reset_request and reset_request.expires_at > datetime.utcnow():
            user = User.query.get(reset_request.user_id)
            user.password = generate_password_hash(form.new_password.data)
            reset_request.is_used = True
            db.session.commit()
            
            flash('Password reset successful! Login now.', 'success')
            return redirect(url_for('auth.login'))
        else:
            flash('Invalid or expired code.', 'danger')
    
    return render_template('reset_password.html', form=form)