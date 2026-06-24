from flask_wtf import FlaskForm
from wtforms import StringField, PasswordField, SubmitField, SelectField, TextAreaField, IntegerField, FileField
from wtforms.validators import DataRequired, Email, Length, EqualTo, Optional

class LoginForm(FlaskForm):
    """Login form for all users"""
    username = StringField('Username / Student ID', 
                          validators=[DataRequired()])
    password = PasswordField('Password', 
                            validators=[DataRequired()])
    role = SelectField('Login As', 
                       choices=[('student', 'Student'), 
                               ('teacher', 'Teacher'), 
                               ('admin', 'Admin')],
                       validators=[DataRequired()])
    submit = SubmitField('Login')


class FeedbackForm(FlaskForm):
    """Feedback submission form"""
    rating = SelectField('Rating', 
                        choices=[(5, '5 - Excellent'), 
                                (4, '4 - Very Good'), 
                                (3, '3 - Good'), 
                                (2, '2 - Poor'), 
                                (1, '1 - Very Poor')],
                        validators=[DataRequired()])
    comment = TextAreaField('Comments (Optional)', 
                           validators=[Length(max=500)])
    submit = SubmitField('Submit Feedback')


class ForgotPasswordForm(FlaskForm):
    """Forgot password - Enter username"""
    username = StringField('Username / Student ID',
                         validators=[DataRequired()])
    submit = SubmitField('Send Reset Code')


class ResetPasswordForm(FlaskForm):
    """Reset password with code"""
    reset_code = StringField('Reset Code', 
                            validators=[DataRequired(), Length(min=6, max=6)])
    new_password = PasswordField('New Password',
                                  validators=[DataRequired(), Length(min=6)])
    confirm_password = PasswordField('Confirm New Password',
                                     validators=[DataRequired(), 
                                                EqualTo('new_password')])
    submit = SubmitField('Reset Password')


class ExcelUploadForm(FlaskForm):
    """Admin: Upload Excel file"""
    file = FileField('Upload Excel File', validators=[DataRequired()])
    submit = SubmitField('Upload & Import')