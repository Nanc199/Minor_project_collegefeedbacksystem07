# Minor_project_collegefeedbacksystem07
# 🎓 College Feedback System

A web-based feedback management system that allows students to submit feedback about faculty, faculty to view their feedback, and admins to manage the entire process — built using **Python (Flask)** and **SQLite**.

## 📖 Introduction

Collecting and analyzing student feedback about faculty is often a manual, paper-based, and time-consuming process in colleges. This project digitizes that entire workflow into a simple, role-based web application.

## ❓ Problem It Solves

Traditional feedback collection (paper forms) is slow, hard to analyze, and lacks transparency. There's no easy way for admins to generate reports or for faculty to get structured feedback in one place.

## ✅ Solution

This system provides:
- A dedicated portal for **Students** to give feedback on faculty/subjects.
- A dashboard for **Faculty** to view feedback received (anonymously).
- Full control for **Admin** to manage students, faculty, and feedback records, and generate reports.

## ✨ Features

- 🔐 **Role-Based Login** – Separate access for Student, Faculty, and Admin.
- 📝 **Feedback Submission** – Students can rate and comment on faculty/subjects.
- 📊 **Faculty Dashboard** – View aggregated feedback and ratings.
- 🛠️ **Admin Panel** – Add/remove students & faculty, view/manage all feedback.
- 🗄️ **Lightweight Database** – Uses SQLite, no complex setup required.
- 💻 **Simple Web Interface** – Built with Flask + HTML templates.

## 🧭 Navigation / How to Use

1. Open the app in your browser.
2. **Login** based on your role — Student, Faculty, or Admin.
3. **Student:** Select subject/faculty → submit rating & comments.
4. **Faculty:** View dashboard → see feedback received.
5. **Admin:** Manage users → view all feedback → generate reports.

## 🛠️ Tech Stack

| Layer      | Technology              |
|------------|--------------------------|
| Backend    | Python, Flask            |
| Database   | SQLite                   |
| Frontend   | HTML, CSS, Jinja2 (Flask templates) |

## 📋 Requirements

- Python 3.x
- pip packages: `flask`, `flask-sqlalchemy` (if used)

## 🚀 Installation / Run Locally

```bash
git clone https://github.com/Nanc199/Minor_project_collegefeedbacksystem07.git
cd Minor_project_collegefeedbacksystem07
pip install -r requirements.txt
python app.py
```

Then open `http://127.0.0.1:5000/` in your browser.

## 📁 Project Structure
```
Minor_project_collegefeedbacksystem07/
├── app.py # Main Flask application
├── database.db # SQLite database
├── templates/ # HTML pages (Jinja2)
├── static/ # CSS, JS, images
└── requirements.txt # Python dependencies
```

## 🎥 Demo
<img width="1920" height="1020" alt="Screenshot 2026-05-18 180937" src="https://github.com/user-attachments/assets/1e4e4f31-3567-4b8f-9c06-1c08a64537af" />
<img width="1809" height="821" alt="Screenshot 2026-05-16 195639" src="https://github.com/user-attachments/assets/c1e97d52-45d6-4578-abdb-2926597ad230" />
<img width="1920" height="914" alt="Screenshot 2026-05-16 195706" src="https://github.com/user-attachments/assets/4f6623ab-fc30-4de8-bc26-4dde7bde25ca" />
<img width="1882" height="775" alt="Screenshot 2026-09-17 194601" src="https://github.com/user-attachments/assets/1ce4909d-250c-4273-855e-6a18094920d7" />
<img width="1911" height="775" alt="Screenshot 2026-09-17 194617" src="https://github.com/user-attachments/assets/43f341be-89dd-499c-a625-ba2e7575e6bc" />
<img width="1889" height="819" alt="Screenshot 2026-09-17 194719" src="https://github.com/user-attachments/assets/c7082706-5c3b-4fd3-8567-7f1ba957d552" />
<img width="1920" height="802" alt="Screenshot 2026-09-17 194748" src="https://github.com/user-attachments/assets/1795bc77-007b-4485-861d-60e1206b8b0f" />
<img width="1920" height="1020" alt="Screenshot 2026-05-17 101506" src="https://github.com/user-attachments/assets/289b0170-6676-425f-a484-75fe69d13f65" />
<img width="1248" height="693" alt="Screenshot 2026-05-15 112224" src="https://github.com/user-attachments/assets/d27988d9-f25b-45b7-a249-8e885726becf" />
<img width="1920" height="1020" alt="Screenshot 2026-05-17 101506" src="https://github.com/user-attachments/assets/833e67ee-1c0c-4452-b883-8ba73e45fbce" />
<img width="1879" height="748" alt="Screenshot 2026-05-08 163804" src="https://github.com/user-attachments/assets/487e6936-906a-438b-882c-1e8d12017195" />
<img width="1270" height="534" alt="Screenshot 2026-05-08 224928" src="https://github.com/user-attachments/assets/c7d5525d-9752-4c13-97a4-03de9e6c353e" />













