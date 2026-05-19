import os
from flask import Blueprint, render_template, request, redirect, url_for, flash, session
from werkzeug.security import check_password_hash, generate_password_hash
from flask_login import login_user, logout_user, login_required, current_user
from authlib.integrations.flask_client import OAuth
from models import User
from extensions import db
from datetime import datetime

auth_bp = Blueprint("auth", __name__)

# ── Google OAuth setup ────────────────────────────────────────────────
oauth = OAuth()

def init_oauth(app):
    oauth.init_app(app)
    oauth.register(
        name="google",
        client_id=os.environ.get("GOOGLE_CLIENT_ID"),
        client_secret=os.environ.get("GOOGLE_CLIENT_SECRET"),
        server_metadata_url="https://accounts.google.com/.well-known/openid-configuration",
        client_kwargs={"scope": "openid email profile"},
    )


# ── Standard login ────────────────────────────────────────────────────
@auth_bp.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        if current_user.is_admin:
            return redirect(url_for("admin.index"))
        return redirect(url_for("dashboard.index"))

    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")

        user = User.query.filter_by(username=username).first()

        if not user or not user.password or not check_password_hash(user.password, password):
            flash("Invalid username or password.", "error")
            return redirect(url_for("auth.login"))

        if not user.active:
            flash("Your account has been deactivated. Please contact an administrator.", "error")
            return redirect(url_for("auth.login"))

        login_user(user)
        if user.is_admin:
            return redirect(url_for("admin.index"))
        return redirect(url_for("dashboard.index"))

    return render_template("login.html")


# ── Standard signup ───────────────────────────────────────────────────
@auth_bp.route("/signup", methods=["GET", "POST"])
def signup():
    if current_user.is_authenticated:
        if current_user.is_admin:
            return redirect(url_for("admin.index"))
        return redirect(url_for("dashboard.index"))

    if request.method == "POST":
        name     = request.form.get("name", "").strip()
        username = request.form.get("username", "").strip().lower()
        password = request.form.get("password", "")

        if not name or not username or not password:
            flash("All fields are required.", "error")
            return redirect(url_for("auth.signup"))

        if len(password) < 6:
            flash("Password must be at least 6 characters.", "error")
            return redirect(url_for("auth.signup"))

        if User.query.filter_by(username=username).first():
            flash("Username already taken.", "error")
            return redirect(url_for("auth.signup"))

        user = User(
            name=name,
            username=username,
            password=generate_password_hash(password),
            is_admin=False,
            active=True,
            date_created=datetime.now().strftime("%Y-%m-%d"),
        )
        db.session.add(user)
        db.session.commit()

        flash("Account created! You can now sign in.", "success")
        return redirect(url_for("auth.login"))

    return render_template("signup.html")


# ── Google OAuth — redirect to Google ────────────────────────────────
@auth_bp.route("/auth/google")
def google_login():
    redirect_uri = url_for("auth.google_callback", _external=True)
    return oauth.google.authorize_redirect(redirect_uri)


# ── Google OAuth — callback ───────────────────────────────────────────
@auth_bp.route("/auth/google/callback")
def google_callback():
    try:
        token = oauth.google.authorize_access_token()
    except Exception:
        flash("Google sign-in failed. Please try again.", "error")
        return redirect(url_for("auth.login"))

    userinfo = token.get("userinfo")
    if not userinfo:
        flash("Could not retrieve your Google account information.", "error")
        return redirect(url_for("auth.login"))

    google_id = userinfo.get("sub")
    email     = userinfo.get("email")
    name      = userinfo.get("name") or email.split("@")[0]

    # ── Find existing user by Google ID or email ──
    user = User.query.filter_by(google_id=google_id).first()

    if not user:
        user = User.query.filter_by(email=email).first()
        if user:
            # Link existing account to Google
            user.google_id = google_id
            db.session.commit()
        else:
            # Create new account via Google
            base_username = email.split("@")[0].lower().replace(".", "")
            username = base_username
            counter = 1
            while User.query.filter_by(username=username).first():
                username = f"{base_username}{counter}"
                counter += 1

            user = User(
                name=name,
                username=username,
                email=email,
                google_id=google_id,
                password=None,   # No password for Google users
                is_admin=False,
                active=True,
                date_created=datetime.now().strftime("%Y-%m-%d"),
            )
            db.session.add(user)
            db.session.commit()

    if not user.active:
        flash("Your account has been deactivated.", "error")
        return redirect(url_for("auth.login"))

    login_user(user)
    if user.is_admin:
        return redirect(url_for("admin.index"))
    return redirect(url_for("dashboard.index"))


# ── Logout ────────────────────────────────────────────────────────────
@auth_bp.route("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("auth.login"))