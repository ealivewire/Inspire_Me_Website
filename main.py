# PROFESSIONAL PROJECT: Inspire Me! Website

# OBJECTIVE: To implement a website which automates the retrieval and distribution of inspiration quotes
#            (can be on a periodic basis if scheduled via web host).

# Import necessary library(ies):
from data import app, db, recognition, SENDER_EMAIL_GMAIL, SENDER_HOST, SENDER_PASSWORD_GMAIL, SENDER_PORT, WEB_LOADING_TIME_ALLOWANCE
from data import Categories, InspirationDataSources, InspirationalQuotes, Subscribers, Users
from data import AddOrEditSubscriberForm, AdminLoginForm, AdminUpdateForm, ContactForm
from datetime import datetime
from dotenv import load_dotenv
import email_validator
from flask import Flask, abort, render_template, redirect, url_for
from flask_bootstrap import Bootstrap5
from flask_login import UserMixin, login_user, LoginManager, current_user, logout_user
from flask_sqlalchemy import SQLAlchemy
from functools import wraps  # Used in 'admin_only" decorator function
from flask_wtf import FlaskForm
import html2text
import os
import random
from selenium import webdriver
from selenium.webdriver.common.by import By
import smtplib
from sqlalchemy import Integer, String, Boolean, ForeignKey, func
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, relationship
import time
import traceback
from typing import List
from werkzeug.security import check_password_hash
from wtforms import BooleanField, EmailField, PasswordField, StringField, SubmitField, TextAreaField
from wtforms.validators import InputRequired, Length, Email
import wx
import wx.lib.agw.pybusyinfo as PBI


# Define variable to be used for showing user dialog and message boxes:
dlg = wx.App()

# Initialize the Flask app. object:
app = Flask(__name__)


# Create needed class "Base":
class Base(DeclarativeBase):
  pass

# NOTE: Additional configurations are launched via the "run_app" function defined below.


# Configure the Flask login manager:
login_manager = LoginManager()
login_manager.init_app(app)


# Implement a user loader callback (to facilitate loading current user into session):
@login_manager.user_loader
def load_user(user_id):
    return retrieve_from_database("load_user", user_id=user_id)


# Implement a decorator function to ensure that only someone who knows the admin password can access the
# "administrative update" functionality of the website:
def admin_only(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        # If not the admin then return abort with 403 error:
        if not (current_user.is_authenticated and current_user.id == 1):
            return abort(403)

        # At this point, user is the admin, so proceed with allowing access to the route:
        return f(*args, **kwargs)

    return decorated_function


# CONFIGURE ROUTES FOR WEB PAGES (LISTED IN HIERARCHICAL ORDER STARTING WITH HOME PAGE, THEN ALPHABETICALLY):
# ***********************************************************************************************************
# Configure route for home page:
@app.route('/',methods=["GET", "POST"])
def home():
    try:
        # Instantiate an instance of the "AddOrEditSubscriberForm" class:
        form = AddOrEditSubscriberForm()

        # If form-level validation has passed, perform additional processing:
        if form.validate_on_submit():
            # Initialize variable used to provide feedback to user:
            result = ""

            # Check if the e-mail address supplied via the form exists in the database:
            email_in_db = retrieve_from_database("get_subscriber_by_email", email=form.txt_email.data)
            if email_in_db == {}:
                result = "An error has occurred.  Subscription cannot be completed at this time."
            elif email_in_db != None:
                result = f"E-mail address {form.txt_email.data} is already subscribed to our mailing list for inspirational quotes."
            else:
                if not update_database("add_subscriber", item_to_process=[], form=form):
                    result = "An error has occurred.  Subscription cannot be completed at this time."
                else:
                    result = f"Thank you! E-mail address {form.txt_email.data} has been added to our mailing list for inspirational quotes."

            # Render the results via the home page:
            return render_template("index.html", result=result, logged_in=current_user.is_authenticated, recognition_web_template=recognition["web_template"])

        # Go to the home page:
        return render_template("index.html", form=form, logged_in=current_user.is_authenticated, recognition_web_template=recognition["web_template"])

    except:
        # Log error into system log file:
        update_system_log("route: '/'", traceback.format_exc())

        # Go to the web page which displays error details to the user:
        return render_template("error.html", activity="route: '/'", details=traceback.format_exc())


# Configure route for "About" web page:
@app.route('/about')
def about():
    try:
        # Go to the "About" page:
        return render_template("about.html", recognition_web_template=recognition["web_template"])

    except:
        # Log error into system log:
        update_system_log("route: '/about'", traceback.format_exc())

        # Go to the web page which displays error details to the user:
        return render_template("error.html", activity="route: '/about'", details=traceback.format_exc())


# Configure route for "add subscriber" web page:
@app.route('/add_subscriber',methods=["GET", "POST"])
@admin_only
def add_subscriber():
    try:
        # Instantiate an instance of the "AddOrEditSubscriberForm" class:
        form = AddOrEditSubscriberForm()

        # If form-level validation has passed, perform additional processing:
        if form.validate_on_submit():
            # Initialize variable used to provide feedback to user:
            result = ""

            # Check if the e-mail address supplied via the form exists in the database:
            email_in_db = retrieve_from_database("get_subscriber_by_email", email=form.txt_email.data)
            if email_in_db == {}:
                result = "An error has occurred.  Subscription cannot be completed at this time."
            elif email_in_db != None:
                result = f"E-mail address {form.txt_email.data} is already subscribed to the mailing list for inspirational quotes."
            else:
                if not update_database("add_subscriber", item_to_process=[], form=form):
                    result = "An error has occurred.  Subscription cannot be completed at this time."
                else:
                    result = f"E-mail address {form.txt_email.data} has been added to the mailing list for inspirational quotes."

            # Go to the web page to render the results:
            return render_template("db_update_result.html", trans_type="Add Subscriber", result=result,
                                   recognition_web_template=recognition["web_template"])

        # Go to the "add subscriber" web page:
        return render_template("add_subscriber.html", form=form, recognition_web_template=recognition["web_template"])

    except:  # An error has occurred.
        # Log error into system log file:
        update_system_log("route: '/add_subscriber'", traceback.format_exc())

        # Go to the web page which displays error details to the user:
        return render_template("error.html", activity="route: '/add_subscriber'", details=traceback.format_exc())


# Configure route for "Administrative Update Login" web page:
@app.route('/admin_login',methods=["GET", "POST"])
def admin_login():
    try:
        # Instantiate an instance of the "AdminLoginForm" class:
        form = AdminLoginForm()

        # Validate form entries upon submittal. If validated, send message:
        if form.validate_on_submit():
            # Capture the supplied e-mail address and password:
            username = form.txt_username.data
            password = form.txt_password.data

            # Check if user account exists under the supplied username:
            user = retrieve_from_database("get_user", username=username)
            if user != None:  # Account exists under the supplied username.
                # Check if supplied password matches the salted/hashed password for that account in the db:
                if check_password_hash(user.password, password):  # Passwords match
                    # Log user in:
                    login_user(user)

                    # Go to the "Administrative Update" page:
                    return redirect(url_for("admin_update"))

                else:  # Passwords do not match.
                    msg_status = "Password is incorrect.  Please try again."

                    # Go to the "Admin Login" page and display the results of e-mail execution attempt:
                    return render_template("admin_login.html", form=form, msg_status=msg_status, recognition_web_template=recognition["web_template"])

            else:  # Account does NOT exist under the supplied username.
                msg_status = "Username is incorrect.  Please try again."
                return render_template("admin_login.html", form=form, msg_status=msg_status,
                                       recognition_web_template=recognition["web_template"])

        # Go to the "Admin Login" page:
        return render_template("admin_login.html", form=form, msg_status="<<Message Being Drafted.>>", recognition_web_template=recognition["web_template"])

    except:  # An error has occurred.
        # Log error into system log:
        update_system_log("route: '/admin_login'", traceback.format_exc())

        # Go to the web page which displays error details to the user:
        return render_template("error.html", activity="route: '/admin_login'", details=traceback.format_exc())


# Configure route for logging out of "Administrative Update":
@app.route('/admin_logout')
def admin_logout():
    try:
        # Log user out:
        logout_user()

        # Go to the home page:
        return redirect(url_for('home'))

    except:  # An error has occurred.
        # Log error into system log:
        update_system_log("route: '/admin_logout'", traceback.format_exc())

        # Go to the web page which displays error details to the user:
        return render_template("error.html", activity="route: '/admin_logout'", details=traceback.format_exc())


# Configure route for "Administrative Update" web page:
@app.route('/admin_update',methods=["GET", "POST"])
@admin_only
def admin_update():
    try:
        # Go to the "Administrative Update" page:
        return render_template("admin_update.html", recognition_web_template=recognition["web_template"])

    except:  # An error has occurred.
        # Log error into system log:
        update_system_log("route: '/admin_update'", traceback.format_exc())

        # Go to the web page which displays error details to the user:
        return render_template("error.html", activity="route: '/admin_update'", details=traceback.format_exc())


# Configure route for "Credits" web page:
@app.route('/credits')
def credits():
    try:
        # Initialize variables for tracking success of obtaining necessary data from the database for this segment:
        success = False
        error_msg = ""
        credit_count = 0
        quote_count = 0

        # Query the table for number of inspirational quotes in the database:
        quote_count = retrieve_from_database("get_quote_count")
        if quote_count == {}:
            error_msg = "An error has occurred.  Data cannot be obtained at this time."
        elif quote_count == 0:
            error_msg = "No credits exist in the database."
        else:  # Count of inspirational quotes was successfully obtained.
            # Query the table for inspiration data sources (credits):
            credits_in_db = retrieve_from_database("get_data_sources")
            if credits_in_db == {}:
                error_msg = "An error has occurred.  Credits cannot be obtained and listed at this time."
            elif credits_in_db == []:
                error_msg = "No credits exist in the database."
                success = False
            else:  # Credits have been successfully retrieved.
                credit_count = len(credits_in_db)
                success = True

        # Go to the web page to render the results:
        return render_template("credits.html", credits=credits_in_db, credit_count=credit_count, quote_count=quote_count, success=success, error_msg=error_msg, recognition_web_template=recognition["web_template"])

    except:  # An error has occurred.
        # Log error into system log file:
        update_system_log("route: '/credits'", traceback.format_exc())

        # Go to the web page which displays error details to the user:
        return render_template("error.html", activity="route: '/credits'", details=traceback.format_exc())


# Configure route for "Contact Us" web page:
@app.route('/contact',methods=["GET", "POST"])
def contact():
    try:
        # Instantiate an instance of the "ContactForm" class:
        form = ContactForm()

        # Validate form entries upon submittal. If validated, send message:
        if form.validate_on_submit():
            # Send message via e-mail:
            msg_status = email_from_contact_page(form)

            # Go to the "Contact Us" page and display the results of e-mail execution attempt:
            return render_template("contact.html", msg_status=msg_status, recognition_web_template=recognition["web_template"])

        # Go to the "Contact Us" page:
        return render_template("contact.html", form=form, msg_status="<<Message Being Drafted.>>", recognition_web_template=recognition["web_template"])

    except:  # An error has occurred.
        # Log error into system log:
        update_system_log("route: '/contact'", traceback.format_exc())

        # Go to the web page which displays error details to the user:
        return render_template("error.html", activity="route: '/contact'", details=traceback.format_exc())


# Configure route for "Delete Subscriber (confirm)" web page:
@app.route('/delete_subscriber_confirm/<subscriber_id>')
@admin_only
def delete_subscriber_confirm(subscriber_id):
    try:
        # Initialize variables to track whether the desired subscriber record has been successfully obtained or if an error has occurred:
        success = False
        error_msg = ""

        # Query the database for information on desired subscriber.  Capture feedback to relay to end user:
        selected_subscriber = retrieve_from_database("get_subscriber_by_id", subscriber_id=subscriber_id)
        if selected_subscriber == {}:
            error_msg = "An error has occurred. Subscriber information cannot be obtained at this time."
        elif selected_subscriber == []:
            error_msg = "No matching records were retrieved."
        else:
            # Indicate that record retrieval has been successfully executed:
            success = True

        # Go to the "Delete Subscriber (confirm)" web page:
        return render_template("delete_subscriber_confirm.html", subscriber=selected_subscriber, success=success, error_msg=error_msg, recognition_web_template=recognition["web_template"])

    except:  # An error has occurred.
        # Log error into system log file:
        update_system_log("route: '/delete_subscriber_confirm'", traceback.format_exc())

        # Go to the web page which displays error details to the user:
        return render_template("error.html", activity="route: '/delete_subscriber_confirm'", details=traceback.format_exc())


# Configure route for "Delete Subscriber (result)" web page:
@app.route('/delete_subscriber_result/<subscriber_id>')
@admin_only
def delete_subscriber_result(subscriber_id):
    try:
        # Delete the desired subscriber record from the database.  Capture feedback to relay to end user:
        if not update_database("delete_subscriber_by_id", item_to_process=[], subscriber_id=subscriber_id):
            result = "An error has occurred. Subscriber has not been deleted."
        else:
            result = "Subscriber has been successfully deleted."

        # Go to the web page to render the results:
        return render_template("db_update_result.html", trans_type="Delete Subscriber", result=result, recognition_web_template=recognition["web_template"])

    except:  # An error has occurred.
        # Log error into system log file:
        update_system_log("route: '/delete_subscriber_result'", traceback.format_exc())

        # Go to the web page which displays error details to the user:
        return render_template("error.html", activity="route: '/delete_subscriber_result'", details=traceback.format_exc())


# Configure route for "edit subscriber" web page:
@app.route('/edit_subscriber/<subscriber_id>',methods=["GET", "POST"])
@admin_only
def edit_subscriber(subscriber_id):
    try:
        # Instantiate an instance of the "AddOrEditSubscriberForm" class:
        form = AddOrEditSubscriberForm()

        # If form-level validation has passed, perform additional processing:
        if form.validate_on_submit():
            # Initialize variable to summarize end result of this transaction attempt:
            result = ""

            # Initialize variable to track whether database update should proceed or not:
            update_db = False

            # Initialize variable to track whether subscriber e-mail address has violated the unique-value constraint:
            unique_subscriber_email_violation = False

            # Check if e-mail address of new subscriber already exists in the db. Capture feedback to relay to end user:
            email_in_db = retrieve_from_database("get_subscriber_by_email", email=form.txt_email.data)
            if email_in_db == {}:
                result = "An error has occurred. Subscriber has not been added."
            else:
                if email_in_db != None:
                    if email_in_db.id != int(subscriber_id):
                        result = f"Email address '{form.txt_email.data}' already exists in the database.  Please go back and enter a unique e-mail address."
                        unique_subscriber_email_violation = True
                    else:
                        # Indicate that database update can proceed:
                        update_db = True
                else:
                    # Indicate that database update can proceed:
                    update_db = True

                # If database update can proceed, then do so:
                if update_db:
                    # Update the desired subscriber record in the database.  Capture feedback to relay to end user:
                    if not update_database("edit_subscriber", item_to_process=[], form=form, subscriber_id=subscriber_id):
                        result = "An error has occurred. Subscriber has not been edited."
                    else:
                        result = "Subscriber has been successfully edited."

            # Go to the web page to render the results:
            return render_template("db_update_result.html", unique_subscriber_email_violation=unique_subscriber_email_violation, trans_type="Edit Subscriber", result=result, recognition_web_template=recognition["web_template"])

        # Initialize variables to track whether a record for the selected subscriber was successfully obtained or if an error has occurred:
        success = False
        error_msg = ""

        # Get information from the database for the selected subscriber. Capture feedback to relay to end user:
        selected_subscriber = retrieve_from_database("get_subscriber_by_id", subscriber_id=subscriber_id)
        if selected_subscriber == {}:
            error_msg = "An error has occurred. Subscriber information cannot be obtained at this time."
        elif selected_subscriber == []:
            error_msg = "No matching records were retrieved."
        else:
            # Populate the form with the retrieved record's contents:
            form.txt_name.data = selected_subscriber.name
            form.txt_email.data = selected_subscriber.email
            form.button_submit.label.text = "Update Subscriber"

            # Indicate that record retrieval and form population have been successful:
            success = True

        # Go to the "Edit Subscriber" web page:
        return render_template("edit_subscriber.html", form=form, subscriber=selected_subscriber, success=success, error_msg=error_msg, recognition_web_template=recognition["web_template"])

    except:  # An error has occurred.
        # Log error into system log file:
        update_system_log("route: '/edit_subscriber'", traceback.format_exc())

        # Go to the web page which displays error details to the user:
        return render_template("error.html", activity="route: '/edit_subscriber'", details=traceback.format_exc())


# Configure route for "existing subscribers" web page:
@app.route('/subscribers')
@admin_only
def subscribers():
    try:
        # Initialize variables to track whether existing subscriber records were successfully obtained or if an error has occurred:
        success = False
        error_msg = ""
        subscriber_count = 0

        # Get information on existing subscribers in the database. Capture feedback to relay to end user:
        existing_subscribers = retrieve_from_database("get_all_subscribers")
        if existing_subscribers == {}:
            error_msg = "An error has occurred. Subscriber information cannot be obtained at this time."
        elif existing_subscribers == []:
            error_msg = "No matching records were retrieved."
        else:
            subscriber_count = len(existing_subscribers)  # Record count to be displayed in the sub-header of the "existing subscribers" web page.

            # Indicate that record retrieval has been successfully executed:
            success = True

        # Go to the web page to render the results:
        return render_template("subscribers.html", subscribers=existing_subscribers, subscriber_count=subscriber_count, success=success, error_msg=error_msg, recognition_web_template=recognition["web_template"])

    except:  # An error has occurred.
        # Log error into system log file:
        update_system_log("route: '/subscribers'", traceback.format_exc())

        # Go to the web page which displays error details to the user:
        return render_template("error.html", activity="route: '/subscribers'", details=traceback.format_exc())


# Configure route for updating inspiration data from external sources:
@app.route('/update_inspirational_data')
@admin_only
def update_inspirational_data():
    try:
        # Perform update of inspiration data in the database.  Capture result of same to relay to user:
        result, success = get_inspirational_data()

        # Go to the web page to render the results:
        return render_template("admin_update.html", result=result, recognition_web_template=recognition["web_template"])

    except:  # An error has occurred.
        # Log error into system log file:
        update_system_log("route: '/update_inspirational_data'", traceback.format_exc())

        # Go to the web page which displays error details to the user:
        return render_template("error.html", activity="route: '/update_inspirational_data'", details=traceback.format_exc())


# DEFINE FUNCTIONS TO BE USED FOR THIS APPLICATION (LISTED IN ALPHABETICAL ORDER BY FUNCTION NAME):
# *************************************************************************************************
def config_database():
    """Function for configuring the database tables supporting this application"""
    global db, app, Categories, InspirationDataSources, InspirationalQuotes, Subscribers, Users

    try:
        # Create the database object using the SQLAlchemy constructor:
        db = SQLAlchemy(model_class=Base)

        # Initialize the app with the extension:
        db.init_app(app)

        # Configure database tables (listed in alphabetical order; class names are sufficiently descriptive):
        class Categories(db.Model):
            __tablename__ = "categories"
            id: Mapped[int] = mapped_column(Integer, primary_key=True)
            category: Mapped[str] = mapped_column(String(20), nullable=False, unique=True)
            children: Mapped[List["InspirationDataSources"]] = relationship(back_populates="parent")

        class InspirationDataSources(db.Model):
            __tablename__ = "inspiration_data_sources"
            id: Mapped[int] = mapped_column(Integer, primary_key=True)
            name: Mapped[str] = mapped_column(String(50), nullable=False, unique=True)
            recognition: Mapped[str] = mapped_column(String(100), nullable=False)
            category_id: Mapped[int] = mapped_column(ForeignKey("categories.id"))
            count: Mapped[int] = mapped_column(Integer, nullable=False)
            static: Mapped[bool] = mapped_column(Boolean, nullable=False)
            url: Mapped[str] = mapped_column(String(250), nullable=False)
            comments: Mapped[str] = mapped_column(String(500))
            children: Mapped[List["InspirationalQuotes"]] = relationship(back_populates="parent")
            parent: Mapped["Categories"] = relationship(back_populates="children")

        class InspirationalQuotes(db.Model):
            __tablename__ = "inspirational_quotes"
            id: Mapped[int] = mapped_column(Integer, primary_key=True)
            quote: Mapped[str] = mapped_column(String(500), nullable=False)
            data_source_id: Mapped[int] = mapped_column(ForeignKey("inspiration_data_sources.id"))
            parent: Mapped["InspirationDataSources"] = relationship(back_populates="children")

        class Subscribers(db.Model):
            id: Mapped[int] = mapped_column(Integer, primary_key=True)
            name: Mapped[str] = mapped_column(String(50), nullable=False)
            email: Mapped[str] = mapped_column(String(50), nullable=False)

        class Users(UserMixin, db.Model):
            __tablename__ = "users"
            id: Mapped[int] = mapped_column(Integer, primary_key=True)
            username: Mapped[str] = mapped_column(String(100), unique=True)
            password: Mapped[str] = mapped_column(String(100))

        # Configure the database per the above.  If needed tables do not already exist in the DB, create them:
        with app.app_context():
            db.create_all()

        # At this point, function is presumed to have executed successfully.  Return\
        # successful-execution indication to the calling function:
        return True

    except:  # An error has occurred.
        # Log error into system log:
        update_system_log("config_database", traceback.format_exc())

        # Return failed-execution indication to the calling function:
        return False


def config_web_forms():
    """Function for configuring the web forms supporting this application's website"""
    global AddOrEditSubscriberForm, AdminLoginForm, AdminUpdateForm, ContactForm

    try:
        # CONFIGURE WEB FORMS (LISTED IN ALPHABETICAL ORDER):
        # Configure 'add/edit subscriber' form:
        class AddOrEditSubscriberForm(FlaskForm):
            txt_name = StringField(label="Name:", validators=[InputRequired(), Length(max=50)])
            txt_email = EmailField(label="E-mail Address:", validators=[InputRequired(), Email(), Length(max=50)])
            button_submit = SubmitField(label="Subscribe")

        # Configure "admin login" form:
        class AdminLoginForm(FlaskForm):
            txt_username = StringField(label="Username:", validators=[InputRequired()])
            txt_password = PasswordField(label="Password:", validators=[InputRequired()])
            button_submit = SubmitField(label="Login")

        # Configure "admin update" form:
        class AdminUpdateForm(FlaskForm):
            chk_approaching_asteroids = BooleanField(label="Approaching Asteroids", default=True)
            chk_confirmed_planets = BooleanField(label="Confirmed Planets", default=True)
            chk_constellations = BooleanField(label="Constellations", default=True)
            chk_mars_photos = BooleanField(label="Photos from Mars", default=True)
            button_submit = SubmitField(label="Begin Update")

        # Configure 'contact us' form:
        class ContactForm(FlaskForm):
            txt_name = StringField(label="Your Name:", validators=[InputRequired(), Length(max=50)])
            txt_email = EmailField(label="Your E-mail Address:", validators=[InputRequired(), Email()])
            txt_message = TextAreaField(label="Your Message:", validators=[InputRequired()])
            button_submit = SubmitField(label="Send Message")

        # At this point, function is presumed to have executed successfully.  Return\
        # successful-execution indication to the calling function:
        return True

    except:  # An error has occurred.
        # Log error into system log:
        update_system_log("config_web_forms", traceback.format_exc())

        # Return failed-execution indication to the calling function:
        return False


def email_from_contact_page(form):
    """Function to process a message that user wishes to e-mail from this website to the website administrator"""
    try:
        # E-mail the message using the contents of the "Contact Us" web page form as input:
        with smtplib.SMTP(SENDER_HOST, port=SENDER_PORT) as connection:
            try:
                # Make connection secure, including encrypting e-mail.
                connection.starttls()
            except:
                # Return failed-execution message to the calling function:
                return "Error: Could not make connection to send e-mails. Your message was not sent."
            try:
                # Login to sender's e-mail server.
                connection.login(SENDER_EMAIL_GMAIL, SENDER_PASSWORD_GMAIL)
            except:
                # Return failed-execution message to the calling function:
                return "Error: Could not log into e-mail server to send e-mails. Your message was not sent."
            else:
                # Send e-mail.
                connection.sendmail(
                    from_addr=SENDER_EMAIL_GMAIL,
                    to_addrs=SENDER_EMAIL_GMAIL,
                    msg=f"Subject: Inspire Me! - E-mail from 'Contact Us' page\n\nName: {form.txt_name.data}\nE-mail address: {form.txt_email.data}\n\nMessage:\n{form.txt_message.data}"
                )
                # Return successful-execution message to the calling function::
                return "Your message has been successfully sent."

    except:  # An error has occurred.
        # Log error into system log:
        update_system_log("email_from_contact_page", traceback.format_exc())

        # Return failed-execution message to the calling function:
        return "An error has occurred. Your message was not sent."


def email_quotes_to_distribution(message, subscribers_list):
    """Function to e-mail a list of randomly-selected quotes (by category) to subscribers"""
    global SENDER_EMAIL_GMAIL, SENDER_HOST, SENDER_PASSWORD_GMAIL, SENDER_PORT

    try:
        # E-mail the message using the 'message' parameter:
        with smtplib.SMTP(SENDER_HOST, port=SENDER_PORT) as connection:
            try:
                # Make connection secure, including encrypting e-mail.
                connection.starttls()
            except:
                # Return failed-execution message to the calling function:
                return "Error: Could not make connection to send e-mails. Message was not sent."
            try:
                # Login to sender's e-mail server.
                connection.login(SENDER_EMAIL_GMAIL, SENDER_PASSWORD_GMAIL)
            except:
                # Return failed-execution message to the calling function:
                return "Error: Could not log into e-mail server to send e-mails. Message was not sent."
            else:
                # Send e-mail.
                connection.sendmail(
                    from_addr=SENDER_EMAIL_GMAIL,
                    to_addrs=subscribers_list,
                    msg=f"Subject: Inspiration from the 'Inspire Me' website\n\n{message}"
                )
                # Return successful-execution message to the calling function::
                return "Success"

    except:  # An error has occurred.
        # Log error into system log:
        update_system_log("email_quotes_to_distribution", traceback.format_exc())

        # Return failed-execution message to the calling function:
        return "An error has occurred. Message not sent."


def find_element(driver, find_type, find_details):
    """Function to find an element via a web-scraping procedure"""
    # NOTE: Error handling is deferred to the calling function:
    if find_type == "xpath":
        return driver.find_element(By.XPATH, find_details)


def get_inspirational_data():
    """Function for getting (via web-scraping) inspirational data and storing such information in the database supporting our website"""
    global dlg

    try:
        data_sources_to_update = retrieve_from_database(trans_type="get_non-static_data_sources")
        if data_sources_to_update == {}:
            return "Error: Data could not be updated at this time.", False
        elif data_sources_to_update == []:
            pass  # Defer to next set of code to produce user feedback:

        # If at least one data source has been flagged for update, perform update.  Otherwise, return appropriate feedback to user:
        if len(data_sources_to_update) == 0:  # No data sources have been flagged for updating.
            return "No data sources have been identified for updating.", True
        else:
            for i in range(0, len(data_sources_to_update)):
                # Get results of obtaining and processing the desired information for the current inspiration data source (use window dialog to keep user informed):
                dlg = wx.App()
                dlg = PBI.PyBusyInfo(f"{data_sources_to_update[i].name}: Update in progress...", title="Administrative Update")

                # Get the inspirational data from each identified data source.
                # If the function called returns an empty directory,
                inspirational_data = get_inspirational_data_details(data_sources_to_update[i].id, data_sources_to_update[i].count, data_sources_to_update[i].name, data_sources_to_update[i].url)
                if inspirational_data == []:
                    update_system_log("get_inspirational_data", f"Error: Data (for source = '{data_sources_to_update[i].name}') cannot be updated at this time.")
                    return f"Error: Data (for source = {data_sources_to_update[i].name}) cannot be updated at this time.", False

                # Delete the existing records in the "inspirational_quotes" database table and update same with the
                # contents that were obtained via the previous step.  If the function called returns a failed-execution
                # indication, update system log and return failed-execution indication to the calling function:
                if not update_database("update_inspirational_quotes", inspirational_data, source=data_sources_to_update[i].id):
                    update_system_log("get_inspirational_data",
                                      "Error: Database could not be updated. Data cannot be updated at this time.")
                    return "Error: Database could not be updated. Data cannot be updated at this time.", False

                # Deactivate the dialog object:
                dlg = None

        # At this point, function is deemed to have executed successfully.  Update system log and return
        # successful-execution indication to the calling function:
        update_system_log("get_inspirational_data", "Successfully updated.")
        return "Inspirational data has been successfully updated.", True

    except:  # An error has occurred.
        # Clear the wx variable:
        dlg = None

        # Log error into system log:
        update_system_log("get_inspirational_data", traceback.format_exc())

        # Return failed-execution indication to the calling function:
        return "An error has occurred. Data cannot be obtained at this time.", False


def get_inspirational_data_details(source, count, name, url):
    """Function for getting (via web-scraping) inspirational quotes from identified data sources"""

    # Define a variable for storing the scraped inspirational data:
    inspirational_data = []

    try:
        # Initiate and configure a Selenium object to be used for scraping the website representing the
        # inspiration-data source. If function failed, update system log and return failed-execution indication to the calling function:
        driver = setup_selenium_driver(url, 1, 1)
        if driver == None:
            update_system_log("get_inspirational_data_details",
                              f"Error: Selenium driver (for {name}) could not be created/configured.")
            return []

        # Pause program execution to allow for website loading time:
        time.sleep(WEB_LOADING_TIME_ALLOWANCE)

        # Scrape data source identified via parameter passed to this function:
        if source == 1 or source == 2 or source == 3 or source == 13 or source == 21 or source == 22 or source == 24:
            for i in range(1, count + 20):
                try:
                    # Scrape element:
                    element_quote = find_element(driver, "xpath",'/html/body/div[4]/div[2]/main/article/div/blockquote[' + str(i) + ']/p')

                    # Add the quote to the "inspirational_data" list:
                    inspirational_data.append(element_quote.text)

                except:
                    continue

        elif source == 4:
            for i in range(1, 21):
                for j in range(1, 21):
                    try:
                        # Scrape element:
                        element_quote = find_element(driver, "xpath",'/html/body/main/article/div[3]/div[2]/div[1]/ul[' + str(i) + ']/li[' + str(j) + ']')

                        # Strip unwanted characters:
                        element_quote_1 = element_quote.text.replace('“', '')
                        element_quote_2 = element_quote_1.replace('"', '')

                        # Add the quote to the "inspirational_data" list:
                        inspirational_data.append(element_quote_2)

                    except:
                        continue

        elif source == 5:
            for i in range(1, 21):
                for j in range(1, 21):
                    try:
                        # Scrape element:
                        element_quote = find_element(driver, "xpath",'/html/body/main/article/div[3]/div[3]/div[1]/ul[' + str(i) + ']/li[' + str(j) + ']')

                        # Strip unwanted characters:
                        element_quote_1 = element_quote.text.replace('”', '')
                        element_quote_2 = element_quote_1.replace('“', '')
                        element_quote_3 = element_quote_2.replace('"', '')

                        # Add the quote to the "inspirational_data" list:
                        inspirational_data.append(element_quote_3)

                    except:
                        continue

        elif source == 6:
            for i in range(7, count + 20):
                try:
                    # Scrape element:
                    element_quote = find_element(driver, "xpath",
                                                 '/html/body/div[2]/div/div/div/div[2]/div/article/div/div/div[2]/p[' + str(i) + ']')

                    # If item is an inspiration quote, strip it of unwanted characters and add it to the list for subsequent update of the database.
                    element_quote_1 = element_quote.text.strip()
                    if element_quote_1 == "":
                        pass  # Item will not be added to the database.
                    else:
                        if element_quote_1[0] == '“' or element_quote_1[0] == '"':  # Item is an inspiration quote.  Therefore, it will be added to the database.
                            # Strip unwanted characters:
                            element_quote_2 = element_quote_1.replace("\n", "")
                            element_quote_3 = element_quote_2.replace('“', '')
                            element_quote_4 = element_quote_3.replace('”–', '-')

                            # Add the quote to the "inspirational_data" list:
                            inspirational_data.append(element_quote_4)

                except:
                    try:
                        # Scrape element:
                        element_quote = find_element(driver, "xpath",
                                                     '/html/body/div[3]/div/div/div/div[2]/div/article/div/div/div[2]/p[' + str(i) + ']')

                        # If item is an inspiration quote, strip it of unwanted characters and add it to the list for subsequent update of the database.
                        element_quote_1 = element_quote.text.strip()
                        if element_quote_1 == "":
                            pass  # Item will not be added to the database.
                        else:
                            if element_quote_1[0] == '“' or element_quote_1[0] == '"':  # Item is an inspiration quote.  Therefore, it will be added to the database.                            # Strip unwanted characters:
                                element_quote_2 = element_quote_1.replace("\n", "")
                                element_quote_3 = element_quote_2.replace('“', '')
                                element_quote_4 = element_quote_3.replace('”–', '-')

                                # Add the quote to the "inspirational_data" list:
                                inspirational_data.append(element_quote_4)

                    except:
                        continue

            for i in range(1, count + 20):
                try:
                    # Scrape element:
                    element_quote = find_element(driver, "xpath",
                                                 '/html/body/div[2]/div/div/div/div[2]/div/article/div/div/div[2]/blockquote[' + str(i) + ']/p')

                    # If item is an inspiration quote, strip it of unwanted characters and add it to the list for subsequent update of the database.
                    element_quote_1 = element_quote.text.strip()
                    if element_quote_1 == "":
                        pass  # Item will not be added to the database.
                    else:
                        if element_quote_1[0] == '“' or element_quote_1[0] == '"':  # Item is an inspiration quote.  Therefore, it will be added to the database.                            # Strip unwanted characters:
                            element_quote_2 = element_quote_1.replace("\n", "")
                            element_quote_3 = element_quote_2.replace('“', '')
                            element_quote_4 = element_quote_3.replace('”–', '-')

                            # Add the quote to the "inspirational_data" list:
                            inspirational_data.append(element_quote_4)

                except:
                    try:
                        # Scrape element:
                        element_quote = find_element(driver, "xpath",
                                                     '/html/body/div[3]/div/div/div/div[2]/div/article/div/div/div[2]/blockquote[' + str(i) + ']/p')

                        # If item is an inspiration quote, strip it of unwanted characters and add it to the list for subsequent update of the database.
                        element_quote_1 = element_quote.text.strip()
                        if element_quote_1 == "":
                            pass  # Item will not be added to the database.
                        else:
                            if element_quote_1[0] == '“' or element_quote_1[0] == '"':  # Item is an inspiration quote.  Therefore, it will be added to the database.                            # Strip unwanted characters:
                                element_quote_2 = element_quote_1.replace("\n", "")
                                element_quote_3 = element_quote_2.replace('“', '')
                                element_quote_4 = element_quote_3.replace('”–', '-')

                                # Add the quote to the "inspirational_data" list:
                                inspirational_data.append(element_quote_4)

                    except:
                        continue

            for i in range(1, count + 20):
                try:
                    # Scrape element:
                    element_quote = find_element(driver, "xpath",
                                                 '/html/body/div[2]/div/div/div/div[2]/div/article/div/div/div[2]/div[' + str(i) + ']/div[1]/a')

                    # If item is an inspiration quote, strip it of unwanted characters and add it to the list for subsequent update of the database.
                    element_quote_1 = element_quote.text.strip()
                    if element_quote_1 == "":
                        pass  # Item will not be added to the database.
                    else:
                        element_quote_2 = element_quote_1.replace("\n", "")
                        element_quote_3 = element_quote_2.replace('“', '')
                        element_quote_4 = element_quote_3.replace('”–', '-')
                        element_quote_5 = element_quote_4.replace('” –', ' -')

                        # Add the quote to the "inspirational_data" list:
                        inspirational_data.append(element_quote_5)
                except:
                    try:
                        # Scrape element:
                        element_quote = find_element(driver, "xpath",
                                                     '/html/body/div[3]/div/div/div/div[2]/div/article/div/div/div[2]/div[' + str(
                                                         i) + ']/div[1]/a')

                        # If item is an inspiration quote, strip it of unwanted characters and add it to the list for subsequent update of the database.
                        element_quote_1 = element_quote.text.strip()
                        if element_quote_1 == "":
                            pass  # Item will not be added to the database.
                        else:
                            element_quote_2 = element_quote_1.replace("\n", "")
                            element_quote_3 = element_quote_2.replace('“', '')
                            element_quote_4 = element_quote_3.replace('”–', '-')
                            element_quote_5 = element_quote_4.replace('” –', ' -')

                            # Add the quote to the "inspirational_data" list:
                            inspirational_data.append(element_quote_5)

                    except:
                        continue

        elif source == 7:
            for i in range(1, 21):
                for j in range(1, 4):
                    for k in range(1, 16):
                        try:
                            # Scrape element:
                            element_quote = find_element(driver, "xpath",'/html/body/div[1]/main/div[2]/section[' + str(i) + ']/section[' + str(j) + ']/div/div/div[2]/div/p[' + str(k) + ']')

                            # Strip unwanted characters:
                            if not (element_quote.text == "" or "RELATED" in element_quote.text):
                                element_quote_1 = element_quote.text.replace('”', '')
                                element_quote_2 = element_quote_1.replace('“', '')
                                element_quote_3 = element_quote_2.replace('"', '')

                                # Add the quote to the "inspirational_data" list:
                                inspirational_data.append(element_quote_3)

                        except:
                            continue

        elif source == 8:
            # The website for this data source responds differently than others.  As such, with the standard window height/width used, elements
            # cannot be found.  Therefore, prior to scraping, maximize driver window so that elements can be found:
            driver.maximize_window()

            # Proceed with scraping:
            for i in range(1, count + 20):
                try:
                    # Scrape element:
                    element_quote = find_element(driver, "xpath",
                                                 '/html/body/div[1]/div[1]/div[1]/div[1]/div[2]/main/article/div/div/div/div/div[3]/div/div[2]/div[1]/blockquote[' + str(i) + ']/p')

                    # If item is an inspiration quote, strip it of unwanted characters and add it to the list for subsequent update of the database.
                    element_quote_1 = element_quote.text.strip()
                    if element_quote_1 == "":
                        pass  # Item will not be added to the database.
                    else:
                        if element_quote_1[0] == '“' or element_quote_1[0] == '"':  # Item is an inspiration quote.  Therefore, it will be added to the database.
                            # Strip unwanted characters:
                            element_quote_2 = element_quote_1.replace("\n", "")
                            element_quote_3 = element_quote_2.replace('“', '')
                            element_quote_4 = element_quote_3.replace('"', '')
                            element_quote_5 = element_quote_4.replace('”. –', '.-')

                            # Add the quote to the "inspirational_data" list:
                            inspirational_data.append(element_quote_5)

                except:
                    continue

        elif source == 9:
            for i in range(1, 31):
                for j in range(1, 31):
                    try:
                        # Scrape element:
                        element_quote = find_element(driver, "xpath",'/html/body/div[2]/main/section/article/div[2]/div[2]/div[1]/div[2]/ol[' + str(i) + ']/li[' + str(j) + ']')

                        # Strip unwanted characters:
                        element_quote_1 = element_quote.text.replace('”', '')
                        element_quote_2 = element_quote_1.replace('“', '')
                        element_quote_3 = element_quote_2.replace('"', '')

                        # Add the quote to the "inspirational_data" list:
                        inspirational_data.append(element_quote_3)

                    except:
                        continue

        elif source == 10:
            for i in range(1, count + 30):
                try:
                    # Scrape element:
                    element_quote = find_element(driver, "xpath",
                                                 '/html/body/div[2]/div[2]/div[1]/main/article/div/div/p[' + str(i) + ']')

                    # If item is an inspiration quote, strip it of unwanted characters and add it to the list for subsequent update of the database.
                    element_quote_1 = element_quote.text.strip()
                    if element_quote_1 == "":
                        pass  # Item will not be added to the database.
                    else:
                        if element_quote_1[0] == '“' or element_quote_1[0] == '"':  # Item is an inspiration quote.  Therefore, it will be added to the database.
                            # Strip unwanted characters:
                            element_quote_2 = element_quote_1.replace("\n", "")
                            element_quote_3 = element_quote_2.replace('“', '')
                            element_quote_4 = element_quote_3.replace('”–', '-')

                            # Add the quote to the "inspirational_data" list:
                            inspirational_data.append(element_quote_4)

                except:
                    continue

        elif source == 11:
            for i in range(1, count + 20):
                try:
                    # Scrape element:
                    element_quote = find_element(driver, "xpath",
                                                 '/html/body/div[1]/div/div[1]/main/article/div/div/ol/li[' + str(i) + ']')

                    # If item is an inspiration quote, strip it of unwanted characters and add it to the list for subsequent update of the database.
                    element_quote_1 = element_quote.text.strip()
                    if element_quote_1 == "":
                        pass  # Item will not be added to the database.
                    else:
                        # Strip unwanted characters:
                        element_quote_2 = element_quote_1.replace("\n", "")
                        element_quote_3 = element_quote_2.replace('“', '')
                        element_quote_4 = element_quote_3.replace('”', '-')

                        # Add the quote to the "inspirational_data" list:
                        inspirational_data.append(element_quote_4)

                except:
                    continue

        elif source == 12:
            for i in range(1, count + 20):
                try:
                    # Scrape element:
                    element_quote = find_element(driver, "xpath",
                                                 '/html/body/div[1]/div[4]/div/main/div/div/div/div/div/article/div/figure[' + str(i) + ']/figcaption')

                    # If item is an inspiration quote, strip it of unwanted characters and add it to the list for subsequent update of the database.
                    element_quote_1 = element_quote.text.strip()
                    if element_quote_1 == "":
                        pass  # Item will not be added to the database.
                    else:
                        # Strip unwanted characters:
                        element_quote_2 = element_quote_1.replace("\n", "")
                        element_quote_3 = element_quote_2.replace('“', '')
                        element_quote_4 = element_quote_3.replace('”', '-')

                        # Add the quote to the "inspirational_data" list:
                        inspirational_data.append(element_quote_4)

                except:
                    continue

        elif source == 14:
            for i in range(7, count + 28):
                try:
                    # Scrape element:
                    element_quote = find_element(driver, "xpath",'/html/body/div[1]/div[4]/div[1]/div[2]/div/div[1]/p[' + str(i) + ']')
                    element_quote_1 = html2text.html2text(element_quote.get_attribute("outerHTML"))

                    if not (":" in element_quote_1 and not ("_" in element_quote_1)):
                        pass  # Item will not be added to the database.
                    else:
                        # Strip unwanted characters:
                        element_quote_2 = element_quote_1.replace("*", "")
                        element_quote_3 = element_quote_2.replace("\n", " ")
                        element_quote_4 = element_quote_3.strip()

                        # Add the quote to the "inspirational_data" list:
                        inspirational_data.append(element_quote_4)

                except:
                    continue

            for i in range(1, 21):
                for j in range(1, 61):
                    try:
                        # Scrape element:
                        element_quote = find_element(driver, "xpath",
                                                     '/html/body/div[1]/div[4]/div[1]/div[2]/div/div[2]/div[' + str(i) + ']/div/p[' + str(j) + ']')
                        element_quote_1 = html2text.html2text(element_quote.get_attribute("outerHTML"))

                        element_quote_2 = element_quote_1.replace("*", "")
                        element_quote_3 = element_quote_2.replace("\n", " ")
                        element_quote_4 = element_quote_3.replace("[", "")
                        element_quote_5 = element_quote_4.replace("]", "")
                        element_quote_6 = element_quote_5.replace('“', "")
                        element_quote_7 = element_quote_6.replace('"', "")
                        element_quote_8 = element_quote_7.replace('”', "")

                        if "(" in element_quote_8 or ")" in element_quote_8:
                            res = "*"
                            while res != "":
                                # initializing substrings
                                sub1 = "("
                                sub2 = ")"

                                # getting index of substrings
                                idx1 = element_quote_8.find(sub1)
                                idx2 = element_quote_8.find(sub2)

                                # length of substring 1 is added to
                                # get string from next character
                                res = element_quote_8[idx1 + len(sub1): idx2]

                                if res == "":
                                    break
                                else:
                                    element_quote_8 = element_quote_8.replace(res, "")

                        element_quote_9 = element_quote_8.replace("()", "")
                        element_quote_10 = element_quote_9.strip()

                        if element_quote_10 == "" or ":" not in element_quote_10 or len(element_quote_10) == 1:
                            pass  # Item will not be added to the database.
                        else:
                            # Add the quote to the "inspirational_data" list:
                            inspirational_data.append(element_quote_10)

                    except:
                        continue

        elif source == 15:
            for i in range(1, count + 20):
                try:
                    # Scrape element:
                    element_quote = find_element(driver, "xpath",
                                                 '/html/body/main/article/div[3]/blockquote[' + str(i) + ']/p')

                    # Add the quote to the "inspirational_data" list:
                    inspirational_data.append(element_quote.text)

                except:
                    continue

        elif source == 16 or source == 17 or source == 18:
            for i in range(1, count + 20):
                try:
                    # Scrape element:
                    element_quote = find_element(driver, "xpath",
                                                 '/html/body/div[1]/div/div/div[1]/main/article/div/div/div/div/p[' + str(i) + ']')

                    # Strip unwanted characters:
                    idx1 = element_quote.text.find(".")
                    res = element_quote.text[0: idx1 + len(".")]
                    element_quote_1 = element_quote.text.replace(res, "")
                    element_quote_2 = element_quote_1.strip()

                    if element_quote_2[0] == '“':
                        element_quote_3 = element_quote_2.replace('“', "")
                        element_quote_4 = element_quote_3.replace('”', "")

                        # Add the quote to the "inspirational_data" list:
                        inspirational_data.append(element_quote_4)

                except:
                    continue

        elif source == 19 or source == 20 or source == 25:
            for i in range(1, count + 20):
                try:
                    # Scrape element:
                    element_quote = find_element(driver, "xpath",
                                                 '/html/body/div[1]/div/div/div/main/div/div/div/div/div[2]/div/div[2]/p[' + str(i) + ']')

                    # Strip unwanted characters:
                    element_quote_1 = element_quote.text.replace("‘", "")
                    element_quote_2 = element_quote_1.replace("’", "")

                    # Add the quote to the "inspirational_data" list:
                    inspirational_data.append(element_quote_2)

                except:
                    continue

        elif source == 23:
            for i in range(1, count + 20):
                try:
                    # Scrape element:
                    element_quote = find_element(driver, "xpath",
                                                 '/html/body/div[1]/div/main/div/section/article/div[2]/p[' + str(i) + ']')
                    element_quote_1 = html2text.html2text(element_quote.get_attribute("outerHTML"))

                    if '“' in element_quote_1 and not ("Rufus" in element_quote_1):
                        # Strip unwanted characters:
                        element_quote_2 = element_quote_1.replace("\n", " ")
                        element_quote_3 = element_quote_2.replace('“', "")
                        element_quote_4 = element_quote_3.replace('"', "")
                        element_quote_5 = element_quote_4.replace('”', "")

                        # Add the quote to the "inspirational_data" list:
                        inspirational_data.append(element_quote_5.strip() + " - Gaius Musonius Rufus")

                except:
                    continue

        elif source == 26:
            for i in range(1, count + 20):
                try:
                    # Scrape element:
                    element_quote = find_element(driver, "xpath", '/html/body/div[1]/main/div[3]/div[1]/div[1]/div[1]/p[' + str(i) + ']')
                    element_quote_1a = element_quote.get_attribute("outerHTML")

                    # initializing substrings
                    sub1 = "</em>"
                    sub2 = "</p>"

                    # getting index of substrings
                    idx1 = element_quote_1a.find(sub1)
                    idx2 = element_quote_1a.find(sub2)

                    # length of substring 1 is added to
                    # get string from next character
                    res = element_quote_1a[idx1 + len(sub1): idx2 + len(sub2) + 1]

                    if res == "":
                        break
                    else:
                        element_quote_1b = element_quote_1a.replace(res, "")

                    element_quote_1c = html2text.html2text(element_quote_1b)
                    element_quote_2 = element_quote_1c.replace("_", "")
                    element_quote_3 = element_quote_2.replace("\n", " ")
                    element_quote_4 = element_quote_3.replace('“', "")
                    element_quote_5 = element_quote_4.replace('”', "")

                    if not ('\\.' in element_quote_5):
                        continue

                    # initializing substrings
                    sub1 = "\\."

                    # getting index of substrings
                    idx1 = element_quote_5.find(sub1)

                    # length of substring 1 is added to
                    # get string from next character
                    res = element_quote_5[0:+ idx1 + len(sub1) + 1]

                    if res == "":
                        break
                    else:
                        element_quote_6 = element_quote_5.replace(res, "")

                    if "(" in element_quote_6 or ")" in element_quote_6:
                        # initializing substrings
                        sub1 = "("
                        sub2 = ")"

                        # getting index of substrings
                        idx1 = element_quote_6.find(sub1)
                        idx2 = element_quote_6.find(sub2)

                        # length of substring 1 is added to
                        # get string from next character
                        res = element_quote_6[idx1 + len(sub1): idx2]

                        if res == "":
                            break
                        else:
                            element_quote_7 = element_quote_6.replace(res, "")
                            element_quote_8 = element_quote_7.replace("()", "")
                    else:
                        element_quote_8 = element_quote_6

                    # Add the quote to the "inspirational_data" list:
                    inspirational_data.append(element_quote_8.strip())

                except:
                    continue

        else:
            pass

        # Close and delete the Selenium driver object:
        driver.close()
        del driver

        # Return the populated "inspirational_data" list to the calling function:
        return inspirational_data

    except:  # An error has occurred.
        # Log error into system log:
        update_system_log("get_inspirational_data_details", traceback.format_exc())

        # Return empty directory as a failed-execution indication to the calling function:
        return []


def retrieve_from_database(trans_type, **kwargs):
    """Function to retrieve data from this application's database based on the type of transaction"""
    try:
        with app.app_context():
            if trans_type == "get_all_subscribers":
                # Retrieve and return all existing subscribers, sorted by name, from the "subscribers" database table:
                return db.session.execute(db.select(Subscribers).order_by(func.lower(Subscribers.name))).scalars().all()

            elif trans_type == "get_categories":
                # Retrieve and return all records from the "categories" database table:
                return db.session.execute(db.select(Categories).order_by(Categories.id)).scalars().all()

            elif trans_type == "get_data_sources":
                # Retrieve and return all existing records, sorted by category and name, from the "inspiration_data_sources" database table
                # (inner-joined with the "categories" database table):
                return db.session.query(InspirationDataSources, Categories).join(Categories, InspirationDataSources.category_id == Categories.id).order_by(Categories.category, InspirationDataSources.name).all()

            elif trans_type == "get_non-static_data_sources":
                # Retrieve and return all existing records, sorted by ID #, from the "inspiration_data_sources" database table, where the static :
                return db.session.execute(db.select(InspirationDataSources).where(InspirationDataSources.static == 0).order_by(InspirationDataSources.id)).scalars().all()

            elif trans_type == "get_quote_count":
                # Retrieve and return all existing records from the "inspirational_quotes" database table:
                quotes = db.session.execute(db.select(InspirationalQuotes)).scalars().all()

                # Return the count of retrieved records to the calling function:
                return len(quotes)

            elif trans_type == "get_quotes_for_category":
                # Capture optional argument:
                category_id = kwargs.get("category_id", None)

                # Retrieve and return all records in the "inspirational_quotes" database table where the data_source_id is
                # associated with the category ID passed to this function.  Use an inner join between the
                # "inspirational_quotes" and "inspiration_data_sources" database tables:
                dataset_join = db.session.query(InspirationalQuotes, InspirationDataSources).join(InspirationDataSources, InspirationalQuotes.data_source_id == InspirationDataSources.id).filter(InspirationDataSources.category_id == category_id).all()

                # Retrieve and return all records in the join where the category ID matches the ID passed to this function:
                return dataset_join

            elif trans_type == "get_subscriber_by_email":
                # Capture optional argument:
                email = kwargs.get("email", None)

                # Retrieve and return record from the database for the desired subscriber (email):
                return db.session.execute(db.select(Subscribers).where(Subscribers.email.ilike(email))).scalar()

            elif trans_type == "get_subscriber_by_id":
                # Capture optional argument:
                subscriber_id = kwargs.get("subscriber_id", None)

                # Retrieve and return the record for the desired ID:
                return db.session.execute(db.select(Subscribers).where(Subscribers.id == subscriber_id)).scalar()

            elif trans_type == "get_subscribers":
                return db.session.execute(db.select(Subscribers).order_by(Subscribers.name)).scalars().all()

            elif trans_type == "get_user":
                # Capture optional argument:
                username = kwargs.get("username", None)

                # Retrieve and return record from the database for the desired username:
                return db.session.execute(db.select(Users).where(Users.username.ilike(username))).scalar()

            elif trans_type == "load_user":
                # Capture optional argument:
                user_id= kwargs.get("user_id", None)

                # Retrieve and return record from the database for the desired user ID:
                return db.session.execute(db.select(Users).where(Users.id == user_id)).scalar()

    except:  # An error has occurred.
        # Log error into system log:
        update_system_log("retrieve_from_database (" + trans_type + ")", traceback.format_exc())

        # Return empty dictionary as a failed-execution indication to the calling function:
        return {}


def run_app():
    """Main function for this application"""
    global app

    try:
        # Load environmental variables from the ".env" file:
        load_dotenv()

        # Configure the SQLite database, relative to the app instance folder:
        app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///inspiration.db"

        # Initialize an instance of Bootstrap5, using the "app" object defined above as a parameter:
        Bootstrap5(app)

        # Retrieve the secret key to be used for CSRF protection:
        app.secret_key = os.getenv("SECRET_KEY_FOR_CSRF_PROTECTION")

        # Configure database tables.  If function failed, update system log and return
        # failed-execution indication to the calling function:
        if not config_database():
            update_system_log("run_app", "Error: Database configuration failed.")
            return False

        # Configure web forms.  If function failed, update system log and return
        # failed-execution indication to the calling function:
        if not config_web_forms():
            update_system_log("run_app", "Error: Web forms configuration failed.")
            return False

        # At this point, function is presumed to have executed successfully.  Return
        # successful-execution indication to the calling function:
        return True

    except:  # An error has occurred.
        # Log error into system log:
        update_system_log("run_app", traceback.format_exc())

        # Return failed-execution indication to the calling function:
        return False


def setup_selenium_driver(url, width, height):
    """Function for initiating and configuring a Selenium driver object"""

    try:
        # Keep Chrome browser open after program finishes:
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_experimental_option("detach", True)

        # Create and configure the Chrome driver (pass above options into the web driver):
        driver = webdriver.Chrome(options=chrome_options)

        # Access the desired URL.
        driver.get(url)

        # Set window position and dimensions, with the latter being large enough to display the website's elements needed:
        driver.set_window_position(0, 0)
        driver.set_window_size(width, height)

        # Return the Selenium driver object to the calling function:
        return driver

    except:  # An error has occurred.
        # Log error into system log:
        update_system_log("setup_selenium_driver", traceback.format_exc())

        # Return failed-execution indication to the calling function:
        return None


def share_quotes_with_distribution():
    """Function to prepare and e-mail a list of randomly-selected quotes (by category) to subscribers"""
    try:
        # Retrieve a list of categories.  If an error occurs or if no records are retrieved, return
        # failed-execution indication to the calling function:
        categories = retrieve_from_database("get_categories")
        if categories == {} or categories == []:
            return "Could not retrieve categories from database"

        # Prepare message introduction:
        message_intro = f"INSPIRATIONAL DATA FOR YOU:\n\n"

        # Initialize variable for storing main body of e-mail to be sent to subscribers:
        message_main_contents = ""

        # For each category retrieved, retrieve all quotes in the database that are assigned said category.
        # If an error occurs, return failed-execution indication to the calling function:
        for i in range(0,len(categories)):
            quotes_in_category = retrieve_from_database("get_quotes_for_category", category_id = categories[i].id)
            if quotes_in_category == {}:
                return f"Could not retrieve quotes from database for category '{categories[i].category}'"

            # Select, at random, one quote from the batch of quotes retrieved above:
            selected_quote = quotes_in_category[random.choice(range(0, len(quotes_in_category)))][0].quote

            # Add that quote (prefixed by the category name) to the main body of e-mail to be sent to subscribers:
            message_main_contents += f"{categories[i].category.upper()}:\n{selected_quote.encode('ascii', 'ignore').decode('ascii')}\n\n"

        # Combine the message intro and the main body of e-mail, and store in a variable:
        message_to_send = message_intro + message_main_contents

        # Retrieve a list of subscribers to whom the e-mail shall be sent.  If an error occurs or if
        # no records are retrieved, return failed-execution indication to the calling function:
        subscribers_from_db = retrieve_from_database("get_subscribers")
        if subscribers_from_db == {} or subscribers_from_db == []:
            return "Could not retrieve subscribers from database"

        # Initialize list variable to gather all subscribers to whom e-mail shall be sent:
        subscribers_list = []

        # Add each retrieved subscriber to the list of subscribers to whom e-mail shall be sent:
        for i in range(0,len(subscribers_from_db)):
            # subscribers += subscribers_from_db[i].email + ","
            subscribers_list.append(subscribers_from_db[i].email)

        # Send the e-mail and return result of same to the calling function:
        return email_quotes_to_distribution(message=message_to_send, subscribers_list=subscribers_list)

    except:  # An error has occurred.
        # Log error into system log:
        update_system_log("share_quotes_with_distribution", traceback.format_exc())

        # Return failed-execution indication to the calling function:
        return "Error in function 'share_quotes_with_distribution'"


def update_database(trans_type, item_to_process, **kwargs):
    """Function to update this application's database based on the type of transaction"""
    try:
        with app.app_context():
            if trans_type == "add_subscriber":
                # Capture optional argument:
                form = kwargs.get("form", None)

                # Upload, to the "subscribers" database table, all contents of the "form" parameter:
                new_records = []
                new_record = Subscribers(
                            name=form.txt_name.data,
                            email=form.txt_email.data
                        )

                new_records.append(new_record)

                db.session.add_all(new_records)
                db.session.commit()

            elif trans_type == "delete_subscriber_by_id":
                # Capture optional argument:
                subscriber_id = kwargs.get("subscriber_id", None)

                # Delete the record associated with the selected ID:
                db.session.query(Subscribers).where(Subscribers.id == subscriber_id).delete()
                db.session.commit()

            elif trans_type == "edit_subscriber":
                # Capture optional arguments:
                form = kwargs.get("form", None)
                subscriber_id = kwargs.get("subscriber_id", None)

                # Edit record for the selected ID, using data in the "form" parameter passed to this function:
                record_to_edit = db.session.query(Subscribers).filter(Subscribers.id == subscriber_id).first()
                record_to_edit.name = form.txt_name.data
                record_to_edit.email = form.txt_email.data

                db.session.commit()

            elif trans_type == "update_inspirational_quotes":
                # Capture optional argument:
                source = kwargs.get("source", None)

                # Delete all records from the "inspirational_quotes" database table for the inspirational data source
                # being processed:
                db.session.execute(db.delete(InspirationalQuotes).where(InspirationalQuotes.data_source_id == source))
                db.session.commit()

                # Upload, to the "inspirational_quotes" database table, all contents of the "item_to_process" parameter:
                new_records = []
                for quote in item_to_process:
                    if quote != "":
                        new_record = InspirationalQuotes(
                            quote=quote,
                            data_source_id=source
                        )

                    new_records.append(new_record)

                db.session.add_all(new_records)
                db.session.commit()

        # Return successful-execution indication to the calling function:
        return True

    except:  # An error has occurred.
        # Log error into system log:
        update_system_log("update_database (" + trans_type + ")", traceback.format_exc())

        # Return failed-execution indication to the calling function:
        return False


def update_system_log(activity, log):
    """Function to update the system log, either to log errors encountered or log successful execution of milestone admin. updates"""
    try:
        # Capture current date/time:
        current_date_time = datetime.now()
        current_date_time_file = current_date_time.strftime("%Y-%m-%d")

        # Update log file.  If log file does not exist, create it:
        with open("log_inspire_me_" + current_date_time_file + ".txt", "a") as f:
            f.write(datetime.now().strftime("%Y-%m-%d @ %I:%M %p") + ":\n")
            f.write(activity + ": " + log + "\n")

        # Close the log file:
        f.close()

    except:
        dlg = wx.App()
        dlg = wx.MessageBox(f"Error: System log could not be updated.\n{traceback.format_exc()}", 'Error',
                            wx.OK | wx.ICON_INFORMATION)


# Run main function for this application:
run_app()

# Destroy the object that was created to show user dialog and message boxes:
dlg.Destroy()

if __name__ == "__main__":
    app.run(debug=True, port=5003)