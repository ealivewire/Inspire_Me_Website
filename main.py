# PROFESSIONAL PROJECT: Inspire Me Website

# OBJECTIVE: To implement a website which automates the retrieval and distribution of daily inspiration quotes.

# Import necessary library(ies):
# import requests

from data import app, db, recognition, SENDER_EMAIL_GMAIL, SENDER_HOST, SENDER_PASSWORD_GMAIL, SENDER_PORT, WEB_LOADING_TIME_ALLOWANCE
from data import Categories, InspirationDataSources, InspirationalQuotes, Subscribers, Users
from data import AdminLoginForm, AdminUpdateForm, ContactForm, DisplayApproachingAsteroidsSheetForm, DisplayConfirmedPlanetsSheetForm, DisplayConstellationSheetForm, DisplayMarsPhotosSheetForm, ViewApproachingAsteroidsForm, ViewConfirmedPlanetsForm, ViewConstellationForm, ViewMarsPhotosForm

from data import data_source
from datetime import datetime, timedelta
from dotenv import load_dotenv
from flask import Flask, abort, render_template, redirect, url_for, request
from flask_bootstrap import Bootstrap5
from flask_login import UserMixin, login_user, LoginManager, current_user, logout_user
from flask_sqlalchemy import SQLAlchemy
from functools import wraps  # Used in 'admin_only" decorator function
from flask_wtf import FlaskForm
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from sqlalchemy import Integer, String, Boolean, Float, DateTime, func, distinct, ForeignKey
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, relationship
from typing import List
# from werkzeug.security import check_password_hash
from wtforms import EmailField, SelectField, StringField, SubmitField, TextAreaField, BooleanField, PasswordField
from wtforms.validators import InputRequired, Length, Email
import collections  # Used for sorting items in the constellations dictionary
# import email_validator
import glob
import math
import os
import smtplib
import time
import traceback
import unidecode
import wx
import wx.lib.agw.pybusyinfo as PBI
# import xlsxwriter
import re
import html2text

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
    return db.session.execute(db.select(Users).where(Users.id == user_id)).scalar()


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
@app.route('/')
def home():
    global db, app, dlg

    try:
        # Go to the home page:
        return render_template("index.html", logged_in=current_user.is_authenticated, recognition_web_template=recognition["web_template"])

    except:
        dlg = wx.MessageBox(f"Error (route: '/'): {traceback.format_exc()}", 'Error', wx.OK | wx.ICON_INFORMATION)
        update_system_log("route: '/'", traceback.format_exc())
        dlg = None


# Configure route for "About" web page:
@app.route('/about')
def about():
    global db, app

    try:
        # Go to the "About" page:
        return render_template("about.html", recognition_web_template=recognition["web_template"])
    except:
        dlg = wx.MessageBox(f"Error (route: '/about'): {traceback.format_exc()}", 'Error', wx.OK | wx.ICON_INFORMATION)
        update_system_log("route: '/about'", traceback.format_exc())
        dlg = None


# Configure route for "Administrative Update Login" web page:
@app.route('/admin_login',methods=["GET", "POST"])
def admin_login():
    global db, app

    try:
        # Instantiate an instance of the "AdminLoginForm" class:
        form = AdminLoginForm()

        # Validate form entries upon submittal. If validated, send message:
        if form.validate_on_submit():
            # Capture the supplied e-mail address and password:
            username = form.txt_username.data
            password = form.txt_password.data

            # Check if user account exists under the supplied username:
            user = db.session.execute(db.select(Users).where(Users.username == username)).scalar()
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
        dlg = wx.MessageBox(f"Error (route: '/admin_login'): {traceback.format_exc()}", 'Error', wx.OK | wx.ICON_INFORMATION)
        update_system_log("route: '/admin_login'", traceback.format_exc())
        dlg = None


# Configure route for logging out of "Administrative Update":
@app.route('/admin_logout')
def admin_logout():
    try:
        # Log user out:
        logout_user()

        # Go to the home page:
        return redirect(url_for('home'))

    except:  # An error has occurred.
        dlg = wx.MessageBox(f"Error (route: '/admin_logout'): {traceback.format_exc()}", 'Error', wx.OK | wx.ICON_INFORMATION)
        update_system_log("route: '/admin_logout'", traceback.format_exc())
        dlg = None


# Configure route for "Administrative Update" web page:
@app.route('/admin_update',methods=["GET", "POST"])
@admin_only
def admin_update():
    global db, app

    try:
        # Instantiate an instance of the "AdminUpdateForm" class:
        form = AdminUpdateForm()

        # Validate form entries upon submittal. Depending on the choices made via the form, perform additional processing:
        if form.validate_on_submit():
            # Initialize variables necessary for capturing errors encountered during desired update:
            update_status_approaching_asteroids = ""
            update_status_confirmed_planets = ""
            update_status_constellations = ""
            update_status_mars_photos = ""

            # Initiate a variable to support dialog for keeping user informed on update status:
            app_wx = wx.App(redirect=False)

            # Execute selected updates:
            if form.chk_approaching_asteroids.data:  # Update to "approaching asteroids" is desired.
                # Get results of obtaining and processing the desired information (use window dialog to keep user informed):
                dlg = PBI.PyBusyInfo("Approaching Asteroids: Update in progress...", title="Administrative Update")
                error_msg_approaching_asteroids, success_approaching_asteroids = get_approaching_asteroids()
                dlg = None
                if success_approaching_asteroids:
                    update_status_approaching_asteroids = "Approaching Asteroids: Successfully updated."
                else:
                    update_status_approaching_asteroids = f"Approaching Asteroids: Update failed ({error_msg_approaching_asteroids})."

            if form.chk_confirmed_planets.data:  # Update to "confirmed planets" is desired.
                # Get results of obtaining and processing the desired information (use window dialog to keep user informed):
                dlg = PBI.PyBusyInfo("Confirmed Planets: Update in progress...", title="Administrative Update")
                error_msg_confirmed_planets, success_confirmed_planets = get_confirmed_planets()
                dlg = None
                if success_confirmed_planets:
                    update_status_confirmed_planets = "Confirmed Planets: Successfully updated."
                else:
                    update_status_confirmed_planets = f"Confirmed Planets: Update failed ({error_msg_confirmed_planets})."

            if form.chk_constellations.data:
                # Get results of obtaining and processing the desired information (use window dialog to keep user informed):
                dlg = PBI.PyBusyInfo("Constellations: Update in progress...", title="Administrative Update")
                error_msg_constellations, success_constellations = get_constellation_data()
                dlg = None
                if success_constellations:
                    update_status_constellations = "Constellations: Successfully updated."
                else:
                    update_status_constellations = f"Constellations: Update failed ({error_msg_constellations})."

            if form.chk_mars_photos.data:
                dlg = PBI.PyBusyInfo("Photos from Mars: Update in progress...", title="Administrative Update")
                error_msg_mars_photos, success_mars_photos = get_mars_photos()
                dlg = None
                if success_mars_photos:
                    update_status_mars_photos = "Photos from Mars: Successfully updated."
                else:
                    update_status_mars_photos = f"Photos from Mars: Update failed ({error_msg_mars_photos})."

            # Prepare final status-update text to pass to the "Administrative Update" page:
            update_status = ""

            # Check if the user has selected at least one of the items to update.  If not, prompt user to select one:
            if not (update_status_approaching_asteroids != "" or update_status_confirmed_planets != "" or update_status_constellations != "" or update_status_mars_photos != ""):
                dlg = wx.MessageBox(f"Please select at least one of the items to update.", 'Administrative Update', wx.OK | wx.ICON_INFORMATION)
                dlg = None

            else:  # User has selected at least one item to update.  Perform selected update(s):
                if update_status_approaching_asteroids != "":
                    update_status += update_status_approaching_asteroids + "\n"
                if update_status_confirmed_planets != "":
                    update_status += update_status_confirmed_planets + "\n"
                if update_status_constellations != "":
                    update_status += update_status_constellations + "\n"
                if update_status_mars_photos != "":
                    update_status += update_status_mars_photos + "\n"

                # Destroy dialog app:
                app_wx.MainLoop()

                # Go to the "Administrative Update" page and display the results of update execution:
                return render_template("admin_update.html", update_status=update_status, recognition_web_template=recognition["web_template"])

        # Go to the "Administrative Update" page:
        return render_template("admin_update.html", form=form, update_status="<<Update Choices to be Made.>>", recognition_web_template=recognition["web_template"])

    except:  # An error has occurred.
        dlg = wx.MessageBox(f"Error (route: '/admin_update'): {traceback.format_exc()}", 'Error', wx.OK | wx.ICON_INFORMATION)
        update_system_log("route: '/admin_update'", traceback.format_exc())
        dlg = None


# # Configure route for "Approaching Asteroids" web page:
# @app.route('/approaching_asteroids',methods=["GET", "POST"])
# def approaching_asteroids():
#     global db, app
#
#     try:
#         # Instantiate an instance of the "ViewApproachingAsteroidsForm" class:
#         form = ViewApproachingAsteroidsForm()
#
#         # Instantiate an instance of the "DisplayApproachingAsteroidsSheetForm" class:
#         form_ss = DisplayApproachingAsteroidsSheetForm()
#
#         # Populate the close approach date listbox with an ordered list of close approach dates represented in the database:
#         list_close_approach_dates = []
#         close_approach_dates = db.session.query(distinct(ApproachingAsteroids.close_approach_date)).order_by(ApproachingAsteroids.close_approach_date).all()
#         for close_approach_date in close_approach_dates:
#             list_close_approach_dates.append(str(close_approach_date)[2:12])
#         form.list_close_approach_date.choices = list_close_approach_dates
#
#         # Populate the approaching-asteroids sheet file listbox with the sole sheet viewable in this scope:
#         form_ss.list_approaching_asteroids_sheet_name.choices = ["ApproachingAsteroids.xlsx"]
#
#         # Validate form entries upon submittal. Depending on the form involved, perform additional processing:
#         if form.validate_on_submit():
#             if form.list_close_approach_date.data != None:
#                 error_msg = ""
#                 # Retrieve the record from the database which pertains to confirmed planets discovered in the selected year:
#                 approaching_asteroids_details = retrieve_from_database(trans_type="approaching_asteroids_by_close_approach_date", close_approach_date=form.list_close_approach_date.data)
#
#                 if approaching_asteroids_details == {}:
#                     error_msg = "Error: Data could not be obtained at this time."
#                 elif approaching_asteroids_details == []:
#                     error_msg = "No matching records were retrieved."
#
#                 # Show web page with retrieved approaching-asteroid details:
#                 return render_template('show_approaching_asteroids_details.html', approaching_asteroids_details=approaching_asteroids_details, close_approach_date=form.list_close_approach_date.data, error_msg=error_msg, recognition_scope_specific=recognition["approaching_asteroids"], recognition_web_template=recognition["web_template"])
#
#             else:
#                 # Open the selected spreadsheet file:
#                 os.startfile(str(form_ss.list_approaching_asteroids_sheet_name.data))
#
#         # Go to the web page to render the results:
#         return render_template('approaching_asteroids.html', form=form, form_ss=form_ss, recognition_scope_specific=recognition["approaching_asteroids"], recognition_web_template=recognition["web_template"])
#
#     except:  # An error has occurred.
#         dlg = wx.MessageBox(f"Error (route: '/approaching_asteroids'): {traceback.format_exc()}", 'Error', wx.OK | wx.ICON_INFORMATION)
#         update_system_log("route: '/approaching_asteroids'", traceback.format_exc())
#         dlg = None
#
#
# # Configure route for "Astronomy Pic of the Day" web page:
# @app.route('/astronomy_pic_of_day')
# def astronomy_pic_of_day():
#     global db, app
#
#     try:
#         # Get details re: the astronomy picture of the day:
#         json, copyright_details, error_msg = get_astronomy_pic_of_the_day()
#
#         # Go to the web page to render the results:
#         return render_template("astronomy_pic_of_day.html", json=json, copyright_details=copyright_details, error_msg=error_msg, recognition_scope_specific=recognition["astronomy_pic_of_day"], recognition_web_template=recognition["web_template"])
#
#     except:  # An error has occurred.
#         dlg = wx.MessageBox(f"Error (route: '/astronomy_pic_of_day'): {traceback.format_exc()}", 'Error', wx.OK | wx.ICON_INFORMATION)
#         update_system_log("route: '/astronomy_pic_of_day'", traceback.format_exc())
#         dlg = None
#
#
# # Configure route for "Confirmed Planets" web page:
# @app.route('/confirmed_planets',methods=["GET", "POST"])
# def confirmed_planets():
#     global db, app
#
#     try:
#         # Instantiate an instance of the "ViewConstellationForm" class:
#         form = ViewConfirmedPlanetsForm()
#
#         # Instantiate an instance of the "DisplayConfirmedPlanetsSheetForm" class:
#         form_ss = DisplayConfirmedPlanetsSheetForm()
#
#         # Populate the discovery year listbox with an ordered (descending) list of discovery years represented in the database:
#         list_discovery_years = []
#         discovery_years = db.session.query(distinct(ConfirmedPlanets.discovery_year)).order_by(ConfirmedPlanets.discovery_year.desc()).all()
#         for year in discovery_years:
#             list_discovery_years.append(int(str(year)[1:5]))
#         form.list_discovery_year.choices = list_discovery_years
#
#         # Populate the confirmed planets sheet file listbox with the sole sheet viewable in this scope:
#         form_ss.list_confirmed_planets_sheet_name.choices = ["ConfirmedPlanets.xlsx"]
#
#         # Validate form entries upon submittal. Depending on the form involved, perform additional processing:
#         if form.validate_on_submit():
#             if form.list_discovery_year.data != None:
#                 error_msg = ""
#                 # Retrieve the record from the database which pertains to confirmed planets discovered in the selected year:
#                 confirmed_planets_details = retrieve_from_database(trans_type="confirmed_planets_by_disc_year", disc_year=form.list_discovery_year.data)
#
#                 if confirmed_planets_details == {}:
#                     error_msg = "Error: Data could not be obtained at this time."
#                 elif confirmed_planets_details == []:
#                     error_msg = "No matching records were retrieved."
#
#                 # Show web page with retrieved confirmed-planet details:
#                 return render_template('show_confirmed_planets_details.html', confirmed_planets_details=confirmed_planets_details, disc_year=form.list_discovery_year.data, error_msg=error_msg, recognition_scope_specific=recognition["confirmed_planets"], recognition_web_template=recognition["web_template"])
#
#             else:
#                 # Open the selected spreadsheet file:
#                 os.startfile(str(form_ss.list_confirmed_planets_sheet_name.data))
#
#         # Go to the web page to render the results:
#         return render_template('confirmed_planets.html', form=form, form_ss=form_ss, recognition_scope_specific=recognition["confirmed_planets"], recognition_web_template=recognition["web_template"])
#
#     except:  # An error has occurred.
#         dlg = wx.MessageBox(f"Error (route: '/confirmed_planets'): {traceback.format_exc()}", 'Error', wx.OK | wx.ICON_INFORMATION)
#         update_system_log("route: '/confirmed_planets'", traceback.format_exc())
#         dlg = None
#
#
# # Configure route for "Constellations" web page:
# @app.route('/constellations',methods=["GET", "POST"])
# def constellations():
#     global db, app
#
#     try:
#         # Instantiate an instance of the "ViewConstellationForm" class:
#         form = ViewConstellationForm()
#
#         # Instantiate an instance of the "DisplayConstellationSheetForm" class:
#         form_ss = DisplayConstellationSheetForm()
#
#         # Populate the constellation name listbox with an ordered list of constellation names from the database:
#         form.list_constellation_name.choices = db.session.execute(db.select(Constellations.name + " (" + Constellations.nickname + ")").order_by(Constellations.name)).scalars().all()
#
#         # Populate the constellation sheet file listbox with the sole sheet viewable in this scope:
#         form_ss.list_constellation_sheet_name.choices = ["Constellations.xlsx"]
#
#         # Validate form entries upon submittal. Depending on the form involved, perform additional processing:
#         if form.validate_on_submit():
#
#             if form.list_constellation_name.data != None:
#                 # Capture selected constellation name:
#                 selected_constellation_name = form.list_constellation_name.data.split("(")[0][:len(form.list_constellation_name.data.split("(")[0])-1]
#
#                 # Retrieve the record from the database which pertains to the selected constellation name:
#                 constellation_details = db.session.execute(db.select(Constellations).where(Constellations.name == selected_constellation_name)).scalar()
#
#                 # Show web page with retrieved constellation details:
#                 return render_template('show_constellation_details.html', constellation_details=constellation_details, recognition_scope_specific=recognition["constellations"], recognition_web_template=recognition["web_template"])
#
#             else:
#                 # Open the selected spreadsheet file:
#                 os.startfile(str(form_ss.list_constellation_sheet_name.data))
#
#         # Go to the web page to render the results:
#         return render_template('constellations.html', form=form, form_ss=form_ss, recognition_scope_specific=recognition["constellations"], recognition_web_template=recognition["web_template"])
#
#     except:  # An error has occurred.
#         dlg = wx.MessageBox(f"Error (route: '/constellations'): {traceback.format_exc()}", 'Error', wx.OK | wx.ICON_INFORMATION)
#         update_system_log("route: '/constellations'", traceback.format_exc())
#         dlg = None
#
#
# Configure route for "Contact Us" web page:
@app.route('/contact',methods=["GET", "POST"])
def contact():
    global db, app

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
        dlg = wx.MessageBox(f"Error (route: '/contact'): {traceback.format_exc()}", 'Error', wx.OK | wx.ICON_INFORMATION)
        update_system_log("route: '/contact'", traceback.format_exc())
        dlg = None


# # Configure route for "Photos from Mars" web page:
# @app.route('/mars_photos',methods=["GET", "POST"])
# def mars_photos():
#     global db, app
#
#     try:
#         # Instantiate an instance of the "ViewConstellationForm" class:
#         form = ViewMarsPhotosForm()
#
#         # Instantiate an instance of the "DisplayMarsPhotosSheetForm" class:
#         form_ss = DisplayMarsPhotosSheetForm()
#
#         # Populate the rover name / earth date combo listbox with an ordered list of such combinations:
#         list_rover_earth_date_combos = []
#         rover_earth_date_combos = db.session.query(distinct(MarsPhotosAvailable.rover_earth_date_combo)).order_by(MarsPhotosAvailable.rover_name, MarsPhotosAvailable.earth_date.desc()).all()
#         for rover_earth_date_combo in rover_earth_date_combos:
#             list_rover_earth_date_combos.append(str(rover_earth_date_combo).split("'")[1])
#         form.list_rover_earth_date_combo.choices = list_rover_earth_date_combos
#
#         # Populate the Mars photos sheet file listbox with all filenames of spreadsheets pertinent to this scope:
#         form_ss.list_mars_photos_sheet_name.choices = glob.glob("Mars Photos*.xlsx")
#
#         # Validate form entries upon submittal. Depending on the form involved, perform additional processing:
#         if form.validate_on_submit():
#             if form.list_rover_earth_date_combo.data != None:
#                 error_msg = ""
#                 # Retrieve the record from the database which pertains to Mars photos taken via the selected rover / earth date combo:
#                 mars_photos_details = retrieve_from_database(trans_type="mars_photos_by_rover_earth_date_combo", rover_earth_date_combo=form.list_rover_earth_date_combo.data)
#
#                 if mars_photos_details == {}:
#                     error_msg = "Error: Data could not be obtained at this time."
#                 elif mars_photos_details == []:
#                     error_msg = "No matching records were retrieved."
#
#                 # Show web page with retrieved photo details:
#                 return render_template('show_mars_photos_details.html', mars_photos_details=mars_photos_details, rover_earth_date_combo=form.list_rover_earth_date_combo.data, error_msg=error_msg, recognition_scope_specific=recognition["mars_photos"], recognition_web_template=recognition["web_template"])
#
#             else:
#                 # Open the selected spreadsheet file:
#                 os.startfile(str(form_ss.list_mars_photos_sheet_name.data))
#
#         # Go to the web page to render the results:
#         return render_template('mars_photos.html', form=form, form_ss=form_ss, recognition_scope_specific=recognition["mars_photos"], recognition_web_template=recognition["web_template"])
#
#     except:  # An error has occurred.
#         dlg = wx.MessageBox(f"Error (route: '/mars_photos'): {traceback.format_exc()}", 'Error', wx.OK | wx.ICON_INFORMATION)
#         update_system_log("route: '/mars_photos'", traceback.format_exc())
#         dlg = None
#
#
# # Configure route for "Space News" web page:
# @app.route('/space_news')
# def space_news():
#     global db, app
#
#     try:
#         # Get results of obtaining and processing the desired information:
#         success, error_msg = get_space_news()
#
#         if success:
#             # Query the table for space news articles:
#             with app.app_context():
#                 articles = db.session.execute(db.select(SpaceNews).order_by(SpaceNews.row_id)).scalars().all()
#                 if articles.count == 0:
#                     success = False
#                     error_msg = "Error: Cannot retrieve article data from database."
#
#         else:
#             articles = None
#
#         # Go to the web page to render the results:
#         return render_template("space_news.html", articles=articles, success=success, error_msg=error_msg, recognition_scope_specific=recognition["space_news"], recognition_web_template=recognition["web_template"])
#
#     except:  # An error has occurred.
#         dlg = wx.MessageBox(f"Error (route: '/space_news'): {traceback.format_exc()}", 'Error', wx.OK | wx.ICON_INFORMATION)
#         update_system_log("route: '/space_news'", traceback.format_exc())
#         dlg = None
#
#
# # Configure route for "Where is ISS" web page:
# @app.route('/where_is_iss')
# def where_is_iss():
#     global db, app
#
#     try:
#         # Get ISS's current location along with a URL to get a map plotting said location:
#         location_address, location_url = get_iss_location()
#
#         # Go to the web page to render the results:
#         return render_template("where_is_iss.html", location_address=location_address, location_url=location_url, has_url=not(location_url == ""), recognition_scope_specific=recognition["where_is_iss"], recognition_web_template=recognition["web_template"])
#
#     except:  # An error has occurred.
#         dlg = wx.MessageBox(f"Error (route: '/where_is_iss'): {traceback.format_exc()}", 'Error', wx.OK | wx.ICON_INFORMATION)
#         update_system_log("route: '/where_is_iss'", traceback.format_exc())
#         dlg = None
#
#
# # Configure route for "Who is in Space Now" web page:
# @app.route('/who_is_in_space_now')
# def who_is_in_space_now():
#     global db, app
#
#     try:
#         # Get results of obtaining a JSON with the desired information:
#         json, has_json = get_people_in_space_now()
#
#         # Go to the web page to render the results:
#         return render_template("who_is_in_space_now.html", json=json, has_json=has_json, recognition_scope_specific=recognition["who_is_in_space_now"], recognition_web_template=recognition["web_template"])
#
#     except:  # An error has occurred.
#         dlg = wx.MessageBox(f"Error (route: '/who_is_in_space_now'): {traceback.format_exc()}", 'Error', wx.OK | wx.ICON_INFORMATION)
#         update_system_log("route: '/who_is_in_space_now'", traceback.format_exc())
#         dlg = None
#

# DEFINE FUNCTIONS TO BE USED FOR THIS APPLICATION (LISTED IN ALPHABETICAL ORDER BY FUNCTION NAME):
# *************************************************************************************************
def config_database():
    """Function for configuring the database tables supporting this website"""
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
        update_system_log("config_database", traceback.format_exc())

        # Return failed-execution indication to the calling function:
        return False


def config_web_forms():
    """Function for configuring the web forms supporting this website"""
    global AdminLoginForm, AdminUpdateForm, ContactForm, DisplayApproachingAsteroidsSheetForm, DisplayConfirmedPlanetsSheetForm, DisplayConstellationSheetForm, DisplayMarsPhotosSheetForm, ViewApproachingAsteroidsForm, ViewConfirmedPlanetsForm, ViewConstellationForm, ViewMarsPhotosForm

    try:
        # CONFIGURE WEB FORMS (LISTED IN ALPHABETICAL ORDER):
        # Configure "admin_update" form:
        class AdminLoginForm(FlaskForm):
            txt_username = StringField(label="Username:", validators=[InputRequired()])
            txt_password = PasswordField(label="Password:", validators=[InputRequired()])
            button_submit = SubmitField(label="Login")

        # Configure "admin_update" form:
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

        # Configure form for viewing "approaching asteroids" spreadsheet:
        class DisplayApproachingAsteroidsSheetForm(FlaskForm):
            list_approaching_asteroids_sheet_name = SelectField("Approaching Asteroids Sheet:", choices=[],
                                                                validate_choice=False)
            button_submit = SubmitField(label="View Approaching Asteroids Spreadsheet")

        # Configure form for viewing "confirmed planets" spreadsheet:
        class DisplayConfirmedPlanetsSheetForm(FlaskForm):
            list_confirmed_planets_sheet_name = SelectField("Confirmed Planets Sheet:", choices=[],
                                                            validate_choice=False)
            button_submit = SubmitField(label="View Confirmed Planets Spreadsheet")

        # Configure form for viewing "constellations" spreadsheet:
        class DisplayConstellationSheetForm(FlaskForm):
            list_constellation_sheet_name = SelectField("Constellation Sheet:", choices=[], validate_choice=False)
            button_submit = SubmitField(label="View Constellations Spreadsheet")

        # Configure form for viewing "Mars photos" spreadsheet (summary or detailed):
        class DisplayMarsPhotosSheetForm(FlaskForm):
            list_mars_photos_sheet_name = SelectField("Mars Photos Sheet:", choices=[], validate_choice=False)
            button_submit = SubmitField(label="View Mars Photos Spreadsheet")

        # Configure form for viewing "approaching asteroids" data online (on dedicated web page):
        class ViewApproachingAsteroidsForm(FlaskForm):
            list_close_approach_date = SelectField("Select Close Approach Date:", choices=[], validate_choice=False)
            button_submit = SubmitField(label="View Details")

        # Configure form for viewing "confirmed planets" data online (on dedicated web page):
        class ViewConfirmedPlanetsForm(FlaskForm):
            list_discovery_year = SelectField("Select Discovery Year:", choices=[], validate_choice=False)
            button_submit = SubmitField(label="View List of Confirmed Planets")

        # Configure form for viewing "constellations" data online (on dedicated web page):
        class ViewConstellationForm(FlaskForm):
            list_constellation_name = SelectField("Select Constellation Name:", choices=[], validate_choice=False)
            button_submit = SubmitField(label="View Details")

        # Configure form for viewing "Mars photos" data online (on dedicated web page):
        class ViewMarsPhotosForm(FlaskForm):
            list_rover_earth_date_combo = SelectField("Select Rover Name / Earth Date Combination:", choices=[],
                                                      validate_choice=False)
            button_submit = SubmitField(label="View List of Photos")

        # At this point, function is presumed to have executed successfully.  Return\
        # successful-execution indication to the calling function:
        return True

    except:  # An error has occurred.
        update_system_log("config_web_forms", traceback.format_exc())

        # Return failed-execution indication to the calling function:
        return False


def email_from_contact_page(form):
    """Function to process a message that user wishes to e-mail from this website to the website administrator."""
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
                    msg=f"Subject: Eye for Space - E-mail from 'Contact Us' page\n\nName: {form.txt_name.data}\nE-mail address: {form.txt_email.data}\n\nMessage:\n{form.txt_message.data}"
                )
                # Return successful-execution message to the calling function::
                return "Your message has been successfully sent."

    except:  # An error has occurred.
        update_system_log("email_from_contact_page", traceback.format_exc())

        # Return failed-execution message to the calling function:
        return "An error has occurred. Your message was not sent."


def find_element(driver, find_type, find_details):
    """Function to find an element via a web-scraping procedure"""
    # NOTE: Error handling is deferred to the calling function:
    if find_type == "xpath":
        return driver.find_element(By.XPATH, find_details)


# def get_approaching_asteroids():
#     """Function that retrieves and processes a list of asteroids based on closest approach to Earth"""
#     # Capture the current date:
#     current_date = datetime.now()
#
#     # Capture the current date + an added window (delta) of the following 7 days:
#     current_date_with_delta = current_date + timedelta(days=7)
#
#     try:
#         # Execute the API request (limit: closest approach <= 7 days from today):
#         response = requests.get(URL_CLOSEST_APPROACH_ASTEROIDS + "?start_date=" + current_date.strftime("%Y-%m-%d") + "&end_date=" + current_date_with_delta.strftime("%Y-%m-%d") + "&api_key=" + API_KEY_CLOSEST_APPROACH_ASTEROIDS)
#
#         # Initialize variable to store collected necessary asteroid data:
#         approaching_asteroids = []
#
#         # If the API request was successful, display the results:
#         if response.status_code == 200:  # API request was successful.
#
#             # Capture desired fields from the returned JSON:
#             for key in response.json()["near_earth_objects"]:
#                 for asteroid in response.json()["near_earth_objects"][key]:
#                     asteroid_dict = {
#                         "id": asteroid["id"],
#                         "name": asteroid["name"],
#                         "absolute_magnitude_h": asteroid["absolute_magnitude_h"],
#                         "estimated_diameter_km_min": asteroid["estimated_diameter"]["kilometers"]["estimated_diameter_min"],
#                         "estimated_diameter_km_max": asteroid["estimated_diameter"]["kilometers"]["estimated_diameter_max"],
#                         "is_potentially_hazardous": asteroid["is_potentially_hazardous_asteroid"],
#                         "close_approach_date": asteroid["close_approach_data"][0]["close_approach_date"],
#                         "relative_velocity_km_per_s": asteroid["close_approach_data"][0]["relative_velocity"]["kilometers_per_second"],
#                         "miss_distance_km": asteroid["close_approach_data"][0]["miss_distance"]["kilometers"],
#                         "orbiting_body": asteroid["close_approach_data"][0]["orbiting_body"],
#                         "is_sentry_object": asteroid["is_sentry_object"],
#                         "url": asteroid["nasa_jpl_url"]
#                         }
#
#                     # Add captured data for each asteroid (as a dictionary) to the "approaching_asteroids" list:
#                     approaching_asteroids.append(asteroid_dict)
#
#             # Delete the existing records in the "approaching_asteroids" database table and update same with
#             # the up-to-date data (from the JSON).  If an error occurred, update system log and return a
#             # failed-execution indication to the calling function:
#             if not update_database("update_approaching_asteroids", approaching_asteroids):
#                 update_system_log("get_approaching_asteroids", "Error: Database could not be updated. Data cannot be obtained at this time.")
#                 return "Error: Database could not be updated. Data cannot be obtained at this time.", False
#
#             # Retrieve all existing records in the "approaching_asteroids" database table. If the function
#             # called returns an empty directory, update system log and return a failed-execution indication
#             # to the calling function:
#             asteroids_data = retrieve_from_database("approaching_asteroids")
#             if asteroids_data == {}:
#                 update_system_log("get_approaching_asteroids", "Error: Data cannot be obtained at this time.")
#                 return "Error: Data cannot be obtained at this time.", False
#
#             # If an empty list was returned, no records satisfied the query.  Therefore, update system log and
#             # return a failed-execution indication to the calling function:
#             elif asteroids_data == []:
#                 update_system_log("get_approaching_asteroids", "No matching records were retrieved.")
#                 return "No matching records were retrieved.", False
#
#             # Create and format a spreadsheet file (workbook) to contain all asteroids data. If execution failed,
#             # update system log and return failed-execution indication to the calling function:
#             if not export_data_to_spreadsheet_standard("approaching_asteroids", asteroids_data):
#                 update_system_log("get_approaching_asteroids", "Error: Spreadsheet creation could not be completed at this time.")
#                 return "Error: Spreadsheet creation could not be completed at this time.", False
#
#             # At this point, function is deemed to have executed successfully.  Update system log and
#             # return successful-execution indication to the calling function:
#             update_system_log("get_approaching_asteroids", "Successfully updated.")
#             return "", True
#
#         else:  # API request failed. Update system log and return failed-execution indication to the calling function:
#             update_system_log("get_approaching_asteroids", "Error: API request failed. Data cannot be obtained at this time.")
#             return "Error: API request failed. Data cannot be obtained at this time.", False
#
#     except:  # An error has occurred.
#         update_system_log("get_approaching_asteroids", traceback.format_exc())
#
#         # Return failed-execution indication to the calling function:
#         return "An error has occurred. Data cannot be obtained at this time.", False
#
#
# def get_confirmed_planets():
#     """Function for getting all needed data pertaining to confirmed planets and store such information in the space database supporting our website"""
#     try:
#         # Execute API request:
#         response = requests.get(URL_CONFIRMED_PLANETS)
#         if response.status_code == 200:
#             # Delete the existing records in the "confirmed_planets" database table and update same with
#             # the up-to-date data (from the JSON).  If execution failed, update system log and return
#             # failed-execution indication to the calling function::
#             # NOTE:  Scope of data: Solution Type = 'Published Confirmed'
#             if not update_database("update_confirmed_planets", response.json()):
#                 update_system_log("get_confirmed_planets", "Error: Database could not be updated. Data cannot be obtained at this time.")
#                 return "Error: Database could not be updated. Data cannot be obtained at this time.", False
#
#             # Retrieve all existing records in the "confirmed_planets" database table. If the function
#             # called returns an empty directory, update system log and return failed-execution indication
#             # to the calling function:
#             confirmed_planets_data = retrieve_from_database("confirmed_planets")
#             if confirmed_planets_data == {}:
#                 update_system_log("get_confirmed_planets", "Error: Data cannot be obtained at this time.")
#                 return "Error: Data cannot be obtained at this time.", False
#
#             # If an empty list was returned, no records satisfied the query.  Therefore, update system log and return
#             # failed-execution indication to the calling function:
#             elif confirmed_planets_data == []:
#                 update_system_log("get_confirmed_planets", "No matching records were retrieved.")
#                 return "No matching records were retrieved.", False
#
#             # Create and format a spreadsheet file (workbook) to contain all confirmed-planet data. If execution
#             # failed, update system log and return failed-execution indication to the calling function:
#             if not export_data_to_spreadsheet_standard("confirmed_planets", confirmed_planets_data):
#                 update_system_log("get_confirmed_planets", "Error: Spreadsheet creation could not be completed at this time.")
#                 return "Error: Spreadsheet creation could not be completed at this time.", False
#
#             # At this point, function is deemed to have executed successfully.  Update system log and return
#             # successful-execution indication to the calling function:
#             update_system_log("get_confirmed_planets", "Successfully updated.")
#             return "", True
#
#         else:  # API request failed.  Update system log and return failed-execution indication to the calling function:
#             update_system_log("get_confirmed_planets","Error: API request failed. Data cannot be obtained at this time.")
#             return "Error: API request failed. Data cannot be obtained at this time.", False
#
#     except:  # An error has occurred.
#         update_system_log("get_confirmed_planets", traceback.format_exc())
#
#         # Return failed-execution indication to the calling function:
#         return "An error has occurred. Data cannot be obtained at this time.", False
#
#
#
def get_inspirational_data():
    """Function for getting (via web-scraping) inspirational data and storing such information in the database supporting our website"""

    try:
        data_sources_to_update = retrieve_from_database(trans_type="get_non-static_data_sources")
        print(len(data_sources_to_update))
        if data_sources_to_update == {}:
            error_msg = "Error: Data could not be obtained at this time."
        elif data_sources_to_update == []:
            error_msg = "No matching records were retrieved."

        # for i in range(0, len(data_source)):
        for i in range(0, len(data_sources_to_update)):

            # Get the inspirational data from each identified data source.
            # If the function called returns an empty directory,
            inspirational_data = get_inspirational_data_details(data_sources_to_update[i].id, data_sources_to_update[i].count, data_sources_to_update[i].name, data_sources_to_update[i].url)
            if inspirational_data == []:
                update_system_log("get_inspirational_data", f"Error: Data (for source = '{data_sources_to_update[i].name}') cannot be obtained at this time.")
                return f"Error: Data (for source = {data_sources_to_update[i].name}) cannot be obtained at this time.", False

            # Delete the existing records in the "inspirational_quotes" database table and update same with the
            # contents that were obtained via the previous step.  If the function called returns a failed-execution
            # indication, update system log and return failed-execution indication to the calling function:
            if not update_database("update_inspirational_quotes", inspirational_data, source=data_sources_to_update[i].id):
                update_system_log("get_constellation_data",
                                  "Error: Database could not be updated. Data cannot be obtained at this time.")
                return "Error: Database could not be updated. Data cannot be obtained at this time.", False

            # # Retrieve all existing records in the "constellations" database table. If the function
            # # called returns an empty directory, update system log and return failed-execution indication to the
            # # calling function:
            # constellations_data = retrieve_from_database("constellations")
            # if constellations_data == {}:
            #     update_system_log("get_constellation_data", "Error: Data cannot be obtained at this time.")
            #     return "Error: Data cannot be obtained at this time.", False
            #
            # # Create and format a spreadsheet file (workbook) to contain all constellation data. If the function called returns
            # # a failed-execution indication, update system log and return a failed-execution indication to the calling function:
            # if not export_data_to_spreadsheet_standard("constellations", constellations_data):
            #     update_system_log("get_constellation_data",
            #                       "Error: Spreadsheet creation could not be completed at this time.")
            #     return "Error: Spreadsheet creation could not be completed at this time.", False

        print("FINISHED")

        # At this point, function is deemed to have executed successfully.  Update system log and return
        # successful-execution indication to the calling function:
        update_system_log("get_inspirational_data", "Successfully updated.")
        return "", True

        # else:  # An error has occurred in processing constellation data.
        #     update_system_log("get_constellation_data", "Error: Data cannot be obtained at this time.")
        #     return "Error: Data cannot be obtained at this time.", False

    except:  # An error has occurred.
        update_system_log("get_inspirational_data", traceback.format_exc())

        # Return failed-execution indication to the calling function:
        return "An error has occurred. Data cannot be obtained at this time.", False


def get_inspirational_data_details(source, count, name, url):
    """Function for getting (via web-scraping) the nickname for each constellation identified"""
    global data_source

    print(source)
    print(count)

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

        if source == 1 or source == 2 or source == 3 or source == 13 or source == 21 or source == 22 or source == 24:
            for i in range(1, count + 20):
                try:
                    element_quote = find_element(driver, "xpath",'/html/body/div[4]/div[2]/main/article/div/blockquote[' + str(i) + ']/p')

                    # Add the quote to the "inspirational_data" list:
                    inspirational_data.append(element_quote.text)

                except:
                    continue

        elif source == 4:
            for i in range(1, 21):
                for j in range(1, 21):
                    try:
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
                    element_quote = find_element(driver, "xpath",
                                                 '/html/body/div[2]/div/div/div/div[2]/div/article/div/div/div[2]/p[' + str(i) + ']')

                    # If item is an inspiration quote, strip it of unwanted characters and add it to the list for subsequent update of the database.
                    element_quote_1 = element_quote.text.strip()
                    # print(f"regular1: {element_quote_1[0] == '“' or element_quote_1[0] == '"'}: {element_quote_1}")
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
                        element_quote = find_element(driver, "xpath",
                                                     '/html/body/div[3]/div/div/div/div[2]/div/article/div/div/div[2]/p[' + str(i) + ']')

                        # If item is an inspiration quote, strip it of unwanted characters and add it to the list for subsequent update of the database.
                        element_quote_1 = element_quote.text.strip()
                        # print(f"regular2: {element_quote_1[0] == '“' or element_quote_1[0] == '"'}: {element_quote_1}")
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
                    element_quote = find_element(driver, "xpath",
                                                 '/html/body/div[2]/div/div/div/div[2]/div/article/div/div/div[2]/blockquote[' + str(i) + ']/p')

                    # If item is an inspiration quote, strip it of unwanted characters and add it to the list for subsequent update of the database.
                    element_quote_1 = element_quote.text.strip()
                    # print(f"block1: {element_quote_1[0] == '“' or element_quote_1[0] == '"'}: {element_quote_1}")
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
                        element_quote = find_element(driver, "xpath",
                                                     '/html/body/div[3]/div/div/div/div[2]/div/article/div/div/div[2]/blockquote[' + str(i) + ']/p')

                        # If item is an inspiration quote, strip it of unwanted characters and add it to the list for subsequent update of the database.
                        element_quote_1 = element_quote.text.strip()
                        # print(f"block2: {element_quote_1[0] == '“' or element_quote_1[0] == '"'}: {element_quote_1}")
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
                    element_quote = find_element(driver, "xpath",
                                                 '/html/body/div[2]/div/div/div/div[2]/div/article/div/div/div[2]/div[' + str(i) + ']/div[1]/a')

                    # If item is an inspiration quote, strip it of unwanted characters and add it to the list for subsequent update of the database.
                    element_quote_1 = element_quote.text.strip()
                    # print(f"image1: {element_quote_1[0] == '“' or element_quote_1[0] == '"'}: {element_quote_1}")
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
                        element_quote = find_element(driver, "xpath",
                                                     '/html/body/div[3]/div/div/div/div[2]/div/article/div/div/div[2]/div[' + str(
                                                         i) + ']/div[1]/a')

                        # If item is an inspiration quote, strip it of unwanted characters and add it to the list for subsequent update of the database.
                        element_quote_1 = element_quote.text.strip()
                        # print(f"image2: {element_quote_1[0] == '“' or element_quote_1[0] == '"'}: {element_quote_1}")
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
            for i in range(1, count + 20):
                try:
                    element_quote = find_element(driver, "xpath",
                                                 '/html/body/div[1]/div[1]/div[1]/div[1]/div[2]/main/article/div/div/div/div/div[3]/div/div[2]/div[1]/blockquote[' + str(i) + ']/p')

                    # If item is an inspiration quote, strip it of unwanted characters and add it to the list for subsequent update of the database.
                    element_quote_1 = element_quote.text.strip()
                    # print(f"block: {element_quote_1[0] == '“' or element_quote_1[0] == '"'}: {element_quote_1}")
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
                    element_quote = find_element(driver, "xpath",'/html/body/div[1]/div[4]/div[1]/div[2]/div/div[1]/p[' + str(i) + ']')

                    element_quote_1 = html2text.html2text(element_quote.get_attribute("outerHTML"))

                    # print(element_quote.get_attribute("outerHTML"))

                    if not (":" in element_quote_1 and not ("_" in element_quote_1)):
                        pass  # Item will not be added to the database.
                    else:
                        # Strip unwanted characters:
                        element_quote_2 = element_quote_1.replace("*", "")
                        element_quote_3 = element_quote_2.replace("\n", " ")
                        element_quote_4 = element_quote_3.strip()

                        # Add the quote to the "inspirational_data" list:
                        # print(f"Block 1: {element_quote_4}")
                        inspirational_data.append(element_quote_4)

                except:
                    continue

            for i in range(1, 21):
                for j in range(1, 61):
                    try:
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
                                    # element_quote_8 = element_quote_8.replace("()", "")

                        element_quote_9 = element_quote_8.replace("()", "")
                        # element_quote_9 = element_quote_8

                        element_quote_10 = element_quote_9.strip()

                        if element_quote_10 == "" or ":" not in element_quote_10 or len(element_quote_10) == 1:
                            pass  # Item will not be added to the database.
                        else:
                            # print(f"Block 2: {element_quote_11}")

                            # Add the quote to the "inspirational_data" list:
                            inspirational_data.append(element_quote_10)

                    except:
                        continue

        elif source == 15:
            for i in range(1, count + 20):
                try:
                    element_quote = find_element(driver, "xpath",
                                                 '/html/body/main/article/div[3]/blockquote[' + str(i) + ']/p')

                    # Add the quote to the "inspirational_data" list:
                    inspirational_data.append(element_quote.text)

                except:
                    continue

        elif source == 16 or source == 17 or source == 18:
            for i in range(1, count + 20):
                try:
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
                        # print(element_quote_4)

                        # Add the quote to the "inspirational_data" list:
                        inspirational_data.append(element_quote_4)

                except:
                    continue

        elif source == 19 or source == 20 or source == 25:
            for i in range(1, count + 20):
                try:
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

                    # print(f"{'\n' in element_quote.get_attribute('outerHTML')}: {element_quote.get_attribute('outerHTML')}")

                    element_quote_2 = element_quote_1c.replace("_", "")
                    element_quote_3 = element_quote_2.replace("\n", " ")
                    # element_quote_4 = element_quote_3.replace("[", "")
                    # element_quote_5 = element_quote_4.replace("]", "")
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

                    # print(element_quote_8.strip())

                    # Add the quote to the "inspirational_data" list:
                    inspirational_data.append(element_quote_8.strip())

                except:
                    continue

        else:
            pass

        print(f"List: {inspirational_data}")
        print(f"Length: {len(inspirational_data)}")

        # Close and delete the Selenium driver object:
        driver.close()
        del driver

        # Return the populated "inspirational_data" list to the calling function:
        return inspirational_data

    except:  # An error has occurred.
        update_system_log("get_inspirational_data_details", traceback.format_exc())

        # Return empty directory as a failed-execution indication to the calling function:
        return []




# def get_space_news():
#     """Function for retrieving the latest space news articles"""
#     try:
#         # Initialize variables to return to calling function:
#         success = True
#         error_message = ""
#
#         # Execute API request:
#         response = requests.get(URL_SPACE_NEWS)
#         if response.status_code == 200:
#             # Delete the existing records in the "space_news" database table and update same with
#             # the newly acquired articles (from the JSON).  If function failed, update system log
#             # and return failed-execution indication to the calling function:
#             if not update_database("update_space_news", response.json()['results']):
#                 update_system_log("update_space_news", "Error: Space news articles cannot be obtained at this time.")
#                 error_message = "Error: Space news articles cannot be obtained at this time."
#                 success = False
#
#         else:  # API request failed. Update system log and return failed-execution calling function:
#             update_system_log("get_space_news", "Error: API request failed. Data cannot be obtained at this time.")
#             error_message = "API request failed. Space news articles cannot be obtained at this time."
#             success = False
#
#     except:  # An error has occurred.
#         update_system_log("get_space_news", traceback.format_exc())
#         error_message = "An error has occurred. Space news articles cannot be obtained at this time."
#         success = False
#
#     finally:
#         # Return results to the calling function:
#         return success, error_message
#


def retrieve_from_database(trans_type, **kwargs):
    """Function to retrieve data from this application's database based on the type of transaction"""
    try:
        with app.app_context():
            if trans_type == "get_categories":
                # Retrieve and return all records from the "categories" database table:
                return db.session.execute(db.select(Categories).order_by(Categories.id)).scalars().all()

            elif trans_type == "get_quotes_for_category":
                # Capture optional argument:
                category_id = kwargs.get("category_id", None)

                # Retrieve and return all records in the "inspirational_quotes" database table where the data_source_id is
                # associated with the category ID passed to this function.  Use an inner join between the
                # "inspirational_quotes" and "inspiration_data_sources" database tables:
                dataset_join = db.session.query(InspirationalQuotes, InspirationDataSources).join(InspirationDataSources, InspirationalQuotes.data_source_id == InspirationDataSources.id).filter(InspirationDataSources.category_id == category_id).all()

                # Retrieve and return all records in the join where the category ID matches the ID passed to this function:
                return dataset_join

            elif trans_type == "get_non-static_data_sources":
                # Retrieve and return all existing records, sorted by ID #, from the "inspiration_data_sources" database table, where the static :
                return db.session.execute(db.select(InspirationDataSources).where(InspirationDataSources.static == 0).order_by(InspirationDataSources.id)).scalars().all()

            elif trans_type == "get_subscribers":
                return db.session.execute(db.select(Subscribers).order_by(Subscribers.name)).scalars().all()

            elif trans_type == "confirmed_planets":
                # Retrieve and return all existing records, sorted by host and planet names. from the "confirmed_planets" database table:
                return db.session.execute(db.select(ConfirmedPlanets).order_by(ConfirmedPlanets.host_name, ConfirmedPlanets.planet_name)).scalars().all()

            elif trans_type == "confirmed_planets_by_disc_year":
                # Capture optional argument:
                disc_year = kwargs.get("disc_year", None)

                # Retrieve and return all existing records, sorted by host and planet names, from the "confirmed_planets" database table where the "discovery_year" field matches the passed parameter:
                return db.session.execute(db.select(ConfirmedPlanets).where(ConfirmedPlanets.discovery_year == disc_year).order_by(ConfirmedPlanets.host_name, ConfirmedPlanets.planet_name)).scalars().all()

            elif trans_type == "constellations":
                # Initialize return variable (dictionary):
                item_to_return = {}

                # Retrieve all existing records from the "constellations" database table:
                constellations_list = db.session.execute(db.select(Constellations)).scalars().all()

                # Populate the "item_to_return" dictionary will all retrieved records from the DB:
                for i in range(0, len(constellations_list)):
                    item_to_return.update({
                        constellations_list[i].name: {
                            "abbreviation": constellations_list[i].abbreviation,
                            "nickname": constellations_list[i].nickname,
                            "url": constellations_list[i].url,
                            "area": constellations_list[i].area,
                            "myth_assoc": constellations_list[i].myth_assoc,
                            "first_appear": constellations_list[i].first_appear,
                            "brightest_star_name": constellations_list[i].brightest_star_name,
                            "brightest_star_url": constellations_list[i].brightest_star_url
                        }
                    })

                # Return the "item to return" dictionary to the calling function:
                return item_to_return

            elif trans_type == "mars_photo_details_compare_with_photos_available":
                # Retrieve all existing records, sorted by rover name/earth date combo and sol, from the "mars_photos_available" database table:
                photos_available_summary = db.session.query(MarsPhotosAvailable).with_entities(MarsPhotosAvailable.rover_earth_date_combo, MarsPhotosAvailable.sol, MarsPhotosAvailable.total_photos).group_by(MarsPhotosAvailable.rover_earth_date_combo, MarsPhotosAvailable.sol).order_by(MarsPhotosAvailable.rover_earth_date_combo, MarsPhotosAvailable.sol).all()

                # Retrieve all existing records, sorted by rover name/earth date combo and sol, from the "mars_photo_details" database table:
                photo_details_summary = db.session.query(MarsPhotoDetails).with_entities(MarsPhotoDetails.rover_earth_date_combo, MarsPhotoDetails.sol,func.count(MarsPhotoDetails.pic_id).label("total_photos")).group_by(MarsPhotoDetails.rover_earth_date_combo, MarsPhotoDetails.sol).order_by(MarsPhotoDetails.rover_earth_date_combo, MarsPhotoDetails.sol).all()

                # Return both retrieved-record lists to the calling function:
                return photos_available_summary, photo_details_summary

            elif trans_type == "mars_photo_details_get_counts_by_rover_and_earth_date":
                # Retrieve and return all existing records, sorted by rover name (asc) and earth date (desc), from the "mars_photos_available" database table:
                return db.session.query(MarsPhotosAvailable).with_entities(MarsPhotosAvailable.rover_name, MarsPhotosAvailable.earth_date, MarsPhotosAvailable.total_photos).group_by(MarsPhotosAvailable.rover_name, MarsPhotosAvailable.earth_date).order_by(MarsPhotosAvailable.rover_name,MarsPhotosAvailable.earth_date.desc()).all()

            elif trans_type == "mars_photo_details":
                # Retrieve and return all existing records, sorted by rover name (asc), earth date (desc), sol (asc), and pic id (asc) from the "mars_photo_details" database table:
                return db.session.execute(db.select(MarsPhotoDetails).order_by(MarsPhotoDetails.rover_name, MarsPhotoDetails.earth_date.desc(), MarsPhotoDetails.sol, MarsPhotoDetails.pic_id)).scalars().all()

            elif trans_type == "mars_photo_details_rover_earth_date_combo":
                # Capture optional arguments:
                rover_name = kwargs.get("rover_name", None)
                earth_date = kwargs.get("earth_date", None)

                # Retrieve and return all existing records, sorted by sol and pic id, from the "mars_photo_details" database table for the rover name and earth date passed to this function:
                return db.session.execute(db.select(MarsPhotoDetails).where(MarsPhotoDetails.rover_earth_date_combo == rover_name + "_" + earth_date).order_by(MarsPhotoDetails.sol, MarsPhotoDetails.pic_id)).scalars().all()

            elif trans_type == "mars_photo_details_rover_earth_date_combo_count":
                # Capture optional arguments:
                rover_name = kwargs.get("rover_name", None)
                earth_date = kwargs.get("earth_date", None)

                # Retrieve all existing records from the "mars_photo_details" database table for the rover name and earth date passed to this function:
                records = db.session.execute(db.select(MarsPhotoDetails).where(MarsPhotoDetails.rover_earth_date_combo == rover_name + "_" + earth_date)).scalars().all()

                # Return the count of retrieved records to the calling function:
                return len(records)

            elif trans_type == "mars_photos_available":
                # Retrieve and return all existing records, sorted by rover name and earth date (latter = descending order) from the "mars_photos_available" database table:
                return db.session.execute(db.select(MarsPhotosAvailable).order_by(MarsPhotosAvailable.rover_name, MarsPhotosAvailable.earth_date.desc())).scalars().all()

            elif trans_type == "mars_photos_by_rover_earth_date_combo":
                # Capture optional argument:
                rover_earth_date_combo = kwargs.get("rover_earth_date_combo", None)

                # Retrieve and return all existing records, sorted by sol and pic id, from the "mars_photo_details" database table where the "rover_earth_date_combo" field matches the passed parameter:
                return db.session.execute(db.select(MarsPhotoDetails).where(MarsPhotoDetails.rover_earth_date_combo == rover_earth_date_combo).order_by(MarsPhotoDetails.sol, MarsPhotoDetails.pic_id)).scalars().all()

            elif trans_type == "mars_rovers":
                # Retrieve and return all existing records, sorted by rover name, from the "mars_rovers" database table where rovers are tagged as active (in terms of data production):
                return db.session.execute(db.select(MarsRovers).where(MarsRovers.active == "Yes").order_by(MarsRovers.rover_name)).scalars().all()

            elif trans_type == "space_news":
                # Retrieve and return all existing records, sorted by article ID, from the "space_news" database table:
                return db.session.execute(db.select(SpaceNews).orderby(SpaceNews.article_id)).scalars().all()

    except:  # An error has occurred.
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
        # failed-execution indication to the calling function::
        if not config_database():
            update_system_log("run_app", "Error: Database configuration failed.")
            return False

        if input("Update DB?").lower() == "y":
            get_inspirational_data()

        #
        # # Configure web forms.  If function failed, update system log and return
        # # failed-execution indication to the calling function::
        # if not config_web_forms():
        #     update_system_log("run_app", "Error: Web forms configuration failed.")
        #     return False

    except:  # An error has occurred.
        update_system_log("run_app", traceback.format_exc())
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
        update_system_log("setup_selenium_driver", traceback.format_exc())

        # Return failed-execution indication to the calling function:
        return None


def update_database(trans_type, item_to_process, **kwargs):
    """Function to update this application's database based on the type of transaction"""
    try:
        with app.app_context():
            if trans_type == "update_inspirational_quotes":
                # Capture optional argument:
                source = kwargs.get("source", None)

                # Delete all records from the "approaching_asteroids" database table for the inspirational data source
                # being processed:
                db.session.execute(db.delete(InspirationalQuotes).where(InspirationalQuotes.data_source_id == source))
                db.session.commit()

                # Upload, to the "approaching_asteroids" database table, all contents of the "item_to_process" parameter:
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

            elif trans_type == "update_confirmed_planets":
                # Capture optional arguments:
                source = kwargs.get("source", None)

                # Delete all records from the "confirmed_planets" database table:
                db.session.execute(db.delete(ConfirmedPlanets))
                db.session.commit()

                # Upload, to the "confirmed_planets" database table, all contents of the "item_to_process" parameter:
                new_records = []
                for i in range(0, len(item_to_process)):
                    new_record = ConfirmedPlanets(
                        host_name=item_to_process[i]["hostname"],
                        host_num_stars=item_to_process[i]["sy_snum"],
                        host_num_planets=item_to_process[i]["sy_pnum"],
                        planet_name=item_to_process[i]["pl_name"],
                        discovery_year=item_to_process[i]["disc_year"],
                        discovery_method=item_to_process[i]["discoverymethod"],
                        discovery_facility=item_to_process[i]["disc_facility"],
                        discovery_telescope=item_to_process[i]["disc_telescope"],
                        url = f"https://exoplanetarchive.ipac.caltech.edu/overview/{item_to_process[i]["pl_name"].replace(" ","%20")}"
                    )

                    new_records.append(new_record)

                db.session.add_all(new_records)
                db.session.commit()

            elif trans_type == "update_constellations":
                # Delete all existing records from the "constellations" database table:
                db.session.query(Constellations).delete()
                db.session.commit()

                # Upload, to the "constellations" database table, all contents of the "item_to_process"
                # parameter (in this case, the "constellations_data" dictionary from the calling function):
                new_records = []
                for key in item_to_process:
                    new_record = Constellations(
                        name=key,
                        abbreviation=item_to_process[key]["abbreviation"],
                        nickname=item_to_process[key]["nickname"],
                        url=item_to_process[key]["url"],
                        area=item_to_process[key]["area"],
                        myth_assoc=item_to_process[key]["myth_assoc"],
                        first_appear=item_to_process[key]["first_appear"],
                        brightest_star_name=item_to_process[key]["brightest_star_name"],
                        brightest_star_url=item_to_process[key]["brightest_star_url"]
                    )
                    new_records.append(new_record)

                db.session.add_all(new_records)
                db.session.commit()

            elif trans_type == "update_mars_photos_available":
                # Delete all existing records from the "mars_photos_available" database table:
                db.session.query(MarsPhotosAvailable).delete()
                db.session.commit()

                # Upload, to the "mars_photos_available" database table, all contents of the "item_to_process"
                # parameter (in this case, the "photos_available" dictionary from the calling function):
                new_records = []
                for key in item_to_process:
                    new_record = MarsPhotosAvailable(
                        rover_earth_date_combo=key,
                        rover_name=item_to_process[key]["rover_name"],
                        sol=int(item_to_process[key]["sol"]),
                        earth_date = item_to_process[key]["earth_date"],
                        cameras=item_to_process[key]["cameras"],
                        total_photos=item_to_process[key]["total_photos"]
                    )
                    new_records.append(new_record)

                db.session.add_all(new_records)
                db.session.commit()

            elif trans_type == "update_mars_photo_details":
                # Upload, to the "mars_photo_details" database table, all contents of the "item_to_process"
                # parameter (in this case, the "photo_details_rover_earth_date_combo" list from the calling function):
                new_records = []
                for i in range(0, len(item_to_process)):
                    new_record = MarsPhotoDetails(
                        rover_earth_date_combo=item_to_process[i]["rover_earth_date_combo"],
                        rover_name=item_to_process[i]["rover_name"],
                        sol=int(item_to_process[i]["sol"]),
                        pic_id=item_to_process[i]["pic_id"],
                        earth_date = item_to_process[i]["earth_date"],
                        camera_name=item_to_process[i]["camera_name"],
                        camera_full_name=item_to_process[i]["camera_full_name"],
                        url=item_to_process[i]["url"]
                    )

                    new_records.append(new_record)

                db.session.add_all(new_records)
                db.session.commit()

            elif trans_type == "update_mars_photo_details_delete_existing":
                # Capture optional arguments:
                rover_name = kwargs.get("rover_name", None)
                earth_date = kwargs.get("earth_date", None)

                # Delete, from the "mars_photo_details" database table, all records where the rover name and
                # earth date collectively match what was passed to this function:
                db.session.execute(db.delete(MarsPhotoDetails).where(MarsPhotoDetails.rover_earth_date_combo == rover_name + "_" + earth_date))
                db.session.commit()

            elif trans_type == "update_space_news":
                # Delete all records from the "space_news" database table:
                db.session.execute(db.delete(SpaceNews))
                db.session.commit()

                # Import the newly acquired articles (from the "item_to_process" list) into the "space_news" database table:
                new_records = []
                for i in range(0, len(item_to_process)):
                    new_record = SpaceNews(
                        article_id=item_to_process[i]["id"],
                        title=item_to_process[i]["title"],
                        url=item_to_process[i]["url"],
                        summary=item_to_process[i]["summary"],
                        news_site=item_to_process[i]["news_site"],
                        date_time_published=datetime.strptime(item_to_process[i]["published_at"], "%Y-%m-%dT%H:%M:%SZ"),
                        date_time_updated=datetime.strptime(item_to_process[i]["updated_at"],"%Y-%m-%dT%H:%M:%S.%fZ")
                    )
                    new_records.append(new_record)

                db.session.add_all(new_records)
                db.session.commit()

        # Return successful-execution indication to the calling function:
        return True

    except:  # An error has occurred.
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
        with open("log_eye_for_space_" + current_date_time_file + ".txt", "a") as f:
            f.write(datetime.now().strftime("%Y-%m-%d @ %I:%M %p") + ":\n")
            f.write(activity + ": " + log + "\n")

        # Close the log file:
        f.close()

    except:
        dlg = wx.MessageBox(f"Error: System log could not be updated.\n{traceback.format_exc()}", 'Error', wx.OK | wx.ICON_INFORMATION)
        dlg = None


def select_random_quotes():
    pass

# Run main function for this application:
run_app()

# get_inspirational_data()

# Destroy the object that was created to show user dialog and message boxes:
dlg.Destroy()

if __name__ == "__main__":
    app.run(debug=True, port=5003)