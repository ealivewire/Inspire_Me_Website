# PROFESSIONAL PROJECT: Inspire Me Website

# OBJECTIVE: To implement a website which automates the retrieval and distribution of daily inspiration quotes.

# Import necessary library(ies):
import requests
from data import app, db, mars_rovers, recognition, spreadsheet_attributes, API_KEY_ASTRONOMY_PIC_OF_THE_DAY, API_KEY_CLOSEST_APPROACH_ASTEROIDS, API_KEY_GET_LOC_FROM_LAT_AND_LON, API_KEY_MARS_ROVER_PHOTOS,  SENDER_EMAIL_GMAIL, SENDER_HOST, SENDER_PASSWORD_GMAIL, SENDER_PORT, URL_ASTRONOMY_PIC_OF_THE_DAY, URL_CLOSEST_APPROACH_ASTEROIDS, URL_CONFIRMED_PLANETS, URL_CONSTELLATION_ADD_DETAILS_1, URL_CONSTELLATION_ADD_DETAILS_2A, URL_CONSTELLATION_ADD_DETAILS_2B, URL_CONSTELLATION_MAP_SITE, URL_GET_LOC_FROM_LAT_AND_LON, URL_ISS_LOCATION, URL_MARS_ROVER_PHOTOS_BY_ROVER, URL_MARS_ROVER_PHOTOS_BY_ROVER_AND_OTHER_CRITERIA, URL_PEOPLE_IN_SPACE_NOW, URL_SPACE_NEWS, WEB_LOADING_TIME_ALLOWANCE
from data import ApproachingAsteroids, ConfirmedPlanets, Constellations, MarsPhotoDetails, MarsPhotosAvailable, MarsRoverCameras, MarsRovers, SpaceNews, Users
from data import AdminLoginForm, AdminUpdateForm, ContactForm, DisplayApproachingAsteroidsSheetForm, DisplayConfirmedPlanetsSheetForm, DisplayConstellationSheetForm, DisplayMarsPhotosSheetForm, ViewApproachingAsteroidsForm, ViewConfirmedPlanetsForm, ViewConstellationForm, ViewMarsPhotosForm
from datetime import datetime, timedelta
from dotenv import load_dotenv
from flask import Flask, abort, render_template, redirect, url_for, request
from flask_bootstrap import Bootstrap5
from flask_login import UserMixin, login_user, LoginManager, current_user, logout_user
from flask_sqlalchemy import SQLAlchemy
from functools import wraps  # Used in 'admin_only" decorator function
from flask_wtf import FlaskForm
from selenium import webdriver
from selenium.webdriver.common.by import By
from sqlalchemy import Integer, String, Boolean, Float, DateTime, func, distinct
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column
from werkzeug.security import check_password_hash
from wtforms import EmailField, SelectField, StringField, SubmitField, TextAreaField, BooleanField, PasswordField
from wtforms.validators import InputRequired, Length, Email
import collections  # Used for sorting items in the constellations dictionary
import email_validator
import glob
import math
import os
import smtplib
import time
import traceback
import unidecode
import wx
import wx.lib.agw.pybusyinfo as PBI
import xlsxwriter

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
    global db, app

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


# Configure route for "Approaching Asteroids" web page:
@app.route('/approaching_asteroids',methods=["GET", "POST"])
def approaching_asteroids():
    global db, app

    try:
        # Instantiate an instance of the "ViewApproachingAsteroidsForm" class:
        form = ViewApproachingAsteroidsForm()

        # Instantiate an instance of the "DisplayApproachingAsteroidsSheetForm" class:
        form_ss = DisplayApproachingAsteroidsSheetForm()

        # Populate the close approach date listbox with an ordered list of close approach dates represented in the database:
        list_close_approach_dates = []
        close_approach_dates = db.session.query(distinct(ApproachingAsteroids.close_approach_date)).order_by(ApproachingAsteroids.close_approach_date).all()
        for close_approach_date in close_approach_dates:
            list_close_approach_dates.append(str(close_approach_date)[2:12])
        form.list_close_approach_date.choices = list_close_approach_dates

        # Populate the approaching-asteroids sheet file listbox with the sole sheet viewable in this scope:
        form_ss.list_approaching_asteroids_sheet_name.choices = ["ApproachingAsteroids.xlsx"]

        # Validate form entries upon submittal. Depending on the form involved, perform additional processing:
        if form.validate_on_submit():
            if form.list_close_approach_date.data != None:
                error_msg = ""
                # Retrieve the record from the database which pertains to confirmed planets discovered in the selected year:
                approaching_asteroids_details = retrieve_from_database(trans_type="approaching_asteroids_by_close_approach_date", close_approach_date=form.list_close_approach_date.data)

                if approaching_asteroids_details == {}:
                    error_msg = "Error: Data could not be obtained at this time."
                elif approaching_asteroids_details == []:
                    error_msg = "No matching records were retrieved."

                # Show web page with retrieved approaching-asteroid details:
                return render_template('show_approaching_asteroids_details.html', approaching_asteroids_details=approaching_asteroids_details, close_approach_date=form.list_close_approach_date.data, error_msg=error_msg, recognition_scope_specific=recognition["approaching_asteroids"], recognition_web_template=recognition["web_template"])

            else:
                # Open the selected spreadsheet file:
                os.startfile(str(form_ss.list_approaching_asteroids_sheet_name.data))

        # Go to the web page to render the results:
        return render_template('approaching_asteroids.html', form=form, form_ss=form_ss, recognition_scope_specific=recognition["approaching_asteroids"], recognition_web_template=recognition["web_template"])

    except:  # An error has occurred.
        dlg = wx.MessageBox(f"Error (route: '/approaching_asteroids'): {traceback.format_exc()}", 'Error', wx.OK | wx.ICON_INFORMATION)
        update_system_log("route: '/approaching_asteroids'", traceback.format_exc())
        dlg = None


# Configure route for "Astronomy Pic of the Day" web page:
@app.route('/astronomy_pic_of_day')
def astronomy_pic_of_day():
    global db, app

    try:
        # Get details re: the astronomy picture of the day:
        json, copyright_details, error_msg = get_astronomy_pic_of_the_day()

        # Go to the web page to render the results:
        return render_template("astronomy_pic_of_day.html", json=json, copyright_details=copyright_details, error_msg=error_msg, recognition_scope_specific=recognition["astronomy_pic_of_day"], recognition_web_template=recognition["web_template"])

    except:  # An error has occurred.
        dlg = wx.MessageBox(f"Error (route: '/astronomy_pic_of_day'): {traceback.format_exc()}", 'Error', wx.OK | wx.ICON_INFORMATION)
        update_system_log("route: '/astronomy_pic_of_day'", traceback.format_exc())
        dlg = None


# Configure route for "Confirmed Planets" web page:
@app.route('/confirmed_planets',methods=["GET", "POST"])
def confirmed_planets():
    global db, app

    try:
        # Instantiate an instance of the "ViewConstellationForm" class:
        form = ViewConfirmedPlanetsForm()

        # Instantiate an instance of the "DisplayConfirmedPlanetsSheetForm" class:
        form_ss = DisplayConfirmedPlanetsSheetForm()

        # Populate the discovery year listbox with an ordered (descending) list of discovery years represented in the database:
        list_discovery_years = []
        discovery_years = db.session.query(distinct(ConfirmedPlanets.discovery_year)).order_by(ConfirmedPlanets.discovery_year.desc()).all()
        for year in discovery_years:
            list_discovery_years.append(int(str(year)[1:5]))
        form.list_discovery_year.choices = list_discovery_years

        # Populate the confirmed planets sheet file listbox with the sole sheet viewable in this scope:
        form_ss.list_confirmed_planets_sheet_name.choices = ["ConfirmedPlanets.xlsx"]

        # Validate form entries upon submittal. Depending on the form involved, perform additional processing:
        if form.validate_on_submit():
            if form.list_discovery_year.data != None:
                error_msg = ""
                # Retrieve the record from the database which pertains to confirmed planets discovered in the selected year:
                confirmed_planets_details = retrieve_from_database(trans_type="confirmed_planets_by_disc_year", disc_year=form.list_discovery_year.data)

                if confirmed_planets_details == {}:
                    error_msg = "Error: Data could not be obtained at this time."
                elif confirmed_planets_details == []:
                    error_msg = "No matching records were retrieved."

                # Show web page with retrieved confirmed-planet details:
                return render_template('show_confirmed_planets_details.html', confirmed_planets_details=confirmed_planets_details, disc_year=form.list_discovery_year.data, error_msg=error_msg, recognition_scope_specific=recognition["confirmed_planets"], recognition_web_template=recognition["web_template"])

            else:
                # Open the selected spreadsheet file:
                os.startfile(str(form_ss.list_confirmed_planets_sheet_name.data))

        # Go to the web page to render the results:
        return render_template('confirmed_planets.html', form=form, form_ss=form_ss, recognition_scope_specific=recognition["confirmed_planets"], recognition_web_template=recognition["web_template"])

    except:  # An error has occurred.
        dlg = wx.MessageBox(f"Error (route: '/confirmed_planets'): {traceback.format_exc()}", 'Error', wx.OK | wx.ICON_INFORMATION)
        update_system_log("route: '/confirmed_planets'", traceback.format_exc())
        dlg = None


# Configure route for "Constellations" web page:
@app.route('/constellations',methods=["GET", "POST"])
def constellations():
    global db, app

    try:
        # Instantiate an instance of the "ViewConstellationForm" class:
        form = ViewConstellationForm()

        # Instantiate an instance of the "DisplayConstellationSheetForm" class:
        form_ss = DisplayConstellationSheetForm()

        # Populate the constellation name listbox with an ordered list of constellation names from the database:
        form.list_constellation_name.choices = db.session.execute(db.select(Constellations.name + " (" + Constellations.nickname + ")").order_by(Constellations.name)).scalars().all()

        # Populate the constellation sheet file listbox with the sole sheet viewable in this scope:
        form_ss.list_constellation_sheet_name.choices = ["Constellations.xlsx"]

        # Validate form entries upon submittal. Depending on the form involved, perform additional processing:
        if form.validate_on_submit():

            if form.list_constellation_name.data != None:
                # Capture selected constellation name:
                selected_constellation_name = form.list_constellation_name.data.split("(")[0][:len(form.list_constellation_name.data.split("(")[0])-1]

                # Retrieve the record from the database which pertains to the selected constellation name:
                constellation_details = db.session.execute(db.select(Constellations).where(Constellations.name == selected_constellation_name)).scalar()

                # Show web page with retrieved constellation details:
                return render_template('show_constellation_details.html', constellation_details=constellation_details, recognition_scope_specific=recognition["constellations"], recognition_web_template=recognition["web_template"])

            else:
                # Open the selected spreadsheet file:
                os.startfile(str(form_ss.list_constellation_sheet_name.data))

        # Go to the web page to render the results:
        return render_template('constellations.html', form=form, form_ss=form_ss, recognition_scope_specific=recognition["constellations"], recognition_web_template=recognition["web_template"])

    except:  # An error has occurred.
        dlg = wx.MessageBox(f"Error (route: '/constellations'): {traceback.format_exc()}", 'Error', wx.OK | wx.ICON_INFORMATION)
        update_system_log("route: '/constellations'", traceback.format_exc())
        dlg = None


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


# Configure route for "Photos from Mars" web page:
@app.route('/mars_photos',methods=["GET", "POST"])
def mars_photos():
    global db, app

    try:
        # Instantiate an instance of the "ViewConstellationForm" class:
        form = ViewMarsPhotosForm()

        # Instantiate an instance of the "DisplayMarsPhotosSheetForm" class:
        form_ss = DisplayMarsPhotosSheetForm()

        # Populate the rover name / earth date combo listbox with an ordered list of such combinations:
        list_rover_earth_date_combos = []
        rover_earth_date_combos = db.session.query(distinct(MarsPhotosAvailable.rover_earth_date_combo)).order_by(MarsPhotosAvailable.rover_name, MarsPhotosAvailable.earth_date.desc()).all()
        for rover_earth_date_combo in rover_earth_date_combos:
            list_rover_earth_date_combos.append(str(rover_earth_date_combo).split("'")[1])
        form.list_rover_earth_date_combo.choices = list_rover_earth_date_combos

        # Populate the Mars photos sheet file listbox with all filenames of spreadsheets pertinent to this scope:
        form_ss.list_mars_photos_sheet_name.choices = glob.glob("Mars Photos*.xlsx")

        # Validate form entries upon submittal. Depending on the form involved, perform additional processing:
        if form.validate_on_submit():
            if form.list_rover_earth_date_combo.data != None:
                error_msg = ""
                # Retrieve the record from the database which pertains to Mars photos taken via the selected rover / earth date combo:
                mars_photos_details = retrieve_from_database(trans_type="mars_photos_by_rover_earth_date_combo", rover_earth_date_combo=form.list_rover_earth_date_combo.data)

                if mars_photos_details == {}:
                    error_msg = "Error: Data could not be obtained at this time."
                elif mars_photos_details == []:
                    error_msg = "No matching records were retrieved."

                # Show web page with retrieved photo details:
                return render_template('show_mars_photos_details.html', mars_photos_details=mars_photos_details, rover_earth_date_combo=form.list_rover_earth_date_combo.data, error_msg=error_msg, recognition_scope_specific=recognition["mars_photos"], recognition_web_template=recognition["web_template"])

            else:
                # Open the selected spreadsheet file:
                os.startfile(str(form_ss.list_mars_photos_sheet_name.data))

        # Go to the web page to render the results:
        return render_template('mars_photos.html', form=form, form_ss=form_ss, recognition_scope_specific=recognition["mars_photos"], recognition_web_template=recognition["web_template"])

    except:  # An error has occurred.
        dlg = wx.MessageBox(f"Error (route: '/mars_photos'): {traceback.format_exc()}", 'Error', wx.OK | wx.ICON_INFORMATION)
        update_system_log("route: '/mars_photos'", traceback.format_exc())
        dlg = None


# Configure route for "Space News" web page:
@app.route('/space_news')
def space_news():
    global db, app

    try:
        # Get results of obtaining and processing the desired information:
        success, error_msg = get_space_news()

        if success:
            # Query the table for space news articles:
            with app.app_context():
                articles = db.session.execute(db.select(SpaceNews).order_by(SpaceNews.row_id)).scalars().all()
                if articles.count == 0:
                    success = False
                    error_msg = "Error: Cannot retrieve article data from database."

        else:
            articles = None

        # Go to the web page to render the results:
        return render_template("space_news.html", articles=articles, success=success, error_msg=error_msg, recognition_scope_specific=recognition["space_news"], recognition_web_template=recognition["web_template"])

    except:  # An error has occurred.
        dlg = wx.MessageBox(f"Error (route: '/space_news'): {traceback.format_exc()}", 'Error', wx.OK | wx.ICON_INFORMATION)
        update_system_log("route: '/space_news'", traceback.format_exc())
        dlg = None


# Configure route for "Where is ISS" web page:
@app.route('/where_is_iss')
def where_is_iss():
    global db, app

    try:
        # Get ISS's current location along with a URL to get a map plotting said location:
        location_address, location_url = get_iss_location()

        # Go to the web page to render the results:
        return render_template("where_is_iss.html", location_address=location_address, location_url=location_url, has_url=not(location_url == ""), recognition_scope_specific=recognition["where_is_iss"], recognition_web_template=recognition["web_template"])

    except:  # An error has occurred.
        dlg = wx.MessageBox(f"Error (route: '/where_is_iss'): {traceback.format_exc()}", 'Error', wx.OK | wx.ICON_INFORMATION)
        update_system_log("route: '/where_is_iss'", traceback.format_exc())
        dlg = None


# Configure route for "Who is in Space Now" web page:
@app.route('/who_is_in_space_now')
def who_is_in_space_now():
    global db, app

    try:
        # Get results of obtaining a JSON with the desired information:
        json, has_json = get_people_in_space_now()

        # Go to the web page to render the results:
        return render_template("who_is_in_space_now.html", json=json, has_json=has_json, recognition_scope_specific=recognition["who_is_in_space_now"], recognition_web_template=recognition["web_template"])

    except:  # An error has occurred.
        dlg = wx.MessageBox(f"Error (route: '/who_is_in_space_now'): {traceback.format_exc()}", 'Error', wx.OK | wx.ICON_INFORMATION)
        update_system_log("route: '/who_is_in_space_now'", traceback.format_exc())
        dlg = None


# DEFINE FUNCTIONS TO BE USED FOR THIS APPLICATION (LISTED IN ALPHABETICAL ORDER BY FUNCTION NAME):
# *************************************************************************************************
def close_workbook(workbook):
    """Function to close a spreadsheet workbook, checking if the file is open"""
    try:
        while True:
            try:
                # Close the workbook.
                workbook.close()

                # Return successful-execution indication to the calling function:
                return True

            except xlsxwriter.exceptions.FileCreateError as e:
                # Inform user that exception has occurred and prompt for confirmation
                # to re-attempt file creation/closure:
                user_answer = wx.MessageBox(f"Spreadsheet file '{workbook.filename}' could not be created.\nPlease close the file if it is open in Excel.\nWould you like to try to write the file again?", 'Administrative Update', wx.YES_NO | wx.ICON_QUESTION)

                if user_answer == 2:  # User wishes to re-attempt file creation/closure:
                    user_answer = None
                    continue
                else:  # User has elected to not re-attempt file creation/closure.
                    user_answer = None
                    # Return failed-execution indication to the calling function.
                    return False

            # Break from the "while" loop:
            break

    except:  # An error has occurred.
        update_system_log("close_workbook", traceback.format_exc())


def config_database():
    """Function for configuring the database tables supporting this website"""
    global db, app, ApproachingAsteroids, ConfirmedPlanets, Constellations, MarsPhotoDetails, MarsPhotosAvailable, MarsRoverCameras, MarsRovers, SpaceNews, Users

    try:
        # Create the database object using the SQLAlchemy constructor:
        db = SQLAlchemy(model_class=Base)

        # Initialize the app with the extension:
        db.init_app(app)

        # Configure database tables (listed in alphabetical order; class names are sufficiently descriptive):
        class ApproachingAsteroids(db.Model):
            id: Mapped[int] = mapped_column(Integer, primary_key=True)
            name: Mapped[str] = mapped_column(String(50), nullable=False)
            absolute_magnitude_h: Mapped[float] = mapped_column(Float, nullable=False)
            estimated_diameter_km_min: Mapped[float] = mapped_column(Float, nullable=False)
            estimated_diameter_km_max: Mapped[float] = mapped_column(Float, nullable=False)
            is_potentially_hazardous: Mapped[bool] = mapped_column(Boolean, nullable=False)
            close_approach_date: Mapped[str] = mapped_column(String(10), nullable=False)
            relative_velocity_km_per_s: Mapped[float] = mapped_column(Float, nullable=False)
            miss_distance_km: Mapped[float] = mapped_column(Float, nullable=False)
            orbiting_body: Mapped[str] = mapped_column(String(20), nullable=False)
            is_sentry_object: Mapped[bool] = mapped_column(Boolean, nullable=False)
            url: Mapped[str] = mapped_column(String(500), nullable=False)

        class ConfirmedPlanets(db.Model):
            row_id: Mapped[int] = mapped_column(Integer, primary_key=True)
            host_name: Mapped[str] = mapped_column(String(50), nullable=False)
            host_num_stars: Mapped[int] = mapped_column(Integer, nullable=False)
            host_num_planets: Mapped[int] = mapped_column(Integer, nullable=False)
            planet_name: Mapped[str] = mapped_column(String(50), unique=True, nullable=False)
            discovery_year: Mapped[int] = mapped_column(Integer, nullable=False)
            discovery_method: Mapped[str] = mapped_column(String(50), nullable=False)
            discovery_facility: Mapped[str] = mapped_column(String(100), nullable=False)
            discovery_telescope: Mapped[str] = mapped_column(String(50), nullable=False)
            url: Mapped[str] = mapped_column(String(500), nullable=False)

        class Constellations(db.Model):
            row_id: Mapped[int] = mapped_column(Integer, primary_key=True)
            name: Mapped[int] = mapped_column(String(20), unique=True, nullable=False)
            abbreviation: Mapped[str] = mapped_column(String(10), unique=False, nullable=False)
            nickname: Mapped[str] = mapped_column(String(30), unique=False, nullable=False)
            url: Mapped[str] = mapped_column(String(500), nullable=False)
            area: Mapped[str] = mapped_column(String(50), unique=False, nullable=False)
            myth_assoc: Mapped[str] = mapped_column(String(500), unique=False, nullable=False)
            first_appear: Mapped[str] = mapped_column(String(50), unique=False, nullable=False)
            brightest_star_name: Mapped[str] = mapped_column(String(40), unique=False, nullable=False)
            brightest_star_url: Mapped[str] = mapped_column(String(40), unique=False, nullable=False)

        class MarsPhotoDetails(db.Model):
            row_id: Mapped[int] = mapped_column(Integer, primary_key=True)
            rover_earth_date_combo = mapped_column(String(32), nullable=False)
            rover_name: Mapped[str] = mapped_column(String(15), nullable=False)
            sol: Mapped[str] = mapped_column(String(15), unique=False, nullable=False)
            pic_id: Mapped[int] = mapped_column(Integer, nullable=False)
            earth_date: Mapped[str] = mapped_column(String(15), nullable=False)
            camera_name: Mapped[str] = mapped_column(String(20), nullable=False)
            camera_full_name: Mapped[str] = mapped_column(String(50), nullable=False)
            url: Mapped[str] = mapped_column(String(500), nullable=False)

        class MarsPhotosAvailable(db.Model):
            row_id: Mapped[int] = mapped_column(Integer, primary_key=True)
            rover_earth_date_combo = mapped_column(String(32), nullable=False)
            rover_name: Mapped[str] = mapped_column(String(15), nullable=False)
            sol: Mapped[str] = mapped_column(String(15), unique=False, nullable=False)
            earth_date: Mapped[str] = mapped_column(String(15), nullable=False)
            cameras: Mapped[str] = mapped_column(String(250), nullable=False)
            total_photos: Mapped[int] = mapped_column(Integer, nullable=False)

        class MarsRoverCameras(db.Model):
            row_id: Mapped[int] = mapped_column(Integer, primary_key=True)
            rover_name: Mapped[str] = mapped_column(String(15), nullable=False)
            camera_name: Mapped[str] = mapped_column(String(20), nullable=False)
            camera_full_name: Mapped[str] = mapped_column(String(50), nullable=False)

        class MarsRovers(db.Model):
            row_id: Mapped[int] = mapped_column(Integer, primary_key=True)
            rover_name: Mapped[str] = mapped_column(String(15), nullable=False)
            active: Mapped[bool] = mapped_column(Boolean, nullable=False)

        class SpaceNews(db.Model):
            row_id: Mapped[int] = mapped_column(Integer, primary_key=True)
            article_id: Mapped[int] = mapped_column(Integer, nullable=False)
            news_site: Mapped[str] = mapped_column(String(30), nullable=False)
            title: Mapped[str] = mapped_column(String(250), nullable=False)
            summary: Mapped[str] = mapped_column(String(500), nullable=False)
            date_time_published: Mapped[datetime] = mapped_column(DateTime, nullable=True)
            date_time_updated: Mapped[datetime] = mapped_column(DateTime, nullable=True)
            url: Mapped[str] = mapped_column(String(500), nullable=False)

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


def create_workbook(workbook_name):
    """Function for creating and returning a spreadsheet workbook for subsequent population/formatting"""
    try:
        # Create and return the workbook:
        return xlsxwriter.Workbook(workbook_name)

    except:  # An error has occurred.
        update_system_log("create_workbook", traceback.format_exc())

        # Return failed-execution indication to the calling function:
        return None


def create_worksheet(workbook, worksheet_name):
    """Function for creating and returning a spreadsheet worksheet for subsequent population/formatting"""
    try:
        # Create and return the worksheet:
        return workbook.add_worksheet(worksheet_name)

    except:  # An error has occurred.
        update_system_log("create_worksheet", traceback.format_exc())

        # Return failed-execution indication to the calling function:
        return None


def delete_mars_photos_workbooks():
    """Function for deleting all spreadsheet workbooks in the application directory which pertain to Mars photos"""
    try:
        # Delete the summary workbook:
        while True:
            try:
                # Delete the summary workbook:
                os.remove("Mars Photos - Summary.xlsx")

            except FileNotFoundError:
                pass

            except PermissionError:
                # Inform user that exception has occurred and prompt for confirmation
                # to re-attempt file deletion:
                user_answer = wx.MessageBox(f"File 'Mars Photos - Summary.xlsx' could not be deleted prior to the upcoming update.\nPlease close the file if open in Excel.\nWould you like to try to delete the file again?", 'Administrative Update', wx.YES_NO | wx.ICON_QUESTION)

                if user_answer == 2:  # User wishes to re-attempt file deletion:
                    user_answer = None
                    continue
                else:  # User has elected to not re-attempt file deletion.
                    user_answer = None
                    # Return failed-execution indication to the calling function.
                    return False

            # Break from the "while" loop:
            break

        # Delete the details workbooks:
        while True:
            try:
                # Delete the details workbooks:
                for f in glob.glob("Mars Photos - Details - *.xlsx"):
                    os.remove(f)

            except FileNotFoundError:
                pass

            except PermissionError:
                # Inform user that exception has occurred and prompt for confirmation
                # to re-attempt file deletion:
                user_answer = wx.MessageBox(
                    f"One or more Mars photo details workbooks could not be deleted prior to the upcoming update.\nPlease close the file(s) if open in Excel.\nWould you like to try to delete the file(s) again?",
                    'Administrative Update', wx.YES_NO | wx.ICON_QUESTION)

                if user_answer == 2:  # User wishes to re-attempt file deletion:
                    user_answer = None
                    continue
                else:  # User has elected to not re-attempt file deletion.
                    user_answer = None
                    # Return failed-execution indication to the calling function.
                    return False

            # Break from the "while" loop:
            break

        # At this point, function is presumed to have executed successfully. Return successful-execution
        # indication to the calling function:
        return True

    except:  # An error has occurred.
        update_system_log("delete_mars_photos_workbooks", traceback.format_exc())

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


def export_data_to_spreadsheet_standard(data_scope, data_to_export):
    """Function to export data to a spreadsheet, with all appropriate formatting applied"""
    try:
        # Capture current date/time:
        current_date_time = datetime.now()
        current_date_time_spreadsheet = current_date_time.strftime("%d-%b-%Y @ %I:%M %p")

        # Create the workbook.  If an error occurred, update system log and return
        # failed-execution indication to the calling function:
        workbook = create_workbook(f"{spreadsheet_attributes[data_scope]["wrkbk_name"]}")
        if workbook == None:
            update_system_log("export_data_to_spreadsheet_standard (" + data_scope + ")", "Error: Workbook could not be created.")
            return False

        # Create the worksheet to contain data from the "data_to_export" variable.  If an error occurred,
        # update system log and return failed-execution indication to the calling function:
        worksheet = create_worksheet(workbook, spreadsheet_attributes[data_scope]["wksht_name"])
        if worksheet == None:
            update_system_log("export_data_to_spreadsheet_standard (" + data_scope + ")", "Error: Worksheet could not be created.")
            return False

        # Add and format the column headers. If an error occurred, update system log and return failed-execution indication to the calling function:
        if not prepare_spreadsheet_main_contents(workbook, worksheet, spreadsheet_attributes[data_scope]["headers"]):
            update_system_log("export_data_to_spreadsheet_standard (" + data_scope + ")", "Error: Spreadsheet headers could not be completely implemented.")
            return False

        # Write/format each item's data into the worksheet, using the "data_to_export" variable as the data source.
        # If an error occurred, update system log and return failed-execution indication to the calling function:
        if data_scope == "constellations":
            i = 3
            for key in data_to_export:
                if not prepare_spreadsheet_main_contents(workbook, worksheet, "constellation_data",dict_name=data_to_export, key=key, i=i):
                    update_system_log("export_data_to_spreadsheet_standard (" + data_scope + ")","Error: Spreadsheet main contents could not be completely implemented.")
                    return False
                i += 1
        else:  # Other items except constellations.
            if not prepare_spreadsheet_main_contents(workbook, worksheet, spreadsheet_attributes[data_scope]["data_to_export_name"], list_name=data_to_export):
                update_system_log("export_data_to_spreadsheet_standard (" + data_scope + ")","Error: Spreadsheet main contents could not be completely implemented.")
                return False

        # Add and format the spreadsheet header row, and implement the following: column widths, footer, page orientation, and margins.
        # If function failed, update system log and return failed-execution indication to the calling function:
        if not prepare_spreadsheet_supplemental_formatting(workbook, worksheet, spreadsheet_attributes[data_scope]["supp_fmtg"], current_date_time_spreadsheet, data_to_export, spreadsheet_attributes[data_scope]["num_cols_minus_one"], spreadsheet_attributes[data_scope]["col_widths"] ):
            update_system_log("export_data_to_spreadsheet_standard (" + data_scope + ")", "Error: Spreadsheet formatting could not be completed.")
            return False

        # Complete file creation/closure of the workbook, checking if the file is open.  If an error occurred or if file is open and user elected to not
        # re-attempt file creation/closure, update system log and return failed-execution indication to the calling function:
        if not close_workbook(workbook):
            update_system_log("export_data_to_spreadsheet_standard (" + data_scope + ")", "Error: Spreadsheet file creation failed.")
            return False

        # Return successful-execution indication to the calling function:
        return True

    except:  # An error has occurred.
        update_system_log("export_data_to_spreadsheet_standard (" + data_scope + ")", traceback.format_exc())

        # Return failed-execution indication to the calling function:
        return False


def export_mars_photos_to_spreadsheet(photos_available, photo_details):
    """Function to export data on available Mars rover photos to a spreadsheet, with all appropriate formatting applied"""
    try:
        # Inform user that export-to-spreadsheet execution will begin:
        dlg = PBI.PyBusyInfo("Photos from Mars: Exporting results to spreadsheet file...", title="Administrative Update")

        # Capture current date/time:
        current_date_time = datetime.now()
        current_date_time_spreadsheet = current_date_time.strftime("%d-%b-%Y @ %I:%M %p")

        # Create the workbook.  If an error occurred, update system log and return
        # failed-execution indication to the calling function:
        photos_available_workbook = create_workbook(f"Mars Photos - Summary.xlsx")
        if photos_available_workbook == None:
            update_system_log("export_mars_photos_to_spreadsheet", "Error: Workbook (photos available summary) not be created.")
            return False

        # Create the worksheet to contain photos-available data from the "photos_available" list of database records.
        # If an error occurred, update system log and return failed-execution indication to the calling function::
        photos_available_worksheet = create_worksheet(photos_available_workbook, f"Summary")
        if photos_available_worksheet == None:
            update_system_log("export_mars_photos_to_spreadsheet", "Error: Worksheet (photos available summary) could not be created.")
            return False

        # Add and format the column headers. If an error occurred, update system log and return failed-execution indication to the calling function:
        if not prepare_spreadsheet_main_contents(photos_available_workbook, photos_available_worksheet, "photos_available_headers"):
            update_system_log("export_mars_photos_to_spreadsheet", "Error: Spreadsheet headers (for photos available summary) could not be completely implemented.")
            return False

        # Populate the "Summary" worksheet with the contents of the "photos_available" list of database records:
        # If function failed, update system log and return failed-execution indication to the calling function:
        if not prepare_spreadsheet_main_contents(photos_available_workbook, photos_available_worksheet,"photos_available_data", list_name=photos_available):
            update_system_log("export_mars_photos_to_spreadsheet","Error: Spreadsheet main contents (for photos available summary) could not be completely implemented.")
            return False

        # Add and format the spreadsheet header row, and implement the following: column widths, footer, page orientation, and margins.
        # If function failed, update system log and return failed-execution indication to the calling function:
        if not prepare_spreadsheet_supplemental_formatting(photos_available_workbook, photos_available_worksheet, "photos_available", current_date_time_spreadsheet, photos_available, 4, (15, 15, 7, 80, 15)):
            update_system_log("export_mars_photos_to_spreadsheet (" + data_scope + ")", "Error: Spreadsheet formatting (for photos available summary) could not be completed.")
            return False

        # Complete file creation/closure of the photos available summary workbook, checking if the file is open.
        # If an error occurred or if file is open and user elected to not re-attempt file creation/closure, update
        # system log and return failed-execution indication to the calling function:
        dlg = PBI.PyBusyInfo("Photos from Mars: Spreadsheet file 'Mars Photos - Summary.xlsx': Saving in progress...", title="Administrative Update")
        if not close_workbook(photos_available_workbook):
            update_system_log("export_mars_photos_to_spreadsheet", "Error: Spreadsheet file creation (photos available summary) failed.")
            return False
        dlg = PBI.PyBusyInfo("Photos from Mars: Spreadsheet file 'Mars Photos - Summary.xlsx': Saving completed...", title="Administrative Update")

        # For each rover, create and format a worksheet to contain details for available photos taken by that rover
        # each earth year.  If function failed, update system log and return failed-execution indication to the calling function:
        rovers_represented = get_mars_photos_summarize_photo_counts_by_rover_and_earth_year()
        if rovers_represented == []:
            update_system_log("export_mars_photos_to_spreadsheet", "Error: 'rovers_represented' data could not be retrieved.")
            return False

        # Initialize variables needed to process photo details using the contents "rovers_represented" variable:
        worksheets_needed = []
        row_start = 0
        row_end = 0

        # Capture all of the photo-details workbooks that need to be created:
        for i in range(0, len(rovers_represented)):
            rover_name = rovers_represented[i][0]
            earth_year = rovers_represented[i][1]
            rover_earth_year_combo = rovers_represented[i][2]

            # Determine whether a particular rover/earth year combo needs to be split up into
            # multiple workbooks (based on whether its contents exceeds 65530 photos):
            if rovers_represented[i][3] <= 65530:
                row_end += rovers_represented[i][3]
                worksheets_needed.append((rover_earth_year_combo, earth_year, rover_name, 1, row_start, row_end))
                row_start = row_end
            else:
                worksheet_to_add = ""
                rover_number_of_sheets_needed = math.ceil(rovers_represented[i][3] / 65530)
                for j in range(0, rover_number_of_sheets_needed):
                    worksheet_to_add = rover_earth_year_combo + "_Part" + str(j + 1)
                    if (j + 1) == rover_number_of_sheets_needed:
                        row_end += rovers_represented[i][3] - 65530
                    else:
                        row_end += 65530
                    worksheets_needed.append((worksheet_to_add, earth_year, rover_name, rover_number_of_sheets_needed, row_start, row_end))
                    row_start = row_end

        # Create and populate photo details workbook for the current rover/earth year combo being processed:
        for i in range(0, len(worksheets_needed)):
            # Create the workbook.  If an error occurred, update system log and return
            # failed-execution indication to the calling function:
            photo_details_workbook = create_workbook(f"Mars Photos - Details - {worksheets_needed[i][0]}.xlsx")
            if photo_details_workbook == None:
                update_system_log("export_mars_photos_to_spreadsheet",
                                  "Error: Workbook (photo details for {worksheets_needed[i][0]}) not be created.")
                return False

            # Create the worksheet to contain photo details data from the "photo_details" list of database records.
            # If an error occurred, update system log and return failed-execution indication to the calling function::
            photo_details_worksheet = create_worksheet(photo_details_workbook, "Details")
            if photo_details_worksheet == None:
                update_system_log("export_mars_photos_to_spreadsheet",
                                  f"Error: Worksheet (photo details for {worksheets_needed[i][0]}) could not be created.")
                return False

            # Add and format the column headers. If an error occurred, update system log and return failed-execution indication to the calling function:
            if not prepare_spreadsheet_main_contents(photo_details_workbook, photo_details_worksheet,"photo_details_headers"):
                update_system_log("export_mars_photos_to_spreadsheet",
                                  f"Error: Spreadsheet headers (photo details for {worksheets_needed[i][0]}) could not be completely implemented.")
                return False

            # Populate the worksheet with its corresponding contents of the "photo_details" list of database records.
            # If function failed, update system log and return failed-execution indication to the calling function:
            if not prepare_spreadsheet_main_contents(photo_details_workbook, photo_details_worksheet, "photo_details_data", list_name=photo_details, worksheet_details=worksheets_needed[i]):
                update_system_log("export_mars_photos_to_spreadsheet",
                                  "Error: Spreadsheet main contents (photo details for {worksheets_needed[i][0]}) could not be completely implemented.")
                return False

            # Add and format the spreadsheet header row, and implement the following: column widths, footer, page orientation, and margins.
            # If function failed, update system log and return failed-execution indication to the calling function:
            if not prepare_spreadsheet_supplemental_formatting(photo_details_workbook, photo_details_worksheet, "photo_details", current_date_time_spreadsheet, photos_available, 4, (15, 15, 7, 15, 30, 50, 80), rover_name=worksheets_needed[i][2], earth_year=worksheets_needed[i][1], rover_earth_year_combo=worksheets_needed[i][0], rover_number_of_sheets_needed=worksheets_needed[i][3]):
                update_system_log("export_mars_photos_to_spreadsheet (" + data_scope + ")",
                                  "Error: Spreadsheet formatting (photo details for {worksheets_needed[i][0]}) could not be completed.")
                return False

            # Complete file creation/closure of the workbook, checking if the file is open.  If an error occurred or if file is open and user elected to not
            # re-attempt file creation/closure, update system log and return failed-execution indication to the calling function:
            dlg = PBI.PyBusyInfo(
                f"Photos from Mars: Spreadsheet file 'Mars Photos - Details - {worksheets_needed[i][0]}.xlsx': Saving in progress...",title="Administrative Update")
            if not close_workbook(photo_details_workbook):
                return False
            dlg = PBI.PyBusyInfo(
                f"Photos from Mars: Spreadsheet file 'Mars Photos - Details - {worksheets_needed[i][0]}.xlsx': Saving completed...",title="Administrative Update")

        # Return successful-execution indication to the calling function:
        dlg = None
        return True

    except:  # An error has occurred.
        update_system_log("export_mars_photos_to_spreadsheet", traceback.format_exc())

        # Return failed-execution indication to the calling function:
        return False


def find_element(driver, find_type, find_details):
    """Function to find an element via a web-scraping procedure"""
    # NOTE: Error handling is deferred to the calling function:
    if find_type == "xpath":
        return driver.find_element(By.XPATH, find_details)


def get_approaching_asteroids():
    """Function that retrieves and processes a list of asteroids based on closest approach to Earth"""
    # Capture the current date:
    current_date = datetime.now()

    # Capture the current date + an added window (delta) of the following 7 days:
    current_date_with_delta = current_date + timedelta(days=7)

    try:
        # Execute the API request (limit: closest approach <= 7 days from today):
        response = requests.get(URL_CLOSEST_APPROACH_ASTEROIDS + "?start_date=" + current_date.strftime("%Y-%m-%d") + "&end_date=" + current_date_with_delta.strftime("%Y-%m-%d") + "&api_key=" + API_KEY_CLOSEST_APPROACH_ASTEROIDS)

        # Initialize variable to store collected necessary asteroid data:
        approaching_asteroids = []

        # If the API request was successful, display the results:
        if response.status_code == 200:  # API request was successful.

            # Capture desired fields from the returned JSON:
            for key in response.json()["near_earth_objects"]:
                for asteroid in response.json()["near_earth_objects"][key]:
                    asteroid_dict = {
                        "id": asteroid["id"],
                        "name": asteroid["name"],
                        "absolute_magnitude_h": asteroid["absolute_magnitude_h"],
                        "estimated_diameter_km_min": asteroid["estimated_diameter"]["kilometers"]["estimated_diameter_min"],
                        "estimated_diameter_km_max": asteroid["estimated_diameter"]["kilometers"]["estimated_diameter_max"],
                        "is_potentially_hazardous": asteroid["is_potentially_hazardous_asteroid"],
                        "close_approach_date": asteroid["close_approach_data"][0]["close_approach_date"],
                        "relative_velocity_km_per_s": asteroid["close_approach_data"][0]["relative_velocity"]["kilometers_per_second"],
                        "miss_distance_km": asteroid["close_approach_data"][0]["miss_distance"]["kilometers"],
                        "orbiting_body": asteroid["close_approach_data"][0]["orbiting_body"],
                        "is_sentry_object": asteroid["is_sentry_object"],
                        "url": asteroid["nasa_jpl_url"]
                        }

                    # Add captured data for each asteroid (as a dictionary) to the "approaching_asteroids" list:
                    approaching_asteroids.append(asteroid_dict)

            # Delete the existing records in the "approaching_asteroids" database table and update same with
            # the up-to-date data (from the JSON).  If an error occurred, update system log and return a
            # failed-execution indication to the calling function:
            if not update_database("update_approaching_asteroids", approaching_asteroids):
                update_system_log("get_approaching_asteroids", "Error: Database could not be updated. Data cannot be obtained at this time.")
                return "Error: Database could not be updated. Data cannot be obtained at this time.", False

            # Retrieve all existing records in the "approaching_asteroids" database table. If the function
            # called returns an empty directory, update system log and return a failed-execution indication
            # to the calling function:
            asteroids_data = retrieve_from_database("approaching_asteroids")
            if asteroids_data == {}:
                update_system_log("get_approaching_asteroids", "Error: Data cannot be obtained at this time.")
                return "Error: Data cannot be obtained at this time.", False

            # If an empty list was returned, no records satisfied the query.  Therefore, update system log and
            # return a failed-execution indication to the calling function:
            elif asteroids_data == []:
                update_system_log("get_approaching_asteroids", "No matching records were retrieved.")
                return "No matching records were retrieved.", False

            # Create and format a spreadsheet file (workbook) to contain all asteroids data. If execution failed,
            # update system log and return failed-execution indication to the calling function:
            if not export_data_to_spreadsheet_standard("approaching_asteroids", asteroids_data):
                update_system_log("get_approaching_asteroids", "Error: Spreadsheet creation could not be completed at this time.")
                return "Error: Spreadsheet creation could not be completed at this time.", False

            # At this point, function is deemed to have executed successfully.  Update system log and
            # return successful-execution indication to the calling function:
            update_system_log("get_approaching_asteroids", "Successfully updated.")
            return "", True

        else:  # API request failed. Update system log and return failed-execution indication to the calling function:
            update_system_log("get_approaching_asteroids", "Error: API request failed. Data cannot be obtained at this time.")
            return "Error: API request failed. Data cannot be obtained at this time.", False

    except:  # An error has occurred.
        update_system_log("get_approaching_asteroids", traceback.format_exc())

        # Return failed-execution indication to the calling function:
        return "An error has occurred. Data cannot be obtained at this time.", False


def get_astronomy_pic_of_the_day():
    """Function to retrieve the astronomy picture of the day"""
    # Initialize variables to be used for returning values to the calling function:
    json = {}
    copyright_details = ""
    error_message = ""

    try:
        # Execute API request:
        url = URL_ASTRONOMY_PIC_OF_THE_DAY + "?api_key=" + API_KEY_ASTRONOMY_PIC_OF_THE_DAY
        response = requests.get(url)

        # If the API request was successful, capture the results:
        if response.status_code == 200:  # API request was successful.
            json = response.json()

            # If there is copyright info. included in the JSON, capture it:
            try:
                copyright_details = f"Copyright: {response.json()["copyright"].replace("\n", "")}"
            except:
                pass

        else:  # API request failed.  Update system log and return failed-execution indication to the calling function:
            update_system_log("get_astronomy_pic_of_the_day", "Error: API request failed. Data cannot be obtained at this time.")
            error_message = "API request failed. Data cannot be obtained at this time."

    except:  # An error has occurred.
        update_system_log("get_astronomy_pic_of_the_day", traceback.format_exc())
        error_message = "An error has occurred. Data cannot be obtained at this time."

    finally:
        # Return results to calling function:
        return json, copyright_details, error_message


def get_confirmed_planets():
    """Function for getting all needed data pertaining to confirmed planets and store such information in the space database supporting our website"""
    try:
        # Execute API request:
        response = requests.get(URL_CONFIRMED_PLANETS)
        if response.status_code == 200:
            # Delete the existing records in the "confirmed_planets" database table and update same with
            # the up-to-date data (from the JSON).  If execution failed, update system log and return
            # failed-execution indication to the calling function::
            # NOTE:  Scope of data: Solution Type = 'Published Confirmed'
            if not update_database("update_confirmed_planets", response.json()):
                update_system_log("get_confirmed_planets", "Error: Database could not be updated. Data cannot be obtained at this time.")
                return "Error: Database could not be updated. Data cannot be obtained at this time.", False

            # Retrieve all existing records in the "confirmed_planets" database table. If the function
            # called returns an empty directory, update system log and return failed-execution indication
            # to the calling function:
            confirmed_planets_data = retrieve_from_database("confirmed_planets")
            if confirmed_planets_data == {}:
                update_system_log("get_confirmed_planets", "Error: Data cannot be obtained at this time.")
                return "Error: Data cannot be obtained at this time.", False

            # If an empty list was returned, no records satisfied the query.  Therefore, update system log and return
            # failed-execution indication to the calling function:
            elif confirmed_planets_data == []:
                update_system_log("get_confirmed_planets", "No matching records were retrieved.")
                return "No matching records were retrieved.", False

            # Create and format a spreadsheet file (workbook) to contain all confirmed-planet data. If execution
            # failed, update system log and return failed-execution indication to the calling function:
            if not export_data_to_spreadsheet_standard("confirmed_planets", confirmed_planets_data):
                update_system_log("get_confirmed_planets", "Error: Spreadsheet creation could not be completed at this time.")
                return "Error: Spreadsheet creation could not be completed at this time.", False

            # At this point, function is deemed to have executed successfully.  Update system log and return
            # successful-execution indication to the calling function:
            update_system_log("get_confirmed_planets", "Successfully updated.")
            return "", True

        else:  # API request failed.  Update system log and return failed-execution indication to the calling function:
            update_system_log("get_confirmed_planets","Error: API request failed. Data cannot be obtained at this time.")
            return "Error: API request failed. Data cannot be obtained at this time.", False

    except:  # An error has occurred.
        update_system_log("get_confirmed_planets", traceback.format_exc())

        # Return failed-execution indication to the calling function:
        return "An error has occurred. Data cannot be obtained at this time.", False



def get_inspiration_data():
    """Function for getting all needed data pertaining to insp and store such information in the space database supporting our website"""

    try:
        # Obtain a list of constellation using the skyfield.api library:
        constellations = dict(load_constellation_names())

        # If a constellation list has been obtained:
        if constellations != {}:
            # Get the nickname for each constellation identified.  If the function called returns an empty directory,
            # update system log and return failed-execution indication to the calling function:
            constellations_data = get_constellation_data_nicknames(constellations)
            if constellations_data == {}:
                update_system_log("get_constellation_data", "Error: Data (nicknames) cannot be obtained at this time.")
                return "Error: Data (nicknames) cannot be obtained at this time.", False

            # Get additional details for each constellation identified.  If the function called returns an empty directory,
            # update system log and return failed-execution indication to the calling function:
            constellations_added_details = get_constellation_data_added_details(constellations)
            if constellations_added_details == {}:
                update_system_log("get_constellation_data",
                                  "Error: Data (added details) cannot be obtained at this time.")
                return "Error: Data (added details) cannot be obtained at this time.", False

            # Get area for each constellation identified.  If the function called returns an empty directory,
            # update system log and return failed-execution indication to the calling function:
            constellations_area = get_constellation_data_area(constellations)
            if constellations_area == {}:
                update_system_log("get_constellation_data", "Error: Data (areas) cannot be obtained at this time.")
                return "Error: Data (areas) cannot be obtained at this time.", False

            # Add the additional details (including area) to the main constellation dictionary:
            for key in constellations_data:
                constellations_data[key]["area"] = constellations_area[key]["area"]
                constellations_data[key]["myth_assoc"] = constellations_added_details[key]["myth_assoc"]
                constellations_data[key]["first_appear"] = constellations_added_details[key]["first_appear"]
                constellations_data[key]["brightest_star_name"] = constellations_added_details[key][
                    "brightest_star_name"]
                constellations_data[key]["brightest_star_url"] = constellations_added_details[key]["brightest_star_url"]

            # Delete the existing records in the "constellations" database table and update same with the
            # contents of the "constellations_data" dictionary.  If the function called returns a failed-execution
            # indication, update system log and return failed-execution indication to the calling function:
            if not update_database("update_constellations", constellations_data):
                update_system_log("get_constellation_data",
                                  "Error: Database could not be updated. Data cannot be obtained at this time.")
                return "Error: Database could not be updated. Data cannot be obtained at this time.", False

            # Retrieve all existing records in the "constellations" database table. If the function
            # called returns an empty directory, update system log and return failed-execution indication to the
            # calling function:
            constellations_data = retrieve_from_database("constellations")
            if constellations_data == {}:
                update_system_log("get_constellation_data", "Error: Data cannot be obtained at this time.")
                return "Error: Data cannot be obtained at this time.", False

            # Create and format a spreadsheet file (workbook) to contain all constellation data. If the function called returns
            # a failed-execution indication, update system log and return a failed-execution indication to the calling function:
            if not export_data_to_spreadsheet_standard("constellations", constellations_data):
                update_system_log("get_constellation_data",
                                  "Error: Spreadsheet creation could not be completed at this time.")
                return "Error: Spreadsheet creation could not be completed at this time.", False

            # At this point, function is deemed to have executed successfully.  Update system log and return
            # successful-execution indication to the calling function:
            update_system_log("get_constellation_data", "Successfully updated.")
            return "", True

        else:  # An error has occurred in processing constellation data.
            update_system_log("get_constellation_data", "Error: Data cannot be obtained at this time.")
            return "Error: Data cannot be obtained at this time.", False

    except:  # An error has occurred.
        update_system_log("get_constellation_data", traceback.format_exc())

        # Return failed-execution indication to the calling function:
        return "An error has occurred. Data cannot be obtained at this time.", False




def get_constellation_data():
    """Function for getting (via web-scraping) inspirational data and storing such information in the database supporting our website"""

    try:
        # If a constellation list has been obtained:
        if constellations != {}:
            # Get the nickname for each constellation identified.  If the function called returns an empty directory,
            # update system log and return failed-execution indication to the calling function:
            constellations_data = get_constellation_data_nicknames(constellations)
            if constellations_data == {}:
                update_system_log("get_constellation_data", "Error: Data (nicknames) cannot be obtained at this time.")
                return "Error: Data (nicknames) cannot be obtained at this time.", False

            # Get additional details for each constellation identified.  If the function called returns an empty directory,
            # update system log and return failed-execution indication to the calling function:
            constellations_added_details = get_constellation_data_added_details(constellations)
            if constellations_added_details == {}:
                update_system_log("get_constellation_data",
                                  "Error: Data (added details) cannot be obtained at this time.")
                return "Error: Data (added details) cannot be obtained at this time.", False

            # Get area for each constellation identified.  If the function called returns an empty directory,
            # update system log and return failed-execution indication to the calling function:
            constellations_area = get_constellation_data_area(constellations)
            if constellations_area == {}:
                update_system_log("get_constellation_data", "Error: Data (areas) cannot be obtained at this time.")
                return "Error: Data (areas) cannot be obtained at this time.", False

            # Add the additional details (including area) to the main constellation dictionary:
            for key in constellations_data:
                constellations_data[key]["area"] = constellations_area[key]["area"]
                constellations_data[key]["myth_assoc"] = constellations_added_details[key]["myth_assoc"]
                constellations_data[key]["first_appear"] = constellations_added_details[key]["first_appear"]
                constellations_data[key]["brightest_star_name"] = constellations_added_details[key][
                    "brightest_star_name"]
                constellations_data[key]["brightest_star_url"] = constellations_added_details[key]["brightest_star_url"]

            # Delete the existing records in the "constellations" database table and update same with the
            # contents of the "constellations_data" dictionary.  If the function called returns a failed-execution
            # indication, update system log and return failed-execution indication to the calling function:
            if not update_database("update_constellations", constellations_data):
                update_system_log("get_constellation_data",
                                  "Error: Database could not be updated. Data cannot be obtained at this time.")
                return "Error: Database could not be updated. Data cannot be obtained at this time.", False

            # Retrieve all existing records in the "constellations" database table. If the function
            # called returns an empty directory, update system log and return failed-execution indication to the
            # calling function:
            constellations_data = retrieve_from_database("constellations")
            if constellations_data == {}:
                update_system_log("get_constellation_data", "Error: Data cannot be obtained at this time.")
                return "Error: Data cannot be obtained at this time.", False

            # Create and format a spreadsheet file (workbook) to contain all constellation data. If the function called returns
            # a failed-execution indication, update system log and return a failed-execution indication to the calling function:
            if not export_data_to_spreadsheet_standard("constellations", constellations_data):
                update_system_log("get_constellation_data",
                                  "Error: Spreadsheet creation could not be completed at this time.")
                return "Error: Spreadsheet creation could not be completed at this time.", False

            # At this point, function is deemed to have executed successfully.  Update system log and return
            # successful-execution indication to the calling function:
            update_system_log("get_constellation_data", "Successfully updated.")
            return "", True

        else:  # An error has occurred in processing constellation data.
            update_system_log("get_constellation_data", "Error: Data cannot be obtained at this time.")
            return "Error: Data cannot be obtained at this time.", False

    except:  # An error has occurred.
        update_system_log("get_constellation_data", traceback.format_exc())

        # Return failed-execution indication to the calling function:
        return "An error has occurred. Data cannot be obtained at this time.", False


def get_constellation_data_added_details(constellations):
    """Function for getting (via web-scraping) additional details for each constellation identified"""

    try:
        # Define a variable for storing the additional details for each constellation (to be scraped from the constellation map website):
        constellations_added_details = {}

        # Constellation "Serpens" is represented via 2 separate entries in the target website (head & tail). Accordingly, define variables to be used
        # as part of the workaround to handle this constellation's data differently than the rest:
        serpens_element_constellation_myth_assoc_text = ""
        serpens_element_constellation_first_appear_text = ""
        serpens_element_constellation_brightest_star_text = ""
        serpens_element_constellation_brightest_star_url = ""

        # Initiate and configure a Selenium object to be used for scraping website for additional constellation details.
        # If function failed, update system log and return failed-execution to the calling function:
        driver = setup_selenium_driver(URL_CONSTELLATION_ADD_DETAILS_1, 1, 1)
        if driver == None:
            update_system_log("get_constellation_data_added_details", "Error: Selenium driver could not be created/configured.")
            return {}

        # Pause program execution to allow for constellation website loading time:
        time.sleep(WEB_LOADING_TIME_ALLOWANCE)

        # Define special variables to handle the 'Serpens' constellation whose data spans 2 entries (head/tail) on the target website:
        serpens_index = 0
        serpens_list = ["Head: ", "Tail: "]

        # Scrape the constellation map website to obtain additional details for each constellation:
        for i in range(1,
                       len(constellations) + 1 + 1):  # Added 1 because the constellation "Serpens" is rep'd by two separate entries on this website
            # Find the element pertaining to the constellation's name. Decode it to normalize to ASCII-based characters:
            element_constellation_name = find_element(driver, "xpath",
                                                      '/html/body/div/div[3]/div[1]/div[1]/div/div[3]/div[2]/table/tbody/tr[' + str(
                                                          i) + ']/td[1]/a')
            element_constellation_name_unidecoded = unidecode.unidecode(
                element_constellation_name.get_attribute("text"))

            # Find the element pertaining to the constellation's mythological association:
            element_constellation_myth_assoc = find_element(driver, "xpath",
                                                            '/html/body/div/div[3]/div[1]/div[1]/div/div[3]/div[2]/table/tbody/tr[' + str(
                                                                i) + ']/td[2]/div')
            element_constellation_myth_assoc_text = element_constellation_myth_assoc.get_attribute("innerHTML")

            # Find the element pertaining to the constellation's first appearance:
            element_constellation_first_appear = find_element(driver, "xpath",
                                                              '/html/body/div/div[3]/div[1]/div[1]/div/div[3]/div[2]/table/tbody/tr[' + str(
                                                                  i) + ']/td[3]/div')
            element_constellation_first_appear_text = element_constellation_first_appear.get_attribute("innerHTML")

            # Find the element pertaining to the constellation's brightest star.  Capture both text and url:
            element_constellation_brightest_star = find_element(driver, "xpath",
                                                                '/html/body/div/div[3]/div[1]/div[1]/div/div[3]/div[2]/table/tbody/tr[' + str(
                                                                    i) + ']/td[5]/a')
            element_constellation_brightest_star_text = element_constellation_brightest_star.get_attribute(
                "text").replace(" ", "").replace("\n", "")
            element_constellation_brightest_star_url = element_constellation_brightest_star.get_attribute("href")

            # Add the additional details collected above to the "constellation added details" dictionary:
            if "Serpens" in element_constellation_name_unidecoded:  # Constellation "Serpens" is represented via 2 separate entries in the target website (head & tail).
                serpens_element_constellation_myth_assoc_text += serpens_list[
                                                                     serpens_index] + element_constellation_myth_assoc_text + " "
                serpens_element_constellation_first_appear_text += serpens_list[
                                                                       serpens_index] + element_constellation_first_appear_text + " "
                serpens_element_constellation_brightest_star_text += serpens_list[
                                                                         serpens_index] + element_constellation_brightest_star_text + " "
                serpens_element_constellation_brightest_star_url += element_constellation_brightest_star_url + " "

                constellations_added_details["Serpens"] = {
                    "myth_assoc": serpens_element_constellation_myth_assoc_text,
                    "first_appear": serpens_element_constellation_first_appear_text,
                    "brightest_star_name": serpens_element_constellation_brightest_star_text,
                    "brightest_star_url": serpens_element_constellation_brightest_star_url
                }

                serpens_index += 1

            else:
                constellations_added_details[element_constellation_name_unidecoded] = {
                    "myth_assoc": element_constellation_myth_assoc_text,
                    "first_appear": element_constellation_first_appear_text,
                    "brightest_star_name": element_constellation_brightest_star_text,
                    "brightest_star_url": element_constellation_brightest_star_url
                }

        # Sort the "constellations_added_details" dictionary in alphabetical order by its key (the constellation's name):
        constellations_added_details = collections.OrderedDict(sorted(constellations_added_details.items()))

        # Close and delete the Selenium driver object:
        driver.close()
        del driver

        # Return the populated "constellations_added_details" dictionary to the calling function:
        return constellations_added_details

    except:  # An error has occurred.
        update_system_log("get_constellation_data_added_details", traceback.format_exc())

        # Return empty directory as a failed-execution indication to the calling function:
        return {}


def get_constellation_data_area(constellations):
    """Function for getting (via web-scraping) the area for each constellation identified"""

    try:
        # Define a variable for storing the area for each constellation (to be scraped from the constellation map website):
        constellations_area = {}

        # Constellation "Serpens" is represented via 2 separate entries in the target website (head & tail). Accordingly, define variable to be used
        # as part of the workaround to handle this constellation's data differently than the rest:
        serpens_element_constellation_area_text = ""

        # Initiate and configure a Selenium object to be used for scraping website for area (page 1).
        # If function failed, update system log and return failed-execution indication to the calling function:
        driver = setup_selenium_driver(URL_CONSTELLATION_ADD_DETAILS_2A, 1, 1)
        if driver == None:
            update_system_log("get_constellation_data_area", "Error: Selenium driver could not be created/configured.")
            return {}

        # Pause program execution to allow for constellation website loading time:
        time.sleep(WEB_LOADING_TIME_ALLOWANCE)

        # Define special variables to handle the 'Serpens' constellation whose data spans 2 entries (head/tail) on the target website:
        serpens_index = 0
        serpens_list = ["Head: ", "Tail: "]

        # Scrape the constellation map website to obtain additional details for each constellation:
        for i in range(1,
                       51):  # Added 1 because the constellation "Serpens" is rep'd by two separate entries on this website

            # Find the element pertaining to the constellation's name. Decode it to normalize to ASCII-based characters:
            try:
                element_constellation_name = find_element(driver, "xpath",
                                                          '/html/body/div/div[3]/div[1]/div[1]/div/div[4]/div[2]/table/tbody/tr[' + str(
                                                              i) + ']/td[1]/a')
            except:  # An added section (e.g., an ad) may displace all subsequent elements by one.
                element_constellation_name = find_element(driver, "xpath",
                                                          '/html/body/div/div[3]/div[1]/div[1]/div/div[5]/div[2]/table/tbody/tr[' + str(
                                                              i) + ']/td[1]/a')
            element_constellation_name_unidecoded = unidecode.unidecode(
                element_constellation_name.get_attribute("text"))

            # Find the element pertaining to the constellation's area. Decode it to normalize to ASCII-based characters:
            try:
                element_constellation_area = find_element(driver, "xpath",
                                                          '/html/body/div/div[3]/div[1]/div[1]/div/div[4]/div[2]/table/tbody/tr[' + str(
                                                              i) + ']/td[2]')
            except:  # An added section (e.g., an ad) may displace all subsequent elements by one.
                element_constellation_area = find_element(driver, "xpath",
                                                          '/html/body/div/div[3]/div[1]/div[1]/div/div[5]/div[2]/table/tbody/tr[' + str(
                                                              i) + ']/td[2]')
            element_constellation_area_text = unidecode.unidecode(
                element_constellation_area.get_attribute("innerHTML")).replace("&nbsp;", " ")

            # Add the area collected above to the "constellations_area" dictionary:
            if "Serpens" in element_constellation_name_unidecoded:  # Constellation "Serpens" is represented via 2 separate entries in the target website (head & tail).
                serpens_element_constellation_area_text += serpens_list[
                                                               serpens_index] + element_constellation_area_text + " "

                constellations_area["Serpens"] = {
                    "area": serpens_element_constellation_area_text
                }

                serpens_index += 1

            else:
                constellations_area[element_constellation_name_unidecoded] = {
                    "area": element_constellation_area_text,
                }

        # Close and delete the Selenium driver object:
        driver.close()
        del driver

        # Initiate and configure a Selenium object to be used for scraping website for area (page 2).
        # If function failed, update system log and return failed-execution indication to the calling function:
        driver = setup_selenium_driver(URL_CONSTELLATION_ADD_DETAILS_2B, 1, 1)
        if driver == None:
            update_system_log("get_constellation_data_area", "Error: Selenium driver could not be created/configured.")
            return {}

        # Pause program execution to allow for constellation website loading time:
        time.sleep(WEB_LOADING_TIME_ALLOWANCE)

        # Scrape the constellation map website to obtain additional details for each constellation:
        for i in range(51,
                       len(constellations) + 1 + 2):  # Added 1 because the constellation "Serpens" is rep'd by two separate entries on this website, and added another because website contains an "Unknown constellation" that should not detract from reaching the end of the "constellations_data" dictionary.

            # Find the element pertaining to the constellation's name. Decode it to normalize to ASCII-based characters:
            try:
                element_constellation_name = find_element(driver, "xpath",
                                                          '/html/body/div/div[3]/div[1]/div[1]/div/div[4]/div[2]/table/tbody/tr[' + str(
                                                              i - 50) + ']/td[1]/a')
            except:  # An added section (e.g., an ad) may displace all subsequent elements by one.
                element_constellation_name = find_element(driver, "xpath",
                                                          '/html/body/div/div[3]/div[1]/div[1]/div/div[5]/div[2]/table/tbody/tr[' + str(
                                                              i - 50) + ']/td[1]/a')
            element_constellation_name_unidecoded = unidecode.unidecode(
                element_constellation_name.get_attribute("text"))

            # Find the element pertaining to the constellation's area. Decode it to normalize to ASCII-based characters:
            try:
                element_constellation_area = find_element(driver, "xpath",
                                                          '/html/body/div/div[3]/div[1]/div[1]/div/div[4]/div[2]/table/tbody/tr[' + str(
                                                              i - 50) + ']/td[2]')
            except:  # An added section (e.g., an ad) may displace all subsequent elements by one.
                element_constellation_area = find_element(driver, "xpath",
                                                          '/html/body/div/div[3]/div[1]/div[1]/div/div[5]/div[2]/table/tbody/tr[' + str(
                                                              i - 50) + ']/td[2]')
            element_constellation_area_text = unidecode.unidecode(
                element_constellation_area.get_attribute("innerHTML")).replace("&nbsp;", " ")

            # Add the area collected above to the "constellations_area" dictionary:
            if "Serpens" in element_constellation_name_unidecoded:  # Constellation "Serpens" is represented via 2 separate entries in the target website (head & tail).
                serpens_element_constellation_area_text += serpens_list[
                                                               serpens_index] + element_constellation_area_text + " "

                constellations_area["Serpens"] = {
                    "area": serpens_element_constellation_area_text
                }

                serpens_index += 1

            else:
                constellations_area[element_constellation_name_unidecoded] = {
                    "area": element_constellation_area_text,
                }

        # Close and delete the Selenium driver object:
        driver.close()
        del driver

        # Sort the "constellations_area" dictionary in alphabetical order by its key (the constellation's name):
        constellations_area = collections.OrderedDict(sorted(constellations_area.items()))

        # Return the populated "constellations_area" dictionary to the calling function:
        return constellations_area

    except:  # An error has occurred.
        update_system_log("get_constellation_data_area", traceback.format_exc())

        # Return empty directory as a failed-execution indication to the calling function:
        return {}


def get_constellation_data_nicknames(constellations):
    """Function for getting (via web-scraping) the nickname for each constellation identified"""

    try:
        # Define a variable for storing the final (sorted) dictionary of data for each constellation
        # (for a better-formatted JSON without the "OrderedDict" qualifier):
        constellations_data = {}

        # Initiate and configure a Selenium object to be used for scraping the constellation map website.
        # If function failed, update system log and return failed-execution indication to the calling function:
        driver = setup_selenium_driver(URL_CONSTELLATION_MAP_SITE, 1, 1)
        if driver == None:
            update_system_log("get_constellation_data_nicknames", "Error: Selenium driver could not be created/configured.")
            return {}

        # Pause program execution to allow for constellation website loading time:
        time.sleep(WEB_LOADING_TIME_ALLOWANCE)

        # Define a variable for storing the nicknames of each constellation (to be scraped from the constellation map website):
        constellation_nicknames = {}

        # Scrape the constellation map website to obtain the nicknames of each constellation:
        for i in range(1, len(constellations) + 1):
            try:
                # Find the element pertaining to the constellation's name:
                element_constellation_name = find_element(driver, "xpath",
                                                          '/html/body/div[3]/section[2]/div/div/div/div[1]/div[' + str(
                                                              i) + ']/div/article/div[2]/header/h2/a')

            except:  # Some of the constellations may use a different path than the above.
                # Find the element pertaining to the constellation's name:
                element_constellation_name = find_element(driver, "xpath",
                                                          '/html/body/div[3]/section[2]/div/div/div/div[1]/div[' + str(
                                                              i) + ']/div/article/div[2]/header/h3/a')

            # From the scraping performed above, decode the constellation's name to normalize to ASCII-based characters:
            element_constellation_name_unidecoded = unidecode.unidecode(element_constellation_name.text)

            # Find the element pertaining to the constellation's nickname. Decode it to normalize to ASCII-based characters:
            element_constellation_nickname = find_element(driver, "xpath",
                                                          '/html/body/div[3]/section[2]/div/div/div/div[1]/div[' + str(
                                                              i) + ']/div/article/div[2]/div/p')
            element_constellation_nickname_unidecoded = unidecode.unidecode(element_constellation_nickname.text)

            # Add the nickname to the "constellation nicknames" dictionary:
            constellation_nicknames[element_constellation_name_unidecoded] = element_constellation_nickname_unidecoded

        # Sort the "constellation_nicknames" dictionary in alphabetical order by its key (the constellation's name):
        constellation_nicknames = collections.OrderedDict(sorted(constellation_nicknames.items()))

        # Close and delete the Selenium driver object:
        driver.close()
        del driver

        # Define a variable for storing the (unsorted) dictionary of data for each constellation:
        constellations_unsorted = {}

        # For each constellation identified, prepare its dictionary entry:
        for key in constellations:
            constellations_unsorted[constellations[key]] = {"abbreviation": key,
                                                            "nickname": constellation_nicknames[constellations[key]],
                                                            "url": "https://www.go-astronomy.com/constellations.php?Name=" +
                                                                   constellations[key].replace(" ", "%20")}

        # Sort the (unsorted) dictionary in alphabetical order by its key (the constellation's name):
        constellations_sorted = collections.OrderedDict(sorted(constellations_unsorted.items()))

        # For each constellation identified, prepare its dictionary entry:
        for key in constellations_sorted:
            constellations_data[key] = {"abbreviation": constellations_sorted[key]["abbreviation"],
                                        "nickname": constellations_sorted[key]["nickname"],
                                        "url": constellations_sorted[key]["url"]}

        # Return the populated "constellations_data" dictionary to the calling function:
        return constellations_data

    except:  # An error has occurred.
        update_system_log("get_constellation_data_nicknames", traceback.format_exc())

        # Return empty directory as a failed-execution indication to the calling function:
        return {}


def get_iss_location():
    """Function to retrieve the current location of the ISS and a link to view the map of same"""
    # Initialize variables to be used for returning values to the calling function:
    location_address = ""
    location_url = ""

    try:
        # Execute API request:
        response = requests.get(URL_ISS_LOCATION)

        # If the API request was successful, capture and process the results:
        if response.status_code == 200:
            latitude = response.json()["iss_position"]["latitude"]
            longitude = response.json()["iss_position"]["longitude"]

            # Execute API request (using the retrieved latitude and longitude), to
            # get a link to a map of the ISS's current location:
            url = URL_GET_LOC_FROM_LAT_AND_LON + "?lat=" + str(latitude) + "&lon=" + str(
                longitude) + "&api_key=" + API_KEY_GET_LOC_FROM_LAT_AND_LON
            response = requests.get(url)

            # If the API request was successful, capture and process the results:
            if response.status_code == 200:  # API request was successful.
                for key in response.json():
                    if key == "error":  # Resulting JSON has an error key (possibly due to current location being over water).
                        if response.json()["error"] == "Unable to geocode":  # ISS may currently be over water.
                            location_address = "No terrestrial address is available.  ISS could be over water at the current time."

                    else:  # Terrestrial address is available.
                        # Display terrestrial address:
                        location_address = response.json()["display_name"]

                    # Break from the 'for' loop:
                    break

                # Prepare and display a link that points to the ISS's current location:
                location_url = "https://maps.google.com/?q=" + str(latitude) + "," + str(longitude)

        else:  # API request failed.  Update system log and return failed-execution indication to the calling function:
            update_system_log("get_iss_location", "Error: API request failed. Data cannot be obtained at this time.")
            location_address = "API request failed. Data cannot be obtained at this time."
            location_url = ""

    except:  # An error has occurred.
        update_system_log("get_iss_location", traceback.format_exc())
        location_address = "An error has occurred. Data cannot be obtained at this time."
        location_url = ""

    finally:
        # Return location address and URL to the calling function:
        return location_address, location_url


def get_mars_photos():
    """Function to retrieve summary and detailed data pertaining to the photos taken by each rover exploring on Mars"""
    global mars_rovers

    try:

        # Delete all existing files pertaining to this scope. If the function failed, update system
        # log and return failed-execution indication to the calling function:
        if not delete_mars_photos_workbooks():
            update_system_log("get_mars_photos", "Error: Previous spreadsheet file(s) could not be deleted.")
            return "Error: Previous spreadsheet file(s) could not be deleted.", False

        # Retrieve, from the database, a list of all rovers that are currently active for purposes of
        # data production.  If the function called returns an empty directory, update system log and return
        # failed-execution indication to the calling function:
        mars_rovers_from_db = retrieve_from_database("mars_rovers")
        if mars_rovers_from_db == {}:
            update_system_log("get_mars_photos", "Error: Data (Mars rovers) cannot be obtained at this time.")
            return "Error: Data (Mars rovers) cannot be obtained at this time.", False

        # If an empty list was returned, no records satisfied the query.  Therefore, update system log and
        # return failed-execution indication to the calling function:
        elif mars_rovers_from_db == []:
            update_system_log("get_mars_photos", "No matching records were retrieved (Mars rovers).")
            return "No matching records were retrieved (Mars rovers).", False

        # Populate "mars_rovers" with list of active rovers per the database:
        for record in mars_rovers_from_db:
            mars_rovers.append(record.rover_name)

        # Inform user that database will be checked for updates:
        dlg = PBI.PyBusyInfo("Photos from Mars: Checking for updates needed...", title="Administrative Update")

        # Prepare a dictionary which summarizes photos available by rover and earth date. If the function returns
        # an empty dictionary, update system log and return failed-execution indication to the calling function:
        photos_available = get_mars_photos_summarize_photos_available({})
        if photos_available == {}:
            update_system_log("get_mars_photos", "Error: Data (photos, summarize available) cannot be obtained at this time.")
            return "Error: Data (photos, summarize available) cannot be obtained at this time.", False

        # Obtain respective dictionaries summarizing photos available and the corresponding contents of the
        # "mars_photo_details" database table. If the function returns an empty directory for the former,
        # update system log and return failed-execution indication to the calling function:
        photos_available_summary, photo_details_summary = retrieve_from_database("mars_photo_details_compare_with_photos_available")
        if photos_available_summary == {}:
            update_system_log("get_mars_photos", "Error: Data (photos, pre-update check) cannot be obtained at this time.")
            return "Error: Data (photos, pre-update check) cannot be obtained at this time.", False

        # If an empty list was returned for "photos available", no records satisfied the query.  Therefore,
        # update system log and return failed-execution indication to the calling function:
        elif photos_available_summary == []:
            update_system_log("get_mars_photos", "No matching records were retrieved (photos, pre-update check).")
            return "No matching records were retrieved (photos, pre-update check).", False

        else:
            # Initialize a variable for capturing rover/earth date combinations for which there is a mismatch
            # between the photos available and the corresponding photo details:
            rover_earth_date_combo_mismatch_between_summaries = []

            # Compare photos available with the corresponding contents of the "mars_photo_details" database table:
            if photos_available_summary == photo_details_summary:  # Database is up to date.  No API requests are needed.
                dlg = PBI.PyBusyInfo("Photos from Mars: Database is up to date. Proceeding to export results to spreadsheet files...", title="Administrative Update")

            else:  # Database (specifically the "mars_photo_details" needs updating.
                dlg = PBI.PyBusyInfo("Photos from Mars: Photo details table needs updating.  Update in progress...", title="Administrative Update")

                # Capture a list of the rover/earth date combinations for which there is a mismatch
                # between the photos available and the corresponding photo details:
                for i in range(0, len(photos_available_summary)):
                    if not(photos_available_summary[i] in photo_details_summary):
                        rover_earth_date_combo_mismatch_between_summaries.append(photos_available_summary[i][0])

            # Perform required database updates based on whether any mismatches were identified above.
            # If the function called returns a failed-execution indication, update system log and return
            # failed-execution indication to the calling function:
            if not get_mars_photos_update_database(photos_available, rover_earth_date_combo_mismatch_between_summaries):
                update_system_log("get_mars_photos", "Error: Data (photos, post-details-update) cannot be obtained at this time.")
                return "Error: Data (photos, post-details-update) cannot be obtained at this time.", False

        # Provide user an update before proceeding to export results to spreadsheet files:
        dlg = PBI.PyBusyInfo("Photos from Mars: Proceeding to export results to spreadsheet files...", title="Administrative Update")

        # Retrieve a list of records from the "mars_photos_available" database table.  If the function
        # called returns a failed-execution indication (i.e., an empty dictionary), update system log and
        # return failed-execution indication to the calling function:
        photos_available = retrieve_from_database("mars_photos_available")
        if photos_available == {}:
            update_system_log("get_mars_photos", "Error: Data (photos, post-update) cannot be obtained at this time.")
            return "Error: Data (photo, post-update) cannot be obtained at this time.", False

        # If an empty list was returned, no records satisfied the query.  Therefore, update system log and
        # return failed-execution indication to the calling function:
        elif photos_available == []:
            update_system_log("get_mars_photos", "No matching records were retrieved (photos, post-update).")
            return "No matching records were retrieved (photos, post-update).", False

        # Retrieve a list of records from the "mars_photo_details database table.  If the function
        # called returns a failed-execution indication (i.e., an empty dictionary), update system log
        # and return failed-execution indication to the calling function:
        photo_details = retrieve_from_database("mars_photo_details")
        if photo_details == {}:
            update_system_log("get_mars_photos", "Error: Data cannot be obtained at this time.")
            return "Error: Data cannot be obtained at this time.", False

        # If an empty list was returned, no records satisfied the query.  Therefore, update system log
        # and return failed-execution indication to the calling function:
        elif photo_details == []:
            update_system_log("get_mars_photos", "No matching records were retrieved.")
            return "No matching records were retrieved.", False

        # Export collected summary and detailed results to a spreadsheet workbook:
        if not export_mars_photos_to_spreadsheet(photos_available, photo_details):
            update_system_log("get_mars_photos","Error: Spreadsheet creation could not be completed at this time.")
            return "Error: Spreadsheet creation could not be completed at this time.", False

        # At this point, function is deemed to have executed successfully.  Update system log and return
        # successful-execution indication to the calling function:
        update_system_log("get_mars_photos", "Successfully updated.")
        return "", True

    except:  # An error has occurred.
        update_system_log("get_mars_photos", traceback.format_exc())

        # Return failed-execution indication to the calling function:
        return "An error has occurred. Data cannot be obtained at this time.", False


def get_mars_photos_summarize_photo_counts_by_rover_and_earth_year():
    """Function to summarize photo counts by rover and earth year.  This supports final spreadsheet creation"""
    try:
        # Get counts (by rover name and earth date) from "mars_photo_details" database table. If the
        # function called returns a failed-execution indication (i.e., an empty dictionary), update system log
        # and return failed-execution indication to the calling function:
        photo_counts = retrieve_from_database("mars_photo_details_get_counts_by_rover_and_earth_date")
        if photo_counts == {}:
            update_system_log("get_mars_photos_summarize_photo_counts_by_rover_and_earth_year",
                              "Error: Data could not be obtained at this time.")
            return []

        # If an empty list was returned, no records satisfied the query.  Therefore, update system log and return
        # failed-execution indication to the calling function:
        elif photo_counts == []:
            update_system_log("get_mars_photos_summarize_photo_counts_by_rover_and_earth_year",
                              "Error: No matching records were retrieved.")
            return []

        # Initialize list variables needed to produce the final results to the calling function:
        photo_count_grouping_1 = []
        photo_count_grouping_2 = []

        # Add a rover name/earth year combo value to the results obtained above, and add all to a list:
        for item in photo_counts:
            photo_count_grouping_1.append([item[0], item[1], item[0] + "_" + item[1][:4], item[2]])

        # Iterate through the list created above, summarize photo counts by rover name/earth year combo, and
        # populate list with summarized data:
        # Capture the first row of data:
        rover_name_a = photo_count_grouping_1[0][0]
        earth_year_a = photo_count_grouping_1[0][1][:4]
        rover_earth_date_combo_a = photo_count_grouping_1[0][2]
        total_photos_a = photo_count_grouping_1[0][3]

        # Capture the next row of date.  Compare the rover/earth year combo with the combo from
        # the previous row.  Iterate through this process until the end of the data set has been
        # reached:
        for i in range(1, len(photo_count_grouping_1)):
            # Capture the next row of data:
            rover_name_b = photo_count_grouping_1[i][0]
            earth_year_b = photo_count_grouping_1[i][1][:4]
            rover_earth_date_combo_b = photo_count_grouping_1[i][2]
            total_photos_b = photo_count_grouping_1[i][3]

            # Compare the rover/earth year combo with the combo from the previous row:
            if rover_earth_date_combo_b != rover_earth_date_combo_a:  # New rover/earth year combo has been reached.
                # Append the final photo count for the previous row (whose final row has been reached) to the
                # final list to be returned to the calling function:
                photo_count_grouping_2.append([rover_name_a, earth_year_a, rover_earth_date_combo_a, total_photos_a])
                rover_name_a = rover_name_b
                earth_year_a = earth_year_b
                rover_earth_date_combo_a = rover_earth_date_combo_b
                total_photos_a = total_photos_b
            else:  # New rover/earth year combo has NOT been reached.
                # Continue tallying the running total for the current combo.
                total_photos_a += total_photos_b

        # Capture the final total photo count for the final rover/earth date combo (whose final row
        # has been reached), and append it to the final list to be returned to the calling function:
        photo_count_grouping_2.append([rover_name_a, earth_year_a, rover_earth_date_combo_a, total_photos_a])

        # Return resulting list to the calling function:
        return photo_count_grouping_2

    except:  # An error has occurred.
        update_system_log("get_mars_photos_summarize_photo_counts_by_rover_and_earth_year", traceback.format_exc())

        # Return empty dictionary as a failed-execution indication to the calling function:
        return []


def get_mars_photos_summarize_photos_available(photos_available):
    """Function to summarize photos available.  This supports final spreadsheet creation"""

    try:
        # Perform the following for each rover that is currently active:
        for rover_name in mars_rovers:
            # Execute the API request:
            url = URL_MARS_ROVER_PHOTOS_BY_ROVER + rover_name + "?api_key=" + API_KEY_MARS_ROVER_PHOTOS
            response = requests.get(url)

            # If API request was successful, capture desired data elements:
            if response.status_code == 200:  # API request was successful.
                i = 0
                for item in response.json()['photo_manifest']['photos']:
                    photos_available[rover_name + "_" + str(item["earth_date"])] = {
                        "sol": item["sol"],
                        "rover_name": rover_name,
                        "earth_date": item["earth_date"],
                        "total_photos": item['total_photos'],
                        "cameras": ','.join(item["cameras"])
                    }

                if photos_available == {}:
                    update_system_log("get_mars_photos_summarize_photos_available",f"No photos are available for Mars rover '{rover_name}'")
                    return {}

            else:  # API request failed.  Update system log and return failed-execution indication to the calling function:
                # Update system log and return failed-execution indication to the calling function:
                update_system_log("get_mars_photos_summarize_photos_available",f"API request failed. No photos are available for Mars rover '{rover_name}'")
                return {}

        # Delete the existing records in the "mars_photos_available" database table and update same with the
        # contents of the "photos_available" dictionary.  If the function returns a failed-execution
        # indication, update system log and return failed-execution indication to the calling function:
        if not update_database("update_mars_photos_available", photos_available):
            update_system_log("get_mars_photos_summarize_photos_available","Error: Database could not be updated. Data cannot be obtained at this time.")
            return {}

        # Return populated "photos_available" dictionary to the calling function:
        return photos_available

    except:  # An error has occurred.
        update_system_log("get_mars_photos_summarize_photos_available", traceback.format_exc())

        # Return empty dictionary as a failed-execution indication to the calling function:
        return {}


def get_mars_photos_update_database(photos_available, rover_earth_date_combo_mismatch_between_summaries):
    try:
        if len(rover_earth_date_combo_mismatch_between_summaries) > 0:
            # For each rover/earth date combo where there is a mismatch between what the API provided and what exists in the database,
            # perform necessary updates to the database to align with what the API indicates is current:
            for i in range(0, len(rover_earth_date_combo_mismatch_between_summaries)):
                dlg = PBI.PyBusyInfo(f"Photos from Mars: {i + 1} of {len(rover_earth_date_combo_mismatch_between_summaries)} rover/earth date combinations needing update ({round((i+1)/len(rover_earth_date_combo_mismatch_between_summaries) * 100, 1)} %)...", title="Administrative Update")
                photo_details_rover_earth_date_combo = []

                # Capture a count of how many records are in the database for the rover-earth date combo.
                # If function failed, update system log and return failed-execution indication to the calling function:
                existing_record_count = retrieve_from_database("mars_photo_details_rover_earth_date_combo_count",
                                                               rover_name=rover_earth_date_combo_mismatch_between_summaries[i].split("_")[0],
                                                               earth_date=rover_earth_date_combo_mismatch_between_summaries[i].split("_")[1])
                if existing_record_count == {}:
                    update_system_log("get_mars_photos_update_database",
                                      f"Error: Pre-update retrieval of existing detail records failed (Rover '{rover_earth_date_combo_mismatch_between_summaries[i].split("_")[0]}', Earth Date {rover_earth_date_combo_mismatch_between_summaries[i].split("_")[1]}).")
                    return "Error: Pre-update retrieval of existing detail records failed (Rover '{rover_earth_date_combo_mismatch_between_summaries[i].split('_')[0]}', Earth Date {rover_earth_date_combo_mismatch_between_summaries[i].split('_')[1]}).", False

                # Capture updated record count for the rover/earth date combo being processed, using the "photos_available" dictionary to represent
                # what the API has indicated is current:
                updated_record_count = photos_available[rover_earth_date_combo_mismatch_between_summaries[i]]["total_photos"]

                # Provide user a progress update:
                dlg = PBI.PyBusyInfo(f"Photos from Mars: Rover '{rover_earth_date_combo_mismatch_between_summaries[i].split("_")[0]}', Earth Date {rover_earth_date_combo_mismatch_between_summaries[i].split("_")[1]} - Total Photos in DB: {existing_record_count}\nRover '{rover_earth_date_combo_mismatch_between_summaries[i].split("_")[0]}', Earth Date {rover_earth_date_combo_mismatch_between_summaries[i].split("_")[1]} - Total Photos (updated from API): {updated_record_count}\nUpdate in progress...", title="Administrative Update")

                # Capture the updated record set (for the rover/earth-date combo being processed) from what the API provided:
                dict_to_add = get_mars_photos_update_from_api(rover_earth_date_combo_mismatch_between_summaries[i].split("_")[0], rover_earth_date_combo_mismatch_between_summaries[i].split("_")[1])
                if dict_to_add != {}:
                    # Delete existing records in DB for this rover/earth date combo.  If function failed,
                    # update system log and return failed-execution indication to the calling function:
                    if not update_database("update_mars_photo_details_delete_existing", {},
                                           rover_name=rover_earth_date_combo_mismatch_between_summaries[i].split("_")[0],
                                           earth_date=rover_earth_date_combo_mismatch_between_summaries[i].split("_")[1]):
                        update_system_log("get_mars_photos_update_database",
                                          f"Error: Pre-update deletion of existing detail records failed (Rover '{rover_earth_date_combo_mismatch_between_summaries[i].split("_")[0]}', Earth Date {rover_earth_date_combo_mismatch_between_summaries[i].split("_")[1]}).")
                        return "Error: Pre-update deletion of existing detail records failed (Rover '{rover_earth_date_combo_mismatch_between_summaries[i].split('_')[0]}', Earth Date {rover_earth_date_combo_mismatch_between_summaries[i].split('_')[1]}).", False

                    # Populate dictionary which will be used to update database with updated detail records for the rover/earth date combo being processed:
                    for j in range(0, len(dict_to_add)):
                        dict_to_add_sub = {
                            "rover_earth_date_combo": dict_to_add[j]["rover"]["name"] + "_" + dict_to_add[j][
                                "earth_date"],
                            "rover_name": dict_to_add[j]["rover"]["name"],
                            "sol": dict_to_add[j]["sol"],
                            "pic_id": dict_to_add[j]["id"],
                            "earth_date": dict_to_add[j]["earth_date"],
                            "camera_name": dict_to_add[j]["camera"]["name"],
                            "camera_full_name": dict_to_add[j]["camera"]["full_name"],
                            "url": dict_to_add[j]["img_src"]
                        }
                        photo_details_rover_earth_date_combo.append(dict_to_add_sub)

                    # Update the "mars_photo_details" database table with the contents of the "photo_details_rover_earth_date_combo" list.
                    # If the function called returns a failed-execution indication, If function failed, update system log and return
                    # failed-execution indication to the calling function:
                    if not update_database("update_mars_photo_details", photo_details_rover_earth_date_combo):
                        update_system_log("get_mars_photos_update_database",
                                          f"Error: Database could not be updated (photo details) (Rover '{rover_earth_date_combo_mismatch_between_summaries[i].split("_")[0]}', Earth Date {rover_earth_date_combo_mismatch_between_summaries[i].split("_")[1]}).")
                        return "Error: Database could not be updated (photo details) (Rover '{rover_earth_date_combo_mismatch_between_summaries[i].split('_')[0]}', Earth Date {rover_earth_date_combo_mismatch_between_summaries[i].split('_')[1]}).", False

                    else:
                        # Inform user that the update has been successfully completed:
                        dlg = PBI.PyBusyInfo(f"Photos from Mars: Rover '{rover_earth_date_combo_mismatch_between_summaries[i].split("_")[0]}', Earth Date {rover_earth_date_combo_mismatch_between_summaries[i].split("_")[1]} - Update complete.",title="Administrative Update")

        # At this point, function is deemed to have executed successfully.  Return successful-execution indication to the calling function:
        dlg = None
        return "", True

    except:  # An error has occurred.
        update_system_log("get_mars_photos_update_database", traceback.format_exc())

        # Return failed-execution indication to the calling function:
        return False


def get_mars_photos_update_from_api(rover_name, earth_date):
    """Function to retrieve, via an API request, photos available for a particular rover/earth date combination"""
    try:
        # Identify the URL which will be used as part of the API request:
        url = URL_MARS_ROVER_PHOTOS_BY_ROVER_AND_OTHER_CRITERIA + rover_name + "/photos/?api_key=" + API_KEY_MARS_ROVER_PHOTOS + "&earth_date=" + earth_date

        # Execute the API request.
        response = requests.get(url)
        if response.status_code == 200:  # API request was successful.
            # Return the retrieved JSON to the calling function:
            return response.json()['photos']

        else:  # API request failed.  Update system log and return failed-execution indication to the calling function:
            # Inform the user that the photos cannot be obtained at this time:
            update_system_log("get_mars_photos_update_from_api", "Error: API request failed. Data (for Rover: '{rover_name}', Earth Date {earth_date}) cannot be obtained at this time.")
            return {}

    except:  # An error has occurred.
        update_system_log("get_mars_photos_update_from_api", traceback.format_exc())

        # Return failed-execution indication to the calling function:
        return {}


def get_people_in_space_now():
    """Function that retrieves a list of people currently in space at the present moment"""
    try:
        # Execute the API request:
        response = requests.get(URL_PEOPLE_IN_SPACE_NOW)

        # If the API request was successful, display the results:
        if response.status_code == 200:  # API request was successful.
            # Sort the resulting JSON by person's name:
            people_in_space_now = collections.OrderedDict(response.json().items())

            # Return the sorted JSON to the calling function:
            return people_in_space_now["people"], True

        else:  # API request failed. Update system log and return failed-execution indication to the calling function:
            update_system_log("get_people_in_space_now", "Error: API request failed. Data cannot be obtained at this time.")
            return "Error: API request failed. Data cannot be obtained at this time.", False

    except:  # An error has occurred.
        update_system_log("get_people_in_space_now", traceback.format_exc())
        return "An error has occurred. Data cannot be obtained at this time.", False


def get_space_news():
    """Function for retrieving the latest space news articles"""
    try:
        # Initialize variables to return to calling function:
        success = True
        error_message = ""

        # Execute API request:
        response = requests.get(URL_SPACE_NEWS)
        if response.status_code == 200:
            # Delete the existing records in the "space_news" database table and update same with
            # the newly acquired articles (from the JSON).  If function failed, update system log
            # and return failed-execution indication to the calling function:
            if not update_database("update_space_news", response.json()['results']):
                update_system_log("update_space_news", "Error: Space news articles cannot be obtained at this time.")
                error_message = "Error: Space news articles cannot be obtained at this time."
                success = False

        else:  # API request failed. Update system log and return failed-execution calling function:
            update_system_log("get_space_news", "Error: API request failed. Data cannot be obtained at this time.")
            error_message = "API request failed. Space news articles cannot be obtained at this time."
            success = False

    except:  # An error has occurred.
        update_system_log("get_space_news", traceback.format_exc())
        error_message = "An error has occurred. Space news articles cannot be obtained at this time."
        success = False

    finally:
        # Return results to the calling function:
        return success, error_message


def prepare_spreadsheet_get_format(workbook, name):
    """Function for identifying the format to be used in formatting content in spreadsheet, based on the type of content involved"""
    # NOTE: Error handling is deferred to the calling function.
    if name == "column_headers":  # Column headers
        return workbook.add_format({"bold": 3, "underline": True, "font_name": "Calibri", "font_size": 11, 'text_wrap': True})

    elif name == "data":  # Main body of data (excluding columns to be treated as active URLs)
        return workbook.add_format({"bold": 0, "font_name": "Calibri", "font_size": 11, 'text_wrap': True})

    elif name == "url":  # URLs
        return workbook.add_format({"bold": 0, "font_color": "blue", "underline": 1, "font_name": "Calibri", "font_size": 11, 'text_wrap': True})

    elif name == "spreadsheet_header":  # Header info. (e.g., title, generation date/time) at beginning of spreadsheet
        return workbook.add_format({"bold": 3, "font_name": "Calibri", "font_size": 16})


def prepare_spreadsheet_main_contents(workbook, worksheet, name, **kwargs):
    """Function for adding and formatting spreadsheet content based on the type of content being worked on"""
    try:
        if name == "approaching_asteroids_data":
            # Capture optional arguments:
            list_name = kwargs.get("list_name", None)

            # Add/format main contents:
            i = 3
            for item in list_name:
                worksheet.write(i, 0, item.close_approach_date, prepare_spreadsheet_get_format(workbook, "data"))
                worksheet.write(i, 1, item.name, prepare_spreadsheet_get_format(workbook, "data"))
                worksheet.write(i, 2, str(item.id), prepare_spreadsheet_get_format(workbook, "data"))
                worksheet.write(i, 3, "{:.2f}".format(round(item.absolute_magnitude_h,2)), prepare_spreadsheet_get_format(workbook, "data"))
                worksheet.write(i, 4, "{:.2f}".format(round(item.estimated_diameter_km_min,2)), prepare_spreadsheet_get_format(workbook, "data"))
                worksheet.write(i, 5, "{:.2f}".format(round(item.estimated_diameter_km_max,2)), prepare_spreadsheet_get_format(workbook, "data"))
                if item.is_potentially_hazardous == 0:
                    worksheet.write(i, 6, "No", prepare_spreadsheet_get_format(workbook, "data"))
                elif item.is_potentially_hazardous == 1:
                    worksheet.write(i, 6, "Yes", prepare_spreadsheet_get_format(workbook, "data"))
                worksheet.write(i, 7, "{:.2f}".format(round(item.relative_velocity_km_per_s,2)), prepare_spreadsheet_get_format(workbook, "data"))
                worksheet.write(i, 8, "{:.2f}".format(round(item.miss_distance_km,2)), prepare_spreadsheet_get_format(workbook, "data"))
                worksheet.write(i, 9, item.orbiting_body, prepare_spreadsheet_get_format(workbook, "data"))
                if item.is_sentry_object == 0:
                    worksheet.write(i, 10, "No", prepare_spreadsheet_get_format(workbook, "data"))
                elif item.is_sentry_object == 1:
                    worksheet.write(i, 10, "Yes", prepare_spreadsheet_get_format(workbook, "data"))
                worksheet.write_url(i, 11, item.url, prepare_spreadsheet_get_format(workbook, "url"), tip="Click here for details.")
                i += 1

        elif name == "approaching_asteroids_headers":
            worksheet.write(2, 0, "Close Approach Date", prepare_spreadsheet_get_format(workbook, "column_headers"))
            worksheet.write(2, 1, "Name", prepare_spreadsheet_get_format(workbook, "column_headers"))
            worksheet.write(2, 2, "ID", prepare_spreadsheet_get_format(workbook, "column_headers"))
            worksheet.write(2, 3, "[H] Absolute Magnitude", prepare_spreadsheet_get_format(workbook, "column_headers"))
            worksheet.write(2, 4, "Estimated Diameter (km) - Min.", prepare_spreadsheet_get_format(workbook, "column_headers"))
            worksheet.write(2, 5, "Estimated Diameter (km) - Max.", prepare_spreadsheet_get_format(workbook, "column_headers"))
            worksheet.write(2, 6, "Is Potentially Hazardous?", prepare_spreadsheet_get_format(workbook, "column_headers"))
            worksheet.write(2, 7, "Relative Velocity (km/s)", prepare_spreadsheet_get_format(workbook, "column_headers"))
            worksheet.write(2, 8, "Miss Distance (km)", prepare_spreadsheet_get_format(workbook, "column_headers"))
            worksheet.write(2, 9, "Orbiting Body", prepare_spreadsheet_get_format(workbook, "column_headers"))
            worksheet.write(2, 10, "Is Sentry Object?", prepare_spreadsheet_get_format(workbook, "column_headers"))
            worksheet.write(2, 11, "URL for Details", prepare_spreadsheet_get_format(workbook, "column_headers"))

        elif name == "confirmed_planets_data":
            # Capture optional arguments:
            list_name = kwargs.get("list_name", None)

            # Add/format main contents:
            i = 3
            for item in list_name:
                worksheet.write(i, 0, item.host_name, prepare_spreadsheet_get_format(workbook, "data"))
                worksheet.write(i, 1, str(item.host_num_stars), prepare_spreadsheet_get_format(workbook, "data"))
                worksheet.write(i, 2, str(item.host_num_planets), prepare_spreadsheet_get_format(workbook, "data"))
                worksheet.write(i, 3, item.planet_name, prepare_spreadsheet_get_format(workbook, "data"))
                worksheet.write(i, 4, str(item.discovery_year), prepare_spreadsheet_get_format(workbook, "data"))
                worksheet.write(i, 5, item.discovery_method, prepare_spreadsheet_get_format(workbook, "data"))
                worksheet.write(i, 6, item.discovery_facility, prepare_spreadsheet_get_format(workbook, "data"))
                worksheet.write(i, 7, item.discovery_telescope, prepare_spreadsheet_get_format(workbook, "data"))
                worksheet.write_url(i, 8, item.url, prepare_spreadsheet_get_format(workbook, "url"), tip="Click here for details.")
                i += 1

        elif name == "confirmed_planets_headers":
            worksheet.write(2, 0, "Host Name", prepare_spreadsheet_get_format(workbook, "column_headers"))
            worksheet.write(2, 1, "# Stars", prepare_spreadsheet_get_format(workbook, "column_headers"))
            worksheet.write(2, 2, "# Planets", prepare_spreadsheet_get_format(workbook, "column_headers"))
            worksheet.write(2, 3, "Planet Name", prepare_spreadsheet_get_format(workbook, "column_headers"))
            worksheet.write(2, 4, "Discovery Year", prepare_spreadsheet_get_format(workbook, "column_headers"))
            worksheet.write(2, 5, "Discovery Method", prepare_spreadsheet_get_format(workbook, "column_headers"))
            worksheet.write(2, 6, "Discovery Facility", prepare_spreadsheet_get_format(workbook, "column_headers"))
            worksheet.write(2, 7, "Discovery Telescope", prepare_spreadsheet_get_format(workbook, "column_headers"))
            worksheet.write(2, 8, "URL for Details", prepare_spreadsheet_get_format(workbook, "column_headers"))

        elif name == "constellation_headers":
            worksheet.write(2, 0, "Name", prepare_spreadsheet_get_format(workbook, "column_headers"))
            worksheet.write(2, 1, "Abbv.", prepare_spreadsheet_get_format(workbook, "column_headers"))
            worksheet.write(2, 2, "Nickname", prepare_spreadsheet_get_format(workbook, "column_headers"))
            worksheet.write(2, 3, "URL for Details", prepare_spreadsheet_get_format(workbook, "column_headers"))
            worksheet.write(2, 4, "Area", prepare_spreadsheet_get_format(workbook, "column_headers"))
            worksheet.write(2, 5, "Mythological Association", prepare_spreadsheet_get_format(workbook, "column_headers"))
            worksheet.write(2, 6, "First Appearance", prepare_spreadsheet_get_format(workbook, "column_headers"))
            worksheet.write(2, 7, "Brightest Star", prepare_spreadsheet_get_format(workbook, "column_headers"))

        elif name == "constellation_data":
            # Capture optional arguments:
            dict_name = kwargs.get("dict_name", None)
            key = kwargs.get("key", None)
            i = kwargs.get("i", None)

            # Add/format main contents:
            worksheet.write(i, 0, key, prepare_spreadsheet_get_format(workbook, "data"))
            worksheet.write(i, 1, dict_name[key]["abbreviation"], prepare_spreadsheet_get_format(workbook, "data"))
            worksheet.write(i, 2, dict_name[key]["nickname"], prepare_spreadsheet_get_format(workbook, "data"))
            worksheet.write_url(i, 3, dict_name[key]["url"])
            worksheet.write(i, 4, dict_name[key]["area"], prepare_spreadsheet_get_format(workbook, "data"))
            worksheet.write(i, 5, dict_name[key]["myth_assoc"], prepare_spreadsheet_get_format(workbook, "data"))
            worksheet.write(i, 6, dict_name[key]["first_appear"], prepare_spreadsheet_get_format(workbook, "data"))
            if key == "Serpens":
                worksheet.write(i, 7,f"{dict_name[key]["brightest_star_name"]}\n{dict_name[key]["brightest_star_url"]}",prepare_spreadsheet_get_format(workbook, "data"))

            else:
                worksheet.write_url(i, 7, dict_name[key]["brightest_star_url"], string=f"{dict_name[key]["brightest_star_name"]}")

        elif name == "photo_details_headers":
            worksheet.write(2, 0, "Rover Name", prepare_spreadsheet_get_format(workbook, "column_headers"))
            worksheet.write(2, 1, "Earth Date", prepare_spreadsheet_get_format(workbook, "column_headers"))
            worksheet.write(2, 2, "SOL", prepare_spreadsheet_get_format(workbook, "column_headers"))
            worksheet.write(2, 3, "Pic ID", prepare_spreadsheet_get_format(workbook, "column_headers"))
            worksheet.write(2, 4, "Camera Name", prepare_spreadsheet_get_format(workbook, "column_headers"))
            worksheet.write(2, 5, "Camera Full Name", prepare_spreadsheet_get_format(workbook, "column_headers"))
            worksheet.write(2, 6, "URL", prepare_spreadsheet_get_format(workbook, "column_headers"))

        elif name == "photo_details_data":
            # Capture optional arguments:
            list_name = kwargs.get("list_name", None)
            worksheet_details = kwargs.get("worksheet_details", None)

            # Add/format main contents:
            i = 3
            dlg = PBI.PyBusyInfo(
                f"Photos from Mars: Exporting results to spreadsheet file 'Mars Photos - Details - {worksheet_details[0]}.xlsx': Processing...",
                title="Administrative Update")

            for j in range(worksheet_details[4], worksheet_details[5]):
                worksheet.write(i, 0, list_name[j].rover_name, prepare_spreadsheet_get_format(workbook, "data"))
                worksheet.write(i, 1, list_name[j].earth_date, prepare_spreadsheet_get_format(workbook, "data"))
                worksheet.write(i, 2, str(list_name[j].sol), prepare_spreadsheet_get_format(workbook, "data"))
                worksheet.write(i, 3, str(list_name[j].pic_id), prepare_spreadsheet_get_format(workbook, "data"))
                worksheet.write(i, 4, list_name[j].camera_name, prepare_spreadsheet_get_format(workbook, "data"))
                worksheet.write(i, 5, list_name[j].camera_full_name, prepare_spreadsheet_get_format(workbook, "data"))
                worksheet.write_url(i, 6, list_name[j].url, prepare_spreadsheet_get_format(workbook, "url"), tip="Click here for photo.")
                i += 1
            dlg = PBI.PyBusyInfo(
                f"Photos from Mars: Exporting results to spreadsheet file 'Mars Photos - Details - {worksheet_details[0]}.xlsx': Completed...",
                title="Administrative Update")

        elif name == "photos_available_headers":
            worksheet.write(2, 0, "Rover Name", prepare_spreadsheet_get_format(workbook, "column_headers"))
            worksheet.write(2, 1, "Earth Date", prepare_spreadsheet_get_format(workbook, "column_headers"))
            worksheet.write(2, 2, "SOL", prepare_spreadsheet_get_format(workbook, "column_headers"))
            worksheet.write(2, 3, "Cameras", prepare_spreadsheet_get_format(workbook, "column_headers"))
            worksheet.write(2, 4, "Total Photos Available", prepare_spreadsheet_get_format(workbook, "column_headers"))

        elif name == "photos_available_data":
            dlg = PBI.PyBusyInfo(
                f"Photos from Mars: Exporting results to spreadsheet file 'Mars Photos - Summary.xlsx': Processing...",
                title="Administrative Update")

            # Capture optional arguments:
            list_name = kwargs.get("list_name", None)

            # Add/format main contents:
            i = 3
            for item in list_name:
                worksheet.write(i, 0, item.rover_name, prepare_spreadsheet_get_format(workbook, "data"))
                worksheet.write(i, 1, item.earth_date, prepare_spreadsheet_get_format(workbook, "data"))
                worksheet.write(i, 2, str(item.sol), prepare_spreadsheet_get_format(workbook, "data"))
                worksheet.write(i, 3, item.cameras, prepare_spreadsheet_get_format(workbook, "data"))
                worksheet.write(i, 4, item.total_photos, prepare_spreadsheet_get_format(workbook, "data"))
                i += 1

            dlg = PBI.PyBusyInfo(
                f"Photos from Mars: Exporting results to spreadsheet file 'Mars Photos - Summary.xlsx': Completed...",
                title="Administrative Update")

        # At this point, function is presumed to have executed succssfully.  Return successful-execution indication
        # to the calling function:
        return True

    except:  # An error has occurred.
        update_system_log("prepare_spreadsheet_main_contents (" + name + ")", traceback.format_exc())

        # Return failed-execution indication to the calling function:
        return False


def prepare_spreadsheet_supplemental_formatting(workbook, worksheet, name, current_date_time, dict_name, num_columns_minus_one, column_widths, **kwargs):
    try:
        # Add an auto-filter:
        worksheet.autofilter(2, 0, len(dict_name) + 2, num_columns_minus_one)

        # Auto-fit the worksheet:
        worksheet.autofit()

        # Set column widths as needed:
        for i in range(0, len(column_widths)):
            worksheet.set_column(i, i, column_widths[i])

        # Set the footer:
        worksheet.set_footer(f"{recognition[name]}\n\n&CFile Name: &F\n&CPage &P of &N")

        # Add and format the spreadsheet header row, and implement the following: footer, page orientation, and margins:
        if name == "approaching_asteroids":
            # Add and format the spreadsheet header row:
            worksheet.merge_range("A1:L1",f"APPROACHING ASTEROIDS DATA (as of {current_date_time}) ({'{:,}'.format(len(dict_name))} Asteroids)", prepare_spreadsheet_get_format(workbook, "spreadsheet_header"))

            # Set page orientation:
            worksheet.set_landscape()

            # Set the margins:
            worksheet.set_margins(0.5, 0.5, 1, 1)  # Left, right, top, bottom

        if name == "confirmed_planets":
            # Add and format the spreadsheet header row:
            worksheet.merge_range("A1:I1",f"CONFIRMED PLANETS DATA (as of {current_date_time}) ({'{:,}'.format(len(dict_name))} Confirmed Planets)", prepare_spreadsheet_get_format(workbook, "spreadsheet_header"))

            # Set page orientation:
            worksheet.set_landscape()

            # Set the margins:
            worksheet.set_margins(0.5, 0.5, 1, 1)  # Left, right, top, bottom

        elif name == "constellations":
            # Add and format the spreadsheet header row:
            worksheet.merge_range("A1:H1",f"CONSTELLATION DATA (as of {current_date_time}) ({'{:,}'.format(len(dict_name))} Constellations)", prepare_spreadsheet_get_format(workbook, "spreadsheet_header"))

            # Set page orientation:
            worksheet.set_landscape()

            # Set the margins:
            worksheet.set_margins(0.5, 0.5, 1, 1)  # Left, right, top, bottom

        elif name == "photo_details":
            # Capture optional arguments:
            rover_name = kwargs.get("rover_name", None)
            earth_year = kwargs.get("earth_year", None)
            rover_earth_year_combo = kwargs.get("rover_earth_year_combo", None)
            rover_number_of_sheets_needed = kwargs.get("rover_number_of_sheets_needed", None)

            # Determine if rover/earth year combo needs multiple sheets:
            part_number = str(rover_earth_year_combo).split("_Part")
            if len(part_number) == 1:
                part_number = ""
            else:
                part_number = f", Part {part_number[len(part_number)-1]} of {rover_number_of_sheets_needed}"

            # Add and format the spreadsheet header row:
            worksheet.merge_range("A1:G1",f"PHOTOS TAKEN BY MARS ROVER '{str(rover_name).upper()}' - Year {str(earth_year)}{part_number} (as of {current_date_time})", prepare_spreadsheet_get_format(workbook, "spreadsheet_header"))

            # Set page orientation:
            worksheet.set_landscape()

            # Set the margins:
            worksheet.set_margins(1, 0.5, 1, 1)  # Left, right, top, bottom

        elif name == "photos_available":
            # Add and format the spreadsheet header row:
            worksheet.merge_range("A1:E1",f"SUMMARY OF PHOTOS TAKEN BY MARS ROVERS (as of {current_date_time})", prepare_spreadsheet_get_format(workbook, "spreadsheet_header"))

            # Set page orientation:
            worksheet.set_portrait()

            # Set the margins:
            worksheet.set_margins(1, 0.5, 1, 1)  # Left, right, top, bottom

        # Freeze panes (for top row and left column):
        worksheet.freeze_panes(3, 1)

        # Identify the rows to print at top of each page:
        worksheet.repeat_rows(0, 2)  # First row, last row

        # Scale the pages to fit within the page boundaries:
        worksheet.fit_to_pages(1, 0)

        # At this point, function is presumed to have executed successfully. Return successful-execution indication
        # to the calling function:
        return True

    except:  # An error has occurred.
        update_system_log("prepare_spreadsheet_supplemental_formatting (" + name + ")", traceback.format_exc())

        # Return failed-execution indication to the calling function:
        return False


def retrieve_from_database(trans_type, **kwargs):
    """Function to retrieve data from this application's database based on the type of transaction"""
    try:
        with app.app_context():
            if trans_type == "approaching_asteroids":
                # Retrieve and return all existing records, sorted by close-approach date, from the "approaching_asteroids" database table:
                return db.session.execute(db.select(ApproachingAsteroids).order_by(ApproachingAsteroids.close_approach_date, ApproachingAsteroids.name)).scalars().all()

            elif trans_type == "approaching_asteroids_by_close_approach_date":
                # Capture optional argument:
                close_approach_date = kwargs.get("close_approach_date", None)

                # Retrieve and return all existing records, sorted by asteroid's name, from the "approaching_asteroids" database table where the "close_approach_date" field matches the passed parameter:
                return db.session.execute(db.select(ApproachingAsteroids).where(ApproachingAsteroids.close_approach_date == close_approach_date).order_by(ApproachingAsteroids.name)).scalars().all()

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
        app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///space.db"

        # Initialize an instance of Bootstrap5, using the "app" object defined above as a parameter:
        Bootstrap5(app)

        # Retrieve the secret key to be used for CSRF protection:
        app.secret_key = os.getenv("SECRET_KEY_FOR_CSRF_PROTECTION")

        # Configure database tables.  If function failed, update system log and return
        # failed-execution indication to the calling function::
        if not config_database():
            update_system_log("run_app", "Error: Database configuration failed.")
            return False

        # Configure web forms.  If function failed, update system log and return
        # failed-execution indication to the calling function::
        if not config_web_forms():
            update_system_log("run_app", "Error: Web forms configuration failed.")
            return False

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
            if trans_type == "update_approaching_asteroids":
                # Delete all records from the "approaching_asteroids" database table:
                db.session.execute(db.delete(ApproachingAsteroids))
                db.session.commit()

                # Upload, to the "approaching_asteroids" database table, all contents of the "item_to_process" parameter:
                new_records = []
                for i in range(0, len(item_to_process)):
                    new_record = ApproachingAsteroids(
                        id=item_to_process[i]["id"],
                        name=item_to_process[i]["name"],
                        absolute_magnitude_h=item_to_process[i]["absolute_magnitude_h"],
                        estimated_diameter_km_min=item_to_process[i]["estimated_diameter_km_min"],
                        estimated_diameter_km_max=item_to_process[i]["estimated_diameter_km_max"],
                        is_potentially_hazardous=item_to_process[i]["is_potentially_hazardous"],
                        close_approach_date=item_to_process[i]["close_approach_date"],
                        relative_velocity_km_per_s=item_to_process[i]["relative_velocity_km_per_s"],
                        miss_distance_km=item_to_process[i]["miss_distance_km"],
                        orbiting_body=item_to_process[i]["orbiting_body"],
                        is_sentry_object=item_to_process[i]["is_sentry_object"],
                        url=item_to_process[i]["url"]
                    )

                    new_records.append(new_record)

                db.session.add_all(new_records)
                db.session.commit()

            elif trans_type == "update_confirmed_planets":
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


# Run main function for this application:
run_app()

# Destroy the object that was created to show user dialog and message boxes:
dlg.Destroy()

if __name__ == "__main__":
    app.run(debug=True, port=5003)