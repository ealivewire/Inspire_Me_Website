import os
import datetime
from dotenv import load_dotenv

# Load environmental variables from the ".env" file:
load_dotenv()

# Define constants to be used for e-mailing content:
SENDER_EMAIL_GMAIL = os.getenv("SENDER_EMAIL_GMAIL")
SENDER_PASSWORD_GMAIL = os.getenv("SENDER_PASSWORD_GMAIL") # App password (for the app "Python e-mail", NOT the normal password for the account).
SENDER_HOST = os.getenv("SENDER_HOST")
SENDER_PORT = str(os.getenv("SENDER_PORT"))

# Define constant for web page loading-time allowance (in seconds) for the web-scrapers:
WEB_LOADING_TIME_ALLOWANCE = 10

# Create a dictionary to store recognition merit by content type:
recognition = {
    "web_template":
        f"Website template created by the Bootstrap team · © {datetime.datetime.now().year}"
}

# Define variable to represent the Flask application object to be used for this application:
app = None

# Define variable to represent the database supporting this application:
db = None

# Initialize class variables for database tables:
Categories = None
InspirationDataSources = None
InspirationalQuotes = None
Subscribers = None
Users = None

# Initialize class variables for web forms:
AdminLoginForm = None
AdminUpdateForm = None
ContactForm = None
AddOrEditSubscriberForm = None