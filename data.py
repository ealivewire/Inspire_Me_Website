import os
import datetime
from dotenv import load_dotenv

# Load environmental variables from the ".env" file:
load_dotenv()

# Define a list for storing metadata re: each website to be scraped:
data_source = [
    {"name": "89 Moving On Quotes To Bring Lightness Into Your Life", "url": "https://wisdomquotes.com/moving-on-quotes/","recognition": f"Courtesy of Wisdom Quotes, Copyright © 2004-{datetime.datetime.now().year}", "count": 85, "category": 1},
    {"name": "83 Serenity Quotes To Bring Quietude Into Your Life", "url": "https://wisdomquotes.com/serenity-quotes/","recognition": f"Courtesy of Wisdom Quotes, Copyright © 2004-{datetime.datetime.now().year}", "count": 83, "category": 2},
    {"name": "75 Letting Go Quotes To Help You Live Free", "url": "https://wisdomquotes.com/letting-go-quotes/", "recognition": f"Courtesy of Wisdom Quotes, Copyright © 2004-{datetime.datetime.now().year}", "count": 76, "category": 1},
    {"name": "70 Quotes for Moving On and Letting Go to Take You To a Better Future", "url": "https://blog.gratefulness.me/moving-on-quotes/","recognition": f"© {datetime.datetime.now().year} - The Life Blog", "count": 75, "category": 1},
    {"name": "Letting Go Quotes: 89 Quotes about Letting Go and Moving On", "url": "https://www.developgoodhabits.com/letting-go-quotes/","recognition": f"© {datetime.datetime.now().year} Oldtown Publishing LLC", "count": 89, "category": 1}
]

# Define constants to be used for e-mailing messages submitted via the "Contact Us" web page:
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

# Define variable to represent the Flask application object to be used for this website:
app = None

# Define variable to represent the database supporting this website:
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
DisplayApproachingAsteroidsSheetForm = None
DisplayConfirmedPlanetsSheetForm = None
DisplayConstellationSheetForm = None
DisplayMarsPhotosSheetForm = None
ViewApproachingAsteroidsForm = None
ViewConfirmedPlanetsForm = None
ViewConstellationForm = None
ViewMarsPhotosForm = None