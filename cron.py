from datetime import datetime
from dotenv import load_dotenv
from flask import Flask
from flask_sqlalchemy import SQLAlchemy
import random
import os
import smtplib
from sqlalchemy import Integer, String, Boolean, ForeignKey
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, relationship
import traceback
from typing import List

# # Load environmental variables from the ".env" file:
# load_dotenv()

# Define constants to be used for e-mailing messages submitted via the "Contact Us" web page:
ADMIN_EMAIL = None
SENDER_EMAIL_GMAIL = None
SENDER_PASSWORD_GMAIL = None
SENDER_HOST = None
SENDER_PORT = None

# Initialize the Flask app. object:
app = Flask(__name__)

db = None

# Create needed class "Base":
class Base(DeclarativeBase):
    pass


# Initialize class variables for database tables:
Categories = None
InspirationDataSources = None
InspirationalQuotes = None
Subscribers = None
Users = None


def config_database():
    """Function for configuring the database tables supporting this script"""
    global db, app, Categories, InspirationDataSources, InspirationalQuotes, Subscribers

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


def email_quotes_to_distribution(message, subscribers):
    """Function to process a message that user wishes to e-mail from this website to the website administrator."""
    global SENDER_EMAIL_GMAIL, SENDER_HOST, SENDER_PASSWORD_GMAIL, SENDER_PORT

    try:
        # E-mail the message using the 'message' parameter:
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
                    to_addrs=subscribers,
                    # to_addrs="ealivewire@gmail.com",
                    # msg=f"Subject: Inspiration from the 'Inspire Me' website"

                    msg = f"Subject: Inspiration from the 'Inspire Me' website\n\n{message}"
                )
                # Return successful-execution message to the calling function::
                return "Your message has been successfully sent."

    except:  # An error has occurred.
        update_system_log("email_quotes_to_distribution", traceback.format_exc())

        # Return failed-execution message to the calling function:
        return "An error has occurred. Your message was not sent."


def retrieve_from_database(trans_type, **kwargs):
    """Function to retrieve data from this application's database based on the type of transaction"""
    global app, db, Categories, InspirationDataSources, InspirationalQuotes, Subscribers

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

            elif trans_type == "get_subscribers":
                return db.session.execute(db.select(Subscribers).order_by(Subscribers.name)).scalars().all()

    except:  # An error has occurred.
        update_system_log("retrieve_from_database (" + trans_type + ")", traceback.format_exc())

        # Return empty dictionary as a failed-execution indication to the calling function:
        return {}


def run_app():
    """Main function for this application"""
    global app, ADMIN_EMAIL, SENDER_EMAIL_GMAIL, SENDER_PASSWORD_GMAIL, SENDER_HOST, SENDER_PORT

    try:
        # # Load environmental variables from the ".env" file:
        load_dotenv()

        # Define constants to be used for e-mailing messages submitted via the "Contact Us" web page:
        ADMIN_EMAIL = os.getenv("ADMIN_EMAIL")
        SENDER_EMAIL_GMAIL = os.getenv("SENDER_EMAIL_GMAIL")
        SENDER_PASSWORD_GMAIL = os.getenv("SENDER_PASSWORD_GMAIL")  # App password (for the app "Python e-mail", NOT the normal password for the account).
        SENDER_HOST = os.getenv("SENDER_HOST")
        SENDER_PORT = int(os.getenv("SENDER_PORT"))

        # Configure the SQLite database, relative to the app instance folder:
        app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///inspiration.db"

        # Configure database tables.  If function failed, update system log and return
        # failed-execution indication to the calling function::
        if not config_database():
            update_system_log("run_app", "Error: Database configuration failed.")

        # Select and share quotes with subscribers:
        share_quotes_with_distribution()

    except:  # An error has occurred.
        update_system_log("run_app", traceback.format_exc())


def share_quotes_with_distribution():
    categories = retrieve_from_database("get_categories")
    if categories == {}:
        exit()
    elif categories == []:
        exit()

    message_intro = f"TEST:\n\n"
    message_main_contents = ""

    for i in range(0,len(categories)):
        quotes_in_category = retrieve_from_database("get_quotes_for_category", category_id = categories[i].id)
        if quotes_in_category == {}:
            exit()

        selected_quote = quotes_in_category[random.choice(range(0, len(quotes_in_category)))][0].quote

        message_main_contents += f"{categories[i].category.upper()}:\n{selected_quote.encode('ascii', 'ignore').decode('ascii')}\n\n"

    message_to_send = message_intro + message_main_contents

    subscribers_from_db = retrieve_from_database("get_subscribers")
    if subscribers_from_db == {}:
        exit()
    elif subscribers_from_db == []:
        exit()

    subscribers = [ADMIN_EMAIL]
    for i in range(0,len(subscribers_from_db)):
        # subscribers += subscribers_from_db[i].email + ","
        subscribers.append(subscribers_from_db[i].email)

    email_quotes_to_distribution(message=message_to_send, subscribers=subscribers)


def update_system_log(activity, log):
    """Function to update the system log, either to log errors encountered or log successful execution of milestone admin. updates"""
    # Capture current date/time:
    current_date_time = datetime.now()
    current_date_time_file = current_date_time.strftime("%Y-%m-%d")

    # Update log file.  If log file does not exist, create it:
    with open("log_inspire_me_" + current_date_time_file + ".txt", "a") as f:
        f.write(datetime.now().strftime("%Y-%m-%d @ %I:%M %p") + ":\n")
        f.write(activity + ": " + log + "\n")

    # Close the log file:
    f.close()


run_app()