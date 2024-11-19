import traceback
from main import share_quotes_with_distribution, update_system_log


def inspire_us():
    """Function which distributes inspirational quotes to subscribers (can execute on a periodic basis if "cron.py" is scheduled for execution via web host)"""
    try:
        # Select and share quotes with subscribers:
        result = share_quotes_with_distribution()
        if result != "Success":  # Sharing function failed.
            update_system_log("inspire_us", f"Failed execution ({result}).")
        else:  # Sharing function succeeded.
            update_system_log("inspire_us", "Successfully executed.")

    except:  # An error has occurred.
        update_system_log("inspire_us", traceback.format_exc())


inspire_us()