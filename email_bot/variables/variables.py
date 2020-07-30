'''
Variables for Robot Framework goes here.
'''
import calendar
from datetime import date

# +
from RPA.Robocloud.Secrets import Secrets

secrets = Secrets()
USER_NAME = secrets.get_secret("credentials")["username"]
PASSWORD = secrets.get_secret("credentials")["password"]
# -

WEEK_DAY_NAME = calendar.day_name[date.today().weekday()]
