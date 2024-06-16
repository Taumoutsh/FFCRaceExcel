import datetime


class Race:
    date: datetime
    place: str
    category: str
    race_name: str
    info_link: str
    cycling_club: str
    department: str

    def __init__(self):
        self.date = None
        self.department = None
        self.info_link = None

