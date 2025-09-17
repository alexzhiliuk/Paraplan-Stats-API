from enum import Enum
from typing import TypedDict


class Statuses(TypedDict):
    ATTENDED_TRIAL: str
    ATTENDED_FREE_TRIAL: str
    WORKED_OUT: str
    SKIP: str
    ATTEND: str


class StatusesEnum(Enum):
    ATTENDED_TRIAL = "78e0eab0-b4d4-9cd2-3c9a-bc862db3bbbc"
    ATTENDED_FREE_TRIAL = "fea4db4a-b812-a27f-1d02-998fc23f76b3"
    WORKED_OUT = "57b3be44-8863-4a04-18e3-492314751701"
    SKIP = "0376516a-2bfc-dbbc-8fe7-9c35e7b18365"
    ATTEND = "a9ff5b2c-f5f9-cb83-a512-9ba807f74fd2"


class TeachersAttendancesStats:

    data = dict()

    @classmethod
    def get_stats(cls):
        result = cls.data.copy()
        cls.data = dict()
        return result

    @classmethod
    def _init_teacher(cls, teacher: str):
        if teacher in cls.data.keys():
            return

        cls.data[teacher] = {
            "statuses": {
                "ATTENDED_TRIAL": 0,
                "WORKED_OUT": 0,
                "SKIP": 0,
                "ATTEND": 0,
            },
            "attendances_count": 0
        }

    @classmethod
    def add_teacher_attendance_stats(cls, attendance):
        teacher = attendance["teacherList"][0]["teacherInfo"]["name"]
        cls._init_teacher(teacher)

        counting = False
        for attendee in attendance["attendeeList"]:
            status_id = attendee["statusId"]
            if status_id == StatusesEnum.ATTENDED_TRIAL.value:
                cls.data[teacher]["statuses"]["ATTENDED_TRIAL"] += 1
                counting = True
            elif status_id == StatusesEnum.WORKED_OUT.value:
                cls.data[teacher]["statuses"]["WORKED_OUT"] += 1
                counting = True
            elif status_id == StatusesEnum.ATTEND.value:
                cls.data[teacher]["statuses"]["ATTEND"] += 1
                counting = True
            elif status_id == StatusesEnum.SKIP.value:
                cls.data[teacher]["statuses"]["SKIP"] += 1

        if counting:
            cls.data[teacher]["attendances_count"] += 1
