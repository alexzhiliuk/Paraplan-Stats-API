import requests
import json
from typing import Literal
from datetime import date, timedelta
import calendar
import openpyxl
import sys

from exceptions import AuthError, CsrfTokenError


class ParaplanAPI:
    BASE_URL = "https://paraplancrm.ru/"
    USER_URL = BASE_URL + "api/open/user"
    LOGIN_URL = BASE_URL + "api/public/login"
    STUDENTS_URL = BASE_URL + "api/open/students/min-info"
    STUDENT_SUBSCRIPTIONS_URL_TEMPLATE = BASE_URL + "/api/open/students/{student_id}/subscriptions/paginated?page=1&size=10"
    STUDENT_CARD_URL_TEMPLATE = "https://paraplancrm.ru/crm/#/students/{student_id}/groups"

    LOGIN_DATA = json.dumps({
        "username": "mdcode.vadim@gmail.com",
        "password": "123123123aa",
        "locale": "RU",
        "loginType": "KIDS_APP",
        "rememberMe": False,
        "captcha": ""
    })
    STUDENTS_DATA = json.dumps({
        "currentOnly": False
    })

    HEADERS = {
        'Content-Type': 'application/json',
    }

    def __init__(self):
        self.session = requests.Session()
        self.session.request("POST", self.LOGIN_URL, headers=self.HEADERS, data=self.LOGIN_DATA)

        if self.session.get(self.USER_URL).status_code != 200:
            raise AuthError("Не удалось залогиниться")

        csrf_token = self._get_csrf_token()
        if not csrf_token:
            raise CsrfTokenError("Не удалось получить CSRF токен")

        self._update_csrf_in_headers(csrf_token)

        self.current_month_period = self._get_month_period("current")
        self.previous_month_period = self._get_month_period("previous")
        self.next_month_period = self._get_month_period("next")
        self.current_week_period = self._get_current_week_period()
        self.after_current_week_period = self._get_period_after_current_week()

    def _get_csrf_token(self):
        return self.session.cookies.get("XSRF-TOKEN")

    def _update_csrf_in_headers(self, csrf_token):
        self.HEADERS["X-XSRF-TOKEN"] = csrf_token

    @staticmethod
    def _get_month_period(period: Literal["current", "previous", "next"]) -> tuple[date, date]:

        today = date.today()
        period_end_date = today - timedelta(days=today.day)

        if period == "current":
            pass
        elif period == "previous":
            period_end_date = period_end_date - timedelta(days=period_end_date.day)
        elif period == "next":
            period_start_date = period_end_date + timedelta(days=1)
            period_last_day_num = calendar.monthrange(period_start_date.year, period_start_date.month)[1]
            period_end_date = date(period_start_date.year, period_start_date.month, period_last_day_num)

        return period_end_date.replace(day=1), period_end_date

    @staticmethod
    def _get_current_week_period() -> tuple[date, date]:
        today = date.today()
        week_start_date = today - timedelta(days=today.weekday())
        week_end_date = week_start_date + timedelta(days=6)

        return week_start_date, week_end_date

    @staticmethod
    def _get_period_after_current_week() -> tuple[date, None]:
        today = date.today()
        week_start_date = today - timedelta(days=today.weekday())
        next_week_start_date = week_start_date + timedelta(days=7)
        return next_week_start_date, None

    @staticmethod
    def _format_subs_end_date(end_date: dict):
        months = {1: "Января", 2: "Февраля", 3: "Марта", 4: "Апреля", 5: "Мая", 6: "Июня",
                  7: "Июля", 8: "Августа", 9: "Сентября", 10: "Октября", 11: "Ноября", 12: "Декабря"}
        return f"{end_date['day']} {months[end_date['month']]} {end_date['year']}"

    @staticmethod
    def _convert_subs_end_date_to_date(end_date: dict):
        return date(end_date["year"], end_date["month"], end_date["day"])

    @staticmethod
    def _get_start_period_parameters(start_date: date):
        return f"&from.day={start_date.day}&from.month={start_date.month}&from.year={start_date.year}"

    @staticmethod
    def _get_end_period_parameters(end_date: date):
        return f"&to.day={end_date.day}&to.month={end_date.month}&to.year={end_date.year}"

    def _get_students(self) -> list:
        response = self.session.post(self.STUDENTS_URL, headers=self.HEADERS, data=self.STUDENTS_DATA)
        return response.json()["studentList"]

    def _get_student_subscriptions(self, student_id: str, period: tuple[date | None, date | None] = None,
                                   filter_by_end_date: bool = False) -> list:
        url = self.STUDENT_SUBSCRIPTIONS_URL_TEMPLATE.format(student_id=student_id)
        if period:
            if period[0]:
                url += self._get_start_period_parameters(period[0])
            if period[1]:
                url += self._get_end_period_parameters(period[1])
        subscriptions = self.session.get(url).json().get("itemList")
        subscriptions = list(filter(lambda item: item["lessonQuantity"] > 1 and item["endDate"], subscriptions))

        if filter_by_end_date and period:
            subscriptions = self._filter_subscriptions_by_end_date(subscriptions, period)

        return subscriptions

    def _filter_subscriptions_by_end_date(self, subscriptions: list, period: tuple[date | None, date | None]) -> list:
        if period[0] and period[1]:
            return list(filter(
                lambda item: period[0] < self._convert_subs_end_date_to_date(item["endDate"]) < period[1],
                subscriptions
            ))
        if period[0]:
            return list(filter(
                lambda item: period[0] < self._convert_subs_end_date_to_date(item["endDate"]),
                subscriptions
            ))
        if period[1]:
            return list(filter(
                lambda item: self._convert_subs_end_date_to_date(item["endDate"]) < period[1],
                subscriptions
            ))

    def _is_student_has_non_renewed_subs_in_month(self, student) -> tuple[bool, str | None]:
        previous_period_subs = self._get_student_subscriptions(student["id"], self.previous_month_period)
        if not previous_period_subs:
            return False, None

        current_period_subs = self._get_student_subscriptions(student["id"], self.current_month_period)
        if current_period_subs:
            return False, None

        return True, self._format_subs_end_date(previous_period_subs[0]["endDate"])

    def get_students_with_non_renewed_subscription_in_month(self) -> list:
        students_with_non_renewed_subscription = []
        students = self._get_students()

        for student in students:
            _, subs_end_date = self._is_student_has_non_renewed_subs_in_month(student)
            if not subs_end_date:
                continue

            students_with_non_renewed_subscription.append(
                {
                    "name": student["name"],
                    "link": self.STUDENT_CARD_URL_TEMPLATE.format(student_id=student["id"]),
                    "subs_end_date": subs_end_date
                }
            )
            print(f"Student {student['id']} processed")

        return students_with_non_renewed_subscription

    def get_students_week_subscriptions_info(self) -> dict:
        students_ids_who_have_non_renewed_subscription = []
        students_ids_who_renewed_subscription = []
        students = self._get_students()

        for student in students:

            if self._get_student_subscriptions(student["id"], self.current_week_period, filter_by_end_date=True):

                if self._get_student_subscriptions(student["id"], self.after_current_week_period, filter_by_end_date=True):

                    students_ids_who_renewed_subscription.append(
                        self.STUDENT_CARD_URL_TEMPLATE.format(student_id=student["id"])
                    )

                else:

                    students_ids_who_have_non_renewed_subscription.append(
                        self.STUDENT_CARD_URL_TEMPLATE.format(student_id=student["id"])
                    )

                print(f"Student {student['id']} processed")

        return {
            "have_non_renewed_subscription": students_ids_who_have_non_renewed_subscription,
            "who_renewed_subscription": students_ids_who_renewed_subscription
        }

    def get_students_with_ending_subscription_in_next_month(self) -> list:
        students_with_ending_subscription_in_next_month = []
        students = self._get_students()

        for student in students:
            subscriptions_ending_in_next_month = self._get_student_subscriptions(
                student["id"],
                self._get_month_period("next"),
                filter_by_end_date=True
            )

            if subscriptions_ending_in_next_month:

                total_price = sum([sub["totalPrice"] for sub in subscriptions_ending_in_next_month])

                students_with_ending_subscription_in_next_month.append({
                    "link": self.STUDENT_CARD_URL_TEMPLATE.format(student_id=student["id"]),
                    "subs_end_date": self._format_subs_end_date(subscriptions_ending_in_next_month[0]["endDate"]),
                    "total_price": total_price
                })

                print(f"Student {student['id']} processed")

        return students_with_ending_subscription_in_next_month

    def create_excel_file_students_with_non_renewed_subscription_in_month(self) -> None:

        students = self.get_students_with_non_renewed_subscription_in_month()

        wb = openpyxl.Workbook()
        ws = wb.worksheets[0]

        ws["A1"] = "Имя ученика"
        ws["B1"] = "Дата окончания абонемента"
        ws["C1"] = "Ссылка на карточку ученика"

        for row_index, student in enumerate(students, start=2):
            ws[f"A{row_index}"] = student["name"]
            ws[f"B{row_index}"] = student["subs_end_date"]
            ws[f"C{row_index}"] = student["link"]

        wb.save(filename="students-month.xlsx")

    def create_excel_file_with_students_week_subscriptions_info(self) -> None:

        subs_info = self.get_students_week_subscriptions_info()

        wb = openpyxl.Workbook()
        ws = wb.worksheets[0]

        ws["A1"] = "Всего"
        ws["A2"] = f"{len(subs_info["have_non_renewed_subscription"]) + len(subs_info["who_renewed_subscription"])}"
        ws["B1"] = "Непродлившие"
        ws["B2"] = f"{len(subs_info["have_non_renewed_subscription"])}"
        ws["C1"] = "Продлившие"
        ws["C2"] = f"{len(subs_info["who_renewed_subscription"])}"

        for row_index, student_link in enumerate(subs_info["have_non_renewed_subscription"], start=3):
            ws[f"B{row_index}"] = student_link

        for row_index, student_link in enumerate(subs_info["who_renewed_subscription"], start=3):
            ws[f"C{row_index}"] = student_link

        wb.save(filename="students-week-info.xlsx")

    def create_excel_students_with_ending_subscription_in_next_month(self) -> None:

        students = self.get_students_with_ending_subscription_in_next_month()

        wb = openpyxl.Workbook()
        ws = wb.worksheets[0]

        ws["A1"] = "Сумма"
        ws["B1"] = "Дата окончания абонемента"
        ws["C1"] = "Ссылка на карточку ученика"

        for row_index, student in enumerate(students, start=2):
            ws[f"A{row_index}"] = student["total_price"]
            ws[f"B{row_index}"] = student["subs_end_date"]
            ws[f"C{row_index}"] = student["link"]

        wb.save(filename="students-predicts.xlsx")


def main():

    if len(sys.argv) < 2:
        print("Не указан тип действия")
        print("Используйте current-month | current-week | next-month")
        return

    if sys.argv[1] not in ["current-month", "current-week", "next-month"]:
        print("Используйте current-month | current-week | next-month")
        return

    paraplan = ParaplanAPI()

    if sys.argv[1] == "current-month":
        paraplan.create_excel_file_students_with_non_renewed_subscription_in_month()
    if sys.argv[1] == "current-week":
        paraplan.create_excel_file_with_students_week_subscriptions_info()
    if sys.argv[1] == "next-month":
        paraplan.create_excel_students_with_ending_subscription_in_next_month()


if __name__ == "__main__":
    try:
        main()
    except AuthError as err:
        print(err)
    except CsrfTokenError as err:
        print(err)
