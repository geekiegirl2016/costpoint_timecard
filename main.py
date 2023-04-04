from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


from datetime import date
from datetime import datetime
from calendar import monthrange
import smtplib



class TIMECARD_READER:
    def __init__(self, costpoint_user, costpoint_password, costpoint_url, costpoint_db, sender_email, email_password, receiver_email, smtp_server='smtp.gmail.com', smtp_port=587, autorun = True, default_row=0, default_hours=8):
        self.today = date.today()
        self.day = date.today().day
        self.month = date.today().month
        self.year = date.today().year
        self.isWeekDay = self.get_ifweekday(self.today)
        self.deltek_day, self.pay_period = self.get_current_deltek_day()
        self.default_row = default_row
        self.default_hours = default_hours
        self.sender_email = sender_email
        self.email_password = email_password
        self.reciever_email = receiver_email
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.costpoint_user = costpoint_user
        self.costpoint_password = costpoint_password
        self.costpoint_url = costpoint_url
        self.costpoint_db = costpoint_db
        if self.pay_period == 1:
            self.days_in_pay_period = 15
        else:
            self.days_in_pay_period = monthrange(self.year, self.month)[1] - 15

        if autorun:
            if self.isWeekDay:
                print("Let's do that timecard.")
                self.bring_up_timecard(audittc=True)
            else:
                print("It is a weekend, chill....")
                print("But let's doublecheck that time card.")
                self.bring_up_timecard(audittc=True)

    def send_email(self, subject, body):
        smtp_server = self.smtp_server
        port = self.smtp_port
        sender_email = self.sender_email
        password = self.email_password
        receiver_email = self.reciever_email

        # Create a message
        subject = subject
        body = body
        message = f'Subject: {subject}\n\n{body}'

        # Send the message
        with smtplib.SMTP(smtp_server, port) as server:
            server.starttls()  # Enable TLS encryption
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_email, message)


    def get_current_deltek_day(self):
        if self.day > 15:
            deltek_day = self.day - 15
            pay_period = 2
        else:
            deltek_day = self.day
            pay_period = 1
        return deltek_day, pay_period

    def get_ifweekday(self, date):
        weekday = date.weekday()
        if weekday >= 5:
            return False
        else:
            return True

    def get_charge_code(self):
        pass

    def bring_up_timecard(self, audittc=False, log_today=False, hours = "8"):  
        '''
        The next 3 lines are used if you need to troubleshoot the script
        with this it won't close the browser so you can see what it does
        '''
        #options = webdriver.ChromeOptions()
        #options.add_experimental_option("detach", True)
        #driver = webdriver.Chrome(options=options)  # Replace with the appropriate driver for your browser
        driver = webdriver.Chrome()
        # navigate to the Deltek login page

        driver.get(self.costpoint_url)

        # add a wait time
        driver.implicitly_wait(3)

        # enter your login credentials
        username = driver.find_element(By.NAME, "USER")
        username.send_keys(self.costpoint_user)
        password = driver.find_element(By.NAME, "CLIENT_PASSWORD")
        password.send_keys(self.costpoint_password)
        databasefield = driver.find_element(By.NAME, "DATABASE")
        databasefield.send_keys(self.costpoint_db)
        driver.find_element(By.ID, "loginBtn").click()


        wait = WebDriverWait(driver, 10)
        wait.until(EC.title_contains('Costpoint'))

        driver.find_element(By.ID, "pushSubNo").click()
        wait = WebDriverWait(driver, 5)
        driver.find_element(By.ID, "goTo").click()
        driver.find_element(By.ID, "bus__TC").click()
        driver.find_element(By.ID, "dpt__TM").click()
        driver.find_element(By.ID, "wrk__Timesheets").click()
        driver.find_element(By.ID, "actvty__TMMTIMESHEET").click()
        wait = WebDriverWait(driver, 10)

        timecard = TIMECARD(self.days_in_pay_period, self.day, self.deltek_day, self.pay_period)


        if log_today:
            current_day = driver.find_element(By.ID, f"DAY{self.deltek_day}_HRS-_{self.default_row}_E")
            if current_day.get_attribute('value'):
                if float(current_day.get_attribute('value')) >= 8:
                    email_subject = "Timecard Good To Go!"
                    email_body = f"You timecard was already good for {self.today}."
                else:
                    email_subject = "Updated but timecard Good To Go!!"
                    email_body = f"I updated your timecard for today, {self.today}."
                    current_day.send_keys(self.default_hours)
            else:
                email_subject = "Updated but timecard Good To Go!!"
                email_body = f"I updated your timecard for today, {self.today}."
                current_day.send_keys(self.default_hours)
            wait = WebDriverWait(driver, 10)
            wait.until(EC.element_to_be_clickable(By.ID, "svCntBttn"))
            driver.find_element(By.ID, "svCntBttn").click()

        if audittc:
            row = 0
            while len(driver.find_elements(By.ID, f"LINE_DESC-_{row}_E")) > 0:
                print(row)
                project_description = driver.find_element(By.ID, f"LINE_DESC-_{row}_E").get_attribute('value')
                print("Project:", project_description)
                project_id = driver.find_element(By.ID, f"UDT02_ID-_{row}_E").get_attribute('value')
                print("ID:", project_id)
                weekhours = {}
                for xday in range(1,self.deltek_day+1):
                    weekend = False
                    if not len(driver.find_elements(By.ID, f"DAY{xday}_HRS-_{row}_E")) > 0:
                        continue
                    if self.pay_period == 1:
                        xdate = datetime.strptime(f"{self.year}-{self.month}-{xday}", "%Y-%m-%d")
                    else:
                        xdate = datetime.strptime(f"{self.year}-{self.month}-{xday+15}", "%Y-%m-%d")

                    if self.get_ifweekday(xdate):
                        weekend = False
                    else:
                        weekend = True
                    day_id = (f"DAY{xday}_HRS-_{row}_E")
                    if weekend:
                        hours = 0
                        weekhours[xday] = "0"
                    else:
                        hours = driver.find_element(By.ID, day_id).get_attribute('value')
                    weekhours[xday] = TIMECARD_DAY(hours, weekend, day_id)

                ROW = TIMECARD_ROW(project_description, project_id, weekhours)
                timecard.add_row(row=ROW)
                row = row + 1
            missing_days = timecard.find_missing_days()
            missing_day_print = []
            if self.pay_period == 1:
                missing_day_print = missing_days
            else:
                for day in missing_days:
                    missing_day_print.append(day + 15)
            for missing_day in missing_days:
                day = driver.find_element(By.ID, f"DAY{missing_day}_HRS-_{self.default_row}_E")
                day.send_keys(self.default_hours)
            if missing_days:
                email_subject = "Updated but timecard Good To Go!"
                email_body = f"You were missing a few days, but I fixed it! {missing_day_print}."
            else:
                email_subject = "Timecard Good To Go!"
                email_body = f"Your timecard was all good! Kickass!"

        driver.find_element(By.ID, "svCntBttn").click()
        myemail = self.send_email(email_subject, email_body)

class TIMECARD_DAY:
    def __init__(self, hours, weekend, id):
        self.hours = hours
        self.weekend = weekend
        self.id = id

class TIMECARD_ROW:
    def __init__(self, name, id, hours):
        self.name = name
        self.id = id
        self.hours = hours

    def get_hours_for_day(self, day):
        day_hours = self.hours.get(day, None)

        if day_hours:
            if not day_hours.hours:
                return 0
            return day_hours.hours
        else:
            return 0

class TIMECARD:
    def __init__(self,days_in_pay_period, current_day, deltek_day, payperiod):
        self.today = date.today()
        self.day = date.today().day
        self.month = date.today().month
        self.year = date.today().year
        self.days_in_pay_period = days_in_pay_period
        self.current_day = current_day
        self.deltek_day = deltek_day
        self.payperiod = payperiod
        self.rows = []

    def add_row(self, row):
        self.rows.append(row)

    def is_there_8_for_day(self, day):
        total = 0
        for row in self.rows:
            total = total + float((row.get_hours_for_day(day)))
        if total >= 8:
            return True
        else:
            return False

    def is_day_weekend(self, realday):
        day = datetime.strptime(f"{self.year}-{self.month}-{realday}", "%Y-%m-%d")
        weekday = day.weekday()
        if weekday >= 5:
            return True
        else:
            return False


    def find_missing_days(self):
        missing_days = []
        print("Getting Missing Days")
        for x in range(1 , self.deltek_day+1):
            if not self.is_there_8_for_day(x):
                if self.payperiod == 1:
                    realday = x
                else:
                    realday = x + 15
                if not self.is_day_weekend(realday):
                    print(f"Hours are missing from Day {realday}")
                    print(f"Filling in day")
                    missing_days.append(x)
        return missing_days

    def get_rows(self):
        for row in self.rows:
            print(row.get_hours_for_day(2))

