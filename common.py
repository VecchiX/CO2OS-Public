# common.py

import time
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd

class Common:
    def __init__(self, username, password, date):
        self.username = username
        self.password = password
        self.date = date

    def login(self):
        self.driver = webdriver.Chrome()
        self.driver.get("https://example.com/login")
        self.driver.find_element(By.ID, "username").send_keys(self.username)
        self.driver.find_element(By.ID, "password").send_keys(self.password)
        self.driver.find_element(By.ID, "login-button").click()
        time.sleep(2)

    def get_data(self):
        date_field = self.driver.find_element(By.ID, "date-field")
        date_field.clear()
        date_field.send_keys(self.date)
        date_field.send_keys(Keys.RETURN)
        time.sleep(2)

        table_element = self.driver.find_element(By.ID, "data-table")
        table_html = table_element.get_attribute("outerHTML")
        df = pd.read_html(table_html)[0]
        return df

    def analyze_data(self, df):
        abnormal_values = df[df['value'] > 100]
        return abnormal_values

    def send_email(self, receiver_email, subject, message):
        msg = MIMEMultipart()
        msg['From'] = self.username
        msg['To'] = receiver_email
        msg['Subject'] = subject
        msg.attach(MIMEText(message, 'plain'))

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(self.username, self.password)
        text = msg.as_string()
        server.sendmail(self.username, receiver_email, text)
        server.quit()

    def logout(self):
        self.driver.find_element(By.ID, "logout-button").click()
        self.driver.quit()

    def main(self):
        attempt = 0
        max_attempts = 2

        while attempt < max_attempts:
            try:
                self.login()
                df = self.get_data()
                abnormal_values = self.analyze_data(df)

                if not abnormal_values.empty:
                    receiver_email = "recipient_email@example.com"
                    subject = "Abnormal Data Detected"
                    message = "Abnormal data detected:\n\n" + str(abnormal_values)
                    self.send_email(receiver_email, subject, message)
                else:
                    print("No abnormal data detected.")

                print("Process completed successfully.")
                break

            except Exception as e:
                print("An error occurred:", e)
                attempt += 1
                if attempt < max_attempts:
                    print("Retrying...")
                else:
                    print("Max attempts reached. Exiting...")
                    receiver_email = "recipient_email@example.com"
                    subject = "Error Occurred"
                    message = "An error occurred: " + str(e)
                    self.send_email(receiver_email, subject, message)
                    break

            finally:
                self.logout()