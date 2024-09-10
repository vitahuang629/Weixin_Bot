# Attendance Notification Bot
This project integrates 飛騰 attendance information with the WeChat Work (企業微信) API to build a bot that automatically sends notifications about late arrivals. The bot is designed to streamline attendance management by instantly informing employees or managers about tardiness.

## Features:
Automated Notifications: Sends automated messages to employees or supervisors when someone is marked as late in the HR system.
Integration: Connects directly to the HR system to retrieve real-time attendance and tardiness information.
WeChat Work API Integration: Utilizes the WeChat Work API to send notifications directly within the company's internal communication platform.
Customizable Message Templates: Allows custom message formatting to fit the needs of the company, such as including details like time of late arrival and employee ID.

## Technologies Used:
Python: Core language for integrating APIs and handling automation.
Feiteng API: To retrieve late arrival data from the attendance system.
WeChat Work API: To send notifications to the company's communication platform.
Requests: For making API calls.
