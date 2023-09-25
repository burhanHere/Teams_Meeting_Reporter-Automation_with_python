import csv
import json
import os
import re

import pygame
import win32com.client


def CheckFile():
    # we can add arguments in this to ask the user where he/show downloads the teams meeting report
    download_folder = os.path.join(os.path.expanduser(
        '~'), 'Downloads')  # getting download folder
    # iterating through file to get the meeting report
    for file in os.listdir(download_folder):
        if ".csv" in file and "Attendance report" in file and "(PROCESSED)" not in file:
            return str(download_folder + '\\' + file)
    return None


def RenameFile(file_path):
    new_name = str(file_path).replace(".csv", "(PROCESSED).csv")
    try:
        os.rename(file_path, new_name)
    except FileExistsError:
        os.remove(new_name)
        os.rename(file_path, new_name)


def GetParticipantEmails(file_path):
    participants_email = set()
    with open(file_path, 'r', encoding='utf-16') as file:
        my_reader = csv.reader(file, delimiter=",")
        for row in my_reader:
            for cell in row:
                temp_list = re.findall(
                    r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,7}\b', cell)
                for i in temp_list:
                    participants_email.add(i)
    return participants_email


def sendEmail(to, sent_on_behalf_of_name, attachment):
    # Create an instance of the Outlook application
    outlook = win32com.client.Dispatch("Outlook.Application")
    # Create a new mail message
    message = outlook.CreateItem(0)
    # Set the subject of the message
    message.Subject = "This is a test message."
    # Set the body of the message
    message.Body = "This is the body of the message."
    # Set the recipient of the message
    message.To = to
    # Set the "From" email address (if needed)
    # Uncomment and modify if necessary
    message.SentOnBehalfOfName = sent_on_behalf_of_name
    # Attach a file (if needed)
    message.Attachments.Add(attachment)
    # Send the message
    message.Send()


def read_email_from_json(file_path):
    try:
        with open(file_path, 'r') as json_file:
            data = json.load(json_file)
            if 'form' in data:
                return data['form']
            else:
                return None  # Return None if the 'form' key is missing
    except FileNotFoundError:
        print(f"File '{file_path}' not found.")
        return None
    except json.JSONDecodeError:
        print(f"Invalid JSON format in file '{file_path}'.")
        return None


def mainFunctionality(meeting_report_file_path):
    print(meeting_report_file_path)  # remove later
    # retrieving participant emails(Both emails one there main account and the one from which they joined the meeting)
    participant_emails = GetParticipantEmails(
        file_path=meeting_report_file_path)
    print(f"{len(participant_emails)}")  # remove later
    print(participant_emails)  # remove later
    print("______________________________")
    print("______________________________")
    # sending email to participant
    sent_on_behalf_of_name = read_email_from_json(
        file_path="./src/sent_on_behalf_of_name.json")
    print(f"sent_on_behalf_of_name: {sent_on_behalf_of_name}")
    for i in participant_emails:
        sendEmail(to=i, sent_on_behalf_of_name=sent_on_behalf_of_name,
                  attachment=meeting_report_file_path)
        print(f"Email sent to {i}")
        print("______________________________")
    # play sound all email sent
    pygame.init()
    pygame.mixer.init()
    sound = pygame.mixer.Sound(".\src\message-incoming-132126.wav")
    sound.play()
    pygame.time.delay(1000)  # Play for 1 seconds (adjust as needed)
    pygame.quit()
    # rename file after processing
    RenameFile(file_path=meeting_report_file_path)
    print("Renamed file for later personal use if needed.")
    # delete the used file is also an option
