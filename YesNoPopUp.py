import tkinter as tk
import pygame
from HelperFunctions import mainFunctionality, RenameFile


def tkinter_main(file_path):
    root = tk.Tk()
    # area on the screen in px which the app will consume
    root.geometry("675x200")
    root.title("Teams Meeting Reporting")  # app title
    root.configure(bg="#66B2FF")  # set app background
    icon_image = tk.PhotoImage(file="./src/icon.png")  # set logo of app
    root.tk_setPalette(background="#007BFF")
    root.iconphoto(True, icon_image)

    def on_yes_click():
        print("Yes clicked")
        my_file_path = file_path_lable.cget("text")
        root.destroy()
        mainFunctionality(meeting_report_file_path=my_file_path)

    def on_no_click():
        print("No clicked")
        my_file_path = file_path_lable.cget("text")
        root.destroy()
        RenameFile(file_path=my_file_path)

    info_lable = tk.Label(
        root, text="Do you want to send meeting report to the participants of the meeting file below?")
    info_lable.configure(bg="#66B2FF", fg="#333333", font=("Helvetica", 12))

    file_path_lable = tk.Label(root, text=file_path)
    file_path_lable.configure(bg="#66B2FF", fg="white",
                              font=("Helvetica", 12), wraplength=600, justify="left")

    note_lable = tk.Label(root, text="NOTE: Make sure that the meeting report file(.csv) is not in use of any other "
                                     "application else the mail will not be sent to any one.")
    note_lable.configure(bg="#FF8C42", fg="#333333", font=("Helvetica", 8))
    button_yes = tk.Button(
        root, text="Yes", command=on_yes_click)
    button_yes.configure(height=1, width=10, borderwidth=3, font=(
        "Helvetica", 12), bg='#EAEAEA', fg="#333333")
    button_no = tk.Button(
        root, text="No", command=on_no_click)
    button_no.configure(height=1, width=10, borderwidth=3, font=(
        "Helvetica", 12), bg='#EAEAEA', fg="#333333")

    info_lable.place(x=25, y=20)
    file_path_lable.place(x=25, y=50)
    note_lable.place(x=25, y=120)
    button_yes.place(x=430, y=150)
    button_no.place(x=540, y=150)

    pygame.init()
    pygame.mixer.init()
    sound = pygame.mixer.Sound("./src/off-hook-tone-43891-Trimed.wav")
    sound.play()
    pygame.time.delay(1000)  # Play for 1 seconds (adjust as needed)
    sound.play()
    pygame.quit()
    root.mainloop()
