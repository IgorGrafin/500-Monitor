# By Igor A. Grafin
# Import system modules
import win32com.client
import pywintypes
import time
# import pygame
# import os
import easygui


# Returns array of config. Server - Database - View
#  Eg. [['192.168.242.10', 'domlog.nsf', '500'], ['192.168.242.10', 'domlog2.nsf', '500']]
def get_config(conf):
    with open(conf) as config:
        array = [row.strip().split("|") for row in config]
    return array


# Append problem documents to the file
def write_log(file, text):
    f = open(file, 'a')
    f.write(text + '\n')


# Returns last unid of document which was found and has been written into file
def get_last_doc(file, probe):
    with open(file) as config:
        array = [row.strip() for row in config]
    # print('debug', array, probe)
    # TODO if the second row in last doc file is empty - it stops here
    if array:
        if len(array[probe]) == 32:
            return array[probe]
        else:
            return "Nothing"
    else:
        return "Nothing"


# Writes last document unids into file
# It opens the file, read it, and then rewrite the file with needed unid
def write_last_doc(file, probe, unid, probe_config):
    # print('start1')
    with open(file) as config:
        temp_array = [row.strip() for row in config]
    # print('temp_array = ', temp_array)
    if temp_array:
        # print('step 1.1')
        temp_array[probe] = unid
        # print('step 1.2, ', temp_array)
    else:
        # print('step 2.1')
        temp_array = ["Nothing"] * len(probe_config)
        temp_array[probe] = unid
        # print('step 2.2 ', temp_array)
    f = open(file, 'w')
    f.write('\n'.join(temp_array))
    f.close()


# Need to initialize file of last document unids, it calls write_last_doc function
def init_last_doc(probe, config):
    Server = probes[probe][0]
    DdPath = probes[probe][1]
    DdView = probes[probe][2]

    # Connect
    notesSession = win32com.client.Dispatch('Lotus.NotesSession')
    try:
        notesSession.Initialize(mail_password)
    except pywintypes.com_error:
        easygui.msgbox('Неверный пароль')
        raise Exception('Cannot access database using this password to ', DdPath, ' on server ', Server)

    try:
        notesDatabase = notesSession.GetDatabase(Server, DdPath)
    except pywintypes.com_error:
        print('Невозможно получить базу данных {}', DdPath)


    # print('init_last_doc: successfully connected to', Server)
    view = notesDatabase.GetView(DdView)
    if not view:
        raise Exception('View ', DdView, ' not found')
    # Get the first document
    view.refresh()
    document_unid = view.GetLastDocument().UniversalID
    write_last_doc(file_last_docs, probe, document_unid, config)


# def document_generator(view_name):
#     # Get view
#     view = notesDatabase.GetView(view_name)
#     if not view:
#         raise Exception('View ', view_name, ' not found')
#     # Get the first document
#     view.refresh()
#     document = view.GetLastDocument()
#     # If the document exists,
#     while document:
#         # Yield it
#         yield document
#         # Get the next document
#         document = view.GetPrevDocument(document)


def send_mail(to, attach):
    outlook = win32com.client.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'grafinia@veb.ru;zagreba@veb.ru'
    mail.Subject = 'Hello from Python'
    mail.Body = 'Look at these afwul errors:'
    mail.HTMLBody = '<h2>OMG! ERROR 500 AGAIN!</h2>'  # this field is optional
    # In case you want to attach a file to the email
    attachment = "D:\GitHub\PycharmProjects\COM-Applications\log.txt"
    mail.Attachments.Add(attachment)
    mail.Send()


def do_init():
    for probe in range(len(probes)):
        if get_last_doc(file_last_docs, probe) == "Nothing":  # if the file of last docs is empty (first run)
            init_last_doc(probe, probes)  # get last documents in the view and write their unids to the file


def do_scenario(probes, file_last_docs, file_log, mail_password, sleep_time):
    alarm = False
    for probe in range(len(probes)):  # "probe" here is a number of row in config file
        Server = probes[probe][0]
        DdPath = probes[probe][1]
        DdView = probes[probe][2]
        LastUnid = get_last_doc(file_last_docs, probe)
        # Connect
        notesSession = win32com.client.Dispatch('Lotus.NotesSession')
        try:
            notesSession.Initialize(mail_password)
        except pywintypes.com_error:
            easygui.msgbox('Неверный пароль')
            easygui.exceptionbox()
            raise Exception('Cannot access database using this password to ', DdPath, ' on server ', Server)

        try:
            notesDatabase = notesSession.GetDatabase(Server, DdPath)
        except pywintypes.com_error:
            # print('Невозможно получить базу данных {} на сервере {}'.format(DdPath, Server))
            raise 'Невозможно получить базу данных {} на сервере {}'.format(DdPath, Server)


        # Get a list of views
        if not notesDatabase.isOpen:
            print('2Невозможно получить базу данных {} на сервере {}'.format(DdPath, Server))
            break
        # print('successfully connected to', Server)
        view = notesDatabase.GetView(DdView)
        if not view:
            raise Exception('View ', DdView, ' not found')
        # Refresh view and get the last document
        view.refresh()
        document = view.GetLastDocument()
        if document.UniversalID != LastUnid:
            print('Step 1. probe = ', probe, probes[probe][0], 'New UNID: ', document.UniversalID, 'Old UNID: ',
                  LastUnid)
            temp_last_unid = document.UniversalID
            while document.UniversalID != LastUnid:
                log_text = str(document.Created) + " " + \
                           str(document.GetItemValue('AuthenticatedUser')[0]) + " " + \
                           str(document.GetItemValue('Request')[0]) + " " + \
                           str(document.GetItemValue('ServerAddress')[0] + " ")
                write_log(file_log, log_text)
                # print(log_text)
                document = view.GetPrevDocument(
                    document)  # FIXME: it stopped here once because there is no document.
            write_last_doc(file_last_docs, probe, temp_last_unid, probes)
            alarm = True
            # for doc in document_generator(DdView):
    if alarm is True:
        print('ALAAAAAAARM!!!!!!!111')

    return alarm
        # os.startfile(file_log)
        # pygame.mixer.init()
        # pygame.mixer.music.load("SilentHill.mp3")
        # pygame.mixer.music.play()
        # while pygame.mixer.music.get_busy():
        #     pygame.time.Clock().tick(1)
        #     pygame.time.wait(10000)
        #     pygame.mixer.music.stop()

    # time.sleep(sleep_time)

# Constants
# sleep_time = 5
# mailTo = 'grafinia@veb.ru;zagreba@veb.ru'
# mail_password = easygui.passwordbox('Введите пароль от Lotus Notes: ')  # TODO Just plain text. Need to make secure password input instead.
# file_last_docs = 'last_docs.txt'
# file_config = "config.conf"
# file_log = "log.txt"
# probes = get_config(file_config)  # List of the config file contains server, database, view
# print(probes)


# Initialization
#do_init()

#  do_scenario()
