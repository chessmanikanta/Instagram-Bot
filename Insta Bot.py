from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd
from openpyxl import load_workbook
import pyautogui
import win10toast
import smtplib
import gtts
import pyttsx3
import os


driver = webdriver.Chrome('C:\Installers\chromedriver.exe')
time.sleep(10)
driver.get("https://www.instagram.com/")
driver.maximize_window()


Done='Hey Mani Activity iiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiiis doneeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeee'
Error='Hey Mani Errorrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrr ocurred'
welcome_message='Hey Mani Good Morning mera dost!!!!!!!!!!!!'


def voice(message):
    engine = pyttsx3.init()
    print(engine)
    engine.say(message)
    engine.runAndWait()

def login(uname,pwd):

# username enter
    ubutton = driver.find_element_by_xpath("//input[@name='username']")
    time.sleep(1)
    ubutton.send_keys(uname)

    time.sleep(1)
#password enter
    pbutton = driver.find_element_by_xpath("//input[@name='password']")
    time.sleep(1)
    pbutton.send_keys(pwd)

#login button click

    login_button = driver.find_element_by_xpath("//Button[@type='submit']")
        #"/html/body/div[1]/section/main/article/div[2]/div[1]/div/form/div[4]/button/div")
    login_button.click()
    time.sleep(7)
    mouse_move()

def instagram_desktop_notification():
#
    notification1 = driver.find_element_by_xpath("/html/body/div[1]/section/main/div/div/div/div/button")
    time.sleep(2)
    notification1.click()
    time.sleep(5)

# Notification for desktop notification

    notification = driver.find_element_by_xpath("/html/body/div[4]/div/div/div/div[3]/button[2]")
    time.sleep(2)
    notification.click()
    time.sleep(5)
    mouse_move()

def mouse_move():
    #pyautogui.moveTo(100, 100, duration=1)
    #pyautogui.moveTo(10, 10, duration=1)
    pyautogui.press('volumedown')
    time.sleep(1)
    pyautogui.press('volumeup')
    time.sleep(5)
def desktop_notification(profile):

#Desktop Notification
    toaster = win10toast.ToastNotifier()
    toaster.show_toast('Instagram Bot' ,'likes and messages are sent to'+ profile,duration=5)


def getting_links_from_messagees(name):
#calling profile function
    going_to_profile(name)

#clicking message button
    message_button = driver.find_element_by_xpath("//Button[@type='button']").click()
    time.sleep(2)

    time.sleep(3)
#clicking the textarea in messeger
    txtarea = driver.find_element_by_xpath(
        "/html/body/div[1]/section/div/div[2]/div/div/div[2]/div[2]/div/div[2]/div/div/div[2]/textarea")
    time.sleep(2)

    a_tag_links=[]
    profile_links=[]
    elements = driver.find_element_by_class_name("uueGX")
    mouse_move()

#loop that brings links from the messages
    for i in range(1, 2):
        driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(2)
        a_tag_links = elements.find_elements_by_tag_name('a')
        print(len(a_tag_links))
        print(a_tag_links)

        profile_links= [m.get_attribute('href') for m in a_tag_links
                 if '.com/' in m.get_attribute('href')]
    print(profile_links)
    mouse_move()
    return profile_links

def going_to_profile(name):
#going to profile
    driver.get("https://www.instagram.com/"+ name +"/")
    time.sleep(2)
    mouse_move()

def links_to_pfnames(profile_links):
#splitting the profile name from profile links
    links=[]
    links=profile_links
    profile_names=[]
    mouse_move()
    dummy=[]
    for i in profile_links:
        m = i.split('https://www.instagram.com/')
        n = m[1]
        q = (n.rstrip('/'))
        print(q)
        dummy.append(q)

# to remove duplicates in the new list

    for i in dummy:
        if i not in profile_names:
            profile_names.append(i)
    return profile_names

def liker_names_in_excelsheet():
#existing likers list
    mouse_move()
    df = pd.read_excel('likes_list.xlsx', sheet_name=0)
    likes_list= df['List'].tolist()
    print(len(likes_list))
    return likes_list

def comparing_new_likers_with_old(profile_names,likes_list):
# removing old names from new list
    list=profile_names
    likes=likers_list
    for n in range(0,10):
        for i in list:
            if i not in likes:
                continue
            else:
                list.remove(i)

    new_profile_names=list
    mouse_move()
    print('final new profile list',new_profile_names)
    #print(list)
    return new_profile_names

def update_excel_likers_list(new_profile_names):
#updating new profile list in the excel sheet
    df=pd.DataFrame(new_profile_names)
# Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter('likes_list.xlsx', engine='openpyxl')
# try to open an existing workbook
    book = load_workbook('likes_list.xlsx')
    writer.book = book
# copy existing sheets
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
#reading the excel sheet
    reader = pd.read_excel(r'likes_list.xlsx')
#updating the excel
    df.to_excel(writer, index=False, header=False, startrow=len(reader) + 1)
    writer.save()
    mouse_move()

def getting_post_links(new_profile_links):
#getting the post links from profiles
    no_of_profiles = len(new_profile_links)
    count = 1
    for g in new_profile_links:
        print('profile started:', g)
        going_to_profile(g)
        time.sleep(3)

        a_tag_links = []
        post_links = []
        mouse_move()
        for i in range(1, 2):
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            a_tag_links = driver.find_elements_by_tag_name('a')
            #count=print('total no of posts from'+ g+' :'+len(a_tag_links))
            post_links = [m.get_attribute('href') for m in a_tag_links if '.com/p/' in m.get_attribute('href')]
        mouse_move()
        post_like(post_links,no_of_profiles,count,g)
        count= count + 1

def post_like(post_links,no_of_profiles,count,profile):
#liking the post
    posts_count=len(post_links)
    q=0
    mouse_move()
    for l in post_links:
        q = q + 1
        print(str(count)+'out of '+str(no_of_profiles)+'profiles')
        print(str(q)+'out of'+str(posts_count)+'posts')
        print(l)
        try:

            driver.get(l)
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

            time.sleep(2)

            # scroll up page
            driver.execute_script("window.scrollTo(0, 0);")

            like1 = driver.find_element_by_class_name('fr66n').click()
            time.sleep(2)
            if  q == posts_count:
                mouse_move()
                sending_messages(profile)
        except Exception as e:
            print(e)
#post_comment()
            if  q == posts_count:
                mouse_move()
                sending_messages(profile)


def post_comment():
    print()
'''
#commenting the post
                comment_button=driver.find_element_by_xpath("/html/body/div[1]/section/main/div/div[1]/article/div[3]/section[1]/span[2]/button")
                time.sleep(2)
                comment_button.click()
                time.sleep(2)
                comment=driver.find_element_by_xpath("/html/body/div[1]/section/main/div/div[1]/article/div[3]/section[3]/div/form/textarea")
                time.sleep(2)
                comm='nice !!!!!!!'
                comment.send_keys(comm)
                time.sleep(4)
                comment1=driver.find_element_by_xpath("/html/body/div[1]/section/main/div/div[1]/article/div[3]/section[3]/div/form/button")
                comment1.click()
                time.s
'''

def sending_messages(profile):
#sending youtube messages

    going_to_profile(profile)

    time.sleep(2)
    try:

        like1 = driver.find_element_by_xpath(
            "//Button[@type='button']").click()
        time.sleep(2)
        txtarea = driver.find_element_by_xpath(
            "/html/body/div[1]/section/div/div[2]/div/div/div[2]/div[2]/div/div[2]/div/div/div[2]/textarea")
        msg1 = "Hi i liked your posts,as i said "
        time.sleep(2)

        txtarea.click()
        txtarea.send_keys(msg1)
        mouse_move()

        time.sleep(2)
        msg_button = driver.find_element_by_xpath(
               "/html/body/div[1]/section/div/div[2]/div/div/div[2]/div[2]/div/div[2]/div/div/div[3]/button")
        msg_button.click()

        msg2 = "https://www.youtube.com/watch?v=TyOxH2355d4"
        time.sleep(2)
        txtarea.click()
        time.sleep(2)
        txtarea.send_keys(msg2)
        time.sleep(2)
        msg_button1 = driver.find_element_by_xpath(
                "/html/body/div[1]/section/div/div[2]/div/div/div[2]/div[2]/div/div[2]/div/div/div[3]/button")
        msg_button1.click()
        mouse_move()

        time.sleep(2)
        msg3 = "can you  please subscribe it"
        time.sleep(2)
        txtarea.send_keys(msg3)

        time.sleep(2)
        msg_button2 = driver.find_element_by_xpath(
                "/html/body/div[1]/section/div/div[2]/div/div/div[2]/div[2]/div/div[2]/div/div/div[3]/button")
        msg_button2.click()
        time.sleep(2)
        desktop_notification(profile)
        print('messages are sent to:', profile)
    except Exception as e:
        print('error occured so couldnt send messages',profile)

def send_email(final_list):
# to send the email
    #mail_id=['chessmanikanta@gmail.com','varuna.nimmala.vn@gmail.com']
    mail_id=['chessmanikanta@gmail.com']
# message to be sent
    #print(final_list)
    message = "Likes and messages are done for: " + ' , '.join(final_list)
    print(message)

    for i in mail_id:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login("Username", "PWD")
        server.sendmail('Username', i , message)
        print('mail sent to '+ i)
        server.quit()


username = '******************'
password = '******************'
name='laughingperson20'
#love_quotes_of_life95
time.sleep(1)
try:
        voice(welcome_message)
        login(username,password)
        time.sleep(1)
        instagram_desktop_notification()
        time.sleep(1)
        profile_links = getting_links_from_messagees(name)
        time.sleep(1)
        profile_names=links_to_pfnames(profile_links)
        time.sleep(1)
        likers_list=liker_names_in_excelsheet()
        time.sleep(1)
        final_list=comparing_new_likers_with_old(profile_names,likers_list)
        time.sleep(1)
        mouse_move()
        update_excel_likers_list(final_list)
        time.sleep(1)
        getting_post_links(final_list)
        send_email(final_list)
        voice(Done)
except Exception as e:
    print(e)
    voice(Error)
