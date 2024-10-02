import subprocess
import wolframalpha
import pyttsx3
import sys
import tkinter
import cv2
import socket
import socketserver
import json
import geocoder
import random
import operator
import pyautogui
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import pyaudio
import PyPDF2
import winshell
import os
import winshell
import instaloader
import pyjokes
import feedparser
import smtplib
import ctypes
import time
from requests import get 
import shutil
from twilio.rest import Client
from clint.textui import progress
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
import pywhatkit as kit
import requests
from pywikihow import search_wikihow
import psutil
import speedtest
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders




# Text  to Speech
engine = pyttsx3.init("sapi5")
"""for voice in engine.getProperty('voices'):
    id='HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_EN-US_DAVID_11.0'
engine.setProperty('voice',id)
engine.say('Hello Welocome')
engine.runAndWait()"""
voices=engine.getProperty('voices')
engine.setProperty('voices',voices[0].id)
engine.setProperty('rate', 200)
"""for voice in voices:
    print(voice.id)
    engine.setProperty('voice', voice.id)
    engine.say("hello sir i am your virtual assistant")
engine.runAndWait()"""


def speak(audio):
    engine.say(audio)
    print(audio)
    engine.runAndWait()
#=====================Wish Me Function==================================================
def wishMe():
	hour = int(datetime.datetime.now().hour)
	if hour>= 0 and hour<12:
		speak("Good Morning Sir !")

	elif hour>= 12 and hour<18:
		speak("Good Afternoon Sir !")

	else:
		speak("Good Evening Sir !")

	assname =("Jarvis 1 point o")
	speak("I am your Assistant")
	speak(assname)

 
	

"""def username():
	speak("What should i call you sir")
	uname = takeCommand()
	speak("Welcome Mister")
	speak(uname)
	columns = shutil.get_terminal_size().columns
	
	print("#####################".center(columns))
	print("Welcome Mr.", uname.center(columns))
	print("#####################".center(columns))
	
	speak("How can i Help you, Sir")"""


        
def takeCommand():
        r = sr.Recognizer()
        with sr.Microphone() as source:
            print("Listening...")
            r.pause_threshold = 1
            r.adjust_for_ambient_noise(source)
            audio = r.listen(source)
            #audio = r.listen(source,timeout=4,phrase_time_limit=7)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language ='en-in')
            print(f"User said: {query}\n")
        except Exception as e:
            #speak("say that again please...")
            return "None"
        query = query.lower()
        return query



# FOR NEWS UPDATES
def news():
    main_url = 'http://newsapi.org/v2/top-headlines?sources-techcrunch&apiKey-ca2ec6a05d7f4ca68eea0331e25373d9'
    
    main_page = requests.get(main_url).json()
    #print(main_page)
    articles = main_page["article"]
    #print(articles)
    head = []
    day=["first","second","third","fourth","fifth","sixth","seventh","eighth","ninth","tenth"]
    for ar in articles:
        head.append(ar["title"])
    for i in range (len(day)):
        #print("today's {day[i]} news is: ", head[i])
        speak(f"today's {day[i]} news is: {head[i]}")
# TO READ PDFs   
def pdf_reader():
    book = open('py3.pdf','rb')
    pdfReader = PyPDF2.PdfFileReader(book) #pip install PyPDF2
    pages = pdfReader.numpages
    speak(f"Total number of pages in this book {pages}")
    speak("Sir please enter the page number i have to read")
    pg = int(input("please enter the page number:"))
    page = pdfReader.getPage(pg)
    text = page.extractText()
    speak(text)
    # jarvis speaking speed should be controlled by user   
   
# TO SEARCH ANY THING IN YOUTUBE   
def search_youtube(query):
    search_url = f"https://www.youtube.com/results?search_query={query.replace(' ', '+')}"
    webbrowser.open(search_url)
    print(f"Searching YouTube for: {query}") 

def sendEmail(to, content):
	server = smtplib.SMTP('smtp.gmail.com', 587)
	server.ehlo()
	server.starttls()
	
	# Enable low security in gmail
	server.login('2k22.cse.2213378@gmail.com', 'Sarthak@123')
	server.sendmail('2k22.cse.2213378@gmail.com', to, content)
	server.close()
#------------------------------------ Mathematical/Bitwise Operation-------------------------------------------------------
def abs(a):
    "same as abs(a)."
    return abs(a)

def add(a,b):
    "same as a+b"
    return a+b

def  and_(a,b):
    "same as a & b."
    return a & b

def floordiv(a,b):
    "same as a // b"
    return a // b

def index(a):
    "same as a.__index __()"
    return a._index_() 

def inv(a):
    "same as ~a."
    return ~a
invert=inv

def lshift(a,b):
    "same as a << b."
    return a << b 

def mod(a,b):
    "same as a % b."
    return a % b


def search_wikihow(query, max_results=10, lang="en"):
    return list(search_wikihow(query, max_results, lang))

def take_photo(save_path):
    # Initialize the camera
    camera = cv2.VideoCapture(0)  # Use the default camera (change the parameter if using a different camera index)

    if not camera.isOpened():
        print("Cannot open camera")
        return

    # Capture a photo
    ret, frame = camera.read()

    if not ret:
        print("Failed to capture image")
        camera.release()
        return

    # Save the captured photo
    cv2.imwrite(save_path, frame)

    # Release the camera
    camera.release()
    print(f"Photo saved at: {save_path}")

# Replace 'path/to/save/photo.jpg' with the desired save location
save_path = 'C:/Users/Akansha/Pictures.jpg'
take_photo(save_path)
   

def run(self):
    #self.TaskExecution()
    speak("please say wakeup to continue")
    while True:
        self.query = self.takeCommand()
        if "wake up" in self.query or "are you there" in self.query or "hello" in self.query:
            self.TaskExecution


def TaskExecution():
    wishMe()
    while True:
        query = takeCommand()
        
        # logic building for tasks
        
         #---------------------------------To open notepad------------------------------------    
        if 'open Notepad' in query:
            npath = "C:\\Windows\\system32\\notepad.exe"
            os.startfile(npath)
        
        #-----------------------------To open adobe reader-----------------------------------    
        elif 'open adobe reader' in query:
            apath = "C:\\Program Files\\Adobe\\Acrobat DC\\Acrobat"
            os.startfile(apath)
            
        #------------------------To open wikipedia---------it  is not working------------------------------
        elif 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences = 2)
            speak("According to Wikipedia")
            print(results)
            speak(results)
        #------------------------------ To open Youtube-------------------------------------------------   
        elif 'search in youtube for' in query:
            parts = query.split("search in youtube for")
            query = parts[1].strip()
            search_youtube(query)
            #speak("Here you go to Youtube\n")
            #webbrowser.open("https://www.youtube.com")
            """if 'search' in query: 
                speak("{MASTER} what do you want to search")    
                query2 = None        
                while query2 is None:  
                    r2 = sr.Recognizer()  
                    with sr.Microphone() as source:
                        print("Listening...")
                        speak("Beep")      
                        audio2 = r2.listen(source, 2)   
                        command = r2.recognize_google(audio2)   
                        print(command)           
                    try:              
                        print("Recognising...")         
                        query2 = r2.recognize_google(audio2, language= 'en-in') 
                        print(f"user said: {query2}\n")                 
                        print(query2)                 
                        break
                    except Exception:         
                        print("Say that again please...")          
                        speak("Say that again please...")          
                        query2 = None 
                        query2 = query2.replace(" ","+")     
                        print("eee")         
                        crome_path = "C:\Program Files\Google\Chrome\Application\chrome.exe %s"       
                        webbrowser.get(crome_path).open(url = "https://www.youtube.com"+query2)  
            else:       
                speak("Here you go to Youtube\n")
                webbrowser.open("https://www.youtube.com")"""
        #--------------------------------------------open Youtube----------------------------------------------
        elif "open youtube" in query:
            speak("Here you go to Youtube\n")
            webbrowser.open("https://www.youtube.com")
                
        #---------------------------------To open facebook-----------------------------------    
        elif 'open facebook' in query:
            speak("here you go to facebook \n")
            webbrowser.open("https://www.facebook.com")
        #------------------------To open stackoverflow-------------------------------------
        elif 'open stackoverflow' in query:
            webbrowser.open("www.stackoverflow.com")
        #---------------------------To open google----------------------------------------    
        elif 'open google' in query:
            speak("Here you go to Google\n")
            speak("Sir, what should i search on google")
            cn = takeCommand().lower()
            webbrowser.open(f"{cn}")
        #------------------------------To send whatsapp message--------------------------------------------    
        
        
        #--------------------------------to open erppsit-------------------------------------
        elif 'open erp' in query:
            webbrowser.open("https://erp.psit.ac.in/")
            
        #----------------send message with twilio-------------------------
        elif "send message" in query:
            speak("sir what should i say")
            msz = takeCommand()
            # Find your Account SID and Auth Token at twilio.com/console
            # and set the environment variables. See http://twil.io/secure
            account_sid = 'AC7a3570a7e546c25edc66a05d1e6cb839'
            auth_token = '08b3c730b26e9141ae7dd34c6a941fb3'
            client = Client(account_sid, auth_token)
            
            twilio_phone_number = '+13146495774'
            recipient_phone_number = '+917088977083'   
            
            client = Client(account_sid, auth_token)
            
            message = client.messages.create(
                body=msz,
                from_=twilio_phone_number,
                to=recipient_phone_number
            )
            print(f"Message sent! Message SID: {message.sid}")
            speak("sir, message has been sent")
        
        
        #------------------------------------------------for Calling to your number----------------------------------------------------------
        elif "make phone call" in query:
            speak("ok, Sir i will calling ")
            # Find your Account SID and Auth Token at twilio.com/console
            # and set the environment variables. See http://twil.io/secure
            account_sid = 'AC7a3570a7e546c25edc66a05d1e6cb839'
            auth_token = '08b3c730b26e9141ae7dd34c6a941fb3'
            client = Client(account_sid, auth_token)
            
            twilio_phone_number = '+13146495774'
            recipient_phone_number = '+917088977083'   
            
            client = Client(account_sid, auth_token)
            
            call = client.calls.create(
                    twiml='<Response><Say>Hello! This is a test call from your virtual assistant.</Say></Response>',
                    to=recipient_phone_number,
                    from_=twilio_phone_number
            )
            speak(f"Call initiated! Call SID: {call.sid}")
  
        #------------------------To play song on youtube-------------------------------------------------------------------------------------    
        elif 'play song on youtube' in query:
            kit.playonyt("see you again")
            
        #-----------------------------to switch the windows----------------------------------------------------------------------------------
        elif 'switch the window' in query:
            pyautogui.keyDown("alt")
            pyautogui.press("tab")
            time.sleep(1)
            pyautogui.keyUp("alt")
            
        #-----------------------------------system volume up-------------------------------------------------------------------------------------
        elif 'volume up' in query:
            pyautogui.press("volumeup")
            
        #-----------------------------------system volume down---------------------------------------------------
        elif 'volume down' in query:
            pyautogui.press("volumedown")
            
        #-----------------s------------------------system volume mute----------------------------------------------
        elif 'volume mute' in query or 'mute' in query:
            pyautogui.press("volumemute")
            
        #------------------------To send E-mail--------------------------------------------------    
        elif "email to ayush" in query:
            """try:
                speak("what should i say")
                content = takeCommand().lower()
                to = "asmpatel221@gmail.com"
                sendEmail(to,content)
                speak("Email has been sent to ayush")
                
            except Exception as e:
                print(e)
                speak("sorry sir, i am not able to sent this mail to ayush")"""
            speak("sir what should i say")
            query = takeCommand()
            if "send a file" in query:
                email = "2k22.cse.2213378@gmail.com"
                password = "Sarthak@123" 
                send_to_email = "aryayadavji786@gmail.com"
                speak("okay sir, what is the subject for this email")
                query = takeCommand().lower()
                subject = query # the  subject in the email
                speak(" and sir, what is the message for this email")
                query2 = takeCommand().lower()
                message = query2 #the message in the email
                speak("sir please enter the corrct path of the file into the shell")
                file_location = input("please enter the path here")  #the file attachment in the mail
                
                speak("please wait, i am sending email now")
                
                msg =MIMEMultipart()
                msg['From'] = email
                msg['To'] = send_to_email
                msg['Subject'] = subject
                
                msg.attach(MIMEText(message, 'plain'))
                
                #setup the attachement
                filename = os.path.basename(file_location)
                attachment = open(file_location, "rb")
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header('content-Disposition',"attachment;filename= %s" % filename)
                # attach the attachment to the MIMEMultipart object
                msg.attach(part)
                
                server=smtplib.SMTP('smtp.gmail.com',587)
                server.startls()
                server.login(email,password)
                text=msg.as_string()
                server.sendmail(email,send_to_email,text)
                server.quit()
                speak("email has been sent to avinash")
                
            else:
                email='2k22.cse.2213378@gmail.com'#your email
                password="Sarthak@123"#your email account password
                send_to_email='2k22.cse.2213525@gmail.com'#whom your sending message to
                message=query# the message in the email
                
                server=smtplib.SMTP('smtp.gmail.com',587)#connect to the server
                server.starttls()#use tls
                server.login(email,password)#login to the email server
                server.login(email,password)#login to the email server
                server.sendmail(email,send_to_email.message)#send the email
                server.quit()#logout of the email server
                speak("email has been sent to avinash ") 
        #---------------------------------------------------------------------------------------------------------------
                     
        #elif "can you search om youtube" in query:
        #      to play something on youtube
        #      kit.playonyt("see you again")
        #---------------------Temperature------------------------------------------
        elif "temperature" in query:
            search = "temperature in kanpur"
            url = f"https://www.google.com/search?q={search}"
            r = requests.get(url)
            data = BeautifulSoup(r.text,"html.parser")
            temp = data.find("div",class_="BNeawe").text
            speak(f"current {search} is {temp}")
        
        #------------------------------------  Activate mode---------------------------------
       # elif "activate how to do " in query:
            #   speak("How to do mode is activate please tell me what you want to know")
            #how = takeCommand()
           # max_results = 1
           # how_to = search_wikihow(how, max_results)
            #assert len(how_to) == 1
           # how_to[0].print()
           # speak(how_to[0].summary)
        elif "activate how to do mod" in query:
            speak("how to do mode is  activated")
            #how = takeCommand()
            #max_result = 1
            #how_to = search_wikihow(how, max_results)
            #assert len(how_to) == 1
            #how_to[0].print()
            #speak(how_to[0].summary)
            while True:
                speak("please tell me what you want to know")
                how = takeCommand()
                try:
                    if "exit" in how or "close" in how:
                        speak("okay sir, how to do mode is closed")
                        break
                    else:
                        max_results = 1
                        how_to = search_wikihow(how, max_results)
                        assert len(how_to) == 1
                        how_to[0].print()
                        speak(how_to[0].summary)
                except Exception as e:
                    speak("sorry sir, I am not able to find this")
                    
                
       
        
        #------------------To open command prompt-----------------------------------------    
        elif 'open command prompt' in query:
            os.system("start cmd")
        #-------------------------To play music-----------------------------------
        elif 'play music' in query or "play song" in query:
           music_dir = "E:\\music"
           songs = os.listdir(music_dir)
           #rd = random.choice(song)
           #for song in songs:
           #    if song.endswith('.mp3'):
           #       os.startfile(os.path.join(music_dir, song))
           print(songs)
           os.startfile(os.path.join(music_dir, songs[0]))
           #sw = takeCommand().lower()
           #sys.exit(f"{sw}")
           
        #-------------------------To find the IP Address----------------------------------------          
        elif "ip address" in query:
             ip = get('https://api.ipify.org').text
             speak(f"your IP address is {ip}") 
             
        #-------------------to find my location  using ip Address------------------ 
        elif "where i am" in query or "where we are" in query:
            speak("wait sir, let me check")
            try:
                """ipAdd = requests.get('https://api.ipify.org').text
                print(ipAdd)
                url = 'https://get.geojs.io/v1/ip/geo'+ipAdd+'.json'
                geo_requests = requests.get(url)
                geo_data = geo_requests.json()
                #print(geo_data)
                city = geo_data['city']
                # state = geo_data['state]
                country = geo_data['country']
                speak(f"Sir i am not sure, but i think we are in {city} city of {country} country")"""
                g = geocoder.ip('me')
            
                if g.ok:
                    location = g.json
                    speak("\nYour current location details:")
                    speak(f"Latitude: {location['lat']}")
                    speak(f"Longitude: {location['lng']}")
                    speak(f"City: {location['city']}")
                    (f"Country: {location['country']}")
                else:
                    print("Unable to fetch location information.")
            except Exception as e:
                speak("Sorry sir, Due to network issue i am not able to find where we are.")
                pass
            
        #---------------------To check a instagram profile----------------------
        elif "instagram profile" in query or "profile on instagram" in query:
            speak("sir please enter the usser name correctly.")
            name = input("Enter username here:")
            webbrowser.open(f"www.instagram.com/{name}")
            speak(f"Sir here is the profile of the user {name}")
            time.sleep(5)
            speak("Sir would you like to download profile picture of this account.")
            condition = takeCommand().lower()
            if "yes" in condition:
                mod = instaloader.Instaloader() #pip install instadownloader
                mod.download_profile(name, profile_pic_only=True)
                speak("I am done sir, profile picture is saved in our main folder. now i am ready for next command")
            else:
                pass
        #-------------------To take screenshot-----------------------------
        elif "take screenshot" in query or "take a screenshot" in query:
            speak("sir, please tell me the name for this screenshot file")
            name = takeCommand().lower()
            speak("please sir hold the screen for few second, i am taking screenshot")
            time.sleep(3)
            img = pyautogui.screenshot()
            img.save(f"{name}.png")
            speak("i am done sir ,the screenshot is saved in our main folder. now i am ready for next command")
        # ----------------------To read PDF file-------------------------    
        elif "read pdf" in query:
            pdf_reader()
        #------------------------------To open camera and save file at given location-------------------------------------------------    
        elif "open camera" in query or "take a photo" in query:
            take_photo(save_path)
            
        # ----------------------------to open camera and click photo and save to the vscode------------------------------    
        elif "camera" in query :
            ec.capture(0, "Jarvis Camera ", "img.jpg")
           
        #------------------------------to close any application----------------------------
        elif 'close notepad' in query:
            speak("okay Sir, closing notepad")
            os.system("taskkill /f /in notepad.exe")
        #--------------------to set an alarm-------------------------------------------
        elif 'set alarm' in query:
            nn = int(datetime.datetime.now().hour)
            if nn==22:
                music_dir = "E:\\music"
                songs = os.listdir(music_dir)
                os.startfile(os.path.join(music_dir, songs[0]))
        #----------------to find a joke------------------------------------------------
        elif 'tell me a joke' in query:
            joke = pyjokes.get_joke()
            speak(joke)
        #---------------------------To shut down the system-------------------------------------------    
        elif ' shut down the system' in query:
            os.system("shutdown /s /t 5")
        #------------To restart the system-------------------------------------------------------   
        elif 'restart the system' in query:
            os.system("shut down /r /t 5")
        #-------To sleep the system----------------------------------------------------------------    
        elif 'sleep the system' in query:
            os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
        #-------------------------To log off or sign out window-----------------------------------------\
        elif 'log off' in query or 'sign out' in query:
            speak("Make sure all the application are closed before sign out")
            time.sleep(5)
            subprocess.call(["shutdown", '/1'])
        #--------------For News----------------------------------------------------   
        elif "tell me news" in query:
            speak("please wait sir, fetching the latest news")
            news()
            
        #-------------------------------------telling the time--------------------------------
        elif "tell me the time" in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"sir, the time is {strTime}")
            
        # ------------------------------------To empty recycle bin-------------------------------------
        elif 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
            speak("Recycle Bin Recycled")    
        #----------------------------To do calculations-----------------------------------
        elif "do some calculations" in query or "can you calculate" in query:
            r = sr.Recognizer()
            with sr.Microphone() as source:
                speak("Say what you want to calculate, example: 3 plus 3")
                print("listening.....")
                r.adjust_for_ambient_noise (source)
                audio = r.listen(source)
            my_string=r.recognize_google(audio)
            print(my_string)
            def get_operator_fn(op):
                    return {'+': operator.add, #plus
                            '-': operator.sub, #minus
                            'x': operator.mul, #multiplied by
                            'divided' :operator.__truediv__, #divided
                            }[op]
            def eval_binary_expr (op1, oper, op2): # 5 plus 8 
                    op1, op2= int(op1), int (op2)
                    return get_operator_fn(oper) (op1, op2) 
            speak("your result is") 
            speak (eval_binary_expr(*(my_string.split())))
        #---------------------------To Hide files and folder-----------------------
        elif "hide all files" in query or "hide this folder" in query or "visible for everyone" in query:
            speak("sir please tell me you want to hide this folder or make it visible for everyone ")
            condition = takeCommand().lower()
            if "hide" in condition:
                os.system("attrib +h /s /d")#os module
                speak("sir, all the files in this folder are now hidden")
            
            elif "visible" in condition:
                os.system("attrib +h /s /d")
                speak("sir, all the files in this folder are now visible to everyone.")
            
            elif "leave it" in condition or "leave for now" in condition:
                speak("ok sir")
                
                
                
        #------------------------------------------------------------------------------        
        elif "hello" in query or "hey" in query:
            speak("hello sir, may i help you with something.")
        #------------------------------------------------------------------------------            
        elif "how are you" in query:
            speak("i am fine sir, what about you.")
        #------------------------------------------------------------------------------    
        elif "also good" in query or "fine" in query:
            speak("that's great to hear from you.")
        #------------------------------------------------------------------------------            
        elif "thank you" in query or "thanks" in query or "thank u" in query:
            speak("it's my pleasure sir.")
        #------------------------------------------------------------------------------       
        elif "you can sleep" in query or "sleep now" in query:
            speak("okay sir, i am going to sleep you can call me anytime.")
            break
        #----------------------------laptop power -----------------------------------------------
        elif "how much power left" in query or "how much power we have" in query or "battery" in query:
            battery = psutil.sensors_battery()
            percentage = battery.percent
            speak(f"sir our system have {percentage} percent battery")
            if percentage>=75:
                speak("we have enough power to continue our work")
            elif percentage>=40 and percentage<=75:
                speak("we shoud connect our system to charging point to charge our battery")
            elif percentage<=15 and percentage<=30:
                speak("we don't have enough power to work, please connect to charging")
            elif percentage<=15:
                speak("we have very low power, please connect to charging the system will shut down very soon")
        
        #-------------------------------------------internet speed----------------------------------------------------------
        elif "internet speed" in query:
            st = speedtest.Speedtest()
            st.get_best_server()
            download_speed = st.download() / 1024 / 1024  # in Mbps
            upload_speed = st.upload() / 1024 / 1024  # in Mbps
            ping = st.results.ping  # in ms

            speak(f"Download Speed: {download_speed:.2f} Mbps")
            speak(f"Upload Speed: {upload_speed:.2f} Mbps")
            speak(f"Ping: {ping} ms")
            #speak("sir we have {download_speed:.2f} bit per second downloading speed and {upload_speed:.2f} bit per second uploading speed and {ping} ms")
            #try:
             #   os.system('cmd /k "speedtest"')
            #except:
            #    speak("there is no internet connection")   
            
        
            
            
if __name__ == "__main__":
    while True:
        permission = takeCommand()
        if "wake up" in permission:
            speak("ok sir i am ready to work")
            TaskExecution()
        elif "goodbye" in permission:
            speak("thanks for using me sir, have a good day.")
            sys.exit()