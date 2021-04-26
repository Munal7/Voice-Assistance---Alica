#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import speech_recognition as sr
import subprocess
import wolframalpha
import pyttsx3
import json
import random
import operator
import datetime
import wikipedia
import webbrowser
import os
import sys
import pyjokes
from PIL import Image
import pyautogui
import winshell
import smtplib
import ctypes
import datetime
import requests
import shutil
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
import screen_brightness_control as sbc
import time
import psutil
import math
import socket
import sounddevice as sd
from scipy.io.wavfile import write
import wavio as  wv
from playsound import playsound


engine = pyttsx3.init('sapi5')
client = wolframalpha.Client('LA8K2L-5JJEXY4J32')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[len(voices)-1].id)

def speak(audio):
    
    engine.setProperty('voice',voices[1].id)
    print('Alica: ' + audio)
    engine.say(audio)
    engine.runAndWait()

def myCommand():
   
    r = sr.Recognizer()                                                                                   
    with sr.Microphone() as source:                                                                       
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)
    try:
        query = r.recognize_google(audio, language='en-in')
        print('User: ' + query + '\n')
        
    except sr.UnknownValueError:
        speak("Sorry sir! I didn't get that! Try typing the command!")
        query = str(input('Command: '))

    return query
    
def main():
    while True:
        speak('Welcome ! THis is ALICA your Voice assistance.')
        speak('Voice Activation Required ')
        speak('Enter User name')
        UserName = input ("Enter Username: ")
        speak('Password please !')
        PassWord = myCommand().lower()

        if UserName == 'Munal' and PassWord == 'school':
            time.sleep(1)
            playsound("success.mp3")
            speak('login successful!')
            logged()
            break

        else:
            playsound("abort.mp3")
            speak('Access denied !')

def logged():
    time.sleep(1)

main()





def greetMe():
    currentH = int(datetime.datetime.now().hour)
    if currentH >= 0 and currentH < 12:
        speak('Good Morning!')

    if currentH >= 12 and currentH < 18:
        speak('Good Afternoon!')


    if currentH >= 18 and currentH < 20:
        speak('Good Evening!')

    if currentH >= 20 and currentH != 0:
        speak('Good Evening!')

greetMe()

speak('Hello Mister Munal Sir ! I am your digital assistant , ALice.')

speak('How can I help you, Sir!')



def record():
    # Sampling frequency
    freq = 44100
  
    # Recording duration
    duration = 10
  
    # Start recorder with the given values 
    # of duration and sample frequency
    recording = sd.rec(int(duration * freq), 
                   samplerate=freq, channels=2)
  
    # Record audio for the given number of seconds
    sd.wait()
  
    # This will convert the NumPy array to an audio
    # file with the given sampling frequency
    write("recording0.wav", freq, recording)
  
    # Convert the NumPy array to audio file
    wv.write("recording1.wav", recording, freq, sampwidth=2)
    
    
def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
     
    # Enable low security in gmail
    server.login('munalthapa7@gmail.com', 'Munal@thapa$123')
    server.sendmail('munalthapa7@gmail.com', to, content)
    server.close()

def tellDay():
      
    # the weekday method is a method from datetime
    # library which helps us to an integer 
    # corresponding to the day of the week
    # this dictionary will help us to map the
    # integer with the day and we will check for
    # the condition and if the condition is true
    # it will return the day
    day = datetime.datetime.today().weekday() + 1
      
    Day_dict = {1: 'Monday', 2: 'Tuesbyeday', 3: 'Wednesday',
                4: 'Thursday', 5: 'Friday', 6: 'Saturday',
                7: 'Sunday'}
      
    if day in Day_dict.keys():
        day_of_the_week = Day_dict[day]
        print(day_of_the_week)
        speak("The day is " + day_of_the_week)
        
def tellTime():
# This method will give the time
    time = str(datetime.datetime.now())
      # the time will be displayed like this "2020-06-05 17:50:14.582630"
    # nd then after slicing we can get time
    print(time)
    hour = time[11:13]
    min = time[14:16]
    speak("The time is sir  " + hour + "Hours and  " + min + "Minutes") 

def getBatteryStatus():
    myBattery = psutil.sensors_battery()
    batteryStatus = math.floor(myBattery.percent)
    speak("The actual battery status is {} percent !".format((str(batteryStatus))))
    
    




if __name__ == '__main__':
      

    while True:
    
        
        query = myCommand();
        query = query.lower()
        
        # Browser Module
        
        
        if 'open youtube' in query:
            speak('okay, here you go !')
            webbrowser.open('www.youtube.com')

        elif 'open google' in query:
            speak('okay, here  you go !')
            webbrowser.open('www.google.co.in')
        
        elif 'open facebook' in query:
            speak('okay here  you go !')
            webbrowser.open('www.facebook.com')
            
        elif 'open instagram' in query:
            speak('okay here  you go !')
            webbrowser.open('www.instagram.com')

        elif 'open gmail' in query:
            speak('okay here  you go !')
            webbrowser.open('www.gmail.com')
            
        elif 'open wikipedia' in query:
            speak('okay here  you go !')
            webbrowser.open('www.wikipedia.com')
            
        elif 'open my mail' in query or 'open my gmail' in query or 'check my mail' in query:
            speak('okay here  you go !')
            webbrowser.open('https://mail.google.com/mail/u/0/#inbox')
            
        elif 'open stack overflow' in query:
            speak("Here you go to Stack Over flow.Happy coding")
            webbrowser.open("stackoverflow.com")  
            
        elif 'search' in query:
             
            query = query.replace("search", "")
            query = query.replace("play", "")         
            webbrowser.open(query)
      
            
        elif "where is" in query:
            query = query.replace("where is", "")
            location = query
            speak("User asked to Locate")
            speak(location)
            webbrowser.open("https://www.google.nl/maps/place/" + location + " ")
            speak('Here is the location of' + location)
        
        # END Browser
        
    
        elif "battery"  in query or "battery status" in query or "Battery status" in query:
            getBatteryStatus()
            time.sleep(1)    
        
        # OS MOdule
        elif 'c drive' in query:
            speak("Opening c drive")
            cpath = "C:\\"
            os.startfile(cpath)
            
        elif 'd drive' in query:
            speak("Opening D drive")
            dpath = "D:\\"
            os.startfile(dpath)
        
        elif 'e drive' in query:
            speak("Opening E drive")
            epath = "E:\\"
            os.startfile(epath) 
            
        elif "create a new folder" in query:
            speak("creating new folder")
            speak("folder name please")
            newfolder=input()
            os.mkdir("E:\\" +newfolder)
            speak("your folder is created")
            
        elif "delete a folder" in query:
            speak("Type a folder name which you want to delete")
            folder=input()
            os.rmdir("E:\\" +folder)
            speak("your folder is deleted")
            
        
        
        elif "control" in query:
            os.system('control')
        
        elif "record my voice" in query:
            speak('okey ! your voice is recording')
            record()
        
        elif "play my recorded sound" in query:
            speak("okey")
            playsound('recording1.wav')
          
            
        elif 'open microsoft word' in query:
            speak("opening microsoft word")
            power = r"E:\\munal\\ms word.DOCM"
            os.startfile(power)
        
        elif 'open power point' in query:
            speak("opening power point")
            power = r"E:\\munal\\power point.pptx"
            os.startfile(power)
            
        elif 'open excel' in query:
            speak("opening excel")
            power = r"E:\\munal\\ms excel.XLSX"
            os.startfile(power)
            
        elif 'open notepad' in query or 'open NotePad' in query:
            speak("openinig notepad")
            power = r"E:\\munal\\munal.txt"
            os.startfile(power)
            
        elif 'open my picture' in query or 'show my picture'in query:
            speak("openinig your picture")
            power = r"E:\\munal\\munal.JPG"
            os.startfile(power)
        
        elif 'open my notes in ms word' in query or 'open my notes in Ms Word' in query:
            speak("openinig your notes")
            power = r"C:\\users\\munal\\final_project\\munal.DOCM"
            os.startfile(power)
        
        elif 'play music' in query or "play song" in query or 'play music for me' in query:
            speak("Here you go with music")
            # music_dir = "G:\\Song"
            music_dir = "E:\\Video"
            songs = os.listdir(music_dir)
            print(songs)   
            random = os.startfile(os.path.join(music_dir, songs[1]))
                  
            speak('Okay, here is your music! Enjoy!')
            
        elif 'play next' in query:
            speak("Here you go with music")
            music_dir = 'E:\\Video'
            songs = os.listdir(music_dir)
            print(songs)
            speak("Playing the next song ")
            os.startfile(os.path.join(music_dir, songs[random.randint(0, len(songs) - 1)]))
           
            
        elif 'play previous' in query:
            music_dir = 'E:\\Video'
            songs = os.listdir(music_dir)
            print(songs)
            speak("Playing the previous song ")
            os.startfile(os.path.join(music_dir, songs[random.randint(0, len(songs) + 1)]))    
       
        elif 'disconnect' in query or 'disconnect wifi' in query or 'offline' in query or 'go offline' in query:
            speak('disconnecting from internet sir..')
            os.system('netsh wlan disconnect')
        # END OS 
       
            
        
        # Ctypes 
        elif 'change my background' in query or 'change background' in query:
            ctypes.windll.user32.SystemParametersInfoW(20,
                                                       0,
                                                       "E:\\munal\\munal.JPG",
                                                       0)
            speak("Background changed succesfully")
            
        elif 'lock window' in query or 'lock my window' in query or 'lock my laptop'in query or 'lock my computer' in query:
                speak("locking the device")
                ctypes.windll.user32.LockWorkStation()   
            
         # END Ctypes   
        
        
      # Brightness Control
        elif "brightness" in query:
            old = sbc.get_brightness()
            if "increase" in query or "raise" in query:
                sbc.set_brightness(old + 25)
                speak(f"increasing brightness")
            elif "max" in query:
                sbc.set_brightness(100)
                speak(f"maximum brightness")
            elif "decrease" in query or "lower" in query:
                speak(f"decreasing brightness")
                sbc.set_brightness(old - 25)
            elif "min" in query:
                speak(f"Mnimun brightness")
                sbc.set_brightness(0)
            else:
                speak(f"Current brightness is {old}")
        
        #datetime
        
        elif 'tell me the day' in query or 'day' in query:
            tellDay()
            
        elif "tell me the time" in query:
            tellTime()
            
        
        #end datetime
        
        #pyautogui and PIL Screenshot
        
        elif "capture my screen" in query or 'screenshot'in query:
            myScreenshot = pyautogui.screenshot()
            playsound("screenshot.wav")
            speak('done !')
            myScreenshot.save('E:/munal/screen.png')
        
        elif "show me the screen shot image" in query:
            speak("okey")
            power = r"E:\\munal\\screen.png"
            os.startfile(power)
            
        #end screen shot
       
        elif "don't listen" in query or "stop listening" in query:
            speak("for how much time you want to stop jarvis from listening commands")
            a = int(input())
            speak('okey!' )
            time.sleep(a)
           
        
        #pyjokes
        
        elif 'joke' in query:
            speak(pyjokes.get_joke())
            
        #End pyjokes
        
        #Winshell
        
        elif 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
            speak("Recycle Bin Recycled")
        
        #END winshell
    
       
        
       #sub process 
        
        elif 'shutdown system' in query:
                speak("Hold On a Sec ! Your system is on its way to shut down")
                subprocess.call('shutdown / p /f') 
            
        elif "restart" in query:
            subprocess.call(["shutdown", "/r"])
             
        elif "hibernate" in query or "sleep" in query:
            speak("Hibernating")
            subprocess.call("shutdown / h")
 
        elif "log off" in query or "sign out" in query:
            speak("Make sure all the application are closed before sign-out")
            time.sleep(5)
            subprocess.call(["shutdown", "/l"])
           
        #END sub process 
        
        #Smtplib 
        elif 'email to munal' in query:
            try:
                speak("What should I say?")
                content = myCommand()
                to = "Receiver email address"   
                sendEmail(to, content)
                speak("Email has been sent !")
            except Exception as e:
                print(e)
                speak("I am not able to send this email")
 
        elif 'send a mail' in query:
            try:
                speak("What should I say?")
                content = myCommand()
                speak("whom should i send")
                to = input()   
                sendEmail(to, content)
                speak("Email has been sent !")
            except Exception as e:
                print(e)
                speak("I am not able to send this email")
                
        #END Smtplib
        elif "calculate" in query:
            app_id = "58G6Y5-X253RUAUY3"
            client = wolframalpha.Client(app_id)
            indx = query.lower().split().index('calculate')
            query = query.split()[indx + 1:]
            res = client.query(' '.join(query))
            answer = next(res.results).text
            print("The answer is " + answer)
            speak("The answer is " + answer)
        
        #SYS module
        
        elif 'nothing' in query or 'abort' in query or 'stop' in query or 'bye bye' in query:
            speak('okay')
            speak('Bye Sir, have a good day.')
            sys.exit()
            
        elif 'bye' in query:
            speak('Bye Sir, have a good day.')
            sys.exit()
            
        #end sys module
        
        # socket module
        
        elif 'pc name' in query or 'computer name' in query:
            hostname = socket.gethostname()
            speak('your '+query+' is:'+hostname)
            
        elif 'ip' in query or 'my ip' in query or 'show ip' in query or 'show my ip' in query:
            hostname = socket.gethostname()  
            IP = socket.gethostbyname(hostname) 
            speak('your ip address is:'+IP)
            
         #end Socket Module
        
        
        elif "weather" in query:
             
            # Google Open weather website
            # to get API of Open weather
            api_key = "330964eb864b0a2f0af0533fcde28906"
            base_url = "http://api.openweathermap.org/data/2.5/weather?"
            speak(" City name ")
            print("City name : ")
            city_name = myCommand()
            complete_url = base_url + "appid =" + api_key + "&q =" + city_name
            response = requests.get(complete_url)
            x = response.json()
             
            if x["cod"] != "404":
                y = x["main"]
                current_temperature = y["temp"]
                current_pressure = y["pressure"]
                current_humidiy = y["humidity"]
                z = x["weather"]
                weather_description = z[0]["description"]
                print(" Temperature (in kelvin unit) = " +str(current_temperature)+"\n atmospheric pressure (in hPa unit) ="+str(current_pressure) +"\n humidity (in percentage) = " +str(current_humidiy) +"\n description = " +str(weather_description))
             
            else:
                speak(" City Not Found ")
                
        elif "write a note" in query:
            speak("What should i write, sir")
            note = myCommand()
            file = open('munal.DOCM', 'w')
            file.write(note)
            speak('Your note is saved succesfully !')
            
        elif "delete my notes" in query or "remove my notes" in query:
            speak("Your Notes Will Be Deleted. Are You Sure")
            ans= myCommand()
            if 'yes' in ans or 'sure' in ans:
                with open("C:\\users\\munal\\final_project\\munal.DOCM", "w") as f:
                    f.write(" ")
                    speak("note has been deleted ! ")
            else:
                with open("C:\\users\\munal\\final_project\\munal.DOCM", "r") as f:
                    f.write(note)
                
                
        elif "show my note" in query or "show my notes" in query:
            speak("Showing Notes")
            file = open("munal.DOCM", "r")
            print(file.read())
            speak(file.read(6))  
            
        

        
        #normal chat
            
        elif 'hello' in query or 'hey'in query or 'hola' in query or 'hi 'in query:
            stMsgs = ['hola !', 'hey !', 'hello sir !','Hiya','I am here....', 'Hi there !']
            speak(random.choice(stMsgs))

        elif "what's up" in query or 'how are you' in query or 'how you doing' in query :
            stMsgs = ['All good here.','I am preety good,Thanks!','Just doing my thing!', 'I am fine! thanks.', 'Nice!', 'I am nice and full of energy','All good','Booring please talk to me','Great !',"I am happy to be here.",'Not too shabby. Thanks for asking.','Hi there. How can i help?','I am good. Thanks for asking']
            speak(random.choice(stMsgs))
            
        elif "k cha hajur" in query or "k cha khabar" in query or 'namasteey' in query:
            stMsgs = ['Thick cha hajurr ','Namastey !', 'I am fine! thanks.','All good','Great !',"I am happy to be here.",'Hi there. How can i help?','I am good. Thanks for asking']
            speak(random.choice(stMsgs))
        
    
        elif "will you be my gf" in query or "will you be my bf" in query:  
            speak("I'm not sure about, may be you should give me some time")
            

       
           
        
            
        elif "who made you" in query or "who created you" in query:
            speak("I have been created by Mister Munal Thapa and his group member.")
            

        elif "who i am" in query:
            speak("you are my definately my Creater.")
 
        elif "why you came to world" in query:
            speak("Thanks to Mister Munal. further It's a secret")
 
        elif 'is love' in query:
            speak("It is 7th sense that destroy all other senses")
 
        elif "who are you" in query:
            speak("I am your virtual assistant created by Mister Munal Thapa Mukesh thapa bilip bir paudel and sharmila chaudary ")
 
        elif 'reason for you' in query:
            speak("I was created as a Majorr project by Mister Munal Thapa Mukesh thapa bilip bir paudel and sharmila chaudary ")
        
        elif "change name" in query:
            speak("What would you like to call me, Sir ")
            assname = myCommand()
            speak("Thanks for naming me")
            
       
        elif "i love you" in query:
            speak("It's hard to understand")     
        

        else:
            query = query
            speak('Searching...')
            try:
                try:
                    res = client.query(query)
                    results = next(res.results).text
                    speak('WOLFRAM-ALPHA says - ')
                    speak('Got it.')
                    speak(results)
                    
                except:
                    results = wikipedia.summary(query, sentences=2)
                    speak('Got it.')
                    speak('WIKIPEDIA says - ')
                    speak(results)
        
            except:
                webbrowser.open('www.google.com')
        
        speak('Next Command! Sir!')
        


        














