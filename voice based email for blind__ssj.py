from email.message import EmailMessage
import select
import email
from email.base64mime import body_decode, body_encode
import speech_recognition as sr
import smtplib
import pyaudio
import platform
import sys
from bs4 import BeautifulSoup
import email
import imaplib
from gtts import gTTS
import pyglet
import os, time
import pyttsx3
from win32com.client import Dispatch #for speaking the content again

engine=pyttsx3.init()
s = Dispatch("SAPI.SpVoice")
mic = sr.Microphone()

def speak(text): #defining the speaking function 
	s.Speak(text)

def talk(text):
    engine.say(text)
    engine.runAndWait()

def get_voice_input(instruction):
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print(instruction)
        audio = r.listen(source)
        try:
            text = r.recognize_google(audio)
            print("Input: ", text)
            speak(text)
            return text
        except sr.UnknownValueError:
            print("Could not understand audio")
            return None

def login():
    # Provide instructions for blind users
    print("Welcome to the Voice Login System for email")
    talk("Welcome to the Voice Login System for email")
    talk("Please say 'login' to initiate login process.")
    while True:
        # Capture voice input to start login process
        command = get_voice_input("Please say 'login' to initiate login process.")
        if command and command.lower() == "login" or "l o g i n":
            break

    # Capture email address and password through voice input
    talk("Please say your user name:")
    user = get_voice_input("Please say your user name:")
    talk("Please say your password:")
    password = get_voice_input("Please say your password:")

    # Authenticate with email provider using email and password
    # You can use appropriate email API or libraries to implement this part
    if user == "admin" and password == "1234":
        print("Login successful!")
        talk("login successful")
        # Proceed with email access logic here
        tts = gTTS(text="Project: Voice based Email", lang='en')
        ttsname=("name.mp3") 
        tts.save(ttsname)

        music = pyglet.media.load(ttsname, streaming = False)
        music.play()

        time.sleep(music.duration)
        os.remove(ttsname)


        #login from os
        login = os.getlogin
        print ("You are logging from : "+login())

        #choices
        print ("1. composed a mail.")
        tts = gTTS(text="option 1. composed a mail.", lang='en')
        ttsname=("hello.mp3") 
        tts.save(ttsname)

        music = pyglet.media.load(ttsname, streaming = False)
        music.play()

        time.sleep(music.duration)
        os.remove(ttsname)

        print ("2. Check your inbox")
        tts = gTTS(text="option 2. Check your inbox", lang='en')
        ttsname=("second.mp3")
        tts.save(ttsname)

        music = pyglet.media.load(ttsname, streaming = False)
        music.play()

        time.sleep(music.duration)
        os.remove(ttsname)

        print ("3.fetch the email")
        tts = gTTS(text="option 3.fetch the email", lang='en')
        ttsname=("third.mp3")
        tts.save(ttsname)

        music = pyglet.media.load(ttsname, streaming = False)
        music.play()

        time.sleep(music.duration)
        os.remove(ttsname)

        

        print ("4.exit from the mail")
        tts = gTTS(text="option 4. exit from the mail", lang='en')
        ttsname=("four.mp3")
        tts.save(ttsname)

        music = pyglet.media.load(ttsname, streaming = False)
        music.play()

        time.sleep(music.duration)
        os.remove(ttsname)


        #this is for input choices
        tts = gTTS(text="Your choice ", lang='en')
        ttsname=("hello.mp3") 
        tts.save(ttsname)

        music = pyglet.media.load(ttsname, streaming = False)
        music.play()

        time.sleep(music.duration)
        os.remove(ttsname)



        #voice recognition part
        r = sr.Recognizer()
        with sr.Microphone() as source:
            print ("Your choice:")
            audio=r.listen(source)
            print("ok done!!")
            speak("ok done") #to speak the content

        try:
            text=r.recognize_google(audio)
            print ("You said : "+text)
            speak(text)
            
            
        except sr.UnknownValueError:
            talk("Google Speech Recognition could not understand audio.")
            
        except sr.RequestError as e:
            print("Could not request results from Google Speech Recognition service; {0}".format(e)) 

        #choices details
        if(text == '1' or text == 'One' or text == 'one' or text =='o n e'):
            r = sr.Recognizer()
            engine = pyttsx3.init() #recognize
            with sr.Microphone() as source:
                talk ("Your message :") #say any message
                audio=r.listen(source)
                print ("ok done!!")
            try:
                text=r.recognize_google(audio)
                print ("You said : "+text)
                speak(text)
                msg = text
            except sr.UnknownValueError:
                print("Google Speech Recognition could not understand audio")
            except sr.RequestError as e:
                print("Could not request results from Google Speech Recognition service; {0}".format(e))    



            def get_info():
                try:
                    with sr.Microphone() as source:
                        print('listening...')
                        voice = r.listen(source)
                        info = r.recognize_google(voice)
                        print(info)
                        talk("you said that")
                        speak(info)
                    
                        
                        return info.lower()
                except:
                    pass


            def send_email(receiver, subject, message):
                server = smtplib.SMTP('smtp.gmail.com', 587)
                server.starttls() # Make sure to give app access in your Google account
                server.login('project0sssj@gmail.com', 'mgnhisipzskvapwe')
                email = EmailMessage()
                email['From'] = 'Sender_Email'
                email['To'] = receiver
                email['Subject'] = subject
                email.set_content(message)
                server.send_message(email)


            email_list = {
            'f l o r a':'itsflorazone@gmail.com',
            's h a r m i':'babusharmi27@gmail.com',
            's a n d y' :'2023025@saec.ac.in',
            's a n t h i y a':'2023026@saec.ac.in'
            }


            def get_email_info():
                talk('To Whom you want to send email')
                name = get_info()
                receiver = email_list[name]
                print(receiver)
                talk('What is the subject of your email?')
                subject = get_info()
                
                talk('Tell me the body in your email')
                message = get_info()
                send_email(receiver, subject, message)
                talk('Your email is sent')
                talk("thank you for using this voice based email")
            get_email_info()
        elif (text == '2' or text == 'tu' or text == 'two' or text == 'Tu' or text == 'o' or text == 't o' or text == 'p o' ):
            print("hello")
            
            mail = imaplib.IMAP4_SSL('imap.gmail.com',993) #this is host and port area.... ssl security
            unm = ('project0sssj@gmail.com')  #username
            psw = ('mgnhisipzskvapwe')  #password
            mail.login(unm,psw)  #login
            stat, total = mail.select('Inbox')  #total number of mails in inbox
            print ("Number of mails in your inbox :"+str(total))
            tts = gTTS(text="Total mails are :"+str(total), lang='en') #voice out
            ttsname=("total.mp3") 
            tts.save(ttsname)
            music = pyglet.media.load(ttsname, streaming = False)
            music.play()
            time.sleep(music.duration)
            os.remove(ttsname)
            
            #seen mails
            unseen = mail.search(None, 'UnSeen') # unseen count
            print ("Number of seen mails :"+str(unseen))
            tts = gTTS(text="Your seen mail :"+str(unseen), lang='en')
            ttsname=("unseen.mp3") 
            tts.save(ttsname)
            music = pyglet.media.load(ttsname, streaming = False)
            music.play()
            time.sleep(music.duration)
            os.remove(ttsname)
            
            #search mails
            result, data = mail.uid('search',None, "ALL")
            inbox_item_list = data[0].split()
            new = inbox_item_list[-1]
            old = inbox_item_list[0]
            result2, email_data = mail.uid('fetch', new, '(RFC822)') #fetch
            raw_email = email_data[0][1].decode("utf-8") #decode
            email_message = email.message_from_string(raw_email)
            print ("From: "+email_message['From'])
            print ("Subject: "+str(email_message['Subject']))
            tts = gTTS(text="From: "+email_message['From']+" And Your subject: "+str(email_message['Subject']), lang='en')
            ttsname=("mail.mp3") 
            tts.save(ttsname)
            music = pyglet.media.load(ttsname, streaming = False)
            music.play()
            time.sleep(music.duration)
            os.remove(ttsname)
            
            #Body part of mails
            stat, total1 = mail.select('Inbox')
            stat, data1 = mail.fetch(total1[0], "(UID BODY[TEXT])")
            msg = data1[0][1]
            soup = BeautifulSoup(msg, "html.parser")
            txt = soup.get_text()
            print ("Body :"+txt)
            tts = gTTS(text="Body: "+txt, lang='en')
            ttsname=("body.mp3") 
            tts.save(ttsname)
            music = pyglet.media.load(ttsname, streaming = False)
            music.play()
            time.sleep(music.duration)
            os.remove(ttsname)
            mail.close()
            mail.logout()
        

        elif ( text == 'three' or text == 't h r e e' or text == 'fetch' or text =='f e t c h' or text == 'h r e'):
            print("hello")
            with mic as source:
    
                r.adjust_for_ambient_noise(source)
                speak("Please say the email address you want to fetch without saying @ gmail.com.")
                print("Please say the email address you want to fetch without saying @ gmail.com.")
                audio = r.listen(source)

            email_address = r.recognize_google(audio)
            rmaildomain = "@gmail.com"
            text_without_space = email_address.replace(" ", "")
            new=text_without_space + rmaildomain 



            # Connect to the IMAP server using SSL
            mail = imaplib.IMAP4_SSL('imap.gmail.com',993) 

            # Login to the email account
            mail.login('project0sssj@gmail.com', 'mgnhisipzskvapwe')

            # Select the inbox folder
            mail.select('inbox')

            # Search for emails sent from the specified email address
            result, data = mail.search(None, f'FROM "{new.lower()}"')

            if not data[0]:
                print(f"No emails found from {new.lower()}.")
                speak(f"No emails found from {new.lower()}.")
            else:
                # Get the email IDs of the searched emails
                email_ids = data[0].split()

                # Fetch the first email in the search results
                result, data = mail.fetch(email_ids[0], "(RFC822)")

                # Parse the email message
                message = email.message_from_bytes(data[0][1])

                # Print the subject and body of the email
                print("Subject: ", message["Subject"])
                print("Body: ", message.get_payload())
                tts = gTTS(text="The subject of the mail is :"+ message["Subject"], lang='en')
                typ, msg_data = mail.search(None, f'(SUBJECT "{message["Subject"]}")')
                if msg_data[0]:
                    email_id = msg_data[0].split()[-1]
                    typ, msg_data = mail.fetch(email_id, '(RFC822)')
                    raw_email = msg_data[0][1]
                    email_message = email.message_from_bytes(raw_email)
                    body = ''
                    if email_message.is_multipart():
                        for part in email_message.walk():
                            if part.get_content_type() == 'text/plain':
                                body += part.get_payload()
                    else:
                        body = email_message.get_payload()
                    # Use pyttsx3 to read the email body out loud
                    engine = pyttsx3.init()
                    engine.say(body)
                    engine.runAndWait()
            mail.close()
            mail.logout()

        elif text =='exit' or text == 'e x i t' or text == 'f o u r' or text == 'four':
            print("the user is exit from the mail successfully")
            talk("the user is exit from the mail successfully")
    else:
        print("Login failed. Please try again.")
        talk("login failed. please try again")

# Main program
login()