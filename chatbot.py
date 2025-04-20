import pyttsx3 as p  # Text to speech
import speech_recognition as sr  # Recognize my voice
import requests  # For making HTTP requests
import randfacts
from selenium import webdriver #automate web pages
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from google.oauth2.credentials import Credentials #calendar 
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os.path


# Initialize the text-to-speech engine
engine = p.init()
engine.setProperty('rate', 150)
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)

def speak(text):
    engine.say(text)
    engine.runAndWait()

class Infow:
    def __init__(self):
        chrome_driver_path = r'C:\webdrivers\chromedriver-win64\chromedriver.exe'
        service = Service(chrome_driver_path)
        chrome_options = Options()
        chrome_options.add_argument('--ignore-certificate-errors')
        chrome_options.add_argument('--allow-insecure-localhost')
        self.driver = webdriver.Chrome(service=service, options=chrome_options)

    def get_info(self, query):
        try:
            self.driver.get('https://www.wikipedia.org')
            search = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="searchInput"]'))
            )
            search.click()
            search.send_keys(query)

            enter = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="search-form"]/fieldset/button'))
            )
            enter.click()

            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//h1'))
            )
            
            print(self.driver.title)  # Print the title of the page
            speak(f"Here is the information I found on Wikipedia about {query}.")
        except Exception as e:
            print(f"An error occurred: {e}")
            speak("Sorry, there was an error while fetching the information.")

    def play_youtube_video(self, video_url):
        try:
            self.driver.get(video_url)  # Open the YouTube video URL
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="movie_player"]//button[@class="ytp-large-play-button ytp-button"]'))
            ).click()  # Click the play button
            speak("Playing the video.")
        except Exception as e:
            print(f"An error occurred: {e}")
            speak("Sorry, I couldn't play the video.")

    def close_driver(self):
        self.driver.quit()  # Close the browser when done

def fetch_news(api_key):
    url = f"https://newsapi.org/v2/top-headlines?country=us&apiKey={api_key}"
    try:
        response = requests.get(url)
        response.raise_for_status()  # Raise an error for bad responses
        news_data = response.json()
        articles = news_data.get('articles', [])
        if articles:
            news_summary = "Here are the top news headlines: "
            for article in articles[:5]:  # Limit to the top 5 articles
                news_summary += f"{article['title']}. "
            return news_summary
        else:
            return "No news articles found."
    except Exception as e:
        print(f"An error occurred while fetching news: {e}")
        return "Sorry, I couldn't fetch the news at the moment."

def fetch_joke():
    url = "https://official-joke-api.appspot.com/jokes/random"
    try:
        response = requests.get(url)
        response.raise_for_status()  # Raise an error for bad responses
        joke_data = response.json()
        joke = f"{joke_data['setup']} ... {joke_data['punchline']}"
        return joke
    except Exception as e:
        print(f"An error occurred while fetching a joke: {e}")
        return "Sorry, I couldn't fetch a joke at the moment."

def fetch_weather(api_key, city):
    url = f"http://api.openweathermap.org/data/2.5/weather?q={city}&appid={api_key}&units=metric"
    try:
        response = requests.get(url)
        response.raise_for_status()
        weather_data = response.json()
        temp = weather_data['main']['temp']
        weather_desc = weather_data['weather'][0]['description']
        return f"The current temperature in {city} is {temp}°C with {weather_desc}."
    except Exception as e:
        print(f"An error occurred while fetching weather: {e}")
        return "Sorry, I couldn't fetch the weather at the moment."

def fetch_stock_price(api_key, symbol):
    url = f"https://www.alphavantage.co/query?function=TIME_SERIES_INTRADAY&symbol={symbol}&interval=5min&apikey={api_key}"
    try:
        response = requests.get(url)
        response.raise_for_status()
        stock_data = response.json()
        latest_time = next(iter(stock_data['Time Series (5min)']))
        latest_data = stock_data['Time Series (5min)'][latest_time]
        return f"The current price of {symbol} is ${latest_data['1. open']}."
    except Exception as e:
        print(f"An error occurred while fetching stock price: {e}")
        return "Sorry, I couldn't fetch the stock price at the moment."

# Google Calendar Integration
SCOPES = ['https://www.googleapis.com/auth/calendar']

def authenticate_google_calendar():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def fetch_upcoming_events(creds):
    service = build('calendar', 'v3', credentials=creds)
    now = '2024-10-03T10:59:07Z'  
    events_result = service.events().list(calendarId='primary', timeMin=now,
                                          maxResults=5, singleEvents=True,
                                          orderBy='startTime').execute()
    events = events_result.get('items', [])
    
    if not events:
        return "No upcoming events found."
    
    upcoming_events = "Upcoming events: "
    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        upcoming_events += f"{start}: {event['summary']}. "
    
    return upcoming_events

# Main interaction with the assistant
if __name__ == "__main__":
    news_api_key = 'e4621b77ee849bcbac43dfeebe35a2aF'  
    weather_api_key = '3d07c874d15486de63eb36c003f80566'  
    stock_api_key = 'QX8DD04T94D8RFRH'  
    r = sr.Recognizer()  # Retrieve audio from mic
    speak("Hello sir, I am Mia, your voice assistant. How are you?")

    with sr.Microphone() as source:
        r.energy_threshold = 10000  # Decrease spectrum of voice
        r.adjust_for_ambient_noise(source, 1.2)
        print("Listening...")
        try:
            audio = r.listen(source)  # Listens for input in microphone
            text = r.recognize_google(audio)  # Send to Google API and transform into text
            print(text)
        except sr.UnknownValueError:
            speak("Sorry, I didn't catch that.")
            exit()
        except sr.RequestError as e:
            speak("Could not request results; check your network connection.")
            exit()

    # Check for greeting response
    if "what" in text.lower() and "about" in text.lower() and "you" in text.lower():
        speak("I am also having a good day sir.")

    speak("What can I do for you?")      

    with sr.Microphone() as source:
        r.energy_threshold = 10000  # Decrease spectrum of voice
        r.adjust_for_ambient_noise(source, 1.2)
        print("Listening...")
        try:
            audio = r.listen(source)
            text2 = r.recognize_google(audio).lower()  # Normalize to lowercase
            print(text2)
        except sr.UnknownValueError:
            speak("Sorry, I didn't catch that.")
            exit()
        except sr.RequestError as e:
            speak("Could not request results; check your network connection.")
            exit()

    # Check for weather request
    if "weather" in text2:
        speak("Which city do you want the weather for?")
        
        with sr.Microphone() as source:
            r.energy_threshold = 10000  # Decrease spectrum of voice
            r.adjust_for_ambient_noise(source, 1.2)
            print("Listening...")
            try:
                audio = r.listen(source)
                city = r.recognize_google(audio).lower()  # Normalize to lowercase
                weather_info = fetch_weather(weather_api_key, city)
                speak(weather_info)

            except sr.UnknownValueError:
                speak("Sorry, I didn't catch that.")
            except sr.RequestError as e:
                speak("Could not request results; check your network connection.")

    # Check for news request
    if "news" in text2:
        speak("Fetching the latest news.")
        news_summary = fetch_news(news_api_key)
        speak(news_summary)

    # Check for joke request
    if "joke" in text2:
        speak("Here's a joke for you.")
        joke = fetch_joke()
        speak(joke)

    # Check for information request
    if "information" in text2:
        speak("You need info related to what topic?")

        with sr.Microphone() as source:
            r.energy_threshold = 10000  # Decrease spectrum of voice
            r.adjust_for_ambient_noise(source, 1.2)
            print("Listening...")
            try:
                audio = r.listen(source)
                infor = r.recognize_google(audio).lower()  # Normalize to lowercase

                speak(f"Searching in Wikipedia for {infor}.")
                assist = Infow()
                assist.get_info(infor)

                # Keep the browser open until the user decides to close it
                speak("Press enter to close the browser.")
                input()  # Wait for user input
                assist.close_driver()  # Ensure the driver is closed after the operation

            except sr.UnknownValueError:
                speak("Sorry, I didn't catch that.")
            except sr.RequestError as e:
                speak("Could not request results; check your network connection.")

    # Check for stock price request
    if "stock" in text2:
        speak("What stock symbol do you want to check?")
        with sr.Microphone() as source:
            r.energy_threshold = 10000
            r.adjust_for_ambient_noise(source, 1.2)
            print("Listening...")
            try:
                audio = r.listen(source)
                stock_symbol = r.recognize_google(audio).upper()
                stock_info = fetch_stock_price(stock_api_key, stock_symbol)
                speak(stock_info)
            except sr.UnknownValueError:
                speak("Sorry, I didn't catch that.")
            except sr.RequestError as e:
                speak("Could not request results; check your network connection.")

    # Check for calendar requests
    if "calendar" in text2 or "events" in text2:
        speak("Let me fetch your upcoming calendar events.") 
        creds = authenticate_google_calendar()
        upcoming_events = fetch_upcoming_events(creds)
        speak(upcoming_events)

    # Random facts 
    if "facts" in text2 or "fact" in text2:
        x = randfacts.get_fact()
        print(x)
        speak("Did you know that," + x)           

    # Check for YouTube video request
    if "play" in text2 and "video" in text2:
        speak("Which video do you want to play?")

        with sr.Microphone() as source:
            r.energy_threshold = 10000  # Decrease spectrum of voice
            r.adjust_for_ambient_noise(source, 1.2)
            print("Listening...")
            try:
                audio = r.listen(source)
                video_title = r.recognize_google(audio).lower()  # Normalize to lowercase
                speak(f"Searching for {video_title} on YouTube.")

                # Create YouTube search URL
                search_url = f"https://www.youtube.com/results?search_query={video_title.replace(' ', '+')}"
                assist = Infow()
                assist.driver.get(search_url)

                # Wait for search results to load and click the first video
                first_video = WebDriverWait(assist.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="video-title"]'))
                )
                first_video.click()

                # Wait for the video page to load and play
                assist.play_youtube_video(assist.driver.current_url)
    
                # Keep the browser open until the user decides to close it
                speak("Press enter to close the browser.")
                input()  # Wait for user input
                assist.close_driver()
            except sr.UnknownValueError:
                speak("Sorry, I didn't catch that.")
            except sr.RequestError as e:
                speak("Could not request results; check your network connection.")