from win32com.client import Dispatch
import requests
import json
cat = ""


def readnews(n):
    try:
        x = int(input("Enter 0 to go to previous menu.\nEnter the corresponding news number to read it in more  detail."
                      "\nTo exit just press enter : "))
    except ValueError:
        x = -1
    if x == 0:
        start()
    elif 0 < x < 11:
        print()
        print(n['articles'][x - 1]['description'])
        print(n['articles'][x - 1]['content'])
        print("For complete news go to this address : ")
        print(n['articles'][x - 1]['url'])
        print()
        try:
            y = int(input("Enter 1 if you want this program to read the news out loud or just press enter to "
                          "continue : "))
        except ValueError:
            y = 0
        if y == 1:
            text_to_speak(n['articles'][x - 1]['description'])
            text_to_speak(n['articles'][x - 1]['content'])
            text_to_speak("For complete news go to this address : ")
            readnews(n)
        else:
            readnews(n)
    else:
        print("Exiting...")
        text_to_speak("Thanks for using our service have a great day")


def text_to_speak(st):
    speak = Dispatch("SAPI.SpVoice")
    speak.Speak(st)


def catsel():
    n = 0
    global cat
    print("Category:-\n"
          "1) Business\n"
          "2) Entertainment\n"
          "3) General\n"
          "4) Health\n"
          "5) Science\n"
          "6) Sports\n"
          "7) Technology")
    print("To select the category of the news enter the corresponding number")
    try:
        n = int(input(":"))
    except ValueError:
        pass
    if n == 1:
        cat = "business"
    elif n == 2:
        cat = "entertainment"
    elif n == 3:
        cat = "general"
    elif n == 4:
        cat = "health"
    elif n == 5:
        cat = "science"
    elif n == 6:
        cat = "sports"
    elif n == 7:
        cat = "technology"
    else:
        print("Invalid Option !!!!\n"
              "Enter valid option!!")
        catsel()


def start():
    global cat
    catsel()
    print("Loading.....")
    url = "https://newsapi.org/v2/top-headlines?country=in&category=" + cat + "&pageSize=10&" \
                                                                              "apiKey=19e7cb69474d4ef9a39619de1dec6e45"
    try:
        news = requests.get(url)
    except Exception as e:
        print("Please connect to internet snd try again!!")
        text_to_speak("No internet!!!")
        print("Press enter after connecting to internet")
        input()
        start()
    else:
        pnews = json.loads(news.text)
        for i in range(10):
            print(f"{i+1}) {pnews['articles'][i]['title']}")
            text_to_speak(f'News no {i+1}')
            text_to_speak(pnews['articles'][i]['title'])
        print("Done.....")
        readnews(pnews)


print("Welcome to top 10 news headline reader mady by shubham Jha!!!!!!!")
text_to_speak("Welcome to top 10 news headline reader by shubham Jha!!!!!!!")
start()
