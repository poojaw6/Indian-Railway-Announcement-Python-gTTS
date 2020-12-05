"""
We imported pyaudio, pydub and gTTS(google text to speech) modules using pip install
"""
import os
import pandas as pd
# pydub to manipulate audio
from pydub import AudioSegment
from gtts import gTTS


def textToSpeech(text, filename):
    # converts text to mp3 file
    myText = str(text)
    language = 'en'
    myObj = gTTS(text=myText, lang=language, slow=False)
    myObj.save(filename)


def mergeAudios(audios):
    # returns pydub AudioSegment
    # first create an emtpy audio
    combined = AudioSegment.empty()
    for audio in audios:
        combined += AudioSegment.from_mp3(audio)
    return combined


def generateSkeleton():
    # will make pieces of audio and stitch them together

    # import main audio file
    audio = AudioSegment.from_mp3('railway.mp3')
    # 1: May I have your attention please
    start = 64500
    finish = 70000
    audioProcessed = audio[start:finish]
    audioProcessed.export("1_eng.mp3", format="mp3")

    # 2: Train number and name
    # 3: From
    start = 76400
    finish = 77200
    audioProcessed = audio[start:finish]
    audioProcessed.export("3_eng.mp3", format="mp3")
    # 4: From city
    # 5: To
    start = 78000
    finish = 78800
    audioProcessed = audio[start:finish]
    audioProcessed.export("5_eng.mp3", format="mp3")
    # 6: To city
    # 7: via
    start = 80000
    finish = 81000
    audioProcessed = audio[start:finish]
    audioProcessed.export("7_eng.mp3", format="mp3")
    # 8: via route
    # 9: is arriving shortly on platform num
    start = 82700
    finish = 86700
    audioProcessed = audio[start:finish]
    audioProcessed.export("9_eng.mp3", format="mp3")
    # 10: platform num


def generateAnnouncement(filename):
    # reads excel file for train details
    df = pd.read_excel(filename)
    print(df)
    for index, item in df.iterrows():
        # Generate 2: Train number and name
        textToSpeech(item['train_no'] + " " + item['train_name'], "2_eng.mp3")
        # Generate 4: From city
        textToSpeech(item['from'], "4_eng.mp3")
        # Generate 6: To city
        textToSpeech(item['to'], "6_eng.mp3")
        # Generate 8: via route
        textToSpeech(item['via'], "8_eng.mp3")
        # Generate 10: platform num
        textToSpeech(item['platform'], "10_eng.mp3")

        # Concatenate all generated audios
        audios = [f"{i}_eng.mp3" for i in range(1, 11)]

        announcement = mergeAudios(audios)
        announcement.export(f"announcement_for_{item['train_no']}.mp3", format="mp3")


if __name__ == '__main__':
    print("Generating Skeleton...")
    generateSkeleton()
    print("Now Generating Announcement...")
    generateAnnouncement('announce_english.xlsx')
