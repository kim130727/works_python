#!/usr/bin/env python3
# NOTE: this example requires PyAudio because it uses the Microphone class

import speech_recognition as sr
from pylab import*
import matplotlib.pyplot as plt
import numpy as np
import numpy
import wave
import sys
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from gtts import gTTS
import pyglet

def bible(sentence):

    blabla = (sentence)
    tts = gTTS(text=blabla, lang='en')
    tts.save('test.mp3')

    while True:
        song = pyglet.media.load('test.mp3')
        song.play()
        question1 = input("다시 들으시겠습니까? (y/n) ")
        if question1 == 'y':
            song = pyglet.media.load('test.mp3')
            song.play()
        else:
            break

    while True:
        # obtain audio from the microphone
        r = sr.Recognizer()
        with sr.Microphone() as source:
            print("음성 인식을 시작합니다.")
            audio = r.listen(source)

        # recognize speech using Google Speech Recognition
        try:
            # for testing purposes, we're just using the default API key
            # to use another API key, use `r.recognize_google(audio, key="GOOGLE_SPEECH_RECOGNITION_API_KEY")`
            # instead of `r.recognize_google(audio)`
            print("인식 결과")
            print(r.recognize_google(audio))

        except sr.UnknownValueError:
            print("Speech Recognition could not understand audio")

        except sr.RequestError as e:
            print("Could not request results from Speech Recognition service; {0}".format(e))

        # write audio to a WAV file
        with open("microphone-results.wav", "wb") as f:
            f.write(audio.get_wav_data())

        spf = wave.open('microphone-results.wav','r')

        #Extract Raw Audio from Wav File
        signal = spf.readframes(-1)
        signal = np.fromstring(signal, 'Int16')
        fs = spf.getframerate()

        #If Stereo
        if spf.getnchannels() == 2:
            print ('Just mono files')
            sys.exit(0)

        Time=np.linspace(0, len(signal)/fs, num=len(signal))

        plt.figure(1)
        plt.title('Signal Wave...')
        plt.plot(Time,signal)
        plt.show()

        documents = (r.recognize_google(audio),sentence)
        tfidf_vectorizer = TfidfVectorizer()
        tfidf_matrix = tfidf_vectorizer.fit_transform(documents)
        cosine = cosine_similarity(tfidf_matrix[0:1], tfidf_matrix)
        cosine = cosine.astype(numpy.float32)
        print(round(cosine[0][1]*100,1))

        question2 = input("다시 말씀하시겠습니까? (y/n) ")

        if question2 == 'y':
            print("")
        else:
            break

bible('In the beginning was the Word')