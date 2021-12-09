from gtts import gTTS

tts = gTTS(text='hello', lang='en')

tts.save('hello.mp3')