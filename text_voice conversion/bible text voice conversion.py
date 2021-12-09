#!/usr/bin/env python3
# NOTE: this example requires PyAudio because it uses the Microphone class

from gtts import gTTS

def japan(mean, text):

    blabla = (mean)
    tts = gTTS(text=blabla, lang='ja')
    tts.save('c:\data automation\ '+text+'.mp3')
    return

japan('アパート', '아파트')
japan('あぶない', '위험하다')
japan('あまり', '나머지')
japan('あらう', '씻다')
japan('あるく', '걷다')
japan('いい', '좋다')
japan('いう', '말하다')



