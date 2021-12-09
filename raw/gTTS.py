from gtts import gTTS
from tempfile import TemporaryFile

tts = gTTS(text='如果你们中间有谁缺乏智慧，他就应该向上帝请求。上帝不求全责备，对所有的人都慷慨施予，因此，上帝是会赐给他智慧的。', lang='zh')
tts.save("ch1_5.mp3")