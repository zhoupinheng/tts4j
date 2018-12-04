# tts4j
tts4j主要是为了降低Java使用tts的难度。它通过jacob调用windows本身的文字阅读服务发声。

目前只使用在windows平台，并且是windows7 以后的版本才能正确使用。

使用方法:
TTS4JUtil.speak(volume,rate,words);

volume:音量取值范围[0-100]

rate:语速:[-10 - +10]

words:需要朗读的字符串


----

To build:

mvn compile

create installation:

mvn install

To cleanup:

mvn clean