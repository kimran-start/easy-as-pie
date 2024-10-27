import speech_recognition as sr
from googletrans import Translator
import os

# 음성 인식 초기화
recognizer = sr.Recognizer()
translator = Translator()

# 마이크로부터 음성 인식
with sr.Microphone() as source:
    print("음성을 말하세요...")
    audio = recognizer.listen(source)

    try:
        # 음성을 텍스트로 변환
        text = recognizer.recognize_google(audio, language='ko-KR')
        print(f"인식된 텍스트: {text}")

        # 텍스트를 영어로 번역
        translated = translator.translate(text, dest='en')
        print(f"번역된 텍스트: {translated.text}")

        # 결과를 메모장에 저장
        with open("translated_text.txt", "w", encoding="utf-8") as f:
            f.write(translated.text)
        
        # 메모장 열기
        os.startfile("translated_text.txt")

    except sr.UnknownValueError:
        print("음성을 인식할 수 없습니다.")
    except sr.RequestError as e:
        print(f"음성 인식 서비스에 접근할 수 없습니다; {e}")
