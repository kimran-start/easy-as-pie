import speech_recognition as sr  # 음성 인식을 위한 라이브러리
from googletrans import Translator  # 번역 기능을 위한 라이브러리
import win32com.client  # 파워포인트와 상호작용을 위한 라이브러리

# 음성 인식기 초기화
recognizer = sr.Recognizer()
microphone = sr.Microphone()

# 번역기 초기화
translator = Translator()

# 파워포인트 애플리케이션에 연결
powerpoint = win32com.client.Dispatch("PowerPoint.Application")

# 슬라이드 쇼가 실행 중인지 확인
if powerpoint.SlideShowWindows.Count > 0:
    presentation = powerpoint.ActivePresentation
    slide = powerpoint.SlideShowWindows[1].View.Slide  # 첫 번째 슬라이드 쇼 윈도우 사용

    # 자막을 표시할 텍스트 상자 생성 (위치와 크기 조정 가능)
    text_box = slide.Shapes.AddTextbox(1, 100, 100, 400, 100)
    text_box.TextFrame.TextRange.Text = "자막이 여기에 표시됩니다."

    try:
        while True:
            with microphone as source:
                print("음성 입력을 듣는 중...")
                recognizer.adjust_for_ambient_noise(source)  # 주변 소음 조정
                audio = recognizer.listen(source)
                try:
                    # 음성을 한국어로 인식
                    spoken_text = recognizer.recognize_google(audio, language="ko-KR")
                    print(f"인식된 텍스트: {spoken_text}")

                    # 인식된 텍스트를 영어로 번역
                    translated_text = translator.translate(spoken_text, src='ko', dest='en').text
                    print(f"번역된 텍스트: {translated_text}")

                    # 파워포인트의 텍스트 상자에 번역된 텍스트 표시
                    text_box.TextFrame.TextRange.Text = translated_text

                except sr.UnknownValueError:
                    print("음성을 이해할 수 없습니다.")
                except sr.RequestError:
                    print("음성 인식 서비스에 요청할 수 없습니다.")
    except KeyboardInterrupt:
        print("음성 입력을 중지합니다.")
else:
    print("슬라이드 쇼가 실행되고 있지 않습니다. 슬라이드 쇼를 시작한 후 다시 실행하세요.")
