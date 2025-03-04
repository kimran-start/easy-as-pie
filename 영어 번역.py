import speech_recognition as sr  # 음성 인식을 위한 라이브러리
import openai  # OpenAI API를 위한 라이브러리
import win32com.client  # 파워포인트와 상호작용을 위한 라이브러리
import threading  # 멀티스레딩을 위한 라이브러리
import queue  # 스레드 간 데이터 전송을 위한 큐

# OpenAI API 키 설정
openai.api_key = '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'  # 여기에 자신의 OpenAI API 키를 입력하세요.

# 음성 인식기 초기화
recognizer = sr.Recognizer()
microphone = sr.Microphone()

# 파워포인트 애플리케이션에 연결
powerpoint = win32com.client.Dispatch("PowerPoint.Application")

# 슬라이드 쇼가 실행 중인지 확인
if powerpoint.SlideShowWindows.Count > 0:
    presentation = powerpoint.ActivePresentation
    slide = powerpoint.SlideShowWindows[1].View.Slide  # 첫 번째 슬라이드 쇼 윈도우 사용

    # 자막을 표시할 텍스트 상자 생성 (가로 길이 늘리기)
    text_box = slide.Shapes.AddTextbox(1, 100, 700, 800, 200)  # 너비를 800으로 설정
    text_box.TextFrame.TextRange.Text = "자막이 여기에 표시됩니다."
    text_box.TextFrame.TextRange.Font.Size = 32  # 글씨 크기 설정

    last_slide_index = slide.SlideIndex  # 현재 슬라이드 인덱스 저장

    # 큐 생성 (음성 인식 결과를 메인 스레드로 전달하기 위해)
    translation_queue = queue.Queue()

    def listen_and_translate():
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
                    response = openai.ChatCompletion.create(
                        model="gpt-3.5-turbo",
                        messages=[
                            {"role": "user", "content": f"Translate this to English: {spoken_text}"}
                        ]
                    )
                    translated_text = response['choices'][0]['message']['content'].strip()
                    print(f"번역된 텍스트: {translated_text}")

                    # 번역된 텍스트를 큐에 추가
                    translation_queue.put(translated_text)

                except sr.UnknownValueError:
                    print("음성을 이해할 수 없습니다.")
                except sr.RequestError:
                    print("음성 인식 서비스에 요청할 수 없습니다.")
                except Exception as e:
                    print(f"오류 발생: {e}")

    # 음성 인식 및 번역 스레드 시작
    thread = threading.Thread(target=listen_and_translate, daemon=True)
    thread.start()

    try:
        while True:
            # 큐에서 번역된 텍스트 가져오기
            if not translation_queue.empty():
                translated_text = translation_queue.get()

                # 현재 슬라이드가 변경되었는지 확인
                current_slide_index = powerpoint.SlideShowWindows[1].View.Slide.SlideIndex
                if current_slide_index != last_slide_index:
                    # 현재 슬라이드로 업데이트
                    slide = powerpoint.SlideShowWindows[1].View.Slide
                    text_box = slide.Shapes.AddTextbox(1, 100, 700, 800, 200)  # 새로운 텍스트 상자 생성
                    text_box.TextFrame.TextRange.Font.Size = 32  # 글씨 크기 설정
                    last_slide_index = current_slide_index  # 슬라이드 인덱스 업데이트

                # 파워포인트의 텍스트 상자에 번역된 텍스트 표시
                text_box.TextFrame.TextRange.Text = translated_text

    except KeyboardInterrupt:
        print("음성 입력을 중지합니다.")

else:
    print("슬라이드 쇼가 실행되고 있지 않습니다. 슬라이드 쇼를 시작한 후 다시 실행하세요.")
