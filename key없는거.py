import speech_recognition as sr  # 음성 인식 라이브러리 임포트
import openai  # OpenAI API 라이브러리 임포트
import os  # 파일 작업을 위한 os 모듈 임포트

# OpenAI API 키 설정
openai.api_key = "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"  # 자신의 OpenAI API 키 입력

# 음성 인식 초기화
recognizer = sr.Recognizer()

# 마이크로부터 음성 인식
with sr.Microphone() as source:  # 마이크를 소스로 설정
    print("음성을 말하세요...")  # 사용자에게 음성을 말하라고 안내
    audio = recognizer.listen(source)  # 음성을 듣고 audio 변수에 저장

    try:
        # 음성을 텍스트로 변환
        text = recognizer.recognize_google(audio, language='ko-KR')  # 한국어로 인식
        print(f"인식된 텍스트: {text}")  # 인식된 텍스트 출력

        # ChatGPT를 통해 텍스트를 영어로 번역
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "user", "content": f"Translate the following text to English: {text}"}
            ]
        )
        translated_text = response['choices'][0]['message']['content'].strip()  # 번역된 텍스트 추출
        print(f"번역된 텍스트: {translated_text}")  # 번역된 텍스트 출력

        # 결과를 메모장에 저장
        with open("translated_text.txt", "w", encoding="utf-8") as f:  # 파일 열기
            f.write(translated_text)  # 번역된 텍스트 파일에 쓰기
        
        # 메모장 열기
        os.startfile("translated_text.txt")  # 메모장 파일 열기

    except sr.UnknownValueError:
        print("음성을 인식할 수 없습니다.")  # 음성이 인식되지 않을 때 메시지 출력
    except sr.RequestError as e:
        print(f"음성 인식 서비스에 접근할 수 없습니다; {e}")  # 서비스 오류 메시지 출력
    except Exception as e:
        print(f"오류 발생: {e}")  # 기타 오류 메시지 출력
