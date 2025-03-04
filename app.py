import speech_recognition as sr
import openai
import win32com.client
import threading
import queue
import pythoncom
import time
from flask import Flask, render_template, jsonify

app = Flask(__name__)

# OpenAI API 키 설정
openai.api_key = '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'

# 음성 인식기 초기화
recognizer = sr.Recognizer()
running = False
powerpoint = None
translation_thread = None
stop_event = threading.Event()  # 중지를 신호하기 위한 스레딩 이벤트 추가

def translate_text(text):
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[
            {"role": "user", "content": f"Translate this to English: {text}"}
        ]
    )
    return response['choices'][0]['message']['content'].strip()

def listen_and_translate():
    global running, powerpoint
    pythoncom.CoInitialize()  # COM 초기화
    
    try:
        # 스레드 내에서 마이크 인스턴스 생성
        microphone = sr.Microphone()
        
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")

        if powerpoint.SlideShowWindows.Count > 0:
            presentation = powerpoint.ActivePresentation
            slide = powerpoint.SlideShowWindows[1].View.Slide
            text_box = slide.Shapes.AddTextbox(1, 100, 700, 800, 200)
            text_box.TextFrame.TextRange.Text = "자막이 여기에 표시됩니다."
            text_box.TextFrame.TextRange.Font.Size = 32
            last_slide_index = slide.SlideIndex

            # 중지 체크를 더 빠르게 하기 위해 짧은 타임아웃 설정
            recognizer.pause_threshold = 0.5
            recognizer.operation_timeout = 3  # 3초 타임아
            
            while running and not stop_event.is_set():
                try:
                    with microphone as source:
                        print("Listening for speech...")
                        # 짧은 조정 기간 설
                        recognizer.adjust_for_ambient_noise(source, duration=0.5)
                        
                        # 듣기 시작하기 전에 중지해야 하는지 확인
                        if not running or stop_event.is_set():
                            break
                            
                        # 중지 이벤트를 주기적으로 확인하기 위해 타임아웃 사용
                        try:
                            audio = recognizer.listen(source, timeout=2, phrase_time_limit=10)
                        except sr.WaitTimeoutError:
                            # 타임아웃이 발생하면 중지해야 하는지 확인
                            continue
                        
                        # 오디오 처리 전에 다시 한 번 중지 여부 확인
                        if not running or stop_event.is_set():
                            break
                            
                        try:
                            spoken_text = recognizer.recognize_google(audio, language="ko-KR")
                            print(f"Recognized text: {spoken_text}")

                            # 번역하기 전에 다시 한 번 중지 여부 확인
                            if not running or stop_event.is_set():
                                break
                                
                            translated_text = translate_text(spoken_text)
                            print(f"Translated text: {translated_text}")

                            # 파워포인트가 여전히 실행 중인지 확인하고 중지 요청이 없을 경우
                            if powerpoint.SlideShowWindows.Count > 0 and running and not stop_event.is_set():
                                current_slide_index = powerpoint.SlideShowWindows[1].View.Slide.SlideIndex
                                if current_slide_index != last_slide_index:
                                    slide = powerpoint.SlideShowWindows[1].View.Slide
                                    text_box = slide.Shapes.AddTextbox(1, 100, 700, 800, 200)
                                    text_box.TextFrame.TextRange.Font.Size = 32
                                    last_slide_index = current_slide_index

                                text_box.TextFrame.TextRange.Text = translated_text
                            else:
                                print("슬라이드 쇼가 닫혔거나 중지 요청 되었습니다.")
                                running = False
                                break

                        except sr.UnknownValueError:
                            print("음성을 이해할 수 없습니다.")
                        except sr.RequestError:
                            print("음성 인식 서비스에서 결과를 요청할 수 없습니다.")
                except Exception as e:
                    print(f"마이크에서 오류: {e}")
                    # 반복적인 오류를 방지하기 위해 짧은 일시 정지
                    time.sleep(0.5)
                    
                # 각 반복 후 중지 조건 확인
                if not running or stop_event.is_set():
                    break
        else:
            print("실행중인 슬라이드 쇼가 없습니다.")
            running = False
    except Exception as e:
        print(f"번역 스레드에서 오류: {e}")
    finally:
        running = False
        try:
            if powerpoint:
                # 가능하면 텍스트 박스를 정리하려고 시도
                if powerpoint.SlideShowWindows.Count > 0:
                    try:
                        text_box.TextFrame.TextRange.Text = ""
                    except:
                        pass
        except:
            pass
        powerpoint = None
        pythoncom.CoUninitialize()  # COM 해제
        print("번역 스레드 종료")

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/start')
def start():
    global running, translation_thread, stop_event
    
    # 중지 이벤트 초기화
    stop_event.clear()
    
    if not running:
        running = True
        translation_thread = threading.Thread(target=listen_and_translate, daemon=True)
        translation_thread.start()
        return jsonify({"status": "started"})
    return jsonify({"status": "already_running"})

@app.route('/stop')
def stop():
    global running, stop_event
    
    # 모든 스레드에 중지 신호 보내기
    running = False
    stop_event.set()
    
    # 메시지가 전송되도록 잠시 대기
    time.sleep(0.5)
    
    return jsonify({"status": "stopped"})

@app.route('/status')
def status():
    return jsonify({
        "running": running,
        "thread_alive": translation_thread.is_alive() if translation_thread else False
    })

if __name__ == '__main__':
    app.run(debug=False)  # 프로덕션에서는 디버그를 False로 설정하여 스레드 문제 방지
