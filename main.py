from flask import Flask, request, send_file
import win32com.client
import pythoncom
import pyttsx3


app = Flask(__name__)


@app.before_request
def initialize():
    if not hasattr(pythoncom, 'CoInitialized'):
        pythoncom.CoInitialize()


@app.route('/convert', methods=['POST'])
def convert_text_to_mp3():
    # 获取前端发送的文本信息
    text = request.form['text']

    # 调用文本转换函数，生成MP3文件
    output_file = text_to_speech(text)

    # 返回生成的MP3文件
    response = send_file(output_file, as_attachment=True)
    response.headers['Access-Control-Allow-Origin'] = '*'
    return response


def text_to_speech(text):
    # 创建一个语音合成对象
    engine = pyttsx3.init()

    # 设置语音属性
    voices = engine.getProperty('voices')
    engine.setProperty('voice', voices[0].id)  # 选择第一个语音
    engine.setProperty('rate', 100)  # 设置语速，范围为0-200
    engine.setProperty('volume', 1.0)  # 设置音量，范围为0.0-1.0
    output_file = "output.mp3"

    # 将文本转换成语音
    engine.save_to_file(text, output_file)
    engine.runAndWait()

    # 创建一个SAPI对象
    speaker = win32com.client.Dispatch("SAPI.SpVoice")

    # 设置语音属性
    speaker.Voice = speaker.GetVoices().Item(0)  # 选择第一个语音
    speaker.Rate = -2  # 设置语速，0为正常速度
    speaker.Volume = 100  # 设置音量，范围0-100

    # 将文本转换为语音
    #speaker.Speak(text)
    print("转换完成！")

    return output_file


if __name__ == '__main__':
    app.run(debug=True)
