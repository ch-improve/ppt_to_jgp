import os
import pythoncom
import requests
from flask import Flask, request, jsonify
import win32com.client

from ppt_to_jgp.settings import host, port

app = Flask(__name__)


def func(input_name, out_dirname):
    powerpoint = win32com.client.Dispatch("Powerpoint.Application")
    powerpoint.Visible = 1
    deck = powerpoint.Presentations.Open(input_name)
    path = os.path.dirname(input_name)
    out_filename = os.path.join(out_dirname, input_name.lstrip(path)).rsplit('.')[0]
    print(">>>>>>>", out_filename)
    deck.SaveAs(out_filename + '.jpg', 17)
    deck.Close()
    powerpoint.Quit()
    for dirpath, dirnames, filenames in os.walk(out_filename):
        print(filenames)
        return filenames


@app.route("/", methods=['post'])
def ppt_to_jpg():
    pythoncom.CoInitialize()
    resp = request.json
    url: str = resp.get('ppt_url')
    code = resp.get('code')
    if not all([url, code]):
        return jsonify(errno='1', errmsg='传入参数错误')
    try:
        ppt_file = requests.get(url).content
    except Exception as e:
        return jsonify(errno='2', errmsg=e)
    file_name = url.rstrip('/')
    try:
        file_name = "-".join(file_name.split('/')[-2:])
    except Exception as e:
        return jsonify(errno='3', errmsg='ppt_url错误')
    dir_name = os.path.dirname(os.path.abspath(__file__))
    input_name = os.path.join(dir_name, "ppt_dir")
    input_name = os.path.join(input_name, file_name)
    with open(input_name, 'wb') as f:
        f.write(ppt_file)
    try:
        file_name_list = func(input_name, os.path.join(dir_name, 'static'))
        n = len(file_name_list)
    except Exception as e:
        return jsonify(errno='4', errmsg='ppt转jgp出错')
    for i in range(n):
        file_name_list[i] = "http://{}:{}/static/".format(host, port) + file_name.split('.')[0] + '/' + file_name_list[i]
    data = {
        "code": code,
        "urls": file_name_list
    }
    print(data)
    return jsonify(data=data, errno='0', errmsg='ok')


if __name__ == '__main__':
    app.run(host=host, port=port, debug=True)


