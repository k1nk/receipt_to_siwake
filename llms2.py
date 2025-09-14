# Copyright (c) 2025 Kenichi Nakatani
# https://note.com/kenichi_nakatani/n/n82d474e5c4d1
# Released under the MIT license
# https://opensource.org/licenses/mit-license.php

import os
import tkinter as tk
from tkinter import filedialog,messagebox
import json
import time
import mimetypes


#Vertex AI
#from google.cloud import aiplatform

from google.genai import Client, types
#from google.genai.types import Part
PROJECT_ID = os.environ["GOOGLE_CLOUD_PROJECT_ID"]
GOOGLE_CLOUD_LOCATION = os.environ["GOOGLE_CLOUD_LOCATION"]

#import vertexai
#from vertexai.generative_models import GenerationConfig, GenerativeModel, Part
#PROJECT_ID = os.environ["GOOGLE_CLOUD_PROJECT_ID"]
#vertexai.init(project=PROJECT_ID)

#問い合わせの待ち時間
QUERY_WAIT_SEC = 6
global ACCEPTED_EXTENSIONS
ACCEPTED_EXTENSIONS = ['.png','.jpg','.jpeg','.webp','.pdf','.heic']

def adjust_type(json_res):

    #PDFファイルの中に複数の検索対象となる文書等が含まれている場合、辞書を含むリストとして結果が返ってくることがある
    if type(json_res) is dict:
        pdf_info=[]
        pdf_info.append(json_res)
        return pdf_info

    if type(json_res) is list:
        return json_res

    # unexpected type
    return []

def get_pdf_json_response_gemini(pdf_path,prompt,response_schema,model_id = "gemini-2.5-flash"):
    return get_image_json_response_gemini(pdf_path,prompt,response_schema,model_id)

def get_image_json_response_gemini(image_path,prompt,response_schema,model_id = "gemini-2.5-flash"):
    """
    Gemini（Vertex AI API経由）を使用して、
    prompt文字列により、pdf_pathにある書類の情報を取得する。

    Parameters
    ----------
    pdf_path : str
        検索対象のPDFファイルのパス（文字列）
    prompt : str
        プロンプト（文字列）
    response_schema : object
        JSONレスポンスの形式

    Returns
    -------
    res : list[dict[str,str]]　書類の内容

    その他の場合
    res : None

    Memo
    -------
    ファイル全体で１回の問い合わせをおこなって、結果を取得している。


    """

    system_prompt = prompt[0]
    user_prompt = prompt[1]

    #model_id = "gemini-2.0-flash-001"
    #model_id = "gemini-2.5-flash"

    #model = GenerativeModel(model_id,
    #        system_instruction=system_prompt,
    #        generation_config={"response_mime_type": "application/json"})
    
    #model = GenerativeModel(model_id,
    #        system_instruction=system_prompt)

    client = Client(
        vertexai=True,
        project=PROJECT_ID
    )
    #client = Client(
    #    vertexai=True,
    #    project=PROJECT_ID,
    #    location=GOOGLE_CLOUD_LOCATION
    #)

    #model = client.get_model(model_id)


    # ファイルの拡張子を取得

    file_extension = os.path.splitext(image_path)[1].lower()
    if not file_extension in ACCEPTED_EXTENSIONS:
        print(f"{image_path} not ends with accepted extensions.")
        return None
    mime_type = mimetypes.types_map.get(file_extension, "image/jpeg")
    
    with open(image_path, 'rb') as image_file:
        image_data = image_file.read()
        image_part = types.Part.from_bytes(data=image_data,mime_type=mime_type)
        #image_part = Part.from_data(image_data, mime_type)

    try:
        #receipt
        #response = model.generate_content([user_prompt, sample_pdf])
        #print(response.text)
    
        #response = model.generate_content(
        #    [user_prompt, image_part],
        #    generation_config=GenerationConfig(
        #        response_mime_type="application/json",
        #        response_schema=response_schema
        #    )
        #)

        response = client.models.generate_content(
            model = model_id,
            contents = [image_part,user_prompt],
            config = types.GenerateContentConfig(
                system_instruction = system_prompt,
                response_mime_type="application/json",
                response_schema=response_schema
            ),
        )

    except Exception as e:
        print("Error:",e)
        return None

    res = json.loads(response.text)
    #pdf_info = adjust_type(res)

    return res

def get_pdf_info(pdf_path,prompt,response_schema,model_id = "gemini-2.5-flash"):
    return get_image_info(pdf_path,prompt,response_schema,model_id)

def get_image_info(image_path,prompt,response_schema,model_id = "gemini-2.5-flash"):
    """
    promptを使って、image_pathにある画像ファイルの内容を調べる
    
    Parameters
    ----------
    image_path : str
        検索対象の画像ファイルのパス（文字列）
    prompt : str
        プロンプト（文字列）
    response_schema : object
        JSONレスポンスの形式

    Returns
    -------
    res : list[dict[str,str]]
            　　エラーの場合は、None

    Memo
    -------
    認識する対象の単位ごとにPDFファイルを分けておくことを推奨する。

    １つのPDFファイルの中に複数の認識する対象（領収書など）がある場合は、
    対象の情報（辞書型）を要素が複数含まれるリストが返されることがあるが、
    使用するモデルによって挙動がことなることがあるので、認識する対象の単位ごとに
    画像ファイルを分けておくことが望ましい。
    """
    
    #Gemini(Vertex AI)
    image_info = get_image_json_response_gemini(image_path,prompt,response_schema,model_id)

    #Other LLMS
    #todo

    return image_info
def get_pdf_info_in(folder_path,prompt,response_schema,wait_time=1,model_id = "gemini-2.5-flash"):
    return get_image_info_in(folder_path,prompt,response_schema,wait_time,model_id)

def get_image_info_in(folder_path,prompt,response_schema,wait_time=1,model_id = "gemini-2.5-flash"):
    """
    promptを使って、フォルダ（folder_path）内のPDFファイルの内容を調べる
    拡張子が下記のものでないファイルやファイル名の先頭が「-」であるファイルは調べない。
    対象となる画像ファイルの拡張子はACCEPTED_EXTENSIONS
    ['.png','.jpg','.jpeg','.webp','.pdf','.heic']

    Parameters
    ----------
    folder_path : str
        検索対象のPDFファイルが入っているフォルダのパス（文字列）
    prompt : str
        プロンプト（文字列）
    response_schema : object
        JSONレスポンスの形式
    wait_time : int , default 1
        １つのファイルを調べてから、次のファイルを調べるまでの待ち時間（秒）

    Returns
    -------
    info : list[dict[str,str]]　書類の内容
            　　エラーの場合は、None

    Memo
    -------
    認識する対象の単位ごとにPDFファイルを分けておくことを推奨する。

    １つの画像ファイルの中に複数の認識する対象（領収書など）がある場合は、
    対象の情報（辞書型）を要素が複数含まれるリストが返されることがあるが、
    使用するモデルによって挙動がことなることがあるので、認識する対象の単位ごとに
    画像ファイルを分けておくことが望ましい。

    """

    info=[]
    for pathname, dirnames, filenames in os.walk(folder_path):
        for filename in filenames:
            if  filename.startswith('-'):
                print(f"{filename} starts with minus.")
                continue
            file_extension = os.path.splitext(filename)[1].lower()
            if not file_extension in ACCEPTED_EXTENSIONS:
                print(f"{filename} not ends with accepted extensions.")
                continue
            image_info = get_image_info(os.path.join(pathname,filename),prompt,response_schema,model_id)
            if image_info is None:
                continue
            info.extend(image_info)
            time.sleep(wait_time)

    return info

def get_user_prompt(prompt_info):
    """
    文書を調べて結果をJSON形式で取得するための、ユーザプロンプトを作成する
    
    Parameters
    ----------
    prompt_info (dict[str, str | dict[str, str]]): ユーザー情報を格納した辞書。
        - "doc_type": 読み込む文書の種類（文字列）
        - "item_type": 読み込む文書内の項目の種類（文字列）、オプション
        - "contents": 調べる内容（文字列）を値としたリスト
        - "comment": プロンプトに付加するコメント（文字列）

    Returns
    -------
    user_prompt : str
        ユーザプロンプト
    
    Examples
    -----
    >>>prompt_info={doc_type:'領収書',
        'contents':['支払日','支払の相手先','合計金額','支払内容'],
        'comment':"合計金額は税込みの金額でお願いします。支払内容が複数ある場合は空白で区切って連結して下さい。支払日は[西暦]/[月]/[日]という形式で出力して下さい。"}

    >>>get_user_prompt(prompt_info)

    この画像が領収書である場合、領収書の支払日と支払の相手先、合計金額、支払いの内容を教えて下さい。
    合計金額は税込みの金額でお願いします。
    支払内容が複数ある場合は空白で区切って連結して下さい。
    支払日は[西暦]/[月]/[日]という形式で出力して下さい。
    """

    user_prompt = f"この画像が{prompt_info['doc_type']}である場合、{prompt_info['doc_type']}に記載されている"

    if prompt_info.get('item_type',None) is not None:
        user_prompt += f"それぞれの{prompt_info['item_type']}について、"

    first_item = True
    for value in prompt_info['contents']:
        if not first_item :
            user_prompt += '、'
        user_prompt += value
        first_item = False

    user_prompt += "の内容を教えて下さい。"
    
    user_prompt += prompt_info['comment']

    #user_prompt += prompt_info['comment']
    #uer_prompt += f"この画像が{prompt_info['doc_type']}でない場合は、空欄のJSON形式を返してください。"

    #prompt = system_prompt + user_prompt

    return user_prompt


    #prompt_info={doc_type:'領収書',
    #    'kv':{'date':'支払日','payee':'支払の相手先','amount':合計金額,'contents':'支払内容'},
    #    'must_filled_key':'amount',
    #    'comment':"合計金額は税込みの金額でお願いします。支払内容が複数ある場合は空白で区切って連結して下さい。支払日は[西暦]/[月]/[日]という形式で出力して下さい。"}

    # この画像が領収書である場合、領収書の支払日と支払の相手先、合計金額、支払いの内容を教えて下さい。
    # JSONのキーはdate, payee, amount, contentsとしてください。
    # dateは支払日、payeeは支払の相手先、amountは合計金額、contentsは支払内容とします。
    # 合計金額は税込みの金額でお願いします。
    # 支払内容が複数ある場合は空白で区切って連結して下さい。
    # 支払日は[西暦]/[月]/[日]という形式で出力して下さい。
    # この画像が領収書でない場合は、空欄のJSON形式を返してください。

def main():

    #検索対象のフォルダを指定
    folder_path = filedialog.askdirectory()
    print(folder_path)

    #システムプロンプトを指定
    system_prompt = """
    あなたは会計事務所の職員です。
    書類の内容の確認を手伝ってください。
    JSON形式で日本語で返答してください。
    """
    print("system_prompt",system_prompt)

    #ユーザプロンプトを指定

    #保険の控除証明書を読み取るサンプル
    comment_str = ("年間の保険料が数字でない場合は、「0」としてください。"
                )

    prompt_info={'doc_type':'保険の控除証明書',
                'item_type':'保険のカテゴリー',
        'contents':['保険会社の名前',
            '保険契約の分類',
            '保険の種類',
            '保険料',
            '配当金',
            '年間の申告額'],
        'comment':comment_str}

    user_prompt = get_user_prompt(prompt_info)
    print("user_prompt",user_prompt)

    #プロンプトを指定
    prompt =[system_prompt,user_prompt]

    response_schema = {
        "type": "array",
        "items": {
            "type": "object",
            "properties": {
                "company_name": {
                   "type": "string",
                   "description":"保険会社の名前",
                },
                "contract": {
                   "type": "string",
                   "description":"保険契約の分類",
                   "enum": [
                       "新生命保険契約",
                       "旧生命保険契約",
                   ],
                },
                "category": {
                   "type": "string",
                   "description":"保険の種類",
                   "enum": [
                       "一般",
                       "個人年金",
                       "介護医療",
                   ],
                },
                "insurance_fee_year": {
                    "type": "array",
                    "description":"保険料",
                    "items": {
                        "type": "object",
                        "properties": {
                            "amount": {
                                "type": "number",
                                "description":"金額",
                            },
                            "unit": {
                                "type": "string",
                                "description":"単位",
                                "enum": [
                                    "円",
                                    "ドル",
                                    "ユーロ",
                                ],
                            },
                        },
                        "required": ["amount"],
                    },
                },
                "dividend": {
                    "type": "number",
                    "description":"配当金",
                },
                "amount_year": {
                    "type": "number",
                    "description":"年間の申告額",
                },                
            },
            "required": ["company_name"],
        },
    }


    #pdf_info = get_pdf_info_in(folder_path,prompt,response_schema)

    for pathname, dirnames, filenames in os.walk(folder_path):
        for filename in filenames:
            #ファイル名のチェック
            if filename.endswith('.pdf') and (not filename.startswith('-')):
                #ファイルから情報を取得
                pdf_path = os.path.join(pathname,filename)
                #print("get_pdf_info")
                pdf_info = get_pdf_info(pdf_path,prompt,response_schema)
                
                if pdf_info is None:
                    time.sleep(QUERY_WAIT_SEC)
                    continue

                print("pdf_info",pdf_info)
                time.sleep(QUERY_WAIT_SEC)

if __name__ == "__main__":
    main()


