# Copyright (c) 2025 Kenichi Nakatani
# https://note.com/kenichi_nakatani/n/n82d474e5c4d1
# Released under the MIT license
# https://opensource.org/licenses/mit-license.php

import os
import sys
import tkinter as tk
from tkinter import filedialog,messagebox
import time
import datetime
import csv

import llms2

#領収書等の借方科目として利用する科目
#消費税　課税
#KAMOKU_KAZEI =["荷造運賃","水道光熱費","旅費交通費","通信費","広告宣伝費","接待交際費","修繕費","消耗品費","車両費","支払手数料","新聞図書費","会議費","雑費","工具器具備品"]
#消費税　非課税
#KAMOKU_HIKAZEI =["損害保険料","利子割引料"]
#消費税　不課税
#KAMOKU_FUKAZEI =["租税公課","法定福利費","給料賃金","諸会費"]
#KAMOKU_ALL = KAMOKU_KAZEI + KAMOKU_HIKAZEI + KAMOKU_FUKAZEI

#消費税　課税
KAMOKU_KAZEI =["外注管理費","水道光熱費","旅費交通費","通信費","広告宣伝費","接待交際費","修繕費","消耗品費","車両費","支払手数料","リース料","雑費","工具器具備品"]
#消費税　非課税
KAMOKU_HIKAZEI =["損害保険料","借入金利子"]
#消費税　不課税
KAMOKU_FUKAZEI =["租税公課","法定福利費","給料手当","諸会費"]
KAMOKU_ALL = KAMOKU_KAZEI + KAMOKU_HIKAZEI + KAMOKU_FUKAZEI


#仕訳の貸方の科目
KAMOKU_Cr ="現金"
KAMOKU_Cr_HOJO =""
#KAMOKU_Cr ="事業主借"
#KAMOKU_Cr_HOJO =""
#KAMOKU_Cr ="短期借入金"
#KAMOKU_Cr_HOJO ="代表者名"

#勘定科目が見つからない場合の科目
KAMOKU_DEFAULT = "仮払金"

#問い合わせの待ち時間
#QUERY_WAIT_SEC = 10

def modify_image_info(image_info):
    """
    image_info内のそれぞれの情報に、
    Dr: デフォルトの場合の借方勘定科目を設定する
    Dr_hojo: 取引先を設定する
    
    """
    for index,res in enumerate(image_info):
        #借方科目の指定がない場合、または、リストの中にない科目である場合
        kamoku = res.get("Dr","") 
        if kamoku=="":
            image_info[index]["Dr"]=KAMOKU_DEFAULT
        elif kamoku not in KAMOKU_ALL:
            image_info[index]["Dr"]=KAMOKU_DEFAULT

        #借方の補助科目に取引先を設定
        payee = res.get('payee',"")
        image_info[index]["Dr_hojo"]=payee  

        #日付をYYYY/MM/DD形式に変換
        orgstr = res.get('date',None)
        #print("orgstr",orgstr,type(orgstr))
        if orgstr is not None:
            adt = datetime.datetime.strptime(orgstr, '%Y-%m-%d')
            newstr = adt.strftime('%Y/%m/%d')
            image_info[index]["date"]=newstr

    return image_info

def add_image_info(image_info,base_res):
    """
    image_info内のそれぞれの情報に共通する情報（base_res）を加える
    辞書型でない要素は削除する
    """
    for index,res in enumerate(image_info):
        if type(res) is not dict:
            continue        
        image_info[index].update(base_res)

    #result_image_info =[]
    #for res in image_info:
    #    if type(res) is not dict:
    #        continue
    #    base_res.update(res)
    #    result_image_info.append(base_res)

    #return result_image_info

    return image_info

def tax_code_correct(tax_code):
    """
    登録番号が一定の書式にしたがっているかどうかを調べる

    Parameters
    ----------
    tax_code : str
        消費税の登録番号

    Returns
    -------
        書式に合致する場合：True
        書式に合致しない場合：False

    Memo
    -------
    消費税の事業者番号はTの後に１３桁の数値が続く数
    例：T1234567890123
    例：Ｔ１２３４５６７８９０１２３

    """
    if tax_code is None:
        return False
    if len(tax_code)==0:
        return False
    if tax_code[0] not in ['T','Ｔ']:
        return False
    houjin_num = tax_code[1:]
    if len(houjin_num)!=13:
        return False

    if houjin_num.isdecimal():
        return True
    return False


def get_default_zeikubun_mf(account):
    """
    与えられた勘定科目について、デフォルトの消費税の税区分を表す文字列（マネーフォワード会計）を返す。

    Parameters
    ----------
    account : str
        勘定科目

    Returns
    -------
        課税の場合："課税仕入 10%"
        非課税の場合："非仕"
        不課税の場合："対象外"

    Memo
    -------
    KAMOKU_KAZEI　に記載されている科目：課税
    KAMOKU_HIKAZEI　に記載されている科目：非課税
    KAMOKU_FUKAZEIに記載されている科目、その他の科目：不課税
    領収書の処理を前提としているので、現在のところ、上記以外の税区分には対応していない。

    """
    if account in KAMOKU_KAZEI:
        return "課税仕入 10%"

    if account in KAMOKU_HIKAZEI:
        return "非仕"

    if account in KAMOKU_FUKAZEI:
        return "対象外"

    return "対象外"

def get_default_zeikubun_yayoi(account,tax_code,kubun_str="区分80%"):
    """
    与えられた勘定科目について、デフォルトの消費税の税区分を表す文字列（マネーフォワード会計）を返す。

    Parameters
    ----------
    account : str
        勘定科目
    tax_code : str
         登録番号
    kubun_str: str
        適格請求書でない場合に税区分に付加する文字列

    Returns
    -------
        課税(適格)の場合："課対仕入込10%適格"
        課税(その他)の場合："課対仕入込10%"にkubun_strを加えた文字列（例：課対仕入込10%区分80%）
        非課税の場合："対象外"
        不課税の場合："対象外"

    Memo
    -------
    KAMOKU_KAZEI　に記載されている科目：課税
    KAMOKU_HIKAZEI　に記載されている科目：非課税
    KAMOKU_FUKAZEIに記載されている科目、その他の科目：不課税
    領収書の処理を前提としているので、現在のところ、上記以外の税区分には対応していない。
    ８％の税率には対応していない。
    """

    if account in KAMOKU_KAZEI:
        if tax_code_correct(tax_code):
            return "課対仕入込10%適格"
        else:
            return "課対仕入込10%"+kubun_str

    if account in KAMOKU_HIKAZEI:
        return "対象外"

    if account in KAMOKU_FUKAZEI:
        return "対象外"

    return "対象外"

def write_yayoi_siwake(index,pdf_info,writer):
    """
    pdf_infoから弥生会計の仕訳をwriterへ書き出す

    Parameters
    ----------
    index : int
        取引No
    pdf_info : list[dict[str,str]]　書類の内容
        - "date": 日付（文字列）
        - "Dr": 借方勘定科目（文字列）
        - "payee": 支払先（文字列）
        - "tax_code": 登録番号（文字列）
        - "Cr": 貸方勘定科目（文字列）
        - "amount": 金額（文字列）
        - "tax_amount": 消費税額（文字列）
    writer : object
        writer オブジェクト

    Memo
    ----------
    支払先は、借方の補助科目として設定しています。借方取引先には設定していません。
    これは、
    ・支払先ごとの金額の集計が可能となる
    ・支払先を取引先として設定すると、事前にソフト上で取引先の設定をおこなってから
    取り込みを行う必要があるので、取り込みの手間が増えてしまう
    ためです。
    """

    for res in pdf_info:
        #取引日
        res_date = res.get('date',"")
        if res_date=="":
            continue

        #借方勘定科目
        account_Dr = res.get('Dr',"")

        #借方補助科目
        account_Dr_hojo = res.get('Dr_hojo',"")

        #貸方勘定科目
        account_Cr = res.get('Cr',"")

        #貸方補助科目
        account_Cr_hojo = res.get('Cr_hojo',"")

        #取引金額(円)
        amount = res.get('amount',0)

        #取引税額（円）
        tax_amount = res.get('tax_amount',0)

        #取引先
        payee = res.get('payee',"")

        #登録番号
        tax_code = res.get('tax_code',"")

        #借方税区分
        zeikubun_Dr = get_default_zeikubun_yayoi(account_Dr,tax_code)

        #貸方税区分
        zeikubun_Cr = get_default_zeikubun_yayoi(account_Cr,None)

        #摘要
        tekiyou = res.get('contents',"")
        if tekiyou is None:
            tekiyou =""
        tekiyou=tekiyou.translate(str.maketrans({'\\': None}))

        writer.writerow({'識別フラグ':"2000",\
                        '伝票No':index,\
                        '決算':"",\
                        '取引日付':res_date,\
                        '借方勘定科目':account_Dr,\
                        '借方補助科目':account_Dr_hojo,\
                        '借方部門':"",\
                        '借方税区分':zeikubun_Dr,\
                        '借方金額':int(amount),\
                        '借方税金額':int(tax_amount),\
                        '貸方勘定科目':account_Cr,\
                        '貸方補助科目':account_Cr_hojo,\
                        '貸方部門':"",\
                        '貸方税区分':zeikubun_Cr,\
                        '貸方金額':int(amount),\
                        '貸方税金額':0,\
                        '摘要':tekiyou,\
                        '番号':"",\
                        '期日':"",\
                        'タイプ':0,\
                        '生成元':"",\
                        '仕訳メモ':"",\
                        '付箋1':"0",        
                        '付箋2':"0",\
                        '調整':"no"})
    return



def write_mf_siwake(index,pdf_info,writer):
    """
    pdf_infoからマネーフォワード会計の仕訳をwriterへ書き出す

    Parameters
    ----------
    index : int
        取引No
    pdf_info : list[dict[str,str]]　書類の内容
        - "date": 日付（文字列）
        - "Dr": 借方勘定科目（文字列）
        - "payee": 支払先（文字列）
        - "tax_code": 登録番号（文字列）
        - "Cr": 貸方勘定科目（文字列）
    writer : object
        writer オブジェクト

    Memo
    ----------
    支払先は、借方の補助科目として設定しています。借方取引先には設定していません。
    これは、
    ・支払先ごとの金額の集計が可能となる
    ・支払先を取引先として設定すると、事前にソフト上で取引先の設定をおこなってから
    取り込みを行う必要があるので、取り込みの手間が増えてしまう
    ためです。
    """

    for res in pdf_info:
        row = []
        #取引No
        row.append(index)
        #取引日
        res_date = res.get('date',"")
        if res_date=="":
            continue
        row.append(res_date)
        #借方勘定科目
        account_Dr = res.get('Dr',"") 
        row.append(account_Dr)
        #借方補助科目
        account_Dr_hojo = res.get('Dr_hojo',"")
        row.append(account_Dr_hojo)        
        #借方部門
        row.append("")
        #借方取引先
        row.append("")
        #row.append(res.get('payee',""))
        #借方税区分
        zeikubun_Dr = get_default_zeikubun_mf(account_Dr)
        row.append(zeikubun_Dr)

        #借方インボイス
        if zeikubun_Dr=="課税仕入 10%":
            tax_code = res.get('tax_code',"")
            if tax_code_correct(tax_code):
                row.append("適格")
            else:
                row.append("80%控除")            
        else:    
            row.append("")

        #借方金額(円)
        amount = res.get('amount',0)
        row.append(amount)

        #借方税額
        row.append("")

        #貸方勘定科目
        account_Cr = res.get('Cr',"") 
        row.append(account_Cr)
        #貸方補助科目
        account_Cr_hojo = res.get('Cr_hojo',"")
        row.append(account_Cr_hojo)
        #貸方部門
        row.append("")
        #貸方取引先
        row.append("")
        #貸方税区分
        zeikubun_Cr = get_default_zeikubun_mf(account_Cr)
        row.append(zeikubun_Cr)

        #貸方インボイス
        if zeikubun_Cr=="課税仕入 10%":
            tax_code = res.get('tax_code',"")
            if tax_code_correct(tax_code):
                row.append("適格")
            else:
                row.append("80%控除")            
        else:    
            row.append("")

        #貸方金額(円)
        row.append(amount)

        #貸方税額
        row.append("")

        #摘要
        tekiyou = res.get('contents',"")
        if tekiyou is None:
            tekiyou =""
        row.append(tekiyou[:200])

        #仕訳メモ
        row.append("")
        #タグ
        row.append("")
        #MF仕訳タイプ
        row.append("")
        #決算整理仕訳
        row.append("")
        #作成日時
        row.append("")
        #作成者
        row.append("")
        #最終更新日時
        row.append("")
        #最終更新者
        row.append("")

        writer.writerow(row)
    return

def write_mf_header(writer):
    """
    マネーフォワード会計の仕訳のヘッダーファイルをwriterへ書き出す

    Parameters
    ----------
    writer : object
        writer オブジェクト
    """

    header = ['取引No','取引日',
    '借方勘定科目','借方補助科目','借方部門','借方取引先','借方税区分','借方インボイス','借方金額(円)','借方税額',
    '貸方勘定科目','貸方補助科目','貸方部門','貸方取引先','貸方税区分','貸方インボイス','貸方金額(円)','貸方税額',
    '摘要','仕訳メモ','タグ','MF仕訳タイプ','決算整理仕訳','作成日時','作成者','最終更新日時','最終更新者']

    writer.writerow(header)

def write_image_info_in(folder_path,prompt,file_path,response_schema,wait_time,model_id,csv_format):
    """
    指定したフォルダの下（そのフォルダの下にフォルダがある場合は、それらの中も含めて）
    にあるファイルの情報を指定したプロンプトでしらべて、その結果をを指定したファイルに
    仕訳として書き出す。
    
    Parameters
    ----------
    folder_path : str
        検索対象のフォルダのパス（文字列）
    prompt : str
        プロンプト（文字列）
    file_path : str
        書き出すファイルのパス（文字列）
    response_schema : dict
        レスポンスのスキーマ（辞書）
    wait_time : int
        待ち時間（秒）
    model_id : str
        モデルID（文字列）
    csv_format : str
        CSVファイルの形式（"MF" or "YAYOI"）

    Returns
    -------
    None

    """

    #info = llms2.get_image_info_in(folder_path,prompt,response_schema,wait_time,model_id)

    with open(file_path, 'w', newline="") as f:

        #write csv header
        if csv_format == "MF":
            writer = csv.writer(f)
            write_mf_header(writer)
        else:
            fieldnames=["識別フラグ","伝票No","決算","取引日付","借方勘定科目","借方補助科目","借方部門","借方税区分","借方金額","借方税金額","貸方勘定科目","貸方補助科目","貸方部門","貸方税区分","貸方金額","貸方税金額","摘要","番号","期日","タイプ","生成元","仕訳メモ","付箋1","付箋2","調整"]
            #fieldnames=["識別フラグ","伝票No","決算","取引日付","借方勘定科目","借方補助科目","借方部門","借方税区分","借方金額","借方税金額","貸方勘定科目","貸方補助科目","貸方部門","貸方税区分","貸方金額","貸方税金額","摘要","番号","期日","タイプ","生成元","仕訳メモ","付箋1","付箋2","調整","借方取引先名","貸方取引先名"]
            writer = csv.DictWriter(f, fieldnames=fieldnames,quotechar='"', quoting=csv.QUOTE_NONNUMERIC)

        #write csv row
        
        index =1 #取引No
        for pathname, dirnames, filenames in os.walk(folder_path):
            for filename in filenames:

                #ファイル名のチェック
                file_extension = os.path.splitext(filename)[1].lower()
                if not file_extension in llms2.ACCEPTED_EXTENSIONS:
                    #print(f"{filename} not ends with accepted extensions.")
                    continue

                if filename.startswith('-'):
                    continue

                #ファイルから情報を取得
                image_path = os.path.join(pathname,filename)

                #print("get_pdf_info")
                #pdf_info = get_pdf_info(pdf_path,prompt)
                image_info = llms2.get_image_info(image_path,prompt,response_schema,model_id)
                
                #print("pdf_info",pdf_info)
                if image_info is None:
                    time.sleep(wait_time)
                    continue

                #print("add_image_info")
                base_res ={'Cr':KAMOKU_Cr,'Cr_hojo':KAMOKU_Cr_HOJO} #貸方勘定科目
                base_res['path']=image_path
                image_info = add_image_info(image_info,base_res)
                #print("image_info_before",image_info)

                #print("modify_image_info")
                image_info = modify_image_info(image_info)
                print("image_info",image_info)

                #仕訳を書き出す
                if csv_format == "MF":
                    write_mf_siwake(index,image_info,writer)
                elif csv_format == "YAYOI":
                    write_yayoi_siwake(index,image_info,writer)
                else:
                    #default
                    write_yayoi_siwake(index,image_info,writer)

                #サーバの負荷を軽減するための
                time.sleep(wait_time)
                index +=1


def display_message():
    """
    起動時のメッセージを表示する

    Returns
    -------
    answer : bool
        「はい」を選択した場合はTrue
        [いいえ] を選択した場合はFalse

    """

    disp_msg = """
    ・領収書の資料が入っているフォルダを選択して下さい。
    フォルダ内の画像ファイル（拡張子 .png .jpg .jpeg .webp .pdf .heic)
    が読み込まれます。
    
    ・実行結果のインポート用ファイルは、そのフォルダの中に、
    「siwakeYAYOI(作成日時).txt」または、
    「siwakeMF(作成日時).csv」
    という名前で作成されます。

    ・画像ファイル以外のファイルは読み込まれません。
    
    ・フォルダ内に別のフォルダがある場合は、そのフォルダ内の資料も
     読み込まれます。
    
    ・領収書ごとに画像ファイルを準備することを想定しています。
    
    ・読み込みの対象から外したい画像ファイルは、ファイル名の先頭に
     「-」(マイナス)記号をつけて下さい。
    
    実行しますか？
    """

    answer = messagebox.askyesno("使い方",disp_msg)
    
    return answer

def get_file_name_to_export(csv_format):
    """
    # 出力用のファイル名を返す
    # csv_format: MF or YAYOI

    Returns
    -------
    file_name : str

    Examples
    -------
    >>>get_file_name_to_export("MF")
    siwakeMF20250104173728.csv

    """
    t_delta = datetime.timedelta(hours=9)
    JST = datetime.timezone(t_delta, 'JST')
    now = datetime.datetime.now(JST)
    # YYYYMMDDhhmmss形式に書式化
    ymdhms = now.strftime('%Y%m%d%H%M%S')
    #print(ymdhms)  # 20211104173728
    if csv_format=="MF":
        file_name = "siwakeMF"+ymdhms+".csv"
    elif csv_format=="YAYOI":
        file_name = "siwakeYAYOI"+ymdhms+".txt"
    else:
        file_name = "siwake"+ymdhms+".csv"

    return file_name


def main():

    #モデルを選択
    default_value = "gemini-2.5-flash"
    model_id = input(f"利用するGeminiのmodel_idを入力してください（Enterキーを押すと: {default_value}）: ") or default_value
    #print(f"入力された値: {model_id}")

    #出力フォーマットを選択
    csv_format_num = input(f"出力フォーマットを選択してください（1:弥生会計CSV, 2:マネーフォワードクラウド会計CSV Enterキーを押すと: 弥生会計CSV）: ") or "1"
    if csv_format_num=="1":
        csv_format = "YAYOI"
    elif csv_format_num=="2":
        csv_format = "MF"
    else:
        csv_format = "MF"
    #print(f"入力された値: {csv_format}")

    #メッセージを表示
    ret = display_message()
    if not ret:
        sys.exit()
    
    #検索対象のフォルダを指定
    folder_path = filedialog.askdirectory()
    print("検索対象のフォルダ:",folder_path)

    #書き込むファイルを指定
    file_name = get_file_name_to_export(csv_format)
    file_path = os.path.join(folder_path,file_name)
    print("出力ファイル:",file_path)

    system_prompt = """
    あなたは会計事務所の職員です。
    書類の内容の確認を手伝ってください。
    JSON形式で日本語で返答してください。
    """

    prompt_info={'doc_type':'領収書',
        'contents':['支払日','登録番号','支払の相手先','合計金額','合計金額に含まれる消費税の金額','支払内容','借方の勘定科目'],
        'comment':"合計金額は税込みの金額でお願いします。支払内容が複数ある場合は空白で区切って連結して下さい。"}

    response_schema = {
        "type": "array",
        "items": {
            "type": "object",
            "properties": {
                "date": {
                   "type": "string",
                   "format": "date",
                   "description":"支払日",
                },
                "tax_code": {
                   "type": "string",
                   "description":"登録番号",
                },
                "payee": {
                   "type": "string",
                   "description":"支払の相手先",
                },
                "amount": {
                    "type": "number",
                    "description":"合計金額",
                },
                "tax_amount": {
                    "type": "number",
                    "description":"合計金額に含まれる消費税の金額",
                },
                "contents": {
                    "type": "string",
                    "description":"支払内容",
                },
                "Dr": {
                    "type": "string",
                    "description":"借方の勘定科目",
                    "enum": KAMOKU_ALL,
                },               
            },
            "required": ["date"],
        },
    }

    user_prompt = llms2.get_user_prompt(prompt_info)
    prompt =[system_prompt,user_prompt]
    wait_time = 1
    
    #info = llms2.get_image_info_in(folder_path,prompt,response_schema,wait_time,model_id)
    #print(info) 

    write_image_info_in(folder_path,prompt,file_path,response_schema,wait_time,model_id,csv_format)

if __name__ == "__main__":
    main()


