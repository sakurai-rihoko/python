import pandas as pd
import os
import re
import difflib

#設問と回答データの移行　葉　2023/03/10

#Questions model
class Question:
    def __init__(self,context,note,weight,time):
        self.context = context
        self.note = note
        self.weight= weight
        self.time = time
QUESTION_REG_TEXT = r"^\d+\."
PATTERN = re.compile(QUESTION_REG_TEXT)
NOTE_REG_TEXT = "*"
QUESTION_REGIONAL = "<"
#DEFINE
MONTHS = ["Apr.","May.","Jun.","Jul.","Aug.","Sep.","Oct.","Nov.","Dec.","Jan.","Feb.","Mar.","May","1月","2月","3月","4月","5月","6月","7月","8月","9月","10月","11月","12月"]
MONTHS_DIC={"1月":"01","2月":"02","3月":"03","4月":"04","5月":"05","6月":"06","7月":"07","8月":"08","9月":"09","10月":"10","11月":"11","12月":"12"}
MONTHS_DIC2={"1月":"Jan","2月":"Feb","3月":"Mar","4月":"Apr","5月":"May","6月":"Jun","7月":"Jul","8月":"Aug","9月":"Sep","10月":"Oct","11月":"Nov","12月":"Dec"}
YEAR = "2022"
YEAR2 = "2023"
LINE_ITEM_QUESTION = "Question "
LINE_ITEM_QUESTION_NOTE = "Question Notes "
LINE_ITEM_ANSWER = "Answer "

def getCompanyName(originName):
    if (originName=="北米その他（B030 ACC）"):
        return "B030"
    else:
        return originName.split(" ")[0]
    
def getAllXlsxFile(path):
    all_files = os.listdir(path)
    xlsx_files = [f for f in all_files if f.endswith('.xlsx')]
    print("処理ファイル：")
    print(xlsx_files)
    return xlsx_files
def getMonths(str):
    start_index = str.index('【') + 1
    end_index = str.index('度')
    if (MONTHS_DIC.get(str[start_index:end_index]) == "01" or MONTHS_DIC.get(str[start_index:end_index]) == "02" or MONTHS_DIC.get(str[start_index:end_index]) == "03") :

        return YEAR2 + MONTHS_DIC.get(str[start_index:end_index])
    else :
        return YEAR + MONTHS_DIC.get(str[start_index:end_index])
# def changeMonths(str):
#     if "月" in str:
#         return MONTHS_DIC2.get(str)
#     else :
#         return df.loc[7,col].rstrip('.')
def questionListCheck(csv_questions_list,text,time):
    if len(csv_questions_list) == False:
        temp = Question(text,"",1,time)
        csv_questions_list.append(temp)
        return
    for question in csv_questions_list:
        seq=difflib.SequenceMatcher(None,text,question.context)
        if seq.ratio() > 0.95:
            question.weight = question.weight + 1
            return
    temp = Question(text,"",1,time)
    csv_questions_list.append(temp)
            
def addNote(csv_questions_list,text):
    csv_questions_list[-1].note=text
def saveToMap(questions,map):
    question_num = 1
    for question in questions:
        questions_dict = {"Time":question.time,"LineItem":LINE_ITEM_QUESTION+str(question_num).zfill(2),"Text":question.context}
        note_dict = {"Time":question.time,"LineItem":LINE_ITEM_QUESTION_NOTE+str(question_num).zfill(2),"Text":question.note}
        map.append(questions_dict)
        map.append(note_dict)
        question_num = question_num + 1

def questionRegionalAdd(csv_questions_list,text):
    csv_questions_list[-1].context=csv_questions_list[-1].context+text

#インポートパラメータ
#パラメータ１：エクセルフォルダ
#パラメータ２：年
XLSX_FILES = getAllXlsxFile("C:\\Users\\A783600\\Desktop\\FY2022-20230222T063746Z-001\\FY2022")
    
sheet_names = ["B020 AAI","B070 Canada","北米その他（B030 ACC）",
               "E010 UK","E250 Iberica","E090 Scandinavia","E110 Denmark","E120 Norway",
               "E020 GmbH","E040 OOO","E050 Poland (Total)","E080 SA","E100 Swiss",
               "E180 Turkey","E140 Italia (IT Block)","E130 AOSA","E160 AAG","E190 AEE",
               "E310 AISE","E320 AAE","L010 天田中国","L060 Beijing","L040 Shanghai","L050 Shenzhen","L020 天田香港",
               "L090 天上机","H010 Singapore","H110 Thailand","H020 Malaysia","H060 Vietnam","H120 Indonesia","T010 Taiwan","T020 Korea",
               "K010 Oceania","P020 India","S010 JHB","R010 Brasil","Q010 Middle East","国内"]
# sheet_names = ["B020 AAI"]

csv_questions_list=[]
csv_common_questions=[]
csv_specific_questions=[]

common_questions_dicts = []
specific_questions_dicts = []
answer_dicts = []

print("インポート用CSVフォーマット処理始めます")
for excel_file_path in XLSX_FILES:
    print(excel_file_path+"処理開始")
    time = getMonths(excel_file_path)
    #インポート用CSVフォーマット_契約
    for sheet_name in sheet_names:
        print("シート"+sheet_name+"処理開始")
        df =  pd.read_excel(excel_file_path, sheet_name=sheet_name, header=None)
        if sheet_name == "国内":
                sheet_name = "A010 国内"
        df = df.iloc[74:300,1:30]
        for index, row in df.iterrows():
            text=""
            # values = [str(row[col_name]) for col_name in df.columns if pd.notna(row[col_name])]
            values =str(row[1])
            text=text.join(values)
            if re.match(QUESTION_REG_TEXT,text):
                questionListCheck(csv_questions_list,text,time)
            if text.find(QUESTION_REGIONAL)!=-1:
                questionRegionalAdd(csv_questions_list,text)
            if text.find(NOTE_REG_TEXT)!=-1:
                addNote(csv_questions_list,text)
            if not re.match(QUESTION_REG_TEXT,text):
                df.at[index,1] = None
        answer_groups_df = df.groupby(df.iloc[:,0].notnull().cumsum())
        answer_groups_list = [pd.DataFrame(data) for _, data in answer_groups_df]
        ans_num  = 1
        for data in answer_groups_list:
            filtered_data = df.iloc[:, 2].apply(lambda x: not pd.isna(x) and isinstance(x, str))
            ans_context = data.iloc[:,2][filtered_data].astype(str).str.cat()
            if ans_context != "":
                ans_dict = {"company":getCompanyName(sheet_name),"Time":time,"LineItem":LINE_ITEM_ANSWER+str(ans_num).zfill(2),"Value":ans_context}
                answer_dicts.append(ans_dict)
                ans_num = ans_num + 1
            
    for question in csv_questions_list:
        if (question.weight >= 10):
            csv_common_questions.append(question)
        if (question.weight < 10):
            csv_specific_questions.append(question)
    saveToMap(csv_common_questions,common_questions_dicts)
    saveToMap(csv_specific_questions,specific_questions_dicts)
    csv_common_questions = []
    csv_specific_questions= []
common_df = pd.DataFrame(common_questions_dicts)
specific_df = pd.DataFrame(specific_questions_dicts)
ans_df = pd.DataFrame(answer_dicts)

if os.path.isfile("「1.4.1設問(全社共通)」画面データ__インポート用CSVフォーマット"+YEAR+".csv"):
    os.remove("「1.4.1設問(全社共通)」画面データ__インポート用CSVフォーマット"+YEAR+".csv")
if os.path.isfile("「1.4.2設問(全社別)」」画面データ__インポート用CSVフォーマット"+YEAR+".csv"):
    os.remove("「1.4.2設問(全社別)」」画面データ__インポート用CSVフォーマット"+YEAR+".csv")
if os.path.isfile("「2.1CSP」画面のコメントデータ__インポート用CSVフォーマット"+YEAR+".csv"):
    os.remove("「「2.1CSP」画面のコメントデータ__インポート用CSVフォーマット"+YEAR+".csv")
print(common_df)
common_df.to_csv("「1.4.1設問(全社共通)」画面データ__インポート用CSVフォーマット"+YEAR+".csv",index=False)
specific_df.to_csv("「1.4.2設問(全社別)」」画面データ__インポート用CSVフォーマット"+YEAR+".csv",index=False)
ans_df.to_csv("「「2.1CSP」画面のコメントデータ__インポート用CSVフォーマット"+YEAR+".csv",index=False)








