import pandas as pd
import os

#インポートスクリプト　葉　2023/02/25



#DEFINE
MONTHS = ["Apr.","May.","Jun.","Jul.","Aug.","Sep.","Oct.","Nov.","Dec.","Jan.","Feb.","Mar."]
MONTHS_DIC={"1月":"01","2月":"02","3月":"03","4月":"04","5月":"05","6月":"06","7月":"07","8月":"08","9月":"09","10月":"10","11月":"11","12月":"12"}

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
    return MONTHS_DIC.get(str[start_index:end_index])

#インポートパラメータ
#パラメータ１：エクセルフォルダ
#パラメータ２：年
XLSX_FILES = getAllXlsxFile("C:\\Users\\A783600\\Desktop\\FY2022-20230222T063746Z-001\\FY2022")
YEAR = "2022"

csv_data_contract=[]
csv_data_backOrder=[]
csv_data_sales=[]

print("インポート用CSVフォーマット処理始めます")
for excel_file_path in XLSX_FILES:
    
    sheet_names = ["B020 AAI","B070 Canada","北米その他（B030 ACC）",
               "E010 UK","E250 Iberica","E090 Scandinavia","E110 Denmark","E120 Norway",
               "E020 GmbH","E040 OOO","E050 Poland (Total)","E080 SA","E100 Swiss",
               "E180 Turkey","E140 Italia (IT Block)","E130 AOSA","E160 AAG","E190 AEE",
               "E310 AISE","E320 AAE","L010 天田中国","L060 Beijing","L040 Shanghai","L050 Shenzhen","L020 天田香港",
               "L090 天上机","H010 Singapore","H110 Thailand","H020 Malaysia","H060 Vietnam","H120 Indonesia","T010 Taiwan","T020 Korea",
               "K010 Oceania","P020 India","S010 JHB","R010 Brasil","Q010 Middle East"]

    #インポート用CSVフォーマット_契約
    for sheet_name in sheet_names:

        df =  pd.read_excel(excel_file_path, sheet_name=sheet_name, header=None)
        df = df.iloc[7:100,3:20]

        output_cols = [col for col in df.columns if df.loc[7,col] in MONTHS]


        for col in output_cols:
            #契約
            output_dict_contract={"":getCompanyName(sheet_name),"":"Contract","":YEAR+getMonths(excel_file_path),"":df.loc[7,col],"After Sales - Service":df.loc[9,col],
                    "After Sales - Tooling":df.loc[10,col],"After Sales - Parts":df.loc[11,col],"Software":df.loc[13,col],
                    "Machine - To Backorder":df.loc[14,col],"Machine - C.S.S.M":df.loc[15,col]}
            csv_data_contract.append(output_dict_contract)
            #受注残
            output_dict_backOrder={"":getCompanyName(sheet_name),"":"BackOrder","":YEAR+getMonths(excel_file_path),"":df.loc[7,col],"(Will be Sold in Current FY 1st Half)":df.loc[30,col],
                    "(Will be Sold in Current FY 2nd Half)":df.loc[31,col],"(Will be Sold in Next FY and After)":df.loc[32,col]}
            csv_data_backOrder.append(output_dict_backOrder)
            #売上
            output_dict_sales={"":getCompanyName(sheet_name),"":"Sales","":YEAR+getMonths(excel_file_path),"":df.loc[7,col],"After Sales - Service":df.loc[40,col],
                    "After Sales - Tooling":df.loc[41,col],"After Sales - Parts":df.loc[42,col],"Software":df.loc[44,col],
                    "Machine - from Previous FY Backorder":df.loc[45,col],"Machine - from Current FY Backorder":df.loc[46,col],"Machine - C.S.S.M":df.loc[47,col]}
            csv_data_sales.append(output_dict_sales)

            contract_df = pd.DataFrame(csv_data_contract)
            backOrder_df = pd.DataFrame(csv_data_backOrder)
            sales_df = pd.DataFrame(csv_data_sales)




print(contract_df)
contract_df.to_csv("インポート用CSVフォーマット_契約.csv")

print(backOrder_df)
backOrder_df.to_csv("インポート用CSVフォーマット_受注残.csv")

print(sales_df)
sales_df.to_csv("インポート用CSVフォーマット_売上.csv")
print("インポート用CSVフォーマット処理完成")






