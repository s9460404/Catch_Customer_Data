from flask import Flask, request, render_template, jsonify
#import pymysql
import pandas as pd
#from openpyxl import load_workbook
#import charts
# 資料庫參數設定
db_settings = {
    "host": "127.0.0.1",
    "port": 3306,
    "user": "root",
    "password": "1qaz@WSX",
    "db": "customer_data",
    "charset": "utf8"
}

app = Flask(__name__)

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route("/merge_submit", methods=['POST'])
def merge_submit():

    if 'clientFile' not in request.files:
        return jsonify({'error': 'No file part'})
    
    file = request.files['clientFile']

    if file:
        database = pd.read_excel('database.xlsx')
        data = pd.read_excel(file)
        # 使用 rename 方法修改列名
        data = data.rename(columns={'日期': 'date'})
        data = data.rename(columns={'姓名': 'name'})
        data = data.rename(columns={'金額': 'money'})

        result = pd.merge(data, database, on='name', how='left')
        # Convert result to JSON format
        result_json = result.to_json(orient='records', date_format='iso', force_ascii=False)
        #result_json = result.to_json(orient='records')
        #print(result_json)
    '''
    try:
        # 建立Connection物件
        conn = pymysql.connect(**db_settings)
        # 建立Cursor物件
        with conn.cursor() as cursor:

        # 0 > 資料寫入mysql
        # 6 > 使用loop來比對輸出
        # 7 > 使用2個dataframe來比對輸出
            kind = 7
            kind_data = pd.read_excel('test.xlsx')
            
            # 1 > mysql新增資料, 2 > mysql查詢*, 3 > mysql條件查詢, 4 > mysql條件修改, 5 > mysql刪除
            if kind == 1:
                # 新增資料SQL語法
                command = "INSERT INTO data(id, name, phone, addres, remark)VALUES(%s, %s, %s, %s, %s)"
                cursor.execute(command, ("1", "testname", "0912465481", "台北市", "無備註"))
                cursor.execute(command, ("2", "testname2", "0912462126", "台北市", "測試"))
                cursor.execute(command, ("3", "testname3", "0912412315", "新北市", "test"))
                cursor.execute(command, ("4", "testname4", "0912416418", "台中市", "123"))
            elif kind == 2:
                # 查詢資料SQL語法
                command = "SELECT * FROM data"
                # 執行指令
                cursor.execute(command)
                # 取得所有資料
                result = cursor.fetchall()
                print(result)
            elif kind == 3:
                # 查詢資料SQL語法
                command = "SELECT * FROM data WHERE addres = %s"
                # 執行指令
                cursor.execute(command, ("台北市"))
                # 取得所有資料
                result = cursor.fetchall()
                print(result)
            elif kind == 4:
                # 修改資料SQL語法
                command = "UPDATE data SET addres = %s WHERE name = %s"
                # 執行指令
                cursor.execute(command, ("高雄市", "testname3"))
            elif kind == 5:
                # 刪除特定資料指令
                command = "DELETE FROM data"
                # 執行指令
                cursor.execute(command)

            # 儲存變更
            conn.commit()
            
            if kind == 0:
                #讀取客戶資料xlsx輸入db
                #customer_data = pd.read_excel('信東公益捐款人明細_0530(含編號版)(1).xlsx')
                customer_data = kind_data
                # 迭代處理每一行資料
                for index, row in customer_data.iterrows():
                    # 檢查資料庫中是否已存在相同id的記錄
                    query = "SELECT COUNT(*) FROM data WHERE 姓名 = %s"
                    cursor.execute(query, (row['姓名'],))
                    result = cursor.fetchone()

                    if result[0] == 0:
                        # 如果資料庫中沒有相同id的記錄，則插入該筆資料
                        insert_query = "INSERT INTO data (id, 姓名, phone, address, remark) VALUES (%s, %s, %s, %s, %s)"
                        data_tuple = ("", row['姓名'], row['電話號碼'], row['收件地址/單位'], row['備註'])
                        cursor.execute(insert_query, data_tuple)
                        conn.commit()
                        print(f"Inserted data with id {row['姓名']} into database.")
                    else:
                        print(f"Skipping data with id {row['姓名']} because it already exists in the database.")

            if kind == 6:
                #讀取捐款資料
                #customer_data = pd.read_excel('test.xlsx')
                customer_data = kind_data
                data = {}
                id_col = []
                date_col = []
                money_col = []
                name_col = []
                phone_col = []
                address_col = []
                # 迭代處理每一行資料
                for index, row in customer_data.iterrows():
                    # 檢查資料庫中是否已存在相同id的記錄
                    if isNaN(row['姓名']) == False :
                        query = "SELECT * FROM data WHERE name = %s"
                        cursor.execute(query, (row['name'],))
                        result = cursor.fetchone()

                        if result[0] != 0:
                            # 如果資料庫中有相同姓名的記錄
                            date_col.append(str(row['日期']))
                            name_col.append(str(result[1]))
                            money_col.append(str(row['金額']))
                            phone_col.append(str(result[2]))
                            address_col.append(str(result[3]))

                d = {'日期':date_col, '姓名':name_col , '金額':money_col, '電話號碼':phone_col, '收件地址/單位':address_col}
                df = pd.DataFrame(d)
                #print(df)
                df.to_excel('output.xlsx', index=False)
                
                # 打開 Excel 檔案並取得 workbook
                wb = load_workbook('output.xlsx')
                ws = wb.active  # 取得目前的工作表

                # 設定電話號碼和地址欄位的列寬度
                ws.column_dimensions['D'].width = 35  # 電話號碼欄位的列寬度
                ws.column_dimensions['E'].width = 50  # 收件地址/單位欄位的列寬度

                # 儲存修改後的 Excel 檔案
                wb.save('output.xlsx')
                # 關閉工作簿
                wb.close()
            elif kind == 7:
                #讀取捐款資料
                query = "SELECT * FROM data"  # 假設要從名為 'data' 的資料表中讀取所有資料
                db_data = pd.read_sql(query, conn)
                
                #customer_data = pd.read_excel('test.xlsx')
                customer_data = kind_data
                result_left = pd.merge(customer_data, db_data, on='name', how='left')
                #print(result_left)
                result_left.to_excel('output.xlsx', index=False)
                # 打開 Excel 檔案並取得 workbook
                wb = load_workbook('output.xlsx')
                ws = wb.active  # 取得目前的工作表

                # 設定電話號碼和地址欄位的列寬度
                ws.column_dimensions['A'].width = 20
                ws.column_dimensions['E'].width = 15  # 身份證字號欄位的列寬度
                ws.column_dimensions['F'].width = 35  # 電話號碼欄位的列寬度
                ws.column_dimensions['G'].width = 50  # 收件地址/單位欄位的列寬度

                # 儲存修改後的 Excel 檔案
                wb.save('output.xlsx')
                # 關閉工作簿
                wb.close()

            # 關閉資料庫連線
            cursor.close()
            conn.close()
    except Exception as ex:
        print(ex)
    '''

    return jsonify({'status': 'success','result': result_json})

@app.route("/database_submit", methods=['POST'])
def database_submit():
    
    if 'clientFile' not in request.files:
        return jsonify({'error': 'No file part'})
    
    file = request.files['clientFile']

    if file:
        database = pd.read_excel('database.xlsx')
        data = pd.read_excel(file)

        # 迭代處理每一行資料
        for index, row in data.iterrows():
            result = database[database['name'] == row['name']]
            if result.empty: #沒有搜尋到 > 表示可新增
                database = database._append(row, ignore_index=True)

        #print(database)
        database.to_excel('database.xlsx', index=False) #輸出xlsx
    
    return {"status":"sucess"}

@app.route("/update_database", methods=['POST'])
def update_database():
    data = request.get_json()  # 取得 POST 過來的 JSON 資料
    new_database = pd.DataFrame(data)
    new_database.to_excel('database.xlsx', index=False) #輸出xlsx

    return {"status":"update sucess"}

@app.route("/load_database", methods=['POST'])
def load_database():
    database = pd.read_excel('database.xlsx')
    result_json = database.to_json(orient='records', date_format='iso', force_ascii=False)
    return jsonify({'status': 'success','result': result_json})

def isNaN(num):
    return num != num

if __name__ == '__main__':
    app.run(debug=True)