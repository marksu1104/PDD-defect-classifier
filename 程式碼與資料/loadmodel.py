import pdfplumber as pr
import pandas as pd
import joblib
import os

def countif(line, base, column_name, real_name):        #判斷瑕疵偵測
    if column_name in line[base]:
        line[real_name] = 1
    else:
        line[real_name] = 0
    return line

def model(dr):                      #導入model
    pdf = pr.open(dr)

    for pg in pdf.pages:            #從pdf抓出表格並轉換成df
        tables = pg.extract_tables()
        if pg == pdf.pages[0]:    
            table = tables[1]
            df = pd.DataFrame(table[1:],columns = table[0])
        else:
            table = tables[0]
            temp = pd.DataFrame(table[1:],columns = table[0])
            df = pd.concat([df, temp], axis=0, ignore_index=True)

    #新增五列            
    df.insert(2, '橫向缺點', 0)
    df.insert(2, '縱向缺點', 0)
    df.insert(2, '面狀輕', 0)
    df.insert(2, '面狀重', 0)
    df.insert(2, '點狀', 0)
    
    #模型採用的
    features = ['X位', '點狀', '面狀重', '面狀輕', '縱向缺點', '橫向缺點', '逆位置',
                '縱向', '橫向', '缺點面積', '混合像', '暗缺點', '明亮缺']

    #呈現的全屬性
    all_features = ['疵號', 'X位', '點狀', '面狀重', '面狀輕', '縱向缺點', '橫向缺點', '逆位置',
                    '縱向', '橫向', '缺點面積', '混合像', '暗缺點', '明亮缺']
    
    #全屬性
    real_all_features = ['疵號', 'X位', '點狀', '面狀重', '面狀輕', '縱向缺點', '橫向缺點', '種類', '逆位置',
                    '縱向', '橫向', '缺點面積', '混合像', '暗缺點', '明亮缺', '結果']
    
    #針對25p2的報告重新命名列名
    df.columns = real_all_features

    #文字轉數字、Nan
    for i in all_features:
        df[i] = pd.to_numeric(df[i], errors='coerce')    
        
    #從種類名稱中判斷有哪些特徵
    df = df.apply(countif, axis=1, args=('種類', '点状', '點狀'))
    df = df.apply(countif, axis=1, args=('種類', '面状重', '面狀重'))
    df = df.apply(countif, axis=1, args=('種類', '淡', '面狀輕'))
    df = df.apply(countif, axis=1, args=('種類', '縱向', '縱向缺點'))
    df = df.apply(countif, axis=1, args=('種類', '橫向', '橫向缺點'))

    #載入模型
    ng = joblib.load("model/NG.m")
    type = joblib.load("model/train_model.m") 

    na = df[df.isnull().T.any()].drop(columns=['結果', '種類'])     #抓出NA，呈現資料時放回
    x = df.dropna(how='any').drop(columns=['結果', '種類'])

    #predict
    p_ng = ng.predict(x[features])
    p_type = type.predict(x[features])

    #類別轉文字
    ng_map = {0:'NG', 1:'放行'}
    x['判斷'] = p_ng
    x['判斷'] = x['判斷'].map(ng_map)
    type_map = {0:'塞版', 1:'套色偏移', 2:'拖墨', 3:'染墨', 4:'死紋', 5:'髒污'}
    x['種類'] = p_type
    x['種類'] = x['種類'].map(type_map)

    #合併NA
    new = pd.concat([x,na],axis=0).sort_values(by=['疵號'])
    
    #取得檔名
    basename = os.path.basename(dr)
    filename = os.path.splitext(basename)[0]

    #寫入excel
    writer = pd.ExcelWriter("結果/"+ filename + "判斷結果.xlsx")
    new.to_excel(writer, sheet_name = '判斷結果')
    df.to_excel(writer, sheet_name = '原始資料')
    writer.close()
    
if __name__ == '__main__':
    model('數據/1.pdf')