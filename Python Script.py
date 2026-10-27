from flask import Flask, request, jsonify
import pandas as pd
import os

app = Flask(__name__)
EXCEL_PATH = '1268-تسقيط.xlsx'

@app.route('/process-excel', methods=['POST'])
def process_excel():
    data = request.json
    dwg_no = str(data.get('dwg_no'))
    mark = str(data.get('mark'))
    qty = data.get('qty')
    grit_type = data.get('gritType')
    current_date = data.get('currentDate')
    image_source = data.get('imageSource')
    
    sheet_name = f"DWG_{dwg_no}"
    
    df_dwg = pd.read_excel(EXCEL_PATH, sheet_name=sheet_name)
    
    mask = (df_dwg['mark'].astype(str) == mark) | (df_dwg['item'].astype(str) == mark)
    df_dwg.loc[mask, 'qty'] = qty
    df_dwg.loc[mask, 'gritType'] = grit_type
    df_dwg.loc[mask, 'currentDate'] = current_date
    
    calculated_area = df_dwg.loc[mask, 'area'].values[0] if not df_dwg.loc[mask].empty else 0
    df_dwg.loc[mask, 'qty'] = ""
    
    df_release = pd.read_excel(EXCEL_PATH, sheet_name='test006')
    next_ser = len(df_release) + 1
    
    new_row = {
        "Ser": next_ser,
        "DWG No": dwg_no,
        "ITEM": mark,
        "Qty": qty,
        "AREA": calculated_area,
        "RESULT": "",
        "Image Source": image_source
    }
    
    df_release = pd.concat([df_release, pd.DataFrame([new_row])], ignore_index=True)
    
    with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_dwg.to_excel(writer, sheet_name=sheet_name, index=False)
        df_release.to_excel(writer, sheet_name='test006', index=False)
        
    return jsonify({"status": "success", "ser": next_ser, "area": calculated_area})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))