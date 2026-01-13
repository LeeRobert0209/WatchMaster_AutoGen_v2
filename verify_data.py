
import pandas as pd
output_path = r'data/维修师数据_清洗版.xlsx'

try:
    df = pd.read_excel(output_path)
    # Check Yancheng entry
    print("Checking Yancheng entries:")
    yancheng = df[df['门店'].str.contains("盐城", na=False)]
    print(yancheng[['姓名', '门店', '匠龄', '匠人独白']].to_string())
    
    print("\nChecking for any names with numbers or '年':")
    suspicious = df[df['姓名'].astype(str).str.contains(r'\d|年', regex=True)]
    if not suspicious.empty:
        print(suspicious[['姓名', '门店']])
    else:
        print("No suspicious names found.")

    print("\nChecking Descriptions with Numbers or '+':")
    # Search for rows where desc contains digits or +
    for idx, row in df.iterrows():
        d1 = str(row['描述1'])
        if '+' in d1 or any(char.isdigit() for char in d1):
            # Print only extensive matches to save space? just print first 5
            if idx < 5:
                print(f"Row {idx} Desc1: {d1}")

    print("\nChecking Experience column for spaces:")
    print(df['匠龄'].unique())
except Exception as e:
    print(f"Error reading output excel: {e}")
