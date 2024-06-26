from PrettyColorPrinter import add_printer
from pdferli import get_pdfdf
import numpy as np
import pandas as pd
import os

path = 'C:\\ProjetorPython\\files\\PR_69699742001710_1000_BU9AL80U_Data-_1_9_2020_Hora-_23_4_16.pdf'
if not os.path.exists(path):
    print(f"Error: File not found at {path}")
else:
    print(f"File found at {path}")
    add_printer(1)
    df = get_pdfdf(path, normalize_content=False)
print(df)
togi = []
for r in np.split(df, df.loc[df.aa_element_type == "LTAnno"].index):
    df2 = r.dropna(subset="aa_size")
    if not df2.empty:
     df3 = df2.sort_values(by="aa_x0")
     togi.append(df3.iloc[:1].copy())
     df4 = pd.concat(togi).copy()
     df4.loc[:, "x0round"] = df4.aa_x0.round(2)
     resultado = []
    for name, group in df4.groupby("x0round"):
        if len(group) > 1:
            group2 = group.reset_index(drop=True)
            group3 = np.split(group2, group2.loc[group2.aa_fontname == "CIDFont+F1"].index)
            for group4 in group3:
               if len(group4) > 1:
                group5 = group4.sort_values(by="bb_hierachy_page")
                t1 = group5.aa_text_line.iloc[0]
                t2 = "\n".join(group5.aa_text_line.iloc[1:].to_list())
                resultado.append((t1, t2))
    df5 = pd.DataFrame(resultado).set_index(0)  # .to_excel('c:\\resultadospdf.xlsx')  

       
