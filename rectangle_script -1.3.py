import shutil, os
import pandas as pd
import sdxf

excel_data = pd.read_excel(r'C:\Bazy\ExportPT\WYKAZ_DXF.xlsx')


file_names = excel_data["NAZWAPLIKU"]
w_a = excel_data["W_A"]
w_b = excel_data["W_B"]
w_c = excel_data["W_C"]
plate_names = excel_data["NAZWABL"]

def draw(x, y, name, folder):
    d = sdxf.Drawing()
    d.layers.append(sdxf.Layer(name="textlayer", color=3))
    d.append(sdxf.Rectangle(point=(0, 0, 0), width=x, height=y, color=2))
    name += '.dxf'
    d.saveas(name)
    shutil.move(name, folder)
    

#os.makedirs(gr)

for j in range(len(plate_names)):
    try:
        os.makedirs(plate_names[j])
        print('utworzono folder', plate_names[j])
        
    except:
        print(f'folder {plate_names[j]} już istnieje')
    
for i in range(len(file_names)):
    try:
        draw(w_b[i], w_c[i], file_names[i], plate_names[i])
        print(f'utworzono plik:{file_names[i]}, w katalogu {plate_names[i]}')
    except:
        print('plik {file_names[i]} już istnieje')
