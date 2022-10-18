import json, pandas as pd
from tabula.io import read_pdf
temp = []
src_file = 'PATH_TO:\myfile\MITRA.pdf'

def scrape_table_sakti(src_file, target_file):
    for pg in range(0, 11):
        print("page %d"%(pg+1))
        tb = read_pdf(src_file, pages=pg+1)
        if pg == 0:
            try:
                tb_target1 = tb[3]
                pds = pd.DataFrame(tb_target1)
            except IndexError: continue
            try:
                tb_target2 = tb[4]
                pds2 = pd.DataFrame(tb_target2)
            except IndexError: continue
            
            try:
                j = {   
                    "nama": pds.loc[0][3],
                    "npwp": pds.loc[1][3],
                    "alamat": pds.loc[2][3],
                    "kode_pos": pds.loc[6][3],
                    "norek": pds.loc[8][3],
                    "nama_rekening": pds.loc[9][3]
                }
            except IndexError: continue
        else:
            try:
                tb_target1 = tb[2]
                pds = pd.DataFrame(tb_target1)
            except IndexError: continue
            try:
                tb_target2 = tb[3]
                pds2 = pd.DataFrame(tb_target2)
            except IndexError: continue
            
            try:
                j = {
                    "nama": pds.loc[0][1],
                    "npwp": pds.loc[1][1],
                    "alamat": pds.loc[2][1],
                    "kode_pos": pds.loc[6][1],
                    "norek": pds.loc[8][1],
                    "nama_rekening": pds.loc[9][1]
                }
            except IndexError: continue

        try:
            k = {
                    "nama": pds2.loc[0][1],
                    "npwp": pds2.loc[1][1],
                    "alamat": pds2.loc[2][1],
                    "kode_pos": pds2.loc[6][1],
                    "norek": pds2.loc[8][1],
                    "nama_rekening": pds2.loc[9][1]
                }
        except IndexError: continue
        
        try:
            temp.append(j)
            print(j)
            temp.append(k)
            print(k)
        except: continue

    with open(target_file,'w') as newtxt:
        a = json.dumps(temp)
        newtxt.write(a)
    newtxt.close()

def overview(src_file):
    for pg in range(0,2):
        print("page %d"%(pg+1))
        tb = read_pdf(src_file, pages=pg+1)
        if pg == 0:
            try:
                tb_target1 = tb[3]
                pds = pd.DataFrame(tb_target1)
            except IndexError: continue
            try:
                tb_target2 = tb[4]
                pds2 = pd.DataFrame(tb_target2)
            except IndexError: continue
            
            try:
                j = {   
                    "nama": pds.loc[0][3],
                    "npwp": pds.loc[1][3],
                    "alamat": pds.loc[2][3],
                    "kode_pos": pds.loc[6][3],
                    "norek": pds.loc[8][3],
                    "nama_rekening": pds.loc[9][3]
                }
            except IndexError: continue
        else:
            try:
                tb_target1 = tb[2]
                pds = pd.DataFrame(tb_target1)
            except IndexError: continue
            try:
                tb_target2 = tb[3]
                pds2 = pd.DataFrame(tb_target2)
            except IndexError: continue
            
            try:
                j = {
                    "nama": pds.loc[0][1],
                    "npwp": pds.loc[1][1],
                    "alamat": pds.loc[2][1],
                    "kode_pos": pds.loc[6][1],
                    "norek": pds.loc[8][1],
                    "nama_rekening": pds.loc[9][1]
                }
            except IndexError: continue

        try:
            k = {
                    "nama": pds2.loc[0][1],
                    "npwp": pds2.loc[1][1],
                    "alamat": pds2.loc[2][1],
                    "kode_pos": pds2.loc[6][1],
                    "norek": pds2.loc[8][1],
                    "nama_rekening": pds2.loc[9][1]
                }
        except IndexError: continue
        
        try:
            temp.append(j)
            print(j)
            temp.append(k)
            print(k)
        except: continue

def json_to_csv(a, b):
    with open(a) as q:
        df = pd.read_json(q, dtype={"npwp":"string", "norek":"string"})
        df_c = pd.DataFrame(df)
        df_c.to_csv(b, sep=";")
        print("File sudah dikonversi ke Excel.")
    q.close()

y = "PATH_TO:\myfile\MITRA.csv"
x = "PATH_TO:\myfile\MITRA.json"

#scrape_table_sakti(src_file, x)
json_to_csv(x,y)

