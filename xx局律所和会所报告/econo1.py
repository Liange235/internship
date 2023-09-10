import pandas as pd
from pdf2image import convert_from_path, convert_from_bytes
import pytesseract
from pypdf import PdfReader
import re
import os
from fnmatch import fnmatch
from joblib import Parallel, delayed
from tqdm import tqdm
# import matplotlib.pyplot as plt
import datetime
starttime = datetime.datetime.now()
# plt.ion()
# plt.show()


def read_report(file, s1, s2):
    i = 0
    key_str = 'Not Found'
    ins = 'Not Found'
    name = 'NotFound'
    try:
        reader = PdfReader(file)
        last_ind = 16 if len(reader.pages)>16 else len(reader.pages)
        images = convert_from_path(file, first_page=0, last_page=last_ind, dpi=600)
        for _ in images:
            text2 = pytesseract.image_to_string(_, lang='chi_sim')  # ocr the img to text
            # text2_filtered = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', text2)
            text2 = text2.replace(' ', '')
            text2_filtered = re.sub(r'[\x00-\x09\x0b-\x1f\x7f-\x9f]', '', text2)
            text2_filtered = text2_filtered.replace(' ', '')
            text2_filtered = re.sub(r'\n{2,}', '\n', text2_filtered)
            if i == 0:
                valid_s = []
                for idx, str in enumerate(text2_filtered.split('\n')):
                    if str:
                        valid_s.append(str)
                for idx, str in enumerate(valid_s):
                    if (fnmatch(str, '*告*')) or (fnmatch(str, '*报*')):
                        break          
                found_indices = [idx]
                ins = valid_s[found_indices[0]+1]
                name = valid_s[0]
            # elif (s3 in text2_filtered) & flag:
                # idx1 = text2_filtered.index(s3)
                # idx2 = text2_filtered.index(s4)
                # name = text2_filtered[idx1+6: idx2-1].replace('\n', '')
                # flag = False
            elif s1[0] in text2_filtered:
                text2_filtered = text2_filtered.replace('\n', '')
                idx1 = text2_filtered.index(s1[0])
                try:
                    idx2 = text2_filtered.index(s2)
                    key_str = text2_filtered[idx1:idx2]
                except ValueError:
                    key_str = text2_filtered[idx1:]
                    text = pytesseract.image_to_string(images[i+1], lang='chi_sim')
                    text = text.replace(' ', '')
                    text_filtered = re.sub(r'[\x00-\x09\x0b-\x1f\x7f-\x9f]', '', text)
                    text_filtered = text_filtered.replace(' ', '')
                    text_filtered = re.sub(r'\n{2,}', '\n', text_filtered)
                    idx2 = text_filtered.index(s2)
                    k2 = text_filtered[:idx2]
                    key_str = key_str + k2
                break
            elif s1[1] in text2_filtered:
                text2_filtered = text2_filtered.replace('\n', '')
                idx1 = text2_filtered.index(s1[1])
                try:
                    idx2 = text2_filtered.index(s2)
                    key_str = text2_filtered[idx1:idx2]
                except ValueError:
                    key_str = text2_filtered[idx1:]
                    text = pytesseract.image_to_string(images[i+1], lang='chi_sim')
                    text = text.replace(' ', '')
                    text_filtered = re.sub(r'[\x00-\x09\x0b-\x1f\x7f-\x9f]', '', text)
                    text_filtered = text_filtered.replace(' ', '')
                    text_filtered = re.sub(r'\n{2,}', '\n', text_filtered)
                    idx2 = text_filtered.index(s2)
                    k2 = text_filtered[:idx2]
                    key_str = key_str + k2
                break
            i += 1
    except Exception as e:
        print("Error: Exception occurred during processing.")
        print(e)
        print(file)
        return key_str, name, ins
    
    return key_str, name, ins

def get_files(dir):
    paths, form, fnames = [], [], []
    for filepath, dirnames, filenames in os.walk(dir):
        for filename in filenames:
            if ('.pdf' in filename):
                abs_p = os.path.join(filepath, filename)
                f = abs_p.split('.')
                # nam = get_company_name(f[0])
                paths.append(abs_p)
                form.append(f[1])
                fnames.append(filename)
    return paths, form, fnames
def my_fun(dir, s1, s2):
    if ('.pdf' in dir):
        con, nam, ins = read_report(dir, s1, s2)
        # print('Proc, '+ dir)
        # para.add_run(_)
        # para.add_run('\n')
        return con, nam, ins
def process_files(dir, s1, s2):
    dirs = []
    abs_path, form, file = get_files(dir)
    # para = document.add_paragraph('会所绝对目录：')
    # res = []
    # with tqdm(total=LEN) as progress_bar:
    #     def my_fun_with_progress(dir, s1, s2, s3, s4):
    #         result = my_fun(dir, s1, s2, s3, s4)
    #         progress_bar.update(1)
    #         return result
    for i, _ in enumerate(abs_path):
        str_block = _.split('/')
        dirs.append('\\'.join(_ for _ in str_block[4:]))
    res = Parallel(n_jobs=32)(delayed(my_fun)(_, s1, s2) for _ in tqdm(abs_path))
    # for _ in abs_path:
    #     res.append(my_fun(_, s1, s2, s3, s4))
    return res, form, file, dirs
######################################The main function######################################
def main():
    c_path = os.getcwd()
    dir = r"/mnt/nvme1n1/data/涉非核查（稳定处）/2020-2022会所核查报告（稳定处）"
    s1 = ["核查结论", "审计结论"]
    s2 = "附件"
    res, form, file, dirs = process_files(dir, s1, s2)
    KeyStr = [_[0] for _ in res]
    CompanyName = [_[1] for _ in res]
    Institution = [_[2] for _ in res]
    # document.save(c_path + '/print_info.docx')
    df1 = pd.DataFrame({'机构': pd.Series(Institution), '企业': pd.Series(CompanyName), '源文件': pd.Series(file),
                        '格式': pd.Series(form), '相对路径': pd.Series(dirs), '核查结论': pd.Series(KeyStr)})
    # 将每个 DataFrame 导出到一个单独的工作表，并应用样式
    filepath = c_path+'/info_statistics1.xlsx'
    writer = pd.ExcelWriter(filepath, engine='openpyxl')
    df1.to_excel(writer, sheet_name='2020-2022会所核查报告(稳定处）')
    writer._save()

    endtime = datetime.datetime.now()
    print (f"Total training time: {(endtime - starttime).seconds:d}")
    # plt.draw()
    # input("Press [enter] to close all the figure windows.")
    # plt.close('all')
    
if __name__ == "__main__":
    main()