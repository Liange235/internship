import os
import pandas as pd
from transformers import AutoTokenizer, AutoModel
    
# def gpt_response(row, flag):
#     conclusion = row['核查结论']
#     if flag:
#         suggestion = row['核查建议']
#         conclusion = conclusion + suggestion
#     query = conclusion+'/n根据上述核查结论'+'/n您认为该公司是否存在资金上违法的情况？/n你只需要告诉我一个结论，不要太长，我重申一遍控制在10个字以内，不要太啰嗦。'
#     response, history = model.chat(tokenizer, query, history=[])
#     return response

def modify_cell(row):
    relative_path = row['相对路径']
    ls_str = relative_path.split('\\')
    output = '\\'.join(ls_str[1:])
    return output

path = os.getcwd()
dir = '/mnt/c/Users/Administrator/Documents/egnail/file/info_statistics.xlsx'
dfs = pd.read_excel(dir, sheet_name=None)
writer = pd.ExcelWriter(dir)
# tokenizer = AutoTokenizer.from_pretrained("/home/egnail/proj/ChatGLM2-6B-model", trust_remote_code=True)
# model = AutoModel.from_pretrained("/home/egnail/proj/ChatGLM2-6B-model", trust_remote_code=True, device='cuda')
# model = model.eval()
for sheet_name, df in dfs.items():
    # suggestion_exist = False
    # for _ in df.columns:
        # if '核查建议' in _:
            # suggestion_exist = True
    # df['Response_by_GPT'] = df.apply(gpt_response, axis=1, args=(suggestion_exist,))
    df['相对路径2'] = df.apply(modify_cell, axis=1)
    df.to_excel(writer, sheet_name=sheet_name, index=False)
writer._save()