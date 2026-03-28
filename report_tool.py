import re
import pandas as pd
import datetime
import tkinter as tk
from tkinter import messagebox
import os
from sqlalchemy import create_engine

def split_part_and_action(text):
    """根据常见施工动名词智能拆分部位和内容"""
    actions = ['平台搭建', '模板打磨', '钢筋绑扎', '安装模板', '破桩头', '成孔检测', 
               '拆模板', '浇筑', '搭建', '打磨', '凿毛', '焊接', '铺装', '回填', '砌']
    actions.sort(key=len, reverse=True)
    
    for action in actions:
        if action in text:
            parts = text.split(action, 1)
            return parts[0].strip(), action + parts[1].strip()
            
    if len(text) > 4:
        return text[:-2], text[-2:]
    return '', text

def process_text():
    raw_text = text_box.get("1.0", tk.END)
    if not raw_text.strip():
        messagebox.showwarning("提示", "请输入微信群的汇报内容！")
        return

    try:
        text = raw_text.replace('：', ':')
        blocks = re.split(r'\n\s*\n', text.strip())
        records = []
        
        for block in blocks:
            if not block.strip(): continue
            record = {}
            
            def extract_field(pattern, block_text):
                match = re.search(pattern, block_text)
                return match.group(1).strip() if match else ''

            record['时间'] = extract_field(r'时间:(.*?)(?=\n|$)', block)
            record['桩号'] = extract_field(r'桩号:(.*?)(?=\n|$)', block)
            record['劳务队伍'] = extract_field(r'队伍:(.*?)(?=\n|$)', block)
            
            personnel_text = extract_field(r'施工人员:(.*?)(?=\n|$)', block)
            machine_text = extract_field(r'机械:(.*?)(?=\n|$)', block)
            
            if '、' in personnel_text and not machine_text:
                parts = personnel_text.split('、')
                personnel_raw = parts[0]
                record['机械'] = parts[1] if len(parts) > 1 else ''
            else:
                personnel_raw = personnel_text
                record['机械'] = machine_text
                
            num_match = re.search(r'\d+', personnel_raw)
            record['施工人数'] = num_match.group() if num_match else ''
            
            content_match = re.search(r'施工内容:(.*?)(?=\n(?:施工人员|机械|队伍)|$)', block, re.DOTALL)
            if content_match:
                raw_contents = content_match.group(1).strip().split('\n')
                buweis, neirongs = [], []
                for line in raw_contents:
                    line = line.strip().replace('\xa0', '')
                    if not line: continue
                    bw, nr = split_part_and_action(line)
                    if bw: buweis.append(bw)
                    if nr: neirongs.append(nr)
                record['部位'] = '；'.join(buweis)
                record['施工内容'] = '；'.join(neirongs)
            else:
                record['部位'] = ''
                record['施工内容'] = ''
                
            record['备注'] = ''
            records.append(record)
            
        df = pd.DataFrame(records)
        columns_order = ['序号', '桩号', '部位', '施工内容', '施工人数', '机械', '劳务队伍', '时间', '备注']
        df['序号'] = range(1, len(df) + 1)
        df = df.reindex(columns=columns_order)
        
        # 1. 导出为 Excel 到用户桌面
        date_str = datetime.datetime.now().strftime("%Y-%m-%d")
        file_name = f"{date_str}二工区台账.xlsx"
        desktop_path = os.path.join(os.path.expanduser("~"), 'Desktop')
        output_path = os.path.join(desktop_path, file_name)
        
        df.to_excel(output_path, index=False)
        
        # 2. 同步写入 PostgreSQL (含异常捕捉，防止对方电脑未配环境导致程序崩溃)
        db_status = ""
        try:
            # 这里的连接信息可以根据你的实际数据库地址修改
            engine = create_engine('postgresql://postgres:password@localhost:5432/highway_project')
            df.to_sql('construction_logs', engine, if_exists='append', index=False)
            db_status = "\n数据已同步至 PostgreSQL。"
        except Exception:
            db_status = "" # 对朋友隐藏数据库连接失败的底层信息
        
        messagebox.showinfo("处理成功", f"表格已自动生成在桌面：\n{file_name}{db_status}")
        text_box.delete("1.0", tk.END) # 清空输入框方便下次使用
        
    except Exception as e:
        messagebox.showerror("处理失败", f"请检查输入格式是否规范。\n错误详情：{str(e)}")

# ================= 可视化界面设置 =================
root = tk.Tk()
root.title("台账自动化生成工具")
root.geometry("550x450")
root.configure(bg="#f0f0f0")

label = tk.Label(root, text="请将微信群内的汇报内容粘贴在下方文本框中：", font=("微软雅黑", 11), bg="#f0f0f0")
label.pack(pady=(20, 10))

text_box = tk.Text(root, width=65, height=18, font=("微软雅黑", 10))
text_box.pack(pady=5)

btn = tk.Button(root, text="一键生成当天 Excel", command=process_text, font=("微软雅黑", 12, "bold"), bg="#2c3e50", fg="white", width=20)
btn.pack(pady=15)

root.mainloop()