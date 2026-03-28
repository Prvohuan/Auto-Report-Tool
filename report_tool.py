import re
import pandas as pd
import datetime
import tkinter as tk
from tkinter import messagebox
import os

def split_part_and_action(text):
    """根据常见施工动名词智能拆分部位和内容"""
    actions = ['浇筑混凝土', '模板拆除', '支架拆除', '平台搭建', '模板打磨', '钢筋绑扎', 
               '安装模板', '破桩头', '成孔检测', '拆模板', '浇筑', '搭建', '打磨', 
               '凿毛', '焊接', '铺装', '回填', '安装', '施工', '开钻', '整理', '砌']
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
        blocks = re.split(r'(?=\n?时间:)', text.strip())
        
        records = []
        for block in blocks:
            block = block.strip()
            if not block or not block.startswith('时间:'): 
                continue
                
            def extract_field(pattern, block_text, flags=0):
                match = re.search(pattern, block_text, flags)
                return match.group(1).strip() if match else ''

            # 分离人员与机械
            pm_text = extract_field(r'施工人员:(.*?)(?=\n|$)', block)
            num_match = re.search(r'(\d+)\s*人?', pm_text)
            personnel_num = num_match.group(1) if num_match else ''
            
            if num_match:
                machine_text = pm_text.replace(num_match.group(0), '', 1)
            else:
                machine_text = pm_text
            machine_text = machine_text.strip(' 、，,')
            
            # 智能拆分施工部位与内容
            content_raw = extract_field(r'施工内容:(.*?)(?=\n(?:施工人员|队伍)|$)', block, re.DOTALL)
            sub_tasks = re.split(r'[，,\n]', content_raw)
            buweis, neirongs = [], []
            
            for task in sub_tasks:
                task = task.strip()
                if not task: continue
                bw, nr = split_part_and_action(task)
                if bw: buweis.append(bw)
                if nr: neirongs.append(nr)
            
            # 【终极防弹设计】：在这里直接把所有列死死固定，绝不会再报 KeyError
            record = {
                '序号': 0,  # 临时占位，下面会统一刷新
                '桩号': extract_field(r'桩号:(.*?)(?=\n|$)', block),
                '部位': '；'.join(buweis),
                '施工内容': '；'.join(neirongs),
                '施工人数': personnel_num,
                '机械': machine_text,
                '劳务队伍': extract_field(r'队伍:(.*?)(?=\n|$)', block),
                '时间': extract_field(r'时间:(.*?)(?=\n|$)', block),
                '备注': ''
            }
            records.append(record)
            
        if not records:
            messagebox.showwarning("提示", "未能识别到有效数据，请确保格式正确。")
            return
            
        # 生成表格并刷新正确的序号
        df = pd.DataFrame(records)
        df['序号'] = range(1, len(df) + 1)
        
        # 导出为 Excel 到用户桌面
        date_str = datetime.datetime.now().strftime("%Y-%m-%d")
        file_name = f"{date_str}二工区台账.xlsx"
        desktop_path = os.path.join(os.path.expanduser("~"), 'Desktop')
        output_path = os.path.join(desktop_path, file_name)
        
        df.to_excel(output_path, index=False)
        
        messagebox.showinfo("处理成功", f"成功解析 {len(records)} 条数据！\n表格已生成在桌面：\n{file_name}")
        text_box.delete("1.0", tk.END) 
        
    except Exception as e:
        messagebox.showerror("处理失败", f"程序遇到错误，详情：{str(e)}")

# ================= 可视化界面设置 =================
root = tk.Tk()
root.title("台账自动化生成工具 (智能解析版)")
root.geometry("600x480")
root.configure(bg="#f0f0f0")

label = tk.Label(root, text="请将微信群内的汇报内容（可连带多条）粘贴在下方：", font=("微软雅黑", 11), bg="#f0f0f0")
label.pack(pady=(20, 10))

text_box = tk.Text(root, width=70, height=20, font=("微软雅黑", 10))
text_box.pack(pady=5)

btn = tk.Button(root, text="一键生成当日 Excel", command=process_text, font=("微软雅黑", 12, "bold"), bg="#2c3e50", fg="white", width=22)
btn.pack(pady=15)

root.mainloop()
