"""从文件选择窗口比较B识别记录与五种样式的人工基准。"""
import hashlib
import json
import os
import tkinter as tk
import uuid
from pathlib import Path
from tkinter import filedialog,messagebox
from 附注更新.评测 import compare_predictions,DEFAULT_SOURCE

def main():
    root=tk.Tk();root.withdraw()
    try:
        filename=filedialog.askopenfilename(title="选择本程序生成的 更新数据语义映射.json",filetypes=[("识别结果","*.json")],parent=root)
        if not filename:return
        source=Path(filename);prediction=json.loads(source.read_text(encoding="utf-8-sig"))
        identity=prediction.get("original_hash") or prediction.get("source_hash")
        if not identity:raise ValueError("该识别记录没有来源哈希，无法确认对应哪一种人工基准。")
        matches=[]
        for style in range(1,6):
            workbook=DEFAULT_SOURCE/"附注样式示例"/f"附注样式{style}.xlsx"
            if hashlib.sha256(workbook.read_bytes()).hexdigest()==identity:matches.append(style)
        if len(matches)!=1:raise ValueError("识别记录并非正式五种样式之一，不能套用这些人工答案。")
        result=compare_predictions(prediction,matches[0])
        output=source.parent/("人工基准比较_"+uuid.uuid4().hex[:8]+".json")
        with output.open("x",encoding="utf-8-sig") as stream:json.dump(result,stream,ensure_ascii=False,indent=2,default=str)
        reread=json.loads(output.read_text(encoding="utf-8-sig"))
        c=reread["counts"]
        text=(f"样式 {matches[0]}，人工业务格 {c['expected']} 个。\n"
              f"ID一致 {c['correct']}，错误 {c['wrong']}，遗漏 {c['missing']}，重复 {c['duplicate']}。\n"
              f"仍需核实维度 {c['dimension_review']} 项。\n\n"
              "ID一致不等于完整语义已确认。\n结果已保存在识别记录旁边。")
        messagebox.showinfo("人工基准比较",text,parent=root)
        os.startfile(str(output.parent))
    except Exception as error:
        messagebox.showerror("未完成比较",str(error),parent=root)
    finally:root.destroy()

if __name__=="__main__":main()
