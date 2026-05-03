import pandas as pd
import json

def xlsx_to_js_array_with_mapping_and_filter(excel_path, column_mapping=None, output_js_path=None):
    """
    读取xlsx文件并转换为JavaScript数组格式，支持中文列名映射和类型筛选
    
    参数:
        excel_path: Excel文件路径
        column_mapping: 列名映射字典，格式: {'中文列名': '英文键名'}
        output_js_path: 输出的JS文件路径（可选）
    """
    try:
        # 读取Excel文件
        df = pd.read_excel(excel_path)
        
        # 默认映射（如果未提供）
        if column_mapping is None:
            column_mapping = {
                '标题': 'title',
                '歌词': 'lyrics',
                '描述': 'description'
            }
        
        # 检查必要的列是否存在
        required_chinese_cols = list(column_mapping.keys())
        missing_cols = [col for col in required_chinese_cols if col not in df.columns]
        if missing_cols:
            raise ValueError(f"Excel文件缺少以下列: {missing_cols}")
        
        # 新增：筛选原创/作词/作曲/合作创作
        # 请根据你的Excel文件实际情况修改下面的类型列名
        type_column = '原创/作词/作曲/合作创作'  # ←←← 修改为你的实际类型列名
        
        if type_column in df.columns:
            # 定义要保留的类型
            allowed_types = ['原创','初创演员']
            
            # 记录筛选前的数量
            original_count = len(df)
            
            # 筛选数据：只保留指定类型的行
            df[type_column] = df[type_column].astype(str)
            df = df[df[type_column].isin(allowed_types)]
            
            # 记录筛选后的数量
            filtered_count = len(df)
            print(f"🔍 筛选前: {original_count} 条记录")
            print(f"🔍 筛选后: {filtered_count} 条记录 (保留类型: {allowed_types})")
            print(f"📋 使用类型列: '{type_column}'")
            
            if filtered_count == 0:
                print("⚠️ 警告: 筛选后没有符合条件的数据!")
                print("   请检查类型列名是否正确，或者Excel中是否有这些类型的值")
                print(f"   类型列中的实际值: {df[type_column].unique()}")
                return None
        else:
            print(f"❌ 错误: 类型列 '{type_column}' 不存在于Excel文件中")
            print(f"   可用的列名: {list(df.columns)}")
            return None
        
        # 转换为字典列表
        question_bank = []
        for _, row in df.iterrows():
            item = {}
            for chinese_col, english_key in column_mapping.items():
                value = row[chinese_col]
                item[english_key] = str(value) if pd.notna(value) else ""
            question_bank.append(item)
        
        # 转换为JavaScript格式
        js_output = "const questionBank = " + json.dumps(
            question_bank, 
            ensure_ascii=False, 
            indent=2,
            separators=(',', ': ')
        ) + ";"
        
        # 输出到文件或直接打印
        if output_js_path:
            with open(output_js_path, 'w', encoding='utf-8') as f:
                f.write(js_output)
            print(f"✅ 已成功导出到: {output_js_path}")
            print(f"📊 共导出了 {len(question_bank)} 条符合条件的记录")
            
            # 显示映射关系
            print("\n📋 使用的列名映射:")
            for chinese, english in column_mapping.items():
                print(f"   '{chinese}' → '{english}'")
        else:
            print(js_output)
            
        return js_output
        
    except Exception as e:
        print(f"❌ 错误: {e}")
        return None

# 使用示例
if __name__ == "__main__":
    # 请将 'input.xlsx' 替换为你的Excel文件路径
    excel_file = "音乐_音乐剧再汇总.xlsx"
    
    # 可选：指定输出的JS文件名
    output_file = "temp_song.js"
    
    # 定义中文列名到英文键名的映射
    column_mapping = {
        '歌名': 'title',     
        '歌词': 'lyrics',     
        '创作初心/歌曲描述': 'description' ,
        '发布时间': 'time',
        '专辑/单曲/EP': 'tag',
        '音频链接': 'listen1',    
        'MV': 'listen2',
        '官摄/MV链接2': 'listen3',
        '原创/作词/作曲/合作创作': 'type'
    }
    
    # 执行转换
    result = xlsx_to_js_array_with_mapping_and_filter(
        excel_file, 
        column_mapping, 
        output_file
    )